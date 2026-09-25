#!/usr/bin/env python3
"""Строит интерактивную диаграмму Ганта из «Графика строительства скважин».

Читает лист «график» книги .xlsb/.xlsx/.xlsm, находит столбцы этапов по
заголовкам (а не по номерам), проверяет данные на противоречия и
собирает автономную HTML-страницу по шаблону template.html.

    python3 build_gantt.py "График.xlsb" -o grafik.html
"""

import argparse
import datetime as dt
from html import escape as html_escape
import json
import re
import sys
import zipfile
from pathlib import Path

from python_calamine import CalamineWorkbook

EPOCH = dt.datetime(1970, 1, 1)
EXCEL_EPOCH = dt.datetime(1899, 12, 30)
TEMPLATE = Path(__file__).with_name("template.html")
DATA_MARK = "/*__DATA__*/null"

# Атрибуты скважины: ключ -> заголовок столбца (после нормализации пробелов).
ATTRS = {
    "fin": "Наличие финансирования",
    "note": "Примечания",
    "queue": "Очередь",
    "rig": "Сцепка Бур Подрядчик Бригада",
    "goal": "Цель бурения",
    "contractor": "Подрядная организация",
    "rigNo": "№ Бригады",
    "rigType": "Тип станка",
    "status": "Статус",
    "field": "Месторождение",
    "pad": "Куст",
    "well": "Скважина",
    "profile": "Профиль",
    "wellType": "Тип скважины",
    "purpose": "Назначение",
    "target": "Цель",
    "layer": "Пласт",
    "frac": "МГРП/ ГРП",
    "fracStages": "Кол-во стадий",
    "footage": "Проходка",
    "tempLaunch": "Запуск по временной схеме",
    "mob": "Мобилизация /демобилизация",
}

# Этапы: ключ -> начало заголовка группы. Порядок = порядок в листе.
STAGES = {
    "kons": "Консервация/простой",
    "montazh": "Монтаж",
    "demontazh": "Демонтаж",
    "peredv": "Передвижка",
    "bur": "Бурение",
    "osvob": "Освобождение устья",
    "negot": "Период неготовности",
    "montKrs": "Монтаж КРС",
    "krs": "КРС.",
    "montGrp": "Монтаж флота ГРП",
    "grp": "ГРП",
    "demGrp": "Демонтаж флота ГРП",
    "gnkt": "ГНКТ",
    "osv": "Освоение",
    "demKrs": "Демонтаж КРС",
    "montGki": "Монтаж ГКИ",
    "gki": "ГКИ",
    "demGki": "Демонтаж ГКИ",
    "obustr": "Обустройство",
    "vnr": "ВНР",
}
MILESTONES = {"okStr": "Окончание строительства", "vvod": "Ввод в эксплуатацию"}

STAGE_NAMES = {
    "kons": "консервация/простой", "montazh": "монтаж", "demontazh": "демонтаж",
    "peredv": "передвижка", "bur": "бурение", "osvob": "освобождение устья",
    "negot": "неготовность КП", "montKrs": "монтаж КРС", "krs": "КРС",
    "montGrp": "монтаж флота ГРП", "grp": "ГРП", "demGrp": "демонтаж флота ГРП",
    "gnkt": "ГНКТ", "osv": "освоение", "demKrs": "демонтаж КРС",
    "montGki": "монтаж ГКИ", "gki": "ГКИ", "demGki": "демонтаж ГКИ",
    "obustr": "обустройство", "vnr": "ВНР",
}
COMPLETION = ["montGrp", "grp", "demGrp", "gnkt", "krs", "osv"]


def norm(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def fold(value):
    """Нормализует смесь латиницы и кириллицы в номерах кустов и скважин."""
    return norm(value).upper().translate(str.maketrans("ABCEHKMOPTX", "АВСЕНКМОРТХ"))


def to_day(value):
    """Дата из ячейки -> дни от 1970-01-01 (float), иначе None."""
    if isinstance(value, dt.datetime):
        return round((value - EPOCH).total_seconds() / 86400, 4)
    if isinstance(value, dt.date):
        return (dt.datetime(value.year, value.month, value.day) - EPOCH).days
    if isinstance(value, (int, float)) and not isinstance(value, bool) and 20000 < value < 80000:
        return round((EXCEL_EPOCH + dt.timedelta(days=value) - EPOCH).total_seconds() / 86400, 4)
    return None


def to_num(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ".").replace("\xa0", "").strip())
    except ValueError:
        return None


def fmt_day(day, with_time=False):
    moment = EPOCH + dt.timedelta(days=day)
    if with_time and (moment.hour or moment.minute):
        return moment.strftime("%d.%m.%Y %H:%M")
    return moment.strftime("%d.%m.%Y")


def cell_text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return norm(value)


def find_schedule_sheet(workbook):
    for name in workbook.sheet_names:
        if norm(name).lower() == "график":
            return name
    for name in workbook.sheet_names:
        rows = workbook.get_sheet_by_name(name).to_python(nrows=12)
        if find_header_row(rows) is not None:
            return name
    sys.exit("Не найден лист с графиком: нужен заголовок со столбцами «Скважина», «Куст» и «Бурение».")


def find_header_row(rows):
    for i, row in enumerate(rows):
        cells = {norm(c) for c in row}
        if {"Скважина", "Куст", "Бурение"} <= cells:
            return i
    return None


def map_columns(header, sub):
    """Возвращает (attrs, stages, milestones) с индексами столбцов."""
    width = max(len(header), len(sub))
    header = [norm(c) for c in header] + [""] * (width - len(header))
    sub = [norm(c) for c in sub] + [""] * (width - len(sub))

    attrs = {}
    for key, title in ATTRS.items():
        for j, text in enumerate(header):
            if text == title:
                attrs[key] = j
                break

    groups = [j for j, text in enumerate(header) if text]
    def group_span(j):
        nxt = next((g for g in groups if g > j), width)
        return range(j, nxt)

    def find_group(prefix, needed):
        # Сначала точное совпадение («Монтаж»), потом по началу («Монтаж КРС / …»).
        # Берём только группу с подзаголовком дат: «ГНКТ» встречается и среди признаков.
        pattern = re.compile(re.escape(prefix) + r"(?![А-Яа-яЁёA-Za-z])")
        exact = [j for j in groups if header[j] == prefix]
        loose = [j for j in groups if j not in exact and pattern.match(header[j])]
        return next((j for j in exact + loose if any(sub[k] == needed for k in group_span(j))), None)

    stages = {}
    for key, prefix in STAGES.items():
        j = find_group(prefix, "Дата начала")
        if j is None:
            continue
        span = group_span(j)
        start = next((k for k in span if sub[k] == "Дата начала"), None)
        end = next((k for k in span if sub[k] == "Дата окончания"), None)
        days = next((k for k in span if sub[k] == "сут"), None)
        if start is not None and end is not None:
            stages[key] = (start, end, days)

    milestones = {}
    for key, title in MILESTONES.items():
        j = find_group(title, "Дата")
        if j is not None:
            milestones[key] = next((k for k in group_span(j) if sub[k] == "Дата"), None)

    missing = [ATTRS[k] for k in ("rig", "pad", "well") if k not in attrs]
    missing += [STAGES[k] for k in ("bur",) if k not in stages]
    if missing:
        sys.exit("В заголовке графика не найдены столбцы: " + ", ".join(missing))
    return attrs, stages, milestones


def file_date(path):
    """Дата сохранения книги из docProps/core.xml, иначе дата изменения файла."""
    try:
        with zipfile.ZipFile(path) as archive:
            core = archive.read("docProps/core.xml").decode("utf-8", "ignore")
        found = re.search(r"<dcterms:modified[^>]*>(\d{4}-\d{2}-\d{2})", core)
        if found:
            return found.group(1)
    except (KeyError, zipfile.BadZipFile, OSError):
        pass
    return dt.date.fromtimestamp(Path(path).stat().st_mtime).isoformat()


def read_wells(rows, header_row, attrs, stages, milestones):
    wells = []
    for r in range(header_row + 2, len(rows)):
        row = rows[r]
        get = lambda j: row[j] if j is not None and j < len(row) else None
        well_name = cell_text(get(attrs["well"]))
        if not well_name:
            continue
        st = {}
        for key, (js, je, jd) in stages.items():
            start, end = to_day(get(js)), to_day(get(je))
            if start is None and end is None:
                continue
            item = {"s": start, "e": end}
            days = to_num(get(jd))
            if days is not None:
                item["d"] = days
            st[key] = item
        if "bur" not in st:
            continue
        record = {"row": r + 1, "well": well_name}
        for key, j in attrs.items():
            if key in ("well",):
                continue
            value = get(j)
            if key in ("footage", "fracStages", "rigNo", "queue"):
                number = to_num(value)
                if number is not None:
                    record[key] = int(number) if number.is_integer() else round(number, 1)
            else:
                text = cell_text(value)
                if text and text != "-":
                    record[key] = text
        for key, j in milestones.items():
            day = to_day(get(j))
            if day is not None:
                record[key] = day
        record["st"] = st
        record["pad"] = record.get("pad", "—")
        record["rig"] = record.get("rig", "—")
        record["label"] = f"{record['pad']}-{well_name}" if well_name.isdigit() else well_name
        record["finOk"] = bool(record.get("fin")) and "не утвержд" not in record["fin"].lower()
        wells.append(record)
    return wells


def check_wells(wells, as_of_day):
    """Ищет противоречия в датах. Возвращает список замечаний."""
    issues = []

    def add(well, level, text):
        issues.append({"well": well["label"], "row": well["row"], "level": level, "text": text})

    for w in wells:
        st = w["st"]
        bur = st["bur"]
        for key, item in st.items():
            s, e = item.get("s"), item.get("e")
            name = STAGE_NAMES[key]
            if s is not None and e is not None:
                span = e - s
                if span < -0.01:
                    add(w, "error", f"{name}: окончание {fmt_day(e)} раньше начала {fmt_day(s)}")
                days = item.get("d")
                if days is not None and span > 0 and abs(span - days) > max(5, 0.2 * days):
                    add(w, "error", f"{name}: по датам {fmt_day(s)}–{fmt_day(e)} выходит {round(span)} сут, "
                                    f"а в столбце «сут» — {round(days)}")
        if bur.get("e") is not None:
            vvod = w.get("vvod")
            if vvod is not None and vvod < bur["e"] - 0.01:
                when = "до начала бурения" if bur.get("s") is not None and vvod < bur["s"] else "раньше окончания бурения"
                add(w, "error", f"ввод в эксплуатацию {fmt_day(vvod)} {when} ({fmt_day(bur['s'])}–{fmt_day(bur['e'])})")
            grp = st.get("grp")
            if grp and grp.get("s") is not None and grp["s"] < bur["e"] - 0.01:
                add(w, "error", f"ГРП начинается {fmt_day(grp['s'])}, до окончания бурения {fmt_day(bur['e'])}")
            completion_end = max((st[k]["e"] for k in COMPLETION if k in st and st[k].get("e") is not None), default=None)
            if vvod is not None and completion_end is not None and vvod < completion_end - 7 and vvod >= bur["e"]:
                add(w, "error", f"ввод {fmt_day(vvod)} раньше окончания освоения/ГРП {fmt_day(completion_end)}")
        status = (w.get("status") or "").lower()
        if status == "бурение" and bur.get("e") is not None and bur["e"] < as_of_day:
            add(w, "warn", f"статус «Бурение», а плановое окончание бурения {fmt_day(bur['e'])} "
                           f"раньше даты файла {fmt_day(as_of_day)}")
        if status.startswith("окончена") and bur.get("e") is not None and bur["e"] > as_of_day + 1:
            add(w, "warn", f"статус «{w['status']}», но окончание бурения {fmt_day(bur['e'])} позже даты файла")
        if not status and bur.get("s") is not None and bur["s"] < as_of_day - 1:
            add(w, "warn", f"бурение по плану началось {fmt_day(bur['s'])}, статус не заполнен")

    # Пересечения работ одной бригады.
    by_rig = {}
    for w in wells:
        for key in ("montazh", "peredv", "bur"):
            item = w["st"].get(key)
            if item and item.get("s") is not None and item.get("e") is not None:
                by_rig.setdefault(w["rig"], []).append((item["s"], item["e"], w, key))
    for rig, items in by_rig.items():
        items.sort(key=lambda x: x[0])
        for (s1, e1, w1, k1), (s2, e2, w2, k2) in zip(items, items[1:]):
            if w1 is not w2 and s2 < e1 - 0.5:
                add(w2, "error", f"{rig}: {STAGE_NAMES[k2]} {w2['label']} с {fmt_day(s2)} пересекается "
                                 f"с работой на {w1['label']} (до {fmt_day(e1)})")

    # Один номер ИК с разными датами.
    decisions = {}
    for w in wells:
        found = re.match(r"(ИКК?\s*\d+)\s+от\s+(\S+)", w.get("fin") or "")
        if found:
            decisions.setdefault(found.group(1), {}).setdefault(found.group(2), []).append(w["label"])
    for number, dates in decisions.items():
        if len(dates) > 1:
            parts = "; ".join(f"«{number} от {d}» — {', '.join(ws)}" for d, ws in dates.items())
            issues.append({"well": None, "row": None, "level": "warn",
                           "text": f"одно решение {number} указано с разными датами: {parts}"})
    return issues


def check_year_totals(rows, header_row, wells):
    """Сверяет столбцы «Проходка Всего за ГГГГг.» с проходкой плановых скважин.

    Для пробуренных и бурящихся скважин в годовых столбцах может стоять факт,
    поэтому они не проверяются.
    """
    sub = [norm(c) for c in rows[header_row + 1]]
    year_cols = []
    for j, text in enumerate(sub):
        found = re.fullmatch(r"Проходка Всего за (\d{4})\s*г\.?", text)
        if found:
            year_cols.append((j, found.group(1)))
    if not year_cols:
        return []

    def metres(value):
        return f"{value:,.1f}".replace(",", " ").replace(".0", "").replace(".", ",")

    issues = []
    for w in wells:
        footage = w.get("footage")
        if not footage or (w.get("status") or "").lower() not in ("", "план", "вне плана"):
            continue
        row = rows[w["row"] - 1]
        by_year = [(year, to_num(row[j]) or 0) for j, year in year_cols if j < len(row)]
        total = sum(v for _, v in by_year)
        if total == 0:
            text = f"проходка {metres(footage)} м не попала в годовые столбцы «Проходка Всего за …г.»"
        elif abs(total - footage) > max(1, 0.02 * footage):
            parts = ", ".join(f"{year} — {metres(v)}" for year, v in by_year if v)
            text = f"в годовых столбцах проходки {metres(total)} м ({parts}) при проходке скважины {metres(footage)} м"
        else:
            continue
        issues.append({"well": w["label"], "row": w["row"], "level": "warn", "text": text})
    return issues


def check_viz_sheet(workbook, wells):
    """Сверяет лист «Данные» (источник макроса «График визуализация») с графиком."""
    if "Данные" not in workbook.sheet_names:
        return None
    rows = workbook.get_sheet_by_name("Данные").to_python()
    head = next((i for i, row in enumerate(rows) if norm(row[0] if row else "") == "Бригада"), None)
    if head is None:
        return None
    control = {}
    if "Управление" in workbook.sheet_names:
        for row in workbook.get_sheet_by_name("Управление").to_python():
            if len(row) > 1:
                control[norm(row[0])] = row[1]
    first = next((to_num(v) for k, v in control.items() if k.startswith("1.")), None)
    years = next((to_num(v) for k, v in control.items() if k.startswith("2.")), None)

    snapshot = {}
    for row in rows[head + 2:]:
        if len(row) > 11 and norm(row[5]):
            s, e = to_day(row[10]), to_day(row[11])
            snapshot[fold(row[5].lstrip("`"))] = (norm(row[0]), s, e)
    current = {}
    for w in wells:
        s, e = w["st"]["bur"].get("s"), w["st"]["bur"].get("e")
        if first and years:
            in_window = any(d is not None and first <= (EPOCH + dt.timedelta(days=d)).year <= first + years - 1
                            for d in (s, e))
            if not in_window:
                continue
        current[fold(w["well"])] = (w["rig"], s, e)

    def same(a, b):
        return a[0] == b[0] and all(x is not None and y is not None and abs(x - y) < 0.05 for x, y in zip(a[1:], b[1:]))

    missing = sorted(k for k in current if k not in snapshot)
    extra = sorted(k for k in snapshot if k not in current)
    changed = sorted(k for k in current if k in snapshot and not same(current[k], snapshot[k]))
    if not (missing or extra or changed):
        return None
    return {"inSheet": len(snapshot), "inSchedule": len(current), "missing": len(missing),
            "extra": len(extra), "changed": len(changed)}


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", help="файл графика (.xlsb, .xlsx, .xlsm)")
    parser.add_argument("-o", "--output", help="куда сохранить HTML (по умолчанию рядом с книгой)")
    parser.add_argument("--json", help="дополнительно сохранить извлечённые данные в JSON")
    args = parser.parse_args()

    path = Path(args.workbook)
    workbook = CalamineWorkbook.from_path(str(path))
    sheet = find_schedule_sheet(workbook)
    rows = workbook.get_sheet_by_name(sheet).to_python()
    header_row = find_header_row(rows)
    attrs, stages, milestones = map_columns(rows[header_row], rows[header_row + 1])
    wells = read_wells(rows, header_row, attrs, stages, milestones)
    if not wells:
        sys.exit("В листе «%s» не найдено ни одной скважины с датами бурения." % sheet)

    as_of = file_date(path)
    as_of_day = to_day(dt.date.fromisoformat(as_of))
    issues = check_year_totals(rows, header_row, wells) + check_wells(wells, as_of_day)
    fields = sorted({w.get("field") for w in wells if w.get("field")})
    data = {
        "meta": {
            "sheet": sheet,
            "asOf": as_of_day,
            "fields": fields,
            "vizSheet": check_viz_sheet(workbook, wells),
        },
        "wells": wells,
        "issues": issues,
    }

    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    template = TEMPLATE.read_text(encoding="utf-8")
    if DATA_MARK not in template:
        sys.exit(f"В шаблоне {TEMPLATE.name} нет метки {DATA_MARK}")
    if len(fields) == 1:
        field = fields[0]
        title = "График бурения " + (field[:-2] + "ого м/р" if field.endswith("ое") else field)
    else:
        title = "График бурения"
    page = template.replace("__TITLE__", html_escape(title)).replace(DATA_MARK, payload)
    output = Path(args.output) if args.output else path.with_suffix(".html")
    output.write_text(page, encoding="utf-8")
    if args.json:
        Path(args.json).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    rigs = sorted({w["rig"] for w in wells})
    print(f"Лист «{sheet}»: {len(wells)} скважин, бригад: {len(rigs)}, замечаний: {len(issues)}")
    print(f"Сохранено: {output}")


if __name__ == "__main__":
    main()
