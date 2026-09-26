#!/usr/bin/env python3
"""Собирает автономную страницу analiz-grafika-bureniya.html из src/app.html.

В страницу встраиваются библиотека чтения Excel (SheetJS Community Edition 0.18.5)
и шрифт IBM Plex Sans, поэтому она открывается с диска и работает без интернета.
Пакеты берутся из registry.npmjs.org один раз, проверяются по sha256 и кэшируются
в .cache/.

    python3 build.py
"""

import base64
import hashlib
import io
import re
import sys
import tarfile
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "src" / "app.html"
OUT = HERE / "analiz-grafika-bureniya.html"
CACHE = HERE / ".cache"
LICENSES = HERE / "licenses"

SHEETJS_TGZ = "https://registry.npmjs.org/xlsx/-/xlsx-0.18.5.tgz"
PLEX_TGZ = "https://registry.npmjs.org/@fontsource-variable/ibm-plex-sans/-/ibm-plex-sans-5.3.0.tgz"
PINNED = {
    "package/dist/xlsx.full.min.js": "c9506197caf809a075b6dee1da0d36fb19da7158ffe8a88e7b0c96c5d8623c99",
    "package/files/ibm-plex-sans-cyrillic-wdth-normal.woff2": "d0c0257292881fee5bffd2671c380772c22e94f78e1e4bd9ad840aa0fcdef5f0",
    "package/files/ibm-plex-sans-latin-wdth-normal.woff2": "1e875856600a34cb44a406a7deef0e3f9ae70e6f636e0f37fbca2248dec2dca4",
}
# Поднаборы шрифта: русский текст и латиница с типографскими знаками.
FONT_SUBSETS = {
    "cyrillic": "U+0301,U+0400-045F,U+0490-0491,U+04B0-04B1,U+2116",
    "latin": "U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,"
             "U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD",
}
NOTICE = """<!--
  Анализ графика бурения — автономная версия: открывается с диска и работает без интернета.
  Встроено: SheetJS Community Edition 0.18.5 (Apache License 2.0, © SheetJS LLC)
  и шрифт IBM Plex Sans (SIL Open Font License 1.1, © IBM Corp.).
  Собрано build.py из src/app.html; тексты лицензий — в папке licenses/.
-->
"""


def fetch(url):
    """Скачивает архив пакета (или берёт из кэша) и возвращает его файлы."""
    CACHE.mkdir(exist_ok=True)
    path = CACHE / url.rsplit("/", 1)[1]
    if not path.exists():
        print(f"Скачиваю {url}")
        with urllib.request.urlopen(url, timeout=120) as response:
            path.write_bytes(response.read())
    with tarfile.open(fileobj=io.BytesIO(path.read_bytes()), mode="r:gz") as archive:
        return {m.name: archive.extractfile(m).read() for m in archive.getmembers() if m.isfile()}


def pinned(files, name):
    data = files[name]
    digest = hashlib.sha256(data).hexdigest()
    if digest != PINNED[name]:
        sys.exit(f"{name}: sha256 {digest} не совпадает с ожидаемым {PINNED[name]}")
    return data


def main():
    sheetjs = fetch(SHEETJS_TGZ)
    plex = fetch(PLEX_TGZ)

    library = pinned(sheetjs, "package/dist/xlsx.full.min.js").decode("utf-8")
    # Внутри <script> нельзя встречать закрывающий тег и открывающий <script: парсер HTML сломается.
    if re.search(r"</?script", library, re.I):
        sys.exit("В библиотеке встретился тег script — встраивать её как есть нельзя")

    faces = []
    for subset, ranges in FONT_SUBSETS.items():
        woff2 = pinned(plex, f"package/files/ibm-plex-sans-{subset}-wdth-normal.woff2")
        faces.append(
            "@font-face{font-family:'IBM Plex Sans';font-style:normal;font-weight:100 700;"
            "font-stretch:75% 100%;font-display:swap;"
            f"src:url(data:font/woff2;base64,{base64.b64encode(woff2).decode('ascii')}) format('woff2');"
            f"unicode-range:{ranges}}}"
        )

    source = SRC.read_text(encoding="utf-8")
    head, marker, body = source.partition('<div class="page">')
    if not marker:
        sys.exit('В src/app.html нет <div class="page">')
    head, count = re.subn(r"<!--fonts-->.*?<!--/fonts-->", lambda m: "<style>" + "".join(faces) + "</style>", head, flags=re.S)
    if count != 1:
        sys.exit("В src/app.html нет блока <!--fonts-->…<!--/fonts-->")
    if body.count("<!--xlsx-->") != 1:
        sys.exit("В src/app.html нет метки <!--xlsx-->")
    body = body.replace("<!--xlsx-->", "<script>\n" + library.strip() + "\n</script>")

    page = (
        "<!doctype html>\n<html lang=\"ru\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        + NOTICE + head.strip() + "\n</head>\n<body>\n" + marker + body.rstrip() + "\n</body>\n</html>\n"
    )
    if re.search(r"<link\b", page) or "fonts.googleapis.com" in page:
        sys.exit("В автономной странице остались внешние ссылки на шрифты")
    OUT.write_text(page, encoding="utf-8")

    LICENSES.mkdir(exist_ok=True)
    (LICENSES / "SheetJS-Apache-2.0.txt").write_bytes(sheetjs["package/LICENSE"])
    (LICENSES / "IBM-Plex-Sans-OFL-1.1.txt").write_bytes(plex["package/LICENSE"])
    print(f"Сохранено: {OUT.name} ({len(page.encode('utf-8')) / 1048576:.2f} МБ)")


if __name__ == "__main__":
    main()
