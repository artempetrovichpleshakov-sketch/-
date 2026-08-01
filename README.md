# DeepSeek API integration

Минимальный пример подключения DeepSeek API через OpenAI-compatible endpoint.

## Что с этим делать

Если коротко: нужно положить API-ключ в `.env`, запустить Python-файл и получить
ответ DeepSeek в терминале.

```bash
cp .env.example .env
# откройте .env и замените your_deepseek_api_key_here на реальный ключ
python deepseek_client.py "Напиши короткое приветствие"
```

API-ключ создаётся в DeepSeek Platform. Код использует официальный base URL
`https://api.deepseek.com` и модель `deepseek-v4-flash`.

## Настройка

1. Создайте API-ключ в кабинете DeepSeek.
2. Скопируйте пример переменных окружения:

   ```bash
   cp .env.example .env
   ```

3. Укажите ключ в `.env`:

   ```env
   DEEPSEEK_API_KEY=your_deepseek_api_key_here
   DEEPSEEK_BASE_URL=https://api.deepseek.com
   DEEPSEEK_MODEL=deepseek-v4-flash
   DEEPSEEK_SYSTEM_PROMPT=You are a helpful assistant.
   ```

`deepseek_client.py` сам читает `.env`, поэтому устанавливать `python-dotenv` не
нужно. Если переменные уже экспортированы в shell, они имеют приоритет над `.env`.

## Запуск

```bash
python deepseek_client.py "Привет! Кратко расскажи, что умеет DeepSeek API"
```

Скрипт отправляет запрос в `/chat/completions`, передает Bearer-токен из
`DEEPSEEK_API_KEY` и печатает текст ответа ассистента.

## Использование в коде

```python
from deepseek_client import ask_deepseek

answer = ask_deepseek("Составь план интеграции API")
print(answer)
```

Модель, base URL, системный промпт, температуру и лимит токенов можно
переопределить аргументами функции или через переменные окружения.
