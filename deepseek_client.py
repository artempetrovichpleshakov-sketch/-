"""Minimal DeepSeek API client.

The client uses DeepSeek's OpenAI-compatible chat completions endpoint and reads
configuration from environment variables by default. For local development it
also loads simple KEY=VALUE pairs from a .env file in the current directory.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any
from urllib import error, request


def load_env_file(path: str | Path = ".env") -> None:
    """Load simple KEY=VALUE pairs from a .env file without external packages."""
    env_path = Path(path)
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env_file()

DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash")
DEFAULT_SYSTEM_PROMPT = os.getenv("DEEPSEEK_SYSTEM_PROMPT", "You are a helpful assistant.")


class DeepSeekError(RuntimeError):
    """Raised when DeepSeek returns an error response."""


def ask_deepseek(
    prompt: str,
    *,
    api_key: str | None = None,
    model: str = DEEPSEEK_MODEL,
    base_url: str = DEEPSEEK_BASE_URL,
    system_prompt: str = DEFAULT_SYSTEM_PROMPT,
    temperature: float = 0.7,
    max_tokens: int = 1024,
) -> str:
    """Send a single user prompt to DeepSeek and return the assistant text."""
    key = api_key or os.getenv("DEEPSEEK_API_KEY")
    if not key:
        raise DeepSeekError(
            "Set DEEPSEEK_API_KEY in .env or export it before calling DeepSeek API."
        )

    payload: dict[str, Any] = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }

    endpoint = base_url.rstrip("/") + "/chat/completions"
    http_request = request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with request.urlopen(http_request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise DeepSeekError(f"DeepSeek API error {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise DeepSeekError(f"Could not reach DeepSeek API: {exc.reason}") from exc

    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise DeepSeekError(f"Unexpected DeepSeek API response: {data}") from exc


def main() -> int:
    """CLI entry point: python deepseek_client.py 'your question'."""
    prompt = " ".join(sys.argv[1:]).strip()
    if not prompt:
        print("Usage: python deepseek_client.py 'your question'", file=sys.stderr)
        return 2

    try:
        print(ask_deepseek(prompt))
    except DeepSeekError as exc:
        print(exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
