"""Транспорт для API ФГИС «Аршин». Ничего предметного."""

from __future__ import annotations

import gzip
import http.cookiejar
import json
import urllib.error
import urllib.request
import zlib
from typing import Any


DEFAULT_HEADERS: dict[str, str] = {
    "User-Agent": (
        "Mozilla/5.0 (Linux; Android 15; Pixel 9) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/153.0.0.0 Mobile Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Referer": "https://fgis.gost.ru/fundmetrology/cm/results",
    "Origin": "https://fgis.gost.ru",
    "sec-ch-ua": '"Google Chrome";v="153", "Not_A Brand";v="8", "Chromium";v="153"',
    "sec-ch-ua-mobile": "?1",
    "sec-ch-ua-platform": '"Android"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}


class ArshinError(RuntimeError):
    """Любая ошибка при обращении к API."""


_COOKIE_JAR = http.cookiejar.CookieJar()
_OPENER = urllib.request.build_opener(
    urllib.request.HTTPCookieProcessor(_COOKIE_JAR)
)


def encode_value(value: str) -> str:
    return (
        value.replace(" ", "+")
             .replace('"', "%22")
             .replace("#", "%23")
             .replace("&", "%26")
             .replace("?", "%3F")
    )


def build_url(base: str, params: list[tuple[str, str]]) -> str:
    parts = [f"{k}={encode_value(str(v))}" for k, v in params if v is not None]
    return f"{base}?{'&'.join(parts)}"


def _decode_body(raw: bytes, encoding: str | None) -> bytes:
    enc = (encoding or "").lower()
    if enc == "gzip":
        return gzip.decompress(raw)
    if enc == "deflate":
        try:
            return zlib.decompress(raw)
        except zlib.error:
            return zlib.decompress(raw, -zlib.MAX_WBITS)
    return raw


def http_get(url: str, headers: dict[str, str], timeout: float) -> bytes:
    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with _OPENER.open(req, timeout=timeout) as resp:
            raw = resp.read()
            return _decode_body(raw, resp.headers.get("Content-Encoding"))
    except urllib.error.HTTPError as e:
        body = e.read()
        try:
            body = _decode_body(body, e.headers.get("Content-Encoding"))
        except Exception:
            pass
        snippet = body.decode("utf-8", errors="replace")[:500]
        raise ArshinError(f"HTTP {e.code} {e.reason}: {snippet}") from e
    except urllib.error.URLError as e:
        raise ArshinError(f"Network error: {e.reason}") from e


def http_get_json(
    url: str,
    *,
    headers: dict[str, str] | None = None,
    timeout: float = 30.0,
) -> Any:
    hdrs = dict(DEFAULT_HEADERS)
    if headers:
        hdrs.update(headers)
    body = http_get(url, hdrs, timeout)
    try:
        return json.loads(body.decode("utf-8", errors="replace"))
    except json.JSONDecodeError as e:
        raise ArshinError(f"Не удалось разобрать JSON: {e}") from e