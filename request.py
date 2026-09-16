"""
Клиент для API ФГИС «Аршин»: /fundmetrology/cm/xcdb/vri/select

Использование из другого скрипта:

    from arshin import search_arshin, search_arshin_raw

    result = search_arshin(
        verification_year="2026",
        mitnumber="24319-05",
        mi_number="185357",
    )
    for doc in result.docs:
        print(doc["vri_id"], doc["mi.number"])
"""

from __future__ import annotations

import gzip
import json
import urllib.error
import urllib.request
import zlib
from dataclasses import dataclass, field
from typing import Any, Iterable


API_URL = "https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select"

DEFAULT_FIELDS = (
    "vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,"
    "mi.modification,mi.number,verification_date,valid_date,"
    "applicability,result_docnum,sticker_num"
)

DEFAULT_SORT = "verification_date desc,org_title asc"

DEFAULT_HEADERS = {
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


# ============================================================
# Результат
# ============================================================
@dataclass
class SearchResult:
    num_found: int
    start: int
    docs: list[dict[str, Any]]
    qtime: int | None = None
    raw: dict[str, Any] = field(default_factory=dict)

    def __bool__(self) -> bool:
        return bool(self.docs)

    def __len__(self) -> int:
        return len(self.docs)

    def __iter__(self) -> Iterable[dict[str, Any]]:
        return iter(self.docs)


class ArshinError(RuntimeError):
    """Любая ошибка при обращении к API (сеть, HTTP, разбор ответа)."""


# ============================================================
# Внутренние помощники
# ============================================================
def _encode_value(value: str) -> str:
    """Кодируем только то, что реально нужно. ':' и '*' оставляем как есть."""
    return (
        value.replace(" ", "+")
             .replace('"', "%22")
             .replace("#", "%23")
             .replace("&", "%26")
             .replace("?", "%3F")
    )


def _build_url(base: str, params: list[tuple[str, str]]) -> str:
    parts = [f"{k}={_encode_value(str(v))}" for k, v in params if v is not None]
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


def _http_get(url: str, headers: dict[str, str], timeout: float) -> bytes:
    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
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


# ============================================================
# Публичные функции
# ============================================================
def search_arshin_raw(
    *,
    q: str = "*",
    fq: str | None = None,
    fl: str = DEFAULT_FIELDS,
    sort: str = DEFAULT_SORT,
    rows: int = 20,
    start: int = 0,
    timeout: float = 30.0,
    extra_fq: list[str] | None = None,
    headers: dict[str, str] | None = None,
) -> SearchResult:
    """
    Низкоуровневый вызов: сам формируешь `q` и `fq`.

    ВАЖНО: сервер принимает ТОЛЬКО ОДИН `fq`. Если передать `extra_fq`
    больше одного значения — он будет молча проигнорирован (кроме первого).
    Основной фильтр по году кладём в `fq`, всё остальное — в `q`.
    """
    params: list[tuple[str, str]] = []
    if fq:
        params.append(("fq", fq))
    if extra_fq:
        # на практике работает только первый, добавляем на случай будущих изменений API
        for f in extra_fq:
            params.append(("fq", f))

    params.extend([
        ("q", q),
        ("fl", fl),
        ("sort", sort),
        ("rows", str(rows)),
        ("start", str(start)),
    ])

    url = _build_url(API_URL, params)
    hdrs = dict(DEFAULT_HEADERS)
    if headers:
        hdrs.update(headers)

    body = _http_get(url, hdrs, timeout)

    try:
        payload = json.loads(body.decode("utf-8", errors="replace"))
    except json.JSONDecodeError as e:
        raise ArshinError(f"Не удалось разобрать JSON: {e}") from e

    resp = payload.get("response") or {}
    header = payload.get("responseHeader") or {}

    return SearchResult(
        num_found=int(resp.get("numFound", 0)),
        start=int(resp.get("start", 0)),
        docs=list(resp.get("docs") or []),
        qtime=header.get("QTime"),
        raw=payload,
    )


def _solr_quote(value: str) -> str:
    """Оборачиваем значение в кавычки, если внутри есть пробелы/дефисы/двоеточия."""
    if any(ch in value for ch in ' -:"'):
        return f'"{value}"'
    return value


def search_arshin(
    *,
    verification_year: str | int | None = None,
    mitnumber: str | None = None,
    mi_number: str | None = None,
    extra_q: str | None = None,
    fl: str = DEFAULT_FIELDS,
    sort: str = DEFAULT_SORT,
    rows: int = 20,
    start: int = 0,
    timeout: float = 30.0,
) -> SearchResult:
    """
    Удобный вызов по конкретным критериям.

    Все параметры необязательные, но хотя бы один из
    (verification_year, mitnumber, mi_number, extra_q) должен быть задан.

    Пример:
        search_arshin(verification_year=2026, mitnumber="24319-05", mi_number="185357")
    """
    if not any([verification_year, mitnumber, mi_number, extra_q]):
        raise ValueError("Нужно задать хотя бы один критерий поиска")

    q_parts: list[str] = []
    if mi_number:
        q_parts.append(f"mi.number:{_solr_quote(str(mi_number))}")
    if mitnumber:
        q_parts.append(f"mi.mitnumber:{_solr_quote(str(mitnumber))}")
    if extra_q:
        q_parts.append(f"({extra_q})")

    q = " AND ".join(q_parts) if q_parts else "*"

    fq = f"verification_year:{verification_year}" if verification_year else None

    return search_arshin_raw(
        q=q, fq=fq, fl=fl, sort=sort, rows=rows, start=start, timeout=timeout,
    )


def search_all_pages(
    *,
    page_size: int = 100,
    max_records: int | None = None,
    **kwargs: Any,
) -> list[dict[str, Any]]:
    """
    Пробегает по всем страницам и возвращает единый список docs.
    Осторожно: total может быть очень большим — ставь max_records.

    Все прочие kwargs пробрасываются в search_arshin().
    """
    if page_size > 100:
        page_size = 100  # Solr обычно не отдаёт больше

    collected: list[dict[str, Any]] = []
    start = 0
    while True:
        res = search_arshin(rows=page_size, start=start, **kwargs)
        collected.extend(res.docs)
        if max_records and len(collected) >= max_records:
            return collected[:max_records]
        if len(res.docs) < page_size or start + page_size >= res.num_found:
            return collected
        start += page_size