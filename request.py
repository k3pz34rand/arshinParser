"""
Первый запрос: поиск ФГИС «Аршин»
    GET /fundmetrology/cm/xcdb/vri/select

Плюс конвейер: для каждого найденного doc берём doc["vri_id"]
и дёргаем карточку из arshin_details. Разбор — в arshin_parse.

Использование:

    from arshin import search_arshin, fetch_details_for_docs
    from arshin_parse import parse_vri_details

    res = search_arshin(
        verification_year="2026",
        mitnumber="24319-05",
        mi_number="185357",
    )
    for doc, payload in fetch_details_for_docs(res.docs):
        det = parse_vri_details(doc["vri_id"], payload)
        print(det.vrf_date, det.valid_date, det.organization)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable

from request_http import (
    ArshinError,
    DEFAULT_HEADERS,
    build_url,
    http_get_json,
)
from request_detail import get_vri_details


API_URL = "https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select"

DEFAULT_FIELDS = (
    "vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,"
    "mi.modification,mi.number,verification_date,valid_date,"
    "applicability,result_docnum,sticker_num"
)

DEFAULT_SORT = "verification_date desc,org_title asc"


__all__ = [
    "API_URL",
    "DEFAULT_FIELDS",
    "DEFAULT_SORT",
    "SearchResult",
    "ArshinError",
    "search_arshin_raw",
    "search_arshin",
    "search_all_pages",
    "iter_search_arshin",
    "fetch_details_for_docs",
    "search_with_details",
    "get_vri_details",
]


# ============================================================
# Результат поиска
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


# ============================================================
# Первый запрос: поиск
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
        for f in extra_fq:
            params.append(("fq", f))

    params.extend([
        ("q", q),
        ("fl", fl),
        ("sort", sort),
        ("rows", str(rows)),
        ("start", str(start)),
    ])

    url = build_url(API_URL, params)
    payload = http_get_json(url, headers=headers, timeout=timeout)

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
    """
    if page_size > 100:
        page_size = 100

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


def iter_search_arshin(
    *,
    page_size: int = 100,
    start: int = 0,
    **kwargs: Any,
) -> Iterable[dict[str, Any]]:
    """Ленивый генератор по страницам поиска."""
    if page_size > 100:
        page_size = 100

    while True:
        res = search_arshin(rows=page_size, start=start, **kwargs)
        if not res.docs:
            return
        yield from res.docs
        if start + page_size >= res.num_found:
            return
        start += page_size


# ============================================================
# Конвейер: первый дёргает второй
# ============================================================
def fetch_details_for_docs(
    docs: Iterable[dict[str, Any]],
    *,
    vri_id_key: str = "vri_id",
    timeout: float = 30.0,
    headers: dict[str, str] | None = None,
) -> Iterable[tuple[dict[str, Any], dict[str, Any]]]:
    """
    Для каждого doc из поиска берёт doc["vri_id"] и дёргает карточку.
    Отдаёт пары (doc, details_payload), где details_payload — СЫРОЙ ответ
    2-го запроса (dict). Разбором занимается arshin_parse.

    Пример:
        for doc, payload in fetch_details_for_docs(res.docs):
            det = parse_vri_details(doc["vri_id"], payload)
    """
    for doc in docs:
        vid = doc.get(vri_id_key)
        if not vid:
            continue
        yield doc, get_vri_details(vid, timeout=timeout, headers=headers)


def search_with_details(
    *,
    rows: int = 20,
    start: int = 0,
    timeout: float = 30.0,
    **search_kwargs: Any,
) -> list[tuple[dict[str, Any], dict[str, Any]]]:
    """
    Один вызов: делает поиск, потом по каждому vri_id тянет карточку.
    Возвращает список пар (search_doc, details_payload).

    Пример:
        for doc, payload in search_with_details(
            verification_year="2026",
            mitnumber="24319-05",
            mi_number="185357",
        ):
            ...
    """
    res = search_arshin(rows=rows, start=start, timeout=timeout, **search_kwargs)
    return list(fetch_details_for_docs(res.docs, timeout=timeout))