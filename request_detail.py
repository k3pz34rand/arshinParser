"""
Второй запрос: карточка поверки /fundmetrology/cm/iaux/vri/{vri_id}.
Отдаёт сырой payload. Разбор — в request_parse.py.
"""

from __future__ import annotations

import urllib.parse
from typing import Any

from request_http import ArshinError, http_get_json


DETAILS_API_URL = "https://fgis.gost.ru/fundmetrology/cm/iaux/vri/{vri_id}"


def get_vri_details(
    vri_id: str,
    *,
    timeout: float = 30.0,
    headers: dict[str, str] | None = None,
) -> dict[str, Any]:
    url = DETAILS_API_URL.format(vri_id=urllib.parse.quote(str(vri_id), safe=""))

    hdrs = {"Referer": f"https://fgis.gost.ru/fundmetrology/cm/results/{vri_id}"}
    if headers:
        hdrs.update(headers)

    payload = http_get_json(url, headers=hdrs, timeout=timeout)
    if not isinstance(payload, dict) or "result" not in payload:
        raise ArshinError(f"Неожиданный ответ карточки: {payload!r}")
    return payload