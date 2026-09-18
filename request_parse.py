"""
Разбор СЫРОГО ответа карточки поверки (из arshin_details.get_vri_details).
Про HTTP и поиск ничего не знает — можно гонять на сохранённом JSON.

Использование:

    from arshin_parse import parse_vri_details

    det = parse_vri_details("1-527477154", payload)
    print(det.mitype_type, det.vrf_date, det.valid_date)
    for eta in det.mi_eta:
        print(eta.notation, eta.rank_title)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# ============================================================
# Модели
# ============================================================
@dataclass
class Mieta:
    """Средство поверки (эталон) из блока means.mieta[]."""
    reg_number: str | None = None
    mitype_number: str | None = None
    mitype_title: str | None = None
    notation: str | None = None
    modification: str | None = None
    manufacture_num: str | None = None
    manufacture_year: int | None = None
    rank_code: str | None = None
    rank_title: str | None = None
    raw: dict[str, Any] = field(default_factory=dict)


@dataclass
class VriDetails:
    """
    Разобранный ответ /iaux/vri/{vri_id}.

    Самые ходовые поля; полный сырой payload всегда в .raw.
    Для групповых поверок (multiMI) в single-полях лежит первый СИ.
    """
    vri_id: str
    mitype_number: str | None = None       # 24319-05
    mitype_type: str | None = None         # "СВ-15"
    mitype_title: str | None = None        # "Счетчики холодной и горячей воды..."
    modification: str | None = None        # СВ-15Х
    manufacture_num: str | None = None     # 185357
    manufacture_year: int | None = None    # 2007

    organization: str | None = None        # поверитель
    sign_cipher: str | None = None         # ДЕЗ
    mi_owner: str | None = None            # физ.лицо / юр.лицо
    vrf_date: str | None = None            # 08.01.2026
    valid_date: str | None = None          # 07.01.2032
    doc_title: str | None = None           # МИ 1592-2015
    vri_type: str | None = None            # "2"
    cert_num: str | None = None            # С-ДЕЗ/08-01-2026/527477154

    mi_eta: list[Mieta] = field(default_factory=list)
    published_at: str | None = None        # vriVerDateStart

    raw: dict[str, Any] = field(default_factory=dict)


# ============================================================
# Разбор
# ============================================================
def parse_vri_details(vri_id: str, payload: dict[str, Any]) -> VriDetails:
    """
    Превращает сырой payload в VriDetails.
    Безопасен к отсутствующим блокам — вернёт None-поля, не упадёт.
    """
    result = payload.get("result") or {}
    mi_info = result.get("miInfo") or {}

    single = mi_info.get("singleMI") or {}
    if not single:
        multi = (mi_info.get("multiMI") or {}).get("mi") or []
        single = multi[0] if multi else {}

    vri = result.get("vriInfo") or {}
    applicable = vri.get("applicable") or {}
    means = result.get("means") or {}
    publication = result.get("publication") or {}

    return VriDetails(
        vri_id=vri_id,
        mitype_number=single.get("mitypeNumber"),
        mitype_type=single.get("mitypeType"),
        mitype_title=single.get("mitypeTitle"),
        modification=single.get("modification"),
        manufacture_num=single.get("manufactureNum"),
        manufacture_year=single.get("manufactureYear"),
        organization=vri.get("organization"),
        sign_cipher=vri.get("signCipher"),
        mi_owner=vri.get("miOwner"),
        vrf_date=vri.get("vrfDate"),
        valid_date=vri.get("validDate"),
        doc_title=vri.get("docTitle"),
        vri_type=vri.get("vriType"),
        cert_num=applicable.get("certNum"),
        mi_eta=[
            Mieta(
                reg_number=x.get("regNumber"),
                mitype_number=x.get("mitypeNumber"),
                mitype_title=x.get("mitypeTitle"),
                notation=x.get("notation"),
                modification=x.get("modification"),
                manufacture_num=x.get("manufactureNum"),
                manufacture_year=x.get("manufactureYear"),
                rank_code=x.get("rankCode"),
                rank_title=x.get("rankTitle"),
                raw=x,
            )
            for x in (means.get("mieta") or [])
        ],
        published_at=publication.get("vriVerDateStart"),
        raw=payload,
    )


def parse_many_vri_details(
    pairs: list[tuple[str, dict[str, Any]]],
) -> list[VriDetails]:
    """Удобно разобрать сразу пачку (vri_id, payload)."""
    return [parse_vri_details(vid, payload) for vid, payload in pairs]