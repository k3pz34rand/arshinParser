import gzip
import json
import sys
import urllib.error
import urllib.request
import zlib

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

API = "https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select"

# --- критерии поиска ---
YEAR = "2026"
MITNUMBER = "24319-05"   # тип СИ (в Аршине это "номер в госреестре")
MI_NUMBER = "185357"     # заводской номер

# Поля, которые хотим получить
FIELDS = ("vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,"
          "mi.modification,mi.number,verification_date,valid_date,"
          "applicability,result_docnum,sticker_num")

# Solr-запрос: оба условия через AND, значение mitnumber в кавычках —
# потому что дефис внутри значения
q = f'mi.number:{MI_NUMBER} AND mi.mitnumber:"{MITNUMBER}"'

# fq — только год (проверено: второй fq игнорируется)
params = [
    ("fq", f"verification_year:{YEAR}"),
    ("q", q),
    ("fl", FIELDS),
    ("sort", "verification_date desc,org_title asc"),
    ("rows", "20"),
    ("start", "0"),
]

# Собираем URL вручную, чтобы `:` и `*` не закодировались в %3A / %2A
def build_url(base, pairs):
    parts = []
    for k, v in pairs:
        # заменяем проблемные символы вручную — так же, как браузер
        v_enc = (v.replace(" ", "+")
                  .replace('"', "%22"))
        parts.append(f"{k}={v_enc}")
    return f"{base}?{'&'.join(parts)}"


url = build_url(API, params)
print("URL:", url)
print("-" * 70)

headers = {
    "User-Agent": ("Mozilla/5.0 (Linux; Android 15; Pixel 9) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/153.0.0.0 Mobile Safari/537.36"),
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

req = urllib.request.Request(url, headers=headers, method="GET")

try:
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read()
        if (resp.headers.get("Content-Encoding") or "").lower() == "gzip":
            raw = gzip.decompress(raw)
        elif (resp.headers.get("Content-Encoding") or "").lower() == "deflate":
            try:
                raw = zlib.decompress(raw)
            except zlib.error:
                raw = zlib.decompress(raw, -zlib.MAX_WBITS)
        text = raw.decode("utf-8", errors="replace")
        print("HTTP:", resp.status, resp.reason)
except urllib.error.HTTPError as e:
    body = e.read()
    try:
        if (e.headers.get("Content-Encoding") or "").lower() == "gzip":
            body = gzip.decompress(body)
    except Exception:
        pass
    print(f"HTTP {e.code} {e.reason}")
    print(body.decode("utf-8", errors="replace"))
    sys.exit(1)
except urllib.error.URLError as e:
    print("Network error:", e.reason)
    sys.exit(2)

print("-" * 70)

data = json.loads(text)
resp_block = data.get("response", {})
docs = resp_block.get("docs", [])

print(f"Найдено записей: {resp_block.get('numFound')}, "
      f"получено: {len(docs)}")
print()

for i, d in enumerate(docs, 1):
    print(f"--- Запись {i} ---")
    print(f"  vri_id:            {d.get('vri_id')}")
    print(f"  Тип СИ (mitnumber): {d.get('mi.mitnumber')}")
    print(f"  Название:          {d.get('mi.mititle')}")
    print(f"  Модификация:       {d.get('mi.modification')}")
    print(f"  Заводской №:       {d.get('mi.number')}")
    print(f"  Поверка:           {d.get('verification_date')}")
    print(f"  Действительна до:  {d.get('valid_date')}")
    print(f"  Пригоден:          {d.get('applicability')}")
    print(f"  Свидетельство:     {d.get('result_docnum')}")
    print(f"  Организация:       {d.get('org_title')}")
    print()