import os, tempfile, time, json, requests, gspread, pytz
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import CellFormat, TextFormat, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

# ════════════════════════════════════════════════════════════════════
#  КОНФИГ — из переменных окружения (GitHub Secrets)
# ════════════════════════════════════════════════════════════════════

WB_TOKEN       = os.environ["WB_TOKEN"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID_2"]

# credentials.json восстанавливается из секрета GOOGLE_JSON
_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(os.environ["GOOGLE_JSON"])
_tmp.flush()
CREDENTIALS_PATH = _tmp.name

SHEET_TITLE      = "Автоматизация_РекламаWB"
PAUSE_SEC        = 0.4   # пауза между запросами к fullstats

URL_COUNT   = "https://advert-api.wildberries.ru/adv/v1/promotion/count"
URL_ADVERTS = "https://advert-api.wildberries.ru/api/advert/v2/adverts"
URL_STATS   = "https://advert-api.wildberries.ru/adv/v3/fullstats"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Колонки строго по ТЗ ─────────────────────────────────────────────

BID_TYPE = {
    "unified":  "Единая ставка",
    "fix":      "Единая ставка",
    "cpm":      "Единая ставка",
    "manual":   "Ручная ставка",
    "cpc":      "Ручная ставка",
    "1":        "Единая ставка",
    "2":        "Ручная ставка",
}

# ════════════════════════════════════════════════════════════════════
#  ШАГ 1: СПИСОК КАМПАНИЙ
#
#  Используем /adv/v1/promotion/count — он содержит advert_list с ID
#  (рабочий метод из Notebook).
#  Дополнительно /api/advert/v2/adverts для названий и bid_type.
# ════════════════════════════════════════════════════════════════════

def fetch_advert_list() -> List[Dict]:
    print("\n" + "─"*60)
    print("  ШАГ 1: Список кампаний")
    print("─"*60)

    # ── Шаг 1а: ID и статусы из /promotion/count ────────────────────
    r, err = safe_get(URL_COUNT, timeout=60)
    if r is None or r.status_code != 200:
        print(f"  ✗ /count: HTTP {r.status_code if r else err}")
        return []

    try:
        raw = r.json()
    except Exception as e:
        print(f"  ✗ JSON: {e}"); return []

    # Собираем {advertId: status} из advert_list
    id_status: Dict[int, int] = {}
    for entry in (raw.get("adverts") or []):
        status = entry.get("status")
        for adv in (entry.get("advert_list") or []):
            cid = adv.get("advertId")
            if cid and isinstance(status, int):
                id_status[int(cid)] = status

    print(f"  /count → {len(id_status)} кампаний")

    # Фильтр: только Активна (9) и На паузе (11)
    filtered = {cid: st for cid, st in id_status.items() if st in (9, 11)}
    print(f"  После фильтра (9+11): {len(filtered)} кампаний")

    # ── Шаг 1б: названия и bid_type из /api/advert/v2/adverts ───────
    print(f"  GET {URL_ADVERTS} ...", end="", flush=True)
    r2, err2 = safe_get(URL_ADVERTS, timeout=60)
    print(f" HTTP {r2.status_code if r2 else err2[:60]}")

    name_map: Dict[int, Dict] = {}   # {advertId: {name, type_str}}
    if r2 and r2.status_code == 200:
        try:
            raw2 = r2.json()
            for item in (raw2.get("adverts") or []):
                cid  = item.get("id")
                if not cid: continue
                settings = item.get("settings") or {}
                bid_type = item.get("bid_type") or ""
                name_map[int(cid)] = {
                    "name":     str(settings.get("name") or ""),
                    "type_str": BID_TYPE.get(str(bid_type), str(bid_type)),
                }
            print(f"  /adverts → {len(name_map)} кампаний с названиями")
        except Exception as e:
            print(f"  ✗ JSON /adverts: {e}")

    # Собираем итоговый список
    campaigns: List[Dict] = []
    for cid, status in filtered.items():
        info = name_map.get(cid, {})
        campaigns.append({
            "advertId":  cid,
            "name":      info.get("name",     str(cid)),
            "type_str":  info.get("type_str", ""),
            "status":    status,
        })

    campaigns.sort(key=lambda x: x["advertId"])
    print(f"\n  ▶ Итого кампаний для запроса статистики: {len(campaigns)}")
    for c in campaigns[:20]:
        sname = CAMP_STATUS.get(c["status"], str(c["status"]))
        print(f"    {c['advertId']:>10}  {c['type_str']:20}  {sname:15}  «{c['name'][:35]}»")
    if len(campaigns) > 20:
        print(f"    ... и ещё {len(campaigns)-20}")
    return campaigns

# ════════════════════════════════════════════════════════════════════
#  ШАГ 2: СТАТИСТИКА
#  GET /adv/v3/fullstats?ids=ID1,ID2&beginDate=YYYY-MM-DD&endDate=YYYY-MM-DD
#  Актуальный endpoint (v2/fullstats задепрекейчен 2026-03-05)
# ════════════════════════════════════════════════════════════════════

BATCH_SIZE = 50    # макс ID в одном запросе
RATE_PAUSE = 6.0   # пауза между батчами (rate limit ~10 req/min)


def fetch_stats_per_id(campaigns: List[Dict]) -> List[Dict]:
    """
    GET /adv/v3/fullstats
    Параметры: ids=ID1,ID2,...  beginDate=YYYY-MM-DD  endDate=YYYY-MM-DD
    Батчами по 50 ID.
    """
    print("\n" + "─"*60)
    print(f"  ШАГ 2: GET {URL_STATS}")
    print(f"  Период   : {DATE_FROM} — {DATE_TO}")
    print(f"  Кампаний : {len(campaigns)}")
    print("─"*60)

    ids  = [c["advertId"] for c in campaigns]
    cmap = {c["advertId"]: c for c in campaigns}
    agg_map: Dict[int, Dict] = {}

    total_batches = (len(ids) + BATCH_SIZE - 1) // BATCH_SIZE

    for start in range(0, len(ids), BATCH_SIZE):
        batch  = ids[start: start + BATCH_SIZE]
        bn     = start // BATCH_SIZE + 1
        params = {
            "ids":       ",".join(str(x) for x in batch),
            "beginDate": DATE_FROM,
            "endDate":   DATE_TO,
        }

        print(f"\n  Батч {bn}/{total_batches}: {len(batch)} ID  "
              f"{batch[:4]}{'...' if len(batch)>4 else ''}",
              end="", flush=True)

        r, err = safe_get(URL_STATS, params=params, timeout=90)

        if r is None:
            print(f"\n  ✗ Сеть: {err[:120]}")
            time.sleep(RATE_PAUSE); continue

        print(f"  HTTP {r.status_code}", end="")

        if r.status_code == 429:
            print("  → 429 жду 65 с ...")
            time.sleep(65)
            r, err = safe_get(URL_STATS, params=params, timeout=90)
            if r is None:
                time.sleep(RATE_PAUSE); continue
            print(f"  retry HTTP {r.status_code}", end="")

        if r.status_code == 200:
            n = _parse_into(r, agg_map)
            print(f"  → {n} кампаний с данными")
        elif r.status_code == 204:
            print("  → нет данных за период")
        else:
            print(f"\n  Ответ: {r.text[:500]}")

        time.sleep(RATE_PAUSE)

    # Собираем итоговые строки
    raw_rows: List[Dict] = []
    for cid in ids:
        camp = cmap[cid]
        agg  = agg_map.get(cid, {})
        raw_rows.append({
            "advertId":  cid,
            "name":      camp["name"],
            "type_str":  camp["type_str"],
            "status":    camp["status"],
            "ok":        bool(agg),
            "views":     agg.get("views",  0),
            "clicks":    agg.get("clicks", 0),
            "ctr":       agg.get("ctr",    0.0),
            "cpc":       agg.get("cpc",    0.0),
            "cr":        agg.get("cr",     0.0),
            "atbs":      agg.get("atbs",   0),
            "orders":    agg.get("orders", 0),
            "sum":       agg.get("sum",    0.0),
            "shks":      agg.get("shks",   0),
        })

    ok_n = sum(1 for r in raw_rows if r["ok"])
    print(f"\n  ▶ Итого: данные есть по {ok_n}/{len(raw_rows)} кампаний")
    return raw_rows
HEADERS = [
    "ID кампании",             #  1  advertId
    "Кампания",                #  2  name
    "Тип",                     #  3  тип расшифрованный
    "Статус",                  #  4  статус расшифрованный
    "Показы",                  #  5  views
    "Клики",                   #  6  clicks
    "CTR %",                   #  7  ctr
    "CPC, ₽",                  #  8  cpc
    "CR (%)",                  #  9  cr
    "Добавления в корзину",    # 10  atbs
    "Заказанные товары, шт",   # 11  orders
    "Затраты, ₽",              # 12  sum
    "Доля затрат (%)",         # 13  вычисляется: sum / total_sum * 100
    "Отмены, шт",              # 14  shks = отмены (возвраты после выкупа)
]
assert len(HEADERS) == 14, f"Ожидается 14 колонок, сейчас {len(HEADERS)}"

# Расшифровка типов и статусов кампаний WB
CAMP_TYPE = {
    4: "Каталог",
    5: "Карточка товара",
    6: "Поиск + каталог",
    8: "АРК",
    9: "Поиск",
}
CAMP_STATUS = {
    4:  "Готова к запуску",
    7:  "Завершена",
    8:  "Отказана",
    9:  "Активна",
    11: "На паузе",
}

# ════════════════════════════════════════════════════════════════════
#  АВТОВЫЧИСЛЕНИЕ ПЕРИОДА (прошлая полная неделя Пн–Вс, МСК)
# ════════════════════════════════════════════════════════════════════

def _last_monday_msk() -> datetime:
    tz  = pytz.timezone("Europe/Moscow")
    now = datetime.now(tz)
    if now.weekday() == 0 and now.hour < 6:
        now -= timedelta(days=1)
    cur_mon = now - timedelta(days=now.weekday())
    return (cur_mon - timedelta(weeks=1)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )

_MON      = _last_monday_msk()
DATE_FROM = _MON.strftime("%Y-%m-%d")
DATE_TO   = (_MON + timedelta(days=6)).strftime("%Y-%m-%d")

# ════════════════════════════════════════════════════════════════════
#  HTTP
# ════════════════════════════════════════════════════════════════════

def _wh() -> Dict[str, str]:
    return {"Authorization": WB_TOKEN, "Content-Type": "application/json"}

def safe_get(url: str, params: Optional[Dict] = None,
             timeout: int = 60) -> Tuple[Optional[requests.Response], str]:
    try:
        r = requests.get(url, headers=_wh(), params=params, timeout=timeout)
        return r, ""
    except Exception as exc:
        return None, str(exc)

# ════════════════════════════════════════════════════════════════════
#  ШАГ 0: ДИАГНОСТИКА /adv/v1/promotion/count
# ════════════════════════════════════════════════════════════════════

def log_promotion_count() -> None:
    print("\n" + "─"*60)
    print("  ШАГ 0: Диагностика /adv/v1/promotion/count")
    print("─"*60)
    r, err = safe_get(URL_COUNT, timeout=30)
    if r is None:
        print(f"  ✗ {err}"); return

    print(f"  HTTP: {r.status_code}")
    if r.status_code != 200:
        print(f"  Ответ: {r.text[:400]}"); return

    try:
        raw = r.json()
    except Exception as e:
        print(f"  JSON: {e}"); return

    total = 0
    for entry in (raw.get("adverts") or []):
        cnt  = int(entry.get("count", 0) or 0)
        st   = entry.get("status", "?")
        name = CAMP_STATUS.get(st, f"статус {st}")
        if cnt:
            print(f"  {name:22}: {cnt} кампаний")
            total += cnt
    print(f"  {'Итого':22}: {total} кампаний")



def _parse_into(r: requests.Response, agg_map: Dict[int, Dict]) -> int:
    """Парсит ответ fullstats → заполняет agg_map. Возвращает кол-во кампаний."""
    try:
        raw = r.json()
    except Exception as e:
        print(f"\n  JSON: {e}"); return 0

    camps = _to_list(raw)
    if not camps:
        print(f"\n  Пустой ответ: {r.text[:200]}")
        return 0

    if not agg_map:   # диагностика один раз
        _diag_stats(camps)

    for camp in camps:
        cid = camp.get("advertId")
        if not cid: continue
        cid = int(cid)
        agg = _aggregate([camp])
        if cid in agg_map:
            for k in ("views","clicks","sum","atbs","orders","shks"):
                agg_map[cid][k] = agg_map[cid].get(k, 0) + agg.get(k, 0)
            # Пересчитываем производные
            v = agg_map[cid]["views"]; c = agg_map[cid]["clicks"]
            s = agg_map[cid]["sum"]
            agg_map[cid]["ctr"] = round(c/v*100, 2) if v else 0.0
            agg_map[cid]["cpc"] = round(s/c, 2)     if c else 0.0
            agg_map[cid]["cr"]  = round(agg_map[cid]["orders"]/c*100, 2) if c else 0.0
        else:
            agg_map[cid] = agg
    return len(camps)


def _to_list(raw: Any) -> List[Dict]:
    if isinstance(raw, list):
        return [x for x in raw if isinstance(x, dict)]
    if isinstance(raw, dict):
        for k in ("data", "adverts", "campaigns", "items"):
            v = raw.get(k)
            if isinstance(v, list):
                return [x for x in v if isinstance(x, dict)]
    return []


def _diag_stats(camps: List[Dict]) -> None:
    c0   = camps[0]
    days = c0.get("days") or c0.get("statistics") or []
    print(f"\n  [ДИА] Ключи кампании: {list(c0.keys())}")
    if days and isinstance(days[0], dict):
        d0   = days[0]
        apps = d0.get("apps") or []
        print(f"  [ДИА] Ключи дня     : {list(d0.keys())}")
        if apps and isinstance(apps[0], dict):
            print(f"  [ДИА] Ключи app     : {list(apps[0].keys())}")


def _v(x: Any) -> float:
    if x is None: return 0.0
    try: return float(x)
    except: return 0.0


def _aggregate(camps: List[Dict]) -> Dict[str, Any]:
    """Суммирует все дни всех кампаний. apps[] агрегируются в один счётчик."""
    t = {k: 0.0 for k in ("views","clicks","sum","atbs","orders","shks")}

    for camp in camps:
        days = camp.get("days") or camp.get("statistics") or []
        for day in days:
            if not isinstance(day, dict): continue
            apps = day.get("apps") or []
            src  = apps if apps else [day]
            for item in src:
                if not isinstance(item, dict): continue
                for k in t:
                    t[k] += _v(item.get(k))

    v = t["views"]; c = t["clicks"]; s = t["sum"]
    return {
        "views":   int(t["views"]),
        "clicks":  int(t["clicks"]),
        "ctr":     round(c / v * 100, 2) if v else 0.0,
        "cpc":     round(s / c, 2)       if c else 0.0,
        "cr":      round(t["orders"] / c * 100, 2) if c else 0.0,
        "atbs":    int(t["atbs"]),
        "orders":  int(t["orders"]),
        "sum":     round(s, 2),
        "shks":    int(t["shks"]),
    }

# ════════════════════════════════════════════════════════════════════
#  СБОРКА ТАБЛИЦЫ
#  Доля затрат считается как sum_i / total_sum * 100
# ════════════════════════════════════════════════════════════════════

def build_rows(raw_rows: List[Dict]) -> List[List[str]]:
    total_sum = sum(r["sum"] for r in raw_rows if r["ok"])

    result = []
    for r in raw_rows:
        share = round(r["sum"] / total_sum * 100, 2) if total_sum > 0 else 0.0
        # Статус: WB UI показывает "Приостановлена" для "На паузе"
        status_map = {9: "Активна", 11: "Приостановлена",
                      4: "Готова к запуску", 7: "Завершена", 8: "Отказана"}
        status_str = status_map.get(r["status"], str(r["status"]))

        row = [
            str(r["advertId"]),          #  1 ID кампании
            r["name"],                   #  2 Кампания
            r["type_str"],               #  3 Тип (Единая/Ручная ставка)
            status_str,                  #  4 Статус
            str(r["views"]),             #  5 Показы
            str(r["clicks"]),            #  6 Клики
            str(r["ctr"]),               #  7 CTR %
            str(r["cpc"]),               #  8 CPC, ₽
            str(r["cr"]),                #  9 CR (%)
            str(r["atbs"]),              # 10 В корзину
            str(r["orders"]),            # 11 Заказанные товары
            str(r["sum"]),               # 12 Затраты, ₽
            str(share),                  # 13 Доля затрат (%)
            str(r["shks"]),              # 14 Отмены, шт
        ]
        assert len(row) == len(HEADERS), f"row={len(row)} != headers={len(HEADERS)}"
        result.append(row)
    return result

# ════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS
#  ws.update(values=..., range_name=...) — без DeprecationWarning
# ════════════════════════════════════════════════════════════════════

_BLK  = {"red": 0.0,  "green": 0.0,  "blue": 0.0}
_WHT  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_NAVY = {"red": 0.07, "green": 0.21, "blue": 0.38}
_RED  = {"red": 0.8,  "green": 0.0,  "blue": 0.0}

FMT_H = CellFormat(backgroundColor=_NAVY,
                   textFormat=TextFormat(bold=True,  fontSize=10, foregroundColor=_WHT))
FMT_D = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_M = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_E = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=11, foregroundColor=_RED))


def _col(n: int) -> str:
    buf = []
    while n > 0:
        n, r = divmod(n-1, 26); buf.append(chr(65+r))
    return "".join(reversed(buf))


def write_sheet(ss: gspread.Spreadsheet, rows: List[List[str]]) -> None:
    """
    Строка 1  — период и кол-во кампаний
    Строка 5  — заголовки (тёмно-синяя шапка)
    Строка 6+ — данные (чёрный обычный)
    """
    try:
        ws = ss.worksheet(SHEET_TITLE)
    except WorksheetNotFound:
        ws = ss.add_worksheet(
            title=SHEET_TITLE,
            rows=max(5000, len(rows) + 20),
            cols=len(HEADERS) + 3,
        )
    ws.clear()

    n  = len(rows)
    lc = _col(len(HEADERS))
    ok_n = sum(1 for r in rows if any(v not in ("0", "0.0", "") for v in r[4:]))

    ws.update(
        values=[[
            f"Реклама  |  Период: {DATE_FROM} – {DATE_TO}  |  "
            f"Кампаний: {n}  |  С данными: {ok_n}  |  "
            f"Выгружено: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        ]],
        range_name="A1",
    )
    format_cell_range(ws, "A1", FMT_M)

    ws.update(values=[HEADERS], range_name=f"A5:{lc}5")
    format_cell_range(ws, f"A5:{lc}5", FMT_H)

    if n == 0:
        ws.update(values=[["Кампании не найдены — смотри консоль"]], range_name="A6")
        format_cell_range(ws, "A6", FMT_E)
        return

    for i in range(0, n, 500):
        chunk = rows[i: i+500]
        r0, r1 = 6+i, 5+i+len(chunk)
        ws.update(values=chunk, range_name=f"A{r0}:{lc}{r1}")
        time.sleep(0.5)

    format_cell_range(ws, f"A6:{lc}{5+n}", FMT_D)
    print(f"\n  ✓ «{SHEET_TITLE}»: {n} строк × {len(HEADERS)} колонок")

# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

def main() -> None:
    print("\n" + "═"*60)
    print("  WB Реклама  |  Автоматизация_РекламаWB  |  v6.0")
    print(f"  Период  : {DATE_FROM} — {DATE_TO}")
    print(f"  Лист    : «{SHEET_TITLE}»")
    print(f"  Колонок : {len(HEADERS)}")
    print("═"*60)


    print("\nПодключаюсь к Google Sheets ...")
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, SCOPES)
        ss    = gspread.Client(auth=creds).open_by_key(SPREADSHEET_ID)
        print("  ✓ OK")
    except Exception as e:
        print(f"  ✗ {e}"); return

    # ── Шаг 0: диагностика ──────────────────────────────────────────
    log_promotion_count()

    # ── Шаг 1: список кампаний (ID из /count, названия из /adverts) ─
    campaigns = fetch_advert_list()
    if not campaigns:
        print("\n  ✗ Кампаний со статусом «Активна» или «На паузе» не найдено")
        write_sheet(ss, [])
        return

    # ── Шаг 2: статистика — только по отфильтрованным кампаниям ─────
    raw_rows = fetch_stats_per_id(campaigns)

    # ── Сборка ──────────────────────────────────────────────────────
    rows = build_rows(raw_rows)

    # ── Итог в консоль ──────────────────────────────────────────────
    ok_rows = [r for r in raw_rows if r["ok"]]
    print("\n" + "═"*60)
    print(f"  Кампаний найдено : {len(campaigns)}")
    print(f"  С данными        : {len(ok_rows)}")
    if ok_rows:
        print(f"  Показы итого     : {sum(r['views']  for r in ok_rows):,}")
        print(f"  Клики итого      : {sum(r['clicks'] for r in ok_rows):,}")
        print(f"  Затраты итого, ₽ : {sum(r['sum']    for r in ok_rows):,.2f}")
        print(f"  Заказы итого     : {sum(r['orders'] for r in ok_rows):,}")

    # ── Запись ──────────────────────────────────────────────────────
    print(f"\n  Записываю в «{SHEET_TITLE}» ...")
    try:
        write_sheet(ss, rows)
    except Exception as e:
        print(f"  ✗ Ошибка записи: {e}")
        import traceback; traceback.print_exc()
        return

    print("═"*60 + "\n")


main()
