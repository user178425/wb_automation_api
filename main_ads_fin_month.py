import os, tempfile

import time, requests, gspread, pytz
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import CellFormat, TextFormat, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from calendar import monthrange
from typing import Any, Dict, List, Optional, Tuple

# ════════════════════════════════════════════════════════════════════
#  КОНФИГУРАЦИИ
# ════════════════════════════════════════════════════════════════════

WB_TOKEN       = os.environ["WB_TOKEN"]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
# credentials.json восстанавливается из секрета GOOGLE_JSON
_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(os.environ["GOOGLE_JSON"])
_tmp.flush()
CREDENTIALS_PATH = _tmp.name
SHEET_TITLE      = "Автоматизация_реклама_фин"

URL_COUNT   = "https://advert-api.wildberries.ru/adv/v1/promotion/count"
URL_ADVERTS = "https://advert-api.wildberries.ru/api/advert/v2/adverts"
URL_UPD     = "https://advert-api.wildberries.ru/adv/v1/upd"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ════════════════════════════════════════════════════════════════════
#  ─── БЛОК ДАТ ────────────────────────────────────────────────────
#
#  По умолчанию: прошлый полный месяц (автовычисление).
#  Чтобы задать период вручную — раскомментируй РУЧНОЙ ВВОД
#  и закомментируй строки АВТОВЫЧИСЛЕНИЕ.
#
# ─────────────────────────────────────────────────────────────────
#  РУЧНОЙ ВВОД (раскомментируй и заполни):
# ─────────────────────────────────────────────────────────────────
# YEAR  = 2026
# MONTH = 2    # 2 = февраль
# ─────────────────────────────────────────────────────────────────
#  АВТОВЫЧИСЛЕНИЕ (прошлый полный месяц):
# ─────────────────────────────────────────────────────────────────
_tz  = pytz.timezone("Europe/Moscow")
_now = datetime.now(_tz)

if _now.month == 1:
    YEAR, MONTH = _now.year - 1, 12
else:
    YEAR, MONTH = _now.year, _now.month - 1

# ─── Даты из выбранного месяца ────────────────────────────────────
_last_day    = monthrange(YEAR, MONTH)[1]
DATE_FROM    = f"{YEAR:04d}-{MONTH:02d}-01"
DATE_TO      = f"{YEAR:04d}-{MONTH:02d}-{_last_day:02d}"
DATE_FROM_RU = f"01.{MONTH:02d}.{YEAR:04d}"
DATE_TO_RU   = f"{_last_day:02d}.{MONTH:02d}.{YEAR:04d}"

# ════════════════════════════════════════════════════════════════════
#  ЗАГОЛОВКИ — 7 колонок строго по ТЗ
# ════════════════════════════════════════════════════════════════════

HEADERS = [
    "ID кампании",        #  1  advertId
    "Кампания",           #  2  campName
    "Раздел",             #  3  advertType → расшифровка
    "Дата списания",      #  4  updTime
    "Источник списания",  #  5  paymentType
    "Сумма",              #  6  updSum
    "Номер документа",    #  7  updNum
]
assert len(HEADERS) == 7

# ════════════════════════════════════════════════════════════════════
#  СПРАВОЧНИКИ
# ════════════════════════════════════════════════════════════════════

CAMP_STATUS = {
    4:  "Готова к запуску",
    7:  "Завершена",
    8:  "Отказана",
    9:  "Активна",
    11: "На паузе",
}

# advertType → название раздела
ADV_TYPE = {
    4: "Каталог",
    5: "Карточка товара",
    6: "Поиск + каталог",
    8: "АРК",
    9: "Поиск",
}

# bid_type → Единая/Ручная ставка
BID_TYPE = {
    # Реальные значения bid_type из API /api/advert/v2/adverts
    "unified":  "Единая ставка",
    "fix":      "Единая ставка",
    "cpm":      "Единая ставка",
    "manual":   "Ручная ставка",
    "cpc":      "Ручная ставка",
    # Числовые варианты (на всякий случай)
    "1":        "Единая ставка",
    "2":        "Ручная ставка",
}

# Ключевые слова в paymentType, которые означают ПОПОЛНЕНИЕ (исключаем)
# Всё остальное (Баланс, Счёт, Промо бонусы и т.п.) — это списания → оставляем
EXCLUDE_PAYMENT_KEYWORDS = [
    "пополнени",   # «Пополнение баланса», «Пополнение счёта» и т.д.
]

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
#  ШАГ 0: ДИАГНОСТИКА
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

# ════════════════════════════════════════════════════════════════════
#  ШАГ 1: СПИСОК КАМПАНИЙ
#  Нужен для обогащения транзакций: ставка (Единая/Ручная)
#  берётся из bid_type кампании, т.к. в /upd её нет
# ════════════════════════════════════════════════════════════════════

def fetch_advert_list() -> Dict[int, Dict]:
    """Возвращает {advertId: {name, type_str}} для всех кампаний."""
    print("\n" + "─"*60)
    print("  ШАГ 1: Список кампаний")
    print("─"*60)

    # ID и статусы из /promotion/count
    r, err = safe_get(URL_COUNT, timeout=60)
    if r is None or r.status_code != 200:
        print(f"  ✗ /count: HTTP {r.status_code if r else err}")
        return {}
    try:
        raw = r.json()
    except Exception as e:
        print(f"  ✗ JSON: {e}"); return {}

    id_status: Dict[int, int] = {}
    for entry in (raw.get("adverts") or []):
        status = entry.get("status")
        for adv in (entry.get("advert_list") or []):
            cid = adv.get("advertId")
            if cid and isinstance(status, int):
                id_status[int(cid)] = status

    print(f"  /count → {len(id_status)} кампаний")

    # Названия и bid_type из /api/advert/v2/adverts
    print(f"  GET {URL_ADVERTS} ...", end="", flush=True)
    r2, err2 = safe_get(URL_ADVERTS, timeout=60)
    print(f" HTTP {r2.status_code if r2 else err2[:60]}")

    result: Dict[int, Dict] = {}
    if r2 and r2.status_code == 200:
        try:
            raw2 = r2.json()
            adverts_list = raw2.get("adverts") or []

            for item in adverts_list:
                cid = item.get("id")
                if not cid: continue
                settings = item.get("settings") or {}
                # bid_type: "unified" / "manual" / "cpm" / "cpc"
                # settings.payment_type: "cpm" / "cpc" — запасной вариант
                bid_type = (
                    item.get("bid_type") or
                    settings.get("payment_type") or
                    ""
                )
                result[int(cid)] = {
                    "name":     str(settings.get("name") or ""),
                    "type_str": BID_TYPE.get(str(bid_type), ""),
                }
            print(f"  /adverts → {len(result)} кампаний с названиями")
        except Exception as e:
            print(f"  ✗ JSON /adverts: {e}")

    # Добиваем кампании которых нет в /adverts (только ID из /count)
    for cid in id_status:
        if cid not in result:
            result[cid] = {"name": str(cid), "type_str": ""}

    print(f"  ▶ Итого кампаний: {len(result)}")
    return result

# ════════════════════════════════════════════════════════════════════
#  ШАГ 2: ИСТОРИЯ СПИСАНИЙ
# ════════════════════════════════════════════════════════════════════

def _fmt_date(raw: str) -> str:
    """ISO → YYYY-MM-DD HH:MM  (например 2026-02-01 23:59)."""
    if not raw: return ""
    try:
        s = raw[:19]  # "2026-02-28T23:59:59"
        dt = datetime.fromisoformat(s)
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return raw

def _fmt_sum(v: Any) -> str:
    if v is None: return "0"
    try:
        f = round(float(v), 2)
        return str(int(f)) if f == int(f) else str(f)
    except Exception:
        return str(v)

def _is_topup(payment_type: str) -> bool:
    """True если это пополнение — такие строки исключаем."""
    pt_lower = payment_type.lower()
    return any(kw in pt_lower for kw in EXCLUDE_PAYMENT_KEYWORDS)


def fetch_finance(camp_map: Dict[int, Dict]) -> List[Dict]:
    print("\n" + "─"*60)
    print(f"  ШАГ 2: GET {URL_UPD}  (один запрос для всего аккаунта)")
    print(f"  Период: {DATE_FROM_RU} — {DATE_TO_RU}")
    print("─"*60)

    params = {
        "from": DATE_FROM,
        "to":   DATE_TO,
        # БЕЗ параметра id — возвращает все транзакции аккаунта
    }

    r, err = safe_get(URL_UPD, params=params, timeout=90)

    if r is None:
        print(f"  ✗ Сеть: {err}"); return []

    print(f"  HTTP {r.status_code}")

    if r.status_code == 429:
        print("  → 429 жду 65 с ...")
        time.sleep(65)
        r, err = safe_get(URL_UPD, params=params, timeout=90)
        if r is None:
            print(f"  ✗ {err}"); return []
        print(f"  retry HTTP {r.status_code}")

    if r.status_code == 204:
        print("  (нет данных за период — 204)"); return []

    if r.status_code != 200:
        print(f"  ✗ HTTP {r.status_code}: {r.text[:400]}"); return []

    try:
        raw = r.json()
    except Exception as e:
        print(f"  ✗ JSON: {e}"); return []

    # Нормализуем в список
    items: List[Dict] = []
    if isinstance(raw, list):
        items = [x for x in raw if isinstance(x, dict)]
    elif isinstance(raw, dict):
        for key in ("data", "upds", "updates", "items"):
            v = raw.get(key)
            if isinstance(v, list):
                items = [x for x in v if isinstance(x, dict)]
                break

    print(f"  Всего транзакций в ответе: {len(items)}")

    # Собираем уникальные paymentType для справки
    payment_types = sorted(set(str(i.get("paymentType", "")) for i in items if i.get("paymentType")))
    print(f"  Все типы paymentType: {payment_types}")

    all_rows: List[Dict] = []
    skipped  = 0

    for item in items:
        payment_type = str(item.get("paymentType") or "")

        # Пропускаем пополнения
        if _is_topup(payment_type):
            skipped += 1
            continue

        advert_id   = int(item.get("advertId") or 0)
        camp_name   = str(item.get("campName") or "")
        adv_type    = int(item.get("advertType") or 0)
        upd_time    = str(item.get("updTime") or "")
        upd_sum     = item.get("updSum")
        upd_num     = item.get("updNum")

        # Раздел = Единая ставка / Ручная ставка — из справочника кампаний
        camp_info = camp_map.get(advert_id, {})
        type_str  = camp_info.get("type_str", "")
        razdel    = type_str  # именно ставка идёт в колонку «Раздел»

        # Если campName пришёл пустым — берём из справочника
        if not camp_name:
            camp_name = camp_info.get("name", str(advert_id))

        all_rows.append({
            "advert_id":    advert_id,
            "camp_name":    camp_name,
            "razdel":       razdel,
            "upd_time":     _fmt_date(upd_time),
            "payment_type": payment_type,
            "upd_sum":      _fmt_sum(upd_sum),
            "upd_num":      str(upd_num) if upd_num is not None else "",
        })

    print(f"  Пополнений исключено  : {skipped}")
    print(f"  ▶ Транзакций к записи : {len(all_rows)}")
    return all_rows

# ════════════════════════════════════════════════════════════════════
#  СБОРКА СТРОК — порядок строго по ТЗ
# ════════════════════════════════════════════════════════════════════

def build_rows(finance: List[Dict]) -> List[List[str]]:
    result = []
    for r in finance:
        row = [
            str(r["advert_id"]),   #  1 ID кампании
            r["camp_name"],        #  2 Кампания
            r["razdel"],           #  3 Раздел  (advertType → Поиск/АРК/…)
            r["upd_time"],         #  4 Дата списания
            r["payment_type"],     #  5 Источник списания (Баланс/Счёт/Бонусы…)
            r["upd_sum"],          #  6 Сумма
            r["upd_num"],          #  7 Номер документа (updNum)
        ]
        assert len(row) == len(HEADERS), f"row={len(row)} != {len(HEADERS)}"
        result.append(row)
    return result

# ════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS
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
        n, r = divmod(n - 1, 26); buf.append(chr(65 + r))
    return "".join(reversed(buf))


def write_sheet(ss: gspread.Spreadsheet, rows: List[List[str]]) -> None:
    needed_rows = max(5000, len(rows) + 20)
    needed_cols = len(HEADERS) + 3

    try:
        ws = ss.worksheet(SHEET_TITLE)
        ws.resize(rows=needed_rows, cols=needed_cols)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_TITLE, rows=needed_rows, cols=needed_cols)
    ws.clear()

    n  = len(rows)
    lc = _col(len(HEADERS))

    ws.update(
        values=[[
            f"Реклама (Финансы)  |  Период: {DATE_FROM_RU} – {DATE_TO_RU}  |  "
            f"Транзакций: {n}  |  "
            f"Выгружено: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        ]],
        range_name="A1",
    )
    format_cell_range(ws, "A1", FMT_M)

    ws.update(values=[HEADERS], range_name=f"A5:{lc}5")
    format_cell_range(ws, f"A5:{lc}5", FMT_H)

    if n == 0:
        ws.update(
            values=[["Данных за период не найдено — смотри консоль"]],
            range_name="A6",
        )
        format_cell_range(ws, "A6", FMT_E)
        return

    for i in range(0, n, 500):
        chunk = rows[i: i + 500]
        r0, r1 = 6 + i, 5 + i + len(chunk)
        ws.update(values=chunk, range_name=f"A{r0}:{lc}{r1}")
        time.sleep(0.5)

    format_cell_range(ws, f"A6:{lc}{5 + n}", FMT_D)
    print(f"\n  ✓ «{SHEET_TITLE}»: {n} строк × {len(HEADERS)} колонок")

# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

MONTH_NAMES = {
    1:"январь",2:"февраль",3:"март",4:"апрель",
    5:"май",6:"июнь",7:"июль",8:"август",
    9:"сентябрь",10:"октябрь",11:"ноябрь",12:"декабрь"
}

def main() -> None:
    print("\n" + "═"*60)
    print("  WB Реклама Финансы  |  Автоматизация_реклама_фин  |  v2.0")
    print(f"  Период  : {MONTH_NAMES[MONTH]} {YEAR}  ({DATE_FROM_RU} — {DATE_TO_RU})")
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

    # Шаг 0: диагностика
    log_promotion_count()

    # Шаг 1: справочник кампаний (для ставки Единая/Ручная)
    camp_map = fetch_advert_list()

    # Шаг 2: один запрос — все транзакции аккаунта за период
    finance = fetch_finance(camp_map)

    # Сборка
    rows = build_rows(finance)

    # Итог
    print("\n" + "═"*60)
    print(f"  Транзакций к записи : {len(rows)}")
    if rows:
        try:
            total = sum(float(r[5]) for r in rows)
            print(f"  Итого сумма, ₽      : {total:,.2f}")
        except Exception:
            pass

    print(f"\n  Записываю в «{SHEET_TITLE}» ...")
    try:
        write_sheet(ss, rows)
    except Exception as e:
        print(f"  ✗ Ошибка записи: {e}")
        import traceback; traceback.print_exc()
        return

    print("═"*60 + "\n")


main()
