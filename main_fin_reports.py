import os, tempfile

import time, requests, gspread, pytz
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import CellFormat, TextFormat, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from calendar import monthrange
from typing import Any, Dict, List, Optional, Tuple

# ════════════════════════════════════════════════════════════════════
#  КОНФИГ
# ════════════════════════════════════════════════════════════════════

WB_TOKEN       = os.environ["WB_TOKEN"]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
# credentials.json восстанавливается из секрета GOOGLE_JSON
_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(os.environ["GOOGLE_JSON"])
_tmp.flush()
CREDENTIALS_PATH = _tmp.name
SHEET_TITLE      = "Автоматизация_фин_отчеты"

# Юридическое лицо — suppliercontract_code в API всегда None
LEGAL_ENTITY = "ИП Буторин Е. И."

URL_REPORTS = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod"

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

_last_day    = monthrange(YEAR, MONTH)[1]
DATE_FROM    = f"{YEAR:04d}-{MONTH:02d}-01"
DATE_TO      = f"{YEAR:04d}-{MONTH:02d}-{_last_day:02d}"
DATE_FROM_RU = f"01.{MONTH:02d}.{YEAR:04d}"
DATE_TO_RU   = f"{_last_day:02d}.{MONTH:02d}.{YEAR:04d}"

# ════════════════════════════════════════════════════════════════════
#  ЗАГОЛОВКИ — 21 колонка строго по ТЗ
# ════════════════════════════════════════════════════════════════════

HEADERS = [
    "No отчёта",                                                          #  1
    "Юридическое лицо",                                                   #  2
    "Дата начала",                                                        #  3
    "Дата конца",                                                         #  4
    "Дата формирования",                                                  #  5
    "Тип отчёта",                                                         #  6
    "Продажа",                                                            #  7
    "В том числе Компенсация скидки по программе лояльности",             #  8
    "К перечислению за товар",                                            #  9
    "Согласованная скидка, %",                                            # 10
    "Стоимость логистики",                                                # 11
    "Стоимость хранения",                                                 # 12
    "Стоимость платной приёмки",                                          # 13
    "Прочие удержания/выплаты",                                           # 14
    "Общая сумма штрафов",                                                # 15
    "Корректировка Вознаграждения Вайлдберриз (ВВ)",                      # 16
    "Стоимость участия в программе лояльности",                           # 17
    "Сумма удержанная за начисленные баллы программы лояльности",         # 18
    "Разовое изменение срока перечисления денежных средств",              # 19
    "Итого к оплате",                                                     # 20
    "Валюта",                                                             # 21
]
assert len(HEADERS) == 21

# ════════════════════════════════════════════════════════════════════
#  СПРАВОЧНИКИ
# ════════════════════════════════════════════════════════════════════

# report_type (число) → название по ТЗ
# 1 = Основной, 2 = По выкупам
REPORT_TYPE_MAP = {
    1: "Основной",
    2: "По выкупам",
}

# currency_name → отображаемое значение
CURRENCY_MAP = {
    "RUB": "руб.",
    "USD": "USD",
    "EUR": "EUR",
}

# ════════════════════════════════════════════════════════════════════
#  HTTP
# ════════════════════════════════════════════════════════════════════

def _wh() -> Dict[str, str]:
    return {"Authorization": WB_TOKEN, "Content-Type": "application/json"}

def safe_get(url: str, params: Optional[Dict] = None,
             timeout: int = 90) -> Tuple[Optional[requests.Response], str]:
    try:
        r = requests.get(url, headers=_wh(), params=params, timeout=timeout)
        return r, ""
    except Exception as exc:
        return None, str(exc)

# ════════════════════════════════════════════════════════════════════
#  УТИЛИТЫ
# ════════════════════════════════════════════════════════════════════

def _col(n: int) -> str:
    buf = []
    while n > 0:
        n, r = divmod(n - 1, 26); buf.append(chr(65 + r))
    return "".join(reversed(buf))

def _fmt_create_dt(raw: str) -> str:
    """
    Дата формирования → 2026-02-28T21:18:02+00:00
    API отдаёт '2026-02-02' (только дата без времени).
    Добавляем T00:00:00+00:00.
    """
    if not raw: return ""
    try:
        s = str(raw).strip()
        # Если уже содержит время — парсим и переформатируем
        if "T" in s:
            # Убираем Z, нормализуем
            s2 = s.replace("Z", "+00:00")
            dt = datetime.fromisoformat(s2)
            return dt.strftime("%Y-%m-%dT%H:%M:%S+00:00")
        else:
            # Только дата — добавляем время
            dt = datetime.fromisoformat(s[:10])
            return dt.strftime("%Y-%m-%dT%H:%M:%S+00:00")
    except Exception:
        return str(raw)

def _fmt_period_dt(raw: str) -> str:
    """
    Даты начала и конца → 2026-02-28T00:00:00+03:00
    API отдаёт '2026-01-26' (только дата).
    Добавляем T00:00:00+03:00.
    """
    if not raw: return ""
    try:
        s = str(raw).strip()
        if "T" in s:
            s2 = s.replace("Z", "+00:00")
            dt = datetime.fromisoformat(s2)
            return dt.strftime("%Y-%m-%dT%H:%M:%S+03:00")
        else:
            dt = datetime.fromisoformat(s[:10])
            return dt.strftime("%Y-%m-%dT%H:%M:%S+03:00")
    except Exception:
        return str(raw)

def _fmt_num(v: Any) -> str:
    """Число с 2 знаками после запятой, 0 при пустом."""
    if v is None or v == "": return "0"
    try:
        f = round(float(v), 2)
        return str(int(f)) if f == int(f) else str(f)
    except Exception:
        return str(v)

def _fmt_pct(v: Any) -> str:
    """Процент — целое число."""
    if v is None or v == "": return "0"
    try:
        return str(int(round(float(v))))
    except Exception:
        return str(v)

def _n(v: Any) -> float:
    if v is None or v == "": return 0.0
    try:    return float(v)
    except: return 0.0

# ════════════════════════════════════════════════════════════════════
#  ПОЛУЧЕНИЕ ОТЧЁТОВ
# ════════════════════════════════════════════════════════════════════

def fetch_reports() -> List[Dict]:
    print("\n── ФИНАНСОВЫЕ ОТЧЁТЫ (/api/v5/supplier/reportDetailByPeriod) ─")
    print(f"  Период: {DATE_FROM_RU} — {DATE_TO_RU}")

    all_items: List[Dict] = []
    rrdid     = 0
    page      = 0
    diag_done = False

    while True:
        page += 1
        params = {
            "dateFrom": DATE_FROM,
            "dateTo":   DATE_TO,
            "rrdid":    rrdid,
            "limit":    100000,
        }
        print(f"  страница {page} (rrdid={rrdid}) ...", end="", flush=True)

        r, err = safe_get(URL_REPORTS, params=params, timeout=90)
        if r is None:
            print(f"  ✗ {err}"); break

        print(f"  HTTP {r.status_code}", end="")

        if r.status_code == 429:
            print("  → 429 жду 65 с ...")
            time.sleep(65)
            r, err = safe_get(URL_REPORTS, params=params, timeout=90)
            if r is None:
                print(f" ✗ {err}"); break
            print(f"  retry HTTP {r.status_code}", end="")

        if r.status_code == 204:
            print("  (нет данных — 204)"); break
        if r.status_code != 200:
            print(f"\n  ✗ HTTP {r.status_code}: {r.text[:400]}"); break

        try:
            items = r.json()
        except Exception as e:
            print(f"  ✗ JSON: {e}"); break

        if not isinstance(items, list):
            print(f"  ✗ Тип ответа: {type(items)}: {str(items)[:200]}"); break

        print(f"  строк: {len(items)}")

        if not items:
            break

        # Диагностика report_type — один раз
        if not diag_done:
            diag_done = True
            rtypes = sorted(set(str(i.get("report_type", "")) for i in items))
            dtypes = sorted(set(str(i.get("doc_type_name", "")) for i in items))

        all_items.extend(items)

        max_rrdid = max(int(i.get("rrd_id") or 0) for i in items)
        if max_rrdid <= rrdid or len(items) < 100000:
            break
        rrdid = max_rrdid
        time.sleep(0.5)

    print(f"\n  ▶ Строк детализации получено: {len(all_items)}")

    # ── ДИАГНОСТИКА: какие rr_dt реально вернул API ──────────────
    if all_items:
        rr_dates = sorted(set(str(i.get("rr_dt") or "")[:10] for i in all_items if i.get("rr_dt")))
        print(f"  ▶ Уникальных rr_dt: {len(rr_dates)}")
        print(f"  ▶ Первая rr_dt   : {rr_dates[0] if rr_dates else chr(8212)}")
        print(f"  ▶ Последняя rr_dt: {rr_dates[-1] if rr_dates else chr(8212)}")
        print(f"  ▶ Все rr_dt: {rr_dates}")
        df_dates = sorted(set(str(i.get("date_from") or "")[:10] for i in all_items if i.get("date_from")))
        dt_dates = sorted(set(str(i.get("date_to")   or "")[:10] for i in all_items if i.get("date_to")))
        print(f"  ▶ date_from диапазон: {df_dates[0] if df_dates else chr(8212)} --- {df_dates[-1] if df_dates else chr(8212)}")
        print(f"  ▶ date_to   диапазон: {dt_dates[0] if dt_dates else chr(8212)} --- {dt_dates[-1] if dt_dates else chr(8212)}")
        print(f"  ▶ Ключи первой строки: {list(all_items[0].keys())}")
    # ─────────────────────────────────────────────────────────────
    return all_items

# ════════════════════════════════════════════════════════════════════
#  АГРЕГАЦИЯ
#
#  Группируем по realizationreport_id (БЕЗ фильтра по типу).
#  Тип отчёта берём из report_type: 1=Основной, 2=По выкупам.
#  Одна строка = один уникальный отчёт.
# ════════════════════════════════════════════════════════════════════

def _extract_oid_from_token(token: str) -> str:
    """Извлекает oid (owner id) из JWT токена — используется как префикс № отчёта."""
    try:
        import base64, json as _json
        payload = token.split('.')[1]
        payload += '=' * (4 - len(payload) % 4)
        data = _json.loads(base64.b64decode(payload))
        return str(data.get("oid", ""))
    except Exception:
        return ""


# Извлекаем oid один раз при загрузке модуля
OID = _extract_oid_from_token(WB_TOKEN)


def aggregate(items: List[Dict]) -> List[Dict]:
    """Суммируем детализацию по каждому дню и типу отчёта (rr_dt + report_type).
    Формат строки совпадает с ручной синей таблицей:
      - Один день + тип = одна строка
      - № отчёта = {oid}{yyyymmdd}  для Основного
                   {oid}{yyyymmdd}1 для По выкупам
      - date_from = date_to = сам rr_dt (дата дня)
      - Итого ~56 строк за месяц (28 дней × 2 типа, но не каждый день есть оба типа)
    """

    unique_rr_dt = sorted(set(str(i.get("rr_dt") or "")[:10] for i in items if i.get("rr_dt")))
    unique_rtypes = sorted(set(str(i.get("report_type") or "") for i in items))



    agg: Dict[str, Dict] = {}  # ключ = "rr_dt_date|report_type"

    for item in items:
        rr_dt_raw = str(item.get("rr_dt") or "").strip()
        day_key   = rr_dt_raw[:10]  # YYYY-MM-DD
        if not day_key or len(day_key) < 10:
            continue

        rt     = int(item.get("report_type") or 0)
        combo  = f"{day_key}|{rt}"

        if combo not in agg:
            rtype   = REPORT_TYPE_MAP.get(rt, str(rt) if rt else "")
            cur_raw = str(item.get("currency_name") or "RUB")
            cur     = CURRENCY_MAP.get(cur_raw.upper(), cur_raw)

            # № отчёта: {oid}{yyyymmdd} для Основного (rt=1), {oid}{yyyymmdd}1 для По выкупам (rt=2)
            date_compact = day_key.replace("-", "")  # 20260201
            if rt == 2:
                rep_id_val = f"{OID}{date_compact}1"
            else:
                rep_id_val = f"{OID}{date_compact}"

            # date_from = date_to = сама дата дня (rr_dt)
            day_dt_str = _fmt_period_dt(day_key)

            agg[combo] = {
                "rep_id":                rep_id_val,
                "legal":                 LEGAL_ENTITY,
                "date_from":             day_dt_str,
                "date_to":               day_dt_str,
                "create_dt":             _fmt_create_dt(str(item.get("create_dt") or "")),
                "doc_type":              rtype,
                "retail_amount":         0.0,
                "loyalty_discount":      0.0,
                "ppvz_for_pay":          0.0,
                "agreed_discount":       0.0,
                "delivery_rub":          0.0,
                "storage_fee":           0.0,
                "acceptance_fee":        0.0,
                "other_deductions":      0.0,
                "penalty":               0.0,
                "ppvz_reward":           0.0,
                "loyalty_cost":          0.0,
                "loyalty_bonus_sum":     0.0,
                "deduction_date_change": 0.0,
                "currency":              cur,
            }

        a = agg[combo]

        # Обновляем create_dt — берём последнее (максимальное)
        cd = str(item.get("create_dt") or "")
        if cd and (not a.get("_create_dt_raw") or cd > a.get("_create_dt_raw", "")):
            a["_create_dt_raw"] = cd
            a["create_dt"]      = _fmt_create_dt(cd)

        # Суммируем финансовые поля
        # Возврат вычитается из retail_amount и ppvz_for_pay (как в ручном отчёте)
        _is_return = str(item.get("doc_type_name") or "") == "Возврат"
        _sign = -1 if _is_return else 1
        a["retail_amount"]         += _sign * _n(item.get("retail_amount"))
        # Компенсация скидки по программе лояльности = cashback_discount
        a["loyalty_discount"]      += _sign * _n(item.get("cashback_discount"))
        a["ppvz_for_pay"]          += _sign * _n(item.get("ppvz_for_pay"))
        if _n(item.get("sale_percent")) != 0 and a["agreed_discount"] == 0:
            a["agreed_discount"]    = _n(item.get("sale_percent"))
        a["delivery_rub"]          += _n(item.get("delivery_rub"))
        a["storage_fee"]           += _n(item.get("storage_fee"))
        a["acceptance_fee"]        += _n(item.get("acceptance"))
        a["other_deductions"]      += (_n(item.get("deduction")) +
                                       _n(item.get("additional_payment")))
        a["penalty"]               += _n(item.get("penalty"))
        # ppvz_vw_nds и rebill_logistic_cost — в ручном отчёте всегда 0, не суммируем
        a["ppvz_reward"]           += 0  # _n(item.get("ppvz_vw_nds"))
        a["loyalty_cost"]          += _n(item.get("supplier_promo"))
        a["loyalty_bonus_sum"]     += _n(item.get("cashback_amount"))
        a["deduction_date_change"] += 0  # _n(item.get("rebill_logistic_cost"))

    # Итого к оплате + сортировка по дате, затем по типу
    result = []
    # Сортировка: сначала все Основной (rt=1) по дате, потом все По выкупам (rt=2) по дате
    # combo = "YYYY-MM-DD|1" или "YYYY-MM-DD|2"
    for combo, a in sorted(agg.items(), key=lambda x: (x[0].split("|")[1], x[0].split("|")[0])):
        a["itogo"] = (a["ppvz_for_pay"]
                      - a["delivery_rub"]
                      - a["storage_fee"]
                      - a["acceptance_fee"]
                      - a["other_deductions"]
                      - a["penalty"]
                      - a["ppvz_reward"]
                      - a["loyalty_cost"]
                      - a["loyalty_bonus_sum"]
                      - a["deduction_date_change"])
        result.append(a)

    print(f"  ▶ Строк (день × тип): {len(result)}")
    return result


def build_rows(reports: List[Dict]) -> List[List[str]]:
    result = []
    for r in reports:
        row = [
            r["rep_id"],                          #  1 No отчёта
            r["legal"],                            #  2 Юридическое лицо
            r["date_from"],                        #  3 Дата начала
            r["date_to"],                          #  4 Дата конца
            r["create_dt"],                        #  5 Дата формирования
            r["doc_type"],                         #  6 Тип отчёта
            _fmt_num(r["retail_amount"]),           #  7 Продажа
            _fmt_num(r["loyalty_discount"]),        #  8 Компенсация скидки лояльности
            _fmt_num(r["ppvz_for_pay"]),            #  9 К перечислению за товар
            _fmt_pct(r["agreed_discount"]),         # 10 Согласованная скидка, %
            _fmt_num(r["delivery_rub"]),            # 11 Стоимость логистики
            _fmt_num(r["storage_fee"]),             # 12 Стоимость хранения
            _fmt_num(r["acceptance_fee"]),          # 13 Стоимость платной приёмки
            _fmt_num(r["other_deductions"]),        # 14 Прочие удержания/выплаты
            _fmt_num(r["penalty"]),                 # 15 Общая сумма штрафов
            _fmt_num(r["ppvz_reward"]),             # 16 Корректировка Вознаграждения ВВ
            _fmt_num(r["loyalty_cost"]),            # 17 Стоимость участия в программе лояльности
            _fmt_num(r["loyalty_bonus_sum"]),       # 18 Сумма удержанная за баллы лояльности
            _fmt_num(r["deduction_date_change"]),   # 19 Разовое изменение срока перечисления
            _fmt_num(r["itogo"]),                   # 20 Итого к оплате
            r["currency"],                          # 21 Валюта
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


def write_sheet(ss: gspread.Spreadsheet, rows: List[List[str]]) -> None:
    needed_rows = max(1000, len(rows) + 20)
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
            f"Финансовые отчёты  |  Период: {DATE_FROM_RU} – {DATE_TO_RU}  |  "
            f"Дней: {n}  |  "
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
    print("  WB Финансовые отчёты  |  Автоматизация_фин_отчеты  |  v2.0")
    print(f"  Период    : {MONTH_NAMES[MONTH]} {YEAR}  ({DATE_FROM_RU} — {DATE_TO_RU})")
    print(f"  Лист      : «{SHEET_TITLE}»")
    print(f"  Колонок   : {len(HEADERS)}")
    print(f"  Юрлицо    : {LEGAL_ENTITY}")
    print("═"*60)

    print("\nПодключаюсь к Google Sheets ...")
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, SCOPES)
        ss    = gspread.Client(auth=creds).open_by_key(SPREADSHEET_ID)
        print("  ✓ OK")
    except Exception as e:
        print(f"  ✗ {e}"); return

    items = fetch_reports()

    if not items:
        print("\n  ✗ Данных нет — проверь токен и период")
        write_sheet(ss, [])
        return

    print("\n── АГРЕГАЦИЯ ────────────────────────────────────────────────")
    reports = aggregate(items)
    rows    = build_rows(reports)

    print("\n" + "═"*60)
    print(f"  Строк детализации  : {len(items)}")
    print(f"  Уникальных дней    : {len(rows)}")
    if rows:
        try:
            total = sum(float(r[19]) for r in rows)
            print(f"  Итого к оплате, ₽  : {total:,.2f}")
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
