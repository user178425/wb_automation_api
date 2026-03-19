# ╔══════════════════════════════════════════════════════════════════╗
# ║  WB Заказы (Месяц) → Автоматизация_заказы  |  v1.3             ║
# ║                                                                  ║
# ║  Данные берутся за полный месяц (по умолчанию — прошлый).       ║
# ║  ✓ Лист: "Автоматизация_заказы"                                 ║
# ║  ✓ 69 колонок строго по ТЗ                                      ║
# ║  ✓ Блок дат — легко менять вручную                              ║
# ║  ✓ v1.3: полный список nmId из /content/v2/get/cards/list       ║
# ║                                                                  ║
# ║  GITHUB VERSION:                                                 ║
# ║  - WB_TOKEN и SPREADSHEET_ID берутся из переменных окружения    ║
# ║  - credentials.json восстанавливается из env GOOGLE_JSON  ║
# ║  - !pip install убран — зависимости в requirements.txt          ║
# ╚══════════════════════════════════════════════════════════════════╝

import os, json, time, tempfile, requests, gspread, pandas as pd
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import CellFormat, TextFormat, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from typing import Any, Dict, List, Optional, Set
import pytz as _pytz
from datetime import timedelta as _td
from calendar import monthrange

# ════════════════════════════════════════════════════════════════════
#  КОНФИГ — из переменных окружения (GitHub Secrets)
# ════════════════════════════════════════════════════════════════════

WB_TOKEN       = os.environ["WB_TOKEN"]          # секрет WB_TOKEN
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]    # секрет SPREADSHEET_ID

# credentials.json восстанавливается из секрета GOOGLE_JSON
_creds_json = os.environ["GOOGLE_JSON"]
_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(_creds_json)
_tmp.flush()
CREDENTIALS_PATH = _tmp.name

SHEET_TITLE = "Автоматизация_заказы"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ════════════════════════════════════════════════════════════════════
#  БЛОК ДАТ
#  По умолчанию: прошлый полный месяц (автовычисление).
#  Чтобы задать вручную — раскомментируй РУЧНОЙ ВВОД:
# ─────────────────────────────────────────────────────────────────
# CUR_YEAR    = 2026
# CUR_MONTH   = 2
# PAST_YEAR   = 2026
# PAST_MONTH  = 1
# ════════════════════════════════════════════════════════════════════

_tz  = _pytz.timezone("Europe/Moscow")
_now = datetime.now(_tz)

if _now.month == 1:
    CUR_YEAR, CUR_MONTH = _now.year - 1, 12
else:
    CUR_YEAR, CUR_MONTH = _now.year, _now.month - 1

if CUR_MONTH == 1:
    PAST_YEAR, PAST_MONTH = CUR_YEAR - 1, 12
else:
    PAST_YEAR, PAST_MONTH = CUR_YEAR, CUR_MONTH - 1


def _month_range(year: int, month: int):
    last_day = monthrange(year, month)[1]
    fr    = f"{year:04d}-{month:02d}-01"
    to    = f"{year:04d}-{month:02d}-{last_day:02d}"
    fr_ru = f"01.{month:02d}.{year:04d}"
    to_ru = f"{last_day:02d}.{month:02d}.{year:04d}"
    return fr, to, fr_ru, to_ru


CUR_FROM_ISO,  CUR_TO_ISO,  CUR_FROM_RU,  CUR_TO_RU  = _month_range(CUR_YEAR,  CUR_MONTH)
PAST_FROM_ISO, PAST_TO_ISO, PAST_FROM_RU, PAST_TO_RU = _month_range(PAST_YEAR, PAST_MONTH)

CUR_FROM    = CUR_FROM_ISO
CUR_TO      = CUR_TO_ISO
PAST_FROM   = PAST_FROM_ISO
PAST_TO     = PAST_TO_ISO
CUR_FROM_DT = f"{CUR_FROM_ISO}T00:00:00"

# ════════════════════════════════════════════════════════════════════
#  ЗАГОЛОВКИ — 69 колонок строго по ТЗ
# ════════════════════════════════════════════════════════════════════

HEADERS = [
    "Артикул продавца",                                          #  1
    "Артикул WB",                                                #  2
    "Название",                                                  #  3
    "Предмет",                                                   #  4
    "Бренд",                                                     #  5
    "Удаленный товар",                                           #  6
    "Рейтинг карточки",                                          #  7
    "Рейтинг по отзывам",                                        #  8
    "Показы",                                                    #  9
    "Показы (предыдущий период)",                                # 10
    "CTR",                                                       # 11
    "CTR (предыдущий период)",                                   # 12
    "Доля карточки в выручке",                                   # 13
    "Доля карточки в выручке (предыдущий период)",               # 14
    "Переходы в карточку",                                       # 15
    "Переходы в карточку (предыдущий период)",                   # 16
    "Положили в корзину",                                        # 17
    "Положили в корзину (предыдущий период)",                    # 18
    "Добавили в отложенные",                                     # 19
    "Добавили в отложенные (предыдущий период)",                 # 20
    "Заказали, шт",                                              # 21
    "Заказали, шт (предыдущий период)",                          # 22
    "Заказали ВБ клуб, шт",                                      # 23
    "Заказали ВБ клуб, шт (предыдущий период)",                  # 24
    "Выкупили, шт",                                              # 25
    "Выкупы, шт (предыдущий период)",                            # 26
    "Выкупили ВБ клуб, шт",                                      # 27
    "Выкупы ВБ клуб, шт (предыдущий период)",                    # 28
    "Отменили, шт",                                              # 29
    "Отменили, шт (предыдущий период)",                          # 30
    "Отменили ВБ клуб, шт",                                      # 31
    "Отменили ВБ клуб, шт (предыдущий период)",                  # 32
    "Конверсия в корзину, %",                                    # 33
    "Конверсия в корзину, % (предыдущий период)",                # 34
    "Конверсия в заказ, %",                                      # 35
    "Конверсия в заказ, % (предыдущий период)",                  # 36
    "Процент выкупа",                                            # 37
    "Процент выкупа (предыдущий период)",                        # 38
    "Процент выкупа ВБ клуб",                                    # 39
    "Процент выкупа ВБ клуб (предыдущий период)",                # 40
    "Заказали на сумму, руб",                                    # 41
    "Заказали на сумму, руб (предыдущий период)",                # 42
    "Заказали на сумму ВБ клуб, ₽",                              # 43
    "Заказали на сумму ВБ клуб, ₽ (предыдущий период)",          # 44
    "Динамика суммы заказов, руб",                               # 45
    "Динамика суммы заказов ВБ клуб, ₽",                         # 46
    "Выкупили на сумму, руб",                                    # 47
    "Выкупили на сумму, руб (предыдущий период)",                # 48
    "Выкупили на сумму ВБ клуб, ₽",                              # 49
    "Выкупили на сумму ВБ клуб, ₽ (предыдущий период)",          # 50
    "Отменили на сумму, руб",                                    # 51
    "Отменили на сумму, руб (предыдущий период)",                # 52
    "Отменили на сумму ВБ клуб, ₽",                              # 53
    "Отменили на сумму ВБ клуб, ₽ (предыдущий период)",          # 54
    "Средняя цена, руб",                                         # 55
    "Средняя цена, руб (предыдущий период)",                     # 56
    "Средняя цена ВБ клуб, ₽",                                   # 57
    "Средняя цена ВБ клуб, ₽ (предыдущий период)",               # 58
    "Среднее количество заказов в день, шт",                     # 59
    "Среднее количество заказов в день, шт (предыдущий период)", # 60
    "Среднее количество заказов в день ВБ клуб, шт",             # 61
    "Среднее количество заказов в день ВБ клуб, шт (предыдущий период)", # 62
    "Остатки склад ВБ, шт",                                      # 63
    "Остатки МП, шт",                                            # 64
    "Сумма остатков на складах, руб",                            # 65
    "Среднее время доставки",                                    # 66
    "Среднее время доставки (предыдущий период)",                # 67
    "Локальные заказы, %",                                       # 68
    "Локальные заказы, % (предыдущий период)",                   # 69
]
assert len(HEADERS) == 69, f"Ожидается 69 колонок, сейчас {len(HEADERS)}"

# ════════════════════════════════════════════════════════════════════
#  ЖУРНАЛ
# ════════════════════════════════════════════════════════════════════

_log: List[str] = []

def _log_status(name: str, code: Any, detail: str = "") -> None:
    icon = "✓" if str(code) == "200" else "✗"
    msg  = f"  {icon} {name}: {code}" + (f"  ({detail})" if detail else "")
    print(msg)
    _log.append(msg)

# ════════════════════════════════════════════════════════════════════
#  HTTP
# ════════════════════════════════════════════════════════════════════

def _wh() -> Dict:
    return {"Authorization": WB_TOKEN, "Content-Type": "application/json"}

def safe_get(label: str, url: str, params=None, timeout=90) -> Optional[requests.Response]:
    try:
        r = requests.get(url, headers=_wh(), params=params, timeout=timeout)
        _log_status(label, r.status_code)
        return r
    except Exception as exc:
        _log_status(label, "ERROR", str(exc)[:80])
        return None

def safe_post(label: str, url: str, body: Any, timeout=90) -> Optional[requests.Response]:
    try:
        r = requests.post(url, headers=_wh(), json=body, timeout=timeout)
        _log_status(label, r.status_code)
        return r
    except Exception as exc:
        _log_status(label, "ERROR", str(exc)[:80])
        return None

# ════════════════════════════════════════════════════════════════════
#  УТИЛИТЫ
# ════════════════════════════════════════════════════════════════════

def _n(d: Any, *keys) -> float:
    cur = d
    for k in keys:
        if not isinstance(cur, dict): return 0.0
        cur = next((cur[a] for a in k if a in cur), None) \
              if isinstance(k, tuple) else cur.get(k)
        if cur is None: return 0.0
    try:    return float(cur)
    except: return 0.0

def _s(d: Any, *keys, default="") -> str:
    cur = d
    for k in keys:
        if not isinstance(cur, dict): return default
        cur = next((cur[a] for a in k if a in cur), None) \
              if isinstance(k, tuple) else cur.get(k)
        if cur is None: return default
    return str(cur) if cur is not None else default

def _i(v: Any) -> str:
    if v is None or v == "": return "0"
    try:    return str(int(round(float(v))))
    except: return "0"

def _col(n: int) -> str:
    buf = []
    while n > 0:
        n, r = divmod(n-1, 26); buf.append(chr(65+r))
    return "".join(reversed(buf))

def _r2(v: Any) -> str:
    if not v: return "0"
    try:
        f = round(float(v), 2)
        return str(int(f)) if f == int(f) else str(f)
    except: return "0"

def _pct(num: float, den: float) -> str:
    if den == 0: return "0"
    return str(int(round(num / den * 100)))

def _pct_buyout(buy: float, order: float, cancel: float) -> str:
    den = order - cancel
    if den <= 0: return "0"
    return str(int(round(buy / den * 100)))

def _fmt_delivery(val: Any) -> str:
    if not val or val == 0: return "0"
    if isinstance(val, dict):
        days  = val.get("days",  0) or 0
        hours = val.get("hours", 0) or 0
        mins  = val.get("mins",  0) or 0
        if days == 0 and hours == 0 and mins == 0: return "0"
        parts = []
        if days:  parts.append(f"{int(days)}д")
        if hours: parts.append(f"{int(hours)}ч")
        if mins:  parts.append(f"{int(mins)}м")
        return " ".join(parts) if parts else "0"
    try:
        secs = float(val)
        if secs <= 0: return "0"
        h = int(secs // 3600)
        d = h // 24; h = h % 24
        parts = []
        if d: parts.append(f"{d}д")
        if h: parts.append(f"{h}ч")
        return " ".join(parts) if parts else "0"
    except: return str(val)

# ════════════════════════════════════════════════════════════════════
#  УДАЛЁННЫЕ ТОВАРЫ  /content/v2/get/cards/trash
# ════════════════════════════════════════════════════════════════════

TRASH_URL = "https://content-api.wildberries.ru/content/v2/get/cards/trash"

def fetch_deleted_cards() -> Dict[int, Dict]:
    print("\n── УДАЛЁННЫЕ ТОВАРЫ (/content/v2/get/cards/trash) ──────────")
    result: Dict[int, Dict] = {}
    cursor: Dict = {"limit": 100}
    page = 0

    while True:
        page += 1
        body = {"settings": {"cursor": cursor, "filter": {"withPhoto": -1}}}
        print(f"  страница {page} ...", end="", flush=True)
        r = safe_post(f"Корзина стр.{page}", TRASH_URL, body, timeout=60)

        if r is None:
            print("\n  ✗ Нет ответа от API"); break
        if r.status_code == 204:
            print("\n  (корзина пуста — 204)"); break
        if r.status_code != 200:
            print(f"\n  ✗ HTTP {r.status_code}: {r.text[:400]}"); break

        try:
            raw = r.json()
        except Exception as e:
            print(f"\n  ✗ JSON: {e}"); break

        cards = raw.get("cards") or []
        cur_  = raw.get("cursor") or {}
        print(f" карточек: {len(cards)}", end="")

        for card in cards:
            if not isinstance(card, dict): continue
            nm = card.get("nmID") or card.get("nmId")
            if not nm: continue
            nm = int(nm)
            vendor   = card.get("vendorCode") or ""
            title    = card.get("title") or card.get("name") or ""
            category = card.get("subjectName") or card.get("object") or ""
            brand    = card.get("brand") or card.get("brandName") or ""
            if not title:
                for ch in (card.get("characteristics") or []):
                    if isinstance(ch, dict):
                        val = ch.get("Наименование") or ch.get("name") or ""
                        if val:
                            title = str(val); break
            result[nm] = {"vendor": vendor, "title": title,
                          "category": str(category), "brand": brand}

        limit = cursor.get("limit", 100)
        if len(cards) < limit:
            print("  ← последняя страница"); break

        next_updated = cur_.get("updatedAt")
        next_nm      = cur_.get("nmID")
        if not next_updated or not next_nm:
            print("  ← курсор исчерпан"); break

        cursor = {"limit": limit, "updatedAt": next_updated, "nmID": next_nm}
        print()
        time.sleep(0.3)

    print(f"\n  ▶ Удалённых товаров из API: {len(result)}")
    return result

# ════════════════════════════════════════════════════════════════════
#  ВСЕ АКТИВНЫЕ КАРТОЧКИ  /content/v2/get/cards/list
# ════════════════════════════════════════════════════════════════════

CARDS_LIST_URL = "https://content-api.wildberries.ru/content/v2/get/cards/list"

def fetch_all_cards() -> Dict[int, Dict]:
    print("\n── ВСЕ КАРТОЧКИ (/content/v2/get/cards/list) ───────────────")
    result: Dict[int, Dict] = {}
    cursor: Dict = {"limit": 100}
    page = 0

    while True:
        page += 1
        body = {"settings": {"cursor": cursor, "filter": {"withPhoto": -1}}}
        print(f"  страница {page} ...", end="", flush=True)
        r = safe_post(f"Карточки стр.{page}", CARDS_LIST_URL, body, timeout=60)

        if r is None:
            print("\n  ✗ Нет ответа от API"); break
        if r.status_code == 204:
            print("\n  (нет карточек — 204)"); break
        if r.status_code != 200:
            print(f"\n  ✗ HTTP {r.status_code}: {r.text[:400]}"); break

        try:
            raw = r.json()
        except Exception as e:
            print(f"\n  ✗ JSON: {e}"); break

        cards = raw.get("cards") or []
        cur_  = raw.get("cursor") or {}
        print(f" карточек: {len(cards)}", end="")

        for card in cards:
            if not isinstance(card, dict): continue
            nm = card.get("nmID") or card.get("nmId")
            if not nm: continue
            nm = int(nm)
            vendor   = card.get("vendorCode") or ""
            title    = card.get("title") or card.get("name") or ""
            category = card.get("subjectName") or card.get("object") or ""
            brand    = card.get("brand") or card.get("brandName") or ""
            if not title:
                for ch in (card.get("characteristics") or []):
                    if isinstance(ch, dict):
                        val = ch.get("Наименование") or ch.get("name") or ""
                        if val:
                            title = str(val); break
            result[nm] = {"vendor": vendor, "title": title,
                          "category": str(category), "brand": brand}

        limit = cursor.get("limit", 100)
        if len(cards) < limit:
            print("  ← последняя страница"); break

        next_updated = cur_.get("updatedAt")
        next_nm      = cur_.get("nmID")
        if not next_updated or not next_nm:
            print("  ← курсор исчерпан"); break

        cursor = {"limit": limit, "updatedAt": next_updated, "nmID": next_nm}
        print()
        time.sleep(0.3)

    print(f"\n  ▶ Активных карточек из API: {len(result)}")
    return result

# ════════════════════════════════════════════════════════════════════
#  ВОРОНКА v3
# ════════════════════════════════════════════════════════════════════

FUNNEL_URL = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

def fetch_funnel(nm_ids: List[int]) -> Dict[int, Dict]:
    print("\n── ВОРОНКА v3 ───────────────────────────────────────────────")
    print(f"  nmIds         : {len(nm_ids)} шт")
    print(f"  Период        : {CUR_FROM_ISO} — {CUR_TO_ISO}")
    print(f"  Прошлый период: {PAST_FROM_ISO} — {PAST_TO_ISO}")
    result: Dict[int, Dict] = {}

    period_block = {
        "selectedPeriod": {"start": CUR_FROM_ISO,  "end": CUR_TO_ISO},
        "pastPeriod":     {"start": PAST_FROM_ISO, "end": PAST_TO_ISO},
    }

    BATCH_SIZE = 100
    batches = [nm_ids[i: i + BATCH_SIZE] for i in range(0, len(nm_ids), BATCH_SIZE)]
    print(f"  Батчей        : {len(batches)} (по {BATCH_SIZE} nmId)")

    for batch_idx, batch_ids in enumerate(batches):
        print(f"\n  ── Батч {batch_idx+1}/{len(batches)} ({len(batch_ids)} nmId) ──")
        try:
            for page in range(1, 101):
                body = {**period_block, "nmIds": batch_ids, "page": page}
                print(f"    page={page} ...", end="", flush=True)
                r = safe_post(f"Воронка б{batch_idx+1} p{page}", FUNNEL_URL, body, timeout=60)

                if r is None or r.status_code in (401, 403):
                    print(f"\n  ✗ Нет доступа"); return result
                if r.status_code == 204:
                    print("\n    (нет данных — 204)"); break
                if r.status_code != 200:
                    print(f"\n    ✗ HTTP {r.status_code}: {r.text[:300]}"); break

                try:
                    raw = r.json()
                except Exception as e:
                    print(f"\n    ✗ JSON: {e}"); break

                if batch_idx == 0 and page == 1:
                    _diag_funnel(raw)

                items    = _pick_items(raw)
                has_more = _pick_has_more(raw)
                print(f" товаров: {len(items)} | hasMore: {has_more}")

                if not items:
                    break

                for item in items:
                    try:
                        nm_id, row = _parse_funnel_item(item)
                        if nm_id:
                            result[nm_id] = row
                    except Exception as e:
                        print(f"      ⚠ parse item: {e}")

                if not has_more:
                    break
                time.sleep(0.4)

        except Exception as e:
            print(f"  ✗ Воронка батч {batch_idx+1}: {e}")

        if batch_idx < len(batches) - 1:
            time.sleep(0.5)

    print(f"\n  ▶ Воронка итого: {len(result)} товаров")
    return result


def _calc_views(open_count: float, cart_count: float, ctr_pct: float) -> str:
    if ctr_pct > 0:
        views = cart_count / (ctr_pct / 100.0)
        return _i(views)
    elif open_count > 0:
        return _i(open_count)
    return "0"


def _parse_funnel_item(item: Dict) -> tuple:
    prod  = item.get("product") or item
    nm_id = int(prod.get("nmId", 0) or prod.get("nmID", 0) or 0)
    if not nm_id:
        return 0, {}

    stat = item.get("statistic") or item.get("statistics") or {}
    sel  = stat.get("selected")  or stat.get("selectedPeriod") or {}
    pst  = stat.get("past")      or stat.get("previousPeriod") or {}

    conv_s = sel.get("conversions") or {}
    conv_p = pst.get("conversions") or {}

    opens_s = _n(sel, "openCount")
    opens_p = _n(pst, "openCount")
    cart_s  = _n(sel, "cartCount")
    cart_p  = _n(pst, "cartCount")
    wish_s  = _n(sel, ("addToWishlist", "wishlistCount"))
    wish_p  = _n(pst, ("addToWishlist", "wishlistCount"))
    ord_c_s = _n(sel, "orderCount")
    ord_c_p = _n(pst, "orderCount")
    buy_c_s = _n(sel, "buyoutCount")
    buy_c_p = _n(pst, "buyoutCount")
    can_c_s = _n(sel, "cancelCount")
    can_c_p = _n(pst, "cancelCount")
    ord_s_s = _n(sel, "orderSum")
    ord_s_p = _n(pst, "orderSum")
    buy_s_s = _n(sel, "buyoutSum")
    buy_s_p = _n(pst, "buyoutSum")
    can_s_s = _n(sel, "cancelSum")
    can_s_p = _n(pst, "cancelSum")
    avg_p_s = _n(sel, "avgPrice")
    avg_p_p = _n(pst, "avgPrice")
    avg_o_s = _n(sel, "avgOrdersCountPerDay")
    avg_o_p = _n(pst, "avgOrdersCountPerDay")

    ctr_s = _i(_n(conv_s, "addToCartPercent"))
    ctr_p = _i(_n(conv_p, "addToCartPercent"))

    views_s = _calc_views(opens_s, cart_s, _n(conv_s, "addToCartPercent"))
    views_p = _calc_views(opens_p, cart_p, _n(conv_p, "addToCartPercent"))

    conv_cart_s = ctr_s
    conv_cart_p = ctr_p
    conv_ord_s  = _i(_n(conv_s, "cartToOrderPercent"))
    conv_ord_p  = _i(_n(conv_p, "cartToOrderPercent"))
    pct_buy_s   = _i(_n(conv_s, "buyoutPercent"))
    pct_buy_p   = _i(_n(conv_p, "buyoutPercent"))

    share_s  = _r2(_n(sel, "shareOrderPercent"))
    share_p  = _r2(_n(pst, "shareOrderPercent"))
    dynamics = _i(ord_s_s - ord_s_p)

    del_s   = _fmt_delivery(sel.get("timeToReady") or sel.get("deliveryTime") or 0)
    del_p   = _fmt_delivery(pst.get("timeToReady") or pst.get("deliveryTime") or 0)
    local_s = _r2(_n(sel, ("localizationPercent", "localPercent")))
    local_p = _r2(_n(pst, ("localizationPercent", "localPercent")))

    stk = prod.get("stocks") or {}
    stk_wb  = int(stk.get("wb",  0) or 0)
    stk_mp  = int(stk.get("mp",  0) or 0)
    stk_sum = int(stk.get("balanceSum", 0) or 0)

    return nm_id, {
        "vendor":       _s(prod, "vendorCode"),
        "nm_id":        nm_id,
        "title":        _s(prod, ("title", "name")),
        "category":     _s(prod, ("subjectName",)),
        "brand":        _s(prod, "brandName"),
        "rating_card":  _r2(_n(prod, ("productRating", "rating"))),
        "rating_fb":    _r2(_n(prod, ("feedbackRating", "reviewRating"))),
        "views_s":      views_s,
        "views_p":      views_p,
        "share_s":      share_s,
        "share_p":      share_p,
        "opens_s":      _i(opens_s),
        "opens_p":      _i(opens_p),
        "cart_s":       _i(cart_s),
        "cart_p":       _i(cart_p),
        "wish_s":       _i(wish_s),
        "wish_p":       _i(wish_p),
        "ord_c_s":      _i(ord_c_s),
        "ord_c_p":      _i(ord_c_p),
        "ord_c_wb_s":   "0",
        "ord_c_wb_p":   "0",
        "buy_c_s":      _i(buy_c_s),
        "buy_c_p":      _i(buy_c_p),
        "buy_c_wb_s":   "0",
        "buy_c_wb_p":   "0",
        "can_c_s":      _i(can_c_s),
        "can_c_p":      _i(can_c_p),
        "can_c_wb_s":   "0",
        "can_c_wb_p":   "0",
        "conv_cart_s":  conv_cart_s,
        "conv_cart_p":  conv_cart_p,
        "conv_ord_s":   conv_ord_s,
        "conv_ord_p":   conv_ord_p,
        "pct_buy_s":    pct_buy_s,
        "pct_buy_p":    pct_buy_p,
        "pct_buy_wb_s": "0",
        "pct_buy_wb_p": "0",
        "ord_s_s":      _i(ord_s_s),
        "ord_s_p":      _i(ord_s_p),
        "ord_s_wb_s":   "0",
        "ord_s_wb_p":   "0",
        "dynamics":     dynamics,
        "dyn_wb":       "0",
        "buy_s_s":      _i(buy_s_s),
        "buy_s_p":      _i(buy_s_p),
        "buy_s_wb_s":   "0",
        "buy_s_wb_p":   "0",
        "can_s_s":      _i(can_s_s),
        "can_s_p":      _i(can_s_p),
        "can_s_wb_s":   "0",
        "can_s_wb_p":   "0",
        "avg_p_s":      _i(avg_p_s),
        "avg_p_p":      _i(avg_p_p),
        "avg_p_wb_s":   "0",
        "avg_p_wb_p":   "0",
        "avg_o_s":      _r2(avg_o_s),
        "avg_o_p":      _r2(avg_o_p),
        "avg_o_wb_s":   "0",
        "avg_o_wb_p":   "0",
        "tot_buy_c":    buy_c_s,
        "stk_wb":       stk_wb,
        "stk_mp":       stk_mp,
        "stk_sum":      stk_sum,
        "del_s":        del_s,
        "del_p":        del_p,
        "local_s":      local_s,
        "local_p":      local_p,
    }


def _diag_funnel(raw: Any) -> None:
    items = _pick_items(raw)
    if not items: return
    prod = items[0].get("product") or items[0]
    stat = items[0].get("statistic") or items[0].get("statistics") or {}
    sel  = stat.get("selected") or stat.get("selectedPeriod") or {}
    print(f"\n  [ДИА] product keys : {list(prod.keys())}")
    print(f"  [ДИА] stat keys    : {list(stat.keys())}")
    print(f"  [ДИА] selected keys: {list(sel.keys())}")


def _pick_items(raw: Any) -> List[Dict]:
    if isinstance(raw, list):
        return [x for x in raw if isinstance(x, dict)]
    if not isinstance(raw, dict): return []
    data = raw.get("data")
    if isinstance(data, dict):
        for k in ("products","items","cards","rows"):
            v = data.get(k)
            if isinstance(v, list): return [x for x in v if isinstance(x, dict)]
    if isinstance(data, list): return [x for x in data if isinstance(x, dict)]
    for k in ("products","items","cards","rows"):
        v = raw.get(k)
        if isinstance(v, list): return [x for x in v if isinstance(x, dict)]
    return []


def _pick_has_more(raw: Any) -> bool:
    if not isinstance(raw, dict): return False
    containers = []
    if isinstance(raw.get("data"), dict): containers.append(raw["data"])
    containers.append(raw)
    for c in containers:
        if not isinstance(c, dict): continue
        for k in ("hasMore","isNextPage"):
            v = c.get(k)
            if v is not None: return bool(v)
    return False

# ════════════════════════════════════════════════════════════════════
#  ОСТАТКИ
# ════════════════════════════════════════════════════════════════════

STOCKS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/stocks"

def fetch_stocks() -> tuple:
    date_dt = f"{CUR_TO_ISO}T23:59:59"
    print("\n── ОСТАТКИ ─────────────────────────────────────────────────")
    print(f"  dateFrom: {date_dt}")
    try:
        r = safe_get("Остатки", STOCKS_URL, params={"dateFrom": date_dt}, timeout=90)
        if r is None or r.status_code != 200:
            print("  ⚠ Остатки недоступны → 0 везде")
            return {}, []

        raw = r.json()
        if not isinstance(raw, list):
            return {}, []

        out: Dict[int, Dict] = {}
        for item in raw:
            try:
                nm = item.get("nmId")
                if not nm: continue
                nm  = int(nm)
                wb  = int(item.get("quantity", 0) or 0)
                prc = float(item.get("Price", 0) or 0)
                val = prc * wb
                if nm not in out:
                    out[nm] = {"wb": 0, "mp": 0, "sum_rub": 0.0}
                out[nm]["wb"]      += wb
                out[nm]["sum_rub"] += val
            except Exception:
                pass

        nm_ids = sorted(out.keys())
        print(f"  ▶ Остатки: {len(out)} nmId")
        return out, nm_ids

    except Exception as e:
        print(f"  ✗ Остатки: {e}")
        return {}, []

# ════════════════════════════════════════════════════════════════════
#  ЗАКАЗЫ (СВЕРКА)
# ════════════════════════════════════════════════════════════════════

ORDERS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/orders"

def fetch_orders_fact() -> Dict[int, int]:
    print("\n── ЗАКАЗЫ (сверка) ─────────────────────────────────────────")
    try:
        r = safe_get("Заказы", ORDERS_URL,
                     params={"dateFrom": CUR_FROM_DT, "flag": 1}, timeout=90)
        if r is None or r.status_code == 204:
            return {}
        if r.status_code != 200:
            return {}

        raw = r.json()
        if not isinstance(raw, list): return {}

        out: Dict[int, int] = {}
        for item in raw:
            try:
                dt = (item.get("date") or item.get("lastChangeDate") or "")[:10]
                if dt and not (CUR_FROM <= dt <= CUR_TO): continue
                if item.get("isCancel"): continue
                nm = item.get("nmId")
                if not nm: continue
                nm = int(nm)
                out[nm] = out.get(nm, 0) + 1
            except Exception:
                pass

        print(f"  ▶ Заказы (факт): {len(out)} nmId, {sum(out.values())} шт")
        return out
    except Exception as e:
        print(f"  ✗ Заказы: {e}")
        return {}

# ════════════════════════════════════════════════════════════════════
#  ПРОДАЖИ (СВЕРКА)
# ════════════════════════════════════════════════════════════════════

SALES_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"

def fetch_sales_fact() -> Dict[int, int]:
    print("\n── ПРОДАЖИ (сверка) ────────────────────────────────────────")
    try:
        r = safe_get("Продажи", SALES_URL,
                     params={"dateFrom": CUR_FROM_DT, "flag": 1}, timeout=90)
        if r is None or r.status_code == 204:
            return {}
        if r.status_code != 200:
            return {}

        raw = r.json()
        if not isinstance(raw, list): return {}

        out: Dict[int, int] = {}
        for item in raw:
            try:
                dt = (item.get("date") or item.get("lastChangeDate") or "")[:10]
                if dt and not (CUR_FROM <= dt <= CUR_TO): continue
                sale_id = str(item.get("saleID") or "")
                if sale_id.startswith(("R", "D")): continue
                nm = item.get("nmId")
                if not nm: continue
                nm = int(nm)
                out[nm] = out.get(nm, 0) + 1
            except Exception:
                pass

        print(f"  ▶ Продажи (факт): {len(out)} nmId, {sum(out.values())} шт")
        return out
    except Exception as e:
        print(f"  ✗ Продажи: {e}")
        return {}

# ════════════════════════════════════════════════════════════════════
#  СБОРКА ТАБЛИЦЫ
# ════════════════════════════════════════════════════════════════════

def build_df(funnel: Dict[int, Dict],
             stocks: Dict[int, Dict],
             orders_f: Dict[int, int],
             sales_f: Dict[int, int],
             all_cards: Dict[int, Dict],
             deleted_meta: Dict[int, Dict],
             deleted_nm_ids: Set[int]) -> pd.DataFrame:
    rows = []

    all_ids = sorted(
        set(all_cards.keys())
        | set(funnel.keys())
        | deleted_nm_ids
        | set(orders_f.keys())
        | set(sales_f.keys())
        | set(stocks.keys())
    )
    print(f"  ▶ build_df: всего nmId = {len(all_ids)}")

    for nm_id in all_ids:
        try:
            f      = funnel.get(nm_id, {})
            stk    = stocks.get(nm_id, {})
            stk_wb = stk.get("wb", 0)
            stk_mp = stk.get("mp", 0)
            stk_ru = stk.get("sum_rub", 0.0)

            is_deleted_api = nm_id in deleted_nm_ids
            deleted = "Да" if is_deleted_api else "Нет"

            c_meta   = all_cards.get(nm_id, {})
            d_meta   = deleted_meta.get(nm_id, {})
            vendor   = f.get("vendor")   or c_meta.get("vendor")   or d_meta.get("vendor",   "")
            title    = f.get("title")    or c_meta.get("title")    or d_meta.get("title",    "")
            category = f.get("category") or c_meta.get("category") or d_meta.get("category", "")
            brand    = f.get("brand")    or c_meta.get("brand")    or d_meta.get("brand",    "")

            views_s = f.get("views_s", "0")
            views_p = f.get("views_p", "0")
            ctr_s   = f.get("conv_cart_s", "0")
            ctr_p   = f.get("conv_cart_p", "0")
            del_s   = f.get("del_s", "0")
            del_p   = f.get("del_p", "0")

            row = [
                vendor,                           #  1
                str(nm_id),                       #  2
                title,                            #  3
                category,                         #  4
                brand,                            #  5
                deleted,                          #  6
                f.get("rating_card",  "0"),       #  7
                f.get("rating_fb",    "0"),       #  8
                views_s,                          #  9
                views_p,                          # 10
                ctr_s,                            # 11
                ctr_p,                            # 12
                f.get("share_s",      "0"),       # 13
                f.get("share_p",      "0"),       # 14
                f.get("opens_s",      "0"),       # 15
                f.get("opens_p",      "0"),       # 16
                f.get("cart_s",       "0"),       # 17
                f.get("cart_p",       "0"),       # 18
                f.get("wish_s",       "0"),       # 19
                f.get("wish_p",       "0"),       # 20
                f.get("ord_c_s",      "0"),       # 21
                f.get("ord_c_p",      "0"),       # 22
                f.get("ord_c_wb_s",   "0"),       # 23
                f.get("ord_c_wb_p",   "0"),       # 24
                f.get("buy_c_s",      "0"),       # 25
                f.get("buy_c_p",      "0"),       # 26
                f.get("buy_c_wb_s",   "0"),       # 27
                f.get("buy_c_wb_p",   "0"),       # 28
                f.get("can_c_s",      "0"),       # 29
                f.get("can_c_p",      "0"),       # 30
                f.get("can_c_wb_s",   "0"),       # 31
                f.get("can_c_wb_p",   "0"),       # 32
                f.get("conv_cart_s",  "0"),       # 33
                f.get("conv_cart_p",  "0"),       # 34
                f.get("conv_ord_s",   "0"),       # 35
                f.get("conv_ord_p",   "0"),       # 36
                f.get("pct_buy_s",    "0"),       # 37
                f.get("pct_buy_p",    "0"),       # 38
                f.get("pct_buy_wb_s", "0"),       # 39
                f.get("pct_buy_wb_p", "0"),       # 40
                f.get("ord_s_s",      "0"),       # 41
                f.get("ord_s_p",      "0"),       # 42
                f.get("ord_s_wb_s",   "0"),       # 43
                f.get("ord_s_wb_p",   "0"),       # 44
                f.get("dynamics",     "0"),       # 45
                f.get("dyn_wb",       "0"),       # 46
                f.get("buy_s_s",      "0"),       # 47
                f.get("buy_s_p",      "0"),       # 48
                f.get("buy_s_wb_s",   "0"),       # 49
                f.get("buy_s_wb_p",   "0"),       # 50
                f.get("can_s_s",      "0"),       # 51
                f.get("can_s_p",      "0"),       # 52
                f.get("can_s_wb_s",   "0"),       # 53
                f.get("can_s_wb_p",   "0"),       # 54
                f.get("avg_p_s",      "0"),       # 55
                f.get("avg_p_p",      "0"),       # 56
                f.get("avg_p_wb_s",   "0"),       # 57
                f.get("avg_p_wb_p",   "0"),       # 58
                f.get("avg_o_s",      "0"),       # 59
                f.get("avg_o_p",      "0"),       # 60
                f.get("avg_o_wb_s",   "0"),       # 61
                f.get("avg_o_wb_p",   "0"),       # 62
                _i(stk_wb),                       # 63
                _i(stk_mp),                       # 64
                _i(stk_ru),                       # 65
                del_s,                            # 66
                del_p,                            # 67
                f.get("local_s",      "0"),       # 68
                f.get("local_p",      "0"),       # 69
            ]
            assert len(row) == len(HEADERS), f"row={len(row)} headers={len(HEADERS)}"
            rows.append(row)
        except Exception as e:
            print(f"  ⚠ build row nmId={nm_id}: {e}")

    df = pd.DataFrame(rows, columns=HEADERS).fillna("0")
    df = df.replace({"": "0", None: "0"})

    df["_sort"] = df["Удаленный товар"].apply(lambda x: 1 if x == "Да" else 0)
    df = df.sort_values("_sort", kind="stable").drop(columns=["_sort"]).reset_index(drop=True)
    return df

# ════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS
# ════════════════════════════════════════════════════════════════════

_BLK  = {"red": 0.0,  "green": 0.0,  "blue": 0.0}
_WHT  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
_NAVY = {"red": 0.07, "green": 0.21, "blue": 0.38}
_RED  = {"red": 0.8,  "green": 0.0,  "blue": 0.0}
_YEL  = {"red": 1.0,  "green": 0.96, "blue": 0.76}

FMT_H   = CellFormat(backgroundColor=_NAVY,
                     textFormat=TextFormat(bold=True,  fontSize=10, foregroundColor=_WHT))
FMT_D   = CellFormat(backgroundColor=_WHT,
                     textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_M   = CellFormat(backgroundColor=_WHT,
                     textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_E   = CellFormat(backgroundColor=_WHT,
                     textFormat=TextFormat(bold=False, fontSize=11, foregroundColor=_RED))
FMT_W   = CellFormat(backgroundColor=_YEL,
                     textFormat=TextFormat(bold=True,  fontSize=10, foregroundColor=_RED))


def write_sheet(ss, df: pd.DataFrame, status_summary: str,
                deleted_nm_ids: Set[int]) -> None:
    try:
        ws = ss.worksheet(SHEET_TITLE)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_TITLE, rows=5000, cols=len(HEADERS)+5)
    ws.clear()

    lc = _col(len(HEADERS))
    n  = len(df)
    del_count = int((df["Удаленный товар"] == "Да").sum()) if n > 0 else 0

    run_day = datetime.now().day
    if run_day <= 21:
        warn_prev = (f"⚠ Данные предыдущего периода ({PAST_FROM_RU}–{PAST_TO_RU}) могут отличаться "
                     f"от ЛК: WB финализирует данные в течение 2–4 недель.")
    else:
        warn_prev = (f"Данные предыдущего периода ({PAST_FROM_RU}–{PAST_TO_RU}) финализированы.")

    ws.update(
        values=[[f"Период: {CUR_FROM_RU} – {CUR_TO_RU}  ({CUR_FROM_ISO} – {CUR_TO_ISO})  |  "
                 f"Предыдущий: {PAST_FROM_RU} – {PAST_TO_RU}  |  "
                 f"Выгружено: {datetime.now().strftime('%d.%m.%Y %H:%M')}"]],
        range_name="A1",
    )
    ws.update(
        values=[[f"Статусы: {status_summary}  |  Удалённых в таблице: {del_count}"]],
        range_name="A2",
    )
    ws.update(values=[[f"Строк: {n}  |  Колонок: {len(HEADERS)}"]], range_name="A3")
    ws.update(values=[[warn_prev]], range_name="A4")

    format_cell_range(ws, "A1", FMT_M)
    format_cell_range(ws, "A2", FMT_W if "нет данных" in status_summary else FMT_M)
    format_cell_range(ws, "A3", FMT_M)
    format_cell_range(ws, "A4", FMT_W if run_day <= 21 else FMT_M)

    ws.update(values=[HEADERS], range_name=f"A5:{lc}5")
    format_cell_range(ws, f"A5:{lc}5", FMT_H)

    if n == 0:
        ws.update(values=[["Нет данных — смотри лог"]], range_name="A6")
        format_cell_range(ws, "A6", FMT_E)
        return

    rows_str = [
        [("0" if (v is None or str(v) in ("nan", "")) else str(v)) for v in row]
        for row in df.values.tolist()
    ]
    for i in range(0, n, 500):
        chunk = rows_str[i: i+500]
        r0, r1 = 6+i, 5+i+len(chunk)
        ws.update(values=chunk, range_name=f"A{r0}:{lc}{r1}")
        time.sleep(0.5)

    format_cell_range(ws, f"A6:{lc}{5+n}", FMT_D)
    print(f"\n  ✓ «{SHEET_TITLE}»: {n} строк × {len(HEADERS)} колонок  "
          f"(удалённых: {del_count})")

# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

def main():
    month_names = {
        1:"январь",2:"февраль",3:"март",4:"апрель",
        5:"май",6:"июнь",7:"июль",8:"август",
        9:"сентябрь",10:"октябрь",11:"ноябрь",12:"декабрь"
    }
    print("\n" + "═"*60)
    print("  WB Заказы (Месяц)  |  Автоматизация_заказы  |  v1.3-gh")
    print(f"  Текущий  : {month_names[CUR_MONTH]} {CUR_YEAR}  ({CUR_FROM_RU} — {CUR_TO_RU})")
    print(f"  Прошлый  : {month_names[PAST_MONTH]} {PAST_YEAR}  ({PAST_FROM_RU} — {PAST_TO_RU})")
    print(f"  Лист     : «{SHEET_TITLE}»")
    print(f"  Колонок  : {len(HEADERS)}")
    print("═"*60)

    print("\nПодключаюсь к Google Sheets ...")
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, SCOPES)
        ss    = gspread.Client(auth=creds).open_by_key(SPREADSHEET_ID)
        print("  ✓ OK")
    except Exception as e:
        print(f"  ✗ {e}"); raise

    deleted_meta   = fetch_deleted_cards()
    deleted_nm_ids: Set[int] = set(deleted_meta.keys())

    all_cards = fetch_all_cards()

    stocks, nm_ids_stocks = fetch_stocks()
    orders_f = fetch_orders_fact()
    sales_f  = fetch_sales_fact()

    all_nm_ids = sorted(
        set(all_cards.keys())
        | set(nm_ids_stocks)
        | set(orders_f.keys())
        | set(sales_f.keys())
        | deleted_nm_ids
    )
    print(f"\n  ▶ nmId всего: {len(all_nm_ids)}")

    funnel = fetch_funnel(all_nm_ids)

    print("\n" + "═"*60)
    s_cards   = f"Карточки: {len(all_cards)} шт"
    s_funnel  = f"Воронка: {'OK' if funnel   else 'нет данных'}"
    s_stocks  = f"Остатки: {'OK' if stocks   else 'нет данных'}"
    s_orders  = f"Заказы:  {'OK' if orders_f else 'нет данных'}"
    s_sales   = f"Продажи: {'OK' if sales_f  else 'нет данных'}"
    s_deleted = f"Удалённые: {len(deleted_nm_ids)} шт"
    for s in (s_cards, s_funnel, s_stocks, s_orders, s_sales, s_deleted):
        print(f"  {'✓' if 'нет данных' not in s else '~'} {s}")
    status_summary = " | ".join([s_cards, s_funnel, s_stocks, s_orders, s_sales, s_deleted])

    print("\n  Собираю таблицу ...")
    df = build_df(funnel, stocks, orders_f, sales_f, all_cards,
                  deleted_meta, deleted_nm_ids)
    del_in_df = int((df["Удаленный товар"] == "Да").sum()) if len(df) > 0 else 0
    print(f"  Строк: {len(df)}  |  Из них удалённых: {del_in_df}")

    print("\n  Записываю в Google Sheets ...")
    write_sheet(ss, df, status_summary, deleted_nm_ids)

    print("\n" + "═"*60)
    print(f"  ✓ Готово!  Лист: «{SHEET_TITLE}»  Строк: {len(df)}")
    print("═"*60 + "\n")


main()
