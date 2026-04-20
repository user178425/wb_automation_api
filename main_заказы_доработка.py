# ╔══════════════════════════════════════════════════════════════════╗
# ║  WB Заказы → Автоматизация_ЗаказыWB  |  v3.0                  ║
# ║                                                                  ║
# ║  ИЗМЕНЕНИЯ v3.0 (рефакторинг показов):                         ║
# ║  ✓ «Увидели карточку» → «Переходы в карточку» (openCount)      ║
# ║  ✓ Убраны дублирующие колонки переходов (были 9/10 и 15/16)    ║
# ║  ✓ Убраны колонки «Конверсия в переход» (11/12) — нет показов  ║
# ║  ✓ Конверсии берутся из conversions API, не считаются вручную  ║
# ║  ✓ wbClub поля берутся из воронки v3 (не 0)                    ║
# ║  ✓ timeToReady берётся из воронки v3 (nm-report убран)         ║
# ║  ✓ 65 колонок вместо 69 (убраны 4 дублирующих/пустых)         ║
# ╚══════════════════════════════════════════════════════════════════╝

import os, tempfile, time, requests, gspread, pandas as pd
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import CellFormat, TextFormat, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from typing import Any, Dict, List, Optional

# ════════════════════════════════════════════════════════════════════
#  КОНФИГ — из переменных окружения (GitHub Secrets)
# ════════════════════════════════════════════════════════════════════

WB_TOKEN       = os.environ["WB_TOKEN"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID_2"]

_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(os.environ["GOOGLE_JSON"])
_tmp.flush()
CREDENTIALS_PATH = _tmp.name

SHEET_TITLE = "Автоматизация_ЗаказыWB"

# ─── Даты периодов ────────────────────────────────────────────────
import pytz as _pytz
from datetime import timedelta as _td

def _last_monday_msk() -> "datetime":
    tz  = _pytz.timezone("Europe/Moscow")
    now = datetime.now(tz)
    if now.weekday() == 0 and now.hour < 6:
        now -= _td(days=1)
    cur_mon = now - _td(days=now.weekday())
    return (cur_mon - _td(weeks=1)).replace(hour=0, minute=0, second=0, microsecond=0)

_MON = _last_monday_msk()

CUR_FROM_ISO  = _MON.strftime("%Y-%m-%d")
CUR_TO_ISO    = (_MON + _td(days=6)).strftime("%Y-%m-%d")
PAST_FROM_ISO = (_MON - _td(weeks=1)).strftime("%Y-%m-%d")
PAST_TO_ISO   = (_MON - _td(days=1)).strftime("%Y-%m-%d")

CUR_FROM_RU   = _MON.strftime("%d.%m.%Y")
CUR_TO_RU     = (_MON + _td(days=6)).strftime("%d.%m.%Y")
PAST_FROM_RU  = (_MON - _td(weeks=1)).strftime("%d.%m.%Y")
PAST_TO_RU    = (_MON - _td(days=1)).strftime("%d.%m.%Y")

CUR_FROM  = CUR_FROM_ISO
CUR_TO    = CUR_TO_ISO
PAST_FROM = PAST_FROM_ISO
PAST_TO   = PAST_TO_ISO

CUR_FROM_DT = f"{CUR_FROM_ISO}T00:00:00"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ════════════════════════════════════════════════════════════════════
#  ЗАГОЛОВКИ ТАБЛИЦЫ (65 колонок)
#
#  Изменения vs v2.x:
#  - Убраны «Увидели карточку» (9,10) — показов в API нет
#  - Убраны «Конверсия в переход» (11,12) — нет смысла без показов
#  - «Переходы в карточку» теперь на позиции 9/10 (openCount)
#  - Конверсии — из поля conversions воронки (готовые от API)
#  - wbClub — из воронки v3 (реальные значения, не 0)
#  - timeToReady — из воронки v3 (nm-report убран полностью)
# ════════════════════════════════════════════════════════════════════

HEADERS = [
    "Артикул продавца",                                        #  1
    "Номенклатура",                                            #  2
    "Название",                                                #  3
    "Категория",                                               #  4
    "Бренд",                                                   #  5
    "Удаленный товар",                                         #  6
    "Рейтинг карточки",                                        #  7
    "Рейтинг по отзывам",                                      #  8
    "Переходы в карточку",                                     #  9
    "Переходы в карточку (предыдущий период)",                 # 10
    "Доля карточки в выручке",                                 # 11
    "Доля карточки в выручке (предыдущий период)",             # 12
    "Положили в корзину",                                      # 13
    "Положили в корзину (предыдущий период)",                  # 14
    "Добавили в отложенные",                                   # 15
    "Добавили в отложенные (предыдущий период)",               # 16
    "Заказали, шт",                                            # 17
    "Заказали, шт (предыдущий период)",                        # 18
    "Заказали ВБ клуб, шт",                                    # 19
    "Заказали ВБ клуб, шт (предыдущий период)",                # 20
    "Выкупили, шт",                                            # 21
    "Выкупы, шт (предыдущий период)",                          # 22
    "Выкупили ВБ клуб, шт",                                    # 23
    "Выкупы ВБ клуб, шт (предыдущий период)",                  # 24
    "Отменили, шт",                                            # 25
    "Отменили, шт (предыдущий период)",                        # 26
    "Отменили ВБ клуб, шт",                                    # 27
    "Отменили ВБ клуб, шт (предыдущий период)",                # 28
    "Конверсия в корзину, %",                                  # 29
    "Конверсия в корзину, % (предыдущий период)",              # 30
    "Конверсия в заказ, %",                                    # 31
    "Конверсия в заказ, % (предыдущий период)",                # 32
    "Процент выкупа",                                          # 33
    "Процент выкупа (предыдущий период)",                      # 34
    "Процент выкупа ВБ клуб",                                  # 35
    "Процент выкупа ВБ клуб (предыдущий период)",              # 36
    "Заказали на сумму, руб",                                  # 37
    "Заказали на сумму, руб (предыдущий период)",              # 38
    "Заказали на сумму ВБ клуб, ₽",                            # 39
    "Заказали на сумму ВБ клуб, ₽ (предыдущий период)",        # 40
    "Динамика суммы заказов, руб",                             # 41
    "Динамика суммы заказов ВБ клуб, ₽",                       # 42
    "Выкупили на сумму, руб",                                  # 43
    "Выкупили на сумму, руб (предыдущий период)",              # 44
    "Выкупили на сумму ВБ клуб, ₽",                            # 45
    "Выкупили на сумму ВБ клуб, ₽ (предыдущий период)",        # 46
    "Отменили на сумму, руб",                                  # 47
    "Отменили на сумму, руб (предыдущий период)",              # 48
    "Отменили на сумму ВБ клуб, ₽",                            # 49
    "Отменили на сумму ВБ клуб, ₽ (предыдущий период)",        # 50
    "Средняя цена, р",                                         # 51
    "Средняя цена, руб (предыдущий период)",                   # 52
    "Средняя цена ВБ клуб, ₽",                                 # 53
    "Средняя цена ВБ клуб, ₽ (предыдущий период)",             # 54
    "Среднее количество заказов в день, шт",                   # 55
    "Среднее количество заказов в день, шт (предыдущий период)",# 56
    "Среднее количество заказов в день ВБ клуб, шт",           # 57
    "Среднее количество заказов в день ВБ клуб, шт (предыдущий период)", # 58
    "Остатки склад ВБ, шт",                                    # 59
    "Остатки МП, шт",                                          # 60
    "Сумма остатков на складах, руб",                          # 61
    "Среднее время доставки",                                  # 62
    "Среднее время доставки (предыдущий период)",              # 63
    "Локальные заказы, %",                                     # 64
    "Локальные заказы, % (предыдущий период)",                 # 65
]

assert len(HEADERS) == 65, f"Ожидается 65 колонок, сейчас {len(HEADERS)}"

# ════════════════════════════════════════════════════════════════════
#  ЖУРНАЛ СТАТУСОВ
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

def safe_get(label: str, url: str, params=None,
             timeout=90) -> Optional[requests.Response]:
    try:
        r = requests.get(url, headers=_wh(), params=params, timeout=timeout)
        _log_status(label, r.status_code)
        return r
    except Exception as exc:
        _log_status(label, "ERROR", str(exc)[:80])
        return None

def safe_post(label: str, url: str, body: Any,
              timeout=90) -> Optional[requests.Response]:
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

def _fmt_delivery(val: Any) -> str:
    """
    Среднее время доставки.
    dict {days, hours, mins} → «Xд Yч Zм»
    число (секунды)          → конвертируем
    """
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
#  МЕТОД 1: ВОРОНКА v3
#  POST /api/analytics/v3/sales-funnel/products
#
#  Что берём из selected/past:
#    openCount            → переходы в карточку (клики)
#    cartCount            → добавили в корзину
#    orderCount/Sum       → заказы
#    buyoutCount/Sum      → выкупы
#    cancelCount/Sum      → отмены
#    avgPrice             → средняя цена
#    avgOrdersCountPerDay → среднее заказов в день
#    shareOrderPercent    → доля в выручке
#    addToWishlist        → отложенные
#    timeToReady          → время доставки {days, hours, mins}
#    localizationPercent  → локальные заказы %
#    wbClub               → полный блок WB Клуб
#    conversions          → готовые конверсии от WB
# ════════════════════════════════════════════════════════════════════

FUNNEL_URL = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

def fetch_funnel(nm_ids: List[int]) -> Dict[int, Dict]:
    print("\n── МЕТОД 1: Воронка v3 ──────────────────────────────────")
    print(f"  nmIds в запросе : {len(nm_ids)} штук")
    print(f"  Период          : {CUR_FROM_ISO} — {CUR_TO_ISO}")
    print(f"  Прошлый период  : {PAST_FROM_ISO} — {PAST_TO_ISO}")
    if nm_ids:
        print(f"  Первые 10 nmId  : {nm_ids[:10]}")

    result: Dict[int, Dict] = {}

    try:
        for page in range(1, 51):
            body = {
                "selectedPeriod": {"start": CUR_FROM_ISO,  "end": CUR_TO_ISO},
                "pastPeriod":     {"start": PAST_FROM_ISO, "end": PAST_TO_ISO},
                "nmIds": nm_ids,
                "page":  page,
            }
            print(f"  page={page} ...", end="", flush=True)
            r = safe_post(f"Воронка p{page}", FUNNEL_URL, body, timeout=60)

            if r is None or r.status_code in (401, 403):
                print(f"\n  ✗ Нет доступа (HTTP {r.status_code if r else '—'})")
                return result

            if r.status_code == 204:
                print("\n  (нет данных — 204)"); break

            if r.status_code != 200:
                print(f"\n  ✗ HTTP {r.status_code}: {r.text[:300]}"); break

            try:
                raw = r.json()
            except Exception as e:
                print(f"\n  ✗ JSON: {e}"); break

            if page == 1:
                _diag_funnel(raw)

            items    = _pick_items(raw)
            has_more = _pick_has_more(raw)
            print(f" товаров: {len(items)} | hasMore: {has_more}")

            if not items:
                print("  ⚠ Пустой список. Полный ответ сервера:")
                import json as _json
                try:    print(_json.dumps(raw, ensure_ascii=False, indent=2)[:2000])
                except: print(repr(raw)[:2000])
                break

            for item in items:
                try:
                    nm_id, row = _parse_funnel_item(item)
                    if nm_id:
                        result[nm_id] = row
                except Exception as e:
                    print(f"    ⚠ parse item: {e}")

            if not has_more:
                break
            time.sleep(0.4)

    except Exception as e:
        print(f"  ✗ Воронка упала с ошибкой: {e}")

    print(f"\n  ▶ Воронка итого: {len(result)} товаров")
    return result


def _parse_funnel_item(item: Dict) -> tuple:
    prod  = item.get("product") or item
    nm_id = int(prod.get("nmId", 0) or prod.get("nmID", 0) or 0)
    if not nm_id:
        return 0, {}

    stat = item.get("statistic") or item.get("statistics") or {}
    sel  = stat.get("selected")  or stat.get("selectedPeriod") or {}
    pst  = stat.get("past")      or stat.get("previousPeriod") or {}

    # ── wbClub — полный блок из воронки v3 ─────────────────────────
    sel_wb = sel.get("wbClub") or {}
    pst_wb = pst.get("wbClub") or {}

    # ── Конверсии — готовые от WB, не считаем вручную ──────────────
    # addToCartPercent   = переходы → корзина (%)
    # cartToOrderPercent = корзина → заказ (%)
    # buyoutPercent      = % выкупа
    sel_conv = sel.get("conversions") or {}
    pst_conv = pst.get("conversions") or {}

    conv_cart_s  = _i(_n(sel_conv, "addToCartPercent"))
    conv_cart_p  = _i(_n(pst_conv, "addToCartPercent"))
    conv_ord_s   = _i(_n(sel_conv, "cartToOrderPercent"))
    conv_ord_p   = _i(_n(pst_conv, "cartToOrderPercent"))
    pct_buy_s    = _i(_n(sel_conv, "buyoutPercent"))
    pct_buy_p    = _i(_n(pst_conv, "buyoutPercent"))
    pct_buy_wb_s = _i(_n(sel_wb,   "buyoutPercent"))
    pct_buy_wb_p = _i(_n(pst_wb,   "buyoutPercent"))

    # ── Переходы в карточку ─────────────────────────────────────────
    opens_s = _n(sel, "openCount")
    opens_p = _n(pst, "openCount")

    # ── Корзина / отложенные ────────────────────────────────────────
    cart_s = _n(sel, "cartCount")
    cart_p = _n(pst, "cartCount")
    wish_s = _n(sel, ("addToWishlist", "wishlistCount"))
    wish_p = _n(pst, ("addToWishlist", "wishlistCount"))

    # ── Заказы шт ───────────────────────────────────────────────────
    ord_c_s    = _n(sel,    "orderCount")
    ord_c_p    = _n(pst,    "orderCount")
    ord_c_wb_s = _n(sel_wb, "orderCount")
    ord_c_wb_p = _n(pst_wb, "orderCount")

    # ── Выкупы шт ───────────────────────────────────────────────────
    buy_c_s    = _n(sel,    "buyoutCount")
    buy_c_p    = _n(pst,    "buyoutCount")
    buy_c_wb_s = _n(sel_wb, "buyoutCount")
    buy_c_wb_p = _n(pst_wb, "buyoutCount")

    # ── Отмены шт ───────────────────────────────────────────────────
    can_c_s    = _n(sel,    "cancelCount")
    can_c_p    = _n(pst,    "cancelCount")
    can_c_wb_s = _n(sel_wb, "cancelCount")
    can_c_wb_p = _n(pst_wb, "cancelCount")

    # ── Суммы заказов ───────────────────────────────────────────────
    ord_s_s    = _n(sel,    "orderSum")
    ord_s_p    = _n(pst,    "orderSum")
    ord_s_wb_s = _n(sel_wb, "orderSum")
    ord_s_wb_p = _n(pst_wb, "orderSum")

    # ── Суммы выкупов ───────────────────────────────────────────────
    buy_s_s    = _n(sel,    "buyoutSum")
    buy_s_p    = _n(pst,    "buyoutSum")
    buy_s_wb_s = _n(sel_wb, "buyoutSum")
    buy_s_wb_p = _n(pst_wb, "buyoutSum")

    # ── Суммы отмен ─────────────────────────────────────────────────
    can_s_s    = _n(sel,    "cancelSum")
    can_s_p    = _n(pst,    "cancelSum")
    can_s_wb_s = _n(sel_wb, "cancelSum")
    can_s_wb_p = _n(pst_wb, "cancelSum")

    # ── Средние ─────────────────────────────────────────────────────
    avg_p_s    = _n(sel,    "avgPrice")
    avg_p_p    = _n(pst,    "avgPrice")
    avg_p_wb_s = _n(sel_wb, "avgPrice")
    avg_p_wb_p = _n(pst_wb, "avgPrice")

    avg_o_s    = _n(sel,    "avgOrdersCountPerDay")
    avg_o_p    = _n(pst,    "avgOrdersCountPerDay")
    avg_o_wb_s = _n(sel_wb, ("avgOrderCountPerDay", "avgOrdersCountPerDay"))
    avg_o_wb_p = _n(pst_wb, ("avgOrderCountPerDay", "avgOrdersCountPerDay"))

    # ── Доля в выручке ──────────────────────────────────────────────
    share_s = _r2(_n(sel, "shareOrderPercent"))
    share_p = _r2(_n(pst, "shareOrderPercent"))

    # ── Динамика ────────────────────────────────────────────────────
    dynamics = _i(ord_s_s    - ord_s_p)
    dyn_wb   = _i(ord_s_wb_s - ord_s_wb_p)

    # ── Время доставки — из воронки v3 ──────────────────────────────
    del_s = _fmt_delivery(sel.get("timeToReady") or 0)
    del_p = _fmt_delivery(pst.get("timeToReady") or 0)

    # ── Локальные заказы ────────────────────────────────────────────
    local_s = _r2(_n(sel, ("localizationPercent", "localPercent")))
    local_p = _r2(_n(pst, ("localizationPercent", "localPercent")))

    return nm_id, {
        "vendor":        _s(prod, "vendorCode"),
        "nm_id":         nm_id,
        "title":         _s(prod, ("title", "name")),
        "category":      _s(prod, ("subjectName",)),
        "brand":         _s(prod, "brandName"),
        "rating_card":   _r2(_n(prod, ("productRating", "rating"))),
        "rating_fb":     _r2(_n(prod, ("feedbackRating", "reviewRating"))),
        "opens_s":       _i(opens_s),
        "opens_p":       _i(opens_p),
        "share_s":       share_s,
        "share_p":       share_p,
        "cart_s":        _i(cart_s),
        "cart_p":        _i(cart_p),
        "wish_s":        _i(wish_s),
        "wish_p":        _i(wish_p),
        "ord_c_s":       _i(ord_c_s),
        "ord_c_p":       _i(ord_c_p),
        "ord_c_wb_s":    _i(ord_c_wb_s),
        "ord_c_wb_p":    _i(ord_c_wb_p),
        "buy_c_s":       _i(buy_c_s),
        "buy_c_p":       _i(buy_c_p),
        "buy_c_wb_s":    _i(buy_c_wb_s),
        "buy_c_wb_p":    _i(buy_c_wb_p),
        "can_c_s":       _i(can_c_s),
        "can_c_p":       _i(can_c_p),
        "can_c_wb_s":    _i(can_c_wb_s),
        "can_c_wb_p":    _i(can_c_wb_p),
        "conv_cart_s":   conv_cart_s,
        "conv_cart_p":   conv_cart_p,
        "conv_ord_s":    conv_ord_s,
        "conv_ord_p":    conv_ord_p,
        "pct_buy_s":     pct_buy_s,
        "pct_buy_p":     pct_buy_p,
        "pct_buy_wb_s":  pct_buy_wb_s,
        "pct_buy_wb_p":  pct_buy_wb_p,
        "ord_s_s":       _i(ord_s_s),
        "ord_s_p":       _i(ord_s_p),
        "ord_s_wb_s":    _i(ord_s_wb_s),
        "ord_s_wb_p":    _i(ord_s_wb_p),
        "dynamics":      dynamics,
        "dyn_wb":        dyn_wb,
        "buy_s_s":       _i(buy_s_s),
        "buy_s_p":       _i(buy_s_p),
        "buy_s_wb_s":    _i(buy_s_wb_s),
        "buy_s_wb_p":    _i(buy_s_wb_p),
        "can_s_s":       _i(can_s_s),
        "can_s_p":       _i(can_s_p),
        "can_s_wb_s":    _i(can_s_wb_s),
        "can_s_wb_p":    _i(can_s_wb_p),
        "avg_p_s":       _i(avg_p_s),
        "avg_p_p":       _i(avg_p_p),
        "avg_p_wb_s":    _i(avg_p_wb_s),
        "avg_p_wb_p":    _i(avg_p_wb_p),
        "avg_o_s":       _r2(avg_o_s),
        "avg_o_p":       _r2(avg_o_p),
        "avg_o_wb_s":    _r2(avg_o_wb_s),
        "avg_o_wb_p":    _r2(avg_o_wb_p),
        "tot_buy_c":     buy_c_s,
        "del_s":         del_s,
        "del_p":         del_p,
        "local_s":       local_s,
        "local_p":       local_p,
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
    diag_keys = ["openCount", "cartCount", "orderCount", "orderSum",
                 "buyoutCount", "buyoutSum", "cancelCount", "cancelSum",
                 "avgPrice", "avgOrdersCountPerDay", "shareOrderPercent",
                 "addToWishlist", "timeToReady", "localizationPercent",
                 "wbClub", "conversions"]
    print("  [ДИА] selected values:")
    for k in diag_keys:
        v = sel.get(k, "—")
        if v != "—":
            print(f"    {k}: {v}")


def _pick_items(raw: Any) -> List[Dict]:
    if isinstance(raw, list):
        return [x for x in raw if isinstance(x, dict)]
    if not isinstance(raw, dict): return []
    data = raw.get("data")
    if isinstance(data, dict):
        for k in ("products", "items", "cards", "rows"):
            v = data.get(k)
            if isinstance(v, list): return [x for x in v if isinstance(x, dict)]
    if isinstance(data, list): return [x for x in data if isinstance(x, dict)]
    for k in ("products", "items", "cards", "rows"):
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
        for k in ("hasMore", "isNextPage"):
            v = c.get(k)
            if v is not None: return bool(v)
    return False

# ════════════════════════════════════════════════════════════════════
#  ОСТАТКИ
#  GET /api/v1/supplier/stocks
# ════════════════════════════════════════════════════════════════════

STOCKS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/stocks"

def fetch_stocks() -> tuple:
    date_dt = f"{CUR_TO_ISO}T23:59:59"
    print("\n── ШАГ 1: Остатки (сбор nmId) ──────────────────────────")
    print(f"  dateFrom: {date_dt}")
    try:
        r = safe_get("Остатки", STOCKS_URL,
                     params={"dateFrom": date_dt}, timeout=90)
        if r is None or r.status_code != 200:
            print("  ⚠ Остатки недоступны → 0 везде")
            return {}, []

        raw = r.json()
        if not isinstance(raw, list):
            print(f"  ⚠ Тип: {type(raw).__name__}")
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
        print(f"  ▶ Остатки итого: {len(out)} уникальных nmId")
        print(f"  ▶ Первые 20 nmId: {nm_ids[:20]}")
        return out, nm_ids

    except Exception as e:
        print(f"  ✗ Остатки упали с ошибкой: {e}")
        return {}, []

# ════════════════════════════════════════════════════════════════════
#  ЗАКАЗЫ (СВЕРКА)
#  GET /api/v1/supplier/orders
# ════════════════════════════════════════════════════════════════════

ORDERS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/orders"

def fetch_orders_fact() -> Dict[int, int]:
    print("\n── МЕТОД 3: Заказы (сверка) ─────────────────────────────")
    try:
        r = safe_get("Заказы", ORDERS_URL,
                     params={"dateFrom": CUR_FROM_DT, "flag": 1}, timeout=90)
        if r is None or r.status_code == 204:
            print("  (нет данных)"); return {}
        if r.status_code != 200:
            print("  ⚠ Заказы (факт) недоступны → везде 0"); return {}

        raw = r.json()
        if not isinstance(raw, list):
            print(f"  ⚠ Тип: {type(raw)}"); return {}

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
        print(f"  ✗ Заказы упали с ошибкой: {e}")
        return {}

# ════════════════════════════════════════════════════════════════════
#  ПРОДАЖИ (СВЕРКА)
#  GET /api/v1/supplier/sales
# ════════════════════════════════════════════════════════════════════

SALES_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"

def fetch_sales_fact() -> Dict[int, int]:
    print("\n── МЕТОД 4: Продажи (сверка) ────────────────────────────")
    try:
        r = safe_get("Продажи", SALES_URL,
                     params={"dateFrom": CUR_FROM_DT, "flag": 1}, timeout=90)
        if r is None or r.status_code == 204:
            print("  (нет данных)"); return {}
        if r.status_code != 200:
            print("  ⚠ Продажи (факт) недоступны → везде 0"); return {}

        raw = r.json()
        if not isinstance(raw, list):
            print(f"  ⚠ Тип: {type(raw)}"); return {}

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
        print(f"  ✗ Продажи упали с ошибкой: {e}")
        return {}

# ════════════════════════════════════════════════════════════════════
#  СБОРКА ТАБЛИЦЫ
# ════════════════════════════════════════════════════════════════════

def build_df(funnel, stocks, orders_f, sales_f) -> pd.DataFrame:
    rows = []
    for nm_id, f in funnel.items():
        try:
            stk    = stocks.get(nm_id, {})
            stk_wb = stk.get("wb", 0)
            stk_mp = stk.get("mp", 0)
            stk_ru = stk.get("sum_rub", 0)

            deleted = "Да" if (stk_wb == 0 and stk_mp == 0
                                and f.get("tot_buy_c", 0) == 0) else "Нет"

            row = [
                f.get("vendor",       ""),       #  1
                str(nm_id),                       #  2
                f.get("title",        ""),        #  3
                f.get("category",     ""),        #  4
                f.get("brand",        ""),        #  5
                deleted,                          #  6
                f.get("rating_card",  "0"),       #  7
                f.get("rating_fb",    "0"),       #  8
                f.get("opens_s",      "0"),       #  9 Переходы в карточку
                f.get("opens_p",      "0"),       # 10 Переходы в карточку (пред.)
                f.get("share_s",      "0"),       # 11
                f.get("share_p",      "0"),       # 12
                f.get("cart_s",       "0"),       # 13
                f.get("cart_p",       "0"),       # 14
                f.get("wish_s",       "0"),       # 15
                f.get("wish_p",       "0"),       # 16
                f.get("ord_c_s",      "0"),       # 17
                f.get("ord_c_p",      "0"),       # 18
                f.get("ord_c_wb_s",   "0"),       # 19
                f.get("ord_c_wb_p",   "0"),       # 20
                f.get("buy_c_s",      "0"),       # 21
                f.get("buy_c_p",      "0"),       # 22
                f.get("buy_c_wb_s",   "0"),       # 23
                f.get("buy_c_wb_p",   "0"),       # 24
                f.get("can_c_s",      "0"),       # 25
                f.get("can_c_p",      "0"),       # 26
                f.get("can_c_wb_s",   "0"),       # 27
                f.get("can_c_wb_p",   "0"),       # 28
                f.get("conv_cart_s",  "0"),       # 29
                f.get("conv_cart_p",  "0"),       # 30
                f.get("conv_ord_s",   "0"),       # 31
                f.get("conv_ord_p",   "0"),       # 32
                f.get("pct_buy_s",    "0"),       # 33
                f.get("pct_buy_p",    "0"),       # 34
                f.get("pct_buy_wb_s", "0"),       # 35
                f.get("pct_buy_wb_p", "0"),       # 36
                f.get("ord_s_s",      "0"),       # 37
                f.get("ord_s_p",      "0"),       # 38
                f.get("ord_s_wb_s",   "0"),       # 39
                f.get("ord_s_wb_p",   "0"),       # 40
                f.get("dynamics",     "0"),       # 41
                f.get("dyn_wb",       "0"),       # 42
                f.get("buy_s_s",      "0"),       # 43
                f.get("buy_s_p",      "0"),       # 44
                f.get("buy_s_wb_s",   "0"),       # 45
                f.get("buy_s_wb_p",   "0"),       # 46
                f.get("can_s_s",      "0"),       # 47
                f.get("can_s_p",      "0"),       # 48
                f.get("can_s_wb_s",   "0"),       # 49
                f.get("can_s_wb_p",   "0"),       # 50
                f.get("avg_p_s",      "0"),       # 51
                f.get("avg_p_p",      "0"),       # 52
                f.get("avg_p_wb_s",   "0"),       # 53
                f.get("avg_p_wb_p",   "0"),       # 54
                f.get("avg_o_s",      "0"),       # 55
                f.get("avg_o_p",      "0"),       # 56
                f.get("avg_o_wb_s",   "0"),       # 57
                f.get("avg_o_wb_p",   "0"),       # 58
                _i(stk_wb),                       # 59
                _i(stk_mp),                       # 60
                _i(stk_ru),                       # 61
                f.get("del_s",        "0"),       # 62
                f.get("del_p",        "0"),       # 63
                f.get("local_s",      "0"),       # 64
                f.get("local_p",      "0"),       # 65
            ]

            assert len(row) == len(HEADERS), f"row={len(row)} headers={len(HEADERS)}"
            rows.append(row)

        except Exception as e:
            print(f"  ⚠ build row nmId={nm_id}: {e}")

    df = pd.DataFrame(rows, columns=HEADERS).fillna("0")
    df = df.replace({"": "0", None: "0"})
    return df

# ════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS
# ════════════════════════════════════════════════════════════════════

_BLK  = {"red": 0.0, "green": 0.0, "blue": 0.0}
_WHT  = {"red": 1.0, "green": 1.0, "blue": 1.0}
_NAVY = {"red": 0.07, "green": 0.21, "blue": 0.38}
_RED  = {"red": 0.8,  "green": 0.0,  "blue": 0.0}
_YEL  = {"red": 1.0,  "green": 0.96, "blue": 0.76}

FMT_H = CellFormat(backgroundColor=_NAVY,
                   textFormat=TextFormat(bold=True,  fontSize=10, foregroundColor=_WHT))
FMT_D = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_M = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=10, foregroundColor=_BLK))
FMT_E = CellFormat(backgroundColor=_WHT,
                   textFormat=TextFormat(bold=False, fontSize=11, foregroundColor=_RED))
FMT_W = CellFormat(backgroundColor=_YEL,
                   textFormat=TextFormat(bold=True,  fontSize=10, foregroundColor=_RED))


def write_sheet(ss, df: pd.DataFrame, status_summary: str) -> None:
    try:
        ws = ss.worksheet(SHEET_TITLE)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_TITLE, rows=5000, cols=len(HEADERS)+5)

    ws.clear()
    lc = _col(len(HEADERS))
    n  = len(df)

    ws.update(
        values=[[f"Период: {CUR_FROM_RU} – {CUR_TO_RU}  "
                 f"({CUR_FROM_ISO} – {CUR_TO_ISO})  |  "
                 f"Предыдущий: {PAST_FROM_RU} – {PAST_TO_RU}  |  "
                 f"Выгружено: {datetime.now().strftime('%d.%m.%Y %H:%M')}"]],
        range_name="A1",
    )
    ws.update(
        values=[[f"Статусы методов: {status_summary}"]],
        range_name="A2",
    )
    ws.update(
        values=[[f"Строк: {n}  |  Колонок: {len(HEADERS)}  |  Лист: {SHEET_TITLE}"]],
        range_name="A3",
    )

    format_cell_range(ws, "A1", FMT_M)
    format_cell_range(ws, "A2", FMT_W if "✗" in status_summary else FMT_M)
    format_cell_range(ws, "A3", FMT_M)

    ws.update(values=[HEADERS], range_name=f"A5:{lc}5")
    format_cell_range(ws, f"A5:{lc}5", FMT_H)

    if n == 0:
        ws.update(values=[["Воронка не вернула данных — смотри консоль"]], range_name="A6")
        format_cell_range(ws, "A6", FMT_E)
        return

    rows = [
        [("0" if (v is None or str(v) in ("nan", "")) else str(v)) for v in row]
        for row in df.values.tolist()
    ]

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

def main():
    print("\n" + "═"*60)
    print("  WB Заказы  |  Автоматизация_ЗаказыWB  |  v3.0")
    print(f"  Текущий  : {CUR_FROM_RU} — {CUR_TO_RU}  ({CUR_FROM_ISO} — {CUR_TO_ISO})")
    print(f"  Прошлый  : {PAST_FROM_RU} — {PAST_TO_RU}")
    print(f"  Лист     : «{SHEET_TITLE}»")
    print(f"  Колонок  : {len(HEADERS)}")
    print("═"*60)

    print("\nПодключаюсь к Google Sheets ...")
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, SCOPES)
        ss    = gspread.Client(auth=creds).open_by_key(SPREADSHEET_ID)
        print("  ✓ OK")
    except Exception as e:
        print(f"  ✗ {e}"); return

    stocks, nm_ids_stocks = fetch_stocks()
    orders_f = fetch_orders_fact()
    sales_f  = fetch_sales_fact()

    all_nm_ids = sorted(set(nm_ids_stocks)
                        | set(orders_f.keys())
                        | set(sales_f.keys()))
    print(f"\n  ▶ nmId объединено из остатков+заказов+продаж: {len(all_nm_ids)}")
    print(f"  ▶ Только в заказах/продажах (удалённые): "
          f"{sorted(set(all_nm_ids) - set(nm_ids_stocks))}")

    # nm-report убран — все данные из воронки v3
    funnel = fetch_funnel(all_nm_ids)

    print("\n" + "═"*60)
    print("  ИТОГ СТАТУСОВ:")
    s_funnel = f"Воронка: {'OK' if funnel   else 'нет данных'}"
    s_stocks = f"Остатки: {'OK' if stocks   else 'нет данных'}"
    s_orders = f"Заказы:  {'OK' if orders_f else 'нет данных'}"
    s_sales  = f"Продажи: {'OK' if sales_f  else 'нет данных'}"
    for s in (s_funnel, s_stocks, s_orders, s_sales):
        print(f"  {'✓' if 'OK' in s else '~'} {s}")
    status_summary = " | ".join([s_funnel, s_stocks, s_orders, s_sales])

    if not funnel:
        print("\n  ✗ Воронка пуста — таблица будет содержать заглушку.")
        print("  Проверь токен и период.")

    print("\n  Собираю таблицу ...")
    df = build_df(funnel, stocks, orders_f, sales_f)
    print(f"  Строк: {len(df)}  Колонок: {len(df.columns)}")

    print("\n  Записываю в Google Sheets ...")
    try:
        write_sheet(ss, df, status_summary)
    except Exception as e:
        print(f"  ✗ Ошибка записи: {e}")
        import traceback; traceback.print_exc()
        return

    print("\n" + "═"*60)
    print(f"  ✓ Готово!  Лист: «{SHEET_TITLE}»  Строк: {len(df)}")
    print("═"*60 + "\n")


main()
