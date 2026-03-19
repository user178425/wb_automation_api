# ╔══════════════════════════════════════════════════════════════════╗
# ║  WB Заказы → Автоматизация_ЗаказыWB  |  v2.2                  ║
# ║                                                                  ║
# ║  ИСПРАВЛЕНИЯ v2.2:                                              ║
# ║  ✓ SHEET_TITLE = "Автоматизация_ЗаказыWB" (по ТЗ)             ║
# ║  ✓ ws.update(values=..., range_name=...) — без DeprecationWarning║
# ║  ✓ value_input_option="USER_ENTERED" во всех записях            ║
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

# credentials.json восстанавливается из секрета GOOGLE_JSON
_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
_tmp.write(os.environ["GOOGLE_JSON"])
_tmp.flush()
CREDENTIALS_PATH = _tmp.name

# ✅ ИСПРАВЛЕНО: имя листа строго по ТЗ
SHEET_TITLE = "Автоматизация_ЗаказыWB"

# ─── Даты периодов ────────────────────────────────────────────────
# statistics-api принимает ISO: YYYY-MM-DD  или  YYYY-MM-DDT00:00:00
# seller-analytics-api (воронка) принимает DD.MM.YYYY
#
# Автовычисление: берём прошлую завершённую неделю (Пн–Вс) по МСК
# Можно задать вручную — раскомментируй и впиши нужные даты:
# CUR_FROM_ISO  = "2026-03-02"
# CUR_TO_ISO    = "2026-03-08"
# PAST_FROM_ISO = "2026-02-23"
# PAST_TO_ISO   = "2026-03-01"

import pytz as _pytz
from datetime import timedelta as _td

def _last_monday_msk() -> "datetime":
    tz  = _pytz.timezone("Europe/Moscow")
    now = datetime.now(tz)
    # Если понедельник и ещё нет 6:00 — предыдущая неделя
    if now.weekday() == 0 and now.hour < 6:
        now -= _td(days=1)
    cur_mon = now - _td(days=now.weekday())
    return (cur_mon - _td(weeks=1)).replace(hour=0, minute=0, second=0, microsecond=0)

_MON = _last_monday_msk()

# ISO-формат для statistics-api
CUR_FROM_ISO  = _MON.strftime("%Y-%m-%d")
CUR_TO_ISO    = (_MON + _td(days=6)).strftime("%Y-%m-%d")
PAST_FROM_ISO = (_MON - _td(weeks=1)).strftime("%Y-%m-%d")
PAST_TO_ISO   = (_MON - _td(days=1)).strftime("%Y-%m-%d")

# DD.MM.YYYY для seller-analytics-api (воронка)
CUR_FROM_RU   = _MON.strftime("%d.%m.%Y")
CUR_TO_RU     = (_MON + _td(days=6)).strftime("%d.%m.%Y")
PAST_FROM_RU  = (_MON - _td(weeks=1)).strftime("%d.%m.%Y")
PAST_TO_RU    = (_MON - _td(days=1)).strftime("%d.%m.%Y")

# Совместимые псевдонимы (statistics-api)
CUR_FROM  = CUR_FROM_ISO
CUR_TO    = CUR_TO_ISO
PAST_FROM = PAST_FROM_ISO
PAST_TO   = PAST_TO_ISO

# Даты для statistics-api с временем
CUR_FROM_DT = f"{CUR_FROM_ISO}T00:00:00"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS = [
    "Артикул продавца",                                        #  1
    "Номенклатура",                                            #  2
    "Название",                                                #  3
    "Категория",                                               #  4
    "Бренд",                                                   #  5
    "Удаленный товар",                                         #  6
    "Рейтинг карточки",                                        #  7
    "Рейтинг по отзывам",                                      #  8
    "Увидели карточку",                                        #  9
    "Увидели карточку (предыдущий период)",                    # 10
    "Конверсия в переход",                                     # 11
    "Конверсия в переход (предыдущий период)",                 # 12
    "Доля карточки в выручке",                                 # 13
    "Доля карточки в выручке (предыдущий период)",             # 14
    "Переходы в карточку",                                     # 15
    "Переходы в карточку (предыдущий период)",                 # 16
    "Положили в корзину",                                      # 17
    "Положили в корзину (предыдущий период)",                  # 18
    "Добавили в отложенные",                                   # 19
    "Добавили в отложенные (предыдущий период)",               # 20
    "Заказали, шт",                                            # 21
    "Заказали, шт (предыдущий период)",                        # 22
    "Заказали ВБ клуб, шт",                                    # 23
    "Заказали ВБ клуб, шт (предыдущий период)",                # 24
    "Выкупили, шт",                                            # 25
    "Выкупы, шт (предыдущий период)",                          # 26
    "Выкупили ВБ клуб, шт",                                    # 27
    "Выкупы ВБ клуб, шт (предыдущий период)",                  # 28
    "Отменили, шт",                                            # 29
    "Отменили, шт (предыдущий период)",                        # 30
    "Отменили ВБ клуб, шт",                                    # 31
    "Отменили ВБ клуб, шт (предыдущий период)",                # 32
    "Конверсия в корзину, %",                                  # 33
    "Конверсия в корзину, % (предыдущий период)",              # 34
    "Конверсия в заказ, %",                                    # 35
    "Конверсия в заказ, % (предыдущий период)",                # 36
    "Процент выкупа",                                          # 37
    "Процент выкупа (предыдущий период)",                      # 38
    "Процент выкупа ВБ клуб",                                  # 39
    "Процент выкупа ВБ клуб (предыдущий период)",              # 40
    "Заказали на сумму, руб",                                  # 41
    "Заказали на сумму, руб (предыдущий период)",              # 42
    "Заказали на сумму ВБ клуб, ₽",                            # 43
    "Заказали на сумму ВБ клуб, ₽ (предыдущий период)",        # 44
    "Динамика суммы заказов, руб",                             # 45
    "Динамика суммы заказов ВБ клуб, ₽",                       # 46
    "Выкупили на сумму, руб",                                  # 47
    "Выкупили на сумму, руб (предыдущий период)",              # 48
    "Выкупили на сумму ВБ клуб, ₽",                            # 49
    "Выкупили на сумму ВБ клуб, ₽ (предыдущий период)",        # 50
    "Отменили на сумму, руб",                                  # 51
    "Отменили на сумму, руб (предыдущий период)",              # 52
    "Отменили на сумму ВБ клуб, ₽",                            # 53
    "Отменили на сумму ВБ клуб, ₽ (предыдущий период)",        # 54
    "Средняя цена, р",                                         # 55
    "Средняя цена, руб (предыдущий период)",                   # 56
    "Средняя цена ВБ клуб, ₽",                                 # 57
    "Средняя цена ВБ клуб, ₽ (предыдущий период)",             # 58
    "Среднее количество заказов в день, шт",                   # 59
    "Среднее количество заказов в день, шт (предыдущий период)",# 60
    "Среднее количество заказов в день ВБ клуб, шт",           # 61
    "Среднее количество заказов в день ВБ клуб, шт (предыдущий период)", # 62
    "Остатки склад ВБ, шт",                                    # 63
    "Остатки МП, шт",                                          # 64
    "Сумма остатков на складах, руб",                          # 65
    "Среднее время доставки",                                  # 66
    "Среднее время доставки (предыдущий период)",              # 67
    "Локальные заказы, %",                                     # 68
    "Локальные заказы, % (предыдущий период)",                 # 69
]
assert len(HEADERS) == 69, f"Ожидается 69 колонок, сейчас {len(HEADERS)}"

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

# ════════════════════════════════════════════════════════════════════
#  МЕТОД 1: ВОРОНКА v3
#  POST /api/analytics/v3/sales-funnel/products
# ════════════════════════════════════════════════════════════════════

FUNNEL_URL = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

def fetch_funnel(nm_ids: List[int]) -> Dict[int, Dict]:
    print("\n── МЕТОД 1: Воронка v3 ──────────────────────────────────")
    print(f"  nmIds в запросе : {len(nm_ids)} штук")
    print(f"  Период          : {CUR_FROM_ISO} — {CUR_TO_ISO}")
    print(f"  Прошлый период  : {PAST_FROM_ISO} — {PAST_TO_ISO}")
    print(f"  Формат полей    : start/end  (YYYY-MM-DD)")
    if nm_ids:
        print(f"  Первые 10 nmId  : {nm_ids[:10]}")
    result: Dict[int, Dict] = {}

    # Воронка требует строго: поля "start"/"end" + формат YYYY-MM-DD
    # (подтверждено ошибкой: "start (field required)", cannot parse "DD.MM.YYYY" as "2006-01-02")
    PERIOD_VARIANTS = [
        # ✅ Единственно верный формат по ответу API
        {
            "selectedPeriod": {"start": CUR_FROM_ISO, "end": CUR_TO_ISO},
            "pastPeriod":     {"start": PAST_FROM_ISO, "end": PAST_TO_ISO},
        },
    ]

    try:
        for v_idx, period_block in enumerate(PERIOD_VARIANTS, 1):
            print(f"\n  Вариант тела #{v_idx}: "
                  f"selectedPeriod={period_block['selectedPeriod']}")
            variant_got_items = False

            for page in range(1, 51):
                body = {
                    **period_block,
                    "nmIds": nm_ids,
                    "page":  page,
                }
                print(f"  page={page} ...", end="", flush=True)
                r = safe_post(f"Воронка v{v_idx} p{page}", FUNNEL_URL, body, timeout=60)

                if r is None or r.status_code in (401, 403):
                    print(f"\n  ✗ Нет доступа (HTTP {r.status_code if r else '—'})")
                    return result   # токен не работает — дальше не пробуем
                if r.status_code == 204:
                    print("\n  (нет данных — 204)")
                    break
                if r.status_code != 200:
                    print(f"\n  ✗ HTTP {r.status_code}: {r.text[:300]}")
                    break

                try:
                    raw = r.json()
                except Exception as e:
                    print(f"\n  ✗ JSON: {e}"); break

                if page == 1 and v_idx == 1:
                    _diag_funnel(raw)

                items    = _pick_items(raw)
                has_more = _pick_has_more(raw)
                print(f" товаров: {len(items)} | hasMore: {has_more}")

                if not items:
                    print("  ⚠ Пустой список. Полный ответ сервера:")
                    import json as _json
                    try:    print(_json.dumps(raw, ensure_ascii=False, indent=2)[:2000])
                    except: print(repr(raw)[:2000])
                    break   # пробуем следующий вариант тела

                variant_got_items = True
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

            if variant_got_items:
                print(f"\n  ✓ Данные получены на варианте #{v_idx}")
                break   # нашли рабочий вариант — дальше не идём

    except Exception as e:
        print(f"  ✗ Воронка упала с ошибкой: {e}")

    print(f"\n  ▶ Воронка итого: {len(result)} товаров")
    return result


def _pct(num: float, den: float) -> str:
    """
    Безопасный процент: num/den*100.
    Результат — ЦЕЛОЕ число (WB округляет до целых в интерфейсе).
    '0' при нулевом знаменателе.
    """
    if den == 0: return "0"
    return str(int(round(num / den * 100)))

def _pct_buyout(buy: float, order: float, cancel: float) -> str:
    """
    Процент выкупа по формуле WB:
    выкупы / (заказы - отмены) * 100
    Знаменатель = заказы минус отмены (не просто заказы).
    """
    den = order - cancel
    if den <= 0: return "0"
    return str(int(round(buy / den * 100)))

def _r2(v: Any) -> str:
    """float с 2 знаками, '0' при None."""
    if not v: return "0"
    try:
        f = round(float(v), 2)
        return str(int(f)) if f == int(f) else str(f)
    except: return "0"

def _fmt_delivery(val: Any) -> str:
    """
    Среднее время доставки:
    Если dict с днями/часами — форматируем в строку "Xд Yч".
    Если число (секунды или часы) — пытаемся разобрать.
    '0' или '-' при нулевом/пустом значении.
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


def _parse_funnel_item(item: Dict) -> tuple:
    prod  = item.get("product") or item
    nm_id = int(prod.get("nmId", 0) or prod.get("nmID", 0) or 0)
    if not nm_id:
        return 0, {}

    stat    = item.get("statistic") or item.get("statistics") or {}
    sel     = stat.get("selected")  or stat.get("selectedPeriod") or {}
    pst     = stat.get("past")      or stat.get("previousPeriod") or {}
    # wbClub и conversions: берём если придут, иначе 0 — см. комментарии ниже
    sel_wb  = sel.get("wbClub") or {}
    pst_wb  = pst.get("wbClub") or {}

    # Реальные ключи из диагностики API:
    # selected keys: period, openCount, cartCount, orderCount, orderSum,
    #   buyoutCount, buyoutSum, cancelCount, cancelSum, avgPrice,
    #   avgOrdersCountPerDay, shareOrderPercent
    # product keys: nmId, title, vendorCode, brandName, subjectId,
    #   subjectName, tags, productRating
    # Нет: viewCount, openCardCount, wbClub, timeToReady, conversions — обрабатываем отдельно

    # ── Просмотры / переходы ────────────────────────────────────────
    # openCount = переходы в карточку (клики)
    # viewCount / openCardCount в этой версии API отсутствует →
    #   "Увидели карточку" = openCount (лучшее приближение)
    opens_s = _n(sel, "openCount")   # переходы тек.
    opens_p = _n(pst, "openCount")   # переходы пред.
    views_s = opens_s                 # нет отдельного поля показов — используем переходы
    views_p = opens_p

    # Конверсия в переход = не вычисляется без просмотров (показов нет) → 0
    conv_open_s = "0"
    conv_open_p = "0"

    # ── Корзина / отложенные ────────────────────────────────────────
    cart_s = _n(sel, "cartCount")
    cart_p = _n(pst, "cartCount")
    wish_s = _n(sel, ("addToWishlist", "wishlistCount"))
    wish_p = _n(pst, ("addToWishlist", "wishlistCount"))

    # ── Заказы шт ───────────────────────────────────────────────────
    ord_c_s = _n(sel, "orderCount")
    ord_c_p = _n(pst, "orderCount")
    # ВБ Клуб: в этой версии API нет отдельного блока wbClub → 0
    ord_c_wb_s = 0.0
    ord_c_wb_p = 0.0

    # ── Выкупы шт ───────────────────────────────────────────────────
    buy_c_s    = _n(sel, "buyoutCount")
    buy_c_p    = _n(pst, "buyoutCount")
    buy_c_wb_s = 0.0
    buy_c_wb_p = 0.0

    # ── Отмены шт ───────────────────────────────────────────────────
    can_c_s    = _n(sel, "cancelCount")
    can_c_p    = _n(pst, "cancelCount")
    can_c_wb_s = 0.0
    can_c_wb_p = 0.0

    # ── Суммы заказов ───────────────────────────────────────────────
    ord_s_s    = _n(sel, "orderSum")
    ord_s_p    = _n(pst, "orderSum")
    ord_s_wb_s = 0.0
    ord_s_wb_p = 0.0

    # ── Суммы выкупов ───────────────────────────────────────────────
    buy_s_s    = _n(sel, "buyoutSum")
    buy_s_p    = _n(pst, "buyoutSum")
    buy_s_wb_s = 0.0
    buy_s_wb_p = 0.0

    # ── Суммы отмен ─────────────────────────────────────────────────
    can_s_s    = _n(sel, "cancelSum")
    can_s_p    = _n(pst, "cancelSum")
    can_s_wb_s = 0.0
    can_s_wb_p = 0.0

    # ── Средние ─────────────────────────────────────────────────────
    avg_p_s    = _n(sel, "avgPrice")
    avg_p_p    = _n(pst, "avgPrice")
    avg_p_wb_s = 0.0
    avg_p_wb_p = 0.0

    avg_o_s    = _n(sel, "avgOrdersCountPerDay")
    avg_o_p    = _n(pst, "avgOrdersCountPerDay")
    avg_o_wb_s = 0.0
    avg_o_wb_p = 0.0

    # ── Конверсии ───────────────────────────────────────────────────
    # Нет блока conversions в ответе → считаем сами
    conv_cart_s  = _pct(cart_s, opens_s)
    conv_cart_p  = _pct(cart_p, opens_p)
    conv_ord_s   = _pct(ord_c_s, cart_s)
    conv_ord_p   = _pct(ord_c_p, cart_p)
    # Процент выкупа по формуле WB: выкупы / (заказы - отмены) * 100
    pct_buy_s    = _pct_buyout(buy_c_s, ord_c_s, can_c_s)
    pct_buy_p    = _pct_buyout(buy_c_p, ord_c_p, can_c_p)
    pct_buy_wb_s = "0"
    pct_buy_wb_p = "0"

    # ── Доля в выручке ──────────────────────────────────────────────
    share_s = _r2(_n(sel, "shareOrderPercent"))
    share_p = _r2(_n(pst, "shareOrderPercent"))

    # ── Динамика ────────────────────────────────────────────────────
    dynamics = _i(ord_s_s - ord_s_p)
    dyn_wb   = "0"

    # ── Время доставки ──────────────────────────────────────────────
    # timeToReady отсутствует в этой версии API → берём из остатков если придёт
    del_s = _fmt_delivery(sel.get("timeToReady") or sel.get("deliveryTime") or 0)
    del_p = _fmt_delivery(pst.get("timeToReady") or pst.get("deliveryTime") or 0)

    # ── Локальные заказы ────────────────────────────────────────────
    local_s = _r2(_n(sel, ("localizationPercent", "localPercent")))
    local_p = _r2(_n(pst, ("localizationPercent", "localPercent")))

    # ── tot_buy_c для признака "Удалённый товар" ────────────────────
    tot_buy_c = buy_c_s

    return nm_id, {
        # Идентификация
        "vendor":       _s(prod, "vendorCode"),
        "nm_id":        nm_id,
        "title":        _s(prod, ("title", "name")),
        "category":     _s(prod, ("subjectName",)),
        "brand":        _s(prod, "brandName"),
        "rating_card":  _r2(_n(prod, ("productRating", "rating"))),
        "rating_fb":    _r2(_n(prod, ("feedbackRating", "reviewRating"))),
        # Видимость
        "views_s":      _i(views_s),
        "views_p":      _i(views_p),
        "conv_open_s":  conv_open_s,
        "conv_open_p":  conv_open_p,
        "share_s":      share_s,
        "share_p":      share_p,
        # Переходы / корзина / отложенные
        "opens_s":      _i(opens_s),
        "opens_p":      _i(opens_p),
        "cart_s":       _i(cart_s),
        "cart_p":       _i(cart_p),
        "wish_s":       _i(wish_s),
        "wish_p":       _i(wish_p),
        # Заказы шт
        "ord_c_s":      _i(ord_c_s),
        "ord_c_p":      _i(ord_c_p),
        "ord_c_wb_s":   _i(ord_c_wb_s),
        "ord_c_wb_p":   _i(ord_c_wb_p),
        # Выкупы шт
        "buy_c_s":      _i(buy_c_s),
        "buy_c_p":      _i(buy_c_p),
        "buy_c_wb_s":   _i(buy_c_wb_s),
        "buy_c_wb_p":   _i(buy_c_wb_p),
        # Отмены шт
        "can_c_s":      _i(can_c_s),
        "can_c_p":      _i(can_c_p),
        "can_c_wb_s":   _i(can_c_wb_s),
        "can_c_wb_p":   _i(can_c_wb_p),
        # Конверсии
        "conv_cart_s":  conv_cart_s,
        "conv_cart_p":  conv_cart_p,
        "conv_ord_s":   conv_ord_s,
        "conv_ord_p":   conv_ord_p,
        "pct_buy_s":    pct_buy_s,
        "pct_buy_p":    pct_buy_p,
        "pct_buy_wb_s": pct_buy_wb_s,
        "pct_buy_wb_p": pct_buy_wb_p,
        # Суммы заказов
        "ord_s_s":      _i(ord_s_s),
        "ord_s_p":      _i(ord_s_p),
        "ord_s_wb_s":   _i(ord_s_wb_s),
        "ord_s_wb_p":   _i(ord_s_wb_p),
        # Динамика
        "dynamics":     dynamics,
        "dyn_wb":       dyn_wb,
        # Суммы выкупов
        "buy_s_s":      _i(buy_s_s),
        "buy_s_p":      _i(buy_s_p),
        "buy_s_wb_s":   _i(buy_s_wb_s),
        "buy_s_wb_p":   _i(buy_s_wb_p),
        # Суммы отмен
        "can_s_s":      _i(can_s_s),
        "can_s_p":      _i(can_s_p),
        "can_s_wb_s":   _i(can_s_wb_s),
        "can_s_wb_p":   _i(can_s_wb_p),
        # Средние
        "avg_p_s":      _i(avg_p_s),
        "avg_p_p":      _i(avg_p_p),
        "avg_p_wb_s":   _i(avg_p_wb_s),
        "avg_p_wb_p":   _i(avg_p_wb_p),
        "avg_o_s":      _r2(avg_o_s),
        "avg_o_p":      _r2(avg_o_p),
        "avg_o_wb_s":   _r2(avg_o_wb_s),
        "avg_o_wb_p":   _r2(avg_o_wb_p),
        # Для удалённого товара
        "tot_buy_c":    tot_buy_c,
        # Доставка / локализация
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
    pst  = stat.get("past")     or stat.get("previousPeriod") or {}
    print(f"\n  [ДИА] product keys : {list(prod.keys())}")
    print(f"  [ДИА] stat keys    : {list(stat.keys())}")
    print(f"  [ДИА] selected keys: {list(sel.keys())}")
    # Выводим значения для диагностики расхождений
    diag_keys = ["openCount","viewCount","openCardCount","cartCount",
                 "orderCount","orderSum","buyoutCount","buyoutSum",
                 "cancelCount","cancelSum","avgPrice","avgOrdersCountPerDay",
                 "shareOrderPercent","timeToReady","wbClub"]
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
#  МЕТОД 1б: NM-REPORT/DETAIL — показы (Увидели карточку) + время доставки
#
#  /v2/nm-report/detail возвращает:
#    openCardCount  = ПОКАЗЫ карточки (то что WB называет «Увидели карточку»)
#    addToCartCount = корзина
#    ordersCount    = заказы
#    buyoutsCount   = выкупы
#    cancelCount    = отмены
#    avgRubPrice    = средняя цена
#    timeToReady    = время доставки
#
#  Запрашиваем ДВА раза: за текущий период и за прошлый период отдельно.
# ════════════════════════════════════════════════════════════════════

NM_REPORT_URL = "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail"

def _fetch_nm_report_for_period(nm_ids: List[int],
                                 date_from: str, date_to: str,
                                 label: str) -> Dict[int, Dict]:
    """Один запрос nm-report за указанный период. Возвращает {nmId: stat_dict}."""
    result: Dict[int, Dict] = {}
    BATCH = 20
    for start in range(0, len(nm_ids), BATCH):
        batch = nm_ids[start: start + BATCH]
        body = {
            "nmIDs":    batch,
            "timezone": "Europe/Moscow",
            "period": {
                "begin": date_from + " 00:00:00",
                "end":   date_to   + " 23:59:59",
            },
            "orderBy": {"field": "ordersSumRub", "mode": "desc"},
            "page": 1,
        }
        r = safe_post(f"nm-report {label} батч {start//BATCH+1}",
                      NM_REPORT_URL, body, timeout=60)
        if r is None or r.status_code != 200:
            if r: print(f"  HTTP {r.status_code}: {r.text[:300]}")
            continue

        try:
            raw = r.json()
        except Exception as e:
            print(f"  JSON: {e}"); continue

        # Диагностика первого батча первого периода
        if start == 0 and label == "тек":
            _diag_nm_report(raw)

        cards = []
        if isinstance(raw, dict):
            data  = raw.get("data") or {}
            cards = (data.get("cards") or data.get("items") or [])
            if not cards: cards = raw.get("cards") or []
        elif isinstance(raw, list):
            cards = raw

        for card in cards:
            if not isinstance(card, dict): continue
            nm = card.get("nmID") or card.get("nmId")
            if not nm: continue
            nm = int(nm)

            # Пробуем разные вложенности ответа
            stat = (card.get("statistics") or {})
            sp   = (stat.get("selectedPeriod") or stat.get("period") or
                    card.get("selectedPeriod") or card)

            result[nm] = {
                # openCardCount = ПОКАЗЫ карточки («Увидели карточку» в интерфейсе WB)
                "openCardCount": int(_n(sp, ("openCardCount", "viewCount")) or 0),
                "addToCart":     int(_n(sp, ("addToCartCount", "cartCount")) or 0),
                "ordersCount":   int(_n(sp, ("ordersCount",   "orderCount")) or 0),
                "buyoutsCount":  int(_n(sp, ("buyoutsCount",  "buyoutCount")) or 0),
                "cancelCount":   int(_n(sp, ("cancelCount",)) or 0),
                "timeToReady":   sp.get("timeToReady") or sp.get("deliveryTime") or 0,
            }
        time.sleep(0.3)

    return result


def _diag_nm_report(raw: Any) -> None:
    """Диагностика структуры nm-report/detail."""
    cards = []
    if isinstance(raw, dict):
        data  = raw.get("data") or {}
        cards = data.get("cards") or data.get("items") or []
    if cards and isinstance(cards[0], dict):
        c0   = cards[0]
        stat = c0.get("statistics") or {}
        sp   = stat.get("selectedPeriod") or stat.get("period") or c0.get("selectedPeriod") or {}
        print(f"\n  [ДИА nm-report] card keys     : {list(c0.keys())[:10]}")
        print(f"  [ДИА nm-report] stat keys     : {list(stat.keys())}")
        print(f"  [ДИА nm-report] selectedPeriod: {list(sp.keys())}")
        print(f"  [ДИА nm-report] openCardCount : {sp.get('openCardCount','—')}")
        print(f"  [ДИА nm-report] timeToReady   : {sp.get('timeToReady','—')}")


def fetch_nm_report(nm_ids: List[int]) -> Dict[int, Dict]:
    """
    Запрашивает nm-report за текущий и прошлый периоды.
    Возвращает {nmId: {"views_s", "views_p", "del_s", "del_p"}}.
    """
    print("\n── МЕТОД 1б: nm-report/detail (показы + время доставки) ──")
    if not nm_ids:
        print("  nmIds пусты — пропускаю"); return {}

    data_s = _fetch_nm_report_for_period(nm_ids, CUR_FROM_ISO,  CUR_TO_ISO,  "тек")
    data_p = _fetch_nm_report_for_period(nm_ids, PAST_FROM_ISO, PAST_TO_ISO, "пред")

    result: Dict[int, Dict] = {}
    all_ids = set(data_s) | set(data_p)
    for nm in all_ids:
        s = data_s.get(nm, {})
        p = data_p.get(nm, {})
        result[nm] = {
            "views_s": str(s.get("openCardCount", 0)),
            "views_p": str(p.get("openCardCount", 0)),
            "del_s":   _fmt_delivery(s.get("timeToReady", 0)),
            "del_p":   _fmt_delivery(p.get("timeToReady", 0)),
        }

    print(f"  ▶ nm-report: {len(result)} товаров (тек: {len(data_s)}, пред: {len(data_p)})")
    return result

STOCKS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/stocks"

def fetch_stocks() -> tuple:
    """
    Остатки на дату КОНЦА отчётного периода (CUR_TO_ISO).
    quantity    = склад WB (FBO) — это «Остатки склад ВБ»
    quantityFull = склад продавца (FBS) — WB в UI не показывает, ставим 0
    Сумма = Price * quantity (без скидки — Price уже розничная цена)
    """
    date_dt = f"{CUR_TO_ISO}T23:59:59"
    print("\n── ШАГ 1: Остатки (сбор nmId) ──────────────────────────")
    print(f"  dateFrom: {date_dt}  (конец отчётного периода)")
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
                # quantity = FBO (склад WB) — именно это показывает WB в интерфейсе
                wb  = int(item.get("quantity", 0) or 0)
                # quantityFull = FBS (склад продавца) — в ручной выгрузке WB = 0
                # Оставляем для информации, но в таблицу пишем 0 (как в WB UI)
                # Price в /stocks — это уже цена с учётом настроек (не надо применять Discount)
                prc = float(item.get("Price", 0) or 0)
                val = prc * wb
                if nm not in out:
                    out[nm] = {"wb": 0, "mp": 0, "sum_rub": 0.0}
                out[nm]["wb"]      += wb
                # МП = 0 (FBS не отображается в WB аналитике как отдельный склад)
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
#  МЕТОД 3: ЗАКАЗЫ (СВЕРКА)
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
#  МЕТОД 4: ПРОДАЖИ (СВЕРКА)
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

def build_df(funnel, stocks, nm_report, orders_f, sales_f) -> pd.DataFrame:
    rows = []
    for nm_id, f in funnel.items():
        try:
            stk    = stocks.get(nm_id, {})
            stk_wb = stk.get("wb", 0)
            stk_mp = stk.get("mp", 0)
            stk_ru = stk.get("sum_rub", 0)
            nm_r   = nm_report.get(nm_id, {})

            deleted = "Да" if (stk_wb == 0 and stk_mp == 0
                                and f.get("tot_buy_c", 0) == 0) else "Нет"

            # views: из nm_report (приоритет) или из воронки
            views_s = nm_r.get("views_s") or f.get("views_s", "0")
            views_p = nm_r.get("views_p") or f.get("views_p", "0")

            # Пересчитываем конверсию в переход с реальными просмотрами
            try:
                v_s = float(views_s); v_p = float(views_p)
                o_s = float(f.get("opens_s", "0") or 0)
                o_p = float(f.get("opens_p", "0") or 0)
                conv_open_s = _pct(o_s, v_s)
                conv_open_p = _pct(o_p, v_p)
            except Exception:
                conv_open_s = f.get("conv_open_s", "0")
                conv_open_p = f.get("conv_open_p", "0")

            # время доставки из nm_report (приоритет) или из воронки
            del_s = nm_r.get("del_s") or f.get("del_s", "0")
            del_p = nm_r.get("del_p") or f.get("del_p", "0")

            row = [
                f.get("vendor",       ""),       #  1 Артикул продавца
                str(nm_id),                       #  2 Номенклатура
                f.get("title",        ""),        #  3 Название
                f.get("category",     ""),        #  4 Категория
                f.get("brand",        ""),        #  5 Бренд
                deleted,                          #  6 Удаленный товар
                f.get("rating_card",  "0"),       #  7 Рейтинг карточки
                f.get("rating_fb",    "0"),       #  8 Рейтинг по отзывам
                views_s,                          #  9 Увидели карточку
                views_p,                          # 10 Увидели карточку (пред.)
                conv_open_s,                      # 11 Конверсия в переход
                conv_open_p,                      # 12 Конверсия в переход (пред.)
                f.get("share_s",      "0"),       # 13 Доля карточки в выручке
                f.get("share_p",      "0"),       # 14 Доля карточки в выручке (пред.)
                f.get("opens_s",      "0"),       # 15 Переходы в карточку
                f.get("opens_p",      "0"),       # 16 Переходы в карточку (пред.)
                f.get("cart_s",       "0"),       # 17 Положили в корзину
                f.get("cart_p",       "0"),       # 18 Положили в корзину (пред.)
                f.get("wish_s",       "0"),       # 19 Добавили в отложенные
                f.get("wish_p",       "0"),       # 20 Добавили в отложенные (пред.)
                f.get("ord_c_s",      "0"),       # 21 Заказали, шт
                f.get("ord_c_p",      "0"),       # 22 Заказали, шт (пред.)
                f.get("ord_c_wb_s",   "0"),       # 23 Заказали ВБ клуб, шт
                f.get("ord_c_wb_p",   "0"),       # 24 Заказали ВБ клуб, шт (пред.)
                f.get("buy_c_s",      "0"),       # 25 Выкупили, шт
                f.get("buy_c_p",      "0"),       # 26 Выкупы, шт (пред.)
                f.get("buy_c_wb_s",   "0"),       # 27 Выкупили ВБ клуб, шт
                f.get("buy_c_wb_p",   "0"),       # 28 Выкупы ВБ клуб, шт (пред.)
                f.get("can_c_s",      "0"),       # 29 Отменили, шт
                f.get("can_c_p",      "0"),       # 30 Отменили, шт (пред.)
                f.get("can_c_wb_s",   "0"),       # 31 Отменили ВБ клуб, шт
                f.get("can_c_wb_p",   "0"),       # 32 Отменили ВБ клуб, шт (пред.)
                f.get("conv_cart_s",  "0"),       # 33 Конверсия в корзину, %
                f.get("conv_cart_p",  "0"),       # 34 Конверсия в корзину, % (пред.)
                f.get("conv_ord_s",   "0"),       # 35 Конверсия в заказ, %
                f.get("conv_ord_p",   "0"),       # 36 Конверсия в заказ, % (пред.)
                f.get("pct_buy_s",    "0"),       # 37 Процент выкупа
                f.get("pct_buy_p",    "0"),       # 38 Процент выкупа (пред.)
                f.get("pct_buy_wb_s", "0"),       # 39 Процент выкупа ВБ клуб
                f.get("pct_buy_wb_p", "0"),       # 40 Процент выкупа ВБ клуб (пред.)
                f.get("ord_s_s",      "0"),       # 41 Заказали на сумму, руб
                f.get("ord_s_p",      "0"),       # 42 Заказали на сумму, руб (пред.)
                f.get("ord_s_wb_s",   "0"),       # 43 Заказали на сумму ВБ клуб, ₽
                f.get("ord_s_wb_p",   "0"),       # 44 Заказали на сумму ВБ клуб, ₽ (пред.)
                f.get("dynamics",     "0"),       # 45 Динамика суммы заказов, руб
                f.get("dyn_wb",       "0"),       # 46 Динамика суммы заказов ВБ клуб, ₽
                f.get("buy_s_s",      "0"),       # 47 Выкупили на сумму, руб
                f.get("buy_s_p",      "0"),       # 48 Выкупили на сумму, руб (пред.)
                f.get("buy_s_wb_s",   "0"),       # 49 Выкупили на сумму ВБ клуб, ₽
                f.get("buy_s_wb_p",   "0"),       # 50 Выкупили на сумму ВБ клуб, ₽ (пред.)
                f.get("can_s_s",      "0"),       # 51 Отменили на сумму, руб
                f.get("can_s_p",      "0"),       # 52 Отменили на сумму, руб (пред.)
                f.get("can_s_wb_s",   "0"),       # 53 Отменили на сумму ВБ клуб, ₽
                f.get("can_s_wb_p",   "0"),       # 54 Отменили на сумму ВБ клуб, ₽ (пред.)
                f.get("avg_p_s",      "0"),       # 55 Средняя цена, р
                f.get("avg_p_p",      "0"),       # 56 Средняя цена, руб (пред.)
                f.get("avg_p_wb_s",   "0"),       # 57 Средняя цена ВБ клуб, ₽
                f.get("avg_p_wb_p",   "0"),       # 58 Средняя цена ВБ клуб, ₽ (пред.)
                f.get("avg_o_s",      "0"),       # 59 Среднее кол-во заказов в день, шт
                f.get("avg_o_p",      "0"),       # 60 Среднее кол-во заказов в день, шт (пред.)
                f.get("avg_o_wb_s",   "0"),       # 61 Среднее кол-во заказов в день ВБ клуб, шт
                f.get("avg_o_wb_p",   "0"),       # 62 Среднее кол-во заказов в день ВБ клуб, шт (пред.)
                _i(stk_wb),                       # 63 Остатки склад ВБ, шт
                _i(stk_mp),                       # 64 Остатки МП, шт
                _i(stk_ru),                       # 65 Сумма остатков на складах, руб
                del_s,                            # 66 Среднее время доставки
                del_p,                            # 67 Среднее время доставки (пред.)
                f.get("local_s",      "0"),       # 68 Локальные заказы, %
                f.get("local_p",      "0"),       # 69 Локальные заказы, % (пред.)
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
#  ✅ ИСПРАВЛЕНО: ws.update(values=..., range_name=...) везде
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

    # ✅ Новый формат: ws.update(values=..., range_name=...)
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

    ws.update(
        values=[HEADERS],
        range_name=f"A5:{lc}5",
    )
    format_cell_range(ws, f"A5:{lc}5", FMT_H)

    if n == 0:
        ws.update(
            values=[["Воронка не вернула данных — смотри консоль"]],
            range_name="A6",
        )
        format_cell_range(ws, "A6", FMT_E)
        return

    rows = [
        [("0" if (v is None or str(v) in ("nan", "")) else str(v)) for v in row]
        for row in df.values.tolist()
    ]
    for i in range(0, n, 500):
        chunk = data_chunk = rows[i: i+500]
        r0, r1 = 6+i, 5+i+len(chunk)
        ws.update(
            values=chunk,
            range_name=f"A{r0}:{lc}{r1}",
        )
        time.sleep(0.5)

    format_cell_range(ws, f"A6:{lc}{5+n}", FMT_D)
    print(f"\n  ✓ «{SHEET_TITLE}»: {n} строк × {len(HEADERS)} колонок")

# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*60)
    print("  WB Заказы  |  Автоматизация_ЗаказыWB  |  v2.8")
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

    # ── Объединяем nmId из ВСЕХ источников ──────────────────────────
    # Остатки дают только товары с ненулевыми остатками (10 из 13).
    # Заказы и продажи содержат удалённые / нулевые товары тоже.
    all_nm_ids = sorted(set(nm_ids_stocks)
                        | set(orders_f.keys())
                        | set(sales_f.keys()))
    print(f"\n  ▶ nmId объединено из остатков+заказов+продаж: {len(all_nm_ids)}")
    print(f"  ▶ Только в заказах/продажах (удалённые): "
          f"{sorted(set(all_nm_ids) - set(nm_ids_stocks))}")

    funnel   = fetch_funnel(all_nm_ids)
    nm_report = fetch_nm_report(all_nm_ids)   # viewCount + timeToReady

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
    df = build_df(funnel, stocks, nm_report, orders_f, sales_f)
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
