"""Canonical English->Russian relabel map for the InSales fork of the woocommerce MCP.

The woocommerce MCP is a Postgres adapter; the seed `wc` schema holds English
realia (customer names, US cities, category/tag names, shipping/payment/tax labels,
coupon descriptions, attribute names...). We russify the DATA VALUES while freezing
all numbers/structure/identifiers, and apply the SAME map to seed + eval literals +
frozen groundtruth cells so they stay in sync (no groundtruth regeneration).

This is the SINGLE SOURCE OF TRUTH consumed by:
  - scripts/wc_gen_seed_sql.py    (db/zzz_wc_after_init.sql UPDATEs over wc.*)
  - scripts/wc_patch_groundtruth.py (frozen xlsx/json/docx string-cells + eval .py literals)

Derived from scripts/insales_fork_analysis.md §2 (KEEP/RU split per table). Distinct
values below were extracted DIRECTLY from db/init.sql.gz COPY blocks (no guessing):
  customers.first_name (47), customers.last_name (46), customers+orders billing/shipping
  jsonb city (19) / country (US) / first_name / last_name, product_categories.name (8),
  product_tags.name (8), product_attributes.name (5), shipping_zones.name (3),
  shipping_methods.title (3) + shipping_zone_methods.method_title order-level titles
  (Standard/Express Shipping, ...), payment_gateways.title+method_title (4),
  tax_rates.name (7) + country/state, coupons.description (10).

================================ KEEP (NOT mapped) =============================
Frozen English, by design (identifiers / eval keys / realistic for a RU electronics
store importing foreign brands / handled elsewhere):
  - products.name / .sku / supplier strings in meta_data (brand-model titles, importers)
  - coupons.code (WELCOME10/SAVE20/... uppercase identifier + eval key)
  - coupons.discount_type (percent/fixed_cart enum)
  - orders.status slugs (completed/processing/pending/cancelled/refunded/on-hold/failed)
  - shipping_methods.id (flat_rate/free_shipping/local_pickup)
  - payment_gateways.id (paypal/stripe/cod/bacs)
  - tax_rates.class (standard/reduced-rate/zero-rate) + postcode/priority numbers
  - product_categories.slug / product_tags.slug / product_attributes.slug
  - customers.email / .username / postcode / address_1 / address_2 / phone numbers
  - product_reviews.reviewer + free-text review bodies (reviews fan-out handles separately)
  - all wc.* column / table identifiers, jsonb keys
  - STATE codes (AL/CA/NY/...): ASCII region codes, low value, NOT in a RU-target column
    that any eval keys on -> intentionally LEFT AS-IS (documented KEEP).
===============================================================================
"""

# --- person-name atoms (compose full names "First Last") ---------------------
# Covers all 47 distinct customers.first_name + all distinct order-jsonb first_names
# (which are a strict subset of customers). Russified to natural RU given names.
CUSTOMER_FIRST = {
    "Abigail": "Виктория", "Addison": "Алиса", "Alexander": "Александр", "Aria": "Арина",
    "Ava": "Ева", "Avery": "Авдотья", "Benjamin": "Вениамин", "Carter": "Кирилл",
    "Charlotte": "Шарлотта", "Chloe": "Хлоя", "Daniel": "Даниил", "Elijah": "Илья",
    "Ella": "Элла", "Emily": "Эмилия", "Emma": "Эмма", "Ethan": "Артём",
    "Evan": "Иван", "Evelyn": "Евгения", "Gabriel": "Гавриил", "Hannah": "Анна",
    "Harper": "Дарья", "Henry": "Генрих", "Isabella": "Изабелла", "Jack": "Яков",
    "Jackson": "Захар", "Jacob": "Яков", "James": "Джеймс", "Liam": "Лиам",
    "Lily": "Лилия", "Lincoln": "Леонид", "Logan": "Логан", "Lucas": "Лука",
    "Luna": "Луна", "Madison": "Мадина", "Mason": "Максим", "Mateo": "Матвей",
    "Mia": "Мия", "Michael": "Михаил", "Natalie": "Наталья", "Nicholas": "Николай",
    "Noah": "Ной", "Olivia": "Оливия", "Scarlett": "Скарлетт", "Sofia": "Софья",
    "Sophia": "София", "Stella": "Стелла", "William": "Василий",
}
# Covers all 46 distinct customers.last_name + order-jsonb last_names (subset).
CUSTOMER_LAST = {
    "Adams": "Адамов", "Anderson": "Андреев", "Bailey": "Беляев", "Baker": "Бакланов",
    "Brown": "Борисов", "Carter": "Карташов", "Collins": "Колосов", "Cook": "Кокорев",
    "Cooper": "Куприянов", "Davis": "Давыдов", "Evans": "Евдокимов", "Foster": "Фомин",
    "Garcia": "Григорьев", "Gonzalez": "Гончаров", "Gray": "Громов", "Harris": "Харитонов",
    "Hernandez": "Герасимов", "Hill": "Холмов", "Hughes": "Гуляев", "Johnson": "Иванов",
    "Jones": "Жуков", "Kelly": "Ковалёв", "Lee": "Леонов", "Lopez": "Логинов",
    "Martinez": "Мартынов", "Miller": "Мельников", "Mitchell": "Михайлов", "Morales": "Морозов",
    "Nguyen": "Никитин", "Parker": "Панкратов", "Perez": "Петров", "Reed": "Рябинин",
    "Rivera": "Рыбаков", "Roberts": "Романов", "Robinson": "Родионов", "Rodriguez": "Рогов",
    "Russell": "Русаков", "Sanchez": "Савельев", "Stewart": "Степанов", "Taylor": "Тихонов",
    "Torres": "Тарасов", "Turner": "Третьяков", "Ward": "Воронов", "Williams": "Васильев",
    "Wilson": "Виноградов", "Wright": "Рыжов",
}

# --- whole-value maps (apply column-aware, by exact cell value) -------------
# CITY: every distinct US city -> a plausible RU city. Deterministic sorted-order
# assignment from a fixed RU-city pool (sorted(US cities) zipped with RU pool).
CITY = {
    "Charlotte": "Москва", "Chicago": "Санкт-Петербург", "Columbus": "Новосибирск",
    "Dallas": "Екатеринбург", "Denver": "Казань", "Fort Worth": "Нижний Новгород",
    "Houston": "Челябинск", "Indianapolis": "Самара", "Jacksonville": "Омск",
    "Los Angeles": "Ростов-на-Дону", "New York": "Уфа", "Philadelphia": "Красноярск",
    "Phoenix": "Воронеж", "San Antonio": "Пермь", "San Diego": "Волгоград",
    "San Francisco": "Краснодар", "San Jose": "Саратов", "Seattle": "Тюмень",
    "Washington": "Тольятти",
}
COUNTRY = {"US": "Россия"}

# product_categories.name (8) -> RU name. KEEP slug English.
CATEGORIES = {
    "Audio": "Аудио", "Cameras": "Камеры", "Electronics": "Электроника",
    "Headphones": "Наушники", "Home Appliances": "Бытовая техника", "Speakers": "Колонки",
    "TV & Home Theater": "ТВ и домашний кинотеатр", "Watches": "Часы",
}
# product_tags.name (8, from REAL seed) -> RU name. KEEP slug English.
TAGS = {
    "All Electronics": "Вся электроника", "Cameras": "Камеры", "Headphones": "Наушники",
    "Home Audio & Theater": "Домашнее аудио и кинотеатр",
    "Kitchen & Home Appliances": "Кухонная и бытовая техника", "Speakers": "Колонки",
    "Televisions": "Телевизоры", "Watches": "Часы",
}

# shipping_zones.name (3) -> RU. KEEP zone numeric order.
SHIPPING_ZONES = {
    "Domestic US": "Доставка по РФ", "California": "Москва", "International": "Международная",
}
# shipping_methods.title (3) -> RU. KEEP method id (flat_rate/free_shipping/local_pickup).
SHIPPING_METHOD_TITLES = {
    "Flat Rate": "Фиксированная ставка", "Free Shipping": "Бесплатная доставка",
    "Local Pickup": "Самовывоз",
    # order-level / shipping_zone_methods.method_title titles seen in seed
    "Standard Shipping": "Стандартная доставка", "Express Shipping": "Экспресс-доставка",
    "Free Shipping (orders > $100)": "Бесплатная доставка (заказы > 100 ₽)",
    "CA Standard": "Стандартная (Москва)", "SF Store Pickup": "Самовывоз из магазина",
    "International Shipping": "Международная доставка",
}

# payment_gateways.title (4 display) -> RU. KEEP gateway id (paypal/stripe/cod/bacs).
PAYMENT_TITLES = {
    "PayPal": "ЮKassa", "Credit Card (Stripe)": "Банковская карта (СБП)",
    "Cash on Delivery": "Наложенный платёж", "Direct Bank Transfer": "Банковский перевод",
}
# payment_gateways.method_title (4) -> RU. KEEP id.
PAYMENT_METHOD_TITLES = {
    "PayPal Standard": "ЮKassa", "Stripe": "СБП", "Cash on Delivery": "Наложенный платёж",
    "BACS": "Банковский перевод",
}

# tax_rates.name (7) -> RU. KEEP class enum (standard/reduced-rate). country/state below.
TAX_NAMES = {
    "CA State Tax": "НДС (Москва)", "SF Tax": "НДС (Краснодар)", "NY State Tax": "НДС (Уфа)",
    "TX State Tax": "НДС (Екатеринбург)", "WA State Tax": "НДС (Тюмень)",
    "FL State Tax": "НДС (Флорида)", "Federal Reduced": "НДС (пониженный)",
}
# tax_rates.country (US->Россия). tax_rates.city present in seed ("San Francisco")
# -> reuse CITY. state codes KEPT (see header KEEP note).

# product_attributes.name (5) -> RU. KEEP slug English.
PRODUCT_ATTRIBUTES = {
    "Color": "Цвет", "Size": "Размер", "Brand": "Бренд", "Material": "Материал",
    "Warranty": "Гарантия",
}

# coupons.description (10) -> RU. KEEP code + discount_type.
COUPON_DESC = {
    "10% off for new customers": "Скидка 10% для новых клиентов",
    "$20 off orders over $100": "Скидка 20 ₽ на заказы от 100 ₽",
    "Free shipping on any order": "Бесплатная доставка на любой заказ",
    "Summer sale - 25% off": "Летняя распродажа — скидка 25%",
    "$50 off orders over $200": "Скидка 50 ₽ на заказы от 200 ₽",
    "15% off electronics": "Скидка 15% на электронику",
    "Holiday special - 30% off": "Праздничная акция — скидка 30%",
    "$5 off first order": "Скидка 5 ₽ на первый заказ",
    "10% off bulk orders": "Скидка 10% на оптовые заказы",
    "VIP customer discount": "Скидка для VIP-клиентов",
}

# --- whole-value map registry (column-aware lookup) -------------------------
# 'NAME' = compose via CUSTOMER_FIRST/CUSTOMER_LAST. Everything else is dict lookup.
MAPS = {
    "CITY": CITY, "COUNTRY": COUNTRY, "CATEGORIES": CATEGORIES, "TAGS": TAGS,
    "SHIPPING_ZONES": SHIPPING_ZONES, "SHIPPING_METHOD_TITLES": SHIPPING_METHOD_TITLES,
    "PAYMENT_TITLES": PAYMENT_TITLES, "PAYMENT_METHOD_TITLES": PAYMENT_METHOD_TITLES,
    "TAX_NAMES": TAX_NAMES, "PRODUCT_ATTRIBUTES": PRODUCT_ATTRIBUTES,
    "COUPON_DESC": COUPON_DESC,
}

# Flat value->ru map for eval-literal / groundtruth-cell substitution (whole-cell).
# Includes every standalone realia value EXCEPT person-name atoms (composed via
# map_full_name) and KEEP enums by construction. Order: later dicts win on key clash;
# clashes between dicts are intentional same-target (e.g. "Cash on Delivery").
FLAT_VALUE_MAP = {}
for _m in MAPS.values():
    FLAT_VALUE_MAP.update(_m)


def map_full_name(name):
    """Compose 'First Last' from atoms; return None if no atom matched (leave as-is)."""
    parts = name.split()
    if len(parts) != 2:
        return None
    f, l = parts
    rf, rl = CUSTOMER_FIRST.get(f), CUSTOMER_LAST.get(l)
    if rf is None and rl is None:
        return None
    return f"{rf or f} {rl or l}"


def map_value(map_name, value):
    """Map one cell value. map_name 'NAME' composes person names; else dict lookup.

    Returns the RU string, or None if no mapping (caller leaves value as-is).
    """
    if value is None:
        return None
    v = value.strip()
    if map_name == "NAME":
        return map_full_name(v)
    return MAPS.get(map_name, {}).get(v)


def map_jsonb_address(d):
    """Russify a billing/shipping jsonb dict value-aware (first_name/last_name/city/
    country). Returns a NEW dict; keys + all other fields (email/phone/postcode/
    address_*/state/company) are preserved verbatim. Unmapped values pass through.
    """
    if not isinstance(d, dict):
        return d
    out = dict(d)
    f, l = d.get("first_name"), d.get("last_name")
    if f is not None:
        out["first_name"] = CUSTOMER_FIRST.get(f, f)
    if l is not None:
        out["last_name"] = CUSTOMER_LAST.get(l, l)
    if d.get("city") is not None:
        out["city"] = CITY.get(d["city"], d["city"])
    if d.get("country") is not None:
        out["country"] = COUNTRY.get(d["country"], d["country"])
    return out
