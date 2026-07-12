"""Canonical English->Russian relabel map for the ClickHouse fork of the snowflake MCP.

The snowflake MCP is a Postgres adapter; the seed `sf_data` schema holds English
realia (names, cities, departments...). We russify the DATA VALUES while freezing
all numbers/structure, and apply the SAME map to seed + eval literals + frozen
groundtruth cells so they stay in sync (no groundtruth regeneration). Column/value
identifiers, enums the eval greps, and brand proper-nouns stay English.

Derived from scripts/clickhouse_relabel_analysis.md (§1, §2, §4). KEEP-list there.
"""

# --- whole-value maps (apply column-aware, by exact cell value) -------------
DEPARTMENTS = {
    "Engineering": "Инженерия", "Finance": "Финансы", "HR": "Кадры",
    "Operations": "Операции", "R&D": "НИОКР", "Sales": "Продажи", "Support": "Поддержка",
}
LOCATIONS = {
    "New York": "Москва", "Sydney": "Санкт-Петербург",
    "Singapore": "Новосибирск", "Toronto": "Екатеринбург",
}
REGIONS = {
    "North America": "Северная Америка", "Europe": "Европа",
    "Asia Pacific": "Азиатско-Тихоокеанский регион",
    "Latin America": "Латинская Америка", "Middle East": "Ближний Восток",
}
SEGMENTS = {
    "Consumer": "Частные клиенты", "Enterprise": "Корпоративный",
    "Government": "Государственный", "SMB": "Малый и средний бизнес",
}
EDUCATION = {
    "Bachelor's": "Бакалавр", "Master's": "Магистр", "PhD": "Кандидат наук",
    "Diploma": "Диплом", "High School": "Среднее образование",
}
MARITAL = {"Married": "В браке", "Single": "Не в браке"}
ROLES = {
    "Account Executive": "Менеджер по работе с клиентами", "Analyst": "Аналитик",
    "Customer Specialist": "Специалист по работе с клиентами", "DevOps Engineer": "DevOps-инженер",
    "HR Executive": "Специалист по кадрам", "HR Manager": "Менеджер по кадрам",
    "ML Engineer": "ML-инженер", "Operations Coordinator": "Координатор операций",
    "Ops Manager": "Руководитель операций", "Researcher": "Исследователь",
    "Sales Manager": "Менеджер по продажам", "Scientist": "Научный сотрудник",
    "Senior Analyst": "Старший аналитик", "Software Engineer": "Инженер-программист",
    "Support Engineer": "Инженер поддержки",
}
SHIP_MODE = {"Economy": "Эконом", "Express": "Экспресс", "Next Day": "На следующий день", "Standard": "Стандарт"}
ORDER_STATUS = {"Cancelled": "Отменён", "Delivered": "Доставлен", "Processing": "В обработке", "Shipped": "Отправлен"}
PRODUCT_TAXONOMY = {"tv, audio & cameras": "ТВ, аудио и камеры", "All Electronics": "Вся электроника"}
TICKET_REPORTER = {"Alice": "Алиса", "Bob": "Борис", "Charlie": "Карл", "Emily": "Эмилия", "John": "Иван"}
ISSUE_TYPE = {
    "Bug": "Ошибка", "Feature Request": "Запрос функции", "Incident": "Инцидент",
    "Maintenance": "Обслуживание", "Performance Issue": "Проблема производительности",
    "Service Request": "Запрос обслуживания", "Technical Issue": "Техническая проблема",
}
REPORT_CHANNEL = {"App": "Приложение", "Email": "Электронная почта", "Phone": "Телефон", "Website": "Сайт"}
SHORT_DESCRIPTION = {
    "Application crash": "Сбой приложения", "Data loss issue": "Потеря данных",
    "New feature request": "Запрос новой функции", "Password reset needed": "Сброс пароля",
    "Printer not working": "Принтер не работает", "Server outage": "Сбой сервера",
    "UI glitch": "Сбой интерфейса", "Unable to login": "Невозможно войти",
}
AGENT_NAME = {
    "Ava": "Ева", "David": "Давид", "Emma": "Эмма", "James": "Яков", "Michael": "Михаил",
    "Olivia": "Оливия", "Sam": "Семён", "Sarah": "Сара", "William": "Василий",
}

# --- person-name atoms (compose full names "First Last") ---------------------
FIRST_ATOMS = {
    "Aisha": "Аиша", "Ananya": "Анания", "Arun": "Арун", "Daniel": "Даниил", "David": "Давид",
    "Emily": "Эмилия", "John": "Иван", "Karen": "Карина", "Kiran": "Кира", "Leo": "Лев",
    "Linda": "Лидия", "Luke": "Лука", "Michael": "Михаил", "Nina": "Нина", "Olivia": "Оливия",
    "Priya": "Полина", "Rohit": "Роман", "Sarah": "Сара", "Sophia": "София", "Vikram": "Виктор",
    # customer-only first names
    "Amelia": "Амелия", "Ava": "Ева", "Charlotte": "Шарлотта", "Emma": "Эмма", "Ethan": "Артём",
    "Harper": "Дарья", "Isabella": "Изабелла", "James": "Яков", "Liam": "Лиам", "Lucas": "Лука",
    "Mason": "Максим", "Mia": "Мия", "Noah": "Ной", "Oliver": "Олег", "William": "Василий",
}
LAST_ATOMS = {
    "Anderson": "Андреев", "Brown": "Борисов", "Davis": "Дмитриев", "Gupta": "Гущин",
    "Iyer": "Игнатьев", "Johnson": "Иванов", "Kumar": "Кузнецов", "Lewis": "Лебедев",
    "Miller": "Морозов", "Patel": "Павлов", "Sharma": "Соколов", "Singh": "Семёнов",
    "Smith": "Смирнов", "Taylor": "Тихонов", "Thomas": "Тимофеев", "Thompson": "Тарасов",
    "Walker": "Волков", "Williams": "Васильев", "Wilson": "Виноградов",
    # customer-only last names
    "Garcia": "Григорьев", "Jones": "Жуков",
}

# --- per-table, per-column realia spec: column -> map name -----------------
# Only columns listed here get russified, and only by exact whole-cell value.
# 'NAME' = compose via FIRST_ATOMS/LAST_ATOMS. Everything else stays English.
MAPS = {
    "DEPARTMENTS": DEPARTMENTS, "LOCATIONS": LOCATIONS, "REGIONS": REGIONS,
    "SEGMENTS": SEGMENTS, "EDUCATION": EDUCATION, "MARITAL": MARITAL, "ROLES": ROLES,
    "SHIP_MODE": SHIP_MODE, "ORDER_STATUS": ORDER_STATUS, "PRODUCT_TAXONOMY": PRODUCT_TAXONOMY,
    "TICKET_REPORTER": TICKET_REPORTER, "ISSUE_TYPE": ISSUE_TYPE, "REPORT_CHANNEL": REPORT_CHANNEL,
    "SHORT_DESCRIPTION": SHORT_DESCRIPTION, "AGENT_NAME": AGENT_NAME,
}
TABLE_REALIA = {
    'HR_ANALYTICS__PUBLIC__DEPARTMENTS': {"DEPARTMENT_NAME": "DEPARTMENTS", "LOCATION": "LOCATIONS"},
    'HR_ANALYTICS__PUBLIC__EMPLOYEES': {
        "EMPLOYEE_NAME": "NAME", "MARITAL_STATUS": "MARITAL", "EDUCATION_LEVEL": "EDUCATION",
        "DEPARTMENT": "DEPARTMENTS", "ROLE": "ROLES",
    },
    # SALARY_HISTORY.CHANGE_REASON -> KEEP (enum); MONTHLY_REVENUE -> all numeric
    'SALES_DW__PUBLIC__CUSTOMERS': {"CUSTOMER_NAME": "NAME", "SEGMENT": "SEGMENTS", "REGION": "REGIONS"},
    'SALES_DW__PUBLIC__ORDERS': {"SHIP_MODE": "SHIP_MODE", "STATUS": "ORDER_STATUS"},
    'SALES_DW__PUBLIC__PRODUCTS': {"CATEGORY": "PRODUCT_TAXONOMY", "SUB_CATEGORY": "PRODUCT_TAXONOMY"},
    #   PRODUCT_NAME -> KEEP (brand-model titles, realistic in RU e-commerce, eval keys on PRODUCT_ID); BRAND -> KEEP
    'SUPPORT_CENTER__PUBLIC__AGENTS': {"AGENT_NAME": "AGENT_NAME"},
    #   TEAM/SKILL_LEVEL -> KEEP (enums)
    'SUPPORT_CENTER__PUBLIC__TICKETS': {
        "REPORTER": "TICKET_REPORTER", "ISSUE_TYPE": "ISSUE_TYPE",
        "REPORT_CHANNEL": "REPORT_CHANNEL", "SHORT_DESCRIPTION": "SHORT_DESCRIPTION",
    },
    #   PRIORITY/STATUS/EVENT/VARIANT -> KEEP (enums)
}

# Flat value->ru map for eval-literal / groundtruth-cell substitution (whole-cell).
# Excludes name atoms (handled compositionally) and KEEP enums by construction.
FLAT_VALUE_MAP = {}
for _m in MAPS.values():
    FLAT_VALUE_MAP.update(_m)


def map_full_name(name):
    """Compose 'First Last' from atoms; return None if no atom matched (leave as-is)."""
    parts = name.split()
    if len(parts) != 2:
        return None
    f, l = parts
    rf, rl = FIRST_ATOMS.get(f), LAST_ATOMS.get(l)
    if rf is None and rl is None:
        return None
    return f"{rf or f} {rl or l}"


def map_value(map_name, value):
    """Map one cell value. map_name 'NAME' composes person names; else dict lookup."""
    v = value.strip()
    if map_name == "NAME":
        return map_full_name(v)
    return MAPS.get(map_name, {}).get(v)
