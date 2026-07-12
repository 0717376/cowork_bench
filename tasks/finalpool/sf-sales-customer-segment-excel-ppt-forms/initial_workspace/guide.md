# Методология анализа сегментов

## Ключевые метрики (Core Metrics)
- Customer Count: COUNT(DISTINCT CUSTOMER_ID) по сегменту
- Total Orders: COUNT(ORDER_ID) по сегменту
- Total Revenue: SUM(TOTAL_AMOUNT), округлить до 2 знаков
- Avg Order Value: AVG(TOTAL_AMOUNT), округлить до 2 знаков
- Orders Per Customer: Total Orders / Customer Count, округлить до 2 знаков
- Avg Discount Pct: AVG(DISCOUNT) * 100, округлить до 2 знаков
- Revenue Share: выручка сегмента / общая выручка * 100, округлить до 1 знака

## Индекс прибыльности (Profitability Index)
Формула: AVG((UNIT_PRICE - UNIT_COST) / UNIT_PRICE * 100)
Округлить до 1 знака. Требует JOIN заказов с таблицей products.

## Индикаторы роста (Growth Indicators)
1. Найдите середину диапазона дат заказов (MIN..MAX)
2. Подсчитайте заказы в каждой половине для каждого сегмента
3. High Growth: > 50% заказов попадают в недавнюю половину
4. Low Growth: <= 50% заказов в недавней половине

## Стратегические категории (Strategic Categories)
По методологии матрицы BCG:
- Сравните вклад каждого сегмента в выручку с медианой по сегментам
- Star: выручка выше медианы + high growth
- Cash Cow: выручка выше медианы + low growth
- Question Mark: выручка ниже медианы + high growth
- Dog: выручка ниже медианы + low growth
