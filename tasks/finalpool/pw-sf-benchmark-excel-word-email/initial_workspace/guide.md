## Руководство по сравнительному анализу (Benchmark Review Guide)

### Источники данных
- **Внешние эталоны (External Benchmarks)**: Industry Analytics Dashboard (внутренний портал) — доступ через веб-браузер
- **Внутренние метрики (Internal Metrics)**: корпоративное хранилище данных (ClickHouse)

### Методология сопоставления
1. Извлеките эталонные значения с дашборда для каждой метрики
2. Запросите внутреннее хранилище данных для соответствующих внутренних значений
3. Рассчитайте отклонения: Internal Value - Industry Average
4. Рассчитайте процент отклонения: (Internal - Industry) / Industry * 100
5. Классифицируйте отклонения по правилам из Benchmark_Context.pdf
6. Присвойте приоритеты на основе классификации

### Сопоставление метрик
| Dashboard Metric | Internal Calculation |
|---|---|
| Avg Salary | AVG of all employee salaries |
| Employee Satisfaction | AVG of job satisfaction scores |
| Revenue Per Employee | Total order revenue / Total employees |
| Avg Order Value | AVG of order total amounts |
| Customer Retention Rate | N/A - not directly calculable |
| SLA Compliance Rate | N/A - not directly calculable |

### Примечания
- Округляйте денежные значения до 2 знаков после запятой
- Округляйте проценты до 1 знака после запятой
- Сортируйте все выходные таблицы по названию метрики в алфавитном порядке, если не указано иное
