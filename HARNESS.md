# HARNESS.md — Проект: price-comparison-skill

## Назначение

Автоматический поиск и сравнение цен на материалы и оборудование у российских поставщиков. Работает с Excel-таблицами, ищет по B2B-каталогам и маркетплейсам.

## Архитектура

```
[Excel input] → [Кэш?] → [Поиск цен] → [Inline Eval] → [Создание Excel] → [Subagent аналоги] → [Итоговый Excel]
```

## Компоненты

| Компонент | Файл | Ответственность |
|-----------|------|-----------------|
| Runner | `scripts/runner_v2.py` | Создание Excel, inline eval |
| Inline Eval | `scripts/inline_eval.py` | Проверки во время работы |
| Eval (post-factum) | `scripts/eval.py` | Проверка готового Excel |
| Matrix Builder | `scripts/matrix_builder.py` | Сборка финального Excel с матрицами |
| Subagent Retry | `scripts/subagent_retry.py` | Retry + fallback для субагентов |
| Cache v2 | `scripts/cache_v2.py` | Кэширование с TTL по категориям |
| Analog Researcher | `scripts/analog_researcher.py` | Исследование аналогов |

## Форматы данных

### Input: results.json
```json
[
  {
    "num": 1,
    "name": "Cisco C9200-24P-E",
    "price1": 106714,
    "supplier1": "shop.nag.ru",
    "url1": "https://shop.nag.ru/cisco-c9200",
    "price2": 182112,
    "supplier2": "ediscom.ru",
    "url2": "https://ediscom.ru/cisco-c9200",
    "analog": "Huawei S5735-S24P4X",
    "analog_price": 99194,
    "analog_supplier": "network.msk.ru",
    "analog_url": "https://network.msk.ru/huawei-s5735",
    "comment": "Аналог дешевле на 26%"
  }
]
```

### Output: Excel
- **Вкладка 1:** Сводная таблица (все цены + аналоги)
- **Вкладка 2:** Аналоги той же марки (отклонения)
- **Вкладка 3:** Аналоги другой марки (отклонения)

## Параметры

- **TTL кэша по категориям:**
  - Камеры: 14 дней
  - Коммутаторы: 7 дней
  - Оптика: 14 дней
  - Кабели: 30 дней
- **Retry субагентов:** 3 попытки + fallback
- **Inline eval:** перед созданием Excel

## Запуск

```bash
cd skills/price-comparison

# Создание основного Excel
python3 scripts/runner_v2.py input.xlsx results.json output/

# Тесты
python3 tests/test_runner.py

# Очистка просроченного кэша
python3 -c "from scripts.cache_v2 import PriceCache; c = PriceCache(); c.cleanup_expired()"
```

## State Machine

```
ANALYZE → VALIDATE → EXECUTE → VERIFY → REPORT

ANALYZE:   Изучить input, определить категории
VALIDATE:  Проверить кэш, структуру данных
EXECUTE:   Поиск цен → Inline eval → Создание Excel → Субагенты
VERIFY:    Post-factum eval + sanity checks
REPORT:    Итоговый Excel + статистика
```

## Чеклист перед запуском

- [ ] Файл results.json валидный
- [ ] Кэш актуален (или очищен)
- [ ] API ключи для поиска доступны
- [ ] Subagent конфигурация корректна
- [ ] Папка output/ доступна для записи

## Ограничения

- Только российские поставщики
- URL должны быть проверены вручную
- Аналоги — на усмотрение заказчика
- Кэш не гарантирует актуальность цен

## Тесты

```bash
# Все тесты
python3 tests/test_runner.py

# Только inline eval
python3 -c "from scripts.inline_eval import inline_eval_all; ..."
```
