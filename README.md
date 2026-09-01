# Price Comparison Skill v7.4

[![OpenClaw](https://img.shields.io/badge/OpenClaw-Skill-blue)](https://openclaw.ai)
[![Version](https://img.shields.io/badge/version-7.4-green)](SKILL.md)

Skill для поиска и сравнения цен на материалы и оборудование. Работает с Excel-таблицами, ищет по B2B-каталогам и маркетплейсам, записывает 2 цены + аналоги.

## Формула 2+1+1

| Тип | Описание | Лимит | Fallback |
|-----|----------|-------|----------|
| **Цена 1** | Оригинал, поставщик 1 | 10 мин | FAIL если не найдена |
| **Цена 2** | Оригинал, поставщик 2 | 10 мин | «Не найдена за 10 мин» |
| **Аналог др. марки** | Другой бренд | 5 мин | «—» (прочерк) |
| **Аналог той же марки** | Тот же бренд | 10 мин | «—» (прочерк) |

## Возможности v7.4

- ✅ **Кликабельные цены** — каждая цена = гиперссылка на товар
- ✅ **25 проверок eval** — 11 FAIL + 14 WARN
- ✅ **Матрицы аналогов** — секции сравнения на отдельных вкладках
- ✅ **Проверка URL** — защита от ссылок на главную/404/поиск
- ✅ **Graceful fallback** — "Цена не указана" вместо блокировки

## Быстрый старт

```bash
# Установка
git clone https://github.com/kimicito/price-comparison-skill.git

# Создание Excel с ценами
python3 scripts/runner_v3.py input.xlsx results.json output/

# Проверка готового файла
python3 scripts/eval.py output/price_comparison_main_*.xlsx
```

## Структура

| Файл | Назначение |
|------|------------|
| `SKILL.md` | Полная документация + инструкции для агента |
| `scripts/runner_v3.py` | Генерация Excel с inline eval |
| `scripts/eval.py` | Post-factum проверка (25 проверок) |
| `scripts/matrix_builder_v3.py` | Сборка матриц аналогов |
| `templates/` | Шаблоны параметров по категориям |

## Архитектура

```
[Excel input] → [Поиск цен] → [Inline Eval] → [Excel + URL] → [Eval] → [Итоговый файл]
                                      ↓
                              [Матрицы аналогов]
```

## Версии

- **v7.4** — URL validation + fallback prices + matrix sheets
- **v7.3** — Anti-hallucination eval checks
- **v7.0** — Formula 2+1+1 + Clickable Links + Full Design
- **v6.0** — Tests, retry/fallback, cache v2
- **v5.2** — Excel Format v2: Tab1 prices, Tab2 same-brand, Tab3 cross-brand

## Лицензия

MIT — свободное использование в проектах OpenClaw.
