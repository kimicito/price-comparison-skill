# Price Comparison Skill

[![OpenClaw](https://img.shields.io/badge/OpenClaw-Skill-blue)](https://openclaw.ai)
[![Version](https://img.shields.io/badge/version-5.0-green)](SKILL.md)

Skill для поиска и сравнения цен на материалы и оборудование у российских поставщиков. Работает с Excel-таблицами, ищет по B2B-каталогам и маркетплейсам, записывает 2 цены от разных поставщиков и находит аналоги для снижения стоимости.

## Возможности

- **Автоопределение категории ТМЦ** — загружает шаблон с нужными параметрами (не 30+ строк, а 5-8 критичных)
- **Два типа аналогов** — другая марка + та же марка с отклонениями
- **Inline eval** — проверка в процессе, не пост-фактум
- **Красная команда** — проверка рисков для каждого аналога
- **TCO-анализ** — стоимость владения, не только цена закупки

## Быстрый старт

```bash
# Установка
clone https://github.com/kimicito/price-comparison-skill.git

# Использование
python3 scripts/runner_v2.py input.xlsx results.json
```

## Структура

| Директория | Назначение |
|------------|------------|
| `SKILL.md` | Полная документация skill (v5.0) |
| `templates/` | Шаблоны параметров по категориям (ip_cameras, fiber_optic и т.д.) |
| `scripts/` | runner_v2.py, matrix_builder.py, eval.py |
| `analogs/` | Готовые матрицы сравнения |
| `subagent-*.md` | Конфигурации субагентов |

## Версии

- **v5.2** — Excel Format v2: Вкладка 1 = все цены, Вкладка 2 = аналоги той же марки, Вкладка 3 = аналоги другой марки
- **v5.1** — Cache + Category templates + Inline eval + Dual analogs + Red team
- **v5.0** — Category templates + Inline eval + Dual analogs + Red team
- **v4.1** — Price reduction focus
- **v4.0** — Same-brand deviation analogs
- **v3.0** — Multi-sheet Excel + subagents

## Лицензия

MIT — свободное использование в проектах OpenClaw.
