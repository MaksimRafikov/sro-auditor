# СРО оценка рисков ОДО и ВВ

**СРО-Аудитор**: проверка лимитов ВВ/ОДО + комплект Knowledge для Custom GPT.

| | |
|--|--|
| **Онлайн** | https://maksimrafikov.github.io/sro-auditor/ |
| **Репозиторий** | https://github.com/MaksimRafikov/sro-auditor |

Источник логики: [чат «СРО проект»](https://chatgpt.com/share/6a5f60a5-ff88-83eb-9344-1654e17f9f08).

## Быстрый старт (локально)

```bash
pip install requests
python tools/sro_server.py          # http://127.0.0.1:8765/sro_checker.html
```

1. **Шаг 1** — введите номер СРО, как в адресе реестра (`reestr.nostroy.ru/sro/`**`263`**`/member/list`), и нажмите **Обновить с НОСТРОЙ**. Реестр членов Excel-файлом не нужен.
2. **Шаг 2** — загрузите свою выгрузку договоров (Excel/CSV). Шаблон: `samples/contracts-template.csv`.
3. **Проверить** → сводка / риски / **Скачать Excel**.

Договоры обрабатываются только в браузере и на сервер не уходят. Наружу ходит
только helper — за публичным реестром НОСТРОЙ.

### Что делает helper

`tools/sro_server.py` раздаёт статику чекера и проксирует НОСТРОЙ (браузеру это
запрещает CORS). Реальную работу выполняет `tools/nostroy_connector.py`:
список членов СРО → карточка каждого члена → нормализация в колонки, которые уже
читает `_sro_logic.js` → кэш в `cache/nostroy/<номер>/`.

Из карточек берутся ИНН, наименование, уровни ВВ/ОДО, состояние права, дата
регистрации в реестре СРО, периоды приостановок (вся история решений) и расчёт
обязательств с датой. Формулировки уровней НОСТРОЙ пишет словами («не превышает
девяносто миллионов рублей») — коннектор переводит их в «до 90 млн руб.», чтобы
не менять формулы ВВ/ОДО в чекере.

Повторное нажатие **Обновить с НОСТРОЙ** докачивает только изменившиеся карточки
(сверка по `last_updated_at`). Первая выгрузка СРО на ~1400 членов — несколько минут.

Без UI:

```bash
python tools/nostroy_connector.py 263           # обновить кэш
python tools/nostroy_connector.py 263 --force   # перекачать всё заново
```

### Если НОСТРОЙ недоступен

В шаге 1 раскройте **«НОСТРОЙ недоступен — загрузить реестр членов вручную»** и
выберите Excel/CSV. Подойдёт и `cache/nostroy/<номер>/members.csv` от прошлой
выгрузки. Онлайн-версия на [GitHub Pages](https://maksimrafikov.github.io/sro-auditor/)
к локальному helper'у не подключается — там работает только ручная загрузка и **Демо-данные**.

Excel-отчёт (как в прежней сверке): листы **Сводка**, **Проверка компаний**, **Договоры с расчётом**, **Риски**, **Ручная проверка** — с автофильтром, подсветкой рисков и форматом сумм.

Логика расчёта: `_sro_logic.js` (подключается из HTML).

## Структура

| Путь | Назначение |
|---|---|
| `sro_checker.html` | UI чекера |
| `_sro_logic.js` | Алгоритм ВВ/ОДО + шаг «Реестр членов из НОСТРОЙ» |
| `tools/sro_server.py` | Локальный helper: статика + API НОСТРОЙ |
| `tools/nostroy_connector.py` | Выгрузка и нормализация реестра членов |
| `cache/nostroy/<номер>/` | Кэш выгрузки: `members.json`, `members.csv`, карточки (в git не идёт) |
| `knowledge/` | GPT Knowledge + Instructions |
| `design-system/sro-auditor/` | Design system (ui-ux-pro-max) |
| `samples/` | Демо CSV + шаблон реестра договоров |
| `AGENTS.md` | Инструкции агенту |
| `ECC-SKILLS.md` | Подключённые skills/MCP |
| `.cursor/skills/` | Junctions → ECC-main + Cursor skills |
| `.cursor/rules/sro-auditor.mdc` | Правила проекта |
| `.mcp.json` | context7 / memory / sequential-thinking |

## Skills и тулзы

См. `ECC-SKILLS.md`. Ключевые для UI:

- `frontend-design`, `frontend-design-direction`, `ui-ux-pro-max`
- `make-interfaces-feel-better`, `frontend-a11y`, `kill-ai-slop`
- `control-ui` / `playwright` / `browser-qa` — проверка в браузере

## GPT MVP

1. Instructions ← `knowledge/06_Инструкция_GPT_СРО_Аудитор.md`
2. Knowledge ← `knowledge/01`–`07`
3. Code Interpreter включить; актуальные Excel — только в чат

Подробнее: `knowledge/README_как_использовать.md`.
