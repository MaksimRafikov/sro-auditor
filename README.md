# СРО оценка рисков ОДО и ВВ

**СРО-Аудитор**: проверка лимитов ВВ/ОДО + комплект Knowledge для Custom GPT.

| | |
|--|--|
| **Онлайн** | https://maksimrafikov.github.io/sro-auditor/ |
| **Репозиторий** | https://github.com/MaksimRafikov/sro-auditor |

Источник логики: [чат «СРО проект»](https://chatgpt.com/share/6a5f60a5-ff88-83eb-9344-1654e17f9f08).

## Быстрый старт

1. Откройте [онлайн-чекер](https://maksimrafikov.github.io/sro-auditor/) или локально `sro_checker.html` (`python -m http.server 8765` → `http://127.0.0.1:8765/sro_checker.html`).
2. Загрузите два Excel/CSV **или** нажмите **Демо-данные**.
3. **Проверить** → сводка / риски / CSV.

Файлы обрабатываются только в браузере — на сервер не уходят.

Логика расчёта: `_sro_logic.js` (подключается из HTML).

## Структура

| Путь | Назначение |
|---|---|
| `sro_checker.html` | UI чекера |
| `_sro_logic.js` | Алгоритм ВВ/ОДО |
| `knowledge/` | GPT Knowledge + Instructions |
| `design-system/sro-auditor/` | Design system (ui-ux-pro-max) |
| `samples/` | Демо CSV |
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
