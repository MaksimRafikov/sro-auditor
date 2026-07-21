# Подключённые skills и тулзы

Junctions в `.cursor/skills/<name>` → источник ниже. Не копируем содержимое ECC — только ссылки.

## ECC-main (`../ECC-main/skills/`)

| Skill | Роль в СРО-Аудитор |
|---|---|
| frontend-design-direction | Направление UI под домен (audit tool, не лендинг) |
| frontend-patterns | Паттерны фронта |
| frontend-a11y | a11y / focus / contrast |
| design-system | Работа с токенами и DS |
| make-interfaces-feel-better | Spacing, hit areas, motion polish |
| dashboard-builder | Операторские вопросы: что сломано / куда смотреть |
| e2e-testing | Playwright E2E паттерны |
| browser-qa | Визуальная/UI проверка |
| click-path-audit | Проверка клик-путей |
| ui-demo | Демо-сценарии UI |
| verification-loop | Повторная верификация отчётов |
| gateguard | Не править логику ВВ/ОДО вслепую |
| orch-build-mvp | Нарезка MVP-срезов |
| prompt-optimizer | Стабильные Instructions для GPT |
| security-scan / safety-guard | Секреты, чувствительные данные |
| documentation-lookup | Актуальные docs (с MCP context7) |
| coding-standards / error-handling | Качество кода |
| product-lens | Продуктовая ясность UX |

## Global Cursor skills (`~/.cursor/skills/`)

| Skill | Роль |
|---|---|
| frontend-design | Базовый skill качественного UI |
| ui-ux-pro-max | Генерация/поиск DS (уже → `design-system/sro-auditor/`) |
| kill-ai-slop | Убрать AI-look |
| deslop | Убрать code slop |
| control-ui | CDP/browser harness |
| playwright | Browser automation |
| browser-qa | UI QA |

## MCP (проектный `.mcp.json`)

- `context7` — документация
- `memory` — память решений по формулам/кейсам
- `sequential-thinking` — разбор спорных сверк

Playwright уже в пользовательском MCP Cursor — секреты не дублируем.

## Design system

Сгенерировано через `ui-ux-pro-max`:

```text
design-system/sro-auditor/MASTER.md
```

Страничные оверрайды: `design-system/sro-auditor/pages/`.
