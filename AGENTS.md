# СРО-Аудитор — Agent Guide

## Mission

Локальный инструмент проверки членов строительной СРО: лимиты **ВВ** (один договор) и **ОДО** (совокупные конкурентные обязательства) + комплект Knowledge для Custom GPT.

Спека логики: `knowledge/`, диалог-источник: `chatgpt-dialog.md`.

## Product surface

| Артефакт | Роль |
|---|---|
| `sro_checker.html` | MVP UI: загрузка 2 Excel → сверка → риски → CSV |
| `knowledge/*` | GPT Instructions + Knowledge |
| `design-system/sro-auditor/MASTER.md` | Визуальный source of truth |
| `samples/*` | Демо CSV |

## Skills wired

Junctions: `.cursor/skills/<name>` → `ECC-main/skills/*` или `~/.cursor/skills/*`.  
Список и приоритет: `ECC-SKILLS.md`.

### Приоритет для агента

1. `safety-guard` / `security-scan` — чувствительные ИНН/договоры; данные не уходят на сервер в HTML-MVP
2. `frontend-design` + `frontend-design-direction` + `ui-ux-pro-max` — UI чекера
3. `make-interfaces-feel-better` / `frontend-a11y` / `kill-ai-slop` — полировка и анти-slop
4. `gateguard` / `verification-loop` — не менять формулы ВВ/ОДО без явного запроса
5. `control-ui` / `playwright` / `browser-qa` / `e2e-testing` — проверка UI
6. `prompt-optimizer` — правки GPT Instructions
7. `orch-build-mvp` — нарезка следующих срезов (HTML → сайт)

## Design direction

**Industrial audit desk** (не SaaS-лендинг):
- тон: технический, плотный, сканируемый оператором;
- палитра: industrial slate + safety orange (`design-system/sro-auditor/MASTER.md`);
- шрифты: Fira Sans + Fira Code;
- запоминаемая деталь: blueprint-сетка + «штампы» рисков;
- motion: staggered reveal + hover 150–250ms; уважать `prefers-reduced-motion`.

## Do not

- Не отправлять загруженные Excel на внешний сервер в HTML-MVP.
- Не зашивать актуальные реестры в GPT Knowledge.
- Не менять пороги ВВ/ОДО из ГрК без явного запроса.
- Не деплоить live без review через `security-scan`.
