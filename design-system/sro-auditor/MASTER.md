# Design System Master File

> **LOGIC:** When building a specific page, first check `design-system/sro-auditor/pages/[page-name].md`.
> If that file exists, its rules **override** this Master file.
> If not, strictly follow the rules below.

---

**Project:** SRO-Auditor  
**Updated:** 2026-07-21  
**Category:** Construction compliance / operator tool  
**Direction:** Industrial audit desk

---

## Product surface

This is a **local audit tool**, not a marketing site.

| Surface | Role |
|---|---|
| `sro_checker.html` | Primary UI: upload Excel → check ВВ/ОДО → risks → CSV |
| Optional intro page | Short handoff only; same visual language, no feature-grid landing |

**Audience:** оператор СРО (повторный ежедневный сценарий).  
**Job:** за 30–60 секунд увидеть критичные отклонения по членам.

---

## Global Rules

### Color Palette

| Role | Hex | CSS Variable |
|------|-----|--------------|
| Primary (slate) | `#64748B` | `--color-primary` |
| Secondary | `#94A3B8` | `--color-secondary` |
| CTA / safety accent | `#F97316` | `--color-cta` |
| CTA hover | `#EA580C` | `--color-cta-hover` |
| Background | `#F1F5F9` | `--color-background` |
| Panel | `#FFFFFF` | `--color-panel` |
| Text | `#334155` | `--color-text` |
| Ink (headings) | `#0F172A` | `--color-ink` |
| Line | `#D8DEE8` | `--color-line` |
| Critical | `#B91C1C` / bg `#FEF2F2` | `--color-crit` |
| Manual review | `#B45309` / bg `#FFFBEB` | `--color-warn` |
| OK | `#047857` / bg `#ECFDF5` | `--color-ok` |

**Color Notes:** Industrial slate + safety orange. Orange reserved for primary action («Проверить») and high-signal accents — not decoration.

### Typography

- **Display / mono:** Fira Code — brand title, ИНН, суммы, коды уровней
- **Body:** Fira Sans — labels, helper text, table cells
- **Mood:** technical, dense, scannable, precise
- **Google Fonts:** [Fira Code + Fira Sans](https://fonts.google.com/share?selection.family=Fira+Code:wght@400;500;600;700|Fira+Sans:wght@300;400;500;600;700)

```css
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600;700&family=Fira+Sans:wght@300;400;500;600;700&display=swap');
```

### Spacing Variables

| Token | Value | Usage |
|-------|-------|-------|
| `--space-xs` | `4px` | Tight gaps, stamp padding |
| `--space-sm` | `8px` | Icon gaps, inline spacing |
| `--space-md` | `16px` | Standard padding |
| `--space-lg` | `24px` | Section padding |
| `--space-xl` | `32px` | Large gaps |
| `--space-2xl` | `48px` | Page top/bottom breathing room |

Avoid oversized hero padding; density beats whitespace.

### Shadow Depths

| Level | Value | Usage |
|-------|-------|-------|
| `--shadow-sm` | `0 1px 2px rgba(15,23,42,0.05)` | Inputs, subtle lift |
| `--shadow-md` | `0 8px 24px rgba(15,23,42,0.08)` | Panels, file dropzones |
| `--shadow-lg` | `0 10px 15px rgba(15,23,42,0.10)` | Sticky summary, dropdowns |

No multi-layer glow. Prefer border (`--color-line`) over heavy shadow.

### Atmosphere

- Cool slate page background
- Blueprint grid (`~28px`) at low opacity
- Soft radial wash of orange/slate in corners (≤10% opacity)
- Risk stamps: **КРИТИЧНО** / **РУЧНАЯ** / **НОРМА** (pill or ink-stamp), never emoji

---

## Style Guidelines

**Style:** Industrial audit desk  
**Keywords:** operator tool, dense table, blueprint, safety orange, stamps, tabular nums, sticky headers  
**Best For:** Local compliance checkers, registry review, Excel→risk workflows  
**Not For:** Fashion portfolios, SaaS marketing landings, dark cyber dashboards

### Page Pattern

**Pattern Name:** Upload → Run → Scan

- **First viewport:** brand mark (secondary) + file slots + «Проверить»
- **After run:** summary strip → tabs/filters → results table → export
- **CTA:** one orange primary only for the check action
- **Brand:** visible, does not overpower the workflow

Preferred composition variants: see `pages/sro_checker.md`.

---

## Component Specs

### Buttons

```css
.btn-primary {
  background: #F97316;
  color: #fff;
  padding: 12px 24px;
  border: none;
  border-radius: 8px;
  font-family: "Fira Sans", system-ui, sans-serif;
  font-weight: 600;
  cursor: pointer;
  transition: background-color 180ms ease, box-shadow 180ms ease;
}
.btn-primary:hover { background: #EA580C; }
.btn-primary:focus-visible { outline: none; box-shadow: 0 0 0 3px rgba(249, 115, 22, 0.35); }

.btn-secondary {
  background: #fff;
  color: #334155;
  border: 1px solid #D8DEE8;
  padding: 12px 24px;
  border-radius: 8px;
  font-weight: 600;
  cursor: pointer;
  transition: border-color 180ms ease, background-color 180ms ease;
}
.btn-secondary:hover { background: #F8FAFC; border-color: #94A3B8; }
```

Do not use `transform: translateY` / `scale` on hover if it shifts layout.

### Panels (not marketing cards)

```css
.panel {
  background: #fff;
  border: 1px solid #D8DEE8;
  border-radius: 12px;
  padding: 20px 24px;
  box-shadow: var(--shadow-sm);
}
```

Use panels for upload zones and result chrome. Avoid nested cards and clickable decorative cards.

### Risk stamps

```css
.stamp {
  display: inline-flex;
  align-items: center;
  font-family: "Fira Code", ui-monospace, monospace;
  font-size: 0.72rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  padding: 4px 8px;
  border-radius: 4px;
  border: 1px solid currentColor;
}
.stamp--crit { color: #B91C1C; background: #FEF2F2; }
.stamp--warn { color: #B45309; background: #FFFBEB; }
.stamp--ok   { color: #047857; background: #ECFDF5; }
```

### Inputs / file zones

```css
.input {
  padding: 12px 16px;
  border: 1px solid #E2E8F0;
  border-radius: 8px;
  font-size: 16px;
  font-family: "Fira Sans", system-ui, sans-serif;
  transition: border-color 180ms ease, box-shadow 180ms ease;
}
.input:focus {
  border-color: #64748B;
  outline: none;
  box-shadow: 0 0 0 3px rgba(100, 116, 139, 0.2);
}

.file-zone {
  border: 1.5px dashed #94A3B8;
  border-radius: 12px;
  background: #F8FAFC;
  padding: 20px;
}
.file-zone[data-ready="true"] {
  border-style: solid;
  border-color: #64748B;
  background: #fff;
}
```

### Tables

- Sticky `thead`
- Numeric columns: `font-variant-numeric: tabular-nums` + Fira Code when dense
- Row hover: soft slate wash, not lift/shadow
- Compact density: prefer readable scan over generous card padding

### Modals (sparingly)

```css
.modal-overlay { background: rgba(15, 23, 42, 0.45); }
.modal {
  background: #fff;
  border: 1px solid #D8DEE8;
  border-radius: 12px;
  padding: 24px;
  box-shadow: var(--shadow-lg);
  max-width: 480px;
  width: 90%;
}
```

Prefer inline status/errors over modal when possible.

---

## Motion

- Stagger fade-up on header / upload / results (~40–60 ms steps)
- Control transitions 150–250 ms
- Respect `prefers-reduced-motion: reduce` (disable stagger and transforms)

---

## Anti-Patterns (Do NOT Use)

- Marketing hero with oversized display type and feature grids
- Purple / pink / indigo gradients; glassmorphism; glow stacks
- Dark “cyber ops” theme for the main checker
- Emoji as icons or risk markers (use SVG + stamps)
- Nested cards / decorative card hover lift
- Layout-shifting hover (`scale`, large `translateY`)
- Low-contrast muted text on slate
- Instant state changes without transition (unless reduced-motion)
- Invisible focus states
- Sending uploaded Excel to an external server in the HTML MVP

---

## Pre-Delivery Checklist

- [ ] First viewport is workflow (upload + run), not a landing pitch
- [ ] Orange used mainly for «Проверить» / critical signal
- [ ] Risk stamps, not emoji
- [ ] Blueprint atmosphere present but quiet
- [ ] Icons from one set (Heroicons / Lucide), SVG only
- [ ] `cursor: pointer` on clickable controls
- [ ] Hover 150–250 ms, no layout shift
- [ ] Text contrast ≥ 4.5:1
- [ ] Visible `:focus-visible`
- [ ] `prefers-reduced-motion` respected
- [ ] Sticky table headers; numeric columns tabular
- [ ] Responsive: 375 / 768 / 1024 / 1440 — no horizontal scroll
- [ ] Files stay in-browser (HTML MVP)
