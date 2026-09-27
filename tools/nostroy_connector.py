#!/usr/bin/env python3
"""Локальный коннектор НОСТРОЙ: номер СРО → кэш на диске → строки для analyze().

Качает реестр членов строительной СРО с reestr.nostroy.ru, складывает сырые
карточки в cache/nostroy/<sro_id>/cards/ и собирает members.json — массив строк
с теми же названиями колонок, которые уже распознаёт mapMembers() в _sro_logic.js.

    python tools/nostroy_connector.py 263
    python tools/nostroy_connector.py 263 --force
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Iterable

import requests

BASE = "https://reestr.nostroy.ru"
CACHE_ROOT = Path(__file__).resolve().parents[1] / "cache" / "nostroy"
PAGE_COUNT = 100
WORKERS = 8
RETRIES = 4
TIMEOUT = 60

ProgressFn = Callable[[str, int, int, str], None]


def _noop(phase: str, done: int, total: int, message: str) -> None:
    pass


# --- HTTP -------------------------------------------------------------------


def make_session(sro_id: int) -> requests.Session:
    sess = requests.Session()
    sess.headers.update(
        {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Origin": BASE,
            "Referer": f"{BASE}/sro/{sro_id}/member/list",
            "User-Agent": "Mozilla/5.0 (compatible; SRO-Auditor/1.0; +local export)",
        }
    )
    return sess


def post_json(sess: requests.Session, path: str, payload: dict) -> dict:
    url = f"{BASE}{path}"
    last_err: Exception | None = None
    for attempt in range(1, RETRIES + 1):
        try:
            resp = sess.post(url, json=payload, timeout=TIMEOUT)
            if resp.status_code >= 500:
                raise requests.HTTPError(f"{resp.status_code} for {url}", response=resp)
            resp.raise_for_status()
            return resp.json()
        except Exception as exc:  # retry any transient failure
            last_err = exc
            time.sleep(0.4 * attempt)
    raise RuntimeError(f"НОСТРОЙ не ответил на {path}: {last_err}")


def fetch_member_list(sess: requests.Session, sro_id: int, progress: ProgressFn) -> list[dict]:
    def page(n: int) -> dict:
        body = post_json(
            sess,
            f"/api/sro/{sro_id}/member/list",
            {"filters": {}, "page": n, "pageCount": str(PAGE_COUNT), "sortBy": {}},
        )
        return body["data"]

    first = page(1)
    members = list(first["data"])
    pages = int(first["countPages"])
    total = int(first["count"])
    if not members:
        raise RuntimeError(f"СРО {sro_id}: реестр членов пуст или номер не найден")
    progress("list", len(members), total, f"список членов: {len(members)}/{total}")
    for n in range(2, pages + 1):
        block = page(n)
        members.extend(block["data"])
        progress("list", len(members), total, f"список членов: {len(members)}/{total}")
        time.sleep(0.1)
    return members


def fetch_card(sess: requests.Session, member_id: int) -> dict:
    body = post_json(sess, f"/api/member/{member_id}/info", {})
    data = body.get("data")
    if not data:
        raise RuntimeError(body.get("message") or "пустая карточка")
    return data


# --- Нормализация -----------------------------------------------------------

_SCALES = (
    ("миллиард", 1_000_000_000, "млрд"),
    ("миллион", 1_000_000, "млн"),
    ("тысяч", 1_000, "тыс"),
)

_WORD_NUMBERS = {
    "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4, "пять": 5,
    "шесть": 6, "семь": 7, "восемь": 8, "девять": 9, "десять": 10,
    "одиннадцать": 11, "двенадцать": 12, "пятнадцать": 15, "двадцать": 20,
    "тридцать": 30, "сорок": 40, "пятьдесят": 50, "шестьдесят": 60,
    "семьдесят": 70, "восемьдесят": 80, "девяносто": 90, "сто": 100,
    "двести": 200, "триста": 300, "четыреста": 400, "пятьсот": 500,
    "шестьсот": 600, "семьсот": 700, "восемьсот": 800, "девятьсот": 900,
}


def parse_level_cost(text: object) -> tuple[str, float | None, bool]:
    """Формулировка уровня НОСТРОЙ → (текст для чекера, лимит ₽, без верхнего предела).

    НОСТРОЙ пишет лимит словами («не превышает девяносто миллионов рублей»),
    а parseLimit() в _sro_logic.js читает «до 90 млн руб.» — переводим здесь,
    чтобы не трогать формулы ВВ/ОДО.
    """
    raw = "" if text is None else str(text).strip()
    if not raw:
        return "", None, False

    s = raw.lower().replace("ё", "е")
    unlimited = bool(re.search(r"и\s*более|свыше|не\s*менее", s))

    scale_value = 1
    scale_short = ""
    for stem, value, short in _SCALES:
        if stem in s:
            scale_value, scale_short = value, short
            break

    digits = re.search(r"\d[\d\s.,]*", s)
    if digits:
        cleaned = digits.group(0).replace(" ", "").replace("\u00a0", "").replace(",", ".")
        try:
            amount = float(cleaned.rstrip("."))
        except ValueError:
            amount = 0.0
    else:
        amount = float(sum(v for w, v in _WORD_NUMBERS.items() if re.search(rf"\b{w}\b", s)))

    if not amount:
        return raw, None, unlimited

    if amount == int(amount):
        amount_text = str(int(amount))
    else:
        amount_text = f"{amount:g}".replace(".", ",")
    tail = f"{amount_text} {scale_short} руб." if scale_short else f"{amount_text} руб."

    if unlimited:
        return f"свыше {tail}", None, True
    return f"до {tail}", amount * scale_value, False


def title(value: object) -> str:
    if isinstance(value, dict):
        return str(value.get("title") or "")
    if value is None:
        return ""
    return str(value)


def fmt_date(value: object) -> str:
    """Дата НОСТРОЙ (ISO или «дд.мм.гггг чч:мм:сс») → дд.мм.гггг."""
    if not value:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    if re.match(r"^\d{2}\.\d{2}\.\d{4}", s):
        return s[:10]
    if "T" in s:
        try:
            return datetime.fromisoformat(s).strftime("%d.%m.%Y")
        except ValueError:
            return s[:10]
    return s.split(" ")[0]


def parse_money(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(" ", "").replace("\u00a0", "")
    if not s:
        return None
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _iso_key(value: object) -> str:
    return str(value or "")


def collect_periods(right: dict) -> tuple[list[str], list[str], list[str]]:
    """История решений по праву → периоды приостановок, ограничений ОДО, лог решений."""
    actions = sorted(right.get("actions") or [], key=lambda a: _iso_key(a.get("basis_date")))
    suspensions: list[str] = []
    restrictions: list[str] = []
    log: list[str] = []
    open_susp: str | None = None
    open_restr: str | None = None

    for action in actions:
        kind = title(action.get("suspension_decision"))
        date = fmt_date(action.get("basis_date"))
        log.append(f"{date} | {kind} | {action.get('basis') or ''}")
        k = kind.lower()
        if "приостановлен" in k:
            open_susp = date
        elif "возобновлен" in k and open_susp:
            suspensions.append(f"{open_susp} — {date}")
            open_susp = None
        elif "прекращен" in k and open_susp:
            suspensions.append(f"{open_susp} — {date} (до прекращения)")
            open_susp = None
        elif k.startswith("ограничение права"):
            open_restr = date
        elif k.startswith("снятие ограничения") and open_restr:
            restrictions.append(f"{open_restr} — {date}")
            open_restr = None

    if open_susp:
        suspensions.append(f"{open_susp} — открыто / без возобновления в реестре")
    if open_restr:
        restrictions.append(f"{open_restr} — открыто")
    return suspensions, restrictions, log


def normalize_card(card: dict) -> dict:
    """Карточка НОСТРОЙ → строка в формате, который читает mapMembers()."""
    right = card.get("right") or {}
    vv_text, vv_limit, vv_inf = parse_level_cost((card.get("responsibility_level_vv") or {}).get("cost"))
    odo_text, odo_limit, odo_inf = parse_level_cost((card.get("responsibility_level_odo") or {}).get("cost"))
    suspensions, restrictions, log = collect_periods(right)
    member_id = card.get("id")

    return {
        # Колонки, которые читает analyze() — названия менять нельзя.
        "ИНН": str(card.get("inn") or "").strip(),
        "Наименование": (card.get("short_description") or card.get("full_description") or "").strip(),
        "Состояние права": title(right.get("right_status")),
        "Уровень ВВ": vv_text,
        "Уровень ОДО": odo_text,
        "Расчёт обязательств": parse_money(card.get("members_total_liability")),
        "Дата расчёта размера обязательств": fmt_date(
            card.get("members_total_liability_last_count_dt_date_string")
        ),
        "Дата регистрации в реестре СРО": fmt_date(
            card.get("registry_registration_date_time_string") or card.get("registry_registration_date")
        ),
        "Периоды приостановок": "; ".join(suspensions),
        # Справочные колонки: analyze() их не использует, они для оператора и отчёта.
        "Полное наименование": (card.get("full_description") or "").strip(),
        "Реестровый номер": str(card.get("registration_number") or ""),
        "Тип члена": title(card.get("member_type")),
        "Регион": title(card.get("region_number")),
        "Формулировка ВВ (НОСТРОЙ)": title(card.get("responsibility_level_vv")),
        "Формулировка ОДО (НОСТРОЙ)": title(card.get("responsibility_level_odo")),
        "Предел ВВ, ₽": "без верхнего предела" if vv_inf else vv_limit,
        "Предел ОДО, ₽": "без верхнего предела" if odo_inf else odo_limit,
        "Взнос КФ ВВ": parse_money(card.get("compensation_fund_fee_vv")),
        "Взнос КФ ОДО": parse_money(card.get("compensation_fund_fee_odo")),
        "Соответствие требованиям": title(card.get("accordance_status")),
        "Дата прекращения членства": fmt_date(card.get("suspension_date")),
        "Основание прекращения": card.get("suspension_reason") or "",
        "Периоды ограничения по конкурентным": "; ".join(restrictions),
        "Решений по праву": len(right.get("actions") or []),
        "История решений": " || ".join(log),
        "Обновлено в НОСТРОЙ": fmt_date(card.get("last_updated_at_date_time_string") or card.get("last_updated_at")),
        "Ссылка НОСТРОЙ": f"{BASE}/member/{member_id}" if member_id else "",
        "Записей по ИНН": 1,
    }


_RIGHT_RANK = {"действует": 0, "приостановлено": 1, "прекращено": 2}


def _right_rank(row: dict) -> int:
    return _RIGHT_RANK.get(row["Состояние права"].strip().lower(), 3)


def _reg_key(row: dict) -> str:
    d = row["Дата регистрации в реестре СРО"]
    return f"{d[6:10]}{d[3:5]}{d[0:2]}" if len(d) == 10 else ""


def _merge_periods(values: Iterable[str]) -> str:
    seen: list[str] = []
    for value in values:
        for chunk in (p.strip() for p in value.split(";")):
            if chunk and chunk not in seen:
                seen.append(chunk)
    return "; ".join(seen)


def dedupe_by_inn(rows: list[dict]) -> list[dict]:
    """Одна строка на ИНН: приоритет действующему праву, затем свежей регистрации.

    В реестре бывает несколько записей на один ИНН (вышел — вступил заново).
    Периоды приостановок и ограничений склеиваем со всех записей: договор,
    попавший в любое историческое окно, должен остаться видимым.
    """
    groups: dict[str, list[dict]] = {}
    for row in rows:
        inn = row["ИНН"]
        if not inn:
            continue
        groups.setdefault(inn, []).append(row)

    out: list[dict] = []
    for group in groups.values():
        group.sort(key=_reg_key, reverse=True)
        group.sort(key=_right_rank)
        primary = dict(group[0])
        if len(group) > 1:
            primary["Записей по ИНН"] = len(group)
            primary["Периоды приостановок"] = _merge_periods(r["Периоды приостановок"] for r in group)
            primary["Периоды ограничения по конкурентным"] = _merge_periods(
                r["Периоды ограничения по конкурентным"] for r in group
            )
            primary["Ссылка НОСТРОЙ"] = "; ".join(r["Ссылка НОСТРОЙ"] for r in group if r["Ссылка НОСТРОЙ"])
        out.append(primary)

    out.sort(key=lambda r: r["Наименование"] or r["ИНН"])
    return out


# --- Кэш --------------------------------------------------------------------


def cache_dir(sro_id: int) -> Path:
    return CACHE_ROOT / str(sro_id)


def extract_sro_names(card: dict) -> tuple[str, str]:
    """Из карточки члена: (полное название СРО, короткое)."""
    sro = card.get("sro") or {}
    if not isinstance(sro, dict):
        return "", ""
    full = re.sub(r"\s+", " ", (sro.get("full_description") or "").strip())
    short = re.sub(r"\s+", " ", (sro.get("short_description") or "").strip())
    return full, short


def peek_sro_names(sro_id: int) -> tuple[str, str]:
    """Подтянуть название СРО из любой сохранённой карточки (для старого кэша)."""
    cards = cache_dir(sro_id) / "cards"
    if not cards.is_dir():
        return "", ""
    for path in cards.glob("*.json"):
        try:
            card = json.loads(path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            continue
        full, short = extract_sro_names(card)
        if full or short:
            return full, short
    return "", ""


def load_cached(sro_id: int) -> dict | None:
    path = cache_dir(sro_id) / "members.json"
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return None
    if not isinstance(data, dict):
        return None
    # Старые выгрузки без названия — один раз добираем из карточки.
    if not data.get("sro_name") and not data.get("sro_short_name"):
        full, short = peek_sro_names(sro_id)
        if full or short:
            data["sro_name"] = full
            data["sro_short_name"] = short
    return data


def _card_path(sro_id: int, member_id: int) -> Path:
    return cache_dir(sro_id) / "cards" / f"{member_id}.json"


def _load_stamps(sro_id: int) -> dict[str, str]:
    path = cache_dir(sro_id) / "card_stamps.json"
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return {}


def _write_json(path: Path, data: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


def _write_csv(path: Path, rows: list[dict]) -> None:
    """Тот же реестр в CSV — открыть в Excel и загрузить вручную, если НОСТРОЙ недоступен."""
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


# --- Основной проход --------------------------------------------------------


def refresh(sro_id: int, progress: ProgressFn = _noop, force: bool = False) -> dict:
    """Скачать реестр СРО, обновить кэш и вернуть payload для чекера."""
    sess = make_session(sro_id)
    progress("list", 0, 0, "запрос списка членов")
    members = fetch_member_list(sess, sro_id, progress)

    stamps = {} if force else _load_stamps(sro_id)
    fresh_stamps: dict[str, str] = {}
    to_fetch: list[int] = []
    cards: list[dict] = []

    for member in members:
        mid = member.get("id")
        if mid is None:
            continue
        stamp = str(member.get("last_updated_at_date_time_string") or "")
        fresh_stamps[str(mid)] = stamp
        path = _card_path(sro_id, mid)
        if not force and path.exists() and stamps.get(str(mid)) == stamp:
            try:
                cards.append(json.loads(path.read_text(encoding="utf-8")))
                continue
            except (ValueError, OSError):
                pass
        to_fetch.append(mid)

    total = len(to_fetch)
    reused = len(cards)
    progress("cards", 0, total, f"карточки: 0/{total} (из кэша {reused})")

    errors: list[dict] = []
    if total:
        def job(mid: int) -> tuple[int, dict | None, str | None]:
            try:
                return mid, fetch_card(sess, mid), None
            except Exception as exc:
                return mid, None, str(exc)

        done = 0
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            futures = [pool.submit(job, mid) for mid in to_fetch]
            for fut in as_completed(futures):
                mid, card, err = fut.result()
                done += 1
                if card is None:
                    errors.append({"id": mid, "error": err})
                    fresh_stamps.pop(str(mid), None)
                else:
                    _write_json(_card_path(sro_id, mid), card)
                    cards.append(card)
                if done % 25 == 0 or done == total:
                    progress("cards", done, total, f"карточки: {done}/{total} (из кэша {reused})")

    if not cards:
        raise RuntimeError(f"СРО {sro_id}: не удалось получить ни одной карточки члена")

    progress("normalize", 0, 0, "нормализация")
    sro_name, sro_short = "", ""
    for card in cards:
        sro_name, sro_short = extract_sro_names(card)
        if sro_name or sro_short:
            break
    rows = dedupe_by_inn([normalize_card(c) for c in cards])
    payload = build_payload(
        sro_id,
        rows,
        cards_failed=len(errors),
        reused=reused,
        sro_name=sro_name,
        sro_short_name=sro_short,
    )

    _write_json(cache_dir(sro_id) / "card_stamps.json", fresh_stamps)
    _write_json(cache_dir(sro_id) / "members.json", payload)
    _write_csv(cache_dir(sro_id) / "members.csv", rows)
    if errors:
        _write_json(cache_dir(sro_id) / "errors.json", errors)
    progress("done", 1, 1, "готово")
    return payload


def build_payload(
    sro_id: int,
    rows: list[dict],
    cards_failed: int = 0,
    reused: int = 0,
    sro_name: str = "",
    sro_short_name: str = "",
) -> dict:
    with_susp = sum(1 for r in rows if r["Периоды приостановок"])
    return {
        "sro_id": sro_id,
        "sro_name": sro_name,
        "sro_short_name": sro_short_name,
        "source": f"{BASE}/sro/{sro_id}/member/list",
        "exported_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "stats": {
            "members": len(rows),
            "with_suspensions": with_susp,
            "right_active": sum(1 for r in rows if _right_rank(r) == 0),
            "right_suspended": sum(1 for r in rows if _right_rank(r) == 1),
            "right_terminated": sum(1 for r in rows if _right_rank(r) == 2),
            "no_vv_level": sum(1 for r in rows if not r["Уровень ВВ"]),
            "no_odo_level": sum(1 for r in rows if not r["Уровень ОДО"]),
            "duplicate_inn": sum(1 for r in rows if r["Записей по ИНН"] > 1),
            "unparsed_levels": sum(
                1 for r in rows if r["Уровень ВВ"] and r["Предел ВВ, ₽"] is None
            ),
            "cards_failed": cards_failed,
            "cards_from_cache": reused,
        },
        "members": rows,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Выгрузка реестра членов строительной СРО с НОСТРОЙ")
    parser.add_argument("sro_id", type=int, help="номер СРО как в reestr.nostroy.ru/sro/<номер>, например 263")
    parser.add_argument("--force", action="store_true", help="перекачать все карточки, игнорируя кэш")
    args = parser.parse_args()

    def show(phase: str, done: int, total: int, message: str) -> None:
        print(f"  [{phase}] {message}", flush=True)

    payload = refresh(args.sro_id, progress=show, force=args.force)
    print(json.dumps(payload["stats"], ensure_ascii=False, indent=2))
    print(f"Кэш: {cache_dir(args.sro_id) / 'members.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
