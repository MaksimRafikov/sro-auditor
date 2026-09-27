#!/usr/bin/env python3
"""Локальный helper СРО-Аудитора: статика чекера + выгрузка НОСТРОЙ по номеру СРО.

Браузер не может сам ходить на reestr.nostroy.ru (CORS), поэтому качает этот
процесс. Реестр договоров остаётся в браузере и на сервер не уходит.

    python tools/sro_server.py            # http://127.0.0.1:8765/sro_checker.html
    python tools/sro_server.py --port 9000 --no-browser

API:
    GET  /api/health                     — чекер понимает, что helper запущен
    GET  /api/nostroy/<sro>              — кэш реестра членов (404, если кэша нет)
    POST /api/nostroy/<sro>/refresh      — запустить выгрузку в фоне
    GET  /api/nostroy/<sro>/progress     — состояние выгрузки
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import webbrowser
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import nostroy_connector as nostroy

ROOT = Path(__file__).resolve().parents[1]
API_VERSION = 1
ALLOWED_ORIGINS = re.compile(r"^(null|https?://(localhost|127\.0\.0\.1)(:\d+)?)$")

_lock = threading.Lock()
_job: dict = {"state": "idle", "sro_id": None, "phase": "", "done": 0, "total": 0, "message": "", "error": None}


def _set_job(**fields: object) -> None:
    with _lock:
        _job.update(fields)


def _job_snapshot() -> dict:
    with _lock:
        return dict(_job)


def start_refresh(sro_id: int, force: bool) -> tuple[bool, str]:
    with _lock:
        if _job["state"] == "running":
            return False, f"уже идёт выгрузка СРО {_job['sro_id']}"
        _job.update(
            {"state": "running", "sro_id": sro_id, "phase": "start", "done": 0, "total": 0,
             "message": "подключение к НОСТРОЙ", "error": None}
        )

    def progress(phase: str, done: int, total: int, message: str) -> None:
        _set_job(phase=phase, done=done, total=total, message=message)

    def worker() -> None:
        try:
            payload = nostroy.refresh(sro_id, progress=progress, force=force)
            _set_job(state="done", phase="done", message="готово", stats=payload["stats"])
        except Exception as exc:
            _set_job(state="error", phase="error", message=str(exc), error=str(exc))

    threading.Thread(target=worker, name=f"nostroy-{sro_id}", daemon=True).start()
    return True, "запущено"


class Handler(SimpleHTTPRequestHandler):
    server_version = "SroAuditorHelper/1"

    def _cors(self) -> None:
        origin = self.headers.get("Origin")
        if origin and ALLOWED_ORIGINS.match(origin):
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")

    def _json(self, status: int, payload: object) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self) -> None:  # статика тоже не должна кэшироваться при правках UI
        if not self.path.startswith("/api/"):
            self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self._cors()
        self.end_headers()

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        match = re.fullmatch(r"/api/nostroy/(\d+)/refresh", path)
        if not match:
            self._json(404, {"error": "неизвестный endpoint"})
            return
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else b""
        try:
            options = json.loads(body) if body else {}
        except ValueError:
            options = {}
        ok, message = start_refresh(int(match.group(1)), bool(options.get("force")))
        self._json(202 if ok else 409, {"ok": ok, "message": message, "job": _job_snapshot()})

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if not path.startswith("/api/"):
            super().do_GET()
            return

        if path == "/api/health":
            self._json(200, {"ok": True, "app": "sro-auditor-helper", "version": API_VERSION})
            return

        progress = re.fullmatch(r"/api/nostroy/(\d+)/progress", path)
        if progress:
            self._json(200, _job_snapshot())
            return

        members = re.fullmatch(r"/api/nostroy/(\d+)", path)
        if members:
            sro_id = int(members.group(1))
            cached = nostroy.load_cached(sro_id)
            if cached is None:
                self._json(404, {"error": f"нет локального кэша по СРО {sro_id}", "sro_id": sro_id})
                return
            self._json(200, cached)
            return

        self._json(404, {"error": "неизвестный endpoint"})

    def log_message(self, fmt: str, *args: object) -> None:
        if self.path.startswith("/api/nostroy") and "progress" in self.path:
            return
        super().log_message(fmt, *args)


_LOOPBACK_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})


def _is_loopback_host(host: str) -> bool:
    h = (host or "").strip().lower()
    if h in _LOOPBACK_HOSTS:
        return True
    # IPv6 loopback in brackets: [::1]
    return h == "[::1]"


def main() -> int:
    parser = argparse.ArgumentParser(description="Локальный helper СРО-Аудитора")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument(
        "--allow-lan",
        action="store_true",
        help="разрешить --host не только localhost (кэш НОСТРОЙ без auth станет доступен в сети)",
    )
    args = parser.parse_args()

    if not _is_loopback_host(args.host) and not args.allow_lan:
        print(
            f"Отказ: --host {args.host!r} открывает кэш реестра (ИНН) в сеть без пароля.\n"
            f"Оставьте 127.0.0.1 или добавьте --allow-lan, если это осознанно.",
            file=sys.stderr,
        )
        return 2

    url = f"http://{args.host}:{args.port}/sro_checker.html"
    server = ThreadingHTTPServer((args.host, args.port), partial(Handler, directory=str(ROOT)))
    print(f"СРО-Аудитор: {url}")
    print(f"Кэш НОСТРОЙ:  {nostroy.CACHE_ROOT}")
    if not _is_loopback_host(args.host):
        print(
            "ВНИМАНИЕ: helper слушает не только localhost — "
            "любой в сети может GET /api/nostroy/<id> и забрать кэш членов."
        )
    print("Ctrl+C — остановить")
    if not args.no_browser:
        threading.Timer(0.6, webbrowser.open, args=(url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nостановлено")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
