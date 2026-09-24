#!/usr/bin/env python3
"""Помощник для массового подключения доменов в cPanel без хранения паролей в коде."""
from __future__ import annotations

import argparse
import os
import re
import time
from pathlib import Path

import requests
from requests.auth import HTTPBasicAuth

HOST = os.getenv("CPANEL_HOST", "").strip()
PORT = int(os.getenv("CPANEL_PORT", "2083"))
USER = os.getenv("CPANEL_USER", "").strip()
PASSWORD = os.getenv("CPANEL_PASSWORD", "").strip()
MAIN_DOMAIN = os.getenv("CPANEL_MAIN_DOMAIN", "").strip().lower()
VERIFY_SSL = os.getenv("CPANEL_VERIFY_SSL", "1") == "1"


def require_config() -> None:
    missing = [name for name, value in (
        ("CPANEL_HOST", HOST), ("CPANEL_USER", USER), ("CPANEL_PASSWORD", PASSWORD)
    ) if not value]
    if missing:
        raise SystemExit("Не заданы переменные: " + ", ".join(missing))


def api_call(module: str, func: str, extra: dict | None = None) -> requests.Response:
    require_config()
    params = {
        "cpanel_jsonapi_apiversion": "2",
        "cpanel_jsonapi_module": module,
        "cpanel_jsonapi_func": func,
    }
    params.update(extra or {})
    return requests.get(
        f"https://{HOST}:{PORT}/json-api/cpanel",
        params=params,
        auth=HTTPBasicAuth(USER, PASSWORD),
        timeout=60,
        verify=VERIFY_SSL,
    )


def normalize(domain: str) -> str:
    value = domain.strip().lower().replace("https://", "").replace("http://", "").strip("/")
    try:
        value = value.encode("idna").decode("ascii")
    except Exception:
        pass
    return value


def valid(domain: str) -> bool:
    return re.match(r"^(?!-)([a-z0-9-]{1,63}\.)+[a-z0-9-]{2,63}$", domain) is not None


def list_domains() -> list[str]:
    response = api_call("AddonDomain", "listaddondomains")
    response.raise_for_status()
    data = response.json().get("cpanelresult", {}).get("data", [])
    found = set()
    for item in data:
        value = item.get("domain") or item.get("domainkey") or item.get("rootdomain")
        if value:
            found.add(str(value).lower())
    return sorted(found)


def add_domain(domain: str, apply: bool = False) -> tuple[str, str, str]:
    domain = normalize(domain)
    if not valid(domain):
        return domain, "skip", "некорректный домен"
    if MAIN_DOMAIN and domain == MAIN_DOMAIN:
        return domain, "skip", "основной домен аккаунта"

    subdomain = re.sub(r"-+", "-", re.sub(r"[^a-z0-9-]", "-", domain.replace(".", "-"))).strip("-")[:55]
    docroot = f"public_html/{domain}"

    if not apply:
        return domain, "dry-run", docroot

    response = api_call("AddonDomain", "addaddondomain", {
        "newdomain": domain,
        "subdomain": subdomain,
        "dir": docroot,
        "ftp_is_optional": "1",
    })
    if response.status_code in (401, 403):
        return domain, "error", "ошибка авторизации"
    response.raise_for_status()
    data = response.json().get("cpanelresult", {})
    event = data.get("event", {})
    ok = str(event.get("result")) == "1"
    message = "добавлен" if ok else (event.get("reason") or data.get("error") or "неизвестная ошибка")
    return domain, "ok" if ok else "error", message


def main() -> None:
    parser = argparse.ArgumentParser(description="Работа с addon-доменами cPanel")
    parser.add_argument("--list", action="store_true", help="показать текущие домены")
    parser.add_argument("--add-file", type=Path, help="файл со списком доменов")
    parser.add_argument("--apply", action="store_true", help="реально добавить домены; без флага только проверка")
    args = parser.parse_args()

    if args.list or not args.add_file:
        for index, domain in enumerate(list_domains(), start=1):
            print(f"{index}. {domain}")

    if args.add_file:
        existing = set(list_domains())
        for raw in args.add_file.read_text(encoding="utf-8").splitlines():
            domain = normalize(raw)
            if not domain or domain.startswith("#"):
                continue
            if domain in existing:
                print(f"[skip] {domain}: уже есть")
                continue
            domain, status, message = add_domain(domain, args.apply)
            print(f"[{status}] {domain}: {message}")
            if status == "ok":
                existing.add(domain)
            time.sleep(0.5)


if __name__ == "__main__":
    main()
