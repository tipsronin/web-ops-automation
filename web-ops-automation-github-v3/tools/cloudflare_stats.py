#!/usr/bin/env python3
"""Сводный отчёт Cloudflare по большому списку сайтов."""
from __future__ import annotations

import argparse
import csv
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

TOKEN = os.getenv("CF_API_TOKEN", "").strip()
API = "https://api.cloudflare.com/client/v4"
GRAPHQL = API + "/graphql"


def headers() -> dict[str, str]:
    if not TOKEN:
        raise SystemExit("Не задан CF_API_TOKEN")
    return {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}


def zone_id(domain: str) -> str | None:
    response = requests.get(
        API + "/zones",
        headers=headers(),
        params={"name": domain, "status": "active", "per_page": 1},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    return data["result"][0]["id"] if data.get("success") and data.get("result") else None


def stats(zone: str, hours: int) -> dict[str, int]:
    end = datetime.now(timezone.utc)
    start = end - timedelta(hours=hours)
    query = """
    query($zoneTag:String!,$start:Time!,$end:Time!){
      viewer{zones(filter:{zoneTag:$zoneTag}){
        httpRequests1hGroups(limit:200,filter:{datetime_geq:$start,datetime_leq:$end}){
          sum{requests pageViews} uniq{uniques}
        }
      }}
    }
    """
    variables = {
        "zoneTag": zone,
        "start": start.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "end": end.strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    response = requests.post(GRAPHQL, headers=headers(), json={"query": query, "variables": variables}, timeout=30)
    response.raise_for_status()
    data = response.json()
    zones = data.get("data", {}).get("viewer", {}).get("zones", [])
    groups = zones[0].get("httpRequests1hGroups", []) if zones else []
    return {
        "visitors": sum(group.get("uniq", {}).get("uniques", 0) for group in groups),
        "pageviews": sum(group.get("sum", {}).get("pageViews", 0) for group in groups),
        "requests": sum(group.get("sum", {}).get("requests", 0) for group in groups),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Cloudflare статистика для списка сайтов")
    parser.add_argument("domains", type=Path)
    parser.add_argument("--csv", type=Path, default=Path("cloudflare_results.csv"))
    args = parser.parse_args()

    rows = []
    for raw in args.domains.read_text(encoding="utf-8").splitlines():
        domain = raw.strip().lower()
        if not domain or domain.startswith("#"):
            continue
        try:
            zone = zone_id(domain)
            if not zone:
                print(f"[skip] {domain}: зона не найдена")
                continue
            day = stats(zone, 24)
            week = stats(zone, 168)
            row = {"domain": domain}
            row.update({f"{key}_24h": value for key, value in day.items()})
            row.update({f"{key}_7d": value for key, value in week.items()})
            rows.append(row)
            print(f"[ok] {domain}: {day['visitors']} посетителей за 24ч / {week['visitors']} за 7д")
        except Exception as exc:
            print(f"[error] {domain}: {exc}")

    fields = [
        "domain", "visitors_24h", "pageviews_24h", "requests_24h",
        "visitors_7d", "pageviews_7d", "requests_7d",
    ]
    with args.csv.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Сохранено: {args.csv}")


if __name__ == "__main__":
    main()
