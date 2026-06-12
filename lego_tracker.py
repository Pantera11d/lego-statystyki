"""
LEGO Market Tracker v5
- Logowanie przez Allegro (refresh token)
- Pobieranie cen ofert z Allegro
- Dane zapisywane do CSV dla Google Sheets
"""

import os
import sys
import json
import base64
import urllib.request
import urllib.parse
import csv
from datetime import date
from pathlib import Path

CSV_FILE = "lego_dane.csv"
EXCEL_FILE = "LEGO_Monitor.xlsx"

CLIENT_ID = os.environ.get("ALLEGRO_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("ALLEGRO_CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("ALLEGRO_REFRESH_TOKEN", "")


def get_token_from_refresh() -> str:
    """Pobiera nowy access token przy użyciu refresh tokena."""
    if not CLIENT_ID or not CLIENT_SECRET:
        print("BLAD: Brak ALLEGRO_CLIENT_ID albo ALLEGRO_CLIENT_SECRET w GitHub Secrets.")
        sys.exit(1)

    creds = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()
    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
    }).encode()

    req = urllib.request.Request(
        "https://allegro.pl/auth/oauth/token",
        data=data,
        headers={
            "Authorization": f"Basic {creds}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )

    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())

        print("Token OK z refresh tokena.")

        new_refresh = result.get("refresh_token", "")
        if new_refresh and new_refresh != REFRESH_TOKEN:
            print(f"NOWY_REFRESH_TOKEN={new_refresh}")
            print("UWAGA: Zaktualizuj ALLEGRO_REFRESH_TOKEN w GitHub Secrets.")

        return result["access_token"]

    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        print(f"BLAD refresh tokena: HTTP {e.code}")
        print(body)
        print("Refresh token mogl wygasnac. Uruchom get_token.py lokalnie i wklej nowy token do GitHub Secrets.")
        sys.exit(1)


def fetch_offers(token: str, set_nr: str) -> list:
    """Pobiera oferty Allegro dla numeru zestawu LEGO."""
    params = urllib.parse.urlencode({
        "phrase": str(set_nr),
        "limit": 100,
        "sort": "price",
    })

    req = urllib.request.Request(
        f"https://api.allegro.pl/offers/listing?{params}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.allegro.public.v1+json",
        },
    )

    try:
        with urllib.request.urlopen(req) as r:
            data = json.loads(r.read())

        items = data.get("items", {})
        regular = items.get("regular", [])
        promoted = items.get("promoted", [])
        return regular + promoted

    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        print(f"BLAD Allegro HTTP {e.code}: {e.reason}")
        print(body[:500])
        return []

    except Exception as e:
        print(f"BLAD pobierania ofert: {e}")
        return []


def parse_stats(items: list) -> dict:
    """Liczy najniższą cenę, liczbę ofert i statystyki popularności."""
    prices = []
    bought_last = 0

    for item in items:
        try:
            amount = item.get("sellingMode", {}).get("price", {}).get("amount")
            if amount not in (None, ""):
                price = float(str(amount).replace(",", "."))
                if price > 0:
                    prices.append(price)
        except (TypeError, ValueError):
            pass

        stats = item.get("stats", {}) or {}
        for key in ("boughtCount", "popularityScore", "watchersCount"):
            value = stats.get(key)
            if value not in (None, ""):
                try:
                    bought_last = max(bought_last, int(float(value)))
                except (TypeError, ValueError):
                    pass

    return {
        "min_price": round(min(prices), 2) if prices else "",
        "offer_count": len(prices),
        "bought_last": bought_last,
    }


def read_sets() -> list:
    """Czyta aktywne zestawy z arkusza LEGO_Monitor.xlsx, zakładka Lista zestawow."""
    try:
        import openpyxl

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Lista zestawow"]
        sets = []

        for row in ws.iter_rows(min_row=4, values_only=True):
            nr = str(row[0]).strip() if row[0] else ""
            name = str(row[1]).strip() if row[1] else ""
            seria = str(row[2]).strip() if row[2] else ""
            aktyw = str(row[4]).strip().upper() if len(row) > 4 and row[4] else ""

            if nr and nr != "None" and aktyw == "TAK":
                sets.append((nr, name, seria))

        return sets

    except Exception as e:
        print(f"BLAD odczytu zestawow: {e}")
        sys.exit(1)


def save_csv(rows: list) -> None:
    """Zapisuje dane do CSV używanego przez Google Sheets."""
    existing = []

    if Path(CSV_FILE).exists():
        with open(CSV_FILE, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            existing = list(reader)

    today = date.today().strftime("%Y-%m-%d")
    existing = [row for row in existing if row.get("Data") != today]
    all_rows = existing + rows

    fieldnames = [
        "Data",
        "Nr setu",
        "Nazwa zestawu",
        "Seria LEGO",
        "Najnizsza cena PLN",
        "Liczba ofert",
        "Kupilo ostatnio",
    ]

    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"Zapisano {len(rows)} wierszy, lacznie w pliku: {len(all_rows)}.")


def main() -> None:
    today = date.today().strftime("%Y-%m-%d")
    print(f"[{today}] Start.")

    if not REFRESH_TOKEN:
        print("BLAD: Brak ALLEGRO_REFRESH_TOKEN w GitHub Secrets.")
        sys.exit(1)

    token = get_token_from_refresh()
    sets = read_sets()
    print(f"Znaleziono {len(sets)} aktywnych zestawow.")

    rows = []

    for nr, name, seria in sets:
        print(f"-> {nr} {name}...", end=" ", flush=True)

        items = fetch_offers(token, nr)
        print(f"znaleziono {len(items)} ofert", end=" | ", flush=True)

        stats = parse_stats(items)
        print(
            f"cena: {stats['min_price']} | "
            f"ofert: {stats['offer_count']} | "
            f"kupilo/score: {stats['bought_last']}"
        )

        rows.append({
            "Data": today,
            "Nr setu": nr,
            "Nazwa zestawu": name,
            "Seria LEGO": seria,
            "Najnizsza cena PLN": stats["min_price"],
            "Liczba ofert": stats["offer_count"],
            "Kupilo ostatnio": stats["bought_last"],
        })

    save_csv(rows)
    print("Gotowe.")


if __name__ == "__main__":
    main()
