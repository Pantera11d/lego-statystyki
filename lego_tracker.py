"""
LEGO Market Tracker v4
- Logowanie przez Allegro (authorization_code + refresh token)
- Token wazny 3 miesiace bez ponownego logowania
- Dane zapisywane do CSV dla Google Sheets
"""

import os, sys, json, base64, urllib.request, urllib.parse, csv
from datetime import date
from pathlib import Path

CSV_FILE   = "lego_dane.csv"
EXCEL_FILE = "LEGO_Monitor.xlsx"

CLIENT_ID     = os.environ.get("ALLEGRO_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("ALLEGRO_CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("ALLEGRO_REFRESH_TOKEN", "")
REDIRECT_URI  = "https://pantera11d.github.io/lego-statystyki/callback"

def get_token_from_refresh() -> str:
    """Pobiera nowy access token uzywajac refresh tokena."""
    creds = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()
    data  = urllib.parse.urlencode({
        "grant_type":    "refresh_token",
        "refresh_token": REFRESH_TOKEN,
    }).encode()
    req = urllib.request.Request(
        "https://allegro.pl/auth/oauth/token", data=data,
        headers={"Authorization": f"Basic {creds}",
                 "Content-Type":  "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            print("Token OK (z refresh tokena).")
            # Wypisz nowy refresh token do logów
            new_refresh = result.get("refresh_token", "")
            if new_refresh and new_refresh != REFRESH_TOKEN:
                print(f"NOWY_REFRESH_TOKEN={new_refresh}")
                print("UWAGA: Zaktualizuj ALLEGRO_REFRESH_TOKEN w GitHub Secrets!")
            return result["access_token"]
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"BLAD refresh tokena: {body}")
        print("Refresh token wygasl. Uruchom get_token.py lokalnie zeby uzyskac nowy.")
        sys.exit(1)

def fetch_offers(token: str, set_nr: str) -> list:
    params = urllib.parse.urlencode({
        "phrase":   f"LEGO {set_nr}",
        "limit":    60,
        "sort":     "price",
        "fallback": "false",
    })
    req = urllib.request.Request(
        f"https://api.allegro.pl/offers/listing?{params}",
        headers={"Authorization": f"Bearer {token}",
                 "Accept":        "application/vnd.allegro.public.v1+json"})
    try:
        with urllib.request.urlopen(req) as r:
            data = json.loads(r.read())
        return (data.get("items", {}).get("regular", []) +
                data.get("items", {}).get("promoted", []))
    except urllib.error.HTTPError as e:
        print(f"  BLAD {e.code}: {e.reason}")
        return []
    except Exception as e:
        print(f"  BLAD: {e}")
        return []

def parse_stats(items: list, set_nr: str) -> dict:
    prices      = []
    bought_last = 0
    for it in items:
        name = it.get("name", "").lower()
        if set_nr.lower() not in name:
            continue
        try:
            p = it.get("sellingMode", {}).get("price", {}).get("amount")
            if p:
                prices.append(float(p))
        except (TypeError, ValueError):
            pass
        stats = it.get("stats", {})
        for key in ["boughtCount", "popularityScore", "watchersCount"]:
            val = stats.get(key)
            if val:
                try:
                    bought_last = max(bought_last, int(val))
                except (TypeError, ValueError):
                    pass
    return {
        "min_price":   round(min(prices), 2) if prices else "",
        "offer_count": len(prices),
        "bought_last": bought_last,
    }

def read_sets() -> list:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Lista zestawow"]
        sets = []
        for row in ws.iter_rows(min_row=4, values_only=True):
            nr    = str(row[0]).strip() if row[0] else ""
            name  = str(row[1]).strip() if row[1] else ""
            seria = str(row[2]).strip() if row[2] else ""
            aktyw = str(row[4]).strip().upper() if len(row) > 4 and row[4] else ""
            if nr and nr != "None" and aktyw == "TAK":
                sets.append((nr, name, seria))
        return sets
    except Exception as e:
        print(f"BLAD odczytu zestawow: {e}")
        sys.exit(1)

def save_csv(rows: list):
    existing = []
    if Path(CSV_FILE).exists():
        with open(CSV_FILE, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            existing = list(reader)
    today = date.today().strftime("%Y-%m-%d")
    existing = [r for r in existing if r.get("Data") != today]
    all_rows = existing + rows
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["Data", "Nr setu", "Nazwa zestawu", "Seria LEGO",
                      "Najnizsza cena PLN", "Liczba ofert", "Kupilo ostatnio"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)
    print(f"Zapisano {len(rows)} wierszy (lacznie: {len(all_rows)}).")

def main():
    today = date.today().strftime("%Y-%m-%d")
    print(f"[{today}] Start.")

    if not REFRESH_TOKEN:
        print("BLAD: Brak ALLEGRO_REFRESH_TOKEN w Secrets.")
        print("Uruchom get_token.py lokalnie zeby uzyskac token.")
        sys.exit(1)

    token = get_token_from_refresh()
    sets  = read_sets()
    print(f"Znaleziono {len(sets)} zestawow.")

    rows = []
    for nr, name, seria in sets:
        print(f"  -> {nr} {name}...", end=" ", flush=True)
        items = fetch_offers(token, nr)
        stats = parse_stats(items, nr)
        print(f"cena: {stats['min_price']} | ofert: {stats['offer_count']} | kupilo: {stats['bought_last']}")
        rows.append({
            "Data":               today,
            "Nr setu":            nr,
            "Nazwa zestawu":      name,
            "Seria LEGO":         seria,
            "Najnizsza cena PLN": stats["min_price"],
            "Liczba ofert":       stats["offer_count"],
            "Kupilo ostatnio":    stats["bought_last"],
        })

    save_csv(rows)
    print("Gotowe!")

if __name__ == "__main__":
    main()
