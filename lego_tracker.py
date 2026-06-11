"""
LEGO Market Tracker
Pobiera dane z Allegro i zapisuje do CSV (dla Google Sheets).
"""

import os, sys, json, base64, urllib.request, urllib.parse, csv
from datetime import date
from pathlib import Path

EXCEL_FILE = "LEGO_Monitor.xlsx"
CSV_FILE   = "lego_dane.csv"

def get_token() -> str:
    CLIENT_ID     = os.environ.get("ALLEGRO_CLIENT_ID", "")
    CLIENT_SECRET = os.environ.get("ALLEGRO_CLIENT_SECRET", "")
    if not CLIENT_ID or not CLIENT_SECRET:
        print("BLAD: Brak CLIENT_ID lub CLIENT_SECRET.")
        sys.exit(1)
    creds = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()
    data  = urllib.parse.urlencode({"grant_type": "client_credentials"}).encode()
    req   = urllib.request.Request(
        "https://allegro.pl/auth/oauth/token", data=data,
        headers={"Authorization": f"Basic {creds}",
                 "Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            print("Token OK.")
            return result["access_token"]
    except urllib.error.HTTPError as e:
        print(f"BLAD tokenu: {e.read().decode()}")
        sys.exit(1)

def fetch_offers(token: str, set_nr: str) -> list:
    params = urllib.parse.urlencode({
        "phrase":   f"LEGO {set_nr}",
        "limit":    100,
        "sort":     "price",
        "fallback": "false",
    })
    req = urllib.request.Request(
        f"https://api.allegro.pl/offers/listing?{params}",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.allegro.public.v1+json"})
    try:
        with urllib.request.urlopen(req) as r:
            data = json.loads(r.read())
        return (data.get("items", {}).get("regular", []) +
                data.get("items", {}).get("promoted", []))
    except Exception as e:
        print(f"  BLAD pobierania: {e}")
        return []

def parse_stats(items: list, set_nr: str) -> dict:
    prices      = []
    bought_last = 0
    for it in items:
        # Filtruj tylko oferty zawierające numer setu w nazwie
        name = it.get("name", "").lower()
        if set_nr.lower() not in name:
            continue
        try:
            prices.append(float(it["sellingMode"]["price"]["amount"]))
        except (KeyError, TypeError, ValueError):
            pass
        stats = it.get("stats", {})
        val = stats.get("boughtCount") or stats.get("popularityScore") or 0
        bought_last = max(bought_last, int(val))
    return {
        "min_price":   round(min(prices), 2) if prices else "",
        "offer_count": len(prices),
        "bought_last": bought_last,
    }

def read_sets() -> list:
    # Czyta zestawy z CSV jesli istnieje, inaczej z xlsx
    sets_file = Path("zestawy.csv")
    if sets_file.exists():
        sets = []
        with open(sets_file, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("Aktywny","").strip().upper() == "TAK":
                    sets.append((row["Nr setu"].strip(),
                                 row["Nazwa zestawu"].strip(),
                                 row["Seria LEGO"].strip()))
        return sets

    # Fallback: czyta z xlsx
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
            if nr and aktyw == "TAK":
                sets.append((nr, name, seria))
        return sets
    except Exception as e:
        print(f"BLAD odczytu zestawow: {e}")
        sys.exit(1)

def save_csv(rows: list):
    file_exists = Path(CSV_FILE).exists()
    # Wczytaj istniejace dane
    existing = []
    if file_exists:
        with open(CSV_FILE, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            existing = list(reader)

    today = str(date.today())
    # Usun dzisiejsze wpisy jesli juz sa (re-run)
    existing = [r for r in existing if r.get("Data") != today]
    # Dodaj nowe
    all_rows = existing + rows

    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["Data","Nr setu","Nazwa zestawu","Seria LEGO",
                      "Najnizsza cena PLN","Liczba ofert","Kupilo ostatnio"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"Zapisano {len(rows)} wierszy do {CSV_FILE} (lacznie {len(all_rows)} wierszy historycznych).")

def main():
    today = str(date.today())
    print(f"[{today}] Start.")

    token = get_token()
    sets  = read_sets()
    print(f"Znaleziono {len(sets)} aktywnych zestawow.")

    rows = []
    for nr, name, seria in sets:
        print(f"  -> {nr} {name}...", end=" ", flush=True)
        items = fetch_offers(token, nr)
        stats = parse_stats(items, nr)
        print(f"cena: {stats['min_price']} zl | ofert: {stats['offer_count']} | kupilo: {stats['bought_last']}")
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
