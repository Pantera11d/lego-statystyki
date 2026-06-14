"""
LEGO Market Tracker v5
- Refresh token z auto-rotacja zapisywana do GitHub Secrets
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

# Dane do aktualizacji GitHub Secret
GH_TOKEN  = os.environ.get("GH_PAT", "")          # Personal Access Token
GH_REPO   = os.environ.get("GH_REPO", "")          # np. "Pantera11d/lego-statystyki"

def get_token_from_refresh() -> tuple:
    """Pobiera nowy access token i (jesli jest) nowy refresh token."""
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
            access_token  = result["access_token"]
            new_refresh   = result.get("refresh_token", "")
            return access_token, new_refresh
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"BLAD refresh tokena: {body}")
        print("Refresh token wygasl calkowicie. Trzeba wygenerowac nowy recznie.")
        sys.exit(1)

# ── Aktualizacja GitHub Secret przez API ─────────────────────────────────────
def update_github_secret(secret_name: str, secret_value: str):
    """Aktualizuje Secret w repo uzywajac GitHub REST API + PyNaCl do szyfrowania."""
    if not GH_TOKEN or not GH_REPO:
        print("UWAGA: Brak GH_PAT lub GH_REPO - nie moge zaktualizowac Secret automatycznie.")
        print(f"NOWY_REFRESH_TOKEN={secret_value}")
        return False

    try:
        from nacl import encoding, public
    except ImportError:
        print("BLAD: brak biblioteki PyNaCl. Zainstaluj: pip install pynacl")
        print(f"NOWY_REFRESH_TOKEN={secret_value}")
        return False

    # 1. Pobierz public key repo
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/public-key",
        headers={
            "Authorization": f"Bearer {GH_TOKEN}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })
    with urllib.request.urlopen(req) as r:
        key_data = json.loads(r.read())

    public_key = public.PublicKey(key_data["key"].encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(public_key)
    encrypted  = sealed_box.encrypt(secret_value.encode("utf-8"))
    encrypted_b64 = base64.b64encode(encrypted).decode("utf-8")

    # 2. Wyslij zaszyfrowany secret
    payload = json.dumps({
        "encrypted_value": encrypted_b64,
        "key_id": key_data["key_id"],
    }).encode("utf-8")

    req2 = urllib.request.Request(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/{secret_name}",
        data=payload,
        method="PUT",
        headers={
            "Authorization": f"Bearer {GH_TOKEN}",
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
            "X-GitHub-Api-Version": "2022-11-28",
        })
    try:
        with urllib.request.urlopen(req2) as r:
            print(f"Secret {secret_name} zaktualizowany automatycznie. (status {r.status})")
            return True
    except urllib.error.HTTPError as e:
        print(f"BLAD aktualizacji secret: {e.read().decode()}")
        return False

# ── Allegro: pobieranie ofert ────────────────────────────────────────────────
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
        sys.exit(1)

    access_token, new_refresh = get_token_from_refresh()

    # Jesli Allegro wydalo nowy refresh token, zapisz go do GitHub Secrets
    if new_refresh and new_refresh != REFRESH_TOKEN:
        print("Otrzymano NOWY refresh token - aktualizuje GitHub Secret...")
        update_github_secret("ALLEGRO_REFRESH_TOKEN", new_refresh)
    else:
        print("Refresh token bez zmian.")

    sets = read_sets()
    print(f"Znaleziono {len(sets)} zestawow.")

    rows = []
    for nr, name, seria in sets:
        print(f"  -> {nr} {name}...", end=" ", flush=True)
        items = fetch_offers(access_token, nr)
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
