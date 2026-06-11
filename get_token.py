"""
URUCHOM TEN PLIK RAZ NA KOMPUTERZE (nie na GitHubie)
Uzyskuje refresh token do Allegro - wazny 3 miesiace.

Jak uzyc:
1. Zainstaluj Python na komputerze
2. Uruchom: python get_token.py
3. Otworzy sie przeglądarka - zaloguj sie na Allegro
4. Skopiuj refresh_token z wyniku
5. Wklej go do GitHub Secrets jako ALLEGRO_REFRESH_TOKEN
"""

import http.server
import urllib.parse
import urllib.request
import webbrowser
import json
import base64
import threading

# ── WPISZ SWOJE DANE ────────────────────────────────────────────────────────
CLIENT_ID     = "WPISZ_SWOJ_CLIENT_ID"
CLIENT_SECRET = "WPISZ_SWOJ_CLIENT_SECRET"
REDIRECT_URI  = "http://localhost:8080"
# ─────────────────────────────────────────────────────────────────────────────

auth_code = None

class CallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if "code" in params:
            auth_code = params["code"][0]
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(b"""
            <html><body style='font-family:Arial;text-align:center;padding:50px'>
            <h2>&#10003; Zalogowano pomyslnie!</h2>
            <p>Mozesz zamknac ta karte i wrocic do terminala.</p>
            </body></html>
            """)
        else:
            self.send_response(400)
            self.end_headers()

    def log_message(self, format, *args):
        pass  # wycisz logi serwera

def get_tokens(code: str) -> dict:
    creds = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()
    data  = urllib.parse.urlencode({
        "grant_type":   "authorization_code",
        "code":          code,
        "redirect_uri":  REDIRECT_URI,
    }).encode()
    req = urllib.request.Request(
        "https://allegro.pl/auth/oauth/token", data=data,
        headers={"Authorization": f"Basic {creds}",
                 "Content-Type":  "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def main():
    if CLIENT_ID == "WPISZ_SWOJ_CLIENT_ID":
        print("BLAD: Wpisz swoj CLIENT_ID i CLIENT_SECRET w pliku get_token.py!")
        return

    # Zbuduj URL logowania
    params = urllib.parse.urlencode({
        "response_type": "code",
        "client_id":     CLIENT_ID,
        "redirect_uri":  REDIRECT_URI,
        "prompt":        "confirm",
    })
    auth_url = f"https://allegro.pl/auth/oauth/authorize?{params}"

    print("=" * 60)
    print("KROK 1: Otwieram przegladarke...")
    print("Zaloguj sie na Allegro i zatwierdz dostep.")
    print("=" * 60)

    # Uruchom lokalny serwer na porcie 8080
    server = http.server.HTTPServer(("localhost", 8080), CallbackHandler)
    thread = threading.Thread(target=server.handle_request)
    thread.start()

    # Otworz przegladarke
    webbrowser.open(auth_url)
    print(f"\nJesli przeglądarka nie otworzyla sie automatycznie, wejdz na:")
    print(f"{auth_url}\n")

    # Czekaj na callback
    thread.join(timeout=120)

    if not auth_code:
        print("BLAD: Nie otrzymano kodu autoryzacji. Sprobuj ponownie.")
        return

    print("Kod autoryzacji otrzymany. Pobieram tokeny...")

    try:
        tokens = get_tokens(auth_code)
        refresh_token = tokens.get("refresh_token", "")
        access_token  = tokens.get("access_token", "")

        print("\n" + "=" * 60)
        print("SUKCES! Twoje tokeny:")
        print("=" * 60)
        print(f"\nACCESS TOKEN (wazny 12h):")
        print(access_token)
        print(f"\nREFRESH TOKEN (wazny 3 miesiace):")
        print(refresh_token)
        print("\n" + "=" * 60)
        print("CO TERAZ ZROBIC:")
        print("1. Wejdz na GitHub -> Settings -> Secrets -> Actions")
        print("2. Dodaj nowy secret:")
        print("   Name:  ALLEGRO_REFRESH_TOKEN")
        print("   Value: (wklej refresh token powyzej)")
        print("=" * 60)

        # Zapisz do pliku dla pewnosci
        with open("tokeny.txt", "w") as f:
            f.write(f"refresh_token={refresh_token}\n")
            f.write(f"access_token={access_token}\n")
        print("\nTokeny zapisane rowniez do pliku: tokeny.txt")
        print("NIE WGRYWAJ tego pliku na GitHub!")

    except Exception as e:
        print(f"BLAD pobierania tokenow: {e}")

if __name__ == "__main__":
    main()
