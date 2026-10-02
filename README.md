# Konwerter e-sprawozdań do Excela

Zamienia e-sprawozdanie finansowe złożone do KRS (plik XML, także podpisany `.xml.xades`)
na czytelny arkusz Excel: bilans i rachunek zysków i strat z **oficjalnymi polskimi nazwami pozycji**
zamiast kodów takich jak `Aktywa_B_III_1_C`.

**▶ Wersja online:** https://ikangela.github.io/konwerter-e-sprawozdan/

Konwersja działa **w całości w przeglądarce** – Python uruchamia się lokalnie dzięki
[Pyodide](https://pyodide.org) (WebAssembly). Plik ze sprawozdaniem nie jest nigdzie wysyłany.

## Funkcje

- odczyt e-sprawozdań jednostek innych (pełny wzór) w wariancie porównawczym i kalkulacyjnym rachunku zysków i strat,
- obsługa plików podpisanych elektronicznie (`.xml.xades`) bez rozpakowywania,
- nazwy pozycji pobrane ze schematu XSD Ministerstwa Finansów, kody w zapisie z ustawy (np. `B.III.1.c`),
- dane za rok bieżący, poprzedni i – jeśli są – poprzedni po przekształceniu,
- **kontrola poprawności**: aktywa = pasywa, zysk netto w rachunku wyników = zysk netto w bilansie,
- Excel w 4 arkuszach (Informacje, Aktywa, Pasywa, RZiS) z formatowaniem kwot i wcięciami.

## Jak to działa

```
 serwer (GitHub Pages)                 przeglądarka użytkownika
 ┌────────────────────┐               ┌──────────────────────────────────┐
 │ index.html, app.js │ ── raz ─────▶ │ Python (Pyodide) + konwerter.py  │
 │ pyodide/           │  (narzędzia)  │    plik XML ──▶ Excel (.xlsx)    │
 │ konwerter.py       │               │    z dysku       na dysk         │
 └────────────────────┘               └──────────────────────────────────┘
          ✗ plik ze sprawozdaniem nigdy nie trafia na serwer
```

| Plik | Rola |
|---|---|
| `konwerter.py` | Rdzeń: bajty XML → bajty .xlsx. Nie korzysta z dysku ani sieci, więc ten sam kod działa w przeglądarce i w terminalu. |
| `app.js`, `index.html` | Interfejs: uruchomienie Pyodide, wybór/przeciąganie pliku, podgląd wyniku, pobranie. |
| `cli.py` | Wersja z linii poleceń. |
| `etykiety.json` | Nazwy 563 pozycji (60 KB) wyciągnięte ze schematu MF (730 KB). |
| `narzedzia/zbuduj_etykiety.py` | Odtwarza `etykiety.json` po wydaniu nowej wersji schematu. |

Kluczowa decyzja: **nazwy pozycji są kluczowane całą ścieżką** (np. `RZiSPor/A/A_J`), a nie samym kodem –
bo ten sam kod (`A`) oznacza w sprawozdaniu różne rzeczy w różnych sekcjach.

## Bezpieczeństwo i prywatność

- plik jest przetwarzany wyłącznie w pamięci przeglądarki; po pobraniu Excela wynik jest automatycznie usuwany z pamięci,
- Pyodide i biblioteki są dołączone do repozytorium (sumy kontrolne zweryfikowane z npm/PyPI) – strona nie pobiera kodu z zewnętrznych CDN,
- polityka **Content-Security-Policy** blokuje połączenia z innymi adresami niż własny serwer (poza czcionkami Google Fonts),
- teksty z pliku zapisywane są w Excelu jako tekst, nie formuły (ochrona przed *formula injection*),
- odporność na złośliwe XML (XXE, „bomba XML”) – sprawdzone testami,
- treść z pliku wyświetlana jest wyłącznie jako tekst (`textContent`), nigdy jako HTML.

## Uruchomienie lokalne

**Strona:** dwuklik w `podglad.bat` (wymaga Node.js) i otwarcie http://127.0.0.1:8765.
Samo otwarcie `index.html` z dysku nie zadziała – przeglądarka wymaga serwera.

**Linia poleceń:**

```
pip install -r requirements.txt
python cli.py sprawozdanie.xml [wynik.xlsx]
```

## Technologie

Python 3 · xml.etree.ElementTree · openpyxl · Pyodide (WebAssembly) · JavaScript · HTML/CSS · GitHub Pages

## Ograniczenia

- obsługiwane są sprawozdania **jednostek innych** (pełny wzór); jednostki małe i mikro mają inne struktury,
- sprawozdania spółek giełdowych w formacie ESEF (`.xhtml`) nie są obsługiwane,
- arkusze obejmują bilans i rachunek zysków i strat (bez przepływów pieniężnych i zestawienia zmian w kapitale).

## Autorka

Angelika Berecka · [ikangela.pl](https://ikangela.pl) · [GitHub](https://github.com/IKANGELA)

Licencja: MIT (zob. `LICENSE`; dołączone biblioteki zachowują własne licencje).
