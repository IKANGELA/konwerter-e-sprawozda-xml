"""
Konwerter e-sprawozdań - wersja z linii poleceń.

    python cli.py <plik.xml lub plik.xml.xades> [wynik.xlsx]
"""

import json
import sys
from pathlib import Path

from konwerter import BladSprawozdania, konwertuj

ETYKIETY = Path(__file__).resolve().parent / "etykiety.json"


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    plik = Path(sys.argv[1])
    etykiety = json.loads(ETYKIETY.read_text(encoding="utf-8"))

    try:
        xlsx, s = konwertuj(plik.read_bytes(), etykiety)
    except BladSprawozdania as e:
        sys.exit(f"Błąd: {e}")

    wynik = Path(sys.argv[2]) if len(sys.argv) > 2 else plik.with_name(
        f"{plik.name.split('.')[0]}.xlsx")
    wynik.write_bytes(xlsx)

    print(f"{s['firma']} (KRS {s['krs']}), {s['data_od']} – {s['data_do']}, RZiS {s['wariant']}")
    for k in s["kontrole"]:
        print(("  OK    " if k["ok"] else "  BŁĄD  ") + k["opis"])
    print(f"Zapisano: {wynik}")


if __name__ == "__main__":
    main()
