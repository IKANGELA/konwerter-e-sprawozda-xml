"""
Wyciąga z oficjalnego schematu XSD Ministerstwa Finansów same nazwy pozycji
i zapisuje je jako mały plik etykiety.json (ok. 30 KB zamiast 730 KB schematu).

Uruchamia się raz, po pobraniu nowej wersji schematu:
    python narzedzia/zbuduj_etykiety.py
"""

import json
from pathlib import Path
import xml.etree.ElementTree as ET

PROJEKT = Path(__file__).resolve().parent.parent
SCHEMAT = PROJEKT / "narzedzia" / "schematy" / "JednostkaInnaStrukturyDanychSprFin_v2-0E.xsd"
WYNIK = PROJEKT / "etykiety.json"

XS = "{http://www.w3.org/2001/XMLSchema}"
# Elementy, od których zaczynają się ścieżki pozycji w sprawozdaniu.
SEKCJE = {"Aktywa", "Pasywa", "RZiSPor", "RZiSKalk"}


def zbuduj(schemat):
    """
    Słownik {ścieżka: nazwa pozycji}, np. {'RZiSPor/A/A_J': '– od jednostek powiązanych'}.
    Kluczem jest cała ścieżka, bo ten sam kod (np. 'A') znaczy co innego
    w rachunku zysków i strat, a co innego w przepływach pieniężnych.
    """
    etykiety = {}

    def przejdz(wezel, sciezka):
        for dziecko in wezel:
            if dziecko.tag == XS + "element" and dziecko.get("name"):
                n = dziecko.get("name")
                nowa = sciezka + [n] if (sciezka or n in SEKCJE) else []
                if nowa:
                    opis = dziecko.find(f"{XS}annotation/{XS}documentation")
                    if opis is not None and opis.text:
                        etykiety.setdefault("/".join(nowa), opis.text.strip())
                przejdz(dziecko, nowa)
            else:
                przejdz(dziecko, sciezka)

    przejdz(ET.parse(schemat).getroot(), [])
    return etykiety


if __name__ == "__main__":
    etykiety = zbuduj(SCHEMAT)
    WYNIK.write_text(json.dumps(etykiety, ensure_ascii=False, separators=(",", ":")),
                     encoding="utf-8")
    print(f"Zapisano {len(etykiety)} etykiet do {WYNIK.name} "
          f"({WYNIK.stat().st_size // 1024} KB)")
