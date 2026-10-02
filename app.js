// Konwerter e-sprawozdań – obsługa strony.
// Wszystkie pliki (Python/Pyodide, biblioteki, kod) są wczytywane z tego samego serwera,
// a polityka CSP w index.html blokuje połączenia z innymi adresami.

const PLIK_PRZYKLADU = "przyklady/meble-przyklad-2025.xml";
const BIBLIOTEKI = [
    "biblioteki/et_xmlfile-2.0.0-py3-none-any.whl",
    "biblioteki/openpyxl-3.1.5-py2.py3-none-any.whl",
];

const $ = (id) => document.getElementById(id);
const strefa = $("strefa"), wejscie = $("plik"), wynik = $("wynik");
let pyodide = null;
let urlPobierania = null;

function ustawStatus(tekst, typ = "") {
    $("status").className = "status " + typ;
    $("status-tekst").textContent = tekst;
}

// --- 1. Start: Python w przeglądarce + biblioteka openpyxl + nasz kod ---
async function uruchom() {
    try {
        pyodide = await loadPyodide({ indexURL: new URL("pyodide/", location.href).href });
        await pyodide.loadPackage(BIBLIOTEKI.map((b) => new URL(b, location.href).href));

        const [kod, etykiety] = await Promise.all([
            fetch("konwerter.py").then((r) => r.text()),
            fetch("etykiety.json").then((r) => r.text()),
        ]);
        pyodide.FS.writeFile("konwerter.py", kod);
        pyodide.globals.set("etykiety_json", etykiety);
        pyodide.runPython(`
import json, base64, konwerter
ETYKIETY = json.loads(etykiety_json)

def przetworz(sciezka):
    with open(sciezka, "rb") as f:
        xlsx, s = konwerter.konwertuj(f.read(), ETYKIETY)
    info = {k: s[k] for k in ("firma", "krs", "data_od", "data_do", "wariant", "rok",
                              "kontrole", "podsumowanie")}
    return json.dumps({"xlsx": base64.b64encode(xlsx).decode(), "info": info})
`);
        strefa.classList.remove("zablokowana");
        wejscie.disabled = false;
        ustawStatus("Gotowe – wybierz plik e-sprawozdania.", "gotowy");

        // przycisk przykładu pokazujemy tylko, jeśli plik przykładu istnieje
        const jest = await fetch(PLIK_PRZYKLADU, { method: "HEAD" }).then((r) => r.ok).catch(() => false);
        $("przyklad").hidden = !jest;
        $("przyklad").disabled = false;
    } catch (e) {
        console.error(e);
        ustawStatus("Nie udało się uruchomić konwertera. Odśwież stronę lub spróbuj w innej przeglądarce.", "blad");
    }
}

// --- 2. Konwersja jednego pliku ---
async function konwertuj(nazwaPliku, bajty) {
    ustawStatus(`Przetwarzam „${nazwaPliku}”…`);
    wynik.classList.remove("widoczny");
    try {
        pyodide.FS.writeFile("/tmp/wejscie.xml", bajty);
        const przetworz = pyodide.globals.get("przetworz");
        const odp = JSON.parse(przetworz("/tmp/wejscie.xml"));
        przetworz.destroy();
        pokazWynik(odp.info, odp.xlsx);
        ustawStatus("Gotowe.", "gotowy");
    } catch (e) {
        console.error(e);
        const msg = String(e.message || e);
        const blad = msg.match(/BladSprawozdania: (.*)/);
        ustawStatus(blad ? blad[1] : "Nie udało się odczytać tego pliku. Czy to na pewno e-sprawozdanie finansowe w XML?", "blad");
    } finally {
        try { pyodide.FS.unlink("/tmp/wejscie.xml"); } catch { /* plik mógł nie powstać */ }
    }
}

// --- 3. Wyświetlenie wyniku (tylko textContent – treść z pliku nigdy nie jest wstawiana jako HTML) ---
const zl = new Intl.NumberFormat("pl-PL", { maximumFractionDigits: 0 });

function pokazWynik(info, xlsxB64) {
    $("firma").textContent = info.firma || "(brak nazwy spółki)";
    $("meta").textContent =
        `KRS ${info.krs || "—"} · okres ${info.data_od} – ${info.data_do} · rachunek wyników: wariant ${info.wariant}`;

    $("liczby").replaceChildren(...Object.entries(info.podsumowanie).map(([etykieta, kwota]) => {
        const div = document.createElement("div");
        div.className = "liczba";
        const span = document.createElement("span");
        span.textContent = `${etykieta} ${info.rok}`;
        const strong = document.createElement("strong");
        strong.textContent = kwota == null ? "—" : `${zl.format(kwota)} zł`;
        if (kwota < 0) strong.classList.add("ujemna");
        div.append(span, strong);
        return div;
    }));

    $("kontrole").replaceChildren(...info.kontrole.map((k) => {
        const li = document.createElement("li");
        li.className = k.ok ? "ok" : "nie";
        li.textContent = k.opis;
        return li;
    }));

    const bajty = Uint8Array.from(atob(xlsxB64), (c) => c.charCodeAt(0));
    if (urlPobierania) URL.revokeObjectURL(urlPobierania);
    urlPobierania = URL.createObjectURL(new Blob([bajty], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }));
    const nazwa = (info.firma || "sprawozdanie").replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_|_$/g, "");
    $("pobierz").href = urlPobierania;
    $("pobierz").download = `${nazwa}_${info.rok || ""}.xlsx`;

    wynik.classList.add("widoczny");
}

// --- 4. Obsługa wyboru pliku i przeciągania ---
const MAKS_ROZMIAR = 20 * 1024 * 1024; // e-sprawozdania mają zwykle 0,3–3 MB

async function wczytaj(plik) {
    if (!plik || !pyodide) return;
    if (plik.size > MAKS_ROZMIAR) {
        ustawStatus("Plik jest za duży (ponad 20 MB) – to raczej nie jest e-sprawozdanie.", "blad");
        return;
    }
    konwertuj(plik.name, new Uint8Array(await plik.arrayBuffer()));
}

wejscie.addEventListener("change", () => wczytaj(wejscie.files[0]));
strefa.addEventListener("dragover", (e) => { e.preventDefault(); strefa.classList.add("nad"); });
strefa.addEventListener("dragleave", () => strefa.classList.remove("nad"));
strefa.addEventListener("drop", (e) => {
    e.preventDefault();
    strefa.classList.remove("nad");
    wczytaj(e.dataTransfer.files[0]);
});
$("kolejny").addEventListener("click", () => { wejscie.value = ""; wejscie.click(); });

// --- Prywatność: po pobraniu usuwamy wynik z pamięci i z ekranu ---
// Kilka sekund opóźnienia, żeby przeglądarka zdążyła zapisać plik.
const OPOZNIENIE_CZYSZCZENIA = 3000;

function wyczyscDane() {
    if (urlPobierania) URL.revokeObjectURL(urlPobierania);   // Excel znika z pamięci
    urlPobierania = null;
    $("pobierz").removeAttribute("href");
    for (const id of ["firma", "meta"]) $(id).textContent = "";
    $("liczby").replaceChildren();
    $("kontrole").replaceChildren();
    wynik.classList.remove("widoczny");
    wejscie.value = "";
    try { pyodide.runPython("import gc; gc.collect()"); } catch { /* nic */ }
    ustawStatus("Plik pobrany. Dane sprawozdania zostały usunięte z pamięci przeglądarki.", "gotowy");
}

$("pobierz").addEventListener("click", () => {
    if (!urlPobierania) return;
    ustawStatus("Pobieranie… za chwilę dane zostaną usunięte z pamięci przeglądarki.");
    setTimeout(wyczyscDane, OPOZNIENIE_CZYSZCZENIA);
});
$("przyklad").addEventListener("click", async () => {
    const r = await fetch(PLIK_PRZYKLADU);
    konwertuj("przykładowe sprawozdanie", new Uint8Array(await r.arrayBuffer()));
});

// --- 5. Osadzenie na ikangela.pl: strona podaje swoją wysokość ramce (iframe).
// Wysyłana jest wyłącznie liczba (wysokość w pikselach), nigdy treść sprawozdania.
new ResizeObserver(() => {
    window.parent.postMessage({ konwerterWysokosc: document.documentElement.scrollHeight }, "*");
}).observe(document.body);

uruchom();
