import pandas as pd
import requests
import os
import json
import math

# ======================================================
# 🔧 TEAMLEADER CONFIG — VIA RAILWAY ENV
# ======================================================

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

API_BASE = "https://api.focus.teamleader.eu"
TOKEN_URL = "https://focus.teamleader.eu/oauth2/access_token"

# 21% BTW
TAX_RATE_21_ID = "94da9f7d-9bf3-04fb-ac49-404ed252c381"

# Vaste kosten
MONTAGE_PER_FRONT = 34.71
INMETEN = 99.17
VRACHT = 60.00

PRIJS_SCHARNIER = 6.5          # voor renovatie-fronten (bestaande flow)
PRIJS_LADE = 184.0             # voor renovatie-fronten (bestaande flow)

# ======================================================
# 🔒 TOKEN MANAGEMENT — AUTOMATISCHE REFRESH + OPSLAAN
# ======================================================

TOKEN_FILE = "/app/refresh_token.txt"   # persistent binnen Railway container

def load_refresh_token():
    """Laadt refresh_token vanaf disk, of fallback naar ENV (1e keer)."""
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return f.read().strip()
    return os.getenv("REFRESH_TOKEN")

def save_refresh_token(token: str):
    """Slaat vernieuwde refresh_token op zodat altijd geldig blijft."""
    with open(TOKEN_FILE, "w") as f:
        f.write(token)

# Laad token bij opstart
REFRESH_TOKEN = load_refresh_token()


def get_access_token():
    """
    Haal een nieuwe access_token op. Sla vernieuwde refresh_token op.
    Werkt onbeperkt zonder opnieuw inloggen.
    """
    global REFRESH_TOKEN

    if not REFRESH_TOKEN:
        raise Exception("Geen refresh_token gevonden — log eerst in via de app.")

    data = {
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }

    resp = requests.post(TOKEN_URL, data=data)

    if resp.status_code != 200:
        raise Exception(f"Kon access_token niet vernieuwen: {resp.text}")

    tokens = resp.json()

    # Teamleader geeft ALTIJD een nieuwe refresh_token terug
    REFRESH_TOKEN = tokens["refresh_token"]
    save_refresh_token(REFRESH_TOKEN)

    return tokens["access_token"]


def request_with_auto_refresh(method: str, url: str, json_data=None, files=None):
    """API wrapper die automatisch token vernieuwt."""
    access_token = get_access_token()

    headers = {"Authorization": f"Bearer {access_token}"}
    if not files:
        headers["Content-Type"] = "application/json"

    resp = requests.request(method, url, headers=headers, json=json_data, files=files)
    return resp


# ======================================================
# 🧮 MODEL- EN PRIJSLOGICA (RENOVATIE-FRONTEN)
# ======================================================

MODEL_INFO = {
    "NOAH":  {"materiaal": "MDF gespoten", "prijs_per_front": 96.69, "passtuk": 189.26},
    "FEDDE": {"materiaal": "MDF gespoten", "prijs_per_front": 96.69, "passtuk": 189.26},
    "DAVE":  {"materiaal": "MDF gespoten", "prijs_per_front": 96.69, "passtuk": 189.26},
    "JOLIE": {"materiaal": "MDF gespoten", "prijs_per_front": 96.69, "passtuk": 189.26},
    "DEX":   {"materiaal": "MDF gespoten", "prijs_per_front": 96.69, "passtuk": 189.26},

    "JACK":  {"materiaal": "Eikenfineer", "prijs_per_front": 113.22, "passtuk": 209.92},
    "CHIEL": {"materiaal": "Eikenfineer", "prijs_per_front": 151.00, "passtuk": 209.92},
    "JAMES": {"materiaal": "Eikenfineer", "prijs_per_front": 195.00, "passtuk": 209.92},

    "SAM":   {"materiaal": "Noten fineer", "prijs_per_front": 169.00, "passtuk": 240.29},
    "DUKE":  {"materiaal": "Noten fineer", "prijs_per_front": 185.00, "passtuk": 240.29},
}

FRONT_DESCRIPTION_CONFIG = {
    "NOAH": {"titel": "Keukenrenovatie model Noah", "materiaal": "MDF Gespoten - vlak", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde wit"},
    "FEDDE": {"titel": "Keukenrenovatie model Fedde", "materiaal": "MDF Gespoten - greeploos", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde wit"},
    "DAVE": {"titel": "Keukenrenovatie model Dave", "materiaal": "MDF Gespoten - 70mm kader", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde wit"},
    "JOLIE": {"titel": "Keukenrenovatie model Jolie", "materiaal": "MDF Gespoten - 25mm kader", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde wit"},
    "DEX": {"titel": "Keukenrenovatie model Dex", "materiaal": "MDF Gespoten - V-groef", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde wit"},

    "JACK": {"titel": "Keukenrenovatie model Jack", "materiaal": "Eiken fineer - vlak", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja"},
    "JAMES": {"titel": "Keukenrenovatie model James", "materiaal": "Eiken fineer - 10mm massief kader", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja"},
    "CHIEL": {"titel": "Keukenrenovatie model Chiel", "materiaal": "Eiken fineer - greeploos", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja"},

    "SAM": {"titel": "Keukenrenovatie model Sam", "materiaal": "Noten fineer - vlak", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja"},
    "DUKE": {"titel": "Keukenrenovatie model Duke", "materiaal": "Noten fineer - greeploos", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja"},
}


def bepaal_model(g2, h2):
    mapping = {
        ("K01 - vlak", "MDF gespoten"): "NOAH",
        ("K02 - greeploos", "MDF gespoten"): "FEDDE",
        ("K04 - 70mm kader", "MDF gespoten"): "DAVE",
        ("K05 - 25mm kader", "MDF gespoten"): "JOLIE",
        ("K13 - Vgroef", "MDF gespoten"): "DEX",

        ("K09 - 10mm kader", "Eikenfineer"): "JAMES",
        ("K02 - greeploos", "Eikenfineer"): "CHIEL",
        ("K01 - vlak", "Eikenfineer"): "JACK",

        ("K01 - vlak", "Noten fineer"): "SAM",
        ("K02 - greeploos", "Noten fineer"): "DUKE",
    }
    return mapping.get((str(g2).strip(), str(h2).strip()), None)


# ======================================================
# 🧮 MAATWERK KASTEN – EXTRA PRIJSLOGICA (m² + tabblad 2)
# ======================================================

# m²-prijzen fronten (maatwerk kasten)
FRONT_M2_PRICES = {
    "NOAH": 76.0,
    "FEDDE": 80.0,
    "DEX": 103.0,
    "DAVE": 103.0,
    "JOLIE": 103.0,
    "JACK": 114.5,
    "CHIEL": 149.50,
    "JAMES": 202.50,
    "SAM": 141.50,
    "DUKE": 186.50,
}

# materiaal-groepen voor zijkanten → vlak model
# MDF  -> NOAH
# EIKEN -> JACK
# NOTEN -> SAM
VLAK_MODEL_PER_MATERIAALGROEP = {
    "MDF": "NOAH",
    "EIKEN": "JACK",
    "NOTEN": "SAM",
}

# staffels voor breedte (mm)
BREEDTE_STAFFELS = [300, 400, 500, 600, 800, 900, 1000, 1200]

# prijs-tabellen uit je FinanceKing Excel (rij -> {staffel: prijs})
PRICE_TABLE = {
    # A-kast geschikt voor lades (rij 7)
    7: {300: 45.0, 400: 47.0, 500: 49.0, 600: 51.0, 800: 59.0, 900: 61.0, 1000: 63.0, 1200: 66.0},
    # A-kast geschikt voor planken (rij 8)
    8: {300: 51.0, 400: 55.0, 500: 60.0, 600: 64.0, 800: 77.0, 900: 81.0, 1000: 85.0, 1200: 92.0},

    # B-kasten (hoge kasten)
    # Leeg 2080–2770 (rij 11)
    11: {300: 85.0, 400: 89.0, 500: 92.0, 600: 97.0, 800: 103.0, 900: 108.0, 1000: 114.0},
    # Leeg 1001–2079 (rij 12)
    12: {300: 80.0, 400: 84.0, 500: 87.0, 600: 90.0, 800: 95.0, 900: 98.0, 1000: 103.0},

    # C-kasten (hangkasten) – corpusprijzen
    # tot 390mm (leeg) (rij 15)
    15: {300: 35.0, 400: 37.0, 500: 39.0, 600: 41.0, 800: 47.0, 900: 48.0, 1000: 50.0, 1200: 56.0},
    # 391–520mm (1 plank inbegrepen) (rij 16)
    16: {300: 39.0, 400: 42.0, 500: 45.0, 600: 47.0, 800: 56.0, 900: 58.0, 1000: 61.0, 1200: 67.0},
    # 521–780mm (2 planken inbegrepen) (rij 17)
    17: {300: 42.0, 400: 45.0, 500: 49.0, 600: 50.0, 800: 57.0, 900: 61.0, 1000: 66.0, 1200: 70.0},
    # vanaf 781mm (2 planken inbegrepen) (rij 18)
    18: {300: 46.0, 400: 48.0, 500: 51.0, 600: 55.0, 800: 61.0, 900: 65.0, 1000: 68.0, 1200: 75.0},

    # Inrichtingen / accessoires
    # B21 = Klepscharnieren (rij 21)
    21: {300: 63.0, 400: 63.0, 500: 63.0, 600: 63.0, 800: 63.0, 900: 63.0, 1000: 63.0, 1200: 63.0},
    # B22 = Plank voor A of B kast
    22: {300: 4.5, 400: 6.3, 500: 7.5, 600: 7.6, 800: 10.6, 900: 11.8, 1000: 13.3, 1200: 15.75},
    # B23 = Lades
    23: {300: 68.0, 400: 68.0, 500: 68.0, 600: 68.0, 800: 68.0, 900: 68.0, 1000: 68.0, 1200: 68.0},
    # B24 = Push to open lade
    24: {300: 109.0, 400: 109.0, 500: 109.0, 600: 109.0, 800: 109.0, 900: 109.0, 1000: 109.0, 1200: 109.0},
    # B25 = Scharnier (per stuk)
    25: {300: 9.7, 400: 9.7, 500: 9.7, 600: 9.7, 800: 9.7, 900: 9.7, 1000: 9.7, 1200: 9.7},
    # B26 = Bestek bak
    26: {300: 33.0, 400: 33.0, 500: 33.0, 600: 33.0, 800: 33.0, 900: 33.0, 1000: 33.0, 1200: 33.0},
    # B27 = Plastic bescherming spoelkast
    27: {300: 33.0, 400: 33.0, 500: 33.0, 600: 33.0, 800: 33.0, 900: 33.0, 1000: 33.0, 1200: 33.0},
    # B28 = Apothekers lade
    28: {300: 546.0, 400: 546.0, 500: 546.0},
    # B29 = Carrouselsysteem
    29: {800: 452.0, 900: 452.0, 1000: 452.0, 1200: 452.0},
}


def _kies_staffel(breedte_mm: float) -> int:
    """
    Kies altijd de eerstvolgende staffel OMHOOG.
    Voorbeeld: 610mm → 800mm.
    """
    if breedte_mm is None or math.isnan(float(breedte_mm)):
        return BREEDTE_STAFFELS[0]
    b = float(breedte_mm)
    for s in BREEDTE_STAFFELS:
        if b <= s:
            return s
    # als groter dan grootste staffel → pak grootste
    return BREEDTE_STAFFELS[-1]


def _prijs_uit_tabel(rij: int, breedte_mm: float) -> float:
    """Haalt prijs op uit PRICE_TABLE op basis van rij + staffelbreedte."""
    staffel = _kies_staffel(breedte_mm)
    tabel = PRICE_TABLE.get(rij, {})
    if staffel in tabel:
        return tabel[staffel]
    # fallback: pak hoogste beschikbare staffel in die rij
    if not tabel:
        return 0.0
    max_staffel = max(tabel.keys())
    return tabel[max_staffel]


def _bepaal_materiaal_groep(frontmodel: str) -> str:
    """
    Zet model om naar materiaal-groep MDF / EIKEN / NOTEN.
    Gebaseerd op MODEL_INFO materiaalvelden.
    """
    info = MODEL_INFO.get(frontmodel)
    if not info:
        return "MDF"
    mat = info["materiaal"].lower()
    if "eiken" in mat:
        return "EIKEN"
    if "noten" in mat:
        return "NOTEN"
    return "MDF"


def _parse_int(value):
    """Robuust omzetten naar int, 'Nee' → 0."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return 0
    s = str(value).strip().lower()
    if s in ("", "nee", "nvt", "n.v.t."):
        return 0
    try:
        return int(float(s.replace(",", ".")))
    except Exception:
        return 0


def _parse_inrichting_counts(tekst: str):
    """
    Haalt aantallen planken en lades uit een tekst als:
    '3x plank, 1x lade'
    """
    import re
    text = str(tekst).lower()
    plank = 0
    lade = 0
    for qty, kind in re.findall(r"(\d+)\s*x?\s*(plank|planken|lade|lades|laden)", text):
        q = int(qty)
        if kind.startswith("plank"):
            plank += q
        else:
            lade += q
    return plank, lade


def _kast_titel(kast_type: str) -> str:
    kast_type = (kast_type or "").strip().upper()
    if kast_type == "A":
        return "Maatwerk onderkast"
    if kast_type == "B":
        return "Maatwerk hoge kast"
    if kast_type == "C":
        return "Maatwerk hangkast"
    return "Maatwerk kast"


def _beschrijving_uit_kastdict(k: dict) -> str:
    """
    Bouwt de beschrijving precies volgens de velden uit tabblad 2.
    (1-op-1 opsomming)
    """
    regels = [
        f"Type kast: {k.get('type_kast', '')}",
        f"Hoogte: {k.get('hoogte', '')} mm",
        f"Breedte: {k.get('breedte', '')} mm",
        f"Diepte: {k.get('diepte', '')} mm",
        f"Hoogte pootje: {k.get('hoogte_pootje', '')} mm",
        f"Zichtbare zijde: {k.get('zichtbare_zijde', '')}",
        f"Inrichting: {k.get('inrichting', '')}",
        f"Scharnieren: {k.get('scharnieren', '')}",
        f"Front: {k.get('frontmodel', '')}",
        f"Front indeling: {k.get('front_indeling', '')}",
        f"Kleur: {k.get('kleur_front', '')}",
        f"Dubbelzijdig afgewerkt: {k.get('dubbelzijdig', '')}",
        f"Handgreep: {k.get('handgreep', '')}",
        f"Afwerking: {k.get('afwerking', '')}",
    ]
    return "\r\n".join(regels)


def _bereken_maatwerk_kastprijs(k: dict) -> float:
    """
    Bereken de totale inkoopprijs voor één maatwerk kast (excl. btw).
    Bestaat uit:
      - corpusprijs (A/B/C) op basis van staffel
      - inrichting (planken, lades, etc.)
      - scharnieren
      - fronten + zichtbare zijden op m²-prijs + 40% opslag
    """
    # Basisgegevens
    kast_type = str(k.get("type_kast", "")).strip().upper()
    hoogte = float(k.get("hoogte") or 0)
    breedte = float(k.get("breedte") or 0)
    diepte = float(k.get("diepte") or 0)
    zichtbare_zijde = str(k.get("zichtbare_zijde") or "").strip().lower()
    inrichting_tekst = str(k.get("inrichting") or "")
    scharnieren_aantal = _parse_int(k.get("scharnieren"))
    frontmodel = str(k.get("frontmodel") or "").strip().upper()
    aantal_fronten = _parse_int(k.get("aantal_fronten"))

    # ---------- 1. CORPUSPRIJS OP BASIS VAN TYPE / HOOGTE ----------
    corpus_rij = None

    if kast_type == "A":
        # A = onderkasten tot 1000mm
        # subtype bepalen via inrichting
        plank_count, lade_count = _parse_inrichting_counts(inrichting_tekst)
        if lade_count > 0:
            # Onderkast geschikt voor lades
            corpus_rij = 7
        elif plank_count > 0:
            # Onderkast geschikt voor planken (incl. 2 planken)
            corpus_rij = 8
        else:
            # fallback: ladekast-prijstabel
            corpus_rij = 7

    elif kast_type == "B":
        # B = hoge kasten tot 2770mm
        if hoogte >= 2080:
            # Hoge kast (leeg) 2080–2770
            corpus_rij = 11
        elif hoogte >= 1001:
            # Hoge kast (leeg) 1001–2079
            corpus_rij = 12

    elif kast_type == "C":
        # C = hangkasten, verschillende hoogtes
        if hoogte <= 390:
            corpus_rij = 15  # geen planken inbegrepen
        elif 391 <= hoogte <= 520:
            corpus_rij = 16  # 1 plank inbegrepen
        elif 521 <= hoogte <= 780:
            corpus_rij = 17  # 2 planken inbegrepen
        else:  # >= 781
            corpus_rij = 18  # 2 planken inbegrepen

    corpus_prijs = 0.0
    if corpus_rij:
        corpus_prijs = _prijs_uit_tabel(corpus_rij, breedte)

    # ---------- 2. INRICHTING – PLANKEN EN LADES ----------
    plank_count, lade_count = _parse_inrichting_counts(inrichting_tekst)

    # inbegrepen planken per type/hoogte
    inbegrepen_planken = 0
    if kast_type == "A":
        # Alleen A-kast geschikt voor planken heeft 2 inbegrepen planken
        if corpus_rij == 8:
            inbegrepen_planken = 2
    elif kast_type == "C":
        if hoogte <= 390:
            inbegrepen_planken = 0
        elif 391 <= hoogte <= 520:
            inbegrepen_planken = 1
        elif hoogte <= 780:
            inbegrepen_planken = 2
        else:
            inbegrepen_planken = 2

    extra_planken = max(0, plank_count - inbegrepen_planken)

    # Plankprijs (B22)
    plank_prijs_per_stuk = _prijs_uit_tabel(22, breedte) if extra_planken > 0 else 0.0
    planken_prijs = extra_planken * plank_prijs_per_stuk

    # Ladeprijs (B23)
    lade_prijs_per_stuk = _prijs_uit_tabel(23, breedte) if lade_count > 0 else 0.0
    lades_prijs = lade_count * lade_prijs_per_stuk

    # ---------- 3. SCHARNIEREN (B25) ----------
    scharnier_prijs_per_stuk = _prijs_uit_tabel(25, breedte) if scharnieren_aantal > 0 else 0.0
    scharnieren_prijs = scharnieren_aantal * scharnier_prijs_per_stuk

    # ---------- 4. FRONTEN + ZIJKANTEN (m²) + 40% OPSLAG ----------
    # Front-oppervlakte (totale frontvlak)
    front_hoogte_m = hoogte / 1000.0
    front_breedte_m = breedte / 1000.0
    front_m2 = front_hoogte_m * front_breedte_m

    # m²-prijs front
    m2_prijs_front = FRONT_M2_PRICES.get(frontmodel, 0.0)
    front_m2_prijs = front_m2 * m2_prijs_front

    # Zichtbare zijden
    zij_m2_totaal = 0.0
    z = zichtbare_zijde
    # we werken met hoogte x diepte voor een zijvlak
    zij_hoogte_m = front_hoogte_m
    zij_diepte_m = diepte / 1000.0
    zij_m2 = zij_hoogte_m * zij_diepte_m

    if "links" in z and "rechts" in z:
        zij_m2_totaal = 2 * zij_m2
    elif "links" in z or "rechts" in z:
        zij_m2_totaal = zij_m2

    # m²-prijs voor vlakke zijkant op basis van materiaal-groep
    materiaal_groep = _bepaal_materiaal_groep(frontmodel)
    vlak_model = VLAK_MODEL_PER_MATERIAALGROEP.get(materiaal_groep, "NOAH")
    vlak_m2_prijs = FRONT_M2_PRICES.get(vlak_model, 0.0)

    zij_prijs = zij_m2_totaal * vlak_m2_prijs

    tot_m2_prijs = front_m2_prijs + zij_prijs

    # 40% opslag voor montage van fronten op de kasten
    tot_m2_prijs_met_opslag = tot_m2_prijs * 1.40

    # ---------- 5. TOTAAL PER KAST ----------
    totaal_kast = corpus_prijs + planken_prijs + lades_prijs + scharnieren_prijs + tot_m2_prijs_met_opslag

    return round(totaal_kast, 2)


def _bereken_maatwerk_kasten(maatwerk_kasten_lijst):
    """
    Loopt door alle kasten (kolommen B t/m K van tabblad 2) en
    berekent per kast de prijs + omschrijving voor Teamleader.
    """
    resultaten = []
    totaal = 0.0

    for kast in maatwerk_kasten_lijst:
        prijs = _bereken_maatwerk_kastprijs(kast)
        totaal += prijs

        resultaten.append({
            "titel": _kast_titel(kast.get("type_kast")),
            "beschrijving": _beschrijving_uit_kastdict(kast),
            "prijs_excl": prijs,
        })

    return resultaten, round(totaal, 2)


# ======================================================
# 📥 EXCEL UITLEZEN
# ======================================================

def lees_excel(path):
    """
    Leest zowel:
      - tabblad 1 (renovatie-fronten)
      - tabblad 2 (maatwerk kasten)
    en retourneert ALLES in één keer.
    """
    # Tabblad 1 = bestaande logica
    df = pd.read_excel(path, sheet_name=0, header=None)

    onderdelen = df.iloc[:, 5].dropna().astype(str).str.upper().tolist()

    g2 = df.iloc[1, 6]
    h2 = df.iloc[1, 7]
    kleur = df.iloc[1, 8]

    klantregels = [
        str(df.iloc[r, 10])
        for r in range(1, 6)
        if pd.notna(df.iloc[r, 10])
    ]

    scharnieren = int(df.iloc[2, 9]) if pd.notna(df.iloc[2, 9]) else 0
    lades = int(df.iloc[4, 9]) if pd.notna(df.iloc[4, 9]) else 0

    project = os.path.splitext(os.path.basename(path))[0]

    # Tabblad 2 = maatwerk kasten
    maatwerk_kasten = []
    try:
        df2 = pd.read_excel(path, sheet_name=1, header=None)
        maatwerk_kasten = _lees_maatwerk_kasten_tab2(df2)
    except Exception:
        # Als tabblad 2 ontbreekt, gewoon geen maatwerk kasten
        maatwerk_kasten = []

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project, maatwerk_kasten


def _lees_maatwerk_kasten_tab2(df2):
    """
    Leest tabblad 2 structuur:
    Kolom A: labels (Hoogte, Breedte, Diepte, etc.)
    Kolommen B t/m K: per kolom één kast.
    """
    # Map labels naar interne veldnamen
    label_to_field = {
        "type kast": "type_kast",
        "hoogte": "hoogte",
        "breedte": "breedte",
        "diepte": "diepte",
        "hoogte pootje": "hoogte_pootje",
        "kleur corpus": "kleur_corpus",
        "zichtbare zijde": "zichtbare_zijde",
        "inrichting": "inrichting",
        "scharnieren": "scharnieren",
        "front": "frontmodel",
        "front indeling": "front_indeling",
        "kleur": "kleur_front",
        "dubbelzijdig afgewerkt": "dubbelzijdig",
        "handgreep": "handgreep",
        "afwerking": "afwerking",
        "aantal fronten": "aantal_fronten",
    }

    # Zoek per label de rij-index
    label_rows = {}
    for r in range(df2.shape[0]):
        val = str(df2.iloc[r, 0]).strip().lower().rstrip(":")
        if val in label_to_field:
            label_rows[val] = r

    # Als er geen 'type kast' label is, kunnen we niets doen
    if "type kast" not in label_rows:
        return []

    kasten = []
    # Kolommen 1.. tot einde (B t/m K)
    for col in range(1, df2.shape[1]):
        type_row = label_rows["type kast"]
        type_val = df2.iloc[type_row, col]
        if pd.isna(type_val) or str(type_val).strip() == "":
            continue  # geen kast in deze kolom

        kast_data = {}
        for label, field in label_to_field.items():
            r = label_rows.get(label)
            if r is None:
                continue
            value = df2.iloc[r, col] if r < df2.shape[0] else None
            if pd.isna(value):
                value = None
            kast_data[field] = value

        kasten.append(kast_data)

    return kasten


# ======================================================
# 🧮 OFFERTE BEREKENING (RENOVATIE + MAATWERK KASTEN)
# ======================================================

def bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades, maatwerk_kasten):
    """
    Bestaande renovatie-berekening + aanvulling met maatwerk kasten.
    """
    info = MODEL_INFO[model]

    fronts = sum(o in ["DEUR", "LADE", "BEDEKKINGSPANEEL"] for o in onderdelen)
    heeft_passtuk = any(o in ["PASSTUK", "PLINT"] for o in onderdelen)
    heeft_anders = any("ANDERS" in o for o in onderdelen)

    passtuk_kosten = info["passtuk"] if heeft_passtuk else 0
    anders_kosten = info["passtuk"] if heeft_anders else 0

    materiaal_totaal = fronts * info["prijs_per_front"]
    montage = fronts * MONTAGE_PER_FRONT
    scharnier_totaal = scharnieren * PRIJS_SCHARNIER
    lades_totaal = lades * PRIJS_LADE

    totaal_excl = (
        materiaal_totaal +
        passtuk_kosten +
        anders_kosten +
        montage +
        INMETEN +
        VRACHT +
        scharnier_totaal +
        lades_totaal
    )

    # ===== MAATWERK KASTEN BEREKENING =====
    maatwerk_regels, maatwerk_totaal_excl = _bereken_maatwerk_kasten(maatwerk_kasten or [])

    totaal_excl += maatwerk_totaal_excl

    btw = totaal_excl * 0.21
    totaal_incl = totaal_excl + btw

    return {
        "project": project,
        "model": model,
        "kleur": kleur,
        "materiaal": info["materiaal"],
        "klantgegevens": klantregels,
        "fronts": fronts,
        "toeslag_passtuk": passtuk_kosten,
        "toeslag_anders": anders_kosten,
        "prijs_per_front": info["prijs_per_front"],
        "scharnieren": scharnieren,
        "scharnier_totaal": scharnier_totaal,
        "lades": lades,
        "lades_totaal": lades_totaal,
        "materiaal_totaal": materiaal_totaal,
        "montage": montage,
        "totaal_excl": totaal_excl,
        "btw": btw,
        "totaal_incl": totaal_incl,

        # Maatwerk kasten
        "maatwerk_kasten": maatwerk_regels,
        "maatwerk_kasten_totaal_excl": maatwerk_totaal_excl,
    }


# ======================================================
# 🧾 TEAMLEADER OFFERTE AANMAKEN
# ======================================================

def maak_teamleader_offerte(deal_id, data, mode):

    url = f"{API_BASE}/quotations.create"

    model = data["model"]
    cfg = FRONT_DESCRIPTION_CONFIG.get(model)
    fronts = data["fronts"]

    klantregels = data["klantgegevens"] + ["", "", "", "", ""]
    klanttekst = (
        f"Naam: {klantregels[0]}\r\n"
        f"Adres: {klantregels[1]}\r\n"
        f"Postcode / woonplaats: {klantregels[2]}\r\n"
        f"Tel: {klantregels[3]}\r\n"
        f"Email: {klantregels[4]}"
    )

    grouped_lines = []

    # -----------------------------  
    # KLANTGEGEVENS  
    # -----------------------------  

    grouped_lines.append({
        "section": {"title": "KLANTGEGEVENS"},
        "line_items": [{
            "quantity": 1,
            "description": "Klantgegevens",
            "extended_description": klanttekst,
            "unit_price": {"amount": 0, "tax": "excluding"},
            "tax_rate_id": TAX_RATE_21_ID,
        }],
    })

    # -----------------------------  
    # PARTICULIER  
    # -----------------------------  

    if mode == "P":
        tekst = []
        toevoegingen = []

        if data["toeslag_passtuk"] > 0:
            toevoegingen.append("inclusief passtukken en/of plinten")

        if data["toeslag_anders"] > 0:
            toevoegingen.append("inclusief licht- en/of sierlijsten")

        extra_text = f" ({', '.join(toevoegingen)})" if toevoegingen else ""

        tekst += [
            f"Aantal fronten: {fronts} fronten{extra_text}",
            f"Materiaal: {cfg['materiaal']}",
            f"Frontdikte: {cfg['frontdikte']}",
            f"Kleur: {data['kleur']}",
            f"Afwerking: {cfg['afwerking']}",
            f"Dubbelzijdig in kleur afwerken: {cfg['dubbelzijdig']}",
            "Inmeten: Ja",
            "Montage: Ja",
            "Handgrepen: Te bepalen",
            "",
            "Prijs is inclusief:",
            "- Demontage oude fronten & materialen",
            "- Inmeten, leveren en montage van de fronten",
        ]

        if data["toeslag_passtuk"] > 0:
            tekst.append("- Montage van passtukken en/of plinten")
        if data["toeslag_anders"] > 0:
            tekst.append("- Montage van licht- en/of sierlijsten")

        tekst.append("- Afvoeren van oude fronten")

        if data["scharnieren"] > 0:
            tekst.append(f"- Inclusief vervangen scharnieren ({data['scharnieren']} stuks)")
        if data["lades"] > 0:
            tekst.append(f"- Inclusief plaatsen maatwerk lades ({data['lades']} stuks)")

        final = "\r\n".join(tekst)

        grouped_lines.append({
            "section": {"title": "KEUKENRENOVATIE"},
            "line_items": [{
                "quantity": 1,
                "description": f"Keukenrenovatie model {model}",
                "extended_description": final,
                "unit_price": {"amount": round(data["totaal_excl"], 2), "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            }],
        })

    # -----------------------------  
    # DEALER  
    # -----------------------------  

    else:
        section = {
            "section": {"title": "KEUKENRENOVATIE"},
            "line_items": []
        }

        section["line_items"].append({
            "quantity": fronts,
            "description": cfg["titel"],
            "extended_description": (
                f"Aantal fronten: {fronts}\r\n"
                f"Materiaal: {cfg['materiaal']}\r\n"
                f"Frontdikte: {cfg['frontdikte']}\r\n"
                f"Kleur: {data['kleur']}\r\n"
                f"Afwerking: {cfg['afwerking']}\r\n"
                f"Dubbelzijdig in kleur afwerken: {cfg['dubbelzijdig']}\r\n"
                "Inmeten: Ja\r\n"
                "Montage: Ja\r\n"
                "\r\n"
                "Fronten worden geleverd zonder scharnieren"
            ),
            "unit_price": {"amount": data["prijs_per_front"], "tax": "excluding"},
            "tax_rate_id": TAX_RATE_21_ID,
        })

        if data["toeslag_passtuk"] > 0:
            section["line_items"].append({
                "quantity": 1,
                "description": "Plinten en/of passtukken",
                "extended_description": "inclusief montage",
                "unit_price": {"amount": data["toeslag_passtuk"], "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            })

        if data["toeslag_anders"] > 0:
            section["line_items"].append({
                "quantity": 1,
                "description": "Licht- en/of sierlijsten",
                "extended_description": "inclusief montage",
                "unit_price": {"amount": data["toeslag_anders"], "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            })

        grouped_lines.append(section)

        # ACCESSOIRES
        if data["scharnieren"] > 0 or data["lades"] > 0:
            acc = {"section": {"title": "ACCESSOIRES"}, "line_items": []}

            if data["scharnieren"] > 0:
                acc["line_items"].append({
                    "quantity": data["scharnieren"],
                    "description": "Scharnieren - Softclose",
                    "extended_description": "Prijs per stuk",
                    "unit_price": {"amount": PRIJS_SCHARNIER, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                })

            if data["lades"] > 0:
                acc["line_items"].append({
                    "quantity": data["lades"],
                    "description": "Maatwerk lades - Softclose",
                    "extended_description": "Prijs per stuk (incl. montage)",
                    "unit_price": {"amount": PRIJS_LADE, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                })

            grouped_lines.append(acc)

        # INMETEN EN MONTAGE
        grouped_lines.append({
            "section": {"title": "INMETEN, LEVEREN & MONTEREN"},
            "line_items": [
                {
                    "quantity": 1,
                    "description": "Inmeten",
                    "extended_description": "Inmeten op locatie",
                    "unit_price": {"amount": INMETEN, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
                {
                    "quantity": fronts,
                    "description": "Montage per front",
                    "extended_description": "Inclusief demontage oude fronten & afvoeren",
                    "unit_price": {"amount": MONTAGE_PER_FRONT, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
                {
                    "quantity": 1,
                    "description": "Vracht- & verpakkingskosten",
                    "extended_description": "Levering op locatie",
                    "unit_price": {"amount": VRACHT, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
            ],
        })

    # -----------------------------
    # MAATWERK KASTEN – NIEUWE SECTIE
    # -----------------------------
    maatwerk_regels = data.get("maatwerk_kasten", [])
    if maatwerk_regels:
        section_mk = {
            "section": {"title": "MAATWERK KASTEN"},
            "line_items": []
        }

        for kast in maatwerk_regels:
            section_mk["line_items"].append({
                "quantity": 1,
                "description": kast["titel"],
                "extended_description": kast["beschrijving"],
                "unit_price": {"amount": kast["prijs_excl"], "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            })

        grouped_lines.append(section_mk)

    payload = {
        "deal_id": deal_id,
        "currency": {"code": "EUR", "exchange_rate": 1.0},
        "grouped_lines": grouped_lines,
        "text": "\u200b"
    }

    resp = request_with_auto_refresh("POST", url, json_data=payload)

    if resp.status_code not in (200, 201):
        raise Exception(f"Offerte NIET aangemaakt: {resp.text}")

    return True
