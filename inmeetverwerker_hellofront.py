import pandas as pd
import requests
import os
import json
import re

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

PRIJS_SCHARNIER = 6.5
PRIJS_LADE = 184.0

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
# 🧮 MODEL- EN PRIJSLOGICA FRONTEN (BESTAAND)
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
# 🧮 MAATWERK KASTEN – PRIJSENGINE
# ======================================================

# m2 prijzen voor fronten (inkoop), op basis van jouw lijst
MAATWERK_FRONT_M2_PRIJZEN = {
    "NOAH": 76.00,
    "FEDDE": 80.00,
    "DEX": 103.00,
    "DAVE": 103.00,
    "JOLIE": 103.00,
    "JACK": 114.50,
    "CHIEL": 149.50,
    "JAMES": 202.50,
    "SAM": 141.50,
    "DUKE": 186.50,
}

# vlak model per materiaal (voor zichtbare zijden)
# MDF gespoten → NOAH, Eiken → JACK, Noten → SAM
MATERIAAL_NAAR_VLAKMODEL = {
    "MDF gespoten": "NOAH",
    "Eikenfineer": "JACK",
    "Noten fineer": "SAM",
}

# Breedte-staffels zoals in je prijslijst (C3 t/m J3)
MAATWERK_BREEDTE_STAFFELS = [300, 400, 500, 600, 800, 900, 1000, 1200]

def _staffel_index_omhoog(maat_mm: float) -> int:
    """
    Zoek index van eerstvolgende staffel naar boven.
    Voorbeeld: 610 mm → staffel 800 mm.
    """
    for idx, grens in enumerate(MAATWERK_BREEDTE_STAFFELS):
        if maat_mm <= grens:
            return idx
    # Als groter dan grootste staffel → laatste staffel
    return len(MAATWERK_BREEDTE_STAFFELS) - 1

# Corpusprijzen uit jouw FinanceKing-lijst (C7..J18)
# Let op: sommige combinaties (bijv. 1200mm bij hoge kasten) zijn in jouw txt niet volledig zichtbaar.
# Voor de veiligheid zijn die hier niet ingevuld; als je ze nodig hebt kun je ze zelf aanvullen.
MAATWERK_CORPUS_PRIJZEN = {
    # A-kasten
    # Aanname op basis van eerdere controles met jou:
    # - Rij 7: onderkast geschikt voor lades
    # - Rij 8: onderkast geschikt voor ovens
    "A_LADE":       [45, 47, 49, 51, 59, 61, 63, 66],  # C7..J7
    "A_OVEN":       [51, 55, 60, 64, 77, 81, 85, 92],  # C8..J8

    # B-kasten (hoge kasten)
    # Rij 11: hoge kast (leeg) 2080–2770mm
    # Rij 12: hoge kast (leeg) 1001–2079mm
    "B_HOOG_2080_2770": [85, 89, 92, 97, 103, 108, 114, None],  # C11..I11, J11 ontbreekt in txt → None
    "B_HOOG_1001_2079": [80, 84, 87, 90, 95, 98, 103, None],   # C12..I12, J12 ontbreekt → None

    # C-kasten (hangkasten) – corpusprijzen per hoogtegroep
    # Rij 15: t/m 390mm
    # Rij 16: 391–520mm
    # Rij 17: 521–780mm
    # Rij 18: vanaf 781mm
    "C_TOT_390":    [35, 37, 39, 41, 47, 48, 50, 56],  # C15..J15
    "C_391_520":    [39, 42, 45, 47, 56, 58, 61, 67],  # C16..J16
    "C_521_780":    [42, 45, 49, 50, 57, 61, 66, 70],  # C17..J17
    "C_VANAF_781":  [46, 48, 51, 55, 61, 65, 68, 75],  # C18..J18
}

# Inrichtingen (rij 21 t/m 29)
MAATWERK_INRICHTING_PRIJZEN = {
    "KLEPSCHARNIER": {
        "per_stuk": 63.0  # B21 = Klepscharnieren, prijs is 63 in alle kolommen
    },
    "PLANK": {
        # rij 22, C22..J22
        "per_staffel": [4.5, 6.3, 7.5, 7.6, 10.6, 11.8, 13.3, 15.75],
    },
    "LADE": {
        # rij 23: 68 in alle kolommen
        "per_stuk": 68.0
    },
    "PUSH_TO_OPEN_LADE": {
        # rij 24: 109 in alle kolommen
        "per_stuk": 109.0
    },
    "SCHARNIER_PER_STUK": {
        # rij 25: 9.7 in alle kolommen
        "per_stuk": 9.7
    },
    "BESTEKBAK": {
        # rij 26: 33 in alle kolommen
        "per_stuk": 33.0
    },
    "PLASTIC_SPOELKAST": {
        # rij 27: 33 in alle kolommen
        "per_stuk": 33.0
    },
    "APOTHEKERSLADE": {
        # rij 28: 546 in alle kolommen
        "per_stuk": 546.0
    },
    "CARROUSEL": {
        # rij 29: 452 voor G..J, overige kolommen niet volledig zichtbaar in txt
        "per_staffel": [452.0, 452.0, 452.0, 452.0, 452.0, 452.0, 452.0, 452.0],
    },
}


def _parse_inrichting(inrichting: str):
    """
    Haalt aantallen uit de inrichting-string.
    Voorbeelden:
    - '2x plank, 1x lade'
    - '3 lades, 1 plank'
    Geeft dict terug met aantallen: {"PLANK": 2, "LADE": 1, ...}
    """
    result = {
        "PLANK": 0,
        "LADE": 0,
        "KLEPSCHARNIER": 0,
        "APOTHEKERSLADE": 0,
        "CARROUSEL": 0,
    }
    if not inrichting:
        return result

    s = str(inrichting).lower()

    # generieke regex: "<aantal> x <type>"
    matches = re.findall(r"(\d+)\s*x?\s*([a-zA-Z]+)", s)
    for aantal_str, soort in matches:
        aantal = int(aantal_str)
        soort = soort.lower()

        if "plank" in soort:
            result["PLANK"] += aantal
        elif "lade" in soort and "apothek" not in soort and "push" not in soort:
            result["LADE"] += aantal
        elif "klep" in soort:
            result["KLEPSCHARNIER"] += aantal
        elif "apothek" in soort:
            result["APOTHEKERSLADE"] += aantal
        elif "carrousel" in soort or "carousel" in soort:
            result["CARROUSEL"] += aantal

    return result


def _bepaal_hoofdtype(type_kast: str, hoogte_mm: float):
    """
    Bepaalt A / B / C hoofdcategorie op basis van type_kast en hoogte.
    """
    t = (type_kast or "").strip().upper()
    if t in ("A", "B", "C"):
        return t

    # Fallback op hoogte (alleen als type niet gevuld is)
    if hoogte_mm <= 1000:
        return "A"
    # B-kasten tot 2770mm
    if hoogte_mm <= 2770:
        return "B"
    # Anders hangen we 'm aan C, maar dat komt in jouw praktijk niet voor
    return "C"


def _bepaal_subtype_code(hoofdtype: str, hoogte_mm: float, inrichting_str: str):
    """
    Bepaalt welke corpusprijs-reeks gebruikt moet worden binnen A/B/C.
    """
    inrichting_info = _parse_inrichting(inrichting_str)

    if hoofdtype == "A":
        # A-kast: tot 1000mm hoog, 3 varianten:
        # - kast geschikt voor ovens
        # - kast geschikt voor lades
        # - kast geschikt voor planken (incl. 2 planken)
        if inrichting_info["LADE"] > 0:
            return "A_LADE"
        elif inrichting_info["PLANK"] > 0:
            # 'geschikt voor planken' variant
            return "A_LADE"  # corpus is technisch meestal hetzelfde; als je een aparte rij hebt kun je die hier mappen
        else:
            # Geen lades/planken genoemd → behandel als ovenkast
            return "A_OVEN"

    if hoofdtype == "B":
        # Hoge kast (leeg) 1001–2079 of 2080–2770
        if hoogte_mm >= 2080:
            return "B_HOOG_2080_2770"
        else:
            return "B_HOOG_1001_2079"

    if hoofdtype == "C":
        # Hangkasten: op basis van hoogte-groepen
        if hoogte_mm <= 390:
            return "C_TOT_390"
        if 391 <= hoogte_mm <= 520:
            return "C_391_520"
        if 521 <= hoogte_mm <= 780:
            return "C_521_780"
        # 781 en hoger
        return "C_VANAF_781"

    return None


def _inbegrepen_planken_c_kast(hoogte_mm: float) -> int:
    """
    Inbegrepen planken bij C-kasten:
    - t/m 390mm → 0
    - 391–520mm → 1
    - 521–780mm → 2
    - vanaf 781mm → 2
    """
    if hoogte_mm <= 390:
        return 0
    if 391 <= hoogte_mm <= 520:
        return 1
    if 521 <= hoogte_mm <= 780:
        return 2
    return 2


def _inbegrepen_planken_a_kast(inrichting_str: str) -> int:
    """
    Voor A-kast 'geschikt voor planken' zijn altijd 2 planken inbegrepen.
    Voor andere A-varianten → 0 inbegrepen planken.
    """
    if "plank" in (inrichting_str or "").lower():
        return 2
    return 0


def bereken_maatwerk_kast(
    type_kast: str,
    hoogte_mm: float,
    breedte_mm: float,
    diepte_mm: float,
    hoogte_poot_mm: float,
    zichtbare_zijde: str,
    inrichting: str,
    scharnieren_stuks: int,
    frontmodel: str,
    aantal_fronten: int,
    kleur_corpus: str,
    dubbelzijdig_afgewerkt: str,
    handgreep: str,
    afwerking: str,
):
    """
    Berekent de inkoopprijs van één maatwerk kast, inclusief:
    - corpus
    - inrichting
    - scharnieren
    - fronten (m2)
    - zichtbare zijkanten (m2 in vlak model)
    - 40% opslag voor montage fronten op kasten
    Geeft een dict terug met prijs + gegevens voor Teamleader.
    """

    frontmodel = (frontmodel or "").strip().upper()
    hoofdtype = _bepaal_hoofdtype(type_kast, hoogte_mm)
    subtype_code = _bepaal_subtype_code(hoofdtype, hoogte_mm, inrichting)

    # --------------------------
    # Corpusprijs (A/B/C + staffel)
    # --------------------------
    breedte_index = _staffel_index_omhoog(breedte_mm)
    corpus_prijs = 0.0
    if subtype_code and subtype_code in MAATWERK_CORPUS_PRIJZEN:
        corpus_reeks = MAATWERK_CORPUS_PRIJZEN[subtype_code]
        corpus_waarde = corpus_reeks[breedte_index]
        if corpus_waarde is None:
            raise Exception(
                f"Geen corpusprijs gedefinieerd voor subtype {subtype_code} bij breedte-staffel {MAATWERK_BREEDTE_STAFFELS[breedte_index]} mm"
            )
        corpus_prijs = float(corpus_waarde)

    # --------------------------
    # Inrichting (planken, lades, e.d.)
    # --------------------------
    inrichting_info = _parse_inrichting(inrichting)

    # Planken
    plank_eenheden = inrichting_info["PLANK"]
    inbegrepen_planken = 0
    if hoofdtype == "C":
        inbegrepen_planken = _inbegrepen_planken_c_kast(hoogte_mm)
    elif hoofdtype == "A":
        inbegrepen_planken = _inbegrepen_planken_a_kast(inrichting)

    extra_planken = max(0, plank_eenheden - inbegrepen_planken)
    plank_prijs_per_staffel = MAATWERK_INRICHTING_PRIJZEN["PLANK"]["per_staffel"][breedte_index]
    planken_prijs = extra_planken * plank_prijs_per_staffel

    # Lades (maatwerk korpus-lades)
    lade_stuks = inrichting_info["LADE"]
    lade_prijs_per_stuk = MAATWERK_INRICHTING_PRIJZEN["LADE"]["per_stuk"]
    lades_prijs = lade_stuks * lade_prijs_per_stuk

    # Eventuele andere inrichtingen kun je op dezelfde manier toevoegen
    # Klepscharnier, apothekerslade, carrousel etc. worden hier (voor nu) niet apart doorgerekend.

    # --------------------------
    # Scharnieren (maatwerk kasten)
    # --------------------------
    scharnier_prijs_per_stuk = MAATWERK_INRICHTING_PRIJZEN["SCHARNIER_PER_STUK"]["per_stuk"]
    scharnieren_prijs = scharnieren_stuks * scharnier_prijs_per_stuk

    # --------------------------
    # Fronten + zichtbare zijkant(en) in m2
    # --------------------------
    # totale front-oppervlakte (ongeacht aantal fronten; verdeling doet er m2-technisch niet toe)
    front_m2 = (hoogte_mm * breedte_mm) / 1_000_000.0

    if frontmodel not in MAATWERK_FRONT_M2_PRIJZEN:
        raise Exception(f"Onbekend frontmodel voor m2-prijs: {frontmodel}")

    front_m2_prijs = MAATWERK_FRONT_M2_PRIJZEN[frontmodel]
    fronten_prijs = front_m2 * front_m2_prijs

    # Zichtbare zijden
    zicht = (zichtbare_zijde or "").lower()
    aantal_zijkanten = 0
    if "links" in zicht and "rechts" in zicht:
        aantal_zijkanten = 2
    elif "beide" in zicht:
        aantal_zijkanten = 2
    elif "links" in zicht or "rechts" in zicht:
        aantal_zijkanten = 1
    else:
        aantal_zijkanten = 0

    zijkant_m2 = 0.0
    zijkanten_prijs = 0.0
    if aantal_zijkanten > 0:
        # zijde-oppervlakte: hoogte x diepte
        zijde_m2 = (hoogte_mm * diepte_mm) / 1_000_000.0
        zijkant_m2 = zijde_m2 * aantal_zijkanten

        # materiaal bepalen via MODEL_INFO (MDF / Eiken / Noten)
        materiaal = MODEL_INFO.get(frontmodel, {}).get("materiaal")
        vlakmodel = MATERIAAL_NAAR_VLAKMODEL.get(materiaal)
        if not vlakmodel:
            raise Exception(f"Geen vlakmodel gedefinieerd voor materiaal: {materiaal}")

        vlak_m2_prijs = MAATWERK_FRONT_M2_PRIJZEN[vlakmodel]
        zijkanten_prijs = zijkant_m2 * vlak_m2_prijs

    # --------------------------
    # 40% opslag voor montage fronten op kasten
    # --------------------------
    totaal_front_en_zijkant = fronten_prijs + zijkanten_prijs
    front_en_zijkant_met_opslag = totaal_front_en_zijkant * 1.40  # +40%

    # --------------------------
    # Totaal inkoopprijs maatwerk kast
    # --------------------------
    totaal_inkoop = (
        corpus_prijs +
        planken_prijs +
        lades_prijs +
        scharnieren_prijs +
        front_en_zijkant_met_opslag
    )

    # Titel op basis van hoofdtype
    if hoofdtype == "A":
        titel = "Maatwerk onderkast"
    elif hoofdtype == "B":
        titel = "Maatwerk hoge kast"
    else:
        titel = "Maatwerk hangkast"

    # Beschrijving EXACT zoals ingevuld in Excel-opzet
    beschrijving_regels = [
        f"Type kast: {type_kast}",
        f"Hoogte: {hoogte_mm} mm",
        f"Breedte: {breedte_mm} mm",
        f"Diepte: {diepte_mm} mm",
        f"Hoogte pootje: {hoogte_poot_mm} mm",
        f"Zichtbare zijde: {zichtbare_zijde}",
        f"Inrichting: {inrichting}",
        f"Scharnieren: {scharnieren_stuks}",
        f"Frontmodel: {frontmodel}",
        f"Aantal fronten: {aantal_fronten}",
        f"Kleur corpus: {kleur_corpus}",
        f"Dubbelzijdig afgewerkt: {dubbelzijdig_afgewerkt}",
        f"Handgreep: {handgreep}",
        f"Afwerking: {afwerking}",
    ]
    beschrijving = "\r\n".join(beschrijving_regels)

    return {
        "hoofdtype": hoofdtype,
        "titel": titel,
        "beschrijving": beschrijving,
        "totaal_inkoop_excl": round(totaal_inkoop, 2),
        "corpus_prijs": round(corpus_prijs, 2),
        "planken_prijs": round(planken_prijs, 2),
        "lades_prijs": round(lades_prijs, 2),
        "scharnieren_prijs": round(scharnieren_prijs, 2),
        "fronten_prijs_m2": round(fronten_prijs, 2),
        "zijkanten_prijs_m2": round(zijkanten_prijs, 2),
        "front_en_zijkant_met_opslag": round(front_en_zijkant_met_opslag, 2),
    }


# ======================================================
# 📥 EXCEL UITLEZEN (BESTAAND – NOG ZONDER TABBLAD 2)
# ======================================================

def lees_excel(path):
    df = pd.read_excel(path, header=None)

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

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project


# ======================================================
# 🧮 OFFERTE BEREKENING (BESTAAND, NU MET OPTIONELE MAATWERK-KASTEN)
# ======================================================

def bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades, maatwerk_kasten=None):
    """
    Bestaande offerte-berekening voor keukenrenovatie-fronten.
    maatwerk_kasten: optionele lijst met reeds berekende kastdicts
                     (zoals uit bereken_maatwerk_kast), wordt alleen
                     in data meegestuurd zodat maak_teamleader_offerte
                     er secties van kan maken.
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
        # nieuwe sleutel: optionele lijst met maatwerk-kasten
        "maatwerk_kasten": maatwerk_kasten or [],
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
    # MAATWERK KASTEN (NIEUW – OPTIONEEL)
    # -----------------------------
    maatwerk_kasten = data.get("maatwerk_kasten") or []
    if maatwerk_kasten:
        sectie_kasten = {
            "section": {"title": "MAATWERK KASTEN"},
            "line_items": []
        }

        for kast in maatwerk_kasten:
            sectie_kasten["line_items"].append({
                "quantity": 1,
                "description": kast["titel"],
                "extended_description": kast["beschrijving"],
                "unit_price": {"amount": kast["totaal_inkoop_excl"], "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            })

        grouped_lines.append(sectie_kasten)

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
