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
# 🧮 MODEL- EN PRIJSLOGICA (FRONTEN)
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

# ======================================================
# 🧮 M²-PRIJZEN VOOR MAATWERK FRONTEN
# ======================================================

M2_FRONT_PRIJZEN = {
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

# vlak model per materiaaltype voor ZIJKANTEN
VLAK_MODEL_PER_MATERIAAL = {
    "MDF gespoten": "NOAH",
    "Eikenfineer": "JACK",
    "Noten fineer": "SAM",
}

# ======================================================
# 🧮 PRIJS-TABELLEN MAATWERK KASTEN (CORPUS + INRICHTING)
# ======================================================

BREEDTE_STAFFELS = [300, 400, 500, 600, 800, 900, 1000, 1200]

# A-kast geschikt voor lades (rij 7 C–J)
A_LADE_KAST = [45, 47, 49, 51, 59, 61, 63, 66]

# Onderkast geschikt voor ovens (rij 8 C–J)
A_OVEN_KAST = [51, 55, 60, 64, 77, 81, 85, 92]

# B-kasten (hoge kasten – leeg)
# 1001–2079 mm (rij 12 C–J)
B_HOOG_1001_2079 = [80, 84, 87, 90, 95, 98, 103, 103]

# 2080–2770 mm (rij 11 C–J)
B_HOOG_2080_2770 = [85, 89, 92, 97, 103, 108, 114, 114]

# C-kasten (hangkasten) – corpus
# t/m 390mm (rij 15 C–J)
C_CORPUS_0_390 = [35, 37, 39, 41, 47, 48, 50, 56]
# 391–520mm (rij 16 C–J)
C_CORPUS_391_520 = [39, 42, 45, 47, 56, 58, 61, 67]
# 521–780mm (rij 17 C–J)
C_CORPUS_521_780 = [42, 45, 49, 50, 57, 61, 66, 70]
# vanaf 781mm (rij 18 C–J)
C_CORPUS_781_PLUS = [46, 48, 51, 55, 61, 65, 68, 75]

# Inrichting / accessoires (C–J)
PLANK_A_OF_B = [4.5, 6.3, 7.5, 7.6, 10.6, 11.8, 13.3, 15.75]  # rij 22
LADES_KAST = [68, 68, 68, 68, 68, 68, 68, 68]  # rij 23
PUSH_TO_OPEN_LADE = [109, 109, 109, 109, 109, 109, 109, 109]  # rij 24
SCHARNIER_PER_STUK_MAATWERK = [9.7, 9.7, 9.7, 9.7, 9.7, 9.7, 9.7, 9.7]  # rij 25
BESTEK_BAK = [33, 33, 33, 33, 33, 33, 33, 33]  # rij 26
SPOELKAST_BESCHERMING = [33, 33, 33, 33, 33, 33, 33, 33]  # rij 27
APOTHEKERS_LADE = [546, 546, 546, 546, 546, 546, 546, 546]  # rij 28
CARROUSEL = [0, 0, 0, 452, 452, 452, 452, 452]  # G–J=452, C–F=0


# ======================================================
# 🔎 HULPFUNCTIES MAATWERK KASTEN
# ======================================================

def _staffel_index(breedte_mm: float) -> int:
    """
    Bepaal index in BREEDTE_STAFFELS volgens jouw regel:
    altijd naar BOVEN afronden naar de eerstvolgende staffel.
    """
    if breedte_mm is None or math.isnan(breedte_mm):
        return 0
    for i, grens in enumerate(BREEDTE_STAFFELS):
        if breedte_mm <= grens:
            return i
    return len(BREEDTE_STAFFELS) - 1


def _safe_float(val):
    try:
        if pd.isna(val):
            return None
        return float(str(val).replace(",", "."))
    except Exception:
        return None


def _safe_int(val):
    try:
        if pd.isna(val) or val == "":
            return 0
        return int(float(str(val).replace(",", ".")))
    except Exception:
        return 0


def _parse_inrichting(inrichting_raw: str):
    """
    Leest een tekst zoals '3x plank, 1x lade' en geeft aantallen terug.
    """
    result = {
        "planken": 0,
        "lades": 0,
        "push_to_open_lades": 0,
        "bestek_bakken": 0,
        "spoelkast_bescherming": 0,
        "apothekers": 0,
        "carrousels": 0,
        "klepscharnieren": 0,
    }

    if not isinstance(inrichting_raw, str):
        return result

    txt = inrichting_raw.lower()
    parts = [p.strip() for p in txt.split(",") if p.strip()]

    for part in parts:
        # hoeveelheid
        aantal = 1
        for token in part.split():
            if token.replace("x", "").isdigit():
                try:
                    aantal = int(token.replace("x", ""))
                    break
                except Exception:
                    pass

        if "plank" in part:
            result["planken"] += aantal
        elif "push" in part and "lade" in part:
            result["push_to_open_lades"] += aantal
        elif "lade" in part:
            result["lades"] += aantal
        elif "bestek" in part:
            result["bestek_bakken"] += aantal
        elif "spoel" in part:
            result["spoelkast_bescherming"] += aantal
        elif "apotheker" in part:
            result["apothekers"] += aantal
        elif "carrousel" in part:
            result["carrousels"] += aantal
        elif "klep" in part:
            result["klepscharnieren"] += aantal

    return result


def _lees_maatwerk_kasten(path: str):
    """
    Leest tabblad 'MAATWERK KASTEN' en geeft een lijst met kast-dicts terug.

    Rijstructuur (EXACT zoals bevestigd):

      5  = TYPE KAST (A,B of C)
      6  = Hoogte
      7  = Breedte
      8  = Diepte
      9  = Hoogte pootje
      10 = Zichtbare zijde
      11 = Inrichting
      12 = Scharnieren
      13 = Frontmodel
      14 = Aantal fronten
      15 = Kleur corpus
      16 = Dubbelzijdig afgewerkt
      17 = Handgreep
      18 = Afwerking
    """
    try:
        df = pd.read_excel(path, sheet_name="MAATWERK KASTEN", header=None)
    except Exception:
        return []

    kasten = []
    max_row_idx = df.shape[0] - 1

    # Mapping naar 0-based index:
    ROW_TYPE            = 4   # rij 5
    ROW_HOOGTE          = 5   # rij 6
    ROW_BREEDTE         = 6   # rij 7
    ROW_DIEPTE          = 7   # rij 8
    ROW_POOTHOOGTE      = 8   # rij 9
    ROW_ZICHTBARE_ZIJDE = 9   # rij 10
    ROW_INRICHTING      = 10  # rij 11
    ROW_SCHARNIEREN     = 11  # rij 12
    ROW_FRONTMODEL      = 12  # rij 13
    ROW_AANTAL_FRONTEN  = 13  # rij 14
    ROW_KLEUR_CORPUS    = 14  # rij 15
    ROW_DUBBELZIJDIG    = 15  # rij 16
    ROW_HANDGREEP       = 16  # rij 17
    ROW_AFWERKING       = 17  # rij 18

    for col in range(1, 11):  # kolommen B..K (0-based: 1..10)

        # check of kast überhaupt gegevens bevat
        rows_to_check = list(range(ROW_TYPE, ROW_AFWERKING + 1))
        filled = any(
            r <= max_row_idx and not pd.isna(df.iloc[r, col])
            for r in rows_to_check
        )
        if not filled:
            continue

        # Uitlezen velden — exact volgens bevestigde Excel-structuur
        kast_type_raw = df.iloc[ROW_TYPE, col]
        kast_type = str(kast_type_raw).strip().upper() if not pd.isna(kast_type_raw) else ""

        hoogte = _safe_float(df.iloc[ROW_HOOGTE, col]) if ROW_HOOGTE <= max_row_idx else None
        breedte = _safe_float(df.iloc[ROW_BREEDTE, col]) if ROW_BREEDTE <= max_row_idx else None
        diepte = _safe_float(df.iloc[ROW_DIEPTE, col]) if ROW_DIEPTE <= max_row_idx else None
        poothoogte = _safe_float(df.iloc[ROW_POOTHOOGTE, col]) if ROW_POOTHOOGTE <= max_row_idx else None

        zichtbare_zijde = ""
        if ROW_ZICHTBARE_ZIJDE <= max_row_idx and not pd.isna(df.iloc[ROW_ZICHTBARE_ZIJDE, col]):
            zichtbare_zijde = str(df.iloc[ROW_ZICHTBARE_ZIJDE, col]).strip()

        inrichting_raw = ""
        if ROW_INRICHTING <= max_row_idx and not pd.isna(df.iloc[ROW_INRICHTING, col]):
            inrichting_raw = str(df.iloc[ROW_INRICHTING, col]).strip()

        scharnieren = 0
        if ROW_SCHARNIEREN <= max_row_idx:
            scharnieren = _safe_int(df.iloc[ROW_SCHARNIEREN, col])

        frontmodel = ""
        if ROW_FRONTMODEL <= max_row_idx and not pd.isna(df.iloc[ROW_FRONTMODEL, col]):
            frontmodel = str(df.iloc[ROW_FRONTMODEL, col]).strip().upper()

        aantal_fronten = 0
        if ROW_AANTAL_FRONTEN <= max_row_idx and not pd.isna(df.iloc[ROW_AANTAL_FRONTEN, col]):
            aantal_fronten = _safe_int(df.iloc[ROW_AANTAL_FRONTEN, col])

        kleur_corpus = ""
        if ROW_KLEUR_CORPUS <= max_row_idx and not pd.isna(df.iloc[ROW_KLEUR_CORPUS, col]):
            kleur_corpus = str(df.iloc[ROW_KLEUR_CORPUS, col]).strip()

        dubbelzijdig = ""
        if ROW_DUBBELZIJDIG <= max_row_idx and not pd.isna(df.iloc[ROW_DUBBELZIJDIG, col]):
            dubbelzijdig = str(df.iloc[ROW_DUBBELZIJDIG, col]).strip()

        handgreep = ""
        if ROW_HANDGREEP <= max_row_idx and not pd.isna(df.iloc[ROW_HANDGREEP, col]):
            handgreep = str(df.iloc[ROW_HANDGREEP, col]).strip()

        afwerking = ""
        if ROW_AFWERKING <= max_row_idx and not pd.isna(df.iloc[ROW_AFWERKING, col]):
            afwerking = str(df.iloc[ROW_AFWERKING, col]).strip()

        # Parse inrichting
        inrichting_parsed = _parse_inrichting(inrichting_raw)

        # Kast-dict
        kast = {
            "kolom_index": col,
            "type": kast_type,
            "hoogte": hoogte,
            "breedte": breedte,
            "diepte": diepte,
            "poothoogte": poothoogte,
            "kleur_corpus": kleur_corpus,
            "zichtbare_zijde": zichtbare_zijde,
            "inrichting_raw": inrichting_raw,
            "inrichting": inrichting_parsed,
            "scharnieren": scharnieren,
            "frontmodel": frontmodel,
            "aantal_fronten": aantal_fronten,
            "dubbelzijdig": dubbelzijdig,
            "handgreep": handgreep,
            "afwerking": afwerking,
        }

        kasten.append(kast)

    return kasten


def _kast_titel(kast_type: str) -> str:
    t = kast_type.upper()
    if t == "A":
        return "Maatwerk onderkast"
    if t == "B":
        return "Maatwerk hoge kast"
    if t == "C":
        return "Maatwerk hangkast"
    return "Maatwerk kast"


def _bereken_maatwerk_kast(kast: dict):
    """
    Berekent inkoop- en verkoopprijs voor één maatwerk kast:
    - corpus + inrichting + scharnieren (inkoop)
    - fronten + zichtbare zijden in m² + 40% opslag (inkoop)
    - verkoopprijs = inkoop / 0.4 (60% marge)
    """

    kast_type = kast.get("type", "").upper()
    hoogte = kast.get("hoogte") or 0
    breedte = kast.get("breedte") or 0
    diepte = kast.get("diepte") or 0

    idx = _staffel_index(breedte)

    inrichting = kast.get("inrichting", {})
    planken = inrichting.get("planken", 0)
    lades = inrichting.get("lades", 0)
    push_lades = inrichting.get("push_to_open_lades", 0)
    bestek = inrichting.get("bestek_bakken", 0)
    spoelbesch = inrichting.get("spoelkast_bescherming", 0)
    apothekers = inrichting.get("apothekers", 0)
    carrousels = inrichting.get("carrousels", 0)
    klepscharnieren = inrichting.get("klepscharnieren", 0)

    scharnieren = kast.get("scharnieren", 0)

    # -----------------------------
    # 1) CORPUSPRIJS OP BASIS VAN TYPE + HOOGTE
    # -----------------------------
    corpus_inkoop = 0.0

    if kast_type == "A":
        # A-kast: onderkast tot 1000mm
        # Subtype bepalen via inrichting: lades of planken of ovens
        inrichting_raw = (kast.get("inrichting_raw") or "").lower()
        if "lade" in inrichting_raw:
            # ladekast
            corpus_inkoop = A_LADE_KAST[idx]
        elif "plank" in inrichting_raw:
            # voor nu zelfde prijslijn als ladekast
            corpus_inkoop = A_LADE_KAST[idx]
        else:
            # ovenkast
            corpus_inkoop = A_OVEN_KAST[idx]

    elif kast_type == "B":
        # B-kast: hoge kast
        if hoogte is None:
            hoogte = 0
        if 1001 <= hoogte <= 2079:
            corpus_inkoop = B_HOOG_1001_2079[idx]
        elif 2080 <= hoogte <= 2770:
            corpus_inkoop = B_HOOG_2080_2770[idx]
        else:
            # buiten bekende range, default op lagere hoge kast
            corpus_inkoop = B_HOOG_1001_2079[idx]

    elif kast_type == "C":
        # C-kast: hangkast
        if hoogte is None:
            hoogte = 0
        if hoogte <= 390:
            corpus_inkoop = C_CORPUS_0_390[idx]
        elif 391 <= hoogte <= 520:
            corpus_inkoop = C_CORPUS_391_520[idx]
        elif 521 <= hoogte <= 780:
            corpus_inkoop = C_CORPUS_521_780[idx]
        else:
            corpus_inkoop = C_CORPUS_781_PLUS[idx]

    # -----------------------------
    # 2) INRICHTING (PLANKEN, LADES, ETC.)
    # -----------------------------
    inrichting_inkoop = 0.0

    # Inbegrepen planken bepalen
    inbegrepen_planken = 0

    if kast_type == "A":
        # A-kast "geschikt voor planken" → 2 planken inbegrepen
        if "plank" in (kast.get("inrichting_raw") or "").lower():
            inbegrepen_planken = 2

    if kast_type == "C":
        # Hangkasten – afhankelijk van hoogte
        if hoogte <= 390:
            inbegrepen_planken = 0
        elif 391 <= hoogte <= 520:
            inbegrepen_planken = 1
        elif 521 <= hoogte <= 780:
            inbegrepen_planken = 2
        else:  # >=781
            inbegrepen_planken = 2

    extra_planken = max(0, planken - inbegrepen_planken)
    if extra_planken > 0:
        inrichting_inkoop += extra_planken * PLANK_A_OF_B[idx]

    if lades > 0:
        inrichting_inkoop += lades * LADES_KAST[idx]

    if push_lades > 0:
        inrichting_inkoop += push_lades * PUSH_TO_OPEN_LADE[idx]

    if bestek > 0:
        inrichting_inkoop += bestek * BESTEK_BAK[idx]

    if spoelbesch > 0:
        inrichting_inkoop += spoelbesch * SPOELKAST_BESCHERMING[idx]

    if apothekers > 0:
        inrichting_inkoop += apothekers * APOTHEKERS_LADE[idx]

    if carrousels > 0:
        inrichting_inkoop += carrousels * CARROUSEL[idx]

    if klepscharnieren > 0:
        # Zelfde prijs als scharnier per stuk
        inrichting_inkoop += klepscharnieren * SCHARNIER_PER_STUK_MAATWERK[idx]

    # Scharnieren (deurscharnieren uit veld "Scharnieren")
    scharnier_inkoop = 0.0
    if scharnieren > 0:
        scharnier_inkoop = scharnieren * SCHARNIER_PER_STUK_MAATWERK[idx]

    corpus_inrichting_inkoop = corpus_inkoop + inrichting_inkoop + scharnier_inkoop

    # -----------------------------
    # 3) FRONTEN + ZIJKANTEN IN M² + 40% OPSLAG
    # -----------------------------

    frontmodel = (kast.get("frontmodel") or "").upper()
    if frontmodel not in M2_FRONT_PRIJZEN:
        # fallback: geen frontprijs → 0
        front_m2_prijs = 0.0
    else:
        front_m2_prijs = M2_FRONT_PRIJZEN[frontmodel]

    # front-oppervlakte (hele kastfront, aantal fronten is puur informatief)
    front_m2 = (hoogte * breedte) / 1_000_000.0 if hoogte and breedte else 0.0
    front_inkoop = front_m2 * front_m2_prijs

    # Zichtbare zijde(n) → gebruik vlak model per materiaal
    materiaal_type = None
    if frontmodel in MODEL_INFO:
        materiaal_type = MODEL_INFO[frontmodel]["materiaal"]

    zij_m2_inkoop = 0.0
    if materiaal_type in VLAK_MODEL_PER_MATERIAAL:
        vlak_model = VLAK_MODEL_PER_MATERIAAL[materiaal_type]
        vlak_m2_prijs = M2_FRONT_PRIJZEN.get(vlak_model, 0.0)

        zichtbaar = (kast.get("zichtbare_zijde") or "").lower()
        links = "links" in zichtbaar
        rechts = "rechts" in zichtbaar
        # als alleen "ja" staat zonder links/rechts → 1 zijde
        if "ja" in zichtbaar and not (links or rechts):
            links = True

        zijden = 0
        if links:
            zijden += 1
        if rechts:
            zijden += 1

        if zijden > 0 and hoogte and diepte:
            zijde_m2 = (hoogte * diepte) / 1_000_000.0
            zij_m2_inkoop = zijden * zijde_m2 * vlak_m2_prijs

    front_en_zijden_inkoop = front_inkoop + zij_m2_inkoop

    # 40% opslag voor montage fronten op kasten
    front_en_zijden_met_opslag_inkoop = front_en_zijden_inkoop * 1.40

    # -----------------------------
    # 4) TOTAAL INKOOP + VERKOOPPRIJS
    # -----------------------------
    totaal_inkoop = corpus_inrichting_inkoop + front_en_zijden_met_opslag_inkoop

    # 60% marge → verkoop = inkoop / 0.4
    if totaal_inkoop > 0:
        verkoop_excl = round(totaal_inkoop / 0.4, 2)
    else:
        verkoop_excl = 0.0

    totaal_inkoop = round(totaal_inkoop, 2)

    # -----------------------------
    # 5) BESCHRIJVING OPBOUWEN (VOLGORDE = EXCEL)
    # -----------------------------
    aantal_fronten = kast.get("aantal_fronten", 0)

    beschrijving_regels = [
        f"Type kast: {kast_type}",
        f"Hoogte: {hoogte:.0f} mm" if hoogte else "Hoogte: -",
        f"Breedte: {breedte:.0f} mm" if breedte else "Breedte: -",
        f"Diepte: {diepte:.0f} mm" if diepte else "Diepte: -",
        f"Hoogte pootje: {kast.get('poothoogte') or '-'}",
        f"Zichtbare zijde: {kast.get('zichtbare_zijde') or '-'}",
        f"Inrichting: {kast.get('inrichting_raw') or '-'}",
        f"Scharnieren: {scharnieren}",
        f"Frontmodel: {frontmodel or '-'}",
        f"Aantal fronten: {aantal_fronten or '-'}",
        f"Kleur corpus: {kast.get('kleur_corpus') or '-'}",
        f"Dubbelzijdig afgewerkt: {kast.get('dubbelzijdig') or '-'}",
        f"Handgreep: {kast.get('handgreep') or 'n.v.t.'}",
        f"Afwerking: {kast.get('afwerking') or '-'}",
    ]
    beschrijving = "\r\n".join(beschrijving_regels)

    return {
        "titel": _kast_titel(kast_type),
        "beschrijving": beschrijving,
        "totaal_inkoop": totaal_inkoop,
        "verkoop_excl": verkoop_excl,
    }


def _bereken_alle_maatwerk_kasten(kasten_lijst):
    """
    Neemt de lijst kasten (ruwe data uit Excel) en rekent alles door.
    Geeft terug:
    - lijst met kastregels (incl. verkoopprijs)
    - totaal verkoop excl. (som van alle kasten)
    """
    regels = []
    totaal_verkoop = 0.0

    for kast in kasten_lijst:
        res = _bereken_maatwerk_kast(kast)
        if res["verkoop_excl"] > 0:
            regels.append(res)
            totaal_verkoop += res["verkoop_excl"]

    totaal_verkoop = round(totaal_verkoop, 2)
    return regels, totaal_verkoop


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
# 📥 EXCEL UITLEZEN
# ======================================================

def lees_excel(path):
    # Hoofdtabblad (fronten)
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

    projectnaam = os.path.splitext(os.path.basename(path))[0]

    # Maatwerk kasten (tabblad "MAATWERK KASTEN")
    maatwerk_kasten_raw = _lees_maatwerk_kasten(path)

    # project-meta doorgeven zodat bereken_offerte de maatwerk-data heeft
    project_meta = {
        "name": projectnaam,
        "maatwerk_kasten": maatwerk_kasten_raw,
    }

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project_meta


# ======================================================
# 🧮 OFFERTE BEREKENING
# ======================================================

def bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades):

    info = MODEL_INFO[model]

    # Projectnaam & maatwerk-data uit project-meta halen
    if isinstance(project, dict):
        projectnaam = project.get("name", "")
        maatwerk_kasten_raw = project.get("maatwerk_kasten", [])
    else:
        projectnaam = project
        maatwerk_kasten_raw = []

    fronts = sum(o in ["DEUR", "LADE", "BEDEKKINGSPANEEL"] for o in onderdelen)
    heeft_passtuk = any(o in ["PASSTUK", "PLINT"] for o in onderdelen)
    heeft_anders = any("ANDERS" in o for o in onderdelen)

    passtuk_kosten = info["passtuk"] if heeft_passtuk else 0
    anders_kosten = info["passtuk"] if heeft_anders else 0

    materiaal_totaal = fronts * info["prijs_per_front"]
    montage = fronts * MONTAGE_PER_FRONT
    scharnier_totaal = scharnieren * PRIJS_SCHARNIER
    lades_totaal = lades * PRIJS_LADE

    totaal_excl_frontdeel = (
        materiaal_totaal +
        passtuk_kosten +
        anders_kosten +
        montage +
        INMETEN +
        VRACHT +
        scharnier_totaal +
        lades_totaal
    )

    # Maatwerk kasten uitrekenen (inkoop → verkoop)
    maatwerk_regels, maatwerk_totaal_verkoop = _bereken_alle_maatwerk_kasten(maatwerk_kasten_raw)

    # Totaal excl. = front-offerte + maatwerk-kasten (allemaal verkoopprijzen)
    totaal_excl = totaal_excl_frontdeel + maatwerk_totaal_verkoop
    btw = totaal_excl * 0.21
    totaal_incl = totaal_excl + btw

    return {
        "project": projectnaam,
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
        # maatwerk informatie voor Teamleader
        "maatwerk_kasten": maatwerk_regels,
        "maatwerk_totaal_verkoop": maatwerk_totaal_verkoop,
        # ook handig om keuken-deel los te hebben
        "totaal_excl_frontdeel": totaal_excl_frontdeel,
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

    maatwerk_kasten = data.get("maatwerk_kasten") or []
    maatwerk_totaal = data.get("maatwerk_totaal_verkoop", 0.0)

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

        # Keukenrenovatie-bedrag ZONDER maatwerk kasten
        keuken_bedrag = round(data["totaal_excl"] - maatwerk_totaal, 2)

        grouped_lines.append({
            "section": {"title": "KEUKENRENOVATIE"},
            "line_items": [{
                "quantity": 1,
                "description": f"Keukenrenovatie model {model}",
                "extended_description": final,
                "unit_price": {"amount": keuken_bedrag, "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            }],
        })

        # MAATWERK KASTEN ALS APARTE SECTIE (variant B)
        if maatwerk_kasten:
            mk_section = {
                "section": {"title": "MAATWERK KASTEN"},
                "line_items": []
            }
            for kast in maatwerk_kasten:
                mk_section["line_items"].append({
                    "quantity": 1,
                    "description": kast["titel"],
                    "extended_description": kast["beschrijving"],
                    "unit_price": {"amount": kast["verkoop_excl"], "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                })
            grouped_lines.append(mk_section)

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

        # -----------------------------
        # MAATWERK KASTEN (NA KEUKENRENOVATIE, VOOR INMETEN/MONTAGE)
        # -----------------------------
        if maatwerk_kasten:
            mk_section = {
                "section": {"title": "MAATWERK KASTEN"},
                "line_items": []
            }
            for kast in maatwerk_kasten:
                mk_section["line_items"].append({
                    "quantity": 1,
                    "description": kast["titel"],
                    "extended_description": kast["beschrijving"],
                    "unit_price": {"amount": kast["verkoop_excl"], "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                })
            grouped_lines.append(mk_section)

        # ACCESSOIRES (FRONT-SCHARNIEREN/LADES)
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
