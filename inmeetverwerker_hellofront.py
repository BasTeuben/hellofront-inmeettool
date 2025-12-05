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
# 🧮 MODEL- EN PRIJSLOGICA (BESTAANDE FRONTEN)
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
# 🧮 MAATWERK KASTEN – M²-PRIJZEN FRONTEN
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

# Voor zijkanten in vlak model per materiaaltype
# MDF = NOAH, EIKEN = JACK, NOTEN = SAM
VLAK_MODEL_PER_MATERIAAL = {
    "MDF gespoten": "NOAH",
    "Eikenfineer": "JACK",
    "Noten fineer": "SAM",
}

# Breedte-staffels in mm
KAST_BREEDTE_STAFFELS = [300, 400, 500, 600, 800, 900, 1000, 1200]

# Prijstabellen uit "Prijslijst maatwerk kasten + inrichting"
KASTPRIJS_A_OVEN = {600: 58.0}

KASTPRIJS_A_LADES = {
    300: 45.0, 400: 47.0, 500: 49.0, 600: 51.0,
    800: 59.0, 900: 61.0, 1000: 63.0, 1200: 66.0,
}

KASTPRIJS_A_PLANK2 = {
    300: 51.0, 400: 55.0, 500: 60.0, 600: 64.0,
    800: 77.0, 900: 81.0, 1000: 85.0, 1200: 92.0,
}

KASTPRIJS_B_2080_2600 = {
    300: 85.0, 400: 89.0, 500: 92.0, 600: 97.0,
    800: 103.0, 900: 108.0, 1000: 114.0,
}

KASTPRIJS_B_1000_1950 = {
    300: 80.0, 400: 84.0, 500: 87.0, 600: 90.0,
    800: 95.0, 900: 98.0, 1000: 103.0,
}

KASTPRIJS_C_390 = {
    300: 35.0, 400: 37.0, 500: 39.0, 600: 41.0,
    800: 47.0, 900: 48.0, 1000: 50.0, 1200: 56.0,
}

KASTPRIJS_C_391_520 = {
    300: 39.0, 400: 42.0, 500: 45.0, 600: 47.0,
    800: 56.0, 900: 58.0, 1000: 61.0, 1200: 67.0,
}

KASTPRIJS_C_521_780 = {
    300: 42.0, 400: 45.0, 500: 49.0, 600: 50.0,
    800: 57.0, 900: 61.0, 1000: 66.0, 1200: 70.0,
}

KASTPRIJS_C_781P = {
    300: 46.0, 400: 48.0, 500: 51.0, 600: 55.0,
    800: 61.0, 900: 65.0, 1000: 68.0, 1200: 75.0,
}

# Inrichting / accessoires
PRIJS_KLEPSCHARNIER = {b: 63.0 for b in KAST_BREEDTE_STAFFELS}
PRIJS_PLANK_AB = {
    300: 4.5, 400: 6.3, 500: 7.5, 600: 7.6,
    800: 10.6, 900: 11.8, 1000: 13.3, 1200: 15.75,
}
PRIJS_LADES_KAST = {b: 68.0 for b in KAST_BREEDTE_STAFFELS}
PRIJS_PUSH_LADE = {b: 109.0 for b in KAST_BREEDTE_STAFFELS}
PRIJS_SCHARNIER_KAST = {b: 9.7 for b in KAST_BREEDTE_STAFFELS}
PRIJS_BESTEK_BAK = {b: 33.0 for b in KAST_BREEDTE_STAFFELS}
PRIJS_PLASTIC_SPOEL = {b: 33.0 for b in KAST_BREEDTE_STAFFELS}
PRIJS_APOTHEKER = {300: 546.0, 400: 546.0, 500: 546.0}
PRIJS_CARROUSEL = {800: 452.0, 900: 452.0, 1000: 452.0, 1200: 452.0}


def kies_breedte_staffel(breedte_mm: float) -> int:
    """Kies eerstvolgende staffel: 610mm → 800mm, enz."""
    if breedte_mm is None:
        return KAST_BREEDTE_STAFFELS[0]
    try:
        b = float(breedte_mm)
    except (TypeError, ValueError):
        return KAST_BREEDTE_STAFFELS[0]

    for s in KAST_BREEDTE_STAFFELS:
        if b <= s:
            return s
    return KAST_BREEDTE_STAFFELS[-1]


def parse_int_safe(v):
    if pd.isna(v):
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    if not s:
        return 0
    m = re.search(r"-?\d+", s)
    return int(m.group(0)) if m else 0


def parse_inrichting(inrichting_str: str):
    """
    Haalt aantallen uit de inrichtingstekst, zoals:
    '3x plank, 1x lade' etc.
    """
    result = {
        "plank": 0,
        "lade": 0,
        "bestekbak": 0,
        "apotheker": 0,
        "carrousel": 0,
        "klepscharnier": 0,
        "push_lade": 0,
        "plastic_spoel": 0,
    }
    if not isinstance(inrichting_str, str):
        return result

    text = inrichting_str.lower()

    def count(patterns):
        total = 0
        for p in patterns:
            for m in re.finditer(r"(\d+)\s*x?\s*" + p, text):
                total += int(m.group(1))
        return total

    # planken
    result["plank"] = count(["plank", "planken"])

    # lades
    result["lade"] = count(["lade", "lades", "laden"])

    # bestekbak
    result["bestekbak"] = count(["bestek", "bestekbak"])

    # apothekerslade
    result["apotheker"] = count(["apotheker", "apothekers"])

    # carrousel
    result["carrousel"] = count(["carrousel"])

    # klepscharnieren
    result["klepscharnier"] = count(["klepscharnier", "klepscharnieren"])

    # push-to-open
    result["push_lade"] = count(["push to open", "push-to-open"])

    # plastic bescherming spoelkast
    result["plastic_spoel"] = count(["plastic bescherming", "spoelkast"])

    # Als er bijv. 'lade' wordt genoemd zonder getal → 1
    if "lade" in text and result["lade"] == 0:
        result["lade"] = 1

    if "plank" in text and result["plank"] == 0:
        result["plank"] = 1

    return result


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

def _lees_fronten_sheet(path):
    """Bestaande logica – tabblad 1 (fronten)."""
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

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project


def _normalize_label(label: str) -> str:
    if not isinstance(label, str):
        label = str(label)
    return label.strip().lower()


def lees_maatwerk_kasten_sheet(path):
    """
    Leest tabblad 2: maatwerk kasten.
    Verwacht in kolom A de labels (Type kast, Hoogte, Breedte, etc.)
    en in kolommen B t/m K per kolom één kast.
    """
    try:
        df = pd.read_excel(path, sheet_name=1, header=None)
    except Exception:
        return []

    labels = df.iloc[:, 0].astype(str).apply(_normalize_label)
    kasten = []

    # Kolommen B t/m K (index 1 t/m 10)
    for col in range(1, min(11, df.shape[1])):
        col_values = df.iloc[:, col]

        # check of er überhaupt iets ingevuld is
        if col_values.isna().all():
            continue

        kast = {
            "kolom": col,  # ter info/debuggen
            "raw": {},
        }

        for row_idx, lab in enumerate(labels):
            value = col_values.iloc[row_idx]
            if (isinstance(value, float) or isinstance(value, int)) and pd.isna(value):
                continue
            if str(value).strip() == "" or str(value).strip().lower() == "nan":
                continue

            kast["raw"][lab] = value

        # beslissen of dit echt een kast is: minimaal type of hoogte/breedte
        has_type = any("type" in k for k in kast["raw"].keys())
        has_hoogte = any("hoogte" in k and "poot" not in k for k in kast["raw"].keys())
        has_breedte = any("breedte" in k for k in kast["raw"].keys())

        if not (has_type or (has_hoogte and has_breedte)):
            continue

        kasten.append(kast)

    return kasten


def lees_excel(path):
    """
    Hoofd-ingang:
    - fronten + klantdata (bestaand gedrag)
    - maatwerk kasten (nieuw) van tabblad 2
    """
    onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = _lees_fronten_sheet(path)
    maatwerk_kasten = lees_maatwerk_kasten_sheet(path)

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project, maatwerk_kasten


# ======================================================
# 🧮 MAATWERK KASTEN BEREKENING
# ======================================================

def _bepaal_kasttype_en_basisprijs(kast, staffel_breedte):
    """
    Bepaalt A/B/C en subtype en levert basis kastprijs (zonder inrichting, fronten, zijkanten).
    """
    raw = kast.get("raw", {})
    # type kast staat in label met 'type kast'
    type_kast = None
    for k, v in raw.items():
        if "type" in k and "kast" in k:
            type_kast = str(v).strip().upper()
            break

    hoogte = None
    for k, v in raw.items():
        if "hoogte" in k and "poot" not in k:
            try:
                hoogte = float(v)
            except (TypeError, ValueError):
                hoogte = None
            break

    # default
    if type_kast not in ("A", "B", "C"):
        # eventueel op hoogte afleiden
        if hoogte is not None and hoogte <= 1000:
            type_kast = "A"
        elif hoogte is not None and hoogte > 1000:
            type_kast = "B"

    basisprijs = 0.0
    subtype = None

    # inrichting uitlezen voor subtype A
    inrichting_str = ""
    for k, v in raw.items():
        if "inrichting" in k:
            inrichting_str = str(v)
            break
    inrichting_counts = parse_inrichting(inrichting_str)

    if type_kast == "A":
        # A-kast: onderkast tot 1000mm
        # subtype op basis van inrichting:
        if inrichting_counts["lade"] > 0:
            subtype = "A_LADES"
            basisprijs = KASTPRIJS_A_LADES.get(staffel_breedte, 0.0)
        elif inrichting_counts["plank"] > 0:
            subtype = "A_PLANK"
            basisprijs = KASTPRIJS_A_PLANK2.get(staffel_breedte, 0.0)
        else:
            subtype = "A_OVEN"
            basisprijs = KASTPRIJS_A_OVEN.get(staffel_breedte, 0.0)

    elif type_kast == "B":
        # hoge kasten
        if hoogte is None:
            subtype = "B_ONBEKEND"
            basisprijs = 0.0
        else:
            if hoogte >= 2080:
                subtype = "B_HOOG_2080_2600"
                basisprijs = KASTPRIJS_B_2080_2600.get(staffel_breedte, 0.0)
            elif hoogte >= 1001:
                subtype = "B_HOOG_1000_1950"
                basisprijs = KASTPRIJS_B_1000_1950.get(staffel_breedte, 0.0)
            else:
                subtype = "B_HOOG_1000_1950"
                basisprijs = KASTPRIJS_B_1000_1950.get(staffel_breedte, 0.0)

    elif type_kast == "C":
        # hangkasten
        if hoogte is None:
            subtype = "C_ONBEKEND"
            basisprijs = 0.0
        else:
            if hoogte <= 390:
                subtype = "C_390"
                basisprijs = KASTPRIJS_C_390.get(staffel_breedte, 0.0)
            elif 391 <= hoogte <= 520:
                subtype = "C_391_520"
                basisprijs = KASTPRIJS_C_391_520.get(staffel_breedte, 0.0)
            elif 521 <= hoogte <= 780:
                subtype = "C_521_780"
                basisprijs = KASTPRIJS_C_521_780.get(staffel_breedte, 0.0)
            else:  # >= 781
                subtype = "C_781P"
                basisprijs = KASTPRIJS_C_781P.get(staffel_breedte, 0.0)

    return type_kast, subtype, basisprijs, inrichting_counts, hoogte


def _bepaal_inbegrepen_planken(type_kast, subtype, hoogte):
    """
    Regels:
    - A-kast planken: kast met 2 planken → 2 inbegrepen
    - C-kasten:
        t/m 390 mm: 0
        391–520: 1
        521–780: 2
        >=781: 2
    """
    if type_kast == "A" and subtype == "A_PLANK":
        return 2

    if type_kast == "C":
        if hoogte is None:
            return 0
        if hoogte <= 390:
            return 0
        elif 391 <= hoogte <= 520:
            return 1
        elif 521 <= hoogte <= 780:
            return 2
        else:
            return 2

    return 0


def _bereken_front_en_zijkanten_prijs(kast, hoogte_mm, breedte_mm, diepte_mm):
    raw = kast.get("raw", {})

    frontmodel = None
    kleur = None
    zichtbare_zijde = ""
    for k, v in raw.items():
        lk = k.lower()
        if "front" in lk and "indeling" not in lk:
            frontmodel = str(v).strip().upper()
        elif "kleur" in lk:
            kleur = str(v).strip()
        elif "zichtbare" in lk:
            zichtbare_zijde = str(v).strip().lower()

    if frontmodel not in M2_FRONT_PRIJZEN:
        # geen frontmodel → geen prijs voor fronten
        return 0.0

    m2_prijs_front = M2_FRONT_PRIJZEN[frontmodel]

    # front-oppervlakte (totaal)
    try:
        h_m = float(hoogte_mm) / 1000.0
        b_m = float(breedte_mm) / 1000.0
        d_m = float(diepte_mm) / 1000.0
    except (TypeError, ValueError):
        return 0.0

    front_m2 = h_m * b_m
    front_prijs = front_m2 * m2_prijs_front

    # materiaal bepalen voor vlak model (voor zijkant)
    info = MODEL_INFO.get(frontmodel)
    zijkant_prijs = 0.0
    if info:
        materiaal = info["materiaal"]  # "MDF gespoten", "Eikenfineer", "Noten fineer"
        vlak_model = VLAK_MODEL_PER_MATERIAAL.get(materiaal)
        if vlak_model:
            m2_prijs_vlak = M2_FRONT_PRIJZEN[vlak_model]

            # zichtbare zijde: ja, links / rechts / beide
            zijden = 0
            txt = zichtbare_zijde
            if "links" in txt and "rechts" in txt:
                zijden = 2
            elif "beide" in txt:
                zijden = 2
            elif "ja" in txt and ("links" in txt or "rechts" in txt):
                zijden = 1
            elif txt.strip().lower() in ("links", "rechts"):
                zijden = 1

            if zijden > 0:
                zij_m2 = h_m * d_m
                zijkant_prijs = zijden * zij_m2 * m2_prijs_vlak

    # 40% opslag voor montage fronten op kasten
    totaal = (front_prijs + zijkant_prijs) * 1.40
    return round(totaal, 2)


def bereken_maatwerk_kasten(maatwerk_kasten_raw):
    """
    Neemt lijst zoals uit lees_maatwerk_kasten_sheet
    en berekent per kast:
    - inkoopprijs
    - verkoopprijs (inkoop / 0.4)
    - type titel ("Maatwerk hoge kast" etc.)
    - beschrijvingstekst (opsomming zoals Excel)
    """
    resultaten = []
    totaal_inkoop = 0.0

    for kast in maatwerk_kasten_raw:
        raw = kast.get("raw", {})

        # basis-gegevens uit raw
        hoogte = None
        breedte = None
        diepte = None
        hoogte_poot = None
        inrichting_str = ""
        zichtbare_zijde = ""
        scharnieren_kast = 0
        frontmodel = ""
        kleur_corpus = ""
        dubbelzijdig = ""
        handgreep = ""
        afwerking = ""
        type_kast_label = ""

        for k, v in raw.items():
            lk = k.lower()

            if "type" in lk and "kast" in lk:
                type_kast_label = str(v).strip()

            elif "hoogte poot" in lk or "hoogte pootje" in lk:
                try:
                    hoogte_poot = float(v)
                except (TypeError, ValueError):
                    hoogte_poot = None

            elif "hoogte" in lk and "poot" not in lk:
                try:
                    hoogte = float(v)
                except (TypeError, ValueError):
                    hoogte = None

            elif "breedte" in lk:
                try:
                    breedte = float(v)
                except (TypeError, ValueError):
                    breedte = None

            elif "diepte" in lk:
                try:
                    diepte = float(v)
                except (TypeError, ValueError):
                    diepte = None

            elif "inrichting" in lk:
                inrichting_str = str(v)

            elif "zichtbare" in lk:
                zichtbare_zijde = str(v)

            elif "scharnier" in lk:
                scharnieren_kast = parse_int_safe(v)

            elif "front indeling" in lk:
                # wordt alleen in beschrijving gebruikt
                pass

            elif "front" in lk and "indeling" not in lk:
                frontmodel = str(v).strip().upper()

            elif "kleur corpus" in lk:
                kleur_corpus = str(v).strip()

            elif lk == "kleur" or "kleur:" in lk:
                # kleur front
                pass

            elif "dubbelzijdig" in lk:
                dubbelzijdig = str(v).strip()

            elif "handgreep" in lk:
                handgreep = str(v).strip()

            elif "afwerking" in lk:
                afwerking = str(v).strip()

        if hoogte is None or breedte is None or diepte is None:
            # onvoldoende data om prijs te berekenen
            continue

        staffel = kies_breedte_staffel(breedte)
        type_kast, subtype, basisprijs, inrichting_counts, hoogte_mm = _bepaal_kasttype_en_basisprijs(
            kast, staffel
        )

        # inbegrepen planken per type
        inbegrepen_planken = _bepaal_inbegrepen_planken(type_kast, subtype, hoogte_mm or hoogte)

        # inrichting extra prijzen
        extra_planken = max(0, inrichting_counts["plank"] - inbegrepen_planken)
        plank_prijs = extra_planken * PRIJS_PLANK_AB.get(staffel, 0.0)

        lade_prijs = inrichting_counts["lade"] * PRIJS_LADES_KAST.get(staffel, 0.0)
        bestek_prijs = inrichting_counts["bestekbak"] * PRIJS_BESTEK_BAK.get(staffel, 0.0)
        apotheker_prijs = inrichting_counts["apotheker"] * PRIJS_APOTHEKER.get(staffel, 0.0)
        carrousel_prijs = inrichting_counts["carrousel"] * PRIJS_CARROUSEL.get(staffel, 0.0)
        klep_prijs = inrichting_counts["klepscharnier"] * PRIJS_KLEPSCHARNIER.get(staffel, 0.0)
        push_prijs = inrichting_counts["push_lade"] * PRIJS_PUSH_LADE.get(staffel, 0.0)
        plastic_prijs = inrichting_counts["plastic_spoel"] * PRIJS_PLASTIC_SPOEL.get(staffel, 0.0)

        # scharnieren volgens kast-prijslijst
        scharnier_prijs = scharnieren_kast * PRIJS_SCHARNIER_KAST.get(staffel, 0.0)

        # fronten + zijkant + 40% opslag
        front_en_zijkant_prijs = _bereken_front_en_zijkanten_prijs(
            kast, hoogte, breedte, diepte
        )

        inkoop = (
            basisprijs +
            plank_prijs +
            lade_prijs +
            bestek_prijs +
            apotheker_prijs +
            carrousel_prijs +
            klep_prijs +
            push_prijs +
            plastic_prijs +
            scharnier_prijs +
            front_en_zijkant_prijs
        )

        verkoop = round(inkoop / 0.4, 2) if inkoop > 0 else 0.0
        totaal_inkoop += inkoop

        # titel per type
        if type_kast == "A":
            titel = "Maatwerk onderkast"
        elif type_kast == "B":
            titel = "Maatwerk hoge kast"
        elif type_kast == "C":
            titel = "Maatwerk hangkast"
        else:
            titel = "Maatwerk kast"

        # beschrijving: opsomming zoals in Excel
        beschrijvingsregels = []
        # we reconstrueren in een vaste volgorde m.b.v. de bekende labels
        raw_l = {k.lower(): v for k, v in raw.items()}

        def get_label(stuk):
            for k, v in raw_l.items():
                if stuk in k:
                    return v
            return ""

        beschrijvingsregels.append(f"Type kast: {type_kast_label or type_kast}")
        beschrijvingsregels.append(f"Hoogte: {hoogte:.0f} mm")
        beschrijvingsregels.append(f"Breedte: {breedte:.0f} mm")
        beschrijvingsregels.append(f"Diepte: {diepte:.0f} mm")
        if hoogte_poot is not None:
            beschrijvingsregels.append(f"Hoogte pootje: {hoogte_poot:.0f} mm")
        beschrijvingsregels.append(f"Kleur corpus: {kleur_corpus}")
        beschrijvingsregels.append(f"Zichtbare zijde: {zichtbare_zijde}")
        beschrijvingsregels.append(f"Inrichting: {inrichting_str}")
        beschrijvingsregels.append(f"Scharnieren: {scharnieren_kast}")
        beschrijvingsregels.append(f"Frontmodel: {frontmodel}")
        fi = get_label("front indeling")
        if fi:
            beschrijvingsregels.append(f"Front indeling: {fi}")
        kleur_front = get_label("kleur")
        if kleur_front:
            beschrijvingsregels.append(f"Kleur: {kleur_front}")
        if dubbelzijdig:
            beschrijvingsregels.append(f"Dubbelzijdig afgewerkt: {dubbelzijdig}")
        if handgreep:
            beschrijvingsregels.append(f"Handgreep: {handgreep}")
        if afwerking:
            beschrijvingsregels.append(f"Afwerking: {afwerking}")

        beschrijving = "\r\n".join(beschrijvingsregels)

        resultaten.append({
            "titel": titel,
            "beschrijving": beschrijving,
            "inkoop": round(inkoop, 2),
            "verkoop": verkoop,
        })

    return resultaten, totaal_inkoop


# ======================================================
# 🧮 OFFERTE BEREKENING (FRONTEN + MAATWERK KASTEN)
# ======================================================

def bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades, maatwerk_kasten_raw):

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

    # Bestaand totaal voor keukenrenovatie (fronten e.d.)
    totaal_excl_fronten = (
        materiaal_totaal +
        passtuk_kosten +
        anders_kosten +
        montage +
        INMETEN +
        VRACHT +
        scharnier_totaal +
        lades_totaal
    )

    # NIEUW: maatwerk kasten
    maatwerk_kasten, maatwerk_inkoop = bereken_maatwerk_kasten(maatwerk_kasten_raw)
    maatwerk_verkoop_totaal = round(maatwerk_inkoop / 0.4, 2) if maatwerk_inkoop > 0 else 0.0

    totaal_excl = totaal_excl_fronten + maatwerk_verkoop_totaal
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
        # maatwerk kasten info
        "maatwerk_kasten": maatwerk_kasten,
        "maatwerk_inkoop_totaal": round(maatwerk_inkoop, 2),
        "maatwerk_verkoop_totaal": maatwerk_verkoop_totaal,
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

        # Maatwerk kasten tekstueel kun je later eventueel nog toevoegen;
        # nu staat hun prijs wel in het totaal.
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

        # MAATWERK KASTEN – NA KEUKENRENOVATIE, VOOR INMETEN/MONTAGE
        maatwerk_kasten = data.get("maatwerk_kasten") or []
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
                    "unit_price": {"amount": kast["verkoop"], "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                })
            grouped_lines.append(mk_section)

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
