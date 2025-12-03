import pandas as pd
import requests
import os
import json

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
# 🧮 MODEL- EN PRIJSLOGICA
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
# 📥 EXCEL UITLEZEN
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
# 🧮 OFFERTE BEREKENING
# ======================================================

def bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades):

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
