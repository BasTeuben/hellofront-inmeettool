import pandas as pd
import requests
import os
import json
from dotenv import load_dotenv

# ======================================================
# 🔧 TEAMLEADER CONFIG
# ======================================================
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TOKEN_FILE = "teamleader_token.json"

API_BASE = "https://api.focus.teamleader.eu"
TOKEN_URL = "https://focus.teamleader.eu/oauth2/access_token"

# 21% BTW-tarief ID
TAX_RATE_21_ID = "94da9f7d-9bf3-04fb-ac49-404ed252c381"

# Vaste kosten
MONTAGE_PER_FRONT = 34.71
INMETEN = 99.17
VRACHT = 60.00

PRIJS_SCHARNIER = 6.5
PRIJS_LADE = 184.0

# ======================================================
# 🔒 TOKEN MANAGEMENT (AUTO REFRESH)
# ======================================================

def read_tokens():
    if not os.path.exists(TOKEN_FILE):
        return None
    try:
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return None


def save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def refresh_access_token():
    tokens = read_tokens()
    if not tokens:
        print("❌ Geen tokenbestand gevonden.")
        return None

    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        print("❌ Geen refresh_token gevonden.")
        return None

    print("🔄 Access token verlopen — vernieuwen...")

    data = {
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token,
    }

    resp = requests.post(TOKEN_URL, data=data)

    if resp.status_code != 200:
        print("❌ Fout bij vernieuwen access_token:")
        print(resp.text)
        return None

    new_tokens = resp.json()
    save_tokens(new_tokens)
    print("🔄 Nieuwe access_token opgeslagen!")
    return new_tokens.get("access_token")


def request_with_auto_refresh(method: str, url: str, json_data=None, files=None):
    tokens = read_tokens()
    if not tokens:
        print("❌ Geen tokens gevonden.")
        return None

    access_token = tokens.get("access_token")

    headers = {"Authorization": f"Bearer {access_token}"}
    if not files:
        headers["Content-Type"] = "application/json"

    resp = requests.request(method, url, headers=headers, json=json_data, files=files)

    if resp.status_code == 401:
        new_token = refresh_access_token()
        if not new_token:
            return None

        headers["Authorization"] = f"Bearer {new_token}"

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
    "NOAH": {"titel": "Keukenrenovatie model Noah", "materiaal": "MDF Gespoten - vlak", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde standaard wit"},
    "FEDDE": {"titel": "Keukenrenovatie model Fedde", "materiaal": "MDF Gespoten - greeploos", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde standaard wit"},
    "DAVE": {"titel": "Keukenrenovatie model Dave", "materiaal": "MDF Gespoten - 70mm kader", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde standaard wit"},
    "JOLIE": {"titel": "Keukenrenovatie model Jolie", "materiaal": "MDF Gespoten - 25mm kader", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde standaard wit"},
    "DEX": {"titel": "Keukenrenovatie model Dex", "materiaal": "MDF Gespoten - V-groef 100mm", "frontdikte": "18mm", "afwerking": "Zijdeglans", "dubbelzijdig": "Nee, binnenzijde standaard wit"},

    "JACK": {"titel": "Keukenrenovatie model Jack", "materiaal": "Eiken fineer - vlak", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja, standaard"},
    "JAMES": {"titel": "Keukenrenovatie model James", "materiaal": "Eiken fineer - 10mm massief kader", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja, standaard"},
    "CHIEL": {"titel": "Keukenrenovatie model Chiel", "materiaal": "Eiken fineer - greeploos", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja, standaard"},

    "SAM": {"titel": "Keukenrenovatie model Sam", "materiaal": "Noten fineer - vlak", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja, standaard"},
    "DUKE": {"titel": "Keukenrenovatie model Duke", "materiaal": "Noten fineer - greeploos", "frontdikte": "19mm", "afwerking": "Monocoat olie", "dubbelzijdig": "Ja, standaard"},
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
    kleur = df.iloc[1, 8]  # I2

    klantregels = []
    for r in range(1, 6):
        waarde = df.iloc[r, 10]
        if pd.notna(waarde):
            klantregels.append(str(waarde))

    # Scharnieren J3 = row 2 col 9
    try:
        scharnieren = int(df.iloc[2, 9]) if not pd.isna(df.iloc[2, 9]) else 0
    except:
        scharnieren = 0

    # Maatwerk lades J5 = row 4 col 9
    try:
        lades = int(df.iloc[4, 9]) if not pd.isna(df.iloc[4, 9]) else 0
    except:
        lades = 0

    project = os.path.splitext(os.path.basename(path))[0]

    return onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project


# ======================================================
# 🧮 BEREKENINGEN
# ======================================================

def euro(x):
    return f"€ {x:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


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
        "prijs_per_front": info["prijs_per_front"],
        "materiaal_totaal": materiaal_totaal,
        "toeslag_passtuk": passtuk_kosten,
        "toeslag_anders": anders_kosten,
        "montage": montage,
        "scharnieren": scharnieren,
        "scharnier_totaal": scharnier_totaal,
        "lades": lades,
        "lades_totaal": lades_totaal,
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

    klanttekst = "\r\n".join(data["klantgegevens"])
    # Bouw vaste layout voor klantgegevens
    klantregels = data["klantgegevens"] + ["", "", "", "", ""]  # voorkomt indexfouten

    klanttekst = (
        f"Naam: {klantregels[0]}\r\n"
        f"Adres: {klantregels[1]}\r\n"
        f"Postcode / woonplaats: {klantregels[2]}\r\n"
        f"Tel: {klantregels[3]}\r\n"
        f"Email: {klantregels[4]}"
    )

    grouped_lines = []

    # =====================================================
    # ✨ EERSTE REGEL: KLANTGEGEVENS
    # =====================================================
# ================================================
# 🔹 SECTIE: KLANTGEGEVENS
# ================================================
    grouped_lines.append({
        "section": {"title": "KLANTGEGEVENS"},
        "line_items": [
            {
                "quantity": 1,
                "description": "Klantgegevens",
                "extended_description": klanttekst,
                "unit_price": {"amount": 0, "tax": "excluding"},
                "tax_rate_id": TAX_RATE_21_ID,
            }
        ],
    })

    # =====================================================
    # 🧍 PARTICULIER — 1 REGELITEM
    # =====================================================
    if mode == "P":

        tekst = []

        # TOEVOEGINGEN ACHTER "AANTAL FRONTEN"
        front_toevoegingen = []

        if data["toeslag_passtuk"] > 0:
            front_toevoegingen.append("inclusief passtukken en/of plinten")

        if data["toeslag_anders"] > 0:
            front_toevoegingen.append("inclusief licht- en/of sierlijsten")

        # Bouw de regel "Aantal fronten"
        if front_toevoegingen:
            extra_text = " (" + ", ".join(front_toevoegingen) + ")"
        else:
            extra_text = ""

        tekst.append(f"Aantal fronten: {fronts} fronten{extra_text}")

        # BASIS-INFORMATIE
        tekst.append(f"Materiaal: {cfg['materiaal']}")
        tekst.append(f"Frontdikte: {cfg['frontdikte']}")
        tekst.append(f"Kleur: {data['kleur']}")
        tekst.append(f"Afwerking: {cfg['afwerking']}")
        tekst.append(f"Dubbelzijdig in kleur afwerken: {cfg['dubbelzijdig']}")
        tekst.append("Inmeten: Ja")
        tekst.append("Montage: Ja")
        tekst.append("Handgrepen: Te bepalen")
        tekst.append("")
        tekst.append("Prijs is inclusief:")
        tekst.append("- Demontage oude fronten & materialen")

        # BASIS MONTAGEREGEL
        tekst.append("- Inmeten, leveren en montage van de fronten")

        # EXTRA'S — ALLEEN WEERGEVEN ALS IN EXCEL AANWEZIG
        if data["toeslag_passtuk"] > 0:
            tekst.append("- Montage van passtukken en/of plinten")

        if data["toeslag_anders"] > 0:
            tekst.append("- Montage van licht- en/of sierlijsten")

        tekst.append("- Afvoeren van oude fronten")

        # SCHARNIEREN + AANTAL
        if data["scharnieren"] > 0:
            tekst.append(f"- Inclusief vervangen scharnieren ({data['scharnieren']} stuks)")

        # MAATWERK LADES + AANTAL
        if data["lades"] > 0:
            tekst.append(f"- Inclusief plaatsen maatwerk lades ({data['lades']} stuks)")

        final_text = "\r\n".join(tekst)


        grouped_lines.append({
            "section": {"title": "KEUKENRENOVATIE"},
            "line_items": [
                {
                    "quantity": 1,
                    "description": f"Keukenrenovatie model {model}",
                    "extended_description": final_text,
                    "unit_price": {"amount": round(data["totaal_excl"], 2), "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                }
            ],
        })

    # =====================================================
    # 🏭 DEALER — MEERDERE REGELS
    # =====================================================
    else:

        # KEUKENRENOVATIE
        section = {"section": {"title": "KEUKENRENOVATIE"}, "line_items": []}

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
                f"Inmeten: Ja\r\n"
                f"Montage: Ja\r\n"
                f"\r\n"
                f"Fronten worden geleverd zonder scharnieren"
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

        # INMETEN, LEVEREN & MONTEREN
        grouped_lines.append({
            "section": {"title": "INMETEN, LEVEREN & MONTEREN"},
            "line_items": [
                {
                    "quantity": 1,
                    "description": "Inmeten",
                    "extended_description": "inmeten op locatie",
                    "unit_price": {"amount": INMETEN, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
                {
                    "quantity": fronts,
                    "description": "Montage per front",
                    "extended_description": "inclusief demontage oude fronten & afvoeren",
                    "unit_price": {"amount": MONTAGE_PER_FRONT, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
                {
                    "quantity": 1,
                    "description": "Vracht- & verpakkingskosten",
                    "extended_description": "Levering op locatie (excl. waddeneilanden)",
                    "unit_price": {"amount": VRACHT, "tax": "excluding"},
                    "tax_rate_id": TAX_RATE_21_ID,
                },
            ],
        })

    payload = {
        "deal_id": deal_id,
        "currency": {"code": "EUR", "exchange_rate": 1.0},
        "grouped_lines": grouped_lines,
        "text": "\u200b",  # begeleidende tekst leeg houden
    }

    print("\n📨 VERZONDEN PAYLOAD:")
    print(json.dumps(payload, indent=2))

    resp = request_with_auto_refresh("POST", url, json_data=payload)

    if not resp:
        print("❌ Geen response!")
        return

    print("➡️ STATUS:", resp.status_code)
    print("➡️ RESPONSE:", resp.text)

    if resp.status_code not in (200, 201):
        print("❌ Offerte NIET aangemaakt.")
        return

    print("✅ Offerte aangemaakt!")


# ======================================================
# 🚀 MAIN
# ======================================================

if __name__ == "__main__":
    pad = input("Sleep het Excel-bestand hierheen en druk Enter: ").strip().replace("\\", "")

    onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = lees_excel(pad)

    model = bepaal_model(g2, h2)
    if not model:
        print("❌ MODEL ONBEKEND")
        exit()

    data = bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

    print("\n📋 SAMENVATTING")
    print("──────────────────────────────────────────")
    print(f"Project:               {data['project']}")
    print(f"Model:                 {data['model']}")
    print(f"Materiaal:             {data['materiaal']}")
    print(f"Kleur:                 {data['kleur']}")
    print(f"Fronten:               {data['fronts']}")
    print(f"Scharnieren:           {data['scharnieren']} × € {PRIJS_SCHARNIER} = {euro(data['scharnier_totaal'])}")
    print(f"Maatwerk lades:        {data['lades']} × € {PRIJS_LADE} = {euro(data['lades_totaal'])}")
    print("──────────────────────────────────────────")
    print(f"Aantal fronten:        {data['fronts']} × {euro(data['prijs_per_front'])} = {euro(data['materiaal_totaal'])}")
    print(f"Plinten/passtukken:    {'Ja' if data['toeslag_passtuk'] else 'Nee'} - {euro(data['toeslag_passtuk'])}")
    print(f"Licht- & sierlijsten:  {'Ja' if data['toeslag_anders'] else 'Nee'} - {euro(data['toeslag_anders'])}")
    print(f"Montage:               {data['fronts']} × {euro(MONTAGE_PER_FRONT)} = {euro(data['montage'])}")
    print(f"Inmeten:               {euro(INMETEN)}")
    print(f"Vracht/verpakking:     {euro(VRACHT)}")
    print("──────────────────────────────────────────")
    print(f"Totaal excl. btw:      {euro(data['totaal_excl'])}")
    print(f"BTW (21%):             {euro(data['btw'])}")
    print(f"Totaal incl. btw:      {euro(data['totaal_incl'])}")
    print("──────────────────────────────────────────")

    if input("Wil je uploaden naar Teamleader? (ja/nee): ").lower() != "ja":
        print("❌ Afgebroken.")
        exit()

    mode = input("Uploaden als Dealer (D) of Particulier (P)? ").strip().upper()
    deal_id = input("Teamleader deal-ID: ").strip()

    maak_teamleader_offerte(deal_id, data, mode)
