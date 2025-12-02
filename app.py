import streamlit as st
import os
import tempfile
import requests
from urllib.parse import urlencode
import inmeetverwerker_hellofront as hf

# ======================================================
# 1. BASISCONFIG
# ======================================================

# Haal ENV-variabelen op
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("REFRESH_TOKEN", "")

REDIRECT_URI = "https://hellofront-inmeettool-production.up.railway.app/"
AUTH_BASE = "https://app.teamleader.eu/oauth2/authorize"
TOKEN_URL = "https://focus.teamleader.eu/oauth2/access_token"

st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand om automatisch een offerte aan te maken in Teamleader.")

# 🔍 Debug: laat ZIEN wat Python daadwerkelijk uit de omgeving leest
st.write("DEBUG – CLIENT_ID uit env:", repr(CLIENT_ID))
st.write("DEBUG – CLIENT_SECRET leeg?:", CLIENT_SECRET == "")
st.write("DEBUG – REDIRECT_URI:", REDIRECT_URI)

# ======================================================
# 2. CALLBACK HANDLING (als Teamleader ?code=... terugstuurt)
# ======================================================
params = st.query_params  # nieuwe API i.p.v. experimental

auth_code = params.get("code", None)

if auth_code:
    st.subheader("Teamleader koppeling")
    st.info("Authorisatiecode ontvangen, tokens worden opgehaald…")

    data = {
        "grant_type": "authorization_code",
        "code": auth_code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
    }

    resp = requests.post(TOKEN_URL, data=data)
    st.write("DEBUG – token response raw:", resp.text)

    if resp.status_code != 200:
        st.error(f"❌ Ophalen tokens mislukt:\n\n{resp.text}")
        st.stop()

    tokens = resp.json()
    new_refresh = tokens.get("refresh_token")
    access_token = tokens.get("access_token")

    if not new_refresh:
        st.error("❌ Geen refresh_token ontvangen van Teamleader.")
        st.write("DEBUG – tokens JSON:", tokens)
        st.stop()

    st.success("✅ Refresh token ontvangen!")
    st.markdown("➡ Zet deze nu in Railway bij `REFRESH_TOKEN`:")
    st.code(new_refresh)

    st.info("Na opslaan in Railway en een nieuwe deploy is de koppeling permanent actief.")
    st.stop()

# ======================================================
# 3. LOGIN-BLOK (als er nog geen REFRESH_TOKEN in env staat)
# ======================================================
if not REFRESH_TOKEN:
    st.warning("⚠️ Je bent nog niet gekoppeld met Teamleader. Verbind hieronder.")

    if not CLIENT_ID or not CLIENT_SECRET:
        st.error(
            "CLIENT_ID of CLIENT_SECRET is leeg in de omgeving.\n\n"
            "Controleer in Railway → Variables of beide variabelen een waarde hebben "
            "en redeploy daarna de app."
        )
    else:
        # Bouw de login-URL dynamisch, met actuele CLIENT_ID en redirect
        params_login = {
            "client_id": CLIENT_ID,
            "redirect_uri": REDIRECT_URI,
            "response_type": "code",
            # Geen offline_access: die gaf eerder invalid_scope in jouw omgeving
            "scope": "companies contacts deals products quotations projects invoices",
        }
        login_url = f"{AUTH_BASE}?{urlencode(params_login)}"

        st.write("DEBUG – login URL:", login_url)
        st.link_button("🔐 Verbind met Teamleader", login_url)
        st.markdown(f"👉 Of klik hier: [{login_url}]({login_url})")

    # Zolang er geen refresh_token is, stoppen we hier
    st.stop()

# ======================================================
# 4. VANAF HIER: TEAMLEADER IS GEKOPPELD → NORMAL APP FLOW
# ======================================================

uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

if uploaded_file:
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.write(uploaded_file.read())
    temp_file.close()

    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_file.name)
    except Exception as e:
        st.error(f"❌ Fout bij uitlezen van Excel: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Onbekend model op basis van G2='{g2}' en H2='{h2}'.")
        st.stop()

    try:
        data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)
    except Exception as e:
        st.error(f"❌ Fout tijdens berekenen van offerte: {e}")
        st.stop()

    st.subheader("Samenvatting")
    st.write(f"**Project:** {data['project']}")
    st.write(f"**Model:** {data['model']} ({data['materiaal']})")
    st.write(f"**Kleur:** {data['kleur']}")
    st.write(f"**Aantal fronten:** {data['fronts']}")
    st.write(f"**Scharnieren:** {data['scharnieren']}")
    st.write(f"**Lades:** {data['lades']}")
    st.write(f"**Totaal excl. btw:** € {data['totaal_excl']:.2f}")
    st.write(f"**Totaal incl. btw:** € {data['totaal_incl']:.2f}")

    st.markdown("---")
    st.subheader("Offerte aanmaken in Teamleader")

    if not deal_id:
        st.info("Vul een deal-ID in om te verzenden naar Teamleader.")
    elif st.button("Maak offerte in Teamleader"):
        try:
            hf.maak_teamleader_offerte(deal_id, data, mode)
            st.success("✅ Offerte succesvol aangemaakt in Teamleader!")
        except Exception as e:
            st.error(f"❌ Er ging iets mis bij het aanmaken van de offerte:\n\n{e}")
else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
