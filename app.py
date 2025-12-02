import streamlit as st
import os
import tempfile
import requests
from urllib.parse import urlencode
import inmeetverwerker_hellofront as hf

# ======================================================
# ⚙️ STREAMLIT CONFIG
# ======================================================
st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand om automatisch een offerte aan te maken in Teamleader.")

# ======================================================
# 🔧 TEAMLEADER CONFIG
# ======================================================
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("REFRESH_TOKEN")   # bij eerste login: leeg

# 🔥 Debug redirect: zeer belangrijk om de invalid_client fout op te sporen
REDIRECT_URI = "https://hellofront-inmeettool-production.up.railway.app/"

st.write("DEBUG — REDIRECT_URI (zoals code denkt):", REDIRECT_URI)

# Teamleader OAuth endpoints
AUTH_BASE = "https://app.teamleader.eu/oauth2/authorize"
TOKEN_URL = "https://focus.teamleader.eu/oauth2/access_token"


# ======================================================
# 🔐 1. HANDLE TEAMLEADER CALLBACK VIA QUERY PARAM
# ======================================================
params = st.experimental_get_query_params()

if "code" in params:
    st.subheader("Teamleader koppeling")

    code = params["code"][0]
    st.info("Bezig met ophalen van tokens…")

    # Token exchange — Teamleader Focus geeft refresh_token zonder extra scopes
    resp = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": REDIRECT_URI,
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
        }
    )

    st.write("DEBUG — token exchange response:", resp.text)

    if resp.status_code != 200:
        st.error(f"❌ Ophalen tokens mislukt:\n\n{resp.text}")
        st.stop()

    tokens = resp.json()
    refresh_token = tokens.get("refresh_token")

    if not refresh_token:
        st.error("❌ Geen refresh_token ontvangen van Teamleader. Controleer app-instellingen.")
        st.stop()

    st.success("✅ Refresh token ontvangen!")
    st.markdown("**➡️ Zet deze nu in Railway → Variables → `REFRESH_TOKEN`**")
    st.code(refresh_token)

    st.info("Ververs de pagina — de koppeling is actief.")
    st.stop()


# ======================================================
# 🔐 2. LOGIN KNOP (alleen zichtbaar als REFRESH_TOKEN leeg is)
# ======================================================
if not REFRESH_TOKEN:
    st.warning("⚠️ Je bent nog niet gekoppeld met Teamleader. Klik hieronder om te verbinden.")

    if st.button("🔐 Verbind met Teamleader"):
        # laat zien wat we nu écht meesturen naar Teamleader
        st.write("DEBUG — redirect die we meesturen:", REDIRECT_URI)

        params = {
            "client_id": CLIENT_ID,
            "redirect_uri": REDIRECT_URI,
            "response_type": "code",
            "scope": "companies contacts deals products quotations projects invoices",
        }

        login_url = f"{AUTH_BASE}?{urlencode(params)}"
        st.write("DEBUG — login URL:", login_url)

        st.markdown(f"[👉 Klik hier om in te loggen bij Teamleader]({login_url})")

    st.stop()


# ======================================================
# 🟢 3. TEAMLEADER IS VERBONDEN → START REGULIERE TOOL
# ======================================================
uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")


# ======================================================
# 📥 EXCEL VERWERKEN
# ======================================================
if uploaded_file:

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.write(uploaded_file.read())
    temp_file.close()

    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_file.name)
    except Exception as e:
        st.error(f"❌ Fout bij uitlezen Excel: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Onbekend model voor G2='{g2}', H2='{h2}'.")
        st.stop()

    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

    st.subheader("Samenvatting")
    st.write(f"**Project:** {data['project']}")
    st.write(f"**Model:** {data['model']} ({data['materiaal']})")
    st.write(f"**Kleur:** {data['kleur']}")
    st.write(f"**Fronten:** {data['fronts']}")
    st.write(f"**Scharnieren:** {data['scharnieren']}")
    st.write(f"**Lades:** {data['lades']}")
    st.write(f"**Totaal excl. btw:** € {data['totaal_excl']:.2f}")
    st.write(f"**Totaal incl. btw:** € {data['totaal_incl']:.2f}")

    st.markdown("---")
    st.subheader("Offerte aanmaken in Teamleader")

    if not deal_id:
        st.info("Vul een deal-ID in om te verzenden naar Teamleader.")

    if deal_id and st.button("Maak offerte in Teamleader"):
        try:
            hf.maak_teamleader_offerte(deal_id, data, mode)
            st.success("✅ Offerte succesvol aangemaakt in Teamleader!")
        except Exception as e:
            st.error(f"❌ Fout bij verwerken in Teamleader:\n\n{e}")

else:
    st.info("Upload een Excel-bestand om te beginnen.")
