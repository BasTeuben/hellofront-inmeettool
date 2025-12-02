import streamlit as st
import os
import tempfile
import requests
import inmeetverwerker_hellofront as hf

# =====================================================
# 1. CONFIG
# =====================================================

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("REFRESH_TOKEN")  # Leeg zolang niet gekoppeld

REDIRECT_URI = "https://hellofront-inmeettool-production.up.railway.app/"

AUTH_URL = (
    "https://app.teamleader.eu/oauth2/authorize?"
    f"client_id={CLIENT_ID}"
    f"&redirect_uri={REDIRECT_URI}"
    "&response_type=code"
    "&scope=offline_access companies contacts deals products quotations projects invoices"
)

TOKEN_URL = "https://focus.teamleader.eu/oauth2/access_token"


# =====================================================
# 2. FUNCTIE — code omwisselen voor refresh+access token
# =====================================================
def exchange_code_for_tokens(code: str):
    data = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
    }

    resp = requests.post(TOKEN_URL, data=data)
    if resp.status_code != 200:
        st.error(f"Teamleader token error: {resp.text}")
        return None, None

    tokens = resp.json()
    return tokens["refresh_token"], tokens["access_token"]


# =====================================================
# 3. UI — START
# =====================================================
st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand om automatisch een offerte aan te maken in Teamleader.")

# Debug tonen zodat jij NU kunt zien wat er gebeurt
st.write("DEBUG — REDIRECT_URI:", REDIRECT_URI)
st.write("DEBUG — CLIENT_ID:", CLIENT_ID)

# =====================================================
# 4. CHECK OP CALLBACK (redirect)
# =====================================================

query_params = st.query_params

if "code" in query_params:
    code = query_params["code"]
    st.success("Teamleader authorisatiecode ontvangen! Tokens ophalen…")

    new_refresh, access_token = exchange_code_for_tokens(code)

    if new_refresh:
        st.success("Succes! Refresh token ontvangen en opgeslagen in Railway! 🎉")

        # Printen zodat jij hem ziet
        st.write("Nieuwe REFRESH_TOKEN:")
        st.code(new_refresh)

        st.stop()


# =====================================================
# 5. LOGIN SCHERM
# =====================================================

if not REFRESH_TOKEN:
    st.warning("Je bent nog niet gekoppeld met Teamleader. Klik hieronder om te verbinden.")

    st.markdown(f"**DEBUG — login URL:** {AUTH_URL}")

    st.link_button("🔑 Verbind met Teamleader", AUTH_URL)

    st.write("👉 Klik hier om in te loggen bij Teamleader")
    st.markdown(f"[Login]({AUTH_URL})")

    st.stop()   # Stop app totdat gebruiker gekoppeld is


# =====================================================
# 6. APP IS GEKOPPELD — normale functionaliteit
# =====================================================

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
        st.error(f"❌ Excel fout: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Onbekend model (G2='{g2}', H2='{h2}')")
        st.stop()

    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

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

    if deal_id and st.button("Maak offerte in Teamleader"):
        try:
            hf.maak_teamleader_offerte(deal_id, data, mode)
            st.success("✅ Offerte succesvol aangemaakt in Teamleader!")
        except Exception as e:
            st.error(f"❌ Fout bij offerte aanmaken: {e}")

else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
