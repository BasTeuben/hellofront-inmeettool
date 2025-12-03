import streamlit as st
import os
import tempfile
import requests
import inmeetverwerker_hellofront as hf

# ======================================================
# 1. BASISCONFIG
# ======================================================

CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("REFRESH_TOKEN", "")  # Moet een geldige token bevatten!

# Als REFRESH_TOKEN leeg is → app kan niet werken
if not REFRESH_TOKEN:
    st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")
    st.title("HelloFront – Inmeet Tool")
    st.error(
        "❌ De Teamleader koppeling is nog niet geconfigureerd.\n\n"
        "Er is **geen REFRESH_TOKEN** gevonden in Railway.\n"
        "Voer de eenmalige OAuth-koppeling uit en vul daarna de REFRESH_TOKEN in Railway in."
    )
    st.stop()

st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand om automatisch een offerte aan te maken in Teamleader.")

# ======================================================
# 2. GEEN OAUTH MEER IN PRODUCTIE
# ======================================================
# Deze hele sectie is weggehaald zodat collega's nooit per ongeluk
# refresh tokens overschrijven of de koppeling opnieuw starten.

# ======================================================
# 3. VANAF HIER: APP IS GEKOPPELD → NORMAL APP FLOW
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

    # Excel uitlezen
    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_file.name)
    except Exception as e:
        st.error(f"❌ Fout bij uitlezen van Excel: {e}")
        st.stop()

    # Model bepalen
    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Onbekend model op basis van G2='{g2}' en H2='{h2}'.")
        st.stop()

    # Berekeningen uitvoeren
    try:
        data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)
    except Exception as e:
        st.error(f"❌ Fout tijdens berekenen van offerte: {e}")
        st.stop()

    # Samenvatting tonen
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
