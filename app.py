# app.py
# Dit is het "scherm" in de browser.
# Dit bestand gebruikt jouw bestaande code uit inmeetverwerker_hellofront.py

import streamlit as st
import os

# === Teamleader secrets ophalen vanuit Streamlit Cloud ===
CLIENT_ID = st.secrets.get("CLIENT_ID")
CLIENT_SECRET = st.secrets.get("CLIENT_SECRET")

# We geven deze waarden door aan je bestaande script via environment variables
if CLIENT_ID:
    os.environ["CLIENT_ID"] = CLIENT_ID
if CLIENT_SECRET:
    os.environ["CLIENT_SECRET"] = CLIENT_SECRET

# We halen de functies uit jouw bestaande script
import inmeetverwerker_hellofront as hf

st.set_page_config(page_title="HelloFront Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand en maak een offerte aan in Teamleader.")

# 1. Excel uploaden
uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

# 2. Kiezen: particulier of dealer
offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

# 3. Deal-ID voor Teamleader
deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

if uploaded_file is not None:
    # Sla het Excel-bestand tijdelijk op
    temp_filename = "upload_inmeet.xlsx"
    with open(temp_filename, "wb") as f:
        f.write(uploaded_file.read())

    # Gebruik je bestaande functies
    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_filename)
    except Exception as e:
        st.error(f"❌ Er ging iets mis bij het uitlezen van het Excel-bestand: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Model onbekend op basis van G2='{g2}' en H2='{h2}'. Controleer het Excel-bestand.")
        st.stop()

    # Bereken de offerte met jouw bestaande logica
    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

    # Samenvatting laten zien
    st.subheader("Samenvatting van de offerte")
    st.write(f"**Project:** {data['project']}")
    st.write(f"**Model:** {data['model']} ({data['materiaal']})")
    st.write(f"**Kleur:** {data['kleur']}")
    st.write(f"**Aantal fronten:** {data['fronts']}")
    st.write(f"**Scharnieren:** {data['scharnieren']}")
    st.write(f"**Maatwerk lades:** {data['lades']}")
    st.write(f"**Totaal excl. btw:** € {data['totaal_excl']:.2f}")
    st.write(f"**Totaal incl. btw:** € {data['totaal_incl']:.2f}")

    st.markdown("---")
    st.subheader("Offerte aanmaken in Teamleader")

    if not deal_id:
        st.info("Vul een Teamleader deal-ID in om de offerte te kunnen aanmaken.")

    # Knop om naar Teamleader te sturen
    if deal_id and st.button("Maak offerte in Teamleader"):
        try:
            response_text = hf.maak_teamleader_offerte(deal_id, data, mode)

            st.markdown("### 🔎 Technische response van Teamleader")
            st.code(str(response_text), language="json")

            if response_text is None:
                st.error("❌ Teamleader gaf geen geldige response terug. Waarschijnlijk ontbreekt er een geldig token op de server.")
            else:
                st.success("✅ Aanroep naar Teamleader is uitgevoerd. Controleer de response hierboven en in Teamleader zelf.")
        except Exception as e:
            st.error(f"❌ Er ging iets mis bij het aanmaken van de offerte: {e}")

else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
