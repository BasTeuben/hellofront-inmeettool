# app.py
# Streamlit webinterface voor jouw bestaande inmeetverwerker_hellofront.py

import streamlit as st
import os
import tempfile
import json
import inmeetverwerker_hellofront as hf

st.set_page_config(page_title="HelloFront Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand en maak automatisch een offerte aan in Teamleader.")

# -----------------------------
# 1. Excel uploaden
# -----------------------------
uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

# -----------------------------
# 2. Offerte type kiezen
# -----------------------------
offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

# -----------------------------
# 3. Deal-ID
# -----------------------------
deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

# ------------------------------------------
# Wanneer gebruiker een bestand uploadt
# ------------------------------------------
if uploaded_file is not None:

    # Sla het bestand tijdelijk op
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        temp_path = tmp.name

    # Lees Excel met jouw bestaande functie
    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_path)
    except Exception as e:
        st.error(f"❌ Fout bij inlezen Excel: {e}")
        st.stop()

    # Model bepalen
    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Model onbekend: G2='{g2}', H2='{h2}'.")
        st.stop()

    # Offerte berekenen
    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

    # ------------------------------------------
    # Samenvatting tonen
    # ------------------------------------------
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
        st.info("Vul boven een deal-ID in om verder te gaan.")

    # ------------------------------------------
    # Knop → Offerte naar Teamleader sturen
    # ------------------------------------------
    if deal_id and st.button("Maak offerte in Teamleader"):

        with st.spinner("Bezig met versturen naar Teamleader..."):
            try:
                resp = hf.maak_teamleader_offerte(deal_id, data, mode)

                # resp kan None zijn → fout
                if resp is None:
                    st.error("❌ Geen geldige response van Teamleader.")
                    st.stop()

                # Response tonen
                st.markdown("### 🔍 Technische response van Teamleader")
                response_text = resp.text
                st.code(response_text)

                # -------------------------------
                # **GOEDE CHECK**
                # Teamleader geeft bij succes terug:
                # {"data":{"type":"quotation","id":"..."}}
                # -------------------------------
                if '"type":"quotation"' in response_text:
                    st.success("✅ Offerte is aangemaakt in Teamleader! Controleer de deal in Teamleader.")
                else:
                    st.error("❌ Teamleader gaf geen geldige bevestiging. Controleer de response hierboven.")

            except Exception as e:
                st.error(f"❌ Er ging iets mis: {e}")

else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
