# app.py
# Dit is het "scherm" dat jouw collega's gebruiken in de browser.

import streamlit as st
import os
import inmeetverwerker_hellofront as hf

st.set_page_config(page_title="HelloFront Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand en maak een offerte aan in Teamleader.")

# ======================================================
# 1. Excel uploaden
# ======================================================

uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

# ======================================================
# 2. Soort offerte
# ======================================================

offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

# ======================================================
# 3. Deal-ID invoeren
# ======================================================

deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

# ======================================================
# Als er een bestand is geüpload → uitlezen
# ======================================================

if uploaded_file is not None:

    temp_filename = "upload_inmeet.xlsx"
    with open(temp_filename, "wb") as f:
        f.write(uploaded_file.read())

    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_filename)
    except Exception as e:
        st.error(f"❌ Fout bij uitlezen Excel-bestand: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Onbekend model op basis van G2='{g2}' en H2='{h2}'.")
        st.stop()

    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

    # ==================================================
    # SAMENVATTING TONEN
    # ==================================================
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
        st.info("Vul eerst een Teamleader deal-ID in.")

    # ==================================================
    # OFFERTES VERSTUREN
    # ==================================================
    if deal_id and st.button("Maak offerte in Teamleader"):

        try:
            response_text = hf.maak_teamleader_offerte(deal_id, data, mode)
        except Exception as e:
            st.error(f"❌ Fout bij aanmaken van de offerte: {e}")
            st.stop()

        st.markdown("### 🔍 Technische response van Teamleader")
        st.code(response_text)

       if '"type":"quotation"' in response_text or '"type": "quotation"' in response_text:
            st.success("✅ Offerte is aangemaakt in Teamleader! Controleer de deal in Teamleader.")
        else:
            st.error("❌ Teamleader gaf geen geldige bevestiging. Controleer de response hierboven.")

else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
