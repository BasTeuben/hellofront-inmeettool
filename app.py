import streamlit as st
import tempfile
import os

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

data = None  # zodat we weten of er al een berekening is

if uploaded_file is not None:
    # Sla het Excel-bestand tijdelijk op
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        temp_filename = tmp.name
        tmp.write(uploaded_file.getbuffer())

    # Gebruik je bestaande functies
    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_filename)
    except Exception as e:
        st.error(f"❌ Er ging iets mis bij het uitlezen van het Excel-bestand: {e}")
        st.stop()

    # Model bepalen
    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Model onbekend op basis van G2='{g2}' en H2='{h2}'. Controleer het Excel-bestand.")
        st.stop()

    # Bereken de offerte met jouw bestaande logica
    try:
        data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)
    except Exception as e:
        st.error(f"❌ Er ging iets mis bij het berekenen van de offerte: {e}")
        st.stop()

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
            status_code, resp_text = hf.maak_teamleader_offerte(deal_id, data, mode)
        except Exception as e:
            st.error(f"❌ Er ging iets mis bij het aanmaken van de offerte (technische fout): {e}")
        else:
            if status_code in (200, 201):
                st.success("✅ Offerte is aangemaakt in Teamleader. Controleer Teamleader voor de details.")
            else:
                st.error(
                    f"❌ Teamleader gaf een foutmelding (status {status_code}). "
                    f"Details:\n\n```{resp_text}```"
                )
else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
