import streamlit as st
import os
import inmeetverwerker_hellofront as hf

st.set_page_config(page_title="HelloFront Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand en maak een offerte aan in Teamleader.")

uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

if uploaded_file is not None:

    temp_filename = "upload_inmeet.xlsx"
    with open(temp_filename, "wb") as f:
        f.write(uploaded_file.read())

    try:
        onderdelen, g2, h2, kleur, klantregels, scharnieren, lades, project = hf.lees_excel(temp_filename)
    except Exception as e:
        st.error(f"❌ Fout bij uitlezen Excel: {e}")
        st.stop()

    model = hf.bepaal_model(g2, h2)
    if not model:
        st.error(f"❌ Model onbekend op basis van: G2='{g2}' en H2='{h2}'")
        st.stop()

    data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)

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
        st.info("Voer een Teamleader deal-ID in.")
    else:
        if st.button("Maak offerte in Teamleader"):

            resp = hf.maak_teamleader_offerte(deal_id, data, mode)

            st.subheader("🔍 Technische response van Teamleader")
            st.code(resp.text)

            if resp.status_code in (200, 201):
                st.success("✅ Offerte is succesvol aangemaakt in Teamleader!")
            else:
                st.error("❌ Teamleader gaf geen geldige bevestiging. Controleer de response hierboven.")
