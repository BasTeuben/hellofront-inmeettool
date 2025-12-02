import streamlit as st
import os
import tempfile
import inmeetverwerker_hellofront as hf

# ------------------------------------------------------
# ⚙️ Streamlit Config
# ------------------------------------------------------
st.set_page_config(page_title="HelloFront – Inmeet Tool", layout="centered")

st.title("HelloFront – Inmeet Tool")
st.write("Upload een Excel-bestand om automatisch een offerte aan te maken in Teamleader.")

# ------------------------------------------------------
# 📤 FILE UPLOADER
# ------------------------------------------------------
uploaded_file = st.file_uploader("Kies een Excel-bestand (.xlsx)", type=["xlsx"])

# Particulier / Dealer keuze
offerte_type = st.radio("Soort offerte", ["Particulier", "Dealer"])
mode = "P" if offerte_type == "Particulier" else "D"

# Deal-ID
deal_id = st.text_input("Teamleader deal-ID")

st.markdown("---")

# ------------------------------------------------------
# 📥 Verwerking
# ------------------------------------------------------
if uploaded_file:

    # Tijdelijk bestand opslaan
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.write(uploaded_file.read())
    temp_file.close()

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

    # Offerte berekenen
    try:
        data = hf.bereken_offerte(onderdelen, model, project, kleur, klantregels, scharnieren, lades)
    except Exception as e:
        st.error(f"❌ Fout tijdens berekenen van offerte: {e}")
        st.stop()

    # ------------------------------------------------------
    # 📊 SAMENVATTING
    # ------------------------------------------------------
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

    # ------------------------------------------------------
    # 📤 Versturen naar Teamleader
    # ------------------------------------------------------
    if deal_id and st.button("Maak offerte in Teamleader"):
        try:
            hf.maak_teamleader_offerte(deal_id, data, mode)
            st.success("✅ Offerte succesvol aangemaakt in Teamleader!")
        except Exception as e:
            st.error(f"❌ Er ging iets mis bij het aanmaken van de offerte (technische fout): {e}")

else:
    st.info("Upload hierboven een Excel-bestand om te beginnen.")
