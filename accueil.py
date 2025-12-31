import streamlit as st

st.set_page_config(
    page_title="Orange Tools",
    page_icon="🍊",
    layout="wide"
)

st.title("🍊 Outils Orange - Bienvenue")

st.markdown("""
## Sélectionnez un outil dans le menu latéral 👈

### Applications disponibles :

**1. 🚀 Préactivations**
- Génération de reporting des préactivations
- Tri sélectif : Clôtures avec Statut / Rejets

**2. 📊 Classement PVT**
- Classement des 7 Directions Régionales
- Analyse des ventes PVT

**3. 📈 Reporting NFC**
- Génération de rapports NFC
- Analyse des données

---

👈 **Utilisez le menu latéral pour accéder aux outils**
""")

st.info("💡 Astuce : Vous pouvez basculer entre les outils en utilisant le menu de navigation à gauche.")