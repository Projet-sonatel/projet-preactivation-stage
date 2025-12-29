import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Orange NFC Synthesis", layout="wide")

st.title("📊 Synthèse NFC par Direction Régionale")

# Mapping officiel fourni
DR_MAPPING = {
    'DV-DRVE_DIRECTION REGIONALE DES VENTES EST': 'DRE',
    'DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE': 'DRC',
    'DV-DRVN_DIRECTION REGIONALE DES VENTES NORD': 'DRN',
    'DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST': 'DRSE',
    'DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2': 'DR2',
    'DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1': 'DR1',
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS"
}

uploaded_file = st.file_uploader("Déposez le fichier WEEKLY STAT NFC", type=["csv", "xlsx", "xlsb"])

if uploaded_file:
    try:
        # 1. Lecture du fichier
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, sep=';')
        elif uploaded_file.name.endswith('.xlsb'):
            df = pd.read_excel(uploaded_file, engine='pyxlsb')
        else:
            df = pd.read_excel(uploaded_file)

        # Nettoyage des noms de colonnes
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Filtrage et Renommage des DR
        # On utilise la colonne 'AGENCE' comme indiqué
        if 'AGENCE' in df.columns:
            # On ne garde que les lignes dont l'AGENCE est dans notre mapping
            df = df[df['AGENCE'].isin(DR_MAPPING.keys())].copy()
            # On renomme avec les abréviations
            df['DR'] = df['AGENCE'].map(DR_MAPPING)
        else:
            st.error("La colonne 'AGENCE' est introuvable dans le fichier.")
            st.stop()

        # 3. Calcul de la Synthèse par DR
        # On groupe par la nouvelle colonne 'DR' et on somme les valeurs
        synthese = df.groupby('DR').agg({
            'OPERATION NFC': 'sum',
            'OPERATION MANUELLE': 'sum',
            'TOTAL OPERATION': 'sum'
        }).reset_index()

        # 4. Calcul du TAUX NFC
        # Formule : (NFC / TOTAL) * 100
        synthese['TAUX NFC'] = (synthese['OPERATION NFC'] / synthese['TOTAL OPERATION']) * 100

        # Tri par performance (Optionnel)
        synthese = synthese.sort_values('TAUX NFC', ascending=False)

        # 5. Affichage du tableau style "Reporting"
        st.subheader("📈 Tableau de Synthèse par DR")

        # Formatage pour l'affichage Streamlit (pour voir les % proprement)
        st.dataframe(
            synthese.style.format({'TAUX NFC': '{:.2f}%'}),
            use_container_width=True
        )

        # 6. Export EXCEL avec mise en forme
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            synthese.to_excel(writer, sheet_name='SYNTHESE DR', index=False)

            workbook = writer.book
            worksheet = writer.sheets['SYNTHESE DR']

            # Formats
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'
            })
            num_format = workbook.add_format({'border': 1, 'align': 'center'})
            pct_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00"%"'})

            # Appliquer les formats aux en-têtes
            for col_num, value in enumerate(synthese.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)

            # Appliquer le format pourcentage à la dernière colonne
            for row_num in range(1, len(synthese) + 1):
                worksheet.write(row_num, 4, synthese.iloc[row_num-1, 4] / 100, pct_format)

        st.download_button(
            label="📥 Télécharger la Synthèse DR (Excel)",
            data=output.getvalue(),
            file_name="Synthese_NFC_par_DR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erreur : {e}")