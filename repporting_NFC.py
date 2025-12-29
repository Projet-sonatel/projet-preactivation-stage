import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Orange NFC - Reporting Officiel", layout="wide")

st.title("📊 Reporting NFC : Synthèse & Détail DR-SADI-RAVT")

# Mapping officiel des DR
DR_MAPPING = {
    'DV-DRVE_DIRECTION REGIONALE DES VENTES EST': 'DRE',
    'DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE': 'DRC',
    'DV-DRVN_DIRECTION REGIONALE DES VENTES NORD': 'DRN',
    'DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST': 'DRSE',
    'DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2': 'DR2',
    'DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1': 'DR1',
    'DV-DRVS_DIRECTION REGIONALE DES VENTES SUD': 'DRS'
}

col1, col2 = st.columns(2)
with col1:
    ref_file = st.file_uploader("1. Déposez le RÉFÉRENTIEL (Mapping)", type=["csv", "xlsx"])
with col2:
    weekly_file = st.file_uploader("2. Déposez le fichier WEEKLY STAT NFC", type=["csv", "xlsx", "xlsb"])

if ref_file and weekly_file:
    try:
        # --- 1. LECTURE ET NETTOYAGE DU RÉFÉRENTIEL ---
        df_ref = pd.read_csv(ref_file) if ref_file.name.endswith('.csv') else pd.read_excel(ref_file)
        df_ref.columns = [str(c).strip() for c in df_ref.columns]
        # On garde une seule ligne par LOGIN pour ne pas multiplier les stats
        df_ref = df_ref[['LOGIN', 'SADI', 'RAVT']].drop_duplicates(subset=['LOGIN'])

        # --- 2. LECTURE DU WEEKLY ---
        if weekly_file.name.endswith('.csv'):
            df_weekly = pd.read_csv(weekly_file, sep=';')
        elif weekly_file.name.endswith('.xlsb'):
            df_weekly = pd.read_excel(weekly_file, engine='pyxlsb')
        else:
            df_weekly = pd.read_excel(weekly_file)

        df_weekly.columns = [str(c).strip() for c in df_weekly.columns]

        # --- 3. TRAITEMENT ---
        # Filtrage et renommage des DR initial
        df_weekly = df_weekly[df_weekly['AGENCE'].isin(DR_MAPPING.keys())].copy()
        df_weekly['DR'] = df_weekly['AGENCE'].map(DR_MAPPING)

        # Jointure INNER pour ne garder que ce qui est mappé (Supprime les "Inconnus")
        df_final = pd.merge(df_weekly, df_ref, on='LOGIN', how='inner')

        # Nettoyage strict des lignes vides ou sans SADI/RAVT
        df_final = df_final.dropna(subset=['SADI', 'RAVT'])
        df_final = df_final[(df_final['SADI'].astype(str).str.strip() != "") &
                            (df_final['RAVT'].astype(str).str.strip() != "")]

        # CORRECTION : Garder seulement le SADI qui correspond au DR du LOGIN
        # Cela évite qu'un SADI apparaisse dans plusieurs DR
        df_final = df_final.drop_duplicates(subset=['LOGIN', 'SADI', 'RAVT', 'DR'])

        # Nettoyer les valeurs numériques nulles ou invalides
        df_final = df_final[
            (df_final['OPERATION NFC'].notna()) &
            (df_final['OPERATION MANUELLE'].notna()) &
            (df_final['TOTAL OPERATION'].notna())
        ]

        # --- 4. GÉNÉRATION EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # FORMATS
            h_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})
            dr_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            sadi_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'indent': 1})
            ravt_fmt = workbook.add_format({'border': 1, 'indent': 2})
            num_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            taux_fmt = workbook.add_format({'border': 1, 'num_format': '0.00', 'align': 'center'})

            headers = ['DR', 'OP NFC', 'OP MANUELLE', 'TOTAL', 'Taux']

            # --- FEUILLE 1 : SYNTHESE DR ---
            ws1 = workbook.add_worksheet('SYNTHESE DR')
            for c, h in enumerate(headers): ws1.write(0, c, h, h_fmt)

            synthese_dr = df_final.groupby('DR').agg({
                'OPERATION NFC': 'sum', 'OPERATION MANUELLE': 'sum', 'TOTAL OPERATION': 'sum'
            }).reset_index()

            for i, r in synthese_dr.iterrows():
                ws1.write(i+1, 0, r['DR'], num_fmt)
                ws1.write(i+1, 1, r['OPERATION NFC'], num_fmt)
                ws1.write(i+1, 2, r['OPERATION MANUELLE'], num_fmt)
                ws1.write(i+1, 3, r['TOTAL OPERATION'], num_fmt)
                t_val = (r['OPERATION NFC'] / r['TOTAL OPERATION'] * 100) if r['TOTAL OPERATION'] > 0 else 0
                ws1.write(i+1, 4, t_val, taux_fmt)
            ws1.set_column('A:E', 18)

            # --- FEUILLE 2 : REPORTING DR-SADI-RAVT ---
            ws2 = workbook.add_worksheet('REPORTING DR-SADI-RAVT')
            for c, h in enumerate(headers): ws2.write(0, c, h, h_fmt)

            curr_row = 1
            # On trie par DR, SADI, RAVT pour une cascade parfaite
            for dr, dr_group in df_final.groupby('DR', sort=True):
                # Vérifier que le groupe n'est pas vide
                if len(dr_group) == 0 or dr_group['TOTAL OPERATION'].sum() == 0:
                    continue

                # Ligne DR
                n_dr, m_dr, t_dr = dr_group['OPERATION NFC'].sum(), dr_group['OPERATION MANUELLE'].sum(), dr_group['TOTAL OPERATION'].sum()
                ws2.write(curr_row, 0, dr, dr_fmt)
                ws2.write(curr_row, 1, n_dr, dr_fmt)
                ws2.write(curr_row, 2, m_dr, dr_fmt)
                ws2.write(curr_row, 3, t_dr, dr_fmt)
                ws2.write(curr_row, 4, (n_dr/t_dr*100) if t_dr > 0 else 0, dr_fmt)
                curr_row += 1

                for sadi, sadi_group in dr_group.groupby('SADI', sort=True):
                    # Vérifier que le groupe n'est pas vide
                    if len(sadi_group) == 0 or sadi_group['TOTAL OPERATION'].sum() == 0:
                        continue

                    # Ligne SADI
                    n_s, m_s, t_s = sadi_group['OPERATION NFC'].sum(), sadi_group['OPERATION MANUELLE'].sum(), sadi_group['TOTAL OPERATION'].sum()
                    ws2.write(curr_row, 0, sadi, sadi_fmt)
                    ws2.write(curr_row, 1, n_s, sadi_fmt)
                    ws2.write(curr_row, 2, m_s, sadi_fmt)
                    ws2.write(curr_row, 3, t_s, sadi_fmt)
                    ws2.write(curr_row, 4, (n_s/t_s*100) if t_s > 0 else 0, sadi_fmt)
                    curr_row += 1

                    for ravt, ravt_group in sadi_group.groupby('RAVT', sort=True):
                        # Vérifier que le groupe n'est pas vide
                        if len(ravt_group) == 0 or ravt_group['TOTAL OPERATION'].sum() == 0:
                            continue

                        # Ligne RAVT
                        n_r, m_r, t_r = ravt_group['OPERATION NFC'].sum(), ravt_group['OPERATION MANUELLE'].sum(), ravt_group['TOTAL OPERATION'].sum()
                        ws2.write(curr_row, 0, ravt, ravt_fmt)
                        ws2.write(curr_row, 1, n_r, ravt_fmt)
                        ws2.write(curr_row, 2, m_r, ravt_fmt)
                        ws2.write(curr_row, 3, t_r, ravt_fmt)
                        ws2.write(curr_row, 4, (n_r/t_r*100) if t_r > 0 else 0, ravt_fmt)
                        curr_row += 1

            # Appliquer la largeur des colonnes SANS format par défaut
            ws2.set_column('A:A', 45)
            ws2.set_column('B:D', 15)
            ws2.set_column('E:E', 15)

            # --- FEUILLE 3 : REPORTING DR-RAVT-PVT ---
            ws3 = workbook.add_worksheet('REPORTING DR-RAVT-PVT')
            for c, h in enumerate(headers): ws3.write(0, c, h, h_fmt)

            # IMPORTANT : Filtrer uniquement les PVT pour cette feuille
            df_pvt = df_final[df_final['ACCUEIL'].astype(str).str.startswith('PVT')].copy()

            curr_row = 1
            # On trie par DR, RAVT, PVT (ACCUEIL)
            for dr, dr_group in df_pvt.groupby('DR', sort=True):
                # Vérifier que le groupe n'est pas vide
                if len(dr_group) == 0 or dr_group['TOTAL OPERATION'].sum() == 0:
                    continue

                # Ligne DR
                n_dr, m_dr, t_dr = dr_group['OPERATION NFC'].sum(), dr_group['OPERATION MANUELLE'].sum(), dr_group['TOTAL OPERATION'].sum()
                ws3.write(curr_row, 0, dr, dr_fmt)
                ws3.write(curr_row, 1, n_dr, dr_fmt)
                ws3.write(curr_row, 2, m_dr, dr_fmt)
                ws3.write(curr_row, 3, t_dr, dr_fmt)
                ws3.write(curr_row, 4, (n_dr/t_dr*100) if t_dr > 0 else 0, dr_fmt)
                curr_row += 1

                for ravt, ravt_group in dr_group.groupby('RAVT', sort=True):
                    # Vérifier que le groupe n'est pas vide
                    if len(ravt_group) == 0 or ravt_group['TOTAL OPERATION'].sum() == 0:
                        continue

                    # Ligne RAVT
                    n_r, m_r, t_r = ravt_group['OPERATION NFC'].sum(), ravt_group['OPERATION MANUELLE'].sum(), ravt_group['TOTAL OPERATION'].sum()
                    ws3.write(curr_row, 0, ravt, sadi_fmt)
                    ws3.write(curr_row, 1, n_r, sadi_fmt)
                    ws3.write(curr_row, 2, m_r, sadi_fmt)
                    ws3.write(curr_row, 3, t_r, sadi_fmt)
                    ws3.write(curr_row, 4, (n_r/t_r*100) if t_r > 0 else 0, sadi_fmt)
                    curr_row += 1

                    for pvt, pvt_group in ravt_group.groupby('ACCUEIL', sort=True):
                        # Vérifier que le groupe n'est pas vide
                        if len(pvt_group) == 0 or pvt_group['TOTAL OPERATION'].sum() == 0:
                            continue

                        # Ligne PVT (ACCUEIL)
                        n_p, m_p, t_p = pvt_group['OPERATION NFC'].sum(), pvt_group['OPERATION MANUELLE'].sum(), pvt_group['TOTAL OPERATION'].sum()
                        ws3.write(curr_row, 0, pvt, ravt_fmt)
                        ws3.write(curr_row, 1, n_p, ravt_fmt)
                        ws3.write(curr_row, 2, m_p, ravt_fmt)
                        ws3.write(curr_row, 3, t_p, ravt_fmt)
                        ws3.write(curr_row, 4, (n_p/t_p*100) if t_p > 0 else 0, ravt_fmt)
                        curr_row += 1

            # Appliquer la largeur des colonnes
            ws3.set_column('A:A', 45)
            ws3.set_column('B:D', 15)
            ws3.set_column('E:E', 15)

        st.success("✅ Fichier corrigé généré avec succès !")
        st.download_button("📥 Télécharger le Reporting Final", output.getvalue(), "Reporting_NFC_Orange_Final.xlsx")

    except Exception as e:
        st.error(f"Erreur : {e}")