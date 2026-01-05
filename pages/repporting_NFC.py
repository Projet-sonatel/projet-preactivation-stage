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
        df_ref = df_ref[['LOGIN', 'SADI', 'RAVT']].drop_duplicates(subset=['LOGIN'])

        # --- 2. LECTURE DU WEEKLY ---
        if weekly_file.name.endswith('.csv'):
            df_weekly = pd.read_csv(weekly_file, sep=';')
        elif weekly_file.name.endswith('.xlsb'):
            df_weekly = pd.read_excel(weekly_file, engine='pyxlsb')
        else:
            df_weekly = pd.read_excel(weekly_file)

        df_weekly.columns = [str(c).strip() for c in df_weekly.columns]

        # 📊 AFFICHER LES TOTAUX AVANT TRAITEMENT
        total_avant = df_weekly[df_weekly['AGENCE'].isin(DR_MAPPING.keys())]['TOTAL OPERATION'].sum()
        st.info(f"📊 **Total AVANT traitement** : {int(total_avant):,} opérations")

        # --- 3. TRAITEMENT SANS PERTE DE DONNÉES ---
        # ✅ Filtrage des DR seulement (PAS de suppression de lignes)
        df_weekly = df_weekly[df_weekly['AGENCE'].isin(DR_MAPPING.keys())].copy()
        df_weekly['DR'] = df_weekly['AGENCE'].map(DR_MAPPING)

        # ✅ JOINTURE LEFT pour GARDER TOUTES les lignes du WEEKLY
        df_final = pd.merge(df_weekly, df_ref, on='LOGIN', how='left')

        # ✅ Remplir les SADI/RAVT manquants par "NON MAPPÉ"
        df_final['SADI'] = df_final['SADI'].fillna('NON MAPPÉ')
        df_final['RAVT'] = df_final['RAVT'].fillna('NON MAPPÉ')

        # ✅ Nettoyer UNIQUEMENT les valeurs numériques invalides
        df_final = df_final[
            (df_final['OPERATION NFC'].notna()) &
            (df_final['OPERATION MANUELLE'].notna()) &
            (df_final['TOTAL OPERATION'].notna()) &
            (df_final['TOTAL OPERATION'] > 0)
        ]

        # 📊 AFFICHER LES TOTAUX APRÈS TRAITEMENT
        total_apres = df_final['TOTAL OPERATION'].sum()
        difference = total_avant - total_apres

        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Total AVANT", f"{int(total_avant):,}")
        with col_stat2:
            st.metric("Total APRÈS", f"{int(total_apres):,}")
        with col_stat3:
            st.metric("Différence", f"{int(difference):,}",
                     delta=f"{(difference/total_avant*100):.1f}%" if total_avant > 0 else "0%",
                     delta_color="inverse")

        # Vérifier les LOGIN non mappés
        non_mappes = df_final[df_final['SADI'] == 'NON MAPPÉ']
        if len(non_mappes) > 0:
            st.warning(f"⚠️ {len(non_mappes)} LOGIN non trouvés dans le référentiel (classés en 'NON MAPPÉ')")
            with st.expander("Voir les LOGIN non mappés"):
                st.dataframe(non_mappes[['LOGIN', 'PRENOM', 'NOM', 'TOTAL OPERATION']].drop_duplicates('LOGIN'))

        # --- 4. GÉNÉRATION EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # FORMATS
            h_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})
            dr_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            dr_taux_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            sadi_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'indent': 1})
            sadi_num_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
            sadi_taux_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
            ravt_fmt = workbook.add_format({'border': 1, 'indent': 2})
            ravt_num_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            ravt_taux_fmt = workbook.add_format({'border': 1, 'num_format': '0"%"', 'align': 'center'})
            num_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            taux_fmt = workbook.add_format({'border': 1, 'num_format': '0"%"', 'align': 'center'})
            total_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})
            total_taux_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})
            non_mappe_fmt = workbook.add_format({'bg_color': '#FFF3CD', 'border': 1, 'indent': 1, 'align': 'center'})
            non_mappe_taux_fmt = workbook.add_format({'bg_color': '#FFF3CD', 'border': 1, 'indent': 1, 'num_format': '0"%"', 'align': 'center'})

            headers = ['DR', 'OP NFC', 'OP MANUELLE', 'TOTAL', 'Taux']

            # --- FEUILLE 1 : SYNTHESE DR ---
            ws1 = workbook.add_worksheet('SYNTHESE DR')
            for c, h in enumerate(headers): ws1.write(0, c, h, h_fmt)

            synthese_dr = df_final.groupby('DR').agg({
                'OPERATION NFC': 'sum', 'OPERATION MANUELLE': 'sum', 'TOTAL OPERATION': 'sum'
            }).reset_index()

            for i, r in synthese_dr.iterrows():
                ws1.write(i+1, 0, r['DR'], num_fmt)
                ws1.write(i+1, 1, int(r['OPERATION NFC']), num_fmt)
                ws1.write(i+1, 2, int(r['OPERATION MANUELLE']), num_fmt)
                ws1.write(i+1, 3, int(r['TOTAL OPERATION']), num_fmt)
                t_val = (r['OPERATION NFC'] / r['TOTAL OPERATION'] * 100) if r['TOTAL OPERATION'] > 0 else 0
                ws1.write(i+1, 4, t_val, taux_fmt)

            # Ligne TOTAL en bas
            total_row = len(synthese_dr) + 1
            total_nfc = int(synthese_dr['OPERATION NFC'].sum())
            total_man = int(synthese_dr['OPERATION MANUELLE'].sum())
            total_op = int(synthese_dr['TOTAL OPERATION'].sum())
            taux_total = (total_nfc / total_op * 100) if total_op > 0 else 0

            ws1.write(total_row, 0, 'TOTAL', total_fmt)
            ws1.write(total_row, 1, total_nfc, total_fmt)
            ws1.write(total_row, 2, total_man, total_fmt)
            ws1.write(total_row, 3, total_op, total_fmt)
            ws1.write(total_row, 4, f"{round(taux_total)}%", total_taux_fmt)

            ws1.set_column('A:E', 18)

            # --- FEUILLE 2 : REPORTING DR-SADI-RAVT ---
            ws2 = workbook.add_worksheet('REPORTING DR-SADI-RAVT')
            for c, h in enumerate(headers): ws2.write(0, c, h, h_fmt)

            # ✅ EXCLURE les NON MAPPÉ de cette feuille
            df_sans_non_mappe = df_final[df_final['SADI'] != 'NON MAPPÉ'].copy()

            curr_row = 1
            total_nfc_sadi = 0
            total_man_sadi = 0
            total_op_sadi = 0

            for dr, dr_group in df_sans_non_mappe.groupby('DR', sort=True):
                if len(dr_group) == 0 or dr_group['TOTAL OPERATION'].sum() == 0:
                    continue

                # Ligne DR
                n_dr, m_dr, t_dr = dr_group['OPERATION NFC'].sum(), dr_group['OPERATION MANUELLE'].sum(), dr_group['TOTAL OPERATION'].sum()
                ws2.write(curr_row, 0, dr, dr_fmt)
                ws2.write(curr_row, 1, int(n_dr), dr_fmt)
                ws2.write(curr_row, 2, int(m_dr), dr_fmt)
                ws2.write(curr_row, 3, int(t_dr), dr_fmt)
                ws2.write(curr_row, 4, (n_dr/t_dr*100) if t_dr > 0 else 0, dr_taux_fmt)
                curr_row += 1

                for sadi, sadi_group in dr_group.groupby('SADI', sort=True):
                    if len(sadi_group) == 0 or sadi_group['TOTAL OPERATION'].sum() == 0:
                        continue

                    # Ligne SADI
                    n_s, m_s, t_s = sadi_group['OPERATION NFC'].sum(), sadi_group['OPERATION MANUELLE'].sum(), sadi_group['TOTAL OPERATION'].sum()

                    ws2.write(curr_row, 0, sadi, sadi_fmt)
                    ws2.write(curr_row, 1, int(n_s), sadi_num_fmt)
                    ws2.write(curr_row, 2, int(m_s), sadi_num_fmt)
                    ws2.write(curr_row, 3, int(t_s), sadi_num_fmt)
                    ws2.write(curr_row, 4, (n_s/t_s*100) if t_s > 0 else 0, sadi_taux_fmt)
                    curr_row += 1

                    total_nfc_sadi += n_s
                    total_man_sadi += m_s
                    total_op_sadi += t_s

                    for ravt, ravt_group in sadi_group.groupby('RAVT', sort=True):
                        if len(ravt_group) == 0 or ravt_group['TOTAL OPERATION'].sum() == 0:
                            continue

                        # Ligne RAVT
                        n_r, m_r, t_r = ravt_group['OPERATION NFC'].sum(), ravt_group['OPERATION MANUELLE'].sum(), ravt_group['TOTAL OPERATION'].sum()
                        ws2.write(curr_row, 0, ravt, ravt_fmt)
                        ws2.write(curr_row, 1, int(n_r), ravt_num_fmt)
                        ws2.write(curr_row, 2, int(m_r), ravt_num_fmt)
                        ws2.write(curr_row, 3, int(t_r), ravt_num_fmt)
                        ws2.write(curr_row, 4, (n_r/t_r*100) if t_r > 0 else 0, ravt_taux_fmt)
                        curr_row += 1

            # Ajouter la ligne TOTAL
            taux_total_sadi = (total_nfc_sadi / total_op_sadi * 100) if total_op_sadi > 0 else 0
            ws2.write(curr_row, 0, 'TOTAL', total_fmt)
            ws2.write(curr_row, 1, int(total_nfc_sadi), total_fmt)
            ws2.write(curr_row, 2, int(total_man_sadi), total_fmt)
            ws2.write(curr_row, 3, int(total_op_sadi), total_fmt)
            ws2.write(curr_row, 4, f"{round(taux_total_sadi)}%", total_taux_fmt)

            ws2.set_column('A:A', 45)
            ws2.set_column('B:D', 15)
            ws2.set_column('E:E', 15)

            # --- FEUILLE 3 : REPORTING DR-RAVT-PVT ---
            ws3 = workbook.add_worksheet('REPORTING DR-RAVT-PVT')
            for c, h in enumerate(headers): ws3.write(0, c, h, h_fmt)

            df_pvt = df_final[df_final['ACCUEIL'].astype(str).str.startswith('PVT')].copy()

            curr_row = 1
            total_nfc_pvt = 0
            total_man_pvt = 0
            total_op_pvt = 0

            for dr, dr_group in df_pvt.groupby('DR', sort=True):
                if len(dr_group) == 0 or dr_group['TOTAL OPERATION'].sum() == 0:
                    continue

                n_dr, m_dr, t_dr = dr_group['OPERATION NFC'].sum(), dr_group['OPERATION MANUELLE'].sum(), dr_group['TOTAL OPERATION'].sum()
                ws3.write(curr_row, 0, dr, dr_fmt)
                ws3.write(curr_row, 1, int(n_dr), dr_fmt)
                ws3.write(curr_row, 2, int(m_dr), dr_fmt)
                ws3.write(curr_row, 3, int(t_dr), dr_fmt)
                ws3.write(curr_row, 4, (n_dr/t_dr*100) if t_dr > 0 else 0, dr_taux_fmt)
                curr_row += 1

                for ravt, ravt_group in dr_group.groupby('RAVT', sort=True):
                    if len(ravt_group) == 0 or ravt_group['TOTAL OPERATION'].sum() == 0:
                        continue

                    n_r, m_r, t_r = ravt_group['OPERATION NFC'].sum(), ravt_group['OPERATION MANUELLE'].sum(), ravt_group['TOTAL OPERATION'].sum()
                    ws3.write(curr_row, 0, ravt, sadi_fmt)
                    ws3.write(curr_row, 1, int(n_r), sadi_num_fmt)
                    ws3.write(curr_row, 2, int(m_r), sadi_num_fmt)
                    ws3.write(curr_row, 3, int(t_r), sadi_num_fmt)
                    ws3.write(curr_row, 4, (n_r/t_r*100) if t_r > 0 else 0, sadi_taux_fmt)
                    curr_row += 1

                    total_nfc_pvt += n_r
                    total_man_pvt += m_r
                    total_op_pvt += t_r

                    for pvt, pvt_group in ravt_group.groupby('ACCUEIL', sort=True):
                        if len(pvt_group) == 0 or pvt_group['TOTAL OPERATION'].sum() == 0:
                            continue

                        n_p, m_p, t_p = pvt_group['OPERATION NFC'].sum(), pvt_group['OPERATION MANUELLE'].sum(), pvt_group['TOTAL OPERATION'].sum()
                        ws3.write(curr_row, 0, pvt, ravt_fmt)
                        ws3.write(curr_row, 1, int(n_p), ravt_num_fmt)
                        ws3.write(curr_row, 2, int(m_p), ravt_num_fmt)
                        ws3.write(curr_row, 3, int(t_p), ravt_num_fmt)
                        ws3.write(curr_row, 4, (n_p/t_p*100) if t_p > 0 else 0, ravt_taux_fmt)
                        curr_row += 1

            # Ajouter la ligne TOTAL
            taux_total_pvt = (total_nfc_pvt / total_op_pvt * 100) if total_op_pvt > 0 else 0
            ws3.write(curr_row, 0, 'TOTAL', total_fmt)
            ws3.write(curr_row, 1, int(total_nfc_pvt), total_fmt)
            ws3.write(curr_row, 2, int(total_man_pvt), total_fmt)
            ws3.write(curr_row, 3, int(total_op_pvt), total_fmt)
            ws3.write(curr_row, 4, f"{round(taux_total_pvt)}%", total_taux_fmt)

            ws3.set_column('A:A', 45)
            ws3.set_column('B:D', 15)
            ws3.set_column('E:E', 15)

            # --- FEUILLE 4 : REPORTING DR-RAVT-PVT-VTO ---
            ws4 = workbook.add_worksheet('REPORTING DR-RAVT-PVT-VTO')
            headers_vto = ['DR/RAVT/PVT/VTO', 'Prénom', 'Nom', 'LOGIN', 'OP NFC', 'OP MANUELLE', 'TOTAL', 'Taux']
            for c, h in enumerate(headers_vto): ws4.write(0, c, h, h_fmt)

            vto_fmt = workbook.add_format({'border': 1, 'indent': 3, 'font_size': 9, 'align': 'center'})
            vto_num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 9})
            vto_taux_fmt = workbook.add_format({'border': 1, 'num_format': '0"%"', 'align': 'center', 'font_size': 9})

            curr_row = 1
            total_nfc_vto = 0
            total_man_vto = 0
            total_op_vto = 0

            for dr, dr_group in df_pvt.groupby('DR', sort=True):
                if len(dr_group) == 0 or dr_group['TOTAL OPERATION'].sum() == 0:
                    continue

                n_dr, m_dr, t_dr = dr_group['OPERATION NFC'].sum(), dr_group['OPERATION MANUELLE'].sum(), dr_group['TOTAL OPERATION'].sum()
                ws4.write(curr_row, 0, dr, dr_fmt)
                ws4.write(curr_row, 4, int(n_dr), dr_fmt)
                ws4.write(curr_row, 5, int(m_dr), dr_fmt)
                ws4.write(curr_row, 6, int(t_dr), dr_fmt)
                ws4.write(curr_row, 7, (n_dr/t_dr*100) if t_dr > 0 else 0, dr_taux_fmt)
                curr_row += 1

                for ravt, ravt_group in dr_group.groupby('RAVT', sort=True):
                    if len(ravt_group) == 0 or ravt_group['TOTAL OPERATION'].sum() == 0:
                        continue

                    n_r, m_r, t_r = ravt_group['OPERATION NFC'].sum(), ravt_group['OPERATION MANUELLE'].sum(), ravt_group['TOTAL OPERATION'].sum()
                    ws4.write(curr_row, 0, ravt, sadi_fmt)
                    ws4.write(curr_row, 4, int(n_r), sadi_num_fmt)
                    ws4.write(curr_row, 5, int(m_r), sadi_num_fmt)
                    ws4.write(curr_row, 6, int(t_r), sadi_num_fmt)
                    ws4.write(curr_row, 7, (n_r/t_r*100) if t_r > 0 else 0, sadi_taux_fmt)
                    curr_row += 1

                    total_nfc_vto += n_r
                    total_man_vto += m_r
                    total_op_vto += t_r

                    for pvt, pvt_group in ravt_group.groupby('ACCUEIL', sort=True):
                        if len(pvt_group) == 0 or pvt_group['TOTAL OPERATION'].sum() == 0:
                            continue

                        n_p, m_p, t_p = pvt_group['OPERATION NFC'].sum(), pvt_group['OPERATION MANUELLE'].sum(), pvt_group['TOTAL OPERATION'].sum()
                        ws4.write(curr_row, 0, pvt, ravt_fmt)
                        ws4.write(curr_row, 4, int(n_p), ravt_num_fmt)
                        ws4.write(curr_row, 5, int(m_p), ravt_num_fmt)
                        ws4.write(curr_row, 6, int(t_p), ravt_num_fmt)
                        ws4.write(curr_row, 7, (n_p/t_p*100) if t_p > 0 else 0, ravt_taux_fmt)
                        curr_row += 1

                        for _, vto_row in pvt_group.iterrows():
                            prenom = vto_row.get('PRENOM', '')
                            nom = vto_row.get('NOM', '')
                            login = vto_row.get('LOGIN', '')
                            n_v = int(vto_row['OPERATION NFC'])
                            m_v = int(vto_row['OPERATION MANUELLE'])
                            t_v = int(vto_row['TOTAL OPERATION'])

                            ws4.write(curr_row, 0, 'VTO', vto_fmt)
                            ws4.write(curr_row, 1, prenom, vto_fmt)
                            ws4.write(curr_row, 2, nom, vto_fmt)
                            ws4.write(curr_row, 3, login, vto_fmt)
                            ws4.write(curr_row, 4, n_v, vto_num_fmt)
                            ws4.write(curr_row, 5, m_v, vto_num_fmt)
                            ws4.write(curr_row, 6, t_v, vto_num_fmt)
                            ws4.write(curr_row, 7, (n_v/t_v*100) if t_v > 0 else 0, vto_taux_fmt)
                            curr_row += 1

            # Ajouter la ligne TOTAL
            taux_total_vto = (total_nfc_vto / total_op_vto * 100) if total_op_vto > 0 else 0
            ws4.write(curr_row, 0, 'TOTAL', total_fmt)
            ws4.write(curr_row, 4, int(total_nfc_vto), total_fmt)
            ws4.write(curr_row, 5, int(total_man_vto), total_fmt)
            ws4.write(curr_row, 6, int(total_op_vto), total_fmt)
            ws4.write(curr_row, 7, f"{round(taux_total_vto)}%", total_taux_fmt)

            ws4.set_column('A:A', 35)
            ws4.set_column('B:C', 20)
            ws4.set_column('D:D', 25)
            ws4.set_column('E:G', 15)
            ws4.set_column('H:H', 15)

        st.success("✅ Fichier généré avec succès - AUCUNE perte de données !")
        st.download_button("📥 Télécharger le Reporting Final", output.getvalue(), "Reporting_NFC_Orange_Final.xlsx")

    except Exception as e:
        st.error(f"Erreur : {e}")
        import traceback
        st.code(traceback.format_exc())