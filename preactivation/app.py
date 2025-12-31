import streamlit as st
import pandas as pd
import io
import re
from filtre_date import *
from repporting_NFC import *


st.set_page_config(page_title="Orange Preactivation Specialist", layout="wide")

st.title("🚀 Générateur de Reporting Préactivations")
st.write("Tri sélectif : Clôtures avec Statut / Rejets avec colonne PREACTIVATION")

uploaded_file = st.file_uploader("Déposez le fichier de ventes global (XLSB ou XLSX)", type=["xlsb", "xlsx"])

if uploaded_file:
    try:
        st.write("⏳ Lecture des données détaillées...")
        engine = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else None

        # Tentative de lecture de la feuille de détail (index 1)
        try:
            df = pd.read_excel(uploaded_file, engine=engine, sheet_name=1)
        except:
            df = pd.read_excel(uploaded_file, engine=engine, sheet_name=0)

        df.columns = [str(c).strip() for c in df.columns]

        # 1. FILTRE : Uniquement les "PREACTIVATION"
        col_filtre = 'preactivateur' if 'preactivateur' in df.columns else 'COMMENTAIRE'
        if col_filtre in df.columns:
            df = df[df[col_filtre].astype(str).str.contains('PREACTIVATION', case=False, na=False)]

        # Conversion intensité
        df['intensite'] = pd.to_numeric(df['intensite'], errors='coerce').fillna(0)

        # 2. SÉPARATION DES DONNÉES
        df_clotures_raw = df[df['intensite'] >= 80].copy()
        df_rejets_raw = df[df['intensite'] < 80].copy()

        # 3. FONCTION POUR EXTRAIRE RAVT ET ACCUEIL - SIMPLIFIÉE
        def extraire_ravt_accueil_simple(colonne_accueil):
            if pd.isna(colonne_accueil):
                return pd.Series(['', ''])

            texte = str(colonne_accueil).strip()

            # Nettoyer les espaces multiples
            texte = re.sub(r'\s+', ' ', texte)

            # 1. Chercher TOUT ce qui est entre parenthèses - c'est le RAVT
            pattern_parentheses = re.compile(r'\(([^)]+)\)')
            match_parentheses = pattern_parentheses.search(texte)

            ravt = ''
            accueil = texte

            if match_parentheses:
                # Ce qui est entre parenthèses est le RAVT
                ravt = match_parentheses.group(1).strip()

                # Retirer les parenthèses et leur contenu pour obtenir l'accueil
                accueil = pattern_parentheses.sub('', texte).strip()

                # Nettoyer l'accueil (supprimer les espaces en trop, parenthèses vides)
                accueil = re.sub(r'\s+', ' ', accueil)
                accueil = accueil.strip()
                accueil = accueil.strip('()')

                # S'assurer que l'accueil ne contient pas de parenthèses vides
                if accueil.endswith('(') or accueil.startswith(')'):
                    accueil = accueil.strip('()')

            # Si pas de parenthèses, tout le texte est l'accueil
            else:
                accueil = texte
                ravt = ''  # Pas de RAVT si pas de parenthèses

            # Nettoyage final
            accueil = accueil.strip()
            ravt = ravt.strip()

            return pd.Series([ravt, accueil])

        # 4. FONCTION POUR FILTRER PAR TYPE D'ACCUEIL (BOUTIQUE ou PVT)
        def filtrer_par_type_accueil(df_source):
            if df_source.empty:
                return df_source

            df_temp = df_source.copy()

            if 'ACCUEIL_VENDEUR' in df_temp.columns:
                # D'abord extraire l'accueil
                df_temp['ACCUEIL_EXTRACT'] = df_temp['ACCUEIL_VENDEUR'].apply(
                    lambda x: extraire_ravt_accueil_simple(x)[1] if pd.notna(x) else ''
                )

                # Filtrer les lignes où ACCUEIL commence par BOUTIQUE ou PVT
                masque = df_temp['ACCUEIL_EXTRACT'].str.upper().str.startswith(('BOUTIQUE', 'PVT'))
                df_filtre = df_temp[masque].copy()

                # Supprimer la colonne temporaire
                df_filtre = df_filtre.drop(columns=['ACCUEIL_EXTRACT'])

                return df_filtre
            else:
                return df_source

        # Appliquer le filtre BOUTIQUE/PVT
        df_clotures_raw = filtrer_par_type_accueil(df_clotures_raw)
        df_rejets_raw = filtrer_par_type_accueil(df_rejets_raw)

        # 5. FONCTION POUR PRÉPARER LES DONNÉES AVEC REGROUPEMENT ET BON CALCUL
        def preparer_donnees_avec_regroupement(df_source, type_donnees='clotures'):
            if df_source.empty:
                return pd.DataFrame()

            # Créer une copie pour ne pas modifier l'original
            df_temp = df_source.copy()

            # Extraire RAVT et ACCUEIL avec la fonction simple
            if 'ACCUEIL_VENDEUR' in df_temp.columns:
                df_temp[['RAVT', 'ACCUEIL']] = df_temp['ACCUEIL_VENDEUR'].apply(extraire_ravt_accueil_simple)

                # Vérifier les RAVT vides
                ravts_vides = df_temp['RAVT'].isna() | (df_temp['RAVT'] == '')
                if ravts_vides.any():
                    st.warning(f"⚠️ Attention : {ravts_vides.sum()} lignes n'ont pas de RAVT (pas de parenthèses)")
            else:
                df_temp['RAVT'] = ''
                df_temp['ACCUEIL'] = df_temp['AGENCE_VENDEUR'] if 'AGENCE_VENDEUR' in df_temp.columns else ''

            # Gestion de la colonne DR avec renommage
            if 'DR' in df_temp.columns:
                dr_column = df_temp['DR']
            elif 'AGENCE_VENDEUR' in df_temp.columns:
                dr_column = df_temp['AGENCE_VENDEUR']
            else:
                dr_column = pd.Series([''] * len(df_temp))

            # Renommage des DR selon les spécifications
            dr_mapping = {
                'DV-DRVE_DIRECTION REGIONALE DES VENTES EST': 'DRE',
                'DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE': 'DRC',
                'DV-DRVN_DIRECTION REGIONALE DES VENTES NORD': 'DRN',
                'DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST': 'DRSE',
                'DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2': 'DR2',
                'DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1': 'DR1'
            }

            # Appliquer le renommage
            df_temp['DR'] = dr_column.replace(dr_mapping)

            # S'assurer que les colonnes nécessaires existent
            colonnes_requises = ['LOGIN_VENDEUR', 'DR', 'RAVT', 'ACCUEIL', 'PRENOM_VENDEUR',
                               'NOM_VENDEUR', 'intensite']
            for col in colonnes_requises:
                if col not in df_temp.columns:
                    df_temp[col] = ''

            # IMPORTANT : Chaque ligne = 1 préactivation
            # Ajouter une colonne de comptage
            df_temp['COMPTE_PREACTIVATION'] = 1

            # Regrouper par LOGIN pour éviter les répétitions
            df_grouped = df_temp.groupby('LOGIN_VENDEUR').agg({
                'DR': 'first',
                'RAVT': 'first',
                'ACCUEIL': 'first',
                'PRENOM_VENDEUR': 'first',
                'NOM_VENDEUR': 'first',
                'COMPTE_PREACTIVATION': 'sum',  # Nombre total de préactivations = nombre de lignes
                'intensite': 'mean'  # Moyenne de l'intensité
            }).reset_index()

            # Renommer les colonnes
            df_grouped = df_grouped.rename(columns={
                'LOGIN_VENDEUR': 'LOGIN',
                'COMPTE_PREACTIVATION': 'PREACTIVATIONS',  # Maintenant c'est le vrai compte
                'intensite': 'CRITERE_INTENSITE'
            })

            # Réorganiser les colonnes pour mettre DR en premier
            colonnes_finales = ['DR', 'RAVT', 'ACCUEIL', 'PRENOM_VENDEUR', 'NOM_VENDEUR',
                              'LOGIN', 'PREACTIVATIONS', 'CRITERE_INTENSITE']

            # S'assurer que toutes les colonnes existent
            for col in colonnes_finales:
                if col not in df_grouped.columns:
                    df_grouped[col] = ''

            df_final = df_grouped[colonnes_finales]

            # Ajouter les colonnes spécifiques selon le type
            if type_donnees == 'clotures':
                df_final['STATUT'] = 'clôturé'
            elif type_donnees == 'rejets':
                df_final['PREACTIVATION'] = 'PREACTIVATION'

            # Trier par CRITERE_INTENSITE par ordre décroissant
            df_final = df_final.sort_values('CRITERE_INTENSITE', ascending=False)

            return df_final

        # --- PRÉPARATION FEUILLE 1 : CLÔTURES ---
        df_clotures_final = preparer_donnees_avec_regroupement(df_clotures_raw, 'clotures')

        # --- PRÉPARATION FEUILLE 2 : REJETS ---
        df_rejets_final = preparer_donnees_avec_regroupement(df_rejets_raw, 'rejets')

        # 7. INTERFACE PRINCIPALE
        st.success(f"✅ Analyse terminée : {len(df_clotures_final)} Logins Clôturés / {len(df_rejets_final)} Logins Rejetés")

        # Aperçu des données - Top 10
        st.subheader("👁️ Top 10 par préactivations")

        if not df_clotures_final.empty:
            # Trier par nombre de préactivations
            df_clotures_trie = df_clotures_final.sort_values('PREACTIVATIONS', ascending=False).head(10)
            st.dataframe(df_clotures_trie[['LOGIN', 'ACCUEIL', 'PREACTIVATIONS', 'DR']], use_container_width=True)

        # 8. GÉNÉRATION DU FICHIER EXCEL
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Onglet 1 - LOGIN CLOTURES
            if not df_clotures_final.empty:
                colonnes_export_clotures = ['DR', 'RAVT', 'ACCUEIL', 'PRENOM_VENDEUR', 'NOM_VENDEUR',
                                          'LOGIN', 'PREACTIVATIONS', 'CRITERE_INTENSITE', 'STATUT']
                df_clotures_export = df_clotures_final[colonnes_export_clotures]
                df_clotures_export.to_excel(writer, sheet_name='LOGIN CLOTURES', index=False)

                # Formatage de la feuille LOGIN CLOTURES
                workbook = writer.book
                worksheet = writer.sheets['LOGIN CLOTURES']

                # Format pour l'en-tête (fond bleu, texte blanc, gras)
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#4472C4',
                    'font_color': 'white',
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })

                # Format pour les cellules "clôturé" (texte rouge)
                statut_format = workbook.add_format({
                    'font_color': 'red',
                    'align': 'center',
                    'border': 1
                })

                # Format pour les cellules normales
                cell_format = workbook.add_format({
                    'border': 1,
                    'align': 'left',
                    'valign': 'vcenter'
                })

                # Format pour les nombres
                number_format = workbook.add_format({
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                # Appliquer le format d'en-tête
                for col_num, value in enumerate(df_clotures_export.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                # Appliquer le format aux données
                for row_num in range(len(df_clotures_export)):
                    for col_num, col_name in enumerate(df_clotures_export.columns):
                        value = df_clotures_export.iloc[row_num, col_num]

                        # Format spécial pour la colonne STATUT (rouge)
                        if col_name == 'STATUT':
                            worksheet.write(row_num + 1, col_num, value, statut_format)
                        # Format pour les colonnes numériques
                        elif col_name in ['PREACTIVATIONS', 'CRITERE_INTENSITE']:
                            worksheet.write(row_num + 1, col_num, value, number_format)
                        else:
                            worksheet.write(row_num + 1, col_num, value, cell_format)

                # Ajuster la largeur des colonnes
                worksheet.set_column('A:A', 10)  # DR
                worksheet.set_column('B:B', 15)  # RAVT
                worksheet.set_column('C:C', 30)  # ACCUEIL
                worksheet.set_column('D:E', 20)  # PRENOM, NOM
                worksheet.set_column('F:F', 20)  # LOGIN
                worksheet.set_column('G:G', 15)  # PREACTIVATIONS
                worksheet.set_column('H:H', 18)  # CRITERE_INTENSITE
                worksheet.set_column('I:I', 12)  # STATUT

            # Onglet 2 - PREACTIVATIONS
            if not df_rejets_final.empty:
                colonnes_export_rejets = ['DR', 'RAVT', 'ACCUEIL', 'PRENOM_VENDEUR', 'NOM_VENDEUR',
                                        'LOGIN', 'PREACTIVATIONS', 'CRITERE_INTENSITE', 'PREACTIVATION']
                df_rejets_export = df_rejets_final[colonnes_export_rejets]
                df_rejets_export.to_excel(writer, sheet_name='PREACTIVATIONS', index=False)

                # Formatage de la feuille PREACTIVATIONS
                worksheet2 = writer.sheets['PREACTIVATIONS']

                # Format pour la colonne PREACTIVATION (texte orange)
                preactivation_format = workbook.add_format({
                    'font_color': '#FF6600',
                    'align': 'center',
                    'border': 1
                })

                # Appliquer le format d'en-tête
                for col_num, value in enumerate(df_rejets_export.columns.values):
                    worksheet2.write(0, col_num, value, header_format)

                # Appliquer le format aux données
                for row_num in range(len(df_rejets_export)):
                    for col_num, col_name in enumerate(df_rejets_export.columns):
                        value = df_rejets_export.iloc[row_num, col_num]

                        # Format spécial pour la colonne PREACTIVATION (orange)
                        if col_name == 'PREACTIVATION':
                            worksheet2.write(row_num + 1, col_num, value, preactivation_format)
                        # Format pour les colonnes numériques
                        elif col_name in ['PREACTIVATIONS', 'CRITERE_INTENSITE']:
                            worksheet2.write(row_num + 1, col_num, value, number_format)
                        else:
                            worksheet2.write(row_num + 1, col_num, value, cell_format)

                # Ajuster la largeur des colonnes
                worksheet2.set_column('A:A', 10)  # DR
                worksheet2.set_column('B:B', 15)  # RAVT
                worksheet2.set_column('C:C', 30)  # ACCUEIL
                worksheet2.set_column('D:E', 20)  # PRENOM, NOM
                worksheet2.set_column('F:F', 20)  # LOGIN
                worksheet2.set_column('G:G', 15)  # PREACTIVATIONS
                worksheet2.set_column('H:H', 18)  # CRITERE_INTENSITE
                worksheet2.set_column('I:I', 18)  # PREACTIVATION

        st.download_button(
            label="📥 Télécharger le Fichier Propre",
            data=output.getvalue(),
            file_name="Reporting_Final_Preactivations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erreur : {e}")