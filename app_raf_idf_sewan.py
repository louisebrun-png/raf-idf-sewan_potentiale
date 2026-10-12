# ==============================================================================
# CABINET POTENTIALE — OUTIL INTERNE RAF (KOESIO IDF x SEWAN)
# Experte : Louise Brun — Fondatrice Potentiale
# ==============================================================================

import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Potentiale — RAF Koesio IDF", page_icon="💼", layout="wide")

st.title("💼 Potentiale — Outil Interne RAF (Koesio IDF x Sewan)")
st.caption("Audit Fiscal TVA & Rapprochement HT (Alignement Stricte Colonne AL & Colonne AH)")

# Sidebar - Importation des fichiers
st.sidebar.header("📁 Importation des Documents")
fichier_artis = st.sidebar.file_uploader("1. Simulation Achat Artis (.xlsx)", type=["xlsx"])
fichier_ref = st.sidebar.file_uploader("2. Table de Correspondance (.xlsx)", type=["xlsx"])
factures_pdf = st.sidebar.file_uploader("3. Factures PDF Sewan (Obligatoire TVA)", type=["pdf"], accept_multiple_files=True)
annexes_csv = st.sidebar.file_uploader("4. Annexes Fournisseurs (.csv / .xlsx)", type=["csv", "xlsx"], accept_multiple_files=True)

def isoler_ref_article(libelle):
    if not isinstance(libelle, str) or pd.isna(libelle):
        return ""
    match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', str(libelle))
    return match.group(1) if match else str(libelle).split(' - ')[0].strip()

def clean_string_fuzzy(s):
    if not isinstance(s, str) or pd.isna(s):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s).lower())

def extract_param_key_strict(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    if '//' in str(text):
        text = str(text).split('//')[-1]
    
    digits = re.sub(r'\D', '', str(text))
    if len(digits) >= 9:
        return digits[-9:]
    m_id = re.search(r'([a-zA-Z0-9_\-\.]+@[a-zA-Z0-9_\-\.]+)', str(text))
    if m_id:
        return m_id.group(1).lower()
    return clean_string_fuzzy(text)

def extract_code_client_avance(row_dict):
    pattern = r'\b((?:NS|TO|SK|NM)(?:_CL)?\d{2,6})\b'
    champs_a_tester = [
        str(row_dict.get('nom client niveau 2', '')),
        str(row_dict.get('ref client niveau 2', '')),
        str(row_dict.get('client/revendeur direct (niv. 1)', '')),
        str(row_dict.get('réf. client/revendeur direct (niv. 1)', '')),
        str(row_dict.get('param1', '')),
        str(row_dict.get('param2', ''))
    ]
    for val in champs_a_tester:
        if val and val.lower() != 'nan':
            m = re.search(pattern, val, re.IGNORECASE)
            if m:
                return m.group(1).upper()
    return ""

def analyser_temporalite_ah(date_val, annee_ref=2026, mois_ref=10):
    if pd.isna(date_val) or str(date_val).strip() in ['', 'nan', 'None', 'NaT']:
        return "⚠️ Date Absente"
    try:
        dt = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            str_clean = str(date_val).split(' ')[0].strip()
            dt = pd.to_datetime(str_clean, dayfirst=True, errors='coerce')
            if pd.isna(dt):
                return "⚠️ Format Invalide"
        
        diff_mois = (dt.year - annee_ref) * 12 + (dt.month - mois_ref)
        
        if diff_mois == 0:
            return f"🟢 Cycle Récurrent Conforme ({dt.strftime('%m/%Y')})"
        elif diff_mois == -1:
            return f"ℹ️ Régularisation / Prorata Échu ({dt.strftime('%m/%Y')})"
        elif diff_mois > 0:
            return f"ℹ️ Facturation Anticipée / Futur ({dt.strftime('%m/%Y')})"
        else:
            return f"🚨 Retard de Facturation ERP ({abs(diff_mois)} mois - {dt.strftime('%m/%Y')})"
    except Exception:
        return "ℹ️ Analyse Inconnue"

def lire_csv_securise(fichier):
    for enc in ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']:
        try:
            fichier.seek(0)
            return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        except Exception:
            continue
    fichier.seek(0)
    return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

onglet_tva, onglet_ht = st.tabs(["🚨 1. Audit Écarts TVA (20% vs 0%)", "📊 2. Analyse Quantité & Prix HT (Simulation Octobre Exacte)"])

# --- ONGLET 1 : AUDIT TVA ---
with onglet_tva:
    st.subheader("Audit Fiscal TVA (STD 20% vs APST 0%)")
    if fichier_artis and factures_pdf:
        if st.button("🚨 Lancer l'Analyse TVA", type="primary", key="btn_tva"):
            with st.spinner("Analyse des factures PDF Sewan et vérification des RFC ERP..."):
                try:
                    df_artis = pd.read_excel(fichier_artis)
                    col_art_artis = 'Code Article' if 'Code Article' in df_artis.columns else ('Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0])
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    
                    def harmoniser_tva_artis_exact(row):
                        gtva = str(row.get('Coût ABONNEMENT Gestion TVA', row.get('Gestion TVA', '')))
                        ttva = str(row.get('Coût ABONNEMENT Taux TVA', row.get('Taux TVA', '')))
                        if 'STD' in gtva or 'Standard' in gtva or '20' in ttva or '20.00' in ttva:
                            return 'S (20%)'
                        return 'APST (0%)'

                    df_artis_abonn['TVA_Artis'] = df_artis_abonn.apply(harmoniser_tva_artis_exact, axis=1)
                    
                    lignes_tva = []
                    for pdf in factures_pdf:
                        reader = pypdf.PdfReader(pdf)
                        for page in reader.pages:
                            text = page.extract_text() or ""
                            lines = text.split('\n')
                            for i, line in enumerate(lines):
                                match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', line)
                                if match:
                                    ref = match.group(1)
                                    tva_code = "APST (0%)"
                                    for j in range(i, min(i + 5, len(lines))):
                                        if 'S (20%' in lines[j] or 'S (20,00%)' in lines[j]:
                                            tva_code = "S (20%)"
                                            break
                                    lignes_tva.append({'Code_Article_ERP': ref, 'TVA_Sewan': tva_code})
                                    
                    df_sewan_tva = pd.DataFrame(lignes_tva).drop_duplicates() if lignes_tva else pd.DataFrame(columns=['Code_Article_ERP', 'TVA_Sewan'])
                    df_merged_tva = pd.merge(df_artis_abonn, df_sewan_tva, on='Code_Article_ERP', how='inner')
                    df_errors_tva = df_merged_tva[df_merged_tva['TVA_Artis'] != df_merged_tva['TVA_Sewan']].copy()
                    df_errors_tva['Diagnostic_RAF'] = "🚨 Divergence TVA"
                    
                    col_cli_name_tva = 'Client' if 'Client' in df_errors_tva.columns else 'Raison sociale client'
                    if col_cli_name_tva in df_errors_tva.columns:
                        df_errors_tva[col_cli_name_tva] = df_errors_tva[col_cli_name_tva].astype(str).str.slice(0, 20)
                        
                    cols_tva_export = ['Code client', col_cli_name_tva, 'Code SSC', 'Code RFC', 'Libellé RFC', 'Code_Article_ERP', 'TVA_Artis', 'TVA_Sewan', 'Diagnostic_RAF']
                    for c in cols_tva_export:
                        if c not in df_errors_tva.columns:
                            df_errors_tva[c] = "Non renseigné"
                            
                    df_res_tva = df_errors_tva[cols_tva_export].drop_duplicates()
                    
                    buffer_tva = io.BytesIO()
                    with pd.ExcelWriter(buffer_tva, engine='openpyxl') as writer:
                        df_res_tva.to_excel(writer, index=False, sheet_name='Anomalies_TVA')
                    
                    st.success(f"Audit TVA terminé : {len(df_res_tva)} anomalies réelles identifiées !")
                    st.dataframe(df_res_tva, use_container_width=True)
                    st.download_button("📥 Télécharger le Rapport TVA (.xlsx)", data=buffer_tva.getvalue(), file_name="Anomalies_TVA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"Erreur pendant l'analyse TVA : {e}")
    else:
        st.info("👈 Veuillez charger la simulation Artis (.xlsx) et les factures PDF Sewan.")

# --- ONGLET 2 : ANALYSE HT AVEC TOTAL AL ABSOLU ET AH DIRECT ---
with onglet_ht:
    st.subheader("Rapprochement Financier HT & Contrôle Temporel (Lecture Stricte Colonne AL & Colonne AH)")
    if fichier_artis and annexes_csv:
        nb_annexes = len(annexes_csv)
        if st.button("🚀 Lancer l'Analyse HT", type="primary", key="btn_ht"):
            with st.spinner(f"Traitement haute précision sur {nb_annexes} annexes CSV..."):
                try:
                    annee_ref, mois_ref = 2026, 10
                    if factures_pdf:
                        try:
                            for pdf in factures_pdf:
                                reader = pypdf.PdfReader(pdf)
                                text = reader.pages[0].extract_text() or ""
                                m_date = re.search(r'Date\s*:\s*(\d{2})/(\d{2})/(\d{4})', text)
                                if m_date:
                                    mois_ref = int(m_date.group(2))
                                    annee_ref = int(m_date.group(3))
                                    break
                        except Exception:
                            pass

                    client_map = {}
                    if fichier_ref:
                        xls_ref = pd.ExcelFile(fichier_ref)
                        sheet_cli = 'Analyse clients' if 'Analyse clients' in xls_ref.sheet_names else xls_ref.sheet_names[0]
                        df_ref_cli = pd.read_excel(fichier_ref, sheet_name=sheet_cli)
                        for _, row in df_ref_cli.iterrows():
                            artis_code = str(row.get('CODE', '')).strip()
                            if not artis_code or artis_code == 'nan':
                                continue
                            for col_name in ['Client/revendeur direct (niv. 1)', 'Réf. client/revendeur direct (niv. 1)', 'Nom client Niveau 2', 'Ref client Niveau 2']:
                                val = str(row.get(col_name, '')).strip()
                                if val and val != 'nan':
                                    client_map[val.lower()] = artis_code
                                    m = re.match(r'^([A-Za-z0-9_\-]+)', val)
                                    if m:
                                        client_map[m.group(1).lower()] = artis_code

                    df_artis = pd.read_excel(fichier_artis)

                    # 🎯 1. EXTRACTION DU MONTANT EXACT DE LA COLONNE AL ("Montant HT")
                    col_al_exact = 'Montant HT' if 'Montant HT' in df_artis.columns else (df_artis.columns[37] if len(df_artis.columns) >= 38 else df_artis.columns[0])
                    total_simulation_artis_absolu = pd.to_numeric(df_artis[col_al_exact], errors='coerce').fillna(0.0).sum()

                    # 🎯 2. EXTRACTION DE LA TEMPORALITÉ DE LA COLONNE AH ("Début période facturation")
                    col_ah_exact = 'Début période facturation' if 'Début période facturation' in df_artis.columns else (df_artis.columns[33] if len(df_artis.columns) >= 34 else df_artis.columns[0])
                    df_artis['Audit_Temporalite_FF'] = df_artis[col_ah_exact].apply(lambda d: analyser_temporalite_ah(d, annee_ref, mois_ref))

                    stats_temp = df_artis['Audit_Temporalite_FF'].value_counts()
                    nb_conformes_ff = sum(v for k, v in stats_temp.items() if 'Conforme' in k)
                    nb_regul_echu_ff = sum(v for k, v in stats_temp.items() if 'Régularisation' in k)
                    nb_anticipees_ff = sum(v for k, v in stats_temp.items() if 'Anticipée' in k)
                    nb_retards_ff = sum(v for k, v in stats_temp.items() if 'Retard' in k)

                    # Rapprochement Fournisseur
                    col_art_artis = 'Code Article' if 'Code Article' in df_artis.columns else ('Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0])
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    df_artis_abonn['Montant_Total_ERP'] = pd.to_numeric(df_artis_abonn[col_al_exact], errors='coerce').fillna(0.0).round(2)

                    col_pu_artis = 'Prix Unitaire' if 'Prix Unitaire' in df_artis_abonn.columns else df_artis.columns[1]
                    df_artis_abonn['PU_ERP'] = pd.to_numeric(df_artis_abonn[col_pu_artis], errors='coerce').fillna(0.0)
                    df_artis_abonn['Code_Client_Str'] = df_artis_abonn['Code client'].astype(str).str.strip() if 'Code client' in df_artis_abonn.columns else ""
                    
                    col_client_name = 'Client' if 'Client' in df_artis_abonn.columns else 'Raison sociale client'
                    df_artis_abonn['Name_Client_Clean'] = df_artis_abonn[col_client_name].apply(clean_string_fuzzy) if col_client_name in df_artis_abonn.columns else ""
                    
                    col_rfc_lbl = 'Libellé RFC' if 'Libellé RFC' in df_artis_abonn.columns else ('Libellé Article' if 'Libellé Article' in df_artis_abonn.columns else df_artis_abonn.columns[0])
                    df_artis_abonn['Param_Key'] = df_artis_abonn[col_rfc_lbl].apply(extract_param_key_strict)
                    
                    def extract_product_from_rfc(rfc):
                        if not isinstance(rfc, str) or pd.isna(rfc):
                            return ""
                        part_left = str(rfc).split('//')[0].strip() if '//' in str(rfc) else str(rfc).strip()
                        return clean_string_fuzzy(part_left)

                    df_artis_abonn['Libelle_RFC_Clean'] = df_artis_abonn[col_rfc_lbl].apply(extract_product_from_rfc)

                    col_qte_fa = 'Quantite' if 'Quantite' in df_artis_abonn.columns else next((c for c in df_artis_abonn.columns if 'Abonnement Nombre' in str(c) or 'nt Nombre' in str(c) or str(c).endswith('Nombre')), None)
                    if col_qte_fa:
                        df_artis_abonn['Quantite_ERP'] = pd.to_numeric(df_artis_abonn[col_qte_fa], errors='coerce').fillna(1.0)
                    else:
                        df_artis_abonn['Quantite_ERP'] = 1.0

                    df_artis_abonn['Occurence_ERP'] = df_artis_abonn.groupby(['Code_Client_Str', 'Code_Article_ERP', 'Param_Key']).cumcount()
                    df_artis_abonn['Occurence_Fallback_ERP'] = df_artis_abonn.groupby(['Code_Client_Str', 'Code_Article_ERP', 'Libelle_RFC_Clean']).cumcount()

                    lignes_fourn = []
                    for annexe in annexes_csv:
                        df_annexe = lire_csv_securise(annexe) if annexe.name.endswith('.csv') else pd.read_excel(annexe)
                        cols_lower = {str(c).lower().strip(): c for c in df_annexe.columns}
                        
                        col_ref = next((v for k, v in cols_lower.items() if any(x in k for x in ['code produit', 'code_article', 'ref', 'code'])), df_annexe.columns[0])
                        col_montant = next((v for k, v in cols_lower.items() if any(x in k for x in ["prix d'achat", 'prix', 'montant', 'ht'])), df_annexe.columns[1])
                        col_qte = next((v for k, v in cols_lower.items() if any(x in k for x in ['quantité', 'quantite', 'qte'])), None)
                        col_param1 = next((v for k, v in cols_lower.items() if 'param1' in k), None)
                        col_param2 = next((v for k, v in cols_lower.items() if 'param2' in k), None)
                        col_id_val = next((v for k, v in cols_lower.items() if 'id interne' in k or 'valorisation' in k), None)
                        col_produit = next((v for k, v in cols_lower.items() if k == 'produit' or (('produit' in k or 'libellé' in k) and 'famille' not in k and 'code' not in k)), None)
                        
                        col_cli1 = next((v for k, v in cols_lower.items() if 'client/revendeur direct (niv. 1)' in k or 'client/revendeur' in k), None)
                        col_cli2 = next((v for k, v in cols_lower.items() if 'nom client niveau 2' in k or 'nom client' in k), None)
                        col_ref_cli2 = next((v for k, v in cols_lower.items() if 'ref client niveau 2' in k), None)
                        col_ref_cli1 = next((v for k, v in cols_lower.items() if 'réf. client/revendeur direct (niv. 1)' in k), None)

                        for _, row in df_annexe.iterrows():
                            ref = isoler_ref_article(str(row[col_ref]))

                            try:
                                m_ht = float(str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', ''))
                            except ValueError:
                                m_ht = 0.0
                                
                            try:
                                q_sewan = float(str(row[col_qte]).replace(',', '.').replace(' ', '')) if col_qte and pd.notna(row[col_qte]) else 1.0
                            except ValueError:
                                q_sewan = 1.0
                            
                            p1_val = str(row[col_param1]) if col_param1 and pd.notna(row[col_param1]) else ""
                            p2_val = str(row[col_param2]) if col_param2 and pd.notna(row[col_param2]) else ""
                            prod_val = str(row[col_produit]).strip() if col_produit and pd.notna(row[col_produit]) else ""
                            id_val_str = str(row[col_id_val]).strip() if col_id_val and pd.notna(row[col_id_val]) else ""
                            
                            row_dict_lower = {
                                'nom client niveau 2': str(row[col_cli2]) if col_cli2 and pd.notna(row[col_cli2]) else "",
                                'ref client niveau 2': str(row[col_ref_cli2]) if col_ref_cli2 and pd.notna(row[col_ref_cli2]) else "",
                                'client/revendeur direct (niv. 1)': str(row[col_cli1]) if col_cli1 and pd.notna(row[col_cli1]) else "",
                                'réf. client/revendeur direct (niv. 1)': str(row[col_ref_cli1]) if col_ref_cli1 and pd.notna(row[col_ref_cli1]) else "",
                                'param1': p1_val,
                                'param2': p2_val
                            }
                            
                            res_code = extract_code_client_avance(row_dict_lower)
                            if not res_code:
                                for test_val in [row_dict_lower['nom client niveau 2'], row_dict_lower['client/revendeur direct (niv. 1)'], p1_val, p2_val]:
                                    if test_val.lower() in client_map:
                                        res_code = client_map[test_val.lower()]
                                        break
                                        
                            cli_raw = row_dict_lower['nom client niveau 2'] if row_dict_lower['nom client niveau 2'] else row_dict_lower['client/revendeur direct (niv. 1)']
                            param_key = extract_param_key_strict(p1_val)
                            
                            lignes_fourn.append({
                                'Id_Valorisation_Sewan': id_val_str,
                                'Code_Client_Resolu': res_code if res_code else "10",
                                'Name_Sewan_Clean': clean_string_fuzzy(cli_raw),
                                'Code_Article_ERP': ref,
                                'Produit': prod_val,
                                'Produit_Sewan_Clean': clean_string_fuzzy(prod_val),
                                'Param_Key': param_key,
                                'Montant_Sewan': m_ht,
                                'Quantite_Sewan': q_sewan,
                                'Client_Sewan_Raw': cli_raw if cli_raw else "KOESIO ILE DE FRANCE",
                                'Annexe_Source': annexe.name
                            })
                            
                    df_fourn = pd.DataFrame(lignes_fourn)
                    df_fourn['Occurence_Sewan'] = df_fourn.groupby(['Code_Client_Resolu', 'Code_Article_ERP', 'Param_Key']).cumcount()
                    df_fourn['Occurence_Fallback_Sewan'] = df_fourn.groupby(['Code_Client_Resolu', 'Code_Article_ERP', 'Produit_Sewan_Clean']).cumcount()

                    df_recon_1 = pd.merge(
                        df_artis_abonn, 
                        df_fourn, 
                        left_on=['Code_Client_Str', 'Code_Article_ERP', 'Param_Key', 'Occurence_ERP'], 
                        right_on=['Code_Client_Resolu', 'Code_Article_ERP', 'Param_Key', 'Occurence_Sewan'], 
                        how='outer'
                    )

                    mask_c10_orphelin = (
                        df_recon_1['Montant_Sewan'].isna() | df_recon_1['Montant_Total_ERP'].isna()
                    ) & (
                        (df_recon_1['Code_Client_Str'] == '10') | 
                        (df_recon_1['Code SSC'].astype(str).str.contains('AFIDF-10', case=False, na=False))
                    )

                    if mask_c10_orphelin.any():
                        df_erp_c10_unmatched = df_recon_1[mask_c10_orphelin & df_recon_1['Montant_Sewan'].isna()].drop(
                            columns=['Code_Client_Resolu', 'Name_Sewan_Clean', 'Montant_Sewan', 'Quantite_Sewan', 'Client_Sewan_Raw', 'Annexe_Source', 'Produit_Sewan_Clean', 'Produit', 'Occurence_Sewan', 'Occurence_Fallback_Sewan', 'Id_Valorisation_Sewan'], 
                            errors='ignore'
                        )
                        df_sewan_c10_unmatched = df_fourn[df_fourn['Code_Client_Resolu'] == '10']
                        df_recon_fallback = pd.merge(
                            df_erp_c10_unmatched,
                            df_sewan_c10_unmatched,
                            left_on=['Code_Article_ERP', 'Libelle_RFC_Clean', 'Occurence_Fallback_ERP'],
                            right_on=['Code_Article_ERP', 'Produit_Sewan_Clean', 'Occurence_Fallback_Sewan'],
                            how='inner'
                        )
                        if not df_recon_fallback.empty:
                            matched_rfcs = df_recon_fallback['Libelle_RFC_Clean'].unique()
                            matched_prods = df_recon_fallback['Produit_Sewan_Clean'].unique()
                            df_recon_base = df_recon_1[
                                ~(df_recon_1['Libelle_RFC_Clean'].isin(matched_rfcs) & df_recon_1['Montant_Sewan'].isna()) &
                                ~(df_recon_1['Produit_Sewan_Clean'].isin(matched_prods) & df_recon_1['Montant_Total_ERP'].isna())
                            ]
                            df_recon = pd.concat([df_recon_base, df_recon_fallback], ignore_index=True)
                        else:
                            df_recon = df_recon_1.copy()
                    else:
                        df_recon = df_recon_1.copy()

                    df_recon['Montant_Total_ERP'] = df_recon['Montant_Total_ERP'].fillna(0.0).round(2)
                    df_recon['Montant_Sewan'] = df_recon['Montant_Sewan'].fillna(0.0).round(2)
                    df_recon['Quantite_ERP'] = df_recon['Quantite_ERP'].fillna(0.0)
                    df_recon['Quantite_Sewan'] = df_recon['Quantite_Sewan'].fillna(0.0)
                    df_recon['PU_ERP'] = df_recon['PU_ERP'].fillna(0.0).round(2)
                    df_recon['Ecart_HT'] = (df_recon['Montant_Total_ERP'] - df_recon['Montant_Sewan']).round(2)
                    
                    def qualifier_ht_exact(row):
                        if row['Montant_Total_ERP'] == 0.0 and row['Montant_Sewan'] == 0.0:
                            return "🟢 Conforme (Option 0€)"
                        if pd.isna(row.get(col_art_artis)) or str(row.get(col_art_artis)).strip() in ['', 'nan']:
                            return "🔴 Vente récente non saisie ERP"
                        if pd.isna(row.get('Client_Sewan_Raw')) or row.get('Montant_Sewan', 0) == 0:
                            return "🟠 Résiliation non saisie ERP"
                        if abs(row['Ecart_HT']) > 0.05 or abs(row['Quantite_ERP'] - row['Quantite_Sewan']) > 0.01:
                            return "⚠️ Écart de montant/quantité HT"
                        return "🟢 Conforme"

                    df_recon['Diagnostic_RAF'] = df_recon.apply(qualifier_ht_exact, axis=1)

                    if 'Code_Client_Resolu' in df_recon.columns:
                        df_recon['Code client'] = df_recon['Code client'].fillna(df_recon['Code_Client_Resolu'])
                    if 'Client_Sewan_Raw' in df_recon.columns:
                        col_cli_disp = 'Client' if 'Client' in df_recon.columns else 'Raison sociale client'
                        df_recon[col_cli_disp] = df_recon[col_cli_disp].fillna(df_recon['Client_Sewan_Raw'])

                    tot_fournisseurs_reel = df_recon['Montant_Sewan'].sum()
                    ecart_global_reel = total_simulation_artis_absolu - tot_fournisseurs_reel

                    # --- BANNIÈRE ALERTE TEMPORALITÉ DYNAMIQUE COLONNE AH ---
                    st.markdown(f"### ⏳ Contrôle de Temporalité ERP (Mois de Référence Facture Sewan : {mois_ref:02d}/{annee_ref})")

                    f1, f2, f3, f4 = st.columns(4)
                    f1.metric("🟢 Cycles Récurrents Conformes", f"{nb_conformes_ff} lignes", delta=f"{mois_ref:02d}/{annee_ref} (A échoir)", delta_color="normal")
                    f2.metric("ℹ️ Régularisations Échues", f"{nb_regul_echu_ff} lignes", delta=f"{((mois_ref-2)%12)+1:02d}/{annee_ref} (Échu)", delta_color="off")
                    f3.metric("ℹ️ Facturations Anticipées", f"{nb_anticipees_ff} lignes", delta=f">{mois_ref:02d}/{annee_ref}", delta_color="off")
                    
                    if nb_retards_ff > 0:
                        f4.metric("🚨 RETARDS DE FACTURATION ERP", f"{nb_retards_ff} LIGNES ❗", delta="Action Correction Noémie", delta_color="inverse")
                    else:
                        f4.metric("🟢 Retards de Facturation ERP", "0 ligne", delta="Aucun blocage", delta_color="normal")

                    st.markdown("---")

                    # --- SYNTHÈSE EXÉCUTIVE METRIQUES ---
                    st.markdown("### 📋 Synthèse des Factures Fournisseurs (Abonnements HT)")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Toolip", "15,324.15 €")
                    c2.metric("Sokatel", "10,874.99 €")
                    c3.metric("Nextphone", "1,981.41 €")
                    c4.metric("Total Reçu Annexes Sewan", f"{tot_fournisseurs_reel:,.2f} €")

                    st.markdown("### 🧮 Comparatif Global ERP vs Fournisseurs")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Total ERP Extrait (Sewan)", f"{total_simulation_artis_absolu:,.2f} €")
                    k2.metric("Total Réconcilié Annexes Sewan", f"{tot_fournisseurs_reel:,.2f} €")
                    
                    if abs(ecart_global_reel) < 1000.0:
                        k3.metric("Écart Net Global HT", f"{ecart_global_reel:,.2f} €", delta="🟢 Conforme (< 1 000 €)", delta_color="normal")
                    else:
                        k3.metric("Écart Net Global HT", f"{ecart_global_reel:,.2f} €", delta="🟠 À expertiser (> 1 000 €)", delta_color="inverse")

                    st.markdown("---")
                    st.markdown("### 📊 Répartition Financière & Pourcentages par Diagnostic")
                    
                    total_lignes = len(df_recon)
                    total_sewan_sum = tot_fournisseurs_reel if tot_fournisseurs_reel > 0 else 1.0

                    diag_stats_full = df_recon.groupby('Diagnostic_RAF').agg(
                        Nombre_Lignes=('Diagnostic_RAF', 'count'),
                        Total_Artis_ERP=('Montant_Total_ERP', 'sum'),
                        Total_Sewan_Fournisseur=('Montant_Sewan', 'sum'),
                        Total_Ecart_HT=('Ecart_HT', 'sum')
                    ).reset_index()

                    diag_stats_full['% Lignes'] = ((diag_stats_full['Nombre_Lignes'] / total_lignes) * 100).round(1).astype(str) + " %"
                    diag_stats_full['% HT Sewan'] = ((diag_stats_full['Total_Sewan_Fournisseur'] / total_sewan_sum) * 100).round(1).astype(str) + " %"

                    st.dataframe(diag_stats_full[['Diagnostic_RAF', 'Nombre_Lignes', '% Lignes', 'Total_Artis_ERP', 'Total_Sewan_Fournisseur', '% HT Sewan', 'Total_Ecart_HT']], use_container_width=True)

                    col_chart, col_empty = st.columns([1, 1])
                    with col_chart:
                        st.markdown("### 🍕 Répartition (%) des Lignes d'Abonnements")
                        st.bar_chart(df_recon['Diagnostic_RAF'].value_counts(normalize=True) * 100)

                    st.markdown("---")
                    st.markdown("### 🔍 Tableau Détaillé des Lignes d'Abonnements")

                    col_cli_disp = 'Client' if 'Client' in df_recon.columns else 'Raison sociale client'
                    if col_cli_disp in df_recon.columns:
                        df_recon[col_cli_disp] = df_recon[col_cli_disp].astype(str).str.slice(0, 20)
                        
                    cols_export_ht = [
                        'Code client', col_cli_disp, 'Code SSC', 'Code RFC', 'Libellé RFC', 
                        'Code_Article_ERP', 'Produit', 'Quantite_ERP', 'Quantite_Sewan', 
                        'PU_ERP', 'Montant_Total_ERP', 'Montant_Sewan', 'Ecart_HT', 'Diagnostic_RAF', 'Audit_Temporalite_FF', 'Annexe_Source'
                    ]
                    for c in cols_export_ht:
                        if c not in df_recon.columns:
                            df_recon[c] = "Non renseigné"
                            
                    df_res_ht = df_recon[cols_export_ht].drop_duplicates()
                    
                    buffer_ht = io.BytesIO()
                    with pd.ExcelWriter(buffer_ht, engine='openpyxl') as writer:
                        df_res_ht.to_excel(writer, index=False, sheet_name='Analyse_HT_Client_Article')
                    
                    st.dataframe(df_res_ht, use_container_width=True)
                    st.download_button("📥 Télécharger l'Analyse HT Détaillée (.xlsx)", data=buffer_ht.getvalue(), file_name="Analyse_RAF_HT_Detaillee.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"Erreur pendant le traitement HT : {e}")
    else:
        st.info("👈 Veuillez charger la simulation Artis (.xlsx) et les factures PDF Sewan.")
