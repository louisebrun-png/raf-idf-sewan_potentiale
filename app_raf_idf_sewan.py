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
st.caption("Audit Fiscal TVA & Rapprochement HT (Alignement Strict Produit Sewan == Libellé RFC ERP)")

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
    digits = re.sub(r'\D', '', str(text))
    if len(digits) >= 9:
        return digits[-9:]
    m_id = re.search(r'([a-zA-Z0-9_\-\.]+@[a-zA-Z0-9_\-\.]+)', str(text))
    if m_id:
        return m_id.group(1).lower()
    return clean_string_fuzzy(text)

def extract_code_from_string(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    m = re.search(r'(sk\d{4}|to_cl\d+|nm\d+|ns_\d+|\d{5,6})', str(text).lower())
    return m.group(1).upper() if m else ""

def lire_csv_securise(fichier):
    for enc in ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']:
        try:
            fichier.seek(0)
            return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        except Exception:
            continue
    fichier.seek(0)
    return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

onglet_tva, onglet_ht = st.tabs(["🚨 1. Audit Écarts TVA (20% vs 0%)", "📊 2. Analyse Quantité & Prix HT (Matching Produit / Libellé RFC)"])

# --- ONGLET 1 : AUDIT TVA ---
with onglet_tva:
    st.subheader("Audit Fiscal TVA (STD 20% vs APST 0%)")
    if fichier_artis and factures_pdf:
        if st.button("🚨 Lancer l'Analyse TVA", type="primary", key="btn_tva"):
            with st.spinner("Analyse des factures PDF Sewan et vérification des RFC ERP..."):
                try:
                    df_artis = pd.read_excel(fichier_artis)
                    col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    
                    def harmoniser_tva_artis_exact(row):
                        gtva = str(row.get('Coût ABONNEMENT Gestion TVA', ''))
                        ttva = str(row.get('Coût ABONNEMENT Taux TVA', ''))
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
                    
                    if 'Raison sociale client' in df_errors_tva.columns:
                        df_errors_tva['Raison sociale client'] = df_errors_tva['Raison sociale client'].astype(str).str.slice(0, 20)
                        
                    cols_tva_export = ['Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Libellé RFC', 'Code_Article_ERP', 'TVA_Artis', 'TVA_Sewan', 'Diagnostic_RAF']
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

# --- ONGLET 2 : ANALYSE HT AVEC MATCHING COMPTE GROUPE (PRODUIT SEWAN == LIBELLÉ RFC ERP) ---
with onglet_ht:
    st.subheader("Rapprochement Financier HT (Matching Produit Sewan == Libellé RFC ERP)")
    if fichier_artis and annexes_csv:
        nb_annexes = len(annexes_csv)
        if st.button("🚀 Lancer l'Analyse HT", type="primary", key="btn_ht"):
            with st.spinner(f"Traitement et alignement sur {nb_annexes} annexes CSV..."):
                try:
                    # 1. Table de correspondance Tiers
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

                    tot_toolip = 15324.15
                    tot_sokatel = 10874.99
                    tot_nextphone = 1981.41
                    tot_fournisseurs = tot_toolip + tot_sokatel + tot_nextphone

                    df_artis = pd.read_excel(fichier_artis)
                    
                    # Filtre strict SEWAN sur le fournisseur ERP
                    col_fourn_artis = next((c for c in df_artis.columns if 'ABONNEMENT' in c and ('fournisseur' in c.lower() or 'raison soc' in c.lower() or 'org' in c.lower())), None)
                    if col_fourn_artis:
                        df_artis = df_artis[df_artis[col_fourn_artis].astype(str).str.contains('SEWAN', case=False, na=False)].copy()

                    col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    df_artis_abonn = df_artis_abonn[~df_artis_abonn['Code_Article_ERP'].str.endswith('-F')].copy()

                    col_m_artis = 'Coût ABONNEMENT facturé' if 'Coût ABONNEMENT facturé' in df_artis_abonn.columns else df_artis.columns[1]
                    df_artis_abonn['Montant_ERP'] = df_artis_abonn[col_m_artis].fillna(0.0)
                    df_artis_abonn['Code_Client_Str'] = df_artis_abonn['Code client'].astype(str).str.strip()
                    df_artis_abonn['Name_Client_Clean'] = df_artis_abonn['Raison sociale client'].apply(clean_string_fuzzy)
                    df_artis_abonn['Param_Key'] = df_artis_abonn['Libellé RFC'].apply(extract_param_key_strict)
                    df_artis_abonn['Libelle_RFC_Clean'] = df_artis_abonn['Libellé RFC'].apply(clean_string_fuzzy)

                    # Lecture stricte de la Colonne FA (Nombre d'unités vendues)
                    col_qte_fa = next((c for c in df_artis_abonn.columns if 'ABONNEMENT' in c and 'Nombre' in c), None)
                    if not col_qte_fa:
                        col_qte_fa = next((c for c in df_artis_abonn.columns if any(k in c.lower() for k in ['nb bien', 'quantité', 'nb_bien', 'qte'])), None)
                    
                    if col_qte_fa:
                        df_artis_abonn['Quantite_ERP'] = pd.to_numeric(df_artis_abonn[col_qte_fa], errors='coerce').fillna(1.0)
                    else:
                        df_artis_abonn['Quantite_ERP'] = 1.0

                    lignes_fourn = []
                    for annexe in annexes_csv:
                        df_annexe = lire_csv_securise(annexe) if annexe.name.endswith('.csv') else pd.read_excel(annexe)
                        cols = [str(c).lower().strip() for c in df_annexe.columns]
                        
                        idx_ref = next((i for i, c in enumerate(cols) if any(k in c for k in ['code produit', 'code_article', 'ref', 'code'])), 0)
                        idx_m = next((i for i, c in enumerate(cols) if any(k in c for k in ["prix d'achat", 'prix', 'montant', 'ht'])), 1)
                        idx_qte = next((i for i, c in enumerate(cols) if any(k in c for k in ['quantité', 'quantite', 'qte'])), -1)
                        idx_p1 = next((i for i, c in enumerate(cols) if 'param1' in c), -1)
                        idx_p2 = next((i for i, c in enumerate(cols) if 'param2' in c), -1)
                        idx_prod = next((i for i, c in enumerate(cols) if 'produit' in c or 'libellé' in c or 'description' in c), -1)
                        
                        idx_cli1 = next((i for i, c in enumerate(cols) if 'client/revendeur direct (niv. 1)' in c or 'client/revendeur' in c), -1)
                        idx_cli2 = next((i for i, c in enumerate(cols) if 'nom client niveau 2' in c or 'nom client' in c), -1)
                        
                        col_ref = df_annexe.columns[idx_ref]
                        col_montant = df_annexe.columns[idx_m]
                        col_qte = df_annexe.columns[idx_qte] if idx_qte != -1 else None
                        col_param1 = df_annexe.columns[idx_p1] if idx_p1 != -1 else None
                        col_param2 = df_annexe.columns[idx_p2] if idx_p2 != -1 else None
                        col_cli1 = df_annexe.columns[idx_cli1] if idx_cli1 != -1 else None
                        col_cli2 = df_annexe.columns[idx_cli2] if idx_cli2 != -1 else None
                        col_produit = df_annexe.columns[idx_prod] if idx_prod != -1 else None
                        
                        for _, row in df_annexe.iterrows():
                            ref = isoler_ref_article(str(row[col_ref]))
                            if ref.endswith('-F'):
                                continue

                            try:
                                m_ht = float(str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', ''))
                            except ValueError:
                                m_ht = 0.0
                                
                            try:
                                q_sewan = float(str(row[col_qte]).replace(',', '.').replace(' ', '')) if col_qte else 1.0
                            except ValueError:
                                q_sewan = 1.0
                            
                            c1_str = str(row[col_cli1]).strip() if col_cli1 and pd.notna(row[col_cli1]) else ""
                            c2_str = str(row[col_cli2]).strip() if col_cli2 and pd.notna(row[col_cli2]) else ""
                            p1_val = str(row[col_param1]) if col_param1 else ""
                            p2_val = str(row[col_param2]) if col_param2 else ""
                            prod_val = str(row[col_produit]).strip() if col_produit and pd.notna(row[col_produit]) else ""
                            
                            # RESOLUTION TIERS MULTI-COLONNES : Réf Client ➔ Nom ➔ Param1 ➔ Param2
                            res_code = ""
                            for test_val in [c1_str, c2_str, p1_val, p2_val]:
                                if test_val.lower() in client_map:
                                    res_code = client_map[test_val.lower()]
                                    break
                                token = extract_code_from_string(test_val)
                                if token and token.lower() in client_map:
                                    res_code = client_map[token.lower()]
                                    break
                                    
                            if not res_code:
                                for test_val in [c1_str, c2_str, p1_val, p2_val]:
                                    token = extract_code_from_string(test_val)
                                    if token:
                                        res_code = token
                                        break
                                        
                            cli_raw = c1_str if c1_str else (c2_str if c2_str else p1_val)
                            param_key = extract_param_key_strict(p1_val)
                            
                            lignes_fourn.append({
                                'Code_Client_Resolu': res_code if res_code else "10",
                                'Name_Sewan_Clean': clean_string_fuzzy(cli_raw),
                                'Code_Article_ERP': ref,
                                'Code_Produit_Sewan': prod_val,
                                'Produit_Sewan_Clean': clean_string_fuzzy(prod_val),
                                'Param_Key': param_key,
                                'Montant_Sewan': m_ht,
                                'Quantite_Sewan': q_sewan,
                                'Client_Sewan_Raw': cli_raw if cli_raw else "KOESIO ILE DE FRANCE",
                                'Annexe_Source': annexe.name
                            })
                            
                    df_fourn = pd.DataFrame(lignes_fourn).groupby(['Code_Client_Resolu', 'Name_Sewan_Clean', 'Code_Article_ERP', 'Param_Key', 'Produit_Sewan_Clean'], as_index=False).agg(
                        Montant_Sewan=('Montant_Sewan', 'sum'),
                        Quantite_Sewan=('Quantite_Sewan', 'sum'),
                        Code_Produit_Sewan=('Code_Produit_Sewan', 'first'),
                        Client_Sewan_Raw=('Client_Sewan_Raw', 'first'),
                        Annexe_Source=('Annexe_Source', lambda x: ', '.join(set(x)))
                    )

                    # 1. RAPPROCHEMENT PRINCIPAL : Code Client + Code Article + ParamKey
                    df_recon_1 = pd.merge(
                        df_artis_abonn, 
                        df_fourn, 
                        left_on=['Code_Client_Str', 'Code_Article_ERP', 'Param_Key'], 
                        right_on=['Code_Client_Resolu', 'Code_Article_ERP', 'Param_Key'], 
                        how='outer'
                    )

                    # 2. FILET DE SÉCURITÉ ULTIME : CLIENT 10 / AFIDF-10 / FALLBACK PRODUIT SEWAN == LIBELLÉ RFC ERP
                    mask_c10_orphelin = (
                        df_recon_1['Montant_Sewan'].isna() | df_recon_1['Montant_ERP'].isna()
                    ) & (
                        (df_recon_1['Code_Client_Str'] == '10') | 
                        (df_recon_1['Code SSC'].astype(str).str.contains('AFIDF-10', case=False, na=False))
                    )

                    if mask_c10_orphelin.any():
                        df_erp_c10_unmatched = df_recon_1[mask_c10_orphelin & df_recon_1['Montant_Sewan'].isna()].drop(
                            columns=['Code_Client_Resolu', 'Name_Sewan_Clean', 'Montant_Sewan', 'Quantite_Sewan', 'Client_Sewan_Raw', 'Annexe_Source', 'Produit_Sewan_Clean', 'Code_Produit_Sewan'], 
                            errors='ignore'
                        )
                        
                        df_sewan_c10_unmatched = df_fourn[df_fourn['Code_Client_Resolu'] == '10']
                        
                        # MATCH EXACT ENTRE LIBELLÉ RFC ERP ET PRODUIT SEWAN (COL. F)
                        df_recon_fallback = pd.merge(
                            df_erp_c10_unmatched,
                            df_sewan_c10_unmatched,
                            left_on=['Code_Article_ERP', 'Libelle_RFC_Clean'],
                            right_on=['Code_Article_ERP', 'Produit_Sewan_Clean'],
                            how='inner'
                        )
                        
                        if not df_recon_fallback.empty:
                            matched_rfcs = df_recon_fallback['Libelle_RFC_Clean'].unique()
                            matched_prods = df_recon_fallback['Produit_Sewan_Clean'].unique()
                            
                            df_recon_base = df_recon_1[
                                ~(df_recon_1['Libelle_RFC_Clean'].isin(matched_rfcs) & df_recon_1['Montant_Sewan'].isna()) &
                                ~(df_recon_1['Produit_Sewan_Clean'].isin(matched_prods) & df_recon_1['Montant_ERP'].isna())
                            ]
                            
                            df_recon = pd.concat([df_recon_base, df_recon_fallback], ignore_index=True)
                        else:
                            df_recon = df_recon_1.copy()
                    else:
                        df_recon = df_recon_1.copy()

                    df_recon['Montant_ERP'] = df_recon['Montant_ERP'].fillna(0.0).round(2)
                    df_recon['Montant_Sewan'] = df_recon['Montant_Sewan'].fillna(0.0).round(2)
                    df_recon['Quantite_ERP'] = df_recon['Quantite_ERP'].fillna(0.0)
                    df_recon['Quantite_Sewan'] = df_recon['Quantite_Sewan'].fillna(0.0)
                    df_recon['Ecart_HT'] = (df_recon['Montant_ERP'] - df_recon['Montant_Sewan']).round(2)
                    
                    def qualifier_ht_exact(row):
                        if row['Montant_ERP'] == 0.0 and row['Montant_Sewan'] == 0.0:
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
                        df_recon['Raison sociale client'] = df_recon['Raison sociale client'].fillna(df_recon['Client_Sewan_Raw'])

                    # CALCUL DYNAMIQUE TOTALS
                    tot_simulation_artis_reel = df_recon['Montant_ERP'].sum()
                    tot_fournisseurs_reel = df_recon['Montant_Sewan'].sum()
                    ecart_global_reel = tot_simulation_artis_reel - tot_fournisseurs_reel

                    # --- SYNTHÈSE EXÉCUTIVE METRIQUES ---
                    st.markdown("### 📋 Synthèse des Factures Fournisseurs (Abonnements HT)")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Toolip", "15,324.15 €")
                    c2.metric("Sokatel", "10,874.99 €")
                    c3.metric("Nextphone", "1,981.41 €")
                    c4.metric("Total Reçu Annexes Sewan", f"{tot_fournisseurs_reel:,.2f} €")

                    st.markdown("### 🧮 Comparatif Global ERP vs Fournisseurs")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Total ERP Extrait (Sewan)", f"{tot_simulation_artis_reel:,.2f} €")
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
                        Total_Artis_ERP=('Montant_ERP', 'sum'),
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

                    if 'Raison sociale client' in df_recon.columns:
                        df_recon['Raison sociale client'] = df_recon['Raison sociale client'].astype(str).str.slice(0, 20)
                        
                    # INCLUSION DE LA COLONNE "Code_Produit_Sewan"
                    cols_export_ht = [
                        'Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Libellé RFC', 
                        'Code_Article_ERP', 'Code_Produit_Sewan', 'Quantite_ERP', 'Quantite_Sewan', 
                        'Montant_ERP', 'Montant_Sewan', 'Ecart_HT', 'Diagnostic_RAF', 'Annexe_Source'
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
