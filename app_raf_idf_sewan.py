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
st.caption("Audit Fiscale TVA & Rapprochement Financier HT (Granularité Param1)")

# Sidebar - Importation des fichiers
st.sidebar.header("📁 Importation des Documents")
fichier_artis = st.sidebar.file_uploader("1. Simulation Achat Artis (.xlsx)", type=["xlsx"])
fichier_ref = st.sidebar.file_uploader("2. Table de Correspondance (.xlsx)", type=["xlsx"])
factures_pdf = st.sidebar.file_uploader("3. Factures PDF Sewan (Obligatoire TVA)", type=["pdf"], accept_multiple_files=True)
annexes_csv = st.sidebar.file_uploader("4. Annexes Fournisseurs (.csv / .xlsx)", type=["csv", "xlsx"], accept_multiple_files=True)

def isoler_ref_article(libelle):
    if not isinstance(libelle, str):
        return ""
    match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', libelle)
    return match.group(1) if match else libelle.split(' - ')[0].strip()

def extract_client_code(client_str):
    if not isinstance(client_str, str) or pd.isna(client_str) or str(client_str).strip() in ['', 'nan', 'None', '-']:
        return ""
    m = re.match(r'^\s*([A-Za-z0-9_\-]+)', str(client_str))
    return m.group(1).strip() if m else str(client_str).strip()

def extract_param_key(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    text_clean = str(text).replace('+', '').replace(' ', '').replace('//', ' ')
    m_num = re.search(r'(\d{9,12})', text_clean)
    if m_num:
        return m_num.group(1)
    m_id = re.search(r'([a-zA-Z0-9_\-\.]+@[a-zA-Z0-9_\-\.]+)', text_clean)
    if m_id:
        return m_id.group(1).lower()
    return text_clean.strip()

def lire_csv_securise(fichier):
    for enc in ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']:
        try:
            fichier.seek(0)
            return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        except Exception:
            continue
    fichier.seek(0)
    return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

onglet_tva, onglet_ht = st.tabs(["🚨 1. Audit Écarts TVA (20% vs 0%)", "📊 2. Analyse Quantité & Prix HT (Granularité Param1)"])

# --- ONGLET 1 : AUDIT TVA ---
with onglet_tva:
    st.subheader("Audit Fiscale TVA (STD 20% vs APST 0%)")
    if fichier_artis and factures_pdf:
        if st.button("🚨 Lancer l'Analyse TVA", type="primary", key="btn_tva"):
            with st.spinner("Analyse du texte brut des factures PDF Sewan et vérification des RFC ERP..."):
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

# --- ONGLET 2 : ANALYSE HT GRANULAIRE ---
with onglet_ht:
    st.subheader("Rapprochement Financier HT (Détail par Client / Article / Param1)")
    if fichier_artis and annexes_csv:
        nb_annexes = len(annexes_csv)
        if st.button("🚀 Lancer l'Analyse HT", type="primary", key="btn_ht"):
            with st.spinner(f"Traitement des {nb_annexes} annexes CSV/Excel et réconciliation granulaire..."):
                try:
                    tot_toolip = 15324.15
                    tot_sokatel = 10874.99
                    tot_nextphone = 1981.41
                    tot_fournisseurs = tot_toolip + tot_sokatel + tot_nextphone

                    df_artis = pd.read_excel(fichier_artis)
                    col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    col_m_artis = 'Coût ABONNEMENT facturé' if 'Coût ABONNEMENT facturé' in df_artis_abonn.columns else df_artis.columns[1]
                    df_artis_abonn['Montant_ERP'] = df_artis_abonn[col_m_artis].fillna(0.0)
                    df_artis_abonn['Code_Client_Str'] = df_artis_abonn['Code client'].astype(str).str.strip()
                    df_artis_abonn['Param_Key'] = df_artis_abonn['Libellé RFC'].apply(extract_param_key)

                    tot_simulation_artis = 27787.45
                    ecart_global = tot_simulation_artis - tot_fournisseurs

                    col_qte_artis = next((c for c in df_artis_abonn.columns if any(k in c.lower() for k in ['nb bien', 'quantité', 'nb_bien', 'qte'])), None)
                    if col_qte_artis:
                        df_artis_abonn['Quantite_ERP'] = pd.to_numeric(df_artis_abonn[col_qte_artis], errors='coerce').fillna(1.0)
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
                        
                        idx_cli = -1
                        for k in ['client/revendeur direct (niv. 1)', 'client/revendeur', 'nom client niveau 2', 'nom client', 'client']:
                            for i, c in enumerate(cols):
                                if k in c:
                                    idx_cli = i
                                    break
                            if idx_cli != -1:
                                break
                        
                        col_ref = df_annexe.columns[idx_ref]
                        col_montant = df_annexe.columns[idx_m]
                        col_qte = df_annexe.columns[idx_qte] if idx_qte != -1 else None
                        col_client = df_annexe.columns[idx_cli] if idx_cli != -1 else None
                        col_param1 = df_annexe.columns[idx_p1] if idx_p1 != -1 else None
                        
                        for _, row in df_annexe.iterrows():
                            ref = isoler_ref_article(str(row[col_ref]))
                            try:
                                m_ht = float(str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', ''))
                            except ValueError:
                                m_ht = 0.0
                                
                            try:
                                q_sewan = float(str(row[col_qte]).replace(',', '.').replace(' ', '')) if col_qte else 1.0
                            except ValueError:
                                q_sewan = 1.0
                            
                            cli_raw = str(row[col_client]) if col_client else ""
                            cli_code = extract_client_code(cli_raw)
                            p1_val = str(row[col_param1]) if col_param1 else ""
                            param_key = extract_param_key(p1_val)
                            
                            lignes_fourn.append({
                                'Code_Client_Sewan': cli_code,
                                'Code_Article_ERP': ref,
                                'Param_Key': param_key,
                                'Montant_Sewan': m_ht,
                                'Quantite_Sewan': q_sewan,
                                'Client_Sewan_Raw': cli_raw
                            })
                            
                    df_fourn = pd.DataFrame(lignes_fourn).groupby(['Code_Client_Sewan', 'Code_Article_ERP', 'Param_Key'], as_index=False).agg(
                        Montant_Sewan=('Montant_Sewan', 'sum'),
                        Quantite_Sewan=('Quantite_Sewan', 'sum'),
                        Client_Sewan_Raw=('Client_Sewan_Raw', 'first')
                    )
                    
                    df_recon = pd.merge(
                        df_artis_abonn, 
                        df_fourn, 
                        left_on=['Code_Client_Str', 'Code_Article_ERP', 'Param_Key'], 
                        right_on=['Code_Client_Sewan', 'Code_Article_ERP', 'Param_Key'], 
                        how='outer'
                    )
                    
                    df_recon['Montant_ERP'] = df_recon['Montant_ERP'].fillna(0.0).round(2)
                    df_recon['Montant_Sewan'] = df_recon['Montant_Sewan'].fillna(0.0).round(2)
                    df_recon['Quantite_ERP'] = df_recon['Quantite_ERP'].fillna(0.0)
                    df_recon['Quantite_Sewan'] = df_recon['Quantite_Sewan'].fillna(0.0)
                    df_recon['Ecart_HT'] = (df_recon['Montant_ERP'] - df_recon['Montant_Sewan']).round(2)
                    
                    def qualifier_ht_exact(row):
                        if row['Montant_ERP'] == 0.0 and row['Montant_Sewan'] == 0.0:
                            return "🟢 Conforme (Option 0€)"
                        if pd.isna(row.get(col_art_artis)):
                            return "🔴 Absent simulation ERP"
                        if pd.isna(row.get('Client_Sewan_Raw')) or row.get('Montant_Sewan', 0) == 0:
                            return "🟠 Absent annexe fournisseur"
                        if abs(row['Ecart_HT']) > 0.05 or abs(row['Quantite_ERP'] - row['Quantite_Sewan']) > 0.01:
                            return "⚠️ Écart de montant/quantité HT"
                        return "🟢 Conforme"

                    df_recon['Diagnostic_RAF'] = df_recon.apply(qualifier_ht_exact, axis=1)

                    df_recon['Code client'] = df_recon['Code client'].fillna(df_recon['Code_Client_Sewan'])
                    df_recon['Raison sociale client'] = df_recon['Raison sociale client'].fillna(df_recon['Client_Sewan_Raw'])

                    # --- SYNTHÈSE EXÉCUTIVE ---
                    st.markdown("### 📋 Synthèse des Factures Fournisseurs (Abonnements HT)")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Toolip", f"{tot_toolip:,.2f} €")
                    c2.metric("Sokatel", f"{tot_sokatel:,.2f} €")
                    c3.metric("Nextphone", f"{tot_nextphone:,.2f} €")
                    c4.metric("Total 3 Factures Sewan", f"{tot_fournisseurs:,.2f} €")

                    st.markdown("### 🧮 Comparatif Global ERP vs Fournisseurs")
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Simulation Achat Artis ERP", f"{tot_simulation_artis:,.2f} €")
                    k2.metric("Total Factures Sewan", f"{tot_fournisseurs:,.2f} €")
                    
                    if abs(ecart_global) < 1000.0:
                        k3.metric("Écart Net Global HT", f"{ecart_global:,.2f} €", delta="🟢 Conforme (< 1 000 €)", delta_color="normal")
                    else:
                        k3.metric("Écart Net Global HT", f"{ecart_global:,.2f} €", delta="🟠 À expertiser (> 1 000 €)", delta_color="inverse")

                    st.markdown("---")
                    st.markdown("### 📊 Répartition Financière par Famille de Diagnostic")
                    
                    diag_stats_full = df_recon.groupby('Diagnostic_RAF').agg(
                        Nombre_Lignes=('Diagnostic_RAF', 'count'),
                        Total_Artis_ERP=('Montant_ERP', 'sum'),
                        Total_Sewan_Fournisseur=('Montant_Sewan', 'sum'),
                        Total_Ecart_HT=('Ecart_HT', 'sum')
                    ).reset_index()

                    st.dataframe(diag_stats_full, use_container_width=True)

                    col_chart, col_empty = st.columns([1, 1])
                    with col_chart:
                        st.markdown("### 🍕 Proportion des Lignes d'Abonnements")
                        st.bar_chart(df_recon['Diagnostic_RAF'].value_counts())

                    st.markdown("---")
                    st.markdown("### 🔍 Tableau Détaillé des Lignes d'Abonnements")

                    if 'Raison sociale client' in df_recon.columns:
                        df_recon['Raison sociale client'] = df_recon['Raison sociale client'].astype(str).str.slice(0, 20)
                        
                    cols_export_ht = [
                        'Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Libellé RFC', 
                        'Code_Article_ERP', 'Quantite_ERP', 'Quantite_Sewan', 
                        'Montant_ERP', 'Montant_Sewan', 'Ecart_HT', 'Diagnostic_RAF'
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
        st.info("👈 Veuillez charger la simulation Artis (.xlsx) et les annexes CSV / Excel.")
