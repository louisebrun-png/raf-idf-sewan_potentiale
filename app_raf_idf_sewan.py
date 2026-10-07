# ==============================================================================
# CABINET POTENTIALE — OUTIL INTERNE RAF (KOESIO IDF x SEWAN)
# Auteur : Louise Brun — Potentiale
# ==============================================================================

import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Potentiale — RAF Koesio IDF", page_icon="💼", layout="wide")

st.title("💼 Potentiale — Outil Interne RAF (Koesio IDF x Sewan)")
st.caption("Audit Financier HT & Cohérence TVA Abonnements Fournisseurs")

# Sidebar - Importation
st.sidebar.header("📁 Importation des Documents")
fichier_artis = st.sidebar.file_uploader("1. Simulation Achat Artis (.xlsx)", type=["xlsx"])
fichier_ref = st.sidebar.file_uploader("2. Table de Correspondance (.xlsx)", type=["xlsx"])
annexes_csv = st.sidebar.file_uploader("3. Annexes Fournisseurs (.csv / .xlsx)", type=["csv", "xlsx"], accept_multiple_files=True)
factures_pdf = st.sidebar.file_uploader("4. Factures PDF Sewan (Optionnel)", type=["pdf"], accept_multiple_files=True)

def isoler_ref_article(libelle):
    if not isinstance(libelle, str):
        return ""
    match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', libelle)
    return match.group(1) if match else libelle.split(' - ')[0].strip()

def lire_csv_securise(fichier):
    encodings = ['utf-8', 'iso-8859-1', 'cp1252', 'latin1']
    for enc in encodings:
        try:
            fichier.seek(0)
            return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        except (UnicodeDecodeError, Exception):
            continue
    fichier.seek(0)
    return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

onglet_ht, onglet_tva = st.tabs(["📊 1. Analyse Quantité & Prix HT (Global)", "🚨 2. Audit Écarts TVA"])

if fichier_artis and (annexes_csv or factures_pdf):
    
    # --- ONGLET 1 : HT ---
    with onglet_ht:
        st.subheader("Rapprochement Financier HT (Conformes & Écarts)")
        if st.button("🚀 Lancer l'Analyse HT & Générer l'Excel", type="primary", key="btn_ht"):
            with st.spinner("Analyse du périmètre HT en cours..."):
                df_artis = pd.read_excel(fichier_artis)
                df_artis_abonn = df_artis.dropna(subset=['Coût ABONNEMENT article']).copy()
                df_artis_abonn['Code_Article_ERP'] = df_artis_abonn['Coût ABONNEMENT article'].apply(isoler_ref_article)
                df_artis_abonn['Montant_ERP'] = df_artis_abonn['Coût ABONNEMENT facturé'].fillna(0.0)
                
                lignes_fourn = []
                
                # Traitement de TOUTES les annexes CSV/Excel chargées avec secours encodage
                if annexes_csv:
                    for annexe in annexes_csv:
                        if annexe.name.endswith('.csv'):
                            df_annexe = lire_csv_securise(annexe)
                        else:
                            df_annexe = pd.read_excel(annexe)
                        
                        col_ref = next((c for c in df_annexe.columns if any(k in str(c).lower() for k in ['ref', 'code', 'article'])), df_annexe.columns[0])
                        col_montant = next((c for c in df_annexe.columns if any(k in str(c).lower() for k in ['montant', 'ht', 'prix', 'total'])), df_annexe.columns[1])
                        col_client = next((c for c in df_annexe.columns if any(k in str(c).lower() for k in ['client', 'raison'])), df_annexe.columns[-1])
                        
                        for _, row in df_annexe.iterrows():
                            ref = isoler_ref_article(str(row[col_ref]))
                            try:
                                m_ht = float(str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', ''))
                            except ValueError:
                                m_ht = 0.0
                            
                            lignes_fourn.append({
                                'Code_Article_ERP': ref,
                                'Montant_Sewan': m_ht,
                                'Client_Sewan': str(row[col_client])
                            })
                
                # Complément par PDF si fournis
                if factures_pdf:
                    for pdf in factures_pdf:
                        reader = pypdf.PdfReader(pdf)
                        for page in reader.pages:
                            text = page.extract_text() or ""
                            lines = text.split('\n')
                            for line in lines:
                                match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', line)
                                if match:
                                    ref = match.group(1)
                                    montants = re.findall(r'\d+[\.,]\d{2}', line)
                                    m_ht = 0.0
                                    if montants:
                                        m_ht = float(montants[-2].replace(',', '.')) if len(montants) >= 2 else float(montants[0].replace(',', '.'))
                                    lignes_fourn.append({
                                        'Code_Article_ERP': ref,
                                        'Montant_Sewan': m_ht,
                                        'Client_Sewan': pdf.name.replace('.pdf', '')
                                    })
                                
                df_fourn = pd.DataFrame(lignes_fourn).groupby(['Code_Article_ERP', 'Client_Sewan'], as_index=False)['Montant_Sewan'].sum()
                
                df_merged = pd.merge(df_artis_abonn, df_fourn, on='Code_Article_ERP', how='outer')
                df_merged['Montant_ERP'] = df_merged['Montant_ERP'].fillna(0.0)
                df_merged['Montant_Sewan'] = df_merged['Montant_Sewan'].fillna(0.0)
                df_merged['Ecart'] = (df_merged['Montant_ERP'] - df_merged['Montant_Sewan']).round(2)
                
                def qualifier_ht(row):
                    if pd.isna(row['Coût ABONNEMENT article']):
                        return "🔴 Absent simulation ERP"
                    if pd.isna(row['Client_Sewan']):
                        return "🟠 Absent annexe fournisseur"
                    if abs(row['Ecart']) > 0.05:
                        return "⚠️ Écart de montant HT"
                    return "🟢 Conforme"

                df_merged['Diagnostic_RAF'] = df_merged.apply(qualifier_ht, axis=1)
                
                cols_export_ht = {
                    'Code client': 'Code_Client',
                    'Raison sociale client': 'Client_ERP',
                    'Client_Sewan': 'Client_Sewan',
                    'Code SSC': 'Code_SSC',
                    'Code RFC': 'Code_RFC',
                    'Code_Article_ERP': 'Code_Article_ERP',
                    'Montant_Sewan': 'Montant_Sewan',
                    'Montant_ERP': 'Montant_ERP',
                    'Ecart': 'Ecart',
                    'Diagnostic_RAF': 'Diagnostic_RAF'
                }
                
                for col in cols_export_ht.keys():
                    if col not in df_merged.columns:
                        df_merged[col] = "Non renseigné"
                        
                df_res_ht = df_merged[list(cols_export_ht.keys())].rename(columns=cols_export_ht).drop_duplicates()
                
                buffer_ht = io.BytesIO()
                with pd.ExcelWriter(buffer_ht, engine='openpyxl') as writer:
                    df_res_ht.to_excel(writer, index=False, sheet_name='Analyse_HT_Global')
                
                st.dataframe(df_res_ht, use_container_width=True)
                st.download_button(
                    label="📥 Télécharger l'Analyse HT (.xlsx)",
                    data=buffer_ht.getvalue(),
                    file_name="Analyse_RAF_HT_Koesio_Sewan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

    # --- ONGLET 2 : TVA ---
    with onglet_tva:
        st.subheader("Audit Fiscale TVA (STD 20% vs APST 0%)")
        if st.button("🚨 Lancer l'Analyse TVA & Générer l'Excel", type="primary", key="btn_tva"):
            with st.spinner("Audit des régimes TVA en cours..."):
                df_artis = pd.read_excel(fichier_artis)
                df_artis_abonn = df_artis.dropna(subset=['Coût ABONNEMENT article']).copy()
                df_artis_abonn['Code_Article_ERP'] = df_artis_abonn['Coût ABONNEMENT article'].apply(isoler_ref_article)
                
                def harmoniser_tva_artis(row):
                    gtva = str(row['Coût ABONNEMENT Gestion TVA'])
                    ttva = str(row['Coût ABONNEMENT Taux TVA'])
                    return 'APST (0%)' if ('APST' in gtva or '0.00' in ttva) else 'S (20%)'

                df_artis_abonn['TVA_Artis'] = df_artis_abonn.apply(harmoniser_tva_artis, axis=1)
                
                lignes_tva = []
                if factures_pdf:
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
                                
                df_sewan_tva = pd.DataFrame(lignes_tva).drop_duplicates(subset=['Code_Article_ERP', 'TVA_Sewan']) if lignes_tva else pd.DataFrame(columns=['Code_Article_ERP', 'TVA_Sewan'])
                df_merged_tva = pd.merge(df_artis_abonn, df_sewan_tva, on='Code_Article_ERP', how='inner')
                df_errors_tva = df_merged_tva[df_merged_tva['TVA_Artis'] != df_merged_tva['TVA_Sewan']].copy()
                df_errors_tva['Diagnostic_RAF'] = "🚨 Divergence TVA"
                df_errors_tva['Action_Preconisee'] = "Aligner la Gestion TVA dans Artis ERP"
                
                cols_export_tva = {
                    'Code client': 'Code_Client',
                    'Raison sociale client': 'Client_ERP',
                    'Code SSC': 'Code_SSC',
                    'Code RFC': 'Code_RFC',
                    'Code_Article_ERP': 'Code_Article_ERP',
                    'Coût ABONNEMENT article': 'Designation_Article',
                    'TVA_Artis': 'TVA_Artis',
                    'TVA_Sewan': 'TVA_Sewan',
                    'Diagnostic_RAF': 'Diagnostic_RAF',
                    'Action_Preconisee': 'Action_Preconisee'
                }
                for col in cols_export_tva.keys():
                    if col not in df_errors_tva.columns:
                        df_errors_tva[col] = "Non renseigné"
                        
                df_res_tva = df_errors_tva[list(cols_export_tva.keys())].rename(columns=cols_export_tva).drop_duplicates()
                
                buffer_tva = io.BytesIO()
                with pd.ExcelWriter(buffer_tva, engine='openpyxl') as writer:
                    df_res_tva.to_excel(writer, index=False, sheet_name='Anomalies_TVA')
                
                st.dataframe(df_res_tva, use_container_width=True)
                st.download_button(
                    label="📥 Télécharger le Rapport TVA (.xlsx)",
                    data=buffer_tva.getvalue(),
                    file_name="Anomalies_TVA_Koesio_Sewan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
else:
    st.info("👈 Veuillez charger la simulation Artis ERP ainsi que les annexes CSV (ou les PDF) dans le menu latéral pour démarrer.")
