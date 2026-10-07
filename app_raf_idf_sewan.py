# ==============================================================================
# CABINET POTENTIALE — OUTIL INTERNE RAF (KOESIO IDF x SEWAN)
# ==============================================================================

import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Potentiale — RAF Koesio IDF", page_icon="💼", layout="wide")

st.title("💼 Potentiale — Outil Interne RAF (Koesio IDF x Sewan)")
st.caption("Audit Financier HT & Cohérence TVA Abonnements Fournisseurs")

# Sidebar - Importation des fichiers
st.sidebar.header("📁 Importation des Documents")
fichier_artis = st.sidebar.file_uploader("1. Simulation Achat Artis (.xlsx)", type=["xlsx"])
fichier_ref = st.sidebar.file_uploader("2. Table de Correspondance (.xlsx)", type=["xlsx"])
annexes_csv = st.sidebar.file_uploader("3. Annexes Fournisseurs (.csv / .xlsx)", type=["csv", "xlsx"], accept_multiple_files=True)
factures_pdf = st.sidebar.file_uploader("4. Factures PDF Sewan (Obligatoire TVA)", type=["pdf"], accept_multiple_files=True)

def isoler_ref_article(libelle):
    if not isinstance(libelle, str):
        return ""
    match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', libelle)
    return match.group(1) if match else libelle.split(' - ')[0].strip()

def lire_csv_securise(fichier):
    # Gestion des encodages Sewan (latin1 / utf-8)
    for enc in ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']:
        try:
            fichier.seek(0)
            return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        except Exception:
            continue
    fichier.seek(0)
    return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

onglet_ht, onglet_tva = st.tabs(["📊 1. Analyse Quantité & Prix HT (Global)", "🚨 2. Audit Écarts TVA"])

# --- ONGLET 1 : ANALYSE HT ---
with onglet_ht:
    st.subheader("Rapprochement Financier HT (Conformes & Écarts)")
    if fichier_artis and annexes_csv:
        if st.button("🚀 Lancer l'Analyse HT", type="primary", key="btn_ht"):
            with st.spinner("Traitement des 5 annexes CSV et de la simulation Artis ERP..."):
                try:
                    df_artis = pd.read_excel(fichier_artis)
                    
                    # Filtrage des lignes abonnements Artis
                    col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    
                    col_m_artis = 'Coût ABONNEMENT facturé' if 'Coût ABONNEMENT facturé' in df_artis.columns else df_artis.columns[1]
                    df_artis_abonn['Montant_ERP'] = df_artis_abonn[col_m_artis].fillna(0.0)
                    
                    lignes_fourn = []
                    for annexe in annexes_csv:
                        if annexe.name.endswith('.csv'):
                            df_annexe = lire_csv_securise(annexe)
                        else:
                            df_annexe = pd.read_excel(annexe)
                        
                        cols = [str(c).lower().strip() for c in df_annexe.columns]
                        
                        # Recherche souple des colonnes Sewan
                        idx_ref = next((i for i, c in enumerate(cols) if any(k in c for k in ['code produit', 'code_article', 'ref', 'code'])), 0)
                        idx_m = next((i for i, c in enumerate(cols) if any(k in c for k in ["prix d'achat", 'prix', 'montant', 'ht'])), 1)
                        idx_cli = next((i for i, c in enumerate(cols) if any(k in c for k in ['nom client', 'client', 'nom utilisateur', 'description'])), -1)
                        
                        col_ref = df_annexe.columns[idx_ref]
                        col_montant = df_annexe.columns[idx_m]
                        col_client = df_annexe.columns[idx_cli]
                        
                        for _, row in df_annexe.iterrows():
                            ref = isoler_ref_article(str(row[col_ref]))
                            try:
                                val_str = str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', '')
                                m_ht = float(val_str)
                            except ValueError:
                                m_ht = 0.0
                            
                            lignes_fourn.append({
                                'Code_Article_ERP': ref,
                                'Montant_Sewan': m_ht,
                                'Client_Sewan': str(row[col_client])
                            })
                            
                    df_fourn = pd.DataFrame(lignes_fourn).groupby(['Code_Article_ERP', 'Client_Sewan'], as_index=False)['Montant_Sewan'].sum()
                    df_merged = pd.merge(df_artis_abonn, df_fourn, on='Code_Article_ERP', how='outer')
                    df_merged['Montant_ERP'] = df_merged['Montant_ERP'].fillna(0.0)
                    df_merged['Montant_Sewan'] = df_merged['Montant_Sewan'].fillna(0.0)
                    df_merged['Ecart'] = (df_merged['Montant_ERP'] - df_merged['Montant_Sewan']).round(2)
                    
                    def qualifier_ht(row):
                        if pd.isna(row.get(col_art_artis)):
                            return "🔴 Absent simulation ERP"
                        if pd.isna(row.get('Client_Sewan')):
                            return "🟠 Absent annexe fournisseur"
                        if abs(row['Ecart']) > 0.05:
                            return "⚠️ Écart de montant HT"
                        return "🟢 Conforme"

                    df_merged['Diagnostic_RAF'] = df_merged.apply(qualifier_ht, axis=1)
                    
                    cols_export = ['Code client', 'Raison sociale client', 'Client_Sewan', 'Code SSC', 'Code RFC', 'Code_Article_ERP', 'Montant_Sewan', 'Montant_ERP', 'Ecart', 'Diagnostic_RAF']
                    for col in cols_export:
                        if col not in df_merged.columns:
                            df_merged[col] = "Non renseigné"
                            
                    df_res = df_merged[cols_export].drop_duplicates()
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_res.to_excel(writer, index=False, sheet_name='Analyse_HT')
                    
                    st.success("Analyse HT terminée avec succès !")
                    st.dataframe(df_res, use_container_width=True)
                    st.download_button("📥 Télécharger l'Analyse HT (.xlsx)", data=buffer.getvalue(), file_name="Analyse_RAF_HT_Koesio_Sewan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"Erreur pendant le traitement HT : {e}")
    else:
        st.info("👈 Charge la simulation Artis (.xlsx) et tes annexes CSV dans le menu à gauche.")

# --- ONGLET 2 : AUDIT TVA (PDF) ---
with onglet_tva:
    st.subheader("Audit Fiscale TVA (STD 20% vs APST 0%)")
    if fichier_artis and factures_pdf:
        if st.button("🚨 Lancer l'Analyse TVA", type="primary", key="btn_tva"):
            with st.spinner("Parsing du texte brut des factures PDF Sewan..."):
                try:
                    df_artis = pd.read_excel(fichier_artis)
                    col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    
                    def harmoniser_tva_artis(row):
                        gtva = str(row.get('Coût ABONNEMENT Gestion TVA', ''))
                        ttva = str(row.get('Coût ABONNEMENT Taux TVA', ''))
                        return 'APST (0%)' if ('APST' in gtva or '0.00' in ttva or '0%' in ttva) else 'S (20%)'

                    df_artis_abonn['TVA_Artis'] = df_artis_abonn.apply(harmoniser_tva_artis, axis=1)
                    
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
                    
                    cols_tva_export = ['Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Code_Article_ERP', 'TVA_Artis', 'TVA_Sewan', 'Diagnostic_RAF']
                    for c in cols_tva_export:
                        if c not in df_errors_tva.columns:
                            df_errors_tva[c] = "Non renseigné"
                            
                    df_res_tva = df_errors_tva[cols_tva_export].drop_duplicates()
                    
                    buffer_tva = io.BytesIO()
                    with pd.ExcelWriter(buffer_tva, engine='openpyxl') as writer:
                        df_res_tva.to_excel(writer, index=False, sheet_name='Anomalies_TVA')
                    
                    st.success("Audit TVA terminé !")
                    st.dataframe(df_res_tva, use_container_width=True)
                    st.download_button("📥 Télécharger le Rapport TVA (.xlsx)", data=buffer_tva.getvalue(), file_name="Anomalies_TVA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"Erreur pendant l'analyse TVA : {e}")
    else:
        st.info("👈 Charge les 3 factures PDF Sewan pour exécuter l'audit TVA.")
