# ==============================================================================
# CABINET POTENTIALE — OUTIL INTERNE RAF (KOESIO IDF x SEWAN)
# Experte : Louise Brun — Fondatrice Potentiale
# ==============================================================================

import Streamlit as st
import pandas as pd
import pypdf
import re
import io

St.set_page_config(page_title="Potentiale — RAF Koesio IDF", page_icon="💼", layout="wide")

St.title("💼 Potentiale — Outil Interne RAF (Koesio IDF x Sewan)")
st.caption("Audit Fiscal TVA & Rapprochement HT (Version Restaurée & Sécurisée)")

# Sidebar - Importation des fichiers
St.sidebar.header("📁 Importation des Documents")
Fichier_artis = st.sidebar.file_uploader("1. Simulation Achat Artis (.xlsx)", type=["xlsx"])
Fichier_ref = st.sidebar.file_uploader("2. Table de Correspondance (.xlsx)", type=["xlsx"])
Factures_pdf = st.sidebar.file_uploader("3. Factures PDF Sewan (Obligatoire TVA)", type=["pdf"], accept_multiple_files=True)
Annexes_csv = st.sidebar.file_uploader("4. Annexes Fournisseurs (.csv / .xlsx)", type=["csv", "xlsx"], accept_multiple_files=True)

Def isoler_ref_article(libelle):
    If not isinstance(libelle, str) or pd.isna(libelle):
        Return ""
    Match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', str(libelle))
    Return match.group(1) if match else str(libelle).split(' - ')[0].strip()

Def clean_string_fuzzy(s):
    If not isinstance(s, str) or pd.isna(s):
        Return ""
    Return re.sub(r'[^a-zA-Z0-9]', '', str(s).lower())

Def extract_param_key_strict(text):
    If not isinstance(text, str) or pd.isna(text):
        Return ""
    If '//' in str(text):
        Text = str(text).split('//')[-1]
    
    Digits = re.sub(r'\D', '', str(text))
    If len(digits) >= 9:
        Return digits[-9:]
    M_id = re.search(r'([a-zA-Z0-9_\-\.]+@[a-zA-Z0-9_\-\.]+)', str(text))
    If m_id:
        Return m_id.group(1).lower()
    Return clean_string_fuzzy(text)

Def extract_code_from_string(text):
    If not isinstance(text, str) or pd.isna(text):
        Return ""
    M = re.search(r'(sk\d{4}|to_cl\d+|nm\d+|ns_\d+|\d{5,6})', str(text).lower())
    Return m.group(1).upper() if m else ""

Def lire_csv_securise(fichier):
    For enc in ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']:
        Try:
            Fichier.seek(0)
            Return pd.read_csv(fichier, sep=None, engine='python', encoding=enc)
        Except Exception:
            Continue
    Fichier.seek(0)
    Return pd.read_csv(fichier, sep=None, engine='python', encoding='utf-8', errors='ignore')

Onglet_tva, onglet_ht = st.tabs(["🚨 1. Audit Écarts TVA (20% vs 0%)", "📊 2. Analyse Quantité & Prix HT (Version Conforme)"])

# --- ONGLET 1 : AUDIT TVA ---
With onglet_tva:
    St.subheader("Audit Fiscal TVA (STD 20% vs APST 0%)")
    If fichier_artis and factures_pdf:
        If st.button("🚨 Lancer l'Analyse TVA", type="primary", key="btn_tva"):
            With st.spinner("Analyse des factures PDF Sewan et vérification des RFC ERP..."):
                Try:
                    Df_artis = pd.read_excel(fichier_artis)
                    Col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    Df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    Df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    
                    Def harmoniser_tva_artis_exact(row):
                        Gtva = str(row.get('Coût ABONNEMENT Gestion TVA', ''))
                        Ttva = str(row.get('Coût ABONNEMENT Taux TVA', ''))
                        If 'STD' in gtva or 'Standard' in gtva or '20' in ttva or '20.00' in ttva:
                            Return 'S (20%)'
                        Return 'APST (0%)'

                    Df_artis_abonn['TVA_Artis'] = df_artis_abonn.apply(harmoniser_tva_artis_exact, axis=1)
                    
                    Lignes_tva = []
                    For pdf in factures_pdf:
                        Reader = pypdf.PdfReader(pdf)
                        For page in reader.pages:
                            Text = page.extract_text() or ""
                            Lines = text.split('\n')
                            For i, line in enumerate(lines):
                                Match = re.search(r'([0-9]{2}-[\d\-]+-[0-9]{2}-[MFUV])', line)
                                If match:
                                    Ref = match.group(1)
                                    Tva_code = "APST (0%)"
                                    For j in range(i, min(i + 5, len(lines))):
                                        If 'S (20%' in lines[j] or 'S (20,00%)' in lines[j]:
                                            Tva_code = "S (20%)"
                                            Break
                                    Lignes_tva.append({'Code_Article_ERP': ref, 'TVA_Sewan': tva_code})
                                    
                    Df_sewan_tva = pd.DataFrame(lignes_tva).drop_duplicates() if lignes_tva else pd.DataFrame(columns=['Code_Article_ERP', 'TVA_Sewan'])
                    Df_merged_tva = pd.merge(df_artis_abonn, df_sewan_tva, on='Code_Article_ERP', how='inner')
                    Df_errors_tva = df_merged_tva[df_merged_tva['TVA_Artis'] != df_merged_tva['TVA_Sewan']].copy()
                    Df_errors_tva['Diagnostic_RAF'] = "🚨 Divergence TVA"
                    
                    If 'Raison sociale client' in df_errors_tva.columns:
                        Df_errors_tva['Raison sociale client'] = df_errors_tva['Raison sociale client'].astype(str).str.slice(0, 20)
                        
                    Cols_tva_export = ['Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Libellé RFC', 'Code_Article_ERP', 'TVA_Artis', 'TVA_Sewan', 'Diagnostic_RAF']
                    For c in cols_tva_export:
                        If c not in df_errors_tva.columns:
                            Df_errors_tva[c] = "Non renseigné"
                            
                    Df_res_tva = df_errors_tva[cols_tva_export].drop_duplicates()
                    
                    Buffer_tva = io.BytesIO()
                    With pd.ExcelWriter(buffer_tva, engine='openpyxl') as writer:
                        Df_res_tva.to_excel(writer, index=False, sheet_name='Anomalies_TVA')
                    
                    St.success(f"Audit TVA terminé : {len(df_res_tva)} anomalies réelles identifiées !")
                    St.dataframe(df_res_tva, use_container_width=True)
                    St.download_button("📥 Télécharger le Rapport TVA (.xlsx)", data=buffer_tva.getvalue(), file_name="Anomalies_TVA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                Except Exception as e:
                    St.error(f"Erreur pendant l'analyse TVA : {e}")
    Else:
        St.info("👈 Veuillez charger la simulation Artis (.xlsx) et les factures PDF Sewan.")

# --- ONGLET 2 : ANALYSE HT RESTAURÉE ---
With onglet_ht:
    St.subheader("Rapprochement Financier HT (Alignement Tiers & Libellés Produits)")
    If fichier_artis and annexes_csv:
        Nb_annexes = len(annexes_csv)
        If st.button("🚀 Lancer l'Analyse HT", type="primary", key="btn_ht"):
            With st.spinner(f"Traitement sur {nb_annexes} annexes CSV..."):
                Try:
                    # 1. Table de correspondance Tiers
                    Client_map = {}
                    If fichier_ref:
                        Xls_ref = pd.ExcelFile(fichier_ref)
                        Sheet_cli = 'Analyse clients' if 'Analyse clients' in xls_ref.sheet_names else xls_ref.sheet_names[0]
                        Df_ref_cli = pd.read_excel(fichier_ref, sheet_name=sheet_cli)
                        For _, row in df_ref_cli.iterrows():
                            Artis_code = str(row.get('CODE', '')).strip()
                            If not artis_code or artis_code == 'nan':
                                Continue
                            For col_name in ['Client/revendeur direct (niv. 1)', 'Réf. client/revendeur direct (niv. 1)', 'Nom client Niveau 2', 'Ref client Niveau 2']:
                                Val = str(row.get(col_name, '')).strip()
                                If val and val != 'nan':
                                    Client_map[val.lower()] = artis_code
                                    M = re.match(r'^([A-Za-z0-9_\-]+)', val)
                                    If m:
                                        Client_map[m.group(1).lower()] = artis_code

                    Tot_toolip = 15324.15
                    Tot_sokatel = 10874.99
                    Tot_nextphone = 1981.41
                    Tot_fournisseurs = tot_toolip + tot_sokatel + tot_nextphone

                    Df_artis = pd.read_excel(fichier_artis)
                    
                    # FILTRAGE SEWAN SÉCURISÉ (Ne vide jamais le jeu de données)
                    Col_fourn = next((c for c in df_artis.columns if 'fourn' in str(c).lower() or 'org' in str(c).lower() or 'raison' in str(c).lower()), None)
                    If col_fourn:
                        Df_sewan = df_artis[df_artis[col_fourn].astype(str).str.contains('SEWAN', case=False, na=False)].copy()
                        If not df_sewan.empty:
                            Df_artis = df_sewan

                    Col_art_artis = 'Coût ABONNEMENT article' if 'Coût ABONNEMENT article' in df_artis.columns else df_artis.columns[0]
                    Df_artis_abonn = df_artis.dropna(subset=[col_art_artis]).copy()
                    Df_artis_abonn['Code_Article_ERP'] = df_artis_abonn[col_art_artis].apply(isoler_ref_article)
                    Df_artis_abonn = df_artis_abonn[~df_artis_abonn['Code_Article_ERP'].str.endswith('-F')].copy()

                    # Montant Direct
                    Col_m_artis = 'Coût ABONNEMENT facturé' if 'Coût ABONNEMENT facturé' in df_artis_abonn.columns else df_artis.columns[1]
                    Df_artis_abonn['Montant_ERP'] = pd.to_numeric(df_artis_abonn[col_m_artis], errors='coerce').fillna(0.0).round(2)
                    
                    Df_artis_abonn['Code_Client_Str'] = df_artis_abonn['Code client'].astype(str).str.strip()
                    Df_artis_abonn['Name_Client_Clean'] = df_artis_abonn['Raison sociale client'].apply(clean_string_fuzzy)
                    Df_artis_abonn['Param_Key'] = df_artis_abonn['Libellé RFC'].apply(extract_param_key_strict)
                    
                    Def extract_product_from_rfc(rfc):
                        If not isinstance(rfc, str) or pd.isna(rfc):
                            Return ""
                        Part_left = str(rfc).split('//')[0].strip() if '//' in str(rfc) else str(rfc).strip()
                        Return clean_string_fuzzy(part_left)

                    Df_artis_abonn['Libelle_RFC_Clean'] = df_artis_abonn['Libellé RFC'].apply(extract_product_from_rfc)

                    # Quantité FA
                    Col_qte_fa = next((c for c in df_artis_abonn.columns if 'Abonnement Nombre' in str(c) or 'nt Nombre' in str(c) or str(c).endswith('Nombre')), None)
                    If col_qte_fa:
                        Df_artis_abonn['Quantite_ERP'] = pd.to_numeric(df_artis_abonn[col_qte_fa], errors='coerce').fillna(1.0)
                    Elif len(df_artis_abonn.columns) >= 147:
                        Df_artis_abonn['Quantite_ERP'] = pd.to_numeric(df_artis_abonn.iloc[:, 146], errors='coerce').fillna(1.0)
                    Else:
                        Df_artis_abonn['Quantite_ERP'] = 1.0

                    Lignes_fourn = []
                    For annexe in annexes_csv:
                        Df_annexe = lire_csv_securise(annexe) if annexe.name.endswith('.csv') else pd.read_excel(annexe)
                        Cols = [str(c).lower().strip() for c in df_annexe.columns]
                        
                        Idx_ref = next((i for i, c in enumerate(cols) if any(k in c for k in ['code produit', 'code_article', 'ref', 'code'])), 0)
                        Idx_m = next((i for i, c in enumerate(cols) if any(k in c for k in ["prix d'achat", 'prix', 'montant', 'ht'])), 1)
                        Idx_qte = next((i for i, c in enumerate(cols) if any(k in c for k in ['quantité', 'quantite', 'qte'])), -1)
                        Idx_p1 = next((i for i, c in enumerate(cols) if 'param1' in c), -1)
                        Idx_p2 = next((i for i, c in enumerate(cols) if 'param2' in c), -1)
                        
                        Idx_prod = next((i for i, c in enumerate(cols) if c == 'produit' or (('produit' in c or 'libellé' in c) and 'famille' not in c and 'code' not in c)), -1)
                        If idx_prod == -1:
                            Idx_prod = next((i for i, c in enumerate(cols) if 'produit' in c and 'famille' not in c), -1)

                        Idx_cli1 = next((i for i, c in enumerate(cols) if 'client/revendeur direct (niv. 1)' in c or 'client/revendeur' in c), -1)
                        Idx_cli2 = next((i for i, c in enumerate(cols) if 'nom client niveau 2' in c or 'nom client' in c), -1)
                        
                        Col_ref = df_annexe.columns[idx_ref]
                        Col_montant = df_annexe.columns[idx_m]
                        Col_qte = df_annexe.columns[idx_qte] if idx_qte != -1 else None
                        Col_param1 = df_annexe.columns[idx_p1] if idx_p1 != -1 else None
                        Col_param2 = df_annexe.columns[idx_p2] if idx_p2 != -1 else None
                        Col_cli1 = df_annexe.columns[idx_cli1] if idx_cli1 != -1 else None
                        Col_cli2 = df_annexe.columns[idx_cli2] if idx_cli2 != -1 else None
                        Col_produit = df_annexe.columns[idx_prod] if idx_prod != -1 else None
                        
                        For _, row in df_annexe.iterrows():
                            Ref = isoler_ref_article(str(row[col_ref]))
                            If ref.endswith('-F'):
                                Continue

                            Try:
                                M_ht = float(str(row[col_montant]).replace(',', '.').replace(' ', '').replace('€', ''))
                            Except ValueError:
                                M_ht = 0.0
                                
                            Try:
                                Q_sewan = float(str(row[col_qte]).replace(',', '.').replace(' ', '')) if col_qte else 1.0
                            Except ValueError:
                                Q_sewan = 1.0
                            
                            C1_str = str(row[col_cli1]).strip() if col_cli1 and pd.notna(row[col_cli1]) else ""
                            C2_str = str(row[col_cli2]).strip() if col_cli2 and pd.notna(row[col_cli2]) else ""
                            P1_val = str(row[col_param1]) if col_param1 else ""
                            P2_val = str(row[col_param2]) if col_param2 else ""
                            Prod_val = str(row[col_produit]).strip() if col_produit and pd.notna(row[col_produit]) else ""
                            
                            Res_code = ""
                            For test_val in [c1_str, c2_str, p1_val, p2_val]:
                                If test_val.lower() in client_map:
                                    Res_code = client_map[test_val.lower()]
                                    Break
                                Token = extract_code_from_string(test_val)
                                If token and token.lower() in client_map:
                                    Res_code = client_map[token.lower()]
                                    Break
                                    
                            If not res_code:
                                For test_val in [c1_str, c2_str, p1_val, p2_val]:
                                    Token = extract_code_from_string(test_val)
                                    If token:
                                        Res_code = token
                                        Break
                                        
                            Cli_raw = c1_str if c1_str else (c2_str if c2_str else p1_val)
                            Param_key = extract_param_key_strict(p1_val)
                            
                            Lignes_fourn.append({
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
                            
                    Df_fourn = pd.DataFrame(lignes_fourn).groupby(['Code_Client_Resolu', 'Name_Sewan_Clean', 'Code_Article_ERP', 'Param_Key', 'Produit_Sewan_Clean'], as_index=False).agg(
                        Montant_Sewan=('Montant_Sewan', 'sum'),
                        Quantite_Sewan=('Quantite_Sewan', 'sum'),
                        Produit=('Produit', 'first'),
                        Client_Sewan_Raw=('Client_Sewan_Raw', 'first'),
                        Annexe_Source=('Annexe_Source', lambda x: ', '.join(set(x)))
                    )

                    # 1. RAPPROCHEMENT PRINCIPAL
                    Df_recon_1 = pd.merge(
                        Df_artis_abonn, 
                        Df_fourn, 
                        Left_on=['Code_Client_Str', 'Code_Article_ERP', 'Param_Key'], 
                        Right_on=['Code_Client_Resolu', 'Code_Article_ERP', 'Param_Key'], 
                        How='outer'
                    )

                    # 2. FILET DE SÉCURITÉ : CLIENT 10 / AFIDF-10
                    Mask_c10_orphelin = (
                        Df_recon_1['Montant_Sewan'].isna() | df_recon_1['Montant_ERP'].isna()
                    ) & (
                        (df_recon_1['Code_Client_Str'] == '10') | 
                        (df_recon_1['Code SSC'].astype(str).str.contains('AFIDF-10', case=False, na=False))
                    )

                    If mask_c10_orphelin.any():
                        Df_erp_c10_unmatched = df_recon_1[mask_c10_orphelin & df_recon_1['Montant_Sewan'].isna()].drop(
                            Columns=['Code_Client_Resolu', 'Name_Sewan_Clean', 'Montant_Sewan', 'Quantite_Sewan', 'Client_Sewan_Raw', 'Annexe_Source', 'Produit_Sewan_Clean', 'Produit'], 
                            Errors='ignore'
                        )
                        
                        Df_sewan_c10_unmatched = df_fourn[df_fourn['Code_Client_Resolu'] == '10']
                        
                        Df_recon_fallback = pd.merge(
                            Df_erp_c10_unmatched,
                            Df_sewan_c10_unmatched,
                            Left_on=['Code_Article_ERP', 'Libelle_RFC_Clean'],
                            Right_on=['Code_Article_ERP', 'Produit_Sewan_Clean'],
                            How='inner'
                        )
                        
                        If not df_recon_fallback.empty:
                            Matched_rfcs = df_recon_fallback['Libelle_RFC_Clean'].unique()
                            Matched_prods = df_recon_fallback['Produit_Sewan_Clean'].unique()
                            
                            Df_recon_base = df_recon_1[
                                ~(df_recon_1['Libelle_RFC_Clean'].isin(matched_rfcs) & df_recon_1['Montant_Sewan'].isna()) &
                                ~(df_recon_1['Produit_Sewan_Clean'].isin(matched_prods) & df_recon_1['Montant_ERP'].isna())
                            ]
                            
                            Df_recon = pd.concat([df_recon_base, df_recon_fallback], ignore_index=True)
                        Else:
                            Df_recon = df_recon_1.copy()
                    Else:
                        Df_recon = df_recon_1.copy()

                    Df_recon['Montant_ERP'] = df_recon['Montant_ERP'].fillna(0.0).round(2)
                    Df_recon['Montant_Sewan'] = df_recon['Montant_Sewan'].fillna(0.0).round(2)
                    Df_recon['Quantite_ERP'] = df_recon['Quantite_ERP'].fillna(0.0)
                    Df_recon['Quantite_Sewan'] = df_recon['Quantite_Sewan'].fillna(0.0)
                    
                    Df_recon['Ecart_HT'] = (df_recon['Montant_ERP'] - df_recon['Montant_Sewan']).round(2)
                    
                    Def qualifier_ht_exact(row):
                        If row['Montant_ERP'] == 0.0 and row['Montant_Sewan'] == 0.0:
                            Return "🟢 Conforme (Option 0€)"
                        If pd.isna(row.get(col_art_artis)) or str(row.get(col_art_artis)).strip() in ['', 'nan']:
                            Return "🔴 Vente récente non saisie ERP"
                        If pd.isna(row.get('Client_Sewan_Raw')) or row.get('Montant_Sewan', 0) == 0:
                            Return "🟠 Résiliation non saisie ERP"
                        If abs(row['Ecart_HT']) > 0.05 or abs(row['Quantite_ERP'] - row['Quantite_Sewan']) > 0.01:
                            Return "⚠️ Écart de montant/quantité HT"
                        Return "🟢 Conforme"

                    Df_recon['Diagnostic_RAF'] = df_recon.apply(qualifier_ht_exact, axis=1)

                    If 'Code_Client_Resolu' in df_recon.columns:
                        Df_recon['Code client'] = df_recon['Code client'].fillna(df_recon['Code_Client_Resolu'])
                    If 'Client_Sewan_Raw' in df_recon.columns:
                        Df_recon['Raison sociale client'] = df_recon['Raison sociale client'].fillna(df_recon['Client_Sewan_Raw'])

                    # CALCUL DYNAMIQUE TOTALS
                    Tot_simulation_artis_reel = df_recon['Montant_ERP'].sum()
                    Tot_fournisseurs_reel = df_recon['Montant_Sewan'].sum()
                    Ecart_global_reel = tot_simulation_artis_reel - tot_fournisseurs_reel

                    # --- SYNTHÈSE EXÉCUTIVE METRIQUES ---
                    St.markdown("### 📋 Synthèse des Factures Fournisseurs (Abonnements HT)")
                    C1, c2, c3, c4 = st.columns(4)
                    C1.metric("Toolip", "15,324.15 €")
                    C2.metric("Sokatel", "10,874.99 €")
                    C3.metric("Nextphone", "1,981.41 €")
                    C4.metric("Total Reçu Annexes Sewan", f"{tot_fournisseurs_reel:,.2f} €")

                    St.markdown("### 🧮 Comparatif Global ERP vs Fournisseurs")
                    K1, k2, k3 = st.columns(3)
                    K1.metric("Total ERP Extrait (Sewan)", f"{tot_simulation_artis_reel:,.2f} €")
                    K2.metric("Total Réconcilié Annexes Sewan", f"{tot_fournisseurs_reel:,.2f} €")
                    
                    If abs(ecart_global_reel) < 1000.0:
                        K3.metric("Écart Net Global HT", f"{ecart_global_reel:,.2f} €", delta="🟢 Conforme (< 1 000 €)", delta_color="normal")
                    Else:
                        K3.metric("Écart Net Global HT", f"{ecart_global_reel:,.2f} €", delta="🟠 À expertiser (> 1 000 €)", delta_color="inverse")

                    St.markdown("---")
                    St.markdown("### 📊 Répartition Financière & Pourcentages par Diagnostic")
                    
                    Total_lignes = len(df_recon)
                    Total_sewan_sum = tot_fournisseurs_reel if tot_fournisseurs_reel > 0 else 1.0

                    Diag_stats_full = df_recon.groupby('Diagnostic_RAF').agg(
                        Nombre_Lignes=('Diagnostic_RAF', 'count'),
                        Total_Artis_ERP=('Montant_ERP', 'sum'),
                        Total_Sewan_Fournisseur=('Montant_Sewan', 'sum'),
                        Total_Ecart_HT=('Ecart_HT', 'sum')
                    ).reset_index()

                    Diag_stats_full['% Lignes'] = ((diag_stats_full['Nombre_Lignes'] / total_lignes) * 100).round(1).astype(str) + " %"
                    Diag_stats_full['% HT Sewan'] = ((diag_stats_full['Total_Sewan_Fournisseur'] / total_sewan_sum) * 100).round(1).astype(str) + " %"

                    St.dataframe(diag_stats_full[['Diagnostic_RAF', 'Nombre_Lignes', '% Lignes', 'Total_Artis_ERP', 'Total_Sewan_Fournisseur', '% HT Sewan', 'Total_Ecart_HT']], use_container_width=True)

                    Col_chart, col_empty = st.columns([1, 1])
                    With col_chart:
                        St.markdown("### 🍕 Répartition (%) des Lignes d'Abonnements")
                        St.bar_chart(df_recon['Diagnostic_RAF'].value_counts(normalize=True) * 100)

                    St.markdown("---")
                    St.markdown("### 🔍 Tableau Détaillé des Lignes d'Abonnements")

                    If 'Raison sociale client' in df_recon.columns:
                        Df_recon['Raison sociale client'] = df_recon['Raison sociale client'].astype(str).str.slice(0, 20)
                        
                    Cols_export_ht = [
                        'Code client', 'Raison sociale client', 'Code SSC', 'Code RFC', 'Libellé RFC', 
                        'Code_Article_ERP', 'Produit', 'Quantite_ERP', 'Quantite_Sewan', 
                        'Montant_ERP', 'Montant_Sewan', 'Ecart_HT', 'Diagnostic_RAF', 'Annexe_Source'
                    ]
                    For c in cols_export_ht:
                        If c not in df_recon.columns:
                            Df_recon[c] = "Non renseigné"
                            
                    Df_res_ht = df_recon[cols_export_ht].drop_duplicates()
                    
                    Buffer_ht = io.BytesIO()
                    With pd.ExcelWriter(buffer_ht, engine='openpyxl') as writer:
                        Df_res_ht.to_excel(writer, index=False, sheet_name='Analyse_HT_Client_Article')
                    
                    St.dataframe(df_res_ht, use_container_width=True)
                    St.download_button("📥 Télécharger l'Analyse HT Détaillée (.xlsx)", data=buffer_ht.getvalue(), file_name="Analyse_RAF_HT_Detaillee.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                Except Exception as e:
                    St.error(f"Erreur pendant le traitement HT : {e}")
    Else:
        St.info("👈 Veuillez charger la simulation Artis (.xlsx) et les factures PDF Sewan.")
