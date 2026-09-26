import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Tableau de Bord", page_icon="📊", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

st.title("📊 Tableau de Bord – Achats Karité")
st.markdown("---")

FICHIER_TRACKER = "SHEA_PURCHASE TRACKER.xlsx"
FICHIER_BONS = "donnees_bons.xlsx"
FICHIER_DECH = "dechargements.xlsx"

@st.cache_data(ttl=30)
def charger_tout():
    data = {}
    # Nouveaux bons (app)
    if os.path.exists(FICHIER_BONS):
        data["nouveaux"] = pd.read_excel(FICHIER_BONS)
    else:
        data["nouveaux"] = pd.DataFrame()
    
    # Déchargements saisis dans l'app
    if os.path.exists(FICHIER_DECH):
        data["dechargements"] = pd.read_excel(FICHIER_DECH)
    else:
        data["dechargements"] = pd.DataFrame()
    
    # Tracker
    if os.path.exists(FICHIER_TRACKER):
        try:
            data["bons_tracker"] = pd.read_excel(FICHIER_TRACKER, sheet_name="Bons de Commande", header=3)
            data["bons_tracker"].columns = [str(c).strip() for c in data["bons_tracker"].columns]
        except Exception:
            data["bons_tracker"] = pd.DataFrame()
        try:
            data["cmd_decharge"] = pd.read_excel(FICHIER_TRACKER, sheet_name="Commandes Dechargees", header=3)
            data["cmd_decharge"].columns = [str(c).strip() for c in data["cmd_decharge"].columns]
        except Exception:
            data["cmd_decharge"] = pd.DataFrame()
        try:
            data["releve"] = pd.read_excel(FICHIER_TRACKER, sheet_name="Releve Comptes Fournisseurs", header=3)
            data["releve"].columns = [str(c).strip() for c in data["releve"].columns]
        except Exception:
            data["releve"] = pd.DataFrame()
    else:
        data["bons_tracker"] = pd.DataFrame()
        data["cmd_decharge"] = pd.DataFrame()
        data["releve"] = pd.DataFrame()
    
    return data

data = charger_tout()
df_nouveaux = data["nouveaux"]
df_dech_app = data["dechargements"]
df_bons_tr = data["bons_tracker"]
df_cmd = data["cmd_decharge"]
df_releve = data["releve"]

# ---------- KPIs ----------
st.subheader("Indicateurs clés")

nb_nouveaux = len(df_nouveaux)
nb_tracker = len(df_bons_tr) if not df_bons_tr.empty else 0
nb_dech_app = len(df_dech_app)
nb_dech_tracker = len(df_cmd) if not df_cmd.empty else 0

poids_app = pd.to_numeric(df_dech_app.get("Poids_Net_KG", 0), errors="coerce").sum() if not df_dech_app.empty else 0
montant_app = pd.to_numeric(df_dech_app.get("Montant_Du", 0), errors="coerce").sum() if not df_dech_app.empty else 0
solde_app = pd.to_numeric(df_dech_app.get("Solde_Restant", 0), errors="coerce").sum() if not df_dech_app.empty else 0

col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    st.metric("Bons (App)", nb_nouveaux)
with col2:
    st.metric("Bons (Tracker)", nb_tracker)
with col3:
    st.metric("Déchargements (App)", nb_dech_app)
with col4:
    st.metric("Déchargements (Tracker)", nb_dech_tracker)
with col5:
    st.metric("Poids net App (KG)", f"{poids_app:,.0f}")
with col6:
    st.metric("Solde restant App (FCFA)", f"{solde_app:,.0f}")

st.markdown("---")

# ---------- Filtres ----------
st.subheader("Filtres")
col_f1, col_f2 = st.columns(2)
with col_f1:
    filtre_fourn = st.text_input("Filtrer par fournisseur", "")
with col_f2:
    filtre_statut = st.selectbox("Statut paiement (déchargements app)", ["Tous", "NON PAYE", "PARTIEL", "PAYE"])

# ---------- Nouveaux bons ----------
st.subheader("📄 Nouveaux bons de commande (format XXXX-BC-SHEA/NAP-2026)")
df_aff = df_nouveaux.copy()
if filtre_fourn and not df_aff.empty and "Nom_Fournisseur" in df_aff.columns:
    df_aff = df_aff[df_aff["Nom_Fournisseur"].astype(str).str.contains(filtre_fourn, case=False, na=False)]
if df_aff.empty:
    st.info("Aucun nouveau bon.")
else:
    st.dataframe(df_aff, width="stretch", height=250)

# ---------- Déchargements app ----------
st.subheader("🚛 Déchargements saisis dans l'application")
df_d = df_dech_app.copy()
if not df_d.empty:
    if filtre_fourn and "Fournisseur" in df_d.columns:
        df_d = df_d[df_d["Fournisseur"].astype(str).str.contains(filtre_fourn, case=False, na=False)]
    if filtre_statut != "Tous" and "Statut_Paiement" in df_d.columns:
        df_d = df_d[df_d["Statut_Paiement"] == filtre_statut]
    st.dataframe(df_d, width="stretch", height=250)
else:
    st.info("Aucun déchargement saisi dans l'app.")

# ---------- Tracker Bons ----------
st.subheader("📋 Bons du Tracker")
if not df_bons_tr.empty:
    cols = [c for c in ["Date", "N° BC", "Code Frs", "Fournisseur", "N° Camion", "Order QTY(MT)", "Verdict Labo", "Statut Livraison"] if c in df_bons_tr.columns]
    df_bt = df_bons_tr[cols] if cols else df_bons_tr
    if filtre_fourn and "Fournisseur" in df_bt.columns:
        df_bt = df_bt[df_bt["Fournisseur"].astype(str).str.contains(filtre_fourn, case=False, na=False)]
    st.dataframe(df_bt.head(150), width="stretch", height=300)
    st.caption(f"Affichage limité à 150 lignes (total : {len(df_bons_tr)})")
else:
    st.warning("Tracker Bons de Commande non chargé.")

# ---------- Relevé ----------
st.subheader("💰 Relevé comptes fournisseurs (Tracker)")
if not df_releve.empty:
    st.dataframe(df_releve, width="stretch", height=250)
else:
    st.info("Relevé non disponible.")

if st.button("🔄 Actualiser"):
    st.cache_data.clear()
    st.rerun()