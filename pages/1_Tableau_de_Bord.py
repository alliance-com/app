import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Dashboard", page_icon="📊", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")

# ==================== FICHIERS ====================
FICHIER_BONS = "donnees_bons.xlsx"
FICHIER_DECH = "dechargements.xlsx"
FICHIER_FRS = "fournisseurs.xlsx"
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FICHIER_MOUVEMENTS = "comptes_fournisseurs.xlsx"
# ==================================================

def fmt_int(n):
    """Nombre entier lisible : 339102300 → 339 102 300"""
    try:
        return f"{int(round(float(n))):,}".replace(",", " ")
    except Exception:
        return "0"

def fmt_fcfa(n):
    return f"{fmt_int(n)} FCFA"

def _num(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

def _txt(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s

def safe_read(path, sheet=None):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        if sheet:
            return pd.read_excel(path, sheet_name=sheet)
        return pd.read_excel(path)
    except Exception:
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()

# ---------- Chargement ----------
df_bons = safe_read(FICHIER_BONS)
df_dech = safe_read(FICHIER_DECH)
df_frs = safe_read(FICHIER_FRS)
df_mvt = safe_read(FICHIER_MOUVEMENTS, "Mouvements")
if df_mvt.empty:
    df_mvt = safe_read(FICHIER_MOUVEMENTS)

# ---------- Calculs ----------
nb_bons = len(df_bons) if not df_bons.empty else 0
nb_frs_actifs = 0
if not df_frs.empty and "Actif" in df_frs.columns:
    nb_frs_actifs = int((df_frs["Actif"].astype(str).str.upper() == "OUI").sum())
elif not df_frs.empty:
    nb_frs_actifs = len(df_frs)

nb_dech = len(df_dech) if not df_dech.empty else 0

# Poids & montants dus (déchargements)
poids_total = 0.0
montant_du = 0.0
montant_paye_dech = 0.0
if not df_dech.empty:
    for col in ("Poids_Net_KG", "Poids_Net", "Poids"):
        if col in df_dech.columns:
            poids_total = df_dech[col].apply(_num).sum()
            break
    for col in ("Montant_Du", "Montant_Total", "Montant"):
        if col in df_dech.columns:
            montant_du = df_dech[col].apply(_num).sum()
            break
    for col in ("Montant_Paye", "Montant_Payé"):
        if col in df_dech.columns:
            montant_paye_dech = df_dech[col].apply(_num).sum()
            break

# Crédits mouvements
montant_credit = 0.0
if not df_mvt.empty:
    for col in df_mvt.columns:
        if str(col).lower() in ("credit", "crédit", "montant_credit"):
            montant_credit = df_mvt[col].apply(_num).sum()
            break

montant_paye = max(montant_paye_dech, montant_credit) if montant_credit else montant_paye_dech
if montant_credit > montant_paye_dech:
    montant_paye = montant_credit

solde = montant_du - montant_paye
pct = (montant_paye / montant_du * 100) if montant_du else 0.0

# Fournisseurs distincts dans les bons
nb_frs_bons = 0
if not df_bons.empty and "Nom_Fournisseur" in df_bons.columns:
    nb_frs_bons = df_bons["Nom_Fournisseur"].dropna().astype(str).nunique()

# ==================== UI ====================
st.title("📊 Tableau de bord – NAP SARL")
st.caption(f"Connecté : **{name}** · Mis à jour : {datetime.now().strftime('%d/%m/%Y %H:%M')}")
st.markdown("---")

# ----- BLOC 1 : Activité -----
st.subheader("Activité")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown("**Bons de commande**")
    st.markdown(f"<p style='font-size:2rem;font-weight:700;margin:0'>{fmt_int(nb_bons)}</p>", unsafe_allow_html=True)
with c2:
    st.markdown("**Fournisseurs actifs**")
    st.markdown(f"<p style='font-size:2rem;font-weight:700;margin:0'>{fmt_int(nb_frs_actifs)}</p>", unsafe_allow_html=True)
with c3:
    st.markdown("**Déchargements**")
    st.markdown(f"<p style='font-size:2rem;font-weight:700;margin:0'>{fmt_int(nb_dech)}</p>", unsafe_allow_html=True)
with c4:
    st.markdown("**Fournisseurs (sur BC)**")
    st.markdown(f"<p style='font-size:2rem;font-weight:700;margin:0'>{fmt_int(nb_frs_bons)}</p>", unsafe_allow_html=True)

st.markdown("---")

# ----- BLOC 2 : Financier (texte complet, pas de métrique tronquée) -----
st.subheader("Situation financière")

col_a, col_b = st.columns(2)
with col_a:
    st.markdown(
        f"""
        <div style="background:#f0f4f8;padding:1.2rem 1.5rem;border-radius:10px;border-left:5px solid #1E5AA0">
            <div style="color:#555;font-size:0.9rem">Poids net livré</div>
            <div style="font-size:1.6rem;font-weight:700;color:#0F2850">{fmt_int(poids_total)} kg</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.write("")
    st.markdown(
        f"""
        <div style="background:#e8f5e9;padding:1.2rem 1.5rem;border-radius:10px;border-left:5px solid #2e7d32">
            <div style="color:#555;font-size:0.9rem">Montant total dû</div>
            <div style="font-size:1.6rem;font-weight:700;color:#1b5e20">{fmt_fcfa(montant_du)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_b:
    st.markdown(
        f"""
        <div style="background:#e3f2fd;padding:1.2rem 1.5rem;border-radius:10px;border-left:5px solid #1565c0">
            <div style="color:#555;font-size:0.9rem">Montant payé</div>
            <div style="font-size:1.6rem;font-weight:700;color:#0d47a1">{fmt_fcfa(montant_paye)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.write("")
    couleur_solde = "#b71c1c" if solde > 0 else "#1b5e20"
    st.markdown(
        f"""
        <div style="background:#fff3e0;padding:1.2rem 1.5rem;border-radius:10px;border-left:5px solid #ef6c00">
            <div style="color:#555;font-size:0.9rem">Solde restant (à payer)</div>
            <div style="font-size:1.6rem;font-weight:700;color:{couleur_solde}">{fmt_fcfa(solde)}</div>
            <div style="color:#666;font-size:0.95rem;margin-top:0.3rem">Taux payé : <b>{pct:.0f} %</b></div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# ----- BLOC 3 : Derniers bons -----
st.subheader("Derniers bons de commande")
if df_bons.empty:
    st.info("Aucun bon enregistré.")
else:
    cols_aff = [c for c in [
        "Reference", "Date_Emission", "Code_Fournisseur", "Nom_Fournisseur",
        "Quantite", "Numero_Tracteur", "Utilisateur"
    ] if c in df_bons.columns]
    derniers = df_bons[cols_aff].tail(10).iloc[::-1]
    st.dataframe(derniers, width="stretch", height=280)

# ----- BLOC 4 : Derniers déchargements -----
st.subheader("Derniers déchargements")
if df_dech.empty:
    st.info("Aucun déchargement enregistré.")
else:
    cols_d = [c for c in df_dech.columns if c in [
        "N_BC", "Reference", "N_Camion", "Numero_Camion", "Fournisseur",
        "Poids_Net_KG", "Montant_Du", "Montant_Paye", "Solde_Restant", "Statut_Paiement"
    ] or True]
    # Afficher les colonnes utiles si présentes
    prefer = ["N_BC", "Reference", "Fournisseur", "N_Camion", "Numero_Camion",
              "Poids_Net_KG", "Montant_Du", "Montant_Paye", "Solde_Restant", "Statut_Paiement"]
    cols_d = [c for c in prefer if c in df_dech.columns]
    if not cols_d:
        cols_d = list(df_dech.columns)[:8]
    aff_d = df_dech[cols_d].tail(10).iloc[::-1].copy()
    # Format nombres en texte complet (évite 1.2e+08 dans le tableau)
    for col in aff_d.columns:
        if any(x in col.lower() for x in ("montant", "poids", "solde", "kg")):
            aff_d[col] = aff_d[col].apply(lambda x: fmt_int(_num(x)) if _num(x) or x == 0 else _txt(x))
    st.dataframe(aff_d, width="stretch", height=280)

st.markdown("---")
st.caption("NAP SARL – Tableau de bord · Chiffres affichés en entier (sans notation scientifique)")