import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
from openpyxl import Workbook

st.set_page_config(page_title="Fournisseurs", page_icon="👤", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

# ==================== CONFIG ====================
FICHIER_FRS = "fournisseurs.xlsx"
COLONNES = [
    "Code", "Nom_Fournisseur", "Telephone", "Commune", "PU_FCFA_KG",
    "Date_Creation", "Cree_Par", "Actif", "Commentaire"
]
# ================================================

def init_fichier():
    if os.path.exists(FICHIER_FRS):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Fournisseurs"
    ws.append(COLONNES)
    wb.save(FICHIER_FRS)

def charger_fournisseurs():
    init_fichier()
    df = pd.read_excel(FICHIER_FRS)
    for col in COLONNES:
        if col not in df.columns:
            if col == "PU_FCFA_KG":
                df[col] = 0.0
            else:
                df[col] = ""
    # Ordre des colonnes
    autres = [c for c in df.columns if c not in COLONNES]
    return df[COLONNES + autres]

def sauvegarder_df(df):
    df.to_excel(FICHIER_FRS, sheet_name="Fournisseurs", index=False)

def prochain_code(df):
    max_n = 0
    if not df.empty and "Code" in df.columns:
        for c in df["Code"].dropna().astype(str):
            m = re.search(r"FRS(\d+)", c.upper())
            if m:
                max_n = max(max_n, int(m.group(1)))
    return f"FRS{max_n + 1:05d}"

def _clean(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s

def _pu(val):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return 0.0
        return float(val)
    except Exception:
        return 0.0

def ajouter_fournisseur(nom, telephone, commune, pu, commentaire):
    df = charger_fournisseurs()
    code = prochain_code(df)
    noms = df["Nom_Fournisseur"].dropna().astype(str).str.strip().str.upper()
    if nom.strip().upper() in noms.values:
        return False, "", "Un fournisseur avec ce nom existe déjà."

    nouvelle = {
        "Code": code,
        "Nom_Fournisseur": nom.strip().upper(),
        "Telephone": telephone.strip(),
        "Commune": commune.strip(),
        "PU_FCFA_KG": float(pu) if pu else 0.0,
        "Date_Creation": datetime.now().strftime("%d/%m/%Y"),
        "Cree_Par": name,
        "Actif": "Oui",
        "Commentaire": commentaire.strip()
    }
    df = pd.concat([df, pd.DataFrame([nouvelle])], ignore_index=True)
    sauvegarder_df(df)
    return True, code, f"Fournisseur **{code}** — {nom.strip().upper()} créé (PU : {pu:.0f} FCFA/kg)."

def mettre_a_jour(code, nom, telephone, commune, pu, commentaire, actif):
    df = charger_fournisseurs()
    mask = df["Code"].astype(str).str.upper() == str(code).upper()
    if not mask.any():
        return False, "Code introuvable."
    idx = df.index[mask][0]
    df.at[idx, "Nom_Fournisseur"] = nom.strip().upper()
    df.at[idx, "Telephone"] = telephone.strip()
    df.at[idx, "Commune"] = commune.strip()
    df.at[idx, "PU_FCFA_KG"] = float(pu) if pu else 0.0
    df.at[idx, "Commentaire"] = commentaire.strip()
    df.at[idx, "Actif"] = "Oui" if actif else "Non"
    sauvegarder_df(df)
    return True, f"**{code}** mis à jour (PU : {float(pu):.0f} FCFA/kg)."

def desactiver(code):
    df = charger_fournisseurs()
    mask = df["Code"].astype(str).str.upper() == str(code).upper()
    if not mask.any():
        return False, "Code introuvable."
    df.loc[mask, "Actif"] = "Non"
    sauvegarder_df(df)
    return True, f"**{code}** désactivé (invisible dans les bons)."

def reactiver(code):
    df = charger_fournisseurs()
    mask = df["Code"].astype(str).str.upper() == str(code).upper()
    if not mask.any():
        return False, "Code introuvable."
    df.loc[mask, "Actif"] = "Oui"
    sauvegarder_df(df)
    return True, f"**{code}** réactivé."

def supprimer_definitif(code):
    df = charger_fournisseurs()
    mask = df["Code"].astype(str).str.upper() == str(code).upper()
    if not mask.any():
        return False, "Code introuvable."
    df = df[~mask]
    sauvegarder_df(df)
    return True, f"**{code}** supprimé définitivement."

# ==================== UI ====================
st.title("👤 Fournisseurs – NAP SARL")
st.caption("Code FRS · Téléphone · Commune · **PU (FCFA/kg)** · Actif / Inactif")
st.markdown("---")

df = charger_fournisseurs()
actifs = df[df["Actif"].astype(str).str.upper() == "OUI"] if not df.empty else df
inactifs = df[df["Actif"].astype(str).str.upper() != "OUI"] if not df.empty else df

c1, c2, c3 = st.columns(3)
c1.metric("Total fiches", len(df))
c2.metric("Actifs", len(actifs))
c3.metric("Inactifs", len(inactifs))

st.markdown("---")

onglet1, onglet2, onglet3, onglet4 = st.tabs([
    "➕ Ajouter",
    "✏️ Mettre à jour",
    "🗑️ Désactiver / Supprimer",
    "📋 Liste"
])

# ---------- AJOUTER ----------
with onglet1:
    st.subheader("Nouveau fournisseur")
    st.info(f"Prochain code automatique : **{prochain_code(df)}**")

    if st.session_state.get("frs_msg_ok"):
        st.success(st.session_state.pop("frs_msg_ok"))

    with st.form("form_ajout_frs", clear_on_submit=True):
        nom = st.text_input("Nom du fournisseur *", placeholder="Ex: VLAVONOU PIERRE")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            telephone = st.text_input("Téléphone", placeholder="97 00 00 00")
        with col_b:
            commune = st.text_input("Commune", placeholder="Parakou")
        with col_c:
            pu = st.number_input("PU (FCFA/kg) *", min_value=0.0, value=0.0, step=1.0,
                                 help="Prix unitaire pour calcul : Montant = Poids × PU")
        commentaire = st.text_input("Commentaire")

        if st.form_submit_button("💾 Enregistrer le fournisseur", type="primary"):
            if not nom.strip():
                st.error("Le nom est obligatoire.")
            elif pu <= 0:
                st.warning("PU à 0 : le montant des déchargements devra être saisi manuellement.")
                ok, code, msg = ajouter_fournisseur(nom, telephone, commune, pu, commentaire)
                if ok:
                    st.session_state["frs_msg_ok"] = f"✅ {msg}"
                    st.rerun()
                else:
                    st.error(msg)
            else:
                ok, code, msg = ajouter_fournisseur(nom, telephone, commune, pu, commentaire)
                if ok:
                    st.session_state["frs_msg_ok"] = f"✅ {msg}"
                    st.rerun()
                else:
                    st.error(msg)

# ---------- METTRE À JOUR ----------
with onglet2:
    st.subheader("Modifier un fournisseur")
    if df.empty:
        st.info("Aucun fournisseur.")
    else:
        options = [
            f"{r['Code']} — {r['Nom_Fournisseur']} (PU {_pu(r.get('PU_FCFA_KG')):.0f})"
            for _, r in df.iterrows()
        ]
        choix = st.selectbox("Fournisseur à modifier", options, key="upd_select")
        code_sel = choix.split("—")[0].strip()
        row = df[df["Code"].astype(str).str.upper() == code_sel.upper()].iloc[0]

        with st.form("form_maj_frs"):
            nom_u = st.text_input("Nom *", value=_clean(row.get("Nom_Fournisseur")))
            col1, col2, col3 = st.columns(3)
            with col1:
                tel_u = st.text_input("Téléphone", value=_clean(row.get("Telephone")))
            with col2:
                com_u = st.text_input("Commune", value=_clean(row.get("Commune")))
            with col3:
                pu_u = st.number_input(
                    "PU (FCFA/kg) *",
                    min_value=0.0,
                    value=_pu(row.get("PU_FCFA_KG")),
                    step=1.0
                )
            com_txt = st.text_input("Commentaire", value=_clean(row.get("Commentaire")))
            actif_u = st.checkbox(
                "Actif (visible dans les bons et déchargements)",
                value=str(row.get("Actif", "Oui")).upper() == "OUI"
            )

            if st.form_submit_button("💾 Enregistrer les modifications", type="primary"):
                if not nom_u.strip():
                    st.error("Le nom est obligatoire.")
                else:
                    ok, msg = mettre_a_jour(code_sel, nom_u, tel_u, com_u, pu_u, com_txt, actif_u)
                    if ok:
                        st.success(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(msg)

        st.caption(
            "Le **PU**, le **téléphone** et la **commune** sont repris automatiquement "
            "dans le bon de commande et le déchargement."
        )

# ---------- DESACTIVER / SUPPRIMER ----------
with onglet3:
    st.subheader("Désactiver ou supprimer")
    st.warning(
        "Préférez **désactiver** (historique conservé). "
        "La suppression définitive est réservée à l'admin."
    )

    if df.empty:
        st.info("Aucun fournisseur.")
    else:
        options = [
            f"{r['Code']} — {r['Nom_Fournisseur']} "
            f"({'Actif' if str(r.get('Actif', '')).upper() == 'OUI' else 'Inactif'})"
            for _, r in df.iterrows()
        ]
        choix = st.selectbox("Fournisseur", options, key="del_select")
        code_sel = choix.split("—")[0].strip()

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔴 Désactiver", width="stretch"):
                ok, msg = desactiver(code_sel)
                st.success(msg) if ok else st.error(msg)
                st.rerun()
        with col2:
            if st.button("🟢 Réactiver", width="stretch"):
                ok, msg = reactiver(code_sel)
                st.success(msg) if ok else st.error(msg)
                st.rerun()

        if username == "admin":
            st.markdown("---")
            st.markdown("**Zone admin — suppression définitive**")
            confirmer = st.checkbox("Je confirme la suppression définitive de ce fournisseur")
            if st.button("🗑️ Supprimer définitivement", type="primary", disabled=not confirmer):
                ok, msg = supprimer_definitif(code_sel)
                st.success(msg) if ok else st.error(msg)
                st.rerun()

# ---------- LISTE ----------
with onglet4:
    st.subheader("Liste des fournisseurs")
    filtre = st.text_input("Rechercher (code ou nom)", "")
    df_aff = df.copy()
    if filtre.strip():
        f = filtre.strip().upper()
        df_aff = df_aff[
            df_aff["Code"].astype(str).str.upper().str.contains(f, na=False) |
            df_aff["Nom_Fournisseur"].astype(str).str.upper().str.contains(f, na=False)
        ]

    # Affichage lisible du PU
    if not df_aff.empty and "PU_FCFA_KG" in df_aff.columns:
        df_show = df_aff.copy()
        df_show["PU_FCFA_KG"] = df_show["PU_FCFA_KG"].apply(
            lambda x: f"{_pu(x):,.0f}".replace(",", " ")
        )
        st.dataframe(df_show, width="stretch", height=450)
    else:
        st.dataframe(df_aff, width="stretch", height=450)

st.markdown("---")
if os.path.exists(FICHIER_FRS):
    with open(FICHIER_FRS, "rb") as f:
        st.download_button(
            "⬇️ Télécharger fournisseurs.xlsx",
            data=f,
            file_name="fournisseurs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("NAP SARL – Fournisseurs FRS · PU utilisé pour Montant = Poids × PU")