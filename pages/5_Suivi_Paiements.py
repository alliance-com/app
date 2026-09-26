import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
from openpyxl import load_workbook
from generer_bon_paiement import generer_bon_paiement

st.set_page_config(page_title="Suivi Paiements", page_icon="💰", layout="wide")

# Vérifier authentification
if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

# ==================== CONFIGURATION ====================
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FICHIER_DECHARGEMENTS = "dechargements.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
# =======================================================

st.title("💰 Suivi des Paiements – NAP SARL")
st.markdown("---")

# ---------- Fonctions ----------
def charger_depenses():
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame()
    try:
        df = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=FEUILLE_DEPENSES, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        return df
    except Exception as e:
        st.error(f"Erreur lecture DEPENSES : {e}")
        return pd.DataFrame()

def charger_dechargements():
    if not os.path.exists(FICHIER_DECHARGEMENTS):
        return pd.DataFrame()
    return pd.read_excel(FICHIER_DECHARGEMENTS)

def prochain_numero_chq(df):
    max_num = 0
    if not df.empty and "No de transaction" in df.columns:
        for val in df["No de transaction"].dropna().astype(str):
            m = re.search(r'CHQ-?(\d+)', val.upper())
            if m:
                max_num = max(max_num, int(m.group(1)))
    return max_num + 1

def prochain_numero_cc(df):
    max_num = 0
    if not df.empty and "No de transaction" in df.columns:
        for val in df["No de transaction"].dropna().astype(str):
            m = re.search(r'CC-?(\d+)', val.upper().replace(" ", ""))
            if m:
                max_num = max(max_num, int(m.group(1)))
    return max_num + 1

def montant_en_lettres(n):
    try:
        n = int(n)
    except Exception:
        return ""
    return f"{n:,} francs CFA".replace(",", " ")

def ajouter_paiement_excel(ligne):
    if not os.path.exists(FICHIER_PAIEMENTS):
        return False, "Fichier Suivi de paiement.xlsx introuvable"
    try:
        wb = load_workbook(FICHIER_PAIEMENTS)
        if FEUILLE_DEPENSES not in wb.sheetnames:
            return False, "Feuille DEPENSES absente"
        ws = wb[FEUILLE_DEPENSES]
        next_row = ws.max_row + 1

        ws.cell(row=next_row, column=2).value = ligne.get("Date")
        ws.cell(row=next_row, column=3).value = ligne.get("No_transaction")
        ws.cell(row=next_row, column=4).value = ligne.get("Type_operation")
        ws.cell(row=next_row, column=5).value = ligne.get("Produit")
        ws.cell(row=next_row, column=6).value = ligne.get("Prestataire")
        ws.cell(row=next_row, column=7).value = ligne.get("Beneficiaire")
        ws.cell(row=next_row, column=8).value = ligne.get("Piece_identite")
        ws.cell(row=next_row, column=9).value = ligne.get("Date_expiration")
        ws.cell(row=next_row, column=10).value = ligne.get("Telephone")
        ws.cell(row=next_row, column=11).value = ligne.get("Ref_doc")
        ws.cell(row=next_row, column=12).value = ligne.get("Description")
        ws.cell(row=next_row, column=13).value = ligne.get("Montant")
        ws.cell(row=next_row, column=14).value = ligne.get("Observation")
        ws.cell(row=next_row, column=15).value = ligne.get("Montant_lettres")
        ws.cell(row=next_row, column=16).value = ligne.get("Mode_paiement")
        ws.cell(row=next_row, column=17).value = ligne.get("Numero_cheque")
        ws.cell(row=next_row, column=18).value = ligne.get("Banque")

        wb.save(FICHIER_PAIEMENTS)
        return True, "Paiement enregistré dans Suivi de paiement.xlsx"
    except Exception as e:
        return False, f"Erreur écriture : {e}"

# ---------- Chargement ----------
df = charger_depenses()
df_dech = charger_dechargements()

# ---------- KPIs ----------
st.subheader("Indicateurs")

total_general = 0
total_shea = 0
nb_cheques = 0
nb_especes = 0
col_montant = None
col_produit = None
col_mode = None

if not df.empty:
    col_montant = next((c for c in df.columns if "montant" in c.lower() and "lettre" not in c.lower()), None)
    col_produit = next((c for c in df.columns if "produit" in c.lower()), None)
    col_mode = next((c for c in df.columns if "mode" in c.lower()), None)

    if col_montant:
        df[col_montant] = pd.to_numeric(df[col_montant], errors="coerce")
        total_general = df[col_montant].sum()
        if col_produit:
            mask_shea = df[col_produit].astype(str).str.upper().str.contains("SHEA", na=False)
            total_shea = df.loc[mask_shea, col_montant].sum()
        if col_mode:
            nb_cheques = df[col_mode].astype(str).str.upper().str.contains("CHEQUE|CHÈQUE", na=False).sum()
            nb_especes = df[col_mode].astype(str).str.upper().str.contains("ESPECE", na=False).sum()

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Total paiements", f"{total_general:,.0f} FCFA")
with col2:
    st.metric("Total SHEA", f"{total_shea:,.0f} FCFA")
with col3:
    st.metric("Nb chèques", nb_cheques)
with col4:
    st.metric("Nb espèces / MM", nb_especes)
with col5:
    non_payes = 0
    if not df_dech.empty and "Statut_Paiement" in df_dech.columns:
        non_payes = (df_dech["Statut_Paiement"].astype(str).str.upper() == "NON PAYE").sum()
    st.metric("Déchargements non payés", non_payes)

st.markdown("---")

# ---------- Filtres ----------
st.subheader("Filtres")
c1, c2, c3, c4 = st.columns(4)
with c1:
    filtre_produit = st.selectbox("Produit", ["Tous", "SHEA", "SOYABEAN", "SERVICE", "EXPORT", "AUTRE"])
with c2:
    filtre_mode = st.selectbox("Mode paiement", ["Tous", "CHEQUE", "ESPECES", "MOBILE MONEY", "VIREMENT"])
with c3:
    filtre_fourn = st.text_input("Fournisseur / Bénéficiaire", "")
with c4:
    annee = st.selectbox("Année", ["2026", "Toutes"])

# ---------- Tableau des paiements ----------
st.subheader("📋 Liste des paiements")
df_aff = df.copy()
if not df_aff.empty:
    if filtre_produit != "Tous" and col_produit:
        if filtre_produit == "AUTRE":
            df_aff = df_aff[~df_aff[col_produit].astype(str).str.upper().str.contains("SHEA|SOYABEAN|SERVICE|EXPORT", na=False)]
        else:
            df_aff = df_aff[df_aff[col_produit].astype(str).str.upper().str.contains(filtre_produit, na=False)]
    if filtre_mode != "Tous" and col_mode:
        df_aff = df_aff[df_aff[col_mode].astype(str).str.upper().str.contains(filtre_mode.replace("È", "E"), na=False)]
    if filtre_fourn:
        mask = pd.Series([False] * len(df_aff))
        for c in df_aff.columns:
            if "fournisseur" in c.lower() or "bénéficiaire" in c.lower() or "beneficiaire" in c.lower():
                mask = mask | df_aff[c].astype(str).str.contains(filtre_fourn, case=False, na=False)
        df_aff = df_aff[mask]

    st.dataframe(df_aff, width="stretch", height=350)
    st.caption(f"{len(df_aff)} ligne(s) affichée(s) / {len(df)} au total")
else:
    st.info("Aucun paiement trouvé dans le fichier.")

st.markdown("---")

# ---------- Nouveau paiement ----------
st.subheader("➕ Enregistrer un nouveau paiement")

options_dech = ["— Saisie libre —"]
if not df_dech.empty and "Statut_Paiement" in df_dech.columns:
    non_payes_df = df_dech[df_dech["Statut_Paiement"].astype(str).str.upper().isin(["NON PAYE", "PARTIEL"])]
    for _, r in non_payes_df.iterrows():
        label = f"{r.get('N_BC', '')} | {r.get('Fournisseur', '')} | Solde: {r.get('Solde_Restant', 0):,.0f} FCFA"
        options_dech.append(label)

choix_dech = st.selectbox("Lier à un déchargement non payé (optionnel)", options_dech)

pre_fourn = ""
pre_montant = 0.0
pre_ref = ""
pre_desc = "PURCHASE OF SHEA"
if choix_dech != "— Saisie libre —" and not df_dech.empty:
    idx = options_dech.index(choix_dech) - 1
    non_payes_df = df_dech[df_dech["Statut_Paiement"].astype(str).str.upper().isin(["NON PAYE", "PARTIEL"])].reset_index(drop=True)
    if idx < len(non_payes_df):
        row = non_payes_df.iloc[idx]
        pre_fourn = str(row.get("Fournisseur", ""))
        pre_montant = float(row.get("Solde_Restant", 0) or 0)
        pre_ref = f"{row.get('N_Camion', '')}_{row.get('Poids_Net_KG', '')} KG"
        pre_desc = f"PURCHASE OF SHEA – BC {row.get('N_BC', '')}"

with st.form("form_paiement", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)

    with col1:
        date_p = st.date_input("Date *", value=datetime.now())
        type_op = st.selectbox("Type d'opération *", ["PURCHASE", "SERVICE", "EXPORT", "ADMIN"])
        produit = st.selectbox("Produit *", ["SHEA", "SOYABEAN", "SERVICE", "FRET EXPORT", "SALARY", "AUTRE"])
        mode = st.selectbox("Mode paiement *", ["CHEQUE", "ESPECES", "MOBILE MONEY", "VIREMENT"])

    with col2:
        prestataire = st.text_input("Prestataire / Fournisseur *", value=pre_fourn)
        beneficiaire = st.text_input("Nom du bénéficiaire *", value=pre_fourn)
        piece_id = st.text_input("N° Pièce d'identité")
        date_exp = st.text_input("Date expiration pièce")
        telephone = st.text_input("Téléphone / Adresse")

    with col3:
        montant = st.number_input("Montant (FCFA) *", min_value=0.0, value=pre_montant, step=1000.0)
        observation = st.selectbox("Observation", ["PAYE", "AVANCE", "100% PAYE", "PARTIEL"])
        numero_cheque = st.text_input("N° Chèque (si chèque)")
        banque = st.selectbox("Banque", ["ECOBANK BENIN", "CAISSE", "BIIC", "MOBILE MONEY", "AUTRE"])
        ref_doc = st.text_input("Réf. doc (camion / ticket)", value=pre_ref)

    description = st.text_input("Description / Objet du paiement *", value=pre_desc)

    if mode == "CHEQUE":
        num_auto = f"CHQ-{prochain_numero_chq(df):04d}-2026"
    else:
        num_auto = f"CC-{prochain_numero_cc(df):04d}-2026"

    no_transaction = st.text_input("N° de transaction *", value=num_auto)

    submitted = st.form_submit_button("💾 Enregistrer le paiement", type="primary", width="stretch")

    if submitted:
        if not prestataire.strip() or not beneficiaire.strip() or montant <= 0 or not description.strip():
            st.error("Champs obligatoires manquants (*).")
        else:
            ligne = {
                "Date": date_p.strftime("%d/%m/%Y"),
                "No_transaction": no_transaction.strip(),
                "Type_operation": type_op,
                "Produit": produit,
                "Prestataire": prestataire.strip(),
                "Beneficiaire": beneficiaire.strip(),
                "Piece_identite": piece_id.strip(),
                "Date_expiration": date_exp.strip(),
                "Telephone": telephone.strip(),
                "Ref_doc": ref_doc.strip(),
                "Description": description.strip(),
                "Montant": montant,
                "Observation": observation,
                "Montant_lettres": montant_en_lettres(montant),
                "Mode_paiement": mode,
                "Numero_cheque": numero_cheque.strip() if mode == "CHEQUE" else "-",
                "Banque": banque
            }
            ok, msg = ajouter_paiement_excel(ligne)
            if ok:
                st.success(f"✅ {msg}")
                st.success(f"Transaction **{no_transaction}** — {montant:,.0f} FCFA")

                # Génération automatique du bon de paiement
                try:
                    chemin_bon = generer_bon_paiement(ligne)
                    st.session_state["dernier_bon_paiement"] = chemin_bon
                    st.success(f"📄 Bon de paiement généré : `{os.path.basename(chemin_bon)}`")
                except Exception as e:
                    st.warning(f"Paiement enregistré, mais erreur génération du bon : {e}")
            else:
                st.error(msg)

# ---------- Téléchargement du dernier bon généré ----------
if st.session_state.get("dernier_bon_paiement") and os.path.exists(st.session_state["dernier_bon_paiement"]):
    st.markdown("---")
    st.subheader("📄 Bon de paiement venant d'être généré")
    chemin = st.session_state["dernier_bon_paiement"]
    with open(chemin, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger le Bon de paiement (Word)",
            data=f,
            file_name=os.path.basename(chemin),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            width="stretch"
        )

# ---------- Réimprimer un bon existant ----------
st.markdown("---")
st.subheader("🖨️ Réimprimer un bon de paiement existant")

if not df.empty and "No de transaction" in df.columns:
    liste_trans = df["No de transaction"].dropna().astype(str).unique().tolist()
    liste_trans = sorted(liste_trans, reverse=True)
    trans_choisie = st.selectbox("Choisir le N° de transaction", ["—"] + liste_trans)

    if trans_choisie and trans_choisie != "—":
        if st.button("📄 Générer / Télécharger ce bon de paiement"):
            row = df[df["No de transaction"].astype(str) == trans_choisie].iloc[0]

            def get(col_candidates, default=""):
                for c in df.columns:
                    for cand in col_candidates:
                        if cand.lower() in str(c).lower():
                            val = row.get(c, default)
                            return "" if pd.isna(val) else str(val)
                return default

            montant_val = get(["montant"], "0")
            try:
                montant_float = float(str(montant_val).replace(" ", "").replace(",", "."))
            except Exception:
                montant_float = 0.0

            data_bon = {
                "Date": get(["date"]),
                "No_transaction": trans_choisie,
                "Type_operation": get(["type"]),
                "Produit": get(["produit"]),
                "Prestataire": get(["prestataire", "fournisseur"]),
                "Beneficiaire": get(["bénéficiaire", "beneficiaire", "nom"]),
                "Piece_identite": get(["piece", "identite"]),
                "Date_expiration": get(["expiration"]),
                "Telephone": get(["telephone", "adresse", "addresse"]),
                "Ref_doc": get(["ref"]),
                "Description": get(["description", "objet"]),
                "Montant": montant_float,
                "Observation": get(["observation"]),
                "Montant_lettres": get(["lettre"]) or montant_en_lettres(montant_float),
                "Mode_paiement": get(["mode"]),
                "Numero_cheque": get(["cheque", "chèque"]),
                "Banque": get(["banque"]),
            }
            try:
                chemin = generer_bon_paiement(data_bon)
                with open(chemin, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Télécharger Bon_Paiement_{trans_choisie}.docx",
                        data=f,
                        file_name=f"Bon_Paiement_{trans_choisie}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_bon_{trans_choisie}"
                    )
                st.success("Bon généré avec succès.")
            except Exception as e:
                st.error(f"Erreur : {e}")

# ---------- Téléchargement Excel ----------
st.markdown("---")
if os.path.exists(FICHIER_PAIEMENTS):
    with open(FICHIER_PAIEMENTS, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger Suivi de paiement.xlsx",
            data=f,
            file_name="Suivi_de_paiement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("NAP SARL – Suivi des paiements (chèques, caisse, mobile money)")