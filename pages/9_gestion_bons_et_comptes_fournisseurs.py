import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Comptes Fournisseurs", page_icon="📒", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")

# ==================== FICHIERS ====================
FICHIER_BONS = "donnees_bons.xlsx"
FICHIER_DECH = "dechargements.xlsx"
FICHIER_FRS = "fournisseurs.xlsx"
FICHIER_MOUVEMENTS = "comptes_fournisseurs.xlsx"  # mouvements débit/crédit
FICHIER_RELEVE = "releve_comptes_fournisseurs.xlsx"
# ==================================================

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

def charger_bons():
    if not os.path.exists(FICHIER_BONS):
        return pd.DataFrame()
    return pd.read_excel(FICHIER_BONS)

def charger_dechargements():
    if not os.path.exists(FICHIER_DECH):
        return pd.DataFrame()
    return pd.read_excel(FICHIER_DECH)

def charger_mouvements():
    if not os.path.exists(FICHIER_MOUVEMENTS):
        return pd.DataFrame()
    try:
        return pd.read_excel(FICHIER_MOUVEMENTS, sheet_name="Mouvements")
    except Exception:
        return pd.read_excel(FICHIER_MOUVEMENTS)

def charger_fournisseurs():
    if not os.path.exists(FICHIER_FRS):
        return pd.DataFrame()
    return pd.read_excel(FICHIER_FRS)

def construire_detail_operations():
    """
    1 ligne = Fournisseur + BC + Camion (+ montants si dispo).
    Sources : bons + déchargements + mouvements de paiement.
    """
    df_b = charger_bons()
    df_d = charger_dechargements()
    df_m = charger_mouvements()

    lignes = []

    if df_b.empty:
        return pd.DataFrame()

    # Index déchargements par N_BC
    dech_par_bc = {}
    if not df_d.empty:
        col_bc = next((c for c in df_d.columns if str(c).upper() in ("N_BC", "REFERENCE", "BON_COMMANDE", "N° BC")), None)
        if col_bc is None:
            for c in df_d.columns:
                if "bc" in str(c).lower() or "reference" in str(c).lower():
                    col_bc = c
                    break
        if col_bc:
            for _, r in df_d.iterrows():
                bc = _txt(r.get(col_bc))
                if not bc:
                    continue
                dech_par_bc.setdefault(bc, []).append(r)

    for _, b in df_b.iterrows():
        ref = _txt(b.get("Reference"))
        if not ref:
            continue
        code_frs = _txt(b.get("Code_Fournisseur"))
        fourn = _txt(b.get("Nom_Fournisseur"))
        camion = _txt(b.get("Numero_Tracteur"))
        quantite_bc = _txt(b.get("Quantite"))
        date_bc = b.get("Date_Emission")
        if isinstance(date_bc, datetime):
            date_bc = date_bc.strftime("%d/%m/%Y")
        else:
            date_bc = _txt(date_bc)

        # Déchargements liés à ce BC
        dechs = dech_par_bc.get(ref, [])
        # Aussi chercher par numéro seul (0301 vs 0301-BC-SHEA/NAP-2026)
        num_seul = ref.split("-")[0] if "-" in ref else ref
        for k, v in dech_par_bc.items():
            if k != ref and (k.startswith(num_seul) or num_seul in k):
                dechs = dechs + v

        if not dechs:
            # Ligne BC sans livraison encore
            lignes.append({
                "Code_FRS": code_frs,
                "Fournisseur": fourn,
                "Bon_de_commande": ref,
                "N_Tracteur_Camion": camion,
                "Statut_BC": "EN COURS",
                "Poids_net_livre_kg": 0.0,
                "Montant_total_FCFA": 0.0,
                "Montant_paye_FCFA": 0.0,
                "Solde_restant_FCFA": 0.0,
                "Pct_paye": 0.0,
                "Date_BC": date_bc,
            })
            continue

        for d in dechs:
            poids = _num(d.get("Poids_Net_KG", d.get("Poids_Net", 0)))
            montant_du = _num(d.get("Montant_Du", d.get("Montant_Total", 0)))
            montant_paye_l = _num(d.get("Montant_Paye", 0))
            solde_l = _num(d.get("Solde_Restant", montant_du - montant_paye_l))
            statut = _txt(d.get("Statut_Paiement", d.get("Statut_Livraison", "DECHARGE")))
            camion_d = _txt(d.get("N_Camion", d.get("Numero_Camion", camion)))

            lignes.append({
                "Code_FRS": code_frs or _txt(d.get("Code_Fournisseur")),
                "Fournisseur": fourn or _txt(d.get("Fournisseur")),
                "Bon_de_commande": ref,
                "N_Tracteur_Camion": camion_d or camion,
                "Statut_BC": statut or "DECHARGE",
                "Poids_net_livre_kg": poids,
                "Montant_total_FCFA": montant_du,
                "Montant_paye_FCFA": montant_paye_l,
                "Solde_restant_FCFA": solde_l if solde_l else (montant_du - montant_paye_l),
                "Pct_paye": (montant_paye_l / montant_du) if montant_du else 0.0,
                "Date_BC": date_bc,
            })

    df_detail = pd.DataFrame(lignes)

    # Enrichir Montant_paye avec mouvements globaux (crédits) répartis n'est pas trivial ;
    # on ajoute une vue synthèse qui agrège les crédits par nom fournisseur.
    return df_detail

def construire_synthese(df_detail, df_mvt):
    """Tableau type image : une ligne par fournisseur, tout calculé."""
    if df_detail is None or df_detail.empty:
        return pd.DataFrame()

    # Crédits (paiements / avances) par fournisseur
    credits = {}
    if df_mvt is not None and not df_mvt.empty:
        col_f = next((c for c in df_mvt.columns if "fourn" in str(c).lower()), None)
        col_c = next((c for c in df_mvt.columns if str(c).lower() in ("credit", "crédit")), None)
        if col_f and col_c:
            for _, r in df_mvt.iterrows():
                f = _txt(r.get(col_f)).upper()
                credits[f] = credits.get(f, 0.0) + _num(r.get(col_c))

    rows = []
    for fourn, g in df_detail.groupby("Fournisseur"):
        if not fourn:
            continue
        bc_uniques = g["Bon_de_commande"].nunique()
        # Rejetés si statut contient rejet
        rejetes = g["Statut_BC"].astype(str).str.lower().str.contains("rejet", na=False)
        nb_rejetes = g.loc[rejetes, "Bon_de_commande"].nunique()
        camions = g["N_Tracteur_Camion"].replace("", pd.NA).dropna().nunique()
        poids = g["Poids_net_livre_kg"].sum()
        total = g["Montant_total_FCFA"].sum()
        # Payé = somme des lignes + crédits mouvements si plus élevé
        paye_lignes = g["Montant_paye_FCFA"].sum()
        paye_mvt = credits.get(fourn.upper(), 0.0)
        paye = max(paye_lignes, paye_mvt) if paye_mvt else paye_lignes
        # Si on a des mouvements globaux, on privilégie le total crédit
        if paye_mvt > 0:
            paye = paye_mvt
        solde = total - paye
        pct = (paye / total) if total else 0.0
        code = ""
        if "Code_FRS" in g.columns:
            codes = g["Code_FRS"].replace("", pd.NA).dropna()
            if len(codes):
                code = codes.iloc[0]
        rows.append({
            "Code_FRS": code,
            "Fournisseur": fourn,
            "Nb_BC_Passes": int(bc_uniques),
            "Nb_BC_Rejetes": int(nb_rejetes),
            "Nb_Camions_Livres": int(camions),
            "Poids_Net_Livre_kg": poids,
            "Montant_Total_FCFA": total,
            "Montant_Paye_FCFA": paye,
            "Solde_Restant_FCFA": solde,
            "Pct_Paye": pct,
        })

    df_s = pd.DataFrame(rows)
    if not df_s.empty:
        df_s = df_s.sort_values("Fournisseur")
    return df_s

def exporter_releve(df_detail, df_synth):
    wb = Workbook()
    # Détail
    ws = wb.active
    ws.title = "Releve_Detaille"
    ws.merge_cells("A1:K1")
    ws["A1"] = "RELEVE DE COMPTES FOURNISSEURS - DETAIL"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = PatternFill("solid", fgColor="244062")

    if not df_detail.empty:
        headers = list(df_detail.columns)
        for c, h in enumerate(headers, 1):
            cell = ws.cell(3, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="203864")
        for r, row in enumerate(df_detail.itertuples(index=False), 4):
            for c, val in enumerate(row, 1):
                ws.cell(r, c, val)

    # Synthèse
    ws2 = wb.create_sheet("Synthese_Fournisseurs")
    ws2.merge_cells("A1:J1")
    ws2["A1"] = "RELEVE DE COMPTES PAR FOURNISSEUR"
    ws2["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws2["A1"].fill = PatternFill("solid", fgColor="244062")

    if not df_synth.empty:
        headers = list(df_synth.columns)
        for c, h in enumerate(headers, 1):
            cell = ws2.cell(3, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="203864")
        for r, row in enumerate(df_synth.itertuples(index=False), 4):
            for c, val in enumerate(row, 1):
                ws2.cell(r, c, val)

    wb.save(FICHIER_RELEVE)
    return FICHIER_RELEVE

# ==================== UI ====================
st.title("📒 Comptes Fournisseurs – NAP SARL")
st.caption("Calcul automatique : Fournisseur → BC → Camion → Livraison → Paiements → Solde")
st.markdown("---")

df_detail = construire_detail_operations()
df_mvt = charger_mouvements()
df_synth = construire_synthese(df_detail, df_mvt)

# KPIs
total_du = df_synth["Montant_Total_FCFA"].sum() if not df_synth.empty else 0
total_paye = df_synth["Montant_Paye_FCFA"].sum() if not df_synth.empty else 0
total_solde = df_synth["Solde_Restant_FCFA"].sum() if not df_synth.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Fournisseurs", len(df_synth))
c2.metric("Total dû", f"{total_du:,.0f} FCFA")
c3.metric("Total payé", f"{total_paye:,.0f} FCFA")
c4.metric("Solde global", f"{total_solde:,.0f} FCFA")

st.markdown("---")

onglet1, onglet2, onglet3 = st.tabs([
    "📊 Relevé par fournisseur",
    "🔍 Détail d'un fournisseur",
    "📋 Opérations (BC × Camion)"
])

# ----- SYNTHÈSE (comme l'image) -----
with onglet1:
    st.subheader("RELEVÉ DE COMPTES PAR FOURNISSEUR")
    if df_synth.empty:
        st.info(
            "Aucune donnée à agréger.\n\n"
            "1. Créez des bons liés à un fournisseur\n"
            "2. Enregistrez des déchargements (montants dus)\n"
            "3. Enregistrez des paiements / avances (crédits)"
        )
    else:
        aff = df_synth.copy()
        aff["Poids_Net_Livre_kg"] = aff["Poids_Net_Livre_kg"].map(lambda x: f"{x:,.0f}")
        aff["Montant_Total_FCFA"] = aff["Montant_Total_FCFA"].map(lambda x: f"{x:,.0f}")
        aff["Montant_Paye_FCFA"] = aff["Montant_Paye_FCFA"].map(lambda x: f"{x:,.0f}")
        aff["Solde_Restant_FCFA"] = aff["Solde_Restant_FCFA"].map(lambda x: f"{x:,.0f}")
        aff["Pct_Paye"] = aff["Pct_Paye"].map(lambda x: f"{x*100:.0f} %")
        st.dataframe(aff, width="stretch", height=420)

        if st.button("⬇️ Générer Excel relevé (détail + synthèse)"):
            path = exporter_releve(df_detail, df_synth)
            st.success(f"Fichier généré : {path}")
            with open(path, "rb") as f:
                st.download_button(
                    "Télécharger releve_comptes_fournisseurs.xlsx",
                    data=f,
                    file_name="releve_comptes_fournisseurs.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ----- DETAIL FOURNISSEUR -----
with onglet2:
    st.subheader("Fiche fournisseur")
    if df_synth.empty:
        st.info("Aucun fournisseur à afficher.")
    else:
        liste = df_synth["Fournisseur"].tolist()
        choix = st.selectbox("Fournisseur", liste)
        row = df_synth[df_synth["Fournisseur"] == choix].iloc[0]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Nb BC", int(row["Nb_BC_Passes"]))
        k2.metric("Camions livrés", int(row["Nb_Camions_Livres"]))
        k3.metric("Montant dû", f"{row['Montant_Total_FCFA']:,.0f}")
        k4.metric("Solde", f"{row['Solde_Restant_FCFA']:,.0f}")

        st.markdown("#### Opérations (BC × Camion)")
        det = df_detail[df_detail["Fournisseur"] == choix]
        st.dataframe(det, width="stretch", height=280)

        st.markdown("#### Paiements / mouvements")
        if df_mvt is not None and not df_mvt.empty:
            col_f = next((c for c in df_mvt.columns if "fourn" in str(c).lower()), None)
            if col_f:
                m = df_mvt[df_mvt[col_f].astype(str).str.upper() == choix.upper()]
                st.dataframe(m if not m.empty else pd.DataFrame({"Info": ["Aucun mouvement"]}), width="stretch")
            else:
                st.caption("Colonne fournisseur introuvable dans les mouvements.")
        else:
            st.caption("Aucun fichier de mouvements. Utilisez la page Crédit / Paiements fournisseurs.")

# ----- DETAIL OPERATIONS -----
with onglet3:
    st.subheader("Toutes les opérations (Fournisseur × BC × Camion)")
    if df_detail.empty:
        st.info("Aucune opération.")
    else:
        st.dataframe(df_detail, width="stretch", height=450)

st.markdown("---")
st.caption(
    "NAP SARL – Comptes fournisseurs automatiques | "
    "Sources : donnees_bons.xlsx + dechargements.xlsx + comptes_fournisseurs.xlsx"
)