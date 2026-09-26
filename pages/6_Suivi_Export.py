import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Suivi Export", page_icon="🚢", layout="wide")

if not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

# ==================== CONFIG ====================
FICHIER = "exports_suivi.xlsx"

PIECES = [
    ("P01", "Facture commerciale", True),
    ("P02", "Packing list", True),
    ("P03", "Certificat d'origine", True),
    ("P04", "Certificat phytosanitaire", True),
    ("P05", "COQ / COW", True),
    ("P06", "Analyse labo / qualité", False),
    ("P07", "Booking confirmé", True),
    ("P08", "VGM / note de poids", True),
    ("P09", "Numéros de scellés", True),
    ("P10", "Déclaration export (EX1)", True),
    ("P11", "BAE / validation douane", True),
    ("P12", "Manifeste export visé", True),
    ("P13", "B/L draft", True),
    ("P14", "B/L final / Telex", True),
    ("P15", "BESC / CTN (si requis)", False),
]

ETAPES = [
    (1, "Contrat / confirmation vente"),
    (2, "Choix transitaire + mandat"),
    (3, "Booking obtenu"),
    (4, "Conteneurs positionnés"),
    (5, "Empotage terminé"),
    (6, "Pesage + VGM"),
    (7, "Scellés posés"),
    (8, "Pièces complètes (checklist)"),
    (9, "Déclaration douane déposée"),
    (10, "Contrôle douane OK / BAE"),
    (11, "Transfert au terminal"),
    (12, "Manifeste export visé"),
    (13, "Embarquement (vu à bord)"),
    (14, "B/L final / Telex reçu"),
    (15, "Envoi docs à acheteur / banque"),
    (16, "Clôture dossier"),
]

# ==================== INIT FICHIER ====================
def init_fichier():
    if os.path.exists(FICHIER):
        return
    wb = Workbook()
    # Dossiers
    ws = wb.active
    ws.title = "Dossiers"
    ws.append([
        "N_Dossier", "Booking", "BL", "Produit", "Destination", "Transitaire",
        "Ligne_Maritime", "Statut", "Pct_Avancement", "Date_Creation",
        "Responsable", "Commentaire"
    ])
    # Pieces
    ws2 = wb.create_sheet("Pieces")
    ws2.append([
        "N_Dossier", "Code_Piece", "Libelle", "Obligatoire", "Statut",
        "Date_Validation", "Valide_Par", "Commentaire"
    ])
    # Etapes
    ws3 = wb.create_sheet("Etapes")
    ws3.append([
        "N_Dossier", "N_Etape", "Libelle", "Statut", "Date_Debut",
        "Date_Fin", "Commentaire"
    ])
    # Historique
    ws4 = wb.create_sheet("Historique")
    ws4.append(["Date_Heure", "N_Dossier", "Utilisateur", "Action", "Detail"])
    wb.save(FICHIER)

def charger_dossiers():
    init_fichier()
    return pd.read_excel(FICHIER, sheet_name="Dossiers")

def charger_pieces(n_dossier=None):
    init_fichier()
    df = pd.read_excel(FICHIER, sheet_name="Pieces")
    if n_dossier:
        df = df[df["N_Dossier"].astype(str) == str(n_dossier)]
    return df

def charger_etapes(n_dossier=None):
    init_fichier()
    df = pd.read_excel(FICHIER, sheet_name="Etapes")
    if n_dossier:
        df = df[df["N_Dossier"].astype(str) == str(n_dossier)]
    return df

def prochain_numero():
    df = charger_dossiers()
    if df.empty:
        return "EXP-2026-0001"
    nums = []
    for v in df["N_Dossier"].dropna().astype(str):
        if "EXP-2026-" in v:
            try:
                nums.append(int(v.split("-")[-1]))
            except Exception:
                pass
    n = max(nums) + 1 if nums else 1
    return f"EXP-2026-{n:04d}"

def ajouter_historique(n_dossier, action, detail):
    wb = load_workbook(FICHIER)
    ws = wb["Historique"]
    ws.append([
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        n_dossier, name, action, detail
    ])
    wb.save(FICHIER)

def creer_dossier(data):
    n_dossier = prochain_numero()
    wb = load_workbook(FICHIER)

    # Dossier
    ws = wb["Dossiers"]
    ws.append([
        n_dossier,
        data.get("Booking", ""),
        data.get("BL", ""),
        data.get("Produit", ""),
        data.get("Destination", ""),
        data.get("Transitaire", ""),
        data.get("Ligne_Maritime", ""),
        "EN COURS",
        0,
        datetime.now().strftime("%d/%m/%Y"),
        name,
        data.get("Commentaire", "")
    ])

    # Pièces
    ws2 = wb["Pieces"]
    for code, lib, oblig in PIECES:
        ws2.append([
            n_dossier, code, lib, "Oui" if oblig else "Non",
            "MANQUANT", "", "", ""
        ])

    # Étapes
    ws3 = wb["Etapes"]
    for n, lib in ETAPES:
        ws3.append([n_dossier, n, lib, "À FAIRE", "", "", ""])

    wb.save(FICHIER)
    ajouter_historique(n_dossier, "Création", f"Dossier créé – Booking {data.get('Booking', '')}")
    return n_dossier

def maj_piece(n_dossier, code_piece, statut, commentaire=""):
    wb = load_workbook(FICHIER)
    ws = wb["Pieces"]
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == str(n_dossier) and str(row[1].value) == str(code_piece):
            row[4].value = statut  # Statut
            if statut == "OK":
                row[5].value = datetime.now().strftime("%d/%m/%Y")
                row[6].value = name
            row[7].value = commentaire
            break
    wb.save(FICHIER)
    ajouter_historique(n_dossier, "Pièce", f"{code_piece} → {statut}")

def maj_etape(n_dossier, n_etape, statut, commentaire=""):
    wb = load_workbook(FICHIER)
    ws = wb["Etapes"]
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == str(n_dossier) and int(row[1].value) == int(n_etape):
            ancien = row[3].value
            row[3].value = statut
            if statut == "EN COURS" and not row[4].value:
                row[4].value = datetime.now().strftime("%d/%m/%Y")
            if statut == "FAIT":
                row[5].value = datetime.now().strftime("%d/%m/%Y")
            row[6].value = commentaire
            break
    wb.save(FICHIER)
    ajouter_historique(n_dossier, "Étape", f"Étape {n_etape} → {statut}")
    recalculer_avancement(n_dossier)

def recalculer_avancement(n_dossier):
    df_e = charger_etapes(n_dossier)
    if df_e.empty:
        return
    total = len(df_e)
    faits = (df_e["Statut"].astype(str) == "FAIT").sum()
    pct = int(round(100 * faits / total)) if total else 0

    # Statut global
    if (df_e["Statut"].astype(str) == "BLOQUÉ").any():
        statut = "BLOQUÉ"
    elif pct >= 100:
        statut = "CLOTURÉ"
    elif (df_e["Statut"].astype(str) == "FAIT").any() or (df_e["Statut"].astype(str) == "EN COURS").any():
        # Si embarquement fait
        emb = df_e[df_e["N_Etape"] == 13]
        if not emb.empty and str(emb.iloc[0]["Statut"]) == "FAIT":
            statut = "EMBARQUÉ"
        else:
            statut = "EN COURS"
    else:
        statut = "EN COURS"

    wb = load_workbook(FICHIER)
    ws = wb["Dossiers"]
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == str(n_dossier):
            row[7].value = statut
            row[8].value = pct
            break
    wb.save(FICHIER)

def pieces_manquantes_obligatoires(n_dossier):
    df = charger_pieces(n_dossier)
    if df.empty:
        return []
    mask = (df["Obligatoire"].astype(str).str.upper() == "OUI") & (df["Statut"].astype(str).str.upper() != "OK")
    return df.loc[mask, "Libelle"].tolist()

def etapes_bloquees(n_dossier):
    df = charger_etapes(n_dossier)
    if df.empty:
        return []
    return df[df["Statut"].astype(str).str.upper() == "BLOQUÉ"][["N_Etape", "Libelle", "Commentaire"]].to_dict("records")

# ==================== UI ====================
st.title("🚢 Suivi Export – par B/L / Booking")
st.caption("Ne rien oublier • Voir ce qui bloque • Corriger")
st.markdown("---")

# ---------- KPIs ----------
df_dossiers = charger_dossiers()
nb_cours = 0
nb_bloque = 0
nb_embarque = 0
nb_cloture = 0
if not df_dossiers.empty and "Statut" in df_dossiers.columns:
    nb_cours = (df_dossiers["Statut"].astype(str) == "EN COURS").sum()
    nb_bloque = (df_dossiers["Statut"].astype(str) == "BLOQUÉ").sum()
    nb_embarque = (df_dossiers["Statut"].astype(str) == "EMBARQUÉ").sum()
    nb_cloture = (df_dossiers["Statut"].astype(str) == "CLOTURÉ").sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total dossiers", len(df_dossiers))
c2.metric("En cours", nb_cours)
c3.metric("Bloqués", nb_bloque)
c4.metric("Embarqués", nb_embarque)
c5.metric("Clôturés", nb_cloture)

st.markdown("---")

onglet1, onglet2, onglet3 = st.tabs([
    "📋 Liste des dossiers",
    "➕ Nouveau dossier",
    "🔍 Fiche détaillée"
])

# ========== ONGLET 1 : LISTE ==========
with onglet1:
    st.subheader("Tous les dossiers export")
    if df_dossiers.empty:
        st.info("Aucun dossier. Créez-en un dans l'onglet « Nouveau dossier ».")
    else:
        filtre = st.selectbox("Filtrer par statut", ["Tous", "EN COURS", "BLOQUÉ", "EMBARQUÉ", "CLOTURÉ"])
        df_aff = df_dossiers.copy()
        if filtre != "Tous":
            df_aff = df_aff[df_aff["Statut"].astype(str) == filtre]
        st.dataframe(df_aff, width="stretch", height=350)

        # Alertes bloqués
        bloques = df_dossiers[df_dossiers["Statut"].astype(str) == "BLOQUÉ"]
        if not bloques.empty:
            st.error(f"⚠️ {len(bloques)} dossier(s) BLOQUÉ(S) – à traiter en priorité")
            for _, r in bloques.iterrows():
                st.write(f"• **{r['N_Dossier']}** | BL: {r.get('BL', '—')} | Transitaire: {r.get('Transitaire', '—')}")

# ========== ONGLET 2 : NOUVEAU ==========
with onglet2:
    st.subheader("Créer un nouveau dossier export")
    with st.form("form_nouveau"):
        col1, col2 = st.columns(2)
        with col1:
            booking = st.text_input("N° Booking *")
            bl = st.text_input("N° B/L (si déjà connu)")
            produit = st.selectbox("Produit", ["SHEA", "SOYABEAN", "AUTRE"])
            destination = st.text_input("Destination *", placeholder="Ex: Chine, Europe…")
        with col2:
            transitaire = st.text_input("Transitaire *", placeholder="Ex: TRANS SOHA")
            ligne = st.text_input("Ligne maritime", placeholder="Ex: MSC, Maersk")
            commentaire = st.text_area("Commentaire", height=80)

        submitted = st.form_submit_button("Créer le dossier", type="primary")
        if submitted:
            if not booking.strip() or not destination.strip() or not transitaire.strip():
                st.error("Booking, Destination et Transitaire sont obligatoires.")
            else:
                n = creer_dossier({
                    "Booking": booking.strip(),
                    "BL": bl.strip(),
                    "Produit": produit,
                    "Destination": destination.strip(),
                    "Transitaire": transitaire.strip(),
                    "Ligne_Maritime": ligne.strip(),
                    "Commentaire": commentaire.strip()
                })
                st.success(f"✅ Dossier **{n}** créé avec checklist pièces + 16 étapes.")
                st.rerun()

# ========== ONGLET 3 : FICHE ==========
with onglet3:
    st.subheader("Fiche détaillée d’un dossier")
    if df_dossiers.empty:
        st.info("Aucun dossier à afficher.")
    else:
        options = df_dossiers["N_Dossier"].astype(str).tolist()
        # Afficher aussi BL si dispo
        labels = []
        for _, r in df_dossiers.iterrows():
            bl_aff = r.get("BL", "") or "—"
            labels.append(f"{r['N_Dossier']} | BL: {bl_aff} | {r.get('Statut', '')}")
        choix_idx = st.selectbox("Choisir le dossier", range(len(options)), format_func=lambda i: labels[i])
        n_dossier = options[choix_idx]

        row = df_dossiers[df_dossiers["N_Dossier"].astype(str) == n_dossier].iloc[0]

        # Infos
        st.markdown(f"### {n_dossier} — **{row.get('Statut', '')}** ({row.get('Pct_Avancement', 0)} %)")
        col_a, col_b, col_c = st.columns(3)
        col_a.write(f"**Booking :** {row.get('Booking', '—')}")
        col_a.write(f"**B/L :** {row.get('BL', '—')}")
        col_b.write(f"**Produit :** {row.get('Produit', '—')}")
        col_b.write(f"**Destination :** {row.get('Destination', '—')}")
        col_c.write(f"**Transitaire :** {row.get('Transitaire', '—')}")
        col_c.write(f"**Ligne :** {row.get('Ligne_Maritime', '—')}")

        # Alertes
        manquantes = pieces_manquantes_obligatoires(n_dossier)
        bloques_e = etapes_bloquees(n_dossier)
        if manquantes:
            st.warning("⚠️ **Pièces obligatoires manquantes :** " + " · ".join(manquantes))
        if bloques_e:
            for b in bloques_e:
                st.error(f"🔴 **Étape {b['N_Etape']} bloquée :** {b['Libelle']} — {b.get('Commentaire', '')}")

        st.markdown("---")

        # --- PIÈCES ---
        st.markdown("#### 📎 Checklist des pièces")
        df_p = charger_pieces(n_dossier)
        for _, p in df_p.iterrows():
            code = p["Code_Piece"]
            lib = p["Libelle"]
            oblig = p["Obligatoire"]
            statut = str(p["Statut"])
            colx, coly, colz = st.columns([4, 2, 3])
            with colx:
                badge = "🔴" if statut != "OK" and str(oblig).upper() == "OUI" else ("✅" if statut == "OK" else "⚪")
                st.write(f"{badge} **{lib}** {'*(obligatoire)*' if str(oblig).upper() == 'OUI' else ''}")
            with coly:
                nouveau = st.selectbox(
                    "Statut", ["MANQUANT", "OK", "N/A"],
                    index=["MANQUANT", "OK", "N/A"].index(statut) if statut in ["MANQUANT", "OK", "N/A"] else 0,
                    key=f"piece_{n_dossier}_{code}"
                )
            with colz:
                if st.button("Enregistrer", key=f"btn_p_{n_dossier}_{code}"):
                    maj_piece(n_dossier, code, nouveau)
                    st.rerun()

        st.markdown("---")

        # --- ÉTAPES ---
        st.markdown("#### 📍 Étapes du process")
        df_e = charger_etapes(n_dossier)
        for _, e in df_e.iterrows():
            n_et = int(e["N_Etape"])
            lib = e["Libelle"]
            statut = str(e["Statut"])
            icon = {"À FAIRE": "⬜", "EN COURS": "🔄", "FAIT": "✅", "BLOQUÉ": "🔴", "N/A": "➖"}.get(statut, "⬜")
            col1, col2, col3 = st.columns([4, 2, 3])
            with col1:
                st.write(f"{icon} **{n_et}. {lib}**")
                if e.get("Date_Debut") or e.get("Date_Fin"):
                    st.caption(f"Début: {e.get('Date_Debut', '—')} | Fin: {e.get('Date_Fin', '—')}")
            with col2:
                nouveau_e = st.selectbox(
                    "Statut", ["À FAIRE", "EN COURS", "FAIT", "BLOQUÉ", "N/A"],
                    index=["À FAIRE", "EN COURS", "FAIT", "BLOQUÉ", "N/A"].index(statut) if statut in ["À FAIRE", "EN COURS", "FAIT", "BLOQUÉ", "N/A"] else 0,
                    key=f"etape_{n_dossier}_{n_et}"
                )
            with col3:
                com = st.text_input("Commentaire / motif blocage", value=str(e.get("Commentaire", "") or ""), key=f"com_{n_dossier}_{n_et}")
                if st.button("Maj étape", key=f"btn_e_{n_dossier}_{n_et}"):
                    # Contrôle anti-oubli : avant de valider étape 8 (pièces complètes)
                    if n_et == 8 and nouveau_e == "FAIT":
                        manq = pieces_manquantes_obligatoires(n_dossier)
                        if manq:
                            st.error("Impossible : pièces obligatoires encore manquantes → " + ", ".join(manq))
                        else:
                            maj_etape(n_dossier, n_et, nouveau_e, com)
                            st.rerun()
                    else:
                        maj_etape(n_dossier, n_et, nouveau_e, com)
                        st.rerun()

        # Mise à jour BL si besoin
        st.markdown("---")
        st.markdown("#### ✏️ Mettre à jour le N° B/L")
        nouveau_bl = st.text_input("N° B/L", value=str(row.get("BL", "") or ""))
        if st.button("Enregistrer le B/L"):
            wb = load_workbook(FICHIER)
            ws = wb["Dossiers"]
            for r in ws.iter_rows(min_row=2):
                if str(r[0].value) == str(n_dossier):
                    r[2].value = nouveau_bl.strip()
                    break
            wb.save(FICHIER)
            ajouter_historique(n_dossier, "BL", f"B/L mis à jour : {nouveau_bl}")
            st.success("B/L enregistré.")
            st.rerun()

# Téléchargement
st.markdown("---")
if os.path.exists(FICHIER):
    with open(FICHIER, "rb") as f:
        st.download_button(
            "⬇️ Télécharger exports_suivi.xlsx",
            data=f,
            file_name="exports_suivi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("NAP SARL – Suivi export par B/L • Checklist pièces & étapes • Alertes automatiques")