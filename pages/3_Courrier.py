import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re

st.set_page_config(page_title="Courrier", page_icon="📨", layout="wide")

# ---------- Auth (multipage) ----------
if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""

# ==================== CONFIG ====================
FICHIER_COURRIER = "courrier.xlsx"
TYPES = ["Arrivée", "Départ"]
NATURES = [
    "Lettre", "Facture", "BL / Connaissement", "Certificat", "Déclaration",
    "Contrat", "Chèque / Paiement", "Relance", "Administratif", "Autre"
]
STATUTS = ["Reçu", "En traitement", "Transmis", "Archivé", "Répondu", "Envoyé"]
PRIORITES = ["Normale", "Urgente", "Confidentielle"]
# =================================================


def charger_courrier():
    colonnes = [
        "ID", "Date", "Type", "Nature", "Objet", "Expediteur", "Destinataire",
        "Reference", "Priorite", "Statut", "Service", "Observations",
        "Heure_Saisie", "Utilisateur"
    ]
    if os.path.exists(FICHIER_COURRIER):
        df = pd.read_excel(FICHIER_COURRIER)
        for c in colonnes:
            if c not in df.columns:
                df[c] = ""
        return df[colonnes]
    df = pd.DataFrame(columns=colonnes)
    df.to_excel(FICHIER_COURRIER, index=False)
    return df


def prochain_id(df):
    if df.empty or "ID" not in df.columns:
        return "CR-0001"
    nums = []
    for v in df["ID"].dropna().astype(str):
        m = re.search(r"(\d+)", v)
        if m:
            nums.append(int(m.group(1)))
    n = max(nums) + 1 if nums else 1
    return f"CR-{n:04d}"


def sauvegarder(df):
    df.to_excel(FICHIER_COURRIER, index=False)


# ---------- Sidebar : nouveau courrier ----------
st.sidebar.header("➕ Nouveau courrier")

df = charger_courrier()
nouvel_id = prochain_id(df)

with st.sidebar.form("form_courrier", clear_on_submit=True):
    type_c = st.radio("Type *", options=TYPES, horizontal=True)
    nature = st.selectbox("Nature *", options=NATURES)
    priorite = st.selectbox("Priorité", options=PRIORITES, index=0)
    statut = st.selectbox(
        "Statut",
        options=STATUTS,
        index=0 if type_c == "Arrivée" else STATUTS.index("Envoyé") if "Envoyé" in STATUTS else 0,
    )
    date_c = st.date_input("Date *", value=date.today())
    objet = st.text_input("Objet *", placeholder="Objet du courrier")
    expediteur = st.text_input(
        "Expéditeur *",
        placeholder="Qui envoie" if type_c == "Arrivée" else "NAP SARL (ou service)",
        value="" if type_c == "Arrivée" else "NAP SARL",
    )
    destinataire = st.text_input(
        "Destinataire *",
        placeholder="Qui reçoit" if type_c == "Départ" else "NAP SARL (ou service)",
        value="NAP SARL" if type_c == "Arrivée" else "",
    )
    reference = st.text_input("Référence / N° pièce", placeholder="BL, facture, n° lettre…")
    service = st.text_input("Service concerné", placeholder="Achat, Logistique, Finance…")
    observations = st.text_area("Observations", placeholder="Notes, suite à donner…")
    utilisateur = st.text_input("Saisi par", value=name)

    submitted = st.form_submit_button("💾 Enregistrer", width="stretch")

    if submitted:
        if not objet.strip():
            st.error("L'objet est obligatoire.")
        elif not expediteur.strip() or not destinataire.strip():
            st.error("Expéditeur et destinataire sont obligatoires.")
        else:
            df = charger_courrier()
            ligne = {
                "ID": nouvel_id,
                "Date": date_c.strftime("%d/%m/%Y"),
                "Type": type_c,
                "Nature": nature,
                "Objet": objet.strip(),
                "Expediteur": expediteur.strip(),
                "Destinataire": destinataire.strip(),
                "Reference": reference.strip(),
                "Priorite": priorite,
                "Statut": statut,
                "Service": service.strip(),
                "Observations": observations.strip(),
                "Heure_Saisie": datetime.now().strftime("%H:%M:%S"),
                "Utilisateur": utilisateur.strip(),
            }
            df = pd.concat([df, pd.DataFrame([ligne])], ignore_index=True)
            sauvegarder(df)
            st.success(f"✅ Courrier **{nouvel_id}** enregistré ({type_c})")
            st.rerun()

# ---------- Zone principale ----------
st.title("📨 Gestion du Courrier — Arrivée & Départ")
st.caption("Registre local des courriers · Source : courrier.xlsx")

df = charger_courrier()

# KPI
nb_arr = len(df[df["Type"] == "Arrivée"]) if not df.empty else 0
nb_dep = len(df[df["Type"] == "Départ"]) if not df.empty else 0
nb_urg = len(df[df["Priorite"] == "Urgente"]) if not df.empty else 0
nb_trait = len(df[df["Statut"].isin(["Reçu", "En traitement"])]) if not df.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Courriers arrivée", nb_arr)
c2.metric("Courriers départ", nb_dep)
c3.metric("Urgents", nb_urg)
c4.metric("À traiter", nb_trait)

st.markdown("---")

# Filtres
f1, f2, f3, f4 = st.columns(4)
with f1:
    filtre_type = st.selectbox("Filtrer type", ["Tous", "Arrivée", "Départ"])
with f2:
    filtre_statut = st.selectbox("Filtrer statut", ["Tous"] + STATUTS)
with f3:
    filtre_nature = st.selectbox("Filtrer nature", ["Tous"] + NATURES)
with f4:
    recherche = st.text_input("Recherche (objet, réf, nom…)", placeholder="Texte libre")

df_f = df.copy()
if not df_f.empty:
    if filtre_type != "Tous":
        df_f = df_f[df_f["Type"] == filtre_type]
    if filtre_statut != "Tous":
        df_f = df_f[df_f["Statut"] == filtre_statut]
    if filtre_nature != "Tous":
        df_f = df_f[df_f["Nature"] == filtre_nature]
    if recherche.strip():
        q = recherche.strip().lower()
        mask = (
            df_f["Objet"].astype(str).str.lower().str.contains(q, na=False)
            | df_f["Reference"].astype(str).str.lower().str.contains(q, na=False)
            | df_f["Expediteur"].astype(str).str.lower().str.contains(q, na=False)
            | df_f["Destinataire"].astype(str).str.lower().str.contains(q, na=False)
            | df_f["ID"].astype(str).str.lower().str.contains(q, na=False)
        )
        df_f = df_f[mask]

# Onglets Arrivée / Départ / Tout
tab_all, tab_arr, tab_dep = st.tabs(["📋 Tous", "📥 Arrivée", "📤 Départ"])


def afficher_table(data):
    if data.empty:
        st.info("Aucun courrier pour ce filtre.")
        return
    st.dataframe(data, width="stretch", height=380, hide_index=True)


with tab_all:
    afficher_table(df_f)
with tab_arr:
    afficher_table(df_f[df_f["Type"] == "Arrivée"] if not df_f.empty else df_f)
with tab_dep:
    afficher_table(df_f[df_f["Type"] == "Départ"] if not df_f.empty else df_f)

# Mise à jour statut
st.markdown("---")
st.subheader("🔄 Mettre à jour le statut")
if df.empty:
    st.info("Aucun courrier à modifier.")
else:
    options_id = df["ID"].astype(str).tolist()
    col_u1, col_u2, col_u3 = st.columns([2, 2, 1])
    with col_u1:
        id_sel = st.selectbox("ID courrier", options=options_id)
    with col_u2:
        nouveau_statut = st.selectbox("Nouveau statut", options=STATUTS, key="upd_statut")
    with col_u3:
        st.write("")
        st.write("")
        if st.button("Appliquer", type="primary", width="stretch"):
            df = charger_courrier()
            df.loc[df["ID"].astype(str) == str(id_sel), "Statut"] = nouveau_statut
            sauvegarder(df)
            st.success(f"Statut de **{id_sel}** → {nouveau_statut}")
            st.rerun()

# Suppression admin
if username == "admin":
    st.markdown("---")
    st.subheader("🗑️ Supprimer un courrier (Admin)")
    if not df.empty:
        id_del = st.selectbox("ID à supprimer", options=df["ID"].astype(str).tolist(), key="del_id")
        if st.button("Supprimer", type="primary"):
            df = charger_courrier()
            df = df[df["ID"].astype(str) != str(id_del)]
            sauvegarder(df)
            st.success(f"{id_del} supprimé.")
            st.rerun()

# Téléchargement
st.markdown("---")
col_dl, col_ref = st.columns([1, 3])
with col_dl:
    if os.path.exists(FICHIER_COURRIER):
        with open(FICHIER_COURRIER, "rb") as f:
            st.download_button(
                "⬇️ Télécharger Excel",
                data=f,
                file_name="courrier.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
with col_ref:
    if st.button("🔄 Actualiser"):
        st.rerun()

st.caption("NAP SARL — Registre courrier arrivée / départ · Fichier : courrier.xlsx")
