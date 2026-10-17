import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
import streamlit_authenticator as stauth
import glob

# ==================== CONFIGURATION ====================
FICHIER_EXCEL = "donnees_bons.xlsx"
FICHIER_FRS = "fournisseurs.xlsx"

# Produit → suffixe de référence automatique
PRODUITS_SUFFIXES = {
    "Shea Nut (Noir de Karité)": "-BC-SHEA/NAP-2026",
    "Soja (Graines de Soja)": "-BC-SOYABEAN/NAP-2026",
    "Cotton Seed (Graines de Coton)": "-BC-COTTON/NAP-2026",
}
# =======================================================

st.set_page_config(page_title="Bons de Commande NAP", page_icon="📄", layout="wide")

# ==================== AUTHENTIFICATION ====================
config = {
    'credentials': {
        'usernames': {
            'admin': {
                'email': 'alliancebodjrenou8@gmail.com',
                'name': 'Admin',
                'password': '69060724@'
            },
            'Manager': {
                'email': 'logistics.napbj@gmail.com',
                'name': 'Deepak',
                'password': 'Deepak2026'
            },
            'Hospice': {
                'email': 'logistics.napbj@gmail.com',
                'name': 'Hospice',
                'password': 'Hospice2026'
            }
        }
    },
    'cookie': {
        'expiry_days': 30,
        'key': 'nap_secret_key_2026_changez_moi',
        'name': 'nap_auth_cookie'
    },
    'preauthorized': {
        'emails': []
    }
}

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

authenticator.login(location='main')

if st.session_state.get("authentication_status") is False:
    st.error("Nom d'utilisateur ou mot de passe incorrect")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("Veuillez entrer vos identifiants pour accéder à l'application")
    st.stop()

name = st.session_state.get("name")
username = st.session_state.get("username")

with st.sidebar:
    st.success(f"Connecté : **{name}**")
    st.caption(f"Compte : `{username}`")
    authenticator.logout(location='sidebar')
    st.divider()

# ---------- Fonctions ----------
def extraire_numero(reference):
    if not reference:
        return 0
    base = os.path.basename(str(reference)).split("-")[0]
    match = re.search(r'(\d+)', base)
    if match:
        return int(match.group(1))
    match_fallback = re.search(r'(\d+)', str(reference))
    return int(match_fallback.group(1)) if match_fallback else 0

def charger_fournisseurs_actifs():
    """Retourne dict { 'FRS00006 — VLAVONOU PIERRE': (code, nom) }"""
    if not os.path.exists(FICHIER_FRS):
        return {}
    df = pd.read_excel(FICHIER_FRS)
    if "Actif" in df.columns:
        df = df[df["Actif"].astype(str).str.upper() == "OUI"]
    options = {}
    for _, row in df.iterrows():
        code = str(row.get("Code", "")).strip()
        nom = str(row.get("Nom_Fournisseur", "")).strip()
        if code and nom and nom.lower() not in ("nan", "none"):
            label = f"{code} — {nom}"
            options[label] = (code, nom)
    return options

def charger_donnees():
    colonnes_ordonnees = [
        "Reference", "Date_Emission", "Heure_Creation", "Utilisateur", "Produit",
        "Code_Fournisseur", "Nom_Fournisseur", "Telephone_Fournisseur", "Commune_Origine",
        "Quantite", "Numero_Tracteur", "Chauffeur",
        "Numero_Permis", "Telephone_Chauffeur"
    ]
    if os.path.exists(FICHIER_EXCEL):
        df = pd.read_excel(FICHIER_EXCEL)
        if "Heure_Creation" not in df.columns:
            df.insert(2, "Heure_Creation", "")
        if "Utilisateur" not in df.columns:
            df.insert(3, "Utilisateur", "")
        if "Produit" not in df.columns:
            df.insert(4, "Produit", "")
        if "Code_Fournisseur" not in df.columns:
            # Insérer après Utilisateur
            idx = list(df.columns).index("Utilisateur") + 1 if "Utilisateur" in df.columns else 4
            df.insert(idx, "Code_Fournisseur", "")
        cols = [c for c in colonnes_ordonnees if c in df.columns] + [c for c in df.columns if c not in colonnes_ordonnees]
        return df[cols]
    else:
        df = pd.DataFrame(columns=colonnes_ordonnees)
        df.to_excel(FICHIER_EXCEL, index=False)
        return df

def obtenir_prochain_numero(df):
    dernier = 0
    if not df.empty and "Reference" in df.columns:
        for ref in df["Reference"].dropna():
            num = extraire_numero(ref)
            if num > dernier:
                dernier = num
    return dernier + 1 if dernier > 0 else 1

def lister_fichiers_word():
    dossiers = ["bons_generes", "."]
    fichiers_valides = []
    for dossier in dossiers:
        for pattern in (
            os.path.join(dossier, "*-BC-SOYABEAN-NAP-2026*.docx"),
            os.path.join(dossier, "*-BC-SHEA-NAP-2026*.docx"),
            os.path.join(dossier, "*-BC-COTTON-NAP-2026*.docx"),
        ):
            for f in glob.glob(pattern):
                nom = os.path.basename(f)
                if any(x in nom.lower() for x in ["template", "prototype", "modele", "model"]):
                    continue
                if re.match(r'^\d{3,}-BC-(SOYABEAN|SHEA|COTTON)-NAP-2026', nom):
                    fichiers_valides.append(f)
    return sorted(set(fichiers_valides), reverse=True)

def supprimer_reference(reference_a_supprimer):
    df = charger_donnees()
    if reference_a_supprimer not in df["Reference"].astype(str).values:
        return False, "Cette référence n'existe pas."
    df = df[df["Reference"].astype(str) != str(reference_a_supprimer)]
    df.to_excel(FICHIER_EXCEL, index=False)
    return True, f"La référence {reference_a_supprimer} a été supprimée avec succès."

# ---------- Sidebar : Formulaire ----------
st.sidebar.header("➕ Nouveau Bon de Commande")
df_actuel = charger_donnees()
prochain_num = obtenir_prochain_numero(df_actuel)
prochain_num_str = f"{prochain_num:04d}"

options_frs = charger_fournisseurs_actifs()

# Produit HORS formulaire → le suffixe se met à jour tout de suite
produit_choisi = st.sidebar.radio(
    "Produit *",
    options=list(PRODUITS_SUFFIXES.keys()),
    index=1,  # Soja par défaut
    help="Le suffixe de référence change selon le produit : "
         "SHEA → -BC-SHEA/NAP-2026 · SOJA → -BC-SOYABEAN/NAP-2026 · COTON → -BC-COTTON/NAP-2026"
)
suffixe_auto = PRODUITS_SUFFIXES[produit_choisi]
st.sidebar.info(f"Référence → `XXXX{suffixe_auto}`")

with st.sidebar.form("form_nouveau_bon", clear_on_submit=True):
    st.markdown("**Référence** *(suffixe selon le produit)*")
    col_num, col_suf = st.columns([1, 2])
    with col_num:
        numero_evolution = st.text_input(
            "N° Évolution *",
            value=prochain_num_str,
            placeholder="0297"
        )
    with col_suf:
        st.text_input("Suffixe (auto)", value=suffixe_auto, disabled=True)

    utilisateur = st.text_input("Utilisateur (Créé par) *", value=name)
    date_emission = st.date_input("Date d'émission", value=datetime.now())

    # --- Fournisseur lié au code FRS ---
    if not options_frs:
        st.error("Aucun fournisseur actif. Créez-en un dans la page Fournisseurs.")
        choix_frs = ""
    else:
        choix_frs = st.selectbox(
            "Fournisseur *",
            options=[""] + list(options_frs.keys()),
            help="Liste issue de fournisseurs.xlsx"
        )

    tel_fourn = st.text_input("Téléphone Fournisseur")
    commune = st.text_input("Commune d'origine")
    quantite = st.text_input("Quantité (ex: 50 MT)")
    tracteur = st.text_input("N° Tracteur / Remorque")
    chauffeur = st.text_input("Nom du Chauffeur")
    permis = st.text_input("N° Permis")
    tel_chauffeur = st.text_input("Téléphone Chauffeur")

    submitted = st.form_submit_button("💾 Enregistrer dans Excel", width="stretch")

    if submitted:
        num_clean = extraire_numero(numero_evolution)
        code_frs, nom_frs = ("", "")
        if choix_frs and choix_frs in options_frs:
            code_frs, nom_frs = options_frs[choix_frs]

        if num_clean == 0:
            st.error("Le numéro d'évolution est obligatoire et doit être numérique !")
        elif not utilisateur.strip():
            st.error("Le nom de l'utilisateur est obligatoire !")
        elif not code_frs or not nom_frs:
            st.error("Vous devez sélectionner un fournisseur dans la liste !")
        elif not produit_choisi:
            st.error("Vous devez sélectionner un produit !")
        else:
            reference = f"{num_clean:04d}{suffixe_auto}"
            df = charger_donnees()
            if reference in df["Reference"].astype(str).values:
                st.error(f"La référence {reference} existe déjà !")
            else:
                heure_creation = datetime.now().strftime("%H:%M:%S")
                nouvelle_ligne = {
                    "Reference": reference,
                    "Date_Emission": date_emission.strftime("%d/%m/%Y"),
                    "Heure_Creation": heure_creation,
                    "Utilisateur": utilisateur.strip(),
                    "Produit": produit_choisi,
                    "Code_Fournisseur": code_frs,
                    "Nom_Fournisseur": nom_frs,
                    "Telephone_Fournisseur": tel_fourn.strip(),
                    "Commune_Origine": commune.strip(),
                    "Quantite": quantite.strip(),
                    "Numero_Tracteur": tracteur.strip(),
                    "Chauffeur": chauffeur.strip(),
                    "Numero_Permis": permis.strip(),
                    "Telephone_Chauffeur": tel_chauffeur.strip()
                }
                df = pd.concat([df, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
                df.to_excel(FICHIER_EXCEL, index=False)
                st.success(f"✅ Bon **{reference}** — {produit_choisi} — {code_frs} {nom_frs}")
                st.rerun()

# ---------- Zone principale ----------
st.title("📄 Gestion des Bons de Commande – NAP SARL")
st.markdown("---")

st.subheader("📊 Tableau des Bons de Commande")
df = charger_donnees()
if df.empty:
    st.info("Aucun bon de commande enregistré pour le moment.")
else:
    st.dataframe(df, width="stretch", height=400)
    st.markdown("---")
    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("🔄 Actualiser le tableau", width="stretch"):
            st.rerun()

    with col2:
        regenerer = st.checkbox(
            "Régénérer aussi les anciens bons (nouveau format 1 page A4)",
            value=False,
            help="Recrée tous les fichiers Word, y compris ceux déjà générés."
        )
        if st.button("📄 Générer les nouveaux Bons Word", width="stretch", type="primary"):
            with st.spinner("Génération en cours..."):
                try:
                    from generer_bons import generer_tous_les_bons
                    generer_tous_les_bons(forcer_tout=regenerer)
                    st.success("Bons Word générés avec succès !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la génération : {e}")
                    import traceback
                    st.code(traceback.format_exc())

    with col3:
        if os.path.exists(FICHIER_EXCEL):
            with open(FICHIER_EXCEL, "rb") as f:
                st.download_button(
                    label="⬇️ Télécharger Excel",
                    data=f,
                    file_name="donnees_bons.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )

# ---------- SUPPRESSION (admin) ----------
if username == "admin":
    st.markdown("---")
    st.subheader("🗑️ Supprimer un Bon de Commande (Admin uniquement)")
    df = charger_donnees()
    if not df.empty:
        liste_references = df["Reference"].astype(str).tolist()
        reference_a_supprimer = st.selectbox(
            "Choisissez la référence à supprimer :",
            options=liste_references,
            key="select_suppression"
        )
        col_del1, col_del2 = st.columns([1, 3])
        with col_del1:
            if st.button("🗑️ Supprimer cette référence", type="primary", width="stretch"):
                success, message = supprimer_reference(reference_a_supprimer)
                if success:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)
        with col_del2:
            st.warning("⚠️ Cette action est irréversible.")
    else:
        st.info("Aucun bon à supprimer.")

# ---------- Téléchargement Word ----------
st.markdown("---")
st.subheader("📥 Télécharger les Bons de Commande (Word)")
col_refresh, _ = st.columns([1, 4])
with col_refresh:
    if st.button("🔄 Actualiser la liste des fichiers"):
        st.rerun()

fichiers_word = lister_fichiers_word()
if not fichiers_word:
    st.warning("Aucun fichier Word trouvé.")
    st.info(
        "1. Cliquez d'abord sur « Générer les nouveaux Bons Word »\n"
        "2. Puis cliquez sur « Actualiser la liste des fichiers »"
    )
else:
    st.success(f"**{len(fichiers_word)} bon(s) disponible(s) :**")
    for fichier in fichiers_word:
        col1, col2 = st.columns([5, 1])
        with col1:
            st.write(f"📄 **{os.path.basename(fichier)}**")
        with col2:
            try:
                with open(fichier, "rb") as f:
                    st.download_button(
                        label="⬇️ Télécharger",
                        data=f,
                        file_name=os.path.basename(fichier),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{fichier}",
                        width="stretch"
                    )
            except Exception as e:
                st.error(f"Erreur : {e}")

st.markdown("---")
st.caption("NAP SARL – Bons de commande multi-produits (SHEA / SOYABEAN / COTTON) + codes FRS")