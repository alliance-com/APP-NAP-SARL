import streamlit as st
import pandas as pd
import requests
import os
import glob

st.set_page_config(page_title="Import Excel vers Supabase", layout="wide")
st.title("📥 Import Excel vers Supabase")

# On prend tes clés que tu as mises dans Secrets
try:
    URL = st.secrets["supabase"]["url"]
    KEY = st.secrets["supabase"]["key"]
except:
    st.error("Tu n'as pas mis tes clés dans Secrets. Va dans Streamlit > Manage App > Settings > Secrets")
    st.stop()

st.success("Connecté à Supabase")
st.write(URL)

headers = {
    "apikey": KEY,
    "Authorization": f"Bearer {KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

# On cherche tous les Excel dans ton projet
fichiers = glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
if not fichiers:
    st.warning("Aucun Excel trouvé dans GitHub. Il faut que ton Excel soit dans GitHub.")
    st.stop()

st.subheader("1. Choisis ton fichier Excel")
fichier = st.selectbox("Fichier", fichiers)

# On lit les feuilles
try:
    excel = pd.ExcelFile(fichier)
    feuille = st.selectbox("Feuille", excel.sheet_names)
    df = pd.read_excel(fichier, sheet_name=feuille)
    df = df.dropna(how="all")
except Exception as e:
    st.error(f"Erreur lecture Excel : {e}")
    st.stop()

st.subheader("2. Aperçu")
st.write(f"{len(df)} lignes trouvées")
st.dataframe(df.head(10), use_container_width=True)
st.write("Colonnes :", list(df.columns))

# On nettoie les colonnes pour Supabase
# Supabase n'aime pas les espaces, les accents
def nettoyer_colonne(c):
    c = str(c).strip().lower()
    c = c.replace(" ", "_").replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c").replace("°","").replace("n°","no")
    return c

df.columns = [nettoyer_colonne(c) for c in df.columns]
st.write("Colonnes nettoyées pour Supabase :", list(df.columns))

st.subheader("3. Choisis la table Supabase")
table = st.text_input("Nom de la table dans Supabase", value="paiements")
st.info("Exemples : paiements / depenses / prestataires. Tu dois avoir créé cette table avant dans Supabase > SQL Editor, sinon ça va planter.")

if st.button("🚀 IMPORTER TOUT DANS SUPABASE"):
    with st.spinner("En train d'importer..."):
        # On remplace les NaN par None
        df = df.where(pd.notnull(df), None)
        # On convertit les dates en texte
        for col in df.columns:
            df[col] = df[col].astype(str).replace("None", None).replace("nan", None)

        records = df.to_dict(orient="records")
        total = len(records)
        ok = 0
        erreurs = 0

        # On envoie par paquet de 100
        for i in range(0, total, 100):
            paquet = records[i:i+100]
            r = requests.post(f"{URL}/rest/v1/{table}", headers=headers, json=paquet)
            if r.status_code in [200, 201]:
                ok += len(paquet)
            else:
                erreurs += len(paquet)
                st.error(f"Erreur paquet {i} : {r.text[:300]}")
                break

        if erreurs == 0:
            st.success(f"✅ TERMINÉ : {ok} lignes importées dans {table}")
            st.balloons()
        else:
            st.warning(f"{ok} lignes OK, {erreurs} lignes en erreur. Regarde l'erreur rouge au-dessus.")
