import streamlit as st
import pandas as pd
import requests
import glob
import numpy as np
import math

st.set_page_config(page_title="Import Excel vers Supabase", layout="wide")
st.title("📥 Import Excel vers Supabase")

try:
    URL = st.secrets["supabase"]["url"]
    KEY = st.secrets["supabase"]["key"]
except:
    st.error("Tu n'as pas mis tes clés dans Secrets.")
    st.stop()

st.success("Connecté à Supabase")

headers = {
    "apikey": KEY,
    "Authorization": f"Bearer {KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

fichiers = glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
if not fichiers:
    st.warning("Aucun Excel trouvé dans GitHub.")
    st.stop()

st.subheader("1. Choisis ton fichier Excel")
fichier = st.selectbox("Fichier", fichiers)

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

def nettoyer_colonne(c):
    c = str(c).strip().lower()
    c = c.replace(" ", "_").replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c").replace("°","").replace("n°","no").replace("(","").replace(")","")
    return c

df.columns = [nettoyer_colonne(c) for c in df.columns]
st.write("Colonnes nettoyées :", list(df.columns))

st.subheader("3. Choisis la table Supabase")
table = st.text_input("Nom de la table dans Supabase", value="paiements")

# C'est ici la correction du NaN
def nettoyer_valeur(v):
    # Case vide
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except:
        pass
    # Nombre infini / NaN
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    # Date
    if isinstance(v, pd.Timestamp):
        return v.isoformat()
    # Texte "nan" / "None" / "nat"
    s = str(v).strip()
    if s.lower() in ["nan", "none", "nat", ""]:
        return None
    return v

if st.button("🚀 IMPORTER TOUT DANS SUPABASE"):
    with st.spinner("En train d'importer..."):
        records = []
        for _, ligne in df.iterrows():
            d = {}
            for col in df.columns:
                d[col] = nettoyer_valeur(ligne[col])
            records.append(d)

        total = len(records)
        ok = 0

        for i in range(0, total, 50):
            paquet = records[i:i+50]
            try:
                r = requests.post(f"{URL}/rest/v1/{table}", headers=headers, json=paquet)
            except Exception as e:
                st.error(f"Erreur envoi : {e}")
                break
            if r.status_code in [200, 201]:
                ok += len(paquet)
                st.write(f"{ok}/{total} lignes OK...")
            else:
                st.error(f"Erreur Supabase : {r.text[:500]}")
                break

        if ok == total:
            st.success(f"✅ TERMINÉ : {ok} lignes importées dans {table}")
            st.balloons()
        else:
            st.warning(f"{ok}/{total} lignes importées. Regarde l'erreur rouge.")
