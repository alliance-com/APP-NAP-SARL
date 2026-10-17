import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime, timedelta

st.set_page_config(page_title="Tableau de Bord", page_icon="📊", layout="wide")

# ---------- Auth (si multipage) ----------
if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

FICHIER_EXCEL = "donnees_bons.xlsx"


def extraire_mt(val):
    """Extrait un nombre (MT) depuis '50 MT', '50', 50, etc."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else 0.0


def parser_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return pd.NaT
    if isinstance(val, datetime):
        return pd.Timestamp(val)
    if isinstance(val, pd.Timestamp):
        return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except Exception:
            continue
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def charger_bons():
    if not os.path.exists(FICHIER_EXCEL):
        return pd.DataFrame()
    df = pd.read_excel(FICHIER_EXCEL)
    if df.empty:
        return df
    # Normalisation
    if "Date_Emission" in df.columns:
        df["Date_Emission"] = df["Date_Emission"].apply(parser_date)
    else:
        df["Date_Emission"] = pd.NaT
    if "Quantite" in df.columns:
        df["Masse_MT"] = df["Quantite"].apply(extraire_mt)
    else:
        df["Masse_MT"] = 0.0
    if "Nom_Fournisseur" not in df.columns:
        df["Nom_Fournisseur"] = ""
    if "Reference" not in df.columns:
        df["Reference"] = ""
    if "Produit" not in df.columns:
        df["Produit"] = ""
    df["Nom_Fournisseur"] = df["Nom_Fournisseur"].fillna("").astype(str)
    df["Produit"] = df["Produit"].fillna("").astype(str)
    return df


# ---------- UI ----------
st.title("📊 Tableau de Bord — Bons de Commande")
st.caption("Source : donnees_bons.xlsx  ·  Actualisez après chaque nouvel enregistrement")

df = charger_bons()

if df.empty:
    st.info("Aucun bon enregistré pour le moment. Créez des bons dans la page **Bons de Commande**.")
    st.stop()

# Filtres
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    produits = ["Tous"] + sorted([p for p in df["Produit"].unique() if p and p.lower() not in ("nan", "none")])
    filtre_prod = st.selectbox("Produit", options=produits)
with col_f2:
    dmin = df["Date_Emission"].min()
    dmax = df["Date_Emission"].max()
    if pd.notna(dmin) and pd.notna(dmax):
        plage = st.date_input(
            "Période",
            value=(dmin.date(), dmax.date()),
            min_value=dmin.date(),
            max_value=max(dmax.date(), datetime.now().date()),
        )
    else:
        plage = None
with col_f3:
    st.write("")
    st.write("")
    if st.button("🔄 Actualiser"):
        st.rerun()

df_f = df.copy()
if filtre_prod != "Tous":
    df_f = df_f[df_f["Produit"] == filtre_prod]
if plage and isinstance(plage, (list, tuple)) and len(plage) == 2:
    deb, fin = plage
    df_f = df_f[
        (df_f["Date_Emission"] >= pd.Timestamp(deb))
        & (df_f["Date_Emission"] <= pd.Timestamp(fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))
    ]

# ---------- KPI ----------
nb_bons = len(df_f)
masse_totale = df_f["Masse_MT"].sum()
nb_frs = df_f["Nom_Fournisseur"].replace("", pd.NA).dropna().nunique()
aujourdhui = pd.Timestamp(datetime.now().date())
nb_aujourdhui = len(df_f[df_f["Date_Emission"].dt.normalize() == aujourdhui]) if "Date_Emission" in df_f else 0
masse_aujourdhui = df_f.loc[df_f["Date_Emission"].dt.normalize() == aujourdhui, "Masse_MT"].sum() if nb_bons else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Bons de commande émis", f"{nb_bons}")
k2.metric("Masse totale (MT)", f"{masse_totale:,.1f}".replace(",", " "))
k3.metric("Fournisseurs distincts", f"{nb_frs}")
k4.metric("Bons aujourd'hui", f"{nb_aujourdhui}", delta=f"{masse_aujourdhui:.1f} MT")

st.markdown("---")

# ---------- Par fournisseur ----------
st.subheader("Nombre de bons et masse par fournisseur")
par_frs = (
    df_f[df_f["Nom_Fournisseur"].str.strip() != ""]
    .groupby("Nom_Fournisseur", as_index=False)
    .agg(Nb_Bons=("Reference", "count"), Masse_MT=("Masse_MT", "sum"))
    .sort_values("Nb_Bons", ascending=False)
)
if par_frs.empty:
    st.warning("Aucun fournisseur renseigné sur la période.")
else:
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.dataframe(
            par_frs.rename(columns={"Nom_Fournisseur": "Fournisseur"}),
            width="stretch",
            hide_index=True,
            height=min(400, 40 + 35 * len(par_frs)),
        )
    with c2:
        st.bar_chart(par_frs.set_index("Nom_Fournisseur")["Nb_Bons"], horizontal=True)

st.markdown("---")

# ---------- Jour / Semaine / Mois ----------
st.subheader("Quantités (MT) — Jour · Semaine · Mois")

df_dates = df_f.dropna(subset=["Date_Emission"]).copy()
if df_dates.empty:
    st.warning("Pas de dates d'émission exploitables.")
else:
    # Jour
    par_jour = (
        df_dates.assign(Jour=df_dates["Date_Emission"].dt.date)
        .groupby("Jour", as_index=False)
        .agg(Nb_Bons=("Reference", "count"), Masse_MT=("Masse_MT", "sum"))
        .sort_values("Jour")
    )

    # Semaine (lundi de la semaine)
    par_semaine = (
        df_dates.assign(Semaine=df_dates["Date_Emission"].dt.to_period("W-SUN").apply(lambda p: p.start_time.date()))
        .groupby("Semaine", as_index=False)
        .agg(Nb_Bons=("Reference", "count"), Masse_MT=("Masse_MT", "sum"))
        .sort_values("Semaine")
    )

    # Mois
    par_mois = (
        df_dates.assign(Mois=df_dates["Date_Emission"].dt.to_period("M").astype(str))
        .groupby("Mois", as_index=False)
        .agg(Nb_Bons=("Reference", "count"), Masse_MT=("Masse_MT", "sum"))
        .sort_values("Mois")
    )

    tab1, tab2, tab3 = st.tabs(["📅 Par jour", "📆 Par semaine", "🗓️ Par mois"])

    with tab1:
        t1, t2 = st.columns([1, 1.2])
        with t1:
            st.dataframe(par_jour, width="stretch", hide_index=True, height=320)
        with t2:
            if not par_jour.empty:
                chart_j = par_jour.set_index("Jour")[["Masse_MT", "Nb_Bons"]]
                st.bar_chart(chart_j)

    with tab2:
        t1, t2 = st.columns([1, 1.2])
        with t1:
            st.dataframe(
                par_semaine.rename(columns={"Semaine": "Début semaine"}),
                width="stretch",
                hide_index=True,
                height=320,
            )
        with t2:
            if not par_semaine.empty:
                st.bar_chart(par_semaine.set_index("Semaine")[["Masse_MT", "Nb_Bons"]])

    with tab3:
        t1, t2 = st.columns([1, 1.2])
        with t1:
            st.dataframe(par_mois, width="stretch", hide_index=True, height=320)
        with t2:
            if not par_mois.empty:
                st.bar_chart(par_mois.set_index("Mois")[["Masse_MT", "Nb_Bons"]])

st.markdown("---")
st.caption("NAP SARL — Tableau de bord bons de commande (masse en MT extraite du champ Quantité)")
