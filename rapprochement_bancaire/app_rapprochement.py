"""
Rapprochement bancaire NAP SARL - relevé PDF Ecobank  vs  classeur Excel interne.
Lancer :  streamlit run app_rapprochement.py
"""
import hashlib
import io
import json
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

import moteur as m
import rapport as rp

DOSSIER = Path(__file__).parent
FICHIER_REFERENCE = DOSSIER / "dernier_controle.json"

ICONES = {m.STATUT_OK: "✅ Rapproché", m.STATUT_VERIF: "🔎 À vérifier",
          m.STATUT_ECART: "⚠️ Écart de montant", m.STATUT_ABSENT: "❌ Absent d'Excel"}

st.set_page_config(page_title="Rapprochement bancaire", page_icon="🏦", layout="wide")

def lire_reference():
    try:
        d = json.loads(FICHIER_REFERENCE.read_text(encoding="utf-8"))
        d["date_fin"] = datetime.strptime(d["date_fin"], "%Y-%m-%d").date()
        return d
    except Exception:
        return None

def ecrire_reference(synthese, ref, compte):
    FICHIER_REFERENCE.write_text(json.dumps({
        "date_fin": synthese["date_fin"].isoformat(), "solde_cloture": synthese["solde_banque_cloture"],
        "ref": ref, "compte": compte, "enregistre_le": datetime.now().isoformat(timespec="seconds")},
        ensure_ascii=False, indent=2), encoding="utf-8")

def solde_banque_avant(banque, meta, d):
    avant = banque[banque["Date"].apply(lambda x: x is not None and x < d)]
    return int(avant["Solde banque"].iloc[-1]) if len(avant) else int(meta.get("solde_ouverture", 0))

@st.cache_data(show_spinner="Lecture du relevé PDF…")
def analyser_pdf(octets):
    meta, banque = m.lire_releve_pdf(io.BytesIO(octets))
    return meta, banque, m.controler_releve(meta, banque)

@st.cache_data(show_spinner="Lecture du fichier Excel…")
def analyser_excel(octets, feuille):
    return m.lire_excel(io.BytesIO(octets), feuille)

@st.cache_data(show_spinner="Rapprochement en cours…")
def calculer(meta, banque, excel, date_depart, solde_depart, tolerance):
    return m.rapprocher(meta, banque, excel, date_depart, solde_depart, tolerance_jours=tolerance)

def f(n):
    return m.fmt_xof(n)

# ----------------------------------------------------------------- barre latérale
st.sidebar.title("🏦 Rapprochement bancaire")
st.sidebar.markdown("**1. Importer les fichiers**")
fichier_pdf = st.sidebar.file_uploader("Relevé de la banque (PDF)", type=["pdf"])
fichier_xlsx = st.sidebar.file_uploader("Fichier Excel interne (feuille STATEMENT)", type=["xlsx", "xlsm"])

with st.sidebar.expander("Options", expanded=False):
    nom_feuille = st.text_input("Nom de la feuille Excel", value="STATEMENT")
    tolerance = st.number_input("Tolérance de date (jours)", min_value=0, max_value=30, value=7,
                                help="Écart maximal entre la date banque et la date Excel pour rapprocher deux lignes.")
    etabli_par = st.text_input("Établi par", value="")
    date_controle = st.date_input("Date du contrôle", value=date.today())

# ----------------------------------------------------------------- page d'accueil
st.title("Rapprochement bancaire - relevé banque / Excel")

if not fichier_pdf or not fichier_xlsx:
    st.info("**Pour commencer :** importez à gauche le **relevé PDF** de la banque et votre **fichier Excel**. "
            "L'application lit les deux fichiers, rapproche les lignes et prépare le rapport.")
    c1, c2, c3 = st.columns(3)
    c1.markdown("### ① Importer\nLe PDF d'Ecobank et votre classeur Excel.")
    c2.markdown("### ② Vérifier\nLes lignes rapprochées, à vérifier, ou absentes d'Excel. Ajoutez vos observations.")
    c3.markdown("### ③ Exporter\nLe rapport PDF (QR code, date, solde de clôture, banque) et le relevé rapproché Excel.")
    ref_prec = lire_reference()
    if ref_prec:
        st.caption(f"Dernier contrôle enregistré : arrêté au {ref_prec['date_fin']:%d/%m/%Y}, "
                   f"solde de clôture {f(ref_prec['solde_cloture'])} XOF (point de départ proposé).")
    st.stop()

# ----------------------------------------------------------------- lecture des fichiers
try:
    octets_pdf = fichier_pdf.getvalue()
    meta, banque, controles = analyser_pdf(octets_pdf)
except Exception as err:
    st.error(f"Impossible de lire le relevé PDF : {err}")
    st.stop()
try:
    excel = analyser_excel(fichier_xlsx.getvalue(), nom_feuille)
except Exception as err:
    st.error(f"Impossible de lire le fichier Excel : {err}")
    st.stop()
if banque.empty:
    st.error("Aucune opération trouvée dans le PDF. Est-ce bien un relevé Ecobank (Statement of Account) ?")
    st.stop()

signature = hashlib.md5(octets_pdf + fichier_xlsx.getvalue()[:5000]).hexdigest()
if st.session_state.get("signature") != signature:
    st.session_state["signature"] = signature
    st.session_state["observations"] = {}

# ----------------------------------------------------------------- point de départ
ref_prec = lire_reference()
debut_pdf = meta.get("periode_debut") or banque["Date"].min()
fin_pdf = meta.get("periode_fin") or banque["Date"].max()

st.sidebar.markdown("---")
st.sidebar.markdown("**2. Point de départ du contrôle**")
if ref_prec and ref_prec["date_fin"] < fin_pdf:
    date_defaut = max(ref_prec["date_fin"] + timedelta(days=1), debut_pdf)
else:
    date_defaut = debut_pdf
date_depart = st.sidebar.date_input("Date de départ", value=date_defaut, key=f"dep_{signature}")

choix = ["Solde de la banque à cette date (automatique)", "Clôture du contrôle précédent", "Saisir un montant"]
dispo_prec = bool(ref_prec and ref_prec["date_fin"] < fin_pdf)
mode = st.sidebar.radio("Solde de départ", choix, index=1 if dispo_prec else 0, key=f"mode_{signature}")
auto = solde_banque_avant(banque, meta, date_depart)
if mode == choix[0]:
    solde_depart = auto
elif mode == choix[1]:
    if ref_prec:
        solde_depart = int(ref_prec["solde_cloture"])
        st.sidebar.caption(f"Contrôle précédent arrêté au {ref_prec['date_fin']:%d/%m/%Y}.")
    else:
        st.sidebar.warning("Aucun contrôle enregistré : solde de la banque utilisé.")
        solde_depart = auto
else:
    solde_depart = int(st.sidebar.number_input("Montant de départ (XOF)", value=int(auto), step=1, format="%d",
                                               key=f"montant_{signature}"))
st.sidebar.metric("Solde de départ retenu", f(solde_depart) + " XOF")

# ----------------------------------------------------------------- calcul
resultat = calculer(meta, banque, excel, date_depart, solde_depart, int(tolerance))
S = resultat["synthese"]
B = resultat["banque"]
E = resultat["excel"]
B_per = B[B["Date"].apply(lambda x: x is not None and S["date_depart"] <= x <= S["date_fin"])].copy()
obs = st.session_state["observations"]

# ----------------------------------------------------------------- bandeau de résultat
st.caption(f"**{meta.get('banque', '')}** · compte {meta.get('compte', '')} · {meta.get('societe', '')} · "
           f"relevé du {debut_pdf:%d/%m/%Y} au {fin_pdf:%d/%m/%Y} · {len(banque)} opérations")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Solde de clôture banque", f(S["solde_banque_cloture"]), help=f"Au {S['date_fin']:%d/%m/%Y}")
k2.metric("Solde Excel recalculé", f(S["solde_excel_cloture"]))
k3.metric("Écart à expliquer", f(S["ecart_total"]))
k4.metric("Écart non expliqué", f(S["non_explique"]))

if abs(S["non_explique"]) < 1:
    st.success("✅ Contrôle bouclé : tout l'écart entre la banque et Excel est expliqué par les lignes ci-dessous.")
else:
    st.error("❌ Il reste un écart non expliqué. Vérifiez la date et le solde de départ dans la barre de gauche.")

ok_lecture = all(ok for ok, _ in controles)
with st.expander(("✅ " if ok_lecture else "❌ ") + "Contrôle de lecture du relevé PDF", expanded=not ok_lecture):
    for ok, msg in controles:
        st.write(("✅ " if ok else "❌ ") + msg)

n1, n2, n3, n4 = st.columns(4)
n1.metric("✅ Rapprochées", S["nb_rapprochees"])
n2.metric("🔎 À vérifier", S["nb_a_verifier"])
n3.metric("⚠️ Écarts de montant", S["nb_ecarts_montant"])
n4.metric("❌ Absentes d'Excel", S["nb_banque_seule"])

onglet_releve, onglet_ecarts, onglet_synthese, onglet_export = st.tabs(
    ["📋 Relevé & observations", "⚠️ À traiter", "📊 Synthèse", "📥 Rapport & exports"])

# ----------------------------------------------------------------- onglet 1 : relevé + observations
with onglet_releve:
    st.write("Ajoutez vos **observations** dans la dernière colonne : elles seront reprises dans le rapport et dans l'Excel.")
    vue = st.radio("Afficher", ["Tout", "À traiter seulement", "Rapprochées seulement"], horizontal=True)
    base = B_per
    if vue == "À traiter seulement":
        base = B_per[B_per["Statut"] != m.STATUT_OK]
    elif vue == "Rapprochées seulement":
        base = B_per[B_per["Statut"] == m.STATUT_OK]

    tableau = pd.DataFrame({
        "N°": base["No"].values,
        "Date": base["Date"].apply(lambda d: d.strftime("%d/%m/%Y")).values,
        "Libellé": (base["Libellé propre"] + base["Bénéficiaire"].apply(lambda x: f" - {x}" if x else "")).values,
        "Chèque": base["N° chèque"].values,
        "Débit": base["Débit"].apply(lambda x: f(x) if x else "").values,
        "Crédit": base["Crédit"].apply(lambda x: f(x) if x else "").values,
        "État": base["Statut"].map(ICONES).values,
        "Ligne Excel": base["Ligne Excel liée"].apply(lambda x: "" if pd.isna(x) else str(int(x))).values,
        "Observation": [obs.get(int(no), det) for no, det in zip(base["No"], base["Détail"])],
    })
    edite = st.data_editor(
        tableau, hide_index=True, width="stretch", height=560, key=f"editeur_{vue}_{signature}",
        disabled=[c for c in tableau.columns if c != "Observation"],
        column_config={"Observation": st.column_config.TextColumn("Observation (modifiable)", width="large")})
    defauts = dict(zip(base["No"].astype(int), base["Détail"]))
    for no, texte in zip(edite["N°"], edite["Observation"]):
        no = int(no)
        texte = "" if texte is None else str(texte)
        if texte != defauts.get(no, ""):
            obs[no] = texte
        else:
            obs.pop(no, None)

# ----------------------------------------------------------------- onglet 2 : à traiter
with onglet_ecarts:
    def montrer(titre, df, colonnes, total_label=None, total=None):
        st.subheader(f"{titre} ({len(df)})")
        if df.empty:
            st.caption("Rien à signaler.")
            return
        st.dataframe(df[colonnes], hide_index=True, width="stretch")
        if total_label:
            st.caption(f"{total_label} : **{f(total)} XOF**")

    def pour_affichage(df):
        d = pd.DataFrame({
            "N°": df["No"], "Date": df["Date"].apply(lambda x: x.strftime("%d/%m/%Y")),
            "Libellé": df["Libellé propre"] + df["Bénéficiaire"].apply(lambda x: f" - {x}" if x else ""),
            "Chèque": df["N° chèque"], "Débit": df["Débit"].apply(lambda x: f(x) if x else ""),
            "Crédit": df["Crédit"].apply(lambda x: f(x) if x else ""),
            "Ligne Excel": df["Ligne Excel liée"].apply(lambda x: "" if pd.isna(x) else str(int(x))),
            "Observation": [obs.get(int(n), dt) for n, dt in zip(df["No"], df["Détail"])]})
        return d

    cols = ["N°", "Date", "Libellé", "Chèque", "Débit", "Crédit", "Observation"]
    absentes = B_per[B_per["Statut"] == m.STATUT_ABSENT]
    montrer("❌ En banque, absentes d'Excel - à saisir dans Excel", pour_affichage(absentes), cols,
            "Effet net sur le solde", S["net_banque_seule"])
    montrer("🔎 Rapprochements à vérifier", pour_affichage(B_per[B_per["Statut"] == m.STATUT_VERIF]),
            cols[:-1] + ["Ligne Excel", "Observation"])
    montrer("⚠️ Écarts de montant", pour_affichage(B_per[B_per["Statut"] == m.STATUT_ECART]),
            cols[:-1] + ["Ligne Excel", "Observation"], "Écart cumulé", S["net_ecarts_montant"])

    seules = E[E["En période"] & (E["Statut"] == m.STATUT_EXCEL_SEUL)]
    st.subheader(f"⏳ Dans Excel, pas encore en banque ({len(seules)})")
    if seules.empty:
        st.caption("Rien à signaler.")
    else:
        aff = pd.DataFrame({"Ligne": seules["Ligne Excel"], "Date": seules["Date"].apply(lambda x: x.strftime("%d/%m/%Y")),
                            "Description": seules["Description"], "Chèque": seules["N° chèque"],
                            "Débit": seules["Débit"].apply(lambda x: f(x) if x else ""),
                            "Crédit": seules["Crédit"].apply(lambda x: f(x) if x else ""), "Motif": seules["Motif"]})
        st.dataframe(aff, hide_index=True, width="stretch")
        st.caption("Chèques non encore encaissés ou virements attendus. Effet net sur le solde Excel : "
                   f"**{f(-S['net_excel_seul'])} XOF**")

# ----------------------------------------------------------------- onglet 3 : synthèse
with onglet_synthese:
    st.subheader("Rapprochement")
    tab = pd.DataFrame({
        "": ["Solde de départ", "+ Crédits de la période", "- Débits de la période", "= Solde au " + f"{S['date_fin']:%d/%m/%Y}"],
        "Banque": [f(S["solde_depart"]), f(S["credits_banque"]), f(S["debits_banque"]), f(S["solde_banque_cloture"])],
        "Excel": [f(S["solde_depart"]), f(S["credits_excel"]), f(S["debits_excel"]), f(S["solde_excel_cloture"])]})
    st.dataframe(tab, hide_index=True, width="stretch")
    st.subheader("Explication de l'écart")
    expl = pd.DataFrame({
        "Poste": ["Écart à expliquer (banque - Excel)", f"Absentes d'Excel ({S['nb_banque_seule']})",
                  f"Excel pas encore en banque ({S['nb_excel_seul']})", f"Écarts de montant ({S['nb_ecarts_montant']})",
                  f"Écritures Excel d'avant la période ({S['nb_hors_periode']})", "Écart sur le solde de départ",
                  "TOTAL EXPLIQUÉ", "ÉCART NON EXPLIQUÉ"],
        "Montant (XOF)": [f(S["ecart_total"]), f(S["net_banque_seule"]), f(-S["net_excel_seul"]), f(S["net_ecarts_montant"]),
                          f(S["net_hors_periode"]), f(S["ecart_ouverture"]), f(S["ecart_explique"]), f(S["non_explique"])]})
    st.dataframe(expl, hide_index=True, width="stretch")
    st.caption(f"Pour information, la colonne SOLDE de votre Excel affiche {f(S['solde_excel_colonne'])} XOF au "
               f"{S['date_fin']:%d/%m/%Y} (elle inclut tout l'historique, pas seulement la période contrôlée).")

# ----------------------------------------------------------------- onglet 4 : exports
with onglet_export:
    remarque = st.text_area("Observations générales (reprises dans le rapport)", height=110,
                            placeholder="Ex. : 3 chèques émis le 14/10 non encore encaissés ; virement Bunge reçu avec écart de change.")
    pdf_octets, reference = rp.generer_rapport_pdf(
        meta, S, B, E, observations=obs, etabli_par=etabli_par, date_controle=date_controle,
        remarque_generale=remarque, controles_releve=controles)
    xlsx_octets = rp.generer_excel(meta, S, B, E, observations=obs)

    c1, c2 = st.columns(2)
    c1.download_button("📄 Télécharger le rapport de contrôle (PDF)", pdf_octets,
                       file_name=f"Rapport_controle_{S['date_fin']:%Y%m%d}.pdf", mime="application/pdf",
                       type="primary", width="stretch")
    c2.download_button("📊 Télécharger le relevé rapproché (Excel)", xlsx_octets,
                       file_name=f"Releve_rapproche_{S['date_fin']:%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    st.caption("L'Excel contient 4 feuilles : Synthèse, Relevé rapproché, **À saisir dans Excel** (lignes prêtes à coller "
               "dans votre feuille STATEMENT) et Pas encore en banque.")

    st.markdown("---")
    st.markdown(f"**Référence du contrôle : `{reference}`**")
    try:
        from reportlab.graphics import renderSVG
        svg = renderSVG.drawToString(rp._qr(rp.texte_qr(meta, S, reference, date_controle), 40 * 2.835))
        st.image(svg, width=150)
    except Exception:
        pass
    with st.expander("Contenu du QR code"):
        st.code(rp.texte_qr(meta, S, reference, date_controle))

    st.markdown("---")
    st.write("Quand le contrôle est terminé, enregistrez-le : la prochaine fois, l'application proposera "
             f"automatiquement le **{S['date_fin'] + timedelta(days=1):%d/%m/%Y}** et le solde de **{f(S['solde_banque_cloture'])} XOF** "
             "comme point de départ.")
    if st.button("💾 Enregistrer ce contrôle comme point de départ du prochain", width="stretch"):
        try:
            ecrire_reference(S, reference, meta.get("compte", ""))
            st.success("Contrôle enregistré.")
        except Exception as err:
            st.error(f"Impossible d'enregistrer : {err}")