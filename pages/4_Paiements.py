import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import re
from io import BytesIO

st.set_page_config(page_title="Paiements", page_icon="💳", layout="wide")

if "authentication_status" in st.session_state and not st.session_state.get("authentication_status"):
    st.warning("Veuillez vous connecter depuis la page d'accueil.")
    st.stop()

name = st.session_state.get("name") or "Utilisateur"
username = st.session_state.get("username") or ""

# ==================== CONFIG ====================
FICHIER_PAIEMENTS = "Suivi de paiement.xlsx"
FEUILLE_DEPENSES = "DEPENSES"
DOSSIER_BONS = "bons_paiement"
os.makedirs(DOSSIER_BONS, exist_ok=True)

TYPES_OP = ["PURCHASE", "EXPORT", "SERVICE", "SALARY", "AUTRE"]
PRODUITS = ["SHEA", "SOYABEAN", "COTTON", "FRET EXPORT", "TRANSFERT DHL", "SERVICE", "ENTRETIEN", "SALARY", "BFU", "AUTRE"]
MODES = ["CHEQUE", "ESPECES", "VIREMENT", "MOBILE MONEY"]
BANQUES = ["ECOBANK BENIN", "CAISSE", "BIIC", "ATLANTIC BANK BENIN", "BGFI", "AUTRE"]
OBSERVATIONS = ["PAYE", "AVANCE", "PARTIEL", "EN ATTENTE"]

COLONNES = [
    "Date", "No de transaction", "Type d'operation", "Produit",
    "Prestataires/Fournisseurs", "Nom du bénéficiaire", "Num. Pieces d'identite",
    "Date d'expiration", "Addresse /Telephone", "Ref doc",
    "Description/Objet du paiement", "Montant", "Observation",
    "Montant en lettre", "Mode paiement", "Numero de cheque", "Banque"
]
# =================================================



def valeurs_uniques(df, colonne, extras=None):
    """Liste déroulante = valeurs déjà saisies dans Excel + extras de base."""
    vals = []
    if extras:
        vals.extend(extras)
    if df is not None and not df.empty and colonne in df.columns:
        for v in df[colonne].dropna().astype(str):
            v = v.strip()
            if v and v.lower() not in ("nan", "none", "-"):
                vals.append(v)
    # unique preserve order
    seen = set()
    out = []
    for v in vals:
        key = v.casefold()
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out


def choix_ou_nouveau(label, options, key_prefix):
    """Liste déroulante (valeurs Excel) + champ libre pour une NOUVELLE valeur.
    Si le champ libre est rempli, il a priorité sur la liste.
    """
    options = list(options) if options else []
    # Ligne vide en tête pour forcer un choix explicite
    options_affichees = [""] + options
    choix = st.selectbox(
        label,
        options=options_affichees,
        key=f"{key_prefix}_sel",
        format_func=lambda x: "— Choisir —" if x == "" else x,
    )
    nouveau = st.text_input(
        f"✏️ Nouvelle valeur (si absente de la liste)",
        key=f"{key_prefix}_new",
        placeholder="Écrire ici pour ajouter une valeur qui n'est pas dans la liste",
        help="Laissez vide pour utiliser la liste. Si vous tapez ici, cette valeur sera utilisée.",
    )
    nouveau = (nouveau or "").strip()
    if nouveau:
        return nouveau
    return (choix or "").strip()

# ==================== ANNUAIRE (listes déroulantes liées à l'Excel) ====================
COL_PRESTA = "Prestataires/Fournisseurs"
COL_BENEF = "Nom du bénéficiaire"
COL_PIECE = "Num. Pieces d'identite"
COL_EXP = "Date d'expiration"
COL_TEL = "Addresse /Telephone"


def _propre(v):
    """Valeur Excel -> texte propre ('' si vide ou placeholder '-')."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = str(v).strip()
    if t.lower() in ("", "nan", "none", "nat", "-"):
        return ""
    # date texte "2031-06-02 00:00:00" -> "02/06/2031"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return t


def _uniques_recents(serie):
    """Valeurs uniques non vides, la plus récente (dernière ligne Excel) en premier."""
    out, vus = [], set()
    for v in reversed(serie.tolist()):
        t = _propre(v)
        if t and t.casefold() not in vus:
            vus.add(t.casefold())
            out.append(t)
    return out


def charger_annuaire(df):
    """Construit l'annuaire depuis la feuille DEPENSES (base de données).
    Retourne {prestataire: DataFrame de ses lignes}, trié par nom."""
    if df is None or df.empty or COL_PRESTA not in df.columns:
        return {}
    annuaire = {}
    tmp = df.copy()
    tmp["_presta"] = tmp[COL_PRESTA].map(_propre)
    tmp = tmp[tmp["_presta"] != ""]
    for nom, sous in tmp.groupby("_presta", sort=True):
        annuaire[nom] = sous
    return annuaire


def extraire_num_trans(val):
    if not val:
        return 0
    m = re.search(r"(\d+)", str(val).replace(" ", ""))
    return int(m.group(1)) if m else 0


def prochain_numero(df, prefix="CHQ"):
    """CHQ-0390-2026 ou CC-0138-2026 selon le mode."""
    annee = datetime.now().year
    nums = []
    if not df.empty and "No de transaction" in df.columns:
        for v in df["No de transaction"].dropna().astype(str):
            if prefix in v.upper() or (prefix == "CHQ" and "CHQ" in v.upper()) or (prefix == "CC" and v.upper().startswith("CC")):
                n = extraire_num_trans(v)
                if n:
                    nums.append(n)
    n = max(nums) + 1 if nums else 1
    if prefix == "CHQ":
        return f"CHQ-{n:04d}-{annee}"
    return f"CC-{n:04d}-{annee}"


def montant_en_lettres(n):
    """Conversion basique FCFA en lettres (français)."""
    try:
        n = int(round(float(n)))
    except Exception:
        return ""
    if n == 0:
        return "Zéro franc CFA"

    uns = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf",
           "dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
    diz = ["", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt"]

    def sous_mille(x):
        if x < 20:
            return uns[x]
        if x < 100:
            d, u = divmod(x, 10)
            if d == 7 or d == 9:  # 70-79, 90-99
                base = 60 if d == 7 else 80
                reste = x - base
                liaison = "-" if reste else ""
                prefix = "soixante" if d == 7 else "quatre-vingt"
                if d == 9:
                    return f"quatre-vingt-{uns[10 + (x - 90)]}" if x >= 90 else prefix
                if reste < 20:
                    return f"{prefix}-{uns[reste]}" if reste else prefix
            if d == 8:
                return f"quatre-vingt-{uns[u]}" if u else "quatre-vingts"
            sep = " et " if u == 1 and d in (2, 3, 4, 5, 6) else ("-" if u else "")
            return f"{diz[d]}{sep}{uns[u]}" if u else diz[d]
        c = x // 100
        r = x % 100
        cent = "cent" if c == 1 else f"{uns[c]} cent"
        if r == 0 and c > 1:
            cent = f"{uns[c]} cents"
        return f"{cent} {sous_mille(r)}".strip() if r else cent

    parts = []
    millions = n // 1_000_000
    reste = n % 1_000_000
    milliers = reste // 1000
    unites = reste % 1000

    if millions:
        if millions == 1:
            parts.append("un million")
        else:
            parts.append(f"{sous_mille(millions)} millions")
    if milliers:
        if milliers == 1:
            parts.append("mille")
        else:
            parts.append(f"{sous_mille(milliers)} mille")
    if unites:
        parts.append(sous_mille(unites))

    texte = " ".join(parts).strip()
    if texte:
        texte = texte[0].upper() + texte[1:] + " francs CFA"
    return texte


def charger_depenses():
    """Charge la feuille DEPENSES.
    Dans le fichier NAP, les en-têtes sont en ligne 3 (header=2).
    """
    if not os.path.exists(FICHIER_PAIEMENTS):
        return pd.DataFrame(columns=COLONNES)
    try:
        xl = pd.ExcelFile(FICHIER_PAIEMENTS)
        sheet = FEUILLE_DEPENSES if FEUILLE_DEPENSES in xl.sheet_names else xl.sheet_names[0]

        # Essayer plusieurs lignes d'en-tête (fichier NAP = ligne 3)
        df = None
        for hdr in (2, 1, 0, 3):
            try:
                tmp = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=hdr)
                cols_str = " ".join(str(c).lower() for c in tmp.columns)
                if "prestataire" in cols_str or "transaction" in cols_str or "bénéficiaire" in cols_str or "beneficiaire" in cols_str:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            df = pd.read_excel(FICHIER_PAIEMENTS, sheet_name=sheet, header=2)

        # Supprimer colonnes vides / Unnamed inutiles
        df = df.dropna(axis=1, how="all")
        df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]

        # Mapping flexible des noms de colonnes
        alias = {
            "date": "Date",
            "no de transaction": "No de transaction",
            "n° de transaction": "No de transaction",
            "numero de transaction": "No de transaction",
            "type d'operation": "Type d'operation",
            "type d'opération": "Type d'operation",
            "type operation": "Type d'operation",
            "produit": "Produit",
            "prestataires/fournisseurs": "Prestataires/Fournisseurs",
            "prestataire": "Prestataires/Fournisseurs",
            "fournisseur": "Prestataires/Fournisseurs",
            "nom du bénéficiaire": "Nom du bénéficiaire",
            "nom du beneficiaire": "Nom du bénéficiaire",
            "bénéficiaire": "Nom du bénéficiaire",
            "beneficiaire": "Nom du bénéficiaire",
            "nom du percepteur": "Nom du bénéficiaire",
            "percepteur": "Nom du bénéficiaire",
            "num. pieces d'identite": "Num. Pieces d'identite",
            "num pieces d'identite": "Num. Pieces d'identite",
            "piece d'identite": "Num. Pieces d'identite",
            "date d'expiration": "Date d'expiration",
            "addresse /telephone": "Addresse /Telephone",
            "adresse /telephone": "Addresse /Telephone",
            "telephone": "Addresse /Telephone",
            "ref doc": "Ref doc",
            "description/objet du paiement": "Description/Objet du paiement",
            "description": "Description/Objet du paiement",
            "objet du paiement": "Description/Objet du paiement",
            "montant": "Montant",
            "observation": "Observation",
            "montant en lettre": "Montant en lettre",
            "mode paiement": "Mode paiement",
            "numero de cheque": "Numero de cheque",
            "banque": "Banque",
        }
        rename = {}
        for c in df.columns:
            key = str(c).strip().lower().replace("  ", " ")
            if key in alias:
                rename[c] = alias[key]
            else:
                for cible in COLONNES:
                    if key == cible.lower():
                        rename[c] = cible
                        break
        df = df.rename(columns=rename)

        for c in COLONNES:
            if c not in df.columns:
                df[c] = ""

        # Enlever lignes totalement vides (sans transaction ni montant)
        if "No de transaction" in df.columns:
            df = df[df["No de transaction"].notna()]
            df = df[df["No de transaction"].astype(str).str.strip().str.lower().isin(["", "nan", "none"]) == False]

        return df[COLONNES].reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur lecture Excel : {e}")
        return pd.DataFrame(columns=COLONNES)


def sauvegarder_ligne(nouvelle):
    """Ajoute une ligne à la fin du tableau de la feuille DEPENSES.
    - repère automatiquement la ligne d'en-têtes (ligne 3, colonne B dans le fichier NAP)
    - écrit sous la dernière opération, colonne par colonne, en recopiant le format
    - agrandit le tableau Excel (Table1) pour inclure la nouvelle ligne
    Les autres feuilles ne sont pas modifiées."""
    from copy import copy
    from openpyxl import load_workbook, Workbook

    if not os.path.exists(FICHIER_PAIEMENTS):
        wb = Workbook()
        ws = wb.active
        ws.title = FEUILLE_DEPENSES
        ws.append(COLONNES)
        wb.save(FICHIER_PAIEMENTS)

    wb = load_workbook(FICHIER_PAIEMENTS)
    ws = wb[FEUILLE_DEPENSES] if FEUILLE_DEPENSES in wb.sheetnames else wb.create_sheet(FEUILLE_DEPENSES)

    def norm(x):
        return str(x or "").strip().lower()

    # 1) ligne d'en-têtes : celle qui contient « No de transaction »
    hdr_row = None
    for r in range(1, 15):
        if any(norm(c.value) == "no de transaction" for c in ws[r]):
            hdr_row = r
            break
    if hdr_row is None:  # feuille vide : on crée les en-têtes en ligne 1
        for i, h in enumerate(COLONNES, 1):
            ws.cell(1, i, h)
        hdr_row = 1

    col_of = {norm(c.value): c.column for c in ws[hdr_row] if c.value}
    col_trans = col_of.get("no de transaction", 1)

    # 2) dernière ligne réellement remplie (colonne N° transaction)
    last = hdr_row
    for r in range(ws.max_row, hdr_row, -1):
        if ws.cell(r, col_trans).value not in (None, ""):
            last = r
            break
    new_row = last + 1

    # 3) écriture colonne par colonne + format de la ligne précédente
    valeurs = {norm(k): v for k, v in nouvelle.items()}
    for h, col in col_of.items():
        val = valeurs.get(h, None)
        if h == "date" and isinstance(val, str):
            try:
                val = datetime.strptime(val, "%d/%m/%Y")
            except ValueError:
                pass
        cell = ws.cell(new_row, col, val if val != "" else None)
        if last > hdr_row:
            src = ws.cell(last, col)
            if src.has_style:
                cell._style = copy(src._style)

    # 4) agrandir le tableau Excel pour inclure la nouvelle ligne
    for t in ws.tables.values():
        t.ref = f"{t.ref.split(':')[0]}:{t.ref.split(':')[1].rstrip('0123456789')}{new_row}"

    wb.save(FICHIER_PAIEMENTS)


# ---------- Sidebar formulaire ----------
st.sidebar.header("➕ Nouveau paiement")

df = charger_depenses()

# Listes issues de l'Excel (historique) + valeurs de base
opt_types = valeurs_uniques(df, "Type d'operation", TYPES_OP)
opt_produits = valeurs_uniques(df, "Produit", PRODUITS)
opt_prestataires = valeurs_uniques(df, "Prestataires/Fournisseurs")
opt_beneficiaires = valeurs_uniques(df, "Nom du bénéficiaire")
opt_descriptions = valeurs_uniques(df, "Description/Objet du paiement")
opt_banques = valeurs_uniques(df, "Banque", BANQUES)
opt_modes = valeurs_uniques(df, "Mode paiement", MODES)

# Champs en listes déroulantes HORS formulaire → "➕ Nouveau…" affiche tout de suite le champ texte
st.sidebar.markdown("**Listes Excel** — pour une valeur absente, écrivez-la dans le champ *Nouvelle valeur*")
type_op = choix_ou_nouveau("Type d'opération *", opt_types, "type_op")
produit = choix_ou_nouveau("Produit *", opt_produits, "produit")
prestataire = choix_ou_nouveau("Prestataire / Fournisseur *", opt_prestataires, "prestataire")
beneficiaire = choix_ou_nouveau("Nom du percepteur / bénéficiaire *", opt_beneficiaires, "benef")
description = choix_ou_nouveau("Description / Objet *", opt_descriptions, "desc")

mode_defaut = st.sidebar.selectbox("Mode (pour N° auto)", opt_modes if opt_modes else MODES, index=0)
prefix = "CHQ" if "CHEQUE" in str(mode_defaut).upper() else "CC"
num_suggere = prochain_numero(df, prefix=prefix)

with st.sidebar.form("form_paiement", clear_on_submit=True):
    date_p = st.date_input("Date *", value=date.today())
    no_trans = st.text_input("N° transaction *", value=num_suggere)
    piece_id = st.text_input("N° pièce d'identité")
    date_exp = st.text_input("Date expiration pièce", placeholder="JJ/MM/AAAA")
    telephone = st.text_input("Adresse / Téléphone")
    ref_doc = st.text_input("Réf. Doc", placeholder="Camion_MT / facture / waybill…")
    montant = st.number_input("Montant (FCFA) *", min_value=0.0, step=100.0, format="%.0f")
    observation = st.selectbox("Observation", OBSERVATIONS)
    mode_pay = st.selectbox("Mode paiement *", opt_modes if opt_modes else MODES)
    num_cheque = st.text_input("N° chèque", placeholder="si chèque")
    banque = st.selectbox("Banque", opt_banques if opt_banques else BANQUES)

    submitted = st.form_submit_button("💾 Enregistrer + générer le bon", width="stretch")

    if submitted:
        if not no_trans.strip() or not prestataire.strip() or not beneficiaire.strip():
            st.error("N° transaction, prestataire et bénéficiaire obligatoires.")
        elif montant <= 0:
            st.error("Montant invalide.")
        elif not description.strip():
            st.error("Description obligatoire.")
        else:
            lettres = montant_en_lettres(montant)
            nouvelle = {
                "Date": date_p.strftime("%d/%m/%Y"),
                "No de transaction": no_trans.strip(),
                "Type d'operation": type_op,
                "Produit": produit,
                "Prestataires/Fournisseurs": prestataire.strip(),
                "Nom du bénéficiaire": beneficiaire.strip(),
                "Num. Pieces d'identite": piece_id.strip(),
                "Date d'expiration": date_exp.strip(),
                "Addresse /Telephone": telephone.strip(),
                "Ref doc": ref_doc.strip(),
                "Description/Objet du paiement": description.strip(),
                "Montant": montant,
                "Observation": observation,
                "Montant en lettre": lettres,
                "Mode paiement": mode_pay,
                "Numero de cheque": num_cheque.strip() or "-",
                "Banque": banque,
            }
            try:
                sauvegarder_ligne(nouvelle)
            except Exception as e:
                st.error(f"Erreur sauvegarde Excel : {e}")
            else:
                # Génération Word
                try:
                    from generer_bon_paiement import generer_bon_paiement
                    data = {
                        "Date": nouvelle["Date"],
                        "No_transaction": nouvelle["No de transaction"],
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
                        "Montant_lettres": lettres,
                        "Mode_paiement": mode_pay,
                        "Numero_cheque": num_cheque.strip() or "-",
                        "Banque": banque,
                        "Utilisateur": name,
                    }
                    chemin = generer_bon_paiement(data)
                    st.success(f"✅ Enregistré + bon généré : `{os.path.basename(chemin)}`")
                except Exception as e:
                    st.warning(f"Enregistré dans Excel, mais erreur Word : {e}")
                    import traceback
                    st.code(traceback.format_exc())
                st.rerun()

# ---------- Main ----------
st.title("💳 Suivi des paiements & bons de paiement")
st.caption(f"Fichier : `{FICHIER_PAIEMENTS}` · Feuille : `{FEUILLE_DEPENSES}` · Bons Word : `{DOSSIER_BONS}/`")

df = charger_depenses()

# KPI
if not df.empty:
    montants = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Nb opérations", len(df))
    k2.metric("Total payé (FCFA)", f"{montants.sum():,.0f}".replace(",", " "))
    k3.metric("Chèques", len(df[df["Mode paiement"].astype(str).str.upper().str.contains("CHEQUE", na=False)]))
    k4.metric("Espèces / MM", len(df[df["Mode paiement"].astype(str).str.upper().str.contains("ESPECES|MOBILE", na=False, regex=True)]))
else:
    st.info("Aucune dépense chargée. Vérifiez que le fichier `Suivi de paiement.xlsx` est dans le dossier du projet.")

st.markdown("---")

# Filtres
f1, f2, f3, f4 = st.columns(4)
with f1:
    ft = st.selectbox("Type", ["Tous"] + TYPES_OP)
with f2:
    fp = st.selectbox("Produit", ["Tous"] + sorted(df["Produit"].dropna().astype(str).unique().tolist()) if not df.empty else ["Tous"])
with f3:
    fm = st.selectbox("Mode", ["Tous"] + MODES)
with f4:
    q = st.text_input("Recherche", placeholder="N°, fournisseur, réf…")

df_f = df.copy()
if not df_f.empty:
    if ft != "Tous":
        df_f = df_f[df_f["Type d'operation"].astype(str) == ft]
    if fp != "Tous":
        df_f = df_f[df_f["Produit"].astype(str) == fp]
    if fm != "Tous":
        df_f = df_f[df_f["Mode paiement"].astype(str).str.upper() == fm]
    if q.strip():
        qq = q.strip().lower()
        mask = df_f.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)
        df_f = df_f[mask]

st.subheader("📋 Journal des dépenses")
if df_f.empty:
    st.warning("Aucune ligne à afficher.")
else:
    st.dataframe(df_f, width="stretch", height=400, hide_index=True)

# Annuaire : listes déroulantes alimentées par l'Excel (lecture seule)
st.markdown("---")
st.subheader("🔎 Annuaire prestataires (données issues de l'Excel)")
annuaire = charger_annuaire(df)
if not annuaire:
    st.info("Aucun prestataire dans l'Excel pour le moment.")
else:
    a1, a2, a3, a4, a5 = st.columns(5)
    with a1:
        presta_sel = st.selectbox("Prestataire / Fournisseur", list(annuaire.keys()), key="ann_presta")
    lignes_p = annuaire[presta_sel]
    with a2:
        benefs = _uniques_recents(lignes_p[COL_BENEF]) or ["—"]
        benef_sel = st.selectbox("Nom du bénéficiaire", benefs, key="ann_benef")
    lignes_b = lignes_p[lignes_p[COL_BENEF].map(_propre) == benef_sel] if benef_sel != "—" else lignes_p
    with a3:
        pieces = _uniques_recents(lignes_b[COL_PIECE]) or ["—"]
        st.selectbox("N° pièce d'identité", pieces, key="ann_piece")
    with a4:
        exps = _uniques_recents(lignes_b[COL_EXP]) or ["—"]
        st.selectbox("Date d'expiration", exps, key="ann_exp")
    with a5:
        tels = _uniques_recents(lignes_b[COL_TEL]) or ["—"]
        st.selectbox("Adresse / Téléphone", tels, key="ann_tel")
    st.caption(f"{len(annuaire)} prestataires · {len(lignes_p)} opération(s) pour « {presta_sel} ». "
               "Les listes se mettent à jour automatiquement à chaque nouvel enregistrement.")

# Régénérer un bon depuis une ligne existante
st.markdown("---")
st.subheader("📄 Générer / régénérer un bon de paiement Word")
if not df.empty:
    options = df["No de transaction"].astype(str).tolist()
    # dedupe preserve order
    seen = set()
    options_u = []
    for o in reversed(options):
        if o not in seen and o and o != "nan":
            seen.add(o)
            options_u.append(o)
    options_u = list(reversed(options_u))

    sel = st.selectbox("N° transaction", options=options_u)
    if st.button("📄 Générer le bon Word", type="primary"):
        lignes = df[df["No de transaction"].astype(str) == str(sel)]
        if lignes.empty:
            st.error("Transaction introuvable.")
        else:
            # Si plusieurs lignes même N°, prendre la première / sommer ?
            row = lignes.iloc[-1]
            try:
                montant = float(pd.to_numeric(row["Montant"], errors="coerce") or 0)
            except Exception:
                montant = 0
            lettres = str(row.get("Montant en lettre", "") or "") or montant_en_lettres(montant)
            data = {
                "Date": str(row.get("Date", "")),
                "No_transaction": str(sel),
                "Type_operation": str(row.get("Type d'operation", "")),
                "Produit": str(row.get("Produit", "")),
                "Prestataire": str(row.get("Prestataires/Fournisseurs", "")),
                "Beneficiaire": str(row.get("Nom du bénéficiaire", "")),
                "Piece_identite": str(row.get("Num. Pieces d'identite", "")),
                "Date_expiration": str(row.get("Date d'expiration", "")),
                "Telephone": str(row.get("Addresse /Telephone", "")),
                "Ref_doc": str(row.get("Ref doc", "")),
                "Description": str(row.get("Description/Objet du paiement", "")),
                "Montant": montant,
                "Observation": str(row.get("Observation", "")),
                "Montant_lettres": lettres,
                "Mode_paiement": str(row.get("Mode paiement", "")),
                "Numero_cheque": str(row.get("Numero de cheque", "")),
                "Banque": str(row.get("Banque", "")),
                "Utilisateur": name,
            }
            try:
                from generer_bon_paiement import generer_bon_paiement
                chemin = generer_bon_paiement(data)
                st.success(f"Bon généré : `{chemin}`")
                with open(chemin, "rb") as f:
                    st.download_button(
                        "⬇️ Télécharger ce bon",
                        data=f,
                        file_name=os.path.basename(chemin),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            except Exception as e:
                st.error(f"Erreur : {e}")
                import traceback
                st.code(traceback.format_exc())

# Liste des bons générés
st.markdown("---")
st.subheader("📥 Bons Word disponibles")
import glob
fichiers = sorted(glob.glob(os.path.join(DOSSIER_BONS, "Bon_Paiement_*.docx")), reverse=True)
if not fichiers:
    st.info("Aucun bon dans le dossier bons_paiement/.")
else:
    for fpath in fichiers[:30]:
        c1, c2 = st.columns([4, 1])
        with c1:
            st.write(f"📄 **{os.path.basename(fpath)}**")
        with c2:
            with open(fpath, "rb") as f:
                st.download_button("⬇️", data=f, file_name=os.path.basename(fpath), key=f"dl_{fpath}", width="stretch")

st.caption("NAP SARL — Suivi paiements + génération bon de paiement (QR) · Placez `Suivi de paiement.xlsx` à côté de l'app")
