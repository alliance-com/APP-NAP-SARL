import pandas as pd
from supabase import create_client
from datetime import datetime

SUPABASE_URL = "https://fwcqtbjrgiyhdqswvryt.supabase.co"
SUPABASE_KEY = "sb_publishable_4mztoC8qcbH8LiI7LAcbog_9pKxvvPr"

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# 1. Lire DEPENSES, en-têtes ligne 3
df = pd.read_excel("Suivi de paiement.xlsx", sheet_name="DEPENSES", header=2)
df = df.dropna(axis=1, how="all")
df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]
print("Colonnes Excel:", df.columns.tolist())
print("Lignes:", len(df))

# 2. Garder seulement les vraies lignes
df = df[df["No de transaction"].notna()]
df = df[df["No de transaction"].astype(str).str.strip().isin(["", "nan", "none"]) == False]

def clean_date(v):
    if pd.isna(v) or str(v).strip() in ("", "-", "nan"):
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    # "02/06/2026" ou "2026-06-02"
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).strftime("%Y-%m-%d")
        except:
            pass
    return None

def clean_text(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    t = str(v).strip()
    if t.lower() in ("nan", "none", "nat"):
        return ""
    return t

# 3. Construire les lignes pour Supabase
lignes = []
for _, r in df.iterrows():
    # Montant en lettre : si c'est une formule, on laisse vide, ton app le recalculera
    en_lettre = clean_text(r.get("Montant en lettre", ""))
    if en_lettre.startswith("="):
        en_lettre = ""

    lignes.append({
        "date_paiement": clean_date(r.get("Date")),
        "no_transaction": clean_text(r.get("No de transaction")),
        "type_operation": clean_text(r.get("Type d'operation")),
        "produit": clean_text(r.get("Produit")),
        "prestataire_fournisseur": clean_text(r.get("Prestataires/Fournisseurs")),
        "nom_beneficiaire": clean_text(r.get("Nom du bénéficiaire")),
        "num_piece_identite": clean_text(r.get("Num. Pieces d'identite")),
        "date_expiration": clean_text(r.get("Date d'expiration")),
        "adresse_telephone": clean_text(r.get("Addresse /Telephone")),
        "ref_doc": clean_text(r.get("Ref doc")),
        "description_objet": clean_text(r.get("Description/Objet du paiement")),
        "montant": float(r.get("Montant") or 0),
        "observation": clean_text(r.get("Observation")),
        "montant_en_lettre": en_lettre,
        "mode_paiement": clean_text(r.get("Mode paiement")),
        "numero_cheque": clean_text(r.get("Numero de cheque")),
        "banque": clean_text(r.get("Banque")),
    })

print("Lignes à importer:", len(lignes))

# 4. Envoyer par paquets de 100
for i in range(0, len(lignes), 100):
    paquet = lignes[i:i+100]
    supabase.table("depenses").insert(paquet).execute()
    print(f"Importé {i+len(paquet)} / {len(lignes)}")

print("✅ FIN DEPENSES")

# 5. VERSEMENTS, pareil
try:
    df2 = pd.read_excel("Suivi de paiement.xlsx", sheet_name="VERSEMENTS", header=2)
    df2 = df2.dropna(axis=1, how="all")
    df2 = df2.loc[:, [c for c in df2.columns if not str(c).startswith("Unnamed")]]
    lignes2 = []
    for _, r in df2.iterrows():
        if pd.isna(r.get("MONTANT")):
            continue
        lignes2.append({
            "date_versement": clean_date(r.get("DATE")),
            "doc_ref": clean_text(r.get("DOC REF")),
            "nom_deposant": clean_text(r.get("NOM DU DEPOSANT")),
            "motif_versement": clean_text(r.get("MOTIF DE VERSEMENT")),
            "montant": float(r.get("MONTANT") or 0),
            "banque": clean_text(r.get("BANQUE")),
        })
    supabase.table("versements").insert(lignes2).execute()
    print(f"✅ VERSEMENTS importés: {len(lignes2)}")
except Exception as e:
    print("Erreur versements:", e)
