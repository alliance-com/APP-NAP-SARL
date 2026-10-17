from docx import Document
from openpyxl import load_workbook
from datetime import datetime
from docx.shared import RGBColor, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from io import BytesIO
import os
import re
import sys

try:
    import qrcode
except ImportError:
    qrcode = None

# ==================== CONFIGURATION ====================
FICHIER_EXCEL = "donnees_bons.xlsx"
MODELE_WORD   = "BON DE COMMANDE SOYABEAN [2].docx"
DOSSIER_SORTIE = "bons_generes"
SUFFIXE_REFERENCE = "-BC-SOYABEAN/NAP-2026"

# Code produit (dans la référence) -> libellé affiché sur le bon
# Le modèle Word parle de « SOYABEAN » : on le remplace selon le produit choisi.
LIBELLES_PRODUIT = {
    "SHEA": "KARITE",
    "SOYABEAN": "SOJA",
    "COTTON": "COTON",
}
# Le modèle Word (BON DE COMMANDE SOYABEAN [2].docx) contient « SOJA » / « SOYABEAN ».
# On remplace à la génération selon le produit — le fichier modèle sur disque ne change pas.
# =======================================================

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

os.makedirs(DOSSIER_SORTIE, exist_ok=True)


# ==================== MISE EN PAGE : 1 SEULE FEUILLE A4 ====================
def _tous_les_paragraphes(doc):
    """Paragraphes du corps + paragraphes dans les tableaux."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _est_paragraphe_vide(p):
    if p.text.strip():
        return False
    return not p._p.xpath(".//w:drawing") and not p._p.xpath(".//w:pict")


def _reduire_paragraphe_vide(p, taille_pt=3):
    """Un paragraphe vide garde la hauteur d'une ligne : on réduit son 'pied de mouche'."""
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(int(taille_pt * 2)))
    for r in p.runs:
        r.font.size = Pt(taille_pt)


def forcer_une_seule_page(doc, echelle_police=1.0):
    """1 seule feuille A4 : contenu bien réparti, espacements serrés et réguliers.
    - Marges stables pour remplir la page sans déborder
    - Lignes de détail (- NOM, - QUANTITÉ…) serrées
    - Titres de section un peu plus aérés
    - Paragraphes vides réduits (plus de grands trous blancs)
    """
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        # Marges un peu plus larges qu'avant pour « remplir » la feuille sans coller au bord
        section.top_margin = Cm(1.3)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        section.header_distance = Cm(0.4)
        section.footer_distance = Cm(0.4)

    titres_section = (
        "FOURNISSEUR", "DETAILS", "MODALIT", "QUALITÉ", "QUALITE", "NOTE",
        "ACHETEUR", "REFERENCE", "BON DE COMMANDE", "CAMPAGNE",
    )

    for p in _tous_les_paragraphes(doc):
        if _est_paragraphe_vide(p):
            # Réduit fortement les trous blancs du modèle
            _reduire_paragraphe_vide(p, taille_pt=2)
            continue

        pf = p.paragraph_format
        pf.line_spacing = 1.08
        pf.keep_together = True
        pf.page_break_before = False
        pf.widow_control = True

        texte = p.text.strip().upper()

        if texte.startswith("-"):
            # Lignes de détail : serrées, lisibles
            pf.space_before = Pt(1)
            pf.space_after = Pt(2)
        elif any(t in texte for t in titres_section):
            # Titres / blocs importants : petit air avant, serré après
            pf.space_before = Pt(6)
            pf.space_after = Pt(3)
        else:
            pf.space_before = Pt(2)
            pf.space_after = Pt(2)

        if echelle_police != 1.0:
            for r in p.runs:
                if r.font.size:
                    r.font.size = Pt(max(8, round(r.font.size.pt * echelle_police * 2) / 2))

    # Tableau qualité : compact mais lisible
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.05


# ==================== PRODUIT ====================
def detecter_code_produit(reference_brute, produit=None):
    """Retourne SHEA / SOYABEAN / COTTON (ou None) d'après la référence, sinon le nom du produit."""
    ref = str(reference_brute or "").upper().replace("\\", "/")
    for code in LIBELLES_PRODUIT:
        if f"-BC-{code}" in ref:
            return code
    mapping = {
        "shea": "SHEA", "karité": "SHEA", "karite": "SHEA",
        "soja": "SOYABEAN", "soy": "SOYABEAN",
        "cotton": "COTTON", "coton": "COTTON",
    }
    if produit:
        p = str(produit).lower()
        for k, v in mapping.items():
            if k in p:
                return v
    return None


def adapter_produit_dans_modele(doc, code_produit):
    """Adapte le texte du modèle (SOJA / SOYABEAN) au produit du bon.
    Le modèle source reste BON DE COMMANDE SOYABEAN [2].docx — on ne le modifie pas sur disque."""
    libelle = LIBELLES_PRODUIT.get(code_produit)
    if not libelle:
        return
    # Mots du modèle à remplacer (casse variable)
    cibles = ["SOYABEAN", "SOJA"]
    for p in _tous_les_paragraphes(doc):
        for r in p.runs:
            original = r.text
            nouveau = original
            for cible in cibles:
                if cible.lower() in nouveau.lower():
                    nouveau = re.sub(cible, libelle, nouveau, flags=re.IGNORECASE)
            if nouveau != original:
                r.text = nouveau


# ==================== QR CODE ====================
def generer_qr_png_bytes(texte: str) -> BytesIO:
    if qrcode is None:
        raise ImportError("Installez qrcode : pip install qrcode[pil]")
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=6,
        border=2,
    )
    qr.add_data(texte)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def inserer_qr_haut_droite(doc, texte_qr: str, largeur_cm: float = 1.6):
    """QR en haut à droite.
    Le modèle contient déjà une image QR ancrée (1,6 cm) : on REMPLACE son contenu
    par le QR du bon -> aucune ligne ajoutée, donc aucun risque de 2e page."""
    png = generer_qr_png_bytes(texte_qr).getvalue()
    rids = doc.element.xpath(".//a:blip/@r:embed")
    if rids:
        doc.part.related_parts[rids[0]]._blob = png
        return
    # Repli : le modèle n'a pas d'image -> QR en ligne dans le 1er paragraphe
    if doc.paragraphs:
        p = doc.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(BytesIO(png), width=Cm(largeur_cm))


# ==================== FILIGRANE ====================
def ajouter_filigrane_centre(doc, reference):
    date_str = datetime.now().strftime("%d/%m/%Y")
    texte = f"STE NAP SARL  •  {reference}  •  {date_str}"

    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False

    for p in header.paragraphs:
        p.clear()
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()

    r = OxmlElement("w:r")
    paragraph._p.append(r)

    rPr = OxmlElement("w:rPr")
    noProof = OxmlElement("w:noProof")
    rPr.append(noProof)
    r.append(rPr)

    pict = OxmlElement("w:pict")
    r.append(pict)

    shape_xml = f'''
    <v:shape xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             id="FiligraneNAP"
             type="#_x0000_t136"
             style="position:absolute;margin-left:0;margin-top:0;width:500pt;height:60pt;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;"
             fillcolor="#C8C8C8"
             stroked="f">
      <v:fill opacity=".40"/>
      <v:textbox inset="0,0,0,0">
        <w:txbxContent>
          <w:p>
            <w:pPr>
              <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:b/>
                <w:sz w:val="28"/>
                <w:szCs w:val="28"/>
                <w:color w:val="A0A0A0"/>
              </w:rPr>
              <w:t xml:space="preserve">{texte}</w:t>
            </w:r>
          </w:p>
        </w:txbxContent>
      </v:textbox>
    </v:shape>
    '''
    shape = etree.fromstring(shape_xml)
    pict.append(shape)


# ==================== OUTILS ====================
def extraire_numero(reference):
    if not reference:
        return 0
    base = os.path.basename(str(reference)).split("-")[0]
    match = re.search(r'(\d+)', base)
    if match:
        return int(match.group(1))
    match_fallback = re.search(r'(\d+)', str(reference))
    return int(match_fallback.group(1)) if match_fallback else 0


def formater_reference(reference_brute, produit=None):
    """Conserve le suffixe produit déjà présent dans la référence Excel.
    Sinon déduit du produit, sinon SUFFIXE_REFERENCE par défaut.
    Ex : SOYABEAN -> 0308-BC-SOYABEAN/NAP-2026 ; SOYABEAN -> 0308-BC-SOYABEAN/NAP-2026
    """
    ref = str(reference_brute or "").strip()
    num = extraire_numero(ref)
    # Si la référence contient déjà un suffixe produit connu, on le garde
    for code in LIBELLES_PRODUIT:
        if f"-BC-{code}/NAP-2026" in ref.replace("\\", "/"):
            if num > 0:
                return f"{num:04d}-BC-{code}/NAP-2026"
            return ref
    # Déduire depuis le nom produit
    code = detecter_code_produit("", produit)
    if num > 0:
        if code:
            return f"{num:04d}-BC-{code}/NAP-2026"
        return f"{num:04d}{SUFFIXE_REFERENCE}"
    return ref


def assainir_nom_fichier(nom):
    if not nom:
        return ""
    nom_propre = re.sub(r'[\\/*?:"<>|]', "", str(nom)).strip()
    nom_propre = re.sub(r'\s+', ' ', nom_propre)
    return nom_propre


def nettoyer_valeur(val):
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    val_str = str(val).strip()
    return "" if val_str.lower() in ("none", "nan") else val_str


def remplacer_dans_paragraphe(paragraph, valeur, couleur=None):
    valeur = nettoyer_valeur(valeur)
    if not valeur:
        return
    trouve = False
    for r in paragraph.runs:
        if '\xa0' in r.text:
            r.text = r.text.replace('\xa0', ' ')
        if '…' in r.text or '..' in r.text:
            if not trouve:
                r.text = str(valeur)
                if couleur:
                    r.font.color.rgb = couleur
                else:
                    if r.font.color and r.font.color.rgb:
                        r.font.color.rgb = None
                trouve = True
            else:
                r.text = ''
        elif trouve:
            if r.text.strip() in ('.', '..', '...'):
                r.text = ''


# ==================== GÉNÉRATION ====================
def generer_tous_les_bons(forcer_tout=False):
    if not os.path.exists(FICHIER_EXCEL):
        print(f"Erreur : Le fichier {FICHIER_EXCEL} est introuvable.")
        return

    if not os.path.exists(MODELE_WORD):
        print(f"Erreur : Le modèle Word {MODELE_WORD} est introuvable.")
        return

    if qrcode is None:
        print("Erreur : module qrcode manquant. Lancez : pip install qrcode[pil]")
        return

    wb = load_workbook(FICHIER_EXCEL)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    compteur = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        data = dict(zip(headers, row))
        ref_brute = data.get("Reference")
        if not ref_brute:
            continue

        nom = data.get("Nom_Fournisseur")
        reference = formater_reference(ref_brute, produit=data.get('Produit'))
        code_produit = detecter_code_produit(reference, data.get('Produit'))
        ref_propre = reference.replace("/", "-").replace("\\", "-").strip()
        fourn_propre = assainir_nom_fichier(nom)

        if fourn_propre:
            nom_fichier = f"{ref_propre} - {fourn_propre}.docx"
        else:
            nom_fichier = f"{ref_propre}.docx"

        chemin_sortie = os.path.join(DOSSIER_SORTIE, nom_fichier)

        ancien_chemin = os.path.join(DOSSIER_SORTIE, f"{ref_propre}.docx")
        if ancien_chemin != chemin_sortie and os.path.exists(ancien_chemin):
            try:
                os.remove(ancien_chemin)
            except Exception:
                pass

        if os.path.exists(chemin_sortie) and not forcer_tout:
            continue

        doc = Document(MODELE_WORD)

        # Filigrane
        ajouter_filigrane_centre(doc, reference)

        # Texte du modèle selon le produit (KARITE / GRAINES DE SOJA / ...)
        adapter_produit_dans_modele(doc, code_produit)

        date_val = data.get("Date_Emission")
        if isinstance(date_val, datetime):
            date_emission = date_val.strftime("%d/%m/%Y")
        elif date_val:
            date_emission = str(date_val).strip()
        else:
            date_emission = datetime.now().strftime("%d/%m/%Y")

        heure_gen = datetime.now().strftime("%H:%M:%S")
        utilisateur_gen = nettoyer_valeur(data.get("Utilisateur")) or "Systeme"

        tel_fourn     = data.get("Telephone_Fournisseur")
        commune       = data.get("Commune_Origine")
        quantite      = data.get("Quantite")
        tracteur      = data.get("Numero_Tracteur")
        chauffeur     = data.get("Chauffeur")
        permis        = data.get("Numero_Permis")
        tel_chauffeur = data.get("Telephone_Chauffeur")

        # --- QR haut droite (remplace le QR du modèle) ---
        texte_qr = (
            f"STE NAP SARL|BC|{reference}|"
            f"{nettoyer_valeur(nom)}|"
            f"{date_emission}|{heure_gen}|{utilisateur_gen}"
        )
        inserer_qr_haut_droite(doc, texte_qr, largeur_cm=1.6)

        section_actuelle = None

        for paragraph in doc.paragraphs:
            texte = paragraph.text.strip()

            if "FOURNISSEUR" in texte and "DETAILS" not in texte:
                section_actuelle = "FOURNISSEUR"
                continue
            elif "DETAILS SUR LE PRODUIT" in texte:
                section_actuelle = "TRANSPORT"
                continue
            elif "MODALIT" in texte:
                section_actuelle = "MODALITES"
                continue

            if "REFERENCE" in texte and ("…" in texte or ".." in texte):
                remplacer_dans_paragraphe(paragraph, reference, RGBColor(255, 0, 0))

            elif ("DATE D" in texte and "MISSION" in texte) or "DATE D’ÉMISSION" in texte or "DATE D'ÉMISSION" in texte:
                if date_emission and paragraph.runs:
                    paragraph.runs[-1].text = f" {date_emission}"

            elif section_actuelle == "FOURNISSEUR":
                if texte.startswith("-   NOM"):
                    remplacer_dans_paragraphe(paragraph, nom)
                elif "TÉLÉPHONE" in texte or "PHONE" in texte:
                    remplacer_dans_paragraphe(paragraph, tel_fourn)
                elif "COMMUNE" in texte:
                    remplacer_dans_paragraphe(paragraph, commune)

            elif section_actuelle == "TRANSPORT":
                if "QUANTITÉ" in texte or "QUANTIT" in texte:
                    remplacer_dans_paragraphe(paragraph, quantite)
                elif "TRACTEUR" in texte or "REMORQUE" in texte:
                    remplacer_dans_paragraphe(paragraph, tracteur)
                elif "CHAUFFEUR" in texte:
                    remplacer_dans_paragraphe(paragraph, chauffeur)
                elif "PERMIS" in texte:
                    remplacer_dans_paragraphe(paragraph, permis)
                elif "TÉLÉPHONE" in texte or "PHONE" in texte:
                    remplacer_dans_paragraphe(paragraph, tel_chauffeur)

        # Mise en page : tout sur UNE seule feuille A4
        forcer_une_seule_page(doc)

        doc.save(chemin_sortie)
        compteur += 1
        print(f"[OK] Bon genere : {chemin_sortie}")

    print("-" * 60)
    if compteur == 0:
        print("Aucun nouveau bon à générer (tous les fichiers sont déjà présents).")
    else:
        print(f"Terminé ! {compteur} bon(s) généré(s) avec succès.")


if __name__ == "__main__":
    forcer = "--tout" in sys.argv or "--force" in sys.argv or "--regenerer" in sys.argv
    generer_tous_les_bons(forcer_tout=forcer)
