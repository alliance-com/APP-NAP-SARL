"""
rapport.py - Sorties du rapprochement : rapport de contrôle PDF (avec QR code) et relevé rapproché Excel.
Dépendances : reportlab (QR code intégré, rien d'autre à installer), openpyxl.
"""
import hashlib
import io
from datetime import datetime

import pandas as pd
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, KeepTogether, PageBreak)

from moteur import (fmt_xof, STATUT_OK, STATUT_VERIF, STATUT_ECART, STATUT_ABSENT, STATUT_EXCEL_SEUL)

VERT = colors.HexColor("#0B6E4F")
GRIS = colors.HexColor("#F2F4F3")
ROUGE = colors.HexColor("#B3261E")
ORANGE = colors.HexColor("#B26A00")

def fd(d):
    return d.strftime("%d/%m/%Y") if d else ""

def reference_controle(meta, synthese, banque):
    base = f"{meta.get('compte')}|{synthese['date_fin']}|{synthese['solde_banque_cloture']}|" \
           f"{synthese['credits_banque']}|{synthese['debits_banque']}|{len(banque)}"
    h = hashlib.sha256(base.encode()).hexdigest()[:8].upper()
    return f"RB-{synthese['date_fin']:%Y%m%d}-{h}"

def texte_qr(meta, synthese, reference, date_controle):
    return (f"{meta.get('societe', 'NAP SARL')} | {meta.get('banque', 'Ecobank')} | Cpt {meta.get('compte')}\n"
            f"Periode {fd(synthese['date_depart'])} au {fd(synthese['date_fin'])}\n"
            f"Solde de cloture {meta.get('devise', 'XOF')} {synthese['solde_banque_cloture']}\n"
            f"Controle du {fd(date_controle)} | Ecart non explique {int(round(synthese['non_explique']))}\n"
            f"Ref {reference}")

def _qr(texte, taille=30 * mm):
    w = QrCodeWidget(texte)
    x0, y0, x1, y1 = w.getBounds()
    d = Drawing(taille, taille, transform=[taille / (x1 - x0), 0, 0, taille / (y1 - y0), 0, 0])
    d.add(w)
    return d

def generer_rapport_pdf(meta, synthese, banque, excel, observations=None, etabli_par="", date_controle=None,
                        remarque_generale="", controles_releve=None):
    observations = observations or {}
    date_controle = date_controle or datetime.now().date()
    ref = reference_controle(meta, synthese, banque)
    devise = meta.get("devise", "XOF")

    ss = getSampleStyleSheet()
    n = ParagraphStyle("n", parent=ss["Normal"], fontName="Helvetica", fontSize=8.5, leading=10.5)
    pt = ParagraphStyle("pt", parent=n, fontSize=7.3, leading=8.8)
    h1 = ParagraphStyle("h1", parent=n, fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=VERT)
    h2 = ParagraphStyle("h2", parent=n, fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=VERT,
                        spaceBefore=9, spaceAfter=4)
    droite = ParagraphStyle("d", parent=n, alignment=2)

    flux = []

    titre = [Paragraph("RAPPORT DE CONTRÔLE BANCAIRE", h1),
             Paragraph(f"{meta.get('societe', 'STE NAP SARL')} - Rapprochement relevé {meta.get('banque', 'Ecobank')} / "
                       f"registre interne", n),
             Spacer(1, 3), Paragraph(f"<b>Réf. {ref}</b>", n)]
    qr_bloc = [_qr(texte_qr(meta, synthese, ref, date_controle)),
               Paragraph("Scanner pour vérifier", ParagraphStyle("c", parent=pt, alignment=1))]
    t = Table([[titre, qr_bloc]], colWidths=[140 * mm, 38 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (1, 0), (1, 0), "CENTER")]))
    flux += [t, Spacer(1, 5)]

    ident = [
        ["Banque", meta.get("banque", "Ecobank").upper(), "Date du contrôle", fd(date_controle)],
        ["Compte", f"{meta.get('compte', '')} ({meta.get('nom_compte', '')})", "Arrêté au", fd(synthese["date_fin"])],
        ["Période", f"{fd(synthese['date_depart'])} au {fd(synthese['date_fin'])}", "Établi par", etabli_par or "-"],
    ]
    ti = Table(ident, colWidths=[22 * mm, 74 * mm, 30 * mm, 52 * mm])
    ti.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 8.5),
        ("FONT", (2, 0), (2, -1), "Helvetica-Bold", 8.5), ("BACKGROUND", (0, 0), (-1, -1), GRIS),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    flux += [ti, Spacer(1, 5)]

    ok = abs(synthese["non_explique"]) < 1
    clot = Table([[Paragraph(f"SOLDE DE CLÔTURE AU {fd(synthese['date_fin'])} - {meta.get('banque', '').upper()}",
                             ParagraphStyle("x", parent=n, fontName="Helvetica-Bold", fontSize=9, textColor=colors.white)),
                   Paragraph(f"{devise} {fmt_xof(synthese['solde_banque_cloture'])}",
                             ParagraphStyle("y", parent=n, fontName="Helvetica-Bold", fontSize=15, leading=18,
                                            textColor=colors.white, alignment=2))]],
                 colWidths=[88 * mm, 90 * mm])
    clot.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), VERT), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                              ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    flux += [clot]

    if controles_releve:
        flux.append(Paragraph("1. Contrôle de lecture du relevé bancaire", h2))
        lignes = [[("OK" if okc else "ANOMALIE"), msg] for okc, msg in controles_releve]
        tc = Table(lignes, colWidths=[20 * mm, 158 * mm])
        st = [("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey)]
        for i, (okc, _) in enumerate(controles_releve):
            st.append(("TEXTCOLOR", (0, i), (0, i), VERT if okc else ROUGE))
            st.append(("FONT", (0, i), (0, i), "Helvetica-Bold", 8))
        tc.setStyle(TableStyle(st))
        flux.append(tc)

    flux.append(Paragraph("2. Rapprochement banque / registre interne (Excel)", h2))
    s = synthese
    corps = [
        ["", "Banque", "Excel interne"],
        ["Solde de départ", fmt_xof(s["solde_depart"]), fmt_xof(s["solde_depart"])],
        ["+ Crédits de la période", fmt_xof(s["credits_banque"]), fmt_xof(s["credits_excel"])],
        ["- Débits de la période", fmt_xof(s["debits_banque"]), fmt_xof(s["debits_excel"])],
        [f"= Solde au {fd(s['date_fin'])}", fmt_xof(s["solde_banque_cloture"]), fmt_xof(s["solde_excel_cloture"])],
    ]
    tr = Table(corps, colWidths=[78 * mm, 50 * mm, 50 * mm])
    tr.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
        ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, 0), (-1, 0), GRIS), ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey)]))
    flux += [tr, Spacer(1, 5)]

    expl = [
        ["Écart à expliquer (banque - Excel)", "", fmt_xof(s["ecart_total"])],
        [f"Opérations en banque absentes d'Excel ({s['nb_banque_seule']})", "à saisir dans Excel", fmt_xof(s["net_banque_seule"])],
        [f"Opérations Excel pas encore en banque ({s['nb_excel_seul']})", "en attente", fmt_xof(-s["net_excel_seul"])],
        [f"Écarts de montant sur opérations rapprochées ({s['nb_ecarts_montant']})", "change / frais", fmt_xof(s["net_ecarts_montant"])],
        [f"Écritures Excel d'avant la période ({s['nb_hors_periode']})", "rapprochées", fmt_xof(s["net_hors_periode"])],
        ["Écart sur le solde de départ (ouverture banque - solde saisi)", "", fmt_xof(s["ecart_ouverture"])],
        ["TOTAL EXPLIQUÉ", "", fmt_xof(s["ecart_explique"])],
        ["ÉCART NON EXPLIQUÉ", "OK - contrôle bouclé" if ok else "À INVESTIGUER", fmt_xof(s["non_explique"])],
    ]
    te = Table(expl, colWidths=[104 * mm, 36 * mm, 38 * mm])
    te.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8.3), ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5), ("FONT", (0, 6), (-1, -1), "Helvetica-Bold", 9),
        ("TEXTCOLOR", (1, 1), (1, 5), colors.grey),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey), ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E3F4EC") if ok else colors.HexColor("#FBE3E1")),
        ("TEXTCOLOR", (1, -1), (2, -1), VERT if ok else ROUGE)]))
    flux.append(te)

    kpi = (f"{s['nb_lignes_banque']} opérations en banque sur la période : <b>{s['nb_rapprochees']}</b> rapprochées, "
           f"<b>{s['nb_a_verifier']}</b> à vérifier, <b>{s['nb_ecarts_montant']}</b> avec écart de montant, "
           f"<b>{s['nb_banque_seule']}</b> absentes d'Excel.")
    flux += [Spacer(1, 4), Paragraph(kpi, n)]

    def tableau(titre, df, colonnes, largeurs, titres):
        if df.empty:
            return []
        data = [[Paragraph(f"<b>{x}</b>", pt) for x in titres]]
        for _, r in df.iterrows():
            data.append([Paragraph(str(r[c]) if str(r[c]) != "nan" else "", pt) for c in colonnes])
        tb = Table(data, colWidths=largeurs, repeatRows=1)
        tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), GRIS), ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                                ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 2),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
        return [Paragraph(titre, h2), tb]

    b = banque.copy()
    b["Obs"] = b.apply(lambda r: observations.get(int(r["No"]), r.get("Détail", "")) or "", axis=1)
    b["D"] = b["Date"].apply(fd)
    b["Déb"] = b["Débit"].apply(lambda x: fmt_xof(x) if x else "")
    b["Cré"] = b["Crédit"].apply(lambda x: fmt_xof(x) if x else "")
    b["Lib"] = b.apply(lambda r: (r["Libellé propre"] + (f" - {r['Bénéficiaire']}" if r["Bénéficiaire"] else "")), axis=1)
    b["Chq"] = b["N° chèque"]
    d, f = s["date_depart"], s["date_fin"]
    b = b[b["Date"].apply(lambda x: x is not None and d <= x <= f)]

    flux.append(PageBreak())
    flux.append(Paragraph("3. Opérations à traiter", h2))
    cols = ["No", "D", "Lib", "Chq", "Déb", "Cré", "Obs"]
    larg = [9 * mm, 19 * mm, 50 * mm, 17 * mm, 22 * mm, 22 * mm, 39 * mm]
    tit = ["N°", "Date", "Libellé", "Chèque", "Débit", "Crédit", "Observation"]
    flux += tableau("3.1 Opérations de la banque absentes d'Excel - à saisir dans Excel",
                    b[b["Statut"] == STATUT_ABSENT], cols, larg, tit)
    flux += tableau("3.2 Rapprochements à vérifier", b[b["Statut"] == STATUT_VERIF], cols, larg, tit)
    flux += tableau("3.3 Écarts de montant", b[b["Statut"] == STATUT_ECART], cols, larg, tit)

    e = excel[excel["En période"] & (excel["Statut"] == STATUT_EXCEL_SEUL)].copy()
    if not e.empty:
        e["D"] = e["Date"].apply(fd)
        e["Déb"] = e["Débit"].apply(lambda x: fmt_xof(x) if x else "")
        e["Cré"] = e["Crédit"].apply(lambda x: fmt_xof(x) if x else "")
        e["L"] = e["Ligne Excel"].astype(str)
        flux += tableau("3.4 Écritures Excel pas encore en banque (chèques non encaissés, virements attendus)",
                        e, ["L", "D", "Description", "N° chèque", "Déb", "Cré", "Motif"],
                        [12 * mm, 19 * mm, 50 * mm, 17 * mm, 22 * mm, 22 * mm, 36 * mm],
                        ["Ligne", "Date", "Description", "Chèque", "Débit", "Crédit", "Motif"])

    if remarque_generale.strip():
        flux += [Paragraph("4. Observations générales", h2), Paragraph(remarque_generale.replace("\n", "<br/>"), n)]

    sig = Table([["Établi par", "Vérifié par", "Direction"], ["\n\n\n", "", ""],
                 [etabli_par or "", "", ""]], colWidths=[59 * mm] * 3)
    sig.setStyle(TableStyle([("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5), ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
                             ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                             ("FONT", (0, 2), (-1, 2), "Helvetica", 8)]))
    flux += [Spacer(1, 12), KeepTogether([sig])]

    def pied(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.grey)
        canvas.drawString(15 * mm, 9 * mm, f"{meta.get('societe', '')} - Rapport de contrôle bancaire - Réf. {ref} - "
                                          f"généré le {datetime.now():%d/%m/%Y %H:%M}")
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, f"Page {doc.page}")
        canvas.restoreState()

    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=13 * mm,
                      bottomMargin=16 * mm, title=f"Rapport de contrôle bancaire {ref}",
                      author=etabli_par or "NAP SARL").build(flux, onFirstPage=pied, onLaterPages=pied)
    return buf.getvalue(), ref

def generer_excel(meta, synthese, banque, excel, observations=None):
    observations = observations or {}
    d, f = synthese["date_depart"], synthese["date_fin"]
    b = banque[banque["Date"].apply(lambda x: x is not None and d <= x <= f)].copy()
    b["Observation"] = b.apply(lambda r: observations.get(int(r["No"]), r.get("Détail", "")) or "", axis=1)

    releve = b[["No", "Date", "Date valeur", "Libellé propre", "Bénéficiaire", "N° chèque", "Débit", "Crédit",
                "Solde banque", "Statut", "Méthode", "Ligne Excel liée", "Observation", "Libellé banque"]].rename(
        columns={"No": "N°", "Libellé propre": "Libellé", "Ligne Excel liée": "Ligne Excel"})

    a = b[b["Statut"] == STATUT_ABSENT]
    a_saisir = pd.DataFrame({
        "Date": a["Date"],
        "DESCRIPTION": a.apply(lambda r: (r["Bénéficiaire"] or r["Libellé propre"]).upper(), axis=1),
        "CREDIT": a["Crédit"].replace(0, pd.NA),
        "DEBIT": a["Débit"].replace(0, pd.NA),
        "SOLDE": "",
        "NUM DE CHEQUE": a["N° chèque"],
        "MOTIF": a["Libellé propre"],
        "CATEGORIE": a["Catégorie suggérée"],
    })

    e = excel[excel["En période"] & (excel["Statut"] == STATUT_EXCEL_SEUL)][
        ["Ligne Excel", "Date", "Description", "N° chèque", "Débit", "Crédit", "Motif", "Catégorie"]]

    resume = pd.DataFrame([
        ("Banque", meta.get("banque")), ("Compte", meta.get("compte")),
        ("Période", f"{d:%d/%m/%Y} au {f:%d/%m/%Y}"),
        ("Solde de départ", synthese["solde_depart"]),
        ("Crédits banque", synthese["credits_banque"]), ("Débits banque", synthese["debits_banque"]),
        ("Solde de clôture banque", synthese["solde_banque_cloture"]),
        ("Solde Excel recalculé", synthese["solde_excel_cloture"]),
        ("Écart à expliquer", synthese["ecart_total"]),
        ("Banque absente d'Excel (net)", synthese["net_banque_seule"]),
        ("Excel pas encore en banque (net)", -synthese["net_excel_seul"]),
        ("Écarts de montant", synthese["net_ecarts_montant"]),
        ("Écritures d'avant la période", synthese["net_hors_periode"]),
        ("Écart de solde de départ", synthese["ecart_ouverture"]),
        ("ÉCART NON EXPLIQUÉ", synthese["non_explique"]),
    ], columns=["Indicateur", "Valeur"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="dd/mm/yyyy", datetime_format="dd/mm/yyyy") as xw:
        resume.to_excel(xw, sheet_name="Synthèse", index=False)
        releve.to_excel(xw, sheet_name="Relevé rapproché", index=False)
        a_saisir.to_excel(xw, sheet_name="À saisir dans Excel", index=False)
        e.to_excel(xw, sheet_name="Pas encore en banque", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                largeur = max(len(str(c.value)) if c.value is not None else 0 for c in col[:60])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, largeur + 2), 60)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                        c.number_format = '#,##0'
    return buf.getvalue()