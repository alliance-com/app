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
MODELE_WORD   = "BON DE COMMANDE KARITE [2].docx"
DOSSIER_SORTIE = "bons_generes"
SUFFIXE_REFERENCE = "-BC-SHEA/NAP-2026"
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


def inserer_qr_haut_droite(doc, texte_qr: str, largeur_cm: float = 2.2):
    """QR en haut à droite (coin supérieur droit)."""
    if doc.paragraphs:
        p = doc.paragraphs[0].insert_paragraph_before()
    else:
        p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.add_picture(generer_qr_png_bytes(texte_qr), width=Cm(largeur_cm))
    # Petite légende sous le QR, aussi à droite
    if doc.paragraphs:
        leg = doc.paragraphs[0].insert_paragraph_before() if False else None
    # Légende juste après le paragraphe QR : on récupère le paragraphe qu'on vient de créer
    # (dernier insert = premier body para)
    p_leg = doc.paragraphs[0].insert_paragraph_before() if False else doc.paragraphs[1] if len(doc.paragraphs) > 1 else None
    # Simpler: add caption on same structure
    # Re-get first paragraph after our QR - actually insert_paragraph_before stacks in reverse
    # Clean approach: one right-aligned block
    pass


def inserer_qr_haut_droite(doc, texte_qr: str, largeur_cm: float = 2.2):
    """QR + légende en haut à droite."""
    # Légende d'abord puis QR : avec insert_before, le dernier inséré se retrouve en haut
    if doc.paragraphs:
        p_leg = doc.paragraphs[0].insert_paragraph_before()
    else:
        p_leg = doc.add_paragraph()
        p_qr = doc.add_paragraph()
        p_qr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p_qr.add_run()
        run.add_picture(generer_qr_png_bytes(texte_qr), width=Cm(largeur_cm))
        p_leg.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p_leg.add_run("Scan – STE NAP SARL")
        r.font.size = Pt(7)
        r.font.italic = True
        return

    p_leg.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_leg.add_run("Scan – STE NAP SARL")
    r.font.size = Pt(7)
    r.font.italic = True

    p_qr = doc.paragraphs[0].insert_paragraph_before()
    p_qr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p_qr.add_run()
    run.add_picture(generer_qr_png_bytes(texte_qr), width=Cm(largeur_cm))


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


def extraire_numero(reference):
    if not reference:
        return 0
    base = os.path.basename(str(reference)).split("-")[0]
    match = re.search(r'(\d+)', base)
    if match:
        return int(match.group(1))
    match_fallback = re.search(r'(\d+)', str(reference))
    return int(match_fallback.group(1)) if match_fallback else 0


def formater_reference(reference_brute):
    num = extraire_numero(reference_brute)
    if num > 0:
        return f"{num:04d}{SUFFIXE_REFERENCE}"
    return str(reference_brute).strip()


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
        reference = formater_reference(ref_brute)
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

        # --- QR haut droite ---
        texte_qr = (
            f"STE NAP SARL|BC|{reference}|"
            f"{nettoyer_valeur(nom)}|"
            f"{date_emission}|{heure_gen}|{utilisateur_gen}"
        )
        inserer_qr_haut_droite(doc, texte_qr, largeur_cm=2.2)

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