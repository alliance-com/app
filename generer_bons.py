from docx import Document
from openpyxl import load_workbook
from datetime import datetime
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import os
import re
import sys

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

def ajouter_filigrane_centre(doc, reference):
    """
    Vrai filigrane au centre de la page (derrière le texte).
    STE NAP SARL + N° référence + date du jour.
    N'ajoute aucune page.
    """
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

        # Filigrane au centre de la page
        ajouter_filigrane_centre(doc, reference)

        date_val = data.get("Date_Emission")
        if isinstance(date_val, datetime):
            date_emission = date_val.strftime("%d/%m/%Y")
        elif date_val:
            date_emission = str(date_val).strip()
        else:
            date_emission = datetime.now().strftime("%d/%m/%Y")

        tel_fourn     = data.get("Telephone_Fournisseur")
        commune       = data.get("Commune_Origine")
        quantite      = data.get("Quantite")
        tracteur      = data.get("Numero_Tracteur")
        chauffeur     = data.get("Chauffeur")
        permis        = data.get("Numero_Permis")
        tel_chauffeur = data.get("Telephone_Chauffeur")

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