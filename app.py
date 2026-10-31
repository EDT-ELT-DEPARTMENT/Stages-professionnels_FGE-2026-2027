# -*- coding: utf-8 -*-
"""Plateforme des stages professionnels et visites en entreprise — FGE 2026-2027."""
from pathlib import Path
from datetime import date, datetime
import os, re, sqlite3, json, io
import html
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

APP_DIR = Path(__file__).resolve().parent
SOURCE_FILE = APP_DIR / "Liste des étudiants_2026-2027.xlsx"
if not SOURCE_FILE.exists():
    candidates = sorted(APP_DIR.glob("Liste des étudiants*.xlsx"))
    if candidates:
        SOURCE_FILE = candidates[0]
TEACHERS_FILE = APP_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx"
if not TEACHERS_FILE.exists():
    candidates = sorted(APP_DIR.glob("Permanents-Vacataires*.xlsx"))
    if candidates:
        TEACHERS_FILE = candidates[0]
DB_FILE = APP_DIR / "demandes_stages.db"
SHEET_NAME = "Feuille1"
PROMOTIONS_STAGE_REMUNERE = {
    "M1CE", "M1ER", "M1RE", "M1ME", "M2CE", "M2ER", "M2RE", "M2ME",
    "M1MCIL", "M2MCIL", "MCIL2", "MCIL3", "ING3EI", "ING3RSE",
    "ING4EI", "ING4RSE", "ING5RSE",
}

st.set_page_config(page_title="Stages professionnels FGE 2026-2027", page_icon="🎓", layout="wide")
st.markdown("""
<style>
.main{background:#f6f8fc}.block-container{padding-top:1.3rem;max-width:1250px}
.hero{padding:22px 26px;border-radius:14px;color:white;background:linear-gradient(115deg,#123b68,#176b87);margin-bottom:20px}
.hero h1{margin:0 0 8px;font-size:27px}.hero p{margin:0;opacity:.94}
.section{font-size:20px;font-weight:700;color:#16456b;border-bottom:2px solid #d8e5f0;padding-bottom:7px;margin:18px 0 12px}
.stButton>button{background:#145b83;color:white;border-radius:8px;min-height:42px;font-weight:600}
</style>""", unsafe_allow_html=True)
st.markdown("""<div class="hero"><h1>🎓 Plateforme de demande de stages professionnels</h1>
<p>Faculté de Génie Électrique — Année universitaire 2026-2027</p></div>""", unsafe_allow_html=True)

def clean_text(value):
    if pd.isna(value): return ""
    return str(value).strip()

def normalize_matricule(value):
    s = clean_text(value)
    if not s: return ""
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return re.sub(r"\s+", "", s)

def norm_promo(value):
    return re.sub(r"\s+", "", clean_text(value).upper())

@st.cache_data
def load_students(path, modified_time):
    df = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=object)
    required = ["Mat. Etudiant", "Nom", "Prénom", "Promotion"]
    missing = [c for c in required if c not in df.columns]
    if missing: raise ValueError("Colonnes absentes : " + ", ".join(missing))
    df["_matricule_key"] = df["Mat. Etudiant"].apply(normalize_matricule)
    return df[df["_matricule_key"] != ""].drop_duplicates("_matricule_key").copy()

def first_value(row, *columns):
    for col in columns:
        if col in row.index:
            value = clean_text(row[col])
            if value and value.lower() not in {"nan", "none"}: return value
    return ""

def student_record(row):
    return {
        "matricule": first_value(row, "Mat. Etudiant"),
        "nom": first_value(row, "Nom"), "prenom": first_value(row, "Prénom"),
        "nom_arabe": first_value(row, "اللقب"), "prenom_arabe": first_value(row, "الإسم"),
        "date_naissance": first_value(row, "Date de naiss.", "Date de naissance"),
        "lieu_naissance": first_value(row, "Lieu de naissance", "Ville de naissance"),
        "wilaya_residence": first_value(row, "Wilaya de résidence"),
        "ville_naissance": first_value(row, "Ville de naissance"),
        "wilaya_naissance": first_value(row, "Wilaya de naissance", "Wilaya de naissance.1"),
        "promotion": first_value(row, "Promotion"), "groupe": first_value(row, "Groupe"),
        "sous_groupe": first_value(row, "Sous groupe"),
        "telephone": first_value(row, "N° de téléphone"),
        "email": first_value(row, "E-mail institutionnel", "email"),
        "mat_bac": first_value(row, "Mat. BAC"),
    }

@st.cache_data
def load_permanent_teachers(path, modified_time):
    df = pd.read_excel(path, sheet_name="Permanents-Vacataires", dtype=object)
    required = ["NOM", "PRÉNOM", "Qualité"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Colonnes absentes du fichier des enseignants : " + ", ".join(missing))
    quality = df["Qualité"].astype(str).str.strip().str.casefold()
    df = df[quality.isin({"permanent", "permanente"})].copy()
    df["_label"] = (df["NOM"].fillna("").astype(str).str.strip() + " " +
                    df["PRÉNOM"].fillna("").astype(str).str.strip()).str.strip()
    df = df[df["_label"] != ""].drop_duplicates("_label")
    return sorted(df["_label"].tolist(), key=lambda x: x.casefold())

def init_db():
    with sqlite3.connect(DB_FILE) as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS demandes (
          id INTEGER PRIMARY KEY AUTOINCREMENT, date_soumission TEXT NOT NULL,
          matricule TEXT NOT NULL, nom TEXT, prenom TEXT, nom_arabe TEXT, prenom_arabe TEXT,
          promotion TEXT, groupe TEXT, sous_groupe TEXT, telephone TEXT, email TEXT,
          wilaya_residence TEXT, ville_naissance TEXT, wilaya_naissance TEXT,
          date_naissance TEXT, lieu_naissance TEXT,
          entreprise TEXT NOT NULL, secteur TEXT, adresse_entreprise TEXT,
          wilaya_entreprise TEXT, telephone_entreprise TEXT, email_entreprise TEXT,
          responsable TEXT, fonction_responsable TEXT, sujet_stage TEXT, domaine_stage TEXT,
          enseignant_superviseur TEXT, programme_stage TEXT,
          date_debut TEXT, date_fin TEXT, observations TEXT,
          type_demande TEXT DEFAULT 'Stage rémunéré',
          participants_json TEXT DEFAULT '[]', statut TEXT DEFAULT 'Brouillon - formulaire à signer',
          signed_pdf BLOB
        )""")
        cols = {r[1] for r in con.execute("PRAGMA table_info(demandes)")}
        migrations = {
            "type_demande": "TEXT DEFAULT 'Stage rémunéré'",
            "participants_json": "TEXT DEFAULT '[]'",
            "statut": "TEXT DEFAULT 'Brouillon - formulaire à signer'",
            "signed_pdf": "BLOB",
            "wilaya_residence": "TEXT", "ville_naissance": "TEXT", "wilaya_naissance": "TEXT",
            "date_naissance": "TEXT", "lieu_naissance": "TEXT",
            "enseignant_superviseur": "TEXT", "programme_stage": "TEXT",
        }
        for col, typ in migrations.items():
            if col not in cols: con.execute(f"ALTER TABLE demandes ADD COLUMN {col} {typ}")
        con.commit()

def admin_password():
    try: return st.secrets.get("ADMIN_PASSWORD", os.getenv("ADMIN_PASSWORD", ""))
    except Exception: return os.getenv("ADMIN_PASSWORD", "")

def build_request_pdf(request_data, participants):
    """Génère un PDF dont les cellules s'adaptent au contenu et passent sur plusieurs pages."""
    from reportlab.platypus import LongTable, KeepTogether

    # Mise en forme arabe : ligatures + ordre droite-à-gauche si les dépendances sont installées.
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        def arabic_display(value):
            value = str(value or "")
            if not re.search(r"[\u0600-\u06FF]", value):
                return value
            return get_display(arabic_reshaper.reshape(value))
    except Exception:
        def arabic_display(value):
            return str(value or "")

    # Police Unicode si disponible (accents et caractères arabes).
    arabic_font = "Helvetica"
    for font_path in [
        APP_DIR / "DejaVuSans.ttf",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf"),
    ]:
        if font_path.exists():
            try:
                if "FGEArabic" not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont("FGEArabic", str(font_path)))
                arabic_font = "FGEArabic"
                break
            except Exception:
                pass

    buffer = io.BytesIO()
    page_width, page_height = A4
    left_margin = right_margin = 1.35 * cm
    usable_width = page_width - left_margin - right_margin
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=right_margin, leftMargin=left_margin,
        topMargin=1.35*cm, bottomMargin=1.35*cm
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="CenterTitleFGE", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#17466b"), fontSize=14, leading=18,
        spaceAfter=3, wordWrap="CJK"
    ))
    styles.add(ParagraphStyle(
        name="SectionFGE", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=11, leading=14, textColor=colors.HexColor("#17466b"),
        spaceBefore=7, spaceAfter=6, keepWithNext=True
    ))
    styles.add(ParagraphStyle(
        name="CellFGE", fontName=arabic_font, fontSize=7.5, leading=10,
        wordWrap="CJK", splitLongWords=1, spaceAfter=0, spaceBefore=0
    ))
    styles.add(ParagraphStyle(
        name="ArabicCellFGE", parent=styles["CellFGE"], alignment=TA_RIGHT,
        fontName=arabic_font
    ))
    styles.add(ParagraphStyle(
        name="LabelFGE", parent=styles["CellFGE"], fontName="Helvetica-Bold",
        textColor=colors.HexColor("#173f5f")
    ))
    styles.add(ParagraphStyle(
        name="HeaderFGE", parent=styles["CellFGE"], fontName="Helvetica-Bold",
        textColor=colors.white, alignment=TA_CENTER, fontSize=7.2, leading=9
    ))
    styles.add(ParagraphStyle(
        name="SmallFGE", parent=styles["BodyText"], fontSize=8, leading=11,
        wordWrap="CJK", splitLongWords=1
    ))

    def para(value, style="CellFGE"):
        value = "" if value is None else str(value)
        if value.strip() == "":
            value = "—"
        has_arabic = bool(re.search(r"[\u0600-\u06FF]", value))
        value = arabic_display(value)
        # Évite qu'un &, < ou > présent dans les données casse le XML de Paragraph.
        selected_style = "ArabicCellFGE" if has_arabic and style == "CellFGE" else style
        return Paragraph(html.escape(value).replace("\n", "<br/>"), styles[selected_style])

    def table_style(header=False, font_size=7.5):
        commands = [
            ("GRID", (0,0), (-1,-1), 0.45, colors.HexColor("#9aa9b5")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]
        if header:
            commands += [
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#17466b")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("REPEAT", (0,0), (-1,0)),
            ]
        return TableStyle(commands)

    story = [
        para("UNIVERSITÉ DJILLALI LIABÈS DE SIDI BEL ABBÈS", "CenterTitleFGE"),
        para("FACULTÉ DE GÉNIE ÉLECTRIQUE", "SectionFGE"),
        para("ANNÉE UNIVERSITAIRE 2026-2027", "SmallFGE"),
        Spacer(1, 8),
        para("DEMANDE DE STAGE PROFESSIONNEL RÉMUNÉRÉ", "CenterTitleFGE"),
        Spacer(1, 8),
        para("Informations des étudiants demandeurs", "SectionFGE")
    ]

    # Les contenus sont des Paragraph : ReportLab calcule la hauteur de chaque cellule,
    # effectue les retours à la ligne et conserve les colonnes dans la largeur imprimable.
    student_widths = [1.75*cm, 2.55*cm, 4.05*cm, 3.0*cm, usable_width - 11.35*cm]
    rows = [[para(x, "HeaderFGE") for x in
             ["N° / matricule", "Rubrique", "Détail", "Rubrique", "Détail"]]]
    for i, student in enumerate(participants, 1):
        get = lambda key: str(student.get(key, "") or "")
        rows.extend([
            [para(f"{i} / {get('matricule')}"), para("Nom et prénom", "LabelFGE"),
             para((get("nom")+" "+get("prenom")).strip()), para("Nom et prénom (arabe)", "LabelFGE"),
             para((get("nom_arabe")+" "+get("prenom_arabe")).strip())],
            [para(""), para("Promotion", "LabelFGE"), para(get("promotion")),
             para("Groupe / Sous-groupe", "LabelFGE"),
             para((get("groupe")+" / "+get("sous_groupe")).strip())],
            [para(""), para("Date de naissance", "LabelFGE"), para(get("date_naissance")),
             para("Lieu de naissance", "LabelFGE"), para(get("lieu_naissance"))],
            [para(""), para("Wilaya de résidence", "LabelFGE"), para(get("wilaya_residence")),
             para("Wilaya de naissance", "LabelFGE"), para(get("wilaya_naissance"))],
            [para(""), para("Téléphone / E-mail", "LabelFGE"),
             para((get("telephone")+" / "+get("email")).strip()),
             para("Ville de naissance", "LabelFGE"), para(get("ville_naissance"))]
        ])
    student_table = LongTable(rows, colWidths=student_widths, repeatRows=1, hAlign="LEFT")
    student_table.setStyle(table_style(header=True))
    student_table.setStyle(TableStyle([
        ("BACKGROUND",(0,1),(0,-1),colors.HexColor("#edf3f8")),
        ("BACKGROUND",(1,1),(1,-1),colors.HexColor("#edf3f8")),
        ("BACKGROUND",(3,1),(3,-1),colors.HexColor("#edf3f8")),
    ]))
    story += [student_table, Spacer(1, 9), para("Encadrement du stage", "SectionFGE")]

    supervisor_rows = [
        [para("Enseignant superviseur (Permanent)", "LabelFGE"),
         para(request_data.get("superviseur", "Non renseigné"))],
        [para("Programme détaillé du stage", "LabelFGE"),
         para(request_data.get("programme", "Non renseigné"))]
    ]
    supervisor_table = Table(supervisor_rows, colWidths=[4.8*cm, usable_width-4.8*cm],
                             hAlign="LEFT")
    supervisor_table.setStyle(table_style())
    supervisor_table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#edf3f8"))
    ]))
    story += [supervisor_table, Spacer(1, 9), para("Entreprise d'accueil", "SectionFGE")]

    fields = [
        ["Entreprise", request_data.get("entreprise",""), "Secteur", request_data.get("secteur","")],
        ["Adresse", request_data.get("adresse",""), "Wilaya / commune", request_data.get("wilaya","")],
        ["Téléphone", request_data.get("tel_ent",""), "E-mail", request_data.get("email_ent","")],
        ["Responsable", request_data.get("responsable",""), "Fonction", request_data.get("fonction","")],
    ]
    enterprise_widths = [2.4*cm, 6.0*cm, 2.8*cm, usable_width-11.2*cm]
    enterprise_table = Table(
        [[para(v, "LabelFGE" if j in (0,2) else "CellFGE") for j,v in enumerate(row)]
         for row in fields],
        colWidths=enterprise_widths, hAlign="LEFT"
    )
    enterprise_table.setStyle(table_style())
    enterprise_table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#edf3f8")),
        ("BACKGROUND",(2,0),(2,-1),colors.HexColor("#edf3f8"))
    ]))
    story += [enterprise_table, Spacer(1, 9), para("Informations sur le stage", "SectionFGE")]

    details = [
        ["Sujet / thème", request_data.get("sujet","")],
        ["Service / domaine", request_data.get("domaine","")],
        ["Période", f'Du {request_data.get("debut","")} au {request_data.get("fin","")}'],
        ["Observations", request_data.get("observations","") or "—"],
    ]
    details_table = Table(
        [[para(label, "LabelFGE"), para(value)] for label,value in details],
        colWidths=[3.5*cm, usable_width-3.5*cm], hAlign="LEFT"
    )
    details_table.setStyle(table_style())
    details_table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#edf3f8"))
    ]))
    story += [
        details_table, Spacer(1, 13),
        para("AVIS FAVORABLE DE L'ÉTABLISSEMENT D'ACCUEIL", "SectionFGE"),
        para("Je soussigné(e), représentant(e) de l'établissement d'accueil, donne un avis favorable à l'accueil du/des étudiant(s) désigné(s) ci-dessus pour le stage indiqué.", "SmallFGE"),
        Spacer(1, 25),
        para("Nom, qualité et signature du responsable :", "SmallFGE"),
        Spacer(1, 25),
        para("Cachet de l'établissement :", "SmallFGE"),
        Spacer(1, 15),
        para("Après signature et cachet, le demandeur principal doit téléverser ce document sur la plateforme pour transmission au Chef de département.", "SmallFGE")
    ]
    doc.build(story)
    return buffer.getvalue()

init_db()
if not SOURCE_FILE.exists():
    st.error(f"Fichier source introuvable : {SOURCE_FILE.name}. Placez-le à côté de app.py.")
    st.stop()
try:
    students = load_students(str(SOURCE_FILE), SOURCE_FILE.stat().st_mtime)
except Exception as exc:
    st.error(f"Erreur de lecture du fichier Excel : {exc}")
    st.stop()
try:
    permanent_teachers = load_permanent_teachers(str(TEACHERS_FILE), TEACHERS_FILE.stat().st_mtime)
except Exception as exc:
    st.error(f"Erreur de lecture du fichier des enseignants permanents : {exc}")
    st.stop()
if not permanent_teachers:
    st.error("Aucun enseignant Permanent trouvé dans le fichier des enseignants.")
    st.stop()

with st.sidebar:
    st.header("Navigation")
    page = st.radio("Rubrique", ["Demande de stage", "Administration"])
    st.divider()
    st.caption("FGE • 2026-2027")

if page == "Demande de stage":
    st.markdown('<div class="section">1. Identification du demandeur principal</div>', unsafe_allow_html=True)
    matricule = st.text_input("Votre matricule étudiant", key="main_matricule",
                              placeholder="Saisissez votre matricule officiel")
    main_student = None
    if matricule.strip():
        found = students[students["_matricule_key"] == normalize_matricule(matricule)]
        if found.empty: st.error("Matricule introuvable dans le fichier source.")
        else:
            main_student = student_record(found.iloc[0])
            if norm_promo(main_student["promotion"]) not in PROMOTIONS_STAGE_REMUNERE:
                st.warning(f"Votre promotion ({main_student['promotion']}) n'est pas éligible au stage rémunéré. Vous êtes orienté(e) vers la visite en entreprise.")
                st.info("La demande groupée de stage rémunéré est réservée aux promotions autorisées. Pour une visite, utilisez la procédure correspondante.")
                st.stop()
            st.success(f"Étudiant reconnu : {main_student['nom']} {main_student['prenom']} — {main_student['promotion']}")

    if main_student:
        with st.container(border=True):
            st.markdown("**Coordonnées du demandeur principal (issues du fichier source)**")
            st.write(f"**{main_student['nom']} {main_student['prenom']}** — {main_student['matricule']}")
            st.write(f"Nom et prénom en arabe : {main_student['nom_arabe']} {main_student['prenom_arabe']}")
            st.write(f"Promotion : {main_student['promotion']} | Groupe : {main_student['groupe']} | Sous-groupe : {main_student['sous_groupe']}")
            st.write(f"Date de naissance : {main_student['date_naissance']} | Lieu de naissance : {main_student['lieu_naissance']}")
            st.write(f"Wilaya de résidence : {main_student['wilaya_residence']} | Wilaya de naissance : {main_student['wilaya_naissance']}")
            st.write(f"Téléphone : {main_student['telephone']} | E-mail : {main_student['email']}")

        st.markdown('<div class="section">2. Ajouter un ou plusieurs camarades</div>', unsafe_allow_html=True)
        st.write("Saisissez les matricules des camarades, un matricule par ligne. Tous les participants doivent appartenir aux promotions autorisées.")
        companions_text = st.text_area("Matricules des camarades (facultatif)", height=110,
                                       placeholder="Exemple :\n252538288004\n262638574206")
        participants = [main_student]
        companion_errors = []
        seen = {normalize_matricule(main_student["matricule"])}
        for raw in companions_text.splitlines():
            mk = normalize_matricule(raw)
            if not mk: continue
            if mk in seen:
                companion_errors.append(f"Matricule répété : {mk}")
                continue
            seen.add(mk)
            match = students[students["_matricule_key"] == mk]
            if match.empty:
                companion_errors.append(f"Matricule introuvable : {mk}")
                continue
            record = student_record(match.iloc[0])
            if norm_promo(record["promotion"]) not in PROMOTIONS_STAGE_REMUNERE:
                companion_errors.append(f"{record['nom']} {record['prenom']} ({record['promotion']}) n'est pas éligible au stage rémunéré.")
                continue
            participants.append(record)

        # Vérification et correction des données de chaque participant avant création du PDF.
        st.markdown('<div class="section">3. Vérifier et corriger les informations des étudiants</div>', unsafe_allow_html=True)
        st.info("Les données ci-dessous proviennent du fichier source. Vous pouvez les corriger avant la génération du PDF. Les corrections concernent uniquement cette demande.")
        corrected_participants = []
        for idx, original in enumerate(participants):
            student = dict(original)
            with st.expander(f"{idx + 1}. {original.get('nom','')} {original.get('prenom','')} — {original.get('matricule','')}", expanded=(idx == 0)):
                a1, a2 = st.columns(2)
                student["nom"] = a1.text_input("Nom", value=student.get("nom", ""), key=f"edit_nom_{idx}_{student.get('matricule','')}")
                student["prenom"] = a2.text_input("Prénom", value=student.get("prenom", ""), key=f"edit_prenom_{idx}_{student.get('matricule','')}")
                student["nom_arabe"] = a1.text_input("Nom en arabe (اللقب)", value=student.get("nom_arabe", ""), key=f"edit_nom_ar_{idx}_{student.get('matricule','')}")
                student["prenom_arabe"] = a2.text_input("Prénom en arabe (الإسم)", value=student.get("prenom_arabe", ""), key=f"edit_prenom_ar_{idx}_{student.get('matricule','')}")
                student["date_naissance"] = a1.text_input("Date de naissance", value=student.get("date_naissance", ""), key=f"edit_date_naiss_{idx}_{student.get('matricule','')}", placeholder="JJ/MM/AAAA")
                student["lieu_naissance"] = a2.text_input("Lieu de naissance", value=student.get("lieu_naissance", ""), key=f"edit_lieu_naiss_{idx}_{student.get('matricule','')}")
                student["wilaya_residence"] = a1.text_input("Wilaya de résidence", value=student.get("wilaya_residence", ""), key=f"edit_wilaya_res_{idx}_{student.get('matricule','')}")
                student["wilaya_naissance"] = a2.text_input("Wilaya de naissance", value=student.get("wilaya_naissance", ""), key=f"edit_wilaya_naiss_{idx}_{student.get('matricule','')}")
                student["ville_naissance"] = a1.text_input("Ville de naissance (source)", value=student.get("ville_naissance", ""), key=f"edit_ville_naiss_{idx}_{student.get('matricule','')}")
                student["telephone"] = a2.text_input("Téléphone", value=student.get("telephone", ""), key=f"edit_tel_{idx}_{student.get('matricule','')}")
                student["email"] = a1.text_input("E-mail", value=student.get("email", ""), key=f"edit_email_{idx}_{student.get('matricule','')}")
            corrected_participants.append(student)
        participants = corrected_participants
        # Le demandeur principal enregistré dans les colonnes dédiées doit aussi refléter ses corrections.
        main_student = participants[0]

        if companion_errors:
            for err in companion_errors: st.error(err)
        else:
            st.success(f"{len(participants)} étudiant(s) éligible(s) dans cette demande.")
        st.dataframe(pd.DataFrame([{
            "Matricule": p["matricule"], "Nom": p["nom"], "Prénom": p["prenom"],
            "Nom arabe": p.get("nom_arabe", ""), "Prénom arabe": p.get("prenom_arabe", ""),
            "Date de naissance": p.get("date_naissance", ""), "Lieu de naissance": p.get("lieu_naissance", ""),
            "Promotion": p["promotion"], "Groupe": p["groupe"], "Sous-groupe": p["sous_groupe"]
        } for p in participants]), use_container_width=True, hide_index=True)

        st.markdown('<div class="section">4. Informations sur l’entreprise et le stage</div>', unsafe_allow_html=True)
        with st.form("stage_group_form"):
            c1,c2 = st.columns(2)
            entreprise = c1.text_input("Nom de l'entreprise *")
            secteur = c2.text_input("Secteur d'activité")
            adresse = c1.text_area("Adresse complète *", height=80)
            wilaya = c2.text_input("Wilaya / commune")
            tel_ent = c1.text_input("Téléphone de l'entreprise")
            email_ent = c2.text_input("E-mail de l'entreprise")
            responsable = c1.text_input("Nom du responsable d'accueil")
            fonction = c2.text_input("Fonction du responsable")
            sujet = st.text_input("Sujet / intitulé du stage *")
            domaine = st.text_input("Service ou domaine d'affectation")
            superviseur = st.selectbox("Enseignant chargé de superviser le stage (Permanent uniquement) *",
                                       options=permanent_teachers,
                                       help="La liste est extraite du fichier des enseignants et filtrée sur la qualité Permanent.")
            programme = st.text_area("Programme détaillé du stage *",
                                     height=150,
                                     placeholder="Décrivez les activités prévues, les travaux à réaliser et les objectifs du stage.")
            d1,d2 = st.columns(2)
            debut = d1.date_input("Date prévue de début", value=date.today())
            fin = d2.date_input("Date prévue de fin", value=date.today())
            observations = st.text_area("Observations")
            create = st.form_submit_button("📄 Générer le formulaire à faire signer", use_container_width=True)

        if create:
            errors = []
            if companion_errors: errors.append("Corrigez les matricules des camarades avant de continuer.")
            if not entreprise.strip(): errors.append("Le nom de l'entreprise est obligatoire.")
            if not adresse.strip(): errors.append("L'adresse de l'entreprise est obligatoire.")
            if not sujet.strip(): errors.append("Le sujet du stage est obligatoire.")
            if not programme.strip(): errors.append("Le programme détaillé du stage est obligatoire.")
            if not superviseur.strip(): errors.append("Veuillez choisir un enseignant superviseur permanent.")
            if superviseur not in permanent_teachers: errors.append("L'enseignant choisi doit être un Permanent.")
            if fin < debut: errors.append("La date de fin doit être postérieure ou égale à la date de début.")
            if errors:
                for e in errors: st.error(e)
            else:
                req = {"entreprise":entreprise.strip(),"secteur":secteur.strip(),"adresse":adresse.strip(),
                       "wilaya":wilaya.strip(),"tel_ent":tel_ent.strip(),"email_ent":email_ent.strip(),
                       "responsable":responsable.strip(),"fonction":fonction.strip(),"sujet":sujet.strip(),
                       "domaine":domaine.strip(),
                       # Ces deux champs sont obligatoires dans le PDF et en base.
                       "superviseur":superviseur.strip(),
                       "programme":programme.strip(),
                       "debut":debut.isoformat(),"fin":fin.isoformat(),
                       "observations":observations.strip()}
                pdf = build_request_pdf(req, participants)
                with sqlite3.connect(DB_FILE) as con:
                    cur = con.execute("""
                    INSERT INTO demandes (date_soumission,matricule,nom,prenom,nom_arabe,prenom_arabe,
                    promotion,groupe,sous_groupe,telephone,email,entreprise,secteur,adresse_entreprise,
                    wilaya_entreprise,telephone_entreprise,email_entreprise,responsable,fonction_responsable,
                    sujet_stage,domaine_stage,enseignant_superviseur,programme_stage,date_debut,date_fin,observations,
                    wilaya_residence,ville_naissance,wilaya_naissance,date_naissance,lieu_naissance,type_demande,participants_json,statut)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (datetime.now().isoformat(timespec="seconds"),main_student["matricule"],
                    main_student["nom"],main_student["prenom"],main_student["nom_arabe"],main_student["prenom_arabe"],
                    main_student["promotion"],main_student["groupe"],main_student["sous_groupe"],
                    main_student["telephone"],main_student["email"],req["entreprise"],req["secteur"],req["adresse"],
                    req["wilaya"],req["tel_ent"],req["email_ent"],req["responsable"],req["fonction"],req["sujet"],
                    req["domaine"],req["superviseur"],req["programme"],req["debut"],req["fin"],req["observations"],
                    main_student["wilaya_residence"],main_student["ville_naissance"],main_student["wilaya_naissance"],
                    main_student["date_naissance"],main_student["lieu_naissance"],
                    "Stage rémunéré",json.dumps(participants,ensure_ascii=False),"Brouillon - formulaire à signer"))
                    rid = cur.lastrowid
                    con.commit()
                st.session_state["active_request_id"] = rid
                st.session_state["active_pdf"] = pdf
                st.rerun()

        # Étape de téléchargement puis téléversement du formulaire signé.
        active_id = st.session_state.get("active_request_id")
        if active_id and st.session_state.get("active_pdf"):
            st.markdown('<div class="section">5. Télécharger, faire signer et déposer le formulaire</div>', unsafe_allow_html=True)
            st.success(f"Formulaire généré — Référence STG-{active_id:05d}.")
            st.download_button("⬇️ Télécharger le formulaire PDF à imprimer et faire signer",
                               data=st.session_state["active_pdf"],
                               file_name=f"Demande_stage_FGE_STG-{active_id:05d}.pdf",
                               mime="application/pdf", use_container_width=True)
            st.info("Imprimez le formulaire, faites signer et cacheter l'avis favorable par l'établissement d'accueil, puis numérisez le document signé en PDF.")
            signed_file = st.file_uploader("Téléverser le formulaire signé (PDF uniquement)", type=["pdf"], key=f"signed_{active_id}")
            if signed_file is not None:
                if signed_file.size > 15 * 1024 * 1024:
                    st.error("Le fichier dépasse la limite de 15 Mo.")
                else:
                    st.success(f"Document prêt : {signed_file.name}")
                    if st.button("📨 Transmettre la demande signée au Chef de département", type="primary"):
                        content = signed_file.getvalue()
                        if not content.startswith(b"%PDF"):
                            st.error("Le fichier téléversé ne semble pas être un PDF valide.")
                        else:
                            with sqlite3.connect(DB_FILE) as con:
                                con.execute("UPDATE demandes SET signed_pdf=?, statut=? WHERE id=?",
                                            (sqlite3.Binary(content),"Transmise au Chef de département",active_id))
                                con.commit()
                            st.session_state.pop("active_request_id",None)
                            st.session_state.pop("active_pdf",None)
                            st.success(f"Demande STG-{active_id:05d} transmise au Chef de département.")
                            st.info("Le statut de votre demande est désormais : Transmise au Chef de département.")
elif main_student:
        st.markdown('<div class="section">2. Demande de visite en entreprise</div>', unsafe_allow_html=True)
        st.info("Votre promotion n'est pas éligible au stage rémunéré. Vous pouvez déposer une demande de visite en entreprise.")
        with st.form("visit_form"):
            v1, v2 = st.columns(2)
            v_ent = v1.text_input("Nom de l'entreprise *")
            v_sector = v2.text_input("Secteur d'activité")
            v_address = v1.text_area("Adresse complète *")
            v_wilaya = v2.text_input("Wilaya / commune")
            v_tel = v1.text_input("Téléphone")
            v_email = v2.text_input("E-mail")
            v_responsable = v1.text_input("Responsable d'accueil")
            v_sujet = v2.text_input("Objet de la visite *")
            v_date = st.date_input("Date souhaitée", value=date.today())
            v_submit = st.form_submit_button("Envoyer la demande de visite", use_container_width=True)
        if v_submit:
            if not v_ent.strip() or not v_address.strip() or not v_sujet.strip():
                st.error("Veuillez renseigner l'entreprise, l'adresse et l'objet de la visite.")
            else:
                with sqlite3.connect(DB_FILE) as con:
                    cur = con.execute("""INSERT INTO demandes
                    (date_soumission,matricule,nom,prenom,nom_arabe,prenom_arabe,promotion,groupe,sous_groupe,
                     telephone,email,entreprise,secteur,adresse_entreprise,wilaya_entreprise,telephone_entreprise,
                     email_entreprise,responsable,sujet_stage,date_debut,date_fin,type_demande,participants_json,statut)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec="seconds"),main_student["matricule"],main_student["nom"],
                     main_student["prenom"],main_student["nom_arabe"],main_student["prenom_arabe"],main_student["promotion"],
                     main_student["groupe"],main_student["sous_groupe"],main_student["telephone"],main_student["email"],
                     v_ent.strip(),v_sector.strip(),v_address.strip(),v_wilaya.strip(),v_tel.strip(),v_email.strip(),
                     v_responsable.strip(),v_sujet.strip(),v_date.isoformat(),v_date.isoformat(),
                     "Visite en entreprise",json.dumps([main_student],ensure_ascii=False),"Demande de visite reçue"))
                    visit_id=cur.lastrowid
                    con.commit()
                st.success(f"Votre demande de visite a été enregistrée. Référence : VIS-{visit_id:05d}.")

else:
    st.title("Administration — Suivi des demandes")
    pwd = st.text_input("Mot de passe administrateur", type="password")
    expected = admin_password()
    if not expected:
        st.warning("Configurez ADMIN_PASSWORD dans les secrets Streamlit.")
    elif pwd and pwd == expected:
        with sqlite3.connect(DB_FILE) as con:
            data = pd.read_sql_query("""SELECT id,date_soumission,matricule,nom,prenom,nom_arabe,prenom_arabe,promotion,
                wilaya_residence,ville_naissance,wilaya_naissance,date_naissance,lieu_naissance,enseignant_superviseur,programme_stage,
                entreprise,sujet_stage,type_demande,statut,participants_json FROM demandes ORDER BY id DESC""",con)
        st.metric("Nombre total de demandes",len(data))
        if data.empty: st.info("Aucune demande enregistrée.")
        else:
            def participants_label(raw):
                try:
                    ps=json.loads(raw or "[]")
                    return " ; ".join(
                        f'{p.get("nom","")} {p.get("prenom","")} ({p.get("matricule","")})'
                        f' | Arabe : {p.get("nom_arabe","")} {p.get("prenom_arabe","")}'
                        f' | Wilaya résidence : {p.get("wilaya_residence","")}'
                        f' | Date naissance : {p.get("date_naissance","")}'
                        f' | Lieu naissance : {p.get("lieu_naissance","")}'
                        f' | Ville naissance : {p.get("ville_naissance","")}'
                        f' | Wilaya naissance : {p.get("wilaya_naissance","")}'
                        for p in ps)
                except Exception: return ""
            data["Étudiants participants"] = data["participants_json"].apply(participants_label)
            st.dataframe(data.drop(columns=["participants_json"]),use_container_width=True,hide_index=True)
            csv=data.drop(columns=["participants_json"]).to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Exporter les demandes CSV",csv,"Demandes_stages_FGE_2026-2027.csv","text/csv")
            st.markdown("### Consulter le formulaire signé")
            selected=st.selectbox("Référence de demande",options=data["id"].tolist(),
                                  format_func=lambda x:f"STG-{x:05d}")
            with sqlite3.connect(DB_FILE) as con:
                row=con.execute("SELECT signed_pdf,statut FROM demandes WHERE id=?",(selected,)).fetchone()
            if row and row[0]:
                st.success(f"Statut : {row[1]}")
                st.download_button("⬇️ Télécharger le formulaire signé",data=bytes(row[0]),
                                   file_name=f"Demande_signee_STG-{selected:05d}.pdf",mime="application/pdf")
            elif row:
                st.warning(f"Statut : {row[1]} — le document signé n'a pas encore été déposé.")
    elif pwd: st.error("Mot de passe incorrect.")

st.divider()
st.caption("Faculté de Génie Électrique — Plateforme de stages professionnels 2026-2027")
