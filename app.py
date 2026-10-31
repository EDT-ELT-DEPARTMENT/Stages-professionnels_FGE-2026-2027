# -*- coding: utf-8 -*-
"""
Plateforme de demande de stages professionnels — FGE 2026-2027
Lancer : streamlit run app.py
Le fichier Excel source doit être placé à côté de app.py.
"""
from pathlib import Path
import os
import re
import sqlite3
from datetime import date
import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
SOURCE_FILE = APP_DIR / "Liste des étudiants_2026-2027.xlsx"
DB_FILE = APP_DIR / "demandes_stages.db"
SHEET_NAME = "Feuille1"
# Seules ces promotions peuvent déposer une demande de stage rémunéré.
PROMOTIONS_STAGE_REMUNERE = {
    "M1CE", "M1ER", "M1RE", "M1ME",
    "M2CE", "M2ER", "M2RE", "M2ME",
    "M1MCIL", "M2MCIL", "MCIL2", "MCIL3",
    "ING3EI", "ING3RSE", "ING4EI", "ING4RSE", "ING5RSE",
}

st.set_page_config(
    page_title="Stages professionnels FGE 2026-2027",
    page_icon="🎓",
    layout="wide",
)

st.markdown("""
<style>
.main {background-color:#f6f8fc;}
.block-container {padding-top:1.5rem; max-width:1200px;}
.hero {padding:22px 26px; border-radius:14px; color:white;
       background:linear-gradient(115deg,#123b68,#176b87); margin-bottom:20px;}
.hero h1 {margin:0 0 8px 0; font-size:27px;}
.hero p {margin:0; opacity:.94;}
.section {font-size:20px; font-weight:700; color:#16456b;
          border-bottom:2px solid #d8e5f0; padding-bottom:7px; margin:18px 0 12px;}
.stButton>button {background:#145b83;color:white;border-radius:8px;
                  min-height:42px;font-weight:600;}
.note {background:#fff7df;padding:12px 15px;border-left:4px solid #e2a400;
       border-radius:6px;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>🎓 Plateforme de demande de stages professionnels</h1>
<p>Faculté de Génie Électrique — Année universitaire 2026-2027</p>
</div>
""", unsafe_allow_html=True)


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_matricule(value):
    """Uniformise les matricules saisis depuis Excel ou le formulaire."""
    s = clean_text(value)
    if not s:
        return ""
    # Évite les .0 ajoutés lorsqu'Excel a interprété un matricule comme nombre.
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\s+", "", s)


@st.cache_data
def load_students(path, modified_time):
    df = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=object)
    required = ["Mat. Etudiant", "Nom", "Prénom"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Colonnes absentes du fichier source : " + ", ".join(missing))
    df["_matricule_key"] = df["Mat. Etudiant"].apply(normalize_matricule)
    df = df[df["_matricule_key"] != ""].copy()
    # Une ligne par matricule : la première occurrence est conservée.
    df = df.drop_duplicates(subset=["_matricule_key"], keep="first")
    return df


def first_value(row, *columns):
    for col in columns:
        if col in row.index:
            value = clean_text(row[col])
            if value and value.lower() not in {"nan", "none"}:
                return value
    return ""


def init_db():
    with sqlite3.connect(DB_FILE) as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS demandes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date_soumission TEXT NOT NULL,
            matricule TEXT NOT NULL,
            nom TEXT, prenom TEXT, nom_arabe TEXT, prenom_arabe TEXT,
            promotion TEXT, groupe TEXT, sous_groupe TEXT,
            telephone TEXT, email TEXT,
            entreprise TEXT NOT NULL, secteur TEXT, adresse_entreprise TEXT,
            wilaya_entreprise TEXT, telephone_entreprise TEXT,
            email_entreprise TEXT, responsable TEXT, fonction_responsable TEXT,
            sujet_stage TEXT, domaine_stage TEXT, date_debut TEXT, date_fin TEXT,
            observations TEXT,
            type_demande TEXT DEFAULT 'Stage rémunéré'
        )
        """)
        # Migration automatique pour les bases créées avec une ancienne version.
        colonnes = [row[1] for row in con.execute("PRAGMA table_info(demandes)").fetchall()]
        if "type_demande" not in colonnes:
            con.execute("ALTER TABLE demandes ADD COLUMN type_demande TEXT DEFAULT 'Stage rémunéré'")
        con.commit()


def get_admin_password():
    try:
        return st.secrets.get("ADMIN_PASSWORD", os.getenv("ADMIN_PASSWORD", ""))
    except Exception:
        return os.getenv("ADMIN_PASSWORD", "")


init_db()

if not SOURCE_FILE.exists():
    st.error(
        "Fichier source introuvable. Placez **Liste des étudiants_2026-2027.xlsx** "
        "dans le même dossier que app.py."
    )
    st.stop()

try:
    students = load_students(str(SOURCE_FILE), SOURCE_FILE.stat().st_mtime)
except Exception as exc:
    st.error(f"Impossible de lire le fichier Excel : {exc}")
    st.stop()

with st.sidebar:
    st.header("Navigation")
    page = st.radio("Choisir une rubrique", ["Demande de stage", "Administration"])
    st.divider()
    st.caption("FGE • Stages professionnels • 2026-2027")

if page == "Demande de stage":
    st.markdown('<div class="section">1. Identification de l’étudiant</div>', unsafe_allow_html=True)
    matricule = st.text_input(
        "Matricule étudiant",
        placeholder="Saisissez votre matricule tel qu’il figure sur votre carte",
        max_chars=30,
    )
    student = None
    if matricule.strip():
        key = normalize_matricule(matricule)
        matches = students[students["_matricule_key"] == key]
        if matches.empty:
            st.error("Aucun étudiant ne correspond à ce matricule dans le fichier source.")
        else:
            student = matches.iloc[0]
            st.success("Matricule reconnu : vos informations ont été récupérées.")

    if student is not None:
        nom = first_value(student, "Nom")
        prenom = first_value(student, "Prénom")
        nom_ar = first_value(student, "اللقب")
        prenom_ar = first_value(student, "الإسم")
        promotion = first_value(student, "Promotion")
        groupe = first_value(student, "Groupe")
        sous_groupe = first_value(student, "Sous groupe")
        telephone = first_value(student, "N° de téléphone")
        email = first_value(student, "E-mail institutionnel", "email")
        mat_bac = first_value(student, "Mat. BAC")
        date_naiss = first_value(student, "Date de naiss.")
        lieu_naiss = first_value(student, "Lieu de naissance", "مكان الميلاد")
        sexe = first_value(student, "Sexe")

        with st.container(border=True):
            st.markdown("**Informations extraites du fichier officiel**")
            c1, c2, c3 = st.columns(3)
            c1.text_input("Nom", value=nom, disabled=True)
            c2.text_input("Prénom", value=prenom, disabled=True)
            c3.text_input("Matricule BAC", value=mat_bac, disabled=True)
            c1.text_input("Nom (arabe)", value=nom_ar, disabled=True)
            c2.text_input("Prénom (arabe)", value=prenom_ar, disabled=True)
            c3.text_input("Promotion", value=promotion, disabled=True)
            c1.text_input("Groupe", value=groupe, disabled=True)
            c2.text_input("Sous-groupe", value=sous_groupe, disabled=True)
            c3.text_input("Sexe", value=sexe, disabled=True)
            c1.text_input("Date de naissance", value=date_naiss, disabled=True)
            c2.text_input("Lieu de naissance", value=lieu_naiss, disabled=True)
            c3.text_input("Téléphone étudiant", value=telephone, disabled=True)
            st.text_input("Adresse e-mail", value=email, disabled=True)

        promotion_key = re.sub(r"\s+", "", promotion.upper())
        stage_remunere = promotion_key in PROMOTIONS_STAGE_REMUNERE

        if stage_remunere:
            st.success(
                f"Votre promotion ({promotion}) est éligible au dépôt d’une demande de stage rémunéré."
            )
            type_demande = "Stage rémunéré"
            titre_formulaire = "Demande de stage rémunéré"
            libelle_sujet = "Intitulé / sujet proposé du stage *"
            bouton = "📨 Envoyer la demande de stage rémunéré"
        else:
            st.warning(
                f"Votre promotion ({promotion or 'non renseignée'}) ne figure pas dans la liste "
                "des promotions autorisées pour les stages rémunérés. "
                "Vous êtes donc orienté(e) vers une demande de visite en entreprise."
            )
            st.info(
                "Cette orientation est automatique et repose sur la promotion indiquée "
                "dans le fichier officiel des étudiants."
            )
            type_demande = "Visite en entreprise"
            titre_formulaire = "Demande de visite en entreprise"
            libelle_sujet = "Objet / thème de la visite *"
            bouton = "📨 Envoyer la demande de visite"

        st.markdown(
            f'<div class="section">2. {titre_formulaire} — entreprise d’accueil</div>',
            unsafe_allow_html=True,
        )
        with st.form("stage_application", clear_on_submit=False):
            a1, a2 = st.columns(2)
            entreprise = a1.text_input("Nom de l’entreprise *")
            secteur = a2.text_input("Secteur d’activité")
            adresse = a1.text_area("Adresse complète de l’entreprise *", height=90)
            wilaya = a2.text_input("Wilaya / commune")
            tel_ent = a1.text_input("Téléphone de l’entreprise")
            email_ent = a2.text_input("E-mail de l’entreprise")
            responsable = a1.text_input("Nom et prénom du responsable d’accueil")
            fonction = a2.text_input("Fonction du responsable")

            if stage_remunere:
                st.markdown('<div class="section">3. Informations sur le stage</div>', unsafe_allow_html=True)
                sujet = st.text_input(libelle_sujet)
                domaine = st.text_input("Service ou domaine d’affectation")
                d1, d2 = st.columns(2)
                debut = d1.date_input("Date prévue de début", value=date.today())
                fin = d2.date_input("Date prévue de fin", value=date.today())
            else:
                st.markdown('<div class="section">3. Informations sur la visite</div>', unsafe_allow_html=True)
                sujet = st.text_input(libelle_sujet)
                domaine = st.text_input("Service / unité à visiter")
                debut = st.date_input("Date souhaitée de la visite", value=date.today())
                fin = debut
            observations = st.text_area("Observations complémentaires")
            st.caption("* Champs obligatoires.")
            submitted = st.form_submit_button(bouton, use_container_width=True)

        if submitted:
            errors = []
            if not entreprise.strip(): errors.append("Le nom de l’entreprise est obligatoire.")
            if not adresse.strip(): errors.append("L’adresse de l’entreprise est obligatoire.")
            if not sujet.strip(): errors.append("L’intitulé ou l’objet est obligatoire.")
            if stage_remunere and fin < debut:
                errors.append("La date de fin doit être postérieure ou égale à la date de début.")
            if errors:
                for err in errors:
                    st.error(err)
            else:
                with sqlite3.connect(DB_FILE) as con:
                    cur = con.execute("""
                    INSERT INTO demandes (
                      date_soumission, matricule, nom, prenom, nom_arabe, prenom_arabe,
                      promotion, groupe, sous_groupe, telephone, email,
                      entreprise, secteur, adresse_entreprise, wilaya_entreprise,
                      telephone_entreprise, email_entreprise, responsable,
                      fonction_responsable, sujet_stage, domaine_stage, date_debut,
                      date_fin, observations, type_demande
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        date.today().isoformat(), normalize_matricule(matricule),
                        nom, prenom, nom_ar, prenom_ar, promotion, groupe, sous_groupe,
                        telephone, email, entreprise.strip(), secteur.strip(), adresse.strip(),
                        wilaya.strip(), tel_ent.strip(), email_ent.strip(), responsable.strip(),
                        fonction.strip(), sujet.strip(), domaine.strip(), debut.isoformat(),
                        fin.isoformat(), observations.strip(), type_demande
                    ))
                    demande_id = cur.lastrowid
                    con.commit()
                st.success(
                    f"Votre {('demande de stage rémunéré' if stage_remunere else 'demande de visite')} "
                    f"a été enregistrée. Numéro de suivi : STG-{demande_id:05d}"
                )
                st.info("Conservez ce numéro pour le suivi administratif de votre demande.")

    else:
        st.markdown(
            '<div class="note">Saisissez votre matricule pour afficher automatiquement '
            'vos coordonnées. Les champs étudiants proviennent du fichier Excel source.</div>',
            unsafe_allow_html=True
        )

else:
    st.title("Administration — Gestion des demandes")
    password = st.text_input("Mot de passe administrateur", type="password")
    expected = get_admin_password()
    if not expected:
        st.warning("Configurez ADMIN_PASSWORD dans les secrets Streamlit ou dans les variables d’environnement.")
    elif password and password == expected:
        with sqlite3.connect(DB_FILE) as con:
            demandes = pd.read_sql_query("SELECT * FROM demandes ORDER BY id DESC", con)
        st.metric("Nombre total de demandes", len(demandes))
        if demandes.empty:
            st.info("Aucune demande enregistrée pour le moment.")
        else:
            st.dataframe(demandes, use_container_width=True, hide_index=True)
            export = demandes.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Télécharger toutes les demandes (CSV)",
                data=export,
                file_name="Demandes_stages_FGE_2026-2027.csv",
                mime="text/csv",
            )
    elif password:
        st.error("Mot de passe incorrect.")

st.divider()
st.caption("Faculté de Génie Électrique — Plateforme de demandes de stages professionnels 2026-2027")
