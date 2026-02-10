# ======================================================
# BOOT STREAMLIT MINIMAL (OBLIGATOIRE)
# ======================================================
import streamlit as st

st.set_page_config(
    page_title="Générateur de planning – Pipeline complet",
    layout="centered"
)

st.title("🗓️ Générateur de planning – Pipeline complet")
st.info("Application initialisée correctement")
st.stop()

# ======================================================
# CHARGEMENT CONTRÔLÉ DES MODULES MÉTIER
# ======================================================
if st.button("🔓 Charger l'application"):

    with st.spinner("Chargement des modules métier…"):
        import pandas as pd
        import os
        import tempfile
        from pathlib import Path
        from openpyxl import load_workbook, Workbook
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.worksheet.datavalidation import DataValidation
        from copy import copy
        import re
        from datetime import datetime
        import unicodedata
        import logging

    st.success("Modules chargés avec succès")

    # ======================================================
    # CONSTANTES (LOGIQUE MÉTIER INCHANGÉE)
    # ======================================================
    MOTIFS_PREDEFINIS = [
        "Accident de travail",
        "Arrêt Maladie",
        "Congé de Maternité",
        "Congé parental d'éducation",
        "Congés Payés",
        "Formation",
        "Mi-temps Thérapeutique",
        "Récupération",
        "Surcroît temporaire d'activité CNR ou",
        "Surcroit temporaire d’activité",
        "Absence injustifiée",
        "Congé d'ancienneté",
        "Congé de Paternité",
        "Congé de présence parentale",
        "Congé Individuel de Formation",
        "Congé sabbatique",
        "Congés Évènements Familiaux",
        "Congés sans solde",
        "Congés spécifiques/trimestriels",
        "Dans l'attente de la nomination du titulaire",
        "Détachement du titulaire sur une tâche exceptionnelle",
        "Mise à pied conservatoire",
        "Mise à pied disciplinaire",
        "Réduction temps travail femme enceinte"
    ]

    # ======================================================
    # HELPERS FICHIERS
    # ======================================================
    def save_uploaded_file(uploaded_file, suffix):
        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(
            temp_dir,
            f"{Path(uploaded_file.name).stem}_{suffix}{Path(uploaded_file.name).suffix}"
        )
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path

    # ======================================================
    # HELPERS TEXTE / NORMALISATION
    # ======================================================
    def noacc_lower(s: str) -> str:
        if s is None:
            return ""
        s = str(s)
        s = "".join(
            c for c in unicodedata.normalize("NFKD", s)
            if not unicodedata.combining(c)
        )
        s = s.replace("\u00a0", " ").replace("\ufeff", "")
        s = re.sub(r"[ \t]+", " ", s).strip().lower()
        return s

    def norm_group(s: str) -> str:
        s = noacc_lower(s).replace("\n", " ")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.replace("remplaçant", "remplacant")
        return s

    def strip_placeholders(s: str) -> str:
        if not isinstance(s, str):
            return ""
        s = s.strip()
        s = re.sub(r"^\s*(nom|pr[ée]nom)\s*[/:\-]?\s*", "", s, flags=re.IGNORECASE)
        while re.match(r"^(nom|pr[ée]nom)\b", s, flags=re.IGNORECASE):
            s = re.sub(r"^\s*(nom|pr[ée]nom)\s*[/:\-]?\s*", "", s, flags=re.IGNORECASE)
        return s.strip()

    # ======================================================
    # STRUCTURES – MAPPING (INCHANGÉ)
    # ======================================================
    def get_structure_mapping():
        data = [
         ("6750404", "EA ADAPAYSAGE BOURG"),
        ("6750405", "EA ADAPAYSAGE HAUT BUGEY"),
        ("6750309", "ESAT BELLEGARDE INDUSTRIE"),
        ("6750313", "ESAT CENTRE DE VIE RURALE"),
        ("6750307", "ESAT LA LECHERE"),
        ("6750305", "ESAT LE PENNESSUY"),
        ("6750311", "ESAT LES ATELIERS DE NIERME"),
        ("6750303", "ESAT LES BROSSES"),
        ("6750301", "ESAT LES DOMBES"),
        ("6750315", "ESAT LES TEPPES"),
        ("6750503", "FAM PRE LA TOUR"),
        ("6750504", "FAM SOUS LA ROCHE"),
        ("6750215", "FOYER BELLEVUE"),
        ("6750212", "FOYER DE TREFFORT"),
        ("6750213", "FOYER COURTES VERNOUX"),
        ("6750203", "FOYER CROIX BLANCHE"),
        ("6750201", "FOYER DE DOMAGNE"),
        ("6750210", "FOYER DE LASSIGNIEU"),
        ("6750207", "FOYER LE SOUS BOIS"),
        ("6750204", "FOYER LE VILLARDOIS"),
        ("6750202", "FOYER LES 4 VENTS"),
        ("6750209", "FOYER LES FLORALIES"),
        ("6750211", "FOYER LES PATIOS"),
        ("6750206", "FOYER LES PRES DE BROU"),
        ("6750214", "FOYER LES SOURDIERES"),
        ("6750208", "FOYER LE VAL FLEURI"),
        ("6750300", "CHAMP D'OR"),
        ("6750102", "IME GEORGES LOISEAU"),
        ("6750105", "IME L'ARMAILLOU"),
        ("6750101", "IME LE PRELION"),
        ("6750103", "IME LES SAPINS"),
        ("6750402", "EA DE BROU"),
        ("6750104", "IME SERVICE LES MUSCARIS"),
        ("6750401", "EA MAISONNETTE"),
        ("6750403", "EA MAISON DES PAYS DE L'AIN"),
        ("6750505", "MAS BELLEVUE"),
        ("6750502", "MAS LES MONTAINES"),
        ("6750501", "MAS MONTPLAISANT"),
        ("6750205", "SAVS LE PASSAGE BG EN B"),
        ("6750001", "ADAPEI DE L'AIN SIEGE SOCIAL"),
        ("6750007", "PCPE"),
        ("6750004", "POLE GEST BOURG EN BRESSE"),
        ("6750005", "POLE DE GESTION OYONNAX"),
        ("6750006", "POLE DE GESTION BELLEY"),
        ("6750003", "POLE GEST FONC TRANSVERSES"),

        # --- SAJ ---
        ("675020902", "SAJ FOYER LES FLORALIES"),
        ("675020102", "SAJ DE DOMAGNE"),
        ("675021402", "SAJ FOYER LES SOURDIERES"),
        ("675020702", "SAJ FOYER SOUS BOIS"),
        ("675021202", "SAJ FOYER DE TREFFORT"),
        ("675020402", "SAJ FOYER LE VILLARDOIS"),
        ("675021002", "SAJ FOYER DE LASSIGNIEU"),

        # --- SAVS ---
        ("675020903", "SAVS FOYER LES FLORALIES"),
        ("675021003", "SAVS FOYER DE LASSIGNIEU"),
        ("675020703", "SAVS SOUS-BOIS"),

        # --- SESSAD ---
        ("675010101", "SESSAD LES DOMBES"),
        ("675010501", "SESSAD INTERLUDE"),
        ("675010201", "SESSAD G LOISEAU"),
        ("67510301",  "SESSAD LES SAPINS"),
        ]
        mapping = {}
        for sirh_id, nom in data:
            sirh_id = str(sirh_id)
            mapping[sirh_id] = nom
            mapping[sirh_id[-3:]] = nom
        return mapping

    STRUCTURE_MAP = get_structure_mapping()

    # ======================================================
    # PIPELINE (PLACEHOLDER – TA LOGIQUE ICI)
    # ======================================================
    def traitement_pipeline_complet(fichier_brut: str) -> str:
        # 👉 ICI tu remets TOUT ton pipeline réel
        return fichier_brut

    # ======================================================
    # UI MÉTIER
    # ======================================================
    st.header("1️⃣ Pipeline complet")

    uploaded_file = st.file_uploader(
        "Importer le planning brut (.xlsx)",
        type=["xlsx"]
    )

    if uploaded_file and st.button("🚀 Lancer le pipeline"):
        raw_path = save_uploaded_file(uploaded_file, "raw")
        with st.spinner("Traitement en cours…"):
            result = traitement_pipeline_complet(raw_path)
        st.success("Traitement terminé")
