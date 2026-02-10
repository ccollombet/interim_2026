# ======================================================
# STREAMLIT – GÉNÉRATEUR DE PLANNING
# VERSION CLOUD SAFE – LOGIQUE MÉTIER STRICTEMENT IDENTIQUE
# ======================================================

import streamlit as st

# ======================================================
# CONFIG STREAMLIT (TOP-LEVEL ULTRA LÉGER)
# ======================================================
st.set_page_config(
    page_title="Générateur de planning – Pipeline complet",
    layout="centered"
)
st.title("🗓️ Générateur de planning – Pipeline complet")

# ======================================================
# IMPORTS LOURDS ISOLÉS
# ======================================================
def lazy_imports():
    global pd, os, tempfile, Path, re, datetime, unicodedata, logging
    global load_workbook, Workbook
    global Font, Alignment, PatternFill
    global Table, TableStyleInfo
    global DataValidation
    global copy

    import pandas as pd
    import os
    import tempfile
    from pathlib import Path
    import re
    from datetime import datetime
    import unicodedata
    import logging

    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.datavalidation import DataValidation
    from copy import copy


# ======================================================
# STRUCTURE MAP (CACHÉ – SAFE CLOUD)
# ======================================================
@st.cache_data(show_spinner=False)
def load_structure_map():
    lazy_imports()

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
    ]

    mapping = {}
    for sirh_id, nom in data:
        sirh_id = str(sirh_id)
        mapping[sirh_id] = nom
        if len(sirh_id) >= 3:
            mapping[sirh_id[-3:]] = nom

    return mapping


# ======================================================
# HELPERS UPLOAD
# ======================================================
def save_uploaded_file(uploaded_file, suffix):
    lazy_imports()
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(
        temp_dir,
        f"{Path(uploaded_file.name).stem}_{suffix}{Path(uploaded_file.name).suffix}"
    )
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path


# ======================================================
# ================== MÉTIER (INCHANGÉ) =================
# 👉 TOUTES TES FONCTIONS MÉTIER SONT ICI
# 👉 AUCUNE MODIFICATION DE LOGIQUE
# ======================================================

# ⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️⬇️
# 👉 ICI TU COLLES **EXACTEMENT** TOUTES TES
# fonctions métier :
# - helpers (noacc_lower, norm_group, etc.)
# - traitement_partie1
# - traitement_partie2
# - traitement_partie3
# - adapter_badakan_version_auto
# - etc.
#
# ⚠️ RIEN N’EST INDENTÉ ICI
# ⚠️ RIEN N’EST EXÉCUTÉ
# ⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️⬆️


# ======================================================
# ORCHESTRATEUR PIPELINE (SEUL POINT D’ENTRÉE)
# ======================================================
def traitement_pipeline_complet(fichier_brut: str) -> str:
    """
    Pipeline complet :
    1) Partie 1 : Nettoyage + planning
    2) Partie 2 : Lecture + interimaire
    Logique STRICTEMENT IDENTIQUE
    """
    lazy_imports()
    STRUCTURE_MAP = load_structure_map()

    fichier_p1 = traitement_partie1(fichier_brut)
    fichier_p2 = traitement_partie2(fichier_p1)

    return fichier_p2


def traitement_badakan_depuis_interimaire(fichier_interimaire: str) -> str:
    lazy_imports()
    return traitement_partie3(fichier_interimaire)


# ======================================================
# UI STREAMLIT
# ======================================================
st.header("1️⃣ Planning, Lecture & Intérimaire")

uploaded_file_full = st.file_uploader(
    " ",
    type=["xlsx"],
    key="upload_full"
)

if uploaded_file_full and st.button("Générer planning + lecture + intérimaire"):
    raw_path = save_uploaded_file(uploaded_file_full, "raw")

    with st.spinner("Pipeline complet en cours…"):
        try:
            fichier_final = traitement_pipeline_complet(raw_path)

            st.success("✅ Pipeline complet terminé")

            with open(fichier_final, "rb") as f:
                st.download_button(
                    "📥 Télécharger le fichier final",
                    data=f,
                    file_name=os.path.basename(fichier_final)
                )

        except Exception as e:
            st.error("❌ Erreur lors du traitement")
            st.exception(e)


st.header("2️⃣ Génération du fichier Badakan")

uploaded_file_3 = st.file_uploader(
    " ",
    type=["xlsx"],
    key="upload3"
)

if st.button("Générer le fichier Badakan"):
    if uploaded_file_3:
        source_p3 = save_uploaded_file(uploaded_file_3, "interimaire")
    else:
        source_p3 = None

    if not source_p3:
        st.error("❌ Aucun fichier interimaire disponible")
    else:
        with st.spinner("Génération du fichier Badakan…"):
            badakan = traitement_badakan_depuis_interimaire(source_p3)

        st.success("✅ Fichier Badakan généré")
        st.download_button(
            "📥 Télécharger Badakan.csv",
            data=open(badakan, "rb"),
            file_name="badakan.csv"
        )
