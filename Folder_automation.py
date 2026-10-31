# -*- coding: utf-8 -*-
#####################################################################
#############        AUTOMATIZACION DE PROJECTS       ###############
#####################################################################
#################    JON GONDRA (08/05/2026)      ###################
#####################################################################

################## USE PYTHON 3.13 OR ABOVE #########################

import os
import shutil
from datetime import date

# ===================================================================
# ========================= ROOT PATHS ==============================
# ===================================================================

PSSE_ROOT = r"C:\PSSE"
DOC_ROOT = os.path.join(PSSE_ROOT, "00 - DOCUMENTATION")
PROJECTS_ROOT = os.path.join(PSSE_ROOT, "01 - PROJECTS")

############### Define Excel Filename-Version (if more than one .xls file available ) otherwise it will pick the remaining by default ###########
NX_WFC_Excel_FileName = "NXWFC_V4.188"
NX_WTG_Excel_FileName = "NXWTG_V8.692"
########################################################################################################

# ===================================================================
# ========================== HELPERS =================================
# ===================================================================

def mkdir(path):
    os.makedirs(path, exist_ok=True)


def safe_copy(src, dst):
    if not os.path.exists(src):
        raise FileNotFoundError(f"SOURCE NOT FOUND: {src}")
    if not os.path.exists(dst):
        raise FileNotFoundError(f"DESTINATION NOT FOUND: {dst}")
    shutil.copy(src, dst)


def find_excel(folder, startswith= None):
    """
    Finds an Excel file in a folder.

    Priority:
      1) Excel starting with 'startswith' (if provided)
      2) Any Excel file if only one exists

    Fails if:
      - No Excel files exist
      - More than one Excel exists in fallback mode
    """
    excel_files = [
        f for f in os.listdir(folder)
        if f.lower().endswith((".xls", ".xlsx", ".xlsm"))
    ]

    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in {folder}")

    # Try model-based match first
    if startswith:
        for f in excel_files:
            indx=-1
            if f.lower().startswith(startswith.lower()):
                indx = indx +1
                return os.path.join(folder, f)
            print(f"Using fallback file {(excel_files[indx])}")   

    # Fallback: only one Excel present
    if len(excel_files) == 1:
        print(
            f"Using fallback file '{excel_files[0]}'"
        )
        return os.path.join(folder, excel_files[0])

    # Ambiguous case
    raise FileNotFoundError(
        f"Multiple Excel files found in {folder}, cannot decide automatically:\n"
        + "\n".join(excel_files)
    )


def find_pdf(folder, contains):
    for f in os.listdir(folder):
        if contains.lower() in f.lower() and f.lower().endswith(".pdf"):
            return os.path.join(folder, f)
    raise FileNotFoundError(f"No PDF containing '{contains}' in {folder}")


def find_wecc_excel(folder):
    for f in os.listdir(folder):
        if "wecc" in f.lower() and f.lower().endswith((".xlsm", ".xls", ".xlsx")):
            return os.path.join(folder, f)
    raise FileNotFoundError(f"No WECC Excel found in {folder}")

# ===================================================================
# ============================== UDM ================================
# ===================================================================

def UDM():
    hoy = date.today()
    fecha = f"{hoy.year}{hoy.month:02d}{hoy.day:02d}"

    project = input("Project name: ").title()
    country = input("Country (3 letters): ").upper()
    gcv = "GCV" + input("GCV number: ").zfill(3)
    rev = input("Version: ").zfill(2)
    model = "N" + input("Model NXXX: ")
    pot = input("Pnom (kW): ")

    folder_name = f"{fecha}_{project}_{gcv}{rev}_{model}_{pot}_PSSE"
    project_root = os.path.join(PROJECTS_ROOT, folder_name)

    # ---- Project folders ----
    sharepoint = os.path.join(project_root, "00 - Project Sharepoint")
    ctrl_wtg = os.path.join(project_root, f"01 - {model}_{pot}_{gcv}_{country}_PSSE_V8.691")  ### EXCEL NAMING ####
    ctrl_wfc = os.path.join(project_root, f"02 - NX_WFC_V4.188_PSSE_{project}")
    example = os.path.join(project_root, "03 - Example")
    working = os.path.join(project_root, "04 - Working Folder")
    test = os.path.join(working, "01_Test")

    for p in [project_root, sharepoint, ctrl_wtg, ctrl_wfc, example, working, test]:
        mkdir(p)

    # ---- Documentation paths ----
    cab_models = os.path.join(DOC_ROOT, "01 - MODELS")
    cab_wtg = os.path.join(cab_models, "01 - WTG")
    cab_wfc = os.path.join(cab_models, "02 - WFC")

    cab_sim = os.path.join(DOC_ROOT, "02 - SIMULATION SCENARIO")
    cab_plb = os.path.join(cab_sim, "03 - PLB")
    cab_ener = os.path.join(cab_sim, "02 - ENERPLOT")

    # ---- Manuals ----
    wtg_manual = find_pdf(os.path.join(DOC_ROOT, "00 - MANUALS", "01 - WTG"), "WTG")
    wfc_manual = find_pdf(os.path.join(DOC_ROOT, "00 - MANUALS", "02 - WFC"), "WFC")
    safe_copy(wtg_manual, ctrl_wtg)
    safe_copy(wfc_manual, ctrl_wfc)

    # ---- WTG model (FIXED HERE) ----
    wtg_excel = find_excel(cab_wtg, NX_WTG_Excel_FileName)
    safe_copy(wtg_excel, ctrl_wtg)
    os.replace(
        os.path.join(ctrl_wtg, os.path.basename(wtg_excel)),
        os.path.join(ctrl_wtg, f"{model}_{pot}_{gcv}_{country}_PSSE_V8.691.xls"), ### EXCEL NAMING ####
    )

    for dll in os.listdir(cab_wtg):
        if dll.lower().endswith(".dll"):
            safe_copy(os.path.join(cab_wtg, dll), ctrl_wtg)

    # ---- WFC model ----
    wfc_excel = find_excel(cab_wfc, NX_WFC_Excel_FileName)
    safe_copy(wfc_excel, ctrl_wfc)
    os.replace(
        os.path.join(ctrl_wfc, os.path.basename(wfc_excel)),
        os.path.join(ctrl_wfc, f"NX_WFC_V4.188_PSSE_{project}.xls"),
    )

    for dll in os.listdir(cab_wfc):
        if dll.lower().endswith(".dll"):
            safe_copy(os.path.join(cab_wfc, dll), ctrl_wfc)

    # ---- Simulation scenario ----
    safe_copy(os.path.join(cab_sim, "01 - DYRE", "DYRE.dyr"), working)
    safe_copy(os.path.join(cab_sim, "04 - SAV_SLD", "WPP.sav"), working)
    safe_copy(os.path.join(cab_sim, "04 - SAV_SLD", "Slider_NX.sld"), working)
    safe_copy(os.path.join(cab_sim, "05 - SCRIPT", "Simulation_scenario.txt"), working)
    safe_copy(os.path.join(cab_sim, "05 - SCRIPT", "01 - PSSE v34", "01_Test.py"), working)

    for sub in os.listdir(cab_plb):
        sub_path = os.path.join(cab_plb, sub)
        if os.path.isdir(sub_path):
            for f in os.listdir(sub_path):
                if f.lower().endswith(".plb"):
                    safe_copy(os.path.join(sub_path, f), working)

    # ---- Enerplot ----
    for f in os.listdir(cab_ener):
        safe_copy(os.path.join(cab_ener, f), test)

    print("✅ UDM PROJECT CREATED SUCCESSFULLY")

# ===================================================================
# ============================== WECC ===============================
# ===================================================================

def WECC():
    hoy = date.today()
    fecha = f"{hoy.year}{hoy.month:02d}{hoy.day:02d}"

    project = input("Project name: ").title()
    gcv = "GCV" + input("GCV number: ").zfill(3)
    rev = input("Version: ").zfill(2)
    model = "N" + input("Model NXXX: ")
    pot = input("Pnom (kW): ")

    folder_name = f"{fecha}_{project}_{gcv}{rev}_{model}_{pot}_PSSE_WECC"
    project_root = os.path.join(PROJECTS_ROOT, folder_name)

    sharepoint = os.path.join(project_root, "00 - Project Sharepoint")
    param = os.path.join(project_root, "01 - Parameterization")
    working = os.path.join(project_root, "02 - Working Folder")
    test = os.path.join(working, "01_Test")

    for p in [project_root, sharepoint, param, working, test]:
        mkdir(p)

    cab_sim = os.path.join(DOC_ROOT, "02 - SIMULATION SCENARIO")
    cab_plb = os.path.join(cab_sim, "03 - PLB")
    cab_ener = os.path.join(cab_sim, "02 - ENERPLOT")

    wecc_excel = find_wecc_excel(os.path.join(DOC_ROOT, "05 - WECC"))
    safe_copy(wecc_excel, param)

    safe_copy(os.path.join(cab_sim, "05 - SCRIPT", "01 - PSSE v34", "01_Test_WECC.py"), working)

    for sub in os.listdir(cab_plb):
        sub_path = os.path.join(cab_plb, sub)
        if os.path.isdir(sub_path):
            for f in os.listdir(sub_path):
                if f.lower().endswith(".plb"):
                    safe_copy(os.path.join(sub_path, f), working)

    for f in os.listdir(cab_ener):
        safe_copy(os.path.join(cab_ener, f), test)

    print("✅ WECC PROJECT CREATED SUCCESSFULLY")

# ===================================================================
# ============================== MAIN ===============================
# ===================================================================

model_type = input("Type of model (UDM / WECC): ").upper()

if model_type == "UDM":
    UDM()
elif model_type == "WECC":
    WECC()
else:
    print("❌ INVALID MODEL TYPE")