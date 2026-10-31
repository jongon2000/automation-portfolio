# -*- coding: utf-8 -*-
#####################################################################
#############        AUTOMATIZATION DE PROJECTS       ###############
#####################################################################
#################    JON GONDRA (14/07/2026)      ###################
#####################################################################

################## USE PYTHON 3.13 OR ABOVE #########################


#########################  GUIDE  ###################################
# Type of model (UDM / WECC): UDM

# Project name: Alpha
# Country (3 letters): CAN
# GCV number: 123
# Version: 01

# PSSE Version (34 / 35 / 36): 36
#####################################################################

# ============================================================
# AVAILABLE WTG VERSIONS
# ============================================================
# 1. NXWTG_V8.691
# 2. NXWTG_V8.692
# 3. NXWTG_V8.693
# ============================================================
# Select WTG version:

# ============================================================
# AVAILABLE WFG VERSIONS
# ============================================================
# 1. NXWTG_V8.691
# 2. NXWTG_V8.692
# 3. NXWTG_V8.693
# ============================================================


import os
import shutil
from datetime import date

# ===================================================================
# ========================= ROOT PATHS ==============================
# ===================================================================

PSSE_ROOT = r"C:\PSSE"
DOC_ROOT = os.path.join(PSSE_ROOT, "00 - DOCUMENTATION")
PROJECTS_ROOT = os.path.join(PSSE_ROOT, "01 - PROJECTS")

############### Define Excel Filename-Version #######################
############### The dlls will be taken correspondingly###############

NX_WFC_Excel_FileName = "NXWFC_V4.191"
NX_WTG_Excel_FileName = "NXWTG_V8.695"

NX_WFC_VERSION = NX_WFC_Excel_FileName.split("_V")[1]
NX_WTG_VERSION = NX_WTG_Excel_FileName.split("_V")[1]

#####################################################################

# ===================================================================
# ========================== HELPERS ================================
# ===================================================================

def mkdir(path):
    os.makedirs(path, exist_ok=True)


def safe_copy(src, dst):
    if not os.path.exists(src):
        raise FileNotFoundError(f"SOURCE NOT FOUND: {src}")

    if not os.path.exists(dst):
        raise FileNotFoundError(f"DESTINATION NOT FOUND: {dst}")

    shutil.copy(src, dst)


def find_excel(folder, startswith=None):
    """
    Finds an Excel file in a folder.

    Priority:
      1) Excel starting with 'startswith'
      2) Any Excel if only one exists
    """

    excel_files = [
        f for f in os.listdir(folder)
        if f.lower().endswith((".xls", ".xlsx", ".xlsm"))
    ]

    if not excel_files:
        raise FileNotFoundError(
            f"No Excel files found in {folder}"
        )

    if startswith:

        for f in excel_files:

            if f.lower().startswith(
                startswith.lower()
            ):
                return os.path.join(folder, f)

    if len(excel_files) == 1:

        print(
            f"Using fallback file '{excel_files[0]}'"
        )

        return os.path.join(
            folder,
            excel_files[0]
        )

    raise FileNotFoundError(
        f"Multiple Excel files found in {folder}, "
        f"cannot decide automatically:\n"
        + "\n".join(excel_files)
    )


def find_pdf(folder, contains):

    for f in os.listdir(folder):

        if (
            contains.lower() in f.lower()
            and f.lower().endswith(".pdf")
        ):
            return os.path.join(folder, f)

    raise FileNotFoundError(
        f"No PDF containing '{contains}' in {folder}"
    )


def find_wecc_excel(folder):

    for f in os.listdir(folder):

        if (
            "wecc" in f.lower()
            and f.lower().endswith(
                (".xls", ".xlsx", ".xlsm")
            )
        ):
            return os.path.join(folder, f)

    raise FileNotFoundError(
        f"No WECC Excel found in {folder}"
    )


def copy_related_dlls(
    source_folder,
    excel_file,
    destination_folder,
    psse_version
):
    """
    Copies only DLLs that belong to the selected
    Excel version AND the selected PSSE version.

    PSSE version is identified by the "_rXX" tag at
    the end of the dll name (e.g. "_r34", "_r35",
    "_r36"). This also matches sub-variants such as
    "_r34.9.6" or "_r34.9.6_IVF19", since they still
    start with "_r34".
    """

    base_name = os.path.splitext(
        os.path.basename(excel_file)
    )[0]

    version_tag = f"_r{psse_version}".lower()

    copied = 0

    for f in os.listdir(source_folder):

        f_lower = f.lower()

        if (
            f_lower.endswith(".dll")
            and f_lower.startswith(
                base_name.lower()
            )
            and version_tag in f_lower
        ):

            safe_copy(
                os.path.join(source_folder, f),
                destination_folder
            )

            copied += 1

    if copied == 0:

        raise FileNotFoundError(
            f"NO DLL FOUND for '{base_name}' "
            f"with PSSE version {psse_version} "
            f"(expected a file containing "
            f"'{version_tag}') in {source_folder}. "
            f"Check that this DLL exists, or pick a "
            f"different Excel version / PSSE version."
        )


def copy_selected_plbs(cab_plb, working):

    required_plbs = [

        # Flat start
        ("00 - FLAT START", "FR.plb"),
        # Voltage
        
        ("01 - VOLTAJE", "HVRT.plb"),
        ("01 - VOLTAJE", "HVRT2.plb"),
        ("01 - VOLTAJE", "V.plb"),

        # Frequency
        ("02 - FREQUENCY", "frerate.plb"),
        ("02 - FREQUENCY", "oF.plb"),
        ("02 - FREQUENCY", "oF_80.plb"),
        ("02 - FREQUENCY", "uF.plb"),

        # Power Factor
        ("03 - POWER FACTOR", "PF.plb"),
    ]

    for folder, filename in required_plbs:

        src = os.path.join(
            cab_plb,
            folder,
            filename
        )

        if os.path.exists(src):

            safe_copy(
                src,
                working
            )

        else:

            print(
                f"WARNING: PLB not found -> {src}"
            )


def get_test_script(
    model_type,
    psse_version
):

    scripts_root = os.path.join(
        DOC_ROOT,
        "02 - SIMULATION SCENARIO",
        "05 - SCRIPT"
    )

    script_map = {

        ("UDM", "34"): (
            "01 - PSSE v34",
            "01_Test V34.py"
        ),

        ("WECC", "34"): (
            "01 - PSSE v34",
            "01_Test_WECC V34.py"
        ),

        ("UDM", "35"): (
            "02 - PSSE v35",
            "01_Test V35.py"
        ),

        ("WECC", "35"): (
            "02 - PSSE v35",
            "01_Test_WECC V35.py"
        ),

        ("UDM", "36"): (
            "03 - PSSE v36",
            "01_Test V36(dummy).py"
        ),

        ("WECC", "36"): (
            "03 - PSSE v36",
            "01_Test_WECC V36(dummy).py"
        ),
    }

    if (
        model_type,
        psse_version
    ) not in script_map:

        raise ValueError(
            f"Unsupported configuration: "
            f"{model_type} PSSE {psse_version}"
        )

    folder, filename = script_map[
        (model_type, psse_version)
    ]

    return os.path.join(
        scripts_root,
        folder,
        filename
    )

# ===================================================================
# ============================== UDM ================================
# ===================================================================

def UDM():

    hoy = date.today()

    fecha = (
        f"{hoy.year}"
        f"{hoy.month:02d}"
        f"{hoy.day:02d}"
    )

    project = input(
        "Project name: "
    ).title()

    country = input(
        "Country (3 letters): "
    ).upper()

    gcv = (
        "GCV"
        + input(
            "GCV number: "
        ).zfill(3)
    )

    rev = input(
        "Version: "
    ).zfill(2)

    psse_version = input(
        "PSSE Version (34 / 35 / 36): "
    ).strip()

    model = (
        "N"
        + input(
            "Model NXXX: "
        )
    )

    pot = input(
        "Pnom (kW): "
    )

    folder_name = (
        f"{fecha}_{project}_{gcv}{rev}_"
        f"{model}_{pot}_PSSE")

    project_root = os.path.join(
        PROJECTS_ROOT,
        folder_name
    )

    # ---------------- FOLDERS ----------------

    sharepoint = os.path.join(
        project_root,
        "00 - Project Sharepoint"
    )

    ctrl_wtg = os.path.join(
    project_root,
    f"01 - {model}_{pot}_{gcv}_{country}_PSSE_V{NX_WTG_VERSION}"
    )

    ctrl_wfc = os.path.join(
    project_root,
    f"02 - NX_WFC_V{NX_WFC_VERSION}_PSSE_{project}"
    )

    example = os.path.join(
        project_root,
        "03 - Example"
    )

    working = os.path.join(
        project_root,
        "04 - Working Folder"
    )

    test = os.path.join(
        working,
        "01_Test"
    )

    for p in [
        project_root,
        sharepoint,
        ctrl_wtg,
        ctrl_wfc,
        example,
        working,
        test
    ]:
        mkdir(p)

    # ---------------- PATHS ----------------

    cab_models = os.path.join(
        DOC_ROOT,
        "01 - MODELS"
    )

    cab_wtg = os.path.join(
        cab_models,
        "01 - WTG"
    )

    cab_wfc = os.path.join(
        cab_models,
        "02 - WFC"
    )

    cab_sim = os.path.join(
        DOC_ROOT,
        "02 - SIMULATION SCENARIO"
    )

    cab_plb = os.path.join(
        cab_sim,
        "03 - PLB"
    )

    cab_ener = os.path.join(
        cab_sim,
        "02 - ENERPLOT"
    )

    # ---------------- MANUALS ----------------

    wtg_manual = find_pdf(
        os.path.join(
            DOC_ROOT,
            "00 - MANUALS",
            "01 - WTG"
        ),
        "WTG"
    )

    wfc_manual = find_pdf(
        os.path.join(
            DOC_ROOT,
            "00 - MANUALS",
            "02 - WFC"
        ),
        "WFC"
    )

    safe_copy(
        wtg_manual,
        ctrl_wtg
    )

    safe_copy(
        wfc_manual,
        ctrl_wfc
    )

    # ---------------- WTG ----------------

    wtg_excel = find_excel(
        cab_wtg,
        NX_WTG_Excel_FileName
    )

    safe_copy(
        wtg_excel,
        ctrl_wtg
    )

    os.replace(
        os.path.join(
            ctrl_wtg,
            os.path.basename(
                wtg_excel
            )
        ),
        os.path.join(
            ctrl_wtg,
            f"{model}_{pot}_{gcv}_{country}_PSSE_V{NX_WTG_VERSION}.xls"
        ),
    )

    copy_related_dlls(
        cab_wtg,
        wtg_excel,
        ctrl_wtg,
        psse_version
    )

    # ---------------- WFC ----------------

    wfc_excel = find_excel(
        cab_wfc,
        NX_WFC_Excel_FileName
    )

    safe_copy(
        wfc_excel,
        ctrl_wfc
    )

    os.replace(
        os.path.join(
            ctrl_wfc,
            os.path.basename(
                wfc_excel
            )
        ),
        os.path.join(
            ctrl_wfc,
            f"NX_WFC_V{NX_WFC_VERSION}_PSSE_{project}.xls"
        ),
    )

    copy_related_dlls(
        cab_wfc,
        wfc_excel,
        ctrl_wfc,
        psse_version
    )

    # ---------------- DLLS -> WORKING FOLDER ----------------
    # The simulation (.sav, DYRE, test script) runs from the
    # Working Folder, so the matching WTG and WFC DLLs must be
    # copied there too, not just into the control folders.
    # Only the DLL matching the selected PSSE version (34/35/36)
    # is copied, identified by the "_r34"/"_r35"/"_r36" tag.

    copy_related_dlls(
        cab_wtg,
        wtg_excel,
        working,
        psse_version
    )

    copy_related_dlls(
        cab_wfc,
        wfc_excel,
        working,
        psse_version
    )

    # ---------------- SCENARIO ----------------

    safe_copy(
    os.path.join(
        cab_sim,
        "01 - DYRE",
        "DYRE.dyr"
    ),
    working
    )

    safe_copy(
    os.path.join(
        cab_sim,
        "04 - SAV_SLD",
        "UDM",
        "WPP.sav"
    ),
    working
    )

    safe_copy(
    os.path.join(
        cab_sim,
        "04 - SAV_SLD",
        "UDM",
        "Slider_NX.sld"
    ),
    working
    )

    safe_copy(
        os.path.join(
            cab_sim,
            "05 - SCRIPT",
            "Simulation_scenario.txt"
        ),
        working
    )

    test_script = get_test_script(
        "UDM",
        psse_version
    )

    safe_copy(
        test_script,
        working
    )

    copy_selected_plbs(
        cab_plb,
        working
    )

    # ---------------- ENERPLOT ----------------

    for f in os.listdir(cab_ener):

        safe_copy(
            os.path.join(
                cab_ener,
                f
            ),
            test
        )

    print(
        "✅ UDM PROJECT CREATED SUCCESSFULLY"
    )

# ===================================================================
# ============================== WECC ===============================
# ===================================================================

def WECC():

    hoy = date.today()

    fecha = (
        f"{hoy.year}"
        f"{hoy.month:02d}"
        f"{hoy.day:02d}"
    )

    project = input(
        "Project name: "
    ).title()

    gcv = (
        "GCV"
        + input(
            "GCV number: "
        ).zfill(3)
    )

    rev = input(
        "Version: "
    ).zfill(2)

    psse_version = input(
        "PSSE Version (34 / 35 / 36): "
    ).strip()

    model = (
        "N"
        + input(
            "Model NXXX: "
        )
    )

    pot = input(
        "Pnom (kW): "
    )

    folder_name = (
        f"{fecha}_{project}_{gcv}{rev}_"
        f"{model}_{pot}_PSSE_WECC"
    )

    project_root = os.path.join(
        PROJECTS_ROOT,
        folder_name
    )

    sharepoint = os.path.join(
        project_root,
        "00 - Project Sharepoint"
    )

    param = os.path.join(
        project_root,
        "01 - Parameterization"
    )

    working = os.path.join(
        project_root,
        "02 - Working Folder"
    )

    test = os.path.join(
        working,
        "01_Test"
    )

    for p in [
        project_root,
        sharepoint,
        param,
        working,
        test
    ]:
        mkdir(p)

    cab_sim = os.path.join(
        DOC_ROOT,
        "02 - SIMULATION SCENARIO"
    )

    cab_plb = os.path.join(
        cab_sim,
        "03 - PLB"
    )

    cab_ener = os.path.join(
        cab_sim,
        "02 - ENERPLOT"
    )

    wecc_excel = find_wecc_excel(
        os.path.join(
            DOC_ROOT,
            "05 - WECC"
        )
    )

    safe_copy(
        wecc_excel,
        param
    )

    safe_copy(
    os.path.join(
        cab_sim,
        "01 - DYRE",
        "DYRE.dyr"
    ),
    working
    )

    safe_copy(
    os.path.join(
        cab_sim,
        "04 - SAV_SLD",
        "WECC",
        "WPP.sav"
    ),
    working
    )

    safe_copy(
    os.path.join(
        cab_sim,
        "04 - SAV_SLD",
        "WECC",
        "Slider_NX.sld"
    ),
    working
    )

    test_script = get_test_script(
        "WECC",
        psse_version
    )

    safe_copy(
        test_script,
        working
    )

    copy_selected_plbs(
        cab_plb,
        working
    )

    for f in os.listdir(cab_ener):

        safe_copy(
            os.path.join(
                cab_ener,
                f
            ),
            test
        )

    print(
        "✅ WECC PROJECT CREATED SUCCESSFULLY"
    )

# ===================================================================
# ============================== MAIN ===============================
# ===================================================================

model_type = input(
    "Type of model (UDM / WECC): "
).upper()

if model_type == "UDM":
    UDM()

elif model_type == "WECC":
    WECC()

else:
    print(
        "❌ INVALID MODEL TYPE"
    )