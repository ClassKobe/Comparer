import os
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from rapidfuzz import fuzz, process

OUTPUT_FILE = "privv_importable.csv"

ITEM_MAP = {
    "102": "Project Feasibility Analysis",
    "107": "Enviromental Assessment",
    "108": "Hazardous Material Survey",
    "111": "Site Survey",
    "201": "Financing Costs",
    "202": "Owner controlled insurance",
    "204": "Builders Risk",
    "303": "Project Manager",
    "304": "Project Manager Reimbursables",
    "308": "Testing & Inspections",
    "310": "Commissioning Agent",
    "315": "Pre-Opening Expenses",
    "401": "Project Architect- A/E",
    "402": "Project Architect- ASR",
    "403": "Project Architect Reimbursables",
    "408": "Geotechnical Engineer",
    "416": "Technology Design",
    "418": "Peer Review",
    "419": "Peer Review Reimbursables",
    "501": "Pre Construction",
    "503.1": "West Side Renovation",
    "503.3": "Facility Work",
    "503.4": "Deferred Maintenance",
    "503.5": "2024 North/East Scope",
    "503.6": "2025 North/East Scope",
    "601": "Signage",
    "602": "Sponsorship Signage & Branding",
    "603": "Food Beverage and Retail Equipment",
    "604": "Facility Operations",
    "605": "Team Operations",
    "606": "Specialty Lighting/ Sports Lighting",
    "608": "Furnishings",
    "609": "Fixed & Loose Seating",
    "703": "Audio System",
    "704": "Video Displays",
    "706": "Video Production",
    "707": "Broadcast Cabling",
    "708": "TV Displays",
    "710": "Point of Sale(POS)",
    "711": "Security",
    "712": "Audio Visual",
    "713": "DAS & WIFI",
    "802": "Impact Fees",
    "901": "Owner Contingency",
    "316": "General Store / Grounds",
}

CLASSIFICATION_RULES = [
    {"code": "403",   "weight": 1000, "fields": ["vendor", "desc"], "keywords": ["populous"],               "note": "Populous always maps to 403"},
    {"code": "316",   "weight": 1000, "fields": ["vendor", "desc"], "keywords": ["general store"],          "note": "general store is 316"},
    {"code": "416",   "weight": 1000, "fields": ["vendor", "desc"], "keywords": ["anthony james partners"], "note": "Anthony James Partners → Technology Design"},
    {"code": "418",   "weight": 500,  "fields": ["ctype"],           "keywords": ["internal commitment"],   "note": "Internal Commitments → Peer Review (418)"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["cmt lab"],                "note": "CMT Laboratories"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hillis"],                 "note": "Hillis-Carnes"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hillis-carnes"],          "note": "Hillis-Carnes (hyphenated)"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hillis carnes"],          "note": "Hillis-Carnes (space)"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hrv conformance"],        "note": "HRV Conformance Verification"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hrv"],                    "note": "HRV (short match)"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["simpson gumpertz"],       "note": "Simpson Gumpertz & Heger"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["testing engineers"],      "note": "Testing Engineers & Consultants"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["trc environmental"],      "note": "TRC Environmental"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["intertek"],               "note": "Intertek PSI"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["professional service industries"], "note": "PSI (full name)"},
    {"code": "308",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["psi"],                    "note": "PSI (abbreviation)"},
    {"code": "310",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["envinity"],               "note": "Envinity"},
    {"code": "310",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["sustainable building"],   "note": "Sustainable Building Partners"},
    {"code": "111",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hrg"],                    "note": "HRG"},
    {"code": "102",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["nations group"],          "note": "Nations Group → Feasibility / PM"},
    {"code": "108",   "weight": 200,  "fields": ["desc"],           "keywords": ["asbestos"],               "note": "Asbestos work → 108"},
    {"code": "108",   "weight": 200,  "fields": ["desc"],           "keywords": ["hazardous"],              "note": "Hazardous material → 108"},
    {"code": "108",   "weight": 150,  "fields": ["desc"],           "keywords": ["pcb"],                    "note": "PCB survey → 108"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["precis engineering"],     "note": "Precis Engineering"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["weber murphy fox"],       "note": "Weber Murphy Fox"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["stahl sheaffer"],         "note": "Stahl Sheaffer Engineering"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["howe engineers"],         "note": "Howe Engineers"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["thornton tomasetti"],     "note": "Thornton Tomasetti"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["florida cosulting"],      "note": "Florida Cosulting (sic) → A/E"},
    {"code": "401",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["florida consulting"],     "note": "Florida Consulting → A/E"},
    {"code": "408",   "weight": 200,  "fields": ["desc"],           "keywords": ["geotechnical"],           "note": "Geotechnical keyword"},
    {"code": "408",   "weight": 150,  "fields": ["desc"],           "keywords": ["soil"],                   "note": "Soil testing → 408"},
    {"code": "416",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["labella"],                "note": "Labella Associates → Technology Design"},
    {"code": "501",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["barton malow"],           "note": "Barton Malow AECOM Hunt"},
    {"code": "501",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bmahajv"],                "note": "BMAHAJV JV"},
    {"code": "501",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bmaha"],                  "note": "BMAHA (partial)"},
    {"code": "501",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["triventure"],             "note": "Triventure → Pre Construction / West Side"},
    {"code": "503.1", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["triventure"],             "note": "Triventure West Side work"},
    {"code": "503.3", "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["leonard s fiore"],        "note": "Leonard S Fiore → Facility Work"},
    {"code": "503.3", "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["gm mccrossin"],           "note": "GM McCrossin → Facility Work"},
    {"code": "503.3", "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["burke"],                  "note": "Burke & Company → Facility Work"},
    {"code": "503.3", "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["mccarl"],                 "note": "S P McCarl → Facility Work"},
    {"code": "601",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["architectural design"],   "note": "AD/S → Signage"},
    {"code": "601",   "weight": 100,  "fields": ["desc"],           "keywords": ["signage"],                "note": "Signage keyword → 601"},
    {"code": "602",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["blair image"],            "note": "Blair Image Elements → Branding"},
    {"code": "602",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["byoglobe"],               "note": "ByoGlobe → Branding"},
    {"code": "602",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["paramount"],              "note": "Paramount & Co → Branding"},
    {"code": "602",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["sc vinyl"],               "note": "SC Vinyl → Sponsorship Signage"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["iowa rotocast"],          "note": "Iowa Rotocast Plastics"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["iowa rotoplastics"],      "note": "Iowa Rotoplastics Inc (alt spelling)"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["iowa roto"],              "note": "Iowa Roto* (catch-all for both spellings)"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["gallery carts"],          "note": "Gallery Carts / Carts of Colorado"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["singer equipment"],       "note": "Singer Equipment"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["carnegie equipment"],     "note": "Carnegie Equipment"},
    {"code": "603",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["11400"],                  "note": "11400 LLC"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["inproduction"],           "note": "InProduction → Facility Ops"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["salient"],                "note": "Salient Engineered Products"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["spicer welding"],         "note": "Spicer Welding & Fabrication"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["cramalot"],               "note": "Cram-A-Lot"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["henry schein"],           "note": "Henry Schein"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bhfoto"],                 "note": "B&H Photo"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bhphoto"],                "note": "B&H Photo variant"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bh foto"],                "note": "B&H Photo (Foto spelling)"},
    {"code": "604",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["bh photo"],               "note": "B&H Photo"},
    {"code": "712",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["dagostino"],              "note": "Dagostino Electronic Services → AV"},
    {"code": "713",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["eplus"],                  "note": "ePlus Technology → DAS & WIFI"},
    {"code": "713",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["wesco"],                  "note": "WESCO Distribution → DAS & WIFI"},
    {"code": "706",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["university video"],       "note": "University Video Services → 706"},
    {"code": "704",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["insane impact"],          "note": "Insane Impact LLC → Video Displays"},
    {"code": "704",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["insane/impact"],          "note": "Insane/Impact (slash variant) → Video Displays"},
    {"code": "704",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["insane"],                 "note": "Insane* catch-all → Video Displays"},
    {"code": "604",   "weight": 50,   "fields": ["desc"],           "keywords": ["penn state general stores"], "note": "PSU General Stores → Facility Ops fallback"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["comunale"],               "note": "S.A. Comunale → Facility Work"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["liberty fire"],           "note": "Liberty Fire → Facility Work"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["thyssenkrupp"],           "note": "ThyssenKrupp Elevator → Facility Work"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["tk elevator"],            "note": "TK Elevator → Facility Work"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["tri-state valve"],        "note": "Tri-State Valve → Facility Work"},
    {"code": "503.3", "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["tristate valve"],         "note": "Tri-State Valve → Facility Work"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["standing stone"],         "note": "Standing Stone Consulting → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["action cleaning"],        "note": "Action Cleaning → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["emerson"],                "note": "Emerson → Facility Ops"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["supra office"],           "note": "Supra Office Solutions → 604"},
    {"code": "316",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["tuckahoe turf"],          "note": "Tuckahoe Turf → 316"},
    {"code": "316",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["dryject"],                "note": "Dryject → 316"},
    {"code": "316",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["aercore"],                "note": "Aer-Core → 316"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["george trailers"],        "note": "George Trailers → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["fulmers storage"],        "note": "Fulmer's Storage Trailers → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["fulmer"],                 "note": "Fulmer's Storage → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["s  c operations"],        "note": "S&C Operations → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["alvarado"],               "note": "Alvarado (turnstiles) → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["triple d enterprise"],    "note": "Triple D Enterprise → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["gxc"],                    "note": "GXC → 604"},
    {"code": "604",   "weight": 150,  "fields": ["vendor", "desc"], "keywords": ["west penn power"],        "note": "West Penn Power → 604"},
    {"code": "608",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["4topps"],                 "note": "4Topps LLC"},
    {"code": "608",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["link interiors"],         "note": "Link Interiors"},
    {"code": "608",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["sedia systems"],          "note": "Sedia Systems"},
    {"code": "703",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["clair global"],           "note": "Clair Global → Audio"},
    {"code": "704",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["insane impact"],          "note": "Insane Impact → Video"},
    {"code": "706",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["vizrt"],                  "note": "Vizrt → Video Production"},
    {"code": "708",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["dobil"],                  "note": "Dobil Laboratories → TV Displays"},
    {"code": "708",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["king systems"],           "note": "King Systems → TV Displays"},
    {"code": "316",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["johnson controls"],       "note": "Johnson Controls → 316"},
    {"code": "316",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["hummer turfgrass"],       "note": "Hummer Turfgrass → 316"},
    {"code": "316",   "weight": 200,  "fields": ["vendor", "desc"], "keywords": ["cdw"],                    "note": "CDW → 316"},
    {"code": "401",   "weight": 100,  "fields": ["desc"],           "keywords": ["design agreement"],       "note": "Design Agreement → A/E"},
    {"code": "401",   "weight": 100,  "fields": ["desc"],           "keywords": ["design services"],        "note": "Design Services IC → A/E"},
    {"code": "501",   "weight": 100,  "fields": ["desc"],           "keywords": ["preconstruction"],        "note": "Pre-construction keyword"},
    {"code": "501",   "weight": 100,  "fields": ["desc"],           "keywords": ["pre construction"],       "note": "Pre-construction keyword"},
    {"code": "503.3", "weight": 100,  "fields": ["desc"],           "keywords": ["construction services"],  "note": "Construction services → 503.3"},
    {"code": "308",   "weight": 100,  "fields": ["desc"],           "keywords": ["inspection"],             "note": "Inspection keyword → 308"},
    {"code": "308",   "weight": 100,  "fields": ["desc"],           "keywords": ["testing"],                "note": "Testing keyword → 308"},
    {"code": "310",   "weight": 100,  "fields": ["desc"],           "keywords": ["commissioning"],          "note": "Commissioning keyword → 310"},
    {"code": "604",   "weight": 100,  "fields": ["desc"],           "keywords": ["operations"],             "note": "Operations keyword → 604"},
    {"code": "604",   "weight": 100,  "fields": ["desc"],           "keywords": ["facility"],               "note": "Facility keyword → 604"},
]

VENDOR_ALIASES = {
    # eBuilder name (lowercased)                                        : PRIVV canonical name(s)
    "insane impact llc":                                                  "Insane/Impact",
    "insane impact":                                                      "Insane/Impact",
    "iowa rotocast plastics inc":                                         "Iowa Rotoplastics Inc",
    "iowa rotocast plastics":                                             "Iowa Rotoplastics Inc",
    "hillis-carnes engineering associates, inc.":                         "Hillis Carnes",
    "hillis-carnes engineering associates":                               "Hillis Carnes",
    "professional service industries inc":                                "Intertek PSI",
    "professional service industries":                                    "Intertek PSI",
    "hrv conformance verification associ":                                "HRV Conformance Verification Associates",
    "hrv conformance verification associates":                            "HRV Conformance Verification Associates",

    # Barton Malow / JV — eBuilder shows the full JV name, PRIVV has multiple vendor entries
    "barton malow aecom hunt alexander":                                  "Triventure",
    "barton malow alexander ii joint venture": [
        "1-CM-GMP-000925000-Barton Malow - Alexander - II Joint Venture -David Peck (dlp50)",
        "Barton Malow - Alexander - II Joint Venture ^",
    ],
    "bmahajv": [
        "1-CM-GMP-000925000-Barton Malow - Alexander - II Joint Venture -David Peck (dlp50)",
        "Barton Malow - Alexander - II Joint Venture ^",
    ],
    "triventure":                                                         "barton malow aecom hunt alexander",

    # Florida Consulting — two spellings on eBuilder side, multiple PRIVV entries
    "florida cosulting": [
        "Florida Cosulting",
        "Florida Consulting, LLC^",
        "1-P Design Agreement_Agreement_Florida Consulting, LLC",
    ],
    "florida consulting": [
        "Florida Cosulting",
        "Florida Consulting, LLC^",
        "1-P Design Agreement_Agreement_Florida Consulting, LLC",
    ],

    # Johnson Controls
    "johnson contrl security solutions":                                  "JOHNSON CONTROLS US HOLDINGS LLC JOHNSON CONTROLS SECURITY SOLUTIONS",
    "johnson controls us holdings llc johnson controls security solutions": "Johnson Control Security Solutions",
    "SGH":                                                                 "Simpson Gumpertz & Heger, Inc.^",
    "AD/S":                                         "ARCHITECTURAL DESIGN & SIGNS INC^",
}

DEFAULT_CODE = "901"
FUZZY_THRESHOLD = 40


def normalize(text):
    lowered = str(text).lower()

    if "nations group" in lowered and "suzan" in lowered:
        return "nations group"

    text = lowered
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def resolve_alias(vendor_raw: str):
    key = vendor_raw.strip().lower()
    if not key:
        return vendor_raw
    if key in VENDOR_ALIASES:
        return VENDOR_ALIASES[key]
    for alias_key, canonical in VENDOR_ALIASES.items():
        if alias_key and (alias_key in key or key in alias_key):
            return canonical
    return vendor_raw


def resolve_alias_list(vendor_raw: str) -> list:
    result = resolve_alias(vendor_raw)
    if isinstance(result, list):
        return result
    return [result]


def convert_date(val):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return ""
        val = str(val).strip()
        if "." in val:
            m, d, y = val.split(".")
        elif "/" in val:
            m, d, y = val.split("/")
        else:
            return val
        return f"{int(m)}/{int(d)}/{int(y)}"
    except Exception:
        return val


def assign_code_weighted(row, lookup: dict) -> tuple:
    vendor = normalize(row.get("Company", ""))
    desc   = normalize(row.get("Description", ""))
    ctype  = normalize(row.get("Commitment Type", ""))

    field_values = {"vendor": vendor, "desc": desc, "ctype": ctype}

    best_score = -1
    best_code  = None
    best_note  = "no match"

    for rule in CLASSIFICATION_RULES:
        matched = all(
            any(kw in field_values.get(f, "") for f in rule["fields"])
            for kw in rule["keywords"]
        )
        if matched and rule["weight"] > best_score:
            best_score = rule["weight"]
            best_code  = rule["code"]
            best_note  = rule.get("note", "")

    if lookup:
        match = process.extractOne(
            vendor,
            lookup.keys(),
            scorer=fuzz.token_set_ratio,
            score_cutoff=FUZZY_THRESHOLD,
        )
        if match:
            matched_key, score, _ = match
            lookup_weight = 150
            if lookup_weight > best_score:
                best_score = lookup_weight
                best_code  = lookup[matched_key]
                best_note  = f"privv lookup (fuzzy {score:.0f}%): {matched_key}"

    if best_code is None:
        return DEFAULT_CODE, "default fallback"

    return best_code, best_note


# ---------------------------------------------------------------------------
# GUI STATE
# ---------------------------------------------------------------------------
ebuilder_files: list = []
privv_file: str = ""


def select_ebuilder_files():
    global ebuilder_files
    files = list(filedialog.askopenfilenames(
        title="Select one or more eBuilder CSV files",
        filetypes=[("CSV Files", "*.csv")]
    ))
    if files:
        ebuilder_files = files
        names = ", ".join(os.path.basename(f) for f in files)
        ebuilder_label.config(text=f"✅ {len(files)} file(s): {names}", fg="green")
    else:
        if not ebuilder_files:
            ebuilder_label.config(text="No eBuilder files selected.", fg="gray")


def select_privv_file():
    global privv_file
    f = filedialog.askopenfilename(
        title="Select PRIVV reference CSV file",
        filetypes=[("CSV Files", "*.csv")]
    )
    if f:
        privv_file = f
        privv_label.config(text=f"✅ Selected: {os.path.basename(f)}", fg="green")
    else:
        if not privv_file:
            privv_label.config(text="No PRIVV file selected.", fg="gray")


def clear_ebuilder_files():
    global ebuilder_files
    ebuilder_files = []
    ebuilder_label.config(text="No eBuilder files selected.", fg="gray")
    log("eBuilder file list cleared.")


def log(msg: str):
    output_box.insert(tk.END, msg + "\n")
    output_box.see(tk.END)


# ---------------------------------------------------------------------------
# COMPARISON
# ---------------------------------------------------------------------------
def run_comparison():
    if not ebuilder_files:
        messagebox.showerror("Error", "Select at least one eBuilder file first.")
        return
    if not privv_file:
        messagebox.showerror("Error", "Select the PRIVV file first.")
        return

    log("\n── Running Vendor Comparison ──────────────────────────────────────")

    privv_df = pd.read_csv(privv_file, dtype=str).fillna("")
    if "Vendor" not in privv_df.columns or "Amount" not in privv_df.columns:
        messagebox.showerror("Error", "PRIVV file must have 'Vendor' and 'Amount' columns.")
        return

    privv_df["_amount_num"] = pd.to_numeric(
        privv_df["Amount"].str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)

    data_list = []
    for f in ebuilder_files:
        df = pd.read_csv(f, header=0, dtype=str).fillna("")
        df = df[~(df == "").all(axis=1)]
        data_list.append(df)
    eb_raw = pd.concat(data_list, ignore_index=True).drop_duplicates()

    if "#" in eb_raw.columns:
        eb_raw = eb_raw.rename(columns={"#": "ID"})

    col_names = [
        "ID", "Description", "Company", "Date", "Status",
        "Commitment Type", "Commitment Amount",
        "Current Commitment", "Projected Commitment",
        "Actuals Approved", "Remaining Balance",
    ]
    if not all(c in eb_raw.columns for c in ["Description", "Current Commitment"]):
        while len(eb_raw.columns) < len(col_names):
            eb_raw[len(eb_raw.columns)] = ""
        eb_raw.columns = col_names[:len(eb_raw.columns)]

    # NOTE: Company is intentionally NOT overwritten here so that both
    # Description and Company are checked independently against PRIVV vendors.

    # Drop summary/totals rows
    eb_raw = eb_raw[eb_raw["Description"].str.strip() != ""].reset_index(drop=True)

    eb_raw["_commit_num"] = pd.to_numeric(
        eb_raw["Current Commitment"].str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)

    # Hard-coded override: map eBuilder JV description to PRIVV vendor name BMAHAJV
    eb_raw.loc[
        eb_raw["Company"].str.contains("Barton Malow - Alexander - II Joint Venture", case=False, na=False) |
        eb_raw["Description"].str.contains("Barton Malow - Alexander - II Joint Venture", case=False, na=False),
        "Company"
    ] = "BMAHAJV"
    eb_raw.loc[
        eb_raw["Company"].str.contains("J.V. MANUFACTURING, INC-CRAMALOT", case=False, na=False) |
        eb_raw["Description"].str.contains("J.V. MANUFACTURING, INC-CRAMALOT", case=False, na=False),
        "Company"
    ] = "Cram-A-Lot"   
    eb_raw.loc[
    eb_raw["Company"].str.contains("Architectural Design", case=False, na=False, regex=False) |
    eb_raw["Description"].str.contains("Architectural Design", case=False, na=False, regex=False),
    "Company"
] = "AD/S"

    eb_raw.loc[
        eb_raw["Company"].str.contains("Simpson Gumpertz", case=False, na=False, regex=False) |
        eb_raw["Description"].str.contains("Simpson Gumpertz", case=False, na=False, regex=False),
        "Company"
    ] = "SGH"  



#J.V. MANUFACTURING, INC-CRAMALOT
#
    # Build resolved-vendor column checking BOTH Description and Company
    # Build resolved-vendor column checking BOTH Description and Company
    # for alias matches, deduplicating the resulting names.
    eb_raw["_resolved_vendors"] = eb_raw.apply(
        lambda r: list({
            name
            for field in [r.get("Description", ""), r.get("Company", "")]
            for name in resolve_alias_list(str(field))
        }),
        axis=1,
    )

    comparison_rows = []
    vendors_seen = {}

    for _, row in privv_df.iterrows():
        vendor_raw = str(row.get("Vendor", "")).strip()
        if not vendor_raw:
            continue
        vendor_key = normalize(vendor_raw)
        if vendor_key not in vendors_seen:
            vendors_seen[vendor_key] = vendor_raw

    for vendor_key, vendor_raw in vendors_seen.items():
        privv_mask = privv_df["Vendor"].apply(lambda v: normalize(str(v)) == vendor_key)
        privv_total = privv_df.loc[privv_mask, "_amount_num"].sum()

        # Check vendor_key against both Description and Company columns
        eb_direct = (
            eb_raw["Description"].apply(lambda d: vendor_key in normalize(str(d))) |
            eb_raw["Company"].apply(lambda c: vendor_key in normalize(str(c)))
        )
        eb_alias = eb_raw["_resolved_vendors"].apply(
            lambda names: any(normalize(n) == vendor_key for n in names)
        )
        eb_mask = eb_direct | eb_alias

        eb_total = eb_raw.loc[eb_mask, "_commit_num"].sum()
        eb_matches = int(eb_mask.sum())

        delta = eb_total - privv_total

        if eb_total == 0 and privv_total == 0:
            status = "⚪ No Data"
        elif abs(delta) < 0.01:
            status = "✅ Match"
        elif eb_total > privv_total:
            status = "🔼 eBuilder Higher"
        else:
            status = "🔽 eBuilder Lower"

        comparison_rows.append({
            "Vendor":         vendor_raw,
            "PRIVV Total":    privv_total,
            "eBuilder Total": eb_total,
            "Delta":          delta,
            "eB Matches":     eb_matches,
            "Status":         status,
        })

    if not comparison_rows:
        log("No vendors found in PRIVV file.")
        return


    all_matched_indices = set()
    for vendor_key, vendor_raw in vendors_seen.items():
        eb_direct = (
            eb_raw["Description"].apply(lambda d: vendor_key in normalize(str(d))) |
            eb_raw["Company"].apply(lambda c: vendor_key in normalize(str(c)))
        )
        eb_alias = eb_raw["_resolved_vendors"].apply(
            lambda names: any(normalize(n) == vendor_key for n in names)
        )
        eb_mask = eb_direct | eb_alias
        all_matched_indices.update(eb_raw[eb_mask].index.tolist())

    unmatched_eb = eb_raw[~eb_raw.index.isin(all_matched_indices)]
    unmatched_total = unmatched_eb["_commit_num"].sum()

    if unmatched_total != 0:
        psu_key = "penn state office of physical plant"
        existing = next(
            (r for r in comparison_rows if r["Vendor"].strip().lower() == psu_key),
            None
        )
        if existing:
            existing["eBuilder Total"] += unmatched_total
            existing["Delta"]           = existing["eBuilder Total"] - existing["PRIVV Total"]
            existing["eB Matches"]     += len(unmatched_eb)
            if abs(existing["Delta"]) < 0.01:
                existing["Status"] = "✅ Match"
            elif existing["eBuilder Total"] > existing["PRIVV Total"]:
                existing["Status"] = "🔼 eBuilder Higher"
            else:
                existing["Status"] = "🔽 eBuilder Lower"
        else:
            comparison_rows.append({
                "Vendor":         "Penn State Office of Physical Plant",
                "PRIVV Total":    0.0,
                "eBuilder Total": unmatched_total,
                "Delta":          unmatched_total,
                "eB Matches":     len(unmatched_eb),
                "Status":         "🔼 eBuilder Higher",
            })


    comparison_rows.sort(key=lambda r: abs(r["Delta"]), reverse=True)
    show_comparison_window(comparison_rows)
    log(f"Comparison complete — {len(comparison_rows)} vendor(s) analyzed.")


def show_comparison_window(rows):
    win = tk.Toplevel(root)
    win.title("Vendor Amount Comparison: PRIVV vs eBuilder")
    win.geometry("1000x520")

    total_privv    = sum(r["PRIVV Total"]    for r in rows)
    total_ebuilder = sum(r["eBuilder Total"] for r in rows)
    total_delta    = total_ebuilder - total_privv
    direction      = "eBuilder is HIGHER" if total_delta > 0 else ("eBuilder is LOWER" if total_delta < 0 else "Exact Match")

    summary_frame = tk.Frame(win, bg="#1e1e2e", padx=10, pady=6)
    summary_frame.pack(fill="x")
    tk.Label(
        summary_frame,
        text=f"  PRIVV Total: ${total_privv:,.2f}     eBuilder Total: ${total_ebuilder:,.2f}     "
             f"Net Delta: ${total_delta:+,.2f}     Overall: {direction}",
        bg="#1e1e2e", fg="white", font=("Consolas", 10, "bold"), anchor="w"
    ).pack(fill="x")

    search_frame = tk.Frame(win, padx=8, pady=4)
    search_frame.pack(fill="x")
    tk.Label(search_frame, text="Filter vendor:").pack(side="left")
    search_var = tk.StringVar()
    tk.Entry(search_frame, textvariable=search_var, width=30).pack(side="left", padx=6)

    cols = ("Vendor", "PRIVV Total", "eBuilder Total", "Delta", "eB Matches", "Status")
    tree_frame = tk.Frame(win)
    tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    tree = ttk.Treeview(
        tree_frame, columns=cols, show="headings",
        yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=20
    )
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    col_widths = {"Vendor": 240, "PRIVV Total": 130, "eBuilder Total": 140, "Delta": 130, "eB Matches": 90, "Status": 160}
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=col_widths.get(c, 120), anchor="e" if c not in ("Vendor", "Status") else "w")

    tree.tag_configure("match",  background="#d4edda")
    tree.tag_configure("higher", background="#fff3cd")
    tree.tag_configure("lower",  background="#f8d7da")
    tree.tag_configure("nodata", background="#e2e3e5")

    def populate_tree(filter_text=""):
        tree.delete(*tree.get_children())
        ft = filter_text.strip().lower()
        for r in rows:
            if ft and ft not in r["Vendor"].lower():
                continue
            tag = "nodata"
            if "Match"   in r["Status"]: tag = "match"
            elif "Higher" in r["Status"]: tag = "higher"
            elif "Lower"  in r["Status"]: tag = "lower"
            tree.insert("", "end", values=(
                r["Vendor"],
                f"${r['PRIVV Total']:,.2f}",
                f"${r['eBuilder Total']:,.2f}",
                f"${r['Delta']:+,.2f}",
                r["eB Matches"],
                r["Status"],
            ), tags=(tag,))

    populate_tree()
    search_var.trace_add("write", lambda *_: populate_tree(search_var.get()))

    vsb.pack(side="right",  fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    legend = tk.Frame(win, padx=8, pady=4)
    legend.pack(fill="x")
    for color, label in [
        ("#d4edda", "✅ Match"),
        ("#fff3cd", "🔼 eBuilder Higher"),
        ("#f8d7da", "🔽 eBuilder Lower"),
        ("#e2e3e5", "⚪ No Data"),
    ]:
        tk.Label(legend, text=f"  {label}  ", bg=color, relief="solid", bd=1, padx=6, pady=2).pack(side="left", padx=4)

    def export_comparison():
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            title="Save comparison as..."
        )
        if path:
            pd.DataFrame(rows).to_csv(path, index=False)
            messagebox.showinfo("Exported", f"Saved to {path}")

    tk.Button(win, text="Export to CSV", command=export_comparison, bg="#0066cc", fg="white", padx=8).pack(pady=6)


# ---------------------------------------------------------------------------
# MAIN PROCESS
# ---------------------------------------------------------------------------
def run_process():
    if not ebuilder_files:
        messagebox.showerror("Error", "Select at least one eBuilder file.")
        return
    if not privv_file:
        messagebox.showerror("Error", "Select the PRIVV reference file.")
        return

    log("Starting process...")

    log(f"Combining {len(ebuilder_files)} eBuilder file(s)...")
    data_list = []
    for f in ebuilder_files:
        df = pd.read_csv(f, header=0, dtype=str).fillna("")
        df = df[~(df == "").all(axis=1)]
        log(f"  Loaded: {os.path.basename(f)} ({len(df)} rows)")
        data_list.append(df)

    data = pd.concat(data_list, ignore_index=True).drop_duplicates()
    log(f"Combined total (after dedup): {len(data)} rows")

    if "#" in data.columns:
        data = data.rename(columns={"#": "ID"})

    col_names = [
        "ID", "Description", "Company", "Date", "Status",
        "Commitment Type", "Commitment Amount",
        "Current Commitment", "Projected Commitment",
        "Actuals Approved", "Remaining Balance",
    ]
    if not all(c in data.columns for c in ["Description", "Current Commitment"]):
        while len(data.columns) < len(col_names):
            data[len(data.columns)] = ""
        data.columns = col_names[:len(data.columns)]

    # run_process keeps Company = Description intentionally for output formatting
   # data["Company"] = data["Description"]

    data = data[data["Description"].str.strip() != ""].reset_index(drop=True)

    ebuilder_df = data.copy()

    privv_df = pd.read_csv(privv_file, dtype=str).fillna("")
    lookup: dict = {}
    for _, row in privv_df.iterrows():
        company = normalize(row.get("Vendor", ""))
        code    = str(row.get("Code", "")).strip()
        if company and code:
            lookup[company] = code

    results = ebuilder_df.apply(lambda r: assign_code_weighted(r, lookup), axis=1)
    ebuilder_df["Cost Code"]  = results.apply(lambda x: x[0])
    ebuilder_df["Match Note"] = results.apply(lambda x: x[1])

    log("Classification summary:")
    for code, grp in ebuilder_df.groupby("Cost Code"):
        label = ITEM_MAP.get(code, "Unknown")
        log(f"  {code} ({label}): {len(grp)} rows")

    unmatched = ebuilder_df[ebuilder_df["Cost Code"] == DEFAULT_CODE]
    if not unmatched.empty:
        log(f"\nRows that fell through to default ({DEFAULT_CODE}):")
        for _, r in unmatched.iterrows():
            log(f"  Vendor: {r['Company']} | Desc: {r['Description'][:60]}")

    ebuilder_df["Final Type"] = ebuilder_df["Cost Code"].apply(
        lambda c: "Change Order" if c in ("315", "316") else "Original"
    )

    ebuilder_df["Amount"] = pd.to_numeric(
        ebuilder_df["Current Commitment"].str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0)

    po_df     = ebuilder_df[ebuilder_df["Commitment Type"] == "Purchase Order"]
    non_po_df = ebuilder_df[ebuilder_df["Commitment Type"] != "Purchase Order"]

    po_exempt = po_df[ po_df["Cost Code"].isin(["315", "316"])]
    po_normal = po_df[~po_df["Cost Code"].isin(["315", "316"])]

    grouped_po = (
        po_normal.groupby(["Company", "Cost Code", "Final Type"], as_index=False)
        .agg({"Amount": "sum", "Description": "first", "Date": "first"})
    )

    final_df = pd.concat(
        [
            grouped_po,
            po_exempt[["Company", "Cost Code", "Final Type", "Amount", "Description", "Date"]],
            non_po_df[["Company", "Cost Code", "Final Type", "Amount", "Description", "Date"]],
        ],
        ignore_index=True,
    )

    output = pd.DataFrame()
    output["Code"]             = final_df["Cost Code"]
    output["Vendor"]           = final_df["Company"]
    output["Description"]      = final_df["Description"]
    output["PO Number"]        = ""
    output["Item Description"] = final_df["Cost Code"].map(ITEM_MAP).fillna("")
    output["Date Committed"]   = final_df["Date"].apply(convert_date)
    output["Type"]             = final_df["Final Type"]
    output["Amount"]           = final_df["Amount"]

    output.to_csv(OUTPUT_FILE, index=False)
    log(f"\n✅ DONE: {OUTPUT_FILE} created with {len(output)} rows")


# ---------------------------------------------------------------------------
# GUI LAYOUT
# ---------------------------------------------------------------------------
root = tk.Tk()
root.title("EBuilder → PRIVV Processor (Weighted Classifier + Fuzzy Match)")
root.resizable(True, True)

frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill="x")

tk.Button(frame, text="Select eBuilder Files", command=select_ebuilder_files, width=22).grid(
    row=0, column=0, padx=5, pady=4, sticky="w"
)
tk.Button(frame, text="Clear", command=clear_ebuilder_files, width=8, fg="red").grid(
    row=0, column=1, padx=2, pady=4, sticky="w"
)
ebuilder_label = tk.Label(frame, text="No eBuilder files selected.", fg="gray", anchor="w")
ebuilder_label.grid(row=0, column=2, padx=5, sticky="w")

tk.Button(frame, text="Select PRIVV File", command=select_privv_file, width=22).grid(
    row=1, column=0, padx=5, pady=4, sticky="w"
)
privv_label = tk.Label(frame, text="No PRIVV file selected.", fg="gray", anchor="w")
privv_label.grid(row=1, column=2, padx=5, sticky="w")

btn_frame = tk.Frame(root)
btn_frame.pack(pady=4)


tk.Button(
    btn_frame, text="Compare Vendors", command=run_comparison,
    bg="#0066cc", fg="white", width=18, padx=6
).pack(side="left", padx=8)

output_box = tk.Text(root, height=18, width=100)
output_box.pack(pady=10, padx=10)

root.mainloop()