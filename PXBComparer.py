import os
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from rapidfuzz import fuzz, process
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    "AJP":                                          ["ANTHONY JAMES PARTNERS LLC ^"," Anthony James"],
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


# Threshold used by find_fuzzy_privv_matches below. Kept as a module-level
# constant (not hardcoded per-vendor) so it can be tuned in one place if
# it's ever too loose/strict across the whole PRIVV file.
PRIVV_VENDOR_FUZZY_THRESHOLD = 87


def find_fuzzy_privv_matches(vendor_label: str, privv_df, threshold: int = PRIVV_VENDOR_FUZZY_THRESHOLD):
    """Return every PRIVV row whose Vendor field is the same company as
    vendor_label, even if spelled slightly differently (e.g. 'Dobil
    Laboratories' vs 'Dobil Laboratories Inc'), AND every row where the
    company name only shows up inside the Description field instead of
    Vendor (PRIVV files sometimes log a generic bucket like 'Penn State
    Office of Physical Plant' as the Vendor and put the real company name
    at the start of the Description, e.g. 'Dobil Laboratories Inc - TV
    monitors purchased...').

    This intentionally does NOT hardcode any vendor name. A row counts as
    a match if, after normalize()-ing the relevant text, either:
      - the Vendor field is identical / a substring match / fuzzy-similar
        to vendor_label (token_sort_ratio >= threshold), or
      - the Description field CONTAINS vendor_label as a substring, or
        scores >= threshold on fuzz.partial_ratio (which finds the best
        matching substring rather than comparing the whole sentence, so a
        short vendor name embedded in a longer description still matches).
    """
    target = normalize(vendor_label)
    if not target:
        return privv_df.iloc[0:0]

    has_vendor = "Vendor" in privv_df.columns
    has_desc = "Description" in privv_df.columns
    if not has_vendor and not has_desc:
        return privv_df.iloc[0:0]

    def _vendor_match(vn: str) -> bool:
        if not vn:
            return False
        if vn == target or vn in target or target in vn:
            return True
        return fuzz.token_sort_ratio(vn, target) >= threshold

    def _desc_match(dn: str) -> bool:
        if not dn:
            return False
        if target in dn:
            return True
        return fuzz.partial_ratio(target, dn) >= threshold

    def _is_match(row):
        if has_vendor and _vendor_match(normalize(str(row.get("Vendor", "")))):
            return True
        if has_desc and _desc_match(normalize(str(row.get("Description", "")))):
            return True
        return False

    mask = privv_df.apply(_is_match, axis=1)
    return privv_df.loc[mask]


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


def normalize_date_to_tuple(val):
    """Parse both eBuilder (MM.DD.YYYY) and Privv (M/D/YYYY) date formats
    into a comparable (year, month, day) tuple. Returns None on failure."""
    try:
        val = str(val).strip()
        if not val:
            return None
        if "." in val:
            parts = val.split(".")
        elif "/" in val:
            parts = val.split("/")
        else:
            return None
        if len(parts) != 3:
            return None
        m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
        return (y, m, d)
    except Exception:
        return None


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


# Preferred display order for the well-known eBuilder columns. This is no
# longer a filter -- _rows_to_entries below keeps EVERY real column from the
# eBuilder row, it just uses this list to decide what order to show the
# familiar ones in before tacking on anything extra.
ENTRY_COLS = [
    "ID", "Description", "Company", "Date", "Status",
    "Commitment Type", "Commitment Amount", "Current Commitment",
    "Projected Commitment", "Actuals Approved", "Remaining Balance",
]


def _rows_to_entries(df_subset):
    """Convert a slice of the eBuilder dataframe into a list of plain dicts
    representing the FULL line item (every real column from the eBuilder
    file) so the drill-down entries window can show all of it.

    Internal helper columns that we add ourselves during processing
    (e.g. "_commit_num", "_resolved_vendors") are excluded since they're
    not part of the original eBuilder row -- everything else is kept.
    """
    real_cols = [c for c in df_subset.columns if not str(c).startswith("_")]
    ordered = [c for c in ENTRY_COLS if c in real_cols] + \
              [c for c in real_cols if c not in ENTRY_COLS]
    return df_subset[ordered].to_dict("records")


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

    _current = pd.to_numeric(
        eb_raw["Current Commitment"].str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)
    _commitment = pd.to_numeric(
        eb_raw["Commitment Amount"].str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)
    # Use Current Commitment when it has a value; fall back to Commitment Amount
    # for rows where Current Commitment is zero or was blank/N/A (e.g. Pending agreements)
    eb_raw["_commit_num"] = _current.where(_current != 0, other=_commitment)

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
    eb_raw.loc[
        eb_raw["Company"].str.contains("Anthony James", case=False, na=False, regex=False) |
        eb_raw["Description"].str.contains("Anthony James", case=False, na=False, regex=False),
        "Company"
    ] = "AJP" 


 



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


    # ── Pre-pass: Clair Global rows must have "Clair Global" in BOTH ───────────
    #    Company AND Description to count as a Clair Global eBuilder entry.
    #    Rows where it only appears in one field are forced to PSU OPP and
    #    excluded from ALL matching loops.
    clair_in_either = (
        eb_raw["Company"].str.contains("Clair Global", case=False, na=False, regex=False) |
        eb_raw["Description"].str.contains("Clair Global", case=False, na=False, regex=False)
    )
    clair_both = (
        eb_raw["Company"].str.contains("Clair Global", case=False, na=False, regex=False) &
        eb_raw["Description"].str.contains("Clair Global", case=False, na=False, regex=False)
    )
    # Rows that mention Clair Global in only one field → pre-assign to PSU OPP,
    # remove them so NO matching loop ever sees them.
    clair_partial_indices = eb_raw[clair_in_either & ~clair_both].index
    clair_partial_total   = eb_raw.loc[clair_partial_indices, "_commit_num"].sum()
    clair_partial_entries = _rows_to_entries(eb_raw.loc[clair_partial_indices])
    if len(clair_partial_indices):
        log(f"  Clair Global partial match ({len(clair_partial_indices)} row(s), "
            f"${clair_partial_total:,.2f}) → Penn State OPP")
    eb_raw = eb_raw[~eb_raw.index.isin(clair_partial_indices)].copy()

    # Rebuild _resolved_vendors after stripping Clair Global partial rows
    eb_raw["_resolved_vendors"] = eb_raw.apply(
        lambda r: list({
            name
            for field in [r.get("Description", ""), r.get("Company", "")]
            for name in resolve_alias_list(str(field))
        }),
        axis=1,
    )

    for vendor_key, vendor_raw in vendors_seen.items():
        privv_mask = privv_df["Vendor"].apply(lambda v: normalize(str(v)) == vendor_key)
        privv_total = privv_df.loc[privv_mask, "_amount_num"].sum()

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
            "_entries":       _rows_to_entries(eb_raw.loc[eb_mask]),
        })

    if not comparison_rows:
        log("No vendors found in PRIVV file.")
        return

    # ── Main matching loop ───────────────────────────────────────────────────
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

    unmatched_eb = eb_raw[~eb_raw.index.isin(all_matched_indices)].copy()

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _add_to_psu(amount, n_rows, entries=None):
        entries = entries or []
        psu_key = "penn state office of physical plant"
        existing = next(
            (r for r in comparison_rows if r["Vendor"].strip().lower() == psu_key),
            None,
        )
        if existing:
            existing["eBuilder Total"] += amount
            existing["Delta"]           = existing["eBuilder Total"] - existing["PRIVV Total"]
            existing["eB Matches"]     += n_rows
            existing.setdefault("_entries", []).extend(entries)
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
                "eBuilder Total": amount,
                "Delta":          amount,
                "eB Matches":     n_rows,
                "Status":         "🔼 eBuilder Higher",
                "_entries":       list(entries),
            })

    def _add_to_new_company(vendor_label, amount, n_rows, entries=None):
        comparison_rows.append({
            "Vendor":         vendor_label,
            "PRIVV Total":    0.0,
            "eBuilder Total": amount,
            "Delta":          amount,
            "eB Matches":     n_rows,
            "Status":         "🔼 eBuilder Higher",
            "_entries":       list(entries or []),
        })

    # Flush the pre-pass Clair Global partial rows into PSU OPP now that
    # comparison_rows is fully built and _add_to_psu is defined.
    if clair_partial_total != 0:
        _add_to_psu(clair_partial_total, len(clair_partial_indices), clair_partial_entries)

    # ── Build PSU OPP date set from PRIVV ────────────────────────────────────
    # Collect every committed-entry date that PRIVV has for Penn State OPP so
    # we can match unmatched eBuilder rows by date as well as description.
    psu_privv_key = "penn state office of physical plant"
    psu_privv_dates: set = set()
    if "Date Committed" in privv_df.columns:
        date_col = "Date Committed"
    elif "Date" in privv_df.columns:
        date_col = "Date"
    else:
        date_col = None

    if date_col:
        psu_privv_mask = privv_df["Vendor"].apply(
            lambda v: normalize(str(v)) == psu_privv_key
        )
        for raw_date in privv_df.loc[psu_privv_mask, date_col]:
            tup = normalize_date_to_tuple(raw_date)
            if tup:
                psu_privv_dates.add(tup)

    if psu_privv_dates:
        log(f"  PSU OPP PRIVV committed dates loaded: {len(psu_privv_dates)} date(s)")

    def _eb_date_matches_psu(eb_date_val) -> bool:
        """Return True if this eBuilder row's date appears in the PSU OPP PRIVV dates."""
        tup = normalize_date_to_tuple(eb_date_val)
        return tup is not None and tup in psu_privv_dates

    # ── Build a flat list of PRIVV vendor names for fuzzy matching ───────────
    # Used below to check whether an "unmatched" eBuilder row actually belongs
    # to a known PRIVV vendor that just wasn't caught by the exact-match pass.
    privv_vendor_names: list = [
        str(row.get("Vendor", "")).strip()
        for _, row in privv_df.iterrows()
        if str(row.get("Vendor", "")).strip()
    ]
    # Deduplicate while preserving order
    _seen_pv: set = set()
    privv_vendor_names_dedup: list = []
    for _pv in privv_vendor_names:
        _k = normalize(_pv)
        if _k not in _seen_pv:
            _seen_pv.add(_k)
            privv_vendor_names_dedup.append(_pv)
    privv_vendor_names = privv_vendor_names_dedup

    UNMATCHED_FUZZY_THRESHOLD = 72  # higher bar than classification fuzzy (40)

    # Vendors handled exclusively by _fix_zero_eb post-pass.
    # Prevent the fuzzy loop from grabbing their rows and mis-grouping them
    # (e.g. "DYWIDAG SYSTEMS INC" scoring high against "KING SYSTEMS" because
    # both share the token "SYSTEMS").
    FUZZY_SKIP_KEYWORDS = {"dywidag", "thornton tomasetti"}

    def _fuzzy_find_privv_vendor(company_str: str, description_str: str):
        """
        Return (privv_vendor_raw, score) if either the Company field or the
        Description field fuzzy-matches a known PRIVV vendor above the threshold,
        OR if the vendor name is found as a substring inside the description.
        Returns (None, 0) if no match is found.
        """
        if not privv_vendor_names:
            return None, 0

        # Skip vendors that are handled exclusively by the _fix_zero_eb post-pass.
        # Without this guard, shared generic tokens like "SYSTEMS" can cause
        # e.g. "DYWIDAG SYSTEMS INC" to fuzzy-match "KING SYSTEMS".
        combined_lower = (company_str + " " + description_str).lower()
        if any(kw in combined_lower for kw in FUZZY_SKIP_KEYWORDS):
            return None, 0

        candidates = []

        # 1. Fuzzy match Company field against all PRIVV vendor names
        if company_str.strip():
            m = process.extractOne(
                normalize(company_str),
                [normalize(v) for v in privv_vendor_names],
                scorer=fuzz.token_set_ratio,
                score_cutoff=UNMATCHED_FUZZY_THRESHOLD,
            )
            if m:
                matched_norm, score, idx = m
                candidates.append((privv_vendor_names[idx], score, "Company fuzzy"))

        # 2. Fuzzy match Description field against all PRIVV vendor names
        if description_str.strip():
            m2 = process.extractOne(
                normalize(description_str),
                [normalize(v) for v in privv_vendor_names],
                scorer=fuzz.token_set_ratio,
                score_cutoff=UNMATCHED_FUZZY_THRESHOLD,
            )
            if m2:
                matched_norm2, score2, idx2 = m2
                candidates.append((privv_vendor_names[idx2], score2, "Description fuzzy"))

        # 3. Substring check: is any PRIVV vendor name contained in the description?
        desc_norm = normalize(description_str)
        for pv in privv_vendor_names:
            pv_norm = normalize(pv)
            if pv_norm and pv_norm in desc_norm:
                candidates.append((pv, 100, "vendor-in-description"))

        # 4. Substring check: is any PRIVV vendor name contained in the company field?
        comp_norm = normalize(company_str)
        for pv in privv_vendor_names:
            pv_norm = normalize(pv)
            if pv_norm and pv_norm in comp_norm:
                candidates.append((pv, 100, "vendor-in-company"))

        if not candidates:
            return None, 0

        # Pick the highest-scoring candidate
        best = max(candidates, key=lambda x: x[1])
        return best[0], best[1]   # (privv_vendor_raw, score)

    def _add_to_existing_privv_vendor(privv_vendor_raw: str, amount: float,
                                       n_rows: int, entries=None):
        """
        Find the comparison_row that belongs to privv_vendor_raw and add the
        eBuilder amount to it.  If for some reason it is not found, fall back
        to creating a new entry so nothing is silently lost.
        """
        target_key = normalize(privv_vendor_raw)
        for r in comparison_rows:
            if normalize(r["Vendor"]) == target_key:
                r["eBuilder Total"] += amount
                r["eB Matches"]     += n_rows
                r.setdefault("_entries", []).extend(entries or [])
                r["Delta"] = r["eBuilder Total"] - r["PRIVV Total"]
                r["Status"] = (
                    "✅ Match"           if abs(r["Delta"]) < 0.01 else
                    "🔼 eBuilder Higher" if r["eBuilder Total"] > r["PRIVV Total"] else
                    "🔽 eBuilder Lower"
                )
                return True
        # Fallback: vendor found via fuzzy but somehow not in comparison_rows yet
        _add_to_new_company(privv_vendor_raw, amount, n_rows, entries)
        return False

    # ── Route each remaining unmatched row ───────────────────────────────────
    # Priority order:
    #   1. Fuzzy / substring match to a known PRIVV vendor → credit that vendor
    #   2. "Penn" in eBuilder Description OR date matches PSU OPP PRIVV → PSU OPP
    #   3. Everything else → new company entry
    for _, urow in unmatched_eb.iterrows():
        row_amount   = float(urow["_commit_num"])
        description  = str(urow.get("Description", ""))
        eb_date      = str(urow.get("Date", ""))
        company_str  = str(urow.get("Company", "")).strip()
        vendor_label = company_str or description[:40].strip()
        entries      = _rows_to_entries(unmatched_eb.loc[[urow.name]])

        # ── Step 1: fuzzy / substring match to a known PRIVV vendor ──────────
        privv_match, match_score = _fuzzy_find_privv_vendor(company_str, description)
        if privv_match:
            log(f"  Unmatched fuzzy→PRIVV (score {match_score}) "
                f"'{vendor_label}' → '{privv_match}'")
            _add_to_existing_privv_vendor(privv_match, row_amount, 1, entries)
            continue

        # ── Step 2: Penn State OPP heuristics ────────────────────────────────
        penn_in_desc = "penn" in description.lower()
        date_matches = _eb_date_matches_psu(eb_date)

        if penn_in_desc or date_matches:
            reason = []
            if penn_in_desc:  reason.append("'Penn' in desc")
            if date_matches:  reason.append(f"date match ({eb_date})")
            log(f"  Unmatched {' & '.join(reason)} → PSU OPP: {vendor_label!r}")
            _add_to_psu(row_amount, 1, entries)
            continue

        # ── Step 3: genuinely new vendor ─────────────────────────────────────
        if not vendor_label:
            vendor_label = "Unknown (eBuilder Only)"
        log(f"  Unmatched → new entry: {vendor_label!r}")
        _add_to_new_company(vendor_label, row_amount, 1, entries)


    # ── Post-pass: fold "The Pennsylvania State University" into PSU OPP ─────
    # Scan all final comparison_rows; any row whose Vendor name matches
    # "Pennsylvania State University" (case-insensitive) has its eBuilder Total
    # merged into Penn State OPP and is then removed from the table.
    AJP_NAMES = {"anthony james partners llc ^", "anthony james partners llc", "anthony james partners", "ajp"}
    ajp_eb_mask = (
        eb_raw["Company"].str.contains("Anthony James", case=False, na=False, regex=False) |
        eb_raw["Description"].str.contains("Anthony James", case=False, na=False, regex=False)
    )
    ajp_eb_total   = eb_raw.loc[ajp_eb_mask, "_commit_num"].sum()
    ajp_eb_entries = _rows_to_entries(eb_raw.loc[ajp_eb_mask])
    ajp_merge_rows = [r for r in comparison_rows if r["Vendor"].strip().lower() in AJP_NAMES]
    if ajp_merge_rows or ajp_eb_total:
        for mr in ajp_merge_rows:
            log(f"  Post-pass: merging '{mr['Vendor']}' "
                f"(${ajp_eb_total:,.2f}, {int(ajp_eb_mask.sum())} row(s)) → Penn State OPP")
        if ajp_eb_total:
            _add_to_psu(ajp_eb_total, int(ajp_eb_mask.sum()), ajp_eb_entries)
        comparison_rows[:] = [
            r for r in comparison_rows
            if r["Vendor"].strip().lower() not in AJP_NAMES
        ]

    def _fix_zero_eb(privv_variants: set, eb_keyword: str, display_name: str = None):
        """
        Post-pass correction for vendors whose eBuilder rows were not caught by
        the main matching loop (e.g. vendor absent from PRIVV entirely, or
        matched via _add_to_new_company before _fix_zero_eb runs).

        Changes vs. original:
          - The `eBuilder Total == 0` guard is removed: we always write the
            correct total, whether the row was previously $0 or already partially
            credited by _add_to_new_company.
          - If no comparison_row matches privv_variants at all (vendor is
            completely absent from PRIVV), a new row is created via
            _add_to_new_company so the amount is never silently lost.
        """
        mask = (
            eb_raw["Company"].str.contains(eb_keyword, case=False, na=False, regex=False) |
            eb_raw["Description"].str.contains(eb_keyword, case=False, na=False, regex=False)
        )
        eb_total = eb_raw.loc[mask, "_commit_num"].sum()
        if not eb_total:
            return

        eb_count   = int(mask.sum())
        eb_entries = _rows_to_entries(eb_raw.loc[mask])

        # Locate an existing row — could be from the PRIVV main loop *or* from a
        # prior _add_to_new_company call (vendor label like "DYWIDAG SYSTEMS INC^"
        # normalises into the privv_variants set).
        target = next(
            (r for r in comparison_rows if r["Vendor"].strip().lower() in privv_variants),
            None,
        )

        if target is not None:
            # Always overwrite: removes the old `== 0` guard that caused the bug.
            target["eBuilder Total"] = eb_total
            target["eB Matches"]     = eb_count
            target["_entries"]       = eb_entries
            target["Delta"]          = target["eBuilder Total"] - target["PRIVV Total"]
            target["Status"] = (
                "✅ Match"           if abs(target["Delta"]) < 0.01 else
                "🔼 eBuilder Higher" if target["eBuilder Total"] > target["PRIVV Total"] else
                "🔽 eBuilder Lower"
            )
            log(f"  Post-pass fix: '{target['Vendor']}' eBuilder total → ${eb_total:,.2f}")
        else:
            # Vendor has no PRIVV row at all — create one so the amount is visible.
            label = display_name or eb_keyword
            _add_to_new_company(label, eb_total, eb_count, eb_entries)
            log(f"  Post-pass fix: created new entry '{label}' → ${eb_total:,.2f}")

    _fix_zero_eb({"dywidag systems inc^", "dywidag systems inc"},
                 "DYWIDAG", display_name="DYWIDAG SYSTEMS INC")
    _fix_zero_eb({"thornton tomasetti inc^", "thornton tomasetti inc", "thornton tomasetti"},
                 "Thornton Tomasetti")

    PSU_NAMES = {"the pennsylvania state university", "pennsylvania state university"}
    psu_merge_rows = [
        r for r in comparison_rows
        if r["Vendor"].strip().lower() in PSU_NAMES
    ]
    if psu_merge_rows:
        for mr in psu_merge_rows:
            merged_amount  = mr["eBuilder Total"]
            merged_matches = mr["eB Matches"]
            merged_entries = mr.get("_entries", [])
            log(f"  Post-pass: merging '{mr['Vendor']}' "
                f"(${merged_amount:,.2f}, {merged_matches} row(s)) → Penn State OPP")
            _add_to_psu(merged_amount, merged_matches, merged_entries)
        # Remove the now-merged rows from the display list
        comparison_rows[:] = [
            r for r in comparison_rows
            if r["Vendor"].strip().lower() not in PSU_NAMES
        ]

    # ── Recompute PRIVV Total with fuzzy matching ───────────────────────────
    # The PRIVV Total set earlier (vendors_seen exact-match loop) misses rows
    # where the company name only appears in Description (e.g. Vendor =
    # "Penn State Office of Physical Plant", Description = "Dobil
    # Laboratories Inc - ..."), or where the Vendor spelling varies slightly
    # ("Dobil Laboratories" vs "Dobil Laboratories Inc"). find_fuzzy_privv_matches
    # is the same lookup used by the entries drilldown window, so recomputing
    # every row's total with it here keeps the summary table and the
    # drilldown in agreement instead of only fixing it after a click.
    for r in comparison_rows:
        fuzzy_matches = find_fuzzy_privv_matches(r["Vendor"], privv_df)
        if "_amount_num" in fuzzy_matches.columns:
            fuzzy_total = fuzzy_matches["_amount_num"].sum()
        else:
            fuzzy_total = pd.to_numeric(
                fuzzy_matches.get("Amount", pd.Series(dtype=str)).astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            ).sum()
        if abs(fuzzy_total - r["PRIVV Total"]) > 0.005:
            log(f"  Fuzzy PRIVV total fix: '{r['Vendor']}' "
                f"${r['PRIVV Total']:,.2f} → ${fuzzy_total:,.2f}")
        r["PRIVV Total"] = fuzzy_total
        r["Delta"] = r["eBuilder Total"] - r["PRIVV Total"]
        r["Status"] = (
            "✅ Match"           if abs(r["Delta"]) < 0.01 else
            "🔼 eBuilder Higher" if r["eBuilder Total"] > r["PRIVV Total"] else
            "🔽 eBuilder Lower"
        )

    comparison_rows.sort(key=lambda r: abs(r["Delta"]), reverse=True)

    # ── DEBUG: Total entries assigned across all companies ────────────────────
    total_assigned_entries = sum(r.get("eB Matches", 0) for r in comparison_rows)
    log("\n── DEBUG: Entry Assignment Summary ────────────────────────────────────")
    log(f"  {'Company':<45} {'Entries':>7}")
    log(f"  {'-'*45} {'-'*7}")
    for r in sorted(comparison_rows, key=lambda x: x.get("eB Matches", 0), reverse=True):
        if r.get("eB Matches", 0) > 0:
            log(f"  {r['Vendor'][:45]:<45} {r['eB Matches']:>7}")
    log(f"  {'-'*45} {'-'*7}")
    log(f"  {'TOTAL ENTRIES ASSIGNED':<45} {total_assigned_entries:>7}")
    #log(f"  Total eBuilder rows loaded: {len(eb_raw)}")

    # ── DEBUG: list eBuilder '#' (ID) codes that were actually matched ──────
    # Built from claimed_ids (every row appended into some company's
    # _entries across ALL passes: direct match, fuzzy, PSU OPP, AJP merge,
    # _fix_zero_eb, etc.). NOTE: "eB Matches" / total_assigned_entries above
    # is a raw SUM across passes — a row that gets matched, then re-routed
    # by a post-pass fix, then folded into another vendor, gets counted
    # multiple times there. That inflated sum can equal len(eb_raw) by
    # coincidence even when some rows were never actually claimed and
    # others were claimed 2-3 times. The discrepancy check below instead
    # uses the DISTINCT set of eBuilder IDs that ended up in some entries
    # list, which is the only number that reflects reality.
    claimed_ids: set = set()
    for r in comparison_rows:
        for entry in r.get("_entries", []):
            eid = str(entry.get("ID", "")).strip()
            if eid:
                claimed_ids.add(eid)

    all_eb_ids = (
        eb_raw["ID"].astype(str).str.strip().tolist()
        if "ID" in eb_raw.columns else []
    )
    matched_ids   = [eid for eid in all_eb_ids if eid in claimed_ids]
    unmatched_ids = [eid for eid in all_eb_ids if eid not in claimed_ids]

    log("────────────────────────────────────────────────────────────────────────\n")
    # ── END DEBUG ─────────────────────────────────────────────────────────────

    show_comparison_window(comparison_rows, privv_df)
    log(f"Comparison complete — {len(comparison_rows)} vendor(s) analyzed.")


def show_comparison_window(rows, privv_df):
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

    # Map tree item iid → row dict so we can look up _entries on click
    iid_to_row: dict = {}

    def populate_tree(filter_text=""):
        tree.delete(*tree.get_children())
        iid_to_row.clear()
        ft = filter_text.strip().lower()
        for r in rows:
            if ft and ft not in r["Vendor"].lower():
                continue
            tag = "nodata"
            if "Match"   in r["Status"]: tag = "match"
            elif "Higher" in r["Status"]: tag = "higher"
            elif "Lower"  in r["Status"]: tag = "lower"
            iid = tree.insert("", "end", values=(
                r["Vendor"],
                f"${r['PRIVV Total']:,.2f}",
                f"${r['eBuilder Total']:,.2f}",
                f"${r['Delta']:+,.2f}",
                r["eB Matches"],
                r["Status"],
            ), tags=(tag,))
            iid_to_row[iid] = r

    def _populate_entries_tree(etree, entries, entry_cols, empty_label):
        if entries:
            for e in entries:
                etree.insert("", "end", values=tuple(
                    e.get(c, "") for c in entry_cols
                ))
        else:
            etree.insert("", "end", values=(empty_label,) + ("",) * (len(entry_cols) - 1))

    def _build_entries_tab(parent, entries, fallback_cols, col_widths, empty_label):
        # Build the column list from whatever fields are actually present
        # on the entries (full line item) rather than a fixed subset, so
        # every field captured for that row gets shown. Order is preserved
        # based on first appearance across all entries.
        if entries:
            entry_cols = []
            seen = set()
            for e in entries:
                for k in e.keys():
                    if k not in seen:
                        seen.add(k)
                        entry_cols.append(k)
        else:
            entry_cols = list(fallback_cols)

        ef = tk.Frame(parent)
        ef.pack(fill="both", expand=True, padx=8, pady=6)

        evsb = ttk.Scrollbar(ef, orient="vertical")
        ehsb = ttk.Scrollbar(ef, orient="horizontal")
        etree = ttk.Treeview(
            ef, columns=entry_cols, show="headings",
            yscrollcommand=evsb.set, xscrollcommand=ehsb.set, height=16
        )
        evsb.config(command=etree.yview)
        ehsb.config(command=etree.xview)

        money_cols = {
            "Commitment Amount", "Current Commitment",
            "Projected Commitment", "Actuals Approved", "Remaining Balance",
            "Amount",
        }
        for c in entry_cols:
            etree.heading(c, text=c)
            etree.column(c, width=col_widths.get(c, 130),
                         anchor="e" if c in money_cols else "w")

        _populate_entries_tree(etree, entries, entry_cols, empty_label)

        evsb.pack(side="right", fill="y")
        ehsb.pack(side="bottom", fill="x")
        etree.pack(fill="both", expand=True)
        return etree

    def show_entries_window(event=None):
        sel = tree.selection()
        if not sel:
            return
        row = iid_to_row.get(sel[0])
        if row is None:
            return
        eb_entries = row.get("_entries", [])

        # Fuzzy-match PRIVV rows for this same company on the fly (no
        # hardcoded vendor names — see find_fuzzy_privv_matches), rather
        # than threading a separate _privv_entries field through every
        # matching pass.
        privv_matches = find_fuzzy_privv_matches(row["Vendor"], privv_df)
        privv_entries = privv_matches.to_dict("records")
        # Use the sum of what's ACTUALLY shown in the PRIVV tab (fuzzy
        # matches), not row["PRIVV Total"] — that field comes from the
        # earlier exact-match aggregation pass and won't reflect rows that
        # only the fuzzy/Description matching here picked up.
        if "_amount_num" in privv_matches.columns:
            privv_total_fuzzy = privv_matches["_amount_num"].sum()
        else:
            privv_total_fuzzy = pd.to_numeric(
                privv_matches.get("Amount", pd.Series(dtype=str)).astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            ).sum()

        detail = tk.Toplevel(win)
        detail.title(f"Entries — {row['Vendor']}")
        detail.geometry("1300x500")

        privv_total_note = ""
        if abs(privv_total_fuzzy - row["PRIVV Total"]) > 0.005:
            privv_total_note = f" (was ${row['PRIVV Total']:,.2f} in summary)"

        hdr = tk.Frame(detail, bg="#1e1e2e", padx=10, pady=6)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text=f"  Vendor: {row['Vendor']}     "
                 f"eBuilder Total: ${row['eBuilder Total']:,.2f}     "
                 f"PRIVV Total: ${privv_total_fuzzy:,.2f}{privv_total_note}     "
                 f"eB Entries: {len(eb_entries)}     "
                 f"PRIVV Entries: {len(privv_entries)}",
            bg="#1e1e2e", fg="white", font=("Consolas", 10, "bold"), anchor="w"
        ).pack(fill="x")

        nb = ttk.Notebook(detail)
        nb.pack(fill="both", expand=True, padx=8, pady=6)

        eb_tab = tk.Frame(nb)
        privv_tab = tk.Frame(nb)
        nb.add(eb_tab, text=f"eBuilder Entries ({len(eb_entries)})")
        nb.add(privv_tab, text=f"PRIVV Entries ({len(privv_entries)})")

        eb_col_widths = {
            "ID": 70, "Description": 260, "Company": 200,
            "Date": 90, "Status": 100, "Commitment Type": 130,
            "Commitment Amount": 130, "Current Commitment": 130,
            "Projected Commitment": 130, "Actuals Approved": 130,
            "Remaining Balance": 130,
        }
        _build_entries_tab(
            eb_tab, eb_entries, ENTRY_COLS, eb_col_widths,
            "(No eBuilder entries found)",
        )

        privv_col_widths = {
            "Vendor": 220, "Description": 280, "Amount": 130,
            "Date": 90, "Status": 100, "Item": 90, "Type": 130,
        }
        _build_entries_tab(
            privv_tab, privv_entries, list(privv_df.columns), privv_col_widths,
            "(No PRIVV entries found)",
        )

        tk.Button(detail, text="Close", command=detail.destroy, padx=10).pack(pady=6)

    tree.bind("<Double-1>", show_entries_window)


    populate_tree()
    search_var.trace_add("write", lambda *_: populate_tree(search_var.get()))

    vsb.pack(side="right",  fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    hint = tk.Label(win, text="💡 Double-click any row to view its eBuilder entries",
                    fg="#555", font=("Segoe UI", 9, "italic"))
    hint.pack(pady=(0, 2))

    legend = tk.Frame(win, padx=8, pady=4)
    legend.pack(fill="x")
    for color, label in [
        ("#d4edda", "✅ Match"),
        ("#fff3cd", "🔼 eBuilder Higher"),
        ("#f8d7da", "🔽 eBuilder Lower"),
        ("#e2e3e5", "⚪ No Data"),
    ]:
        tk.Label(legend, text=f"  {label}  ", bg=color, relief="solid", bd=1, padx=6, pady=2).pack(side="left", padx=4)

    STATUS_FILL = {
        "✅ Match":           "C6EFCE",
        "🔼 eBuilder Higher": "FFEB9C",
        "🔽 eBuilder Lower":  "FFC7CE",
        "⚪ No Data":          "E2E3E5",
    }
    HEADER_FILL = PatternFill("solid", start_color="1E1E2E", end_color="1E1E2E")
    HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF")
    BASE_FONT   = Font(name="Calibri")
    THIN = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def _style_header(ws, headers):
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"

    def _autosize(ws, n_cols, min_w=10, max_w=55):
        for c in range(1, n_cols + 1):
            letter = get_column_letter(c)
            longest = max(
                (len(str(cell.value)) for cell in ws[letter] if cell.value is not None),
                default=min_w,
            )
            ws.column_dimensions[letter].width = max(min_w, min(longest + 2, max_w))

    def export_comparison():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save comparison as..."
        )
        if not path:
            return

        wb = openpyxl.Workbook()

        # ── Summary sheet ────────────────────────────────────────────────
        ws = wb.active
        ws.title = "Summary"
        headers = ["Vendor", "eBuilder Total", "PRIVV Total", "Delta",
                   "Status", "eB Matches"]
        _style_header(ws, headers)

        money_cols = {2, 3, 4}
        for r_idx, r in enumerate(rows, start=2):
            values = [
                r.get("Vendor", ""),
                r.get("eBuilder Total", 0) or 0,
                r.get("PRIVV Total", 0) or 0,
                r.get("Delta", 0) or 0,
                r.get("Status", ""),
                r.get("eB Matches", 0) or 0,
            ]
            fill = PatternFill("solid", start_color=STATUS_FILL.get(r.get("Status", ""), "FFFFFF"))
            for c_idx, val in enumerate(values, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = BASE_FONT
                cell.border = BORDER
                cell.fill = fill
                if c_idx in money_cols:
                    cell.number_format = '$#,##0.00;($#,##0.00);"-"'
                if c_idx == 5:
                    cell.alignment = Alignment(horizontal="center")
        _autosize(ws, len(headers))

        # ── eBuilder Entries sheet ───────────────────────────────────────
        eb_sheet = wb.create_sheet("eBuilder Entries")
        eb_headers = ["Vendor"] + list(ENTRY_COLS)
        _style_header(eb_sheet, eb_headers)
        r_idx = 2
        for r in rows:
            for e in r.get("_entries", []):
                row_vals = [r.get("Vendor", "")] + [e.get(c, "") for c in ENTRY_COLS]
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = eb_sheet.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = BASE_FONT
                    cell.border = BORDER
                r_idx += 1
        _autosize(eb_sheet, len(eb_headers))

        # ── PRIVV Entries sheet (same fuzzy match used in the drilldown) ──
        privv_sheet = wb.create_sheet("PRIVV Entries")
        privv_cols = list(privv_df.columns.drop("_amount_num", errors="ignore"))
        privv_headers = ["Matched To Vendor"] + privv_cols
        _style_header(privv_sheet, privv_headers)
        r_idx = 2
        for r in rows:
            matches = find_fuzzy_privv_matches(r["Vendor"], privv_df)
            for _, prow in matches.iterrows():
                row_vals = [r.get("Vendor", "")] + [prow.get(c, "") for c in privv_cols]
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = privv_sheet.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = BASE_FONT
                    cell.border = BORDER
                r_idx += 1
        _autosize(privv_sheet, len(privv_headers))

        wb.save(path)
        messagebox.showinfo("Exported", f"Saved to {path}")

    tk.Button(win, text="Export to Excel", command=export_comparison, bg="#0066cc", fg="white", padx=8).pack(pady=6)


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


#06/24/2026 Current code handles debugging and display unaccounted for rows. question is how ot minimize un accounted for rows


#06/24/2026 Current code handles debugging and display unaccounted for rows. question is how ot minimize un accounted for rows
