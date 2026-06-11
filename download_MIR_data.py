"""
Download Croatian MIR data from the ECB Data Portal API and create a lighter
CSV with human-readable labels from the MIR DSD codelists.

Default output file:
    MIR_podaci_HR_enhanced.csv

Requirements:
    pip install pandas requests

Main points:
    - Default mode replaces ECB dimension/attribute codes with labels.
    - This avoids duplicate CODE + CODE_LABEL columns and reduces the number
      of fields shown in browser-based CSV/pivot viewers.
    - A switch near the top of the script can restore the older behaviour.
    - Fails loudly if no labels are added, instead of silently producing
      an unchanged CSV.
"""

from __future__ import annotations

import io
import sys
import xml.etree.ElementTree as ET
from typing import Dict, Iterable, Optional, Tuple

import pandas as pd
import requests


# -----------------------------
# User-editable configuration
# -----------------------------

DATA_URL = "https://data-api.ecb.europa.eu/service/data/MIR/"
SERIES_KEY = "M.HR.......EUR.N"

# Important: references=all is needed so that the DSD response includes the
# referenced codelists, not only the structural shell.
DSD_URL = (
    "https://data-api.ecb.europa.eu/service/datastructure/ECB/ECB_MIR1"
    "?references=all&format=structure"
)

RAW_OUTPUT_CSV = "MIR_podaci_HR.csv"
ENHANCED_OUTPUT_CSV = "MIR_podaci_HR_enhanced.csv"

# Output mode for the enhanced CSV:
#   "labels_only"      = replace code values with labels; no *_LABEL duplicates.
#   "codes_and_labels" = keep code columns and add adjacent *_LABEL columns.
#   "codes_only"       = raw ECB-style output only, saved as ENHANCED_OUTPUT_CSV.
OUTPUT_MODE = "labels_only"

# Set this to True only if you also want to write MIR_podaci_HR.csv locally.
# Your GitHub workflow does not need to commit this raw file.
SAVE_RAW_CSV = False

PREFERRED_LANGUAGE = "en"
REQUEST_TIMEOUT_SECONDS = 120

# Explicit MIR dimension-to-codelist mapping shown on the ECB MIR Data Portal
# structure page. This makes the script robust even if parsing the component
# mapping from the SDMX XML varies slightly.
MANUAL_COMPONENT_TO_CODELIST: Dict[str, str] = {
    "FREQ": "CL_FREQ",
    "REF_AREA": "CL_AREA_EE",
    "BS_REP_SECTOR": "CL_BS_REP_SECTOR",
    "BS_ITEM": "CL_BS_ITEM",
    "MATURITY_NOT_IRATE": "CL_MATURITY_ORIG",
    "DATA_TYPE_MIR": "CL_DATA_TYPE_MIR",
    "AMOUNT_CAT": "CL_AMOUNT_CAT",
    "BS_COUNT_SECTOR": "CL_BS_COUNT_SECTOR",
    "CURRENCY_TRANS": "CL_CURRENCY",
    "IR_BUS_COV": "CL_IR_BUS_COV",
    "OBS_STATUS": "CL_OBS_STATUS",
    "OBS_CONF": "CL_OBS_CONF",
}


# -----------------------------
# Small XML/SDMX helpers
# -----------------------------


def local_name(tag: str) -> str:
    """Return the XML local name, stripping any namespace."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def iter_by_local_name(root: ET.Element, wanted_name: str) -> Iterable[ET.Element]:
    """Yield descendants whose tag has the requested local name."""
    for element in root.iter():
        if local_name(element.tag) == wanted_name:
            yield element


def get_best_name(element: ET.Element, preferred_language: str = "en") -> Optional[str]:
    """
    Extract the best SDMX Name from an element.

    ECB SDMX structures usually hold labels in child Name elements, often with
    xml:lang="en". If the preferred language is not found, the first available
    Name is used.
    """
    names: list[Tuple[Optional[str], str]] = []
    for child in element:
        if local_name(child.tag) == "Name" and child.text and child.text.strip():
            lang = child.attrib.get("{http://www.w3.org/XML/1998/namespace}lang")
            names.append((lang, child.text.strip()))

    for lang, text in names:
        if lang == preferred_language:
            return text

    return names[0][1] if names else None


def first_ref_id(parent: ET.Element) -> Optional[str]:
    """Return the first nested SDMX Ref id, if any."""
    for element in parent.iter():
        if local_name(element.tag) == "Ref":
            ref_id = element.attrib.get("id")
            if ref_id:
                return ref_id
    return None


def get_component_id(component: ET.Element) -> Optional[str]:
    """
    Return a Dimension/Attribute concept id.

    Some SDMX XML variants put this directly in an id attribute; others identify
    it under ConceptIdentity/Ref. This supports both patterns.
    """
    direct_id = component.attrib.get("id")
    if direct_id:
        return direct_id

    for child in component:
        if local_name(child.tag) == "ConceptIdentity":
            return first_ref_id(child)

    return None


def find_enumeration_codelist_id(component: ET.Element) -> Optional[str]:
    """
    Find the codelist ID referenced by a Dimension or Attribute component.
    """
    for element in component.iter():
        if local_name(element.tag) == "Enumeration":
            return first_ref_id(element)
    return None


def parse_codelists(root: ET.Element, preferred_language: str = "en") -> Dict[str, Dict[str, str]]:
    """
    Parse all SDMX codelists in the DSD.

    Returns:
        {
            "CL_FREQ": {"M": "Monthly", ...},
            ...
        }
    """
    codelists: Dict[str, Dict[str, str]] = {}

    for codelist in iter_by_local_name(root, "Codelist"):
        codelist_id = codelist.attrib.get("id")
        if not codelist_id:
            continue

        code_map: Dict[str, str] = {}
        for child in codelist:
            if local_name(child.tag) != "Code":
                continue

            code_id = child.attrib.get("id")
            if not code_id:
                continue

            label = get_best_name(child, preferred_language=preferred_language) or code_id
            code_map[code_id] = label

        if code_map:
            codelists[codelist_id] = code_map

    return codelists


def parse_component_to_codelist_map(root: ET.Element) -> Dict[str, str]:
    """Parse mapping from SDMX component IDs to codelist IDs."""
    component_to_codelist: Dict[str, str] = {}

    for component in root.iter():
        component_type = local_name(component.tag)
        if component_type not in {"Dimension", "TimeDimension", "DataAttribute", "Attribute"}:
            continue

        component_id = get_component_id(component)
        if not component_id:
            continue

        codelist_id = find_enumeration_codelist_id(component)
        if codelist_id:
            component_to_codelist[component_id] = codelist_id

    return component_to_codelist


# -----------------------------
# ECB retrieval functions
# -----------------------------


def get_ecb_mir_croatia() -> Optional[pd.DataFrame]:
    """Download the Croatian MIR subset from the ECB API as a DataFrame."""
    headers = {"Accept": "text/csv"}
    request_url = DATA_URL + SERIES_KEY

    print(f"Requesting data from ECB API for key: MIR.{SERIES_KEY}...")
    response = requests.get(request_url, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)

    if response.status_code != 200:
        print(f"Failed to fetch data. HTTP Status Code: {response.status_code}")
        print(response.text)
        return None

    print("Success. Loading data into a Pandas DataFrame...")
    df = pd.read_csv(io.StringIO(response.text), low_memory=False)

    if "KEY" in df.columns and "TIME_PERIOD" in df.columns:
        df = df.sort_values(by=["KEY", "TIME_PERIOD"]).reset_index(drop=True)

    return df


def get_mir_dsd_root() -> ET.Element:
    """Download and parse the MIR DSD XML structure, including codelists."""
    headers = {
        "Accept": "application/vnd.sdmx.structure+xml;version=2.1, application/xml, text/xml"
    }

    print("Requesting MIR DSD / codelists from ECB API...")
    print(f"DSD URL: {DSD_URL}")
    response = requests.get(DSD_URL, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
    response.raise_for_status()

    return ET.fromstring(response.content)


# -----------------------------
# Label enrichment
# -----------------------------


def add_codelist_labels(
    df: pd.DataFrame,
    dsd_root: ET.Element,
    preferred_language: str = "en",
    keep_original_codes: bool = True,
) -> pd.DataFrame:
    """
    Add human-readable label columns to a MIR CSV DataFrame.

    For each CSV column that has a codelist in the MIR DSD, this adds a new
    column named <COLUMN>_LABEL. Original code columns are kept by default.
    """
    codelists = parse_codelists(dsd_root, preferred_language=preferred_language)
    parsed_component_map = parse_component_to_codelist_map(dsd_root)

    # Manual mapping is used as a fallback and also normalises the mapping to
    # the ECB MIR CSV column names.
    component_to_codelist = dict(parsed_component_map)
    component_to_codelist.update(MANUAL_COMPONENT_TO_CODELIST)

    print(f"Parsed {len(codelists)} codelists from DSD.")
    print(f"Parsed {len(parsed_component_map)} component mappings from DSD XML.")

    if not codelists:
        raise RuntimeError(
            "No codelists were found in the DSD response. The usual cause is "
            "that the DSD URL was fetched without references=all."
        )

    enhanced = df.copy()
    added_columns: list[str] = []
    replaced_columns: list[str] = []
    mapped_value_count = 0

    for column in list(df.columns):
        codelist_id = component_to_codelist.get(column)
        if not codelist_id:
            continue

        code_map = codelists.get(codelist_id)
        if not code_map:
            print(f"Warning: codelist {codelist_id} for column {column} was not found in DSD.")
            continue

        # Map using string representation, because some codelist values are
        # numeric-looking codes such as 0 or 2240, while pandas may infer int.
        source_values = df[column].map(lambda x: None if pd.isna(x) else str(x))
        label_values = source_values.map(code_map)
        mapped_value_count += int(label_values.notna().sum())

        # Keep unmapped values visible rather than creating blank labels.
        label_values = label_values.fillna(source_values)

        if keep_original_codes:
            label_column = f"{column}_LABEL"
            insert_after = enhanced.columns.get_loc(column) + 1
            enhanced.insert(insert_after, label_column, label_values)
            added_columns.append(label_column)
        else:
            enhanced[column] = label_values
            replaced_columns.append(column)

    if keep_original_codes:
        print(f"Added {len(added_columns)} label columns.")
        if added_columns:
            print("Label columns:", ", ".join(added_columns))
    else:
        print(f"Replaced codes with labels in {len(replaced_columns)} columns.")
        if replaced_columns:
            print("Replaced columns:", ", ".join(replaced_columns))

    if mapped_value_count == 0:
        raise RuntimeError(
            "No code values were mapped to labels. The enhanced CSV would be "
            "identical to the raw CSV, so the script is stopping. Check the "
            "GitHub Actions log for missing codelists or changed ECB DSD IDs."
        )

    return enhanced


def add_series_label_column(df: pd.DataFrame) -> pd.DataFrame:
    """Add a compact SERIES_LABEL column from available MIR dimension columns.

    Works in both output modes:
      - codes_and_labels: uses *_LABEL columns;
      - labels_only: uses the original dimension columns after their values
        have already been replaced by labels.
    """
    dimension_order = [
        "FREQ",
        "REF_AREA",
        "BS_REP_SECTOR",
        "BS_ITEM",
        "MATURITY_NOT_IRATE",
        "DATA_TYPE_MIR",
        "AMOUNT_CAT",
        "BS_COUNT_SECTOR",
        "CURRENCY_TRANS",
        "IR_BUS_COV",
    ]

    result = df.copy()

    # Prefer *_LABEL columns when they exist. Otherwise use the dimension
    # columns themselves, which is correct in labels_only mode.
    source_columns = []
    for col in dimension_order:
        label_col = f"{col}_LABEL"
        if label_col in result.columns:
            source_columns.append(label_col)
        elif col in result.columns:
            source_columns.append(col)

    if not source_columns:
        return result

    def combine_labels(row: pd.Series) -> str:
        parts = []
        for col in source_columns:
            value = row.get(col)
            if pd.notna(value) and str(value).strip():
                parts.append(str(value).strip())
        return " | ".join(parts)

    insert_at = result.columns.get_loc("KEY") + 1 if "KEY" in result.columns else 0

    if "SERIES_LABEL" not in result.columns:
        result.insert(insert_at, "SERIES_LABEL", result.apply(combine_labels, axis=1))

    return result


def simplify_for_labels_only_viewer(df: pd.DataFrame) -> pd.DataFrame:
    """Return a compact labels-only CSV suitable for the browser viewer.

    The coded dimension/attribute columns keep their original column names but
    contain human-readable labels. Technical columns that are usually not
    useful as browser filters are dropped when present.
    """
    technical_columns_to_drop = [
        "STRUCTURE",
        "STRUCTURE_ID",
        "ACTION",
    ]

    result = df.drop(
        columns=[col for col in technical_columns_to_drop if col in df.columns],
        errors="ignore",
    ).copy()

    # Keep the most useful columns near the front for Excel and web viewers.
    preferred_front_order = [
        "KEY",
        "SERIES_LABEL",
        "TIME_PERIOD",
        "OBS_VALUE",
        "FREQ",
        "REF_AREA",
        "BS_REP_SECTOR",
        "BS_ITEM",
        "MATURITY_NOT_IRATE",
        "DATA_TYPE_MIR",
        "AMOUNT_CAT",
        "BS_COUNT_SECTOR",
        "CURRENCY_TRANS",
        "IR_BUS_COV",
    ]

    front = [col for col in preferred_front_order if col in result.columns]
    rest = [col for col in result.columns if col not in front]
    return result[front + rest]

def enhance_mir_dataframe(df: pd.DataFrame, output_mode: str = OUTPUT_MODE) -> pd.DataFrame:
    """Download the DSD and return a MIR DataFrame in the selected output mode."""
    if output_mode not in {"labels_only", "codes_and_labels", "codes_only"}:
        raise ValueError(
            "OUTPUT_MODE must be one of: labels_only, codes_and_labels, codes_only"
        )

    if output_mode == "codes_only":
        return df.copy()

    dsd_root = get_mir_dsd_root()
    keep_original_codes = output_mode == "codes_and_labels"
    enhanced = add_codelist_labels(
        df,
        dsd_root,
        preferred_language=PREFERRED_LANGUAGE,
        keep_original_codes=keep_original_codes,
    )
    enhanced = add_series_label_column(enhanced)

    if output_mode == "labels_only":
        enhanced = simplify_for_labels_only_viewer(enhanced)

    return enhanced


# -----------------------------
# Main script
# -----------------------------


def main() -> int:
    mir_hr_df = get_ecb_mir_croatia()
    if mir_hr_df is None:
        return 1

    if SAVE_RAW_CSV:
        # Optional raw ECB CSV output for reproducibility.
        mir_hr_df.to_csv(RAW_OUTPUT_CSV, index=False, encoding="utf-8-sig")
        print(f"Saved raw data to: {RAW_OUTPUT_CSV}")
    else:
        print("Skipping raw CSV output because SAVE_RAW_CSV = False.")

    try:
        enhanced_df = enhance_mir_dataframe(mir_hr_df, output_mode=OUTPUT_MODE)
    except requests.HTTPError as exc:
        print("Failed to fetch MIR DSD from ECB API.")
        print(exc)
        return 1
    except ET.ParseError as exc:
        print("Failed to parse MIR DSD XML.")
        print(exc)
        return 1
    except RuntimeError as exc:
        print("Failed to create a labelled/enhanced CSV.")
        print(exc)
        return 1

    # UTF-8 with BOM opens cleanly in many Excel installations.
    enhanced_df.to_csv(ENHANCED_OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Saved enhanced data to: {ENHANCED_OUTPUT_CSV}")
    print(f"Output mode: {OUTPUT_MODE}")
    print(f"Output columns: {len(enhanced_df.columns)}")
    print("Columns:", ", ".join(enhanced_df.columns))

    return 0


if __name__ == "__main__":
    sys.exit(main())
