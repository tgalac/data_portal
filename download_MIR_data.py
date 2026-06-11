"""
Download Croatian MIR data from the ECB Data Portal API and create an
"enhanced" CSV with human-readable labels from the MIR DSD codelists.

Default output files:
    MIR_podaci_HR.csv
    MIR_podaci_HR_enhanced.csv

Requirements:
    pip install pandas requests
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

# The MIR data-structure definition contains the codelists used to decode
# dimension/attribute code values in the CSV returned by the data endpoint.
DSD_URL = "https://data-api.ecb.europa.eu/service/datastructure/ECB/ECB_MIR1?format=structure"

RAW_OUTPUT_CSV = "MIR_podaci_HR.csv"
ENHANCED_OUTPUT_CSV = "MIR_podaci_HR_enhanced.csv"

PREFERRED_LANGUAGE = "en"
REQUEST_TIMEOUT_SECONDS = 120


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


def find_enumeration_codelist_id(component: ET.Element) -> Optional[str]:
    """
    Find the codelist ID referenced by a Dimension or Attribute component.

    In SDMX-ML 2.1 this is typically:
        Dimension/LocalRepresentation/Enumeration/Ref id="CL_..."
    """
    for enum in component.iter():
        if local_name(enum.tag) != "Enumeration":
            continue

        for ref in enum.iter():
            if local_name(ref.tag) == "Ref":
                ref_id = ref.attrib.get("id")
                if ref_id:
                    return ref_id

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
    """
    Parse the mapping from MIR CSV component IDs to codelist IDs.

    Example output:
        {
            "FREQ": "CL_FREQ",
            "REF_AREA": "CL_AREA_EE",
            "BS_ITEM": "CL_BS_ITEM",
            ...
        }
    """
    component_to_codelist: Dict[str, str] = {}

    for component in root.iter():
        component_type = local_name(component.tag)
        if component_type not in {"Dimension", "DataAttribute", "Attribute"}:
            continue

        component_id = component.attrib.get("id")
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
    """Download and parse the MIR DSD XML structure."""
    headers = {
        "Accept": "application/vnd.sdmx.structure+xml;version=2.1, application/xml, text/xml"
    }

    print("Requesting MIR DSD / codelists from ECB API...")
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
    column named <COLUMN>_LABEL. By default, original code columns are kept.

    If keep_original_codes=False, the original code columns are replaced by
    labels instead. This is less safe for machine processing, so the default is
    True.
    """
    codelists = parse_codelists(dsd_root, preferred_language=preferred_language)
    component_to_codelist = parse_component_to_codelist_map(dsd_root)

    enhanced = df.copy()
    added_columns: list[str] = []
    replaced_columns: list[str] = []

    for column in list(df.columns):
        codelist_id = component_to_codelist.get(column)
        if not codelist_id:
            continue

        code_map = codelists.get(codelist_id)
        if not code_map:
            continue

        # Map using string representation, because some codelist values are
        # numeric-looking codes such as 0 or 2240, while pandas may infer int.
        source_values = df[column].map(lambda x: None if pd.isna(x) else str(x))
        label_values = source_values.map(code_map)

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

    return enhanced


def add_series_label_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add a compact SERIES_LABEL column from available dimension label columns.

    This is optional but useful in Excel filters/pivots because it gives one
    human-readable description per series row. Existing ECB TITLE/TITLE_COMPL
    columns are kept unchanged.
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

    label_columns = [f"{col}_LABEL" for col in dimension_order if f"{col}_LABEL" in df.columns]
    if not label_columns:
        return df

    result = df.copy()

    def combine_labels(row: pd.Series) -> str:
        parts = []
        for col in label_columns:
            value = row.get(col)
            if pd.notna(value) and str(value).strip():
                parts.append(str(value).strip())
        return " | ".join(parts)

    if "KEY" in result.columns:
        insert_at = result.columns.get_loc("KEY") + 1
    else:
        insert_at = 0

    if "SERIES_LABEL" not in result.columns:
        result.insert(insert_at, "SERIES_LABEL", result.apply(combine_labels, axis=1))

    return result


def enhance_mir_dataframe(df: pd.DataFrame, keep_original_codes: bool = True) -> pd.DataFrame:
    """Download the DSD and return a label-enhanced MIR DataFrame."""
    dsd_root = get_mir_dsd_root()
    enhanced = add_codelist_labels(
        df,
        dsd_root,
        preferred_language=PREFERRED_LANGUAGE,
        keep_original_codes=keep_original_codes,
    )
    enhanced = add_series_label_column(enhanced)
    return enhanced


# -----------------------------
# Main script
# -----------------------------


def main() -> int:
    mir_hr_df = get_ecb_mir_croatia()
    if mir_hr_df is None:
        return 1

    # Save raw ECB CSV output for reproducibility.
    mir_hr_df.to_csv(RAW_OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Saved raw data to: {RAW_OUTPUT_CSV}")

    try:
        enhanced_df = enhance_mir_dataframe(mir_hr_df, keep_original_codes=True)
    except requests.HTTPError as exc:
        print("Failed to fetch MIR DSD from ECB API.")
        print(exc)
        return 1
    except ET.ParseError as exc:
        print("Failed to parse MIR DSD XML.")
        print(exc)
        return 1

    # UTF-8 with BOM opens cleanly in many Excel installations.
    enhanced_df.to_csv(ENHANCED_OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Saved enhanced data to: {ENHANCED_OUTPUT_CSV}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
