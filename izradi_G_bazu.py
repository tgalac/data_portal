#!/usr/bin/env python3
"""Pretvara prezentacijske tablice G1, G2, G3, G5 i G6 u lean CSV bazu.

Zrno izlazne baze jest datum pomnozen ekonomskim dimenzijama. Izvorna tablica
nije dimenzija, nego pomocni atribut. Pet mjera jesu nominalna kamatna stopa na
nove poslove, nominalna kamatna stopa na stanja, efektivna kamatna stopa, iznos
novih poslova i iznos stanja.

Kod preklapanja detaljne tablice G1-G3 imaju prednost pred G6. Za iznos novih
poslova u G2 koristi se iznos iz bloka nominalnih stopa, a iznos uz efektivnu
stopu se ne izvozi. CSV je prilagodjen hrvatskom Excelu: UTF-8 BOM, tocka-zarez
kao delimiter i decimalni zarez. Prazne vrijednosti i izvorni znak "–" se ne
izvoze kao posebna opazanja. Datumski stupci otkrivaju se iz knjige pri svakom
pokretanju, pa novi mjeseci ne zahtijevaju promjenu koda. Svi podatkovni blokovi
moraju imati jednak mjesečni raspon; nepotpuno ažurirana knjiga zaustavlja se
jasnom pogreškom.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


UKUPNO = "UKUPNO"
NP = "NIJE_PRIMJENJIVO"
NOVI = "NOVI_POSLOVI_TIJEKOM_MJESECA"
STANJE = "STANJE_NA_KRAJU_MJESECA"


@dataclass(frozen=True)
class SeriesMeta:
    pozicija: str
    sektor: str = UKUPNO
    podskup: str = UKUPNO
    instrument: str = UKUPNO
    namjena: str = NP
    dospijece: str = NP
    otkazni_rok: str = NP
    fiksiranje: str = NP
    velicina: str = NP
    osnova: str = NOVI
    fusnota_a: bool = False


@dataclass(frozen=True)
class BlockSpec:
    sheet: str
    vrsta_stope: str
    koncept_tablice: str
    date_row_rate: int
    date_row_amount: int
    status_row_rate: int
    status_row_amount: int
    amount_offset: int
    rows: dict[int, SeriesMeta]


@dataclass(frozen=True)
class Candidate:
    value: str
    source: str
    priority: int


def dep(
    sektor: str = UKUPNO,
    instrument: str = UKUPNO,
    *,
    dospijece: str = NP,
    otkazni_rok: str = NP,
    osnova: str = NOVI,
    fusnota_a: bool = False,
) -> SeriesMeta:
    return SeriesMeta(
        pozicija="DEPOZITI",
        sektor=sektor,
        instrument=instrument,
        namjena=NP,
        dospijece=dospijece,
        otkazni_rok=otkazni_rok,
        fiksiranje=NP,
        velicina=NP,
        osnova=osnova,
        fusnota_a=fusnota_a,
    )


def loan(
    sektor: str,
    *,
    podskup: str = UKUPNO,
    instrument: str = "KREDITI",
    namjena: str = UKUPNO,
    dospijece: str = UKUPNO,
    fiksiranje: str = UKUPNO,
    velicina: str = UKUPNO,
    osnova: str = NOVI,
    fusnota_a: bool = False,
) -> SeriesMeta:
    return SeriesMeta(
        pozicija="KREDITI",
        sektor=sektor,
        podskup=podskup,
        instrument=instrument,
        namjena=namjena,
        dospijece=dospijece,
        otkazni_rok=NP,
        fiksiranje=fiksiranje,
        velicina=velicina,
        osnova=osnova,
        fusnota_a=fusnota_a,
    )


def g1_rows() -> dict[int, SeriesMeta]:
    h, n = "KUCANSTVA", "NEFINANCIJSKA_DRUSTVA"
    r: dict[int, SeriesMeta] = {
        10: dep(h, "PREKONOCNI_DEPOZITI", osnova=STANJE),
        11: dep(h, "TRANSAKCIJSKI_RACUNI", osnova=STANJE),
        12: dep(h, "STEDNI_DEPOZITI", osnova=STANJE),
        13: dep(h, "OROCENI_DEPOZITI", dospijece=UKUPNO, fusnota_a=True),
        14: dep(h, "OROCENI_DEPOZITI", dospijece="DO_3_MJESECA"),
        15: dep(h, "OROCENI_DEPOZITI", dospijece="OD_3_DO_6_MJESECI"),
        16: dep(h, "OROCENI_DEPOZITI", dospijece="OD_6_MJESECI_DO_1_GODINE"),
        17: dep(h, "OROCENI_DEPOZITI", dospijece="OD_1_DO_2_GODINE"),
        18: dep(h, "OROCENI_DEPOZITI", dospijece="VISE_OD_2_GODINE"),
        19: dep(h, "DEPOZITI_U_OTKAZNOM_ROKU", otkazni_rok=UKUPNO),
        20: dep(h, "DEPOZITI_U_OTKAZNOM_ROKU", otkazni_rok="DO_3_MJESECA"),
        21: dep(h, "DEPOZITI_U_OTKAZNOM_ROKU", otkazni_rok="VISE_OD_3_MJESECA"),
        23: dep(n, "PREKONOCNI_DEPOZITI", osnova=STANJE),
        24: dep(n, "TRANSAKCIJSKI_RACUNI", osnova=STANJE),
        25: dep(n, "STEDNI_DEPOZITI", osnova=STANJE),
        26: dep(n, "OROCENI_DEPOZITI", dospijece=UKUPNO, fusnota_a=True),
        27: dep(n, "OROCENI_DEPOZITI", dospijece="DO_3_MJESECA"),
        28: dep(n, "OROCENI_DEPOZITI", dospijece="OD_3_DO_6_MJESECI"),
        29: dep(n, "OROCENI_DEPOZITI", dospijece="OD_6_MJESECI_DO_1_GODINE"),
        30: dep(n, "OROCENI_DEPOZITI", dospijece="OD_1_DO_2_GODINE"),
        31: dep(n, "OROCENI_DEPOZITI", dospijece="VISE_OD_2_GODINE"),
        32: dep(UKUPNO, "REPO_POSLOVI", dospijece=UKUPNO),
    }
    return r


def g2_nominal_rows() -> dict[int, SeriesMeta]:
    h = "KUCANSTVA"
    other = "OSTALE_NAMJENE"
    cash = "GOTOVINSKI_NENAMJENSKI"
    r: dict[int, SeriesMeta] = {
        9: loan(h, instrument="REVOLVING_PREKORACENJA_I_KARTICE", dospijece=NP, fiksiranje=NP, osnova=STANJE),
        10: loan(h, instrument="REVOLVING_KREDITI", dospijece=NP, fiksiranje=NP, osnova=STANJE),
        11: loan(h, instrument="PREKORACENJA_PO_TRANSAKCIJSKOM_RACUNU", dospijece=NP, fiksiranje=NP, osnova=STANJE),
        12: loan(h, instrument="KREDITI_PO_KREDITNIM_KARTICAMA", dospijece=NP, fiksiranje=NP, osnova=STANJE),
        13: loan(h, podskup="OBRTNICI", instrument="REVOLVING_PREKORACENJA_I_KARTICE", dospijece=NP, fiksiranje=NP, osnova=STANJE),
        14: loan(h, namjena="POTROSACKI", fusnota_a=True),
        15: loan(h, namjena="POTROSACKI", fiksiranje="PROMJENJIVA_ILI_DO_1_GODINE"),
        16: loan(h, namjena="POTROSACKI", fiksiranje="OD_1_DO_5_GODINA"),
        17: loan(h, namjena="POTROSACKI", fiksiranje="VISE_OD_5_GODINA"),
        18: loan(h, namjena="POTROSACKI", fiksiranje="FIKSNA_CIJELI_VIJEK"),
        19: loan(h, namjena="STAMBENI", fusnota_a=True),
        20: loan(h, namjena="STAMBENI", fiksiranje="PROMJENJIVA_ILI_DO_1_GODINE"),
        21: loan(h, namjena="STAMBENI", fiksiranje="OD_1_DO_5_GODINA"),
        22: loan(h, namjena="STAMBENI", fiksiranje="OD_5_DO_10_GODINA"),
        23: loan(h, namjena="STAMBENI", fiksiranje="VISE_OD_10_GODINA"),
        24: loan(h, namjena="STAMBENI", fiksiranje="FIKSNA_CIJELI_VIJEK"),
        25: loan(h, namjena="STAMBENI", dospijece="KRATKOROCNO", fiksiranje="FIKSNA_CIJELI_VIJEK"),
        26: loan(h, namjena="STAMBENI", dospijece="DUGOROCNO", fiksiranje="FIKSNA_CIJELI_VIJEK"),
        27: loan(h, namjena=other, fusnota_a=True),
        28: loan(h, namjena=other, fiksiranje="PROMJENJIVA_ILI_DO_1_GODINE"),
        29: loan(h, namjena=other, fiksiranje="OD_1_DO_5_GODINA"),
        30: loan(h, namjena=other, fiksiranje="VISE_OD_5_GODINA"),
        31: loan(h, namjena=cash),
        32: loan(h, namjena=cash, fiksiranje="PROMJENJIVA_ILI_DO_1_GODINE"),
        33: loan(h, namjena=cash, fiksiranje="OD_1_DO_5_GODINA"),
        34: loan(h, namjena=cash, fiksiranje="VISE_OD_5_GODINA"),
        35: loan(h, namjena=cash, fiksiranje="FIKSNA_CIJELI_VIJEK"),
        36: loan(h, podskup="OBRTNICI", namjena=other),
    }
    return r


def g2_effective_rows() -> dict[int, SeriesMeta]:
    h = "KUCANSTVA"
    return {
        75: loan(h, namjena="POTROSACKI", fusnota_a=True),
        76: loan(h, namjena="STAMBENI", fusnota_a=True),
        77: loan(h, namjena="OSTALE_NAMJENE"),
        78: loan(h, namjena="GOTOVINSKI_NENAMJENSKI"),
    }


def g3_rows() -> dict[int, SeriesMeta]:
    n = "NEFINANCIJSKA_DRUSTVA"
    r: dict[int, SeriesMeta] = {
        9: loan(n, instrument="REVOLVING_PREKORACENJA_I_KARTICE", dospijece=NP, fiksiranje=NP, velicina=NP, osnova=STANJE),
        10: loan(n, instrument="REVOLVING_I_PREKORACENJA", dospijece=NP, fiksiranje=NP, velicina=NP, osnova=STANJE),
        11: loan(n, instrument="KREDITI_PO_KREDITNIM_KARTICAMA", dospijece=NP, fiksiranje=NP, velicina=NP, osnova=STANJE),
    }
    fixation_rows = {
        1: "PROMJENJIVA_ILI_DO_3_MJESECA",
        2: "OD_3_MJESECA_DO_1_GODINE",
        3: "OD_1_DO_3_GODINE",
        4: "OD_3_DO_5_GODINA",
        5: "OD_5_DO_10_GODINA",
        6: "VISE_OD_10_GODINA",
    }
    for base, size in ((12, "DO_0_25_MIL_EUR"), (19, "OD_0_25_DO_1_MIL_EUR"), (26, "VISE_OD_1_MIL_EUR")):
        r[base] = loan(n, velicina=size, fusnota_a=True)
        for offset, fixation in fixation_rows.items():
            r[base + offset] = loan(n, velicina=size, fiksiranje=fixation)
    return r


def g5_rows() -> dict[int, SeriesMeta]:
    h, n = "KUCANSTVA", "NEFINANCIJSKA_DRUSTVA"
    r: dict[int, SeriesMeta] = {
        9: dep(UKUPNO, UKUPNO, dospijece=UKUPNO, osnova=STANJE),
        10: dep(UKUPNO, "OROCENI_DEPOZITI", dospijece=UKUPNO, osnova=STANJE, fusnota_a=True),
        11: dep(h, "OROCENI_DEPOZITI", dospijece=UKUPNO, osnova=STANJE),
        12: dep(h, "OROCENI_DEPOZITI", dospijece="KRATKOROCNO", osnova=STANJE),
        13: dep(h, "OROCENI_DEPOZITI", dospijece="DO_3_MJESECA", osnova=STANJE),
        14: dep(h, "OROCENI_DEPOZITI", dospijece="OD_3_DO_6_MJESECI", osnova=STANJE),
        15: dep(h, "OROCENI_DEPOZITI", dospijece="OD_6_MJESECI_DO_1_GODINE", osnova=STANJE),
        16: dep(h, "OROCENI_DEPOZITI", dospijece="DUGOROCNO", osnova=STANJE),
        17: dep(h, "OROCENI_DEPOZITI", dospijece="OD_1_DO_2_GODINE", osnova=STANJE),
        18: dep(h, "OROCENI_DEPOZITI", dospijece="VISE_OD_2_GODINE", osnova=STANJE),
        19: dep(n, "OROCENI_DEPOZITI", dospijece=UKUPNO, osnova=STANJE),
        20: dep(n, "OROCENI_DEPOZITI", dospijece="KRATKOROCNO", osnova=STANJE),
        21: dep(n, "OROCENI_DEPOZITI", dospijece="DO_3_MJESECA", osnova=STANJE),
        22: dep(n, "OROCENI_DEPOZITI", dospijece="OD_3_DO_6_MJESECI", osnova=STANJE),
        23: dep(n, "OROCENI_DEPOZITI", dospijece="OD_6_MJESECI_DO_1_GODINE", osnova=STANJE),
        24: dep(n, "OROCENI_DEPOZITI", dospijece="DUGOROCNO", osnova=STANJE),
        25: dep(n, "OROCENI_DEPOZITI", dospijece="OD_1_DO_2_GODINE", osnova=STANJE),
        26: dep(n, "OROCENI_DEPOZITI", dospijece="VISE_OD_2_GODINE", osnova=STANJE),
        27: dep(UKUPNO, "REPO_POSLOVI", dospijece=UKUPNO, osnova=STANJE),
        28: loan(UKUPNO, osnova=STANJE),
        29: loan(h, osnova=STANJE),
        30: loan(h, namjena="STAMBENI", osnova=STANJE, fusnota_a=True),
        31: loan(h, namjena="STAMBENI", dospijece="KRATKOROCNO", osnova=STANJE),
        32: loan(h, namjena="STAMBENI", dospijece="DUGOROCNO", osnova=STANJE),
        33: loan(h, namjena="STAMBENI", dospijece="OD_1_DO_5_GODINA", osnova=STANJE),
        34: loan(h, namjena="STAMBENI", dospijece="VISE_OD_5_GODINA", osnova=STANJE),
        35: loan(h, namjena="POTROSACKI_I_OSTALI", osnova=STANJE, fusnota_a=True),
        36: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="KRATKOROCNO", osnova=STANJE, fusnota_a=True),
        37: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="DUGOROCNO", osnova=STANJE, fusnota_a=True),
        38: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="OD_1_DO_5_GODINA", osnova=STANJE),
        39: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="VISE_OD_5_GODINA", osnova=STANJE),
        40: loan(h, podskup="OBRTNICI", namjena="POTROSACKI_I_OSTALI", osnova=STANJE),
        41: loan(n, osnova=STANJE, fusnota_a=True),
        42: loan(n, dospijece="KRATKOROCNO", osnova=STANJE),
        43: loan(n, dospijece="DUGOROCNO", osnova=STANJE),
        44: loan(n, dospijece="OD_1_DO_5_GODINA", osnova=STANJE),
        45: loan(n, dospijece="VISE_OD_5_GODINA", osnova=STANJE),
    }
    return r


def g6_rows() -> dict[int, SeriesMeta]:
    h, n = "KUCANSTVA", "NEFINANCIJSKA_DRUSTVA"
    r: dict[int, SeriesMeta] = {
        9: dep(UKUPNO, UKUPNO, dospijece=UKUPNO),
        10: dep(UKUPNO, "OROCENI_DEPOZITI", dospijece=UKUPNO, fusnota_a=True),
        11: dep(h, "OROCENI_DEPOZITI", dospijece=UKUPNO),
        12: dep(h, "OROCENI_DEPOZITI", dospijece="KRATKOROCNO"),
        13: dep(h, "OROCENI_DEPOZITI", dospijece="DUGOROCNO"),
        14: dep(n, "OROCENI_DEPOZITI", dospijece=UKUPNO),
        15: dep(n, "OROCENI_DEPOZITI", dospijece="KRATKOROCNO"),
        16: dep(n, "OROCENI_DEPOZITI", dospijece="DUGOROCNO"),
        17: dep(UKUPNO, "REPO_POSLOVI", dospijece=UKUPNO),
        18: loan(UKUPNO),
        19: loan(h),
        20: loan(h, namjena="STAMBENI", fusnota_a=True),
        21: loan(h, namjena="STAMBENI", dospijece="KRATKOROCNO"),
        22: loan(h, namjena="STAMBENI", dospijece="DUGOROCNO"),
        23: loan(h, namjena="STAMBENI", dospijece="OD_1_DO_5_GODINA"),
        24: loan(h, namjena="STAMBENI", dospijece="VISE_OD_5_GODINA"),
        25: loan(h, namjena="POTROSACKI_I_OSTALI", fusnota_a=True),
        26: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="KRATKOROCNO", fusnota_a=True),
        27: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="DUGOROCNO", fusnota_a=True),
        28: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="OD_1_DO_5_GODINA"),
        29: loan(h, namjena="POTROSACKI_I_OSTALI", dospijece="VISE_OD_5_GODINA"),
        30: loan(h, podskup="OBRTNICI", namjena="POTROSACKI_I_OSTALI"),
        31: loan(n, fusnota_a=True),
        32: loan(n, dospijece="KRATKOROCNO"),
        33: loan(n, dospijece="DUGOROCNO"),
        34: loan(n, dospijece="OD_1_DO_5_GODINA"),
        35: loan(n, dospijece="VISE_OD_5_GODINA"),
    }
    return r


BLOCKS = (
    BlockSpec("G1", "NOMINALNA", "NOVI_POSLOVI", 8, 37, 7, 36, 29, g1_rows()),
    BlockSpec("G2", "NOMINALNA", "NOVI_POSLOVI", 8, 41, 7, 40, 33, g2_nominal_rows()),
    BlockSpec("G2", "EFEKTIVNA", "NOVI_POSLOVI", 74, 83, 73, 82, 9, g2_effective_rows()),
    BlockSpec("G3", "NOMINALNA", "NOVI_POSLOVI", 8, 37, 7, 36, 29, g3_rows()),
    BlockSpec("G5", "NOMINALNA", "STANJA", 8, 50, 7, 49, 42, g5_rows()),
    BlockSpec("G6", "NOMINALNA", "NOVI_POSLOVI", 8, 40, 7, 39, 32, g6_rows()),
)


DIMENSION_FIELDS = (
    "datum",
    "pozicija",
    "sektor_protustranke",
    "podskup_protustranke",
    "instrument",
    "namjena_kredita",
    "izvorno_dospijece",
    "otkazni_rok",
    "pocetno_fiksiranje_kamatne_stope",
    "velicina_kredita",
)

MEASURE_FIELDS = (
    "nominalna_kamatna_stopa_novi_posao",
    "nominalna_kamatna_stopa_stanje",
    "efektivna_kamatna_stopa",
    "iznos_novi_posao",
    "iznos_stanje",
)

FIELDS = DIMENSION_FIELDS + MEASURE_FIELDS + ("izvorne_tablice",)


def normalize_space(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").replace("\t", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def cell_kind(value: Any) -> str:
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return "NUMERIC"
    if value == "–":
        return "DASH"
    if value in (None, ""):
        return "BLANK"
    raise ValueError(f"Neocekivana vrijednost u podatkovnoj celiji: {value!r}")


def csv_number(value: Any) -> str:
    if cell_kind(value) != "NUMERIC":
        return ""
    # Excel racuna s najvise 15 znacajnih znamenaka. Time se cuva puna
    # smislena preciznost, ali se ne prenose binarni artefakti poput
    # 1.6201416000000002 umjesto 1.6201416.
    text = format(value, ".15g") if isinstance(value, float) else format(Decimal(str(value)), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", ",")




def date_columns(ws: Any, row: int) -> list[int]:
    cols = [c for c in range(3, ws.max_column + 1) if isinstance(ws.cell(row, c).value, datetime)]
    if not cols:
        raise ValueError(f"{ws.title}: nisu pronadjeni datumi u retku {row}.")
    expected = list(range(cols[0], cols[-1] + 1))
    if cols != expected:
        raise ValueError(f"{ws.title}: datumski stupci u retku {row} nisu neprekinuti.")
    return cols


def validate_monthly(dates: Iterable[datetime], context: str) -> None:
    values = list(dates)
    for previous, current in zip(values, values[1:]):
        distance = (current.year - previous.year) * 12 + current.month - previous.month
        if distance != 1:
            raise ValueError(f"{context}: datumi nisu mjesecno neprekinuti: {previous} -> {current}")


def candidate_priority(sheet: str) -> int:
    """Detaljne tablice G1-G3 i tablica stanja G5 imaju prednost pred G6."""
    return 1 if sheet == "G6" else 0


def choose_candidate(candidates: list[Candidate], measure: str, key: tuple[str, ...]) -> Candidate:
    best_priority = min(candidate.priority for candidate in candidates)
    best = [candidate for candidate in candidates if candidate.priority == best_priority]
    values = {candidate.value for candidate in best}
    if len(values) != 1:
        details = [(candidate.value, candidate.source) for candidate in best]
        raise ValueError(f"Kolizija jednakopravnih izvora za {measure}, kljuc {key}: {details}")
    return sorted(best, key=lambda candidate: candidate.source)[0]


def build_rows_from_workbook(workbook: Any, input_label: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    required = {"Metodologija", "G1", "G2", "G3", "G5", "G6"}
    missing = required.difference(workbook.sheetnames)
    if missing:
        raise ValueError(f"Nedostaju ocekivani listovi: {sorted(missing)}")

    groups: dict[tuple[str, ...], dict[str, Any]] = {}
    source_series = sum(len(block.rows) for block in BLOCKS)
    possible = numeric = dash = blank = ignored_effective_amounts = 0
    common_dates: list[datetime] | None = None

    for block in BLOCKS:
        ws = workbook[block.sheet]
        rate_cols = date_columns(ws, block.date_row_rate)
        amount_cols = date_columns(ws, block.date_row_amount)
        if rate_cols != amount_cols:
            raise ValueError(f"{block.sheet} {block.vrsta_stope}: datumski stupci stopa i iznosa nisu jednaki.")

        rate_dates = [ws.cell(block.date_row_rate, c).value for c in rate_cols]
        amount_dates = [ws.cell(block.date_row_amount, c).value for c in amount_cols]
        if rate_dates != amount_dates:
            raise ValueError(f"{block.sheet} {block.vrsta_stope}: datumi stopa i iznosa nisu jednaki.")
        validate_monthly(rate_dates, f"{block.sheet} {block.vrsta_stope}")
        if common_dates is None:
            common_dates = rate_dates
        elif rate_dates != common_dates:
            raise ValueError(
                f"{block.sheet} {block.vrsta_stope}: datumski raspon nije jednak ostalim podatkovnim blokovima."
            )

        for rate_row, meta in sorted(block.rows.items()):
            amount_row = rate_row + block.amount_offset
            raw_rate_label = str(ws.cell(rate_row, 2).value or "")
            raw_amount_label = str(ws.cell(amount_row, 2).value or "")
            if normalize_space(raw_rate_label) != normalize_space(raw_amount_label):
                raise ValueError(
                    f"{block.sheet}: oznake retka stope {rate_row} i iznosa {amount_row} nisu jednake."
                )

            for col, date in zip(rate_cols, rate_dates):
                possible += 1
                rate_value = ws.cell(rate_row, col).value
                amount_value = ws.cell(amount_row, col).value
                rate_kind = cell_kind(rate_value)
                amount_kind = cell_kind(amount_value)
                if rate_kind != amount_kind:
                    raise ValueError(
                        f"{block.sheet} redak {rate_row}, {date:%Y-%m-%d}: "
                        "raspolozivost stope i iznosa nije jednaka."
                    )
                if rate_kind == "BLANK":
                    blank += 1
                    continue
                elif rate_kind == "DASH":
                    dash += 1
                    continue

                numeric += 1
                dimensions = {
                    "datum": date.strftime("%Y-%m-%d"),
                    "pozicija": meta.pozicija,
                    "sektor_protustranke": meta.sektor,
                    "podskup_protustranke": meta.podskup,
                    "instrument": meta.instrument,
                    "namjena_kredita": meta.namjena,
                    "izvorno_dospijece": meta.dospijece,
                    "otkazni_rok": meta.otkazni_rok,
                    "pocetno_fiksiranje_kamatne_stope": meta.fiksiranje,
                    "velicina_kredita": meta.velicina,
                }
                key = tuple(dimensions[field] for field in DIMENSION_FIELDS)
                group = groups.setdefault(
                    key,
                    {"dimensions": dimensions, "candidates": defaultdict(list)},
                )
                priority = candidate_priority(block.sheet)

                if block.vrsta_stope == "EFEKTIVNA":
                    rate_measure = "efektivna_kamatna_stopa"
                elif meta.osnova == NOVI:
                    rate_measure = "nominalna_kamatna_stopa_novi_posao"
                else:
                    rate_measure = "nominalna_kamatna_stopa_stanje"
                group["candidates"][rate_measure].append(
                    Candidate(csv_number(rate_value), block.sheet, priority)
                )

                # Iznos uz EKS namjerno se ne izvozi; nominalni blok G2 jest
                # kanonski izvor mjere iznos_novi_posao.
                if block.vrsta_stope == "EFEKTIVNA":
                    ignored_effective_amounts += 1
                    continue
                amount_measure = "iznos_novi_posao" if meta.osnova == NOVI else "iznos_stanje"
                group["candidates"][amount_measure].append(
                    Candidate(csv_number(amount_value), block.sheet, priority)
                )

    output: list[dict[str, Any]] = []
    selected_measure_counts = {measure: 0 for measure in MEASURE_FIELDS}
    discarded_lower_priority = 0
    for key in sorted(groups):
        group = groups[key]
        row = dict(group["dimensions"])
        sources: set[str] = set()
        for measure in MEASURE_FIELDS:
            candidates = group["candidates"].get(measure, [])
            if not candidates:
                row[measure] = ""
                continue
            chosen = choose_candidate(candidates, measure, key)
            row[measure] = chosen.value
            sources.add(chosen.source)
            selected_measure_counts[measure] += 1
            discarded_lower_priority += sum(
                candidate.priority > chosen.priority for candidate in candidates
            )
        row["izvorne_tablice"] = "|".join(sorted(sources))
        output.append(row)

    summary = {
        "input": input_label,
        "prvi_datum": common_dates[0].strftime("%Y-%m-%d") if common_dates else None,
        "zadnji_datum": common_dates[-1].strftime("%Y-%m-%d") if common_dates else None,
        "broj_mjeseci": len(common_dates or []),
        "broj_izvornih_serija": source_series,
        "broj_mogucih_redaka": possible,
        "broj_numerickih_parova_u_izvoru": numeric,
        "broj_izvornih_znakova_minus": dash,
        "broj_praznih_parova_u_izvoru": blank,
        "broj_ignoriranih_iznosa_uz_EKS": ignored_effective_amounts,
        "broj_kandidata_odbacenih_zbog_prioriteta": discarded_lower_priority,
        "broj_nepraznih_mjera": selected_measure_counts,
        "broj_izvezenih_redaka": len(output),
    }
    return output, summary


def build_rows(input_path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    workbook = load_workbook(input_path, data_only=True, read_only=False)
    return build_rows_from_workbook(workbook, str(input_path))


def write_csv(rows: list[dict[str, Any]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=FIELDS,
            delimiter=";",
            quotechar='"',
            quoting=csv.QUOTE_MINIMAL,
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Ulazna Excel knjiga G tablice.xlsx")
    parser.add_argument("output", type=Path, nargs="?", default=Path("G_podaci.csv"), help="Izlazni CSV")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    rows, summary = build_rows(args.input)
    if summary["broj_izvornih_serija"] != 142:
        raise ValueError(
            f"Ocekivane su 142 izvorne serije, pronadjeno je {summary['broj_izvornih_serija']}."
        )
    write_csv(rows, args.output)
    summary["output"] = str(args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
