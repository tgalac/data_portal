#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Preuzima HNB tablice G1, G2, G3, G5 i G6 s javne HNB stranice
i spaja ih u jednu Excel datoteku "G tablice.xlsx".

Značajke:
- sam pronalazi aktualne HNB poveznice na stranici
- svaki izvorni list kopira u zaseban list: G1, G2, G3, G5, G6
- dodaje prvi list "Metodologija"
- u list "Metodologija" dodaje text box s metodološkim tekstom sa stranice
- uklanja boju kartica listova
- isključuje prikaz crta rešetke
- koristi LibreOffice headless radi boljeg očuvanja izvornog formatiranja
"""

from __future__ import annotations

import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from xml.etree import ElementTree as ET


HNB_PAGE_URL = (
    "https://www.hnb.hr/statistika/statisticki-podaci/financijski-sektor/"
    "druge-monetarne-financijske-institucije/kreditne-institucije/kamatne-stope"
)

TABLES = ["G1", "G2", "G3", "G5", "G6"]
OUTPUT_FILE = "G tablice.xlsx"


def get_html(url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 Chrome/120 Safari/537.36"
        )
    }
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    response.encoding = response.apparent_encoding
    return response.text


def find_table_links(html: str, base_url: str) -> dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    links: dict[str, str] = {}

    for a in soup.find_all("a", href=True):
        text = " ".join(a.get_text(" ", strip=True).split())

        for table in TABLES:
            # Traži linkove čiji tekst počinje s "Tablica G1", "Tablica G2" itd.
            if re.search(rf"\bTablica\s+{table}\b", text, flags=re.IGNORECASE):
                links[table] = urljoin(base_url, a["href"])

    missing = [t for t in TABLES if t not in links]
    if missing:
        raise RuntimeError(
            f"Nisam pronašao poveznice za: {', '.join(missing)}. "
            "Provjeri je li HNB promijenio strukturu stranice."
        )

    return {t: links[t] for t in TABLES}


def extract_methodology_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    # Ukloni skripte i stilove.
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    raw_lines = soup.get_text("\n", strip=True).splitlines()
    lines = [line.strip() for line in raw_lines if line.strip()]

    # Traži prvi red koji počinje s "Metodologija".
    start_index = None
    for i, line in enumerate(lines):
        if line.startswith("Metodologija"):
            start_index = i
            break

    if start_index is None:
        raise RuntimeError(
            'Nisam pronašao redak koji počinje s "Metodologija". '
            "Moguće je da je HNB promijenio strukturu stranice."
        )

    # Tipični početak podnožja / kraja relevantnog sadržaja.
    end_markers = [
        "Skriveno",
        "HRVATSKA NARODNA BANKA",
        "Trg hrvatskih velikana",
        "Kontakt",
        "Pristupačnost",
    ]

    end_index = len(lines)
    for i in range(start_index + 1, len(lines)):
        if any(lines[i].startswith(marker) for marker in end_markers):
            end_index = i
            break

    methodology_lines = lines[start_index:end_index]
    methodology = "\n".join(methodology_lines).strip()

    if len(methodology) < 100:
        raise RuntimeError(
            "Metodološki tekst je pronađen, ali je neočekivano kratak. "
            "Provjeri ekstrakciju HTML-a."
        )

    return methodology

def download_file(url: str, target_path: Path) -> None:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 Chrome/120 Safari/537.36"
        )
    }

    with requests.get(url, headers=headers, stream=True, timeout=120) as response:
        response.raise_for_status()
        with open(target_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if chunk:
                    f.write(chunk)

    if target_path.stat().st_size == 0:
        raise RuntimeError(f"Preuzeta datoteka je prazna: {target_path}")


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_for_libreoffice(port: int, timeout_seconds: int = 30) -> None:
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return
        except OSError:
            time.sleep(0.5)

    raise RuntimeError("LibreOffice se nije pokrenuo na vrijeme.")


def start_libreoffice(port: int, user_profile_dir: Path) -> subprocess.Popen:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice nije pronađen. Instaliraj LibreOffice ili ga dodaj u PATH."
        )

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--nodefault",
        "--norestore",
        f"-env:UserInstallation=file://{user_profile_dir.as_posix()}",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager",
    ]

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    wait_for_libreoffice(port)
    return proc


def uno_property(name: str, value):
    import uno
    from com.sun.star.beans import PropertyValue

    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def path_to_file_url(path: Path) -> str:
    import uno

    return uno.systemPathToFileUrl(str(path.resolve()))


def make_workbook_with_libreoffice(
    xls_paths: dict[str, Path],
    methodology_text: str,
    output_path: Path,
    port: int,
) -> None:
    import uno

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx,
    )
    ctx = resolver.resolve(
        f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    hidden = (uno_property("Hidden", True),)

    # Nova prazna Calc radna knjiga.
    out_doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, hidden)
    out_sheets = out_doc.Sheets

    # Prvi list: Metodologija.
methodology_sheet = out_sheets.getByIndex(0)
methodology_sheet.Name = "Metodologija"

# Isključi gridlines na listu Metodologija.
try:
    out_doc.CurrentController.setActiveSheet(methodology_sheet)
    out_doc.CurrentController.ShowGrid = False
except Exception:
    pass

# Column A width: 120 points.
# LibreOffice UNO koristi 1/100 mm.
# 120 pt = 120 / 72 inch = 1.6667 inch = 42.333 mm = approx. 4233 in 1/100 mm.
methodology_sheet.Columns.getByName("A").Width = 4233

# Kopiraj metodološki tekst u ćeliju A1.
cell_a1 = methodology_sheet.getCellByPosition(0, 0)
cell_a1.String = methodology_text

# Osnovno formatiranje ćelije A1.
# Ovo ne može savršeno preslikati HTML source formatting, ali čuva tekst i retke
# te daje Excelu normalan prikaz.
cell_a1.IsTextWrapped = True
cell_a1.CharFontName = "Arial"
cell_a1.CharHeight = 10

# Prvi redak metodologije učini bold.
# LibreOffice cell rich text zna biti osjetljiv, pa je ovo u try bloku.
try:
    cursor = cell_a1.createTextCursor()
    cursor.gotoStart(False)
    first_newline = methodology_text.find("\n")

    if first_newline != -1:
        cursor.goRight(first_newline, True)
    else:
        cursor.gotoEnd(True)

    cursor.CharWeight = 150  # bold
except Exception:
    pass

# Wrap text za cijeli stupac A.
try:
    methodology_sheet.Columns.getByName("A").IsTextWrapped = True
except Exception:
    pass

# Autofit visine redaka na cijelom listu Metodologija.
try:
    cursor = methodology_sheet.createCursor()
    cursor.gotoEndOfUsedArea(True)
    used_range = cursor.RangeAddress

    for r in range(used_range.StartRow, used_range.EndRow + 1):
        methodology_sheet.Rows.getByIndex(r).OptimalHeight = True
except Exception:
    pass

    # Uvezi prvi list iz svake izvorne datoteke.
    position = 1
    for table in TABLES:
        src_path = xls_paths[table]
        src_doc = desktop.loadComponentFromURL(path_to_file_url(src_path), "_blank", 0, hidden)

        try:
            src_sheet = src_doc.Sheets.getByIndex(0)
            src_sheet_name = src_sheet.Name

            out_sheets.importSheet(src_doc, src_sheet_name, position)
            imported_sheet = out_sheets.getByIndex(position)
            imported_sheet.Name = table

            # Ukloni boju kartice lista, ako LibreOffice podržava svojstvo.
            try:
                imported_sheet.TabColor = -1
            except Exception:
                pass

            # Isključi gridlines.
            try:
                out_doc.CurrentController.setActiveSheet(imported_sheet)
                out_doc.CurrentController.ShowGrid = False
            except Exception:
                pass

            # Uključi Wrap text za stupac B.
            try:
                col_b_range = imported_sheet.Columns.getByName("B")
                col_b_range.IsTextWrapped = True
            except Exception:
                pass

            # Automatski prilagodi visinu redaka.
            # Pokušaj najprije za sve korištene retke, a ako to ne uspije, za prvih 2000.
            try:
                cursor = imported_sheet.createCursor()
                cursor.gotoEndOfUsedArea(True)
                used_range = cursor.RangeAddress
                first_row = used_range.StartRow
                last_row = used_range.EndRow

                for r in range(first_row, last_row + 1):
                    imported_sheet.Rows.getByIndex(r).OptimalHeight = True
            except Exception:
                try:
                    for r in range(0, 2000):
                        imported_sheet.Rows.getByIndex(r).OptimalHeight = True
                except Exception:
                    pass

            position += 1
        finally:
            src_doc.close(True)

    # Još jednom prođi sve listove i isključi gridlines.
    try:
        controller = out_doc.CurrentController
        for i in range(out_sheets.Count):
            controller.setActiveSheet(out_sheets.getByIndex(i))
            controller.ShowGrid = False
    except Exception:
        pass

    # Spremi kao XLSX.
    store_props = (
        uno_property("FilterName", "Calc MS Excel 2007 XML"),
        uno_property("Overwrite", True),
    )

    out_doc.storeAsURL(path_to_file_url(output_path), store_props)
    out_doc.close(True)

def patch_xlsx_sheet_xml(xlsx_path: Path) -> None:
    """
    Nakon što LibreOffice napravi XLSX, izravno se uređuju sheet XML datoteke:
    - uklanjaju se tabColor elementi
    - showGridLines se postavlja na 0

    Ovo ne otvara i ponovno ne sprema workbook preko openpyxl-a, jer bi to moglo
    izbaciti text box/drawing objekte.
    """

    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ET.register_namespace("", ns["x"])

    tmp_path = xlsx_path.with_suffix(".tmp.xlsx")

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                root = ET.fromstring(data)

                # Ukloni <tabColor .../> iz <sheetPr>.
                sheet_pr = root.find("x:sheetPr", ns)
                if sheet_pr is not None:
                    for tab_color in list(sheet_pr.findall("x:tabColor", ns)):
                        sheet_pr.remove(tab_color)

                # Postavi showGridLines="0" na svim sheetView elementima.
                for sheet_view in root.findall(".//x:sheetView", ns):
                    sheet_view.set("showGridLines", "0")

                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            zout.writestr(item, data)

    tmp_path.replace(xlsx_path)


def main() -> None:
    output_path = Path(OUTPUT_FILE).resolve()

    print("Dohvaćam HNB stranicu...")
    html = get_html(HNB_PAGE_URL)

    print("Pronalazim poveznice na G tablice...")
    links = find_table_links(html, HNB_PAGE_URL)

    print("Izdvajam metodološki tekst...")
    methodology_text = extract_methodology_text(html)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        downloads_dir = tmp_dir / "downloads"
        downloads_dir.mkdir()

        xls_paths: dict[str, Path] = {}

        for table, url in links.items():
            target = downloads_dir / f"{table}.xls"
            print(f"Preuzimam {table}: {url}")
            download_file(url, target)
            xls_paths[table] = target

        lo_profile = tmp_dir / "lo-profile"
        lo_profile.mkdir()

        port = find_free_port()
        print("Pokrećem LibreOffice headless...")
        lo_proc = start_libreoffice(port, lo_profile)

        try:
            print(f"Kreiram {output_path.name}...")
            make_workbook_with_libreoffice(
                xls_paths=xls_paths,
                methodology_text=methodology_text,
                output_path=output_path,
                port=port,
            )
        finally:
            lo_proc.terminate()
            try:
                lo_proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                lo_proc.kill()

    print("Podešavam gridlines i boje kartica listova u XLSX XML-u...")
    patch_xlsx_sheet_xml(output_path)

    print(f"Gotovo: {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"GREŠKA: {exc}", file=sys.stderr)
        sys.exit(1)
