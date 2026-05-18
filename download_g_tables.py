#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Preuzima HNB tablice G1, G2, G3, G5 i G6 s javne HNB stranice
i spaja ih u jednu Excel datoteku "G tablice.xlsx".

Značajke:
- sam pronalazi aktualne HNB poveznice na stranici
- svaki izvorni list kopira u zaseban list: G1, G2, G3, G5, G6
- dodaje prvi list "Metodologija"
- u list "Metodologija" upisuje tekst s HNB stranice počevši od retka koji počinje s "Metodologija"
- svaki redak/odlomak metodologije ide u zaseban redak stupca A
- redci koji počinju s "Metodologija" ili "Tablica G" su bold
- prije svakog retka koji počinje s "Tablica G" dodaje se prazan red
- stupac A na listu Metodologija širok je približno 700 px
- u svim G listovima uključuje Wrap text za stupac B
- prilagođava visinu redaka
- uklanja boju kartica listova
- isključuje prikaz crta rešetke
"""

from __future__ import annotations

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
from xml.etree import ElementTree as ET

import requests
from bs4 import BeautifulSoup


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

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    raw_lines = soup.get_text("\n", strip=True).splitlines()
    lines = [line.strip() for line in raw_lines if line.strip()]

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

    methodology = "\n".join(lines[start_index:end_index]).strip()

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
    from com.sun.star.beans import PropertyValue

    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def path_to_file_url(path: Path) -> str:
    import uno

    return uno.systemPathToFileUrl(str(path.resolve()))


def set_gridlines_off(out_doc, sheet) -> None:
    try:
        out_doc.CurrentController.setActiveSheet(sheet)
        out_doc.CurrentController.ShowGrid = False
    except Exception:
        pass


def set_rows_optimal_height(sheet, first_row: int, last_row: int) -> None:
    try:
        for r in range(first_row, last_row + 1):
            sheet.Rows.getByIndex(r).OptimalHeight = True
    except Exception:
        pass


def px_to_hundredths_mm(px: int, dpi: int = 96) -> int:
    """LibreOffice column widths use 1/100 mm. 700 px at 96 DPI ≈ 18521."""
    return round(px / dpi * 25.4 * 100)


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

    # ------------------------------------------------------------------
    # List Metodologija
    # ------------------------------------------------------------------
    methodology_sheet = out_sheets.getByIndex(0)
    methodology_sheet.Name = "Metodologija"

    set_gridlines_off(out_doc, methodology_sheet)

    # Column A width: approximately 700 pixels.
    methodology_sheet.Columns.getByName("A").Width = px_to_hundredths_mm(700)

    try:
        methodology_sheet.Columns.getByName("A").IsTextWrapped = True
    except Exception:
        pass

    methodology_lines = [
        line.strip()
        for line in methodology_text.splitlines()
        if line.strip()
    ]

    output_row = 0
    for line in methodology_lines:
        # Add an empty row immediately above each "Tablica G..." heading,
        # except if it would be the first row in the sheet.
        if line.startswith("Tablica G") and output_row > 0:
            output_row += 1

        cell = methodology_sheet.getCellByPosition(0, output_row)  # A1, A2, A3...
        cell.String = line
        cell.IsTextWrapped = True
        cell.CharFontName = "Arial"
        cell.CharHeight = 10

        if line.startswith("Metodologija") or line.startswith("Tablica G"):
            cell.CharWeight = 150  # bold

        output_row += 1

    if output_row > 0:
        set_rows_optimal_height(methodology_sheet, 0, output_row - 1)

    # ------------------------------------------------------------------
    # Uvoz G tablica
    # ------------------------------------------------------------------
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

            try:
                imported_sheet.TabColor = -1
            except Exception:
                pass

            set_gridlines_off(out_doc, imported_sheet)

            # Wrap text za stupac B.
            try:
                imported_sheet.Columns.getByName("B").IsTextWrapped = True
            except Exception:
                pass

            # Autofit visine redaka u korištenom području.
            try:
                cursor = imported_sheet.createCursor()
                cursor.gotoEndOfUsedArea(True)
                used_range = cursor.RangeAddress
                set_rows_optimal_height(
                    imported_sheet,
                    used_range.StartRow,
                    used_range.EndRow,
                )
            except Exception:
                set_rows_optimal_height(imported_sheet, 0, 2000)

            position += 1

        finally:
            src_doc.close(True)

    # Još jednom isključi gridlines na svim listovima.
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
    """

    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ET.register_namespace("", ns["x"])

    tmp_path = xlsx_path.with_suffix(".tmp.xlsx")

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                root = ET.fromstring(data)

                sheet_pr = root.find("x:sheetPr", ns)
                if sheet_pr is not None:
                    for tab_color in list(sheet_pr.findall("x:tabColor", ns)):
                        sheet_pr.remove(tab_color)

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
