#!/usr/bin/env python3
"""Render a .xlsx to PDF (and optionally page images) to close the visual feedback loop.

openpyxl writes cells but can't show you the printed result. We use LibreOffice
(``soffice --headless``) as the reference renderer to catch what validation can't:
columns cut off at a page edge, a chart that didn't render, number formats that
overflow, or a sheet that explodes into dozens of pages. With ``--render`` each
PDF page is rasterized to PNG for inspection.

IMPORTANT — spreadsheet pagination is print-settings dependent. Page count and
column breaks follow each sheet's **print area, page setup, and fit-to-page
scaling**. A workbook with no print setup can paginate very differently from how
it looks on screen. If you care about the printed layout, set print area /
``fitToWidth`` on the sheet first (openpyxl ``ws.page_setup``), then render. By
default LibreOffice exports **all** sheets.

Requires LibreOffice (``soffice``). Page rendering additionally needs Poppler
(``pdftoppm``); page counting falls back to ``pypdf`` then ``pdfinfo``.

Examples
--------
    # page count only (across all sheets, per current print settings)
    python scripts/render_xlsx.py book.xlsx

    # render every page to PNG at 120 dpi for inspection
    python scripts/render_xlsx.py book.xlsx --render --dpi 120

    # render just pages 1-2
    python scripts/render_xlsx.py book.xlsx --render --pages 1-2
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile


def _find(*names: str) -> str | None:
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def convert_to_pdf(xlsx_path: str, outdir: str, timeout: int = 240) -> str:
    """Convert ``xlsx_path`` to PDF in ``outdir`` via LibreOffice. Return the PDF path."""
    soffice = _find("soffice", "libreoffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install it (Rocky/RHEL: "
            "`sudo dnf install -y libreoffice-calc libreoffice-core`; "
            "Debian/Ubuntu: `sudo apt-get install -y libreoffice-calc`)."
        )
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        env = dict(os.environ, HOME=profile)
        proc = subprocess.run(
            [soffice, "--headless", f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", outdir, xlsx_path],
            capture_output=True, text=True, timeout=timeout, env=env,
        )
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError(f"PDF not produced.\nstdout: {proc.stdout}\nstderr: {proc.stderr}")
    return pdf


def page_count(pdf_path: str) -> int:
    """Best-effort page count: pypdf, then Poppler ``pdfinfo``."""
    try:
        from pypdf import PdfReader
        return len(PdfReader(pdf_path).pages)
    except Exception:
        pass
    pdfinfo = _find("pdfinfo")
    if pdfinfo:
        out = subprocess.run([pdfinfo, pdf_path], capture_output=True, text=True).stdout
        for line in out.splitlines():
            if line.lower().startswith("pages"):
                return int(line.split()[-1])
    raise RuntimeError("Could not determine page count (install pypdf or poppler-utils).")


def render_pages(pdf_path: str, outdir: str, dpi: int, pages: str | None) -> list[str]:
    """Rasterize PDF pages to PNG via Poppler ``pdftoppm``. Return the image paths."""
    pdftoppm = _find("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("pdftoppm not found. Install Poppler (poppler-utils).")
    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    prefix = os.path.join(outdir, stem)
    cmd = [pdftoppm, "-png", "-r", str(dpi)]
    if pages:
        first, _, last = pages.partition("-")
        cmd += ["-f", first, "-l", last or first]
    cmd += [pdf_path, prefix]
    subprocess.run(cmd, check=True, capture_output=True)
    return sorted(
        os.path.join(outdir, f) for f in os.listdir(outdir)
        if f.startswith(stem + "-") and f.endswith(".png")
    )


def main() -> int:
    ap = argparse.ArgumentParser(description="Render a .xlsx to PDF/PNG to verify printed layout.")
    ap.add_argument("xlsx", help="Path to the .xlsx file")
    ap.add_argument("--outdir", default=None, help="Output dir (default: alongside the xlsx)")
    ap.add_argument("--render", action="store_true", help="Also rasterize pages to PNG")
    ap.add_argument("--dpi", type=int, default=110, help="PNG resolution (default 110)")
    ap.add_argument("--pages", default=None, help="Page range to render, e.g. '1-2' or '3'")
    ap.add_argument("--expect-pages", type=int, default=None,
                    help="Assert the workbook prints to exactly N pages; exit 2 if not")
    args = ap.parse_args()

    xlsx = os.path.abspath(args.xlsx)
    if not os.path.exists(xlsx):
        print(json.dumps({"ok": False, "error": f"not found: {xlsx}"})); return 1
    outdir = os.path.abspath(args.outdir) if args.outdir else os.path.dirname(xlsx)

    try:
        pdf = convert_to_pdf(xlsx, outdir)
        pages = page_count(pdf)
        images = render_pages(pdf, outdir, args.dpi, args.pages) if args.render else []
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)})); return 1

    result = {
        "ok": True, "pdf": pdf, "pages": pages, "images": images,
        "note": "Page count follows each sheet's print area / page setup / fit-to-page scaling.",
    }
    if args.expect_pages is not None:
        result["expected_pages"] = args.expect_pages
        result["pages_match"] = (pages == args.expect_pages)
    print(json.dumps(result, indent=2))
    if args.expect_pages is not None and pages != args.expect_pages:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
