#!/usr/bin/env python3
"""Render a .pptx to PDF (and optionally slide images) to close the visual feedback loop.

python-pptx has no rendering engine, so a deck can be structurally valid yet look
wrong — text overflowing a box, an image off the slide, an unreadable color on a
fill. We use LibreOffice (``soffice --headless``) as the reference renderer: one
PDF page per slide. With ``--render`` each slide is rasterized to PNG so the agent
can *see* the deck before delivering it.

Requires LibreOffice (``soffice``). Slide rendering additionally needs Poppler
(``pdftoppm``); slide counting falls back to ``pypdf`` then ``pdfinfo``.

Examples
--------
    # slide count only
    python scripts/render_pptx.py deck.pptx

    # render every slide to PNG at 120 dpi for inspection
    python scripts/render_pptx.py deck.pptx --render --dpi 120

    # render just slides 1-3
    python scripts/render_pptx.py deck.pptx --render --slides 1-3
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


def convert_to_pdf(pptx_path: str, outdir: str, timeout: int = 240) -> str:
    """Convert ``pptx_path`` to PDF in ``outdir`` via LibreOffice. Return the PDF path."""
    soffice = _find("soffice", "libreoffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install it (Rocky/RHEL: "
            "`sudo dnf install -y libreoffice-impress libreoffice-core`; "
            "Debian/Ubuntu: `sudo apt-get install -y libreoffice-impress`)."
        )
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        env = dict(os.environ, HOME=profile)
        proc = subprocess.run(
            [soffice, "--headless", f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", outdir, pptx_path],
            capture_output=True, text=True, timeout=timeout, env=env,
        )
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError(f"PDF not produced.\nstdout: {proc.stdout}\nstderr: {proc.stderr}")
    return pdf


def slide_count(pdf_path: str) -> int:
    """Best-effort slide count (one PDF page per slide): pypdf, then Poppler ``pdfinfo``."""
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
    raise RuntimeError("Could not determine slide count (install pypdf or poppler-utils).")


def render_slides(pdf_path: str, outdir: str, dpi: int, slides: str | None) -> list[str]:
    """Rasterize PDF pages (slides) to PNG via Poppler ``pdftoppm``. Return the image paths."""
    pdftoppm = _find("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("pdftoppm not found. Install Poppler (poppler-utils).")
    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    prefix = os.path.join(outdir, stem)
    cmd = [pdftoppm, "-png", "-r", str(dpi)]
    if slides:
        first, _, last = slides.partition("-")
        cmd += ["-f", first, "-l", last or first]
    cmd += [pdf_path, prefix]
    subprocess.run(cmd, check=True, capture_output=True)
    return sorted(
        os.path.join(outdir, f) for f in os.listdir(outdir)
        if f.startswith(stem + "-") and f.endswith(".png")
    )


def main() -> int:
    ap = argparse.ArgumentParser(description="Render a .pptx to PDF/PNG to verify slide layout.")
    ap.add_argument("pptx", help="Path to the .pptx file")
    ap.add_argument("--outdir", default=None, help="Output dir (default: alongside the pptx)")
    ap.add_argument("--render", action="store_true", help="Also rasterize slides to PNG")
    ap.add_argument("--dpi", type=int, default=110, help="PNG resolution (default 110)")
    ap.add_argument("--slides", default=None, help="Slide range to render, e.g. '1-3' or '2'")
    ap.add_argument("--expect-slides", type=int, default=None,
                    help="Assert the deck has exactly N slides; exit 2 if not")
    args = ap.parse_args()

    pptx = os.path.abspath(args.pptx)
    if not os.path.exists(pptx):
        print(json.dumps({"ok": False, "error": f"not found: {pptx}"})); return 1
    outdir = os.path.abspath(args.outdir) if args.outdir else os.path.dirname(pptx)

    try:
        pdf = convert_to_pdf(pptx, outdir)
        slides = slide_count(pdf)
        images = render_slides(pdf, outdir, args.dpi, args.slides) if args.render else []
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)})); return 1

    result = {"ok": True, "pdf": pdf, "slides": slides, "images": images}
    if args.expect_slides is not None:
        result["expected_slides"] = args.expect_slides
        result["slides_match"] = (slides == args.expect_slides)
    print(json.dumps(result, indent=2))
    if args.expect_slides is not None and slides != args.expect_slides:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
