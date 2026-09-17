"""Refresh the final DE-Cividis paper artifacts.

This file is now the safe entry point for the recovered paper.  It does not
rewrite the paper body from an old template.  Instead it refreshes experiment
tables, regenerates the figures used in the paper, reapplies the Word layout
polish, and exports the final PDF when LibreOffice is available.

Usage:
    python code/write_paper_decividis_final.py
    python code/write_paper_decividis_final.py --with-experiments
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CODE_DIR = ROOT / "code"
PAPER_DIR = ROOT / "\u8bba\u6587"
DOCX = PAPER_DIR / (
    "\u8ba1\uff08\u521b\uff092302\u73ed-\u5218\u4e16\u7fd4-XXXXXXX-"
    "\u6570\u5b57\u56fe\u50cf\u5904\u7406\u6280\u672f\u7ed3\u8bfe\u8bba\u6587.docx"
)
PDF = DOCX.with_suffix(".pdf")


def run_python(script_name: str) -> None:
    script = CODE_DIR / script_name
    if not script.exists():
        raise FileNotFoundError(script)
    print(f"[run] {script_name}")
    subprocess.run([sys.executable, str(script)], cwd=str(ROOT), check=True)


def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def export_pdf() -> None:
    soffice = find_soffice()
    if not soffice:
        print("[warn] LibreOffice not found; DOCX was refreshed but PDF was not exported.")
        return
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)

    tmp_dir = CODE_DIR / "_pdf_export_tmp"
    tmp_dir.mkdir(exist_ok=True)
    tmp_docx = tmp_dir / "paper.docx"
    shutil.copy2(DOCX, tmp_docx)

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(tmp_dir),
        str(tmp_docx),
    ]
    print("[run] export pdf")
    subprocess.run(cmd, cwd=str(ROOT), check=True)
    tmp_pdf = tmp_dir / "paper.pdf"
    if not tmp_pdf.exists():
        raise FileNotFoundError(tmp_pdf)
    try:
        shutil.copy2(tmp_pdf, PDF)
        print(f"[saved] {PDF}")
    except PermissionError:
        fallback_pdf = PDF.with_name(PDF.stem + "_\u4fee\u8ba2\u7248.pdf")
        shutil.copy2(tmp_pdf, fallback_pdf)
        print(f"[warn] target PDF is in use; saved fallback PDF: {fallback_pdf}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--with-experiments",
        action="store_true",
        help="Recompute the full metric CSV files before refreshing figures and layout.",
    )
    parser.add_argument("--skip-pdf", action="store_true")
    args = parser.parse_args()

    if args.with_experiments:
        run_python("run_final_experiments.py")
    run_python("make_polished_figures.py")
    run_python("polish_cjig_layout.py")
    if not args.skip_pdf:
        export_pdf()


if __name__ == "__main__":
    main()
