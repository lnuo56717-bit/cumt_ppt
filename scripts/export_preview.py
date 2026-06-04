#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Export PPTX slides to PNG previews using PowerPoint or LibreOffice.

On Windows with PowerPoint installed, this uses COM automation. Otherwise it
tries LibreOffice/soffice. If neither is available, it exits with a clear
message and nonzero status.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def export_with_powerpoint(pptx: Path, out_dir: Path, width: int, height: int) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # type: ignore
    except Exception:
        try:
            import comtypes.client  # type: ignore
        except Exception:
            return False
        app = comtypes.client.CreateObject("PowerPoint.Application")
    else:
        app = win32com.client.Dispatch("PowerPoint.Application")

    presentation = app.Presentations.Open(str(pptx), False, False, True)
    presentation.Export(str(out_dir), "PNG", width, height)
    presentation.Close()
    return True


def export_with_libreoffice(pptx: Path, out_dir: Path) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "png",
        "--outdir",
        str(out_dir),
        str(pptx),
    ]
    subprocess.run(cmd, check=True)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Export PPTX preview PNGs.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("preview_dir", type=Path)
    parser.add_argument("--width", type=int, default=1600)
    parser.add_argument("--height", type=int, default=900)
    args = parser.parse_args()

    args.preview_dir.mkdir(parents=True, exist_ok=True)
    try:
        if export_with_powerpoint(args.pptx.resolve(), args.preview_dir.resolve(), args.width, args.height):
            print(f"Exported with PowerPoint to {args.preview_dir}")
            return 0
    except Exception as exc:
        print(f"PowerPoint export failed: {exc}")

    try:
        if export_with_libreoffice(args.pptx.resolve(), args.preview_dir.resolve()):
            print(f"Exported with LibreOffice to {args.preview_dir}")
            return 0
    except Exception as exc:
        print(f"LibreOffice export failed: {exc}")

    print("No supported preview exporter found. Install Microsoft PowerPoint on Windows or LibreOffice/soffice.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
