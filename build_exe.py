"""Build Windows executables using PyInstaller.

This script creates a standalone executable for the GUI entry
point. PyInstaller must be installed in the active environment.
Run on Windows:
    python build_exe.py
The built binary will appear in the project root directory.
"""
from __future__ import annotations

import os
from pathlib import Path
import PyInstaller.__main__  # type: ignore

ROOT = Path(__file__).resolve().parent

# Data files required at runtime
DATA_FILES = ["rules.yaml", "INSTRUCTION.md"]
TESSERACT_DIR = ROOT / "tesseract"


def _build(target: str, *, windowed: bool, name: str) -> None:
    """Run PyInstaller for a single target.

    Args:
        target: The entry-point script relative to the project root.
        windowed: Whether to build a GUI application without console.
    """

    opts: list[str] = [
        str(ROOT / target),
        "--onefile",
        "--name",
        name,
        "--distpath",
        str(ROOT),
    ]
    if windowed:
        opts.append("--windowed")

    sep = os.pathsep  # ";" on Windows, ":" on POSIX
    for f in DATA_FILES:
        opts += ["--add-data", f"{ROOT / f}{sep}." ]

    if TESSERACT_DIR.exists():
        opts += ["--add-data", f"{TESSERACT_DIR}{sep}tesseract"]

    PyInstaller.__main__.run(opts)


def main() -> None:
    _build("main_gui.py", windowed=True, name="IFCChecks")


if __name__ == "__main__":
    main()
