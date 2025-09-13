"""Build Windows executables using PyInstaller.

This script creates two standalone executables for the CLI and GUI entry
points. PyInstaller must be installed in the active environment.
Run on Windows:
    python build_exe.py
The built binaries will appear in the ``dist`` directory.
"""
from __future__ import annotations

import os
from pathlib import Path
import PyInstaller.__main__  # type: ignore

ROOT = Path(__file__).resolve().parent

# Data files required at runtime
DATA_FILES = ["rules.yaml", "INSTRUCTION.md"]


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
    ]
    if windowed:
        opts.append("--windowed")

    sep = os.pathsep  # ";" on Windows, ":" on POSIX
    for f in DATA_FILES:
        opts += ["--add-data", f"{ROOT / f}{sep}." ]

    PyInstaller.__main__.run(opts)


def main() -> None:
    _build("main_cli.py", windowed=False, name="xmlchecks_cli")
    _build("main_gui.py", windowed=True, name="xmlchecks_gui")


if __name__ == "__main__":
    main()
