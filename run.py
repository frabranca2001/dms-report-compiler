"""Build the DMS reports. Edit the SETTINGS below, then just run this file
(F5 in VS Code / Visual Studio, or double-click "Build reports.bat").
"""
# ======================================================================
# SETTINGS - edit these, nothing else
# Paths can be absolute (r"C:\Reports\template.docm") or relative to
# the folder this file is in.
# ======================================================================

# The DMS template (.docx or .docm). It must contain a paragraph with the
# text {{SCADE_CHAPTERS}} where the SCADE chapters have to go.
TEMPLATE = r"examples\dms_template.docx"

# Excel sheet listing the reports to build: one row per report.
#   column "scade_report": the SCADE report file (relative to the sheet);
#                          .docx, .rtf or .doc (.rtf/.doc need USE_WORD = True)
#   column "output":       name of the report to create (optional)
#   other columns:         values for the {{placeholders}} in the template
REPORTS_SHEET = r"examples\reports.xlsx"

# Folder where the finished reports are written.
OUTPUT_FOLDER = r"output"

EXPORT_PDF = True          # also save a PDF next to each report
USE_WORD = True            # let Word refresh table of contents / page numbers
OPEN_OUTPUT_FOLDER = True  # open the output folder when finished

# ---- advanced (usually leave as is) ----------------------------------
ADVANCED = {
    # text of the marker paragraph in the template
    "marker": "SCADE_CHAPTERS",
    # regex of the first chapter to take; None = first level-1 heading
    "start_heading": None,
    # regex of a level-1 heading where to stop (not included); None = end
    "end_heading": None,
    # SCADE style name -> template style name, when the names differ
    "style_map": {},
    # "auto" / True / False: remove chapter numbers typed by hand
    "strip_typed_numbers": "auto",
    # keep section breaks of the SCADE report (e.g. landscape pages)
    "keep_section_breaks": False,
}
# ======================================================================

import importlib.util
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def ensure_packages():
    """Install the required libraries the first time, so no prompt is needed."""
    needed = {"lxml": "lxml", "openpyxl": "openpyxl"}
    if USE_WORD:
        needed["win32com"] = "pywin32"
    missing = [pkg for mod, pkg in needed.items()
               if importlib.util.find_spec(mod) is None]
    if missing:
        print("First run: installing", ", ".join(missing), "...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        importlib.invalidate_caches()


def main():
    ensure_packages()
    sys.path.insert(0, str(HERE))
    import dms_compiler

    path = lambda p: str(p if Path(p).is_absolute() else HERE / p)
    out = path(OUTPUT_FOLDER)
    args = ["--template", path(TEMPLATE), "--batch", path(REPORTS_SHEET),
            "--out", out]
    if EXPORT_PDF and USE_WORD:
        args.append("--pdf")
    if not USE_WORD:
        args.append("--no-word")
    code = dms_compiler.main(args, config=ADVANCED)
    if OPEN_OUTPUT_FOLDER and os.path.isdir(out):
        os.startfile(out)
    return code


if __name__ == "__main__":
    sys.exit(main())
