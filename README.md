# DMS report compiler

Copies the DMS template and inserts the chapters of a SCADE report into it.

## Use

1. Open `run.py` and edit the SETTINGS at the top (template, Excel sheet, output folder).
2. Press **F5** (VS Code / Visual Studio), or double-click `Build reports.bat`.

SCADE reports can be `.docx`, `.rtf` or `.doc`; `.rtf` and `.doc` files are
converted with Word first, so they need Word installed and `USE_WORD = True`.

The finished reports appear in the output folder. Missing Python libraries are
installed automatically on the first run.

## Preparing the template

- Put a paragraph containing `{{SCADE_CHAPTERS}}` where the SCADE chapters go.
- Write manual chapters before/after it. Use `{{name}}` for words that change
  per report; each name is a column in the Excel sheet.
- Wrap optional chapters between a `{{#if name}}` paragraph and a `{{/if}}`
  paragraph; they are kept when the column says yes/true/x.

## Files

- `run.py` - settings + start button
- `dms_compiler.py` - the compiler
- `make_examples.py` - regenerates the synthetic reports in `examples/`
- `input/` - put real reports here; this folder is ignored by git
