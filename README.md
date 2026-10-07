# eDoc Generator

A Windows desktop app for turning fixed-format calibration Excel reports into filled company approval Excel documents and merged PDFs.

**One report produces one approval workbook and one final PDF. Five reports produce five sets.** The final PDF contains the approval pages first, followed by the calibration certificate. You choose the Excel template, input folder, output folder and any number of cell mappings in the app. Save a mapping profile and reuse it for later batches.

## Requirements

- Windows 10 or 11, with desktop Microsoft Excel installed and activated.
- The actual company approval template in `.xlsx`, `.xlsm`, `.xlsb` or `.xls` format.
- Reports with the mapped worksheets and cells in consistent locations.
- For source-code setup/building: 64-bit Python 3.12 and an internet connection to install dependencies.
- For an already built executable: Python is **not** required; Microsoft Excel still is.

This is a document preparation tool. It does not grant company approval, sign a certificate, or upload company data to a server. It runs locally. Macro-dependent templates and workbooks with external links/data connections are not supported.

## Download the Windows executable

Open this repository's **Actions** tab → **Test and build Windows app** → a successful run → download **EDocGenerator-Windows** under Artifacts. Extract the ZIP and run `EDocGenerator.exe`.

Version tags such as `v1.0.0` also publish the executable and its SHA-256 checksum to GitHub Releases. The executable is unsigned. A successful build is not a substitute for testing your actual template in desktop Excel.

If a workflow is not present or has not run, use the source setup or local build below. GitHub's Windows runner builds the executable but does not have desktop Excel to verify PDF rendering.

## Run from source

1. Download this repository as a ZIP, then **extract it** into a local folder.
2. Install 64-bit Python 3.12 from [python.org](https://www.python.org/downloads/windows/), with the Python launcher enabled.
3. Double-click `setup_windows.bat` once.
4. Double-click `run_windows.bat` to open the app.

Use a local folder you can write to. Do not run scripts from inside an unextracted ZIP. If your company computer restricts installations, use its approved installation process.

## First-time setup

### 1. Files and folders

Choose:

- **Approval Excel template:** your existing company-format approval document.
- **Input report folder:** only the calibration reports you want processed.
- **Output folder:** a separate folder outside the input folder.
- **Certificate source:** export each report using Excel, attach an existing matching PDF, or create approval pages only.

For existing PDFs, `ACC-001.xlsx` matches exactly one `ACC-001.pdf` in the chosen certificate folder. The filename is used for matching; the app cannot independently prove that the PDF is for the correct instrument. Review your naming convention and generated PDFs. With recursive inputs, certificate subfolders must mirror input subfolders.

You may specify an **Approval PDF worksheet** or **Report PDF worksheet**. Leave either blank to export the whole workbook using Excel's saved print settings. Set the print area, orientation, margins, headers and scaling in Excel beforehand. Save those settings. The app preserves the template rather than redesigning its layout.

### 2. Cell mappings

Click **Inspect worksheets**, choose a sample report, then **Add mapping**. Select the exact report worksheet and cell, and the approval worksheet and destination cell. Repeat for as many fields as you need.

The following addresses are examples only; they are not your actual company mappings:

| Field | Report cell | Approval cell | Mode |
| --- | --- | --- | --- |
| Serial number | Report!B4 | Approval!F6 | text |
| Calibration date | Report!D5 | Approval!F7 | value |
| Sensitivity | Report!C12 | Approval!D14 | value |
| Measured value | Report!C13 | Approval!D15 | value |
| Uncertainty | Report!C14 | Approval!D16 | value |

- **Value** copies the underlying number, string or Excel-recognized date. The destination retains the template's formatting. A date destination should already use a date format. A source date stored as a plain unformatted serial number is treated as a number; use date formatting in the report or text mode.
- **Text** copies Excel's displayed text, including leading zeros, rounded decimals and units. A report showing `00012` stays `00012`. `####` display errors are rejected; widen the source column or switch to value mode.
- **Required** rejects this report if the source is blank. Optional blanks clear the destination's placeholder.
- **Allow replacing a formula** is off by default. Turn it on only for a destination formula you deliberately intend to replace.
- For merged cells, the top-left cell is used. Two destinations resolving to the same merged cell are rejected.

Click **Preview sample** to see extracted values and validate destination cells without changing files. Verify this preview against the original report, then **Save profile**. Profiles are readable JSON files. The last generation profile is restored automatically on the same computer. You can load another profile for a different report format.

### 3. Generate

Click **Generate documents**. The app processes all matching Excel reports, shows success/failure for each, and continues after an individual report fails. Excel lock files beginning `~$` and the selected template are excluded. Subfolder processing is optional.

The app copies source reports and the template before processing, opens reports read-only, and never saves the originals. Every batch has a new folder and every report has its own folder. Successful report outputs are published only after generation and merging finish.

Example output:

```text
Output/
  edoc_20261007_063000_ab12cd34/
    batch.json
    001_ACC-001_a12b34c56d/
      approval.xlsx
      approval.pdf
      certificate.pdf
      final.pdf
      audit.json
    002_ACC-002_b23c45d67e/
      ...
```

`approval.xlsx` keeps the original template's extension when it is `.xlsm`, `.xls` or `.xlsb`. The app always keeps the filled workbook, final PDF and audit. Keeping component PDFs is optional and enabled by default.

`batch.json` records failures and reports not processed. `audit.json` records the source/template hashes, mapped values, final hash and page count. These files can contain sensitive instrument identifiers and local paths. Keep output folders under your normal company document controls; `.gitignore` helps exclude workbooks/PDFs but does not replace reviewing what you commit.

**Stop after current report** lets the current report finish and then stops the batch. Closing the app while processing offers the same graceful stop. There is no forced termination of Excel. If Excel is blocked by an Office dialog, the current operation can still wait; resolve the dialog in Excel. Workbooks requiring passwords are not supported.

## Calculation and PDF fidelity

Source reports are read in manual calculation mode; external links are not refreshed. Save and calculate your reports in Excel before processing so formula results are current. The filled approval workbook is fully recalculated with Excel's dependency scheduler before saving/exporting, and saved in automatic calculation mode. Complex formulas, custom fonts, printer-specific layouts, VBA functions and cross-sheet dependencies require a real-template check; this app cannot certify their correctness.

The Excel export respects saved print areas. If the PDF has blank pages or clipped text, adjust the workbook's print setup and preview it in Excel. A generated PDF can still need visual checking for long identifiers or values that do not fit the template cells.

Merging PDFs does not preserve a certificate's digital-signature validity. Keep the original signed PDF separately if its signature matters. The merged PDF is the combined document copy.

## Build the `.exe` locally

On Windows, double-click `build_windows.bat`. It creates a virtual environment, installs pinned dependencies, runs the tests and uses PyInstaller. The result is `dist/EDocGenerator.exe`. Only distribute an executable built from a trusted commit. Building on Linux does not produce a Windows executable.

## Publish the source to the provided empty repository

The included `publish_github.ps1` targets `Raman0925/Edoc-generator`. If the source has not yet been uploaded, install [Git](https://git-scm.com/downloads/win) and [GitHub CLI](https://cli.github.com/), run the source setup above, and configure your Git commit name/email. From PowerShell in the extracted folder:

```powershell
.\publish_github.ps1
```

The script uses GitHub's browser login if needed, runs tests, initializes a local repository and pushes to the empty remote. It refuses to overwrite an existing local Git history or remote branch. No credentials are embedded in the project. After pushing, the included Actions workflow builds the Windows executable. Company workbooks and generated PDFs are not part of the upload.

## Validation

```powershell
python -m unittest discover -s tests -v
```

The portable automated tests exercise configuration validation, original-file preservation, five-report batches, failure isolation, cancellation, certificate matching, PDF order, PDF encryption rejection, dates, numeric precision and literal strings. The batch tests use an injected fake Excel adapter and real PDF operations; they do not claim to verify Office rendering.

To run the included real Excel smoke test on an Office-equipped Windows PC:

```powershell
$env:EDOC_RUN_EXCEL_TESTS = '1'
python -m unittest discover -s tests -v
```

That test creates temporary workbooks, checks a leading-zero serial and high-precision number, exports/merges PDFs, and checks the original report. It is skipped on ordinary GitHub runners. Before production use, also compare a generated PDF from your actual approval template and a known report against the manually prepared document.

## Learn the code

Start with [LEARNING.md](LEARNING.md). `edoc/config.py` defines mappings and validates them. `edoc/excel.py` controls a private Excel instance. `edoc/engine.py` loops over reports and produces audited outputs. `edoc/gui.py` lets you select paths and edit mappings, while keeping Excel work off the UI thread. `launcher.py` starts the app.

## Troubleshooting

| Problem | Action |
| --- | --- |
| Could not start Excel | Confirm desktop Excel is installed; open it once and complete activation/setup. |
| Worksheet not found | Inspect worksheets; use the exact name in the mapping. |
| Required value empty | Check the mapped source cell and the report format. |
| Formula destination blocked | Choose the intended input cell, or explicitly permit formula replacement for that mapping. |
| Protected approval sheet | Use an authorized unlocked template copy. The app does not bypass protection. |
| Missing certificate PDF | Match the Excel report's base filename and, if enabled, its relative subfolder. |
| Extra pages / clipping | Set print areas, scaling and margins in Excel and save them. |
| Different formula result | Recalculate and save reports in Excel first; check template formula dependencies. |
| Output path error | Use a short local path such as `C:/Calibration/Output` and check write permissions. |

Company sample workbooks are intentionally not included. Configure the mappings and verify the print layout with your own files.
