# OPC Importer — AI Project Handoff

- Handoff version: 1.0
- Date: 2026-10-18 (Asia/Tehran)
- Authoritative application baseline: `OPC_Importer_v1.26_2026-10-18.py`
- Project state: closing this development session. The user responded positively after the v1.26 filename-language fix, then requested this handoff. Treat v1.26 as the accepted baseline; this is not a claim that every Windows/Visio scenario has been tested.
- User language: Persian. Respond in Persian, with practical explanations and downloadable complete scripts.

## Instructions to the next AI

Read the actual v1.26 source before modifying it. This handoff describes intent and implementation but is not a substitute for source. Preserve all existing features unless the user explicitly changes them. Do not recreate the app from scratch or switch back to VBA/add-ins.

Deliver each revision as a complete Python file with an English filename, incremented version, and date: `OPC_Importer_v1.27_YYYY-MM-DD.py`. Update the window title and run instructions too. Keep older versions as separate artifacts. The user prefers action, concise communication, and no repeated confirmation for already-requested changes.

Do not claim real Visio rendering was verified using Python mocks. The development runtime was Linux; desktop Visio automation must be verified on Windows. In particular, do not repeatedly guess at text-direction fixes without inspecting the generated VSDX.

## Purpose and architecture

A standalone Windows Python/Tkinter application converts Excel operation process chart (OPC) tables into editable Visio documents and/or PDFs. It also exports existing Visio files to PDF. It uses `win32com.client.DispatchEx('Visio.InvisibleApp')` to create an isolated background Visio process. No VBA macro or UserForm is needed.

The project originally began as a VBA macro, then migrated to Python. A Visio add-in was considered and explicitly abandoned. Automatic updating/synchronization of an existing diagram from Excel is NOT implemented.

Dependencies:

- Windows and installed, activated desktop Microsoft Visio with PDF export support.
- Dubai font installed on Windows for generated diagrams.
- Python packages: `openpyxl`, `pywin32`, and `pypdf` for PDF merging.
- Tkinter is used for the UI.
- Microsoft Excel itself is not required.
- Even PDF-only conversion still requires desktop Visio.

The user had a broken `py` launcher. Prefer `python` commands.

## Running and packaging

From the directory containing the script:

```bat
python -m pip install openpyxl pywin32 pypdf
python OPC_Importer_v1.26_2026-10-18.py
```

For a Windows EXE, build on Windows:

```bat
python -m pip install --upgrade pyinstaller openpyxl pywin32 pypdf
python -m PyInstaller --clean --noconfirm --onefile --windowed --collect-all openpyxl --collect-all pypdf --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --name OPC_Importer_v1.26_2026-10-18 OPC_Importer_v1.26_2026-10-18.py
```

Distribute the executable from `dist`. Python and bundled Python packages should not need separate end-user installation. Visio and Dubai are external prerequisites and are not bundled. Rebuild the EXE after each script update; using an old EXE will retain old behavior. No Windows EXE was built or independently verified in the AI runtime.

## Excel contract

Supported inputs: `.xlsx` and `.xlsm`; no `.xls` support. Each selected worksheet becomes one Visio page. Each workbook produces its own Visio document when that output is enabled.

The first eight columns must appear in this order, with headers on the first row (comparison is case-insensitive):

| Column | Meaning / rules |
|---|---|
| ID | Required unique step ID within a sheet; matching is case-insensitive. |
| Activity | Required activity/shape type. |
| InsideText | Text inside the symbol. |
| SideText | Description beside the symbol. |
| Lane | Positive integer; blank means 1. Used lane numbers must be consecutive starting at 1. |
| LaneTitle | Lane heading; conflicting nonempty titles for one lane cause an error. |
| NextID | Explicit target ID, or automatically the next nonempty Excel data row's ID. |
| NoID | Optional explicit alternative target, only for Control/decision. Never inferred. |

Primary activity names:

| English | Persian | Shape |
|---|---|---|
| Operation | عملیات | Circle |
| QC | بازرسی | Square |
| Warehouse | انبار | Downward triangle |
| Control | کنترل | Diamond |

Additional aliases are in `KINDS`. The source normalizes case, whitespace and common Arabic/Persian letter variants.

Important semantics:

- Blank `NextID` follows Excel row order in BOTH single-lane and multi-lane charts, including across lanes. It does not mean the next numeric ID or next row in the same lane.
- Explicit `NextID` always takes precedence.
- The final row's blank `NextID` remains unconnected.
- Entirely empty rows are skipped.
- With current shorthand there is no explicit sentinel for “stop here despite later rows.” Do not claim blank NextID creates an intermediate terminal node.
- Merged input cells should be avoided; Activity must be present on every step row.
- Formula cells use saved cached values (`data_only=True`). The application does not calculate Excel formulas.
- Unknown target IDs, missing/duplicate IDs, unknown/empty Activity, invalid lanes, forbidden NoID, self-connections, and conflicting forward-order cycles are validated.
- Current preflight records the first raised validation error per worksheet, not an exhaustive list of every bad cell.

## Diagram layout and formatting

- A4 portrait, fixed page size, autosize disabled, margin 10 mm.
- Chart anchored at the top left and scaled down to fit available space.
- Every lane begins at the same top position. Within each lane, steps follow Excel row order.
- Cross-lane connectors NEVER push the destination lane downward.
- Forward-link validation still runs before lane-local visual levels are assigned. Keep validation and visual layout conceptually separate.
- Dynamic connectors, right-angle routing. These are real Visio connectors, not drawn line/arrow substitutes.
- Page names are sequential numbers; no page-number label at the top.
- Footer spans usable page width, framed and divided into three areas: process name (sheet title), document name (workbook filename stem), and a two-row symbol guide.
- Guide first row: square/inspection and triangle/warehouse; second row: circle/operation and diamond/control.
- Document name is derived from the INPUT Excel filename, without extension, not the selected output filename.
- Generated text uses Dubai for BOTH `Char.Font` and `Char.ComplexScriptFont`, and corresponding sizes. Font IDs are resolved through the current document's Fonts collection.
- Text is generally centered. LTR is not the same setting as left alignment.

## Critical text-direction history — do not regress

Most text uses automatic/context direction: first actual strong letter determines RTL or LTR. Initial digits, spaces, punctuation, and emoji are ignored. Persian starts RTL; English starts LTR. Text with no strong letter falls back to RTL. The manual RTL/LTR selector was removed by explicit user request.

Document-name text is explicitly LTR regardless of general automatic direction. The final fix adds English character language for ASCII LTR text before and after insertion (`Char.LangID = 1033`, en-US), across existing character-formatting rows.

Observed defect: `TST02-P` visually appeared like `TST-02P`; the uploaded `ANT02-CAD-A1000.vsdx` displayed the hyphen in the wrong position although caret traversal suggested correct logical order.

Actual inspection of the user's VSDX established:

- Page 1, shape ID 87 contained the exact logical string `ANT02-CAD-A1000` (plus a trailing paragraph newline).
- The shape inherited text style 3 (“Normal”).
- Normal style had Dubai for Latin and complex-script fonts and paragraph Flags = 0 (LTR).
- The name shape's character row explicitly had `LangID = fa-IR`.
- Thus the stored string was not wrongly copied. The source uses direct text assignment, not the clipboard.

v1.26 explicitly sets English character language on ASCII LTR text. User said “good” and requested project closure after that delivery. Do not overstate isolated causal proof: actual Windows rendering was not available to the AI.

Prior fixes retained in the baseline:

- Set paragraph direction BEFORE assigning text, then apply formatting afterwards.
- Document default style uses LTR; individual Persian text resolves to RTL.
- Apply `Para.Flags` with `FormulaForceU` and cover subsequent paragraph rows.
- A compatibility helper clears `Char.RTLText` where that cell is exposed. Do not assume it exists in every Visio version; the code checks availability.
- Remove Unicode bidi controls from the displayed filename only, preserving visible characters. The actual filesystem filename is unchanged.

Rejected approach: adding LRE/LRM/PDF or other hidden direction characters around the filename. They appeared as unwanted characters on the user's system. DO NOT add them again. Do not reverse text, move hyphens manually, or rasterize the filename as an unrequested workaround.

Existing-Visio PDF export intentionally preserves source formatting. It does NOT apply the importer’s font/direction fixes. For a changed importer to affect a diagram, generate it again from Excel.

## Main UI and preflight

- Multi-file Excel selection; workbook groups with sheet checkboxes.
- Global/per-workbook select-all and clear controls, and workbook removal.
- Sheets initially unselected.
- All sheets are checked when a workbook is loaded; errors appear under affected sheets.
- Header counts: total files, ready to build, files with any detected error (independent of checkbox selection), and files with no selected sheet.
- Counts overlap: a workbook can contain an invalid unselected sheet and still be ready to build from valid selected sheets. Error count is workbook-level, not error-cell count.
- “Recheck files” refreshes results after Excel edits, preserving selection by sheet name.
- Create rechecks before proceeding. Invalid selected sheets block conversion until fixed or deselected. Invalid unselected sheets do not block valid selected sheets.
- Preflight does not guarantee success with Visio installation, COM, file permissions or PDF export. Validation currently runs synchronously on the UI thread; large workbooks may pause the UI.
- Mouse-wheel scrolling is routed from nested labels/checkboxes to the list canvas only when the pointer is in the list. Windows wheel and Button-4/5 handlers are present.

## Output controls and overwrite behavior

Independent controls: save Visio, save individual PDFs, and save merged PDF. At least one is required.

- PDF-only: disable Visio output.
- Merged-only: disable Visio and individual PDFs, enable merged PDF.
- Merged PDF follows input-list order; each workbook contributes selected sheet pages in worksheet-list order.
- PDFs for merging are staged temporarily; unwanted individual PDFs are not published.
- PDF export uses `ExportAsFixedFormat(1, path, 1, 0)` — PDF, print intent, all foreground pages.
- When any input fails, the merged output is not produced; selected successful individual outputs may remain.
- Existing output names trigger custom Replace / Cancel, per output. Cancel skips that output, not all jobs. All outputs canceled => no conversion.
- No `_OPC` suffix and no automatic `_2` suffix for ordinary batch outputs in the current version.
- Inputs with identical stems in a batch are rejected to prevent ambiguous targets; user must process them separately.
- Output publication uses a temporary file in the target directory then `os.replace`. Originals are not overwritten until publication.
- Replacement choices are made before background generation; there is no second race-condition check for files created by another process afterwards.

## Existing Visio → PDF exporter

A separate button opens a dedicated dialog:

- Accepts one or more `.vsdx`, `.vsdm`, `.vsd` files.
- Independent individual-PDF and merged-PDF options; individual enabled by default.
- Does not read Excel or generate diagrams.
- Opens each input in an isolated background Visio instance with OpenEx flags 2|64|128|256: read-only, hidden, macros disabled, no workspace restoration.
- Exports all foreground pages without changing layout, fonts, language or text.
- Source documents are never saved. Closes the document and dedicated Visio process in finally blocks.
- Supports Replace / Cancel and merged-only output through the same output pipeline.
- Existing source fonts/layout/page sizes are preserved; A4/Dubai rules apply to newly generated diagrams only.

## Code map

| Function/class | Responsibility |
|---|---|
| `read_steps` | Parse, normalize, infer blank links, validate, then assign lane-local visual levels. |
| `apply_dubai`, `font` | Latin/complex-script font and size, alignment, text direction. |
| `detect_text_direction`, `set_paragraph_direction`, `apply_text_direction` | Context direction and paragraph handling. |
| `set_document_direction` | Default LTR style and Dubai font. |
| `set_ltr_character_runs` | Optional compatibility clearing of RTL character formatting. |
| `set_english_text_language` | v1.26 English language correction for ASCII LTR text. |
| `textbox` | Shape creation, direction/language before text insertion and after formatting. |
| `document_display_name` | Filename stem and removal of bidi controls only. |
| `symbol` | Circle/square/polygon geometry. |
| `legend`, `draw_page` | Footer, page layout, shapes, dynamic connectors. |
| `export_visio` | Excel-to-Visio COM work; optional save and PDF export. |
| `inspect_workbook` | Per-sheet preflight result collection. |
| `plan_batch_outputs` | Names and duplicate-stem rejection. |
| `merge_pdf_files`, `publish_file` | PDF merge and output publication. |
| `process_outputs` | Current shared pipeline, staging, skipped outputs, merging and results. |
| `export_existing_visio` | Read-only existing-document PDF export adapter. |
| `ImportWindow` | Tkinter UI, validation display, dialogs, thread/queue handling. |
| `run_batch` | Older helper retained in source; current UI uses `process_outputs`. |

## Earlier bugs worth remembering

- Avoid SAFEARRAY/polyline marshaling issues: triangle/diamond geometry is built with scalar ShapeSheet operations.
- Do not use `RowCount` on Style objects. Existing helpers use named cell-existence checks.
- Paragraph section index is 4, not 6 (Scratch). Do not restore that early mistake.
- No invented `PagefitScale` cell; page sizing uses actual page/print settings.
- Font IDs are document-specific; do not hardcode IDs or reuse them across documents.
- Keep operations on the isolated Visio instance; never quit the user's existing session.
- COM initialization/cleanup occurs inside the worker context.

## Verification and remaining limits

During development, syntax, parser/validation behavior, automatic links, independent lane levels, output combinations, temporary-file cleanup, overwrite skipping, mock COM calls, and filename transformation logic were checked. PDF merging was tested with real generated PDFs, including page order and failure preservation. Source VSDX XML was inspected for the naming bug.

Desktop Visio rendering, Windows Tk UI behavior, and a packaged EXE were not directly exercised by the AI. The user performed iterative Windows testing. There is no maintained comprehensive automated test suite; temporary development checks are not part of the deliverable.

Suggested checks when making the NEXT requested change (do not reopen completed work unnecessarily):

1. Generate a workbook named `ANT02-CAD-A1000.xlsx` and verify the title visually and via VSDX XML.
2. Confirm Dubai for Persian and English, with no added direction characters.
3. Multi-lane input: blank NextID crosses lanes while every lane starts at top.
4. Merged-only output produces only the requested final PDF.
5. Cancel preserves an existing output; Replace publishes the new one.
6. Existing-Visio export leaves source bytes unchanged.

No further change is requested at handoff. Preserve this baseline and resume only from the user's next task.
