---
name: ccdc-digest-character-report
description: Scan a folder of CCDC digest Excel workbooks for characters rejected by the CCDC ETL character rule, then create a read-only Excel report with exact cell locations, original values, Unicode details, and values with invalid characters removed. Use when a user provides a digest folder by pasted path or drag-and-drop and does not want to run CCDC ETL manually.
---

# CCDC Digest Character Report

Turn one digest folder into a plain-language summary and a filterable Excel report. Do not ask the user to run commands or understand the ETL.

## Input

Accept either:

- a folder path pasted into the message; or
- a folder reference created by dragging the folder into the message.

Resolve the supplied path or folder reference to an absolute local directory before scanning. If the dropped item cannot be resolved as a readable directory, ask the user to paste the folder path or drag the folder again. If no folder was supplied, ask only for the folder and explain that they may drag it into the message or paste its path. Do not ask them to run a terminal command.

## Run the scanner

1. Resolve the folder to an absolute path and confirm it is a readable directory.
2. Call `load_workspace_dependencies` to obtain the bundled Node executable and `node_modules` path.
3. Create a temporary runtime directory, copy `scripts/scan-digests.mjs` into it, and create a `node_modules` symlink there pointing to the bundled dependency directory. This lets the script load `@oai/artifact-tool` without installing anything.
4. Run the copied script with `<bundled-node> <temporary-script> --input <absolute-folder-path>`. Add `--output <absolute-xlsx-path>` only when the user requested a specific destination. Otherwise the scanner writes to `<input-folder>/ccdc-character-reports/special-character-report-<timestamp>.xlsx`.
5. Parse the final JSON line printed by the scanner. Treat exit code `1` as a successful scan with findings, `0` as a clean scan, and `2` as a fatal failure.
6. Re-open the generated report with the spreadsheet tooling. Check the Summary values, inspect representative rows in `Special Characters`, and render the populated report sheets for visual verification.
7. Remove only the temporary runtime directory created for this run. Never remove source or report files.

If the default report location is not writable, choose a writable output path and tell the user where the report was saved. Ask for approval if access to the supplied folder requires it.

## Fixed character policy

The bundled scanner reproduces the CCDC ETL character rule without starting or importing CCDC ETL:

- Convert a defined value to a string and apply JavaScript `trim()`.
- An empty result passes.
- CR (`U+000D`), LF (`U+000A`), and printable ASCII (`U+0020` through `U+007E`) pass.
- Every other Unicode code point is reported.

The report's `Value Without Special Characters` is the trimmed value with every rejected code point deleted. It is advisory. Do not describe it as an ETL transformation.

## Scope and safety

- Scan regular, non-hidden `.xlsx` files directly inside the supplied folder, in filename order.
- Scan the Data Resource, Dataset Information, and mapped Digest fields checked by CCDC ETL.
- Match digest sheets by exact dataset name first, then use the ETL positional fallback.
- Keep scanning after a workbook parse or mapping failure and record the error.
- Never overwrite, repair, rename, move, or delete source workbooks.
- Do not create corrected workbooks unless the user explicitly requests them.

## Response

Use non-technical language. State:

- how many workbooks were found and scanned;
- how many affected cells and invalid-character occurrences were found;
- whether any workbook could not be scanned; and
- a clickable link to the generated `.xlsx` report.

If there are no findings, still return the report. Do not expose internal runtime setup unless troubleshooting is needed.
