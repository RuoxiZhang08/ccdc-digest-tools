# CCDC Digest Tools

This small Codex plugin scans a folder of CCDC digest Excel workbooks for characters rejected by the CCDC ETL character rule. It creates a read-only Excel report containing the file, sheet, cell, field, original value, Unicode details, and a proposed value with invalid characters removed.

The plugin is designed for non-technical users. They provide a folder path in Codex; they do not need the CCDC source repository or a local ETL installation.

## Install from GitHub

Replace `GITHUB-OWNER` with the repository owner's GitHub username:

```sh
codex plugin marketplace add GITHUB-OWNER/ccdc-digest-tools
codex plugin add ccdc-digest-tools@ccdc-digest-tools
```

Then ask Codex:

> Scan `/path/to/my/digest/folder` with `$ccdc-digest-character-report`.

By default, the report is saved inside the selected digest folder under `ccdc-character-reports/`. If that folder is not writable, Codex chooses a writable destination and tells the user where it saved the report.

## Safety

The scanner does not edit, rename, move, or delete source workbooks. The suggested cleaned values in the report are advisory and are not an ETL transformation.

## Repository structure

- `.agents/plugins/marketplace.json` — Codex marketplace catalog
- `plugins/ccdc-digest-tools/.codex-plugin/plugin.json` — plugin manifest
- `plugins/ccdc-digest-tools/skills/ccdc-digest-character-report/` — skill instructions and scanner
