#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const REPORT_FOLDER = "ccdc-character-reports";
const ASCII_RULE = /^[\r\n\x20-\x7E]+$/;
const MAX_EXCEL_TEXT = 32767;

const RESOURCE_FIELDS = [
  [10, 1, "Initial Submission Date", "initial_submission_date"],
  [10, 3, "Resource ID", "id"],
  [10, 4, "Resource Name", "resource_name"],
  [10, 5, "Resource URI", "resource_uri"],
  [15, 1, "Resource Description", "description"],
  [15, 2, "Resource Type", "resource_type"],
  [15, 3, "Data Content Type", "data_content_type"],
  [15, 4, "Resource POC", "poc"],
  [15, 5, "Resource POC Email", "poc_email"],
  [15, 6, "Analytics", "analytics"],
  [15, 7, "Visualization", "visualization"],
  [15, 8, "Pediatric Specific", "pediatric_specific"],
  [15, 9, "API", "api"],
  [18, 1, "Data Update Date", "data_update_date"],
  [18, 5, "Resource Contact URL", "resource_contact_url"],
];

const DATASET_FIELDS = [
  [3, "Dataset Abbreviated Name", "dataset_name"],
  [4, "Dataset Full Name", "dataset_full_name"],
  [5, "Dataset Description", "description"],
  [6, "Primary Dataset Scope", "dataset_scope"],
  [7, "Dataset POC", "poc"],
  [8, "POC Email", "poc_email"],
  [9, "Reference", "published_in"],
];

const DIGEST_FIELDS = [
  [2, "Parent Digest ID", "parent_digest_level"],
  [3, "Digest ID", "digest_level"],
  [4, "Core Element", "core_element"],
  [5, "Data Element", "data_element"],
  [6, "Data Element Value", "element_value"],
  [7, "Statistic Type", "statistic_type"],
  [8, "Statistic Value", "statistic_value"],
];

function parseArgs(argv) {
  const result = { input: "", output: "" };
  for (let index = 0; index < argv.length; index += 1) {
    const item = argv[index];
    if (item !== "--input" && item !== "--output") throw new Error(`Unknown argument: ${item}`);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) throw new Error(`Missing value for ${item}`);
    result[item.slice(2)] = value;
    index += 1;
  }
  if (!result.input) throw new Error("A digest folder is required with --input");
  return result;
}

function timestamp(date = new Date()) {
  return date.toISOString().replace(/[:.]/g, "-");
}

function readableTimestamp(date) {
  return `${date.toLocaleString("en-US", {
    timeZone: "UTC",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  })} UTC`;
}

function normalizedValue(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function invalidOccurrences(value) {
  const normalized = normalizedValue(value);
  if (!normalized || ASCII_RULE.test(normalized)) return { normalized, occurrences: [] };
  const occurrences = [];
  let position = 0;
  for (const character of Array.from(normalized)) {
    position += 1;
    if (character === "\r" || character === "\n" || /^[\x20-\x7E]$/.test(character)) continue;
    occurrences.push({
      character,
      codePoint: `U+${character.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`,
      position,
    });
  }
  return { normalized, occurrences };
}

function displayCharacter(character) {
  const names = { "\t": "TAB", "\u00A0": "NO-BREAK SPACE", "\u200B": "ZERO WIDTH SPACE" };
  return names[character] || character;
}

function cleanValue(normalized) {
  return Array.from(normalized)
    .filter((character) => character === "\r" || character === "\n" || /^[\x20-\x7E]$/.test(character))
    .join("");
}

function columnName(oneBasedColumn) {
  let current = oneBasedColumn;
  let result = "";
  while (current > 0) {
    current -= 1;
    result = String.fromCharCode(65 + (current % 26)) + result;
    current = Math.floor(current / 26);
  }
  return result;
}

function valueAt(data, oneBasedRow, oneBasedColumn) {
  return data[oneBasedRow - 1]?.[oneBasedColumn - 1];
}

function rowHasValue(row) {
  return Array.isArray(row) && row.some((value) => value !== null && value !== undefined && value !== "");
}

function safeText(value) {
  if (value === undefined || value === null) return "";
  let result = String(value);
  if (result.length > MAX_EXCEL_TEXT) result = `${result.slice(0, MAX_EXCEL_TEXT - 14)}… [truncated]`;
  if (/^[=+\-@]/.test(result)) result = `'${result}`;
  return result;
}

function worksheetData(sheet) {
  const range = sheet?.getUsedRange();
  return range ? range.values : [];
}

function makeFinding({ filename, sheetName, sheetIndex, datasetName = "", row, column, field, etlField, value }) {
  const { normalized, occurrences } = invalidOccurrences(value);
  if (occurrences.length === 0) return null;
  const distinct = [];
  const seen = new Set();
  for (const occurrence of occurrences) {
    const key = `${occurrence.character}:${occurrence.codePoint}`;
    if (!seen.has(key)) {
      seen.add(key);
      distinct.push(`${displayCharacter(occurrence.character)} (${occurrence.codePoint})`);
    }
  }
  return {
    filename,
    sheetName,
    sheetIndex,
    datasetName,
    row,
    column: columnName(column),
    cell: `${columnName(column)}${row}`,
    field,
    etlField,
    originalValue: safeText(value),
    specialCharacters: distinct.join(", "),
    unicodeCodePoints: occurrences.map((item) => item.codePoint).join(", "),
    characterPositions: occurrences.map((item) => item.position).join(", "),
    cleanedValue: safeText(cleanValue(normalized)),
    occurrenceCount: occurrences.length,
  };
}

function selectDigestSheets(sheets, datasetNames) {
  const digestSheets = sheets.slice(2);
  const selected = new Array(datasetNames.length);
  const claimed = new Set();
  datasetNames.forEach((datasetName, index) => {
    if (!datasetName) return;
    const named = digestSheets.find((sheet) => sheet.name === datasetName && !claimed.has(sheet));
    if (named) {
      selected[index] = named;
      claimed.add(named);
    }
  });
  datasetNames.forEach((datasetName, index) => {
    if (!datasetName || selected[index]) return;
    let available = digestSheets[index];
    if (claimed.has(available)) {
      const unclaimed = digestSheets.filter((sheet) => !claimed.has(sheet));
      if (unclaimed.length !== 1) throw new Error(`Digest sheet for ${datasetName} is ambiguous`);
      [available] = unclaimed;
    }
    if (!available) throw new Error(`Digest sheet for ${datasetName} is missing`);
    selected[index] = available;
    claimed.add(available);
  });
  return selected;
}

function scanWorkbook(workbook, filename) {
  const sheets = workbook.worksheets.items;
  const findings = [];
  const errors = [];
  const add = (details) => {
    const finding = makeFinding({ filename, ...details });
    if (finding) findings.push(finding);
  };

  const resource = sheets[0];
  if (!resource) {
    errors.push("Missing Data Resource sheet");
  } else {
    const data = worksheetData(resource);
    for (const [row, column, field, etlField] of RESOURCE_FIELDS) {
      add({ sheetName: resource.name, sheetIndex: 0, row, column, field, etlField, value: valueAt(data, row, column) });
    }
  }

  const dataset = sheets[1];
  if (!dataset) {
    errors.push("Missing Dataset Information sheet");
    return { findings, errors, sheetCount: sheets.length };
  }

  const datasetData = worksheetData(dataset);
  for (let row = 2; row <= datasetData.length; row += 1) {
    const datasetName = valueAt(datasetData, row, 3) || `Dataset_${row - 1}`;
    for (const [column, field, etlField] of DATASET_FIELDS) {
      add({ sheetName: dataset.name, sheetIndex: 1, datasetName, row, column, field, etlField, value: valueAt(datasetData, row, column) });
    }
  }

  const datasetNames = datasetData.slice(1).map((row, index) => {
    if (!rowHasValue(row)) return undefined;
    return row?.[2] || `Dataset_${index + 1}`;
  });

  let selected;
  try {
    selected = selectDigestSheets(sheets, datasetNames);
  } catch (error) {
    errors.push(error.message);
    return { findings, errors, sheetCount: sheets.length };
  }

  datasetNames.forEach((datasetName, index) => {
    if (!datasetName) return;
    const sheet = selected[index];
    if (!sheet) {
      errors.push(`Digest sheet for ${datasetName} is missing`);
      return;
    }
    const data = worksheetData(sheet);
    for (let row = 2; row <= data.length; row += 1) {
      for (const [column, field, etlField] of DIGEST_FIELDS) {
        add({ sheetName: sheet.name, sheetIndex: sheets.indexOf(sheet), datasetName, row, column, field, etlField, value: valueAt(data, row, column) });
      }
    }
  });

  return { findings, errors, sheetCount: sheets.length };
}

function setHeader(range) {
  range.format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF", name: "Arial", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
  };
}

function formatBody(range) {
  range.format = {
    font: { color: "#222222", name: "Arial", size: 10 },
    verticalAlignment: "top",
  };
}

function writeTableSheet(workbook, name, headers, rows, widths, tableName) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const matrix = [headers, ...rows];
  sheet.getRange("A1").write(matrix);
  const lastColumn = columnName(headers.length);
  const lastRow = matrix.length;
  setHeader(sheet.getRange(`A1:${lastColumn}1`));
  if (lastRow > 1) {
    formatBody(sheet.getRange(`A2:${lastColumn}${lastRow}`));
    const table = sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  sheet.freezePanes.freezeRows(1);
  widths.forEach((width, index) => {
    sheet.getRange(`${columnName(index + 1)}:${columnName(index + 1)}`).format.columnWidth = width;
  });
  if (lastRow > 1) sheet.getRange(`A2:${lastColumn}${lastRow}`).format.wrapText = true;
  return sheet;
}

async function writeReport({ inputPath, outputPath, runAt, files, findings, errors }) {
  const workbook = Workbook.create();
  const summary = workbook.worksheets.add("Summary");
  summary.showGridLines = false;
  const affectedFiles = new Set(findings.map((item) => item.filename)).size;
  const occurrences = findings.reduce((sum, item) => sum + item.occurrenceCount, 0);
  const scanned = files.filter((item) => item.status !== "ERROR").length;
  const summaryRows = [
    ["CCDC Digest Character Report", ""],
    ["Input folder", safeText(inputPath)],
    ["Scan time", readableTimestamp(runAt)],
    ["Workbooks found", files.length],
    ["Workbooks scanned", scanned],
    ["Workbooks with findings", affectedFiles],
    ["Workbooks with errors", files.filter((item) => item.status === "ERROR").length],
    ["Affected cells", findings.length],
    ["Invalid-character occurrences", occurrences],
    ["Character rule", "After trim, allow CR, LF, and printable ASCII U+0020–U+007E."],
    ["Clean-value rule", "Delete rejected code points from the trimmed value. Source workbooks are unchanged."],
  ];
  summary.getRange("A1").write(summaryRows);
  summary.getRange("A1:B1").format = { font: { bold: true, color: "#1F4E78", name: "Arial", size: 14 } };
  summary.getRange("A2:A11").format = { font: { bold: true, color: "#222222", name: "Arial", size: 10 }, fill: "#D9EAF7" };
  summary.getRange("B2:B11").format = { font: { color: "#222222", name: "Arial", size: 10 }, wrapText: true };
  summary.getRange("A:A").format.columnWidth = 30;
  summary.getRange("B:B").format.columnWidth = 74;
  summary.getRange("A1:B11").format.verticalAlignment = "top";

  const findingRows = findings.map((item) => [
    safeText(item.filename), safeText(item.sheetName), item.sheetIndex, safeText(item.datasetName), item.row, item.column, item.cell,
    item.field, item.etlField, item.originalValue, item.specialCharacters, item.unicodeCodePoints,
    item.characterPositions, item.cleanedValue, item.occurrenceCount,
  ]);
  writeTableSheet(workbook, "Special Characters", [
    "Filename", "Sheet", "Sheet Index", "Dataset", "Row", "Column", "Cell", "Field", "ETL Field",
    "Original Value", "Special Characters", "Unicode Code Points", "Character Positions",
    "Value Without Special Characters", "Occurrence Count",
  ], findingRows, [28, 24, 11, 20, 9, 9, 11, 24, 22, 42, 28, 25, 20, 42, 12], "SpecialCharactersTable");

  const fileRows = files.map((item) => [
    safeText(item.filename), item.status, item.sheetCount ?? "", item.affectedCells,
    item.occurrenceCount, item.errorMessage || "",
  ]);
  writeTableSheet(workbook, "Files", [
    "Filename", "Status", "Sheet Count", "Affected Cells", "Invalid-character Occurrences", "Error Message",
  ], fileRows, [30, 13, 13, 14, 25, 58], "FilesTable");

  const errorRows = errors.map((item) => [safeText(item.filename), item.stage, safeText(item.message)]);
  writeTableSheet(workbook, "Errors", ["Filename", "Stage", "Message"], errorRows, [30, 20, 70], "ErrorsTable");

  writeTableSheet(workbook, "Rules", ["Rule", "Definition", "Source"], [[
    "CCDC ETL character rule",
    "Convert the value to text, trim it, allow blank, CR, LF, and printable ASCII U+0020–U+007E; report every other code point.",
    "CCDC-ETL/common/utils.js containsSpecialCharacters",
  ], [
    "Value Without Special Characters",
    "Delete rejected code points from the trimmed value. This report-only suggestion is not an ETL transformation.",
    "CCDC Digest Character Report",
  ]], [34, 90, 52], "RulesTable");

  workbook.recalculate();
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  return { affectedFiles, occurrences, scanned };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = path.resolve(args.input);
  const stat = await fs.stat(inputPath).catch((error) => {
    throw new Error(`Cannot access input folder ${inputPath}: ${error.message}`);
  });
  if (!stat.isDirectory()) throw new Error(`Input path is not a folder: ${inputPath}`);

  const outputPath = path.resolve(args.output || path.join(
    inputPath,
    REPORT_FOLDER,
    `special-character-report-${timestamp()}.xlsx`,
  ));
  if (path.extname(outputPath).toLowerCase() !== ".xlsx") throw new Error("Output path must end in .xlsx");

  const entries = await fs.readdir(inputPath, { withFileTypes: true });
  const filenames = entries
    .filter((entry) => entry.isFile()
      && !entry.name.startsWith(".")
      && !entry.name.startsWith("~$")
      && /\.xlsx$/i.test(entry.name))
    .map((entry) => entry.name)
    .sort((left, right) => left.localeCompare(right));

  const sourcePaths = new Set(await Promise.all(
    filenames.map((filename) => fs.realpath(path.join(inputPath, filename))),
  ));
  let outputIdentity;
  try {
    outputIdentity = await fs.realpath(outputPath);
  } catch {
    outputIdentity = path.resolve(outputPath);
  }
  if (sourcePaths.has(outputIdentity)) {
    throw new Error("Output path must not overwrite an input workbook");
  }

  const files = [];
  const findings = [];
  const errors = [];
  for (const filename of filenames) {
    const filePath = path.join(inputPath, filename);
    try {
      const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
      const result = scanWorkbook(workbook, filename);
      findings.push(...result.findings);
      result.errors.forEach((message) => errors.push({ filename, stage: "Workbook mapping", message }));
      const occurrenceCount = result.findings.reduce((sum, item) => sum + item.occurrenceCount, 0);
      files.push({
        filename,
        status: result.errors.length ? "ERROR" : result.findings.length ? "FINDINGS" : "CLEAN",
        sheetCount: result.sheetCount,
        affectedCells: result.findings.length,
        occurrenceCount,
        errorMessage: result.errors.join("; "),
      });
    } catch (error) {
      errors.push({ filename, stage: "Workbook read", message: error.message });
      files.push({ filename, status: "ERROR", sheetCount: "", affectedCells: 0, occurrenceCount: 0, errorMessage: error.message });
    }
  }

  findings.sort((left, right) => left.filename.localeCompare(right.filename)
    || left.sheetIndex - right.sheetIndex
    || left.row - right.row
    || left.column.localeCompare(right.column));

  const runAt = new Date();
  const totals = await writeReport({ inputPath, outputPath, runAt, files, findings, errors });
  const summary = {
    ok: true,
    inputPath,
    outputPath,
    workbooksFound: files.length,
    workbooksScanned: totals.scanned,
    workbooksWithFindings: totals.affectedFiles,
    workbooksWithErrors: files.filter((item) => item.status === "ERROR").length,
    affectedCells: findings.length,
    invalidCharacterOccurrences: totals.occurrences,
  };
  console.log(JSON.stringify(summary));
  process.exitCode = findings.length > 0 || errors.length > 0 ? 1 : 0;
}

main().catch((error) => {
  console.error(JSON.stringify({ ok: false, error: error.message }));
  process.exitCode = 2;
});
