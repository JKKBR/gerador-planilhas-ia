import type { ExportReport, Formula, ValidationIssue, WorkbookModel } from "./types";

const ALLOWED_FUNCTIONS = new Set([
  "IF", "IFERROR", "AND", "OR", "NOT", "SUM", "SUMIF", "SUMIFS",
  "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "AVERAGE", "AVERAGEIF",
  "AVERAGEIFS", "MIN", "MAX", "ROUND", "ROUNDUP", "ROUNDDOWN",
  "ABS", "FILTER", "SORT", "UNIQUE", "XLOOKUP", "INDEX", "MATCH",
  "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "TODAY",
  "DATE", "YEAR", "MONTH", "DAY",
]);

function normalizeText(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function trimHumanText(value: unknown) {
  return typeof value === "string" && !value.trimStart().startsWith("=")
    ? value.trim().replace(/\s+/g, " ")
    : value;
}

function safeSheetName(name: string) {
  return name.replace(/[\\/*?:[\]]/g, " ").trim().slice(0, 31) || "Planilha";
}

function uniqueSheetName(name: string, used: Set<string>) {
  const base = safeSheetName(name);
  let candidate = base;
  let index = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` ${index}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function legacyPreparedRowCount(model: WorkbookModel, sheetName: string, sheetId: string) {
  const category = normalizeText(model.category);
  const identity = normalizeText(`${sheetName} ${sheetId}`);
  if (identity.includes("moviment")) return 1000;
  if (category === "financas" && (identity.includes("lancamento") || identity.includes("transacao"))) return 1000;
  if (category === "orcamento" && (identity.includes("item") || identity.includes("produto orcado"))) return 500;
  if (["cadastro", "produto", "peca", "fornecedor", "cliente"].some((term) => identity.includes(term))) return 300;
  return 0;
}

function requestedRows(model: WorkbookModel, sheet: WorkbookModel["sheets"][number]) {
  return Math.max(0, Math.min(5000, sheet.preparedRows ?? legacyPreparedRowCount(model, sheet.name, sheet.id)));
}

function cleanFormula(expression: string, rowNumber?: number) {
  const withoutEquals = expression.trim().replace(/^=/, "").replace(/;/g, ",");
  return rowNumber ? withoutEquals.replaceAll("{row}", String(rowNumber)) : withoutEquals;
}

function toExcelValue(value: unknown) {
  const normalized = trimHumanText(value);
  if (typeof normalized === "string" && normalized.trim().startsWith("=")) {
    return { formula: cleanFormula(normalized) };
  }
  return normalized ?? null;
}

function formulaMode(formula: Formula) {
  return formula.mode ?? "fill";
}

function isFormulaObject(value: unknown) {
  return typeof value === "object" && value !== null && "formula" in value;
}

function referencedSheets(expression: string) {
  return [...expression.matchAll(/'([^']+)'!/g)].map((match) => match[1]);
}

function functionsIn(expression: string) {
  return [...expression.matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].map((match) => match[1].toUpperCase());
}

function validateModel(model: WorkbookModel): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const sheetNames = new Set<string>();

  for (const sheet of model.sheets) {
    const sheetNameKey = safeSheetName(sheet.name).toLowerCase();
    if (sheetNames.has(sheetNameKey)) {
      issues.push({ severity: "error", code: "DUPLICATE_SHEET", sheet: sheet.name, message: "Há duas abas com o mesmo nome após a normalização." });
    }
    sheetNames.add(sheetNameKey);

    const columnKeys = new Set<string>();
    for (const column of sheet.columns) {
      if (columnKeys.has(column.key)) {
        issues.push({ severity: "error", code: "DUPLICATE_COLUMN", sheet: sheet.name, message: `A chave de coluna "${column.key}" está duplicada.` });
      }
      columnKeys.add(column.key);
    }

    for (const formula of sheet.formulas ?? []) {
      const mode = formulaMode(formula);
      if (mode === "fill") {
        const fill = formula as Extract<Formula, { mode?: "fill" }>;
        if (!columnKeys.has(fill.key)) {
          issues.push({ severity: "error", code: "MISSING_FORMULA_COLUMN", sheet: sheet.name, message: `A fórmula usa a coluna inexistente "${fill.key}".` });
        }
        const from = fill.fromRow ?? 2;
        const to = fill.toRow ?? Math.max(2, requestedRows(model, sheet) + 1);
        if (from < 2 || to < from || to > 5001) {
          issues.push({ severity: "error", code: "INVALID_FORMULA_RANGE", sheet: sheet.name, message: `Intervalo de fórmula inválido: ${from} até ${to}.` });
        }
      } else {
        const single = formula as Extract<Formula, { mode: "single" | "dynamic" }>;
        if (!/^[A-Z]{1,3}[1-9][0-9]*$/i.test(single.targetCell)) {
          issues.push({ severity: "error", code: "INVALID_TARGET_CELL", sheet: sheet.name, cell: single.targetCell, message: "Célula de destino inválida." });
        }
      }

      for (const fn of functionsIn(formula.expression)) {
        if (!ALLOWED_FUNCTIONS.has(fn)) {
          issues.push({ severity: "warning", code: "UNSUPPORTED_FORMULA", sheet: sheet.name, message: `A função ${fn} não está na lista validada e deve ser revisada.` });
        }
      }
      for (const referenced of referencedSheets(formula.expression)) {
        if (![...sheetNames, ...model.sheets.map((item) => safeSheetName(item.name).toLowerCase())].includes(safeSheetName(referenced).toLowerCase())) {
          issues.push({ severity: "error", code: "MISSING_REFERENCED_SHEET", sheet: sheet.name, message: `A fórmula referencia a aba inexistente "${referenced}".` });
        }
      }
    }
  }
  return issues;
}

function rangeCells(range: string) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return { startCol: match[1], startRow: Number(match[2]), endCol: match[3], endRow: Number(match[4]) };
}

export async function exportXlsx(model: WorkbookModel): Promise<ExportReport> {
  const initialIssues = validateModel(model);
  const initialErrors = initialIssues.filter((issue) => issue.severity === "error");
  const warnings = initialIssues.filter((issue) => issue.severity === "warning");
  if (initialErrors.length) return { errors: initialErrors, warnings, formulaCount: 0, downloaded: false };

  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  const usedSheetNames = new Set<string>();
  const runtimeIssues: ValidationIssue[] = [];
  let formulaCount = 0;

  workbook.creator = "Gerador de Planilhas IA";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;
  workbook.calcProperties.calcMode = "auto";

  const sheetNameMap = new Map<string, string>();
  for (const sheet of model.sheets) {
    sheetNameMap.set(sheet.name, uniqueSheetName(sheet.name, usedSheetNames));
  }

  for (const sheet of model.sheets) {
    const worksheet = workbook.addWorksheet(sheetNameMap.get(sheet.name)!, { views: [{ state: "frozen", ySplit: 1 }] });
    worksheet.columns = sheet.columns.map((column) => ({
      header: column.label.trim().replace(/\s+/g, " "),
      key: column.key,
      width: column.width || Math.max(12, Math.min(30, column.label.length + 5)),
    }));

    const header = worksheet.getRow(1);
    header.height = 24;
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1D4ED8" } };
    header.alignment = { vertical: "middle", horizontal: "center", wrapText: true };

    for (const sourceRow of sheet.rows) {
      worksheet.addRow(sheet.columns.map((column) => toExcelValue(sourceRow[column.key])));
    }

    const fillFormulas = (sheet.formulas ?? []).filter((formula) => formulaMode(formula) === "fill") as Extract<Formula, { mode?: "fill" }>[];
    const formulaKeys = new Set(fillFormulas.map((formula) => formula.key));
    const totalDataRows = Math.max(sheet.rows.length, requestedRows(model, sheet));
    let lastRowNumber = Math.max(1, totalDataRows + 1);
    for (const formula of fillFormulas) lastRowNumber = Math.max(lastRowNumber, formula.toRow ?? 0);

    for (let rowNumber = 2; rowNumber <= lastRowNumber; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      row.height = 20;
      const isPreparedBlankRow = rowNumber > sheet.rows.length + 1;
      sheet.columns.forEach((column, index) => {
        const cell = row.getCell(index + 1);
        cell.alignment = { vertical: "middle", horizontal: ["text", "status"].includes(column.type) ? "left" : "right" };
        cell.border = { bottom: { style: "hair", color: { argb: "FFE2E8F0" } } };
        if (isPreparedBlankRow && (column.inputMode ?? "manual") === "manual") cell.font = { color: { argb: "FF0000FF" } };
      });
      if (rowNumber % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    }

    for (const formula of fillFormulas) {
      const columnIndex = sheet.columns.findIndex((column) => column.key === formula.key) + 1;
      if (columnIndex < 1) continue;
      const fromRow = formula.fromRow ?? 2;
      const toRow = formula.toRow ?? Math.max(2, requestedRows(model, sheet) + 1);
      for (let rowNumber = fromRow; rowNumber <= toRow; rowNumber++) {
        const cell = worksheet.getRow(rowNumber).getCell(columnIndex);
        cell.value = { formula: cleanFormula(formula.expression, rowNumber) };
        cell.font = { color: { argb: "FF000000" } };
        formulaCount++;
      }
    }

    const pointFormulas = (sheet.formulas ?? []).filter((formula) => formulaMode(formula) !== "fill") as Extract<Formula, { mode: "single" | "dynamic" }>[];
    for (const formula of pointFormulas) {
      if (formula.mode === "dynamic" && formula.clearRange) {
        const parsed = rangeCells(formula.clearRange);
        if (parsed) {
          const target = formula.targetCell.toUpperCase();
          for (let row = parsed.startRow; row <= parsed.endRow; row++) {
            for (let col = worksheet.getColumn(parsed.startCol).number; col <= worksheet.getColumn(parsed.endCol).number; col++) {
              const cell = worksheet.getCell(row, col);
              if (cell.address !== target) cell.value = null;
            }
          }
        }
      }
      worksheet.getCell(formula.targetCell).value = { formula: cleanFormula(formula.expression) };
      formulaCount++;
    }

    for (const column of sheet.columns) {
      if ((column.inputMode ?? "manual") !== "calculated") continue;
      const columnIndex = sheet.columns.findIndex((item) => item.key === column.key) + 1;
      const hasFillRule = fillFormulas.some((formula) => formula.key === column.key);
      if (!hasFillRule) {
        for (let rowNumber = 2; rowNumber <= sheet.rows.length + 1; rowNumber++) {
          const value = worksheet.getRow(rowNumber).getCell(columnIndex).value;
          if (value !== null && value !== undefined && value !== "" && !isFormulaObject(value)) {
            runtimeIssues.push({ severity: "error", code: "FIXED_VALUE_IN_CALCULATED_COLUMN", sheet: sheet.name, cell: worksheet.getRow(rowNumber).getCell(columnIndex).address, message: `A coluna calculada "${column.label}" contém valor fixo.` });
          }
        }
      }
    }

    sheet.columns.forEach((column, index) => {
      const excelColumn = worksheet.getColumn(index + 1);
      if (column.type === "currency") excelColumn.numFmt = 'R$ #,##0.00';
      if (column.type === "date") excelColumn.numFmt = "dd/mm/yyyy";
      if (column.type === "number") excelColumn.numFmt = "0.00";
    });

    if (sheet.columns.length && lastRowNumber > 1) {
      worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: lastRowNumber, column: sheet.columns.length } };
    }

    const statusColumnIndex = sheet.columns.findIndex((column) => column.type === "status") + 1;
    if (statusColumnIndex > 0 && lastRowNumber > 1) {
      const letter = worksheet.getColumn(statusColumnIndex).letter;
      worksheet.addConditionalFormatting({
        ref: `${letter}2:${letter}${lastRowNumber}`,
        rules: [
          { type: "containsText", operator: "containsText", text: "ESGOTADO", priority: 1, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFECACA" } }, font: { color: { argb: "FF991B1B" }, bold: true } } } as any,
          { type: "containsText", operator: "containsText", text: "DEFEITO", priority: 2, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFECACA" } }, font: { color: { argb: "FF991B1B" }, bold: true } } } as any,
          { type: "containsText", operator: "containsText", text: "REPOR", priority: 3, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } }, font: { color: { argb: "FF92400E" }, bold: true } } } as any,
          { type: "containsText", operator: "containsText", text: "REPARANDO", priority: 4, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } }, font: { color: { argb: "FF92400E" }, bold: true } } } as any,
          { type: "containsText", operator: "containsText", text: "OK", priority: 5, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } }, font: { color: { argb: "FF166534" }, bold: true } } } as any,
          { type: "containsText", operator: "containsText", text: "PRONTO", priority: 6, style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } }, font: { color: { argb: "FF166534" }, bold: true } } } as any,
        ],
      });
    }
  }

  const errors = runtimeIssues.filter((issue) => issue.severity === "error");
  const allWarnings = [...warnings, ...runtimeIssues.filter((issue) => issue.severity === "warning")];
  if (errors.length) return { errors, warnings: allWarnings, formulaCount, downloaded: false };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${model.name.replace(/[^a-zA-Z0-9À-ÿ_-]+/g, "_")}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
  return { errors: [], warnings: allWarnings, formulaCount, downloaded: true };
}
