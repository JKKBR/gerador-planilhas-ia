import type {
  Formula,
  Sheet,
  WorkbookModel,
} from "./types";

export type SheetSummary = {
  id: string;
  name: string;
  purpose?: string;
  columnCount: number;
  exampleRowCount: number;
  preparedRows: number;
  manualColumns: string[];
  calculatedColumns: string[];
  fillFormulaCount: number;
  singleFormulaCount: number;
  dynamicFormulaCount: number;
};

export type WorkbookSummary = {
  name: string;
  category: WorkbookModel["category"];
  description: string;
  sheetCount: number;
  totalColumns: number;
  totalExampleRows: number;
  totalPreparedRows: number;
  totalFormulas: number;
  sheets: SheetSummary[];
  warnings: string[];
};

function formulaMode(
  formula: Formula
): "fill" | "single" | "dynamic" {
  return formula.mode ?? "fill";
}

function inferredPreparedRows(sheet: Sheet): number {
  if (
    typeof sheet.preparedRows === "number" &&
    Number.isFinite(sheet.preparedRows)
  ) {
    return Math.max(0, sheet.preparedRows);
  }

  return sheet.rows.length;
}

function summarizeSheet(sheet: Sheet): SheetSummary {
  const formulas = sheet.formulas ?? [];

  const manualColumns = sheet.columns
    .filter(
      (column) =>
        (column.inputMode ?? "manual") === "manual"
    )
    .map((column) => column.label);

  const calculatedColumns = sheet.columns
    .filter(
      (column) =>
        column.inputMode === "calculated"
    )
    .map((column) => column.label);

  return {
    id: sheet.id,
    name: sheet.name,
    purpose: sheet.purpose,
    columnCount: sheet.columns.length,
    exampleRowCount: sheet.rows.length,
    preparedRows: inferredPreparedRows(sheet),
    manualColumns,
    calculatedColumns,
    fillFormulaCount: formulas.filter(
      (formula) => formulaMode(formula) === "fill"
    ).length,
    singleFormulaCount: formulas.filter(
      (formula) => formulaMode(formula) === "single"
    ).length,
    dynamicFormulaCount: formulas.filter(
      (formula) => formulaMode(formula) === "dynamic"
    ).length,
  };
}

export function buildWorkbookSummary(
  workbook: WorkbookModel
): WorkbookSummary {
  const sheets = workbook.sheets.map(summarizeSheet);

  const warnings: string[] = [];

  for (const sheet of sheets) {
    if (sheet.preparedRows === 0) {
      warnings.push(
        `A aba "${sheet.name}" não possui linhas preparadas.`
      );
    }

    if (
      sheet.calculatedColumns.length > 0 &&
      sheet.fillFormulaCount === 0 &&
      sheet.singleFormulaCount === 0 &&
      sheet.dynamicFormulaCount === 0
    ) {
      warnings.push(
        `A aba "${sheet.name}" possui colunas calculadas sem regras de fórmula.`
      );
    }

    if (sheet.dynamicFormulaCount > 0) {
      warnings.push(
        `A aba "${sheet.name}" possui fórmula dinâmica. O arquivo incluirá instruções para ajuste manual, se necessário.`
      );
    }
  }

  return {
    name: workbook.name,
    category: workbook.category,
    description: workbook.description,
    sheetCount: sheets.length,
    totalColumns: sheets.reduce(
      (total, sheet) =>
        total + sheet.columnCount,
      0
    ),
    totalExampleRows: sheets.reduce(
      (total, sheet) =>
        total + sheet.exampleRowCount,
      0
    ),
    totalPreparedRows: sheets.reduce(
      (total, sheet) =>
        total + sheet.preparedRows,
      0
    ),
    totalFormulas: sheets.reduce(
      (total, sheet) =>
        total +
        sheet.fillFormulaCount +
        sheet.singleFormulaCount +
        sheet.dynamicFormulaCount,
      0
    ),
    sheets,
    warnings,
  };
}
