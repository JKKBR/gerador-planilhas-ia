import type {
  Formula,
  WorkbookModel,
} from "./types";

import {
  formatWorkbookValidationErrors,
  validateWorkbook,
} from "./workbook-schema";

export type QualityStatus =
  | "ready"
  | "warning"
  | "blocked";

export type WorkbookQualityReport = {
  status: QualityStatus;
  title: string;
  errors: string[];
  warnings: string[];
  sheetCount: number;
  columnCount: number;
  exampleRowCount: number;
  preparedRowCount: number;
  formulaCount: number;
  dynamicFormulaCount: number;
};

function formulaMode(
  formula: Formula
): "fill" | "single" | "dynamic" {
  return formula.mode ?? "fill";
}

export function inspectWorkbookQuality(
  input: unknown
): WorkbookQualityReport {
  const validation = validateWorkbook(input);

  if (!validation.success) {
    return {
      status: "blocked",
      title: "Download bloqueado",
      errors: formatWorkbookValidationErrors(
        validation.error
      ),
      warnings: [],
      sheetCount: 0,
      columnCount: 0,
      exampleRowCount: 0,
      preparedRowCount: 0,
      formulaCount: 0,
      dynamicFormulaCount: 0,
    };
  }

  const workbook =
    validation.data as WorkbookModel;

  const errors: string[] = [];
  const warnings: string[] = [];

  let columnCount = 0;
  let exampleRowCount = 0;
  let preparedRowCount = 0;
  let formulaCount = 0;
  let dynamicFormulaCount = 0;

  for (const sheet of workbook.sheets) {
    columnCount += sheet.columns.length;
    exampleRowCount += sheet.rows.length;
    preparedRowCount +=
      sheet.preparedRows ?? sheet.rows.length;

    const formulas = sheet.formulas ?? [];
    formulaCount += formulas.length;

    const formulaKeys = new Set(
      formulas
        .filter(
          (formula) =>
            formulaMode(formula) === "fill" &&
            "key" in formula
        )
        .map((formula) =>
          "key" in formula ? formula.key : ""
        )
        .filter(Boolean)
    );

    for (const formula of formulas) {
      if (formulaMode(formula) === "dynamic") {
        dynamicFormulaCount += 1;
      }
    }

    const calculatedColumns =
      sheet.columns.filter(
        (column) =>
          column.inputMode === "calculated"
      );

    for (const column of calculatedColumns) {
      const hasFillFormula =
        formulaKeys.has(column.key);

      const hasPointFormula =
        formulas.some(
          (formula) =>
            formulaMode(formula) !== "fill"
        );

      if (!hasFillFormula && !hasPointFormula) {
        errors.push(
          `A aba "${sheet.name}" possui a coluna calculada "${column.label}" sem regra de fórmula.`
        );
      }
    }

    if (
      (sheet.preparedRows ?? 0) === 0 &&
      sheet.rows.length === 0
    ) {
      warnings.push(
        `A aba "${sheet.name}" não possui linhas preparadas nem exemplos.`
      );
    }

    if (sheet.columns.length >= 18) {
      warnings.push(
        `A aba "${sheet.name}" possui ${sheet.columns.length} colunas e pode exigir rolagem horizontal.`
      );
    }
  }

  if (dynamicFormulaCount > 0) {
    warnings.push(
      `${dynamicFormulaCount} fórmula(s) dinâmica(s) possuem instruções de ajuste manual no arquivo.`
    );
  }

  if (formulaCount === 0) {
    warnings.push(
      "A planilha não possui regras de fórmula."
    );
  }

  const status: QualityStatus =
    errors.length > 0
      ? "blocked"
      : warnings.length > 0
        ? "warning"
        : "ready";

  const title =
    status === "blocked"
      ? "Download bloqueado"
      : status === "warning"
        ? "Pronto com avisos"
        : "Pronto para baixar";

  return {
    status,
    title,
    errors,
    warnings,
    sheetCount: workbook.sheets.length,
    columnCount,
    exampleRowCount,
    preparedRowCount,
    formulaCount,
    dynamicFormulaCount,
  };
}
