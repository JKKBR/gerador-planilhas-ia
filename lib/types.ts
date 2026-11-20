export type CellValue = string | number | boolean | null;
export type ColumnType = "text" | "number" | "currency" | "date" | "status";
export type InputMode = "manual" | "calculated";

export type Column = {
  key: string;
  label: string;
  type: ColumnType;
  width?: number;
  inputMode?: InputMode;
  required?: boolean;
};

export type FillFormula = {
  mode?: "fill";
  key: string;
  expression: string;
  fromRow?: number;
  toRow?: number;
};

export type SingleFormula = {
  mode: "single" | "dynamic";
  targetCell: string;
  expression: string;
  clearRange?: string;
};

export type Formula = FillFormula | SingleFormula;

export type Sheet = {
  id: string;
  name: string;
  purpose?: string;
  preparedRows?: number;
  columns: Column[];
  rows: Record<string, CellValue>[];
  formulas?: Formula[];
};

export type WorkbookModel = {
  name: string;
  category: "Estoque" | "Finanças" | "Orçamento" | "Personalizada";
  description: string;
  sheets: Sheet[];
};

export type ValidationIssue = {
  severity: "error" | "warning";
  code:
    | "DUPLICATE_SHEET"
    | "DUPLICATE_COLUMN"
    | "MISSING_FORMULA_COLUMN"
    | "INVALID_FORMULA_RANGE"
    | "INVALID_TARGET_CELL"
    | "MISSING_REFERENCED_SHEET"
    | "UNSUPPORTED_FORMULA"
    | "FORMULA_NOT_FILLED"
    | "FIXED_VALUE_IN_CALCULATED_COLUMN"
    | "SPILL_RANGE_BLOCKED";
  sheet: string;
  cell?: string;
  message: string;
};

export type ExportReport = {
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
  formulaCount: number;
  downloaded: boolean;
};
