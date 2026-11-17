export type CellValue = string | number | boolean | null;
export type Column = { key: string; label: string; type: "text"|"number"|"currency"|"date"|"status"; width?: number };
export type Formula = { key: string; expression: string };
export type Sheet = { id: string; name: string; purpose?: string; columns: Column[]; rows: Record<string,CellValue>[]; formulas?: Formula[] };
export type WorkbookModel = { name: string; category: "Estoque"|"Finanças"|"Orçamento"|"Personalizada"; description: string; sheets: Sheet[] };
