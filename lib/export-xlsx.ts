import type {
  ExportReport,
  Formula,
  ValidationIssue,
  WorkbookModel,
} from "./types";

const ALLOWED_FUNCTIONS = new Set([
  "IF", "IFERROR", "AND", "OR", "NOT", "SUM", "SUMIF", "SUMIFS",
  "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "AVERAGE", "AVERAGEIF",
  "AVERAGEIFS", "MIN", "MAX", "ROUND", "ROUNDUP", "ROUNDDOWN",
  "ABS", "FILTER", "SORT", "UNIQUE", "XLOOKUP", "INDEX", "MATCH",
  "AGGREGATE", "ROW", "ROWS", "COLUMN", "COLUMNS", "LEFT", "RIGHT",
  "MID", "LEN", "TRIM", "UPPER", "LOWER", "TODAY", "DATE", "YEAR",
  "MONTH", "DAY",
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

function localizeInstructionFormulaPtBr(expression: string) {
  let localized = cleanFormula(expression);

  const translations: Array<[RegExp, string]> = [
    [/\bFILTER\s*\(/gi, "FILTRO("],
    [/\bSORT\s*\(/gi, "CLASSIFICAR("],
    [/\bUNIQUE\s*\(/gi, "ÚNICO("],
    [/\bXLOOKUP\s*\(/gi, "PROCX("],

    // Coloque funções maiores antes das menores.
    [/\bIFERROR\s*\(/gi, "SEERRO("],
    [/\bCOUNTIFS\s*\(/gi, "CONT.SES("],
    [/\bCOUNTIF\s*\(/gi, "CONT.SE("],
    [/\bCOUNTA\s*\(/gi, "CONT.VALORES("],
    [/\bCOUNT\s*\(/gi, "CONT.NÚM("],
    [/\bSUMIFS\s*\(/gi, "SOMASES("],
    [/\bSUMIF\s*\(/gi, "SOMASE("],
    [/\bSUM\s*\(/gi, "SOMA("],
    [/\bAVERAGEIFS\s*\(/gi, "MÉDIASES("],
    [/\bAVERAGEIF\s*\(/gi, "MÉDIASE("],
    [/\bAVERAGE\s*\(/gi, "MÉDIA("],

    [/\bINDEX\s*\(/gi, "ÍNDICE("],
    [/\bMATCH\s*\(/gi, "CORRESP("],
    [/\bAGGREGATE\s*\(/gi, "AGREGAR("],
    [/\bCOLUMNS\s*\(/gi, "COLS("],
    [/\bCOLUMN\s*\(/gi, "COL("],
    [/\bROWS\s*\(/gi, "LINS("],
    [/\bROW\s*\(/gi, "LIN("],

    [/\bROUNDUP\s*\(/gi, "ARREDONDAR.PARA.CIMA("],
    [/\bROUNDDOWN\s*\(/gi, "ARREDONDAR.PARA.BAIXO("],
    [/\bROUND\s*\(/gi, "ARRED("],

    [/\bIF\s*\(/gi, "SE("],
    [/\bAND\s*\(/gi, "E("],
    [/\bOR\s*\(/gi, "OU("],
    [/\bNOT\s*\(/gi, "NÃO("],
    [/\bTODAY\s*\(/gi, "HOJE("],
    [/\bYEAR\s*\(/gi, "ANO("],
    [/\bMONTH\s*\(/gi, "MÊS("],
    [/\bDAY\s*\(/gi, "DIA("],
  ];

  for (const [pattern, replacement] of translations) {
    localized = localized.replace(pattern, replacement);
  }

  // Fórmulas digitadas manualmente no Excel em português usam ;.
  return localized.replace(/,/g, ";");
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

function columnToNumber(column: string) {
  return column.toUpperCase().split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function numberToColumn(number: number) {
  let result = "";
  let value = number;
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function parseRange(range: string) {
  const match = range.trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return {
    startCol: match[1].toUpperCase(),
    startRow: Number(match[2]),
    endCol: match[3].toUpperCase(),
    endRow: Number(match[4]),
  };
}

type ParsedFilter = {
  sourceSheet: string;
  sourceStartCol: string;
  sourceStartRow: number;
  sourceEndCol: string;
  sourceEndRow: number;
  criteriaSheet: string;
  criteriaCol: string;
  criteriaStartRow: number;
  criteriaEndRow: number;
  criteriaValue: string;
};

function parseFilterFormula(expression: string): ParsedFilter | null {
  const formula = cleanFormula(expression);
  const pattern = /^(?:_xlfn\._xlws\.)?FILTER\(\s*'([^']+)'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\s*,\s*'([^']+)'!\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)\s*=\s*"([^"]*)"\s*,\s*""\s*\)$/i;
  const match = formula.match(pattern);
  if (!match) return null;
  return {
    sourceSheet: match[1],
    sourceStartCol: match[2].toUpperCase(),
    sourceStartRow: Number(match[3]),
    sourceEndCol: match[4].toUpperCase(),
    sourceEndRow: Number(match[5]),
    criteriaSheet: match[6],
    criteriaCol: match[7].toUpperCase(),
    criteriaStartRow: Number(match[8]),
    criteriaEndRow: Number(match[9]),
    criteriaValue: match[10],
  };
}

function escapeFormulaText(value: string) {
  return value.replace(/"/g, '""');
}

function buildLegacyFilterFormula(
  parsed: ParsedFilter,
  targetStartCol: string,
  targetStartRow: number,
  targetCol: string,
  targetRow: number
) {
  const occurrence = `ROWS($${targetStartCol}$${targetStartRow}:${targetStartCol}${targetRow})`;
  const sourceColumnIndex = `COLUMNS($${targetStartCol}:${targetCol})`;
  const source = `'${parsed.sourceSheet}'!$${parsed.sourceStartCol}$${parsed.sourceStartRow}:$${parsed.sourceEndCol}$${parsed.sourceEndRow}`;
  const criteria = `'${parsed.criteriaSheet}'!$${parsed.criteriaCol}$${parsed.criteriaStartRow}:$${parsed.criteriaCol}$${parsed.criteriaEndRow}`;
  const firstCriteriaCell = `'${parsed.criteriaSheet}'!$${parsed.criteriaCol}$${parsed.criteriaStartRow}`;
  const criterion = escapeFormulaText(parsed.criteriaValue);
  return `IFERROR(INDEX(${source},AGGREGATE(15,6,(ROW(${criteria})-ROW(${firstCriteriaCell})+1)/(${criteria}="${criterion}"),${occurrence}),${sourceColumnIndex}),"")`;
}

function validateModel(model: WorkbookModel): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const normalizedNames = model.sheets.map((sheet) => safeSheetName(sheet.name).toLowerCase());

  model.sheets.forEach((sheet, sheetIndex) => {
    if (normalizedNames.indexOf(normalizedNames[sheetIndex]) !== sheetIndex) {
      issues.push({ severity: "error", code: "DUPLICATE_SHEET", sheet: sheet.name, message: "Há duas abas com o mesmo nome após a normalização." });
    }

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
        const point = formula as Extract<Formula, { mode: "single" | "dynamic" }>;
        if (!/^[A-Z]{1,3}[1-9][0-9]*$/i.test(point.targetCell)) {
          issues.push({ severity: "error", code: "INVALID_TARGET_CELL", sheet: sheet.name, cell: point.targetCell, message: "Célula de destino inválida." });
        }
      }

      for (const fn of functionsIn(formula.expression)) {
        if (!ALLOWED_FUNCTIONS.has(fn)) {
          issues.push({ severity: "warning", code: "UNSUPPORTED_FORMULA", sheet: sheet.name, message: `A função ${fn} não está na lista validada e deve ser revisada.` });
        }
      }
      for (const referenced of referencedSheets(formula.expression)) {
        if (!normalizedNames.includes(safeSheetName(referenced).toLowerCase())) {
          issues.push({ severity: "error", code: "MISSING_REFERENCED_SHEET", sheet: sheet.name, message: `A fórmula referencia a aba inexistente "${referenced}".` });
        }
      }
    }
  });
  return issues;
}

type FormulaHelpRow = {
  targetSheet: string;
  target: string;
  type: string;
  purpose: string;
  originalFormula: string;
  automaticFormula: string;
  steps: string;
  expected: string;
};

function addInstructionsSheet(workbook: any, rows: FormulaHelpRow[]) {
  if (!rows.length) return;

  const instructionsSheetName = "Instruções e Fórmulas";
  const existingInstructionsSheet = workbook.getWorksheet(instructionsSheetName);

  // A IA pode ter criado essa aba. Remove a versão anterior para que o
  // exportador gere uma única aba padronizada, sem conflito de nomes.
  if (existingInstructionsSheet) {
    workbook.removeWorksheet(existingInstructionsSheet.id);
  }

  const worksheet = workbook.addWorksheet(instructionsSheetName, {
    views: [{ state: "frozen", ySplit: 5 }],
  });

  worksheet.mergeCells("A1:H1");
  worksheet.getCell("A1").value = "Instruções e fórmulas da planilha";
  worksheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1D4ED8" } };
  worksheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  worksheet.getRow(1).height = 28;

  worksheet.mergeCells("A2:H2");
  worksheet.getCell("A2").value = "A planilha usa fórmulas tradicionais para listar todos os registros filtrados. As fórmulas FILTER originais são mantidas abaixo apenas como referência para inserção manual em versões compatíveis do Excel.";
  worksheet.getCell("A2").alignment = { wrapText: true, vertical: "middle" };
  worksheet.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
  worksheet.getRow(2).height = 44;

  worksheet.mergeCells("A3:H3");
  worksheet.getCell("A3").value = "Se optar pela fórmula dinâmica: limpe todo o intervalo de saída, clique apenas na célula inicial, cole a fórmula pela barra de fórmulas e pressione Enter. Não arraste a fórmula.";
  worksheet.getCell("A3").alignment = { wrapText: true, vertical: "middle" };
  worksheet.getCell("A3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
  worksheet.getRow(3).height = 40;

  const headers = ["Aba de destino", "Célula/intervalo", "Tipo", "Finalidade", "Fórmula original para copiar", "Fórmula automática usada", "Passo a passo", "Resultado esperado"];
  worksheet.getRow(5).values = headers;
  const header = worksheet.getRow(5);
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
  header.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  header.height = 34;

  for (const item of rows) {
    const row = worksheet.addRow([
      item.targetSheet,
      item.target,
      item.type,
      item.purpose,
      `'=${localizeInstructionFormulaPtBr(item.originalFormula)}`,
      `'=${item.automaticFormula}`,
      item.steps,
      item.expected,
    ]);
    row.alignment = { vertical: "top", wrapText: true };
    row.getCell(5).font = { color: { argb: "FF008000" } };
    row.getCell(6).font = { color: { argb: "FF000000" } };
    row.height = 88;
  }

  worksheet.columns = [
    { width: 20 }, { width: 18 }, { width: 20 }, { width: 34 },
    { width: 70 }, { width: 90 }, { width: 58 }, { width: 34 },
  ];
  worksheet.autoFilter = { from: "A5", to: `H${Math.max(6, worksheet.rowCount)}` };
}

export async function exportXlsx(model: WorkbookModel): Promise<ExportReport> {
  const initialIssues = validateModel(model);
  const errors = initialIssues.filter((issue) => issue.severity === "error");
  const warnings = initialIssues.filter((issue) => issue.severity === "warning");
  if (errors.length) return { errors, warnings, formulaCount: 0, downloaded: false };

  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  const usedSheetNames = new Set<string>();
  const runtimeIssues: ValidationIssue[] = [];
  const helpRows: FormulaHelpRow[] = [];
  let formulaCount = 0;

  workbook.creator = "Gerador de Planilhas IA";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const sheetNameMap = new Map<string, string>();
  for (const sheet of model.sheets) {
    sheetNameMap.set(sheet.name, uniqueSheetName(sheet.name, usedSheetNames));
  }

  for (const sheet of model.sheets) {
    const worksheet = workbook.addWorksheet(sheetNameMap.get(sheet.name)!, {
      views: [{ state: "frozen", ySplit: 1 }],
    });

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
      const preparedBlank = rowNumber > sheet.rows.length + 1;
      sheet.columns.forEach((column, index) => {
        const cell = row.getCell(index + 1);
        cell.alignment = { vertical: "middle", horizontal: ["text", "status"].includes(column.type) ? "left" : "right" };
        cell.border = { bottom: { style: "hair", color: { argb: "FFE2E8F0" } } };
        if (preparedBlank && (column.inputMode ?? "manual") === "manual" && !formulaKeys.has(column.key)) {
          cell.font = { color: { argb: "FF0000FF" } };
        }
      });
      if (rowNumber % 2 === 0) {
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      }
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
      helpRows.push({
        targetSheet: sheet.name,
        target: `${worksheet.getColumn(columnIndex).letter}${fromRow}:${worksheet.getColumn(columnIndex).letter}${toRow}`,
        type: "Fórmula por linha",
        purpose: `Calcular automaticamente a coluna ${sheet.columns[columnIndex - 1].label}.`,
        originalFormula: cleanFormula(formula.expression, fromRow),
        automaticFormula: cleanFormula(formula.expression, fromRow),
        steps: `1. Abra a aba ${sheet.name}. 2. Clique em ${worksheet.getColumn(columnIndex).letter}${fromRow}. 3. Cole a fórmula. 4. Copie até a linha ${toRow}.`,
        expected: `A coluna ${sheet.columns[columnIndex - 1].label} deve ser calculada automaticamente.`,
      });
    }

    const pointFormulas = (sheet.formulas ?? []).filter((formula) => formulaMode(formula) !== "fill") as Extract<Formula, { mode: "single" | "dynamic" }>[];
    for (const formula of pointFormulas) {
      if (formula.mode === "dynamic") {
        const parsedFilter = parseFilterFormula(formula.expression);
        const parsedOutput = formula.clearRange ? parseRange(formula.clearRange) : null;
        if (parsedFilter && parsedOutput) {
          const firstAutomatic = buildLegacyFilterFormula(
            parsedFilter,
            parsedOutput.startCol,
            parsedOutput.startRow,
            parsedOutput.startCol,
            parsedOutput.startRow
          );

          // O ExcelJS grava fórmulas, mas não as calcula. Para que a aba já
          // abra preenchida, salvamos também o resultado inicial em cache.
          const sourceModelSheet = model.sheets.find(
            (candidate) => candidate.name === parsedFilter.sourceSheet
          );
          const criteriaColumnIndex = columnToNumber(parsedFilter.criteriaCol) - 1;
          const sourceStartColumnIndex = columnToNumber(parsedFilter.sourceStartCol) - 1;
          const criteriaKey = sourceModelSheet?.columns[criteriaColumnIndex]?.key;
          const matchingRows = sourceModelSheet && criteriaKey
            ? sourceModelSheet.rows.filter(
                (sourceRow) =>
                  String(sourceRow[criteriaKey] ?? "").trim() ===
                  parsedFilter.criteriaValue.trim()
              )
            : [];

          for (let rowNumber = parsedOutput.startRow; rowNumber <= parsedOutput.endRow; rowNumber++) {
            const matchIndex = rowNumber - parsedOutput.startRow;
            for (let colNumber = columnToNumber(parsedOutput.startCol); colNumber <= columnToNumber(parsedOutput.endCol); colNumber++) {
              const targetCol = numberToColumn(colNumber);
              const sourceOffset = colNumber - columnToNumber(parsedOutput.startCol);
              const sourceKey = sourceModelSheet?.columns[
                sourceStartColumnIndex + sourceOffset
              ]?.key;
              const cachedResult = sourceKey
                ? matchingRows[matchIndex]?.[sourceKey] ?? ""
                : "";

              // Grava o resultado filtrado como valor estático. Isso evita que
              // o Excel recalcule uma fórmula incompatível e apague a listagem.
              // A fórmula original e a alternativa permanecem disponíveis na
              // aba "Instruções e Fórmulas" para aplicação manual.
              worksheet.getCell(`${targetCol}${rowNumber}`).value = cachedResult;
              if (cachedResult !== "") {
                worksheet.getCell(`${targetCol}${rowNumber}`).font = {
                  color: { argb: "FF008000" },
                };
              }
            }
          }
          helpRows.push({
            targetSheet: sheet.name,
            target: `${formula.targetCell} ou ${formula.clearRange}`,
            type: "Filtro automático com alternativa compatível",
            purpose: `Listar todos os registros em que o critério seja "${parsedFilter.criteriaValue}".`,
            originalFormula: cleanFormula(formula.expression),
            automaticFormula: firstAutomatic,
            steps: `Uso automático: nenhuma ação. Para testar FILTER manualmente: 1. Abra a aba ${sheet.name}. 2. Limpe ${formula.clearRange}. 3. Clique somente em ${formula.targetCell}. 4. Copie a fórmula original desta aba de instruções. 5. Cole na barra de fórmulas e pressione Enter. 6. Não arraste.`,
            expected: `Todos os registros com ${parsedFilter.criteriaValue} devem aparecer, não apenas o primeiro.`,
          });
          continue;
        }

        runtimeIssues.push({
          severity: "warning",
          code: "UNSUPPORTED_FORMULA",
          sheet: sheet.name,
          cell: formula.targetCell,
          message: "A fórmula dinâmica não pôde ser convertida para a alternativa compatível; foi mantida como fórmula única.",
        });
      }

      worksheet.getCell(formula.targetCell).value = { formula: cleanFormula(formula.expression) };
      formulaCount++;
      helpRows.push({
        targetSheet: sheet.name,
        target: formula.targetCell,
        type: formula.mode === "dynamic" ? "Matriz dinâmica" : "Fórmula única",
        purpose: "Executar o cálculo definido para esta célula.",
        originalFormula: cleanFormula(formula.expression),
        automaticFormula: cleanFormula(formula.expression),
        steps: `1. Abra a aba ${sheet.name}. 2. Clique em ${formula.targetCell}. 3. Cole a fórmula pela barra de fórmulas. 4. Pressione Enter.`,
        expected: "A célula deve exibir o resultado calculado.",
      });
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

    let statusColumnIndex = sheet.columns.findIndex((column) => normalizeText(column.key) === "status" || normalizeText(column.label) === "status") + 1;
    if (statusColumnIndex < 1) statusColumnIndex = sheet.columns.findIndex((column) => column.type === "status") + 1;
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

  addInstructionsSheet(workbook, helpRows);

  const runtimeErrors = runtimeIssues.filter((issue) => issue.severity === "error");
  const allWarnings = [...warnings, ...runtimeIssues.filter((issue) => issue.severity === "warning")];
  if (runtimeErrors.length) return { errors: runtimeErrors, warnings: allWarnings, formulaCount, downloaded: false };

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
