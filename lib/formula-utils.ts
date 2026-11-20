export type ParsedFilterFormula = {
  sourceSheet: string;
  sourceStartColumn: string;
  sourceStartRow: number;
  sourceEndColumn: string;
  sourceEndRow: number;
  criteriaSheet: string;
  criteriaColumn: string;
  criteriaStartRow: number;
  criteriaEndRow: number;
  criteriaValue: string;
};

export function cleanFormula(
  expression: string,
  rowNumber?: number
): string {
  const withoutEquals = expression
    .trim()
    .replace(/^=/, "")
    .replace(/;/g, ",");

  if (rowNumber === undefined) {
    return withoutEquals;
  }

  return withoutEquals.replaceAll("{row}", String(rowNumber));
}

export function normalizeText(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

export function localizeInstructionFormulaPtBr(
  expression: string
): string {
  let localized = cleanFormula(expression);

  const translations: Array<[RegExp, string]> = [
    [/\bAVERAGEIFS\s*\(/gi, "MÉDIASES("],
    [/\bAVERAGEIF\s*\(/gi, "MÉDIASE("],
    [/\bAVERAGE\s*\(/gi, "MÉDIA("],

    [/\bCOUNTIFS\s*\(/gi, "CONT.SES("],
    [/\bCOUNTIF\s*\(/gi, "CONT.SE("],
    [/\bCOUNTA\s*\(/gi, "CONT.VALORES("],
    [/\bCOUNT\s*\(/gi, "CONT.NÚM("],

    [/\bSUMIFS\s*\(/gi, "SOMASES("],
    [/\bSUMIF\s*\(/gi, "SOMASE("],
    [/\bSUM\s*\(/gi, "SOMA("],

    [/\bROUNDDOWN\s*\(/gi, "ARREDONDAR.PARA.BAIXO("],
    [/\bROUNDUP\s*\(/gi, "ARREDONDAR.PARA.CIMA("],
    [/\bROUND\s*\(/gi, "ARRED("],

    [/\bXLOOKUP\s*\(/gi, "PROCX("],
    [/\bFILTER\s*\(/gi, "FILTRO("],
    [/\bSORT\s*\(/gi, "CLASSIFICAR("],
    [/\bUNIQUE\s*\(/gi, "ÚNICO("],

    [/\bIFERROR\s*\(/gi, "SEERRO("],
    [/\bIF\s*\(/gi, "SE("],
    [/\bAND\s*\(/gi, "E("],
    [/\bOR\s*\(/gi, "OU("],
    [/\bNOT\s*\(/gi, "NÃO("],

    [/\bAGGREGATE\s*\(/gi, "AGREGAR("],
    [/\bCOLUMNS\s*\(/gi, "COLS("],
    [/\bCOLUMN\s*\(/gi, "COL("],
    [/\bROWS\s*\(/gi, "LINS("],
    [/\bROW\s*\(/gi, "LIN("],
    [/\bINDEX\s*\(/gi, "ÍNDICE("],
    [/\bMATCH\s*\(/gi, "CORRESP("],

    [/\bTODAY\s*\(/gi, "HOJE("],
    [/\bYEAR\s*\(/gi, "ANO("],
    [/\bMONTH\s*\(/gi, "MÊS("],
    [/\bDAY\s*\(/gi, "DIA("],
  ];

  for (const [pattern, replacement] of translations) {
    localized = localized.replace(pattern, replacement);
  }

  return localized.replace(/,/g, ";");
}

export function parseFilterFormula(
  expression: string
): ParsedFilterFormula | null {
  const formula = cleanFormula(expression);

  const pattern =
    /^(?:_xlfn\._xlws\.)?FILTER\(\s*'([^']+)'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\s*,\s*'([^']+)'!\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)\s*=\s*"([^"]*)"\s*,\s*""\s*\)$/i;

  const match = formula.match(pattern);

  if (!match) {
    return null;
  }

  return {
    sourceSheet: match[1],
    sourceStartColumn: match[2].toUpperCase(),
    sourceStartRow: Number(match[3]),
    sourceEndColumn: match[4].toUpperCase(),
    sourceEndRow: Number(match[5]),
    criteriaSheet: match[6],
    criteriaColumn: match[7].toUpperCase(),
    criteriaStartRow: Number(match[8]),
    criteriaEndRow: Number(match[9]),
    criteriaValue: match[10],
  };
}

export function isReservedSheetName(name: string): boolean {
  return normalizeText(name) ===
    normalizeText("Instruções e Fórmulas");
}
