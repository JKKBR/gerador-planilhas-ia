import { describe, expect, it } from "vitest";

import {
  cleanFormula,
  isReservedSheetName,
  localizeInstructionFormulaPtBr,
  normalizeText,
  parseFilterFormula,
} from "../lib/formula-utils";

describe("cleanFormula", () => {
  it("remove o sinal de igual inicial", () => {
    expect(cleanFormula("=SUM(A1:A5)")).toBe("SUM(A1:A5)");
  });

  it("converte ponto e vírgula para vírgula no formato técnico", () => {
    expect(cleanFormula("=SOMA(A1;B1)")).toBe("SOMA(A1,B1)");
  });

  it("substitui {row} pela linha solicitada", () => {
    const formula =
      'IF(A{row}="","",MAX(0,E{row}-D{row}))';

    expect(cleanFormula(formula, 8)).toBe(
      'IF(A8="","",MAX(0,E8-D8))'
    );
  });
});

describe("localizeInstructionFormulaPtBr", () => {
  it("traduz FILTER para FILTRO e usa ponto e vírgula", () => {
    const formula =
      'FILTER(\'Equipamentos\'!A2:H21,' +
      '\'Equipamentos\'!G2:G21="Reparando","")';

    expect(localizeInstructionFormulaPtBr(formula)).toBe(
      'FILTRO(\'Equipamentos\'!A2:H21;' +
      '\'Equipamentos\'!G2:G21="Reparando";"")'
    );
  });

  it("traduz COUNTIF para CONT.SE", () => {
    const formula =
      'COUNTIF(\'Equipamentos\'!$G$2:$G$21,"Defeito")';

    expect(localizeInstructionFormulaPtBr(formula)).toBe(
      'CONT.SE(\'Equipamentos\'!$G$2:$G$21;"Defeito")'
    );
  });

  it("traduz IFERROR antes de IF", () => {
    const formula = 'IFERROR(IF(A2="","",B2/C2),"")';

    expect(localizeInstructionFormulaPtBr(formula)).toBe(
      'SEERRO(SE(A2="";"";B2/C2);"")'
    );
  });

  it("traduz SUMIFS para SOMASES", () => {
    const formula =
      'SUMIFS(\'Peças\'!$D$2:$D$51,' +
      '\'Peças\'!$B$2:$B$51,A2,' +
      '\'Peças\'!$C$2:$C$51,B2)';

    expect(localizeInstructionFormulaPtBr(formula)).toBe(
      'SOMASES(\'Peças\'!$D$2:$D$51;' +
      '\'Peças\'!$B$2:$B$51;A2;' +
      '\'Peças\'!$C$2:$C$51;B2)'
    );
  });
});

describe("parseFilterFormula", () => {
  it("interpreta uma regra FILTER válida", () => {
    const parsed = parseFilterFormula(
      'FILTER(\'Equipamentos\'!A2:H21,' +
      '\'Equipamentos\'!G2:G21="Reparando","")'
    );

    expect(parsed).toEqual({
      sourceSheet: "Equipamentos",
      sourceStartColumn: "A",
      sourceStartRow: 2,
      sourceEndColumn: "H",
      sourceEndRow: 21,
      criteriaSheet: "Equipamentos",
      criteriaColumn: "G",
      criteriaStartRow: 2,
      criteriaEndRow: 21,
      criteriaValue: "Reparando",
    });
  });

  it("aceita referências absolutas", () => {
    const parsed = parseFilterFormula(
      'FILTER(\'Equipamentos\'!$A$2:$H$21,' +
      '\'Equipamentos\'!$G$2:$G$21="Defeito","")'
    );

    expect(parsed?.criteriaValue).toBe("Defeito");
    expect(parsed?.criteriaColumn).toBe("G");
    expect(parsed?.sourceEndRow).toBe(21);
  });

  it("retorna null para fórmula que não seja FILTER", () => {
    expect(
      parseFilterFormula(
        'COUNTIF(\'Equipamentos\'!G2:G21,"Pronto")'
      )
    ).toBeNull();
  });
});

describe("normalizeText", () => {
  it("remove acentos e espaços excedentes", () => {
    expect(normalizeText("  Instruções   e Fórmulas  ")).toBe(
      "instrucoes e formulas"
    );
  });
});

describe("isReservedSheetName", () => {
  it("identifica a aba reservada sem diferenciar acentos", () => {
    expect(isReservedSheetName("Instruções e Fórmulas")).toBe(true);
    expect(isReservedSheetName("instrucoes e formulas")).toBe(true);
  });

  it("permite nomes comuns", () => {
    expect(isReservedSheetName("Equipamentos")).toBe(false);
    expect(isReservedSheetName("Resumo")).toBe(false);
  });
});
