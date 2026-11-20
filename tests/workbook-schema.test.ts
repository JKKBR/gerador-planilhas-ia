import { describe, expect, it } from "vitest";

import {
  formatWorkbookValidationErrors,
  validateWorkbook,
} from "../lib/workbook-schema";

function validWorkbook(): any {
  return {
    name: "Controle de Equipamentos",
    category: "Personalizada",
    description: "Planilha de teste.",
    sheets: [
      {
        id: "equipamentos",
        name: "Equipamentos",
        preparedRows: 20,
        columns: [
          {
            key: "patrimonio",
            label: "Patrimônio",
            type: "text",
            inputMode: "manual",
          },
          {
            key: "status",
            label: "Status",
            type: "status",
            inputMode: "manual",
          },
        ],
        rows: [
          {
            patrimonio: "PAT-001",
            status: "Pronto",
          },
        ],
        formulas: [],
      },
    ],
  };
}

describe("workbookSchema", () => {
  it("aceita um workbook válido", () => {
    const result = validateWorkbook(validWorkbook());

    expect(result.success).toBe(true);
  });

  it("rejeita a aba reservada ao exportador", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].name =
      "Instruções e Fórmulas";

    const result = validateWorkbook(workbook);

    expect(result.success).toBe(false);

    if (!result.success) {
      expect(
        formatWorkbookValidationErrors(result.error)
      ).toContain(
        'sheets.0.name: A aba "Instruções e Fórmulas" é reservada ao exportador.'
      );
    }
  });

  it("rejeita nomes de abas duplicados", () => {
    const workbook = validWorkbook();

    workbook.sheets.push({
      ...workbook.sheets[0],
      id: "equipamentos-2",
      name: "equipamentos",
    });

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita IDs de abas duplicados", () => {
    const workbook = validWorkbook();

    workbook.sheets.push({
      ...workbook.sheets[0],
      name: "Equipamentos 2",
    });

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita chaves de coluna duplicadas", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].columns.push({
      key: "status",
      label: "Situação",
      type: "status",
      inputMode: "manual",
    });

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita fórmula fill com coluna inexistente", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].formulas = [
      {
        mode: "fill",
        key: "total_inexistente",
        expression: "A{row}+1",
        fromRow: 2,
        toRow: 20,
      },
    ];

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita intervalo fill invertido", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].formulas = [
      {
        mode: "fill",
        key: "status",
        expression: 'IF(A{row}="","","OK")',
        fromRow: 20,
        toRow: 2,
      },
    ];

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita fórmula dynamic sem clearRange", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].formulas = [
      {
        mode: "dynamic",
        targetCell: "A2",
        expression:
          'FILTER(\'Equipamentos\'!A2:B21,' +
          '\'Equipamentos\'!B2:B21="Pronto","")',
      },
    ];

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita célula de destino inválida", () => {
    const workbook = validWorkbook();

    workbook.sheets[0].formulas = [
      {
        mode: "single",
        targetCell: "linha dois",
        expression: "SUM(A2:A10)",
      },
    ];

    expect(validateWorkbook(workbook).success).toBe(false);
  });

  it("rejeita quantidade excessiva de abas", () => {
    const workbook = validWorkbook();

    workbook.sheets = Array.from(
      { length: 9 },
      (_, index) => ({
        ...workbook.sheets[0],
        id: `aba-${index}`,
        name: `Aba ${index}`,
      })
    );

    expect(validateWorkbook(workbook).success).toBe(false);
  });
});
