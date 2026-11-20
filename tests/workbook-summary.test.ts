import { describe, expect, it } from "vitest";

import { buildWorkbookSummary } from "../lib/workbook-summary";
import type { WorkbookModel } from "../lib/types";

function createWorkbook(): WorkbookModel {
  return {
    name: "Controle de Equipamentos",
    category: "Personalizada",
    description:
      "Controle de equipamentos e manutenção.",
    sheets: [
      {
        id: "equipamentos",
        name: "Equipamentos",
        purpose: "Cadastro principal.",
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
          {
            patrimonio: "PAT-002",
            status: "Reparando",
          },
        ],
        formulas: [],
      },
      {
        id: "em-reparo",
        name: "Em Reparo",
        purpose:
          "Consulta dos equipamentos em reparo.",
        preparedRows: 20,
        columns: [
          {
            key: "patrimonio",
            label: "Patrimônio",
            type: "text",
            inputMode: "calculated",
          },
          {
            key: "status",
            label: "Status",
            type: "status",
            inputMode: "calculated",
          },
        ],
        rows: [],
        formulas: [
          {
            mode: "dynamic",
            targetCell: "A2",
            clearRange: "A2:B21",
            expression:
              'FILTER(\'Equipamentos\'!A2:B21,' +
              '\'Equipamentos\'!B2:B21="Reparando","")',
          },
        ],
      },
      {
        id: "resumo",
        name: "Resumo",
        purpose: "Indicadores.",
        preparedRows: 4,
        columns: [
          {
            key: "indicador",
            label: "Indicador",
            type: "text",
            inputMode: "manual",
          },
          {
            key: "quantidade",
            label: "Quantidade",
            type: "number",
            inputMode: "calculated",
          },
        ],
        rows: [
          {
            indicador: "Total",
            quantidade: null,
          },
        ],
        formulas: [
          {
            mode: "single",
            targetCell: "B2",
            expression:
              'COUNTIF(\'Equipamentos\'!A2:A21,"<>")',
          },
        ],
      },
    ],
  };
}

describe("buildWorkbookSummary", () => {
  it("resume os totais do workbook", () => {
    const summary =
      buildWorkbookSummary(createWorkbook());

    expect(summary.name).toBe(
      "Controle de Equipamentos"
    );
    expect(summary.sheetCount).toBe(3);
    expect(summary.totalColumns).toBe(6);
    expect(summary.totalExampleRows).toBe(3);
    expect(summary.totalPreparedRows).toBe(44);
    expect(summary.totalFormulas).toBe(2);
  });

  it("separa colunas manuais e calculadas", () => {
    const summary =
      buildWorkbookSummary(createWorkbook());

    const equipamentos = summary.sheets.find(
      (sheet) =>
        sheet.name === "Equipamentos"
    );

    const emReparo = summary.sheets.find(
      (sheet) =>
        sheet.name === "Em Reparo"
    );

    expect(equipamentos?.manualColumns).toEqual([
      "Patrimônio",
      "Status",
    ]);

    expect(emReparo?.calculatedColumns).toEqual([
      "Patrimônio",
      "Status",
    ]);
  });

  it("conta fórmulas por modo", () => {
    const summary =
      buildWorkbookSummary(createWorkbook());

    const emReparo = summary.sheets.find(
      (sheet) =>
        sheet.name === "Em Reparo"
    );

    const resumo = summary.sheets.find(
      (sheet) =>
        sheet.name === "Resumo"
    );

    expect(emReparo?.dynamicFormulaCount).toBe(1);
    expect(resumo?.singleFormulaCount).toBe(1);
    expect(resumo?.fillFormulaCount).toBe(0);
  });

  it("avisa sobre fórmulas dinâmicas", () => {
    const summary =
      buildWorkbookSummary(createWorkbook());

    expect(
      summary.warnings.some((warning) =>
        warning.includes(
          'A aba "Em Reparo" possui fórmula dinâmica'
        )
      )
    ).toBe(true);
  });

  it("avisa sobre coluna calculada sem fórmula", () => {
    const workbook = createWorkbook();

    workbook.sheets.push({
      id: "pendencias",
      name: "Pendências",
      preparedRows: 10,
      columns: [
        {
          key: "resultado",
          label: "Resultado",
          type: "text",
          inputMode: "calculated",
        },
      ],
      rows: [],
      formulas: [],
    });

    const summary =
      buildWorkbookSummary(workbook);

    expect(
      summary.warnings.some((warning) =>
        warning.includes(
          'A aba "Pendências" possui colunas calculadas sem regras de fórmula'
        )
      )
    ).toBe(true);
  });

  it("usa as linhas de exemplo quando preparedRows não existe", () => {
    const workbook = createWorkbook();

    delete workbook.sheets[0].preparedRows;

    const summary =
      buildWorkbookSummary(workbook);

    expect(
      summary.sheets[0].preparedRows
    ).toBe(2);
  });
});
