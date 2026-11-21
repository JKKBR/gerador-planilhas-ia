import { describe, expect, it } from "vitest";

import {
  inspectWorkbookQuality,
} from "../lib/workbook-quality";

import type {
  WorkbookModel,
} from "../lib/types";

function createWorkbook(): WorkbookModel {
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
        preparedRows: 2,
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

describe("inspectWorkbookQuality", () => {
  it("calcula os totais do relatório", () => {
    const report = inspectWorkbookQuality(
      createWorkbook()
    );

    expect(report.sheetCount).toBe(3);
    expect(report.columnCount).toBe(6);
    expect(report.exampleRowCount).toBe(3);
    expect(report.preparedRowCount).toBe(42);
    expect(report.formulaCount).toBe(2);
    expect(report.dynamicFormulaCount).toBe(1);
  });

  it("libera com aviso quando há fórmula dinâmica", () => {
    const report = inspectWorkbookQuality(
      createWorkbook()
    );

    expect(report.status).toBe("warning");
    expect(report.errors).toEqual([]);
    expect(
      report.warnings.some((warning) =>
        warning.includes("fórmula(s) dinâmica(s)")
      )
    ).toBe(true);
  });

  it("bloqueia estrutura inválida", () => {
    const report = inspectWorkbookQuality({
      name: "",
      sheets: [],
    });

    expect(report.status).toBe("blocked");
    expect(report.errors.length).toBeGreaterThan(0);
  });

  it("bloqueia coluna calculada sem fórmula", () => {
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

    const report =
      inspectWorkbookQuality(workbook);

    expect(report.status).toBe("blocked");
    expect(
      report.errors.some((error) =>
        error.includes(
          'coluna calculada "Resultado" sem regra de fórmula'
        )
      )
    ).toBe(true);
  });

  it("avisa quando não existem fórmulas", () => {
    const workbook = createWorkbook();

    workbook.sheets = [
      workbook.sheets[0],
    ];

    const report =
      inspectWorkbookQuality(workbook);

    expect(report.status).toBe("warning");
    expect(
      report.warnings
    ).toContain(
      "A planilha não possui regras de fórmula."
    );
  });

  it("avisa sobre aba vazia sem linhas preparadas", () => {
    const workbook = createWorkbook();

    workbook.sheets.push({
      id: "vazia",
      name: "Vazia",
      preparedRows: 0,
      columns: [
        {
          key: "campo",
          label: "Campo",
          type: "text",
          inputMode: "manual",
        },
      ],
      rows: [],
      formulas: [],
    });

    const report =
      inspectWorkbookQuality(workbook);

    expect(
      report.warnings.some((warning) =>
        warning.includes(
          'A aba "Vazia" não possui linhas preparadas'
        )
      )
    ).toBe(true);
  });
});
