import {
  describe,
  expect,
  it,
} from "vitest";

import {
  formatApiErrorMessage,
  readApiError,
} from "../lib/api-error";

describe("readApiError", () => {
  it("lê mensagem e detalhes", () => {
    const result = readApiError(
      {
        error:
          "A IA retornou uma estrutura inválida.",
        details: [
          "sheets.0.name: aba duplicada",
          "sheets.1.formulas: clearRange ausente",
        ],
      },
      502
    );

    expect(result).toEqual({
      message:
        "A IA retornou uma estrutura inválida.",
      details: [
        "sheets.0.name: aba duplicada",
        "sheets.1.formulas: clearRange ausente",
      ],
    });
  });

  it("usa mensagem HTTP como alternativa", () => {
    const result = readApiError(
      {},
      503
    );

    expect(result.message).toBe(
      "Não foi possível gerar a planilha. HTTP 503."
    );
  });

  it("ignora detalhes que não sejam texto", () => {
    const result = readApiError(
      {
        error: "Falha.",
        details: [
          "erro válido",
          123,
          null,
          {},
        ],
      },
      502
    );

    expect(result.details).toEqual([
      "erro válido",
    ]);
  });

  it("limita os detalhes a dez itens", () => {
    const result = readApiError(
      {
        details: Array.from(
          { length: 15 },
          (_, index) =>
            `Erro ${index + 1}`
        ),
      },
      502
    );

    expect(result.details).toHaveLength(10);
  });
});

describe("formatApiErrorMessage", () => {
  it("formata os detalhes em linhas", () => {
    const message =
      formatApiErrorMessage(
        {
          error: "Estrutura inválida.",
          details: [
            "Aba duplicada.",
            "Fórmula inválida.",
          ],
        },
        502
      );

    expect(message).toBe(
      [
        "Estrutura inválida.",
        "• Aba duplicada.",
        "• Fórmula inválida.",
      ].join("\n")
    );
  });
});
