import { NextResponse } from "next/server";

import { template } from "@/lib/templates";
import type { WorkbookModel } from "@/lib/types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Category = "Estoque" | "Finanças" | "Orçamento" | "Personalizada";

const validCategories: Category[] = [
  "Estoque",
  "Finanças",
  "Orçamento",
  "Personalizada",
];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isValidWorkbook(value: unknown): value is WorkbookModel {
  if (!isRecord(value)) return false;

  if (
    typeof value.name !== "string" ||
    typeof value.description !== "string" ||
    typeof value.category !== "string" ||
    !validCategories.includes(value.category as Category) ||
    !Array.isArray(value.sheets) ||
    value.sheets.length === 0 ||
    value.sheets.length > 8
  ) {
    return false;
  }

  for (const sheet of value.sheets) {
    if (!isRecord(sheet)) return false;

    if (
      typeof sheet.id !== "string" ||
      typeof sheet.name !== "string" ||
      !Array.isArray(sheet.columns) ||
      sheet.columns.length === 0 ||
      sheet.columns.length > 20 ||
      !Array.isArray(sheet.rows) ||
      sheet.rows.length > 30 ||
      (sheet.formulas !== undefined && !Array.isArray(sheet.formulas))
    ) {
      return false;
    }

    for (const column of sheet.columns) {
      if (!isRecord(column)) return false;
      if (
        typeof column.key !== "string" ||
        typeof column.label !== "string" ||
        typeof column.type !== "string" ||
        !["text", "number", "currency", "date", "status"].includes(column.type)
      ) {
        return false;
      }
    }

    for (const row of sheet.rows) {
      if (!isRecord(row)) return false;
    }
  }

  return true;
}

function cleanJsonResponse(text: string) {
  return text
    .trim()
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();
}

function getGeminiError(responseText: string) {
  try {
    const parsed = JSON.parse(responseText);
    return parsed?.error?.message || parsed?.message || responseText;
  } catch {
    return responseText;
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const category = body.category as Category;
    const instruction = body.instruction;
    const current = body.current;

    if (!validCategories.includes(category)) {
      return NextResponse.json({ error: "Categoria inválida." }, { status: 400 });
    }

    if (typeof instruction !== "string" || !instruction.trim()) {
      return NextResponse.json(
        { error: "Descreva a planilha que deseja criar." },
        { status: 400 }
      );
    }

    if (instruction.length > 8000) {
      return NextResponse.json(
        { error: "O pedido ultrapassou o limite de 8.000 caracteres." },
        { status: 400 }
      );
    }

    if (instruction.trim() === "__diagnostico__") {
      return NextResponse.json({
        ok: true,
        routeVersion: "texto-json-v4",
        model: process.env.GEMINI_MODEL || "não configurado",
        hasApiKey: Boolean(process.env.GEMINI_API_KEY),
      });
    }

    const apiKey = process.env.GEMINI_API_KEY;

    if (!apiKey) {
      return NextResponse.json({
        workbook: current || template(category),
        warning: "Modo de demonstração: GEMINI_API_KEY não foi configurada.",
      });
    }

    const model = process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";
    const currentWorkbook = current
      ? JSON.stringify(current)
      : "Nenhuma planilha anterior.";

    const prompt = `
Você cria estruturas de planilhas para Microsoft Excel.
Retorne somente JSON válido, sem Markdown e sem explicações.

Categoria: ${category}
Pedido: ${instruction}
Planilha atual: ${currentWorkbook}

Retorne um objeto com esta estrutura:
{
  "name": "Nome da planilha",
  "category": "${category}",
  "description": "Descrição curta",
  "sheets": [
    {
      "id": "identificador",
      "name": "Nome da aba",
      "purpose": "Finalidade",
      "columns": [
        { "key": "chave", "label": "Título", "type": "text", "width": 18 }
      ],
      "rows": [
        { "chave": "Valor" }
      ],
      "formulas": [
        { "key": "coluna_calculada", "expression": "C{row}+D{row}-E{row}" }
      ]
    }
  ]
}

Tipos permitidos: text, number, currency, date, status.
Use português do Brasil nos textos visíveis.
Use chaves sem espaços e sem acentos.
Crie de 1 a 8 abas, até 20 colunas e até 15 linhas de exemplo por aba.
Inclua Dashboard quando houver indicadores.
Para estoque, relacione cadastro, movimentações, saldo e reposição.
Use fórmulas do Excel em inglês e {row} para representar a linha.
Não gere macros, VBA ou links externos.
`;

    const endpoint =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      `${encodeURIComponent(model)}:generateContent`;

    const geminiResponse = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: [{ text: prompt }],
          },
        ],
      }),
    });

    const responseText = await geminiResponse.text();

    if (!geminiResponse.ok) {
      const geminiError = getGeminiError(responseText);
      console.error("Erro retornado pelo Gemini:", {
        status: geminiResponse.status,
        model,
        response: responseText,
      });

      return NextResponse.json(
        {
          error: `Gemini respondeu com HTTP ${geminiResponse.status}: ${geminiError}`,
        },
        { status: 502 }
      );
    }

    const responseData = JSON.parse(responseText);
    const generatedText = responseData?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (typeof generatedText !== "string" || !generatedText.trim()) {
      return NextResponse.json(
        { error: "O Gemini não retornou a estrutura da planilha." },
        { status: 502 }
      );
    }

    let workbook: unknown;

    try {
      workbook = JSON.parse(cleanJsonResponse(generatedText));
    } catch {
      console.error("JSON inválido:", generatedText);
      return NextResponse.json(
        { error: "O Gemini retornou um JSON inválido. Tente novamente." },
        { status: 502 }
      );
    }

    if (!isValidWorkbook(workbook)) {
      console.error("Estrutura inválida:", JSON.stringify(workbook));
      return NextResponse.json(
        { error: "A IA retornou uma estrutura de planilha inválida." },
        { status: 502 }
      );
    }

    return NextResponse.json({
      workbook,
      model,
      routeVersion: "texto-json-v4",
    });
  } catch (error) {
    console.error("Erro interno em /api/generate:", error);
    const message = error instanceof Error ? error.message : "Erro desconhecido.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
