import { NextResponse } from "next/server";
import type { WorkbookModel } from "@/lib/types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Category = "Estoque" | "Finanças" | "Orçamento" | "Personalizada";

const categories: Category[] = [
  "Estoque",
  "Finanças",
  "Orçamento",
  "Personalizada",
];

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validWorkbook(value: unknown): value is WorkbookModel {
  if (!isObject(value) || typeof value.name !== "string") return false;
  if (
    typeof value.description !== "string" ||
    typeof value.category !== "string"
  ) {
    return false;
  }

  if (
    !Array.isArray(value.sheets) ||
    value.sheets.length < 1 ||
    value.sheets.length > 8
  ) {
    return false;
  }

  return value.sheets.every(
    (sheet) =>
      isObject(sheet) &&
      typeof sheet.id === "string" &&
      typeof sheet.name === "string" &&
      Array.isArray(sheet.columns) &&
      sheet.columns.length > 0 &&
      Array.isArray(sheet.rows)
  );
}

function parseModelJson(text: string): unknown {
  const cleaned = text
    .trim()
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch {
    const firstBrace = cleaned.indexOf("{");
    const lastBrace = cleaned.lastIndexOf("}");

    if (firstBrace === -1 || lastBrace === -1 || lastBrace <= firstBrace) {
      throw new Error("Resposta sem objeto JSON.");
    }

    return JSON.parse(cleaned.slice(firstBrace, lastBrace + 1));
  }
}

function getGeminiError(raw: string) {
  try {
    const parsed = JSON.parse(raw);
    return parsed?.error?.message || parsed?.message || raw;
  } catch {
    return raw;
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const category = body.category as Category;
    const instruction = String(body.instruction || "").trim();
    const current = body.current;

    if (!categories.includes(category)) {
      return NextResponse.json(
        { error: "Categoria inválida." },
        { status: 400 }
      );
    }

    if (!instruction) {
      return NextResponse.json(
        { error: "Digite uma solicitação." },
        { status: 400 }
      );
    }

    if (instruction === "__diagnostico__") {
      return NextResponse.json({
        ok: true,
        routeVersion: "generate-v2-json-v3",
        model: process.env.GEMINI_MODEL || "gemini-3.5-flash-lite",
        hasApiKey: Boolean(process.env.GEMINI_API_KEY),
      });
    }

    const apiKey = process.env.GEMINI_API_KEY;

    if (!apiKey) {
      return NextResponse.json(
        { error: "GEMINI_API_KEY não configurada." },
        { status: 500 }
      );
    }

    const model = process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";

    const prompt = `Retorne somente JSON válido, sem Markdown.
Crie ou altere uma planilha Excel da categoria ${category}.

Pedido do usuário:
${instruction}

Planilha atual:
${current ? JSON.stringify(current) : "nenhuma"}

Formato obrigatório:
{"name":"Nome","category":"${category}","description":"Descrição","sheets":[{"id":"aba1","name":"Aba","purpose":"Finalidade","columns":[{"key":"campo","label":"Campo","type":"text","width":18}],"rows":[{"campo":"Exemplo"}],"formulas":[]}]}

Tipos de coluna permitidos: text, number, currency, date, status.
Máximo de 8 abas, 20 colunas e 15 linhas de exemplo por aba.
Use fórmulas Excel em inglês com {row} nas colunas calculadas.
Não use macros, VBA, links externos ou tabelas estruturadas do Excel.

Regras para linhas preparadas:
- Movimentações devem considerar as linhas 2 até 1001.
- Itens do Orçamento devem considerar as linhas 2 até 501.
- Lançamentos Financeiros devem considerar as linhas 2 até 1001.
- Cadastros devem considerar as linhas 2 até 301.
- Nas linhas vazias, fórmulas calculadas devem retornar vazio até que o campo principal seja preenchido.

Regras para abas de Resumo, Indicadores e Dashboard:
- Nunca preencha totais calculados com valores fixos.
- Todos os indicadores calculados devem ser fórmulas reais.
- Coloque a fórmula diretamente no valor da célula em rows, como uma string começando com =.
- Use nomes de funções do Excel em inglês e separadores por vírgula.
- Fórmulas de Resumo devem consultar todo o intervalo preparado.
- Use referências absolutas, como $J$2:$J$501.
- Não limite os resumos às linhas com dados de exemplo.
- Quando houver espaços no nome da aba, use aspas simples, como 'Itens do Orçamento'!$J$2:$J$501.
- Para valores por categoria, use SUMIF ou SUMIFS.
- Para totais gerais, use SUM.
- Para contagens, use COUNTIF ou COUNTA.

Exemplos para orçamento:
- Total final: =SUM('Itens do Orçamento'!$J$2:$J$501)
- Total de serviços: =SUMIF('Itens do Orçamento'!$B$2:$B$501,"Serviço",'Itens do Orçamento'!$J$2:$J$501)
- Quantidade de itens: =COUNTIF('Itens do Orçamento'!$A$2:$A$501,"<>")

Preserve abas e fórmulas existentes que não foram solicitadas para alteração.`;

    const endpoint =
      `https://generativelanguage.googleapis.com/v1beta/models/` +
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
        generationConfig: {
          responseMimeType: "application/json",
          temperature: 0.1,
          maxOutputTokens: 16000,
        },
      }),
      cache: "no-store",
    });

    const raw = await geminiResponse.text();

    if (!geminiResponse.ok) {
      const geminiError = getGeminiError(raw);
      console.error("Gemini v2 error", geminiResponse.status, raw);

      return NextResponse.json(
        {
          error: `Gemini HTTP ${geminiResponse.status}: ${geminiError}`,
        },
        { status: 502 }
      );
    }

    const envelope = JSON.parse(raw);
    const candidate = envelope?.candidates?.[0];
    const finishReason = candidate?.finishReason;

    if (finishReason === "MAX_TOKENS" || finishReason === "LENGTH") {
      return NextResponse.json(
        {
          error:
            "A resposta do Gemini ficou grande demais. Tente solicitar menos abas ou menos dados de exemplo.",
        },
        { status: 502 }
      );
    }

    const text = candidate?.content?.parts
      ?.map((part: { text?: string }) => part.text || "")
      .join("");

    if (typeof text !== "string" || !text.trim()) {
      return NextResponse.json(
        { error: "O Gemini não retornou texto." },
        { status: 502 }
      );
    }

    let workbook: unknown;

    try {
      workbook = parseModelJson(text);
    } catch {
      return NextResponse.json(
        { error: "O Gemini retornou JSON inválido." },
        { status: 502 }
      );
    }

    if (!validWorkbook(workbook)) {
      return NextResponse.json(
        { error: "Estrutura de planilha inválida." },
        { status: 502 }
      );
    }

    return NextResponse.json({
      workbook,
      model,
      routeVersion: "generate-v2-json-v3",
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Erro desconhecido.";

    console.error("generate-v2 internal error", error);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
