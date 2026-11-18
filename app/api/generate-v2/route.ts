import { NextResponse } from "next/server";
import type { WorkbookModel } from "@/lib/types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Category = "Estoque" | "Finanças" | "Orçamento" | "Personalizada";
const categories: Category[] = ["Estoque", "Finanças", "Orçamento", "Personalizada"];

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validWorkbook(value: unknown): value is WorkbookModel {
  if (!isObject(value) || typeof value.name !== "string") return false;
  if (typeof value.description !== "string" || typeof value.category !== "string") return false;
  if (!Array.isArray(value.sheets) || value.sheets.length < 1 || value.sheets.length > 8) return false;
  return value.sheets.every((sheet) =>
    isObject(sheet) &&
    typeof sheet.id === "string" &&
    typeof sheet.name === "string" &&
    Array.isArray(sheet.columns) &&
    sheet.columns.length > 0 &&
    Array.isArray(sheet.rows)
  );
}

function parseModelJson(text: string): unknown {
  return JSON.parse(
    text.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim()
  );
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const category = body.category as Category;
    const instruction = String(body.instruction || "").trim();
    const current = body.current;

    if (!categories.includes(category)) return NextResponse.json({ error: "Categoria inválida." }, { status: 400 });
    if (!instruction) return NextResponse.json({ error: "Digite uma solicitação." }, { status: 400 });

    if (instruction === "__diagnostico__") {
      return NextResponse.json({
        ok: true,
        routeVersion: "generate-v2-sem-schema",
        model: process.env.GEMINI_MODEL || "gemini-3.5-flash-lite",
        hasApiKey: Boolean(process.env.GEMINI_API_KEY),
      });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return NextResponse.json({ error: "GEMINI_API_KEY não configurada." }, { status: 500 });

    const model = process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";
    const prompt = `Retorne somente JSON válido, sem Markdown.
Crie ou altere uma planilha Excel da categoria ${category}.
Pedido: ${instruction}
Planilha atual: ${current ? JSON.stringify(current) : "nenhuma"}
Formato:
{"name":"Nome","category":"${category}","description":"Descrição","sheets":[{"id":"aba1","name":"Aba","purpose":"Finalidade","columns":[{"key":"campo","label":"Campo","type":"text","width":18}],"rows":[{"campo":"Exemplo"}],"formulas":[]}]}
Tipos: text, number, currency, date, status. Máximo de 8 abas, 20 colunas e 15 linhas por aba.
Para estoque, relacione cadastro, movimentações, saldo, reposição e dashboard.
Use fórmulas Excel em inglês com {row}. Não use macros nem links externos.`;

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
    const geminiResponse = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": apiKey },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      cache: "no-store",
    });

    const raw = await geminiResponse.text();
    if (!geminiResponse.ok) {
      console.error("Gemini v2 error", geminiResponse.status, raw);
      return NextResponse.json({ error: `Gemini HTTP ${geminiResponse.status}: ${raw}` }, { status: 502 });
    }

    const envelope = JSON.parse(raw);
    const text = envelope?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (typeof text !== "string") return NextResponse.json({ error: "O Gemini não retornou texto." }, { status: 502 });

    let workbook: unknown;
    try { workbook = parseModelJson(text); }
    catch { return NextResponse.json({ error: "O Gemini retornou JSON inválido." }, { status: 502 }); }

    if (!validWorkbook(workbook)) return NextResponse.json({ error: "Estrutura de planilha inválida." }, { status: 502 });

    return NextResponse.json({ workbook, model, routeVersion: "generate-v2-sem-schema" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Erro desconhecido.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
