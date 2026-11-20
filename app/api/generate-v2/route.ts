import { NextResponse } from "next/server";
import {
  formatWorkbookValidationErrors,
  validateWorkbook,
} from "@/lib/workbook-schema";
import type { WorkbookModel } from "@/lib/types";


export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Category = "Estoque" | "Finanças" | "Orçamento" | "Personalizada";
const categories: Category[] = ["Estoque", "Finanças", "Orçamento", "Personalizada"];

// Modelos de contingência em ordem de prioridade.
const FALLBACK_MODELS = [
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-3.8-flash",
];

const ALLOWED_COLUMN_TYPES = new Set(["text", "number", "currency", "date", "status"]);
const ALLOWED_FORMULA_MODES = new Set(["fill", "single", "dynamic"]);
const ALLOWED_FUNCTIONS = new Set([
  "IF", "IFERROR", "AND", "OR", "NOT", "SUM", "SUMIF", "SUMIFS",
  "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "AVERAGE", "AVERAGEIF",
  "AVERAGEIFS", "MIN", "MAX", "ROUND", "ROUNDUP", "ROUNDDOWN",
  "ABS", "FILTER", "SORT", "UNIQUE", "XLOOKUP", "INDEX", "MATCH",
  "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "TODAY",
  "DATE", "YEAR", "MONTH", "DAY",
]);

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseModelJson(text: string): unknown {
  const cleaned = text.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const firstBrace = cleaned.indexOf("{");
    const lastBrace = cleaned.lastIndexOf("}");
    if (firstBrace === -1 || lastBrace <= firstBrace) throw new Error("Resposta sem objeto JSON.");
    return JSON.parse(cleaned.slice(firstBrace, lastBrace + 1));
  }
}

function normalizeText(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizeWorkbook(workbook: WorkbookModel): WorkbookModel {
  return {
    ...workbook,
    name: normalizeText(workbook.name),
    description: normalizeText(workbook.description),
    sheets: workbook.sheets.map((sheet) => ({
      ...sheet,
      id: normalizeText(sheet.id),
      name: normalizeText(sheet.name),
      purpose: sheet.purpose ? normalizeText(sheet.purpose) : undefined,
      preparedRows: Math.min(5000, Math.max(0, sheet.preparedRows ?? 0)),
      columns: sheet.columns.map((column) => ({
        ...column,
        key: normalizeText(column.key),
        label: normalizeText(column.label),
        inputMode: column.inputMode ?? "manual",
      })),
      rows: sheet.rows.map((row) => Object.fromEntries(
        Object.entries(row).map(([key, value]) => [key, typeof value === "string" && !value.trimStart().startsWith("=") ? normalizeText(value) : value])
      )),
      formulas: (sheet.formulas ?? []).map((formula) => ({
        ...formula,
        expression: formula.expression.trim(),
      })),
    })),
  };
}

function formulaFunctions(expression: string) {
  return [...expression.matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].map((match) => match[1].toUpperCase());
}

function workbookWarnings(workbook: WorkbookModel) {
  const warnings: string[] = [];
  const sheetNames = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));

  for (const sheet of workbook.sheets) {
    const keys = new Set(sheet.columns.map((column) => column.key));
    for (const formula of sheet.formulas ?? []) {
      if ((formula.mode ?? "fill") === "fill" && !("key" in formula && keys.has(formula.key))) {
        warnings.push(`${sheet.name}: fórmula ligada a uma coluna inexistente.`);
      }
      for (const fn of formulaFunctions(formula.expression)) {
        if (!ALLOWED_FUNCTIONS.has(fn)) warnings.push(`${sheet.name}: função ${fn} precisa de revisão.`);
      }
      for (const match of formula.expression.matchAll(/'([^']+)'!/g)) {
        if (!sheetNames.has(match[1].toLowerCase())) warnings.push(`${sheet.name}: a fórmula referencia a aba inexistente "${match[1]}".`);
      }
    }
  }
  return [...new Set(warnings)];
}

function getGeminiError(raw: string) {
  try {
    const parsed = JSON.parse(raw);
    return parsed?.error?.message || parsed?.message || raw;
  } catch {
    return raw;
  }
}

async function callGeminiWithFallback(apiKey: string, prompt: string, primaryModel?: string) {
  const modelsQueue = primaryModel
    ? [primaryModel, ...FALLBACK_MODELS.filter((model) => model !== primaryModel)]
    : FALLBACK_MODELS;
  let lastStatus = 502;
  let lastErrorMessage = "Falha ao comunicar com a API do Gemini.";

  for (const currentModel of modelsQueue) {
    const cleanModel = currentModel.trim();
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(cleanModel)}:generateContent`;
    for (let attempt = 1; attempt <= 2; attempt++) {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": apiKey },
        body: JSON.stringify({
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: {
            responseMimeType: "application/json",
            temperature: 0.05,
            maxOutputTokens: 16000,
          },
        }),
        cache: "no-store",
      });
      const raw = await response.text();
      if (response.ok) return { ok: true as const, raw, usedModel: cleanModel };

      lastStatus = response.status;
      lastErrorMessage = getGeminiError(raw).slice(0, 1000);
      console.warn("Falha na chamada Gemini", { model: cleanModel, attempt, status: response.status, error: lastErrorMessage });
      if (response.status === 404) break;
      if ((response.status === 503 || response.status === 429) && attempt === 1) {
        await sleep(1500);
        continue;
      }
      if (response.status === 400 || response.status === 403) {
        return { ok: false as const, status: lastStatus, error: lastErrorMessage };
      }
      break;
    }
  }
  return { ok: false as const, status: lastStatus, error: lastErrorMessage };
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
        routeVersion: "generate-v2-zod-v1",
        model: process.env.GEMINI_MODEL || FALLBACK_MODELS[0],
        fallbackModels: FALLBACK_MODELS,
        hasApiKey: Boolean(process.env.GEMINI_API_KEY),
      });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return NextResponse.json({ error: "GEMINI_API_KEY não configurada." }, { status: 500 });

    const prompt = `Retorne somente JSON válido, sem Markdown.
Crie ou altere uma planilha Excel da categoria ${category}.
Pedido do usuário:
${instruction}
Planilha atual:
${current ? JSON.stringify(current) : "nenhuma"}

FORMATO OBRIGATÓRIO:
{"name":"Nome","category":"${category}","description":"Descrição","sheets":[{"id":"aba1","name":"Aba","purpose":"Finalidade","preparedRows":300,"columns":[{"key":"campo","label":"Campo","type":"text","width":18,"inputMode":"manual"}],"rows":[{"campo":"Exemplo"}],"formulas":[{"mode":"fill","key":"total","expression":"IF(A{row}=\"\",\"\",B{row}*C{row})","fromRow":2,"toRow":301},{"mode":"dynamic","targetCell":"A2","clearRange":"A2:L1001","expression":"FILTER('Origem'!A2:L1001,'Origem'!J2:J1001=\"Reparando\",\"\")"}]}]}

REGRAS:
- Tipos: text, number, currency, date, status.
- inputMode: manual ou calculated.
- Máximo: 8 abas, 20 colunas, 15 linhas de exemplo, 5000 linhas preparadas.
- preparedRows é a capacidade de linhas de dados, sem contar o cabeçalho.
- Defina preparedRows explicitamente em toda aba de entrada ou resumo repetitivo.
- Use mode=fill para fórmulas repetidas por linha. Use {row} em toda referência relativa.
- Use fromRow e toRow. Para 1000 registros, use fromRow=2 e toRow=1001.
- Use mode=single para fórmula em uma célula única.
- Use mode=dynamic apenas para FILTER, SORT, UNIQUE ou matrizes; grave somente em targetCell e informe clearRange.
- Não coloque números fixos em colunas calculated.
- Não coloque fórmulas repetitivas diretamente em rows. Coloque-as em formulas com mode=fill.
- Valores pontuais em rows podem começar com =, mas prefira mode=single.
- Use funções em inglês e vírgulas.
- Use apenas: IF, IFERROR, AND, OR, NOT, SUM, SUMIF, SUMIFS, COUNT, COUNTA, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, AVERAGEIFS, MIN, MAX, ROUND, ROUNDUP, ROUNDDOWN, ABS, FILTER, SORT, UNIQUE, XLOOKUP, INDEX, MATCH, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, TODAY, DATE, YEAR, MONTH, DAY.
- Referencie o intervalo preparado completo, não somente exemplos.
- Nomes de abas com espaços devem usar aspas simples.
- Nas linhas vazias, fórmulas fill devem retornar vazio conforme os campos manuais usados como critério.
- Abas de filtro devem usar a mesma ordem de colunas da aba de origem.
- Não use macros, VBA, links externos ou tabelas estruturadas.
- Preserve abas e fórmulas não solicitadas.
- Remova espaços extras de nomes, status, peças, gerações e critérios.
- Se a planilha tiver resumos por peça/categoria, crie uma fórmula fill para cada coluna calculada até preparedRows+1.
- Nunca crie uma aba chamada "Instruções e Fórmulas". Essa aba é criada exclusivamente pelo exportador.
`;

    const result = await callGeminiWithFallback(apiKey, prompt, process.env.GEMINI_MODEL || FALLBACK_MODELS[0]);
    if (!result.ok) return NextResponse.json({ error: `Gemini HTTP ${result.status}: ${result.error}` }, { status: 502 });

    const envelope = JSON.parse(result.raw);
    const candidate = envelope?.candidates?.[0];
    if (["MAX_TOKENS", "LENGTH"].includes(candidate?.finishReason)) {
      return NextResponse.json({ error: "A resposta ficou grande demais. Solicite menos abas ou exemplos." }, { status: 502 });
    }
    const text = candidate?.content?.parts?.map((part: { text?: string }) => part.text || "").join("");
    if (typeof text !== "string" || !text.trim()) return NextResponse.json({ error: "O Gemini não retornou texto." }, { status: 502 });

    let parsed: unknown;
    try {
      parsed = parseModelJson(text);
    } catch {
      return NextResponse.json({ error: "O Gemini retornou JSON inválido." }, { status: 502 });
    }
const validation = validateWorkbook(parsed);

if (!validation.success) {
  const details = formatWorkbookValidationErrors(
    validation.error
  );

  console.error(
    "Workbook inválido retornado pelo Gemini",
    details
  );

  return NextResponse.json(
    {
      error:
        "A IA retornou uma estrutura de planilha inválida.",
      details: details.slice(0, 10),
    },
    {
      status: 502,
    }
  );
}

if (validation.data.category !== category) {
  return NextResponse.json(
    {
      error:
        "A categoria retornada pela IA não corresponde à categoria solicitada.",
      details: [
        `Categoria solicitada: ${category}`,
        `Categoria retornada: ${validation.data.category}`,
      ],
    },
    {
      status: 502,
    }
  );
}

const validatedWorkbook =
  validation.data as WorkbookModel;

const workbook = normalizeWorkbook(
  validatedWorkbook
);
    const warnings = workbookWarnings(workbook);
    return NextResponse.json({
      workbook,
      warnings,
      warning: warnings.length ? `Planilha gerada com ${warnings.length} aviso(s). Revise antes de baixar.` : undefined,
      model: result.usedModel,
      routeVersion: "generate-v2-zod-v1",
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Erro desconhecido.";
    console.error("generate-v2 internal error", error);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
