import { NextResponse } from "next/server";

import { template } from "@/lib/templates";
import { WorkbookModel } from "@/lib/types";

export const runtime = "nodejs";

type Category =
  | "Estoque"
  | "Finanças"
  | "Orçamento"
  | "Personalizada";

const validCategories: Category[] = [
  "Estoque",
  "Finanças",
  "Orçamento",
  "Personalizada",
];

function isRecord(
  value: unknown
): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value)
  );
}

function validateWorkbook(
  value: unknown
): value is WorkbookModel {
  if (!isRecord(value)) {
    return false;
  }

  if (typeof value.name !== "string") {
    return false;
  }

  if (typeof value.description !== "string") {
    return false;
  }

  if (
    typeof value.category !== "string" ||
    !validCategories.includes(value.category as Category)
  ) {
    return false;
  }

  if (
    !Array.isArray(value.sheets) ||
    value.sheets.length === 0 ||
    value.sheets.length > 8
  ) {
    return false;
  }

  for (const sheet of value.sheets) {
    if (!isRecord(sheet)) {
      return false;
    }

    if (
      typeof sheet.id !== "string" ||
      typeof sheet.name !== "string"
    ) {
      return false;
    }

    if (
      !Array.isArray(sheet.columns) ||
      sheet.columns.length === 0 ||
      sheet.columns.length > 20
    ) {
      return false;
    }

    if (
      !Array.isArray(sheet.rows) ||
      sheet.rows.length > 30
    ) {
      return false;
    }

    if (
      sheet.formulas !== undefined &&
      !Array.isArray(sheet.formulas)
    ) {
      return false;
    }

    for (const column of sheet.columns) {
      if (!isRecord(column)) {
        return false;
      }

      if (
        typeof column.key !== "string" ||
        typeof column.label !== "string" ||
        typeof column.type !== "string"
      ) {
        return false;
      }

      if (
        ![
          "text",
          "number",
          "currency",
          "date",
          "status",
        ].includes(column.type)
      ) {
        return false;
      }
    }

    for (const row of sheet.rows) {
      if (!isRecord(row)) {
        return false;
      }
    }
  }

  return true;
}

function removeMarkdownCodeBlock(text: string) {
  return text
    .trim()
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "*)
    .replace(/\s*```$/, "")
    *trim();
}

export async function P*ST(request: Request) {
  try {
   *const body = await request.json();*
    const category = body.categor*;
    const instruction = body.ins*ruction;
    const current = body.*urrent;

    if (
      !validCate*ories.includes(category as Categor*)
    ) {
      return NextRespons*.json(
        {
          error: *Categoria inválida.",
        },
 *      {
          status: 400,
   *    }
      );
    }

    if (
   *  !instruction ||
      typeof ins*ruction !== "string"
    ) {
     *return NextResponse.json(
        *
          error: "Descreva a plan*lha desejada.",
        },
       *{
          status: 400,
        }*      );
    }

    if (instructio*.length > 8000) {
      return Nex*Response.json(
        {
         *error:
            "O pedido ultra*assou o limite de 8.000 caracteres*",
        },
        {
          *tatus: 400,
        }
      );
   *}

    const apiKey = process.env.*EMINI_API_KEY;

    if (!apiKey) {*      return NextResponse.json({
 *      workbook: current || templat*(category),
        warning:
     *    "Modo de demonstração: a chave*do Gemini não foi encontrada.",
  *   });
    }

    const model =
  *   process.env.GEMINI_MODEL ||
   *  "gemini-3.5-flash-lite";

    co*st currentWorkbook = current
     *? JSON.stringify(current)
      : *Nenhuma planilha anterior.";

    *onst prompt = `
Você cria estrutur*s de planilhas para Microsoft Exce*.

Retorne SOMENTE JSON válido.
Nã* use Markdown.
Não use bloco de có*igo.
Não adicione explicações ante* ou depois do JSON.

CATEGORIA:
${*ategory}

PEDIDO DO USUÁRIO:
${ins*ruction}

PLANILHA ATUAL:
${curren*Workbook}

Se existir uma planilha*atual:
- Preserve abas, colunas e *nformações não mencionadas.
- Apli*ue somente as alterações solicitad*s.
- Não remova funcionalidades ex*stentes sem pedido explícito.

Cri* um objeto exatamente com esta est*utura:

{
  "name": "Nome da plani*ha",
  "category": "${category}",
* "description": "Descrição curta",*  "sheets": [
    {
      "id": "i*entificador_sem_espacos",
      "n*me": "Nome da aba",
      "purpose*: "Finalidade da aba",
      "colu*ns": [
        {
          "key": *chave_sem_espacos_e_sem_acentos",
*         "label": "Título visível"*
          "type": "text",
       *  "width": 18
        }
      ],
 *    "rows": [
        {
          *chave_sem_espacos_e_sem_acentos": *Valor"
        }
      ],
      "f*rmulas": [
        {
          "ke*": "coluna_calculada",
          "*xpression": "C{row}+D{row}-E{row}"*        }
      ]
    }
  ]
}

Tip*s permitidos para as colunas:
- te*t
- number
- currency
- date
- sta*us

Regras:
- Use português do Bra*il nos textos visíveis.
- Use chav*s sem espaços e sem acentos.
- Cri* entre 1 e 8 abas.
- Crie no máxim* 20 colunas em cada aba.
- Crie no*máximo 15 linhas de exemplo em cad* aba.
- Inclua somente abas relaci*nadas ao pedido.
- Inclua Dashboar* quando houver indicadores numéric*s.
- Inclua dados coerentes entre *s abas.
- Para estoque, relacione *adastro, movimentações, saldo e re*osição.
- Use fórmulas do Excel em*inglês.
- Use {row} nas fórmulas p*ra representar o número da linha.
* O objeto rows deve usar as mesmas*chaves definidas em columns.
- Use*números verdadeiros para valores n*méricos, sem "R$" no valor.
- Não *ere macros.
- Não gere código VBA.*- Não gere links externos.
- Não c*ie propriedades além das indicadas*
`;

    const endpoint =
      "https://generativelanguage.googleapis.com/" +
      `v1beta/models/${encodeURIComponent(model)}` +
      ":generateContent";

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
      parts: [
        {
          text: prompt,
        },
      ],
    },
  ],
}),
    });

    const responseText =
      await geminiResponse.text();

    if (!geminiResponse.ok) {
      console.error("Erro retornado pelo Gemini:", {
        status: geminiResponse.status,
        model,
        response: responseText,
      });

      let geminiMessage = responseText;

      try {
        const parsedError =
          JSON.parse(responseText);

        geminiMessage =
          parsedError?.error?.message ||
          parsedError?.message ||
          responseText;
      } catch {
        // Mantém a resposta original.
      }

      return NextResponse.json(
        {
          error:
            `Gemini respondeu com HTTP ` +
            `${geminiResponse.status}: ` +
            geminiMessage,
        },
        {
          status: 502,
        }
      );
    }

    const geminiData = JSON.parse(responseText);

    const generatedText =
      geminiData?.candidates?.[0]?.content
        ?.parts?.[0]?.text;

    if (
      !generatedText ||
      typeof generatedText !== "string"
    ) {
      console.error(
        "Resposta sem conteúdo:",
        responseText
      );

      return NextResponse.json(
        {
          error:
            "O Gemini não retornou a estrutura da planilha.",
        },
        {
          status: 502,
        }
      );
    }

    const cleanedText =
      removeMarkdownCodeBlock(generatedText);

    let workbook: unknown;

    try {
      workbook = JSON.parse(cleanedText);
    } catch (error) {
      console.error(
        "JSON inválido retornado pelo Gemini:",
        generatedText
      );

      return NextResponse.json(
        {
          error:
            "O Gemini respondeu, mas o JSON retornado é inválido.",
        },
        {
          status: 502,
        }
      );
    }

    if (!validateWorkbook(workbook)) {
      console.error(
        "Estrutura de planilha inválida:",
        JSON.stringify(workbook)
      );

      return NextResponse.json(
        {
          error:
            "A IA retornou uma planilha com estrutura inválida. Tente novamente com um pedido menor.",
        },
        {
          status: 502,
        }
      );
    }

    return NextResponse.json({
      workbook,
      model,
    });
  } catch (error) {
    console.error(
      "Erro interno em /api/generate:",
      error
    );

    const message =
      error instanceof Error
        ? error.message
        : "Erro interno desconhecido.";

    return NextResponse.json(
      {
        error: message,
      },
      {
        status: 500,
      }
    );
  }
}
