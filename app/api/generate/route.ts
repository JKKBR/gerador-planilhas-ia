import { NextResponse } from "next/server";

import { workbookSchema } from "@/lib/schema";
import { template } from "@/lib/templates";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const body = await request.json();

    const category = body.category;
    const instruction = body.instruction;
    const current = body.current;

    if (!instruction || typeof instruction !== "string") {
      return NextResponse.json(
        {
          error: "Descreva a planilha desejada.",
        },
        {
          status: 400,
        }
      );
    }

    if (instruction.length > 8000) {
      return NextResponse.json(
        {
          error: "O pedido ultrapassou o limite de 8.000 caracteres.",
        },
        {
          status: 400,
        }
      );
    }

    const apiKey = process.env.GEMINI_API_KEY;

    if (!apiKey) {
      return NextResponse.json({
        workbook: current || template(category),
        warning:
          "Modo de demonstração: GEMINI_API_KEY não foi encontrada.",
      });
    }

    const model =
      process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";

    const prompt = `
Você é um projetista de planilhas Excel para uso interno.

Categoria selecionada:
${category}

Solicitação:
${instruction}

${
  current
    ? `Planilha atual:
${JSON.stringify(current)}

Preserve tudo o que não foi solicitado e aplique somente as alterações necessárias.`
    : "Crie uma planilha útil do zero."
}

Regras:
- Use português do Brasil.
- Crie no máximo 8 abas.
- Crie no máximo 20 colunas por aba.
- Inclua até 20 linhas de exemplo.
- Crie abas relacionadas entre si.
- Inclua uma aba Dashboard quando houver dados numéricos.
- Use chaves de coluna sem espaços e sem acentos.
- Use fórmulas do Excel em inglês.
- Nas fórmulas, use {row} para representar a linha.
- Não crie macros.
- Não crie links externos.
- Não produza código executável.
- Retorne somente o objeto JSON solicitado.
`;

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
            parts: [
              {
                text: prompt,
              },
            ],
          },
        ],
        generationConfig: {
          responseMimeType: "application/json",
          responseJsonSchema: workbookSchema,
          temperature: 0.2,
        },
      }),
    });

    const responseText = await geminiResponse.text();

    if (!geminiResponse.ok) {
      console.error("Erro retornado pelo Gemini:", {
        status: geminiResponse.status,
        model,
        response: responseText,
      });

      let geminiMessage = responseText;

      try {
        const parsedError = JSON.parse(responseText);

        geminiMessage =
          parsedError?.error?.message ||
          parsedError?.message ||
          responseText;
      } catch {
        // Mantém a resposta textual original.
      }

      return NextResponse.json(
        {
          error: `Gemini respondeu com HTTP ${geminiResponse.status}: ${geminiMessage}`,
          model,
        },
        {
          status: 502,
        }
      );
    }

    const geminiData = JSON.parse(responseText);

    const generatedText =
      geminiData?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!generatedText) {
      console.error(
        "Resposta do Gemini sem conteúdo:",
        responseText
      );

      return NextResponse.json(
        {
          error:
            "O Gemini respondeu, mas não retornou a estrutura da planilha.",
          model,
        },
        {
          status: 502,
        }
      );
    }

    let workbook;

    try {
      workbook = JSON.parse(generatedText);
    } catch {
      console.error(
        "JSON inválido retornado pelo Gemini:",
        generatedText
      );

      return NextResponse.json(
        {
          error:
            "O Gemini respondeu, mas o conteúdo não era um JSON válido.",
          model,
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
    console.error("Erro interno em /api/generate:", error);

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
