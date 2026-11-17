import { NextResponse } from "next/server";
import { workbookSchema } from "@/lib/schema";
import { template } from "@/lib/templates";
export const runtime="nodejs";
export async function POST(req:Request){
 try{
  const {category,instruction,current}=await req.json();
  if(!instruction || typeof instruction!=="string") return NextResponse.json({error:"Descreva a planilha desejada."},{status:400});
  if(instruction.length>4000) return NextResponse.json({error:"Pedido muito longo."},{status:400});
  const key=process.env.GEMINI_API_KEY;
  if(!key) return NextResponse.json({workbook: current || template(category), warning:"Modo de demonstração: configure GEMINI_API_KEY para usar a IA."});
  const model=process.env.GEMINI_MODEL || "gemini-2.5-flash";
  const prompt=`Você é um projetista de planilhas Excel para uso interno. Retorne SOMENTE o workbook no esquema fornecido. Categoria: ${category}. Pedido: ${instruction}. ${current?`Planilha atual: ${JSON.stringify(current)}. Preserve tudo que não foi solicitado e aplique a alteração.`:"Crie uma planilha útil do zero."} Regras: nomes e conteúdo em pt-BR; no máximo 8 abas, 20 colunas e 30 linhas de exemplo; inclua abas com nexo entre si; inclua Dashboard quando houver dados numéricos; chaves de coluna sem espaços; fórmulas Excel em inglês usando {row}, por exemplo C{row}*D{row}; não crie macros, links externos nem código.`;
  const url=`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  const response=await fetch(url,{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({contents:[{parts:[{text:prompt}]}],generationConfig:{responseMimeType:"application/json",responseJsonSchema:workbookSchema,temperature:0.2}})});
  if(!response.ok){ const detail=await response.text(); console.error(detail); return NextResponse.json({error:"Não foi possível gerar com a IA. Verifique a chave e o modelo."},{status:502}); }
  const data=await response.json(); const text=data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if(!text) return NextResponse.json({error:"A IA não retornou uma planilha."},{status:502});
  return NextResponse.json({workbook:JSON.parse(text)});
 }catch(e){console.error(e);return NextResponse.json({error:"Erro ao processar o pedido."},{status:500});}
}
