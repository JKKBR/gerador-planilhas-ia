# Planilha IA

MVP de um gerador interno de planilhas Excel sem macros. Inclui modelos de Estoque, Finanças, Orçamento e Personalizada, edição por linguagem natural com Gemini, preview e exportação `.xlsx`.

## Requisitos

- Node.js 20 ou superior
- Uma chave da Gemini API

## Instalação local

```bash
npm install
cp .env.example .env.local
npm run dev
```

Abra `http://localhost:3000`.

No `.env.local`, configure:

```env
GEMINI_API_KEY=sua_chave
GEMINI_MODEL=gemini-2.5-flash
```

Sem a chave, o projeto abre em modo de demonstração com os modelos locais.

## Publicar no GitHub

```bash
git init
git add .
git commit -m "MVP do gerador de planilhas"
git branch -M main
git remote add origin SEU_REPOSITORIO
git push -u origin main
```

## Publicar na Vercel

1. Importe o repositório do GitHub na Vercel.
2. Em **Settings > Environment Variables**, adicione `GEMINI_API_KEY` e `GEMINI_MODEL`.
3. Faça um novo deploy.

Nunca use `NEXT_PUBLIC_` na chave do Gemini. A chamada é feita somente em `/api/generate` no servidor.

## O que já funciona

- Quatro modelos iniciais
- Alteração do workbook por texto
- Saída estruturada da IA
- Preview por abas
- Histórico local com desfazer
- Exportação `.xlsx`
- Tabelas, fórmulas, filtros, estilos e alertas simples
- Layout responsivo

## Limitações do MVP

- Não possui login nem persistência.
- O histórico desaparece ao atualizar a página.
- O gráfico do preview é visual; o XLSX exporta as tabelas e dados do Dashboard, não um gráfico nativo editável.
- Fórmulas geradas pela IA devem ser revisadas antes do uso operacional.
- Não gera macros.

## Próximas melhorias

- Supabase Auth e banco de dados
- Salvamento de projetos e versões
- Importação de `.xlsx`
- Gráficos nativos no arquivo por um serviço de geração compatível
- Permissões por operador
- Testes automatizados e limites por usuário
