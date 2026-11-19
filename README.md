# 📊 Planilha IA

Gerador de planilhas Excel orientado por inteligência artificial.

O projeto transforma solicitações escritas em português em estruturas de planilhas editáveis, com abas relacionadas, colunas, fórmulas, filtros, alertas e áreas de resumo. A proposta é facilitar a criação de controles administrativos para pessoas e pequenos negócios que sabem quais informações precisam acompanhar, mas não dominam fórmulas avançadas do Excel.

> **Status:** protótipo funcional em evolução. O ambiente hospedado não é divulgado publicamente nesta fase, para evitar uso não controlado da API e consumo indevido de créditos.

---

## 🎯 Objetivo

Criar uma alternativa simples, rápida e acessível para prototipar planilhas personalizadas sem começar do zero.

A pessoa seleciona uma categoria, descreve a necessidade em linguagem natural, revisa a estrutura no navegador e exporta o resultado em formato `.xlsx`.

O projeto não pretende substituir sistemas empresariais completos. O foco atual é apoiar controles administrativos de baixa e média complexidade, especialmente quando uma planilha bem estruturada é suficiente para resolver o problema.

---

### Fallback entre modelos Gemini

A integração utiliza uma fila de modelos de contingência.

Ordem padrão:

1. `gemini-3.5-flash-lite`
2. `gemini-3.1-flash-lite`
3. `gemini-3.8-flash`

O modelo definido em `GEMINI_MODEL` é usado primeiro. Se ocorrer
indisponibilidade temporária, limite de requisições ou modelo não
encontrado, o servidor pode tentar os modelos seguintes.

A resposta da API interna informa qual modelo gerou a planilha.

O fallback melhora a disponibilidade, mas não substitui autenticação,
rate limiting, controle de cotas ou monitoramento de consumo.

---

## ✨ Recursos atuais

- Geração de planilhas a partir de instruções em português;
- Modelos iniciais de **Estoque**, **Finanças**, **Orçamento** e **Personalizada**;
- Criação de múltiplas abas relacionadas;
- Pré-visualização das abas antes do download;
- Fórmulas editáveis no Excel;
- Formatação de moeda, datas, números e status;
- AutoFiltro normal nos intervalos preparados;
- Linhas extras preparadas para preenchimento manual;
- Alertas visuais para situações como `OK`, `REPOR` e `ESGOTADO`;
- Histórico local com opção de desfazer a última alteração;
- Exportação em `.xlsx` sem macros;
- Layout responsivo para computador e dispositivos móveis;
- Modo de diagnóstico da integração com a IA.

### Linhas preparadas

O exportador pode preparar intervalos adicionais para que novas informações sejam inseridas manualmente sem a necessidade de recriar a planilha:

- Movimentações: até **1.000 linhas**;
- Lançamentos financeiros: até **1.000 linhas**;
- Itens de orçamento: até **500 linhas**;
- Cadastros, clientes, fornecedores, produtos, equipamentos e licenças: até **300 linhas**.

As colunas calculadas podem receber fórmulas antecipadamente. Enquanto o campo principal da linha estiver vazio, o resultado permanece em branco.

---

## 🧩 Exemplos de uso

### Controle de estoque

```text
Crie um controle de estoque de peças de informática com cadastro,
entradas, saídas, saldo atual, estoque mínimo, reposição e resumo.
```

### Controle financeiro

```text
Crie uma planilha financeira com lançamentos, categorias,
planejamento mensal, valores previstos, valores realizados e resumo.
```

### Orçamento

```text
Crie um orçamento de serviços com quantidade, valor unitário,
desconto percentual, subtotal, total por item e resumo geral.
```

### Controle de licenças

```text
Crie abas Office, Windows e Resumo para controlar patrimônio,
configuração do equipamento, serial, tipo de licença, método de
ativação e status de funcionamento.
```

---

## 🛠️ Tecnologias

- [Next.js](https://nextjs.org/) com App Router;
- [React](https://react.dev/);
- [TypeScript](https://www.typescriptlang.org/);
- [Gemini API](https://ai.google.dev/gemini-api/docs) para interpretação das solicitações;
- [ExcelJS](https://github.com/exceljs/exceljs) para geração dos arquivos `.xlsx`;
- [Lucide React](https://lucide.dev/) para ícones;
- [Vercel](https://vercel.com/) para hospedagem do ambiente de demonstração privado.

---

## 🏗️ Visão geral da arquitetura

```text
Solicitação do usuário
        ↓
Interface Next.js
        ↓
Route Handler no servidor
        ↓
Gemini API
        ↓
Estrutura JSON do workbook
        ↓
Preview no navegador
        ↓
ExcelJS
        ↓
Download do arquivo .xlsx
```

A chave da API permanece no servidor. Ela não deve ser criada com o prefixo `NEXT_PUBLIC_` e nunca deve ser enviada ao navegador.

---

## 📁 Estrutura principal

```text
app/
├── api/
│   └── generate-v2/
│       └── route.ts
├── globals.css
├── layout.tsx
└── page.tsx

lib/
├── export-xlsx.ts
├── templates.ts
└── types.ts
```

A estrutura pode mudar conforme a implementação do importador de arquivos e dos módulos de transformação.

---

## ⚙️ Instalação local

### Requisitos

- Node.js 20 ou superior;
- npm;
- Uma chave da Gemini API.

### Clonar e instalar

```bash
git clone https://github.com/JKKBR/gerador-planilhas-ia.git
cd gerador-planilhas-ia
npm install
```

### Variáveis de ambiente

Crie o arquivo `.env.local` na raiz do projeto:

```env
GEMINI_API_KEY=sua_chave_da_api
GEMINI_MODEL=seu_modelo_compativel
```

Não publique o arquivo `.env.local` e não registre a chave em commits, capturas de tela ou logs.

### Executar

```bash
npm run dev
```

Abra [o ambiente local do Next.js](http://localhost:3000).

### Validar a compilação

```bash
npm run build
```

O build deve ser executado antes de cada publicação para detectar erros de TypeScript e incompatibilidades.

---

## 🧪 Diagnóstico da integração

No campo do Assistente, envie:

```text
__diagnostico__
```

A aplicação deve informar:

- Versão da rota ativa;
- Modelo configurado;
- Presença da chave da API.

O diagnóstico não deve mostrar o valor da chave.

---

## 🔐 Segurança e controle de custos

Este repositório é público para fins de estudo e portfólio, mas o ambiente hospedado permanece privado nesta fase.

Medidas atuais e recomendadas:

- Não divulgar a URL de produção;
- Manter a chave da API somente no servidor;
- Não registrar chaves, documentos ou conteúdo sensível nos logs;
- Limitar o tamanho dos prompts;
- Limitar a quantidade de abas, colunas e linhas de exemplo geradas pela IA;
- Revisar fórmulas e dados antes do uso operacional;
- Usar informações fictícias em demonstrações públicas;
- Adicionar autenticação, cotas e rate limiting antes de permitir acesso externo.

Para planilhas de licenças, senhas ou documentos, prefira armazenar apenas identificadores mascarados. Um arquivo Excel não deve ser tratado como cofre de credenciais.

---

## ⚠️ Limitações atuais

- Não possui login ou controle de acesso;
- Não possui banco de dados ou persistência de projetos;
- O histórico é mantido somente durante a sessão atual;
- Não importa planilhas existentes nesta versão;
- Não gera macros ou VBA;
- Não trabalha com arquivos protegidos por senha;
- Não cria Tabelas Dinâmicas ou Power Query;
- Os gráficos mostrados no preview podem não ser exportados como gráficos nativos editáveis;
- Fórmulas produzidas pela IA devem ser revisadas antes do uso real;
- Solicitações muito grandes podem ultrapassar o limite de resposta do modelo;
- Resumos dependem de referências e intervalos corretos;
- O arquivo exportado não usa Tabelas Estruturadas do Excel nesta fase;
- A aplicação não substitui uma auditoria financeira, contábil, fiscal ou de segurança.

---

## 🗺️ Roadmap

### 1. Importação somente leitura de `.xlsx`

- Upload de arquivo existente;
- Validação de extensão e tamanho;
- Leitura de abas, valores, cabeçalhos e fórmulas;
- Preview limitado no navegador;
- Original preservado.

### 2. Seleção de abas e contexto compacto

- Escolha das abas relevantes;
- Detecção e correção da linha de cabeçalho;
- Até cinco registros de exemplo por aba;
- Até dez padrões de fórmula por aba;
- Envio de uma representação resumida para a IA.

### 3. Segurança e mascaramento

- Identificação de chaves, documentos, e-mails e números de série;
- Mascaramento configurável;
- Visualização do contexto que será enviado;
- Processamento local sempre que possível.

### 4. Plano de transformação controlado

A IA deverá retornar ações conhecidas, em vez de reescrever todos os registros:

- Copiar aba;
- Renomear aba;
- Adicionar coluna;
- Criar resumo;
- Mapear colunas;
- Usar um arquivo existente apenas como modelo.

### 5. Aplicação local das alterações

- Preservar registros existentes;
- Copiar fórmulas;
- Criar novas abas;
- Migrar dados entre estruturas;
- Gerar sempre outro arquivo, sem sobrescrever o original.

### 6. Compatibilidade e qualidade

- Verificação de referências inválidas;
- Validação de fórmulas;
- Melhoria de formatos de data, moeda e porcentagem;
- Testes de abertura no Excel;
- Testes automatizados de regressão.

### 7. Autenticação e limites

- Login;
- Cotas por usuário;
- Rate limiting;
- Limite de tokens e tamanho de arquivo;
- Monitoramento do consumo da API;
- Bloqueio de uso abusivo.

### 8. Projetos e versões

- Salvamento de projetos;
- Histórico de versões;
- Comparação antes e depois;
- Desfazer transformações;
- Modelos reutilizáveis.

### 9. Gráficos nativos

- Gráficos editáveis no arquivo final;
- Dados de apoio validados;
- Melhor integração entre Dashboard e exportação.

---

## 🔄 Visão futura do importador

```text
Upload do .xlsx
      ↓
Leitura local do arquivo
      ↓
Detecção de abas, cabeçalhos, dados e fórmulas
      ↓
Seleção e mascaramento
      ↓
Contexto compacto enviado à IA
      ↓
Plano de transformação
      ↓
Aplicação local sobre os dados
      ↓
Novo arquivo .xlsx
```

Princípio central:

> A IA propõe a estrutura e as ações. O código preserva e movimenta os dados.

Essa separação reduz o consumo de tokens e evita que milhares de registros precisem ser enviados e devolvidos pelo modelo.

---

## ✅ Critérios para a próxima versão

A primeira versão do importador será considerada pronta quando conseguir:

- Importar um arquivo `.xlsx` sem macros;
- Listar abas, cabeçalhos, valores e fórmulas;
- Permitir selecionar uma aba;
- Copiar uma aba com ou sem os dados;
- Preservar fórmulas simples;
- Criar uma aba de resumo;
- Gerar um arquivo novo;
- Manter o original intacto;
- Mascarar informações sensíveis;
- Evitar o envio do arquivo completo à IA.

---

## 🤝 Contribuição

O projeto está em fase experimental. Sugestões podem ser registradas em **Issues**, preferencialmente contendo:

- Descrição objetiva do problema;
- Passos para reproduzir;
- Resultado esperado;
- Resultado observado;
- Capturas sem dados sensíveis;
- Navegador e versão do Excel utilizados.

Antes de enviar código:

```bash
npm install
npm run build
```

Nunca inclua chaves de API, planilhas reais, documentos pessoais ou informações internas em commits e Issues.

---

## 📌 Uso responsável

As planilhas produzidas devem ser revisadas antes do uso operacional. Resultados financeiros, fiscais, contábeis, jurídicos, de estoque ou de licenciamento não devem ser considerados corretos sem conferência humana.

O projeto é uma demonstração de automação e apoio à produtividade. Cada organização continua responsável por seus dados, controles, permissões e processos de validação.

---

## 📄 Licença

Este repositório ainda não declara uma licença de software específica. Até que um arquivo `LICENSE` seja adicionado, não presuma permissão automática para redistribuição, uso comercial ou criação de versões derivadas.

---

## 👤 Autor

Desenvolvido como projeto pessoal de estudo e portfólio, com foco em inteligência artificial aplicada à organização de processos e criação de planilhas.
