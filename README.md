# Gerador de Planilhas com IA

Aplicação web que transforma solicitações em linguagem natural em planilhas Excel editáveis. O projeto foi criado para facilitar a elaboração de controles administrativos para pessoas que sabem o que precisam acompanhar, mas não dominam a criação de abas, fórmulas, filtros e resumos no Excel.

O fluxo principal é:

```text
Descrever a necessidade
→ gerar a estrutura com IA
→ validar o workbook
→ revisar o entendimento
→ confirmar a aplicação
→ analisar a qualidade
→ exportar o arquivo .xlsx
```

> **Status:** funcional para uso pessoal, aprendizado, portfólio e testes controlados. O projeto continua em desenvolvimento e ainda não deve ser considerado uma solução empresarial definitiva.

## Objetivos

- Criar planilhas novas a partir de instruções em português;
- Reduzir a necessidade de conhecimento avançado em Excel;
- Produzir arquivos `.xlsx` editáveis e preparados para uso futuro;
- Tornar fórmulas, avisos e limitações mais transparentes;
- Permitir a criação de controles personalizados sem depender de modelos fixos;
- Servir como projeto prático de aprendizado em IA, TypeScript, Next.js, testes e automação de Excel.

## Casos de uso

A ferramenta pode gerar estruturas para:

- Estoque e reposição;
- Finanças e despesas;
- Orçamentos;
- Equipamentos e patrimônio;
- Manutenção e reparos;
- Licenças de software;
- Peças de reposição;
- Cadastros administrativos;
- Resumos e indicadores;
- Processos personalizados.

## Funcionalidades implantadas

### Geração por linguagem natural

O usuário descreve a planilha desejada em português. A IA cria o nome, a descrição, as abas, as colunas, os exemplos, as linhas preparadas e as regras de cálculo.

### Categorias

A interface possui quatro categorias principais:

- Estoque;
- Finanças;
- Orçamento;
- Personalizada.

### Modelos Gemini de contingência

A API utiliza uma fila de modelos compatíveis. Se o modelo principal estiver indisponível ou retornar determinados erros temporários, o servidor pode tentar outro modelo configurado.

A rota também possui um comando de diagnóstico que informa:

- Versão ativa da rota;
- Modelo principal;
- Existência da chave da API;
- Modelos de contingência configurados.

### Validação estrutural com Zod

Toda resposta da IA é tratada como dado não confiável e validada antes de ser utilizada.

A validação verifica, entre outros pontos:

- Nome, descrição e categoria;
- Quantidade de abas;
- Nomes e IDs duplicados;
- Quantidade e tipos de colunas;
- Chaves de colunas duplicadas;
- Limites de linhas e exemplos;
- Fórmulas associadas a colunas inexistentes;
- Intervalos e células de destino inválidos;
- Fórmulas dinâmicas sem `clearRange`;
- Uso indevido da aba reservada de instruções.

### Modos de fórmula

O modelo interno suporta três modos.

#### `fill`

Replica a fórmula pelas linhas preparadas.

```text
IF(A{row}="","",MAX(0,E{row}-D{row}))
```

O marcador `{row}` é substituído pelo número correto de cada linha.

#### `single`

Grava uma fórmula em uma única célula, geralmente usada para indicadores e resumos.

```text
COUNTIF('Equipamentos'!$H$2:$H$41,"Pronto")
```

#### `dynamic`

Representa consultas automáticas entre abas, como equipamentos em reparo ou com defeito.

```text
FILTER('Equipamentos'!A2:K41,'Equipamentos'!H2:H41="Reparando","")
```

Como fórmulas de matriz dinâmica podem variar conforme a versão e o comportamento do Excel, o exportador também prepara valores iniciais e instruções de uso manual.

### Linhas preparadas

Cada aba pode informar a quantidade de linhas disponíveis para registros futuros. Fórmulas do tipo `fill` são aplicadas até o final do intervalo, evitando que o usuário precise copiá-las manualmente.

Exemplos de capacidade:

- Equipamentos e movimentações em grande volume;
- Centenas de peças ou cadastros;
- Itens de orçamento;
- Quantidades personalizadas definidas pelo pedido.

### Abas filtradas com valores iniciais

Abas como `Em Reparo` e `Com Defeito` podem abrir com os registros correspondentes já visíveis. A fórmula original continua documentada para ativação manual, quando necessário.

### Aba `Instruções e Fórmulas`

O exportador cria automaticamente uma única aba reservada com:

- Aba de destino;
- Célula ou intervalo;
- Tipo de fórmula;
- Finalidade;
- Fórmula para copiar;
- Fórmula técnica;
- Passo a passo;
- Resultado esperado.

A IA é instruída a não criar essa aba. O exportador também protege contra duplicação.

### Fórmulas manuais em português

As fórmulas técnicas permanecem em inglês para o processamento interno. As fórmulas destinadas à inserção manual são apresentadas em português do Brasil.

Exemplo técnico:

```excel
=FILTER('Equipamentos'!A2:H21,'Equipamentos'!G2:G21="Reparando","")
```

Exemplo mostrado ao usuário:

```excel
=FILTRO('Equipamentos'!A2:H21;'Equipamentos'!G2:G21="Reparando";"")
```

### Confirmação do entendimento

A nova estrutura não substitui imediatamente a planilha atual. Antes de aplicar, a interface apresenta o painel **Confira o entendimento**, com:

- Nome e descrição;
- Quantidade de abas;
- Quantidade de colunas;
- Linhas preparadas;
- Regras de fórmula;
- Colunas manuais;
- Colunas calculadas;
- Avisos importantes.

O usuário pode escolher:

- **Voltar e ajustar**;
- **Confirmar e aplicar**.

### Relatório de qualidade antes do download

O download não deve ocorrer no primeiro clique. A ferramenta analisa a estrutura e apresenta um relatório com três estados:

- **Pronto para baixar**;
- **Pronto com avisos**;
- **Download bloqueado**.

O relatório apresenta:

- Abas;
- Colunas;
- Linhas preparadas;
- Regras de fórmula;
- Erros críticos;
- Avisos de compatibilidade ou revisão.

O arquivo só é gerado após **Confirmar download** ou **Baixar mesmo assim**.

### Detalhes dos erros da API

A interface pode apresentar a mensagem principal e os detalhes retornados pela validação.

```text
A IA retornou uma estrutura de planilha inválida.

• sheets.1.formulas: fórmula dinâmica sem clearRange
• sheets.2.columns.3.key: chave duplicada
```

### Histórico e desfazer

A interface mantém um histórico básico durante a sessão. O usuário pode aplicar uma nova versão e desfazer a última alteração.

### Exportação XLSX

A exportação utiliza ExcelJS e pode incluir:

- Múltiplas abas;
- Cabeçalhos formatados;
- Larguras de coluna;
- AutoFiltro comum;
- Congelamento da primeira linha;
- Fórmulas reais;
- Formatação de moeda, datas e números;
- Formatação condicional de status;
- Linhas preparadas;
- Aba de instruções;
- Solicitação de recálculo ao abrir.

### Testes automatizados

O projeto utiliza Vitest. Os testes cobrem:

- Limpeza de fórmulas;
- Substituição de `{row}`;
- Tradução de fórmulas para português;
- Interpretação de `FILTER`;
- Identificação da aba reservada;
- Validação estrutural com Zod;
- Resumo de confirmação;
- Relatório de qualidade;
- Leitura e formatação de erros da API;
- Configuração do ambiente de testes.

### Verificação automática

O comando principal é:

```bash
npm run check
```

Ele executa:

```text
TypeScript
→ testes automatizados
→ build de produção
```

O GitHub Actions executa verificações em branches e Pull Requests. A Vercel pode ser usada para validar um Preview antes do merge na `main`.

## A ferramenta já é funcional?

Sim. A aplicação já é funcional para:

- Criar planilhas novas;
- Gerar múltiplas abas;
- Preparar linhas futuras;
- Aplicar fórmulas por linha e por célula;
- Criar abas automáticas;
- Criar resumos e indicadores;
- Revisar a estrutura antes da aplicação;
- Validar a qualidade antes do download;
- Exibir avisos e erros estruturados;
- Exportar arquivos `.xlsx` editáveis;
- Orientar ajustes manuais quando necessário.

O uso atual é recomendado para:

- Uso pessoal;
- Aprendizado;
- Portfólio;
- Testes controlados;
- Criação de modelos administrativos;
- Pequenos controles de estoque, orçamento, equipamentos e manutenção.

A revisão humana ainda é recomendada antes de utilizar o arquivo em processos críticos.

## Limitações atuais

- Não importa ou remodela planilhas existentes;
- Não preserva macros;
- Não trabalha com Power Query;
- Não cria Tabelas Dinâmicas nativas;
- Gráficos e dashboards ainda são limitados;
- Não possui autenticação;
- Não possui rate limiting;
- Não possui armazenamento permanente;
- O histórico é temporário e local;
- Algumas fórmulas modernas podem exigir ajuste manual;
- O resultado pode variar conforme a versão do Excel;
- A qualidade ainda depende parcialmente da interpretação da IA;
- O ambiente hospedado permanece privado para proteger a API e os créditos.

## Instalação local

```bash
git clone <URL_DO_REPOSITORIO>
cd gerador-planilhas-ia
npm install
```

Crie um arquivo `.env.local`:

```env
GEMINI_API_KEY=sua_chave
GEMINI_MODEL=modelo_principal
```

Inicie o ambiente:

```bash
npm run dev
```

Abra o endereço informado pelo Next.js.

> Nunca envie `.env.local` para o repositório.

## Scripts

```bash
npm run dev
npm run typecheck
npm run test
npm run test:watch
npm run build
npm run check
```

## Fluxo de desenvolvimento

```text
Criar branch
→ implementar
→ executar npm run check
→ push
→ Pull Request
→ GitHub Actions
→ Preview da Vercel
→ validação manual
→ merge na main
```

## Próximas etapas prioritárias

### 1. Correção automática de estruturas inválidas

Quando a primeira resposta for rejeitada, o servidor poderá enviar os erros ao modelo e solicitar uma única correção completa.

```text
Resposta da IA
→ validação Zod
→ erros estruturais
→ uma tentativa automática de correção
→ nova validação
```

### 2. Tratamento diferenciado de falhas

Separar:

- Erros temporários da API;
- Falhas de autenticação;
- Limites de uso;
- JSON inválido;
- Workbook inválido;
- Fórmulas inválidas.

### 3. Perguntas de esclarecimento

Quando o pedido tiver mais de uma interpretação, a interface poderá perguntar antes da geração.

```text
“Total de Peças” significa:

1. Somar unidades;
2. Contar tipos cadastrados;
3. Mostrar os dois.
```

### 4. Testes do XLSX real

Gerar o arquivo em memória e reabri-lo para verificar:

- Abas;
- Cabeçalhos;
- Fórmulas;
- Formatos;
- AutoFiltro;
- Linhas preparadas;
- Aba de instruções;
- Ausência de duplicações;
- Valores filtrados.

### 5. Separar criação, auditoria e download

```text
Criar workbook
→ gerar buffer
→ reabrir e auditar
→ liberar download
```

### 6. Segurança de fórmulas

Bloquear:

- Referências externas;
- Links para outros arquivos;
- Funções de acesso à Web;
- Conexões não solicitadas;
- Fórmulas potencialmente perigosas.

### 7. Controle de consumo

Antes de uma publicação aberta:

- Rate limiting;
- Limite por IP;
- Limite por sessão;
- Limite diário;
- Limite de tamanho do prompt;
- Controle de chamadas simultâneas;
- Proteção contra automação abusiva;
- Monitoramento de créditos.

## Possíveis atualizações futuras

### Histórico de versões

- Visualizar versões;
- Restaurar uma versão;
- Comparar mudanças;
- Baixar uma versão anterior.

### Comparação antes e depois

- Abas adicionadas ou removidas;
- Colunas adicionadas ou removidas;
- Fórmulas alteradas;
- Mudanças nas linhas preparadas.

### Testes de interface

Simular o fluxo completo com testes de ponta a ponta:

```text
Pedido
→ resposta
→ confirmação
→ aplicação
→ relatório
→ download
→ desfazer
```

### Acessibilidade

- Controle de foco;
- Fechar painéis com `Esc`;
- Retorno do foco ao botão de origem;
- Navegação por teclado;
- `aria-live` para mensagens;
- Melhor contraste;
- Bloqueio da navegação atrás de modais.

### Importação de XLSX

Escopo inicial previsto:

- Arquivos `.xlsx`;
- Sem macros;
- Sem senha;
- Sem Power Query;
- Sem conexões externas;
- Preservação de dados e fórmulas simples;
- Geração de um arquivo novo.

### Remodelação de planilhas

Fluxo futuro:

```text
Upload
→ leitura local
→ análise das abas
→ mascaramento de dados sensíveis
→ pedido de alteração
→ plano de transformação
→ novo arquivo
```

### Mascaramento de dados

Ocultar antes do envio à IA:

- Chaves de licença;
- CPF e CNPJ;
- E-mails;
- Telefones;
- Números de série;
- Tokens;
- Senhas;
- Outros campos sensíveis.

### Gráficos e dashboards

Possibilidades:

- Colunas;
- Linhas;
- Pizza;
- Comparações por período;
- Indicadores visuais;
- Dashboards administrativos.

### Autenticação e contas

- Login;
- Limites por usuário;
- Histórico persistente;
- Modelos privados;
- Preferências;
- Área administrativa;
- Controle de consumo.

## Segurança e privacidade

- A chave Gemini fica somente no servidor;
- O arquivo `.env.local` não deve ser versionado;
- Dados sensíveis não devem ser usados nos exemplos;
- Logs devem evitar prompts completos e informações pessoais;
- Arquivos exportados devem ser revisados antes de uso crítico;
- O ambiente público deve permanecer desativado até existir proteção contra consumo abusivo.

## Tecnologias

- Next.js;
- React;
- TypeScript;
- Gemini API;
- ExcelJS;
- Zod;
- Vitest;
- GitHub Actions;
- Vercel.

## Posicionamento

O objetivo não é substituir plataformas profissionais completas de planilhas com IA.

O projeto busca oferecer uma experiência guiada para pessoas que sabem o que precisam controlar, mas não dominam a criação de abas, fórmulas, filtros e resumos no Excel.

```text
Descrever a necessidade
→ revisar o entendimento
→ confirmar a estrutura
→ verificar a qualidade
→ baixar um Excel editável
```

As prioridades do projeto são:

- Simplicidade;
- Português do Brasil;
- Arquivos Excel editáveis;
- Transparência das fórmulas;
- Validação;
- Segurança;
- Personalização;
- Aprendizado técnico.

## Aviso

Este projeto está em desenvolvimento. As planilhas são geradas por inteligência artificial e devem ser revisadas antes de serem utilizadas em processos financeiros, legais, contábeis, operacionais ou administrativos críticos.
