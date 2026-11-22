## Estado atual do projeto

O Gerador de Planilhas com IA é um projeto funcional desenvolvido para transformar solicitações em linguagem natural em arquivos Excel editáveis.

A ferramenta permite descrever o controle desejado, revisar a estrutura interpretada pela IA, confirmar a aplicação, verificar a qualidade do arquivo e exportar a planilha em formato `.xlsx`.

O projeto está funcional para uso pessoal, testes e geração controlada de planilhas novas. O ambiente hospedado ainda não está sendo divulgado publicamente, principalmente para evitar consumo não autorizado da API e dos créditos disponíveis.

> Status atual: funcional em fase de desenvolvimento e validação.

---

## Funcionalidades implantadas

### Geração de planilhas com IA

O usuário pode descrever, em português, o tipo de planilha que deseja criar.

A ferramenta pode gerar modelos para:

- Controle de estoque;
- Controle financeiro;
- Orçamentos;
- Controle de equipamentos;
- Manutenção e reparos;
- Controle de licenças;
- Peças de reposição;
- Cadastros personalizados;
- Outros processos administrativos.

A resposta da IA é convertida em uma estrutura interna antes da criação do arquivo Excel.

---

### Categorias de geração

O sistema possui quatro categorias principais:

- Estoque;
- Finanças;
- Orçamento;
- Personalizada.

Cada categoria ajuda a orientar a criação da estrutura e das fórmulas.

---

### Múltiplos modelos Gemini

A integração possui uma fila de modelos de contingência.

Quando o modelo principal estiver indisponível ou sujeito a determinados erros temporários, o servidor pode tentar outro modelo configurado.

O diagnóstico da rota informa:

- Versão ativa;
- Modelo principal;
- Existência da chave da API;
- Modelos de contingência configurados.

---

### Validação estrutural com Zod

As respostas da IA são tratadas como dados não confiáveis e validadas antes de serem utilizadas.

A validação verifica, entre outros pontos:

- Nome e descrição da planilha;
- Categoria;
- Quantidade de abas;
- Nome e ID das abas;
- Colunas obrigatórias;
- Tipos de coluna;
- Chaves duplicadas;
- Fórmulas associadas a colunas inexistentes;
- Intervalos inválidos;
- Células de destino inválidas;
- Fórmulas dinâmicas sem intervalo de expansão;
- Uso indevido da aba reservada de instruções.

Estruturas inválidas são rejeitadas antes de chegar ao exportador.

---

### Tipos de fórmula

O modelo interno suporta três modos de fórmula.

#### Fórmulas `fill`

São replicadas automaticamente por várias linhas.

Exemplos:

- Subtotal;
- Valor de desconto;
- Saldo;
- Quantidade para reposição;
- Situação do estoque;
- Totais por peça e geração.

As fórmulas podem utilizar `{row}` como marcador da linha:

```text
IF(A{row}="","",MAX(0,E{row}-D{row}))
