"use client";

import { useMemo, useState } from "react";
import {
  Download,
  Plus,
  RotateCcw,
  Send,
  Sparkles,
  Table2,
} from "lucide-react";

import { exportXlsx } from "@/lib/export-xlsx";
import { buildWorkbookSummary } from "@/lib/workbook-summary";
import { template } from "@/lib/templates";
import type { WorkbookModel } from "@/lib/types";

const categories = [
  "Estoque",
  "Finanças",
  "Orçamento",
  "Personalizada",
] as const;

export default function Home() {
  const [category, setCategory] =
    useState<(typeof categories)[number]>("Estoque");
  const [book, setBook] = useState<WorkbookModel>(() => template("Estoque"));
  const [active, setActive] = useState(0);
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(
    "Escolha um modelo e descreva o que deseja alterar."
  );
  const [history, setHistory] = useState<WorkbookModel[]>([]);

  const [pendingBook, setPendingBook] =
    useState<WorkbookModel | null>(null);

  const pendingSummary = useMemo(
    () =>
      pendingBook
        ? buildWorkbookSummary(pendingBook)
        : null,
    [pendingBook]
  );
  const [warnings, setWarnings] = useState<string[]>([]);

  const sheet = book.sheets[active] || book.sheets[0];

  const chart = useMemo(() => {
    if (!sheet) return [];

    const labelKey =
      sheet.columns.find((column) => column.type === "text")?.key ||
      sheet.columns[0]?.key;
    const numberKey = sheet.columns.find((column) =>
      ["number", "currency"].includes(column.type)
    )?.key;

    if (!labelKey || !numberKey) return [];

    return sheet.rows
      .slice(0, 8)
      .map((row) => ({
        label: String(row[labelKey] ?? "Item"),
        value: Number(row[numberKey] ?? 0),
      }))
      .filter((item) => Number.isFinite(item.value));
  }, [sheet]);

  async function generate() {
    const instruction = text.trim();

    if (!instruction) {
      setMessage("Digite uma solicitação antes de enviar.");
      return;
    }

    setLoading(true);
    setMessage("Enviando pedido para a IA...");

    try {
      const response = await fetch("/api/generate-v2", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          category,
          instruction,
          current: book,
        }),
      });

      const responseText = await response.text();
      let data: any;

      try {
        data = JSON.parse(responseText);
      } catch {
        throw new Error(
          `A API retornou uma resposta inválida. HTTP ${response.status}.`
        );
      }

      if (!response.ok) {
        throw new Error(
          data.error ||
            `Não foi possível gerar a planilha. HTTP ${response.status}.`
        );
      }

      if (data.ok && data.routeVersion) {
        setMessage(
          [
            "Diagnóstico concluído.",
            `Rota: ${data.routeVersion}.`,
            `Modelo: ${data.model}.`,
            `Chave configurada: ${data.hasApiKey ? "Sim" : "Não"}.`,
          ].join(" ")
        );
        setText("");
        return;
      }

      if (!data.workbook || !Array.isArray(data.workbook.sheets)) {
        throw new Error(
          "A API respondeu, mas não retornou uma planilha válida."
        );
      }

      const generatedBook =
        data.workbook as WorkbookModel;

      setWarnings(
        Array.isArray(data.warnings)
          ? data.warnings
          : []
      );
      setPendingBook(generatedBook);
      setMessage(
        data.warning ||
          `Estrutura gerada com o modelo ${
            data.model || "configurado"
          }. Confira o entendimento antes de aplicar.`
      );
    } catch (error) {
      console.error("Erro ao gerar planilha:", error);
      const errorMessage =
        error instanceof Error
          ? error.message
          : "Ocorreu um erro desconhecido.";
      setMessage(`Erro: ${errorMessage}`);
    } finally {
      setLoading(false);
    }
  }

  function confirmGeneratedWorkbook() {
    if (!pendingBook) return;

    setHistory((currentHistory) => [
      ...currentHistory,
      book,
    ]);
    setBook(pendingBook);
    setPendingBook(null);
    setActive(0);
    setText("");
    setMessage(
      "Estrutura confirmada. Revise o preview antes de baixar."
    );
  }

  function adjustGeneratedWorkbook() {
    setPendingBook(null);
    setMessage(
      "A estrutura não foi aplicada. Ajuste a solicitação e envie novamente."
    );
  }

  function changeCategory(nextCategory: (typeof categories)[number]) {
    setCategory(nextCategory);
    setBook(template(nextCategory));
    setActive(0);
    setHistory([]);
    setWarnings([]);
    setMessage(
      `Modelo ${nextCategory} carregado. Agora descreva sua necessidade.`
    );
  }

  function undo() {
    const previous = history.at(-1);
    if (!previous) return;

    setBook(previous);
    setHistory((currentHistory) => currentHistory.slice(0, -1));
    setActive(0);
    setMessage("Última alteração desfeita.");
  }

  async function downloadWorkbook() {
    setMessage("Validando fórmulas e preparando o arquivo...");
    try {
      const report = await exportXlsx(book);
      if (!report.downloaded) {
        const details = report.errors
          .slice(0, 3)
          .map((issue) => `${issue.sheet}${issue.cell ? ` ${issue.cell}` : ""}: ${issue.message}`)
          .join(" | ");
        setMessage(`Download bloqueado: ${details || "foram encontrados erros de validação."}`);
        return;
      }
      setMessage(
        `Arquivo gerado com ${report.formulaCount.toLocaleString("pt-BR")} fórmula(s)` +
          (report.warnings.length ? ` e ${report.warnings.length} aviso(s).` : ".")
      );
    } catch (error) {
      setMessage(`Erro ao exportar: ${error instanceof Error ? error.message : "erro desconhecido"}.`);
    }
  }

  const max = Math.max(1, ...chart.map((item) => item.value));

  return (
    <main>
      <header>
        <div className="brand">
          <div className="logo">
            <Table2 />
          </div>
          <div>
            <b>Planilha IA</b>
            <span>Gerador interno para Excel</span>
          </div>
        </div>
        <button className="download" onClick={() => void downloadWorkbook()}>
          <Download size={18} /> Baixar .xlsx
        </button>
      </header>

      <section className="hero">
        <div>
          <span className="eyebrow">
            <Sparkles size={15} /> Planilhas sob medida
          </span>
          <h1>Crie, ajuste e baixe sua planilha.</h1>
          <p>
            Escolha um modelo, converse com a IA e confira o resultado antes de
            abrir no Excel.
          </p>
        </div>
      </section>

      <div className="categoryRow">
        {categories.map((item) => (
          <button
            key={item}
            className={category === item ? "cat active" : "cat"}
            onClick={() => changeCategory(item)}
          >
            {item}
          </button>
        ))}
      </div>

      <section className="workspace">
        <aside className="chat">
          <div className="panelTitle">
            <div>
              <h2>Assistente</h2>
              <p>{message}</p>
            </div>
            <button
              className="iconBtn"
              title="Desfazer"
              disabled={!history.length}
              onClick={undo}
            >
              <RotateCcw size={18} />
            </button>
          </div>

          <div className="suggestions">
            <button
              onClick={() =>
                setText(
                  "Adicione uma aba Dashboard com indicadores e gráfico dos itens mais utilizados."
                )
              }
            >
              <Plus size={14} /> Dashboard
            </button>
            <button
              onClick={() =>
                setText(
                  "Adicione alertas de reposição e destaque itens abaixo do estoque mínimo."
                )
              }
            >
              <Plus size={14} /> Alertas
            </button>
            <button
              onClick={() =>
                setText(
                  "Adicione uma nova coluna de observações em todas as abas relevantes."
                )
              }
            >
              <Plus size={14} /> Coluna
            </button>
          </div>

          <div className="conversation">
            <div className="bubble">
              Exemplo: “Crie um estoque de peças de informática, com entradas,
              saídas, reposição e os itens mais usados.”
            </div>
          </div>

          <div className="composer">
            <textarea
              value={text}
              onChange={(event) => setText(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter" && !event.shiftKey) {
                  event.preventDefault();
                  void generate();
                }
              }}
              placeholder="Descreva a planilha ou a alteração..."
            />
            <button
              type="button"
              onClick={() => void generate()}
              disabled={loading || !text.trim()}
              title="Enviar pedido"
            >
              {loading ? <span className="spinner" /> : <Send size={18} />}
            </button>
          </div>
        </aside>

        <section className="preview">
          <div className="previewHead">
            <div>
              <span>PREVIEW</span>
              <h2>{book.name}</h2>
            </div>
            <small>{book.description}</small>
          </div>

          {warnings.length > 0 && (
            <div className="bubble">
              {warnings.slice(0, 3).join(" | ")}
            </div>
          )}
          <div className="tabs">
            {book.sheets.map((item, index) => (
              <button
                className={index === active ? "tab active" : "tab"}
                key={item.id}
                onClick={() => setActive(index)}
              >
                {item.name}
              </button>
            ))}
          </div>

          <div className="sheetWrap">
            <table>
              <thead>
                <tr>
                  {sheet?.columns.map((column) => (
                    <th key={column.key}>{column.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheet?.rows.length ? (
                  sheet.rows.map((row, index) => (
                    <tr key={index}>
                      {sheet.columns.map((column) => (
                        <td key={column.key}>
                          {row[column.key] === null ||
                          row[column.key] === undefined
                            ? ""
                            : String(row[column.key])}
                        </td>
                      ))}
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td className="empty" colSpan={sheet?.columns.length || 1}>
                      Nenhum dado de exemplo. A estrutura está pronta.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {chart.length > 0 && (
            <div className="chart">
              <h3>Visão rápida</h3>
              {chart.map((item, index) => (
                <div className="barRow" key={index}>
                  <span>{item.label}</span>
                  <div>
                    <i
                      style={{
                        width: `${Math.max(2, (item.value / max) * 100)}%`,
                      }}
                    />
                  </div>
                  <b>{item.value.toLocaleString("pt-BR")}</b>
                </div>
              ))}
            </div>
          )}
        </section>
      </section>

      {pendingBook && pendingSummary && (
        <section className="confirmationPanel">
          <h2>Confira o entendimento</h2>
          <h3>{pendingSummary.name}</h3>
          <p>{pendingSummary.description}</p>

          <p>
            {pendingSummary.sheetCount} abas ·{" "}
            {pendingSummary.totalColumns} colunas ·{" "}
            {pendingSummary.totalPreparedRows} linhas preparadas ·{" "}
            {pendingSummary.totalFormulas} fórmulas
          </p>

          <ul>
            {pendingSummary.sheets.map((item) => (
              <li key={item.id}>
                <b>{item.name}</b>:{" "}
                {item.columnCount} colunas,{" "}
                {item.preparedRows} linhas preparadas
              </li>
            ))}
          </ul>

          {pendingSummary.warnings.length > 0 && (
            <div>
              <b>Avisos:</b>
              <ul>
                {pendingSummary.warnings.map(
                  (warning) => (
                    <li key={warning}>{warning}</li>
                  )
                )}
              </ul>
            </div>
          )}

          <div className="confirmationActions">
            <button
              type="button"
              onClick={adjustGeneratedWorkbook}
            >
              Voltar e ajustar
            </button>

            <button
              type="button"
              onClick={confirmGeneratedWorkbook}
            >
              Confirmar e aplicar
            </button>
          </div>
        </section>
      )}

      <footer>
        Sem macros · Exportação em .xlsx · Revise fórmulas e dados antes do uso
        operacional
      </footer>
    </main>
  );
}
