import type { WorkbookModel } from "./types";

function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function safeSheetName(name: string) {
  return name.replace(/[\\/*?:[\]]/g, " ").trim().slice(0, 31) || "Planilha";
}

function uniqueSheetName(name: string, used: Set<string>) {
  const base = safeSheetName(name);
  let candidate = base;
  let index = 2;

  while (used.has(candidate.toLowerCase())) {
    const suffix = ` ${index}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }

  used.add(candidate.toLowerCase());
  return candidate;
}

function preparedRowCount(
  workbook: WorkbookModel,
  sheetName: string,
  sheetId: string
) {
  const category = normalizeText(workbook.category);
  const identity = normalizeText(`${sheetName} ${sheetId}`);

  if (identity.includes("moviment")) {
    return 1000;
  }

  if (
    category === "financas" &&
    (identity.includes("lancamento") || identity.includes("transacao"))
  ) {
    return 1000;
  }

  if (
    category === "orcamento" &&
    (identity.includes("item") || identity.includes("produto orcado"))
  ) {
    return 500;
  }

  if (
    identity.includes("cadastro") ||
    identity.includes("produto") ||
    identity.includes("peca") ||
    identity.includes("fornecedor") ||
    identity.includes("cliente")
  ) {
    return 300;
  }

  return 0;
}

function cleanFormula(expression: string, rowNumber: number) {
  return expression
    .trim()
    .replace(/^=/, "")
    .replaceAll("{row}", String(rowNumber));
}

function guardedFormula(
  expression: string,
  rowNumber: number,
  guardColumnLetter: string
) {
  const formula = cleanFormula(expression, rowNumber);
  return `IF($${guardColumnLetter}${rowNumber}="","",${formula})`;
}

function toExcelValue(value: unknown) {
  if (
    typeof value === "string" &&
    value.trim().startsWith("=")
  ) {
    return {
      formula: value.trim().slice(1),
    };
  }

  return value ?? null;
}
export async function exportXlsx(model: WorkbookModel) {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  const usedSheetNames = new Set<string>();

  workbook.creator = "Gerador de Planilhas IA";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  for (const sheet of model.sheets) {
    const worksheet = workbook.addWorksheet(
      uniqueSheetName(sheet.name, usedSheetNames),
      {
        views: [{ state: "frozen", ySplit: 1 }],
      }
    );

    worksheet.columns = sheet.columns.map((column) => ({
      header: column.label,
      key: column.key,
      width:
        column.width ||
        Math.max(12, Math.min(30, column.label.length + 5)),
    }));

    const header = worksheet.getRow(1);
    header.height = 24;
    header.font = {
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1D4ED8" },
    };
    header.alignment = {
      vertical: "middle",
      horizontal: "center",
    };

for (const sourceRow of sheet.rows) {
  worksheet.addRow(
    sheet.columns.map((column) =>
      toExcelValue(sourceRow[column.key])
    )
  );
}

    const requestedPreparedRows = preparedRowCount(
      model,
      sheet.name,
      sheet.id
    );
    const totalDataRows = Math.max(
      sheet.rows.length,
      requestedPreparedRows
    );
    const lastRowNumber = Math.max(1, totalDataRows + 1);

    const formulaKeys = new Set(
      (sheet.formulas || []).map((formula) => formula.key)
    );
    const guardColumnIndex = Math.max(
      1,
      sheet.columns.findIndex(
        (column) => !formulaKeys.has(column.key)
      ) + 1
    );
    const guardColumnLetter = worksheet.getColumn(guardColumnIndex).letter;

    for (let rowNumber = 2; rowNumber <= lastRowNumber; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const isPreparedBlankRow = rowNumber > sheet.rows.length + 1;

      row.height = 20;

      for (
        let columnIndex = 1;
        columnIndex <= sheet.columns.length;
        columnIndex += 1
      ) {
        const cell = row.getCell(columnIndex);

        cell.alignment = {
          vertical: "middle",
          horizontal:
            sheet.columns[columnIndex - 1].type === "text" ||
            sheet.columns[columnIndex - 1].type === "status"
              ? "left"
              : "right",
        };

        cell.border = {
          bottom: {
            style: "hair",
            color: { argb: "FFE2E8F0" },
          },
        };

        if (isPreparedBlankRow && !formulaKeys.has(sheet.columns[columnIndex - 1].key)) {
          cell.font = { color: { argb: "FF0000FF" } };
        }
      }

      for (const formula of sheet.formulas || []) {
        const formulaColumnIndex =
          sheet.columns.findIndex(
            (column) => column.key === formula.key
          ) + 1;

        if (formulaColumnIndex > 0) {
          const formulaCell = row.getCell(formulaColumnIndex);
          formulaCell.value = {
            formula: isPreparedBlankRow
              ? guardedFormula(
                  formula.expression,
                  rowNumber,
                  guardColumnLetter
                )
              : cleanFormula(formula.expression, rowNumber),
          };
          formulaCell.font = { color: { argb: "FF000000" } };
        }
      }

      if (rowNumber % 2 === 0) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF8FAFC" },
        };
      }
    }

    sheet.columns.forEach((column, index) => {
      const excelColumn = worksheet.getColumn(index + 1);

      if (column.type === "currency") {
        excelColumn.numFmt = 'R$ #,##0.00';
      }

      if (column.type === "date") {
        excelColumn.numFmt = "dd/mm/yyyy";
      }

      if (column.type === "number") {
        excelColumn.numFmt = "0.00";
      }
    });

    if (sheet.columns.length > 0 && lastRowNumber > 1) {
      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: {
          row: lastRowNumber,
          column: sheet.columns.length,
        },
      };
    }

    const statusColumnIndex =
      sheet.columns.findIndex(
        (column) => column.type === "status"
      ) + 1;

    if (statusColumnIndex > 0 && lastRowNumber > 1) {
      const letter = worksheet.getColumn(statusColumnIndex).letter;
      const range = `${letter}2:${letter}${lastRowNumber}`;

      worksheet.addConditionalFormatting({
        ref: range,
        rules: [
          {
            type: "containsText",
            operator: "containsText",
            text: "ESGOTADO",
            priority: 1,
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFECACA" },
              },
              font: {
                color: { argb: "FF991B1B" },
                bold: true,
              },
            },
          } as any,
          {
            type: "containsText",
            operator: "containsText",
            text: "REPOR",
            priority: 2,
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFEF3C7" },
              },
              font: {
                color: { argb: "FF92400E" },
                bold: true,
              },
            },
          } as any,
          {
            type: "containsText",
            operator: "containsText",
            text: "OK",
            priority: 3,
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFDCFCE7" },
              },
              font: {
                color: { argb: "FF166534" },
                bold: true,
              },
            },
          } as any,
        ],
      });
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${model.name.replace(
    /[^a-zA-Z0-9À-ÿ_-]+/g,
    "_"
  )}.xlsx`;

  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}
