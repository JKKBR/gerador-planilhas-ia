import type { WorkbookModel } from "./types";

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
      const values = sheet.columns.map(
        (column) => sourceRow[column.key] ?? null
      );
      worksheet.addRow(values);
    }

    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);

      for (const formula of sheet.formulas || []) {
        const columnIndex =
          sheet.columns.findIndex(
            (column) => column.key === formula.key
          ) + 1;

        if (columnIndex > 0) {
          row.getCell(columnIndex).value = {
            formula: formula.expression.replaceAll(
              "{row}",
              String(rowNumber)
            ),
          };
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

    if (sheet.columns.length > 0 && worksheet.rowCount > 1) {
      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: {
          row: worksheet.rowCount,
          column: sheet.columns.length,
        },
      };
    }

    const statusColumnIndex =
      sheet.columns.findIndex(
        (column) => column.type === "status"
      ) + 1;

    if (statusColumnIndex > 0 && worksheet.rowCount > 1) {
      const letter = worksheet.getColumn(statusColumnIndex).letter;
      const range = `${letter}2:${letter}${worksheet.rowCount}`;

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
              font: { color: { argb: "FF991B1B" }, bold: true },
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
              font: { color: { argb: "FF92400E" }, bold: true },
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
              font: { color: { argb: "FF166534" }, bold: true },
            },
          } as any,
        ],
      });
    }

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
        cell.alignment = {
          ...cell.alignment,
          vertical: "middle",
        };
      });
    });
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
