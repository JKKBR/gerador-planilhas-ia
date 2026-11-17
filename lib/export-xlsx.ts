import { WorkbookModel } from "./types";

function safeName(name: string) {
  return (
    name
      .replace(/[\\/*?:[\]]/g, " ")
      .trim()
      .slice(0, 31) || "Planilha"
  );
}

export async function exportXlsx(model: WorkbookModel) {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Gerador de Planilhas IA";
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  for (const sheet of model.sheets) {
    const worksheet = workbook.addWorksheet(safeName(sheet.name), {
      views: [
        {
          state: "frozen",
          ySplit: 1,
        },
      ],
    });

    worksheet.columns = sheet.columns.map((column) => ({
      header: column.label,
      key: column.key,
      width:
        column.width ||
        Math.max(12, Math.min(28, column.label.length + 5)),
    }));

    const header = worksheet.getRow(1);

    header.font = {
      bold: true,
      color: {
        argb: "FFFFFFFF",
      },
    };

    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: "FF1D4ED8",
      },
    };

    header.alignment = {
      vertical: "middle",
    };

    header.height = 24;

    if (sheet.rows.length > 0) {
      worksheet.addTable({
        name: `Tabela_${sheet.id}`
          .replace(/[^A-Za-z0-9_]/g, "_")
          .slice(0, 30),
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        style: {
          theme: "TableStyleMedium2",
          showRowStripes: true,
        },
        columns: sheet.columns.map((column) => ({
          name: column.label,
        })),
        rows: sheet.rows.map((row) =>
          sheet.columns.map((column) => row[column.key] ?? null)
        ),
      });
    }

    sheet.rows.forEach((_, index) => {
      const excelRow = index + 2;
      const row = worksheet.getRow(excelRow);

      for (const formula of sheet.formulas || []) {
        const columnIndex =
          sheet.columns.findIndex(
            (column) => column.key === formula.key
          ) + 1;

        if (columnIndex > 0) {
          row.getCell(columnIndex).value = {
            formula: formula.expression.replaceAll(
              "{row}",
              String(excelRow)
            ),
          };
        }
      }
    });

    sheet.columns.forEach((column, index) => {
      const excelColumn = worksheet.getColumn(index + 1);

      if (column.type === "currency") {
        excelColumn.numFmt = 'R$ #,##0.00';
      }

      if (column.type === "date") {
        excelColumn.numFmt = "dd/mm/yyyy";
      }
    });

    worksheet.autoFilter = {
      from: {
        row: 1,
        column: 1,
      },
      to: {
        row: Math.max(1, worksheet.rowCount),
        column: sheet.columns.length,
      },
    };

    const statusColumnIndex =
      sheet.columns.findIndex(
        (column) => column.type === "status"
      ) + 1;

    if (statusColumnIndex > 0 && worksheet.rowCount > 1) {
      const columnLetter =
        worksheet.getColumn(statusColumnIndex).letter;

      worksheet.addConditionalFormatting({
        ref: `${columnLetter}2:${columnLetter}${worksheet.rowCount}`,
        rules: [
          {
            type: "containsText",
            operator: "containsText",
            text: "REPOR",
            priority: 1,
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                bgColor: {
                  argb: "FFFFC7CE",
                },
                fgColor: {
                  argb: "FFFFC7CE",
                },
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
