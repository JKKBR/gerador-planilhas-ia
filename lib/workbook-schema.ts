import { z } from "zod";

import { isReservedSheetName } from "./formula-utils";

export const categorySchema = z.enum([
  "Estoque",
  "Finanças",
  "Orçamento",
  "Personalizada",
]);

export const cellValueSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
]);

export const columnSchema = z.object({
  key: z.string().trim().min(1).max(80),
  label: z.string().trim().min(1).max(120),
  type: z.enum([
    "text",
    "number",
    "currency",
    "date",
    "status",
  ]),
  width: z.number().min(6).max(80).optional(),
  inputMode: z
    .enum(["manual", "calculated"])
    .optional(),
  required: z.boolean().optional(),
});

export const fillFormulaSchema = z.object({
  mode: z.literal("fill").optional(),
  key: z.string().trim().min(1),
  expression: z.string().trim().min(1),
  fromRow: z.number().int().min(2).max(5001).optional(),
  toRow: z.number().int().min(2).max(5001).optional(),
});

export const pointFormulaSchema = z.object({
  mode: z.enum(["single", "dynamic"]),
  targetCell: z
    .string()
    .trim()
    .regex(
      /^[A-Z]{1,3}[1-9][0-9]*$/i,
      "Célula de destino inválida."
    ),
  expression: z.string().trim().min(1),
  clearRange: z
    .string()
    .trim()
    .regex(
      /^[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*$/i,
      "Intervalo de limpeza inválido."
    )
    .optional(),
});

export const formulaSchema = z.union([
  fillFormulaSchema,
  pointFormulaSchema,
]);

export const sheetSchema = z
  .object({
    id: z.string().trim().min(1).max(80),
    name: z.string().trim().min(1).max(31),
    purpose: z.string().trim().max(300).optional(),
    preparedRows: z
      .number()
      .int()
      .min(0)
      .max(5000)
      .optional(),
    columns: z.array(columnSchema).min(1).max(20),
    rows: z
      .array(z.record(z.string(), cellValueSchema))
      .max(15),
    formulas: z.array(formulaSchema).max(100).optional(),
  })
  .superRefine((sheet, context) => {
    if (isReservedSheetName(sheet.name)) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["name"],
        message:
          'A aba "Instruções e Fórmulas" é reservada ao exportador.',
      });
    }

    const columnKeys = new Set<string>();

    sheet.columns.forEach((column, index) => {
      if (columnKeys.has(column.key)) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["columns", index, "key"],
          message: `A chave "${column.key}" está duplicada.`,
        });
      }

      columnKeys.add(column.key);
    });

    for (const formula of sheet.formulas ?? []) {
      const mode = formula.mode ?? "fill";

      if (mode === "fill" && "key" in formula) {
        if (!columnKeys.has(formula.key)) {
          context.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["formulas"],
            message:
              `A fórmula referencia a coluna inexistente "${formula.key}".`,
          });
        }

        const fromRow = formula.fromRow ?? 2;
        const toRow = formula.toRow ?? 2;

        if (toRow < fromRow) {
          context.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["formulas"],
            message:
              "A linha final da fórmula não pode ser menor que a inicial.",
          });
        }
      }

      if (
        mode === "dynamic" &&
        (
          !("clearRange" in formula) ||
          !formula.clearRange
        )
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["formulas"],
          message:
            "Fórmulas dinâmicas devem informar clearRange.",
        });
      }
    }
  });

export const workbookSchema = z
  .object({
    name: z.string().trim().min(1).max(120),
    category: categorySchema,
    description: z.string().trim().min(1).max(500),
    sheets: z.array(sheetSchema).min(1).max(8),
  })
  .superRefine((workbook, context) => {
    const sheetNames = new Set<string>();
    const sheetIds = new Set<string>();

    workbook.sheets.forEach((sheet, index) => {
      const normalizedName = sheet.name
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toLowerCase();

      if (sheetNames.has(normalizedName)) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["sheets", index, "name"],
          message: `A aba "${sheet.name}" está duplicada.`,
        });
      }

      if (sheetIds.has(sheet.id)) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["sheets", index, "id"],
          message: `O ID "${sheet.id}" está duplicado.`,
        });
      }

      sheetNames.add(normalizedName);
      sheetIds.add(sheet.id);
    });
  });

export type ValidatedWorkbook = z.infer<
  typeof workbookSchema
>;

export function validateWorkbook(input: unknown) {
  return workbookSchema.safeParse(input);
}

export function formatWorkbookValidationErrors(
  error: z.ZodError
): string[] {
  return error.issues.map((issue) => {
    const path =
      issue.path.length > 0
        ? issue.path.join(".")
        : "workbook";

    return `${path}: ${issue.message}`;
  });
}
