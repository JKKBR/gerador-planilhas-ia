const fs = require("node:fs");
const path = require("node:path");

const file = path.resolve(process.cwd(), "lib/export-xlsx.ts");
if (!fs.existsSync(file)) {
  console.error("Arquivo não encontrado:", file);
  process.exit(1);
}

let source = fs.readFileSync(file, "utf8");
const oldBlock = `function addInstructionsSheet(workbook: any, rows: FormulaHelpRow[]) {
  if (!rows.length) return;
  const worksheet = workbook.addWorksheet("Instruções e Fórmulas", {
    views: [{ state: "frozen", ySplit: 5 }],
  });`;

const newBlock = `function addInstructionsSheet(workbook: any, rows: FormulaHelpRow[]) {
  if (!rows.length) return;

  const instructionsSheetName = "Instruções e Fórmulas";
  const existingInstructionsSheet = workbook.getWorksheet(instructionsSheetName);

  // A IA pode ter criado essa aba. Remove a versão anterior para que o
  // exportador gere uma única aba padronizada, sem conflito de nomes.
  if (existingInstructionsSheet) {
    workbook.removeWorksheet(existingInstructionsSheet.id);
  }

  const worksheet = workbook.addWorksheet(instructionsSheetName, {
    views: [{ state: "frozen", ySplit: 5 }],
  });`;

if (source.includes(newBlock)) {
  console.log("A correção já está aplicada.");
  process.exit(0);
}

if (!source.includes(oldBlock)) {
  console.error("O bloco esperado não foi encontrado. Faça a substituição manual descrita no README.");
  process.exit(1);
}

source = source.replace(oldBlock, newBlock);
fs.writeFileSync(file, source, "utf8");
console.log("Correção aplicada em lib/export-xlsx.ts");
