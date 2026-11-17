import { WorkbookModel } from "./types";
function safeName(s:string){return s.replace(/[\\/*?:\[\]]/g," ").slice(0,31)||"Planilha"}
export async function exportXlsx(model:WorkbookModel){
 const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook();
 wb.creator="Gerador de Planilhas IA"; wb.created=new Date(); wb.calcProperties.fullCalcOnLoad=true;
 for(const s of model.sheets){
  const ws=wb.addWorksheet(safeName(s.name),{views:[{state:"frozen",ySplit:1}]});
  ws.columns=s.columns.map(c=>({header:c.label,key:c.key,width:c.width||Math.max(12,Math.min(28,c.label.length+5))}));
  const header=ws.getRow(1); header.font={bold:true,color:{argb:"FFFFFFFF"}}; header.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1D4ED8"}}; header.alignment={vertical:"middle"}; header.height=24;
  if(s.rows.length){ws.addTable({name:("Tabela_"+s.id).replace(/[^A-Za-z0-9_]/g,"_").slice(0,30),ref:"A1",headerRow:true,totalsRow:false,style:{theme:"TableStyleMedium2",showRowStripes:true},columns:s.columns.map(c=>({name:c.label})),rows:s.rows.map(r=>s.columns.map(c=>r[c.key]??null))});}
  s.rows.forEach((_,i)=>{ const row=ws.getRow(i+2); const excelRow=i+2; for(const f of s.formulas||[]){ const col=s.columns.findIndex(c=>c.key===f.key)+1; if(col>0) row.getCell(col).value={formula:f.expression.replaceAll("{row}",String(excelRow))}; }});
  s.columns.forEach((c,i)=>{const col=ws.getColumn(i+1); if(c.type==="currency") col.numFmt='R$ #,##0.00'; if(c.type==="date") col.numFmt='dd/mm/yyyy';});
  ws.autoFilter={from:{row:1,column:1},to:{row:Math.max(1,ws.rowCount),column:s.columns.length}};
  const statusCol=s.columns.findIndex(c=>c.type==="status")+1;
  if(statusCol>0 && ws.rowCount>1){ const letter=ws.getColumn(statusCol).letter; ws.addConditionalFormatting({ref:`${letter}2:${letter}${ws.rowCount}`,rules:[{type:"containsText",operator:"containsText",text:"REPOR",priority:1,style:{fill:{type:"pattern",pattern:"solid",bgColor:{argb:"FFFFC7CE"},fgColor:{argb:"FFFFC7CE"}}} as any]}); }
 }
 const buffer=await wb.xlsx.writeBuffer(); const blob=new Blob([buffer as BlobPart],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
 const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`${model.name.replace(/[^a-zA-Z0-9À-ÿ_-]+/g,"_")}.xlsx`; a.click(); URL.revokeObjectURL(a.href);
}
