"use client";
import { useMemo, useState } from "react";
import { Download, Plus, RotateCcw, Send, Sparkles, Table2 } from "lucide-react";
import { template } from "@/lib/templates"; import { WorkbookModel } from "@/lib/types"; import { exportXlsx } from "@/lib/export-xlsx";
const categories=["Estoque","Finanças","Orçamento","Personalizada"] as const;
export default function Home(){
 const [category,setCategory]=useState<(typeof categories)[number]>("Estoque"); const [book,setBook]=useState<WorkbookModel>(()=>template("Estoque"));
 const [active,setActive]=useState(0); const [text,setText]=useState(""); const [loading,setLoading]=useState(false); const [message,setMessage]=useState("Escolha um modelo e descreva o que deseja alterar."); const [history,setHistory]=useState<WorkbookModel[]>([]);
 const sheet=book.sheets[active]||book.sheets[0];
 const chart=useMemo(()=>{if(!sheet)return[]; const label=sheet.columns.find(c=>c.type==="text")?.key||sheet.columns[0]?.key; const num=sheet.columns.find(c=>["number","currency"].includes(c.type))?.key; if(!num)return[]; return sheet.rows.slice(0,8).map(r=>({label:String(r[label]??"Item"),value:Number(r[num]??0)})).filter(x=>Number.isFinite(x.value));},[sheet]);
 async function generate(){if(!text.trim())return; setLoading(true);setMessage("Gerando e validando a planilha...");try{const res=await fetch("/api/generate",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({category,instruction:text,current:book})});const data=await res.json();if(!res.ok)throw new Error(data.error);setHistory(h=>[...h,book]);setBook(data.workbook);setActive(0);setText("");setMessage(data.warning||"Alteração aplicada. Revise o preview antes de baixar.");}catch(e:any){setMessage(e.message||"Erro ao gerar.");}finally{setLoading(false)}}
 function changeCategory(c:any){setCategory(c);setBook(template(c));setActive(0);setHistory([]);setMessage(`Modelo ${c} carregado. Agora descreva sua necessidade.`)}
 function undo(){const prev=history.at(-1);if(prev){setBook(prev);setHistory(h=>h.slice(0,-1));setActive(0);setMessage("Última alteração desfeita.")}}
 const max=Math.max(1,...chart.map(x=>x.value));
 return <main>
  <header><div className="brand"><div className="logo"><Table2/></div><div><b>Planilha IA</b><span>Gerador interno para Excel</span></div></div><button className="download" onClick={()=>exportXlsx(book)}><Download size={18}/> Baixar .xlsx</button></header>
  <section className="hero"><div><span className="eyebrow"><Sparkles size={15}/> Planilhas sob medida</span><h1>Crie, ajuste e baixe sua planilha.</h1><p>Escolha um modelo, converse com a IA e confira o resultado antes de abrir no Excel.</p></div></section>
  <div className="categoryRow">{categories.map(c=><button key={c} className={category===c?"cat active":"cat"} onClick={()=>changeCategory(c)}>{c}</button>)}</div>
  <section className="workspace">
   <aside className="chat"><div className="panelTitle"><div><h2>Assistente</h2><p>{message}</p></div><button className="iconBtn" title="Desfazer" disabled={!history.length} onClick={undo}><RotateCcw size={18}/></button></div>
    <div className="suggestions"><button onClick={()=>setText("Adicione uma aba Dashboard com indicadores e gráfico dos itens mais utilizados.")}><Plus size={14}/> Dashboard</button><button onClick={()=>setText("Adicione alertas de reposição e destaque itens abaixo do estoque mínimo.")}><Plus size={14}/> Alertas</button><button onClick={()=>setText("Adicione uma nova coluna de observações em todas as abas relevantes.")}><Plus size={14}/> Coluna</button></div>
    <div className="conversation"><div className="bubble">Exemplo: “Crie um estoque de peças de informática, com entradas, saídas, reposição e os itens mais usados.”</div></div>
    <div className="composer"><textarea value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();generate()}}} placeholder="Descreva a planilha ou a alteração..."/><button onClick={generate} disabled={loading||!text.trim()}>{loading?<span className="spinner"/>:<Send size={18}/>}</button></div>
   </aside>
   <section className="preview"><div className="previewHead"><div><span>PREVIEW</span><h2>{book.name}</h2></div><small>{book.description}</small></div>
    <div className="tabs">{book.sheets.map((s,i)=><button className={i===active?"tab active":"tab"} key={s.id} onClick={()=>setActive(i)}>{s.name}</button>)}</div>
    <div className="sheetWrap"><table><thead><tr>{sheet?.columns.map(c=><th key={c.key}>{c.label}</th>)}</tr></thead><tbody>{sheet?.rows.length?sheet.rows.map((r,i)=><tr key={i}>{sheet.columns.map(c=><td key={c.key}>{r[c.key]===null||r[c.key]===undefined?"":String(r[c.key])}</td>)}</tr>):<tr><td className="empty" colSpan={sheet?.columns.length}>Nenhum dado de exemplo. A estrutura está pronta.</td></tr>}</tbody></table></div>
    {chart.length>0&&<div className="chart"><h3>Visão rápida</h3>{chart.map((x,i)=><div className="barRow" key={i}><span>{x.label}</span><div><i style={{width:`${Math.max(2,x.value/max*100)}%`}}/></div><b>{x.value.toLocaleString("pt-BR")}</b></div>)}</div>}
   </section>
  </section>
  <footer>Sem macros · Exportação em .xlsx · Revise fórmulas e dados antes do uso operacional</footer>
 </main>
}
