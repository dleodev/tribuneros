import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from "recharts";
import { LayoutGrid, ShoppingBag, TrendingDown, Package, Users, Upload, Settings, Search, Truck, ArrowUpRight, ArrowDownRight, Plus, Minus, X, Check, Trash2, Pencil, FileSpreadsheet, Download, Ticket, ChevronDown, ChevronRight, Boxes, RefreshCw, Trophy, Tag, ShoppingCart, Eye, History, Database, BarChart3, Megaphone, MapPin, AlertTriangle, Wallet, PackageCheck } from "lucide-react";

// ---------- TEMA ----------
const C = {
  bg:"#F5F5FA", card:"#FFFFFF", ink:"#1A1A2E", soft:"#6B6B80", faint:"#A0A0B4",
  violet:"#7C5CFC", violetDeep:"#5B3FD6", violetSoft:"#EEE9FF", violetSoft2:"#F6F3FF",
  green:"#27AE60", greenSoft:"#E7F8EE", orange:"#F2994A", orangeSoft:"#FDF1E7",
  red:"#EB5757", redSoft:"#FCEBEB", blue:"#5AA9E6", blueSoft:"#EAF4FC",
  border:"#ECECF3", lav:"#B9A6FF",
};
const FONT = "'Inter', system-ui, -apple-system, sans-serif";
const PIE = [C.violet, C.lav, C.orange, C.blue, C.green];

const ars = n => "$" + Math.round(Number(n)||0).toLocaleString("es-AR");
const today = () => new Date().toISOString().split("T")[0];
const stamp = () => new Date().toISOString().slice(0,16).replace("T"," ");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,5);
const r2 = n => Math.round(n*100)/100;
const norm = s => (s||"").trim().toLowerCase();

const DEFAULT_RATES = { mp:5.57, cpt:2.0, iibb:1.5, c3:14.75 };

const CANALES = ["Tiendanube","Offline","Instagram","WhatsApp"];
const METODOS = ["Mercado Pago","Tarjeta crédito","Transferencia","Efectivo"];
const CUOTAS = ["1","2 sin interés","3 sin interés"];
const TALLES = ["—","XS","S","M","L","XL","XXL"];
const ENVIO_QUIEN = ["Vos (absorbés)","Cliente","Sin envío","A coordinar"];

function computeVenta(v, rates) {
  const isMP = v.metodo==="Mercado Pago" || v.metodo==="Tarjeta crédito";
  const isTN = v.canal==="Tiendanube";
  const is3c = v.cuotas==="3 sin interés";
  const ret = v.retencion!=null && v.retencion!=="" ? +v.retencion : (isMP ? r2(v.bruto*rates.iibb/100) : 0);
  const mp  = v.cargo_mp!=null && v.cargo_mp!=="" ? +v.cargo_mp : (isMP ? r2(v.bruto*rates.mp/100) : 0);
  const cpt = v.cargo_plataforma!=null && v.cargo_plataforma!=="" ? +v.cargo_plataforma : (isTN ? r2(v.bruto*rates.cpt/100) : 0);
  const i3  = v.interes3c!=null && v.interes3c!=="" ? +v.interes3c : (is3c ? r2(v.bruto*rates.c3/100) : 0);
  const env = +(v.costo_envio||0);
  const absorbe = v.quien_envio !== "Cliente";
  const cobrar = (+v.bruto||0) - ret - mp - cpt - i3;
  const final = cobrar - (absorbe ? env : 0);
  return { ...v, retencion:ret, cargo_mp:mp, cargo_plataforma:cpt, interes3c:i3, costo_envio:env, neto_cobrar:r2(cobrar), neto_final:r2(final) };
}

const migrateVenta = v => {
  if (Array.isArray(v.items) && v.items.length) return v;
  return { ...v, items: [{ producto: v.producto||"", talle: v.talle||"—", cant: 1 }] };
};
const itemsLabel = v => {
  const its = v.items && v.items.length ? v.items : [{ producto:v.producto, talle:v.talle, cant:1 }];
  return its.map(i=>`${i.producto||"—"}${i.talle&&i.talle!=="—"?" ("+i.talle+")":""}${(+i.cant||1)>1?" ×"+i.cant:""}`).join("  ·  ");
};

const mergeStock = (list) => {
  const out = []; const idx = {};
  list.forEach(s=>{
    const k = norm(s.producto)+"|"+(s.talle||"—");
    if(idx[k]!=null){ const t=out[idx[k]]; t.ingresadas=(+t.ingresadas||0)+(+s.ingresadas||0); t.vendidas=(+t.vendidas||0)+(+s.vendidas||0); if(!t.costo&&s.costo)t.costo=s.costo; if(!t.notas&&s.notas)t.notas=s.notas; }
    else { idx[k]=out.length; out.push({...s}); }
  });
  return out;
};

function buildClientes(ventas, manual) {
  const map = {};
  const F = ["dni","email","tel","direccion","ciudad","cp","provincia"];
  ventas.forEach(v=>{
    if(!v.cliente || v.cliente.startsWith("(")) return;
    const key = norm(v.cliente);
    const c = map[key] || { key, nombre:v.cliente, dni:"", email:"", tel:"", direccion:"", ciudad:"", cp:"", provincia:"", notas:"", compras:0, total:0, neto:0, ultima:"" };
    c.compras++; c.total += +v.bruto||0; c.neto += +v.neto_final||0;
    if(!c.ultima || (v.fecha && v.fecha>c.ultima)) c.ultima = v.fecha||c.ultima;
    F.forEach(f=>{ if(!c[f] && v[f]) c[f]=v[f]; });
    map[key]=c;
  });
  (manual||[]).forEach(m=>{
    const key = m.key || norm(m.nombre);
    const c = map[key] || { key, nombre:m.nombre, compras:0, total:0, neto:0, ultima:"" };
    ["nombre","dni","email","tel","direccion","ciudad","cp","provincia","notas"].forEach(f=>{ if(m[f]!==undefined && m[f]!=="") c[f]=m[f]; });
    map[key]=c;
  });
  return Object.values(map).sort((a,b)=>b.neto-a.neto);
}

// Sin datos de ejemplo a propósito: una base nueva arranca vacía, con
// estados vacíos en cada pantalla. Los datos reales se cargan por afuera
// (ver Importar/Exportar) o a mano desde cada módulo.

const load = async (k, def) => { try { const r = await window.storage.get(k); return r?.value ? JSON.parse(r.value) : def; } catch { return def; } };
const save = (k, d) => { try { return window.storage.set(k, JSON.stringify(d)); } catch {} };
const downloadXLSX = (wb, name) => { const out=XLSX.write(wb,{bookType:"xlsx",type:"array"}); const b=new Blob([out],{type:"application/octet-stream"}); const u=URL.createObjectURL(b); const a=document.createElement("a"); a.href=u; a.download=name; a.click(); URL.revokeObjectURL(u); };

const CANAL_CLR = { Tiendanube:C.violet, Offline:C.ink, Instagram:C.orange, WhatsApp:C.green };

// ---------- UI base ----------
function Badge({ children, color, bg }) { return <span style={{ background:bg, color, fontSize:11, fontWeight:600, padding:"3px 9px", borderRadius:8, whiteSpace:"nowrap" }}>{children}</span>; }
function Card({ children, style }) { return <div style={{ background:C.card, borderRadius:18, boxShadow:"0 1px 3px rgba(20,20,50,0.04), 0 8px 24px rgba(20,20,50,0.04)", border:`1px solid ${C.border}`, ...style }}>{children}</div>; }
function KPI({ icon:Icon, label, value, sub, color, bg }) {
  return (
    <Card style={{ padding:"18px 20px" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14 }}>
        <div style={{ width:40, height:40, borderRadius:12, background:bg, display:"flex", alignItems:"center", justifyContent:"center" }}><Icon size={20} color={color} /></div>
      </div>
      <div style={{ fontSize:24, fontWeight:800, color:C.ink, letterSpacing:-0.5 }}>{value}</div>
      <div style={{ fontSize:13, color:C.soft, marginTop:2 }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:C.faint, marginTop:2 }}>{sub}</div>}
    </Card>
  );
}
const inputSt = { width:"100%", boxSizing:"border-box", background:"#FAFAFD", border:`1px solid ${C.border}`, borderRadius:10, color:C.ink, fontSize:13.5, padding:"9px 12px", fontFamily:FONT, outline:"none" };
const labelSt = { fontSize:11, color:C.soft, fontWeight:600, display:"block", marginBottom:5 };
function Field({ label, children, span }) { return <div style={span?{gridColumn:"1/-1"}:{}}><label style={labelSt}>{label||"\u00A0"}</label>{children}</div>; }
function Inp({ label, span, ...p }) { return <Field label={label} span={span}><input style={inputSt} {...p}/></Field>; }
function Sel({ label, options, span, ...p }) { return <Field label={label} span={span}><select style={inputSt} {...p}>{options.map(o=> typeof o==="string"?<option key={o}>{o}</option>:<option key={o.v} value={o.v}>{o.l}</option>)}</select></Field>; }
function Btn({ children, onClick, kind="primary", style, disabled, ...p }) {
  const v = { primary:{ background:C.violet, color:"#fff" }, ghost:{ background:"#fff", border:`1px solid ${C.border}`, color:C.soft }, soft:{ background:C.violetSoft, color:C.violetDeep }, danger:{ background:C.redSoft, color:C.red } }[kind];
  return <button onClick={onClick} disabled={disabled} {...p} style={{ borderRadius:10, fontSize:13, fontWeight:600, padding:"9px 16px", cursor:disabled?"default":"pointer", opacity:disabled?0.55:1, border:"none", fontFamily:FONT, display:"inline-flex", alignItems:"center", gap:6, ...v, ...style }}>{children}</button>;
}
function Switch({ on, onClick }) {
  return (
    <button onClick={onClick} style={{ display:"inline-flex", alignItems:"center", gap:7, background:on?C.greenSoft:C.orangeSoft, border:"none", borderRadius:20, padding:"4px 10px 4px 4px", cursor:"pointer", fontFamily:FONT }}>
      <span style={{ width:18, height:18, borderRadius:"50%", background:on?C.green:C.orange, display:"flex", alignItems:"center", justifyContent:"center" }}>{on ? <Truck size={11} color="#fff"/> : <Package size={11} color="#fff"/>}</span>
      <span style={{ fontSize:12, fontWeight:600, color:on?C.green:C.orange }}>{on?"Despachado":"Pendiente"}</span>
    </button>
  );
}
function Modal({ title, sub, onClose, children, wide }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:200, background:"rgba(26,26,46,0.45)", backdropFilter:"blur(4px)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <Card style={{ width:"100%", maxWidth:wide?720:520, maxHeight:"92vh", overflowY:"auto", padding:24 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:18 }}>
          <div><div style={{ fontSize:19, fontWeight:800, color:C.ink }}>{title}</div>{sub && <div style={{ fontSize:13, color:C.soft, marginTop:2 }}>{sub}</div>}</div>
          <button onClick={onClose} style={{ background:"#F2F2F7", border:"none", borderRadius:10, width:32, height:32, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}><X size={16} color={C.soft}/></button>
        </div>
        {children}
      </Card>
    </div>
  );
}
function Toast({ t }) { if(!t) return null; const err=t.type==="err"; return <div style={{ position:"fixed", top:16, right:16, zIndex:999, background:"#fff", borderLeft:`4px solid ${err?C.red:C.green}`, boxShadow:"0 8px 30px rgba(0,0,0,0.12)", borderRadius:12, padding:"12px 18px", fontSize:13, fontWeight:600, color:C.ink, fontFamily:FONT }}>{t.msg}</div>; }
function Title({ children }) { return <h2 style={{ fontSize:22, fontWeight:800, color:C.ink, margin:0 }}>{children}</h2>; }
function Sec({ children }) { return <div style={{ fontSize:11, fontWeight:700, color:C.violet, textTransform:"uppercase", letterSpacing:1, margin:"16px 0 10px" }}>{children}</div>; }
function IconBtn({ icon:Icon, onClick, color, bg, title }) {
  return <button title={title} onClick={onClick} style={{ width:28, height:28, borderRadius:8, border:`1px solid ${C.border}`, background:bg||"#fff", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}><Icon size={14} color={color||C.soft}/></button>;
}

// ====== VENTAS FORM ======
function VentaForm({ initial, rates, onSave, onClose, notify }) {
  const blank = { pedido:"", fecha:today(), cliente:"", dni:"", email:"", tel:"", direccion:"", ciudad:"", cp:"", provincia:"", items:[{producto:"",talle:"L",cant:1}], canal:"Tiendanube", metodo:"Mercado Pago", cuotas:"1", bruto:"", retencion:null, cargo_mp:null, cargo_plataforma:null, interes3c:null, costo_envio:0, quien_envio:"Vos (absorbés)", desp:false, acred:"", cupon:"", coment:"" };
  const start = { ...blank, ...(initial||{}) };
  if (!Array.isArray(start.items) || !start.items.length) start.items = [{producto:start.producto||"",talle:start.talle||"L",cant:1}];
  const [f, setF] = useState(start);
  const busy = useRef(false);
  const sf = (k,v) => setF(p=>({ ...p, [k]:v }));
  const recalc = (next) => { const t = computeVenta({ ...f, ...next, retencion:null, cargo_mp:null, cargo_plataforma:null, interes3c:null }, rates); setF(p=>({ ...p, ...next, retencion:t.retencion, cargo_mp:t.cargo_mp, cargo_plataforma:t.cargo_plataforma, interes3c:t.interes3c })); };
  const setItem = (i,k,v) => setF(p=>({ ...p, items:p.items.map((it,idx)=>idx===i?{...it,[k]:v}:it) }));
  const addItem = () => setF(p=>({ ...p, items:[...p.items, {producto:"",talle:"L",cant:1}] }));
  const delItem = (i) => setF(p=>({ ...p, items:p.items.length>1?p.items.filter((_,idx)=>idx!==i):p.items }));
  const calc = computeVenta(f, rates);

  const submit = () => {
    if (busy.current) return;
    if(!f.cliente || !f.bruto){ notify&&notify("Completá cliente y monto pagado","err"); return; }
    if(!f.items.some(i=>i.producto.trim())){ notify&&notify("Cargá al menos un producto","err"); return; }
    busy.current = true;
    const clean = { ...f, items:f.items.filter(i=>i.producto.trim()).map(i=>({...i, cant:+i.cant||1})) };
    onSave(computeVenta(clean, rates));
  };

  return (
    <Modal wide title={f.id?`Pedido #${f.pedido||"—"}`:"Nueva venta"} sub="Cliente · varios productos · cobro real · cupón · despacho" onClose={onClose}>
      <Sec>Pedido</Sec>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:6 }}>
        <Inp label="N° Pedido" value={f.pedido} onChange={e=>sf("pedido",e.target.value)} placeholder="100 (Tiendanube)" />
        <Inp label="Fecha" type="date" value={f.fecha} onChange={e=>sf("fecha",e.target.value)} />
        <Sel label="Canal" value={f.canal} options={CANALES} onChange={e=>recalc({canal:e.target.value})} />
      </div>
      <div style={{ fontSize:11, color:C.faint, marginBottom:10 }}>Tiendanube numera desde 100. Para ventas offline/WhatsApp usá un código propio (ej: <b>OFF-1</b>).</div>

      <Sec>Cliente</Sec>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        <Inp label="Nombre completo *" value={f.cliente} onChange={e=>sf("cliente",e.target.value)} />
        <Inp label="DNI / CUIT" value={f.dni} onChange={e=>sf("dni",e.target.value)} />
        <Inp label="Email" value={f.email} onChange={e=>sf("email",e.target.value)} />
        <Inp label="Teléfono" value={f.tel} onChange={e=>sf("tel",e.target.value)} />
        <Inp label="Dirección" value={f.direccion} onChange={e=>sf("direccion",e.target.value)} />
        <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr", gap:8 }}>
          <Inp label="Ciudad" value={f.ciudad} onChange={e=>sf("ciudad",e.target.value)} />
          <Inp label="CP" value={f.cp} onChange={e=>sf("cp",e.target.value)} />
          <Inp label="Prov." value={f.provincia} onChange={e=>sf("provincia",e.target.value)} />
        </div>
      </div>

      <Sec>Productos (podés cargar varios)</Sec>
      {f.items.map((it,i)=>(
        <div key={i} style={{ display:"grid", gridTemplateColumns:"3fr 1fr 0.8fr auto", gap:8, marginBottom:8, alignItems:"end" }}>
          <Field label={i===0?"Producto / camiseta *":""}><input style={inputSt} value={it.producto} onChange={e=>setItem(i,"producto",e.target.value)} placeholder="Ej: Argentina 2026 Titular" /></Field>
          <Field label={i===0?"Talle":""}><select style={inputSt} value={it.talle} onChange={e=>setItem(i,"talle",e.target.value)}>{TALLES.map(t=><option key={t}>{t}</option>)}</select></Field>
          <Field label={i===0?"Cant":""}><input style={inputSt} type="number" min="1" value={it.cant} onChange={e=>setItem(i,"cant",e.target.value)} /></Field>
          <div style={{ paddingBottom:1 }}><IconBtn icon={Trash2} onClick={()=>delItem(i)} color={C.red} bg={C.redSoft}/></div>
        </div>
      ))}
      <Btn kind="soft" style={{ padding:"7px 13px", fontSize:12.5 }} onClick={addItem}><Plus size={14}/>Agregar producto</Btn>

      <Sec>Cobro, comisiones y cupón</Sec>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12 }}>
        <Sel label="Método de pago" value={f.metodo} options={METODOS} onChange={e=>recalc({metodo:e.target.value})} />
        <Sel label="Cuotas" value={f.cuotas} options={CUOTAS} onChange={e=>recalc({cuotas:e.target.value})} />
        <Inp label="Pagado por el cliente *" type="number" value={f.bruto} onChange={e=>recalc({bruto:e.target.value})} placeholder="0" />
        <Inp label="Retención IIBB (SIRTAC)" type="number" value={f.retencion??""} onChange={e=>sf("retencion",e.target.value)} />
        <Inp label="Interés 3 cuotas" type="number" value={f.interes3c??""} onChange={e=>sf("interes3c",e.target.value)} />
        <Inp label="Cargo Mercado Pago" type="number" value={f.cargo_mp??""} onChange={e=>sf("cargo_mp",e.target.value)} />
        <Inp label="Cargo plataforma (CPT)" type="number" value={f.cargo_plataforma??""} onChange={e=>sf("cargo_plataforma",e.target.value)} />
        <Inp label="Cupón usado (código)" value={f.cupon} onChange={e=>sf("cupon",e.target.value)} placeholder="MUNDIAL2026 / vacío" />
        <Inp label="Fecha de acreditación" type="date" value={f.acred} onChange={e=>sf("acred",e.target.value)} />
      </div>

      <Sec>Envío y despacho</Sec>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, alignItems:"end" }}>
        <Inp label="Costo de envío" type="number" value={f.costo_envio} onChange={e=>sf("costo_envio",e.target.value)} />
        <Sel label="Quién paga el envío" value={f.quien_envio} options={ENVIO_QUIEN} onChange={e=>sf("quien_envio",e.target.value)} />
        <div><label style={labelSt}>Estado de despacho</label><Switch on={f.desp} onClick={()=>sf("desp",!f.desp)} /></div>
      </div>

      {f.bruto>0 && (
        <div style={{ marginTop:18, background:C.violetSoft2, borderRadius:14, padding:16 }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr auto", gap:"5px 16px", fontSize:13 }}>
            <span style={{ color:C.soft }}>Pagó el cliente</span><b style={{ textAlign:"right" }}>{ars(f.bruto)}</b>
            {calc.retencion>0 && <><span style={{ color:C.soft }}>− Retención IIBB</span><span style={{ textAlign:"right", color:C.red }}>−{ars(calc.retencion)}</span></>}
            {calc.interes3c>0 && <><span style={{ color:C.soft }}>− Interés 3 cuotas</span><span style={{ textAlign:"right", color:C.red }}>−{ars(calc.interes3c)}</span></>}
            {calc.cargo_mp>0 && <><span style={{ color:C.soft }}>− Mercado Pago</span><span style={{ textAlign:"right", color:C.red }}>−{ars(calc.cargo_mp)}</span></>}
            {calc.cargo_plataforma>0 && <><span style={{ color:C.soft }}>− Plataforma (CPT)</span><span style={{ textAlign:"right", color:C.red }}>−{ars(calc.cargo_plataforma)}</span></>}
            <span style={{ color:C.violetDeep, fontWeight:700, borderTop:`1px solid ${C.border}`, paddingTop:6 }}>Total a recibir (sin envío)</span>
            <b style={{ textAlign:"right", color:C.violetDeep, borderTop:`1px solid ${C.border}`, paddingTop:6 }}>{ars(calc.neto_cobrar)}</b>
            {calc.quien_envio!=="Cliente" && calc.costo_envio>0 && <><span style={{ color:C.soft }}>− Envío que absorbés</span><span style={{ textAlign:"right", color:C.red }}>−{ars(calc.costo_envio)}</span></>}
            <span style={{ color:C.green, fontWeight:800, borderTop:`2px solid ${C.green}`, paddingTop:6 }}>GANANCIA NETA (con envío)</span>
            <b style={{ textAlign:"right", color:C.green, fontSize:17, borderTop:`2px solid ${C.green}`, paddingTop:6 }}>{ars(calc.neto_final)}</b>
          </div>
        </div>
      )}

      <div style={{ display:"flex", gap:10, marginTop:20 }}>
        <Btn style={{ flex:1, justifyContent:"center" }} onClick={submit}>Guardar venta</Btn>
        <Btn kind="ghost" onClick={onClose}>Cancelar</Btn>
      </div>
    </Modal>
  );
}

// ====== MÓDULO VENTAS ======
function Ventas({ ventas, setVentas, rates, notify, query }) {
  const [form, setForm] = useState(null);
  const [filtro, setFiltro] = useState("Todos");
  const [del, setDel] = useState(null);

  const guardar = async (v) => { const vv={...v,id:v.id||uid()}; const list=v.id?ventas.map(x=>x.id===v.id?vv:x):[vv,...ventas]; setVentas(list); await save("tbx_ventas",list); setForm(null); notify(v.id?"Venta actualizada":"Venta registrada"); };
  const eliminar = async (id) => { const l=ventas.filter(v=>v.id!==id); setVentas(l); await save("tbx_ventas",l); setDel(null); notify("Eliminada"); };
  const toggleDesp = async (id) => { const l=ventas.map(v=>v.id===id?{...v,desp:!v.desp}:v); setVentas(l); await save("tbx_ventas",l); };

  const q = (query||"").toLowerCase().trim();
  let lista = filtro==="Todos"?ventas: filtro==="Pendientes de despacho"?ventas.filter(v=>!v.desp): ventas.filter(v=>v.canal===filtro);
  if(q) lista = lista.filter(v => [v.cliente,itemsLabel(v),v.pedido,v.tel,v.cupon,v.canal].some(x=>String(x||"").toLowerCase().includes(q)));

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18 }}>
        <Title>Ventas {q && <span style={{ fontSize:14, color:C.faint, fontWeight:500 }}>· buscando "{query}"</span>}</Title>
        <Btn onClick={()=>setForm({})}><Plus size={16}/>Nueva venta</Btn>
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
        {["Todos","Pendientes de despacho",...CANALES].map(t=>(
          <button key={t} onClick={()=>setFiltro(t)} style={{ background:filtro===t?C.violet:"#fff", color:filtro===t?"#fff":C.soft, border:`1px solid ${filtro===t?C.violet:C.border}`, borderRadius:20, fontSize:12.5, fontWeight:600, padding:"6px 14px", cursor:"pointer", fontFamily:FONT }}>{t}</button>
        ))}
      </div>
      <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
        {lista.length===0 && <div style={{ color:C.faint, fontSize:14, padding:"30px 0", textAlign:"center" }}>Sin resultados.</div>}
        {lista.map(v=>{ const its=v.items&&v.items.length?v.items:[{producto:v.producto,talle:v.talle,cant:1}]; const totalCam=its.reduce((s,i)=>s+(+i.cant||1),0); const absorbe=v.quien_envio!=="Cliente"&&(+v.costo_envio||0)>0; return (
          <Card key={v.id} style={{ padding:"14px 18px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:12 }}>
              <div style={{ flex:1 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6, flexWrap:"wrap" }}>
                  {v.pedido && v.pedido!=="—" && <Badge color={C.violetDeep} bg={C.violetSoft}>#{v.pedido}</Badge>}
                  <span style={{ fontWeight:700, color:C.ink, fontSize:14.5 }}>{v.cliente}</span>
                  <Badge color={CANAL_CLR[v.canal]} bg="#F4F4F9">{v.canal}</Badge>
                  {totalCam>1 && <Badge color={C.blue} bg={C.blueSoft}>{totalCam} prendas</Badge>}
                  {v.cuotas==="3 sin interés" && <Badge color={C.orange} bg={C.orangeSoft}>3 cuotas</Badge>}
                  {v.cupon && v.cupon.trim() && <Badge color={C.green} bg={C.greenSoft}><span style={{display:"inline-flex",alignItems:"center",gap:3}}><Ticket size={11}/>{v.cupon}</span></Badge>}
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:2, marginBottom:6 }}>
                  {its.map((it,j)=>(<div key={j} style={{ fontSize:13, color:C.soft }}>{it.producto}{it.talle&&it.talle!=="—"?` · talle ${it.talle}`:""}{(+it.cant||1)>1?` · ×${it.cant}`:""}</div>))}
                </div>
                <div style={{ display:"flex", gap:14, fontSize:12, color:C.faint, flexWrap:"wrap" }}>
                  <span>{v.fecha}</span>{v.tel && <span>{v.tel}</span>}{v.acred && <span>acredita {v.acred}</span>}{v.coment && <span style={{ fontStyle:"italic" }}>{v.coment}</span>}
                </div>
              </div>
              <div style={{ textAlign:"right", minWidth:150 }}>
                <div style={{ fontSize:11, color:C.faint }}>Ganancia neta {absorbe?"(con envío)":""}</div>
                <div style={{ fontSize:18, fontWeight:800, color:C.green }}>{ars(v.neto_final)}</div>
                <div style={{ fontSize:11, color:C.soft, marginTop:2 }}>a recibir {ars(v.neto_cobrar)}</div>
                {absorbe && <div style={{ fontSize:11, color:C.red }}>envío −{ars(v.costo_envio)}</div>}
                <div style={{ fontSize:11, color:C.faint }}>pagó {ars(v.bruto)}</div>
                <div style={{ marginTop:8 }}><Switch on={v.desp} onClick={()=>toggleDesp(v.id)} /></div>
              </div>
            </div>
            <div style={{ display:"flex", gap:6, marginTop:10, justifyContent:"flex-end" }}>
              <Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>setForm(v)}><Pencil size={13}/>Editar</Btn>
              {del===v.id ? <><Btn kind="danger" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>eliminar(v.id)}><Check size={13}/>Sí</Btn><Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>setDel(null)}>No</Btn></> : <Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12, color:C.red }} onClick={()=>setDel(v.id)}><Trash2 size={13}/></Btn>}
            </div>
          </Card>
        );})}
      </div>
      {form!==null && <VentaForm initial={form} rates={rates} onSave={guardar} onClose={()=>setForm(null)} notify={notify} />}
    </div>
  );
}

// ====== PEDIDO FORM (proveedor) ======
function PedidoForm({ initial, onSave, onClose, notify }) {
  const [f, setF] = useState(initial?.id ? { ...initial, items:initial.items?.map(i=>({...i, talle:i.talle||"—"}))||[] } : { id:uid(), desc:"", costo:"", envio:"", items:[{prod:"",talle:"—",cant:1}] });
  const setItem = (i,k,v) => setF(p=>({ ...p, items:p.items.map((it,idx)=>idx===i?{...it,[k]:v}:it) }));
  const addItem = () => setF(p=>({ ...p, items:[...p.items,{prod:"",talle:"—",cant:1}] }));
  const delItem = (i) => setF(p=>({ ...p, items:p.items.length>1?p.items.filter((_,idx)=>idx!==i):p.items }));
  const total = (+f.costo||0)+(+f.envio||0);
  return (
    <Modal wide title={initial?.id?"Editar pedido":"Nuevo pedido al mayorista"} sub="Detalle completo de camisetas, talles y cantidades" onClose={onClose}>
      <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr", gap:12 }}>
        <Inp label="Descripción del pedido *" value={f.desc} onChange={e=>setF(p=>({...p,desc:e.target.value}))} placeholder="Ej: Pedido 2 · Titular + varias" />
        <Inp label="Costo mercadería" type="number" value={f.costo} onChange={e=>setF(p=>({...p,costo:e.target.value}))} />
        <Inp label="Envío" type="number" value={f.envio} onChange={e=>setF(p=>({...p,envio:e.target.value}))} />
      </div>
      <Sec>Camisetas del pedido</Sec>
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {f.items.map((it,i)=>(
          <div key={i} style={{ display:"grid", gridTemplateColumns:"3fr 1fr 0.8fr auto", gap:8, alignItems:"center" }}>
            <input style={inputSt} value={it.prod} onChange={e=>setItem(i,"prod",e.target.value)} placeholder="Nombre de la camiseta" />
            <select style={inputSt} value={it.talle} onChange={e=>setItem(i,"talle",e.target.value)}>{TALLES.map(t=><option key={t}>{t}</option>)}</select>
            <input style={inputSt} type="number" min="1" value={it.cant} onChange={e=>setItem(i,"cant",e.target.value)} />
            <IconBtn icon={Trash2} onClick={()=>delItem(i)} color={C.red} bg={C.redSoft}/>
          </div>
        ))}
      </div>
      <Btn kind="soft" style={{ marginTop:10, padding:"7px 13px", fontSize:12.5 }} onClick={addItem}><Plus size={14}/>Agregar camiseta</Btn>
      <div style={{ marginTop:14, fontSize:13, color:C.soft }}>Total del pedido: <b style={{ color:C.red }}>{ars(total)}</b></div>
      <div style={{ display:"flex", gap:10, marginTop:18 }}>
        <Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!f.desc.trim()){notify("Falta la descripción","err");return;} onSave({ ...f, costo:+f.costo||0, envio:+f.envio||0, items:f.items.filter(i=>i.prod.trim()).map(i=>({...i,cant:+i.cant||1})) }); }}>Guardar pedido</Btn>
        <Btn kind="ghost" onClick={onClose}>Cancelar</Btn>
      </div>
    </Modal>
  );
}

// ====== PROVEEDORES ======
function Proveedores({ prov, setProv, stock, setStock, movs, setMovs, notify }) {
  const [open, setOpen] = useState(null);
  const [form, setForm] = useState(null);
  const [pedForm, setPedForm] = useState(null);
  const [del, setDel] = useState(null);
  const [delPed, setDelPed] = useState(null);
  const totalP = p => (p.pedidos||[]).reduce((s,x)=>s+(+x.costo||0)+(+x.envio||0),0);
  const cam = p => (p.pedidos||[]).reduce((s,x)=>s+(x.items||[]).reduce((a,i)=>a+(+i.cant||0),0),0);
  const persist = async (l) => { setProv(l); await save("tbx_proveedores",l); };
  const logMov = async (entries) => { const stamped=entries.map(e=>({ id:uid(), fecha:stamp(), ...e })); const l=[...stamped,...movs]; setMovs(l); await save("tbx_movs",l); };
  const guardar = async (x) => { const xx={...x,id:x.id||uid(),pedidos:x.pedidos||[]}; const l=x.id?prov.map(i=>i.id===x.id?xx:i):[xx,...prov]; await persist(l); setForm(null); notify("Guardado"); };
  const eliminar = async (id) => { await persist(prov.filter(i=>i.id!==id)); setDel(null); notify("Eliminado"); };

  // al guardar un pedido, suma su stock y registra movimientos
  const savePedido = async (provId, ped) => {
    const provObj = prov.find(p=>p.id===provId);
    const prevPed = (provObj?.pedidos||[]).find(x=>x.id===ped.id);
    const l = prov.map(p=>{ if(p.id!==provId) return p; const peds=p.pedidos||[]; const exists=peds.some(x=>x.id===ped.id); return { ...p, pedidos: exists?peds.map(x=>x.id===ped.id?ped:x):[...peds,ped] }; });
    await persist(l);
    if(!prevPed){ // pedido nuevo → entra a stock
      let st=[...stock]; const movEntries=[];
      (ped.items||[]).forEach(it=>{ const t=it.talle||"—"; const ex=st.find(s=>norm(s.producto)===norm(it.prod)&&(s.talle||"—")===t); if(ex) st=st.map(s=>s.id===ex.id?{...s,ingresadas:(+s.ingresadas||0)+(+it.cant||0)}:s); else st=[{ id:uid(), producto:it.prod, talle:t, ingresadas:+it.cant||0, vendidas:0, costo:0, notas:"De "+(provObj?.nombre||"") }, ...st]; movEntries.push({ producto:it.prod, talle:t, tipo:"ingreso", cantidad:+it.cant||0, origen:"Proveedor: "+(provObj?.nombre||"") }); });
      const m=mergeStock(st); setStock(m); await save("tbx_stock",m);
      if(movEntries.length) await logMov(movEntries);
      notify("Pedido guardado y sumado al stock");
    } else notify("Pedido actualizado");
    setPedForm(null); setOpen(provId);
  };
  const delPedido = async (provId, pedId) => { const l = prov.map(p=> p.id!==provId ? p : { ...p, pedidos:(p.pedidos||[]).filter(x=>x.id!==pedId) }); await persist(l); setDelPed(null); notify("Pedido eliminado"); };
  const granTotal = prov.reduce((s,p)=>s+totalP(p),0);

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
        <Title>Proveedores</Title>
        <Btn onClick={()=>setForm({})}><Plus size={16}/>Nuevo mayorista</Btn>
      </div>
      <div style={{ fontSize:13, color:C.soft, marginBottom:18 }}>Total comprado a proveedores: <b style={{ color:C.red }}>{ars(granTotal)}</b> · cada pedido nuevo se suma automáticamente al stock</div>
      <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
        {prov.map(p=>(
          <Card key={p.id} style={{ padding:"16px 18px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
              <div>
                <div style={{ fontWeight:700, color:C.ink, fontSize:15 }}>{p.nombre}</div>
                <div style={{ fontSize:12.5, color:C.faint }}>{p.notas}{p.contacto && p.contacto!=="—" ? ` · ${p.contacto}` : ""}</div>
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ fontWeight:800, color:C.red }}>{ars(totalP(p))}</div>
                <div style={{ fontSize:11, color:C.faint }}>{cam(p)} camisetas · {(p.pedidos||[]).length} pedido(s)</div>
              </div>
            </div>
            <div style={{ display:"flex", gap:6, marginTop:10, flexWrap:"wrap" }}>
              <Btn kind="soft" style={{ padding:"5px 12px", fontSize:12 }} onClick={()=>setOpen(open===p.id?null:p.id)}>{open===p.id?<ChevronDown size={14}/>:<ChevronRight size={14}/>}Ver detalle</Btn>
              <Btn kind="ghost" style={{ padding:"5px 12px", fontSize:12 }} onClick={()=>setPedForm({provId:p.id, pedido:null})}><Plus size={13}/>Agregar pedido</Btn>
              <Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>setForm(p)}><Pencil size={13}/></Btn>
              {del===p.id ? <><Btn kind="danger" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>eliminar(p.id)}>Sí</Btn><Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12 }} onClick={()=>setDel(null)}>No</Btn></> : <Btn kind="ghost" style={{ padding:"5px 10px", fontSize:12, color:C.red }} onClick={()=>setDel(p.id)}><Trash2 size={13}/></Btn>}
            </div>
            {open===p.id && (
              <div style={{ marginTop:12, borderTop:`1px solid ${C.border}`, paddingTop:12 }}>
                {(p.pedidos||[]).length===0 && <div style={{ fontSize:12.5, color:C.faint }}>Sin pedidos cargados. Usá “Agregar pedido”.</div>}
                {(p.pedidos||[]).map((ped)=>(
                  <div key={ped.id} style={{ marginBottom:16, background:"#FAFAFD", borderRadius:12, padding:"12px 14px" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8, gap:8 }}>
                      <span style={{ fontWeight:600, color:C.ink, fontSize:13.5 }}>{ped.desc}</span>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <span style={{ fontSize:12.5, color:C.soft }}>{ars(ped.costo)} + envío {ars(ped.envio)} = <b style={{color:C.red}}>{ars((+ped.costo||0)+(+ped.envio||0))}</b></span>
                        <IconBtn icon={Pencil} onClick={()=>setPedForm({provId:p.id, pedido:ped})}/>
                        {delPed===ped.id ? <><IconBtn icon={Check} onClick={()=>delPedido(p.id,ped.id)} color={C.red} bg={C.redSoft}/><IconBtn icon={X} onClick={()=>setDelPed(null)}/></> : <IconBtn icon={Trash2} onClick={()=>setDelPed(ped.id)} color={C.red}/>}
                      </div>
                    </div>
                    <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                      {(ped.items||[]).map((it,j)=>(
                        <span key={j} style={{ fontSize:12, background:"#fff", border:`1px solid ${C.border}`, borderRadius:8, padding:"4px 9px", color:C.soft }}>{it.prod}{it.talle&&it.talle!=="—"?` · ${it.talle}`:""} <b style={{ color:C.ink }}>×{it.cant}</b></span>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>
        ))}
      </div>

      {form!==null && (
        <Modal title={form.id?"Editar mayorista":"Nuevo mayorista"} sub="Los pedidos se cargan con “Agregar pedido”" onClose={()=>setForm(null)}>
          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
            <Inp label="Nombre *" value={form.nombre||""} onChange={e=>setForm(p=>({...p,nombre:e.target.value}))} />
            <Inp label="Contacto" value={form.contacto||""} onChange={e=>setForm(p=>({...p,contacto:e.target.value}))} />
            <Inp label="Notas" value={form.notas||""} onChange={e=>setForm(p=>({...p,notas:e.target.value}))} />
          </div>
          <div style={{ display:"flex", gap:10, marginTop:18 }}><Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!form.nombre){notify("Falta el nombre","err");return;} guardar(form); }}>Guardar</Btn><Btn kind="ghost" onClick={()=>setForm(null)}>Cancelar</Btn></div>
        </Modal>
      )}
      {pedForm!==null && <PedidoForm initial={pedForm.pedido} onSave={(ped)=>savePedido(pedForm.provId, ped)} onClose={()=>setPedForm(null)} notify={notify} />}
    </div>
  );
}

// ====== IMPORTAR / EXPORTAR (respaldo completo) ======
function Importar({ data, apply, rates, notify }) {
  const [preview, setPreview] = useState(null);

  const exportAll = () => {
    const wb = XLSX.utils.book_new();
    const add = (rows, name) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length?rows:[{}]), name);
    add(data.ventas.map(v=>({ ID:v.id, Pedido:v.pedido, Fecha:v.fecha, Cliente:v.cliente, "DNI/CUIT":v.dni, Email:v.email, "Teléfono":v.tel, "Dirección":v.direccion, Ciudad:v.ciudad, CP:v.cp, Provincia:v.provincia, Productos:itemsLabel(v), Canal:v.canal, "Método":v.metodo, Cuotas:v.cuotas, Pagado:v.bruto, "Retención IIBB":v.retencion, "Interés 3c":v.interes3c, "Cargo MP":v.cargo_mp, "Cargo plataforma":v.cargo_plataforma, "Costo envío":v.costo_envio, "Quién envío":v.quien_envio, "Neto a recibir":v.neto_cobrar, "Ganancia neta":v.neto_final, "Cupón":v.cupon, Despachado:v.desp?"Sí":"No", "Acreditación":v.acred, Comentarios:v.coment, _items:JSON.stringify(v.items||[]) })), "Ventas");
    add(data.egresos.map(e=>({ ID:e.id, Fecha:e.fecha, "Descripción":e.desc, "Categoría":e.cat, Monto:e.monto, Notas:e.notas })), "Egresos");
    add(data.prov.map(p=>({ ID:p.id, Nombre:p.nombre, Contacto:p.contacto, Notas:p.notas })), "Proveedores");
    const pedRows = []; data.prov.forEach(p=>(p.pedidos||[]).forEach(ped=>pedRows.push({ ProveedorID:p.id, ProveedorNombre:p.nombre, PedidoID:ped.id, "Descripción":ped.desc, Costo:ped.costo, "Envío":ped.envio, _items:JSON.stringify(ped.items||[]) })));
    add(pedRows, "Pedidos");
    add(data.stock.map(s=>({ ID:s.id, Producto:s.producto, Talle:s.talle, Ingresadas:s.ingresadas, Vendidas:s.vendidas, "Stock actual":(+s.ingresadas||0)-(+s.vendidas||0), Costo:s.costo, Notas:s.notas })), "Stock");
    add(buildClientes(data.ventas, data.clientes).map(c=>({ Key:c.key, Nombre:c.nombre, DNI:c.dni, Email:c.email, "Teléfono":c.tel, "Dirección":c.direccion, Ciudad:c.ciudad, CP:c.cp, Provincia:c.provincia, Notas:c.notas, Compras:c.compras, "Ganancia neta":c.neto })), "Clientes");
    add(data.movs.map(m=>({ ID:m.id, Fecha:m.fecha, Producto:m.producto, Talle:m.talle, Tipo:m.tipo, Cantidad:m.cantidad, Origen:m.origen })), "Movimientos");
    add([{ "MP %":rates.mp, "CPT %":rates.cpt, "IIBB %":rates.iibb, "Interés 3c %":rates.c3 }], "Tasas");
    downloadXLSX(wb, `Tribuneros_respaldo_${today()}.xlsx`);
    notify("Excel descargado con todos los módulos");
  };

  const onFile = (e) => {
    const file = e.target.files?.[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type:"array" });
        const get = (name) => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval:"" }) : null;
        const J = s => { try { return JSON.parse(s||"[]"); } catch { return []; } };
        const parsed = {}; const counts = {};

        const vR = get("Ventas");
        if(vR){ parsed.ventas = vR.filter(r=>r.Cliente||r._items).map(r=>{ let items=J(r._items); if(!items.length) items=[{producto:r.Productos||"",talle:"—",cant:1}]; return computeVenta(migrateVenta({ id:r.ID||uid(), pedido:String(r.Pedido||""), fecha:String(r.Fecha||today()), cliente:String(r.Cliente||""), dni:String(r["DNI/CUIT"]||""), email:String(r.Email||""), tel:String(r["Teléfono"]||""), direccion:String(r["Dirección"]||""), ciudad:String(r.Ciudad||""), cp:String(r.CP||""), provincia:String(r.Provincia||""), items, canal:String(r.Canal||"Tiendanube"), metodo:String(r["Método"]||"Mercado Pago"), cuotas:String(r.Cuotas||"1"), bruto:+r.Pagado||0, retencion:r["Retención IIBB"]!==""?+r["Retención IIBB"]:null, interes3c:r["Interés 3c"]!==""?+r["Interés 3c"]:null, cargo_mp:r["Cargo MP"]!==""?+r["Cargo MP"]:null, cargo_plataforma:r["Cargo plataforma"]!==""?+r["Cargo plataforma"]:null, costo_envio:+r["Costo envío"]||0, quien_envio:String(r["Quién envío"]||"Vos (absorbés)"), cupon:String(r["Cupón"]||""), desp:/s[íi]/i.test(String(r.Despachado)), acred:String(r["Acreditación"]||""), coment:String(r.Comentarios||"") }), rates); }); counts.Ventas=parsed.ventas.length; }

        const eR = get("Egresos");
        if(eR){ parsed.egresos = eR.filter(r=>r["Descripción"]).map(r=>({ id:r.ID||uid(), fecha:String(r.Fecha||today()), desc:String(r["Descripción"]), cat:String(r["Categoría"]||"Otro"), monto:+r.Monto||0, notas:String(r.Notas||"") })); counts.Egresos=parsed.egresos.length; }

        const pR = get("Proveedores"); const pedR = get("Pedidos");
        if(pR){ const provs=pR.filter(r=>r.Nombre).map(r=>({ id:r.ID||uid(), nombre:String(r.Nombre), contacto:String(r.Contacto||""), notas:String(r.Notas||""), pedidos:[] })); if(pedR) pedR.forEach(r=>{ const pr=provs.find(p=>p.id===r.ProveedorID)||provs.find(p=>norm(p.nombre)===norm(r.ProveedorNombre)); if(pr) pr.pedidos.push({ id:r.PedidoID||uid(), desc:String(r["Descripción"]||""), costo:+r.Costo||0, envio:+r["Envío"]||0, items:J(r._items) }); }); parsed.prov=provs; counts.Proveedores=provs.length; }

        const sR = get("Stock");
        if(sR){ parsed.stock = mergeStock(sR.filter(r=>r.Producto).map(r=>({ id:r.ID||uid(), producto:String(r.Producto), talle:String(r.Talle||"—"), ingresadas:+r.Ingresadas||0, vendidas:+r.Vendidas||0, costo:+r.Costo||0, notas:String(r.Notas||"") }))); counts.Stock=parsed.stock.length; }

        const cR = get("Clientes");
        if(cR){ parsed.clientes = cR.filter(r=>r.Nombre).map(r=>({ key:r.Key||norm(r.Nombre), nombre:String(r.Nombre), dni:String(r.DNI||""), email:String(r.Email||""), tel:String(r["Teléfono"]||""), direccion:String(r["Dirección"]||""), ciudad:String(r.Ciudad||""), cp:String(r.CP||""), provincia:String(r.Provincia||""), notas:String(r.Notas||"") })); counts.Clientes=parsed.clientes.length; }

        const mR = get("Movimientos");
        if(mR){ parsed.movs = mR.filter(r=>r.Producto).map(r=>({ id:r.ID||uid(), fecha:String(r.Fecha||""), producto:String(r.Producto), talle:String(r.Talle||"—"), tipo:String(r.Tipo||"ingreso"), cantidad:+r.Cantidad||0, origen:String(r.Origen||"") })); counts.Movimientos=parsed.movs.length; }

        const tR = get("Tasas");
        if(tR && tR.length){ const t=tR[0]; parsed.rates={ mp:+t["MP %"]||rates.mp, cpt:+t["CPT %"]||rates.cpt, iibb:+t["IIBB %"]||rates.iibb, c3:+t["Interés 3c %"]||rates.c3 }; counts.Tasas=1; }

        if(!Object.keys(parsed).length){ notify("El archivo no tiene hojas reconocibles","err"); return; }
        setPreview({ parsed, counts }); notify("Archivo leído");
      } catch(err){ notify("No se pudo leer el archivo","err"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const restore = async () => { await apply(preview.parsed); setPreview(null); notify("Datos restaurados desde el Excel"); };

  return (
    <div>
      <Title>Importar / Exportar</Title>
      <p style={{ color:C.soft, fontSize:14, maxWidth:600, margin:"8px 0 20px" }}>Descargá <b>todo</b> en un solo Excel (ventas, egresos, proveedores, pedidos, stock, clientes, movimientos y tasas), o subí un respaldo para restaurarlo. Útil para guardar copia o pasar los datos a otra compu.</p>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:20 }}>
        <Card style={{ padding:22 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}><Download size={18} color={C.violet}/><div style={{ fontWeight:700, color:C.ink }}>Descargar todo</div></div>
          <div style={{ fontSize:12.5, color:C.soft, marginBottom:14 }}>Genera un .xlsx con una hoja por módulo. Cada vez exporta el estado actual.</div>
          <Btn onClick={exportAll}><FileSpreadsheet size={16}/>Descargar Excel completo</Btn>
        </Card>
        <Card style={{ padding:22 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}><Upload size={18} color={C.violet}/><div style={{ fontWeight:700, color:C.ink }}>Subir respaldo</div></div>
          <div style={{ fontSize:12.5, color:C.soft, marginBottom:14 }}>Elegí un Excel exportado desde acá. Reemplaza los datos actuales por los del archivo.</div>
          <label style={{ display:"inline-flex" }}>
            <span style={{ background:C.violet, color:"#fff", borderRadius:10, fontSize:13, fontWeight:600, padding:"9px 16px", cursor:"pointer", display:"inline-flex", alignItems:"center", gap:6 }}><Upload size={16}/>Elegir archivo .xlsx</span>
            <input type="file" accept=".xlsx,.xls" onChange={onFile} style={{ display:"none" }} />
          </label>
        </Card>
      </div>

      {preview && (
        <Card style={{ padding:22 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, gap:12, flexWrap:"wrap" }}>
            <div style={{ fontWeight:700, color:C.ink, display:"flex", alignItems:"center", gap:8 }}><Database size={18} color={C.violet}/>Listo para restaurar</div>
            <div style={{ display:"flex", gap:8 }}><Btn onClick={restore}><Check size={16}/>Restaurar todo</Btn><Btn kind="ghost" onClick={()=>setPreview(null)}>Cancelar</Btn></div>
          </div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
            {Object.entries(preview.counts).map(([k,v])=>(<Badge key={k} color={C.violetDeep} bg={C.violetSoft}>{k}: {v}</Badge>))}
          </div>
          <div style={{ fontSize:12, color:C.red, marginTop:12 }}>⚠ Al restaurar se reemplazan los datos actuales con los del archivo.</div>
        </Card>
      )}
    </div>
  );
}

// ====== EGRESOS ======
function Egresos({ egresos, setEgresos, notify }) {
  const [form, setForm] = useState(null);
  const [del, setDel] = useState(null);
  const CATS = ["Mercadería","Envío","Insumos","Fijo mensual","Publicidad","Otro"];
  const guardar = async (x)=>{ const xx={...x,id:x.id||uid(),monto:+x.monto||0}; const l=x.id?egresos.map(i=>i.id===x.id?xx:i):[xx,...egresos]; setEgresos(l); await save("tbx_egresos",l); setForm(null); notify("Guardado"); };
  const eliminar = async (id)=>{ const l=egresos.filter(i=>i.id!==id); setEgresos(l); await save("tbx_egresos",l); setDel(null); notify("Eliminado"); };
  const total = egresos.reduce((s,e)=>s+(+e.monto||0),0);
  const CAT_CLR = { Mercadería:C.violet, Envío:C.blue, Insumos:C.faint, "Fijo mensual":C.orange, Publicidad:C.green, Otro:C.faint };
  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
        <Title>Egresos</Title><Btn onClick={()=>setForm({})}><Plus size={16}/>Nuevo egreso</Btn>
      </div>
      <div style={{ fontSize:13, color:C.soft, marginBottom:18 }}>Total: <b style={{ color:C.red }}>{ars(total)}</b></div>
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {egresos.map(e=>(
          <Card key={e.id} style={{ padding:"12px 16px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}><span style={{ fontWeight:600, color:C.ink, fontSize:14 }}>{e.desc}</span><Badge color={CAT_CLR[e.cat]||C.faint} bg="#F4F4F9">{e.cat}</Badge></div>
              <div style={{ fontSize:12, color:C.faint, marginTop:2 }}>{e.fecha}{e.notas && ` · ${e.notas}`}</div>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <b style={{ color:C.red }}>{ars(e.monto)}</b>
              <Btn kind="ghost" style={{ padding:"5px 9px", fontSize:12 }} onClick={()=>setForm(e)}><Pencil size={13}/></Btn>
              {del===e.id ? <><Btn kind="danger" style={{ padding:"5px 9px", fontSize:12 }} onClick={()=>eliminar(e.id)}>Sí</Btn><Btn kind="ghost" style={{ padding:"5px 9px", fontSize:12 }} onClick={()=>setDel(null)}>No</Btn></> : <Btn kind="ghost" style={{ padding:"5px 9px", fontSize:12, color:C.red }} onClick={()=>setDel(e.id)}><Trash2 size={13}/></Btn>}
            </div>
          </Card>
        ))}
      </div>
      {form!==null && (
        <Modal title={form.id?"Editar egreso":"Nuevo egreso"} onClose={()=>setForm(null)}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <Inp span label="Descripción *" value={form.desc||""} onChange={e=>setForm(p=>({...p,desc:e.target.value}))} placeholder="Ej: Meta Ads campaña mundial" />
            <Sel label="Categoría" value={form.cat||"Mercadería"} options={CATS} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} />
            <Inp label="Monto *" type="number" value={form.monto||""} onChange={e=>setForm(p=>({...p,monto:e.target.value}))} />
            <Inp label="Fecha" type="date" value={form.fecha||today()} onChange={e=>setForm(p=>({...p,fecha:e.target.value}))} />
            <Inp label="Notas" value={form.notas||""} onChange={e=>setForm(p=>({...p,notas:e.target.value}))} />
          </div>
          <div style={{ display:"flex", gap:10, marginTop:18 }}><Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!form.desc||!form.monto){notify("Falta descripción y monto","err");return;} guardar(form); }}>Guardar</Btn><Btn kind="ghost" onClick={()=>setForm(null)}>Cancelar</Btn></div>
        </Modal>
      )}
    </div>
  );
}

// ====== CLIENTES ======
function ClienteForm({ initial, onSave, onClose, notify }) {
  const [f, setF] = useState({ ...initial });
  const sf = (k,v) => setF(p=>({ ...p, [k]:v }));
  return (
    <Modal wide title={f.nombre||"Cliente"} sub={`${f.compras||0} compra(s) · ganancia neta ${ars(f.neto||0)} · pagó ${ars(f.total||0)}`} onClose={onClose}>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:6, background:C.violetSoft2, borderRadius:12, padding:14 }}>
        <div><div style={{ fontSize:11, color:C.soft }}>Compras</div><div style={{ fontSize:18, fontWeight:800, color:C.ink }}>{f.compras||0}</div></div>
        <div><div style={{ fontSize:11, color:C.soft }}>Ganancia neta</div><div style={{ fontSize:18, fontWeight:800, color:C.green }}>{ars(f.neto||0)}</div></div>
        <div><div style={{ fontSize:11, color:C.soft }}>Última compra</div><div style={{ fontSize:15, fontWeight:700, color:C.ink }}>{f.ultima||"—"}</div></div>
      </div>
      <Sec>Datos del cliente (editables)</Sec>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        <Inp span label="Nombre completo" value={f.nombre||""} onChange={e=>sf("nombre",e.target.value)} />
        <Inp label="DNI / CUIT" value={f.dni||""} onChange={e=>sf("dni",e.target.value)} />
        <Inp label="Teléfono" value={f.tel||""} onChange={e=>sf("tel",e.target.value)} />
        <Inp span label="Email" value={f.email||""} onChange={e=>sf("email",e.target.value)} />
        <Inp span label="Dirección" value={f.direccion||""} onChange={e=>sf("direccion",e.target.value)} />
        <Inp label="Ciudad" value={f.ciudad||""} onChange={e=>sf("ciudad",e.target.value)} />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          <Inp label="CP" value={f.cp||""} onChange={e=>sf("cp",e.target.value)} />
          <Inp label="Provincia" value={f.provincia||""} onChange={e=>sf("provincia",e.target.value)} />
        </div>
        <Inp span label="Notas internas" value={f.notas||""} onChange={e=>sf("notas",e.target.value)} placeholder="Ej: cliente frecuente, talle habitual L, prefiere retiro..." />
      </div>
      <div style={{ display:"flex", gap:10, marginTop:18 }}>
        <Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!f.nombre || !f.nombre.trim()){notify("Falta el nombre","err");return;} onSave(f); }}>Guardar datos</Btn>
        <Btn kind="ghost" onClick={onClose}>Cancelar</Btn>
      </div>
    </Modal>
  );
}

function Clientes({ ventas, clientes, setClientes, notify }) {
  const [form, setForm] = useState(null);
  const lista = buildClientes(ventas, clientes);
  const guardar = async (c) => {
    const key = c.key || norm(c.nombre);
    const rec = { key, nombre:c.nombre, dni:c.dni||"", email:c.email||"", tel:c.tel||"", direccion:c.direccion||"", ciudad:c.ciudad||"", cp:c.cp||"", provincia:c.provincia||"", notas:c.notas||"" };
    const exists = clientes.some(m=>(m.key||norm(m.nombre))===key);
    const l = exists ? clientes.map(m=>(m.key||norm(m.nombre))===key?rec:m) : [rec, ...clientes];
    setClientes(l); await save("tbx_clientes", l); setForm(null); notify("Datos del cliente guardados");
  };
  return (
    <div>
      <Title>Clientes</Title>
      <p style={{ color:C.soft, fontSize:13, margin:"8px 0 18px" }}>Se arma desde tus ventas y guarda los datos más completos. Tocá “Ver detalles” para verlos todos y editarlos. {lista.length} cliente(s).</p>
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {lista.map((c)=>(
          <Card key={c.key} style={{ padding:"13px 16px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:10 }}>
              <div style={{ minWidth:0 }}>
                <div style={{ fontWeight:700, color:C.ink, fontSize:14 }}>{c.nombre}</div>
                <div style={{ fontSize:12, color:C.faint }}>{[c.tel,c.ciudad,c.provincia].filter(Boolean).join(" · ")||"sin datos de contacto"}</div>
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:12, flexShrink:0 }}>
                <div style={{ textAlign:"right" }}><div style={{ fontWeight:800, color:C.green }}>{ars(c.neto)}</div><div style={{ fontSize:11, color:C.faint }}>{c.compras} compra(s)</div></div>
                <Btn kind="soft" style={{ padding:"6px 12px", fontSize:12 }} onClick={()=>setForm(c)}><Eye size={14}/>Ver detalles</Btn>
              </div>
            </div>
          </Card>
        ))}
        {lista.length===0 && <div style={{ color:C.faint, fontSize:14, padding:"24px 0", textAlign:"center" }}>Todavía no hay clientes.</div>}
      </div>
      {form!==null && <ClienteForm initial={form} onSave={guardar} onClose={()=>setForm(null)} notify={notify} />}
    </div>
  );
}

// ====== DASHBOARD ======
function Dashboard({ ventas, egresos, estadisticas }) {
  const audit = estadisticas?.rentabilidad_auditoria;
  const potencial = estadisticas?.ganancia_potencial_stock;
  const bruto = ventas.reduce((s,v)=>s+(+v.bruto||0),0);
  const neto = ventas.reduce((s,v)=>s+(+v.neto_final||0),0);
  const aRecibir = ventas.reduce((s,v)=>s+(+v.neto_cobrar||0),0);
  const egr = egresos.reduce((s,e)=>s+(+e.monto||0),0);
  const resultado = neto - egr;
  const comMP = ventas.reduce((s,v)=>s+(+v.cargo_mp||0),0);
  const ret = ventas.reduce((s,v)=>s+(+v.retencion||0),0);
  const int3 = ventas.reduce((s,v)=>s+(+v.interes3c||0),0);
  const env = ventas.reduce((s,v)=>s+(v.quien_envio!=="Cliente"?(+v.costo_envio||0):0),0);
  const pend = ventas.filter(v=>!v.desp).length;
  const porCanal = CANALES.map(c=>({ name:c, value: ventas.filter(v=>v.canal===c).reduce((s,v)=>s+(+v.neto_final||0),0) })).filter(x=>x.value>0);
  const cats = ["Mercadería","Envío","Insumos","Fijo mensual","Publicidad"];
  const porCat = cats.map(c=>({ name:c, value: egresos.filter(e=>e.cat===c).reduce((s,e)=>s+(+e.monto||0),0) })).filter(x=>x.value>0);

  const rk = {};
  ventas.forEach(v=>{ const its=v.items&&v.items.length?v.items:[{producto:v.producto,cant:1}]; its.forEach(it=>{ const k=(it.producto||"").trim(); if(!k) return; rk[k]=(rk[k]||0)+(+it.cant||1); }); });
  const ranking = Object.entries(rk).map(([name,cant])=>({name,cant})).sort((a,b)=>b.cant-a.cant).slice(0,8);
  const maxRk = ranking.length ? ranking[0].cant : 1;
  const MEDAL = ["#F2C94C","#BDBDBD","#CD7F32"];

  return (
    <div>
      <Title>Resumen</Title>
      <div style={{ color:C.soft, fontSize:13, marginTop:2, marginBottom:18 }}>Estado general de Tribuneros · Junio 2026</div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(155px,1fr))", gap:14, marginBottom:16 }}>
        <KPI icon={ShoppingBag} label="Cobrado a clientes" value={ars(bruto)} color={C.violet} bg={C.violetSoft}/>
        <KPI icon={ShoppingCart} label="Total a recibir" value={ars(aRecibir)} sub="neto de comisiones, sin envío" color={C.violetDeep} bg={C.violetSoft}/>
        <KPI icon={ArrowUpRight} label="Ganancia neta" value={ars(neto)} sub="con envíos restados" color={C.green} bg={C.greenSoft}/>
        <KPI icon={TrendingDown} label="Egresos (caja)" value={ars(egr)} color={C.red} bg={C.redSoft}/>
        <KPI icon={Package} label="Despachos pendientes" value={pend} color={C.orange} bg={C.orangeSoft}/>
      </div>

      {(audit || potencial) && (
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))", gap:14, marginBottom:16 }}>
          {audit && (
            <KPI icon={Trophy} label="Margen sobre lo vendido (auditoría)" value={`${audit.margen_pct}%`} sub={ars(audit.waterfall[audit.waterfall.length-1]?.monto || 0)+" · costo de lo que se vendió, no caja"} color={C.green} bg={C.greenSoft}/>
          )}
          {potencial && (
            <KPI icon={PackageCheck} label="Ganancia potencial en stock" value={ars(potencial.ganancia_potencial)} sub={`si vendés las ${potencial.unidades} unidades al ritmo actual`} color={C.violetDeep} bg={C.violetSoft}/>
          )}
        </div>
      )}
      <div style={{ display:"grid", gridTemplateColumns:"1.4fr 1fr", gap:14, marginBottom:16 }}>
        <Card style={{ padding:22 }}>
          <div style={{ fontSize:14, color:C.soft, fontWeight:600 }}>Resultado del mes</div>
          <div style={{ fontSize:34, fontWeight:800, color:resultado>=0?C.green:C.red, letterSpacing:-1, margin:"4px 0 2px" }}>{ars(resultado)}</div>
          <div style={{ fontSize:12, color:C.faint, marginBottom:14 }}>{resultado>=0?"Ganancia neta":"Todavía recuperando la inversión en stock"}</div>
          <div style={{ height:180 }}>
            <ResponsiveContainer>
              <BarChart data={[{n:"Ganancia neta",v:neto},{n:"Egresos",v:egr}]}>
                <XAxis dataKey="n" tick={{ fontSize:12, fill:C.soft }} axisLine={false} tickLine={false}/><YAxis hide/>
                <Tooltip formatter={v=>ars(v)} />
                <Bar dataKey="v" radius={[8,8,0,0]}><Cell fill={C.green}/><Cell fill={C.red}/></Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <KPI icon={Settings} label="Comisiones MP" value={ars(comMP)} color={C.violet} bg={C.violetSoft}/>
          <KPI icon={Settings} label="Retenciones IIBB" value={ars(ret)} color={C.orange} bg={C.orangeSoft}/>
          <KPI icon={Settings} label="Interés 3 cuotas" value={ars(int3)} color={C.red} bg={C.redSoft}/>
          <KPI icon={Truck} label="Envíos absorbidos" value={ars(env)} color={C.blue} bg={C.blueSoft}/>
        </div>
      </div>

      <Card style={{ padding:22, marginBottom:16 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
          <Trophy size={18} color="#F2C94C"/><div style={{ fontWeight:700, color:C.ink }}>Camisetas más vendidas</div>
        </div>
        {ranking.length===0 && <div style={{ fontSize:13, color:C.faint }}>Todavía no hay ventas cargadas.</div>}
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {ranking.map((r,i)=>(
            <div key={r.name} style={{ display:"flex", alignItems:"center", gap:12 }}>
              <div style={{ width:24, height:24, borderRadius:7, flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:800, color:i<3?"#fff":C.soft, background:i<3?MEDAL[i]:"#F2F2F7" }}>{i+1}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                  <span style={{ fontSize:13, fontWeight:600, color:C.ink, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{r.name}</span>
                  <span style={{ fontSize:13, fontWeight:700, color:C.violetDeep, flexShrink:0, marginLeft:8 }}>{r.cant} u.</span>
                </div>
                <div style={{ height:7, borderRadius:4, background:"#F2F2F7", overflow:"hidden" }}>
                  <div style={{ width:`${Math.round(r.cant/maxRk*100)}%`, height:"100%", borderRadius:4, background:`linear-gradient(90deg,${C.violet},${C.lav})` }}/>
                </div>
              </div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <Card style={{ padding:22 }}>
          <div style={{ fontWeight:700, color:C.ink, marginBottom:8 }}>Ganancia por canal</div>
          <div style={{ height:230 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={porCanal} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}>{porCanal.map((e,i)=><Cell key={i} fill={PIE[i%PIE.length]}/>)}</Pie>
                <Tooltip formatter={v=>ars(v)}/><Legend iconType="circle" wrapperStyle={{ fontSize:12 }}/>
              </PieChart>
            </ResponsiveContainer>
          </div>
        </Card>
        <Card style={{ padding:22 }}>
          <div style={{ fontWeight:700, color:C.ink, marginBottom:8 }}>Egresos por categoría</div>
          <div style={{ height:230 }}>
            <ResponsiveContainer>
              <BarChart data={porCat} layout="vertical" margin={{ left:20 }}>
                <XAxis type="number" hide/><YAxis type="category" dataKey="name" tick={{ fontSize:12, fill:C.soft }} axisLine={false} tickLine={false} width={90}/>
                <Tooltip formatter={v=>ars(v)}/><Bar dataKey="value" radius={[0,8,8,0]} fill={C.violet}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ====== ESTADISTICAS ======
function agrupar(ventas, campo, valorFn) {
  const m = {};
  ventas.forEach(v=>{
    const k = v[campo] || "—";
    m[k] = (m[k]||0) + (valorFn ? valorFn(v) : 1);
  });
  return Object.entries(m).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
}

function SectionTitle({ children, icon:Icon }) {
  return <div style={{ display:"flex", alignItems:"center", gap:8, margin:"22px 0 12px", fontWeight:800, fontSize:15, color:C.ink }}>{Icon && <Icon size={17} color={C.violet}/>}{children}</div>;
}

function BarList({ data, fmt, color }) {
  const max = Math.max(1, ...data.map(d=>d.value));
  return (
    <div style={{ display:"flex", flexDirection:"column", gap:9 }}>
      {data.map(d=>(
        <div key={d.name} style={{ display:"grid", gridTemplateColumns:"140px 1fr auto", alignItems:"center", gap:10, fontSize:12.5 }}>
          <span style={{ color:C.soft, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={d.name}>{d.name}</span>
          <div style={{ background:C.violetSoft2, borderRadius:5, height:14, position:"relative" }}>
            <div style={{ width:`${Math.max(4,d.value/max*100)}%`, height:"100%", borderRadius:5, background:color||C.violet }}/>
          </div>
          <span style={{ fontFamily:"monospace", color:C.ink, fontWeight:600 }}>{fmt?fmt(d.value):d.value}</span>
        </div>
      ))}
      {!data.length && <div style={{ fontSize:12.5, color:C.faint }}>Sin datos.</div>}
    </div>
  );
}

function Estadisticas({ ventas, estadisticas }) {
  const tn = ventas.filter(v=>v.canal==="Tiendanube");

  const porEquipoDetectado = (() => {
    const eq = { "Boca":"Boca Juniors", "River":"River Plate", "Argentina":"Selección Argentina", "Selecci":"Selección Argentina" };
    const m = {};
    tn.forEach(v=>(v.items||[]).forEach(it=>{
      const nombre = (it.producto||"");
      const clave = Object.keys(eq).find(k=>nombre.includes(k));
      const label = clave ? eq[clave] : "Otros / internacional";
      m[label] = (m[label]||0) + (+it.cant||1);
    }));
    return Object.entries(m).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
  })();

  const porTalle = (() => {
    const m = {};
    tn.forEach(v=>(v.items||[]).forEach(it=>{ const t=it.talle||"—"; if(t==="—") return; m[t]=(m[t]||0)+(+it.cant||1); }));
    return Object.entries(m).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
  })();

  const porModelo = (() => {
    const m = {};
    tn.forEach(v=>(v.items||[]).forEach(it=>{ const k=it.producto||"—"; m[k]=(m[k]||0)+(+it.cant||1); }));
    return Object.entries(m).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value).slice(0,10);
  })();

  const porMedio = agrupar(tn, "metodo", v=>+v.bruto||0);
  const porForma = agrupar(tn, "forma_pago", v=>+v.bruto||0);
  const porCuotas = agrupar(tn.filter(v=>v.metodo==="Mercado Pago"), "cuotas");
  const porProvincia = agrupar(tn, "provincia");
  const porEnvio = agrupar(tn, "quien_envio");
  const conCupon = tn.filter(v=>v.cupon).length;

  const ads = estadisticas?.ads;
  const caidos = estadisticas?.pedidos_caidos || [];
  const mp = estadisticas?.mercadopago;
  const audit = estadisticas?.rentabilidad_auditoria;

  return (
    <div>
      <Title>Estadísticas</Title>
      <div style={{ color:C.soft, fontSize:13, marginTop:2, marginBottom:6 }}>
        Todo lo que se pudo sacar de tus ventas reales{estadisticas ? ` y de la auditoría (actualizada ${estadisticas.actualizado})` : ""}.
      </div>

      <SectionTitle icon={ShoppingBag}>Qué se vendió</SectionTitle>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Unidades por equipo</div><BarList data={porEquipoDetectado}/></Card>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Unidades por talle</div><BarList data={porTalle}/></Card>
      </div>
      <Card style={{ padding:20, marginTop:14 }}>
        <div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Top 10 modelos por unidades</div>
        <BarList data={porModelo}/>
      </Card>

      <SectionTitle icon={Wallet}>Cómo te pagaron</SectionTitle>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Monto por medio de pago</div><BarList data={porMedio} fmt={ars}/></Card>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Monto por forma de pago</div><BarList data={porForma} fmt={ars} color={C.orange}/></Card>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginTop:14 }}>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Cuotas con Mercado Pago</div><BarList data={porCuotas} color={C.blue}/></Card>
        <Card style={{ padding:20 }}>
          <div style={{ fontWeight:700, marginBottom:6, fontSize:13 }}>Cupones</div>
          <div style={{ fontSize:28, fontWeight:800, color:C.ink }}>{conCupon}/{tn.length}</div>
          <div style={{ fontSize:12, color:C.faint }}>pedidos con cupón aplicado</div>
          {estadisticas?.cupones && <div style={{ fontSize:12, color:C.soft, marginTop:8 }}>TRIBUNEROS10 sumó {ars(estadisticas.cupones.tribuneros10_descuento)} en descuentos · {estadisticas.cupones.regalo10_cargados} cupones "REGALO" cargados, {estadisticas.cupones.regalo10_usados} usados.</div>}
        </Card>
      </div>

      <SectionTitle icon={MapPin}>Dónde están tus compradores</SectionTitle>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Pedidos por provincia</div><BarList data={porProvincia}/></Card>
        <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Método de envío</div><BarList data={porEnvio} color={C.green}/></Card>
      </div>

      {!!caidos.length && (
        <>
          <SectionTitle icon={AlertTriangle}>Pedidos caídos, pendientes y reembolsados</SectionTitle>
          <div className="callout" style={{ background:C.orangeSoft, border:`1px solid ${C.border}`, borderRadius:12, padding:"12px 16px", fontSize:12.5, color:C.ink, marginBottom:12 }}>
            <strong>{caidos.length} pedidos, {ars(estadisticas.pedidos_caidos_total)}</strong> en el mismo período que nunca se cobraron — no están en el export de ventas, salen de la API de Tiendanube.
          </div>
          <Card style={{ padding:0, overflow:"hidden" }}>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5 }}>
                <thead><tr style={{ textAlign:"left", color:C.faint, fontSize:11, textTransform:"uppercase" }}>
                  <th style={{ padding:"10px 14px" }}>Pedido</th><th style={{ padding:"10px 14px" }}>Fecha</th><th style={{ padding:"10px 14px" }}>Producto</th><th style={{ padding:"10px 14px" }}>Estado</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Monto</th>
                </tr></thead>
                <tbody>
                  {caidos.map(c=>(
                    <tr key={c.pedido} style={{ borderTop:`1px solid ${C.border}` }}>
                      <td style={{ padding:"9px 14px" }}>#{c.pedido}</td>
                      <td style={{ padding:"9px 14px" }}>{c.fecha}</td>
                      <td style={{ padding:"9px 14px", color:C.soft }}>{c.producto}</td>
                      <td style={{ padding:"9px 14px" }}><Badge color={C.red} bg={C.redSoft}>{c.estado_pago}</Badge></td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{ars(c.monto)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}

      {ads && (
        <>
          <SectionTitle icon={Megaphone}>Publicidad — cuenta {ads.cuenta.nombre}</SectionTitle>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:12, marginBottom:14 }}>
            <KPI icon={Megaphone} label="Gasto en Ads" value={"US$"+ads.gasto_usd} sub={ars(ads.gasto_ars)} color={C.orange} bg={C.orangeSoft}/>
            <KPI icon={ShoppingCart} label="Compras atribuidas (píxel)" value={ads.compras_atribuidas_pixel} sub={`vs. ${tn.length} reales`} color={C.red} bg={C.redSoft}/>
            <KPI icon={ArrowUpRight} label="ROAS real" value={ads.roas_real+"x"} sub={`Meta reporta ${ads.roas_meta}x`} color={C.green} bg={C.greenSoft}/>
            <KPI icon={Eye} label="Agregar al carrito" value={ads.add_to_cart_atribuidos} color={C.blue} bg={C.blueSoft}/>
          </div>
          <Card style={{ padding:0, overflow:"hidden", marginBottom:14 }}>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5 }}>
                <thead><tr style={{ textAlign:"left", color:C.faint, fontSize:11, textTransform:"uppercase" }}>
                  <th style={{ padding:"10px 14px" }}>Campaña</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Gasto</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Impresiones</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Clics</th><th style={{ padding:"10px 14px", textAlign:"right" }}>CTR</th><th style={{ padding:"10px 14px" }}>Resultado</th><th style={{ padding:"10px 14px", textAlign:"right" }}>ROAS</th>
                </tr></thead>
                <tbody>
                  {ads.campañas.map(c=>(
                    <tr key={c.nombre} style={{ borderTop:`1px solid ${C.border}` }}>
                      <td style={{ padding:"9px 14px" }}>{c.nombre}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{c.gasto_usd?`US$${c.gasto_usd}`:"—"}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{c.impresiones.toLocaleString("es-AR")}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{c.clics.toLocaleString("es-AR")}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{c.ctr?c.ctr+"%":"—"}</td>
                      <td style={{ padding:"9px 14px", color:C.soft }}>{c.resultado}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{c.roas?c.roas+"x":"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
            <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Gasto y resultados por género</div><BarList data={ads.genero.map(g=>({name:g.genero, value:g.resultados}))} color={C.lav}/></Card>
            <Card style={{ padding:20 }}><div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Gasto y resultados por edad</div><BarList data={ads.edad.map(e=>({name:e.edad, value:e.resultados}))} color={C.blue}/></Card>
          </div>
        </>
      )}

      {mp && (
        <>
          <SectionTitle icon={Wallet}>Mercado Pago</SectionTitle>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
            <Card style={{ padding:20 }}>
              <div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Comisión real sobre {ars(mp.bruto_mp)}</div>
              <div style={{ fontSize:12.5, color:C.soft, lineHeight:1.9 }}>
                Comisión: <strong style={{ color:C.red }}>−{ars(mp.comision_mp)}</strong><br/>
                IIBB: <strong style={{ color:C.red }}>−{ars(mp.iibb)}</strong><br/>
                Neto acreditado: <strong style={{ color:C.ink }}>{ars(mp.neto_acreditado)}</strong><br/>
                Descuento real: <strong>{mp.descuento_real_pct}%</strong>
              </div>
            </Card>
            <Card style={{ padding:20 }}>
              <div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>Movimientos sin confirmar</div>
              <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                {mp.movimientos_sin_confirmar.map((m,i)=>(
                  <div key={i} style={{ fontSize:12, display:"flex", justifyContent:"space-between", gap:8 }}>
                    <span style={{ color:C.soft }}>{m.desc} <span style={{ color:C.faint }}>({m.fecha})</span></span>
                    <span style={{ fontFamily:"monospace", color:m.monto<0?C.red:C.green, flexShrink:0 }}>{ars(m.monto)}</span>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        </>
      )}

      {audit && (
        <>
          <SectionTitle icon={Trophy}>Rentabilidad real (auditoría)</SectionTitle>
          <Card style={{ padding:20, marginBottom:14 }}>
            <div style={{ fontWeight:700, marginBottom:10, fontSize:13 }}>De la facturación al margen</div>
            <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
              {audit.waterfall.map((w,i)=>(
                <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12.5, fontWeight:w.total?800:400, color:w.total?C.ink:C.soft, borderTop:w.total?`1px solid ${C.border}`:"none", paddingTop:w.total?8:0 }}>
                  <span>{w.concepto}</span>
                  <span style={{ fontFamily:"monospace", color:w.monto<0?C.red:(w.total?C.green:C.ink) }}>{ars(w.monto)}</span>
                </div>
              ))}
            </div>
          </Card>
          <Card style={{ padding:0, overflow:"hidden" }}>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5 }}>
                <thead><tr style={{ textAlign:"left", color:C.faint, fontSize:11, textTransform:"uppercase" }}>
                  <th style={{ padding:"10px 14px" }}>Equipo</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Unidades</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Ingresos</th><th style={{ padding:"10px 14px", textAlign:"right" }}>Margen est.</th><th style={{ padding:"10px 14px", textAlign:"right" }}>% margen</th>
                </tr></thead>
                <tbody>
                  {audit.margen_por_equipo.map(m=>(
                    <tr key={m.equipo} style={{ borderTop:`1px solid ${C.border}` }}>
                      <td style={{ padding:"9px 14px" }}>{m.equipo}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{m.unidades}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{ars(m.ingresos)}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace", color:C.green }}>{ars(m.margen_est)}</td>
                      <td style={{ padding:"9px 14px", textAlign:"right", fontFamily:"monospace" }}>{m.margen_pct}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}

      {!!estadisticas?.ambiguedades_abiertas?.length && (
        <>
          <SectionTitle icon={AlertTriangle}>Para confirmar</SectionTitle>
          <Card style={{ padding:20 }}>
            <ul style={{ margin:0, paddingLeft:18, fontSize:12.5, color:C.soft, lineHeight:1.9 }}>
              {estadisticas.ambiguedades_abiertas.map((a,i)=><li key={i}>{a}</li>)}
            </ul>
          </Card>
        </>
      )}
    </div>
  );
}

// ====== STOCK ======
function StockAddForm({ onSave, onClose, notify }) {
  const [producto, setProducto] = useState("");
  const [costo, setCosto] = useState("");
  const [notas, setNotas] = useState("");
  const [rows, setRows] = useState([{talle:"M",cant:1},{talle:"L",cant:1}]);
  const setRow = (i,k,v) => setRows(r=>r.map((x,idx)=>idx===i?{...x,[k]:v}:x));
  const addRow = () => setRows(r=>[...r,{talle:"M",cant:1}]);
  const delRow = (i) => setRows(r=>r.length>1?r.filter((_,idx)=>idx!==i):r);
  return (
    <Modal wide title="Agregar modelo con varios talles" sub="Un modelo de camiseta · todos los talles y cantidades de una vez" onClose={onClose}>
      <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:12 }}>
        <Inp label="Producto / modelo *" value={producto} onChange={e=>setProducto(e.target.value)} placeholder="Ej: Argentina 2026 Titular" />
        <Inp label="Costo unitario" type="number" value={costo} onChange={e=>setCosto(e.target.value)} />
        <Inp span label="Notas (proveedor, versión...)" value={notas} onChange={e=>setNotas(e.target.value)} />
      </div>
      <Sec>Talles y cantidades</Sec>
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {rows.map((r,i)=>(
          <div key={i} style={{ display:"grid", gridTemplateColumns:"1fr 1fr auto", gap:8, alignItems:"center" }}>
            <select style={inputSt} value={r.talle} onChange={e=>setRow(i,"talle",e.target.value)}>{TALLES.filter(t=>t!=="—").map(t=><option key={t}>{t}</option>)}</select>
            <input style={inputSt} type="number" min="0" value={r.cant} onChange={e=>setRow(i,"cant",e.target.value)} placeholder="Cantidad" />
            <IconBtn icon={Trash2} onClick={()=>delRow(i)} color={C.red} bg={C.redSoft}/>
          </div>
        ))}
      </div>
      <Btn kind="soft" style={{ marginTop:10, padding:"7px 13px", fontSize:12.5 }} onClick={addRow}><Plus size={14}/>Agregar talle</Btn>
      <div style={{ display:"flex", gap:10, marginTop:18 }}>
        <Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!producto.trim()){notify("Falta el modelo","err");return;} const recs=rows.filter(r=>+r.cant>0); if(!recs.length){notify("Cargá al menos un talle con cantidad","err");return;} onSave(producto.trim(), +costo||0, notas, recs); }}>Guardar todo</Btn>
        <Btn kind="ghost" onClick={onClose}>Cancelar</Btn>
      </div>
    </Modal>
  );
}

function HistorialModal({ movs, onClose }) {
  const totalPedidoProv = movs.filter(m=>m.tipo==="ingreso" && /^proveedor/i.test(m.origen||"")).reduce((s,m)=>s+(+m.cantidad||0),0);
  const totalIngreso = movs.filter(m=>m.tipo==="ingreso").reduce((s,m)=>s+(+m.cantidad||0),0);
  const totalEgreso = movs.filter(m=>m.tipo==="egreso").reduce((s,m)=>s+(+m.cantidad||0),0);
  return (
    <Modal wide title="Historial de stock" sub="Todo lo que entró y salió, con su origen" onClose={onClose}>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:10, marginBottom:16 }}>
        <div style={{ background:C.violetSoft2, borderRadius:12, padding:14 }}><div style={{ fontSize:11, color:C.soft }}>Pedido a proveedores</div><div style={{ fontSize:22, fontWeight:800, color:C.violetDeep }}>{totalPedidoProv}</div></div>
        <div style={{ background:C.greenSoft, borderRadius:12, padding:14 }}><div style={{ fontSize:11, color:C.soft }}>Total ingresado</div><div style={{ fontSize:22, fontWeight:800, color:C.green }}>{totalIngreso}</div></div>
        <div style={{ background:C.redSoft, borderRadius:12, padding:14 }}><div style={{ fontSize:11, color:C.soft }}>Total egresado</div><div style={{ fontSize:22, fontWeight:800, color:C.red }}>{totalEgreso}</div></div>
        <div style={{ background:C.blueSoft, borderRadius:12, padding:14 }}><div style={{ fontSize:11, color:C.soft }}>Neto histórico</div><div style={{ fontSize:22, fontWeight:800, color:C.blue }}>{totalIngreso-totalEgreso}</div></div>
      </div>
      <div style={{ display:"flex", flexDirection:"column", gap:6, maxHeight:360, overflowY:"auto" }}>
        {movs.length===0 && <div style={{ fontSize:13, color:C.faint, textAlign:"center", padding:"20px 0" }}>Sin movimientos todavía.</div>}
        {movs.slice(0,250).map(m=>{ const ing=m.tipo==="ingreso"; return (
          <div key={m.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 12px", background:"#FAFAFD", borderRadius:9 }}>
            <span style={{ width:26, height:26, borderRadius:7, flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center", background:ing?C.greenSoft:C.redSoft }}>{ing?<Plus size={14} color={C.green}/>:<Minus size={14} color={C.red}/>}</span>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontSize:13, fontWeight:600, color:C.ink, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{m.producto}{m.talle&&m.talle!=="—"?` · ${m.talle}`:""}</div>
              <div style={{ fontSize:11, color:C.faint }}>{m.fecha} · {m.origen}</div>
            </div>
            <b style={{ color:ing?C.green:C.red, fontSize:14, flexShrink:0 }}>{ing?"+":"−"}{m.cantidad}</b>
          </div>
        );})}
      </div>
    </Modal>
  );
}

function Stock({ stock, setStock, prov, ventas, movs, setMovs, notify }) {
  const [editForm, setEditForm] = useState(null);
  const [addForm, setAddForm] = useState(false);
  const [hist, setHist] = useState(false);
  const [del, setDel] = useState(null);
  const [quickTalle, setQuickTalle] = useState({});
  const actual = s => (+s.ingresadas||0)-(+s.vendidas||0);
  const persist = async (l) => { const m=mergeStock(l); setStock(m); await save("tbx_stock",m); };
  const logMov = async (entries) => { const stamped=entries.map(e=>({ id:uid(), fecha:stamp(), ...e })); const l=[...stamped,...movs]; setMovs(l); await save("tbx_movs",l); };

  const guardarEdit = async (x) => {
    const old = stock.find(i=>i.id===x.id);
    const xx={...x,id:x.id||uid(),ingresadas:+x.ingresadas||0,vendidas:+x.vendidas||0,costo:+x.costo||0};
    await persist(x.id?stock.map(i=>i.id===x.id?xx:i):[xx,...stock]);
    if(old){ const d=xx.ingresadas-(+old.ingresadas||0); if(d>0) logMov([{producto:xx.producto,talle:xx.talle,tipo:"ingreso",cantidad:d,origen:"Edición manual"}]); else if(d<0) logMov([{producto:xx.producto,talle:xx.talle,tipo:"egreso",cantidad:-d,origen:"Edición manual"}]); }
    else logMov([{producto:xx.producto,talle:xx.talle,tipo:"ingreso",cantidad:xx.ingresadas,origen:"Carga manual"}]);
    setEditForm(null); notify("Guardado");
  };
  const guardarNuevo = async (producto, costo, notas, rows) => { const recs=rows.map(r=>({ id:uid(), producto, talle:r.talle, ingresadas:+r.cant, vendidas:0, costo, notas })); await persist([...recs, ...stock]); logMov(recs.map(r=>({producto, talle:r.talle, tipo:"ingreso", cantidad:r.ingresadas, origen:"Carga manual"}))); setAddForm(false); notify(`${recs.length} talle(s) agregados`); };
  const eliminar = async (id) => { await persist(stock.filter(i=>i.id!==id)); setDel(null); notify("Eliminado"); };
  const addOne  = async (id)=>{ const s=stock.find(x=>x.id===id); await persist(stock.map(x=>x.id===id?{...x,ingresadas:(+x.ingresadas||0)+1}:x)); if(s) logMov([{producto:s.producto,talle:s.talle,tipo:"ingreso",cantidad:1,origen:"Ajuste manual"}]); };
  const subOne  = async (id)=>{ const s=stock.find(x=>x.id===id); if(!s||(+s.ingresadas||0)<=(+s.vendidas||0)) return; await persist(stock.map(x=>x.id!==id?x:{...x,ingresadas:(+x.ingresadas||0)-1})); logMov([{producto:s.producto,talle:s.talle,tipo:"egreso",cantidad:1,origen:"Ajuste manual"}]); };
  const sellOne = async (id)=>{ const s=stock.find(x=>x.id===id); if(s && actual(s)<=0){ notify("Sin stock de ese talle","err"); return; } await persist(stock.map(x=>x.id===id?{...x,vendidas:(+x.vendidas||0)+1}:x)); if(s) logMov([{producto:s.producto,talle:s.talle,tipo:"egreso",cantidad:1,origen:"Venta rápida"}]); };
  const changeTalle = (id, t) => persist(stock.map(s=>s.id===id?{...s,talle:t}:s));

  const addSize = async (producto, talle, notas) => {
    const ex = stock.find(s=>norm(s.producto)===norm(producto) && (s.talle||"—")===talle);
    if(ex) await persist(stock.map(s=>s.id===ex.id?{...s,ingresadas:(+s.ingresadas||0)+1}:s));
    else await persist([{ id:uid(), producto, talle, ingresadas:1, vendidas:0, costo:0, notas:notas||"" }, ...stock]);
    logMov([{producto, talle, tipo:"ingreso", cantidad:1, origen:"Carga rápida"}]);
    notify(`+1 ${talle} en ${producto}`);
  };
  const sellSize = async (producto, talle) => {
    const ex = stock.find(s=>norm(s.producto)===norm(producto) && (s.talle||"—")===talle);
    if(!ex || actual(ex)<=0){ notify(`Sin stock de ${talle}`,"err"); return; }
    await persist(stock.map(s=>s.id===ex.id?{...s,vendidas:(+s.vendidas||0)+1}:s));
    logMov([{producto, talle, tipo:"egreso", cantidad:1, origen:"Venta rápida"}]);
    notify(`Vendiste 1 (${talle})`);
  };

  const importar = async () => {
    let l=[...stock]; const movEntries=[];
    prov.forEach(p=>(p.pedidos||[]).forEach(ped=>(ped.items||[]).forEach(it=>{
      const t=it.talle||"—";
      if(!l.some(s=>norm(s.producto)===norm(it.prod) && (s.talle||"—")===t)){ l.unshift({ id:uid(), producto:it.prod, talle:t, ingresadas:it.cant, vendidas:0, costo:0, notas:"De "+p.nombre }); movEntries.push({producto:it.prod, talle:t, tipo:"ingreso", cantidad:it.cant, origen:"Proveedor: "+p.nombre}); }
    })));
    await persist(l); if(movEntries.length) logMov(movEntries);
    notify(movEntries.length?`${movEntries.length} ítems importados de proveedores`:"Ya estaban todos cargados");
  };

  const compradoProv = prov.reduce((s,p)=>s+(p.pedidos||[]).reduce((a,ped)=>a+(ped.items||[]).reduce((b,it)=>b+(+it.cant||0),0),0),0);
  const ingTot = stock.reduce((s,x)=>s+(+x.ingresadas||0),0);
  const venTot = stock.reduce((s,x)=>s+(+x.vendidas||0),0);
  const actTot = stock.reduce((s,x)=>s+actual(x),0);
  const valor = stock.reduce((s,x)=>s+actual(x)*(+x.costo||0),0);

  const groups = []; const gmap = {};
  stock.forEach(s=>{ const k=norm(s.producto); if(!gmap[k]){ gmap[k]={ producto:s.producto, notas:s.notas, rows:[] }; groups.push(gmap[k]); } gmap[k].rows.push(s); });

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, gap:8, flexWrap:"wrap" }}>
        <Title>Stock interno</Title>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
          <Btn kind="ghost" onClick={()=>setHist(true)}><History size={15}/>Historial</Btn>
          <Btn kind="soft" onClick={importar}><RefreshCw size={15}/>Importar de proveedores</Btn>
          <Btn onClick={()=>setAddForm(true)}><Plus size={16}/>Agregar modelo</Btn>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:14, marginBottom:18 }}>
        <KPI icon={Package} label="Comprado a proveedores" value={compradoProv} sub="camisetas" color={C.violet} bg={C.violetSoft}/>
        <KPI icon={ShoppingBag} label="Ventas registradas" value={ventas.length} sub="reconciliación" color={C.orange} bg={C.orangeSoft}/>
        <KPI icon={Boxes} label="Stock actual" value={actTot} sub={`${ingTot} ingresadas − ${venTot} vendidas`} color={C.green} bg={C.greenSoft}/>
        <KPI icon={ArrowUpRight} label="Valorizado (al costo)" value={ars(valor)} color={C.blue} bg={C.blueSoft}/>
      </div>

      {stock.length===0 && <Card style={{ padding:"34px", textAlign:"center", color:C.faint, fontSize:14 }}>Sin stock. Importá de proveedores o agregá un modelo.</Card>}

      <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
        {groups.map(g=>{
          const totalG = g.rows.reduce((s,r)=>s+actual(r),0);
          const agotado = totalG<=0;
          const qt = quickTalle[g.producto] || "M";
          return (
            <Card key={g.producto} style={{ padding:"16px 18px" }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:10, marginBottom:12 }}>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                    <span style={{ fontWeight:700, color:C.ink, fontSize:15 }}>{g.producto}</span>
                    {agotado && <Badge color={C.red} bg={C.redSoft}><span style={{display:"inline-flex",alignItems:"center",gap:3}}><Tag size={11}/>SIN STOCK</span></Badge>}
                  </div>
                  {g.notas && <div style={{ fontSize:12, color:C.faint, marginTop:2 }}>{g.notas}</div>}
                  <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:8 }}>
                    {g.rows.map(r=>{ const a=actual(r); return (
                      <span key={r.id} style={{ fontSize:11.5, fontWeight:700, padding:"3px 9px", borderRadius:8, background:a>0?C.greenSoft:C.redSoft, color:a>0?C.green:C.red }}>{r.talle} · {a>0?a:"0"}</span>
                    );})}
                  </div>
                </div>
                <div style={{ textAlign:"right", flexShrink:0 }}>
                  <div style={{ fontSize:20, fontWeight:800, color:agotado?C.red:C.green }}>{totalG}</div>
                  <div style={{ fontSize:11, color:C.faint }}>en stock</div>
                </div>
              </div>

              <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                {g.rows.map(r=>{ const a=actual(r); return (
                  <div key={r.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 10px", background:"#FAFAFD", borderRadius:10, flexWrap:"wrap" }}>
                    <select value={r.talle} onChange={e=>changeTalle(r.id, e.target.value)} title="Cambiar talle" style={{ ...inputSt, width:62, padding:"5px 6px", fontSize:12.5, fontWeight:700, background:"#fff" }}>{TALLES.map(t=><option key={t}>{t}</option>)}</select>
                    <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                      <button onClick={()=>subOne(r.id)} style={{ width:26, height:26, borderRadius:7, border:`1px solid ${C.border}`, background:"#fff", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}><Minus size={13} color={C.soft}/></button>
                      <span style={{ minWidth:30, textAlign:"center", fontWeight:800, fontSize:14, color:a>0?C.green:C.red }}>{a}</span>
                      <button onClick={()=>addOne(r.id)} style={{ width:26, height:26, borderRadius:7, border:`1px solid ${C.border}`, background:"#fff", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}><Plus size={13} color={C.soft}/></button>
                    </div>
                    <Btn kind={a>0?"soft":"ghost"} disabled={a<=0} style={{ padding:"6px 11px", fontSize:12 }} onClick={()=>sellOne(r.id)}><ShoppingCart size={13}/>Vendí 1</Btn>
                    {a<=0 && <Badge color={C.red} bg={C.redSoft}>agotado</Badge>}
                    <div style={{ flex:1 }}/>
                    <span style={{ fontSize:11.5, color:C.faint }}>{r.ingresadas} ing · {r.vendidas} vend{r.costo?` · ${ars(a*r.costo)}`:""}</span>
                    <IconBtn icon={Pencil} onClick={()=>setEditForm(r)}/>
                    {del===r.id ? <><IconBtn icon={Check} onClick={()=>eliminar(r.id)} color={C.red} bg={C.redSoft}/><IconBtn icon={X} onClick={()=>setDel(null)}/></> : <IconBtn icon={Trash2} onClick={()=>setDel(r.id)} color={C.red}/>}
                  </div>
                );})}
              </div>

              <div style={{ display:"flex", alignItems:"center", gap:8, marginTop:12, paddingTop:12, borderTop:`1px dashed ${C.border}`, flexWrap:"wrap" }}>
                <span style={{ fontSize:12, color:C.faint, fontWeight:600 }}>Acción rápida:</span>
                <select value={qt} onChange={e=>setQuickTalle(q=>({...q,[g.producto]:e.target.value}))} style={{ ...inputSt, width:80, padding:"6px 8px", fontWeight:600 }}>{TALLES.filter(t=>t!=="—").map(t=><option key={t}>{t}</option>)}</select>
                <Btn kind="soft" style={{ padding:"6px 12px", fontSize:12 }} onClick={()=>addSize(g.producto, qt, g.notas)}><Plus size={13}/>Agregar stock</Btn>
                <Btn kind="ghost" style={{ padding:"6px 12px", fontSize:12 }} onClick={()=>sellSize(g.producto, qt)}><ShoppingCart size={13}/>Vendí una</Btn>
              </div>
            </Card>
          );
        })}
      </div>

      {addForm && <StockAddForm onSave={guardarNuevo} onClose={()=>setAddForm(false)} notify={notify} />}
      {hist && <HistorialModal movs={movs} onClose={()=>setHist(false)} />}
      {editForm!==null && (
        <Modal title="Editar talle" onClose={()=>setEditForm(null)}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
            <Inp span label="Producto *" value={editForm.producto||""} onChange={e=>setEditForm(p=>({...p,producto:e.target.value}))} />
            <Sel label="Talle" value={editForm.talle||"—"} options={TALLES} onChange={e=>setEditForm(p=>({...p,talle:e.target.value}))} />
            <Inp label="Costo unitario" type="number" value={editForm.costo||""} onChange={e=>setEditForm(p=>({...p,costo:e.target.value}))} />
            <Inp label="Ingresadas" type="number" value={editForm.ingresadas||""} onChange={e=>setEditForm(p=>({...p,ingresadas:e.target.value}))} />
            <Inp label="Vendidas" type="number" value={editForm.vendidas||""} onChange={e=>setEditForm(p=>({...p,vendidas:e.target.value}))} />
            <Inp span label="Notas" value={editForm.notas||""} onChange={e=>setEditForm(p=>({...p,notas:e.target.value}))} />
          </div>
          <div style={{ display:"flex", gap:10, marginTop:18 }}><Btn style={{ flex:1, justifyContent:"center" }} onClick={()=>{ if(!editForm.producto){notify("Falta el producto","err");return;} guardarEdit(editForm); }}>Guardar</Btn><Btn kind="ghost" onClick={()=>setEditForm(null)}>Cancelar</Btn></div>
        </Modal>
      )}
    </div>
  );
}

// ====== CONFIG ======
function Config({ rates, setRates, notify }) {
  const [r, setR] = useState(rates);
  const guardar = async () => { setRates(r); await save("tbx_rates", r); notify("Tasas actualizadas"); };
  const campos = [["mp","Cargo Mercado Pago %"],["cpt","Cargo plataforma / CPT %"],["iibb","Retención IIBB (SIRTAC) %"],["c3","Interés 3 cuotas sin interés %"]];
  return (
    <div>
      <Title>Configuración de comisiones</Title>
      <p style={{ color:C.soft, fontSize:13, margin:"8px 0 18px" }}>Estos % autocompletan las comisiones de cada venta. Ajustalos cuando cambien.</p>
      <Card style={{ padding:22, maxWidth:420 }}>
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>{campos.map(([k,l])=>(<Inp key={k} label={l} type="number" value={r[k]} onChange={e=>setR(p=>({...p,[k]:+e.target.value}))} />))}</div>
        <Btn style={{ marginTop:16, width:"100%", justifyContent:"center" }} onClick={guardar}>Guardar tasas</Btn>
      </Card>
    </div>
  );
}

// ====== APP ======
const NAV = [
  { id:"resumen", label:"Resumen", icon:LayoutGrid },
  { id:"estadisticas", label:"Estadísticas", icon:BarChart3 },
  { id:"ventas", label:"Ventas", icon:ShoppingBag },
  { id:"egresos", label:"Egresos", icon:TrendingDown },
  { id:"proveedores", label:"Proveedores", icon:Package },
  { id:"stock", label:"Stock", icon:Boxes },
  { id:"clientes", label:"Clientes", icon:Users },
  { id:"importar", label:"Importar / Exportar", icon:Database },
  { id:"config", label:"Configuración", icon:Settings },
];

export default function App() {
  const [tab, setTab] = useState("resumen");
  const [query, setQuery] = useState("");
  const [ventas, setVentas] = useState([]);
  const [egresos, setEgresos] = useState([]);
  const [prov, setProv] = useState([]);
  const [stock, setStock] = useState([]);
  const [clientes, setClientes] = useState([]);
  const [movs, setMovs] = useState([]);
  const [rates, setRates] = useState(DEFAULT_RATES);
  const [estadisticas, setEstadisticas] = useState(null);
  const [ready, setReady] = useState(false);
  const [toast, setToast] = useState(null);
  const notify = (msg,type) => { setToast({msg,type}); setTimeout(()=>setToast(null),2600); };

  const applyBackup = async (d) => {
    if(d.ventas){ setVentas(d.ventas); await save("tbx_ventas", d.ventas); }
    if(d.egresos){ setEgresos(d.egresos); await save("tbx_egresos", d.egresos); }
    if(d.prov){ setProv(d.prov); await save("tbx_proveedores", d.prov); }
    if(d.stock){ setStock(d.stock); await save("tbx_stock", d.stock); }
    if(d.clientes){ setClientes(d.clientes); await save("tbx_clientes", d.clientes); }
    if(d.movs){ setMovs(d.movs); await save("tbx_movs", d.movs); }
    if(d.rates){ setRates(d.rates); await save("tbx_rates", d.rates); }
  };

  useEffect(()=>{
    const l=document.createElement("link"); l.rel="stylesheet"; l.href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap"; document.head.appendChild(l);
    const rts=DEFAULT_RATES;
    (async()=>{
      try{
        const [v,e,p,st,cl,mv,rt,es]=await Promise.all([load("tbx_ventas",[]),load("tbx_egresos",[]),load("tbx_proveedores",[]),load("tbx_stock",[]),load("tbx_clientes",[]),load("tbx_movs",[]),load("tbx_rates",rts),load("tbx_estadisticas",null)]);
        if(es) setEstadisticas(es);
        const rFinal = rt || rts;
        if(Array.isArray(v)) setVentas(v.map(x=>computeVenta(migrateVenta(x), rFinal)));
        if(Array.isArray(e)) setEgresos(e);
        if(Array.isArray(p)) setProv(p.map(pp=>({ ...pp, pedidos:(pp.pedidos||[]).map(ped=>({ id:ped.id||uid(), ...ped, items:(ped.items||[]).map(it=>({ talle:"—", ...it })) })) })));
        if(Array.isArray(st)) setStock(mergeStock(st));
        if(Array.isArray(cl)) setClientes(cl);
        if(Array.isArray(mv)) setMovs(mv);
        setRates(rFinal);
      }catch(err){}
      setReady(true);
    })();
  },[]);

  if(!ready) return <div style={{ background:C.bg, minHeight:500, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:FONT, color:C.violet, fontWeight:700 }}>Cargando…</div>;

  return (
    <div style={{ background:C.bg, fontFamily:FONT, minHeight:640, display:"flex", borderRadius:18, overflow:"hidden" }}>
      <Toast t={toast}/>
      <div style={{ width:230, background:C.card, borderRight:`1px solid ${C.border}`, padding:"22px 16px", flexShrink:0 }}>
        <div style={{ display:"flex", alignItems:"center", gap:9, padding:"0 8px 22px" }}>
          <div style={{ width:34, height:34, borderRadius:10, background:`linear-gradient(135deg,${C.violet},${C.lav})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:800 }}>T</div>
          <div style={{ fontWeight:800, fontSize:17, color:C.ink }}>Tribuneros</div>
        </div>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>{ setTab(n.id); if(n.id!=="ventas") setQuery(""); }} style={{ width:"100%", display:"flex", alignItems:"center", gap:11, padding:"10px 12px", marginBottom:3, border:"none", borderRadius:11, cursor:"pointer", fontFamily:FONT, fontSize:13.5, fontWeight:600, background:tab===n.id?C.violetSoft:"transparent", color:tab===n.id?C.violetDeep:C.soft }}>
            <n.icon size={18}/>{n.label}
          </button>
        ))}
      </div>
      <div style={{ flex:1, overflowY:"auto", maxHeight:"100vh" }}>
        <div style={{ padding:"20px 28px", borderBottom:`1px solid ${C.border}`, background:"rgba(255,255,255,0.6)", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontSize:18, fontWeight:800, color:C.ink }}>{NAV.find(n=>n.id===tab)?.label}</div>
            <div style={{ fontSize:12.5, color:C.faint }}>Gestión y control de e-commerce</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8, background:"#F2F2F7", borderRadius:12, padding:"8px 14px" }}>
            <Search size={15} color={C.faint}/>
            <input value={query} onChange={e=>{ setQuery(e.target.value); if(e.target.value) setTab("ventas"); }} placeholder="Buscar cliente, producto, pedido…" style={{ border:"none", background:"transparent", outline:"none", fontFamily:FONT, fontSize:13, color:C.ink, width:210 }} />
            {query && <button onClick={()=>setQuery("")} style={{ border:"none", background:"transparent", cursor:"pointer", display:"flex" }}><X size={14} color={C.faint}/></button>}
          </div>
        </div>
        <div style={{ padding:28 }}>
          {tab==="resumen" && <Dashboard ventas={ventas} egresos={egresos} estadisticas={estadisticas}/>}
          {tab==="estadisticas" && <Estadisticas ventas={ventas} estadisticas={estadisticas}/>}
          {tab==="ventas" && <Ventas ventas={ventas} setVentas={setVentas} rates={rates} notify={notify} query={query}/>}
          {tab==="egresos" && <Egresos egresos={egresos} setEgresos={setEgresos} notify={notify}/>}
          {tab==="proveedores" && <Proveedores prov={prov} setProv={setProv} stock={stock} setStock={setStock} movs={movs} setMovs={setMovs} notify={notify}/>}
          {tab==="stock" && <Stock stock={stock} setStock={setStock} prov={prov} ventas={ventas} movs={movs} setMovs={setMovs} notify={notify}/>}
          {tab==="clientes" && <Clientes ventas={ventas} clientes={clientes} setClientes={setClientes} notify={notify}/>}
          {tab==="importar" && <Importar data={{ventas,egresos,prov,stock,clientes,movs,rates}} apply={applyBackup} rates={rates} notify={notify}/>}
          {tab==="config" && <Config rates={rates} setRates={setRates} notify={notify}/>}
        </div>
      </div>
    </div>
  );
}