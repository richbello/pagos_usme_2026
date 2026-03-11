import { useState, useRef } from "react";

const API = "http://localhost:3001/api";

const MODULES = [
  {
    id: 1, title: "Extracción Cuentas Bancarias",
    subtitle: "Extraccion_ctas_bancarias_desdepdf.py",
    color: "#1d4ed8", colorDark: "#1e3a8a", emoji: "📄",
    description: "Extrae datos de comprobantes de pago desde PDFs y genera un consolidado Excel con 27 campos.",
    fileInputs: [
      { key: "CARPETA_PDF", label: "📁 Archivos PDF (comprobantes de pago)", accept: ".pdf", multiple: true },
    ],
    textInputs: [
      { key: "NOMBRE_EXCEL", label: "Nombre del Excel de salida", placeholder: "Consolidado_Pagos_Lote1.xlsx" },
    ],
    outputs: ["Archivo PDF","Ciudad y Fecha","Documento No.","Nombre Contratista","NIT Entidad","Cédula","Por Concepto","Periodo Desde/Hasta","Banco","No. Cuenta","La Suma De ($)","Pago No."],
    steps: ["Selecciona todos los PDFs de comprobantes","Escribe el nombre del Excel de salida","Clic en Ejecutar Módulo","Descarga el Excel generado"],
  },
  {
    id: 2, title: "Extraer Datos DEFFEB21 PDF",
    subtitle: "extraer-datos_DEFFEB21_pdf.py",
    color: "#7c3aed", colorDark: "#4c1d95", emoji: "🗄️",
    description: "Procesa PDFs de causación extrayendo montos, retenciones, fechas y datos de contratos.",
    fileInputs: [
      { key: "ruta_carpeta", label: "📁 Archivos PDF (causaciones)", accept: ".pdf", multiple: true },
    ],
    textInputs: [
      { key: "SALIDA", label: "Nombre del Excel de salida", placeholder: "Extraccion-grupo3_feb.xlsx" },
    ],
    outputs: ["Archivo","Contrato CPS","Contratista","NIT/CC","PAGO No.","PERIODO","DEL/AL","Valor Bruto","BASE RETEICA","Neto a Pagar"],
    steps: ["Selecciona los PDFs de causación","Escribe el nombre del Excel de salida","Clic en Ejecutar Módulo","Descarga el Excel generado"],
  },
  {
    id: 3, title: "Plantilla Pagos DeepSeek",
    subtitle: "plantilla_pagos_deepseek.py",
    color: "#059669", colorDark: "#064e3b", emoji: "📊",
    description: "Genera la plantilla SAP con bloques C, P40 y P31 a partir del consolidado de extracción.",
    fileInputs: [
      { key: "ruta_entrada", label: "📄 Consolidado de entrada (.xlsx)", accept: ".xlsx,.xls", multiple: false },
    ],
    textInputs: [
      { key: "SALIDA", label: "Nombre del archivo de salida", placeholder: "PLANTILLA_PAGOS_GENERADAFEB.xlsx" },
    ],
    outputs: ["Tipo Registro C/P40/P31","Clave Contab.","No Identificación","Importe","RP Doc Presupuestal","Asignación","Texto Pago No.","Indicador RETEICA","Base imponible","Importe retención"],
    steps: ["Sube el Excel generado por el módulo 2","Escribe el nombre del archivo de salida","Clic en Ejecutar Módulo","Descarga la plantilla SAP"],
  },
  {
    id: 4, title: "Planilla CRP PAC",
    subtitle: "planilla_crp_pac.py",
    color: "#ea580c", colorDark: "#7f1d1d", emoji: "📈",
    description: "Cruza pagos con CRP y PAC para verificar disponibilidad presupuestal por rubro y fondo.",
    fileInputs: [
      { key: "ruta_pagos", label: "📄 Plantilla de Pagos (.xlsx)",  accept: ".xlsx,.xls", multiple: false },
      { key: "ruta_crp",   label: "📄 Reporte CRP (.xlsx)",         accept: ".xlsx,.xls", multiple: false },
      { key: "ruta_pac",   label: "📄 Reporte PAC (.xlsx)",          accept: ".xlsx,.xls", multiple: false },
    ],
    textInputs: [
      { key: "SALIDA", label: "Nombre del archivo de resultado", placeholder: "RESULTADO_FEB22.xlsx" },
    ],
    outputs: ["CRP por pago","Rubro y Fondo","Total a Pagar","Disponibilidad PAC","ALCANZA/NO ALCANZA","PAC otros meses","Total General","Excel con colores de alerta"],
    steps: ["Sube la Plantilla de Pagos","Sube el Reporte CRP","Sube el Reporte PAC","Escribe el nombre del resultado","Clic en Ejecutar Módulo","Descarga el Excel de análisis"],
  },
];

async function uploadFiles(files) {
  const form = new FormData();
  for (const f of files) form.append("files", f);
  const res = await fetch(`${API}/upload`, { method: "POST", body: form });
  return res.json();
}

async function runScript(scriptName, params, onLog) {
  onLog(`▶ Ejecutando: ${scriptName}\n${"─".repeat(50)}\n`);
  const res = await fetch(`${API}/run`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ scriptName, params }),
  });
  const reader  = res.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let fullLog   = "";
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = decoder.decode(value);
    fullLog += chunk;
    onLog(chunk);
  }
  const match = fullLog.match(/__EXCEL__:(.+)/);
  if (match) return match[1].split(",").map(s => s.trim()).filter(Boolean);
  return [];
}

export default function App() {
  const [active,        setActive]        = useState(1);
  const [tab,           setTab]           = useState("upload");
  const [logs,          setLogs]          = useState("");
  const [running,       setRunning]       = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState({});
  const [textParams,    setTextParams]    = useState({});
  const [serverOk,      setServerOk]      = useState(null);
  const [excelFiles,    setExcelFiles]    = useState([]);
  const [copied,        setCopied]        = useState(false);
  const logsRef = useRef(null);
  const mod = MODULES.find(m => m.id === active);

  const addLog = (txt) => {
    setLogs(prev => {
      const next = prev + txt;
      setTimeout(() => { if (logsRef.current) logsRef.current.scrollTop = logsRef.current.scrollHeight; }, 50);
      return next;
    });
  };

  const checkServer = async () => {
    try {
      await fetch(`${API}/excels`, { signal: AbortSignal.timeout(2000) });
      setServerOk(true);
    } catch { setServerOk(false); }
  };

  const handleSelect = (id) => {
    setActive(id); setTab("upload"); setLogs("");
    setUploadedFiles({}); setTextParams({}); setExcelFiles([]);
    checkServer();
  };

  const handleFileChange = (key, files) =>
    setUploadedFiles(prev => ({ ...prev, [key]: Array.from(files) }));

  // Construir params comunes (texto)
  const buildParams = (pathMap = {}) => {
    const params = {};
    Object.entries(uploadedFiles).forEach(([key, files]) => {
      if (files.length === 0) return;
      if (files.length === 1) {
        const nombre = files[0].name;
        params[key] = pathMap[nombre] || `uploads_temp/${nombre}`;
      } else {
        params[key] = "uploads_temp";
        if (key === "CARPETA_PDF") { params["CARPETA_PDF"] = "uploads_temp"; params["CARPETA_EXCEL"] = "uploads_temp"; }
        if (key === "ruta_carpeta") { params["ruta_carpeta"] = "uploads_temp"; }
      }
    });
    Object.entries(textParams).forEach(([k, v]) => {
      if (k === "SALIDA" && v) {
        if (mod.id === 1) params["NOMBRE_EXCEL"] = v;
        else if (mod.id === 2) params["archivo_salida"] = `uploads_temp/${v}`;
        else if (mod.id === 3) params["ruta_destino"] = `uploads_temp/${v}`;
        else if (mod.id === 4) params["ruta_salida"]  = `uploads_temp/${v}`;
      } else { params[k] = v; }
    });
    return params;
  };

  // Ejecutar CON subida de archivos
  const handleRun = async () => {
    setRunning(true); setLogs(""); setExcelFiles([]); setTab("logs");
    try {
      const allFiles = Object.values(uploadedFiles).flat();
      let pathMap = {};
      if (allFiles.length > 0) {
        addLog(`📤 Subiendo ${allFiles.length} archivo(s)...\n`);
        const result = await uploadFiles(allFiles);
        if (result.ok) {
          addLog(`✅ Archivos listos en servidor\n\n`);
          (result.rutas || []).forEach(ruta => {
            const nombre = ruta.split(/[/\\]/).pop();
            pathMap[nombre] = ruta;
          });
        } else {
          addLog(`❌ Error subiendo archivos: ${result.error}\n`);
        }
      }
      const params = buildParams(pathMap);
      const generados = await runScript(mod.subtitle, params, addLog);
      if (generados && generados.length > 0) {
        setExcelFiles(generados);
        addLog(`\n${"─".repeat(50)}\n📊 ${generados.length} archivo(s) listo(s) para descargar\n`);
        setTab("download");
      }
    } catch (err) {
      addLog(`\n❌ Error: ${err.message}\n`);
      addLog(`💡 ¿Está corriendo el servidor? Ejecuta: node server.js\n`);
    }
    setRunning(false);
  };

  // Ejecutar SIN subir — usa archivos ya en uploads_temp
  const handleRunExisting = async () => {
    setRunning(true); setLogs(""); setExcelFiles([]); setTab("logs");
    try {
      addLog(`♻️ Usando archivos existentes en uploads_temp...\n\n`);
      const params = buildParams();
      // Para módulos que usan carpeta
      if (mod.id === 1) { params["CARPETA_PDF"] = "uploads_temp"; params["CARPETA_EXCEL"] = "uploads_temp"; }
      if (mod.id === 2) { params["ruta_carpeta"] = "uploads_temp"; }
      const generados = await runScript(mod.subtitle, params, addLog);
      if (generados && generados.length > 0) {
        setExcelFiles(generados);
        addLog(`\n${"─".repeat(50)}\n📊 ${generados.length} archivo(s) listo(s) para descargar\n`);
        setTab("download");
      }
    } catch (err) {
      addLog(`\n❌ Error: ${err.message}\n`);
      addLog(`💡 ¿Está corriendo el servidor? Ejecuta: node server.js\n`);
    }
    setRunning(false);
  };

  const handleDownload = (filename) => {
    window.open(`${API}/download/${encodeURIComponent(filename)}`, "_blank");
  };

  const doCopy = () => {
    navigator.clipboard.writeText(`python ${mod.subtitle}`);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const totalFiles = Object.values(uploadedFiles).flat().length;

  const tabList = [
    ["upload",   "📁 Cargar Archivos"],
    ["outputs",  "📤 Salidas"],
    ["steps",    "📋 Pasos"],
    ["logs",     `🖥️ Consola${logs ? " ●" : ""}`],
    ["download", `⬇️ Descargar${excelFiles.length > 0 ? ` (${excelFiles.length})` : ""}`],
  ];

  return (
    <div style={{ fontFamily: "'Segoe UI',sans-serif", minHeight: "100vh", background: "#f1f5f9" }}>

      {/* NAVBAR */}
      <div style={{ background: "linear-gradient(90deg,#0f172a,#1e3a5f,#0f172a)", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", boxShadow: "0 2px 16px rgba(0,0,0,0.5)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 26 }}>🏛️</span>
          <div>
            <div style={{ color: "#fff", fontWeight: 800, fontSize: 16 }}>FDL USME 2026</div>
            <div style={{ color: "#93c5fd", fontSize: 12 }}>Sistema Gestión de Pagos · 4 Módulos</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {serverOk === true  && <span style={{ fontSize: 11, color: "#4ade80", background: "#052e16", padding: "3px 10px", borderRadius: 20 }}>🟢 Servidor activo</span>}
          {serverOk === false && <span style={{ fontSize: 11, color: "#fca5a5", background: "#450a0a", padding: "3px 10px", borderRadius: 20 }}>🔴 Ejecuta: node server.js</span>}
          <div style={{ display: "flex", gap: 8 }}>
            {MODULES.map(m => (
              <button key={m.id} onClick={() => handleSelect(m.id)} style={{
                width: 34, height: 34, borderRadius: "50%", border: "none", cursor: "pointer",
                fontWeight: 800, fontSize: 14,
                background: active === m.id ? m.color : "#1e293b",
                color: active === m.id ? "#fff" : "#94a3b8",
                transform: active === m.id ? "scale(1.18)" : "scale(1)",
                transition: "all 0.2s",
                boxShadow: active === m.id ? `0 0 12px ${m.color}99` : "none",
              }}>{m.id}</button>
            ))}
          </div>
        </div>
      </div>

      {/* BODY */}
      <div style={{ maxWidth: 1200, margin: "0 auto", padding: 20, display: "grid", gridTemplateColumns: "300px 1fr", gap: 20 }}>

        {/* LEFT */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {MODULES.map(m => (
            <div key={m.id} onClick={() => handleSelect(m.id)} style={{
              borderRadius: 14, border: `2px solid ${active === m.id ? m.color : "#e2e8f0"}`,
              boxShadow: active === m.id ? `0 0 0 3px ${m.color}33` : "0 1px 4px rgba(0,0,0,0.07)",
              cursor: "pointer", overflow: "hidden", background: "#fff", transition: "all 0.2s",
            }}>
              <div style={{ background: `linear-gradient(135deg,${m.color},${m.colorDark})`, padding: "11px 14px", display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 18, background: "rgba(255,255,255,0.18)", borderRadius: 8, padding: "2px 7px" }}>{m.emoji}</span>
                <div style={{ flex: 1 }}>
                  <div style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>{m.title}</div>
                  <div style={{ color: "rgba(255,255,255,0.6)", fontSize: 10 }}>{m.subtitle}</div>
                </div>
                <span style={{ background: "rgba(255,255,255,0.18)", color: "#fff", fontSize: 10, fontWeight: 700, borderRadius: 20, padding: "2px 8px" }}>#{m.id}</span>
              </div>
              <div style={{ padding: "9px 13px" }}>
                <div style={{ color: "#64748b", fontSize: 11, lineHeight: 1.5 }}>{m.description}</div>
              </div>
            </div>
          ))}
          {/* Pipeline */}
          <div style={{ background: "linear-gradient(135deg,#0f172a,#1e293b)", borderRadius: 14, padding: 18, color: "#fff" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 2, textTransform: "uppercase", marginBottom: 14 }}>🔄 Flujo de Trabajo</div>
            {[{id:1,l:"PDFs → Cuentas Bancarias",c:"#1d4ed8"},{id:2,l:"PDFs → Datos DEFFEB21",c:"#7c3aed"},{id:3,l:"Consolidado → Plantilla SAP",c:"#059669"},{id:4,l:"Plantilla + CRP + PAC",c:"#ea580c"}]
              .map((item,i,arr)=>(
              <div key={item.id}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ background:item.c, width:28, height:28, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:800, flexShrink:0 }}>{item.id}</div>
                  <div style={{ fontSize:12, fontWeight:600 }}>{item.l}</div>
                </div>
                {i<arr.length-1&&<div style={{ paddingLeft:40, color:"#475569", fontSize:16, lineHeight:1, margin:"3px 0" }}>↓</div>}
              </div>
            ))}
            <div style={{ marginTop:14, paddingTop:10, borderTop:"1px solid #1e293b", fontSize:10, color:"#64748b" }}>⚠️ Ejecutar en orden: 1 → 2 → 3 → 4</div>
          </div>
        </div>

        {/* RIGHT */}
        <div style={{ background: "#fff", borderRadius: 16, border: "1px solid #e2e8f0", boxShadow: "0 2px 12px rgba(0,0,0,0.08)", overflow: "hidden", display: "flex", flexDirection: "column" }}>

          {/* Header */}
          <div style={{ background:`linear-gradient(135deg,${mod.color},${mod.colorDark})`, padding:"20px 24px" }}>
            <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:8 }}>
              <span style={{ fontSize:28, background:"rgba(255,255,255,0.18)", borderRadius:12, padding:"5px 10px" }}>{mod.emoji}</span>
              <div>
                <div style={{ color:"#fff", fontWeight:800, fontSize:20 }}>{mod.title}</div>
                <div style={{ color:"rgba(255,255,255,0.6)", fontSize:12 }}>{mod.subtitle}</div>
              </div>
            </div>
            <div style={{ color:"rgba(255,255,255,0.88)", fontSize:13, lineHeight:1.6 }}>{mod.description}</div>
          </div>

          {/* Tabs */}
          <div style={{ display:"flex", borderBottom:"1px solid #f1f5f9", background:"#f8fafc", flexWrap:"wrap" }}>
            {tabList.map(([key,label])=>(
              <button key={key} onClick={()=>setTab(key)} style={{
                flex:1, padding:"11px 6px", fontSize:11, fontWeight:600, border:"none", cursor:"pointer",
                background: tab===key ? `linear-gradient(135deg,${mod.color},${mod.colorDark})` : "transparent",
                color: tab===key ? "#fff" : "#64748b",
                transition:"all 0.15s", whiteSpace:"nowrap",
              }}>{label}</button>
            ))}
          </div>

          {/* Content */}
          <div style={{ padding:24, flex:1, overflowY:"auto" }}>

            {/* UPLOAD */}
            {tab==="upload" && (
              <div>
                <div style={{ fontSize:12, color:"#94a3b8", marginBottom:18 }}>
                  📌 Selecciona los archivos y haz clic en <strong>Ejecutar Módulo</strong>.
                  {serverOk===false && <div style={{ marginTop:8, padding:"8px 12px", background:"#fef2f2", borderRadius:8, color:"#ef4444", fontSize:11 }}>⚠️ Servidor inactivo. Abre CMD y ejecuta: <code>node server.js</code></div>}
                </div>

                {mod.fileInputs.map(inp => {
                  const files = uploadedFiles[inp.key] || [];
                  return (
                    <div key={inp.key} style={{ marginBottom:18, padding:16, background:"#f8fafc", borderRadius:12, border:`2px dashed ${files.length>0?mod.color:"#e2e8f0"}`, transition:"border 0.2s" }}>
                      <div style={{ fontSize:12, fontWeight:700, color:"#475569", marginBottom:10 }}>{inp.label}</div>
                      <label style={{ display:"inline-flex", alignItems:"center", gap:8, cursor:"pointer", background:mod.color, color:"#fff", padding:"9px 18px", borderRadius:8, fontSize:12, fontWeight:700, userSelect:"none" }}>
                        {inp.multiple ? "📂 Seleccionar archivos" : "📄 Seleccionar archivo"}
                        <input type="file" accept={inp.accept} multiple={!!inp.multiple} style={{ display:"none" }}
                          onChange={e => handleFileChange(inp.key, e.target.files)} />
                      </label>
                      {files.length>0 && (
                        <div style={{ marginTop:10 }}>
                          <div style={{ fontSize:11, color:"#16a34a", fontWeight:700, marginBottom:4 }}>✅ {files.length} archivo(s) seleccionado(s)</div>
                          <div style={{ maxHeight:80, overflowY:"auto" }}>
                            {files.slice(0,6).map((f,i)=><div key={i} style={{ fontSize:11, color:"#475569", padding:"1px 0" }}>📎 {f.name}</div>)}
                            {files.length>6 && <div style={{ fontSize:11, color:"#94a3b8" }}>...y {files.length-6} más</div>}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}

                {mod.textInputs && mod.textInputs.map(inp=>(
                  <div key={inp.key} style={{ marginBottom:16 }}>
                    <label style={{ display:"block", fontSize:12, fontWeight:600, color:"#475569", marginBottom:6 }}>{inp.label}</label>
                    {inp.type==="checkbox" ? (
                      <label style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer" }}>
                        <input type="checkbox" checked={!!textParams[inp.key]}
                          onChange={e=>setTextParams({...textParams,[inp.key]:e.target.checked})}
                          style={{ width:16, height:16 }} />
                        <span style={{ fontSize:12, color:"#64748b" }}>Activar modo debug</span>
                      </label>
                    ) : (
                      <input type="text" value={textParams[inp.key]||""} placeholder={inp.placeholder}
                        onChange={e=>setTextParams({...textParams,[inp.key]:e.target.value})}
                        style={{ width:"100%", padding:"9px 13px", border:"2px solid #e2e8f0", borderRadius:8, fontSize:12, fontFamily:"monospace", color:"#334155", background:"#f8fafc", boxSizing:"border-box", outline:"none" }} />
                    )}
                  </div>
                ))}

                {/* BOTONES */}
                <div style={{ marginTop:24, display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
                  {/* Botón principal: subir + ejecutar */}
                  <button onClick={handleRun} disabled={running} style={{
                    display:"flex", alignItems:"center", gap:10,
                    background: running ? "#94a3b8" : `linear-gradient(135deg,${mod.color},${mod.colorDark})`,
                    color:"#fff", border:"none", borderRadius:10, padding:"13px 28px",
                    fontSize:14, fontWeight:800, cursor: running?"not-allowed":"pointer",
                    boxShadow: running?"none":`0 4px 14px ${mod.color}66`, transition:"all 0.2s",
                  }}>
                    {running ? "⏳ Ejecutando..." : `▶ Ejecutar Módulo ${mod.id}`}
                  </button>

                  {/* Botón secundario: usar archivos existentes */}
                  <button onClick={handleRunExisting} disabled={running} title="Ejecuta el script usando los archivos que ya están en uploads_temp, sin subir nada nuevo" style={{
                    display:"flex", alignItems:"center", gap:8,
                    background: running ? "#e2e8f0" : "#f1f5f9",
                    color: running ? "#94a3b8" : "#475569",
                    border:`2px solid ${running ? "#e2e8f0" : "#cbd5e1"}`,
                    borderRadius:10, padding:"11px 20px",
                    fontSize:13, fontWeight:700, cursor: running?"not-allowed":"pointer",
                    transition:"all 0.2s",
                  }}>
                    ♻️ Usar archivos existentes
                  </button>

                  {totalFiles>0 && <span style={{ fontSize:12, color:"#64748b" }}>📎 {totalFiles} archivo(s) listos</span>}
                  {excelFiles.length>0 && (
                    <button onClick={()=>setTab("download")} style={{ display:"flex", alignItems:"center", gap:6, background:"#dcfce7", color:"#16a34a", border:"2px solid #16a34a", borderRadius:10, padding:"11px 20px", fontSize:13, fontWeight:800, cursor:"pointer" }}>
                      ⬇️ Descargar Excel ({excelFiles.length})
                    </button>
                  )}
                </div>

                {/* Nota explicativa */}
                <div style={{ marginTop:12, padding:"10px 14px", background:"#f0f9ff", borderRadius:8, border:"1px solid #bae6fd", fontSize:11, color:"#0369a1" }}>
                  💡 <strong>¿Archivos muy pesados?</strong> Usa <strong>♻️ Usar archivos existentes</strong> si ya copiaste los archivos directamente a <code>uploads_temp</code> — ejecuta el script sin intentar subirlos por el navegador.
                </div>

                <div style={{ marginTop:14, padding:12, background:"#f1f5f9", borderRadius:10, borderLeft:`4px solid ${mod.color}` }}>
                  <div style={{ fontSize:11, color:"#64748b", marginBottom:6, fontWeight:700 }}>💡 O ejecuta manualmente en CMD:</div>
                  <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                    <code style={{ fontSize:12, color:"#1e293b", fontFamily:"monospace", flex:1 }}>python {mod.subtitle}</code>
                    <button onClick={doCopy} style={{ fontSize:11, padding:"5px 12px", background:copied?"#dcfce7":"#e2e8f0", color:copied?"#16a34a":"#475569", border:"none", borderRadius:6, cursor:"pointer", fontWeight:700 }}>
                      {copied?"✅ Copiado":"📋 Copiar"}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* OUTPUTS */}
            {tab==="outputs" && (
              <div>
                <div style={{ fontSize:12, color:"#94a3b8", marginBottom:16 }}>Campos que genera este módulo en el Excel de salida:</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
                  {mod.outputs.map((o,i)=>(
                    <span key={i} style={{ fontSize:11, padding:"5px 13px", borderRadius:20, fontWeight:600, background:mod.color+"18", color:mod.color, border:`1px solid ${mod.color}44` }}>{o}</span>
                  ))}
                </div>
              </div>
            )}

            {/* STEPS */}
            {tab==="steps" && (
              <div>
                <div style={{ fontSize:12, color:"#94a3b8", marginBottom:16 }}>Sigue estos pasos para ejecutar el módulo:</div>
                {mod.steps.map((step,i)=>(
                  <div key={i} style={{ display:"flex", alignItems:"flex-start", gap:12, marginBottom:16 }}>
                    <div style={{ background:mod.color, color:"#fff", fontSize:12, fontWeight:700, width:28, height:28, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, marginTop:1 }}>{i+1}</div>
                    <span style={{ fontSize:13, color:"#374151", lineHeight:1.6 }}>{step}</span>
                  </div>
                ))}
              </div>
            )}

            {/* LOGS */}
            {tab==="logs" && (
              <div>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                  <div style={{ fontSize:12, color:"#94a3b8" }}>🖥️ Salida en tiempo real del script Python</div>
                  <button onClick={()=>setLogs("")} style={{ fontSize:11, padding:"5px 12px", background:"#f1f5f9", color:"#64748b", border:"none", borderRadius:6, cursor:"pointer", fontWeight:700 }}>🗑️ Limpiar</button>
                </div>
                <div ref={logsRef} style={{ background:"#0f172a", color:"#e2e8f0", fontSize:12, borderRadius:12, padding:18, minHeight:300, maxHeight:450, overflowY:"auto", fontFamily:"monospace", whiteSpace:"pre-wrap", lineHeight:1.7 }}>
                  {logs || <span style={{ color:"#475569" }}>Aquí aparecerá la salida cuando ejecutes el módulo...</span>}
                  {running && <span style={{ color:"#4ade80" }}>█</span>}
                </div>
                {excelFiles.length>0 && (
                  <button onClick={()=>setTab("download")} style={{ marginTop:14, width:"100%", background:"linear-gradient(135deg,#16a34a,#15803d)", color:"#fff", border:"none", borderRadius:10, padding:"12px", fontSize:14, fontWeight:800, cursor:"pointer" }}>
                    ⬇️ Ver archivos para descargar ({excelFiles.length})
                  </button>
                )}
              </div>
            )}

            {/* DOWNLOAD */}
            {tab==="download" && (
              <div>
                <div style={{ fontSize:12, color:"#94a3b8", marginBottom:20 }}>
                  {excelFiles.length>0 ? "✅ Archivos listos para descargar:" : "Ejecuta el módulo primero para generar los archivos."}
                </div>
                {excelFiles.length>0 ? (
                  <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
                    {excelFiles.map((f,i)=>(
                      <div key={i} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"16px 20px", background:"#f0fdf4", borderRadius:12, border:"2px solid #16a34a" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                          <span style={{ fontSize:28 }}>📊</span>
                          <div>
                            <div style={{ fontWeight:700, fontSize:14, color:"#15803d" }}>{f}</div>
                            <div style={{ fontSize:11, color:"#64748b" }}>Excel generado por Módulo {mod.id}</div>
                          </div>
                        </div>
                        <button onClick={()=>handleDownload(f)} style={{
                          display:"flex", alignItems:"center", gap:8,
                          background:"linear-gradient(135deg,#16a34a,#15803d)",
                          color:"#fff", border:"none", borderRadius:9, padding:"10px 22px",
                          fontSize:13, fontWeight:800, cursor:"pointer",
                          boxShadow:"0 4px 12px rgba(22,163,74,0.4)",
                        }}>
                          ⬇️ Descargar
                        </button>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div style={{ textAlign:"center", padding:40, color:"#94a3b8" }}>
                    <div style={{ fontSize:48, marginBottom:12 }}>📭</div>
                    <div style={{ fontSize:14 }}>Aún no hay archivos generados</div>
                    <button onClick={()=>setTab("upload")} style={{ marginTop:16, background:mod.color, color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontSize:13, fontWeight:700, cursor:"pointer" }}>
                      Ir a Cargar Archivos
                    </button>
                  </div>
                )}
              </div>
            )}

          </div>
        </div>
      </div>
    </div>
  );
}
