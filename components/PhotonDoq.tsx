"use client";
// @ts-nocheck
import React, { useState, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════
// ICONS
// ═══════════════════════════════════════
const Spinner = ({ size = 16, color = "currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" style={{ animation: "spin 1s linear infinite" }}><path d="M21 12a9 9 0 1 1-6.219-8.56" /></svg>
);

// ═══════════════════════════════════════
// VALIDATION ENGINE
// ═══════════════════════════════════════
function validateNIF(nif) {
  if (!nif) return { valid: false, msg: "Vazio" };
  const clean = String(nif).replace(/\s/g, "");
  if (!/^\d{9}$/.test(clean)) return { valid: false, msg: "NIF deve ter 9 dígitos" };
  const firstDigits = ["1", "2", "3", "5", "6", "7", "8", "9"];
  if (!firstDigits.includes(clean[0])) return { valid: false, msg: "Primeiro dígito inválido" };
  let sum = 0;
  for (let i = 0; i < 8; i++) sum += parseInt(clean[i]) * (9 - i);
  const mod = sum % 11;
  const check = mod < 2 ? 0 : 11 - mod;
  if (parseInt(clean[8]) !== check) return { valid: false, msg: "Dígito de controlo inválido" };
  return { valid: true, msg: "NIF válido" };
}

function parseDate(str) {
  if (!str) return null;
  const parts = String(str).match(/(\d{2})[\/\-.](\d{2})[\/\-.](\d{4})/);
  if (parts) return new Date(parts[3], parts[2] - 1, parts[1]);
  const iso = String(str).match(/(\d{4})[\/\-.](\d{2})[\/\-.](\d{2})/);
  if (iso) return new Date(iso[1], iso[2] - 1, iso[3]);
  return null;
}

function validateFields(campos, tipo) {
  const warnings = [];
  if (campos.emitente_nif) {
    const r = validateNIF(campos.emitente_nif);
    if (!r.valid) warnings.push({ field: "emitente_nif", msg: r.msg, severity: "error" });
  }
  if (campos.destinatario_nif) {
    const r = validateNIF(campos.destinatario_nif);
    if (!r.valid) warnings.push({ field: "destinatario_nif", msg: r.msg, severity: "error" });
  }
  const dataDoc = parseDate(campos.data_documento);
  const dataVenc = parseDate(campos.data_vencimento);
  if (dataDoc && dataVenc && dataVenc < dataDoc) {
    warnings.push({ field: "data_vencimento", msg: "Vencimento anterior à emissão", severity: "warning" });
  }
  const base = parseFloat(campos.valor_sem_iva);
  const iva = parseFloat(campos.iva);
  const total = parseFloat(campos.valor_total);
  if (!isNaN(base) && !isNaN(iva) && !isNaN(total)) {
    const expected = Math.round((base + iva) * 100) / 100;
    const actual = Math.round(total * 100) / 100;
    if (Math.abs(expected - actual) > 0.02) {
      warnings.push({ field: "valor_total", msg: `Base+IVA=${expected.toFixed(2)} ≠ Total=${actual.toFixed(2)}`, severity: "warning" });
    }
  }
  return warnings;
}

// ═══════════════════════════════════════
// DEFAULT DATA
// ═══════════════════════════════════════
const DOC_TYPES = {
  fatura: { label: "Fatura", icon: "💰", color: "#a855f7" },
  guia_transporte: { label: "Guia Transporte", icon: "🚛", color: "#2563eb" },
  guia_remessa: { label: "Guia Remessa", icon: "📦", color: "#7c3aed" },
  certificado: { label: "Certificado", icon: "📜", color: "#7c3aed" },
  contrato: { label: "Contrato", icon: "📝", color: "#b45309" },
  relatorio: { label: "Relatório", icon: "🔧", color: "#dc2626" },
  outro: { label: "Outro", icon: "📄", color: "#64748b" },
};

const FIELD_LABELS = {
  numero_documento: "Nº Documento", data_documento: "Data", data_vencimento: "Vencimento",
  emitente_nome: "Emitente", emitente_nif: "NIF Emitente", emitente_morada: "Morada Emitente",
  destinatario_nome: "Destinatário", destinatario_nif: "NIF Destinatário",
  valor_sem_iva: "Base Tributável", iva: "IVA", valor_total: "Valor Total",
  descricao_servico: "Descrição", referencia: "Referência", matricula: "Matrícula",
  local_carga: "Local Carga", local_descarga: "Local Descarga", peso: "Peso",
  certificacao: "Certificação", validade: "Validade", moeda: "Moeda",
};

const DEFAULT_PROFILES = [
  { id: "fatura", label: "Fatura", icon: "💰", color: "#a855f7", fields: [
    { key: "numero_documento", label: "Nº Documento", required: true },
    { key: "data_documento", label: "Data", required: true },
    { key: "data_vencimento", label: "Vencimento", required: false },
    { key: "emitente_nome", label: "Emitente", required: true },
    { key: "emitente_nif", label: "NIF Emitente", required: true },
    { key: "destinatario_nome", label: "Destinatário", required: true },
    { key: "destinatario_nif", label: "NIF Destinatário", required: true },
    { key: "valor_sem_iva", label: "Base Tributável", required: true },
    { key: "iva", label: "IVA", required: true },
    { key: "valor_total", label: "Valor Total", required: true },
    { key: "descricao_servico", label: "Descrição", required: false },
  ]},
  { id: "guia_transporte", label: "Guia Transporte", icon: "🚛", color: "#2563eb", fields: [
    { key: "numero_documento", label: "Nº Guia", required: true },
    { key: "data_documento", label: "Data", required: true },
    { key: "emitente_nome", label: "Expedidor", required: true },
    { key: "destinatario_nome", label: "Destinatário", required: true },
    { key: "local_carga", label: "Local Carga", required: true },
    { key: "local_descarga", label: "Local Descarga", required: true },
    { key: "matricula", label: "Matrícula", required: false },
  ]},
  { id: "certificado", label: "Certificado", icon: "📜", color: "#7c3aed", fields: [
    { key: "numero_documento", label: "Nº Certificado", required: true },
    { key: "data_documento", label: "Data Emissão", required: true },
    { key: "validade", label: "Validade", required: false },
    { key: "emitente_nome", label: "Entidade Emissora", required: true },
    { key: "certificacao", label: "Certificação", required: true },
  ]},
  { id: "contrato", label: "Contrato", icon: "📝", color: "#b45309", fields: [
    { key: "data_documento", label: "Data", required: true },
    { key: "emitente_nome", label: "Parte 1", required: true },
    { key: "destinatario_nome", label: "Parte 2", required: true },
    { key: "valor_total", label: "Valor", required: false },
    { key: "descricao_servico", label: "Objeto", required: true },
  ]},
  { id: "outro", label: "Outro", icon: "📄", color: "#64748b", fields: [
    { key: "data_documento", label: "Data", required: false },
    { key: "emitente_nome", label: "Emitente", required: false },
    { key: "descricao_servico", label: "Descrição", required: false },
  ]},
];

function makeCompany(name) {
  return { id: "c" + Date.now(), name: name || "Nova Empresa", nif: "", sector: "", notes: "", profiles: JSON.parse(JSON.stringify(DEFAULT_PROFILES)) };
}

// ═══════════════════════════════════════
// PROMPT BUILDER
// ═══════════════════════════════════════
function buildPrompt(company) {
  const types = company.profiles.map(p => p.id).join("|");
  const allKeys = new Set();
  company.profiles.forEach(p => p.fields.forEach(f => allKeys.add(f.key)));
  const fieldsJson = [...allKeys].map(k => `"${k}": "valor ou null"`).join(", ");

  return `Analisa este documento com OCR e extrai toda a informação.
EMPRESA: ${company.name || "N/A"} | NIF: ${company.nif || "N/A"} | Setor: ${company.sector || "N/A"}
${company.notes ? "NOTAS: " + company.notes : ""}
Retorna APENAS JSON puro (sem markdown, sem backticks):
{"tipo_documento":"${types}","confianca":0.95,"resumo":"Descrição breve","texto_completo":"Todo o texto","campos":{${fieldsJson}},"linhas_detalhe":[{"descricao":"","quantidade":"","unidade":"","preco_unitario":"","total_linha":""}]}
Regras: campos não encontrados=null | Datas DD/MM/AAAA | Valores numéricos sem € | NIF só dígitos | linhas_detalhe só se existirem | JSON puro`;
}

// ═══════════════════════════════════════
// STORAGE
// ═══════════════════════════════════════
function sGet(k: string) { try { const r = localStorage.getItem(k); return r ? JSON.parse(r) : null; } catch { return null; } }
function sSet(k: string, v: any) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.warn(e); } }

// ═══════════════════════════════════════
// EDITABLE FIELD COMPONENT
// ═══════════════════════════════════════
function EditableField({ label, value, fieldKey, onSave, validation }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value || "");
  const inputRef = useRef(null);

  useEffect(() => { if (editing && inputRef.current) inputRef.current.focus(); }, [editing]);
  useEffect(() => { setVal(value || ""); }, [value]);

  const save = () => { onSave(fieldKey, val); setEditing(false); };
  const cancel = () => { setVal(value || ""); setEditing(false); };

  return (
    <div style={{ padding: "7px 9px", borderRadius: 5, background: "#0c0818", position: "relative", border: validation ? `1px solid ${validation.severity === "error" ? "#dc2626" : "#d97706"}` : "1px solid transparent" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: "#5e5478" }}>{label}</span>
        {!editing && (
          <button onClick={() => setEditing(true)} style={{ background: "none", border: "none", cursor: "pointer", color: "#5e5478", fontSize: 10, padding: "0 2px" }} title="Editar">✎</button>
        )}
      </div>
      {editing ? (
        <div style={{ display: "flex", gap: 4, marginTop: 3 }}>
          <input ref={inputRef} value={val} onChange={e => setVal(e.target.value)} onKeyDown={e => { if (e.key === "Enter") save(); if (e.key === "Escape") cancel(); }}
            style={{ flex: 1, padding: "4px 6px", borderRadius: 4, border: "1px solid #2a2340", background: "#18122e", color: "#f4f0ff", fontSize: 12, outline: "none" }} />
          <button onClick={save} style={{ background: "#a855f7", color: "#fff", border: "none", borderRadius: 4, padding: "4px 8px", fontSize: 10, cursor: "pointer", fontWeight: 600 }}>OK</button>
          <button onClick={cancel} style={{ background: "none", border: "1px solid #2a2340", borderRadius: 4, padding: "4px 6px", fontSize: 10, cursor: "pointer", color: "#5e5478" }}>✕</button>
        </div>
      ) : (
        <p style={{ fontSize: 13, fontWeight: 500, marginTop: 2, wordBreak: "break-word", color: "#f4f0ff" }}>{value || "—"}</p>
      )}
      {validation && (
        <div style={{ fontSize: 9, marginTop: 3, color: validation.severity === "error" ? "#f87171" : "#fbbf24", display: "flex", alignItems: "center", gap: 3 }}>
          {validation.severity === "error" ? "⛔" : "⚠️"} {validation.msg}
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════
// DASHBOARD COMPONENT
// ═══════════════════════════════════════
function Dashboard({ results }) {
  if (!results.length) return (
    <div style={S.empty}><p style={{ fontSize: 36 }}>📊</p><p style={{ fontSize: 12, color: "#5e5478", marginTop: 8 }}>Processa documentos para ver análises.</p></div>
  );

  const byType = {};
  const bySupplier = {};
  const byMonth = {};
  let totalValue = 0;
  let totalDocs = results.length;
  let avgConfidence = 0;

  results.forEach(r => {
    const t = r.tipo_documento || "outro";
    byType[t] = (byType[t] || 0) + 1;
    const sup = r.campos?.emitente_nome || "Desconhecido";
    bySupplier[sup] = (bySupplier[sup] || 0) + 1;
    if (r.campos?.valor_total) { const v = parseFloat(r.campos.valor_total); if (!isNaN(v)) totalValue += v; }
    avgConfidence += (r.confianca || 0);
    const d = parseDate(r.campos?.data_documento);
    if (d) { const mk = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); byMonth[mk] = (byMonth[mk] || 0) + 1; }
  });
  avgConfidence = Math.round(avgConfidence / totalDocs * 100);

  const topSuppliers = Object.entries(bySupplier).sort((a, b) => b[1] - a[1]).slice(0, 6);
  const typeEntries = Object.entries(byType).sort((a, b) => b[1] - a[1]);
  const monthEntries = Object.entries(byMonth).sort((a, b) => a[0].localeCompare(b[0]));
  const typeColors = ["#a855f7", "#2563eb", "#7c3aed", "#b45309", "#dc2626", "#64748b", "#7c3aed"];

  return (
    <div style={{ animation: "fadeIn .3s" }}>
      {/* KPI Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, marginBottom: 16 }}>
        {[
          { label: "Total documentos", value: totalDocs, color: "#a855f7" },
          { label: "Valor faturado", value: "€" + totalValue.toLocaleString("pt-PT", { minimumFractionDigits: 2 }), color: "#2563eb" },
          { label: "Confiança média", value: avgConfidence + "%", color: avgConfidence > 85 ? "#059669" : "#d97706" },
          { label: "Fornecedores", value: Object.keys(bySupplier).length, color: "#7c3aed" },
        ].map((kpi, i) => (
          <div key={i} style={{ padding: "12px 14px", borderRadius: 8, background: "#110e20", border: "1px solid #2a2340" }}>
            <p style={{ fontSize: 10, color: "#5e5478", marginBottom: 4 }}>{kpi.label}</p>
            <p style={{ fontSize: 20, fontWeight: 700, color: kpi.color, lineHeight: 1 }}>{kpi.value}</p>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {/* By Type */}
        <div style={{ background: "#110e20", borderRadius: 8, border: "1px solid #2a2340", padding: 14 }}>
          <h4 style={{ fontSize: 12, fontWeight: 600, marginBottom: 10, color: "#8b82a8" }}>Por tipo de documento</h4>
          {typeEntries.map(([type, count], i) => {
            const dt = DOC_TYPES[type] || DOC_TYPES.outro;
            const pct = Math.round(count / totalDocs * 100);
            return (
              <div key={type} style={{ marginBottom: 8 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 3 }}>
                  <span style={{ color: "#f4f0ff" }}>{dt.icon} {dt.label}</span>
                  <span style={{ color: "#5e5478" }}>{count} ({pct}%)</span>
                </div>
                <div style={{ height: 6, borderRadius: 3, background: "#2a2340" }}>
                  <div style={{ height: 6, borderRadius: 3, background: typeColors[i % typeColors.length], width: pct + "%", transition: "width 0.5s" }} />
                </div>
              </div>
            );
          })}
        </div>

        {/* By Supplier */}
        <div style={{ background: "#110e20", borderRadius: 8, border: "1px solid #2a2340", padding: 14 }}>
          <h4 style={{ fontSize: 12, fontWeight: 600, marginBottom: 10, color: "#8b82a8" }}>Top fornecedores</h4>
          {topSuppliers.map(([name, count], i) => {
            const pct = Math.round(count / totalDocs * 100);
            return (
              <div key={i} style={{ marginBottom: 8 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 3 }}>
                  <span style={{ color: "#f4f0ff", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: "70%" }}>{name}</span>
                  <span style={{ color: "#5e5478" }}>{count}</span>
                </div>
                <div style={{ height: 6, borderRadius: 3, background: "#2a2340" }}>
                  <div style={{ height: 6, borderRadius: 3, background: "#a855f7", width: pct + "%", transition: "width 0.5s" }} />
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Monthly */}
      {monthEntries.length > 1 && (
        <div style={{ background: "#110e20", borderRadius: 8, border: "1px solid #2a2340", padding: 14, marginTop: 12 }}>
          <h4 style={{ fontSize: 12, fontWeight: 600, marginBottom: 10, color: "#8b82a8" }}>Volume mensal</h4>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 4, height: 100 }}>
            {monthEntries.map(([m, c]) => {
              const max = Math.max(...monthEntries.map(e => e[1]));
              const h = Math.max(8, Math.round(c / max * 90));
              return (
                <div key={m} style={{ flex: 1, textAlign: "center" }}>
                  <div style={{ height: h, background: "linear-gradient(180deg, #a855f7, #7c3aed)", borderRadius: "3px 3px 0 0", margin: "0 auto", maxWidth: 40 }} />
                  <p style={{ fontSize: 9, color: "#5e5478", marginTop: 4 }}>{m.slice(5)}/{m.slice(2, 4)}</p>
                  <p style={{ fontSize: 10, fontWeight: 600, color: "#f4f0ff" }}>{c}</p>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════
export default function App() {
  const [tab, setTab] = useState("process");
  const [companies, setCompanies] = useState([]);
  const [activeCo, setActiveCo] = useState(null);
  const [editing, setEditing] = useState(null);
  const [ready, setReady] = useState(false);

  const [files, setFiles] = useState([]);
  const [statuses, setStatuses] = useState([]);
  const [results, setResults] = useState([]);
  const [expanded, setExpanded] = useState({});
  const [busy, setBusy] = useState(false);
  const [allResults, setAllResults] = useState([]);
  const inputRef = useRef(null);

  // Search/filter
  const [searchQ, setSearchQ] = useState("");
  const [filterType, setFilterType] = useState("");
  const [filterCompany, setFilterCompany] = useState("");

  useEffect(() => {
    (() => {
      const cos = (sGet("pdq_cos")) || [];
      const all = (sGet("pdq_all")) || [];
      setCompanies(cos);
      setAllResults(all);
      if (cos.length) setActiveCo(cos[0]);
      setReady(true);
    })();
  }, []);

  const saveCos = (l) => { setCompanies(l); sSet("pdq_cos", l); };
  const saveAll = (l) => { const t = l.slice(0, 500); setAllResults(t); sSet("pdq_all", t); };

  const addFiles = (nf) => {
    const valid = [...nf].filter(f => f.type.startsWith("image/") || f.type === "application/pdf");
    if (valid.length) { setFiles(p => [...p, ...valid]); setStatuses(p => [...p, ...valid.map(() => "pending")]); }
  };

  const toB64 = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(f); });

  const process = async () => {
    if (!activeCo || busy) return;
    setBusy(true);
    const prompt = buildPrompt(activeCo);
    const nr = [];
    for (let i = 0; i < files.length; i++) {
      if (statuses[i] !== "pending") continue;
      setStatuses(p => { const n = [...p]; n[i] = "working"; return n; });
      try {
        const b64 = await toB64(files[i]);
        const isPdf = files[i].type === "application/pdf";
        const content = [
          isPdf ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }
                : { type: "image", source: { type: "base64", media_type: files[i].type || "image/png", data: b64 } },
          { type: "text", text: prompt },
        ];
        const resp = await fetch("/api/ocr", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 4096, messages: [{ role: "user", content }] }),
        });
        if (!resp.ok) throw new Error("API: " + resp.status);
        const data = await resp.json();
        const raw = (data.content || []).map(c => c.text || "").join("");
        let parsed;
        try { parsed = JSON.parse(raw.replace(/```json|```/g, "").trim()); }
        catch { parsed = { tipo_documento: "outro", confianca: 0, resumo: "Erro ao interpretar", texto_completo: raw, campos: {}, linhas_detalhe: [] }; }
        parsed.fileName = files[i].name;
        parsed.company = activeCo.name;
        parsed.companyId = activeCo.id;
        parsed.at = new Date().toISOString();
        parsed.edited = false;
        nr.push(parsed);
        setStatuses(p => { const n = [...p]; n[i] = "done"; return n; });
      } catch (err) {
        nr.push({ tipo_documento: "outro", confianca: 0, resumo: "Erro: " + err.message, campos: {}, fileName: files[i].name, company: activeCo.name, companyId: activeCo.id, at: new Date().toISOString() });
        setStatuses(p => { const n = [...p]; n[i] = "error"; return n; });
      }
    }
    setResults(p => [...p, ...nr]);
    const ex = {}; nr.forEach((_, i) => { ex[results.length + i] = true; }); setExpanded(p => ({ ...p, ...ex }));
    saveAll([...nr, ...allResults]);
    setBusy(false);
  };

  const updateField = (resultIdx, key, val) => {
    setResults(prev => {
      const next = [...prev];
      next[resultIdx] = { ...next[resultIdx], campos: { ...next[resultIdx].campos, [key]: val }, edited: true };
      return next;
    });
  };

  const updateLineItem = (resultIdx, lineIdx, key, val) => {
    setResults(prev => {
      const next = [...prev];
      const lines = [...(next[resultIdx].linhas_detalhe || [])];
      lines[lineIdx] = { ...lines[lineIdx], [key]: val };
      next[resultIdx] = { ...next[resultIdx], linhas_detalhe: lines, edited: true };
      return next;
    });
  };

  // Filtered results for search tab
  const filteredAll = useMemo(() => {
    let arr = allResults;
    if (searchQ) {
      const q = searchQ.toLowerCase();
      arr = arr.filter(r =>
        (r.fileName || "").toLowerCase().includes(q) ||
        (r.resumo || "").toLowerCase().includes(q) ||
        (r.campos?.emitente_nome || "").toLowerCase().includes(q) ||
        (r.campos?.destinatario_nome || "").toLowerCase().includes(q) ||
        (r.campos?.numero_documento || "").toLowerCase().includes(q) ||
        (r.campos?.descricao_servico || "").toLowerCase().includes(q)
      );
    }
    if (filterType) arr = arr.filter(r => r.tipo_documento === filterType);
    if (filterCompany) arr = arr.filter(r => r.companyId === filterCompany);
    return arr;
  }, [allResults, searchQ, filterType, filterCompany]);

  // ═══ EXPORT MODAL STATE ═══
  const [exportModal, setExportModal] = useState(null); // null or { data: [...] }

  // ═══ ERP EXPORT ENGINES ═══
  const ERP_FORMATS = {
    generic: { label: "Excel Genérico", icon: "📊", desc: "Formato universal com todos os campos" },
    toconline: { label: "TOConline", icon: "🟢", desc: "CSV pronto para importação em Compras" },
    odoo: { label: "Odoo", icon: "🟣", desc: "CSV compatível com import account.move" },
    phc: { label: "PHC GO", icon: "🔵", desc: "CSV formatado para importação PHC" },
    primavera: { label: "Primavera", icon: "🟠", desc: "CSV para módulo de Compras V10" },
    saft: { label: "SAF-T (PT)", icon: "🇵🇹", desc: "XML universal português para qualquer ERP" },
  };

  const fmtDate = (d, fmt) => {
    if (!d) return "";
    const s = String(d);
    const parts = s.match(/(\d{2})[\/\-.](\d{2})[\/\-.](\d{4})/);
    const iso = s.match(/(\d{4})[\/\-.](\d{2})[\/\-.](\d{2})/);
    let day, mon, year;
    if (parts) { day = parts[1]; mon = parts[2]; year = parts[3]; }
    else if (iso) { year = iso[1]; mon = iso[2]; day = iso[3]; }
    else return s;
    if (fmt === "iso") return `${year}-${mon}-${day}`;
    if (fmt === "pt") return `${day}/${mon}/${year}`;
    if (fmt === "slash") return `${day}/${mon}/${year}`;
    return `${year}-${mon}-${day}`;
  };

  const downloadFile = (content, filename, type = "text/csv;charset=utf-8;") => {
    try {
      const bom = type.includes("csv") ? "\uFEFF" : "";
      const blob = new Blob([bom + content], { type });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 500);
    } catch (e) {
      // Fallback: data URI
      const encoded = encodeURIComponent(content);
      const a = document.createElement("a");
      a.href = "data:" + type + "," + encoded;
      a.download = filename;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
    }
  };

  const exportForERP = (data, format) => {
    const invoices = data.filter(r => r.campos && (r.tipo_documento === "fatura" || r.campos.valor_total));
    const all = data;
    const date = new Date().toISOString().slice(0, 10);
    const coName = activeCo?.name?.replace(/\s+/g, "_") || "export";

    if (format === "generic") {
      const wb = XLSX.utils.book_new();
      const rows = all.map(r => {
        const o = { Ficheiro: r.fileName, Empresa: r.company, Tipo: r.tipo_documento, "Confiança": Math.round((r.confianca || 0) * 100) + "%" };
        if (r.campos) Object.entries(r.campos).forEach(([k, v]) => { if (v && v !== "null") o[FIELD_LABELS[k] || k] = v; });
        o["Editado"] = r.edited ? "Sim" : "Não";
        o["Processado"] = r.at ? new Date(r.at).toLocaleString("pt-PT") : "";
        return o;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Resultados");
      const det = [];
      all.forEach(r => (r.linhas_detalhe || []).forEach(l => det.push({ Ficheiro: r.fileName, Empresa: r.company, "Descrição": l.descricao, Qtd: l.quantidade, "Un.": l.unidade, "P.Unit.": l.preco_unitario, Total: l.total_linha })));
      if (det.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(det), "Linhas");
      const txt = all.filter(r => r.texto_completo).map(r => ({ Ficheiro: r.fileName, Texto: r.texto_completo }));
      if (txt.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(txt), "Texto OCR");
      XLSX.writeFile(wb, `PhotonDoq_${coName}_${date}.xlsx`);
    }

    else if (format === "toconline") {
      // TOConline purchase document import format
      const header = ["Tipo Doc.", "Nº Documento", "Data Documento", "Data Vencimento", "NIF Fornecedor", "Nome Fornecedor", "Descrição", "Base Tributável", "Taxa IVA (%)", "Valor IVA", "Total", "Moeda"];
      const rows = invoices.map(r => {
        const c = r.campos || {};
        const base = parseFloat(c.valor_sem_iva) || 0;
        const iva = parseFloat(c.iva) || 0;
        const taxRate = base > 0 ? Math.round(iva / base * 100) : 23;
        return [
          "FT", c.numero_documento || "", fmtDate(c.data_documento, "iso"), fmtDate(c.data_vencimento, "iso"),
          c.emitente_nif || "", c.emitente_nome || "", c.descricao_servico || r.resumo || "",
          base.toFixed(2), taxRate, iva.toFixed(2), (parseFloat(c.valor_total) || 0).toFixed(2), "EUR"
        ];
      });
      // Line items sheet
      const lineHeader = ["Nº Documento Origem", "Descrição Linha", "Quantidade", "Unidade", "Preço Unitário", "Taxa IVA (%)"];
      const lineRows = [];
      invoices.forEach(r => {
        (r.linhas_detalhe || []).forEach(l => {
          lineRows.push([r.campos?.numero_documento || "", l.descricao || "", l.quantidade || "1", l.unidade || "un", l.preco_unitario || "", "23"]);
        });
      });

      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.aoa_to_sheet([header, ...rows]);
      ws1["!cols"] = header.map(() => ({ wch: 18 }));
      XLSX.utils.book_append_sheet(wb, ws1, "Documentos Compra");
      if (lineRows.length) {
        const ws2 = XLSX.utils.aoa_to_sheet([lineHeader, ...lineRows]);
        XLSX.utils.book_append_sheet(wb, ws2, "Linhas Detalhe");
      }
      XLSX.writeFile(wb, `TOConline_Import_${coName}_${date}.xlsx`);
    }

    else if (format === "odoo") {
      // Odoo account.move import format - using XLSX for reliable download
      const header = ["move_type", "partner_id/name", "partner_id/vat", "ref", "invoice_date", "invoice_date_due", "invoice_line_ids/name", "invoice_line_ids/quantity", "invoice_line_ids/price_unit", "invoice_line_ids/tax_ids", "currency_id"];
      const rows = [];
      invoices.forEach(r => {
        const c = r.campos || {};
        const lines = r.linhas_detalhe?.length ? r.linhas_detalhe : [{ descricao: c.descricao_servico || "Serviço", quantidade: "1", preco_unitario: c.valor_sem_iva || c.valor_total || "0" }];
        lines.forEach((l, i) => {
          rows.push([
            "in_invoice",
            i === 0 ? (c.emitente_nome || "") : "",
            i === 0 ? ("PT" + (c.emitente_nif || "")) : "",
            i === 0 ? (c.numero_documento || "") : "",
            i === 0 ? fmtDate(c.data_documento, "iso") : "",
            i === 0 ? fmtDate(c.data_vencimento, "iso") : "",
            l.descricao || "Serviço",
            l.quantidade || "1",
            l.preco_unitario || c.valor_sem_iva || "0",
            "IVA 23%",
            "EUR"
          ]);
        });
      });
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      ws["!cols"] = header.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(wb, ws, "account_move_import");
      XLSX.writeFile(wb, `Odoo_Import_${coName}_${date}.xlsx`);
    }

    else if (format === "phc") {
      // PHC GO import format - purchase documents
      const header = ["TipoDocumento", "NumDocumento", "DataDocumento", "DataVencimento", "CodFornecedor", "NomeFornecedor", "NIFFornecedor", "Descricao", "ValorBase", "TaxaIVA", "ValorIVA", "ValorTotal", "Moeda", "Observacoes"];
      const rows = invoices.map(r => {
        const c = r.campos || {};
        const base = parseFloat(c.valor_sem_iva) || 0;
        const iva = parseFloat(c.iva) || 0;
        return [
          "VFA", c.numero_documento || "", fmtDate(c.data_documento, "iso"), fmtDate(c.data_vencimento, "iso"),
          "", c.emitente_nome || "", c.emitente_nif || "", c.descricao_servico || "",
          base.toFixed(2), base > 0 ? Math.round(iva / base * 100) : 23, iva.toFixed(2),
          (parseFloat(c.valor_total) || 0).toFixed(2), "EUR", `Processado PhotonDoq AI ${date}`
        ];
      });
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      ws["!cols"] = header.map(() => ({ wch: 16 }));
      XLSX.utils.book_append_sheet(wb, ws, "Compras PHC");
      // Line items
      const lh = ["NumDocOrigem", "CodArtigo", "Descricao", "Quantidade", "Unidade", "PrecoUnitario", "TaxaIVA"];
      const lr = [];
      invoices.forEach(r => (r.linhas_detalhe || []).forEach(l => {
        lr.push([r.campos?.numero_documento || "", "", l.descricao || "", l.quantidade || "1", l.unidade || "UN", l.preco_unitario || "", "23"]);
      }));
      if (lr.length) { const ws2 = XLSX.utils.aoa_to_sheet([lh, ...lr]); XLSX.utils.book_append_sheet(wb, ws2, "Linhas"); }
      XLSX.writeFile(wb, `PHC_Import_${coName}_${date}.xlsx`);
    }

    else if (format === "primavera") {
      // Primavera V10 purchase import format
      const header = ["TipoDoc", "Serie", "NumDoc", "DataDoc", "DataVenc", "Entidade", "NomeEntidade", "NumContrib", "Descricao", "TotalMerc", "TotalIva", "TotalDoc", "Moeda", "CodMoeda"];
      const rows = invoices.map(r => {
        const c = r.campos || {};
        return [
          "VFA", "A", c.numero_documento || "", fmtDate(c.data_documento, "iso"), fmtDate(c.data_vencimento, "iso"),
          "", c.emitente_nome || "", c.emitente_nif || "", c.descricao_servico || "",
          (parseFloat(c.valor_sem_iva) || 0).toFixed(2), (parseFloat(c.iva) || 0).toFixed(2),
          (parseFloat(c.valor_total) || 0).toFixed(2), "Euro", "EUR"
        ];
      });
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      ws["!cols"] = header.map(() => ({ wch: 16 }));
      XLSX.utils.book_append_sheet(wb, ws, "Compras Primavera");
      const lh = ["NumDocOrigem", "Artigo", "Descricao", "Quantidade", "Unidade", "PrecUnit", "TaxaIva", "Armazem"];
      const lr = [];
      invoices.forEach(r => (r.linhas_detalhe || []).forEach(l => {
        lr.push([r.campos?.numero_documento || "", "", l.descricao || "", l.quantidade || "1", l.unidade || "UN", l.preco_unitario || "", "23", "A1"]);
      }));
      if (lr.length) { const ws2 = XLSX.utils.aoa_to_sheet([lh, ...lr]); XLSX.utils.book_append_sheet(wb, ws2, "Linhas"); }
      XLSX.writeFile(wb, `Primavera_Import_${coName}_${date}.xlsx`);
    }

    else if (format === "saft") {
      // SAF-T PT simplified XML for purchase documents
      const now = new Date().toISOString();
      const coNif = activeCo?.nif || "000000000";
      const coFullName = activeCo?.name || "Empresa";
      let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
      xml += `<!-- SAF-T PT - Ficheiro gerado por PhotonDoq AI (Macro Consulting) -->\n`;
      xml += `<!-- Importação de documentos de compra extraídos por OCR -->\n`;
      xml += `<AuditFile xmlns="urn:OECD:StandardAuditFile-Tax:PT_1.04_01">\n`;
      xml += `  <Header>\n`;
      xml += `    <AuditFileVersion>1.04_01</AuditFileVersion>\n`;
      xml += `    <CompanyID>${coNif}</CompanyID>\n`;
      xml += `    <TaxRegistrationNumber>${coNif}</TaxRegistrationNumber>\n`;
      xml += `    <CompanyName>${coFullName}</CompanyName>\n`;
      xml += `    <FiscalYear>${new Date().getFullYear()}</FiscalYear>\n`;
      xml += `    <DateCreated>${date}</DateCreated>\n`;
      xml += `    <SoftwareCertificateNumber>0</SoftwareCertificateNumber>\n`;
      xml += `    <ProductCompanyTaxID>000000000</ProductCompanyTaxID>\n`;
      xml += `    <ProductID>PhotonDoq AI/Macro Consulting</ProductID>\n`;
      xml += `    <ProductVersion>3.0</ProductVersion>\n`;
      xml += `  </Header>\n`;
      xml += `  <MasterFiles>\n`;
      // Suppliers
      const suppliers = {};
      invoices.forEach(r => { const nif = r.campos?.emitente_nif; if (nif && nif !== "null") suppliers[nif] = r.campos?.emitente_nome || ""; });
      Object.entries(suppliers).forEach(([nif, name]) => {
        xml += `    <Supplier>\n`;
        xml += `      <SupplierID>${nif}</SupplierID>\n`;
        xml += `      <SupplierTaxID>${nif}</SupplierTaxID>\n`;
        xml += `      <CompanyName>${name}</CompanyName>\n`;
        xml += `    </Supplier>\n`;
      });
      xml += `  </MasterFiles>\n`;
      xml += `  <SourceDocuments>\n`;
      xml += `    <Invoices>\n`;
      xml += `      <NumberOfEntries>${invoices.length}</NumberOfEntries>\n`;
      const totalDebit = invoices.reduce((a, r) => a + (parseFloat(r.campos?.valor_total) || 0), 0);
      xml += `      <TotalDebit>${totalDebit.toFixed(2)}</TotalDebit>\n`;
      xml += `      <TotalCredit>0.00</TotalCredit>\n`;
      invoices.forEach((r, i) => {
        const c = r.campos || {};
        xml += `      <Invoice>\n`;
        xml += `        <InvoiceNo>${c.numero_documento || "DOC/" + (i + 1)}</InvoiceNo>\n`;
        xml += `        <InvoiceDate>${fmtDate(c.data_documento, "iso")}</InvoiceDate>\n`;
        xml += `        <InvoiceType>FT</InvoiceType>\n`;
        xml += `        <SupplierID>${c.emitente_nif || ""}</SupplierID>\n`;
        xml += `        <DocumentTotals>\n`;
        xml += `          <TaxPayable>${(parseFloat(c.iva) || 0).toFixed(2)}</TaxPayable>\n`;
        xml += `          <NetTotal>${(parseFloat(c.valor_sem_iva) || 0).toFixed(2)}</NetTotal>\n`;
        xml += `          <GrossTotal>${(parseFloat(c.valor_total) || 0).toFixed(2)}</GrossTotal>\n`;
        xml += `        </DocumentTotals>\n`;
        (r.linhas_detalhe || []).forEach((l, li) => {
          xml += `        <Line>\n`;
          xml += `          <LineNumber>${li + 1}</LineNumber>\n`;
          xml += `          <Description>${l.descricao || "Serviço"}</Description>\n`;
          xml += `          <Quantity>${l.quantidade || "1"}</Quantity>\n`;
          xml += `          <UnitPrice>${l.preco_unitario || "0"}</UnitPrice>\n`;
          xml += `          <CreditAmount>${l.total_linha || "0"}</CreditAmount>\n`;
          xml += `        </Line>\n`;
        });
        xml += `      </Invoice>\n`;
      });
      xml += `    </Invoices>\n`;
      xml += `  </SourceDocuments>\n`;
      xml += `</AuditFile>`;
      downloadFile(xml, `SAFT_PT_${coName}_${date}.xml`, "application/xml;charset=utf-8;");
    }

    setExportModal(null);
  };

  const reset = () => { setFiles([]); setStatuses([]); setResults([]); setExpanded({}); };

  if (!ready) return <div style={S.center}><Spinner size={20} /><span style={{ marginLeft: 8, color: "#8b82a8" }}>A carregar...</span></div>;

  return (
    <div style={S.root}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        input,textarea,select{font-family:inherit}
        input::placeholder{color:#5e5478}
        ::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:#2a2340;border-radius:2px}
      `}</style>

      {/* HEADER */}
      <div style={S.header}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ width: 30, height: 30, borderRadius: 8, background: "linear-gradient(135deg,#a855f7,#7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 14, fontWeight: 700 }}>⚡</div>
          <span style={{ fontSize: 15, fontWeight: 700 }}>Photon<span style={{ color: "#a855f7" }}>Doq</span></span>
          <span style={{ fontSize: 9, color: "#2a2340", fontWeight: 500, letterSpacing: "0.08em" }}>v3.0</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          {activeCo && <span style={S.badge}>{activeCo.name}</span>}
          {companies.length > 1 && <select value={activeCo?.id || ""} onChange={e => setActiveCo(companies.find(c => c.id === e.target.value))} style={S.select}>{companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>}
        </div>
      </div>

      {/* TABS */}
      <div style={S.tabs}>
        {[["process", "⚡ Processar"], ["search", "🔍 Pesquisa"], ["dashboard", "📊 Dashboard"], ["companies", "🏢 Empresas"]].map(([id, label]) => (
          <button key={id} onClick={() => { setTab(id); setEditing(null); }} style={{ ...S.tab, color: tab === id ? "#a855f7" : "#5e5478", borderBottom: tab === id ? "2px solid #a855f7" : "2px solid transparent", fontWeight: tab === id ? 600 : 400 }}>{label}</button>
        ))}
      </div>

      {/* CONTENT */}
      <div style={S.content}>
        <div style={{ maxWidth: 920, margin: "0 auto" }}>

          {/* ═══ PROCESS ═══ */}
          {tab === "process" && (
            <div style={{ animation: "fadeIn .3s" }}>
              {!companies.length ? (
                <div style={S.empty}>
                  <p style={{ fontSize: 40 }}>🏢</p>
                  <h3 style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Configura a primeira empresa</h3>
                  <p style={{ fontSize: 12, color: "#5e5478", marginBottom: 16 }}>O OCR precisa de contexto para extrair os campos certos.</p>
                  <button onClick={() => setTab("companies")} style={S.btnPrimary}>+ Adicionar Empresa</button>
                </div>
              ) : (
                <>
                  <div style={S.dropzone} onClick={() => inputRef.current?.click()}
                    onDragOver={e => { e.preventDefault(); e.currentTarget.style.borderColor = "#a855f7"; }}
                    onDragLeave={e => { e.currentTarget.style.borderColor = "#2a2340"; }}
                    onDrop={e => { e.preventDefault(); e.currentTarget.style.borderColor = "#2a2340"; addFiles(e.dataTransfer.files); }}>
                    <input ref={inputRef} type="file" accept="image/*,.pdf" multiple onChange={e => { addFiles(e.target.files); e.target.value = ""; }} style={{ display: "none" }} />
                    <p style={{ fontSize: 28, marginBottom: 6 }}>📎</p>
                    <p style={{ fontSize: 14, fontWeight: 600 }}>Arrastar ficheiros ou clicar</p>
                    <p style={{ fontSize: 11, color: "#5e5478", marginTop: 4 }}>Perfil: <strong style={{ color: "#a855f7" }}>{activeCo?.name}</strong></p>
                  </div>

                  {files.length > 0 && (
                    <div style={{ marginTop: 14 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                        <span style={{ fontSize: 11, color: "#5e5478", fontWeight: 600 }}>{files.length} ficheiro(s)</span>
                        {!busy && <button onClick={reset} style={S.btnGhost}>Limpar</button>}
                      </div>
                      {files.map((f, i) => (
                        <div key={i} style={S.fileRow}>
                          <span style={{ fontSize: 13 }}>📄</span>
                          <span style={{ flex: 1, fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                          <span style={{ fontSize: 10, color: "#5e5478" }}>{(f.size / 1024).toFixed(0)}KB</span>
                          {statuses[i] === "working" && <Spinner size={13} color="#a855f7" />}
                          {statuses[i] === "done" && <span style={{ color: "#059669" }}>✓</span>}
                          {statuses[i] === "error" && <span style={{ color: "#dc2626" }}>✗</span>}
                          {statuses[i] === "pending" && !busy && <button onClick={() => { setFiles(p => p.filter((_, j) => j !== i)); setStatuses(p => p.filter((_, j) => j !== i)); }} style={S.btnIcon}>×</button>}
                        </div>
                      ))}
                      {!busy && statuses.some(s => s === "pending") && <button onClick={process} style={S.btnProcess}>⚡ Processar com perfil {activeCo?.name}</button>}
                      {busy && <div style={S.processing}><Spinner size={14} color="#a855f7" /> A IA está a extrair dados...</div>}
                    </div>
                  )}

                  {/* RESULTS WITH EDITING + VALIDATION */}
                  {results.length > 0 && (
                    <div style={{ marginTop: 24 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                        <h3 style={{ fontSize: 15, fontWeight: 700 }}>Resultados ({results.length})</h3>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button onClick={() => setExportModal({ data: results })} style={S.btnExport}>📥 Exportar para ERP</button>
                          <button onClick={reset} style={S.btnSecondary}>Nova sessão</button>
                        </div>
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, marginBottom: 12 }}>
                        {[
                          ["Documentos", results.length, "#a855f7"],
                          ["Campos extraídos", results.reduce((a, r) => a + Object.values(r.campos || {}).filter(v => v && v !== "null").length, 0), "#2563eb"],
                          ["Alertas", results.reduce((a, r) => a + validateFields(r.campos || {}, r.tipo_documento).length, 0), results.reduce((a, r) => a + validateFields(r.campos || {}, r.tipo_documento).length, 0) > 0 ? "#d97706" : "#059669"],
                        ].map(([l, v, c], i) => (
                          <div key={i} style={S.stat}><span style={{ fontSize: 20, fontWeight: 700, color: c }}>{v}</span><span style={{ fontSize: 10, color: "#5e5478" }}>{l}</span></div>
                        ))}
                      </div>

                      {results.map((r, ri) => {
                        const dt = DOC_TYPES[r.tipo_documento] || DOC_TYPES.outro;
                        const fields = Object.entries(r.campos || {}).filter(([_, v]) => v && v !== "null");
                        const warnings = validateFields(r.campos || {}, r.tipo_documento);
                        const warnMap = {};
                        warnings.forEach(w => { warnMap[w.field] = w; });
                        const isExp = expanded[ri];

                        return (
                          <div key={ri} style={{ ...S.card, animation: "fadeIn .3s" }}>
                            <div onClick={() => setExpanded(p => ({ ...p, [ri]: !p[ri] }))} style={S.cardHead}>
                              <span style={{ fontSize: 20 }}>{dt.icon}</span>
                              <div style={{ flex: 1 }}>
                                <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 2, flexWrap: "wrap" }}>
                                  <span style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: dt.color, background: dt.color + "18", padding: "2px 6px", borderRadius: 3 }}>{dt.label}</span>
                                  <span style={{ fontSize: 10, color: "#5e5478" }}>{Math.round((r.confianca || 0) * 100)}%</span>
                                  {r.edited && <span style={{ fontSize: 9, color: "#a855f7", background: "#a855f718", padding: "1px 5px", borderRadius: 3, fontWeight: 600 }}>EDITADO</span>}
                                  {warnings.length > 0 && <span style={{ fontSize: 9, color: "#d97706", background: "#d9770618", padding: "1px 5px", borderRadius: 3, fontWeight: 600 }}>{warnings.length} alerta{warnings.length > 1 ? "s" : ""}</span>}
                                </div>
                                <span style={{ fontSize: 12, fontWeight: 500 }}>{r.fileName}</span>
                              </div>
                              <span style={{ fontSize: 10, color: "#5e5478" }}>{fields.length}</span>
                              <span style={{ color: "#5e5478", fontSize: 12 }}>{isExp ? "▲" : "▼"}</span>
                            </div>

                            {isExp && (
                              <div style={{ padding: "12px 14px", borderTop: "1px solid #2a2340" }}>
                                {r.resumo && <p style={{ fontSize: 12, color: "#8b82a8", marginBottom: 10, fontStyle: "italic" }}>{r.resumo}</p>}
                                {fields.length > 0 && (
                                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))", gap: 6 }}>
                                    {fields.map(([k, v]) => (
                                      <EditableField
                                        key={k}
                                        label={FIELD_LABELS[k] || k}
                                        value={v}
                                        fieldKey={k}
                                        onSave={(key, val) => updateField(ri, key, val)}
                                        validation={warnMap[k]}
                                      />
                                    ))}
                                  </div>
                                )}

                                {r.linhas_detalhe?.length > 0 && (
                                  <div style={{ marginTop: 10, borderRadius: 6, overflow: "auto", border: "1px solid #2a2340" }}>
                                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                                      <thead><tr style={{ background: "#0c0818" }}>
                                        {["Descrição", "Qtd", "Un.", "P.Unit.", "Total"].map(h => <th key={h} style={S.th}>{h}</th>)}
                                      </tr></thead>
                                      <tbody>
                                        {r.linhas_detalhe.map((l, li) => (
                                          <tr key={li} style={{ borderTop: "1px solid #2a2340" }}>
                                            {["descricao", "quantidade", "unidade", "preco_unitario", "total_linha"].map(k => (
                                              <td key={k} style={S.td}>
                                                <input
                                                  value={l[k] || ""}
                                                  onChange={e => updateLineItem(ri, li, k, e.target.value)}
                                                  style={{ background: "none", border: "none", color: "#f4f0ff", fontSize: 11, width: "100%", outline: "none", padding: "2px 0" }}
                                                />
                                              </td>
                                            ))}
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                )}

                                {r.texto_completo && (
                                  <details style={{ marginTop: 10 }}>
                                    <summary style={{ fontSize: 11, color: "#5e5478", cursor: "pointer" }}>Texto OCR completo</summary>
                                    <pre style={S.ocrText}>{r.texto_completo}</pre>
                                  </details>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </>
              )}
            </div>
          )}

          {/* ═══ SEARCH ═══ */}
          {tab === "search" && (
            <div style={{ animation: "fadeIn .3s" }}>
              <h2 style={{ fontSize: 17, fontWeight: 700, marginBottom: 12 }}>Pesquisa e Filtros</h2>
              <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
                <input value={searchQ} onChange={e => setSearchQ(e.target.value)} placeholder="Pesquisar por nome, emitente, NIF, descrição..."
                  style={{ ...S.input, flex: "1 1 200px" }} />
                <select value={filterType} onChange={e => setFilterType(e.target.value)} style={{ ...S.select, padding: "7px 8px" }}>
                  <option value="">Todos os tipos</option>
                  {Object.entries(DOC_TYPES).map(([k, v]) => <option key={k} value={k}>{v.icon} {v.label}</option>)}
                </select>
                <select value={filterCompany} onChange={e => setFilterCompany(e.target.value)} style={{ ...S.select, padding: "7px 8px" }}>
                  <option value="">Todas as empresas</option>
                  {companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 12, color: "#5e5478" }}>{filteredAll.length} resultado{filteredAll.length !== 1 ? "s" : ""}</span>
                {filteredAll.length > 0 && <button onClick={() => setExportModal({ data: filteredAll })} style={S.btnExport}>📥 Exportar filtrados</button>}
              </div>

              {!filteredAll.length ? (
                <div style={S.empty}><p style={{ fontSize: 36 }}>🔍</p><p style={{ fontSize: 12, color: "#5e5478", marginTop: 8 }}>{allResults.length ? "Sem resultados para este filtro." : "Processa documentos primeiro."}</p></div>
              ) : (
                <div style={{ display: "grid", gap: 3 }}>
                  {filteredAll.slice(0, 50).map((r, i) => {
                    const dt = DOC_TYPES[r.tipo_documento] || DOC_TYPES.outro;
                    return (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 6, background: "#110e20", border: "1px solid #2a2340" }}>
                        <span style={{ fontSize: 16 }}>{dt.icon}</span>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                            <span style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.fileName}</span>
                            <span style={{ fontSize: 9, color: dt.color, background: dt.color + "18", padding: "1px 5px", borderRadius: 3, fontWeight: 600, flexShrink: 0 }}>{dt.label}</span>
                          </div>
                          <div style={{ display: "flex", gap: 8, marginTop: 2, fontSize: 10, color: "#5e5478" }}>
                            <span>{r.company}</span>
                            {r.campos?.emitente_nome && <span>• {r.campos.emitente_nome}</span>}
                            {r.campos?.valor_total && <span>• €{r.campos.valor_total}</span>}
                          </div>
                        </div>
                        <div style={{ textAlign: "right", flexShrink: 0 }}>
                          <span style={{ fontSize: 10, color: "#5e5478" }}>{r.at ? new Date(r.at).toLocaleDateString("pt-PT") : ""}</span>
                          <div style={{ fontSize: 10, color: (r.confianca || 0) > 0.8 ? "#059669" : "#d97706" }}>{Math.round((r.confianca || 0) * 100)}%</div>
                        </div>
                      </div>
                    );
                  })}
                  {filteredAll.length > 50 && <p style={{ fontSize: 11, color: "#5e5478", textAlign: "center", padding: 8 }}>A mostrar 50 de {filteredAll.length} resultados.</p>}
                </div>
              )}
            </div>
          )}

          {/* ═══ DASHBOARD ═══ */}
          {tab === "dashboard" && <Dashboard results={allResults} />}

          {/* ═══ COMPANIES ═══ */}
          {tab === "companies" && !editing && (
            <div style={{ animation: "fadeIn .3s" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <div>
                  <h2 style={{ fontSize: 17, fontWeight: 700 }}>Empresas</h2>
                  <p style={{ fontSize: 11, color: "#5e5478", marginTop: 2 }}>Perfis e campos de extração por empresa</p>
                </div>
                <button onClick={() => setEditing(makeCompany())} style={S.btnPrimary}>+ Nova Empresa</button>
              </div>
              {!companies.length && <div style={S.empty}><p style={{ fontSize: 36 }}>🏢</p><p style={{ fontSize: 12, color: "#5e5478", marginTop: 8 }}>Clica "Nova Empresa" para começar.</p></div>}
              {companies.map(co => (
                <div key={co.id} style={{ ...S.companyRow, border: activeCo?.id === co.id ? "1px solid #a855f740" : "1px solid #2a2340" }} onClick={() => setActiveCo(co)}>
                  <div style={{ width: 38, height: 38, borderRadius: 8, background: "#a855f712", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🏢</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <span style={{ fontSize: 13, fontWeight: 600 }}>{co.name}</span>
                      {activeCo?.id === co.id && <span style={{ fontSize: 9, color: "#a855f7", background: "#a855f718", padding: "1px 5px", borderRadius: 3, fontWeight: 700 }}>ATIVA</span>}
                    </div>
                    <span style={{ fontSize: 10, color: "#5e5478" }}>{co.nif ? "NIF " + co.nif + " · " : ""}{co.profiles?.length || 0} perfis</span>
                  </div>
                  <button onClick={e => { e.stopPropagation(); setEditing(JSON.parse(JSON.stringify(co))); }} style={S.btnSmall}>Editar</button>
                  <button onClick={e => { e.stopPropagation(); const next = companies.filter(c => c.id !== co.id); saveCos(next); if (activeCo?.id === co.id) setActiveCo(next[0] || null); }} style={{ ...S.btnSmall, color: "#dc2626" }}>×</button>
                </div>
              ))}
            </div>
          )}

          {tab === "companies" && editing && (
            <div style={{ animation: "fadeIn .3s" }}>
              <button onClick={() => setEditing(null)} style={S.btnGhost}>← Voltar</button>
              <div style={{ ...S.card, marginTop: 12, padding: 16 }}>
                <h3 style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>Dados da Empresa</h3>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  {[["name", "Nome", "Ex: Elmafe"], ["nif", "NIF", "501234567"], ["sector", "Setor", "Ex: Estruturas Metálicas"]].map(([k, l, ph]) => (
                    <div key={k}><label style={S.label}>{l}</label><input value={editing[k] || ""} onChange={e => setEditing(p => ({ ...p, [k]: e.target.value }))} placeholder={ph} style={S.input} /></div>
                  ))}
                  <div><label style={S.label}>Instruções IA</label><input value={editing.notes || ""} onChange={e => setEditing(p => ({ ...p, notes: e.target.value }))} placeholder="Instruções especiais" style={S.input} /></div>
                </div>
              </div>
              <div style={{ ...S.card, marginTop: 12, padding: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700 }}>Perfis ({editing.profiles?.length || 0})</h3>
                  <button onClick={() => setEditing(p => ({ ...p, profiles: [...(p.profiles || []), { id: "t" + Date.now(), label: "Novo Tipo", icon: "📄", color: "#64748b", fields: [{ key: "descricao", label: "Descrição", required: false }] }] }))} style={S.btnSmallGreen}>+ Tipo</button>
                </div>
                {(editing.profiles || []).map((prof, pi) => (
                  <div key={pi} style={{ background: "#0c0818", borderRadius: 6, padding: "8px 10px", marginBottom: 6, border: "1px solid #2a2340" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
                      <span style={{ fontSize: 16 }}>{prof.icon}</span>
                      <input value={prof.label} onChange={e => { const ps = [...editing.profiles]; ps[pi] = { ...ps[pi], label: e.target.value }; setEditing(p => ({ ...p, profiles: ps })); }} style={{ ...S.input, flex: 1, fontWeight: 600, border: "none", background: "none" }} />
                      <button onClick={() => setEditing(p => ({ ...p, profiles: p.profiles.filter((_, j) => j !== pi) }))} style={{ ...S.btnIcon, color: "#dc2626" }}>×</button>
                    </div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                      {(prof.fields || []).map((f, fi) => (
                        <span key={fi} style={S.fieldTag}>
                          <span style={{ color: f.required ? "#a855f7" : "#8b82a8" }}>{f.label}</span>
                          <button onClick={() => { const ps = [...editing.profiles]; ps[pi] = { ...ps[pi], fields: ps[pi].fields.filter((_, j) => j !== fi) }; setEditing(p => ({ ...p, profiles: ps })); }} style={{ background: "none", border: "none", color: "#5e5478", cursor: "pointer", fontSize: 10, marginLeft: 2 }}>×</button>
                        </span>
                      ))}
                      <button onClick={() => { const ps = [...editing.profiles]; ps[pi] = { ...ps[pi], fields: [...(ps[pi].fields || []), { key: "campo_" + Date.now(), label: "Novo Campo", required: false }] }; setEditing(p => ({ ...p, profiles: ps })); }} style={S.addFieldBtn}>+ Campo</button>
                    </div>
                  </div>
                ))}
              </div>
              <button onClick={() => {
                const idx = companies.findIndex(c => c.id === editing.id);
                const next = idx >= 0 ? companies.map((c, i) => i === idx ? editing : c) : [...companies, editing];
                saveCos(next);
                if (!activeCo || activeCo.id === editing.id) setActiveCo(editing);
                setEditing(null);
              }} style={{ ...S.btnProcess, marginTop: 14 }}>💾 Guardar Empresa</button>
            </div>
          )}

        </div>
      </div>

      {/* ═══ EXPORT MODAL ═══ */}
      {exportModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 20 }}
          onClick={(e) => { if (e.target === e.currentTarget) setExportModal(null); }}>
          <div style={{ background: "#110e20", borderRadius: 12, border: "1px solid #2a2340", padding: 24, maxWidth: 520, width: "100%", maxHeight: "80vh", overflow: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div>
                <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 2 }}>Exportar para ERP</h3>
                <p style={{ fontSize: 11, color: "#5e5478" }}>{exportModal.data.length} documento(s) · {exportModal.data.filter(r => r.campos?.valor_total).length} com valores</p>
              </div>
              <button onClick={() => setExportModal(null)} style={S.btnIcon}>✕</button>
            </div>

            <div style={{ display: "grid", gap: 6 }}>
              {Object.entries(ERP_FORMATS).map(([key, fmt]) => (
                <button
                  key={key}
                  onClick={() => exportForERP(exportModal.data, key)}
                  style={{
                    display: "flex", alignItems: "center", gap: 12, padding: "12px 14px",
                    background: "#0c0818", border: "1px solid #2a2340", borderRadius: 8,
                    cursor: "pointer", textAlign: "left", width: "100%",
                    transition: "border-color 0.2s",
                  }}
                  onMouseEnter={e => e.currentTarget.style.borderColor = "#a855f7"}
                  onMouseLeave={e => e.currentTarget.style.borderColor = "#2a2340"}
                >
                  <span style={{ fontSize: 24 }}>{fmt.icon}</span>
                  <div style={{ flex: 1 }}>
                    <span style={{ fontSize: 13, fontWeight: 600, color: "#f4f0ff", display: "block" }}>{fmt.label}</span>
                    <span style={{ fontSize: 11, color: "#5e5478" }}>{fmt.desc}</span>
                  </div>
                  <span style={{ fontSize: 11, color: "#a855f7", fontWeight: 600 }}>Exportar →</span>
                </button>
              ))}
            </div>

            <div style={{ marginTop: 14, padding: "10px 12px", background: "#a855f710", borderRadius: 6, border: "1px solid #a855f720" }}>
              <p style={{ fontSize: 11, color: "#a855f7", lineHeight: 1.5 }}>
                <strong>Dica:</strong> O ficheiro exportado está formatado para importação direta no ERP selecionado. Abre o módulo de Compras/Faturas do teu software e usa a função "Importar" para carregar o ficheiro.
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════
// STYLES
// ═══════════════════════════════════════
const S: Record<string, React.CSSProperties> = {
  root: { fontFamily: "'IBM Plex Sans',system-ui,sans-serif", background: "#0c0818", color: "#f4f0ff", minHeight: "100vh", display: "flex", flexDirection: "column" },
  center: { display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", background: "#0c0818", color: "#8b82a8", fontFamily: "system-ui" },
  header: { background: "#110e20", borderBottom: "1px solid #2a2340", padding: "0 16px", height: 50, display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 },
  badge: { fontSize: 11, color: "#a855f7", background: "#a855f715", padding: "3px 9px", borderRadius: 5, fontWeight: 600, border: "1px solid #a855f725" },
  select: { background: "#18122e", border: "1px solid #2a2340", borderRadius: 5, padding: "4px 6px", fontSize: 11, color: "#8b82a8", outline: "none" },
  tabs: { background: "#110e20", borderBottom: "1px solid #2a2340", display: "flex", padding: "0 16px", flexShrink: 0, overflowX: "auto" },
  tab: { background: "none", border: "none", cursor: "pointer", padding: "10px 14px", fontSize: 12, transition: "all .2s", whiteSpace: "nowrap" },
  content: { flex: 1, overflow: "auto", padding: "20px 16px 50px" },
  dropzone: { border: "2px dashed #2a2340", borderRadius: 12, padding: "36px 20px", textAlign: "center", cursor: "pointer", background: "#0c1120", transition: "border-color .2s" },
  fileRow: { display: "flex", alignItems: "center", gap: 8, padding: "7px 10px", borderRadius: 6, background: "#110e20", border: "1px solid #2a2340", marginBottom: 3 },
  btnProcess: { width: "100%", marginTop: 10, padding: "11px", background: "linear-gradient(135deg,#a855f7,#7c3aed)", color: "#0c0818", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", textAlign: "center", display: "block" },
  processing: { width: "100%", marginTop: 10, padding: "11px", background: "#110e20", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, fontSize: 12, color: "#8b82a8" },
  card: { borderRadius: 8, background: "#110e20", border: "1px solid #2a2340", marginBottom: 6, overflow: "hidden" },
  cardHead: { display: "flex", alignItems: "center", gap: 10, padding: "12px 14px", cursor: "pointer" },
  stat: { padding: "10px 12px", borderRadius: 6, background: "#110e20", border: "1px solid #2a2340", display: "flex", flexDirection: "column", alignItems: "center" },
  th: { padding: "5px 8px", textAlign: "left", fontWeight: 600, color: "#5e5478", fontSize: 9, textTransform: "uppercase" },
  td: { padding: "4px 8px" },
  ocrText: { marginTop: 6, padding: 8, borderRadius: 5, background: "#0c0818", fontSize: 10, lineHeight: 1.5, maxHeight: 140, overflowY: "auto", whiteSpace: "pre-wrap", color: "#8b82a8", fontFamily: "monospace", border: "none" },
  empty: { textAlign: "center", padding: "40px 20px", background: "#0c1120", borderRadius: 10, border: "1px dashed #2a2340" },
  companyRow: { display: "flex", alignItems: "center", gap: 10, padding: "12px 14px", borderRadius: 8, background: "#110e20", marginBottom: 6, cursor: "pointer", transition: "all .2s" },
  label: { fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "#5e5478", marginBottom: 3, display: "block" },
  input: { width: "100%", padding: "7px 9px", borderRadius: 5, border: "1px solid #2a2340", background: "#0c0818", color: "#f4f0ff", fontSize: 12, outline: "none" },
  fieldTag: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 3, background: "#18122e", border: "1px solid #2a2340", fontSize: 10 },
  addFieldBtn: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 3, border: "1px dashed #2a2340", fontSize: 10, color: "#5e5478", cursor: "pointer", background: "none" },
  btnPrimary: { padding: "8px 14px", background: "#a855f7", color: "#0c0818", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" },
  btnSecondary: { padding: "6px 10px", background: "#18122e", color: "#8b82a8", border: "1px solid #2a2340", borderRadius: 6, fontSize: 11, cursor: "pointer" },
  btnExport: { padding: "6px 10px", background: "#a855f7", color: "#0c0818", border: "none", borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: "pointer" },
  btnGhost: { background: "none", border: "none", color: "#5e5478", fontSize: 11, cursor: "pointer" },
  btnSmall: { background: "#18122e", border: "1px solid #2a2340", borderRadius: 5, padding: "4px 8px", fontSize: 10, color: "#8b82a8", cursor: "pointer" },
  btnSmallGreen: { background: "#a855f718", border: "1px solid #a855f730", borderRadius: 5, padding: "4px 8px", fontSize: 10, color: "#a855f7", cursor: "pointer", fontWeight: 600 },
  btnIcon: { background: "none", border: "none", cursor: "pointer", color: "#5e5478", fontSize: 14, padding: 2 },
};