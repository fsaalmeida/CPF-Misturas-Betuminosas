import { Fragment, jsx, jsxs } from "react/jsx-runtime";
import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, ComposedChart, Bar, Legend } from "recharts";
import { ShieldCheck, Wrench, HardHat, Plus, Pencil, Trash2, ChevronLeft, LogOut, Factory, Users, X, MapPin, Beaker, Power, PowerOff, Upload, FileSpreadsheet, CheckCircle2, ClipboardList, Mountain, Droplet, PackagePlus, Box, ChevronRight, Construction, Building2, FlaskConical, Search, ArrowUpDown, AlertTriangle, Calculator, Eye, Package, TrendingUp, History, Truck, Copy, Grid3x3, Settings, Fuel, Archive, FileText, Image } from "lucide-react";
const ROLES = ["Administrador", "Gestor", "Operador", "Or\xE7amentista", "Convidado"];
const PERMISSOES_DEFEITO = {
  "Administrador": { centrosProducao: { visualizar: true, editar: true } },
  "Gestor": { centrosProducao: { visualizar: true, editar: true } },
  "Operador": { centrosProducao: { visualizar: true, editar: false } },
  "Or\xE7amentista": { centrosProducao: { visualizar: true, editar: false } },
  "Convidado": { centrosProducao: { visualizar: true, editar: false } }
};
const TIPOS_MATERIAL_DEFEITO = ["Agregado", "Betume", "Filler Comercial", "Aditivo", "Outro"];
const TIPOS_DESCONTO_DEFEITO = ["ISP", "Desconto comercial", "B\xF3nus bom pagamento", "B\xF3nus quantidade"];
const TIPOS_CUSTO_EXTRA_DEFEITO = ["Transporte", "Otimiza\xE7\xE3o"];
const roleStyle = {
  "Administrador": { bg: "bg-amber-100", text: "text-amber-800", icon: ShieldCheck },
  "Gestor": { bg: "bg-sky-100", text: "text-sky-800", icon: Wrench },
  "Operador": { bg: "bg-stone-200", text: "text-stone-700", icon: HardHat },
  "Or\xE7amentista": { bg: "bg-emerald-100", text: "text-emerald-800", icon: Calculator },
  "Convidado": { bg: "bg-stone-100", text: "text-stone-500", icon: Eye }
};
const genId = () => Math.random().toString(36).slice(2, 10);
const formatArticleCode = (center, m) => m?.codigoManual || (center?.codigo && m?.numero ? `${center.codigo}${String(m.numero).padStart(2, "0")}` : null);
const normalizeHeader = (h) => String(h ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
const isSocorpenaCliente = (c) => !!c && (String(c.numero || "").trim() === "9999" || normalizeHeader(c.designacao).includes("socorpena"));
const normalizeIncidencias = (val) => {
  if (Array.isArray(val)) return val;
  if (typeof val === "string" && val.trim()) return [{ id: genId(), descricao: val.trim(), resolucaoData: "", resolucaoDescricao: "" }];
  return [];
};
const exportarListaExcel = (nomeFicheiro, nomeFolha, headers, linhas) => {
  const wsData = [headers, ...linhas];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = headers.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  const nomeFolhaSeguro = (nomeFolha || "Dados").replace(/[:\\/?*[\]]/g, "-").slice(0, 31) || "Dados";
  XLSX.utils.book_append_sheet(wb, ws, nomeFolhaSeguro);
  XLSX.writeFile(wb, nomeFicheiro);
};
const matchesSearch = (query, ...fields) => {
  const words = normalizeHeader(query).split(/\s+/).filter(Boolean);
  if (words.length === 0) return true;
  const target = normalizeHeader(fields.filter(Boolean).join(" "));
  return words.every((w) => target.includes(w));
};
const abreviarFornecedor = (nome) => {
  if (!nome) return "";
  const primeira = nome.split(/[\s,;-]+/).filter(Boolean)[0];
  return primeira || nome;
};
const materialLabel = (m) => m?.fornecedor ? `${m.designacao} (${abreviarFornecedor(m.fornecedor)})` : m?.designacao || "";
const DESCONTO_CATEGORIAS = ["Desconto comercial", "B\xF3nus bom pagamento", "B\xF3nus quantidade", "Outro"];
const DESCONTO_CATEGORIAS_CONSUMIVEL = ["ISP", "Desconto comercial", "B\xF3nus bom pagamento", "B\xF3nus quantidade", "Outro"];
const UNIDADES_CUSTO_EQUIPAMENTO = ["\u20AC/hora", "\u20AC/dia", "\u20AC/tonelada", "\u20AC/km"];
const normUnidade = (u) => String(u || "").replace(/\s+/g, "").toLowerCase();
const UNIDADES_CUSTO_MAO_OBRA = ["\u20AC/hora", "\u20AC/dia", "\u20AC/m\xEAs"];
const UNIDADES_CONSUMO_COMBUSTIVEL = ["Lt/Ton", "Kg/Ton", "Ton/Ton", "m\xB3/Ton", "kWh/Ton"];
const unidadeBaseCombustivel = (unidadeConsumo) => (unidadeConsumo || "Lt/Ton").split("/")[0];
const custosExtraLista = (m) => {
  if (Array.isArray(m?.custosExtra) && m.custosExtra.length > 0) return m.custosExtra;
  const legado = [];
  if (parseFloat(m?.custoTransporte) > 0) legado.push({ nome: "Transporte", valor: m.custoTransporte });
  if (parseFloat(m?.custoOtimizacao) > 0) legado.push({ nome: "Otimiza\xE7\xE3o", valor: m.custoOtimizacao });
  return legado;
};
const calcularPrecoFinal = (m) => {
  let preco = parseFloat(m.preco) || 0;
  (m.descontos || []).forEach((d) => {
    if (d.aplicarNoCalculo === false) return;
    const valor = parseFloat(d.valor) || 0;
    preco = d.tipo === "fixo" ? Math.max(0, preco - valor) : Math.max(0, preco * (1 - valor / 100));
  });
  const extra = custosExtraLista(m).reduce((s, c) => s + (parseFloat(c.valor) || 0), 0);
  return preco + extra;
};
const normalizarHistoricoPrecos = (material) => {
  if (Array.isArray(material?.historicoPrecos) && material.historicoPrecos.length > 0) return material.historicoPrecos;
  if (material?.preco !== void 0 && material?.preco !== null && material?.preco !== "") {
    return [{
      id: `legacy-${material.id}`,
      preco: material.preco,
      descontos: material.descontos || [],
      custoTransporte: material.custoTransporte || 0,
      dataEntradaVigor: material.dataEntradaVigor || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10)
    }];
  }
  return [];
};
const precoVigente = (material, dataRef) => {
  const historico = normalizarHistoricoPrecos(material);
  if (historico.length === 0) return null;
  if (!dataRef) {
    return [...historico].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""))[0];
  }
  const aplicaveis = historico.filter((p) => p.dataEntradaVigor && p.dataEntradaVigor <= dataRef).sort((a, b) => b.dataEntradaVigor.localeCompare(a.dataEntradaVigor));
  if (aplicaveis.length > 0) return aplicaveis[0];
  return [...historico].sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""))[0];
};
const historicoOrdenado = (material) => [...normalizarHistoricoPrecos(material)].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
const taxaVigenteEmData = (historico, dataRef) => {
  const lista = historico || [];
  if (lista.length === 0) return null;
  if (!dataRef) return [...lista].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""))[0];
  const aplicaveis = lista.filter((h) => h.dataEntradaVigor && h.dataEntradaVigor <= dataRef).sort((a, b) => b.dataEntradaVigor.localeCompare(a.dataEntradaVigor));
  if (aplicaveis.length > 0) return aplicaveis[0];
  return [...lista].sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""))[0];
};
const VALIDADE_ANOS = 5;
const PENEIROS_DOP = ["40", "31,5", "20", "16", "14", "12,5", "10", "8", "6,3", "4", "2", "1", "0,5", "0,25", "0,125", "0,063"];
const SILO_COLS = [
  { key: "s1", label: "Silo 1" },
  { key: "s2", label: "Silo 2" },
  { key: "s3", label: "Silo 3" },
  { key: "s4", label: "Silo 4" },
  { key: "s5", label: "Silo 5" },
  { key: "s6", label: "Silo 6" }
];
const TRABALHO_COLS = [
  { key: "bypass", label: "By-passe" },
  { key: "crivo1", label: "", num: 1 },
  { key: "crivo2", label: "17/25", num: 2 },
  { key: "crivo3", label: "10/17", num: 3 },
  { key: "crivo4", label: "5/10", num: 4 },
  { key: "crivo5", label: "0/5", num: 5 },
  { key: "fillerRec", label: "Filler Rec." },
  { key: "fillerCom", label: "Filler Com." },
  { key: "aditivo1", label: "Aditivo 1" },
  { key: "aditivo2", label: "Aditivo 2" },
  { key: "aditivo3", label: "Aditivo 3" },
  { key: "ligante", label: "Ligante" }
];
const TRABALHO_MATERIAL_KEYS = ["fillerCom", "aditivo1", "aditivo2", "aditivo3", "ligante"];
const DIAS_UTEIS_ANO_PADRAO = 264;
const calcularMovimentosStock = ({ center, produtoId, categoria, rececoes, diarias, formulas, ajustesStock }) => {
  const movimentos = [];
  (rececoes || []).filter((r) => r.centroId === center.id && r.categoria === categoria && (r.produtoId === produtoId && !r.substituiId || r.substituiId === produtoId)).forEach((r) => {
    const tipo = r.substituiId === produtoId ? "Rece\xE7\xE3o (substituto)" : "Rece\xE7\xE3o";
    movimentos.push({ id: `rec-${r.id}`, data: r.data, tipo, quantidade: parseFloat(r.quantidade) || 0, produtoRecebidoId: r.substituiId === produtoId ? r.produtoId : null });
  });
  if (categoria !== "consumiveis") {
    const formulasPorId = Object.fromEntries((formulas || []).map((f) => [f.id, f]));
    (diarias || []).filter((d) => d.centroId === center.id).forEach((d) => {
      (d.linhas || []).forEach((l) => {
        const formula = formulasPorId[l.artigoId];
        const toneladas = parseFloat(l.toneladas) || 0;
        if (!formula || toneladas <= 0) return;
        const kgNaoAgregadoPorTonelada = TRABALHO_MATERIAL_KEYS.reduce((s, key) => s + (parseFloat(formula.trabalho?.[key]?.design) || 0), 0);
        const kgAgregadoPorTonelada = Math.max(0, 1e3 - kgNaoAgregadoPorTonelada);
        SILO_COLS.forEach((c) => {
          const silo = formula.silos?.[c.key];
          const pct = parseFloat(silo?.pct) || 0;
          if (silo?.materialId === produtoId && pct > 0) {
            const kgPorTonelada = pct / 100 * kgAgregadoPorTonelada;
            movimentos.push({ id: `diaria-${d.id}-${c.key}`, data: d.dataInicio, tipo: "Consumo (Produ\xE7\xE3o)", quantidade: -(toneladas * (kgPorTonelada / 1e3)) });
          }
        });
        TRABALHO_COLS.forEach((c) => {
          const item = formula.trabalho?.[c.key];
          const kgPorTonelada = parseFloat(item?.design) || 0;
          if (item?.materialId === produtoId && kgPorTonelada > 0) {
            movimentos.push({ id: `diaria-${d.id}-${c.key}`, data: d.dataInicio, tipo: "Consumo (Produ\xE7\xE3o)", quantidade: -(toneladas * (kgPorTonelada / 1e3)) });
          }
        });
      });
    });
  }
  if (categoria === "consumiveis" && center.parametrizacao?.blocoTermico?.combustivelId === produtoId) {
    const historicoTaxa = center.parametrizacao.blocoTermico.historico || [];
    (diarias || []).filter((d) => d.centroId === center.id).forEach((d) => {
      const toneladas = (d.linhas || []).reduce((s, l) => s + (parseFloat(l.toneladas) || 0), 0);
      if (toneladas <= 0) return;
      const taxa = taxaVigenteEmData(historicoTaxa, d.dataInicio);
      if (!taxa) return;
      movimentos.push({ id: `diaria-${d.id}-bloco`, data: d.dataInicio, tipo: "Consumo (Bloco T\xE9rmico)", quantidade: -(toneladas * (parseFloat(taxa.valor) || 0)) });
    });
  }
  (ajustesStock || []).filter((a) => a.centroId === center.id && a.produtoId === produtoId && a.categoria === categoria).forEach((a) => {
    movimentos.push({ id: a.id, data: a.data, tipo: "Ajuste", quantidade: parseFloat(a.quantidade) || 0, motivo: a.motivo, utilizador: a.utilizador, dataRegisto: a.dataRegisto });
  });
  movimentos.sort((a, b) => (a.data || "").localeCompare(b.data || ""));
  let saldo = 0;
  return movimentos.map((m) => {
    saldo += m.quantidade;
    return { ...m, saldo };
  });
};
const getTrabalhoLabel = (center, key) => center?.trabalhoLabels?.[key] ?? TRABALHO_COLS.find((c) => c.key === key)?.label ?? key;
const emptyDosagem = () => Object.fromEntries([...SILO_COLS, ...TRABALHO_COLS].map((c) => [c.key, { design: "", pct: "" }]));
const calcValidade = (dataExecucao) => {
  if (!dataExecucao) return null;
  const d = /* @__PURE__ */ new Date(dataExecucao + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  d.setFullYear(d.getFullYear() + VALIDADE_ANOS);
  return d.toISOString().slice(0, 10);
};
const formatDatePT = (iso) => {
  if (!iso) return "\u2014";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const formatDateTimePT = (iso) => {
  if (!iso) return "\u2014";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "\u2014";
  return d.toLocaleString("pt-PT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
};
const isoToDatetimeLocal = (iso) => {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
const datetimeLocalToIso = (v) => {
  if (!v) return null;
  const [datePart] = v.split("T");
  const [y, m, d] = (datePart || "").split("-").map(Number);
  const dt = new Date(v);
  if (isNaN(dt.getTime())) return null;
  if (dt.getFullYear() !== y || dt.getMonth() + 1 !== m || dt.getDate() !== d) return null;
  return dt.toISOString();
};
const isDataCalendarioValida = (v) => {
  if (!v) return false;
  const [y, m, d] = v.split("-").map(Number);
  const dt = /* @__PURE__ */ new Date(`${v}T00:00:00Z`);
  if (isNaN(dt.getTime())) return false;
  return dt.getUTCFullYear() === y && dt.getUTCMonth() + 1 === m && dt.getUTCDate() === d;
};
function RoleBadge({ role }) {
  const s = roleStyle[role] || roleStyle["Operador"];
  const Icon = s.icon;
  return /* @__PURE__ */ jsxs("span", { className: `inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold ${s.bg} ${s.text}`, children: [
    /* @__PURE__ */ jsx(Icon, { size: 13, strokeWidth: 2.5 }),
    role
  ] });
}
function LoginForm({ users, onLogin, onDefinePassword }) {
  const [stage, setStage] = useState("username");
  const [username, setUsername] = useState("");
  const [matchedUser, setMatchedUser] = useState(null);
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [lembrar, setLembrar] = useState(true);
  const [error, setError] = useState("");
  const digitsOnly = (v) => v.replace(/\D/g, "").slice(0, 4);
  const submitUsername = () => {
    setError("");
    if (!username.trim()) return setError("Indique o nome de utilizador");
    const found = users.find((u) => normalizeHeader(u.nome) === normalizeHeader(username));
    if (!found) return setError("Utilizador n\xE3o encontrado");
    setMatchedUser(found);
    setStage(found.pin ? "password" : "setup");
  };
  const submitPassword = () => {
    setError("");
    if (password.length !== 4) return setError("A palavra-passe tem 4 n\xFAmeros");
    const ok = onLogin(matchedUser, password, lembrar);
    if (ok === false) {
      setError("Palavra-passe incorreta");
      setPassword("");
    }
  };
  const submitSetup = () => {
    setError("");
    if (password.length !== 4) return setError("A palavra-passe deve ter 4 n\xFAmeros");
    if (password !== confirmPassword) return setError("As palavras-passe n\xE3o coincidem");
    onDefinePassword(matchedUser, password);
  };
  const irParaEsqueci = () => {
    setStage("setup");
    setPassword("");
    setConfirmPassword("");
    setError("");
  };
  const voltar = () => {
    setStage("username");
    setMatchedUser(null);
    setPassword("");
    setConfirmPassword("");
    setError("");
  };
  return /* @__PURE__ */ jsxs("div", { className: "bg-stone-900 border border-stone-800 rounded-xl p-6", children: [
    stage === "username" && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Nome de utilizador" }), children: /* @__PURE__ */ jsx(
        "input",
        {
          value: username,
          onChange: (e) => setUsername(e.target.value),
          onKeyDown: (e) => e.key === "Enter" && submitUsername(),
          className: inputCls,
          placeholder: "O seu nome",
          autoFocus: true
        }
      ) }),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-500 text-sm mb-3", children: error }),
      /* @__PURE__ */ jsx("button", { onClick: submitUsername, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Continuar" })
    ] }),
    stage === "password" && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-stone-400 text-sm mb-3", children: [
        "Ol\xE1, ",
        /* @__PURE__ */ jsx("span", { className: "text-white font-medium", children: matchedUser.nome })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Palavra-passe (4 n\xFAmeros)" }), children: /* @__PURE__ */ jsx(
        "input",
        {
          value: password,
          onChange: (e) => setPassword(digitsOnly(e.target.value)),
          onKeyDown: (e) => e.key === "Enter" && submitPassword(),
          type: "password",
          inputMode: "numeric",
          className: `${inputCls} font-mono-data tracking-widest`,
          placeholder: "\u2022\u2022\u2022\u2022",
          autoFocus: true
        }
      ) }),
      /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 mb-4 cursor-pointer", children: [
        /* @__PURE__ */ jsx("input", { type: "checkbox", checked: lembrar, onChange: (e) => setLembrar(e.target.checked), className: "w-4 h-4 accent-amber-600 cursor-pointer" }),
        /* @__PURE__ */ jsx("span", { className: "text-sm text-stone-400", children: "Lembrar-me neste dispositivo por 24 horas" })
      ] }),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-500 text-sm mb-3", children: error }),
      /* @__PURE__ */ jsx("button", { onClick: submitPassword, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Entrar" }),
      /* @__PURE__ */ jsx("button", { onClick: irParaEsqueci, className: "w-full mt-3 py-1 text-sm text-amber-500 hover:text-amber-400", children: "Esqueci-me da palavra-passe" }),
      /* @__PURE__ */ jsx("button", { onClick: voltar, className: "w-full py-2 text-sm text-stone-500 hover:text-stone-300", children: "Voltar" })
    ] }),
    stage === "setup" && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-stone-400 text-sm mb-3", children: [
        "Ol\xE1, ",
        /* @__PURE__ */ jsx("span", { className: "text-white font-medium", children: matchedUser.nome }),
        " \u2014 escolha uma nova palavra-passe de 4 n\xFAmeros"
      ] }),
      /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Nova palavra-passe" }), children: /* @__PURE__ */ jsx(
        "input",
        {
          value: password,
          onChange: (e) => setPassword(digitsOnly(e.target.value)),
          type: "password",
          inputMode: "numeric",
          className: `${inputCls} font-mono-data tracking-widest`,
          placeholder: "\u2022\u2022\u2022\u2022",
          autoFocus: true
        }
      ) }),
      /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Confirmar palavra-passe" }), children: /* @__PURE__ */ jsx(
        "input",
        {
          value: confirmPassword,
          onChange: (e) => setConfirmPassword(digitsOnly(e.target.value)),
          onKeyDown: (e) => e.key === "Enter" && submitSetup(),
          type: "password",
          inputMode: "numeric",
          className: `${inputCls} font-mono-data tracking-widest`,
          placeholder: "\u2022\u2022\u2022\u2022"
        }
      ) }),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-500 text-sm mb-3", children: error }),
      /* @__PURE__ */ jsx("button", { onClick: submitSetup, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Confirmar e entrar" }),
      /* @__PURE__ */ jsx("button", { onClick: voltar, className: "w-full mt-2 py-2 text-sm text-stone-500 hover:text-stone-300", children: "Voltar" })
    ] })
  ] });
}
function ConfirmDialog({ message, onConfirm, onCancel }) {
  return /* @__PURE__ */ jsx("div", { className: "fixed inset-0 bg-stone-950/60 flex items-center justify-center z-[60] p-4", children: /* @__PURE__ */ jsx("div", { className: "bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden border border-stone-200", children: /* @__PURE__ */ jsxs("div", { className: "p-6", children: [
    /* @__PURE__ */ jsx("div", { className: "w-10 h-10 rounded-full bg-red-50 flex items-center justify-center mb-4", children: /* @__PURE__ */ jsx(Trash2, { className: "text-red-600", size: 18 }) }),
    /* @__PURE__ */ jsx("p", { className: "text-stone-800 font-medium mb-6", children: message }),
    /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
      /* @__PURE__ */ jsx("button", { onClick: onCancel, className: "flex-1 py-2.5 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Cancelar" }),
      /* @__PURE__ */ jsx("button", { onClick: onConfirm, className: "flex-1 py-2.5 rounded-lg bg-red-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-red-700", children: "Remover" })
    ] })
  ] }) }) });
}
function Modal({ title, subtitle, onClose, children, wide, fullscreen }) {
  if (fullscreen) {
    return /* @__PURE__ */ jsx("div", { className: "fixed top-0 right-0 bottom-0 left-60 bg-stone-950/60 flex z-50 p-4", children: /* @__PURE__ */ jsxs("div", { className: "bg-stone-50 rounded-2xl shadow-2xl w-full h-full overflow-hidden border border-stone-200 flex flex-col", children: [
      /* @__PURE__ */ jsxs("div", { className: "bg-stone-900 px-6 py-5 relative shrink-0", children: [
        /* @__PURE__ */ jsx("button", { onClick: onClose, className: "absolute top-4 right-4 text-stone-400 hover:text-white", children: /* @__PURE__ */ jsx(X, { size: 18 }) }),
        subtitle && /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-500 uppercase mb-1", children: subtitle }),
        /* @__PURE__ */ jsx("h3", { className: "font-display text-xl text-white font-semibold", children: title })
      ] }),
      /* @__PURE__ */ jsx("div", { className: "p-6 overflow-y-auto flex-1 min-h-0", children })
    ] }) });
  }
  return /* @__PURE__ */ jsx("div", { className: "fixed inset-0 bg-stone-950/60 flex items-center justify-center z-50 p-4", children: /* @__PURE__ */ jsxs("div", { className: `bg-stone-50 rounded-2xl shadow-2xl w-full ${wide ? "max-w-3xl" : "max-w-md"} overflow-hidden border border-stone-200 max-h-screen flex flex-col`, children: [
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-900 px-6 py-5 relative shrink-0", children: [
      /* @__PURE__ */ jsx("button", { onClick: onClose, className: "absolute top-4 right-4 text-stone-400 hover:text-white", children: /* @__PURE__ */ jsx(X, { size: 18 }) }),
      subtitle && /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-500 uppercase mb-1", children: subtitle }),
      /* @__PURE__ */ jsx("h3", { className: "font-display text-xl text-white font-semibold", children: title })
    ] }),
    /* @__PURE__ */ jsx("div", { className: "p-6 overflow-y-auto flex-1 min-h-0", children })
  ] }) });
}
function Field({ label, children }) {
  return /* @__PURE__ */ jsxs("label", { className: "block mb-4", children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: label }),
    children
  ] });
}
const inputCls = "w-full px-3 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500 text-sm";
function App() {
  const [loading, setLoading] = useState(true);
  const [users, setUsers] = useState([]);
  const [centers, setCenters] = useState([]);
  const [mixtures, setMixtures] = useState([]);
  const [proveniencias, setProveniencias] = useState([]);
  const [diarias, setDiarias] = useState([]);
  const [formulas, setFormulas] = useState([]);
  const [avarias, setAvarias] = useState([]);
  const [materiais, setMateriais] = useState([]);
  const [consumiveis, setConsumiveis] = useState([]);
  const [equipamentos, setEquipamentos] = useState([]);
  const [maoDeObra, setMaoDeObra] = useState([]);
  const [rececoes, setRececoes] = useState([]);
  const [ajustesStock, setAjustesStock] = useState([]);
  const [logotipo, setLogotipoState] = useState("");
  const [dopConfig, setDopConfig] = useState(null);
  const [clientes, setClientes] = useState([]);
  const [fornecedores, setFornecedores] = useState([]);
  const [centrosCusto, setCentrosCusto] = useState([]);
  const [permissoes, setPermissoes] = useState(PERMISSOES_DEFEITO);
  const [tiposMaterial, setTiposMaterial] = useState([]);
  const [tiposDesconto, setTiposDesconto] = useState([]);
  const [tiposCustoExtra, setTiposCustoExtra] = useState([]);
  const [perfisPersonalizados, setPerfisPersonalizados] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);
  const [activeCenterId, setActiveCenterId] = useState(null);
  const [view, setView] = useState("centers");
  const [selectedCenterId, setSelectedCenterId] = useState(null);
  const [selectedClienteId, setSelectedClienteId] = useState(null);
  const [modal, setModal] = useState(null);
  const [saveError, setSaveError] = useState("");
  const writeQueueRef = useRef({});
  const [confirmDialog, setConfirmDialog] = useState(null);
  const askConfirm = (message, onConfirm) => setConfirmDialog({ message, onConfirm });
  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [u, c, m, p, cl, cc, d, f, av, mat, forn, cons, rec, equip, ajust, logo, dop, mo, perm, tipMat, tipDesc, tipCustoExtra, perfisPers] = await Promise.allSettled([
        window.storage.get("users", true),
        window.storage.get("centers", true),
        window.storage.get("mixtures", true),
        window.storage.get("proveniencias", true),
        window.storage.get("clientes", true),
        window.storage.get("centrosCusto", true),
        window.storage.get("diarias", true),
        window.storage.get("formulas", true),
        window.storage.get("avarias", true),
        window.storage.get("materiais", true),
        window.storage.get("fornecedores", true),
        window.storage.get("consumiveis", true),
        window.storage.get("rececoes", true),
        window.storage.get("equipamentos", true),
        window.storage.get("ajustesStock", true),
        window.storage.get("logotipo", true),
        window.storage.get("dopConfig", true),
        window.storage.get("maoDeObra", true),
        window.storage.get("permissoes", true),
        window.storage.get("tiposMaterial", true),
        window.storage.get("tiposDesconto", true),
        window.storage.get("tiposCustoExtra", true),
        window.storage.get("perfisPersonalizados", true)
      ]);
      setUsers(u.status === "fulfilled" && u.value ? JSON.parse(u.value.value) : []);
      setCenters(c.status === "fulfilled" && c.value ? JSON.parse(c.value.value) : []);
      setMixtures(m.status === "fulfilled" && m.value ? JSON.parse(m.value.value) : []);
      setProveniencias(p.status === "fulfilled" && p.value ? JSON.parse(p.value.value) : []);
      setClientes(cl.status === "fulfilled" && cl.value ? JSON.parse(cl.value.value) : []);
      setCentrosCusto(cc.status === "fulfilled" && cc.value ? JSON.parse(cc.value.value) : []);
      setDiarias(d.status === "fulfilled" && d.value ? JSON.parse(d.value.value) : []);
      setFormulas(f.status === "fulfilled" && f.value ? JSON.parse(f.value.value) : []);
      setAvarias(av.status === "fulfilled" && av.value ? JSON.parse(av.value.value) : []);
      setMateriais(mat.status === "fulfilled" && mat.value ? JSON.parse(mat.value.value) : []);
      setFornecedores(forn.status === "fulfilled" && forn.value ? JSON.parse(forn.value.value) : []);
      setConsumiveis(cons.status === "fulfilled" && cons.value ? JSON.parse(cons.value.value) : []);
      setRececoes(rec.status === "fulfilled" && rec.value ? JSON.parse(rec.value.value) : []);
      setEquipamentos(equip.status === "fulfilled" && equip.value ? JSON.parse(equip.value.value) : []);
      setAjustesStock(ajust.status === "fulfilled" && ajust.value ? JSON.parse(ajust.value.value) : []);
      setLogotipoState(logo.status === "fulfilled" && logo.value ? JSON.parse(logo.value.value) : "");
      setDopConfig(dop.status === "fulfilled" && dop.value ? JSON.parse(dop.value.value) : null);
      setMaoDeObra(mo.status === "fulfilled" && mo.value ? JSON.parse(mo.value.value) : []);
      setPerfisPersonalizados(perfisPers.status === "fulfilled" && perfisPers.value ? JSON.parse(perfisPers.value.value) : []);
      const permissoesGuardadas = perm.status === "fulfilled" && perm.value ? JSON.parse(perm.value.value) : {};
      const permissoesFinais = {};
      ROLES.forEach((r) => {
        permissoesFinais[r] = { ...PERMISSOES_DEFEITO[r], ...permissoesGuardadas[r] || {} };
      });
      setPermissoes(permissoesFinais);
      const tiposGuardados = tipMat.status === "fulfilled" && tipMat.value ? JSON.parse(tipMat.value.value) : null;
      if (tiposGuardados && tiposGuardados.length > 0) {
        setTiposMaterial(tiposGuardados);
      } else {
        const tiposSemente = TIPOS_MATERIAL_DEFEITO.map((nome) => ({ id: genId(), nome }));
        setTiposMaterial(tiposSemente);
        window.storage.set("tiposMaterial", JSON.stringify(tiposSemente), true).catch(() => {
        });
      }
      const descontosGuardados = tipDesc.status === "fulfilled" && tipDesc.value ? JSON.parse(tipDesc.value.value) : null;
      if (descontosGuardados && descontosGuardados.length > 0) {
        setTiposDesconto(descontosGuardados);
      } else {
        const descontosSemente = TIPOS_DESCONTO_DEFEITO.map((nome) => ({ id: genId(), nome }));
        setTiposDesconto(descontosSemente);
        window.storage.set("tiposDesconto", JSON.stringify(descontosSemente), true).catch(() => {
        });
      }
      const custosExtraGuardados = tipCustoExtra.status === "fulfilled" && tipCustoExtra.value ? JSON.parse(tipCustoExtra.value.value) : null;
      if (custosExtraGuardados && custosExtraGuardados.length > 0) {
        setTiposCustoExtra(custosExtraGuardados);
      } else {
        const custosExtraSemente = TIPOS_CUSTO_EXTRA_DEFEITO.map((nome) => ({ id: genId(), nome }));
        setTiposCustoExtra(custosExtraSemente);
        window.storage.set("tiposCustoExtra", JSON.stringify(custosExtraSemente), true).catch(() => {
        });
      }
    } catch (e) {
      console.error(e);
    }
    setLoading(false);
  }, []);
  useEffect(() => {
    load();
  }, [load]);
  useEffect(() => {
    if (currentUser || loading || users.length === 0) return;
    try {
      const guardada = JSON.parse(localStorage.getItem("sessaoLembrada") || "null");
      if (guardada && guardada.expiraEm > Date.now()) {
        const user = users.find((u) => u.id === guardada.userId);
        if (user) setCurrentUser(user);
        else localStorage.removeItem("sessaoLembrada");
      } else if (guardada) {
        localStorage.removeItem("sessaoLembrada");
      }
    } catch {
      localStorage.removeItem("sessaoLembrada");
    }
  }, [users, loading, currentUser]);
  const persist = (key, value, setter) => {
    setter(value);
    const anterior = writeQueueRef.current[key] || Promise.resolve();
    const atual = anterior.catch(() => {
    }).then(() => window.storage.set(key, JSON.stringify(value), true)).catch((e) => {
      setSaveError("N\xE3o foi poss\xEDvel guardar. Tente novamente.");
      console.error(e);
    });
    writeQueueRef.current[key] = atual;
    return atual;
  };
  const exportarBackupCompleto = () => {
    const dados = {
      versaoBackup: 1,
      exportadoEm: (/* @__PURE__ */ new Date()).toISOString(),
      users,
      centers,
      mixtures,
      proveniencias,
      clientes,
      centrosCusto,
      diarias,
      formulas,
      avarias,
      materiais,
      fornecedores,
      consumiveis,
      rececoes,
      equipamentos,
      maoDeObra,
      ajustesStock,
      logotipo,
      dopConfig,
      tiposMaterial,
      tiposDesconto,
      tiposCustoExtra,
      perfisPersonalizados
    };
    const blob = new Blob([JSON.stringify(dados, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `backup_gestao_centrais_${(/* @__PURE__ */ new Date()).toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const importarBackupCompleto = (dados) => {
    const mapa = [
      ["users", dados.users, setUsers],
      ["centers", dados.centers, setCenters],
      ["mixtures", dados.mixtures, setMixtures],
      ["proveniencias", dados.proveniencias, setProveniencias],
      ["clientes", dados.clientes, setClientes],
      ["centrosCusto", dados.centrosCusto, setCentrosCusto],
      ["diarias", dados.diarias, setDiarias],
      ["formulas", dados.formulas, setFormulas],
      ["avarias", dados.avarias, setAvarias],
      ["materiais", dados.materiais, setMateriais],
      ["fornecedores", dados.fornecedores, setFornecedores],
      ["consumiveis", dados.consumiveis, setConsumiveis],
      ["rececoes", dados.rececoes, setRececoes],
      ["equipamentos", dados.equipamentos, setEquipamentos],
      ["maoDeObra", dados.maoDeObra, setMaoDeObra],
      ["ajustesStock", dados.ajustesStock, setAjustesStock],
      ["logotipo", dados.logotipo, setLogotipoState],
      ["dopConfig", dados.dopConfig, setDopConfig],
      ["tiposMaterial", dados.tiposMaterial, setTiposMaterial],
      ["tiposDesconto", dados.tiposDesconto, setTiposDesconto],
      ["tiposCustoExtra", dados.tiposCustoExtra, setTiposCustoExtra],
      ["perfisPersonalizados", dados.perfisPersonalizados, setPerfisPersonalizados]
    ];
    mapa.forEach(([key, value, setter]) => {
      if (value !== void 0) persist(key, value, setter);
    });
    setModal(null);
  };
  const isAdmin = currentUser?.role === "Administrador";
  const canManageArticles = isAdmin || currentUser?.role === "Gestor";
  const podeRegistar = canManageArticles || currentUser?.role === "Operador";
  const podeVerCustos = canManageArticles || currentUser?.role === "Or\xE7amentista";
  const podeVerClientesFornecedores = canManageArticles || currentUser?.role === "Or\xE7amentista" || currentUser?.role === "Operador";
  const podeCriarFornecedores = canManageArticles || currentUser?.role === "Or\xE7amentista";
  const podeGerirObras = canManageArticles || currentUser?.role === "Or\xE7amentista";
  const podeVerMateriais = isAdmin || currentUser?.role === "Or\xE7amentista";
  const podeAtualizarPrecoMateriais = isAdmin || currentUser?.role === "Or\xE7amentista";
  const temPermissao = (menu, acao) => {
    if (isAdmin) return true;
    return !!permissoes[currentUser?.role]?.[menu]?.[acao];
  };
  const podeVerCentros = temPermissao("centrosProducao", "visualizar");
  const podeEditarCentros = temPermissao("centrosProducao", "editar");
  const attemptLogin = (user, pin, lembrar) => {
    if (user.pin === pin) {
      setCurrentUser(user);
      setView("centers");
      if (lembrar) {
        localStorage.setItem("sessaoLembrada", JSON.stringify({ userId: user.id, expiraEm: Date.now() + 24 * 60 * 60 * 1e3 }));
      } else {
        localStorage.removeItem("sessaoLembrada");
      }
      return true;
    }
    return false;
  };
  const definePin = (user, pin) => {
    setUsers((prev) => {
      const updated = { ...prev.find((u) => u.id === user.id) || user, pin };
      const atualizado = prev.map((u) => u.id === user.id ? updated : u);
      persistRaw("users", atualizado);
      setCurrentUser(updated);
      localStorage.setItem("sessaoLembrada", JSON.stringify({ userId: user.id, expiraEm: Date.now() + 24 * 60 * 60 * 1e3 }));
      return atualizado;
    });
    setView("centers");
  };
  const resetUserPin = (id) => {
    const target = users.find((u) => u.id === id);
    askConfirm(`Repor a palavra-passe de "${target?.nome}"? Vai definir uma nova no pr\xF3ximo acesso.`, () => {
      setUsers((prev) => {
        const atualizado = prev.map((u) => u.id === id ? { ...u, pin: null } : u);
        persistRaw("users", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const setUserPin = (id, novoPin) => {
    setUsers((prev) => {
      const atualizado = prev.map((u) => u.id === id ? { ...u, pin: novoPin } : u);
      persistRaw("users", atualizado);
      return atualizado;
    });
  };
  const createFirstAdmin = (nome, email, pin) => {
    const admin = { id: genId(), nome, email, pin, role: "Administrador" };
    persist("users", [admin], setUsers);
    setCurrentUser(admin);
    setPinTarget(null);
  };
  const logout = () => {
    setCurrentUser(null);
    setActiveCenterId(null);
    setView("centers");
    localStorage.removeItem("sessaoLembrada");
  };
  const saveUser = (data) => {
    setUsers((prev) => {
      const atualizado = data.id ? prev.map((u) => u.id === data.id ? { ...u, ...data } : u) : [...prev, { ...data, id: genId() }];
      console.log("[DIAGN\xD3STICO] Array completo de utilizadores a gravar no Firebase:", JSON.stringify(atualizado));
      persistRaw("users", atualizado);
      return atualizado;
    });
    setModal(null);
  };
  const deleteUser = (id) => {
    const admins = users.filter((u) => u.role === "Administrador");
    const target = users.find((u) => u.id === id);
    if (target?.role === "Administrador" && admins.length <= 1) {
      setSaveError("N\xE3o \xE9 poss\xEDvel remover o \xFAnico Administrador.");
      setTimeout(() => setSaveError(""), 4e3);
      return;
    }
    askConfirm(`Remover o utilizador "${target?.nome}"?`, () => {
      persist("users", users.filter((u) => u.id !== id), setUsers);
      setConfirmDialog(null);
    });
  };
  const updatePermissao = (role, menu, acao, valor) => {
    if (role === "Administrador") return;
    setPermissoes((prev) => {
      const atualizado = { ...prev, [role]: { ...prev[role], [menu]: { ...prev[role]?.[menu], [acao]: valor } } };
      persistRaw("permissoes", atualizado);
      return atualizado;
    });
  };
  const saveTipoMaterial = (data) => {
    setTiposMaterial((prev) => {
      const atualizado = data.id ? prev.map((t) => t.id === data.id ? { ...t, ...data } : t) : [...prev, { ...data, id: genId() }];
      persistRaw("tiposMaterial", atualizado);
      return atualizado;
    });
  };
  const deleteTipoMaterial = (id) => {
    const target = tiposMaterial.find((t) => t.id === id);
    const emUso = materiais.filter((m) => m.tipoMaterialId === id).length;
    askConfirm(
      emUso > 0 ? `Eliminar o tipo "${target?.nome}"? ${emUso} material${emUso > 1 ? "ais est\xE3o" : " est\xE1"} atualmente com este tipo e ficar\xE1(\xE3o) sem tipo definido.` : `Eliminar o tipo "${target?.nome}"?`,
      () => {
        setTiposMaterial((prev) => {
          const atualizado = prev.filter((t) => t.id !== id);
          persistRaw("tiposMaterial", atualizado);
          return atualizado;
        });
        setConfirmDialog(null);
      }
    );
  };
  const saveTipoDesconto = (data) => {
    setTiposDesconto((prev) => {
      const atualizado = data.id ? prev.map((t) => t.id === data.id ? { ...t, ...data } : t) : [...prev, { ...data, id: genId() }];
      persistRaw("tiposDesconto", atualizado);
      return atualizado;
    });
  };
  const deleteTipoDesconto = (id) => {
    const target = tiposDesconto.find((t) => t.id === id);
    askConfirm(`Eliminar o tipo de desconto "${target?.nome}"? Descontos j\xE1 registados com este nome mant\xEAm o texto, s\xF3 deixa de aparecer nas novas escolhas.`, () => {
      setTiposDesconto((prev) => {
        const atualizado = prev.filter((t) => t.id !== id);
        persistRaw("tiposDesconto", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const saveTipoCustoExtra = (data) => {
    setTiposCustoExtra((prev) => {
      const atualizado = data.id ? prev.map((t) => t.id === data.id ? { ...t, ...data } : t) : [...prev, { ...data, id: genId() }];
      persistRaw("tiposCustoExtra", atualizado);
      return atualizado;
    });
  };
  const deleteTipoCustoExtra = (id) => {
    const target = tiposCustoExtra.find((t) => t.id === id);
    askConfirm(`Eliminar o tipo de custo extra "${target?.nome}"? Custos j\xE1 registados com este nome mant\xEAm o texto, s\xF3 deixa de aparecer nas novas escolhas.`, () => {
      setTiposCustoExtra((prev) => {
        const atualizado = prev.filter((t) => t.id !== id);
        persistRaw("tiposCustoExtra", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const saveNovoPerfil = (nome) => {
    setPerfisPersonalizados((prev) => {
      const atualizado = [...prev, { id: genId(), nome }];
      persistRaw("perfisPersonalizados", atualizado);
      return atualizado;
    });
    setModal(null);
  };
  const deletePerfilPersonalizado = (id) => {
    const target = perfisPersonalizados.find((p) => p.id === id);
    const emUso = users.filter((u) => u.role === target?.nome);
    askConfirm(
      emUso.length > 0 ? `Eliminar o perfil "${target?.nome}"? ${emUso.length} utilizador${emUso.length > 1 ? "es t\xEAm" : " tem"} este perfil e ficar\xE1(\xE3o) sem acesso at\xE9 lhe(s) atribuir outro.` : `Eliminar o perfil "${target?.nome}"?`,
      () => {
        setPerfisPersonalizados((prev) => {
          const atualizado = prev.filter((p) => p.id !== id);
          persistRaw("perfisPersonalizados", atualizado);
          return atualizado;
        });
        setConfirmDialog(null);
      }
    );
  };
  const saveCenter = (data) => {
    if (data.id) {
      persist("centers", centers.map((c) => c.id === data.id ? { ...c, ...data } : c), setCenters);
    } else {
      persist("centers", [...centers, { ...data, id: genId() }], setCenters);
    }
    setModal(null);
  };
  const importCenters = (rows) => {
    const existentes = new Set(centers.map((c) => normalizeHeader(c.codigo)));
    const novos = rows.filter((r) => r.codigo && !existentes.has(normalizeHeader(r.codigo))).map((r) => ({ id: genId(), nome: r.nome, codigo: r.codigo.toUpperCase(), localizacao: r.localizacao || "", ativo: true }));
    persist("centers", [...centers, ...novos], setCenters);
    setModal(null);
  };
  const deleteCenter = (id) => {
    const target = centers.find((c) => c.id === id);
    askConfirm(`Remover o centro "${target?.nome}" e todos os seus artigos?`, () => {
      persist("centers", centers.filter((c) => c.id !== id), setCenters);
      persist("mixtures", mixtures.filter((m) => m.centroId !== id), setMixtures);
      persist("proveniencias", proveniencias.filter((p) => p.centroId !== id), setProveniencias);
      persist("diarias", diarias.filter((d) => d.centroId !== id), setDiarias);
      persist("formulas", formulas.filter((f) => f.centroId !== id), setFormulas);
      persist("avarias", avarias.filter((a) => a.centroId !== id), setAvarias);
      persist("rececoes", rececoes.filter((r) => r.centroId !== id), setRececoes);
      persist("materiais", materiais.map((m) => Array.isArray(m.centrosIds) ? { ...m, centrosIds: m.centrosIds.filter((cid) => cid !== id) } : m), setMateriais);
      setView("centers");
      setConfirmDialog(null);
    });
  };
  const toggleCenterStatus = (id) => {
    persist("centers", centers.map((c) => c.id === id ? { ...c, ativo: !(c.ativo !== false) } : c), setCenters);
  };
  const updateTrabalhoLabel = (centroId, key, label) => {
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, trabalhoLabels: { ...c.trabalhoLabels || {}, [key]: label } } : c), setCenters);
  };
  const addProducaoAnual = (centroId, ano, valor) => {
    const entrada = { id: genId(), ano: parseInt(ano, 10), valor: parseFloat(valor), dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: currentUser?.nome || "" };
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, producaoAnualHistorico: [...c.parametrizacao?.producaoAnualHistorico || [], entrada] } } : c), setCenters);
    setModal(null);
  };
  const editarProducaoAnual = (centroId, entryId, campos) => {
    const atualizado = centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, producaoAnualHistorico: (c.parametrizacao?.producaoAnualHistorico || []).map((h) => h.id === entryId ? { ...h, ...campos } : h) } } : c);
    persist("centers", atualizado, setCenters);
    setModal({ type: "historicoProducao", data: atualizado.find((c) => c.id === centroId) });
  };
  const deleteProducaoAnual = (centroId, entryId) => {
    askConfirm("Remover este registo de produ\xE7\xE3o anual estimada?", () => {
      const atualizado = centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, producaoAnualHistorico: (c.parametrizacao?.producaoAnualHistorico || []).filter((h) => h.id !== entryId) } } : c);
      persist("centers", atualizado, setCenters);
      setModal({ type: "historicoProducao", data: atualizado.find((c) => c.id === centroId) });
      setConfirmDialog(null);
    });
  };
  const setCombustivelBloco = (centroId, blocoKey, combustivelId) => {
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, [blocoKey]: { ...c.parametrizacao?.[blocoKey] || {}, combustivelId } } } : c), setCenters);
  };
  const addTaxaBloco = (centroId, blocoKey, valor, dataEntradaVigor) => {
    const nova = { id: genId(), valor: parseFloat(valor), dataEntradaVigor, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: currentUser?.nome || "" };
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, [blocoKey]: { ...c.parametrizacao?.[blocoKey] || {}, historico: [...c.parametrizacao?.[blocoKey]?.historico || [], nova] } } } : c), setCenters);
    setModal(null);
  };
  const editarTaxaBloco = (centroId, blocoKey, entryId, campos) => {
    const atualizado = centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, [blocoKey]: { ...c.parametrizacao?.[blocoKey] || {}, historico: (c.parametrizacao?.[blocoKey]?.historico || []).map((h) => h.id === entryId ? { ...h, ...campos } : h) } } } : c);
    persist("centers", atualizado, setCenters);
    setModal((prev) => ({ type: "historicoTaxa", data: { center: atualizado.find((c) => c.id === centroId), blocoKey, titulo: prev?.data?.titulo, unidade: prev?.data?.unidade } }));
  };
  const deleteTaxaBloco = (centroId, blocoKey, entryId) => {
    askConfirm("Remover este registo de consumo?", () => {
      const atualizado = centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, [blocoKey]: { ...c.parametrizacao?.[blocoKey] || {}, historico: (c.parametrizacao?.[blocoKey]?.historico || []).filter((h) => h.id !== entryId) } } } : c);
      persist("centers", atualizado, setCenters);
      setModal((prev) => ({ type: "historicoTaxa", data: { center: atualizado.find((c) => c.id === centroId), blocoKey, titulo: prev?.data?.titulo, unidade: prev?.data?.unidade } }));
      setConfirmDialog(null);
    });
  };
  const setEnergiaTipo = (centroId, tipo) => {
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, energiaTipo: tipo } } : c), setCenters);
  };
  const updateMaoDeObraItens = (centroId, itens) => {
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, maoDeObraItens: itens } } : c), setCenters);
  };
  const updateEquipamentosItens = (centroId, itens) => {
    persist("centers", centers.map((c) => c.id === centroId ? { ...c, parametrizacao: { ...c.parametrizacao || {}, equipamentosItens: itens } } : c), setCenters);
  };
  const saveMixture = (data) => {
    if (data.id) {
      persist("mixtures", mixtures.map((m) => m.id === data.id ? { ...m, ...data } : m), setMixtures);
    } else {
      const center = centers.find((c) => c.id === data.centroId);
      const numero = center?.proximoNumero || 1;
      persist("mixtures", [...mixtures, { ...data, id: genId(), numero }], setMixtures);
      if (center) {
        persist("centers", centers.map((c) => c.id === center.id ? { ...c, proximoNumero: numero + 1 } : c), setCenters);
      }
    }
    setModal(null);
  };
  const deleteMixture = (id) => {
    askConfirm("Remover este artigo?", () => {
      persist("mixtures", mixtures.filter((m) => m.id !== id), setMixtures);
      setConfirmDialog(null);
    });
  };
  const deleteAllMixtures = (centroId) => {
    persist("mixtures", mixtures.filter((m) => m.centroId !== centroId), setMixtures);
    setModal(null);
  };
  const toggleMixtureStatus = (id) => {
    persist("mixtures", mixtures.map((m) => m.id === id ? { ...m, ativo: !(m.ativo !== false) } : m), setMixtures);
  };
  const importArticles = (centroId, rows) => {
    const novos = rows.map((r) => ({
      id: genId(),
      centroId,
      designacao: r.designacao,
      codigoManual: r.codigo,
      ativo: true,
      referenciaRelatorio: r.referenciaRelatorio || "",
      dataExecucao: r.dataExecucao || ""
    }));
    persist("mixtures", [...mixtures, ...novos], setMixtures);
    setModal(null);
  };
  const saveProveniencia = (data) => {
    if (data.id) {
      persist("proveniencias", proveniencias.map((p) => p.id === data.id ? { ...p, ...data } : p), setProveniencias);
    } else {
      persist("proveniencias", [...proveniencias, { ...data, id: genId(), ativo: true }], setProveniencias);
    }
    setModal(null);
  };
  const deleteProveniencia = (id) => {
    askConfirm("Remover esta proveni\xEAncia?", () => {
      persist("proveniencias", proveniencias.filter((p) => p.id !== id), setProveniencias);
      setConfirmDialog(null);
    });
  };
  const toggleProvenienciaStatus = (id) => {
    persist("proveniencias", proveniencias.map((p) => p.id === id ? { ...p, ativo: !(p.ativo !== false) } : p), setProveniencias);
  };
  const diariaChanged = (original, novo) => {
    const camposSimples = ["dataInicio", "dataFim", "turno", "observacoes"];
    if (camposSimples.some((k) => String(original?.[k] || "") !== String(novo?.[k] || ""))) return true;
    const normalizarLinhas = (arr) => (arr || []).map((l) => ({ artigoId: l.artigoId || "", clienteId: l.clienteId || "", centroCustoId: l.centroCustoId || "", toneladas: String(l.toneladas ?? "") }));
    if (JSON.stringify(normalizarLinhas(original?.linhas)) !== JSON.stringify(normalizarLinhas(novo?.linhas))) return true;
    const normalizarIncidencias = (arr) => (arr || []).map((i) => ({ descricao: (i.descricao || "").trim(), resolucaoData: i.resolucaoData || null, resolucaoDescricao: i.resolucaoDescricao || "" }));
    if (JSON.stringify(normalizarIncidencias(original?.incidencias)) !== JSON.stringify(normalizarIncidencias(novo?.incidencias))) return true;
    return false;
  };
  const saveDiaria = (data) => {
    if (data.id) {
      const original = diarias.find((d) => d.id === data.id);
      if (original && !diariaChanged(original, data)) {
        setModal(null);
        return;
      }
      const agora = (/* @__PURE__ */ new Date()).toISOString();
      const entrada = { id: genId(), data: agora, utilizador: currentUser?.nome || "", utilizadorId: currentUser?.id || "" };
      persist("diarias", diarias.map((d) => d.id === data.id ? { ...d, ...data, historico: [...d.historico || [], entrada] } : d), setDiarias);
    } else {
      const agora = (/* @__PURE__ */ new Date()).toISOString();
      const entrada = { id: genId(), data: agora, utilizador: currentUser?.nome || "", utilizadorId: currentUser?.id || "" };
      persist("diarias", [...diarias, { ...data, id: genId(), criadoPor: currentUser?.nome || "", historico: [entrada] }], setDiarias);
    }
    setModal(null);
  };
  const deleteHistoricoDiaria = (diariaId, entryId) => {
    askConfirm("Remover este registo do hist\xF3rico?", () => {
      setDiarias((prev) => {
        const atualizado = prev.map((d) => d.id === diariaId ? { ...d, historico: (d.historico || []).filter((h) => h.id !== entryId) } : d);
        persistRaw("diarias", atualizado);
        setModal({ type: "diaria", data: atualizado.find((d) => d.id === diariaId) });
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const editarUtilizadorHistoricoDiaria = (diariaId, entryId, novoUserId) => {
    const novoUser = users.find((u) => u.id === novoUserId);
    setDiarias((prev) => {
      const atualizado = prev.map((d) => d.id === diariaId ? { ...d, historico: (d.historico || []).map((h) => h.id === entryId ? { ...h, utilizador: novoUser?.nome || h.utilizador, utilizadorId: novoUserId } : h) } : d);
      persistRaw("diarias", atualizado);
      setModal({ type: "diaria", data: atualizado.find((d) => d.id === diariaId) });
      return atualizado;
    });
  };
  const editarDataHistoricoDiaria = (diariaId, entryId, novaDataIso) => {
    setDiarias((prev) => {
      const atualizado = prev.map((d) => d.id === diariaId ? { ...d, historico: (d.historico || []).map((h) => h.id === entryId ? { ...h, data: novaDataIso } : h) } : d);
      persistRaw("diarias", atualizado);
      setModal({ type: "diaria", data: atualizado.find((d) => d.id === diariaId) });
      return atualizado;
    });
  };
  const deleteDiaria = (id) => {
    askConfirm("Remover esta di\xE1ria de produ\xE7\xE3o?", () => {
      persist("diarias", diarias.filter((d) => d.id !== id), setDiarias);
      setConfirmDialog(null);
    });
  };
  const editarIncidenciaDiaria = (diariaId, incidentId, novaDescricao) => {
    persist("diarias", diarias.map((d) => d.id === diariaId ? { ...d, incidencias: normalizeIncidencias(d.incidencias).map((inc) => inc.id === incidentId ? { ...inc, descricao: novaDescricao } : inc) } : d), setDiarias);
    setModal(null);
  };
  const deleteIncidenciaDiaria = (diariaId, incidentId) => {
    askConfirm("Remover esta incid\xEAncia da di\xE1ria?", () => {
      persist("diarias", diarias.map((d) => d.id === diariaId ? { ...d, incidencias: normalizeIncidencias(d.incidencias).filter((inc) => inc.id !== incidentId) } : d), setDiarias);
      setConfirmDialog(null);
    });
  };
  const addAjusteStock = (ajuste) => {
    const novo = { ...ajuste, id: genId(), dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: currentUser?.nome || "" };
    persist("ajustesStock", [...ajustesStock, novo], setAjustesStock);
    setModal((prev) => ({ type: "historicoStock", data: prev?.data }));
  };
  const deleteAjusteStock = (id) => {
    askConfirm("Remover este ajuste de stock?", () => {
      persist("ajustesStock", ajustesStock.filter((a) => a.id !== id), setAjustesStock);
      setConfirmDialog(null);
    });
  };
  const saveLogotipo = (base64) => {
    persist("logotipo", base64, setLogotipoState);
    setModal(null);
  };
  const removeLogotipo = () => {
    askConfirm("Remover o log\xF3tipo?", () => {
      persist("logotipo", "", setLogotipoState);
      setConfirmDialog(null);
    });
  };
  const saveDopConfig = (data) => {
    persist("dopConfig", data, setDopConfig);
    setModal(null);
  };
  const saveDopFormula = (formulaId, dopData) => {
    const entrada = { id: genId(), ...dopData, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: currentUser?.nome || "" };
    persist("formulas", formulas.map((f) => f.id === formulaId ? { ...f, dopHistorico: [...f.dopHistorico || [], entrada] } : f), setFormulas);
  };
  const deleteDopHistorico = (formulaId, entryId) => {
    askConfirm("Remover esta vers\xE3o do hist\xF3rico da DoP?", () => {
      persist("formulas", formulas.map((f) => f.id === formulaId ? { ...f, dopHistorico: (f.dopHistorico || []).filter((h) => h.id !== entryId) } : f), setFormulas);
      setConfirmDialog(null);
    });
  };
  const importDiarias = (centroId, grupos) => {
    const novas = grupos.map((g) => ({
      id: genId(),
      centroId,
      dataInicio: g.dataInicio,
      dataFim: g.dataFim,
      linhas: g.linhas,
      criadoPor: currentUser?.nome || ""
    }));
    persist("diarias", [...diarias, ...novas], setDiarias);
    setModal(null);
  };
  const formulaChanged = (original, novo) => {
    const camposSimples = ["codigo", "estudo", "dataEstudo", "designacao", "central", "observacoes"];
    if (camposSimples.some((k) => String(original?.[k] || "") !== String(novo?.[k] || ""))) return true;
    if (JSON.stringify(original?.silos || {}) !== JSON.stringify(novo?.silos || {})) return true;
    if (JSON.stringify(original?.trabalho || {}) !== JSON.stringify(novo?.trabalho || {})) return true;
    return false;
  };
  const saveFormula = (data) => {
    if (data.id) {
      const original = formulas.find((f) => f.id === data.id);
      if (original && !formulaChanged(original, data)) {
        setModal(null);
        return;
      }
      const agora = (/* @__PURE__ */ new Date()).toISOString();
      const entrada = { id: genId(), data: agora, utilizador: currentUser?.nome || "" };
      persist("formulas", formulas.map((f) => f.id === data.id ? { ...f, ...data, dataAlteracao: agora.slice(0, 10), historico: [...f.historico || [], entrada] } : f), setFormulas);
    } else {
      const agora = (/* @__PURE__ */ new Date()).toISOString();
      const entrada = { id: genId(), data: agora, utilizador: currentUser?.nome || "" };
      persist("formulas", [...formulas, { ...data, id: genId(), dataAlteracao: agora.slice(0, 10), historico: [entrada] }], setFormulas);
    }
    setModal(null);
  };
  const duplicarFormula = (formula) => {
    const agora = (/* @__PURE__ */ new Date()).toISOString();
    const artigo = mixtures.find((m) => m.id === formula.artigoId);
    const codigoBase = artigo?.codigoManual || "";
    const irmas = formulas.filter((f) => f.artigoId === formula.artigoId);
    const novoCodigo = irmas.length === 0 ? codigoBase : `${codigoBase}.${irmas.length + 1}`;
    const novaFormula = {
      ...formula,
      id: genId(),
      codigo: novoCodigo,
      dataAlteracao: agora.slice(0, 10),
      historico: [{ id: genId(), data: agora, utilizador: currentUser?.nome || "" }],
      dop: void 0,
      dopHistorico: []
    };
    delete novaFormula.dop;
    persist("formulas", [...formulas, novaFormula], setFormulas);
    setModal({ type: "formula", data: novaFormula });
  };
  const deleteFormula = (id) => {
    askConfirm("Remover esta f\xF3rmula?", () => {
      persist("formulas", formulas.filter((f) => f.id !== id), setFormulas);
      setConfirmDialog(null);
    });
  };
  const deleteFormulas = (ids) => {
    if (!ids.length) return;
    askConfirm(`Remover ${ids.length} f\xF3rmula${ids.length > 1 ? "s" : ""}? Esta a\xE7\xE3o n\xE3o pode ser desfeita.`, () => {
      setFormulas((prev) => {
        const atualizado = prev.filter((f) => !ids.includes(f.id));
        persistRaw("formulas", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const toggleIncluirCustos = (formulaId) => {
    setFormulas((prev) => {
      const atualizado = prev.map((f) => f.id === formulaId ? { ...f, incluirEmCustosTodas: f.incluirEmCustosTodas === false ? true : false } : f);
      persistRaw("formulas", atualizado);
      return atualizado;
    });
  };
  const importFormulas = (centroId, registos) => {
    const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
    const novas = registos.map((r) => ({ ...r, id: genId(), centroId, dataAlteracao: r.dataAlteracao || hoje }));
    persist("formulas", [...formulas, ...novas], setFormulas);
    setModal(null);
  };
  const recalcularDataAlteracao = (f) => {
    const datas = (f.historico || []).map((h) => h.data).filter(Boolean);
    if (datas.length === 0) return f.dataAlteracao;
    const maisRecente = datas.reduce((max, d) => d > max ? d : max, datas[0]);
    const dt = new Date(maisRecente);
    if (isNaN(dt.getTime())) return f.dataAlteracao;
    const pad = (n) => String(n).padStart(2, "0");
    return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
  };
  const deleteHistoricoFormula = (formulaId, entryId) => {
    askConfirm("Remover este registo do hist\xF3rico de altera\xE7\xF5es?", () => {
      const atualizado = formulas.map((f) => {
        if (f.id !== formulaId) return f;
        const semEntrada = { ...f, historico: (f.historico || []).filter((h) => h.id !== entryId) };
        return { ...semEntrada, dataAlteracao: recalcularDataAlteracao(semEntrada) };
      });
      persist("formulas", atualizado, setFormulas);
      setModal({ type: "formula", data: atualizado.find((f) => f.id === formulaId) });
      setConfirmDialog(null);
    });
  };
  const editarDataHistoricoFormula = (formulaId, entryId, novaData) => {
    setFormulas((prev) => {
      const atualizado = prev.map((f) => {
        if (f.id !== formulaId) return f;
        const comDataEditada = { ...f, historico: (f.historico || []).map((h) => h.id === entryId ? { ...h, data: novaData } : h) };
        return { ...comDataEditada, dataAlteracao: recalcularDataAlteracao(comDataEditada) };
      });
      window.storage.set("formulas", JSON.stringify(atualizado), true).catch((e) => {
        setSaveError("N\xE3o foi poss\xEDvel guardar. Tente novamente.");
        console.error(e);
      });
      setModal({ type: "formula", data: atualizado.find((f) => f.id === formulaId) });
      return atualizado;
    });
  };
  const saveAvaria = (data) => {
    if (data.id) {
      persist("avarias", avarias.map((a) => a.id === data.id ? { ...a, ...data } : a), setAvarias);
    } else {
      persist("avarias", [...avarias, { ...data, id: genId(), criadoPor: currentUser?.nome || "" }], setAvarias);
    }
    setModal(null);
  };
  const deleteAvaria = (id) => {
    askConfirm("Remover esta incid\xEAncia?", () => {
      persist("avarias", avarias.filter((a) => a.id !== id), setAvarias);
      setConfirmDialog(null);
    });
  };
  const saveResolucao = (item, resolucaoData, resolucaoDescricao) => {
    if (item.origem === "Di\xE1ria") {
      persist("diarias", diarias.map((d) => {
        if (d.id !== item.diariaId) return d;
        const lista = normalizeIncidencias(d.incidencias).map((inc) => inc.id === item.incidentId ? { ...inc, resolucaoData, resolucaoDescricao } : inc);
        return { ...d, incidencias: lista };
      }), setDiarias);
    } else {
      persist("avarias", avarias.map((a) => a.id === item.id ? { ...a, resolucaoData, resolucaoDescricao } : a), setAvarias);
    }
    setModal(null);
  };
  const removerResolucao = (item) => {
    askConfirm('Reverter esta incid\xEAncia para "Por resolver"? A resolu\xE7\xE3o registada ser\xE1 apagada.', () => {
      if (item.origem === "Di\xE1ria") {
        persist("diarias", diarias.map((d) => {
          if (d.id !== item.diariaId) return d;
          const lista = normalizeIncidencias(d.incidencias).map((inc) => inc.id === item.incidentId ? { ...inc, resolucaoData: null, resolucaoDescricao: "" } : inc);
          return { ...d, incidencias: lista };
        }), setDiarias);
      } else {
        persist("avarias", avarias.map((a) => a.id === item.id ? { ...a, resolucaoData: null, resolucaoDescricao: "" } : a), setAvarias);
      }
      setConfirmDialog(null);
      setModal(null);
    });
  };
  const saveMaterial = (data) => {
    if (data.id) {
      persist("materiais", materiais.map((m) => m.id === data.id ? { ...m, designacao: data.designacao, fornecedor: data.fornecedor, centrosIds: data.centrosIds, tipoMaterialId: data.tipoMaterialId } : m), setMateriais);
    } else {
      const novoMaterial = {
        id: genId(),
        designacao: data.designacao,
        fornecedor: data.fornecedor,
        centrosIds: data.centrosIds,
        tipoMaterialId: data.tipoMaterialId,
        historicoPrecos: [{ id: genId(), preco: data.preco, descontos: data.descontos, custosExtra: data.custosExtra, dataEntradaVigor: data.dataEntradaVigor }]
      };
      persist("materiais", [...materiais, novoMaterial], setMateriais);
    }
    setModal(null);
  };
  const deleteMaterial = (id) => {
    askConfirm("Remover este material constituinte e todo o seu hist\xF3rico de pre\xE7os?", () => {
      persist("materiais", materiais.filter((m) => m.id !== id), setMateriais);
      setConfirmDialog(null);
    });
  };
  const addPrecoMaterial = (materialId, precoEntry) => {
    persist("materiais", materiais.map((m) => m.id === materialId ? { ...m, historicoPrecos: [...normalizarHistoricoPrecos(m), { id: genId(), ...precoEntry }] } : m), setMateriais);
    setModal(null);
  };
  const deletePrecoMaterial = (materialId, precoId) => {
    askConfirm("Remover esta atualiza\xE7\xE3o do hist\xF3rico de pre\xE7os?", () => {
      const atualizado = materiais.map((m) => m.id === materialId ? { ...m, historicoPrecos: (m.historicoPrecos || []).filter((p) => p.id !== precoId) } : m);
      persist("materiais", atualizado, setMateriais);
      setModal({ type: "historicoMaterial", data: atualizado.find((m) => m.id === materialId) });
      setConfirmDialog(null);
    });
  };
  const editarPrecoMaterial = (materialId, precoId, precoEntry) => {
    const atualizado = materiais.map((m) => m.id === materialId ? { ...m, historicoPrecos: normalizarHistoricoPrecos(m).map((p) => p.id === precoId ? { ...p, ...precoEntry } : p) } : m);
    persist("materiais", atualizado, setMateriais);
    setModal({ type: "historicoMaterial", data: atualizado.find((m) => m.id === materialId) });
  };
  const duplicarMaterial = (material) => {
    const novo = {
      ...material,
      id: genId(),
      designacao: `${material.designacao} (c\xF3pia)`,
      historicoPrecos: normalizarHistoricoPrecos(material).map((p) => ({ ...p, id: genId() }))
    };
    persist("materiais", [...materiais, novo], setMateriais);
    setModal({ type: "material", data: novo });
  };
  const importMateriaisPrecos = (rows) => {
    let working = materiais.map((m) => ({ ...m, historicoPrecos: normalizarHistoricoPrecos(m) }));
    const grupos = {};
    rows.forEach((r) => {
      const chave = normalizeHeader(r.designacao);
      if (!grupos[chave]) grupos[chave] = { designacaoOriginal: r.designacao, fornecedor: "", entradas: [] };
      if (r.fornecedor && !grupos[chave].fornecedor) grupos[chave].fornecedor = r.fornecedor;
      grupos[chave].entradas.push({
        id: genId(),
        preco: r.preco,
        descontos: (r.descontos || []).map((d) => ({ id: genId(), ...d })),
        custosExtra: r.custosExtra || [],
        dataEntradaVigor: r.dataEntradaVigor
      });
    });
    Object.values(grupos).forEach((grupo) => {
      grupo.entradas.sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""));
      const idx = working.findIndex((m) => normalizeHeader(m.designacao) === normalizeHeader(grupo.designacaoOriginal));
      if (idx >= 0) {
        working[idx] = { ...working[idx], historicoPrecos: [...working[idx].historicoPrecos, ...grupo.entradas] };
      } else {
        working.push({
          id: genId(),
          designacao: grupo.designacaoOriginal,
          fornecedor: grupo.fornecedor || "",
          centrosIds: "todos",
          historicoPrecos: grupo.entradas
        });
      }
    });
    persist("materiais", working, setMateriais);
    setModal(null);
  };
  const saveConsumivel = (data) => {
    if (data.id) {
      persist("consumiveis", consumiveis.map((c) => c.id === data.id ? { ...c, designacao: data.designacao, fornecedor: data.fornecedor, centrosIds: data.centrosIds, unidadeCusto: data.unidadeCusto } : c), setConsumiveis);
    } else {
      const novoConsumivel = {
        id: genId(),
        designacao: data.designacao,
        fornecedor: data.fornecedor,
        centrosIds: data.centrosIds,
        unidadeCusto: data.unidadeCusto,
        historicoPrecos: [{ id: genId(), preco: data.preco, descontos: data.descontos, custosExtra: data.custosExtra, dataEntradaVigor: data.dataEntradaVigor }]
      };
      persist("consumiveis", [...consumiveis, novoConsumivel], setConsumiveis);
    }
    setModal(null);
  };
  const deleteConsumivel = (id) => {
    askConfirm("Remover este consum\xEDvel e todo o seu hist\xF3rico de pre\xE7os?", () => {
      persist("consumiveis", consumiveis.filter((c) => c.id !== id), setConsumiveis);
      setConfirmDialog(null);
    });
  };
  const addPrecoConsumivel = (consumivelId, precoEntry) => {
    persist("consumiveis", consumiveis.map((c) => c.id === consumivelId ? { ...c, historicoPrecos: [...normalizarHistoricoPrecos(c), { id: genId(), ...precoEntry }] } : c), setConsumiveis);
    setModal(null);
  };
  const deletePrecoConsumivel = (consumivelId, precoId) => {
    askConfirm("Remover esta atualiza\xE7\xE3o do hist\xF3rico de pre\xE7os?", () => {
      persist("consumiveis", consumiveis.map((c) => c.id === consumivelId ? { ...c, historicoPrecos: (c.historicoPrecos || []).filter((p) => p.id !== precoId) } : c), setConsumiveis);
      setConfirmDialog(null);
    });
  };
  const editarPrecoConsumivel = (consumivelId, precoId, precoEntry) => {
    const atualizado = consumiveis.map((c) => c.id === consumivelId ? { ...c, historicoPrecos: normalizarHistoricoPrecos(c).map((p) => p.id === precoId ? { ...p, ...precoEntry } : p) } : c);
    persist("consumiveis", atualizado, setConsumiveis);
    setModal({ type: "historicoConsumivel", data: atualizado.find((c) => c.id === consumivelId) });
  };
  const duplicarConsumivel = (consumivel) => {
    const novo = {
      ...consumivel,
      id: genId(),
      designacao: `${consumivel.designacao} (c\xF3pia)`,
      historicoPrecos: normalizarHistoricoPrecos(consumivel).map((p) => ({ ...p, id: genId() }))
    };
    persist("consumiveis", [...consumiveis, novo], setConsumiveis);
    setModal({ type: "consumivel", data: novo });
  };
  const importConsumiveisPrecos = (rows) => {
    let working = consumiveis.map((c) => ({ ...c, historicoPrecos: normalizarHistoricoPrecos(c) }));
    const grupos = {};
    rows.forEach((r) => {
      const chave = normalizeHeader(r.designacao);
      if (!grupos[chave]) grupos[chave] = { designacaoOriginal: r.designacao, fornecedor: "", entradas: [] };
      if (r.fornecedor && !grupos[chave].fornecedor) grupos[chave].fornecedor = r.fornecedor;
      grupos[chave].entradas.push({
        id: genId(),
        preco: r.preco,
        descontos: (r.descontos || []).map((d) => ({ id: genId(), ...d })),
        custosExtra: r.custosExtra || [],
        dataEntradaVigor: r.dataEntradaVigor
      });
    });
    Object.values(grupos).forEach((grupo) => {
      grupo.entradas.sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""));
      const idx = working.findIndex((c) => normalizeHeader(c.designacao) === normalizeHeader(grupo.designacaoOriginal));
      if (idx >= 0) {
        working[idx] = { ...working[idx], historicoPrecos: [...working[idx].historicoPrecos, ...grupo.entradas] };
      } else {
        working.push({
          id: genId(),
          designacao: grupo.designacaoOriginal,
          fornecedor: grupo.fornecedor || "",
          centrosIds: "todos",
          historicoPrecos: grupo.entradas
        });
      }
    });
    persist("consumiveis", working, setConsumiveis);
    setModal(null);
  };
  const saveEquipamento = (data) => {
    if (data.id) {
      persist("equipamentos", equipamentos.map((e) => e.id === data.id ? { ...e, designacao: data.designacao, fornecedor: data.fornecedor, centrosIds: data.centrosIds, unidadeCusto: data.unidadeCusto } : e), setEquipamentos);
    } else {
      const novoEquipamento = {
        id: genId(),
        designacao: data.designacao,
        fornecedor: data.fornecedor,
        centrosIds: data.centrosIds,
        unidadeCusto: data.unidadeCusto,
        historicoPrecos: [{ id: genId(), preco: data.preco, descontos: [], custoTransporte: 0, dataEntradaVigor: data.dataEntradaVigor }]
      };
      persist("equipamentos", [...equipamentos, novoEquipamento], setEquipamentos);
    }
    setModal(null);
  };
  const deleteEquipamento = (id) => {
    askConfirm("Remover este equipamento e todo o seu hist\xF3rico de pre\xE7os?", () => {
      persist("equipamentos", equipamentos.filter((e) => e.id !== id), setEquipamentos);
      setConfirmDialog(null);
    });
  };
  const addPrecoEquipamento = (equipamentoId, precoEntry) => {
    persist("equipamentos", equipamentos.map((e) => e.id === equipamentoId ? { ...e, historicoPrecos: [...normalizarHistoricoPrecos(e), { id: genId(), ...precoEntry, descontos: [], custoTransporte: 0 }] } : e), setEquipamentos);
    setModal(null);
  };
  const deletePrecoEquipamento = (equipamentoId, precoId) => {
    askConfirm("Remover esta atualiza\xE7\xE3o do hist\xF3rico de pre\xE7os?", () => {
      const atualizado = equipamentos.map((e) => e.id === equipamentoId ? { ...e, historicoPrecos: (e.historicoPrecos || []).filter((p) => p.id !== precoId) } : e);
      persist("equipamentos", atualizado, setEquipamentos);
      setModal({ type: "historicoEquipamento", data: atualizado.find((e) => e.id === equipamentoId) });
      setConfirmDialog(null);
    });
  };
  const editarPrecoEquipamento = (equipamentoId, precoId, precoEntry) => {
    const atualizado = equipamentos.map((e) => e.id === equipamentoId ? { ...e, historicoPrecos: normalizarHistoricoPrecos(e).map((p) => p.id === precoId ? { ...p, ...precoEntry, descontos: [], custoTransporte: 0 } : p) } : e);
    persist("equipamentos", atualizado, setEquipamentos);
    setModal({ type: "historicoEquipamento", data: atualizado.find((e) => e.id === equipamentoId) });
  };
  const duplicarEquipamento = (equipamento) => {
    const novo = {
      ...equipamento,
      id: genId(),
      designacao: `${equipamento.designacao} (c\xF3pia)`,
      historicoPrecos: normalizarHistoricoPrecos(equipamento).map((p) => ({ ...p, id: genId() }))
    };
    persist("equipamentos", [...equipamentos, novo], setEquipamentos);
    setModal({ type: "equipamento", data: novo });
  };
  const importEquipamentosPrecos = (rows) => {
    let working = equipamentos.map((e) => ({ ...e, historicoPrecos: normalizarHistoricoPrecos(e) }));
    const grupos = {};
    rows.forEach((r) => {
      const chave = normalizeHeader(r.designacao);
      if (!grupos[chave]) grupos[chave] = { designacaoOriginal: r.designacao, fornecedor: "", entradas: [] };
      if (r.fornecedor && !grupos[chave].fornecedor) grupos[chave].fornecedor = r.fornecedor;
      grupos[chave].entradas.push({ id: genId(), preco: r.preco, descontos: [], custoTransporte: 0, dataEntradaVigor: r.dataEntradaVigor });
    });
    Object.values(grupos).forEach((grupo) => {
      grupo.entradas.sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""));
      const idx = working.findIndex((e) => normalizeHeader(e.designacao) === normalizeHeader(grupo.designacaoOriginal));
      if (idx >= 0) {
        working[idx] = { ...working[idx], historicoPrecos: [...working[idx].historicoPrecos, ...grupo.entradas] };
      } else {
        working.push({
          id: genId(),
          designacao: grupo.designacaoOriginal,
          fornecedor: grupo.fornecedor || "",
          centrosIds: "todos",
          historicoPrecos: grupo.entradas
        });
      }
    });
    persist("equipamentos", working, setEquipamentos);
    setModal(null);
  };
  const saveMaoDeObra = (data) => {
    if (data.id) {
      persist("maoDeObra", maoDeObra.map((e) => e.id === data.id ? { ...e, designacao: data.designacao, fornecedor: data.fornecedor, centrosIds: data.centrosIds, unidadeCusto: data.unidadeCusto } : e), setMaoDeObra);
    } else {
      const novo = {
        id: genId(),
        designacao: data.designacao,
        fornecedor: data.fornecedor,
        centrosIds: data.centrosIds,
        unidadeCusto: data.unidadeCusto,
        historicoPrecos: [{ id: genId(), preco: data.preco, descontos: [], custoTransporte: 0, dataEntradaVigor: data.dataEntradaVigor }]
      };
      persist("maoDeObra", [...maoDeObra, novo], setMaoDeObra);
    }
    setModal(null);
  };
  const deleteMaoDeObra = (id) => {
    askConfirm("Remover esta categoria de m\xE3o de obra e todo o seu hist\xF3rico de pre\xE7os?", () => {
      persist("maoDeObra", maoDeObra.filter((e) => e.id !== id), setMaoDeObra);
      setConfirmDialog(null);
    });
  };
  const addPrecoMaoObra = (id, precoEntry) => {
    persist("maoDeObra", maoDeObra.map((e) => e.id === id ? { ...e, historicoPrecos: [...normalizarHistoricoPrecos(e), { id: genId(), ...precoEntry, descontos: [], custoTransporte: 0 }] } : e), setMaoDeObra);
    setModal(null);
  };
  const deletePrecoMaoObra = (id, precoId) => {
    askConfirm("Remover esta atualiza\xE7\xE3o do hist\xF3rico de pre\xE7os?", () => {
      const atualizado = maoDeObra.map((e) => e.id === id ? { ...e, historicoPrecos: (e.historicoPrecos || []).filter((p) => p.id !== precoId) } : e);
      persist("maoDeObra", atualizado, setMaoDeObra);
      setModal({ type: "historicoMaoObra", data: atualizado.find((e) => e.id === id) });
      setConfirmDialog(null);
    });
  };
  const editarPrecoMaoObra = (id, precoId, precoEntry) => {
    const atualizado = maoDeObra.map((e) => e.id === id ? { ...e, historicoPrecos: normalizarHistoricoPrecos(e).map((p) => p.id === precoId ? { ...p, ...precoEntry, descontos: [], custoTransporte: 0 } : p) } : e);
    persist("maoDeObra", atualizado, setMaoDeObra);
    setModal({ type: "historicoMaoObra", data: atualizado.find((e) => e.id === id) });
  };
  const duplicarMaoDeObra = (item) => {
    const novo = {
      ...item,
      id: genId(),
      designacao: `${item.designacao} (c\xF3pia)`,
      historicoPrecos: normalizarHistoricoPrecos(item).map((p) => ({ ...p, id: genId() }))
    };
    persist("maoDeObra", [...maoDeObra, novo], setMaoDeObra);
    setModal({ type: "maodeobra", data: novo });
  };
  const importMaoDeObraPrecos = (rows) => {
    let working = maoDeObra.map((e) => ({ ...e, historicoPrecos: normalizarHistoricoPrecos(e) }));
    const grupos = {};
    rows.forEach((r) => {
      const chave = normalizeHeader(r.designacao);
      if (!grupos[chave]) grupos[chave] = { designacaoOriginal: r.designacao, fornecedor: "", entradas: [] };
      if (r.fornecedor && !grupos[chave].fornecedor) grupos[chave].fornecedor = r.fornecedor;
      grupos[chave].entradas.push({ id: genId(), preco: r.preco, descontos: [], custoTransporte: 0, dataEntradaVigor: r.dataEntradaVigor });
    });
    Object.values(grupos).forEach((grupo) => {
      grupo.entradas.sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""));
      const idx = working.findIndex((e) => normalizeHeader(e.designacao) === normalizeHeader(grupo.designacaoOriginal));
      if (idx >= 0) {
        working[idx] = { ...working[idx], historicoPrecos: [...working[idx].historicoPrecos, ...grupo.entradas] };
      } else {
        working.push({
          id: genId(),
          designacao: grupo.designacaoOriginal,
          fornecedor: grupo.fornecedor || "",
          centrosIds: "todos",
          historicoPrecos: grupo.entradas
        });
      }
    });
    persist("maoDeObra", working, setMaoDeObra);
    setModal(null);
  };
  const saveRececao = (data) => {
    if (data.id) {
      persist("rececoes", rececoes.map((r) => r.id === data.id ? { ...r, ...data } : r), setRececoes);
    } else {
      persist("rececoes", [...rececoes, { ...data, id: genId() }], setRececoes);
    }
    setModal(null);
  };
  const deleteRececao = (id) => {
    askConfirm("Remover este registo de rece\xE7\xE3o?", () => {
      persist("rececoes", rececoes.filter((r) => r.id !== id), setRececoes);
      setConfirmDialog(null);
    });
  };
  const saveCliente = (data) => {
    if (data.id) {
      persist("clientes", clientes.map((c) => c.id === data.id ? { ...c, ...data } : c), setClientes);
    } else {
      persist("clientes", [...clientes, { ...data, id: genId() }], setClientes);
    }
    setModal(null);
  };
  const deleteCliente = (id) => {
    const target = clientes.find((c) => c.id === id);
    askConfirm(`Remover o cliente "${target?.designacao}" e todos os seus centros de custo?`, () => {
      persist("clientes", clientes.filter((c) => c.id !== id), setClientes);
      persist("centrosCusto", centrosCusto.filter((cc) => cc.clienteId !== id), setCentrosCusto);
      setView("clientes");
      setConfirmDialog(null);
    });
  };
  const importClientes = (rows) => {
    const novos = rows.map((r) => ({ id: genId(), numero: r.numero, nif: r.nif, designacao: r.designacao, morada: r.morada || "" }));
    persist("clientes", [...clientes, ...novos], setClientes);
    setModal(null);
  };
  const saveFornecedor = (data) => {
    if (data.id) {
      persist("fornecedores", fornecedores.map((f) => f.id === data.id ? { ...f, ...data } : f), setFornecedores);
    } else {
      persist("fornecedores", [...fornecedores, { ...data, id: genId() }], setFornecedores);
    }
    setModal(null);
  };
  const deleteFornecedor = (id) => {
    const target = fornecedores.find((f) => f.id === id);
    askConfirm(`Remover o fornecedor "${target?.designacao}"?`, () => {
      persist("fornecedores", fornecedores.filter((f) => f.id !== id), setFornecedores);
      setConfirmDialog(null);
    });
  };
  const importFornecedores = (rows) => {
    const novos = rows.map((r) => ({ id: genId(), numero: r.numero, nif: r.nif, designacao: r.designacao, morada: r.morada || "" }));
    persist("fornecedores", [...fornecedores, ...novos], setFornecedores);
    setModal(null);
  };
  const saveCentroCusto = (data) => {
    if (data.id) {
      persist("centrosCusto", centrosCusto.map((cc) => cc.id === data.id ? { ...cc, ...data } : cc), setCentrosCusto);
    } else {
      persist("centrosCusto", [...centrosCusto, { ...data, id: genId(), ativo: true }], setCentrosCusto);
    }
    setModal(null);
  };
  const persistRaw = (key, value) => {
    const anterior = writeQueueRef.current[key] || Promise.resolve();
    const atual = anterior.catch(() => {
    }).then(() => window.storage.set(key, JSON.stringify(value), true)).catch((e) => {
      setSaveError("N\xE3o foi poss\xEDvel guardar. Tente novamente.");
      console.error(e);
    });
    writeQueueRef.current[key] = atual;
    return atual;
  };
  const deleteCentroCusto = (id) => {
    askConfirm("Remover este centro de custo?", () => {
      setCentrosCusto((prev) => {
        const atualizado = prev.filter((cc) => cc.id !== id);
        persistRaw("centrosCusto", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const deleteAllCentrosCusto = (ids) => {
    if (!ids.length) return;
    askConfirm(`Remover ${ids.length} obra${ids.length > 1 ? "s" : ""}? Esta a\xE7\xE3o n\xE3o pode ser desfeita.`, () => {
      setCentrosCusto((prev) => {
        const atualizado = prev.filter((cc) => !ids.includes(cc.id));
        persistRaw("centrosCusto", atualizado);
        return atualizado;
      });
      setConfirmDialog(null);
    });
  };
  const toggleCentroCustoStatus = (id) => {
    persist("centrosCusto", centrosCusto.map((cc) => cc.id === id ? { ...cc, ativo: !(cc.ativo !== false) } : cc), setCentrosCusto);
  };
  const importCentrosCusto = (clienteId, rows) => {
    const novos = rows.map((r) => ({
      id: genId(),
      clienteId,
      codigo: r.codigo || "",
      designacao: r.designacao,
      codigoPostal: r.codigoPostal || "",
      localidade: r.localidade || "",
      ativo: true
    }));
    setCentrosCusto((prev) => {
      const atualizado = [...prev, ...novos];
      persistRaw("centrosCusto", atualizado);
      return atualizado;
    });
    setModal(null);
  };
  const fontStyle = /* @__PURE__ */ jsx("style", { children: `
      @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');
      .font-display { font-family: 'Oswald', sans-serif; }
      .font-mono-data { font-family: 'JetBrains Mono', monospace; }
    ` });
  if (loading) {
    return /* @__PURE__ */ jsxs("div", { className: "min-h-screen bg-stone-950 flex items-center justify-center", children: [
      fontStyle,
      /* @__PURE__ */ jsx("div", { className: "text-amber-500 font-display tracking-widest text-sm animate-pulse", children: "A CARREGAR..." })
    ] });
  }
  if (users.length === 0) {
    return /* @__PURE__ */ jsxs("div", { className: "min-h-screen bg-stone-950 flex items-center justify-center p-4", children: [
      fontStyle,
      /* @__PURE__ */ jsxs("div", { className: "w-full max-w-sm", children: [
        /* @__PURE__ */ jsxs("div", { className: "text-center mb-6", children: [
          /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.3em] text-amber-500 uppercase mb-2", children: "Misturas Betuminosas" }),
          /* @__PURE__ */ jsx("h1", { className: "font-display text-2xl text-white font-semibold", children: "Configura\xE7\xE3o Inicial" }),
          /* @__PURE__ */ jsx("p", { className: "text-stone-400 text-sm mt-2", children: "Crie a conta de Administrador para come\xE7ar" })
        ] }),
        /* @__PURE__ */ jsx(SetupForm, { onSubmit: createFirstAdmin })
      ] })
    ] });
  }
  if (!currentUser) {
    return /* @__PURE__ */ jsxs("div", { className: "min-h-screen bg-white p-4 md:p-8 flex items-center justify-center", children: [
      fontStyle,
      /* @__PURE__ */ jsxs("div", { className: "w-full max-w-sm", children: [
        /* @__PURE__ */ jsxs("div", { className: "text-center mb-8", children: [
          logotipo && /* @__PURE__ */ jsx("img", { src: logotipo, alt: "Log\xF3tipo", className: "max-h-16 mx-auto mb-4" }),
          /* @__PURE__ */ jsx("p", { className: "font-display text-2xl tracking-[0.15em] text-amber-600 uppercase mb-2 font-semibold", children: "Misturas Betuminosas" }),
          /* @__PURE__ */ jsx("h1", { className: "font-display text-lg text-stone-800 font-medium uppercase tracking-widest", children: "Identifica\xE7\xE3o" })
        ] }),
        /* @__PURE__ */ jsx(LoginForm, { users, onLogin: attemptLogin, onDefinePassword: definePin })
      ] })
    ] });
  }
  const needsCenterSelect = currentUser.role === "Operador" && !activeCenterId;
  if (needsCenterSelect) {
    const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
    const escalasHoje = (currentUser.escalasCentros || []).filter((e) => e.dataInicio <= hoje && (!e.dataFim || e.dataFim >= hoje));
    const centrosPermitidosIds = new Set(escalasHoje.map((e) => e.centroId));
    const activeCenters = centers.filter((c) => c.ativo !== false && centrosPermitidosIds.has(c.id));
    return /* @__PURE__ */ jsxs("div", { className: "min-h-screen bg-stone-950 p-4 md:p-8", children: [
      fontStyle,
      /* @__PURE__ */ jsxs("div", { className: "max-w-2xl mx-auto pt-8", children: [
        /* @__PURE__ */ jsxs("div", { className: "text-center mb-8", children: [
          /* @__PURE__ */ jsxs("p", { className: "font-display text-xs tracking-[0.3em] text-amber-500 uppercase mb-2", children: [
            "Ol\xE1, ",
            currentUser.nome
          ] }),
          /* @__PURE__ */ jsx("h1", { className: "font-display text-3xl text-white font-semibold", children: "Onde est\xE1 a trabalhar hoje?" }),
          /* @__PURE__ */ jsx("p", { className: "text-stone-400 text-sm mt-2", children: "Escolha o centro de produ\xE7\xE3o para esta sess\xE3o \u2014 pode trocar a qualquer momento" })
        ] }),
        activeCenters.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-stone-900 border border-dashed border-stone-700 rounded-xl p-10 text-center", children: [
          /* @__PURE__ */ jsx(Factory, { className: "mx-auto text-stone-600 mb-3", size: 30 }),
          /* @__PURE__ */ jsx("p", { className: "text-stone-400 text-sm", children: "N\xE3o tem nenhum centro atribu\xEDdo para hoje. Pe\xE7a a um Administrador para o adicionar ao calend\xE1rio de acesso." })
        ] }) : /* @__PURE__ */ jsx("div", { className: "grid sm:grid-cols-2 gap-3", children: activeCenters.map((c) => /* @__PURE__ */ jsx(
          "button",
          {
            onClick: () => setActiveCenterId(c.id),
            className: "bg-stone-900 border border-stone-800 rounded-xl p-4 text-left hover:border-amber-600 transition-colors group",
            children: /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3", children: [
              /* @__PURE__ */ jsx("div", { className: "w-9 h-9 rounded-lg bg-stone-800 flex items-center justify-center group-hover:bg-amber-600/20", children: /* @__PURE__ */ jsx(Factory, { className: "text-amber-500", size: 16 }) }),
              /* @__PURE__ */ jsxs("div", { children: [
                /* @__PURE__ */ jsx("span", { className: "font-display text-white font-medium block group-hover:text-amber-500", children: c.nome }),
                c.localizacao && /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-500", children: c.localizacao })
              ] })
            ] })
          },
          c.id
        )) }),
        /* @__PURE__ */ jsxs("button", { onClick: logout, className: "w-full mt-6 flex items-center justify-center gap-2 py-2 text-sm text-stone-500 hover:text-stone-300", children: [
          /* @__PURE__ */ jsx(LogOut, { size: 14 }),
          " Trocar de utilizador"
        ] })
      ] })
    ] });
  }
  const selectedCenter = centers.find((c) => c.id === selectedCenterId);
  const workingCenter = centers.find((c) => c.id === activeCenterId);
  const selectedCliente = clientes.find((c) => c.id === selectedClienteId);
  const clienteSocorpena = clientes.find(isSocorpenaCliente);
  const centrosCustoSocorpena = centrosCusto.filter((cc) => clientes.some((c) => c.id === cc.clienteId && isSocorpenaCliente(c)));
  return /* @__PURE__ */ jsxs("div", { className: "min-h-screen bg-stone-100 flex", children: [
    fontStyle,
    /* @__PURE__ */ jsxs("aside", { className: "w-60 bg-stone-900 flex flex-col shrink-0", children: [
      /* @__PURE__ */ jsxs("div", { className: "px-5 py-6 bg-white", children: [
        logotipo && /* @__PURE__ */ jsx("img", { src: logotipo, alt: "Log\xF3tipo", className: "max-h-12 mb-3" }),
        /* @__PURE__ */ jsx("p", { className: "font-display text-[10px] tracking-[0.25em] text-amber-600 uppercase mb-1", children: "Misturas Betuminosas" }),
        /* @__PURE__ */ jsx("h1", { className: "font-display text-lg text-stone-900 font-semibold leading-tight", children: "Gest\xE3o de Centrais" })
      ] }),
      /* @__PURE__ */ jsxs("nav", { className: "flex-1 px-3 py-4 space-y-1", children: [
        canManageArticles && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("users"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "users" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Users, { size: 17 }),
              " Utilizadores"
            ]
          }
        ),
        podeVerCentros && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => {
              setView("centers");
              setSelectedCenterId(null);
            },
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "centers" || view === "centerDetail" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Factory, { size: 17 }),
              " Centros de Produ\xE7\xE3o"
            ]
          }
        ),
        podeVerClientesFornecedores && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => {
              setView("clientes");
              setSelectedClienteId(null);
            },
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "clientes" || view === "clienteDetail" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Building2, { size: 17 }),
              " Clientes"
            ]
          }
        ),
        /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("obras"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "obras" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(MapPin, { size: 17 }),
              " Listagem Obras SCRPN"
            ]
          }
        ),
        podeVerClientesFornecedores && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("fornecedores"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "fornecedores" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Truck, { size: 17 }),
              " Fornecedores"
            ]
          }
        ),
        podeVerMateriais && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("materiais"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "materiais" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Package, { size: 17 }),
              " Materiais Constituintes"
            ]
          }
        ),
        podeVerMateriais && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("consumiveis"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "consumiveis" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Fuel, { size: 17 }),
              " Combust\xEDveis/Energia"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("equipamentos"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "equipamentos" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(Wrench, { size: 17 }),
              " Equipamentos"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setView("maodeobra"),
            className: `w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${view === "maodeobra" ? "bg-amber-600 text-white" : "text-stone-300 hover:bg-stone-800"}`,
            children: [
              /* @__PURE__ */ jsx(HardHat, { size: 17 }),
              " M\xE3o de Obra"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setModal({ type: "dopConfig" }),
            className: "w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors text-stone-300 hover:bg-stone-800",
            children: [
              /* @__PURE__ */ jsx(ShieldCheck, { size: 17 }),
              " Defini\xE7\xF5es DoP"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setModal({ type: "logotipo" }),
            className: "w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors text-stone-300 hover:bg-stone-800",
            children: [
              /* @__PURE__ */ jsx(Image, { size: 17 }),
              " Log\xF3tipo da Empresa"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => setModal({ type: "backup" }),
            className: "w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors text-stone-300 hover:bg-stone-800",
            children: [
              /* @__PURE__ */ jsx(Archive, { size: 17 }),
              " Backup"
            ]
          }
        )
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "p-3 bg-white", children: [
        currentUser.role === "Operador" && workingCenter && /* @__PURE__ */ jsxs("button", { onClick: () => setActiveCenterId(null), className: "w-full text-left bg-stone-100 rounded-lg p-3 mb-2 hover:bg-stone-200", children: [
          /* @__PURE__ */ jsx("p", { className: "text-[10px] uppercase tracking-wide text-stone-500 font-semibold mb-0.5", children: "A trabalhar em" }),
          /* @__PURE__ */ jsxs("p", { className: "font-display text-sm text-amber-600 font-medium flex items-center gap-1.5", children: [
            /* @__PURE__ */ jsx(Factory, { size: 13 }),
            " ",
            workingCenter.nome
          ] }),
          /* @__PURE__ */ jsx("p", { className: "text-[10px] text-stone-500 mt-0.5", children: "Trocar de centro" })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg p-3 mb-2", children: [
          /* @__PURE__ */ jsx("p", { className: "font-display text-sm text-stone-900 font-medium", children: currentUser.nome }),
          /* @__PURE__ */ jsx("div", { className: "mt-1.5", children: /* @__PURE__ */ jsx(RoleBadge, { role: currentUser.role }) })
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: logout, className: "w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-stone-500 hover:text-stone-900 hover:bg-stone-100", children: [
          /* @__PURE__ */ jsx(LogOut, { size: 15 }),
          " Sair"
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("main", { className: "flex-1 overflow-y-auto", children: [
      saveError && /* @__PURE__ */ jsx("div", { className: "bg-red-50 border-b border-red-200 text-red-700 text-sm px-6 py-2", children: saveError }),
      view === "users" && canManageArticles && /* @__PURE__ */ jsx(
        UsersView,
        {
          users,
          isAdmin,
          currentUserRole: currentUser?.role,
          perfisPersonalizados,
          onAdd: () => setModal({ type: "user" }),
          onEdit: (u) => setModal({ type: "user", data: u }),
          onDelete: deleteUser,
          onBack: () => setView("centers"),
          onSavePerfil: saveNovoPerfil,
          onDeletePerfil: deletePerfilPersonalizado
        }
      ),
      view === "centers" && /* @__PURE__ */ jsx(
        CentersView,
        {
          centers: currentUser.role === "Operador" ? centers.filter((c) => {
            const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
            const escalasHoje = (currentUser.escalasCentros || []).filter((e) => e.dataInicio <= hoje && (!e.dataFim || e.dataFim >= hoje));
            return escalasHoje.some((e) => e.centroId === c.id);
          }) : centers,
          mixtures,
          diarias,
          podeEditar: podeEditarCentros,
          onAdd: () => setModal({ type: "center" }),
          onOpen: (id) => {
            setSelectedCenterId(id);
            setView("centerDetail");
          },
          onImport: importCenters
        }
      ),
      view === "centerDetail" && selectedCenter && (currentUser.role !== "Operador" || (() => {
        const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
        const escalasHoje = (currentUser.escalasCentros || []).filter((e) => e.dataInicio <= hoje && (!e.dataFim || e.dataFim >= hoje));
        return escalasHoje.some((e) => e.centroId === selectedCenter.id);
      })()) && /* @__PURE__ */ jsx(
        CenterDetail,
        {
          center: selectedCenter,
          mixtures: mixtures.filter((m) => m.centroId === selectedCenter.id),
          proveniencias: proveniencias.filter((p) => p.centroId === selectedCenter.id),
          diarias: diarias.filter((d) => d.centroId === selectedCenter.id),
          formulas: formulas.filter((f) => f.centroId === selectedCenter.id),
          avarias: avarias.filter((a) => a.centroId === selectedCenter.id),
          clientes,
          centrosCusto,
          canManage: canManageArticles,
          podeRegistar,
          podeVerCustos,
          isAdmin,
          currentUserRole: currentUser?.role,
          onBack: () => {
            setView("centers");
            setSelectedCenterId(null);
          },
          onEditCenter: () => setModal({ type: "center", data: selectedCenter }),
          onDeleteCenter: () => deleteCenter(selectedCenter.id),
          onToggleStatus: () => toggleCenterStatus(selectedCenter.id),
          onAddMixture: () => setModal({ type: "mixture", data: { centroId: selectedCenter.id } }),
          onEditMixture: (m) => setModal({ type: "mixture", data: m }),
          onDeleteMixture: deleteMixture,
          onToggleMixtureStatus: toggleMixtureStatus,
          onOpenEliminarTodosArtigos: () => setModal({ type: "eliminarTodosArtigos", data: { centroId: selectedCenter.id, count: mixtures.filter((m) => m.centroId === selectedCenter.id).length } }),
          onImport: () => setModal({ type: "import", data: { centroId: selectedCenter.id } }),
          onAddProveniencia: () => setModal({ type: "proveniencia", data: { centroId: selectedCenter.id } }),
          onEditProveniencia: (p) => setModal({ type: "proveniencia", data: p }),
          onDeleteProveniencia: deleteProveniencia,
          onToggleProvenienciaStatus: toggleProvenienciaStatus,
          onAddDiaria: () => setModal({ type: "diaria", data: { centroId: selectedCenter.id } }),
          onEditDiaria: (d) => setModal({ type: "diaria", data: d }),
          onDeleteDiaria: deleteDiaria,
          onImportDiarias: () => setModal({ type: "importDiarias", data: { centroId: selectedCenter.id } }),
          onAddFormula: () => setModal({ type: "formula", data: { centroId: selectedCenter.id } }),
          onEditFormula: (f) => setModal({ type: "formula", data: f }),
          onDeleteFormula: deleteFormula,
          onImportFormulas: () => setModal({ type: "importFormulas", data: { centroId: selectedCenter.id } }),
          onDuplicateFormula: duplicarFormula,
          onDeleteAllFormulas: (centroId) => deleteFormulas(formulas.filter((f) => f.centroId === centroId).map((f) => f.id)),
          onDeleteSelectedFormulas: deleteFormulas,
          onToggleIncluirCustos: toggleIncluirCustos,
          onAddAvaria: () => setModal({ type: "avaria", data: { centroId: selectedCenter.id } }),
          onEditAvaria: (a) => setModal({ type: "avaria", data: a }),
          onDeleteAvaria: deleteAvaria,
          onOpenDiaria: (d) => setModal({ type: "diaria", data: d }),
          onOpenResolucao: (item) => setModal({ type: "resolucao", data: item }),
          onEditIncidenciaDiaria: (item) => setModal({ type: "editarIncidenciaDiaria", data: item }),
          onDeleteIncidenciaDiaria: deleteIncidenciaDiaria,
          ajustesStock,
          onOpenHistoricoStock: (data) => setModal({ type: "historicoStock", data }),
          onOpenAtualizarProducao: () => setModal({ type: "atualizarProducao", data: selectedCenter }),
          onOpenHistoricoProducao: () => setModal({ type: "historicoProducao", data: selectedCenter }),
          consumiveis,
          equipamentos,
          maoDeObra,
          onUpdateMaoDeObraItens: updateMaoDeObraItens,
          onUpdateEquipamentosItens: updateEquipamentosItens,
          nomeUtilizadorAtual: currentUser?.nome,
          logotipo,
          onSetCombustivel: setCombustivelBloco,
          onOpenAtualizarTaxa: (blocoKey, titulo, unidade) => setModal({ type: "atualizarTaxa", data: { center: selectedCenter, blocoKey, titulo, unidade } }),
          onOpenHistoricoTaxa: (blocoKey, titulo, unidade) => setModal({ type: "historicoTaxa", data: { center: selectedCenter, blocoKey, titulo, unidade } }),
          onSetEnergiaTipo: setEnergiaTipo,
          materiais,
          tiposMaterial,
          rececoes: rececoes.filter((r) => r.centroId === selectedCenter.id),
          onAddRececao: (categoria) => setModal({ type: "rececao", data: { centroId: selectedCenter.id, categoria } }),
          onEditRececao: (r) => setModal({ type: "rececao", data: r }),
          onDeleteRececao: deleteRececao
        }
      ),
      view === "clientes" && podeVerClientesFornecedores && /* @__PURE__ */ jsx(
        ClientesView,
        {
          clientes,
          centrosCusto,
          canManage: canManageArticles,
          onAdd: () => setModal({ type: "cliente" }),
          onImport: () => setModal({ type: "importClientes" }),
          onOpen: (id) => {
            setSelectedClienteId(id);
            setView("clienteDetail");
          },
          onDelete: deleteCliente,
          onBack: () => setView("centers")
        }
      ),
      view === "fornecedores" && podeVerClientesFornecedores && /* @__PURE__ */ jsx(
        FornecedoresView,
        {
          fornecedores,
          canManage: canManageArticles,
          podeCriarFornecedores,
          onAdd: () => setModal({ type: "fornecedor" }),
          onEdit: (f) => setModal({ type: "fornecedor", data: f }),
          onImport: () => setModal({ type: "importFornecedores" }),
          onDelete: deleteFornecedor,
          onBack: () => setView("centers")
        }
      ),
      view === "materiais" && podeVerMateriais && /* @__PURE__ */ jsx(
        MateriaisView,
        {
          materiais,
          centers,
          tiposMaterial,
          tiposDesconto,
          tiposCustoExtra,
          isAdmin,
          onAdd: () => setModal({ type: "material" }),
          onEdit: (m) => setModal({ type: "material", data: m }),
          onDelete: deleteMaterial,
          onAddPreco: (m) => setModal({ type: "precoMaterial", data: m }),
          onViewHistorico: (m) => setModal({ type: "historicoMaterial", data: m }),
          onUpdatePricePicker: () => setModal({ type: "escolherMaterial" }),
          onImportPrecos: () => setModal({ type: "importMateriais" }),
          onDuplicate: duplicarMaterial,
          onBack: () => setView("centers"),
          onSaveTipoMaterial: saveTipoMaterial,
          onDeleteTipoMaterial: deleteTipoMaterial,
          onSaveTipoDesconto: saveTipoDesconto,
          onDeleteTipoDesconto: deleteTipoDesconto,
          onSaveTipoCustoExtra: saveTipoCustoExtra,
          onDeleteTipoCustoExtra: deleteTipoCustoExtra
        }
      ),
      view === "consumiveis" && podeVerMateriais && /* @__PURE__ */ jsx(
        MateriaisView,
        {
          tipo: "consumivel",
          materiais: consumiveis,
          centers,
          tiposDesconto,
          tiposCustoExtra,
          isAdmin,
          onAdd: () => setModal({ type: "consumivel" }),
          onEdit: (m) => setModal({ type: "consumivel", data: m }),
          onDelete: deleteConsumivel,
          onAddPreco: (m) => setModal({ type: "precoConsumivel", data: m }),
          onViewHistorico: (m) => setModal({ type: "historicoConsumivel", data: m }),
          onUpdatePricePicker: () => setModal({ type: "escolherConsumivel" }),
          onImportPrecos: () => setModal({ type: "importConsumiveis" }),
          onDuplicate: duplicarConsumivel,
          onBack: () => setView("centers"),
          onSaveTipoDesconto: saveTipoDesconto,
          onDeleteTipoDesconto: deleteTipoDesconto,
          onSaveTipoCustoExtra: saveTipoCustoExtra,
          onDeleteTipoCustoExtra: deleteTipoCustoExtra
        }
      ),
      view === "equipamentos" && isAdmin && /* @__PURE__ */ jsx(
        MateriaisView,
        {
          tipo: "equipamento",
          materiais: equipamentos,
          centers,
          isAdmin,
          onAdd: () => setModal({ type: "equipamento" }),
          onEdit: (m) => setModal({ type: "equipamento", data: m }),
          onDelete: deleteEquipamento,
          onAddPreco: (m) => setModal({ type: "precoEquipamento", data: m }),
          onViewHistorico: (m) => setModal({ type: "historicoEquipamento", data: m }),
          onUpdatePricePicker: () => setModal({ type: "escolherEquipamento" }),
          onImportPrecos: () => setModal({ type: "importEquipamentos" }),
          onDuplicate: duplicarEquipamento,
          onBack: () => setView("centers")
        }
      ),
      view === "maodeobra" && isAdmin && /* @__PURE__ */ jsx(
        MateriaisView,
        {
          tipo: "maodeobra",
          materiais: maoDeObra,
          centers,
          isAdmin,
          onAdd: () => setModal({ type: "maodeobra" }),
          onEdit: (m) => setModal({ type: "maodeobra", data: m }),
          onDelete: deleteMaoDeObra,
          onAddPreco: (m) => setModal({ type: "precoMaoObra", data: m }),
          onViewHistorico: (m) => setModal({ type: "historicoMaoObra", data: m }),
          onUpdatePricePicker: () => setModal({ type: "escolherMaoObra" }),
          onImportPrecos: () => setModal({ type: "importMaoObra" }),
          onDuplicate: duplicarMaoDeObra,
          onBack: () => setView("centers")
        }
      ),
      view === "obras" && /* @__PURE__ */ jsx(
        ObrasView,
        {
          cliente: clienteSocorpena,
          centrosCusto: centrosCustoSocorpena,
          isAdmin: podeGerirObras,
          onAdd: () => setModal({ type: "centroCusto", data: { clienteId: clienteSocorpena.id } }),
          onEdit: (cc) => setModal({ type: "centroCusto", data: cc }),
          onDelete: deleteCentroCusto,
          onToggleStatus: toggleCentroCustoStatus,
          onImport: () => setModal({ type: "importCentrosCusto", data: { clienteId: clienteSocorpena.id } }),
          onDeleteAll: () => deleteAllCentrosCusto(centrosCustoSocorpena.map((cc) => cc.id)),
          onDeleteSelected: (ids) => deleteAllCentrosCusto(ids),
          onBack: () => setView("centers")
        }
      ),
      view === "clienteDetail" && selectedCliente && podeVerClientesFornecedores && /* @__PURE__ */ jsx(
        ClienteDetail,
        {
          cliente: selectedCliente,
          centrosCusto: centrosCusto.filter((cc) => cc.clienteId === selectedCliente.id),
          canManage: canManageArticles,
          onBack: () => {
            setView("clientes");
            setSelectedClienteId(null);
          },
          onEditCliente: () => setModal({ type: "cliente", data: selectedCliente }),
          onDeleteCliente: () => deleteCliente(selectedCliente.id),
          onAddCentroCusto: () => setModal({ type: "centroCusto", data: { clienteId: selectedCliente.id } }),
          onEditCentroCusto: (cc) => setModal({ type: "centroCusto", data: cc }),
          onDeleteCentroCusto: deleteCentroCusto,
          onToggleCentroCustoStatus: toggleCentroCustoStatus,
          onImportCentrosCusto: () => setModal({ type: "importCentrosCusto", data: { clienteId: selectedCliente.id } })
        }
      )
    ] }),
    modal?.type === "user" && /* @__PURE__ */ jsx(UserModal, { data: modal.data, centers, isAdmin, currentUserRole: currentUser?.role, perfisPersonalizados, onSave: saveUser, onResetPin: resetUserPin, onSetPin: setUserPin, onClose: () => setModal(null) }),
    modal?.type === "logotipo" && /* @__PURE__ */ jsx(LogotipoModal, { logotipoAtual: logotipo, onSave: saveLogotipo, onRemove: removeLogotipo, onClose: () => setModal(null) }),
    modal?.type === "dopConfig" && /* @__PURE__ */ jsx(DopConfigModal, { dopConfig, onSave: saveDopConfig, onClose: () => setModal(null) }),
    modal?.type === "backup" && /* @__PURE__ */ jsx(BackupModal, { onExport: exportarBackupCompleto, onImport: importarBackupCompleto, onClose: () => setModal(null) }),
    modal?.type === "center" && /* @__PURE__ */ jsx(CenterModal, { data: modal.data, onSave: saveCenter, onClose: () => setModal(null) }),
    modal?.type === "mixture" && /* @__PURE__ */ jsx(MixtureModal, { data: modal.data, onSave: saveMixture, onClose: () => setModal(null) }),
    modal?.type === "eliminarTodosArtigos" && /* @__PURE__ */ jsx(
      EliminarTudoModal,
      {
        titulo: "Eliminar Todos os Artigos",
        aviso: "Isto vai apagar TODOS os artigos deste centro. As f\xF3rmulas j\xE1 ligadas a estes artigos ficam sem artigo \u2014 vai ter de as voltar a ligar depois de importar os novos.",
        count: modal.data.count,
        palavraConfirmacao: "ELIMINAR",
        onConfirm: () => deleteAllMixtures(modal.data.centroId),
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "import" && /* @__PURE__ */ jsx(ImportModal, { centroId: modal.data.centroId, onImport: importArticles, onClose: () => setModal(null) }),
    modal?.type === "proveniencia" && /* @__PURE__ */ jsx(ProvenienciaModal, { data: modal.data, onSave: saveProveniencia, onClose: () => setModal(null) }),
    modal?.type === "diaria" && /* @__PURE__ */ jsx(
      DiariaModal,
      {
        data: modal.data,
        artigos: formulas.filter((f) => f.centroId === modal.data.centroId),
        clientes,
        centrosCusto,
        diarias: diarias.filter((d) => d.centroId === modal.data.centroId),
        avarias: avarias.filter((a) => a.centroId === modal.data.centroId),
        rececoes: rececoes.filter((r) => r.centroId === modal.data.centroId),
        materiais,
        consumiveis,
        center: centers.find((c) => c.id === modal.data.centroId),
        logotipo,
        isAdmin,
        onDeleteHistorico: deleteHistoricoDiaria,
        onEditHistoricoUtilizador: editarUtilizadorHistoricoDiaria,
        onEditHistoricoData: editarDataHistoricoDiaria,
        users,
        readOnly: !!modal.data.id && !canManageArticles && modal.data.dataInicio !== (/* @__PURE__ */ new Date()).toISOString().slice(0, 10),
        onSave: saveDiaria,
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "importDiarias" && /* @__PURE__ */ jsx(
      ImportDiariasModal,
      {
        centroId: modal.data.centroId,
        artigos: formulas.filter((f) => f.centroId === modal.data.centroId),
        clientes,
        centrosCusto,
        onImport: importDiarias,
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "formula" && /* @__PURE__ */ jsx(
      FormulaModal,
      {
        data: modal.data,
        readOnly: !canManageArticles,
        isAdmin,
        materiaisDisponiveis: materiais.filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(modal.data.centroId)),
        tiposMaterial,
        center: centers.find((c) => c.id === modal.data.centroId),
        artigos: mixtures.filter((m) => m.centroId === modal.data.centroId),
        formulasDoCentro: formulas.filter((f) => f.centroId === modal.data.centroId),
        onUpdateLabel: (key, label) => updateTrabalhoLabel(modal.data.centroId, key, label),
        onDeleteHistorico: deleteHistoricoFormula,
        onEditHistoricoData: editarDataHistoricoFormula,
        dopConfig,
        logotipo,
        onSaveDop: saveDopFormula,
        onDeleteDopHistorico: deleteDopHistorico,
        nomeUtilizadorAtual: currentUser?.nome,
        onSave: saveFormula,
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "importFormulas" && /* @__PURE__ */ jsx(ImportFormulasModal, { centroId: modal.data.centroId, center: selectedCenter, materiais, onImport: importFormulas, onClose: () => setModal(null) }),
    modal?.type === "avaria" && /* @__PURE__ */ jsx(AvariaModal, { data: modal.data, onSave: saveAvaria, onClose: () => setModal(null) }),
    modal?.type === "resolucao" && /* @__PURE__ */ jsx(ResolucaoModal, { item: modal.data, canManage: canManageArticles, onSave: saveResolucao, onRemoveResolucao: removerResolucao, onClose: () => setModal(null) }),
    modal?.type === "editarIncidenciaDiaria" && /* @__PURE__ */ jsx(EditarIncidenciaDiariaModal, { item: modal.data, onSave: editarIncidenciaDiaria, onClose: () => setModal(null) }),
    modal?.type === "historicoStock" && /* @__PURE__ */ jsx(
      HistoricoStockModal,
      {
        center: modal.data.center,
        produtoId: modal.data.produtoId,
        categoria: modal.data.categoria,
        materiais,
        consumiveis,
        rececoes,
        diarias,
        formulas,
        ajustesStock,
        isAdmin,
        onAddAjuste: (movimentos, unidade) => setModal({ type: "novoAjusteStock", data: { ...modal.data, movimentos, unidade } }),
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "novoAjusteStock" && /* @__PURE__ */ jsx(
      NovoAjusteModal,
      {
        centroId: modal.data.center.id,
        produtoId: modal.data.produtoId,
        categoria: modal.data.categoria,
        movimentos: modal.data.movimentos,
        unidade: modal.data.unidade,
        onSave: addAjusteStock,
        onClose: () => setModal({ type: "historicoStock", data: { center: modal.data.center, produtoId: modal.data.produtoId, categoria: modal.data.categoria } })
      }
    ),
    modal?.type === "rececao" && /* @__PURE__ */ jsx(
      RececaoModal,
      {
        data: modal.data,
        produtos: (() => {
          const cat = modal.data.categoria;
          const base = (["betumes", "agregados", "filler"].includes(cat) ? materiais : consumiveis).filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(modal.data.centroId));
          const nomeTipo = cat === "betumes" ? "Betume" : cat === "agregados" ? "Agregado" : cat === "filler" ? "Filler Comercial" : null;
          if (!nomeTipo) return base;
          const tid = tiposMaterial.find((t) => normalizeHeader(t.nome) === normalizeHeader(nomeTipo))?.id;
          return tid ? base.filter((m) => !m.tipoMaterialId || m.tipoMaterialId === tid) : base;
        })(),
        onSave: saveRececao,
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "atualizarProducao" && /* @__PURE__ */ jsx(AtualizarProducaoModal, { center: modal.data, onSave: addProducaoAnual, onClose: () => setModal(null) }),
    modal?.type === "historicoProducao" && /* @__PURE__ */ jsx(HistoricoProducaoModal, { center: modal.data, isAdmin, onEdit: (entry) => setModal({ type: "editarProducao", data: { center: modal.data, entry } }), onDelete: deleteProducaoAnual, onClose: () => setModal(null) }),
    modal?.type === "editarProducao" && /* @__PURE__ */ jsx(EditarProducaoModal, { center: modal.data.center, entry: modal.data.entry, onSave: editarProducaoAnual, onClose: () => setModal({ type: "historicoProducao", data: modal.data.center }) }),
    modal?.type === "atualizarTaxa" && /* @__PURE__ */ jsx(
      AtualizarTaxaModal,
      {
        center: modal.data.center,
        blocoKey: modal.data.blocoKey,
        titulo: modal.data.titulo,
        unidade: modal.data.unidade,
        onSave: (centroId, valor, data) => addTaxaBloco(centroId, modal.data.blocoKey, valor, data),
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "historicoTaxa" && /* @__PURE__ */ jsx(
      HistoricoTaxaModal,
      {
        center: modal.data.center,
        blocoKey: modal.data.blocoKey,
        titulo: modal.data.titulo,
        unidade: modal.data.unidade,
        isAdmin,
        onEdit: (entry) => setModal({ type: "editarTaxa", data: { center: modal.data.center, blocoKey: modal.data.blocoKey, titulo: modal.data.titulo, unidade: modal.data.unidade, entry } }),
        onDelete: (centroId, entryId) => deleteTaxaBloco(centroId, modal.data.blocoKey, entryId),
        onClose: () => setModal(null)
      }
    ),
    modal?.type === "editarTaxa" && /* @__PURE__ */ jsx(
      EditarTaxaModal,
      {
        center: modal.data.center,
        entry: modal.data.entry,
        unidade: modal.data.unidade,
        onSave: (centroId, entryId, campos) => editarTaxaBloco(centroId, modal.data.blocoKey, entryId, campos),
        onClose: () => setModal({ type: "historicoTaxa", data: { center: modal.data.center, blocoKey: modal.data.blocoKey, titulo: modal.data.titulo, unidade: modal.data.unidade } })
      }
    ),
    modal?.type === "cliente" && /* @__PURE__ */ jsx(ClienteModal, { data: modal.data, onSave: saveCliente, onClose: () => setModal(null) }),
    modal?.type === "material" && /* @__PURE__ */ jsx(MaterialModal, { data: modal.data, centers, fornecedores, tiposDesconto, tiposCustoExtra, tiposMaterial, onSave: saveMaterial, onClose: () => setModal(null) }),
    modal?.type === "precoMaterial" && /* @__PURE__ */ jsx(AtualizarPrecoModal, { material: modal.data, tiposDesconto, tiposCustoExtra, onSave: addPrecoMaterial, onClose: () => setModal(null) }),
    modal?.type === "historicoMaterial" && /* @__PURE__ */ jsx(HistoricoPrecosModal, { material: modal.data, isAdmin, onDelete: deletePrecoMaterial, onEdit: (material, entry) => setModal({ type: "editarPreco", data: { material, entry } }), onClose: () => setModal(null) }),
    modal?.type === "editarPreco" && /* @__PURE__ */ jsx(EditarPrecoModal, { material: modal.data.material, entry: modal.data.entry, tiposDesconto, tiposCustoExtra, onSave: editarPrecoMaterial, onClose: () => setModal({ type: "historicoMaterial", data: modal.data.material }) }),
    modal?.type === "escolherMaterial" && /* @__PURE__ */ jsx(EscolherMaterialModal, { materiais, onChoose: (m) => setModal({ type: "precoMaterial", data: m }), onClose: () => setModal(null) }),
    modal?.type === "importMateriais" && /* @__PURE__ */ jsx(ImportMateriaisModal, { materiaisExistentes: materiais, onImport: importMateriaisPrecos, onClose: () => setModal(null) }),
    modal?.type === "consumivel" && /* @__PURE__ */ jsx(MaterialModal, { data: modal.data, centers, fornecedores, tiposDesconto, tiposCustoExtra, tipo: "consumivel", onSave: saveConsumivel, onClose: () => setModal(null) }),
    modal?.type === "precoConsumivel" && /* @__PURE__ */ jsx(AtualizarPrecoModal, { material: modal.data, tiposDesconto, tiposCustoExtra, tipo: "consumivel", onSave: addPrecoConsumivel, onClose: () => setModal(null) }),
    modal?.type === "historicoConsumivel" && /* @__PURE__ */ jsx(HistoricoPrecosModal, { material: modal.data, isAdmin, onDelete: deletePrecoConsumivel, onEdit: (material, entry) => setModal({ type: "editarPrecoConsumivel", data: { material, entry } }), onClose: () => setModal(null) }),
    modal?.type === "editarPrecoConsumivel" && /* @__PURE__ */ jsx(EditarPrecoModal, { material: modal.data.material, entry: modal.data.entry, tiposDesconto, tiposCustoExtra, tipo: "consumivel", onSave: editarPrecoConsumivel, onClose: () => setModal({ type: "historicoConsumivel", data: modal.data.material }) }),
    modal?.type === "escolherConsumivel" && /* @__PURE__ */ jsx(EscolherMaterialModal, { materiais: consumiveis, tipo: "consumivel", onChoose: (c) => setModal({ type: "precoConsumivel", data: c }), onClose: () => setModal(null) }),
    modal?.type === "importConsumiveis" && /* @__PURE__ */ jsx(ImportMateriaisModal, { materiaisExistentes: consumiveis, tipo: "consumivel", onImport: importConsumiveisPrecos, onClose: () => setModal(null) }),
    modal?.type === "equipamento" && /* @__PURE__ */ jsx(MaterialModal, { data: modal.data, centers, fornecedores, tiposDesconto, tiposCustoExtra, tipo: "equipamento", onSave: saveEquipamento, onClose: () => setModal(null) }),
    modal?.type === "precoEquipamento" && /* @__PURE__ */ jsx(AtualizarPrecoModal, { material: modal.data, tiposDesconto, tiposCustoExtra, tipo: "equipamento", onSave: addPrecoEquipamento, onClose: () => setModal(null) }),
    modal?.type === "historicoEquipamento" && /* @__PURE__ */ jsx(HistoricoPrecosModal, { material: modal.data, isAdmin, onDelete: deletePrecoEquipamento, onEdit: (material, entry) => setModal({ type: "editarPrecoEquipamento", data: { material, entry } }), onClose: () => setModal(null) }),
    modal?.type === "editarPrecoEquipamento" && /* @__PURE__ */ jsx(EditarPrecoModal, { material: modal.data.material, entry: modal.data.entry, tiposDesconto, tiposCustoExtra, tipo: "equipamento", onSave: editarPrecoEquipamento, onClose: () => setModal({ type: "historicoEquipamento", data: modal.data.material }) }),
    modal?.type === "escolherEquipamento" && /* @__PURE__ */ jsx(EscolherMaterialModal, { materiais: equipamentos, tipo: "equipamento", onChoose: (eq) => setModal({ type: "precoEquipamento", data: eq }), onClose: () => setModal(null) }),
    modal?.type === "importEquipamentos" && /* @__PURE__ */ jsx(ImportMateriaisModal, { materiaisExistentes: equipamentos, tipo: "equipamento", onImport: importEquipamentosPrecos, onClose: () => setModal(null) }),
    modal?.type === "maodeobra" && /* @__PURE__ */ jsx(MaterialModal, { data: modal.data, centers, fornecedores, tiposDesconto, tiposCustoExtra, tipo: "maodeobra", onSave: saveMaoDeObra, onClose: () => setModal(null) }),
    modal?.type === "precoMaoObra" && /* @__PURE__ */ jsx(AtualizarPrecoModal, { material: modal.data, tiposDesconto, tiposCustoExtra, tipo: "maodeobra", onSave: addPrecoMaoObra, onClose: () => setModal(null) }),
    modal?.type === "historicoMaoObra" && /* @__PURE__ */ jsx(HistoricoPrecosModal, { material: modal.data, isAdmin, onDelete: deletePrecoMaoObra, onEdit: (material, entry) => setModal({ type: "editarPrecoMaoObra", data: { material, entry } }), onClose: () => setModal(null) }),
    modal?.type === "editarPrecoMaoObra" && /* @__PURE__ */ jsx(EditarPrecoModal, { material: modal.data.material, entry: modal.data.entry, tiposDesconto, tiposCustoExtra, tipo: "maodeobra", onSave: editarPrecoMaoObra, onClose: () => setModal({ type: "historicoMaoObra", data: modal.data.material }) }),
    modal?.type === "escolherMaoObra" && /* @__PURE__ */ jsx(EscolherMaterialModal, { materiais: maoDeObra, tipo: "maodeobra", onChoose: (mo) => setModal({ type: "precoMaoObra", data: mo }), onClose: () => setModal(null) }),
    modal?.type === "importMaoObra" && /* @__PURE__ */ jsx(ImportMateriaisModal, { materiaisExistentes: maoDeObra, tipo: "maodeobra", onImport: importMaoDeObraPrecos, onClose: () => setModal(null) }),
    modal?.type === "importClientes" && /* @__PURE__ */ jsx(ImportClientesModal, { onImport: importClientes, onClose: () => setModal(null) }),
    modal?.type === "fornecedor" && /* @__PURE__ */ jsx(FornecedorModal, { data: modal.data, onSave: saveFornecedor, onClose: () => setModal(null) }),
    modal?.type === "importFornecedores" && /* @__PURE__ */ jsx(ImportClientesModal, { onImport: importFornecedores, tituloEntidade: "Fornecedores", onClose: () => setModal(null) }),
    modal?.type === "centroCusto" && /* @__PURE__ */ jsx(CentroCustoModal, { data: modal.data, centrosCusto, onSave: saveCentroCusto, onClose: () => setModal(null) }),
    modal?.type === "importCentrosCusto" && /* @__PURE__ */ jsx(
      ImportCentrosCustoModal,
      {
        clienteId: modal.data.clienteId,
        clienteDesignacao: clientes.find((c) => c.id === modal.data.clienteId)?.designacao || "",
        centrosCustoExistentes: centrosCusto,
        onImport: importCentrosCusto,
        onClose: () => setModal(null)
      }
    ),
    confirmDialog && /* @__PURE__ */ jsx(
      ConfirmDialog,
      {
        message: confirmDialog.message,
        onConfirm: confirmDialog.onConfirm,
        onCancel: () => setConfirmDialog(null)
      }
    )
  ] });
}
function SetupForm({ onSubmit }) {
  const [nome, setNome] = useState("");
  const [email, setEmail] = useState("");
  const [pin, setPin] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const submit = () => {
    if (!nome.trim()) return setError("Indique o seu nome");
    if (!email.trim() || !/^\S+@\S+\.\S+$/.test(email.trim())) return setError("Indique um email v\xE1lido");
    if (pin.length !== 4) return setError("A palavra-passe tem 4 n\xFAmeros");
    if (pin !== confirm) return setError("As palavras-passe n\xE3o coincidem");
    onSubmit(nome.trim(), email.trim(), pin);
  };
  return /* @__PURE__ */ jsxs("div", { className: "bg-stone-900 border border-stone-800 rounded-xl p-6", children: [
    /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Nome (tamb\xE9m usado para entrar)" }), children: /* @__PURE__ */ jsx("input", { value: nome, onChange: (e) => setNome(e.target.value), className: inputCls, placeholder: "O seu nome" }) }),
    /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Email" }), children: /* @__PURE__ */ jsx("input", { value: email, onChange: (e) => setEmail(e.target.value), type: "email", className: inputCls, placeholder: "nome@empresa.pt" }) }),
    /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Palavra-passe (4 n\xFAmeros)" }), children: /* @__PURE__ */ jsx("input", { value: pin, onChange: (e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 4)), type: "password", inputMode: "numeric", className: `${inputCls} font-mono-data tracking-widest`, placeholder: "\u2022\u2022\u2022\u2022" }) }),
    /* @__PURE__ */ jsx(Field, { label: /* @__PURE__ */ jsx("span", { className: "text-stone-400", children: "Confirmar palavra-passe" }), children: /* @__PURE__ */ jsx("input", { value: confirm, onChange: (e) => setConfirm(e.target.value.replace(/\D/g, "").slice(0, 4)), type: "password", inputMode: "numeric", className: `${inputCls} font-mono-data tracking-widest`, placeholder: "\u2022\u2022\u2022\u2022" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-500 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Criar Administrador" })
  ] });
}
function UsersView({ users, isAdmin, currentUserRole, perfisPersonalizados, onAdd, onEdit, onDelete, onBack, onSavePerfil, onDeletePerfil }) {
  const [mostrarPerfis, setMostrarPerfis] = useState(false);
  const scopeText = { "Gestor": "Todos os centros", "Operador": "Por calend\xE1rio", "Or\xE7amentista": "Apenas consulta", "Convidado": "Apenas consulta" };
  const usersOrdenados = [...users].sort(
    (a, b) => ROLES.indexOf(a.role) - ROLES.indexOf(b.role) || (a.nome || "").localeCompare(b.nome || "", "pt", { sensitivity: "base" })
  );
  const podeEditar = (u) => isAdmin || currentUserRole === "Gestor" && u.role === "Operador";
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-6", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Equipa" }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: "Utilizadores" })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarPerfis(true), className: "flex items-center gap-1.5 px-4 py-2.5 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo Perfil"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo Utilizador"
        ] })
      ] })
    ] }),
    mostrarPerfis && /* @__PURE__ */ jsx(
      GerirTiposModal,
      {
        titulo: "Perfis Personalizados",
        subtitulo: "Utilizadores",
        descricao: /* @__PURE__ */ jsxs(Fragment, { children: [
          "Al\xE9m dos perfis base (Administrador, Gestor, Operador, Or\xE7amentista, Convidado), pode criar outros nomes de perfil aqui.",
          /* @__PURE__ */ jsxs("span", { className: "block mt-1 text-amber-700", children: [
            "Aten\xE7\xE3o: um perfil novo fica dispon\xEDvel para escolher, mas ",
            /* @__PURE__ */ jsx("strong", { children: "n\xE3o tem permiss\xF5es pr\xF3prias definidas" }),
            " \u2014 funciona como o mais restrito, at\xE9 me pedir para lhe atribuir regras espec\xEDficas (visualizar centros, criar di\xE1rias, etc.)."
          ] })
        ] }),
        tipos: perfisPersonalizados || [],
        placeholderNovo: "Nome do novo perfil (ex: Chefe de Central)",
        onSave: (data) => onSavePerfil(data.nome),
        onDelete: onDeletePerfil,
        onClose: () => setMostrarPerfis(false)
      }
    ),
    /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: usersOrdenados.map((u, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-4 ${i !== usersOrdenados.length - 1 ? "border-b border-stone-100" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900", children: u.nome }),
        u.email && /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400", children: u.email }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mt-1.5", children: [
          /* @__PURE__ */ jsx(RoleBadge, { role: u.role }),
          scopeText[u.role] && /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-400", children: scopeText[u.role] }),
          !u.pin && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded-full", children: "Palavra-passe por definir" })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-1", children: [
        podeEditar(u) && /* @__PURE__ */ jsx("button", { onClick: () => onEdit(u), title: isAdmin ? "Editar" : "Gerir calend\xE1rio de acessos", className: "p-2 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: isAdmin ? /* @__PURE__ */ jsx(Pencil, { size: 16 }) : /* @__PURE__ */ jsx(ClipboardList, { size: 16 }) }),
        isAdmin && /* @__PURE__ */ jsx("button", { onClick: () => onDelete(u.id), className: "p-2 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 16 }) })
      ] })
    ] }, u.id)) })
  ] });
}
function ObrasView({ cliente, centrosCusto, isAdmin, onAdd, onEdit, onDelete, onToggleStatus, onImport, onDeleteAll, onDeleteSelected, onBack }) {
  const [query, setQuery] = useState("");
  const [sortBy, setSortBy] = useState("codigo");
  const [sortDir, setSortDir] = useState("asc");
  const [selecionadas, setSelecionadas] = useState(/* @__PURE__ */ new Set());
  if (!cliente) {
    return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-3xl", children: [
      /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
        /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
        " Voltar"
      ] }),
      /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Cliente 9999" }),
      /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold mb-6", children: "Listagem Obras SCRPN" }),
      /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
        /* @__PURE__ */ jsx(MapPin, { className: "mx-auto text-stone-300 mb-3", size: 36 }),
        /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
          "Ainda n\xE3o existe o cliente ",
          /* @__PURE__ */ jsx("strong", { children: "9999 \u2014 Socorpena" }),
          ". ",
          isAdmin ? "Crie-o em Clientes para come\xE7ar a registar obras." : "Pe\xE7a a um Administrador para o criar em Clientes."
        ] })
      ] })
    ] });
  }
  const toggleSort = (field) => {
    if (sortBy === field) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else {
      setSortBy(field);
      setSortDir("asc");
    }
  };
  const filtradas = (query ? centrosCusto.filter((cc) => matchesSearch(query, cc.codigo, cc.designacao, cc.localidade)) : centrosCusto).slice().sort((a, b) => {
    const va = String(a[sortBy] || "");
    const vb = String(b[sortBy] || "");
    const cmp = va.localeCompare(vb, "pt", { numeric: true, sensitivity: "base" });
    return sortDir === "asc" ? cmp : -cmp;
  });
  const todasVisiveisSelecionadas = filtradas.length > 0 && filtradas.every((cc) => selecionadas.has(cc.id));
  const toggleTodas = () => {
    setSelecionadas((prev) => {
      if (todasVisiveisSelecionadas) {
        const novo2 = new Set(prev);
        filtradas.forEach((cc) => novo2.delete(cc.id));
        return novo2;
      }
      const novo = new Set(prev);
      filtradas.forEach((cc) => novo.add(cc.id));
      return novo;
    });
  };
  const toggleUma = (id) => {
    setSelecionadas((prev) => {
      const novo = new Set(prev);
      if (novo.has(id)) novo.delete(id);
      else novo.add(id);
      return novo;
    });
  };
  const eliminarSelecionadas = () => {
    onDeleteSelected([...selecionadas]);
    setSelecionadas(/* @__PURE__ */ new Set());
  };
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: [
          "Cliente ",
          cliente.numero,
          " \u2014 ",
          cliente.designacao
        ] }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: "Listagem Obras SCRPN" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => exportarListaExcel(
              "listagem_obras.xlsx",
              "Obras",
              ["C\xF3digo", "Designa\xE7\xE3o", "Cliente", "Local", "C\xF3digo Postal", "Estado"],
              [...centrosCusto].sort((a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true })).map((cc) => [cc.codigo || "", cc.designacao || "", cliente.designacao || "", cc.localidade || "", cc.codigoPostal || "", cc.ativo !== false ? "Ativo" : "Inativo"])
            ),
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50",
            children: [
              /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
              " Exportar Excel"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
          selecionadas.size > 0 ? /* @__PURE__ */ jsxs(
            "button",
            {
              onClick: eliminarSelecionadas,
              className: "flex items-center gap-1.5 px-3.5 py-2 bg-red-600 text-white rounded-lg text-sm font-semibold hover:bg-red-700",
              children: [
                /* @__PURE__ */ jsx(Trash2, { size: 15 }),
                " Eliminar Selecionadas (",
                selecionadas.size,
                ")"
              ]
            }
          ) : /* @__PURE__ */ jsxs(
            "button",
            {
              onClick: onDeleteAll,
              disabled: centrosCusto.length === 0,
              className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-red-200 text-red-600 rounded-lg text-sm font-semibold hover:bg-red-50 disabled:opacity-40 disabled:cursor-not-allowed",
              children: [
                /* @__PURE__ */ jsx(Trash2, { size: 15 }),
                " Eliminar Todas"
              ]
            }
          ),
          /* @__PURE__ */ jsxs("button", { onClick: onImport, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
            /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
            " Importar Excel"
          ] }),
          /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
            /* @__PURE__ */ jsx(Plus, { size: 16 }),
            " Nova Obra"
          ] })
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "relative flex-1", children: [
        /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
        /* @__PURE__ */ jsx(
          "input",
          {
            value: query,
            onChange: (e) => setQuery(e.target.value),
            placeholder: "Pesquisar por c\xF3digo, designa\xE7\xE3o ou localidade...",
            className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
          }
        ),
        query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
      ] }),
      /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-400 shrink-0 hidden sm:inline", children: "Ordenar:" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => toggleSort("codigo"),
          className: `flex items-center gap-1 px-3 py-2.5 rounded-lg text-sm font-semibold border shrink-0 ${sortBy === "codigo" ? "bg-amber-50 border-amber-300 text-amber-700" : "bg-white border-stone-300 text-stone-600 hover:bg-stone-50"}`,
          children: [
            "N\xBA ",
            sortBy === "codigo" && /* @__PURE__ */ jsx(ArrowUpDown, { size: 13, className: sortDir === "desc" ? "rotate-180" : "" })
          ]
        }
      ),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => toggleSort("designacao"),
          className: `flex items-center gap-1 px-3 py-2.5 rounded-lg text-sm font-semibold border shrink-0 ${sortBy === "designacao" ? "bg-amber-50 border-amber-300 text-amber-700" : "bg-white border-stone-300 text-stone-600 hover:bg-stone-50"}`,
          children: [
            "Nome ",
            sortBy === "designacao" && /* @__PURE__ */ jsx(ArrowUpDown, { size: 13, className: sortDir === "desc" ? "rotate-180" : "" })
          ]
        }
      )
    ] }),
    centrosCusto.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(MapPin, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 obras registadas." })
    ] }) : filtradas.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhuma obra encontrada para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: [
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 px-5 py-2 border-b border-stone-100 bg-stone-50/60", children: [
        /* @__PURE__ */ jsx("input", { type: "checkbox", checked: todasVisiveisSelecionadas, onChange: toggleTodas, className: "w-4 h-4 accent-amber-600 cursor-pointer" }),
        /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-500", children: "Selecionar todas as vis\xEDveis" })
      ] }),
      filtradas.map((cc, i) => {
        const obraAtiva = cc.ativo !== false;
        return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-3.5 ${i !== filtradas.length - 1 ? "border-b border-stone-100" : ""} ${obraAtiva ? "" : "opacity-60"} ${selecionadas.has(cc.id) ? "bg-amber-50/50" : ""}`, children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
            isAdmin && /* @__PURE__ */ jsx("input", { type: "checkbox", checked: selecionadas.has(cc.id), onChange: () => toggleUma(cc.id), className: "w-4 h-4 accent-amber-600 cursor-pointer shrink-0" }),
            cc.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-1 rounded shrink-0", children: cc.codigo }),
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
                /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: cc.designacao }),
                /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${obraAtiva ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: obraAtiva ? "Ativo" : "Inativo" })
              ] }),
              /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500", children: [cc.codigoPostal, cc.localidade].filter(Boolean).join(" ") })
            ] })
          ] }),
          /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
            isAdmin && /* @__PURE__ */ jsx(
              "button",
              {
                onClick: () => onToggleStatus(cc.id),
                title: obraAtiva ? "Desativar" : "Ativar",
                className: `p-1.5 rounded-lg ${obraAtiva ? "text-stone-400 hover:text-red-600 hover:bg-red-50" : "text-stone-400 hover:text-emerald-600 hover:bg-emerald-50"}`,
                children: obraAtiva ? /* @__PURE__ */ jsx(PowerOff, { size: 14 }) : /* @__PURE__ */ jsx(Power, { size: 14 })
              }
            ),
            /* @__PURE__ */ jsx("button", { onClick: () => onEdit(cc), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
            isAdmin && /* @__PURE__ */ jsx("button", { onClick: () => onDelete(cc.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] }, cc.id);
      })
    ] })
  ] });
}
function ClientesView({ clientes, centrosCusto, canManage, onAdd, onImport, onOpen, onDelete, onBack }) {
  const [query, setQuery] = useState("");
  const filtrados = (query ? clientes.filter((c) => matchesSearch(query, c.designacao, c.nif, c.numero)) : clientes).slice().sort((a, b) => (a.numero || "").localeCompare(b.numero || "", "pt", { numeric: true, sensitivity: "base" }));
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Base de Dados" }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: "Clientes" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => exportarListaExcel(
              "clientes.xlsx",
              "Clientes",
              ["N\xFAmero", "Designa\xE7\xE3o", "NIF", "Morada"],
              [...clientes].sort((a, b) => (a.numero || "").localeCompare(b.numero || "", "pt", { numeric: true, sensitivity: "base" })).map((c) => [c.numero || "", c.designacao || "", c.nif || "", c.morada || ""])
            ),
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50",
            children: [
              /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
              " Exportar Excel"
            ]
          }
        ),
        canManage && /* @__PURE__ */ jsxs("button", { onClick: onImport, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        canManage && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo Cliente"
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "relative mb-4", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          placeholder: "Pesquisar por nome, NIF ou n\xFAmero...",
          className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
        }
      ),
      query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
    ] }),
    clientes.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Building2, { className: "mx-auto text-stone-300 mb-3", size: 36 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 clientes registados." })
    ] }) : filtrados.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhum cliente encontrado para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: filtrados.map((c, i) => {
      const count = centrosCusto.filter((cc) => cc.clienteId === c.id).length;
      return /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => onOpen(c.id),
          className: `w-full flex items-center justify-between px-5 py-4 text-left hover:bg-stone-50 ${i !== filtrados.length - 1 ? "border-b border-stone-100" : ""}`,
          children: [
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
              /* @__PURE__ */ jsx("div", { className: "w-9 h-9 rounded-lg bg-amber-50 flex items-center justify-center shrink-0", children: /* @__PURE__ */ jsx(Building2, { className: "text-amber-600", size: 16 }) }),
              /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
                /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
                  c.numero && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: [
                    "N\xBA ",
                    c.numero
                  ] }),
                  /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: c.designacao })
                ] }),
                /* @__PURE__ */ jsxs("p", { className: "text-xs font-mono-data text-stone-400", children: [
                  "NIF ",
                  c.nif,
                  " \xB7 ",
                  count,
                  " centro",
                  count !== 1 ? "s" : "",
                  " de custo"
                ] })
              ] })
            ] }),
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
              canManage && /* @__PURE__ */ jsx(
                "span",
                {
                  onClick: (e) => {
                    e.stopPropagation();
                    onDelete(c.id);
                  },
                  className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg",
                  children: /* @__PURE__ */ jsx(Trash2, { size: 15 })
                }
              ),
              /* @__PURE__ */ jsx(ChevronRight, { className: "text-stone-300", size: 18 })
            ] })
          ]
        },
        c.id
      );
    }) })
  ] });
}
function FornecedoresView({ fornecedores, canManage, podeCriarFornecedores, onAdd, onEdit, onImport, onDelete, onBack }) {
  const [query, setQuery] = useState("");
  const filtrados = (query ? fornecedores.filter((f) => matchesSearch(query, f.designacao, f.nif, f.numero)) : fornecedores).slice().sort((a, b) => (a.numero || "").localeCompare(b.numero || "", "pt", { numeric: true, sensitivity: "base" }));
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Base de Dados" }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: "Fornecedores" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => exportarListaExcel(
              "fornecedores.xlsx",
              "Fornecedores",
              ["N\xFAmero", "Designa\xE7\xE3o", "NIF", "Morada"],
              [...fornecedores].sort((a, b) => (a.numero || "").localeCompare(b.numero || "", "pt", { numeric: true, sensitivity: "base" })).map((f) => [f.numero || "", f.designacao || "", f.nif || "", f.morada || ""])
            ),
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50",
            children: [
              /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
              " Exportar Excel"
            ]
          }
        ),
        canManage && /* @__PURE__ */ jsxs("button", { onClick: onImport, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        podeCriarFornecedores && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo Fornecedor"
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "relative mb-4", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          placeholder: "Pesquisar por nome, NIF ou n\xFAmero...",
          className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
        }
      ),
      query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
    ] }),
    fornecedores.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Truck, { className: "mx-auto text-stone-300 mb-3", size: 36 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 fornecedores registados." })
    ] }) : filtrados.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhum fornecedor encontrado para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: filtrados.map((f, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-3.5 ${i !== filtrados.length - 1 ? "border-b border-stone-100" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
        /* @__PURE__ */ jsx("div", { className: "w-9 h-9 rounded-lg bg-amber-50 flex items-center justify-center shrink-0", children: /* @__PURE__ */ jsx(Truck, { className: "text-amber-600", size: 16 }) }),
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
            f.numero && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: [
              "N\xBA ",
              f.numero
            ] }),
            /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: f.designacao })
          ] }),
          /* @__PURE__ */ jsxs("p", { className: "text-xs font-mono-data text-stone-400", children: [
            "NIF ",
            f.nif,
            f.morada ? ` \xB7 ${f.morada}` : ""
          ] })
        ] })
      ] }),
      canManage && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => onEdit(f), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("button", { onClick: () => onDelete(f.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] })
    ] }, f.id)) })
  ] });
}
function MateriaisView({ tipo, materiais, centers, tiposMaterial, tiposDesconto, tiposCustoExtra, isAdmin, onAdd, onEdit, onDelete, onAddPreco, onViewHistorico, onUpdatePricePicker, onImportPrecos, onDuplicate, onBack, onSaveTipoMaterial, onDeleteTipoMaterial, onSaveTipoDesconto, onDeleteTipoDesconto, onSaveTipoCustoExtra, onDeleteTipoCustoExtra }) {
  const [query, setQuery] = useState("");
  const [mostrarTipos, setMostrarTipos] = useState(false);
  const [mostrarDescontos, setMostrarDescontos] = useState(false);
  const [mostrarCustosExtra, setMostrarCustosExtra] = useState(false);
  const isConsumivel = tipo === "consumivel";
  const isEquipamento = tipo === "equipamento";
  const isMaoDeObra = tipo === "maodeobra";
  const isMaterial = !isConsumivel && !isEquipamento && !isMaoDeObra;
  const nomeSingular = isConsumivel ? "Combust\xEDvel" : isEquipamento ? "Equipamento" : isMaoDeObra ? "Categoria de M\xE3o de Obra" : "Material";
  const nomeBotaoNovo = isConsumivel ? "Artigo" : isMaoDeObra ? "Categoria" : nomeSingular;
  const nomePlural = isConsumivel ? "Combust\xEDveis/Energia" : isEquipamento ? "Equipamentos" : isMaoDeObra ? "M\xE3o de Obra" : "Materiais Constituintes";
  const IconeTipo = isConsumivel ? Fuel : isEquipamento ? Wrench : isMaoDeObra ? HardHat : Package;
  const filtrados = (query ? materiais.filter((m) => matchesSearch(query, m.designacao, m.fornecedor)) : materiais).slice().sort(
    (a, b) => isConsumivel || isEquipamento || isMaoDeObra ? (a.designacao || "").localeCompare(b.designacao || "", "pt", { numeric: true, sensitivity: "base" }) : (a.fornecedor || "").localeCompare(b.fornecedor || "", "pt", { sensitivity: "base" }) || (a.designacao || "").localeCompare(b.designacao || "", "pt", { numeric: true, sensitivity: "base" })
  );
  const aplicacaoTexto = (m) => {
    if (m.centrosIds === "todos") return "Todos os centros";
    const ids = Array.isArray(m.centrosIds) ? m.centrosIds : [];
    if (ids.length === 0) return "Nenhum centro";
    if (ids.length === 1) return centers.find((c) => c.id === ids[0])?.nome || "1 centro";
    return `${ids.length} centros`;
  };
  const descontoTexto = (d) => `${d.tipo === "fixo" ? `${parseFloat(d.valor || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC` : `${d.valor || 0}%`} (${d.categoria === "Outro" && d.outroTexto ? d.outroTexto : d.categoria})${d.aplicarNoCalculo === false ? " [n\xE3o entra no c\xE1lculo]" : ""}`;
  const custosExtraTexto = (h) => custosExtraLista(h).map((c) => `${c.nome}: ${parseFloat(c.valor || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC`).join(" | ");
  const exportar = () => {
    const linhas = [];
    filtrados.forEach((m) => {
      const historico = [...normalizarHistoricoPrecos(m)].sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || ""));
      const vigente = precoVigente(m);
      if (historico.length === 0) {
        linhas.push([m.designacao || "", m.fornecedor || "", "", "", "", "", "", m.unidadeCusto || "", aplicacaoTexto(m)]);
        return;
      }
      historico.forEach((h) => {
        const descontosTxt = (h.descontos || []).map((d) => descontoTexto(d)).join(" | ");
        linhas.push([
          m.designacao || "",
          m.fornecedor || "",
          h.dataEntradaVigor || "",
          h.preco ?? "",
          custosExtraTexto(h),
          descontosTxt,
          h.id === vigente?.id ? "Sim" : "",
          m.unidadeCusto || "",
          aplicacaoTexto(m)
        ]);
      });
    });
    exportarListaExcel(
      `${normalizeHeader(nomePlural).replace(/[^a-z0-9]+/g, "_")}_com_historico.xlsx`,
      nomePlural,
      ["Designa\xE7\xE3o", "Fornecedor", "Data Entrada em Vigor", "Pre\xE7o", "Custos Extra", "Descontos", "Pre\xE7o Atual?", "Unidade", "Aplica-se a"],
      linhas
    );
  };
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Base de Dados" }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: nomePlural })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        isAdmin && isMaterial && /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarTipos(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(Settings, { size: 15 }),
          " Gerir Tipos"
        ] }),
        isAdmin && (isMaterial || isConsumivel) && /* @__PURE__ */ jsxs(Fragment, { children: [
          /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarDescontos(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
            /* @__PURE__ */ jsx(Settings, { size: 15 }),
            " Gerir Descontos"
          ] }),
          /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarCustosExtra(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
            /* @__PURE__ */ jsx(Settings, { size: 15 }),
            " Gerir Custos Extra"
          ] })
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: exportar, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Exportar Excel"
        ] }),
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onImportPrecos, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: onUpdatePricePicker, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
          " Atualizar Pre\xE7o"
        ] }),
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo ",
          nomeBotaoNovo
        ] })
      ] })
    ] }),
    mostrarTipos && /* @__PURE__ */ jsx(TiposMaterialModal, { tiposMaterial, materiais, onSave: onSaveTipoMaterial, onDelete: onDeleteTipoMaterial, onClose: () => setMostrarTipos(false) }),
    mostrarDescontos && /* @__PURE__ */ jsx(
      GerirTiposModal,
      {
        titulo: "Gerir Descontos",
        subtitulo: nomePlural,
        descricao: "Tipos de desconto dispon\xEDveis ao registar o pre\xE7o \u2014 ex: Desconto comercial, B\xF3nus bom pagamento.",
        tipos: tiposDesconto,
        placeholderNovo: "Novo tipo de desconto (ex: B\xF3nus fideliza\xE7\xE3o)",
        onSave: onSaveTipoDesconto,
        onDelete: onDeleteTipoDesconto,
        onClose: () => setMostrarDescontos(false)
      }
    ),
    mostrarCustosExtra && /* @__PURE__ */ jsx(
      GerirTiposModal,
      {
        titulo: "Gerir Custos Extra",
        subtitulo: nomePlural,
        descricao: "Tipos de custo extra que se somam ao pre\xE7o \u2014 ex: Transporte, Otimiza\xE7\xE3o.",
        tipos: tiposCustoExtra,
        placeholderNovo: "Novo custo extra (ex: Seguro)",
        onSave: onSaveTipoCustoExtra,
        onDelete: onDeleteTipoCustoExtra,
        onClose: () => setMostrarCustosExtra(false)
      }
    ),
    /* @__PURE__ */ jsxs("div", { className: "relative mb-4", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          placeholder: "Pesquisar por designa\xE7\xE3o ou fornecedor...",
          className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
        }
      ),
      query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
    ] }),
    materiais.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(IconeTipo, { className: "mx-auto text-stone-300 mb-3", size: 36 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        "Ainda n\xE3o h\xE1 ",
        nomePlural.toLowerCase(),
        " registados."
      ] })
    ] }) : filtrados.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhum resultado encontrado para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: filtrados.map((m, i) => {
      const vigente = precoVigente(m);
      const nHistorico = (m.historicoPrecos || []).length;
      return /* @__PURE__ */ jsx("div", { className: `px-5 py-3.5 ${i !== filtrados.length - 1 ? "border-b border-stone-100" : ""}`, children: /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between gap-3", children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: m.designacao }),
          /* @__PURE__ */ jsx("span", { className: "inline-block text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-stone-100 text-stone-600 mt-0.5 mb-0.5", children: aplicacaoTexto(m) }),
          isEquipamento && m.unidadeCusto && /* @__PURE__ */ jsx("span", { className: "inline-block text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-amber-50 text-amber-700 mt-0.5 mb-0.5 ml-1", children: m.unidadeCusto }),
          isMaoDeObra && m.unidadeCusto && /* @__PURE__ */ jsx("span", { className: "inline-block text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-amber-50 text-amber-700 mt-0.5 mb-0.5 ml-1", children: m.unidadeCusto }),
          isConsumivel && m.unidadeCusto && /* @__PURE__ */ jsx("span", { className: "inline-block text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-amber-50 text-amber-700 mt-0.5 mb-0.5 ml-1", children: m.unidadeCusto }),
          vigente ? /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500", children: [
            m.fornecedor && /* @__PURE__ */ jsxs(Fragment, { children: [
              m.fornecedor,
              " \xB7 "
            ] }),
            /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-700", children: [
              calcularPrecoFinal(vigente).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
              " \u20AC",
              isConsumivel ? `/${unidadeBaseCombustivel(m.unidadeCusto)}` : ""
            ] }),
            (vigente.descontos && vigente.descontos.length > 0 || custosExtraLista(vigente).length > 0) && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
              " ",
              "(base ",
              parseFloat(vigente.preco || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
              " \u20AC",
              (vigente.descontos || []).map((d) => ` \u2212 ${descontoTexto(d)}`).join(""),
              custosExtraLista(vigente).map((c) => ` + ${c.nome.toLowerCase()} ${parseFloat(c.valor || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC`).join(""),
              ")"
            ] }),
            vigente.dataEntradaVigor && /* @__PURE__ */ jsxs(Fragment, { children: [
              " \xB7 em vigor desde ",
              formatDatePT(vigente.dataEntradaVigor)
            ] })
          ] }) : /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
            m.fornecedor,
            " \xB7 Sem pre\xE7o registado"
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
          /* @__PURE__ */ jsx("button", { onClick: () => onAddPreco(m), title: "Atualizar pre\xE7o", className: "p-1.5 text-stone-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg", children: /* @__PURE__ */ jsx(TrendingUp, { size: 15 }) }),
          /* @__PURE__ */ jsxs("button", { onClick: () => onViewHistorico(m), title: "Ver hist\xF3rico de pre\xE7os", className: "relative p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: [
            /* @__PURE__ */ jsx(History, { size: 15 }),
            nHistorico > 0 && /* @__PURE__ */ jsx("span", { className: "absolute -top-1 -right-1 w-4 h-4 rounded-full bg-amber-600 text-white text-[9px] font-bold flex items-center justify-center", children: nHistorico })
          ] }),
          isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsx("button", { onClick: () => onDuplicate(m), title: `Duplicar ${nomeSingular.toLowerCase()} (incluindo hist\xF3rico de pre\xE7os)`, className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Copy, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => onEdit(m), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => onDelete(m.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] })
      ] }) }, m.id);
    }) })
  ] });
}
function ClienteDetail({ cliente, centrosCusto, canManage, onBack, onEditCliente, onDeleteCliente, onAddCentroCusto, onEditCentroCusto, onDeleteCentroCusto, onToggleCentroCustoStatus, onImportCentrosCusto }) {
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-4 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Clientes"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between mb-8", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-1", children: [
          /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase", children: "Cliente" }),
          cliente.numero && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-0.5 rounded", children: [
            "N\xBA ",
            cliente.numero
          ] })
        ] }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: cliente.designacao }),
        /* @__PURE__ */ jsxs("p", { className: "text-sm font-mono-data text-stone-500 mt-1", children: [
          "NIF ",
          cliente.nif
        ] }),
        cliente.morada && /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 flex items-start gap-1 mt-1", children: [
          /* @__PURE__ */ jsx(MapPin, { size: 13, className: "mt-0.5 shrink-0" }),
          " ",
          cliente.morada
        ] })
      ] }),
      canManage && /* @__PURE__ */ jsxs("div", { className: "flex gap-1", children: [
        /* @__PURE__ */ jsx("button", { onClick: onEditCliente, className: "p-2 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 16 }) }),
        /* @__PURE__ */ jsx("button", { onClick: onDeleteCliente, className: "p-2 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 16 }) })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
        /* @__PURE__ */ jsx(MapPin, { size: 17, className: "text-amber-600" }),
        " Centros de Custo"
      ] }),
      canManage && /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs("button", { onClick: onImportCentrosCusto, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: onAddCentroCusto, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 15 }),
          " Novo Centro de Custo"
        ] })
      ] })
    ] }),
    centrosCusto.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(MapPin, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 centros de custo registados para este cliente." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: centrosCusto.map((cc, i) => {
      const obraAtiva = cc.ativo !== false;
      return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-3.5 ${i !== centrosCusto.length - 1 ? "border-b border-stone-100" : ""} ${obraAtiva ? "" : "opacity-60"}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
          cc.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-1 rounded shrink-0", children: cc.codigo }),
          /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
              /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: cc.designacao }),
              /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${obraAtiva ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: obraAtiva ? "Ativo" : "Inativo" })
            ] }),
            /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500", children: [cc.codigoPostal, cc.localidade].filter(Boolean).join(" ") })
          ] })
        ] }),
        canManage && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
          /* @__PURE__ */ jsx(
            "button",
            {
              onClick: () => onToggleCentroCustoStatus(cc.id),
              title: obraAtiva ? "Desativar" : "Ativar",
              className: `p-1.5 rounded-lg ${obraAtiva ? "text-stone-400 hover:text-red-600 hover:bg-red-50" : "text-stone-400 hover:text-emerald-600 hover:bg-emerald-50"}`,
              children: obraAtiva ? /* @__PURE__ */ jsx(PowerOff, { size: 14 }) : /* @__PURE__ */ jsx(Power, { size: 14 })
            }
          ),
          /* @__PURE__ */ jsx("button", { onClick: () => onEditCentroCusto(cc), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
          /* @__PURE__ */ jsx("button", { onClick: () => onDeleteCentroCusto(cc.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
        ] })
      ] }, cc.id);
    }) })
  ] });
}
function ClienteModal({ data, onSave, onClose }) {
  const [numero, setNumero] = useState(data?.numero || "");
  const [nif, setNif] = useState(data?.nif || "");
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [morada, setMorada] = useState(data?.morada || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!numero.trim()) return setError("Indique o n\xFAmero do cliente");
    if (!nif.trim()) return setError("Indique o NIF");
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o");
    onSave({ id: data?.id, numero: numero.trim(), nif: nif.trim(), designacao: designacao.trim(), morada: morada.trim() });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Cliente" : "Novo Cliente", subtitle: "Base de Dados de Clientes", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "N\xBA Cliente", children: /* @__PURE__ */ jsx("input", { value: numero, onChange: (e) => setNumero(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "Ex: 1045", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "NIF", children: /* @__PURE__ */ jsx("input", { value: nif, onChange: (e) => setNif(e.target.value.replace(/\D/g, "").slice(0, 9)), className: `${inputCls} font-mono-data tracking-wide`, placeholder: "500000000" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: "Nome do cliente / empresa" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Morada", children: /* @__PURE__ */ jsx("textarea", { value: morada, onChange: (e) => setMorada(e.target.value), className: inputCls, rows: 2, placeholder: "Rua, n\xFAmero, c\xF3digo postal, localidade" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function FornecedorModal({ data, onSave, onClose }) {
  const [numero, setNumero] = useState(data?.numero || "");
  const [nif, setNif] = useState(data?.nif || "");
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [morada, setMorada] = useState(data?.morada || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!numero.trim()) return setError("Indique o n\xFAmero do fornecedor");
    if (!nif.trim()) return setError("Indique o NIF");
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o");
    onSave({ id: data?.id, numero: numero.trim(), nif: nif.trim(), designacao: designacao.trim(), morada: morada.trim() });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Fornecedor" : "Novo Fornecedor", subtitle: "Base de Dados de Fornecedores", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "N\xBA Fornecedor", children: /* @__PURE__ */ jsx("input", { value: numero, onChange: (e) => setNumero(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "Ex: 1045", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "NIF", children: /* @__PURE__ */ jsx("input", { value: nif, onChange: (e) => setNif(e.target.value.replace(/\D/g, "").slice(0, 9)), className: `${inputCls} font-mono-data tracking-wide`, placeholder: "500000000" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: "Nome do fornecedor / empresa" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Morada", children: /* @__PURE__ */ jsx("textarea", { value: morada, onChange: (e) => setMorada(e.target.value), className: inputCls, rows: 2, placeholder: "Rua, n\xFAmero, c\xF3digo postal, localidade" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function CentroCustoModal({ data, centrosCusto, onSave, onClose }) {
  const [codigo, setCodigo] = useState(data?.codigo || "");
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [codigoPostal, setCodigoPostal] = useState(data?.codigoPostal || "");
  const [localidade, setLocalidade] = useState(data?.localidade || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o");
    const codigoTrim = codigo.trim();
    if (codigoTrim) {
      const duplicado = (centrosCusto || []).some((cc) => cc.id !== data?.id && String(cc.codigo || "").trim().toLowerCase() === codigoTrim.toLowerCase());
      if (duplicado) return setError(`J\xE1 existe uma obra com o n\xFAmero "${codigoTrim}". Escolha outro n\xFAmero.`);
    }
    onSave({ id: data?.id, clienteId: data.clienteId, codigo: codigoTrim, designacao: designacao.trim(), codigoPostal: codigoPostal.trim(), localidade: localidade.trim() });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Centro de Custo" : "Novo Centro de Custo", subtitle: "Centro de Custo do Cliente", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "C\xF3digo (opcional)", children: /* @__PURE__ */ jsx("input", { value: codigo, onChange: (e) => setCodigo(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "Ex: 117" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: "Ex: Obra Estrada Municipal 123", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "C\xF3digo Postal", children: /* @__PURE__ */ jsx("input", { value: codigoPostal, onChange: (e) => setCodigoPostal(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "0000-000" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Localidade", children: /* @__PURE__ */ jsx("input", { value: localidade, onChange: (e) => setLocalidade(e.target.value), className: inputCls, placeholder: "Ex: Vila Real" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function CentersView({ centers, mixtures, diarias, podeEditar, onAdd, onOpen, onImport }) {
  const [mostrarGrafico, setMostrarGrafico] = useState(false);
  const [mostrarImport, setMostrarImport] = useState(false);
  const exportar = () => {
    const linhas = [...centers].sort((a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true, sensitivity: "base" })).map((c) => [c.nome || "", c.codigo || "", c.localizacao || "", c.ativo !== false ? "Ativo" : "Inativo"]);
    exportarListaExcel("centros_producao.xlsx", "Centros", ["Nome", "C\xF3digo", "Localiza\xE7\xE3o", "Estado"], linhas);
  };
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-5xl", children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-6", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase mb-1", children: "Rede de Produ\xE7\xE3o" }),
        /* @__PURE__ */ jsx("h2", { className: "font-display text-2xl text-stone-900 font-semibold", children: "Centros de Produ\xE7\xE3o" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarGrafico(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
          " Ver Gr\xE1fico"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: exportar, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Exportar Excel"
        ] }),
        podeEditar && /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarImport(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(Upload, { size: 15 }),
          " Importar Excel"
        ] }),
        podeEditar && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 16 }),
          " Novo Centro"
        ] })
      ] })
    ] }),
    mostrarGrafico && /* @__PURE__ */ jsx(GraficoProducaoModal, { diarias: diarias || [], onClose: () => setMostrarGrafico(false) }),
    mostrarImport && /* @__PURE__ */ jsx(ImportCentersModal, { centersExistentes: centers, onImport, onClose: () => setMostrarImport(false) }),
    centers.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-12 text-center", children: [
      /* @__PURE__ */ jsx(Factory, { className: "mx-auto text-stone-300 mb-3", size: 36 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 centros registados." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "grid sm:grid-cols-2 lg:grid-cols-3 gap-4", children: [...centers].sort((a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true, sensitivity: "base" })).map((c) => {
      const count = mixtures.filter((m) => m.centroId === c.id).length;
      const ativo = c.ativo !== false;
      return /* @__PURE__ */ jsxs("button", { onClick: () => onOpen(c.id), className: `bg-white border rounded-xl p-5 text-left transition-all group ${ativo ? "border-stone-200 hover:border-amber-400 hover:shadow-md" : "border-stone-200 opacity-60"}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between mb-3", children: [
          /* @__PURE__ */ jsx("div", { className: `w-10 h-10 rounded-lg flex items-center justify-center ${ativo ? "bg-amber-50 group-hover:bg-amber-100" : "bg-stone-100"}`, children: /* @__PURE__ */ jsx(Factory, { className: ativo ? "text-amber-600" : "text-stone-400", size: 18 }) }),
          /* @__PURE__ */ jsx("span", { className: `inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold uppercase tracking-wide ${ativo ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: ativo ? "Ativo" : "Inativo" })
        ] }),
        /* @__PURE__ */ jsxs("h3", { className: "font-display font-semibold text-stone-900 mb-1 flex items-center gap-2", children: [
          c.nome,
          c.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-[10px] font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded", children: c.codigo })
        ] }),
        c.localizacao && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 flex items-center gap-1 mb-2", children: [
          /* @__PURE__ */ jsx(MapPin, { size: 11 }),
          " ",
          c.localizacao
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "text-xs font-mono-data text-stone-400", children: [
          count,
          " artigo",
          count !== 1 ? "s" : "",
          " registado",
          count !== 1 ? "s" : ""
        ] })
      ] }, c.id);
    }) })
  ] });
}
const CENTER_MENU = [
  { key: "producao", label: "Di\xE1ria de Produ\xE7\xE3o", desc: "Registo di\xE1rio, destino e toneladas", icon: ClipboardList, ready: true },
  { key: "formulas", label: "Lista de F\xF3rmulas", desc: "F\xF3rmulas de trabalho e dosifica\xE7\xE3o", icon: FlaskConical, ready: true },
  { key: "incidencias", label: "Incid\xEAncias / Avarias", desc: "Ocorr\xEAncias registadas nas di\xE1rias ou diretamente", icon: AlertTriangle, ready: true },
  { key: "agregados", label: "Rece\xE7\xE3o de Agregados", desc: "Entradas de agregados, por data e quantidade", icon: Mountain, ready: true },
  { key: "betumes", label: "Rece\xE7\xE3o de Betumes", desc: "Entradas de ligante betuminoso", icon: Droplet, ready: true },
  { key: "filler", label: "Rece\xE7\xE3o de Filler Comercial", desc: "Entradas de filler", icon: PackagePlus, ready: true },
  { key: "consumiveis", label: "Rece\xE7\xE3o de Combust\xEDveis", desc: "Entradas de consum\xEDveis diversos", icon: Box, ready: true },
  { key: "artigos", label: "Artigos", desc: "Cat\xE1logo de artigos do centro", icon: Beaker, ready: true },
  { key: "stocks", label: "Stocks", desc: "Stock atual por produto, calculado a partir das rece\xE7\xF5es", icon: Archive, ready: true },
  { key: "parametrizacao", label: "Parametriza\xE7\xE3o de Produ\xE7\xE3o", desc: "Condi\xE7\xF5es e estimativas para o c\xE1lculo de custos", icon: Settings, ready: true }
];
function CenterDetail({ center, mixtures, proveniencias, diarias, formulas, avarias, clientes, centrosCusto, canManage, podeRegistar, podeVerCustos, isAdmin, currentUserRole, onBack, onEditCenter, onDeleteCenter, onToggleStatus, onAddMixture, onEditMixture, onDeleteMixture, onToggleMixtureStatus, onImport, onOpenEliminarTodosArtigos, onAddProveniencia, onEditProveniencia, onDeleteProveniencia, onToggleProvenienciaStatus, onAddDiaria, onEditDiaria, onDeleteDiaria, onImportDiarias, onAddFormula, onEditFormula, onDeleteFormula, onImportFormulas, onDuplicateFormula, onDeleteAllFormulas, onDeleteSelectedFormulas, onToggleIncluirCustos, onAddAvaria, onEditAvaria, onDeleteAvaria, onOpenDiaria, onOpenResolucao, onEditIncidenciaDiaria, onDeleteIncidenciaDiaria, onOpenAtualizarProducao, onOpenHistoricoProducao, consumiveis, equipamentos, maoDeObra, onUpdateMaoDeObraItens, onUpdateEquipamentosItens, nomeUtilizadorAtual, logotipo, onSetCombustivel, onOpenAtualizarTaxa, onOpenHistoricoTaxa, onSetEnergiaTipo, materiais, tiposMaterial, rececoes, onAddRececao, onEditRececao, onDeleteRececao, ajustesStock, onOpenHistoricoStock }) {
  const materiaisPorTipo = (materiaisLista, nomeTipo) => {
    const tid = (tiposMaterial || []).find((t) => normalizeHeader(t.nome) === normalizeHeader(nomeTipo))?.id;
    if (!tid) return materiaisLista;
    return materiaisLista.filter((m) => !m.tipoMaterialId || m.tipoMaterialId === tid);
  };
  const [section, setSection] = useState(null);
  const ativo = center.ativo !== false;
  if (section === "artigos") {
    return /* @__PURE__ */ jsx(
      ArtigosSection,
      {
        center,
        mixtures,
        canManage,
        isAdmin,
        onBack: () => setSection(null),
        onAddMixture,
        onEditMixture,
        onDeleteMixture,
        onToggleMixtureStatus,
        onImport,
        onOpenEliminarTodosArtigos
      }
    );
  }
  if (section === "agregados") {
    const materiaisDoCentro = materiaisPorTipo((materiais || []).filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(center.id)), "Agregado");
    return /* @__PURE__ */ jsx(
      RececaoSection,
      {
        center,
        categoria: "agregados",
        titulo: "Rece\xE7\xE3o de Agregados",
        produtos: materiaisDoCentro,
        rececoes: (rececoes || []).filter((r) => r.categoria === "agregados"),
        canManage,
        podeRegistar,
        onBack: () => setSection(null),
        onAdd: () => onAddRececao("agregados"),
        onEdit: onEditRececao,
        onDelete: onDeleteRececao
      }
    );
  }
  if (section === "producao") {
    return /* @__PURE__ */ jsx(
      ProducaoSection,
      {
        center,
        diarias,
        avarias,
        clientes,
        centrosCusto,
        canManage,
        podeRegistar,
        isAdmin,
        formulas,
        materiais,
        equipamentos,
        maoDeObra,
        consumiveis,
        logotipo,
        onBack: () => setSection(null),
        onAdd: onAddDiaria,
        onEdit: onEditDiaria,
        onDelete: onDeleteDiaria,
        onImport: onImportDiarias,
        onOpenDiaria,
        onOpenResolucao
      }
    );
  }
  if (section === "formulas") {
    return /* @__PURE__ */ jsx(
      FormulasSection,
      {
        center,
        formulas,
        materiais,
        equipamentos,
        maoDeObra,
        consumiveis,
        logotipo,
        canManage,
        podeVerCustos,
        isAdmin,
        currentUserRole,
        onBack: () => setSection(null),
        onAdd: onAddFormula,
        onEdit: onEditFormula,
        onDelete: onDeleteFormula,
        onImport: onImportFormulas,
        onDuplicate: onDuplicateFormula,
        onDeleteAll: onDeleteAllFormulas,
        onDeleteSelected: onDeleteSelectedFormulas,
        onToggleIncluirCustos,
        nomeUtilizadorAtual
      }
    );
  }
  if (section === "incidencias") {
    return /* @__PURE__ */ jsx(
      IncidenciasSection,
      {
        center,
        diarias,
        avarias,
        canManage,
        isAdmin,
        onBack: () => setSection(null),
        onAdd: onAddAvaria,
        onEdit: onEditAvaria,
        onDelete: onDeleteAvaria,
        onOpenDiaria,
        onOpenResolucao,
        onEditIncidenciaDiaria,
        onDeleteIncidenciaDiaria
      }
    );
  }
  if (section === "parametrizacao") {
    return /* @__PURE__ */ jsx(
      ParametrizacaoSection,
      {
        center,
        canManage,
        isAdmin,
        onBack: () => setSection(null),
        onOpenAtualizar: onOpenAtualizarProducao,
        onOpenHistorico: onOpenHistoricoProducao,
        consumiveis,
        equipamentos,
        maoDeObra,
        onUpdateMaoDeObraItens,
        onUpdateEquipamentosItens,
        nomeUtilizadorAtual,
        onSetCombustivel,
        onOpenAtualizarTaxa,
        onOpenHistoricoTaxa,
        onSetEnergiaTipo
      }
    );
  }
  if (section === "stocks") {
    return /* @__PURE__ */ jsx(
      StocksSection,
      {
        center,
        materiais: materiais || [],
        consumiveis: consumiveis || [],
        rececoes: rececoes || [],
        diarias: diarias || [],
        formulas: formulas || [],
        ajustesStock: ajustesStock || [],
        isAdmin,
        onOpenHistoricoStock: (produtoId, categoria) => onOpenHistoricoStock({ center, produtoId, categoria }),
        onBack: () => setSection(null)
      }
    );
  }
  if (section === "betumes") {
    const materiaisDoCentro = materiaisPorTipo((materiais || []).filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(center.id)), "Betume");
    return /* @__PURE__ */ jsx(
      RececaoSection,
      {
        center,
        categoria: "betumes",
        titulo: "Rece\xE7\xE3o de Betumes",
        produtos: materiaisDoCentro,
        rececoes: (rececoes || []).filter((r) => r.categoria === "betumes"),
        canManage,
        podeRegistar,
        onBack: () => setSection(null),
        onAdd: () => onAddRececao("betumes"),
        onEdit: onEditRececao,
        onDelete: onDeleteRececao
      }
    );
  }
  if (section === "filler") {
    const materiaisDoCentro = materiaisPorTipo((materiais || []).filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(center.id)), "Filler Comercial");
    return /* @__PURE__ */ jsx(
      RececaoSection,
      {
        center,
        categoria: "filler",
        titulo: "Rece\xE7\xE3o de Filler Comercial",
        produtos: materiaisDoCentro,
        rececoes: (rececoes || []).filter((r) => r.categoria === "filler"),
        canManage,
        podeRegistar,
        onBack: () => setSection(null),
        onAdd: () => onAddRececao("filler"),
        onEdit: onEditRececao,
        onDelete: onDeleteRececao
      }
    );
  }
  if (section === "consumiveis") {
    const consumiveisDoCentro = (consumiveis || []).filter((m) => m.centrosIds === "todos" || Array.isArray(m.centrosIds) && m.centrosIds.includes(center.id));
    return /* @__PURE__ */ jsx(
      RececaoSection,
      {
        center,
        categoria: "consumiveis",
        titulo: "Rece\xE7\xE3o de Combust\xEDveis",
        produtos: consumiveisDoCentro,
        rececoes: (rececoes || []).filter((r) => r.categoria === "consumiveis"),
        canManage,
        podeRegistar,
        onBack: () => setSection(null),
        onAdd: () => onAddRececao("consumiveis"),
        onEdit: onEditRececao,
        onDelete: onDeleteRececao
      }
    );
  }
  if (section) {
    const item = CENTER_MENU.find((m) => m.key === section);
    return /* @__PURE__ */ jsx(PlaceholderSection, { title: item.label, onBack: () => setSection(null) });
  }
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-5xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-4 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Centros de Produ\xE7\xE3o"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between mb-6", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-1", children: [
          /* @__PURE__ */ jsx("p", { className: "font-display text-xs tracking-[0.2em] text-amber-600 uppercase", children: "Centro de Produ\xE7\xE3o" }),
          /* @__PURE__ */ jsx("span", { className: `inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-semibold uppercase tracking-wide ${ativo ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: ativo ? "Ativo" : "Inativo" })
        ] }),
        /* @__PURE__ */ jsxs("h2", { className: "font-display text-2xl text-stone-900 font-semibold flex items-center gap-2", children: [
          center.nome,
          center.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-0.5 rounded", children: center.codigo })
        ] }),
        center.localizacao && /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 flex items-center gap-1 mt-1", children: [
          /* @__PURE__ */ jsx(MapPin, { size: 13 }),
          " ",
          center.localizacao
        ] })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1", children: [
        /* @__PURE__ */ jsx(
          "button",
          {
            onClick: onToggleStatus,
            title: ativo ? "Desativar centro" : "Ativar centro",
            className: `p-2 rounded-lg ${ativo ? "text-stone-400 hover:text-red-600 hover:bg-red-50" : "text-stone-400 hover:text-emerald-600 hover:bg-emerald-50"}`,
            children: ativo ? /* @__PURE__ */ jsx(PowerOff, { size: 16 }) : /* @__PURE__ */ jsx(Power, { size: 16 })
          }
        ),
        /* @__PURE__ */ jsx("button", { onClick: onEditCenter, className: "p-2 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 16 }) }),
        /* @__PURE__ */ jsx("button", { onClick: onDeleteCenter, className: "p-2 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 16 }) })
      ] })
    ] }),
    !ativo && /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 border border-stone-200 rounded-lg px-4 py-3 mb-6 text-sm text-stone-600", children: [
      "Este centro est\xE1 ",
      /* @__PURE__ */ jsx("strong", { children: "inativo" }),
      ". Continua vis\xEDvel e consult\xE1vel, mas foi assinalado como fora de opera\xE7\xE3o."
    ] }),
    /* @__PURE__ */ jsx("div", { className: "grid sm:grid-cols-2 gap-3", children: CENTER_MENU.filter((item) => (item.key !== "parametrizacao" || isAdmin) && (currentUserRole !== "Or\xE7amentista" || item.key === "formulas")).map((item) => {
      const Icon = item.icon;
      return /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => setSection(item.key),
          className: "bg-white border border-stone-200 rounded-xl p-4 text-left hover:border-amber-400 hover:shadow-md transition-all group flex items-center gap-4",
          children: [
            /* @__PURE__ */ jsx("div", { className: "w-11 h-11 rounded-lg bg-amber-50 flex items-center justify-center shrink-0 group-hover:bg-amber-100", children: /* @__PURE__ */ jsx(Icon, { className: "text-amber-600", size: 20 }) }),
            /* @__PURE__ */ jsxs("div", { className: "min-w-0 flex-1", children: [
              /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
                /* @__PURE__ */ jsx("h3", { className: "font-display font-semibold text-stone-900", children: item.label }),
                !item.ready && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide text-stone-400 bg-stone-100 px-1.5 py-0.5 rounded-full shrink-0", children: "Em breve" })
              ] }),
              /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500 truncate", children: item.desc })
            ] }),
            /* @__PURE__ */ jsx(ChevronRight, { className: "text-stone-300 group-hover:text-amber-500 shrink-0", size: 18 })
          ]
        },
        item.key
      );
    }) })
  ] });
}
function CombustivelBlocoCard({ center, blocoKey, titulo, descricao, consumiveis, isAdmin, onSetCombustivel, onOpenAtualizarTaxa, onOpenHistoricoTaxa, bg, campoLabel, unidade, campoValorLabel, acaoLabel }) {
  const bloco = center.parametrizacao?.[blocoKey] || {};
  const combustivelId = bloco.combustivelId || "";
  const combustivelAtual = (consumiveis || []).find((c) => c.id === combustivelId);
  const un = combustivelAtual?.unidadeCusto || unidade || "L/t";
  const rotuloCampo = campoLabel || "Combust\xEDvel";
  const rotuloValor = campoValorLabel || "Consumo";
  const rotuloAcao = acaoLabel || "Consumo";
  const historicoTaxa = [...bloco.historico || []].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || "") || (b.dataRegisto || "").localeCompare(a.dataRegisto || ""));
  const taxaVigente = historicoTaxa[0] || null;
  return /* @__PURE__ */ jsxs("div", { className: `${bg} border border-stone-200 rounded-xl p-5 mb-4`, children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: titulo }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: descricao }),
    /* @__PURE__ */ jsx(Field, { label: rotuloCampo, children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: combustivelId, materiais: consumiveis || [], isAdmin, disabled: !isAdmin, onChange: (v) => onSetCombustivel(center.id, blocoKey, v) }) }),
    combustivelId && (taxaVigente ? /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-lg px-3 py-2.5 mb-3 text-sm text-stone-600 border border-stone-200", children: [
      rotuloValor,
      " em vigor: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        parseFloat(taxaVigente.valor).toLocaleString("pt-PT"),
        " ",
        un
      ] }),
      taxaVigente.dataEntradaVigor && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
        " (desde ",
        formatDatePT(taxaVigente.dataEntradaVigor),
        ")"
      ] })
    ] }) : /* @__PURE__ */ jsxs("div", { className: "bg-amber-50 rounded-lg px-3 py-2.5 mb-3 text-sm text-amber-700", children: [
      "Ainda n\xE3o h\xE1 ",
      rotuloValor.toLowerCase(),
      " registado para ",
      combustivelAtual?.designacao || `este ${rotuloCampo.toLowerCase()}`,
      "."
    ] })),
    combustivelId && /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
      isAdmin && /* @__PURE__ */ jsxs("button", { onClick: () => onOpenAtualizarTaxa(blocoKey, titulo, un), className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700", children: [
        /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
        " Atualizar ",
        rotuloAcao
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: () => onOpenHistoricoTaxa(blocoKey, titulo, un), className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg border border-stone-300 text-stone-700 text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(History, { size: 15 }),
        " Ver Hist\xF3rico de ",
        rotuloAcao
      ] })
    ] })
  ] });
}
function ValorBlocoCard({ center, blocoKey, titulo, descricao, unidade, isAdmin, onOpenAtualizarTaxa, onOpenHistoricoTaxa, bg }) {
  const bloco = center.parametrizacao?.[blocoKey] || {};
  const historicoTaxa = [...bloco.historico || []].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || "") || (b.dataRegisto || "").localeCompare(a.dataRegisto || ""));
  const taxaVigente = historicoTaxa[0] || null;
  return /* @__PURE__ */ jsxs("div", { className: `${bg} border border-stone-200 rounded-xl p-5 mb-4`, children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: titulo }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: descricao }),
    taxaVigente ? /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-lg px-3 py-2.5 mb-3 text-sm text-stone-600 border border-stone-200", children: [
      "Valor em vigor: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        parseFloat(taxaVigente.valor).toLocaleString("pt-PT"),
        " ",
        unidade
      ] }),
      taxaVigente.dataEntradaVigor && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
        " (desde ",
        formatDatePT(taxaVigente.dataEntradaVigor),
        ")"
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-amber-50 rounded-lg px-3 py-2.5 mb-3 text-sm text-amber-700", children: "Ainda n\xE3o h\xE1 valor registado." }),
    /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
      isAdmin && /* @__PURE__ */ jsxs("button", { onClick: () => onOpenAtualizarTaxa(blocoKey, titulo, unidade), className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700", children: [
        /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
        " Atualizar Valor"
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: () => onOpenHistoricoTaxa(blocoKey, titulo, unidade), className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg border border-stone-300 text-stone-700 text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(History, { size: 15 }),
        " Ver Hist\xF3rico"
      ] })
    ] })
  ] });
}
function ParametrizacaoSection({ center, canManage, isAdmin, onBack, onOpenAtualizar, onOpenHistorico, consumiveis, equipamentos, maoDeObra, onUpdateMaoDeObraItens, onUpdateEquipamentosItens, nomeUtilizadorAtual, onSetCombustivel, onOpenAtualizarTaxa, onOpenHistoricoTaxa, onSetEnergiaTipo }) {
  const anoAtual = (/* @__PURE__ */ new Date()).getFullYear();
  const historico = [...center.parametrizacao?.producaoAnualHistorico || []].sort((a, b) => (a.ano || 0) - (b.ano || 0) || (a.dataRegisto || "").localeCompare(b.dataRegisto || ""));
  const valorVigenteAno = (anoRef) => {
    const doAno = historico.filter((h) => h.ano === anoRef);
    if (doAno.length === 0) return null;
    return [...doAno].sort((a, b) => (b.dataRegisto || "").localeCompare(a.dataRegisto || ""))[0];
  };
  const vigenteAtual = valorVigenteAno(anoAtual);
  const energiaTipo = center.parametrizacao?.energiaTipo || "transformador";
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-3xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mb-6 flex items-start justify-between gap-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
          /* @__PURE__ */ jsx(Settings, { size: 17, className: "text-amber-600" }),
          " Parametriza\xE7\xE3o de Produ\xE7\xE3o"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mt-1", children: "Condi\xE7\xF5es e estimativas usadas no c\xE1lculo do custo das misturas betuminosas deste centro \u2014 base para os custos indiretos." })
      ] }),
      canManage && /* @__PURE__ */ jsxs("button", { onClick: () => {
        const linhas = [];
        [
          ["custoCertificacao", "Certifica\xE7\xE3o (\u20AC/ano)"],
          ["custoEstaleiro", "Aluguer de Estaleiro (\u20AC/ano)"],
          ["custoQAS", "Qualidade/Ambiente/Seguran\xE7a (\u20AC/ano)"],
          ["custoAmortizacao", "Amortiza\xE7\xE3o da Central (\u20AC/ano)"],
          ["custoMudancaCentral", "Mudan\xE7a de Central (\u20AC/ano)"],
          ["custoManutencaoCentral", "Manuten\xE7\xE3o da Central (\u20AC/ano)"],
          ["custoControloLab", "Controlo Laboratorial (\u20AC/t)"],
          ["custoAluguerCentral", "Aluguer da Central (\u20AC/t)"]
        ].forEach(([key, nome]) => {
          (center.parametrizacao?.[key]?.historico || []).forEach((h) => linhas.push([nome, h.dataEntradaVigor || "", h.valor ?? "", ""]));
        });
        [["blocoTermico", "Bloco T\xE9rmico"], ["queimadorTamborSecador", "Queimador Tambor Secador"], ["energiaPrincipal", "Energia Principal"], ["energiaAuxiliar", "Energia Auxiliar"]].forEach(([key, nome]) => {
          const bloco = center.parametrizacao?.[key];
          if (!bloco) return;
          const combustivel = (consumiveis || []).find((c) => c.id === bloco.combustivelId);
          (bloco.historico || []).forEach((h) => linhas.push([`${nome}${combustivel ? ` (${combustivel.designacao})` : ""}`, h.dataEntradaVigor || "", h.valor ?? "", combustivel?.unidadeCusto || ""]));
        });
        (center.parametrizacao?.producaoAnualHistorico || []).forEach((h) => linhas.push([`Produ\xE7\xE3o Anual Estimada ${h.ano || ""}`, (h.dataRegisto || "").slice(0, 10), h.valor ?? "", "t"]));
        (center.parametrizacao?.equipamentosItens || []).forEach((item) => {
          const eq = (equipamentos || []).find((e) => e.id === item.equipamentoId);
          (item.historico || []).forEach((h) => linhas.push([`Equipamento: ${eq?.designacao || "\u2014"} \u2014 Horas Di\xE1rias`, h.dataEntradaVigor || "", h.horasDiarias ?? "", "h/dia"]));
        });
        (center.parametrizacao?.maoDeObraItens || []).forEach((item) => {
          const mo = (maoDeObra || []).find((m) => m.id === item.maoDeObraId);
          (item.historico || []).forEach((h) => linhas.push([`M\xE3o de Obra: ${mo?.designacao || "\u2014"} \u2014 Quantidade`, h.dataEntradaVigor || "", h.quantidade ?? "", "elementos"]));
        });
        exportarListaExcel("parametrizacao_producao.xlsx", "Parametriza\xE7\xE3o", ["Par\xE2metro", "Data Entrada em Vigor", "Valor", "Unidade"], linhas);
      }, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50 shrink-0", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Exportar Excel"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-stone-200 rounded-xl p-5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Produ\xE7\xE3o Anual Estimada" }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Usada para repartir os custos indiretos fixos (2.1) por tonelada produzida. Cada ano tem a sua pr\xF3pria estimativa, e cada atualiza\xE7\xE3o fica registada no hist\xF3rico." }),
      vigenteAtual ? /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-3 text-sm text-stone-600", children: [
        "Estimativa em vigor para ",
        anoAtual,
        ": ",
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
          parseFloat(vigenteAtual.valor).toLocaleString("pt-PT"),
          " t"
        ] })
      ] }) : /* @__PURE__ */ jsxs("div", { className: "bg-amber-50 rounded-lg px-3 py-2.5 mb-3 text-sm text-amber-700", children: [
        "Ainda n\xE3o h\xE1 estimativa registada para ",
        anoAtual,
        "."
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onOpenAtualizar, className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
          " Atualizar Produ\xE7\xE3o"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: onOpenHistorico, className: "flex items-center gap-1.5 px-4 py-2.5 rounded-lg border border-stone-300 text-stone-700 text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(History, { size: 15 }),
          " Ver Hist\xF3rico de Previs\xE3o"
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsx(
      CombustivelBlocoCard,
      {
        center,
        blocoKey: "blocoTermico",
        titulo: "Bloco T\xE9rmico",
        descricao: "Combust\xEDvel gasto no bloco t\xE9rmico, e consumo estimado por tonelada produzida \u2014 a unidade (Lt/Ton, Kg/Ton, Ton/Ton...) vem definida no pr\xF3prio combust\xEDvel. Usado para o custo indireto por tonelada (2.2).",
        consumiveis,
        isAdmin,
        onSetCombustivel,
        onOpenAtualizarTaxa,
        onOpenHistoricoTaxa,
        bg: "bg-stone-200"
      }
    ),
    /* @__PURE__ */ jsx(
      CombustivelBlocoCard,
      {
        center,
        blocoKey: "queimadorTamborSecador",
        titulo: "Queimador Tambor Secador",
        descricao: "Combust\xEDvel gasto no queimador do tambor secador, e consumo estimado por tonelada produzida \u2014 a unidade (Lt/Ton, Kg/Ton, Ton/Ton...) vem definida no pr\xF3prio combust\xEDvel. Usado para o custo indireto por tonelada (2.2).",
        consumiveis,
        isAdmin,
        onSetCombustivel,
        onOpenAtualizarTaxa,
        onOpenHistoricoTaxa,
        bg: "bg-white"
      }
    ),
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-stone-200 rounded-xl p-5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Energia" }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Fonte de energia el\xE9trica da central \u2014 posto transformador (rede) ou gerador \u2014 e o respetivo consumo de combust\xEDvel por tonelada produzida." }),
      /* @__PURE__ */ jsxs("div", { className: "flex rounded-lg border border-stone-300 overflow-hidden mb-4 max-w-sm", children: [
        /* @__PURE__ */ jsx(
          "button",
          {
            type: "button",
            disabled: !isAdmin,
            onClick: () => onSetEnergiaTipo(center.id, "transformador"),
            className: `flex-1 py-2 text-xs font-semibold ${energiaTipo === "transformador" ? "bg-amber-100 text-amber-800" : "bg-white text-stone-500"}`,
            children: "Posto Transformador"
          }
        ),
        /* @__PURE__ */ jsx(
          "button",
          {
            type: "button",
            disabled: !isAdmin,
            onClick: () => onSetEnergiaTipo(center.id, "gerador"),
            className: `flex-1 py-2 text-xs font-semibold border-l border-stone-300 ${energiaTipo === "gerador" ? "bg-amber-100 text-amber-800" : "bg-white text-stone-500"}`,
            children: "Gerador"
          }
        )
      ] }),
      /* @__PURE__ */ jsx(
        CombustivelBlocoCard,
        {
          center,
          blocoKey: "energiaPrincipal",
          titulo: energiaTipo === "gerador" ? "Gerador Principal" : "Posto Transformador",
          descricao: energiaTipo === "gerador" ? "Fonte de energia e consumo por tonelada do gerador principal." : "Fonte de energia e consumo por tonelada associado ao posto transformador.",
          consumiveis,
          isAdmin,
          onSetCombustivel,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          campoLabel: "Fonte",
          bg: "bg-stone-200"
        }
      ),
      energiaTipo === "gerador" && /* @__PURE__ */ jsx(
        CombustivelBlocoCard,
        {
          center,
          blocoKey: "energiaAuxiliar",
          titulo: "Gerador Auxiliar",
          descricao: "Fonte de energia e consumo por tonelada do gerador auxiliar.",
          consumiveis,
          isAdmin,
          onSetCombustivel,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          campoLabel: "Fonte",
          bg: "bg-white"
        }
      )
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-200 border border-stone-200 rounded-xl p-5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Valores Fixos" }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500 mb-3", children: "Valores globais anuais, repartidos pela produ\xE7\xE3o anual estimada para chegar ao custo por tonelada." }),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoCertificacao",
          titulo: "Certifica\xE7\xE3o (\u20AC)",
          descricao: "Valor global anual de certifica\xE7\xE3o.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-white"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoEstaleiro",
          titulo: "Aluguer de Estaleiro (\u20AC)",
          descricao: "Valor global anual de aluguer de estaleiro.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-stone-100"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoQAS",
          titulo: "Qualidade/Ambiente/Seguran\xE7a (\u20AC)",
          descricao: "Valor global anual de Qualidade, Ambiente e Seguran\xE7a.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-white"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoAmortizacao",
          titulo: "Amortiza\xE7\xE3o da Central (\u20AC)",
          descricao: "Valor global anual de amortiza\xE7\xE3o da central.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-stone-100"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoMudancaCentral",
          titulo: "Mudan\xE7a de Central (\u20AC)",
          descricao: "Valor global anual relativo \xE0 mudan\xE7a de central.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-white"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoManutencaoCentral",
          titulo: "Manuten\xE7\xE3o da Central (\u20AC)",
          descricao: "Valor global anual de manuten\xE7\xE3o da central.",
          unidade: "\u20AC",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-stone-100"
        }
      ),
      /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg p-4 mt-3", children: [
        /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Equipamento" }),
        /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 mb-2", children: [
          "Escolha o(s) equipamento(s) afetos a este centro (podem ser v\xE1rios). Para equipamento em \u20AC/hora, o custo anual = ",
          DIAS_UTEIS_ANO_PADRAO,
          " dias \xD7 horas di\xE1rias \xD7 valor/hora \u2014 o que varia e tem hist\xF3rico s\xE3o as horas di\xE1rias de trabalho. Equipamento em \u20AC/dia usa ",
          DIAS_UTEIS_ANO_PADRAO,
          " \xD7 valor/dia; em \u20AC/tonelada usa-se o valor diretamente."
        ] }),
        /* @__PURE__ */ jsx(
          EquipamentosItensCard,
          {
            center,
            equipamentos: equipamentos || [],
            isAdmin,
            nomeUtilizadorAtual,
            onUpdateItens: (itens) => onUpdateEquipamentosItens(center.id, itens)
          }
        )
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-lg p-4 mt-3", children: [
        /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "M\xE3o de Obra" }),
        /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 mb-2", children: [
          "Custo anual = ",
          DIAS_UTEIS_ANO_PADRAO,
          " dias \xFAteis (m\xE9dia fixa) \xD7 10 h/dia (m\xE9dia) \xD7 valor da categoria (\u20AC/hora) \xD7 quantidade \u2014 repartido pela produ\xE7\xE3o anual estimada para chegar ao custo por tonelada. Categorias em \u20AC/dia usam ",
          DIAS_UTEIS_ANO_PADRAO,
          " \xD7 valor \xD7 quantidade, sem a m\xE9dia de horas. O que varia e fica com hist\xF3rico \xE9 o n\xFAmero de elementos (quantidade) afetos \xE0 produ\xE7\xE3o em cada categoria."
        ] }),
        /* @__PURE__ */ jsx(
          MaoDeObraItensCard,
          {
            center,
            maoDeObra: maoDeObra || [],
            isAdmin,
            nomeUtilizadorAtual,
            onUpdateItens: (itens) => onUpdateMaoDeObraItens(center.id, itens)
          }
        )
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-stone-200 rounded-xl p-5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Custos Vari\xE1veis" }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500 mb-3", children: "Custos por tonelada produzida." }),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoControloLab",
          titulo: "Controlo Laboratorial (\u20AC/tonelada)",
          descricao: "Custo por tonelada de controlo laboratorial.",
          unidade: "\u20AC/t",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-stone-100"
        }
      ),
      /* @__PURE__ */ jsx(
        ValorBlocoCard,
        {
          center,
          blocoKey: "custoAluguerCentral",
          titulo: "Aluguer da Central (\u20AC/tonelada)",
          descricao: "Custo por tonelada produzida de aluguer da central (alternativa \xE0 amortiza\xE7\xE3o, se a central for alugada).",
          unidade: "\u20AC/t",
          isAdmin,
          onOpenAtualizarTaxa,
          onOpenHistoricoTaxa,
          bg: "bg-white"
        }
      )
    ] }),
    /* @__PURE__ */ jsx("div", { className: "bg-stone-50 border border-dashed border-stone-300 rounded-xl p-5 text-sm text-stone-400", children: 'Estes par\xE2metros j\xE1 entram no c\xE1lculo do custo final por tonelada \u2014 veja a "Ficha de Custo" de cada f\xF3rmula, em Lista de F\xF3rmulas.' })
  ] });
}
function EquipamentosItensCard({ center, equipamentos, isAdmin, nomeUtilizadorAtual, onUpdateItens }) {
  const itens = center.parametrizacao?.equipamentosItens || [];
  const [novoEquipId, setNovoEquipId] = useState("");
  const [novasHoras, setNovasHoras] = useState("");
  const [novaData, setNovaData] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [itemAtualizar, setItemAtualizar] = useState(null);
  const [itemHistorico, setItemHistorico] = useState(null);
  const nomeEquip = (id) => equipamentos.find((e) => e.id === id)?.designacao || "Equipamento removido";
  const unidadeEquip = (id) => equipamentos.find((e) => e.id === id)?.unidadeCusto || "";
  const valorVigenteEquip = (id) => {
    const eq = equipamentos.find((e) => e.id === id);
    if (!eq) return null;
    return precoVigente(eq);
  };
  const horasVigentes = (item) => {
    const hist = [...item.historico || []].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
    return hist[0] || null;
  };
  const precisaHoras = (id) => unidadeEquip(id) === "\u20AC/hora";
  const adicionar = () => {
    if (!novoEquipId) return;
    if (itens.some((i) => i.equipamentoId === novoEquipId)) return;
    const horas = parseFloat(novasHoras) || 0;
    const entrada = { id: genId(), horasDiarias: horas, dataEntradaVigor: novaData, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: nomeUtilizadorAtual || "" };
    onUpdateItens([...itens, { id: genId(), equipamentoId: novoEquipId, historico: [entrada] }]);
    setNovoEquipId("");
    setNovasHoras("");
    setNovaData((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  };
  const remover = (id) => onUpdateItens(itens.filter((i) => i.id !== id));
  return /* @__PURE__ */ jsxs("div", { className: "mt-2", children: [
    itens.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Ainda n\xE3o h\xE1 equipamento atribu\xEDdo a este centro." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-3", children: itens.map((item, i) => {
      const vigente = valorVigenteEquip(item.equipamentoId);
      const hVigente = horasVigentes(item);
      const unidade = unidadeEquip(item.equipamentoId);
      return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== itens.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsx("p", { className: "text-stone-800 font-medium truncate", children: nomeEquip(item.equipamentoId) }),
          /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
            vigente && /* @__PURE__ */ jsxs(Fragment, { children: [
              calcularPrecoFinal(vigente).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
              " ",
              unidade
            ] }),
            unidade === "\u20AC/hora" && /* @__PURE__ */ jsxs(Fragment, { children: [
              " \xB7 Horas di\xE1rias: ",
              /* @__PURE__ */ jsx("span", { className: "font-mono-data text-stone-600 font-semibold", children: hVigente ? hVigente.horasDiarias : "\u2014" }),
              hVigente?.dataEntradaVigor && /* @__PURE__ */ jsxs(Fragment, { children: [
                " (desde ",
                formatDatePT(hVigente.dataEntradaVigor),
                ")"
              ] })
            ] })
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
          unidade === "\u20AC/hora" && /* @__PURE__ */ jsx("button", { onClick: () => setItemHistorico(item), title: "Ver hist\xF3rico de horas di\xE1rias", className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(History, { size: 14 }) }),
          isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
            unidade === "\u20AC/hora" && /* @__PURE__ */ jsx("button", { onClick: () => setItemAtualizar(item), title: "Atualizar horas di\xE1rias", className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(TrendingUp, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => remover(item.id), title: "Remover equipamento deste centro", className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] })
      ] }, item.id);
    }) }),
    isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-2", children: [
      /* @__PURE__ */ jsx("div", { className: "flex-1", children: /* @__PURE__ */ jsx(Field, { label: "Equipamento", children: /* @__PURE__ */ jsxs("select", { value: novoEquipId, onChange: (e) => setNovoEquipId(e.target.value), className: inputCls, children: [
        /* @__PURE__ */ jsx("option", { value: "", children: "Escolha um equipamento..." }),
        equipamentos.filter((e) => !itens.some((i) => i.equipamentoId === e.id)).map((e) => /* @__PURE__ */ jsx("option", { value: e.id, children: e.designacao }, e.id))
      ] }) }) }),
      precisaHoras(novoEquipId) && /* @__PURE__ */ jsx("div", { className: "w-28", children: /* @__PURE__ */ jsx(Field, { label: "Horas/dia", children: /* @__PURE__ */ jsx("input", { value: novasHoras, onChange: (e) => setNovasHoras(e.target.value), type: "number", step: "0.5", min: "0", className: `${inputCls} font-mono-data` }) }) }),
      precisaHoras(novoEquipId) && /* @__PURE__ */ jsx("div", { className: "w-40", children: /* @__PURE__ */ jsx(Field, { label: "Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: novaData, onChange: (e) => setNovaData(e.target.value), className: inputCls }) }) }),
      /* @__PURE__ */ jsx("button", { onClick: adicionar, disabled: !novoEquipId, className: "mb-4 px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700 shrink-0 disabled:opacity-40 disabled:cursor-not-allowed", children: "Adicionar" })
    ] }),
    itemAtualizar && /* @__PURE__ */ jsx(
      AtualizarHorasEquipamentoModal,
      {
        nomeEquipamento: nomeEquip(itemAtualizar.equipamentoId),
        valorAtual: horasVigentes(itemAtualizar),
        onSave: (horasDiarias, dataEntradaVigor) => {
          const entrada = { id: genId(), horasDiarias, dataEntradaVigor, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: nomeUtilizadorAtual || "" };
          onUpdateItens(itens.map((i) => i.id === itemAtualizar.id ? { ...i, historico: [...i.historico || [], entrada] } : i));
          setItemAtualizar(null);
        },
        onClose: () => setItemAtualizar(null)
      }
    ),
    itemHistorico && /* @__PURE__ */ jsx(
      HistoricoHorasEquipamentoModal,
      {
        nomeEquipamento: nomeEquip(itemHistorico.equipamentoId),
        historico: itens.find((i) => i.id === itemHistorico.id)?.historico || [],
        isAdmin,
        onDelete: (entryId) => onUpdateItens(itens.map((i) => i.id === itemHistorico.id ? { ...i, historico: (i.historico || []).filter((h) => h.id !== entryId) } : i)),
        onEdit: (entryId, novasHoras2, novaData2) => onUpdateItens(itens.map((i) => i.id === itemHistorico.id ? { ...i, historico: (i.historico || []).map((h) => h.id === entryId ? { ...h, horasDiarias: novasHoras2, dataEntradaVigor: novaData2 } : h) } : i)),
        onClose: () => setItemHistorico(null)
      }
    )
  ] });
}
function AtualizarHorasEquipamentoModal({ nomeEquipamento, valorAtual, onSave, onClose }) {
  const [horas, setHoras] = useState("");
  const [dataEntradaVigor, setDataEntradaVigor] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!horas || parseFloat(horas) < 0) return setError("Indique as horas di\xE1rias de trabalho");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    if (valorAtual && parseFloat(horas) === parseFloat(valorAtual.horasDiarias) && dataEntradaVigor === valorAtual.dataEntradaVigor) {
      return setError("Sem altera\xE7\xF5es em rela\xE7\xE3o ao valor j\xE1 em vigor");
    }
    onSave(parseFloat(horas), dataEntradaVigor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Horas Di\xE1rias", subtitle: nomeEquipamento, onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Horas Di\xE1rias de Trabalho", children: /* @__PURE__ */ jsx("input", { value: horas, onChange: (e) => setHoras(e.target.value), type: "number", step: "0.5", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Em Vigor Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EditarHorasEquipamentoModal({ nomeEquipamento, entrada, onSave, onClose }) {
  const [horas, setHoras] = useState(String(entrada.horasDiarias ?? ""));
  const [dataEntradaVigor, setDataEntradaVigor] = useState(entrada.dataEntradaVigor || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!horas || parseFloat(horas) < 0) return setError("Indique as horas di\xE1rias de trabalho");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    onSave(parseFloat(horas), dataEntradaVigor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Registo de Horas", subtitle: nomeEquipamento, onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Horas Di\xE1rias de Trabalho", children: /* @__PURE__ */ jsx("input", { value: horas, onChange: (e) => setHoras(e.target.value), type: "number", step: "0.5", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Em Vigor Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar Altera\xE7\xE3o" })
  ] });
}
function HistoricoHorasEquipamentoModal({ nomeEquipamento, historico, isAdmin, onDelete, onEdit, onClose }) {
  const ordenado = [...historico].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
  const [entradaAEditar, setEntradaAEditar] = useState(null);
  return /* @__PURE__ */ jsxs(Modal, { title: "Hist\xF3rico de Horas Di\xE1rias", subtitle: nomeEquipamento, onClose, children: [
    ordenado.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: "Ainda n\xE3o h\xE1 registos." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: ordenado.map((h, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== ordenado.length - 1 ? "border-b border-stone-100" : ""} ${i === 0 ? "bg-amber-50/40" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("p", { className: "font-mono-data font-semibold text-stone-800", children: [
          h.horasDiarias,
          " ",
          /* @__PURE__ */ jsx("span", { className: "font-sans font-normal text-stone-400 text-xs", children: "horas/dia" })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
          "Em vigor desde ",
          formatDatePT(h.dataEntradaVigor),
          " \xB7 registado ",
          formatDateTimePT(h.dataRegisto),
          " por ",
          h.utilizador || "\u2014"
        ] })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => setEntradaAEditar(h), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("button", { onClick: () => onDelete(h.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] })
    ] }, h.id)) }),
    entradaAEditar && /* @__PURE__ */ jsx(
      EditarHorasEquipamentoModal,
      {
        nomeEquipamento,
        entrada: entradaAEditar,
        onSave: (horasDiarias, dataEntradaVigor) => {
          onEdit(entradaAEditar.id, horasDiarias, dataEntradaVigor);
          setEntradaAEditar(null);
        },
        onClose: () => setEntradaAEditar(null)
      }
    )
  ] });
}
function MaoDeObraItensCard({ center, maoDeObra, isAdmin, nomeUtilizadorAtual, onUpdateItens }) {
  const itens = center.parametrizacao?.maoDeObraItens || [];
  const [novoArtigoId, setNovoArtigoId] = useState("");
  const [novaQuantidade, setNovaQuantidade] = useState("1");
  const [novaData, setNovaData] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [itemAtualizar, setItemAtualizar] = useState(null);
  const [itemHistorico, setItemHistorico] = useState(null);
  const nomeArtigo = (id) => maoDeObra.find((m) => m.id === id)?.designacao || "Categoria removida";
  const unidadeArtigo = (id) => maoDeObra.find((m) => m.id === id)?.unidadeCusto || "";
  const valorVigenteArtigo = (id) => {
    const artigo = maoDeObra.find((m) => m.id === id);
    if (!artigo) return null;
    return precoVigente(artigo);
  };
  const quantidadeVigente = (item) => {
    const hist = [...item.historico || []].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
    return hist[0] || null;
  };
  const adicionar = () => {
    if (!novoArtigoId) return;
    if (itens.some((i) => i.maoDeObraId === novoArtigoId)) return;
    const qtd = parseFloat(novaQuantidade) || 1;
    const entrada = { id: genId(), quantidade: qtd, dataEntradaVigor: novaData, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: nomeUtilizadorAtual || "" };
    onUpdateItens([...itens, { id: genId(), maoDeObraId: novoArtigoId, historico: [entrada] }]);
    setNovoArtigoId("");
    setNovaQuantidade("1");
    setNovaData((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  };
  const remover = (id) => onUpdateItens(itens.filter((i) => i.id !== id));
  return /* @__PURE__ */ jsxs("div", { className: "mt-2", children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5 mt-3", children: "Categorias de M\xE3o de Obra Afetas \xE0 Produ\xE7\xE3o Deste Centro" }),
    itens.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Ainda n\xE3o h\xE1 categorias de m\xE3o de obra atribu\xEDdas a este centro." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-3", children: itens.map((item, i) => {
      const vigente = valorVigenteArtigo(item.maoDeObraId);
      const qVigente = quantidadeVigente(item);
      return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== itens.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsx("p", { className: "text-stone-800 font-medium truncate", children: nomeArtigo(item.maoDeObraId) }),
          /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
            "Quantidade atual: ",
            /* @__PURE__ */ jsx("span", { className: "font-mono-data text-stone-600 font-semibold", children: qVigente ? qVigente.quantidade : "\u2014" }),
            qVigente?.dataEntradaVigor && /* @__PURE__ */ jsxs(Fragment, { children: [
              " (desde ",
              formatDatePT(qVigente.dataEntradaVigor),
              ")"
            ] }),
            vigente && /* @__PURE__ */ jsxs(Fragment, { children: [
              " \xB7 ",
              calcularPrecoFinal(vigente).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
              " ",
              unidadeArtigo(item.maoDeObraId)
            ] })
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
          /* @__PURE__ */ jsx("button", { onClick: () => setItemHistorico(item), title: "Ver hist\xF3rico de quantidade", className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(History, { size: 14 }) }),
          isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsx("button", { onClick: () => setItemAtualizar(item), title: "Atualizar quantidade", className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(TrendingUp, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => remover(item.id), title: "Remover categoria deste centro", className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] })
      ] }, item.id);
    }) }),
    isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-2", children: [
      /* @__PURE__ */ jsx("div", { className: "flex-1", children: /* @__PURE__ */ jsx(Field, { label: "Categoria", children: /* @__PURE__ */ jsxs("select", { value: novoArtigoId, onChange: (e) => setNovoArtigoId(e.target.value), className: inputCls, children: [
        /* @__PURE__ */ jsx("option", { value: "", children: "Escolha uma categoria de m\xE3o de obra..." }),
        maoDeObra.filter((m) => !itens.some((i) => i.maoDeObraId === m.id)).map((m) => /* @__PURE__ */ jsx("option", { value: m.id, children: m.designacao }, m.id))
      ] }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-24", children: /* @__PURE__ */ jsx(Field, { label: "Quantidade", children: /* @__PURE__ */ jsx("input", { value: novaQuantidade, onChange: (e) => setNovaQuantidade(e.target.value), type: "number", step: "1", min: "1", className: `${inputCls} font-mono-data` }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-40", children: /* @__PURE__ */ jsx(Field, { label: "Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: novaData, onChange: (e) => setNovaData(e.target.value), className: inputCls }) }) }),
      /* @__PURE__ */ jsx("button", { onClick: adicionar, className: "mb-4 px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700 shrink-0", children: "Adicionar" })
    ] }),
    itemAtualizar && /* @__PURE__ */ jsx(
      AtualizarQuantidadeMaoObraModal,
      {
        nomeCategoria: nomeArtigo(itemAtualizar.maoDeObraId),
        valorAtual: quantidadeVigente(itemAtualizar),
        onSave: (quantidade, dataEntradaVigor) => {
          const entrada = { id: genId(), quantidade, dataEntradaVigor, dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: nomeUtilizadorAtual || "" };
          onUpdateItens(itens.map((i) => i.id === itemAtualizar.id ? { ...i, historico: [...i.historico || [], entrada] } : i));
          setItemAtualizar(null);
        },
        onClose: () => setItemAtualizar(null)
      }
    ),
    itemHistorico && /* @__PURE__ */ jsx(
      HistoricoQuantidadeMaoObraModal,
      {
        nomeCategoria: nomeArtigo(itemHistorico.maoDeObraId),
        historico: itens.find((i) => i.id === itemHistorico.id)?.historico || [],
        isAdmin,
        onDelete: (entryId) => onUpdateItens(itens.map((i) => i.id === itemHistorico.id ? { ...i, historico: (i.historico || []).filter((h) => h.id !== entryId) } : i)),
        onEdit: (entryId, quantidade, dataEntradaVigor) => onUpdateItens(itens.map((i) => i.id === itemHistorico.id ? { ...i, historico: (i.historico || []).map((h) => h.id === entryId ? { ...h, quantidade, dataEntradaVigor } : h) } : i)),
        onClose: () => setItemHistorico(null)
      }
    )
  ] });
}
function AtualizarQuantidadeMaoObraModal({ nomeCategoria, valorAtual, onSave, onClose }) {
  const [quantidade, setQuantidade] = useState("");
  const [dataEntradaVigor, setDataEntradaVigor] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!quantidade || parseFloat(quantidade) < 0) return setError("Indique a nova quantidade");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    if (valorAtual && parseFloat(quantidade) === parseFloat(valorAtual.quantidade) && dataEntradaVigor === valorAtual.dataEntradaVigor) {
      return setError("Sem altera\xE7\xF5es em rela\xE7\xE3o ao valor j\xE1 em vigor");
    }
    onSave(parseFloat(quantidade), dataEntradaVigor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Quantidade", subtitle: nomeCategoria, onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Nova Quantidade", children: /* @__PURE__ */ jsx("input", { value: quantidade, onChange: (e) => setQuantidade(e.target.value), type: "number", step: "1", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Em Vigor Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EditarQuantidadeMaoObraModal({ nomeCategoria, entrada, onSave, onClose }) {
  const [quantidade, setQuantidade] = useState(String(entrada.quantidade ?? ""));
  const [dataEntradaVigor, setDataEntradaVigor] = useState(entrada.dataEntradaVigor || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!quantidade || parseFloat(quantidade) < 0) return setError("Indique a quantidade");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    onSave(parseFloat(quantidade), dataEntradaVigor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Registo de Quantidade", subtitle: nomeCategoria, onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Quantidade", children: /* @__PURE__ */ jsx("input", { value: quantidade, onChange: (e) => setQuantidade(e.target.value), type: "number", step: "1", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Em Vigor Desde", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar Altera\xE7\xE3o" })
  ] });
}
function HistoricoQuantidadeMaoObraModal({ nomeCategoria, historico, isAdmin, onDelete, onEdit, onClose }) {
  const ordenado = [...historico].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
  const [editando, setEditando] = useState(null);
  return /* @__PURE__ */ jsxs(Modal, { title: "Hist\xF3rico de Quantidade", subtitle: nomeCategoria, onClose, children: [
    ordenado.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: "Ainda n\xE3o h\xE1 registos." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: ordenado.map((h, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== ordenado.length - 1 ? "border-b border-stone-100" : ""} ${i === 0 ? "bg-amber-50/40" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("p", { className: "font-mono-data font-semibold text-stone-800", children: [
          h.quantidade,
          " ",
          /* @__PURE__ */ jsx("span", { className: "font-sans font-normal text-stone-400 text-xs", children: "elemento(s)" })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
          "Em vigor desde ",
          formatDatePT(h.dataEntradaVigor),
          " \xB7 registado ",
          formatDateTimePT(h.dataRegisto),
          " por ",
          h.utilizador || "\u2014"
        ] })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => setEditando(h), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("button", { onClick: () => onDelete(h.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] })
    ] }, h.id)) }),
    editando && /* @__PURE__ */ jsx(
      EditarQuantidadeMaoObraModal,
      {
        nomeCategoria,
        entrada: editando,
        onSave: (quantidade, dataEntradaVigor) => {
          onEdit(editando.id, quantidade, dataEntradaVigor);
          setEditando(null);
        },
        onClose: () => setEditando(null)
      }
    )
  ] });
}
function IncidenciasSection({ center, diarias, avarias, canManage, isAdmin, onBack, onAdd, onEdit, onDelete, onOpenDiaria, onOpenResolucao, onEditIncidenciaDiaria, onDeleteIncidenciaDiaria }) {
  const fmtDate = (iso) => iso ? formatDatePT(iso) : "\u2014";
  const daDiaria = diarias.flatMap((d) => normalizeIncidencias(d.incidencias).map((inc) => ({
    key: `diaria-${d.id}-${inc.id}`,
    diariaId: d.id,
    incidentId: inc.id,
    centroId: d.centroId,
    data: d.dataInicio,
    descricao: inc.descricao,
    origem: "Di\xE1ria",
    turno: d.turno,
    sourceDiaria: d,
    resolucaoData: inc.resolucaoData,
    resolucaoDescricao: inc.resolucaoDescricao
  })));
  const manuais = avarias.map((a) => ({
    key: a.id,
    id: a.id,
    centroId: a.centroId,
    data: a.data,
    descricao: a.descricao,
    origem: "Manual",
    raw: a,
    resolucaoData: a.resolucaoData,
    resolucaoDescricao: a.resolucaoDescricao
  }));
  const combinadas = [...daDiaria, ...manuais].sort((x, y) => (y.data || "").localeCompare(x.data || ""));
  const diasEntre = (a, b) => Math.round((new Date(b) - new Date(a)) / 864e5);
  const diasEmEspera = (dataReportada) => Math.max(0, diasEntre(dataReportada, (/* @__PURE__ */ new Date()).toISOString().slice(0, 10)));
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
          /* @__PURE__ */ jsx(AlertTriangle, { size: 17, className: "text-amber-600" }),
          " Incid\xEAncias / Avarias"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mt-1", children: "Alimentado automaticamente pelas di\xE1rias com incid\xEAncias, ou registado diretamente aqui." })
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700 shrink-0", children: [
        /* @__PURE__ */ jsx(Plus, { size: 16 }),
        " Nova Incid\xEAncia"
      ] })
    ] }),
    combinadas.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(AlertTriangle, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 incid\xEAncias ou avarias registadas neste centro." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "space-y-3", children: combinadas.map((item) => {
      const resolvida = !!item.resolucaoData;
      const dias = resolvida && item.data ? diasEntre(item.data, item.resolucaoData) : null;
      return /* @__PURE__ */ jsx("div", { className: "bg-white border border-stone-200 rounded-xl p-4", children: /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between gap-3", children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-1 flex-wrap", children: [
            /* @__PURE__ */ jsx("span", { className: "font-display font-medium text-stone-900", children: fmtDate(item.data) }),
            /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${item.origem === "Di\xE1ria" ? "bg-slate-700 text-slate-100" : "bg-amber-100 text-amber-700"}`, children: item.origem === "Di\xE1ria" ? `Parte Di\xE1ria${item.turno ? ` \xB7 ${item.turno}` : ""}` : "Introdu\xE7\xE3o Manual" }),
            /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${resolvida ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"}`, children: resolvida ? `Resolvido em ${dias} dia${dias !== 1 ? "s" : ""}` : `${diasEmEspera(item.data)} dia${diasEmEspera(item.data) !== 1 ? "s" : ""} \xE0 espera` })
          ] }),
          /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-600 whitespace-pre-wrap", children: item.descricao }),
          resolvida && /* @__PURE__ */ jsxs("div", { className: "mt-2 pt-2 border-t border-stone-100", children: [
            /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
              "Resolvido em ",
              fmtDate(item.resolucaoData)
            ] }),
            item.resolucaoDescricao && /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-600 whitespace-pre-wrap mt-0.5", children: item.resolucaoDescricao })
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
          /* @__PURE__ */ jsx(
            "button",
            {
              onClick: () => onOpenResolucao(item),
              title: resolvida ? "Editar resolu\xE7\xE3o" : "Registar resolu\xE7\xE3o",
              className: "p-1.5 text-stone-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg",
              children: /* @__PURE__ */ jsx(CheckCircle2, { size: 16 })
            }
          ),
          item.origem === "Di\xE1ria" ? /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsx("button", { onClick: () => onOpenDiaria(item.sourceDiaria), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", title: "Ver di\xE1ria", children: /* @__PURE__ */ jsx(ChevronRight, { size: 16 }) }),
            canManage && /* @__PURE__ */ jsxs(Fragment, { children: [
              /* @__PURE__ */ jsx(
                "button",
                {
                  onClick: () => onEditIncidenciaDiaria({ diariaId: item.diariaId, incidentId: item.incidentId, descricao: item.descricao }),
                  className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg",
                  title: "Editar descri\xE7\xE3o",
                  children: /* @__PURE__ */ jsx(Pencil, { size: 14 })
                }
              ),
              /* @__PURE__ */ jsx(
                "button",
                {
                  onClick: () => onDeleteIncidenciaDiaria(item.diariaId, item.incidentId),
                  className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg",
                  title: "Apagar incid\xEAncia",
                  children: /* @__PURE__ */ jsx(Trash2, { size: 14 })
                }
              )
            ] })
          ] }) : /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsx("button", { onClick: () => onEdit(item.raw), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
            canManage && /* @__PURE__ */ jsx("button", { onClick: () => onDelete(item.raw.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] })
      ] }) }, item.key);
    }) })
  ] });
}
function FormulasSection({ center, formulas, materiais, equipamentos, maoDeObra, consumiveis, logotipo, canManage, podeVerCustos, isAdmin, currentUserRole, onBack, onAdd, onEdit, onDelete, onImport, onDuplicate, onDeleteAll, onDeleteSelected, onToggleIncluirCustos, nomeUtilizadorAtual }) {
  const [query, setQuery] = useState("");
  const [fichaCustoFormula, setFichaCustoFormula] = useState(null);
  const [mostrarListaCustos, setMostrarListaCustos] = useState(false);
  const [selecionadas, setSelecionadas] = useState(/* @__PURE__ */ new Set());
  const fmtDate = (iso) => iso ? formatDatePT(iso) : "\u2014";
  const filtradas = query ? formulas.filter((f) => matchesSearch(query, f.codigo, f.designacao, f.estudo, f.central, f.observacoes)) : formulas;
  const sorted = [...filtradas].sort(
    (a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true, sensitivity: "base" }) || (a.designacao || "").localeCompare(b.designacao || "", "pt", { sensitivity: "base" })
  );
  const todasVisiveisSelecionadas = sorted.length > 0 && sorted.every((f) => selecionadas.has(f.id));
  const toggleTodas = () => {
    setSelecionadas((prev) => {
      if (todasVisiveisSelecionadas) {
        const novo2 = new Set(prev);
        sorted.forEach((f) => novo2.delete(f.id));
        return novo2;
      }
      const novo = new Set(prev);
      sorted.forEach((f) => novo.add(f.id));
      return novo;
    });
  };
  const toggleUma = (id) => setSelecionadas((prev) => {
    const novo = new Set(prev);
    if (novo.has(id)) novo.delete(id);
    else novo.add(id);
    return novo;
  });
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-5xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
          /* @__PURE__ */ jsx(FlaskConical, { size: 17, className: "text-amber-600" }),
          " Lista de F\xF3rmulas"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mt-1", children: "F\xF3rmulas de trabalho e dosifica\xE7\xE3o de agregados a frio deste centro." })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2 shrink-0", children: [
        isAdmin && (selecionadas.size > 0 ? /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => {
              onDeleteSelected([...selecionadas]);
              setSelecionadas(/* @__PURE__ */ new Set());
            },
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-red-600 text-white rounded-lg text-sm font-semibold hover:bg-red-700",
            children: [
              /* @__PURE__ */ jsx(Trash2, { size: 15 }),
              " Eliminar Selecionadas (",
              selecionadas.size,
              ")"
            ]
          }
        ) : formulas.length > 0 && /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => onDeleteAll(center.id),
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-red-200 text-red-600 rounded-lg text-sm font-semibold hover:bg-red-50",
            children: [
              /* @__PURE__ */ jsx(Trash2, { size: 15 }),
              " Eliminar Tudo"
            ]
          }
        )),
        podeVerCustos && /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarListaCustos(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(Calculator, { size: 15 }),
          " Custos de Todas as F\xF3rmulas"
        ] }),
        canManage && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 15 }),
          " Nova F\xF3rmula"
        ] })
      ] })
    ] }),
    mostrarListaCustos && /* @__PURE__ */ jsx(
      ListaCustosFormulasModal,
      {
        center,
        formulas,
        materiais,
        equipamentos,
        maoDeObra,
        consumiveis,
        logotipo,
        nomeUtilizadorAtual,
        onOpenFormula: (f) => {
          setMostrarListaCustos(false);
          onEdit(f);
        },
        onClose: () => setMostrarListaCustos(false)
      }
    ),
    /* @__PURE__ */ jsxs("div", { className: "relative mb-4 mt-4", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          placeholder: "Pesquisar por c\xF3digo, designa\xE7\xE3o, estudo ou n\xFAmero...",
          className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
        }
      ),
      query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
    ] }),
    formulas.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(FlaskConical, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 f\xF3rmulas registadas neste centro." })
    ] }) : sorted.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhuma f\xF3rmula encontrada para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: [
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between gap-2 px-5 py-2 border-b border-stone-100 bg-stone-50/60", children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
          /* @__PURE__ */ jsx("input", { type: "checkbox", checked: todasVisiveisSelecionadas, onChange: toggleTodas, className: "w-4 h-4 accent-amber-600 cursor-pointer shrink-0" }),
          /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-500", children: "Selecionar todas as vis\xEDveis" })
        ] }),
        /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-400", children: 'A caixa \xE0 direita escolhe se a f\xF3rmula aparece em "Custos de Todas as F\xF3rmulas"' })
      ] }),
      sorted.map((f, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-3.5 hover:bg-stone-50 ${i !== sorted.length - 1 ? "border-b border-stone-100" : ""} ${selecionadas.has(f.id) ? "bg-amber-50/50" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0 flex-1", children: [
          isAdmin && /* @__PURE__ */ jsx("input", { type: "checkbox", checked: selecionadas.has(f.id), onChange: () => toggleUma(f.id), className: "w-4 h-4 accent-amber-600 cursor-pointer shrink-0" }),
          currentUserRole === "Or\xE7amentista" ? /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0 flex-1", children: [
            /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-1 rounded shrink-0", children: f.codigo || "\u2014" }),
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: f.designacao || "\u2014" }),
              /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
                f.estudo && /* @__PURE__ */ jsxs(Fragment, { children: [
                  "Estudo ",
                  f.estudo,
                  " \xB7 "
                ] }),
                "Alterado em ",
                fmtDate(f.dataAlteracao)
              ] }),
              f.observacoes && /* @__PURE__ */ jsxs("p", { className: "text-xs text-amber-700 truncate mt-0.5", children: [
                "Obs: ",
                f.observacoes
              ] })
            ] })
          ] }) : /* @__PURE__ */ jsxs("button", { onClick: () => onEdit(f), className: "flex items-center gap-3 min-w-0 text-left flex-1", children: [
            /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-2 py-1 rounded shrink-0", children: f.codigo || "\u2014" }),
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: f.designacao || "\u2014" }),
              /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
                f.estudo && /* @__PURE__ */ jsxs(Fragment, { children: [
                  "Estudo ",
                  f.estudo,
                  " \xB7 "
                ] }),
                "Alterado em ",
                fmtDate(f.dataAlteracao)
              ] }),
              f.observacoes && /* @__PURE__ */ jsxs("p", { className: "text-xs text-amber-700 truncate mt-0.5", children: [
                "Obs: ",
                f.observacoes
              ] })
            ] })
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
          isAdmin && /* @__PURE__ */ jsx(
            "input",
            {
              type: "checkbox",
              checked: f.incluirEmCustosTodas !== false,
              onChange: () => onToggleIncluirCustos(f.id),
              title: 'Incluir em "Custos de Todas as F\xF3rmulas"',
              className: "w-4 h-4 accent-amber-600 cursor-pointer mr-1"
            }
          ),
          podeVerCustos && /* @__PURE__ */ jsx(
            "span",
            {
              onClick: () => setFichaCustoFormula(f),
              title: "Ficha de Custo",
              className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg cursor-pointer",
              children: /* @__PURE__ */ jsx(Calculator, { size: 14 })
            }
          ),
          isAdmin && /* @__PURE__ */ jsx(
            "span",
            {
              onClick: () => onDuplicate(f),
              title: "Duplicar f\xF3rmula",
              className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg cursor-pointer",
              children: /* @__PURE__ */ jsx(Copy, { size: 14 })
            }
          ),
          isAdmin && /* @__PURE__ */ jsx(
            "span",
            {
              onClick: () => onDelete(f.id),
              className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg cursor-pointer",
              children: /* @__PURE__ */ jsx(Trash2, { size: 14 })
            }
          )
        ] })
      ] }, f.id))
    ] }),
    fichaCustoFormula && /* @__PURE__ */ jsx(FichaCustoModal, { formula: fichaCustoFormula, center, materiais: materiais || [], equipamentos: equipamentos || [], maoDeObra: maoDeObra || [], consumiveis: consumiveis || [], logotipo, onClose: () => setFichaCustoFormula(null) })
  ] });
}
function ExportarProducaoMensalModal({ center, diarias, clientes, centrosCusto, formulas, materiais, equipamentos, maoDeObra, consumiveis, logotipo, onClose }) {
  const [mesAno, setMesAno] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 7));
  const [clienteId, setClienteId] = useState("");
  const nomeCliente = (id) => clientes.find((c) => c.id === id)?.designacao || "\u2014";
  const nomeObra = (id) => {
    const cc = centrosCusto.find((x) => x.id === id);
    return cc ? `${cc.codigo ? cc.codigo + " - " : ""}${cc.designacao}` : "";
  };
  const linhasDoMesSemFiltroCliente = (diarias || []).filter((d) => d.centroId === center.id && (d.dataInicio || "").slice(0, 7) === mesAno).flatMap((d) => (d.linhas || []).map((l) => l.clienteId));
  const clientesComProducao = clientes.filter((c) => linhasDoMesSemFiltroCliente.includes(c.id)).sort((a, b) => (a.designacao || "").localeCompare(b.designacao || "", "pt", { sensitivity: "base" }));
  const linhas = (diarias || []).filter((d) => d.centroId === center.id && (d.dataInicio || "").slice(0, 7) === mesAno).flatMap((d) => (d.linhas || []).map((l) => ({ ...l, data: d.dataInicio, diariaId: d.id }))).filter((l) => !clienteId || l.clienteId === clienteId);
  const linhasComCusto = linhas.map((l) => {
    const formula = (formulas || []).find((f) => f.id === l.artigoId);
    const toneladas = parseFloat(l.toneladas) || 0;
    const misturaLabel = formula?.codigo ? `${formula.codigo} - ${l.artigoDesignacao || formula.designacao || ""}` : l.artigoDesignacao || "\u2014";
    if (!formula) return { ...l, toneladas, misturaLabel, custoUnitario: null, custoTotal: null, semFormula: true };
    const { totalGeral } = calcularCustoFormula(formula, center, materiais || [], equipamentos || [], maoDeObra || [], l.data, consumiveis || []);
    return { ...l, toneladas, misturaLabel, custoUnitario: totalGeral, custoTotal: totalGeral * toneladas };
  }).sort((a, b) => (a.data || "").localeCompare(b.data || ""));
  const totalToneladas = linhasComCusto.reduce((s, l) => s + l.toneladas, 0);
  const totalCusto = linhasComCusto.reduce((s, l) => s + (l.custoTotal || 0), 0);
  const temSemFormula = linhasComCusto.some((l) => l.semFormula);
  const [ano, mesNum] = mesAno.split("-");
  const nomeMes = (/* @__PURE__ */ new Date(`${mesAno}-01T00:00:00Z`)).toLocaleDateString("pt-PT", { month: "long", year: "numeric", timeZone: "UTC" });
  const exportar = () => {
    const linhasHtml = linhasComCusto.map((l) => `
      <tr>
        <td>${formatDatePT(l.data)}</td>
        <td>${nomeCliente(l.clienteId)}</td>
        <td>${nomeObra(l.centroCustoId) || "\u2014"}</td>
        <td>${l.misturaLabel}</td>
        <td style="text-align:right">${l.toneladas.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} t</td>
        <td style="text-align:right">${l.custoUnitario !== null ? l.custoUnitario.toLocaleString("pt-PT", { maximumFractionDigits: 2 }) + " \u20AC/t" : "\u2014"}</td>
        <td style="text-align:right; font-weight:bold">${l.custoTotal !== null ? l.custoTotal.toLocaleString("pt-PT", { maximumFractionDigits: 2 }) + " \u20AC" : "\u2014"}</td>
      </tr>`).join("");
    const html = `<!DOCTYPE html>
      <html><head><meta charset="utf-8"><title>Produ\xE7\xE3o Mensal \u2014 ${center?.nome || ""}</title>
      <style>
        body { font-family: 'Arial Narrow', Arial, sans-serif; color: #1c1917; padding: 32px; max-width: 1000px; margin: 0 auto; font-size: 13px; }
        .logo { max-height: 55px; margin-bottom: 16px; }
        h1 { font-size: 18px; margin-bottom: 2px; text-transform: capitalize; }
        .sub { color: #78716c; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
        th, td { border: 1px solid #d6d3d1; padding: 6px 8px; }
        th { background: #f5f5f4; text-align: left; font-size: 10px; text-transform: uppercase; }
        .totais { margin-top: 16px; padding: 12px 16px; background: #fef3c7; border-radius: 8px; display: flex; justify-content: space-between; font-weight: bold; }
      </style></head>
      <body>
        ${logotipo ? `<img class="logo" src="${logotipo}" alt="Log\xF3tipo" />` : ""}
        <h1>Produ\xE7\xE3o Mensal${clienteId ? ` \u2014 ${nomeCliente(clienteId)}` : ""}</h1>
        <p class="sub">${center?.nome || ""} \u2014 ${nomeMes}</p>
        <table>
          <tr><th>Data</th><th>Cliente</th><th>Obra</th><th>Mistura</th><th style="text-align:right">Toneladas</th><th style="text-align:right">Custo Unit.</th><th style="text-align:right">Custo Total</th></tr>
          ${linhasHtml || '<tr><td colspan="7" style="text-align:center">Sem produ\xE7\xE3o neste per\xEDodo</td></tr>'}
        </table>
        <div class="totais">
          <span>Total: ${totalToneladas.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} t</span>
          <span>Custo Total: ${totalCusto.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC</span>
        </div>
      </body></html>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Producao_Mensal_${mesAno}_${center?.nome || "centro"}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Exportar Produ\xE7\xE3o Mensal", subtitle: center?.nome || "", onClose, wide: true, children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-3 mb-5", children: [
      /* @__PURE__ */ jsx("div", { className: "w-44", children: /* @__PURE__ */ jsx(Field, { label: "M\xEAs", children: /* @__PURE__ */ jsx("input", { type: "month", value: mesAno, onChange: (e) => setMesAno(e.target.value), className: inputCls }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-72", children: /* @__PURE__ */ jsx(Field, { label: "Cliente", children: /* @__PURE__ */ jsx(ClienteSearchSelect, { value: clienteId, clientes: clientesComProducao, permitirVazio: true, placeholder: "Todos os Clientes", onChange: setClienteId }) }) }),
      /* @__PURE__ */ jsxs("button", { onClick: exportar, className: "mb-4 flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileText, { size: 15 }),
        " Exportar PDF"
      ] })
    ] }),
    temSemFormula && /* @__PURE__ */ jsx("p", { className: "text-xs text-red-600 mb-3", children: "H\xE1 linhas de produ\xE7\xE3o associadas a uma f\xF3rmula que j\xE1 n\xE3o existe \u2014 n\xE3o \xE9 poss\xEDvel calcular o custo dessas linhas." }),
    linhasComCusto.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-400 py-6", children: [
      "Sem produ\xE7\xE3o registada neste centro para ",
      nomeMes,
      clienteId ? ` e cliente "${nomeCliente(clienteId)}"` : "",
      "."
    ] }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Data" }),
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Cliente" }),
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Obra" }),
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Mistura" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Toneladas" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo Unit." }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo Total" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasComCusto.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasComCusto.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-500", children: formatDatePT(l.data) }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-700", children: nomeCliente(l.clienteId) }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-500 truncate max-w-[140px]", children: nomeObra(l.centroCustoId) || "\u2014" }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-800", children: l.misturaLabel }),
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-right font-mono-data", children: [
          l.toneladas.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " t"
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data text-stone-500", children: l.custoUnitario !== null ? `${l.custoUnitario.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC/t` : "\u2014" }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold text-stone-800", children: l.custoTotal !== null ? `${l.custoTotal.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, `${l.diariaId}-${l.id}-${i}`)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-3 mt-4", children: [
      /* @__PURE__ */ jsxs("span", { className: "text-sm font-semibold text-amber-800", children: [
        "Total: ",
        totalToneladas.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " t"
      ] }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-lg text-amber-800", children: [
        totalCusto.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " \u20AC"
      ] })
    ] })
  ] });
}
function GraficoProducaoModal({ diarias, center, onClose }) {
  const anos = [...new Set(diarias.filter((d) => d.dataInicio).map((d) => new Date(d.dataInicio).getFullYear()))].sort((a, b) => b - a);
  const anoAtual = (/* @__PURE__ */ new Date()).getFullYear();
  const [ano, setAno] = useState(anos.includes(anoAtual) ? anoAtual : anos[0] || anoAtual);
  const diaDoAno = (isoDate) => Math.round((new Date(isoDate) - /* @__PURE__ */ new Date(`${ano}-01-01`)) / 864e5) + 1;
  const bissexto = ano % 4 === 0 && ano % 100 !== 0 || ano % 400 === 0;
  const diasNoAno = bissexto ? 366 : 365;
  const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  const ticksMeses = MESES.map((_, i) => diaDoAno(new Date(ano, i, 1).toISOString().slice(0, 10)));
  const porDia = {};
  diarias.filter((d) => d.dataInicio && new Date(d.dataInicio).getFullYear() === ano).forEach((d) => {
    const total = (d.linhas || []).reduce((s, l) => s + (parseFloat(l.toneladas) || 0), 0);
    porDia[d.dataInicio] = (porDia[d.dataInicio] || 0) + total;
  });
  let acumulado = 0;
  const chartData = Object.keys(porDia).sort().map((data) => {
    acumulado += porDia[data];
    return { dia: diaDoAno(data), dataLabel: formatDatePT(data), diaria: Math.round(porDia[data] * 100) / 100, acumulado: Math.round(acumulado * 100) / 100 };
  });
  const totalAno = acumulado;
  return /* @__PURE__ */ jsxs(Modal, { title: "Produ\xE7\xE3o do Ano", subtitle: `${center?.nome || "Todos os Centros"} \u2014 Di\xE1ria (barras) e Acumulada (linha)`, onClose, wide: true, children: [
    anos.length > 1 && /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Ano:" }),
      anos.map((a) => /* @__PURE__ */ jsx(
        "button",
        {
          onClick: () => setAno(a),
          className: `px-3 py-1.5 rounded-lg text-sm font-semibold ${a === ano ? "bg-amber-600 text-white" : "bg-stone-100 text-stone-600 hover:bg-stone-200"}`,
          children: a
        },
        a
      ))
    ] }),
    chartData.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 py-8 text-center", children: [
      "Ainda n\xE3o h\xE1 di\xE1rias registadas em ",
      ano,
      "."
    ] }) : /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "bg-amber-50 rounded-lg px-4 py-2.5 mb-4 flex items-center justify-between", children: [
        /* @__PURE__ */ jsxs("span", { className: "text-sm font-semibold text-amber-800", children: [
          "Total produzido em ",
          ano
        ] }),
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-amber-800", children: [
          totalAno.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " t"
        ] })
      ] }),
      /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 340, children: /* @__PURE__ */ jsxs(ComposedChart, { data: chartData, margin: { top: 8, right: 16, bottom: 0, left: 0 }, children: [
        /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", stroke: "#e7e5e4" }),
        /* @__PURE__ */ jsx(
          XAxis,
          {
            dataKey: "dia",
            type: "number",
            domain: [1, diasNoAno],
            ticks: ticksMeses,
            tickFormatter: (v) => MESES[ticksMeses.indexOf(v)] || "",
            tick: { fontSize: 11, fill: "#78716c" },
            tickLine: false,
            axisLine: { stroke: "#e7e5e4" }
          }
        ),
        /* @__PURE__ */ jsx(
          YAxis,
          {
            yAxisId: "left",
            tick: { fontSize: 11, fill: "#78716c" },
            tickLine: false,
            axisLine: false,
            width: 65,
            tickFormatter: (v) => `${v.toLocaleString("pt-PT")} t`,
            label: { value: "Prod Di\xE1ria", angle: -90, position: "insideLeft", style: { textAnchor: "middle", fontSize: 11, fill: "#78716c" } }
          }
        ),
        /* @__PURE__ */ jsx(
          YAxis,
          {
            yAxisId: "right",
            orientation: "right",
            tick: { fontSize: 11, fill: "#78716c" },
            tickLine: false,
            axisLine: false,
            width: 75,
            tickFormatter: (v) => `${v.toLocaleString("pt-PT")} t`,
            label: { value: "Prod Acumulada Anual", angle: -90, position: "insideRight", style: { textAnchor: "middle", fontSize: 11, fill: "#1d4ed8" } }
          }
        ),
        /* @__PURE__ */ jsx(
          Tooltip,
          {
            labelFormatter: (v, payload) => payload?.[0]?.payload?.dataLabel || "",
            formatter: (v, name) => [`${v.toLocaleString("pt-PT")} t`, name === "diaria" ? "Produ\xE7\xE3o do dia" : "Acumulado do ano"],
            contentStyle: { fontSize: 12, borderRadius: 8, borderColor: "#e7e5e4" }
          }
        ),
        /* @__PURE__ */ jsx(Legend, { formatter: (value) => value === "diaria" ? "Produ\xE7\xE3o di\xE1ria" : "Acumulado do ano", wrapperStyle: { fontSize: 12 } }),
        /* @__PURE__ */ jsx(Bar, { yAxisId: "left", dataKey: "diaria", fill: "#fbbf24", radius: [3, 3, 0, 0], barSize: 4 }),
        /* @__PURE__ */ jsx(Line, { yAxisId: "right", type: "monotone", dataKey: "acumulado", stroke: "#1d4ed8", strokeWidth: 2.5, dot: false })
      ] }) })
    ] })
  ] });
}
function ProducaoSection({ center, diarias, avarias, clientes, centrosCusto, canManage, podeRegistar, isAdmin, onBack, onAdd, onEdit, onDelete, onImport, onOpenDiaria, onOpenResolucao, formulas, materiais, equipamentos, maoDeObra, consumiveis, logotipo }) {
  const [searchDate, setSearchDate] = useState("");
  const [mostrarGrafico, setMostrarGrafico] = useState(false);
  const [mostrarExportarMensal, setMostrarExportarMensal] = useState(false);
  const totalDiariaCalc = (d) => (d.linhas || []).reduce((s, l) => s + (parseFloat(l.toneladas) || 0), 0);
  const turnoDe = (d) => d.turno || (d.dataFim === d.dataInicio ? "Diurno" : d.dataFim > d.dataInicio ? "Noturno" : "");
  const totalDiaria = totalDiariaCalc;
  const fmt = (n) => n.toLocaleString("pt-PT", { maximumFractionDigits: 2 });
  const fmtDate = (iso) => iso ? formatDatePT(iso) : "\u2014";
  const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  const diasAberto = (dataInicio) => Math.max(0, Math.round((new Date(hoje) - new Date(dataInicio)) / 864e5));
  const pendentes = [
    ...diarias.flatMap((d) => normalizeIncidencias(d.incidencias).filter((inc) => !inc.resolucaoData).map((inc) => ({
      key: `diaria-${d.id}-${inc.id}`,
      diariaId: d.id,
      incidentId: inc.id,
      data: d.dataInicio,
      descricao: inc.descricao,
      origem: "Di\xE1ria",
      turno: d.turno,
      sourceDiaria: d
    }))),
    ...(avarias || []).filter((a) => !a.resolucaoData).map((a) => ({ key: a.id, id: a.id, data: a.data, descricao: a.descricao, origem: "Manual", raw: a }))
  ].sort((x, y) => (x.data || "").localeCompare(y.data || ""));
  const dataOrdenacaoDiaria = (d) => {
    const tIni = d.dataInicio ? (/* @__PURE__ */ new Date(`${d.dataInicio}T00:00:00Z`)).getTime() : NaN;
    if (isNaN(tIni)) return 0;
    if (!d.dataFim || d.dataFim === d.dataInicio) return tIni;
    const tFim = (/* @__PURE__ */ new Date(`${d.dataFim}T00:00:00Z`)).getTime();
    if (isNaN(tFim)) return tIni;
    return (tIni + tFim) / 2;
  };
  const ULTIMAS = 30;
  const todasOrdenadas = [...diarias].sort((a, b) => dataOrdenacaoDiaria(b) - dataOrdenacaoDiaria(a));
  const sorted = searchDate ? todasOrdenadas.filter((d) => d.dataInicio === searchDate) : todasOrdenadas.slice(0, ULTIMAS);
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-5xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    pendentes.length > 0 && /* @__PURE__ */ jsxs("div", { className: "mb-6", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-sm text-stone-700 font-semibold flex items-center gap-2 mb-2", children: [
        /* @__PURE__ */ jsx(AlertTriangle, { size: 15, className: "text-red-500" }),
        " Pendentes (",
        pendentes.length,
        ")"
      ] }),
      /* @__PURE__ */ jsx("div", { className: "bg-red-50/60 border border-red-100 rounded-xl overflow-hidden", children: pendentes.map((item, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-4 py-2.5 ${i !== pendentes.length - 1 ? "border-b border-red-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0 flex items-center gap-2", children: [
          /* @__PURE__ */ jsx("span", { className: "text-xs font-mono-data text-stone-500 shrink-0", children: fmtDate(item.data) }),
          /* @__PURE__ */ jsx("span", { className: "text-sm text-stone-700 truncate", children: item.descricao })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 shrink-0", children: [
          /* @__PURE__ */ jsxs("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-red-100 text-red-700", children: [
            diasAberto(item.data),
            " dia",
            diasAberto(item.data) !== 1 ? "s" : "",
            " em aberto"
          ] }),
          /* @__PURE__ */ jsx("button", { onClick: () => onOpenResolucao(item), title: "Registar resolu\xE7\xE3o", className: "p-1 text-stone-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg", children: /* @__PURE__ */ jsx(CheckCircle2, { size: 14 }) }),
          item.origem === "Di\xE1ria" && /* @__PURE__ */ jsx("button", { onClick: () => onOpenDiaria(item.sourceDiaria), title: "Ver di\xE1ria", className: "p-1 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(ChevronRight, { size: 14 }) })
        ] })
      ] }, item.key)) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
        /* @__PURE__ */ jsx(ClipboardList, { size: 17, className: "text-amber-600" }),
        " Di\xE1rias de Produ\xE7\xE3o"
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onImport, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: () => {
          const nomeCliente = (id) => clientes.find((c) => c.id === id)?.designacao || "";
          const nomeObra = (id) => centrosCusto.find((x) => x.id === id)?.designacao || "";
          const nomeMistura = (artigoId) => (formulas || []).find((f) => f.id === artigoId)?.designacao || "";
          const linhas = [];
          [...diarias].sort((a, b) => dataOrdenacaoDiaria(b) - dataOrdenacaoDiaria(a)).forEach((d) => {
            const producaoLinhas = d.linhas || [];
            if (producaoLinhas.length === 0) {
              linhas.push([d.dataInicio || "", d.dataFim || "", turnoDe(d), "", "", "", "", d.observacoes || "", d.criadoPor || ""]);
              return;
            }
            producaoLinhas.forEach((l) => {
              linhas.push([
                d.dataInicio || "",
                d.dataFim || "",
                turnoDe(d),
                nomeCliente(l.clienteId),
                nomeObra(l.centroCustoId),
                l.artigoDesignacao || nomeMistura(l.artigoId),
                l.toneladas || "",
                d.observacoes || "",
                d.criadoPor || ""
              ]);
            });
          });
          exportarListaExcel(
            "diarias_producao.xlsx",
            "Di\xE1rias",
            ["Data In\xEDcio", "Data Fim", "Turno", "Cliente", "Obra", "Artigo", "Toneladas", "Observa\xE7\xF5es", "Criado Por"],
            linhas
          );
        }, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Exportar Excel"
        ] }),
        /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarGrafico(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(TrendingUp, { size: 15 }),
          " Ver Gr\xE1fico"
        ] }),
        canManage && /* @__PURE__ */ jsxs("button", { onClick: () => setMostrarExportarMensal(true), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileText, { size: 15 }),
          " Exportar Produ\xE7\xE3o Mensal"
        ] }),
        podeRegistar && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 15 }),
          " Nova Di\xE1ria"
        ] })
      ] })
    ] }),
    mostrarGrafico && /* @__PURE__ */ jsx(GraficoProducaoModal, { diarias, center, onClose: () => setMostrarGrafico(false) }),
    mostrarExportarMensal && /* @__PURE__ */ jsx(
      ExportarProducaoMensalModal,
      {
        center,
        diarias,
        clientes,
        centrosCusto,
        formulas,
        materiais,
        equipamentos,
        maoDeObra,
        consumiveis,
        logotipo,
        onClose: () => setMostrarExportarMensal(false)
      }
    ),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "relative flex-1 max-w-xs", children: [
        /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 15 }),
        /* @__PURE__ */ jsx(
          "input",
          {
            type: "date",
            value: searchDate,
            onChange: (e) => setSearchDate(e.target.value),
            className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
          }
        ),
        searchDate && /* @__PURE__ */ jsx("button", { onClick: () => setSearchDate(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
      ] }),
      !searchDate && /* @__PURE__ */ jsxs("span", { className: "text-xs text-stone-400", children: [
        "A mostrar as \xFAltimas ",
        sorted.length,
        " di\xE1ria",
        sorted.length !== 1 ? "s" : "",
        " de ",
        todasOrdenadas.length
      ] })
    ] }),
    todasOrdenadas.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(ClipboardList, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 di\xE1rias de produ\xE7\xE3o registadas neste centro." })
    ] }) : sorted.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        "Nenhuma di\xE1ria encontrada com data de in\xEDcio ",
        searchDate,
        "."
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: sorted.map((d, i) => {
      const turno = turnoDe(d);
      return /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => onEdit(d),
          className: `w-full flex items-center justify-between px-5 py-3.5 text-left hover:bg-stone-50 ${i !== sorted.length - 1 ? "border-b border-stone-100" : ""}`,
          children: [
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
              /* @__PURE__ */ jsxs("span", { className: "font-display font-medium text-stone-900 truncate", children: [
                d.dataInicio,
                d.dataFim && d.dataFim !== d.dataInicio ? ` \u2192 ${d.dataFim}` : ""
              ] }),
              turno && /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${turno === "Diurno" ? "bg-amber-100 text-amber-700" : "bg-slate-700 text-slate-100"}`, children: turno }),
              normalizeIncidencias(d.incidencias).length > 0 && /* @__PURE__ */ jsx("span", { title: "Tem incid\xEAncias/avarias registadas", className: "shrink-0", children: /* @__PURE__ */ jsx(AlertTriangle, { size: 14, className: "text-red-500" }) })
            ] }),
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 shrink-0", children: [
              /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-sm text-amber-700 font-semibold", children: [
                fmt(totalDiaria(d)),
                " t"
              ] }),
              canManage && /* @__PURE__ */ jsx(
                "span",
                {
                  onClick: (e) => {
                    e.stopPropagation();
                    onDelete(d.id);
                  },
                  className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg",
                  children: /* @__PURE__ */ jsx(Trash2, { size: 14 })
                }
              )
            ] })
          ]
        },
        d.id
      );
    }) })
  ] });
}
function RececaoSection({ center, categoria, titulo, produtos, rececoes, canManage, podeRegistar, onBack, onAdd, onEdit, onDelete }) {
  const sorted = [...rececoes].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  const nomeProduto = (id) => produtos.find((p) => p.id === id)?.designacao || "\u2014";
  const fornecedorProduto = (id) => produtos.find((p) => p.id === id)?.fornecedor || "";
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
        /* @__PURE__ */ jsx(Droplet, { size: 17, className: "text-amber-600" }),
        " ",
        titulo
      ] }),
      podeRegistar && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
        /* @__PURE__ */ jsx(Plus, { size: 15 }),
        " Nova Rece\xE7\xE3o"
      ] })
    ] }),
    sorted.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Droplet, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 rece\xE7\xF5es registadas neste centro." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: sorted.map((r, i) => /* @__PURE__ */ jsxs(
      "button",
      {
        onClick: () => canManage && onEdit(r),
        className: `w-full flex items-center justify-between px-5 py-3.5 text-left ${canManage ? "hover:bg-stone-50" : "cursor-default"} ${i !== sorted.length - 1 ? "border-b border-stone-100" : ""}`,
        children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
            /* @__PURE__ */ jsx("span", { className: "font-display font-medium text-stone-900 shrink-0", children: formatDatePT(r.data) }),
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-700 truncate", children: nomeProduto(r.produtoId) }),
              fornecedorProduto(r.produtoId) && /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 truncate", children: fornecedorProduto(r.produtoId) }),
              r.substituiId && /* @__PURE__ */ jsxs("p", { className: "text-[11px] text-amber-600 truncate", children: [
                "substitui ",
                nomeProduto(r.substituiId)
              ] })
            ] })
          ] }),
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 shrink-0", children: [
            /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-sm text-amber-700 font-semibold", children: [
              parseFloat(r.quantidade || 0).toLocaleString("pt-PT"),
              " ",
              categoria === "consumiveis" ? "L" : "t"
            ] }),
            canManage && /* @__PURE__ */ jsx("span", { onClick: (e) => {
              e.stopPropagation();
              onDelete(r.id);
            }, className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ]
      },
      r.id
    )) })
  ] });
}
function RececaoModal({ data, produtos, onSave, onClose }) {
  const [dataRececao, setDataRececao] = useState(data?.data || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [produtoId, setProdutoId] = useState(data?.produtoId || "");
  const [quantidade, setQuantidade] = useState(data?.quantidade ?? "");
  const [substituiId, setSubstituiId] = useState(data?.substituiId || "");
  const [error, setError] = useState("");
  const unidade = data.categoria === "consumiveis" ? "litros" : "toneladas";
  const produtoSubstituido = produtos.find((p) => p.id === substituiId);
  const submit = () => {
    if (!dataRececao) return setError("Indique a data");
    if (!produtoId) return setError("Escolha o produto");
    if (!quantidade || parseFloat(quantidade) <= 0) return setError(`Indique a quantidade em ${unidade}`);
    if (substituiId === produtoId) return setError("O produto recebido n\xE3o pode substituir-se a si pr\xF3prio");
    onSave({ id: data?.id, centroId: data.centroId, categoria: data.categoria, data: dataRececao, produtoId, quantidade: parseFloat(quantidade), substituiId: substituiId || null });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Rece\xE7\xE3o" : "Nova Rece\xE7\xE3o", subtitle: { betumes: "Rece\xE7\xE3o de Betumes", agregados: "Rece\xE7\xE3o de Agregados", filler: "Rece\xE7\xE3o de Filler Comercial", consumiveis: "Rece\xE7\xE3o de Combust\xEDveis" }[data.categoria] || "Rece\xE7\xE3o", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Data", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataRececao, onChange: (e) => setDataRececao(e.target.value), className: inputCls, autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Produto", children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: produtoId, materiais: produtos, onChange: setProdutoId }) }),
    /* @__PURE__ */ jsx(Field, { label: `Quantidade (${unidade})`, children: /* @__PURE__ */ jsx("input", { value: quantidade, onChange: (e) => setQuantidade(e.target.value), type: "number", step: "0.01", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0.00" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Substitui (opcional)", children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: substituiId, materiais: produtos.filter((p) => p.id !== produtoId), onChange: setSubstituiId }) }),
    /* @__PURE__ */ jsxs("p", { className: "text-[11px] text-stone-400 -mt-2 mb-4", children: [
      "Preencha s\xF3 se este material est\xE1 a ser recebido como alternativa a outro (ex: por escassez). A entrada em stock e o consumo passam a contar para ",
      /* @__PURE__ */ jsx("strong", { children: produtoSubstituido ? produtoSubstituido.designacao : "o produto que estava previsto na f\xF3rmula" }),
      ", mantendo o registo de que foi entregue um material diferente."
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function StocksSection({ center, materiais, consumiveis, rececoes, diarias, formulas, ajustesStock, isAdmin, onOpenHistoricoStock, onBack }) {
  const rececoesDoCentro = rececoes.filter((r) => r.centroId === center.id);
  const blocoTermico = center.parametrizacao?.blocoTermico || {};
  const formulasPorId = Object.fromEntries((formulas || []).map((f) => [f.id, f]));
  const consumoProducao = {};
  (diarias || []).forEach((d) => {
    if (d.centroId !== center.id) return;
    (d.linhas || []).forEach((l) => {
      const formula = formulasPorId[l.artigoId];
      const toneladas = parseFloat(l.toneladas) || 0;
      if (!formula || toneladas <= 0) return;
      const kgNaoAgregadoPorTonelada = TRABALHO_MATERIAL_KEYS.reduce((s, key) => s + (parseFloat(formula.trabalho?.[key]?.design) || 0), 0);
      const kgAgregadoPorTonelada = Math.max(0, 1e3 - kgNaoAgregadoPorTonelada);
      SILO_COLS.forEach((c) => {
        const silo = formula.silos?.[c.key];
        const pct = parseFloat(silo?.pct) || 0;
        if (silo?.materialId && pct > 0) {
          const kgPorTonelada = pct / 100 * kgAgregadoPorTonelada;
          consumoProducao[silo.materialId] = (consumoProducao[silo.materialId] || 0) + toneladas * (kgPorTonelada / 1e3);
        }
      });
      TRABALHO_COLS.forEach((c) => {
        const item = formula.trabalho?.[c.key];
        const kgPorTonelada = parseFloat(item?.design) || 0;
        if (item?.materialId && kgPorTonelada > 0) {
          consumoProducao[item.materialId] = (consumoProducao[item.materialId] || 0) + toneladas * (kgPorTonelada / 1e3);
        }
      });
    });
  });
  const linhas = {};
  rececoesDoCentro.forEach((r) => {
    if (!linhas[r.produtoId]) linhas[r.produtoId] = { produtoId: r.produtoId, categoria: r.categoria };
  });
  Object.keys(consumoProducao).forEach((materialId) => {
    if (!linhas[materialId] && materiais.some((m) => m.id === materialId)) linhas[materialId] = { produtoId: materialId, categoria: "materiais" };
  });
  if (blocoTermico.combustivelId && !linhas[blocoTermico.combustivelId]) {
    linhas[blocoTermico.combustivelId] = { produtoId: blocoTermico.combustivelId, categoria: "consumiveis" };
  }
  (ajustesStock || []).filter((a) => a.centroId === center.id).forEach((a) => {
    if (!linhas[a.produtoId]) linhas[a.produtoId] = { produtoId: a.produtoId, categoria: a.categoria };
  });
  const nomeProduto = (categoria, id) => {
    const lista = categoria === "consumiveis" ? consumiveis : materiais;
    return lista.find((p) => p.id === id)?.designacao || "\u2014";
  };
  const categoriaLabel = { agregados: "Agregado", betumes: "Betume", consumiveis: "Combust\xEDvel", materiais: "Material" };
  const linhasCalculadas = Object.values(linhas).map((l) => {
    const movimentos = calcularMovimentosStock({ center, produtoId: l.produtoId, categoria: l.categoria, rececoes, diarias, formulas, ajustesStock });
    const entradas = movimentos.filter((m) => m.quantidade > 0).reduce((s, m) => s + m.quantidade, 0);
    const saidas = -movimentos.filter((m) => m.quantidade < 0).reduce((s, m) => s + m.quantidade, 0);
    const stock = movimentos.length > 0 ? movimentos[movimentos.length - 1].saldo : 0;
    return { ...l, entradas, saidas, stock };
  }).sort((a, b) => nomeProduto(a.categoria, a.produtoId).localeCompare(nomeProduto(b.categoria, b.produtoId)));
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mb-6", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
        /* @__PURE__ */ jsx(Archive, { size: 17, className: "text-amber-600" }),
        " Stocks"
      ] }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mt-1", children: "Calculado automaticamente: rece\xE7\xF5es menos consumo (produ\xE7\xE3o, para Agregados/Betumes; toneladas produzidas \xD7 taxa do Bloco T\xE9rmico, para Combust\xEDveis), com possibilidade de ajuste manual. Clique numa linha para ver o hist\xF3rico." })
    ] }),
    linhasCalculadas.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Archive, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 rece\xE7\xF5es ou consumo registados neste centro." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: linhasCalculadas.map((l, i) => {
      const un = l.categoria === "consumiveis" ? "L" : "t";
      return /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => onOpenHistoricoStock(l.produtoId, l.categoria),
          className: `w-full text-left flex items-center justify-between px-5 py-3.5 hover:bg-stone-50 ${i !== linhasCalculadas.length - 1 ? "border-b border-stone-100" : ""}`,
          children: [
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
                /* @__PURE__ */ jsx("p", { className: "font-display font-medium text-stone-900 truncate", children: nomeProduto(l.categoria, l.produtoId) }),
                /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-stone-100 text-stone-600 shrink-0", children: categoriaLabel[l.categoria] })
              ] }),
              /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400", children: [
                "Entradas: ",
                /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-600", children: [
                  l.entradas.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
                  " ",
                  un
                ] }),
                " ",
                "\xB7 Sa\xEDdas: ",
                /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-600", children: [
                  l.saidas.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
                  " ",
                  un
                ] })
              ] })
            ] }),
            /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 shrink-0", children: [
              /* @__PURE__ */ jsxs("span", { className: `font-mono-data font-semibold text-lg ${l.stock < 0 ? "text-red-600" : "text-stone-800"}`, children: [
                l.stock.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
                " ",
                un
              ] }),
              /* @__PURE__ */ jsx(History, { size: 16, className: "text-stone-300" })
            ] })
          ]
        },
        l.produtoId
      );
    }) }),
    /* @__PURE__ */ jsx("div", { className: "bg-stone-50 border border-dashed border-stone-300 rounded-xl p-4 mt-4 text-xs text-stone-400", children: "Agregados e Betumes: sa\xEDda calculada a partir das di\xE1rias de produ\xE7\xE3o \u2014 para cada tonelada produzida, reparte-se pela dosifica\xE7\xE3o de agregados a frio (%) e pela f\xF3rmula de trabalho a quente (kg por tonelada) da f\xF3rmula usada nessa linha. Combust\xEDveis: sa\xEDda = total de toneladas produzidas neste centro \xD7 taxa em vigor no Bloco T\xE9rmico (litros/tonelada), definida em Parametriza\xE7\xE3o de Produ\xE7\xE3o." })
  ] });
}
function AgregadosSection({ center, proveniencias, isAdmin, onBack, onAdd, onEdit, onDelete, onToggleStatus }) {
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-4xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between mb-2", children: [
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
          /* @__PURE__ */ jsx(Mountain, { size: 17, className: "text-amber-600" }),
          " Proveni\xEAncias de Agregados"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mt-1", children: "Origens (pedreiras/fornecedores) usadas nas rece\xE7\xF5es de agregados deste centro." })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onAdd, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700 shrink-0", children: [
        /* @__PURE__ */ jsx(Plus, { size: 15 }),
        " Nova Proveni\xEAncia"
      ] })
    ] }),
    proveniencias.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center mt-4", children: [
      /* @__PURE__ */ jsx(Mountain, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 proveni\xEAncias configuradas neste centro." }),
      !isAdmin && /* @__PURE__ */ jsx("p", { className: "text-stone-400 text-xs mt-1", children: "S\xF3 o Administrador pode configurar proveni\xEAncias." })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden mt-4", children: proveniencias.map((p, i) => {
      const provAtiva = p.ativo !== false;
      return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-5 py-3.5 ${i !== proveniencias.length - 1 ? "border-b border-stone-100" : ""} ${provAtiva ? "" : "opacity-60"}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
            /* @__PURE__ */ jsx("span", { className: "font-display font-medium text-stone-900 truncate", children: p.designacao }),
            /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${provAtiva ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: provAtiva ? "Ativo" : "Inativo" })
          ] }),
          /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500 truncate", children: p.fornecedor })
        ] }),
        isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
          /* @__PURE__ */ jsx(
            "button",
            {
              onClick: () => onToggleStatus(p.id),
              title: provAtiva ? "Desativar" : "Ativar",
              className: `p-1.5 rounded-lg ${provAtiva ? "text-stone-400 hover:text-red-600 hover:bg-red-50" : "text-stone-400 hover:text-emerald-600 hover:bg-emerald-50"}`,
              children: provAtiva ? /* @__PURE__ */ jsx(PowerOff, { size: 14 }) : /* @__PURE__ */ jsx(Power, { size: 14 })
            }
          ),
          /* @__PURE__ */ jsx("button", { onClick: () => onEdit(p), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
          /* @__PURE__ */ jsx("button", { onClick: () => onDelete(p.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
        ] })
      ] }, p.id);
    }) }),
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 border border-stone-200 rounded-lg px-4 py-3 mt-6 text-sm text-stone-600 flex items-center gap-2", children: [
      /* @__PURE__ */ jsx(Construction, { size: 16, className: "text-stone-400 shrink-0" }),
      "O lan\xE7amento das rece\xE7\xF5es de agregados propriamente ditas vai ser adicionado aqui a seguir."
    ] })
  ] });
}
function PlaceholderSection({ title, onBack }) {
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-3xl", children: [
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " Voltar"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-16 text-center", children: [
      /* @__PURE__ */ jsx(Construction, { className: "mx-auto text-stone-300 mb-4", size: 36 }),
      /* @__PURE__ */ jsx("h2", { className: "font-display text-xl text-stone-800 font-semibold mb-1", children: title }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Esta sec\xE7\xE3o ainda n\xE3o est\xE1 dispon\xEDvel. Vamos constru\xED-la a seguir." })
    ] })
  ] });
}
function ArtigosSection({ center, mixtures, canManage, isAdmin, onBack, onAddMixture, onEditMixture, onDeleteMixture, onToggleMixtureStatus, onImport, onOpenEliminarTodosArtigos }) {
  const [query, setQuery] = useState("");
  const contagemCodigos = {};
  mixtures.forEach((m) => {
    const cod = (m.codigoManual || "").trim();
    if (cod) contagemCodigos[cod] = (contagemCodigos[cod] || 0) + 1;
  });
  const codigosDuplicados = Object.keys(contagemCodigos).filter((c) => contagemCodigos[c] > 1);
  const artigosDuplicados = mixtures.filter((m) => codigosDuplicados.includes((m.codigoManual || "").trim()));
  const filtrados = query ? mixtures.filter((m) => matchesSearch(query, m.designacao, formatArticleCode(center, m))) : mixtures;
  const sorted = [...filtrados].sort((a, b) => {
    const na = typeof a.numero === "number" ? a.numero : null;
    const nb = typeof b.numero === "number" ? b.numero : null;
    if (na !== null && nb !== null) return na - nb;
    if (na !== null) return -1;
    if (nb !== null) return 1;
    return (a.codigoManual || "").localeCompare(b.codigoManual || "", "pt", { numeric: true });
  });
  return /* @__PURE__ */ jsxs("div", { className: "p-8 max-w-5xl", children: [
    artigosDuplicados.length > 0 && /* @__PURE__ */ jsxs("div", { className: "bg-red-50 border border-red-200 rounded-xl p-4 mb-4", children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm font-semibold text-red-700 mb-1.5 flex items-center gap-1.5", children: [
        /* @__PURE__ */ jsx(AlertTriangle, { size: 15 }),
        " ",
        codigosDuplicados.length,
        " c\xF3digo(s) repetido(s) entre artigos diferentes"
      ] }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-red-600 mb-2", children: "Estes artigos t\xEAm o mesmo c\xF3digo mas s\xE3o registos diferentes \u2014 reveja e corrija manualmente:" }),
      /* @__PURE__ */ jsx("ul", { className: "text-xs text-red-700 space-y-0.5", children: artigosDuplicados.map((m) => /* @__PURE__ */ jsxs("li", { children: [
        "\xB7 ",
        /* @__PURE__ */ jsx("span", { className: "font-mono-data font-semibold", children: m.codigoManual }),
        " \u2014 ",
        m.designacao
      ] }, m.id)) })
    ] }),
    /* @__PURE__ */ jsxs("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 hover:text-amber-600 mb-6 font-medium", children: [
      /* @__PURE__ */ jsx(ChevronLeft, { size: 16 }),
      " ",
      center.nome
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-4", children: [
      /* @__PURE__ */ jsxs("h3", { className: "font-display text-lg text-stone-800 font-semibold flex items-center gap-2", children: [
        /* @__PURE__ */ jsx(Beaker, { size: 17, className: "text-amber-600" }),
        " Artigos"
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        isAdmin && mixtures.length > 0 && /* @__PURE__ */ jsxs("button", { onClick: onOpenEliminarTodosArtigos, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-red-200 text-red-600 rounded-lg text-sm font-semibold hover:bg-red-50", children: [
          /* @__PURE__ */ jsx(Trash2, { size: 15 }),
          " Eliminar Tudo"
        ] }),
        /* @__PURE__ */ jsxs(
          "button",
          {
            onClick: () => exportarListaExcel(
              "artigos.xlsx",
              "Artigos",
              ["C\xF3digo", "Designa\xE7\xE3o", "Refer\xEAncia Relat\xF3rio", "Data Execu\xE7\xE3o", "Estado"],
              sorted.map((m) => [formatArticleCode(center, m) || m.codigoManual || "", m.designacao || "", m.referenciaRelatorio || "", m.dataExecucao || "", m.ativo !== false ? "Ativo" : "Inativo"])
            ),
            className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50",
            children: [
              /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
              " Exportar Excel"
            ]
          }
        ),
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onImport, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
          /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
          " Importar Excel"
        ] }),
        isAdmin && /* @__PURE__ */ jsxs("button", { onClick: onAddMixture, className: "flex items-center gap-1.5 px-3.5 py-2 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
          /* @__PURE__ */ jsx(Plus, { size: 15 }),
          " Novo Artigo"
        ] })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "relative mb-4", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 16 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          placeholder: "Pesquisar por c\xF3digo ou designa\xE7\xE3o...",
          className: "w-full pl-9 pr-9 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-800 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
        }
      ),
      query && /* @__PURE__ */ jsx("button", { onClick: () => setQuery(""), className: "absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
    ] }),
    mixtures.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Beaker, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsx("p", { className: "text-stone-500 text-sm", children: "Ainda n\xE3o h\xE1 artigos registados neste centro." })
    ] }) : sorted.length === 0 ? /* @__PURE__ */ jsxs("div", { className: "bg-white border border-dashed border-stone-300 rounded-xl p-10 text-center", children: [
      /* @__PURE__ */ jsx(Search, { className: "mx-auto text-stone-300 mb-3", size: 30 }),
      /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-sm", children: [
        'Nenhum artigo encontrado para "',
        query,
        '".'
      ] })
    ] }) : /* @__PURE__ */ jsx("div", { className: "bg-white rounded-xl border border-stone-200 overflow-hidden", children: sorted.map((m, i) => {
      const artigoAtivo = m.ativo !== false;
      const validade = calcValidade(m.dataExecucao);
      let validadeStatus = null;
      if (validade) {
        const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
        const diasRestantes = (new Date(validade) - new Date(hoje)) / 864e5;
        if (diasRestantes < 0) validadeStatus = { label: "ETI Expirado", cls: "bg-red-100 text-red-700" };
        else if (diasRestantes <= 90) validadeStatus = { label: "ETI a expirar", cls: "bg-amber-100 text-amber-700" };
        else validadeStatus = { label: "ETI v\xE1lido", cls: "bg-emerald-100 text-emerald-700" };
      } else {
        validadeStatus = { label: "Sem ensaio", cls: "bg-red-100 text-red-700" };
      }
      const alerta = !m.dataExecucao || validadeStatus.label === "ETI Expirado";
      return /* @__PURE__ */ jsxs("div", { className: `px-5 py-3.5 ${i !== sorted.length - 1 ? "border-b border-stone-100" : ""} ${artigoAtivo ? "" : "opacity-60"}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between", children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
            formatArticleCode(center, m) && /* @__PURE__ */ jsx("span", { className: `font-mono-data text-xs font-semibold px-2 py-1 rounded shrink-0 ${alerta ? "bg-red-50 text-red-700" : "bg-amber-50 text-amber-700"}`, children: formatArticleCode(center, m) }),
            /* @__PURE__ */ jsx("span", { className: `font-display font-medium truncate ${!artigoAtivo ? "text-red-600 line-through decoration-red-600 decoration-2" : "text-stone-900"}`, children: m.designacao }),
            /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${artigoAtivo ? "bg-emerald-100 text-emerald-700" : "bg-stone-200 text-stone-500"}`, children: artigoAtivo ? "Ativo" : "Inativo" }),
            codigosDuplicados.includes((m.codigoManual || "").trim()) && /* @__PURE__ */ jsxs("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 bg-red-100 text-red-700 flex items-center gap-1", children: [
              /* @__PURE__ */ jsx(AlertTriangle, { size: 10 }),
              " C\xF3digo repetido"
            ] })
          ] }),
          isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
            /* @__PURE__ */ jsx(
              "button",
              {
                onClick: () => onToggleMixtureStatus(m.id),
                title: artigoAtivo ? "Desativar artigo" : "Ativar artigo",
                className: `p-1.5 rounded-lg ${artigoAtivo ? "text-stone-400 hover:text-red-600 hover:bg-red-50" : "text-stone-400 hover:text-emerald-600 hover:bg-emerald-50"}`,
                children: artigoAtivo ? /* @__PURE__ */ jsx(PowerOff, { size: 14 }) : /* @__PURE__ */ jsx(Power, { size: 14 })
              }
            ),
            /* @__PURE__ */ jsx("button", { onClick: () => onEditMixture(m), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => onDeleteMixture(m.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mt-1.5 pl-0.5 text-xs text-stone-500", children: [
          m.referenciaRelatorio && /* @__PURE__ */ jsx("span", { className: "font-mono-data", children: m.referenciaRelatorio }),
          m.dataExecucao ? /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsxs("span", { children: [
              "Ensaio: ",
              formatDatePT(m.dataExecucao)
            ] }),
            validade && /* @__PURE__ */ jsxs("span", { children: [
              "V\xE1lido at\xE9 ",
              formatDatePT(validade)
            ] })
          ] }) : /* @__PURE__ */ jsx("span", { className: "text-red-500", children: "Sem data de execu\xE7\xE3o do ensaio" }),
          /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full ${validadeStatus.cls}`, children: validadeStatus.label })
        ] })
      ] }, m.id);
    }) })
  ] });
}
function LogotipoModal({ logotipoAtual, onSave, onRemove, onClose }) {
  const [preview, setPreview] = useState(logotipoAtual || "");
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) return setError("Escolha um ficheiro de imagem (PNG, JPG, SVG...)");
    if (file.size > 1024 * 1024) return setError("A imagem deve ter menos de 1 MB");
    setError("");
    const reader = new FileReader();
    reader.onload = (evt) => setPreview(evt.target.result);
    reader.readAsDataURL(file);
  };
  const submit = () => {
    if (!preview) return setError("Escolha uma imagem primeiro");
    onSave(preview);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Log\xF3tipo da Empresa", subtitle: "Usado nos documentos exportados (ex: PDF da Di\xE1ria)", onClose, children: [
    /* @__PURE__ */ jsxs(
      "div",
      {
        onClick: () => fileInputRef.current?.click(),
        className: "border border-dashed border-stone-300 rounded-lg p-6 text-center mb-4 cursor-pointer hover:bg-stone-50",
        children: [
          preview ? /* @__PURE__ */ jsx("img", { src: preview, alt: "Pr\xE9-visualiza\xE7\xE3o do log\xF3tipo", className: "max-h-24 mx-auto mb-2" }) : /* @__PURE__ */ jsx(Image, { className: "mx-auto text-stone-300 mb-2", size: 32 }),
          /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: preview ? "Clique para escolher outra imagem" : "Clique para escolher uma imagem" }),
          /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mt-1", children: "PNG, JPG ou SVG, at\xE9 1 MB" }),
          /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: "image/*", onChange: handleFile, className: "hidden" })
        ]
      }
    ),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full mb-2 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" }),
    logotipoAtual && /* @__PURE__ */ jsx("button", { onClick: onRemove, className: "w-full py-2.5 rounded-lg border border-red-200 text-red-600 font-display font-semibold tracking-wide uppercase text-xs hover:bg-red-50", children: "Remover Log\xF3tipo" })
  ] });
}
function BackupModal({ onExport, onImport, onClose }) {
  const [error, setError] = useState("");
  const [sucesso, setSucesso] = useState(false);
  const [confirmando, setConfirmando] = useState(false);
  const [dadosPendentes, setDadosPendentes] = useState(null);
  const fileInputRef = useRef(null);
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        if (!parsed || typeof parsed !== "object" || !parsed.versaoBackup) {
          throw new Error("formato inv\xE1lido");
        }
        setDadosPendentes(parsed);
        setConfirmando(true);
      } catch (err) {
        setError("Ficheiro inv\xE1lido \u2014 escolha um ficheiro de backup exportado por esta aplica\xE7\xE3o.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };
  const confirmarImportacao = () => {
    onImport(dadosPendentes);
    setConfirmando(false);
    setDadosPendentes(null);
    setSucesso(true);
    setTimeout(() => setSucesso(false), 3e3);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Backup", subtitle: "Seguran\xE7a dos dados da aplica\xE7\xE3o", onClose, wide: true, children: [
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-stone-200 rounded-xl p-5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Exportar Backup Completo" }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mb-3", children: "Descarrega um \xFAnico ficheiro com todos os dados da aplica\xE7\xE3o \u2014 centros, f\xF3rmulas, clientes, obras, di\xE1rias, utilizadores, materiais, combust\xEDveis, equipamentos, m\xE3o de obra e todas as parametriza\xE7\xF5es. Guarde-o num s\xEDtio seguro (ex: computador ou Google Drive) e repita regularmente." }),
      /* @__PURE__ */ jsxs("button", { onClick: onExport, className: "flex items-center gap-1.5 px-4 py-2.5 bg-amber-600 text-white rounded-lg text-sm font-semibold hover:bg-amber-700", children: [
        /* @__PURE__ */ jsx(FileText, { size: 16 }),
        " Exportar Backup Completo"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "bg-white border border-stone-200 rounded-xl p-5", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Importar Backup" }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-red-600 mb-3 font-medium", children: "Aten\xE7\xE3o: importar um backup substitui TODOS os dados atuais da aplica\xE7\xE3o pelos dados do ficheiro. Use apenas para recuperar dados perdidos." }),
      /* @__PURE__ */ jsxs("button", { onClick: () => fileInputRef.current?.click(), className: "flex items-center gap-1.5 px-4 py-2.5 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(Upload, { size: 16 }),
        " Escolher Ficheiro de Backup"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".json", onChange: handleFile, className: "hidden" }),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error }),
      sucesso && /* @__PURE__ */ jsxs("p", { className: "text-emerald-600 text-sm mt-3 flex items-center gap-1.5", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 15 }),
        " Backup importado com sucesso."
      ] })
    ] }),
    confirmando && dadosPendentes && /* @__PURE__ */ jsx("div", { className: "fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4", children: /* @__PURE__ */ jsxs("div", { className: "bg-white rounded-xl p-6 max-w-sm w-full", children: [
      /* @__PURE__ */ jsx("p", { className: "font-display font-semibold text-stone-900 mb-2", children: "Confirmar Importa\xE7\xE3o" }),
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-600 mb-1", children: [
        "Ficheiro de ",
        dadosPendentes.exportadoEm ? formatDateTimePT(dadosPendentes.exportadoEm) : "data desconhecida",
        "."
      ] }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-red-600 mb-4", children: "Isto substitui TODOS os dados atuais desta aplica\xE7\xE3o. N\xE3o pode ser desfeito. Tem a certeza?" }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: confirmarImportacao, className: "flex-1 py-2.5 rounded-lg bg-red-600 text-white text-sm font-semibold hover:bg-red-700", children: "Sim, substituir tudo" }),
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setConfirmando(false);
          setDadosPendentes(null);
        }, className: "flex-1 py-2.5 rounded-lg border border-stone-300 text-stone-600 text-sm font-semibold hover:bg-stone-50", children: "Cancelar" })
      ] })
    ] }) })
  ] });
}
function DopConfigModal({ dopConfig, onSave, onClose }) {
  const c = dopConfig || {};
  const [fabricanteNome, setFabricanteNome] = useState(c.fabricanteNome || "Socorpena - Engenharia e Constru\xE7\xE3o, SA");
  const [fabricanteMorada, setFabricanteMorada] = useState(c.fabricanteMorada || "");
  const [mandatario, setMandatario] = useState(c.mandatario || "N\xE3o aplic\xE1vel.");
  const [avcp, setAvcp] = useState(c.avcp || "Sistema 2+");
  const [organismoNotificado, setOrganismoNotificado] = useState(c.organismoNotificado || "");
  const [documentoAvaliacaoEuropeu, setDocumentoAvaliacaoEuropeu] = useState(c.documentoAvaliacaoEuropeu || "N\xE3o aplic\xE1vel.");
  const [avaliacaoTecnicaEuropeia, setAvaliacaoTecnicaEuropeia] = useState(c.avaliacaoTecnicaEuropeia || "N\xE3o aplic\xE1vel.");
  const [signatarioNome, setSignatarioNome] = useState(c.signatarioNome || "");
  const [signatarioCargo, setSignatarioCargo] = useState(c.signatarioCargo || "Representante da Gest\xE3o para o Controlo de Produ\xE7\xE3o em F\xE1brica");
  const [localEmissao, setLocalEmissao] = useState(c.localEmissao || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!fabricanteNome.trim()) return setError("Indique o nome do fabricante");
    onSave({
      fabricanteNome: fabricanteNome.trim(),
      fabricanteMorada: fabricanteMorada.trim(),
      mandatario: mandatario.trim(),
      avcp: avcp.trim(),
      organismoNotificado: organismoNotificado.trim(),
      documentoAvaliacaoEuropeu: documentoAvaliacaoEuropeu.trim(),
      avaliacaoTecnicaEuropeia: avaliacaoTecnicaEuropeia.trim(),
      signatarioNome: signatarioNome.trim(),
      signatarioCargo: signatarioCargo.trim(),
      localEmissao: localEmissao.trim()
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Defini\xE7\xF5es DoP", subtitle: "Dados fixos usados em todas as Declara\xE7\xF5es de Desempenho", onClose, wide: true, children: [
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-4", children: "Estes dados s\xE3o comuns a todas as f\xF3rmulas \u2014 s\xF3 a Norma Harmonizada e os dados de desempenho variam, e s\xE3o preenchidos em cada f\xF3rmula." }),
    /* @__PURE__ */ jsx(Field, { label: "Fabricante (nome)", children: /* @__PURE__ */ jsx("input", { value: fabricanteNome, onChange: (e) => setFabricanteNome(e.target.value), className: inputCls }) }),
    /* @__PURE__ */ jsx(Field, { label: "Fabricante (morada, inclui o Centro de Produ\xE7\xE3o se aplic\xE1vel)", children: /* @__PURE__ */ jsx("textarea", { value: fabricanteMorada, onChange: (e) => setFabricanteMorada(e.target.value), className: inputCls, rows: 2, placeholder: "Ex: Zona Industrial do Entroncamento N.\xBA 10, 4870-118 Ribeira de Pena" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Mandat\xE1rio", children: /* @__PURE__ */ jsx("input", { value: mandatario, onChange: (e) => setMandatario(e.target.value), className: inputCls }) }),
    /* @__PURE__ */ jsx(Field, { label: "Sistema de Avalia\xE7\xE3o e Verifica\xE7\xE3o (AVCP)", children: /* @__PURE__ */ jsx("input", { value: avcp, onChange: (e) => setAvcp(e.target.value), className: inputCls }) }),
    /* @__PURE__ */ jsx(Field, { label: "Organismo(s) Notificado(s)", children: /* @__PURE__ */ jsx("textarea", { value: organismoNotificado, onChange: (e) => setOrganismoNotificado(e.target.value), className: inputCls, rows: 2, placeholder: "Ex: EIC - Empresa Internacional de Certifica\xE7\xE3o, Organismo Notificado n\xBA. 1515\nCertificado de Conformidade do Controlo de Produ\xE7\xE3o em F\xE1brica n\xBA 1515-CPR-0255" }) }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Documento de Avalia\xE7\xE3o Europeu", children: /* @__PURE__ */ jsx("input", { value: documentoAvaliacaoEuropeu, onChange: (e) => setDocumentoAvaliacaoEuropeu(e.target.value), className: inputCls }) }),
      /* @__PURE__ */ jsx(Field, { label: "Avalia\xE7\xE3o T\xE9cnica Europeia", children: /* @__PURE__ */ jsx("input", { value: avaliacaoTecnicaEuropeia, onChange: (e) => setAvaliacaoTecnicaEuropeia(e.target.value), className: inputCls }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Nome de quem assina", children: /* @__PURE__ */ jsx("input", { value: signatarioNome, onChange: (e) => setSignatarioNome(e.target.value), className: inputCls }) }),
      /* @__PURE__ */ jsx(Field, { label: "Cargo de quem assina", children: /* @__PURE__ */ jsx("input", { value: signatarioCargo, onChange: (e) => setSignatarioCargo(e.target.value), className: inputCls }) })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Local de emiss\xE3o (para a data/assinatura)", children: /* @__PURE__ */ jsx("input", { value: localEmissao, onChange: (e) => setLocalEmissao(e.target.value), className: inputCls, placeholder: "Ex: Ribeira de Pena" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function UserModal({ data, centers, isAdmin, currentUserRole, perfisPersonalizados, onSave, onResetPin, onSetPin, onClose }) {
  const [nome, setNome] = useState(data?.nome || "");
  const [primeiroNome, setPrimeiroNome] = useState(data?.primeiroNome || "");
  const [ultimoNome, setUltimoNome] = useState(data?.ultimoNome || "");
  const [assinatura, setAssinatura] = useState(data?.assinatura || "");
  const [email, setEmail] = useState(data?.email || "");
  const [role, setRole] = useState(data?.role || "Operador");
  const [escalas, setEscalas] = useState(data?.escalasCentros || []);
  const [novaEscalaCentroId, setNovaEscalaCentroId] = useState("");
  const [novaEscalaInicio, setNovaEscalaInicio] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [novaEscalaFim, setNovaEscalaFim] = useState("");
  const [error, setError] = useState("");
  const [novoPin, setNovoPin] = useState("");
  const [pinDefinido, setPinDefinido] = useState(false);
  const fileInputRef = useRef(null);
  const soCalendario = !isAdmin && currentUserRole === "Gestor" && data?.role === "Operador";
  const submitNovoPin = () => {
    if (!/^\d{4}$/.test(novoPin)) return setError("A palavra-passe tem de ter exatamente 4 d\xEDgitos");
    onSetPin(data.id, novoPin);
    setNovoPin("");
    setError("");
    setPinDefinido(true);
    setTimeout(() => setPinDefinido(false), 2e3);
  };
  const handleAssinaturaFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) return setError("Escolha um ficheiro de imagem (PNG, JPG...)");
    if (file.size > 512 * 1024) return setError("A imagem deve ter menos de 500 KB");
    setError("");
    const reader = new FileReader();
    reader.onload = (evt) => setAssinatura(evt.target.result);
    reader.readAsDataURL(file);
  };
  const submit = () => {
    if (!nome.trim()) return setError("Indique o nome (login)");
    if (!email.trim() || !/^\S+@\S+\.\S+$/.test(email.trim())) return setError("Indique um email v\xE1lido");
    let escalasFinais = escalas;
    if (role === "Operador" && novaEscalaCentroId && novaEscalaInicio) {
      escalasFinais = [...escalas, { id: genId(), centroId: novaEscalaCentroId, dataInicio: novaEscalaInicio, dataFim: novaEscalaFim || null }];
    }
    const payload = {
      id: data?.id,
      nome: nome.trim(),
      primeiroNome: primeiroNome.trim(),
      ultimoNome: ultimoNome.trim(),
      assinatura,
      email: email.trim(),
      role,
      escalasCentros: role === "Operador" ? escalasFinais : []
    };
    onSave(payload);
  };
  const addEscala = () => {
    if (!novaEscalaCentroId) return setError("Escolha o centro");
    if (!novaEscalaInicio) return setError("Indique a data de in\xEDcio");
    if (novaEscalaFim && novaEscalaFim < novaEscalaInicio) return setError("A data de fim n\xE3o pode ser anterior \xE0 de in\xEDcio");
    setEscalas([...escalas, { id: genId(), centroId: novaEscalaCentroId, dataInicio: novaEscalaInicio, dataFim: novaEscalaFim || null }]);
    setNovaEscalaCentroId("");
    setNovaEscalaInicio((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
    setNovaEscalaFim("");
    setError("");
  };
  const removeEscala = (id) => setEscalas(escalas.filter((e) => e.id !== id));
  const roleHint = {
    "Administrador": "Acesso total: utilizadores, centros e artigos em toda a aplica\xE7\xE3o.",
    "Gestor": "Gere artigos em todos os centros. N\xE3o cria/edita centros nem gere utilizadores.",
    "Operador": "O acesso aos centros \xE9 definido pelo calend\xE1rio abaixo \u2014 pode ter mais do que um centro atribu\xEDdo em simult\xE2neo.",
    "Or\xE7amentista": "Acesso de consulta por agora. Vai ter um menu pr\xF3prio de or\xE7amenta\xE7\xE3o futuramente.",
    "Convidado": "Acesso de consulta apenas \u2014 n\xE3o cria nem edita nada na aplica\xE7\xE3o."
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Utilizador" : "Novo Utilizador", subtitle: "Equipa", onClose, children: [
    soCalendario && /* @__PURE__ */ jsxs("div", { className: "bg-amber-50 border border-amber-200 rounded-lg px-3 py-2.5 mb-4", children: [
      /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-800", children: nome }),
      email && /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-500", children: email }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-700 mt-1", children: "S\xF3 pode gerir o calend\xE1rio de acesso aos centros desta pessoa \u2014 para alterar nome, email ou perfil, pe\xE7a a um Administrador." })
    ] }),
    !soCalendario && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs(Field, { label: "Nome", children: [
        /* @__PURE__ */ jsx("input", { value: nome, onChange: (e) => setNome(e.target.value), className: inputCls, placeholder: "Nome completo", autoFocus: true }),
        /* @__PURE__ */ jsx("span", { className: "block text-xs text-stone-400 mt-1", children: "Este nome \xE9 tamb\xE9m o nome de utilizador usado para entrar na aplica\xE7\xE3o." })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Primeiro Nome (opcional)", children: /* @__PURE__ */ jsx("input", { value: primeiroNome, onChange: (e) => setPrimeiroNome(e.target.value), className: inputCls, placeholder: "Ex: Jo\xE3o" }) }),
        /* @__PURE__ */ jsx(Field, { label: "\xDAltimo Nome (opcional)", children: /* @__PURE__ */ jsx("input", { value: ultimoNome, onChange: (e) => setUltimoNome(e.target.value), className: inputCls, placeholder: "Ex: Silva" }) })
      ] }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 -mt-2 mb-4", children: "Usados s\xF3 para a assinatura manuscrita dos documentos (Di\xE1ria, etc.) \u2014 n\xE3o afetam o login." }),
      /* @__PURE__ */ jsxs(Field, { label: "Email", children: [
        /* @__PURE__ */ jsx("input", { value: email, onChange: (e) => setEmail(e.target.value), type: "email", className: inputCls, placeholder: "nome@empresa.pt" }),
        /* @__PURE__ */ jsx("span", { className: "block text-xs text-stone-400 mt-1", children: "Usado como contacto/identifica\xE7\xE3o." })
      ] }),
      /* @__PURE__ */ jsxs(Field, { label: "Assinatura", children: [
        /* @__PURE__ */ jsxs("div", { onClick: () => fileInputRef.current?.click(), className: "border border-dashed border-stone-300 rounded-lg p-4 text-center cursor-pointer hover:bg-stone-50", children: [
          assinatura ? /* @__PURE__ */ jsx("img", { src: assinatura, alt: "Assinatura", className: "max-h-16 mx-auto" }) : /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400", children: "Clique para carregar uma imagem da assinatura (PNG, at\xE9 500 KB)" }),
          /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: "image/*", onChange: handleAssinaturaFile, className: "hidden" })
        ] }),
        /* @__PURE__ */ jsxs("span", { className: "block text-xs text-stone-400 mt-1", children: [
          "Sem assinatura carregada, os documentos usam o ",
          primeiroNome || ultimoNome ? `nome (${`${primeiroNome} ${ultimoNome}`.trim()})` : "nome de login",
          " num estilo manuscrito, a azul."
        ] }),
        assinatura && /* @__PURE__ */ jsx("button", { type: "button", onClick: () => setAssinatura(""), className: "text-xs text-red-600 hover:text-red-700 mt-1", children: "Remover assinatura" })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Perfil", children: /* @__PURE__ */ jsxs("select", { value: role, onChange: (e) => setRole(e.target.value), disabled: !isAdmin, className: `${inputCls} disabled:bg-stone-100`, children: [
        ROLES.map((r) => /* @__PURE__ */ jsx("option", { value: r, children: r }, r)),
        (perfisPersonalizados || []).map((p) => /* @__PURE__ */ jsx("option", { value: p.nome, children: p.nome }, p.id))
      ] }) }),
      /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 -mt-2 mb-4", children: [
        roleHint[role],
        !isAdmin && " S\xF3 o Administrador pode alterar o papel."
      ] })
    ] }),
    role === "Operador" && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Calend\xE1rio de Acesso aos Centros" }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Defina em que per\xEDodos esta pessoa tem acesso a cada centro. Pode sobrepor per\xEDodos para dar acesso a mais do que um centro ao mesmo tempo. Sem data de fim = acesso cont\xEDnuo a partir da data de in\xEDcio." }),
      /* @__PURE__ */ jsx("div", { className: "space-y-2 mb-3", children: escalas.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 italic", children: "Ainda sem centros atribu\xEDdos \u2014 esta pessoa n\xE3o vai ver nenhum centro dispon\xEDvel." }) : escalas.slice().sort((a, b) => (a.dataInicio || "").localeCompare(b.dataInicio || "")).map((e) => /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between border border-stone-200 rounded-lg px-3 py-2 text-sm bg-stone-50/60", children: [
        /* @__PURE__ */ jsxs("div", { children: [
          /* @__PURE__ */ jsx("span", { className: "font-medium text-stone-800", children: centers.find((c) => c.id === e.centroId)?.nome || "Centro removido" }),
          /* @__PURE__ */ jsxs("span", { className: "text-stone-400 text-xs ml-2", children: [
            formatDatePT(e.dataInicio),
            " \u2192 ",
            e.dataFim ? formatDatePT(e.dataFim) : "cont\xEDnuo"
          ] })
        ] }),
        /* @__PURE__ */ jsx("button", { onClick: () => removeEscala(e.id), className: "text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
      ] }, e.id)) }),
      /* @__PURE__ */ jsxs("div", { className: "border border-dashed border-stone-300 rounded-lg p-3 mb-4", children: [
        /* @__PURE__ */ jsx("div", { className: "grid grid-cols-1 gap-2 mb-2", children: /* @__PURE__ */ jsx(Field, { label: "Centro", children: /* @__PURE__ */ jsxs("select", { value: novaEscalaCentroId, onChange: (e) => setNovaEscalaCentroId(e.target.value), className: inputCls, children: [
          /* @__PURE__ */ jsx("option", { value: "", children: "Escolha o centro..." }),
          centers.map((c) => /* @__PURE__ */ jsx("option", { value: c.id, children: c.nome }, c.id))
        ] }) }) }),
        /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-2 mb-2", children: [
          /* @__PURE__ */ jsx(Field, { label: "In\xEDcio", children: /* @__PURE__ */ jsx("input", { type: "date", value: novaEscalaInicio, onChange: (e) => setNovaEscalaInicio(e.target.value), className: inputCls }) }),
          /* @__PURE__ */ jsx(Field, { label: "Fim (opcional)", children: /* @__PURE__ */ jsx("input", { type: "date", value: novaEscalaFim, onChange: (e) => setNovaEscalaFim(e.target.value), className: inputCls }) })
        ] }),
        /* @__PURE__ */ jsxs("button", { type: "button", onClick: addEscala, className: "w-full py-2 rounded-lg border border-stone-300 text-stone-600 text-sm font-semibold hover:bg-stone-50 flex items-center justify-center gap-1.5", children: [
          /* @__PURE__ */ jsx(Plus, { size: 15 }),
          " Adicionar per\xEDodo"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-[11px] text-stone-400 mt-1.5", children: "Clique aqui para adicionar este per\xEDodo \xE0 lista acima \u2014 s\xF3 o que estiver na lista fica guardado." })
      ] })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" }),
    data?.id && isAdmin && /* @__PURE__ */ jsxs("div", { className: "mt-3 pt-3 border-t border-stone-200", children: [
      /* @__PURE__ */ jsx("p", { className: "text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1", children: "Palavra-passe atual" }),
      data.pin ? /* @__PURE__ */ jsx("p", { className: "font-mono-data text-lg tracking-widest text-stone-800 mb-2", children: data.pin }) : /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 mb-2", children: "Ainda n\xE3o foi definida \u2014 esta pessoa define-a no primeiro acesso." }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-2 mb-2", children: [
        /* @__PURE__ */ jsx(Field, { label: "Definir nova (4 d\xEDgitos)", children: /* @__PURE__ */ jsx(
          "input",
          {
            value: novoPin,
            onChange: (e) => setNovoPin(e.target.value.replace(/\D/g, "").slice(0, 4)),
            type: "text",
            inputMode: "numeric",
            maxLength: 4,
            className: `${inputCls} font-mono-data tracking-widest`,
            placeholder: "0000"
          }
        ) }),
        /* @__PURE__ */ jsx("button", { onClick: submitNovoPin, className: "mb-4 px-4 py-2.5 rounded-lg bg-stone-800 text-white text-sm font-semibold hover:bg-stone-900 shrink-0", children: pinDefinido ? "Definida \u2713" : "Definir" })
      ] }),
      /* @__PURE__ */ jsx("button", { onClick: () => onResetPin(data.id), className: "w-full py-2.5 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-xs hover:bg-stone-50", children: "Repor palavra-passe (obriga a definir uma nova no pr\xF3ximo acesso)" })
    ] })
  ] });
}
function ImportClientesModal({ onImport, onClose, tituloEntidade }) {
  const entidade = tituloEntidade || "Clientes";
  const entidadeSingular = entidade === "Fornecedores" ? "fornecedor" : "cliente";
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setRows(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const stripped = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
        const numeroKey = headerKeys.find((k) => ["n", "no", "numero"].includes(stripped(k)));
        const nifKey = headerKeys.find((k) => normalizeHeader(k) === "nif");
        const designacaoKey = headerKeys.find((k) => normalizeHeader(k) === "designacao");
        const moradaKey = headerKeys.find((k) => normalizeHeader(k) === "morada");
        if (!numeroKey || !nifKey || !designacaoKey) {
          setError('N\xE3o encontrei as colunas "N\xBA", "NIF" e "Designa\xE7\xE3o" na primeira linha da folha.');
          setRows(null);
          return;
        }
        const parsed = json.map((r) => ({
          numero: String(r[numeroKey] ?? "").trim(),
          nif: String(r[nifKey] ?? "").trim(),
          designacao: String(r[designacaoKey] ?? "").trim(),
          morada: moradaKey ? String(r[moradaKey] ?? "").trim() : ""
        })).filter((r) => r.numero || r.nif || r.designacao);
        if (parsed.length === 0) {
          setError("N\xE3o encontrei linhas com dados.");
          setRows(null);
          return;
        }
        setRows(parsed);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const confirm = () => {
    setImporting(true);
    onImport(rows);
  };
  const baixarModelo = () => {
    const exemplo = entidade === "Fornecedores" ? ["1001", "500123456", "Britas & Agregados, Lda.", "Zona Industrial, Lote 3, 5000-000 Vila Real"] : ["1045", "500987654", "Constru\xE7\xF5es Exemplo, S.A.", "Rua Principal 10, 4000-000 Porto"];
    const wsData = [["N\xBA", "NIF", "Designa\xE7\xE3o", "Morada"], exemplo];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 10 }, { wch: 14 }, { wch: 30 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, entidade);
    XLSX.writeFile(wb, `modelo_importacao_${entidade.toLowerCase()}.xlsx`);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: `Importar ${entidade}`, subtitle: "Ficheiro Excel", onClose, children: [
    !rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro .xlsx, .xls ou .csv com as colunas ",
        /* @__PURE__ */ jsx("strong", { children: "N\xBA" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "NIF" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Designa\xE7\xE3o" }),
        " na primeira linha. Uma coluna ",
        /* @__PURE__ */ jsx("strong", { children: "Morada" }),
        " \xE9 opcional."
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: baixarModelo, className: "w-full mb-3 flex items-center justify-center gap-1.5 py-2.5 rounded-lg border border-stone-300 text-stone-600 text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Descarregar modelo (.xlsx)"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        rows.length,
        " ",
        entidadeSingular,
        rows.length !== 1 ? "s" : "",
        " encontrado",
        rows.length !== 1 ? "s" : "",
        ' em "',
        fileName,
        '"'
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-64 overflow-y-auto mb-4", children: rows.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center gap-3 px-3 py-2 text-sm ${i !== rows.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: [
          "N\xBA ",
          r.numero || "\u2014"
        ] }),
        /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs text-stone-500 shrink-0", children: r.nif || "\u2014" }),
        /* @__PURE__ */ jsx("span", { className: "text-stone-700 truncate", children: r.designacao || "\u2014" })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRows(null);
          setFileName("");
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirm, disabled: importing, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function ImportModal({ centroId, onImport, onClose }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const asDate = (v) => {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v ?? "").trim();
  };
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setRows(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const flat = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
        const codigoKey = headerKeys.find((k) => flat(k) === "codigo");
        const designacaoKey = headerKeys.find((k) => flat(k) === "designacao");
        const referenciaKey = headerKeys.find((k) => flat(k).includes("referencia"));
        const dataExecucaoKey = headerKeys.find((k) => flat(k).includes("data") && (flat(k).includes("execucao") || flat(k).includes("ensaio")));
        if (!codigoKey || !designacaoKey) {
          setError('N\xE3o encontrei as colunas "C\xF3digo" e "Designa\xE7\xE3o" na primeira linha da folha.');
          setRows(null);
          return;
        }
        const parsed = json.map((r) => ({
          codigo: String(r[codigoKey] ?? "").trim(),
          designacao: String(r[designacaoKey] ?? "").trim(),
          referenciaRelatorio: referenciaKey ? String(r[referenciaKey] ?? "").trim() : "",
          dataExecucao: dataExecucaoKey ? asDate(r[dataExecucaoKey]) : ""
        })).filter((r) => r.codigo || r.designacao);
        if (parsed.length === 0) {
          setError("N\xE3o encontrei linhas com dados.");
          setRows(null);
          return;
        }
        setRows(parsed);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const confirm = () => {
    setImporting(true);
    onImport(centroId, rows);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Artigos", subtitle: "Ficheiro Excel", onClose, children: [
    !rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro .xlsx, .xls ou .csv com as colunas ",
        /* @__PURE__ */ jsx("strong", { children: "C\xF3digo" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Designa\xE7\xE3o" }),
        " na primeira linha. As colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Refer\xEAncia do Relat\xF3rio" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Data de Execu\xE7\xE3o" }),
        " s\xE3o opcionais \u2014 se existirem, a validade (+",
        VALIDADE_ANOS,
        " anos) \xE9 calculada automaticamente."
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        rows.length,
        " artigo",
        rows.length !== 1 ? "s" : "",
        " encontrado",
        rows.length !== 1 ? "s" : "",
        ' em "',
        fileName,
        '"'
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-64 overflow-y-auto mb-4", children: rows.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `px-3 py-2 text-sm ${i !== rows.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3", children: [
          /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: r.codigo || "\u2014" }),
          /* @__PURE__ */ jsx("span", { className: "text-stone-700 truncate", children: r.designacao || "\u2014" })
        ] }),
        (r.referenciaRelatorio || r.dataExecucao) && /* @__PURE__ */ jsxs("div", { className: "text-xs text-stone-400 pl-1 mt-0.5", children: [
          r.referenciaRelatorio && /* @__PURE__ */ jsx("span", { className: "font-mono-data", children: r.referenciaRelatorio }),
          r.dataExecucao && /* @__PURE__ */ jsxs("span", { children: [
            r.referenciaRelatorio ? " \xB7 " : "",
            "Ensaio: ",
            formatDatePT(r.dataExecucao),
            " \xB7 V\xE1lido at\xE9 ",
            formatDatePT(calcValidade(r.dataExecucao))
          ] })
        ] })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRows(null);
          setFileName("");
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirm, disabled: importing, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function ImportCentersModal({ centersExistentes, onImport, onClose }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setRows(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const flat = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
        const nomeKey = headerKeys.find((k) => flat(k) === "nome");
        const codigoKey = headerKeys.find((k) => flat(k) === "codigo");
        const localKey = headerKeys.find((k) => flat(k).includes("localizacao") || flat(k).includes("local"));
        if (!nomeKey || !codigoKey) {
          setError('N\xE3o encontrei as colunas "Nome" e "C\xF3digo" neste ficheiro. Confirme que segue o modelo.');
          setRows(null);
          return;
        }
        const existentesSet = new Set((centersExistentes || []).map((c) => normalizeHeader(c.codigo)));
        const parsed = json.map((r) => ({ nome: String(r[nomeKey] || "").trim(), codigo: String(r[codigoKey] || "").trim(), localizacao: localKey ? String(r[localKey] || "").trim() : "" })).filter((r) => r.nome && r.codigo);
        const novos = parsed.filter((r) => !existentesSet.has(normalizeHeader(r.codigo)));
        setRows({ total: parsed.length, novos });
      } catch (err) {
        setError("N\xE3o consegui ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const baixarModelo = () => {
    const wsData = [
      ["Nome", "C\xF3digo", "Localiza\xE7\xE3o"],
      ["Central de Vila Real", "BB", "Vila Real"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 30 }, { wch: 10 }, { wch: 20 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Centros");
    XLSX.writeFile(wb, "modelo_importacao_centros.xlsx");
  };
  const confirm = () => {
    setImporting(true);
    onImport(rows.novos);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Centros de Produ\xE7\xE3o", subtitle: "Excel", onClose, children: [
    !rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro .xlsx, .xls ou .csv com as colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Nome" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "C\xF3digo" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Localiza\xE7\xE3o" }),
        " (esta \xFAltima \xE9 opcional). Centros cujo c\xF3digo j\xE1 exista s\xE3o ignorados automaticamente."
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: baixarModelo, className: "flex items-center gap-1.5 mb-4 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Descarregar Modelo"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs("button", { onClick: () => fileInputRef.current?.click(), className: "w-full py-3 rounded-lg border-2 border-dashed border-stone-300 text-stone-500 text-sm font-semibold hover:bg-stone-50 flex items-center justify-center gap-2", children: [
        /* @__PURE__ */ jsx(Upload, { size: 16 }),
        " Escolher Ficheiro"
      ] })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error }),
    rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-600 mb-2", children: [
        "Ficheiro: ",
        /* @__PURE__ */ jsx("strong", { children: fileName })
      ] }),
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-600 mb-4", children: [
        rows.total,
        " linha",
        rows.total !== 1 ? "s" : "",
        " encontrada",
        rows.total !== 1 ? "s" : "",
        ", ",
        /* @__PURE__ */ jsx("strong", { children: rows.novos.length }),
        " novo",
        rows.novos.length !== 1 ? "s" : "",
        " centro",
        rows.novos.length !== 1 ? "s" : "",
        " a criar",
        rows.total !== rows.novos.length && ` (${rows.total - rows.novos.length} j\xE1 existente(s), ignorado(s))`,
        "."
      ] }),
      rows.novos.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 mb-4", children: "Nada de novo para importar." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-4 max-h-64 overflow-y-auto", children: rows.novos.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `px-3 py-2 text-sm ${i !== rows.novos.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsx("span", { className: "font-mono-data text-amber-700 font-semibold mr-2", children: r.codigo }),
        /* @__PURE__ */ jsx("span", { className: "text-stone-700", children: r.nome }),
        r.localizacao && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
          " \xB7 ",
          r.localizacao
        ] })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRows(null);
          setFileName("");
        }, className: "flex-1 py-2.5 rounded-lg border border-stone-300 text-stone-600 text-sm font-semibold hover:bg-stone-50", children: "Escolher outro ficheiro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirm, disabled: rows.novos.length === 0 || importing, className: "flex-1 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700 disabled:opacity-50", children: importing ? "A importar..." : `Importar ${rows.novos.length}` })
      ] })
    ] })
  ] });
}
function CenterModal({ data, onSave, onClose }) {
  const [nome, setNome] = useState(data?.nome || "");
  const [codigo, setCodigo] = useState(data?.codigo || "");
  const [localizacao, setLocalizacao] = useState(data?.localizacao || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!nome.trim()) return setError("Indique o nome do centro");
    if (!codigo.trim()) return setError("Indique o c\xF3digo do centro");
    onSave({ id: data?.id, nome: nome.trim(), codigo: codigo.trim().toUpperCase(), localizacao: localizacao.trim(), ativo: data?.ativo !== false });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Centro" : "Novo Centro", subtitle: "Centro de Produ\xE7\xE3o", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Nome do centro", children: /* @__PURE__ */ jsx("input", { value: nome, onChange: (e) => setNome(e.target.value), className: inputCls, placeholder: "Ex: Central de Vila Real" }) }),
    /* @__PURE__ */ jsxs(Field, { label: "C\xF3digo do centro", children: [
      /* @__PURE__ */ jsx("input", { value: codigo, onChange: (e) => setCodigo(e.target.value.toUpperCase().slice(0, 6)), className: `${inputCls} font-mono-data tracking-widest uppercase`, placeholder: "Ex: BB" }),
      /* @__PURE__ */ jsxs("span", { className: "block text-xs text-stone-400 mt-1", children: [
        "Usado para gerar o c\xF3digo dos artigos deste centro (ex: ",
        codigo || "BB",
        "01, ",
        codigo || "BB",
        "02, ... ",
        codigo || "BB",
        "90)"
      ] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Localiza\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: localizacao, onChange: (e) => setLocalizacao(e.target.value), className: inputCls, placeholder: "Ex: Vila Real" }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EliminarTudoModal({ titulo, aviso, count, palavraConfirmacao, onConfirm, onClose }) {
  const [texto, setTexto] = useState("");
  const confirmado = texto.trim().toUpperCase() === palavraConfirmacao;
  return /* @__PURE__ */ jsxs(Modal, { title: titulo, subtitle: "A\xE7\xE3o irrevers\xEDvel", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "bg-red-50 border border-red-200 rounded-lg p-3 mb-4 flex items-start gap-2", children: [
      /* @__PURE__ */ jsx(AlertTriangle, { size: 18, className: "text-red-600 shrink-0 mt-0.5" }),
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("p", { className: "text-sm text-red-700 font-semibold mb-1", children: [
          "Vai eliminar ",
          count,
          " registo",
          count !== 1 ? "s" : "",
          ". N\xE3o \xE9 poss\xEDvel desfazer esta a\xE7\xE3o."
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-xs text-red-600", children: aviso })
      ] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: `Para confirmar, escreva "${palavraConfirmacao}"`, children: /* @__PURE__ */ jsx("input", { value: texto, onChange: (e) => setTexto(e.target.value), className: `${inputCls} font-mono-data uppercase`, placeholder: palavraConfirmacao, autoFocus: true }) }),
    /* @__PURE__ */ jsx(
      "button",
      {
        onClick: onConfirm,
        disabled: !confirmado,
        className: "w-full py-3 rounded-lg bg-red-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-red-700 disabled:opacity-40 disabled:cursor-not-allowed",
        children: "Eliminar Definitivamente"
      }
    )
  ] });
}
function MixtureModal({ data, onSave, onClose }) {
  const [codigoManual, setCodigoManual] = useState(data?.codigoManual || "");
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [referenciaRelatorio, setReferenciaRelatorio] = useState(data?.referenciaRelatorio || "");
  const [dataExecucao, setDataExecucao] = useState(data?.dataExecucao || "");
  const [error, setError] = useState("");
  const validade = calcValidade(dataExecucao);
  const submit = () => {
    if (!codigoManual.trim()) return setError("Indique o c\xF3digo do artigo");
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o do artigo");
    onSave({
      id: data?.id,
      centroId: data.centroId,
      codigoManual: codigoManual.trim(),
      designacao: designacao.trim(),
      ativo: data?.ativo !== false,
      referenciaRelatorio: referenciaRelatorio.trim(),
      dataExecucao
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Artigo" : "Novo Artigo", subtitle: "Base de Dados de Artigos", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "C\xF3digo", children: /* @__PURE__ */ jsx("input", { value: codigoManual, onChange: (e) => setCodigoManual(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "Ex: BB01-02", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: "Ex: AC 14 bin 35/50" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Refer\xEAncia do Relat\xF3rio (Ensaio Tipo Inicial)", children: /* @__PURE__ */ jsx("input", { value: referenciaRelatorio, onChange: (e) => setReferenciaRelatorio(e.target.value), className: inputCls, placeholder: "Ex: ETI-2024-018" }) }),
    /* @__PURE__ */ jsx(Field, { label: "Data de Execu\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataExecucao, onChange: (e) => setDataExecucao(e.target.value), className: inputCls }) }),
    validade && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 -mt-2 mb-4", children: [
      "Validade calculada (",
      VALIDADE_ANOS,
      " anos): ",
      /* @__PURE__ */ jsx("strong", { className: "text-stone-700", children: formatDatePT(validade) })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm font-semibold bg-red-50 border border-red-200 rounded-lg px-3 py-2 mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function ProvenienciaModal({ data, onSave, onClose }) {
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [fornecedor, setFornecedor] = useState(data?.fornecedor || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o");
    if (!fornecedor.trim()) return setError("Indique o fornecedor");
    onSave({ id: data?.id, centroId: data.centroId, designacao: designacao.trim(), fornecedor: fornecedor.trim(), ativo: data?.ativo !== false });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Proveni\xEAncia" : "Nova Proveni\xEAncia", subtitle: "Proveni\xEAncias de Agregados", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: "Ex: Pedreira do Vale", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Fornecedor", children: /* @__PURE__ */ jsx("input", { value: fornecedor, onChange: (e) => setFornecedor(e.target.value), className: inputCls, placeholder: "Ex: Britas & Agregados, Lda." }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function ImportDiariasModal({ centroId, artigos, clientes, centrosCusto, onImport, onClose }) {
  const [grupos, setGrupos] = useState(null);
  const [erros, setErros] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const asDate = (v) => {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v ?? "").trim();
  };
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setGrupos(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const flat = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
        const dataInicioKey = headerKeys.find((k) => flat(k) === "datainicio");
        const dataFimKey = headerKeys.find((k) => flat(k) === "datafim");
        const artigoKey = headerKeys.find((k) => flat(k) === "artigo");
        const clienteKey = headerKeys.find((k) => flat(k) === "cliente");
        const centroCustoKey = headerKeys.find((k) => flat(k) === "centrodecusto" || flat(k) === "obra");
        const toneladasKey = headerKeys.find((k) => flat(k).startsWith("tonelada"));
        if (!dataInicioKey || !dataFimKey || !artigoKey || !clienteKey || !centroCustoKey || !toneladasKey) {
          setError('N\xE3o encontrei todas as colunas esperadas: "Data In\xEDcio", "Data Fim", "Artigo", "Cliente", "Centro de Custo" (ou "Obra") e "Toneladas".');
          setGrupos(null);
          return;
        }
        const errosLocais = [];
        const gruposMap = {};
        json.forEach((r, idx) => {
          const dataInicio = asDate(r[dataInicioKey]);
          const dataFim = asDate(r[dataFimKey]) || dataInicio;
          const artigoTxt = String(r[artigoKey] ?? "").trim();
          const clienteTxt = String(r[clienteKey] ?? "").trim();
          const centroCustoTxt = String(r[centroCustoKey] ?? "").trim();
          const toneladas = parseFloat(r[toneladasKey]);
          if (!dataInicio && !artigoTxt && !clienteTxt) return;
          const artigo = artigos.find((a) => a.designacao.trim().toLowerCase() === artigoTxt.toLowerCase());
          const cliente = clientes.find((c) => c.designacao.trim().toLowerCase() === clienteTxt.toLowerCase());
          const precisaCentroCusto = isSocorpenaCliente(cliente);
          const centroCusto = precisaCentroCusto ? centrosCusto.find((cc) => clientes.some((c) => c.id === cc.clienteId && isSocorpenaCliente(c)) && cc.designacao.trim().toLowerCase() === centroCustoTxt.toLowerCase()) : null;
          const problemas = [];
          if (!dataInicio) problemas.push("data in\xEDcio em falta");
          if (!artigo) problemas.push(`artigo "${artigoTxt}" n\xE3o encontrado neste centro`);
          if (!cliente) problemas.push(`cliente "${clienteTxt}" n\xE3o encontrado`);
          if (precisaCentroCusto && !centroCusto) problemas.push(`centro de custo "${centroCustoTxt}" n\xE3o encontrado para este cliente`);
          if (!toneladas || toneladas <= 0) problemas.push("toneladas inv\xE1lidas");
          if (problemas.length > 0) {
            errosLocais.push({ linha: idx + 2, motivo: problemas.join("; ") });
            return;
          }
          const key = `${dataInicio}|${dataFim}`;
          if (!gruposMap[key]) gruposMap[key] = { dataInicio, dataFim, linhas: [] };
          gruposMap[key].linhas.push({
            id: genId(),
            artigoId: artigo.id,
            artigoDesignacao: artigo.designacao,
            clienteId: cliente.id,
            centroCustoId: centroCusto ? centroCusto.id : "",
            toneladas: String(toneladas)
          });
        });
        const gruposFinal = Object.values(gruposMap);
        if (gruposFinal.length === 0 && errosLocais.length === 0) {
          setError("N\xE3o encontrei linhas com dados.");
          setGrupos(null);
          return;
        }
        setGrupos(gruposFinal);
        setErros(errosLocais);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setGrupos(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const totalLinhas = (grupos || []).reduce((s, g) => s + g.linhas.length, 0);
  const confirmImport = () => {
    setImporting(true);
    onImport(centroId, grupos);
  };
  const baixarModelo = () => {
    const exemploArtigo = artigos[0]?.designacao || "AC20 base 35/50 (MB)";
    const exemploCliente = clientes[0]?.designacao || "Nome do Cliente";
    const exemploObra = centrosCusto[0]?.designacao || "";
    const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
    const wsData = [
      ["Data In\xEDcio", "Data Fim", "Artigo", "Cliente", "Centro de Custo", "Toneladas"],
      [hoje, hoje, exemploArtigo, exemploCliente, exemploObra, 100]
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 32 }, { wch: 28 }, { wch: 28 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Di\xE1rias");
    XLSX.writeFile(wb, "modelo_importacao_diarias.xlsx");
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Di\xE1rias de Produ\xE7\xE3o", subtitle: "Ficheiro Excel", onClose, children: [
    !grupos && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Colunas esperadas: ",
        /* @__PURE__ */ jsx("strong", { children: "Data In\xEDcio" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Data Fim" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Artigo" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Cliente" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Centro de Custo" }),
        " (ou Obra) e ",
        /* @__PURE__ */ jsx("strong", { children: "Toneladas" }),
        ". Linhas com a mesma Data In\xEDcio/Fim juntam-se na mesma di\xE1ria. Artigo, Cliente e Centro de Custo t\xEAm de corresponder exatamente aos j\xE1 registados."
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: baixarModelo, className: "flex items-center gap-1.5 mb-4 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Descarregar Modelo"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    grupos && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        grupos.length,
        " di\xE1ria",
        grupos.length !== 1 ? "s" : "",
        " \xB7 ",
        totalLinhas,
        " linha",
        totalLinhas !== 1 ? "s" : "",
        " prontas a importar"
      ] }),
      erros.length > 0 && /* @__PURE__ */ jsxs("div", { className: "bg-red-50 border border-red-200 rounded-lg px-3 py-2 mb-3 text-xs text-red-700", children: [
        /* @__PURE__ */ jsxs("p", { className: "font-semibold mb-1", children: [
          erros.length,
          " linha",
          erros.length !== 1 ? "s" : "",
          " da folha foi",
          erros.length !== 1 ? "ram" : "",
          " ignorada",
          erros.length !== 1 ? "s" : "",
          ":"
        ] }),
        /* @__PURE__ */ jsx("ul", { className: "space-y-0.5 max-h-24 overflow-y-auto", children: erros.map((e, i) => /* @__PURE__ */ jsxs("li", { children: [
          "Linha ",
          e.linha,
          ": ",
          e.motivo
        ] }, i)) })
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-64 overflow-y-auto mb-4", children: grupos.map((g, i) => /* @__PURE__ */ jsxs("div", { className: `px-3 py-2 text-sm ${i !== grupos.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("p", { className: "font-semibold text-stone-800", children: [
          g.dataInicio,
          g.dataFim !== g.dataInicio ? ` \u2192 ${g.dataFim}` : ""
        ] }),
        g.linhas.map((l, j) => /* @__PURE__ */ jsxs("p", { className: "text-stone-500 text-xs pl-2", children: [
          l.artigoDesignacao,
          " \xB7 ",
          l.toneladas,
          " t"
        ] }, j))
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setGrupos(null);
          setErros([]);
          setFileName("");
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirmImport, disabled: importing || grupos.length === 0, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function FornecedorSearchSelect({ value, fornecedores, disabled, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const filtered = query ? fornecedores.filter((f) => matchesSearch(query, f.designacao, f.numero, f.nif)) : fornecedores;
  return /* @__PURE__ */ jsxs("div", { className: "relative", children: [
    /* @__PURE__ */ jsxs("div", { className: "relative", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none", size: 13 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: open ? query : value || "",
          onChange: (e) => {
            const v = e.target.value;
            setQuery(v);
            onChange(v);
            if (!open) setOpen(true);
          },
          onFocus: () => {
            if (!disabled) {
              setQuery(value || "");
              setOpen(true);
            }
          },
          onBlur: () => setTimeout(() => setOpen(false), 150),
          disabled,
          placeholder: "Pesquisar fornecedor (nome, n\xBA ou NIF)...",
          className: `${inputCls} pl-7 disabled:bg-stone-100`
        }
      )
    ] }),
    open && !disabled && /* @__PURE__ */ jsx("div", { className: "absolute z-20 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-stone-200 rounded-lg shadow-lg", children: fornecedores.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-sm text-stone-400", children: "Ainda n\xE3o h\xE1 fornecedores registados" }) : filtered.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-sm text-stone-400", children: "Nenhum fornecedor encontrado" }) : filtered.map((f) => /* @__PURE__ */ jsxs(
      "button",
      {
        type: "button",
        onMouseDown: () => {
          onChange(f.designacao);
          setOpen(false);
          setQuery("");
        },
        className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 flex items-center gap-2",
        children: [
          f.numero && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: f.numero }),
          /* @__PURE__ */ jsx("span", { className: "truncate text-stone-700", children: f.designacao })
        ]
      },
      f.id
    )) })
  ] });
}
function ClienteSearchSelect({ value, clientes, disabled, permitirVazio, placeholder, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const selected = clientes.find((c) => c.id === value);
  const filtered = query ? clientes.filter((c) => matchesSearch(query, c.designacao, c.numero, c.nif)) : clientes;
  return /* @__PURE__ */ jsxs("div", { className: "relative", children: [
    /* @__PURE__ */ jsxs("div", { className: "relative", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none", size: 13 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: open ? query : selected ? selected.designacao : "",
          onChange: (e) => {
            setQuery(e.target.value);
            if (!open) setOpen(true);
          },
          onFocus: () => {
            if (!disabled) {
              setQuery("");
              setOpen(true);
            }
          },
          onBlur: () => setTimeout(() => setOpen(false), 150),
          disabled,
          placeholder: placeholder || "Pesquisar cliente (nome, n\xBA ou NIF)...",
          className: `${inputCls} pl-7 disabled:bg-stone-100`
        }
      )
    ] }),
    open && !disabled && /* @__PURE__ */ jsxs("div", { className: "absolute z-20 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-stone-200 rounded-lg shadow-lg", children: [
      permitirVazio && /* @__PURE__ */ jsx(
        "button",
        {
          type: "button",
          onMouseDown: () => {
            onChange("");
            setOpen(false);
            setQuery("");
          },
          className: "w-full text-left px-3 py-2 text-sm text-stone-400 italic hover:bg-amber-50 border-b border-stone-100",
          children: "\u2014 Todos os Clientes \u2014"
        }
      ),
      filtered.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-sm text-stone-400", children: "Nenhum cliente encontrado" }) : filtered.map((c) => /* @__PURE__ */ jsxs(
        "button",
        {
          type: "button",
          onMouseDown: () => {
            onChange(c.id);
            setOpen(false);
            setQuery("");
          },
          className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 flex items-center gap-2",
          children: [
            c.numero && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: c.numero }),
            /* @__PURE__ */ jsx("span", { className: "truncate text-stone-700", children: c.designacao })
          ]
        },
        c.id
      ))
    ] })
  ] });
}
function CentrosAplicacaoSelect({ value, centers, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const isTodos = value === "todos";
  const selectedIds = Array.isArray(value) ? value : [];
  const selectedCenters = centers.filter((c) => selectedIds.includes(c.id));
  const disponiveis = centers.filter((c) => !selectedIds.includes(c.id));
  const filtrados = query ? disponiveis.filter((c) => matchesSearch(query, c.nome)) : disponiveis;
  const todosVisivel = !isTodos && (!query || matchesSearch(query, "todos os centros"));
  const escolherTodos = () => {
    onChange("todos");
    setQuery("");
    setOpen(false);
  };
  const escolherCentro = (id) => {
    onChange(isTodos ? [id] : [...selectedIds, id]);
    setQuery("");
  };
  const remover = (id) => onChange(selectedIds.filter((x) => x !== id));
  const removerTodos = () => onChange([]);
  return /* @__PURE__ */ jsxs("div", { className: "relative", children: [
    /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap gap-1.5 mb-1.5", children: [
      isTodos && /* @__PURE__ */ jsxs("span", { className: "inline-flex items-center gap-1 text-xs font-semibold bg-amber-100 text-amber-800 px-2 py-1 rounded-full", children: [
        "Todos os centros",
        /* @__PURE__ */ jsx("button", { type: "button", onClick: removerTodos, className: "hover:text-red-600", children: /* @__PURE__ */ jsx(X, { size: 12 }) })
      ] }),
      !isTodos && selectedCenters.map((c) => /* @__PURE__ */ jsxs("span", { className: "inline-flex items-center gap-1 text-xs font-semibold bg-stone-100 text-stone-700 px-2 py-1 rounded-full", children: [
        c.nome,
        /* @__PURE__ */ jsx("button", { type: "button", onClick: () => remover(c.id), className: "hover:text-red-600", children: /* @__PURE__ */ jsx(X, { size: 12 }) })
      ] }, c.id))
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "relative", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none", size: 13 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => {
            setQuery(e.target.value);
            if (!open) setOpen(true);
          },
          onFocus: () => setOpen(true),
          onBlur: () => setTimeout(() => setOpen(false), 150),
          disabled: isTodos,
          placeholder: isTodos ? "Aplica-se a todos os centros" : "Escreva para procurar um centro...",
          className: `${inputCls} pl-7 disabled:bg-stone-100`
        }
      )
    ] }),
    open && !isTodos && /* @__PURE__ */ jsxs("div", { className: "absolute z-20 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-stone-200 rounded-lg shadow-lg", children: [
      todosVisivel && /* @__PURE__ */ jsx(
        "button",
        {
          type: "button",
          onMouseDown: escolherTodos,
          className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 font-semibold text-amber-700 border-b border-stone-100",
          children: "\u2605 Todos os centros"
        }
      ),
      filtrados.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-sm text-stone-400", children: "Nenhum centro encontrado" }) : filtrados.map((c) => /* @__PURE__ */ jsx(
        "button",
        {
          type: "button",
          onMouseDown: () => escolherCentro(c.id),
          className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 text-stone-700",
          children: c.nome
        },
        c.id
      ))
    ] })
  ] });
}
function MaterialSearchSelect({ value, materiais, disabled, compact, alignRight, isAdmin, placeholder, semAbreviar, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const selected = materiais.find((m) => m.id === value);
  const rotulo = (m) => semAbreviar ? m?.fornecedor ? `${m.designacao} (${m.fornecedor})` : m?.designacao || "" : materialLabel(m);
  const q = query;
  const filtered = q ? materiais.filter((m) => matchesSearch(q, m.designacao, m.fornecedor, m.observacoes)) : materiais;
  const expandido = compact && open;
  const inputSizeCls = compact && !expandido ? "text-[11px] py-1.5 pl-6" : "pl-8 py-2.5 text-sm";
  return /* @__PURE__ */ jsx("div", { className: "relative", children: /* @__PURE__ */ jsxs("div", { className: expandido ? `absolute z-40 -top-2 ${alignRight ? "-right-2" : "-left-2"} w-96 bg-white rounded-xl shadow-2xl ring-2 ring-amber-400 p-2` : "relative", children: [
    /* @__PURE__ */ jsxs("div", { className: "relative", children: [
      /* @__PURE__ */ jsx(Search, { className: `absolute ${compact && !expandido ? "left-1.5" : "left-3"} top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none`, size: compact && !expandido ? 11 : 15 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: open ? query : selected ? rotulo(selected) : "",
          onChange: (e) => {
            setQuery(e.target.value);
            if (!open) setOpen(true);
          },
          onFocus: () => {
            if (!disabled) {
              setQuery("");
              setOpen(true);
            }
          },
          onBlur: () => setTimeout(() => setOpen(false), 150),
          disabled,
          placeholder: placeholder || (compact ? "Material..." : "Pesquisar material..."),
          className: `${inputCls} ${inputSizeCls} disabled:bg-stone-100`
        }
      )
    ] }),
    open && !disabled && /* @__PURE__ */ jsxs("div", { className: `${expandido ? "relative mt-1" : "absolute z-20 mt-1"} w-full max-h-64 overflow-y-auto bg-white border border-stone-200 rounded-lg shadow-lg`, children: [
      value && /* @__PURE__ */ jsx(
        "button",
        {
          type: "button",
          onMouseDown: () => {
            onChange("");
            setOpen(false);
            setQuery("");
          },
          className: "w-full text-left px-3 py-2 text-sm text-stone-400 italic hover:bg-red-50 hover:text-red-600 border-b border-stone-100",
          children: "\u2014 Vazio \u2014"
        }
      ),
      materiais.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-xs text-stone-400", children: "Sem materiais aplic\xE1veis a este centro" }) : filtered.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-xs text-stone-400", children: "Nenhum material encontrado" }) : filtered.map((m) => {
        const vigente = precoVigente(m);
        return /* @__PURE__ */ jsxs(
          "button",
          {
            type: "button",
            onMouseDown: () => {
              onChange(m.id);
              setOpen(false);
              setQuery("");
            },
            className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 flex items-center justify-between gap-2",
            children: [
              /* @__PURE__ */ jsxs("span", { className: "min-w-0 truncate", children: [
                /* @__PURE__ */ jsx("span", { className: "text-stone-700", children: rotulo(m) }),
                m.observacoes && /* @__PURE__ */ jsxs("span", { className: "block text-[11px] text-amber-700 truncate", children: [
                  "Obs: ",
                  m.observacoes
                ] })
              ] }),
              isAdmin && vigente && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs text-stone-400 shrink-0", children: [
                calcularPrecoFinal(vigente).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
                "\u20AC"
              ] })
            ]
          },
          m.id
        );
      })
    ] })
  ] }) });
}
function CentroCustoSearchSelect({ value, centrosCusto, disabled, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const selected = centrosCusto.find((cc) => cc.id === value);
  const filtered = query ? centrosCusto.filter((cc) => matchesSearch(query, cc.designacao, cc.codigo)) : centrosCusto;
  return /* @__PURE__ */ jsxs("div", { className: "relative", children: [
    /* @__PURE__ */ jsxs("div", { className: "relative", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none", size: 13 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: open ? query : selected ? selected.codigo ? `${selected.codigo} - ${selected.designacao}` : selected.designacao : "",
          onChange: (e) => {
            setQuery(e.target.value);
            if (!open) setOpen(true);
          },
          onFocus: () => {
            if (!disabled) {
              setQuery("");
              setOpen(true);
            }
          },
          onBlur: () => setTimeout(() => setOpen(false), 150),
          disabled,
          placeholder: "Pesquisar obra (nome ou n\xBA)...",
          className: `${inputCls} pl-7 disabled:bg-stone-100`
        }
      )
    ] }),
    open && !disabled && /* @__PURE__ */ jsx("div", { className: "absolute z-20 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-stone-200 rounded-lg shadow-lg", children: filtered.length === 0 ? /* @__PURE__ */ jsx("div", { className: "px-3 py-2 text-sm text-stone-400", children: "Nenhuma obra encontrada" }) : filtered.map((cc) => /* @__PURE__ */ jsxs(
      "button",
      {
        type: "button",
        onMouseDown: () => {
          onChange(cc.id);
          setOpen(false);
          setQuery("");
        },
        className: "w-full text-left px-3 py-2 text-sm hover:bg-amber-50 flex items-center gap-2",
        children: [
          cc.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: cc.codigo }),
          /* @__PURE__ */ jsx("span", { className: "truncate text-stone-700", children: cc.designacao })
        ]
      },
      cc.id
    )) })
  ] });
}
function EditarDataHoraHistoricoModal({ entrada, onSave, onClose }) {
  const dtLocal = isoToDatetimeLocal(entrada.data);
  const [dataParte, setDataParte] = useState(dtLocal.slice(0, 10) || "");
  const [horaParte, setHoraParte] = useState(dtLocal.slice(11, 16) || "00:00");
  const [error, setError] = useState("");
  const submit = () => {
    if (!dataParte) return setError("Indique a data");
    if (!horaParte) return setError("Indique a hora");
    const iso = datetimeLocalToIso(`${dataParte}T${horaParte}`);
    if (!iso) return setError("Essa data n\xE3o existe no calend\xE1rio \u2014 verifique o dia e o m\xEAs");
    onSave(iso);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Data e Hora", subtitle: "Hist\xF3rico de Registos", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Data", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataParte, onChange: (e) => setDataParte(e.target.value), className: inputCls, autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Hora", children: /* @__PURE__ */ jsx("input", { type: "time", value: horaParte, onChange: (e) => setHoraParte(e.target.value), className: inputCls }) })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EditarUtilizadorHistoricoModal({ users, valorAtualId, onSave, onClose }) {
  const [userId, setUserId] = useState(valorAtualId || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!userId) return setError("Escolha a pessoa que realmente fez este registo");
    onSave(userId);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Indicar Outra Pessoa", subtitle: "Corrigir quem fez este registo", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Pessoa", children: /* @__PURE__ */ jsxs("select", { value: userId, onChange: (e) => setUserId(e.target.value), className: inputCls, children: [
      /* @__PURE__ */ jsx("option", { value: "", children: "Escolha a pessoa..." }),
      users.map((u) => /* @__PURE__ */ jsx("option", { value: u.id, children: u.nome }, u.id))
    ] }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
const nomeParaAssinatura = (user, fallback) => {
  const composto = `${user?.primeiroNome || ""} ${user?.ultimoNome || ""}`.trim();
  return composto || user?.nome || fallback || "\u2014";
};
function AssinaturaDisplay({ nome, assinatura, className }) {
  if (assinatura) {
    return /* @__PURE__ */ jsx("img", { src: assinatura, alt: `Assinatura de ${nome}`, className: className || "h-9" });
  }
  return /* @__PURE__ */ jsx("span", { className: `text-blue-700 ${className || "text-xl"}`, style: { fontFamily: "'Brush Script MT', 'Segoe Script', 'Lucida Handwriting', cursive" }, children: nome || "\u2014" });
}
function DiariaModal({ data, artigos, clientes, centrosCusto, diarias, avarias, rececoes, materiais, consumiveis, center, logotipo, isAdmin, onDeleteHistorico, onEditHistoricoUtilizador, onEditHistoricoData, users, readOnly, onSave, onClose }) {
  const emptyMistura = () => ({ id: genId(), artigoId: "", toneladas: "" });
  const emptyGrupo = () => ({ id: genId(), clienteId: "", centroCustoId: "", misturas: [emptyMistura()] });
  const [dataInicio, setDataInicio] = useState(data?.dataInicio || "");
  const [dataFim, setDataFim] = useState(data?.dataFim || data?.dataInicio || "");
  const [grupos, setGrupos] = useState(() => {
    if (!data?.linhas?.length) return [emptyGrupo()];
    const lista = [];
    data.linhas.forEach((l) => {
      let grupo = lista.find((g) => g.clienteId === l.clienteId && g.centroCustoId === l.centroCustoId);
      if (!grupo) {
        grupo = { id: genId(), clienteId: l.clienteId, centroCustoId: l.centroCustoId, misturas: [] };
        lista.push(grupo);
      }
      grupo.misturas.push({ id: l.id || genId(), artigoId: l.artigoId, toneladas: l.toneladas });
    });
    return lista;
  });
  const [incidenciasList, setIncidenciasList] = useState(() => normalizeIncidencias(data?.incidencias));
  const [observacoes, setObservacoes] = useState(data?.observacoes || "");
  const [error, setError] = useState("");
  const [entradaAssinaturaEditar, setEntradaAssinaturaEditar] = useState(null);
  const [entradaDataEditar, setEntradaDataEditar] = useState(null);
  const artigosOrdenados = [...artigos].sort((a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true, sensitivity: "base" }));
  const usaCentroCusto = (clienteId) => isSocorpenaCliente(clientes.find((x) => x.id === clienteId));
  const rececoesDoDia = (rececoes || []).filter((r) => r.data && dataInicio && r.data === dataInicio);
  const historicoDiariaOrdenado = [...data?.historico || []].sort((a, b) => (a.data || "").localeCompare(b.data || ""));
  const nomeProdutoRececao = (r) => {
    const lista = r.categoria === "consumiveis" ? consumiveis || [] : materiais || [];
    return lista.find((p) => p.id === r.produtoId)?.designacao || "\u2014";
  };
  const fornecedorRececao = (r) => {
    const lista = r.categoria === "consumiveis" ? consumiveis || [] : materiais || [];
    return lista.find((p) => p.id === r.produtoId)?.fornecedor || "";
  };
  const unidadeRececao = (r) => r.categoria === "consumiveis" ? "L" : "t";
  const categoriaLabelRececao = { agregados: "Agregado", betumes: "Betume", consumiveis: "Combust\xEDvel" };
  const updateGrupo = (grupoId, field, value) => {
    setGrupos(grupos.map((g) => {
      if (g.id !== grupoId) return g;
      const updated = { ...g, [field]: value };
      if (field === "clienteId") updated.centroCustoId = "";
      return updated;
    }));
  };
  const updateMistura = (grupoId, misturaId, field, value) => {
    setGrupos(grupos.map((g) => g.id !== grupoId ? g : {
      ...g,
      misturas: g.misturas.map((m) => m.id === misturaId ? { ...m, [field]: value } : m)
    }));
  };
  const addMistura = (grupoId) => setGrupos(grupos.map((g) => g.id === grupoId ? { ...g, misturas: [...g.misturas, emptyMistura()] } : g));
  const removeMistura = (grupoId, misturaId) => setGrupos(grupos.map((g) => g.id !== grupoId ? g : {
    ...g,
    misturas: g.misturas.length > 1 ? g.misturas.filter((m) => m.id !== misturaId) : g.misturas
  }));
  const addGrupo = () => setGrupos([...grupos, emptyGrupo()]);
  const removeGrupo = (grupoId) => setGrupos(grupos.length > 1 ? grupos.filter((g) => g.id !== grupoId) : grupos);
  const addIncidencia = () => setIncidenciasList([...incidenciasList, { id: genId(), descricao: "", resolucaoData: "", resolucaoDescricao: "" }]);
  const removeIncidencia = (id) => setIncidenciasList(incidenciasList.filter((i) => i.id !== id));
  const updateIncidencia = (id, descricao) => setIncidenciasList(incidenciasList.map((i) => i.id === id ? { ...i, descricao } : i));
  const linhasFlat = grupos.flatMap((g) => g.misturas.map((m) => ({ ...m, clienteId: g.clienteId, centroCustoId: g.centroCustoId })));
  const totalToneladas = linhasFlat.reduce((s, l) => s + (parseFloat(l.toneladas) || 0), 0);
  const turno = dataInicio && dataFim ? dataFim === dataInicio ? "Diurno" : dataFim > dataInicio ? "Noturno" : "" : "";
  const dataFimInvalida = !!(dataInicio && dataFim && (dataFim < dataInicio || !isDataCalendarioValida(dataInicio) || !isDataCalendarioValida(dataFim)));
  const anoRef = dataInicio ? new Date(dataInicio).getFullYear() : (/* @__PURE__ */ new Date()).getFullYear();
  const acumuladoAnual = (diarias || []).filter((d) => d.id !== data?.id && d.dataInicio && new Date(d.dataInicio).getFullYear() === anoRef && d.dataInicio <= dataInicio).reduce((s, d) => s + (d.linhas || []).reduce((s2, l) => s2 + (parseFloat(l.toneladas) || 0), 0), 0) + totalToneladas;
  const dataReferenciaPendentes = dataInicio || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  const diasAberto = (dIni) => Math.max(0, Math.round((new Date(dataReferenciaPendentes) - new Date(dIni)) / 864e5));
  const pendentes = [
    ...(diarias || []).filter((d) => d.id !== data?.id).flatMap((d) => normalizeIncidencias(d.incidencias).filter((inc) => !inc.resolucaoData).map((inc) => ({ key: `diaria-${d.id}-${inc.id}`, data: d.dataInicio, descricao: inc.descricao }))),
    ...(avarias || []).filter((a) => !a.resolucaoData).map((a) => ({ key: a.id, data: a.data, descricao: a.descricao }))
  ].sort((x, y) => (x.data || "").localeCompare(y.data || ""));
  const [avisos, setAvisos] = useState([]);
  const calcularAvisos = () => {
    const lista = [];
    if (!data?.id) {
      const hoje = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
      const diffDias = Math.round(Math.abs(/* @__PURE__ */ new Date(`${dataInicio}T00:00:00Z`) - /* @__PURE__ */ new Date(`${hoje}T00:00:00Z`)) / 864e5);
      if (diffDias > 1) lista.push(`A data de in\xEDcio (${formatDatePT(dataInicio)}) \xE9 diferente da data de hoje em ${diffDias} dias \u2014 confirme que n\xE3o h\xE1 engano.`);
    }
    const conflito = (diarias || []).some((d) => d.id !== data?.id && d.centroId === data.centroId && d.dataInicio === dataInicio && d.dataFim === dataFim);
    if (conflito) lista.push("J\xE1 existe uma di\xE1ria registada com este per\xEDodo (in\xEDcio e fim) neste centro.");
    return lista;
  };
  const executarGravacao = () => {
    const linhasFinal = grupos.flatMap((g) => g.misturas.map((m) => ({
      id: m.id,
      artigoId: m.artigoId,
      clienteId: g.clienteId,
      centroCustoId: usaCentroCusto(g.clienteId) ? g.centroCustoId : "",
      toneladas: m.toneladas,
      artigoDesignacao: artigos.find((a) => a.id === m.artigoId)?.designacao || ""
    })));
    const incidenciasFinal = incidenciasList.map((i) => ({ ...i, descricao: i.descricao.trim() })).filter((i) => i.descricao);
    onSave({ id: data?.id, centroId: data.centroId, dataInicio, dataFim, turno, linhas: linhasFinal, incidencias: incidenciasFinal, observacoes: observacoes.trim() });
  };
  const submit = () => {
    if (!dataInicio) return setError("Indique a data de in\xEDcio");
    if (!isDataCalendarioValida(dataInicio)) return setError("A data de in\xEDcio n\xE3o existe no calend\xE1rio \u2014 verifique o dia e o m\xEAs");
    if (!dataFim) return setError("Indique a data de fim");
    if (!isDataCalendarioValida(dataFim)) return setError("A data de fim n\xE3o existe no calend\xE1rio \u2014 verifique o dia e o m\xEAs");
    if (dataFim < dataInicio) return setError("A data de fim n\xE3o pode ser anterior \xE0 data de in\xEDcio");
    for (const g of grupos) {
      if (!g.clienteId) return setError("Escolha o cliente em todos os grupos");
      if (usaCentroCusto(g.clienteId) && !g.centroCustoId) return setError("Escolha o centro de custo em todos os grupos");
      for (const m of g.misturas) {
        if (!m.artigoId) return setError("Escolha o artigo em todas as misturas");
        if (!m.toneladas || parseFloat(m.toneladas) <= 0) return setError("Indique as toneladas em todas as misturas");
      }
    }
    const novosAvisos = calcularAvisos();
    if (novosAvisos.length > 0) {
      setAvisos(novosAvisos);
      setError("");
      return;
    }
    executarGravacao();
  };
  const exportarPDF = () => {
    const linhas = linhasFlat;
    const nomeCliente = (id) => clientes.find((c) => c.id === id)?.designacao || "\u2014";
    const nomeObra = (id) => {
      const cc = centrosCusto.find((x) => x.id === id);
      if (!cc) return "";
      return cc.codigo ? `${cc.codigo} - ${cc.designacao}` : cc.designacao;
    };
    const clientesAgrupados = [];
    linhas.forEach((l) => {
      let cg = clientesAgrupados.find((x) => x.clienteId === l.clienteId);
      if (!cg) {
        cg = { clienteId: l.clienteId, obras: [] };
        clientesAgrupados.push(cg);
      }
      let og = cg.obras.find((x) => x.centroCustoId === l.centroCustoId);
      if (!og) {
        og = { centroCustoId: l.centroCustoId, misturas: [] };
        cg.obras.push(og);
      }
      og.misturas.push(l);
    });
    const BORDA_MISTURA = "border-bottom:1px solid #f5f5f4;";
    const BORDA_OBRA = "border-bottom:1px solid #a8a29e;";
    const BORDA_CLIENTE = "border-bottom:2px solid #292524;";
    const linhasHtml = clientesAgrupados.flatMap((cg) => {
      const clienteRowspan = cg.obras.reduce((s, og) => s + og.misturas.length, 0);
      let clienteCellPendente = `<td rowspan="${clienteRowspan}" style="vertical-align:top; border-right:1px solid #d6d3d1; font-weight:600; ${BORDA_CLIENTE}">${nomeCliente(cg.clienteId)}</td>`;
      return cg.obras.flatMap((og, ogIdx) => {
        const ultimaObraDoCliente = ogIdx === cg.obras.length - 1;
        return og.misturas.map((m, mIdx) => {
          const ultimaMisturaDaObra = mIdx === og.misturas.length - 1;
          const borda = ultimaMisturaDaObra && ultimaObraDoCliente ? BORDA_CLIENTE : ultimaMisturaDaObra ? BORDA_OBRA : BORDA_MISTURA;
          const clienteCell = clienteCellPendente;
          clienteCellPendente = "";
          const obraCell = mIdx === 0 ? `<td rowspan="${og.misturas.length}" style="vertical-align:top; border-right:1px solid #f5f5f4; ${ultimaObraDoCliente ? BORDA_CLIENTE : BORDA_OBRA}">${usaCentroCusto(cg.clienteId) ? nomeObra(og.centroCustoId) || "\u2014" : "\u2014"}</td>` : "";
          return `
      <tr>
        ${clienteCell}${obraCell}
        <td style="${borda}">${artigos.find((a2) => a2.id === m.artigoId)?.designacao || "\u2014"}</td>
        <td style="text-align:right; ${borda}">${parseFloat(m.toneladas || 0).toLocaleString("pt-PT")} t</td>
      </tr>`;
        });
      });
    }).join("");
    const incidenciasHtml = incidenciasList.filter((i) => i.descricao?.trim()).map((i) => `
      <li>${i.descricao}${i.resolucaoData ? ` \u2014 resolvido em ${formatDatePT(i.resolucaoData)}${i.resolucaoDescricao ? `: ${i.resolucaoDescricao}` : ""}` : " (por resolver)"}</li>`).join("");
    const pendentesHtml = pendentes.map((p) => `
      <li>${p.descricao} \u2014 <strong>${diasAberto(p.data)} dia${diasAberto(p.data) !== 1 ? "s" : ""} em aberto</strong> (desde ${formatDatePT(p.data)})</li>`).join("");
    const assinaturaHtml = historicoDiariaOrdenado.map((h, i) => {
      const userDoRegisto = (users || []).find((u) => u.id === h.utilizadorId);
      const nomeAssinatura = nomeParaAssinatura(userDoRegisto, h.utilizador);
      const imgAssinatura = userDoRegisto?.assinatura;
      const assinaturaVisual = imgAssinatura ? `<img src="${imgAssinatura}" style="height:32px; vertical-align:middle;" />` : `<span style="font-family:'Brush Script MT','Segoe Script','Lucida Handwriting',cursive; font-size:20px; color:#1d4ed8;">${nomeAssinatura}</span>`;
      return `
      <div style="display:flex; align-items:center; gap:12px; padding:4px 0; font-size:12px;">
        <span style="width:90px; color:#78716c; text-transform:uppercase; font-size:10px;">${i === 0 ? "Executado por:" : "Editado por:"}</span>
        ${assinaturaVisual}
        <span style="color:#78716c;">${formatDateTimePT(h.data)}</span>
      </div>`;
    }).join("");
    const rececoesHtml = rececoesDoDia.map((r) => `
      <tr>
        <td>${formatDatePT(r.data)}</td>
        <td>${categoriaLabelRececao[r.categoria] || "\u2014"}</td>
        <td>${nomeProdutoRececao(r)}</td>
        <td>${fornecedorRececao(r) || "\u2014"}</td>
        <td style="text-align:right">${parseFloat(r.quantidade || 0).toLocaleString("pt-PT")} ${unidadeRececao(r)}</td>
      </tr>`).join("");
    const html = `<!DOCTYPE html>
      <html><head><meta charset="utf-8"><title>Di\xE1ria de Produ\xE7\xE3o \u2014 ${formatDatePT(dataInicio)}</title>
      <style>
        body { font-family: 'Arial Narrow', Arial, sans-serif; color: #1c1917; padding: 32px; max-width: 800px; margin: 0 auto; }
        h1 { font-size: 20px; margin-bottom: 2px; }
        .sub { color: #78716c; font-size: 12px; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 20px; }
        h2 { font-size: 14px; text-transform: uppercase; letter-spacing: 0.05em; color: #57534e; border-bottom: 1px solid #e7e5e4; padding-bottom: 4px; margin-top: 28px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 8px; }
        th { text-align: left; font-size: 11px; text-transform: uppercase; color: #57534e; background: #f5f5f4; padding: 6px 4px; border-bottom: 1px solid #d6d3d1; }
        td { padding: 6px 4px; border-bottom: 1px solid #f5f5f4; }
        .total { font-weight: bold; text-align: right; padding: 6px 4px; font-size: 14px; background: #f5f5f4; text-transform: uppercase; letter-spacing: 0.05em; }
        ul { font-size: 13px; padding-left: 18px; }
        .meta { font-size: 13px; color: #44403c; margin-bottom: 4px; }
        .aviso { margin-top: 36px; padding: 10px 14px; background: #fef3c7; border-radius: 8px; font-size: 12px; color: #92400e; }
        .logo { max-height: 60px; margin-bottom: 12px; }
        @media print { .aviso { display: none; } }
      </style></head>
      <body>
        ${logotipo ? `<img class="logo" src="${logotipo}" alt="Log\xF3tipo" />` : ""}
        <h1>Di\xE1ria de Produ\xE7\xE3o</h1>
        <p class="sub">${center?.nome || ""}</p>
        <p class="meta"><strong>Per\xEDodo:</strong> ${formatDatePT(dataInicio)} \u2192 ${formatDatePT(dataFim)} \xB7 <strong>Turno:</strong> ${turno || "\u2014"}</p>

        <h2>Produ\xE7\xE3o</h2>
        <table>
          <tr><th>Cliente</th><th>Centro de Custo</th><th>Artigo</th><th style="text-align:right">Toneladas</th></tr>
          ${linhasHtml}
          <tr><td colspan="3" class="total">Total do dia</td><td class="total" style="text-align:right">${totalToneladas.toLocaleString("pt-PT")} t</td></tr>
        </table>
        <p class="meta" style="text-align:right; margin-top:4px;">Acumulado do ano (${anoRef}): <strong>${acumuladoAnual.toLocaleString("pt-PT")} t</strong></p>

        ${rececoesDoDia.length > 0 ? `
        <h2>Rece\xE7\xE3o de Materiais Constituintes</h2>
        <table>
          <tr><th>Data</th><th>Categoria</th><th>Produto</th><th>Fornecedor</th><th style="text-align:right">Quantidade</th></tr>
          ${rececoesHtml}
        </table>` : ""}

        ${incidenciasHtml ? `
        <h2>Incid\xEAncias / Avarias</h2>
        <ul>${incidenciasHtml}</ul>` : ""}

        ${pendentesHtml ? `
        <h2>Pendentes de Resolu\xE7\xE3o</h2>
        <ul>${pendentesHtml}</ul>` : ""}

        ${observacoes.trim() ? `
        <h2>Observa\xE7\xF5es</h2>
        <p class="meta">${observacoes.trim().replace(/\n/g, "<br>")}</p>` : ""}

        ${assinaturaHtml ? `
        <h2>Assinatura \u2014 Hist\xF3rico de Registos</h2>
        ${assinaturaHtml}` : ""}

        <p class="aviso">Para guardar como PDF: abra este ficheiro no navegador e use Ficheiro \u2192 Imprimir \u2192 Guardar como PDF (ou Ctrl/Cmd+P).</p>
      </body></html>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `diaria-producao-${dataInicio || "sem-data"}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? readOnly ? "Di\xE1ria de Produ\xE7\xE3o" : "Editar Di\xE1ria de Produ\xE7\xE3o" : "Nova Di\xE1ria de Produ\xE7\xE3o", subtitle: "Di\xE1ria de Produ\xE7\xE3o", onClose, fullscreen: true, children: [
    data?.id && /* @__PURE__ */ jsx("div", { className: "flex justify-end mb-3", children: /* @__PURE__ */ jsxs("button", { onClick: exportarPDF, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
      /* @__PURE__ */ jsx(FileText, { size: 15 }),
      " Exportar Documento"
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap items-end gap-4 mb-4 pb-4 border-b border-stone-200", children: [
      /* @__PURE__ */ jsx("div", { className: "w-44", children: /* @__PURE__ */ jsx(Field, { label: "Data In\xEDcio", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataInicio, onChange: (e) => setDataInicio(e.target.value), disabled: readOnly, className: `${inputCls} disabled:bg-stone-100` }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-44", children: /* @__PURE__ */ jsx(Field, { label: "Data Fim", children: /* @__PURE__ */ jsx(
        "input",
        {
          type: "date",
          value: dataFim,
          onChange: (e) => setDataFim(e.target.value),
          disabled: readOnly,
          className: `${inputCls} disabled:bg-stone-100 ${dataFimInvalida ? "border-red-400 focus:ring-red-400 focus:border-red-400" : ""}`
        }
      ) }) }),
      turno && /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-4", children: [
        /* @__PURE__ */ jsx("span", { className: "text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Turno:" }),
        /* @__PURE__ */ jsx("span", { className: `text-xs font-semibold uppercase tracking-wide px-2 py-0.5 rounded-full ${turno === "Diurno" ? "bg-amber-100 text-amber-700" : "bg-slate-700 text-slate-100"}`, children: turno })
      ] }),
      dataFimInvalida && /* @__PURE__ */ jsx("span", { className: "block text-xs text-red-600 mb-4", children: "A data de fim n\xE3o pode ser anterior \xE0 data de in\xEDcio" })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Produ\xE7\xE3o por Cliente" }),
    /* @__PURE__ */ jsx("div", { className: "space-y-4 mb-3", children: grupos.map((g) => {
      const centrosCustoDoCliente = usaCentroCusto(g.clienteId) ? centrosCusto.filter((cc) => clientes.some((c) => c.id === cc.clienteId && isSocorpenaCliente(c))) : centrosCusto.filter((cc) => cc.clienteId === g.clienteId);
      const subtotalGrupo = g.misturas.reduce((s, m) => s + (parseFloat(m.toneladas) || 0), 0);
      return /* @__PURE__ */ jsxs("div", { className: "border border-stone-300 rounded-lg p-3 bg-white relative", children: [
        !readOnly && grupos.length > 1 && /* @__PURE__ */ jsx("button", { onClick: () => removeGrupo(g.id), className: "absolute top-2 right-2 text-stone-400 hover:text-red-600", title: "Remover cliente", children: /* @__PURE__ */ jsx(X, { size: 15 }) }),
        /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-2 mb-3 pr-6", children: [
          /* @__PURE__ */ jsx(ClienteSearchSelect, { value: g.clienteId, clientes, disabled: readOnly, onChange: (v) => updateGrupo(g.id, "clienteId", v) }),
          usaCentroCusto(g.clienteId) && /* @__PURE__ */ jsx(CentroCustoSearchSelect, { value: g.centroCustoId, centrosCusto: centrosCustoDoCliente, disabled: readOnly, onChange: (v) => updateGrupo(g.id, "centroCustoId", v) })
        ] }),
        /* @__PURE__ */ jsx("div", { className: "space-y-2", children: g.misturas.map((m) => /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap items-center gap-2", children: [
          /* @__PURE__ */ jsx("div", { className: "w-64 shrink-0", children: /* @__PURE__ */ jsx(
            MaterialSearchSelect,
            {
              value: m.artigoId,
              materiais: artigosOrdenados.map((a) => ({ ...a, fornecedor: a.codigo })),
              placeholder: "Mistura (f\xF3rmula)...",
              disabled: readOnly,
              onChange: (v) => updateMistura(g.id, m.id, "artigoId", v),
              semAbreviar: true,
              compact: true
            }
          ) }),
          /* @__PURE__ */ jsx(
            "input",
            {
              value: m.toneladas,
              onChange: (e) => updateMistura(g.id, m.id, "toneladas", e.target.value),
              type: "number",
              step: "0.01",
              min: "0",
              disabled: readOnly,
              className: `${inputCls} font-mono-data disabled:bg-stone-100 w-32 shrink-0`,
              placeholder: "Toneladas"
            }
          ),
          !readOnly && g.misturas.length > 1 && /* @__PURE__ */ jsx("button", { onClick: () => removeMistura(g.id, m.id), className: "text-stone-400 hover:text-red-600 shrink-0", children: /* @__PURE__ */ jsx(X, { size: 15 }) })
        ] }, m.id)) }),
        !readOnly && /* @__PURE__ */ jsxs("button", { onClick: () => addMistura(g.id), className: "w-full mt-2 py-1.5 rounded-lg border border-dashed border-stone-300 text-stone-500 text-xs font-semibold hover:bg-stone-50 flex items-center justify-center gap-1.5", children: [
          /* @__PURE__ */ jsx(Plus, { size: 13 }),
          " Adicionar mistura para este cliente"
        ] }),
        g.misturas.length > 1 && /* @__PURE__ */ jsxs("p", { className: "text-right text-xs text-stone-500 mt-2", children: [
          "Subtotal: ",
          /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-700", children: [
            subtotalGrupo.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
            " t"
          ] })
        ] })
      ] }, g.id);
    }) }),
    !readOnly && /* @__PURE__ */ jsxs("button", { onClick: addGrupo, className: "w-full mb-4 py-2 rounded-lg border border-dashed border-amber-400 text-amber-700 text-sm font-semibold hover:bg-amber-50 flex items-center justify-center gap-1.5", children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Adicionar cliente"
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-2.5 mb-2", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-amber-800", children: "Total da di\xE1ria" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-amber-800", children: [
        totalToneladas.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " t"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 mb-6", children: [
      "Acumulado do ano (",
      anoRef,
      "): ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-700", children: [
        acumuladoAnual.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " t"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "pt-4 border-t border-stone-200", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Incid\xEAncias / Avarias" }),
      /* @__PURE__ */ jsx("div", { className: "space-y-2 mb-2", children: incidenciasList.map((inc) => {
        const resolvida = !!inc.resolucaoData;
        return /* @__PURE__ */ jsxs("div", { className: "border border-stone-200 rounded-lg p-2.5 bg-stone-50/60 relative", children: [
          !readOnly && /* @__PURE__ */ jsx("button", { onClick: () => removeIncidencia(inc.id), className: "absolute top-2 right-2 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(X, { size: 14 }) }),
          /* @__PURE__ */ jsx(
            "textarea",
            {
              value: inc.descricao,
              onChange: (e) => updateIncidencia(inc.id, e.target.value),
              disabled: readOnly,
              rows: 2,
              spellCheck: "true",
              lang: "pt-PT",
              className: `${inputCls} disabled:bg-stone-100 ${!readOnly ? "pr-7" : ""}`,
              placeholder: "Descreva a incid\xEAncia ou avaria..."
            }
          ),
          inc.descricao.trim() && /* @__PURE__ */ jsx("span", { className: `inline-flex mt-1.5 text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full ${resolvida ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"}`, children: resolvida ? `Resolvido em ${formatDatePT(inc.resolucaoData)}` : "Por resolver" })
        ] }, inc.id);
      }) }),
      !readOnly && /* @__PURE__ */ jsxs("button", { onClick: addIncidencia, className: "w-full mb-4 py-2 rounded-lg border border-dashed border-stone-300 text-stone-500 text-sm font-semibold hover:bg-stone-50 flex items-center justify-center gap-1.5", children: [
        /* @__PURE__ */ jsx(Plus, { size: 15 }),
        " Adicionar incid\xEAncia"
      ] }),
      pendentes.length > 0 && /* @__PURE__ */ jsxs("div", { className: "mb-4", children: [
        /* @__PURE__ */ jsxs("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: [
          "Pendentes de Resolu\xE7\xE3o (",
          pendentes.length,
          ")"
        ] }),
        /* @__PURE__ */ jsx("div", { className: "bg-red-50/60 border border-red-100 rounded-lg overflow-hidden", children: pendentes.map((p, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2 text-sm ${i !== pendentes.length - 1 ? "border-b border-red-100" : ""}`, children: [
          /* @__PURE__ */ jsx("span", { className: "text-stone-700 truncate", children: p.descricao }),
          /* @__PURE__ */ jsxs("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-red-100 text-red-700 shrink-0 ml-2", children: [
            diasAberto(p.data),
            " dia",
            diasAberto(p.data) !== 1 ? "s" : ""
          ] })
        ] }, p.key)) })
      ] }),
      rececoesDoDia.length > 0 && /* @__PURE__ */ jsxs("div", { className: "mb-4", children: [
        /* @__PURE__ */ jsxs("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: [
          "Rece\xE7\xE3o de Materiais Constituintes (",
          rececoesDoDia.length,
          ")"
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-2", children: "Compilado automaticamente a partir das rece\xE7\xF5es registadas neste centro, no per\xEDodo desta di\xE1ria." }),
        /* @__PURE__ */ jsx("div", { className: "bg-white border border-stone-200 rounded-lg overflow-hidden", children: rececoesDoDia.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `px-3 py-2 text-sm ${i !== rececoesDoDia.length - 1 ? "border-b border-stone-100" : ""}`, children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 min-w-0 flex-wrap", children: [
            /* @__PURE__ */ jsx("span", { className: "text-stone-400 text-xs shrink-0", children: formatDatePT(r.data) }),
            /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-stone-100 text-stone-600 shrink-0", children: categoriaLabelRececao[r.categoria] || "\u2014" })
          ] }),
          /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between gap-2 mt-0.5", children: [
            /* @__PURE__ */ jsxs("span", { className: "text-stone-700 truncate", children: [
              nomeProdutoRececao(r),
              fornecedorRececao(r) && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
                " \xB7 ",
                fornecedorRececao(r)
              ] })
            ] }),
            /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-amber-700 font-semibold shrink-0", children: [
              parseFloat(r.quantidade || 0).toLocaleString("pt-PT"),
              " ",
              unidadeRececao(r)
            ] })
          ] })
        ] }, r.id)) })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "pt-4 border-t border-stone-200", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Observa\xE7\xF5es" }),
      /* @__PURE__ */ jsx("div", { className: "bg-stone-50 border border-stone-200 rounded-lg p-3 mb-6", children: /* @__PURE__ */ jsx(
        "textarea",
        {
          value: observacoes,
          onChange: (e) => setObservacoes(e.target.value),
          disabled: readOnly,
          rows: 4,
          spellCheck: "true",
          lang: "pt-PT",
          className: `${inputCls} disabled:bg-stone-100 bg-white`,
          placeholder: "Notas gerais sobre esta di\xE1ria..."
        }
      ) })
    ] }),
    data?.id && /* @__PURE__ */ jsxs("div", { className: "pt-4 border-t border-stone-200 mb-6", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Assinatura \u2014 Hist\xF3rico de Registos" }),
      historicoDiariaOrdenado.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400", children: "Ainda sem registos." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: historicoDiariaOrdenado.map((h, i) => {
        const userDoRegisto = (users || []).find((u) => u.id === h.utilizadorId);
        const ultimo = i === historicoDiariaOrdenado.length - 1 && historicoDiariaOrdenado.length > 1;
        return /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== historicoDiariaOrdenado.length - 1 ? "border-b border-stone-100" : ""} ${ultimo ? "bg-amber-50/40" : ""}`, children: [
          /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3 min-w-0", children: [
            /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide text-stone-400 shrink-0 w-20", children: i === 0 ? "Executado por:" : "Editado por:" }),
            /* @__PURE__ */ jsx(AssinaturaDisplay, { nome: nomeParaAssinatura(userDoRegisto, h.utilizador), assinatura: userDoRegisto?.assinatura, className: "h-8" }),
            /* @__PURE__ */ jsx("span", { className: "text-stone-500 text-xs shrink-0", children: formatDateTimePT(h.data) })
          ] }),
          /* @__PURE__ */ jsx("div", { className: "flex items-center gap-2 shrink-0", children: isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
            /* @__PURE__ */ jsx("button", { onClick: () => setEntradaAssinaturaEditar(h), title: "Indicar outra pessoa", className: "p-1 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(Pencil, { size: 13 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => setEntradaDataEditar(h), title: "Editar data e hora", className: "p-1 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(History, { size: 13 }) }),
            /* @__PURE__ */ jsx("button", { onClick: () => onDeleteHistorico(data.id, h.id), title: "Apagar registo", className: "p-1 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 13 }) })
          ] }) })
        ] }, h.id);
      }) })
    ] }),
    entradaAssinaturaEditar && /* @__PURE__ */ jsx(
      EditarUtilizadorHistoricoModal,
      {
        users: users || [],
        valorAtualId: entradaAssinaturaEditar.utilizadorId,
        onSave: (novoUserId) => {
          onEditHistoricoUtilizador(data.id, entradaAssinaturaEditar.id, novoUserId);
          setEntradaAssinaturaEditar(null);
        },
        onClose: () => setEntradaAssinaturaEditar(null)
      }
    ),
    entradaDataEditar && /* @__PURE__ */ jsx(
      EditarDataHoraHistoricoModal,
      {
        entrada: entradaDataEditar,
        onSave: (novaIso) => {
          onEditHistoricoData(data.id, entradaDataEditar.id, novaIso);
          setEntradaDataEditar(null);
        },
        onClose: () => setEntradaDataEditar(null)
      }
    ),
    avisos.length > 0 && /* @__PURE__ */ jsxs("div", { className: "bg-amber-50 border border-amber-200 rounded-lg p-3 mb-3", children: [
      avisos.map((a, i) => /* @__PURE__ */ jsxs("p", { className: "text-sm text-amber-800 flex items-start gap-1.5 mb-1 last:mb-0", children: [
        /* @__PURE__ */ jsx(AlertTriangle, { size: 15, className: "shrink-0 mt-0.5" }),
        " ",
        a
      ] }, i)),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2 mt-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setAvisos([]);
          executarGravacao();
        }, className: "px-3 py-1.5 rounded-lg bg-amber-600 text-white text-xs font-semibold hover:bg-amber-700", children: "Confirmar e gravar mesmo assim" }),
        /* @__PURE__ */ jsx("button", { onClick: () => setAvisos([]), className: "px-3 py-1.5 rounded-lg border border-stone-300 text-stone-600 text-xs font-semibold hover:bg-stone-50", children: "Corrigir" })
      ] })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    !readOnly && /* @__PURE__ */ jsx("button", { onClick: submit, disabled: dataFimInvalida, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-50 disabled:cursor-not-allowed", children: "Guardar" })
  ] });
}
const calcularCustoFormula = (formula, center, materiais, equipamentos, maoDeObra, dataRef, consumiveis) => {
  const anoRef = new Date(dataRef).getFullYear();
  const nomeMaterial = (id) => materiais.find((m) => m.id === id)?.designacao || "\u2014";
  const kgNaoAgregadoPorTonelada = TRABALHO_MATERIAL_KEYS.reduce((s, key) => s + (parseFloat(formula.trabalho?.[key]?.design) || 0), 0);
  const kgAgregadoPorTonelada = Math.max(0, 1e3 - kgNaoAgregadoPorTonelada);
  const linhasDiretos = [];
  SILO_COLS.forEach((c) => {
    const silo = formula.silos?.[c.key];
    const pct = parseFloat(silo?.pct) || 0;
    if (!silo?.materialId || pct <= 0) return;
    const material = materiais.find((m) => m.id === silo.materialId);
    const kgPorTonelada = pct / 100 * kgAgregadoPorTonelada;
    const entry = material ? precoVigente(material, dataRef) : null;
    const precoUnitario = entry ? calcularPrecoFinal(entry) : null;
    linhasDiretos.push({
      key: `silo-${c.key}`,
      origem: c.label,
      designacao: nomeMaterial(silo.materialId),
      kgPorTonelada,
      precoUnitario,
      custo: precoUnitario !== null ? kgPorTonelada / 1e3 * precoUnitario : null
    });
  });
  TRABALHO_MATERIAL_KEYS.forEach((key) => {
    const item = formula.trabalho?.[key];
    const kgPorTonelada = parseFloat(item?.design) || 0;
    if (!item?.materialId || kgPorTonelada <= 0) return;
    const material = materiais.find((m) => m.id === item.materialId);
    const entry = material ? precoVigente(material, dataRef) : null;
    const precoUnitario = entry ? calcularPrecoFinal(entry) : null;
    linhasDiretos.push({
      key: `trabalho-${key}`,
      origem: getTrabalhoLabel(center, key),
      designacao: nomeMaterial(item.materialId),
      kgPorTonelada,
      precoUnitario,
      custo: precoUnitario !== null ? kgPorTonelada / 1e3 * precoUnitario : null
    });
  });
  const totalDiretos = linhasDiretos.reduce((s, l) => s + (l.custo || 0), 0);
  const temSemPreco = linhasDiretos.some((l) => l.precoUnitario === null);
  const custoCombustao = (blocoKey, nomeBase) => {
    const bloco = center.parametrizacao?.[blocoKey] || {};
    if (!bloco.combustivelId) return null;
    const combustivelObj = (consumiveis || []).find((c) => c.id === bloco.combustivelId);
    if (!combustivelObj) return null;
    const consumoEntry = taxaVigenteEmData(bloco.historico, dataRef);
    const consumo = consumoEntry ? parseFloat(consumoEntry.valor) : null;
    const precoEntry = precoVigente(combustivelObj, dataRef);
    const precoUnitario = precoEntry ? calcularPrecoFinal(precoEntry) : null;
    const unidadeConsumo = combustivelObj.unidadeCusto || "";
    let custoPorTonelada = null;
    let aviso = null;
    if (consumo === null) {
      aviso = "Sem consumo definido para esta data.";
    } else if (precoUnitario === null) {
      aviso = "Sem pre\xE7o em vigor nesta data para o combust\xEDvel selecionado.";
    } else {
      custoPorTonelada = consumo * precoUnitario;
    }
    return { key: blocoKey, nome: `${nomeBase} (${combustivelObj.designacao})`, consumo, unidadeConsumo, precoUnitario, custoPorTonelada, aviso };
  };
  const linhasCombustao = [
    custoCombustao("blocoTermico", "Bloco T\xE9rmico"),
    custoCombustao("queimadorTamborSecador", "Queimador Tambor Secador")
  ].filter(Boolean);
  const totalCombustao = linhasCombustao.reduce((s, l) => s + (l.custoPorTonelada || 0), 0);
  const producaoAnualHistorico = center.parametrizacao?.producaoAnualHistorico || [];
  const producaoDoAno = (() => {
    const doAno = producaoAnualHistorico.filter((h) => h.ano === anoRef);
    if (doAno.length === 0) return null;
    return [...doAno].sort((a, b) => (b.dataRegisto || "").localeCompare(a.dataRegisto || ""))[0];
  })();
  const converterAnualParaTonelada = (blocoKey, nome) => {
    const bloco = center.parametrizacao?.[blocoKey] || {};
    const taxa = taxaVigenteEmData(bloco.historico, dataRef);
    if (!taxa) return null;
    const valorAnual = parseFloat(taxa.valor) || 0;
    if (valorAnual <= 0) return null;
    const custoPorTonelada = producaoDoAno ? valorAnual / (parseFloat(producaoDoAno.valor) || 1) : null;
    return {
      key: blocoKey,
      nome,
      valor: valorAnual,
      unidade: "\u20AC/ano",
      custoPorTonelada,
      aviso: !producaoDoAno ? `Sem Produ\xE7\xE3o Anual Estimada para ${anoRef} \u2014 indique-a em Parametriza\xE7\xE3o de Produ\xE7\xE3o para converter para \u20AC/tonelada.` : null
    };
  };
  const linhasEquipamento = [];
  const itensEquipamento = center.parametrizacao?.equipamentosItens || [];
  itensEquipamento.forEach((item) => {
    const equipamentoObj = (equipamentos || []).find((e) => e.id === item.equipamentoId);
    if (!equipamentoObj) return;
    const entry = precoVigente(equipamentoObj, dataRef);
    const valor = entry ? calcularPrecoFinal(entry) : null;
    const unidade = equipamentoObj.unidadeCusto || "";
    const horasEntry = taxaVigenteEmData(item.historico, dataRef);
    const horasDiarias = horasEntry ? parseFloat(horasEntry.horasDiarias) || 0 : null;
    let custoPorTonelada = null;
    let aviso = null;
    if (valor === null) {
      aviso = "Sem pre\xE7o em vigor nesta data.";
    } else if (normUnidade(unidade) === normUnidade("\u20AC/tonelada")) {
      custoPorTonelada = valor;
    } else if (normUnidade(unidade) === normUnidade("\u20AC/hora")) {
      if (!horasDiarias) {
        aviso = 'Sem "Horas Di\xE1rias de Trabalho" definidas (ou definidas a 0) para esta data.';
      } else if (!producaoDoAno) {
        aviso = `Sem Produ\xE7\xE3o Anual Estimada para ${anoRef}.`;
      } else {
        const custoAnual = DIAS_UTEIS_ANO_PADRAO * horasDiarias * valor;
        custoPorTonelada = custoAnual / (parseFloat(producaoDoAno.valor) || 1);
      }
    } else if (normUnidade(unidade) === normUnidade("\u20AC/dia")) {
      if (!producaoDoAno) {
        aviso = `Sem Produ\xE7\xE3o Anual Estimada para ${anoRef}.`;
      } else {
        const custoAnual = DIAS_UTEIS_ANO_PADRAO * valor;
        custoPorTonelada = custoAnual / (parseFloat(producaoDoAno.valor) || 1);
      }
    } else {
      aviso = `Definido em ${unidade || "\u2014"} \u2014 mude a unidade do equipamento para \u20AC/hora, \u20AC/dia ou \u20AC/tonelada.`;
    }
    linhasEquipamento.push({ key: item.id, nome: equipamentoObj.designacao, valor, unidade, custoPorTonelada, aviso });
  });
  const totalEquipamento = linhasEquipamento.reduce((s, l) => s + (l.custoPorTonelada || 0), 0);
  const linhasFixos = [
    converterAnualParaTonelada("custoCertificacao", "Certifica\xE7\xE3o"),
    converterAnualParaTonelada("custoEstaleiro", "Aluguer de Estaleiro"),
    converterAnualParaTonelada("custoQAS", "Qualidade/Ambiente/Seguran\xE7a"),
    converterAnualParaTonelada("custoAmortizacao", "Amortiza\xE7\xE3o da Central"),
    converterAnualParaTonelada("custoMudancaCentral", "Mudan\xE7a de Central"),
    converterAnualParaTonelada("custoManutencaoCentral", "Manuten\xE7\xE3o da Central")
  ].filter(Boolean);
  const totalFixos = linhasFixos.reduce((s, l) => s + (l.custoPorTonelada || 0), 0);
  const linhasVariaveis = [];
  const blocoControloLab = center.parametrizacao?.custoControloLab || {};
  const taxaControloLab = taxaVigenteEmData(blocoControloLab.historico, dataRef);
  if (taxaControloLab) {
    const valor = parseFloat(taxaControloLab.valor) || 0;
    linhasVariaveis.push({ key: "custoControloLab", nome: "Controlo Laboratorial", valor, unidade: "\u20AC/t", custoPorTonelada: valor, aviso: null });
  }
  const blocoAluguerCentral = center.parametrizacao?.custoAluguerCentral || {};
  const taxaAluguerCentral = taxaVigenteEmData(blocoAluguerCentral.historico, dataRef);
  if (taxaAluguerCentral) {
    const valor = parseFloat(taxaAluguerCentral.valor) || 0;
    linhasVariaveis.push({ key: "custoAluguerCentral", nome: "Aluguer da Central", valor, unidade: "\u20AC/t", custoPorTonelada: valor, aviso: null });
  }
  const totalVariaveis = linhasVariaveis.reduce((s, l) => s + (l.custoPorTonelada || 0), 0);
  const HORAS_POR_DIA = 10;
  const diasUteis = DIAS_UTEIS_ANO_PADRAO;
  const itensMaoDeObra = center.parametrizacao?.maoDeObraItens || [];
  const linhasMaoDeObra = itensMaoDeObra.map((item) => {
    const artigo = (maoDeObra || []).find((m) => m.id === item.maoDeObraId);
    if (!artigo) return null;
    const entry = precoVigente(artigo, dataRef);
    const valorUnitario = entry ? calcularPrecoFinal(entry) : null;
    const unidade = artigo.unidadeCusto || "";
    const quantidadeEntry = taxaVigenteEmData(item.historico, dataRef);
    const quantidade = quantidadeEntry ? parseFloat(quantidadeEntry.quantidade) || 0 : null;
    let custoAnual = null;
    let aviso = null;
    if (quantidade === null) {
      aviso = "Sem quantidade definida para esta data.";
    } else if (valorUnitario === null) {
      aviso = "Sem pre\xE7o em vigor nesta data.";
    } else if (normUnidade(unidade) === normUnidade("\u20AC/hora")) {
      custoAnual = diasUteis * HORAS_POR_DIA * valorUnitario * quantidade;
    } else if (normUnidade(unidade) === normUnidade("\u20AC/dia")) {
      custoAnual = diasUteis * valorUnitario * quantidade;
    } else if (normUnidade(unidade) === normUnidade("\u20AC/m\xEAs")) {
      custoAnual = 12 * valorUnitario * quantidade;
    } else {
      aviso = `Unidade ${unidade || "\u2014"} n\xE3o suportada neste c\xE1lculo.`;
    }
    const custoPorTonelada = custoAnual !== null && producaoDoAno ? custoAnual / (parseFloat(producaoDoAno.valor) || 1) : null;
    if (custoAnual !== null && !producaoDoAno) aviso = `Sem Produ\xE7\xE3o Anual Estimada para ${anoRef}.`;
    return { key: item.id, nome: artigo.designacao, quantidade, valorUnitario, unidade, custoAnual, custoPorTonelada, aviso };
  }).filter(Boolean);
  const totalMaoDeObra = linhasMaoDeObra.reduce((s, l) => s + (l.custoPorTonelada || 0), 0);
  const totalGeral = totalDiretos + totalCombustao + totalEquipamento + totalMaoDeObra + totalFixos + totalVariaveis;
  const arred2 = (v) => v === null || v === void 0 ? v : Math.round((v + Number.EPSILON) * 100) / 100;
  const arredLinhas = (arr) => arr.map((l) => ({ ...l, custoPorTonelada: arred2(l.custoPorTonelada) }));
  return {
    anoRef,
    producaoDoAno,
    linhasDiretos: linhasDiretos.map((l) => ({ ...l, custo: arred2(l.custo) })),
    totalDiretos: arred2(totalDiretos),
    temSemPreco,
    linhasCombustao: arredLinhas(linhasCombustao),
    totalCombustao: arred2(totalCombustao),
    linhasEquipamento: arredLinhas(linhasEquipamento),
    totalEquipamento: arred2(totalEquipamento),
    linhasFixos: arredLinhas(linhasFixos),
    totalFixos: arred2(totalFixos),
    linhasVariaveis: arredLinhas(linhasVariaveis),
    totalVariaveis: arred2(totalVariaveis),
    linhasMaoDeObra: arredLinhas(linhasMaoDeObra),
    totalMaoDeObra: arred2(totalMaoDeObra),
    totalGeral: arred2(totalGeral)
  };
};
function ListaCustosFormulasModal({ center, formulas, materiais, equipamentos, maoDeObra, consumiveis, logotipo, nomeUtilizadorAtual, onOpenFormula, onClose }) {
  const [dataRef, setDataRef] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [fatorK, setFatorK] = useState("1");
  const [exportarQue, setExportarQue] = useState("tudo");
  const k = parseFloat(fatorK) || 1;
  const linhas = [...formulas].filter((f) => f.incluirEmCustosTodas !== false).sort((a, b) => (a.codigo || "").localeCompare(b.codigo || "", "pt", { numeric: true, sensitivity: "base" })).map((f) => {
    const { totalGeral } = calcularCustoFormula(f, center, materiais, equipamentos, maoDeObra, dataRef, consumiveis);
    return { id: f.id, codigo: f.codigo, designacao: f.designacao, observacoes: f.observacoes || "", custo: totalGeral, venda: Math.round((totalGeral * k + Number.EPSILON) * 100) / 100 };
  });
  const exportar = () => {
    const mostrarCusto = exportarQue !== "so-venda";
    const mostrarVenda = exportarQue !== "so-custo";
    const agora = /* @__PURE__ */ new Date();
    const rodapeTexto = `Exportado em ${formatDateTimePT(agora.toISOString())} por ${nomeUtilizadorAtual || "\u2014"}`;
    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const margemEsq = 15;
    const margemDir = 15;
    const pageWidthTopo = doc.internal.pageSize.getWidth();
    let y = 15;
    const alturaCabecalho = 18;
    if (logotipo) {
      try {
        const formatoMatch = /^data:image\/(\w+);/.exec(logotipo);
        const formato = formatoMatch ? formatoMatch[1].toUpperCase().replace("JPG", "JPEG") : "PNG";
        const propsImg = doc.getImageProperties(logotipo);
        const proporcao = propsImg.width && propsImg.height ? propsImg.width / propsImg.height : 3;
        doc.addImage(logotipo, formato, margemEsq, y, alturaCabecalho * proporcao, alturaCabecalho);
      } catch {
      }
    }
    const nomeCentral = center?.codigo ? `${center.codigo} \u2014 ${center.nome || ""}` : center?.nome || "";
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(28, 25, 23);
    doc.text("Custo Final das Misturas", pageWidthTopo - margemDir, y + 4, { align: "right" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(87, 83, 78);
    doc.text(nomeCentral, pageWidthTopo - margemDir, y + 9.5, { align: "right" });
    doc.setTextColor(120, 113, 108);
    doc.text(`Data: ${formatDatePT(dataRef)}${mostrarVenda && exportarQue !== "so-venda" ? ` \xB7 Fator K: ${k.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : ""}`, pageWidthTopo - margemDir, y + 14.5, { align: "right" });
    y += alturaCabecalho + 3;
    doc.setDrawColor(214, 211, 209);
    doc.setLineWidth(0.2);
    doc.line(margemEsq, y, pageWidthTopo - margemDir, y);
    y += 4;
    const head = [["C\xF3digo", "Designa\xE7\xE3o", ...mostrarCusto ? ["Custo (\u20AC/t)"] : [], ...mostrarVenda ? ["Venda (\u20AC/t)"] : []]];
    const body = linhas.map((l) => [
      l.codigo || "",
      l.designacao + (l.observacoes ? `
Obs: ${l.observacoes}` : ""),
      ...mostrarCusto ? [l.custo.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " \u20AC"] : [],
      ...mostrarVenda ? [l.venda.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " \u20AC"] : []
    ]);
    autoTable(doc, {
      head,
      body,
      startY: y,
      margin: { left: margemEsq, right: margemEsq, bottom: 16 },
      styles: { font: "helvetica", fontSize: 9, cellPadding: 2.2, lineColor: [214, 211, 209], lineWidth: 0.1 },
      headStyles: { fillColor: [245, 245, 244], textColor: [87, 83, 78], fontStyle: "bold", fontSize: 8 },
      alternateRowStyles: { fillColor: [250, 250, 249] },
      columnStyles: Object.fromEntries(
        [mostrarCusto, mostrarVenda].map((show, i) => show ? [2 + i, { halign: "right" }] : null).filter(Boolean)
      ),
      didParseCell: (data) => {
        if (data.column.index === 1 && data.cell.raw && data.cell.raw.includes("\nObs:")) {
          data.cell.styles.fontSize = 8;
        }
      },
      didDrawPage: () => {
        const pageHeight2 = doc.internal.pageSize.getHeight();
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);
        doc.setTextColor(168, 162, 158);
        doc.text(rodapeTexto, margemEsq, pageHeight2 - 8);
      }
    });
    const totalPaginas = doc.internal.getNumberOfPages();
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    for (let i = 1; i <= totalPaginas; i++) {
      doc.setPage(i);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7);
      doc.setTextColor(168, 162, 158);
      doc.text(`P\xE1gina ${i} de ${totalPaginas}`, pageWidth - margemEsq, pageHeight - 8, { align: "right" });
    }
    doc.save(`Custos_Formulas_${center?.nome || "centro"}.pdf`);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Custos de Todas as F\xF3rmulas", subtitle: center?.nome || "", onClose, wide: true, children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-3 mb-3", children: [
      /* @__PURE__ */ jsx("div", { className: "w-52", children: /* @__PURE__ */ jsx(Field, { label: "Custo \xE0 data", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataRef, onChange: (e) => setDataRef(e.target.value), className: inputCls }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-32", children: /* @__PURE__ */ jsx(Field, { label: "Fator K", children: /* @__PURE__ */ jsx("input", { type: "number", step: "0.01", min: "0", value: fatorK, onChange: (e) => setFatorK(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "1,20" }) }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-3 mb-5", children: [
      /* @__PURE__ */ jsx("div", { className: "w-56", children: /* @__PURE__ */ jsx(Field, { label: "Exportar PDF com", children: /* @__PURE__ */ jsxs("select", { value: exportarQue, onChange: (e) => setExportarQue(e.target.value), className: inputCls, children: [
        /* @__PURE__ */ jsx("option", { value: "tudo", children: "Custo e Venda" }),
        /* @__PURE__ */ jsx("option", { value: "so-custo", children: "S\xF3 Custo" }),
        /* @__PURE__ */ jsx("option", { value: "so-venda", children: "S\xF3 Venda" })
      ] }) }) }),
      /* @__PURE__ */ jsxs("button", { onClick: exportar, className: "mb-4 flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileText, { size: 15 }),
        " Exportar PDF"
      ] })
    ] }),
    linhas.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-6", children: "Ainda n\xE3o h\xE1 f\xF3rmulas neste centro." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "C\xF3digo" }),
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Designa\xE7\xE3o" }),
        /* @__PURE__ */ jsxs("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: [
          "Custo",
          /* @__PURE__ */ jsx("span", { className: "block font-normal normal-case text-[10px] text-stone-400", children: "\u20AC/t" })
        ] }),
        /* @__PURE__ */ jsxs("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: [
          "Venda",
          /* @__PURE__ */ jsx("span", { className: "block font-normal normal-case text-[10px] text-stone-400", children: "\u20AC/t" })
        ] })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhas.map((l, i) => /* @__PURE__ */ jsxs(
        "tr",
        {
          onClick: () => onOpenFormula(formulas.find((f) => f.id === l.id)),
          className: `cursor-pointer hover:bg-amber-50 ${i !== linhas.length - 1 ? "border-b border-stone-100" : ""}`,
          children: [
            /* @__PURE__ */ jsx("td", { className: "px-3 py-2 font-mono-data text-amber-700 font-semibold", children: l.codigo }),
            /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
              l.designacao,
              l.observacoes && /* @__PURE__ */ jsxs("span", { className: "block text-xs text-amber-700 font-normal mt-0.5", children: [
                "Obs: ",
                l.observacoes
              ] })
            ] }),
            /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-right font-mono-data text-stone-600", children: [
              l.custo.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
              " \u20AC"
            ] }),
            /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-right font-mono-data font-semibold text-stone-800", children: [
              l.venda.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
              " \u20AC"
            ] })
          ]
        },
        l.id
      )) })
    ] }) })
  ] });
}
function FichaCustoModal({ formula, center, materiais, equipamentos, maoDeObra, consumiveis, logotipo, onClose }) {
  const [dataRef, setDataRef] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [fatorK, setFatorK] = useState("1");
  const {
    anoRef,
    linhasDiretos,
    totalDiretos,
    temSemPreco,
    linhasCombustao,
    totalCombustao,
    linhasEquipamento,
    totalEquipamento,
    linhasFixos,
    totalFixos,
    linhasVariaveis,
    totalVariaveis,
    linhasMaoDeObra,
    totalMaoDeObra,
    totalGeral
  } = calcularCustoFormula(formula, center, materiais, equipamentos, maoDeObra, dataRef, consumiveis);
  const k = parseFloat(fatorK) || 1;
  const totalComFatorK = Math.round((totalGeral * k + Number.EPSILON) * 100) / 100;
  const exportarFichaCusto = () => {
    const linhaHtml = (nome, valorTxt, custoTxt, aviso) => `
      <tr><td>${nome}${aviso ? `<br><span style="font-size:10px;color:#b45309">${aviso}</span>` : ""}</td><td style="text-align:right">${valorTxt}</td><td style="text-align:right; font-weight:bold">${custoTxt}</td></tr>`;
    const tabelaDiretos = linhasDiretos.map((l) => `
      <tr>
        <td>${l.origem}</td><td>${l.designacao}</td>
        <td style="text-align:right">${l.kgPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td style="text-align:right">${l.precoUnitario !== null ? l.precoUnitario.toLocaleString("pt-PT", { maximumFractionDigits: 3 }) + " \u20AC" : "sem pre\xE7o"}</td>
        <td style="text-align:right; font-weight:bold">${l.custo !== null ? l.custo.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " \u20AC" : "\u2014"}</td>
      </tr>`).join("");
    const tabelaCombustao = linhasCombustao.map((l) => linhaHtml(l.nome, l.consumo !== null ? `${l.consumo.toLocaleString("pt-PT", { maximumFractionDigits: 3 })} ${l.unidadeConsumo}` : "\u2014", l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014", l.aviso)).join("");
    const tabelaEquip = linhasEquipamento.map((l) => linhaHtml(l.nome, l.valor !== null ? `${l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${l.unidade}` : "\u2014", l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014", l.aviso)).join("");
    const tabelaMaoObra = linhasMaoDeObra.map((l) => linhaHtml(`${l.nome} (qtd. ${l.quantidade ?? "\u2014"})`, l.valorUnitario !== null ? `${l.valorUnitario.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${l.unidade}` : "\u2014", l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014", l.aviso)).join("");
    const tabelaFixos = linhasFixos.map((l) => linhaHtml(l.nome, `${l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/ano`, l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014", l.aviso)).join("");
    const tabelaVariaveis = linhasVariaveis.map((l) => linhaHtml(l.nome, `${l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${l.unidade}`, l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014", l.aviso)).join("");
    const html = `<!DOCTYPE html>
      <html><head><meta charset="utf-8"><title>Ficha de Custo \u2014 ${formula.codigo}</title>
      <style>
        body { font-family: 'Arial Narrow', Arial, sans-serif; color: #1c1917; padding: 32px; max-width: 900px; margin: 0 auto; font-size: 13px; }
        .logo { max-height: 55px; margin-bottom: 16px; }
        h1 { font-size: 18px; margin-bottom: 2px; }
        .sub { color: #78716c; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 20px; }
        h2 { font-size: 13px; text-transform: uppercase; letter-spacing: 0.05em; color: #57534e; border-bottom: 1px solid #e7e5e4; padding-bottom: 4px; margin-top: 22px; }
        table { width: 100%; border-collapse: collapse; margin-top: 6px; font-size: 12px; }
        th, td { border: 1px solid #d6d3d1; padding: 5px 7px; }
        th { background: #f5f5f4; text-align: left; font-size: 10px; text-transform: uppercase; }
        .subtotal { text-align: right; font-weight: bold; margin-top: 4px; font-size: 13px; }
        .totalfinal { margin-top: 24px; padding: 14px 18px; background: #fef3c7; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }
        .totalfinal .label { font-weight: bold; text-transform: uppercase; font-size: 12px; }
        .totalfinal .valor { font-weight: bold; font-size: 20px; }
      </style></head>
      <body>
        ${logotipo ? `<img class="logo" src="${logotipo}" alt="Log\xF3tipo" />` : ""}
        <h1>Ficha de Custo</h1>
        <p class="sub">${formula.codigo} \u2014 ${formula.designacao} \xB7 ${center?.nome || ""}</p>
        <p style="font-size:12px; color:#44403c;">Custo \xE0 data: <strong>${formatDatePT(dataRef)}</strong> \xB7 Fator K: <strong>${k.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></p>

        <h2>1 \u2014 Custos Diretos (materiais constituintes)</h2>
        <table>
          <tr><th>Origem</th><th>Material</th><th style="text-align:right">kg/t</th><th style="text-align:right">Pre\xE7o</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaDiretos || '<tr><td colspan="5" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal Diretos: ${totalDiretos.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <h2>2 \u2014 Bloco T\xE9rmico e Queimador Tambor Secador</h2>
        <table>
          <tr><th>Origem</th><th style="text-align:right">Consumo</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaCombustao || '<tr><td colspan="3" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal Bloco T\xE9rmico/Queimador: ${totalCombustao.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <h2>3 \u2014 Custo Equipamento</h2>
        <table>
          <tr><th>Equipamento</th><th style="text-align:right">Valor</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaEquip || '<tr><td colspan="3" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal Equipamento: ${totalEquipamento.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <h2>4 \u2014 M\xE3o de Obra</h2>
        <table>
          <tr><th>Categoria</th><th style="text-align:right">Valor</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaMaoObra || '<tr><td colspan="3" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal M\xE3o de Obra: ${totalMaoDeObra.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <h2>5 \u2014 Custos Fixos</h2>
        <table>
          <tr><th>Custo</th><th style="text-align:right">Valor Anual</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaFixos || '<tr><td colspan="3" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal Custos Fixos: ${totalFixos.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <h2>6 \u2014 Custos Vari\xE1veis</h2>
        <table>
          <tr><th>Custo</th><th style="text-align:right">Valor</th><th style="text-align:right">Custo (\u20AC/t)</th></tr>
          ${tabelaVariaveis || '<tr><td colspan="3" style="text-align:center">\u2014</td></tr>'}
        </table>
        <p class="subtotal">Subtotal Custos Vari\xE1veis: ${totalVariaveis.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t</p>

        <div class="totalfinal">
          <span class="label">Total Geral (sem Fator K): ${totalGeral.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC/t<br>Total Final (com Fator K \xD7 ${k.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })})</span>
          <span class="valor">${totalComFatorK.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC / t</span>
        </div>
      </body></html>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Ficha_Custo_${formula.codigo}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Ficha de Custo", subtitle: `${formula.codigo} \u2014 ${formula.designacao}`, onClose, wide: true, children: [
    /* @__PURE__ */ jsx("div", { className: "flex justify-end mb-3", children: /* @__PURE__ */ jsxs("button", { onClick: exportarFichaCusto, className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
      /* @__PURE__ */ jsx(FileText, { size: 15 }),
      " Exportar PDF"
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-end gap-3 mb-5", children: [
      /* @__PURE__ */ jsx("div", { className: "w-52", children: /* @__PURE__ */ jsx(Field, { label: "Custo \xE0 data", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataRef, onChange: (e) => setDataRef(e.target.value), className: inputCls }) }) }),
      /* @__PURE__ */ jsx("div", { className: "w-32", children: /* @__PURE__ */ jsx(Field, { label: "Fator K", children: /* @__PURE__ */ jsx("input", { type: "number", step: "0.01", min: "0", value: fatorK, onChange: (e) => setFatorK(e.target.value), className: `${inputCls} font-mono-data`, placeholder: "1,20" }) }) }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-4", children: "Usa os pre\xE7os e taxas em vigor nesta data, com a composi\xE7\xE3o atual da f\xF3rmula. O Fator K multiplica o total final (ex: 1,20 para acrescentar 20%)." })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "1 \u2014 Custos Diretos (materiais constituintes)" }),
    linhasDiretos.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Esta f\xF3rmula ainda n\xE3o tem componentes com material associado." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Origem" }),
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Material" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "kg / tonelada" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Pre\xE7o (\u20AC/un.)" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasDiretos.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasDiretos.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-500", children: l.origem }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-stone-800", children: l.designacao }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.kgPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.precoUnitario !== null ? `${l.precoUnitario.toLocaleString("pt-PT", { maximumFractionDigits: 3 })} \u20AC` : /* @__PURE__ */ jsx("span", { className: "text-red-500", children: "sem pre\xE7o" }) }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custo !== null ? `${l.custo.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    temSemPreco && /* @__PURE__ */ jsx("p", { className: "text-xs text-red-600 mb-3", children: "H\xE1 materiais sem pre\xE7o em vigor nesta data \u2014 o total fica incompleto at\xE9 indicar um pre\xE7o para eles." }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal Diretos" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalDiretos.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "2 \u2014 Bloco T\xE9rmico e Queimador Tambor Secador" }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-2", children: "Consumo de combust\xEDvel em vigor nesta data (definido em Parametriza\xE7\xE3o de Produ\xE7\xE3o) \xD7 pre\xE7o em vigor do combust\xEDvel selecionado." }),
    linhasCombustao.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Ainda n\xE3o h\xE1 combust\xEDvel selecionado para o Bloco T\xE9rmico ou o Queimador Tambor Secador (configure em Parametriza\xE7\xE3o de Produ\xE7\xE3o)." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Origem" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Consumo" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasCombustao.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasCombustao.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
          l.nome,
          l.aviso && /* @__PURE__ */ jsx("span", { className: "block text-[11px] text-amber-600 mt-0.5", children: l.aviso })
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.consumo !== null ? `${l.consumo.toLocaleString("pt-PT", { maximumFractionDigits: 3 })} ${l.unidadeConsumo}` : "\u2014" }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal Bloco T\xE9rmico/Queimador" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalCombustao.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "3 \u2014 Custo Equipamento" }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-2", children: "P\xE1 carregadora e outro equipamento vari\xE1vel \u2014 repartidos pela produ\xE7\xE3o anual estimada para chegar ao \u20AC/tonelada." }),
    linhasEquipamento.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Ainda n\xE3o h\xE1 custos de equipamento configurados em Parametriza\xE7\xE3o de Produ\xE7\xE3o." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Item" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Valor" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasEquipamento.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasEquipamento.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
          l.nome,
          l.aviso && /* @__PURE__ */ jsx("span", { className: "block text-[11px] text-amber-600 mt-0.5", children: l.aviso })
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.valor !== null ? `${l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${l.unidade}` : "\u2014" }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal Equipamento" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalEquipamento.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "4 \u2014 M\xE3o de Obra" }),
    /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 mb-2", children: [
      DIAS_UTEIS_ANO_PADRAO,
      " dias \xFAteis (m\xE9dia fixa) \xD7 10 h/dia (categorias em \u20AC/hora) ou s\xF3 os dias (categorias em \u20AC/dia) \xD7 valor \xD7 quantidade em vigor nesta data, repartido pela produ\xE7\xE3o anual estimada."
    ] }),
    linhasMaoDeObra.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Ainda n\xE3o h\xE1 categorias de m\xE3o de obra atribu\xEDdas a este centro (configure em Parametriza\xE7\xE3o de Produ\xE7\xE3o)." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Categoria" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Quantidade" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Valor" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasMaoDeObra.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasMaoDeObra.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
          l.nome,
          l.aviso && /* @__PURE__ */ jsx("span", { className: "block text-[11px] text-amber-600 mt-0.5", children: l.aviso })
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.quantidade }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data", children: l.valorUnitario !== null ? `${l.valorUnitario.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${l.unidade}` : "\u2014" }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal M\xE3o de Obra" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalMaoDeObra.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "5 \u2014 Custos Fixos" }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-2", children: "Valores globais anuais, repartidos pela Produ\xE7\xE3o Anual Estimada." }),
    linhasFixos.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Ainda n\xE3o h\xE1 custos fixos com valor registado para esta data (configure em Parametriza\xE7\xE3o de Produ\xE7\xE3o)." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Valor Anual" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasFixos.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasFixos.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
          l.nome,
          l.aviso && /* @__PURE__ */ jsx("span", { className: "block text-[11px] text-amber-600 mt-0.5", children: l.aviso })
        ] }),
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-right font-mono-data", children: [
          l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
          " \u20AC"
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal Custos Fixos" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalFixos.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "6 \u2014 Custos Vari\xE1veis" }),
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-2", children: "Controlo Laboratorial e Aluguer da Central j\xE1 em \u20AC/tonelada \u2014 aplicados diretamente, sem passar pela Produ\xE7\xE3o Anual Estimada." }),
    linhasVariaveis.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 py-4", children: "Ainda n\xE3o h\xE1 custos vari\xE1veis com valor registado para esta data (configure em Parametriza\xE7\xE3o de Produ\xE7\xE3o)." }) : /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-2", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "text-left px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Valor" }),
        /* @__PURE__ */ jsx("th", { className: "text-right px-3 py-2 text-xs font-semibold uppercase tracking-wide text-stone-500", children: "Custo (\u20AC/t)" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: linhasVariaveis.map((l, i) => /* @__PURE__ */ jsxs("tr", { className: i !== linhasVariaveis.length - 1 ? "border-b border-stone-100" : "", children: [
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-stone-800", children: [
          l.nome,
          l.aviso && /* @__PURE__ */ jsx("span", { className: "block text-[11px] text-amber-600 mt-0.5", children: l.aviso })
        ] }),
        /* @__PURE__ */ jsxs("td", { className: "px-3 py-2 text-right font-mono-data", children: [
          l.valor.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
          " ",
          l.unidade
        ] }),
        /* @__PURE__ */ jsx("td", { className: "px-3 py-2 text-right font-mono-data font-semibold", children: l.custoPorTonelada !== null ? `${l.custoPorTonelada.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} \u20AC` : "\u2014" })
      ] }, l.key)) })
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-5", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Subtotal Custos Vari\xE1veis" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalVariaveis.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-stone-100 rounded-lg px-4 py-2.5 mb-2", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-stone-700", children: "Total Geral (1 a 5), sem Fator K" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        totalGeral.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-3 mb-4", children: [
      /* @__PURE__ */ jsxs("span", { className: "text-sm font-semibold text-amber-800", children: [
        "Total Final (com Fator K \xD7 ",
        k.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        ")"
      ] }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-lg text-amber-800", children: [
        totalComFatorK.toLocaleString("pt-PT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        " \u20AC / t"
      ] })
    ] })
  ] });
}
function DopFormulaModal({ formula, center, dopConfig, logotipo, isAdmin, nomeUtilizadorAtual, onSave, onDeleteHistorico, onClose }) {
  const cfg = dopConfig || {};
  const [historicoLocal, setHistoricoLocal] = useState(formula.dopHistorico || []);
  const historicoOrdenado2 = [...historicoLocal].sort((a, b) => (b.dataRegisto || "").localeCompare(a.dataRegisto || ""));
  const dop = historicoOrdenado2[0] || {};
  const [codigoDop, setCodigoDop] = useState(dop.codigoDop || "");
  const [norma, setNorma] = useState(dop.norma || "");
  const [utilizacaoPrevista, setUtilizacaoPrevista] = useState(dop.utilizacaoPrevista || "Constru\xE7\xE3o e manuten\xE7\xE3o de estradas, pistas de aeroporto e outras \xE1reas de circula\xE7\xE3o de transporte.");
  const [porosidadeVmax, setPorosidadeVmax] = useState(dop.porosidadeVmax || "");
  const [porosidadeVmin, setPorosidadeVmin] = useState(dop.porosidadeVmin || "");
  const [itsr, setItsr] = useState(dop.itsr || "");
  const [temperatura, setTemperatura] = useState(dop.temperatura || "");
  const [granulometria, setGranulometria] = useState(() => Object.fromEntries(PENEIROS_DOP.map((p) => [p, dop.granulometria?.[p] || { tolInf: "", tolSup: "", li: "", valor: "", ls: "" }])));
  const [ligante, setLigante] = useState(dop.ligante || { tolInf: "", tolSup: "", bmin: "", li: "", valor: "", ls: "" });
  const [defTaxaDnd, setDefTaxaDnd] = useState(dop.defTaxaDnd !== false);
  const [defTaxa, setDefTaxa] = useState(dop.defTaxa || "");
  const [defProfDnd, setDefProfDnd] = useState(dop.defProfDnd !== false);
  const [defProf, setDefProf] = useState(dop.defProf || "");
  const [error, setError] = useState("");
  const [guardado, setGuardado] = useState(false);
  const [semAlteracao, setSemAlteracao] = useState(false);
  const updateGranulometria = (peneiro, campo, valor) => setGranulometria((prev) => ({ ...prev, [peneiro]: { ...prev[peneiro], [campo]: valor } }));
  const updateLigante = (campo, valor) => setLigante((prev) => ({ ...prev, [campo]: valor }));
  const dadosDop = {
    codigoDop,
    norma,
    utilizacaoPrevista,
    porosidadeVmax,
    porosidadeVmin,
    itsr,
    temperatura,
    granulometria,
    ligante,
    defTaxaDnd,
    defTaxa,
    defProfDnd,
    defProf
  };
  const dopFieldsChanged = (a, b) => {
    const camposSimples = ["codigoDop", "norma", "utilizacaoPrevista", "porosidadeVmax", "porosidadeVmin", "itsr", "temperatura", "defTaxa", "defProf"];
    if (camposSimples.some((k) => String(a?.[k] || "") !== String(b?.[k] || ""))) return true;
    if (!!a?.defTaxaDnd !== !!b?.defTaxaDnd) return true;
    if (!!a?.defProfDnd !== !!b?.defProfDnd) return true;
    if (JSON.stringify(a?.granulometria || {}) !== JSON.stringify(b?.granulometria || {})) return true;
    if (JSON.stringify(a?.ligante || {}) !== JSON.stringify(b?.ligante || {})) return true;
    return false;
  };
  const submit = () => {
    if (!norma.trim()) return setError("Indique a Norma Harmonizada");
    if (historicoOrdenado2.length > 0 && !dopFieldsChanged(dop, dadosDop)) {
      setError("");
      setSemAlteracao(true);
      setTimeout(() => setSemAlteracao(false), 2e3);
      return;
    }
    onSave(dadosDop);
    setHistoricoLocal((prev) => [...prev, { ...dadosDop, id: genId(), dataRegisto: (/* @__PURE__ */ new Date()).toISOString(), utilizador: nomeUtilizadorAtual || "" }]);
    setError("");
    setGuardado(true);
    setTimeout(() => setGuardado(false), 2e3);
  };
  const removerHistorico = (entryId) => {
    onDeleteHistorico(entryId);
    setHistoricoLocal((prev) => prev.filter((h) => h.id !== entryId));
  };
  const exportarDop = (camposOverride) => {
    const campos = camposOverride || dadosDop;
    const { codigoDop: codigoDop2, norma: norma2, utilizacaoPrevista: utilizacaoPrevista2, porosidadeVmax: porosidadeVmax2, porosidadeVmin: porosidadeVmin2, itsr: itsr2, temperatura: temperatura2, granulometria: granulometria2, ligante: ligante2, defTaxaDnd: defTaxaDnd2, defTaxa: defTaxa2, defProfDnd: defProfDnd2, defProf: defProf2 } = campos;
    const hoje = /* @__PURE__ */ new Date();
    const dataEmissao = hoje.toLocaleDateString("pt-PT", { day: "numeric", month: "long", year: "numeric" });
    const linhasGranulometria = PENEIROS_DOP.map((p) => {
      const g = granulometria2[p] || {};
      if (!g.li && !g.valor && !g.ls && !g.tolInf && !g.tolSup) return "";
      return `
        <tr>
          <td>${g.tolInf || ""}</td><td>${g.tolSup || ""}</td>
          <td style="text-align:center">${p} mm</td>
          <td style="text-align:center">${g.li || ""}</td>
          <td style="text-align:center; font-weight:bold">${g.valor || ""}</td>
          <td style="text-align:center">${g.ls || ""}</td>
        </tr>`;
    }).join("");
    const html = `<!DOCTYPE html>
      <html><head><meta charset="utf-8"><title>DoP \u2014 ${formula.codigo}</title>
      <style>
        body { font-family: 'Arial Narrow', Arial, sans-serif; color: #1c1917; padding: 32px; max-width: 900px; margin: 0 auto; font-size: 13px; }
        .logo { max-height: 55px; margin-bottom: 16px; }
        h1 { font-size: 17px; text-align: center; margin-bottom: 24px; text-transform: uppercase; }
        .num { font-weight: bold; }
        .bloco { margin-bottom: 16px; }
        .titulo { font-weight: bold; text-transform: uppercase; margin-bottom: 4px; }
        table { width: 100%; border-collapse: collapse; margin-top: 6px; margin-bottom: 10px; }
        th, td { border: 1px solid #78716c; padding: 4px 6px; font-size: 12px; }
        th { background: #f5f5f4; }
        .semlinhas td, .semlinhas th { border: none; padding: 2px 4px; }
        .assinatura { margin-top: 40px; }
      </style></head>
      <body>
        ${logotipo ? `<img class="logo" src="${logotipo}" alt="Log\xF3tipo" />` : ""}
        <h1>Declara\xE7\xE3o de Desempenho</h1>

        <div class="bloco">
          <p class="titulo"><span class="num">1-</span> C\xF3digo de identifica\xE7\xE3o \xFAnico do produto-tipo:</p>
          <p>${formula.codigo} - ${formula.designacao}</p>
          <p>${codigoDop2}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">2-</span> Utiliza\xE7\xE3o(\xF5es) prevista(as):</p>
          <p>${utilizacaoPrevista2}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">3-</span> Fabricante:</p>
          <p>${cfg.fabricanteNome || ""}<br>${(cfg.fabricanteMorada || "").replace(/\n/g, "<br>")}<br>Centro de Produ\xE7\xE3o: ${center?.codigo || ""} - ${center?.nome || ""}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">4-</span> Mandat\xE1rio:</p>
          <p>${cfg.mandatario || "N\xE3o aplic\xE1vel."}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">5-</span> Sistema de Avalia\xE7\xE3o e Verifica\xE7\xE3o da Regularidade do Desempenho (AVCP):</p>
          <p>${cfg.avcp || ""}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">6A-</span> Norma Harmonizada:</p>
          <p>${norma2}</p>
          <p class="titulo" style="margin-top:8px">Organismo(s) Notificado(s):</p>
          <p>${(cfg.organismoNotificado || "").replace(/\n/g, "<br>")}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">6B-</span> Documento de Avalia\xE7\xE3o Europeu:</p>
          <p>${cfg.documentoAvaliacaoEuropeu || "N\xE3o aplic\xE1vel."}</p>
          <p class="titulo" style="margin-top:8px">Avalia\xE7\xE3o T\xE9cnica Europeia:</p>
          <p>${cfg.avaliacaoTecnicaEuropeia || "N\xE3o aplic\xE1vel."}</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">7-</span> Desempenho Declarado:</p>

          <p style="margin-top:10px"><strong>Porosidade</strong> (compacta\xE7\xE3o dos provetes de ensaio conforme refer\xEAncia C.1.3 da EN 13108-20):</p>
          <table class="semlinhas">
            <tr><td>M\xE1xima</td><td style="text-align:right">Vm\xE1x ${porosidadeVmax2} (${porosidadeVmax2}%)</td></tr>
            <tr><td>M\xEDnima</td><td style="text-align:right">Vmin ${porosidadeVmin2} (${porosidadeVmin2}%)</td></tr>
          </table>

          <p><strong>Sensibilidade \xE0 \xE1gua</strong> (conforme EN 12697-12): <strong>ITSR ${itsr2} %</strong></p>
          <p><strong>Temperatura da mistura:</strong> ${temperatura2}</p>

          <p style="margin-top:10px"><strong>Granulometria</strong> (% de material passado) \u2014 Peneiros (s\xE9rie base + s\xE9rie 2):</p>
          <table>
            <tr><th colspan="2">Toler\xE2ncias (%)</th><th>Peneiro</th><th>LI</th><th>%</th><th>LS</th></tr>
            ${linhasGranulometria || '<tr><td colspan="6" style="text-align:center">\u2014</td></tr>'}
          </table>

          <table class="semlinhas">
            <tr><td><strong>Percentagem de ligante</strong></td>
              <td>Tol. ${ligante2.tolInf || ""}/${ligante2.tolSup || ""}</td>
              <td>Bmin ${ligante2.bmin || ""}</td>
              <td style="text-align:right">LI ${ligante2.li || ""} \u2014 <strong>${ligante2.valor || ""}</strong> \u2014 LS ${ligante2.ls || ""}</td>
            </tr>
          </table>

          <p style="margin-top:10px"><strong>Resist\xEAncia \xE0 deforma\xE7\xE3o permanente</strong> (refer\xEAncia D.1.6 da EN 13108-20):</p>
          <table class="semlinhas">
            <tr><td>Taxa de deforma\xE7\xE3o (WTSAIR)</td><td style="text-align:right">${defTaxaDnd2 ? "DND" : defTaxa2}</td></tr>
            <tr><td>Profundidade de rodeira m\xE1xima (PRDAIR)</td><td style="text-align:right">${defProfDnd2 ? "DND" : defProf2}</td></tr>
          </table>
          <p style="font-size:11px; color:#78716c">DND - Desempenho N\xE3o Declarado</p>
        </div>

        <div class="bloco">
          <p class="titulo"><span class="num">8-</span> Documenta\xE7\xE3o T\xE9cnica Adequada e/ou Documenta\xE7\xE3o T\xE9cnica Espec\xEDfica:</p>
          <p>O desempenho do produto identificado acima est\xE1 em conformidade com o conjunto de desempenhos declarados.<br>
          A presente declara\xE7\xE3o de desempenho \xE9 emitida em conformidade com as disposi\xE7\xF5es europeias, sob a exclusiva responsabilidade do fabricante identificado acima.</p>
        </div>

        <div class="assinatura">
          <p>${cfg.localEmissao || ""}, ${dataEmissao}</p>
          <table class="semlinhas" style="margin-top:30px"><tr>
            <td style="width:60%"></td>
            <td style="text-align:center">
              <p>${cfg.signatarioNome || ""}</p>
              <p style="font-size:11px">${cfg.signatarioCargo || ""}</p>
            </td>
          </tr></table>
        </div>
      </body></html>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `DoP_${formula.codigo}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const inputSm = "w-full px-2 py-1 rounded border border-stone-300 text-xs focus:outline-none focus:ring-1 focus:ring-amber-500";
  return /* @__PURE__ */ jsxs(Modal, { title: "Declara\xE7\xE3o de Desempenho (DoP)", subtitle: `${formula.codigo} \u2014 ${formula.designacao}`, onClose, fullscreen: true, children: [
    /* @__PURE__ */ jsx("div", { className: "flex justify-end mb-3", children: /* @__PURE__ */ jsxs("button", { onClick: () => exportarDop(), className: "flex items-center gap-1.5 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
      /* @__PURE__ */ jsx(FileText, { size: 15 }),
      " Exportar DoP"
    ] }) }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4", children: [
      /* @__PURE__ */ jsx(Field, { label: "C\xF3digo espec\xEDfico da DoP", children: /* @__PURE__ */ jsx("input", { value: codigoDop, onChange: (e) => setCodigoDop(e.target.value), className: inputCls, placeholder: "Ex: BB0119-72025" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Norma Harmonizada", children: /* @__PURE__ */ jsx("input", { value: norma, onChange: (e) => setNorma(e.target.value), className: inputCls, placeholder: "Ex: EN 13108-1 - Misturas betuminosas..." }) })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Utiliza\xE7\xE3o Prevista", children: /* @__PURE__ */ jsx("textarea", { value: utilizacaoPrevista, onChange: (e) => setUtilizacaoPrevista(e.target.value), className: inputCls, rows: 2, spellCheck: "true", lang: "pt-PT" }) }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5 mt-4", children: "Desempenho Declarado" }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 sm:grid-cols-4 gap-3 mb-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Porosidade Vm\xE1x (%)", children: /* @__PURE__ */ jsx("input", { value: porosidadeVmax, onChange: (e) => setPorosidadeVmax(e.target.value), className: `${inputCls} font-mono-data` }) }),
      /* @__PURE__ */ jsx(Field, { label: "Porosidade Vmin (%)", children: /* @__PURE__ */ jsx("input", { value: porosidadeVmin, onChange: (e) => setPorosidadeVmin(e.target.value), className: `${inputCls} font-mono-data` }) }),
      /* @__PURE__ */ jsx(Field, { label: "ITSR (%)", children: /* @__PURE__ */ jsx("input", { value: itsr, onChange: (e) => setItsr(e.target.value), className: `${inputCls} font-mono-data` }) }),
      /* @__PURE__ */ jsx(Field, { label: "Temperatura da Mistura", children: /* @__PURE__ */ jsx("input", { value: temperatura, onChange: (e) => setTemperatura(e.target.value), className: inputCls, placeholder: "150\xB0C a 190\xB0C" }) })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Granulometria (% de material passado)" }),
    /* @__PURE__ */ jsx("div", { className: "overflow-x-auto mb-4 border border-stone-200 rounded-lg", children: /* @__PURE__ */ jsxs("table", { className: "w-full text-xs", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-stone-100", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "p-1.5 text-left", children: "Peneiro" }),
        /* @__PURE__ */ jsx("th", { className: "p-1.5", children: "Tol. inf." }),
        /* @__PURE__ */ jsx("th", { className: "p-1.5", children: "Tol. sup." }),
        /* @__PURE__ */ jsx("th", { className: "p-1.5", children: "LI" }),
        /* @__PURE__ */ jsx("th", { className: "p-1.5", children: "% (valor)" }),
        /* @__PURE__ */ jsx("th", { className: "p-1.5", children: "LS" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: PENEIROS_DOP.map((p) => /* @__PURE__ */ jsxs("tr", { className: "border-t border-stone-100", children: [
        /* @__PURE__ */ jsxs("td", { className: "p-1.5 font-mono-data font-semibold text-stone-700", children: [
          p,
          " mm"
        ] }),
        /* @__PURE__ */ jsx("td", { className: "p-1", children: /* @__PURE__ */ jsx("input", { value: granulometria[p]?.tolInf || "", onChange: (e) => updateGranulometria(p, "tolInf", e.target.value), className: inputSm }) }),
        /* @__PURE__ */ jsx("td", { className: "p-1", children: /* @__PURE__ */ jsx("input", { value: granulometria[p]?.tolSup || "", onChange: (e) => updateGranulometria(p, "tolSup", e.target.value), className: inputSm }) }),
        /* @__PURE__ */ jsx("td", { className: "p-1", children: /* @__PURE__ */ jsx("input", { value: granulometria[p]?.li || "", onChange: (e) => updateGranulometria(p, "li", e.target.value), className: inputSm }) }),
        /* @__PURE__ */ jsx("td", { className: "p-1", children: /* @__PURE__ */ jsx("input", { value: granulometria[p]?.valor || "", onChange: (e) => updateGranulometria(p, "valor", e.target.value), className: inputSm }) }),
        /* @__PURE__ */ jsx("td", { className: "p-1", children: /* @__PURE__ */ jsx("input", { value: granulometria[p]?.ls || "", onChange: (e) => updateGranulometria(p, "ls", e.target.value), className: inputSm }) })
      ] }, p)) })
    ] }) }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Percentagem de Ligante" }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-3 sm:grid-cols-6 gap-2 mb-4", children: [
      /* @__PURE__ */ jsx(Field, { label: "Tol. inf.", children: /* @__PURE__ */ jsx("input", { value: ligante.tolInf, onChange: (e) => updateLigante("tolInf", e.target.value), className: inputSm }) }),
      /* @__PURE__ */ jsx(Field, { label: "Tol. sup.", children: /* @__PURE__ */ jsx("input", { value: ligante.tolSup, onChange: (e) => updateLigante("tolSup", e.target.value), className: inputSm }) }),
      /* @__PURE__ */ jsx(Field, { label: "Bmin", children: /* @__PURE__ */ jsx("input", { value: ligante.bmin, onChange: (e) => updateLigante("bmin", e.target.value), className: inputSm }) }),
      /* @__PURE__ */ jsx(Field, { label: "LI", children: /* @__PURE__ */ jsx("input", { value: ligante.li, onChange: (e) => updateLigante("li", e.target.value), className: inputSm }) }),
      /* @__PURE__ */ jsx(Field, { label: "% (valor)", children: /* @__PURE__ */ jsx("input", { value: ligante.valor, onChange: (e) => updateLigante("valor", e.target.value), className: inputSm }) }),
      /* @__PURE__ */ jsx(Field, { label: "LS", children: /* @__PURE__ */ jsx("input", { value: ligante.ls, onChange: (e) => updateLigante("ls", e.target.value), className: inputSm }) })
    ] }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Resist\xEAncia \xE0 Deforma\xE7\xE3o Permanente" }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "border border-stone-200 rounded-lg p-3", children: [
        /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 text-xs font-semibold text-stone-600 mb-2", children: [
          /* @__PURE__ */ jsx("input", { type: "checkbox", checked: defTaxaDnd, onChange: (e) => setDefTaxaDnd(e.target.checked), className: "accent-amber-600" }),
          " Taxa de deforma\xE7\xE3o (WTSAIR) \u2014 Desempenho N\xE3o Declarado (DND)"
        ] }),
        !defTaxaDnd && /* @__PURE__ */ jsx("input", { value: defTaxa, onChange: (e) => setDefTaxa(e.target.value), className: inputCls, placeholder: "Valor" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "border border-stone-200 rounded-lg p-3", children: [
        /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 text-xs font-semibold text-stone-600 mb-2", children: [
          /* @__PURE__ */ jsx("input", { type: "checkbox", checked: defProfDnd, onChange: (e) => setDefProfDnd(e.target.checked), className: "accent-amber-600" }),
          " Profundidade de rodeira m\xE1xima (PRDAIR) \u2014 DND"
        ] }),
        !defProfDnd && /* @__PURE__ */ jsx("input", { value: defProf, onChange: (e) => setDefProf(e.target.value), className: inputCls, placeholder: "Valor" })
      ] })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: guardado ? "Nova vers\xE3o guardada \u2713" : semAlteracao ? "Sem altera\xE7\xF5es a guardar" : "Guardar Nova Vers\xE3o" }),
    /* @__PURE__ */ jsxs("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5 mt-6", children: [
      "Hist\xF3rico de Atualiza\xE7\xF5es (",
      historicoOrdenado2.length,
      ")"
    ] }),
    /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden", children: historicoOrdenado2.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 px-4 py-4", children: "Ainda sem vers\xF5es guardadas." }) : historicoOrdenado2.map((h, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-4 py-2.5 text-sm ${i !== historicoOrdenado2.length - 1 ? "border-b border-stone-100" : ""} ${i === 0 ? "bg-amber-50/40" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 min-w-0", children: [
        /* @__PURE__ */ jsx("span", { className: "text-stone-700 font-medium truncate", children: h.norma || "Sem norma indicada" }),
        h.codigoDop && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs text-stone-400 truncate", children: [
          "\xB7 ",
          h.codigoDop
        ] }),
        i === 0 && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700 shrink-0", children: "Vers\xE3o atual" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 text-xs text-stone-400 shrink-0", children: [
        /* @__PURE__ */ jsxs("span", { children: [
          formatDateTimePT(h.dataRegisto),
          " \u2014 ",
          h.utilizador || "\u2014"
        ] }),
        /* @__PURE__ */ jsx("button", { onClick: () => exportarDop(h), title: "Exportar esta vers\xE3o", className: "p-1 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(FileText, { size: 13 }) }),
        isAdmin && /* @__PURE__ */ jsx("button", { onClick: () => removerHistorico(h.id), title: "Apagar esta vers\xE3o", className: "p-1 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 13 }) })
      ] })
    ] }, h.id)) })
  ] });
}
function EditarDataHistoricoFormulaModal({ entry, onSave, onClose }) {
  const [valor, setValor] = useState(isoToDatetimeLocal(entry.data));
  const [error, setError] = useState("");
  const submit = () => {
    const iso = datetimeLocalToIso(valor);
    if (!iso) return setError("Essa data n\xE3o existe no calend\xE1rio \u2014 verifique o dia e o m\xEAs");
    onSave(iso);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Data do Registo", subtitle: "Hist\xF3rico de Altera\xE7\xF5es", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Data e Hora", children: /* @__PURE__ */ jsx("input", { type: "datetime-local", value: valor, onChange: (e) => setValor(e.target.value), className: inputCls, autoFocus: true }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function FormulaModal({ data, readOnly, isAdmin, materiaisDisponiveis, tiposMaterial, center, artigos, formulasDoCentro, onUpdateLabel, onDeleteHistorico, onEditHistoricoData, dopConfig, logotipo, onSaveDop, onDeleteDopHistorico, nomeUtilizadorAtual, onSave, onClose }) {
  const tipoIdPorNome = (nome) => (tiposMaterial || []).find((t) => normalizeHeader(t.nome) === normalizeHeader(nome))?.id;
  const filtrarPorTipo = (nomeTipo) => {
    const tid = tipoIdPorNome(nomeTipo);
    if (!tid) return materiaisDisponiveis || [];
    return (materiaisDisponiveis || []).filter((m) => !m.tipoMaterialId || m.tipoMaterialId === tid);
  };
  const materiaisAgregado = filtrarPorTipo("Agregado");
  const materiaisFiller = filtrarPorTipo("Filler Comercial");
  const materiaisAditivo = filtrarPorTipo("Aditivo");
  const materiaisBetume = filtrarPorTipo("Betume");
  const materiaisParaTrabalho = (key) => {
    if (key === "fillerRec" || key === "fillerCom") return materiaisFiller;
    if (key === "aditivo1" || key === "aditivo2" || key === "aditivo3") return materiaisAditivo;
    if (key === "ligante") return materiaisBetume;
    return materiaisAgregado;
  };
  const [dopOpen, setDopOpen] = useState(false);
  const [artigoId, setArtigoId] = useState(data?.artigoId || "");
  const [estudo, setEstudo] = useState(data?.estudo || "");
  const [dataEstudo, setDataEstudo] = useState(data?.dataEstudo || "");
  const [central, setCentral] = useState(data?.central || "");
  const [observacoes, setObservacoes] = useState(data?.observacoes || "");
  const [silos, setSilos] = useState(() => {
    const base = data?.silos || {};
    return Object.fromEntries(SILO_COLS.map((c) => {
      const entry = base[c.key] || {};
      return [c.key, { materialId: entry.materialId || "", pct: entry.materialId ? entry.pct || "" : "" }];
    }));
  });
  const [trabalho, setTrabalho] = useState(() => Object.fromEntries(TRABALHO_COLS.map((c) => [c.key, {
    design: data?.trabalho?.[c.key]?.design || "",
    materialId: data?.trabalho?.[c.key]?.materialId || ""
  }])));
  const [labelDrafts, setLabelDrafts] = useState({});
  const labelValue = (key) => labelDrafts[key] ?? getTrabalhoLabel(center, key);
  const handleLabelChange = (key, value) => setLabelDrafts((prev) => ({ ...prev, [key]: value }));
  const handleLabelBlur = (key) => {
    if (labelDrafts[key] !== void 0) onUpdateLabel(key, labelDrafts[key]);
  };
  const [error, setError] = useState("");
  const [editHistEntry, setEditHistEntry] = useState(null);
  const artigoSelecionado = (artigos || []).find((a) => a.id === artigoId);
  const designacao = artigoSelecionado?.designacao || "";
  const artigoMudou = artigoId !== (data?.artigoId || "");
  const codigoBase = artigoSelecionado?.codigoManual || "";
  const irmasDoArtigo = (formulasDoCentro || []).filter((f) => f.artigoId === artigoId && f.id !== data?.id);
  const codigoCalculado = artigoId ? irmasDoArtigo.length === 0 ? codigoBase : `${codigoBase}.${irmasDoArtigo.length + 1}` : "";
  const codigoInicial = data?.id && !artigoMudou && data.codigo ? data.codigo : codigoCalculado;
  const [codigoManualState, setCodigoManualState] = useState(codigoInicial);
  const codigo = artigoMudou ? codigoCalculado : codigoManualState;
  const updateGrid = (setter, state, key, field, value) => {
    setter({ ...state, [key]: { ...state[key], [field]: value } });
  };
  const updateSiloMaterial = (key, materialId) => {
    setSilos((prev) => ({ ...prev, [key]: { materialId, pct: materialId ? prev[key]?.pct : "" } }));
  };
  const totalSilosPct = SILO_COLS.reduce((s, c) => s + (parseFloat(silos[c.key]?.pct) || 0), 0);
  const totalTrabalho = TRABALHO_COLS.reduce((s, c) => s + (parseFloat(trabalho[c.key]?.design) || 0), 0);
  const silosOk = Math.round(totalSilosPct) === 100;
  const trabalhoOk = Math.round(totalTrabalho) === 1e3;
  const submit = () => {
    if (!artigoId) return setError("Escolha o artigo \u2014 n\xE3o pode haver f\xF3rmula sem artigo");
    if (!silosOk) return setError("A soma da Dosifica\xE7\xE3o de Agregados a Frio tem de ser 100%");
    if (!trabalhoOk) return setError("A soma da F\xF3rmula de Trabalho a Quente tem de ser 1000 kg");
    onSave({
      id: data?.id,
      centroId: data.centroId,
      artigoId,
      codigo,
      estudo: estudo.trim(),
      dataEstudo,
      designacao,
      central: central.trim ? central.trim() : central,
      silos,
      trabalho,
      observacoes: observacoes.trim()
    });
  };
  const gridCls = "w-full px-2 py-1.5 rounded-md border border-stone-300 bg-white text-stone-800 focus:outline-none focus:ring-1 focus:ring-amber-500 text-xs disabled:bg-stone-100 disabled:text-stone-400";
  const historicoOrdenado2 = [...data?.historico || []].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar F\xF3rmula" : "Nova F\xF3rmula", subtitle: "Lista de F\xF3rmulas", onClose, wide: true, children: [
    /* @__PURE__ */ jsx(Field, { label: "Artigo", children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: artigoId, materiais: (artigos || []).map((a) => ({ ...a, fornecedor: a.codigoManual })), disabled: readOnly, semAbreviar: true, onChange: setArtigoId }) }),
    /* @__PURE__ */ jsx("p", { className: "text-[11px] text-stone-400 -mt-2 mb-3", children: 'A designa\xE7\xE3o e o c\xF3digo da f\xF3rmula v\xEAm do artigo escolhido \u2014 n\xE3o \xE9 poss\xEDvel haver f\xF3rmula sem artigo. Se o mesmo artigo tiver mais do que uma f\xF3rmula, o c\xF3digo fica com ".2", ".3"... a seguir ao c\xF3digo do artigo.' }),
    data?.id && /* @__PURE__ */ jsxs("button", { type: "button", onClick: () => setDopOpen(true), className: "w-full mb-3 flex items-center justify-center gap-1.5 py-2.5 rounded-lg border border-amber-300 text-amber-700 text-sm font-semibold hover:bg-amber-50", children: [
      /* @__PURE__ */ jsx(ShieldCheck, { size: 15 }),
      " Declara\xE7\xE3o de Desempenho (DoP)"
    ] }),
    dopOpen && /* @__PURE__ */ jsx(
      DopFormulaModal,
      {
        formula: data,
        center,
        dopConfig,
        logotipo,
        isAdmin,
        nomeUtilizadorAtual,
        onSave: (dopData) => onSaveDop(data.id, dopData),
        onDeleteHistorico: (entryId) => onDeleteDopHistorico(data.id, entryId),
        onClose: () => setDopOpen(false)
      }
    ),
    !artigoId && (data?.codigo || data?.designacao) && /* @__PURE__ */ jsxs("p", { className: "text-xs text-amber-700 bg-amber-50 rounded-lg px-3 py-2 -mt-1 mb-3", children: [
      "Esta f\xF3rmula ainda n\xE3o est\xE1 ligada a um artigo (dados antigos: ",
      /* @__PURE__ */ jsx("strong", { children: data.codigo }),
      " \u2014 ",
      data.designacao,
      "). Escolha o artigo correspondente acima para atualizar."
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 sm:grid-cols-4 gap-3", children: [
      /* @__PURE__ */ jsxs(Field, { label: "F\xF3rmula (c\xF3digo)", children: [
        /* @__PURE__ */ jsx(
          "input",
          {
            value: codigo,
            onChange: isAdmin && !artigoMudou ? ((e) => setCodigoManualState(e.target.value)) : void 0,
            disabled: !isAdmin || artigoMudou,
            className: `${inputCls} font-mono-data ${!isAdmin || artigoMudou ? "bg-stone-100 text-stone-500" : ""}`,
            placeholder: "Escolha o artigo..."
          }
        ),
        isAdmin && !artigoMudou && /* @__PURE__ */ jsxs("span", { className: "block text-[11px] text-stone-400 mt-1", children: [
          "Calculado automaticamente (",
          codigoCalculado || "\u2014",
          "), mas pode corrigir aqui se necess\xE1rio."
        ] })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Estudo", children: /* @__PURE__ */ jsx("input", { value: estudo, onChange: (e) => setEstudo(e.target.value), disabled: readOnly, className: `${inputCls} disabled:bg-stone-100`, placeholder: "A16.165.01" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Data do Estudo", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEstudo, onChange: (e) => setDataEstudo(e.target.value), disabled: readOnly, className: `${inputCls} disabled:bg-stone-100` }) }),
      /* @__PURE__ */ jsx(Field, { label: "N\xBA (Central)", children: /* @__PURE__ */ jsx("input", { value: central, onChange: (e) => setCentral(e.target.value), disabled: readOnly, className: `${inputCls} font-mono-data disabled:bg-stone-100`, placeholder: "2" }) })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, disabled: true, className: `${inputCls} bg-stone-100 text-stone-500`, placeholder: "Vem do artigo escolhido" }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-1.5 mt-2", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-amber-700", children: "Dosifica\xE7\xE3o de Agregados a Frio" }),
      /* @__PURE__ */ jsxs("span", { className: `text-xs font-mono-data font-semibold ${silosOk ? "text-emerald-600" : "text-amber-600"}`, children: [
        totalSilosPct.toLocaleString("pt-PT", { maximumFractionDigits: 1 }),
        "%",
        !silosOk && " (tem de somar 100%)"
      ] })
    ] }),
    /* @__PURE__ */ jsx("div", { className: "border-2 border-amber-200 bg-amber-50/40 rounded-lg overflow-visible mb-4 divide-y divide-amber-200", children: SILO_COLS.map((c) => {
      const temMaterial = !!silos[c.key]?.materialId;
      return /* @__PURE__ */ jsxs("div", { className: `flex items-center gap-3 px-3 py-2 ${temMaterial ? "bg-white" : "bg-amber-50/60"}`, children: [
        /* @__PURE__ */ jsx("span", { className: `text-[11px] font-semibold uppercase tracking-wide w-16 shrink-0 ${temMaterial ? "text-stone-500" : "text-stone-400"}`, children: c.label }),
        /* @__PURE__ */ jsx("div", { className: "flex-1 min-w-0", children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: silos[c.key]?.materialId || "", materiais: materiaisAgregado, disabled: readOnly, isAdmin, onChange: (v) => updateSiloMaterial(c.key, v), compact: true, semAbreviar: true }) }),
        /* @__PURE__ */ jsx(
          "input",
          {
            value: silos[c.key]?.pct || "",
            onChange: (e) => updateGrid(setSilos, silos, c.key, "pct", e.target.value),
            disabled: readOnly || !temMaterial,
            className: `${gridCls} w-20 shrink-0 ${!temMaterial ? "text-stone-300" : ""}`,
            placeholder: temMaterial ? "%" : ""
          }
        )
      ] }, c.key);
    }) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between mb-1.5", children: [
      /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-sky-700", children: "F\xF3rmula de Trabalho a Quente" }),
      /* @__PURE__ */ jsxs("span", { className: `text-xs font-mono-data font-semibold ${trabalhoOk ? "text-emerald-600" : "text-amber-600"}`, children: [
        totalTrabalho.toLocaleString("pt-PT", { maximumFractionDigits: 1 }),
        " kg",
        !trabalhoOk && " (tem de somar 1000 kg)"
      ] })
    ] }),
    /* @__PURE__ */ jsx("p", { className: "text-[11px] text-stone-400 -mt-1 mb-1.5", children: "As designa\xE7\xF5es destas colunas s\xE3o partilhadas por todas as f\xF3rmulas deste centro \u2014 alterar aqui altera em todas." }),
    /* @__PURE__ */ jsx("div", { className: "border-2 border-sky-200 bg-sky-50/40 rounded-lg overflow-visible mb-4", children: /* @__PURE__ */ jsx("div", { className: "grid", style: { gridTemplateColumns: `repeat(4, minmax(0,1fr))` }, children: TRABALHO_COLS.map((c, idx) => {
      const ligadoAMaterial = TRABALHO_MATERIAL_KEYS.includes(c.key);
      return /* @__PURE__ */ jsxs("div", { className: `border-l border-t border-sky-200 ${idx % 4 === 0 ? "border-l-0" : ""} p-2 bg-white`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-center gap-1 mb-1", children: [
          c.num && /* @__PURE__ */ jsxs("span", { className: "relative inline-flex items-center justify-center w-5 h-5 rounded bg-amber-100 text-amber-700 shrink-0", title: "Malha do crivo: d = dimens\xE3o menor, D = dimens\xE3o maior (ex: 5/15)", children: [
            /* @__PURE__ */ jsx(Grid3x3, { size: 12, strokeWidth: 2.5 }),
            /* @__PURE__ */ jsx("span", { className: "absolute -top-1.5 -right-1.5 w-3.5 h-3.5 rounded-full bg-amber-600 text-white text-[8px] font-bold flex items-center justify-center leading-none", children: c.num })
          ] }),
          /* @__PURE__ */ jsx(
            "input",
            {
              value: labelValue(c.key),
              onChange: (e) => handleLabelChange(c.key, e.target.value),
              onBlur: () => handleLabelBlur(c.key),
              disabled: readOnly,
              title: c.num ? "d/D: dimens\xE3o menor / dimens\xE3o maior (ex: 5/15)" : void 0,
              placeholder: c.num ? "d/D" : void 0,
              className: `w-14 text-[10px] font-semibold ${c.num ? "" : "uppercase"} tracking-wide text-stone-600 text-center bg-transparent border-b border-dashed border-stone-300 focus:outline-none focus:border-amber-500 disabled:text-stone-500 disabled:border-transparent placeholder:text-stone-400 placeholder:normal-case`
            }
          ),
          /* @__PURE__ */ jsx("span", { className: "text-[10px] text-stone-400 shrink-0", children: "(kg)" })
        ] }),
        ligadoAMaterial && /* @__PURE__ */ jsx("div", { className: "mb-1", children: /* @__PURE__ */ jsx(MaterialSearchSelect, { value: trabalho[c.key]?.materialId || "", materiais: materiaisParaTrabalho(c.key), disabled: readOnly, alignRight: idx % 4 === 3, isAdmin, onChange: (v) => updateGrid(setTrabalho, trabalho, c.key, "materialId", v), compact: true }) }),
        /* @__PURE__ */ jsx("input", { value: trabalho[c.key]?.design || "", onChange: (e) => updateGrid(setTrabalho, trabalho, c.key, "design", e.target.value), disabled: readOnly, className: gridCls, placeholder: "Quantidade" })
      ] }, c.key);
    }) }) }),
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-600 mb-1.5", children: "Hist\xF3rico de Altera\xE7\xF5es" }),
    /* @__PURE__ */ jsx("div", { className: "border-2 border-stone-300 rounded-lg overflow-hidden mb-4 max-h-40 overflow-y-auto", children: historicoOrdenado2.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 px-3 py-2.5", children: "Ainda sem registos de altera\xE7\xE3o." }) : historicoOrdenado2.map((h) => /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between px-3 py-2 text-xs border-b border-stone-200 last:border-b-0 bg-stone-50/60", children: [
      /* @__PURE__ */ jsxs("span", { className: "text-stone-600", children: [
        formatDateTimePT(h.data),
        " \u2014 ",
        /* @__PURE__ */ jsx("span", { className: "font-medium text-stone-800", children: h.utilizador || "\u2014" })
      ] }),
      isAdmin && /* @__PURE__ */ jsxs("div", { className: "flex gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("button", { type: "button", onClick: () => setEditHistEntry(h), className: "p-1 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(Pencil, { size: 12 }) }),
        /* @__PURE__ */ jsx("button", { type: "button", onClick: () => onDeleteHistorico(data.id, h.id), className: "p-1 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 12 }) })
      ] })
    ] }, h.id)) }),
    editHistEntry && /* @__PURE__ */ jsx(
      EditarDataHistoricoFormulaModal,
      {
        entry: editHistEntry,
        onSave: (novaData) => {
          onEditHistoricoData(data.id, editHistEntry.id, novaData);
          setEditHistEntry(null);
        },
        onClose: () => setEditHistEntry(null)
      }
    ),
    /* @__PURE__ */ jsx(Field, { label: "Obs:", children: /* @__PURE__ */ jsx("textarea", { value: observacoes, onChange: (e) => setObservacoes(e.target.value), disabled: readOnly, rows: 3, spellCheck: "true", lang: "pt-PT", className: `${inputCls} disabled:bg-stone-100`, placeholder: "Observa\xE7\xF5es sobre esta f\xF3rmula..." }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    !readOnly && /* @__PURE__ */ jsx("button", { onClick: submit, disabled: !silosOk || !trabalhoOk, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-50 disabled:cursor-not-allowed", children: "Guardar" })
  ] });
}
function ImportFormulasModal({ centroId, center, materiais, onImport, onClose }) {
  const [registos, setRegistos] = useState(null);
  const [avisos, setAvisos] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const asText = (v) => String(v ?? "").trim();
  const asDate = (v) => {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return asText(v);
  };
  const resolverMaterial = (nome) => {
    const alvo = normalizeHeader(nome);
    if (!alvo) return null;
    const encontrado = materiais.find((m) => normalizeHeader(m.designacao) === alvo);
    return encontrado?.id || null;
  };
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const headerIdx = rows.findIndex((r) => normalizeHeader(r[1]) === "formula");
        if (headerIdx === -1) {
          setError('N\xE3o encontrei o cabe\xE7alho "F\xF3rmula" (coluna B) neste ficheiro. Confirme que segue um dos dois modelos suportados.');
          setRegistos(null);
          return;
        }
        const headerRow = rows[headerIdx] || [];
        const formatoNovo = normalizeHeader(headerRow[6]) === "silo 1";
        const col = (r, idx) => asText(r[idx]);
        const avisosNaoEncontrados = /* @__PURE__ */ new Set();
        const parsed = [];
        let i = headerIdx + 1;
        while (i < rows.length) {
          const row1 = rows[i] || [];
          const codigo = asText(row1[1]);
          if (!codigo) break;
          const row2 = rows[i + 1] || [];
          if (formatoNovo) {
            const silos = {};
            SILO_COLS.forEach((c, idx) => {
              const colIdx = 6 + idx;
              const nome = col(row1, colIdx);
              const pct = col(row2, colIdx);
              const materialId = nome ? resolverMaterial(nome) : null;
              if (nome && !materialId) avisosNaoEncontrados.add(nome);
              silos[c.key] = { materialId, pct };
            });
            const trabalho = {};
            TRABALHO_COLS.forEach((c, idx) => {
              const colIdx = 6 + SILO_COLS.length + idx;
              const nome = col(row1, colIdx);
              const design = col(row2, colIdx);
              const materialId = nome ? resolverMaterial(nome) : null;
              if (nome && !materialId) avisosNaoEncontrados.add(nome);
              trabalho[c.key] = { materialId, design };
            });
            const offCentral = 6 + SILO_COLS.length + TRABALHO_COLS.length;
            parsed.push({
              codigo,
              estudo: col(row1, 2),
              dataEstudo: asDate(row1[3]),
              designacao: col(row1, 4),
              central: col(row1, offCentral),
              dataAlteracao: asDate(row1[offCentral + 1]),
              observacoes: col(row1, offCentral + 2),
              silos,
              trabalho
            });
          } else {
            parsed.push({
              codigo,
              estudo: col(row1, 2),
              dataEstudo: asDate(row1[3]),
              designacao: col(row1, 4),
              central: col(row1, 20),
              dataAlteracao: asDate(row1[21]),
              silos: {
                s1: { design: col(row1, 6), pct: col(row2, 6) },
                s2: { design: col(row1, 7), pct: col(row2, 7) },
                s3: { design: col(row1, 8), pct: col(row2, 8) },
                s4: { design: col(row1, 9), pct: col(row2, 9) },
                s5: { design: col(row1, 10), pct: col(row2, 10) },
                byPass: { design: col(row1, 11), pct: col(row2, 11) }
              },
              trabalho: {
                c1725: { design: col(row1, 12), pct: col(row2, 12) },
                c1017: { design: col(row1, 13), pct: col(row2, 13) },
                c0610: { design: col(row1, 14), pct: col(row2, 14) },
                c006: { design: col(row1, 15), pct: col(row2, 15) },
                fillerRec: { design: col(row1, 16), pct: col(row2, 16) },
                fillerCom: { design: col(row1, 17), pct: col(row2, 17) },
                aditivo: { design: col(row1, 18), pct: col(row2, 18) },
                ligante: { design: col(row1, 19), pct: col(row2, 19) }
              }
            });
          }
          i += 2;
        }
        if (parsed.length === 0) {
          setError("N\xE3o encontrei nenhuma f\xF3rmula com c\xF3digo preenchido a seguir ao cabe\xE7alho.");
          setRegistos(null);
          return;
        }
        setAvisos(formatoNovo ? [...avisosNaoEncontrados] : []);
        setRegistos(parsed);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx ou .xls v\xE1lido, num dos dois formatos suportados.");
        setRegistos(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const confirmImport = () => {
    setImporting(true);
    onImport(centroId, registos);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Lista de F\xF3rmulas", subtitle: "Ficheiro Excel", onClose, children: [
    !registos && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro no formato da ",
        /* @__PURE__ */ jsx("strong", { children: "Lista de F\xF3rmulas por Posto de Produ\xE7\xE3o" }),
        " (documento da empresa) ",
        /* @__PURE__ */ jsx("strong", { children: "ou" }),
        ' um ficheiro exportado pelo bot\xE3o "Exportar Excel" desta aplica\xE7\xE3o \u2014 ambos s\xE3o reconhecidos automaticamente.'
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    registos && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        registos.length,
        " f\xF3rmula",
        registos.length !== 1 ? "s" : "",
        " encontrada",
        registos.length !== 1 ? "s" : "",
        ' em "',
        fileName,
        '"'
      ] }),
      avisos.length > 0 && /* @__PURE__ */ jsxs("div", { className: "flex items-start gap-2 mb-3 text-amber-700 bg-amber-50 rounded-lg px-3 py-2 text-sm", children: [
        /* @__PURE__ */ jsx(AlertTriangle, { size: 16, className: "shrink-0 mt-0.5" }),
        /* @__PURE__ */ jsxs("span", { children: [
          "N\xE3o encontrei estes materiais na base de dados \u2014 as f\xF3rmulas importam-se na mesma, mas esses componentes ficam por ligar: ",
          /* @__PURE__ */ jsx("strong", { children: avisos.join(", ") })
        ] })
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-64 overflow-y-auto mb-4", children: registos.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center gap-3 px-3 py-2 text-sm ${i !== registos.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: r.codigo }),
        /* @__PURE__ */ jsx("span", { className: "text-stone-700 truncate", children: r.designacao || "\u2014" })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRegistos(null);
          setFileName("");
          setAvisos([]);
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirmImport, disabled: importing, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function ImportCentrosCustoModal({ clienteId, clienteDesignacao, centrosCustoExistentes, onImport, onClose }) {
  const [rows, setRows] = useState(null);
  const [ignoradas, setIgnoradas] = useState(0);
  const [duplicadas, setDuplicadas] = useState(0);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const flat = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
  const matchesCliente = (fileCliente) => {
    const a = normalizeHeader(fileCliente);
    const b = normalizeHeader(clienteDesignacao);
    if (!a || !b) return false;
    if (a.length < 4 || b.length < 4) return a === b;
    return a.includes(b) || b.includes(a);
  };
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setRows(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const codigoKey = headerKeys.find((k) => flat(k) === "codigo");
        const designacaoKey = headerKeys.find((k) => flat(k) === "designacao");
        const clienteKey = headerKeys.find((k) => flat(k) === "cliente");
        const localKey = headerKeys.find((k) => flat(k) === "local" || flat(k) === "localidade");
        const cpKey = headerKeys.find((k) => flat(k) === "codigopostal");
        if (!designacaoKey) {
          setError('N\xE3o encontrei a coluna "Designa\xE7\xE3o" na primeira linha da folha.');
          setRows(null);
          return;
        }
        const codigosExistentes = new Set((centrosCustoExistentes || []).map((cc) => String(cc.codigo || "").trim().toLowerCase()).filter(Boolean));
        const codigosNesteFicheiro = /* @__PURE__ */ new Set();
        let ignoradasCount = 0;
        let duplicadasCount = 0;
        const parsed = [];
        json.forEach((r) => {
          const designacao = String(r[designacaoKey] ?? "").trim();
          if (!designacao) return;
          const clienteTxt = clienteKey ? String(r[clienteKey] ?? "").trim() : "";
          if (clienteKey && (!clienteTxt || !matchesCliente(clienteTxt))) {
            ignoradasCount++;
            return;
          }
          const codigo = codigoKey ? String(r[codigoKey] ?? "").trim() : "";
          const codigoNorm = codigo.toLowerCase();
          if (codigo && (codigosExistentes.has(codigoNorm) || codigosNesteFicheiro.has(codigoNorm))) {
            duplicadasCount++;
            return;
          }
          if (codigo) codigosNesteFicheiro.add(codigoNorm);
          parsed.push({
            codigo,
            designacao,
            localidade: localKey ? String(r[localKey] ?? "").trim() : "",
            codigoPostal: cpKey ? String(r[cpKey] ?? "").trim() : ""
          });
        });
        if (parsed.length === 0) {
          setError(`N\xE3o encontrei nenhuma linha nova para "${clienteDesignacao}" neste ficheiro.`);
          setRows(null);
          return;
        }
        setRows(parsed);
        setIgnoradas(ignoradasCount);
        setDuplicadas(duplicadasCount);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const confirm = () => {
    setImporting(true);
    onImport(clienteId, rows);
  };
  const baixarModelo = () => {
    const wsData = [
      ["C\xF3digo", "Designa\xE7\xE3o", "Cliente", "Local", "C\xF3digo Postal"],
      ["4130", "Grande Repara\xE7\xE3o de Pavimento Grupo III (A7 Selho_Fafe Sul e A11 Guimar\xE3es Oeste_Selho)", clienteDesignacao, "Guimar\xE3es", "4800-000"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 10 }, { wch: 50 }, { wch: 28 }, { wch: 18 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Obras");
    XLSX.writeFile(wb, "modelo_importacao_obras.xlsx");
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Centros de Custo", subtitle: `Cliente: ${clienteDesignacao}`, onClose, children: [
    !rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro .xlsx, .xls ou .csv com as colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Designa\xE7\xE3o" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Cliente" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Local" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "C\xF3digo Postal" }),
        " (a coluna ",
        /* @__PURE__ */ jsx("strong", { children: "C\xF3digo" }),
        " \xE9 opcional). S\xF3 s\xE3o importadas as linhas cujo ",
        /* @__PURE__ */ jsx("strong", { children: "Cliente" }),
        ' corresponda a "',
        clienteDesignacao,
        '" \u2014 as restantes s\xE3o ignoradas automaticamente.'
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: baixarModelo, className: "flex items-center gap-1.5 mb-4 px-3.5 py-2 bg-white border border-stone-300 text-stone-700 rounded-lg text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Descarregar Modelo"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        rows.length,
        " centro",
        rows.length !== 1 ? "s" : "",
        " de custo encontrado",
        rows.length !== 1 ? "s" : "",
        ' para "',
        clienteDesignacao,
        '"'
      ] }),
      ignoradas > 0 && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 mb-1", children: [
        ignoradas,
        " linha",
        ignoradas !== 1 ? "s" : "",
        " de outros clientes foi",
        ignoradas !== 1 ? "ram" : "",
        " ignorada",
        ignoradas !== 1 ? "s" : "",
        "."
      ] }),
      duplicadas > 0 && /* @__PURE__ */ jsxs("p", { className: "text-xs text-amber-600 mb-3", children: [
        duplicadas,
        " linha",
        duplicadas !== 1 ? "s" : "",
        " ignorada",
        duplicadas !== 1 ? "s" : "",
        " por j\xE1 existir",
        duplicadas !== 1 ? "em" : "",
        " uma obra com esse n\xFAmero."
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-64 overflow-y-auto mb-4", children: rows.map((r, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center gap-3 px-3 py-2 text-sm ${i !== rows.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        r.codigo && /* @__PURE__ */ jsx("span", { className: "font-mono-data text-xs font-semibold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded shrink-0", children: r.codigo }),
        /* @__PURE__ */ jsx("span", { className: "text-stone-700 truncate", children: r.designacao }),
        r.localidade && /* @__PURE__ */ jsx("span", { className: "text-stone-400 text-xs shrink-0", children: r.localidade })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRows(null);
          setFileName("");
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirm, disabled: importing, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function AvariaModal({ data, onSave, onClose }) {
  const [dataOcorrencia, setDataOcorrencia] = useState(data?.data || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [descricao, setDescricao] = useState(data?.descricao || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!dataOcorrencia) return setError("Indique a data");
    if (!descricao.trim()) return setError("Descreva a incid\xEAncia/avaria");
    onSave({ id: data?.id, centroId: data.centroId, data: dataOcorrencia, descricao: descricao.trim() });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? "Editar Incid\xEAncia" : "Nova Incid\xEAncia", subtitle: "Incid\xEAncias / Avarias", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Data", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataOcorrencia, onChange: (e) => setDataOcorrencia(e.target.value), className: inputCls, autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Descri\xE7\xE3o", children: /* @__PURE__ */ jsx("textarea", { value: descricao, onChange: (e) => setDescricao(e.target.value), className: inputCls, rows: 4, spellCheck: "true", lang: "pt-PT", placeholder: "Descreva a incid\xEAncia ou avaria..." }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function ResolucaoModal({ item, canManage, onSave, onRemoveResolucao, onClose }) {
  const [resolucaoData, setResolucaoData] = useState(item.resolucaoData || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [resolucaoDescricao, setResolucaoDescricao] = useState(item.resolucaoDescricao || "");
  const [error, setError] = useState("");
  const jaResolvida = !!item.resolucaoData;
  const dataInvalida = !!(item.data && resolucaoData && resolucaoData < item.data);
  const dias = item.data ? Math.round((new Date(resolucaoData) - new Date(item.data)) / 864e5) : null;
  const submit = () => {
    if (!resolucaoData) return setError("Indique a data de resolu\xE7\xE3o");
    if (dataInvalida) return setError("A data de resolu\xE7\xE3o n\xE3o pode ser anterior \xE0 da incid\xEAncia");
    if (!resolucaoDescricao.trim()) return setError("Descreva brevemente o que foi feito");
    onSave(item, resolucaoData, resolucaoDescricao.trim());
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Registar Resolu\xE7\xE3o", subtitle: "Incid\xEAncias / Avarias", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-4", children: [
      /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 mb-1", children: [
        "Incid\xEAncia de ",
        item.data
      ] }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-700", children: item.descricao })
    ] }),
    /* @__PURE__ */ jsxs(Field, { label: "Data de Resolu\xE7\xE3o", children: [
      /* @__PURE__ */ jsx(
        "input",
        {
          type: "date",
          value: resolucaoData,
          onChange: (e) => setResolucaoData(e.target.value),
          autoFocus: true,
          className: `${inputCls} ${dataInvalida ? "border-red-400 focus:ring-red-400 focus:border-red-400" : ""}`
        }
      ),
      dataInvalida && /* @__PURE__ */ jsxs("span", { className: "block text-xs text-red-600 mt-1", children: [
        "N\xE3o pode ser anterior \xE0 data da incid\xEAncia (",
        item.data,
        ")"
      ] })
    ] }),
    !dataInvalida && dias !== null && dias >= 0 && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 -mt-2 mb-4", children: [
      "Tempo decorrido: ",
      dias,
      " dia",
      dias !== 1 ? "s" : "",
      " (calculado automaticamente)"
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "O que foi feito", children: /* @__PURE__ */ jsx("textarea", { value: resolucaoDescricao, onChange: (e) => setResolucaoDescricao(e.target.value), className: inputCls, rows: 3, spellCheck: "true", lang: "pt-PT", placeholder: "Breve descri\xE7\xE3o da resolu\xE7\xE3o..." }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" }),
    jaResolvida && canManage && /* @__PURE__ */ jsx("button", { onClick: () => onRemoveResolucao(item), className: "w-full mt-2 py-2.5 rounded-lg border border-red-300 text-red-600 text-sm font-semibold hover:bg-red-50", children: 'Reverter para "Por Resolver"' })
  ] });
}
function EditarIncidenciaDiariaModal({ item, onSave, onClose }) {
  const [descricao, setDescricao] = useState(item.descricao || "");
  const [error, setError] = useState("");
  const submit = () => {
    if (!descricao.trim()) return setError("Descreva a incid\xEAncia");
    onSave(item.diariaId, item.incidentId, descricao.trim());
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Incid\xEAncia", subtitle: "Parte Di\xE1ria", onClose, children: [
    /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-3", children: "Esta incid\xEAncia veio de uma di\xE1ria de produ\xE7\xE3o. A altera\xE7\xE3o afeta s\xF3 a descri\xE7\xE3o, mantendo a data e o turno originais." }),
    /* @__PURE__ */ jsx(Field, { label: "Descri\xE7\xE3o", children: /* @__PURE__ */ jsx("textarea", { value: descricao, onChange: (e) => setDescricao(e.target.value), className: inputCls, rows: 3, autoFocus: true }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar altera\xE7\xF5es" })
  ] });
}
function CustosExtraEditor({ custosExtra, onChange, tipos }) {
  const nomesDisponiveis = (tipos || []).map((t) => t.nome);
  const addCusto = () => {
    onChange([...custosExtra, { id: genId(), nome: nomesDisponiveis[0] || "", valor: "" }]);
  };
  const updateCusto = (id, field, value) => onChange(custosExtra.map((c) => c.id === id ? { ...c, [field]: value } : c));
  const removeCusto = (id) => onChange(custosExtra.filter((c) => c.id !== id));
  return /* @__PURE__ */ jsxs(Fragment, { children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Custos Extra" }),
    /* @__PURE__ */ jsx("div", { className: "space-y-2 mb-2", children: custosExtra.map((c) => /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
      /* @__PURE__ */ jsxs("select", { value: c.nome, onChange: (e) => updateCusto(c.id, "nome", e.target.value), className: `${inputCls} flex-1`, children: [
        nomesDisponiveis.length === 0 && /* @__PURE__ */ jsx("option", { value: "", children: "\u2014 Nenhum tipo definido \u2014" }),
        nomesDisponiveis.map((n) => /* @__PURE__ */ jsx("option", { value: n, children: n }, n))
      ] }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: c.valor,
          onChange: (e) => updateCusto(c.id, "valor", e.target.value),
          type: "number",
          step: "0.01",
          min: "0",
          className: `${inputCls} font-mono-data w-28`,
          placeholder: "0.00"
        }
      ),
      /* @__PURE__ */ jsx("button", { type: "button", onClick: () => removeCusto(c.id), className: "text-stone-400 hover:text-red-600 shrink-0", children: /* @__PURE__ */ jsx(X, { size: 16 }) })
    ] }, c.id)) }),
    /* @__PURE__ */ jsxs("button", { type: "button", onClick: addCusto, className: "w-full mb-4 py-2 rounded-lg border border-dashed border-stone-300 text-stone-500 text-sm font-semibold hover:bg-stone-50 flex items-center justify-center gap-1.5", children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Adicionar custo extra"
    ] })
  ] });
}
function DescontosEditor({ descontos, onChange, categorias }) {
  const listaCategorias = categorias || DESCONTO_CATEGORIAS;
  const addDesconto = () => {
    if (descontos.length >= 3) return;
    onChange([...descontos, { id: genId(), tipo: "%", valor: "", categoria: listaCategorias[0], outroTexto: "", aplicarNoCalculo: true }]);
  };
  const updateDesconto = (id, field, value) => onChange(descontos.map((d) => d.id === id ? { ...d, [field]: value } : d));
  const removeDesconto = (id) => onChange(descontos.filter((d) => d.id !== id));
  return /* @__PURE__ */ jsxs(Fragment, { children: [
    /* @__PURE__ */ jsx("span", { className: "block text-xs font-semibold uppercase tracking-wide text-stone-500 mb-1.5", children: "Descontos (at\xE9 3)" }),
    /* @__PURE__ */ jsx("div", { className: "space-y-2 mb-2", children: descontos.map((d) => /* @__PURE__ */ jsxs("div", { className: "border border-stone-200 rounded-lg p-3 bg-stone-50/60 relative", children: [
      /* @__PURE__ */ jsx("button", { type: "button", onClick: () => removeDesconto(d.id), className: "absolute top-2 right-2 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(X, { size: 15 }) }),
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-2 mb-2 pr-6", children: [
        /* @__PURE__ */ jsxs("div", { className: "flex rounded-lg border border-stone-300 overflow-hidden", children: [
          /* @__PURE__ */ jsx(
            "button",
            {
              type: "button",
              onClick: () => updateDesconto(d.id, "tipo", "%"),
              className: `flex-1 py-2 text-xs font-semibold ${d.tipo === "%" ? "bg-amber-100 text-amber-800" : "bg-white text-stone-500"}`,
              children: "%"
            }
          ),
          /* @__PURE__ */ jsx(
            "button",
            {
              type: "button",
              onClick: () => updateDesconto(d.id, "tipo", "fixo"),
              className: `flex-1 py-2 text-xs font-semibold border-l border-stone-300 ${d.tipo === "fixo" ? "bg-amber-100 text-amber-800" : "bg-white text-stone-500"}`,
              children: "Valor fixo (\u20AC)"
            }
          )
        ] }),
        /* @__PURE__ */ jsx(
          "input",
          {
            value: d.valor,
            onChange: (e) => updateDesconto(d.id, "valor", e.target.value),
            type: "number",
            step: "0.01",
            min: "0",
            className: `${inputCls} font-mono-data`,
            placeholder: d.tipo === "%" ? "%" : "\u20AC"
          }
        )
      ] }),
      /* @__PURE__ */ jsx("select", { value: d.categoria, onChange: (e) => updateDesconto(d.id, "categoria", e.target.value), className: `${inputCls} ${d.categoria === "Outro" ? "mb-2" : ""}`, children: listaCategorias.map((c) => /* @__PURE__ */ jsx("option", { value: c, children: c }, c)) }),
      d.categoria === "Outro" && /* @__PURE__ */ jsx("input", { value: d.outroTexto, onChange: (e) => updateDesconto(d.id, "outroTexto", e.target.value), className: `${inputCls} mb-2`, placeholder: "Descreva o tipo de desconto" }),
      /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 cursor-pointer", children: [
        /* @__PURE__ */ jsx(
          "input",
          {
            type: "checkbox",
            checked: d.aplicarNoCalculo !== false,
            onChange: (e) => updateDesconto(d.id, "aplicarNoCalculo", e.target.checked),
            className: "w-4 h-4 accent-amber-600 cursor-pointer"
          }
        ),
        /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-500", children: "Considerar no c\xE1lculo do custo da mistura" })
      ] }),
      d.aplicarNoCalculo === false && /* @__PURE__ */ jsx("p", { className: "text-[11px] text-amber-600 mt-1 pl-6", children: "Fica registado s\xF3 para mapas comparativos \u2014 n\xE3o entra no pre\xE7o usado na Ficha de Custo." })
    ] }, d.id)) }),
    descontos.length < 3 && /* @__PURE__ */ jsxs("button", { type: "button", onClick: addDesconto, className: "w-full mb-4 py-2 rounded-lg border border-dashed border-stone-300 text-stone-500 text-sm font-semibold hover:bg-stone-50 flex items-center justify-center gap-1.5", children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Adicionar desconto"
    ] })
  ] });
}
function validarDescontos(descontos, setError) {
  for (const d of descontos) {
    if (d.valor === "" || parseFloat(d.valor) < 0) {
      setError("Indique o valor de cada desconto");
      return false;
    }
    if (d.categoria === "Outro" && !d.outroTexto.trim()) {
      setError('Descreva o tipo de desconto "Outro"');
      return false;
    }
  }
  return true;
}
function GerirTiposModal({ titulo, subtitulo, descricao, tipos, placeholderNovo, onSave, onDelete, onClose }) {
  const [novoNome, setNovoNome] = useState("");
  const [editandoId, setEditandoId] = useState(null);
  const [editandoNome, setEditandoNome] = useState("");
  const [error, setError] = useState("");
  const adicionar = () => {
    if (!novoNome.trim()) return setError("Indique o nome");
    if (tipos.some((t) => t.nome.trim().toLowerCase() === novoNome.trim().toLowerCase())) {
      return setError("J\xE1 existe um com esse nome");
    }
    onSave({ nome: novoNome.trim() });
    setNovoNome("");
    setError("");
  };
  const iniciarEdicao = (t) => {
    setEditandoId(t.id);
    setEditandoNome(t.nome);
    setError("");
  };
  const guardarEdicao = () => {
    if (!editandoNome.trim()) return setError("Indique o nome");
    if (tipos.some((t) => t.id !== editandoId && t.nome.trim().toLowerCase() === editandoNome.trim().toLowerCase())) {
      return setError("J\xE1 existe um com esse nome");
    }
    onSave({ id: editandoId, nome: editandoNome.trim() });
    setEditandoId(null);
    setError("");
  };
  return /* @__PURE__ */ jsxs(Modal, { title: titulo, subtitle: subtitulo, onClose, children: [
    descricao && /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mb-4", children: descricao }),
    /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-4", children: tipos.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 px-3 py-4 text-center", children: "Ainda n\xE3o h\xE1 nenhum definido." }) : tipos.map((t, i) => /* @__PURE__ */ jsx("div", { className: `flex items-center justify-between px-3 py-2.5 ${i !== tipos.length - 1 ? "border-b border-stone-100" : ""}`, children: editandoId === t.id ? /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 flex-1", children: [
      /* @__PURE__ */ jsx(
        "input",
        {
          value: editandoNome,
          onChange: (e) => setEditandoNome(e.target.value),
          className: `${inputCls} flex-1`,
          autoFocus: true,
          onKeyDown: (e) => e.key === "Enter" && guardarEdicao()
        }
      ),
      /* @__PURE__ */ jsx("button", { onClick: guardarEdicao, className: "text-xs font-semibold text-amber-700 hover:text-amber-800 shrink-0", children: "Guardar" }),
      /* @__PURE__ */ jsx("button", { onClick: () => setEditandoId(null), className: "text-xs text-stone-400 hover:text-stone-600 shrink-0", children: "Cancelar" })
    ] }) : /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm text-stone-700", children: t.nome }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("span", { onClick: () => iniciarEdicao(t), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg cursor-pointer", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("span", { onClick: () => onDelete(t.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg cursor-pointer", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] })
    ] }) }, t.id)) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
      /* @__PURE__ */ jsx(
        "input",
        {
          value: novoNome,
          onChange: (e) => setNovoNome(e.target.value),
          placeholder: placeholderNovo,
          className: `${inputCls} flex-1`,
          onKeyDown: (e) => e.key === "Enter" && adicionar()
        }
      ),
      /* @__PURE__ */ jsx("button", { onClick: adicionar, className: "px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700 shrink-0", children: "Adicionar" })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
  ] });
}
function TiposMaterialModal({ tiposMaterial, materiais, onSave, onDelete, onClose }) {
  const [novoNome, setNovoNome] = useState("");
  const [editandoId, setEditandoId] = useState(null);
  const [editandoNome, setEditandoNome] = useState("");
  const [error, setError] = useState("");
  const contarUso = (id) => materiais.filter((m) => m.tipoMaterialId === id).length;
  const adicionar = () => {
    if (!novoNome.trim()) return setError("Indique o nome do tipo");
    if (tiposMaterial.some((t) => t.nome.trim().toLowerCase() === novoNome.trim().toLowerCase())) {
      return setError("J\xE1 existe um tipo com esse nome");
    }
    onSave({ nome: novoNome.trim() });
    setNovoNome("");
    setError("");
  };
  const iniciarEdicao = (t) => {
    setEditandoId(t.id);
    setEditandoNome(t.nome);
    setError("");
  };
  const guardarEdicao = () => {
    if (!editandoNome.trim()) return setError("Indique o nome do tipo");
    if (tiposMaterial.some((t) => t.id !== editandoId && t.nome.trim().toLowerCase() === editandoNome.trim().toLowerCase())) {
      return setError("J\xE1 existe um tipo com esse nome");
    }
    onSave({ id: editandoId, nome: editandoNome.trim() });
    setEditandoId(null);
    setError("");
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Gerir Tipos de Material", subtitle: "Materiais Constituintes", onClose, children: [
    /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 mb-4", children: 'Usados para filtrar as pesquisas de materiais nas f\xF3rmulas e nas rece\xE7\xF5es \u2014 ao escolher, por exemplo, "Agregado", s\xF3 aparecem materiais desse tipo.' }),
    /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden mb-4", children: tiposMaterial.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-400 px-3 py-4 text-center", children: "Ainda n\xE3o h\xE1 tipos definidos." }) : tiposMaterial.map((t, i) => /* @__PURE__ */ jsx("div", { className: `flex items-center justify-between px-3 py-2.5 ${i !== tiposMaterial.length - 1 ? "border-b border-stone-100" : ""}`, children: editandoId === t.id ? /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 flex-1", children: [
      /* @__PURE__ */ jsx(
        "input",
        {
          value: editandoNome,
          onChange: (e) => setEditandoNome(e.target.value),
          className: `${inputCls} flex-1`,
          autoFocus: true,
          onKeyDown: (e) => e.key === "Enter" && guardarEdicao()
        }
      ),
      /* @__PURE__ */ jsx("button", { onClick: guardarEdicao, className: "text-xs font-semibold text-amber-700 hover:text-amber-800 shrink-0", children: "Guardar" }),
      /* @__PURE__ */ jsx("button", { onClick: () => setEditandoId(null), className: "text-xs text-stone-400 hover:text-stone-600 shrink-0", children: "Cancelar" })
    ] }) : /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
        /* @__PURE__ */ jsx("span", { className: "text-sm text-stone-700", children: t.nome }),
        contarUso(t.id) > 0 && /* @__PURE__ */ jsxs("span", { className: "text-xs text-stone-400 ml-2", children: [
          contarUso(t.id),
          " material",
          contarUso(t.id) > 1 ? "ais" : ""
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-1 shrink-0", children: [
        /* @__PURE__ */ jsx("span", { onClick: () => iniciarEdicao(t), className: "p-1.5 text-stone-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg cursor-pointer", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("span", { onClick: () => onDelete(t.id), className: "p-1.5 text-stone-400 hover:text-red-600 hover:bg-red-50 rounded-lg cursor-pointer", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] })
    ] }) }, t.id)) }),
    /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
      /* @__PURE__ */ jsx(
        "input",
        {
          value: novoNome,
          onChange: (e) => setNovoNome(e.target.value),
          placeholder: "Novo tipo (ex: Agregado)",
          className: `${inputCls} flex-1`,
          onKeyDown: (e) => e.key === "Enter" && adicionar()
        }
      ),
      /* @__PURE__ */ jsx("button", { onClick: adicionar, className: "px-4 py-2.5 rounded-lg bg-amber-600 text-white text-sm font-semibold hover:bg-amber-700 shrink-0", children: "Adicionar" })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
  ] });
}
function MaterialModal({ data, centers, fornecedores, tipo, tiposMaterial, tiposDesconto, tiposCustoExtra, onSave, onClose }) {
  const isNew = !data?.id;
  const isConsumivel = tipo === "consumivel";
  const isEquipamento = tipo === "equipamento";
  const isMaoDeObra = tipo === "maodeobra";
  const isMaterial = !isConsumivel && !isEquipamento && !isMaoDeObra;
  const isUnidadeSimples = isEquipamento || isMaoDeObra;
  const nomeSingular = isConsumivel ? "Combust\xEDvel" : isEquipamento ? "Equipamento" : isMaoDeObra ? "Categoria de M\xE3o de Obra" : "Material Constituinte";
  const nomePlural = isConsumivel ? "Combust\xEDveis/Energia" : isEquipamento ? "Equipamentos" : isMaoDeObra ? "M\xE3o de Obra" : "Materiais Constituintes";
  const [designacao, setDesignacao] = useState(data?.designacao || "");
  const [fornecedor, setFornecedor] = useState(data?.fornecedor || "");
  const [tipoMaterialId, setTipoMaterialId] = useState(data?.tipoMaterialId || "");
  const [aplicacao, setAplicacao] = useState(data?.centrosIds === "todos" ? "todos" : Array.isArray(data?.centrosIds) ? data.centrosIds : []);
  const [preco, setPreco] = useState("");
  const [unidadeCusto, setUnidadeCusto] = useState(data?.unidadeCusto || (isMaoDeObra ? UNIDADES_CUSTO_MAO_OBRA[0] : UNIDADES_CUSTO_EQUIPAMENTO[0]));
  const [unidadeConsumo, setUnidadeConsumo] = useState(data?.unidadeCusto || UNIDADES_CONSUMO_COMBUSTIVEL[0]);
  const [descontos, setDescontos] = useState([]);
  const [custosExtra, setCustosExtra] = useState([]);
  const [dataEntradaVigor, setDataEntradaVigor] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const precoFinal = calcularPrecoFinal({ preco, descontos, custosExtra });
  const submit = () => {
    if (!designacao.trim()) return setError("Indique a designa\xE7\xE3o");
    if (aplicacao !== "todos" && aplicacao.length === 0) return setError('Escolha pelo menos um centro, ou "Todos os centros"');
    if (isNew) {
      if (!preco || parseFloat(preco) < 0) return setError("Indique o pre\xE7o");
      if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor do pre\xE7o");
      if (!isUnidadeSimples && !validarDescontos(descontos, setError)) return;
    }
    onSave({
      id: data?.id,
      designacao: designacao.trim(),
      fornecedor: fornecedor.trim(),
      centrosIds: aplicacao,
      ...isMaterial ? { tipoMaterialId: tipoMaterialId || null } : {},
      ...isUnidadeSimples ? { unidadeCusto } : {},
      ...isConsumivel ? { unidadeCusto: unidadeConsumo } : {},
      ...isNew ? {
        preco: parseFloat(preco),
        descontos: isUnidadeSimples ? [] : descontos.map((d) => ({ ...d, valor: parseFloat(d.valor) })),
        custosExtra: isUnidadeSimples ? [] : custosExtra.map((c) => ({ ...c, valor: parseFloat(c.valor) || 0 })),
        dataEntradaVigor
      } : {}
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: data?.id ? `Editar ${nomeSingular}` : `Novo ${isConsumivel ? "Artigo" : nomeSingular}`, subtitle: nomePlural, onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Designa\xE7\xE3o", children: /* @__PURE__ */ jsx("input", { value: designacao, onChange: (e) => setDesignacao(e.target.value), className: inputCls, placeholder: isConsumivel ? "Ex: Gas\xF3leo Rodovi\xE1rio" : isEquipamento ? "Ex: P\xE1 Carregadora" : isMaoDeObra ? "Ex: Encarregado de Central" : "Ex: Bitume 35/50", autoFocus: true }) }),
    /* @__PURE__ */ jsx(Field, { label: "Fornecedor", children: /* @__PURE__ */ jsx(FornecedorSearchSelect, { value: fornecedor, fornecedores: fornecedores || [], onChange: setFornecedor }) }),
    isMaterial && /* @__PURE__ */ jsx(Field, { label: "Tipo de Material", children: /* @__PURE__ */ jsxs("select", { value: tipoMaterialId, onChange: (e) => setTipoMaterialId(e.target.value), className: inputCls, children: [
      /* @__PURE__ */ jsx("option", { value: "", children: "\u2014 N\xE3o definido \u2014" }),
      (tiposMaterial || []).map((t) => /* @__PURE__ */ jsx("option", { value: t.id, children: t.nome }, t.id))
    ] }) }),
    /* @__PURE__ */ jsx(Field, { label: "Aplica-se a", children: /* @__PURE__ */ jsx(CentrosAplicacaoSelect, { value: aplicacao, centers, onChange: setAplicacao }) }),
    isUnidadeSimples && /* @__PURE__ */ jsx(Field, { label: "Unidade de Custo", children: /* @__PURE__ */ jsx("select", { value: unidadeCusto, onChange: (e) => setUnidadeCusto(e.target.value), className: inputCls, children: (isMaoDeObra ? UNIDADES_CUSTO_MAO_OBRA : UNIDADES_CUSTO_EQUIPAMENTO).map((u) => /* @__PURE__ */ jsx("option", { value: u, children: u }, u)) }) }),
    isConsumivel && /* @__PURE__ */ jsxs(Field, { label: "Unidade de Consumo", children: [
      /* @__PURE__ */ jsx("select", { value: unidadeConsumo, onChange: (e) => setUnidadeConsumo(e.target.value), className: inputCls, children: UNIDADES_CONSUMO_COMBUSTIVEL.map((u) => /* @__PURE__ */ jsx("option", { value: u, children: u }, u)) }),
      /* @__PURE__ */ jsx("span", { className: "block text-xs text-stone-400 mt-1", children: "Usada no Bloco T\xE9rmico e nos blocos de Energia da Parametriza\xE7\xE3o de Produ\xE7\xE3o, para saber em que unidade se mede o consumo por tonelada." })
    ] }),
    isNew ? /* @__PURE__ */ jsxs(Fragment, { children: [
      isUnidadeSimples ? /* @__PURE__ */ jsx(Field, { label: "Pre\xE7o inicial (\u20AC)", children: /* @__PURE__ */ jsx("input", { value: preco, onChange: (e) => setPreco(e.target.value), type: "number", step: "0.01", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0.00" }) }) : /* @__PURE__ */ jsxs(Fragment, { children: [
        /* @__PURE__ */ jsx(Field, { label: isConsumivel ? `Pre\xE7o inicial (\u20AC/${unidadeBaseCombustivel(unidadeConsumo)})` : "Pre\xE7o inicial (\u20AC)", children: /* @__PURE__ */ jsx("input", { value: preco, onChange: (e) => setPreco(e.target.value), type: "number", step: "0.01", min: "0", className: `${inputCls} font-mono-data`, placeholder: "0.00" }) }),
        /* @__PURE__ */ jsx(CustosExtraEditor, { custosExtra, onChange: setCustosExtra, tipos: tiposCustoExtra }),
        /* @__PURE__ */ jsx(DescontosEditor, { descontos, onChange: setDescontos, categorias: [...(tiposDesconto || []).map((t) => t.nome), "Outro"] })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Data de Entrada em Vigor", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
      !isUnidadeSimples && /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-2.5 mb-4", children: [
        /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-amber-800", children: "Pre\xE7o final (com descontos, transporte e otimiza\xE7\xE3o)" }),
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-amber-800", children: [
          precoFinal.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " \u20AC"
        ] })
      ] })
    ] }) : /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mb-4", children: 'Para atualizar o pre\xE7o, feche esta janela e use o bot\xE3o "Atualizar pre\xE7o" na lista \u2014 isso preserva o hist\xF3rico de pre\xE7os anteriores.' }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EscolherMaterialModal({ materiais, tipo, onChoose, onClose }) {
  const [query, setQuery] = useState("");
  const nomeSingular = tipo === "consumivel" ? "consum\xEDvel" : tipo === "equipamento" ? "equipamento" : tipo === "maodeobra" ? "categoria de m\xE3o de obra" : "material";
  const filtrados = query ? materiais.filter((m) => matchesSearch(query, m.designacao, m.fornecedor)) : materiais;
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Pre\xE7o", subtitle: `Escolha o ${nomeSingular}`, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "relative mb-3", children: [
      /* @__PURE__ */ jsx(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 text-stone-400", size: 15 }),
      /* @__PURE__ */ jsx(
        "input",
        {
          value: query,
          onChange: (e) => setQuery(e.target.value),
          autoFocus: true,
          placeholder: "Pesquisar por designa\xE7\xE3o ou fornecedor...",
          className: `${inputCls} pl-9`
        }
      )
    ] }),
    /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-72 overflow-y-auto", children: materiais.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 px-3 py-4 text-center", children: [
      "Ainda n\xE3o h\xE1 ",
      nomeSingular,
      "s registados."
    ] }) : filtrados.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 px-3 py-4 text-center", children: [
      "Nenhum ",
      nomeSingular,
      " encontrado."
    ] }) : filtrados.map((m) => {
      const vigente = precoVigente(m);
      return /* @__PURE__ */ jsxs(
        "button",
        {
          type: "button",
          onClick: () => onChoose(m),
          className: "w-full text-left px-3 py-2.5 hover:bg-amber-50 border-b border-stone-100 last:border-b-0 flex items-center justify-between gap-2",
          children: [
            /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
              /* @__PURE__ */ jsx("p", { className: "text-sm font-medium text-stone-800 truncate", children: m.designacao }),
              m.fornecedor && /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 truncate", children: m.fornecedor })
            ] }),
            vigente && /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-xs text-stone-500 shrink-0", children: [
              calcularPrecoFinal(vigente).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
              " \u20AC"
            ] })
          ]
        },
        m.id
      );
    }) })
  ] });
}
function AtualizarPrecoModal({ material, tipo, tiposDesconto, tiposCustoExtra, onSave, onClose }) {
  const isEquipamento = tipo === "equipamento" || tipo === "maodeobra";
  const isConsumivel = tipo === "consumivel";
  const rotuloPreco = isConsumivel ? `Pre\xE7o (\u20AC/${unidadeBaseCombustivel(material.unidadeCusto)})` : "Pre\xE7o (\u20AC)";
  const atual = precoVigente(material);
  const custosExtraIniciais = () => custosExtraLista(atual).map((c) => ({ ...c, id: genId(), valor: String(c.valor) }));
  const [preco, setPreco] = useState(atual ? String(atual.preco ?? "") : "");
  const [descontos, setDescontos] = useState(atual?.descontos ? atual.descontos.map((d) => ({ ...d, id: genId() })) : []);
  const [custosExtra, setCustosExtra] = useState(custosExtraIniciais());
  const [dataEntradaVigor, setDataEntradaVigor] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const precoOriginal = atual ? String(atual.preco ?? "") : "";
  const normalizarDescontos = (arr) => JSON.stringify((arr || []).map((d) => ({ tipo: d.tipo, valor: parseFloat(d.valor || 0), categoria: d.categoria, outroTexto: d.outroTexto || "" })));
  const normalizarCustosExtra = (arr) => JSON.stringify((arr || []).map((c) => ({ nome: c.nome, valor: parseFloat(c.valor || 0) })).sort((a, b) => a.nome.localeCompare(b.nome)));
  const semAlteracao = !!atual && preco === precoOriginal && (isEquipamento || normalizarCustosExtra(custosExtra) === normalizarCustosExtra(custosExtraIniciais())) && (isEquipamento || normalizarDescontos(descontos) === normalizarDescontos(atual.descontos));
  const precoFinal = calcularPrecoFinal({ preco, descontos, custosExtra });
  const submit = () => {
    if (!preco || parseFloat(preco) < 0) return setError("Indique o pre\xE7o");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor do pre\xE7o");
    if (!isEquipamento && !validarDescontos(descontos, setError)) return;
    if (semAlteracao) return setError("N\xE3o h\xE1 nenhuma altera\xE7\xE3o em rela\xE7\xE3o ao pre\xE7o atual.");
    onSave(material.id, {
      preco: parseFloat(preco),
      descontos: isEquipamento ? [] : descontos.map((d) => ({ ...d, valor: parseFloat(d.valor) })),
      custosExtra: isEquipamento ? [] : custosExtra.map((c) => ({ ...c, valor: parseFloat(c.valor) || 0 })),
      dataEntradaVigor
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Pre\xE7o", subtitle: material.designacao, onClose, children: [
    atual && /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-4 text-sm text-stone-600", children: [
      "Pre\xE7o atual: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold", children: [
        calcularPrecoFinal(atual).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " \u20AC"
      ] }),
      atual.dataEntradaVigor && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
        " (em vigor desde ",
        formatDatePT(atual.dataEntradaVigor),
        ")"
      ] }),
      /* @__PURE__ */ jsx("br", {}),
      /* @__PURE__ */ jsx("span", { className: "text-xs text-stone-400", children: "Os campos abaixo v\xEAm pr\xE9-preenchidos com estes valores \u2014 altere s\xF3 o que mudou." })
    ] }),
    isEquipamento ? /* @__PURE__ */ jsx(Field, { label: "Pre\xE7o (\u20AC)", children: /* @__PURE__ */ jsx(
      "input",
      {
        value: preco,
        onChange: (e) => setPreco(e.target.value),
        type: "number",
        step: "0.01",
        min: "0",
        autoFocus: true,
        className: `${inputCls} font-mono-data ${preco === precoOriginal && atual ? "text-stone-400" : "text-stone-800"}`,
        placeholder: "0.00"
      }
    ) }) : /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsx(Field, { label: rotuloPreco, children: /* @__PURE__ */ jsx(
        "input",
        {
          value: preco,
          onChange: (e) => setPreco(e.target.value),
          type: "number",
          step: "0.01",
          min: "0",
          autoFocus: true,
          className: `${inputCls} font-mono-data ${preco === precoOriginal && atual ? "text-stone-400" : "text-stone-800"}`,
          placeholder: "0.00"
        }
      ) }),
      /* @__PURE__ */ jsx(CustosExtraEditor, { custosExtra, onChange: setCustosExtra, tipos: tiposCustoExtra }),
      /* @__PURE__ */ jsx(DescontosEditor, { descontos, onChange: setDescontos, categorias: [...(tiposDesconto || []).map((t) => t.nome), "Outro"] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Data de Entrada em Vigor", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    !isEquipamento && /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-2.5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-amber-800", children: "Pre\xE7o final (com descontos, transporte e otimiza\xE7\xE3o)" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-amber-800", children: [
        precoFinal.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " \u20AC"
      ] })
    ] }),
    semAlteracao && /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600 mb-3", children: "Ainda n\xE3o alterou nada em rela\xE7\xE3o ao pre\xE7o atual." }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx(
      "button",
      {
        onClick: submit,
        disabled: semAlteracao,
        className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-50 disabled:cursor-not-allowed",
        children: "Guardar nova atualiza\xE7\xE3o"
      }
    )
  ] });
}
function EditarPrecoModal({ material, entry, tipo, tiposDesconto, tiposCustoExtra, onSave, onClose }) {
  const isEquipamento = tipo === "equipamento" || tipo === "maodeobra";
  const isConsumivel = tipo === "consumivel";
  const rotuloPreco = isConsumivel ? `Pre\xE7o (\u20AC/${unidadeBaseCombustivel(material.unidadeCusto)})` : "Pre\xE7o (\u20AC)";
  const custosExtraIniciais = () => custosExtraLista(entry).map((c) => ({ ...c, id: genId(), valor: String(c.valor) }));
  const [preco, setPreco] = useState(String(entry.preco ?? ""));
  const [descontos, setDescontos] = useState(entry.descontos ? entry.descontos.map((d) => ({ ...d, id: genId() })) : []);
  const [custosExtra, setCustosExtra] = useState(custosExtraIniciais());
  const [dataEntradaVigor, setDataEntradaVigor] = useState(entry.dataEntradaVigor || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const precoFinal = calcularPrecoFinal({ preco, descontos, custosExtra });
  const submit = () => {
    if (!preco || parseFloat(preco) < 0) return setError("Indique o pre\xE7o");
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor do pre\xE7o");
    if (!isEquipamento && !validarDescontos(descontos, setError)) return;
    onSave(material.id, entry.id, {
      preco: parseFloat(preco),
      descontos: isEquipamento ? [] : descontos.map((d) => ({ ...d, valor: parseFloat(d.valor) })),
      custosExtra: isEquipamento ? [] : custosExtra.map((c) => ({ ...c, valor: parseFloat(c.valor) || 0 })),
      dataEntradaVigor
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Atualiza\xE7\xE3o de Pre\xE7o", subtitle: material.designacao, onClose, children: [
    isEquipamento ? /* @__PURE__ */ jsx(Field, { label: "Pre\xE7o (\u20AC)", children: /* @__PURE__ */ jsx("input", { value: preco, onChange: (e) => setPreco(e.target.value), type: "number", step: "0.01", min: "0", autoFocus: true, className: `${inputCls} font-mono-data`, placeholder: "0.00" }) }) : /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsx(Field, { label: rotuloPreco, children: /* @__PURE__ */ jsx("input", { value: preco, onChange: (e) => setPreco(e.target.value), type: "number", step: "0.01", min: "0", autoFocus: true, className: `${inputCls} font-mono-data`, placeholder: "0.00" }) }),
      /* @__PURE__ */ jsx(CustosExtraEditor, { custosExtra, onChange: setCustosExtra, tipos: tiposCustoExtra }),
      /* @__PURE__ */ jsx(DescontosEditor, { descontos, onChange: setDescontos, categorias: [...(tiposDesconto || []).map((t) => t.nome), "Outro"] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Data de Entrada em Vigor", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) }),
    !isEquipamento && /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between bg-amber-50 rounded-lg px-4 py-2.5 mb-4", children: [
      /* @__PURE__ */ jsx("span", { className: "text-sm font-semibold text-amber-800", children: "Pre\xE7o final (com descontos, transporte e otimiza\xE7\xE3o)" }),
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-amber-800", children: [
        precoFinal.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " \u20AC"
      ] })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar altera\xE7\xF5es" })
  ] });
}
function HistoricoPrecosModal({ material, isAdmin, onDelete, onEdit, onClose }) {
  const historico = historicoOrdenado(material);
  const emVigorId = precoVigente(material)?.id;
  const descontoTexto = (d) => `${d.tipo === "fixo" ? `${parseFloat(d.valor || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC` : `${d.valor || 0}%`} (${d.categoria === "Outro" && d.outroTexto ? d.outroTexto : d.categoria})${d.aplicarNoCalculo === false ? " [n\xE3o entra no c\xE1lculo]" : ""}`;
  const chartData = [...historico].sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || "")).map((p) => ({ data: formatDatePT(p.dataEntradaVigor), preco: Math.round(calcularPrecoFinal(p) * 100) / 100 }));
  const precos = chartData.map((d) => d.preco);
  const chartDomain = precos.length > 0 ? [Math.max(0.01, Math.floor(Math.min(...precos) * 0.9 * 100) / 100), Math.ceil(Math.max(...precos) * 1.1 * 100) / 100] : ["auto", "auto"];
  return /* @__PURE__ */ jsx(Modal, { title: "Hist\xF3rico de Pre\xE7os", subtitle: material.designacao, onClose, wide: true, children: historico.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: "Ainda n\xE3o h\xE1 pre\xE7os registados para este material." }) : /* @__PURE__ */ jsxs(Fragment, { children: [
    chartData.length >= 2 && /* @__PURE__ */ jsx("div", { className: "mb-5 -ml-2", children: /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 180, children: /* @__PURE__ */ jsxs(LineChart, { data: chartData, margin: { top: 8, right: 16, bottom: 0, left: 0 }, children: [
      /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", stroke: "#e7e5e4" }),
      /* @__PURE__ */ jsx(XAxis, { dataKey: "data", tick: { fontSize: 11, fill: "#78716c" }, tickLine: false, axisLine: { stroke: "#e7e5e4" } }),
      /* @__PURE__ */ jsx(
        YAxis,
        {
          domain: chartDomain,
          scale: "log",
          tick: { fontSize: 11, fill: "#78716c" },
          tickLine: false,
          axisLine: false,
          width: 50,
          tickFormatter: (v) => `${v.toLocaleString("pt-PT")} \u20AC`
        }
      ),
      /* @__PURE__ */ jsx(
        Tooltip,
        {
          formatter: (v) => [`${v.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC`, "Pre\xE7o final"],
          contentStyle: { fontSize: 12, borderRadius: 8, borderColor: "#e7e5e4" }
        }
      ),
      /* @__PURE__ */ jsx(Line, { type: "monotone", dataKey: "preco", stroke: "#d97706", strokeWidth: 2.5, dot: { r: 4, fill: "#d97706" }, activeDot: { r: 6 } })
    ] }) }) }),
    /* @__PURE__ */ jsx("div", { className: "space-y-2", children: historico.map((p) => /* @__PURE__ */ jsxs("div", { className: `border rounded-lg p-3 relative ${p.id === emVigorId ? "border-amber-300 bg-amber-50/60" : "border-stone-200"}`, children: [
      /* @__PURE__ */ jsx("div", { className: "absolute top-2 right-2 flex gap-1", children: isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
        /* @__PURE__ */ jsx("button", { onClick: () => onEdit(material, p), className: "text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
        /* @__PURE__ */ jsx("button", { onClick: () => onDelete(material.id, p.id), className: "text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
      ] }) }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 pr-12", children: [
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
          calcularPrecoFinal(p).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " \u20AC"
        ] }),
        p.id === emVigorId && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-amber-600 text-white", children: "Em vigor" })
      ] }),
      /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-500 mt-0.5", children: [
        "Desde ",
        formatDatePT(p.dataEntradaVigor),
        " \xB7 base ",
        parseFloat(p.preco || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " \u20AC",
        (p.descontos || []).map((d) => ` \u2212 ${descontoTexto(d)}`).join(""),
        custosExtraLista(p).map((c) => ` + ${c.nome.toLowerCase()} ${parseFloat(c.valor || 0).toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC`).join("")
      ] })
    ] }, p.id)) })
  ] }) });
}
function ImportMateriaisModal({ materiaisExistentes, tipo, onImport, onClose }) {
  const nomeSingular = tipo === "consumivel" ? "Combust\xEDvel" : tipo === "equipamento" ? "Equipamento" : tipo === "maodeobra" ? "M\xE3o de Obra" : "Material";
  const nomePlural = tipo === "consumivel" ? "Combust\xEDveis/Energia" : tipo === "equipamento" ? "Equipamentos" : tipo === "maodeobra" ? "M\xE3o de Obra" : "Materiais Constituintes";
  const [rows, setRows] = useState(null);
  const [grupos, setGrupos] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const asDate = (v) => {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v ?? "").trim();
  };
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json.length === 0) {
          setError("A folha est\xE1 vazia.");
          setRows(null);
          return;
        }
        const headerKeys = Object.keys(json[0]);
        const flat = (k) => normalizeHeader(k).replace(/[^a-z0-9]/g, "");
        const designacaoKey = headerKeys.find((k) => flat(k) === "designacao");
        const fornecedorKey = headerKeys.find((k) => flat(k) === "fornecedor");
        const precoKey = headerKeys.find((k) => flat(k) === "preco");
        const transporteKey = headerKeys.find((k) => flat(k).includes("transporte"));
        const dataKey = headerKeys.find((k) => flat(k) === "data" || flat(k).includes("entradaemvigor") || flat(k).includes("datavigor"));
        const descontoCols = [1, 2, 3].map((i) => ({
          valorKey: headerKeys.find((k) => flat(k) === `desconto${i}`),
          tipoKey: headerKeys.find((k) => flat(k) === `tipo${i}`),
          categoriaKey: headerKeys.find((k) => flat(k) === `categoria${i}`)
        }));
        if (!designacaoKey || !precoKey || !dataKey) {
          setError('N\xE3o encontrei as colunas "Designa\xE7\xE3o", "Pre\xE7o" e "Data" na primeira linha da folha.');
          setRows(null);
          return;
        }
        const interpretarTipo = (txt) => {
          const t = normalizeHeader(txt);
          return t.includes("fix") || t.includes("valor") || t.includes("\u20AC") ? "fixo" : "%";
        };
        const resolverCategoria = (txt) => {
          const t = String(txt ?? "").trim();
          if (!t) return { categoria: DESCONTO_CATEGORIAS[0], outroTexto: "" };
          const found = DESCONTO_CATEGORIAS.find((c) => normalizeHeader(c) === normalizeHeader(t));
          if (found && found !== "Outro") return { categoria: found, outroTexto: "" };
          return { categoria: "Outro", outroTexto: t };
        };
        const parsed = [];
        json.forEach((r) => {
          const designacao = String(r[designacaoKey] ?? "").trim();
          const precoTxt = String(r[precoKey] ?? "").trim();
          const dataEntradaVigor = asDate(r[dataKey]);
          if (!designacao || !precoTxt || !dataEntradaVigor) return;
          const descontos = [];
          descontoCols.forEach((col) => {
            if (!col.valorKey) return;
            const valorTxt = String(r[col.valorKey] ?? "").trim();
            if (!valorTxt) return;
            const valor = parseFloat(valorTxt) || 0;
            if (valor <= 0) return;
            const tipo2 = col.tipoKey ? interpretarTipo(r[col.tipoKey]) : "%";
            const { categoria, outroTexto } = resolverCategoria(col.categoriaKey ? r[col.categoriaKey] : "");
            descontos.push({ tipo: tipo2, valor, categoria, outroTexto });
          });
          parsed.push({
            designacao,
            fornecedor: fornecedorKey ? String(r[fornecedorKey] ?? "").trim() : "",
            preco: parseFloat(precoTxt) || 0,
            descontos,
            custoTransporte: transporteKey ? parseFloat(r[transporteKey]) || 0 : 0,
            dataEntradaVigor
          });
        });
        if (parsed.length === 0) {
          setError("N\xE3o encontrei linhas v\xE1lidas (Designa\xE7\xE3o, Pre\xE7o e Data preenchidos).");
          setRows(null);
          return;
        }
        const map = {};
        parsed.forEach((r) => {
          const chave = normalizeHeader(r.designacao);
          if (!map[chave]) map[chave] = { designacao: r.designacao, entradas: [] };
          map[chave].entradas.push(r);
        });
        const gruposPreview = Object.values(map).map((g) => {
          g.entradas.sort((a, b) => a.dataEntradaVigor.localeCompare(b.dataEntradaVigor));
          const existente = materiaisExistentes.find((m) => normalizeHeader(m.designacao) === normalizeHeader(g.designacao));
          return { ...g, existente: !!existente };
        });
        setRows(parsed);
        setGrupos(gruposPreview);
      } catch (err) {
        console.error(err);
        setError("N\xE3o foi poss\xEDvel ler este ficheiro. Confirme que \xE9 um .xlsx, .xls ou .csv v\xE1lido.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };
  const confirmImport = () => {
    setImporting(true);
    onImport(rows);
  };
  const baixarModelo = () => {
    const wsData = [
      ["Designa\xE7\xE3o", "Fornecedor", "Pre\xE7o", "Desconto 1", "Tipo 1", "Categoria 1", "Desconto 2", "Tipo 2", "Categoria 2", "Desconto 3", "Tipo 3", "Categoria 3", "Transporte", "Data"],
      ["Bitume 35/50", "Repsol", 450, 5, "%", "Desconto comercial", 2, "%", "B\xF3nus bom pagamento", "", "", "", 20, "2026-01-01"],
      ["Bitume 35/50", "Repsol", 470, 5, "%", "Desconto comercial", "", "", "", "", "", "", 20, "2026-04-01"],
      ["Filler Calc\xE1rio", "Fillercarb", 35, 3, "Fixo", "B\xF3nus quantidade", "", "", "", "", "", "", 8, "2026-01-15"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 22 }, { wch: 16 }, { wch: 9 }, { wch: 11 }, { wch: 8 }, { wch: 20 }, { wch: 11 }, { wch: 8 }, { wch: 20 }, { wch: 11 }, { wch: 8 }, { wch: 20 }, { wch: 11 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Materiais");
    XLSX.writeFile(wb, `modelo_importacao_${nomePlural.toLowerCase().replace(/\s+/g, "_")}.xlsx`);
  };
  const totalNovos = (grupos || []).filter((g) => !g.existente).length;
  const totalAtualizados = (grupos || []).filter((g) => g.existente).length;
  return /* @__PURE__ */ jsxs(Modal, { title: "Importar Lista de Pre\xE7os", subtitle: nomePlural, onClose, wide: true, children: [
    !rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm text-stone-500 mb-4", children: [
        "Escolha um ficheiro .xlsx, .xls ou .csv com as colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Designa\xE7\xE3o" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Pre\xE7o" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Data" }),
        " (data de entrada em vigor). As colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Fornecedor" }),
        " e ",
        /* @__PURE__ */ jsx("strong", { children: "Transporte" }),
        " s\xE3o opcionais. Para descontos (at\xE9 3), use pares de colunas ",
        /* @__PURE__ */ jsx("strong", { children: "Desconto 1" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Tipo 1" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Categoria 1" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Desconto 2" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Tipo 2" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Categoria 2" }),
        ", ",
        /* @__PURE__ */ jsx("strong", { children: "Desconto 3" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Tipo 3" }),
        "/",
        /* @__PURE__ */ jsx("strong", { children: "Categoria 3" }),
        ' \u2014 Tipo \xE9 "%" ou "Fixo"; Categoria pode ser Desconto comercial, B\xF3nus bom pagamento, B\xF3nus quantidade, ou outro texto livre. Linhas com a ',
        /* @__PURE__ */ jsx("strong", { children: "mesma designa\xE7\xE3o" }),
        " e datas diferentes s\xE3o tratadas como ",
        /* @__PURE__ */ jsx("strong", { children: "atualiza\xE7\xF5es do mesmo material" }),
        ", n\xE3o materiais separados."
      ] }),
      /* @__PURE__ */ jsxs("button", { onClick: baixarModelo, className: "w-full mb-3 flex items-center justify-center gap-1.5 py-2.5 rounded-lg border border-stone-300 text-stone-600 text-sm font-semibold hover:bg-stone-50", children: [
        /* @__PURE__ */ jsx(FileSpreadsheet, { size: 15 }),
        " Descarregar modelo (.xlsx)"
      ] }),
      /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", onChange: handleFile, className: "hidden" }),
      /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => fileInputRef.current?.click(),
          className: "w-full border-2 border-dashed border-stone-300 rounded-xl p-8 text-center hover:border-amber-400 hover:bg-amber-50/40 transition-colors",
          children: [
            /* @__PURE__ */ jsx(Upload, { className: "mx-auto text-stone-400 mb-2", size: 28 }),
            /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-stone-700", children: fileName || "Clique para escolher o ficheiro" })
          ]
        }
      ),
      error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mt-3", children: error })
    ] }),
    rows && /* @__PURE__ */ jsxs(Fragment, { children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 mb-3 text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 text-sm font-medium", children: [
        /* @__PURE__ */ jsx(CheckCircle2, { size: 16 }),
        " ",
        rows.length,
        " linha",
        rows.length !== 1 ? "s" : "",
        " \xB7 ",
        totalNovos,
        " material",
        totalNovos !== 1 ? "is" : "",
        " novo",
        totalNovos !== 1 ? "s" : "",
        " \xB7 ",
        totalAtualizados,
        " material",
        totalAtualizados !== 1 ? "is" : "",
        " atualizado",
        totalAtualizados !== 1 ? "s" : ""
      ] }),
      /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-72 overflow-y-auto mb-4", children: grupos.map((g, i) => /* @__PURE__ */ jsxs("div", { className: `px-3 py-2 text-sm ${i !== grupos.length - 1 ? "border-b border-stone-100" : ""}`, children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
          /* @__PURE__ */ jsx("span", { className: "font-medium text-stone-800", children: g.designacao }),
          /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full shrink-0 ${g.existente ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"}`, children: g.existente ? `+${g.entradas.length} atualiza\xE7\xE3o${g.entradas.length !== 1 ? "\xF5es" : ""}` : "Novo material" })
        ] }),
        /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 pl-0.5", children: g.entradas.map((e) => {
          const descTxt = (e.descontos || []).map((d) => ` \u2212 ${d.tipo === "fixo" ? `${d.valor}\u20AC` : `${d.valor}%`} (${d.categoria === "Outro" ? d.outroTexto : d.categoria})`).join("");
          return `${formatDatePT(e.dataEntradaVigor)}: ${e.preco.toLocaleString("pt-PT", { maximumFractionDigits: 2 })} \u20AC${descTxt}`;
        }).join(" \xB7 ") })
      ] }, i)) }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx("button", { onClick: () => {
          setRows(null);
          setGrupos(null);
          setFileName("");
        }, className: "flex-1 py-3 rounded-lg border border-stone-300 text-stone-600 font-display font-semibold tracking-wide uppercase text-sm hover:bg-stone-50", children: "Escolher outro" }),
        /* @__PURE__ */ jsx("button", { onClick: confirmImport, disabled: importing, className: "flex-1 py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700 disabled:opacity-60", children: importing ? "A importar..." : "Confirmar importa\xE7\xE3o" })
      ] })
    ] })
  ] });
}
function AtualizarProducaoModal({ center, onSave, onClose }) {
  const anoAtual = (/* @__PURE__ */ new Date()).getFullYear();
  const historico = center.parametrizacao?.producaoAnualHistorico || [];
  const [ano, setAno] = useState(anoAtual);
  const [valor, setValor] = useState("");
  const [error, setError] = useState("");
  const valorAtualDoAno = (anoRef) => {
    const doAno = historico.filter((h) => h.ano === anoRef);
    if (doAno.length === 0) return null;
    return [...doAno].sort((a, b) => (b.dataRegisto || "").localeCompare(a.dataRegisto || ""))[0];
  };
  const referencia = valorAtualDoAno(ano);
  const submit = () => {
    if (!ano || ano < 2e3 || ano > 2100) return setError("Indique um ano v\xE1lido");
    if (!valor || parseFloat(valor) <= 0) return setError("Indique a produ\xE7\xE3o estimada em toneladas");
    onSave(center.id, ano, valor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Produ\xE7\xE3o", subtitle: center.nome, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Ano", children: /* @__PURE__ */ jsx("input", { value: ano, onChange: (e) => setAno(parseInt(e.target.value, 10) || ""), type: "number", className: `${inputCls} font-mono-data`, autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Produ\xE7\xE3o estimada (toneladas)", children: /* @__PURE__ */ jsx("input", { value: valor, onChange: (e) => setValor(e.target.value), type: "number", step: "1", min: "0", className: `${inputCls} font-mono-data`, placeholder: "Ex: 45000" }) })
    ] }),
    referencia && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 -mt-2 mb-4", children: [
      "J\xE1 existe uma estimativa para ",
      ano,
      ": ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-600", children: [
        parseFloat(referencia.valor).toLocaleString("pt-PT"),
        " t"
      ] }),
      " \u2014 guardar aqui acrescenta uma nova atualiza\xE7\xE3o ao hist\xF3rico desse ano."
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EditarProducaoModal({ center, entry, onSave, onClose }) {
  const [ano, setAno] = useState(entry.ano);
  const [valor, setValor] = useState(String(entry.valor ?? ""));
  const [dataRegisto, setDataRegisto] = useState((entry.dataRegisto || (/* @__PURE__ */ new Date()).toISOString()).slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!ano || ano < 2e3 || ano > 2100) return setError("Indique um ano v\xE1lido");
    if (!valor || parseFloat(valor) <= 0) return setError("Indique a produ\xE7\xE3o estimada em toneladas");
    if (!dataRegisto) return setError("Indique a data de registo");
    onSave(center.id, entry.id, {
      ano: parseInt(ano, 10),
      valor: parseFloat(valor),
      dataRegisto: `${dataRegisto}T${(entry.dataRegisto || "").slice(11) || "00:00:00.000Z"}`
    });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Estimativa de Produ\xE7\xE3o", subtitle: center.nome, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Ano", children: /* @__PURE__ */ jsx("input", { value: ano, onChange: (e) => setAno(parseInt(e.target.value, 10) || ""), type: "number", className: `${inputCls} font-mono-data`, autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Produ\xE7\xE3o estimada (toneladas)", children: /* @__PURE__ */ jsx("input", { value: valor, onChange: (e) => setValor(e.target.value), type: "number", step: "1", min: "0", className: `${inputCls} font-mono-data` }) })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Data de Registo", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataRegisto, onChange: (e) => setDataRegisto(e.target.value), className: inputCls }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar altera\xE7\xF5es" })
  ] });
}
function HistoricoProducaoModal({ center, isAdmin, onEdit, onDelete, onClose }) {
  const historico = [...center.parametrizacao?.producaoAnualHistorico || []].sort((a, b) => (a.ano || 0) - (b.ano || 0) || (a.dataRegisto || "").localeCompare(b.dataRegisto || ""));
  const anoAtual = (/* @__PURE__ */ new Date()).getFullYear();
  const vigentePorAno = {};
  historico.forEach((h) => {
    const atual = vigentePorAno[h.ano];
    if (!atual || (h.dataRegisto || "") > (atual.dataRegisto || "")) vigentePorAno[h.ano] = h;
  });
  const chartData = Object.values(vigentePorAno).sort((a, b) => (a.ano || 0) - (b.ano || 0)).map((h) => ({ ano: String(h.ano), valor: Math.round((parseFloat(h.valor) || 0) * 100) / 100 }));
  const valores = chartData.map((d) => d.valor);
  const chartDomain = valores.length > 0 ? [Math.floor(Math.min(...valores) * 0.9), Math.ceil(Math.max(...valores) * 1.1)] : ["auto", "auto"];
  return /* @__PURE__ */ jsx(Modal, { title: "Hist\xF3rico de Previs\xE3o de Produ\xE7\xE3o", subtitle: center.nome, onClose, wide: true, children: historico.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: "Ainda n\xE3o h\xE1 estimativas de produ\xE7\xE3o registadas." }) : /* @__PURE__ */ jsxs(Fragment, { children: [
    chartData.length >= 2 && /* @__PURE__ */ jsx("div", { className: "mb-5 -ml-2", children: /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 180, children: /* @__PURE__ */ jsxs(LineChart, { data: chartData, margin: { top: 8, right: 16, bottom: 0, left: 0 }, children: [
      /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", stroke: "#e7e5e4" }),
      /* @__PURE__ */ jsx(XAxis, { dataKey: "ano", tick: { fontSize: 11, fill: "#78716c" }, tickLine: false, axisLine: { stroke: "#e7e5e4" } }),
      /* @__PURE__ */ jsx(
        YAxis,
        {
          domain: chartDomain,
          tick: { fontSize: 11, fill: "#78716c" },
          tickLine: false,
          axisLine: false,
          width: 55,
          tickFormatter: (v) => `${v.toLocaleString("pt-PT")} t`
        }
      ),
      /* @__PURE__ */ jsx(
        Tooltip,
        {
          formatter: (v) => [`${v.toLocaleString("pt-PT")} t`, "Produ\xE7\xE3o estimada"],
          contentStyle: { fontSize: 12, borderRadius: 8, borderColor: "#e7e5e4" }
        }
      ),
      /* @__PURE__ */ jsx(Line, { type: "monotone", dataKey: "valor", stroke: "#d97706", strokeWidth: 2.5, dot: { r: 4, fill: "#d97706" }, activeDot: { r: 6 } })
    ] }) }) }),
    /* @__PURE__ */ jsx("div", { className: "space-y-2", children: historico.map((h) => /* @__PURE__ */ jsxs("div", { className: "border border-stone-200 rounded-lg p-3 flex items-center justify-between", children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3", children: [
        /* @__PURE__ */ jsx("span", { className: "font-mono-data font-semibold text-stone-800", children: h.ano }),
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data text-stone-600", children: [
          parseFloat(h.valor).toLocaleString("pt-PT"),
          " t"
        ] }),
        h.id === vigentePorAno[h.ano]?.id && h.ano === anoAtual && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700", children: "Em vigor" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
        /* @__PURE__ */ jsxs("span", { className: "text-xs text-stone-400", children: [
          formatDateTimePT(h.dataRegisto),
          " \u2014 ",
          h.utilizador || "\u2014"
        ] }),
        isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
          /* @__PURE__ */ jsx("button", { onClick: () => onEdit(h), className: "p-1.5 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
          /* @__PURE__ */ jsx("button", { onClick: () => onDelete(center.id, h.id), className: "p-1.5 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
        ] })
      ] })
    ] }, h.id)) })
  ] }) });
}
function AtualizarTaxaModal({ center, blocoKey, titulo, unidade, onSave, onClose }) {
  const un = unidade || "L/t";
  const bloco = center.parametrizacao?.[blocoKey] || {};
  const historico = bloco.historico || [];
  const atual = [...historico].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""))[0] || null;
  const [valor, setValor] = useState("");
  const [dataEntradaVigor, setDataEntradaVigor] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!valor || parseFloat(valor) <= 0) return setError(`Indique o valor em ${un}`);
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    if (atual && parseFloat(valor) === parseFloat(atual.valor) && dataEntradaVigor === atual.dataEntradaVigor) {
      return setError("Sem altera\xE7\xF5es em rela\xE7\xE3o ao valor j\xE1 em vigor");
    }
    onSave(center.id, valor, dataEntradaVigor);
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Atualizar Valor", subtitle: `${titulo} \u2014 ${center.nome}`, onClose, children: [
    atual && /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-4 text-sm text-stone-600", children: [
      "Valor atual: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold", children: [
        parseFloat(atual.valor).toLocaleString("pt-PT"),
        " ",
        un
      ] }),
      atual.dataEntradaVigor && /* @__PURE__ */ jsxs("span", { className: "text-stone-400", children: [
        " (desde ",
        formatDatePT(atual.dataEntradaVigor),
        ")"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: `Valor (${un})`, children: /* @__PURE__ */ jsx("input", { value: valor, onChange: (e) => setValor(e.target.value), type: "number", step: "0.01", min: "0", className: `${inputCls} font-mono-data`, placeholder: "Ex: 6.5", autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Data de Entrada em Vigor", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar" })
  ] });
}
function EditarTaxaModal({ center, entry, unidade, onSave, onClose }) {
  const un = unidade || "L/t";
  const [valor, setValor] = useState(String(entry.valor ?? ""));
  const [dataEntradaVigor, setDataEntradaVigor] = useState(entry.dataEntradaVigor || (/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [error, setError] = useState("");
  const submit = () => {
    if (!valor || parseFloat(valor) <= 0) return setError(`Indique o valor em ${un}`);
    if (!dataEntradaVigor) return setError("Indique a data de entrada em vigor");
    onSave(center.id, entry.id, { valor: parseFloat(valor), dataEntradaVigor });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Editar Valor", subtitle: center.nome, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: `Valor (${un})`, children: /* @__PURE__ */ jsx("input", { value: valor, onChange: (e) => setValor(e.target.value), type: "number", step: "0.01", min: "0", className: `${inputCls} font-mono-data`, autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Data de Entrada em Vigor", children: /* @__PURE__ */ jsx("input", { type: "date", value: dataEntradaVigor, onChange: (e) => setDataEntradaVigor(e.target.value), className: inputCls }) })
    ] }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar altera\xE7\xF5es" })
  ] });
}
function HistoricoTaxaModal({ center, blocoKey, titulo, unidade, isAdmin, onEdit, onDelete, onClose }) {
  const un = unidade || "L/t";
  const bloco = center.parametrizacao?.[blocoKey] || {};
  const historico = [...bloco.historico || []].sort((a, b) => (b.dataEntradaVigor || "").localeCompare(a.dataEntradaVigor || ""));
  const vigenteId = historico[0]?.id;
  const chartData = [...bloco.historico || []].sort((a, b) => (a.dataEntradaVigor || "").localeCompare(b.dataEntradaVigor || "")).map((h) => ({ data: formatDatePT(h.dataEntradaVigor), valor: Math.round((parseFloat(h.valor) || 0) * 100) / 100 }));
  const valores = chartData.map((d) => d.valor);
  const chartDomain = valores.length > 0 ? [Math.floor(Math.min(...valores) * 0.9 * 100) / 100, Math.ceil(Math.max(...valores) * 1.1 * 100) / 100] : ["auto", "auto"];
  return /* @__PURE__ */ jsx(Modal, { title: "Hist\xF3rico de Valor", subtitle: `${titulo} \u2014 ${center.nome}`, onClose, wide: true, children: historico.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500", children: "Ainda n\xE3o h\xE1 valores registados." }) : /* @__PURE__ */ jsxs(Fragment, { children: [
    chartData.length >= 2 && /* @__PURE__ */ jsx("div", { className: "mb-5 -ml-2", children: /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 180, children: /* @__PURE__ */ jsxs(LineChart, { data: chartData, margin: { top: 8, right: 16, bottom: 0, left: 0 }, children: [
      /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", stroke: "#e7e5e4" }),
      /* @__PURE__ */ jsx(XAxis, { dataKey: "data", tick: { fontSize: 11, fill: "#78716c" }, tickLine: false, axisLine: { stroke: "#e7e5e4" } }),
      /* @__PURE__ */ jsx(
        YAxis,
        {
          domain: chartDomain,
          tick: { fontSize: 11, fill: "#78716c" },
          tickLine: false,
          axisLine: false,
          width: 55,
          tickFormatter: (v) => `${v.toLocaleString("pt-PT")} ${un}`
        }
      ),
      /* @__PURE__ */ jsx(
        Tooltip,
        {
          formatter: (v) => [`${v.toLocaleString("pt-PT")} ${un}`, "Valor"],
          contentStyle: { fontSize: 12, borderRadius: 8, borderColor: "#e7e5e4" }
        }
      ),
      /* @__PURE__ */ jsx(Line, { type: "monotone", dataKey: "valor", stroke: "#d97706", strokeWidth: 2.5, dot: { r: 4, fill: "#d97706" }, activeDot: { r: 6 } })
    ] }) }) }),
    /* @__PURE__ */ jsx("div", { className: "space-y-2", children: historico.map((h) => /* @__PURE__ */ jsxs("div", { className: `border rounded-lg p-3 flex items-center justify-between ${h.id === vigenteId ? "border-amber-300 bg-amber-50/60" : "border-stone-200"}`, children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-3", children: [
        /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
          parseFloat(h.valor).toLocaleString("pt-PT"),
          " ",
          un
        ] }),
        /* @__PURE__ */ jsxs("span", { className: "text-stone-400 text-sm", children: [
          "desde ",
          formatDatePT(h.dataEntradaVigor)
        ] }),
        h.id === vigenteId && /* @__PURE__ */ jsx("span", { className: "text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-amber-600 text-white", children: "Em vigor" })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
        /* @__PURE__ */ jsxs("span", { className: "text-xs text-stone-400", children: [
          formatDateTimePT(h.dataRegisto),
          " \u2014 ",
          h.utilizador || "\u2014"
        ] }),
        isAdmin && /* @__PURE__ */ jsxs(Fragment, { children: [
          /* @__PURE__ */ jsx("button", { onClick: () => onEdit(h), className: "p-1.5 text-stone-400 hover:text-amber-600", children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
          /* @__PURE__ */ jsx("button", { onClick: () => onDelete(center.id, h.id), className: "p-1.5 text-stone-400 hover:text-red-600", children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
        ] })
      ] })
    ] }, h.id)) })
  ] }) });
}
function HistoricoStockModal({ center, produtoId, categoria, materiais, consumiveis, rececoes, diarias, formulas, ajustesStock, isAdmin, onAddAjuste, onClose }) {
  const lista = categoria === "consumiveis" ? consumiveis : materiais;
  const produto = lista.find((p) => p.id === produtoId);
  const un = categoria === "consumiveis" ? "L" : "t";
  const movimentos = calcularMovimentosStock({ center, produtoId, categoria, rececoes, diarias, formulas, ajustesStock });
  const movimentosDesc = [...movimentos].reverse();
  const stockAtual = movimentos.length > 0 ? movimentos[movimentos.length - 1].saldo : 0;
  const chartData = movimentos.map((m) => ({ data: formatDatePT(m.data), saldo: Math.round(m.saldo * 100) / 100 }));
  const saldos = chartData.map((d) => d.saldo);
  const chartDomain = saldos.length > 0 ? [Math.floor(Math.min(0, ...saldos) * 1.1), Math.ceil(Math.max(...saldos, 0) * 1.1 || 1)] : ["auto", "auto"];
  const tipoCor = { "Rece\xE7\xE3o": "text-emerald-600", "Rece\xE7\xE3o (substituto)": "text-emerald-600", "Ajuste": "text-amber-600", "Consumo (Produ\xE7\xE3o)": "text-stone-500", "Consumo (Bloco T\xE9rmico)": "text-stone-500" };
  return /* @__PURE__ */ jsxs(Modal, { title: "Hist\xF3rico de Stock", subtitle: `${produto?.designacao || "\u2014"} \u2014 ${center.nome}`, onClose, wide: true, children: [
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-4 text-sm text-stone-600 flex items-center justify-between", children: [
      /* @__PURE__ */ jsx("span", { children: "Stock atual" }),
      /* @__PURE__ */ jsxs("span", { className: `font-mono-data font-semibold text-lg ${stockAtual < 0 ? "text-red-600" : "text-stone-800"}`, children: [
        stockAtual.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " ",
        un
      ] })
    ] }),
    chartData.length >= 2 && /* @__PURE__ */ jsx("div", { className: "mb-5 -ml-2", children: /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 180, children: /* @__PURE__ */ jsxs(LineChart, { data: chartData, margin: { top: 8, right: 16, bottom: 0, left: 0 }, children: [
      /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", stroke: "#e7e5e4" }),
      /* @__PURE__ */ jsx(XAxis, { dataKey: "data", tick: { fontSize: 11, fill: "#78716c" }, tickLine: false, axisLine: { stroke: "#e7e5e4" } }),
      /* @__PURE__ */ jsx(
        YAxis,
        {
          domain: chartDomain,
          tick: { fontSize: 11, fill: "#78716c" },
          tickLine: false,
          axisLine: false,
          width: 55,
          tickFormatter: (v) => `${v.toLocaleString("pt-PT")} ${un}`
        }
      ),
      /* @__PURE__ */ jsx(
        Tooltip,
        {
          formatter: (v) => [`${v.toLocaleString("pt-PT")} ${un}`, "Stock"],
          contentStyle: { fontSize: 12, borderRadius: 8, borderColor: "#e7e5e4" }
        }
      ),
      /* @__PURE__ */ jsx(Line, { type: "monotone", dataKey: "saldo", stroke: "#d97706", strokeWidth: 2.5, dot: { r: 3, fill: "#d97706" }, activeDot: { r: 6 } })
    ] }) }) }),
    isAdmin && /* @__PURE__ */ jsxs("button", { onClick: () => onAddAjuste(movimentos, un), className: "w-full mb-4 py-2.5 rounded-lg border border-dashed border-amber-300 text-amber-700 text-sm font-semibold hover:bg-amber-50 flex items-center justify-center gap-1.5", children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Novo Ajuste"
    ] }),
    /* @__PURE__ */ jsx("div", { className: "border border-stone-200 rounded-lg overflow-hidden max-h-72 overflow-y-auto", children: movimentosDesc.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-sm text-stone-500 px-3 py-4 text-center", children: "Ainda n\xE3o h\xE1 movimentos registados." }) : movimentosDesc.map((m, i) => /* @__PURE__ */ jsxs("div", { className: `flex items-center justify-between px-3 py-2.5 text-sm ${i !== movimentosDesc.length - 1 ? "border-b border-stone-100" : ""}`, children: [
      /* @__PURE__ */ jsxs("div", { className: "min-w-0", children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
          /* @__PURE__ */ jsx("span", { className: "text-stone-700 font-medium", children: formatDatePT(m.data) }),
          /* @__PURE__ */ jsx("span", { className: `text-[10px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full bg-stone-100 ${tipoCor[m.tipo] || "text-stone-500"}`, children: m.tipo })
        ] }),
        m.tipo === "Ajuste" && m.motivo && /* @__PURE__ */ jsx("p", { className: "text-xs text-stone-400 mt-0.5", children: m.motivo }),
        m.tipo === "Ajuste" && /* @__PURE__ */ jsxs("p", { className: "text-[11px] text-stone-300 mt-0.5", children: [
          formatDateTimePT(m.dataRegisto),
          " \u2014 ",
          m.utilizador || "\u2014"
        ] }),
        m.tipo === "Rece\xE7\xE3o (substituto)" && m.produtoRecebidoId && /* @__PURE__ */ jsxs("p", { className: "text-xs text-stone-400 mt-0.5", children: [
          "recebido: ",
          (categoria === "consumiveis" ? consumiveis : materiais).find((p) => p.id === m.produtoRecebidoId)?.designacao || "\u2014"
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "text-right shrink-0", children: [
        /* @__PURE__ */ jsxs("span", { className: `font-mono-data font-semibold ${m.quantidade < 0 ? "text-red-600" : "text-emerald-600"}`, children: [
          m.quantidade >= 0 ? "+" : "",
          m.quantidade.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " ",
          un
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "text-[11px] text-stone-400", children: [
          "saldo ",
          m.saldo.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
          " ",
          un
        ] })
      ] })
    ] }, m.id)) })
  ] });
}
function NovoAjusteModal({ centroId, produtoId, categoria, movimentos, unidade, onSave, onClose }) {
  const [data, setData] = useState((/* @__PURE__ */ new Date()).toISOString().slice(0, 10));
  const [stockReal, setStockReal] = useState("");
  const [motivo, setMotivo] = useState("");
  const [error, setError] = useState("");
  const un = unidade || (categoria === "consumiveis" ? "L" : "t");
  const saldoCalculado = (() => {
    const aplicaveis = (movimentos || []).filter((m) => m.data && m.data <= data);
    if (aplicaveis.length === 0) return 0;
    return aplicaveis[aplicaveis.length - 1].saldo;
  })();
  const diferenca = stockReal !== "" && !isNaN(parseFloat(stockReal)) ? parseFloat(stockReal) - saldoCalculado : null;
  const submit = () => {
    if (!data) return setError("Indique a data");
    if (stockReal === "" || isNaN(parseFloat(stockReal))) return setError("Indique o stock real contado nesta data");
    if (diferenca === 0) return setError("O stock real j\xE1 corresponde ao stock calculado \u2014 n\xE3o h\xE1 nenhum ajuste a fazer");
    if (!motivo.trim()) return setError("Descreva o motivo do ajuste (ex: contagem f\xEDsica)");
    onSave({ centroId, produtoId, categoria, data, quantidade: diferenca, motivo: motivo.trim() });
  };
  return /* @__PURE__ */ jsxs(Modal, { title: "Novo Ajuste de Stock", subtitle: "Para corrigir desvios face \xE0 contagem real", onClose, children: [
    /* @__PURE__ */ jsx(Field, { label: "Data da contagem", children: /* @__PURE__ */ jsx("input", { type: "date", value: data, onChange: (e) => setData(e.target.value), className: inputCls, autoFocus: true }) }),
    /* @__PURE__ */ jsxs("div", { className: "bg-stone-100 rounded-lg px-3 py-2.5 mb-3 text-sm text-stone-600", children: [
      "Stock calculado nesta data: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold text-stone-800", children: [
        saldoCalculado.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " ",
        un
      ] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: `Stock real contado (${un})`, children: /* @__PURE__ */ jsx("input", { value: stockReal, onChange: (e) => setStockReal(e.target.value), type: "number", step: "0.01", className: `${inputCls} font-mono-data`, placeholder: "Ex: 437.5" }) }),
    diferenca !== null && /* @__PURE__ */ jsxs("div", { className: `rounded-lg px-3 py-2.5 mb-4 text-sm ${diferenca === 0 ? "bg-stone-100 text-stone-500" : diferenca > 0 ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`, children: [
      "Ajuste a registar: ",
      /* @__PURE__ */ jsxs("span", { className: "font-mono-data font-semibold", children: [
        diferenca >= 0 ? "+" : "",
        diferenca.toLocaleString("pt-PT", { maximumFractionDigits: 2 }),
        " ",
        un
      ] })
    ] }),
    /* @__PURE__ */ jsx(Field, { label: "Motivo", children: /* @__PURE__ */ jsx("textarea", { value: motivo, onChange: (e) => setMotivo(e.target.value), className: inputCls, rows: 3, spellCheck: "true", lang: "pt-PT", placeholder: "Ex: contagem f\xEDsica de fim de m\xEAs..." }) }),
    error && /* @__PURE__ */ jsx("p", { className: "text-red-600 text-sm mb-3", children: error }),
    /* @__PURE__ */ jsx("button", { onClick: submit, className: "w-full py-3 rounded-lg bg-amber-600 text-white font-display font-semibold tracking-wide uppercase text-sm hover:bg-amber-700", children: "Guardar Ajuste" })
  ] });
}
export {
  App as default
};
