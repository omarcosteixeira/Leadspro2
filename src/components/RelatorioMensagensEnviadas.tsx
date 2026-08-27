import React, { useState, useMemo } from "react";
import { MensagemEnviadaLog } from "../types";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
  AreaChart,
  Area,
} from "recharts";
import {
  MessageSquare,
  Mail,
  Bot,
  Calendar,
  Filter,
  Users,
  Search,
  Download,
  Clock,
  TrendingUp,
  Database,
  GraduationCap,
  Sparkles,
  Send,
  X,
  History as HistoryIcon,
} from "lucide-react";
import { cn } from "../lib/utils";
import * as XLSX from "xlsx";

const PIE_COLORS = ["#10b981", "#8b5cf6", "#0ea5e9", "#f59e0b", "#ec4899", "#3b82f6", "#14b8a6"];

interface RelatorioMensagensEnviadasProps {
  mensagensEnviadasLog: MensagemEnviadaLog[];
}

export function RelatorioMensagensEnviadas({
  mensagensEnviadasLog = [],
}: RelatorioMensagensEnviadasProps) {
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [baseFilter, setBaseFilter] = useState("");
  const [tipoEnvioFilter, setTipoEnvioFilter] = useState<string>("");
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedCandidateHistory, setSelectedCandidateHistory] = useState<{
    nome: string;
    logs: MensagemEnviadaLog[];
  } | null>(null);

  // Extract unique filter lists
  const uniqueCursos = useMemo(() => {
    return Array.from(
      new Set(
        mensagensEnviadasLog
          .map((m) => m.curso?.trim())
          .filter(Boolean)
      )
    ).sort();
  }, [mensagensEnviadasLog]);

  const uniqueBases = useMemo(() => {
    return Array.from(
      new Set(
        mensagensEnviadasLog
          .map((m) => m.base?.trim())
          .filter(Boolean)
      )
    ).sort();
  }, [mensagensEnviadasLog]);

  // Helper to parse date from firestore timestamp or string
  const parseDate = (val: any): Date | null => {
    if (!val) return null;
    if (val.toDate && typeof val.toDate === "function") return val.toDate();
    if (val.seconds) return new Date(val.seconds * 1000);
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  };

  // Filtered dataset
  const filteredData = useMemo(() => {
    return mensagensEnviadasLog
      .filter((log) => {
        const d = parseDate(log.dataHora);
        if (startDate) {
          if (!d) return false;
          const start = new Date(startDate + "T00:00:00");
          if (d < start) return false;
        }
        if (endDate) {
          if (!d) return false;
          const end = new Date(endDate + "T23:59:59");
          if (d > end) return false;
        }
        if (cursoFilter && log.curso !== cursoFilter) return false;
        if (baseFilter && log.base !== baseFilter) return false;
        if (tipoEnvioFilter && log.tipoEnvio !== tipoEnvioFilter) return false;
        if (searchTerm) {
          const s = searchTerm.toLowerCase();
          const matchNome = log.nome?.toLowerCase().includes(s);
          const matchTel = log.telefone?.includes(s);
          const matchCurso = log.curso?.toLowerCase().includes(s);
          const matchBase = log.base?.toLowerCase().includes(s);
          const matchUser = log.usuarioNome?.toLowerCase().includes(s);
          if (!matchNome && !matchTel && !matchCurso && !matchBase && !matchUser) return false;
        }
        return true;
      })
      .sort((a, b) => {
        const dateA = parseDate(a.dataHora)?.getTime() || 0;
        const dateB = parseDate(b.dataHora)?.getTime() || 0;
        return dateB - dateA;
      });
  }, [mensagensEnviadasLog, startDate, endDate, cursoFilter, baseFilter, tipoEnvioFilter, searchTerm]);

  // Candidate contact frequency map (how many times each candidate received messages in filtered data)
  const candidateFrequencyMap = useMemo(() => {
    const map = new Map<string, number>();
    filteredData.forEach((item) => {
      const key = item.nome?.toLowerCase().trim() || item.telefone || "desconhecido";
      map.set(key, (map.get(key) || 0) + 1);
    });
    return map;
  }, [filteredData]);

  // General counts
  const totalEnviadas = filteredData.length;
  const totalWhats = filteredData.filter(
    (d) => d.tipoEnvio === "whats"
  ).length;
  const totalMalaDireta = filteredData.filter(
    (d) => d.tipoEnvio === "maladireta"
  ).length;
  const totalBot = filteredData.filter(
    (d) => d.tipoEnvio === "bot_automatico"
  ).length;

  const uniqueCandidatesCount = useMemo(() => {
    const set = new Set<string>();
    filteredData.forEach((item) => {
      const key = item.nome?.toLowerCase().trim() || item.telefone;
      if (key) set.add(key);
    });
    return set.size;
  }, [filteredData]);

  // Course distribution
  const byCourseData = useMemo(() => {
    const acc: Record<string, { total: number; whats: number; mala: number; bot: number }> = {};
    filteredData.forEach((item) => {
      const curso = item.curso || "Não Informado";
      if (!acc[curso]) acc[curso] = { total: 0, whats: 0, mala: 0, bot: 0 };
      acc[curso].total += 1;
      if (item.tipoEnvio === "whats") acc[curso].whats += 1;
      else if (item.tipoEnvio === "maladireta") acc[curso].mala += 1;
      else if (item.tipoEnvio === "bot_automatico") acc[curso].bot += 1;
    });

    return Object.entries(acc)
      .map(([name, data]) => ({
        name,
        total: data.total,
        WhatsApp: data.whats,
        "Mala Direta": data.mala,
        "Bot Auto": data.bot,
      }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);
  }, [filteredData]);

  // Base distribution
  const byBaseData = useMemo(() => {
    const acc: Record<string, { total: number; whats: number; mala: number; bot: number }> = {};
    filteredData.forEach((item) => {
      const base = item.base || "Base Geral";
      if (!acc[base]) acc[base] = { total: 0, whats: 0, mala: 0, bot: 0 };
      acc[base].total += 1;
      if (item.tipoEnvio === "whats") acc[base].whats += 1;
      else if (item.tipoEnvio === "maladireta") acc[base].mala += 1;
      else if (item.tipoEnvio === "bot_automatico") acc[base].bot += 1;
    });

    return Object.entries(acc)
      .map(([name, data]) => ({
        name,
        total: data.total,
        WhatsApp: data.whats,
        "Mala Direta": data.mala,
        "Bot Auto": data.bot,
      }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);
  }, [filteredData]);

  // By Type Distribution for Pie Chart
  const byTypeData = useMemo(() => {
    return [
      { name: "WhatsApp", value: totalWhats, color: "#10b981" },
      { name: "Mala Direta", value: totalMalaDireta, color: "#8b5cf6" },
      { name: "Bot Automático", value: totalBot, color: "#0ea5e9" },
    ].filter((item) => item.value > 0);
  }, [totalWhats, totalMalaDireta, totalBot]);

  // By Day / Timeline
  const timelineData = useMemo(() => {
    const acc: Record<string, { date: string; dateObj: Date; whats: number; mala: number; bot: number; total: number }> = {};
    filteredData.forEach((item) => {
      const d = parseDate(item.dataHora);
      if (!d) return;
      const key = d.toLocaleDateString("pt-BR");
      if (!acc[key]) {
        acc[key] = {
          date: key,
          dateObj: new Date(d.getFullYear(), d.getMonth(), d.getDate()),
          whats: 0,
          mala: 0,
          bot: 0,
          total: 0,
        };
      }
      acc[key].total += 1;
      if (item.tipoEnvio === "whats") acc[key].whats += 1;
      else if (item.tipoEnvio === "maladireta") acc[key].mala += 1;
      else if (item.tipoEnvio === "bot_automatico") acc[key].bot += 1;
    });

    return Object.values(acc)
      .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime())
      .slice(-15);
  }, [filteredData]);

  // Export to Excel
  const handleExportExcel = () => {
    const exportRows = filteredData.map((item, idx) => {
      const d = parseDate(item.dataHora);
      const dataStr = d ? d.toLocaleDateString("pt-BR") : "N/A";
      const horaStr = d ? d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }) : "N/A";
      const freqKey = item.nome?.toLowerCase().trim() || item.telefone || "";
      const totalVezes = candidateFrequencyMap.get(freqKey) || 1;

      return {
        "#": idx + 1,
        "Nome do Candidato": item.nome || "Não Informado",
        Telefone: item.telefone || "",
        Curso: item.curso || "Não Informado",
        Base: item.base || "Base Geral",
        "Tipo de Envio":
          item.tipoEnvio === "whats"
            ? "WhatsApp"
            : item.tipoEnvio === "maladireta"
            ? "Mala Direta"
            : "Bot Automático",
        Data: dataStr,
        Hora: horaStr,
        "Total de Envios para o Candidato": totalVezes,
        "Usuário / Operador": item.usuarioNome || "Sistema",
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Mensagens Enviadas");
    XLSX.writeFile(workbook, `Relatorio_Mensagens_Enviadas_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  // Quick Preset Filters
  const setQuickDate = (preset: "today" | "7days" | "month" | "all") => {
    const today = new Date();
    const todayStr = today.toISOString().split("T")[0];
    if (preset === "today") {
      setStartDate(todayStr);
      setEndDate(todayStr);
    } else if (preset === "7days") {
      const past = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
      setStartDate(past.toISOString().split("T")[0]);
      setEndDate(todayStr);
    } else if (preset === "month") {
      const past = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
      setStartDate(past.toISOString().split("T")[0]);
      setEndDate(todayStr);
    } else {
      setStartDate("");
      setEndDate("");
    }
  };

  // Open Candidate Detailed History Modal
  const openCandidateHistory = (candidateName: string) => {
    const norm = candidateName.toLowerCase().trim();
    const history = mensagensEnviadasLog
      .filter((m) => m.nome?.toLowerCase().trim() === norm)
      .sort((a, b) => {
        const dA = parseDate(a.dataHora)?.getTime() || 0;
        const dB = parseDate(b.dataHora)?.getTime() || 0;
        return dB - dA;
      });
    setSelectedCandidateHistory({ nome: candidateName, logs: history });
  };

  return (
    <div className="space-y-8">
      {/* Header & Controls Section */}
      <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-6">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 border-b border-slate-100 pb-4">
          <div>
            <div className="flex items-center gap-2">
              <div className="p-2 bg-emerald-50 text-emerald-600 rounded-xl">
                <MessageSquare size={20} />
              </div>
              <h3 className="text-xl font-bold text-slate-800">
                Relatório de Mensagens Enviadas (Whats / Mala Direta / Bot)
              </h3>
            </div>
            <p className="text-slate-500 text-xs mt-1">
              Acompanhe quantas vezes, os dias, as horas e os canais utilizados nos envios por candidato, base e curso.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-4 py-2 rounded-xl text-xs transition-all shadow-sm"
              title="Exportar dados filtrados para Excel"
            >
              <Download size={15} />
              <span>Exportar Excel</span>
            </button>
          </div>
        </div>

        {/* Filters Panel */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          {/* Data Inicio */}
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
              Data Início
            </label>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all"
            />
          </div>

          {/* Data Fim */}
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
              Data Fim
            </label>
            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all"
            />
          </div>

          {/* Filtro por Curso */}
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
              Filtrar por Curso
            </label>
            <select
              value={cursoFilter}
              onChange={(e) => setCursoFilter(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all truncate"
            >
              <option value="">Todos os Cursos ({uniqueCursos.length})</option>
              {uniqueCursos.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </div>

          {/* Filtro por Base */}
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
              Filtrar por Base
            </label>
            <select
              value={baseFilter}
              onChange={(e) => setBaseFilter(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all truncate"
            >
              <option value="">Todas as Bases ({uniqueBases.length})</option>
              {uniqueBases.map((b) => (
                <option key={b} value={b}>
                  {b}
                </option>
              ))}
            </select>
          </div>

          {/* Filtro por Tipo de Envio */}
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
              Tipo de Envio
            </label>
            <select
              value={tipoEnvioFilter}
              onChange={(e) => setTipoEnvioFilter(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-medium"
            >
              <option value="">Todos os Canais</option>
              <option value="whats">💬 WhatsApp Manual</option>
              <option value="maladireta">✉️ Mala Direta</option>
              <option value="bot_automatico">🤖 Bot Automático ARGO'S</option>
            </select>
          </div>
        </div>

        {/* Quick Date Presets & Search */}
        <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-2">
          <div className="flex flex-wrap items-center gap-1.5 w-full sm:w-auto">
            <span className="text-[11px] font-bold text-slate-400 mr-1 flex items-center gap-1">
              <Calendar size={13} /> Período:
            </span>
            <button
              onClick={() => setQuickDate("today")}
              className={cn(
                "px-2.5 py-1 rounded-lg text-xs font-semibold transition-all",
                startDate === new Date().toISOString().split("T")[0] &&
                  endDate === new Date().toISOString().split("T")[0]
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "bg-slate-100 hover:bg-slate-200 text-slate-600"
              )}
            >
              Hoje
            </button>
            <button
              onClick={() => setQuickDate("7days")}
              className="px-2.5 py-1 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-600 transition-all"
            >
              Últimos 7 dias
            </button>
            <button
              onClick={() => setQuickDate("month")}
              className="px-2.5 py-1 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-600 transition-all"
            >
              Últimos 30 dias
            </button>
            <button
              onClick={() => setQuickDate("all")}
              className="px-2.5 py-1 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-600 transition-all"
            >
              Todo o Período
            </button>
            {(startDate || endDate || cursoFilter || baseFilter || tipoEnvioFilter || searchTerm) && (
              <button
                onClick={() => {
                  setStartDate("");
                  setEndDate("");
                  setCursoFilter("");
                  setBaseFilter("");
                  setTipoEnvioFilter("");
                  setSearchTerm("");
                }}
                className="px-2.5 py-1 rounded-lg text-xs font-bold text-rose-600 bg-rose-50 hover:bg-rose-100 transition-all ml-1"
              >
                Limpar Filtros
              </button>
            )}
          </div>

          <div className="relative w-full sm:w-72">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={15} />
            <input
              type="text"
              placeholder="Buscar por candidato, telefone..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-4 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-emerald-500 transition-all"
            />
          </div>
        </div>
      </div>

      {/* KPI Cards Dashboard */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
        {/* Total Geral */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
          <div className="p-3.5 rounded-2xl bg-slate-900 text-white">
            <Send size={22} />
          </div>
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Envios</p>
            <h3 className="text-2xl font-black text-slate-900 mt-0.5">{totalEnviadas}</h3>
            <span className="text-[11px] font-semibold text-slate-500">mensagens registradas</span>
          </div>
        </div>

        {/* Total WhatsApp Manual */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
          <div className="p-3.5 rounded-2xl bg-emerald-500 text-white">
            <MessageSquare size={22} />
          </div>
          <div>
            <p className="text-xs font-bold text-emerald-600 uppercase tracking-wider">WhatsApp</p>
            <h3 className="text-2xl font-black text-slate-900 mt-0.5">{totalWhats}</h3>
            <span className="text-[11px] font-semibold text-slate-500">
              {totalEnviadas > 0 ? ((totalWhats / totalEnviadas) * 100).toFixed(1) : 0}% do total
            </span>
          </div>
        </div>

        {/* Total Mala Direta */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
          <div className="p-3.5 rounded-2xl bg-purple-500 text-white">
            <Mail size={22} />
          </div>
          <div>
            <p className="text-xs font-bold text-purple-600 uppercase tracking-wider">Mala Direta</p>
            <h3 className="text-2xl font-black text-slate-900 mt-0.5">{totalMalaDireta}</h3>
            <span className="text-[11px] font-semibold text-slate-500">
              {totalEnviadas > 0 ? ((totalMalaDireta / totalEnviadas) * 100).toFixed(1) : 0}% do total
            </span>
          </div>
        </div>

        {/* Total Bot Automático */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
          <div className="p-3.5 rounded-2xl bg-sky-500 text-white">
            <Bot size={22} />
          </div>
          <div>
            <p className="text-xs font-bold text-sky-600 uppercase tracking-wider">Bot Automático</p>
            <h3 className="text-2xl font-black text-slate-900 mt-0.5">{totalBot}</h3>
            <span className="text-[11px] font-semibold text-slate-500">
              {totalEnviadas > 0 ? ((totalBot / totalEnviadas) * 100).toFixed(1) : 0}% do total
            </span>
          </div>
        </div>

        {/* Candidatos Únicos */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
          <div className="p-3.5 rounded-2xl bg-amber-500 text-white">
            <Users size={22} />
          </div>
          <div>
            <p className="text-xs font-bold text-amber-600 uppercase tracking-wider">Candidatos</p>
            <h3 className="text-2xl font-black text-slate-900 mt-0.5">{uniqueCandidatesCount}</h3>
            <span className="text-[11px] font-semibold text-slate-500">
              {uniqueCandidatesCount > 0 ? (totalEnviadas / uniqueCandidatesCount).toFixed(1) : 0} envios/cand.
            </span>
          </div>
        </div>
      </div>

      {/* Visual Charts Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Total por Curso */}
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <GraduationCap className="text-emerald-600" size={18} />
              <h4 className="text-sm font-bold text-slate-800">Total por Curso (Top 10)</h4>
            </div>
            <span className="text-xs font-bold text-slate-400">{byCourseData.length} cursos</span>
          </div>

          <div className="h-[260px] w-full">
            {byCourseData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={byCourseData} layout="vertical" margin={{ left: 10, right: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                  <XAxis type="number" tick={{ fontSize: 10, fill: "#64748b" }} />
                  <YAxis
                    dataKey="name"
                    type="category"
                    width={90}
                    tick={{ fontSize: 10, fill: "#334155", fontWeight: 600 }}
                    tickFormatter={(val) => (val.length > 14 ? `${val.substring(0, 12)}...` : val)}
                  />
                  <Tooltip
                    contentStyle={{
                      borderRadius: "12px",
                      border: "none",
                      boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                    }}
                  />
                  <Legend wrapperStyle={{ fontSize: "10px", fontWeight: "bold" }} />
                  <Bar dataKey="WhatsApp" stackId="a" fill="#10b981" radius={[0, 0, 0, 0]} />
                  <Bar dataKey="Mala Direta" stackId="a" fill="#8b5cf6" radius={[0, 0, 0, 0]} />
                  <Bar dataKey="Bot Auto" stackId="a" fill="#0ea5e9" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-slate-400 text-xs italic">
                Nenhum dado com os filtros atuais.
              </div>
            )}
          </div>
        </div>

        {/* Total por Base */}
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Database className="text-purple-600" size={18} />
              <h4 className="text-sm font-bold text-slate-800">Total por Base (Top 10)</h4>
            </div>
            <span className="text-xs font-bold text-slate-400">{byBaseData.length} bases</span>
          </div>

          <div className="h-[260px] w-full">
            {byBaseData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={byBaseData} layout="vertical" margin={{ left: 10, right: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                  <XAxis type="number" tick={{ fontSize: 10, fill: "#64748b" }} />
                  <YAxis
                    dataKey="name"
                    type="category"
                    width={90}
                    tick={{ fontSize: 10, fill: "#334155", fontWeight: 600 }}
                    tickFormatter={(val) => (val.length > 14 ? `${val.substring(0, 12)}...` : val)}
                  />
                  <Tooltip
                    contentStyle={{
                      borderRadius: "12px",
                      border: "none",
                      boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                    }}
                  />
                  <Legend wrapperStyle={{ fontSize: "10px", fontWeight: "bold" }} />
                  <Bar dataKey="WhatsApp" stackId="a" fill="#10b981" />
                  <Bar dataKey="Mala Direta" stackId="a" fill="#8b5cf6" />
                  <Bar dataKey="Bot Auto" stackId="a" fill="#0ea5e9" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-slate-400 text-xs italic">
                Nenhum dado com os filtros atuais.
              </div>
            )}
          </div>
        </div>

        {/* Proporção por Canal & Evolução Diária */}
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <TrendingUp className="text-blue-600" size={18} />
              <h4 className="text-sm font-bold text-slate-800">Distribuição por Canal</h4>
            </div>
          </div>

          <div className="h-[260px] w-full">
            {byTypeData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={byTypeData}
                    cx="50%"
                    cy="50%"
                    innerRadius={50}
                    outerRadius={80}
                    paddingAngle={4}
                    dataKey="value"
                  >
                    {byTypeData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    contentStyle={{
                      borderRadius: "12px",
                      border: "none",
                      boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                    }}
                  />
                  <Legend wrapperStyle={{ fontSize: "10px", fontWeight: "bold" }} />
                </PieChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-slate-400 text-xs italic">
                Nenhum dado com os filtros atuais.
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Timeline Chart if available */}
      {timelineData.length > 0 && (
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Clock className="text-indigo-600" size={18} />
              <h4 className="text-sm font-bold text-slate-800">Evolução dos Envios por Dia</h4>
            </div>
            <span className="text-xs font-bold text-slate-400">Últimos dias com atividade</span>
          </div>
          <div className="h-[220px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={timelineData}>
                <defs>
                  <linearGradient id="colorWhats" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#10b981" stopOpacity={0.4} />
                    <stop offset="95%" stopColor="#10b981" stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="colorMala" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#8b5cf6" stopOpacity={0.4} />
                    <stop offset="95%" stopColor="#8b5cf6" stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="colorBot" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.4} />
                    <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: "#64748b" }} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: "#64748b" }} />
                <Tooltip
                  contentStyle={{
                    borderRadius: "12px",
                    border: "none",
                    boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                  }}
                />
                <Legend wrapperStyle={{ fontSize: "10px", fontWeight: "bold" }} />
                <Area type="monotone" dataKey="whats" name="WhatsApp" stroke="#10b981" fillOpacity={1} fill="url(#colorWhats)" />
                <Area type="monotone" dataKey="mala" name="Mala Direta" stroke="#8b5cf6" fillOpacity={1} fill="url(#colorMala)" />
                <Area type="monotone" dataKey="bot" name="Bot Auto" stroke="#0ea5e9" fillOpacity={1} fill="url(#colorBot)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {/* Detailed History Table */}
      <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
          <div>
            <h4 className="text-lg font-bold text-slate-800">
              Histórico Detalhado dos Candidatos Contactados
            </h4>
            <p className="text-xs text-slate-500">
              Exibindo {filteredData.length} registros com os filtros aplicados
            </p>
          </div>
          <div className="text-xs font-bold text-slate-400 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100">
            Clique no candidato para ver todo o histórico de envios
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                <th className="px-5 py-3.5 w-12 text-slate-400">#</th>
                <th className="px-5 py-3.5">Candidato</th>
                <th className="px-5 py-3.5">Curso</th>
                <th className="px-5 py-3.5">Base / Origem</th>
                <th className="px-5 py-3.5">Canal de Envio</th>
                <th className="px-5 py-3.5">Data & Hora</th>
                <th className="px-5 py-3.5">Frequência</th>
                <th className="px-5 py-3.5">Operador / Usuário</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs">
              {filteredData.map((item, idx) => {
                const d = parseDate(item.dataHora);
                const dataStr = d ? d.toLocaleDateString("pt-BR") : "N/A";
                const horaStr = d
                  ? d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", second: "2-digit" })
                  : "";
                const freqKey = item.nome?.toLowerCase().trim() || item.telefone || "";
                const totalVezes = candidateFrequencyMap.get(freqKey) || 1;

                return (
                  <tr key={item.id || idx} className="hover:bg-slate-50/70 transition-all">
                    <td className="px-5 py-3.5 text-slate-400 font-mono font-bold">
                      {idx + 1}
                    </td>
                    <td className="px-5 py-3.5">
                      <button
                        onClick={() => openCandidateHistory(item.nome)}
                        className="text-left font-bold text-slate-900 hover:text-emerald-600 hover:underline transition-colors flex flex-col"
                        title="Ver histórico completo deste candidato"
                      >
                        <span>{item.nome || "Não Informado"}</span>
                        {item.telefone && (
                          <span className="text-[11px] text-slate-400 font-normal font-mono">
                            {item.telefone}
                          </span>
                        )}
                      </button>
                    </td>
                    <td className="px-5 py-3.5 text-slate-700 font-medium max-w-[180px] truncate">
                      {item.curso || "Não informado"}
                    </td>
                    <td className="px-5 py-3.5 text-slate-600 font-medium max-w-[150px] truncate">
                      <span className="bg-slate-100 px-2 py-0.5 rounded-md text-[11px]">
                        {item.base || "Base Geral"}
                      </span>
                    </td>
                    <td className="px-5 py-3.5">
                      {item.tipoEnvio === "whats" && (
                        <span className="inline-flex items-center gap-1 bg-emerald-100 text-emerald-700 font-bold px-2.5 py-1 rounded-full text-[11px]">
                          <MessageSquare size={12} /> WhatsApp
                        </span>
                      )}
                      {item.tipoEnvio === "maladireta" && (
                        <span className="inline-flex items-center gap-1 bg-purple-100 text-purple-700 font-bold px-2.5 py-1 rounded-full text-[11px]">
                          <Mail size={12} /> Mala Direta
                        </span>
                      )}
                      {item.tipoEnvio === "bot_automatico" && (
                        <span className="inline-flex items-center gap-1 bg-sky-100 text-sky-700 font-bold px-2.5 py-1 rounded-full text-[11px]">
                          <Bot size={12} /> Bot Automático
                        </span>
                      )}
                    </td>
                    <td className="px-5 py-3.5 whitespace-nowrap">
                      <div className="font-semibold text-slate-800">{dataStr}</div>
                      <div className="text-[10px] text-slate-400">{horaStr}</div>
                    </td>
                    <td className="px-5 py-3.5">
                      <button
                        onClick={() => openCandidateHistory(item.nome)}
                        className={cn(
                          "px-2.5 py-0.5 rounded-full font-black text-[10px] transition-all",
                          totalVezes > 1
                            ? "bg-amber-100 text-amber-800 hover:bg-amber-200"
                            : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                        )}
                        title="Ver histórico detalhado"
                      >
                        {totalVezes}x {totalVezes > 1 ? "envios" : "envio"}
                      </button>
                    </td>
                    <td className="px-5 py-3.5 text-slate-500 font-medium">
                      {item.usuarioNome || "Sistema"}
                    </td>
                  </tr>
                );
              })}

              {filteredData.length === 0 && (
                <tr>
                  <td colSpan={8} className="px-6 py-12 text-center text-slate-400 italic">
                    Nenhum registro de mensagem enviada encontrado com os filtros selecionados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Candidate Full History Modal */}
      {selectedCandidateHistory && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full max-h-[85vh] overflow-hidden flex flex-col border border-slate-100 animate-in fade-in zoom-in duration-200">
            {/* Modal Header */}
            <div className="p-6 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="p-2.5 bg-emerald-500/20 text-emerald-400 rounded-2xl border border-emerald-500/30">
                  <HistoryIcon size={22} />
                </div>
                <div>
                  <h3 className="text-lg font-bold">Histórico Completo de Envios</h3>
                  <p className="text-xs text-slate-300">
                    Candidato: <span className="font-bold text-white">{selectedCandidateHistory.nome}</span>
                  </p>
                </div>
              </div>
              <button
                onClick={() => setSelectedCandidateHistory(null)}
                className="text-slate-400 hover:text-white p-1 rounded-lg transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto space-y-4 flex-1">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 flex items-center justify-between">
                <div>
                  <span className="text-xs font-bold text-slate-400 uppercase">Total de vezes contatado</span>
                  <p className="text-2xl font-black text-slate-900">
                    {selectedCandidateHistory.logs.length}x
                  </p>
                </div>
                <div className="flex gap-2">
                  <div className="text-right">
                    <span className="text-[10px] font-bold text-emerald-600 uppercase">WhatsApp</span>
                    <p className="font-bold text-sm text-slate-800">
                      {selectedCandidateHistory.logs.filter((l) => l.tipoEnvio === "whats").length}
                    </p>
                  </div>
                  <div className="text-right pl-3 border-l border-slate-200">
                    <span className="text-[10px] font-bold text-purple-600 uppercase">Mala Direta</span>
                    <p className="font-bold text-sm text-slate-800">
                      {selectedCandidateHistory.logs.filter((l) => l.tipoEnvio === "maladireta").length}
                    </p>
                  </div>
                  <div className="text-right pl-3 border-l border-slate-200">
                    <span className="text-[10px] font-bold text-sky-600 uppercase">Bot Auto</span>
                    <p className="font-bold text-sm text-slate-800">
                      {selectedCandidateHistory.logs.filter((l) => l.tipoEnvio === "bot_automatico").length}
                    </p>
                  </div>
                </div>
              </div>

              <div className="space-y-2.5">
                <h5 className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                  Timeline de Disparos e Envios
                </h5>
                <div className="divide-y divide-slate-100 border border-slate-100 rounded-2xl overflow-hidden">
                  {selectedCandidateHistory.logs.map((log, idx) => {
                    const d = parseDate(log.dataHora);
                    const dataStr = d ? d.toLocaleDateString("pt-BR") : "N/A";
                    const horaStr = d
                      ? d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", second: "2-digit" })
                      : "";

                    return (
                      <div
                        key={log.id || idx}
                        className="p-4 bg-white hover:bg-slate-50 flex items-center justify-between gap-3 transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <div className="text-xs font-mono font-bold text-slate-400 w-5">
                            #{selectedCandidateHistory.logs.length - idx}
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              {log.tipoEnvio === "whats" && (
                                <span className="inline-flex items-center gap-1 bg-emerald-100 text-emerald-700 font-bold px-2 py-0.5 rounded-md text-[10px]">
                                  <MessageSquare size={11} /> WhatsApp
                                </span>
                              )}
                              {log.tipoEnvio === "maladireta" && (
                                <span className="inline-flex items-center gap-1 bg-purple-100 text-purple-700 font-bold px-2 py-0.5 rounded-md text-[10px]">
                                  <Mail size={11} /> Mala Direta
                                </span>
                              )}
                              {log.tipoEnvio === "bot_automatico" && (
                                <span className="inline-flex items-center gap-1 bg-sky-100 text-sky-700 font-bold px-2 py-0.5 rounded-md text-[10px]">
                                  <Bot size={11} /> Bot Automático
                                </span>
                              )}
                              <span className="text-xs text-slate-600 font-medium">
                                Base: {log.base || "Base Geral"}
                              </span>
                            </div>
                            <div className="text-[11px] text-slate-400 mt-1">
                              Curso: <span className="text-slate-600 font-medium">{log.curso || "Não informado"}</span> • Operador: <span className="text-slate-600">{log.usuarioNome || "Sistema"}</span>
                            </div>
                          </div>
                        </div>

                        <div className="text-right">
                          <div className="text-xs font-bold text-slate-800">{dataStr}</div>
                          <div className="text-[11px] font-mono text-slate-400">{horaStr}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-100 flex justify-end">
              <button
                onClick={() => setSelectedCandidateHistory(null)}
                className="px-5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold rounded-xl text-xs transition-colors"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
