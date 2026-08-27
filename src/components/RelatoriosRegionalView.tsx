import React, { useState, useEffect, useMemo } from "react";
import {
  BarChart3,
  Download,
  Printer,
  Search,
  Filter,
  Users,
  Building,
  Briefcase,
  Gift,
  CheckCircle2,
  PieChart,
  FileSpreadsheet,
} from "lucide-react";
import { Funcionario, ChecklistItem, UserProfile } from "../types";
import { db, COLLECTIONS } from "../firebase";
import { collection, onSnapshot } from "firebase/firestore";
import * as XLSX from "xlsx";

interface Props {
  profile: UserProfile;
  onToast: (msg: string, type?: "success" | "error") => void;
  uniqueUnidades?: string[];
}

const MESES = [
  { id: "01", nome: "Janeiro" },
  { id: "02", nome: "Fevereiro" },
  { id: "03", nome: "Março" },
  { id: "04", nome: "Abril" },
  { id: "05", nome: "Maio" },
  { id: "06", nome: "Junho" },
  { id: "07", nome: "Julho" },
  { id: "08", nome: "Agosto" },
  { id: "09", nome: "Setembro" },
  { id: "10", nome: "Outubro" },
  { id: "11", nome: "Novembro" },
  { id: "12", nome: "Dezembro" },
];

export function RelatoriosRegionalView({ profile, onToast, uniqueUnidades = [] }: Props) {
  const [funcionarios, setFuncionarios] = useState<Funcionario[]>([]);
  const [checklist, setChecklist] = useState<ChecklistItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [unitFilter, setUnitFilter] = useState("Todas");
  const [statusFilter, setStatusFilter] = useState("Todos");

  useEffect(() => {
    const unsubFunc = onSnapshot(collection(db, COLLECTIONS.FUNCIONARIOS), (snap) => {
      const list = snap.docs.map((d) => ({
        id: d.id,
        ...d.data(),
      })) as Funcionario[];
      setFuncionarios(list);
      setLoading(false);
    });

    const unsubCheck = onSnapshot(collection(db, COLLECTIONS.CHECKLIST), (snap) => {
      const list = snap.docs.map((d) => ({
        id: d.id,
        ...d.data(),
      })) as ChecklistItem[];
      setChecklist(list);
    });

    return () => {
      unsubFunc();
      unsubCheck();
    };
  }, []);

  // Stats by Unit
  const statsByUnit = useMemo(() => {
    const counts: Record<string, number> = {};
    funcionarios.forEach((f) => {
      const unit = f.unidade?.trim() || "Sem Unidade";
      counts[unit] = (counts[unit] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [funcionarios]);

  // Stats by Cargo
  const statsByCargo = useMemo(() => {
    const counts: Record<string, number> = {};
    funcionarios.forEach((f) => {
      const cargo = f.cargo?.trim() || "Consultor SM";
      counts[cargo] = (counts[cargo] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [funcionarios]);

  // Birthdays by Month
  const birthdaysByMonth = useMemo(() => {
    const counts: Record<string, number> = {};
    MESES.forEach((m) => (counts[m.id] = 0));

    funcionarios.forEach((f) => {
      if (!f.dataNascimento) return;
      let month = "";
      if (f.dataNascimento.includes("-")) {
        month = f.dataNascimento.split("-")[1];
      } else if (f.dataNascimento.includes("/")) {
        month = f.dataNascimento.split("/")[1]?.padStart(2, "0");
      }
      if (counts[month] !== undefined) {
        counts[month]++;
      }
    });

    return MESES.map((m) => ({
      id: m.id,
      nome: m.nome,
      total: counts[m.id] || 0,
    }));
  }, [funcionarios]);

  // Filtered List
  const filteredList = useMemo(() => {
    return funcionarios.filter((f) => {
      if (unitFilter !== "Todas" && f.unidade !== unitFilter) return false;
      if (statusFilter !== "Todos" && (f.status || "Ativo") !== statusFilter) return false;
      if (search.trim()) {
        const term = search.toLowerCase();
        return (
          f.nome?.toLowerCase().includes(term) ||
          f.matricula?.toLowerCase().includes(term) ||
          f.cargo?.toLowerCase().includes(term) ||
          f.email?.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [funcionarios, unitFilter, statusFilter, search]);

  const handleExportConsolidatedExcel = () => {
    const dataToExport = filteredList.map((f) => ({
      Nome: f.nome,
      Matricula: f.matricula,
      Cargo: f.cargo || "Consultor SM",
      Unidade: f.unidade || "Geral",
      Telefone: f.telefone || "",
      Email: f.email || "",
      "Data Nascimento": f.dataNascimento || "",
      Status: f.status || "Ativo",
      Observações: f.observacao || "",
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio_Regional_SM");
    XLSX.writeFile(wb, `Relatorio_Regional_SM_${new Date().toISOString().split("T")[0]}.xlsx`);
    onToast("Relatório exportado com sucesso!", "success");
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto pb-16">
      {/* Header */}
      <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-6 print:hidden">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="px-3 py-1 bg-blue-50 text-blue-700 text-xs font-extrabold rounded-full border border-blue-200 flex items-center gap-1.5">
              <BarChart3 size={13} />
              Inteligência e Relatórios
            </span>
          </div>
          <h1 className="text-2xl font-black text-slate-900">Relatórios do Servidor Regional</h1>
          <p className="text-xs text-slate-500 mt-1">
            Consolidado executivo com estatísticas de equipe, aniversários e rotinas operacionais
          </p>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={handlePrint}
            className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer"
          >
            <Printer size={15} />
            <span>Imprimir / PDF</span>
          </button>

          <button
            onClick={handleExportConsolidatedExcel}
            className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition flex items-center gap-2 cursor-pointer"
          >
            <Download size={15} />
            <span>Exportar XLS</span>
          </button>
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Colaboradores SM</p>
            <p className="text-2xl font-black text-slate-900 mt-1">{funcionarios.length}</p>
            <p className="text-[11px] text-slate-500 mt-0.5">Base Regional SM</p>
          </div>
          <div className="w-12 h-12 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center">
            <Users size={22} />
          </div>
        </div>

        <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Polos & Unidades</p>
            <p className="text-2xl font-black text-sky-600 mt-1">{statsByUnit.length}</p>
            <p className="text-[11px] text-slate-500 mt-0.5">Com equipe alocada</p>
          </div>
          <div className="w-12 h-12 bg-sky-50 text-sky-600 rounded-2xl flex items-center justify-center">
            <Building size={22} />
          </div>
        </div>

        <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Funções / Cargos</p>
            <p className="text-2xl font-black text-purple-600 mt-1">{statsByCargo.length}</p>
            <p className="text-[11px] text-slate-500 mt-0.5">Cargos mapeados</p>
          </div>
          <div className="w-12 h-12 bg-purple-50 text-purple-600 rounded-2xl flex items-center justify-center">
            <Briefcase size={22} />
          </div>
        </div>

        <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Conclusão Checklist</p>
            <p className="text-2xl font-black text-emerald-600 mt-1">
              {checklist.length > 0
                ? `${Math.round(
                    (checklist.filter((c) => c.concluido).length / checklist.length) * 100
                  )}%`
                : "100%"}
            </p>
            <p className="text-[11px] text-slate-500 mt-0.5">Rotinas executadas</p>
          </div>
          <div className="w-12 h-12 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center">
            <CheckCircle2 size={22} />
          </div>
        </div>
      </div>

      {/* Visual Analytics */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Unidades Breakdown */}
        <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
              <Building size={20} />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900">Distribuição por Unidade</h2>
              <p className="text-xs text-slate-500">Quantidade de colaboradores em cada polo</p>
            </div>
          </div>

          <div className="space-y-3 pt-2">
            {statsByUnit.length === 0 ? (
              <p className="text-xs text-slate-400 py-6 text-center">Nenhuma unidade cadastrada.</p>
            ) : (
              statsByUnit.map(([unit, count]) => {
                const pct = funcionarios.length > 0 ? Math.round((count / funcionarios.length) * 100) : 0;
                return (
                  <div key={unit} className="space-y-1">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-bold text-slate-700">{unit}</span>
                      <span className="font-bold text-slate-900">{count} ({pct}%)</span>
                    </div>
                    <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden">
                      <div
                        className="h-full bg-blue-600 rounded-full transition-all duration-500"
                        style={{ width: `${pct}%` }}
                      />
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>

        {/* Aniversários no Ano */}
        <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
              <Gift size={20} />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900">Aniversários por Mês</h2>
              <p className="text-xs text-slate-500">Distribuição das celebrações ao longo do ano</p>
            </div>
          </div>

          <div className="grid grid-cols-3 sm:grid-cols-4 gap-2.5 pt-2">
            {birthdaysByMonth.map((m) => (
              <div
                key={m.id}
                className={`p-3 rounded-2xl border text-center transition ${
                  m.total > 0
                    ? "bg-amber-50/70 border-amber-200/80"
                    : "bg-slate-50 border-slate-100"
                }`}
              >
                <span className="block text-[11px] font-bold text-slate-600 truncate">{m.nome}</span>
                <span className={`block text-lg font-black mt-0.5 ${m.total > 0 ? "text-amber-700" : "text-slate-300"}`}>
                  {m.total}
                </span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Table Consolidated */}
      <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 print:hidden">
          <div>
            <h2 className="text-base font-bold text-slate-900">Listagem Detalhada da Equipe Regional</h2>
            <p className="text-xs text-slate-500">Filtre e analise os dados de todos os colaboradores</p>
          </div>

          <div className="flex flex-col sm:flex-row items-center gap-3">
            <div className="relative w-full sm:w-60">
              <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Buscar..."
                className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
              />
            </div>

            <select
              value={unitFilter}
              onChange={(e) => setUnitFilter(e.target.value)}
              className="w-full sm:w-auto px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
            >
              <option value="Todas">Todas as Unidades</option>
              {uniqueUnidades.map((u) => (
                <option key={u} value={u}>
                  {u}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50/80 text-slate-500 text-[11px] font-extrabold uppercase tracking-wider border-b border-slate-100">
                <th className="px-6 py-4">Nome</th>
                <th className="px-6 py-4">Matrícula</th>
                <th className="px-6 py-4">Cargo</th>
                <th className="px-6 py-4">Unidade</th>
                <th className="px-6 py-4">Aniversário</th>
                <th className="px-6 py-4">Telefone</th>
                <th className="px-6 py-4">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs">
              {filteredList.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-6 py-8 text-center text-slate-400">
                    Nenhum colaborador encontrado com os filtros selecionados.
                  </td>
                </tr>
              ) : (
                filteredList.map((f) => (
                  <tr key={f.id} className="hover:bg-slate-50/60 transition-colors">
                    <td className="px-6 py-4 font-bold text-slate-900">{f.nome}</td>
                    <td className="px-6 py-4 text-slate-500 font-mono text-[11px]">{f.matricula || "-"}</td>
                    <td className="px-6 py-4 text-slate-700">{f.cargo || "Consultor SM"}</td>
                    <td className="px-6 py-4 text-slate-600">{f.unidade || "Geral"}</td>
                    <td className="px-6 py-4 text-amber-800 font-bold">
                      {f.dataNascimento
                        ? f.dataNascimento.includes("-")
                          ? `${f.dataNascimento.split("-")[2]}/${f.dataNascimento.split("-")[1]}`
                          : f.dataNascimento
                        : "-"}
                    </td>
                    <td className="px-6 py-4 text-slate-600">{f.telefone || "-"}</td>
                    <td className="px-6 py-4">
                      <span
                        className={`px-2 py-0.5 rounded-full text-[10px] font-extrabold uppercase ${
                          f.status === "Inativo"
                            ? "bg-rose-50 text-rose-600"
                            : "bg-emerald-50 text-emerald-700"
                        }`}
                      >
                        {f.status || "Ativo"}
                      </span>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

export default RelatoriosRegionalView;
