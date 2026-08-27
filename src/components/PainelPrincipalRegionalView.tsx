import React, { useState, useMemo, useEffect } from "react";
import {
  Calendar,
  Gift,
  ExternalLink,
  Search,
  Users,
  Link2,
  CheckCircle2,
  Clock,
  Sparkles,
  Phone,
  Copy,
  Building,
  Briefcase,
  ChevronRight,
  Filter,
  CheckSquare,
} from "lucide-react";
import { LinkUtil, Funcionario, UserProfile, ChecklistItem } from "../types";
import { db, COLLECTIONS } from "../firebase";
import { collection, onSnapshot, query } from "firebase/firestore";

interface Props {
  profile: UserProfile;
  onToast: (msg: string, type?: "success" | "error") => void;
  links?: LinkUtil[];
  onNavigateTab?: (tab: string) => void;
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

export function PainelPrincipalRegionalView({
  profile,
  onToast,
  links: propLinks,
  onNavigateTab,
}: Props) {
  const [activeSubTab, setActiveSubTab] = useState<"aniversariantes" | "links" | "visaoGeral">("visaoGeral");
  const [funcionarios, setFuncionarios] = useState<Funcionario[]>([]);
  const [linksList, setLinksList] = useState<LinkUtil[]>(propLinks || []);
  const [checklist, setChecklist] = useState<ChecklistItem[]>([]);
  const [loading, setLoading] = useState(true);

  // Month selector for birthdays
  const currentMonthNumber = String(new Date().getMonth() + 1).padStart(2, "0");
  const currentDay = new Date().getDate();
  const [selectedMonth, setSelectedMonth] = useState<string>(currentMonthNumber);
  const [searchBirthday, setSearchBirthday] = useState("");
  const [searchLink, setSearchLink] = useState("");
  const [linkCategoryFilter, setLinkCategoryFilter] = useState("Todas");

  // Fetch Funcionários SM and Links from Firestore in Realtime
  useEffect(() => {
    const unsubFunc = onSnapshot(
      collection(db, COLLECTIONS.FUNCIONARIOS),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as Funcionario[];
        setFuncionarios(list);
        setLoading(false);
      },
      (err) => {
        console.error("Erro ao buscar funcionários:", err);
        setLoading(false);
      }
    );

    const unsubLinks = onSnapshot(
      collection(db, COLLECTIONS.LINKS),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as LinkUtil[];
        setLinksList(list);
      },
      (err) => {
        console.error("Erro ao buscar links úteis:", err);
      }
    );

    const unsubChecklist = onSnapshot(
      collection(db, COLLECTIONS.CHECKLIST),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as ChecklistItem[];
        setChecklist(list);
      },
      (err) => {
        console.error("Erro ao buscar checklist:", err);
      }
    );

    return () => {
      unsubFunc();
      unsubLinks();
      unsubChecklist();
    };
  }, []);

  // Filter birthdays by selected month
  const aniversariantesMes = useMemo(() => {
    return funcionarios
      .filter((f) => {
        if (!f.dataNascimento) return false;
        // Parse date format (YYYY-MM-DD or DD/MM/YYYY)
        let month = "";
        let day = 0;

        if (f.dataNascimento.includes("-")) {
          const parts = f.dataNascimento.split("-");
          month = parts[1];
          day = parseInt(parts[2], 10);
        } else if (f.dataNascimento.includes("/")) {
          const parts = f.dataNascimento.split("/");
          month = parts[1]?.padStart(2, "0");
          day = parseInt(parts[0], 10);
        }

        const matchMonth = month === selectedMonth;
        if (!matchMonth) return false;

        if (searchBirthday.trim()) {
          const term = searchBirthday.toLowerCase();
          const matchName = f.nome?.toLowerCase().includes(term);
          const matchUnit = f.unidade?.toLowerCase().includes(term);
          const matchRole = f.cargo?.toLowerCase().includes(term);
          return matchName || matchUnit || matchRole;
        }

        return true;
      })
      .sort((a, b) => {
        const getDay = (dateStr?: string) => {
          if (!dateStr) return 0;
          if (dateStr.includes("-")) return parseInt(dateStr.split("-")[2] || "0", 10);
          if (dateStr.includes("/")) return parseInt(dateStr.split("/")[0] || "0", 10);
          return 0;
        };
        return getDay(a.dataNascimento) - getDay(b.dataNascimento);
      });
  }, [funcionarios, selectedMonth, searchBirthday]);

  // Aniversariantes de Hoje
  const aniversariantesHoje = useMemo(() => {
    return funcionarios.filter((f) => {
      if (!f.dataNascimento) return false;
      let month = "";
      let day = 0;
      if (f.dataNascimento.includes("-")) {
        const parts = f.dataNascimento.split("-");
        month = parts[1];
        day = parseInt(parts[2], 10);
      } else if (f.dataNascimento.includes("/")) {
        const parts = f.dataNascimento.split("/");
        month = parts[1]?.padStart(2, "0");
        day = parseInt(parts[0], 10);
      }
      return month === currentMonthNumber && day === currentDay;
    });
  }, [funcionarios, currentMonthNumber, currentDay]);

  // Useful links filtered
  const linkCategories = useMemo(() => {
    const cats = new Set<string>();
    linksList.forEach((l) => {
      if (l.local) cats.add(l.local);
    });
    return ["Todas", ...Array.from(cats)];
  }, [linksList]);

  const filteredLinks = useMemo(() => {
    return linksList.filter((l) => {
      const matchCat = linkCategoryFilter === "Todas" || l.local === linkCategoryFilter;
      if (!matchCat) return false;
      if (searchLink.trim()) {
        const term = searchLink.toLowerCase();
        return (
          l.nome?.toLowerCase().includes(term) ||
          l.url?.toLowerCase().includes(term) ||
          l.local?.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [linksList, linkCategoryFilter, searchLink]);

  const checklistStats = useMemo(() => {
    const total = checklist.length;
    const concluidos = checklist.filter((c) => c.concluido).length;
    const pendentes = total - concluidos;
    const pct = total > 0 ? Math.round((concluidos / total) * 100) : 0;
    return { total, concluidos, pendentes, pct };
  }, [checklist]);

  const formatBirthDate = (dateStr?: string) => {
    if (!dateStr) return "";
    if (dateStr.includes("-")) {
      const [year, month, day] = dateStr.split("-");
      return `${day}/${month}`;
    }
    return dateStr.slice(0, 5);
  };

  const handleSendWhatsAppCongrats = (func: Funcionario) => {
    if (!func.telefone) {
      onToast("Funcionário não possui telefone cadastrado.", "error");
      return;
    }
    const cleanPhone = func.telefone.replace(/\D/g, "");
    const firstName = func.nome.split(" ")[0];
    const message = encodeURIComponent(
      `Olá ${firstName}! 🎂🎉 Parabéns pelo seu aniversário! Desejamos muitas felicidades, saúde, sucesso e realizações em sua jornada na equipe Regional SM!`
    );
    window.open(`https://wa.me/55${cleanPhone}?text=${message}`, "_blank");
  };

  const handleCopyLink = (url: string) => {
    navigator.clipboard.writeText(url);
    onToast("Link copiado para a área de transferência!", "success");
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto pb-16">
      {/* Banner / Welcome Header */}
      <div className="bg-gradient-to-r from-[#031d44] via-[#052b66] to-[#0a3d8f] rounded-3xl p-6 md:p-8 text-white shadow-xl border border-blue-900/40 relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-sky-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="absolute bottom-0 right-32 w-64 h-64 bg-emerald-500/10 rounded-full blur-2xl pointer-events-none" />

        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div>
            <div className="flex items-center gap-2 mb-2">
              <span className="px-3 py-1 bg-sky-500/20 text-sky-300 text-xs font-extrabold rounded-full border border-sky-400/30 flex items-center gap-1.5">
                <Sparkles size={13} className="text-sky-300" />
                Servidor Regional SM
              </span>
              <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-300 text-xs font-bold rounded-full border border-emerald-400/30">
                {profile.role}
              </span>
            </div>
            <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight">
              Olá, {profile.name}!
            </h1>
            <p className="text-blue-200 text-sm mt-1 max-w-2xl">
              Painel de Gestão e Rotinas do Servidor Regional. Acompanhe aniversariantes do mês, links úteis e rotinas operacionais da equipe.
            </p>
          </div>

          <div className="flex items-center gap-3">
            <div className="bg-white/10 backdrop-blur-md px-4 py-3 rounded-2xl border border-white/10 text-center">
              <span className="block text-2xl font-black text-white">{funcionarios.length}</span>
              <span className="text-[11px] font-bold text-blue-200 uppercase tracking-wider">Funcionários SM</span>
            </div>
            <div className="bg-white/10 backdrop-blur-md px-4 py-3 rounded-2xl border border-white/10 text-center">
              <span className="block text-2xl font-black text-amber-300">{aniversariantesHoje.length}</span>
              <span className="text-[11px] font-bold text-blue-200 uppercase tracking-wider">Aniversário Hoje</span>
            </div>
          </div>
        </div>
      </div>

      {/* Sub-tab Navigation */}
      <div className="flex flex-wrap items-center gap-2 bg-slate-100/80 p-1.5 rounded-2xl w-fit border border-slate-200">
        <button
          onClick={() => setActiveSubTab("visaoGeral")}
          className={`px-5 py-2.5 rounded-xl font-bold text-xs transition-all flex items-center gap-2 cursor-pointer ${
            activeSubTab === "visaoGeral"
              ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
              : "text-slate-600 hover:text-slate-900"
          }`}
        >
          <Sparkles size={16} />
          <span>Visão Geral</span>
        </button>
        <button
          onClick={() => setActiveSubTab("aniversariantes")}
          className={`px-5 py-2.5 rounded-xl font-bold text-xs transition-all flex items-center gap-2 cursor-pointer ${
            activeSubTab === "aniversariantes"
              ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
              : "text-slate-600 hover:text-slate-900"
          }`}
        >
          <Gift size={16} className="text-amber-500" />
          <span>Aniversariantes do Mês</span>
          {aniversariantesMes.length > 0 && (
            <span className="px-2 py-0.5 bg-amber-100 text-amber-700 text-[10px] font-extrabold rounded-full">
              {aniversariantesMes.length}
            </span>
          )}
        </button>
        <button
          onClick={() => setActiveSubTab("links")}
          className={`px-5 py-2.5 rounded-xl font-bold text-xs transition-all flex items-center gap-2 cursor-pointer ${
            activeSubTab === "links"
              ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
              : "text-slate-600 hover:text-slate-900"
          }`}
        >
          <Link2 size={16} className="text-sky-500" />
          <span>Links Úteis</span>
          <span className="px-2 py-0.5 bg-slate-200 text-slate-700 text-[10px] font-extrabold rounded-full">
            {linksList.length}
          </span>
        </button>
      </div>

      {/* TODAY'S BIRTHDAY ALERT (If any) */}
      {aniversariantesHoje.length > 0 && (
        <div className="bg-gradient-to-r from-amber-500/15 via-orange-500/10 to-amber-500/15 border-2 border-amber-400/40 rounded-3xl p-5 md:p-6 shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-amber-500 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-amber-500/30 shrink-0">
              <Gift size={24} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 bg-amber-500 text-white text-[10px] font-black uppercase tracking-wider rounded-md">
                  Hoje é dia de festa!
                </span>
                <span className="text-xs text-amber-800 font-bold">
                  {new Date().toLocaleDateString("pt-BR", { day: "numeric", month: "long" })}
                </span>
              </div>
              <h3 className="text-lg font-black text-slate-900 mt-1">
                {aniversariantesHoje.map((a) => a.nome).join(", ")}
              </h3>
              <p className="text-xs text-slate-600">
                {aniversariantesHoje.length === 1
                  ? "está fazendo aniversário hoje! Envie uma mensagem de felicitações."
                  : "estão fazendo aniversário hoje! Deseje parabéns à equipe."}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {aniversariantesHoje.map((a) => (
              <button
                key={a.id}
                onClick={() => handleSendWhatsAppCongrats(a)}
                className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl shadow-md shadow-emerald-600/20 transition flex items-center gap-1.5 cursor-pointer"
              >
                <Phone size={14} />
                <span>Parabenizar {a.nome.split(" ")[0]}</span>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 1. VISÃO GERAL (Cards & Combined Sections) */}
      {/* ========================================================================= */}
      {activeSubTab === "visaoGeral" && (
        <div className="space-y-8">
          {/* Quick Metrics Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Funcionários SM</p>
                <p className="text-2xl font-black text-slate-900 mt-1">{funcionarios.length}</p>
                <p className="text-[11px] text-slate-500 mt-0.5">Cadastrados no regional</p>
              </div>
              <div className="w-12 h-12 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center">
                <Users size={22} />
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Aniversariantes ({MESES.find((m) => m.id === currentMonthNumber)?.nome})</p>
                <p className="text-2xl font-black text-amber-600 mt-1">{aniversariantesMes.length}</p>
                <p className="text-[11px] text-slate-500 mt-0.5">Neste mês</p>
              </div>
              <div className="w-12 h-12 bg-amber-50 text-amber-600 rounded-2xl flex items-center justify-center">
                <Gift size={22} />
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Links Úteis</p>
                <p className="text-2xl font-black text-sky-600 mt-1">{linksList.length}</p>
                <p className="text-[11px] text-slate-500 mt-0.5">Atalhos rápidos</p>
              </div>
              <div className="w-12 h-12 bg-sky-50 text-sky-600 rounded-2xl flex items-center justify-center">
                <Link2 size={22} />
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Checklist SM</p>
                <p className="text-2xl font-black text-emerald-600 mt-1">
                  {checklistStats.pct}%
                </p>
                <p className="text-[11px] text-slate-500 mt-0.5">{checklistStats.concluidos}/{checklistStats.total} Concluídos</p>
              </div>
              <div className="w-12 h-12 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center">
                <CheckSquare size={22} />
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* Left Column: Aniversariantes do Mês Preview */}
            <div className="lg:col-span-7 bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between mb-6">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
                      <Gift size={20} />
                    </div>
                    <div>
                      <h2 className="text-base font-bold text-slate-900">
                        Aniversariantes de {MESES.find((m) => m.id === selectedMonth)?.nome}
                      </h2>
                      <p className="text-xs text-slate-500">Comemorações da equipe neste mês</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setActiveSubTab("aniversariantes")}
                    className="text-xs font-bold text-blue-600 hover:text-blue-700 flex items-center gap-1 cursor-pointer"
                  >
                    <span>Ver todos</span>
                    <ChevronRight size={14} />
                  </button>
                </div>

                {aniversariantesMes.length === 0 ? (
                  <div className="text-center py-10 bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                    <Gift size={32} className="text-slate-300 mx-auto mb-2" />
                    <p className="text-xs font-bold text-slate-600">Nenhum aniversariante encontrado neste mês.</p>
                    <p className="text-[11px] text-slate-400 mt-0.5">Cadastre a data de nascimento dos colaboradores em Administração.</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 max-h-[360px] overflow-y-auto pr-1">
                    {aniversariantesMes.slice(0, 6).map((func) => {
                      const isToday = aniversariantesHoje.some((a) => a.id === func.id);
                      return (
                        <div
                          key={func.id}
                          className={`p-4 rounded-2xl border transition-all flex flex-col justify-between ${
                            isToday
                              ? "bg-amber-50/80 border-amber-300 shadow-sm"
                              : "bg-slate-50 hover:bg-slate-100/80 border-slate-100"
                          }`}
                        >
                          <div className="flex items-start justify-between gap-2">
                            <div>
                              <span className="font-bold text-slate-900 text-sm line-clamp-1">
                                {func.nome}
                              </span>
                              <div className="flex items-center gap-1.5 text-[11px] text-slate-500 mt-0.5">
                                <Building size={12} className="shrink-0" />
                                <span className="truncate">{func.unidade || "Geral"}</span>
                              </div>
                              {func.cargo && (
                                <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                                  <Briefcase size={12} className="shrink-0" />
                                  <span className="truncate">{func.cargo}</span>
                                </div>
                              )}
                            </div>
                            <div className="px-2.5 py-1 bg-white rounded-xl shadow-xs border border-slate-200/80 text-center shrink-0">
                              <span className="block text-xs font-black text-amber-600">
                                {formatBirthDate(func.dataNascimento)}
                              </span>
                            </div>
                          </div>

                          {func.telefone && (
                            <button
                              onClick={() => handleSendWhatsAppCongrats(func)}
                              className="mt-3 w-full py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 text-[11px] font-bold rounded-lg transition flex items-center justify-center gap-1 cursor-pointer"
                            >
                              <Phone size={12} />
                              <span>Parabenizar no WhatsApp</span>
                            </button>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {onNavigateTab && (
                <div className="mt-6 pt-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
                  <span>Quer cadastrar mais colaboradores?</span>
                  <button
                    onClick={() => onNavigateTab("admin")}
                    className="font-bold text-blue-600 hover:underline cursor-pointer"
                  >
                    Ir para Administração →
                  </button>
                </div>
              )}
            </div>

            {/* Right Column: Links Úteis Preview */}
            <div className="lg:col-span-5 bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between mb-6">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center">
                      <Link2 size={20} />
                    </div>
                    <div>
                      <h2 className="text-base font-bold text-slate-900">Links Úteis</h2>
                      <p className="text-xs text-slate-500">Acesso rápido aos sistemas e portais</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setActiveSubTab("links")}
                    className="text-xs font-bold text-blue-600 hover:text-blue-700 flex items-center gap-1 cursor-pointer"
                  >
                    <span>Ver todos</span>
                    <ChevronRight size={14} />
                  </button>
                </div>

                {linksList.length === 0 ? (
                  <div className="text-center py-10 bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                    <Link2 size={32} className="text-slate-300 mx-auto mb-2" />
                    <p className="text-xs font-bold text-slate-600">Nenhum link cadastrado ainda.</p>
                    <p className="text-[11px] text-slate-400 mt-0.5">Cadastre links rápidos na aba Administração.</p>
                  </div>
                ) : (
                  <div className="space-y-2.5 max-h-[360px] overflow-y-auto pr-1">
                    {linksList.slice(0, 6).map((link) => (
                      <div
                        key={link.id}
                        className="p-3.5 bg-slate-50 hover:bg-blue-50/50 rounded-2xl border border-slate-100 hover:border-blue-200 transition-all flex items-center justify-between gap-3 group"
                      >
                        <div className="min-w-0">
                          <p className="font-bold text-slate-900 text-xs truncate group-hover:text-blue-600 transition-colors">
                            {link.nome}
                          </p>
                          {link.local && (
                            <span className="inline-block mt-0.5 px-2 py-0.5 bg-slate-200/80 text-slate-600 text-[10px] font-bold rounded-md uppercase">
                              {link.local}
                            </span>
                          )}
                        </div>

                        <div className="flex items-center gap-1.5 shrink-0">
                          <button
                            onClick={() => handleCopyLink(link.url)}
                            title="Copiar Link"
                            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-white rounded-lg transition cursor-pointer"
                          >
                            <Copy size={14} />
                          </button>
                          <a
                            href={link.url.startsWith("http") ? link.url : `https://${link.url}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="p-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg shadow-xs transition flex items-center gap-1 cursor-pointer"
                          >
                            <ExternalLink size={14} />
                          </a>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {onNavigateTab && (
                <div className="mt-6 pt-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
                  <span>Gerenciar todos os links</span>
                  <button
                    onClick={() => onNavigateTab("admin")}
                    className="font-bold text-blue-600 hover:underline cursor-pointer"
                  >
                    Acessar Links Úteis →
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 2. ABA ANIVERSARIANTES DO MÊS */}
      {/* ========================================================================= */}
      {activeSubTab === "aniversariantes" && (
        <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm space-y-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-6">
            <div>
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
                  <Gift size={22} />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-slate-900">Aniversariantes do Mês</h2>
                  <p className="text-xs text-slate-500">Consulte e celebre os aniversários da equipe Regional SM</p>
                </div>
              </div>
            </div>

            {/* Search Input */}
            <div className="relative w-full md:w-72">
              <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={searchBirthday}
                onChange={(e) => setSearchBirthday(e.target.value)}
                placeholder="Buscar colaborador..."
                className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-amber-500 outline-none transition"
              />
            </div>
          </div>

          {/* Month Filter Badges */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-2 scrollbar-none">
            {MESES.map((m) => {
              const isSelected = selectedMonth === m.id;
              const isCurrent = currentMonthNumber === m.id;
              return (
                <button
                  key={m.id}
                  onClick={() => setSelectedMonth(m.id)}
                  className={`px-4 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all flex items-center gap-1.5 cursor-pointer ${
                    isSelected
                      ? "bg-amber-500 text-white shadow-md shadow-amber-500/20"
                      : isCurrent
                      ? "bg-amber-50 text-amber-800 border border-amber-200 hover:bg-amber-100"
                      : "bg-slate-50 text-slate-600 hover:bg-slate-100 border border-slate-100"
                  }`}
                >
                  <span>{m.nome}</span>
                  {isCurrent && <span className="w-1.5 h-1.5 bg-amber-600 rounded-full" />}
                </button>
              );
            })}
          </div>

          {/* List of Birthday Celebrants */}
          {aniversariantesMes.length === 0 ? (
            <div className="text-center py-16 bg-slate-50 rounded-3xl border border-dashed border-slate-200 space-y-3">
              <div className="w-16 h-16 bg-amber-50 text-amber-400 rounded-full flex items-center justify-center mx-auto">
                <Gift size={32} />
              </div>
              <h3 className="text-base font-bold text-slate-800">
                Nenhum aniversariante encontrado em {MESES.find((m) => m.id === selectedMonth)?.nome}
              </h3>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                Adicione ou atualize a data de nascimento dos funcionários na aba Administração para visualizar seus aniversários aqui.
              </p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {aniversariantesMes.map((func) => {
                const isToday = aniversariantesHoje.some((a) => a.id === func.id);
                return (
                  <div
                    key={func.id}
                    className={`p-5 rounded-3xl border transition-all flex flex-col justify-between relative overflow-hidden ${
                      isToday
                        ? "bg-gradient-to-br from-amber-50 via-orange-50/50 to-white border-amber-300 shadow-md ring-2 ring-amber-400/30"
                        : "bg-white hover:border-slate-300 border-slate-100 shadow-sm"
                    }`}
                  >
                    {isToday && (
                      <div className="absolute top-0 right-0 bg-amber-500 text-white text-[9px] font-black uppercase tracking-wider px-3 py-1 rounded-bl-xl shadow-xs">
                        Aniversário Hoje! 🎂
                      </div>
                    )}

                    <div className="space-y-3">
                      <div className="flex items-start justify-between gap-2">
                        <div className="w-10 h-10 bg-gradient-to-tr from-amber-500 to-orange-500 text-white rounded-2xl flex items-center justify-center font-black text-sm shadow-md shadow-amber-500/20">
                          {func.nome.charAt(0)}
                        </div>
                        <div className="px-3 py-1 bg-amber-100/80 text-amber-800 rounded-xl text-xs font-black">
                          {formatBirthDate(func.dataNascimento)}
                        </div>
                      </div>

                      <div>
                        <h4 className="font-extrabold text-slate-900 text-sm">{func.nome}</h4>
                        <div className="flex flex-col gap-1 text-xs text-slate-500 mt-1">
                          <span className="flex items-center gap-1.5">
                            <Building size={13} className="text-slate-400 shrink-0" />
                            <span>{func.unidade || "Sem unidade"}</span>
                          </span>
                          {func.cargo && (
                            <span className="flex items-center gap-1.5">
                              <Briefcase size={13} className="text-slate-400 shrink-0" />
                              <span>{func.cargo}</span>
                            </span>
                          )}
                          {func.email && (
                            <span className="text-[11px] text-slate-400 truncate">{func.email}</span>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="mt-4 pt-3 border-t border-slate-100">
                      {func.telefone ? (
                        <button
                          onClick={() => handleSendWhatsAppCongrats(func)}
                          className="w-full py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl shadow-sm transition flex items-center justify-center gap-1.5 cursor-pointer"
                        >
                          <Phone size={14} />
                          <span>Enviar Parabéns (WhatsApp)</span>
                        </button>
                      ) : (
                        <span className="block text-center text-[11px] text-slate-400 py-1 font-medium">
                          Telefone não informado
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* 3. ABA LINKS ÚTEIS (Reflected view) */}
      {/* ========================================================================= */}
      {activeSubTab === "links" && (
        <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm space-y-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-6">
            <div>
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center">
                  <Link2 size={22} />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-slate-900">Links Úteis do Sistema</h2>
                  <p className="text-xs text-slate-500">Portais, sistemas operacionais e manuais cadastrados</p>
                </div>
              </div>
            </div>

            {/* Filters */}
            <div className="flex flex-col sm:flex-row items-center gap-3">
              <div className="relative w-full sm:w-64">
                <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  value={searchLink}
                  onChange={(e) => setSearchLink(e.target.value)}
                  placeholder="Buscar links..."
                  className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>

              {linkCategories.length > 1 && (
                <select
                  value={linkCategoryFilter}
                  onChange={(e) => setLinkCategoryFilter(e.target.value)}
                  className="w-full sm:w-auto px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                >
                  {linkCategories.map((c) => (
                    <option key={c} value={c}>
                      {c === "Todas" ? "Todas as Categorias" : c}
                    </option>
                  ))}
                </select>
              )}
            </div>
          </div>

          {filteredLinks.length === 0 ? (
            <div className="text-center py-16 bg-slate-50 rounded-3xl border border-dashed border-slate-200 space-y-3">
              <div className="w-16 h-16 bg-sky-50 text-sky-400 rounded-full flex items-center justify-center mx-auto">
                <Link2 size={32} />
              </div>
              <h3 className="text-base font-bold text-slate-800">Nenhum link encontrado</h3>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                Os links cadastrados na aba Administração aparecerão organizados aqui para acesso rápido de toda a equipe.
              </p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredLinks.map((link) => (
                <div
                  key={link.id}
                  className="p-5 bg-slate-50/70 hover:bg-blue-50/40 rounded-3xl border border-slate-100 hover:border-blue-200 transition-all flex flex-col justify-between gap-4 group shadow-xs"
                >
                  <div className="space-y-2">
                    <div className="flex items-start justify-between gap-2">
                      <span className="w-9 h-9 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center font-black text-xs shrink-0">
                        <ExternalLink size={16} />
                      </span>
                      {link.local && (
                        <span className="px-2.5 py-0.5 bg-slate-200 text-slate-700 text-[10px] font-extrabold rounded-md uppercase">
                          {link.local}
                        </span>
                      )}
                    </div>
                    <div>
                      <h4 className="font-bold text-slate-900 text-sm group-hover:text-blue-600 transition-colors">
                        {link.nome}
                      </h4>
                      <p className="text-xs text-slate-400 truncate mt-0.5">{link.url}</p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 pt-2 border-t border-slate-200/60">
                    <button
                      onClick={() => handleCopyLink(link.url)}
                      className="flex-1 py-2 bg-white hover:bg-slate-100 text-slate-600 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      <Copy size={13} />
                      <span>Copiar Link</span>
                    </button>
                    <a
                      href={link.url.startsWith("http") ? link.url : `https://${link.url}`}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="flex-1 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-xs transition flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      <span>Abrir</span>
                      <ExternalLink size={13} />
                    </a>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default PainelPrincipalRegionalView;
