import React, { useState, useEffect, useMemo } from "react";
import {
  CheckSquare,
  Square,
  Plus,
  Trash2,
  Filter,
  Search,
  CheckCircle2,
  Clock,
  AlertCircle,
  Calendar,
  Building,
  Tag,
  Sparkles,
  ListTodo,
} from "lucide-react";
import { db, COLLECTIONS } from "../firebase";
import {
  collection,
  onSnapshot,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  serverTimestamp,
} from "firebase/firestore";
import { ChecklistItem, UserProfile } from "../types";

interface Props {
  profile: UserProfile;
  onToast: (msg: string, type?: "success" | "error") => void;
  uniqueUnidades?: string[];
}

export function CheckListRegionalView({ profile, onToast, uniqueUnidades = [] }: Props) {
  const [items, setItems] = useState<ChecklistItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<"Todos" | "Pendentes" | "Concluidos">("Todos");
  const [categoryFilter, setCategoryFilter] = useState("Todas");
  const [isAdding, setIsAdding] = useState(false);

  // Form State
  const [titulo, setTitulo] = useState("");
  const [descricao, setDescricao] = useState("");
  const [categoria, setCategoria] = useState<
    "Abertura" | "Atendimento" | "Matrículas" | "Auditoria" | "Fechamento" | "Geral"
  >("Atendimento");
  const [prioridade, setPrioridade] = useState<"Alta" | "Média" | "Baixa">("Média");
  const [periodo, setPeriodo] = useState<"Diário" | "Semanal" | "Mensal">("Diário");
  const [unidade, setUnidade] = useState("");

  // Subscribe to Checklist in real-time
  useEffect(() => {
    const unsub = onSnapshot(
      collection(db, COLLECTIONS.CHECKLIST),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as ChecklistItem[];

        // Sort: uncompleted first, then newest
        list.sort((a, b) => {
          if (a.concluido !== b.concluido) {
            return a.concluido ? 1 : -1;
          }
          const timeA = a.createdAt?.seconds || 0;
          const timeB = b.createdAt?.seconds || 0;
          return timeB - timeA;
        });

        setItems(list);
        setLoading(false);
      },
      (err) => {
        console.error("Erro ao carregar checklist:", err);
        onToast("Erro ao carregar checklist.", "error");
        setLoading(false);
      }
    );

    return () => unsub();
  }, []);

  const handleAddItem = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!titulo.trim()) {
      onToast("O título do item é obrigatório.", "error");
      return;
    }

    try {
      await addDoc(collection(db, COLLECTIONS.CHECKLIST), {
        titulo: titulo.trim(),
        descricao: descricao.trim(),
        categoria,
        prioridade,
        periodo,
        unidade: unidade.trim(),
        responsavelNome: profile.name,
        responsavelRole: profile.role,
        concluido: false,
        createdAt: serverTimestamp(),
      });

      onToast("Item adicionado ao checklist!", "success");
      setTitulo("");
      setDescricao("");
      setCategoria("Atendimento");
      setPrioridade("Média");
      setPeriodo("Diário");
      setUnidade("");
      setIsAdding(false);
    } catch (err: any) {
      console.error(err);
      onToast(`Erro ao adicionar item: ${err.message}`, "error");
    }
  };

  const handleToggleConcluido = async (item: ChecklistItem) => {
    try {
      const nextStatus = !item.concluido;
      const updateData: any = {
        concluido: nextStatus,
        updatedAt: serverTimestamp(),
      };

      if (nextStatus) {
        updateData.concluidoEm = serverTimestamp();
        updateData.concluidoPorNome = profile.name;
        updateData.concluidoPorUid = profile.uid;
      } else {
        updateData.concluidoEm = null;
        updateData.concluidoPorNome = null;
        updateData.concluidoPorUid = null;
      }

      await updateDoc(doc(db, COLLECTIONS.CHECKLIST, item.id), updateData);
      onToast(nextStatus ? "Item marcado como concluído!" : "Item reaberto.");
    } catch (err: any) {
      onToast(`Erro ao atualizar item: ${err.message}`, "error");
    }
  };

  const handleDeleteItem = async (id: string, itemTitle: string) => {
    if (!window.confirm(`Deseja remover "${itemTitle}" do checklist?`)) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.CHECKLIST, id));
      onToast("Item removido do checklist.", "success");
    } catch (err: any) {
      onToast(`Erro ao remover: ${err.message}`, "error");
    }
  };

  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      if (statusFilter === "Pendentes" && item.concluido) return false;
      if (statusFilter === "Concluidos" && !item.concluido) return false;
      if (categoryFilter !== "Todas" && item.categoria !== categoryFilter) return false;

      if (search.trim()) {
        const term = search.toLowerCase();
        return (
          item.titulo?.toLowerCase().includes(term) ||
          item.descricao?.toLowerCase().includes(term) ||
          item.unidade?.toLowerCase().includes(term)
        );
      }
      return true;
    });
  }, [items, statusFilter, categoryFilter, search]);

  const stats = useMemo(() => {
    const total = items.length;
    const concluidos = items.filter((i) => i.concluido).length;
    const pendentes = total - concluidos;
    const pct = total > 0 ? Math.round((concluidos / total) * 100) : 0;
    return { total, concluidos, pendentes, pct };
  }, [items]);

  const getPriorityColor = (p?: string) => {
    if (p === "Alta") return "bg-rose-50 text-rose-700 border-rose-200";
    if (p === "Média") return "bg-amber-50 text-amber-700 border-amber-200";
    return "bg-slate-100 text-slate-600 border-slate-200";
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto pb-16">
      {/* Header & Progress Card */}
      <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="px-3 py-1 bg-emerald-50 text-emerald-700 text-xs font-extrabold rounded-full border border-emerald-200 flex items-center gap-1">
              <CheckSquare size={13} />
              Operacional Regional
            </span>
          </div>
          <h1 className="text-2xl font-black text-slate-900">Checklist Operacional SM</h1>
          <p className="text-xs text-slate-500 mt-1">
            Acompanhe a realização das rotinas de abertura, atendimento e fechamento da regional em tempo real.
          </p>
        </div>

        <div className="flex items-center gap-4">
          <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 text-center min-w-[130px]">
            <span className="block text-2xl font-black text-emerald-600">{stats.pct}%</span>
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">
              {stats.concluidos}/{stats.total} Concluídos
            </span>
          </div>

          <button
            onClick={() => setIsAdding(true)}
            className="px-5 py-3 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-2xl shadow-md shadow-blue-600/20 transition flex items-center gap-2 cursor-pointer h-fit"
          >
            <Plus size={16} />
            <span>Novo Item</span>
          </button>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="bg-white rounded-3xl p-5 border border-slate-100 shadow-sm flex flex-col md:flex-row items-center justify-between gap-4">
        {/* Status Tabs */}
        <div className="flex items-center gap-1.5 bg-slate-100 p-1.5 rounded-2xl border border-slate-200 w-full sm:w-auto">
          {(["Todos", "Pendentes", "Concluidos"] as const).map((st) => (
            <button
              key={st}
              onClick={() => setStatusFilter(st)}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                statusFilter === st
                  ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <span>{st}</span>
              <span className="px-1.5 py-0.2 bg-slate-200/80 text-[10px] rounded-full">
                {st === "Todos" ? stats.total : st === "Pendentes" ? stats.pendentes : stats.concluidos}
              </span>
            </button>
          ))}
        </div>

        {/* Search & Category Filter */}
        <div className="flex flex-col sm:flex-row items-center gap-3 w-full md:w-auto">
          <div className="relative w-full sm:w-64">
            <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Buscar tarefas..."
              className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>

          <select
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            className="w-full sm:w-auto px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
          >
            <option value="Todas">Todas as Categorias</option>
            <option value="Abertura">Abertura</option>
            <option value="Atendimento">Atendimento</option>
            <option value="Matrículas">Matrículas</option>
            <option value="Auditoria">Auditoria</option>
            <option value="Fechamento">Fechamento</option>
            <option value="Geral">Geral</option>
          </select>
        </div>
      </div>

      {/* Modal Add Checklist Item */}
      {isAdding && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl p-6 md:p-8 max-w-lg w-full shadow-2xl border border-slate-100 space-y-6">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                  <CheckSquare size={20} />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-slate-900">Novo Item do Checklist</h3>
                  <p className="text-xs text-slate-500">Defina a tarefa e parâmetros de acompanhamento</p>
                </div>
              </div>
              <button
                onClick={() => setIsAdding(false)}
                className="p-2 text-slate-400 hover:text-slate-600 rounded-xl cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleAddItem} className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                  Título da Tarefa *
                </label>
                <input
                  type="text"
                  required
                  value={titulo}
                  onChange={(e) => setTitulo(e.target.value)}
                  placeholder="Ex: Vistoria da Sala de Matrícula e Computadores"
                  className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                  Descrição / Instruções
                </label>
                <textarea
                  rows={2}
                  value={descricao}
                  onChange={(e) => setDescricao(e.target.value)}
                  placeholder="Detalhes ou passos para verificação..."
                  className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none resize-none"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                    Categoria
                  </label>
                  <select
                    value={categoria}
                    onChange={(e) => setCategoria(e.target.value as any)}
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                  >
                    <option value="Abertura">Abertura</option>
                    <option value="Atendimento">Atendimento</option>
                    <option value="Matrículas">Matrículas</option>
                    <option value="Auditoria">Auditoria</option>
                    <option value="Fechamento">Fechamento</option>
                    <option value="Geral">Geral</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                    Prioridade
                  </label>
                  <select
                    value={prioridade}
                    onChange={(e) => setPrioridade(e.target.value as any)}
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                  >
                    <option value="Alta">Alta</option>
                    <option value="Média">Média</option>
                    <option value="Baixa">Baixa</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                    Periodicidade
                  </label>
                  <select
                    value={periodo}
                    onChange={(e) => setPeriodo(e.target.value as any)}
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                  >
                    <option value="Diário">Diário</option>
                    <option value="Semanal">Semanal</option>
                    <option value="Mensal">Mensal</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                    Unidade / Polo
                  </label>
                  <input
                    type="text"
                    value={unidade}
                    onChange={(e) => setUnidade(e.target.value)}
                    placeholder="Todas ou nome da unidade"
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
              </div>

              <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsAdding(false)}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition cursor-pointer"
                >
                  Adicionar Tarefa
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Checklist List */}
      <div className="space-y-3">
        {filteredItems.length === 0 ? (
          <div className="text-center py-16 bg-white rounded-3xl border border-dashed border-slate-200 space-y-3">
            <div className="w-16 h-16 bg-emerald-50 text-emerald-500 rounded-full flex items-center justify-center mx-auto">
              <CheckSquare size={32} />
            </div>
            <h3 className="text-base font-bold text-slate-800">Nenhuma tarefa encontrada</h3>
            <p className="text-xs text-slate-500 max-w-md mx-auto">
              {statusFilter === "Concluidos"
                ? "Nenhum item concluído ainda. Marque itens na lista para registrá-los como concluídos."
                : "Clique em 'Novo Item' para adicionar tarefas e procedimentos ao checklist regional."}
            </p>
          </div>
        ) : (
          filteredItems.map((item) => (
            <div
              key={item.id}
              className={`p-5 rounded-3xl border transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-4 ${
                item.concluido
                  ? "bg-slate-50/70 border-slate-200/80 opacity-75"
                  : "bg-white border-slate-100 shadow-sm hover:border-slate-200"
              }`}
            >
              <div className="flex items-start gap-4">
                <button
                  onClick={() => handleToggleConcluido(item)}
                  className={`mt-0.5 w-6 h-6 rounded-lg flex items-center justify-center transition-all cursor-pointer ${
                    item.concluido
                      ? "bg-emerald-600 text-white shadow-xs"
                      : "border-2 border-slate-300 hover:border-blue-600 text-transparent"
                  }`}
                >
                  <CheckCircle2 size={16} className={item.concluido ? "block" : "hidden"} />
                </button>

                <div className="space-y-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <h3
                      className={`font-bold text-sm ${
                        item.concluido ? "line-through text-slate-500" : "text-slate-900"
                      }`}
                    >
                      {item.titulo}
                    </h3>
                    <span
                      className={`px-2 py-0.5 text-[10px] font-black uppercase rounded-md border ${getPriorityColor(
                        item.prioridade
                      )}`}
                    >
                      {item.prioridade || "Média"}
                    </span>
                    {item.categoria && (
                      <span className="px-2 py-0.5 bg-blue-50 text-blue-700 text-[10px] font-bold rounded-md">
                        {item.categoria}
                      </span>
                    )}
                    {item.periodo && (
                      <span className="px-2 py-0.5 bg-purple-50 text-purple-700 text-[10px] font-bold rounded-md">
                        {item.periodo}
                      </span>
                    )}
                  </div>

                  {item.descricao && (
                    <p className="text-xs text-slate-500 max-w-2xl">{item.descricao}</p>
                  )}

                  <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-400 pt-1">
                    {item.unidade && (
                      <span className="flex items-center gap-1 text-slate-600 font-medium">
                        <Building size={12} />
                        <span>{item.unidade}</span>
                      </span>
                    )}
                    {item.concluido && item.concluidoPorNome && (
                      <span className="flex items-center gap-1 text-emerald-700 font-bold">
                        <CheckCircle2 size={12} />
                        <span>Concluído por {item.concluidoPorNome}</span>
                      </span>
                    )}
                  </div>
                </div>
              </div>

              <div className="flex items-center justify-end gap-2 shrink-0">
                <button
                  onClick={() => handleDeleteItem(item.id, item.titulo)}
                  className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-xl transition cursor-pointer"
                  title="Excluir item"
                >
                  <Trash2 size={16} />
                </button>
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

export default CheckListRegionalView;
