import React, { useState, useMemo, useEffect } from "react";
import { SolicitacaoManutencao, UserProfile } from "../types";
import { db, COLLECTIONS, handleFirestoreError, OperationType } from "../firebase";
import {
  collection,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  onSnapshot,
  serverTimestamp,
} from "firebase/firestore";
import {
  Plus,
  Trash2,
  Check,
  X,
  ClipboardList,
  AlertTriangle,
  CheckCircle2,
  Send,
  User,
  MapPin,
  Clock,
  Wrench,
  Search,
} from "lucide-react";


interface Props {
  profile: UserProfile | null;
  onToast: (m: string, t?: "success" | "error") => void;
}

export function SolicitacoesManutencaoView({ profile, onToast }: Props) {
  const [solicitacoes, setSolicitacoes] = useState<SolicitacaoManutencao[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [filterStatus, setFilterStatus] = useState<string>("all");
  const [isAdding, setIsAdding] = useState(false);
  const [newDescricao, setNewDescricao] = useState("");
  const [newLocal, setNewLocal] = useState("");
  const [observacoes, setObservacoes] = useState<{ [key: string]: string }>({});

  useEffect(() => {
    const unsub = onSnapshot(
      collection(db, COLLECTIONS.SOLICITACOES_MANUTENCAO),
      (snap) => {
        const data = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as SolicitacaoManutencao[];
        // Sort by createdAt descending
        data.sort((a, b) => {
          const tA = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
          const tB = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
          return tB - tA;
        });
        setSolicitacoes(data);
      },
      (err) => {
        handleFirestoreError(err, OperationType.LIST, "solicitações de manutenção");
        onToast("Erro ao carregar solicitações de manutenção", "error");
      }
    );
    return () => unsub();
  }, [onToast]);

  const isFinanceiro = useMemo(() => {
    if (!profile) return false;
    return profile.role === "Financeiro" || profile.role === "Admin Master";
  }, [profile]);

  const filteredSolicitacoes = useMemo(() => {
    return solicitacoes.filter((s) => {
      const matchStatus = filterStatus === "all" || s.status === filterStatus;
      const matchSearch =
        s.descricao.toLowerCase().includes(searchTerm.toLowerCase()) ||
        s.local.toLowerCase().includes(searchTerm.toLowerCase()) ||
        s.solicitanteNome.toLowerCase().includes(searchTerm.toLowerCase());
      return matchStatus && matchSearch;
    });
  }, [solicitacoes, filterStatus, searchTerm]);

  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!profile) return;
    if (!newDescricao.trim() || !newLocal.trim()) {
      onToast("Preencha todos os campos obrigatórios.", "error");
      return;
    }

    try {
      await addDoc(collection(db, COLLECTIONS.SOLICITACOES_MANUTENCAO), {
        descricao: newDescricao.trim(),
        local: newLocal.trim(),
        status: "Pendente",
        solicitanteId: profile.uid,
        solicitanteNome: profile.name,
        createdAt: serverTimestamp(),
      });
      setIsAdding(false);
      setNewDescricao("");
      setNewLocal("");
      onToast("Solicitação enviada com sucesso!", "success");
    } catch (err: any) {
      handleFirestoreError(err, OperationType.CREATE, "solicitação de manutenção");
      onToast(`Erro ao criar solicitação: ${err.message}`, "error");
    }
  };

  const handleUpdateStatus = async (
    id: string,
    newStatus: SolicitacaoManutencao["status"]
  ) => {
    if (!isFinanceiro) return;
    try {
      const obs = observacoes[id]?.trim();
      const updateData: any = {
        status: newStatus,
        updatedAt: serverTimestamp(),
      };
      if (obs) {
        updateData.observacoesFinanceiro = obs;
      }
      
      await updateDoc(doc(db, COLLECTIONS.SOLICITACOES_MANUTENCAO, id), updateData);
      onToast(`Status atualizado para ${newStatus}`, "success");
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, "solicitação de manutenção");
      onToast(`Erro ao atualizar status: ${err.message}`, "error");
    }
  };

  const handleDelete = async (id: string, solicitanteId: string) => {
    if (!profile) return;
    // Only financeiro or the person who created it can delete
    if (!isFinanceiro && profile.uid !== solicitanteId) {
      onToast("Sem permissão para excluir", "error");
      return;
    }
    if (window.confirm("Deseja realmente excluir esta solicitação?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.SOLICITACOES_MANUTENCAO, id));
        onToast("Solicitação excluída com sucesso!");
      } catch (err: any) {
        handleFirestoreError(err, OperationType.DELETE, "solicitação de manutenção");
        onToast(`Erro ao excluir solicitação: ${err.message}`, "error");
      }
    }
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case "Aprovado":
        return "bg-green-100 text-green-700 border-green-200";
      case "Rejeitado":
        return "bg-rose-100 text-rose-700 border-rose-200";
      case "Em Andamento":
        return "bg-amber-100 text-amber-700 border-amber-200";
      case "Concluído":
        return "bg-blue-100 text-blue-700 border-blue-200";
      case "Pendente":
      default:
        return "bg-slate-100 text-slate-700 border-slate-200";
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <Wrench className="text-blue-600" size={28} />
            Solicitações de Manutenção
          </h2>
          <p className="text-sm text-slate-500 font-medium mt-1">
            Solicite pequenas manutenções para o seu setor.
          </p>
        </div>
        <button
          onClick={() => setIsAdding(true)}
          className="flex items-center justify-center space-x-2 bg-blue-600 hover:bg-blue-700 text-white font-bold px-5 py-3 rounded-xl text-sm transition-all shadow-md"
        >
          <Plus size={18} />
          <span>Nova Solicitação</span>
        </button>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-4">
        <div className="flex flex-col md:flex-row gap-4 items-center">
          <div className="flex-1 w-full relative">
            <Search
              size={18}
              className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Buscar por descrição, local ou solicitante..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm bg-slate-50 focus:bg-white transition-colors"
            />
          </div>
          <select
            value={filterStatus}
            onChange={(e) => setFilterStatus(e.target.value)}
            className="px-4 py-2.5 rounded-xl border border-slate-200 text-sm font-medium focus:ring-2 focus:ring-blue-500 outline-none w-full md:w-auto bg-slate-50 focus:bg-white cursor-pointer"
          >
            <option value="all">Todos os Status</option>
            <option value="Pendente">Pendente</option>
            <option value="Aprovado">Aprovado</option>
            <option value="Rejeitado">Rejeitado</option>
            <option value="Em Andamento">Em Andamento</option>
            <option value="Concluído">Concluído</option>
          </select>
        </div>
      </div>

      {isAdding && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-lg w-full max-h-[90vh] flex flex-col overflow-hidden">
            <div className="p-6 border-b border-slate-100 flex items-center justify-between shrink-0 bg-slate-50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-blue-100 flex items-center justify-center text-blue-600">
                  <Wrench size={20} />
                </div>
                <div>
                  <h3 className="text-xl font-bold text-slate-800">
                    Nova Solicitação
                  </h3>
                  <p className="text-xs text-slate-500 font-medium">
                    Detalhe o problema e o local
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsAdding(false)}
                className="p-2 text-slate-400 hover:bg-slate-200 hover:text-slate-600 rounded-full transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-6 overflow-y-auto">
              <form id="manutencao-form" onSubmit={handleAdd} className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1 ml-1">
                    Local / Setor *
                  </label>
                  <div className="relative">
                    <MapPin
                      size={16}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                    />
                    <input
                      type="text"
                      required
                      placeholder="Ex: Sala 2, Recepção, Banheiro Masculino..."
                      value={newLocal}
                      onChange={(e) => setNewLocal(e.target.value)}
                      className="w-full pl-10 pr-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm transition-all"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 uppercase mb-1 ml-1">
                    Descrição do Problema *
                  </label>
                  <textarea
                    required
                    placeholder="Descreva o que precisa ser consertado ou verificado..."
                    value={newDescricao}
                    onChange={(e) => setNewDescricao(e.target.value)}
                    rows={4}
                    className="w-full p-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm transition-all resize-none"
                  />
                </div>
              </form>
            </div>

            <div className="p-6 border-t border-slate-100 bg-slate-50 shrink-0">
              <div className="flex gap-3">
                <button
                  type="button"
                  onClick={() => setIsAdding(false)}
                  className="flex-1 px-4 py-3 rounded-xl font-bold text-slate-600 bg-white border border-slate-200 hover:bg-slate-50 transition-colors"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  form="manutencao-form"
                  className="flex-1 px-4 py-3 rounded-xl font-bold text-white bg-blue-600 hover:bg-blue-700 transition-colors flex items-center justify-center gap-2"
                >
                  <Send size={18} />
                  <span>Enviar Solicitação</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {filteredSolicitacoes.map((s) => (
          <div
            key={s.id}
            className="bg-white rounded-2xl shadow-sm border border-slate-200 flex flex-col overflow-hidden hover:shadow-md transition-shadow group"
          >
            <div className="p-5 border-b border-slate-100 flex items-start justify-between bg-slate-50/50">
              <div className="flex-1 pr-4">
                <div className="flex items-center gap-2 mb-2">
                  <span
                    className={`px-2.5 py-1 rounded-md text-[10px] font-black uppercase tracking-wider border ${getStatusColor(
                      s.status
                    )}`}
                  >
                    {s.status}
                  </span>
                </div>
                <h4 className="text-base font-bold text-slate-800 line-clamp-2">
                  {s.descricao}
                </h4>
              </div>
              {(isFinanceiro || (profile && profile.uid === s.solicitanteId && s.status === 'Pendente')) && (
                <button
                  onClick={() => handleDelete(s.id, s.solicitanteId)}
                  className="p-2 text-slate-400 hover:bg-rose-100 hover:text-rose-600 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                  title="Excluir"
                >
                  <Trash2 size={16} />
                </button>
              )}
            </div>

            <div className="p-5 space-y-4 flex-1">
              <div className="grid grid-cols-1 gap-3">
                <div className="flex items-center gap-2 text-sm text-slate-600 bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                  <MapPin size={16} className="text-slate-400 shrink-0" />
                  <span className="font-medium truncate">{s.local}</span>
                </div>
                <div className="flex items-center gap-2 text-sm text-slate-600 bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                  <User size={16} className="text-slate-400 shrink-0" />
                  <span className="font-medium truncate">{s.solicitanteNome}</span>
                </div>
                <div className="flex items-center gap-2 text-sm text-slate-600 bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                  <Clock size={16} className="text-slate-400 shrink-0" />
                  <span className="font-medium">
                    {s.createdAt?.toDate
                      ? s.createdAt.toDate().toLocaleDateString("pt-BR")
                      : "Recente"}
                  </span>
                </div>
              </div>

              {s.observacoesFinanceiro && (
                <div className="mt-4 p-3 rounded-xl bg-amber-50 border border-amber-100">
                  <div className="flex items-center gap-1.5 mb-1 text-amber-700">
                    <ClipboardList size={14} />
                    <span className="text-[10px] font-bold uppercase tracking-wider">
                      Obs. Financeiro
                    </span>
                  </div>
                  <p className="text-sm text-amber-900 font-medium">
                    {s.observacoesFinanceiro}
                  </p>
                </div>
              )}
            </div>

            {isFinanceiro && (
              <div className="p-5 border-t border-slate-100 bg-slate-50 space-y-3">
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-bold text-slate-700 ml-1">Observações (opcional)</label>
                  <input
                    type="text"
                    placeholder="Ex: Peça comprada, agendado para sexta..."
                    value={observacoes[s.id] ?? s.observacoesFinanceiro ?? ""}
                    onChange={(e) => setObservacoes(prev => ({ ...prev, [s.id]: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>
                
                {s.status === "Pendente" && (
                  <div className="flex gap-2">
                    <button
                      onClick={() => handleUpdateStatus(s.id, "Aprovado")}
                      className="flex-1 bg-green-100 hover:bg-green-200 text-green-700 py-2 rounded-lg text-xs font-bold transition-colors flex items-center justify-center gap-1"
                    >
                      <Check size={14} /> Aprovar
                    </button>
                    <button
                      onClick={() => handleUpdateStatus(s.id, "Rejeitado")}
                      className="flex-1 bg-rose-100 hover:bg-rose-200 text-rose-700 py-2 rounded-lg text-xs font-bold transition-colors flex items-center justify-center gap-1"
                    >
                      <X size={14} /> Rejeitar
                    </button>
                  </div>
                )}
                
                {s.status === "Aprovado" && (
                  <button
                    onClick={() => handleUpdateStatus(s.id, "Em Andamento")}
                    className="w-full bg-amber-100 hover:bg-amber-200 text-amber-700 py-2 rounded-lg text-xs font-bold transition-colors flex items-center justify-center gap-1"
                  >
                    <AlertTriangle size={14} /> Marcar Em Andamento
                  </button>
                )}

                {s.status === "Em Andamento" && (
                  <button
                    onClick={() => handleUpdateStatus(s.id, "Concluído")}
                    className="w-full bg-blue-100 hover:bg-blue-200 text-blue-700 py-2 rounded-lg text-xs font-bold transition-colors flex items-center justify-center gap-1"
                  >
                    <CheckCircle2 size={14} /> Concluir Manutenção
                  </button>
                )}
                
                {(s.status === "Concluído" || s.status === "Rejeitado") && (
                   <button
                    onClick={() => handleUpdateStatus(s.id, "Pendente")}
                    className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg text-xs font-bold transition-colors"
                  >
                     Voltar para Pendente
                  </button>
                )}
              </div>
            )}
          </div>
        ))}
        {filteredSolicitacoes.length === 0 && (
          <div className="col-span-full py-16 flex flex-col items-center justify-center bg-white rounded-2xl border border-dashed border-slate-300">
            <div className="w-16 h-16 rounded-full bg-slate-50 flex items-center justify-center text-slate-400 mb-4">
              <Wrench size={32} />
            </div>
            <p className="text-slate-500 font-medium text-center max-w-sm">
              Nenhuma solicitação encontrada para os filtros atuais.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
