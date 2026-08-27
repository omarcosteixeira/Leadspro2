import React, { useState, useEffect } from "react";
import { Lead, UserProfile } from "../types";
import { X, Phone, Mail, Building2, Calendar, MapPin, Tag, Plus, MessageSquare, History, CheckSquare, FileText, Check, Clock } from "lucide-react";
import { cn, formatPhone } from "../lib/utils";
import { doc, updateDoc, collection, addDoc, serverTimestamp, query, where, orderBy, onSnapshot } from "firebase/firestore";
import { db, COLLECTIONS } from "../firebase";

interface LeadProfileDrawerProps {
  lead: Lead;
  onClose: () => void;
  profile: UserProfile;
  onToast: (m: string, t?: "success" | "error") => void;
}

export default function LeadProfileDrawer({ lead, onClose, profile, onToast }: LeadProfileDrawerProps) {
  const [activeTab, setActiveTab] = useState("dados");
  const [formData, setFormData] = useState<Partial<Lead>>(lead);
  const [historyLogs, setHistoryLogs] = useState<any[]>([]);
  const [tasks, setTasks] = useState<any[]>([]);
  const [newTaskTitle, setNewTaskTitle] = useState("");

  useEffect(() => {
    if (activeTab !== 'tarefas') return;
    const q = query(
      collection(db, "crm_tasks"),
      where("leadId", "==", lead.id),
      orderBy("createdAt", "desc")
    );
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setTasks(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() })));
    });
    return () => unsubscribe();
  }, [activeTab, lead.id]);

  const handleAddTask = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTaskTitle.trim()) return;
    
    try {
      await addDoc(collection(db, "crm_tasks"), {
        leadId: lead.id,
        titulo: newTaskTitle.trim(),
        status: "PENDENTE",
        responsavelId: profile.uid,
        prioridade: "MEDIA",
        createdAt: serverTimestamp()
      });
      setNewTaskTitle("");
      
      await addDoc(collection(db, "crm_history"), {
        leadId: lead.id,
        dataHora: serverTimestamp(),
        usuarioId: profile.uid,
        usuarioNome: profile.name,
        acao: "NOVA_TAREFA",
        detalhes: `Criou a tarefa: "${newTaskTitle.trim()}"`,
        origem: "Tarefas"
      });
      
    } catch (e) {
      console.error(e);
      onToast("Erro ao criar tarefa", "error");
    }
  };
  
  const handleToggleTask = async (taskId: string, currentStatus: string) => {
    const newStatus = currentStatus === "PENDENTE" ? "CONCLUIDA" : "PENDENTE";
    await updateDoc(doc(db, "crm_tasks", taskId), {
      status: newStatus
    });
  };


  useEffect(() => {
    if (activeTab !== 'historico') return;
    const q = query(
      collection(db, "crm_history"),
      where("leadId", "==", lead.id),
      orderBy("dataHora", "desc")
    );
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setHistoryLogs(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() })));
    });
    return () => unsubscribe();
  }, [activeTab, lead.id]);


  const tabs = [
    { id: "dados", label: "Dados do Lead", icon: FileText },
    { id: "whatsapp", label: "WhatsApp", icon: MessageSquare },
    { id: "historico", label: "Histórico", icon: History },
    { id: "tarefas", label: "Tarefas", icon: CheckSquare },
  ];

  const handleSave = async () => {
    try {
      await updateDoc(doc(db, COLLECTIONS.LEADS, lead.id), {
        ...formData,
        ultimaInteracao: serverTimestamp()
      });
      
      await addDoc(collection(db, "crm_history"), {
        leadId: lead.id,
        dataHora: serverTimestamp(),
        usuarioId: profile.uid,
        usuarioNome: profile.name,
        acao: "ATUALIZACAO_PERFIL",
        detalhes: "Atualizou os dados do perfil do lead",
        origem: "Perfil"
      });

      onToast("Lead atualizado com sucesso!", "success");
    } catch (error) {
      console.error(error);
      onToast("Erro ao atualizar lead", "error");
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex justify-end bg-slate-900/50 backdrop-blur-sm">
      <div className="w-full max-w-4xl bg-white h-full shadow-2xl flex flex-col animate-in slide-in-from-right duration-300">
        
        {/* Header */}
        <div className="p-6 border-b border-slate-100 flex items-start justify-between bg-slate-50/50">
          <div className="flex gap-4">
            <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-2xl flex items-center justify-center font-bold text-2xl shadow-sm">
              {lead.nome.charAt(0).toUpperCase()}
            </div>
            <div>
              <h2 className="text-2xl font-bold text-slate-800">{lead.nome}</h2>
              <div className="flex items-center gap-3 mt-2 text-sm text-slate-500">
                <span className="flex items-center gap-1"><Phone size={14} /> {formatPhone(lead.telefone)}</span>
                {lead.email && <span className="flex items-center gap-1"><Mail size={14} /> {lead.email}</span>}
              </div>
              <div className="flex items-center gap-2 mt-3">
                <span className={cn(
                  "px-2.5 py-1 rounded-md text-[10px] font-bold tracking-wider uppercase",
                  formData.temperatura === 'QUENTE' ? "bg-red-100 text-red-600" :
                  formData.temperatura === 'MORNO' ? "bg-yellow-100 text-yellow-600" :
                  "bg-blue-100 text-blue-600"
                )}>
                  {formData.temperatura || 'FRIO'}
                </span>
                <span className="px-2.5 py-1 rounded-md text-[10px] font-bold tracking-wider uppercase bg-slate-100 text-slate-600">
                  {formData.etapaFunil || 'NOVO LEAD'}
                </span>
                <span className="px-2.5 py-1 rounded-md text-[10px] font-bold tracking-wider uppercase bg-emerald-100 text-emerald-600">
                  Score: {formData.score || 0}
                </span>
              </div>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-200 rounded-full transition-colors text-slate-400">
            <X size={20} />
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-slate-200 bg-white px-6">
          {tabs.map(t => (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={cn(
                "px-6 py-4 text-sm font-bold border-b-2 transition-colors flex items-center gap-2",
                activeTab === t.id 
                  ? "border-blue-600 text-blue-600" 
                  : "border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300"
              )}
            >
              <t.icon size={16} />
              {t.label}
            </button>
          ))}
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto bg-slate-50 p-6 custom-scrollbar">
          {activeTab === 'dados' && (
            <div className="space-y-6 max-w-3xl">
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 space-y-4">
                <h3 className="font-bold text-slate-800 text-lg mb-4">Informações Principais</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Nome</label>
                    <input 
                      type="text" 
                      value={formData.nome || ""} 
                      onChange={e => setFormData({...formData, nome: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Telefone</label>
                    <input 
                      type="text" 
                      value={formData.telefone || ""} 
                      onChange={e => setFormData({...formData, telefone: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">E-mail</label>
                    <input 
                      type="email" 
                      value={formData.email || ""} 
                      onChange={e => setFormData({...formData, email: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">CPF</label>
                    <input 
                      type="text" 
                      value={formData.cpf || ""} 
                      onChange={e => setFormData({...formData, cpf: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Curso de Interesse</label>
                    <input 
                      type="text" 
                      value={formData.cursoInteresse || ""} 
                      onChange={e => setFormData({...formData, cursoInteresse: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Modalidade</label>
                    <select 
                      value={formData.modalidade || ""} 
                      onChange={e => setFormData({...formData, modalidade: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    >
                      <option value="">Selecione...</option>
                      <option value="EAD">EAD</option>
                      <option value="PRESENCIAL">PRESENCIAL</option>
                      <option value="FLEX">FLEX</option>
                    </select>
                  </div>
                </div>
              </div>

              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 space-y-4">
                <h3 className="font-bold text-slate-800 text-lg mb-4">Classificação e Funil</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Etapa do Funil</label>
                    <select 
                      value={formData.etapaFunil || "NOVO LEAD"} 
                      onChange={e => setFormData({...formData, etapaFunil: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    >
                      {["NOVO LEAD", "CONTATO INICIADO", "EM ATENDIMENTO", "INTERESSADO", "PROPOSTA ENVIADA", "FOLLOW-UP", "NEGOCIAÇÃO", "AGUARDANDO PAGAMENTO", "CONVERTIDO", "PERDIDO"].map(e => (
                        <option key={e} value={e}>{e}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">Temperatura</label>
                    <select 
                      value={formData.temperatura || "FRIO"} 
                      onChange={e => setFormData({...formData, temperatura: e.target.value as any})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    >
                      <option value="FRIO">Frio</option>
                      <option value="MORNO">Morno</option>
                      <option value="QUENTE">Quente</option>
                    </select>
                  </div>
                  <div className="space-y-1 col-span-2">
                    <label className="text-xs font-bold text-slate-500 uppercase">Origem/Campanha</label>
                    <input 
                      type="text" 
                      value={formData.origem || formData.acao || ""} 
                      onChange={e => setFormData({...formData, origem: e.target.value})}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div className="space-y-1 col-span-2">
                    <label className="text-xs font-bold text-slate-500 uppercase">Observações Gerais</label>
                    <textarea 
                      value={formData.observacoes || ""} 
                      onChange={e => setFormData({...formData, observacoes: e.target.value})}
                      rows={4}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none resize-none"
                    />
                  </div>
                </div>
              </div>
              
              <div className="flex justify-end pt-4 pb-12">
                <button onClick={handleSave} className="bg-blue-600 text-white px-8 py-3 rounded-xl font-bold hover:bg-blue-700 transition-colors shadow-lg shadow-blue-500/20">
                  Salvar Alterações
                </button>
              </div>
            </div>
          )}

          {activeTab === 'whatsapp' && (
            <div className="flex flex-col h-full bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
               <div className="flex-1 p-12 text-center flex flex-col items-center justify-center space-y-4">
                  <MessageSquare size={48} className="text-slate-300" />
                  <p className="text-slate-500">Módulo de WhatsApp será integrado aqui (Caixa de Entrada unificada).</p>
               </div>
            </div>
          )}

          {activeTab === 'historico' && (
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-0 overflow-hidden">
              <div className="p-4 border-b border-slate-100 bg-slate-50">
                <h3 className="font-bold text-slate-800">Histórico de Ações</h3>
              </div>
              <div className="p-0">
                {historyLogs.length > 0 ? (
                  <div className="divide-y divide-slate-100">
                    {historyLogs.map(log => (
                      <div key={log.id} className="p-4 flex gap-4 hover:bg-slate-50/50 transition-colors">
                        <div className="mt-1">
                          <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-xs">
                            {log.usuarioNome?.charAt(0) || '?'}
                          </div>
                        </div>
                        <div className="flex-1 space-y-1">
                          <p className="text-sm text-slate-800">
                            <span className="font-bold">{log.usuarioNome}</span> {log.detalhes}
                          </p>
                          <div className="flex items-center gap-2 text-xs text-slate-400">
                            <span className="font-bold uppercase">{log.acao}</span>
                            <span>•</span>
                            <span>{log.dataHora?.seconds ? new Date(log.dataHora.seconds * 1000).toLocaleString() : 'Recente'}</span>
                            <span>•</span>
                            <span>{log.origem}</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="p-12 text-center text-slate-400">
                    <History size={32} className="mx-auto mb-2 opacity-50" />
                    <p>Nenhum histórico registrado.</p>
                  </div>
                )}
              </div>
            </div>
          )}

          {activeTab === 'tarefas' && (
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[500px]">
              <div className="p-4 border-b border-slate-100 bg-slate-50">
                <h3 className="font-bold text-slate-800">Tarefas e Lembretes</h3>
              </div>
              
              <div className="flex-1 overflow-y-auto p-4 space-y-2">
                {tasks.length > 0 ? (
                  tasks.map(task => (
                    <div key={task.id} className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-100 rounded-xl hover:border-slate-200 transition-colors">
                      <button 
                        onClick={() => handleToggleTask(task.id, task.status)}
                        className={cn(
                          "w-6 h-6 rounded-md flex items-center justify-center border transition-colors shrink-0",
                          task.status === 'CONCLUIDA' ? "bg-emerald-500 border-emerald-500 text-white" : "border-slate-300 text-transparent hover:border-blue-500"
                        )}
                      >
                        <Check size={14} />
                      </button>
                      <div className="flex-1">
                        <p className={cn("text-sm font-medium", task.status === 'CONCLUIDA' ? "text-slate-400 line-through" : "text-slate-700")}>
                          {task.titulo}
                        </p>
                        <div className="flex items-center gap-2 mt-1">
                          <span className={cn(
                            "text-[10px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wide",
                            task.status === 'PENDENTE' ? "bg-amber-100 text-amber-700" : "bg-slate-200 text-slate-500"
                          )}>
                            {task.status}
                          </span>
                        </div>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-slate-400">
                    <CheckSquare size={32} className="mb-2 opacity-50" />
                    <p>Nenhuma tarefa pendente.</p>
                  </div>
                )}
              </div>
              
              <form onSubmit={handleAddTask} className="p-4 border-t border-slate-100 bg-white flex gap-2">
                <input
                  type="text"
                  placeholder="Adicionar nova tarefa..."
                  value={newTaskTitle}
                  onChange={e => setNewTaskTitle(e.target.value)}
                  className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <button 
                  type="submit"
                  disabled={!newTaskTitle.trim()}
                  className="bg-blue-600 text-white px-4 py-2 rounded-xl font-bold hover:bg-blue-700 disabled:opacity-50 transition-colors"
                >
                  <Plus size={20} />
                </button>
              </form>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
