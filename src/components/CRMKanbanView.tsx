import React, { useState, useMemo } from "react";
import { Lead, UserProfile } from "../types";
import { cn, formatPhone } from "../lib/utils";
import { Search, User, Phone, Mail, Clock, Calendar, MessageSquare, MoreVertical, Building2 } from "lucide-react";
import { doc, updateDoc, collection, addDoc, serverTimestamp } from "firebase/firestore";
import { db, COLLECTIONS } from "../firebase";

const KANBAN_STAGES = [
  "NOVO LEAD",
  "CONTATO INICIADO",
  "EM ATENDIMENTO",
  "INTERESSADO",
  "PROPOSTA ENVIADA",
  "FOLLOW-UP",
  "NEGOCIAÇÃO",
  "AGUARDANDO PAGAMENTO",
  "CONVERTIDO",
  "PERDIDO"
];

interface CRMKanbanViewProps {
  leads: Lead[];
  profile: UserProfile;
  onLeadClick: (lead: Lead) => void;
  onToast: (m: string, t?: "success" | "error") => void;
}

export default function CRMKanbanView({ leads, profile, onLeadClick, onToast }: CRMKanbanViewProps) {
  const [searchTerm, setSearchTerm] = useState("");

  const filteredLeads = useMemo(() => {
    return leads.filter(l => 
      l.nome?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      l.telefone?.includes(searchTerm) ||
      l.cursoInteresse?.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [leads, searchTerm]);

  // Group leads by stage
  const columns = useMemo(() => {
    const cols: Record<string, Lead[]> = {};
    KANBAN_STAGES.forEach(s => cols[s] = []);
    
    filteredLeads.forEach(lead => {
      let stage = lead.etapaFunil || "NOVO LEAD";
      if (!KANBAN_STAGES.includes(stage)) {
        if (lead.status === 'Convertido') stage = 'CONVERTIDO';
        else if (lead.status === 'Não Interessado' || lead.status === 'Sem retorno') stage = 'PERDIDO';
        else if (lead.status === 'Interessado') stage = 'INTERESSADO';
        else stage = "NOVO LEAD";
      }
      cols[stage].push(lead);
    });
    return cols;
  }, [filteredLeads]);

  const handleDragStart = (e: React.DragEvent, leadId: string) => {
    e.dataTransfer.setData("leadId", leadId);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = async (e: React.DragEvent, targetStage: string) => {
    e.preventDefault();
    const leadId = e.dataTransfer.getData("leadId");
    if (!leadId) return;

    try {
      const leadRef = doc(db, COLLECTIONS.LEADS, leadId);
      await updateDoc(leadRef, {
        etapaFunil: targetStage,
        ultimaInteracao: serverTimestamp()
      });
      
      // Criar log de histórico
      await addDoc(collection(db, "crm_history"), {
        leadId,
        dataHora: serverTimestamp(),
        usuarioId: profile.uid,
        usuarioNome: profile.name,
        acao: "ALTEROU_ETAPA",
        detalhes: `Moveu o lead para a etapa ${targetStage}`,
        origem: "Kanban"
      });

      onToast("Lead movido com sucesso!", "success");
    } catch (error) {
      console.error(error);
      onToast("Erro ao mover lead", "error");
    }
  };

  return (
    <div className="flex flex-col h-full bg-slate-50">
      <div className="p-4 border-b border-slate-200 bg-white flex items-center justify-between">
        <h2 className="text-lg font-bold text-slate-800">Funil de Vendas</h2>
        <div className="relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input
            type="text"
            placeholder="Pesquisar leads..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="pl-9 pr-4 py-2 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-64"
          />
        </div>
      </div>

      <div className="flex-1 overflow-x-auto overflow-y-hidden p-6 custom-scrollbar">
        <div className="flex gap-6 h-full items-start" style={{ width: 'max-content' }}>
          {KANBAN_STAGES.map(stage => (
            <div 
              key={stage}
              className="flex flex-col w-80 max-h-full bg-slate-100/50 rounded-2xl border border-slate-200/60"
              onDragOver={handleDragOver}
              onDrop={(e) => handleDrop(e, stage)}
            >
              <div className="p-4 border-b border-slate-200 flex items-center justify-between bg-white rounded-t-2xl shadow-sm">
                <h3 className="font-bold text-slate-700 text-sm">{stage}</h3>
                <span className="bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full text-xs font-bold">
                  {columns[stage].length}
                </span>
              </div>
              <div className="flex-1 overflow-y-auto p-3 space-y-3 custom-scrollbar">
                {columns[stage].map(lead => (
                  <div 
                    key={lead.id}
                    draggable
                    onDragStart={(e) => handleDragStart(e, lead.id)}
                    onClick={() => onLeadClick(lead)}
                    className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 hover:border-blue-300 hover:shadow-md transition-all cursor-pointer group relative"
                  >
                    <div className="flex justify-between items-start mb-2">
                      <p className="font-bold text-slate-800 text-sm truncate pr-6">{lead.nome}</p>
                    </div>
                    <div className="space-y-1.5 mb-3">
                      <div className="flex items-center text-xs text-slate-500 gap-2">
                        <Phone size={12} />
                        <span>{formatPhone(lead.telefone)}</span>
                      </div>
                      {lead.cursoInteresse && (
                        <div className="flex items-center text-xs text-slate-500 gap-2">
                          <Building2 size={12} />
                          <span className="truncate">{lead.cursoInteresse}</span>
                        </div>
                      )}
                    </div>
                    <div className="flex items-center justify-between mt-3 pt-3 border-t border-slate-100">
                      <div className="flex items-center gap-1">
                        <div className={cn(
                          "w-2 h-2 rounded-full",
                          lead.temperatura === 'QUENTE' ? "bg-red-500" :
                          lead.temperatura === 'MORNO' ? "bg-yellow-500" :
                          "bg-blue-400"
                        )} />
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                          {lead.temperatura || 'FRIO'}
                        </span>
                      </div>
                      <span className="text-[10px] text-slate-400 bg-slate-50 px-2 py-1 rounded-md">
                        {lead.origem || lead.acao || 'Desconhecida'}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
