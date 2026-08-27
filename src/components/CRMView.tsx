import React, { useState } from "react";
import { Lead, BaseEntry, FiesProuniEntry, GapEntry, UserProfile } from "../types";
import { Kanban, MessageSquare, Plus } from "lucide-react";
import CRMKanbanView from "./CRMKanbanView";
import CRMWhatsAppView from "./CRMWhatsAppView";
import LeadProfileDrawer from "./LeadProfileDrawer";
import { cn } from "../lib/utils";

interface CRMViewProps {
  leads: Lead[];
  bases: BaseEntry[];
  fiesProuni: FiesProuniEntry[];
  gap: GapEntry[];
  profile: UserProfile;
  onSendBot: (tel: string, msg: string, contactName?: string) => Promise<void>;
  onToast: (m: string, t?: "success" | "error") => void;
}

export default function CRMView({
  leads,
  bases,
  fiesProuni,
  gap,
  profile,
  onSendBot,
  onToast
}: CRMViewProps) {
  const [activeTab, setActiveTab] = useState<"kanban" | "whatsapp">("kanban");
  const [selectedLead, setSelectedLead] = useState<Lead | null>(null);

  return (
    <div className="flex flex-col h-full bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden relative">
      {/* CRM Header & Sub-nav */}
      <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="bg-blue-600 text-white p-2 rounded-xl">
            <Kanban size={20} />
          </div>
          <h2 className="text-xl font-black text-slate-800 tracking-tight">CRM & Atendimento</h2>
        </div>
        
        <div className="flex bg-slate-200/50 p-1 rounded-xl">
          <button
            onClick={() => setActiveTab("kanban")}
            className={cn(
              "flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all",
              activeTab === "kanban" ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
            )}
          >
            <Kanban size={16} />
            Funil (Kanban)
          </button>
          <button
            onClick={() => setActiveTab("whatsapp")}
            className={cn(
              "flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all",
              activeTab === "whatsapp" ? "bg-white text-emerald-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
            )}
          >
            <MessageSquare size={16} />
            WhatsApp (Inbox)
          </button>
        </div>
      </div>

      {/* Main Content Area */}
      <div className="flex-1 overflow-hidden">
        {activeTab === "kanban" && (
          <CRMKanbanView 
            leads={leads} 
            profile={profile} 
            onLeadClick={(lead) => setSelectedLead(lead)} 
            onToast={onToast} 
          />
        )}
        {activeTab === "whatsapp" && (
          <CRMWhatsAppView 
            leads={leads}
            bases={bases}
            fiesProuni={fiesProuni}
            gap={gap}
            profile={profile}
            onSendBot={onSendBot}
            onToast={onToast}
          />
        )}
      </div>

      {/* Profile Drawer Overlay */}
      {selectedLead && (
        <LeadProfileDrawer 
          lead={selectedLead} 
          onClose={() => setSelectedLead(null)} 
          profile={profile}
          onToast={onToast}
        />
      )}
    </div>
  );
}
