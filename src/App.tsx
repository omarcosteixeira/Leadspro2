import React, { useState, useEffect, useMemo } from "react";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { initializeApp, getApp } from "firebase/app";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  sendPasswordResetEmail,
  updatePassword,
  updateProfile,
  getAuth,
  User,
} from "firebase/auth";
import {
  collection,
  query,
  onSnapshot,
  addDoc,
  updateDoc,
  doc,
  deleteDoc,
  serverTimestamp,
  where,
  or,
  limit,
  orderBy,
  getDoc,
  setDoc,
  getDocs,
  writeBatch,
} from "firebase/firestore";
import {
  LayoutDashboard,
  UserPlus,
  History,
  Database,
  GraduationCap,
  Settings,
  LogOut,
  Plus,
  Trash2,
  ExternalLink,
  CheckCircle2,
  XCircle,
  Phone,
  Search,
  Users,
  User as UserIcon,
  TrendingUp,
  Calendar,
  Download,
  Upload,
  Menu,
  X,
  Check,
  ChevronRight,
  AlertCircle,
  FileText,
  Clock,
  Calculator,
  LayoutGrid,
  List,
  ShieldCheck,
  Megaphone,
  Sun,
  Edit2,
  Share2,
  Edit,
  Save,
  MapPin,
  Lock,
  Unlock,
  Circle,
  KeyRound,
  Building2,
  MessageSquare,
  PhoneOutgoing,
  Mail,
  Globe,
  Copy,
  Bot,
  Send,
  Bell,
  Monitor,
  Maximize,
  Cloud,
  RefreshCw,
  Play,
  Pause,
  ChevronUp,
  ChevronDown,
  Target,
  Cake,
  CheckSquare,
  Square,
  Coins,
  BookOpen,
  Briefcase,
  Boxes,
  Smartphone,
  Chrome,
  BarChart3,
  Eye,
  EyeOff,
  UserMinus,
  Wrench,
  Hash,
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import {
  auth,
  db,
  COLLECTIONS,
  handleFirestoreError,
  OperationType,
  secondaryAuth,
  firebaseConfigPrincipal,
  firebaseConfigComercial,
} from "./firebase";
import {
  cn,
  formatPhone,
  getWhatsAppUrl,
  validateCPF,
  formatCPF,
} from "./lib/utils";
import * as XLSX from "xlsx";
import { EmailMarketingView } from "./components/EmailMarketingView";
import { RelatoriosView } from "./components/RelatoriosView";
import { ControleConcorrenciaView } from "./components/ControleConcorrenciaView";
import Mapa3D from "./components/Mapa3D";
import {
  UserProfile,
  SalesContact,
  Lead,
  BaseEntry,
  GapEntry,
  PlannerTask,
  LinkUtil,
  UserRole,
  FiesProuniEntry,
  FiesProuniVaga,
  Campanha,
  BomDiaCaptacao,
  ForecastCaptacao,
  BomDiaMetrics,
  PeriodoCaptacao,
  CalendarioAcao,
  EmpresaParceira,
  WhatsAppMessage,
  MapaoAcademicoEntry,
  BaseDisparoEntry,
  BotConfig,
  MetaDia,
  MetaSM,
  MetaCurso,
  QgLigacao,
  SolicitacaoFolga,
  CursoDisponivel,
  InsumoPedido,
  InsumoEstoque,
  InsumoBaixa,
  InsumoPedidoComercial,
  InsumoEstoqueComercial,
  IsencaoEntry,
  ControleConcorrencia,
  PedidoCursoEntry,
  Ligacao,
  AnalysisScheme,
  PeriodAnalysis,
  SolicitacaoManutencao,
  MensagemEnviadaLog
} from "./types";
import { OPENROUTER_MODELS } from "./ai-config";
import CrescimentoAnualAdmin from "./components/CrescimentoAnualAdmin";
import { ProfileModal } from "./components/ProfileModal";
import { PublicRegistrationForm } from "./components/PublicRegistrationForm";
import { FormulariosView } from "./components/FormulariosView";
import { PublicCustomForm } from "./components/PublicCustomForm";
import { PublicInsumoForm } from "./components/PublicInsumoForm";
import { PublicMaintenanceForm } from "./components/PublicMaintenanceForm";
import { PublicPedidoCursoForm } from "./components/PublicPedidoCursoForm";
import { MessageTemplateModal } from "./components/MessageTemplateModal";
import { CursosDisponiveisView } from "./components/CursosDisponiveisView";
import { ControleInsumosView } from "./components/ControleInsumosView";
import { SolicitacoesManutencaoView } from "./components/SolicitacoesManutencaoView";
import { ControleInsumosComercialView } from "./components/ControleInsumosComercialView";
import { WhatsAppMessageEditor } from "./components/WhatsAppMessageEditor";
import { AdminFuncionariosView } from "./components/AdminFuncionariosView";
import { IsencoesView } from "./components/IsencoesView";
import { WhatsAppMessageSelector } from "./components/WhatsAppMessageSelector";
import { MultiSelect } from "./components/MultiSelect";
import { EvasaoView } from "./components/EvasaoView";
import NovasOportunidadesView from "./components/NovasOportunidadesView";
import ControleLigacoesView from "./components/ControleLigacoesView";
import CRMView from "./components/CRMView";
import MetaSMView from "./components/MetaSMView";
import MetaCursosView from "./components/MetaCursosView";

// --- Helpers ---
export const replaceMessageVariables = (
  template: string,
  lead: any,
): string => {
  if (!template) return "";
  let text = template;
  text = text.replace(/\[nome\]/gi, lead.nome || "");
  text = text.replace(/\[curso\]/gi, lead.curso || lead.cursoInteresse || "");
  text = text.replace(/\[matr[iÃ­]cula\]/gi, lead.numeroMatricula || "");

  // Novas variÃ¡veis
  text = text.replace(
    /\[unidade\]/gi,
    lead.unidade || lead.nome_unidade || "nossa unidade",
  );
  text = text.replace(
    /\[data_contato\]/gi,
    new Date().toLocaleDateString("pt-BR"),
  );

  const hour = new Date().getHours();
  const saudacao =
    hour < 12 ? "Bom dia" : hour < 18 ? "Boa tarde" : "Boa noite";
  text = text.replace(/\[saudacao\]/gi, saudacao);

  if (lead.missingDocs) {
    text = text.replace(
      /\[pendencias\]/gi,
      Array.isArray(lead.missingDocs)
        ? lead.missingDocs.join(", ")
        : lead.missingDocs,
    );
  }

  return text;
};

const exportToExcel = (data: any[], fileName: string) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

const exportToCSV = (data: any[], fileName: string) => {
  if (data.length === 0) return;
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(","),
    ...data.map((row) =>
      headers.map((header) => JSON.stringify(row[header])).join(","),
    ),
  ].join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", `${fileName}.csv`);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const importFromExcel = (file: File, callback: (data: any[]) => void) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const bstr = e.target?.result;
    const workbook = XLSX.read(bstr, { type: "binary" });
    const worksheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    callback(data);
  };
  reader.readAsBinaryString(file);
};

// Component WhatsAppMessageSelector moved to src/components/WhatsAppMessageSelector.tsx

// --- Constants ---
const HOLIDAYS = [
  "2024-01-01",
  "2024-03-29",
  "2024-04-21",
  "2024-05-01",
  "2024-05-30",
  "2024-07-09",
  "2024-09-07",
  "2024-10-12",
  "2024-11-02",
  "2024-11-15",
  "2024-11-20",
  "2024-12-25",
  "2025-01-01",
  "2025-04-18",
  "2025-04-21",
  "2025-05-01",
  "2025-06-19",
  "2025-09-07",
  "2025-10-12",
  "2025-11-02",
  "2025-11-15",
  "2025-11-20",
  "2025-12-25",
  "2026-01-01",
  "2026-04-03",
  "2026-04-21",
  "2026-05-01",
  "2026-06-04",
  "2026-09-07",
  "2026-10-12",
  "2026-11-02",
  "2026-11-15",
  "2026-11-20",
  "2026-12-25",
];

const getWorkingDaysRemaining = (endDateStr: string) => {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const dataFim = new Date(endDateStr);
  dataFim.setHours(0, 0, 0, 0);

  if (dataFim < hoje) return 0;

  let count = 0;
  let curDate = new Date(hoje.getTime());
  // Start counting from today if it's a working day
  while (curDate <= dataFim) {
    const dayOfWeek = curDate.getDay(); // 0 = Sunday
    const dateString = curDate.toISOString().split("T")[0];
    const isSunday = dayOfWeek === 0;
    const isHoliday = HOLIDAYS.includes(dateString);

    if (!isSunday && !isHoliday) {
      count++;
    }
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
};

const getWorkingDaysBetween = (startDateStr: string, endDateStr: string) => {
  const start = new Date(startDateStr);
  start.setHours(0, 0, 0, 0);
  const end = new Date(endDateStr);
  end.setHours(0, 0, 0, 0);

  if (end < start) return 0;

  let count = 0;
  let curDate = new Date(start.getTime());
  while (curDate <= end) {
    const dayOfWeek = curDate.getDay();
    const dateString = curDate.toISOString().split("T")[0];
    const isSunday = dayOfWeek === 0;
    const isHoliday = HOLIDAYS.includes(dateString);

    if (!isSunday && !isHoliday) {
      count++;
    }
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
};

const formatLocalDateString = (dateStr: string) => {
  if (!dateStr) return "";
  const dateOnly = dateStr.split("T")[0];
  if (dateOnly.includes("-")) {
    const parts = dateOnly.split("-");
    if (parts.length === 3) {
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return dateStr;
};

export const ROLES: Record<string, UserRole> = {
  ADMIN_MASTER: "Admin Master",
  PROMOTOR: "Promotor",
  FDV: "FDV",
  SALA_MATRICULA: "Sala de MatrÃ­cula",
  QG: "QG",
  LIDER_FDV: "LÃ­der/FDV",
  SSA: "SSA",
  GESTOR_UNIDADE: "Gestor Unidade",
  GESTOR_COMERCIAL: "Gestor Comercial",
  ACADEMICO: "AcadÃªmico",
  PROMOTOR_RUA: "Promotor/rua",
  GESTOR_COMERCIAL_COMERCIAL: "Gerente Comercial (Comercial)",
  FDV_COMERCIAL: "FDV (Comercial)",
  FINANCEIRO: "Financeiro",
  TECNICO: "TÃ©cnico",
};

const VIEW_PERMISSIONS: Record<string, UserRole[]> = {
  dashboard: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  formularios: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  relatorios: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  cadastro: [
    ROLES.ADMIN_MASTER,
    ROLES.PROMOTOR,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  historico: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  bases: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
  ],
  gap: [ROLES.ADMIN_MASTER, ROLES.SALA_MATRICULA, ROLES.LIDER_FDV, ROLES.GESTOR_UNIDADE],
  fiesProuni: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
  ],
  campanhas: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  calendario: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  empresas: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.QG,
    ROLES.SSA,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  calculo: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.SSA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  mapao: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  basesDisparo: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
  ],
  basesRenovacao: [ROLES.ADMIN_MASTER, ROLES.LIDER_FDV, ROLES.SSA],
  avisos: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.ACADEMICO,
  ],
  emailMarketing: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  controleConcorrencia: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FINANCEIRO,
  ],
  admin: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL,
  ],
  crm: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
  ],
  controlePagamentos: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
  ],
  evasao: [
    ROLES.ADMIN_MASTER,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.SALA_MATRICULA,
  ],
  cursos: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  solicitacaoManutencao: Object.values(ROLES),
  controleInsumos: [
    ROLES.ADMIN_MASTER,
    ROLES.ACADEMICO,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
    ROLES.GESTOR_UNIDADE,
    ROLES.LIDER_FDV,
  ],
  controleInsumosComercial: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FINANCEIRO,
  ],
  isencoes: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  controleLigacoes: [
    ROLES.ADMIN_MASTER,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
  ],
};

// --- Components ---
function PasswordChangeModal({ onComplete }: { onComplete: () => void }) {
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (newPassword.length < 6) {
      setError("A senha deve ter pelo menos 6 caracteres.");
      return;
    }
    if (newPassword !== confirmPassword) {
      setError("As senhas nÃ£o coincidem.");
      return;
    }

    setLoading(true);
    setError("");
    try {
      if (auth.currentUser) {
        await updatePassword(auth.currentUser, newPassword);
        onComplete();
      }
    } catch (err: any) {
      setError("Erro ao atualizar senha. Tente sair e entrar novamente.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="bg-white p-8 rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full"
      >
        <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-2xl flex items-center justify-center mb-6">
          <KeyRound size={32} />
        </div>
        <h2 className="text-2xl font-bold text-slate-900 mb-2">
          Troca de Senha ObrigatÃ³ria
        </h2>
        <p className="text-slate-500 mb-6">
          Para sua seguranÃ§a, vocÃª deve alterar sua senha padrÃ£o antes de
          continuar.
        </p>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1">
              Nova Senha
            </label>
            <input
              type="password"
              required
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              placeholder="MÃ­nimo 6 caracteres"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1">
              Confirmar Nova Senha
            </label>
            <input
              type="password"
              required
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>

          {error && (
            <div className="p-3 bg-rose-50 text-rose-600 text-xs font-bold rounded-xl flex items-center space-x-2">
              <AlertCircle size={14} />
              <span>{error}</span>
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-blue-600 text-white py-4 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 disabled:opacity-50"
          >
            {loading ? "Atualizando..." : "Atualizar Senha"}
          </button>
        </form>
      </motion.div>
    </div>
  );
}

const Toast = ({
  message,
  type,
  onClose,
}: {
  message: string;
  type: "success" | "error";
  onClose: () => void;
}) => (
  <motion.div
    initial={{ x: 100, opacity: 0 }}
    animate={{ x: 0, opacity: 1 }}
    exit={{ x: 100, opacity: 0 }}
    className={cn(
      "fixed top-5 right-5 z-50 p-4 rounded-lg shadow-lg flex items-center space-x-2 text-white",
      type === "success" ? "bg-emerald-600" : "bg-rose-600",
    )}
  >
    {type === "success" ? (
      <CheckCircle2 size={20} />
    ) : (
      <AlertCircle size={20} />
    )}
    <span className="font-medium">{message}</span>
    <button onClick={onClose} className="ml-2 hover:opacity-80">
      <X size={16} />
    </button>
  </motion.div>
);

function MapaoAcademicoView({
  mapao,
  onToast,
  profile,
}: {
  mapao: MapaoAcademicoEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
}) {
  const [showModal, setShowModal] = useState(false);
  const [editingEntry, setEditingEntry] = useState<MapaoAcademicoEntry | null>(
    null,
  );

  const defaultDisciplina = {
    codDisc: "",
    disciplina: "",
    dia: "Segunda-feira",
    horario: "",
    turma: "",
    tipoDisciplina: "PRESENCIAL",
    professor: "",
    matricula: "",
    observacao: "",
    linkAula: "",
  };

  const [formData, setFormData] = useState<Partial<MapaoAcademicoEntry>>({
    modalidade: "Presencial",
    tipoCurso: "GRADUACAO",
    periodo: "",
    semestre: "",
    disciplinas: [{ ...defaultDisciplina }],
  });
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});
  
  // Filters
  const [filterCurso, setFilterCurso] = useState("");
  const [filterPeriodo, setFilterPeriodo] = useState("");
  const [filterSemestre, setFilterSemestre] = useState("");
  const [filterTipoCurso, setFilterTipoCurso] = useState("");

  const toggleExpand = (id: string) => {
    setExpandedCards(prev => ({ ...prev, [id]: !prev[id] }));
  };


  const handleShareCard = async (cardId: string, cardName: string) => {
    const element = document.getElementById(`mapao-card-${cardId}`);
    if (!element) return;
    try {
      onToast("Gerando PDF, aguarde...", "success");
      const canvas = await html2canvas(element, { backgroundColor: '#ffffff', scale: 1, useCORS: true, logging: false });
      if (canvas.width === 0 || canvas.height === 0) {
        throw new Error("Canvas gerado estÃ¡ vazio (tamanho 0). O elemento pode estar oculto.");
      }
      const imgData = canvas.toDataURL('image/jpeg', 0.9);
      if (typeof jsPDF !== "function") throw new Error("jsPDF nÃ£o carregado corretamente");
      // Fallback for jspdf format
      const pdf = new jsPDF({
        orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
        unit: 'px',
        format: [canvas.width, canvas.height]
      });
      pdf.addImage(imgData, 'JPEG', 0, 0, canvas.width, canvas.height);
      const safeName = cardName.replace(/[^a-zA-Z0-9]/g, '_');
      pdf.save(`mapao-${safeName}.pdf`);
      onToast("PDF gerado com sucesso!", "success");
    } catch (error: any) {
      console.error('Failed to generate PDF', error);
      onToast(`Erro ao gerar PDF: ${error.message || "Erro desconhecido"}`, "error");
    }
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();

    const isDuplicate = mapao.some(
      (m) =>
        m.curso?.toLowerCase() === formData.curso?.toLowerCase() &&
        m.modalidade === formData.modalidade &&
        m.periodo === formData.periodo &&
        m.id !== editingEntry?.id,
    );

    if (isDuplicate) {
      onToast(
        "Este curso/modalidade/perÃ­odo jÃ¡ estÃ¡ cadastrado no MapÃ£o.",
        "error",
      );
      return;
    }

    try {
      if (editingEntry) {
        await updateDoc(doc(db, COLLECTIONS.MAPAO_ACADEMICO, editingEntry.id), {
          ...formData,
          createdAt: serverTimestamp(),
        });
        onToast("Registro atualizado!");
      } else {
        await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
          ...formData,
          createdAt: serverTimestamp(),
        });
        onToast("Registro cadastrado!");
      }
      setShowModal(false);
      setEditingEntry(null);
      setFormData({
        modalidade: "Presencial",
        tipoCurso: "GRADUACAO",
        periodo: "",
        semestre: "",
        disciplinas: [{ ...defaultDisciplina }],
      });
    } catch (err: any) {
      onToast("Erro ao salvar.", "error");
    }
  };

  const handleDuplicate = async (entry: MapaoAcademicoEntry) => {
    try {
      const { id, ...data } = entry;
      await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
        ...data,
        createdAt: serverTimestamp(),
      });
      onToast("Registro duplicado!");
    } catch (err: any) {
      onToast("Erro ao duplicar.", "error");
    }
  };

  const handleAddDisciplina = () => {
    if (formData.disciplinas && formData.disciplinas.length < 7) {
      setFormData((prev) => ({
        ...prev,
        disciplinas: [...(prev.disciplinas || []), { ...defaultDisciplina }],
      }));
    }
  };

  const handleRemoveDisciplina = (index: number) => {
    const newDisciplinas = [...(formData.disciplinas || [])];
    newDisciplinas.splice(index, 1);
    setFormData((prev) => ({ ...prev, disciplinas: newDisciplinas }));
  };

  const handleChangeDisciplina = (
    index: number,
    field: string,
    value: string,
  ) => {
    const newDisciplinas: any = [...(formData.disciplinas || [])];
    newDisciplinas[index][field] = value;
    if (field === "dia" && value === "Virtual") {
      newDisciplinas[index].horario = "";
    }
    setFormData((prev) => ({ ...prev, disciplinas: newDisciplinas }));
  };

  const handleExport = () => {
    const exportData: any[] = [];
    mapao.forEach(m => {
      const disciplinas = m.disciplinas || [];
      if (disciplinas.length === 0) {
        exportData.push({
          Modalidade: m.modalidade,
          Curso: m.curso,
          PerÃ­odo: m.periodo,
          "Tipo Curso": m.tipoCurso,
          "CÃ³d. Disciplina": "",
          Disciplina: "",
          Dia: "",
          HorÃ¡rio: "",
          Turma: "",
          "Tipo Disciplina": "",
          Professor: "",
          MatrÃ­cula: "",
          ObservaÃ§Ã£o: "",
          "Link Aula": ""
        });
      } else {
        disciplinas.forEach(d => {
          exportData.push({
            Modalidade: m.modalidade,
            Curso: m.curso,
            PerÃ­odo: m.periodo,
            "Tipo Curso": m.tipoCurso,
            "CÃ³d. Disciplina": d.codDisc,
            Disciplina: d.disciplina,
            Dia: d.dia,
            HorÃ¡rio: d.horario,
            Turma: d.turma,
            "Tipo Disciplina": d.tipoDisciplina,
            Professor: d.professor,
            MatrÃ­cula: d.matricula,
            ObservaÃ§Ã£o: d.observacao,
            "Link Aula": d.linkAula || ""
          });
        });
      }
    });
    exportToExcel(exportData, "Mapao_Academico");
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return "";
        };

        const map = new Map<string, any>();
        
        data.forEach((row: any) => {
          const curso = String(getVal(row, "curso") || "").trim();
          const modalidade = String(getVal(row, "modalidade") || "").trim();
          const periodo = String(getVal(row, "perÃ­odo", "periodo") || "").trim();
          
          if (!curso) return; 
          
          const key = `${curso.toLowerCase()}|${modalidade.toLowerCase()}|${periodo.toLowerCase()}`;
          
          if (!map.has(key)) {
            map.set(key, {
              curso,
              modalidade: modalidade || "Presencial",
              periodo,
              tipoCurso: String(getVal(row, "tipo curso", "tipocurso") || "GRADUACAO").trim(),
              disciplinas: []
            });
          }
          
          const entry = map.get(key);
          const disciplina = String(getVal(row, "disciplina") || "").trim();
          
          if (disciplina) {
            entry.disciplinas.push({
              codDisc: String(getVal(row, "cÃ³d. disciplina", "cod disciplina", "coddisc") || "").trim(),
              disciplina,
              dia: String(getVal(row, "dia") || "").trim() || "Segunda-feira",
              horario: String(getVal(row, "horÃ¡rio", "horario") || "").trim(),
              turma: String(getVal(row, "turma") || "").trim(),
              tipoDisciplina: String(getVal(row, "tipo disciplina", "tipodisciplina") || "").trim() || "PRESENCIAL",
              professor: String(getVal(row, "professor") || "").trim(),
              matricula: String(getVal(row, "matrÃ­cula", "matricula") || "").trim(),
              observacao: String(getVal(row, "observaÃ§Ã£o", "observacao") || "").trim(),
              linkAula: String(getVal(row, "link aula", "linkaula") || "").trim()
            });
          }
        });
        
        let importedCount = 0;
        const newEntries = Array.from(map.values());
        
        for (const item of newEntries) {
          const isDuplicate = mapao.some(
            (m) =>
              m.curso?.toLowerCase() === item.curso?.toLowerCase() &&
              m.modalidade?.toLowerCase() === item.modalidade?.toLowerCase() &&
              m.periodo?.toLowerCase() === item.periodo?.toLowerCase()
          );

          if (!isDuplicate) {
            await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
              ...item,
              createdAt: serverTimestamp(),
            });
            importedCount++;
          }
        }
                
        onToast(`${importedCount} novos registros importados com sucesso!`, "success");
      } catch (err: any) {
        onToast("Erro ao importar arquivo.", "error");
      }
    });
    
    e.target.value = ''; // Reset input
  };

  const canEdit = true;

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
            MapÃ£o AcadÃªmico
          </h2>
          <p className="text-sm text-slate-500">
            GestÃ£o de cursos, disciplinas e horÃ¡rios
          </p>
        </div>

        {canEdit && (
          <div className="flex gap-2">
            <input
              type="file"
              accept=".xlsx, .xls"
              className="hidden"
              id="import-mapao-file"
              onChange={handleImport}
            />
            <label
              htmlFor="import-mapao-file"
              className="cursor-pointer bg-slate-100 text-slate-700 px-4 py-2.5 rounded-2xl font-bold hover:bg-slate-200 transition-all flex items-center space-x-2"
            >
              <Upload size={18} />
              <span className="hidden sm:inline">Importar</span>
            </label>
            <button
              onClick={handleExport}
              className="bg-slate-100 text-slate-700 px-4 py-2.5 rounded-2xl font-bold hover:bg-slate-200 transition-all flex items-center space-x-2"
            >
              <Download size={18} />
              <span className="hidden sm:inline">Exportar</span>
            </button>
            <button
              onClick={() => {
                setEditingEntry(null);
                setFormData({
                  modalidade: "Presencial",
                  tipoCurso: "GRADUACAO",
                  disciplinas: [{ ...defaultDisciplina }],
                });
                setShowModal(true);
              }}
              className="bg-blue-600 text-white px-6 py-2.5 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 flex items-center space-x-2"
            >
              <Plus size={20} />
              <span>Novo Cadastro</span>
            </button>
          </div>
        )}

      </div>

      {/* Filter Bar */}
      <div className="bg-white p-4 rounded-3xl border border-slate-100 shadow-sm flex flex-col md:flex-row gap-4 mb-6">
        <div className="flex-1 relative">
          <Search size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Filtrar por Curso"
            value={filterCurso}
            onChange={(e) => setFilterCurso(e.target.value)}
            className="w-full pl-11 pr-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          />
        </div>
        <div className="flex-1">
          <input
            type="text"
            placeholder="Filtrar por PerÃ­odo"
            value={filterPeriodo}
            onChange={(e) => setFilterPeriodo(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          />
        </div>
        <div className="w-full md:w-48 shrink-0">
          <select
            value={filterSemestre}
            onChange={(e) => setFilterSemestre(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          >
            <option value="">Qualquer Semestre</option>
            <option value="1Âº">1Âº</option>
            <option value="2Âº">2Âº</option>
            <option value="3Âº">3Âº</option>
            <option value="4Âº">4Âº</option>
            <option value="5Âº">5Âº</option>
            <option value="6Âº">6Âº</option>
            <option value="7Âº">7Âº</option>
            <option value="8Âº">8Âº</option>
            <option value="9Âº">9Âº</option>
            <option value="10Âº">10Âº</option>
          </select>
        </div>
        <div className="w-full md:w-48 shrink-0">
          <select
            value={filterTipoCurso}
            onChange={(e) => setFilterTipoCurso(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          >
            <option value="">Qualquer Tipo</option>
            <option value="GRADUACAO">GRADUAÃ‡ÃƒO</option>
            <option value="TECNICO">TÃ‰CNICO</option>
          </select>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 gap-6">
        {mapao.filter((entry) => {
          const matchCurso = !filterCurso || entry.curso?.toLowerCase().includes(filterCurso.toLowerCase());
          const matchPeriodo = !filterPeriodo || entry.periodo?.toLowerCase().includes(filterPeriodo.toLowerCase());
          const matchSemestre = !filterSemestre || entry.semestre === filterSemestre;
          const matchTipo = !filterTipoCurso || entry.tipoCurso === filterTipoCurso;
          return matchCurso && matchPeriodo && matchSemestre && matchTipo;
        }).map((entry) => {
          const disciplinasList = entry.disciplinas || [];
          const isExpanded = expandedCards[entry.id];
          return (
            <motion.div
              id={`mapao-card-${entry.id}`}
              key={entry.id}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className={cn(
                "p-6 rounded-[2rem] border shadow-sm transition-all relative group flex flex-col gap-4 bg-white",
                entry.tipoCurso === "GRADUACAO"
                  ? "bg-white border-blue-100"
                  : "bg-white border-emerald-100",
              )}
            >
              <div className="flex justify-between items-start">
                <div className="flex flex-wrap gap-2">
                  <span
                    className={cn(
                      "px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider",
                      entry.tipoCurso === "GRADUACAO"
                        ? "bg-blue-100 text-blue-600"
                        : "bg-emerald-100 text-emerald-600",
                    )}
                  >
                    {entry.tipoCurso}
                  </span>
                  <span className="px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider bg-slate-100 text-slate-600">
                    {entry.modalidade}
                  </span>
                </div>
                <div className="flex space-x-1 shrink-0 bg-white/50 p-1 rounded-xl shadow-sm border border-slate-100/50 backdrop-blur-sm">
                  <button
                    onClick={() => handleShareCard(entry.id, entry.curso || "Curso")}
                    className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                    title="Baixar como PDF"
                  >
                    <FileText size={14} />
                  </button>
                  {canEdit && (
                    <>
                      <button
                        onClick={() => {
                          setEditingEntry(entry);
                          setFormData({
                            ...entry,
                            semestre: entry.semestre || "",
                            disciplinas:
                              disciplinasList.length > 0
                                ? disciplinasList
                                : [{ ...defaultDisciplina }],
                          });
                          setShowModal(true);
                        }}
                        className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                        title="Editar"
                      >
                        <Edit2 size={14} />
                      </button>
                      <button
                        onClick={() => handleDuplicate(entry)}
                        className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                        title="Duplicar"
                      >
                        <Copy size={14} />
                      </button>
                      <button
                        onClick={async () => {
                          if (window.confirm("Excluir?"))
                            await deleteDoc(
                              doc(db, COLLECTIONS.MAPAO_ACADEMICO, entry.id),
                            );
                        }}
                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                        title="Excluir"
                      >
                        <Trash2 size={14} />
                      </button>
                    </>
                  )}
                </div>
              </div>

              <div className="flex justify-between items-center gap-4 cursor-pointer" onClick={() => toggleExpand(entry.id)}>
                <div>
                  <h3 className="text-xl font-bold text-slate-900 leading-tight mb-1">
                    {entry.curso}
                  </h3>
                  <p className="text-xs text-slate-500 font-bold uppercase tracking-widest">
                    {entry.periodo} {entry.semestre ? ` - ${entry.semestre}` : ""}
                  </p>
                </div>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleExpand(entry.id);
                  }}
                  className="bg-slate-50 hover:bg-slate-100 p-2 rounded-full text-slate-500 transition-colors shrink-0"
                >
                  {isExpanded ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                </button>
              </div>

              {isExpanded && (
              <div className="flex-1 space-y-3 mt-2">
                {disciplinasList.map((disc, idx) => (
                  <div
                    key={idx}
                    className="bg-slate-50 p-4 rounded-2xl border border-slate-100 flex flex-col gap-2 relative overflow-hidden"
                  >
                    <div className="flex justify-between items-start gap-2">
                      <div>
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-1">
                          {disc.codDisc}
                        </p>
                        <p className="text-sm font-bold text-slate-800 leading-tight">
                          {disc.disciplina}
                        </p>
                      </div>
                      <span className={cn(
                        "px-2 py-1 text-[9px] font-bold uppercase tracking-wider rounded-lg shrink-0",
                        disc.tipoDisciplina === "PRESENCIAL" ? "bg-purple-100 text-purple-600" :
                        disc.tipoDisciplina === "TEAMS" || disc.tipoDisciplina === "ONLINE" ? "bg-blue-100 text-blue-600" :
                        "bg-orange-100 text-orange-600"
                      )}>
                        {disc.tipoDisciplina || "PRESENCIAL"}
                      </span>
                    </div>

                    <p className="text-xs text-slate-600 font-medium mt-1">
                      <span className="text-slate-400 font-normal">Prof:</span> {disc.professor || "-"}
                    </p>

                    <div className="flex flex-wrap gap-2 mt-2">
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Calendar size={12} className="text-blue-500" />
                        <span className="text-[10px] font-bold">
                          {disc.dia || "-"}
                        </span>
                      </div>
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Clock size={12} className="text-amber-500" />
                        <span className="text-[10px] font-bold">
                          {disc.horario || "-"}
                        </span>
                      </div>
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Users size={12} className="text-emerald-500" />
                        <span className="text-[10px] font-bold">
                          {disc.turma || "-"}

                        </span>
                      </div>
                    </div>

                    {disc.observacao && (
                      <div className="mt-3 text-xs text-slate-600 bg-white p-2.5 rounded-xl border border-slate-200 border-dashed">
                        {disc.observacao}
                      </div>
                    )}
                    {disc.linkAula && (
                      <a href={disc.linkAula} target="_blank" rel="noopener noreferrer" className="mt-2 flex items-center gap-1.5 text-xs font-bold text-blue-600 hover:text-blue-700 transition-colors w-fit bg-blue-50 px-3 py-1.5 rounded-lg">
                        <ExternalLink size={14} /> Link da Aula
                      </a>
                    )}
                  </div>
                ))}
                {disciplinasList.length === 0 && (
                  <div className="bg-slate-50 border border-slate-100 border-dashed rounded-2xl py-8 flex flex-col items-center justify-center text-slate-400">
                    <GraduationCap size={24} className="mb-2 opacity-50" />
                    <p className="text-xs italic">Nenhuma disciplina cadastrada.</p>
                  </div>
                )}
              </div>
              )}
            </motion.div>
          );
        })}
      </div>

      {showModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white rounded-[2.5rem] shadow-2xl max-w-3xl w-full p-8 overflow-y-auto max-h-[90vh] custom-scrollbar"
          >
            <div className="flex justify-between items-center mb-6 sticky top-0 bg-white py-2 z-10 border-b border-slate-100">
              <h3 className="text-2xl font-bold text-slate-900">
                {editingEntry ? "Editar Curso" : "Novo Cadastro AcadÃªmico"}
              </h3>
              <button
                type="button"
                onClick={() => setShowModal(false)}
                className="p-2 hover:bg-slate-100 rounded-full transition-all text-slate-400"
              >
                <X size={24} />
              </button>
            </div>

            <form onSubmit={handleSave} className="space-y-8">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pb-6 border-b border-slate-100">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    PerÃ­odo
                  </label>
                  <input
                    type="text"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.periodo || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, periodo: e.target.value })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Semestre
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.semestre || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, semestre: e.target.value })
                    }
                    required
                  >
                    <option value="">Selecione...</option>
                    <option value="1Âº">1Âº</option>
                    <option value="2Âº">2Âº</option>
                    <option value="3Âº">3Âº</option>
                    <option value="4Âº">4Âº</option>
                    <option value="5Âº">5Âº</option>
                    <option value="6Âº">6Âº</option>
                    <option value="7Âº">7Âº</option>
                    <option value="8Âº">8Âº</option>
                    <option value="9Âº">9Âº</option>
                    <option value="10Âº">10Âº</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Tipo de Curso
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.tipoCurso}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        tipoCurso: e.target.value as any,
                      })
                    }
                    required
                  >
                    <option value="GRADUACAO">GRADUAÃ‡ÃƒO</option>
                    <option value="TECNICO">TÃ‰CNICO</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Modalidade
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.modalidade}
                    onChange={(e) =>
                      setFormData({ ...formData, modalidade: e.target.value })
                    }
                    required
                  >
                    <option value="Presencial">Presencial</option>
                    <option value="EAD">EAD</option>
                    <option value="Semipresencial">Semipresencial</option>
                  </select>
                </div>

                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Nome do Curso
                  </label>
                  <input
                    type="text"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.curso || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, curso: e.target.value })
                    }
                    required
                  />
                </div>
              </div>

              <div>
                <div className="flex justify-between items-center mb-4">
                  <h4 className="font-bold text-slate-800 text-lg">
                    Disciplinas do Curso
                  </h4>
                  {(formData.disciplinas?.length || 0) < 7 && (
                    <button
                      type="button"
                      onClick={handleAddDisciplina}
                      className="text-sm flex items-center gap-1 text-blue-600 hover:text-blue-700 font-bold bg-blue-50 px-3 py-1.5 rounded-xl"
                    >
                      <Plus size={16} /> Adicionar (
                      {formData.disciplinas?.length || 0}/7)
                    </button>
                  )}
                </div>

                <div className="space-y-4">
                  {formData.disciplinas?.map((disc, idx) => (
                    <div
                      key={idx}
                      className="bg-slate-50 border border-slate-200 p-5 rounded-2xl relative"
                    >
                      {formData.disciplinas &&
                        formData.disciplinas.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveDisciplina(idx)}
                            className="absolute top-4 right-4 text-rose-400 hover:text-rose-600 hover:bg-rose-50 p-1.5 rounded-lg transition-colors"
                          >
                            <Trash2 size={16} />
                          </button>
                        )}
                      <h5 className="text-xs font-bold uppercase text-slate-400 mb-4">
                        Disciplina {idx + 1}
                      </h5>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            CÃ³digo
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.codDisc}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "codDisc",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Disciplina
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.disciplina}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "disciplina",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Dia da Semana
                          </label>
                          <select
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.dia}
                            onChange={(e) =>
                              handleChangeDisciplina(idx, "dia", e.target.value)
                            }
                            required
                          >
                            {[
                              "Segunda-feira",
                              "TerÃ§a-feira",
                              "Quarta-feira",
                              "Quinta-feira",
                              "Sexta-feira",
                              "SÃ¡bado",
                              "Virtual",
                            ].map((d) => (
                              <option key={d} value={d}>
                                {d}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            HorÃ¡rio{" "}
                            {disc.dia === "Virtual" ? "(NÃ£o se aplica)" : ""}
                          </label>
                          <input
                            type="text"
                            placeholder={
                              disc.dia === "Virtual"
                                ? "Virtual"
                                : "Ex: 19:00 - 22:00"
                            }
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500 disabled:opacity-50 disabled:bg-slate-100"
                            value={disc.horario}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "horario",
                                e.target.value,
                              )
                            }
                            required={disc.dia !== "Virtual"}
                            disabled={disc.dia === "Virtual"}
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Turma
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.turma}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "turma",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Tipo Disciplina
                          </label>
                          <select
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.tipoDisciplina}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "tipoDisciplina",
                                e.target.value,
                              )
                            }
                            required
                          >
                            <option value="PRESENCIAL">Presencial</option>
                            <option value="ONLINE">Online</option>
                            <option value="TEAMS">Teams</option>
                          </select>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex space-x-4 pt-4 sticky bottom-0 bg-white py-4 border-t border-slate-100 z-10">
                <button
                  type="submit"
                  className="flex-1 bg-blue-600 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-100 hover:bg-blue-700 transition-all"
                >
                  {editingEntry ? "Salvar AlteraÃ§Ãµes" : "Cadastrar"}
                </button>
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="px-8 bg-slate-100 text-slate-600 font-bold py-4 rounded-2xl hover:bg-slate-200 transition-all border border-slate-200"
                >
                  Cancelar
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function BasesDisparoView({
  bases,
  onToast,
}: {
  bases: BaseDisparoEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
}) {
  const [showModal, setShowModal] = useState(false);
  const [filterDate, setFilterDate] = useState(
    new Date().toISOString().split("T")[0],
  );
  const [formData, setFormData] = useState<Partial<BaseDisparoEntry>>({
    data: new Date().toISOString().split("T")[0],
    totalDisparos: 0,
    positivos: 0,
    negativos: 0,
  });

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      await addDoc(collection(db, COLLECTIONS.BASES_DISPARO), {
        ...formData,
        createdAt: serverTimestamp(),
      });
      onToast("Base registrada!");
      setShowModal(false);
      setFormData({
        data: new Date().toISOString().split("T")[0],
        totalDisparos: 0,
        positivos: 0,
        negativos: 0,
      });
    } catch (err: any) {
      onToast("Erro ao registrar.", "error");
    }
  };

  const filteredBases = bases.filter((b) => b.data === filterDate);

  const totalDisparos = filteredBases.reduce(
    (acc, b) => acc + b.totalDisparos,
    0,
  );
  const totalPositivos = filteredBases.reduce((acc, b) => acc + b.positivos, 0);
  const totalNegativos = filteredBases.reduce((acc, b) => acc + b.negativos, 0);

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
            Bases de Disparo
          </h2>
          <p className="text-sm text-slate-500">
            MÃ©tricas diÃ¡rias de disparos e conversÃ£o
          </p>
        </div>
        <div className="flex items-center space-x-4">
          <input
            type="date"
            className="px-4 py-2.5 rounded-2xl border border-slate-200 text-sm font-bold text-slate-700"
            value={filterDate}
            onChange={(e) => setFilterDate(e.target.value)}
          />
          <button
            onClick={() => setShowModal(true)}
            className="bg-blue-600 text-white px-6 py-2.5 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 flex items-center space-x-2"
          >
            <Plus size={20} />
            <span>Registrar Base</span>
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-slate-400 uppercase mb-1">
            Total de Disparos
          </p>
          <p className="text-3xl font-black text-blue-600">{totalDisparos}</p>
        </div>
        <div className="bg-emerald-50 p-6 rounded-3xl border border-emerald-100 shadow-sm">
          <p className="text-xs font-bold text-emerald-500 uppercase mb-1">
            Total Positivos
          </p>
          <p className="text-3xl font-black text-emerald-600">
            {totalPositivos}
          </p>
          <p className="text-xs font-bold text-emerald-500 mt-2">
            Taxa:{" "}
            {totalDisparos > 0
              ? ((totalPositivos / totalDisparos) * 100).toFixed(1)
              : 0}
            %
          </p>
        </div>
        <div className="bg-rose-50 p-6 rounded-3xl border border-rose-100 shadow-sm">
          <p className="text-xs font-bold text-rose-500 uppercase mb-1">
            Total Negativos
          </p>
          <p className="text-3xl font-black text-rose-600">{totalNegativos}</p>
        </div>
      </div>

      <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-6 border-b border-slate-100">
          <h3 className="text-xl font-bold text-slate-900">Listagem DiÃ¡ria</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="bg-slate-50 text-slate-500 font-bold uppercase tracking-wider">
                <th className="px-6 py-4">Nome da Base</th>
                <th className="px-6 py-4">Total Disparos</th>
                <th className="px-6 py-4">Positivos</th>
                <th className="px-6 py-4">Negativos</th>
                <th className="px-6 py-4">ConversÃ£o</th>
                <th className="px-6 py-4">AÃ§Ã£o</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-sm">
              {filteredBases.map((b) => (
                <tr key={b.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-6 py-4 font-bold text-slate-900">
                    {b.nomeBase}
                  </td>
                  <td className="px-6 py-4 font-bold text-blue-600">
                    {b.totalDisparos}
                  </td>
                  <td className="px-6 py-4 font-bold text-emerald-600">
                    {b.positivos}
                  </td>
                  <td className="px-6 py-4 font-bold text-rose-600">
                    {b.negativos}
                  </td>
                  <td className="px-6 py-4 font-bold text-slate-700">
                    {b.totalDisparos > 0
                      ? ((b.positivos / b.totalDisparos) * 100).toFixed(1)
                      : 0}
                    %
                  </td>
                  <td className="px-6 py-4">
                    <button
                      onClick={async () => {
                        if (window.confirm("Excluir?"))
                          await deleteDoc(
                            doc(db, COLLECTIONS.BASES_DISPARO, b.id),
                          );
                      }}
                      className="text-rose-500 hover:bg-rose-100 p-2 rounded-xl transition-all"
                    >
                      <Trash2 size={18} />
                    </button>
                  </td>
                </tr>
              ))}
              {filteredBases.length === 0 && (
                <tr>
                  <td
                    colSpan={6}
                    className="px-6 py-12 text-center text-slate-400 italic"
                  >
                    Nenhum registro para esta data.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {showModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-[2.5rem] shadow-2xl max-w-lg w-full p-8"
          >
            <div className="flex justify-between items-center mb-8">
              <h3 className="text-2xl font-bold text-slate-900">
                Registrar MÃ©tricas da Base
              </h3>
              <button
                onClick={() => setShowModal(false)}
                className="p-2 hover:bg-slate-100 rounded-full transition-all text-slate-400"
              >
                <X size={24} />
              </button>
            </div>

            <form onSubmit={handleSave} className="space-y-6">
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                  Data do Disparo
                </label>
                <input
                  type="date"
                  className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                  value={formData.data}
                  onChange={(e) =>
                    setFormData({ ...formData, data: e.target.value })
                  }
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                  Nome da Base
                </label>
                <input
                  type="text"
                  className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                  value={formData.nomeBase}
                  onChange={(e) =>
                    setFormData({ ...formData, nomeBase: e.target.value })
                  }
                  required
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Total
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.totalDisparos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        totalDisparos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Positivos
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.positivos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        positivos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Negativos
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.negativos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        negativos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
              </div>

              <div className="flex space-x-4 pt-4">
                <button
                  type="submit"
                  className="flex-1 bg-blue-600 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-100 hover:bg-blue-700 transition-all"
                >
                  Registrar
                </button>
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="px-8 bg-slate-100 text-slate-600 font-bold py-4 rounded-2xl hover:bg-slate-200 transition-all"
                >
                  Cancelar
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

const StatCard = ({
  title,
  value,
  icon: Icon,
  color,
  trend,
}: {
  title: string;
  value: string | number;
  icon: any;
  color: string;
  trend?: string;
}) => (
  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
    <div>
      <p className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-1">
        {title}
      </p>
      <p className="text-3xl font-bold text-slate-900">{value}</p>
      {trend && (
        <p className="text-xs text-emerald-600 font-medium mt-1 flex items-center">
          <TrendingUp size={12} className="mr-1" /> {trend}
        </p>
      )}
    </div>
    <div className={cn("p-4 rounded-2xl", color)}>
      <Icon size={24} className="text-white" />
    </div>
  </div>
);

// --- Main App ---

function CampanhasView({
  campanhas,
  onToast,
}: {
  campanhas: Campanha[];
  onToast: (m: string, t?: "success" | "error") => void;
}) {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isDetailModalOpen, setIsDetailModalOpen] = useState(false);
  const [editingCampanha, setEditingCampanha] = useState<Campanha | null>(null);
  const [selectedCampanha, setSelectedCampanha] = useState<Campanha | null>(
    null,
  );
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [startDateFilter, setStartDateFilter] = useState("");
  const [endDateFilter, setEndDateFilter] = useState("");

  const getEffectiveStatus = (camp: Campanha) => {
    const today = new Date().toISOString().split("T")[0];
    if (today < camp.dataInicio) return "Pendente";
    if (today > camp.dataFim) return "Finalizada";
    return "Ativa";
  };

  const filteredCampanhas = useMemo(() => {
    return campanhas.filter((camp) => {
      const effectiveStatus = getEffectiveStatus(camp);
      const matchesSearch = camp.nome
        .toLowerCase()
        .includes(searchTerm.toLowerCase());
      const matchesStatus = !statusFilter || effectiveStatus === statusFilter;

      let matchesDate = true;
      if (startDateFilter && endDateFilter) {
        matchesDate =
          camp.dataInicio <= endDateFilter && camp.dataFim >= startDateFilter;
      } else if (startDateFilter) {
        matchesDate = camp.dataFim >= startDateFilter;
      } else if (endDateFilter) {
        matchesDate = camp.dataInicio <= endDateFilter;
      }
      return matchesSearch && matchesStatus && matchesDate;
    });
  }, [campanhas, searchTerm, statusFilter, startDateFilter, endDateFilter]);

  const handleSave = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const payload = {
      nome: formData.get("nome") as string,
      dataInicio: formData.get("dataInicio") as string,
      dataFim: formData.get("dataFim") as string,
      objetivo: formData.get("objetivo") as string,
      updatedAt: serverTimestamp(),
    };

    const isDuplicate = campanhas.some(
      (c) =>
        c.nome.toLowerCase() === payload.nome.toLowerCase() &&
        c.id !== editingCampanha?.id,
    );
    if (isDuplicate) {
      onToast("JÃ¡ existe uma campanha com este nome.", "error");
      return;
    }

    try {
      if (editingCampanha) {
        await updateDoc(
          doc(db, COLLECTIONS.CAMPANHAS, editingCampanha.id),
          payload,
        );
        onToast("Campanha atualizada!");
      } else {
        await addDoc(collection(db, COLLECTIONS.CAMPANHAS), {
          ...payload,
          createdAt: serverTimestamp(),
        });
        onToast("Campanha criada!");
      }
      setIsModalOpen(false);
      setEditingCampanha(null);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, COLLECTIONS.CAMPANHAS);
      onToast("Erro ao salvar campanha.", "error");
    }
  };

  const handleExport = () => {
    const data = filteredCampanhas.map((c) => ({
      Nome: c.nome,
      "Data InÃ­cio": c.dataInicio,
      "Data Fim": c.dataFim,
      Status: c.status,
      Objetivo: c.objetivo,
    }));
    exportToExcel(data, "Campanhas");
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const batch = data.map((item) => {
          const rawStatus = String(getVal(item, "Status", "status") || "")
            .trim()
            .toLowerCase();
          const finalStatus =
            rawStatus === "ativa"
              ? "Ativa"
              : rawStatus === "inativa"
                ? "Inativa"
                : rawStatus === "pendente"
                  ? "Pendente"
                  : "Ativa";

          return {
            nome: String(getVal(item, "Nome", "nome") || "").trim(),
            dataInicio: String(
              getVal(item, "Data InÃ­cio", "dataInicio", "data_inicio") || "",
            ).trim(),
            dataFim: String(
              getVal(item, "Data Fim", "dataFim", "data_fim") || "",
            ).trim(),
            status: finalStatus,
            objetivo: String(getVal(item, "Objetivo", "objetivo") || "").trim(),
            createdAt: serverTimestamp(),
          };
        });

        let imported = 0;
        let skipped = 0;
        const inserted = new Set();
        for (const entry of batch) {
          if (!entry.nome) continue;
          const isDup =
            campanhas.some(
              (c) => c.nome.trim().toLowerCase() === entry.nome.toLowerCase(),
            ) ||
            Array.from(inserted).some(
              (name: any) =>
                String(name).toLowerCase() === entry.nome.toLowerCase(),
            );
          if (!isDup) {
            await addDoc(collection(db, COLLECTIONS.CAMPANHAS), entry);
            inserted.add(entry.nome);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} campanhas importadas! ${skipped > 0 ? `${skipped} ignoradas.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar campanhas.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-2xl font-bold text-slate-800">Campanhas</h2>
        <div className="flex space-x-2">
          <button
            onClick={() => {
              setEditingCampanha(null);
              setIsModalOpen(true);
            }}
            className="bg-blue-600 text-white px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
          >
            <Plus size={20} />
            <span>Nova Campanha</span>
          </button>
          <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
            <Upload size={18} />
            <span>Importar</span>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleImport}
              className="hidden"
            />
          </label>
          <button
            onClick={handleExport}
            className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
          >
            <Download size={18} />
            <span>Exportar</span>
          </button>
        </div>
      </div>

      <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex flex-wrap gap-4 items-center">
        <div className="relative flex-1 min-w-[200px]">
          <Search
            className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
            size={18}
          />
          <input
            type="text"
            placeholder="Buscar por nome..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
          />
        </div>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className="px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm"
        >
          <option value="">Todos os Status</option>
          <option value="Ativa">Ativa</option>
          <option value="Pendente">Pendente</option>
          <option value="Finalizada">Finalizada</option>
        </select>
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
            PerÃ­odo:
          </span>
          <input
            type="date"
            value={startDateFilter}
            onChange={(e) => setStartDateFilter(e.target.value)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm font-semibold text-slate-600"
            title="Data de InÃ­cio"
          />
          <span className="text-slate-400 text-xs font-bold">atÃ©</span>
          <input
            type="date"
            value={endDateFilter}
            onChange={(e) => setEndDateFilter(e.target.value)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm font-semibold text-slate-600"
            title="Data de Fim"
          />
          {(startDateFilter || endDateFilter) && (
            <button
              onClick={() => {
                setStartDateFilter("");
                setEndDateFilter("");
              }}
              className="text-xs font-bold text-red-500 hover:text-red-700 hover:underline transition-all cursor-pointer px-2"
            >
              Limpar
            </button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {filteredCampanhas.map((camp) => {
          const effectiveStatus = getEffectiveStatus(camp);
          return (
            <div
              key={camp.id}
              className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 flex flex-col justify-between cursor-pointer hover:shadow-md transition-all"
              onClick={() => {
                setSelectedCampanha(camp);
                setIsDetailModalOpen(true);
              }}
            >
              <div>
                <div className="flex justify-between items-start mb-4">
                  <h3 className="text-lg font-bold text-slate-900">
                    {camp.nome}
                  </h3>
                  <span
                    className={cn(
                      "px-2 py-1 rounded-full text-[10px] font-bold uppercase",
                      effectiveStatus === "Ativa"
                        ? "bg-emerald-100 text-emerald-600"
                        : effectiveStatus === "Pendente"
                          ? "bg-blue-100 text-blue-600"
                          : "bg-slate-100 text-slate-600",
                    )}
                  >
                    {effectiveStatus}
                  </span>
                </div>
                <p className="text-sm text-slate-500 mb-4 line-clamp-2">
                  {camp.objetivo}
                </p>
                <div className="flex items-center space-x-4 text-xs text-slate-400">
                  <div className="flex items-center space-x-1">
                    <Calendar size={14} />
                    <span>
                      {camp.dataInicio} - {camp.dataFim}
                    </span>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
        {filteredCampanhas.length === 0 && (
          <div className="col-span-full py-12 text-center text-slate-400 italic">
            Nenhuma campanha encontrada.
          </div>
        )}
      </div>

      <AnimatePresence>
        {isDetailModalOpen && selectedCampanha && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white p-8 rounded-3xl shadow-xl w-full max-w-lg space-y-6"
            >
              <h2 className="text-2xl font-bold text-slate-900">
                {selectedCampanha.nome}
              </h2>
              <div className="space-y-4">
                <div>
                  <p className="text-xs font-bold text-slate-500 uppercase">
                    PerÃ­odo
                  </p>
                  <p className="text-sm text-slate-700">
                    {selectedCampanha.dataInicio} - {selectedCampanha.dataFim}
                  </p>
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-500 uppercase">
                    Objetivo
                  </p>
                  <p className="text-sm text-slate-700">
                    {selectedCampanha.objetivo}
                  </p>
                </div>
              </div>
              <div className="flex space-x-4 pt-4">
                <button
                  onClick={() => setIsDetailModalOpen(false)}
                  className="flex-1 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold hover:bg-slate-200 transition-all"
                >
                  Fechar
                </button>
                <button
                  onClick={() => {
                    setEditingCampanha(selectedCampanha);
                    setIsDetailModalOpen(false);
                    setIsModalOpen(true);
                  }}
                  className="flex-1 py-3 bg-blue-600 text-white rounded-xl font-bold hover:bg-blue-700 transition-all"
                >
                  Editar Campanha
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden"
            >
              <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-bold text-slate-900">
                  {editingCampanha ? "Editar Campanha" : "Nova Campanha"}
                </h3>
                <button
                  onClick={() => setIsModalOpen(false)}
                  className="text-slate-400 hover:text-slate-600"
                >
                  <X size={24} />
                </button>
              </div>
              <form onSubmit={handleSave} className="p-6 space-y-4">
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome da Campanha
                  </label>
                  <input
                    name="nome"
                    defaultValue={editingCampanha?.nome}
                    required
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">
                      InÃ­cio
                    </label>
                    <input
                      type="date"
                      name="dataInicio"
                      defaultValue={editingCampanha?.dataInicio}
                      required
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">
                      Fim
                    </label>
                    <input
                      type="date"
                      name="dataFim"
                      defaultValue={editingCampanha?.dataFim}
                      required
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Status
                  </label>
                  <select
                    name="status"
                    defaultValue={editingCampanha?.status || "Ativa"}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  >
                    <option value="Ativa">Ativa</option>
                    <option value="Pausada">Pausada</option>
                    <option value="Finalizada">Finalizada</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Objetivo
                  </label>
                  <textarea
                    name="objetivo"
                    defaultValue={editingCampanha?.objetivo}
                    rows={3}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <button
                  type="submit"
                  className="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
                >
                  {editingCampanha ? "Salvar AlteraÃ§Ãµes" : "Criar Campanha"}
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function FiesProuniView({
  data,
  vagas = [],
  onToast,
  profile,
  whatsappMessages,
  periodos,
  botConfig,
  onSendBot,
  onMassSendBot,
}: {
  data: FiesProuniEntry[];
  vagas?: FiesProuniVaga[];
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
  whatsappMessages: WhatsAppMessage[];
  periodos: PeriodoCaptacao[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
}) {
  const [activeTab, setActiveTab] = useState<"lista" | "informacoes">(
    "informacoes",
  );
  const [searchTerm, setSearchTerm] = useState("");
  const [periodoFilter, setPeriodoFilter] = useState("");
  const [tipoFilter, setTipoFilter] = useState("");
  const [listaFilter, setListaFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [bolsaFilter, setBolsaFilter] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [situacaoFilter, setSituacaoFilter] = useState("");
  const [vagasPeriodoFilter, setVagasPeriodoFilter] = useState("");
  const [vagasMetodologiaFilter, setVagasMetodologiaFilter] = useState("");
  const [vagasBolsaFilter, setVagasBolsaFilter] = useState("");
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingEntry, setEditingEntry] = useState<FiesProuniEntry | null>(
    null,
  );
  const [isVagaModalOpen, setIsVagaModalOpen] = useState(false);
  const [editingVaga, setEditingVaga] = useState<FiesProuniVaga | null>(null);
  const [cpfInput, setCpfInput] = useState("");

  const isAdmin = true;

  useEffect(() => {
    if (editingEntry) {
      setCpfInput(formatCPF(editingEntry.cpf));
    } else {
      setCpfInput("");
    }
  }, [editingEntry, isModalOpen]);

  const filteredData = data.filter((item) => {
    // Restrict visibility to entries from the same unit, unless admin/gestor
    if (
      profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
      profile.role !== ROLES.GESTOR_COMERCIAL &&
      profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
      profile.role !== ROLES.SSA
    ) {
      if (profile.unidade && item.unidade && item.unidade !== profile.unidade) {
        return false;
      }
    }

    const matchesSearch =
      (item.nome || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.cpf || "").includes(searchTerm) ||
      (item.curso || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.lista &&
        item.lista.toLowerCase().includes(searchTerm.toLowerCase())) ||
      (item.responsavelEntrevista &&
        item.responsavelEntrevista
          .toLowerCase()
          .includes(searchTerm.toLowerCase())) ||
      (item.status &&
        item.status.toLowerCase().includes(searchTerm.toLowerCase()));
    const matchesPeriodo = !periodoFilter || item.periodo === periodoFilter;
    const matchesTipo = !tipoFilter || item.tipo === tipoFilter;
    const matchesLista = !listaFilter || item.lista === listaFilter;
    const matchesStatus = !statusFilter || item.status === statusFilter;
    const matchesBolsa = !bolsaFilter || item.bolsa === bolsaFilter;
    const matchesCurso = !cursoFilter || item.curso === cursoFilter;
    const matchesSituacao = !situacaoFilter || item.situacao === situacaoFilter;
    return (
      matchesSearch &&
      matchesPeriodo &&
      matchesTipo &&
      matchesLista &&
      matchesStatus &&
      matchesBolsa &&
      matchesCurso &&
      matchesSituacao
    );
  });

  const uniqueListas = Array.from(
    new Set(data.map((i) => i.lista).filter(Boolean)),
  ).sort();
  const uniqueStatuses = Array.from(
    new Set(data.map((i) => i.status).filter(Boolean)),
  ).sort();
  const uniquePeriodos = Array.from(
    new Set(data.map((i) => i.periodo).filter(Boolean)),
  ).sort();
  const uniqueCursos = Array.from(
    new Set(data.map((i) => i.curso).filter(Boolean)),
  ).sort();
  const uniqueSituacoes = Array.from(
    new Set(data.map((i) => i.situacao).filter(Boolean)),
  ).sort();

  const stats = {
    total: filteredData.length,
    pendentes: filteredData.filter((i) => i.docsEntreguesStatus === "Pendente")
      .length,
    parcial: filteredData.filter((i) => i.docsEntreguesStatus === "Parcial")
      .length,
    entregaram: filteredData.filter((i) => i.docsEntreguesStatus === "Sim")
      .length,
    comInscricao: filteredData.filter((i) => i.inscricaoSales).length,
    comMatricula: filteredData.filter((i) => i.numeroMatricula).length,
    emAnalise: filteredData.filter((i) => i.digitalizaStatus === "Em AnÃ¡lise")
      .length,
    concluido: filteredData.filter((i) => i.digitalizaStatus === "ConcluÃ­do")
      .length,
  };

  const safeVagas = Array.isArray(vagas) ? vagas : [];

  const filteredVagas = safeVagas.filter((item) => {
    if (
      profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
      profile.role !== ROLES.GESTOR_COMERCIAL &&
      profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
      profile.role !== ROLES.SSA
    ) {
      if (profile.unidade && item.unidade && item.unidade !== profile.unidade) {
        return false;
      }
    }
    const matchesPeriodo =
      !vagasPeriodoFilter || item.periodo === vagasPeriodoFilter;
    const matchesMetodologia =
      !vagasMetodologiaFilter || item.metodologia === vagasMetodologiaFilter;
    const matchesBolsa = !vagasBolsaFilter || item.bolsa === vagasBolsaFilter;

    return matchesPeriodo && matchesMetodologia && matchesBolsa;
  });

  const uniqueMetodologias = Array.from(
    new Set(safeVagas.map((i) => i.metodologia).filter(Boolean)),
  ).sort();

  const vagasStats = {
    totalVagas: filteredVagas.reduce(
      (acc, curr) => acc + (Number(curr?.vagas) || 0),
      0,
    ),
    total100: filteredVagas
      .filter((v) => v?.bolsa === "100%")
      .reduce((acc, curr) => acc + (Number(curr?.vagas) || 0), 0),
    total50: filteredVagas
      .filter((v) => v?.bolsa === "50%")
      .reduce((acc, curr) => acc + (Number(curr?.vagas) || 0), 0),
  };

  const handleSaveVaga = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const payload = {
      periodo: formData.get("periodo") as string,
      codCurso: formData.get("codCurso") as string,
      curso: formData.get("curso") as string,
      turno: formData.get("turno") as string,
      metodologia: formData.get("metodologia") as string,
      bolsa: formData.get("bolsa") as "50%" | "100%",
      vagas: parseInt(formData.get("vagas") as string, 10) || 0,
      unidade: (formData.get("unidade") as string) || "",
    };

    try {
      if (editingVaga) {
        await updateDoc(
          doc(db, COLLECTIONS.FIES_PROUNI_VAGAS, editingVaga.id),
          {
            ...payload,
            updatedAt: serverTimestamp(),
          },
        );
        onToast("Vaga atualizada com sucesso!", "success");
      } else {
        await addDoc(collection(db, COLLECTIONS.FIES_PROUNI_VAGAS), {
          ...payload,
          createdAt: serverTimestamp(),
        });
        onToast("Vaga cadastrada com sucesso!", "success");
      }
      setIsVagaModalOpen(false);
      setEditingVaga(null);
    } catch (err) {
      handleFirestoreError(
        err,
        editingVaga ? OperationType.UPDATE : OperationType.CREATE,
        COLLECTIONS.FIES_PROUNI_VAGAS,
      );
    }
  };

  const handleDeleteVaga = async (id: string) => {
    if (!window.confirm("Deseja realmente excluir esta vaga?")) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI_VAGAS, id));
      onToast("Vaga excluÃ­da com sucesso!");
    } catch (err) {
      handleFirestoreError(
        err,
        OperationType.DELETE,
        COLLECTIONS.FIES_PROUNI_VAGAS,
      );
    }
  };

  const handleSave = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const cpf = formData.get("cpf") as string;

    if (!validateCPF(cpf)) {
      onToast("CPF invÃ¡lido. Por favor, verifique os 11 dÃ­gitos.", "error");
      return;
    }

    const cleanCpf = cpf.replace(/\D/g, "");
    const isDuplicate = data.some(
      (item) => item.cpf === cleanCpf && item.id !== editingEntry?.id,
    );
    if (isDuplicate) {
      onToast("Este CPF jÃ¡ estÃ¡ cadastrado no FIES/Prouni.", "error");
      return;
    }

    const payload = {
      nome: formData.get("nome") as string,
      cpf: cpf.replace(/\D/g, ""), // Store only digits
      telefone: formData.get("telefone") as string,
      email: formData.get("email") as string,
      endereco: formData.get("endereco") as string,
      status: formData.get("status") as string,
      tipo: formData.get("tipo") as "FIES" | "PROUNI",
      bolsa: formData.get("bolsa") as "PARCIAL" | "INTEGRAL",
      situacao: formData.get("situacao") as "Candidato" | "Aluno (mesmo curso)" | "Aluno (outro curso)",
      cotaPPI: formData.get("cotaPPI") as "Sim" | "NÃ£o",
      metodologia: formData.get("metodologia") as string,
      curso: formData.get("curso") as string,
      inscricaoSales: formData.get("inscricaoSales") as string,
      numeroMatricula: formData.get("numeroMatricula") as string,
      tcbAssinado: formData.get("tcbAssinado") === "on",
      digitalizaStatus: formData.get("digitalizaStatus") as any,
      docsEntreguesStatus: formData.get("docsEntreguesStatus") as any,
      sisprouniStatus: formData.get("sisprouniStatus") as any,
      responsavelEntrevista: formData.get("responsavelEntrevista") as string,
      dataEntrevista: formData.get("dataEntrevista") as string,
      observacao: formData.get("observacao") as string,
      periodo: formData.get("periodo") as string,
      lista: formData.get("lista") as string,
      posicaoRanking: formData.get("posicaoRanking") as string,
      documentosEntregues:
        (formData.get("documentos") as string)
          ?.split(",")
          .map((s) => s.trim())
          .filter(Boolean) || [],
      updatedAt: serverTimestamp(),
    };

    try {
      if (editingEntry) {
        await updateDoc(
          doc(db, COLLECTIONS.FIES_PROUNI, editingEntry.id),
          payload,
        );
        onToast("Registro atualizado!");
      } else {
        await addDoc(collection(db, COLLECTIONS.FIES_PROUNI), {
          ...payload,
          unidade: profile.unidade || "",
          createdAt: serverTimestamp(),
        });
        onToast("Registro cadastrado!");
      }
      setIsModalOpen(false);
      setEditingEntry(null);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, COLLECTIONS.FIES_PROUNI);
      onToast("Erro ao salvar registro.", "error");
    }
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      onToast("Importando registros...");
      let successCount = 0;
      let errorCount = 0;

      const getVal = (row: any, ...keys: string[]) => {
        const rowKeys = Object.keys(row);
        for (const key of keys) {
          const foundKey = rowKeys.find(
            (k) => k.toLowerCase() === key.toLowerCase(),
          );
          if (foundKey && row[foundKey] !== undefined) return row[foundKey];
        }
        return undefined;
      };

      for (const row of data) {
        try {
          const rawCpf = String(getVal(row, "CPF", "cpf") || "");
          const cpf = rawCpf.replace(/\D/g, "");
          if (!cpf) continue;

          const payload = {
            nome: String(getVal(row, "Nome", "nome") || ""),
            cpf,
            telefone: String(getVal(row, "Telefone", "telefone") || ""),
            email: String(getVal(row, "Email", "email") || ""),
            endereco: String(getVal(row, "EndereÃ§o", "Endereco", "endereco") || ""),
            status: String(getVal(row, "Status", "status") || "Pendente"),
            tipo: (String(getVal(row, "Tipo", "tipo") || "PROUNI").toUpperCase() === "FIES"
              ? "FIES"
              : "PROUNI") as "FIES" | "PROUNI",
            bolsa: String(getVal(row, "Bolsa", "bolsa") || "INTEGRAL").toUpperCase().includes("PARCIAL") ? "PARCIAL" : "INTEGRAL",
            situacao: String(getVal(row, "SituaÃ§Ã£o", "Situacao", "situaÃ§Ã£o", "situacao") || "Candidato")
              .toLowerCase()
              .includes("outro curso")
              ? "Aluno (outro curso)"
              : String(getVal(row, "SituaÃ§Ã£o", "Situacao", "situaÃ§Ã£o", "situacao") || "")
                  .toLowerCase()
                  .includes("mesmo curso")
              ? "Aluno (mesmo curso)"
              : "Candidato",
            cotaPPI: String(getVal(row, "Cota PPI", "cota ppi", "Cota_PPI", "cotappi") || "").toLowerCase().includes("sim") ? "Sim" : "NÃ£o",
            curso: String(getVal(row, "Curso", "curso") || ""),
            posicaoRanking: String(getVal(row, "Ranking", "ranking") || ""),
            lista: String(getVal(row, "Lista", "lista") || ""),
            periodo: String(getVal(row, "Periodo", "PerÃ­odo", "periodo", "perÃ­odo") || ""),
            metodologia: String(getVal(row, "Metodologia", "metodologia") || ""),
            responsavelEntrevista: String(getVal(row, "ResponsÃ¡vel Entrevista", "Responsavel Entrevista", "responsavel entrevista") || ""),
            dataEntrevista: String(getVal(row, "Data Entrevista", "data entrevista") || ""),
            docsEntreguesStatus: String(
              getVal(row, "Status Docs", "status docs") || "Pendente",
            ) as any,
            inscricaoSales: String(getVal(row, "InscriÃ§Ã£o Sales", "Inscricao Sales", "inscricao sales") || ""),
            numeroMatricula: String(getVal(row, "NÃºmero MatrÃ­cula", "Numero Matricula", "numero matricula") || ""),
            digitalizaStatus: String(
              getVal(row, "Status Digitaliza", "status digitaliza") || "Pendente",
            ) as any,
            sisprouniStatus: String(getVal(row, "SISPROUNI", "sisprouni") || "Pendente") as any,
            tcbAssinado: String(getVal(row, "TCB Assinado", "tcb assinado")).toLowerCase() === "sim",
            documentosEntregues: String(getVal(row, "Documentos Entregues", "documentos entregues") || "")
              .split(",")
              .map((s) => s.trim())
              .filter(Boolean),
            observacao: String(getVal(row, "ObservaÃ§Ã£o", "Observacao", "observacao") || ""),
            unidade: profile.unidade || "",
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
          };

          await addDoc(collection(db, COLLECTIONS.FIES_PROUNI), payload);
          successCount++;
        } catch (err) {
          console.error("Erro ao importar registro Fies/Prouni:", err);
          errorCount++;
        }
      }

      onToast(
        `ImportaÃ§Ã£o concluÃ­da: ${successCount} sucesso, ${errorCount} erros`,
        successCount > 0 ? "success" : "error",
      );
    });
    e.target.value = "";
  };

  const handleExport = () => {
    const exportData = filteredData.map((item) => ({
      Nome: item.nome,
      CPF: item.cpf,
      Telefone: item.telefone || "",
      Email: item.email || "",
      EndereÃ§o: item.endereco || "",
      Status: item.status || "",
      Tipo: item.tipo,
      Bolsa: item.bolsa,
      SituaÃ§Ã£o: item.situacao || "",
      "Cota PPI": item.cotaPPI || "",
      Curso: item.curso,
      Ranking: item.posicaoRanking || "",
      Lista: item.lista || "",
      Periodo: item.periodo || "",
      Metodologia: item.metodologia || "",
      "ResponsÃ¡vel Entrevista": item.responsavelEntrevista || "",
      "Data Entrevista": item.dataEntrevista || "",
      "Status Docs": item.docsEntreguesStatus || "",
      "InscriÃ§Ã£o Sales": item.inscricaoSales || "",
      "NÃºmero MatrÃ­cula": item.numeroMatricula || "",
      "Status Digitaliza": item.digitalizaStatus,
      SISPROUNI: item.sisprouniStatus || "Pendente",
      "TCB Assinado": item.tcbAssinado ? "Sim" : "NÃ£o",
      "Documentos Entregues": item.documentosEntregues?.join(", ") || "",
      ObservaÃ§Ã£o: item.observacao || "",
    }));
    exportToExcel(exportData, "Fies_Prouni");
  };

  const handleExportVagas = () => {
    const exportData = vagas.map((v) => ({
      PerÃ­odo: v.periodo || "",
      "Cod. Curso": v.codCurso || "",
      Curso: v.curso || "",
      Turno: v.turno || "",
      Metodologia: v.metodologia || "",
      Bolsa: v.bolsa || "",
      Vagas: v.vagas || 0,
      Unidade: v.unidade || "",
    }));
    exportToExcel(exportData, "Fies_Prouni_Vagas");
  };

  const handleImportVagas = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      onToast("Importando vagas...");
      let successCount = 0;
      let errorCount = 0;

      const getVal = (row: any, ...keys: string[]) => {
        const rowKeys = Object.keys(row);
        for (const key of keys) {
          const foundKey = rowKeys.find(
            (k) => k.toLowerCase() === key.toLowerCase(),
          );
          if (foundKey && row[foundKey] !== undefined) return row[foundKey];
        }
        return undefined;
      };

      for (const row of data) {
        try {
          const payload = {
            periodo: String(getVal(row, "PerÃ­odo", "Periodo", "perÃ­odo", "periodo") || ""),
            codCurso: String(getVal(row, "Cod. Curso", "cod curso", "cod. curso", "codCurso") || ""),
            curso: String(getVal(row, "Curso", "curso") || ""),
            turno: String(getVal(row, "Turno", "turno") || ""),
            metodologia: String(getVal(row, "Metodologia", "metodologia") || ""),
            bolsa: String(getVal(row, "Bolsa", "bolsa") || "") as "50%" | "100%",
            vagas: parseInt(String(getVal(row, "Vagas", "vagas")), 10) || 0,
            unidade: String(getVal(row, "Unidade", "unidade") || ""),
            createdAt: serverTimestamp(),
          };

          if (payload.curso && payload.periodo && payload.bolsa) {
            await addDoc(
              collection(db, COLLECTIONS.FIES_PROUNI_VAGAS),
              payload,
            );
            successCount++;
          }
        } catch (err) {
          console.error("Erro ao importar vaga:", err);
          errorCount++;
        }
      }

      onToast(
        `ImportaÃ§Ã£o concluÃ­da: ${successCount} sucesso, ${errorCount} erros`,
        successCount > 0 ? "success" : "error",
      );
    });
    e.target.value = "";
  };

  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);

  const handleBulkDelete = async () => {
    if (selectedEntries.length === 0) return;
    if (
      window.confirm(
        `Deseja excluir ${selectedEntries.length} registros Fies/Prouni selecionados?`,
      )
    ) {
      try {
        for (const id of selectedEntries) {
          await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI, id));
        }
        onToast(`${selectedEntries.length} registros removidos.`);
        setSelectedEntries([]);
      } catch (err: any) {
        onToast("Erro ao excluir registros.", "error");
      }
    }
  };

  const handleDeleteIndividual = async (id: string) => {
    if (window.confirm("Deseja excluir este registro?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI, id));
        onToast("Registro removido.");
      } catch (err: any) {
        onToast("Erro ao excluir registro.", "error");
      }
    }
  };

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries([...selectedEntries, id]);
    } else {
      setSelectedEntries(selectedEntries.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(filteredData.map((b) => b.id));
    } else {
      setSelectedEntries([]);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div className="flex items-center gap-4">
          <h2 className="text-2xl font-bold text-gray-800">
            Acompanhamento Fies/Prouni
          </h2>
          <div className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab("informacoes")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "informacoes" ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}
            >
              InformaÃ§Ãµes
            </button>
            <button
              onClick={() => setActiveTab("lista")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "lista" ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}
            >
              Lista
            </button>
          </div>
        </div>
        <div className="flex space-x-2">
          {activeTab === "lista" ? (
            <>
              <button
                onClick={() => {
                  setEditingEntry(null);
                  setIsModalOpen(true);
                }}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg flex items-center space-x-2 hover:bg-indigo-700 transition-colors"
              >
                <Plus size={20} />
                <span>Novo Cadastro</span>
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar Lista</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  className="hidden"
                  onChange={handleImport}
                />
              </label>
              <button
                onClick={handleExport}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
            </>
          ) : (
            <>
              <button
                onClick={() => {
                  setEditingVaga(null);
                  setIsVagaModalOpen(true);
                }}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg flex items-center space-x-2 hover:bg-indigo-700 transition-colors"
              >
                <Plus size={20} />
                <span>Nova Vaga</span>
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar Vagas</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  className="hidden"
                  onChange={handleImportVagas}
                />
              </label>
              <button
                onClick={handleExportVagas}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
            </>
          )}
        </div>
      </div>

      {activeTab === "lista" && (
        <>
          {/* Dashboard Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Total Candidatos"
              value={stats.total}
              icon={Users}
              color="bg-blue-500"
            />
            <StatCard
              title="Pendentes Doc"
              value={stats.pendentes}
              icon={AlertCircle}
              color="bg-red-500"
            />
            <StatCard
              title="Docs Parciais"
              value={stats.parcial}
              icon={Clock}
              color="bg-amber-500"
            />
            <StatCard
              title="Docs Entregues"
              value={stats.entregaram}
              icon={CheckCircle2}
              color="bg-green-500"
            />
            <StatCard
              title="Com InscriÃ§Ã£o"
              value={stats.comInscricao}
              icon={FileText}
              color="bg-indigo-500"
            />
            <StatCard
              title="Com MatrÃ­cula"
              value={stats.comMatricula}
              icon={GraduationCap}
              color="bg-purple-500"
            />
            <StatCard
              title="Em AnÃ¡lise"
              value={stats.emAnalise}
              icon={Clock}
              color="bg-amber-500"
            />
            <StatCard
              title="Docs OK"
              value={stats.concluido}
              icon={ShieldCheck}
              color="bg-emerald-500"
            />
          </div>

          {/* Filters */}
          <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex flex-wrap gap-4">
            <div className="flex-1 min-w-[200px] relative">
              <Search
                className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                size={18}
              />
              <input
                type="text"
                placeholder="Pesquisar por nome, CPF ou curso..."
                className="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={periodoFilter}
              onChange={(e) => setPeriodoFilter(e.target.value)}
            >
              <option value="">Todos os PerÃ­odos</option>
              {uniquePeriodos.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 max-w-[200px]"
              value={cursoFilter}
              onChange={(e) => setCursoFilter(e.target.value)}
            >
              <option value="">Todos os Cursos</option>
              {uniqueCursos.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={tipoFilter}
              onChange={(e) => setTipoFilter(e.target.value)}
            >
              <option value="">Fies & Prouni</option>
              <option value="FIES">Apenas FIES</option>
              <option value="PROUNI">Apenas PROUNI</option>
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={bolsaFilter}
              onChange={(e) => setBolsaFilter(e.target.value)}
            >
              <option value="">Todas as Bolsas</option>
              <option value="INTEGRAL">INTEGRAL</option>
              <option value="PARCIAL">PARCIAL</option>
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={situacaoFilter}
              onChange={(e) => setSituacaoFilter(e.target.value)}
            >
              <option value="">Todas as SituaÃ§Ãµes</option>
              {uniqueSituacoes.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={listaFilter}
              onChange={(e) => setListaFilter(e.target.value)}
            >
              <option value="">Todas as Listas</option>
              {uniqueListas.map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="">Todos os Status</option>
              {uniqueStatuses.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead className="bg-gray-50 border-b border-gray-100">
                  <tr>
                    <th className="px-6 py-4">
                      <input
                        type="checkbox"
                        checked={
                          selectedEntries.length === filteredData.length &&
                          filteredData.length > 0
                        }
                        onChange={(e) => toggleSelectAll(e.target.checked)}
                      />
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Candidato
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Lista/Status
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Tipo/Bolsa
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Curso/Metodologia
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      DocumentaÃ§Ã£o
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Digitaliza
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      TCB
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600 flex items-center gap-4">
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkDelete}
                          className="text-rose-600 font-bold hover:underline"
                        >
                          excluir selecionados
                        </button>
                      )}
                      {selectedEntries.length > 0 && botConfig.url && (
                        <button
                          onClick={() => {
                            const selectedObjs = data.filter((g) =>
                              selectedEntries.includes(g.id),
                            );
                            const payloads = selectedObjs.map((item) => {
                              const isMatAcadOk =
                                item.numeroMatricula &&
                                item.numeroMatricula.trim().length > 0;
                              const type = isMatAcadOk
                                ? "fiesProuni_1"
                                : "fiesProuni_0";
                              const msgTemplate = whatsappMessages.find(
                                (m) =>
                                  m.tipo === type || m.tipo === "fiesProuni",
                              );
                              const text = msgTemplate
                                ? replaceMessageVariables(
                                    msgTemplate.texto,
                                    item,
                                  )
                                : `OlÃ¡ ${item.nome}, tudo bem?`;
                              return {
                                telefone: item.telefone,
                                message: text,
                                nome: item.nome,
                              };
                            });
                            onMassSendBot(payloads);
                            setSelectedEntries([]);
                          }}
                          className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                        >
                          <Bot size={14} /> Em Massa
                        </button>
                      )}
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filteredData.map((item) => (
                    <tr
                      key={item.id}
                      className="hover:bg-gray-50/50 transition-colors"
                    >
                      <td className="px-6 py-4">
                        <input
                          type="checkbox"
                          checked={selectedEntries.includes(item.id)}
                          onChange={(e) =>
                            toggleSelect(item.id, e.target.checked)
                          }
                        />
                      </td>
                      <td className="px-6 py-4">
                        <div className="font-medium text-gray-900">
                          {item.nome || "Sem nome"}
                        </div>
                        <div className="text-[10px] font-bold text-indigo-500">
                          Ranking: {item.posicaoRanking || "-"}
                        </div>
                        <div className="text-xs text-gray-500">
                          {formatCPF(item.cpf || "")}
                        </div>
                        <div className="text-xs text-gray-400">
                          {item.periodo}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <div className="text-sm font-bold text-indigo-600">
                          {item.lista || "-"}
                        </div>
                        <div className="text-[10px] text-slate-500 uppercase font-bold">
                          {item.status || "Sem Status"}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-xs font-bold ${item.tipo === "FIES" ? "bg-blue-100 text-blue-700" : "bg-orange-100 text-orange-700"}`}
                        >
                          {item.tipo}
                        </span>
                        <div className="text-xs text-gray-500 mt-1 font-bold">
                          {item.bolsa}
                        </div>
                        {item.situacao && (
                          <div className="text-[10px] text-gray-400 mt-1">
                            Sit.: {item.situacao}
                          </div>
                        )}
                        {item.cotaPPI && (
                          <div className="text-[10px] text-gray-400">
                            PPI: {item.cotaPPI}
                          </div>
                        )}
                      </td>
                      <td className="px-6 py-4">
                        <div className="text-sm text-gray-700">
                          {item.curso}
                        </div>
                        <div className="text-xs text-gray-500">
                          {item.metodologia}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-[10px] font-bold ${
                            item.docsEntreguesStatus === "Sim"
                              ? "bg-green-100 text-green-700"
                              : item.docsEntreguesStatus === "Parcial"
                                ? "bg-amber-100 text-amber-700"
                                : "bg-red-100 text-red-700"
                          }`}
                        >
                          {item.docsEntreguesStatus || "Pendente"}
                        </span>
                        <div className="text-[10px] text-slate-400 mt-1">
                          {item.documentosEntregues?.length || 0} docs
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-xs font-medium ${
                            item.digitalizaStatus === "ConcluÃ­do"
                              ? "bg-green-100 text-green-700"
                              : item.digitalizaStatus === "Em AnÃ¡lise"
                                ? "bg-amber-100 text-amber-700"
                                : "bg-gray-100 text-gray-600"
                          }`}
                        >
                          {item.digitalizaStatus}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        {item.tcbAssinado ? (
                          <CheckCircle2 className="text-green-500" size={20} />
                        ) : (
                          <Clock className="text-gray-300" size={20} />
                        )}
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center space-x-2">
                          <button
                            onClick={() => {
                              setEditingEntry(item);
                              setIsModalOpen(true);
                            }}
                            className="text-indigo-600 hover:text-indigo-800 font-medium text-sm p-2 hover:bg-indigo-50 rounded-lg transition-all"
                            title="Editar"
                          >
                            <Edit2 size={18} />
                          </button>
                          {item.telefone && (
                            <>
                              {botConfig.url && (
                                <button
                                  onClick={() => {
                                    const isMatAcadOk =
                                      item.numeroMatricula &&
                                      item.numeroMatricula.trim().length > 0;
                                    const type = isMatAcadOk
                                      ? "fiesProuni_1"
                                      : "fiesProuni_0";
                                    const msgObj = whatsappMessages.find(
                                      (m) =>
                                        m.tipo === type ||
                                        m.tipo === "fiesProuni",
                                    );
                                    const msg = replaceMessageVariables(
                                      msgObj
                                        ? msgObj.texto
                                        : `OlÃ¡ [nome], tudo bem?`,
                                      item,
                                    );
                                    onSendBot(item.telefone, msg);
                                  }}
                                  className="text-blue-600 hover:text-blue-800 p-2 hover:bg-blue-50 rounded-lg transition-all"
                                  title="Enviar pelo Bot ARGO'S"
                                >
                                  <Bot size={18} />
                                </button>
                              )}
                              <a
                                href={getWhatsAppUrl(
                                  item.telefone,
                                  (() => {
                                    const isMatAcadOk =
                                      item.numeroMatricula &&
                                      item.numeroMatricula.trim().length > 0;
                                    const type = isMatAcadOk
                                      ? "fiesProuni_1"
                                      : "fiesProuni_0";
                                    const msg = whatsappMessages.find(
                                      (m) =>
                                        m.tipo === type ||
                                        m.tipo === "fiesProuni",
                                    );
                                    if (msg)
                                      return replaceMessageVariables(
                                        msg.texto,
                                        item,
                                      );
                                    return `OlÃ¡ ${item.nome}, tudo bem?`;
                                  })(),
                                )}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-emerald-600 hover:text-emerald-800 p-2 hover:bg-emerald-50 rounded-lg transition-all"
                                title="Enviar WhatsApp"
                              >
                                <MessageSquare size={18} />
                              </a>
                            </>
                          )}
                          <button
                            onClick={() => handleDeleteIndividual(item.id)}
                            className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                            title="Excluir"
                          >
                            <Trash2 size={18} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {activeTab === "informacoes" && (
        <div className="space-y-6">
          {/* Filters for Vagas */}
          <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex flex-wrap gap-4">
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Semestre / PerÃ­odo
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasPeriodoFilter}
                onChange={(e) => setVagasPeriodoFilter(e.target.value)}
              >
                <option value="">Todos os Semestres</option>
                {periodos.map((p) => (
                  <option key={p.id} value={p.nome}>
                    {p.nome}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Metodologia
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasMetodologiaFilter}
                onChange={(e) => setVagasMetodologiaFilter(e.target.value)}
              >
                <option value="">Todas as Metodologias</option>
                {uniqueMetodologias.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Bolsa
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasBolsaFilter}
                onChange={(e) => setVagasBolsaFilter(e.target.value)}
              >
                <option value="">Todas as Bolsas</option>
                <option value="50%">50%</option>
                <option value="100%">100%</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
              <div className="p-3 bg-blue-500 rounded-xl text-white">
                <FileText size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">
                  Total de Vagas
                </p>
                <h3 className="text-2xl font-bold text-slate-900">
                  {vagasStats.totalVagas}
                </h3>
              </div>
            </div>
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4 border-l-4 border-l-emerald-500">
              <div className="p-3 bg-emerald-500 rounded-xl text-white">
                <ShieldCheck size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">
                  Bolsas 100%
                </p>
                <h3 className="text-2xl font-bold text-emerald-600">
                  {vagasStats.total100}
                </h3>
              </div>
            </div>
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4 border-l-4 border-l-blue-500">
              <div className="p-3 bg-blue-500 rounded-xl text-white">
                <BookOpen size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">Bolsas 50%</p>
                <h3 className="text-2xl font-bold text-blue-600">
                  {vagasStats.total50}
                </h3>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider">
                    <th className="p-4 font-bold border-b border-slate-100">
                      PerÃ­odo
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Cod. Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Turno
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Metodologia
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Bolsa
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Vagas
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-right">
                      AÃ§Ãµes
                    </th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-50">
                  {filteredVagas.length === 0 ? (
                    <tr>
                      <td
                        colSpan={8}
                        className="p-8 text-center text-slate-400"
                      >
                        Nenhuma vaga cadastrada.
                      </td>
                    </tr>
                  ) : (
                    filteredVagas.map((vaga) => (
                      <tr
                        key={vaga.id}
                        className="hover:bg-slate-50/50 transition-colors"
                      >
                        <td className="p-4 font-medium text-slate-700">
                          {vaga.periodo}
                        </td>
                        <td className="p-4 font-mono text-slate-500">
                          {vaga.codCurso}
                        </td>
                        <td className="p-4 font-medium text-slate-800">
                          {vaga.curso}
                        </td>
                        <td className="p-4 text-center text-slate-600">
                          {vaga.turno}
                        </td>
                        <td className="p-4 text-center text-slate-600">
                          {vaga.metodologia}
                        </td>
                        <td className="p-4 text-center">
                          <span
                            className={`inline-block px-2.5 py-1 rounded-full text-xs font-bold ${vaga.bolsa === "100%" ? "bg-emerald-100 text-emerald-700" : "bg-blue-100 text-blue-700"}`}
                          >
                            {vaga.bolsa}
                          </span>
                        </td>
                        <td className="p-4 text-center font-bold text-slate-800">
                          {vaga.vagas}
                        </td>
                        <td className="p-4 text-right">
                          <div className="flex justify-end gap-2">
                            <button
                              onClick={() => {
                                setEditingVaga(vaga);
                                setIsVagaModalOpen(true);
                              }}
                              className="text-slate-400 hover:text-blue-600 p-2 hover:bg-blue-50 rounded-lg transition-all"
                            >
                              <Edit2 size={18} />
                            </button>
                            <button
                              onClick={() => handleDeleteVaga(vaga.id)}
                              className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                            >
                              <Trash2 size={18} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Modal */}
      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center sticky top-0 bg-white z-10">
                <h3 className="text-xl font-bold text-gray-800">
                  {editingEntry
                    ? "Editar Registro"
                    : "Novo Cadastro Fies/Prouni"}
                </h3>
                <button
                  onClick={() => setIsModalOpen(false)}
                  className="text-gray-400 hover:text-gray-600"
                >
                  <X size={24} />
                </button>
              </div>
              <form onSubmit={handleSave} className="p-6 space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Nome Completo
                    </label>
                    <input
                      name="nome"
                      defaultValue={editingEntry?.nome}
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      CPF
                    </label>
                    <input
                      name="cpf"
                      value={cpfInput}
                      onChange={(e) =>
                        setCpfInput(formatCPF(e.target.value || ""))
                      }
                      required
                      placeholder="000.000.000-00"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Telefone
                    </label>
                    <input
                      name="telefone"
                      defaultValue={editingEntry?.telefone}
                      onChange={(e) => {
                        e.target.value = formatPhone(e.target.value);
                      }}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Email
                    </label>
                    <input
                      name="email"
                      type="email"
                      defaultValue={editingEntry?.email}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      EndereÃ§o
                    </label>
                    <input
                      name="endereco"
                      defaultValue={editingEntry?.endereco}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status
                    </label>
                    <select
                      name="status"
                      defaultValue={editingEntry?.status || "Pendente"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Aprovado">Aprovado</option>
                      <option value="Reprovado">Reprovado</option>
                      <option value="Em AnÃ¡lise">Em AnÃ¡lise</option>
                      <option value="Desistente">Desistente</option>
                      <option value="NÃ£o compareceu">NÃ£o compareceu</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Tipo
                    </label>
                    <select
                      name="tipo"
                      defaultValue={editingEntry?.tipo || "PROUNI"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="FIES">FIES</option>
                      <option value="PROUNI">PROUNI</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Bolsa
                    </label>
                    <select
                      name="bolsa"
                      defaultValue={editingEntry?.bolsa || "INTEGRAL"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="INTEGRAL">INTEGRAL</option>
                      <option value="PARCIAL">PARCIAL</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      SituaÃ§Ã£o
                    </label>
                    <select
                      name="situacao"
                      defaultValue={editingEntry?.situacao || "Candidato"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Candidato">Candidato</option>
                      <option value="Aluno (mesmo curso)">Aluno (mesmo curso)</option>
                      <option value="Aluno (outro curso)">Aluno (outro curso)</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Cota PPI
                    </label>
                    <select
                      name="cotaPPI"
                      defaultValue={editingEntry?.cotaPPI || "NÃ£o"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Sim">Sim</option>
                      <option value="NÃ£o">NÃ£o</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      PerÃ­odo
                    </label>
                    <input
                      name="periodo"
                      defaultValue={editingEntry?.periodo}
                      placeholder="Ex: 2025.1"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Lista
                    </label>
                    <input
                      name="lista"
                      defaultValue={editingEntry?.lista}
                      placeholder="Ex: Lista 1"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      PosiÃ§Ã£o no Ranking
                    </label>
                    <input
                      name="posicaoRanking"
                      defaultValue={editingEntry?.posicaoRanking}
                      placeholder="Ex: 15Âº"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Curso
                    </label>
                    <input
                      name="curso"
                      defaultValue={editingEntry?.curso}
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Metodologia
                    </label>
                    <input
                      name="metodologia"
                      defaultValue={editingEntry?.metodologia}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      InscriÃ§Ã£o Sales
                    </label>
                    <input
                      name="inscricaoSales"
                      defaultValue={editingEntry?.inscricaoSales}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      NÃºmero MatrÃ­cula
                    </label>
                    <input
                      name="numeroMatricula"
                      defaultValue={editingEntry?.numeroMatricula}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status Digitaliza
                    </label>
                    <select
                      name="digitalizaStatus"
                      defaultValue={
                        editingEntry?.digitalizaStatus || "NÃ£o Postado"
                      }
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="NÃ£o Postado">NÃ£o Postado</option>
                      <option value="Em AnÃ¡lise">Em AnÃ¡lise</option>
                      <option value="ConcluÃ­do">ConcluÃ­do</option>
                      <option value="Documento reprovado">
                        Documento reprovado
                      </option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status Documentos
                    </label>
                    <select
                      name="docsEntreguesStatus"
                      defaultValue={
                        editingEntry?.docsEntreguesStatus || "Pendente"
                      }
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Parcial">Parcial</option>
                      <option value="Sim">Sim (Tudo Entregue)</option>
                      <option value="NÃ£o compareceu">NÃ£o compareceu</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      SISPROUNI
                    </label>
                    <select
                      name="sisprouniStatus"
                      defaultValue={editingEntry?.sisprouniStatus || "Pendente"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Aprovado">Aprovado</option>
                      <option value="Reprovado">Reprovado</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      ResponsÃ¡vel Entrevista
                    </label>
                    <input
                      name="responsavelEntrevista"
                      defaultValue={
                        editingEntry?.responsavelEntrevista || profile.name
                      }
                      readOnly={!isAdmin}
                      className={`w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 ${!isAdmin ? "bg-slate-50 text-slate-500" : ""}`}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Data Entrevista
                    </label>
                    <input
                      name="dataEntrevista"
                      type="date"
                      defaultValue={
                        editingEntry?.dataEntrevista ||
                        new Date().toISOString().split("T")[0]
                      }
                      readOnly={!isAdmin}
                      className={`w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 ${!isAdmin ? "bg-slate-50 text-slate-500" : ""}`}
                    />
                  </div>
                  <div className="flex items-center space-x-2 pt-8">
                    <input
                      type="checkbox"
                      name="tcbAssinado"
                      defaultChecked={editingEntry?.tcbAssinado}
                      className="w-5 h-5 text-indigo-600 rounded focus:ring-indigo-500"
                    />
                    <label className="text-sm font-medium text-gray-700">
                      TCB Assinado
                    </label>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Documentos Entregues (separados por vÃ­rgula)
                  </label>
                  <input
                    name="documentos"
                    defaultValue={editingEntry?.documentosEntregues?.join(", ")}
                    placeholder="Ex: RG, CPF, Diploma"
                    className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    ObservaÃ§Ãµes / O que falta
                  </label>
                  <textarea
                    name="observacao"
                    defaultValue={editingEntry?.observacao}
                    rows={3}
                    className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div className="pt-4">
                  <button
                    type="submit"
                    className="w-full bg-indigo-600 text-white py-3 rounded-xl font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-200"
                  >
                    {editingEntry ? "Salvar AlteraÃ§Ãµes" : "Cadastrar Candidato"}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}

        {isVagaModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-lg"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-white rounded-t-2xl">
                <h3 className="text-xl font-bold text-gray-800">
                  {editingVaga ? "Editar Vaga" : "Nova Vaga"}
                </h3>
                <button
                  onClick={() => setIsVagaModalOpen(false)}
                  className="text-gray-400 hover:text-gray-600 transition-colors"
                >
                  <X size={24} />
                </button>
              </div>

              <form onSubmit={handleSaveVaga} className="p-6 space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      PerÃ­odo
                    </label>
                    <input
                      name="periodo"
                      required
                      defaultValue={editingVaga?.periodo}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Cod. Curso
                    </label>
                    <input
                      name="codCurso"
                      required
                      defaultValue={editingVaga?.codCurso}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Curso
                    </label>
                    <input
                      name="curso"
                      required
                      defaultValue={editingVaga?.curso}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Turno
                    </label>
                    <select
                      name="turno"
                      required
                      defaultValue={editingVaga?.turno}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="ManhÃ£">ManhÃ£</option>
                      <option value="Tarde">Tarde</option>
                      <option value="Noite">Noite</option>
                      <option value="Integral">Integral</option>
                      <option value="EAD">EAD</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Metodologia
                    </label>
                    <select
                      name="metodologia"
                      required
                      defaultValue={editingVaga?.metodologia}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="Presencial">Presencial</option>
                      <option value="EAD">EAD</option>
                      <option value="HÃ­brido">HÃ­brido</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Bolsa
                    </label>
                    <select
                      name="bolsa"
                      required
                      defaultValue={editingVaga?.bolsa}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="50%">50%</option>
                      <option value="100%">100%</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Vagas
                    </label>
                    <input
                      name="vagas"
                      type="number"
                      required
                      min="0"
                      defaultValue={editingVaga?.vagas}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                </div>
                <div className="pt-4">
                  <button
                    type="submit"
                    className="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 transition-colors shadow-lg shadow-blue-200"
                  >
                    {editingVaga ? "Salvar AlteraÃ§Ãµes" : "Cadastrar Vaga"}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

export default function App() {
  const [analysisSchemes, setAnalysisSchemes] = useState<AnalysisScheme[]>([]);
  const [user, setUser] = useState<User | null>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [currentView, setCurrentView] = useState(() => {
    const params = new URLSearchParams(window.location.search);
    return params.get("view") || "cadastro";
  });
  const [toast, setToast] = useState<{
    message: string;
    type: "success" | "error";
  } | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [insumosBaixas, setInsumosBaixas] = useState<InsumoBaixa[]>([]);
  const [isOnline, setIsOnline] = useState<boolean>(
    typeof navigator !== "undefined" ? navigator.onLine : true,
  );

  useEffect(() => {
    if (typeof window === "undefined") return;
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);

    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  // Data States
  const [salesContacts, setSalesContacts] = useState<SalesContact[]>([]);
  const [mensagensEnviadasLog, setMensagensEnviadasLog] = useState<MensagemEnviadaLog[]>([]);
  const [leads, setLeads] = useState<Lead[]>([]);
  const [users, setUsers] = useState<UserProfile[]>([]);
  const [bases, setBases] = useState<BaseEntry[]>([]);
  const [gap, setGap] = useState<GapEntry[]>([]);
  const [isencoes, setIsencoes] = useState<IsencaoEntry[]>([]);
  const [solicitacoesManutencao, setSolicitacoesManutencao] = useState<SolicitacaoManutencao[]>([]);
  const [fiesProuni, setFiesProuni] = useState<FiesProuniEntry[]>([]);
  const [fiesProuniVagas, setFiesProuniVagas] = useState<FiesProuniVaga[]>([]);
  const [campanhas, setCampanhas] = useState<Campanha[]>([]);
  const [bomDia, setBomDia] = useState<BomDiaCaptacao[]>([]);
  const [forecast, setForecast] = useState<ForecastCaptacao[]>([]);
  const [metaDia, setMetaDia] = useState<MetaDia[]>([]);
  const [metaSM, setMetaSM] = useState<MetaSM[]>([]);
  const [metaCursos, setMetaCursos] = useState<MetaCurso[]>([]);
  const [qgLigacoes, setQgLigacoes] = useState<QgLigacao[]>([]);
  const [planner, setPlanner] = useState<PlannerTask[]>([]);
  const [periodos, setPeriodos] = useState<PeriodoCaptacao[]>([]);
  const [calendarioAcoes, setCalendarioAcoes] = useState<CalendarioAcao[]>([]);
  const [empresasParceiras, setEmpresasParceiras] = useState<EmpresaParceira[]>(
    [],
  );
  const [controleConcorrencia, setControleConcorrencia] = useState<
    ControleConcorrencia[]
  >([]);
  const [whatsappMessages, setWhatsappMessages] = useState<WhatsAppMessage[]>(
    [],
  );
  const [activeWhatsappTemplates, setActiveWhatsappTemplates] = useState<
    Record<string, string>
  >({});
  const [links, setLinks] = useState<LinkUtil[]>([]);
  const [mapao, setMapao] = useState<MapaoAcademicoEntry[]>([]);
  const [basesDisparo, setBasesDisparo] = useState<BaseDisparoEntry[]>([]);
  const [basesRenovacao, setBasesRenovacao] = useState<BaseEntry[]>([]);
  const [cursos, setCursos] = useState<CursoDisponivel[]>([]);
  const uniqueUnidades = useMemo(() => {
    return Array.from(
      new Set((cursos || []).map((c) => c.nomeUnidade).filter(Boolean)),
    ).sort();
  }, [cursos]);
  const [pedidosCursos, setPedidosCursos] = useState<PedidoCursoEntry[]>([]);
  const [ligacoes, setLigacoes] = useState<Ligacao[]>([]);
  const [insumosPedidos, setInsumosPedidos] = useState<InsumoPedido[]>([]);
  const [insumosEstoque, setInsumosEstoque] = useState<InsumoEstoque[]>([]);
  const [insumosPedidosComercial, setInsumosPedidosComercial] = useState<
    InsumoPedidoComercial[]
  >([]);
  const [insumosEstoqueComercial, setInsumosEstoqueComercial] = useState<
    InsumoEstoqueComercial[]
  >([]);
  const [botConfig, setBotConfig] = useState<BotConfig>({
    url: "",
    active: false,
  });
  const [botStatuses, setBotStatuses] = useState<
    Record<
      string,
      {
        status: string;
        pairingCode?: string;
        qrCode?: string;
        qrUrl?: string;
        active?: boolean;
      }
    >
  >({});
  const [showInjectModal, setShowInjectModal] = useState(false);
  const [injectBotNumber, setInjectBotNumber] = useState("");
  const [injectSessionData, setInjectSessionData] = useState("");
  const [initialActionData, setInitialActionData] =
    useState<Partial<CalendarioAcao> | null>(null);
  const [activePopup, setActivePopup] = useState<{
    title: string;
    message: string;
  } | null>(null);
  const [isProfileModalOpen, setIsProfileModalOpen] = useState(false);
  const [massSendProgress, setMassSendProgress] = useState<{
    total: number;
    sent: number;
    active: boolean;
    info: string;
  }>({ total: 0, sent: 0, active: false, info: "" });
  const [isMassSendPaused, setIsMassSendPaused] = useState(false);
  const massSendControlRef = React.useRef({ paused: false, cancelled: false });

  const showToast = (
    message: string,
    type: "success" | "error" = "success",
  ) => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  const showPopup = (title: string, message: string) => {
    setActivePopup({ title, message });
  };

  const canView = (view: string) => {
    if (!profile) return false;
    if (
      profile.email === "canaldonutri@gmail.com" ||
      profile.email === "marcos.teixeira@estacio.br" ||
      profile.role === "Admin Master"
    ) {
      return true;
    }
    const isComercial =
      localStorage.getItem("servidor_selected") === "comercial";
    if (profile.role === ROLES.FINANCEIRO) {
      if (isComercial) {
        return ["controlePagamentos", "controleInsumosComercial"].includes(view);
      } else {
        return VIEW_PERMISSIONS[view]?.includes(profile.role) || false;
      }
    }
    return VIEW_PERMISSIONS[view]?.includes(profile.role) || false;
  };

  const callBotApi = async (
    path: string,
    options: { method?: "GET" | "POST"; body?: any } = {},
  ) => {
    // Determine the exact URL to fetch from, using the requested Railway API directly for send actions
    const directUrl =
      path === "/api/send"
        ? "https://argoscliente-production-170b.up.railway.app/api/send"
        : botConfig.url
          ? `${botConfig.url.endsWith("/") ? botConfig.url.slice(0, -1) : botConfig.url}${path}`
          : `https://argoscliente-production-170b.up.railway.app${path}`;

    const fetchOptions: RequestInit = {
      method: options.method || "GET",
      headers: {
        "Content-Type": "application/json",
      },
    };
    if (options.method === "POST" && options.body) {
      fetchOptions.body = JSON.stringify(options.body);
    }

    const response = await fetch(directUrl, fetchOptions);
    if (!response.ok) {
      const isJson = response.headers
        .get("content-type")
        ?.includes("application/json");
      const json = isJson ? await response.json().catch(() => ({})) : {};
      throw new Error(
        json.error ||
          json.message ||
          `Erro ao conectar ao Bot (${response.status})`,
      );
    }

    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      const text = await response.text();
      throw new Error(
        `O Bot no Railway retornou uma resposta inesperada (formato nÃ£o-JSON). O bot pode estar offline ou em reinicializaÃ§Ã£o.`,
      );
    }

    const resData = await response.json();

    // Support either direct raw JSON responses or wrapper structures with { success: boolean, data?: any }
    if (
      resData !== null &&
      typeof resData === "object" &&
      "success" in resData
    ) {
      if (!resData.success) {
        throw new Error(resData.data?.error || resData.error || `Falha no bot`);
      }
      return "data" in resData ? resData.data : resData;
    }

    return resData;
  };

  const sendAppWhatsApp = async (recipientPhone: string, message: string) => {
    let rawPhone = recipientPhone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }
    if (!rawPhone) return;

    try {
      const finalMessage =
        message +
        "\n\nPor favor nÃ£o responder nesse whatsapp. Pois ele Ã© apenas um numero de assistÃªncia de envio.";
      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: "5524993346717",
          number: rawPhone,
          message: finalMessage,
          force: true,
          manual: true,
        },
      });
      console.log(`WhatsApp sent to ${rawPhone} via bot 5524993346717`);
    } catch (err) {
      console.error("Error sending WhatsApp notification:", err);
    }
  };

  const sendAppTelegram = async (
    telegramHandleOrId: string,
    message: string,
  ) => {
    if (!telegramHandleOrId) return;
    const targetUrl = botConfig?.telegramBotUrl || "";
    const apiKey = botConfig?.telegramApiKey || "";
    if (!targetUrl) {
      console.log("Telegram Bot URL not configured in botConfig.");
      return;
    }
    try {
      const chatId = telegramHandleOrId.trim();
      const response = await fetch("/api/bot-proxy", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          targetUrl,
          method: "POST",
          headers: {
            "x-api-key": apiKey,
          },
          body: {
            chatId,
            mensagem: message,
          },
        }),
      });
      const resData = await response.json();
      if (!response.ok || !resData.success) {
        console.error(
          "Failed to send Telegram message at root:",
          resData.error || "Unknown error",
        );
      } else {
        console.log(`Telegram message sent to ${chatId}`);
      }
    } catch (err) {
      console.error("Error calling Telegram bot-proxy at root:", err);
    }
  };

  const handleSendBotMessage = async (
    telefone: string,
    message: string,
    contactName?: string,
  ) => {
    const currentBotNumber = profile?.botNumber;
    let safeBotNumber = currentBotNumber
      ? currentBotNumber.replace(/\D/g, "")
      : "";

    // Auto-fallback: if the user's personal bot number is offline, not active,
    // or not set, look for any online bot session in the system to route the dispatch.
    const isUserBotOnline =
      safeBotNumber && (botStatuses as any)[safeBotNumber]?.status === "online";

    if (!isUserBotOnline) {
      const firstOnlineBot = Object.entries(botStatuses).find(
        ([_, info]) => (info as any)?.status === "online",
      )?.[0];
      if (firstOnlineBot) {
        console.log(
          `Fallback bot activated: Routing message via active online session: ${firstOnlineBot}`,
        );
        safeBotNumber = firstOnlineBot;
      } else if (!safeBotNumber) {
        showToast(
          "VocÃª ainda nÃ£o tem um nÃºmero de WhatsApp configurado (AdministraÃ§Ã£o -> GestÃ£oPro) e nenhum bot estÃ¡ ativo no momento.",
          "error",
        );
        return;
      }
    }

    // Format phone: remove non-numeric, strip leading zero if present
    let rawPhone = telefone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    // Add country code if not present and has standard length
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }

    try {
      const isTargetBot = safeBotNumber === "5524993346717";
      const finalMessage = isTargetBot
        ? message +
          "\n\nPor favor nÃ£o responder nesse whatsapp. Pois ele Ã© apenas um numero de assistÃªncia de envio."
        : message;

      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: safeBotNumber,
          number: rawPhone,
          message: finalMessage,
          contactName: contactName || "",
          force: true,
          manual: true,
        },
      });
      showToast("Mensagem enviada com sucesso pelo Bot ARGO'S!");

      // Log to CRM
      try {
        let sentiment = "Neutro";
        try {
          const res = await fetch("/api/crm/sentiment", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: message }),
          });
          const data = await res.json();
          if (data.success && data.sentiment) {
            sentiment = data.sentiment;
          }
        } catch (e) {
          console.error("Error analyzing sentiment:", e);
        }

        const msgData = {
          text: message,
          senderId: profile?.uid || "system",
          senderName: profile?.name || "System",
          senderRole: profile?.role || "",
          receiverPhone: rawPhone,
          timestamp: serverTimestamp(),
          type: "sent",
          status: "sent",
        };
        await addDoc(collection(db, COLLECTIONS.MESSAGES), msgData);

        // Update or Create conversation
        await setDoc(
          doc(db, COLLECTIONS.CONVERSATIONS, rawPhone),
          {
            contactPhone: rawPhone,
            contactName: contactName || rawPhone,
            lastMessage: message,
            lastMessageTimestamp: serverTimestamp(),
            unreadCount: 0,
            unidade: profile?.unidade || "",
            sentiment,
          },
          { merge: true },
        );
      } catch (err) {
        console.error("Error logging message to CRM:", err);
      }

      // Automatic Status Transition Logic upon message sent
      try {
        const phonesMatch = (p1?: string, p2?: string): boolean => {
          if (!p1 || !p2) return false;
          const c1 = p1.replace(/\D/g, "");
          const c2 = p2.replace(/\D/g, "");
          if (c1 === c2) return true;
          const s1 = c1.startsWith("55")
            ? c1.substring(2)
            : c1.startsWith("0")
              ? c1.substring(1)
              : c1;
          const s2 = c2.startsWith("55")
            ? c2.substring(2)
            : c2.startsWith("0")
              ? c2.substring(1)
              : c2;
          if (s1 === s2) return true;
          if (s1.length >= 8 && s2.length >= 8) {
            const last8_1 = s1.slice(-8);
            const last8_2 = s2.slice(-8);
            const ddd1 = s1.substring(0, 2);
            const ddd2 = s2.substring(0, 2);
            if (last8_1 === last8_2 && ddd1 === ddd2) return true;
          }
          return false;
        };

        const matchedLeads = leads.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedBases = bases.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedBasesRenovacao = basesRenovacao.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedFiesProuni = fiesProuni.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );

        const existsInGap = gap.some((g) => {
          if (phonesMatch(g.telefone, telefone)) return true;
          const matchedCpf =
            matchedLeads.find((l) => l.cpf)?.cpf ||
            matchedBases.find((b) => b.cpf)?.cpf ||
            matchedBasesRenovacao.find((br) => br.cpf)?.cpf ||
            matchedFiesProuni.find((fp) => fp.cpf)?.cpf;
          if (matchedCpf && g.cpf) {
            const c1 = matchedCpf.replace(/\D/g, "");
            const c2 = g.cpf.replace(/\D/g, "");
            if (c1 && c1 === c2) return true;
          }
          return false;
        });

        // 1. Process matched LEADS
        for (const lead of matchedLeads) {
          if (existsInGap) {
            if (lead.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.LEADS, lead.id), {
                status: "Convertido",
              });
            }
          } else if (lead.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.LEADS, lead.id), {
              status: "Sem retorno",
            });
          }
        }

        // 2. Process matched BASES
        for (const entry of matchedBases) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), {
                status: "Convertido",
              });
            }
          } else if (entry.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        // 3. Process matched BASES_RENOVACAO
        for (const entry of matchedBasesRenovacao) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.BASES_RENOVACAO, entry.id), {
                status: "Convertido",
              });
            }
          } else if (entry.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.BASES_RENOVACAO, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        // 4. Process matched FIES_PROUNI
        for (const entry of matchedFiesProuni) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.FIES_PROUNI, entry.id), {
                status: "Convertido",
              });
            }
          } else if (
            entry.status &&
            entry.status.toLowerCase() === "pendente"
          ) {
            await updateDoc(doc(db, COLLECTIONS.FIES_PROUNI, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        let tipoContato = "outro";
        let baseName = "";
        if (matchedLeads.length > 0) {
          tipoContato = "leads";
        } else if (matchedBases.length > 0) {
          tipoContato = "bases";
          baseName = matchedBases[0].nomeBase;
        } else if (matchedBasesRenovacao.length > 0) {
          tipoContato = "bases_renovacao";
          baseName = matchedBasesRenovacao[0].nomeBase;
        } else if (matchedFiesProuni.length > 0) {
          tipoContato = "fies_prouni";
        } else if (existsInGap) {
          tipoContato = "gap";
        }

        await addDoc(collection(db, COLLECTIONS.BOT_REPORTS), {
          userId: profile?.uid || "unknown",
          userName: profile?.nome || "UsuÃ¡rio Desconhecido",
          userRole: profile?.role || "unknown",
          telefone,
          tipoContato,
          baseName,
          sentAt: serverTimestamp(),
        });

        const logCandidateName =
          contactName ||
          matchedLeads[0]?.nome ||
          matchedBases[0]?.nome ||
          matchedBasesRenovacao[0]?.nome ||
          matchedFiesProuni[0]?.nome ||
          telefone;
        const logCurso =
          matchedLeads[0]?.cursoInteresse ||
          matchedBases[0]?.curso ||
          matchedBasesRenovacao[0]?.curso ||
          matchedFiesProuni[0]?.curso ||
          "NÃ£o informado";
        const logBase =
          baseName ||
          (matchedLeads.length > 0
            ? matchedLeads[0].acao || "HistÃ³rico Leads"
            : tipoContato);

        await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
          leadId:
            matchedLeads[0]?.id ||
            matchedBases[0]?.id ||
            matchedBasesRenovacao[0]?.id ||
            matchedFiesProuni[0]?.id ||
            "",
          nome: logCandidateName,
          telefone: rawPhone || telefone,
          curso: logCurso,
          base: logBase,
          tipoEnvio: "bot_automatico",
          dataHora: serverTimestamp(),
          usuarioId: profile?.uid || "system",
          usuarioNome: profile?.nome || profile?.name || "Bot ARGO'S",
          unidade: profile?.unidade || "",
        });
      } catch (statusErr: any) {
        console.error(
          "[Auto Status Update] Failed to update statuses or log report:",
          statusErr,
        );
      }
    } catch (err: any) {
      showToast(`Erro ao enviar mensagem: ${err.message}`, "error");
    }
  };

  const sendSilentWhatsApp = async (telefone: string, message: string) => {
    const currentBotNumber = profile?.botNumber;
    let safeBotNumber = currentBotNumber
      ? currentBotNumber.replace(/\D/g, "")
      : "";

    const isUserBotOnline =
      safeBotNumber && (botStatuses as any)[safeBotNumber]?.status === "online";

    if (!isUserBotOnline) {
      const firstOnlineBot = Object.entries(botStatuses).find(
        ([_, info]) => (info as any)?.status === "online",
      )?.[0];
      if (firstOnlineBot) {
        safeBotNumber = firstOnlineBot;
      } else if (!safeBotNumber) {
        return;
      }
    }

    let rawPhone = telefone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }

    try {
      const isTargetBot = safeBotNumber === "5524993346717";
      const finalMessage = isTargetBot
        ? message +
          "\n\nPor favor nÃ£o responder nesse whatsapp. Pois ele Ã© apenas um numero de assistÃªncia de envio."
        : message;

      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: safeBotNumber,
          number: rawPhone,
          message: finalMessage,
          force: true,
          manual: true,
        },
      });
    } catch (err) {
      console.error(err);
    }
  };

  const handleMassSendBotMessages = async (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => {
    if (massSendProgress.active) {
      showToast("JÃ¡ existe um envio em massa em andamento.", "error");
      return;
    }

    if (messages.length === 0) return;
    if (
      !window.confirm(
        `Deseja iniciar o envio em massa via bot para ${messages.length} contatos?`,
      )
    )
      return;

    massSendControlRef.current = { paused: false, cancelled: false };
    setIsMassSendPaused(false);

    setMassSendProgress({
      total: messages.length,
      sent: 0,
      active: true,
      info: "Iniciando...",
    });

    const waitWithCheck = async (seconds: number, labelPrefix: string) => {
      for (let s = 0; s < seconds; s++) {
        if (massSendControlRef.current.cancelled) return;
        while (
          massSendControlRef.current.paused &&
          !massSendControlRef.current.cancelled
        ) {
          setMassSendProgress((prev) => ({
            ...prev,
            info: `RobÃ´ Pausado... (${prev.sent}/${messages.length})`,
          }));
          await new Promise((resolve) => setTimeout(resolve, 500));
        }
        if (massSendControlRef.current.cancelled) return;
        const remaining = seconds - s;
        setMassSendProgress((prev) => ({
          ...prev,
          info: `${labelPrefix} (${remaining}s restantes)... (${prev.sent}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    };

    let sentCount = 0;
    for (let i = 0; i < messages.length; i++) {
      if (massSendControlRef.current.cancelled) {
        break;
      }

      while (
        massSendControlRef.current.paused &&
        !massSendControlRef.current.cancelled
      ) {
        setMassSendProgress((prev) => ({
          ...prev,
          info: `RobÃ´ Pausado... (${sentCount}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 500));
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      if (i > 0) {
        if (sentCount % 5 === 0) {
          await waitWithCheck(120, "Pausa de 2 min");
        } else {
          await waitWithCheck(30, "Aguardando cooldown");
        }
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      while (
        massSendControlRef.current.paused &&
        !massSendControlRef.current.cancelled
      ) {
        setMassSendProgress((prev) => ({
          ...prev,
          info: `RobÃ´ Pausado... (${sentCount}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 500));
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      setMassSendProgress((prev) => ({
        ...prev,
        sent: sentCount,
        info: `Enviando... (${sentCount + 1}/${messages.length})`,
      }));

      try {
        await handleSendBotMessage(
          messages[i].telefone,
          messages[i].message,
          messages[i].nome,
        );
      } catch (e) {
        console.error("Error sending bot message in mass: ", e);
      }
      sentCount++;
      setMassSendProgress((prev) => ({
        ...prev,
        sent: sentCount,
      }));
    }

    const wasCancelled = massSendControlRef.current.cancelled;
    setMassSendProgress({ total: 0, sent: 0, active: false, info: "" });
    setIsMassSendPaused(false);

    if (wasCancelled) {
      showToast("Envio em massa cancelado pelo usuÃ¡rio.", "error");
    } else {
      showToast("Envio em massa concluÃ­do!", "success");
    }
  };

  // Subscribe to Analysis Schemes
  useEffect(() => {
    const q = collection(db, COLLECTIONS.CRESCIMENTO_ANUAL);
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const list = snapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      })) as AnalysisScheme[];
      setAnalysisSchemes(list);
    }, (err) => {
      console.log("Crescimento Anual snapshot error:", err);
    });
    return () => unsubscribe();
  }, []);

  const handleSaveAnalysisScheme = async (scheme: Partial<AnalysisScheme>) => {
    try {
      if (scheme.id) {
        const { id, ...data } = scheme;
        await updateDoc(doc(db, COLLECTIONS.CRESCIMENTO_ANUAL, id), {
          ...data,
          updatedAt: serverTimestamp(),
        });
        showToast("AnÃ¡lise salva com sucesso!");
      } else {
        await addDoc(collection(db, COLLECTIONS.CRESCIMENTO_ANUAL), {
          ...scheme,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        });
        showToast("AnÃ¡lise criada com sucesso!");
      }
    } catch (err: any) {
      console.error("Error saving analysis scheme:", err);
      showToast("Erro ao salvar anÃ¡lise.", "error");
    }
  };

  const handleDeleteAnalysisScheme = async (id: string) => {
    if (window.confirm("Deseja excluir esta anÃ¡lise?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.CRESCIMENTO_ANUAL, id));
        showToast("AnÃ¡lise excluÃ­da.");
      } catch (err: any) {
        console.error("Error deleting analysis scheme:", err);
        showToast("Erro ao excluir anÃ¡lise.", "error");
      }
    }
  };

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        try {
          // 1. Try to get profile by UID
          let userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));

          if (!userDoc.exists()) {
            // 2. If not found by UID, try to find by email (for pre-registered users)
            const q = query(
              collection(db, COLLECTIONS.USERS),
              where("email", "==", user.email),
            );
            const querySnap = await getDocs(q);

            if (!querySnap.empty) {
              // Found by email, use this document
              const existingDoc = querySnap.docs[0];
              const data = existingDoc.data();

              // If the document ID is not the UID, we should ideally migrate it
              // but for now we'll just use it. Wait, if we use it, rules might fail
              // because rules expect path/.../users/{uid}.
              // So we MUST migrate it to a document with UID as ID.
              await setDoc(doc(db, COLLECTIONS.USERS, user.uid), {
                ...data,
                uid: user.uid,
                updatedAt: serverTimestamp(),
              });

              // Delete the old document if it had a different ID
              if (existingDoc.id !== user.uid) {
                try {
                  await deleteDoc(doc(db, COLLECTIONS.USERS, existingDoc.id));
                } catch (e) {
                  console.warn(
                    "Could not delete old user document, likely due to rules. Skipping.",
                    e,
                  );
                }
              }

              userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));
            } else {
              // 3. Create default profile if not exists at all
              let role = ROLES.PROMOTOR;
              let servidor: "principal" | "comercial" = "principal";
              let name = user.email!.split("@")[0];

              if (user.displayName) {
                const parts = user.displayName.split("|");
                name = parts[0] || name;
                if (parts.length > 1 && parts[1] === "comercial") {
                  servidor = "comercial";
                  role = "Promotor/rua" as any;
                }
              }

              if (
                user.email === "marcos.teixeira@estacio.br" ||
                user.email === "canaldonutri@gmail.com"
              ) {
                role = ROLES.ADMIN_MASTER;
              } else {
                const allUsers = await getDocs(
                  query(collection(db, COLLECTIONS.USERS), limit(1)),
                );
                if (allUsers.empty) {
                  role = (
                    servidor === "comercial"
                      ? "Gerente Comercial (Comercial)"
                      : ROLES.LIDER_FDV
                  ) as any;
                }
              }

              const newProfile = {
                uid: user.uid,
                email: user.email!,
                name,
                role,
                servidor,
                mustChangePassword: false, // Default for self-signup
                createdAt: serverTimestamp(),
                dashboardWidgets: { stats: true, links: true, planner: true },
              };
              await setDoc(doc(db, COLLECTIONS.USERS, user.uid), newProfile);
              userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));
            }
          }

          if (userDoc.exists()) {
            const data = userDoc.data() as UserProfile;
            if (
              data.email === "marcos.teixeira@estacio.br" ||
              data.email === "canaldonutri@gmail.com"
            ) {
              data.role = ROLES.ADMIN_MASTER;
            }
            setProfile({ uid: user.uid, ...data } as UserProfile);
          }
          setUser(user);
        } catch (error: any) {
          console.error("Error fetching/creating profile details:", {
            code: error.code,
            message: error.message,
            stack: error.stack,
          });
          showToast(`Erro ao carregar perfil: ${error.message}`, "error");
          setUser(null);
          setProfile(null);
        }
      } else {
        setUser(null);
        setProfile(null);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;

    // Listeners for users require auth
    let unsubUsers = () => {};
    if (user && profile) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let usersQuery = query(collection(db, COLLECTIONS.USERS));
      if (isRestricted) {
        usersQuery = query(
          usersQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubUsers = onSnapshot(
        usersQuery,
        (snap) => {
          setUsers(
            snap.docs.map((d) => ({ uid: d.id, ...d.data() }) as UserProfile),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.USERS),
      );
    }

    let unsubPlanner = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubPlanner = onSnapshot(
        collection(db, COLLECTIONS.PLANNER),
        (snap) => {
          setPlanner(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as PlannerTask),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.PLANNER),
      );
    }

    let unsubLinks = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubLinks = onSnapshot(
        collection(db, COLLECTIONS.LINKS),
        (snap) => {
          setLinks(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as LinkUtil),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.LINKS),
      );
    }

    let unsubSalesContacts = () => {};
    let unsubLeads = () => {};
    let unsubMensagensEnviadas = () => {};
    if (profile) {
      let leadsQuery;
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      if (
        [
          ROLES.ADMIN_MASTER,
          ROLES.LIDER_FDV,
          ROLES.SALA_MATRICULA,
          ROLES.QG,
        ].includes(profile.role)
      ) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_UNIDADE) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("unidade", "==", profile.unidade || ""),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_COMERCIAL_COMERCIAL) {
        // Gerente Comercial (Comercial) ver everything in Comercial
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("servidor", "==", "comercial"),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.FDV_COMERCIAL) {
        // FDV (Comercial) sees their own leads or those from their linked promontors.
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("linkadoA", "==", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.FDV) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("linkadoA", "==", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_COMERCIAL) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("promotorRole", "in", [ROLES.PROMOTOR, ROLES.FDV]),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.PROMOTOR ||
        profile.role === ROLES.PROMOTOR_RUA
      ) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("promotorId", "==", user!.uid),
          orderBy("createdAt", "desc"),
        );
      } else {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("promotorId", "==", "none"),
          orderBy("createdAt", "desc"),
        );
      }

      if (isRestricted) {
        if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          // Already restricted by promotorId/linkadoA above, but we keep unit filter to be safe or bypass it.
          // The user said ONLY what they or linked promotor filled.
          // If we add the unit filter, it might exclude their own leads if they are in a different unit (unlikely).
          // But to be strict with "SÃ“ PODERÃ VE", we keep the current query which is already restricted to UID.
        } else {
          leadsQuery = query(
            leadsQuery,
            where("unidade", "==", profile.unidade || "Matriz"),
          );
        }
      }

      unsubLeads = onSnapshot(
        leadsQuery,
        (snap) => {
          setLeads(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Lead));
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.LEADS),
      );
      
      unsubSalesContacts = onSnapshot(
        query(collection(db, COLLECTIONS.SALES_CONTACTS), orderBy("createdAt", "desc")),
        (snap) => {
          setSalesContacts(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as SalesContact));
        },
        (err) => console.error("Error loading sales contacts:", err)
      );

      unsubMensagensEnviadas = onSnapshot(
        collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG),
        (snap) => {
          setMensagensEnviadasLog(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MensagemEnviadaLog));
        },
        (err) => console.error("Error loading mensagens enviadas log:", err)
      );
    }

    let unsubBases = () => {};
    if (profile && VIEW_PERMISSIONS.bases.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let basesQuery;
      if (isRestricted) {
        if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          basesQuery = query(
            collection(db, COLLECTIONS.BASES),
            or(
              where("promotorId", "==", user!.uid),
              where("linkadoA", "==", user!.uid),
            ),
            orderBy("createdAt", "desc"),
          );
        } else {
          basesQuery = query(
            collection(db, COLLECTIONS.BASES),
            where("unidade", "==", profile.unidade || "Matriz"),
            orderBy("createdAt", "desc"),
          );
        }
      } else {
        basesQuery = query(
          collection(db, COLLECTIONS.BASES),
          orderBy("createdAt", "desc"),
        );
      }

      unsubBases = onSnapshot(
        basesQuery,
        (snap) => {
          setBases(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BaseEntry),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.BASES),
      );
    }

    let unsubGap = () => {};
    if (profile && VIEW_PERMISSIONS.gap.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let gapQuery = query(
        collection(db, COLLECTIONS.GAP),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        gapQuery = query(
          gapQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubGap = onSnapshot(
        gapQuery,
        (snap) => {
          setGap(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as GapEntry));
        },
        (err) => handleFirestoreError(err, OperationType.LIST, COLLECTIONS.GAP),
      );
    }

    let unsubIsencoes = () => {};
    let unsubSolicitacoesManutencao = () => {};
    if (profile && VIEW_PERMISSIONS.isencoes.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let isencoesQuery = query(
        collection(db, COLLECTIONS.ISENCOES),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        isencoesQuery = query(
          isencoesQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubIsencoes = onSnapshot(
        isencoesQuery,
        (snap) => {
          setIsencoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as IsencaoEntry),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.ISENCOES),
      );

      unsubSolicitacoesManutencao = onSnapshot(
        collection(db, COLLECTIONS.SOLICITACOES_MANUTENCAO),
        (snap) => {
          setSolicitacoesManutencao(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as SolicitacaoManutencao),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.SOLICITACOES_MANUTENCAO),
      );
    }

    let unsubPedidosCursos = () => {};
    if (
      profile &&
      (VIEW_PERMISSIONS.historico.includes(profile.role) ||
        VIEW_PERMISSIONS.relatorios.includes(profile.role))
    ) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let pcQuery = query(
        collection(db, COLLECTIONS.PEDIDO_CURSOS),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        pcQuery = query(
          pcQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubPedidosCursos = onSnapshot(
        pcQuery,
        (snap) => {
          setPedidosCursos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as PedidoCursoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.PEDIDO_CURSOS,
          ),
      );
    }

    let unsubFiesProuni = () => {};
    let unsubFiesProuniVagas = () => {};
    if (profile && VIEW_PERMISSIONS.fiesProuni.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        profile.role !== ROLES.SSA &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let fpQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI),
        orderBy("createdAt", "desc"),
      );
      let fpvQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI_VAGAS),
        orderBy("createdAt", "desc"),
      );

      if (isRestricted) {
        fpQuery = query(
          fpQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
        fpvQuery = query(
          fpvQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubFiesProuni = onSnapshot(
        fpQuery,
        (snap) => {
          setFiesProuni(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as FiesProuniEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.FIES_PROUNI,
          ),
      );
      unsubFiesProuniVagas = onSnapshot(
        fpvQuery,
        (snap) => {
          setFiesProuniVagas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as FiesProuniVaga),
          );
        },
        (err) => console.error("Error fetching FIES_PROUNI_VAGAS:", err),
      );
    }

    let unsubCampanhas = () => {};
    if (profile && VIEW_PERMISSIONS.campanhas.includes(profile.role)) {
      unsubCampanhas = onSnapshot(
        collection(db, COLLECTIONS.CAMPANHAS),
        (snap) => {
          setCampanhas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Campanha),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.CAMPANHAS),
      );
    }

    let unsubBomDia = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubBomDia = onSnapshot(
        collection(db, COLLECTIONS.BOM_DIA),
        (snap) => {
          setBomDia(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BomDiaCaptacao),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.BOM_DIA),
      );
    }

    let unsubForecast = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubForecast = onSnapshot(
        collection(db, COLLECTIONS.FORECAST),
        (snap) => {
          setForecast(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as ForecastCaptacao,
            ),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.FORECAST),
      );
    }

    let unsubMetaDia = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaDia = onSnapshot(
        collection(db, COLLECTIONS.META_DIA),
        (snap) => {
          setMetaDia(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaDia),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_DIA),
      );
    }

    let unsubMetaSM = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaSM = onSnapshot(
        collection(db, COLLECTIONS.META_SM),
        (snap) => {
          setMetaSM(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaSM),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_SM),
      );
    }

    let unsubMetaCursos = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaCursos = onSnapshot(
        collection(db, COLLECTIONS.META_CURSOS),
        (snap) => {
          setMetaCursos(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaCurso),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_CURSOS),
      );
    }

    let unsubQgLigacoes = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubQgLigacoes = onSnapshot(
        collection(db, COLLECTIONS.QG_LIGACOES),
        (snap) => {
          setQgLigacoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as QgLigacao),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.QG_LIGACOES,
          ),
      );
    }

    let unsubPeriodos = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubPeriodos = onSnapshot(
        collection(db, COLLECTIONS.PERIODO_CAPTACAO),
        (snap) => {
          setPeriodos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as PeriodoCaptacao,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.PERIODO_CAPTACAO,
          ),
      );
    }

    let unsubCalendario = () => {};
    if (
      profile &&
      (VIEW_PERMISSIONS.calendario.includes(profile.role) ||
        VIEW_PERMISSIONS.controlePagamentos.includes(profile.role) ||
        canView("controlePagamentos"))
    ) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let calendarioQuery;
      if (
        [
          ROLES.ADMIN_MASTER,
          ROLES.LIDER_FDV,
          ROLES.SALA_MATRICULA,
          ROLES.GESTOR_UNIDADE,
          ROLES.GESTOR_COMERCIAL,
          ROLES.FINANCEIRO,
          ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
        ].includes(profile.role)
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.FDV ||
        profile.role === ROLES.FDV_COMERCIAL
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          or(
            where("creatorId", "==", user!.uid),
            where("creatorRole", "==", ROLES.PROMOTOR),
            where("creatorRole", "==", ROLES.PROMOTOR_RUA),
            where("colaboradorId", "==", user!.uid),
            where("colaboradoresIds", "array-contains", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.PROMOTOR ||
        profile.role === ROLES.PROMOTOR_RUA
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          or(
            where("creatorId", "==", user!.uid),
            where("colaboradorId", "==", user!.uid),
            where("promotoresSelecionados", "array-contains", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          where("creatorId", "==", "none"),
          orderBy("createdAt", "desc"),
        );
      }

      if (isRestricted) {
        calendarioQuery = query(
          calendarioQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubCalendario = onSnapshot(
        calendarioQuery,
        (snap) => {
          setCalendarioAcoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as CalendarioAcao),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.CALENDARIO_ACOES,
          ),
      );
    }

    let unsubEmpresas = () => {};
    if (profile && VIEW_PERMISSIONS.empresas.includes(profile.role)) {
      let empresasQuery = query(collection(db, COLLECTIONS.EMPRESAS_PARCEIRAS));

      const isRestricted = ![
        ROLES.ADMIN_MASTER,
        ROLES.GESTOR_COMERCIAL,
        ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
        ROLES.FINANCEIRO,
      ].includes(profile.role);

      if (isRestricted) {
        if (profile.role === ROLES.GESTOR_UNIDADE) {
          empresasQuery = query(
            empresasQuery,
            where(
              "unidadesVinculadas",
              "array-contains",
              profile.unidade || "",
            ),
          );
        } else if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          empresasQuery = query(
            empresasQuery,
            or(
              where(
                "unidadesVinculadas",
                "array-contains",
                profile.unidade || "",
              ),
              where("consultorId", "==", user!.uid),
              where("creatorId", "==", user!.uid),
            ),
          );
        }
      }

      unsubEmpresas = onSnapshot(
        empresasQuery,
        (snap) => {
          setEmpresasParceiras(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as EmpresaParceira,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.EMPRESAS_PARCEIRAS,
          ),
      );
    }

    let unsubControleConcorrencia = () => {};
    if (
      profile &&
      VIEW_PERMISSIONS.controleConcorrencia.includes(profile.role)
    ) {
      unsubControleConcorrencia = onSnapshot(
        collection(db, COLLECTIONS.CONTROLE_CONCORRENCIA),
        (snap) => {
          setControleConcorrencia(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as ControleConcorrencia,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.CONTROLE_CONCORRENCIA,
          ),
      );
    }

    let unsubWhatsApp = () => {};
    if (user) {
      unsubWhatsApp = onSnapshot(
        collection(db, COLLECTIONS.WHATSAPP_MESSAGES),
        (snap) => {
          setWhatsappMessages(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as WhatsAppMessage,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.WHATSAPP_MESSAGES,
          ),
      );
    }

    let unsubMapao = () => {};
    if (profile && VIEW_PERMISSIONS.mapao.includes(profile.role)) {
      unsubMapao = onSnapshot(
        collection(db, COLLECTIONS.MAPAO_ACADEMICO),
        (snap) => {
          setMapao(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as MapaoAcademicoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.MAPAO_ACADEMICO,
          ),
      );
    }

    let unsubBasesDisparo = () => {};
    if (profile && VIEW_PERMISSIONS.basesDisparo.includes(profile.role)) {
      unsubBasesDisparo = onSnapshot(
        collection(db, COLLECTIONS.BASES_DISPARO),
        (snap) => {
          setBasesDisparo(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as BaseDisparoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.BASES_DISPARO,
          ),
      );
    }

    let unsubBasesRenovacao = () => {};
    if (profile && VIEW_PERMISSIONS.basesRenovacao.includes(profile.role)) {
      unsubBasesRenovacao = onSnapshot(
        collection(db, COLLECTIONS.BASES_RENOVACAO),
        (snap) => {
          setBasesRenovacao(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BaseEntry),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.BASES_RENOVACAO,
          ),
      );
    }

    let unsubCursos = () => {};
    if (profile && VIEW_PERMISSIONS.cursos.includes(profile.role)) {
      unsubCursos = onSnapshot(
        collection(db, COLLECTIONS.CURSOS),
        (snap) => {
          setCursos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as CursoDisponivel,
            ),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.CURSOS),
      );
    }

    let unsubInsumosPedidos = () => {};
    let unsubInsumosEstoque = () => {};
    if (profile && VIEW_PERMISSIONS.controleInsumos.includes(profile.role)) {
      unsubInsumosPedidos = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_PEDIDOS),
        (snap) => {
          setInsumosPedidos(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoPedido),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_PEDIDOS,
          ),
      );

      unsubInsumosEstoque = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_ESTOQUE),
        (snap) => {
          setInsumosEstoque(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoEstoque),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_ESTOQUE,
          ),
      );
    }

    let unsubInsumosPedidosComercial = () => {};
    let unsubInsumosEstoqueComercial = () => {};
    let unsubInsumosBaixas = () => {};
    if (
      profile &&
      VIEW_PERMISSIONS.controleInsumosComercial.includes(profile.role)
    ) {
      const isGerenteOrAdmin =
        profile.role === ROLES.ADMIN_MASTER ||
        profile.role === "Admin Master" ||
        profile.role === "Gerente Comercial (Comercial)" ||
        profile.role === "Gestor Comercial";

      const qPedidosComercial = isGerenteOrAdmin
        ? collection(db, COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL)
        : query(
            collection(db, COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL),
            where("solicitanteId", "==", profile.uid),
          );

      const qEstoqueComercial = isGerenteOrAdmin
        ? collection(db, COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL)
        : query(
            collection(db, COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL),
            where("ownerId", "==", profile.uid),
          );

      unsubInsumosPedidosComercial = onSnapshot(
        qPedidosComercial,
        (snap) => {
          setInsumosPedidosComercial(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as InsumoPedidoComercial,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL,
          ),
      );

      unsubInsumosEstoqueComercial = onSnapshot(
        qEstoqueComercial,
        (snap) => {
          setInsumosEstoqueComercial(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as InsumoEstoqueComercial,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL,
          ),
      );

      unsubInsumosBaixas = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_BAIXAS),
        (snap) => {
          setInsumosBaixas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoBaixa),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_BAIXAS,
          ),
      );
    }

    const unsubLigacoes = onSnapshot(
      collection(db, COLLECTIONS.CONTROLE_LIGACOES),
      (snap) => {
        setLigacoes(
          snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Ligacao),
        );
      },
      (err) =>
        handleFirestoreError(
          err,
          OperationType.LIST,
          COLLECTIONS.CONTROLE_LIGACOES,
        ),
    );

    return () => {
      unsubUsers();
      unsubPlanner();
      unsubLinks();
      unsubLeads();
      unsubBases();
      unsubGap();
      unsubIsencoes();
      unsubSolicitacoesManutencao();
      unsubPedidosCursos();
      unsubFiesProuni();
      unsubFiesProuniVagas();
      unsubCampanhas();
      unsubBomDia();
      unsubForecast();
      unsubMetaDia();
      unsubMetaSM();
      unsubMetaCursos();
      unsubQgLigacoes();
      unsubPeriodos();
      unsubCalendario();
      unsubEmpresas();
      unsubWhatsApp();
      unsubMapao();
      unsubBasesDisparo();
      unsubBasesRenovacao();
      unsubCursos();
      unsubControleConcorrencia();
      unsubInsumosPedidos();
      unsubInsumosEstoque();
      unsubInsumosPedidosComercial();
      unsubInsumosEstoqueComercial();
      unsubInsumosBaixas();
      unsubLigacoes();
      unsubMensagensEnviadas();
    };
  }, [user, profile]);

  useEffect(() => {
    const unsubBotConfig = onSnapshot(
      doc(db, COLLECTIONS.BOT_CONFIG, "main"),
      (snap) => {
        if (snap.exists()) {
          setBotConfig({ id: snap.id, ...snap.data() } as BotConfig);
        } else {
          setBotConfig({ url: "", active: false });
        }
      },
      (err) => {
        console.warn("Could not load botConfig publicly:", err);
      },
    );
    return () => unsubBotConfig();
  }, []);

  useEffect(() => {
    // Test connection to Firestore as per instructions
    const testConnection = async () => {
      try {
        const { getDocFromServer, doc } = await import("firebase/firestore");
        await getDocFromServer(
          doc(db, COLLECTIONS.BOT_CONFIG, "connection_test"),
        );
        console.log("Firestore connection test: OK");
      } catch (err) {
        console.warn(
          "Firestore connection test check (expected error if doc doesn't exist):",
          err,
        );
      }
    };
    testConnection();
  }, []);

  useEffect(() => {
    let intervalId: NodeJS.Timeout;

    const checkBotStatus = async () => {
      try {
        const data = await callBotApi("/api/status");
        if (data && data.bots) {
          setBotStatuses(data.bots);
        }
      } catch (e: any) {
        console.debug("Bot check fail via proxy:", e.message);
      }
    };

    checkBotStatus();
    intervalId = setInterval(checkBotStatus, 3000);
    return () => clearInterval(intervalId);
  }, [botConfig.url]);

  const knownLeadsRef = React.useRef<Set<string> | null>(null);
  const knownCampanhasRef = React.useRef<Set<string> | null>(null);

  useEffect(() => {
    if (!profile) return;
    if (
      profile.role !== ROLES.LIDER_FDV &&
      profile.role !== ROLES.SALA_MATRICULA
    )
      return;

    if (knownLeadsRef.current === null) {
      knownLeadsRef.current = new Set(leads.map((l) => l.id!));
      return;
    }

    let hasNew = false;
    leads.forEach((l) => {
      if (!knownLeadsRef.current!.has(l.id!)) {
        knownLeadsRef.current!.add(l.id!);
        hasNew = true;
      }
    });

    if (hasNew) {
      showPopup("Novo Lead!", "Um novo lead foi adicionado no HistÃ³rico.");
    }
  }, [leads, profile]);

  useEffect(() => {
    if (!profile) return;
    if (
      profile.role !== ROLES.LIDER_FDV &&
      profile.role !== ROLES.SALA_MATRICULA
    )
      return;

    if (knownCampanhasRef.current === null) {
      knownCampanhasRef.current = new Set(campanhas.map((c) => c.id!));
      return;
    }

    let hasNew = false;
    campanhas.forEach((c) => {
      if (!knownCampanhasRef.current!.has(c.id!)) {
        knownCampanhasRef.current!.add(c.id!);
        hasNew = true;
      }
    });

    if (hasNew) {
      showPopup("Nova Campanha!", "Uma nova campanha foi adicionada.");
    }
  }, [campanhas, profile]);

  useEffect(() => {
    if (profile && !canView(currentView)) {
      const availableViews = [
        "dashboard",
        "cadastro",
        "historico",
        "bases",
        "gap",
        "fiesProuni",
        "mapao",
        "cursos",
        "basesDisparo",
        "campanhas",
        "calendario",
        "empresas",
        "calculo",
        "emailMarketing",
        "admin",
        "controlePagamentos",
      ];
      const firstAvailable = availableViews.find((v) => canView(v));
      if (firstAvailable) {
        setCurrentView(firstAvailable);
      }
    }
  }, [profile, currentView]);

  const handleSaveLigacao = async (ligacao: Partial<Ligacao>) => {
    try {
      await addDoc(collection(db, COLLECTIONS.CONTROLE_LIGACOES), {
        ...ligacao,
        createdAt: serverTimestamp(),
      });
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, COLLECTIONS.CONTROLE_LIGACOES);
      throw err;
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-[#01112c] flex items-center justify-center">
        <motion.div
          animate={{ rotate: 360 }}
          transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
          className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full"
        />
      </div>
    );
  }

  const searchParams = new URLSearchParams(window.location.search);
  const publicFormId = searchParams.get("formId");

  if (publicFormId) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicCustomForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "pedido-insumos") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicInsumoForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "manutencao-publica") {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicMaintenanceForm />
      </div>
    );
  }

  if (currentView === "desconto") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicRegistrationForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "pedido-curso") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicPedidoCursoForm onToast={showToast} />
      </div>
    );
  }

  if (!user) {
    return <AuthScreen onToast={showToast} botConfig={botConfig} />;
  }

  if (profile?.blocked) {
    return (
      <div className="min-h-screen bg-[#01112c] flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-3xl shadow-xl border border-rose-100 text-center max-w-md">
          <XCircle size={64} className="text-rose-500 mx-auto mb-4" />
          <h2 className="text-2xl font-bold text-slate-900">
            Acesso Bloqueado
          </h2>
          <p className="text-slate-500 mt-2">
            Sua conta foi suspensa. Entre em contato com o administrador para
            mais informaÃ§Ãµes.
          </p>
          <button
            onClick={() => signOut(auth)}
            className="mt-6 w-full bg-slate-900 text-white py-3 rounded-xl font-bold hover:bg-slate-800 transition-all"
          >
            Sair
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#01112c] flex">
      <AnimatePresence>
        {toast && (
          <Toast
            message={toast.message}
            type={toast.type}
            onClose={() => setToast(null)}
          />
        )}
      </AnimatePresence>

      <AnimatePresence>
        {activePopup && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[200] flex items-center justify-center p-4">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white p-6 md:p-8 rounded-3xl shadow-xl border border-slate-100 max-w-sm w-full text-center relative"
            >
              <button
                onClick={() => setActivePopup(null)}
                className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors"
              >
                <X size={20} />
              </button>
              <div className="mx-auto w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-4 border-4 border-blue-50">
                <Bell size={32} className="text-blue-600" />
              </div>
              <h3 className="text-xl font-bold text-slate-900 mb-2">
                {activePopup.title}
              </h3>
              <p className="text-sm text-slate-600 mb-6">
                {activePopup.message}
              </p>
              <button
                onClick={() => setActivePopup(null)}
                className="w-full bg-blue-600 text-white font-bold py-3 rounded-xl hover:bg-blue-700 transition shadow-lg shadow-blue-200"
              >
                Ciente
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {massSendProgress.active && (
          <motion.div
            initial={{ y: 100, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 100, opacity: 0 }}
            className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-white px-6 py-4 rounded-2xl shadow-2xl border border-slate-200 z-[300] flex flex-col items-center gap-2 max-w-sm w-[90%]"
          >
            <div className="flex items-center gap-3 w-full">
              <div className="p-2 bg-blue-100 text-blue-600 rounded-full animate-pulse">
                <Bot size={20} />
              </div>
              <div className="flex-1">
                <h4 className="font-bold text-slate-800 text-sm">
                  Disparo em Massa (Bot)
                </h4>
                <p className="text-xs text-slate-500">
                  {massSendProgress.info}
                </p>
              </div>
              <div className="font-bold text-blue-600">
                {(
                  (massSendProgress.sent / (massSendProgress.total || 1)) *
                  100
                ).toFixed(0)}
                %
              </div>
            </div>
            <div className="w-full bg-slate-100 rounded-full h-2 mt-2 overflow-hidden">
              <div
                className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                style={{
                  width: `${(massSendProgress.sent / (massSendProgress.total || 1)) * 100}%`,
                }}
              />
            </div>

            <div className="flex gap-2 w-full mt-2">
              <button
                type="button"
                onClick={() => {
                  const newPaused = !isMassSendPaused;
                  massSendControlRef.current.paused = newPaused;
                  setIsMassSendPaused(newPaused);
                  showToast(newPaused ? "RobÃ´ pausado!" : "RobÃ´ retomado!");
                }}
                className={`flex-1 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-xl text-xs font-bold transition-all border ${
                  isMassSendPaused
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                    : "bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100"
                }`}
              >
                {isMassSendPaused ? (
                  <>
                    <Play size={14} /> Retomar
                  </>
                ) : (
                  <>
                    <Pause size={14} /> Pausar
                  </>
                )}
              </button>
              <button
                type="button"
                onClick={() => {
                  if (
                    window.confirm(
                      "Deseja realmente cancelar o envio em massa?",
                    )
                  ) {
                    massSendControlRef.current.cancelled = true;
                    massSendControlRef.current.paused = false;
                    setIsMassSendPaused(false);
                    showToast("Cancelando envio em massa...");
                  }
                }}
                className="bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100 py-1.5 px-3 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5"
              >
                <X size={14} /> Cancelar
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {profile?.mustChangePassword && (
        <PasswordChangeModal
          onComplete={async () => {
            try {
              if (user) {
                await updateDoc(doc(db, COLLECTIONS.USERS, user.uid), {
                  mustChangePassword: false,
                  updatedAt: serverTimestamp(),
                });
                setProfile((prev) =>
                  prev ? { ...prev, mustChangePassword: false } : null,
                );
                showToast("Senha atualizada com sucesso!");
              }
            } catch (err: any) {
              showToast("Erro ao atualizar status do perfil.", "error");
            }
          }}
        />
      )}

      <AnimatePresence>
        {isProfileModalOpen && profile && (
          <ProfileModal
            isOpen={isProfileModalOpen}
            onClose={() => setIsProfileModalOpen(false)}
            profile={profile}
            setProfile={setProfile}
            botConfig={botConfig}
            botStatuses={botStatuses}
            onToast={showToast}
          />
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showInjectModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[9999]"
          >
            <motion.div
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden"
            >
              <div className="bg-blue-600 p-6 text-white flex justify-between items-center">
                <div className="flex items-center gap-3">
                  <div className="bg-white/20 p-2 rounded-xl">
                    <MessageSquare size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold leading-none">Conectar via InjeÃ§Ã£o</h3>
                    <p className="text-blue-100 text-[10px] mt-1 font-medium">MODO AVANÃ‡ADO / ANTI-BLOQUEIO</p>
                  </div>
                </div>
                <button
                  onClick={() => setShowInjectModal(false)}
                  className="hover:bg-white/20 p-2 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
              
              <div className="p-6 space-y-4 max-h-[70vh] overflow-y-auto custom-scrollbar">
                <div className="bg-amber-50 border border-amber-200 p-4 rounded-2xl text-[11px] text-amber-800 leading-relaxed shadow-sm">
                  <p className="font-bold flex items-center gap-2 mb-2 text-amber-900">
                    <AlertCircle size={14} />
                    COMO USAR (SOP):
                  </p>
                  <ol className="list-decimal ml-4 space-y-2">
                    <li>Abra o <strong>WhatsApp Web oficial</strong> em uma aba anÃ´nima do Chrome.</li>
                    <li>FaÃ§a o login normal pelo celular (QR Code ou NÃºmero).</li>
                    <li>Com o WhatsApp aberto, clique na extensÃ£o <strong>PESK Linker</strong> e copie o JSON.</li>
                    <li>Cole o cÃ³digo abaixo e clique em Injetar.</li>
                    <li><strong>IMPORTANTE:</strong> Feche a aba do WhatsApp Web imediatamente apÃ³s o sucesso.</li>
                  </ol>
                </div>
                
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5 ml-1">NÃºmero do Bot</label>
                  <div className="relative">
                    <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                      <Hash size={16} />
                    </div>
                    <input 
                      type="text" 
                      value={injectBotNumber} 
                      onChange={e => setInjectBotNumber(e.target.value.replace(/\D/g, ''))}
                      className="w-full bg-slate-50 border border-slate-200 pl-10 pr-4 py-3 rounded-2xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-sm font-bold"
                      placeholder="5524999999999"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5 ml-1">Dados da SessÃ£o (JSON)</label>
                  <textarea 
                    rows={6}
                    value={injectSessionData} 
                    onChange={e => setInjectSessionData(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 px-4 py-3 rounded-2xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all font-mono text-[10px] resize-none"
                    placeholder='Cole aqui o JSON gerado pela extensÃ£o...'
                  />
                </div>

                <button
                  onClick={async () => {
                    if (!injectBotNumber || !injectSessionData) {
                      showToast("Por favor, preencha o nÃºmero e cole o JSON.", "error");
                      return;
                    }
                    try {
                      let sessionDataObj;
                      try {
                        sessionDataObj = JSON.parse(injectSessionData);
                      } catch (e) {
                        showToast("JSON invÃ¡lido! Copie novamente da extensÃ£o.", "error");
                        return;
                      }

                      await callBotApi("/api/inject", {
                        method: "POST",
                        body: { 
                          botNumber: injectBotNumber,
                          sessionData: sessionDataObj
                        }
                      });
                      
                      showToast("Sucesso! SessÃ£o injetada. O bot estÃ¡ iniciando...", "success");
                      setShowInjectModal(false);
                      setInjectSessionData("");
                      
                      // Refresh status after injection
                      setTimeout(async () => {
                        try {
                          const data = await callBotApi("/api/status");
                          if (data && data.bots) setBotStatuses(data.bots);
                        } catch (e) {}
                      }, 4000);
                    } catch (err: any) {
                      showToast(`Erro na injeÃ§Ã£o: ${err.message}`, "error");
                    }
                  }}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-600/20 transition-all flex items-center justify-center gap-2 group"
                >
                  <span>Injetar SessÃ£o e Conectar</span>
                  <ChevronRight size={18} className="group-hover:translate-x-1 transition-transform" />
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 w-64 bg-[#011a3c] border-r border-[#092e5c] transform transition-transform duration-300 ease-in-out lg:translate-x-0 lg:static lg:inset-0",
          isSidebarOpen ? "translate-x-0" : "-translate-x-full",
        )}
      >
        <div className="h-full flex flex-col">
          <div className="p-6 flex items-center space-x-3">
            {botConfig?.loginLogo ? (
              <img
                src={botConfig.loginLogo}
                alt="Logo"
                className="w-full max-h-12 object-contain drop-shadow-md"
                referrerPolicy="no-referrer"
              />
            ) : (
              <>
                <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
                  <TrendingUp size={24} />
                </div>
                <h1 className="text-xl font-bold text-white tracking-tight">
                  GestÃ£o Oeste pro
                </h1>
              </>
            )}
          </div>

          <nav className="flex-1 px-4 space-y-1 overflow-y-auto pr-2">
            {[
              { id: "dashboard", label: "Rotina", icon: LayoutDashboard },
              { id: "relatorios", label: "RelatÃ³rios", icon: BarChart3 },
              { id: "cadastro", label: "Novo Lead", icon: UserPlus },
              { id: "historico", label: "HistÃ³rico", icon: History },
              { id: "crm", label: "CRM WhatsApp", icon: MessageSquare },
              { id: "bases", label: "Bases", icon: Database },
              { id: "gap", label: "GAP AcadÃªmico", icon: GraduationCap },
              {
                id: "isencoes",
                label: "Acompanhamento de IsenÃ§Ãµes",
                icon: ShieldCheck,
              },
              { id: "fiesProuni", label: "Fies/Prouni", icon: FileText },
              { id: "mapao", label: "MapÃ£o AcadÃªmico", icon: MapPin },
              { id: "cursos", label: "Cursos DisponÃ­veis", icon: BookOpen },
              { id: "basesDisparo", label: "Bases de Disparo", icon: Globe },
              { id: "basesRenovacao", label: "Base LÃ­quida", icon: Database },
              { id: "campanhas", label: "Campanhas", icon: Megaphone },
              { id: "calendario", label: "Plano de AÃ§Ã£o", icon: Calendar },
              { id: "empresas", label: "Empresas Parceiras", icon: Building2 },
              {
                id: "controleConcorrencia",
                label: "Controle de ConcorrÃªncia",
                icon: Target,
              },
              { id: "controleLigacoes", label: "Controle de LigaÃ§Ãµes", icon: Phone },
              { id: "evasao", label: "EvasÃ£o", icon: UserMinus },
              {
                id: "calculo",
                label: "CÃ¡lculo de RemuneraÃ§Ã£o",
                icon: Calculator,
              },
              {
                id: "controlePagamentos",
                label: "Controle de Pagamentos",
                icon: Coins,
              },
              {
                id: "solicitacaoManutencao",
                label: "SolicitaÃ§Ã£o de ManutenÃ§Ã£o",
                icon: Wrench,
              },
              {
                id: "controleInsumos",
                label: "Controle de Insumos",
                icon: Boxes,
              },
              {
                id: "controleInsumosComercial",
                label: "Controle de Insumos (Comercial)",
                icon: Boxes,
              },
              {
                id: "emailMarketing",
                label: "Envio de e-mail Marketing",
                icon: Mail,
              },
              { id: "admin", label: "AdministraÃ§Ã£o", icon: Settings },
            ].map(
              (item) =>
                canView(item.id) && (
                  <button
                    key={item.id}
                    onClick={() => {
                      setCurrentView(item.id);
                      setIsSidebarOpen(false);
                    }}
                    className={cn(
                      "w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold transition-all",
                      currentView === item.id
                        ? "bg-blue-500/10 text-white"
                        : "text-slate-400 hover:bg-[#082a5c] hover:text-white",
                    )}
                  >
                    <item.icon size={20} />
                    <span>{item.label}</span>
                  </button>
                ),
            )}
          </nav>

          <div className="p-4 border-t border-[#092e5c]">
            <div className="bg-[#082a5c]/50 p-4 rounded-2xl mb-4">
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">
                UsuÃ¡rio
              </p>
              <p className="text-sm font-bold text-white truncate">
                {profile?.name}
              </p>
              <span className="inline-block mt-1 px-2 py-0.5 bg-blue-500/20 text-blue-400 text-[10px] font-bold rounded-full">
                {profile?.role}
              </span>
            </div>

            <div className="space-y-1">
              <button
                onClick={async () => {
                  if (user?.email) {
                    try {
                      await sendPasswordResetEmail(auth, user.email);
                      showToast("E-mail de redefiniÃ§Ã£o enviado!");
                    } catch (err: any) {
                      showToast("Erro ao enviar e-mail.", "error");
                    }
                  }
                }}
                className="w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold text-slate-400 hover:bg-[#082a5c] hover:text-white transition-all"
              >
                <KeyRound size={20} />
                <span>Trocar Senha</span>
              </button>

              <button
                onClick={() => signOut(auth)}
                className="w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold text-rose-400 hover:bg-rose-500/10 transition-all"
              >
                <LogOut size={20} />
                <span>Sair do Sistema</span>
              </button>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header className="h-16 bg-[#011a3c] border-b border-[#092e5c] flex items-center justify-between px-4 lg:px-8 shrink-0">
          <button
            onClick={() => setIsSidebarOpen(true)}
            className="lg:hidden p-2 text-slate-400 hover:bg-[#082a5c] rounded-lg"
          >
            <Menu size={24} />
          </button>
          <div className="flex-1 lg:flex-none flex items-center space-x-3 flex-wrap gap-y-1">
            <h2 className="text-lg font-bold text-white capitalize ml-2 lg:ml-0">
              {currentView.replace("-", " ")}
            </h2>
            <span className="px-2.5 py-1 bg-gradient-to-r from-blue-600 to-sky-500 text-white text-[10px] font-extrabold rounded-md shadow-sm uppercase tracking-wider">
              Servidor:{" "}
              {localStorage.getItem("servidor_selected") === "comercial"
                ? "Comercial"
                : "SM"}
            </span>
            {isOnline ? (
              <span className="flex items-center space-x-1.5 px-2.5 py-1 bg-emerald-500/10 text-emerald-400 text-[10px] font-extrabold rounded-md border border-emerald-500/20 shadow-sm uppercase tracking-wider">
                <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse"></span>
                <span>Online / Sincronizado</span>
              </span>
            ) : (
              <span className="flex items-center space-x-1.5 px-2.5 py-1 bg-amber-500/10 text-amber-400 text-[10px] font-extrabold rounded-md border border-amber-500/20 shadow-sm uppercase tracking-wider animate-pulse">
                <span className="w-1.5 h-1.5 bg-amber-400 rounded-full"></span>
                <span>Sem ConexÃ£o (Modo Cache Offline)</span>
              </span>
            )}
          </div>
          <div className="flex items-center space-x-4">
            <div className="hidden md:flex items-center space-x-2 text-sm text-slate-400">
              <Calendar size={16} />
              <span>{new Date().toLocaleDateString("pt-BR")}</span>
            </div>
            <button
              onClick={() => setIsProfileModalOpen(true)}
              className="flex items-center space-x-1.5 bg-[#082a5c]/50 hover:bg-[#082a5c] text-white px-3 py-1.5 rounded-xl border border-[#092e5c] text-sm font-semibold transition-all active:scale-95 cursor-pointer"
            >
              <UserIcon size={15} className="text-slate-300" />
              <span>Perfil</span>
            </button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-4 lg:p-8">
          <AnimatePresence mode="wait">
            <motion.div
              key={currentView}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
            >
              {currentView === "dashboard" && (
                <DashboardView
                  leads={leads}
                  planner={[...planner]}
                  links={links}
                  profile={profile!}
                  onToast={showToast}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}
                  users={users}
                />
              )}
              {currentView === "relatorios" && (
                <RelatoriosView
                  leads={leads}
                  bases={bases}
                  fiesProuni={fiesProuni}
                  calendarioAcoes={calendarioAcoes}
                  pedidosCursos={pedidosCursos}
                  empresasParceiras={empresasParceiras}
                  insumosPedidos={insumosPedidos}
                  insumosEstoque={insumosEstoque}
                  insumosBaixas={insumosBaixas}
                  isencoes={isencoes}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  ligacoes={ligacoes}
                  solicitacoesManutencao={solicitacoesManutencao}
                  salesContacts={salesContacts}
                  mensagensEnviadasLog={mensagensEnviadasLog}
                  analysisSchemes={analysisSchemes}
                  profile={profile!}
                  onToast={showToast}
                />
              )}
              {currentView === "cadastro" && (
                <CadastroView
                  onToast={showToast}
                  profile={profile!}
                  calendarioAcoes={calendarioAcoes}
                  uniqueUnidades={uniqueUnidades}
                />
              )}
              {currentView === "historico" && (
                <HistoricoView
                  leads={leads}
                  profile={profile!}
                  onToast={showToast}
                  users={users}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  gap={gap}
                  basesRenovacao={basesRenovacao}
                  calendarioAcoes={calendarioAcoes}
                  pedidosCursos={pedidosCursos}
                  mensagensEnviadasLog={mensagensEnviadasLog}
                />
              )}
              {currentView === "crm" && (
                ["Admin Master", "LÃ­der/FDV"].includes(profile?.role || "") ? (
                  <CRMView
                    leads={leads}
                    bases={bases}
                    fiesProuni={fiesProuni}
                    gap={gap}
                    profile={profile!}
                    onSendBot={handleSendBotMessage}
                    onToast={showToast}
                  />
                ) : (
                  <div className="p-12 text-center bg-white rounded-3xl border border-slate-100 shadow-sm max-w-2xl mx-auto mt-10">
                    <div className="w-20 h-20 bg-slate-50 text-slate-300 rounded-full flex items-center justify-center mx-auto mb-6">
                      <MessageSquare size={40} />
                    </div>
                    <h3 className="text-2xl font-bold text-slate-800">Acesso Restrito</h3>
                    <p className="text-slate-500 mt-3 text-lg">A aba CRM estÃ¡ disponÃ­vel apenas para Perfis LÃ­der/FDV e Administradores enquanto estiver em fase de testes.</p>
                    <button 
                      onClick={() => setCurrentView('dashboard')}
                      className="mt-8 bg-blue-600 text-white px-8 py-3 rounded-xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-500/20"
                    >
                      Voltar ao InÃ­cio
                    </button>
                  </div>
                )
              )}
              {currentView === "bases" && (
                <BasesView
                  bases={bases}
                  profile={profile!}
                  onToast={showToast}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  gap={gap}
                  basesRenovacao={basesRenovacao}
                  mensagensEnviadasLog={mensagensEnviadasLog}
                />
              )}
              {currentView === "gap" && (
                <GapView
                  gap={gap}
                  onToast={showToast}
                  profile={profile}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  calendarioAcoes={calendarioAcoes}
                />
              )}
              {currentView === "isencoes" && (
                <IsencoesView
                  isencoes={isencoes}
                  gap={gap}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "fiesProuni" && (
                <FiesProuniView
                  data={fiesProuni}
                  vagas={fiesProuniVagas}
                  onToast={showToast}
                  profile={profile!}
                  whatsappMessages={whatsappMessages}
                  periodos={periodos}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                />
              )}
              {currentView === "mapao" && (
                <MapaoAcademicoView
                  mapao={mapao}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "cursos" && (
                <CursosDisponiveisView
                  cursos={cursos}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "basesDisparo" && (
                <BasesDisparoView bases={basesDisparo} onToast={showToast} />
              )}
              {currentView === "basesRenovacao" && (
                <BasesRenovacaoView
                  bases={basesRenovacao}
                  onToast={showToast}
                  profile={profile}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                />
              )}
              {currentView === "campanhas" && (
                <CampanhasView campanhas={campanhas} onToast={showToast} />
              )}
              {currentView === "calculo" && <CalculoRemuneracaoView />}
              {currentView === "emailMarketing" && (
                <EmailMarketingView onToast={showToast} />
              )}
              {currentView === "controlePagamentos" && (
                <ControlePagamentosView
                  calendarioAcoes={calendarioAcoes}
                  users={users}
                  onToast={showToast}
                  profile={profile}
                />
              )}
              {currentView === "solicitacaoManutencao" && (
                <SolicitacoesManutencaoView
                  profile={profile}
                  onToast={showToast}
                  users={users}
                />
              )}
              {currentView === "controleInsumos" && (
                <ControleInsumosView
                  pedidos={insumosPedidos}
                  estoque={insumosEstoque}
                  profile={profile!}
                  onToast={showToast}
                  botConfig={botConfig}
                />
              )}
              {currentView === "controleInsumosComercial" && (
                <ControleInsumosComercialView
                  pedidos={insumosPedidosComercial}
                  estoque={insumosEstoqueComercial}
                  profile={profile!}
                  onToast={showToast}
                  botConfig={botConfig}
                />
              )}
              {currentView === "calendario" && (
                <CalendarioAcoesView
                  data={calendarioAcoes}
                  onToast={showToast}
                  profile={profile!}
                  initialData={initialActionData}
                  onClearInitialData={() => setInitialActionData(null)}
                  users={users}
                  callBotApi={callBotApi}
                  leads={leads}
                  gap={gap}
                  onSendWhatsApp={sendAppWhatsApp}
                />
              )}
              {currentView === "empresas" && (
                <EmpresasParceirasView
                  data={empresasParceiras}
                  leads={leads}
                  acoes={calendarioAcoes}
                  onToast={showToast}
                  cursos={cursos}
                  users={users}
                  onSendWhatsApp={sendAppWhatsApp}
                  botConfig={botConfig}
                  uniqueUnidades={uniqueUnidades}
                  profile={profile!}
                  onGenerateAction={(empresa) => {
                    setInitialActionData({
                      nome: `AÃ§Ã£o na empresa ${empresa.nome}`,
                      local: empresa.endereco,
                      observacao: `ResponsÃ¡vel: ${empresa.responsavel}\nTelefone: ${empresa.telefone}`,
                    });
                    setCurrentView("calendario");
                  }}
                />
              )}
              {currentView === "controleConcorrencia" && (
                <ControleConcorrenciaView
                  data={controleConcorrencia}
                  onToast={showToast}
                />
              )}
              {currentView === "controleLigacoes" && (
              <ControleLigacoesView
                leads={leads}
                bases={bases}
                acoes={calendarioAcoes}
                ligacoes={ligacoes}
                fiesProuni={fiesProuni}
                gap={gap}
                profile={profile!}
                onSaveLigacao={handleSaveLigacao}
                onToast={showToast}
              />
            )}

            {currentView === "evasao" && (
                <EvasaoView profile={profile} onToast={showToast} />
              )}
              {currentView === "admin" && (
                <AdminView
                  profile={profile}
                  users={users}
                  links={links}
                  onToast={showToast}
                  leads={leads}
                  bases={bases}
                  gap={gap}
                  planner={[...planner]}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  whatsappMessages={whatsappMessages}
                  activeWhatsappTemplates={activeWhatsappTemplates}
                  setActiveWhatsappTemplates={setActiveWhatsappTemplates}
                  empresasParceiras={empresasParceiras}
                  botConfig={botConfig}
                  botStatuses={botStatuses}
                  setBotStatuses={setBotStatuses}
                  callBotApi={callBotApi}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}
                  cursos={cursos}
                  uniqueUnidades={uniqueUnidades}
                  analysisSchemes={analysisSchemes}
                  onSaveAnalysisScheme={handleSaveAnalysisScheme}
                  onDeleteAnalysisScheme={handleDeleteAnalysisScheme}
                  setShowInjectModal={setShowInjectModal}
                />
              )}
            </motion.div>
          </AnimatePresence>

          <footer className="mt-12 py-6 border-t border-slate-200 text-center">
            <p className="text-sm text-slate-500 font-medium">
              Sistema Criado por{" "}
              <span className="font-bold text-slate-900">Agencia Argo's</span> -
              <a
                href={getWhatsAppUrl(
                  "24992777019",
                  "Gostaria de realizar um orÃ§amento para um sistema",
                )}
                target="_blank"
                rel="noopener noreferrer"
                className="ml-1 text-blue-600 hover:underline font-bold"
              >
                Telefone: (24) 99277-7019
              </a>
            </p>
          </footer>
        </div>
      </main>

      {/* Mobile Sidebar Overlay */}
      {isSidebarOpen && (
        <div
          className="fixed inset-0 bg-slate-900/50 z-30 lg:hidden backdrop-blur-sm"
          onClick={() => setIsSidebarOpen(false)}
        />
      )}
    </div>
  );
}

// --- View Components ---

function AvisosView() {
  return null;
}

function AuthScreen({
  onToast,
  botConfig,
}: {
  onToast: (m: string, t?: "success" | "error") => void;
  botConfig?: BotConfig;
}) {
  const [isLogin, setIsLogin] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [servidor, setServidor] = useState<"principal" | "comercial">(
    (localStorage.getItem("servidor_selected") as "principal" | "comercial") ||
      "principal",
  );
  const [loading, setLoading] = useState(false);

  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallGuide, setShowInstallGuide] = useState(false);
  const [showApkModal, setShowApkModal] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const isStandalone =
      window.matchMedia("(display-mode: standalone)").matches ||
      (window.navigator as any).standalone ||
      document.referrer.includes("android-app://");
    setIsAppInstalled(isStandalone);

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener(
        "beforeinstallprompt",
        handleBeforeInstallPrompt,
      );
    };
  }, []);

  const handleInstallClick = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: { outcome: string }) => {
        if (choiceResult.outcome === "accepted") {
          console.log("User accepted the PWA install prompt");
          setIsAppInstalled(true);
        }
        setDeferredPrompt(null);
      });
    } else {
      setShowInstallGuide((prev) => !prev);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    if (!isLogin && password.length < 6) {
      onToast("A senha deve ter pelo menos 6 caracteres.", "error");
      setLoading(false);
      return;
    }
    try {
      if (isLogin) {
        try {
          // Attempt login on the CURRENTLY SELECTED server
          await signInWithEmailAndPassword(auth, email, password);
        } catch (err: any) {
          // If user login fails with invalid credentials or user not found,
          // let's check programmatically if the credentials are valid on the OTHER server!
          const isUserNotFound =
            err.code === "auth/user-not-found" ||
            err.code === "auth/invalid-credential";
          if (isUserNotFound) {
            const currentSelected = servidor;
            const alternativeServer =
              currentSelected === "principal" ? "comercial" : "principal";

            // Build the alternative config
            const altConfig =
              alternativeServer === "comercial"
                ? firebaseConfigComercial
                : firebaseConfigPrincipal;

            // Resolve alternative app
            let altApp;
            try {
              altApp = getApp("alternative_login_check");
            } catch {
              altApp = initializeApp(altConfig, "alternative_login_check");
            }
            const altAuth = getAuth(altApp);

            try {
              // Attempt login on the ALTERNATIVE server
              await signInWithEmailAndPassword(altAuth, email, password);
              // SUCCESS on the other server! Let's update localStorage and reload to apply the active configuration
              localStorage.setItem("servidor_selected", alternativeServer);
              onToast(
                `Login bem sucedido! Redirecionando para o Servidor ${alternativeServer === "principal" ? "Principal" : "Comercial"}...`,
                "success",
              );
              setTimeout(() => {
                window.location.reload();
              }, 1200);
              return;
            } catch (altErr) {
              // Failed on both servers, throw the original authentication error
              throw err;
            }
          } else {
            throw err;
          }
        }
      } else {
        const userCred = await createUserWithEmailAndPassword(
          auth,
          email,
          password,
        );
        // Pack the chosen servidor into displayName so App.tsx can extract it
        await updateProfile(userCred.user, {
          displayName: `${name}|${servidor}`,
        });
        onToast("Conta criada com sucesso!");
      }
    } catch (err: any) {
      console.error("Auth error details (AuthScreen):", {
        code: err.code,
        message: err.message,
        stack: err.stack,
      });
      let friendlyMessage = err.message;
      if (
        err.code === "auth/user-not-found" ||
        err.code === "auth/invalid-credential"
      ) {
        friendlyMessage =
          "E-mail ou senha invÃ¡lidos em ambos os servidores (Principal / Comercial).";
      } else if (err.code === "auth/wrong-password") {
        friendlyMessage = "Senha incorreta.";
      }
      onToast(`Erro: ${friendlyMessage}`, "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#011430] flex flex-col md:flex-row relative overflow-hidden font-sans text-white">
      {/* Absolute Ambient Background Lights */}
      <div className="absolute top-0 right-0 w-[500px] h-[500px] rounded-full bg-blue-500/15 blur-[120px] pointer-events-none z-0" />
      <div className="absolute bottom-0 left-0 w-[500px] h-[500px] rounded-full bg-cyan-500/10 blur-[130px] pointer-events-none z-0" />

      {/* LEFT COLUMN: Login panel container */}
      <div className="w-full md:w-[42%] lg:w-[38%] xl:w-[34%] bg-[#011a3c] border-r border-[#092e5c] p-8 sm:p-12 md:p-16 flex flex-col justify-between relative z-10 shadow-2xl min-h-screen">
        <div className="my-auto space-y-8">
          <div>
            {botConfig?.loginLogo ? (
              <div className="mb-6 flex">
                <img
                  src={botConfig.loginLogo}
                  alt="Logo"
                  className="max-h-32 w-full object-contain drop-shadow-lg"
                  referrerPolicy="no-referrer"
                />
              </div>
            ) : (
              <>
                <div className="w-16 h-16 bg-gradient-to-tr from-sky-500 to-blue-600 rounded-2xl flex items-center justify-center text-white shadow-lg shadow-sky-500/20 mb-6">
                  <TrendingUp size={32} />
                </div>
                <h2 className="text-3xl font-extrabold text-white tracking-tight">
                  GestÃ£o Oeste pro
                </h2>
              </>
            )}
            <p className="text-slate-400 mt-2 text-sm">
              {isLogin
                ? "Bem-vindo de volta! Insira suas credenciais:"
                : "Preencha os dados e crie sua conta agora:"}
            </p>
          </div>

          {/* Servidor Selector (Principal vs Comercial) */}
          <div className="flex bg-[#032554] p-1.5 rounded-2xl border border-[#0b3c7c] shadow-inner">
            <button
              type="button"
              onClick={() => {
                if (servidor !== "principal") {
                  localStorage.setItem("servidor_selected", "principal");
                  window.location.reload();
                }
              }}
              className={`flex-1 py-3 text-xs font-extrabold rounded-xl transition-all cursor-pointer ${servidor === "principal" ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20" : "text-slate-400 hover:text-white"}`}
            >
              Principal
            </button>
            <button
              type="button"
              onClick={() => {
                if (servidor !== "comercial") {
                  localStorage.setItem("servidor_selected", "comercial");
                  window.location.reload();
                }
              }}
              className={`flex-1 py-3 text-xs font-extrabold rounded-xl transition-all cursor-pointer ${servidor === "comercial" ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20" : "text-slate-400 hover:text-white"}`}
            >
              Comercial
            </button>
          </div>

          {/* Alerta de ConfiguraÃ§Ã£o se os projetos forem idÃªnticos no modo Principal */}
          {servidor === "principal" &&
            firebaseConfigPrincipal.apiKey === "AIzaSyBexxjzDAuNSgY90rlVqpz4AQZDE-QwSG4" && (
              <div className="p-3 bg-amber-500/10 border border-amber-500/20 rounded-xl text-amber-200 text-[10px] font-bold leading-relaxed flex items-start gap-2 animate-pulse">
                <AlertCircle size={14} className="flex-shrink-0 mt-0.5" />
                <p>
                  O servidor Principal estÃ¡ usando as credenciais do servidor
                  Comercial. Configure as variÃ¡veis de ambiente (API Key e App
                  ID) do projeto gestaopro-761e1 para habilitar o acesso.
                </p>
              </div>
            )}

          <form onSubmit={handleSubmit} className="space-y-4 font-sans">
            {!isLogin && (
              <div>
                <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                  Nome Completo
                </label>
                <input
                  type="text"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  required
                  className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                  placeholder="Seu nome"
                />
              </div>
            )}
            <div>
              <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                Email
              </label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                placeholder="seu@email.com"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                Senha
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
                className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                placeholder="â€¢â€¢â€¢â€¢â€¢â€¢â€¢â€¢"
              />
            </div>

            {isLogin && (
              <div className="flex justify-end">
                <button
                  type="button"
                  onClick={async () => {
                    let resetEmail = email;
                    if (!resetEmail) {
                      const inputEmail = window.prompt(
                        "Por favor, digite seu e-mail para receber o link de redefiniÃ§Ã£o de senha:",
                      );
                      if (!inputEmail) return;
                      resetEmail = inputEmail;
                    }
                    try {
                      await sendPasswordResetEmail(auth, resetEmail);
                      onToast(
                        "E-mail de redefiniÃ§Ã£o enviado com sucesso! Verifique sua caixa de entrada.",
                        "success",
                      );
                    } catch (err: any) {
                      onToast(
                        "Erro ao enviar e-mail. Verifique se o endereÃ§o Ã© vÃ¡lido.",
                        "error",
                      );
                    }
                  }}
                  className="text-xs font-bold text-sky-400 hover:text-sky-300 hover:underline transition-colors cursor-pointer"
                >
                  Esqueci minha senha
                </button>
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-gradient-to-tr from-sky-500 to-blue-600 hover:from-sky-600 hover:to-blue-700 text-white font-bold py-4 rounded-xl shadow-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed active:scale-[0.98] cursor-pointer"
            >
              {loading
                ? "Processando..."
                : isLogin
                  ? "Entrar no Sistema"
                  : "Criar Minha Conta"}
            </button>
          </form>

          <div className="mt-8 text-center pt-2 border-t border-[#092e5c]">
            <button
              onClick={() => setIsLogin(!isLogin)}
              className="text-sm font-semibold text-sky-400 hover:text-sky-300 hover:underline cursor-pointer"
            >
              {isLogin
                ? "NÃ£o tem uma conta? Cadastre-se"
                : "JÃ¡ tem uma conta? FaÃ§a login"}
            </button>
          </div>

          {/* Android App Promotion Card on Login */}
          {!isAppInstalled && (
            <div className="mt-8 pt-6 border-t border-[#092e5c] space-y-4">
              <div className="bg-[#032554]/60 p-5 rounded-2xl border border-sky-500/10 text-white relative overflow-hidden transition-all duration-300">
                <div className="flex items-start space-x-3.5">
                  <div className="p-2.5 bg-sky-950/80 rounded-xl border border-sky-500/20 text-emerald-400 flex items-center justify-center shadow shrink-0">
                    <Smartphone size={24} />
                  </div>
                  <div className="space-y-1">
                    <h4 className="text-sm font-extrabold text-white leading-tight">
                      Instalar Aplicativo (Android)
                    </h4>
                    <p className="text-xs text-slate-300 leading-relaxed font-semibold">
                      Deseja usar no celular? Instale o App para usar{" "}
                      <strong className="text-emerald-400 font-extrabold">
                        com ou sem internet
                      </strong>
                      . Sincroniza automÃ¡tico ao conectar.
                    </p>
                  </div>
                </div>

                <div className="flex items-center space-x-2 mt-4">
                  <button
                    onClick={handleInstallClick}
                    className="flex-1 flex items-center justify-center space-x-1.5 px-3 py-2.5 bg-emerald-500 hover:bg-emerald-600 text-slate-950 font-extrabold text-xs rounded-lg shadow-md hover:scale-[1.01] transition-all cursor-pointer"
                  >
                    <Download size={14} />
                    <span>Instalar no Aparelho</span>
                  </button>
                  <button
                    onClick={() => setShowInstallGuide(!showInstallGuide)}
                    className="px-3 py-2.5 bg-white/10 hover:bg-white/15 text-slate-100 font-bold text-xs rounded-lg transition-all cursor-pointer flex-1"
                  >
                    InstruÃ§Ãµes
                  </button>
                </div>

                {showInstallGuide && (
                  <div className="mt-4 pt-4 border-t border-[#092e5c] space-y-3 text-xs">
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 1:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Abra este endereÃ§o no{" "}
                        <strong className="text-white font-bold">
                          Google Chrome
                        </strong>{" "}
                        do seu Android.
                      </p>
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 2:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Toque nos{" "}
                        <strong className="text-white font-bold">
                          trÃªs pontinhos (â‹®)
                        </strong>{" "}
                        no canto superior direito.
                      </p>
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 3:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Selecione{" "}
                        <strong className="text-emerald-400 font-extrabold">
                          "Instalar aplicativo"
                        </strong>{" "}
                        ou{" "}
                        <strong className="text-emerald-400 font-extrabold">
                          "Adicionar Ã  tela inicial"
                        </strong>
                        .
                      </p>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Humble system credits info */}
        <div className="text-center text-[10px] text-slate-500 font-mono tracking-widest mt-6">
          OESTE HUNTER Â© {new Date().getFullYear()}
        </div>
      </div>

      {/* RIGHT COLUMN: The majestic interactive Oeste Hunter logo artwork or Custom Logo */}
      <div className="hidden md:flex flex-1 items-center justify-center bg-[#01112c] p-12 relative overflow-hidden z-0">
        {/* Subtle grid mesh backdrop */}
        <div className="absolute inset-0 bg-[radial-gradient(#082a5c_1px,transparent_1px)] [background-size:24px_24px] opacity-25" />

        {/* Animated ambient outer halo circles */}
        <div className="absolute w-[600px] h-[600px] border border-[#0d4182]/20 rounded-full animate-pulse" />
        <div className="absolute w-[800px] h-[800px] border border-[#0d4182]/10 rounded-full opacity-60" />

        {/* SVG ART Container */}
        <div className="relative z-10 w-full flex justify-center">
          {botConfig?.loginLogo ? (
            <img
              src={botConfig.loginLogo}
              alt="Logo Promocional"
              className="w-full max-w-[560px] aspect-square rounded-3xl object-contain drop-shadow-[0_35px_60px_rgba(14,116,253,0.35)] border border-slate-700/40 p-12 bg-[#011a3c]/50 animate-fade-in"
              referrerPolicy="no-referrer"
            />
          ) : (
            /* Oeste Hunter Badge SVG */
            <svg
              viewBox="0 0 1000 1000"
              className="w-full max-w-[560px] aspect-square drop-shadow-[0_25px_60px_rgba(14,116,253,0.35)] select-none"
            >
              <defs>
                <linearGradient
                  id="blueRingGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="100%"
                >
                  <stop offset="0%" stopColor="#0a397a" />
                  <stop offset="50%" stopColor="#125cb5" />
                  <stop offset="100%" stopColor="#082c5f" />
                </linearGradient>
                <linearGradient
                  id="wolfEyeGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="0%"
                >
                  <stop offset="0%" stopColor="#00f0ff" />
                  <stop offset="100%" stopColor="#00a8ff" />
                </linearGradient>
                <linearGradient
                  id="muzzleGrad"
                  x1="0%"
                  y1="0%"
                  x2="0%"
                  y2="100%"
                >
                  <stop offset="0%" stopColor="#ffffff" />
                  <stop offset="100%" stopColor="#cfd8dc" />
                </linearGradient>
                <linearGradient
                  id="bannerGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="0%"
                >
                  <stop offset="0%" stopColor="#010f24" />
                  <stop offset="50%" stopColor="#051c3d" />
                  <stop offset="100%" stopColor="#010d21" />
                </linearGradient>
                <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
                  <feGaussianBlur stdDeviation="8" result="blur" />
                  <feComposite in="SourceGraphic" in2="blur" operator="over" />
                </filter>
                <filter
                  id="eyeGlow"
                  x="-50%"
                  y="-50%"
                  width="200%"
                  height="200%"
                >
                  <feGaussianBlur stdDeviation="6" result="blur" />
                  <feComposite in="SourceGraphic" in2="blur" operator="over" />
                </filter>
              </defs>

              {/* Target Crosshairs Reticle */}
              <g stroke="#ffffff" strokeWidth="2.5" opacity="0.3">
                {/* Vertical Crosshair Line */}
                <line x1="500" y1="20" x2="500" y2="980" />
                {/* Horizontal Crosshair Line */}
                <line x1="20" y1="500" x2="980" y2="500" />

                {/* Target ticks (top, bottom, left, right) */}
                <line x1="500" y1="80" x2="520" y2="80" />
                <line x1="500" y1="140" x2="515" y2="140" />
                <line x1="500" y1="200" x2="520" y2="200" />

                <line x1="500" y1="920" x2="520" y2="920" />
                <line x1="500" y1="860" x2="515" y2="860" />
                <line x1="500" y1="800" x2="520" y2="800" />

                <line x1="80" y1="500" x2="80" y2="520" />
                <line x1="140" y1="500" x2="140" y2="515" />
                <line x1="200" y1="500" x2="200" y2="520" />

                <line x1="920" y1="500" x2="920" y2="520" />
                <line x1="860" y1="500" x2="860" y2="515" />
                <line x1="800" y1="500" x2="800" y2="520" />
              </g>

              {/* Target Reticle Outer Box ticks */}
              <rect
                x="480"
                y="40"
                width="40"
                height="20"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="480"
                y="940"
                width="40"
                height="20"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="40"
                y="480"
                width="20"
                height="40"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="940"
                y="480"
                width="20"
                height="40"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />

              {/* 1. Outer target circle with dashes */}
              <circle
                cx="500"
                cy="500"
                r="445"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                strokeDasharray="16 20"
                opacity="0.35"
              />

              {/* 2. Concentric circle borders */}
              <circle
                cx="500"
                cy="500"
                r="415"
                fill="none"
                stroke="#ffffff"
                strokeWidth="2.5"
                opacity="0.4"
              />

              {/* 3. Main Thick Ring Outer Rim */}
              <circle
                cx="500"
                cy="500"
                r="400"
                fill="none"
                stroke="#ffffff"
                strokeWidth="4"
              />

              {/* 4. The Mighty Blue Ring Body */}
              <circle
                cx="500"
                cy="500"
                r="348"
                fill="none"
                stroke="url(#blueRingGrad)"
                strokeWidth="100"
              />

              {/* 5. Inner Rim of the Blue Ring */}
              <circle
                cx="500"
                cy="500"
                r="298"
                fill="none"
                stroke="#ffffff"
                strokeWidth="4"
              />

              {/* 6. Main Inner Graphic Backdrop (Turquoise circle) */}
              <circle
                cx="500"
                cy="500"
                r="294"
                fill="#009be1"
                stroke="#ffffff"
                strokeWidth="2"
              />
              <circle cx="500" cy="500" r="275" fill="#0388c7" />

              {/* Curves for circular text alignment */}
              {/* Path for 'OESTE' arched on top (left-to-right) */}
              <path
                id="topArchPath"
                d="M 160,500 A 340,340 0 0,1 840,500"
                fill="none"
              />

              {/* Path for 'OESTE HUNTER' arched on bottom (right-to-left) */}
              <path
                id="bottomArchPath"
                d="M 840,500 A 340,340 0 0,1 160,500"
                fill="none"
              />

              {/* Arched Texts */}
              <text
                fontFamily="'Inter', sans-serif"
                fontWeight="900"
                fontSize="75"
                fill="#ffffff"
                letterSpacing="18"
              >
                <textPath
                  href="#topArchPath"
                  startOffset="50%"
                  textAnchor="middle"
                >
                  OESTE
                </textPath>
              </text>

              <text
                fontFamily="'Inter', sans-serif"
                fontWeight="900"
                fontSize="44"
                fill="#ffffff"
                letterSpacing="14"
              >
                <textPath
                  href="#bottomArchPath"
                  startOffset="50%"
                  textAnchor="middle"
                >
                  OESTE HUNTER
                </textPath>
              </text>

              {/* ======================================= */}
              {/* WOLF HEAD INTERIOR ELEMENT MASCOT ART   */}
              {/* ======================================= */}
              <g id="wolfMascot" transform="translate(0, -35)">
                {/* Wolf Ears Behind */}
                {/* Left ear dark back */}
                <polygon
                  points="350,330 435,420 380,480"
                  fill="#020f2b"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                {/* Left ear internal blue */}
                <polygon points="365,345 425,415 385,465" fill="#0096e6" />

                {/* Right ear dark back */}
                <polygon
                  points="650,330 565,420 620,480"
                  fill="#020f2b"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                {/* Right ear internal blue */}
                <polygon points="635,345 575,415 615,465" fill="#0096e6" />

                {/* Wolf Forehead and Cheek structure */}
                {/* Base Head polygon */}
                <polygon
                  points="500,380 340,500 370,625 500,680 630,625 660,500"
                  fill="#03112b"
                />

                {/* White outer framing highlights (Cheek fur) */}
                {/* Left cheek outer fluff */}
                <polygon
                  points="340,500 310,560 385,585"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="310,560 330,630 400,610"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Right cheek outer fluff */}
                <polygon
                  points="660,500 690,560 615,585"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="690,560 670,630 600,610"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Intermediate Blue Shadows Cheeks */}
                <polygon points="340,500 385,585 410,515" fill="#0a3c7c" />
                <polygon points="660,500 615,585 590,515" fill="#0a3c7c" />

                {/* Side Dark fur shades */}
                <polygon
                  points="410,515 385,585 440,590 460,530"
                  fill="#00183b"
                />
                <polygon
                  points="590,515 615,585 560,590 540,530"
                  fill="#00183b"
                />

                {/* Center forehead wolf shield (cyan core) */}
                <polygon points="500,380 460,470 500,510" fill="#00bdff" />
                <polygon points="500,380 540,470 500,510" fill="#00bdff" />

                <polygon points="500,395 470,470 500,500" fill="#ffffff" />
                <polygon points="500,395 530,470 500,500" fill="#ffffff" />

                {/* Wolf Eyes Areas (Black framing masks) */}
                <polygon
                  points="420,490 470,510 460,535 410,515"
                  fill="#010614"
                />
                <polygon
                  points="580,490 530,510 540,535 590,515"
                  fill="#010614"
                />

                {/* Fierce Cyan Eyes */}
                <polygon
                  points="432,498 462,510 452,525 430,512"
                  fill="url(#wolfEyeGrad)"
                  filter="url(#eyeGlow)"
                />
                <polygon
                  points="568,498 538,510 548,525 570,512"
                  fill="url(#wolfEyeGrad)"
                  filter="url(#eyeGlow)"
                />

                {/* Wolf Nose Bridge */}
                <polygon
                  points="500,510 460,530 475,590 500,610"
                  fill="#020f26"
                />
                <polygon
                  points="500,510 540,530 525,590 500,610"
                  fill="#020f26"
                />

                <polygon
                  points="500,510 480,530 485,585 500,600"
                  fill="#0080cf"
                />
                <polygon
                  points="500,510 520,530 515,585 500,600"
                  fill="#0080cf"
                />

                {/* Muzzle (White muzzle side facets) */}
                <polygon
                  points="500,610 440,590 445,635 500,665"
                  fill="url(#muzzleGrad)"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="500,610 560,590 555,635 500,665"
                  fill="url(#muzzleGrad)"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Black Nose Tip */}
                <polygon points="500,620 475,605 525,605" fill="#010614" />
                <polygon points="500,620 485,635 515,635" fill="#010614" />
                <polygon
                  points="475,605 525,605 515,635 485,635"
                  fill="#010614"
                />
                {/* Nose shine */}
                <circle cx="500" cy="612" r="3" fill="#ffffff" />
              </g>

              {/* Left and Right Side Banners - NPS and CAPTAÃ‡ÃƒO DE ALUNOS */}
              {/* LEFT BANNER (NPS) */}
              <g id="leftBanner">
                <polygon
                  points="6,505 130,505 110,615 6,615 30,560"
                  fill="#07336e"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                <polygon
                  points="12,515 120,515 104,605 12,605"
                  fill="#0b4594"
                />

                <text
                  x="63"
                  y="578"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="46"
                  fill="#ffffff"
                  letterSpacing="1"
                >
                  NPS
                </text>
              </g>

              {/* RIGHT BANNER (CAPTAÃ‡ÃƒO DE ALUNOS) */}
              <g id="rightBanner">
                <polygon
                  points="994,505 870,505 890,615 994,615 970,560"
                  fill="#07336e"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                <polygon
                  points="988,515 880,515 896,605 988,605"
                  fill="#0b4594"
                />

                <text
                  x="934"
                  y="555"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="20"
                  fill="#ffffff"
                  letterSpacing="2"
                >
                  CAPTAÃ‡ÃƒO
                </text>
                <text
                  x="934"
                  y="583"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="18"
                  fill="#ffffff"
                  letterSpacing="1"
                >
                  DE ALUNOS
                </text>
              </g>

              {/* Giant Horizontal Bottom Ribbon - HUNTER */}
              <g id="hunterBanner" transform="translate(0, 10)">
                {/* Banner Ribbon shadow back folds */}
                <polygon points="180,685 240,685 220,625" fill="#01050e" />
                <polygon points="820,685 760,685 780,625" fill="#01050e" />

                {/* Front Main Banner Body */}
                <polygon
                  points="180,625 820,625 790,735 210,735"
                  fill="url(#bannerGrad)"
                  stroke="#ffffff"
                  strokeWidth="6"
                />

                {/* Inner stroke accent */}
                <polygon
                  points="195,635 805,635 780,725 220,725"
                  fill="none"
                  stroke="#2575fc"
                  strokeWidth="3.5"
                  opacity="0.8"
                />

                {/* Bold Athletics display text - HUNTER */}
                <text
                  x="500"
                  y="702"
                  textAnchor="middle"
                  fontFamily="'Impact', 'Arial Black', 'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="105"
                  fill="#ffffff"
                  letterSpacing="5"
                  filter="url(#glow)"
                >
                  HUNTER
                </text>
              </g>
            </svg>
          )}
        </div>
      </div>

      {/* APK Information Modal */}
      {showApkModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative overflow-hidden">
            {/* Header / Icon */}
            <div className="mx-auto w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-6 shadow-inner">
              <Download size={32} className="text-blue-600" />
            </div>

            <h3 className="text-xl font-black text-slate-800 text-center mb-2">
              Download do Arquivo APK
            </h3>
            <p className="text-sm text-slate-500 text-center font-medium leading-relaxed mb-6">
              O projeto nativo Android foi gerado e configurado usando
              Capacitor.
            </p>

            <div className="space-y-4 bg-slate-50 p-5 rounded-2xl border border-slate-100 mb-6">
              <div className="flex items-start gap-3">
                <div className="bg-emerald-100 text-emerald-600 p-1.5 rounded-lg shrink-0 mt-0.5">
                  <Smartphone size={18} />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-slate-800">
                    1. InstalaÃ§Ã£o Imediata (Via Chrome)
                  </h4>
                  <p className="text-xs text-slate-600 mt-1 font-medium">
                    Recomendado: Feche esta aba e clique em{" "}
                    <strong className="text-emerald-600">
                      "Instalar no Aparelho"
                    </strong>{" "}
                    na tela de login (usando o Google Chrome no seu celular)
                    para instalaÃ§Ã£o automÃ¡tica PWA/WebAPK direta no aparelho.
                  </p>
                </div>
              </div>

              <div className="w-full h-px bg-slate-200"></div>

              <div className="flex items-start gap-3">
                <div className="bg-amber-100 text-amber-600 p-1.5 rounded-lg shrink-0 mt-0.5">
                  <Download size={18} />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-slate-800">
                    2. Desenvolvedores (CompilaÃ§Ã£o Nativa)
                  </h4>
                  <p className="text-xs text-slate-600 mt-1 font-medium">
                    Devido as limitaÃ§Ãµes do ambiente Cloud, o arquivo{" "}
                    <strong className="font-bold">.apk</strong> real precisa ser
                    compilado localmente: Exporte os arquivos do app, abra a
                    pasta{" "}
                    <code className="bg-slate-200 px-1 py-0.5 rounded font-mono text-amber-800">
                      android/
                    </code>{" "}
                    no Android Studio e compile o APK.
                  </p>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={() => setShowApkModal(false)}
                className="flex-1 px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold transition-all cursor-pointer"
              >
                Voltar Ã  Tela Inicial
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function DashboardView({
  leads,
  planner,
  links,
  profile,
  onToast,
  campanhas,
  bomDia,
  forecast,
  periodos,
  metaDia,
  metaSM,
  metaCursos,
  qgLigacoes,
  users,
}: {
  leads: Lead[];
  planner: PlannerTask[];
  links: LinkUtil[];
  profile: UserProfile;
  onToast: (m: string, t?: "success" | "error") => void;
  campanhas: Campanha[];
  bomDia: BomDiaCaptacao[];
  forecast: ForecastCaptacao[];
  periodos: PeriodoCaptacao[];
  metaDia: MetaDia[];
  metaSM: MetaSM[];
  metaCursos: MetaCurso[];
  qgLigacoes: QgLigacao[];
  users: UserProfile[];
}) {
  const [isCustomizing, setIsCustomizing] = useState(false);
  const [linksSearchTerm, setLinksSearchTerm] = useState("");
  const [linksFilterLocal, setLinksFilterLocal] = useState("");
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallGuide, setShowInstallGuide] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const isStandalone =
      window.matchMedia("(display-mode: standalone)").matches ||
      (window.navigator as any).standalone ||
      document.referrer.includes("android-app://");
    setIsAppInstalled(isStandalone);

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener(
        "beforeinstallprompt",
        handleBeforeInstallPrompt,
      );
    };
  }, []);

  const handleInstallClick = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: { outcome: string }) => {
        if (choiceResult.outcome === "accepted") {
          console.log("User accepted the PWA install prompt");
          setIsAppInstalled(true);
        }
        setDeferredPrompt(null);
      });
    } else {
      setShowInstallGuide(true);
    }
  };

  const defaultWidgets = {
    stats: false,
    links: true,
    planner: true,
    campanhas: false,
    bomDia: true,
    forecast: true,
    periodo: true,
    qgLigacoes: true,
    metaSM: true,
    metaCursos: true,
    aniversarios: true,
  };
  const widgets = profile?.dashboardWidgets
    ? { ...defaultWidgets, ...profile.dashboardWidgets }
    : defaultWidgets;

  const currentMonthNum = new Date().getMonth() + 1; // 1-12
  const monthNamesPt = [
    "Janeiro",
    "Fevereiro",
    "MarÃ§o",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];
  const currentMonthName = monthNamesPt[currentMonthNum - 1];

  const currentDayNum = new Date().getDate();
  const checkIsToday = (dob: string) => {
    const parts = dob.split("-");
    if (parts.length !== 3) return false;
    return (
      parseInt(parts[2], 10) === currentDayNum &&
      parseInt(parts[1], 10) === currentMonthNum
    );
  };

  const birthdaysThisMonth = (users || [])
    .filter((u) => {
      if (u.blocked) return false;
      if (!u.dataNascimento) return false;
      const dateParts = u.dataNascimento.split("-");
      if (dateParts.length !== 3) return false;
      const birthMonth = parseInt(dateParts[1], 10);
      return birthMonth === currentMonthNum;
    })
    .sort((a, b) => {
      const dayA = parseInt(a.dataNascimento!.split("-")[2], 10);
      const dayB = parseInt(b.dataNascimento!.split("-")[2], 10);
      return dayA - dayB;
    });

  const today = new Date().toISOString().split("T")[0];
  const activePeriod = periodos.find(
    (p) => today >= p.inicioInscricao && today <= p.fimMatFin,
  );

  // Find meta for today, or find the latest meta as a fallback
  const todayEntry = metaDia.find((m) => m.data === today);
  const latestEntry =
    metaDia.length > 0
      ? [...metaDia].sort((a, b) => b.data.localeCompare(a.data))[0]
      : null;
  const activeMeta = todayEntry || latestEntry;

  const days = [
    "Segunda-feira",
    "TerÃ§a-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "SÃ¡bado",
    "Domingo",
  ];

  const toggleWidget = async (
    key: keyof NonNullable<UserProfile["dashboardWidgets"]>,
  ) => {
    try {
      const newWidgets = { ...widgets, [key]: !widgets[key] };
      await updateDoc(doc(db, COLLECTIONS.USERS, profile.uid), {
        dashboardWidgets: newWidgets,
      });
      onToast("PreferÃªncias salvas!");
    } catch (err: any) {
      onToast("Erro ao salvar preferÃªncias.", "error");
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex justify-between items-center">
        <h2 className="text-2xl font-bold text-slate-800">Dashboard</h2>
        <div className="flex items-center space-x-4">
          <button
            onClick={() => setIsCustomizing(true)}
            className="flex items-center space-x-2 px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-50 transition-all"
          >
            <Settings size={18} />
            <span>Personalizar</span>
          </button>
        </div>
      </div>

      {/* Android App Promotion Card */}
      {!isAppInstalled && (
        <div
          id="android-app-prompt-card"
          className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 p-6 rounded-3xl text-white shadow-xl relative overflow-hidden border border-slate-700/60 transition-all duration-300"
        >
          {/* Decorative design bubbles */}
          <div className="absolute -top-16 -right-16 w-32 h-32 bg-emerald-500/10 rounded-full blur-2xl"></div>
          <div className="absolute -bottom-16 -left-16 w-32 h-32 bg-blue-500/10 rounded-full blur-2xl"></div>

          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 relative z-10">
            <div className="flex items-start space-x-4">
              <div className="p-3.5 bg-slate-800/80 rounded-2xl border border-slate-700/80 flex items-center justify-center shadow-lg transform hover:scale-105 transition-all shrink-0">
                <img
                  src="/icon.svg"
                  alt="GestÃ£o Oeste"
                  className="w-12 h-12 rounded-xl object-contain"
                />
              </div>
              <div className="space-y-1">
                <div className="flex items-center space-x-2">
                  <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-400 text-[10px] font-black uppercase tracking-wider rounded-md border border-emerald-500/30">
                    InstalaÃ§Ã£o Android
                  </span>
                  <span className="text-xs text-slate-400 font-bold uppercase tracking-wider">
                    Suporte Offline Completo
                  </span>
                </div>
                <h3 className="text-xl font-black tracking-tight leading-none text-white">
                  Instalar Aplicativo GestÃ£o Oeste no Android
                </h3>
                <p className="text-sm text-slate-300 max-w-2xl mt-1.5 leading-relaxed font-semibold">
                  Trabalhe de qualquer lugar! FaÃ§a pedidos de insumos e
                  visualize dados{" "}
                  <strong className="text-emerald-400 font-bold">
                    com ou sem internet
                  </strong>
                  . Ao voltar a ter conexÃ£o, o sistema sincroniza
                  automaticamente com o servidor.
                </p>
              </div>
            </div>

            <div className="flex items-center space-x-3 shrink-0 self-end md:self-center">
              <button
                onClick={handleInstallClick}
                className="flex items-center space-x-2 px-5 py-3 bg-emerald-500 hover:bg-emerald-600 active:transform active:scale-95 text-slate-950 font-extrabold text-sm rounded-xl shadow-lg shadow-emerald-500/20 transition-all cursor-pointer"
              >
                <Smartphone size={18} />
                <span>Instalar Aplicativo</span>
              </button>
              <button
                onClick={() => setShowInstallGuide(!showInstallGuide)}
                className="flex items-center space-x-2 px-4 py-3 bg-white/10 hover:bg-white/15 text-slate-100 font-bold text-sm rounded-xl transition-all cursor-pointer"
              >
                <span>InstruÃ§Ãµes</span>
              </button>
            </div>
          </div>

          {/* Expanded Step-by-Step Installation Guide */}
          {showInstallGuide && (
            <div className="mt-6 pt-6 border-t border-slate-700/60 grid grid-cols-1 md:grid-cols-3 gap-4 text-sm animate-fade-in">
              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  1
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Chrome size={14} className="text-emerald-400" /> No Google
                  Chrome
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Abra este site no seu aparelho Android utilizando o navegador{" "}
                  <strong className="text-emerald-400 font-bold">
                    Google Chrome
                  </strong>
                  .
                </p>
              </div>

              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  2
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Smartphone size={14} className="text-emerald-400" /> Menu de
                  OpÃ§Ãµes
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Toque nos{" "}
                  <strong className="text-white font-bold">
                    trÃªs pontinhos (â‹®)
                  </strong>{" "}
                  localizados no canto superior direito do navegador Chrome.
                </p>
              </div>

              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  3
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Download size={14} className="text-emerald-400" /> Instalar
                  App
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Selecione{" "}
                  <strong className="text-emerald-400 font-bold">
                    "Instalar aplicativo"
                  </strong>{" "}
                  ou{" "}
                  <strong className="text-emerald-400 font-bold">
                    "Adicionar Ã  tela de inÃ­cio"
                  </strong>
                  . Um atalho oficial serÃ¡ criado no seu telefone!
                </p>
              </div>

              <div className="col-span-1 md:col-span-3 flex justify-end mt-2 animate-fade-in">
                <button
                  onClick={() => setShowInstallGuide(false)}
                  className="px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-xs text-slate-300 rounded-lg hover:text-white transition-all font-bold cursor-pointer border border-slate-700"
                >
                  Fechar InstruÃ§Ãµes
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {(() => {
        const todayDateObj = new Date(today + "T12:00:00Z");
        const dayOfWeek = todayDateObj.getUTCDay();
        const startOfWeek = new Date(todayDateObj);
        startOfWeek.setUTCDate(todayDateObj.getUTCDate() - dayOfWeek);
        const endOfWeek = new Date(startOfWeek);
        endOfWeek.setUTCDate(startOfWeek.getUTCDate() + 6);
        const startOfWeekStr = startOfWeek.toISOString().split("T")[0];
        const endOfWeekStr = endOfWeek.toISOString().split("T")[0];

        const thisWeekMetas = metaDia.filter(m => m.data >= startOfWeekStr && m.data <= endOfWeekStr);
        if (thisWeekMetas.length === 0) return null;

        const weekTotYTD = thisWeekMetas.reduce((acc, item) => acc + item.ytdPresencial + item.ytdSemipresencial + item.ytdDigital, 0);
        const weekTotReal = thisWeekMetas.reduce((acc, item) => acc + item.realizadoPresencial + item.realizadoSemipresencial + item.realizadoDigital, 0);
        const weekTotAA = thisWeekMetas.reduce((acc, item) => acc + item.aaPresencial + item.aaSemipresencial + item.aaDigital, 0);

        let statusText = "Abaixo da Meta";
        let statusColor = "bg-rose-50 text-rose-600 border-rose-100";
        if (weekTotYTD > 0 && weekTotReal > weekTotYTD) {
          statusText = "Meta Superada!";
          statusColor = "bg-emerald-50 text-emerald-600 border-emerald-100";
        } else if (weekTotYTD > 0 && weekTotReal === weekTotYTD) {
          statusText = "Meta Atingida";
          statusColor = "bg-blue-50 text-blue-600 border-blue-100";
        }

        return (
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 space-y-4 mb-6">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center">
              <div>
                <div className="flex items-center space-x-2 text-slate-900">
                  <Target size={20} className="text-indigo-600" />
                  <h3 className="text-lg font-bold">
                    Acompanhamento de Meta Semanal
                  </h3>
                </div>
                <p className="text-xs text-slate-400 font-medium mt-1">
                  Semana: <span className="font-bold">{new Date(startOfWeekStr + "T12:00:00Z").toLocaleDateString("pt-BR")}</span> a <span className="font-bold">{new Date(endOfWeekStr + "T12:00:00Z").toLocaleDateString("pt-BR")}</span>
                </p>
              </div>
              <span className={cn("px-3 py-1.5 rounded-full text-xs font-bold border mt-2 sm:mt-0", statusColor)}>
                {statusText}
              </span>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 pt-2">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-slate-500 mb-1">Ano Anterior (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-slate-700">{weekTotAA}</span>
                </div>
              </div>
              <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-blue-600 mb-1">Meta (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-blue-700">{weekTotYTD}</span>
                </div>
              </div>
              <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-emerald-600 mb-1">Realizado (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-emerald-700">{weekTotReal}</span>
                </div>
              </div>
              <div className="bg-amber-50/50 p-4 rounded-2xl border border-amber-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-amber-600 mb-1">Atingimento</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-amber-700">
                    {weekTotYTD > 0 ? ((weekTotReal / weekTotYTD) * 100).toFixed(1) : 0}%
                  </span>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {activeMeta && (
        <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 space-y-4">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center">
            <div>
              <div className="flex items-center space-x-2 text-slate-900">
                <Target size={20} className="text-blue-600" />
                <h3 className="text-lg font-bold">
                  Acompanhamento de Meta DiÃ¡ria
                </h3>
              </div>
              <p className="text-xs text-slate-400 font-medium mt-1">
                Referente ao dia:{" "}
                <span className="font-bold">
                  {new Date(activeMeta.data + "T00:00:00").toLocaleDateString(
                    "pt-BR",
                  )}
                </span>
                {activeMeta.data === today
                  ? " (Hoje)"
                  : " (Ãšltima meta registrada)"}
              </p>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;

              let statusText = "Abaixo da Meta";
              let statusColor = "bg-rose-50 text-rose-600 border-rose-100";
              if (totReal > totYTD) {
                statusText = "Meta Superada!";
                statusColor =
                  "bg-emerald-50 text-emerald-600 border-emerald-100";
              } else if (totReal === totYTD) {
                statusText = "Meta Atingida";
                statusColor = "bg-blue-50 text-blue-600 border-blue-100";
              }

              return (
                <span
                  className={cn(
                    "px-3 py-1.5 rounded-full text-xs font-bold border mt-2 sm:mt-0",
                    statusColor,
                  )}
                >
                  {statusText}
                </span>
              );
            })()}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-4 pt-2">
            <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
              <span className="text-[10px] font-bold text-slate-400 uppercase">
                Boletos NecessÃ¡rios (YTD)
              </span>
              <span className="text-2xl font-black text-slate-800 mt-2">
                {activeMeta.ytdPresencial +
                  activeMeta.ytdSemipresencial +
                  activeMeta.ytdDigital}
              </span>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;

              let color = "text-rose-600";
              if (totReal > totYTD) color = "text-emerald-600";
              else if (totReal === totYTD) color = "text-blue-600";

              return (
                <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-400 uppercase">
                    Total Realizado
                  </span>
                  <span className={cn("text-2xl font-black mt-2", color)}>
                    {totReal}
                  </span>
                </div>
              );
            })()}

            <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
              <span className="text-[10px] font-bold text-slate-400 uppercase">
                Ano Anterior (A.A)
              </span>
              <span className="text-2xl font-black text-slate-500 mt-2">
                {activeMeta.aaPresencial +
                  activeMeta.aaSemipresencial +
                  activeMeta.aaDigital}
              </span>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;
              const pct = totYTD > 0 ? (totReal / totYTD) * 100 : 0;

              let pctBg = "bg-rose-50 text-rose-700";
              if (totReal > totYTD) pctBg = "bg-emerald-50 text-emerald-700";
              else if (totReal === totYTD) pctBg = "bg-blue-50 text-blue-700";

              return (
                <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-400 uppercase">
                    Aproveitamento
                  </span>
                  <div className="flex items-baseline space-x-2 mt-2">
                    <span
                      className={cn(
                        "text-xl font-extrabold px-2.5 py-0.5 rounded-lg",
                        pctBg,
                      )}
                    >
                      {pct.toFixed(0)}%
                    </span>
                  </div>
                </div>
              );
            })()}
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-2 border-t border-slate-100">
            {[
              {
                label: "Presencial",
                ytd: activeMeta.ytdPresencial,
                real: activeMeta.realizadoPresencial,
                aa: activeMeta.aaPresencial,
                accent: "border-l-4 border-l-blue-500",
              },
              {
                label: "Semipresencial",
                ytd: activeMeta.ytdSemipresencial,
                real: activeMeta.realizadoSemipresencial,
                aa: activeMeta.aaSemipresencial,
                accent: "border-l-4 border-l-orange-500",
              },
              {
                label: "Digital",
                ytd: activeMeta.ytdDigital,
                real: activeMeta.realizadoDigital,
                aa: activeMeta.aaDigital,
                accent: "border-l-4 border-l-indigo-500",
              },
              {
                label: "Curso TÃ©cnico",
                ytd: activeMeta.ytdTecnico || 0,
                real: activeMeta.realizadoTecnico || 0,
                aa: activeMeta.aaTecnico || 0,
                accent: "border-l-4 border-l-emerald-500",
              },
            ].map((modal, idx) => {
              let color = "text-rose-600";
              if (modal.real > modal.ytd) color = "text-emerald-600";
              else if (modal.real === modal.ytd) color = "text-blue-600";

              return (
                <div
                  key={idx}
                  className={cn(
                    "bg-slate-50/30 p-3 rounded-xl border border-slate-100 flex justify-between items-center",
                    modal.accent,
                  )}
                >
                  <div>
                    <span className="text-xs font-bold text-slate-700">
                      {modal.label}
                    </span>
                    <span className="block text-[10px] text-slate-400">
                      Ano Ant: {modal.aa}
                    </span>
                  </div>
                  <div className="text-right">
                    <span className="text-[10px] font-bold text-slate-400 block uppercase">
                      Meta / Real
                    </span>
                    <span className="text-xs font-semibold text-slate-600">
                      {modal.ytd}
                    </span>
                    <span className="mx-1 text-slate-300">/</span>
                    <span className={cn("text-xs font-bold", color)}>
                      {modal.real}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Aniversariantes do MÃªs Widget */}
      {widgets.aniversarios !== false && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-rose-500 mb-6">
            <Cake size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Aniversariantes do MÃªs ({currentMonthName})
            </h3>
          </div>
          {birthdaysThisMonth.length > 0 ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-4">
              {birthdaysThisMonth.map((u) => {
                const bday = parseInt(u.dataNascimento!.split("-")[2], 10);
                const isToday = checkIsToday(u.dataNascimento!);
                return (
                  <div
                    key={u.uid}
                    className={cn(
                      "p-4 rounded-2xl border transition-all flex items-center justify-between",
                      isToday
                        ? "bg-rose-50/50 border-rose-200 shadow-sm shadow-rose-50"
                        : "bg-slate-50/50 border-slate-100 hover:border-slate-200",
                    )}
                  >
                    <div className="flex items-center space-x-3 overflow-hidden">
                      <div
                        className={cn(
                          "w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shrink-0",
                          isToday
                            ? "bg-rose-600 text-white animate-bounce"
                            : "bg-blue-50 text-blue-600",
                        )}
                      >
                        {u.name.charAt(0).toUpperCase()}
                      </div>
                      <div className="overflow-hidden">
                        <h4 className="font-bold text-slate-800 text-sm truncate">
                          {u.name}
                        </h4>
                        <p className="text-[10px] text-slate-400 font-semibold truncate uppercase tracking-wider">
                          {u.role}
                        </p>
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      {isToday ? (
                        <span className="inline-block px-2 py-1 bg-amber-100 text-amber-800 text-[10px] font-black rounded-lg uppercase tracking-wide animate-pulse">
                          Hoje! ðŸŽ‰
                        </span>
                      ) : (
                        <span className="text-xs font-black text-slate-500">
                          Dia {bday}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="text-center py-8 bg-slate-50/50 border border-dashed border-slate-200 rounded-2xl">
              <Cake size={32} className="mx-auto text-slate-300 mb-2" />
              <p className="text-sm text-slate-400 font-semibold">
                Nenhum aniversariante registrado neste mÃªs de {currentMonthName}
                .
              </p>
            </div>
          )}
        </section>
      )}

      {/* Meta SM Dashboard Card */}
      {widgets.metaSM && (metaSM && metaSM.length > 0) && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta SM
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            {[...metaSM].sort((a,b) => (b.semestre || "").localeCompare(a.semestre || "")).map(m => (
              <React.Fragment key={m.id}>
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <div>
                    <p className="text-xs font-bold text-slate-500 uppercase">Semestre {m.semestre}</p>
                    <p className="text-lg font-black text-slate-900">{m.realizado}</p>
                    <p className="text-xs text-slate-400">Realizado SM</p>
                  </div>
                  <div className="w-full text-right mt-2 pt-2 border-t border-slate-200">
                    <p className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                      Atualizado: {m.createdAt?.seconds ? new Date(m.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : m.createdAt ? new Date(m.createdAt).toLocaleDateString("pt-BR") : "-"}
                    </p>
                  </div>
                </div>
                <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Meta Dia</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaDia || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-blue-100 rounded-full h-1.5">
                    <div className="bg-blue-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaDia || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-blue-400">Meta: {m.metaDia}</p>
                    <p className="text-xs font-bold text-blue-500">{((m.realizado / (m.metaDia || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
                <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP A.A</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaAA || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-emerald-100 rounded-full h-1.5">
                    <div className="bg-emerald-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaAA || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-emerald-400">Ano Anterior: {m.metaAA}</p>
                    <p className="text-xs font-bold text-emerald-500">{((m.realizado / (m.metaAA || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
                <div className="bg-purple-50/50 p-4 rounded-2xl border border-purple-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Final</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaFinal || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-purple-100 rounded-full h-1.5">
                    <div className="bg-purple-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaFinal || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-purple-400">Meta: {m.metaFinal}</p>
                    <p className="text-xs font-bold text-purple-500">{((m.realizado / (m.metaFinal || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
              </React.Fragment>
            ))}
          </div>
        </section>
      )}

      {/* Meta Cursos Dashboard Card */}
      {widgets.metaCursos && (metaCursos && metaCursos.length > 0) && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta Cursos
            </h3>
          </div>
          <div className="grid grid-cols-1 xl:grid-cols-2 2xl:grid-cols-3 gap-6">
            {[...metaCursos].sort((a,b) => (b.semestre || "").localeCompare(a.semestre || "") || (a.curso || "").localeCompare(b.curso || "")).map(m => (
              <div key={m.id} className="bg-slate-50 border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                <div className="bg-indigo-600 p-4 flex justify-between items-center">
                  <h4 className="font-bold text-white text-sm uppercase tracking-wider">{m.curso}</h4>
                  <span className="bg-white/20 text-white text-[10px] font-bold px-2 py-1 rounded-full uppercase">{m.semestre}</span>
                </div>
                <div className="p-4">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-slate-400 font-bold uppercase tracking-tighter">
                        <th className="text-left pb-2">Indicador</th>
                        <th className="text-center pb-2">INSC</th>
                        <th className="text-center pb-2">FINANC</th>
                        <th className="text-center pb-2">ACAD</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {[
                        { label: "Meta Final", key: "metaFinal", color: "text-slate-600" },
                        { label: "Meta Dia", key: "metaDia", color: "text-slate-600" },
                        { label: "Ano Anterior", key: "metaAA", color: "text-slate-400" },
                        { label: "Realizado", key: "realizado", color: "text-emerald-600 font-bold" }
                      ].map((row, idx) => (
                        <tr key={idx} className="hover:bg-white/50 transition-colors">
                          <td className="py-2 font-semibold text-slate-500">{row.label}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.inscritos?.[row.key as keyof typeof m.inscritos] ?? m[row.key as keyof typeof m] ?? 0}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.financeiro?.[row.key as keyof typeof m.financeiro] ?? "-"}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.academico?.[row.key as keyof typeof m.academico] ?? "-"}</td>
                        </tr>
                      ))}
                      
                      {/* Calculated Gaps */}
                      {[
                        { label: "Gap Meta Dia", metaKey: "metaDia" },
                        { label: "Gap Ano Ant.", metaKey: "metaAA" },
                        { label: "Gap Final", metaKey: "metaFinal" }
                      ].map((row, idx) => (
                        <tr key={`calc-${idx}`} className="bg-slate-100/50">
                          <td className="py-1.5 font-bold text-[9px] text-slate-400 uppercase">{row.label}</td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              const real = m.inscritos?.realizado ?? m.realizado ?? 0;
                              const meta = m.inscritos?.[row.metaKey as keyof typeof m.inscritos] ?? m[row.metaKey as keyof typeof m] ?? 0;
                              const gap = real - meta;
                              if (gap === 0 && !m.inscritos) return <span className="text-slate-600">0</span>;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              if (!m.financeiro) return <span className="text-slate-400">-</span>;
                              const real = m.financeiro.realizado || 0;
                              const meta = m.financeiro[row.metaKey as keyof typeof m.financeiro] || 0;
                              const gap = real - meta;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              if (!m.academico) return <span className="text-slate-400">-</span>;
                              const real = m.academico.realizado || 0;
                              const meta = m.academico[row.metaKey as keyof typeof m.academico] || 0;
                              const gap = real - meta;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="bg-slate-100 p-2 text-right">
                  <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                    Atualizado: {m.createdAt?.seconds ? new Date(m.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : m.createdAt ? new Date(m.createdAt).toLocaleDateString("pt-BR") : "-"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      
      {/* Bom Dia CaptaÃ§Ã£o (Complete - All cards) */}
      {widgets.bomDia && bomDia.filter((b) => !b.oculto).length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center space-x-2 text-emerald-600">
              <Sun size={24} />
              <h3 className="text-xl font-bold text-slate-900">
                Bom Dia CaptaÃ§Ã£o
              </h3>
            </div>
          </div>
          <div className="grid grid-cols-1 xl:grid-cols-2 2xl:grid-cols-3 gap-6">
            {bomDia
              .filter((b) => !b.oculto)
              .map((card) => (
                <div
                  key={card.id}
                  className="bg-slate-50 rounded-2xl border border-slate-100 overflow-hidden"
                >
                  <div className="bg-emerald-600 p-4">
                    <h4 className="font-bold text-white text-sm uppercase tracking-wider">
                      {card.titulo}
                    </h4>
                  </div>
                  <div className="p-4">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-slate-400 font-bold uppercase tracking-tighter">
                          <th className="text-left pb-2">Indicador</th>
                          <th className="text-center pb-2">INSC</th>
                          <th className="text-center pb-2">MAT FIN</th>
                          <th className="text-center pb-2">MAT ACAD</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {[
                          {
                            label: "Meta Final",
                            data: card.metaFinal,
                            color: "text-slate-600",
                          },
                          {
                            label: "Meta Dia",
                            data: card.metaDia,
                            color: "text-slate-600",
                          },
                          {
                            label: "Ano Anterior",
                            data: card.anoAnterior,
                            color: "text-slate-400",
                          },
                          {
                            label: "Real",
                            data: card.real,
                            color: "text-emerald-600 font-bold",
                          },
                        ].map((row, idx) => (
                          <tr
                            key={idx}
                            className="hover:bg-white/50 transition-colors"
                          >
                            <td className="py-2 font-semibold text-slate-500">
                              {row.label}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.insc ?? 0}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.matFin ?? 0}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.matAcad ?? 0}
                            </td>
                          </tr>
                        ))}
                        {/* Calculated Rows */}
                        {[
                          {
                            label: "% Meta Dia",
                            calc: (m: keyof BomDiaMetrics) =>
                              card.metaDia && card.metaDia[m] > 0 && card.real
                                ? `${((card.real[m] / card.metaDia[m]) * 100).toFixed(0)}%`
                                : "0%",
                            color: "text-blue-600 font-bold",
                          },
                          {
                            label: "% Ano Ant.",
                            calc: (m: keyof BomDiaMetrics) =>
                              card.anoAnterior &&
                              card.anoAnterior[m] > 0 &&
                              card.real
                                ? `${((card.real[m] / card.anoAnterior[m]) * 100).toFixed(0)}%`
                                : "0%",
                            color: "text-slate-500 font-bold",
                          },
                          {
                            label: "Gap Meta Dia",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaDia) return 0;
                              const val = card.real[m] - card.metaDia[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaDia) return "text-slate-600";
                              const val = card.real[m] - card.metaDia[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                          {
                            label: "Gap Ano Ant.",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.anoAnterior) return 0;
                              const val = card.real[m] - card.anoAnterior[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.anoAnterior) return "text-slate-600";
                              const val = card.real[m] - card.anoAnterior[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                          {
                            label: "Gap Meta Final",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaFinal) return 0;
                              const val = card.real[m] - card.metaFinal[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaFinal) return "text-slate-600";
                              const val = card.real[m] - card.metaFinal[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                        ].map((row, idx) => (
                          <tr key={`calc-${idx}`} className="bg-slate-100/50">
                            <td className="py-1.5 font-bold text-[9px] text-slate-400 uppercase">
                              {row.label}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("insc")
                                  : row.color,
                              )}
                            >
                              {row.calc("insc")}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("matFin")
                                  : row.color,
                              )}
                            >
                              {row.calc("matFin")}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("matAcad")
                                  : row.color,
                              )}
                            >
                              {row.calc("matAcad")}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="bg-slate-100 p-2 text-right">
                    <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                      Atualizado: {card.createdAt?.seconds ? new Date(card.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : card.createdAt ? new Date(card.createdAt).toLocaleDateString("pt-BR") : "-"}
                    </span>
                  </div>
                </div>
              ))}
          </div>
        </section>
      )}

      {/* QG LigaÃ§Ãµes Widget */}
      {widgets.qgLigacoes !== false && qgLigacoes && qgLigacoes.length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-xl font-bold text-slate-900 flex items-center">
              <span className="bg-emerald-100 text-emerald-600 p-2 rounded-xl mr-3">
                <Phone size={20} />
              </span>
              QG LigaÃ§Ãµes
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
            {qgLigacoes.map((qg) => (
              <div
                key={qg.id}
                className="bg-slate-50 border border-slate-100 rounded-2xl p-4 flex flex-col justify-between items-start"
              >
                <div className="flex items-center space-x-2 text-emerald-600 mb-2 font-bold">
                  <Phone size={16} />
                  <span>{qg.nome}</span>
                </div>
                <div className="text-sm font-semibold text-slate-700">
                  {Array.isArray(qg.diaSemana)
                    ? qg.diaSemana.join(", ")
                    : qg.diaSemana}
                </div>
                <div className="text-xs text-slate-500 font-medium bg-emerald-100/50 px-2 py-1 rounded-md mt-2">
                  {qg.horario}
                </div>
                <div className="w-full text-right mt-2 pt-2 border-t border-slate-200/60">
                  <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                    Atualizado: {qg.createdAt?.seconds ? new Date(qg.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : qg.createdAt ? new Date(qg.createdAt).toLocaleDateString("pt-BR") : "-"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      {/* Forecasts (Complete - All cards) */}
      {widgets.forecast && forecast.filter((f) => !f.oculto).length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-xl font-bold text-slate-900">
              Forecasts de CaptaÃ§Ã£o
            </h3>
            <TrendingUp size={24} className="text-blue-600" />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
            {[...forecast]
              .filter((f) => !f.oculto)
              .sort((a, b) => a.nome.localeCompare(b.nome))
              .map((f) => {
                const percFech =
                  f.metaFechamento > 0
                    ? ((f.realizado / f.metaFechamento) * 100).toFixed(1)
                    : "0";
                const gapFech = f.realizado - f.metaFechamento;
                const dataFim = new Date(f.dataFim);
                const hoje = new Date();
                const diffTime = dataFim.getTime() - hoje.getTime();
                const diasRestantes = Math.max(
                  1,
                  Math.ceil(diffTime / (1000 * 60 * 60 * 24)),
                );
                const pacing = (Math.abs(gapFech) / diasRestantes).toFixed(1);

                return (
                  <div
                    key={f.id}
                    className="bg-slate-50 p-5 rounded-2xl border border-slate-100"
                  >
                    <div className="flex justify-between items-start mb-4">
                      <div>
                        <h4 className="font-bold text-slate-900">{f.nome}</h4>
                        <p className="text-[10px] text-slate-500 font-medium">
                          AtÃ©{" "}
                          {f.dataFim
                            .split("T")[0]
                            .split("-")
                            .reverse()
                            .join("/")}
                        </p>
                      </div>
                      <span
                        className={`text-[10px] font-bold px-2 py-1 rounded-full ${Number(percFech) >= 100 ? "bg-emerald-100 text-emerald-600" : "bg-blue-100 text-blue-600"}`}
                      >
                        {percFech}%
                      </span>
                    </div>

                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Realizado
                        </p>
                        <p className="text-lg font-bold text-emerald-600">
                          {f.realizado || 0}
                        </p>
                      </div>
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Meta
                        </p>
                        <p className="text-lg font-bold text-slate-700">
                          {f.metaFechamento || 0}
                        </p>
                      </div>
                    </div>

                    <div className="space-y-3 pt-4 border-t border-slate-200/60">
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-slate-400">
                          Meta Dia YTD
                        </span>
                        <span className="text-xs font-bold text-slate-700">
                          {f.metaDiaYTD || 0}
                        </span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-rose-400">
                          Gap Fechamento
                        </span>
                        <span
                          className={`text-xs font-bold ${gapFech > 0 ? "text-emerald-600" : gapFech < 0 ? "text-rose-600" : "text-slate-600"}`}
                        >
                          {gapFech > 0 ? "+" : ""}
                          {gapFech}
                        </span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-blue-400">
                          Pacing (por dia)
                        </span>
                        <span className="text-xs font-bold text-blue-600">
                          {pacing}
                        </span>
                      </div>
                      <div className="flex justify-between items-center bg-slate-200/50 p-2 rounded-lg mt-2">
                        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                          Dias Restantes
                        </span>
                        <span className="text-xs font-bold text-slate-800">
                          {diasRestantes}
                        </span>
                      </div>
                      <div className="w-full text-right mt-2 pt-2 border-t border-slate-200/60">
                        <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                          Atualizado: {f.createdAt?.seconds ? new Date(f.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : f.createdAt ? new Date(f.createdAt).toLocaleDateString("pt-BR") : "-"}
                        </span>
                      </div>
                    </div>
                  </div>
                );
              })}
          </div>
        </section>
      )}

      {widgets.periodo && periodos.length > 0 && (
        <section>
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-xl font-bold text-slate-900">
              PerÃ­odos da CaptaÃ§Ã£o
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {periodos.map((p) => {
              const isActive =
                today >= p.inicioInscricao && today <= p.fimMatFin;
              return (
                <div
                  key={p.id}
                  className={cn(
                    "bg-white p-5 rounded-3xl shadow-sm border transition-all",
                    isActive
                      ? "border-blue-500 ring-4 ring-blue-50"
                      : "border-slate-100",
                  )}
                >
                  <div className="flex items-center justify-between mb-4">
                    <div className="flex items-center space-x-3">
                      <div
                        className={cn(
                          "p-2 rounded-xl",
                          isActive
                            ? "bg-blue-600 text-white"
                            : "bg-blue-100 text-blue-600",
                        )}
                      >
                        <Calendar size={20} />
                      </div>
                      <h4 className="font-bold text-slate-900">{p.nome}</h4>
                    </div>
                    {isActive && (
                      <span className="px-2 py-1 bg-emerald-100 text-emerald-600 text-[10px] font-bold rounded-full uppercase">
                        Ativo
                      </span>
                    )}
                  </div>
                  <div className="space-y-4">
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          InscriÃ§Ã£o
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioInscricao)} -{" "}
                          {formatLocalDateString(p.fimInscricao)}
                        </p>
                      </div>
                      <span className="text-xs font-bold text-blue-600">
                        {getWorkingDaysBetween(
                          p.inicioInscricao,
                          p.fimInscricao,
                        )}{" "}
                        dias
                      </span>
                    </div>
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Mat Fin
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioMatFin)} -{" "}
                          {formatLocalDateString(p.fimMatFin)}
                        </p>
                      </div>
                      <div className="text-right">
                        <span className="text-xs font-bold text-blue-600 block">
                          {getWorkingDaysBetween(p.inicioMatFin, p.fimMatFin)}{" "}
                          dias Ãºteis
                        </span>
                        <span className="text-[10px] font-bold text-slate-500 block">
                          {getWorkingDaysRemaining(p.fimMatFin)} dias restantes
                        </span>
                      </div>
                    </div>
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Mat Acad
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioMatAcad)} -{" "}
                          {formatLocalDateString(p.fimMatAcad)}
                        </p>
                      </div>
                      <span className="text-xs font-bold text-blue-600">
                        {getWorkingDaysBetween(p.inicioMatAcad, p.fimMatAcad)}{" "}
                        dias
                      </span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {widgets.links && (
        <section>
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-4 gap-4">
            <h3 className="text-xl font-bold text-slate-900">Links Ãšteis</h3>
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <select
                value={linksFilterLocal}
                onChange={(e) => setLinksFilterLocal(e.target.value)}
                className="px-4 py-2 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm appearance-none"
              >
                <option value="">Todos os Locais</option>
                {Array.from(new Set(links.map(l => l.local).filter(Boolean))).map(local => (
                  <option key={local} value={local}>{local}</option>
                ))}
              </select>
              <div className="relative w-full sm:w-64">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                  <Search size={18} />
                </span>
                <input
                  type="text"
                  placeholder="Buscar links..."
                  value={linksSearchTerm}
                  onChange={(e) => setLinksSearchTerm(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm"
                />
              </div>
            </div>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {links
              .filter(link => 
                (!linksFilterLocal || link.local === linksFilterLocal) &&
                (!linksSearchTerm || 
                  link.nome.toLowerCase().includes(linksSearchTerm.toLowerCase()) || 
                  link.url.toLowerCase().includes(linksSearchTerm.toLowerCase()) || 
                  (link.local && link.local.toLowerCase().includes(linksSearchTerm.toLowerCase()))
                )
              )
              .map((link) => (
              <a
                key={link.id}
                href={link.url}
                target="_blank"
                rel="noopener noreferrer"
                className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex items-center space-x-3 hover:border-blue-200 hover:bg-blue-50/30 transition-all group"
              >
                <div className="p-2 bg-blue-100 text-blue-600 rounded-lg group-hover:bg-blue-600 group-hover:text-white transition-all shrink-0">
                  <ExternalLink size={18} />
                </div>
                <div className="flex flex-col overflow-hidden">
                  <span className="font-bold text-slate-700 truncate">
                    {link.nome}
                  </span>
                  {link.local && (
                    <span className="text-[10px] text-slate-400 font-medium truncate uppercase tracking-wider">
                      {link.local}
                    </span>
                  )}
                </div>
              </a>
            ))}
            {links.length === 0 && (
              <p className="text-slate-400 text-sm italic">
                Nenhum link cadastrado.
              </p>
            )}
          </div>
        </section>
      )}

      {widgets.planner && (
        <section>
          <h3 className="text-xl font-bold text-slate-900 mb-4">
            Planner da Semana
          </h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-4">
            {days.map((day) => {
              const tasks = planner.filter((t) => t.dayOfWeek === day);
              return (
                <div
                  key={day}
                  className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col"
                >
                  <div className="bg-slate-50 px-4 py-2 border-b border-slate-100">
                    <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                      {day.split("-")[0]}
                    </span>
                  </div>
                  <div className="p-4 flex-1 space-y-2">
                    {tasks.length > 0 ? (
                      tasks.map((task) => (
                        <div
                          key={task.id}
                          className="p-2 bg-blue-50 border-l-4 border-blue-600 rounded-r-lg"
                        >
                          <p className="text-xs font-bold text-blue-900">
                            {task.atendenteName}
                          </p>
                          <p className="text-[10px] text-blue-600 font-medium">
                            {task.baseName}
                          </p>
                        </div>
                      ))
                    ) : (
                      <div className="h-full flex items-center justify-center">
                        <p className="text-xs text-slate-300 italic">Folga</p>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {/* Customization Modal */}
      <AnimatePresence>
        {isCustomizing && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden"
            >
              <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-bold text-slate-900">
                  Personalizar Dashboard
                </h3>
                <button
                  onClick={() => setIsCustomizing(false)}
                  className="text-slate-400 hover:text-slate-600"
                >
                  <X size={24} />
                </button>
              </div>
              <div className="p-6 space-y-4">
                <p className="text-sm text-slate-500 mb-4">
                  Escolha quais blocos vocÃª deseja visualizar na sua tela
                  principal.
                </p>

                {[
                  {
                    id: "periodo",
                    label: "PerÃ­odos da CaptaÃ§Ã£o",
                    icon: Calendar,
                  },
                  { id: "bomDia", label: "Bom Dia CaptaÃ§Ã£o", icon: Sun },
                  { id: "forecast", label: "Forecasts", icon: TrendingUp },
                  { id: "links", label: "Links Ãšteis", icon: ExternalLink },
                  { id: "planner", label: "Planner da Semana", icon: Calendar },
                  { id: "qgLigacoes", label: "QG LigaÃ§Ãµes", icon: Phone },
                  {
                    id: "aniversarios",
                    label: "Aniversariantes do MÃªs",
                    icon: Cake,
                  },
                ].map((item) => (
                  <button
                    key={item.id}
                    onClick={() => toggleWidget(item.id as any)}
                    className={cn(
                      "w-full flex items-center justify-between p-4 rounded-xl border transition-all",
                      widgets[item.id as keyof typeof widgets]
                        ? "bg-blue-50 border-blue-200 text-blue-900"
                        : "bg-white border-slate-100 text-slate-500",
                    )}
                  >
                    <div className="flex items-center space-x-3">
                      <item.icon size={20} />
                      <span className="font-bold">{item.label}</span>
                    </div>
                    <div
                      className={cn(
                        "w-10 h-6 rounded-full relative transition-all",
                        widgets[item.id as keyof typeof widgets]
                          ? "bg-blue-600"
                          : "bg-slate-200",
                      )}
                    >
                      <div
                        className={cn(
                          "absolute top-1 w-4 h-4 bg-white rounded-full transition-all",
                          widgets[item.id as keyof typeof widgets]
                            ? "left-5"
                            : "left-1",
                        )}
                      />
                    </div>
                  </button>
                ))}
              </div>
              <div className="p-6 bg-slate-50 border-t border-slate-100">
                <button
                  onClick={() => setIsCustomizing(false)}
                  className="w-full bg-slate-900 text-white font-bold py-3 rounded-xl hover:bg-slate-800 transition-all"
                >
                  ConcluÃ­do
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function CadastroView({
  onToast,
  profile,
  calendarioAcoes = [],
  uniqueUnidades = [],
}: {
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
  calendarioAcoes?: CalendarioAcao[];
  uniqueUnidades?: string[];
}) {
  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || "NÃ£o informado",
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };
  
  const [formData, setFormData] = useState({
    acao: "",
    acaoId: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    cursoInteresse: "",
  });
  const [loading, setLoading] = useState(false);
  const [activeForm, setActiveForm] = useState<"lead" | "promotor">("lead");
  const [promotorData, setPromotorData] = useState({
    nome: "",
    email: "",
    cpf: "",
    dataNascimento: "",
    phone: "",
    chavePix: "",
    unidade: "",
  });

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    // Duplicate check
    const cleanCpf = formData.cpf.replace(/\D/g, "");
    const cleanTelefone = formData.telefone.replace(/\D/g, "");

    if (cleanCpf) {
      const qCpf = query(
        collection(db, COLLECTIONS.LEADS),
        where("cpf", "==", cleanCpf),
      );
      const snapCpf = await getDocs(qCpf);
      if (!snapCpf.empty) {
        onToast(
          "AtenÃ§Ã£o: Este CPF jÃ¡ possui um lead cadastrado no sistema.",
          "error",
        );
        return;
      }
    } else if (cleanTelefone) {
      const qTel = query(
        collection(db, COLLECTIONS.LEADS),
        where("telefone", "==", cleanTelefone),
      );
      const snapTel = await getDocs(qTel);
      if (!snapTel.empty) {
        onToast(
          "AtenÃ§Ã£o: Este Telefone jÃ¡ possui um lead cadastrado no sistema.",
          "error",
        );
        return;
      }
    }

    setLoading(true);
    try {
      const newLeadData: any = {
        ...formData,
        cpf: cleanCpf,
        telefone: cleanTelefone,
        converted: false,
        createdAt: serverTimestamp(),
        promotorId: profile.uid,
        promotorName: profile.name,
        promotorRole: profile.role,
        unidade: profile.unidade || "",
        servidor: profile.servidor || "principal",
      };

      if (profile.linkadoA) {
        newLeadData.linkadoA = profile.linkadoA;
      }

      await addDoc(collection(db, COLLECTIONS.LEADS), newLeadData);

      if (newLeadData.acaoId && newLeadData.acaoId !== "manual") {
        try {
          const qLeads = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", newLeadData.acaoId),
          );
          const snapLeads = await getDocs(qLeads);
          await updateDoc(
            doc(db, COLLECTIONS.CALENDARIO_ACOES, newLeadData.acaoId),
            {
              leadsFeitos: snapLeads.size,
            },
          );
        } catch (error) {
          console.error("Error auto-updating action leadsCount:", error);
        }
      }

      onToast("Lead cadastrado com sucesso!");
      setFormData({
        acao: "",
        acaoId: "",
        nome: "",
        telefone: "",
        email: "",
        cpf: "",
        cursoInteresse: "",
      });
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handlePromotorSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    const cleanCpf = promotorData.cpf.replace(/\D/g, "");
    const cleanPhone = promotorData.phone.replace(/\D/g, "");
    const cleanEmail = promotorData.email.trim();

    if (!promotorData.nome || !cleanEmail || !cleanPhone) {
      onToast(
        "Por favor, preencha todos os campos obrigatÃ³rios (Nome, Email e Telefone).",
        "error",
      );
      return;
    }

    if (!cleanEmail.includes("@")) {
      onToast("Formato de email invÃ¡lido.", "error");
      return;
    }

    setLoading(true);
    try {
      // 1. Create promoter in Auth with standard base password using secondaryAuth
      const userCredential = await createUserWithEmailAndPassword(
        secondaryAuth,
        cleanEmail,
        "123456",
      );
      await updateProfile(userCredential.user, {
        displayName: `${promotorData.nome}|comercial`,
      });
      const newUid = userCredential.user.uid;

      // 2. Create profile matching promoter/rua rules
      const profileData: any = {
        uid: newUid,
        name: promotorData.nome,
        email: cleanEmail,
        cpf: cleanCpf,
        dataNascimento: promotorData.dataNascimento,
        role: ROLES.PROMOTOR_RUA, // 'Promotor/rua'
        servidor: "comercial", // specified for commercial
        phone: cleanPhone,
        unidade: promotorData.unidade,
        chavePix: promotorData.chavePix,
        blocked: false,
        mustChangePassword: true,
        linkadoA: profile.uid, // linked to the creator FDV
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      // 3. Save profile document
      await setDoc(doc(db, COLLECTIONS.USERS, newUid), profileData);

      // 4. Sign out from secondary auth to avoid trace
      await signOut(secondaryAuth);

      onToast(
        "Promotor/rua cadastrado com sucesso! Senha padrÃ£o: 123456",
        "success",
      );
      setPromotorData({
        nome: "",
        email: "",
        cpf: "",
        dataNascimento: "",
        phone: "",
        chavePix: "",
        unidade: "",
      });
      setActiveForm("lead");
    } catch (err: any) {
      console.error("Auth error details (Promoter Registration):", err);
      let errorMsg = err.message;
      if (
        err.code === "auth/email-already-in-use" ||
        err.message?.includes("email-already-in-use")
      ) {
        errorMsg = "Este email jÃ¡ estÃ¡ em uso.";
      } else if (
        err.code === "auth/weak-password" ||
        err.message?.includes("weak-password")
      ) {
        errorMsg =
          "A senha de cadastro padrÃ£o deve conter pelo menos 6 caracteres.";
      } else if (
        err.code === "auth/invalid-email" ||
        err.message?.includes("invalid-email")
      ) {
        errorMsg = "EndereÃ§o de email invÃ¡lido.";
      }
      onToast(`Erro ao criar promotor: ${errorMsg}`, "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-2xl mx-auto">
      <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
        {profile?.role === ROLES.FDV_COMERCIAL && (
          <div className="flex space-x-2 bg-slate-50 p-1.5 rounded-2xl mb-6 border border-slate-100">
            <button
              type="button"
              onClick={() => setActiveForm("lead")}
              className={`flex-1 py-3 text-xs font-bold rounded-xl transition-all cursor-pointer flex items-center justify-center space-x-2 ${
                activeForm === "lead"
                  ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20"
                  : "text-slate-400 hover:text-slate-700"
              }`}
            >
              <UserPlus size={16} />
              <span>Cadastrar Novo Lead</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveForm("promotor")}
              className={`flex-1 py-3 text-xs font-bold rounded-xl transition-all cursor-pointer flex items-center justify-center space-x-2 ${
                activeForm === "promotor"
                  ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20"
                  : "text-slate-400 hover:text-slate-700"
              }`}
            >
              <Users size={16} />
              <span>Cadastrar Promotor de Rua</span>
            </button>
          </div>
        )}

        {activeForm === "lead" ? (
          <>
            <h3 className="text-2xl font-bold text-slate-900 mb-6">
              Cadastrar Novo Lead
            </h3>
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="md:col-span-2 space-y-2">
                  <label className="block text-sm font-bold text-slate-700">
                    AÃ§Ã£o / Origem
                  </label>
                  {calendarioAcoes && calendarioAcoes.length > 0 ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <span className="block text-xs font-semibold text-slate-500 mb-1">
                          Selecionar do CalendÃ¡rio
                        </span>
                        <select
                          value={formData.acaoId || ""}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val === "manual") {
                              setFormData({
                                ...formData,
                                acaoId: "manual",
                                acao: "",
                              });
                            } else {
                              const matched = calendarioAcoes.find(
                                (a) => a.id === val,
                              );
                              setFormData({
                                ...formData,
                                acaoId: val,
                                acao: matched ? matched.nome : "",
                              });
                            }
                          }}
                          className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm bg-white"
                        >
                          <option value="">Selecione...</option>
                          {calendarioAcoes.map((act) => (
                            <option key={act.id} value={act.id}>
                              {act.nome} ({act.dataInicio})
                            </option>
                          ))}
                          <option value="manual">
                            Outro (Digitar manualmente)
                          </option>
                        </select>
                      </div>
                      {(formData.acaoId === "manual" || !formData.acaoId) && (
                        <div>
                          <span className="block text-xs font-semibold text-slate-500 mb-1">
                            Digitar Nome da AÃ§Ã£o/Origem
                          </span>
                          <input
                            type="text"
                            required={
                              !formData.acaoId || formData.acaoId === "manual"
                            }
                            value={formData.acao}
                            onChange={(e) =>
                              setFormData({ ...formData, acao: e.target.value })
                            }
                            className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
                            placeholder="Ex: Facebook, Panfletagem, etc."
                          />
                        </div>
                      )}
                    </div>
                  ) : (
                    <input
                      type="text"
                      required
                      value={formData.acao}
                      onChange={(e) =>
                        setFormData({ ...formData, acao: e.target.value })
                      }
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                      placeholder="Ex: Evento Junino, Facebook, etc."
                    />
                  )}
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome do Candidato
                  </label>
                  <input
                    type="text"
                    required
                    value={formData.nome}
                    onChange={(e) =>
                      setFormData({ ...formData, nome: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Nome completo"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Telefone (WhatsApp)
                  </label>
                  <input
                    type="tel"
                    required
                    value={formData.telefone}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        telefone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="DDD + NÃºmero"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    E-mail (Opcional)
                  </label>
                  <input
                    type="email"
                    value={formData.email}
                    onChange={(e) =>
                      setFormData({ ...formData, email: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    CPF (Opcional)
                  </label>
                  <input
                    type="text"
                    value={formData.cpf}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="000.000.000-00"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Curso de Interesse
                  </label>
                  <input
                    type="text"
                    value={formData.cursoInteresse}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        cursoInteresse: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Ex: AdministraÃ§Ã£o, Direito..."
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-200 transition-all disabled:opacity-50 flex items-center justify-center space-x-2"
              >
                <Plus size={20} />
                <span>{loading ? "Salvando..." : "Salvar Lead"}</span>
              </button>
            </form>
          </>
        ) : (
          <>
            <h3 className="text-2xl font-bold text-slate-900 mb-2">
              Cadastrar Promotor de Rua
            </h3>
            <p className="text-xs text-slate-500 mb-6 font-medium">
              Os promotores cadastrados por vocÃª ficarÃ£o automaticamente
              vinculados ao seu perfil de FDV e herdarÃ£o todas as regras de
              visualizaÃ§Ã£o do sistema.
            </p>

            <form onSubmit={handlePromotorSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome Completo *
                  </label>
                  <input
                    type="text"
                    required
                    value={promotorData.nome}
                    onChange={(e) =>
                      setPromotorData({ ...promotorData, nome: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Nome completo do promotor"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Email (Google institucional ou pessoal) *
                  </label>
                  <input
                    type="email"
                    required
                    value={promotorData.email}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        email: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Telefone / WhatsApp *
                  </label>
                  <input
                    type="tel"
                    required
                    value={promotorData.phone}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        phone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="(00) 00000-0000"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    CPF (Opcional)
                  </label>
                  <input
                    type="text"
                    value={promotorData.cpf}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="000.000.000-00"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Data de Nascimento (Opcional)
                  </label>
                  <input
                    type="date"
                    value={promotorData.dataNascimento}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        dataNascimento: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Unidade *
                  </label>
                  <select
                    required
                    value={promotorData.unidade}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        unidade: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm bg-white"
                  >
                    <option value="">Selecione uma unidade</option>
                    {uniqueUnidades.map((u) => (
                      <option key={u} value={u}>
                        {u}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Chave PIX (Opcional)
                  </label>
                  <input
                    type="text"
                    value={promotorData.chavePix}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        chavePix: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="CPF, E-mail, Telefone ou AleatÃ³ria"
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-200 transition-all disabled:opacity-50 flex items-center justify-center space-x-2"
              >
                <Plus size={20} />
                <span>
                  {loading ? "Cadastrando..." : "Cadastrar Promotor de Rua"}
                </span>
              </button>
            </form>
          </>
        )}
      </div>
    </div>
  );
}

function HistoricoView({
  leads,
  profile,
  onToast,
  users,
  whatsappMessages,
  botConfig,
  onSendBot,
  onMassSendBot,
  gap,
  basesRenovacao,
  calendarioAcoes = [],
  pedidosCursos = [],
  mensagensEnviadasLog = [],
}: {
  leads: Lead[];
  profile: UserProfile;
  onToast: (m: string, t?: "success" | "error") => void;
  users: UserProfile[];
  whatsappMessages: WhatsAppMessage[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
  gap: GapEntry[];
  basesRenovacao: BaseEntry[];
  calendarioAcoes?: CalendarioAcao[];
  pedidosCursos?: PedidoCursoEntry[];
  mensagensEnviadasLog?: MensagemEnviadaLog[];
}) {
  const [searchTerm, setSearchTerm] = useState("");
  const [courseFilter, setCourseFilter] = useState("");
  const [baseFilter, setBaseFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState("");
  const [promotorFilter, setPromotorFilter] = useState("");
  const [selectorOpen, setSelectorOpen] = useState(false);
  const [selectedLead, setSelectedLead] = useState<Lead | null>(null);
  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);
  const [massSelectorOpen, setMassSelectorOpen] = useState(false);
  const [isAddMsgModalOpen, setIsAddMsgModalOpen] = useState(false);
  const [newMsgData, setNewMsgData] = useState({ modelName: "", texto: "" });
  const [msgLoading, setMsgLoading] = useState(false);
  const [invalidLeadIds, setInvalidLeadIds] = useState<Set<string>>(new Set());
  const [blockedFilter, setBlockedFilter] = useState<
    "all" | "blocked" | "unblocked"
  >("all");

  const [editModalOpen, setEditModalOpen] = useState(false);
  const [editingLead, setEditingLead] = useState<Lead | null>(null);
  const [historicoSubTab, setHistoricoSubTab] = useState<
    "dashboard" | "lista" | "pedidos_cursos"
  >("dashboard");

  const [editFormData, setEditFormData] = useState({
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    cursoInteresse: "",
    acao: "",
    acaoId: "",
  });

  const handleVerificacao = () => {
    const invalidIds = new Set<string>();
    leads.forEach((lead) => {
      let match = false;

      if (
        gap.some(
          (g) =>
            (g.cpf &&
              lead.cpf &&
              g.cpf.replace(/\D/g, "") === lead.cpf.replace(/\D/g, "")) ||
            (g.telefone &&
              lead.telefone &&
              g.telefone.replace(/\D/g, "") ===
                lead.telefone.replace(/\D/g, "")) ||
            g.nome.toLowerCase().trim() === lead.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (
        !match &&
        basesRenovacao.some(
          (b) =>
            (b.cpf &&
              lead.cpf &&
              b.cpf.replace(/\D/g, "") === lead.cpf.replace(/\D/g, "")) ||
            (b.telefone &&
              lead.telefone &&
              b.telefone.replace(/\D/g, "") ===
                lead.telefone.replace(/\D/g, "")) ||
            b.nome.toLowerCase().trim() === lead.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (match) {
        invalidIds.add(lead.id);
      }
    });
    setInvalidLeadIds(invalidIds);
    onToast(
      `VerificaÃ§Ã£o concluÃ­da: ${invalidIds.size} leads jÃ¡ estÃ£o cadastrados em GAP/Base LÃ­quida.`,
      "success",
    );
  };

  const uniqueCursos = useMemo(() => {
    return Array.from(
      new Set(leads.map((l) => l.cursoInteresse).filter(Boolean)),
    ).sort();
  }, [leads]);

  const uniqueBases = useMemo(() => {
    return Array.from(new Set(leads.map((l) => l.acao).filter(Boolean))).sort();
  }, [leads]);

  const uniqueStatuses = [
    "Pendente",
    "Sem retorno",
    "Interessado",
    "NÃ£o Interessado",
    "Convertido",
    "Contato via Sales",
  ];

  const uniquePromotores = useMemo(() => {
    return Array.from(
      new Set(leads.map((l) => l.promotorName).filter(Boolean)),
    ).sort();
  }, [leads]);

  const isAdmin = [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ].includes(profile.role);

  const handleAddCustomMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMsgData.texto.trim()) return;
    setMsgLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
        tipo: "historico",
        texto: newMsgData.texto,
        nome: newMsgData.modelName || undefined,
        createdAt: serverTimestamp(),
      });
      onToast("Mensagem de histÃ³rico salva!");
      setNewMsgData({ modelName: "", texto: "" });
      setIsAddMsgModalOpen(false);
    } catch (err: any) {
      console.error("Erro ao salvar mensagem:", err);
      onToast(`Erro ao salvar mensagem: ${err.message}`, "error");
    } finally {
      setMsgLoading(false);
    }
  };

  const handleInsertDefaultHistoricoMessages = async () => {
    try {
      const existing = whatsappMessages.filter((m) => m.tipo === "historico");
      if (existing.length > 0) {
        if (
          !window.confirm(
            "JÃ¡ existem mensagens para HistÃ³rico. Deseja adicionar as mensagens padrÃµes mesmo assim?",
          )
        ) {
          return;
        }
      }

      const defaults = [
        "OlÃ¡ [nome], tudo bem? Vimos aqui seu interesse no curso de [curso]. Podemos te ajudar?",
        "Oi [nome], aqui Ã© da faculdade! Recebemos seu cadastro sobre o curso de [curso]. Qual o melhor horÃ¡rio para conversarmos?",
        "OlÃ¡ [nome]! Qual a sua dÃºvida sobre o curso de [curso]?",
      ];

      for (const texto of defaults) {
        await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
          tipo: "historico",
          texto,
          createdAt: serverTimestamp(),
        });
      }
      onToast("Mensagens padrÃµes de histÃ³rico inseridas!");
    } catch (err: any) {
      onToast("Erro ao inserir mensagens padrÃµes.", "error");
    }
  };

  const filteredLeads = useMemo(() => {
    return leads
      .filter((l) => {
        // Gestor Unidade filtering
        if (profile.role === "Gestor Unidade") {
          if (!profile.unidade || l.unidade !== profile.unidade) {
            return false;
          }
        }

        const matchesSearch =
          !searchTerm ||
          l.nome.toLowerCase().includes(searchTerm.toLowerCase()) ||
          l.telefone.includes(searchTerm) ||
          l.acao.toLowerCase().includes(searchTerm.toLowerCase());
        const matchesCourse =
          !courseFilter || l.cursoInteresse === courseFilter;
        const matchesBase =
          baseFilter.length === 0 || baseFilter.includes(l.acao);
        const matchesStatus = !statusFilter || l.status === statusFilter;
        const matchesPromotor =
          !promotorFilter || l.promotorName === promotorFilter;
        const isBlocked = invalidLeadIds.has(l.id);
        const matchesBlocked =
          blockedFilter === "all" ||
          (blockedFilter === "blocked" && isBlocked) ||
          (blockedFilter === "unblocked" && !isBlocked);
        return (
          matchesSearch &&
          matchesCourse &&
          matchesBase &&
          matchesStatus &&
          matchesPromotor &&
          matchesBlocked
        );
      })
      .sort(
        (a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0),
      );
  }, [
    leads,
    searchTerm,
    courseFilter,
    baseFilter,
    statusFilter,
    promotorFilter,
    blockedFilter,
    invalidLeadIds,
  ]);

  const stats = useMemo(() => {
    const total = filteredLeads.length;
    const conv = filteredLeads.filter((l) => l.converted).length;
    const userLeads = filteredLeads.filter(
      (l) => l.promotorId === profile.uid,
    ).length;

    // Stats by Course (Top 5)
    const courseGroups: Record<string, number> = {};
    filteredLeads.forEach((l) => {
      const c = l.cursoInteresse || "NÃ£o Informado";
      courseGroups[c] = (courseGroups[c] || 0) + 1;
    });
    const byCourse = Object.entries(courseGroups)
      .map(([name, count]) => ({
        name,
        count,
        percentage: total > 0 ? ((count / total) * 100).toFixed(1) : "0",
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // Stats by Status
    const statusGroups: Record<string, number> = {
      Pendente: 0,
      Convertido: 0,
      "Sem retorno": 0,
      Interessado: 0,
      "NÃ£o Interessado": 0,
      "Contato via Sales": 0,
    };
    filteredLeads.forEach((l) => {
      const s = l.converted ? "Convertido" : l.status || "Pendente";
      if (statusGroups[s] !== undefined) statusGroups[s] += 1;
      else statusGroups["Pendente"] += 1;
    });
    const byStatus = Object.entries(statusGroups).map(([name, count]) => ({
      name,
      count,
      percentage: total > 0 ? ((count / total) * 100).toFixed(1) : "0",
    }));

    return {
      total,
      conv,
      userLeads,
      rate: total > 0 ? ((conv / total) * 100).toFixed(1) : "0",
      byCourse,
      byStatus,
    };
  }, [filteredLeads, profile]);

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries((prev) => [...prev, id]);
    } else {
      setSelectedEntries((prev) => prev.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(
        filteredLeads.filter((l) => !invalidLeadIds.has(l.id)).map((l) => l.id),
      );
    } else {
      setSelectedEntries([]);
    }
  };

  const handleStatusChange = async (id: string, newStatus: string) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.LEADS, id), { status: newStatus });
      onToast("Status atualizado!");
    } catch (err: any) {
      handleFirestoreError(
        err,
        OperationType.UPDATE,
        `${COLLECTIONS.LEADS}/${id}`,
      );
      onToast("Erro ao atualizar status.", "error");
    }
  };

  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || 'NÃ£o informado',
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };

  const handleSendViaWhats = async (lead: Lead) => {
    try {
      await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
        leadId: lead.id,
        nome: lead.nome || "Sem Nome",
        telefone: lead.telefone || "",
        curso: lead.cursoInteresse || "NÃ£o informado",
        base: lead.acao || "HistÃ³rico Leads",
        tipoEnvio: "whats",
        dataHora: serverTimestamp(),
        usuarioId: profile.uid || "",
        usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
        unidade: lead.unidade || profile.unidade || "",
      });
      onToast(`Envio via WhatsApp para ${lead.nome} registrado no relatÃ³rio!`, "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar envio via WhatsApp.", "error");
    }
  };

  const handleSendViaMalaDireta = async (lead: Lead) => {
    try {
      await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
        leadId: lead.id,
        nome: lead.nome || "Sem Nome",
        telefone: lead.telefone || "",
        curso: lead.cursoInteresse || "NÃ£o informado",
        base: lead.acao || "HistÃ³rico Leads",
        tipoEnvio: "maladireta",
        dataHora: serverTimestamp(),
        usuarioId: profile.uid || "",
        usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
        unidade: lead.unidade || profile.unidade || "",
      });
      onToast(`Envio via Mala Direta para ${lead.nome} registrado no relatÃ³rio!`, "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar envio via Mala Direta.", "error");
    }
  };

  const handleBulkSendViaWhats = async () => {
    if (selectedEntries.length === 0) return;
    const selectedLeads = filteredLeads.filter((l) => selectedEntries.includes(l.id));
    try {
      for (const lead of selectedLeads) {
        await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
          leadId: lead.id,
          nome: lead.nome || "Sem Nome",
          telefone: lead.telefone || "",
          curso: lead.cursoInteresse || "NÃ£o informado",
          base: lead.acao || "HistÃ³rico Leads",
          tipoEnvio: "whats",
          dataHora: serverTimestamp(),
          usuarioId: profile.uid || "",
          usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
          unidade: lead.unidade || profile.unidade || "",
        });
      }
      onToast(`${selectedLeads.length} envios via WhatsApp registrados com sucesso!`, "success");
    } catch (e) {
      console.error(e);
      onToast("Erro ao registrar envios em massa via WhatsApp.", "error");
    }
  };

  const handleBulkSendViaMalaDireta = async () => {
    if (selectedEntries.length === 0) return;
    const selectedLeads = filteredLeads.filter((l) => selectedEntries.includes(l.id));
    try {
      for (const lead of selectedLeads) {
        await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
          leadId: lead.id,
          nome: lead.nome || "Sem Nome",
          telefone: lead.telefone || "",
          curso: lead.cursoInteresse || "NÃ£o informado",
          base: lead.acao || "HistÃ³rico Leads",
          tipoEnvio: "maladireta",
          dataHora: serverTimestamp(),
          usuarioId: profile.uid || "",
          usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
          unidade: lead.unidade || profile.unidade || "",
        });
      }
      onToast(`${selectedLeads.length} envios via Mala Direta registrados com sucesso!`, "success");
    } catch (e) {
      console.error(e);
      onToast("Erro ao registrar envios em massa via Mala Direta.", "error");
    }
  };

  const handleMoveToGap = async (lead: Lead) => {
    try {
      await addDoc(collection(db, COLLECTIONS.GAP), {
        nome: lead.nome,
        telefone: lead.telefone,
        matAcad: false,
        documentos: {},
        leadId: lead.id,
        createdAt: serverTimestamp(),
      });
      onToast("Candidato movido para o GAP!");
    } catch (err: any) {
      handleFirestoreError(err, OperationType.CREATE, COLLECTIONS.GAP);
      onToast("Erro ao mover para o GAP.", "error");
    }
  };

  const handleDeleteLead = async (id: string) => {
    if (
      !window.confirm("Tem certeza que deseja excluir este lead do histÃ³rico?")
    )
      return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.LEADS, id));
      onToast("Lead excluÃ­do com sucesso!", "success");
      setSelectedEntries((prev) => prev.filter((s) => s !== id));
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao excluir lead.", "error");
    }
  };

  const handleBulkDelete = async () => {
    if (
      !window.confirm(
        `Tem certeza que deseja excluir ${selectedEntries.length} lead(s) do histÃ³rico?`,
      )
    )
      return;
    try {
      const firestoreBatch = writeBatch(db);
      selectedEntries.forEach((id) => {
        firestoreBatch.delete(doc(db, COLLECTIONS.LEADS, id));
      });
      await firestoreBatch.commit();
      onToast(
        `${selectedEntries.length} lead(s) excluÃ­do(s) com sucesso!`,
        "success",
      );
      setSelectedEntries([]);
    } catch (err) {
      console.error(err);
      onToast("Erro ao excluir leads em massa.", "error");
    }
  };

  const handleExport = () => {
    const data = filteredLeads.map((l) => ({
      Nome: l.nome,
      Telefone: l.telefone,
      CPF: l.cpf || "",
      Email: l.email || "",
      Curso: l.cursoInteresse || "",
      Acao: l.acao,
      Promotor: l.promotorName,
      Status: l.converted ? "Convertido" : "Pendente",
      Data: l.createdAt?.seconds
        ? new Date(l.createdAt.seconds * 1000).toLocaleDateString()
        : "",
    }));
    exportToExcel(data, "Historico_Leads");
  };

  const handleExportMalaDireta = () => {
    const data = filteredLeads.map((l) => ({
      Nome: l.nome,
      Email: l.email || "",
    }));
    exportToExcel(data, "Mala_Direta_Leads");
  };

  const handleExportSMS = () => {
    const data = filteredLeads.map((l) => {
      let tel = l.telefone.replace(/\D/g, "");
      if (tel.length > 0 && !tel.startsWith("55")) {
        tel = "55" + tel;
      }
      return { Telefone: tel };
    });
    exportToCSV(data, "SMS_Leads");
  };

  const handleEditClick = (lead: Lead) => {
    setEditingLead(lead);
    setEditFormData({
      nome: lead.nome,
      telefone: lead.telefone,
      email: lead.email || "",
      cpf: lead.cpf || "",
      cursoInteresse: lead.cursoInteresse || "",
      acao: lead.acao,
      acaoId: lead.acaoId || "",
    });
    setEditModalOpen(true);
  };

  const handleSaveEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingLead) return;
    try {
      const prevAcaoId = editingLead.acaoId;
      const newAcaoId = editFormData.acaoId;

      await updateDoc(doc(db, COLLECTIONS.LEADS, editingLead.id), {
        nome: editFormData.nome,
        telefone: editFormData.telefone,
        cpf: editFormData.cpf,
        cursoInteresse: editFormData.cursoInteresse,
        acao: editFormData.acao,
        acaoId: newAcaoId || "",
      });

      if (prevAcaoId && prevAcaoId !== "manual" && prevAcaoId !== newAcaoId) {
        try {
          const qLeadsOld = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", prevAcaoId),
          );
          const snapOld = await getDocs(qLeadsOld);
          await updateDoc(doc(db, COLLECTIONS.CALENDARIO_ACOES, prevAcaoId), {
            leadsFeitos: snapOld.size,
          });
        } catch (err) {
          console.error(err);
        }
      }

      if (newAcaoId && newAcaoId !== "manual") {
        try {
          const qLeadsNew = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", newAcaoId),
          );
          const snapNew = await getDocs(qLeadsNew);
          await updateDoc(doc(db, COLLECTIONS.CALENDARIO_ACOES, newAcaoId), {
            leadsFeitos: snapNew.size,
          });
        } catch (err) {
          console.error(err);
        }
      }

      onToast("Lead atualizado com sucesso!", "success");
      setEditModalOpen(false);
      setEditingLead(null);
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao editar lead.", "error");
    }
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const batch = data.map((item) => {
          const rawStatus = String(getVal(item, "Status", "status") || "")
            .trim()
            .toLowerCase();
          const isConverted =
            rawStatus === "convertido" || getVal(item, "converted") === true;

          return {
            nome: String(getVal(item, "Nome", "nome") || "").trim(),
            telefone: String(
              getVal(item, "Telefone", "telefone") || "",
            ).replace(/\D/g, ""),
            cpf: String(getVal(item, "CPF", "cpf") || "").replace(/\D/g, ""),
            cursoInteresse: String(
              getVal(item, "Curso", "cursoInteresse", "curso") || "",
            ).trim(),
            acao: String(
              getVal(item, "Acao", "acao", "AÃ§Ã£o", "aÃ§Ã£o") || "ImportaÃ§Ã£o",
            ).trim(),
            promotorId: "import",
            promotorName: String(
              getVal(item, "Promotor", "promotorName") || "Sistema",
            ).trim(),
            converted: isConverted,
            unidade: profile.unidade || "",
            createdAt: serverTimestamp(),
          };
        });

        let imported = 0;
        let skipped = 0;
        const insertedCpfs = new Set();
        const insertedTels = new Set();

        for (const entry of batch) {
          const isDupCpf =
            entry.cpf &&
            (leads.some((l) => l.cpf === entry.cpf) ||
              insertedCpfs.has(entry.cpf));
          const isDupTel =
            entry.telefone &&
            (leads.some((l) => l.telefone === entry.telefone) ||
              insertedTels.has(entry.telefone));

          if (!isDupCpf && !isDupTel) {
            await addDoc(collection(db, COLLECTIONS.LEADS), entry);
            if (entry.cpf) insertedCpfs.add(entry.cpf);
            if (entry.telefone) insertedTels.add(entry.telefone);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} leads importados! ${skipped > 0 ? `${skipped} ignorados por duplicidade.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar leads.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      {/* Tab Navigation */}
      <div className="flex items-center gap-1 bg-white p-1 rounded-2xl shadow-sm border border-slate-100 w-fit">
        <button
          onClick={() => setHistoricoSubTab("dashboard")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "dashboard"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <BarChart3 size={18} />
          <span>Dashboard</span>
        </button>
        <button
          onClick={() => setHistoricoSubTab("lista")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "lista"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <List size={18} />
          <span>Lista de Leads</span>
        </button>
        <button
          onClick={() => setHistoricoSubTab("pedidos_cursos")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "pedidos_cursos"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <GraduationCap size={18} />
          <span>Pedidos de Cursos</span>
        </button>
      </div>

      {historicoSubTab === "dashboard" && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Total de Leads"
              value={stats.total}
              icon={Users}
              color="bg-blue-500"
            />
            <StatCard
              title="Convertidos"
              value={stats.conv}
              icon={CheckCircle2}
              color="bg-emerald-500"
            />
            <StatCard
              title="Taxa de Conv."
              value={`${stats.rate}%`}
              icon={TrendingUp}
              color="bg-purple-500"
            />
            <StatCard
              title="Meus Leads"
              value={stats.userLeads}
              icon={UserPlus}
              color="bg-amber-500"
            />
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <Target size={18} className="text-blue-500" />
                Status dos Leads
              </h3>
              <div className="space-y-3">
                {stats.byStatus.map((s) => (
                  <div key={s.name} className="space-y-1">
                    <div className="flex justify-between text-xs font-semibold">
                      <span className="text-slate-600 flex items-center gap-1.5">
                        <span
                          className={cn(
                            "w-2 h-2 rounded-full",
                            s.name === "Convertido" && "bg-emerald-400",
                            s.name === "Pendente" && "bg-amber-400",
                            s.name === "Interessado" && "bg-blue-400",
                            s.name === "NÃ£o Interessado" && "bg-rose-400",
                            s.name === "Sem retorno" && "bg-slate-400",
                            s.name === "Contato via Sales" && "bg-purple-400",
                          )}
                        />
                        {s.name}
                      </span>
                      <span className="text-slate-800 font-bold">
                        {s.count}{" "}
                        <span className="text-slate-400 font-normal">
                          ({s.percentage}%)
                        </span>
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                      <div
                        className={cn(
                          "h-full rounded-full transition-all",
                          s.name === "Convertido" && "bg-emerald-400",
                          s.name === "Pendente" && "bg-amber-400",
                          s.name === "Interessado" && "bg-blue-400",
                          s.name === "NÃ£o Interessado" && "bg-rose-400",
                          s.name === "Sem retorno" && "bg-slate-400",
                            s.name === "Contato via Sales" && "bg-purple-400",
                        )}
                        style={{ width: `${s.percentage}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <GraduationCap size={18} className="text-blue-500" />
                Cursos de Interesse (Top 5)
              </h3>
              <div className="space-y-3">
                {stats.byCourse.map((p) => (
                  <div key={p.name} className="space-y-1">
                    <div className="flex justify-between text-xs font-semibold">
                      <span className="text-slate-600 truncate max-w-[200px]">
                        {p.name}
                      </span>
                      <span className="text-slate-800 font-bold">
                        {p.count}{" "}
                        <span className="text-slate-400 font-normal">
                          ({p.percentage}%)
                        </span>
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                      <div
                        className="h-full bg-blue-500 rounded-full transition-all"
                        style={{ width: `${p.percentage}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {historicoSubTab === "lista" && (
        <>
          <div className="flex justify-between items-center">
            <h2 className="text-2xl font-bold text-slate-800">
              HistÃ³rico de Leads
            </h2>
            <div className="flex space-x-2">
              {[ROLES.ADMIN_MASTER, ROLES.LIDER_FDV].includes(profile.role) && (
                <button
                  onClick={handleVerificacao}
                  className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold"
                  title="Verificar se leads existem no GAP ou Base LÃ­quida"
                >
                  <Search size={18} />
                  <span>VerificaÃ§Ã£o</span>
                </button>
              )}
              <button
                onClick={() => setIsAddMsgModalOpen(true)}
                className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold"
              >
                <Plus size={18} />
                <span>Inserir Mensagens</span>
              </button>
              <button
                onClick={handleInsertDefaultHistoricoMessages}
                className="bg-slate-50 text-slate-400 px-3 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-100 transition-all text-[10px] font-bold"
                title="Inserir Mensagens PadrÃµes"
              >
                <MessageSquare size={14} />
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleImport}
                  className="hidden"
                />
              </label>
              <button
                onClick={handleExport}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
              <button
                onClick={handleExportMalaDireta}
                className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
              >
                <Mail size={18} />
                <span>Mala Direta</span>
              </button>
              <button
                onClick={handleExportSMS}
                className="bg-orange-50 text-orange-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-orange-100 transition-all text-sm font-bold shadow-sm"
              >
                <MessageSquare size={18} />
                <span>SMS (CSV)</span>
              </button>
            </div>
          </div>

          <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="p-6 border-b border-slate-100 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
              <h3 className="text-xl font-bold text-slate-900 whitespace-nowrap font-sans tracking-tight">
                Lista de Leads
              </h3>
              <div className="flex flex-wrap items-center gap-2 w-full xl:w-auto xl:justify-end">
                <div className="relative flex-1 min-w-[200px] xl:flex-none">
                  <Search
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                    size={16}
                  />
                  <input
                    type="text"
                    placeholder="Buscar por nome, telefone..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all text-xs"
                  />
                </div>
                <MultiSelect
                  options={uniqueBases}
                  selectedValues={baseFilter}
                  onChange={setBaseFilter}
                  placeholder="Todas as Origens / AÃ§Ãµes"
                  allLabel="Todas as Origens"
                />
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 max-w-[150px] lg:max-w-[200px] truncate"
                  value={courseFilter}
                  onChange={(e) => setCourseFilter(e.target.value)}
                >
                  <option value="">Todos os Cursos</option>
                  {uniqueCursos.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="">Todos os Status</option>
                  {uniqueStatuses.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                  value={blockedFilter}
                  onChange={(e) => setBlockedFilter(e.target.value as any)}
                >
                  <option value="all">VerificaÃ§Ã£o: Todos</option>
                  <option value="blocked">VerificaÃ§Ã£o: Bloqueados</option>
                  <option value="unblocked">VerificaÃ§Ã£o: Ativos</option>
                </select>
                {isAdmin && (
                  <select
                    className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 max-w-[150px] lg:max-w-[200px] truncate"
                    value={promotorFilter}
                    onChange={(e) => setPromotorFilter(e.target.value)}
                  >
                    <option value="">Todos os Promotores</option>
                    {uniquePromotores.map((p) => (
                      <option key={p} value={p}>
                        {p}
                      </option>
                    ))}
                  </select>
                )}
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                    <th className="px-6 py-4 w-12">
                      <input
                        type="checkbox"
                        checked={
                          filteredLeads.filter((l) => !invalidLeadIds.has(l.id))
                            .length > 0 &&
                          selectedEntries.length ===
                            filteredLeads.filter(
                              (l) => !invalidLeadIds.has(l.id),
                            ).length
                        }
                        onChange={(e) => toggleSelectAll(e.target.checked)}
                      />
                    </th>
                    <th className="px-3 py-4 w-12 text-slate-400">#</th>
                    <th className="px-6 py-4">Candidato</th>
                    <th className="px-6 py-4">AÃ§Ã£o / Origem</th>
                    <th className="px-6 py-4">Promotor</th>
                    <th className="px-6 py-4">Status</th>
                    <th className="px-6 py-4 flex flex-col gap-1.5">
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkSendViaWhats}
                          className="text-emerald-600 font-bold hover:underline py-1 px-2 bg-emerald-50 rounded-lg flex items-center gap-1 text-xs"
                          title="Registrar envio via Whats para todos selecionados"
                        >
                          <MessageSquare size={13} /> Whats ({selectedEntries.length})
                        </button>
                      )}
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkSendViaMalaDireta}
                          className="text-purple-600 font-bold hover:underline py-1 px-2 bg-purple-50 rounded-lg flex items-center gap-1 text-xs"
                          title="Registrar envio via Mala Direta para todos selecionados"
                        >
                          <Mail size={13} /> Mala Direta ({selectedEntries.length})
                        </button>
                      )}
                      {selectedEntries.length > 0 && botConfig.url && (
                        <button
                          onClick={() => setMassSelectorOpen(true)}
                          className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1 text-xs"
                        >
                          <Bot size={13} /> Bot Em Massa
                        </button>
                      )}
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkDelete}
                          className="text-rose-600 font-bold hover:underline py-1 px-2 bg-rose-50 rounded-lg flex items-center gap-1 text-xs"
                        >
                          <Trash2 size={13} /> Excluir ({selectedEntries.length})
                        </button>
                      )}
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredLeads.map((lead, index) => (
                    <tr
                      key={lead.id}
                      className={cn(
                        "hover:bg-slate-50/50 transition-all",
                        invalidLeadIds.has(lead.id) && "bg-rose-50/50",
                      )}
                    >
                      <td className="px-6 py-4">
                        <input
                          type="checkbox"
                          disabled={invalidLeadIds.has(lead.id)}
                          checked={selectedEntries.includes(lead.id)}
                          onChange={(e) =>
                            !invalidLeadIds.has(lead.id) &&
                            toggleSelect(lead.id, e.target.checked)
                          }
                        />
                      </td>
                      <td className="px-3 py-4 text-xs font-bold text-slate-400 font-mono">
                        {index + 1}
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex flex-col">
                          <span className="font-bold text-slate-900">
                            {lead.nome}
                          </span>
                          <span className="text-xs text-slate-500">
                            {formatPhone(lead.telefone)}
                          </span>
                          {lead.cursoInteresse && (
                            <span className="text-xs text-slate-600 font-medium">
                              Curso: {lead.cursoInteresse}
                            </span>
                          )}
                          {lead.empresa && (
                            <span className="text-[11px] text-indigo-600 font-bold mt-0.5 bg-indigo-50/60 border border-indigo-100/40 px-2 py-0.5 rounded-md self-start">
                              Empresa: {lead.empresa}
                            </span>
                          )}
                          {lead.cpf && (
                            <span className="text-xs text-slate-400">
                              CPF: {formatCPF(lead.cpf)}
                            </span>
                          )}
                          {(() => {
                            const leadLogs = (mensagensEnviadasLog || []).filter(
                              (l) => l.leadId === lead.id || (l.nome && lead.nome && l.nome.toLowerCase().trim() === lead.nome.toLowerCase().trim()) || (l.telefone && lead.telefone && l.telefone.replace(/\D/g, "") === lead.telefone.replace(/\D/g, ""))
                            );
                            const countW = leadLogs.filter((l) => l.tipoEnvio === "whats" || l.tipoEnvio === "bot_automatico").length;
                            const countM = leadLogs.filter((l) => l.tipoEnvio === "maladireta").length;
                            if (countW === 0 && countM === 0) return null;
                            return (
                              <div className="flex flex-wrap gap-1 mt-1.5">
                                {countW > 0 && (
                                  <span className="text-[10px] bg-emerald-100/90 text-emerald-800 font-bold px-1.5 py-0.5 rounded flex items-center gap-0.5" title="Total de envios via WhatsApp para este lead">
                                    <MessageSquare size={10} /> Whats: {countW}x
                                  </span>
                                )}
                                {countM > 0 && (
                                  <span className="text-[10px] bg-purple-100/90 text-purple-800 font-bold px-1.5 py-0.5 rounded flex items-center gap-0.5" title="Total de envios via Mala Direta para este lead">
                                    <Mail size={10} /> Mala: {countM}x
                                  </span>
                                )}
                              </div>
                            );
                          })()}
                        </div>
                      </td>
                      <td className="px-6 py-4 text-sm text-slate-600">
                        {lead.acao}
                      </td>
                      <td className="px-6 py-4 text-sm text-slate-600 font-medium">
                        {lead.promotorName}
                      </td>
                      <td className="px-6 py-4">
                        <select
                          value={lead.status || "Pendente"}
                          onChange={(e) =>
                            handleStatusChange(lead.id, e.target.value)
                          }
                          className={cn(
                            "px-3 py-1 rounded-full text-[10px] font-bold uppercase transition-all border-none focus:ring-0",
                            lead.status === "Convertido"
                              ? "bg-emerald-100 text-emerald-600"
                              : lead.status === "Interessado"
                                ? "bg-blue-100 text-blue-600"
                                : lead.status === "NÃ£o Interessado"
                                  ? "bg-rose-100 text-rose-600"
                                  : lead.status === "Sem retorno"
                                    ? "bg-slate-100 text-slate-600"
                                    : lead.status === "Contato via Sales" ? "bg-purple-100 text-purple-600"
                                    : "bg-amber-100 text-amber-600",
                          )}
                        >
                          <option value="Pendente">Pendente</option>
                          <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
                          <option value="Interessado">Interessado</option>
                          <option value="NÃ£o Interessado">
                            NÃ£o Interessado
                          </option>
                          <option value="Convertido">Convertido</option>
                        </select>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center space-x-3">
                          {!invalidLeadIds.has(lead.id) && (
                            <button
                              onClick={() => {
                                setSelectedLead(lead);
                                setSelectorOpen(true);
                              }}
                              className="inline-flex items-center space-x-1 text-emerald-600 font-bold text-sm hover:text-emerald-700"
                            >
                              <MessageSquare size={14} />
                              <span>WhatsApp</span>
                            </button>
                          )}
                          <button
                            onClick={() => handleSendViaWhats(lead)}
                            className="inline-flex items-center space-x-1 text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200/60 px-2.5 py-1 rounded-lg text-xs font-bold transition-all"
                            title="Registrar Envio via WhatsApp (pode clicar vÃ¡rias vezes para gerar histÃ³rico no relatÃ³rio)"
                          >
                            <MessageSquare size={13} />
                            <span>Enviado Whats</span>
                          </button>
                          <button
                            onClick={() => handleSendViaMalaDireta(lead)}
                            className="inline-flex items-center space-x-1 text-purple-700 bg-purple-50 hover:bg-purple-100 border border-purple-200/60 px-2.5 py-1 rounded-lg text-xs font-bold transition-all"
                            title="Registrar Envio via Mala Direta (pode clicar vÃ¡rias vezes para gerar histÃ³rico no relatÃ³rio)"
                          >
                            <Mail size={13} />
                            <span>Enviado Mala Direta</span>
                          </button>
                          <button
                            onClick={() => handleContatoViaSales(lead, 'Leads')}
                            className="inline-flex items-center space-x-1 text-sky-600 font-bold text-sm hover:text-sky-700 bg-sky-50 px-2 py-1 rounded-lg"
                            title="Registrar Contato via Sales"
                          >
                            <PhoneOutgoing size={14} />
                            <span>Sales</span>
                          </button>
                          {lead.status === "Convertido" && !invalidLeadIds.has(lead.id) && (
                            <button
                              onClick={() => handleMoveToGap(lead)}
                              className="text-purple-600 hover:text-purple-700 font-bold text-sm flex items-center space-x-1"
                              title="Mover para GAP AcadÃªmico"
                            >
                              <GraduationCap size={14} />
                            </button>
                          )}
                          <button
                            onClick={() => handleEditClick(lead)}
                            className="text-slate-400 hover:text-blue-600 transition-colors"
                            title="Editar Lead"
                          >
                            <Edit2 size={16} />
                          </button>
                          <button
                            onClick={() => handleDeleteLead(lead.id)}
                            className="text-slate-400 hover:text-rose-600 transition-colors"
                            title="Excluir Lead"
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {filteredLeads.length === 0 && (
                    <tr>
                      <td
                        colSpan={7}
                        className="px-6 py-12 text-center text-slate-400 italic"
                      >
                        Nenhum lead encontrado.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {historicoSubTab === "pedidos_cursos" && (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <div>
              <h2 className="text-2xl font-bold text-slate-800">
                Pedidos de Cursos
              </h2>
              <p className="text-slate-500 text-sm mt-1">
                Gere o link e acompanhe os cursos solicitados.
              </p>
            </div>
            <button
              onClick={() => {
                const url = `${window.location.origin}${window.location.pathname}?view=pedido-curso`;
                navigator.clipboard.writeText(url);
                onToast(
                  "Link copiado para a Ã¡rea de transferÃªncia!",
                  "success",
                );
              }}
              className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-200 flex items-center space-x-2"
            >
              <Copy size={18} />
              <span>Gerar Link do FormulÃ¡rio</span>
            </button>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider">
                    <th className="p-4 font-bold border-b border-slate-100">
                      Nome
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Telefone
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Data
                    </th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-50">
                  {pedidosCursos.length === 0 ? (
                    <tr>
                      <td
                        colSpan={4}
                        className="p-8 text-center text-slate-400"
                      >
                        Nenhum pedido registrado.
                      </td>
                    </tr>
                  ) : (
                    pedidosCursos.map((pedido) => (
                      <tr
                        key={pedido.id}
                        className="hover:bg-slate-50/50 transition-colors"
                      >
                        <td className="p-4 font-medium text-slate-700">
                          {pedido.nome}
                        </td>
                        <td className="p-4 text-slate-600">
                          {pedido.telefone}
                        </td>
                        <td className="p-4 text-slate-800 font-semibold">
                          {pedido.curso}
                        </td>
                        <td className="p-4 text-slate-500">
                          {pedido.createdAt
                            ? new Date(
                                pedido.createdAt.toDate(),
                              ).toLocaleDateString()
                            : "-"}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      <WhatsAppMessageSelector
        isOpen={selectorOpen}
        onClose={() => setSelectorOpen(false)}
        leadName={selectedLead?.nome || ""}
        leadCurso={selectedLead?.cursoInteresse || ""}
        messages={whatsappMessages.filter((m) => m.tipo === "historico")}
        onSelect={(msg) => {
          if (selectedLead) {
            window.open(getWhatsAppUrl(selectedLead.telefone, msg), "_blank");
          }
        }}
        botConfig={botConfig}
        onSendBot={(msg, contactName) => {
          if (selectedLead) {
            onSendBot(
              selectedLead.telefone,
              Array.isArray(msg) ? msg[0] : msg,
              contactName || selectedLead.nome,
            );
          }
        }}
      />

      <WhatsAppMessageSelector
        isOpen={massSelectorOpen}
        onClose={() => setMassSelectorOpen(false)}
        leadName="Candidatos"
        messages={whatsappMessages.filter((m) => m.tipo === "historico")}
        onSelect={(msg) => {
          // not used for mass send
        }}
        botConfig={botConfig}
        onSendBot={(msgTemplates) => {
          const templates = Array.isArray(msgTemplates) ? msgTemplates : [msgTemplates];
          const selectedLeadObjs = leads.filter(
            (l) => selectedEntries.includes(l.id) && !invalidLeadIds.has(l.id),
          );
          const messagesPayload = selectedLeadObjs.map((l, idx) => {
            const template = templates[idx % templates.length];
            return {
              telefone: l.telefone,
              message: replaceMessageVariables(template, l),
            };
          });
          onMassSendBot(messagesPayload);
          setMassSelectorOpen(false);
          setSelectedEntries([]);
        }}
        forceBotOnly={true}
      />

      <MessageTemplateModal
        isOpen={isAddMsgModalOpen}
        onClose={() => setIsAddMsgModalOpen(false)}
        tipo="historico"
        onToast={onToast}
        availableVariables={[
          { key: "[nome]", label: "Nome do Lead", previewValue: "JoÃ£o Silva" },
          {
            key: "[curso]",
            label: "Curso",
            previewValue: "Engenharia de Software",
          },
          {
            key: "[unidade]",
            label: "Unidade",
            previewValue: "Unidade Central",
          },
          {
            key: "[data_contato]",
            label: "Data",
            previewValue: new Date().toLocaleDateString("pt-BR"),
          },
          { key: "[saudacao]", label: "SaudaÃ§Ã£o", previewValue: "Bom dia" },
        ]}
      />

      {editModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.95 }}
            className="bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden"
          >
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
              <h3 className="text-xl font-bold text-slate-900">Editar Lead</h3>
              <button
                onClick={() => {
                  setEditModalOpen(false);
                  setEditingLead(null);
                }}
                className="text-slate-400 hover:text-slate-600"
              >
                <X size={24} />
              </button>
            </div>
            <form onSubmit={handleSaveEdit} className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome
                  </label>
                  <input
                    required
                    value={editFormData.nome}
                    onChange={(e) =>
                      setEditFormData({ ...editFormData, nome: e.target.value })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Telefone
                  </label>
                  <input
                    required
                    value={editFormData.telefone}
                    onChange={(e) => {
                      e.target.value = formatPhone(e.target.value);
                      setEditFormData({
                        ...editFormData,
                        telefone: e.target.value,
                      });
                    }}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    E-mail (Opcional)
                  </label>
                  <input
                    type="email"
                    value={editFormData.email}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        email: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    CPF
                  </label>
                  <input
                    value={editFormData.cpf}
                    onChange={(e) => {
                      e.target.value = formatCPF(e.target.value);
                      setEditFormData({ ...editFormData, cpf: e.target.value });
                    }}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Curso
                  </label>
                  <input
                    value={editFormData.cursoInteresse}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        cursoInteresse: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div className="col-span-2 space-y-2">
                  <label className="block text-xs font-bold text-slate-500">
                    Origem / AÃ§Ã£o
                  </label>
                  {calendarioAcoes && calendarioAcoes.length > 0 ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <span className="block text-[10px] font-semibold text-slate-400 mb-1">
                          Selecionar do CalendÃ¡rio
                        </span>
                        <select
                          value={editFormData.acaoId || ""}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val === "manual") {
                              setEditFormData({
                                ...editFormData,
                                acaoId: "manual",
                                acao: "",
                              });
                            } else {
                              const matched = calendarioAcoes.find(
                                (a) => a.id === val,
                              );
                              setEditFormData({
                                ...editFormData,
                                acaoId: val,
                                acao: matched ? matched.nome : "",
                              });
                            }
                          }}
                          className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm bg-white"
                        >
                          <option value="">Selecione...</option>
                          {calendarioAcoes.map((act) => (
                            <option key={act.id} value={act.id}>
                              {act.nome} ({act.dataInicio})
                            </option>
                          ))}
                          <option value="manual">
                            Outro (Digitar manualmente)
                          </option>
                        </select>
                      </div>
                      {(editFormData.acaoId === "manual" ||
                        !editFormData.acaoId) && (
                        <div>
                          <span className="block text-[10px] font-semibold text-slate-400 mb-1">
                            Digitar Nome da AÃ§Ã£o/Origem
                          </span>
                          <input
                            type="text"
                            required={
                              !editFormData.acaoId ||
                              editFormData.acaoId === "manual"
                            }
                            value={editFormData.acao}
                            onChange={(e) =>
                              setEditFormData({
                                ...editFormData,
                                acao: e.target.value,
                              })
                            }
                            className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm"
                            placeholder="Ex: Facebook, Panfletagem, etc."
                          />
                        </div>
                      )}
                    </div>
                  ) : (
                    <input
                      value={editFormData.acao}
                      onChange={(e) =>
                        setEditFormData({
                          ...editFormData,
                          acao: e.target.value,
                        })
                      }
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  )}
                </div>
              </div>
              <button
                type="submit"
                className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-xl font-bold transition flex items-center justify-center space-x-2"
              >
                <span>Salvar AlteraÃ§Ãµes</span>
              </button>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function BasesView({
  bases,
  onToast,
  whatsappMessages,
  botConfig,
  onSendBot,
  onMassSendBot,
  gap,
  basesRenovacao,
  profile,
  mensagensEnviadasLog = [],
}: {
  bases: BaseEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
  whatsappMessages: WhatsAppMessage[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
  gap: GapEntry[];
  basesRenovacao: BaseEntry[];
  profile: UserProfile;
  mensagensEnviadasLog?: MensagemEnviadaLog[];
}) {
  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || "NÃ£o informado",
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };
  
  const [formData, setFormData] = useState({
    nomeBase: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    curso: "",
    produto: "GraduaÃ§Ã£o" as "GraduaÃ§Ã£o" | "TÃ©cnico" | "PÃ³s-graduaÃ§Ã£o",
    numeroOportunidade: "",
    semestre: "",
    periodo: "",
    metodologia: "",
    formaIngresso: "",
    numeroMatricula: "",
  });
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [baseFilter, setBaseFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState("");
  const [produtoFilter, setProdutoFilter] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [semestreFilter, setSemestreFilter] = useState("");
  const [selectorOpen, setSelectorOpen] = useState(false);
  const [selectedEntry, setSelectedEntry] = useState<BaseEntry | null>(null);
  const [massSelectorOpen, setMassSelectorOpen] = useState(false);
  const [isAddMsgModalOpen, setIsAddMsgModalOpen] = useState(false);
  const [newMsgData, setNewMsgData] = useState({ modelName: "", texto: "" });
  const [invalidBaseIds, setInvalidBaseIds] = useState<Set<string>>(new Set());
  const [blockedFilter, setBlockedFilter] = useState<
    "all" | "blocked" | "unblocked"
  >("all");

  // New States for Sub-tabs and Candidates Editing
  const [basesSubTab, setBasesSubTab] = useState<
    "dashboard" | "lista" | "novo"
  >("dashboard");
  const [editingCandidate, setEditingCandidate] = useState<BaseEntry | null>(
    null,
  );
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editFormData, setEditFormData] = useState({
    nomeBase: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    curso: "",
    produto: "GraduaÃ§Ã£o" as "GraduaÃ§Ã£o" | "TÃ©cnico" | "PÃ³s-graduaÃ§Ã£o",
    numeroOportunidade: "",
    semestre: "",
    periodo: "",
    metodologia: "",
    formaIngresso: "",
    numeroMatricula: "",
    status: "Pendente" as
      | "Pendente"
      | "Interessado"
      | "Convertido"
      | "NÃ£o tem interesse"
      | "Sem retorno"
      | "Contato via Sales",
  });

  // Memoized aggregations for Dashboard basic metrics
  const statsByBase = useMemo(() => {
    const groups: {
      [key: string]: {
        total: number;
        converted: number;
        interested: number;
        pending: number;
      };
    } = {};
    bases.forEach((b) => {
      const baseName = b.nomeBase || "Sem Nome";
      if (!groups[baseName]) {
        groups[baseName] = {
          total: 0,
          converted: 0,
          interested: 0,
          pending: 0,
        };
      }
      groups[baseName].total += 1;
      if (b.status === "Convertido") groups[baseName].converted += 1;
      if (b.status === "Interessado") groups[baseName].interested += 1;
      if (b.status === "Pendente") groups[baseName].pending += 1;
    });

    return Object.entries(groups)
      .map(([name, data]) => ({
        name,
        total: data.total,
        converted: data.converted,
        interested: data.interested,
        pending: data.pending,
        conversionRate:
          data.total > 0
            ? ((data.converted / data.total) * 100).toFixed(1)
            : "0",
      }))
      .sort((a, b) => b.total - a.total);
  }, [bases]);

  const statsByProduct = useMemo(() => {
    const groups: { [key: string]: number } = {
      GraduaÃ§Ã£o: 0,
      TÃ©cnico: 0,
      "PÃ³s-graduaÃ§Ã£o": 0,
    };
    bases.forEach((b) => {
      const p = b.produto || "GraduaÃ§Ã£o";
      if (groups[p] !== undefined) {
        groups[p] += 1;
      } else {
        groups[p] = 1;
      }
    });
    return Object.entries(groups).map(([name, count]) => ({
      name,
      count,
      percentage:
        bases.length > 0 ? ((count / bases.length) * 100).toFixed(1) : "0",
    }));
  }, [bases]);

  const statsByStatus = useMemo(() => {
    const groups: { [key: string]: number } = {
      Pendente: 0,
      Interessado: 0,
      Convertido: 0,
      "NÃ£o tem interesse": 0,
      "Sem retorno": 0,
      "Contato via Sales": 0,
    };
    bases.forEach((b) => {
      const s = b.status || "Pendente";
      if (groups[s] !== undefined) {
        groups[s] += 1;
      }
    });
    return Object.entries(groups).map(([name, count]) => ({
      name,
      count,
      percentage:
        bases.length > 0 ? ((count / bases.length) * 100).toFixed(1) : "0",
    }));
  }, [bases]);

  const handleEditSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingCandidate) return;

    setLoading(true);
    try {
      const cleanCpf = editFormData.cpf
        ? editFormData.cpf.replace(/\D/g, "")
        : "";
      const cleanTelefone = editFormData.telefone.replace(/\D/g, "");

      const updatedData = {
        ...editFormData,
        cpf: cleanCpf,
        telefone: cleanTelefone,
        updatedAt: serverTimestamp(),
      };

      await updateDoc(
        doc(db, COLLECTIONS.BASES, editingCandidate.id),
        updatedData,
      );

      // If conversion status toggled to Convertido, check and sync with GAP
      if (
        editFormData.status === "Convertido" &&
        editingCandidate.status !== "Convertido" &&
        !invalidBaseIds.has(editingCandidate.id)
      ) {
        const q = query(
          collection(db, COLLECTIONS.GAP),
          where("cpf", "==", cleanCpf || ""),
        );
        const snap = await getDocs(q);
        if (snap.empty && cleanCpf) {
          await addDoc(collection(db, COLLECTIONS.GAP), {
            nome: editFormData.nome,
            telefone: cleanTelefone,
            cpf: cleanCpf,
            produto: editFormData.produto,
            numeroOportunidade: editFormData.numeroOportunidade,
            curso: editFormData.curso,
            metodologia: editFormData.metodologia,
            formaIngresso: editFormData.formaIngresso,
            semestre: editFormData.semestre,
            matAcad: false,
            documentos: {},
            createdAt: serverTimestamp(),
          });
          onToast(
            "Candidato atualizado e enviado para o GAP (Convertido)!",
            "success",
          );
        } else {
          onToast("Status atualizado com sucesso!", "success");
        }
      } else {
        onToast("InformaÃ§Ãµes do candidato atualizadas com sucesso!", "success");
      }

      setIsEditModalOpen(false);
      setEditingCandidate(null);
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleVerificacao = () => {
    const invalidIds = new Set<string>();
    bases.forEach((base) => {
      let match = false;

      if (
        gap.some(
          (g) =>
            (g.cpf &&
              base.cpf &&
              g.cpf.replace(/\D/g, "") === base.cpf.replace(/\D/g, "")) ||
            (g.telefone &&
              base.telefone &&
              g.telefone.replace(/\D/g, "") ===
                base.telefone.replace(/\D/g, "")) ||
            g.nome.toLowerCase().trim() === base.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (
        !match &&
        basesRenovacao.some(
          (b) =>
            (b.cpf &&
              base.cpf &&
              b.cpf.replace(/\D/g, "") === base.cpf.replace(/\D/g, "")) ||
            (b.telefone &&
              base.telefone &&
              b.telefone.replace(/\D/g, "") ===
                base.telefone.replace(/\D/g, "")) ||
            b.nome.toLowerCase().trim() === base.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (match) {
        invalidIds.add(base.id);
      }
    });
    setInvalidBaseIds(invalidIds);
    onToast(
      `VerificaÃ§Ã£o concluÃ­da: ${invalidIds.size} contatos jÃ¡ estÃ£o cadastrados em GAP/Base LÃ­quida.`,
      "success",
    );
  };

  const filteredBases = bases.filter((b) => {
    // Gestor Unidade filtering
    if (profile.role === "Gestor Unidade") {
      if (!profile.unidade || b.unidade !== profile.unidade) {
        return false;
      }
    }

    const matchesSearch = b.nome
      .toLowerCase()
      .includes(searchTerm.toLowerCase());
    const matchesBase =
      baseFilter.length === 0 || baseFilter.includes(b.nomeBase);
    const matchesStatus = !statusFilter || b.status === statusFilter;
    const matchesProduto = !produtoFilter || b.produto === produtoFilter;
    const matchesCurso =
      !cursoFilter || b.curso.toLowerCase().includes(cursoFilter.toLowerCase());
    const matchesSemestre =
      !semestreFilter ||
      (b.semestre &&
        b.semestre.toLowerCase().includes(semestreFilter.toLowerCase()));

    const isBlocked = invalidBaseIds.has(b.id);
    const matchesBlocked =
      blockedFilter === "all" ||
      (blockedFilter === "blocked" && isBlocked) ||
      (blockedFilter === "unblocked" && !isBlocked);

    return (
      matchesSearch &&
      matchesBase &&
      matchesStatus &&
      matchesProduto &&
      matchesCurso &&
      matchesSemestre &&
      matchesBlocked
    );
  });
  const uniqueBases = Array.from(new Set(bases.map((b) => b.nomeBase))).sort();
  const uniqueProdutos = ["GraduaÃ§Ã£o", "TÃ©cnico", "PÃ³s-graduaÃ§Ã£o"];
  const uniqueCursos = Array.from(new Set(bases.map((b) => b.curso))).sort();
  const uniqueSemestres = Array.from(
    new Set(bases.map((b) => b.semestre).filter(Boolean)),
  ).sort();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    const cleanCpf = formData.cpf ? formData.cpf.replace(/\D/g, "") : "";
    const cleanTelefone = formData.telefone.replace(/\D/g, "");

    const isDuplicate = bases.some(
      (b) =>
        (cleanCpf && b.cpf === cleanCpf) ||
        (!cleanCpf && cleanTelefone && b.telefone === cleanTelefone),
    );

    if (isDuplicate) {
      onToast("Registro jÃ¡ existe na base (verificado CPF/Telefone).", "error");
      return;
    }

    setLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.BASES), {
        ...formData,
        status: "Pendente",
        unidade: profile.unidade || "",
        promotorId: profile.uid,
        linkadoA: profile.linkadoA || null,
        createdAt: serverTimestamp(),
      });
      onToast("Registro salvo na base!");
      setFormData({
        nomeBase: "",
        nome: "",
        telefone: "",
        email: "",
        cpf: "",
        curso: "",
        produto: "GraduaÃ§Ã£o",
        numeroOportunidade: "",
        semestre: "",
        periodo: "",
        metodologia: "",
        formaIngresso: "",
        numeroMatricula: "",
      });
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleAddCustomMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMsgData.texto.trim()) return;
    setLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
        tipo: "bases",
        texto: newMsgData.texto,
        nome: newMsgData.modelName || undefined,
        createdAt: serverTimestamp(),
      });
      onToast("Mensagem de base salva!");
      setNewMsgData({ modelName: "", texto: "" });
      setIsAddMsgModalOpen(false);
    } catch (err: any) {
      console.error("Erro ao salvar mensagem:", err);
      onToast(`Erro ao salvar mensagem: ${err.message}`, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleInsertDefaultBasesMessages = async () => {
    try {
      const existing = whatsappMessages.filter((m) => m.tipo === "bases");
      if (existing.length > 0) {
        if (
          !window.confirm(
            "JÃ¡ existem mensagens para Bases. Deseja adicionar as mensagens padrÃµes mesmo assim?",
          )
        ) {
          return;
        }
      }

      const defaults = [
        "OlÃ¡ [nome], vi que vocÃª tem interesse no curso de [curso]. Vamos tirar suas dÃºvidas?",
        "Oi [nome], aqui Ã© da faculdade! Recebemos sua solicitaÃ§Ã£o sobre o curso de [curso]. Qual o melhor horÃ¡rio para conversarmos?",
        "Tudo bem, [nome]? Preparamos uma oferta especial para vocÃª comeÃ§ar o curso de [curso] ainda este semestre! Vamos lÃ¡?",
      ];

      for (const texto of defaults) {
        await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
          tipo: "bases",
          texto,
          createdAt: serverTimestamp(),
        });
      }
      onToast("Mensagens padrÃµes de base inseridas!");
    } catch (err: any) {
      onToast("Erro ao inserir mensagens padrÃµes.", "error");
    }
  };

  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);

  const handleBulkDelete = async () => {
    if (selectedEntries.length === 0) return;
    if (
      window.confirm(
        `Deseja excluir ${selectedEntries.length} registros selecionados?`,
      )
    ) {
      try {
        for (const id of selectedEntries) {
          await deleteDoc(doc(db, COLLECTIONS.BASES, id));
        }
        onToast(`${selectedEntries.length} registros removidos.`);
        setSelectedEntries([]);
      } catch (err: any) {
        onToast("Erro ao excluir registros.", "error");
      }
    }
  };

const handleSendViaWhats = async (entry: BaseEntry) => {
    try {
      await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
        leadId: entry.id,
        nome: entry.nome || "Sem Nome",
        telefone: entry.telefone || "",
        curso: entry.curso || "NÃ£o informado",
        base: entry.nomeBase || "Bases",
        tipoEnvio: "whats",
        dataHora: serverTimestamp(),
        usuarioId: profile.uid || "",
        usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
        unidade: entry.unidade || profile.unidade || "",
      });
      onToast(`Envio via WhatsApp para ${entry.nome} registrado no relatÃ³rio!`, "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar envio via WhatsApp.", "error");
    }
  };

  const handleSendViaMalaDireta = async (entry: BaseEntry) => {
    try {
      await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
        leadId: entry.id,
        nome: entry.nome || "Sem Nome",
        telefone: entry.telefone || "",
        curso: entry.curso || "NÃ£o informado",
        base: entry.nomeBase || "Bases",
        tipoEnvio: "maladireta",
        dataHora: serverTimestamp(),
        usuarioId: profile.uid || "",
        usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
        unidade: entry.unidade || profile.unidade || "",
      });
      onToast(`Envio via Mala Direta para ${entry.nome} registrado no relatÃ³rio!`, "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar envio via Mala Direta.", "error");
    }
  };

  const handleBulkSendViaWhats = async () => {
    if (selectedEntries.length === 0) return;
    const selectedBases = filteredBases.filter((b) => selectedEntries.includes(b.id));
    try {
      for (const entry of selectedBases) {
        await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
          leadId: entry.id,
          nome: entry.nome || "Sem Nome",
          telefone: entry.telefone || "",
          curso: entry.curso || "NÃ£o informado",
          base: entry.nomeBase || "Bases",
          tipoEnvio: "whats",
          dataHora: serverTimestamp(),
          usuarioId: profile.uid || "",
          usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
          unidade: entry.unidade || profile.unidade || "",
        });
      }
      onToast(`${selectedBases.length} envios via WhatsApp registrados com sucesso!`, "success");
    } catch (e) {
      console.error(e);
      onToast("Erro ao registrar envios em massa via WhatsApp.", "error");
    }
  };

  const handleBulkSendViaMalaDireta = async () => {
    if (selectedEntries.length === 0) return;
    const selectedBases = filteredBases.filter((b) => selectedEntries.includes(b.id));
    try {
      for (const entry of selectedBases) {
        await addDoc(collection(db, COLLECTIONS.MENSAGENS_ENVIADAS_LOG), {
          leadId: entry.id,
          nome: entry.nome || "Sem Nome",
          telefone: entry.telefone || "",
          curso: entry.curso || "NÃ£o informado",
          base: entry.nomeBase || "Bases",
          tipoEnvio: "maladireta",
          dataHora: serverTimestamp(),
          usuarioId: profile.uid || "",
          usuarioNome: profile.nome || profile.name || "UsuÃ¡rio",
          unidade: entry.unidade || profile.unidade || "",
        });
      }
      onToast(`${selectedBases.length} envios via Mala Direta registrados com sucesso!`, "success");
    } catch (e) {
      console.error(e);
      onToast("Erro ao registrar envios em massa via Mala Direta.", "error");
    }
  };

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries([...selectedEntries, id]);
    } else {
      setSelectedEntries(selectedEntries.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(
        filteredBases.filter((b) => !invalidBaseIds.has(b.id)).map((b) => b.id),
      );
    } else {
      setSelectedEntries([]);
    }
  };

  const handleStatusChange = async (entry: BaseEntry, status: string) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), { status });

      if (status === "Convertido" && !invalidBaseIds.has(entry.id)) {
        // Logic for transferring to GAP
        const q = query(
          collection(db, COLLECTIONS.GAP),
          where("cpf", "==", entry.cpf || ""),
        );
        const snap = await getDocs(q);
        if (snap.empty && entry.cpf) {
          await addDoc(collection(db, COLLECTIONS.GAP), {
            nome: entry.nome,
            telefone: entry.telefone,
            cpf: entry.cpf,
            produto: entry.produto,
            numeroOportunidade: entry.numeroOportunidade,
            curso: entry.curso,
            metodologia: entry.metodologia,
            formaIngresso: entry.formaIngresso,
            semestre: entry.semestre,
            matAcad: false,
            documentos: {},
            createdAt: serverTimestamp(),
          });
          onToast("Candidato convertido e enviado para GAP!");
        } else {
          onToast("Status atualizado!");
        }
      } else {
        onToast("Status da base atualizado!");
      }
    } catch (err: any) {
      onToast(err.message, "error");
    }
  };

  const handleDeleteBase = async (id: string) => {
    if (window.confirm("Deseja excluir este registro da base?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.BASES, id));
        onToast("Registro removido.");
      } catch (err: any) {
        onToast("Erro ao excluir registro.", "error");
      }
    }
  };

  const handleExport = () => {
    const data = filteredBases.map((b) => ({
      Nome: b.nome,
      Telefone: b.telefone,
      Email: b.email || "",
      CPF: b.cpf || "",
      Curso: b.curso,
      Produto: b.produto || "GraduaÃ§Ã£o",
      "NÂº Oportunidade": b.numeroOportunidade || "",
      Semestre: b.semestre || "",
      Periodo: b.periodo || "",
      Metodologia: b.metodologia || "",
      "Forma de Ingresso": b.formaIngresso || "",
      "NÂº MatrÃ­cula": b.numeroMatricula || "",
      Base: b.nomeBase,
      Status: b.status,
      Data: b.createdAt?.seconds
        ? new Date(b.createdAt.seconds * 1000).toLocaleDateString()
        : "",
    }));
    exportToExcel(data, "Base_Candidatos");
  };

  const handleExportMalaDireta = () => {
    const data = filteredBases.map((b) => ({
      Nome: b.nome,
      Email: b.email || "",
    }));
    exportToExcel(data, "Mala_Direta_Bases");
  };

  const handleExportSMS = () => {
    const data = filteredBases.map((b) => {
      let tel = b.telefone.replace(/\D/g, "");
      if (tel.length > 0 && !tel.startsWith("55")) {
        tel = "55" + tel;
      }
      return { Telefone: tel };
    });
    exportToCSV(data, "SMS_Bases");
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const normalizeProduto = (val: string) => {
          if (!val) return "GraduaÃ§Ã£o";
          const lower = val.trim().toLowerCase();
          if (lower.includes("gradua")) return "GraduaÃ§Ã£o";
          if (lower.includes("tecnic") || lower.includes("tÃ©cnic"))
            return "TÃ©cnico";
          if (lower.includes("pos") || lower.includes("pÃ³s"))
            return "PÃ³s-graduaÃ§Ã£o";
          return val;
        };

        const normalizeMetodologia = (val: string) => {
          if (!val) return "";
          const lower = val.trim().toLowerCase();
          if (lower === "ead") return "EAD";
          if (lower === "presencial") return "Presencial";
          if (lower === "semipresencial") return "Semipresencial";
          if (lower === "flex") return "Flex";
          if (lower === "hibrido" || lower === "hÃ­brido") return "HÃ­brido";
          if (lower === "digital") return "Digital";
          return val;
        };

        const normalizeStatusBase = (val: string) => {
          if (!val) return "Pendente";
          const lower = val.trim().toLowerCase();
          if (lower === "pendente") return "Pendente";
          if (lower === "matriculado") return "Matriculado";
          if (
            lower === "ligacao efetuada" ||
            lower === "ligaÃ§Ã£o efetuada" ||
            lower.includes("liga")
          )
            return "LigaÃ§Ã£o Efetuada";
          if (lower === "sem interesse" || lower.includes("sem inter"))
            return "Sem Interesse";
          return val.charAt(0).toUpperCase() + val.slice(1);
        };

        const batch = data.map((item) => ({
          nome: String(getVal(item, "Nome", "nome") || "").trim(),
          telefone: String(getVal(item, "Telefone", "telefone") || "").replace(
            /\D/g,
            "",
          ),
          cpf: String(getVal(item, "CPF", "cpf") || "").replace(/\D/g, ""),
          curso: String(getVal(item, "Curso", "curso") || "").trim(),
          produto: normalizeProduto(
            String(getVal(item, "Produto", "produto") || ""),
          ),
          numeroOportunidade: String(
            getVal(
              item,
              "NÂº Oportunidade",
              "numeroOportunidade",
              "oportunidade",
            ) || "",
          ).trim(),
          semestre: String(getVal(item, "Semestre", "semestre") || "").trim(),
          periodo: String(
            getVal(item, "Periodo", "periodo", "perÃ­odo") || "",
          ).trim(),
          metodologia: normalizeMetodologia(
            String(getVal(item, "Metodologia", "metodologia") || ""),
          ),
          formaIngresso: String(
            getVal(item, "Forma de Ingresso", "formaIngresso", "ingresso") ||
              "",
          ).trim(),
          numeroMatricula: String(
            getVal(
              item,
              "NÂº MatrÃ­cula",
              "numeroMatricula",
              "matricula",
              "matrÃ­cula",
            ) || "",
          ).trim(),
          nomeBase: String(
            getVal(item, "Base", "nomeBase") || "Importado",
          ).trim(),
          status: normalizeStatusBase(
            String(getVal(item, "Status", "status") || ""),
          ),
          createdAt: serverTimestamp(),
        }));

        let imported = 0;
        let skipped = 0;
        const insertedCpfs = new Set();
        const insertedTels = new Set();

        for (const entry of batch) {
          const isDupCpf =
            entry.cpf &&
            (bases.some((b) => b.cpf === entry.cpf) ||
              insertedCpfs.has(entry.cpf));
          const isDupTel =
            entry.telefone &&
            (bases.some((b) => b.telefone === entry.telefone) ||
              insertedTels.has(entry.telefone));

          if (!isDupCpf && !isDupTel) {
            await addDoc(collection(db, COLLECTIONS.BASES), entry);
            if (entry.cpf) insertedCpfs.add(entry.cpf);
            if (entry.telefone) insertedTels.add(entry.telefone);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} registros importados com sucesso! ${skipped > 0 ? `${skipped} ignorados por duplicidade.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar dados.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      {/* Header and Actions */}
      <div className="flex flex-col md:flex-row justify-between items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 flex items-center gap-2">
            <Database className="text-blue-600" size={28} />
            Acompanhamento de Bases
          </h2>
          <p className="text-sm text-slate-500">
            Gerencie e analise as bases de captaÃ§Ã£o de candidatos da sua
            unidade.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {[ROLES.ADMIN_MASTER, ROLES.LIDER_FDV].includes(profile.role) && (
            <button
              onClick={handleVerificacao}
              className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold shadow-sm"
              title="Verificar se contatos existem no GAP ou Base LÃ­quida"
            >
              <Search size={18} />
              <span>VerificaÃ§Ã£o</span>
            </button>
          )}
          <button
            onClick={() => setIsAddMsgModalOpen(true)}
            className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
          >
            <Plus size={18} />
            <span>Inserir Mensagens</span>
          </button>
          <button
            onClick={handleInsertDefaultBasesMessages}
            className="bg-slate-50 text-slate-400 px-3 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-100 transition-all text-[10px] font-bold shadow-sm"
            title="Inserir Mensagens PadrÃµes"
          >
            <MessageSquare size={14} />
          </button>
          <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer shadow-sm">
            <Upload size={18} />
            <span>Importar</span>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleImport}
              className="hidden"
            />
          </label>
          <button
            onClick={handleExport}
            className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold shadow-sm"
          >
            <Download size={18} />
            <span>Exportar Excel</span>
          </button>
          <button
            onClick={handleExportMalaDireta}
            className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
          >
            <Mail size={18} />
            <span>Mala Direta</span>
          </button>
          <button
            onClick={handleExportSMS}
            className="bg-orange-50 text-orange-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-orange-100 transition-all text-sm font-bold shadow-sm"
          >
            <MessageSquare size={18} />
            <span>SMS (CSV)</span>
          </button>
        </div>
      </div>

      {/* Elegant Sub-tabs */}
      <div className="flex border-b border-slate-100 gap-2 overflow-x-auto">
        <button
          onClick={() => setBasesSubTab("dashboard")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "dashboard"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <LayoutDashboard size={16} />
          <span>Painel Geral (Dashboard)</span>
        </button>
        <button
          onClick={() => setBasesSubTab("lista")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "lista"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <Database size={16} />
          <span>Lista de Candidatos</span>
        </button>
        <button
          onClick={() => setBasesSubTab("novo")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "novo"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <UserPlus size={16} />
          <span>Novo Registro</span>
        </button>
      </div>

      {/* Dashboard Sub-tab */}
      {basesSubTab === "dashboard" && (
        <div className="space-y-6" id="bases-dashboard-view">
          {/* Main Hero KPI Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-blue-50 text-blue-600 rounded-xl">
                <Users size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-slate-400 uppercase tracking-wider block">
                  Total de Cadastros
                </span>
                <span className="text-2xl font-black text-slate-800">
                  {bases.length}
                </span>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-emerald-50 text-emerald-600 rounded-xl">
                <CheckCircle2 size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-emerald-500 uppercase tracking-wider block">
                  Convertidos
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Convertido").length}
                  </span>
                  <span className="text-xs font-bold text-emerald-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Convertido")
                            .length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-blue-50 text-blue-500 rounded-xl">
                <TrendingUp size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-blue-500 uppercase tracking-wider block">
                  Interessados
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Interessado").length}
                  </span>
                  <span className="text-xs font-bold text-blue-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Interessado")
                            .length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-amber-50 text-amber-600 rounded-xl">
                <Clock size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-amber-500 uppercase tracking-wider block">
                  Pendentes
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Pendente").length}
                  </span>
                  <span className="text-xs font-bold text-amber-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Pendente").length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* 2-Column Bento Breakdown */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Column 1: Performance por Base */}
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <Database size={18} className="text-blue-500" />
                Desempenho por Base de Origem
              </h3>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase pb-2">
                      <th className="pb-2">Nome da Base</th>
                      <th className="pb-2 text-center">Registros</th>
                      <th className="pb-2 text-center">ConversÃµes</th>
                      <th className="pb-2 text-right">ConversÃ£o (%)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {statsByBase.length === 0 ? (
                      <tr>
                        <td
                          colSpan={4}
                          className="py-4 text-center text-slate-400 italic"
                        >
                          Nenhuma base registrada ainda.
                        </td>
                      </tr>
                    ) : (
                      statsByBase.slice(0, 10).map((b) => (
                        <tr key={b.name} className="hover:bg-slate-50/50">
                          <td className="py-3 font-semibold text-slate-700">
                            {b.name}
                          </td>
                          <td className="py-3 text-center font-bold text-slate-600">
                            {b.total}
                          </td>
                          <td className="py-3 text-center text-emerald-600 font-bold">
                            {b.converted}
                          </td>
                          <td className="py-3 text-right">
                            <span className="inline-block px-2 py-0.5 rounded-full font-black bg-emerald-50 text-emerald-700 text-[10px]">
                              {b.conversionRate}%
                            </span>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Column 2: Status & Product distributions */}
            <div className="space-y-6">
              {/* Distribution of Statuses */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <Target size={18} className="text-blue-500" />
                  DistribuiÃ§Ã£o de Status dos Candidatos
                </h3>
                <div className="space-y-3">
                  {statsByStatus.map((s) => (
                    <div key={s.name} className="space-y-1">
                      <div className="flex justify-between text-xs font-semibold">
                        <span className="text-slate-600 flex items-center gap-1.5">
                          <span
                            className={cn(
                              "w-2 h-2 rounded-full",
                              s.name === "Pendente" && "bg-slate-400",
                              s.name === "Interessado" && "bg-blue-400",
                              s.name === "Convertido" && "bg-emerald-400",
                              s.name === "NÃ£o tem interesse" && "bg-rose-400",
                              s.name === "Sem retorno" && "bg-orange-400",
                              s.name === "Contato via Sales" && "bg-purple-400",
                            )}
                          />
                          {s.name}
                        </span>
                        <span className="text-slate-800 font-bold">
                          {s.count}{" "}
                          <span className="text-slate-400 font-normal">
                            ({s.percentage}%)
                          </span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                        <div
                          className={cn(
                            "h-full rounded-full transition-all",
                            s.name === "Pendente" && "bg-slate-400",
                            s.name === "Interessado" && "bg-blue-400",
                            s.name === "Convertido" && "bg-emerald-400",
                            s.name === "NÃ£o tem interesse" && "bg-rose-400",
                            s.name === "Sem retorno" && "bg-orange-400",
                              s.name === "Contato via Sales" && "bg-purple-400",
                          )}
                          style={{ width: `${s.percentage}%` }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Distribution of Products */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <GraduationCap size={18} className="text-blue-500" />
                  DistribuiÃ§Ã£o por Produto AcadÃªmico
                </h3>
                <div className="space-y-3">
                  {statsByProduct.map((p) => (
                    <div key={p.name} className="space-y-1">
                      <div className="flex justify-between text-xs font-semibold">
                        <span className="text-slate-600">{p.name}</span>
                        <span className="text-slate-800 font-bold">
                          {p.count}{" "}
                          <span className="text-slate-400 font-normal">
                            ({p.percentage}%)
                          </span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-blue-500 rounded-full transition-all"
                          style={{ width: `${p.percentage}%` }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* New Cadastro Sub-tab */}
      {basesSubTab === "novo" && (
        <div className="max-w-xl mx-auto">
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
            <h3 className="text-xl font-bold text-slate-900 mb-4">
              Novo Registro em Base
            </h3>
            <form onSubmit={handleSubmit} className="space-y-4">
              <input
                placeholder="Nome da Base (Ex: Junho 2024)"
                required
                value={formData.nomeBase}
                onChange={(e) =>
                  setFormData({ ...formData, nomeBase: e.target.value })
                }
                className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              />
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Nome"
                  required
                  value={formData.nome}
                  onChange={(e) =>
                    setFormData({ ...formData, nome: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="Telefone"
                  required
                  value={formData.telefone}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      telefone: formatPhone(e.target.value),
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Email (Opcional)"
                  type="email"
                  value={formData.email}
                  onChange={(e) =>
                    setFormData({ ...formData, email: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="CPF"
                  value={formData.cpf}
                  onChange={(e) =>
                    setFormData({ ...formData, cpf: formatCPF(e.target.value) })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="NÂ° Oportunidade"
                  required
                  value={formData.numeroOportunidade}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      numeroOportunidade: e.target.value,
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Semestre"
                  required
                  value={formData.semestre}
                  onChange={(e) =>
                    setFormData({ ...formData, semestre: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <select
                  value={formData.produto}
                  onChange={(e) =>
                    setFormData({ ...formData, produto: e.target.value as any })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                >
                  {uniqueProdutos.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Metodologia"
                  required
                  value={formData.metodologia}
                  onChange={(e) =>
                    setFormData({ ...formData, metodologia: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="Forma de Ingresso"
                  required
                  value={formData.formaIngresso}
                  onChange={(e) =>
                    setFormData({ ...formData, formaIngresso: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="PerÃ­odo"
                  value={formData.periodo}
                  onChange={(e) =>
                    setFormData({ ...formData, periodo: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="NÂº MatrÃ­cula"
                  value={formData.numeroMatricula}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      numeroMatricula: e.target.value,
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <input
                placeholder="Curso"
                required
                value={formData.curso}
                onChange={(e) =>
                  setFormData({ ...formData, curso: e.target.value })
                }
                className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              />
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-amber-600 hover:bg-amber-700 text-white font-bold py-3 rounded-xl transition-all disabled:opacity-50"
              >
                {loading ? "Salvando..." : "Adicionar Ã  Base"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* Candidates List Sub-tab */}
      {basesSubTab === "lista" && (
        <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
          <div className="p-6 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
            <h3 className="text-xl font-bold text-slate-900">
              Bases a Trabalhar
            </h3>
            <div className="flex flex-wrap gap-2">
              <div className="relative">
                <Search
                  className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                  size={16}
                />
                <input
                  type="text"
                  placeholder="Buscar por nome..."
                  className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 w-48"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>
              <MultiSelect
                options={uniqueBases}
                selectedValues={baseFilter}
                onChange={setBaseFilter}
                placeholder="Todas as Bases"
                allLabel="Todas as Bases"
              />
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={produtoFilter}
                onChange={(e) => setProdutoFilter(e.target.value)}
              >
                <option value="">Todos os Produtos</option>
                {uniqueProdutos.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={cursoFilter}
                onChange={(e) => setCursoFilter(e.target.value)}
              >
                <option value="">Todos os Cursos</option>
                {uniqueCursos.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={semestreFilter}
                onChange={(e) => setSemestreFilter(e.target.value)}
              >
                <option value="">Todos os Semestres</option>
                {uniqueSemestres.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
              >
                <option value="">Todos Status</option>
                <option value="Pendente">Pendente</option>
                <option value="Interessado">Interessado</option>
                <option value="Convertido">Convertido</option>
                <option value="NÃ£o tem interesse">NÃ£o tem interesse</option>
                <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={blockedFilter}
                onChange={(e) => setBlockedFilter(e.target.value as any)}
              >
                <option value="all">VerificaÃ§Ã£o: Todos</option>
                <option value="blocked">VerificaÃ§Ã£o: Bloqueados</option>
                <option value="unblocked">VerificaÃ§Ã£o: Ativos</option>
              </select>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                  <th className="px-6 py-4 w-12 text-center">#</th>
                  <th className="px-6 py-4 w-12">
                    <input
                      type="checkbox"
                      checked={
                        filteredBases.filter((b) => !invalidBaseIds.has(b.id))
                          .length > 0 &&
                        selectedEntries.length ===
                          filteredBases.filter((b) => !invalidBaseIds.has(b.id))
                            .length
                      }
                      onChange={(e) => toggleSelectAll(e.target.checked)}
                    />
                  </th>
                  <th className="px-6 py-4">Nome</th>
                  <th className="px-6 py-4">Base</th>
                  <th className="px-6 py-4">Status</th>
                  <th className="px-6 py-4 flex flex-wrap items-center gap-2">
                    {selectedEntries.length > 0 && (
                      <button
                        onClick={handleBulkSendViaWhats}
                        className="text-emerald-600 font-bold hover:underline py-1 px-2 bg-emerald-50 rounded-lg flex items-center gap-1 text-xs"
                        title="Registrar envio via Whats para todos os selecionados"
                      >
                        <MessageSquare size={13} /> Whats ({selectedEntries.length})
                      </button>
                    )}
                    {selectedEntries.length > 0 && (
                      <button
                        onClick={handleBulkSendViaMalaDireta}
                        className="text-purple-600 font-bold hover:underline py-1 px-2 bg-purple-50 rounded-lg flex items-center gap-1 text-xs"
                        title="Registrar envio via Mala Direta para todos os selecionados"
                      >
                        <Mail size={13} /> Mala Direta ({selectedEntries.length})
                      </button>
                    )}
                    {selectedEntries.length > 0 && botConfig.url && (
                      <button
                        onClick={() => setMassSelectorOpen(true)}
                        className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1 text-xs"
                      >
                        <Bot size={13} /> Bot Em Massa
                      </button>
                    )}
                    {selectedEntries.length > 0 && (
                      <button
                        onClick={handleBulkDelete}
                        className="text-rose-600 font-bold hover:underline text-xs"
                      >
                        Excluir ({selectedEntries.length})
                      </button>
                    )}
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredBases.map((entry, index) => (
                  <tr
                    key={entry.id}
                    className={cn(
                      "hover:bg-slate-50/50 transition-all",
                      invalidBaseIds.has(entry.id) && "bg-rose-50/50",
                    )}
              xœì½ÛrÜH² ø®¯æVwgV‘É‹¤:Ý)R<EUó´$rEªzvÙ2	™ “(!l ÉKñÐl^gæaçym¶wÖìX³6³¶y9‡²_0Ÿ°îqC\$I©¨î‚U‰‰@\=<<Ü=<Ü	!äéâxÖËŒ’ (^“h£3=_ú–L/–‘2:/—FQZF99ÎÒri˜%!K-’ Œ–­¬°×ó¢ã®›Ë8£sòY½r·¾\†óõËÛÔzœNg¥ç#!åÅªD£Ãì¼ãÍÆE0L¢pº~$qø,(¢Ý°èŸE ‘_ôã°ç>´,^DI4*£pÊÄQÑÓQ2£V•déöIŽ£ËnÔ#¾!ã³P×MòË_Ö-³ñ8‰h?e™EõË GeŸ¥ç­Â7€eÏ”Þád‡ñ©Zâ8,Ã–FYâ-åŠij]˜ý›••š: «°Òlùçp}Ûš£+|-©yÜ®#£Y^d·è‰˜qM
±þ¡ü(Z:_Z#“ri•ú,¦dL—.–VÛt²<hG€“¤[“Ý™£Õ•éù;“ Éù«í‚£~pµ!5«·š—éñÝö±:Z2ÅÉ€%²ÒL†cže²äÙHn¸t<Kê{¶÷_Èåq–O‚~we·kGwWà)¢IT”ù`Ä0™5€ˆç˜@FO?T¦Qgav@™ÎòiÒ ™çF€áý„pY_êä/{ÙíâI.kjeiQ’!l/³qA6Hw¥E0†vÒÓ8ƒ>þgrô®×?Ž yõ ï&´Ñ¤ŸDA¸’"6P¬§›ÐÍç°Ú*ð¥÷ËìevåÛÐ¥n¯üÁG!+ñæéñÊUZjS×ê{?¦	îîòž/I§£6S“Ë¿çÃd=i„õP©ü=@Z ]@U®Œ§?£ýéœeÑÁÑYŸ†Yù>˜•¨x”uz ót\ž´ìÄ«9:1	’ Œó¨Zµ“®(_AÐ‹6ñ½G ªYž’T]E<[Ãj÷r<r#^¥»sÿqóâåÝ~Ê:]ß0m¼ŽÔ ò˜IˆÐòo¸T Ò~mRè¡ ?œä›ãÀñ@Ž)ã2³2HHêÂ|–-ù="ÍÖtJ¦A Ði‡A™5Â€ëUT 8øã,€m¨ˆŽ{uå
XVù@€êê¼FTŸ†UÌÍ«»›NàÕ©áIŸnf^ÁJ"ÏéRºÅäq¢Ï	V+¦äÕçš’†¨–^õºÞÚkê_PâŒâDe¿­	ÇO/œ3iØ(ag0ÏœÛ*ƒrVÜ‰tùBqi…¬ãq)—6=¿ŒK \ŽÒºuÚ,×ªä±’±P(+0›•IœFK)îãÃ,£œþî,ÖT®‚Œmcû4KµS/ýC¿4ÁÁÀœùÚÜEÒ€Ä4ÌZ5K™qÙ*}›¿Ñí,=ò2nÙ¦²Ié;Ôü-¿¾þoÔ1!1w;`çY¡Œš¾Íßö4„5ËÓvÃÎrDûªYþ~#pÃ[F	üADE«æ«íGÛ{ê›÷Í%F6-ã,å”¤ZOÅ¯õe–£u*N?U^æ®HÁÓ§Õï¹«q ]é³³×ÈXsvEÅÁ§ÊKU‘wžZIM½€Ýœn"w¶gÖè¹üg“¾µ†c[ÎÊ2Kkæv¶$}„­Q¨%¤ˆÊEÇ|ÁºQ+)…²|o¥Ý2‡Í¯®ÈUwdr
EµÎ&5ð^%',‹VÉ? ‰ð¶]‡ùn¶þÑ•O/ÍKQ^QÈ:Óe6™¾~¶¯	àÐ°ïã€vŽOs¾¤§”¡¨¼	rUœ{¼ÂçÅØ=o"Xñe$ŒoW¨~‰ÉìN	QãŠ_
’Ä?Õ\ÎxãÊä„‰ìšØf Œ rðýôúOy€$ýLþGXð*¸þk2’f@³€ÙÁ×¬çk¼fþ8ö°Ç†1•SÆú^fMHv\BYŠIhŸ¡ø.Ïñ©Ò1JtRØ›ø‡Ÿ ™T¡õ'À'EÚ”^ÿÈÄ7qÀ'º…KáJ“.Q­÷+üQüª–±ããÅíö¬€# þ|¼BœrØ$YZ_lNfþ¹Þ?©noVŽ³8·ÛŸØ¤s>é3Ns¼ÄNk.osÍNÔ‚á¥^dùäy Ô§žÉ84pàT§V€aeÕrmÊøÀÔ«·(M`‹‚ô¥M©ÑôX”ÁóÂ6%ð¸W–Á—6¥¦yÎJYŽ¿Ò’ßåA8®ÿ$†&Î`³Ïö¦Y^ÎR˜ñ°°õ¥M§ÄÁš¨E	¶;{’âçf-JN@Z	³$Ç(­$µ©™î¦c”«d´ÄVJaö*(a™%J™Ü
ŽTNèR;–“rp]ù«¦»[à’}•…AÒFj¨‘L*/ÏmÂMÓrO\KŽÁ<¿E6`¾MÇ€\ª¡ç§ÜXÅš ØßÖPìO²Û>bTRò×ÊÉ 5Õ7=ÒA-tP:¨iêM@]¹Ã<(Nî t>9Òs;½ç€Ï%;Œ¨_ðƒ¾êÏ%¿¯»êfBÏ(FYr ÛôÆå·¾rè$V×4³;ÃÖ°pÂ{0_GéÉõ3”É€@Œ€É‘se’Ð€	'Á¹Ý0+úwwsÐm˜…zNHD[;5Ñ8Ñ^¡RþërùkÂY¹œ#BéùzY´}ëd‹ž–‹>ãÖ!k|…$N.­qMØÌç³¢Œ/Ä«TªÿD˜GÀ~£ažM‘€åÈ¾N—\iÇIv¶t±„gÛšºÉìTxv-Êå¸v’Â	LÕÙü:£dœ/ÑO²î“8ÚO&K¿&AO°SqJŽaoÆ¿?fÙÿ†³< {Í:0B83¨þ6
PH«€ëBG d•gt@ƒÖ0‹zCñCBXj·õ“‡jä´Ú£Ç™ÎãJ—:O#åk+Wî	ÕÔ.Jen%¶Ÿ<´†â¡ø-nÇ|$…{vqåhàÈìØ§-)LRU²§DdªËRrQŽ•ªHb–[ò’cjþƒ2-ðÝÛ'ìEaÃÆ]ä× ü³á$.7.Ù‹°c)J Ú2)ÓuÌi.ƒq‡ÿÁ1h˜ª×5Š”®ÃR³(%–PÞóè¤×“`%¹H²ÑG—†D7%œ½f›¯QV
Bå¬¯4Ÿ¶êìß
›Y`c7Ü[XýqÃ¶ìü(ˆÉ±áôºå)q{q´ßï«ÍûjEfÕ™}%®Ü§Ïîa)3Í)?0yX«Ì6p_ÐÕjlâ×ä9ï0cƒvæÌ’uæœ¹{²¨ÁØ	àU”c}> ÿ4KO2²¶²öÈUÀAB]+–¦{Ì">¦g¿ïè~·¨n¡´P™hhûâç=E·C¡LºÇXVoÿÓÕJ1ÇÌÕ©µkXñüíPØ{ŠÁ;KT­ÙÝ›Ž€£\oƒÈTkêÆd²Òì÷S…øïcïÎ£É4ÉþqŒ£î²É„ÍÛû/>!v¡ïhz|ß—žF(‚~¦¯Ÿ—!ý÷#êÊOŠžö‘Î}ÃVçqÔß(Ý½§»MÏï1C[s9ö'$³ü÷g\ýœ¸ºÏÏ¼çÅÖšË.ŒãGë÷ç* ]¼
¤_&òq8âÂB7l/a§øã,â¸Pô'Á´ÛRÝ»ÏtW3Œ.6.§WbÒ§W5÷R§þ{?ufÎ®Ë:Óçû°¬„ÑÆ=Þêopÿt‹R1€ùy/ø¬{A”_ÿ%Í“;Ö½ÏÅa×^Ÿÿ	w
iYõ3N~Nœ|U¡ý¤h©ÃÝ7ÔÔM÷~FÏÏ‰ž8;Þô6Ž?)’jÖ–÷MMûÐŸõskÑÐ|öú/h?{thÒ–÷¾á©mü3¦~VÑ‰YjßÊaNíDw…ŸPì×º‰[^b¿“ìwpy½ÅÅu;m®*?ûôFŒ3ÑHEƒDjŠÈ,Ë$±Á €š>$ÓP—ã[ÙlF[wß€íjì³S[]›ËBu.UÏmÝxž-ce…%.¢IlRãoùu@ij¬®)4Ú®º×1Œi%Aî@ßÍ†¦é)¨Eªkz*µI„ Ñ»€³7—n¾U%/ç(pâÆè“Ð†˜Ç"=M‰ž²i0ŠK¼¡Ù¾b¤d“t`áž
e°ÑtÈ€¿çdoxà¶ÿ+Ý†ˆo"ky}yùù¯(¬‹»éâ¢8÷º ËÅ.á'˜¹c¨zŠ«.+"±êL·lqUÙÑ!sE¤º¾ØTî=ê™é±•[¹L¨g—¢Ý€ë²ZUxÂF_l\R„GåoBG8¡Þø¸Ë?¼˜Óé©Ð`£pLŠ±E‡ÐŸÖ­žA¦Îb@‘³~† æBLÍÛ<ÑËIë·E‚í,’Îûa¤;aªú¥¦aVÂ†r7.åO} iø,ã#XD…e0*qÎæŽ¬Í<°ðÆÈµ•çÁE?.è_ÐMïÑÊ;XBØ=£€ÒYœ_½Ä0½@°–ŸÎ½H&@¿Z.”WF^ÿbéH¦[¡ìŸcoŠ-‡haÛSaa
stYŠïdÃž[¥0dùS}¤¾¿{bU,&ú% loøCÁ}iN©Ý!Ÿ Ûó¡pãò C?ªXÔ³{#&f?¸À :cöå%‹$ÏÌŠ-(/wÈ/ªw~)ðÎŽp?&TY¦&Þ%Çû> ÜÕ*G©ïƒ<Æ=°èŠ¦IbYY1£ìÄ±Ò®´•¦Á,KÙ:`tÂ€–Ó¿dŒ\úÔvÞ)9ä†sA«{ir±q‰w…+‹]HÛÞ?vIÙO9lbRŽR¥Žœ$ Ž‡á«b,9Î+mô&UØ5³[dï	O³¾–µÊ³ €ÅÈ¨E‚Ó Np
åTn\isd¢ËÇèø”#œÌwÖ­l*åÃgq11ÛÊ4Í£Ó8:û%Èû
[&ÙìÇÀÈz¥¿^Š¶é6Ëæ(7 ïFÍ[á$N©ßæÀªÐ=4nŠåÝ[–¡i`<ÙÆËÎAÒ08w_€¸ïGLÒówu½I£3òå”õÌ<_`)¤ãn²go:½vð/‚YŒm
0»`0±—\ÿÉ þ»
	¥jk}Ù¹¨$ƒ
«ØÓãY:¢’1½þü&J³SìË÷ÐUäPôÇ¦8¶ãÏižÉ§ÄÇÜ1Mna¬':ìE¡B‹®D61  <Ä%´¼½éN¤ p]$å&Œ¿˜FÐ\‡ oåy–wèª>ÍâðIÕ»y[Dù>{yâèê€ìkWöz$~>Q=*qŽDŸ`×¬^ÞhS¤j½Ó@0à6&Ù¥Keÿ`¥ŸT»…H@b!k'WGï¶J+WŒCdœÓ3ìvAq‘ŽÐe5íð€*ÖH–Çãh¢uœ­$˜¹¦‚³ .I†Ï³”Op ê†ÃE²½÷òåÎöáîÞëƒþÁÖËƒ÷Û{¯·¶€}®%ou7ˆŸkD~f[œø¢ot|Äw{—åöt"}/Ê³j_¨B5PqJOjBÕï	‰RwÁš
· d€_ Æh³L¦]¹Þ«½—cq×Ö8‰Køâî=`5ž-Àz—Î«¹"£ .`;¨JÀ9Î’¨O×~·[ÞO$Èd{§E}l”/%ÞäÎNH<::–×Â`ÿ*Üw€I³‚ºÔ¸â·ºï(Ü”°éoÕŠ~EB¼R£sùÂ¦R¼J«5ÍK*†õ˜ÑÃë?Rô/û×-–Æ–c—Å°hª²Å‘C‘ ?‹Dã°¯£5¥œ­°U8ˆ¹&ƒBø%û­¸âÃx"
òÑÉa”OK&^µbŽZÉíÊ±Ó2Ïä«ZfQ€£wO_'Z¤šy¥ü’PÓ*Ÿ8¥à¾šRS’b€Rn»z¯)%fOí©–T[¶â}MeKÓ„(Âð¢É&_h – (ª;ŸVJO^›)ç.º¸óúY|ñ¢“ý­¯¸È,ÿkùª/2ÉÂ(Á½b7UÒ@±\V§9\1…H.72R²¼L¾ƒ)Ër"˜?–0“~GÕßîû9A&}ëE:±Äì"¿â–j(_ ¼‘AÕ¶p‘ÂGÐÙ+N/r¢)Rë¨`«H·0ž_ä!¥\\­h=#§ËZôFþ¯¢ZÖº_õI‘¿vÕÌ"7È‚ºÞ¬'ÐêGG-ÂÀªÑ–?«G¸Û`0¯>;jb×
ÄP’Àê¢	FŒ9T%w3L¥õªlL§$Ðÿ  TBUº™êë^¡Ñ%Þ'#ˆŽQ²5ÌD>…f²˜3AØª„w•¦Ò®i+™ÙN‹uÌ”MÇy6é¢xt•]¶´©†­kíz½~ûnW!4º-6Ty¤;£S6õEÇžþÎ¬‰ª}×(¢øû%€aTÈvw­bºe,¡gª íQv±jÌàØ™×‘ŠQšú&BžY¯S‚BõQF„´çÑq0KD}2ö¶·==†ÚG½nj¯®ð@H½Ÿ˜5É‹ëJu5Ñƒ´®ÄÅóu-EUn\ DW®*ÍÒ +»þË_âá€^ˆÔ^µ!ï‚šYï)--½7Ê*ÄwÎ¹‹®âF¡t´Ú$SýFxÐúáúO$:1 IÐñ.ñ1õž…wJ—e#6›-º`º9ã)˜?Õ³ÞM$°g[ ½Ùy½÷ýÖöÖž&‚õûýcËCØ]¸ÞPð‘*‚ä4“@‚Í6§Êºt*€(Â…bIb"UO±…º<tƒŽF2h‚.hàã6”æk¤6S‚ µšR>NI4Q5oIê<ÔHŽb‚àKŸk,!Ø­4H’
ñÔTuÂ\Zt³ÁqTXŒ0¤ÕI)|6K>2g‡9TˆŸy@nüzÚ"ÃÜ|4üp„:ž|R_|xÑ¬kØ±ãœ|eEoeÕ_IGz;|@'aVl~óÀø»
èê"¦sK0HC’›§%úA[ô!…®ûÿ¯_ðxôÑ“+K=}³ÿ¡ÍØòh’ÆèðƒRYý@æ9tÒ²M©¼râš¶ƒùÅ¡¢c!kdÈ·\iDD[½­!´°9ßÉ%Á2¨«À¯sØñ]AEŽj–êF·•$8À;”îòRãV”®4MN·Ùk;Î’ÂÇ Ä«h{å¦ÓBé8›¢qPkì—nQÕ5W³cqÎþ	’øGXÐ-qmÈ©J•×
F“kP¬ŽA¨h\5éÜ3d[íf§ç#@w@Rì}^P‹~çŽˆA{ZÀ}öãŽŒKEË€ç?Èµúð_r¯)g1ÔtÏ‡’·Z:g:wh»ìÞf\—-Dê¾`,†Ž·;¦ˆíVƒ{íÉy=N´÷2t»Ê~¥2$C¿CìŽu„v°ÆùuG·ÇWÆãñqÍø¾Jr”ƒätC¨$Xú•ÀÎˆ¢Áa¶s>Š’nH•W:ãïå1C(…v'…¤ØÅè¼Ö{øêå.^ØI ¸iùÔB4Ô Aqi¿ŒïÅfÿhå]Å…,`bÅœ°tÚð+Y·é9› 8ÍA_Í¬ehíû€nyvÆu`û]HtÛm:%~ù ðÞð õ´Ö£ý+l|G¾³™VIhuB•¼r Eš–JÝ´Gueeß n=µÆºˆÁU¶	B4z$ÞßÑ½Í¡¸Cr%‹–Çm¬$2²´ÂT=y`LAŠÈ[GT)Àº§Aâ éU§ ƒì‘F	lÓ—•B	X‡Ð£nš¿Ò?u˜Î·…¦ö\¥Ë50|ØúÆÔ3fÄ`ÑŒÔÞ4µ1Í
wÓë¿¾Ú-}Ð{þ d­fN!óÏÞ]MSg£ÍGU÷ÎÖs7ôXæ);ëƒD)³_%ÖE›_gñýCMhÝ«|¯5ÙOâaNÃ5Š‰æÉ×aéUM¿I5µ…ñ]¢+¥žó”[àÛc8w6'HÊ¢ÃTÔÚÐQl"öV®¯”T³°¶Æ”š’xŒ['‰Ž¡’ :ªÎ‘—®Å†ÜÊ
Ç"u»ûKYóŽ¨¹·ÕàœÂ"sø¨^+‘6”êN‚|«ì®àt¾NåFöýZ$ñ(ê®öê0oHåÊ©2Î­åuæ¦úâöFlÏ§9™µ²Ì˜…QÑNG„ŒW7ÏJWæ¬F°»X•ÈZU'Ô¼¬˜ÎWKÒC•híSœ³ià¦±UÈ`7Xé•µºŸí®›³uh¦:ˆH…Ÿ¹ƒëÃt¶Âsb;¼Ù’."¯\k·d0M´a#Í¬v£vžÌÿµg q²Ò}:¡%jí"~×ÍŒWjÀ#&‚e¥¡ý¤>:-û¯)d]lA”Prc”*›QÃÐû6Û–Ã Q«bñ»gÒak©:€biœï YIÐƒª²=;Ã¤þ“³Þ–P84Ãó	ŠK³6˜‰Ç@oª£Žæ•ÃÅZÒçXfº¬Ø¯FLks®ÃÅjù–D%U#´¼_y¢})>Æ°÷¸ž;…6 ÐöôÅLqlÚóåƒMÈÈ'3*b(Õó¡ JwÏžC²¥§yôTahÍˆ]ÝUN&«óa~ú(K9‘:@%0fv\b Ý‚1:»UV‡—Í}ÓÎ6õò5½Dð*½”%T8sÖVBÞÙàý6…þ›Lº"b»
¬5ÐBÊ7_¹
Úˆ«Â2ƒQÇío¾ÑôºFœ=Ùœì_BÏ©dûðÕ¥hI=o‰Ý(4ËP<‡â+ë)Y!›Xœ'\‘xô‚²$¤G×tûî ‡øWªµ|C½,ïÔNãBùµ²¥ÚVnÏŠ2›p{ë[Ø6H¬¬,ÓúÔÚŒÓSýÈïîÔÿÛ­Ãƒ­ýý÷¯v¶¾ÛÑ­šñÂÈ€ßþzŸK£z:MâÌN›ÚÊwiQ‡T\*Ÿn}(ÿ*Jq&Æ1<=£ôÃøÊØ¯ÉÀOñ]®i<DÑ­™%âìïDtZéñ zá²zþà);€E¤ÑÀª¸«3o®8¦d†cªT3‹+Îãkùv†ö$hÚ¿a]Z°îÚùn:ðO‰ÒGÞ†8)Šçä*
€ßQ9>’60è´0P«OøY@’ž“çh´]åžaŽ×³ñ¶Ãf®(âÉ¦Î+UôK#ÁêJÃþj@ÙœP«2™—Þ¥!ìö†;…ô³ˆn@*’0þ€¤¸FŽ³˜,úä{ZbäAZÆ9-vŒ:‰òú¯ç1H9ê˜:{±lh™PšŸ†1»¬ÎÎÂŠø4à‘¦±~uÁžd?D¼89 ðäZ[e”²¬¬¥²m?f¬>µ.è °k+ /Ñz)$§Q ¶ +'y…£91nu®î†Ê¶ ³œÒÞ„ËÕ·2›dª¨Jy
ØdõmzŽÃâjk¥5åŽõà½šñ 2(ƒd+™¥ÔÒÑÍò‰ÌËñÖ‘S¥+•ébeê«z~av‹fÝi½QªW»µT5L;®Y·š>ND\µ_K¿%N/(øz•À`jôwžÕÇd1Ï¡Ä9³hÅ]s…„ìNON¿YY!Ô/ënšåÁTó´B•Æ/¯ÿòÇ HÕ†Ð?2¬ÎŒ%i/t9
1<´8÷cÊéŽ0ô“@­ò 	1Ä%´x§Qþ¤ÏÅé‡?
»IVÔbD„íVw?‚ÀV¯Ü>8ö$ŠÐ±¿¶b²Øè»|EÉekÇIw¹ý¨h›-¾¨Â˜É úðf uV”à<Z]™ž¿sƒT†A¶ Eö9é©™á á³ |!ëðµ½¦U±œµ@›·@3ZEcjã|iš1Õ¿1Ê·SêY 	Ó¸Xd#˜Ç} óáƒ'û¦ƒ›`4Š¦åF§žç‹ÿ˜Y*Ÿlª¥‚éÓF/ûªWcL“åL®Å*`F6ÍØ¾º²bº\ºùÄ*Þ×nN=žggi›ye#tÍ«Øš×ÝÛWC0Ð5=è#{SòÅÿÅ¤€+·K —ËºUWî›Ææ¢Á^	;]ŒØ^Q†³ÃÙ”²°ÂŒ>‚À±tcç&C+„é!rh Ž;§nÃ“³åNÂ£Œl2V˜“+KæëmgBÝþ>Û\¨;÷<³!y¹;™	Ñ‹°çB²†Ÿq&²ièçÞæ¼ó@e°»ÞÇ\¨<ügœŽ<+n1 NÀ:O#3Öyí,Ã"KfÐIÏzMûŒˆâ‚îi#˜eë?Âxüu““h‡!´ùyñƒï‚©ÞšÃm][Ô ]°C'™\ý¬	|²ru£^œJË¸eT@™”O¹/‚¤öË±ù™]4>èr·«uzY•R{äk8ˆ¦/âó(ì®Úþd­Qò¦H	ðÄ8¨‡@Ý*ªgLÑõ“ðúr™SÌ4C˜¿ÎN‘rq³ðh¢ž‚jPQ%bâ	ÄîÂîÀîdºµ(•ZHó.Æ«VºF£V÷×z&çí	2Ã*Ë‹¾0äÎ”µ‹oÚõ<oqÛåq `º9nãä˜y2ÆcfÈ¯¼H§Æªßã£lò½5sé˜MÛU¨7&kÂlŸ -|_7LX›ÀÙv»?Ý„9¼¢·¾4»ùøƒLÏ;Žïž›®êsã`Ò_ÚŒ¶Ø“ïp]¢_ãÜ;ƒÞÞnñÕ­ýâ&­%ü÷ÿŽ'nŠ5Õ-(bS”Š[Çú(_Ü}ÞU%-'o>¿þèw·›X_ôº/nFýa4LPzƒ{Þ’õÁ9ï?<íáÎgsÎ›îøšþv\MwD‡ù—¶-°¶YÔª¥ðÍ×um ¶Û!dM µû-èu£ÈÈ.ÜL$tÆœ¾<èŒý%	ƒÎøu±-ãZØcSÎ›µŠ*ZoQi6 =q?ð¬©1P…¥:«	P±%­¡®ÿoÍJÜT¦9ŽRPs¨“¤Žh¾ó#¿zÕl uO"þQÓ9”j¿¿5-¯iÏa³Msê©êÌ\—ä0†ArÓÐ=Ô†ŠªÞlÃ¶Óðkíí‚9y«# Rubk	¸ZfÓ¥Õå5²D1•õ‚&èvöÎÂO@¿5S[	¤.žF=ŸÍŠ M“Q5‚Ø_7Ài²ôàmRM'|ÄI]¯\­ìˆÉUC¼Î–ýÚ+®V^›7VÝ«)~6€Úµ7½š%e|àâG{Sl\*ŽøÌ„ÿê=²VÎ ý;’æÖÌ¦k24ƒÿž™nòñrùmêóY§ï®¡ê1‹~>Ä0Ç$8QÕceÓÞn¹»­Ç
kõéÁË:Oš£ŽD¼±Êæâºçâ¹]·¯&·íáµ¿ˆyWœ‰¶™uÅYñ]Í9sjÙ0ã,›ïQ›ùÉù¹ç{ô÷9ßº³Ö6S®»™¾«Y—ÎG&^æcs_´™ûBÎ}áž{G|Ð¿‹¹WÜ·šy%ÿÌ;«Ðê‡½e`Ð[…½Ó` ?]PZÏg$#…)ëDò–(a:dZŠÝÈõ›ÒEy¡ž2oaüXÈÁZÐbŸ	Š+˜hyb,Ýoqé>–zuM·šú_Ö—Ë“9kpÆJnŒËMÝ ³swlnî$qÃ}`HLGœª;Q×ó†nõ8s»lSði…Ùô	YQáúÑU3âõ<së$·yËà¸ç-#¨Þ\˜â,j±àãqyJÍš~ùK§>Ú§³a]y¬uÏ)±tÒö«Z‚Lc…;[Ž[™“Ýúp"(ª~iÝãòF¡¥ñKíÄèÉøIýYžÜœþ0ŽŽ+,ÕcBW^*ðB±hõ¡kD¹‘ ‹dì‹\;ß¬¬?ËJíºÙ™ZpsãZ2–›[—c³X/‡Yx¡Â¶+ ÷K„ÿð‡“¾t¸ÓŒ˜k× wîcAK;$2a©ðÛÚ¼Ø¸ðxeù±y3 ]áõ2tSÕÊÔ©R|TmžNRÐ¡0 ßU÷UimÛu=ò4R³¶Û•-Ñ²SzÓõ­ºÚJ_$­6,ç–uw «»Rè)CØÅ­X;5³úpäv[‰1Ú7‰j:¡™ÑJž®¹Îc©–}pÏs{ÇiÅkuFqvâÙ1êFÏoº³ºäoÞÑHšÁâÙ1µV”˜1·¨T]+GWp@û•þc×6æ0Moèå§……ð<|kPLgù4i †ÌspðŽ~h8TþŸæ¥‚ò‚œ~Ï·eU$ÊeÉ|“øÈ°ÇÌ†=\Ã1’Jm·ï4Ø®é¯b(l<54sœ—£ÔÌ”«*oi‹ßšúŠë»è±ºÛ‚LìÊ¾ÔùåSÚ§šËšsµ§*’š›¬néªú\*ª§æö´ËçÆÍ¸¹Zuè¬š[—Wœ4io®vUõV}ƒ¬Iå’›q­¦Y˜5Y›æ9½•Íó÷ºæ=4ÒË…ÜTê©`^E¨§šù¢žJ\ŠQ/ÈíÌÞ½f®N|F…jÕC×yø6ç®WÃ–ú6¤yÔ>í¡,å¢ëroç, j7|Ù¯Úª=lû'ó2®%ÕNj€6¯²£•—­e¡DXî:ŽªN32ç,:dK¦@ã„Ð	Û¯˜Íb«õ«Öj¨âã…¡…jtuæ°œ5Ü±?¢åqò“Ü§`n?Þø£AÏ9Ëô¶ÌÞ¬ghãÖz–9-øŒS\EÚiVe8Õ³ôe$U¶SÕG…¸C¬²w*)|¼À:Ìƒâd­2ÕòÁµ\nÂi«]†Ñ—Î#[Ö)·­Ûµ²dÛ€ŸQ– l\>vÏ…ƒ°ë§J¦Ô.ÿã‘ÊnÀ½ŽÒ“Ù¤
›¥^ÆeÏ
zˆƒ~9}_ÞÆæ	9Õ½ê'‰ô °m§ —‚ÚòmD–ÜQ„^í/ªWÒí0VÑ´kûóxXeO¢ dRŽéz“’HæiXÏLQ¬ÜTõäÊ.oÙ¸BU…¹CÇbãÒç2±…§Äž
 Ì¤pG{Ð:hzÂå3"•b’Þæ‰^®Š$E°EÒy?L‚ôcGÛý²ÕOyÒ²q)êHÃgÁ"zi+ƒQ‰³7ÿpdmæz÷ÆÈÅÂåÆýË º‰ã=ZyGøÃ, tgZoEÊÅž&`-Ï¿\ÌØëuKÆ:œò.›Î6lEqû­²Í~Ü½)ÞF“)’ÖÂ2ªßePde–•Âtºå;Lú‘ú®…]b‹)	 Ûþà	¯a÷ä‚†5Ô|‰Úí‰IØ.¨w¨«ìd+Áxçþ_‡z@£;‚äÕ;ßDßéü=wŠh
ÕåäÄ»¼xß„Éàèó}Ç¸]Ñô"I¬‹Ìì®zâXUWÚªÒ`–¥çM0`§åô/#WM0Qìgù(‚V÷Òäbã%Ç*ßJã	Œ—G>RwRzÔ'¢J	ùòMŸ‰WÚèM
àsy¬–Âµ»a­[­Zêtã’ÿP§AœàdÊIÝ¸<ÒfËDœÑÅ€t˜×WKãB]¾÷„‘êÅÊ…Ž·ãèŒšŽ0ZPŽ2B6ûÑŒ©p¥¿^ŠÆé>­ËöD,£æ­p§Ta·_¬
ÝcãJüÃ{ë	pb4Ï³‘m”‚¤apî¾`Àž÷#&Kù;„÷·zƒ± [ÄcB@NÂWí¡3-—ž½éôÚÁ¿f!b™:˜ÆCQ˜€@wÈðßUX¸\q Îõ%¹TX¸WÏRêß}P÷ã`ŠµsÇŸÓ<£Qá§¹ýašÜ¢X1NrØ‹Bƒ0A•†€¤ÙÖ(£;ÒÑ»ÅWÂ`O(qDi/ïÄ€t'UÌàràPÌFèÌ·CþY:ßÅ…~šÅ4ðïò€¼-¢|Ÿ½<qô@nƒµ+‡4 ÏÄÏ'êè Ge”T}‚­R	j\±F›Z42Ñ;..¡Md—.•-…•~Rm "©†¬\!Ñï¶ÒŠlÈ¼­¤õŠ±‘2(¸¸³¨ß‘QCw;ŒóåeFÓcfOK‹l‹·šÚÅŒEëæE][•­þ¢a»_×;TÛSSjJÂ@…JÉWjJMIÀe¥Ôwâ­¦D\Èi‘íWò]+UíÌ¼`Æè&^„jŽÊ%A‹á.Ö¬™t–$O»ø¯Z¿¹¨xÔ¯oû–QäUpÌ†‡ÁP€‹½ie;aPœ³ é¢O`3
:OÙÚQ¾->àÌ#¯šÚ€
vCÖ¿çJ‚ÖÀATò>Ðˆx6ÚpO²3Y¼@ö†ØJvBMk•ÙâÓ­&µ‡úÅ‹,ÿ='jøÕïx›Flíà¤L£Q|”ðPþU4Éºª‚3Æ&qå{QˆMÐX,´ò²ø}\žtñÓû*Ò_¿Èò²ÛÉÐ`ìQ0UiïWª ŠK«O´âCGñ¡¿¸Rš§°¶ú	e¶³É4È#^kå®žÅ\Y$G&˜Þõ¬ •uª3 …©0 Wôà*J¤_ˆF˜ò¯<&¤jÕnÔï¶ö7e6o¾¨æJ8‚8ÄXFkºn£MthÜŒ{nÇõùÐfX)‰Œ‘^ô?(uÖA·¢.€Ø>˜ºj¼‡ÁO« ô‹ÂXp ¬L–DAj`°ëSc¨X#ë÷ûC1^#Ö’£Ø‚0MEAƒ=WsV7º­„ÆË¾ƒA	]9P?¦-`œÙ¸¯t¤q`rÎyŸõÅkTU+˜ó…<äw–Ç4Ü­ö&a¶^î¼ßÞ{}¸µ}¨Gãà­îØøO4ñ”Ÿ™rA|ÑU*¾ÛúîéBd ¯2
+RRíU
ÓÃð8¥ÀBU¦e ¹uì$ë¼L¬1qb ,V‚Ä¼±\1äÚÎýgwÞp!„T;¼ê±IúÑ˜Ö)6y"Tb5]"…:Z/Òe‘Lœ{,bY\áNe]l¢ÅkåJJ¤hhD¢%S$[Æd'EQ‘€Â±þ†˜ÌÞ¯Ô¶z•½@_¡£,×-äL€³+ «uÞ|§Áˆ´¥ïa<M²	vé9ûÅÓ£WkgçõÎ+žrs(;Âþý~÷Å’ ÍvÒ"NñË>¾‰¶è¹DŒ\”ßƒ-nã_64edhéãc½”38%ê‹Tœå‹8_Œ/(Îì}äeõØUŒ6ÊÕ8îóì”É@µž¶PÿÚ©ÜÜ‘ï¢Ü+mdÞû‡CÏÑ{§ÓaÀÇ²Tu]Í;¢Æ<PÈ7TÃ+¸
æØP9hî]ØáW=¿¬ðÒBíËÛ^¬š^T;·h¶~%>€¬ÆãÑù|vA¥^ºsû'vœg³)U	Q…Cåw\@Ã(rµ!é“²ªdEÐÇÃë?RŠ£2©³ý×bi¬ù‘+”).dùN0:;¤ mŠ#|«Q¢®Ò5ŒÁÈ¸ÝË y=R}øfCðÔ‚˜sï€ý®ñm—•è±}û(`«‚îÌÒò}›~‘Ý…ïâ/LBe¨G©PQ.‰ ´ GõÉ‹!j½^Ã\pc½;šê½ß)Ó)lé”¤­ñdatc¤$àÔœm§»PÂŸ~R2 °ý›fÂÀÌ_’48×ÜT3ëŸ¬Z+0™
›uŠ5¯ÿþÛø¯ðÛõÊeû1PpypwZ"NX>X
œ¹Æ·Ÿ‹Æ&R/Ô1ƒ'.ˆœÓA*6–/8YFWPCB•3¡'ÈSFS¨RÒ‰œ›,EÎ¢|-Ÿ*»öêx´RÌêY%ª5C£(“©™e#ƒ×ÍªFùÑ]×Èb5šº–U%èþRõÙÝÊøWdAQãêÔ€~ÐG¨ô¸*Õ\Ç«¶«)‚õ–ù§M_ãZQ«y­ý¸@vêóK1rŒéÍÇýJd¡m©”HÓfûÚëÙõà€D=T2öT&É_…Áf]ÊúY¨FCqJ]UÇ‹ÄÖu¨vðwr®m¦M?”M£}«J!ÁBÆ	ûðÄ¨´Ì
„Ôå•¾N‚‚]¾‹ SQt¥ÔãÑÉ!•Û….­–	ƒ=ä@Þ¹×«sAp`ÓÝêÝƒˆ'ØIxdýrWHu;&X8¥T.1(´ÍNäÊ‘ŸG	y±Í7#ö‚§þÕÔý÷AÁW	aY¬	™“©’¹š¤u©‚±š¤O‰«Wm#fê4T*}”ª/„Ùÿ]ÅuR¦9 –KÕæ‡LÏ„™†ªVbD\YœÐ­—ÛyÒèŒ‹zÔ¥*76Q½W™¾tE	»l(ïèÈF!²2mÚlÍšäEh¿ju ûåÐ7=Ù0ùAÿ„72+ŠçUËÍÑeÙp8ÖJ9ƒO1-¡Ì3}‡¡°|òøøŸÅ4ËCdŸ*²=ð8vªæ\°ã_€¨ª±œï
ŒUÍpTMÐ‘³–‰¡o6NS:Æ!
°ŸQe´ÌOH:=ß¡È9äèeqtÑWÎnu2ÑþDbLo“
ÔÜÓU\YwXÂ/ßÓÜ‹]ÓXÅEQiƒ©ÚÝx/¥DÁ64ßF'r3Áê=·ÄaùéOÁã±>Óirñ†òa?±']´qwà‰Ü*ü†çå"£uüF‹:D±b®x@>#ûH"‚¾i†SªÒ?e õó ö÷Z.Ô‘]ò¡¸&hÛ¿ü¥¸–£«y{Æ¾1)ÆTõè0ªÅMÄy^ü~µ£í%´Å™¬9I,gŸÆ·‚ˆ{}Øßz³õìú?½>àáÙÉÿü¯ÿùÿ [ã,Èi6ºþWrýg qÙÁðì£8ƒ,ÿçFuþƒëÿDÏÆ9Ð˜¶|<KN`}‘É,.Xë<”;‹j…Aíé¡IŠa¹Fñ(¡1¹h\w@qâã4ÅÁ4ðÿþèõƒÿó¿þ—ÿ‹l¥PÌ§ÐÐ"Ê¯ÿDbú¸H¬ªÌBŒƒ†9ÈòfƒyŠn÷©Á›©wØû­íNÍ­%ù²ÁxVI}ƒaŸgƒ O´4	âI9P¾YPE“Rå?R(}<ÁÁæ†9Á‚é	Bx xÅH€èÙ‹ÙíåìŽ¯}@èÉs\ÈüÐ•ë¿Œã’Ÿb_·÷_oÈ?ÂÿÈ „ÿòô*Twýo…æ,¢ñ,ãTd|ð`‹™yÄ0CÐéQ”@2ÔW1C"Àƒûd§—=ÓèÂIYN‹ÁòrÊøñ>/¿ÌÆYúAîtR–ÀÅô	×Ì\+…gþ€ˆä+ÕGÊ" çÀøk a Dâú_' 1Ä	(˜}\ û0³Ãë?§ÅÂ±Ú¶†F€ª„ßò0Z*re(:äƒv â$¹À1ÁÖC6ftBåa¹ût:ÚÁ¬€²¯Š±ˆ‚±l²WQ­CøùEEe3ÕV=²A.³k`WpAÙ}«rçÁ‚ÌÈ0šl¢u0FÖÃÀÓˆ"q0-9	d¬ç6²HØòè8B:Bï÷ÆäÅcgºPÇWšÇš>ÙE)-‡fÄ¸ØRlô8‹	Àq|HyÁöðpåú_S x$¤Ë~<ÃUh|ùW—*d~Èâ´‹ìJ:Ô÷×4fâx„‹•2†,RÄ%?± H}ÊÁÓ(ñÃ?0ˆœQà`0#Ð·ŠÃÈƒÀØ~è.ˆp”nõuœÆ™sn…¸G.ð938™»¦=ÉßÀ,ÿ!ýCúúßÿMQFªI7âCÙZ@.¨/Î#§y
ž@úÜÚÞÙ=Ü!Ï÷Ú¼Ù:Ü[dÀŠ‘fŸj%' ¥ô4¦Pš˜ˆ¿€rðÂö‚Õ­Bÿ÷·öPñÁÁÖ²Gö÷Þn½¤Ml½~¾ûÚ¨èçH^Îé³-I¥¦ÁM¯Ò“h	flIn<´¥µÿïÈIïý¯owÈÁÎ[ºIà¯×¿Ý{ë`~ÿÓÙ"«ñûÿ}QæÞ9€·wÉ«]ÈËJôi•±Ê—»4Ï+	²óüíöÖ6°ë0úå`ç»·0üIÁ£Ü¡Ùw_¿Ýz£OÐ+0°?qŠ(ŒÑG]‹P"<òF4Ø9P8Ð¨&äö<:fIÙÕ™ÙŠ}Lc¬†|#›ÚkŸ³¯Ýå?<_ÃúêÐ °bM"Ø Z°KpT¢u%.¤%W&P• ÇæitR-½jÿº‰‰Šâ¦Û•ÃƒõFuà”h‹Ôžª^…ìjv}<´¼ô%+ß{=¹(¼°-­Êf«gä Û°cÜT…›¢uÏ§¤’ m—e“}Åt†ËgÒþNß¶5«Œ{U'¶VÑf‹Ú[ÊÌÃRÜÈÜrêEF…ùPkêYòoý‚ZAŸ)”9ôÄ’dõÖ›éW¨*s"9?•^B7#RälÍt¬½ý¬Ý8¤V…ªÈÕ/ê¦x©]‘¦<Ž¯Žv×¥ÁU#”,V¤b‚‚`°º”7\ÖI´;ŽOÓTôÑ•))ÿø]0ÝK%É/ô»±Mó¡e>;F¨ÛaÆÆ¿ÃûÒp{±HEONŒ£9–.ï¡V¢ÍRØÞz¹óúùÖ›Ý½÷[Û{;z_Œ;PÃ8â¬xÅ=x_úè²@í¶:±ªÊ¨gAÖmç8ÿÄy®(Ìeõ¢Îr»™|}²™¬ªÅD²ŽØ	éw0‘JWšçšüó¨ßH1î‡—Nª{%ü£#ü¯n¹H§Ý²^¤T-ñq[1*ö¬5&Œ´BÝŒ§á">ãÅª‡#í£iÄHqA3d)•1#›,®—œ_é]`Ü²Ü¡æf²UR­)eÿW§Ñ­ßF›êêrj:Í‹ÖíÝ Ûúl#z%Ç¸yƒƒ‡7Ð#dÖ°W¼”¸ÀÓí=1ò×Ô˜/œM=yÔÊ ³„=šÈ£Å‹ïz
¯6@bÑ—›+Zà…Qà±•TžØíõ¢à¾ÛE{Eg²Â!~Á›öÎê†­«ã‚¯o	úX`>à¸/îh°ã=¼¦!†n*¥NntÂµ Œ–/} äÚ§¿óÄ,
S#‹ª^fõw%ü*|(»Ì§j’ƒÊ{5
„éCRWv¯§™ñZTë‰+½2ì/¯±J©¬”Ò@cJ•4ƒZñ°LcåÉË`VãëU£«Šâèž˜™¡Ó½ª÷UfÙ{û‡+ ÃGÜ5²Hó×võŽ±šÒg;‰3¬î}rŠF°ažßMWiF^ùÙ‘¾‚
òÕ¥ì»zãgÄHtËx/Ó¬(âÓ(™P¥¿<H¯™½¶iåÎ9î:‰$ D"à¹}"÷Å×tWk×0å¾8¶®_€88`b­4Û£´Ò’V¿¨,R÷ÅmÉÚ.u©`›m—Ü6K¤Hãÿqå.Y+÷JÝSÇjà[=_yz  öYÚ-mçÕKì‹ÍUZm-W{òØí“HÉIe)¹>>£i+šÈ()v+Fví¹"Ì5ròƒK¥6zÜ©žbjLä‡¯.é—+Ôv53"Î}löé¹(7ÁÕL[¯>Tü‹ÒJïªY„b”D­Ãlç|%ÝÞié ú¿GØD“x”ùVÓîD¬&©Þbî€©‚ký·‡¯^î¢Oý„ªÐŸZkE]ÜW„×`|/`h+ï*-ý&V†(,6ü"Ï&¬ÓÔ÷Pµá´-^—Vä1ø÷ô>F7ÏÎøu/àÐK½ó¸ë¿ƒ|†Ý¤j‚…²;ÃwÜ›1›éŠÃýÜAP%¯\;æHó‘öè£nÝHÑêÖSk|;1¸Ê6aŸ„FÄ»e-NÇÕ<juª‡žW–VÄ©aƒRÆyƒ^FjÑw£æ|©¾	Pªãn<´!±	¥U,Ò•ÙÝ'ÕOHí|éóïñô«Ó³µ,ºçjÎ×i*´m4všTm Hv
4ÄÓTœvžìcc–t(Ü©¥•I‡› VO’WeÚ<î•˜œÈçD¯7K5f`îqhêJ…HgôÚé@†ª[K­×V	ª.T1-¶olF”t!^~jjYÊ¾Î±ð­â Â¶v7\¸ ÍT_OuÐY•à(Û'~×V¨Éãm`¬0tí)¯H»š5¤ý6ÛÌ$jõ`B,~÷ÌõO»Î¹Ôµ=´Ú¡]¶[·¸BG»qW®¬îûó µ'mf€sƒéµŸ×Áßs4kitî¼²>?\k³è'¼®õ_EyoïÇ9`'O%äÆ¤÷ŸM´=sÐ|ÝI%>Tô§,[„Úõ•'Ú—âc<šø&šB{PH× ©Û­žO× uÕ>¸”%”±5®\OAÕ%TÕ‰ö¡+OB¥Yµ8Ç¬„|›²¨Ã3õ.Ž:E1ŽN)J™¦ži'¤†ZÁßG¯Ec(Ã.¡‡¢×&ã;ß±›Ëõ{e¸A¡«Ó§^ÑËUÃÖFY£ná5p\þæ6?øpä6r>°Y‡ÎT
-©Ên0fÅzá«‡Ýü ®H<N3¦Uüª¦§ÿhrâÍÌyWrfkXc„¬œ‡d5c`1‡íKßÊÐ—Ë_“Ã`H^§ñ˜˜}½,ÀÖB‹ÆDWïg'N¦Š÷µó„'Ðñ3tï
‹Œa7Î–Žc5R¬í3ÝŽ³(½cu¯WšËšH7žA¬á¹{­ÿØ
ÖŒÞï+øº«t–K?^L&¨úg¬ûÍ*¾Ì·"úƒ!‡Y2¿h¦5;ô€tŒ¹V°A]èª ¤9ò~äÛ'A^>Üm8pgÎðŸ‹¡˜ñm¯îsN"sdv?'õíþNÞË¸(ëçsPúOµÂ˜³§{n¿ôâ³îX¿™Ð¸ró8$øF4,€v“AõºF’±òúÈ-w¯ÏlC‡øò‡ÔïUL¼ˆS´XŒsk:•Pêè@]\˜~þcà6.Ñ•¦Þ:—å¡/Tûò\–:n4¬«M]½eŽ8Ü¦g—Ûq>J¢5¿E¬•[v»ôÎ†Ð•½ÚÎ+®C<ÝO²‘5²ªßBß¶×‡Áy€>ŠNØ=ÆŸöY÷prõ‹îžællo§þîWkú¯/GšÒ´ˆ´U³FWÍ·æª1êP6ðoµÜ·mËÝŠQ´~òÐmŒqiœqHuM†ÞØØ®Hë‡T™­=g$e«+Ésdóâá,ffÏÚ]feá™£Z>yhÔCó::}©yAaêØÂt˜×L£}teqåjÆ°ÔÉ´ý0+Êøøbi•gQ”êÑüŠh×† u·”ø|q§ûkâÑ:½_k9 óéœapø_šiKúêÃ +¦p7SIá£† tzjCRQ¦ù*QÜ³ˆj("ÏW‹ri^Ô"âÕÖRÕ$	Ñšá§on$¸¢mBã^"Ež¥åÕe‡tüC¨kJFáMQkYlµ{‰»–p!sõ_ÎÆ¡×Æ”òÇZ5–öÅt¢ÅÆ4×A¶ò8b}‡€¥þE•{‡Óz5vNX«Zjný¹åº¼ƒUykòö+²f=å0.——ä,Ë“UJhXùÁ}Ï˜ËÄ×ùÁÁåÈh³._8ãÁÌ-·Õ½þuGüØê±ÜÉã9¦-xŽé—ÀstžŠ~~†Mfú™7™éßû&#vE¨®ÝQæ¡•Ó{L+½	ÚkïÊ§žaÚ*C5S§q.T•š‚ãÉš…áH®}4¶óïoI±*¢¶V/ŠÒùCàúbj¶{[k£odtØë«m»¯>n;~õ±lúÕ§É¾_ëHƒ­¿Ö¨m÷¯>Þ; êSsÀî–çn€6V×=ê1nèq!ÃŽØ±¼uþÃ¥Æžg€ªdU‘\©¹RÞÒÚ0ü­N¡Ú+†|Ç~2+8»±¶â®Ê4½ÛüÚYî¦ü¾è«ëÔÒ•ÇFàûÛÂeÕ†‹­§hš/M³ØAhßNiô;Ö[…È.?@óm…ë1ÚO:p«¼˜"-ŠËË)>ÁhMËNÿ<)Î	þqešDÍ57.U{N×v£î|lµ2Yc\_¦ó67tÞj±;Æºç	7.yLœºu„BagnähDÑ‡Aeâò‚Šûº?á
U}-SãÝêãªø¨d–HÊÓpUZ›¶‹Ávud}ç"Ú;>Ö°º€ÉZšN‹d¨€‡"qšÄiääªÜ½è¼Œ1yAÃ£–ó¹4ä§ùžß¨Îùf³äèÔ|¤¤ò±kM„¹b7uãÍ©«c­jA–Ÿggi[BÄF<eöñ„5ìœ"N?šó Þ:F˜„Ê1ÂCýTá[÷Yœ‹šR{Î€h”oPƒÕ~¿ï;’>"MÔ¨(i7ªb@‹Ü]iàN+²–®ž›£ŽrÐê„Éš"r(§ÈÙ¬Ä¥r\Šš…Ñ¬ T¨¼´;l„Ööþ?°¤Í6°’Aé¾PP±8:n@h®CÛ C‹µ÷… Ä¢Ù”Úê0t:Où Ö—Ù‡†üª`ñTyiY\DTè<¿Z´Â.<5SÜÙ¤0÷
¢î´ýk¨r1ÛjUÙ¿´™›äìs³a?Ì4/¼­–›Zà…[ýÑ|·ˆAóH#ý1ìVZÒûáS+©eUµ>Êm¶¬.»Åµ£mâ¸é©øÕ² r@ô´ú=!¬ÃOé»nÊ8­+x)’z-g¯¥>ÝÎ&Ôˆ‰ùÞkY’^_}z5uÍâÌ½ áÃVŒ}½jÞlDæ>_B“¶þµD‰Üq>@'3‰ŽK—Æ¦<‰‚Ð©3/s§Ô÷XúðhuezþNÑÌ¦¨ÌÇS>èFQÎâÐ©6â10—ZŒ>"gK«®#?^Ê§%bÓÑÐ‹ÃìÜÁƒ3n¸UÊì©‰iªÞœçÉÆeýqäJV¼üçZ0âRV¤B„ŸœûD¥<™o¾`…
g[7*Lv.É4Ï_…¸ž.¸Öùk¨nËÎ_Úsjm
äÕã	ûJ¯5h§EV<ê1t;Uà^?Fë‰ZD|«žÇr¥Rºœ*¿¼µÕ¥
Gîj¨^ÿX=Z-ñxñººÃ¬ÜFùãþ,OîÜþƒ®êÎXïö†?z˜CÃ£¤û1'}µÒ±õ¦iÎƒ³ÓàõhØ7µ«Nÿ ¾§¬ñâ{¸ÑÛi?´ÜTAêpVâ~®zÈÒW°"€qx–•]‘†BõÁ˜ý˜gQÊcï¢Db´ŠúØ5å¨^ò'ÉØgz³Å¼ ÊÖGT¯¾3!6“cWŠÜlIû(2¤çöfÙÃß¬—Ã,¼Pá
Ü°'K„ÿlš[·oyË‰dì7~Ê3RjÚ#¼)ù°@=|0¯£,?6µæîYôrQeèÙ	ý³WÏxµg½æËKÍ¤§©º%bòBµËSå“dõ‹Äb—êÖ«÷›×*°ÑÉjÓo7˜—}þƒ'µvLÖy–ÓHä7Ü¯Õ{]sµÆPÎ	‹4]Ôh×Ñôø^ô†Zd”û'°ß™·joÓ¿±ûŒbÚ9Þ$
ãÙD…ç?´ž]êÅÿvóÛxyÖoåÔf”\Öj’ãã‡&u#l3Æª>L®†¹½ûÑòÿýÇÿÇèCý˜ÚªaéÔ,|Ü!šaSƒx
Þ¬DÐþÀÓ£ÛÎ~”îù8Ëƒ)¿¢>	àuéhmY´K¯·9Ëð,×4©>”;ƒšW—þW‡yØ´(ñ‘J‹¼ƒÙìè…º‡½æZêØ#|š@6ÇE2|:Tš”[©®‘såãoŒuØþV{š`Ñ ›Úmi#¾µí	]³ˆrÚ×4ß6k Žõ">Ü÷bñ©—Ûh7îýqž¾¨?‰áî”´ rsIÚêí‡‹7ßc“vtKÃ¹ýèÇ>õêêûZmO]Ü´l‚fñ9Ú£~÷zV…ª³a÷ãˆªˆiZ}w¡Ti¾Ç*HÏªªBi:øPOÏ–nS$DßãDÜ§‡mZnðÛºÛA‚ùÜžf\×87aÅ…ˆì‘äéK;;¤œ#·ÞI¨Þ]v@háÛï!ìJ¥¬€½~{ó»Æµ•~ÜÊÈLãa{MëƒÚ[ùk…„¹»3§ƒúÌaÍpËNÎc¬ 7å2\P¿Ï¿'×ÙüÖí×w{Æb(á/^dùïyPÅ®Ë¯š£ünÁjÈòº›,Õs3->ÓèViÿ iö…æDÑãOk¿.ØdÃ˜ÚºôQGO¸–½éõ¿\ÿˆÆÁ1·?%øÖi’­ä¥"®h¯Y'ÖÄ”Þ=ji7ÖZ¢SÃåµêaçW•zÕåîÒ~ª£6Ã7b«Òôê[åé±Uy#NS^Ñ²ÍWä”áÚ—å7:Uqâà,»r§èÛ•«
³¢î¾G»ÉÇ_ôßc\ôcµÔÄ$ð=ÖM@·n¬ˆË‚ºÎ±UY{ˆÄßíKa|¢ªÜnØ¦äUšßxuQ©on‚ÿÈAðéÝëRž~`Ûöbµ$ÜH ‚ü†¤¯µ#ÖŸ€ø2ƒf¬ò]0mwè´\1 -­Y4hÓÔÛB›Ù³ÜÜ‡yPœÜ¼ý›ûÜ§¯¹tØ©Ñ°:~}ùº»~öÉtóV=£,9€ý}ãò±~læêã8Ïc8u ä”òÍ‡¼,Bt;ËãDöïØ®«|ÔÁ¾|FM:ç¿{¶•Æ Fû°DéÈ¨à2®è›k&-¥}|…ÔOp¹´B~Äubó›ÂQÅKn”Š£h0„¿ažM‘®å.1Ðì$ÃÕÖw»¡ˆSX‹A‚þ"øÓYY$Å(H`+^éÿæ±›Z¹UYnÕ](:Ë¹[ji¿¸Y.;÷À+€¦°U»ÓàÅ˜&tÓÃ-Œ‡¶©ñ¼Þ$d+×=>§uö—j`\Ôp°Í‰HcQzõU^“ç‹ÎãÒÅå‡©ÝhlW·vØÈö7Qv—ÊŽ³/iüÂ‡À#/á¯#ú~¿'È:«ËÒƒÙp‚h®Çívß@³ú©kÜàÕQ“gh&CmPÑJ|5Ú‡oôIªû22Qô1ó-ås–„^„%ÛÙdŠÑL½DßyÁ_~®7ŽÊ£?£ù·C~(!U×Û×Ìaø¤I Zg!pjgÀ,{+kµQsJç¼î»…Âµ¾}{^×?½>7mï¿¸¨SkuSÌ© <±Áí‡ª˜€Ë¼ÕTWög´¼!ZŠxO÷7…vê¾"h¥MÓ,ù~ÆÓêÛ'ÂS·ÛÅª³õhÚpèo¢a£ÁÝO‹…•†Uß›™Ê†~5JÛ¹AðVÔÞ!‚·Š›ºFP*¬=b»g‹)U}ú'Ù,,×Rãß×õèòë§/Í/lM~[½z?ø–ô»‘ÇÄùÔÏÙÝpÜ{Êg!fæîwƒ"Õ¡Û½GÃ—ÍÎù€¬­¬=ê{¯~X%ÂÇ~¬'ÍŸ­”î{W_š(!€?¦(&÷•‹Ò!~fŸä·O„‚V0èÏ‚ˆšUÊ}EEÓžægd”ß>™„ùoÍV°ŸB¾lq¿î>—ªMÖß6ÞŸ£À-æû8ÅÉÒÝ›RW5‰ÖH{‰Fi ß¾5Ê¢Í/Œ$ÕAÍf;š¦XƒÎ3ÙZ¤†»³Ž[°
˜Õk0"‚Œi‘Ó nö@9	3²M¡sý' O}×Z\n¾†I9TÓÂw@ç¹†ó³’¹nxáƒ·®hI¼c2	ÒYtÌhãîgSdö´&aâš¢cíKùcRèO“)'Íã
Mî~ØLL0xÂl˜ó8NÃ6—u	étòƒ~ÌîÿÀµO‹á|Ö™kÙm1ir›â3\¿ÓÙlÈQkŒ‹ÏÝìw2ªzým¬&JgyÄd–‡a‡
äTì;m.ácn/ìV|0*[\†W;D¯¼C1tG$è!m¾ØÝÀ¼Ôf…téïPo7ª×0)²-[wÇX‡/§CÍ#Ø›¡Ùg÷y<¦&j¬^šzß®ïM®Xž&O]s‡Rw¦ûšFá^ìÁgb0P›jªlãÓ–÷òx¹ÍÝ”.¶`¥ñ0'W8ˆæ{™â˜Áf-Ô‰kÂi7Ò\j.ïXìùL;Í<2ŽxZ‘f|j·x,Í÷xfÙÇE²¤ÇI„áî&‹$*G–gwó©½2×‚jÜÔ/C¶ñY¤iµÌ‹³sak*Û…_õÈÓò~É'BUvNÑ	ù§Y§Ù¢‚9õÈâE¯SÅ»–È.ü0ÂYPdÿ0Zº3*¦3~ÜÅÒC-8‘á@³&Pœ+æ•§C~ŒC©Á‚MÒ9’S¼3îö >€ÛŠY}í*•æd/õìNo=½Ñ,ò+K}Lñº{I5Á9qÁ:.„,W7Bì›1V¢†ÚëËÞ+0ërq(Á:cíš72C…çâ¸ÁÓðL+jQºXå†Ñ„ÚY¼–‘ûâ±}ÝSxbæòÞt·‚ijR÷´[ÐhÓh„±Çøh
½—¬èæ¤;zÊägô£¼n{dÞæI×ÌÊ0Á:õ1¼õ33ÔºH:ï‡I~ìø“DAÈºx[4mé±µ-©)c|`!©¤­kÌ¯u–ž§7.åOÜÔã0ƒ÷"‚·a
‡ç >ª‚¨fÇIÆR2C¬ÛÞÒ›gÊ*‚½³•Þ¢ê¬v:ÌÒ\¹g[]ÆøG’(ûx–Ž¨ô·­ËÄßÇÑegCÎ¥féa$NóãJâO~ßì¹Ì´sžïê©³"ÊüM¦@g‚b?ÈGQœè/ûˆ:P‰<@oMc|C¼©¾Á‚“¿Ù”ˆ5³øàj@D'Ú‚ìèÝ“ªÛÀ–M„/©ERnÉ.f£¬_I'Êó,ïP|9ÍâðI5Èy½ßg/OôC-0|[×›~
u"9yâ	ëjMQøh±¾[à‚’;,M$±Œô6ÿ90X40Ma"7Öq‰ÇÊ'YˆPønçB`ïà°ó„àõJHÒr…Åh?¡S“¸ˆÖ!õé1=ë%üe€)‚÷ï‚)E\yuª°cÓj%.fA:/g	ZShëŠ-JF2ªèz‹zø¼w€ Dt€u;Hø Ðò2ùm”LÅ8Ž'9Á  °åÅ#é¥ƒ¤°b(IšI¶Äùe6¢s1ÿQÐL7L€ŸitFv{
¡aE/ k#ìC/ þ7xïòõÊõ§°ðO wLEs¾Â4@oÈj¯?¥½ò²»}¥£ƒ£(ï‡§T•³<ÅÐáØ±«¥¯.ióøªºú€¹®(ÐXýA>.
:\6øjÖ  à*Pìa”÷¤{6cüSèç¥ûÅ4‰ËngItÉ0Í#Žq€?ì‰Žv:úhupszöšvUs´òNÞ0ÐÒWßõÈYu~[ex§°~ÁaI_$`aJpÜ:8mü ¥z&4Ù­ˆgyðcœÄAÊ±G@vSâ¿LÓÏçÎ-!Í‹™h" tµ,~¯*¿WÞ©(#w6ÎÃL,+v„.zÇ •Ò,ÖøÐà=þÑSš§ŸGTôë.ÿáù2àbG%ÏÛ/ÿ‹ßÇ°”pô”Jd–Ùõ¢»ê¨@¹Â¿º‚»²ëÃjÅ.Tõxüø«KþÊ@$¶e>6þaé¨‚ °àáý)`ØÍÓœÆ Ýî7Ï
ž:á­úœqie£~Yo¯êüU…ÉÐÔxå»ôK˜óZ¥›0~6Ô¤CÂ@t)du²”jJ‚«HOªTázúI–ã&«'fCØ9OQ‰±© 	<B‰0ËwCÿ·¨Ø¥ûûÌv/Ê €`…_åQk˜™Ù®,Ò§Ç¢ÂÒ?¤_3ëàkòÕ%
ã«á¥ê
ü‚,ýà ´ÊIA‹ü6ËéÑ¯Z‡=Å¸^´cPa‚)AÁe|ª©àè,õá	“ac]íÄ"RfäÅóï¼å£8ˆ’ÄéÇˆ~aµª ¡LoÌ‚?_^Y7}ùø÷ãûne›¾IŽ\ß+µ‚H_‘)èŠ ŸÊêf4<<Eöé6åÖ@Gïzì°³;£X0ëÏøAfJÞ¤("1ÿÙç$¯’88W€2	ùðõÖi\Ð˜×¯³Ó€l•j]6mçq_#ní%×Z$_šð*S!®¾^øCúvÀ";•çcÇYLF´ º ï GT›‘ XÑÑõ¿¾ºÄé¾Âê÷(ÇÁi–£v0›Lƒô$":7+ƒä$*”ªú*A„ÑÓ.ÄØJHA”Z•8õ]TÀÂ#o¹£ª
¥Jè Ü¯
2~ZW€Üç—MtiŽËÜPæUyËÇLÒxBšHÅ&Ö»fý<K˜ãÖ7{/wúßíî½yÿöõîó­ç;8Ë3ÑöLï™*Æi"œ‚‡ÖKDFµÓ¦ˆŠãæYmŒÒðêU+Ä"ÝƒY ÀÞÓÑL´ÓÍ¦QB	ƒžxU@âJÇ¸OŒn{ñHî"XüÀüuÅ÷5AÉÖ*J¶/	>édQòxOn®b/eÄÅ½GTQíj³™1íªÉS0…=—R¦Úúlr"
ÏAºªö°R&R­p\}w!ßÍ	šVstúDZ!
»â”åñÐ³ãøx/´¤G°×ò,ÍÆy0	4<ôâ =tS/çÃ¶‡}…ÆåxT®í|b´OKWÅÆväË \>’µ½÷jçÍöîÖËê—z"ÞXLŸb¤¢¶'mØÀÛ`««ŠêØj •0Zó»ŒCÙD^ík¤Vo¢b
™®ÿt1ÖL£]¨‰ŽÜ´ÜÔzâ ÔTzž…Bºˆ€Ú· ÎÕÆ= #Dp\5íüîÝ~:]»<„(	+äY0ú8¦çN,à…óêj<òhcØ6¶NXÚ4+Ê%†Ôr  S0ÚÎñ1­RÏ
„8M5â_Ý[$.Hœ£?´ïB@åÀ¦öñèØV¾á¨¤!MÍJ%AQ¾™¥4Z`Q@ß`É¸Pá°@îv0Ç{6®÷ªvê=m°Rí³úÒìì0ž V…1Îº*Ú²-².	]ÿÖ#ëäá
>*vŠëÓftºhÙéE¡±â}èU /3˜RÈÃFLÝJÏµôøÔÃÊÃùT»¸´´´„’Š˜²JgÃÊG¢$Õ0d7öO@ÝçÁÅ3Z€ªÕpV™nÌ–ÄÉÒª±g*Å®J}[P­!ï‚¤zøÈƒñ™wM„æèŸñC„—ÑdÁT±Î:-B§c:É8‹3€QXBü¸H¶÷^¾ÜÙ>ÜÝ{}ÐßÞz¹óúùÖ›Ý½÷[Û{;‹‚.cpM‹#uug@ÝøëWª€PÍÉß¾|­`°"d[Œ¿sïóðÿÕ¦ËèÚ»©É™·hõˆíML:îpüÚÀl'×ÿ­A¨¥©{xÊ¼±2ÈÆ Š4žÅÀM‰í84ÊÐ…Ên•×&ÐÂ‚%nÜXpuÌ•!¼iâR«U¨QlÕahˆ¯b&m±U<NñU<Íb¬·`KáV<Nƒåù…]jB¯7)üú‘o.Á×Š>P”(ÆÊ6f:‘=·Åã2A±ÓÌýýÊK•„t,eÁÂ"R¦XßNln•ÝÞF‡Û‹ÓâqŠÕu$Õ)N« šW¬Ö{2/Yu‹ÖŸš¶ÞJ”ÏM‘T“e!X;xŠI;ªê¡§7”®[×iªšöb({ZQÏ¶ÈÕZùO‚ÓX)«jˆáÍDLö˜H¢ü&#¼}CºQž› ÀáÃ”ô©Uˆ‰AœDœþ	9t‰BDµnòH†Øü mÔÌ¯ª¿4%›µ>A{¸4Ë…N.àÇèZšØÅZéfsï¤¡[´yOI£X#ëh’iÜo*Ð¼ˆF'CxOD½C?5‡BEœÓÉ™fKl©_ë3ý5¡Iüc6ƒuƒd·qba¡ö®z}Ÿ/*0Îž<ü›y Ä
Û€ÎfäX".³âH…_ø|œa Ï}ÃÇâ”_ÿüDÖ­b¡ú·®ô²?ïóôùÌû¼†Æí·yÐÀJ*ŽÃ >Ï,¬†Í–l@R¡¡¦hX+ýÔÌ€¹ûøÌuÏeŸXÄ½Ý»›«¯4•rO(Q+¢µ¤Ù¬ìzµÌ‹Tg«›¥1…3Zå¢<­’Ù¥-’#jÑËÀ»žb§vTÐ ¢/èš`–J‚jc¹NÛë`ÀRÑ|ØYXLÒÙ™ ŽžQš°M·öë¿„ü5HGQ‚Ôô´Kk¢ýã=QGéÎ¶™ªÛ}š•p•„RÁ[5¥¾0Õ#‘ÕÁ¡¦ÙV§²4 †QvGM©)[aH-¼èö¢å¯î5ˆÆØÝ¶ •k,E›8ÝN™›)?eR›_¥î mÁfÃÃ`HkÞªÞµz;aPœ³ éü&°u^¯òÍ¬þ4ŽÎ¨Ë¥’Öþ½|Õ+©õBµ,AÅ]4êEsäb;›¥¨¦×€B»±IºÌ®\?1í&t½$ê-S­0²|ªmÒ€¬<ÑÚ}–a„ú–Ñ†Ýh—<çk·])3þZ¼i¨Â(3¸ÞTË:5ˆ*[ä‰×ŠÇ«Ò±˜«QP2ÆÉXTd–J‰9‚*ûŸOŸDeÀ¡ˆ ÁŒI¨óì¦Å(ërMÑž±„)‡Ö¼[Íƒ†`X•°êcyOƒ$Ë«Z<­Ñ\{ù(3oñ¸’_kócÚ»¹û§åÂ
<ùÊxšÉCdh„mè´§ò7ôõû¸ˆK1y‘~Å@íµñIï9]]/¢Úù²YnÈ%­<YÂ•ºÒ«Y|ãq}|Å^s.hö™…˜VIÜâÌößì½ÚÞŒû>¾óv«²|X¨êÅÿ?   ÿÿì}ÙrG’à»¾"Y¦m+¨ÂEPl‡A$¨f¯@Íš±ib¢*¤XUYYE ÂàöaöaÍöe·ŸÆvÌÚl_Öll_ù'ó»Ÿ°áqG ‰T+ºETfÆî~ˆ€à2oŽ[–°zïÉz÷?úâ‡ç{‡Gû\ˆe{òôÅÞ‹GûO^ªS²'«Í`¶. ~ÑrŠ5â’KZ‡¸0ôÞî1Ù§žÁ½W6#¦+e ù]@¿Öh†ùó‚À5t> ër<Åú§ :T“¿ÁÇ]ÀÇ!­ÍhK5£“MÍs#—Z	4°¬NFõ¼ê=V‚Y;;yiçåôWÏI_ÙùtÂ¬g¶„2Z‰ Ùæ[!Ý¼ä’NÂµ#>óR~§å¥€ Fæälh¹ß6²›þ“Ln!¥C2÷|]’ÅÒ	d¬KZÞöÒ
ºcÒâX‡ŒÜí»dtÊaàîÝCë³òé \^lxž+IJKTå&[«ÓÈÝ~ÊŒâ)Ó¹Ð¤.±¼svˆtÇbxÑšŒ<6ŠrXãX/üì²Ñ±Zh6{Á1>;^ÍhWe±âh-F:¤=;08Â£U9ùìByù”ª°é–;Z~­ )Gvš¤Ã$Í¬ÙlûéÄ¦Zè;R]ÒÅìÖCÆ%BõÒîm3%PÁ(B˜OoX^…`pI0Í>óÑq9,§—Ê°©ÉNÈÞ¤Ê=˜ïƒîÊ"ùwX4dØƒQ9^>¥¢è¿þfð6X<Wg>fë-™uÚ°e½iõ¬:/êGyStz%ð.ƒ¢é*Sq3Ç‚bäi”«j]É7Z¡&'5{Ç- y
XÚåæ®)ux±…o°	0›ÐEµÌ‹ÆîR–8Ût¿«nbµk’UÙ„#ƒÕÚ1'ÑÄæË-+Ktí"ïj7¾hµ*NÝè„	ý;y¾Óe¿ÆPL°Ý¦TQËË+sd^5·à§¶,1lC
¬`ß¨ÅµvzØÚ6ËËý%Ž;Û™Õ&GÌ•+Ò'OZW›2®èx¾V7fE|ÏKÈ0w‚ýš6k¿ÔáÚþöZÛµ "Áâ÷„Ï‹Qeœ(…¶ú”z#ÖI
—×êø¡ÏBäRûp#« ;Â	®DH-2wÌkvIÕûÝ§gÜ{
= u©Lœ™·:FþÎúÒÍB7ï÷3ÞÉ~?û}Öå–¹~Î Äš­H'+†iÖî·Æ1bî–ãDBÛ|¶Ù>Ð·›LîC¼«–j™}YÈ¾ÊVWV_V=èîÝUpð×YéçøòèrBx§ƒ¢_Õƒ-a¡Ï¤š;dŒW×¬+æòWõ~Þ?ëoˆPC?7(è%kù½-éª:àE„ï±«»‚–dox=W\ánjt)o:àJiZ,rZàÍŠÃO—ŠQQçÃø#í0_1FM¯ÀRû6U‡a£–æÈ¨CXesœp¥¯Ô¢œ$Ñ#ÁÏðªõÌÂÝ7dŒ÷LZmpˆY4f|3M˜þ´GFX—„&±×ÊòŒy~sÁ‡ÿ–ùÖï–éWí™æÓ_LŠ‚ÎS76¤Â]SJU]×Zßšªž°[ÌŽiŸŽ{¬¶%ŠƒÈ¯5ØC.-¿Ø¸{,½^¯ÑNàÞn7½v'Ýf?¯ÅÝµ¾…†3ø{¥ÅlW&Xts‚‡e¸\¤¾"µ=W€,Þ=.NòÙp*Žºâ˜ðµr	^	aGëCW7q†ø	‘†Éœ†Œò"À»­Ñs‡æ\0@å[fA%ÔŠ7"ójÍxåª¨)ä7cäµ›q%VA.¶Hh‚åÔðHÞT1]Ü&èt°t¢— LSM	Z¤“À9BX•‘¾ÌÄ	ì“AK#Â²Ü£, çtUz8Ï¸~ä` n%T™„pE£IWÖwmð•ÍãÙdV…8\7ä¦ô«¸Æš¥^¥é¾™-¦BÏvöÙPu´ÒÚk¬T9 se\Àïjêº†HB¦~¢å®æº?}ükV\€ªnùVê¼£×	›£sçsJSÈ¶ºÔš6:¨7n+µjCLÓoutJÄfQ½Òz+‡Í•Å8àÐñ63Ð<¬î©Ñ9&–—³=á‘ÈÀçÂ”
U@ž’“…&èÂ70v0M÷Éãï2òƒ~©‹\’Âa93–Ç@éWJL¥€Œ€®©¼q$9pÿ	Zf¡7¨[`MxýPT´­]%0±GÁ2ˆïªøµ1¶{f1Mš„Ú™ÙX™õVäOî›­ê'@i)ü¹	|Ðß‚]±Å`Ð€8S	PY!ÉÈÂÆÌ<6@P*Çeëáó<p;»iÇlœÈ’Ð^ÔÁ¯“öá£ÆáæBßPú¢Z°1Ï†Y>eC ezÐ	”5{L+xób‹n3°kSYë_ù…¦øLõïúuª‹²õœ’‚‰lü…žG¿ÙÔ;“ç6.+¢Ö§C$
âßuFÜ”ušÂoÓe±­Ègz(Þ^ûnªÑûhäâÙwÇüsß$îtÃ·†Vƒ~Õ/¯ë½aÝuo±Â·S±K§ÀMRÒQôæÇwQ¸2Æo€Ü]½šú–UÔ‡×ž”5SÚß7-®ã¨—ä\Iµ{ÿxðôh_}ò66'1,4€æñ1j.O¦”AÖéé°xÅM÷¦S@Ðã~¡œ´˜pƒhêÜêºˆEýZyG„Ú±Ô=~ög5¸2Û¬Ô}»Ü ìöQÈ¤›âbú=•YÞsêz£õè­YLž´v°S™Z£ÚMÑßùcQæXoÄß¦Û-·V`…x“®s,ÞëÇ«”CÍ1á PÃ«È„r6Ë¬Ã¬Éa)š†W¾í2<fÝ&{aM«mÙÁ78¸w>XÝpŒa«’	¼o¾¿Ö¯]ÎRò~Þ£FkC:µ¾&sçìTEÞØjj=%hêÌ†Ô˜lW¤iR–“3õ´Ê
iÌ4 œ“Vä”2ó ’dì8+@¯÷:V—¨ÌÁ*„<r«	ÖHÏ°)xF„hÐ öfÈFÂÉ® PÁ=xõµ¦…þN™ ;=Í·ŸLÃzUG`ƒapœ ÊcxÊÙQÝ³S‹æ×.ŸCS›†@å•˜]3¤Îë» ^¤{ïºÖÐ¾obÇêY¤MkF]Ý	õ©¾l¿ü”›Š@>H	À|hàTžº{ƒ,o<R·Ì¸8²ï(u¾ÍuLµle§åW&EÔ´	üôŽ—P‹B­@7³/¯du×ïn´J¬óñ•aLßã®pÔš(äd)ïò 'Üü”@Òó¬¸€é¨A¦&dl»àî[Þ†ßl‘´ÍôE"«8™*ÑÓÅ[Þ‰C¦NÁÇxºòô	Ü¿˜TÔ%™{7]ÐO\ÉÉ¸v¤8RÓIÙ@ªÕ0êÐ*žŽ?þœ™:›ö¥¾™&õ,à:€ÆÕ™¥Æøð’-™Á;ûîžÓÅFl¡V“cj5ü*I¬(›££jÿ¢_»jÆÁ<Q`„¨Á[ÇñéÏVàéH¬€¼abÙèÓÖž?{
!ßö‡ ­g@Î¤4†çf·÷få­f; /-çƒ%mø	ÁÓ¬÷4‹ÜH#1ãfÙÜ2Ä;tëêœÂë"Ð¦÷Åe£üœ;atXYRâ
$‚_†B)¨ÇçÔ˜|›eÈ†ÙB B	I¤J^9L¸ûžöè=rW@ê6ßš,®eL­þE›¿û4úF<¿5)³´Ù0òèÕéöÄ<¯,­1ÓßØÔcŠ/¶3µjÖEf>çø"y‚Mµ
bŒ­.­‰ 5+8†í¡Î‚}åC&ˆ_}ãoÇnßœs·pÙ<’ÛtÛèâ›N_íYÒ5¹i?	B!@l
«÷·JÌ
Ù'kÅaŒæ<¨&Ì05.­\ÒØ9b•ØR{}|¤=m®€<÷y6C*Úht\oAê]Yõº‘õ‚S›aƒR__ecƒàóØd%ÆœZÚ,1Y%‡\s‘€ÁšC¨8ÍÖÅ›Ø&°·AÑÁ8Y”O?”ìÙj‡¸ ²Ôä<a°ÍÊ¦OÈïírq-:…”èBÅ4SxuoÊtºM·¦,¯=±¯-F#˜_¦¬Åk{ÎÌ#„Ëo­¢?ÿ5.DÑïk Á0tOµW¾1¾4ïËÉÄþ Ü‘öX!ˆ†tXLul«‘ZÐK¢Ä–RìR“Þa[È¹°gIê(Ò«kj©°GB8sûÖ}Æ/p8RcéåM÷Ý—WªFKeW O&,q›J2›“0î![]>²NÛ,‚è%©±u/3¹Ê¿ÿ½Á*Øwàø²[9¿p92±"6‰–®3²”·qIï‘3ž€*¦ŠõN¾¸ÎÊÓ1Ü5äMy8¿~‡\K¶;×ð¦kÑ—ÀÉFã¬-`;žn3ÉûÅÒåÒK÷jù«ì(?Î^äÊSv§ùÕ²˜&»øÉ°¸ ˆf	tÓŠBÖ-­Bô[ï–<ˆp·kÃ¬9#4ì|©¡qŠWWV²ó¥“rªEöucù‚™JÙ/–šî3ºš[#&«êõUlîDÏ0Ö Àò`¹·¡íå¡”µà½f¨^+kÎ>ÑW=´Â.UÞÄâóyžŠ_4ÓÚŠ˜€šÕ~ƒÔ"‹W&kªæH×Úú6¯É™«ž®gMù™³Õ‡×f\ç-Èþ±ŠØÞÆÛz!™·“OuYï>Ý|V‚º}hí G.¼Ot[BâTà‰ X[QŠ£8ËuZ—ƒþY"ä¦!h¤mªÇµlxª=®gCíñ]úûV|ð-`}‘þY3>-§CÒàÕ©UÓa/"O-zTÑÖöZ`ûJH9ì¯T={[ÂÇ†½ôË­z««”‡:*uº=XÊº?,ÖüJé7ì³R^õ˜+{úK¸î÷þŽ2½÷v“ÕjÃ Í†÷\ð9ŸT_'³š¬ÌM;Ë¬R¾+à.µ»LÏÀßéWÃY ß%YžÓê¦ýZÚÉ=7¬	ð¾?!çŠ#‚;ý}¯«&6ã&–£ob˜É@Ek÷<°qU‡Æ =0$[$§ŽŒþl]¯™ÐïÔ)›"(I]£ã¥û™‡\:Õ“˜˜8Í	„f$.Ä ˆJ`V{ ËgëÎ <4cé‡Ô˜ÂÉÌ«~_\èb.±vV‘v°~ÑYüqÖLËðV<=/Š1›î‹†-@SŒJXO…œ;SÊV˜ÝÛðÖèÖ	ÜRçœÀæùO€ÛÉ¸œhÙ×Øâ©9fsåƒÅ:¤àQj`DWÜÈåúª“u|Ý7u_45¦¾X‘£ù"apsýü-†‡ø*CÆ#Ÿ,0;§‹•I0‚½”°‰'C‚$ÎÊ¡³~€#•{‡à°ÍÔ9c­]²Ðfâ`çÏ²à_åfzIhÈÕUv^ ;œ»åz—]ûJ{€Ü»žN÷Œ.-ùÌ)Á³ü²šMå¡o^’pTN˜Þp„{TM²{Ý] š9jÝ•EÒ£Ó
1ý(DgGôóg@ÃÓŸOÿÞÑ°À³Ú¹1ˆsÛ Íé'Œ4½/ŒÇ…k¿<‚Ih,YDHä@·*üœ|6lÒßuu¿iïgTÜ••2Tr /»ìdi=ó–é*"dBp¹Sp$¾va ÑuÜB×vëlÍÙðÐ'”öüíR–½æãJ·‘å³5¤é‰Õlˆ©Bˆ%$,¶_Ô£—ïµZú”gÅ_få¤@º1I˜)ì•âÎë|’B„ãPe`']þNcTãtÝÍŠäŠh•3iá/C	ldg¢aN²I¼£òl-îÏ])ºHLkˆ_›Q¦¸“®ðU$S{_
a-Â‹õ.ô`=§ ’GêaŠpEâ\×w vðtyFGæ.¿Ì¹ÀÔ1ø/ºÀ´Ÿã?Ãî|Õó,ž¥Å´gµÝµ6ýÃ¹3Ï)ù·+ë:šQ8_Þ]Ð²_CY
<$k•{z¹fO+2_ Ã”mŸRÊ²xÐ”ÚàL¢wæ‡ùq1ÄfhcÅ„A˜£ûôÊ˜#ÿð¬Y[ugÍ½¼êÏê¦ª—&Dÿª1ÿzB½l„ÁŒÎÀS~Ïìãš·JÐ¾Äv9h’Tlßåý~1™nwzÃæb1ƒ?X6¥TÙsûJ×ã çµLGC×­õ–ÑuƒcûD‘_K”7? °:Ö !ao<®ÎÇ©°ÀFìƒ|sø8ø ëJAÖ“q<aÖõ›AC:¿áeÕö7Ü=±	³º2¹x‹ó¶@Lf8Xèet=¢â‚WECøË&¯aÔš³º í•
l³3'oá=“7ÕpFæ|XœLÉDM+Â-.¯eKÄèX.éspØž•€ß€Qô­amL†dWœ‘É.êmªÈ˜U3¦©×ëõ°âIº›ÄpˆÂ6ÝBPÅCY¢+UÈieNHÈ&d¹Éf¯‘ÍŽñš:¥ˆêÏšMÐ$å´yôöSŽpžƒÉ/° ¶ï#²”u¿4¬Y/ñ]ÀÂã:UñÕvÜXºK†-¹)¾òÈºcèýŽWÙ^Õ„jm±	¡Þ8wŽ*0ô"ÿ'Þl-³ˆÄÔ‚ILg~‰)o“
M©£–k± ì	—‘¾[%'1!ƒ‰OÊu—˜zÿæoÃ‡jlá¹~up¢›FLq¶÷ê:¿ìsi‚…â¯éŒH{dC†@ù¶ª†E>^XpO‰˜œš9D¿Dv—½¿ÔÞòì+ØY¸(×??ŸãÆbþ·²¯tÏÎiÛJâgí*á¡ý³ß\Œh°‘úAÇ*lõìéµv6;úSršáÏŽ¦/—^^ÝìˆŸþÒ7Û+)—†Æ1´-ngç¡sÝÌR¡;`#Ç7Á¯–»kãšF«Ó+¹täY¶»‘Ü“"8¤'·TÒŸ×s[f_Eð	§¿ë‚*Q:2Øøõ°£rÇ?+GçäÉ/N¹}óÏØáŒñÝ>*ªÏr25–æ3áOlmÇO†&
3¸Dª8÷bÛ7m©Í}õ+YhW¶u…^R§f>‚j¥†ËQxÇ´LIÛ®WåA°[£Šy“Au†ôâ	ù@ùPéjcÞû¡t¡_ÒÕ‘ˆÛ.}x0¼È×"t5iöT§õd‹*žB“—†býòjà®Ê…)ƒ=òÔfšñ¼Y¡&hÔÓ¬a:Þ*Õ·ærÝë`›ôbÎ–ð¦ÞÂ¡WkéŒ¥^T#. Ùåéy¢Ö¤q#zÒL…e=ï-Á3íh¦!í!ä…‡x¸…t¸V+¸¦M³ÇHÃ7mÕ²„mÆÙÁS´]ßú½žqYíƒ*¬W³ÌVíÁ/@iÖtå7ºuÄÕ”Oû1Ì°CJdÚ!9žZ9)ó±î¼æ¿KÌO‚æ%=¤®lzDpœ¢Ç‹R?Z9Íž¸^Äqíë&ÝÛ /…{çuwÝ8	Òx)ÜKÍM]u¸ ð] þárÍ¿?ó	o¹Ôc.WŽŒm%Ü`–­‡3DFê	t¥Ü|á@‰¥]¤Šh¡MÇ÷64”Fd)c£ÃÀÃÆÎ7$tªr‹KÅ«‰PÅ’l)%Œ*–¼¡F­¥XÅÒ.ZÉ-,ƒ'˜IlÑ ¬ñA ¡Yç‚åÖZâ°@ÐÖH5/t	Ý©Š;ÌvëB¸F+ãþµ±êÐ0®X²róº‚¹b	qÞÍ+Œ†tM¨Ž.DX&„}Å’'ÖN ¡´X×XÚEª¸ø÷ó	"5>vpþ(ÙóD:eÏƒ±~Ütã]ñ€½HE~ÛÃlÀVÞ±5ÊV5kWE×Ó†ßbgøõ5¡ô `Rê—RóïíŽºïSÁâ•ƒ€	!6KôgL—YÜêÕysvs3¾I•†(ónW2Ò¿³º¿_Âí4¸và¥!¸Ô{Š“Ë¥Í_•þ*ù7Yƒþ{áž—JôØx(å gþïK²–ˆš¶Jº˜aÃ–2d†tmmeeùATº¥yP`µñG­:þ†×wKR«ýôiøMç®$ÍsŸG«D­oÔŽ5„2Ö}Í&×g]²­ÿvVGË=ì¯D&isztÕ[±	&XÌg•E¯–<ƒ{zŽÙëÒr]/¸Ëãè¢ííú¹G4¿ö
åßàùßzåðo°O"®0õç…ú°­!óšfi‹#a@Çé¶ÁÖÊ
»‹ @?Ò§×Ç–EìkKÈªxNa©-œ°ÔZDS-aF´…BK\ËŒ
 ³rpPc)DzY¢×s¤*?û)fºòÛR<¬q@å4–Ì"|!#o²þ>”"‡„>LuM•ô¸}Ÿp›lÎkðÜæFöÞ-=Ï'¯Ê± (˜-.ÏH'B b*˜ý§³‚š{¹l¢Ys	oäN¬âÒªaä·óª¨?þ­xì·Œ‰@?¹ÖÃÂ3ƒÝµ½:¡7þÔ=0h9p¿ÎMÁBÐuÂ•–ÿI9¢^7«jÕLï–²/ã]!U/\¿ó1|Žq³6ƒ}_þ*{^Lsj1Í=z(—¡FN¡àWûÆ†”Ûpìž[½DÙ?‹hYhÍ«¥ÕæXãÎ(©Ã•À^z‹ð¤iµw—Q-¤ 'Xà,v;CÓx7çÙm®˜ Ç™•€[’ˆ“Àó¿ÖøåV@í”;Zƒ¹ö`¥ƒ¢™§,ÈQ5µÜ¿šâB0Íã¦¦A2i2€.M÷¢|1/’Aùè<¾…\¸Ý{„¬+Ð?¬N0A6•Ú°Ðia©M+	öfÖ¥Uj÷}H™ñanÐŸ²àÓJ¸$óƒË<güˆVÆ/'Í~ªbÞ´…•Ôk-Ø°Ââk~2°fØ´ˆE&£56>Æ¿ñNg‡bÇŸuwKŸh-)nX´’é“´ßf:Ÿâ¯”¾€ç?Ì‡âÅ„Ô	àlÅQ.ŽOÙqÖä”Ó|XöƒÿE1>›T@Ï‚ÎJàr<<3°•üe=¾þB+¦ƒ*:kTÖ2yM¶dLûÌŽžùòøGˆèeØbq´ì¤ìK† õˆvPè–P…aÛvš…49!é
+V$ÊÝÞ›‰š8Ôà 1)©3Ò!h2˜…E!uîÊ»FÀ/,A/ü)èƒO$*Mƒ.‡ÅiÌ[E3à€Ï…íÂèÍÊ«“‹ÖÈõéqÞ]Y¤ÿë­¬-¼ÍÔˆô%,pc´,9NpÛ¥×ªÄd–§DðN’ÛÍ$eláÑÖLó¥Ù™f˜@¯žü“Ä-ÒÃ	’¼L©ª$·wr‰¨>öÃ¸z²žØe%êÌf# ¬l§À%¼žÒÖàw'fWÇâÔç@‹y ìYÿ,¯÷¦Ý•g›V¯á`ñû’‹’yßøVÕSPdae^@¥:-ujm+O€¸€„ZLqóR<Íµä7\tÃœs9, ­ï†ÜM1)Œ•geÿ ­F-¿¹£bP&ì†iM„ccqªî`ŒÉ`ê?³89#úAz2u…Ç7­NO‡2JüÞt
ã¾Ð.Z¤LLÚ¦N›Ò¹è”qì‹:ÂsL.©vc§6£ŸCA¦t%'ê¥½±ýË¥"ÿöT,BÃ¾ÖÕ³ÄK$Œ’?…Éâr-q°It.ú®o:Ø™i‹Æò9üË,¯q¸ÔLÑSø(fµ4w#‰T­ö4æªó*¯§e¿œT3j^²7ƒE25Mm8¦½¨çLÃhÚ0šSFÍAMÅÍó›ÕUŠà?nLÅÒU¨I‡k­¿Ë_eûìx–Ñ6J\xïúí=hÍ¨:Œð°o³°é”óU~š7Ù$nB!;ÒŽ˜ÏÖÎMWÝ7â¬Ü9xõuç-¸¢¡Çw¦fÒ‚ê’ÓwTJÂ¬s8ÀŠ¶9£@¢çsÒÙTÎE¤¹Ø•fð .)3!I»9c‰3s”ŒK”ðD&mžb\DÕÞ›ö´;D¥²O*Íyà©ƒéŽÞ@‘ÆAN6I%1·\ôÅÂþŽ&s¬R4e3j»Ó Ø§7aŸTJäDj»sÛ£Štj¯RÄÀÓÈš~äJ?í'¤)ÿ#lexa:cÆÿ¨ºu.Fý3’wüÖŽÖß_Ì,f³Õ›ÒwÏêÖÔ{Ú=?å¦tûñJ7"Ú·L²ç%Øó‘k›2äqK671æ0û'Æ¶ïÄYÄÊÙMí¨};`¸¥:¿ÊÄ%¿ ¥7x(JåÏâVÒvÚÕÍ‰nÈß5‰oEàÛíqÀ'![;µ'í¸º™§zŸ«s_bØ4®³b§Q9Þî¬¶*‘_ín„¸ë7v©Ñk›u\Ã…oáíÄ.£IW¢÷Þv²œ´Rõ ö»c7[™cG½ Kj;½_»ÛéVÉ×|ÄkÒå®’¡1O8vj5ÏÓj7>
Wi„UAQ|Ž	5Jã8Åo‘Õí™T`Ž«Ù”ÚâŒ«q¡û6\Õ|»ÃÉ·ðr6­£n±TJÏs¢¹µŒQ"”PO®®-ëö”k«ÿl8:úÝï2ýy.“EK%z]S‰ö[Å­Òs	üL1t‹¨x¤Ø¸éÁNš	9Í.5ìæBSË±•PÎÇ»…yè2Õlâß!ðÙ*G§AèjêþöÌVRóátûêÝ’/û&$û}¶ê³íA&†»VåQF«ãašú0+áý^'E]2T‘óØ%áYª%ñ*T0ˆ¶|
–,‘RfdÀã‘iÌ?¨iéý89L£ÏÛŽ‡ùø}l†0ðjRŒ	”Œ«”	ÀCð”ã¦˜.­° Ÿyÿýòý•¬šäýrJh	ƒ²%Æw‹·VT0þ:® ¤Žá^ŠˆvùO¤ˆ5¿‰§t*à“XÎf,Aüy»TùÐKG•AœˆÞ’=d–½de±L}íã~ž®™zÂpþK	à+ 	^ÁhÑÎú5Ü=}j¹h¾(øéÅÓƒ»sŒÖ„O„s4b,?sgÉÂ0Hý“„ÐT‰é^Â8RÛh,GáP[gTs˜yz1Œ5¥"™ÁŸ&9q¶‚yfÕæ7àv²{Ï,«˜¹Z~i”©wb	S“¯CÌôÛšþfð^$y¹½‰[ÍÃ­¯iooK”„‚·c¤G#	EÃq*š' 	R”¤]H\’PxhöEù·¿ºdÔt.oØŸ ”Ñ¦ŠäØ°×&Û¬­èÈ%$vçPèþA}˜¥)i.–òá—©å‘’ëÞÛ#Û¥ò}ûâ7Çò)šÞ0£›<ã÷rÅ˜`3‚K	é™Ôz`]È~$Ø˜ÔBÍ\Ðž!$í®µðNWWÌ^%y¨â-Å­B{ºÁÏ§ùñ°@N/Ì­Qq"Õœ¨ó	ÄfkzFt‰¦µÏfW{ì*Nž7¢uhL»E~§C˜¨Î·\^Î¨K‡­åéYjU€c4§8íŠ*ÇmJ‰hQéeŽBŽ½Vcÿßþž“÷5ÖžEÞšWƒK¬¢¶‘3Xm¾;›x€H:°¡ F d™nò—ÁÜ@Tuq‚ìåø§gƒ"™€¯©$/b,Å|‰ñöÂ"8»;Èye#H<ÜGË'ÏZìÀÚÊ3OÂ 	£hŠ~»Ñ²µóÎuSÿ\ózèšÓG×Ü^º"~ºæðÔ•%øêJõÖÕÒ_—.VP– af=&4ºéJÑ‹ñ»êºÕí¦ËaÇ&é®œ¢’¹ý8eÙ»ìÆ®œn‡+,…Å+,µŠ q[¢–Bpk‹]‚ÝJÉÌÙ›Z-nëC¶ü:nð#¶ÁjXš_\ÃÒ„6,ÝÜìõ†¡ˆT'ÐpDIv·7I¤z€„%J´ûËj‚`À„A–é¶D5huó
lÐÊæÛàµÞðJ¼"ñý.yÅ,Ešó‡yéÍÒ‘¡ ÍŠ—K©e”(H7ˆ%Š·Ž©uÄ(H7ˆ)-r¤–Ñ£ 9@B™Ìâ’Š´DiîhTDDª„Þù‚V%Å±¸í\"xÖÊb¶žÐÁQ® Ý8Ò¤¹¢]AŠGÔ*ê¤[ˆ|iÎèWÒvKQ°XUóDÂ‚tãhXZEÄÒZ½IT,HsFÆ‚”¶D-"dAºq”,HsEÊ‚”6¤[Š˜eUuÓ¨YVu,rVúŒ#±µZÑBmð‰k„ˆÃžbÂ0–6ÓhûkÎ	ÑÁæ›’é KsËYJœ–[`Éb–>ÍsÅ9CM‰uæo¶Í&kóLkpþ¸gæŠ})£µŠéâ Aš3¤´ÍR8.¤±Ñh…Éú¯ ÂŒŠÓ˜¡ÑäÍ›ŠÝœvó&W¨„3k~ƒXé!Öhîû¦9Îæ­Ã­Ar'½.*VS1ðN9|»ÁŒ_À©ì&Sž²fÎùœ÷ø]·ïJ¥­¶mo€} •?œäãí«kc5“4yx5Ÿ†>ïLªVÏ¹Žž[Ã‡¥[Òó}õ+ù¶0¾HVL“‚¼óµ«@dèî^•±»péH#Ë‹b +ËK•pÃœ÷ßêj(º½¢Ÿ »D¡ˆn)½8gèå˜  |¸}u%Tô7³•Å\!o¥÷‡ŒüäãrDúj”Z•¥V±"æ…/S â
O£übéÔs…ª5ÄªÅôö¥Ï–ÞüaåÃÙ[-":3Ö"L–ô¡â€‚‘%±xuˆ²Xv_¨ºB—ƒD‘]‚]VDLÔ|ÕÎ.D.EØ:3s’?7ã¹Çdõ=¹=<'¹ÊD#ÒF1=ˆã?ÅÌ:üDÑƒU·àæ©¨ÎŽGåtûŠ1!ì	º²ÂÃÙ}µ9.Ù û1žöé£zæ€#‹¡±i©ù¨‚`PYŠ—Tð’åŒE,Í¾ò pÚ	OCÆèuñÂ#yür,.N‚Š[É×éæUHºe‹üâÅ²Ù¾öÝ'¨Ø	+0p –%æœ{MÞz}(F²û›Ù·ÃrúézöˆàººÊª™X°§c‚§3*­ô„){@1²&¸ÍË Ü	$gÙQ9©˜!!ûÀ¡8ÇqeZ‰‹ÿ”ÒB¤Å-¢Ö‘‹!.²`ìæýú'º8nÜû1)2övŒn©îñÕ°¡	0Ùæm:Ÿ®!pnçs¸ÃþÆŠ‡/áçÚsG›M8çü¾÷g du_NöñAÅmï:Gà˜¾ó‚·ÿÌé„Ù±vIH,$Ë‘¹É¶…X³ýkÂyH§˜äoÇó 	jÇ„:u'ˆÄÇ{‰ËyÅtíò”ð'ˆ¹.+‘íÉºÉY¼¨‹>*0¢UEªØ´ó‡Ñ¢A‚ÃÛFŒ–kÍ·ûÑbgGºT¯áCõBC^9Û"0“2ø•Øa .O¢îèî©ulOæ/kÈ¿‚µÄ”ydA‡EÐ<c O%¨‹Š¨ëÀ5€»fÃiêgÄ3IhI©Ž{šï_uÂ/˜º2À£:åÇ)1½ò¸¢](ô\„Ûÿásã¼ALN ÿZçó2Ýa_fÌs¨Fûí‘C&ÂB(½¹O—k×ÕýLfäåÍ}ä9þi1¶–Á¦âsƒcÒçOˆ©Òéo¬¾Ýÿ±ª?þµ.}ŠR·¿Ór€_<ùeý§žR»àï
<ƒÈÏ)…¦ög¿¸Ì™ÙÔÞ©Ð™[?3©ó«:ÿø/9“9Ï#\F>üÌ ÷[3h+¯Ÿ€JqTsÂSèd•ˆít‹ŒTt÷A(ø†›äÕ04N‚×—Ù_ÈŽ"#üø×ì”Ðvïä2DxhôÂîŠ_Ùröäñ÷MvP4rxÿø×E9h?'Çþò0$zeGãfûÊ0<.¡ÝZegz‡"!zÊì \¾xkæ¼¹FªS€Ý`‚›xÁ¨âïnöÆ[0¬óð
€<Ä„”5™PÓ¯i¤éíÌ»ª\øè<_s£JìÄAÁUR$³« Û6;nÂÜ¬,B6³ÚNÊº™‚L?È]Êu'¶»ú‡7+o©„:Ö{ÚÌIR‹lòÐ6é'Õª·²a1%Ð{1åw3ôú"p]âï¿UUTßÝ±¬Ñ­…„ë—âw(ìEÜŸMÏ}•-C‡ƒ²ìž]r› â×5<}™!Hµòó•—ðo¡n±lwøæN,ÁîžäM,Å,qtT\Ù¬˜;:T¹aK†côžÍZä«ç[{ïÝ•Á€	›¹4J9 +ëcáòáðð(Û£
líì¢¾r.›)®¿þ>U`žÓT`„«ÔìºCYP,šðáž4ÓÙOWbÔÂ¸x¾¸ »)™|?>3Á"sØø¼e°b#*£ëÏv+Ê!|Ú›1Ù,þ·í8Çv\óBÍœBç¹©mf*”+o¦kžcž^Uƒ·ŠÄÃ÷`Èž=.á†&Ï^‘“,¨Ì•?}|÷àË…«Nˆv…Gªƒ7>î2ØÒÄg]M!nPÐF¨{Hgý|H•úªŒ*°O–³ûg‹ÙòßÃ3Ð9^]9óCmØGß§€æši1!z+ˆf*4Ü*|º¨°•û‰ßP¡£6×'®nôü*ÆŸ ›’Œ^Öÿ6ûQ5%h°Ð¯ró2÷!ŸøÖMs°òÛÆ½ÉÆ=ø2#,üâím^?OÁ¯²´¥pqTÃ±@úùÁƒ_ÛœÐ-²<"ñ»¯^·ƒÐÞhß÷ÏŠþûãêÂ¿N4G1Ð÷­åüéÓÝ»Žÿ*¹{ø¨îrûpp°œKh!¹÷‰•ìœ ÏùÏZØvÛ Lª&UlU”J=ŒÑ¿b3«ïš¢Ù“ûôuÎôê.Yæ`Çöh„>j;ŸOëòxVÖšï±l,,íçf”ýœ‚ ÃèWÿöñ;fpQ9,ÜHeê`ßåM	†V¯!÷œÞ)÷
Ws£¢’Më¼ÿ ù¼\G˜š°;PC­ &Ûsø°7¸ü\©%ÖïJÿ»Á«¢Td. Q¸vÐGÎoƒ¼ßU6.š)Õ‡øPàK5"sZF^ÈoBG¿œäC¿ˆød¹‘\¬ò6ÐQÇ›Ð2¡`¶Jú	&à]"za<áÑ“\À²›È’¹œá›c–ô½‡yL6Ê€$:×ÛqIQGÜb$ÞrŠÇ–igÙS\¡RK}X)Á[:îeI[»ä@Þ·ºx,µ[B–ÃI³2wñpÍKÇšØÄn|¬qí;uË] r@ˆÍ3ô¸}”HiÎmEÒ9Ï€Û„üžohoÓÖµ¥i–Ò9U•<ŽRH¥Uw­Fs¥ìó	D'G¿á"Æt=b
NLëqçËqÆY Ä`òŠf$oH-èÊòCÞÈã.}
Æ†4s¹ÏÌAO:H¬§tü"EÃœÇÃ¾@Š»…OóoÖ˜8KsÀ${ä0µÑÒ™˜Åô²:‰`©5hAÒÀË=š
I6ÁRÈ@J[fÉídÚ†ä¸×{ÓîÊBoZ½†Ä#r€è&¶ñÀfdeá¸ŒÆ£ñ¸æh	?-ðÈKø‰A·ôNŒF]#¦²¢ÿi#Mq	)æo3¸€‘ÉØízË¡eÂb¾—TM:B°8fÑØ%Ùûý&ËO`ë“·pi9¸;±ŸrCå žÝO>Šm„7á^È:(@“‘zLW~¥"¾A¶–Ïîû/EæÖÀŠ
ªîHù#Ëžsß,â7r›•«¦Þ¨ÄïT0c(åžøÖB[µâlÛp´­9Ïww’x{2ßýIÄgBZ4¯;¾E1ïQÂ«g¸GŒÌÙnöno6­FÿJ0eµ™}yº‘þO¯ßE* l	½Â·
Y-7é)¤6”Z*dÄdYìJ–ûWœ:®²†•b”oföLÝDR7üåÐ«Poýœ¬á(ýÓG±-Ëÿ†d?$Ë·Q+4ëÕË…ôëÅ²ÆTýÌxvNUÊÛ·LõË^žÐÈaÝ|úñfëÙëƒgÑ™¢Z¼~ünQC§4v»ãŸ)”£Ø?%r²±_ß‘‘Ã©¦ƒì1R4û}¶êŸ‰èt*¹a—`ô“mgoñ7©ˆ,
¾!ÝÛÊbÄùJG;ÅÅw­ûìEU>`sø×\°1–£¡n§ñN¢FOˆl¦d]Ÿµ·Íh#_›±/ò¡ôõ><¿DX’t§Ý®“ôÃ|øœ¤ÃÅ÷Å@]¥ï²(	Aoé~ÿÞàÉÛ~¿µ¬<ì›ß°hê›øÀ_ÊÌ\¯¿øâd6fƒÙ·íR¿/‹sÊo£#àBè‘ö—÷«B=Uã£*o¦ìçwÅ˜Ì†ˆøï¨rŽÊ=ãvÚ¢ì!™¬<Ë§ÍÞdoŽ«é£j|RžÒÌãò/³âõ˜ºøU¥&uuR‹Å/®73ÑËM1
1ˆ7)ÐŽïnR{C;OÞˆheµ×Ï+öe3ëŽ6³†F_Ì¦$3à~¿hêK¸¨ëªîP´õ¡*ß ß¤¾O¡;N¿Œrlr 7ðCÞïYwè\‘`"þŠYôSŸ6’£;9«Æ…êóˆô5?•/h› 9R6Å´½ó>Õ¤‚oÅïoœißµ°¶ùä›Ý"ðÄìx…¥ûxpT‹Ó:½ {Rös
iÛYÞ\ŽûœŒMy–?Rßô/©å+ýn^jÄl˜ï¹U,pÍ †‘YÊ]IÔ¨{¢0<|£¦Êz±|RþCq‰—Ùcß´2¬?¢%eÖu~«7¬N»1+0áÀYÞl
9HÝ³šFëPõ:’¬èCb{yZ_-Z8PiwNzdþF]YË_3W`!Ÿçå4;)¦ý³ng™{™tb‰¬ôÅegQ£Í£bzV‘Eê¼zyx¤]‘œ‘ÝE@uÓ ãWvZ}Dh@‡Ê'4 ÂòM5Ö*¸V?!PÊfö§Ã—/zlÙË“Kãì'§X?ùzæééÝÅçaŠH×ØBë¥®õÖ%³<›mó„6*Æ ¯wpÈõÔx-Ov×îšP}bIÄõ`¾ÔRP“ßª÷ †÷xÙGUº]½ AŠ¹t²ó$'»—pÇÝ³™O±õŒ©-Ðj(è¿¿WçãŒ¡D•Wöô:+†ÄÜ®Ànxg7˜f”Ûdó{ýNÕÄ€?# Ô?#èµ®Ý-ÆÆ×Ù§ýëÊqrd#¦É¨ Çj?]S‡lÞ4AÕgGE=ZðP>ÿHpòá”àùn‡íNQ†j~BµLX)í…^n‹õNlÃóÆ¨dÆ0®VËkýMj5Mq:+ièv½?ÖKoe•YeÏ('TR«ï‘ûþ-WMJéãÉàƒVÙñé•*_6ÏIÃ—“b¼È"ÄÈgc•TÀ^³qœ(/j1cø+£í&ÔBÊì°-?¿àÀ.çž{uáÄ“Ï½ùå›·;Ý7o±š)gFeÚ{|Öø„-/gÏÉReˆ?4@•:²’ÞÌN«×­ºý<Ÿ(÷z}ú£F:§¬Z>ÏäÍŽˆáÃFÜÕ‡";Ê7³aIÚûÐ ¯W“ò† ¨IUÓ—äè¬ú”ÓR¤íÈžx2Zï@u9åïT…ô‘TÅÞWP9ß£M‡/gÃÁÂA?CÆ)mó{ùh6ÚÏë­j!U²:¤3^LíÞ†•^Œª®®AÈÕ!÷ê:¿ìVO yrdÍ‹i˜d&Q`zmƒžÂB?îÛŠ Ñ|¼ á3Ä
sbµÐkÈðº^ÌÞ@}oåÂ°íwÂ\#wùÉãï+‡c$“À qƒØH87#À›òý]vl`ÁÜ¥õ™!:à”àYuNú¥¨ ÷AÓ›VôS¡0énÙ<ªÄ‡Eý–žõ¤ž1øJeãŽÉ•i2Ù¸Ò­ìLGÎ
š7v,“ª;šUè¹™9ˆ“§qü/ï±6¹¦&}Ö73mÒÅªÓW|ÙÉïý“²–Æòob"YE§\håí‰ý%#qˆUÿÆ­CCv5®">&?‚u(Þ–ªÏŠMT4ê¤¶aŠŒ°ä0:gÅp“KNûdQF%ÄrP;…ð·G	=Î/é çIü,©†¨ó¥º }Ç½©8e+ú™†¦ÙÎŒÞIž¸¢Zµ´×]4[öU¶º²²"D£›žìzëƒòää¨þ·Xî…—¼$P´¤÷O½·«àsò<ŸžõN†áùdÅËYzE:÷@þ³vÁ„YÇ…úÒyñê‹Ô—boXÔÓÄµ€9`L¡ˆ-ZÄý‘¢W¹B@ãŒ1²ñÙ`@¨Ã&Íµ³­n¨Æ+Tüö°øPá¶ƒ`‡Óÿ:&BãÓ¡
õD¥Ù$ïAAà|L£OáÅ¨)ÉHN¬ÚÛãS¦SWWUòáÂJú(e•ôé¡¡`W’A?â5ª<ºÎæµ~òÕ‡þuÂÈ÷È!ÔŽŽË)†ZÕ ¯FFË>hãå/äˆùsdÌ<WÒ¨×FÍ×0u­‘_ÃauŽŒ8‡»JmÀìYŽ—=F†Ë2yGkKŠªÊ2úÔ€Ì<b8 ò•/ÅPD k]?yÕÐOÖÇáŒBX"ò×
y|ÿGŠâ)†”5!"Ã\#=&-›žœk•‡Ej|M3±S¤›•JB”}6¤|lÊté“$°FWýî þ#¬ð£—Ïží?:zúòÅaoÿù«ƒýÃ½Ã^í<Úz°wWNºìÇ@f›ªuµÈ\Ó~oºI-´Š2)6št—µv;|Œäî9äž"Öê°¿IHç¥‚{YÎúY^Éâ5ïhœñ…¤?çëŸÉSþÐÃ…ˆE=¢~;…8 c(YÖtûÛYÓÏ¥µŽˆÜbR|Ô®qtŒ†Ž'?¢Ì«UÌiZ€`c4YãeÙêvéÒ¥•pjçœ’¬Þ[XõëÃñ²vš“U½*d\òÄ¢ÆfÉ5dû•6>­Afqš1d²%D‚–.YÉ†w~2ø@ LÊF°±èŒ-|–™¹ËRëL¢A¢tç©ƒˆùR¬¬•UMŸùÁœ'ó
;6Rœ(™àGÌAóæì¸"Ç]ÝÖ V|L`{@v¶Ü™ {Ìî´Ù½Ÿ” ÛPBDâµ«ã7€/aØ?ãbùY¹u[0é^ÕÝU;Þ£Íì$‹6o;OèoéQ>îÃ›ŒGThäE^±ðŽó7Ãc¨±J<MÑÍðrFPwJ#îfê@ÙPÝ¯j†ÕçªœÕþm]*æ­ž•¶ê·Èÿaþ¡Pd¿ØÌŠ¼?íiŸ åtëGÏŸÑ§a[sG#WXÊC¦ÇÅINp…I›à~™“<8‚=á]ÒÕY]“RGô~ÅE¾ÌU5"ŒQ.ŒuÖ"Ú£#ÄThÁ‘õp‡½zs»úµKÑýF³ò“ürXåÊ"â‚G`1‘Åxá4ÙíÀ»Ž&Þ|	¿LÉ)ŸiÑ>a%á¶í„^¾šÅÄûP™«H¬¦b”—NYúÍÍCF:ø{¬Ìq^.Ê.ÁÞbùû<¯™Ÿ½ÅòËñûçù¤±Kˆ÷hú¦ršÒ?b¥ý·‹ª/h«„s#ïÐ6LnÙjFÿè›“Ã|X “B?`¥\þJ¹h–÷òT$7ä¦—fæÖÜ4e»Q¾ÿšoW&¤òù.f’#Æß€Ã¾3— ìfì¼†ÊæñŒ].‚•oH%^âü¼<ªàÜ8EY¢ÈgÍ!:h2s<»½ÒDq0f­{ÈæOÿš  ™AR®8¢Ü^ä VÏ¨çÚ)çdcëó¢::Öxk¶cüˆ´½r”#5¥®Q‡ýx’»‡WùþÏ±VáX«ÕÂç¹ºKgN é®^¬¡ö*ËÞ½%¸ƒ¯Èö-³ÂR÷“`2—É‘?W“BW­ú—W: ^3<¬†¶|?§Uäp¡•ŠÞ;SAê*é¯E—Nc÷;ƒÈR]"xiªÙŽú9kÐœ’÷³ÙÔ·dPbv·ëi:kLÄuú¤¥|È×çÁ‰èe{	3±u¼cMÆÖòñÎÜâSÃêb[äƒ°&H{ Ç7QK}W N.²GˆÁ4ŽóS¦ì‡^ßfœ£ñŠv}Héàå³ýÃÞwû‡G/~xôòùþÁ£§{ÏÔ/7Üq´°w{€Ä®kŒ!«Näpl `pC¿¥îÉŠ><ŸT&Çkå¤Ö×©Zp¤áæWdßØ¬1bö€…v1vîW{ä¬Feý/*ŠRt¿#$‡œå¾úóøÏã—Ã]Ì¾‚6ÙXi{_Ýûóø5¡Yc(ª ZƒüÈfï}yÅgáZyE¦þ$ÿPÕ‹æPæ6˜{w¾¶î®·-;ôômó9–ŒlfgÑèv¾áº)\CZˆ.´¨Ö"cá{ëEÄñÔí¬e i™K„",smÕoGM!óQI	t~¯£U**qª`ýÍ`LúÕ«5$0%†`Z®9lH–ÉÔ€8ÝôXó¿ò¿ò¿"ÓoüÊoüÊoüÊoüÊ¯•_Q"œ_‘âN];ZÓˆæMõg¡oË:á½g·OJR´÷MÃ’_a+ÒjMgì@zÿxðôh_}Œox>Wi@\û7Ì/tüÒŸõû1èBk÷+J­ÂÒâ:/Çƒê¼GtêQ·ó¸hŠ	\]ô‡à Ø1ÑöngAM.à€B°Ûj`hëê¬]Œªå ïé«jnƒ§ØâüjàÙ9¹$?¨¼No2°üÈJÈ9ñ-…4áøBS’1UÓÙêÔÖj};¾·WÌZ'W‡ÝðalÚ„iÁ/­õ•£{'UÁîe|Í	7>QBŒ”À5ÙãðŠ½'üÍ—WÞž\ËðŸLÞNTæÍ®ä†™"¥† ˆ)ƒ 5Û'S±¢V]cÊP·&…¸ÌŒ.ýþ÷.R€òîË+=³6t>?þôA m“ã °£÷ÞE¡ZÓÔm«c¾¸§&`?" ˜gÐºèQ
ìê ºAÍ¶3W½ Ÿ#—ÑT_èI7vÍ"®'h<zñêOì­R	ßzMóh–"Ó‘¼¤„ânQ|íˆÏFe3³viv`Ÿ]AB^zñ(ßÓKÅÿRñoüŽQ|þ–ß+Â7v›hVûˆß#ÒQ3µãû¡v‹g)ù˜õÈë@ºŸ7õ³^®Sç*FÙ:ÏÊñûŒ^P²¼â¾ËÆ®ìT>úle”—o™¸¿«x	ózÜ,%m£Õ¥_g3¦nõcUŽ» ãâzNìi­GÕþE¿vì.f‚è4?H#q¶7ÜÝðt$vƒT[`Î”âÂSpUh.HÅ¼B÷Ô ÏÍnïÍÊ[…¿ïÁK…ÖÙ{Úð“º±þSqIâGb0†Ñˆ‰d¥Ê÷÷9˜
wëêœ"“EÑ¼/.e
íx­æ'çÿPPåí—Ç?|EKA=¾ùø²áGõðxÀÌyåhÔèî{Ú£÷È&©Û|kÊ_,~fW¶IŽ†¤Ñ7âù-åÀéB!M­Í<>Ï+KkÄB·Í]}ÐpåO\V^OfÝùáÕT×ï‘²_–.Ò7Îr¹¥)Äm¤q+!Q;ÏÏÌ¤v˜jOS-ŒV²)‡–vcjzží·ª¡NJ•/­ò]ÐhJ«eœk“hK.ô|Æl£µAÊŒúñáú—‘ç ë§cÊlg
	0ò+ù_w#³¬šN×M:¤ gNCöŽ„Êm^
Ë;†;ý„OÐ‰Àç¤Ó'7F_,D`ÒúÊò æ‡‹žòiŠª†±ä›Q_ÇØþ¸²§[?~Cºn¨Ðb1ZÜÅr€Š‘KÓIÓ›dú ?¿9‰ÊÜ5æ$ßûAã¼y$åþjT¤Œ%¹ !lÌ66¯Ö¢År@ê5“aIŽ‹×¿CÌéÖï’Ë²›µ`ì&]deRy¦¥ÈA‚C]v‚öènÏ„Ò¢° }1g”©¦¡•ÀÙ*ášjJHF%”òªÓÂbfhJªGöYsµêæ[¥Q™Ò	qv5K´µº ÎºË~¼|º·|Ðv&ð”Ö#Ss®ÎqUOt‘éŒžZ¹êg`™•hÊÈä)ŽÖ.ôDåoøÐbi…2):v"„º¥viäùùzŽ3¥x¨JWM™uô„¥Ñä™°TXSšµN²tt†®«ñ‚åYp/š²ÄÚ_(Kn÷]V‚Î•¡ù*_Èå±úe>*Má”ÙRçóÅLW&N^!M	7¹=vÐç
EÝt˜põo­T<”Â¹ËÞÅ86œ¹‘ZöVÿ\Rlu3IcAšzÒŸ…§)ß„/Íûr2±?p6kLÚc…„»“ÐNÛd—Öô¼MOßªºVŒ-¥O,Í:Æô4 Ú¿P#}¤-zÒ;Ë›®öÍ>jÝ£íØ'þi”ÐÖìÃ¼è©Ó×Aº ½‚ÊÄ"X9¿p!š&ï¾¼-iÒ_ö
€éˆÏù"3wïä‹ë¬<%²õÞQØ¼ÖÔHæ¼²àM×)‚_MØe˜ØùÜ¼>n^Ñh@v@'=ÆŽæ 6=’ÐºáVÖ.7YZ÷¹ÕÔÃpà~2×œ˜¦[ßÎÊá€ Äµ¬!aûjíþµéiÔñKéúðÝ:[s<¯~=5Ÿ½@âª
‘d&E’fýËgkV‹7Ž«ô,<©Ú­ð‹i¸³ž@;u	7€->|ü×qYÙNÓ`?Î'*^Îˆl4YZÕý¦ŽjÇ1îÁÕŠTh¸1žŠlï¦‘èS;p™á§	º¤ü(	iÕÉ%BB1ˆæ°ÛŒÌÂX1aÈ¯¼l3_²‘ØaÎ°§åtH‚õ”ÂÊ#ÒÑÆlÒñêú,¿¬fÓï ÛF«¯m‡½¸«ØyŒº˜úÙŒ¶ú9,Ø3êÇ+²`àr¬ÕR¹xY¹HHÐ Ž†‘M×ó˜Ö3›Î[Ž¨ÍØL¨Ëå‹¥wâw	,§PÖIk‚_g 8+îüSgõT±ˆO”ë¿]#×‰¼òbF!G}sÆbMÑª;En4ñ`˜ú­×Pb
 !ñSÎÙ`üQGñÌO7Ü˜Ù».ï÷‹Ét»Ó»6‹ü±³(oîú ÓtÎÊÁ °BìZìƒãY?¸qô{ø£áþn¶°š/öÈÊ ûqu>NYW6Bl]]À6P\¼åÇÙáy	§¿Z‹à†òÂ=Žã^gUõñ…SvÉ”tÉØåî| QÉ1ANMÙlãÈ^1‚aLò ÝC*Š•Û äJóô`e%L«6e~ÚÂ9RÏ¶&Ëé®¯-Ò¥Mƒ¾Öÿï¿ÿ—ÿÌèsÑÅØ`máP¸µ¡ž0ïvi‚„!¸l´sŸöªý×ÿóÿíŸA&ÅÔ~¾…Ó<š~ªË§uñS^D††ÿýŸÿ—´È™ÎÙˆR›¹ÉºÔU`Þ,¸ˆúÊö¢'LHó»f°hzÀ 	&BÄVzv_I‚l–ÖîI<ÝÞá%lh½±”)Cgo0*Ç æ	¼Œ=NŒ¦û <â¦sÞ€rîÒX7ûêœrÀ—ãAyZ)WÏn©Û{ë¨dñ Ìº`R8ð&{©2…_‘[¼J¥Ï lÅ(À‚	cHZ~´±"Ðª3h4m6ÒÅ{X+u1„ÝTØÌ;›??_œ7ÕpFzQÛ	PN«ÉÒêòZÆ–”öä’¾0ÏÀ60H&ÒÏ`P»]§ªCŠÕYaz}A¢„©E5£Úª½^Ï®‡†Rþóì‚Êp¾oÇ/D ¥!Y·lRÎ^@ì‹§dœsû³f.~`·«áô¼“Ëq±4†‹_œùqìpg‘èÀ«Ùh°©7Ð`ÁX¨´`4N:P¡ëý@Äd_XKv×g÷ ”¶Å´Ô]‡Í	¡E	…­D^<¡·Öç‚A.t°á	öTÑ Û|¼Üàõ÷·µÌ¾Ei‡;êwrqC‘p‡<IN ½RŸpGþL.ljöíÐGá€ÎW	!&n"‘°>Í`ÝØßÊ®@üH¦m$bÇç¼Gòô=BîÀ¿ÉE˜Àú'¹wì·Ãþ~þà+ÜÊûþ[`ÃÇkèap~­@{e†]c
ˆ3oÔOÑë9»–³{í‰_8Ãúzã†‡ü|àVégÝbùNäE¬RŸ%ØVq~äÊ‰PÃ ·IƒÜFBnãƒ\4žø¯rŸ<þþV@Vz…NƒUëW¤vì¡vØÔý†¥O>,ËÜü< Ò.Y®P®pù‰)&	I3¬sèšy½ŸOµÇ‡ìX
 þ 9›Þ’,†4µ¥ Î>‡Õ¢œ»^©ÈFk@¡+&:Tˆè&lñÅ3£bPÎ\=%H,ò»s¯ º3IjXiaséüjX®¤kvô‘V1Ô‹D¿€!à*ŸÖ3Xvdb|åC¨™/ð4â‘“¸SÛŒ°ÝÖ¤¼	Yß¨jµØ÷´ÁbTÔùp@·‹›Ÿ,O¤{%¦@Ñ¤«±àéº_{ ”_u¨¶q7î	}~—S+ä´˜X]|óóO­÷é“›[_ªÌ
ÉÖ/ ²¢éO|VQ2™VCÞ÷óO­íâÆ“ûM‹Š;Û7ëÆc½5­“—ædº–7@ÚPf¥Ðÿ@ý7CÔ_°WÓxÑ¨ùó !ˆ$oõ´%ŠÅ®±^ùØ5þ2Ë`]ÑnùáÝL»RÝk?íTª{7óÎºõ02ï´w3ñ<Žà¼Šg8ÇÔ3ÑøÝÌ=ïXèYîß;/´7 y	þÌh$ëjœ½ÊÇÅPS¿¤yüÞÌp]!
ÌœN$N›Ö˜pÑŒ6éïº:ò¯gù¸ÐŸäƒâéØ%9øö™á@#GÛCÍcÚfÁÇ«ñéNÀÙYJž§“uÜµæÚ
 ¤¤Îj]‡`=W$ƒÀÊO)%a.!#z¾”Ï¦•œÎbìªbe«H~¥(–\Õ(Ü]õº<„‹òV{™dL|’<KzcÅb¼2—½ð’8ýÂvðã¢å5hžÐ«dŸb¶2·4³¶/ÂøTÞ×¦R9ä¬Ñ7rš4e>ßŒiV2¦®‰_FO›Ù-rtkÎ„,k“eÑl\	ù?Ô7¥uâK’‚=®mÌùgd)¸oÃlZžwÆMã,ŒfË(QlE–ÄâŒõ#û]+gK?¯À?ëxPeWH¼ÂkP½è/
Lt‡q‹„=ð†þˆ]$f4M ð	A4AO‡,p’c¼¤ÕU{Ã¡@_àñÓ¨¼WýsÙí’:\OB,!TAz²¢ÿˆ„öì´ÍÞ¸^–E‹A§«h™ÌrÅJÇ“Ý³F«wí¹§ï¨ñöÍº»W×ùe¢w…‘ýˆ\@JQÿxfÏß¶ë®óî:xõdìJ)±×´[EìfÜúrI„ìQÎq9y¬$ˆ‡àƒ¿ëIˆªýZ.#ˆ¦´®pSÈw;²î—(ÊYxçÞ5¡¦	¬,j›K†æNÃgü5ó6‰ùí‹${©ã‚¼pì á”„%rÁTàº¢žrŸê§ÅTÊ<©Æ?í…[Ò
ÊkOºOèÍ#ëÆjD-wYÒåK¨F¬:®cR&ó(`³ÿœ{ÑÌ1Ã·SÒÀ$“:Ð8È´ñdpL?€ðªëºå5ÔB˜%¥/%žŽ <±çüHX=¡¢Ý?+úï«ê^dd¸±6Y•ì­ Õ½AÏ{ƒ)}ßÏ~"õ`3\k[%FÏE×üÓCsƒí+µx—!i^>@$‚&F/EÂ	Q°JÞ-"*3²ìƒ?–Pr$RŠ†ñ&2
B0Ý„»*¢N†s½Èû!¸"é–wX„ÓŽ?' |Fþ³Ì™= Ç\‚.Ò•ëÈÙh5ùð¬9À *…¤¶ììÙõl©.OÏ¨©†£DS^ 82ÕÃRËÆÇ^ì		4ÏÔ]ƒ0Âœ/HÑ¸.2ž‰cNöbwÄ æQ[£iÖÔ.ûkÕöb=Ö†÷u!hf‹,y…üÍš—,@BI$vÆ€ÒW‡G¤ËÛõÒ/aH¿Ùk¦y=÷œ¨Õw¶î—¹ºÎyÑ_²9ð&w@XŠªnÀ$ça i>‹>E*1}gëÞ‡ŽÇ«d—\õ~ÉOfi^™‘J1é‘J	gi=¹>?pÎÒ)qb§ Ö‡dèJÙfg¶pUSšJð©Z²0k~£C+·W¨å¹Ë•Õ#Ìˆ³DûI¦ÒÁûŽ¼pãnfØ”þÝ|Š¸-LÓ=;—žGÎë|Bƒ R%ØËUïÓj‰Y`É?ãþ†E_¯­uÝÎà”bjÆ\˜ÐÁê
¡€%„ïâY#âU˜Ô:å	ê–°T×‹¯1Æb¾Y’gi½;Gm,­/âN,Ò•®§//‚.ïÝtÏ©g!mØ%Ë›AŠ#€ˆ¢äØ…bŽ=Dª}W»zõ*|K€‹§Õ¡"zgƒF¼ý„6oÈÍMÂ BØ>6ºR’§”J$9Qü'1àZ!ÀÐ­Td¯eJÆÖ;>ßC“¡z¨'z+×ú¤¨”}BßmEq9™V3µë¸ÎºWºñq~I%×LF¬µ¤€Z˜@#È %µi<óŒëâKÒanEI2q Ž¯ä±£Žº—Çë0çPôÑ†8>üª˜éßÿÛÿø¿ÿöÏ7X+wÎ†EæK f½(b˜#ÜqŒ2ÃÊãÊF·“Ô4ý*;*%œ	iƒl&ã=KR\° ‡ùì;Þ»‹,9×©9rJc:»QÖÿ  ÿÿì}ÛnÜH–à{ET¢¶‘ª’R7ÛåÊ’ì‘%Ù­ÙÒHru<†MeRíÌdÉ”­Öê3,v±h`AÐOy™WýÉ~Á~Âž7Fqc*%ÛUTYI2îqâœ'ÎE„ŒóQŒ)ýjz’%ñ	•x•Ç	³òz.	Bª5øÁËvf±a‚ gú5w"û¡†ê[ñU6zêpi6 ˜Å]å,×êÚ§Ñ¤g±ê‚˜^žÐëEZq[F8wå_.ï‚ÉvêÌäaVÞ9áòvƒ¸)P£gÛßTÍ®dy³‘öòZ¡“¤Ðµ¤¸ Rï“6[ï-´Þ‡ð…˜¬Ð®¬¨G”DÛØA{ãÖw^ê"³¡ãñŽÆ…œìîì`yí%vsüÈqc%Ô«#úÆyY5:;®A¹0üaÉoÈZ™µè–câÄc7ñh†–Û7èÆ@Xü|ÀíÃÒó(Ì
”´ÒÙdÔƒ²-_4œT(dÝÁ Çû‰f6lçê¦#¿T‚ßÎ¦ÉwUH*oN&›røåZxº&"H+³üéXüK%Nõ×¥¶twê¥ÁÀ•YþTK 3„¬tÈøÍ%œ?•UÅi¨2Æ„CØºí¨çþŒV&gõxÃ¼æ ß›Î}‡Ê|5¢‰ÞpáÃîÉ0ú¸p¶°ràìÉ ý 'tjÃt ²Àw;3J¬Žg*=õJòYbnjFÉ4Ÿ*;|*üß_Z|Pó<\Ìòhox…¢BJ; á›œÎè  ûAþ£•Ò6® VPMaç§ºÓžj-žÉtøÇ…}Â¤L+ðd'Š€¥·k¡9%ñ	&EPF/~sËOÎ²ød]«Ð½2ÌÁÕzëÍñ ½wÃaÖ[£4Ç£}CSq–ÙÕY
’§Ô"ŠwZUOZ•»Kc$qbÐÈ1z0¯ßpNölÇ>Ë1%?Ý]¸L!4:šž“!8g
_´Æ_€ù¼^¹ÀËÍH‡.=pzàÚþcE]µ!ŒÑ¥º- s«8yÌÂ”Ë*Šeéè	lXf;Î®‰­¥Žú.üQ ñ‡³óeOÛ<8ß›°À£~”Ù£&iÙ) <ƒÌÈuPâ†·š›uý-jz˜W•µ«›¿’®ßÐ__=´è±†/šå’¡ea¸Êç”/¨Ññ 6 S
ŽO$ó€!k±›ÎskÅYõVdÚ Œ±¦¤Ï«åå)Ï%I]\µ±6ÐK]ó³™
æ„BCxVjavÕ¶ uZV;ÞY[,ÎB‡ÕzÄÐ‘EÅ[n³
ÔP÷PÍ&¬HÔ´³ºY±Ý´‡¡È#¾‹Ã‹âáß£SÃÖ#ÄÿdÂ\	¼ÏÀ¼hæµâ8í_¨Âjø-\þ£ª“f–7²oÄ4+Ç²®iì19lÙeÖmà5y¤½¶Õ©9¨@8uk¯¡"&ef›TD".™©ÉZÑ7¢ b	*S)m31äÆ…¶ˆ2>YÛì¥PãC–™ ²ÔÄ‘¥fæˆ,Me”(ŠÚLƒŠb´`ôiK²ä±sä™|$6,MgùÈÒTö,…ÏŠÛ’%Ðy,dîÄ:’%Q†OÞä’5[ÍÀ„Z"‰dqr®§fz}Î<ÁênÆ1ÕÄÛ¥–Yy¯è3ª›…È#dlf+*7RÇÔJFÔ;,8!$J—9ÛþÐj;¦~Š*ªäaf4,Mm+ÒãfF+AuÞÌ*VïY°íG`­š‘ì}ÝFVq¬H^Å¤Ðb;+Ò·W·0Yù×§Ö\³¡hØæm¸u§³pc[< 2faëF»9k{7GçÚ¼‰îÍÈîÍ5klßD¯faÿ†iV6p˜¦²ƒš‰-&Ò
b]¦·‹ÃÌÖ„ÛÇÑjýÌâž[zîeoÂlÝBÇçcr‚íÞhƒ·ÄÜÌÚ.}6±†Ã4BÌ,ãB€Âv. ¯ÇºÎ[ÃL5­ge…G[Ô·þÕ°*a†7á’ñe<(Âò[hÔ+3°0©[ÕÐ¢=„€šjÚØ3 ‡ó•“ÜÕëæ<³˜ì;0ç)•–ƒhÅíÍ^EEû†ÓçP¬hUa¿R³{fÈÍ»2Ÿn3˜9T]VÖZ‡`¬–Á‹5›‡{7°yÐìEšš:„XÖÔ®©í;´‹å´,ø;ä´ã¹[
Èo¢+g¾áôáßŽ33C0WV™åP]ëŠÑÂÛoUýúyRs k,':ÚMW¼~ûÛÿý6ßžU
p=ÕÄùÔMµ„XÒ¦‰{¬w«‡NÝÊ`ìnÕ×w¦àÛRurÜE¼¨®ª +à0ìÖÍß6K¸iœÂ5»ô—Toö®á–\s÷Å¬ws7q¶IŸ±«¸€¡ñÉ7=ª7ëè`2ûš7Ì3TŠ;5Í·Eªú¦¿6Å©(¿‰JüÝÅïÈóh´±YÉÏIüAq]Ò;„Õ½XÉê–Ò°ˆ±®©©£ÃëXøLÿ¨ï£^Ã{úG}/®äwúë2¸4+oèÕ¼éˆ]þcN©`Ï¬“¥õËê5·rN\WeË<K	NNÚTQ)ÊçÈA<N³Â7ãRÞŸWçÝâ”E÷JO}ƒc‹9y6Iú1ÙŒ²~î‰Nï®Ê§ZcÜ9|&ã=™<$Õë!“WJ7MßÁ?èúã/«šï­•¥†N>»çÐM‘ðœ
w»ñð8TŒÎb8 êùÖÏîZ7†T¬Ì
^„ÛÝ@ˆCGLãë¿ç2°Õ*é'4¦{Ä°=4ƒÇtï—IlR¢ë1LîÔhÀ<el0$Oa|$%'qï,b÷ýˆ0Êº†™Æl./šuø,ï—o@-. €(ÝÍ›IÖÄw¨eø:'¤2´†pºQÄ#S Í) U™† úƒTí=$d£?²ë¿dIÄ 5‚W§)@.-Ø€Eb zË :M´¼'­Z‚¦¼’ýL“öÏ–i"\nãìôúßG°”7‡L99árù~0Ó	FIÞq†Èîe×+’^Ä¡3)hð.C”ã!†…AÆå oµ®b7É’žæª/‰s@ðŠ³>7‡1‚¡eìfÞÓDWÑ®o"Ý½Q÷ð¡ Ïc³ã¡Tž¾þ{–ÐÅÛàn!ŒÆ±¶nu?ó!Aâ*q°]:J ÒZÛ–ó®Ù{¡ mÜš¥4´œˆ³ðk^mÛ…'wH`:PàßÜUc$;Ž¾œ®çÂg|Òª[2Høë0¼Ì]¯s ¬4 ÀñÝ°Œ/âÑÙd‰ÃužÉzZU#âtþ­­!“X½ÚgŸá™éqo‘ÏÊ|‹X¸à¨&«lÆB¦²Š’ô?‘’öÓ­ŸçšVÄýX´óÆ%ñL§IÆS5.4lÅ½4ƒcej5s²—Ç¢EÂìÅšUÐ@àÎí°+ø|‚"\Ö¢]èc³ú²ˆƒ¬F[dff[sÙ0èÌPq€A¦›r•5ô£‹¼RXjÈüdwvÆ
3ñÍ^U¯“6LOü3z^‹FAÁÄ¾á%æ„aYk¡å’^bìoï¿ëlÕ•àõwrØ$£~î³âuÃX0&6P­|G N–nXÌÞÔ”!¾„=Ô‹èE›WÞÁE¾¶=7:Ã<—¨ H©5fŒý;,ÐÊ¥ÝOZŽ^^9ÀÆc@ˆÉeDˆ)È“²ãÃ¬oàÊÀh×WmY+W½‡´°¿¡JJ~S› }Ž†Ö6õq¨©ÆàÊ!Þ¡0m¾ØÿÇn¸ÙMàè¼—1¯[òJÆ  ¤jp²ªømW]Þ£³¥±â;ipjè¢oÃ)©ž9}÷ÖFÂK.WœEJ*OÚHÂ½ÕMíyÎ}€fœS7€æh¦‚i
Ý!ÚR¸“€æš‡ŠºP@^RÞÚå¶©•ª€jôÛÜú˜AÓ`ZžA†©¸…¢¯6­Ê£«Ã›Áúb˜ã!‘C'çâ°0ª og>MžÛÚlÍ\)²äw¨8õ¦59	ÕörŽJ8ôH;è÷¶çm.UÇ€5¯€f+NÅ« Íkàj Ùb
 NþÂ†ºïý¾þXmw†š ……ÏT5'ƒãy”'Jª8ÓËbøÙß(>’
³Œÿ¡!S{‰'îY˜lõ	Þ©´sŸML˜^ñmŒF¦bá(´±¹˜Ó¨–¡Éå¦cP,Üf¬©ÕXÀ'ÛŽ5³»¹ýX64Ý4X‘Ñ`4Ö„íüTø¹‰¶4­"8–ªO+pbM¤‰©‰N&&]/3há™îæËqŸŠ¼P†Ë58ç5³ö(hè°YEsx)Ï/—k¾›SyMW”“…”éî£Q2DúN^Âý8·˜TŸ
ø•Yš	Öùÿûÿü/Y©[6ˆ)L˜æübàN:¸}°3;ÐÕëÐ&õiîÔäÄÌÒþïÿüo¢Êß$œéÎnÖŒ1|5p“îl¾­ïàèý,fh&@·s:J³/Þ‚üÈMoÛÂË6¶oÁäµqÁ4…{eñÚÞÐƒ%c8C d›ép<ˆ‹ ßVAünCóZ&Ìf ¦·Ûv`²®³ÉÆÃjåa±ó0«åhÊvúwíQ·Kx‘¢v×Ú!LÊ€&N{„TÍLMÆYz’âÇ,Ä,ÓF˜ŒÈó(Gç³Ú¦Ö$mQé?ê*¬SuTÇiœÔIrº~)ê¦G)4ˆô‡þm£Ÿô ´£LmÁzc@õIi)³e½8É¢ÊRû6½²¹¦"kËïC1Ðòr/“²Öª¡v¾Ið/A_–ÈŸ‘#òº;Çs5™Œzïñ‚ùþö³tŒ+Y]­dm˜"¢è@ËØLF0kÑ`ýòRDgê’¥y’ãU<üêüx¿Žà8mÓÊ,Ë2ËõñÇ¤hÔ‚Oó”k˜Â/®‡Æè¨UÍâí¼úqéüì5ÑSjøU¬o¦ŒêÒ¼ÖQ/Qw˜ïi}kàŽþ8^çoZèž÷®xAÄ0©„´S]ÄÝGÌ+Ä»²ËN¢A1«]¾h¢†õ–JYÚífªd£Bæ9VÝIÕ‰›Iº†âfC÷’þz‹ß†?…æÙ<œq1Žú0:7ÞHV W˜Ýœ§FFe+F£2µ.«zri[œq‰(VÉ·Ç,½FïGžXÉk‹´ykçÜªG´¿¨§`çšúñI4?Gƒ	 ¾M{Ôf²ø—I’Å}kC4Ž(b»Pô=,šù+ºá£?átRPç«®c‹…³sEøà€ê<·¨"4-¸Õ‹Æ@Óã3pœ­·––:KìÿEøiÁåœõK„{²:=¾§…³òÒ‘˜Æ.Vmÿ®aÖ£©ì†Zfgí[MUxZø­»ªtx§¼!TJöo&àiŸúµtŒ29§CnµQKtxw:µEö5¸¸"—~Tþn\¦Åýž¤ÞtóIñå#ù³q%ºÐî‘&¡òU¶¶È ôË Ò{´Íì”uz;RsT=5EÑüƒÿJw$·jµÔû#ü·qQædýýÓ¸ð“,ý9n=bUû£Œõt‹,VéxjBà_Xc·¶?vÉvÂ·ý<ù9ÊâwŽæÝ,×—vÔ"‰Ä1‘Gf°Sqs(UœµN¦Aþ^¿g'Ûè³ôA‚úDµžy2¹!Æãuõ+(„‚Â÷Löa„½dH[¸E‹§]QÞJŒ³Ø{wX¬áå\È:Ñ,k]Æx_‘—úí¶!¶Ÿjï õàæ¤ÎPÕ—Ã_"ˆÎŠ×ûüDeÛÔ§ñõ¿ÝæDøMžšJ¿Ë¿*˜ú4ëý„ú¾ÅÅfŽ©§]jîÖúëBß\ôF5=nóæ†60µ¤ÍãüK\è_¯D`7½'£ýæN{­R°"ï±‰iJ””RÅ8ï..ž¦içt°8„’‹Nç×%¦ú„ œõa4ˆ©Ÿ¯ÿ†¶µ·Iypõi{7ZACøÉ±Ì¯pÜØÌRÙ­€7ËÎ5Tá&õ›é0Î¨_ÝEòtëgËØPUÅÉÞxî¾{(v]!ýtË®î8œÞÔŽfŠïn¥‚êÁ«Ý,,/¡ñØm_MÉU¢ÝàÆÊq\e—YŠ«…œß7SÏpè…áœc»m„ÿFêräÅÓ4cKƒ Ì?€×=€
3øïL’þ•€öäVÇ†LˆÛ®<¶³—>Ø	ÓQö†tò8;Oú0Þ¼-¹Ýšïs"}‰n7£<n{#],þe²;pß
Ö‘Atœ§ƒ	ì ôÉÇö"/,/®j-@7Ê}¼Dê¡C·ßEÍŸên{Í¬¾{1I…B÷_ãEÀ±¥ÇÌ™ÁöD<Œ\þ…ÎKÌçEé®hÌUhÍŽ…îz•‘ßAZ±NÑËÒf@8iN.'£ä—I,êRœ¥:=¢„8»U|ÖÐáº´=¹»Tn 8EEß›ˆÉs²• ”ñúoçq’Û|A¹ì¢]~6láÐ¸‡é ¯> Uí`ŸÒÐ¶1ìZEÏÍµ3k0âr5>f]cÌšíØšÝÔ@NîÐWnî»ëÅ½÷ÇéG7B¦¹âþºÏL02&5Â®³>7Þo$¢Æ„ž%SÄGäóÉ%ù+Ñû¶>0¦g‘¶¶^½ö×ïžÇ`ŠÂé1ÌÄÅ} ÞðþcÜ09­Ò|6"¤á(E/Rm3:¼r3î†<‡Ê:
æNŸè“
™;Õ$ßdÀGÖë¨Ã`Â!ŒÖàÇÓ»‘ŽÈ@Ü»T@È?Šõ4ôeÊ[ExÜóTÍÞäG—ê5&7àaCLM0"åŠúíQ¬l\ËÀf#Û
(H€)¨AwaT?²LÈ7‰c-aþkÜ¼y™ºäU§Ó©vt^4ö:ÄØ„8qÆìˆ¬ˆïEb
+6k˜§&VsØáÚ,qY½V.Úb½kŽ]²ˆÆÓ)¶‚…ÎQ-+..=lÏÚ.§6`¦•ÃŽ{ÈêrS:Mc²@€«*CZr~z¤ÎôHèô€›eNÅ/šq%ÔRÏ`ÓxÎ1¦'b‹ˆ¹v§†Ü½bîh2jtöé¯JYõK%«„§µE£å¯Ì€|õ»ßLFTèB6£JrâáCÕõ¢M¨ÛŒïcˆõUÁy3Iç)a¿_¦äÔuFŒŽrajµ[Ì6/5œŠd<`Åžó{9Yº™ñ@L³>‡i[¥,×9ïd‚“o?¤QÑæ=íd1•*·[ó­yÒê´ææPó|é'µÑ%½Ñk{´ŽÅE‚ÐÕ.
ä1ÆSJž`P‡ïÈóëÿä5Ã‰ï€‰Q9'"{ý]Ù>vêj^™eÑ…×ê¬0•8<f°¦ègþ}Ì&è0Y™ÞºSßƒÎšá)--¼–ÏKªšhéÜêñz¥ã5ñ¾=9Øå¯¯æ:¬Ø6]jæä\c³jN£KógÃC>”˜Æè8[c˜.4ÑÁÚÙJ-ûj¨™òæõ_(È³8xìuSŠµÅ³¥±ºHÆ¿Ñ*†Sv‘”9|†?Ñq”|L™Ó”ÝiTp“©Í’£ä(KIŽáƒâ•Öç"Ò‘8=/ZÓØc!ÕÉÉC¹"ºc÷iîc«®ám‚‹ºØhZ!”¸iÛ·†ŒAƒ{ŽšlIƒÞƒz"-rð­Ê˜Y+»Ï¨96l"Óâ2†á)ÿb¼zaÙC®[tóÒyó©Æ|³"ocV›b6¼‘Ñù
uYj½	XÀïNö‰l©Ý4ï3(¹ ‰ƒ‘ ru@0A‘ –~0ªÍ ÷ÅZ.ûíñç?‹F@©½pà` ÷íÜ¿?ê®äÁUƒ$@áÉ´ÚÁn× È(|§Ù•XD
 ÛâgCâ©t–E£ÛÎ`ûÕë3®42&÷—é3Ì}Éu »C.uÆ­-¹ÛèšƒYÖ™15‹K×Æ°Ð¦«•§×7ò¨×ÿÝ»ÛÇaÀg|ôò-ÎhÎË•í…ÖõAñHµ¨¡ÕYM*aÌ‘Eƒ‚WBœhS Í&óú¯„†Î£PG£YbøF¼H,Íh¸K
&ý4ï}•‰r`O3øsžœâV£ú?ù¤¼cb&:²‚j¸÷•úO’Q¤Þ3Õ™LígåŒH]pÑc!ž¸.<Àñ,ËñªùÐÜa}‡Ùñ˜þ –ÿ †Á®ÄŸ2.
Í—·’ÁdÄ=^lö$…±àïgQ‘Gãñó8Ï£SV)ó)öGþå(Ž¿ÑO@76ì_Ehyá{ˆ÷»ãÌí¬î‰þXú¼Ú'ø4Œ‹ˆ >¿Ø•>ýrº›œÒô´¸|¯Íát÷ê5Û(\äI~Ø;‹‡1Ÿ`t±³¡}`ï·âA\¾ v–~Ø½‹{õt4ÿ»«.Q–²Kð"}Ÿ=ÿ
‡ÊÁà'±¼ÚÇWTÖHW»KÒ^É€½ä+ß%ía—°“ý<)Ã2Ÿô0DOjnÅY–f-JÕÏÓ¤ÿ“€¨þ°š(¼t)¦ÚÙ{ÀÓ%Ï¢±òŠR—ì³GQþž}‘pÕ%›ü'¯›ÂTNÿnFãEì›€º.íé/ý»€DhýÒ?W³K(ØmÈ7,›^»@¬z€×ÄÔqáÈON8ÆRPao“£¢w¶F÷1{wSX7è&^3W»¶QÛÝª{:1|ƒàòŸ?Õ ø± YHl9NÚSÞú»”I`NQ»Ž£7Ó~ü¸úí—Ìòúe6¨½eó¯Ó@oÄÞ_‰yÐ7yƒùåÕ«£«ŽP¥}¤¾Ñ:Flµ}äbô˜0Q¢7¡k “pÖÕÆÄô‡pÁ/×¥}ÜòÏ¶èvßß;<jýDÐ%åc‘\a1ºÿ¥“<^ƒ·´5Ž>»ä9ûÁ@‡aRöòðyùŽáTöžþfŸJÛ%ÿÌ~‹ýÙã%hn¦~ÓÁ±WÓv‰ŽGŠ«ã^˜™œÿ€‚þ
×ô´Á"Ò{ô“_CmI_L°¹`×€ÔŸ|ÑA%Wšt·‚¢ãyÂ‰Ž5É-]úÖY’i–ôRºÖmÓ_€ªéß“$Î÷‘•KÊo€NÏ…C‘Gm¥Ux,=…*1vê #ø˜ë(cC¹ËG×åN7Êg†Îå3¥oòIàõ²zàô'åc‘Å	jJgðŽÃ²|.ÁRéÕà4*éiY9/z“¯ŽªÈ¸Ç\FlŒ&BE™|-3Û6Ú#Û1lÖ~VŸÅ¸B«…7râðZCÎ²¶LÈÄ€ùó"gK¥¾Ñ–ë0$½„Ì§8öW¯	Õ^Û ¥\õF­Ò]Ãµnî¶Q©Š¡Ö§ô*—Ý[(/ |„„¡,4›>lÀD[N µOò¨/ —a2úÅEr89Î{YrÇÒ”°“=’šcw1¢½ä£oµ}r[X»Û@œÒI.^(8¡KùýƒeRT/±l&~±bàk¦cÛîÃ†ÚÜÛÝÝÞ<ÚÙ{qØ9ÜÛÝÙÜ9ÚØÜØ{óto÷Ù†ˆvËJOà("†µŽøoó³´·ã¿ÂÖÎù—Ê¥=«U}Q™ƒçéôÓ×_LáW¥‹ˆázcWÞv:X®ƒ¾|Ûså‡«¹9<Õ¡K‰¿ŠË“Â©8“,/Ì9¾ 2ª’Ì‡ýìä³ÝŽæÉqMNrñ'*ã2=–‘qËk¦j‘'Pä8¤¿NaeXs
­VãW7\G \UÛÀ¤Ü/´J¹Š”=7,!`‰ýÖnmãÂ·*¡€zÞmÍVÁ”#÷IL¯¢Áöœ¸+“{D»#c.H·â^ÒyK¸LùÅ¨Çù¾»5p™çì(cÈ,û^².lrà"§‰ÆqŽ&ÅY‡Ýœxh*÷¬
J†—	ê8U‹¡»^¤O‡”1´•¡ŠÊ¨/Tå©Î;EázÅsÆ*PªÕð'¹LÅª/_wë}£î=¨/º -	Œâò–NðA|‚­¤=?Â™e:JŸ£QØ‡†gØ‚ZÊjçÈdËU¢‚ˆ¯Ú>Úwt+óaÌ†ŽY»êÜ”¹Xã°Q‹³s8]Â>„‡ãû({‘Kµ¡·Ä¸4&&ß^Š‘æüÂ“¯ýö’‡»=—„˜÷˜9NÎ8F­«ÎÛ²»ò”-^ñ¹¼.xBºµyHòKeÅj›Îîl±UØ¦–+#ÆI³“(x#V4‡±vðÂœôËÓÕ•²Q_%Èb1hi¿2ÏÜ0kïÜô=É_Ž9z÷“ˆ× ¿4VQÁÐ+TbŠX0Ôìbû€híGÏwwð®c{#ö¨¦‘@…'ë¥³|Îw^-qý,Š"ð¥Žð5¾íà	#ùÀl|DH°Çý“V¹drövÉItžRæF1.ˆ²_&ÉyJö·žÖ—h_IN¢2ó*¡â86LèÐ»|79Æù¢»4¢üv‹~Xè*Û™;Ïéq4øcš½³=vdì| O‡Yªz+ÿzýÑ»¼Ó¤“þ	ð´q§—£wÑÇÅArœãttÞå‹ß^ÊšagæPÝýÄªì bêßåo+È)Ê²èâÉØ®LöžN½ò¡]Á›P«Ì+Û„å¬4A8h_Ò]­ò«¹Î˜ÛDepÔ£2d¨«UâL`°I?%ð~ù'ø³¶Ž­tF“á>ÊàÕ÷ßÏÕø©1|S{…ÂüíD!½œÚ¤€dFE™²aö#èÌ&ûÖ®ÂLG¬·¼|‡
ßÏ†?9*-€O8°ÎuÞ¥É¨Ý"-¥::äïayÿu´°°@ö¯ÿr
g%@uÉÿ:zK¾/[ûž´þ)
G®úâqlNûL{&%Q"‹’—°oÔÉ¤¾Œ£øƒ,YV¯ò{ù¡]ùò;]Cô»±Ìøq5ÍuýŸ}¹“¶»0F
ZÌÖ²Ñ±j#Á€­LNª$ &êùdïèÍæÞ‹§;Ï`o%µæÌ¢¥Ê”t•	Pyê"§ð‡„
y »u	âˆòC¶ IÚ ó«˜#B =Ö™”]Ú…1Ü¿i5%Z&Êô¶N™ /¸ »3d²Ø«·5E/Iƒß0£Î¶Öüú¨«Ñ;<Ýïgñ9œýøV>4â¸þƒ	ç;PM2Â²â˜…r¤f²ˆÙªTßÙˆ¢ý@jêÀœ:KÊhÊÌJ^1ÁšG¥ŽFå€É7µkà÷ñ5!_«’’³Ë§”æjCúFRÜ½OÌÿ˜gíV2ÐXDU?ÁGÐcà"ãœ;9Á¥r´ø°	ÝÕVÆMwaŸ 8>ß2äß>á#ëJ!—éëÇ7E!çkã%t:8‘õF)¢åœ'UvèÀ²ØÔ}¡’Â3uðÏF¾”ðåÁnûÄ”%q6«#GeÐîçæÃSÞ&ªÞÁÏJUF>aÑ#Ñc%ÕüìÍ¸v‹ehUêÅŸoüéÍw¶Žþ 5Ü[Z²gúÃöÎ³?™s!mÿôÑÂ‹vœþ6e:‹1þ(ÏÅ~ªªv#à²Êñü&ë,5—ƒÙŒ‹7úÝº2ØEbì#&1™¹ž©ªÓ`µëÂ^òæ)³hî'k˜w“O÷"³d—š? «Õ¦ÐÑƒ6Î
Ï#Û3w‰ƒcñÙ)Vˆ§ÆíÖJ¿‚8=PÀ4ðºÓÏ¢ls ¸ÌcÈ"øöp^@F}À\€[4”])R±k9úNƒ-uÞ7TÂ1G›WSËa]pQî}N7©{ÙÉ @ù¿ò¨Õ§=_™PEµ­dLtÅšÆ”ùszÜp÷B-ueÄdam^•®ÊI¯BÜ¨œñ‚§¤K¯ÒØ²ôÜm2‘’êvÕßÿˆíÉ8¹b’?#ßx{¼"ÒôÅš¬Ñi¹Ež«Üb%‡
”E{Á~k¬Ù%QáT6ã$°ƒ´ÑŸ¶x]Ü8EÖ·]>kì¥ÐØ¦äºÊNq]
Ñ/þ¨wŽ~±›”›Ñî1v EÓ.öNþÇïázgõ£…í¿_UÛg—q²}öhh_NãêFxºØ¡¤¸¥å‡“dhxË²?Š§ÉHË[}%3nô¢~5§ú®:¦a"†Áž£(’b2P:†7uO´ðB=ŽÓˆtCÞ+ö“5»$¶¼8-ÒÔ×LÒ,¼à¾à>UçBèÓˆÙÏÞUE2³CW€ž`ÉfDï';‡{‡WÃS>$€0ŽZs¯–^—åž&Ãf…ø<þËÑJš¢ù†®OÜ;c¾üõ•z*»åÎênÄgm7Ö•®ê²‡"H<|Fyþ!Íú²ÒMÃ‡×\¶Ñ¨Ür¼êÃKn<õ¥¶Pº™9aº•®ÈS§þÚ-Ðó\T
MkØWß¸NÉÕñùW6à¶úF«J×ÓÞ¦6¯[èmý]mªÚd¾ŽsýµnþJë:çZXµ¦òÑ°ë˜³Éf‰"n_˜ Zà›ã¢èÞÊMdøE‡ñ0«²|‘ÕY¾G0ÆStÕ¢ÖU}%+©~ˆ¢£¸'UKW_ÉÒÕ0+iþ,‹ú“ˆ’e^LïË™©5â©Ò¤‡|©‡|ëjXw½âÊ­RG«¼r?ÁuŒF‘®÷G1êØ»Ì}õÈŒêË*^q“…K
«â~_E³;Ê‹›È-¶Pýê&·+bŠ¸Ð†ï¶˜ËC!™0Ü‘D\@Q²Ór‹ñF©"ƒr)©í'fµÙVr«ßÖº²ãêå´JAã¦¬7d›S;]Ýƒ¦ŽëyôÎûË×2™á¯Æ’UŒÄ¦Qðz÷%Ê¯¦;
Vóh]”ÇÔEþ‘*”èýt+¿VËÕ±œ£³–:ªÑØJCß}T²XGá«Ç˜Q¯­<ÆS_ImVe5UõåÄl:?ß>Úx³µ³1Oô*Q#B=ms´R¾RdòÜŒeÉV‚Ö>QyvŽŒggL5&CÅ:llQ¿sh–‰AÍÍk2…N§S)U±ÜrïxWkæñÊKE#jÐt0'ÄÁ¼Îˆø5qqE¬JÃ«u|õqH¬ú*çÃë5½¶rK¬¦*Äk2½¶rN|M\’˜GÛ7¥‚CRT“åÌ¹‡
\~™Ðïtd3JVì.ÆV©<ì\Î—Âii¹äû2«äÀ´Œü­+ÊÜ7À‹ÿüìÍîÎ³Í½íC‰Ë^LƒÿùÁâÌ[‹O²ˆÉÀeÏ?*ƒ»m©Zñ8„#ËWa¬Õ2Õ«× *óÞ`;j™r?2³Ã–¬^(—Å’Q?ýÐáæ•íÖ`Ôw@\ÇÑ)t*F{Ï,>M 0UÑzùo“Õë_u3XéÓ!¸x7À|]ï@´M;Ô\L^›m>°YÌ6Åx$+'zRÎô<§·ÂtFX=²hÚVqˆiÎ^nÀlME([5ÈµCW¨ŽgeÂ_æfL]"eûè?M±Á¿Ê7gëW@™·Vh9²ã½Á$ÉÐâ¶4ÊøQÚ3PŸðÑ?FKgØL¿u„ÿŒS4KŽQ4?‰Sè”ê¨ôo9`uIÃ6‚²¨&$'…vþúoú>è(èÍ¾
„OêÓ$£Ö01•=ªd
Ê¨Hyfˆêi¡*ÉÖöîöÑ¶úýí·—µ!\-~{‰nÊßºI•¼™âk!¦Ý¤^råæ?häˆÙ±
•Çš]Þ´Ž»;/þ©¤çX%å„_ˆ™æ‚À¸ÙÔí†¹fDœiCú$uÚ¾©®† Ÿ´WA—y·ƒ `ý™ûv €ÕÝžì=×¼¬–N_SŒ]^¥)ïïMë•‘-õhl'ÖêŒ Š8pxÓBoÄ#ÊiñC4Ÿ˜r Õ;LLSÜcŠbî2é"LsŸ‰©É¦:sû´Ua]ùüúm#ä®×F·³oEíMwîÓ½ƒíÍÃ#¹uEE¶ÍûTñå!Ò¶¯œïþ9g´ƒÅÀƒyƒÍ(ÙË×N”CtžÌ¦¼Ceß£c2Ü¥cªß§‹Ìµ;uur¦Øo,Ã7Wu™á~…­ýÝ/¶ª‡EG¾*Ã‘EŒÍˆÞ5M–,>˜êZ>\\š>7\QÞ¿&Êt‡>Å‚nììmí½ÙÜØ§¦Ž¦•åzN7^Ù8»þŽÒµ´¬-'š°*Pa2+Q•ÅªZS¼ŒéµE¡ª,Q}ás
3O¨g‡Š^y©7«ˆY¥s+±äwr…i¬o^é$“tE›PÝ`Ñ5–ð¤B]T—X˜Ln±0=\I•V®‰:H¹ºýøÙ~õ‡{/:Lxœ\PEßy¦5@Væ^—pŒÆ]ÝVñ]žŽ¤«`µ‘	U8~y°ËUð÷ŽQ[uŽ±y-käPØÄêE³ŒZCÅâM?ý ¬Þ2o¢Ñi½¡KõæÛË0vÕÁQ¼µö`pï&(™e'§±EŠàY:A¦ÑoV¬®©ç;éŸWó ŒA$ìÁr Ú{Rç}˜ê8ÊÐ½¯ê£öò•Â®\R·¥›8Ÿ¢ÃOx%$+øý…ä]òŠY“ç0q†»k÷úoÐ>†´Â'îL¤gFßeTˆTFlËŸs­×*_.{£9B){ô”½f½"m@M“ašÏ}‚þ1çJÏè“§×ÞíV§l¾æ´ÅÜIî§ì¤8ëQ…4*¾û3Ç}Ë¨s'_}ò9|ÊÞñ«Àèî§Šûò©t…¾ù=án„*½‘oï¸GŠc¥²GÚÅÅÝ÷I¸—*;$¸ï>,½$ûâ>®”N	Æ±ÿi1ô¯Uv«Æûá@ñî;¥:õ*ûuT¾Eï‡wß-æ™¬ì
—srý?Š8ù,Hõb¦tOØð ”í¢1Ñ'¡3Ò9šFjàíís5¡óVóíVvu³üDÄ·OÞ_îOa)G<ögÑèúßì8N«òµˆ·Uà‰ë›‚z¤£>±øÏ2â›î´Ž:y˜›c.)Šè¸7Ø´ˆ†qƒÌ¤Š,m"§/¢Ž©>Û,?šâ¡S/Wrê—oñÇëþË©ÛtÆÁÒY4®º„ÿöRkB8îDpèn‚u§’]8µ,Q“¸öaßú¨KÊ°"JÀ ¬PFXòùÆ×k½z«N˜îTœ®… 5O=6ÒœÜHõ,~©OLyv!¿ÿ½
k9“Ò„OÕj(Åyë¸^=×½›ä²ŽãâC´HUïüg«u¿äo0Ñ}ßCGyœÓë_<[­´húUß(ªÞ=saP3¥6P{™¶Î£{ÉyvE`Ì+qjø©NÊ`’»‚L³˜s/ÒóTN¯) €)ìWÝ³~Ž*‡ÿ*\ÀÒÝÝóxtpc¦q¡8‹£¾!,L‘U–Jnu=ê‘!ònƒ ´ÚFù;ÿ^ëzC[[,Îš•ÚFŸtÍ‹mî?m^è(†9MGSô’SÔ)úe§ióbÌãiórìÜd.o³zP#@­èöYm `bá‚ð¶H$˜.™§C>¶FÛË62Ã[®=#šC3*d»7²…úlÕIž‚vzé ÍòÖ¼¥ð¤CñÄ}¤E¸—²4Ç*ï/YÊC"YB°}ËâYzc”*±âª#lvµ–ÉF˜9­Äk©Å¡©Þ¨ ü±BÌ>æÎÞ°¾xëôÎ¢l£h/Y£p:ãpÒÏÕ4ÁTÕÐW/\‘H}„f@×ô‹$ÛéWWé;uìiîü´ßŸPž~Á.óÎ:1>C?Š³í†}S²ˆ¾Vxá² ñ /ôÈc‡¯jÜ,¬V•sÛWêšõÄûÑ­8>³¡@LŽ`ãöâï”;’sÒ4– Š‘ÇjÕØ[K•xg„qE¶àÂÜÀðÅh‹Î&Ãhm‘}h	ÝJËª-òãWFWî(Æ˜ÃU¯·ÇsŽ@ÆVg…Ÿn¾'Pð	·6_Ý(!ÀÒÁ±'šo†KvqË†›·öv·¹&k×™ëÌÕÙ¯ ¨î¬¬ƒÅ„-­`KË:_EWõÕòÒøãkÓ‰ÌÊ¢b’lªs—¬äñèÓ;±ÄJñ0Î¢A¿,,^`y{÷¬+çBº’ï†N?¤€©mtê^$ç©•~º¸³YÂB0'n
è'k±d©"§qc3‹?8xýä+Wõ#¼Ã8i®]týLSºra Mâ&d:ï/¨œ¢<ÀÑrÛN+$I1€ÊYœ÷Œýìù]ÈlíŸâ‹ì‚%=0ˆ’”ì– ñJŽ¦`P±Mš8ÎL–™•ñˆ+‡c5jã4Ó‹Ýb['‰#¿sz±Ž•O5·N˜·2„c«.ù¦Ä\ÁÅ"4 %HLak}Á\4#”jPº¡¡{mO–/Ý$„‘JOY‰xåì°Êüétc+ÎéÈ(åx"ìy$eåîeY{9¢ñ›ƒ€íéº¾
w›Tç^ÌUÙ¦hŒ’¸`M
þæføˆ›Ý	qþ”(é(‹ò³á¤)$D&q­‰%†|(¬­ÞVÐ»wÈéKÅÛ~íFóW¹ä#l‚dtŽ#’xü¸„rQ‚·Êý,£ü0CÁÍŸq½bÃ±ÝZ¦¸Øh¿6üd” ±ñúå%I×JŠjð“÷"<Ó-u~¼oâ8¢ðµRË²Ô²©ˆé¾‘ßÎ£ôJÖpÿˆA«•º‚ù[½}ämÜð’%Òa¶zÉß;ð‹›Í¡¦ï" åæo uŠiB–Íù's¤iuÈ¶oÙók¨Öb¨(NŽ‡I!ä¾ßfêPM<°´µÅÔß©'xþØŽE`¨#*qÖÁÊX¤”Ñ…æ°DO#jU"ÚDOÙí¾kÍ•šmÄžŠo«…éËÒT^-M_†”îOªeáUHIT±	%Ÿj%ú× žœEçñ~ò±Öþ>¤Ž¨òi«uˆ÷!u§ó²U­D~©E
õ:øë˜¬N/Žïüe¯Ì f”%@—14©<¬³°÷|ïhïàÍÁË³?{:¬²<*õâÇ‚z'Åµ£æ™‘[íÔ¡Ðjvþ({aÜ½Æƒw…ÈC†{u„hDuv†jiD“²ÂuqgEax¼°l•• F*»ËL-¼mØÒ©c~9PÕ:CN6Ð‹&ÒzNê3ôû3“„ª«ãºI¬ëzÕg,êYf­‘,äÉ²hw½˜TÑãV‘¡y2¼û²L½ÐŽk×ßòJoî?%í½15þ˜Å!7]r¤Î—KÜR[%Á7OíZô(DþQ#¯|š³‰ƒl} %ŸÁ”ÇÙzkii©Ãÿ_°Kë»F9Û~LJë¶¯ÂÑ9‘’çæPª·àØßîÊU>Ô€Q‘ÞÎr³óGã5¤ÅnËìcí7Å3·	:,Ö^Zš#KK‰ÙÐØnýˆìlösX÷da+>Ozñí@^yÎj}²è§CØ†óìýûËË?ŠDÚ‡)sK7ºþÏaœ¡p8XÀŽkÒ~‰^‘KŠ.ñ`Æ®ÿBRÑ$'ÂtÐÐádR’'y‘q”E6‹CãBKþ†D'{-3êX€düYƒ=×"í}5Ñ­‹È"yºõsi‹4Í6p*±} dwAâÝlö bs/­wUõ´C%t$„'>-‹ª}òèõT´ÕhEe=»/Çx&{žÞÛ4|\ú=ŸÍö Êý·öTÆÕæ]Šu·‰î­°iÒ9³ÂŠ°&äúh.µm4/-!×¹	s×xØ‹GÐÞ\Þn¡«ž¤ŸfoØ‚ÄýÖÊ5µzÒ"Ó¥ö‚bF¥{vY"KÜÝÈ+g&h\·õåöØ‹úŠ#zoT cÔ¦0É‹Ù$
¨>E#ôm“zò¾.ÍH«Q!õd`©&>Õßxçú3ž½ð±~sè©|¡J¡_ÉØe×$"m²ÁBQ#[IT¸%vJ^K4¦-ò¦ã“*‚ù(¦æ¿øÝPäí½\ùù(4Zó'è2[7'ìOªN o6÷žolîlìÂÔ6.íÐÀõlð“þy³-¨UžØèâÙ³Ý1½!mú›IÍ]jr¿r°‰w¹dçO·.%—Æ7»(ùÈ6÷ŸÎ³Ë£ùR“Nø,nâ¨¸þ{–XÔz?±ÞÈ“ö? öÁUãC+ÈÎÖÜmÃ‡Th¢äg#a2ÍßL Â¡¸Ä¤ü9UÚ1·UŸ ›…‹…ÕY»OWr:d7™–<³ 7Î‘]Ñ	Ó55¾ÅR¯ªœXy5w¥©%šç’ßÿ^«£m¶CvDbÇ¥úN÷FCýh'+!–e(ýb…†QÖKsØÉGôÑù(ïƒÛ9¯¢Áöî§£I‘%ÿpŠï;pÂ¬-5”Ú\~*-MÚ¸CSÓtÚš˜¦ÐØÄt{Z›˜,ˆv¶Ú›DQ[tZ"9”‘ºž€¦³Œ×ö’“Ó{LtCôæWÁÆÔ%­ƒ¨À‹©¬K¼hÖ7å_ë·@ª«ÚuCÐxÞná("*À¥çœ?p4ôx÷­±1Þø­g~ÞnÎ”“x">Æ­bSÆ‚øø›pM£ñ‚}r€šÏl ®Úk4Usèøb
Öó¥®s­¦1Üé‡¥yU€Ù<¹ êº÷Ê7þŠyìtç²MYsB÷âwdò–»d+ÉàôDÄêvSM¾[tŠl,
Íe*U›…3÷ ÃÆpUç25·OÄ¤D×²§oFkIßiSµLgN‹3²FxÊú$Æ˜„OkoF”sLÔ‡	%”¦ã­æ0¥9y@z€<zð6Î;^A%­Ž9d÷gõÎ=ÑìÙÓ•yó”‰z3£ú²
za>Í|uëñÌl‰±–0EcøÚ, ÀI\ôÎ‚–`1'€p‡-Œ?,0ýÝ éöwÓ0.ÎRôð¸¿wxT/!èü)Îòn`0
Ô·-¹€QÓZv'ò¾t–-›º¤â×>´¯4–ž‘ ¢u`%ÊîÖösh%â¦©˜ßrc\Ö©~¹ÚYD>x®Ã|÷’qä³(-S„´h›™bè‡@ ˜É :s?ùÐEu?g!÷´xÙA~ÃOu0¤-4ÊŸ7Û^ô‰B?°wyÜV{Ð-vƒ–ŠBUØÌÀÁ¸W#(ÚSëpXØÑAž‹“ó¸ßul!%Æ/C-gÀ(„©TÛêÔc­H'Ÿ34Ô†cìòÊÒ\ž†m—0»L€,Žþç™ŽgYúÚlÕâXÚÓÛ=‰i¬¦Ù(Ð£*#%£óë¿’~DÚ8z¾]úö’ÍÎU§ÓiÍuÈNž§4:Æ€)h¥½4Ëbz’cœ›>uEšäl³cðZÀýçâ
+2$8‘ ©ƒ&itÑ8Go¯Y<$EZˆªiSt9Ü‡âÐSh#- 'œ>†??Oç\–×B @/4jfáÄh³²…O¢ó4›'“"ƒ¤cÖ©Ö6ÓÒÚ–Ç¨õÕ"Ñq”|LçÉ/Àá]ÿ•À™ý¿"‰‰ˆia\hT66ôá1ðRðŸAœóvF€ìå~ËP@iç&l³7`"Àøá<vž49ô#-¼JÐ\ áb½ÂA8ÀyìA¢Z²%ýøÁ„"Ò ¾jñœ1n³ZC=p³³ß>ŽÚK7x„JLÓ€>5 ¶2¢5—)ÑóK=”õžV®Jwƒ
n>œD Lþvœß>u4y‰ÃÀ¯LžûÛ±ºÞu7õTh¬¢PHGÝó8¾”w=øðƒâ¢X¹YîÜ§ñ»Ð'Z¢À×ºûO{¤	©4‡I)Å–z$¡"™áÒß»ŒÍÖÑã,ŽÞSŸôU¸
ªL/Sªµ˜EšÞ'Ek.dUi÷*+;,Ð©YQÆ²*Ôu5;ö»"§Æª~tÈÂ«iƒú1R ,íDÈœÖ{‹òãæ@D;J6öw.[Ä!ß®4ÌV::}ô,MO1ÙáËHŽØ:®-òïMêæ¯—äüN¹œ>
®áq‚,Ö;`"	k.°¾ÍA:é3†#Ð'Eõªla“£¤ä.	MQöÌŠôÕ¾–&—nøŠ‹I>ìÒßÈŠÃïÒû¼@G+åë‹z—°Ò¹ k6K‚zÂ˜yë­³¢çÝÅEq@CaÞ £³íSºHxwˆb¦œÊš*ûœçÁÏ4
ÏÏ“øÃcºÐ½býï/SxZøáÁr¼ìsíT&f‰´Þzs<ˆFïÃË_oRFXË|¤ÚnóROÊÒ%£A2Š¼—™b—Qa•úd¢¢^ÌîÜË«ú;¹®„cÇåñÇ×¡£	ÂC4¡ûÄg˜ì¡iW–;ìK›w:Ã®çŽ¾ìÑiJœ¨þ½•U¯Ë³2ý:öÅiÇ#}c°W¿®!uµ”n›!ä“o%ÓüÆŽÝë™ó1Ü:<ÏMÔð‚(ó¼´§RïˆÊwJ6—^™˜.vÚ¿'<þ*Ê¤iÈm%§¸I5µ„öðúo£d˜j7€KL5qàêmˆŸçlàÛ[$“ŸÚŠíóE!NUE@‚LŸÝ+êe%ÓIAññÈaY.’Óá£[«@æñzvÄäWe,SŽÇƒ¸¿~™ÔÄ¾‰n¢Y™ø™D”}—¥WÕ—=„Š®Ø¬&ý÷ýªqå™€—§	uÂçœ	à¦Öë~#LkÔO;Nñû"¥µQ”fŽ’«}ßâZ`%¦J)\ÒJ.×¼-z(šßej©öêÊR9^¢È
¯0\QEh¯çF{môlzŠ‡ÓªøE‰Î@™ë1t°âü×{2ü(¢±˜-ì§Œ#¦¦tâYÿ¢ßÎÔøbY¡Ì½ŸêÒ{Dñ‡\¡1ìµ¤1ìñ¡ALº¢ÈIWIIò7Öà¦x|>xÍÏO›„{€ñÏè¿ÃLxf­?à­s \ðžÁN’=Bùi,øž¤×[Kd‰¬Üƒÿ|%PÊöÍk˜ÍM<ÖÜÝ¯£âÌ¡@ÛÝæ¡ÑD–ÝÏ«•¥ÞejXìI¿8[¿Rìg¡úë­çË+äÇó•á¹wÖYZ.<èü¸ú–W;ï?è~g¹siøñ~ ËVWVw!ü ÷z~€o«««À<øñÿ½Ú¹÷ Êî®vVï‘å4¡Ÿ:Ë?®UVYýóy* Zw–ƒ˜’zø‹º ½PÍœÅ)—\óH-ñfUîÏÓœIDE¼^ä:N3æóïÓ©Û‘|žD~i°+{fÎp·/Ä²q­o†CÅÔ×Å&0×1ºôÎaà¨®N¶K|ÃÙÞLádT¬
çiïúßÉ8ÅIN0NO”‘PÓä#j"8«ÓÕµ©³`Ñ•n¼,G+[¶®©ËÕKXPˆøúßR8¿Á¢Ó^âM¸ê<Ê5ŽSé5L=]NRh?Ú+êÅ+]fÆWjÃŠ‡rq•N
âæŒs¢óOHX3ã‡Ÿ)l §z|ý×!ƒîÈ§¶ï_ÈŽ×€B¨ 4}sS„7ÙQ“esJÚp•ñ ñ =î	ì¹ªy”(T¾Ý},¤D…^CÒ¶|å–Ñ½Jn‰‡_#L‘éBÂ ú¢—tÃÒMr[MŠ³õû\ÌD¿¼‰–ü¶eqéòÃþPõ™ÂåÃµ™†r#M¦©ô}¦QØá{…m”O­°ãüîú¸ª"Çº¯†Ñý¡v™ˆB(—l«"ÉÒè—8…Ê™\„Í­‰/Ü"ð…;±™Z[í…­Å*Ë5×®X'J|û¯Nôô5ÂI½zÿSZž¡tNv¸ä*¿±'†€wé„ÛåûÂˆ¸ª >(ª5Tc ¸*`.	¬!/DZWA)wŠôY20cŸ–ÅbÆÅ*..’Í,è#è}0Ù – Z€rèL?‚âÑxLPJrž&x)p:ÂÏé¤`gkKÝ<ôÔ»Im
·“?Æ«öhopŠþ˜g”QÝ(¹W#{»áaZ½|ikyeõÞý®Ð¨6¦…‚Ùî3ó©¶>|tÅ7F¡Ä ºxAñ ØÙºqeí‡ãøÃKX®ubè šÏY¬„Åí~HÐ n§Ç´Ïšk0Þe¡~ÚÎ?|Øë¨„úq€@;4
·ÍÌŽ:Äª0Üó¥Óœ9Èðf!œfüHÔöÇ?ò®€aÔöÆ0òÖ%ã…RŠoÏ7Ö‡éI”–§hÈì˜9ŠøúE—Eœ°åE4·&•Ó4+e	«ÄíìÑ”ûäÇuðT‡\Iü´µn@–4Š×Œdêi’Q_Ö6­|[iÏN5úð±<O6÷vw·7vö^v^nÎsdæ˜ze˜¶LvVï¡¼%X ŒšÜ‚ßª£~g=8¸é••N˜ÙX[nX‡C ý”îŸdØ’?@¦„Þ ŽF ”îEJö&E[#Ø–VH<t“×åcèo8@ZrÒÆAæ?'ñ‡¹nËE‚{)¢”yà/Ç.g–•?8p}ÇO–—þlN×½P£ôÃã¹…è8Ê¡óÌ*¨k¸H‰ðB4€ýÖ¿XHFÀ,Ì9%×I‹:ágŒ÷»ë¿ ‰'þ;Ä[gÑn)pB Ï¤ìºÄÉ#g²•áí×hk·mí·é†Ý™µ™Îíø¼å j¿Ýµû´‘Ð¾†3«¥Ï\¾Üpf¿Ý5ûD1@¾FóøÍãó‰Gñ5šG­À×h´Ò_·'ÿÏ=šÇgãkør,_Ãoh*º¨ÄLð4ÎRO#›ÿq·ÑèíìùYDEôZº7`lÄ;
oèô<“‡7
rHnæ0$ÐáM¶{ã¸µ©>ü‹(,S‚¶Û : îr#3â™Qv±^·ädšPÊö¨¹ŸÉ8H·H¾H·½~_ÆJxâÝâúøBMæÓçµ´¿íC›ôŽÎ©Xy›á…Êo9Y…J—
Ã—QM³¢cÜ'.¢Ãç-]mxí9}‹×À„eX¿d¯Äëú%ÿ¡(SÊ˜ÙD…îÜÔû"c¬Aö{úF·’¨Ú¢ÅÚê¡†8ÖøüÕ‡ñ–Ó 2Œ,¤a¡ª/Ó4¾y‡¢I‘V°”Ñ?›3| ºzÔŒ›ë‹ðgÏÙÜ ï²!@ùA|ŠƒTó“-sxº.Š *ßÏ˜â±±d•÷0© ç	N+†±ØŽÏ8«¢Él–Í1q.^-¥ÛÇi†NÌ9Ì1Î}Èxv›hi;}‡vLk£³Aö3tMÅ·˜?ŠÊÖrtZ+/Š~@Éƒ˜z™î§Õâ™øP	vû0&ãj×õ—îî*€!ÖPÃPJ`e8êYHŽƒ?¹ žûÊ”]VÊ¾úŠc'®ÿÚ%½”÷ò(¦O¨±·äîªVúT²ì°V\ö8¨
á×ÏžeQÂ¯œ§9{Ù‹BÆ`¬>¸öMu•; ¬Æ«9‹ ;¦øñF/´U2ÏÆòÈzqböíSÄZÑ»}Ÿ2Æ’õ/«L“â
OõjWc¸ëv@ˆ}—CPyÂ­ôÃhFÒ‹Î=$ì„METw1ààžêg†PW—7¼nªf˜ãÐ_ŽëßZî<g0r†Š–fÞ9êõâq±Þê|äç	þñ2Ùöð€u¹„SÝ° TuH!>ç;¯–^»”l¿Álsž°I	¯§Y:d;	ËÌntƒÛÊs)‹MX1,Ù¯¥NPýO˜®| é	aÕ»îíüöÇt Pß«2­×sÔŸE2šX_Dâ¼GÚã¦@>×>e5”†|{*Ñ%/&èž¦MËW’×sNŒ*’ÆzèV9•ÀÜˆ^¯‘	¬¼Ê§Ô§ ÂT„Oƒ«b×Ót:\õ[y¢àiá<H}>Ç>Æª4¶©éÐ5Ö™ªàÁrö¥>XÉí„ÖX—Îy5®±JK<^•«1lú*GÔ`ë[+6ómQ€µ~CØÈt6'˜¬6$,1Eÿ¨ßGë`LâP³µx¾}´ñfkgcn^ |; AÐ¾ÿÞ­wRZø4}t³,0¥þ†ãªƒiíKbü=iÁú21AÎ‡Ë˜Óš/½MX»a7¨Ý“´Z?¡ÕâB3wedµRQ‹œØ(­ˆYÌ)ìò‹ ùú0ÑZqùÏ`îgEF406íŽãö\§Hw÷„m'’¢Ý:jÍogÛ:gaÝ~†ÁšÏÈXs×éº«Áy­Ä×ÑI¬\=ðgª?G›’b¸ÚôgªS G›:nv®wXNÖ·Á§9>(/,©®ùù0ìÈ»‰úƒôJÒOÌ‘Bl‡ÜÊµºÉ-2uH ¸€cZoôû|g_™$ß9²[We™ûÝÕÛ%_ý°^?Z—‡C®ýLò_ÚÒf4ˆÑ€N9gW×YH	ZÖ».æ²ZßÙ½´:î	½gw»Ò¹Óè¤tPÎ×—Ê’ÍT6P™ BH§Ó)«Ÿç BD-jV^W•ÛÊÕÛÊFŽÇ]w›F
nöòaºtP=’ýâDt0êh€0óÑâC³£–jÓ~~5'ÇÊÕ9÷llÜGˆæ¤gcëúFµ$p0·íð`\ë—†>áwU»ôƒ#ÈÔ%gòÐW‹9ò½•Õ4•ÔévÓÒœô6-¦JÇíš†Vl¨g¨úÔâ`÷±¾­`ÖÉ ¯ð ê_Ž¶n¢$Žo PÚlqkÌcãâSÁTåÞæN€êB<Üß.ì(w&MÀÇp2i¶š–SÅ”•LPÆ›¼Y•ñµ9ØÇeÝðÅtF}_tÑ&EÊ5MgÊ¥A¶çi?0mfO	z‹%„+7áá°t†¸¸¼d,³ÔR	“1l[È¦¡c>¹Û ]†2†òöy«È„ç. TeþÒùP1ƒüÉ;‡íÈyÈkœKo¹Fó)äÍéÊ^™ÁdÔONS1ƒüÉ;ƒ®äÔ¹2çÌ^ Ñdqãœ,ªVÊµ&J n>Så•¸gª\½‘SåÊdœ*{FS¥Ê=œóUyÌØx’rsò'?‚óu©Dq¾œf$ç.U›¾×\w—ÅpbÍä•”0ë	Z®?½îNîiaIlêÊaqÑXÑìçóžÉÅ»Ub	KƒÂ›U'ÓH×ÏîMaÊT­ÞŽeS0éÓ
Ç»­öÆ(%ÈÎ&i6G¾³uèFªàLÔ3¢L–]›Ûö®&ôyÅ–1ŠÐG mzB›‚?ež×®«—`+“°[–tÁ‘3«ÀkyyW±Xq•wùÝ¶9ËlÂàÝ¡ñÄÌ6ØI[¨î~¡¨Ñ¾c`_·Œ’>ë-Sª"ŒP%ú"6ˆjûIrk®]5à}6ÛMðë¦SRSK0sˆƒzl%›úƒÛVÌm'òž^FÑ‰Væ88újƒ`ª–û)Ã¤Ùdk~ƒj²s0þ÷†ˆmö±E¸+óéØÄ~æ;ìº‰›b¨¦¼ÅpZG¬V8ú®DÏrZg5³ÌªÆxpÛeMÐáÿ  ÿÿì}[sÛÈ•ð_is½ù¨¬H]lOe<*Z’=JlK‘ä\ÊŸk‘ „à  -¯¢§}Øß1µ©¤jž¶¾—¼ê}çt7€Ð—’’h»H¢}9}îcA]À@Tƒ€­oáŒ®ÿ7FÝ¬<§
£½jšm@•«ŽX^ût€6kf á³¶Øe 5òÃ³ô¼Ó5©B„Æ‰¡²éùiCçàù”ƒJýæ‹íô‡yQTô²ó&Z»Â×é¹ï´h,+p–~Í  g½¼Ã¤J²0ÅŠò¡Å¾^KÏ©=J%vøß)5ŽaaüïGëk›fCÇÓîkE]‚3HÏk\XãŽ)Ð<Stïn{BRgm›RY•|;Ó’Š`Û´öz»·3ŠÑ¬¬Ï½yI¦š*jÓŸ˜ª|dž—0zÞðdZßô{` ¾uL~žFƒê»’ë|dò
Ò,“Ó9Uèø^7ÖŽÿZ7åì'³äâx›'—¿!1è“Îã:¼8­$A¿yr!‘X,w†VØ1Nå` 0HMUOáp´”Ít˜¶ÂÞo@lÛe¹nÅi»í­²S®{?å^i]žÎÒÇÌÛ^ì·…«š1Ðß#ïäˆà‘Pi”özÖtš8TÉEÈbÏŸ® qBi³>Ú–Ï–lê¤µ°S‹~¾1ˆ)œ;S2>»Z³±’ºÐ7§êÈBë2Óö 
uïÆÜî^²Á®Ü #}ÇôžM:êvÏØ{m=›†BK#væ§Â(ð!â€[$ðk¡özŸý¼°‰‘x¥ù>ÎÖÆùm1¡ÏZvDñÁ™o——jÁÓý»Ù‹ÉÊ«­*ïRqr¦Ïåïûš¿¯<f%~ÓKä:ò W–ó“C[óÉ¥¶ïÒRÊA(ØëªaÅÕr~kÖ)ÁãE³©ÿ89¯27Zaú1l$W]L}ˆ½‰µH6¶Ë<ôÅñ ¼ÇHâ0’éd}}‹ÿÏ™PCj^pB‰¯‘‘5­IÚyzÔ2¦Õ“K7Rùü‰úæTØP‚ ˜¦w¡ËØg,[Z§ÉzÊk_¸ÝÕÏ2ãñf.(¦ý §½6ã`•èu	¬‚_RêfTÈ°Ø.ëv8Y¼ì²RŠ-—¨¥ÃŒäÚ< -AŒfÚQRt„™òG\ AÀY7Î¶i‹ ×y‘Û ³e½§7‹¶–uÕKfE,Ÿ þø…]ÿOƒTß
ø4V¾„h`ãa'~ÑÔâ,ë¸	\°äK^B$àÈî¨éêÎ(©¶;¹öƒOè¾«*Øe½7qé?…u/áÍ§dIÕô'&iUÛà€ECÅ'„j&•e½7¢øT¿„(œ±X3H“„Çj»5ÁM É",!¹}js]Éð47ê¸äž
WŸÃ=®ŸÉ¥.ßXé°*Ý'Ì­²FÛ.s)à»¹åS›‰äãN %[e˜r#%>¬¶z"DîWdO¢©t&G”M¤ÉÊmÀ4ì]ÍX³'P‡©äNœMkŒM—\‘Øu^U·.W£VÃÚ`S´ãÍ¤CÃfJú8ãÍ6	%dYÙÔ`[Ê#4Ó*`«§ œC³QJS©‘§¬«<ÌâRiqÙh3ï,ï®&Ï4IMð‚:`³¥ÌÊWbÓçìœ•Ì6sNàr[[cÇý8Âø‘ˆWQdÉ8ŠÒóÑGR÷A8ˆ>t>ÄID&i4±$=­¶SÿÜ{ðÄbvNÏ"Ñh› MZm%þbÃå6»Ê•x»Gëj}ja¶½ž:ëª‘vÒ-Üæ&š›Ó`LÔ®r³$ÆB#µ'›3/²>E]‰!á¯…Ã S{1ÖÚÚó½Çü‹þhÄlâÇc/DfnÌ+Êû	ÆÖœÐ·‰àg«]zŠ¸/”j9×þÈO}L»NîÈ0÷z“Ç¡Ã)uODÓå|o6‚tÆlÒ‰ ôçÒÐ'¶,›{ƒ‰´4á¦2¾òúçAÔ%YƒÉªið·˜~¤ƒÞlëÜ‹a@¸²{–]©&«ƒax*þÙ¹*’»LçT„³TDÞÂÈˆ%qEñˆ]è7…òˆf©” ùA£ï€á1«çÍƒoµA½ê•r¤•gÌÕI<{œyýÈ¯5”ñ,g±]lÏÛ+Úi‰¯§ÐE3×²¯ž•ráÕÂ&å|ã¸ó@Ÿš—ÉÔ7×¯ôi´õ*¬,UÁÅ™zº´­šê¢|ÎðysÒK}Ñr_J…ÑÚ$o¸6GþwuŽâQ…
£±»a$ƒ •'^èm±7ÆçQìÅAdgž’
cs"m‘…¼6_¡“Gø-À2]ÈìFÊ.ˆ¼zº¬Ýô²Äºê†Ü‚±­òÂ+€EÝæ¨’€³š«JB~ ]¼)¨“ \I™O(ÿfUÞÆ%¬• V¥–K¨T¨ÇÓe¤Ä$‰´UŸ›ÕW¸spÝ¼$ðcCÀÕ1øOm+ÆúÔØ4‰­EkûgpÀ^gè±gÄè­?¾þ»ó©?N½8%<„î§Ža³Ý]ÿt
Ü”ùÝhçbz@æož!{«hBŒL/P„1HÑû¶
4õœ/Ú

º*²•Ú°•øÇõÑfîÞ‘.ÃŸýs¿ÿÃita“ø3þ ‚.sÆ¡„ ùü„ŸMÀ#‰5\9ª”“p‹âTÆ¨hU,íì òL•ÎÅ¶¬âcÎ¢ª.åëóG€´–kÑ8˜JŒ¢ùhøñØ=è®T{ØœJ)çFZ~µj%| ïp]T2ÒB@wãkyô°é42Z—ygP‹ä`³ã?»·5¥Þ)¢IÎ¿%¡üßF1*uÒè\Ìj0ž¯¤Wqë¤D¶X~Õdd‘Še×ÔáSå|©Œ¬þËÅ%±TqPÎFd_.6•¥ª*º“l–…öÅ»²¬ôñPÓsK¹)¿äžÌÛ‹ I=´×å…z;e•Aá¶€”<1Ú»,”ÅÓOÚ¶ýKöÉ|Šµì“¨h’h®õæ½Î¥ìf]3:ý¹g¶«Añ'·{¸LÉíúÞÀÉd1Kz	—ÜVv»bÿßV²ÐÕ–LI´,IËd-è™‘ì_ŒYÊßeµUÖgba/Ž½Ý áÿIËr9ÔîN²ÍÊOw¿‚°äíý¶*ý·R•’Ñ¥òmàS²Š3Mhv¿pŠçSC‡m]Œä¬Ý\'"´ó907Ò‰Ì‘¢Uá’Ð¥
’„.øž7å^æòåw™­ƒEËÅ?0çÕÜhü'3ÏHÏ£[DNLw+.ŒhWßœ=¹Ów‘ân2Ã­Í]î5W¾—@Ô±1ËäßãØ]§O{{mØ×ì £Í‚»<>8Ãhtæqÿ›¼g»ŽÎEúÎþ4Ž–¼Nüøu0`OÈLç]åÛíî4ü®Öo
?ÆÑ0 Qê	ÿ”t‡€	Úí)Ó»ñé”ßPÎQ*Æªipäÿ8õ“³ˆzƒqfŸ3Ýu;öÔP%´DÜSæ“'òz£Äÿ]íYåÑ.à,_lZßÈ^3’G>DOÉ#¼¸þ¤±µg»jÕ!Æü°¦nÁíT4ý(îØqYý¥—Æ×?÷§#¯Åþö·Zÿ¶q e>*(_½/4þêâà« ^AàWµÍ©<¯ß«çp®QÌv€Ç¼tk}i¦Ž1÷nÎ{²vþçJ—,âøæivÝÖÁ<#Çî™ÀöJwi&~8 ¶k0o
¦z‡´öž¶ÈÁw“ÔK§‰˜ñ!„[Z²ì®HQ¶~ñ½É$Ä>˜ÿÅ=¨js6¿8ö¿|¸ˆùýib~±ò¥`¸Â)<iq$ÕØ¶#‘u‹æ9„s¾õ=Ô2ÃÌ’J½œü56ç¢dLv."¹Sš´ž…^Ôlr×K„uèTŒ{˜ÙÎˆO_)Nb›#“Nk‡›•¤7ãC­7£Å9×\SÎnÇÓy¿	—É~e®U!o[UÓšÍÜlÞÓO«ØŠ9Ž¼ÌàæÖ\ªØc±Õá?ø;‹}?ÌáO|¢Â::ìqäo~Ê`˜áÞ;Ã1ùÅÁ!¯
ño"þEÀß§~¾cð+±‹?í—ÕooË2É_ÖÜ:É="»@NƒÿG&K`¯âŸ+Ý
sqj¬œùé’¬œ¢Kî\Ø"PzÆ5\ÿCSm/;©ºÅSþâ²{¢Â8í¬w'UˆdÀG#f/öX]ÔƒYFS4ÐZò>Ôß˜§`™bÔ±[ãÿæ‰aÍ€k~Aîõç¶/e¨†åcè<m]Þ™<¾eCu€¨ÿEbrÏ5™|Þ´N"À7hJ(UQÄ‚·X—zí°2í&&‡ÖŽ›™†ÖºªF/ñÓc.Áˆm‡—SlS—ïN®UÜ}›ö=QÞËå§¡U+·ÍZjWÅÇE¡Ÿ¥ÔEÖá¶X«rÅja'kÉŒ#^½3m›YEh>%»zq»Ú÷—£ÈCF½§êÃ–ekx›ÄÈ%ôŒp±Y¡è&¦ãÈ~9ßù`ØÉŒ<6í•2//Æø¶d„Ý½YÝnªe\et~—:Ñv¼8öÏ¼p±¤D/º]½õº1;€¦ëš¦‚à&P'×ÙQéíùö•
«½§.òxà6W^ ©aCAáð/‚Ó ˜Úè”%S½å
€aÒ;FÁÀƒ‘à‚øð±u»gÙ!Ê¸üzd¯Ex÷dC™||²ßž>”Ug£~¯F–ÌnÙ­8Ò\ s·x/,úhž;â 3  fqÈqvLæè}èwTxö~Ï±YÐ÷Òà½ÅÉ5Œ ¥³÷wø%Éî–øn£Röæ¹<•ì.æ—UtdÙRÔ€Ð±ZÑp’Ó)×Ë­¤»U‡±™®mF}Á¾Åbõ/¡~’žTV°'LvÖ«³íîl´#){FË™QÚh‡BŸg,çx-æ°­­±oýÑ©‹È¤ä¥˜Ïg¸»»k/_®ýaJ¢÷Ó«¢aUÀÁœºt“ÜJ¸IO©ÈçØ…Ü&B(ZÝd2
Òv«Órº/ÈWð®*a~àè‡œî»û—¼ß›Í·WkÙßÊßëo¯Þ9ÇÙb‡S„ãØµ±¹êbãbIB®2Ø¬Žp–
ÄæB$uŸ¨/;»˜d­*OŸOÁ=â‡ÎcvÿW<PÊLn&‡q¶Äé­_“|]:†jËS±W¢ðv·îÅ½´½¾âFÂEk½Öøík÷Å"%”ü¤‰ã×åÝàºIò6™7Š¶\»Ú’°ÕÜìM¿U¡¹ú&VE>râƒÂ;ç†)ù06U¯„<I¨¿¥¼r]Q³RGÍ±ßªoQ[vv)ðÝÒ¿@Èiî¬HYú˜É4žŒ†Q~þ-ì±Ú¶
ÍN>ÿDÅ¬~QÖ.Õ½¡Kä9ÎùérÔÎGÖb†r9abÎol—*+–±½<Ëñ~ˆñXÎ\ÙbS¨XÀ™œV‘›—^ÿãfçvS;û,/t[ç…É±wÑùÀ¥Äxö‘%—ˆi£Bm¡!”=È/á÷ª`NÚ<àž[÷õý¯Zn¦ ¤Yz*RçÑp¶qƒt¥fÒá,‘£ñ±!)d_r§ìõh%Ï_Ž½:#aÂã(G|¼"¦h~õ+é^"ué<YM»Tô	8Î_f#Ó"ð©(žV†)+ K\òÁ=©l9ŒbžÙH›óMû®Z¢½"QÁ›ßjxo!m9Zõ&gœ1U†¢îÜÜT€ÈþX{Y"§”¦c’s<‹V¶¶6@(Y¤E?øRŸÚ,é°fyƒõ®³®FËäŒ*²
 HjµQ8Ù•¢Erj¡Z…µ¡#%òf Že6•
ÐWƒŒé}ùµ§7Ô™Ëç?sP×:k»Úí‚zæ»W„EùµËæ™ÄÅéÜ(ˆ“•Jæ(ñÚ³éÞÈä‡_’` ËP¤™Ìâi¼óKâiþ"›=âÊ8C¸[þc=è-ûIïPÏR‹ˆÓvW¦µÒ6ÇÉFãÝÀ#æ©~|ƒyª¥|þ9k½n0su3ßÊ,ßÏS¾uJöç/´4™ù÷Mò8Ë‘o6‰³x‰;ƒ³|Î$V§A:Y/36öQòl7,“>Œ{)|™ýÙë{ø›]ÙiÌ3„F=„=û0XÒgöþóä¡¶g®åª¾›,Ó¥¯LI¦L5È0­ñ¬¥Û³%Cu&á3æÚûú$ÂÜëIÆ$<¼2êO[†8NÞN®Æ»Â%°_›Èž1…Ÿ1Ÿ%=_)òÞÅÛéžô®ÿûú¿ØÓ×ìðhïxïÕÎ~ïÛ|ÔÝÐ1WEz?qn]qçu0KJìWÆ+2wŸø¸šãJ
jÇ­Yd>iÆ¦fà«AÞ¹U“IY‡gð?‚Íñ¨Órˆ§xócÜU€báQŒ§ÅW¤[~âé{4½5(Í>[ä,œs5=T™iqDå4CóÎS¡
Ö¹öÂˆY¼¥ù"ù±NôÈô@e‚¢åü$t§'SLIöÏ€X[¿ˆ'—²g>Ø¯0V¹å	”0yg[«,†/LÃ‘k1ñùCr‡ßº`ÒúåËà;¯wt?H‰…£ÔL0ÍÂæ7bMI[³œT˜ŒÂ=ö`Ã"÷í¿:Þa¿6ÊC–t¼‚&šÈ8NÑ:aû­¹i±Ifèì2sNžL‰ï³¦ 6ÆÏÀ¢!_ü'Ä:»ñB[ÙÆ•.²žÆ¬ñ4ôìæÂ›vŽ=a¯ø´Ë$Ù* “DÑJäßúäe{ßn9(aà6ùäqvòôÄ“@Ûâ Ÿ8Ó‚1â½«Ô¬Qá?ZÎÂ’MÊÌ*mRX%kás´ŠY•u÷¨ìeï„=Ûõ›ñvKØLÈÏ_ðÙçÏ2mÈŒÆ–£õvz»_Po·‡ÒPø§}68M¨u¿ µF)?V
»SþÜâjKèâë¹§§ª,¨¦âÌúý¦\ÄJ[Q"sóv×†Õ¸V@‹Õ¡ðs¬ßp¶5ÜÁþÍª»‹1ª!à"ì*–MƒŠåR˜G…¾«s›AÙ…Ýad•¬Öjôt]á5KùšJÌÌ¾”Ò½˜Q¬QZ™Tª£ÜÙr	_óQß¡»ù1¦äçÂ=ÎFy{jý~ÌçŒg÷G"§¯Ùj9ši™ g6åØltK1àòùgW'n°Í»8x‡’y–wR¾±uYÞÿ4?j£¿Êàm$ã¥«¾©=Z·§ë¦ø>™wŸ,GßÆc£o“Ý‰Í^ò1ì3L§ñG+Sè}ð‚Ä¹Þúv¶ NWÙÎÁ‹{;'û¯Ž»O^~·»ß[eï:˜;ƒõ§£®Ê=>œø4heÈ£ð$ò\…ï8“qÿR™º ðt%ƒˆ»ŸðoÑ]ð
ËÖÄ`¯ûnÆi]ÁF¦ýsÖöãØ^Û2›~k/Ž#,oåñâW1{$A–#¥‹Y¦`¨(¶…¥›nQ£ëUM—d(ÆMô.49²‰¨•óx!³ÐcGü øÉð“‰ñXä·¦…6ºç`øzï£0–î=¼¿.}g"—·†0ÅD¼ ÊƒxÜníú‰ÿ=¦ßá‰í™Ÿ ‡‹ËÛn­ØÁÏ…Y2Ü2à%œ¸e!ØÅ‘y!¿1üNÇþ8zw¸kOØ ^È-ÔD¸*ÎÖîe¶·|¤ûhÓXÌsW‹’
¥kš•T˜(î¼¤ð$ö’ó9iaSÙ·üM&L”2_é}üJœ2H(ôR¢Þlõ¬P~¬eLeÝ*<v¥fU=M•†avùzÍS¨!öû0•ÏÖµ.U.Îõ™Ü<Å	õY¾Ÿ@o¸ûiþMÔ|ä›uBÍ^ãvCÍŸ4!;QÉì†ZD{óúÐŒ,¯tÓhÿø@J’+YÎœ“ÖÊ›uc9!Ý<ï0R&úëÉ.Ê†‡P¢	þ ×ò—Çüþ9@S˜šüå¹žVziQ3¨ªnBcçœRk)À`+Gêº;3—’6ƒx€è Àš^ Ëiá0š½ÏX{¯Ãh¤="ûbUÞýÅ»ŒÎ«`¯ûÞz“,óäæúæÃ®&Ó§Éu´ò%Ñ)z^Àôæ±ýðúç¾¦Øxs j~îuÛÕ¤
„¾0À2 Ø
¸¥ºøÃçBW+.!œ¹¿9ì•óË†À–@2oqÖ–©¾5³‰Ùy¡˜ìÜr #•½Ô;|AK¨;Êð¥¸\4XxS$•/à6’+„¹¥º²¬¹W•?Gð›×m¦$rZ¼gè+Ž¸œgÜ.2;q h§jkþâCPûe»—°^¼_ŒsŒ«¤ ¡˜€¹¹| ½p€»`€±B5wýC™Ä	µ!æÄõæ~îäýæ¾À‡±vÆÂ®Ì2BÎ‰ÌÒùßL`–Ž9AêÎöÚ™»>÷&³w>ôú ³ôìqíŒ©§>oˆ±”Aƒ"Z ¾|0³¼Õ®¤›DqÚn{«ì”kÜ=®jëŽÐIÊß‰Æ/öÛ§üKmÉî¬ÀÐj«ÍêíÆ}„ã'†§*bû†­Äl†ðÒ‚‰ek¥Î+ì×vëÏ‚ÐÞ°å3Úb­uc~übîNÎÉœÄŒ(Xø"Î¼‰XS_Ü©½Ø˜,_3¼äÈOxl¬ LÒŸ£ê®÷~{Aˆ†ŒažÈÕ±·ü¶Yw¶˜í7OjÓµnrËš“Òîè¥—žw½Ó¤-7oŠ7ö.Ú«åÍ°'ÍRÒ¸ÅÎJö¢(š£î€­â@“ræ¸¦ÉøPIîÜ¤¡h°(“ ¥#s“qJDÒ|"…ªÖ<eKñáNË•{®œZâ·~a»µÖZ¹ºl1G¢ßŽû‘ËüJ:¿	€©Àpž¾Å	9v-ß¢çåýØhS³êƒæ™ãõ”ÜÏ•šŒ›"w`¹ˆ’ÆYJñ’çXAºšH‚œmLžÀXÝ)[ŠcüÝ•½Ø¾T^9œÏéêßí{âLj·(0¡”¼â¯Þeº[ƒfÝUÂ¤¬o©À`mÆ÷/3ëQ—°¥…®ÂoÍOŽeË÷ÝÆI)µàç®ÖÖ`º›B%äð[Ü«‹˜–’Q4ÓÐYhKv¤¨>ËM88Š–VµF«½77Iæü­ŸjtR‰%­·bBPÈýÍª¸Šê(£ØDŠV!ÍO†“ž­‡œ,jö–0Ó¬Ýh¤JÑ\W·½ò!3/5©nc´@sg/šÛ=kâ]Š¦óMvp´·Ó;>YeC·gzÑ¨õ2”X˜!)¦hÖ˜¢‘n1L¦hïrWÖû—Ã…EÌ4œ5=Š¦h×é
»¡#‚¹ö›ÂJ‹ð)š+Ô§h<‘6KDêÙ$r¨h®Muc`õ¾PÀ°xDAÙô”åjøqd×,½À@($T¢¥¨wZT•™¶©wðµÑécã¨­¢-F)“k•bTWÑÇwÕßNÞ¢ZHXtzÑb÷’R”bÑ”eñáhE#ð»åµÅ1¼ÎêvIÞV+Áp˜ºú†ÊÚšŒ/›Œ¼0ôãO"¼¬™ƒˆ>¸Œ‡|ŠEWö´æB‹‘‘ƒÝ|ˆÌmÂôRY^§7	#w±ëe/ÆõÌé?'Çt¸ÏÉ§ŒeÔ	}‰\˜¦ž¢K’g3@á§Þ0A]qzÃpu
“X$HU€‡eã‰Rp„±
iÇþØ› An_õ£x†öýnê¬ó,ßaë½¤´‹lûgð"¯3ôƒX_c :ñãë¿;žùãÔ‹Sç3AèzæVåzäú§SS¥7™wß”„Û×Ñ„3,Ü‹dp•ÏàÊÄÀ]üL¾^#éúÕSøâÍ\TxòôÜÖŸf¢áQnÛA•ÆÙ4®GÚ|qk®¾Ü`VÒ3“ÇkÆ ÎÒyŸÏÃÍV
Š4/[»Ÿ,ÐOt«4ç³%‘&þ <ñç$J>ç$Sr‚.R%³+ùÈn46?òÖ¤r:Râ6w€Ý‹ƒaÛ+Ø¢Ö)ÿvªüÖT‘¡Nˆ'Blwóääybuò\Œ‹§E«³P§XŒ•+•ï´ë™Þ.'˜‡SÀ^Ô\\ž¦0›P6Ï4¬ï¥X+f²TPÓµ¡³Ž3]›h3(÷ujýÃ½W¯öŽVÙ„¦Õ'è sMþ>¬¦‰ß®qv(¤oAíÐ7SB»Ð6 7k–õ"ÆÝ(1-ú	ü¢È2Ò®¦’³¦²Å³ê“Å†Î¬OÞœ7å’.móÍiš³­¼5… ØÞEd\’C
"ñùÊ·4’¥y“´gˆAG³#l}³&ý8¢“®¶ £² ¸:†B·æ4ð·P©Æ”bJœ¥Dµ@=6Gišúµ
¸­8ˆ¾gô 'Ý2l•›f¡µå;hy°2EšM›AÚ1ó4Ú›ü |“9±Þäƒð¢/¤lKÚwÐkŠ§Z*Èã¥[u~¿ ˜ÕSœO‚½ôRÌÔö…Éï„½ä…ö–ñ"«óû\ä»†ÛO…ø,/Àæ“û@ë'Nv°ßº#¼+ºƒ‡°Œ÷¸4Á_À]¾kÈý„(Ï²‚l1»_ ¼6 =‹ôÚ¼!YÔ ýýâDs’»·ð’ ¿D÷ Y³*JgÖÞ¼d¦ì…Rg0çˆ(çb¹¼Ž¤e1q{“˜}Iœž$‹ð#1ÚoèC‹°¥³š‘›{IhjÑLô5h'U»ÀÊ•;—q,Uj	b5VÚ5¬0'EëRËiŸŠúVvÚ>Ø}1ÊKµV¨tæ?Ã<~ìú_©T)DÑ,;ö‰ ˜ÐSÌ]Ù(KZåE®²Ò\¿œ¿"Ü.  Ä0Ë8¿Äl?1p:âÍà†GH@ñÑ›HÎzÛ‘²ÑíÄ×Ø…OëÀ·w´°{ðÝNïð¤·Ó; zò9ñr/¾Ü¡‰ìÉgƒ3«ßûðYòšøï¹¼÷Ì÷Áä¹·D~{Î½4ñ&¢ãÞWtÇ½ÏIÂÞSÿ,–%éžF a=?ø?ÇNÛBá)>Zõfq>1Ê¿À÷£1|`.£àìúŸa?ð˜§ÁhA­XäýIå|uÏðDƒøÊBÒÇPofÐ
6}«Å¬aÓ°¾>zÁcÕ'vä£ÞGÖîîãwpT:ÄmÑ¼j—k¶‰µ¸ž*ùâ§é$ÙZƒË<íÀì„Q'+êâ51ð‡Þt”þIètO£T OwÌºÜ§£iœTßAQha××ñˆ=©è_»À.Û"á ÄrØ_ýÊòÌ=ñP7§ÉŸƒô¼Í·
vªµ2[×„÷5v´Óó|+ÞeCÝ¿ß]½³QËŠ'è9¹‹™~â®fÓ 4^:ôd×•XH0+‰Ÿ:9}õ÷“ïv^=Û¾ÊZ˜$¿åàP\¼¬iKî•‹ÕÔmKÙñûÛßØÐ%Î´ "Çà —na¦ J'ÁÓÞ'mÇ*éû.ÙØ‡Ãß‚ÍG/`«o¢œW“h©‘—NEPïž_k˜3‰”'é]–)š|˜Õ»	oèÂ¾%Þ™åÈHHŒd[“á#ïiv«~@q«Fýi²…""p¥Ê‡<÷g4MGAèî}›¶Á&ä’‡ÊBˆ'îU°åÌ[ûÀcÇ›Ž=N_'€åý y`5;0+”º0³¹pS;šz€@¢É}× 7	Ú­5o¬Á}M§‰}†4è> bPã©ÞcÇ|hqœB7@BÜÒ:øCË	èÉ´ß‡;13¨ç×wa—÷faÁè•Ø‡Ú;GÇkÃ!.ue‹äüÔ$‚ŸÆð$×=x <³ñ,@¼ü¸þ‰ý8õOÑÕeˆ82X~TÙ0ó<ø¹ES‹Ê†ÏÍu€q)HÈMaô!ö&”ðØLræ‰8€)$N³¬jv¨Ë(Òõ¡âð0NL¸Ä[`@'Ò¦¤¸= ?ÀÕ*`GCmÿbK;bÆ[ý)ôêN']…^éj¨ÑYÉUk%”’ì“¢ã–@÷©[¦äCÜ€0³sî·»%%˜ƒ‰MZÔ>8Ü{utðúdïè;øõ»?ìýµ©pc–U„¤2•|€¥ë³$«$?t¢¸ó~£Óívõ›ä’(_`ì?ø9þ4(,	*‚4 £ˆƒ6T$CÍ{CÝvMwÅÅeÛé•Ã^ mç®««Û’ûgGÔ³rÉV¶—Ì![h ìµäó4W¬`¤¹²6®èž…"çÒˆÙ&L»Êh÷«hÄ;I®yÚïÄvY¸n½Öèß@¿<“ë £÷pÍþÐ¥¦nâ}Œà»JÖÎD¿ßÓØŽ}xïõÿÆA”0Ÿa2GÉ
€hì½¿þGR>óÏ•¾Æs$·	`¾Þ¦„Ï˜7.÷…T4«,e¾Í“¸Jš>q"÷k+_à'N;Z% ËIBdÑ­Ü4²+3ú‚Þ%>É,M:<oP­È¼âÊ·Zß#ZpÐé8Ûäö¡7ˆA¾ÚbÏý1Èºl³»ÎžÁÏWléþ þüåÁîÞ‹cá²6¶'@RsŽÑE-ÃHüƒÅlÌµfW0Û\uHfLOÈ[‘ÙøÜc?Â•Ê¶p HƒÓ‹žré%p®‰ŸdÆ.LT”|®4NQäÂNàÎboÌÚRù²hÏl2Ú¢êò7W³ùáûÀ‹;Þû ‰š‰€©\$¬M g°TU,*mûd?ñ¯¼6š•eù	¸áF‘Ì$7MÊåÜ²I}¡æw ¬•”• èOF>Hop*~8˜D@sÓBìÆY.ÔðhB¹´ÃeŠÍutš¤úÙj9ýð<×\ÔßE6¢Ñª7ó;÷¾K|þïtâÇðwP¾÷ÝÆæƒÙˆÝ'¡íÔMös#vŸ•¦Sª3wi)H^i†_ß¾ ^2Æ"ˆûy2qkíÁQ×?½\DÏŸ(u	3§oÚsH£ÌZ¼DKçÆ÷ý4IƒáÇÎ©p£/—KÖfy”ƒ•ýGµï±%`á9Z<ÿ‡óÇ í×ÿÛ?£ß-º,¶_Í:»þû ïë=Òt~¥–eÚ¼Ýîöíu.&.÷‚ƒ~³~­Ìú~œúÿ‰ê´ roþþƒFó“ƒÝÞ1ª’ì úq0I# Ý!ÆÌn³6×àK_oây¨—@µ=×ìG€€Æ“ nŠÏ’v(Hù\ÏŸ{èÂ$Rˆÿ¯ÿX>JV5ßl‘®¸J…@ƒß™Ö*!ûðØOÏ£ÁkŸ8Ø¹ë+“ëïµzXÅÓc/¯:úÜOåg¸5 Ê¼‡“ë²“h€*#åh‡'}x†ø“/ÝS„¦)ó_ÁØw
d’Ï6çªaÞ°ÿÂ}ÈOÚ—®}š¡À!ycs-¯GC=lÄr†óD³Ä•”]zb ×•dY¦ZmjMf¢œØ yù#‰„räß>äHe'G*&à=¡"sÆaÇçÑ‡ýð{¿Ÿ¾„{3j#j‰oÓ¸HÍ¹¡šTÍwWï(ÜÿþøàÕo"•"JÉl:±l‚'ì¸’­Ýà·2Êh2S",1bmlü6k[³úÆI‘o:va1w@—¯¦ãS`‹žà:º±Ï…ãöÚÿÝ];[E‘Ð~¯3ž@Báª|JžMW±Ñ	K^FGê€‚B(^®·„ Ýž²ØìÙðç`à´C@7À0†Ñl[ ê°ä,p¿HÞáX2i‚¡ãZ5‚YïlêÅ(ßŒÎ¦!òXa6öÇ#@È”ðÎI®­±gQÜçÕ9½gýs¿ÿó†(h<€—Â¾Ì±·¢²F]ÈŒíf…²©e£åI5¼À|´_ýŠÚ…“O(õ°kÌQÑÛýÞ‚Oáªc|µÊ¬¯¯ß%ëUx”?ktŸV|ª»ÝîgÊŒÅ sÏË=ˆAæeþC”ox%¨íLlƒÙ™Ú Ô@ÿ€Sä»º?øÔ¸fàŽ¤ÄîÈÏÒs®]gÛ¦7šxÒq5- =@~FÀ+?<æ$c6Pzí§hP”³ò}èaÒAÃk¸wÈ	(ûÁØÓ©€Œ9VØ–i±öd•†Š¦egûï‡iøÚ#@#Ô´ßä¤l•á0zKäá1ÔÍ>á’î0vdÑžZ“ì‰6íÔÕ‰Mf¶Pú¬h;„?'Â‘CYš™S•ì3œxiÓÄëÔ¶•¹sŒõi5JWÏ'ÛÝâß
ëÇïtw9kÎêkØŒ[Ê½ò7»È˜ŠXºhE]iëQÛšrjsAžVóXR¨V´­Ö<'–1³IúkF0Ù&E‚º3Ò €9È'Ã'a"ÎÙÖöÐ¢Ê¯†I	^'ÜYdÌSv‰’;¯ÞŽüd¢7x×f¹†Ó$îº‚7MäPóí>ÔkºŒx¿çËºQo6á–‘×^š§ðÈcíÞ4hœo³-1“JÍ¸¶ÄõæNUPo†$œ„ÌÞ&2{9û÷¡ó d1ÛÐšÏÄw#œ%}æ%[?¦|¤w-›ù2EA¢5Ì¾Po¹é÷žìN®kTY±hBPfqb-BµšwN§n¨Ú^Ç^Mûª+–[³CÀöæ]Æ´tï+êÝ[&.7	KMgíÔè”I­HáÑ^éH}ÔZ³±hõxÑˆDŸú˜“3ÅCoi<$1¤Å/½|WàéŒ›,’û–êýK‡¶»RÛÅK‰€ùPîBøÏ£«s1¾­Þwâˆù³ÎQŒ±6EoùYvÏ=ùÏÒdÓºzG98"/¢Nr»›F¯‘{Úî©½bàZ»{Ç€‡¾ÝÛÙß=°¦èË9 Ô³ÍÚLum³FH­W´fV|µ½3Xô…iU«Òö#u÷*Ü&èÌ²FeÓ(uz³Ö”ªÎm¢W[S=½ÚfÐÙ«Í­¿WYgª¶w™ÉO=ïÌÖßmpî©ÑÌÚxµ5ç¶ÊZúÆ„}Z}µ™ÔôR}Öl¸FŠûRÇU¶áVâë_ETç«m60Í<,$º’ÈªQ6”j#jùÕÖdˆOÒž#ò`µ$ @º‹ *5W„FIï•ÚJ¥¬ 4I.Ò¼µW:¤%Ò‰‘2
Q1Ó¬VxÞÍY3¼x’ÀSšÍjÓðq9/VÊj|M…qUÕÅ‹ÖRGeœÂAó£—¬®FwYâ…gÒ~èll²süÇÌ`~¥èQ9çkVrÊ^Œq“é(ñ©Š»#”þ|çƒ‘Í‡d¡*ëF#S±nWâí¢	K:·þá÷¾°³ó4?â.½0™rê‡øIoª.Ç’KÛºƒ–Sèó ¦èú¼ƒÌ€…è`ÉdEÀ{z§Á(Hy^«×‰Ï<ÒpùúÙáÞñØ‹ ü „‡¹JC›6sCmÉ”!AC‘ØÔ,¢ÜLÀi~‚ÿ4óñ®½™¨ù-Ô…D@Ÿ¦QçÈŸŒ>²ö~"/4Ú,è7“ç”3æ´?c˜lOdÚlÂ¥¶ØDÄ÷Ç•ínà^ó­–ƒmoÏ>š7Ã<G(iÊ±ÙÚÙ½ÒfXm‡å¶¶Æ&i0àrô¥~•Ü¹Â˜·'±ÿ^dhÂév»]ìØ„ÁT¬¥[Eù6e„fìp–ê5ßýfÝk°6ãHä«•• ± %Bü»Ð"4.çŠ†ç„ÅBC}ømhKö{‚Û(kKF@X0èKþTà¯mÖ:xÅ•ªÏž9š–Û'"óµÇL·1K,súpäƒTÌ‘×ÿ¡™æj´a&ê0}˜ŸB¨áÞ"qÃ¬ƒ5D7¡»¡>¨Zx:D:œE=ï<b:¿-³ªe+Q-!®Æ|ÿ²M`Œ`×V¤™(óèÈŒB¹Jh	¢*tš˜Óª›)÷PdP8ï<@Ï±UïQÞÁ¼’Š²•Å£‹Û:>&ß«‹Î#¾yê7äÝ#Téª-þ4Ý±‰ì´CÚ;âXsJâÈÀ$(4ùñŒ’Ï…,¶.ÐÚí0¯;Ç°e7a\Ã;ïÒÖk¢Ui¶ÕjR¿]ÊRJLÑß6¬4lñ¿ãèþmÕ 5®"4wA'l/ó¼>2ãªÿ³¸|Ý[aø3–xÂöÜý°øl|ý¯QLFQ"óâ%åÌCBeÈóÃÈgQè¶,º½­^¥à3›œv6ñôà¿ëRóžôãh4:õb/½Žûòv±—,@ò<HR˜?™”§òï¾…ï0Õ-ÿþ·Ø×ßòç>–ì&¦-‡Æ¼ó‰2ìSùYŽ¸ë¥OMß`HX´2àóÞ!/®|ýÏqi®ÏAœšzxw¼I“á‡Ÿâõ”·<ƒ/×òoå+ž#ÿ€©ñ~| ½÷ú^TÙöâúç§ÁÀ+^R5EØÞòVd·„-µ•d¶j#¹×;`)Ì,ÜRÅ(&²Wr4üÏ²¶Ý‰wÚoÉÈ´»"ÄeßRÃW£¦—#K³©úßkÔÐ¥†šJ9’·È“^u\Y/Öi!—e^YÅ’/Íê`Ø˜Á`+!\¹NÝhÖ«•Õ6o8cT_1€¯1*gjøcú¯llŽ5ÕWmœ‘ ßÓ`R]Å«šGt˜VóØ“Ñˆ¼TºŸõ¨ù1"‰úã¬aœ¼ÄÜ–‰?5>A›sù/`Òœ6X¦\¥æ¹K“Ùè‹aŒ%ÓÓ8…d‹é)qK2<ìÒzé¥ñõÏýéÈ+æï±ƒ?˜žßð@nîƒ3ýØN“W©¥eßKôsöM¯R\Ò~[kÙÆ9·¡FÖ×F!ô3_Ÿg@½?okt‡*au&§ØøRg$`5²V—%ã‘õÖí“0ÌIUnÂžˆWð¥jAx›e…p_Ê>ù,EúîqG$VÌV/³o1>±nè÷æ­-móSnkpá0´ŠùÃ Èÿì°ÖäýKùâ«ïî_Â¶Z“rùÉtÔ,Ì{Z_q>™2ÊIk1 k˜í6ìá}ŒÃñQ¸Å‘|ÿ[DvÚº„ÑØß·ÈöXü)ˆ1Çâ»Ã›½2&(ÔÛ% ðtìþ@¤ø¬\<AV*y#áâ­îåq^rˆ— A·vvÅ¥2În¨:©­˜?è6B¿Pk-ç‘„” aß]”ºÖ!*ŠÓ<m»²=Ê9üÍZFÈ¼’œ=âý#<°ª55ŽmëèØV
Ö—·CvÝéÒVˆ"vEž<!#ÎgëÒæ™T>Êçkú–ÓJk)/£éÚ0¾U¬IH%åùCÇ¬µ.xÕ×—ùÍ¹"®¤²¨¬°AL¢ðúç÷þ¨8Ò®8üÉ‡är%%Bµ¸‡<‡ƒðÔì€n©á“5Y9CÅP–tËY+ªùˆXLÇ)Ôõ&nf\%Ts¿U‰ôtçS²Ã}[*GG•«7º(ùp+*™*|ì‚é¸a²ÜBë]ÖØ”ŽI¨'d/£’µâbŠŠ/‹çrSÓpÂ@QÆù¨¼uâk×nÃÖž‰DåÛw²ZçIÙ°·¹¶§ú5ÅY©ËCê8WiŽ“3:L6ÌÉõAÞÊ³ryƒ9Òøš‘°$QdèÁÕjHõŸ¿í÷¿{¹w|Ü{¾wL„Mõ2€DÔeHiÀÎÿ«MÈïxÎ¨Ú+¤X@>uÄ~ìÏXe>k$g’É¢È¶F>B
Y O<C¢kLž ŸŸ›¬6å‰”VöjnÊ«šg‚+½:÷:“¯Í*¹QË¨ËY¸È|“|AŠî_““µ°LªTÞD‚—œõþ¡ãLÁ‘Å]å¢Üª93Ž¦IB³î¡¡8¡X(•G¶sÈ–©µë¹²0ËëlE»TÅ}{æœê\Ü"µ¾ŸžÇ·É¾æqH)¬Ui¤sV |X±FˆY–0Q:	puÏÑ\ª¦ò 5UO³$=i !çÚ\"‘h@¥á8Ï{'EéØ°>j¹ñDÀÅûhÙƒ³¶ˆÔ8M<ïgN‡Ó0N³8àÔDSa‡Îa5Ês)ù«ìÌÈn²DÙòÈüd‰¾j«V•Ù¤HÑõ:[ï µÝs?êm[q Íu†…SS'Q’yŠ@ôÁ7TÔª7ƒãÀã¦ÕsH+'%î)eÛ½þ'JÿKzûéš ÷î¯$ÈFB+•ñ7G´_ƒ¬/•œ/­½‹þhÄêvn·h¯ôÈ‡alÿV“£HâbÓÒÍìO€HÒl.«7qæ|«%Lß¤ù	{S;Ñã¶ªÅcaòï.5õÊ«Ö8í<’ÒH°%¾’t³kˆoJ>ld­kÖ¤ —]~Š„‡ä}^Ê¨Aõx'bÐÍÝþLæD,Mî{ƒ bÇ{¹‚·|Í];NÎöM;!£7BÇ–rW®sõæÍ“Ëw¸$Lý€!sZùBž·[…¯9óœ¥0záìÌÎ••>%;t¾*ºAeýñ¥mUÌ–Ž¥ÝÜ!€JºæÖøÎ¢ómFÇBÅÈ”¨Aš5®ýÕþD_ ÔçPDï)D—ê ÷Ÿ_LŽ+$‘œ+æŸx3Ê¾8®[‡ÓYy­›€ï™ÁrÖªµj[(X5QšKf)/]{»ªzQ³öVôVîÅ!3º2‚Û= lŽ?¨ŸX¯jN3ý°K³«ñJú
½’¾ª…¸J—+’[™,ù­ã§^œ9z?Ø$˜ÜNGÚ²+ª?ƒ’±Ì9Q«VPûáTÔOÉRDýë¿c²SÇx>rc `³÷˜ƒŒMâë¼EXŠEÏãdä’-ýÌÑ•õvUy-ŠRûL‘èêR¸Mf2­«-Á1†iTIius1˜2ÙCè;¡u%i“Yçœ3ò¤ØY@¼²ñÕ Ç™Â÷AzˆŽMcItóòhSùcø	ò±?Ð*||9‰€f%lP…ÿý^·2ÉI6ªÀ¢Ù_“¥Kç€h56•T³•2#ãSmÞ®xžïÃÀc{bùÕ7êMG_c7Àmu„PRcÊÝö~œ<»v<Æà¬<*6Þxý÷(YŠ\øï$]ÿœò»ùhUžwQ!ƒ°sÞyó@$ò«Í©\¡¨é“Æ»á™Üm“Ù‹\!7-eV­KðÆ`ÜÖò\HUz"2NÔ&­šåSÁŒ'~jáÄ‘ó†ÚTAÖ:0rˆf~ª²‚-e—Ì\ä,Lª‘)½ê	'!âXLX¹œ¬ãB7Jæ;y‹Þ•P¦)ryfk"ýìë”·úUyrº2?”Y\5˜opÌûãÔÃ-hKp(X·Ï¢Óï}žqsˆµX…CG1".£ªjãUèD2høÏ–¶–Taë+Fg¬mœ¢™o6»ð7KQ$IÇ¼_Y‡¥Îá†¿‘Z2–,ãm2	BVc>ZßpZ}¼­It©½£B^/F­owŸ‡­°>F·E`Œ0PpS`¢[˜Oí­õ J€d‘ìéOêÐrä¤¦“Qä0Þ„/Z—æ!)pr~“êxP88¸Ž¸. ú×?uF	õ%ã\KÇ'IKñ$
8ð"Ï‡øªf§Õ—É¥hÝ’_ÈB«<×Q&‹Ê…ÍÓðFçáŒ3FÚâQ^¿ïOÒ'­îd0Ôý®l’ÌÓ¢y¨ð£ÿF>,ã5?vø5ïtäžTW^¸ŽÓuü¦Fè*}Q“ÁÌB×(:û"mAÛdŸˆÙØ:7þÂÊÞ/Dîª"º>óÇXêuìÅ}QZÅ”á÷ÂþÄBÆy8©N
þ¨™F¼^vÂ¦Éôú§8ˆ’ù…1Ja×ÇŽDFf×E}LÙ¢…:Æ^ËMŽ„SkvÀõ7g—­þ‹¾R¨šANhž,{à%çþ ta`¯fcGX]W}ïëÛ dƒK6ð¦èR¤„G>ñAEL¥É4¼•÷«åQ†{HÍ0¢7HFánì@ÿ,ÌÀÄûÜ=6kWziŸƒ{>,Œõ%¸eNÏ~›&MH®“žN&ã"MãÒ-Vx9á)ýâyÜuƒ{Ê¸øŒYñ.dTAÊTÜHZ%úê‡Öm“aïJ©€ô1xŠ\O$÷½‘>ý¸?hs*ÕÁit8Ío­lw±ÂõäÃÓ¦3;÷Ú9‹R›“^ø—ÌGîk¿Ö?Dâ8ûîV FêC\°B–9àÅd’ÑÂŒÖ:`­"±–ACÇýYÅ0l—„jqá'ˆcÈX„0l6Ò×’¨9K“h-F•?UüÊU€§Fýêv™éÂýK6ÈâÖ	t§Ç^Âe£$Á£\ÌBÍªÖƒÞNÑ™Öœ !dÓ¸`zï},YF,ñ§h#¬ ²Î¿È`ã!UPÉ”8ÒÈ…‘a~=±'¨ˆAƒU8€É$[ìðÕóUöûÃç8ÿã?=gí—ÞE—=zùTçùIO\¨÷¿ºD,y¤0ð?èfk\§^?e÷¦mä?kõ˜¥×ElíŠº‡±ï„¨Î½}íG™1ño¥¢í€ ¯hE‰ØõÏXh@‡g%Ú¶øM›<rm:ñÓýòík£þÕb2w9XQÔØ¢Í©ÌÍå"Àe \—;^Ôá/ARP‹fõ8€­yQ {;œŽFÖçsåv.of^¨ˆ;X2E”hw±hèZAJÐŸ«»3hÎ`™¬ìvßº247ŽFv˜Ô@¹ÈÚlx~¶DglSe5ƒœZqç-¾ûJçâ[Ö¨ðoy|¡jÏÛ¬ùÿVäD	ÕÓŠ#yØ/¹ê!W†Ô¼ØÌ»ZÂ w°j`ÞüÛúÆÆÃëot-FÅÕÅª©RUqæYô«º*ám&‚Z\Ö ï‘)†ÎD˜m4ÎÎS‘uôA^13¥®mTO7ÿeSå\f\Ö›ßæ1“HíË¹r€ÅWÍÊy¼ÿòõ‹Þõ_ÿ×ÛÝc/žï¿ÒÓnâƒî]‡k	ÞÄpó¼QçÿÐßþ·ÿ÷Øûncr±ÊAXaŠŸWÞ²7X!ãŒÏ·ƒ|ëÖÆW“‹ïðŸ·$à~~ÄüHz;\€NÄ	 wë˜àø£Èls…2e1*ñ)&¶¾öB8$šÓ\ã33ŽûOÔ÷š}Ñ¼ˆÙìÍ™ÝõÕéyn¡ZWŽ§5>Úõ ¾Ñ à¡:1Ð•Ýñ@{~³.slqxæ7Ô{Ð×þ³ûC ~|wöñI+Œ:ÙW¦N!Æì¢d‘¬‚b°4ê¤1ÆÑ¸“üðQ¨l£ºèH‘ë{Gg6q’ÇK [WùõÄ-Ç5Éo×‡sTA˜òZ«¤ˆ¼ŒiÑ“s‰ð±I¬“édc×Z§MâäSÜyÜ9ªÀ9}/ó^J¦^‚ÛÌÙîÉðÿ  ÿÿ tçvxœì=ÛrÜ¶’ïþ
x*åâ$3#ilù2Ö¥INtŽl©$9•-•Ë¡††1‡dHŽ.GÑìÛ©}Hå!O[çV?¶Ý¸ A%Ù»9»‡•ËDîF£»Ñ€1Ÿ•…tí©áùÚ£j9“qäçù;JW;yêiÿª?$iÆÍúôœÆEÞ“˜’¾†|]ìT»¨6vÑ?E™ô_’“³~ùí¿\\$Y2‹ô£3r’dÍÄÿD•‹‹²ð¿ÐA["©n_•Ý¾0»m¤Þ2™ýgÜØÝô²è)¶OØïã—éåþ“#³Èœ&qÑŸ&qBŠÌ
ã³þEˆ¤`}W¿·}x´M¾ÿîhû€ü×ïd¸8|îF¯±È*XYÈé¸“Xuoäð®}øpNü²ººJ:QÊ;äÉâ•àZ>üb”¤ýçŠäO/#’Oü ¹èçS'Ó—€&Sÿ²ÑBÝéeßŸ‰A›Î§½$?Ïò"<½êŸÐâ‚Ò˜@ÇÓ\RzR!åÊäi…WÐãÆI:›^®íâ°Éí¿4ÌW&OÍæ®iÀ;¬e’Ftw2+Š$®p3‰7£püiõÚë’Õ5rí˜L9-¶ËÆ½DºûÚ]ï½`u®IœLéˆt:=2Ë"þ#JÆ>ûInð77•"›<œÏ€?“äœf#­ð9ò×©NÂÌóå¢?N¢$Ë;VUéÞôã1üŒÀo»ý¤"Úœ&d÷¦^²¡à4É¦@éÃÙÉ4,V¯'~Dt#V7IÂÿ ¾$ŸŽØï,¹ g~
ºäè¹-GaœÎ
Ë4í<ÐlµóXa<£¿ÌÂŒVñ¹ÍèêuÌÙ8@&ÚlüÏ –G™Èl¸'D€
?;£Å€µ¬·´†ß_"éeÿIqm‘³&†kª™úÏòQ†jk¨¿œ@oL¿%³4åªUÈˆIŽ…»RôýÁîý
sáþôdéÿ"9wQ1o/ÃLõ£®M[‹ˆL_Mrb
eôOHN§"/®R4Ê˜‚±Ég.‰¬g¥.Åé×±ÍMú‘¾T1µe,xr5¯L{™xÙª×XÀÖÉÊ¡NIþx9\¼*YÙf¹Qh²Ë¥œWPû6-ß¥)‹Ëþ¤+ÚI8þÓ…+f4ƒ1X«.³ES?õ¼ˆ	\eÝ…Î*‹Ì'zµzÂ qÍ«§ª´1R`›2,—ç	$Z6gP'¿:M¹0¹mz c\]²§n‹-kW+HRç²ÃºXÀ>\` Äµ‚ÃòiBï2'–eÁ­aÅÙ”ÌÒ”fc?§h
/Õà[vîB¹	éŠ®qÎ¬ØÅC2¸y Áæ>ŸÂ²tÉœÂ¢ØÍÌkŸ#>–Õè6EEÛlä" V¼WM¥°#]ý;I|Æ$>­J¥šMÓ+Õ$¸n¦:´ç>n‰YAº-¶ôò¦ÆƒuÚ™â[îøùU<&Í<ò/|À$ -èV2öü÷¤G6÷vw·7vöÞvwÞýõÈÚª[ËÃ$>Jü¼ð:L{gtšœ‡A2è|ödIn±G”|6öe~>ÞŸAÎIì(ì6ywð‡Og13œ²0©¸ÅÁ4Œßh~é…Ö«`Üêµø¡c4‹Ã_fô}~@óÕkó½¬¹0ÅqFóq8]•lÄ3?²ÑÜ´¾3´5Tòñ„NìÃ«<ÌyÁ14p$øÿ£¢Yo‹	<Öä¿êê¶Ú³ÈIú7å'$:™å`ë^§YrFôñM•ômú;ñÇŸ`ýnü(ç…þ0MV«EÌ£U¸‚5ö
«<¾e¨JéÙææío¾!þfDc%uÙe'Ïy'¶7üÍ(¥ÛÏÀ;OÈ_÷Þ‘1 Iã !E$9@‚ñ¿	Lò–QŸ¤~æå
I0ÑÉ4<Ë|8§Mve³3_‰9•ºRç<fÀi8ßGQ¡/“`Î¥ÍÃ˜>ÀS%+CÝ	Pbb;óLn]_Î3ÿW¶’‹8JüÀ0ïíJÌ”û™'—uåÖÌ‚¾nb)
iNQkª¸\£öT©Œñ}ÚŽ4/ü’b¥¹-%Z¬3F5XÂA-,~¿O6&YQ°È¾æ3¦ßàû#Ë>…¡©še-¬´Ÿ%iÎÌ°á`J¢vÜ'4‘M­ÀOŽ? å€*>½‡ÿís½ÉË…Æo:"yuë#ðxLsÐÅ`Ò,K²3€Î“0@@¡}×6¡.Æ4Ù8éešdÁ%•iT÷0<ÇÈ*9þÐ“H«7j¯ì¼÷èfÔDž.§Oç9Î©Ÿ'G4›öÐP>T¯ èé° ÉëpCKÀÐi
Ë©ÿ&Œ€l[/Ñ!W:h!±¾;XZf?ö_tÖøÚÂ>ö¸”Ü,×?Ô
Œ¶Ë6°Ý} Ê>ï„Žg¹øâÅÂ¿á@¬yBï1£gÀ
­Ç­À¤ƒ U°g´Ø
E>2ÄÔ3@7+epÒÞ5[9<%ÞcYÇ¼jŒÛb–Åd‰[¶¼l=†)°…H(¨oHçhqqÄþ‘–0¯Oõú¼á†ÚAxzzÆ ½õ‹ÉÀ?É=:€¡a!ø}’—oÈ-ÿ*—cFžjnx ÐÉ×ä¹úÏðY·¨ˆÑ‰±† /ßy²±.Y'Kd¤Çª7ÍSÿ
%ú ¹È9oÞ‚;àéŽ¯9‹`ûñÕñ6_êæ¼€‰µí'žç³)i84ŒM¼| ">sæÐ4)Î4õ‰—Ñü>fƒÆ”¼”>Jcåô=xB¦WÚq#¥1HY	LÀV†‘NÁ² cX±¦ÄÃIHã±ÏL¹"›Ñ®jÁ«¶¯ÚZc¯JŒ  9«J³½“Ÿ9Í3@8ŒÏ›1TgƒY°nö¯-p0Ï}èh>´ÀPÔß	ºLZøa„œu™©µ_;$d“îž
å5©§}š$à˜È³²ì¦Øñ„«¦ý³‚¬j`Bnâ«T¸T€Rüªk¬×›´^Ð"ß8CiåÕÕû7 m)Tª'OQ«ðwâ¤®G´¢7á´Â´s?J2%íŠÞf1ŒcÑdtÜ(f€4"!86`ÅðÌ¤*kÐl×RÈì ˆô¬«·ò\«N£œºaž0¯ZÁ¼4`@¹µZZ4¡–+ƒÅªHŽ’Â^‚}_+¨jý$Š|°²q’°•GÎ›„(v;£öñ‡.‹ã[Ó'ßÌœz~¸æyß&`:øqE$Nƒs°
Sø0âj ¢7»>`f“à7[?o3™ÒlâîbeTø,M`´ª¢‡uÔ®Îé,Ÿxú„æCß	FRzÂ@ŸÔRsÁ1åÔs))£¢€MÏŠ~M¡q?ÒUâú †"ô\8Xa€÷²Á	¿?Ib»ËX‡tJÆ4BÏÞ^Z°ã	pŠ8þžƒÃû8,¬f@20¨–þº^GUÁ.³Â[ìŠä=F±7ýM…o,OÂ	œƒÇI½¥®Qò>¨Æq˜ú‘±˜E­Ï4ýsž@aáƒr}˜â«¸fÔqN%Üá5vÃY1±†¸Nj!?'aìuz¤Óµ€F é¶@gÇÞ^oµ¤êtùØ+H±Ô‹

åê=oV»kV¦7`¬×dl¨éœÉ3v%wûwàœ!E˜2N‘”…öuRKÞˆTWA«",„#kaÔ«H¬—•Ê»Rs/ƒiXv«•±UÍZÓz•‹½fŽ(½(ÉÿCšëHf0÷È±eèö¸˜|à6˜”èæÜþ,É DÚÞþqNÃ=oè˜£XØŠ$—…,Q8Ôr/©ÙgJ¸"‡´XáÆÏš'ÖùÒúÍ*öw6Ð•VÛøA`—;	²‘eþÕàêy Ö•”Ñz.ÉÁ¢|¨øK–¯‰4hp8DoÆ€øJêeÉ…1"ÞæÔ/Æîlk¶$Ôè‹(×ÝäB*×hÇhÐÜ+v³†1{°197Ü¦RïÜ–²^´¡‹€6v#¨Àõóì×™#êÑË3~6 ^;¨ËæÖƒX¨ëÀš…¬šçè‡O‹Zô ®]rYHLˆ•.,Ož˜”ïb´ò•#¥O	Kì1ÎQF~ìˆŽ‚ÑQû ?Âü€úÁ^]‰ÑËèW©ûöv· ¿?nî½Ý>ØÜÙØ•¢UWùpcwããÛ£ƒÍ÷»rr¾½ý”ÇÌb<ÃmXó"SŠßê&'úªçMTÔPIXíª\Æ‡\ß„÷åc1&ñ‰}ÓU…¦Ò, !óÍ*‡r	y­+¾ÔZ]Í·µÆRÛôMé«Ôö œËJŠ óÚ¯ &¨U§¤U×©U¢õÍŠñ¥I«*Ç §æð,Øà·ðx9môÆ˜¥`"Ð}>«¤¾!bcœµQ:'""(E¿ôIôœ£vÁL8-åªÈ®¬E†wŽ^šñb†¢Ï€ýa'î”Á-ddG²¤×šeÆÀçÄAX(æÆlƒ“0Pm”R3ŒJ3èX£ßi(ÉÏª}ž{ÀÛ¯Ë=ØÜØÝ~·µq°³÷qcsoû°WR¡§¡R;¢‘‰¿ÂAÑPå-ˆ‚q•ª™†f–˜¬ãdJòî,$Ñ_ÛÒª‚V'Í²rFªÆ·³Z+ÌÆFwlWlYˆV™U¾¸Ív(¶/ÇÌÁ÷%-ä£„C@ec*±œ5nÍiþú¶ô¸3¹˜—|e‰Ä¸i*	ÜÁjºATÖ=¢=entYƒùÕe•ý¯áeùm+¼ý;z•X¡â`v¾g‘* Þ†4„“Ÿ¾ºÎxë†°
?•‡šÇ	-jÖ’Ö*3·Á¦¶Ú5ÍX­þo¬rYSÚsZ­-`ÖÚûüVÍåÞðŠoÂü/YÇpn:? wB0Hå|Õe•ªîN‡¯[<V¥ê¹\ ÑÞ^vûJ¸Ù&w‚´ÊåüPÓŒÕ¶¦^9Ãl³ò"É>åÊüwÌ
èüœ'ñÇ"ùÈ>yº¼v_WàO’ä“	Ž%Á}ñTmû£Ÿ¦ [Eûj<²µžQÂ*é´tîv*ïÖìó"
Vmêå'ÙäÇ²½_]«½ Œ×ìî²5Þò4
AuºÇ‹n—Q~ùS¥{¥ghä·ÿ	ëáû˜ÈXæ"ruñy´˜h:ƒ5Gv×¤»„À©R—z[¦q\/|M¾³ló¯d¢CsfÅ4(3+à·yð&ÙI³vòÅŠ}Zc2¬¤žÍTyNªÏ&Eå Ž qËÞÝÂdØ2ñEÙª$½äj.J"püÃ@ÀŒÀÃ,!ð?àú”-8åz¡¨újâ“êÙf²@%GÅÊOÑ'=I¥UR³ytJ3?
Xj†Jº…/jó5.ûËÕ¤òÔ—•¬!iœÓè´ÏvoQlØKƒR1Ødå«Tòy’Æ¶œ,ÚÜdä±Ó5ìTó„ÎŠÒÅZ Ýß!›~ä“ä,Ëÿƒ“$ï/aRù:$Ñ™öúÌ9-ôæR ©J+ªË.wŸ²S	ö
\ÙÎ„qÌ-ó±­“Š'Q2þdM
\ÎÁvOr²Ñ3<3ÄÜN—iH²>>ñ”FmÚŽ{,O¥Ö€—Ì·GôÒ‘-\ås~ìÃŽŸvYí52„©î¸dž-ýÓ‚?=>–ï*Ü’Lé}þI¤ë¹C¾*Âð›áíMxIoØd”¥ò:lîuºó¤æŸ…ÝRÍß“áe¤eÞ?·µm>¿Õþ_0¼<’ÞÌîÎ¶µ˜ïVt¹(üI8_ð©Âw6/Às—  iñ†í5w” ªPÆÅ]‹	Î×	†AR9´¸iqîk„°Ì®yFX‚MVRµæ…çÔ–¾A`q^ƒóOò$š…#z
B4X&E’ö—†¤ÏìUF‰+V`-îV›Ê5ÊÛÊå‡H±u»Mã¸î·³|-¶*ÔQÌ•×‚-ƒÁ æ,o¹»Ñæo™öê™v›ŽëŠË7ÒäÏnò®ƒeãÄdÝ¹]í$e)šxzWM„º“½¼§ÕœpÑ¿áC;õÛpÚ·zÍ‡í¬²ÃÛ¬d66ŽÚPÚÈ¶ˆùq~|Õ†æ—£¶\r:]Ê ƒ:~†Î-'ç—b†Ðó$ë‹KZæœ`HRÏç¤g›ŠkG	f™Á?‚®9ñ0m	L’ƒýÝ•2§%–è´¦"³øÚ’år+@xsÃá ”¡Ï*zú&e«9®Õÿ—à}Áã'Û–Â£ö¸ÖÅ¶—,i+}r_L5 KZ7 ÷Ò´DÑÿ $ëìm$Y?žpõê_"léÎ2JÛ™©¡k#ùIìr±çõe×ìN^óÆäüóÆñ­Ú^bÝªÍ^'ªÎ×ªÕ~äŸ€­iÅúL&|i¾«[0&a ³Ÿ°6¸M'ÙÔ×oƒ¸vlÖ³ÍëE²n¾ês.	³Öˆ_õŸ¶r¹Ê!{ÅÛÜqBÞÑx2›²ŒšÛ½¿àæìiˆ_›S-o0¹ÛÂÖNçÉZ|6iV@ß ¸0o-<<’ÇIdŒlïˆeÂáqÓ#}MÕ~äÇÚî$¡€CÆUgRøqkvú?ÇBüŸ†ÀÒJk^gSd™uºó¨cIq—Œy¯ÌI›•Â?‰¨CO
ñeîi.Nº¢ï$i_=YYL¨XSBê×…RÅöOªSƒßž·Ä®…¹sxÚqÃÆJ‘9ï€é¢Ï
±ñLq+Åän€ÒC»;¤ÜÂ¿;ä&&¼c–û]AuÐY“Qœ‡µÂrî>=Ÿ™xo¶~èÞ§–QpwÀ±ÄÝsîI®·2;k*Éà­f‹s¶¸‚ŸüÑÒ\?|h"ßá!2E¦aêüxÈ®‰ªˆ°»ÕÊÜz4PZÑXTWU¡$Á•Ž/¨VÐ8ý+"~X¬¼+Q»æÈ1€kG¦PrÑ#apY{ŸŒÈ8ÒO¶:oáõ®ÀÅ³Txpõ¨óƒ÷¦úSÀL	"2#Ü[,óØqÕî õ1É++¼atÝÏÈ¤¨+«!Žì]Ú(À6½ÂS»«˜w„ƒ¹éuÍÄðñæ§zè0ßƒ5/€ÕxÕ);˜)òÈJÙ×“'µµ›“D_?r yöë“³\\]†¹VÐ«LÆÃ!YÓ!+i¹ùÉ}Óîø\c7øtÔ®¿Z¨—÷zöj[(‰ýä	Ñ®EY.×B9næÂÇ=Ó¹1.¾KûÛ~VŠÀ­Zê¶3`4²j¿¶&'>Ø3:—Ð÷’²ù™I¥[95ÆL¡ñ±ðØU#žB¶¤³,4]%Þ_TÜö3bð|UóW„nÂ¸†ÑøÔó¿rœõ-ÔßüÆ¿ã
iÑH”:`{_‘rE \IÖ¸ô„ÎÚŽç!/Ç{#R·Å”%–ÞUe­6ãZ‹Jó÷òy¢Q6‹ÇÀá/™QÛrF!ô¡†—u8¶%V}sI¥öUïÁUÝ€j£ñÒ"÷÷ÞóLá‡aÛYc¨ðÔâÉÜ.µí¯{t,:Ó2”oæö(—ïÓ[ó‚chQK:,K¹¢LS5ôeI(æY1­*â!Ÿ{ª×E[+S~¹õ”—Éä÷œïóÀçOvt&¤¿uƒ¢IÔæÞn*0–Ëbäj©áâÖ¨¹†TF*m_æe÷;˜šæ%;jÉO“/ØÉ•þÜw¼Ú(ÍÁè><>yü
9ÑlÂá£læ¹5™õfÞËY*÷Ð#bENçb7‡ºó8é·¿—ìëËòrn…kÃ?ifä<' ;$­ØâÇáiœÎðô_>÷÷S®mÍõ¸’ÙíI€óL´ìržœÖÜÆ<†¬lD4+6ÃlÉû¾—–kî\Õ W|8×È”¹z±<‡s_»‚GÉJ5ÙÆÂ1”ZyäçÙaw›~zè³Ž®Ì…Óƒƒ¦c¢®¿ŸO…ò@Ó—¡ƒy6ê3Q¢-wùá«/30G:ÀƒÍã,ì­oÅü,Í.™[ê/ñn\,4¢ uÛ£ýÓ°˜cK\Ï[‰Ë¨\O;9ÜÓ³‹ÈX¤°±1~m”—W|ËGFÙvÅ8mý…
u×|ËÏØŸñ&ò &†©?ÊwpÆ›^™›jFàIO4p‡1õOD—#ìÎ¥‰¼Ï­rj¿…á¢†ÿÍüñ3X§™´ƒKj¯½3îÚ5fëò&€/„tyêâ(!Ü”qs/Ô*!|ÔZý­*+¬=\Ä¸ö|ŒÛÍ uíX”ï6!´¡’Žú{°T‹lž8a¡¬ä‚Ç¶qBšþàù°í‰¼!|P>æß±h%sÚ¶qÜA1¼üÑ·Xæûl%Œ¶Ó,/ÚÕl£dÚÐ>sLëV¯KŠ¶Ó¼^)©óÁÚðöºº
´"x¿¸x¶«ŒwaXZ»%Öï2ÒÌ‹ù4jús#úSû@øSýûYf'ŸÓ¾KŒ¤)¦DÊ…D».ƒws6ÊJÃ¥ônJë¥æ®ý¹áÌUöG2ˆ	6Ý8šÝþø=ýÚ’0/=âòøó{çCãOª!ÌsÙEübmh•³Œy.|‰y˜—_óÅ™›åú“••ÞS2ì¤d–ÃÖ”"§r6Ëróþÿ  ÿÿ L6ªÖ