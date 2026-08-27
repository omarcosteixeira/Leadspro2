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
  SolicitacaoManutencao
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
  REGIONAL: "Regional",
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
    ROLES.REGIONAL,
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
    ROLES.REGIONAL,
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
    ROLES.REGIONAL,
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
  checklist: [
    ROLES.ADMIN_MASTER,
    ROLES.REGIONAL
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
    ROLES.REGIONAL,
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
              { id: "checklist", label: "CheckList", icon: CheckSquare },
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
              {currentView === "checklist" && (
                <div className="p-8">
                  <h2 className="text-2xl font-bold mb-4 text-slate-800">CheckList</h2>
                  <p className="text-slate-500">ConteÃºdo em breve.</p>
                </div>
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
    metaDia: true,
    aniversarios: true,
  };
  const isRegional = profile?.role === ROLES.REGIONAL;
  
  let widgets = profile?.dashboardWidgets
    ? { ...defaultWidgets, ...profile.dashboardWidgets }
    : defaultWidgets;
    
  if (isRegional) {
    widgets = {
      ...defaultWidgets,
      stats: false,
      planner: false,
      qgLigacoes: false,
      periodo: false,
      bomDia: false,
      forecast: false,
      metaDia: false,
      metaCursos: false,
      metaSM: false,
      aniversarios: true,
      links: true
    };
  }

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
                    <th className="px-6 py-4 flex flex-col gap-2">
                      {selectedEntries.length > 0 && botConfig.url && (
                        <button
                          onClick={() => setMassSelectorOpen(true)}
                          className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                        >
                          <Bot size={14} /> Em Massa
                        </button>
                      )}
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkDelete}
                          className="text-rose-600 font-bold hover:underline py-1 px-2 bg-rose-50 rounded-lg flex items-center gap-1"
                        >
                          <Trash2 size={14} /> Excluir ({selectedEntries.length}
                          )
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
                  <th className="px-6 py-4 flex items-center gap-4">
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
                        onClick={() => setMassSelectorOpen(true)}
                        className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                      >
                        <Bot size={14} /> Em Massa
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
                  >
                    <td className="px-6 py-4 text-center font-bold text-slate-400 text-xs">
                      {index + 1}
                    </td>
                    <td className="px-6 py-4">
                      <input
                        type="checkbox"
                        disabled={invalidBaseIds.has(entry.id)}
                        checked={selectedEntries.includes(entry.id)}
                        onChange={(e) =>
                          !invalidBaseIds.has(entry.id) &&
                          toggleSelect(entry.id, e.target.checked)
                        }
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-900">
                          {entry.nome}
                        </span>
                        <span className="text-xs text-slate-500">
                          {entry.curso}
                        </span>
                        <div className="flex items-center space-x-2 mt-1 flex-wrap gap-y-1">
                          {entry.telefone && (
                            <span className="text-[10px] text-slate-400 font-bold">
                              {entry.telefone}
                            </span>
                          )}
                          {entry.cpf && (
                            <span className="text-[10px] text-slate-500 font-bold px-2 py-0.5 bg-slate-100 rounded-full">
                              CPF: {formatCPF(entry.cpf)}
                            </span>
                          )}
                          {entry.semestre && (
                            <span className="text-[10px] text-blue-500 font-bold px-2 py-0.5 bg-blue-50 rounded-full">
                              {entry.semestre}
                            </span>
                          )}
                          {entry.periodo && (
                            <span className="text-[10px] text-purple-500 font-bold px-2 py-0.5 bg-purple-50 rounded-full">
                              {entry.periodo}
                            </span>
                          )}
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 text-sm text-slate-600">
                      {entry.nomeBase}
                    </td>
                    <td className="px-6 py-4">
                      <select
                        value={entry.status}
                        onChange={(e) =>
                          handleStatusChange(entry, e.target.value)
                        }
                        className={cn(
                          "px-2 py-1 rounded-lg text-xs font-bold outline-none border-none",
                          entry.status === "Pendente" &&
                            "bg-slate-100 text-slate-600",
                          entry.status === "Interessado" &&
                            "bg-blue-100 text-blue-600",
                          entry.status === "Convertido" &&
                            "bg-emerald-100 text-emerald-600",
                          entry.status === "NÃ£o tem interesse" &&
                            "bg-rose-100 text-rose-600",
                          entry.status === "Sem retorno" &&
                            "bg-orange-100 text-orange-600",
                          entry.status === "Contato via Sales" &&
                            "bg-purple-100 text-purple-600",
                        )}
                      >
                        <option value="Pendente">Pendente</option>
                        <option value="Interessado">Interessado</option>
                        <option value="Convertido">Convertido</option>
                        <option value="NÃ£o tem interesse">
                          NÃ£o tem interesse
                        </option>
                        <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
                      </select>
                    </td>
                    <td className="px-6 py-4 flex items-center space-x-2">
                      {!invalidBaseIds.has(entry.id) && (
                        <button
                          onClick={() => {
                            setSelectedEntry(entry);
                            setSelectorOpen(true);
                          }}
                          className="text-emerald-600 font-bold text-sm flex items-center space-x-1 hover:text-emerald-700"
                        >
                          <MessageSquare size={14} />
                          <span>WhatsApp</span>
                        </button>
                      )}
                      <button
                        onClick={() => handleContatoViaSales(entry, entry.nomeBase || 'Bases')}
                        className="text-sky-600 font-bold text-sm flex items-center space-x-1 hover:text-sky-700 bg-sky-50 px-2 py-1 rounded-lg ml-2"
                        title="Registrar Contato via Sales"
                      >
                        <PhoneOutgoing size={14} />
                        <span>Sales</span>
                      </button>
                      <button
                        onClick={() => {
                          setEditingCandidate(entry);
                          setEditFormData({
                            nomeBase: entry.nomeBase || "",
                            nome: entry.nome || "",
                            telefone: entry.telefone || "",
                            email: entry.email || "",
                            cpf: entry.cpf || "",
                            curso: entry.curso || "",
                            produto: entry.produto || "GraduaÃ§Ã£o",
                            numeroOportunidade: entry.numeroOportunidade || "",
                            semestre: entry.semestre || "",
                            periodo: entry.periodo || "",
                            metodologia: entry.metodologia || "",
                            formaIngresso: entry.formaIngresso || "",
                            numeroMatricula: entry.numeroMatricula || "",
                            status: entry.status || "Pendente",
                          });
                          setIsEditModalOpen(true);
                        }}
                        className="text-blue-500 hover:text-blue-700 p-2 hover:bg-blue-50 rounded-lg transition-all"
                        title="Editar Candidato"
                      >
                        <Edit2 size={16} />
                      </button>
                      <button
                        onClick={() => handleDeleteBase(entry.id)}
                        className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                      >
                        <Trash2 size={16} />
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
                      Nenhum registro encontrado com os filtros aplicados.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Editing Candidate Modal */}
      {isEditModalOpen && editingCandidate && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden my-8 animate-in fade-in zoom-in duration-200">
            {/* Modal Header */}
            <div className="flex justify-between items-center p-6 border-b border-slate-100">
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <Edit2 size={20} className="text-blue-600" />
                Editar Candidato
              </h3>
              <button
                onClick={() => {
                  setIsEditModalOpen(false);
                  setEditingCandidate(null);
                }}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Form */}
            <form onSubmit={handleEditSubmit} className="p-6 space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome da Base *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.nomeBase}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        nomeBase: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="Ex: Junho 2024"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome do Candidato *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.nome}
                    onChange={(e) =>
                      setEditFormData({ ...editFormData, nome: e.target.value })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Telefone *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.telefone}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        telefone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
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
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="exemplo@gmail.com"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    CPF
                  </label>
                  <input
                    type="text"
                    value={editFormData.cpf}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    NÂº Oportunidade
                  </label>
                  <input
                    type="text"
                    value={editFormData.numeroOportunidade}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        numeroOportunidade: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Curso *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.curso}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        curso: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
    xœì½ËrI’ xçWb³³#2 ÀLfWb@ ÌBI`	0kwQÒá <éåîG¡!²×™ÃfÎ+³µ½"-Ý"%2"-sécóOöæVÕžj€d1»Ò%“7·§šššªšš*côYK£“8eÃ4*Š×Ñ8^ïœ¤Ùð+ã«réª`§Ù¤\:ÉÒ‘H)Ò¨Œ—ž¬¬°ñÉÒjçÙ#xöól4+3öUàëÚ2o0Tp­ˆÓxX«¼ˆÒY¼~’òE–·£2êOE3·ÁÙdë<šœA™nÜcëáŽ2VÄå©³{S‘±~¿O[_¬Ì(»5`q¿Œò³¸ìóÎ³¨`ÑäºªÜm/ø!<82[—K§³4eÓ«¥oÙôzéqÿ	Ë³Ùd–®Rv’å£8—ää=†És9†ÙÎŠA6+Ód/M²I,“d‰è8Ÿí“³¥Ëó¤Œ;þ„a{3›$¿ŸÅŠþ8šv»Sœ
Ö­€ÁZ6-“lÂ>Ä×ë7Ó[5éÓÛªÙƒV¦a!¢‰ÚÂe{½P¹µe‚~™µåQrñì‘ŸÎ“5ýËê ÇE™Çs¯«d2…—Uy=…®aBËXÿ~–äñ¨õš,d?·E©úå®ÊŸçrMÖòç‰´ûqþñOÙ(ûd(Ü)â<>|nH)»õN²O‹“¯âÀžfgIôgEË±éÇç†š¤k¿ 'û´è‰³±QÌv'gy\^ÚyŠ½Q=ùÜÐÔêÜ/ˆÊ>-¢¾þ·e¯¢6øá,ýó’ÒÉlçv&Á¾|nxêtïLeŸXt*£rV°QÆ¶¢É(E?½v¢àM~nˆ(zõ—®›PŠ1oýú4
ž©_õÊ§ø.À(egäe®J¶²ÉEœ—	Öa~ÏUÅëÿôÆ,‘}€ùisUy %s`ó	t‹¼˜Jüa”¸¸.’ˆDi\ðÑØIu=hÔÀÔ›å¯8ÅNfe™M
öÕ²‹«H|(ªž¦ñûqV”Ééõ ;‹¦Kß°i	¨+ñ­´Qu` —¢ÉÀÈÄ®&>‡0hBš? Ià:²ðZŠ°[ Mx•¢toOº§QZÄ½§UÙ1s29Sd/îN`I³ß†Ö3‘^Æd…%.âqâRãïÄªÔ€bç ô@'á2'” Ì£I‘ J,AæÈ>!\1ã4Ê¨"À<ÿô³“qd;FI¤ñhý&Í¢@ô! ÅÉBŠÓ²ï( xÊßØp*Îª\.G>ÄG`¨m\,‘‚åƒWäªž²i4LÊk ™íà«FÊ6Xî PM‡ä{Î6Sh)úøÿ¬t"UXËkËÈËÓ4'“õÚ»Uk~í·çQYlN§¯èžÅœ|d)’—ÉúM!?à›é)®º¬ˆÕªƒs@òÉÅe²§q4âó+«‹G;“2¿ÞèO²qÌþþïY§cgÞšåEæåbj(»æýlÒ-<£/Öo.‘†GÑ?Mp’ºÝ1á¸_&ÓŒ­¯¯³ÎITÀ´õ(4ÄèãâÌ£CÉ)ëZÝê9dê2¹ìg:`.ÔÔ¼ÍS»\¿„7Xñ"ÃvYçÝIM>t,ÂdúEHÓIVÂ†ršœ­ßèŸö &£ç™Á"âæ3,qÎæŽ®Í=°¨Œ“k3Ï£ë~Rð¿ 8Þ£•cXBØ=§ é,Î¯Ý
b˜] 	XËÏæ^$c _-Ê+'oõbéh¦›PöŸcïŠ-‡ñxŠÛSáa
LLQ"%¾³unIa>Éú¦úˆ¾?õ*Vý ¶wò#ÖÏª aÍz÷DN€AŽ²&“a:ÅE÷¤ŸŒzìË/ÙB2¦,=‡ªvGEÿ<’)õüÞ¨‰Ù®q€Î¸ýgyé"KFWfÅ†”×€;‚ì¯Ì{?'gåù±ÍŽ 9Ë'¤–Ù€¥•KNö} u@ÃX¢ÔQžàXtUÓ‹,í¹…q…aåþJ»µVš³l"Ö ì¬œÕKÆÉu`Om÷è˜ä Èæ0†V÷&éõúM™ÏâÀÊßœ$cï>ðþ1°Kd?•°QˆÉ9JJ%I ^gšã¼µFïR…]7»G`&a=¯ËµlUv˜E,Fùƒ‰.¢$Å)ÔS¹~sdÍ‘‹.âkàSŽp2;îTs)>¿Æm{”1Äl/Ó4/’øò”l ï+l™d³?DNÖ[ûõFµÍ·yh\7Ç¹xwjÞ“IR”œ“ú‡¬ãUÚl4uT3º·"CÓÀd6¶|c¥ƒ÷ˆ{ôn($½ê¡n¢¡7“ø’m£œÒë—ÙËl#¾ÀR˜œu; =ÓéµƒÍFÑ0²¦à Ó˜½yØK?þÑþ±AB­ÚZ[.*Í Âjöôt6rÉéoñ&ždØ— =®ÈáèMIlÇŸÓ<’Ï‰»)bšÞÂD1ItÄ¡B‹nL71àà<Ä'´²½ëŽ¬àp]dåŒ¿˜‡Ð\‡oçy–wøª¾È’ÑSÓ»{[Äù¾xyèê€9ì†hW÷z¤~>¥•8GªO°kšÂm¨T«wÒÆd¬»tCöQú©Ù-T];»=:FØ’Vn‡(6¸s`jÒXj7~H"®Û€Ý.*®'CÖ•pÅËòä,[+	¦E¯©è2JJFÛÙÊ§¸ uG'‹lkïåË­ÃÝ½×ýƒÍ—;ï¶ö^nn ûl¥luw4P?aÇ7-¶8õÅÞè|Ôw—åÍdà¯JãÅyVëR¸*™ð“šQFV¿ 	©;aM6d€_ &h³§]½ÞÍÞ+±¸ëkœ`ï?ãÄèú03Àj<Z€õ®1\VsË†Q9<g]Àv>QFÀ9ÎÒ¸Ï×~÷[ÞO,Êt{9óúÒÇFåR’M>’ì„Æ££S¥µÅýS©p“fEŒŠëX*~qºp-ÃJ•`Ó©ÞÌª”x»¦yNOÉ‹˜Jõª­Ö:ßäfj'Š
'fôðã?'ÉP¼ìü—béŒä=ãâêÞ4ËK¹M™¦Œ-Žn\B¨ëøY%:‡}«)r¶">T‘ –šá—â·`Ã‡ÉEåÃóÃ8/
–L½ZÅ:ZÉíÎ±ó2Ïõ+-³&(ÀÑñ3Å×©¹fž”? 	5­Ê‰#÷iJMIŽ¤Ü–y¯)¥föÔJª-kxßEWÙÒ4!D^tÙäkÄzÃEuç3£ô”µ¹rîbˆ;¯ïÇ/ÙßúJ€ÛÌzñ¿Ö¯öògãl§¸÷qìæÊC¾8–ëê„¸s‰Ï"¥Ü(HÉò2û¦,Ë™bþD.ÀLþU#r»ïç@…ôméb‰ÙT~¹Þ‘øŸè—(ïd Ú)òqø(:{+éå#EŽA4Ejb5â ù&ós~ñ2Î·`Ô]ubeäb³¢íŒ’.[  Ùº¬Â,k)®r`¬ðñ™ODþVt:T³<ˆ\gt½XòÔN?jQÖPµüE=SõUÀÜ|ÔÄ×¼ê!	¢.ž`Ì•än†©¶^ÕÙ”Ú“ „êûòK£ËÑ©Uý±+tº$û$1MéTlŒÒ­Q<påºÉjNÜta¯ÙUžÊ»f­da;­Ö±P6æÙ¸‹âÑA\vÅÒæz±®	ÚõzýöÝ.!4¶-6Tydíê‹dS_ìéÇnM|Pí»Æ¥º_
N…bw¯®UMwO‘¸çªhÒãì¢iÌáØøaÔ¦¾‰‘gEÖkç„`B(ã>Êˆ¶ŸF³TÕ§jb{[ÓS¨Mqq}à²Ø†õÚ—Ú¨îòï¶—	MÏP:OÝš%G«S\] «+I±=›¦ÉP¨Ü¸
@‰®^U–¥AWwýË/q†p@/TjÏ,GÈ»@3Û=å¥KÝoU…ú.9wÕUÜ(HGÍ ™ê7‚™ÎØÿÈâ+ø³IÄÇÃº $§Pí8ö_,ëF|6[-tÅtKÆS1]Ô–É¬w‘Àžo€öfçõÞ›[›{–Öï÷O=;ewalî-øhHQz‘i Áf›se_º D¸ –$¾P¡Rí_¸àËÃ0øh¨ÁlAŸ°°Aš¯‘ÄP\	‚×êJø%	|*¤	Ó¼'QÐy¨‘ÕÁ—¾Ô0xB °[“(MâÔ¤:a)-†Ùà$.<FÒê¤‡>Ÿ¥¶¡4'’âç%Zc/X žµÈ0·<¢Ž'›ó‹÷ÛqÿÁº†;ÉÙ7áêo•L{Ïk}±ñ^ÍƒàïÐé"æsË0HG,;uOKìƒ>±èG¸îGøý‚Ç£ž^YôôMÍþû6cËãqv‘ÀÀúïIeõG 5˜ÐM(Hë6C¤ò6ˆkevv–Ê£B@n2"Ú¸óxø!†„¹åÚH#?÷(z{C:éas2:ÖK$†ePW^%çˆã»‚‹f–êF·™¦8À”%ÙÜ
éJãÐôt»½–ßL„	!ácPâ%ÚÞE½é´P:Î¦hÔûy[ü,‘ Ÿh®fÇ’œ3ü¥É`A/´PÄµ!§A(	Â&D9#ƒÃÎä:«ãªÙµ†ØHlµ^z ’âïóŠZô;DÚÓÑ+Ü‘q©è‰xþƒ\kþkŽã5ç,N,Ýó¡æ-N<3ðyÉŸr“.¸
)[¨Ô}ÅX)K…¸ŒZ¯S£Ã»åqv³šó Bª•c_q!'êÚ™ýýeHNè ;_Ç»Â;h±-N	ÛŸŒÇ1Rßg$G=HI7”JB¤ß*ìŒ9f;WÃ8íŽ¸òªÃgü>æå£ÐîX¡»ã‚×Ú¯_½ÜÅ‹;) wR>ó5HP\Û/ã{±Ñ?Z96\È&æD¤ó†_€\)ºÍÏÙYÀ1X†öj-Ck?D|£È³Ky¬{Ø‡øZ“X ÛaÓ(ñÈ…÷N~RÏKa=ÖÙ?a[à;ò-˜ÍµJ’p@³@¨ª”•(&#×R©û÷èƒ­áìÔm§ÖX	¸ê6AèƒFÔû1ßkÑLº€;¤T²XyÂÆJZ!£K¦êé#g
&ˆü°uÄFÖ½ˆÒ M7^€ºG%ðM_RT
EúPãXBk=uêæùþ©#t&¸-4µ*]Æ¨é èÍ¼oB=U[³¢šÑÚ›¦6¦Yn`úñ_ŠªÚ=}ÐSþ d­fŽ¿ùgï¡¦L¨³ÑæÃÔ½³¹†žÈ<gýI”’2û&±¦(Úü‹Øjª@ë^Rð¾Öd?ONr¼I 'Z&ü“H75ýZ%ÕÔ6JÎ’Òêø¶L¹ˆ=Frgs¢Ö¡<(:LU­9ÅÆjoµàúŠ¤º…­5FjJ“3Ü:Y|
•D£¨Cup¼|-6ä&+‹tè/ö—ºæUs=n“ë&!Â¢sTQ¼V¢m*Pª?<òÍ²»‚Óùv:ÕÙ×ük‘&Ã¸»Ú«Ã¼Î(¯sNUp¦h-o3§øÕ—´7{>Ï¹(¬ÅeÆ,‚Šv:=‰ftó4º²`5ŠÝÅªTVSRóZ°:_+‰*¬`.éWÁ›n[…~ƒF¯lÕ%øìpmÒœ­Ã3ÕAD+üÜÜf°™Û‘Õè–*!R ÊÊ­eKÓÄvÒ|qÁËá7êçÉª¿ölæœ'@itŸAh)…[»¨ßu3£Ä•ð¨‰YùDX?¹ƒNËþ[
Ù[Ð%Hnì©²5½oó°}9­Z0!Q¿{.ö–j (žÆù•H‚¨ªÛó3Œë?ëm‰ æÄ¡ö˜OQ\þ[´!¤H<zcŽ:šWŽk<Hœ™ù²¿1­Í¹Ž«õ[—RTÑò~å©õ¥øÀÞç|zî	´…¶¦§(fªcÓ^U>Ø„œ|:#C¹žQ¾{ö’-?Íã' …‚pÈO­]r2iÎ‡åé£.XDt€ü.ƒÉ¸ÄÀ»cv«4‡—Í}³Î6íò5½Dð’^êÎ’µÕäw6d¿]¡ÿî“¼ažÀÚ-´A¾U•3 °Fl
ëN·¿þÚÒØqñHdwr>ò)='Éöþ‹Õ=oIÝ(,ËP<‡’+ë[aX\&Ü²äè/eÙˆ]óí»ÿžâß¾7kùŽzYÙ-¨Û©ÑÊV(Õ6G£­YQfcio}Û•Æ2­Ï­Í$=µüî@ý·¿Þ<<ØÜß÷jçà`óûÛª/Œäí¯w¹Ö1ÒÓinçvÚ=Ð&ßµERq­|º÷¡ü«x‚30vŽáù}dÆc¿&?]¤êrMã!ŠmÍ¬¯·xÇªÓ¤ÇèEÈêù}EÙ,"rD«â¡Î¼¥â˜“‰©ZÍ¬®¯)ÊííIÐ´Ý»´àÝµ«º}À?$Neê¤üž“Tv UGåøtþNÛÀŒ '…ƒZ}&Ïª ’üœ<G£m“{r¼ž·Æ0sE‘Œ7l^ÉÐ/‹Ó•®&Füµ€:sÂ­Êt^~—†‰ÛW‹°ú`á ýý,¢1ÃB²Æ±	®‘Ó,‘E}ö/qåÑ¤Lr^ì":‹`<yÄ2(?þËURBSg/ÑmÃ"cNó'£D\VgaEr¢F¬Ÿ.ØóìÇxAv'žÜjb³Œ'"«hiíA[Ñ2Q­úÀ ìÃ
ÀKôŒ_Š Éié-èÊ5@81NSà¦sõ0T¶•”ö.\®½•ù$“¢*ç)`“µ·é9‹ÍÖÊkÊë¡òjÆ#c"PFéf:›pKGûhS,î§:¯ÄÛ@NJWŒé¢1õ¥ž_„Ý¢[÷$ÊÞêi·–LÃ¼ã–u«ëã¤˜FÃxézéWÚoIÐ
þƒ^%Øx4à¿óìR»F9‰ËË8žØž&ÆÑÕÒ%:¨_-Ex,ƒ®'¾%ÞQÖÎ¿¡ÍONiØ“Óß®¬0îCtw’]æÑÔò´Â•Æ/?þé÷3@ÓÆòù7¤ÅÊ‘auz8ÄUÆc«£ÇCKp?ælíƒ:²:[j•Géhé‰tt¤ÞÑù‡ñBœ~øŽ=`® ›ö¢jA_+Ä7l·¶û6½rëøàØ’È
Œ×oVuë:óZƒ†'ÏvåŠÒËvm™§SàÜ~Ô´Í_TåŒâÑ· ßÜ ÆuMœG«+Ó«ã0H‘Š–)wå ŠíKÒS3ÊAÃïgQ«©øÖ™Š |}¯igÊE–è´ö1sw4ãU´Á1¡6Î—¦™Pý;£|;åžš0MŠE>‚U¸>|ðdßup‡ñ´\ïô¯Òâj‘á7‹ñÉF-\Ÿ6¾çÉh;®œœiòœÉµXÂÈ¦ÛWWV\—KwŸXâ}íîÔc;»œ´™W1ÂÐ¼úˆmyý±½}¹$þ8†ÿàæUÀè`ÿ2¯ßú›’S^¹‘cÓ¥ï4 ¿1n— 
!—u8ö¾él.ì9<Åo{¿q:ã¯(ÇÙálÊÙGØ	aÆ†@àXºL°sãg+cì94ÇÓ·ã†)Ø‹oôNFÊ£èÇ¯<h7„9¹­nÉ}½ïLÐíï“ÍÝ¹ç™ÍË=ÈL¨^ü?š5ü„3‘åHC?íDÈ6ç.ƒ=ìdÈžæ‚òðŸp:ò¬¸Çd€8ë<¹ˆî§)´ 7½:èD'E–Î “({•KÜ ï0"ÄÝ³F0ëÖÿ C`Â™¨ËI´ÃÞü¼øÁØ÷ÑÔn-à¶®-jð.øˆa“L©~¶¾[¹½S/®
Ò2n(ã2à)÷E”‚Ô~Óaž#?·‹Îg]îv­N/S)µÇ¾b€ƒhºñ"¹ŠGÝUßŸ¬7JÆþJ(R"<1Žê!P·Šê™WtýIø{¹Ì)f~ëR®ì)—4ÇôÔ‚
•ˆá­ÑÝ
¿ ©øMñvR¸í™nn<róuéOI0¬»s5 ]cWÛÜs9ïŠ 3Ò©²¾©Î€]<it¦l]|³®ç‘³hÇ!²çò¸FpÝ·qr,<ã13ä'/Ú©1õ{\#c´`ƒ]¾·f.³é»
­Œ	š0ß'hß×Ö<Y!ßÔ¾	xEo}m&v÷P‡°÷Ÿ…À÷Š›®ô1†pÜ^¦Ü?‡—®={ž>|~n3ÚbO~Àu‰V|s?œž>ôâãv…b&¡î<þì&­%ü·ÿŽ'nÄšê±)JÅ=‰c}”‰ŸÝ}ÚU¥-'ï>¿ÕÑïî7±UÑë~v3ZFÃeepÏûA²>8ççO¸sÆÙœ3Æf8¾fu?®f8¢ÃüKÛXÛ,jj)|÷u]¨í~YhíóÇÆôºQd—î&òS¦‡”åm‡Ÿ¯0ŒoQÛ¢1®…?6rÞl…¬ ªhÞÂh6 ßTÄýÀ³¦Æ@žê¬&@Å¦¶†úø[Vâ®2-p”ê„¢˜C¤uDóU«WÝP÷¤â5CQûømiy]{ŸmšSOåPgáº,b‡yt¥ç0]e±Ñ³@í¨¨êÍ6|;j­±¿]'ouT«®Al-WËlº´ºü˜-qLåC½æ	¶¿³ÈÐïÜÅÔV©gQÏç³bÀEÓdT ö×pš.ý-ð@Š6QÓ‰*âD×«T+brÕ¯Ë¥oU)®¯ŒÍ«íˆÕ?@Ú›^ÍÒ29ñ£‚½)Öoˆ#>·å†{‡¬ÆdõŽd¹…u³ÙÚŸÍßà¿ç®›||€\¾DÛ‚ú|Þé{h¨vÌ¢o>b¸cRœ(õXÙ´·{înë±Â[}vð²Î3€fÃ(”#‘ÊXesqÝsñÜ!Ž»ª.·]Ákÿ,æ8m3ëÄYñCÍ¹pjÙ0ã"“˜ïa›ùêù†ç{ø—9ß¶³Ö6Sn»™~¨Y×ÎG&^çs_´™ûBÏ}žû@|Ð¿ˆ¹'î[Í<Éÿ ó.*¬õÃƒÞ30è½B‚>h0Ð?_ÐZÏg$£…)ïDò–(adZŽÝÈõ»ÒEyG#<eÞÂù‰’ƒ­ ÅU&(¡`¢å¹³t¿Ã¥û-°Ô«m«©ÿem¹<Ÿ³†`¬äÆ¸ÜÜâIvŽÍ-$®‡™ëˆ“º]°poèš'˜;d›‚O»(Ì®OHCi”ëÇP=Áˆ×óÌ¬KÜæ-ƒãž·Œ¢zsaJ8°¨Ä‚O…ËSnÖôå—A}t•ÎJ=Žu´ñXžSæé,´í—Y‚Bc…;[Ž[Y“Ãúp¦(R¿´áqUF¡åñKýÄèéøIýYžÞœÕaWXÌãBW_*¨„.bÑ*êC3r#A1éYUäÚùfeíyVZ×%ØÎ˜áÐ¢‡˜›Ð’´ÜÝº›ÅZy’®)Ü`»z¿tÍäêpÒ7wš±píš ô®ªXÐÒ‰ÌCªü¶6/vî<YY~âÞhBx­…©
µ2ª¿5›gÔ t8Ø×l5|DUzÛv]*©ÙÛmˆdK¬Ù©½éV­ºÚJ_d­6¬à–õp «»RXQ†‰‹!V±vjfúHä[©1ú7‰j:a™Ñjž®¹Ác©–}¯âöNÐŠ×ëqvR±cÔ^Þts¦¹TÝ| U iKÅŽiµBbÆÜs ZuMŽ®à€ö+ý'¡m,`šÞÐËŸÊóð½A1åÓ´:ÏÀ!;ú“@# òoü4/Ôäì«xU[–!Q!Kæ»t ŠW˜ÙˆGªa$Fr© íÖS9¾kzÅ«8
›Šš9Î›á¤™;
)W)oé‹ß–úJê»ø±zØ‚
&qe_k„ªåSÞ§šËšsµGIÍMš[º”CŸ«A¢zjnÏº|îÜŒ›«Õ€Îª¹u}ÅÉ’öæj—ª·êM’KnÎ´šf+Àì¨ÉÚ4/é­n^¾×5_A#+¹»*@+*˜WZQÍ|
ÑŠJBŠÑJû™+÷š¹:ñ	ª¦‡¡óõmÎ]¯†-­ÚæQKTi½`)×Ý{»`ªÝ¨Ê~ÛVíáÛ?¹—q=©v\´y•­¼DX%8¥Âr×qTuš‘9g1 [3'„NØþZØ[­¿n­†*>\;Z¨@§3‡à¬áŽý-ÏXŸ§€àUÚ áöãMu4è9g™ß–Ù›•gÚ¸µžeI>á›H;ÍªŒ zö[{i•í”ú¨Pwˆ){× ’Â§X‡yTœ?6¦ZUp­W˜púªÁaôMðCÄ–Êmk~­"Ù7`Æg˜¥€ë7OÂs ìö©’+õ£Ëÿd‚rp¯ãÉùllÂ&ÅŒ¯ã²g?ÄA¿œ‰¾/ïc÷„œë^í“D~ ØÎ¶S‘KEmå6¢Ë&î(J¯'öê•t+Œ%švkMö4ŽFBÊ±"]op)<Û™¹!Š—›«žBÙõ­ ¿Pˆ SX:t,Öoª\&¶ð”Ø£pp ÀŒ‹p´«ƒ®'\é01C ‚P©&émžÚåL$)†í,²Î»“4š|èX»?ñ'k~ê“–õýÓÀdô<“#XD/me4,qöæŽ®Í]ïƒqr‰p¹IÁÿ
€nàxVŽÙ ¸Hgq¦íVhP.ñ4kyþåâÆ^¯[2ÞáTå²élÁV”Œ`¿%Ûì§ÀÝ»âÍa<ž"i-*BF•ê»ŠLf™æÓ­ßaÒè»vIT¬¦ü%€nïäÇŠ@ñTž\ð°†–/Q¿=5	ûÑ5÷µîõ@œl¥ï*ÀÿÛð@ˆjtGP€ý•y—›è±ÍßK§ˆ®Pa.'§•ËKö}Àd‰>?Dy‚ûHÑUM/²Ô»È,îª§Uuk­*fÙDà¼ 	ì¬œÕËÃÉULÔ"{§Y>Œ¡Õ½Iz½~ƒ‚R`•oN’1ŒWF>¢;©„B=î‘RB¹ü×gâ­5z—T¹<¦¥pí®{ëÖª–»]¿‘?háè"JRœL=©ë7GÖl¹ˆó!¾°Žðúêi\¸Ë·“ž02Æ] x¹Ðñv_rÓñFÊQFÈfpc*ÜÚ¯7ªq¾ÏCëº=‹Å©ys4N&\·_¼
Ãc“Jª‡÷¶"À‰Ó¼ÌÆ¶P.ˆÒ†Á…û‚{Þ…,UÝ!¼¿ÕÐŒu Ùbjr¾Êhi¹ôüM§×þE4!–Ñ)8À4ŠÂºCv€l°pÙp Áõ¥¹TX¸·Î&Ü¿/:(úáîÎ¢)Ö.1NóŒG9„Ÿîö‡iz‹Å$É/„a‚j2$Í6‡ß‘ŽŽÝ˜lx€=á,Ä§½²Ö›˜ÁåÀ¡˜Ñ™o‡ý½v¾‹ý"Kxà?Ùå{[Äù¾xyèÿ€9Ü†hWiÀž«ŸOéè Geœš>ÁVI‚ÖhÃŠF¦zgÁe %´±îÒÙRDé§fQ	H5tíìáˆ~·I+°!óI‰2ê­`#uPpu)fÑ¾#CCw;‚ó•e†ÓSaOË‹l©·šÖÅŒEïæE][ÆVÑ±Ý¯kOÒöhJMIX@ ¨)ùŠ¦Ô”\&¥¾Wo5%’BïL‹b¿ÒïV)³3Ë‚ñ(A7ñ*Ts\î+†»Z[°f&³4}ÖÅiUòä"ñ¨_ßö=£ÈSpÌN£.ñf•íŒ¢âü$‹ò_ô)lFQç™X;äÛâ#É<ÊªE ¨`w$ú·M¬âRvð€FÅ³±†{ž]êâ²7bÄ^rjt®)³%§›&µ‡úõ‹,ÿ­$jøéÇcÁ¦t;8)Óx˜œ&CÊ¿ŠÇY—*Ø%cìW¹ÕˆHÐõ'ÒðH+/‹ß&åy?½3‘þúE–—Ýn´ÈNÆã(P¥“önÅU\Z}j?	?©.NJËÑV?å¬ÀV6žFy,k5îêEÌ•Evä‚é¸ç©0Ö©Á€®2À1€^±ƒ«HUA"Þ;aÊ¿¨0!¥1p¨Ý*2¨ßoîoè6bÞª¢š“pÉc8­Ùº61Ð¡q7î¹×ç}›aMøHtŒô¢ÿžÔY/Ý/ˆºbû`êÔxƒŸš ô‹ÊXp ¬L–ÆÑÄÁe×Gc¨x#ë÷ûOFj¼N¬¥@.°šŠ‚{6sV7ºÍ”ÇË~€A)]9P?¡-œÙYŸt¤q`zÎeŸíÅëT™,ùBò;ËîÆ‹V{—°G›/wÞmí½>ÜÜ:´£qÈVwlò'šxêÏB¹ ¾Ø*Ã†ªï¾~CzºPø«ŽÂŠ”ÔúÂ•Âü0<™p`#*Ó
Ü;v’w^¦Ö˜:1Ð@ 1oì£PL#½¶óê³»Êp!Œ™žzlÒ>F,¦@tJLž
•h¦K¥pGkêE»,²‚‰KE"K(Ü©®KL´z5®¤TŠåF%:Q2U²çaLwR…U	(ÛoˆÉâý–îœ°-ð«ìú
fùhMÑBÉÔø$»²ZçÍ÷Œx@[þ>J¦i6Æ.m‹_2=žàjíì¼Þy%SÎaŽe‡Ø¿_Ãï¾CM²I‘LðË>¾©¶ø¹ŸDŒ\•ßƒ-ná_1422´ô©b½H„Èœ
‰ú¢§Qù"™¨ˆ/Îgö>È²vì*Aõj<ëËìœÉ@µ[hí7wìû8wÃJ;™÷~#áÐôÃiÃôDÈ±,™®Ó¼CnÌS …|Ã5¼Šûàà>‘Õƒ–Þ…~…ÑóËŠŒ!­Ô¾²íEÓô"íÜ¢Ûú­bø ²Ççóù5—zùÎ]=±gy6›r•ÀW	T>à:‰s Wëš>‘U=`+Š>~ü§á„ã¨Nêìü—béŒýQ*”9.dùN4<W;¤ mŠ£|«q¢Ni
c&Fpòn÷:@^™_¯+žZs	ò½“a¿ëÇrÛ%zbß>šD°UAwf“òØŽ¾Í¿èîÂwõ‚&¡2Ô£TÔKF#/Èa>Ubˆ„Z¯×0×ÒXï¦zï7d:•-IÚ<›,ŒnŒH"ðøŽæl;Ý	ú“’Âöo¸	7¿%kp®¹A3ÛŸ¼Z˜\…ºbÍÀÿ¢
ÿ‹ßøOøízåÙ~\^ÜÁ–©S ‘V³Ôøös`ÑÄDÚ…:nðÄUBr:HÅÎôN–“!ÔqåLchÃ1ò”q!ªœt"ç¦Kñc„Ë8ßBË'c×nŽGbÖÎª9P«E™-hÍ¬¼n˜õÇpR#‹ÕXêZQ•¢ûëHæs¸;œñ7d¨qmjÀ?Ø#$=6¥Ú Aêxi»–"ØnY~Ú¨jÜ*ê5oµŸÈN}r©FŽ1½å¸_©,¼-J‰,mvU{=¿ª‡KÆ•iòg0Ø­‹¬Ÿ3ŽSÖÐøªê>x$¶®CµûCu'çÚfÚôƒlí[%…	Nø¬?@Œš”Yº¹µ×ÁyT¢+w`*Š®’z2:9¤rb»ÐåÕ
a°‡Èqx]‰:ô6Ý5oâD2îÀN"{0`âW¸B®ÛqÁ")%¹Ä@h›ŸªÈU ?'>úb[ÕŒø_`j’ú›!Ðý÷AÅW)aY¬)™SQÉœ&Ñ£.*Ó$ûxJµh^­X¨Ó@P1ú(ª/„Ùÿá:¹ÓP+¤j„CæçJ„œŒ¨VcDb,NøÖ+ŒíÀÀ2Hi|)E=îRU»¨Þ3¦/]UBá®Ê1ßÅ(TV¡M›MGQ³&yÚ7­t¿ú¦m•“¥É¢ÑÝ‚Ìªâ¹i¹9º¬ŽÄZ-gÈ)æ%È<ówŠÈ§ÿ^M³>D®RE¶žÄÎi.;ùˆ*åüP`4u6Ã‘š "ç-Gßìœ¦tœC`?cc´,OH:½ªC‘;sèÑë8>êè¢OÎîu2ÑþDâŒÞ:&¨¹ç«ØXwxÂ¯ÜÓÂ‹ÝÒX%Ea´AŽ€dv7ÙK-Qˆ­j£S¹…`õNZâˆüü§âñD¢é4½~#ù°ŸØ“.Ú¸ðDoÕ†Wå¢ uòF¢Z1·² Ÿ‘}`1Cß´Ã¹
•ôÏâ@H«yPÿ{-È®ùP\¼í/¿T×rl5oÏÙ7ÆÅW=Œjq	ž¿[íX{	¯ƒ¢³:“u'Iäìó¸ã^ñÊBï÷7ßl>ÿø_Èðììþ·ÿüŸØæY–Gì"~ügöñŸX„ÆylÃ³“²ü—ÿŒêüGÿ#?o8Ë¶À´åg³ôÖÏ’2‚µ.C¹‹¨VÔžšL0,×0¦<&ë(N|2I†I´ ü¿ÿzýèÑÿüoÿõÿb›“(æ#4´ˆód‰}\$VUf#Œƒ†9Ìò	Ì,óTÝîsƒ1SÇØûÍ!ïNÏ¯%ù³ÁxVÉ}£“(¹Ê@žxi%)’r |³QÄM¤ÊÿÀ¡Tôñ›;Éœœ#$€JÆq‚ˆŸÍˆ˜mÐ^.îxÈÚŒŸ<'…Î]ùø§³¤ç§Ø×­ýìköàÿHåŠPB‚e~ªûø¯‡æ,â³ÙdÌ˜LTÆG6…™G3Æ)$I}•à04<z´Cv~Ùs]8/Ëi1X^.’ˆŒ_áóòËì,›¼×;–%p1ý„kf®•"3¿GDöõ‘²è9ð#þ:H‘øøÏc€âÌ>,°}˜Ù“ÿ4)Þ«XÑ¶ÍÑˆáÒP•Iá[¦qKE©E‡|ÐTœ¦×8&Ø:ã+È&ŒN¸<¬wŸNÇ:#˜PöUqV	DÅX6Ù«Pëy~a¡nÆL é‘rÝ;Á²+}]¨òàÁ‚ÌØI<Þ@ë`Œ¬‡¦1GâhZJÈX/md‘°åñiŒtŒÞïÉŠ'Ît¡Ž/,5}¶‹R:ZEU˜‘àb›`£§YÂ î‚ãCZ€Ì¶‡‡+ÿyø²?›áŠ*,¾ü‹
™³dÒEv¥ ê{Šk³ñ<ÂÅÊC+’RžX0¤¾çàyœøœÁ?0ˆ\Pàè`Æ o%‡‘GXü‘`û¡» ÂqºÕ·qg.¸â¹ çÌuàäîšþ$³ü»Éï&¯ÿí_‰2b`&Ý‰åwj¹ ¾:Dœ–)xYÝàæÖÎîáÛÞchkðfópoQ +Aš}¨•ž”&	‡Ò4ÂäHý”ƒ±g0¬nú¿¿yp°7€Š6ß°=¶¿÷æpó%obóõöî6´aèçP_Îé‹-‰RÓhˆ¦W“óx	flIo<¼¥Çÿöß‘“Þû_ßî°ƒ·|“À_¯½	,öæÁüþ»M¶ÉV!ã;ÿÇ¢Î½s ÿní²W»W”èó*¿Á*_îò<¯4<ØÎöÛ­Í-`×aüËÁÎ÷oaø“ƒF¹Ã³ï¾~»ùÆž¥W`Q£º%¡Äxä6h*°s8@8Ð¸&ä¶ŸF³´ìÚÌìÅ>¡1¦!ßØ†õÚ—ìkwùwÛËg°¾:ü À&¬X“
6H«Sv	J¬®$…¶¢”Ê®£RäØ=NF\KOí_70‘(nº]=<Xo\Î‰¶JíQõ*d_ ÙíñðòÚ–®D}ïõô@xa=_ZÕÍšgìG Û°cFÒTEš¢uÏç¤’ m—u“}b:#å3mgoÛŽã^êÄÀ×ê Úlr{æBY˜axŠ[O½ÊH˜ZSÏ“[èh}¡l sXKRÔ[oJd_¡2æDz~Œ^Â6#"r¶e:ÖÞ~ÖnZ«ÂUäôÝo¬"Ú”'pâÕ±îº´1¸j„’‡Á@‚Áê"o¸¬;ãhÐî>iL£èc+SRþþûhº—"J’_Ûwc›æÃÊ|yŒP·#9Œõuy;Gö¥áöb1Q=8q—È±te­m–ÂÖæË×Û›ov÷ÞmnííØ}qî@dÀgÅ‹8áè!ûÒG—´Ûtb©Ê¨çA6l8ÿÄy6”æÒ¼ÐYn7“¯ãËŸl&Rµ˜HÑ"!ý&’t¥y¡ÉO1öç~¨séÄÜ+‘ámËE>ížõ"ï$µ`Ä'lÅHìYkLy…¶#>AÃE|*ŒM3FÞG×ˆ‘ã‚eÈ¨RŒ1£˜,©—œ_é]`Ü²< æ²UjÖÙÿé4†õá[hSm.· ¦Ó½ˆàÝ~±°½Ï>òÐ+9ÎÍ|<Œº#³†½’¥Ôžnï©“¸¦Æ|£Ù´"­0KÙ³ ‰<Z¼T]O‘ÕFH,úzsE+¼0
<6I•‰Ý^(
î»]d±Wl&kt‚_ð¦}°º“ÖÕIACÖ·}4'˜$î«;âx¯i¨¡»J©óHH-€` õK 9†ö9Åï<u‹ÂÔè¢ÔË¬®DþA…g·‘ù¤&9ˆ¡²wP£B˜>$uu÷z–i¬…ZOÜÚ•aeeˆU¤²RK5’*yZñ°L£ñ‚TÉ`šñõÌèLQÝS73tºgzo2ëÞû‡<ÐXïî8ê®‘G’„¿žˆ‰«w‚ÕÔ>{ÄIœcu_%§X„–ùÃt…H3úÊÏŽöT°/ntßéŸq” Ñ-“3x™fE‘\Äé˜+½äåA~ÈíµO+w®p×qH¤"%ŠÛ'z_|ÍwÅ3ëÆ¡ÞÏ¼ë „X«AÕö¨­´´Õ/*‹è¾Øá-yÛ¥-l‰íRÚf©ÔmüfÜ%[å^Ñ=õŒ¾µóu7à jŸåÝ²v^»Ä¾Ú\µÕ–Ó²Ù“ÏÂ>‰HN.Kéõñ	M[ÑD†¤Ä­Ýµm"Ì5rÊƒKR?î¤§˜ùþ‹þåõ…]ËŒHr}~.*Mp-ÓÖÛ÷†!0zWË"£\ jf;WÃ8íŽø– ÿ;„M<N†YÕjÚ«Õ¤Õ[Â0Wp­ýúðÕË]ô©¿“rú3oÍ¡¨‹ûŠòŒïmåØhé0Ñ¢ˆtÞð‹<‹NsßJÕ†#°¶x[ZÑÇà?ðûÝ<»”×½€S@,uD|t\á®Jüò9v[j	dw†ï¸7c6×'”„ú¹ƒ:¡JY¹uÌ¡‘æïÑÛº‘£/Ôm§ÖøvpÕmÂ>	©wÏ
ZŽÓ<´:êaEæÕ¥‰¸¢5LjÐ'œq^ç„WZôÝXs¹TßD(ÕI7ÖÄ„ò*ùÊÇìˆîcóRû_úòÃ;<ýêô|-‹í¹Zò5N…¶/ƒ&A“JfÉN†x–ŠÓÏ“}hÌr…ò µô2ÙpSÀêiòJ¦­Â½’åœØÕáf‰ Æ,
Ì=	MA¹¬ƒ_ûbÈ`jðµÔvmFP¡Š;hµ}c3ªd)ðòSSËZöŽEnýØÌØ ±µ‡á"½ ux¦úzÌMÀ`UŠWàlŸú][¡%·1a6øÚ#¯5ëHûmö™H´êÁ„Dýî¹ëŸv©5j{èµÃ»ì·îq…<~ã¡\YÝ÷9æAkOÚÌ€ä9Ò[??þ	ÏÑ¬§ÑyðêúªáZ›Å>áå¨ÿªÊW60®ü8ìô©„Þ˜ìïÕgmÏ,_wZIý9Ë£v}å©õ¥øL§î¹‰N =(dkèvkç³5H]Ú‡²„s ¾ÆUê)¸ºÄ‚UXºú$T›U«sL#äû”…ÏÕw„8
èWÄ:E”2M=³NHµBu+Õ(ƒÀv=4 P½vßùŽÝB®ßá‡®Ì*õŠ]ÎÛeºEÖ qùë¯->ØUüà#‘ÛÉùÈÿå:s)PµD•!Ò`p”èF®q3ð½N¸eÉÙ$ZÈO5=ý÷Ü0À’ïf†,»’[Ã#d"@JRÔ%¶_/}§CÜ,Å£ö:ºHÎ„ÙWË
lÍ!´x|Atõ~yélJ|¼?¾JYq¿Dçñ¡°Èvãré4¡‘b}Ÿé~œEí«K¼^Yn,k"ÝTâ1Sž»÷ŸxÁšÑû½ñ‡o»J·i¹öã%dÓ?gÝo˜ø2ß©è#†fé™úÅ3=ö#@XÇ	‘ë´….!Ë‘÷ó(ß:òòåÀýWŽwá[Åuˆï{uŸs…#³ÏsEß>ßÉ{™eý¼aÎ¡sÿ©^8wölÏí7•øl;Öo&4¡\gy2bøF4,€vãy}ÌÒ3òúm0Zî^ŸÙ‚9ð•!$¹ß®˜x‘LÐb1É½é$¡ÔÑ2º¸pýü'À#¬ß +M/¼t.Ë×;'U¡Ú—çê°Ö¡H£a[mê­pÄî0?»ÜJòa?®î·ŠµrÏ®ã`—ßÙPºÒ¢WÛyâ:¤¢ûi6ôFfú¡!ô}{}]Eè£èB€½¢Ã¸ñó>ÛNnÿê}¸ç‡¹ÛÛiu÷M@ÀšþÛË‘§4-"kÕ<æ«æ;wÕ8uü;k¯Ú¶õÆîÅ(Z;ÿÆmŒqi‚qHuO*cc‡B ­re6!zÁHÊÖP’mdó’“Y"Ìž­»Ìdá¹£Z>ÿÆiÍû&ÐéËŠPÇA‡eÍ<ÊpÑGW·¡fª–™¶gE™œ^/ÄåeOìh~E<NjC†ƒ[ê |Uq§ûOjâñ:+¿Ör îÓ¹Äà6ð¿4Ó—ôé# «¦p7£¤ðÛ† tv4ˆ!3”i¾Jˆ{UGäùj!—æU-*.Pm-5P+ƒ$!Zü¬Â›;	6t¡MhÜ¤È³Iy{Óaê!Ô5¥£ðNPkYlµ{ƒ»–r!sûWU8‡^Sª:Öª³´/9¦3+6¦»²•§)ëódXZ½È¡òÊá´^sÑªÕ…Z†Û~î¹.`U>Äš¼ÿŠ¬YEyŒËÍ»LFåù€+%,¬|_}¯20Weßà?W £ÏºüÌan‰h¼¬îÃð¨;’ÇV?Ë!ì	žcÚ‚ç˜þxŽÎ3ÕÏO°ÉL?ñ&3ýKßdÔ.B„êÚeZ9ýŒiee‚õÚ»­RÏm•£š©ÓÀ*¥†®àxþØÃp$×U4¶óïoi±*¢ö¸^å=ª[S³EØÛZ}'cÀ^Ÿ>¾í>}ÂvüôñlúéÓdßou¤ÁÖßjÔ·û§Oå úÔÜð»Uq7Àkèž€yœÛ öã\È¨GXÞ6ÿRãN¯€3@U2U$×Dj6Ê[^†¿µ)T{Åp€ïØOg…d7¯ƒ«
Mï–¼v–‡)UôÕ5n‚Ê'ðý}á²êÃÅ×Îs4Í—¦Y ?¼¿o§<ú]…Ö›BdW Um…k	ÚOp«¼ž"-JRÏË)>ÑpOËõNÿ*-®þ	ešÄÍ5×o¨=gh»¡;ŸØ@½LÞ×–ù¼ÍM ƒ·ZüN’ÝóŒÖodLœºu„BagnähDÕ‡Ae’òšŠ0ÿº<‘
U{-sã]óqU}$™5’Ê4\•Þ¦b°CYÛ¹Ž÷NO-ì„.`²•æƒÓã*à¡H2I“IäªÂ½è¼L0{ÁÃ£–s[òsˆü oTçòƒ°Y
tj>RÒùÄ5&ÂlØMGÝxw*$êxÜ‚µ ËÛÙå¤-!#ž2Wñ„5ì§¿ó Þ;FÈ1Â7ö©Âwá³¸5åöœç Ñ8_ç«ý~¿êTHûˆtQÃPÒnlb@«Ü]màÎ+ò–®ž[¢9hÂä19È)r6+q)‚7AÍÂpVÐF*$/í¡µµÿ¢XÚfXé t?SP‰8:a@X®CÛ ÃŠµ÷3ˆG1²)·Õ0ètžÉA­-‹ù©`ñŒ¼´,®"*tž©_-zaž¹)áŠ€lr˜{qwÚÕkÈ¸˜mµŠLöŸ	ÚÌMrö¥Ùp5Ì,/¼­–-ð3…[ý±|·ˆAógFúãØ­´¤þÂg^RËªj}”ûlY]v;kGÛÔqÓ3õ«eAr@ôÌüž‹Öá§vˆÝ7uœÖ/x©’z-g¯¥>ÛÊÆÜˆIøÞkY’__}v7ÍâÌ½á7­ûzÕ¼ÛˆÎ}µ„&}ýk‰yà|€OfŸ–!MyG£ Î¼ÌƒRßKè{¢dÀ£Õ•éÕ1ÑÌ¦¨ÌÇS>è†Q.“QPm$;â`.·ý–].­†Žüd©*-‘x„®ˆ‡^<É®ªÏ"dpÆõ°JY<51MéÍy™ì\F°Ÿ@þgl¥²@õ¹–G'œ¸”†T¨ð“sŸ¨”çóÍ¬PålëN…¹ÓÎe¦™æù«P×Ó×:æ¶ìü¥+N­]Ü<a_ùµë´ÈëA…FÇ<ŽnÇî­Æ(æ!q‹ˆïèy¬TÚ ¥Ë¹ò«²¶º³TåÈ†ê­k…VK=•xÝ Ý“¬ÜBùgýYž>¸«ºÌ£œÿˆÞíüXØa’áÇœöÕÊCÆÖ›¦nÜþM£kÔ£aßhWƒþAªž²ÆGHÕ#}ŒÂNû¡å¦
&g%áç¶× ‰lò
VÄ0Ï³²« ÒP¨>s°ó,J}ìS¹(‘­¢>ö19ª×üIzVez·Å¼ QÊÖo¹^}gÌl.ÇNŠÜmIWQdHÏìÍr³Vžd£k
Wà¶€=Yºfò‡fÓÂº}Ï[N¬cG„iðS#å¦=Ê›RÐÃ÷:ÊòWkžÅJ.ªUì„Õ³WÏxµg½óUIÍ´§©º%âòBµË“òIºúEæ±Kuëµò[¥U.`cÕæßî0!ûüOjí˜¼ó¬ ‘ÈßVà¾y¨÷ºæj¡‚Ri¶¨Ñ®;ÃéégÑn‘QîŸÃ~çÞª½OÿjÄ>g³ÎñÆñ(™)<ÿ¦õìr/þ÷›ßÆË³ÕVNmF)eaÛÈ°¡&=>yhR7Â6c4õé`r5ÌíÃ†±ÿïÿüœ>Ô©Ý¨–NÍÒÀ'‚ 6µqˆ§âÍAí*ztßÙ¯ÒçA>.óh*¯¨#x]:z¼‚À¬ÚM¥·9Ëñ,×4©>œ;ƒš×–þW‡yØ´(ñÑJ‹¼ƒÙüè…¶‡½æZêØ#|š@6ÇE2|:\š”[1×È¥òñouØþV‹xš`Ñ Ömm£¾óí	C°ˆ
Ú×4ß6k Žõ">Ò÷bñ©—Ûx7>ú<}¡<‰‘î”¬ rsIÚîíGŠ7?`“~tKÇ¹ÿØÇ>õêîûš¶GC7-[§ F|Žö¸ß½žW!u6~qC‰˜fÕ÷J•æ{¬Šô¬RJÓÁ==kXºM‘«ž âÖ8=lÓrƒ÷ØÖÕøÜçþ4sj¸öÀ¹‘"ºGš§k,ì9Gn½“p½»î€ÒÂ·ßCÄ•J]xýîîwk*û¸U™ÆÃöšÖµ÷ò×
	swgNúÌaÍpÏNÎc¬`72\ ßçß“ël~ëöë‡=cq”ð×/²ü·2¨b7äW-P~·5dyÝMóÜM‹/4º&ío,Í¾Òœ=þ”XXWëòMvŒ©½KuôDºaÙ›~üÇÿ#æq°ÕAÌýO	¾šd“¼\ÄUí5ëÄš˜Ò‡G-ëÆZKtj¸¼fq~eÔ«!w—þcŽÚßˆ­Jó«oÆÓc«2úFœ¥¼âe›¯È‘áú—å7:ª8©ŽP9XqåŽèÛ•3…EÑp„ªÇºÉ'…ªèUsÑOÔR“ êñn†ucíÐ@]´uŽ­ÊŠØC¢ þn_
ã™r»£6%oÛÐüÆ«‹¤¾¹	þ·‚ÏïÎx—òìÛ¶«5áFåw$ÕXøq;bý_a°"ŒU¾¦íN ƒ–+´µ5‹mžz_h{–;‚û0Šó‚w5Ã>$¯Ò×ÜìÔxXj}ùZ¸~ñÉuófža–Àþ¾~ó¤z~læêcÁIžÇqê È(U5Õà‚´ÛY'²ÿpÀ]åãæþå3nÒ9ÿÝ³ÍIòg¼»A<:Ü$†¾…fÒSÚ'Wñˆû	.—VØpøü¦rT _ñ’§bÑðZ'œÀßQžM‘®å!1Ðì8ÃÕÖ»¡H&°£ýEÈ§¶²ÈŠa”ÂV¼ÒÿÛ'aJ	HXåVu¹Õp¡ø*)çn©¥90ü’f¹âÜ¯ ºvÀ^íAƒgšÐM´0>ñMçõ&¡[	¸î©r*Qg/pCã¢†ClNL‹ò«¯úš¼\t.]Bþwä—ÚÆwuë‡laãÐ²»8ÁZ*Iãÿ¦||[Iøëˆ~µßdƒÕe“ƒÙÉÑÜŽÛ¾ƒfõRKƒ×@MCs0jƒŠ–Pâ«Ñ|# ORÛ—‘‹¢O„o©*gIøàEX¶•§Í´’è/øëÏõÆQyü{`âêíPJè@Õõö5s>Y¨ÅY	œÖ0Ë••µÚ¨%¥Þ
¯º…"µUûö¼®z}j$ÚÚñy N­-Ô]1§V€ªˆî?\Ý l¤ \î­¦º²¿ åÑRÅ{ú<pSi§>W5Ú4Ë’ï<5ß~"<»]4­GÓ†Cîþ¼Xh4¬öÞ,T~fèW£Ô¸Ÿ„ÊŠÚ;D¨¬â®®H…µGlŸÙâÃcJªOÿI6ÅõÔøŸëzùõ³—æÏlMþ<¶~1ôóà[,ÐFSçS¿dÃQHï)Ÿ„˜5š»?Š˜C·ÏI_6;Wöxåñ·ýÊë†?¬Rác?	V©“æŸ­È	÷gW?4!!€?	¦“‚Ï•‹²!~aŸô·Ÿ½`ÐŸ-«”Ï]{š_QûÉ$Ìm¶‚ý)äË÷ë>á’Údýåaãçs¸)ÁüLp2FëîM¹«š´
ÖH{ƒF“Q|ûæ0‹4¿p’¨ƒš:w4M±ƒg²µ6Hwg·`	˜é5AÆµÈi ·x”? œ2¶Å¡óñ žú®µ¸:Ü|—rPÓÂw@ç¹†ð³²õ¹nxáƒ·®xI¼c2Ž&³(í¸ÑÆÃÏ<¦ÈâiMÂÔ£4UÇÚ—ªŽIa?M¦œ<O(4yø31Æ àñfÃ]˜§ÉdÔæ².cÝˆO~ÔOÄý˜£6ãi1œO:s-»­&MAnCý†ë:›9jqñy˜ýNGU¯¿ÕDé<˜Âò0bâP]¨}§Íe |ÜíEÜŠ†e‹Ëð´CüÊ;CwDŠÊ×æ‹íÐÌËmVX—ÿêíN ê5LŠîDËÁÖÝ1¶‡#á+éPóöfhöÙÝNÎ¸‰š(‡×Âã¦Þ·ë{Ó…+‘§É“E×Ý¡è.Ðt_sÁ)Üká£‘=øDê3ÃMµ€-|Úò^žœÅas7ÒÅŒB£t¡áä
Ñ|/S3ø¬¸&œó'Í¥æòŽ%žO´ÓÌ#ã¨§iÉO­ñV§ù~ï'Yöa‘íG“Ó4ÆpwãE—CÏ³»ûÔ^™kA5îê—¡Ûx,Ò´ZæÅÙ¹°µF•Â¯zäiy¿ä'BUv.Ð	û»Ù$™d‹sê‘¥M**>´DÞpáGÎ‚› W£E +õ£bãÇ]/}c'rhÖŠÅ¼ªèP5Æª¡Ô`ÁëDéÞ{ à61«¯]%ªÒœmâ¥žHÜé­§7–E¾±Ô¯Ž)^w/©&8'.ØÀ…es#Ä¿ã%Z¨½¶\yfM/¬3±®y#3TT\w"x:žiU-¤+IU®;MÐÎâµ„¬ˆÃý;èâžÂS7WåMw/˜¦%%IO»Fp0‡{LŽ¦°{)Z€nŽ‹³@O…üŒ~”×Ñm¯‚ÌÛ<íVÓøfX§=¦K·³Ë~†c†ZYçÝIM>tª“ÆÑH8t©lÑµ¥Ç"Ü¶¤¦Œgð…´’¶®±j­³ö<½~£ºàæ‡¼¼%S8¼ ðQ„D5;EH
–Rx
h„tÝþ–Þ<S^ì—Hz‹ª³ÚépK;xpžmºŒñ&	Pöél2äÒß–-ÿÄ—œI.5›f@âðç4Ï0®$þ”÷Í¶u¦-˜ó|×Nq^àx<:ûQ>Œ“<BÙGÜHä) zsšàâùNÿS¢ÖÌâ£ÛSXCˆ²£ã§¦ÛÀ–•/©EVn É.fÃ!¬_Y'Îó,ïp|¹È’ÑS3È{½ß/OíC-0|[³›~u"9y	ëZMqøX‰¾{à‚’;"M%‰ŒzŠÿœF,Z˜§‘ë¸Ácåól„Pø~çC`ïà°ó”áõJHŽ&×ì‹ñ~B§ÆI¯Aê³§jz ×Kø+: SïßGSŽ¸
òtª°cÓsn%®fA;/	VShëV,JA2Lt½E;|Þ1à `ÅÝ>(´¼Ì~§S`1Žã	GÎ0(lyÉP{é`Ø@1”$Ï¤Ûâü2ò9‹…ÿ(h¦;0L€Ÿ“ø’ÿìö¡E¯¡kcÔ‡:^ "üïðÞ•ëUêOaáŸCé˜Šç|…i€_³Õ^Êzåe÷1Pô•ŽUx];Ee?*Jåq9Ë':;v»ôÅo@U·ï1×-š¨?¶£ë‚WÞÌ \Š}ç=ížÍÿº€Åeé~1M“²ÛYRB2Ìó¨cœ ÄßôTG;{´6¸%={Í» ª9Z9Ö7¬ôÕã[b«ÁoUÙ)l§_HXòX˜·N?x©žMq+âyý!I“h"±GAvCã¿LÓ‚ÌWœ{BZsÑDèvYý^%¿WŽ)ÊèDŒó0SËJÜá‹>0H2BžÅú¢Ç?~ÊÀóôó˜‹~Ýåßm/.vè(eÞ~ø_ü6¥„+ G*ÑYf'¢ÝÕ@ä
ÿê
îÊ¡«†]0õ¿òä‹ù*@¤¶e96ùa×“¡ °à£MþSÁ°›§9M@ºÝo žž6á5ýŒ.£¤tÈ²S¿®·g:k0š:;‹óÝ1ú¥ÌyMé&Œ_5é0P]©NŠ:ÅJš’â*²“Œ*ÜN?ÏrÜdíÄìvÎTblØ@N Ä(ËwGÕßâb—ïgâ³Ø½8 ‚~ÕG­£ÌÍvë‘þdrš!*ünò»ÉWBã:øŠ}q#€"øjøÂ)ýÀ€_EàtC‚–œô°È¯³œýÒú$¤ø,Æõâƒ
G ˜Bt\&VÇ`©÷O…ëjŸ!±2c/¶(ðpœ“()XšL>Äü‹¨•‚†3u²1þryIdÝ¨Ê'¿ßwmú;
}7j'¿"SÐU=d @?Éê4|tìÒmÎ­!€ŽŽ{â°³;ãX0ëÏäAf2Ò¼9.HUDbù³/Iž‘8$W€2	{ÿÕæERð˜×¯³‹ˆm–j]6måI4Š¾BÜÚK?þq‘}h"«œ€qûÕÂï&oÇ,²}>vš%lÈ¢» îðrÄæÜŒEÀŠ?þóà‹œî[¬~€r]d9j³ñ4šœÇ,Cçfe”žÇ©ªÿÞ"‚Æè˜5p%Æ!QjUãÔ÷q½•ŽªJ•Ð	@¸¿.Øøi]œ ¹//›ØÒœ”¹¡Ì«hä-/3Éã	Y"•˜XGìšõó,Ž[ßì½Ü9è¿sp¸÷æÝÛ×»Û›Û;8Ë3ÕöÌîã,ŽàáXô‘‘vÚQqÜ2«Q^½j…X¬{0‹Ø{6š©vÚ Ù4N9aðÑ¡
hC]	’÷þ©ÓíJ<²†»ÈV'?rÝÊ}MQ²Ç†’ík‚Ïº Y””<ÞÓ›«ÚKq	ï&ª]m67¦™<‚	¢(ì¹œ2ÕÖç“UxÒeÚ#`åL$­pœ¾‡ïîÍª¹º}b­?…]q*òTÐ³Óä÷x/´äG°×ól’åÑ8Z°ð°ý¡»z¹*lû¦Oh\ŽGåÖÞÁ'Aû¬t*6¶#_áª"Y[{¯vÞlín¾4¿è‰xc1K|"Äˆ¢võNÚ°·ÁÖPæØj •0Vó»‚CSÙ@^í+¤Voâb
™>þñ"¬™E»Þs½i…©õ8@©¨ô<…u= µïAœÍÆ=à#DpÜ6íüáÝ~:]»<”(	+äy4üpÆÏDÀ+çÕ%Ôxäñ8Á°mbˆ´iV”K¸å @§µÓSZEÏ
”8Í5â_Û[$.hœã?¬ïJ@• ¦÷ñêØædôFvÂHC–Z”J£¢|3›ðhE}?€%àB…Ã. ¹ÛÁïÄ¸ÞQíÔ;Þ Qí‹ú&Ùåa2F­
	 c\v-Tµd[e]Rºù­ÇÖØ7+øPìT#¶§ÍétÑ²Ó‹Jc%ûÐ3 /3˜RÈ#F\ÝJ/´ôäÔÃÊÃù¤]\ZZZBIEM[åŠ³“ÊÇª$×dwöO@Ýíèú9/ÀÕj8«B7æKb‹liÕÙ3ÉBñ«²Bßš¡ª5ä]ÐT¿yp>Ë®©ÐýKyˆð2ŸÀ„Æ°ÖÙ¦EètÌæ Åq0
Ëÿ?Yd[{/_îlîî½>èom¾Üy½½ùfwïÝæÖÞÎÁ¢¢Ë\ÓãHCÝp7þöÇ-Ìœüû—¯	!Ûcüƒ{_ÿo6]A×æØM]Î¼­@kw@mojÒq‡“×6a«8ÿøB-/ÈÝÃsæM”a@v0P¤³YÜ”Úî€CãÝ(¢<àfùñŸ´°à‰w\så¯@š¤Ôêj[m:â«šI_lUOP|UO³[Y°¥p«ž ÁòüÂ.…ZƒÐkÃM¿ÕÈ7—àFÅ*P”ce3È(ž{ÈÀê	™ øinŠý~[I•”t¬eÁÂ#R®XßNln•½*$¼÷§Õ«ëHjPœ¦ šW¬¶{2/Y‹Ö?5m½—(­ž»"©"'ËJ°ðãvTµ‚žÞQºn]Ü¦©hÚ‹¡âiE=Û"WkqTå?.b`¥|¬ª!†w1Åã"	ùÍ†xû†uã<wA€Ã‡)és«	:/¢$%ýSrè’*”ˆêÝäÑ±ûÚ¨™_ª¿t%›Ç}†öp3hV
RÀ‰NÑµ´°‹µÒÍ&æÞ™ŒÂ¢Í‹d¼ÈÅ]G“LÎxWæE<<xÂÏD¤±;ô‹Pó(ÔpÈ%li6Õ–ú•=Ó_1J”&Èf°nì6N,,ÔÞm¯_¥Ç‹Œ³§ÿ¦Q±Â6 ³;Õˆ+¬8ÒEå×>Ÿfè3ÆAŸ&ð1Ç…8åÃÿ~"ëfX¨þ½…+»ì/û<>ñ>o¡qûm^!4°’dÁ±è$J®2«aó‡%±‰ÒPs4¬•ŒþÜÌ€»ûå,uÏ²O,âÞ^¹›Ó¿·–J¹LÆœ¨q‰ZÒlVv+µÌ‹\gk›¥	…3Zåª<¯RØ¥-²#nÑ+ÀqØ©<€è¾&„…%I 6–k¼½|a†æÃÎ"b`²ÎÎuô‚¢ð„-¾µüÓH¾F“aœâ5=ëòšxÿdOòîl¹©¶Ý§[‰TI
ÞÒ”úÂ\/ŒDÖMó­Nui@§ìM©)™›£·ðâ7Ä‹•ßÜkP‰»-bA,’k"Åš8ÛNYš)?r›_Rw„¶à³“Ãè„×¼iÞ­z;£¨8?É¢|Äç7…­#êÈzÉ7·ú‹$¾ä.—J^ûúÕ®|Hë…jEÅ]4êEsäb+›MPMo…wcƒu…]¹}bÚMùzIé-S«0²|Ô6iÀVžZí>Ï0B}ËhÃî´+žÏæk×€n_’­Þ,T”LÜ)ïÔ²Ž¦a¤Ežz5<žIÄbà®þEEÉ'?Q‘E*g$ìª`â<}—‘„"6€ü…00f&PæÙÃ<©Ë5E{Æ"R¦Vóa5‚aUÊªOä½ˆÒ,7µT´ÆsíåÃh”Uæ±ÏA(ùµ5?®½[¸V.¬ "_™L3}ˆˆ÷Tÿ†¾þI©&/¶¯Ð^;ŸìžóÕõ"®Ÿ1Ë¹´•§H¸¥+ÝÌâv‚ÇõÉEœTšCHiÀ²Ï,Ô´êHâg¶ÿfïÕðfü ¸êã»7o7åÃ‚©W[2Ë€à:lNÞ,õ.èz7·_í¾~÷jóàpçÔ†²½Ø}½ùzkg÷Íž‘’+²ºæÜÌ/^Î°FRsÉëP§–ÝÛ‚Ð}Òþ¹²1Ýèï* ú-Ù3,Ì¿+
ò‹ÎoøÄV!ˆ§!†~)ÀÂÓä/øñSàÇ£:«riË4C·Mâ¹Al—¤D0°,ÝFi^“*Á7Y7;$ºyåþKsò$7Ý˜ifG)CJÔnÛr‰¶n|*·o|èND
y­¿#y9"à%s×â…­ì¶ÿ$›[hÓ!ûn]ÒÅk:E7È¦.‘¼ówˆ®éŽ½7uÈÊ=—¬â5ò¸……`}N>ZAËk^Å‘$ßKLå6[ØT§•{~YÅk@F¹ÐV]yïØ!Q¸¦;Ã¬ÉÊã’(5nªâu€>ûltS-<›;á!>»¹"žÑ­ÊaÅƒµXyø6ÝÀàî<X•—Ï…P—oSUÜšcV@t!ãÈŽh:ì@Òâ6›{º¸S­ì¹-é";"½°t\*T/ïÞº0UŒ"†ù$|Ãò2£K‚’]€|t’¤Iym.6ìÖ&7î)ðú>Ú®,Â¿i\À°Gãd²|ÆUÑ9üBð!X<Â«‰…¸ë­™uÞ²eý2{™]ÆùVTÄÝ^?AÞe]sUÜÎÑ3Œ<¯ƒsUsWò”t«zR»w‚ÑB”çˆE77lu¨Ç‹õž† `7AUµÂ‹WÆnp–˜üï¦›¡Ú‰fU7áé`I;6mjŽ¹ü²ºD×-"è.9ñf0³âÕÝLÙßiùŽê~­¡Øh»ÎwE’WVŽ×‘eÕò·’>ÄrpÆ¸,-°Á}«{ÔDzX[·Ëëõ¥D†gëÌiS’@áÊ5Ð§Š.Ì]m›q5ŽÇÐksraW$×¼Æ{%¸É¼Y7‘âµûí-	ÛÕ#R¤À"\@òœðU<Î,‰RY«—Ü1ÝR¤¾–Ò‡¡‘Ëï‡[YÕ¶£œàjÑÔ2Ø	¢Ù%SïS·O/¥÷.ê2™$3ïtþÎ†ÚÍB7™ìäpÈ¾f]ys ¢rV¶¢0¬XWƒH»Ï-1âÎ-[âD‹¶%´`Û~C— ¦ô!Þ5Sµ,¾ôØWlueø²ìFwï®¢ƒ¿ÎJÇ"8'×‡×SàÞÄÃ,­©úB«ùÆxs+ºbO–ïDÃs5ÿö~Ñ šTû¥hùˆŸ–tÍ‡z‘k¬çtWí%ìHÖsÃðw@ö¥¨è +¥Z±h°`JÊÕá'gKñ8Î£t„þH;ÂWŒUÓ> ,¿ßfê‘8lÕ!Œ¬:Ô­lInèL-j ©)~FVM3s÷SÊXéN:mHŒY´ >`h	3,û0Â<=I$››gÂ£ðÂBÿXø¦gËü+yçùhÂ4Î1è<wóàb*ž55a©©ë¶GúVdy	h·ÈNxŸNú¢¶%NƒàWÏö@j‹Ô/1®ÂK¿ß/ˆ^Ùí¢?_Ç¡Ûâç­:»¦KhQP8‹¿QÚEÌvs‹/Nô°Œ‡‹ÜW$Ys1ê0m;>fi©D	j8¦|­\£WB\Ñtèæ$ÎR?´a:§¥ã¼òn+0zIá‚9{vAý–]Ð(µšÑyI3•úSÔVrÔ7cåu›ñ5NA©¶hÑ„ÈI¨Ð¼™bTÝ¦ètB-ÐM¯…2Í´5h­Øœ§„5y‚f&N9bŸö,eYàôrº© (œ3i9ÚÂºä˜“ W4žvu}·Î¾¤ØžMS¼õ+áº !ÍØWI‹5Ç¼ŠØ¾Ù‚–0¡«.ôÙ2õ¬­Ò$9T*qXYðÄ\×RIaR‰VºšëvþîãY|…¦vôæÞR1'Ä<¨F—ÎçŒ¥{ë’4mu6îµ’!¶³oõl¬H@Ñ$‘ÞêaKc1‰8|¼Å-³3:/ÐÄò2ÛTÙ@ CŸ%Wª >…ÀD 	ž|ƒ`ç‘Òt_lÿÐcðƒÉãˆ*IQXÎŒµ¨ýJ)P*Ì¨±5Õ'Ž#ì?dVvƒôöG¨‰JÿÜokÃ(LÜQˆê»)~kmÁ.F´IÁ{fn#NfÚŠþ)}³eÃ7ñ)î´ÿÐ‚ÜF>ìo,ŽØšpÐÂ8ÛvYñÑ‘…­Øy$n ¢TËµÃ—yðtv`)ÇÙh¢x”Eø«¬Æü¶Õz>j2Rn.è‚¢Kˆ[Á6y6dQ)†€Æôh¨k¨uìe1­èÍKLºs™A›RbM¿ÊMõ™¿Òïô8Õ'Ù4§ÞÁT6™@óÐcQ•Í¤Ù<·uXÑ¸`«lˆTÁðwÊˆÛºN[ùm»,vùlÅµ§×U'ÕÁóèÀÁsÕó§>I®9Ó­?5t¬6ýªâu+OØêÎºjO±êO§šjN’Z5žüTT×‡O€üj^Í}ËšÝGÈk/’\íïØ7®¬ËQ{ Wr‚„¢vÿ·ovwÌÇZÞÆåá4…ÅæÐP!<¾FÍçÉŒƒ¥È:;Kã}yuq³,‘@O†±8y1åÑ¶¹¥¶ˆqþÖxGÄD,õÅÏá,G—Av{ê+wß®€8}Tz'í&$¾*à:Ë¯®#Ò£c»˜–H»– ìUfæÅªv zP%?H**üÓFüñünùµ"+$›ôcÉ^oUIRâ˜p„¤a¿ ’Í²ë°ksXfOW¾î3<vÝ6{á€Õ½Ù!8ºw~³úÄ»ŒˆK ø­~K_o}ÎRó~•¢ÆÜéÌüÚÌ·RÍö&f“ôhÔêÌR~™ï®è«I,™ºÌX¬/3€s"eQO©32½Mà=ÎÉëBÇéR)ó¨
lò‚Õ4ÔH3 VtF…h {?b£ñdC!
šà¾Ùÿb…‹þN…"»-yºÛz²/Ö›:jX¨ILùòVTÍjBÁÃîÔ¢ýµ+ah[Ó VÞ(èÎ@·=· o ˜H÷
Ýò^‡÷¹ÇZ1I¢¾í†Æúmn/…Ë/:“W	Õ€Zp·4ò*o»zƒ¬O<Ú.™I|©tß»óCÎƒbªuûAvZZDbMP=Aïe	3)üè€}q£«»}¯YožÁômÇx„cæÄ'ÇxW8‘×OÓ KŒX|…àÈQ§¦tlèî[ŸY¿Å$x›í'	f§F2W»DŸª·* Bãé{È£ Ü¹šfÜ%™6óOÒÈÉ:vä8ÚÒIŽBÐf5
:¼ŠÝÉÇ?ÌÔ¸‡úv6&it ¿¿”æÌÚâF}Ø²¥¸ð.¾ûrºZˆs˜ÕD!³y”¤fTÀè0Û¹Æi×@l¯'*ŠðŽ_xëx>ýÅìŽÕè&Ÿ1­ýúðÕË]ù¶“ÆH´žyó„z&Ðß‹þÑÊ1¹;€‰ŽóÁ„7üè´è=Ç¢ÒXÅ:Y¶×ƒq!‚nž]r|]Ä½éC|]?ç^QJü&æA"äa(–ÂzªœÃw¼³ŒÙB—l1Pì„P'T)+î~à=ú8+€ºíT›Åu®ó[ÿªÍ/¿ÄFÔû±½3ë;VZ½O,óêÒ„™~êjPO8½XgfÖœ5€|$é5FòÄ;Õ&ˆ5>1»¼&@jQiŒXCž{ä ’GO«ÛÄÉqÛ·aîNŠ-½L×­.u†fÍB×ôræý‚PmŠ³ccæ …î“3ã8F¦	3ÂŒ‹”ÂÃØ‹x¤GìWkOÇíXáö<4Û³RÑm„ìã´mwåÔÓKÖ=¯6ëv›ú†úRvhòcÔ4]‰SÇšC<BW)1×ž$ÜpÎ1‹BT…œöSõfh¸ËÀÚÑñr²*¯ÞÞ%â5°>Ì
ñ;Áu©­{ðBà¶,¨›>…ßs´+ÕµAòM+æ™êHu¼m†@÷m¾4uyò&¾Î1ÅüRœr`LÖœG)—¢Ÿþˆ*QèyAî¹UàÊSëKñ!™NÝÊM´'
a4¤ƒ¸¤Ô–lµh—Ä7[¾»„5ù¶CŒöâÑ6ŠüèšßTÇÄ+lœ‘{ê.>‡p$QVcéŸGE÷ý7¦FKåV %ñÈ;•0—“°Î!ç:|vYÕK¨qî^2=Ë_m±
î8>rÚœü_žNŒaÄ&ÕÒ-“!KeÛ—td<…UÂë½N¸eÉÙÏ¢¢/<œß¾KÎ'×È¦sÕ—É†pÖ–°O·˜FÃxézé;K÷fù+v°×ÑEr&Î4¿ZV`r‹Ÿ¦ñ Š%´M‹sY·´ŠÑoe¼[xQán_¥¬8‡=ìr©ã¯®¬°Ë¥Ó¤$‘}ýX¾xM%~PKm÷]âÃŠÉjz}3œØ+±b1Àòw<Àrÿ	Ú+C)“à½v¨^‡*gbÓ7=tÂ7ÞÅ–pKÏÔ/žéñŠPDµµèÀÂ*ÉfMŒ¨íÐÚó(™+/¿aEò€Ùê¯ní¸Îk<ý¶ŠØÞÆ;÷D
o'Ÿë$ŠÞULàÿ  ÿÿì}ÙnI’à{E(QÛHV“ÉKT©Ø<ÀÒQ­©¤ÕQ³€Z(™Af”23²#3E²9ü‚}˜‡}X`_vûi°°/öU2?0û	ëæ÷a~DdR%UË»«Š~™››™›Ûñ	,àãÌíCk_ä"ú$!'±%ä)N µ•¥8Js°¯Îê²ŸÁ¿Ö»™22íªŸ[ÙðLû¹]µŸwèÒß¶òƒïèsŒÏ‚ø¬œI‡/©M­‡½0<‹<õ`èQC[;vjId€ý+¡å°ßRóì}‰;öÒ¯7­nR¨´éöB`Ý+ë“a±å°0J_pÌÊx=4bntî/‘ºßùÊìÞ&CjÕÇú>VÁç¼RcÌk²2‹–y¥|_À%\êp™ÐÏ†óÀ¸K²<gÕ¢ãVÚÉ#7¼	ð±?$çŠ—„vúÇ^WÓÄM*GŸÄ(“AŠ¶(í¹cÓ«M@ºcH>±H
N'ý`[o™2°ïÔ9›b(wI[£ãµÛ™‡]:Í“XŒ™8Ý	‚f#®Ä „JPV{ëƒmgRž±Œc	êLá|Ì›~W\ìb!±~6‘~°qQ(þ<ŸÎÊSˆV<;/Š1÷Å”-À´•°ž93v@ÊV„ŒÝÛñ¶è¶	ÒRçœàæ€ü#ÐítRN<‚ì•klñŒ¬|s°D‡ô>JDÌèŠ;¹\_u²Žoá®n‹®Æ4k 3r4¿F$n®ÿÓŠ¿ÇðÔoeÊxä•…fçt±2)E°—21ñtHˆÄ ì>ëG8Ò¸w:	Û,ëÕBP„6G;ÿ'+þUžÎ.	¹ºÊÎË>äc‡s·±\o³k_m’{ÃóbÅò¡ËK>sNð8¿¬æ3yèkË^–fw á¾¬&ÙŽ½Ç–Æ¨eã”5‹îÆ*é‡qˆY‡˜}¢s ÆùÈðì#“áÙß:tV;7in¢9û„‰¦÷ñsåÚ¯`KR9Ð­
ÿI>õwéßuuÒÞÏ¨º!ª+e¤ä!^vÝÉÚvæQ,!ÚUDÉ„Ðr¡§àD|ë6" ¢ë¸‡®íÞ`ËÙð0&”÷|‹)Ëžóq¥ƒÛÉú`ézâ!5;T³„B“Åž”¨ÄõûEF½–?åYñçy9)aL …=òbÜyOR˜p«ê¤ëßÉbŒjœ¯;ªYQ\­
&-âEcD#AŒìL4ÊIP6IvT‘­Åý¹«E…)cõët”)éäŽ«|ÅTÂÞ–JX‹ñbcDÁÁ•läÁA3âh=L®(\êúÔ….ÿÐÑ¹Ë7-˜ÿE˜Žàs\àÇØý€oºzþ‰giqíYmw­ÍøpîÄÌs
Æ~ÄíÊ¶NfÍ—w´î7P×ÀËÚDõ‚_nÙ`Eà:LÁÙ6pR‘åƒ¦Ü½æÇÅƒÐÎ†‰ƒ £ÛôÊ€‘zÔ6]¨¹—W'ózZÕk“
²Õ5¡Q6ÂhF!ðˆß3û¤æ½¬/±]@Z„%•ÃÛwùÉI1™íwzÃéÅjÿÁ>#XJ=÷¯tkÐ0êpYÛù‘hèº5Þ2ºmplŸ(ök©òÚ#kc+öÆýê|œŠlÆ>\À7‡O‚ˆÞ¨d;É`ÀQÇa]¿4´ó;^Q=aÃÝøëÍÉÅ\¶f2ŸÀÁ
T/£cÑU<+¦D¾œæu"ŽZ0«Ò_ù¾À6;òÞ3ùñ´Î	Ì‡ÅéŒ jViq}+[£(FçrI˜“Ãö¬D¤øŒ’èëc2$»b@€]ÔûÔ1«æÌR¯×ëaÄ’7‰ÑEmº…àŠ/d®4!§E$!¡˜å&›½F6;†Ä[ê”B0êd>ÝKAROû!þ±Ó~ÊÎs0ùv äö½G–²>)ïC6J|°ô¸NS|µ0–î’aKîä@Š¯<²îy¿áG¶W5¡V[ 4çÁË
½Èÿ	À§{ëìDcêËÁ4¦s¿Æ”÷I•¦4PËµXö×‘—¾[%ÿ 1%Ã‰OËõ˜úøÚà·C5·ô\¿:¼ÕÍT€8„ÛGu_ö ¸4†ÁÂð×F¤¡½L²!S |WUÃ"¯¬¸§DLOÍ¢o"»ËÞ_joyöì,\•ë‡Ïç¸±Ø…ÿRö•Ù9m[éYü¬]%"´ö›‹16S?êX•M§žãgz+†ŸÍþ+¹	Íñç@³—K¯/nÄŸþÚ‹í•”KCãŒÚËÙy¨Æ\wóA”TèÎ#ÔÈ	ãMè«îÚ¸¦ÑÚôj.}–®C7Å¤HéùÚH*éÿÖs[f_Ep€Ó¿ë‚Q:
2Øøõ°£zœÆ?.GçäÅ¯N¹}óÏØáœñÝ1*ªC9™ŒKó™È'¶µã'Ã…\"Wl½ØÆöM[js_ýJÚÕm]¡—T©ÙÁ Y©¡ÇrÞ1+SÒ·•@Å†EìÞ¨bÑdP[„!½xB^P9T†Úh{?”®ôKº:yÛe,…ß5HB]š=ÍéU=ŸÀECOaIŠKÃ°~}3pWeNŒâ”!y*ê3ËxÞ­0“têéÖÈ0ï•Ú[óN¹íu°Oz1gkxSoáÐ«µtÁŠr/j°ìòŒ<ÑjÒ¸½ÙG¦â²þí’ð™4ŽÓP”òÂC<Ü¿BY ¯Uç
¯i×ìg¤ãEp[õ,q›Ivð+Ú/Ãoý^Ï¸¬öáƒÇÖkYf›öà ôÓtã7ºuÄÕ”Ïú1,°CIÚ¡8‘Z9+ó‰î¼¿+,N‚%=d®lFDp‚¢Ç«Ò8Z=-Ž^x^Å	íë=Ú ¯…GçuÁÝ4	Jº¥¦Àn:\QÄ.€øp¹ßŸ«ùD´\1—
Gæ¶î0Ë¶ÃDfêIt•ÜÍ|áÏ@‰•C¤‰h¥]'ö64•Fd)	c£ÓÀÓÆ¶›’:U…‡Å“¥âÍÄ¨bEö”’F+ÞT#ÞR¬båmd	ËàIf›F4)k|HjÖvS°ÂZKHÚiæ…ER·Fšâ³Ý¶Ð®ÑÆx|m¬94+V¬€Ü¼­„d®XA‚wó£)]š£‘†–	i_±âÉµè(-×5V‘&–€ÿÞd>I¤æÇNÃŸ%»ÝDdPö<˜ëÇ-×1ÙOØ‹4ä÷½1ÜlãÛ¢lSó&qMt=}ø=vö@^ßæ@w.¥~-5ßì¡Ç>"^Ù¸bP¢g8H\g±T½¬óé`	0Š9ß¤jC”û÷+ß .ÇïÖp?n¸ 6×zOÀðqr¹¶¡Åk¡Ú_¥ÿ&kpòt¸ç¥Rý6Ê9è™ÿÇ’¬%b¦­Š®fØ±µ™¡]ÛÚØX¿ÕniXkü§ÖÂÛ[’Và*Æ?}~×¹« KóÜçÑ&QïÀe ‚Ú±†RÆº# p}Þõ!ßúïæåh´ÜÃ~õJHmÜéÑUlyÄ'˜P1ŸW½ZòL^Ldì9æ¯Këu½Rà!Ï£‹öwè—Ñï}‚°W)ÿÿþWÿ{%ò
Óp^¬û²¨i–µ¸prœîl­¬ð» ô=¼>±,â_›ÀXB^Å-1„•¦xÂJCl]5ÄÑŠ9¬p+3ª€ÌÊþEÀŒ•ëe…^Ï‘¦üâ”˜ëÊ]î?Hé°&•c°TXs(gˆñ…œ¼ÉúûHŠœrû0Ó5UÓ¯àö½ÂIl²;¯!sï˜Ù{·ô$Ÿ<+Ç‚£`¾¸üC
AŠ©V8àöŸ.
jáå²‰æÍ.¼‘;±f†K-LÂÄïàYQø—ªïñß2 ¾r½‡Ed{h!:zuJoüix`°ràñ~›‚•`è„+íû‡åˆxÝ¬FªU³¼]Ë¾Š…4½rýÖ'ð9ÎÍ1þ¾þuö¤˜åÔcšGôP!C/…5‚_]í›Rï Ò±{n!ô~<e?ÑºÐ›×J-jÍ±ÅƒQÒ€+½ŒIÓ$jï.£VHÁH°ÀYìv F¦ña¶ÙmP®œ Ç™@X’H Œ?	øÇ¯5~¹P;å†Ö Ä®=Téy1g,ÉQ„45Ü¿šáÂ0Íã¦fA>0Òd?€]šíEübQ$ƒúÑ6±…\¼Ã»Ä¬+°?¬N1E6ÕÚ°Ôia­M#önÖ¥ª$Ô,ïû
ãÃÜ ?bÉ§•r1¢Hæ—6güˆUÆ/‡'‚Ì~ª˜bÞ4Å•Ôk‚-„Ø¸Âòk}2¸Ø´ŒE&³5N}‚~ã.ÅŽ>ïî†1ÑrÜ°j%Óô`J„é¼Æ_){Ï-~XÅ«	­àÙ†c\ÙqpÊY>,O‚ÿ‡b<˜TBÏ‚B%p9†Ì
l%]O¬¿ÐŠé¨ŠBêZ&¯È–ŒYŸÙÙ3Ÿÿ½ÿ@,–]”ÿ`É¨ô±
ÝªT#lÛÎ²%'Ý`ÅÊDyØ{=ÑS‡:ì'¦ %mF]?aYH„»2Å®‘ð+FÒ	Æà…jÓ`ÈauóVÑL8à‹_a‡0z½ñÓæäâ§-òO}vœw7Véÿz[+o2µ"c	+ÜÄÜÃ
-KÜvé­*5™)¢“…´Äv7I6ˆhk–ÎùÚl 9&Ð«'?øO‹õp†$/ScfƒªÈí\#j}7nž¬vY‰³Ù	+Û%p	¯—´u'ô]£ÇÉ“9Ô©89Ð D<;äõÑ¬»“ã¬j³ê,_rU÷Ã@lU½UÖ§Á¨Ô ¥N«Mõ	p‹î^Š—VK¾à¢Óæ\ÊyÎÒúnÈÝÓÂX_ã¢ì·ÒkÔŠ›;*ú%Â¤HÇÆòTÝÀ“ÑÔfq¾ŒØéÅ´Jœß¬:;Ê,ñG³xTŒO„uÑ*bÒ6uH[ñ)ãØuDæ˜\R+íÆNmF?;†*‚MéFN4JûÔŽ/—Jü›s±ûF7Ï‘4JþfkHÈµÄÉ&ñ¹4ì»R²}è`g–=šËçÅŸçy]ˆ{Äí eŠ^ÂG1«§Ö$rµ&ÔÓ€UçY^ÏÊ“rRÍ©{ÉÑ^ÉÜ4µã˜õ¢þeEÓ¦°œ2Z¶úh&nž_onRr øqc&þ€$]EÉ*t¸ÖÆ»þuö€Ï2zÀF‰+âC_ÎÑƒ¶ŒšÃˆû¶›Î9Ÿågù4›Ä](ä@š‰ í|íÜrÕ}-ÎÊçÏ¾é¼P4ôøÎÌLH@]rúŽÚ¢ SiÀ˜UaXÕ&g(ô|N›*¹ˆÒJ,Pe>øB@àjâ>S’4ƒ+\°hQ3®QÂZ›j\DÍÞ§ÍhvˆJŸTiyà¥ƒÙŽúÞD‘ÆA¿¥˜¤Š€-W}±´¿£I‹UŠæ¢œŽšî4(ñéÎ"â“*‰R‚(MwnsR‘ÎíU‰8xŸ¦¹ÒOûÉ)ÆÊÿ [™&^˜ÍY†ñ¿®.g‹YDöŽŸÂšñúÛ«ÙÕìîj¶¹asúî nÌ½Ûðîöœ›òí;(TbÚKfÙmv;vms†<îÉæ&fÿÀ¤ÃæƒD¼œÝÒŒÛ7C†…8=ðùM¦.ù8½!CQ.?ˆ{IÛåPw'Z“¿ißˆÁ7Ûã@OB¾viÎÚWps3Oó¾Pç¾Â¨iÜfÅ.£r¼ßÙlT#¿ 5šÝñÐo&îR§×&+ê„†ßÂÛ…]F“¡Dï½íbi¥æAÍwÇa¶ÑbGý@—Ôz¾v·ËRÙW;æÕ†u¹ŒkA6ô>	Ç.àp­v‹£p“føØÅ˜Pã4N€PüYÝžIcÈ°šÏ¨/Î¸zlÃMý‡/aw¸éžÎgu4,–*‰êà6'š¥}eB	íñê:ÐC°mO½¦öÏF £ßþ6Ó·rY´L¢·5“h¿WŒÑ+=—ÀŸ)Žn7=ÙÉtBN³kSvs¡™åØF(†äãÝÂ<õ
G5Ÿø7F}öÊÑY»¦õÉþ@+Œ©ùp¶õö!ù.û
 ’ý.Ûôùö €á¡Uy–Ñêøg Ó	@%¼ßëâ´¨kÂ†*r»$2Kµ&…*IÃžÏÀ’•é "}žßˆ€>þI¥÷óä,ÆŸ÷;?óñ»†0ñjRŒ	–Œ« à)xÊñ´˜­m°$ŸùÉ»õÛY5ÉOÊá%ËÖ˜Ü-žZYÁøã¸’:"„G)2Úå!•DÆ¨ö.žvÒ©@LB^a=¸±éçr	¨Š¡—N*ƒ4q8\’?d–=eƒe¹L}ýã~ž¡™vÂpþO	+€^ÁlÑýîž15\4_–(üôâÁÍF[Âá˜ÈÏÂY²4Òþ$!5U¢†D†×†4Ž4Æ6šËQÔÖÕœ f^ƒ]“AM­HfÈ§IAœí…`‘Y5øÂNvo™u•Ð#WË¯2íN,eŠcqòMH˜^(°­oER”ÛEbÜjn}cHx»,UŠÞŠ‘ž$”Çi¨MB¤<)I³&¼$¡ðÔ$ìŠÿn¿uÙ¨\Þð?A9£ÍÉ±a®M$µÙÚÐ‰	ÊHì6Î¡Ò þ…Æ0K3Ò]¬ås"/SÏ#%·½·GvHåÛ>ñÅïŽåKRu½aN7yÆïåŠ1¡f„–Ö3!¤ôÀºýH¨1i…º¹ #CX,:\káo\[1{•ä¡ŠZª-¥öt“ŸÏòãaœ^XX£âTš9Tæ4‰ÍÞl@t‰fµÏg7{ìNžw¢	uèxL¿Å~gC Tç€{.¯g4¤ÃÞúlÚÐ-(N³ª*°F“Z"[TzC¢s¯ÕÜÿÕ?rò¼ÆÐÚ³È{³ãª‰5Ô4skÍwgOEG6Áæ‘Of»ü¡Ã0wS]œ!{%þYßYÁ J&k*)Š+±Xb¼¿°
Îr^Ù	&ÑŠÉ³;°6ŠÌ“0A‚Ä(™¢ïZ¶fÑ¹ÏÕ6BWË]­£tEâtµˆÔ•%ÄêJÖÕ0^—®VPž aa=¦4†éJ±‹ñ‡êZêvóåp`“ôPNQÍGë8NYö6[8”Ó2iH@±ÂJX½ÂJ£ËRµ°Â[[íVºJ¦åhXµt¸¯ÙòÛ¸ÃØ7 ¨a¥½º†•…”6¬,îöº`*"54Q’ßíb)‰Ô´D‰~¿a]M”0(2-KUƒ6×Vaƒ6ÖFmƒ7ÔPyÃñªpÄû›”CºéÎ–¥,Kd†‚Ò&;¯—š!
JÃ,QPÈ%ª7Î¥qÆ((d‚’–9
JÃìQP	 ¡NfI	IUšf¢‚Ò:‘‘*at¾¤U	Ucy,–ý•Hžµ±šm'°A–+(gº‚Ò*Û”xV(²^AYBæ+(-³_AI›Ø’²`±¦ÚdÂ‚²p6,(2bi½.’JËÌXPÒ–¨A†,(gÉ‚Ò*S”´)-)c–ÕÔ¢Y³¬æXæ¬tˆ#¹µñBmò‰k„¨ÃõcÊ0VvÓx’û«%@Dv°v IÐ²ÒZGÈJ"X–šÁŠ›Å,Ì­òœ!¦ä:ówÛd“5Èy¦uØ>ï”V¹Ï ¤Q´F9Ð ,!”–¹Ð ¤M¬uN4(á¼hPäF£&Û¿‚	3ªN¿c¦F“7o*wsÚÍ›(Ü Î¬ù¹2ÒS¬Ñ¯ü›ZœÍ§[ƒâ½.ú*WSÑ÷‚Þ- ñ8•-òô”môó(Ì[ÞsàwÝ¾+•¦Ö:´¼öª\þÅ$ï_í\«™dÉÃ›ù4ìyø`R­zøç‘ëèÖ>¬,ÉÎGŒÕoäÛÁ\ü"Ÿb–ä1˜è˜]"Ãv÷ªä„ÝÅKGY^}ÝX^š@þãüä]¿®&@¢k°+úP—(ö`
Ñ=eçL½”÷¯®„‰þn¶±šA(d"ãmô¾ÝÁØO>.Gd¬F­MYk«b^ø2*nð4Ê/ÖÎÁ<1¨ÚB¼ZÌhP{°öúÛ÷ƒ76YDlf¬E˜¬Ý	ØCÅ3 3Kbùê+d±ì¾Tu…® c»<»:¬Þ‹œ¨ûª/]ˆ]&ª°uaæ4NýÒŒ£ä“Õ÷|í‘yü4É5&š™¾0Žé!ÿ%æÖágŠªº7ßHCÕøÅüxTÎö¯˜Â~aS·PVD8»­6Ç%ct_ ÎÓ>{TÈ1²›–Æ˜¥èqI%/YÏXÆÒìk§ƒð0äŒ^&2’÷0Á/÷Çââ$h¸•|n^Å€¦[v°Ê/^,Ÿíkß]p‚‰ðr YbÁ¹·ä¨7†‚á$ûàb7ûnXÎþB†žÝ#´®®²j.ìÑ˜ÄÙœj+=iƒCÅTŒ¦¬	%nó
7‚ÉYö²œT,ÆÐ}àXÁã¸1Šƒ­†ÆÅJi`ÒàÑ@ëÈÅPY¸~ó‚ýƒL]N7î}™”{£{ª{b5|hB¶y›ÎáÁ-¤îÑÍê†ì¿±êáKø¶I{nh³‰àœ?–ã“9dÝ§F}|X±ì]ç(Ów^ðöŸE !£4;Ö.	©…d=›l_¨5§b|Ó„t2(&ùoçó 	ZÇ„u#„ÄUÇG‰ëy¸ù”ñ'¨¹-«ÑíÉ¶ÉY¼¨‹ª4ThD›Š4±k&‹þ	(—M­ÐZl÷…Ébç@(ºÔ®á=#õJœB^9Û" I™üJl‹0—§Ñpt·Ô:6¿'ó×5ô_ÁVbÆ†<³è Ã*hþa`L%˜‹ô‹hè 5@»æÃYêg$2IhI©{Zì_u"/¸º6ÀFmÊNSbvåqC»PèVŒÛÿâs“¼ALN þ¬ÎÛ
ÝáXf,r˜Fûí‘C&"B(»¹OWj×ÍýLaä•Í}äK9þi	¶—Á§âsÃc2æO‰©ÑéVïnƒÿPÕþZ—>C©eàï¬ð×AO~Yÿé¢§´.ø›BÏ  ò1µÐÔÿì×93ŸÚU:ócëg¦u~Vçþ)g:ç6ÊeäÅGF°xÜª0šA_y]øTŠG¤j‰O¡“U"µÓ=2RÉÝ/„¡wn’77ÀÑ8	_Ÿf&;ŠÌðÃ_³S0B;¼‘Ë¡Ñw»,6bmdëÙÃû?N³çÅtBïþú¾(Û örì/_„T¯ìh<Ý¿24÷Kè·„^Ù™^ä!†LˆÞ£2;(ýß¦w^^#5(Àa°@Å]¼bÔð÷0{í­¶ùDH  €eÓˆŽL*¨iW4Óô~æ]U®|ž¯¹1¥vâ¨à§*-’9TÐÎí›7qn^öW)›Yk§e=N?(]Êõ ¶‡ú‹×o¨†:6zÚÀ$©G<´OúJõêmlXÌö^ÌøÝ½¾\—øÇo5DMÁûÀ°@-kk%áú¥Xà…Ý ˆû£ëÖ÷(QÝ28p!ËnéØ%·	ª~©PÃàËaªõ=_épÿñ&º`ébŒå»Ã7wbv÷$7hb-æ‰£“‚äŠÌgÅÜÑ¡ÊÈ[2£÷lÖ"‡D=ßÚ{ï®LøÌ`QjÀ€¬¬O„Ë‡ÃÇ £ìw^VàkgWõÕsÅLqýõ·ió¤˜å`#X¥e×ê‚bÙDhî@„çð¤¹Î~º£ÎÅíò‚¦|äKüñ™)®ÃÎç­ƒQ9]¶[QNáÓÞŒÉnñ_¶c‹í¸åÅš–JçÖÜ63ÊU4Ó-Ï„±H¯ªÃ¥†;qÂð#8²g÷K¸¡É³gä$&så_À?¾ûü«•«NÈv…gªC7‘>î[™ìGYâ³¡f7(è#Ô}Ag'ùõU5`Ÿ¬g·«ÙòÏÝØonüXŽÑ÷)¹é¬˜zŒf)4Â*|º¤°Qø‰/¤Ð1›k@7wz~ãO@LI¦FOëÿ›ýe5#d¨Ð¯ró²ð!ŸøÖM°òeã.²qŸ•~uy›×/Ó@òëYÚR„8ªáX ãüàÉ¯mIh‰"CHüá«·í$´íû“Aqòî¸ºð¯ý¢èëûÖ
þôéî]'~•Ü=|V7¹}8:XÁ%´<úÄFvNg@þ±¶Ù6³*DH[åRw#NôÏdõ]SL[JŸ¾Á™QÝ¥ÈØÍÐG}çóY]ÏËZ‹=–…§}kAÙ/)6Œ¾õo`—äÃÂmATfõ]¿CÑŒPáõ
Ïé¹W¹Â„••lVç'ï ‘ÏK uD¨	‡5tÐ
k²àŽÃŸ+µÄú]©ÿâ0xU”Š,¤C"ŠÐûÈù­ŸO!¶C•‹éŒÚC¼/ûð¦˜–‡rÄ›ÂqRNò¡?DXn&«¾t4p ä&´\¨šm’qE’	x—ˆ^Oxvã¤°ì&²da gEøæ˜}ïaQ“2 ˆÁÂõv<BR4Æ÷‰÷œ±‡ÆÚÙç)¡P©'‡>­”dÈ÷²¢­]r"ï¥.+Í–•ÄtÒl†,\<\óÒ¹&vqŸkÜúÃ.Ý²O¨ìÓ bm¦÷%-¸­(º ç™p“”ßí¦ö&m]f%]RUÅ(U TZ3ñÐjô«”}Þ 1(âäèÂ#BÄ˜¡GLÅ‰é=Nó|9¡À¸”˜L^ñŒä©%]Y¿«2‚ÁyÜ¥¿‚¹!ÍÂBî3wP'’’ë)="¿(Ñ4çñ´/PâaáÓâÀ›-&B©ÞA!¸·CS;ƒ‰YB/Ë¡“ˆVP£½ÌÔ£©˜%Àa'] $ ”´e–Ò!Æ`mHŽÛy}4ën¬ôfÕ+8@Ü#ˆnbŸ‘lÆ§,—Ñy4W‹žðÓÏ¼„ŸtOï´É¨iÔÕ0â*+ÆŸ6Ó” ‘Pbñ6ƒfÀo×[­Vó=¥fÒù’ÝÀ1‹æ.(ÉÞ?™fù)l}òt!-û7§öSa¨|
ÀÁíä£ø7FjpóîÅ¬çX2Òˆé*®T$6ÈÞúà¶ÿR¤µVTQuCÆYö‚ûf‘ ¸‘Û”¨^5õF%~§‚9C©ðÄKKmÕH²m"Ñ6
äÜîî$ñö¤ÝýI$fBZ6¯¾E1ïQÂ«g„GŒÀì0{{4ŸU£%”²ÚÍ¾ºÝ@÷Èøg×o#±„ÞÆñÛÂ@@V+Ìc@{
¥	§–1]»’åñU€¦Ž«lJ˜J1Êw3R‹hêZðÈ_Ž¼
óÖÏ™ÀÒ?}Û0°ü"ûY¾‘Y¯].”_/•5@õ‘élKSÊå{¦úu/iæ°n>ûð¿³íìÕóÇSŸ rEµdýøÝ¢FNiî2vÇ¦$PŽRÿ”ÌÉÆ~}Kf§ Ùc¤jö»lÓŸ>
g0è&\ rÂ.Àé'ÛÏ^êo
©ˆ.*¾&ÃÓ( b$øOJGVÕÅ7mûì%U¾ `-âk.Ø˜È1¥a§ñA¢NOˆn@²­ƒDím3ÛÈ7fî‹|(c½ÏÄ_"-IzÐn7Hú‹|ø‚¤ÃÅÅ@C¥±,	ÁhéþøÞÉÛ~¾·®"ì›ï°lêxÁÊ	º^ÿæ7§ó1›ÌÛ/õÇ²8§ò:)„IaïÑ”qùIU¨_Õøe•OgìÏï‹1†ÈøÏ¨qŽúzÎý´EÝX?ÈgÓ£ÉžW³{Õø´<£Ë?Ï‹WcâWÕšÔÕi9,Vs½›‰QîŠYˆI¼¦Düp—ª0Ø:xòDd³(«£“¼bïøTv³îh7›Òìá«ÙŒ|L0øä¤˜Ni,á¢®«ºCÉÖûªìÿ™ø.}
ÃqÆeÔcÀÑÀò~È†CaE^‚‹ø36g1Nlä‹îdP5æk~&Ð>Ár¤œ{Ð÷ÁïuP“¾ÿÞû¡h…õÍo‹àóãžîãþËbXœÕùè‚¹§åIN1m?Ë§—ãÎÆfü“?ÐØôO©ç+Ÿ}oÍjÌ|˜o¹M¬pË F‘ÙxÉUÉ Ô¬{¢2™<¼£®Êzµ|Rþ]q‰×9bï´:l<¢'åÖmy«7¬Îº 8HD6›Á¤íyM³u¨ÎzÉVô)±½<«/H+¨‡´“ß¨+[cß×,ÔxHççy9ËN‹ÙÉ ÛY'Ó^'ƒX#+}qÙYÕxó¨˜*²HgO_¼Ô®HdwTÝ5Ø8ä•„F¿$< C*åš‰aýçi5Ö¸VB¢”Ýì/žþÐcË^ž^g?	býxä™gldtkdžkD("Cc­×ºÖ°!™õ´ÍÚ¨¾ÚÁ1×Óâµ<Ù]»kB#ö‰%kÔx©¤¨&ßUï oñº=Nªt¿z‚”ré"dçaNv/‘Ž+ºg3‰žbë =Ðf(ê¿¿WçãŒ‘Dõ­éuV	Š¹CÝðÖî1Ë¨´Éà{ýVµÄ?#t2 äµ®Ý-Ææ×y@ÇwB8?äÉ‘Hœ&³‚¯¨ýtM>°5x=-©¼,êÑ*ˆ€/äO	M~1#t¾Ûa»SÔ¡‰šR+VK{ ×ÛcH}Ðßð|j42gWkå•þ$µ™iq6/iêv}<ÖCoc•ÙÏ('\Rkïžûü7MJãiÿ½ÖØCñ+2*U¿œ>!mŸNŠñ*Ë#«¤ÆðŠ\ŒãLyUËÃC°x7áþRæ€m!øsõ7Ù%ìyTÎ<9ìÍ‡È,_¿9è¾~ƒµtOÓ0ÓžãPã [_Ïž¥Êúx0 *du$£1˜ƒVûI>QáôöôF‹¦¬Ygòä@äða3éê}‘½Ìw³aIú{?Y®&ßçSB &UM’£³SNk‘Jt Gâ—Ñ{šË©|§¤?ISìyó=:íp$àõl<xO$è‡àÈ8£}þ(šžäuŸ¶­&ÙÓ™,¦vï”ÕRŒª®nAÈÍ!ê:¿ìQOyrdÍ^³.ÉL£ÀìÚú=E'V„}Üw!£ùxÒgˆæÌj¥7%Óg<èz5{í½‘Ã¶/Ü	s‹Üõ‡÷d¬œ
Ì‘ !“±™piF 7•û»ìØÀ’¹Ë8êsCuÀ9(™ÀãêœŒjQ ƒ¦7«èfBaòÝrzO˜¿(ê÷´ö¼'íŒ©ÂWwL©LÓÉÈÎ•me‡P:rVÐ¢±c©¶£Ÿ
;7óKgâäÇyÿ/±\3F“õÝLºXuúˆ/;ùûÁé)YKcù@61‰¬âÓ-´¾í‰ý%3qˆUÿ½Û†FífÜD¦<‚(Þ—jCÅf*wR[„ˆNEFDòA1œ pÉiŸ,Ê¨„\j§ùö¥ B÷óKºÈy?Kj§!|©.ÀÞñh&BÙ†~¦á¯é`ö3£ÁwòM_Ü@Q«Z:ê.úYöu¶¹±±!T£»žÏWôÞûåééËRÄßb_¯ô`ºä!Á¢5}|ê¹Ý‡É“|6è+"óÉ†×³.ŒŠîŽü×Öíÿe{ë7ÈàÅ£ß[Ò ¾GÃ¢ž%®À€	…"·XwŒ¿êÈgÌ‘ÍÏFšh@6éWûÙæŽÚa¼A%o‹÷Ån;u8ûðÏcB!49šP¿¨6›|û¼ x>¦Ù§ðjÔˆ”|HN¬ÚÓã3fSWWSjäÃ••ô§ÔUÒ_w»’LúoQ}£Ûl^ë'_}êß$ÌüˆBíŒáø¼±/ÅT«ôÕÈlÙm¾üœ1ÿ™3ÿ*iÖÛ	³æk˜ºÖÈŒ/‹á°:GfœÃ]¥6aö[Î—ýŒL—}ä­5-5)jv*ëèSR2¿Ó•¯|(¦"YëöÉ›†}²>gÂ‘q­ˆ¡÷ $žRø~Y&’1zÀ-Ò3Ò²Ù àR«¤8,Sã+ú;E*µY©4d@@ÙkCËÇ@¦k‡˜&uz¿:éöá"
ß{úøñƒ{/=ýáEïÁ“gÏ¼8zñÓ³£ç÷<z~ô®œVtÝAÌvUïj‘¹¥ýÑl—zh5dRm4é"
®kívøÉ¿xä[ŠY«Ãþ.a—
ïem8ëgy%«×| =rÆšZüœ¯g|&¦ò1ž¡‡+‹zDãv
u &P²OG0ìïæÓ“\zëˆÌ-f%%×AëšDÇxèxò3*¼ZÕœ®Ú1FÓE0Y–È¡m—OÑ,]Z§u.)Éæµ…Õ¾>¯h§YÕ›Bæ%O,jn–^Cö_ióÓªÑdf§C·!{B4!ødéÒ™MÏêNÇûï	†IÝ6]°…×òc²Ô:“h˜(Ãyê(b>+k}ªÀg¾0ád¾#SaÇFJ¥|Åâ!äï~>Wä¸ºÛÌŠ	n÷ÉÎ–;lÙ6ƒÑ»I	¶%D$_»:~úý}1.!—Ÿõµî&Ã«º{ £Z`Ç{´›#)¢µíÇ	ý=ÝËÇ'Åp‘ùˆüW,½cûnx5Öˆ§+ºžÎ	éNéÄÝL¨jûYÍ¨z«ÆiåPëßÕÕø/EÛæYm«}‹ý¿ÈßŠí»Ùó"?™õ@‡ô€ ålï/Ÿ<¦¿†lÍ]à)Ý/NsB+LÞ÷ËœåÁì!ÿÙ%C×5©õ’Þ¯¸Ä—…ªF”1*„±ÁZd{t”˜Š,8º°WïîPÿEýRô¸Ñ¬þ$¿VyŸŠ¤ºXNd1_8Mv;ð¬£©·„\Â/Sr*gšU´WXM¸m;¥—¯f5ñ<Tç¹¿Wä¬¥b”—N]úýš§Œt*ðçXã¼$R”]ƒ=Å¾?áùxÍïÙSìûa9~÷$ŸLíâ9Ú‡¾©œ®ô—XmÅÿíªêÚ+ÇœÎÈ3´SZ¶ºÑ_ú`ò"(Pè¬–+_©Íò>BžŠä†ÜõïÒÌÜš»æOÙoTî¿æÛ•)©<ÊA¾‹™æˆÉ7°o è€C ,Bœ×(@9½?g—Ë…å§d„’.qy^Upiœ’,1äµg´Ž9ŸÃ^i’8˜³6<äóÇÍŠ0ÐÌ 	)7Qa/rP«g4ò”s²±íÀüP?k2Ší˜<¢&m¯•HM­kTáÅñ@?ž†ôÅîáU¾ ‡ÿÄs¬5F8Öj­p8 W×béL šáêÅj²ìíS0’;øŠlß2+!Ý!?	.sù˜ùsº¸jÕ¿ºÒðšEäa-Ô°åOrÚDzÐH¿è½5­¤­’þØ0P„yé<ö·¿5˜,µ%‚‡¦™‘¨Ÿ‹Ó3²ã>ú”M{K†%æp»Î”Va° ®Ó"Œjp(·O‚€èeG	Ø;>°€±·~|Ð >3¬.6±U>	@Úr|­ÔÙ÷Ðä"»‡„àBã8?cÆ~èõ1`&9ca×w@”ž?}üàEïû/^>}þÓ½§O<¿÷èè±úËMw­ìÝ ±ësÈªS9ÞÐw©;GŠâý÷Æ§•)ñZ_Ro‡?¿†ëT-9Òp÷k²olÑq{ÀR»;÷ë#rV£ºþ*‚ŠSt¿',‡œåV¾þÓøOã§Ã]Í¾†>Ù\i_ßúÓøáYc¨ª°ZÃüÈfï}uÅ¡p½<# ?ÍßWõªF9”û‚æÞ¯-‡»ëmÏ}}Û¼Å’‘Íì,ÝÎ®›¢5¤‡èâAjý(1VÞ7^DœN-g-DË\"”`™k«þvìÐt2Ÿ•Ô@ç·:Z£¢§	6Þ¼ßÁä¤r³†¡ÄP¬SÏ5GÉ2yƒP§›k¾È+_ä•/òŠ,_ä•/òÊyå‹¼òk•W”‚—W¤ºS·ŽÖ,¢ùKÓüYØÛ²AxïßÙ-ÅÃ’TkÁ¦cù^Q+ÒkM!~ ½¿þèåõ2¢¼áß¹FâÚÊœðøBÇ/ýÙ¸ïƒ-´v¿¢Ì*,+®órÜ¯Î{ÔA§u;÷‹iñ3Á«‹“!HqLô}ØYQàÑ\Q¨vSm]õ¯‹Qõ¾ìç=}õQËmˆ[¼§sƒ?¤™DvcA.ÉT_§wX~d%$L|K!]8~£É˜¦élujkµ¾›ßÙ+f­“kÃnÄ06}Â´ä—ÖúÊÙ½•¦`·2¾æDŽ¨!fJðšìqxÄžùæ«+ïH®eúO¦o§*óé¡”†™!¥‡ ‰)Ã >uÛ' ØP«®‰ e¸¿;“C,%3cH¿ûK„¢¼ýêJÿX…ç‡{ Ûä8 âè­·Q¬Ö,uàê˜ïß)ÜS´Ì3è]Œ(wu}pAÝö3×Œ½ ¯î#—ÑÔ_ØI7vÍ"®'jÜûáÙÙSe$Þüš~£]XŠ^ÊKJø@Ü-Š·ñÚh¬c~¬]DšxÀ® á[zñ(ŸÓKÅÿTñwüŽQ¼þŽß+Â;v›h6{ß#ÒY3³ãýíÏ2ò1Û‘×t?ï"æ9f»Ü¦Î5Œ2>ë<.Çï2zAÉ¾÷•ØgìÊN}G[ÊË·LÜßU¼†y=nÖ’¾ÑêÒ¯³3·ú¹*Ç]Àqq='ö4ÃÖ—Õƒ‹“bØU¸»š	¦3ýI:‰³½áî†G#±¤Ùö A¨
ÄrAæz¤ø==ì½Þx£è÷-x¨È:{N;~XW#6~ê#.YüHLÆp1‰¬4ùþ1Wán]Sb²
*šwÅåT¹B;Q+„ûÉùßÔxûéñÏ„^ÑZÐŽïÄFÞ½†Ïð£ú)D<`nÈ¼q4kt÷Ñ;äF“´m>5õ/–<
Ð•}’£!éôµøý†Ê¡®Öæ7>…ÿVÖÖ˜…–n›‡ú 9àÊ¿pXy=™ußçCDVSC¿E>ã²l‘~ï,×{
‘JÜG÷­óï™;´Sýi&cª…qÁJ6åÐòñÑnLÍÈ3¢ãV54Hiò¥Q>³+]i­Œsˆ¶æBÿÎ€6ú¡6Iù¡>E|ºþeä_õ¡Ó1•ö3Eû•ò¯»‘Ù§šM·MzAQÏ#öŽ„Æí^ŠÊ;†;'	¯`×?É OîG±É[è#+r€‚W	<â`Šš†±âQ›ÑXÇØþ¼²ÁŠ-ŒŸ¾!C7Lh±‰=b_€Šñ•f“¦w%+Øôóü|q¹kÌY¾÷…&x¿‘8”û›Qß í`"Éˆ†ˆ1û\­E‹}¥7Kr,Xí¸ñ½`A·.øà¯,¿Y'ÀoÒ%V&—gVŠ%8Ðe'dîöL-
O :¢Ì4mÎÐ·T4bØ@b8*±”7mœV3ÃRRýd¯‘5W«n>U•)ƒgèQ³¹D{«à¬»þ§ûëg« ¼ççM! ´™fœ­ÇM=ÑE¦g0zjå¦ŸeVF )3“§8Úº°•Ã‹K+ŒIÑI°!´-­KC(ÏÏ‡8Òs:˜)ÃÓ@SºAh
LÔÑz0ŒF“!a™°¦tkdéì[WãûfÅ½hÊGh;|¡"¹=bt5X
+ÃòU>ËcËü©,…S ¥Îç«™nLœ¼Bšnrì Ï;†ºé8áÚßZ/¨z(E r—½‹Il¸p#­ì­ñ¹¬Øf’Å‚tõ¤®htœb6¦~ÞLß•“‰ý‚‹YcÒ«$Â]h’„vÚ&»´¦çmz"XÁä~0Õµrl){béÖ1¦§°þ…éOºÐb$½A>íjïì£Ö-Ú}â_È¢„öfæÅpH›¾á(Ò¸èT&Áúò7î_ˆ¥ÉÛ¯®DOšö—=dºês¾È,Å[ùà:+Ï‚Ïzo)n^kf$-¯,x×uŠâWSv.v¾0¯wd˜W4ÐIÏ±£ˆMÏ$´m„•µëMÖ¶}a5õ4xœÌ-'§éÞwórØ'$q+›Ž°µuûÚŒ4êÄ¥tcøî¶œÈÇ[F\O-fï·H^U¡’Ì¤JÒl}°eõ8qó¸ÊˆÀ"’ªÝ¿˜†;ë	ôS—pS Ôâý‡—•Ýé$ ûgË<QñzFf£ÉÚ¦7uT;1ˆñ®V¦B#Œ‘ˆTdG7dŸêØ‰ËŒ8­HÒ%GI(´H¯ÎW"%ÃhŽ»Ó‘™B«&ùU”mK6’;Ì™ö¬œÉB°‘R\¹G:5»t¢º>Î/«ùì{HÃ¶ÑæÝk;`/*¶Í‚ÑS}Áh¯ŸÃ‚=¦q¼"!Ç-•K‘•‹¤ÚhŸév³znóy+µ™›	¹|±vçFâ.#‰å©ÓiøÙp>gÃ…?VO‹8 Üøõè¹AäUd39›;0D›.ˆÜlâÁ4õ{¯&`Ä@B:ãG\²Áø£âYœn¸1³w]~rRLfûÞÅpz±šÁìOT4wýÐOi:ƒ²ß/¬»–øàDÖný>„ähº¿ÅV‹ÅYÙ bß¯ÎÇ)ëÊfˆ­«‹Øé1“+@‚·—ùqöâ¼„Ó_­epCeÇ‰?¯‹ªÎúøÒ)»lJ†dìòpŠ½°¨Îä˜'J¦l±qd¯‡˜¡0&{á!ÇÊmv¥Åº³±æU»ò{:
"9ÒÈ¶¦Èé°®o,Ö¥A_ëÿ÷?ÿÛeüŠ…èbb°¶ð*,gmh$Ì›]š c.Ü§½jÿýÿþÇ¿ý#è¤˜¹ÑÇ[8-¢é§º|Ú?åEddøßÿñÿHÿˆœÙ‘(­™§Y—†
Ì§+.¡¾²£Æ€ªÃIä$Òâ®"šž0hŒ‰0±žÆW² [¤µGc2OwôFz	[¯Dn åÊÐ9êÊ1˜y‚,cÏãé> ¸œwÅ' ‡œÇt×Í±:§…ðå¸_žU
åÕo·ÖrÐÞ:*Y2….„”;	|š=Õ'™"¯ÈŽ-Y¥‹ògP¶b`ÅÄ1$	-?ÚXhÕ4š‹6éê=¬—ºÂn*láÁÏ/çÇÓj8'£ƒ¬í)gÕdms}+cKJGrI˜g`¤é°ƒÇ hÝnSÏÕ¡ÅêŒ¬0½¾†$QÂÕ¢šSkÕ^¯g·ÀSC©øyöÁI
eß·“ãŽ"‘Ò¬[6©…d¯' öåS2Î¹'óé.\üÀnW?DÐGˆN6,ÇÅÚ.~qá?xÄ±ÓE²of£þ®ú¹ƒ&ÆR¥“¤qÖ*]o2&ûÒZ²»>{x¢´=f¥îê8¦hAÝ,J(¶h5bøâI½µÝ
W»ÐQÂÆ$ÙSE“lóùò8‚4ÞßÞ:{­¤Y¨¿“«†„ä—”ÒG í	äŸÉ•MË¾úS ó5B˜	Å›H&¬O`3X7öKÙHÉ´Ídìøœ÷Hž¾Gh°Àøwrð€þ'¹ìwÀþûù£¯*ïû—‚ÀFŒ×4Ô5ÒàüZ‘öÊL»ÆçÞ¬Ÿ¢šÖs~-¡{íÉ_8ÇúFã¦‡ü|ðVÙg-G±â'Ê"V­Ïm«¸<råd¨a˜;MÃÜ©ÄÜ©sÑ|â¿NÌ}xÿÇ¥ ¬Œ
†«2ÿÖ¯IíÜCÍ¨+˜û+
Kù¨,/ðq6d]²"C¡\áúSMÒfXçÐ-óz;ži?ï²c) øälº$]éj4JA=œ}6›E9w½
S‘·€BWLi±MÙê;JgFE¿œ»vJPXæwç^Ag’Ô±²Âæ‚ÒùÍ° \ÉÐì.ê#½b¤Éþ~3CÐUþÚÎüiÙÀø2Ê‡H=/ˆ4âÑ“¸ ŠmFØn[Rß„¬ïæ‡Z+ö}í°u>ìÓí¢áfÆåÉt¯Ô(™t-<C@—â†ò«Õ7N£ñÎ=©Ïo´2ANÀêê›Zï-Ò'[‘_ª	Î
ÍÖ/€²¢ëOª(›Š€ÕÐ÷}|ÐjÙ.îžUw6enÖÇvc^'/Í	¸ÖwÀÚ0f¥Øÿ‚ö#Ôß´W`¼˜*øyT’‹a=í	¦b‰klT>q ¿ÌÀrTE7´[¿{3`W¦{ÍÁNµº7w6¬»¸ÓÜàyÁDx•Ï°è™jüf`ÏCz6„¤÷Îí	X^B<3šÉºgÏòq1ÔÌ/é7þhf¸­‹f
N'§ÍkL¼˜Žvéßuu”(9ÞÎòq9¤?ÍûÅ£±Ërðí;2ÓFŽ¶/´ˆi»$¯Æg`kd)ù7¬ã®5·V %%Vëë¹*7PyJ	sèùZ>ŸUœÅØ5ÅÊ<^PüFQ¬¸¦Qx¸4ëu}Wåmöv2)˜ø4y–ôÎ†%xe®xáw$qÆ…íàûÅt”×`yB¯R}ŠùÊ,	²v,Â8(ok ”G	5úD‚I3æóALó’1mMü0Òyd÷ÈÑm:º¬ML—E?ãFè,þ¡¾)ý¤_’ê¹rmSÎÿ<'KÁcf³êìlXX´3îgQ4[G‰R+²$–d¬qØßÕ¸r¶ô“
â³ŽûUv…ä+¼£Ñ«¾þ ÀTw¨Gð¸DØƒÿaìø%AaNÓòùBŒtÈ'9ÎKZ[uq4
ò?Æ{Ù?—Ý.iÃ$Ä
ÂdÔ)+û(èxÀOÛeYôºŠÖÉ¬P¬t>Ù-k¶úÀÑ‘{ÆŽ:o/6Ü£ºÎ/{µ¸+œì_CæR‹ÆÇ3Gþ¦Ùpg×Á«'cWJ½fÝ*r7ãÞ7PK"dr‰ËùÖø”£¤@ñþqôw#	Q³_‹Áe„p Hë
w…|+)°S!ë~…’œ•·î]ê`š Ê¢¾¹daé4|Æß2o“Ø‘ß¾H²—Ê 0.Ê‹ÀMIX"M­+ê©~VÌ¤Î“ZüÓQ¸5­¤¼6xÐ}BoÙX0Q#ê¹ËŠ®_ºC-bÕpÓ2™G[üçÒ‹æŽ8¾‘&™´ÆÉ@¦-ˆçÇõ¯º®[ßB=„YQÖðRá*{ÎD9&Ú'ƒâäÝquAÃ‹ŒŒ0Ö¦H¢Š½¤¹7Øyï0£ïÛÙ_È`=Ø·ÚV…ñs14?xèEÿJ->d(šÅ€W…¥ ‰ñKQpF¬‚²w‹‰ŠÄŒl#ûð”‰ä¢áI¼ŽÌ‚0Lw"á¡Š¬“á¯Þ„gä}‡0\QtÏF;,Æ‹YÇŸT,wfO'èq× ‹rårvT#ZK>:kN0@J¡¨-»{v;[«Ë³uÕpÌ¨bÊk GCf:cxjÙôØK=¡à‰æ™¹kGXð©×UÆ 3qÊÉláTÁ"jki´#(Í†‚úe£†ÂlÇ†Âè¾®Íl•%o?Ùò²((k€â£ÎRúÚð¨ty¿^þáeé7{ÓY^ÏÀ@Ç='jí¶ý:W78cúk¶~Ç”ˆHQÕSpÉ¹èšCÑgH%À7ØŒ>t<Þ$»ŒÐªwk~6KÛðêŒT‰iTI8KëÅùK–N­H»¨>ÃVÊv;³5€›šÑT‚†O•Ð*u @lù­¯½:@í››\Y=ÃŒ8K42ÕÞvô…;7aSû·8ˆ´-ÌÓ=;—žGÎë|Bƒ R%ÔË5ïÓZY`
É?ãñ†C_¯¯ußÎä”cjÎ\˜ÒÁA
a€%”ïâ·Æ:Å£0«Ê“<4,á©®ß`‚E;(I‡³´Ñˆ£6–6q'J×3–‚!ïÝrËig%mØ%Ë[@ŠÍ#@ˆ¢ìØÍ…bÎ=Äª}W»zó*}K@Š§Í¡*zgƒF¢ý„6o(ÌMÂBÔ>‹À
])ÍS
„YN”þI
øm#º•ŠìµLéØzÇg‘ã{Ø®wúqp¦·q­E <!üÄÖS—c‘™i-S¿Žë¬{¥kïç—T¹rÝÇtÄZoDZ ÕÂ!(›`¨Íã™gÜ_²s+J–á8ˆ[ u|Õ0uÔ…¼<^‡%ÿ€¡¨6Äñ!WÅ,ÿþ?þ×üÛ?.°V.Ì†Eîk f½(bh‘î8Æ™aåqã
cØIfš~“UÎ„´CÉøÈ’, áF^ûŽ÷î"CÎmêN£‚Ò˜Á.cÚHã-ã#-}W—Å)Õx©ãn¼‘K’ˆªƒ?pÙ±,F€„8í×<ÀŠü‡joÅWsŒ´²l 5K¸Ée®;iëÍ&½ŒUÌôê”^/Ò†»z2Â•ëørELöã
“ß¦yuFaÂ9äÍ"Gq,QcdÛßV-¯fE?Ëî%|ëÅÎ¬"C+g—„ÕÇ´ÍÞ{cô)r!/¶k+Q%ÑþvFÒÞ¸ñÝÁ’—†Ølê|¢³	ç ¸»<\Þ{Es£-N9m´R½F)blžWn®ÑåIÚ…á7qGVjÈ°€wŽiÒÄÁ=2µšˆ}Ã]dÎ	Œ%.Ü<.=ÉËá²PÉ¨]ÏÇ'¤n‡áM'•ŠYaÒ“ge–6mžçjÑ™_iÉ?žÎÀ»*Jd•‹³É¦¾Z‹ÈÐDiÊ¿œˆ¥å©þ²Ôžá¶^j‘\ƒò/µÔ	Ú1$ee@Ç×züX#²Õi`2‰€°®ïhäþŒ6›¦gOxÃ¢æ@Ü›ÞNÀdÞ.H6Ñ>íŽ1åkƒµ­Û\Àž«srB§><I ð°3ãÒxÆiT“Ï
S3.Û|lqþøLÄÞÙX¿ãD¶³üøÞf©†èÅ€ÐÒŽ‚ðMN!zNh0(~´IZÒæ•tÀJj)íüäí±[‰œHÚÑŸõIÓ2me„>0D’²ô~+Ô¤ $1Å¤HÊ¥oaýÉ .N÷Ã+Ã\íw~:æãwa<¬‹á~g\U“b\@¬cÒUQ×~3FV’ô)NFNËŽ¤eÝ]¢™xÄMb‘ƒF0wo8Û_£}?¬Ž‹˜ç˜ö=Ý]°L)<:o/ŒÉœKÅ/Úâ¯ÁbQ¯BèåËfdb—ÌØ¹\9Œó!Í®ÚÇèRÝ’…Mœ"fiÆe–aY5þ„lXæ;Î®‰½µ‘@}—ñ,HþÅÑró|ùÐÓ‡ÀáûYàq?¯ýY“ŒÏ)"|O XgG4@I7ÂfnÞõ÷¼pì0¯­µsÝ_³Ý¸£¿¹zàÑãM_ -Ë¥@ËÒp©ßT.pGgùñ°@Š§Rx€”µ„°cç¹½Ù Èû(Àfµ1)4×”4ây½¹Ùòü¡Xêú¶O´!£4Í1o3Ÿ©dI(5…§Õ
ó«ö%­3>õÓ½õÙ uZ€.[×¢å6k@OuOš¹GV$oÚ†ÌYÝ¬ÚãêR‘ç|§W…Ã¿G§Ž ÿ
lo„<¯d^÷`óÞì¸ê_ê’Õ"è·v™ñ?l›4\ÿÝÈ¿Ê²|U[mü¡|„jï6ˆº<ÒÑ%ûê8* "ŽëuT„¢A–H“Z‚H %Kõ Ù›õQ”y’ÊXµ}.†Ü¹Ð—Q&¦
»²’ê|ÈJ#DVš8"²ÒÌ‘•VN‰¢ªÏ51©z–¡Œ1kIV"~Žü£˜I¢Ï#+í<YiåÿÈJ:TÂ¾¬Ä.â!óQ¼#Y‰8eÄôM!íPSU°×L˜%ÒœHž çfif×ü&ÙÜ“£ÞVVfê^1®fL47K›QDÉØÌVvœ¤Tnd"¥SŽiwŠXä„PÎ@Ó¸ÉÅö»^ÇØ	Sd™’§¹Ñ°ÒÚV”ÃfN+Im.ækŽ,Ù÷#±UÃIvÇô‘Õfœä#+JÔ0)uç¤øÎŠ²ðŠšñÎÐ#+ëHjÍ8¦mÞ†[·‡Ûâ	˜±_7:Ìeû»×ÐçMoI~o!¨5ð}£Z†ÿ”eùÀAiå'&´_8(¢•$º´÷‹ƒ’,Ö¤ûÇÑfãÂž[FîNoÒ|ÝRçr’ýÞh‡7$Ü,Û.™|6ñ†ƒ²FÌ<ãR€Âw.áÛˆw]´…¥ZZ/Ëöšhoý«UÒoÒ5ã›pP$[(îU`XPo-ÁÃÄõ,rÈ¢?…€^kì% CŠûJ 7$õ¦;Ï2€ýÜy”Ñr¯¸9èY&Ú‚/`:iYU%Q?eÙ½4â]™O†¶!nvÈJ§‡t<!‚€×Ó!ÙcÑñy¸½€Ïƒá/ÒÔÕ!Å³Æ¹¦FûXK°¬Åôãù¸ßD[g¾Éé#¾—æ†€7fA9ÕÖÚrZxû•n_¿š9dÑúÅF»éŠ»·¿M1@‹ßï‹mY¥„ÐSM‚O-j%ÄŠ&±>l6<ƒ±»ÕØØ™oG·É	W‰’:Û$(ÁV,!`ØGX·ø}[»pm‚Â5»ôš(®×.4Üÿ  ÿÿì}ÛnÜH²àû|Ev¡wPê–XºØnwµdY’=:G¶4’Ü=Ã¦ª(‰vU±šdÉÖhõ‹»X°ÀÁ Ÿûr^õ'ûû	‘7f’™ÉdI%ËnÝV1™÷ŒŒˆŒŒË5—\s÷Ù¬ws7q¶I¿aWq3î,B)àã“ozToÖÑÁÇìkÞ0ÏP*îT4ß:TõMO6Å©(¾‰JüÝÎwäy8	Ú‰Ø¬äç8ú 8Š.éBrŒ^¬deSiXD‹XÓÔ…ÔÑáu,|¦Ôô°—DNÿ¨éâJ~»¿&ƒë@³ò†^Í›ŒØå?æ”z öÌ:YZ»(§¨¹•sâšzh,òXæYJp2Ò¦ŠJa6Gö£q’æu3.åýYyÞ-NYt¯ôÔ78¶˜‘g“¸‘0íg5ÑI|âÝ•ùTkŒ;‡Ï$c¼'“‡¤ªÁa5dòrá¦©ä;ø]ÜáeUó½µ¼xÍÐÉ§÷º	2žSán'¥€ŠÑùÏ~¤Q50ßjçôž¡ucHÅÒ¬àE¸±Ýu„8tÄ4¾úG&[­~Lcê°WëÐCó1x {¿N"“]i`r§Fæá(eƒ!Yã#	9Žz§!»Gè‡„iPV=0ÜhÌæâ¢Y‡Ïâ~y jqÝ DénÞˆÓÞ º@-Â×9!•¡5„Óõ<™mN©Êì4Õ\ jï!!£ý‘]ý-C©!$$ ¹´b‰êŒtšhy×NZµMy%{G“ö¯,9ÒD¸ÜFéÉÕ¿`)¯™rrÂåÒý`&2ã,p†Èî¥W¿åq/äÐç4x—À¡1ÊñÃÂ £b×„ÎJB‰«Ø‰³œ$Ç„¹ê‹£¼â¬ÏÍaÜŒ‰H‚¡eìfµ1¦‰®¢]ÝDº+z£îáCžGfÇCû¨<}õ4¦‹·ÎÝ
B)ŒceÝª~æ}‚Ä•â`ºt˜ä@¤µ¶,g]³÷BÚ¸5Khh9%bá×jµ]lžÜ!é@e€sW‘ì8úbºžŸ	ðI«nÑ á¯Âðw½Î²$Ð T ÇwÃ2¾ˆF§“a(cÔy&ëiYˆÓEø·\´‚LbõrŸë·ÈÍ˜n‡ñ¹Sæ[ÄbÀG5i\e32•U”¤ÿ‰„´Ÿnþ<×´"îÇ¢5.‰gz8M2žªqa d3ê%)+«™“­¸<u³kVA;·Ãb¬àó	ŠpY‹v¡ÍêË"²m‘3Ûâ˜Ë†Ao{„ás=C®¢†~xž•
K™ŸìÎÎXa&¾ÙÄ«ê5Ò†é‰~FÏkáèÜ+˜Ø7¼Äœ0,k-´\ÒëA„ýÍ£Ý£w]‚­º2c¼þ ƒM2êgu6B¼nÁÄÊ5ïÀÉâõ‹Ù›š2Ä—°‡z¾hóÊ\TàkÛss¾3Ìs‰
ò„ZcFØ¿ƒ­\Ú­q¾ðd¿åèå¥ljñqâãeHˆ²ãý¬¯áÊÀh×WmY)W¾‡´°¿¾JJõ¦6^ú­mªãP5RÁ•}¼Cá³ñbïŸ»þf7ž£«½Œ1xÝ’W2 U€“UÉo»êò-ßIƒ£@}óNIÕ§‘Ó÷ÚÚ¨SxÉåŠ³HRÅIIxmuS{^§sï¡çÔ 9ši„à3…îmÉßI@sÍCE]È#/)ní¼rÛÔŽ
•G5úmn}L¯i0-Ï ýTÜ|ÑF­6­ªFW‡6½õÅð9ˆ†Dœ‰ÃŠÇ¨¼¼ÕiòÌj³5s¥Èžz‡ŠSoZ““@_m/ç¨„c@O¿€´ƒž¾ñáþýö¼ÍE ê°âÐlÅ©x´y\ñ4[ôCÔÉŸß0}÷}½¯?VÛ­a‡&haáî"€²æ¤w<âDIgzi?ûëù§@R~–ñ?4dj/ðÄ}&ÞÀD†[}‚w*í¬Î&ÆO¯x$£‘©˜?
ml.æ4ªehr©©Åó·kj5æqàñCÁÞ¶cÍ¬Ç®o?Ö„õd7Vdtãz5a;?~n¢-M«ðŽ%ƒêÓ
\ûXÓxécâÓD']/Óká™îæËqŸŠ¼P†Ë58ç5³v(hè°YE5æðRž%—*¾›SyMW”“…”éîÃQ<DúN^üý8·˜TŸ
ø•Yº¬óÿþ×ÿþŸ²R·l?`šó³;é`ö`gv( «W¡MêÓÜ2¨É‰¹Hû¿ÿã¿Š*—p¦;7˜=¬cøjà&Ý=Ø|[ß2ÀÑúYÌÐ ÝöÉ(I?Kxóò#7½m/ÛØ¾ŸZ|¦°saI¯íM1ØÃag€l#ŽQîáÛÊ‹ßmh>CËx‚Ùˆéí¶øX×Ùdãaµò°Øy˜Õr4e;ý»öªÛ%¼HP»kí&E@§=B¢f¦&	ã49ŽÑã MË´ÞÆ#ò<ÌÐy`Él¶©5I[Túº
kTcAÕQ’'uŸ¬]ÈŸº)Ça¢ý¡[ïÇ= í0U[d0‚ÞBuãcFRXÊì…i/ŠÓ°´Ôu›^ÙÜS‘Õu†å÷ …hy±†qQkY†P9ßÄx /‹ä¯ÈÕº;Çs5™{ïñ‚ùþöÓdŒ+iU­du˜ ¢ ålÆ#˜µp°vq!¢3uÉâ<Éð*~?Þ¯"8NÛ´2K²ÌRµ@ô1ÎµP§yÊ5Lá×CctÔªfñv^ý¸xvúšèÎ)µFêU¬¯§ŒêÒ¼ÖQ/Qu˜ïi}iàŽþ8^ç)-tÏ‡{W$±L*¡íTq¯#æ%â]ÚeÇá ‹Œ˜Õ._4QÃjËF¥¬¿ív3U²Q!sˆ«î$êDÍ$]Eq³¡{q­ÅoÃŸBólLŽ†¸…G}žo$KÐ+Ìî{ÎS#£²e£Q™Z—U½
¹´-N¸D”«ä»Æ,½Fï‡5±’W;´ykçÜªG´¿¨§`çšúÑq8ä?‡ƒ	 ¾M×¨Í¤Ñ¯“8úÖ†hQÄv®è{X4ó—uÂËFÂÉ$§ÎW]ÇgçŠ(ð	ÀÕyf¨"4-¸Õ‹Æ@Ó£Sp”®µƒEöþ_\p9gýaÃ‡¬Jïiá¬êéÈLc«¶Û0[£©ì†Zfg]¶šªð´ð[uUéðNyM¨”ìß€§}êW“1rÈäŒ¹ÕzD-Ñ!%
‚`µÃ¾zWäÒŠß«Ñ´¸Á›Ô›nÞ#)¾|$6®DÚ=Ò$Tu•­v€~ÄCz¶™²NßÀŽÔUOMQ4ÿà_èŽäR¡–:p„ÿ6.Êœ¬?¢~’&£¿F­Gìïµ?ŠXO3d±
ïÀS*ÿÂ»µõ±K¶ú¾íçÉÏa½s4ÿe³\ŸÛ	P‹$6ÇD©Á~\LÅõ¡TqÖ:-˜zù{ýz4¼18ÙBŸ¥3	êÕLxV“É15^W¿‚‚/(rÏÁdFØ‹Çá€´E BˆpX<íú‹òvPbœEµ¨QÀb”s!kDs°¬}tã}E^ê·YCl?ÑÞ-@ëþõI¡ªÏ†?G½)^ïî‰Ê¶¨Oã«›åDøMžšJ¿Ë_L}šõ~BýFÏp±™cêi—š»µþºÐ×½QMYÞÜÐ¦–´Õ8$ÿúË•ìÄ£÷d” ßÜi¯U|@
Vä=61-P‰òžRªÓ<gÝNç$I‚“Ag%;A|YbªO0ÀY„ƒˆúøùê7´­%åÁÕ§í]|há'Ã2_à¸±™¥²™€7ËN	5Tá&õÉ0J©_Ýyºù³el¨*ãäÚxîu÷PìºBúé–]Ýv8½©ÍßÝJåƒW#º9XXZDã±Y_MÉUö¢ÝàÆÊq\e—YŠ«…Œß7SÏpè…áŒc»mDýÔÅ Îò§IÊ–5@™€Z÷ b(Ìà?˜ÄýK'ìÍ­Ž™·]ÖØÎ^ÔÙÀN˜Žrm(ÐIEéYÜ‡y¬ÍÛ’;Ð­ù>äÉKt#¸fQ»Æ0ÒeÀR¿LvÇ î[Á*2²d0€>áØž'ã…¥Î2Y Öt£œÓïxä• Rºý.jîøTuÛkvèdõÝ‹TØ t_ñ5î Ž-<f&Èf°'¢aèò/tV`î(ËwEc®Bkv,dp×«ŒüÖÒ²uŠ^6ÂaHsús1Å¿N"Q—â,ÕéÅÇÙ­â³†×¥íÉÝ¥rË Å)*züÞ@Lž‘Í¥ŒW¿Eqfóå²‹vùÙ°…s@ã¦ƒ¾ò€”µƒë”Ö€¶a×*zn®Y—C¨ñëûcÖlÇ®èDÐdè¦rr‡¾rs¿ø°»°ÞiÔ{”|t#dš+ê¯Õ™‚	FÆ¡FØuÖçÆûDÔø çDÉñÕùŠdƒ’ü•è}[L}0&g‘¶¶^½®¯ß=Þ…Ó=b˜‰‹û ¼á)üïÇ¸áã´JóðÙ ˆ†Ã½HµÍèðÒÍ¸ª9TVQ0wúDßj¡¹S³|d­ºa€:&Âh~<kb7ÒÕ ÷.åòøb=}ƒò–÷<U±7ùÑ¥zp¼°!>M0"åŠÖÛ£*XËÛ¸–Í& F¶åQ SP0îÂØ«8>üÈ2!ß0$Žµøù¯qóæÅÓ%¯‚ (wt^4öÚÇØ„8qÆ<ìˆ=¬ˆo	EâãVlV?OM5XÍa‡k³ÄeõZ¹h‹õ®9vI/Œ§3Sls£ZR\\Ö°uuÖvµ3­v¼Öˆ¬*`!7¥Ó4&;¸¢2¤ç§GêüAÔ‰N¸YæàDü¢—}-õ6áàcz"¶™kwjØÈýØ+æŽ&£F§aŸžTÈª_JY%<­vŒ–¿2; òåþp<Q¡Ù(ÉÙ†U×4¡n3¾!ÖWYçÍ8™§H„ý~˜v’Q×:Ê…©yÔn1Ø¼Ôp2Èãñ€{Î_ìådAè6dÆ1Íú<&m•°\g¼#	N¾Yôt„y›÷4H#*Un·æ[ó¤´ææPó|ñ'µÑ%½Ñk{´ŽN‡ 4ÃAµ‹yŒñ”'Ôá;òüê?xÍpâÛgbTÎ‰È^W´ºœWfYtáµ:+Le)¬é9ú™?Có£	:LVf‡7„îÔ·Gù xA3<¥¥…×òyIU³ü-[=^¯t¼&àÛ“ýž|9°n`Ût©™“sÍª8.ÌŸYiüùPb£ãla:×D«§Ë•ì+¾fÊW£ Ïâà1°×M)V;§ËJcUŒ£U¦ì4$	søÂ£0þ˜0§)=ºÓ¨à&U›%9FÉQ—’Ã{Å=*¬ÏE¤#qzîXÓØc!UÉÉC¹"ºc÷iîcË®ám‚‹ªØhZ!”¸iÛ·‚ŒAƒ{ŽŠlIƒÞóƒÖDZ"dÿ[•1³>VvŸQslØD¦ÅeÃS&þÅxõÂ²û\·èæ¥óæSù6fYÞÆ¬46Ålx#£óê²Tzë°Î'€ß­ìÙR»iÞ	fPr#Aåª€`‚"A,ëÁ¨b4ƒÞ+¹ì·Çw~:F@©$8p0û‡vî¿>ê®äÁUƒ$@áÉ´ÚÁnW È(|§Ù•XD
 ÛâgCâ©t–E£ÛÊr`ûåë3®42&÷—é3Ì}Éu »C.tÆ­-¹Û
èšƒYV™15‡…kcXhÓÕÊÓ«yÔ«ÿV»ÛÇ~Àg|­å[œÑœ—JÛ­ë½â‘jQCË³&šTÂ:˜#‹z;.…8Ñ¦ šLæÕß	G¡ŽF³Äðx4þX’Òp—LúI}•	3`ORøsŸàV£ú?Ù¤¸cb&:²‚j¸÷•úãQ¨Þ3U™LígéŒH]pÑc!ž¸.<Àñ,ÍðªùÐÜaMÃˆìøL Ë‡ C`WâO…æK†›qˆ¿`2¢¯6{œÀXð÷‡Ó0ÏÂñøy”eá	«”ùû…9Œ†cÄoôÐuûWZ^øâ}àîÀøs{ «{¢§ K?€¤õqŒoÃ(ù ðçÁsñ‹]ùáÛ¯';ñ	IO‹Ët]h§»W¯éØFáà<‹³ƒÞi4Œø£‹uíKßŒQnø‚<Øiòa{ô.êåÔÓÑü.»DYÊ.Á‹ô=öBþ*ƒŸÄòj_QY#]í.UH{™Ç–ÈW¾KÚÃ.a'ûy’?†d6éaˆžÔÜŠÒ4I[”ªŸ%qÿ''Püa5QxéRLµ5ÊÓs–
ÀÓ%ÏÂ±’Ä©KöØÃ0{Ï¾H¸ê’þ“×Ma*§7ÂqŽ"öM@]öô—þ]@"´Ç~éŸËÀÙ%ìÖe
Ëf×.« ÁU1u\8ò“Ž±Tàmr˜÷NWé>fi…ÕhnâUsõ¸Ûi•ÝÐ-»§È7Î!ÿùS‚…Ä¶‘ã¤=å½¡¿™æ±ë8Œñu#éGËß~M-É/ÓA%•Í?$%	€Þˆ¥_ŠyÐ7yƒùåÕ«£+P¥}¤u£uŒØ>jûÈÅèñad@ÂDÞ„®LÂiWÓÂ¿@\wšôqË?Û:¤Û}o÷à°õA—”i@Dr‰Åèþ”2Œ³hRik}vÉsöƒÃ¤,ñày‘Æp*K§¿Ù§ÁvÉŸÙo±?{¼ÍÍÔ/`:8ö*aÚ.Ññ¨@qUÜ3“ñ°QÐ_áªžá‘6XDz~²âk¨-î‹	6¬¢ó
á“/ú#¨äR“î–Rx4o@8á‘&¹¥Kß:³<Iã^B×š¢múP5ý{GÙ²rq‘á°ÁÉ™p(ò¨­T¢
¥§P¥3ÆN@½`„ selˆ"wù*ðºLàt£xgè\¾Sú&ß^/ªNR¼æi£F¡ti–å{–J¯'añ:HNŠâÈyÑ›|uTjâÁs™ÜÝcž$ÖG¡9È†¼¿VÛMÚ+ÛHl1îWŸE¸<B=«…r>ñ¶CN¾¶zÈ0îG@²<c+¨¦h«xâ^LéèSœ’W¯	^Û ¡Ìöz¥ÒÃµnîÍQ©ŠaÜ§ô†—]g(	?DzÏ0ZÌçM_Öa¢Î€[§/€ñ'YØðÌŠ0Ñ}§C&GY/à´švà‡AR«cì.º—ì5c¹¶ŽagkW¨™SøÎÅ{f>"¢)¿–°LŠê<–ÍÄ¯0VŒ‡ÍToÛ}Øg»;;[‡Û»/‚ƒÝííÃõõÝ7Owwž­‹ ¸¬ôN(bXkˆGá8;Mrqiþ« wíŒ)Ýå³jPu<xž Ÿô¸ÂübzÀ*¹DÌÐI»’X.@¿í¹âÃåÜ›ªÐ¥„eÅåIà0	K–flIæÃ~äl·ÃyrTÑK`£^´Ã	‹pMeÀÜâö©\ä	9ò)ÂoYX™ÖœBÂÕ0Åå×Æ(7Ø60)ö­R®"åÚKX" ßÚ­-üCøV%0PÏº­yÂ*˜²bdâš‰©[ ØžWhrhWgÌ3éfÔ‹ûo	—);õ8[Ãwc·.óœKe|šeßKŽ†MOä4ÑðÎá$?Ø…ZŽg©bÏª° dx£êS¹ê¶ëEútH)C[)j®ŒúBƒžªÂS®W<g¬…P“ËôX¤P•y™Ü­özý Þ½è‚¶$0Š;]:ÁûÑ1¶’ôêÎ¼((}?„1`µaj)ªW “-W
B¾j{höÑ-Í‡1úkíªsSäbÃF.-JÏàÐ	ûZŽì£ìE~*.´‰Þ
ãBšˆ|{!FšñO@fhdOÓöÛ6î]B bÞcæO9eà¶.ƒ·Ewåá[$ñ¹¼æXEºµy¤òeÅ*›Îîl±UØ¦–+#ÆI³“0qyCV4ƒ±xÎÏÿÅ¡ëRÙ¨¯bd±´´×?žgÞ™µ47}³—cŽƒžGý8ä5è‰Æ*JØZb…
L	>›ÝNl­þéðùÎ6^l"äÁU¨Le­ðïÙãàÕ"WÛ¢(u´€É˜àU
#ùÀƒ|„H°;ãþq«X29û »ä8<K(s£Ø„é¯“ø,!{›O«Ë@´†/%'Qšy•‰Pq&tè]¶á|Ñ]Ñ;~»E?,ô í‰ÌÁ³Ar~IÒ÷QºËN’Áúvö ª·Â°×½Ë‚Þ ™ô§‚^2ì„ïÂA|”átï²Î·²fØ™TwI?±*@LÁð]ö¶„œÂ4ÏŸL€íJeïéÔ+Ú%¼	µÊ¼²MX^ÀJ„ƒö=ÐÕ*¿œÆì8'*ƒ -C]­gƒMÚø)†ô¥ŸàÏê¶Œ&Ã=-AÒ÷ßÏUø©1|S{…ÂüíX!½œÚ$€dFy‘²aöCèÌûÖ®ÂL‡¬·¼|@eòŒgÃŸ•ÀÀ· Î±sÁ»$µ[¤¥TG‡ü=,ï¿ŽÈÞÕßNà¨.¾$ð¯£·äû¢µïIë_‘¢päª/Çæ´Ï´gR@äi¨…}£¾'õeEdÉb°z•ßËíÒ—ÇØ1è¢_ØEÆ·ˆ«i®«ÿè£$ž¬³Ø…1RÐb&hŽU›	le2
T1QÏ'»‡o6v_<Ý~{(ñ¨5g–8•¦¤«L€ÊSû9…?$TöØ­KGªt°HÒ˜—XÅXØÙè±Î$4;Ð.íþM«)Ñ2Q¦·UÊ}ÁÜÁ‰h/ßVh½;
|gÀŒ:ÛZq÷# ®BïðÐ¿—FgpöãZùnP”ã’ùG$œ? šx„eÄ1åHÍd³•©"¦Ùˆ¢ý@jêÀœ:KÊhŠÌJ^1ÁšéG©ŽFå€É7µkà÷1™¯UAÉÙÇSJsµ!}#)n@¯³_âü´ÝŠ‡ Ô tÜh=.2Á¹“\zOG‹›Ð]meÜtö	Jé³í!Cþíc>²®rQ¿~|Sôy¾6ÞM'ƒ³Yo.ZÎ¹pRe‡,‹MíÓ•\ž)À?ëÙ&PÂ—û;ícS–dÄÙ¬v„•AéŸ[…Ox›t¨zs?+UMø„…#TD!”Tó³7gàÚ-–¡Uª_¾þ—7¿loþ	j¸·¸hÏô§­íg:4çBÚþ!î£áí8ýmÊtaXRž‹½üTÖøFÀe•=âùMF[j.9³uoô»5e°bì#>b 2s5SYÕÁjî…½äÍ?RfÑÜOÖ0ï&Ÿî³d—šß£«å¦ÐˆAg…ç‘í™»ÄÁ1ÿˆì+DƒSãvk¹_Aœ(`šHúiøm —yŒdÿÑÎÈ¨˜Ëp‹¦ƒ¢+y"v-GãÑ	`°Åàá}C%s´y5•ÖéÞçt“zr”ßñ›J}Úû¥	PTKÑJÊDW©i¼A‘?£Çw/ÔR—FLæ×æe!à*Dð†äÁ=Êÿ§!xJºdñ2-JÏÍ’‰”T·«vø–øGl7ÇÀÉå“pÿùÆÙñŠHÓCk²F§åuz®r‹–êUPíû­±fdD…S-ØŒ“tÀ~’^HrØâuq›YßVñ®±—B‘C2˜’ë*:ÅU,D¿ø«Þ5:"øÅnJPnF»ÇØU€=¥žïÿEï!	¶úáÂ1†ß/Ëí³;:Ù>{5´/§‡qu#<]lPRÜ“âÃq<4¤²ìÏÃüi<Òò–“dÆõ^Ø/çTÓÊã`Š'bìÍ0Š<Î'¥cxS÷AïÙcè8T7ä½b?Y³‹bûÈûtßá(YG5Î8IýîóîSy.„š˜ñ^»ªHf¶é
Ð,ÙÄŒèeû`÷€âjxËÆƒÆakîÕâë¢ÜÓxØ¬ŸÇÿ|¸‰ƒ’ƒ¦¨G¦Ðõ‰z§ì˜'_ª§±Qî¬îF|×vcU«º!{(‚ÄÃg˜e’´/+Ý0|xÍeZÀ-Çë ®½äÆSµ…Ò­¿ðÈ	Ó­tEž:õd·@wÌsQ)4­aOMq’Ë#âó¯lÀ-5E«JWÓÞ¦6¯[è-=­²
e%³ºŽsµµnž¤u§¹V­é…|5ì:&Âl²YÂ›Æˆøæ8Ïû†T¹‰ßÂð ÆccU–/²:Ë÷Æx‚\ÔºÊI²’ò‡0<Œz€qµt9I–.€YI²giØŸ„”Ü(óbJ/f¦úÕˆS¤¦“
2Q™ZÔ°îzÅ/”­RG«¼r?ÆuG¡®G1ŒõØ»È}ùÈŒê‹*^q“…
«âõ~_E³ÛJ‚ÏÍä[¨zõƒv”[%1E ´aÚ&ó„($†;’ðœ(
vZn1Þ(UdP.%µýÄŒ9ÛJnõ»ÂZ—v\µœ–A)hÜ”Õâ†lsj§Ë{ÐÔq=Þùúò•L¦AÔWcÉªFâÓ(øG½ûŽÅWS‡Ëy´.J„cê"ÿHJô~:Š_ËåªXÎÑYKe„hl%‡¡ïu”²XGQW1£^[qŒ§.”4Ú¬ÊjÊêÊ‰Ùt ¾u¸þfs{}žèU¢F„zÚæh¥HRdòÜŒeÉfŒF@aqvgg|*L†*Š7tØØÂ~æÐ,ƒš›×d
AT†A
U,·¡v¼Š6óxå¥¢ÆµhÚ›â`^eDüš¿¸¸"V¥…¿áÕ:¾ÖqH¬ú2çÃë5%[¹%VS™â5™’­œŸE—$æÑöÍÅE©`ãUäAóù¡W½LèºˆG²+vŒc«TžŽ|.æKá´´\2½È*90-#OubE™ûxñÏÏÞìl?[ßØÝ:¨±èÅ4ÈñÏÏgN\ê$‹ø¸ìÂÊàf"õQ+ŽH½pd1ð2ŒµZF zõÚ @EÞklGm SîGfaØ’%{å²øC<ê'nuÙnmF}Äuž@§"4M£“
S­—ñ6Y½þU71ƒ•>í‚Kí& ˜¯jâí‹¶i‡š‹É+³Ív³M…@ÉŠ‰ž3=Ïé­°¨QV,š¶ebš³—[û0[DÊÖBríÐå«ãYšð—Ù„ÙXDÙ>ºÄOSl¨_eŠ›€³u‚K Ì[+´œÙÑÇÞ`§h	[eü(íŒ¨Oøè£4ì¦ß:ÂÆ	Z+G¨
šGÀ)…:*ý[X]R¿ ,ª	ÉÉE¡¿úMß‚Þì«@ø¤>SjQÙ£J¦ ŒŠ”wa†¨žæ!ª’lníln©ßß~{QÂeçÛô^þÖMªäÍ_1í&õ’K7ÿAJÜë¡Py¬ÙEàMë¸³ýâ_
zŽ5PRNø…˜i.è Œ›MÝn˜ë†ˆ3í£OŸ¤NÛ7åÕä“öÊë2o6ÖŸ	¹g¬î¦0ðd÷¹vàeµ%|M1vq•¦¤{Þ;›Ö*#›êÑØN¬Õ(>`ÏáMw½(§¥¢ùÄ)ßaâ3Å=¦(Öè.“.Â4÷™ø4¹ÓTgÎcŸ¶J¬+ŸßzÛ¹kÅµÑlö­¨½éÎ}º»¿µ±~p(·®¨È¶yŸ*.>Äsí+g¥vÿŠœ7´ƒÅÀ½yÍ(ÙKc×N”CtžÌ¦¼Ceß£ãc¸KÇ§zŸ.2WîÔÕÉ™b¿I°ôßp\Õå÷›/líí¬¿x±µ_>¬(Ú82ÙªG16#z×4}X²høàSÕòaàâÒô¹æŠòþ5YP¦;ô)tk{ws÷ÍÆú5u4­,×sºöÊFéÕo8J×Ò²¶œhÂª@…Y‰ª(VÖšâeLÉ…ª¢D9ý: ÃçÆfžP‡%½òBoV³JŸWb3Èîû
Ÿ±¾y¥ï,|¤G-Ú„ê‹Ö¨±„Ç%ê¢zÊÂÇä-£ã+©ÒÊ5QÉW·?Û¯þù`÷EÀ„ññ9UôgZdyîuÇh<ÑÕmßeÉHzV™P…ã—û;\÷µeQç›×²†…ýP¬^œ¦ÔÂ*)ýäƒ°
xËw¼	G'iø†.Õ›o/üHØe€£x+jíÁàÞLP0ËNN"‹ÕÃ[²ôLƒâ,[=$S‡xÒm¯æXcKØcè@µ÷0ÒÎû0ÕQ˜¢×_ÕuíÅ+…]¹ n
7p>E? $$+˜†þB².yÅ¬¿És˜Ž(ÅÝµsõ´‘®ð;)¢œÑ´”
‘Šàƒmùs®õZåËeo4ÿ(Ež²dÖ+ÒÔ4&ÙÜ§íßÁské'k×P~ulm0 JÝ4DäéÕß?œé²ÌWÃ˜;É=ñçIªôFE„Ÿ`u„ÿuîdÒ'Ÿ3á;¨è¿noª¸¿ RWÜp;»žpWE¥ÞÈÔ[î‘âÓ©è‘v9rû}ž­Š	¿ËF/â>A§¸{-¥S‚9íZL ]{]ÃªñÎBøn¼ýN©þÄŠ~©èxñö»Åœ¢BvF®þ{Å·Ljêi#u¦¦tUØÄí ñÒ'¡9Ò›Fv uö\Ô4sXñ+Wt{£øDÄ·ON6¹W>…Õ`ÜxDPä†£«³ã>­Ê×"XŽ§½orêúãâ?‹ tºÃ<ê`bnŽ¹ÃÈÃ£J(cS%Y2q):Z2ÚÀS†„¦ºAl³ühšˆ^½\qJ¸xk‰^u©N=¹³ÓÃ(ù†ã²—úo/´&„/QôF‡®.XwJYÐ}TËÈ‰;…öë£1uIéD‰a€Ê OuîúõZ/ßª¦û9§+A!HÍS×4§Ä[R_èSœ›Èÿ¨€ÂjÆ$D¾ñ\`µšÇnq;ªÖDÏ”ï&Yü…£(ÿE#-xa9`ÀéJÕUºg|è¾ï¡}¥<JêõwNWJ-ãU7ŠªóÏÜ'TÌ¸ÍqÔXc¦­sí€crž]A!³R(ˆJP ª3˜d®¸×,Þ‹ä,‘ÓkŠq`ŠDVuö_‚£’à¡°‡tü<Ä†½©ÈO£°oˆT“§¥¥’[]ÄdÜ Ní€¶Q>ÂNÅ¿×z„žØV;ùi³R[è¯y±½§ÍF0§ÉhŠ^rŠ:E?Ãô$i^Œy[m^n§Ì!5­ÆÙ1ÔjŽž¨ÕF °&Î	ÿaŽ‚Ïó²È"ÚVh{ÑFjHdÑ+£9Z¤B¶{#[ôÑV•ä)h§—’4kÍ[
O(ê#-Â½”&VÑ¹¿h)cŒÒd‰
Ÿ÷-‹géQ¢iÄŠ+ŽHÞåZ>,<$§ôæ¤B¦Ê†b¤j£‚þñ×1û˜9£ŠÃúâ-è†ézÞ^´u†¥ŸËQq¼©ª¡?®^¸‚£:úÌ€®é H¶ÓA¯ÔwêTÔÜù[i¿7>¦<ý‚!‚ç­ub|Š>o¶öMÉ‚[á…Çî4ÆãÎÐ#¾Ê¡¼°ZVnS\©kUTÞnÉéšâãˆnˆ!$P±#9'MÃ{â¡yâ¡V¶X
ÁFWd‹wìÀ,æ_Œ°(Ñèt2W;ìCãàìVZVn‘GA¿”p0ºtVÆ®zk{<çˆ­Ì`õ¦ðÓõ÷Ê >áVÀæËû %XÃ~2ø,öDóÍpÁ.Ùp³öþîÎÖ—À¤Íà:•`ÖAuú 5Ð}•uð¢ø`KËØÒ’ÎWÑU}µ´8þøÚt"³²¨øH6Õ¹VòxôíŠØÃJÑ0JÃA¿(,°¼½{Ö•s!]ÉwC§ŸÀÀÔƒ7:”Ïã³ÄJ?]ÜÙMÂ‚7'nŠ1(k±ÆˆOINãÆf_6pðú©®\ÙKðLã,¤¹•ÑuCMÏ¥h7!ÓyNåÅŽ&Èm38)	\@œÇù *g¡çSr ô?´çw!³Õ‰Î÷±B”ôÀ JR²[âÖ+9š‚AÉ.jâ83YfV†H.ŽÕ@’ÓL/v‹i‹ÇŽüÎéÅ:–?ÕÜ:aÞÊdŽ­ºä›s9‹Ð€” 1Q„­ÕsÑ_ªAé††îµ=Y$ºI%"¥°¢²‘äì°Êüétc3ÊŽéH)åx"^ìÕHŠÊÝË²úrDCJ{1Úòuë*ÜiRKxqãÈ¨l<É6Ec”Äk
Pð”ëá#nò'ÄùS¢¤Ã4ÌNo'M!!2‰kM,1äCamù¶‚Þ-¸£`_(žöôk7š¿ÌíÄaÄ# Cp‘ÄãÇE”‹¼Uî§Éå‡)
nþŠ«X+6Ø­Õa‚‹@û•áÇ£×..H¼VœŸSc£¬â™n1øñ¾‰ãGS8ÒJ-ÉRK¦"¦ûF~;3?.Ð+YÃý#Æa._4–ê6ægzûÈÛ¸æ${4Jo„ÙòM$Owà7›CÌº‹€–›¿@Ô)¦	!X6ç_ÌÁ¯Õ!Ûv¾eÏ¯¢Š‹¡¢dt09Æ¹{Xø~›™EùáAu ­M¦zO½Ðó×v$‚RR!ˆ³nÐÆ¢´ŒÎ5g)ú3¢-¢MôÒÝnaZk®pm#öT|[.L}JS	x¹4Mô)Ý—ËB’OITï
%Ÿr%úW¯žœ†gÑ^ü±ÒžîSGTù$‡å:DºOGIÎ<|•+‘|j‘Bg½žìS“ÕéÅ1­¾ì¥¤Ñ„³ è"~'•‡{û»Ïww÷ßì¿\7ûÒ§Ã*Ê£² ~g¨wR|P;jî‘¹UN
­fç¢ÆÝk<x—ˆŒ0¢¸WEˆFTgça¨&F4)+\w–´†GKVY	j ²Û¸ÌÄÂ;Ñ†-Š1Þ˜U­1äd½èW`"­ç¤>C¿?3I¨º:®›Äª®‡Q}Æ¢žeÖz…<Yí¶“*zÌdš·!Ãóq]–©Úqíú{^é½§¤½;¦†=³8äºKŽÔ¹ñr‰[j«$Àûæ©]‰\…È?Ìaä¥Os6q­ã äS˜ò(]k-..üÿ»´þ÷i”³íG¤à±fx%ŽÎ‰t<7‡R½7Àþ~W^¨ò¡ü(ód6ËÍÎ×›–ÙÃÚ¯‹gf	:%,Ö^\œ#‹‹‰ÙÐØgnýl¯“ösX÷xa3:‹{Ñl ¯8g5†>YôÓ!lÁyöþý¥¥ÅCÚ	s‰7ºúa”¢ñ±?XÀŽ+Ò~‰^’sKŠîø`Æ®þFÑ$'Â¤Ð°åd’Åy’q˜†6‹ãB/#iˆþ0RgX€d|§Ážë?‘öŒšèÖE¤Cžnþ\Ø"M³œ
ClYAã]à¥xw3{€	±¹—Ö;²zÚ‰:Â“:-‹ª}«Ñë)i«Ñ"ŠÊ{w_ŽñLö&jzoÓðqé÷Ü™íA•ûgöTÆÕæ]Šu³D÷VØ4éœYaEXr}4—Útš–kŽÜ„¹Š< ìÅ£wo.o·ÐMPÜOÒ7lA¢~keÎšZ=i‘éR{A1£Ò=»,‘=ÜÕÉ+g&h\·­Ë]c/ZWÑ{£7Iî¤“Ð£úxŽÐ¯NR“÷uaFZŽH©?Öà™êÃ§ú›Ú¹¾Ã³ç?#Öo=•ÏT)ô³ ;ìZ€„¤ ²@ÖYld+‰
3b§äµDcÚ"o:>Ù¡Â›bjþèž"ÃHãèa«žB£µZâä ]fëæ„ýIùÂ	€àÍÆîó­ýíõ˜ÚÆ¥¸5ü¸Öl‹Cj•'6ºx¯Ùî˜^‰6ýÍ¤æ.5¹/là].ÙÛþËÌ¥äâÒ¸ñf%ï€ì`cïé<»<š/Ä0É„Ïâú 
ó«¤±E­÷î‹•ðFž´ÿ	±®ZN¶7çfR! 1|ˆ’wFÂdš¿‡â“ògTiÇÜVul>ÎVnÚ}‚¿’ÓóöÉ´ä™¹qŽìŠN(˜®¨ñu
=¾²rb)iîRSK4í%ü£VGÛl!†ìˆÄŽKÕî†úÑNV&B,ËPúÅ
Ã´—d°âèôŸPÞ68ò¯¢Âöî'£IžÆÿt‚éœ0+gKÍ¦6—ŸJK“6îÐÔÄg:mM|¦ÐØÄgvZ›øXíÍjoEmÑi‰äPF6èz~˜Î2^ÛKNNï1ÑmÐ‹X½
6>]t²ÔàÅTÖ%^4ë›ò¯Õ[ Õeíº!h<o7q!àŽ’³Î8z¼ûÖˆØoü¶f~ÞnÎb”“h">ÆÌbSÆÕ7ášFã=ûä µ:³ªj¯ÑTÍ¡ã‹·ž/uÛk5áN?,Í8 ¨V˜Í“Ë  ª{¯|3à¡˜÷Àîð@§q.Û”U'tw¾#»ì·Ô%›q
§'"V‡°›jò]Ç)²±(4O¡Ú,É{6ú«:OsûD|”È^öç›QÉZ²î´ŽO¹L0ˆF'ù)Y%jÊÖIŒñþ´k3¢ä˜c¢>L(¡4o5‡Ñ(ÉÈÒäÑƒÔ(j•´:æ¾>kíÜ-¸Ÿý¹4ožâ¡ÞÌ¨¾¬‚^˜O³ººõXj¶‡±–0EcøÚ,˜Áq”÷N½– ŽcÀ¸ÃÆ˜þ®×t×wŸa”Ÿ&èáqo÷àÐ«^BÐùS”f]Ï&`¨o[r#¶µììëžK¿lhÙÔ%%Ÿú¾}¥qüŒ¨=+Q¶p·²Ÿ}+7M]Ïü–+ïÒ°NÕË-ßÎ"òÁsýæ»Ã:‹Òâ	‘m1Sý°ã	s>ù<@gî§:tQÞÏ‡ÀYÈ=-ä7ê©†³ …€FÕçõÃ6„×}¢Ðì]µÕÞztK†ü ¥"úP63p0¢rôFûÓzv´@€ç¢ø,êw=ÁG›O	ñË0Ï)0Êa*Õ6¼:õX+d“#††ÚpŒ]Z^œàmØv	³‹åÏá_á<ãÓñü4M>P›­JMûóvWb$«I:J&ô(BÊHñèìêoƒ¸’öŸŸï@—¾½`³sAk. ÛY–ÐÈ¦ •ô’4è=J†1vúÔiœ±ÍŽs÷Ÿ‰+¬PÈàD‚¤š¤‘M£½½¦ÑäI.ª¦ñSÑåpŠCO¡¤ œpúzý,œqeX^j ½Ðˆ!˜M„2£Íd
‡gI:O&y IÆ¬S­-¦¥µ%Qûê1ªEÂ£0þ˜Ì“_Ã»ú;3û"#:D¯CÈÐ ­l lèÃ#à¥à?1‚(ûæír-÷Ã˜X†2 J{87~›½	 Æç‘ó¤É¡ j¡]¼æ	ë½ÖðÂÎc§õÐ*-éÇ¾ ‘zõU‹%1£ÕªA£ý®ã¨]q|½G¨ÄSõèS`+¢is™=¿TÃhßêiå’ tg0ð¡àæÉqÀTßŽó»Ó§Ž&/qøOÍýíX]ïª›€êSh¬ PHGÕó8&Ê»|ùAqQ¬\ˆˆ,÷iì0ô‰–Fƒð#ðµîþÓi‚@*Í¢a\ÈB±å‡5’Pñ˜á²¾v›­£Gi¾§>/é«péU™^¦PjÅ”0ÎÏó$¼óÖœÏªÒî•Vv˜£S³¼ˆ£•«ëjv\?ì’œ«úÑ!/?ëŒÐ¿ˆŠ  Ùcqh'|æ´Ú[”7"ÚQ²¾·}Ñ"ùv©a`¶’ÑÉ£gIr2ˆÈ6_FrÈÖqµÃ¿7©˜¿^œ…4ê<årú(¸†×	²Xï€‰$¬9Ïú6É¤Ï"à‘†qC½*[Øä()¹Kä8è>ËÌŠôåÇ-M.Ýð“lØ¥¿‘‡ß…÷yŽ–‹äóz—°Ü÷€U›%AõÁx}k­Ó<gÝNGÐP˜7HÆèlû„.Þ¢˜)£²¦Ò>çyð3Í‡Âó³8úð˜.t/_;ÁûËÞ~x°-Õ¹v*f‰´Özs4GïýË_kF^Ë|¤ÚnóR}”¥‹Gƒx-Ô^fŠU\Bm„ê“ˆŠz1o¸s/®êìäºfŽ—Æ_ûŽÆlYÐ@„îCœa²'„:¤]Zn¿C,mÞé»š;ü¼7D?¢!0q¢ú÷–Wj]žÏ—±/NÒ(éƒ%}Y;Cêj7(Ý6CÈ'ßJ5Nógt:vó¬g®FˆáÖiày®£†çE™_à¥=•rÔŽØ¡|§ds©áÓ%ÃN×ï‰Å£iÈmÆ'¸I5µ„öðê·Q<L´@‹%¦úpàòmH=ÏÙÀ··xL~jK¶Ï×…8U	60T|v/«/”•@HL&9ÅÇ#‡e¹xœÝZ2O­gG|êU‹§Ç…£AÔ_»ˆ+âº‰n¢Yšø‰(%ú.K¯¨‰=Š®Øs¬&ùõëUãŠ3/=Nbê„Ï9u€o˜ÚZ÷£ø0òÇ´°Fý$?Sï‹”ÖîEQš9J.÷}“kx˜*¥pI»/¹\­m±†¢Õ»L-Ô^]YJÇKYáƒ+ª]ë¹Ñ^=›žàá´,~Q¢3yPæj¬8;Äõž?Šh,fû)ãˆ©O2©Y‹úEŸÍÔÔÅ²B™{?U¤÷ˆâ¹DcX²¤1ìõ¡ALº¬ÈI—IIò7Öà¦x|uðš4!÷ ãŸÒ‡©ðÌZÀ[ç@¹P{;ŽôUOcñÀ÷$ù¸ÖZ$‹dùüWW¥lïÑ¼†yÐÜÀcÍõÐýê8ÌOë/h»;À<ôÂ1šˆÃ²×ójE©w@™û%îç§k@ŠëY¨þZëùÒ2ùñly¸Hî‹KÃ…Á+áei%xxÿAXð{8ËÁýÅeàÇ<ø,?¬,/¬ì@øAîõ‚~€o+++À<øñÿ½Ü{ ewV‚•{déÍFè§`éÇe²Â*"+½6O@ëÎ²QRQ §—ªù€¡8å’«5RK¼ÙF•û³$cQ¯—D¹Ž“”ù¼Àûtêv$›'a½4XÈ•kfÎp·/Ä²%qmÝûŠ©¯þŽ-L`®#téÁÀQ1\l—<øš³½‘ÀÉ¨X
Î’ÞÕ¿“q‚“cœž0%	 ¦ÉGÔDpV§«kSg1À¢)*ÝxYŽV¶l]—«ö`-@!¢«Kàü‹N{‰34áªó(×8Jhœv¦×0õt9I¡ýh¯¨/w™_¡](ÊÊÅe:)ˆ›3Î¡Î?}õ aÍŒ~¦°œêÑÕß‡F¸K :µýú…j(¤€
BÓ77•±€Ïq“5Y6× ¤í W÷Ð÷ÐãþÃžûPÍ£ÔC¡òí&Àè»Ðc!%*¬5Ä mËWný¸V©Á-ñ¨×ãðSAdº0¨¾è%Ý°tûÜ‡“üÔG=Ã>7¢_ÞDK~Ë²¸tùa¨úL~ŠòþÚLC¹–&ÓTú>Ó(ìð½Â6Ê§VØq~w}õ\U’cÝWÃèþP¹LD!”K¶U’diô‹Kœ|åL.ÂæV‚Ä·üZáNl¦ÖV{ak±Òò_ÎµK–Ä±ßþk„ýùá¤Z½ÿ),OPzN¶¸ä*¿³'†€·é„Ûå×…qUA}P”k(Ç@qUÀ\XC^ˆ µ®*<‚R<òä€Y20cŸ–ÅbÆÅ*v:d# úzß ÌEÖåƒ¨ :Ó¡x8”’œ%1^
œŒðs2ÉÙÙÚR7=õnP›ƒÜíäñª=Úœ¢_âü”2ªë÷êbbdo×k˜ÖZ¾´µ´¼rïþWhTÓÂÁì÷˜ùT[Ÿ _]ñQ(1Ï_Ð@<v¶n\Z{Áá8úð–k:€æsV+aqC»ï4ˆÛé±í³æLí²Ð ?mg„¾ìu”Bý8@ í¨¶ÍÔŽÚÇªÐßó¥ÓœÙËðz!œn&ø‘þ¨]ÿ¨vd£vm£Úºd¼PJñíù†Àú0=‰Âò™3G_=ïR ˆÒC ¶,‡ã¶Ã¤’ašf¥,a•¸ý‘=šrß‚ü¸ž
âPƒ+‰Ÿ¶Ö­È‚Fñš‘L=SêËÚ&°’|3éÙ©F>öæÉÆîÎÎÖÆáöî‹ƒàåÁÖþÁ<GfŽ©W†iËd·a­=”·€Q‘[ð[õqØOá¬‡7½²Ò	3kËëp ¤ŸÒýã»SðÈ4 ƒÐDá€Ò½(PÉî$okÛÒj‰‡nòÚ¢|ýgáHKFÚ8Èìç8ú0×m¹Hp/At2üåØåL"Â²ò®ÏáøÉòÒŸÍéz-Ô(ý¨qƒÜÂtfÐyfÔÆ5ìP"¼`¿õÏâÑ0sNÉõcÒ¢Nøãýîêohâ‰ÿñÖÇY´[œÈS)»n qª‘3ÙÀÊú5ÚÚl£­ý~#Ý°;³63Â™ÈPûý®Ý§„ö5œYå¹Ûàòù†3ûý®Ù'Šò5šÇ×hw'Å×h•_£yÐJ¿lOþw=šÇŠÆñ5üF1–¯á7´%Ý Tb&øg ©'¡Íÿ¸Ûht6{þ&¢"ÖZº7¿glÄ[
oèô|#!¯ä\?Ì¡O Ãël÷ÆqkS}øQX¦m·A´GÜåFfÄ7DÙÅzÍ6ÈÉ4¡”íQsïÉ8H3¤ºH³^¿Ïc%jâÍp}êBMæŸ»µ´¿ïCôŽÎ©X9ËðBÅ·È¬B¥…á‹°‡¦Ù‡áîåáÁó–®6¼úœ¦â50aÖ.ØßKqÁºvÁ(Ê”²¦F6P¡;35Ä¾ÈÆØ+kýž¾ÑÍ8,·h±¶z¨!ŽU>Fõa¼å4¨#iØc¨êË4ïAÞáÇ…p’'%,eôÏæŸ¨®5ãúºÁ"üÙs6w Æ»lÈP¾ Å Õ£ÇüdÓž®‹"À˜Ê÷S¦xl,Yæ=L*ÈÆyÂƒÓ²a,v„SgœUÒd6Ëæ˜8¯“­ã$E'ææç>d<»M´Ç´‡Aß¡ÓZÖÉ^Š¢©øó‡a‘`-G§µTð<ï{”Ü¨—é~R.žŠ•`·¢a<.w]OtwßPÁ³†b†jäP<+ÃáPÏBrüÍ= µô¼®LÑeµ ìk]qìäáÕß{£¸—ð^Fô5öÝ]Õ
B_½JÖŠË{UB!üêÙÂ³4ìO˜áµ€ó$c‰½ÐgÆZè½«Q`ßTW±üj¼œ³º#Š/oô¢A[Å!ól,o ‰¼¡'fØuŠ8@«¢!z·ïSÆX²>"±Ì4)®ðT¯v†»j„ØwÉ' •‘'ÜL>ŒI(½èÜCÂNØT„Uî©zfðuuyíÁë¦j†¹0ýå¸:ðí¡eàÎs#g¨hiæÃ^/çk­àã û8OðO-“mX•K8ÕsBÕQ×ˆBà{ö8xµøÚ¥dûf›«	›Óùzš&C¶“°Ì<áF7¸­j.e1 	«#B#€EûµÔ1ªÿ	Ó•$9&¬z×½]½ý1(Ô÷ª…Ì@ëõõg&ÃñpÞ#éqS :×>e5”†ê.öT£K^LÐ=M›–/1$¯çœU<ë¡WXæT<k4p#z½F>Æ³ò2ŸR‚Sá?®ŠM\OÓépÕoå‰¼§…ó Õù‹ÿD«ÒØ¦¦C7ÖXeª¼ËÙ—ê`%·ã?Zc]:çÕt¸Æ*,™÷xU®Æ°éËQƒ­o­ØÌ·5FÖúg#ÓÙœàcµ!aSôû}´¾ FaÀ$[‹ç[‡ëo6·×çæÂ¯q Ú÷ß»õN
«:MÝl L©¿á¸ê`Zû’OZ°¾LLñ!Á2fß´æoÖnØÍ*÷$­ÖOhµ¸ÐÌ]YmƒTÔ"'6ÊCKbs`
»üÂK>>L´V\þ3˜ûY‘ÑŒ@»c Ã¨=äÉöÁ®0°²ñ ÎÛ­ÃÖðv¶½¡sÖíWb¬ùŒl€5w•®»:à×J|‘ÄÊÕƒúLUâçhSRW›õ™ªÈÑ¦Ž›ëí—Ó‚õ­GðiŽÏF'JâƒEÕ2?úy7Pp€^Iú±9Rˆí[ºV7¹E¦	÷pLë¢õ~ŸïìK“äÛ Gvëª,‘a¿ë£ºb»ä«Ö«GëÂã°ÏµŸIþK[ÚÐ)çìò:)AËz×Å\öâQë;»—VÇ=aíÙÝ®tî4:)”óõ¥²d3•õT&(AQý<§ %"jQ³ªuõQº­\ñ¹­läxÜu·i¤àf/¦KÕ#‰Ñ/NH£‰3_ušµ”›®÷èWqr¬\sÏÆÆ}„hNz6¶Þ¡bTK§sÛÆ•~iè~§a¹K?8‚L]p&_}•±˜#ß[YMSIn7-ÍIoÓbªtÜ®ihÅ6žz†ªO-v«Û
ö`•ð
o¢þóáæl!Jâø ¥± Í·Â<6.>L•îmn¨>!ÔÈÃýlaG¹3i>†“I³Õ´œ*¦¬d*€2ÞäÝX“ÍÁ>.ª†/¦3êûè¼‹6)R®i:P.²=Oúá€i3×” ·8PB¸rgˆ¥EcÉ˜¥–J˜Œ¹`ÛB6ÝóÉÕ€Üè2”1Ä·Ï[I&ì9w¥Jó—¤È‡ŠäoµsèÑŽœG¼Æ¹¬-×h>…¼Ùs"]ÙK3úñI"f¿ÕÎ «9u®LÆ9³h4YÁ8'‹*ƒrm‰¨›ÏTq%^3U®ÞÈ©re2N•½@£©RåÎùªˆ¼=fl<IÇ¹9ù[=‚«ëRâêrš‘œ»Teú^s,Ü]Ã‰U“WRÂ¬'h¹ ~Öº;¹§…%±©ÿ)‡ÅŽ)°¢ÙÏç=“‹w«Ä–…#6«">N
.¦‘®vNïMaÊT­ÌÆ²É›‰¬Ó
Ç»­öú(!ëÈÎÆI:G¾³uèZªàLÔ3¢L–]›»6ì]EèóŠ-c¢6 Úä˜6Š<¯]W/ÞV&~·ìÑGÎ¬r ¯åå]ÉbÅUÞaä7ks–›	ƒw‹Æ7¶iàÀNÚBu÷3Ý0@>ó#øºe”çNo™Ba„º(ág±i@TÙO’[síª) ïÎl79À¯›NyšZ‚™CTc+ÙÔÜ¶bn;1Ï÷ô2ŠN´2ÇÞÑWSµÜO&Í&[«7Ø¡&;á ã¯‹Øfÿ'²¨ we^/#›ØÏ|‡]5qSÕ”T§uˆÁj…£ïRô,§uV3Ë¬rŒ·]ÖMt0lD50Øú¬ÑÕ?R”ÍòuÊ`¡Ðú0,»1TÕÅ“±O‡ h“¡¡NP9«K.„Ô ä§U›®q"J¥I—Kd4ƒü9•êÎgÓË ¨¨eŽ÷
«ùiöh,OKp&/~í `f”3º>Ã^‘d¡‹%áC‹*%¬vòSßZˆú;¢Fq#Œþ¾¿ØY¶_t<	^*âr88Ãé¹CÊáØ\ ]§‹µ—w·Ý!.³vuI%ßN·¸ ØÕ­­õÍÛé‹Ï¥™.Ï½~q‰¦šJbÓ™wLÞ·÷‹]zÎ¸3­GëŒ=°U é©‰©³àÏÕü(éŸ«v$·pNøÒÔÉmJt†|/ZcÇ¯šº,>ÙÏ@@.€·Y»øÁ‹A/<¬Â[ãñ–$î5w.Ä‹Ieh…£Tž †ƒô×õÇHÙl‹é
ìý
ŽMC÷Y.È’4o·ÃyrDeïGT+- î,#ô¼¦Q›©ªY]Ñ0ù=òN5\Ì*Oòõu§;M¬JSrÜƒËÜ%$îQ‚ß†9³¶y^íNÝ«DÅìÔ¡‡/'1EíÌh—Ïõ­Ü±zñŸœ²"‹_‘©¦Qhýü®Ûë‡l¹WnPÐÆÌšM
šfÏZºÓ!O'#vËråì"‘áCÄ»<Îàk!ö:Ë¬z^ø°šh¤ùZÏicÿº„É³æÙ="{©õ·KCµàê>Âbî`²ÜÁj«Ì»”‚jÝçÒöVi{zi’EM+Ô*2ï
.ëÇ«vú“Ë]!ÞùM)¡¸ïŽ?ªš—ÃùuV}ŒÇ‹Ç%þ£ä¼ÌÜÓa"©¸€ÉbFÉ‡4;ƒdãs!M_j2ò}Œ$-™»ô¿Z‡ºhR³C	%6Ã-kZã|áÉ~ËêVÝJåeŽêä”ØP ¬Ì¬B'ØÚò¶´J“Í”×=p·ªŸ£ÇÃè¹ èö
v»3ee%ë
uŒ‚nRßÉ(‘a±½«ÓQ;È¢±‹ÞÈ'Ø²F|
LI®íø9ˆ1Ôà·”>Ö˜™Ò,u á€ÓNœkÒn\¯‹Ü>	:»«ût¶hë®ŽúNc§iËg€?~gÛÿó Õ³@ŸÇÈï ¸ñ™°¿kŽ€KqîêŽ˜.¸ãC¾ƒH Æ»£¡h½GIõù$ÛþÁà3Úïªö®î€YlúÏaÜwpçûxI5”÷tÒª>ŸÜ4T|Fˆ r¥rWwÅL$ŸËàï JðöXl¨¤‰Ãcõù4b‚Y ÉMÜ„Èû©åEÅÃÓµQÇÕT¸üö9S-øB6µ¾c¹ZÀ<WŸ°´ßtaw¾›ß|} yé¸{€’+2Œþx9>,?UGˆT¯ÈíDS)ìmQR<ÌM–¼öÃÞeßˆ•ûßjJ¾§“ãcr®èYôº¢n“¯F£„µÁ¤ë›J††Íéã”4]'?’º°©Á´è54“*àSuAyÉ†æ¦Òpžn0.½š)ŽKÚàDmS.ï4m—gÚNMð‚Za³¡LËWâcöÙ9-*™®vŸÀúÓéƒ^š ýHB£(’l˜$ùéàÜ«ø‡xÔO>­â0ñ&y2v8=-?GÑixSÇ¬wµšEìñ›£{Ôò£ñKh.·Ü×6Jöv÷ÕøåÀlj=-,-Þ/[Úqµp—R{ê9BXìªz–Äh¤’³9óÂãS4`aPµÐ28üõ’Ñqœ}KÒÚÚó.$ÑÇÞ`§d¥Ãp„ÌÜF”2´­³(¡?ö?W¼h-—ç¼øÄØPæs½¢<B·ëÞ	ú^o’
ùÎ	{L>ß›ÕÀ•1›ò8ôË¬~èáÍ½AGZsSnß€ yõ[?	¼¬AgU7ø]ŽÎýAoºqn¥0F \bŸ‰-ÕdtPuÅ?“	ñ£#·@E¤Ê´¤"4åNP¶x7FFŽ+Š,îC¿Í”‡=ŽH	†yT†\U;oj|k4êUJáHKyìÑI=Ù‰OÂ^U‡ZÂxê^lo6ŒçíítØ€WÝ˜¬™+ÞWO4_x3‡±îOb˜.¬˜RÓ€âÜ™úòâ¥Ù¶Y„%\ü™­ihrÛjˆ.úçgóÛ8â‹êe}"ŒV:9“Øÿ  ÿÿì}ËrG–è¾¿"…Ñø‚n|èÑnØ2")™mQdTwOð*¤"P ËPpUA¢†æjówç˜EGO„WwÓËË»çä£*³*_€%+£Û"€|çÉóÊó(TRù8îìEU…šÄãvÃH*O‚IÐ!'Æçq$Qlîg‘”
cs m’…<7_¡“Gø-ÀmNºØ¤]`qõtQ»ýÓ.xæU7ÄLl™ž,êæ°@–œÕBYòiãMYBžéJòxBù7«ü6ÞÂ\	rfPßt	¥õxºd€”˜¦±6ës½ü
\w¢ Å1üXpuŒ þ§…o+ÆüÔX4­YiôÂ38à 5£$0bôÆq˜\ÿÝYëÏ³ É<jEw­l¶»Òõ/§ÀM™+ìÄc8S¿u†è­¬012½ÔCúðòÞ·e ©Æ|ÑfPÐe‘-å†-ù?n¬×ð~4s÷ŽpöÏÃþ§ñ…M¢uÂA	]æŒC;š€ä3Sz66Ï;I,+¨áÊQ%Ÿ„[÷eŒŠRÆÒÎ2ÏTj\lË*Vs&Uu)_¯H8¤u»m€ƒa„¡Ä|4Í=þ¹ÝÀj‹S)åÜHË¯V­„„ø®ŠJBª 4W0cˆ£‡Eo¤!h°ÎðM’ƒÅŽÿìÖÖ@”z£ˆ:1ÿn	åÿ.NPÑ¨“FbV£ñb)½Š[Ç%²åò«F˜ðF¹ ¨²¸¦+çëËÈê¿\^Kål„ør¹¡,eUÑ‰fYh_±+U¥O€šž÷›òsìÉ¼<‹Ò,À÷š£<QoKU¹y(Ü–’FëËBY\3ü¤mÛ?GŸÌ§X‰>‰Z€:æßbÜë\Ê®×TÐéO=²]Šo8¸ÝýÛÜ®L£¤+¸ä}E·+öÿe)
Qa‰”ä%ii’,¢…d$»@AcÖŸÒ±¬o•Õ™Øì&Ið®¥ô_´,—Cíæ$[D­Ýþ!Ž&MÌ oo×)µ[ÞJeJæ(•n’U˜kBóÛ…ûX>Õ4ØÖ½‹yk××‰0í|Ìµt"óC$+e¸ôhRI&8Î‰ÚÊœ~AËü:X”\ü“sQÍÆ~RXFjlÝŠ ïÀtïÅ„ßÕ7çnç´]ô17™ãÖ²Íj5§ÞK jß˜ÛdßãØ]§M{{mØ×l £‚{{lp†ñè, ö7yËf³ðýY’ -y‘†É‹h@™ÎÛÒ·[íY4øºÒn?&ñ0Qêý”¶‡€	šÍi†ÍètÔÔ¥¬/¦šGáO³0Í0Šh0GñYè®›Iø“†*áKÄi>y Ïa0JÃ¯+u¥ªmÀY!Û´.ŽHö™É#ï"KfÞ=<»þ¤±µ';iTBŒñ;aMí4†Û¨húQÞ]/QY}?È’ë_û³QÐ ?ÿ\iß4v ÍGãÐ{Lã//Î°ê%~UÙœR}ý^=…s²„<¡©[«K35L¨usÞ’4ó?WT\²ŒC ›§Ùu[óŒ»gÛ+Ý¥™†“°]Û€y3¸0å;¤}ïi²¼I;Í‚l–²BG¸¥ÊËî
e«?˜N@ìƒÅîbGå7góÀIøàÃe|ög©y`éK-ÀP…S4xÔ HªÅ°m‹#ë†Ÿå&Ìù.PËØƒ™¥¥|9ù06ã¢tìm\äeNiÒzzQó“»^"¬jDï¡b<ÀhÐÈ6P€}z(‰m^ŒL:­mú¬Ä­ïk­-Æ¹æœröw<õ{à2½_™sUðÛVÖ´Š™›Ÿ÷ôSÄ,¶lŽ£@<¸¹õ—2öXnvø þÎ’0œäðÇ>ùÂ:lGIn~Ì`(pïC…˜üæàˆW…ø·'þÁßÇ~‚`ðSXŠå‚ŸöËò·ïëe’Vÿu’ZDwœý;‰ßž,½J2¬§Ü
sqj^9óÓõzåd¥Ç¹sö‚ÔÊ¨øãúšl{â¤ª/žü×»'Ú(Œ³ÖzûqR…H|4b`2’€TE=˜e<ÃZkGaJ»ú™0q
–Éz³nñ·5úß\ 1¬yjpÍ/È½>¡Ü6¡©Œ€Õ0°´¥­Ë:sƒú·lÈ¥ã¿HMæ¹¦'Ÿ“ÆqøŸ
G–A$±à%æÃ¥B‡^ûÃ^™†ö'&‡ÖŽ>3­9te^f=*Á°m‡Á}Þ¦._Œ\Kª¸»6í{*Kå§¡U+·ErWÉÆE¢ŸJè"kwÒ(]±ŠÛÉW•`HÆ¯^›¶Í¬"4Ÿ’]½8½]åûËQ £Þ•õa†—e«{ÇÈ
zF¸Ø,QtÓq¿œo¿5ìgF¾2í•4¯`q´tMJº{³ºÝ”Ë¸Ìèì.u&(¬lIž“ALR…^´Ûú×ëÚì >]W4fUr-ŽJÿžo_){µäEþuÜ¦Ê 5dÈh"#áEtSŸ’t¦¹`˜†N£Q4 'ø#>´oÝîYvÈg\v=¼Õ2¬{DW&ñ»ÑÒ§„²ªl´ÑîÕÈ’Ù_vK†4ÀÜÁíÞ“~š§†8À ÀŸYr\½GÓZ†ÀÝ ž¿‡?Qlõƒ,zc±FruÃHéüívI¼¹Å¿Ûh£$F^ÈRÉnb~YFG–-ðò`:VkÁù1r]-ŠîVîÆötm{Ôgì[ÂVïaåÀÔÏQÚåÊ
òˆðÆzu¶ý…[ôvÄeOCo93ê×Û!Óç:Ë9^K‚9,kkä»p4EêÂ")FÆÉÛîÂÎÎÚþþÚ¿Añ˜ký8Á¬h˜pÐ³C“n/³ú¤'eäsìBþ&B(¾ÚéteÍF«á4_àCÐ¦2a¾çh‡œîë»—´ÝÉæË«5ñ÷†ô÷úË«×Î~:dà0Šp›Gn@,®ü€X¨X"@È•&‹ÕîAÍTX\ˆ¤jUàegS“¬Uåéã)¸{|ÛúŠœÃÿ%•ÉreKœÖúÉ×¥c(—ü8¥÷J´Þj÷Ïƒ¤›5×WÜH¸(»}í¾¸C¤äU½kš8~]Üª›ôÞ&óFù-×®¶ôX‡ln¶¦_Êªð¹ú&Vå}äž™uÎ#¯àÃXd½J4¡®HLýÍm¤`ÈuIÍRHÃ~«¾E.âì2à»¹}“ÓÜQ‘Daú˜é,™Ž$†‘þƒÍí±\:…f'ï‡~òîÅ¬~‘‹¬]Ê{ãÓ­/ /tfÏO—¡vÞ³3¨é„=c~c¹”Y1ÁöÒ(Ç{ôÇrÆÊf›â‹œ‘Áý22`	²ëÜìÜnjgŸDã¥në¢09.Zo©”˜Ì&}dÉ9bÚ(Q›…kˆÏä—ðY0÷Ú<àžwõí¯n¦ ¤–3rë©H•Ä‡³¤+…0[“Ž0c‰ý³5©H!ûz7Ãã+y>8~¨5´ aÌâ(ï‡}ü0DLÒ*|ñ7/áºô¬ÍM»”ô	ØÏßæëChèT$K«ZÝ¨
ˆå—¼ó€+[ã„F6ÒÆ|ÓŽU	´W*8ù£†÷fÒ–Ã¡U_`rÆûÊP¾;ãAnÞ”€ÈäþXÏ["÷IMÁûôŽñÌŠúÚZ¡O‹~4¹>µ^ÐaõâëMg]Å/’3_Ù” …S«ÂÈNñÉ©…ü*¬uQÈ›:ªoØ¾TÀÿp5ÈØ¿-½¶ì”’šzï@æ¼þ'êZcmWy¿ .l÷
·¨°àãö‚ùßqv:7
âÞJ%³—x¥ËùtoÞä‡^’h ‹P¤™Ìòi¼³Ž%ð4Èöqeœ†ÁÝ-ÿ±êô&~ÒT£T<â´Í¥é_­4Í~r§ñx'
<ãTuƒqª¹|~9k­n0ru=ÛJïç1Ý:)úóvšÈÎôû:qœyÏ7Ä™âŽàÌë™Äê,Êf#KàeBÆ!*ƒ&4Ú‰&iŸ:Fƒ¾vûÁ þ&W¶N`‹tLâ.ÂõÆž¿Lé3ûEâPÛ£NWbU˜(ÓÊW¦ Ó¦jD˜ÖXVÂíÙ‚¡:ƒðcí}sœ0aîÅT0	÷¯ŒúÓ†Á’·ãë_ñ®C	äKÙ3†ð3ð³„çSB ï^tÈv÷ð¸{ýŸ×ÿq@¿ ‡G»½ÝçÛ{ÝgdóA{CÇ\áýØ¹µÙ×Á¬W`?¯ðØ}ìãjŽO|BP;nÍ2ƒðñglß|(Ð·j")ëðìá;@°9ÕbZ
1P‹¦¡1W£¨d<-ºr Ýüß£i­AiöÙ"gáœ«©Ri¦Å©a†§D¬síNbb­øžæ‹äÇ:Ñ#S…Ò)Dóùqèöœ1ÅÙ?ƒ bmý"]ò–mø`¿Â˜å–ü‘\Âøm¬Ñ]˜†#×bâóûÞ	þÀè‚Ië—/ƒî¼ÞÐýü¾/œOÎÓ,lv#Ö´•—““Q˜çÂlXä¾½ç½mò¥Q²„ãµ4VXÄÑÉ_'Ìb¿56-NÀ,]fÎÉ“)ð½(`£ÿüé"àŸ¼kÔï&˜ØÒ6®´‘õ4V°úÓøG7gÖ„°säyN ©’d«€ì%:°¢kÍi{_vœ”ÐqÓûä–qvüôXM mIÔOaÁ™ñÆ•jÖ(ŽÐ-ga‰&ef•6}X%kâ³·ŠY•õáQÙ~÷˜<Ù{þ›Ñòž°“Ÿ?ã³OŸ	mÈgŒFnFënww>£4ZÞJC=àgœöÉà4¦ÖýŒÔj…üX)ÞòzËË-¡ó¯§–ž²² ŠS´ûƒšÄJ›QB˜y»sCðlT« EªPø)æo¸·5ÜÁþ+ª»“1xdCÀEØU,›Ë%{eú®>Ìme63¸‘•¢ZËÞÓU…×<é3üTbföE	÷b>F¶Fþ*`ˆ¤RîÍ`nðÀ«ˆÙš?‹ûÁÍÍ{Ô™’žµ87<ÊÛC»è÷c1c<»=’wxñÊ[-…@3ñ&èâE:6Ý’pi‹ü³«}°Í›8xåy–6’¾±5d²´þi®j£¿KŠàm$ãÊUß”‡¬ÛÃuûØ>™w£oã+£m“ÝÉšƒôÝ¤O\0%ï¬Lað6ˆ2oÐåz'îÛÙþTœ®’íƒgÏv·÷ž÷Úö_íìuW	Ç»æÎÅ ÆýÙ(ƒ«r‡vÇ>-ÀZòxr.‹Â×”É¸{)ÍMPh¸’ALÍOè·h.x…ik’0HÐ~=ç´®`#³þ9i†IbÏm)¦ßØM’Ó[4ùUBÞDi$b¤´1Êt'6·tÓ-ªu½Êá’É¸=­MŽl"jé<öcd`2	È= z2ôd<þ­i¡µîy	¾Ù}‡Ê½‡Ññkå;¹|oÈ CpÄ¢ì0JÆÍÆN˜†?`øØž„)p¸¸¼­ÆŠü\˜Eà–M	àÄ-KÁ.ŽÈù¡w:	Çñ¸Ãm{ÀùBvPáÊ8[¹—boé¼î£Mc±È]-R*(×T¤T˜Jî¢¤ð8	Òóia]ÙWýFJä+½ŸÂ)ƒ„ÒB;!ÉëÍ–Ï
eà¯´Œ)Ï[…Ç.å¬ª†©Ò0Ì.[¯E5$a¦òÉš V¥Êå™ >á›'¡>É÷è5?Í¿©c‚š÷|³F¨b·j^Ó„ìXn$³jáíMóCê°¼ÒÎâ½Þ—$WDÌœãÆÊÉº1÷n^´.ýÛñÊ†J(ÑDÿ€k©Cå±°Ð4ÉÌ{¦§¥VCTUåM¨mœ£´×Y°*á ©ëBìÌJÚøDD.ÐìôMNƒQ1ž1÷^ƒÑð÷ñÅ*¿ûË7]TÁ^µ½¦"òäæúæý¶&Ò§Ét´ô¥§Qô¢€è- {“ë_ûšdãõ†©ùQ¸×mW=*úÒ Ë€`KàæA`|²‹pø¼…FÕŠ·Î€ÜßöÊù‰Û†Àn€kqÒ–©º5ó‰Ùx¡˜ìÜí@G2{©7øŒ–< îH0à·àrÑàvÀ›$©|·Å\!ÌÝ: SeÍÛyeù÷S¿EÍf‘Ób=sOŸqÄe<ã6‘ÙN"I;UYógëµŸØ½”t³èÍrŒc\)<’	˜ÓÑ˜Óø'p'0¦A(Ç®¿Ïƒ8¡6Ä¸ÞÜÎ¼ßÜø0Ò,ìÊ<=äœÈ<ÿ•Àæi˜¤ö|ÃÎÝôi0¿ñaÐð˜§e—jgL-õqCŒ©j$1Ðñå	P@ñBòR»’v'Y³¬’Sªq¨ª­=B#©p;Oƒ$lžÒ/µ)»E.€¡õ­VäÛMúÇµJbù–¬+b4C´`bÉšÒx…|I`WP±þ$ºÍ[<£i¬ããsGprN¾à$æ\@ÑÁÒqLÙˆ<p«2°1X>ëféQ˜Ò8Ø˜+ ˜¤¿Æ	"Ôàü2¢	>dó@®Ž½¥·Íº³Ål¿}T™®u“Ö˜”îpGûAvÞNÓ&ß¼8(úÝ8¸hn¬ª›aš%¤q‹™ì9XR4GÞ[Æ:éÌ~#Dñ¾Ü¹NBV`Q¦G€|HGä&ã”<Hó‰ªZð¨/yž•[Wì5¸pjiØtVü!Ž&ÍÆZcåê²A~[î*—ù•þx~ S‚á<|‹rìZ¾eÏ+‡I0øMÍªZtfŽá}b?—r2n²Øj&)Œ3—â9Ï±‚t%4YlLÀXÞ)[ˆcüÝ½Ø¾Tš9œÎéê_í{âj·,0ñIyE'^&¼·én	Íº«“²ŽR‚ÁÊŒï^
ë[–—°¡…®ÂnÍOŽeóñÞÇIxSjÆÏ]­¬Ît7…J¼Ý/l~¯¼/ÏÀ°Þ.EÑØ0‰¶xCÕ§Z˜ã ü‚ÀÊ¯Ñ2Ããßš>Iæü‚_;ùÑI&–~­¥'‰œù,+Ë¨Î§›ƒHQ*®"^óãî$^u«.'Ëš½ÅÍT”õT)Šëê–½WzDæõª[-ø™³Åm‹.J-—¢èlÓŸínw{Ç«dè¶L/Šo¾ÉfèåS«KLQ¼nŸ§›LQ^ç¦¬w/‡Kó˜©9k/š¢Ô\§ËíÆ·æ°ÌµßVZ®ƒOQ\®>EðäµYZ¯ ¯–u<‡ŠâÚT7–ï‹–|P¶ÈrÙ}É³g×ÜzŽP^Pò–ò½Ó:§*ñ¬°å{×=jù¢6úXÛk«(K¤‘ž”IòµÊ‰Ž§WWQjûwUG÷Þ¢ŠK˜ 
zQ“bxï¥ORŠeS–å»£ÅƒßU]Ô–Çð:³GØ%y[®Ãaêò#2#hs",Ñ¿l:
&“0ù(ÜËêˆèË¨Ë×![tiO+Ö!~>2¼³›w‘yŸŽ0ÝŒçƒ×ÀéxÂð]lb`\Ï‚ös¼O‡ù¯eL‹ Oè³çÂ0õM’8“„ 
?F€	ªŠÓ†«S˜Ä2Aª<DôÿÙKÁáÆ(¤õÂq0©)0¸}Tz¼;þ5¼©³Î¸}‡­·’Ò.²ÑÏ`  5£DŸc *‡Éõßuþ<’ÌY'š¸êô`U®*×¿œš2½ñ¸û¦ dX¾‰§”a¡V$ƒ+<ƒ+wi°3ùfõ¤kWmÀ4àßˆ5s‘á)Ðs[Ÿmš=ùæ‘mTiœÍ’ª§Íg³æòà†gU =sY¼
pžÆÈû|f¶\Pô³²µÛÉýD³Js<[/ÒD+z'ZÏƒDñzN2Å'è"U¬š…\ñ*;ñ@Ú\å¥Iä4¤ÄmnG »ÃfP°E+¤¥þv*ýVW‘¡Jˆ§Bl7ó¤äyj5ò\Ž‰§E«³T£XŒ•+åcÚõLÖ.'˜‡SÀ^Ö\\–¦0›P¶È4¬ãú¼VÌõRá®uœáÚX™C¹¯Së>ë>¾{´J¦~Z}t®ÉßƒÕÔÑâÛ5Î…ô{PF;tÅõ”Ðn´ÀÍše½ˆñaÈ€í~S
dîéWSŠYSÚâyõÉlCçÖ'o.rI¶ùæ4Íb+ß›Bmï2".ñ®
ûüÅ[Z<ÈÒ¢AÚbÐÑìšÉßîMÒ~qÑI—[Ð‘Y]C¢[sø÷©ÆbŠ%E5A=GjšêµŠ¨-;ˆ~`´ ÷ºeXJ7ÍBkÕ;h©Xš¢ß›ƒ´cæi´7ùžz“9±Þäƒð¢O¤lÚÿá ×äOu« wo5ØÊóûÀ¬žâ|dh?È0RÛgÄü $hŸ&Ú»Yžßoà"h¸ýXˆÏíØ|r¿hýÈÉæáûLwøŠîà!ÜÆ{¬Lð7p—?4ä~D”ç¶‚l1»ß ¼Ö =Ë´ÚüB<%¨AûûÙ8ÈÏ8ˆïÞÒS‚þÍƒæz()Is'
Ò¹¢rÁ‚= ¢\ ‹ÛeuÄ_S·5‰Ù–ÄiI²;ãûímH`¶ð`ÖgäúVš\´ž&ú´Óò»ÀÊ•;—±/Yjqb5fÚ5¬0'EžëRÕ°€Y~+»‰Feì¶êR­*ñÏ0Ž¹þgFe
QËŽ}$ Æô‹B—èå––ºÈU¢ÌõóùKÂí €us{! çW€ ›íGNC¼9Ìð<BøØèM9g½åÙè6â«mÂ§5àÛ=Ú;Ø9xµÝ=<înw<-ùœÆx¹_nÐämÉgƒ3«ßÛðYòêØï¹¬÷Ì÷Ád¹w‹ìöÞžYL§ž†{ý÷>%	{o’…g	OI÷8	ëèéÁÿê9el…{ ÙhU~>è1*¼Àûñ>F0—Qtvýß“~ É¢!Ð‚J>°
÷÷§¥óÕÕ¡ñ?(qCý3ƒVx´é[-Ï6ë‹£gÔW}:%GA4z¼#Íîá~G¥CÜÍ«v¹æ7/E-.Å„§[¼ó,›¦5¸Ì³ÜÉÖ$n%lEm¼¢†.á0˜²¿0îiœ1àiÏ’‘Y—ûx4Kr‚:(*-Ìâú"‘G%ýkØ¥qÓB$”˜wûÅ–:wX¥v
gé_£ì¼I·
vª±2_Ó”¶56´Óó|+^‹®î^²ï®^Û¨eÉôïÈÜÄL?qWÅ4 +‡¾Â}‚Ìýº1f%3'§¢Ïþ~üjûàù“½§«¤AòÅÅ;Áš:|¯\¬£niCØ7äçŸÉ0¥Î° ,Æà ›u0R¥ãhŒaïÇÓ¦cŽð}—dÂáw`óÑ
Ø*ÅÛ€(çÕ8ÚCjd34¸cã×jÆLòŠ“ôZÄFJ™&fÕ!w/a„6ì[œ…WŽ‘lk2übä=ÍfÕ÷|Ìª‡q–vPD®TúÇþŒgÙ(š„€»'¡ÍÛð&ä’‡jÊBˆ'î”°ƒÏ™7ö€ÇN2”¾NË‡QòÀª80+”º0³¹p“ðÍ@ Qä>‚kÐFÍÆZ0Öà¾f³Ô>C?èî 1(êxªwHv ŽShO ·4¾o8=õûp'æõüú.íò>Á(,è½’„ðŸæöÁQoí`8Ä¥®tÈAÎOMcøÑ`5©î! á™ÄˆgâùPáúòÓ,<Å¨6AÔ@‘ÁíGâ3ƒŸ¿hj±ƒúð¹¹®y# T
brÓ$~›SŸðXLræ1;€m.$N³¬j6¨Ê(Üô¡dð0ÎŒL8Ç[ C'Ò¦´¸= ?ÀÕJ`’Ä
CÍð¢£íQðVãp­Ú³i[â…WÚÚjtV|ÕZ	E‘}24Übè>sË”´‹f¶ÏàV`·¸s0'G@S€5wŸ¼8Þ=z¿¾ú~÷ßê
7fY…I*SXÉ[Xº0Y%ý±'­7­v»­¯n’Kâ|I€±¿ßQüiPXÖTi F;­)¨p†š¶†ºéšîŠ‹Ë¶Ó1_{	üµ».¯®Ã÷ÏŽ¨çå’­l¯7‡l¡d¯Á¯X ¹b#M•-@°pÅw,É8—ZÌ¶Ç´ËŒvß¸ŠZü·“äš¨ýÞ@lo×­·ÃZýÛèçgrý ƒ“øM \s8®©›@cø.F â¹ó Ñïuµ6“Æ½þŸ$ŠS’h’ÎÆqº šo®ÿ‘ªgþ©’Ã}`<G|›€ æë­KøŒqãr[HIcÑQídIØBÖ#L¬ã2iúÈIß¯N¾Àœv4 ËIBlÑ­Ü4²K3úŒÞ9>/M:<oP­ð¸ìÊ7ßö-DØél,6¹y¯:äi8Y—l¶×ÉXàùŠ-ÜÀ ÄŸïìì>ë1“µ±= ’‹pŒ&j#ÑS°1Õ>˜MÁlsÕ½@cxBúÛRˆ$ÈÆçù	®”ØnÀ ÎRLzJIdÂºþe¥a*»0PQú©Ò8I‘ÿÃ	œ%Á˜4¹òeÙ2žù5ÊøU•¿©š-œ¼‰‚¤¼‰Ò¸ž˜ñEÂšñ	céà/U¥•¦}²ŸŠø§®Íï•åöpÃòz&¹iRÎç&&õ™š aMQV‚ ?… ½Á©„“Á4šƒœb7Êr¤†G3A‘KÛPl¨ÛC§éƒT?Y-g/œœçš‹
á»hÁF´€!ú€êMàüÎƒWiHÿ;›†	üÝ”¼ÚØ¼7±û(´ºÉ~jÄî“Òtru¦á.Ý
’§Ìð3áû€„¯€—Ž1	âß~L<À\{pTÑõ/o ×ÑÆ§J]ìÒ7m9¤MbñZ|+ˆ–ÎŒï‡YšEÃw­SfF¯¦KÖFyä©ö£Z‚÷•Åaá)¾xþå%z í×ÿÛ¿¢Ý-š,6÷fD­ðMÔõi:»RË2mÖŠnsûÚÆö:—ù‡‡~½~ãpLúa’…ÿŽê|A¥Öü	üƒæÇ;ÝªRq ý$šf1€î}f·H“jð¹‰o0ÎÔK Úžjöc@@ãi7%$i;e´Õóçº0‰Œ:âÂÿ&×ÿ,§+Žœo6Ï—_O†@ƒÝ™ÆªGôáq˜ÇƒiôŽ	ìÜù•½óï5º˜Å3 û×¿œE}j§r„³Üe‹ÞÀÉµÉq<@•‘t´Ã˜’><Cü)äæ)LÓ$ìWÐ·ÁÙËf‹sÕ0oØf>¦ÍK×>Í‘àÐ{csÍ¯GM‹=,žéñ1þfñ+QMz ×¥`"R­6´ƒ&²v¡6¨Dþˆ#¡ù7)RÙÎ‘Š	8]O|‘9Cã0ÞyüvoòCØÏöáÞŒšÈZüÛ4&Rn¨&TDýÅä;Š÷ê<¿ñMô¥ˆ\2›A,›â	;®dc':Ã­ŒÍ@fŠ¹%ÆäÁƒ?ŠÒ™×6Ž‹|³±‹°¹º|>Ÿ[ô×ÑNB*7×þ÷ÎÚÙ*Š„ö{-xÖ‰?Pæ#Pò¬c¸ŠÅŸ°äit¸hÀ(„dåúž ÛR‹=þL œöPÁ°„ 7[DÝ–œ	î—É;ô8“Æ:ªU!˜tÏfA‚òÍèl6Až3Ì ÃFþ|Ù'¼s’kkäIœôi¶CJïIÿ<ìÿH‚!
÷`PØ÷Ù÷–@Ö¨k ™±Y/Q¶oÚèeYR‹‚˜ööÅ´×6œ|ê“»Â­Ýã| Wã«Uro}}ýC²^…E9³³FóiÉ¦ºÝn¢ÌØY2÷¢ÜëdQöá÷,}ÃsFmçbÌÆÔ¥Úœ"ßÕþ1|—¢ÆU€;’:Ú£pr–S…ì:Ù2Å¸Ñø“ŽËa‘ íò3ª ž‡“s`N³Òk?Ã÷w Eþ!Q¤ïCã“†¡VÜJ@©èÜÆNdŒ)°B:¦ÅÚƒU2^˜–-ö?œdIj í5ŒPÓ<ÉIÙ*‰&Ãø¥''#$@P7ûˆ~HÛÃh2°#‹æÌd•Y» ®V?H,<²…ÔfEÃØ!ü9ïÊÚÑÜœ*gŸáÄ•Mc¬SÛ’*RãkmÙKWõÓ­vqÈ/ÙëÇ×º»,Š3ûcÇ–RëŸ|d“«A-©«1"mÕk[“îC..ÈÓj…jIÛjsbéSLÚ£½¦ÓÛ$PwæÕ)`ï“¡“0‘çl+{hQåW:Ã Ž/4O¸³ˆ˜'í’Oì¼j9
Ó©þÁ»2Ë5œ¦ç®KxÓD5#h÷¡š›¨Ðe$³	ð{!Ïu²ñ ·Ì{íÊ<™EivgYìÇùÖÛ3©ÔôkLP-îPÕbÒèÀIÈìm"³·‘³o[÷”³íóûNóˆFG€³ôŸ¹òÖ!ý›ªÏü™òA¬ÔŒ¾P-ùÓ?î½·9¹®øÊŠEa‚"{÷ÌE(—Aý&ÐèÔ-•‹ÁŠ v?Ìê n;ìŠj©wXN^¦¥}W"P¯_úabµpXª;k§FG-^j„¼oI
÷í¥†¾U­9‹bQÅ“èûVsrÌ±øÂòÐ«ô‡$Æk±Å —¯<-¸É"¸¯BPï^¢8´ÕæÚ.XŠ9Ì7€rÂî]‹ñ hÕ¶Ó BÌ/Ç	úÚ­ùgÞ\<ôä?ó'›ÆÕkŸƒóäEäInµ³ørOÛÀ=5W@Lð¸ÆÎnðÐw»Û{;Ö}¢ø3>ùlE™+¯­(¡õŠRï_.¯/úìiU«üí‡ëîe¸å¡3Å—MóÉÓ+J]ªºð½\êêéå2‡Î^.ný½\¼u¦ry-žüäóoýíç^“Í­—K}nKÕÒ×&ìKÐêËÅ¤¦çê³zÝÕRÜ+WÉ†[‰¯ÊS/—ùÀTXXptÅ‘U­h(åâ©å—K=ò¬éWÏ“«Ò]Ä •©¹$4rz/åVR¢‚úIrY”`Ô®r<HK¸£O/žŠ™z¹ÂófÎœáEMžÒü,!—óbJ|Pã0%ÆUVL%+ZK•qMž³ºÝ¥ÂÏ¥|ÛÚØ$çø3ƒùPÒ£RÎ×¬ääƒI4ÆuLg£4ôUÜ…C ôçÛo9ˆlÞ÷_EcõÑÈ”¬Ûx»(ì%¾þD“BöÎNÃü°»ô&Â`ÊY8ÁOú¡òr,±´­‹1h9™>`Ê_Ÿw Ì Èc¬xÏà4Ekõ"IàÕ]¾~r¸Ûûž<‹&?”P7WþÐ¦ÜPY²ÏÎxACØT,¢Ü\Àù0À.üÿSÏÆ»2²§æ·Pzú,‹[GátôŽ4÷º>òB©mÉ‚~=9pAIP0§ýY’Àd»,Òf.µ‰ÀÆ<¾ß­lµ£÷šn5ïlkkþÞ‚9º 1Bý™¦\+ÖNî(›a};TËÚ9˜fÑ8‚ËÑçúUïÆ%Æ¼9MÂ7,ê@N·ÝncÃ:¦ôZÚ©)ÚðÑ¤ê±Ã"Ôk¾ûõšW`mÎžjÈW++5 bIJ„,>;…B‹PS¸,Œ+j6\–
5õáïC[²×eÜ†ª-aA§/þS¿¶Hãà9Uª<yâhª–D§kOˆncn±tNéÃQR10GAÿÇzš«%Ð†¹¨ÃBôaq
Q †;ËÄóvV9Ü„îÆ·¢ü
DãÐ!rˆ¨Ð¢,êyëyÛú£Êªª¯D•€¸šçû»—MÐÈØ1vm…?	‹ñ(”«<_‚|:užÓÊ›É÷EP8oÝCË¶Uèï¡î`žIEÚÊ¢êò¶ŽöI÷ê¢õ€nžüÍ†÷îy?¨ú«¶hmÃ&o£¯½óìkA©ƒƒ701
í]]PrÏsñ[—øÚíx^wöa‹nÂ¸†1ció5ùei¶åj’¿½•©”ˆ¤¿­™h<èÐ¿“ø-þmÕ ÕÎ"´pB',ûy\ñÕÿÂ/_7ªÁÎOXž†I8éG!_ÿs”EÓQœò¸x©yˆ©i\bèù,N"ÒeA··•Ó+%|&ÓÓÖ&žü»Î5ïi?‰G£Ó AðÒë¸/O´‹½$
”çQšÁüû1È¤4”|÷|‡¡né—ðßI‡|ó­÷Ny71ÝhÞ5ÆO¥nóÏ¼Ç hhú]Â¢¥Ÿviråëÿ+s}
âÔ,Àû»Lët?ŒÂô¯w$ò¾\Ë¿åC<‰Fá1 Síýx ¿	úA\Úòìú×ŸfÑ ()?EØFyÉ¢[Â–ÚR2[µ‘Ôê:°$fff©¬Ù«z93þWžÛî88m²Q™vg„¸ì[røjÔ´ìrˆp1›²ý½F­¤0ÔdŠÈ‘¼EžÊk¤Êz¶N¹TyeÙ‹*ò`Ø˜Æ`K.\¹NÝø¬WI«mêßp8F¯ºb _c TÊÔÐjú¯llŽ5Õgmœ“ðK¾§Á¤:!ºŒW5Ut˜VSmŒÁhX\*ÝÏzÔü"‰juÖ0NžcnËÄkøÍ¹Œü—0iJ,S.SóÜ©ÉlôÅ°BÒÙé3œBÚ!zJÜà»4öƒ,¹þµ?Åürð½©þÞd òàî8:Ó÷­QàÔÙp™ZZö]¡Ÿóoz™âzíw±µ–m\p*dÝqm$B?÷õ©p¾÷çe¥J{¡X‡hSŠ uFV!kUÙ˜3¢µnŸØÃWå¦ä‚.UÂ[D$ÂÝçmòY²ðÝã6öÀH,›­^fï:±vhwò’¹–6é© ·5¸p<´²ùC'Èÿì0×äÝK>ðÕ«»—Ðƒ-×$_~z5³ûžVWœÏc®H…ÜGÒCÀf»…@ ûå1¡ÇÐñ¬
·8æ#Ð¿™g§­É$‡v‹lÕ¢ô/Q‚1;äto®zeP¨ä @Ã±‡â³t5ÂñY©ô„ÃÅKÝj?û"Æ
$°ãÖÎ®¸4pAÆâ†Ê“ê(`Rü ÛýB­´”GbR‡}wRêJû¥¨(Ns?´5jÊö çð7+!óLrvtëìážU­©1l[GÃ6ÅXŸÞÙu§I[!ŠØyü„,Œ8­K›gRùH5ÎïWô-§%7V%.£äéZÓ¿•­‰I%åù}Ç¬µ&x5Õ×—ùÍe±"®¸²¨™®A”NãÉõ¯oÂQ3r„]qØ“9ÉeJê	ÕìÒÌRo<°º%‡(<s†Œ¡,á–E)²ù0_LÇ)Tõ×{\õ~PÍ|§äééŽ§d‡û¶”ŽŽ*Wo´øÄÃ-©x¨ðq8ˆfãšÁr­·ª;±))Ò#œ=Š(ÅÅd_
Ïe¦À§á„"=Œ³*¿Uâ-ˆµ«÷Ãæž‰YæÛ×<[åIÉïÉ†û¹ÐoOõ)jŠ>DªK%yœÆ«~†“sLÖŒÉ÷AÞÊ£rƒ·§5ð5#ö’äk ãï\-»Tÿõ»îq¯{xøj·×ë>Ýíùuâéík)Âd Ž}—Á¥;ÿ/&¼¦1£rh/‘bù¾=ö“pÎ,ó¢x“x™,‹œa©e#$‘5ðžgèi“È§çÆ³M,¤•=››4TýHpÊÐ¹ÕVdróM£Îgá"óuâIºMLÖâ5`Z¦ò&Ìø¸ô¨÷-gáw•oˆt«ŒHt8š¥Â…fÝéBãc„b¡@VÙÎ![|¤Öî­[äÊâY^÷VP”KYÜ·GÎ)ÏÅ-RëÛéy|›ìkîÇ+ä‘5+7Ò
„÷sÔñÈ3ÊŠA#®î9šSÕ”*ú†ê©¤§$4ä\›K$b¨Ôà`2ÎóNÁIù4¬™U-4p1ž_ô`Q–§ŽåýÜápjÂ©g	œ+2ìøsXµâß\rþJœ™·™¬§ìqd¼ìd=mäRÎ*³é#EWólùXÈåŽãñ£Z¶$Ú\gX5õgI' ™gÈ ÄoCCF­j1|U7{Ž×Ê½÷(‘Àv®ÿ¥Š%ôýM€{‰NóW/{ÈZB+íÕÛãoo¿Q_J1_»ýÑ,JäíÜjøY¼ú{†0<}‡Ÿñ°T¼ñ°··—¸X×ƒt‚Ïì€pÒü\V-ìÌiÖ—0}áÏOØÚ·‘¿ßV9£xŽcþÝö½rÄª1ÎZ¸tÇlÄi(ÝÀèìÅ†Í[ë*
ôÄå÷‘ð°xYŸ+5<d=ÚÈÓéÀÏÝ^Gó'÷ÝA”Å‰c\ªàU¯¹ë`ÇéÙÞ n£hÇŒÐ°EmJÕb®ÖT ytù—„¡ÐeN+_ðóv+ƒãÉÊ<çDi¿‰qvfãŽÒ‚†Rå:_•¿áËúã MYÌæ†¥íÜ À—t-¬ñGç[Ž-…ŠyS¢aÖ¨öW{ø“ÿ¹F8‡"ÿ–Lt)ß ÿö‹+‚½ý
½HÎ	GièysÊ¾µ8®÷§óòZ7ßsƒå¼Ykå²T°ª£4çÌRžºöýªêYÎÚ÷ª ·r/™ÑÜne3üAýÄzYs*ôÃ.Í®Æ*é!Z%=¬¸¸J—)’[™Ìù­ÞO³ †Þ÷6=žÜFGÚ´+²=ƒ±Ì9Q©fP‡“ËŸ"BDÃë¿c°SG!rc `“7ƒŒL“ë¼‰IXŠE/bdä’-íÌÞ•ÕvUÀæE©­Sººdf“ÂA¦‘a¶%8ÆI—BZÝœ&EvÚŽ™k#µ|2ËþœúBÛÁˆ—6¾ìô8—»ãH Ñ1£i$OšmpÊA>	‡¾
_Nc Y)”á¯Û.MrªÂFX4ûkzéÒ Z›Õl)ÍÈøT·kêÓ}d—-¿<¢þéèl¸­Š5&ßíà§YD£k'ctÎÀÌ£lãaÄë¿Çé*SäÂ¿ÓxtýkFsìæ¢Uy~ˆédMZç­“{,_eNjB†"§O– ìNÎøn›ž½¼³1äOK¢ÃòëŒ›Z~
©RK@DÆ‰Ú¤U³|Ê˜ñ4Ì,œ8rÞÀP›r#ð\FÑÌO•VÐ‘vÉÌEÎÃ¤™ÒK žpÌÅTÍÀÊå<`º¹Po¾Ó·èµÊ€‰Ly§g¶Æ ÒÏ¾JyË_U˜'§)ó}ÅUƒùv ÇtÈŸgn1@[Š]ÁºCŸþÒˆ›CÌÈèÄ*êè<Nq…èUUé¯D'*”AÃçè4x¶°µ^¶ÆÚÆ)šùf³	½µ@’tÌûe„žu˜ê®ñá`h°©coÓi4!æ£ñ-5¡ÕûÛšD—Ê%òz1j|{¸óÄØm€õ>ºªÑwÀ@ÁeLna>•Q«.>N œUD²§?©'@Ë‘“šMGq0@K"‚h=\˜c¤ OÀÉùMª^àAáàà:â"B¸€Á$¼¸þ¥5¢H¨/q(‚p],ŸÄ_Š§qDÏày<ÄÃÊ{¸_îp\Ê¢üžh•Æ:² \Ø<`4ÎaÆ8Ãh¤Môûá4{ÔhOCÝïÒ&ñ8-šJ…ü3…°ŒôØuâ× JƒÓQ8xT^yµr§ëøMÐ¥|Q‘ÁÌB×(>û,mAÙdŸHÈ3Øê:7þŽ&¥½_ŠÜUFt}Ž1Õë8Hú,µJ (#ì…ý…Lèkà¤F8)ø£ÔeÓ|Ù)™¥³ë_’(NÆ|»~ådd6]Ôû”-[¨#äßä˜µŠ®Ž,.[õ}¦P9‚Ð<Xö HÏÃra`¯æcGH	]—}ïêÛ(edƒK6ð¦hR$¹GÞcþA…O¥éi¸“·«Ä‘º»ïaDÿ Ov’àì Ú0RóØ¬&tê¥=¬÷|
XóKÐ—9=ûmšL<õ2\òt„Œ‹4J·˜áå˜†tˆ§	r×æ(ãb³âÉ¨Œ”!¨0¸á´ŠµÕw­Û&ÃÞ)¡€ô1x†\O$÷ÝQˆ>~·7hR*ÕÂi´(Ío¬lµ1ÃõÞ‡§f6îµsèÿ¨2'½ðÏ™ŠÜ×¾ÔWòâ8jÛî– †ëC\°â-À‹éIF3Ú×k¾ŽµrîÏ+†a¹,!T‹	¿‡8†LEÃb{Cú†5gjí‹€Qåï+~å*ÀS£‰~y»ÌôGâþ9d1ëºÓM’ ¥²Qàß8³P³ªµ 79‚ûèLëN²iT°›oBLY4‰IÎð°0È:û"ÃW¡Äá\ñ$6Ì¯Ëö1ø`5ÀdÒ9|þt•üéð)Î¿÷—§¤¹\´ÉƒýÇ:ËOÿÀ…zû«KÄ’‡@
£ð­Þh¶ÂuêõSvkÚZö³V‹Yÿ¼ˆ–÷0	ƒÕ©aˆ¢ýHˆÓø¶	WTâ; À+¾¢ÄäúWL4 C£mYì¦M¹6†Ùžzûš¨µ<™»¬|ÔØ¬,¨ÌfÅe"@e \—Û_Ôa/á¥ fÅjq [ó¬ öæd6YëçÊí\ÞV¨ˆ;H:C”h7±¨iZá ?Wwh°ì­ìfÅ#ì¾ueøÜ8ÙaRå,j³¥ãÅÙÝc›,«äÔ’9oñÝC‰¯ªQ¡ßRÿBù=o³bÿ[’u$TO+ŽøaïSÕC®©<ðb1ìj	ƒÞ`Àª1€}8ù—õû÷Ö_:èJ8 &Šª‹å§JYÅ™GÑ/ëª˜µsjq½§ÀˆÌÐu&Æh£Itvž±¨£÷òÔˆâ)um£|ºù/›2çz_pY'Ì}&‘Ú«±r€Å—Ÿ4”³··ÿâY÷ú?¯ÿã€ìì’gO÷žëi·ç£	Ü»Õœ$pó‚QëÿèoþËFx/ø*xµ1½X¥ ‰
¬I†ŸW^’ÌqFçÛB¾µ³ñpzñ
ÿó’€Ü²w‰Co‹
Ð);än¿c±‘m¦P¦(F
Ÿbbë+‚¨@!Ñæ*Ÿ™itÒ$k¶EF BŠÙ!š3›ëËÓ.Zô…Ši])ž’Ôøø®ð
ª ]®Ñ´÷èë<Æ…gzCƒ{}ÍÃ¿(I8*&‡1pgï5&qK|ejdbÌ&JFÉ*(
€meq+KÈ0‰Ç­ôÇwLeWEGÉ±±wtf'©¿°åp•_LÝr\øv}8}‚*SžË‚a•‘—1,zša,Ñê26Mt2]hìÚ×i“8ù8·Þ wŽ*pÊGßÖKé,HÑb{€1Ûƒ(íÔ”dœB‰@)›:Ü”£®M½¤XDzÏ–å:Ñ¯Ö×])X1¨¯Jd©ønR>ì‹aÿ kÝ½xú÷5SÐO <|…;ùJ“{Ô@F6vuñ»½ã]òÝ‹çÇ»Gäÿýƒl®o>tSÏe>ÐE“Ó:/tølSï•ŽÑ‡MÏW9ŠËl—‚58¯öófdMÉõÿÉÂ(­b”K4¬E–jéÔ
&]‡x›Ý¢s³õž‡oiá(‚(™%#öÇ(îôO­[¢F°Ð¿!ÞW†"œ»FÊ¨¼½—ÆÐØ;“>ð¿	%Ó(ÑÐÖ±ú
Á©‚:Í‚Oz³Óq”‰òî`€{uU¤œ“OÇEÐO´à¨bªÓÙ+&–
¤¼ð$üiy(}ÍÃpLØ1R¡ò±”#ªÇÞn·yãUªÝaÅZÚ»ïÆì4ÕíX«»£/ŽžÍ·¡pæßOz‘>Åí|†ˆ4¦Ô¸e´RÞÛÒ&R<¢7‘õßLŽŒ>ÂíÔ"rö<—RSÞ>CH-Õ¨©º$«)¦•ãY§ @±—:¯2êUØù¦¼	·ikªq³Êš”*rþf±¯|¬,“Oî3•H®þàRô4©èá/)/Â¦kƒhjÍ$¨õHë]+ùŽM¶ÿt $óu¦n2hõÅÒóôÚÜwãÓÙHKvèÆ‡AhÄ°‚1
‰1ãz‰³ý¤¤Ø[f!¬Ü”¨Í4iß-†ÓÐäô‚?æ»Â…I¹Âüëù_–ìoJ%®Ñ¬ù×°8Å!ÕÌöHFÿÆøn-µ™¡ÌQ³Š ©šgSÝÜ&jÜ%7E-•òßýÊ‘htïTÏöžßƒíÇ4Æ3,^n{;c‡,r<æ Ë:%dÇä›hiÅ&5Ô‡‡³	³
Žâ´·£‚±*×“`¹ü„Í–™c’Þ~Y”µÙŽ6¾¥ž:×ÿÄ—c˜$”µ5ÖœÛVÞ´î`MžHþ…o¥IrxtÉÿr6Aˆ“hÂôÑ¥ú¹¨¹æžb?	Ó~DM÷»“Y0*Os»ô;¶4•´Žq°ÕïÒ(í±/®”¥ §…+Á»JEµÞÅX“ýeªë±0dÉf#íÖ?)~ÂM'³-§IŒ6Tw®ª[ï3¾² Ûãu5
t¢^Ãû£CUäuGhgÅQ°Äc:u’^x6”vý÷@ÁîÛ3–yÈ)+0¹§ì›ñ§ÞÁsæXÚåÜŒšÚË Ñ0œ:pÌVé(Í'IÍ¶£³„åÚj+sžÚØq»¤¢FÓÒÂœ
2UÛC·h§ÆûË7Lsò›‚U¤‡XyƒViÓ†Fvræ2¡BÉ…Kjúf'~;‘­Þt¦XZA8ã áà¤cJõï¯~³ò’dIï]ÑI”þ»RY£jügÞ†#ôÛœáVì 4ûî„yVªçØhm´Z-êDÄ£ßiƒ3æŒ•Šßà÷ßÑ'€!,-¯YÔÂJhR.X_4[ìØíÇaÚ!ÛÒA|BmXeÂO/àŸC†7Ù÷cvHsÜ!i†ŠˆU’muH#õÑ¶¦A~ÎC‰ ßø&ŽØcß-¥O¨‹ª`ºÎðb'A’J1ª~MÍ
È#rBsÌÑIçŸøTW‹ÁWwÕ±m3ôa&¾'i$ýóã0¯¢|ÑË?¾„`¤^E³ÁøSÞ†ûŽ=¡¹Ôh³]ù¹å7ä(q³žm< þ¡ñ-£-ôÇÕßqk1#h7K¥Î{ÒJßEØï!lÂ>$ìÏÒ`ðÌˆ.@ñm“¶GLÂ38
iÄ#éuxÓ¼íY˜íDpDH¦žÀt“f`NÒgIÄ ñE}”Ñî°ªÂOœl0€§ú‚Þ'pvpy«ß“Æñúz‡þO¬~(×g[j¢áÝ´¡Ñ~·ƒÓ´¢=~	rQ‹¤Å§JËà]*ZöÃhÔÌ»[#M@èëäKò0ÿÏæý•L«Àzák _>oŠÎVÈÙ ¼s¬z%íù4x‡}¿MÙÙìƒÕ”å7V r6Ê¨mÛÉKz_øÔÕ{Õk7èŸ7›½’ŠH‰}^ý(àã,f‡–ÿÄ¾ÓÜó0fÙçßZ¹˜Æôª0
Ø±;' eEcêä…–ÂÌÍ+Õ™†“~À2¢ÍfÞCi!´ÚaÞ×VûG}YÌ8oÈ÷œV“ƒÓØž'"ÑÜŒNuÖžñDrtö_—šÓˆh±Gk>C^o°R™É Ì‚h”’¼.eµKÂc’¥zŽ¼:#É6†ç1È³r¿øNŠX
|<a¨éð,#¤fnÚüWp5Sóˆàgc}­“4
räÝ3j~ ¡‘~QZ—·(ÇHzjæ»•{ÌÓ°J¤¯žDãÊ¡½	Fq’C{¾ßê×°ŽõrCºÝl“ÆIðkÓ¯±Á}uWi‡0M¿!p‡Ô`“î¯È½<”ªÓÈwÚ6•6ôjó•Ò›O£uµÕƒõÊ‰õÑhò8Î‚î?¾/óVÕúñh ——„Rqoböõ Áno ä)Åö­|ùfêÕC­•Èú8Ö!˜T@b8xs„‘&)HÀ8We"r·[mÊÆQ~²óÒÜŽÇaÒðÑ
g¥T@Ó 8è¼Š¬Ëÿ`è¼=¥çJâ¶ô½AG@ÞO`®˜Ž¯áªq‹~þYƒDWÈÜwd”¸ÕÆ°´Ñs4qêãÙgxÚö‡çñ¤ÜÁ¿£=ôÂ1é‡#”ìõ­£‹RÛþ9œ|7Ç¿sx1‰²R7¯%*ûœlÉuò*8dÒÍšë+í,~Êÿí EVá÷%IBÛ8‰3ln¨¾	tå‰&ýhŒ”óë AÔòM“NÑdCÙf¡jÍ¥‚o~jJíUBÃ @ã,+Ti‰[ÄØ¢ýCMšu°ì‡ÑÑ4Â&+ÓÙ+?oy‘T¹ƒ¶öÊ¤¨ÅJe
õvÝj}ÍÊõ†Ë5é1§ðŒCñƒ»þNNô/@‰.ßÒr°Ö‚äuH•
–*!ì”£\E`dù»yWj$p‹a¥ï(US€Vå:±—Ø‘/Š¿Ä¿œ‘f8’2Ì«ä¤Äè®20yÉx*`)QÌ¡ÄŠ4£QŠ’7ÆGb‰ß©wZŒž 5Ødd€)j™”dçÁA˜â¢H/Ì¾aÌÏ·ÂóVbäî7©ðßI[FZ+Øg;Êßk7½ýÞµÑD·	ÍVÄÎH#ÛAµl©øW¾ZFÃEàà£)â™Æ“ø­²"žÅ}o˜°-ñ’P·-@®Ïâ·¹¶;Žfƒ0mB»ZC¹=Ø™¸w$¡ÔÚ}åÜ‹´t®0Ö®(Þ ’ýó­êÐiòŒ?+¾Öì.½?Ò²bÁ4@é²ÈãR;Í8ìÈ§(iLãÈKëÈòã`%Ë_¨(>óÕŠlRò•(=ê9
ÍOY££¨`ä©½”•QzÄsIñÕíWûžíöÚ€¿_mìïmïuŸ	Ð2UîuŸu_íwö¶_<ëŠË¹ý@}ÞŸá³Ðü|"ã3]N”12DÏÔ5	*q®=ÿ^è‡t¿qéKßŒê˜øOô7UH(­„¤Éüþ‡‚„|-#>@J
-Ùi¯´c×W…¬b!.+äâê¿21¾[æv9´Ê8µºi«¿SºåëU¾“,Ue3XÍïðl<°rfŒÃc5®Äµ‘k(€;<d·JàÂí	h…pÂ5‚ô™DþEÇO™I§\ÉN¸lzlp”ÒÊ/Ê(´qÐVøÄ½By#©ŒÊš,!µ&‰²p‡„ªb®Ô>x¼È¼jÚí¶2@ÁHû÷R0Jâç¼ÿr {ÉÆv÷ÙîóîÑÞÁ«îöÁnoµØ…Ui*ÆuÔùçsÈ÷07÷à ×¶¸iEJÕkwµxf\šî¬Üyž‘”w˜ð+ÃU£Ÿã…“«ì	p—¾Pì^ô©€ß´B%³PY¹JÔÔ/)'šÝw"ˆyq®‘ŠQlp«ÉQQ÷8…C*F5¨\]T9Üû›òktQü¶¡Y
‘I»"`6¾£š*Ø¼®`„a&¯ï^&L‡uEhÉ‰¹'IœÐ£Ä-I½Rvû¿âr¿*+Õ?æqÄºEMÁÏIµvà°Öðï ÏÐj:ñ†U|‚žÑyE¸iü¥‚JªæÑÝZ©*î4Ýbºª¼žNâý$×GWûdBT¹¸ù5£µKW¯¸ae¶òmœü˜ž‡TÞøÛ³ÞßÚ³:?¤ñäU¿¢?5ex]ùºÒþ4ŽT›ã7¯@|)‚•¦SÀ­¼ÿ|=¢·UåZIÞKíkg.Ýªc¾M¢,.,´òZtùªèïÕÝËü-õ5{½ƒ¥9ð)Ž"Œ‘ÛX9YyÕ¾¥¯+Ãçxæý¯ÿèaï˜x°TDdèb9XŒw0÷j:œwq>íŠ;~ÜåÚ—ä;`[7ÿrM:Ø-+ÆƒÂ²þVý•²­qÙø¢ì7üÍùfÅôdS5uQü)íáõ0$·v¾éiø"™)c×Í`ÌšhÀ•faÒÑ‹rÀð)%8½GU‹Q®ÌG±BS*6*%û‰8ÉF*^¶àeË´ËÇaŒjñåŒö˜m­l´aÌ¥-ö8GÃ}½E°¡Ÿ¨T¾å˜Jö*óJf¤±+.‹t7éö”Í5Ê¦ªc$ÞŠBÄZ#ßî‘í ¤–KR‰ï˜Ž•øŽ£3éã}íµPCF>(ÌŠLFùzÌÜ/¡QÙ !»G
é’ƒ'"YºHÎwS²;
Ï˜òLó²¹ŒÅ6ýÆDç£ÙŽ~-÷^Ù¹ÛnÅøç’‹Êm*ü”}B4>~Î o¹Ê®ºfáe‹ Õ%÷Ö B0>…1ÔÆ	‰.>o	4°e=Ô@ÃÑÝ ¹>=‰.ÂAss¥„Ô­Ù DxµQ‰Éò±·@óþÿ  ÿÿ l˜.xœìZÝn7¾ÏS°ƒ Zd)ñ&ue®³ÙíÅ6†“XBÍÐ›™á„äØR=L±Å^ôrŸÀ/¶‡äüÿK²{Q¢h<#žÃóÇÃïœáü	Ê÷Lb]ìÑ_±ËÐ]â–›2‹óü—Þ"ÇÃBü€}rf]{d…¨$¾°XCì•=µæ…åfŠYžX’•´Ÿ¯<tÍiÃÇæ¹H¿'>áØsí¿—ø tõm|"9uÄH*%Ràñ]w0Ž8	=d`¬#dYÃmQœªrcÐ.{?fÏ%ÝCû-nì»%èŒ»„ÇÿØÂÃ’Ø“ãcÄY¸ÄÕjŠ%Hxg?6Ñºd¢z­„±OfÃý¸GaH¸†G`;ç@ªåð˜ó)£+Yð"’!— ×ôþ7N±0¡ðây£ÞË>~×Š<‚Ïë`3þ½¡žäL /Ç	ß²Yv
AÕ x‘Ó»Þr0¹6$Àf±ç›£`ÒD„K">GTàœ!Š–*ëyÃ©‹Ôÿl‡yÂž ß=Í_ ©ràd¦·¤+ïæÎ²äù^æE`a\CNd¡=O‘ºÆký"gšN%ž‚þJÎ6“WÅ—Ä¡AÉ¥\‡±+Ê<uô-Á-„ŸYßEÂÁ…Œ£KÎ|&?Bç÷¿ßÿ‹!¡wÌ£•8d4•9Ýb/ñ„¶Æ{Âými.–8¸92Dgs$ˆ|—Î‘Äü†È‘f4,Sçzg_G£ñ‰BÎ!3trKRMm8OsáÑœ…¦‰Ót#˜G®Bpš0<ží6*)l‚Äk2
‹¤Ê,úÁjöS9U›h+'!â§ìÊØÌÄ9XíÂûXú¯y‚’±$T¬ûØ|u˜µ“³AŸÖ%i ƒucÁGsr".·CF!šyÑSå,>c¡b›Þ¾Öü=s•ÑŠí*ÐàoW“€$W—/‡³±!éà¤(¬yÌá©Çž”°HŽžêé )êzÐÐËHô¼w¹ùÞœ€ðŸ1«>ñ˜*¬gð\0¬nÍS”¼é}€]œ1HÞôf`àoCüêÉœÜÀBý#ù*7óêÏ.åNeÏûÿÄç2Âi4…Ð&
èçˆˆ‘ÃÁÀ8P»gPšž-ý‰¬Wo‹žß–×ÐëÄ¿Uù56¬bö¦P­}¬¢ö÷xX³°µÂgðÝ~íÁßKêÂîGš‡Áôã>ör@ws­ƒž¸WìNŒ<ÜÈ%:;;CÇèÛ‚é«5ç$†µŽJ0<-ž÷*¹R,lxx7¦\¨Z?`ùH¹N@‰ L&¦CŸ#H-×ÔQåvGuÕSØŠ¦lññÊ¾³}ù+G’U» \ÂÚ¸°o•Xð—ÊÞaŒ¯!ø‰
`ÀÞVÉ=P»ö‚Þ¬ÄðÒÃÖ/Æådàš)r˜bNjÔv!
…z©~	®)¸´Âm`]ÄD$²†]Ö)Eñ¶ú>°ZÛÌ$^x¤&OÆáu™‡CUjÏCí”Ø¾bä™\ì–¶D’_Ç/“k/ª[C³ü09W?×—¨/ZºA”(¼.©€„…]±‚B:Ô•sâfc¹Ü0©Ðv§|º’Ý)áH¼…¤DW»’æS€5Oº8‡qù;ã»³°æWãà ¹ÿí–xhðæõOÃ}xÀÁ¥Žª]	Ïã#nWº×Xîk.No–Òšÿ„=(ë_Ó#-Ç¥ [jwK]óÓÓÂ+·íÒÅWí-¿ÿÜ!1…|@:ÿ èòAƒ"1îNOæÞÚÀÛJnP3!uUÚ‚¹ë¼¼Z!ãØkÿQÊ`©õ[|âÒÈ¯Q xŽÅîŽuWEmj´ÃCHä¢3;‘HÃoç­ƒÂTwÐø¤óÂsÛ|ä(fUht éþ¡Þ†è+4ŽBì‚É¹Lulµ°rñºÄÈ»	Øt,€ÑÇ§¥ÌÖ~ºÑBª?àÇíÇfj*ÞÂ™çÂi|V;`ë‘[ùõÑ,[ëÙ³ÆÙ¦ØNƒMc®¬€ûæI!'2âA VÎ§Ú÷ÈÀdÐY­Št¿w•ÊêÙ Â¤úDËöc Ç–HƒvãõB¨a-<8ÍÔ'…ÎÆ…uÔÈ!3ö³gHøœ	ÍgzÒHU)ÓÌ¨ßéŒÇt—ÇLºõ©¥ésFX¬Æ_{›Sµ²*.aíIŠù5¤Ê£œ0Óbh5TÄ=GƒºwÕJ‚ 4Pn	#z¹\?¿¬6¸ËãTÓc-%7ŠºMâG«ÑìUüez6s¨«ò¿Jw^·1õQ	rÛ;¤ê: ¯OÞT¯d+žw	Ÿ€Ç½…‡\Z-“$½\ÂJ{‹š¢ÕvYEi×™*_·J$¼×ƒ&S 9K*=Õ…È«JWM2ö5V3‡NS¥ßU÷ðj@õÉxZh—bµÞÞkð ´Ö\‹²T¬¶ËÎ%sŸ¿öX8^LdL¶+ÆeÆ^«µ8%¢©":Î÷%ker¢øn1µ¬UŠÐ*-«Æý‡ÞêMÝÖÊ–?é½å%Ù¹ƒÙžû½‹¼{³«bò üm¡Ö|£Ð\Ù“ÑIkÕq²'­$u_jì‘@òïØ—l$BÊe[êjÀKõÃèFƒ5._Î¨¬7î\P‰Ô!Ñ€OÓŸüðµòD;„S#ÅÌ35z3U¶Õ¹Š¿·©qŠJÓNé:¬ÛåI,ïÿ¹*¬Çõeç„M¡>iwdW F¹%ºÔW6#@½XBùûÉîÄÖ&A)Éïÿ`.ƒâ™ï/èšñå´v]™{„ËÊ$wWN¶åû*ªÖS@Ïµ:¥3/¦-²½q…é’ei²RwÄTÔÞªÕañn‡í [®ýö°ÚewáòÍÁba’Ü´éaGÉ©Å|;z…e‰¾Þ}ËÝ®|î XÍu€ƒáñû/àGz½.Ÿóå‹–ÄvRü¤>~uŒBý²Ðiú<zg_SÙ%6]'ñ‡¬+žLoHè¿“Ë?«ÞºSØÊlh³B6¶dó#é<¾Óå(Q}Ú¦Æ¡u­ËüðˆD\(4ô0ó#­jûM_?ªOù‹õmÌü ×hi8ì´‰¡ ³èVUæ€aà’ªÿU·þÚx¹»ßFwó=>ýÖnuÙ!¢ï¼Ôi8=–Ð¦‘vr˜È$[
nömk¬Ù%»‰V¾
`:Ø¥*¹ÔÖž«¾v·Äýv D²û6ðÖ
Q~±Û†È©Š,Ä qR¹VGu|›'`º•Åîˆ{¸´­Ÿ òc¶ˆ¤d]ÐWýyB´´²Á@Sê|:ÛTís_lÛ‹$
¡Î —x­ò«9}»õ5#ÿ‰¥»fËhrŸaú’	ÙofŸ$ÓÇö.êZ‡{¶É,Ú‡.Wõ&‘ÚMÖÇ·›ê)ÐË PýªÃ³ßäSTÉÚ=ýóíî$§9xÑm£ÙØl±.[Õ~ÍÆ¶µ^~Øbx—I[O	e	ÊJfÕ¼ëøP–—¬ºÉÐKFjžÓCËp]*ãÓ¤‰	˜Îñ¢û?\|”»Ç†i Uz‡Ë_O·yË¦*t•ìqÿb~¶ú
FW	ß#b«ò~©½›UùÕhbu%£|)Yßak»"—ÞÙÌÞ§ÂÂÛ'ÿ  ÿÿ Ar(Ã