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
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Produto *
                  </label>
                  <select
                    value={editFormData.produto}
              xœì½ËrI’ xçWb³³#2 ÀLfW	b@ ÌBI`	0kwQÒ#ÂxÒ#<ÊÝBCd¯3‡9ÌœWfk{EZºEJdDZæÒÇæŸìÌ'¬ª=Õþ Ébv¥K&nnO555U55UÆðÉ¦ÛgÑô4Þ¸îÆ=¶ñô>E\îŽ“òy–Ov¢2ê^Wäc¬ßïÇ$çreÆYžçe6`q¿ŒòÓ¸ìŸGé<fQÁ¢éUU¹›^ðÃM0u”FEñ*šÄ‹•“yš²ÙåÊ·lvµò°ÿˆåÙ|:ŽÇ+—)fù8ÎåŸ•"ÊxåáÚ+ãËr¥˜°“l4/Ù¼L“i¼2Í¦±L’%†Ðñ•GP`xºrq–”q'ÐŸ0l¯çÓä÷óø@@£èO¢Y·;Ã©`Ý
<Éfe’MÙ‡øjãzvÃ8ÜðWÕìA+³0„ ²UQ[¸l¯*÷dµˆÓxTúež¬Ž“ó§ütž¨)†qJ'j˜f£ð—@yZ®³t,§‚Ïz2\Yï„û|Oâ¢ÌcöU°ë¼Å`_’él^«,¯fÐ5ìBhbËãßÏ“<?Êé¡«¢_È>†gåÏ·(U¿ÜUùó\Ž¡ÉZý<‘ö Î?þ)gŸeCX9‹óúð¹!¥ìÖ/8É>-N¾ŒK {š&ÑŸ-'¦Ÿj’®ý‚žìÓ¢'ÎNÄÆ1Û›žæqQüyiç	öFõäsCS«s¿ *û´ˆúêßþ•½ŒJØàGóôÏKJ§óIœgØ™ûò¹á©Ó½_0•}bÑ©ŒÊyÁÆÛŽ¦ãd•ÙÂR”ÛKB¼ÉÏE¯þÒuJÑ æ­sCŸ¦PÁSõ«^yàßƒ¸Eã¬ó”¼,TÉv6=ó2Á:Ìï…ªxõñ2€Þ„%²0 ?m¡*¡d¬`>…n‘S‰?Œ×y±Ã(>;©®˜`¢“z½úg¢Øp^–Ù´`_­º¸ŠÄ‡¢êI_²çE™œ\­ °Óh¶ò›•€ºßJU× z)šŒLìjâsS&¤Éè®#¯u {Ò„—Ù8J÷gñ´{¥EÜ{\•3'ÓSEöâî–d0ûMh=éeLV(Pâ"ž$.5þN¬J(v–Bt.sB	Ê<š	¢ÄÊdŽ¼ðáÂÓ(N£<€*Ì‹OO1N’ Û1NŠh˜Æãë4‹Æ Ñû€'[)NË¾£€â)cÃ©8ªr±2ûc¡¶I±2B
–s^’«z>ÈfÑ()¯€f¶ƒ¯)ÛdX¸ç€Bl46ï9ÛJ¡¥èã?~ü°Ò}ˆTMD`-?YE^ž¦9™¬×ÞZóO~{•ÅÖlö‰îi|ÈÉGf")p™l\ò¾™žâªËŠX­:X0‡$Ÿ\\&{Gc>¿²ºx¼;-ó«Íþ4›Äìïÿžu:væíy^d^î¦†²kÑoÀæ ÝÂ1úbãúáixý“'©ÛðNúe2ËØÆÆë£¦­G¡!Fà˜§JNX×êVÏ!S	 ÈE?CÐs¡¦æMžÚåú%¼Ázˆ—¶³Ì:ï†i4ýÐ±“é!MÃ¬„å$9Ý¸Ö?íLÇÏ29‚e6ÂÍgTâœ->]›{`Q1'×VžGWý¤à@7q¼Çkoa	a÷œ¤³8¿v+ˆav&`­>]x‘L€~¶\(/¼Õ‹¥£™nBÙ:Œ½-¶Å“nO…‡)01E‰l”øÎ6ü¹%…ù$ëw˜êcúþö±W±šè °ýáX?¨„5ëÝ¡œ ƒ	dM¦£t>Ž‹î°ŸŒ{ìË/ÙR2¦,?ƒªöÆEÿ,’)õüÞ¨‰9ˆ®p€Î¸ýgyé2KÆ—fÅ†”×€;†ì¯Ì{?§§åÙ[›sžO=H-³K+—œìû ê€F±D©¢<Á=°èª¦—YÚsã
ÃÊý•vc­4fÙT¬A'ØY9«—Œ“ëÐžÚîñ[’ƒ 7l˜£ZÝŸ¦W×e>+kšL`¼ÀûÇÀ.‘ýTÂF!&ç()u”$p|<~YœjŽóÆ½KöÜìY€™„õ¼!×²UÙQ°åZ$:’§POåÆõ±5G.º|ˆ¯€O9ÆÉ|Ûq§šKùðùnÛãŒ!f{™fy|žÄ? dy_bËì0›ÿ!r²ÞØ¯×ªm¾ÍCãº9ÎÀ»SóÖx’L“¢äœÔ?d¯ÂðÐæS ©ãšÑ½š&³±màó(m\¸/@Ü£w#!éUwu½™Ælå”^¿Ì^d#ñõ–Âô´ÛìÙëN¯ü‹h>ŽF‘5‡˜&ÀìÍÃ~úñðß$Ôª­'«ÁE¥TX­ÀžžÌ§#.#ý-^ÇÓìûò´Ç9ý±)‰íøs–g@ò9ñq7ELÓ[˜(&‰Žx!ThùÁÍ€é&¼œ‡8æ„V¶7`ÝÉ€®Ë¬Ü„ñóÑšë0à-ã<Ïò_ÕçY2~lz7`oŠ8?/]0‡ÝíêÞCÔÏÇt Ð£çHõ	vMóBx£M•jõÎÁ@Ú˜Lt—®Éþ!J?6»…J@b¡kg7Ço¶¤•Á!Šî˜š4–Ú’ˆë6`·‹Š«éˆue‡\±Æ²<9'VÇÅJ‚iÑk*ºˆ’’EãñN6‚ò)n€AÝñp™mï¿x±»}´·ÿê°¸õb÷ðÝöþ«£­í£C`ŸÍ¢”­îê'ìø¡Å§¾Øúîï²œ¢™üUi¼8Ïj}áB
×@%S~R3ÎÈê !uç1¬©ñ€ðÄà£m6¢É¬«×»Ù{%w}ìý§œ˜]eX§CK°Þ5†ËjnØ(*Gg¬ØÎ'ÊH 8ÇY÷ù:Àï~Ë»ð‰E™n/g^_úØ¨\J²É’Ðxt|¢´¶¸*î[À¤y£â:–Š_œ.\Ë°R%Åtª73*%žÀ®i^G³ò"¦R½j«µÎ÷ ¹¹Ú‰¢ÂI€=úøO£i2/ÿ¥X9%9dÏ¸¸º?ËòRnS¦)c‹£W†*Á:~V‰Îa_ÇjŠœ­ˆU$ˆ¥&ƒCø…ømØða²@Gùèì(Î'Ë‚%S¯V±N‡–Arûœsì¼Ì3ýJË<àøíSÅ×©¹fž”?$	5­Ê‰#hJMIŽ¤Ü¶y¯)¥föÔJª-kxßeWÙÒ4!D^vÙä+ÄzÃEuçS£ô”µ¹rîrˆ;¯ïÇ/ÙßúJ€ÛÌzñ¿Ò¯öòg“l§¸÷qìæÊC¾8–ëê„¸s‰Ï"¥Ü(HÉê*û¦,Ë™bþD.ÀLþU#r»ïç@…ôméb‰Ù—T~¹Þ‘øõË”w2Pm‹ù8|½‘ôò"Ç š"µŽ±q€|“ù9¿xçÛ0ê®:±2r±YÑvFI—­€lCVa–µW90ÖøøÌ'"+:ªYDn°%ºÞ¬
ùj§µÈŽÕXË_Ô3S_ÌÍç@M|Íë¡.’ êâ	6ÀÌPIîf˜jëUÝ˜MI =ù@¨>±/¿4ºZÕ»B§K²OÓ”NÅÆ(ÝÅ7QN¡›¬æÄMö*‘]å©¼kÖJ¶ÓjeÓIžMº(ÆeW,m®‡ëš ]¯×/`ßíBcÛbC•ÇÖ®¾L6õåÀžþÖ­‰ª}×8¢T÷KÃ©PìîÕµªéî)÷,RM{œ]49û!?ã1Œ:ÐÔ×1ò¬ÈzížƒLeÜGÒvâ“hžªúT#lo{vµ).®\Û´^ûRÕ]ýÝÎ*0!°)âJç±[Ó‘äàhuŠ«Tbu%)væ³4	•› W(ÑÕ«Ê²4èê®ù%ÎŽè…Jí™åy—hf»§¼t©û­ªPß%ç®ºŠé¨Ù4SýZ0ÓûñãY|	¿c6øxX„‚äJ¢ÇÁóUÝˆÏf«…®˜nÉx*æ¯‹Ú2™õ6Ø³­CÀ^ï¾Úÿak{kßÁúýþ‰g‡¡ì.Œ­Á©"JÏ3$Øls®làKwÉ „Ä’Ä*Tªâ|yØ2x‚-hà6Hó5ÒƒŠ+AðZ])Ÿ $O…4aš÷$
:5’£š øÒ—Ovk¥©A<‚šT',¥Å0œÄ…ÇCZ´áÂgóôÃ”æ$C’CBüÜ²DkìÔ³æ–£‘ç€#Ôñäs~ñ~'.â#X×°c'9ûâ:\ý’©aá9`Â¢/6ß«yü:]Ä|n¹Éâ˜e'îi‰}Ð'ý˜Ã×ýÿ¯_ðxôÑÓ+‹ž¾©Ùßfly<ÉÎXÿ=©¬þ óº	iÝfˆTÞq­ÌNOSyTèÑMÆDw>Ä0”[®4òs¢·7$ “6'ã·z‰Ä°ê*pÁ«äq|Wp‘ÃÌRÝè¶Òxƒ²$0›[!]išžn·×’ƒá›‰0!$|J¼DÛ»¬7JÇùƒZc?o‹Ÿ%äÍÕìX’s†¢4ù,è¥Š¸6ä4%AØ„(§`dpØ™\‡buB#¢Ö‹­v³Ó«"@÷@Rü}^Q‹~çžˆA{Z  º{‰;2.=‘Ïk­ÂÍq¼âœÅÐÒ=iÞbèéœÏHþ”›¼ètÁUHÙB¥(ÆÂHÁX*ÄetÐz²Þ-ó°›=ÔœR­Šªkgö÷—”!Ò;@v¾Žw„wÐb[œ¶=>c6¤
¾ÏHŽz’n(•„H¿QØs48Êv/GqÚsåU‡Ïø;}Ì%ÊG¡½‰B!-v	:Æ¯'¿>zùb/ì¦ ÜiùÔC4Ô Aqm¿ŒïÅfÿxí­áB–0Ñ0'"7üäJÑm~Î¦ÈŽÁ2d°W³hZû!âEž]ÈcØÃ>ÄWšÄÝ›Ž@‰ß@>(¼?üH=/…õXgÿ„mïÈ·`6×*IÂÍ¡N¨RV ˜Ž]K¥îÞ£¶2„³oP·Zc]$àªÛ¡=Vïoù^‹fŠÐÜ!¥’ÅÊ6VÒ
]š0U8S0Eä‡­#6
°îy”hºéôdÐ=²(oú’"0 R(Ò‡'òZCè±S7ÏoôO¡3Ám¡©½Pé2FLEoæ}ê¨ÚšÕŒÖÞ4µ1ËŠp³ÿRTÕîéƒûó k5s„ü->{÷5eB6¦îÝ­0ôDæ™8ëO¢””90‰5EÑæ7XüÐþPSZ÷’‚Ïñµ&ûY2Ìñ&žh™üñO"ÝÔôk•TSÛ89MJ«ã;2åx öÉ-ˆZ‡r¯è0Sµ64ä›¨½Õ‚ëK’ê¶Ö©)MNqëdñ	T£ÕÁòòµØ›¬p,Ò¡<¼Ø_èšwUÍõ¸M®›„‹ÎQE]ðZ‰¶!¨@©þè,Ê·ÊîNç›ÙLod_ó¯EšŒâîz¯ó†œQÞàœªàLÑZÞfNñª/io$ö|žsYX‹!ËŒYítzÍèæiteÁj»‹U©¬¦:¥æµ`%t¾VUXÁ\Ò®‚6Ü4¶
ü^ÙªKðÙáÚ¤9[‡gªƒˆVø¹;¸=Ì`+2'¶#«Ñ-UB ¤@”•[Ê–¦‰7ì¤ùâ‚—ÃoÔÏ“UíÙÌ9O
€Òè>ƒÐR
·vQ¿ëfF‰+5àQ!²ò‰°~r–ý·²!¶ JÜØRe3j8zßæaûr$Zµ`B¢~÷\:ì-Õ P<ó= +‘+PU·çg˜Ô
ÖÛÌ‰C3ì1Ÿ¢¸ü·hCH‘xôÚu4¯)Öx8'2óe%~5bZ›s)Vë·4.¥¨£åýÚcëKñ!½Ïù õÜSh
mÏNPÌTÇ¦½ª|°	9ùtF"†r=
¢|÷ì$[~šÇO@-
áˆŸZºädÒœËÓG]*°ˆè ù]“9p‰wÆìVi/›ûfmÚåkz‰à%½Ô%(œ%k«!ÈïlÈ~»Bÿí&yÂ<µZhƒ|«*g@`ØÖœ$nýµ¥°5ââ‘Èîä|àÿRzN’íý×ª%zÞ’(ºQX–¡x%WÖS¶Æ6±¸L¸aÉ)Ð^Ê²1?ºæÛwÿ=?Ä¿yoÖò-õ²²[P;¶S£•­PªmÇÛó¢Ì&ÒÞú¶+eZŸ[›IzjùÝßúo½ut¸upðîåîááÖ÷»¶U3^ÈÛ_ïr­c¤§ÓÜ Îí´{ M¾k‹:¤âZùtçCù—ñg`âÃó3úÈ>Œ7Æ~M~ºHÕåšÆCÛšY#^!nñNT§IÐ‹ÕóûŠ²XDäˆVÅ}yKÅ1'3SµšY]9_S<”ÛÚ“ iÿ†wiÁ»kWuû0€Hœ>Ê6ÔIùS<'7¨ì( –ªŽÊñéü¶™(@OµúLžU$ù9yŽFÛ&÷,çx=o;L`æŠ"™lÚ¼’¡_	¦+]MŒøku,æ„[•é¼ü.·¯–aõÁÂ
úûyD7b†…dŒ?bS\#'Y"6Žúì^â4Ê£i™ä¼ØytÁxòˆe Q~ü—Ë¤„:¦Î~¢Û‡E&œæOÇ‰¸¬.ÎÂŠä<*DX?]°gÙñ’ìN<¹ÕÄVOEVÑÒÛ‡¶¢?d¢>Zô€Ø5€—è¿’Ó(Ò[Ò•¿Õ<áÄ8E@NL›ÎÕýPÙtVRÚÛp¹öVæ“LŠªœ§€MÖÞ¦8,6[+¯)¬‡Ê«Œ‰@¥[é|Ê-í£M±¸ë¼o9)]1¦‹ÆÔ—z~v‹nÝÓ({Mª§ÝZ1óŽ[Ö­®“bâ•«•_i¿%A/(øz•`“ñ€ÿÎ³íe—q<µ=ML¢Ë•tP1¹\‰ðX]O|K¼£<9û†6#<9¥aON»¶Æ¸_ÑÝiv‘G3ËÓ
W¿øø§ßÏEL«gß+G†ÕéáW­6BŽB-Áý˜³A¶#êÈêt¨U¥ã•GÒÑ‘zGçÆgqúá;ö€¹„lÚ[ˆª}­ß °ÝÚîGØôÊ­ãƒã H"+@2Þ¸^ÿÕëÌë	4<}º'W”^¶OVy:^ÀíG-@Ûlñu@U>À(}+ úÍí j\×„Ày¼¾6»|)RÑ2ån¡@±Izjf@9høý<Êc5ß:S„¯ï5íT¹ÈÖ>fnf¼Š68&ÔÆùÊ,ªg”ofÜ³@¦I±ÈG°
÷Â‡žì»n¢Ñ(ž•þeZ\.3üãf1>Ù¨¥‚ëÓ†À÷,cÇ•“3Mž3¹«@Ù4cûúÚšëréöK¼¯ÝžzìdÓ6ó*FšW±-¯?¶·/—ÄŸÇÀðÜ¼
ì_æõ[SrÊ+7rl¶òà7Æí@!ä²çÁÞ7ÍÅ‚=‡§ømï7Ngüå8;œÏ8û;!ÌØè+	vn2t¶2ÆŽc@qÜ9-p;n˜‚½øFïÔi¤Ü0Š~üÊsvM˜“›ê–Ü×»ÎÝþ>Ù\Ð{‘ÙÐ¼Ü½Ì„êÅßøs¡YÃO8YŽ4ôÓN„lsÑyà2ØýN†ìI`.(ÿ	§#ÏŠ;Lˆ°Î“ó˜á>q’BrÓ«ƒN4,²tDÙ«\áž y?€!.èž6‚Y·þÎD]N¢†ðæÅÆ¾fvk·umQƒwÁG›dJõ³%ðØÚÍ­zqY–qË0@™”O¹Ï£¤öëóù¹]t>3èr·kuz•J©=öDÓçÉe<î®ûþd½Q2öWB‘á‰qTºUTÏ4¸¢ëOÂØËeAY0ó[—reçH¹¤Yx<¡§ T¨DïxhîVøHÅoŠ·›¢Àm7Ètsã‘3èxœoHJ‚Ñ`ÝÝËí{¸öðÛþÃžËyW™‘N•õEHuìâI£3eëâ›u=œE;‘=—Ç5‚€ëæ¸“cáÉ™!?yÑN©ßã£ìò½5s˜MßUheL Ð„ù>A[ø¾n˜°æÉ
ù¦þóMXÀ+zèk3±ÛÏ€:„½û,¾WÜt¥1„ãö2åÁ¼tíÙóÜðáós›Ñ{ò=®K´âkœûÑìä¾·+3	]pçñg7i-‰à¿ýw<q#ÖTw ˆMQ*îHë£Lüì&èÓ®*m9yûù­Ž~w·‰­Š^÷³›Ñê0.(¥©ï}C²>8ççO¸ÆÙ\0Æf8¾fu?®f8¢ÃâKÛXÛ,jj)|ûu]¨ínYhíóÇÆôºQd—n'òS¦û”åm‡Ÿ¯0ŒoQÛ¢1®…?6rÞl…¬ ªhÞÂh6 ßTÄýÀ³¦Æ@žê¬&@Å–¶†úø[Vâ®2-p”ê„¢X@¤uD‹U«WÝP÷¤â5CQûømiy]{ŸmZPOåPgáº,bGy4ŒÒ3˜†®²ØèY vTTõf¾FµÖØß.„“·:ªU× ¶–€«e6[Y_}ÈV8¦ò¡^ñÛÎÀßYä	èwîbj+Ô…À³¨ç³y1à¢i2ªFûë8KWþx E›¨éDq¢ëUª•1¹jˆ×ÅÊ·¿ªWWÆæÕvÄêŠŸ íM/çi™†øQÁÞ×ÄŸÛ€ò?Ã½‡CVã²zG²ÜÂºÙlíO†æoðß3×M>>@._ mA}>ïô=4T;fÑ7Ÿ1Ü1)N”z¬lÚÛ=w·õXá­>;xYç)@3ƒaÊ‘He¬²…¸î…xîÇ]Õ—Û®àµóNœ‰¶™uâ¬ø¾æ\8µl˜q‘IÌ÷¨Í|ô|Âó=úËœoÛYk›)·ÝLß×¬kç£¯ó‰¹/ÚÌ}¡ç¾Ï} >è_ÄÜ÷Æ­fžä¿—yVúÖáAïôN!Aï5èŸ/h­3’ÑÂ”w¢yK”02-Çnäú]é¢<‹£±ž2oaüHÉÁVÐâ*”P0ÑòÌYºßáÒýXêõ‡¶ÕÔÿòdµ<[°†`¬äÆ¸ÜÜâ0»Çæ–N7Â†ÌuÄIÝ‰†.X¸7tÍÌ²MÁ§]f×'¤¡4Êõc¨ž`ÄëEæÖ%Hn‹–Áq/ZFQ½…0%XÔ?bÁ§Âå)7kúòË >ºJg¥Ç:Úx¬Ï)ótÚöË,A¡±Â-Ç­,ŒÉa}8ÓN©_Úð¸*£Ðòø¥~bôtü¤þ<OïÎê0Ž+,æq¡«/TB±hõ¡¹‘ ‹ô´*ríb³òäYVZ×%Øî„áÐ¢û˜›Ð’´ÜÝº›Å“r˜¯(Ü`»z¿rÅäêpÒ×wš±píš ô.«XÐÒ‰ÌCªü¶6/vî<Z[}äÞhBøI9SjeT)~k6Ï ©èp°¯Ùzøˆªô¶íºU4R³¶ÛÉ–X²S{Ó­Zuµ;•.¾ÌZmXÁ-ëþ Vw¥°¢C¬bíÔÌô‘È¶:ScôoÕtÂ2£Õ<]s‚ÇR-û_Åí ¯×âì¤bÇ¨½¼éæ,Ls¨ºù@ª@Ò–ŠÓj…ÄŒ¹ã@µêš]!Àí×úBÛXÀ4½¡—?-,”çá;ƒb6Ïgi0tž[€Cvô'F@åßøiQ*¨/ÈÙWñª¶,C¢B–Ì·é@®0³TÃHŒäRAÛ­§r|×ôŠWq654sœ×£i52wR®SÞÒ¿-õ•Ôwñcõ°+>LâÊ¾ÖUË§¼O5—5j*’š›4·t)‡¾PƒDõÔÜžuùÜ¹·P«UsëúŠ“%í-Ô.UoÕ7(š$—Üœh5ÍV€ÙQ“µi^Ò[Ý¼|¯k¾‚FVr!·U€VT°¨"´¢šÅ¢•„£• ÷3Wî5uâ*TMCçêÛ‚»^[Zµ!-¢–¨ÒzÁR®º!÷vÁT»Q•ý¦­ÚÃ·r/ãzRí¤h‹*;Zy‰°JpJ…å®ã¨ê4#Îb0@¶f
,N°ýµ°"¶ZÝZU|¸r´P-€Ng+ÀYÃûZž± ?1IÁ«´ÂíÇëêhÐÎ2¿-³?/O3´qk=Ë’|Â)6‘všUAõì·ö2Ò*ÛõQ¡îSö®A%…O%°Žò¨8{hLµªàZ®0áôUƒ!Ãèëà†ˆ-”ÛžøµŠdß€ŸQ–l\?
ÏE€°Û§J®Ô.ÿ“QÊaÀ½Š§gó‰	›O1:¼ŽËžüýrF<$ú¾¼ŒÝr®{µOù`;ÛNE.µ•Ûˆ.›¸£(½žØ_¨WÒí0–hÚ­}Hx<4ÙÓ8)ÇŠt½ÉI¤ð4lgæ†(^n®z
e×·‚üB!‚LaéÐ±Ø¸®r™ØÂSbÂEÀ 3)ÂÑ¬ºžp¥ÃÄB¥š¤7yj—3‘¤¶³Ì:ï†i4ýÐ±vâOÖüÔ'-×ú§=€éøY&G°Œ^ÚÊhTâì->]›»Þ+ãäár“‚ÿ ÝÄñ¯½eüá Å™¶[¡A¹ÄÓ¬ÕÅ—‹{½nÉx‡S•Ë¦³[Q2†ý–l³Ÿwo‹7Gñd†¤µ¨Uªï:(2™eR˜O·~‡I?¦ïVØ%Q±šò ºýáâ-,¨<¹àa-_¢~{j¢+îjÃë8ÙJ1Þe€ÿ·áÕèŽ¡ û+ó.7Ñ·6/"ºB…¹œœV./Ù÷“A2$úüå	î#EW5½ÌRï"³¸«žVÕµª,˜eSó‚&8°³rV/'WM0Q‹ìdù(†V÷§éÕÆ5
JU¾5M&0^ùˆî¤6
õ¸ODJ	åòO\Ÿ‰7Öè]
Påò˜–Âµ»á­[«ZîtãZþ …£ó(Iq2õ¤n\[³å"Î‡øjÀ:Âë«§qá.ßLzÂÈwâåBÇÛI|ÁMÇ-(G!›ÿÁ©pc¿^«Æù>­ëöT,§æ­ñ$™ryDÜ~ñ*MF(©Þ›Š 'Nó2ÛF¹ Jîìy7²Tu‡ðþVCo0Öd‹eL¨ÈIø*£=tfåÊ³×^;øÑ|ŒXF§àÓd(
èÙþ[ƒ…«†®/Í¥ÂÂ½yðàd>åþ}ÑAÑÐwÏpÍ°v‰áøs–g<Ê!üt·?LÓ[”(&IŽx!4TÓ1 i¶5ÊøŽtüvùÁÍ€É†ØÎBsÚ+;1`Ý‰‰\nŠùùvØßkç»¸ÐÏ³„þ“]°7Eœˆ—Çþ˜Ãmˆvõì™úù˜ŽzTÆ©él•$¨±a6­hdªw\RB›è.]“-E”~l6•€TC×ÎnŽèw›´â 2o“”(Ã¡Þ6RW—b–í;24x·#8_Yf4;ö´¼È¶z«)a]ÌXön^Ôµelõ—ÛýºöÄÑ m¦Ô”„€“’/iJMIÀeRê{õVS")ôÎ´,ö+ýn•2;³,t¯B5Çå.I°b¸«µkf:OÓ§]ü—V%¯A.úõmß1Š<×á|x¸Ä›U¶3ŽŠ³aåc¾èSØŒ¢ÎS±vÈ·å’y”U‹@PÁÞXôo‡$XÆ¥ìàS Šgc÷,»ÐÅdoÄˆ½ä Ôè\SfKN7Mjõ«çYþ[IÔ<ðÓoe›FlÐíà¤ÌâQr’ŒHx(ÿ2žd]ª`—Œ±K\å^T#"A<ÖŸHÃ#­¼,~›”g]üôÎDúëY^v»Ñ2:Œ=
ÆQ J'íÝš	ª¸²þØ*>V'¥eŠh«ŸrV`;›Ì¢<–µwõ"æÊ2;vÁô¶ç©0Ö©Á€®2À1€^³ƒ«HUA"Þ;aÊ¿¨0!¥1p¨Ý*2¨ßolê6bÞª¢š“pÉc8­Ùº61Ð¡q7î¹×ç}›aMùHtŒô¢ÿžÔY/Ý-ˆºbû`êÔxƒŸš ôËÊXp ¬L–ÆÑÔÁe×Gc¨x#ë÷ûOÆj¼N¬¥@.°šŠ‚{6sV7º­”ÇË¾‡A)]9P?¡-œÙiŸt¤q`zÎeŸíÅëT™,ùBò;ËîÆ‹V{›°G‡[/vßmï¿:ÚÚ>²£qÈV÷ lò'šxêÏB¹ ¾Ø*Ã†ªï¾~CzºPø«ŽÂŠ”ÔúÂ•Âü0<™r`c*Ó
Ü9v’w^¦Ö˜:1Ð@– ±hì£PL#½¶óê³»Êp!Œ™žzlÒ>F,¦@tJLž
•h¦K¥pGkêE»,²‚‰KE"K(Ü©®KL´z5®¤TŠåF%:Q2U²çaLwR…U	(ÛoˆÉâý†îœ°-ð«ìú
eùø‰¢…’©ðIvdµÎëï-ñ€¶ü}œÌÒl‚]Ú¿dz<ÅÕÚÙ}µûR¦œÁÊŽ°¿†ß}5†4šf»Ó"™â—|Smñs+>ˆ¹*1¾[ÜÆ¿bhddhéSÅz‘38õE+N£òy2U_œ/(ÎìeíØU‚6êÕxÚ—Ù9“j=¶ÐþÚ1nîØ÷qî†•v2ïÿFÂ¡è=†Ó†éˆ0cY1]§yGÜ˜§ 
ùškx÷ÁÁ#|"ª-½ü
£ç—5CZ©}eÛË¦éeÚ¹e·õÅðd-Ïç³+.õò»zbOól>ã*c®0¨üv€hç@®64}"«zÀÖ}<úøO£)ÇQÔ9øø/ÅÊ)!ú£T(s\ÈòÝht¦vH; ÚFùVãDÒÆLŒàäÜîu€¼3¾ÞP<µ"æäûÃa¿ëÇrÛ%zbß>žF°UAwæÓò­}›ÑÝ…ïê/LBe¨G1¨¨—ŒF^Ã|ªÄ	µ^¯a®¥±Þ=MõþoÈt*[:’´u:YÝ‘Dàñ1ÍÙvºþô'%„íßtn~3JÖà\s“f¶?yµ0¹
›uÄš7€ÿEþÿ¾ñŸðÛõÊ²ý8(¸º
¸ƒ;-S§ "¬g©ñíçÀ¢‰‰´uÜà‰Kª„ätŠêœ,'C(¨!ãÊ™ÆÐ†ä)ãB(T9éDÎM—âÇq¾–OÆ®ÝÅ¬Us V3<Š2[ÒšY12xÝ45êá*¤F«±Ôµ¢*E÷7Ìçpw8ãoÈQãÚÔ€°GHzlJµ‚ÔñÒv-E°Ý²ü´YÕ¸UÔkÞj?)ú4äRczËq¿TYx[”YÚìªöz~=8 U—Œ+*ÓäÏ`°[Y?Kf4§¬¡ñUÕ	|ðHl]‡j÷‡êN.´Í´éÙ4Ú·J
96œði~€5-³!u}c¯ƒ³¨DWî"ÀT]-$õdtrHåÄv©Ë«Â`9·áu%ê\ÒØt×¼‰{É¤;‰ìÁ€-‰_á
¹nÇ‹¤”ä¡m~ª"WüœøJè‹mU3â!¨Hêo†@÷oÜ_¥„e}°¦dNy<D%sšDº¨`L“ìã)Õ¢yµ6b¡NAÅè£¨¾fÿ7†ëäLs@­ª-™Ÿ#(r:¦Z5Ž‰±8á[¯0¶wË ¥ñ…õ¸KUiì¢zÏ˜¾tU	…»b(où(F¡²
mÚ|6Žš5ÉËÐ¾iu ûÐ7í¨l˜ü(MþodVÏMËÍÑeÅp$Öj9CN1/Aæ™¿ÃPD>}|ü÷jšõ!r•*²=ð$vLs!ØÉ/@Ti,çû£©³ŽÔ]9o™8úfç4¥ã¢ û£eyBÒéUŠÜò˜C^ÇñQG}r~p§“‰ö'§üðÖ1©@Í=_ÅÆºÃ~åž^ì–Æ*)
£r$³»É^j‰BlhUÊ-«wÒGäç?'úÍféÕkaÈ‡ýÄžtÑÆ=€'z«¨6ü»,—­“7ZèÕŠ¹‘ YøŒì‹ú¦ÎU¨¤*@ZÍƒúßk¹Ð@vÍ‡âšàmù¥º–c«y{Î¾1)N¹ê1`T‹›Hð¼øÝzÇÚKxÕ™¬;I"gŸÇ÷‚ˆWz°õzëÙÇÿøêP†ggÿó¿ýçÿÄ¶N³<bçÙèã?³ÿÄ"4Îc»ž}”då¿ügTç?øøùyÃi´¦-?§g°¾Ødž”¬uÊ]DµÂ öüÐdŠa¹FÉ(å1¹x\w@qâ“i2J¢%hàÿý¯Ðëþçû¯ÿÛšf@i0¡¡Eœü#KDè[à"±ª2c4Ìùc–Oaf1˜§êvŸ¼ˆ™z‹½ßñîñœñZRß9Œg•üØ7FÉe6xðàòÄ+“(I‘”å›#®h"Uþ¥¢'8ØÜ0gXpz† (™Ä	 ~6#b¶A{¹¸ã!k0~òœ:?tåãŸN“RœŸb_·ž³¯Ù€ÿ#•+B		þ•9øU¨îã¿V˜³ˆOçÓq0c2U<Øf	Ìtz§$õe‚ÃÐðàÁAÙùeÏitá¬,gÅ`uµH"2~|…Ï«/²Ólú^ïtZ–ÀÅô®™…VŠÌüÙÔGÊ2 çÀøë aDâã?O bˆP0û°Ä`f‡ÿiZ,½W;±¢m[ã1Ã¤ *“Â·>Lã–ŠRŠù ¨8M¯pL°uÆ—MpyXï>ŽuF0/ ìËâ´ˆŠ±l²W¡Ö!òüÂBÝŒ™@Ó#ä:»v‚dW úºPåÁ)‚™±a<ÙDë`Œ¬‡f1GâhVJÈX/md‘°åñIŒtŒÞïÉŠ'Ît¡Ž/,5}¶‡R:ZEU˜‘àb›b£'YÂ î‚ãCZ€Ì¶‡‡+ÿy
ù²?ãŠ*,¾ü‹k
™³dÚEv¥ ê{‚k³ñ<ÂÅÊC+’RžX0¤¾çàyœøœÂ?0ˆ\Pàè`Æ o%‡‘GSXü‘`û¡» ÂqºÕ·qg.¸â¹$çÌuàäîšþ$³ü»éï¦¯þí_‰2b`&Ý‰åwj	¹ ¾:Dœ–)xYÝàÖöîÞÑ.ÛÙghkðzëhY +Aš}¨•ž”¦ç	‡Ò,ÂäHý”ƒ±g0¬nú°ux¸?€Š·^³}v°ÿúhëobëÕÎÞ´aèçH_Îé‹-‰RÓh„¦WÓ³xflEo<¼¥‡ÿöß‘“Þÿ_ßì²ÃÝ7|“À_¯~½,öÖá>üþ»-¶ÅÖ!ã»ÿÇ²Î½{ÿnï±—{W”èó*¿Á*_ìñ</5<ØîÎ›í­m`×aüËáî÷o`ø“ƒF¹Ë³ï½z³õÚž¥W`Q£º%¡Äxä6h*°{8@8Ð¸&ä¶ŸDó´ìÚÌìÅ>¡1¦!ßØ¦õÚ—ìkwõw;«§°¾:ü À&¬X“
6H«Sv	J¬®$…¶¢”Ê®£RäØ=NÆ\KOí_71‘(nº]=<Xo\Î‰¶JíQõ*d_¢ÙíñðòÚ–®D}ïõô@xa=_ZÕÍšgìG Û°cFÒTEš¢uÏç¤’ mWu“}b:#å3mgoÛŽã^êÄÀ×ê Úlq{æBY˜axŠ[O½ÊH˜ZSÏ“[èh}¡l sXKRÔ[oJd_¡2æDz~Œ^Â6#"r¶e:ÖÞ~ÖnZ«ÂUäôÝ¯­"Ú”'pâÕ±îº´1¸j„’‡Á@‚Áê"o¸¬;“h
Ðî>iL£èc+SRþþûh¶Ÿ"J’_Ùwc›æÃÊ|qŒP·#9Œy;Gö¥áöb1U=8q—È±te­m–ÂöÖ‹ÝW;[¯÷ößmmïïÚ}qî@3àˆ³âyœpô}é£ËÚm:±TeÔó ¶ƒœâ<Ê si^è,·›ÉWñÅO6“©ZL¤èˆ?‘~IºÒ<Ðä§˜GûFŠs?Ô¹tbî•Èð¿¶å"ŸvÏz‘w’Z0â¶b$ö¬5&Œ¼BÛŒŸ á">Æ‹¦™€#ï£kÄÈqÁ2dT)Æ˜QL–ÔK.®ô.0nYPsÙ*5kŠìÿtÃúðm´©6—[PÓé^Dðn¿ØØÞgyè•çæ>FÝ@‘YÃ^ÉRêO·÷ØÉ\Sc¾ñ|V‘‡V˜¥ìYÐD-^ª®§Èj#$}½¹¢•^›¤ÊÄn¯÷Ý.²Øk6“5â¼i¬nØº:)hÈúV æóÄ}uGCïá55tW)uI£©´~é 'Ð>§øÇnQ˜]”z™µSÂ•È?¨ðáì62ŸÔ$1TöjTÓ‡¤®î^Ï2’µPë‰»2ì¯¬±ŠTVji ±FR%Ï@K öi4^*L3¾ž)Š£{ìf†N÷LïMfÝ{ÿG+àÃGÝ5òH’ð×1qõN°šÚg8‰s¬î«ä‹0`Ã2˜®iF_ùÙÕ¾‚
öÅµî;½ñ3‰$ºer
/³¬(’ó8p¥—¼<È¯¹½öiåî%î:‰T$ D"PqûDï‹¯ø®xj]Ã8Òûâ©wýÄÁkµ1¨Úµ•–¶úEeÝ;¼%o»´¥‚m±]JÛ,•z¨ÿO»d«ÜKº§žÒÀ·v¾òü @í³¼[ÖÎk—8P›«¶ÚrZ6{òiØ'ÉÉe)½>>¡i+šÈ”¸£»¶C„¹¦CNypIjãÇôÓb"ßqÍ¿Ü ¾°k™Iîc³ÏÏE¥	®eÚzóÞð/¤FïjY„b”D­£l÷r§Ý1¿ÓÒô‡°‰'É(«ZM{µš´zK¸æ
®'¿>zùb}êï¦\…þÔ[s(êâ¾¢¼ã{C[{k´ôK˜hQD:oøyžMD§¹ï¥jÃX[¼-­ècðø}Œnž]Èë^À) –:">~[á®Jüò9v[j	dw†ï¸7c6×'”„ú¹ƒ:¡JY¹uÌ¡‘æïÑÛº‘£/Ôm§ÖøvpÕmÂ>	«wÏ
ZŽÓ<´:êaEæÕ¥‰¸¢5LjÐCÎ8opÂ+H-ún¬€¹\ª¯#”ê¤kHbByË|åcvD÷‰ù	©}†/}ùáž~uz¾–Åö\-ù'BÛ—A“ I%³‚d§@C<KÅéçÉ>4f¹B‡ÂyZz™l¸)`õ4y%ÓVá^IÈ‰rNìêp³DPcæž„¦ \ˆÖÁ¯}±d05øZj»6#¨†PÅ´Ú¾±U2„xù©©e-ûÇ"·~lHfl€ŠØÚÃp‘^€:<S}=æ&`°*Å+p¶Oý®­Ð’ÇÛÀ˜0|í‘×
H‡šu¤ý6ûÌ$Zõ`B¢~÷ÜõO»Î…Ôµ=ôÚá]ö[÷¸Â@¿ñP®¬îûó µ'mf@rƒé­Ÿÿ„¿hÖÓèÜxu}Õp­ÍbŸð†rÔUå+˜T~\ vúTBoLö÷ê³‰¶g–¯;­¤À‡‹þœe‹Q»¾öØúR|Hf3÷ƒÜD§Ð²5Ht»µóÙ¤.íCHYÂ9_ã*õ\]bÁªN¬]}ªÍªÕ9¦ò}ÊB‡çê;BtŠ+b"J™¦žY'¤ŽZ¡º•j‹Aà»†P¨^»ŒïbÇn!×ïÆpƒC×f•zÅ.g†m²FÝ"k¸üõ×ì*~ð‘Èíä|àÿò¹¨Z¢Êi08ÎŠ%ô#W¸ø^'Ü°ätš	­
ä§šžþ{n`É‰·3C–]É…­a2 %)êÆÀÛ¯V¾Ó¡®W¿bGÑ½ŠÎ“Sa`öÕª[s-_]½_œA:›ï/SVœAÇ/Ðy|(,2†Ý¸X9Ih¤XßgºgQ{Çê¯W–ËšH7ƒxÈ”çî‡ýG^°fô~oüáÛ®ÒmZ®ýx	™ÀôÏY÷›&¾Ìw*úˆ€¡„Yzª~ñLýÐÖqBäzÁm¡Ë@Èräý,Ê·Ï¢¼üF9pÿ•ãÀ]8ÃßQCqâû^ÝœDáÈìóœ@Ñ·Ïwò^$EY?o˜ƒsèÜªÎÀ=Ûsûu%>ÛŽõ›	M(×ižŒþƒ Åd`^²ô”¼~Œ–û¯ÏlC‡øÊGÜïWL<O¦h±˜äÞt’Pêè@]\¸~þà6®Ñ•¦Þ:—åaU¨öÕ…:¬u(ÒhØV›†z+q„;ÌÏ.·“|”Æ«û­b­Ü±ëx Øåw6”®´èÕvž¸©è~š¼‘™~Gh}×^E—ú(:`¯è0nü¼Ï¶‡“›¿zîùQ.ÆöfVÝ}°¦ÿörä)M‹ÈZ5ùªùÎ]5NdÿÎÚÀ«¶m½±{1Šžœ}ã‡6Æ¸4Á8¤¿‚º&ÃÊØØ¡HOŽ¸2›½`$ek(É²yÉpž³gë.3Yxî¨VÏ¾ñFZAó¾	túÚò‚"Ô±EEÐaY32\ôÑ•ÅM¨™ª€¥A¦íÇyQ&'W+Ã¸¼ˆã©Í¯ˆ'ImÒppK€¯*îtÿQM<0^gå×ZÀ}:Üþ·‚fú’>}`õÑîf”~Û€Î®€1d†2-V	qÏ¢ªáˆ¼X-äÒ¼ªEÅª­¥&je$DkŸUxs‹ Á†.´	{y>-o®;¬S=„º¦tÞ)j-ëƒ­v¯q×R.dnþª*gãÐkcJUÇZu–öÇtfÅÆt×C¶ò$b}–ŒK«9T^9œÖ«±s&ZµºPËpÛÏ×å=¬ÊûX“w_‘5ë±(¯€q¹¾fÉ¸<p¥„…•ï«£ïUæªâüà‡à
dôY—Ÿ9ã!Ì-·Õ½þuGòØê'b9¤“=ÁsÌZð³ŸÏÑyªúù	6™Ù'Þdfé›ŒÚEˆP]»£,B+gŸ1­¬L°^{7Uê¡­rT3u˜àB¥ÔÐÏzŽäºŠÆvžâý--Ve@ÔÖ‹¢¼GÕ!p«bj¶{[k£ïdØëÓÇ·Ý§OØŽŸ>žM?}šìû­Ž4Øú[úvÿô©¼@Ÿšû ~·*îXcÝ0sÀ~œáˆËÛæ?BjÜÙ%p¨J¦ŠäšHÍFyËkÃð·6…j¯ðé¼ìÆÃµ`pU¡éÝ–×Îò0å¯Š¾ú„ÛB† òÈ	|W¸¬ûpñµóMó•Y–èïï›~W¡õ¦Ù“hU[á“í'¸U^Í%©çåŸh4ŠgåF§™—Ëÿ„²Mâæš×Ôž3´ÝÐOl ^&oŒOVù¼-L ƒ·ZüN“ÝóŒ7®eLœºu„BagaähDÕ‡Ae’òŠŠ0ÿº<‘
U{-sã]óq]}$™5’Ê4\•Þ¦b°Cy²{ïŸœXØ	]Àd+Í§Ç?
TÀC‘dš&Ó8ÈU…{Ñy‘`öœ‡F-çŽ6äçùAÞ¨Îåa³èÔb¤¤ò‰kM„Ù°›ŽºñöTHÔñ°jA–w²‹i[B$F¼e®â	kØ9"N»àA¼wŒ0“c„oìS…ïÂgq!jÊí9Ï ¢q¾ÁVûý~Õ©öé¢†¡¤ÝØÄ€V¹»ÚÀWä-];<·DrÐ„ÉC"rSäl^âR9nŠš…Ñ¼ TH^Ú6Bkûày5°´Í6°ÒAé~¦ qtÂ€°\‡¶†kïgbd3n«#`Ðé<•ƒz²*>4ä§‚ÅSòÒ²¸Š¨Ðyª~µ,è…]xê¦„+²É1`áÄÝiW¯!ãb¶Õ*2Ù&h³0É9fÃÕ0³¼ð¶Zn´ÀÏnuôÇòAÜ VÍŸ	@éc·Ò’ø7ŸzI-«ªõQî³euÙ=î¬mSÇMOÕ¯–ÉÑSó{!BX‡ŸÚ!vÜÔqZÿ½à¥6Jêµœ¼–út;›p#&á{¯eI~}õéaÜ\44‹0÷
„ß´bìëUón#:÷å
z˜ôõ¯%Jäó>™i|R†46åYƒ:ó2J},¡ï‘’××f—o‰†`>Ce>žò@7ú€ˆr‘Œƒj#Ùs¹Åè·ìbe=tä'KUi‰Ä#tE<ôâ0»¬>‹Á7Â*eñÔÄ4¥7çe²sÁ~ùŸ²µÊÕçZpâRR¡ÂO.|¢Rž-6_°B•³­[æN;W™fš¯B]OW\ëâ5˜Û²‹—®8µvróT„}å×¬Ó"¯ó8º¸·£˜w†Ä-"¾£ç±Riƒ”.çÊ¯ÊÚêÎR•#wª·z¬Z-õTâut‡Y¹òOûó<½7pWt™G9ÿ½ÛþXØa’áÇœöÕÊCÆÖ›¦nÜþÍ¢+Ô£aßhWƒþAªž²ÆGHÕ#}ŒÂNû¡å¦
¦g%áç¦× ‰lúVÄ!0Ï²²« ÒP¨>s°‹,J}ìS¹(‘­£>ö!9ª×üIzZez»Åü@£”­ßr½úî„!Ø\Ž¹Ý’®¢ÈžØ›Õ
þæI9ÌÆW®Àm{²rÅäÍ¦…uûž·œXÇŽÓà§*FÊM{”7¥*, ‡îu”ÕG®Ö<<‹•\T9®Ø	«g¯žñjÏzæ«’šiOSuKÄå…j—'å“tõËÌc—êÖkå·J«\ÀÆ «Í¿ÝbBöøžÔÚ1yçYA#‘¿­À}óP7îuÍÕC;¤,ÒlQ£]wF³“Ï¢7Ü"£<8ƒýÎ½U{—þÕˆ}Î(fãMâq2ŸPxþMëÙå^üï6¿—g«­œÚŒRÊÂ¶‘aCMz|òÐ¤n„mÆhêÓÁäj˜Ûûcÿßÿùÿ8}¨S»Q5,š¥O8A3lj!ãOÅ›5‚ÚTôè®³_¥Ïƒ|\äÑL^QŸDðºrüpY´ëJos2:–ãY®iR%:|8wÿ4!¯-ý5®ò°iQâ£9”y-²--ùÑm{ÍµÔ±Gø4l‹døt¸4(·f®‘Kåãß:ë°ý­ñ4Á¢@7­Û0Ú6F%|çÛ†a´¯i¾mÖ@ëE8|¤ïÅâS/·ñn|.ô'xúBy#Ý)Yä’6´=ÜÛo~À&ýè–Žsÿ±}êÕÜ÷5m†.nZ¶NA7Œøíq¿{=¯Bêl8üâ†1Íªï>”*Í÷XéY§*”¦ƒzzÖ°t›"!V=AÄ­qzØ¦åï±­«ñ$¸ÏÝiæ Ôpís#&.Dt4O×X:Ø!rŽÜz'ázwÝ¥…o¿‡ˆ+•ºñúÝíï×
Töq« 3‡í55¬jîå¯îÎ‚ôYÀšáŽ\ÄXÁn*d¸@¿/¾'×ÙüÖí×÷{Æâ(á¯žgùoePÅnÈ¯Z ü^!jÈòº›,æ¹_htMÚßXš}¥9!zü±°®Öå›ìS{—>êè‰tÃ²?ûøÿGÌã`«ƒ˜»Ÿ|4É&y¹ˆ«ÚkÖ‰51¥÷ZÖµ–èÔpyÍ<âüÊ¨WCî.ýÇµ9¾[•æWßŒ§ÇVeô8KyÅË6_‘#Ãõ/Ë5ntTqR# r°âÊÑ)¶+g
‹¢áUu“O2
UÑªç¢Ÿ¨¥&&AÕãÝëÆÚ¡º,hë[•±‡DAüÝ¾Æ'2åöÆmJÞ´¡ùWI}üoŸßñ.åÙ¶m/VkÂ*ÊoIª±ðÃvÄú' ¾Â`E«|ÍÚ -WhkkÚ<õ®Ðö,·÷Qg÷ïj†-|H^¥¯¹Ø©ñ°:Õúò'áúÅ'×Í›yFYzûûÆõ£êù°™ëo$yÇ© ? TÕ|TƒWDÒngeœÈþý;t•›#ø—Ï¸IçâwÏ¶¦	ÈŸñìñtäTpúšIOiŸ\Æcî'¸\YcÀuâó›ÊQ|ÅKnœŠE£h0„¿ã<›!]ËC.b ÙI†«­vC‘La-F)ú‹7Nlm™£(…­x­ÿ·Â”8°Ê­ërëáBñeR.ÜRKs`ø%ÍrÅ¹^tí€½Úƒ/Î4¡›ia<ôMõ&¡[	¸î©r*Qg/pMã¢†ClNL‹ò«¯úš¼\t.]Bþwä—ÚÆwuë‡laãÐ²»8ÁZ*Iãÿ¦||[Iøëˆ~µßdƒÕeÓÃùp‚hnÇíßA³ú©¥Ák ¦Š¡9˜µAE+(ñÕh¾Ð'©íËÈEÑGÂ·T•³$|ð",ÛÎ&3ŒfZIôƒüõçzã¨<þ=0
qõv(%t êzûšŸ,	ÔŠâ¬NëŒ˜åÊÊZmÔ’Òo…WÝB‘ÚŽª}{Q×?½>5m<ÿ<P§Öê¶˜S+@UÄ÷®n6R .÷VS]Ù_Ðò–h©â=}¸©´SŸ+‚mšeÉ÷žšo?ž†Ý.šÎÖ£iÃ¡¿‹†w^,4V{o*?3ô«QjÜÍBeEí"TVq[×¤ÂÚ#¶Ïlñá1%Õ§ÿ$›…ÇâzjüÏu=†üúÙKóg¶&[¿úyð-è÷#©ó©_²ûá(¤÷”OBÌÍÝïEÌ¡Ûg$Ž/›ÝË{¸öðÛ~åuÃŸV©ð±Ÿ«ÔIóOVä„û³Ç«ŸšÀŸSˆIÁçÊEÙ†¿°OúÛO„‚^0èO‚ˆ–UÊçŠŠ®=Í/È¨¿ýdæ¿6[Áþòe‹ûuŸƒpIm²þò°ñó9
Ü`~H¦8ãˆu÷gÜUMZë¤½F#‚é8¾}k”Åš_8IÔAÍf;š¦XƒÁ3ÙZ¤†»³[°ÌôŒŠ ãZä4€[<‡ÊPÎÆÛæÐùøG O}×Z\n¾†K9¨ia‹; ‹Ü
ÃGøÙFÙÆB7¼ðÁ[W¼$Þ1™DÓy”vÜhãágSdñ´&aêQšªcíKUÇ¤°Ÿ&SNž'š<üˆ™˜`ðx³á.Ì“d:nsY—±nÄ'?ê'âþÌQ›ñ´Î'¹–ÝV“¦ ·©~	Ãõ{Í†µÆ¸øÜÏ~§£ª×ßÆj¢tžGLay1q¨ÀÎÕ¾Óæ2>îö"nÅG£²ÅexÚ!~åŠ¡;"EåkóÅvèæå6+¬Ëõö¦	Põ&Ew¢å`ëîÛÃ‘ð•t¨yûs4ûìî$§ÜDM”ÃkáqSïÛõ½éÂ•ÈÓäÉ¢ëîPthº¯¹äîµpÑÈ|"õ‰™á¦ZÀ–	>mu?ONã°¹ébF¡QºPpr…ƒh¾—©Ž|Ö‚N\N‡ù“æRyÇÏ'Úi‘qÔÓ‚Œ4ƒä§Öx«ÇÓ|?‡÷a–}XfÑô$1ÜÝd™ÅåÈóìî>µWæZPÛúeèÁ6^#‹4­–Eqv!l­Qe‡ð«yZÞ/ù‰ÆC•ÝsôFÂþn>M¦Ù2Áœzd©D“J§Š÷-‘7\ø„³à&ÈÕÃhèJ=Á¨˜ÁøqW+ßXÁ‰š5âB1¯*:T±j(5X°É:‡QzŽwÆÁè¸MÌêkW‰ª4g[x©'wzëée‘o,õ«cŠ×ÝKª	Î‰6p!dÕÜñoÆx‰j?Y­¼óD/¬3±®y#3TT\w"x:žiU-¤+IUn8MÐÎâµ„¬ˆÃý;èâžÂc7WåMw/˜¦%%IO»Fp8‹G{LŽ¦°{)Z€nNŠÓ@O…üŒ~”7Ðm¯‚Ì›<íVÓøfX§=¦·³‹~†c†Z—YçÝ0¦:ÕƒIãh,ºT¶èÚÒcn[RSÆ3øÀBZI[×XµÖY{žÞ¸Ö?]psÃÞËÞ„)^ ø¨
B¢š $K)<4BºnKož)¯öÎK$½EÕYít¸¥<¸	Ï6]ÆøG“(ûd>qéoÛ–‰HâÎÎŽ%—šM2 qøs–gWÊûf;:Ó6Ìy¾g§Î‹8/ðG<™‰Šƒ(ÅI¡¿ìcî@$ò ½5KðñÆ|ƒ§‹)QkfùÁÍ€©N¬!DÙñÛÇ¦ÛÀ–M”/©eVnÉ.æ£¬_Y'Îó,ïp|9Ï’ñc3È{½?/íC-0|{b7ýêDrò8(Öµšâð±}÷À%wEšJô6ÿ9‹0X´0O"7ÖqÇÊgÙ¡ðýî‡ÀÁþáQç1Ãë•M¯Øãý„NM’"~©O«é\/à¯è L¼Í8â*ÈÓ©ÂŽÍÎ¸•¸ší¼\$XM!h ­±(É<6Ñõ–íðyo? ˆè +îvðA¡ÕUöë8‹©pO8r†AA`ËKFÚK›ÂŠ¡$y&ÝçÙˆÏY,üGA3Ýñ€aüœÆüg·G(z]ƒã>Ôñá‡÷®\¯R
ÿ²HÇT<çKLÔøš­÷ú3îÐ+/»¢¯u¬ÂãèÊ)*ûQQ*Ëy>ÅÐáØ±›•/®yóøªºy¹n8ÐDýÑx¼]|¸bðfÖ  à*Pìaœ÷´{6gü3è—¥ûÅ,MÊngEuÉ0Ï£Žq–€ÓSítìÑÚà–ôìï‚¨æxí­¾a`¥¯¿í±¶üöP•‘Âvú…„%Ñ€…)ÁqÛàôñƒ—ê¹Ð·"žåÑ’4‰¦{d75þûÀD0-É|À¹#¤e1M€nVÕïuò{í-E½ƒˆqejY‰{ |ÑIFÈ³xãC_€#ôøÇOxž~sÑ¯»ú»UÀÅ¥ÌÛ/ÿ‹ß&°”pôH%:Ë|(zÑ]T@®ð¯¯á®ú°nØSÿûG¾¸–¯Dj[–c“ß8ðvQq5>Þâ?»9pš³¤Ûƒâiài^ÓÏè"JJ‡,;õëz{¦ó7“¡©ÓÓ8ß› _jÀœW”nÂøÅpP“	Õ¥±ê¤¨SÜ ¤))®";É¨Âíô³,ÇMÖNÌ†°sž£cÓþ x4„ã,ßW‹‹=¾Ÿ‰Ïb÷â VøUµŽ37ÛGú“éI†¨ð»éï¦_	ëà+öÅµ Šà«á§ô~AÐ	ZrRÐÃ"¿Îr~ôKë“â²×‹w*ƒ`
EÐEp™œ[Eƒ¥Þ?r l¬ë}†XÄÊŒ=ßù¡À{Àq>J¢¤`i2ýó/¢V
ÎÔÉÆ<øËå%‘u³*Ÿü>|ß3¶é›ì8ôÝ8¨0œ@þŠLAWõd ý$«[Ððñ9²gH·9·† :~Û‡Ý9Ç‚y.2“±æÍqAªÂ ËŸ}IòŒÄ!¹”IØû¯¶Î“‚Ç¼~•Gl«ÄPè²i;O¢qôâÖ~úñËì+@Yådˆ›¯–~7}3‰`‘ëó±“,a#^Ýp‡w#6çf,VtôñŸ_\ãtß`õ ”“è<ËQ;˜MfÑô,f:7+£ô,.HUý÷F4&@Ç¬+1Ö)ˆRë§¾XxìtTePª„N ÂýuÁæÀOëâÈ}yÙÄ–æ¤Ìe^FS oy˜IOÈ©ÄÄ:b×¼Ÿg©pÜúzÿÅîaÿûÝÃ£ý×ïÞ¼ÚÛÙÚÙÅYž«^ˆ°gvÏ¨g‰p'¢—ˆŒ´Ó®ˆŠã–Y}Œ²ðêe+ÄbÝÃy¤ÀÞ³ÑLµÓÍfqÊ	ƒžxU@êJÄ¸÷nWâ‘5Üe>°:9øûëFîkŠ’=4”ì@|ÖÈ¢¤äñžÞ\Õ^*ˆKx0Qíj³¹1íÌäLEaÏå”©¶>Ÿœ¨Â.Ó+g"i]€ãô=„|·'hVÍmÐèh…ü)ìŠ3‘§‚ž$¿Ç{¡%?2€½n”gÓì4&Ñ’…‡•8èÝÕËUaÛ7}Bãr<*·öÎ>	Úg¥S±±ùrWÉÚÞ¹ûz{oë…ùEOÄ‹Yâ!Fµ«wÒ†¼¶†ª0ÇV­„±šßšêÈ&òj_!µz3Èôñç±`Í,Úõž›èèM+L­'JÝ@¥Y(¬‹è¨}âl6î!‚ã¦içïöÐéÚå¡DIX!Ï¢Ñ‡S~î$^q8¯¯ Æ#'	†mëD¤Í²¢\XÀ- :…Ð¨ížœ`Ð*zV Äiî¬ÿÚÞÚ qIãÿa}WªÄ 1m¼¯UÇ¶¦ã×²F²´Ð¢Tåëù”G,
èû!,y *ö ÈÝæx'ÆõŽj§Þñj_Ô7Í.Ž’	jPH ã¢k1 ª= Û*ëŠÒ…Èo=ö„}³†ÅN5b{ÚœN-;½¬4V²=ÃJñ2ƒ)…<â`ÄÕ­ôBKON=¬<œOÚuÀ¥••”TÔô°u®8ÆP>V%¹¶@¨˜ »³êîDWÏx®VÃYº1_[f+ëÎžIŠ_•úÖP­!ï’¦zøÈƒóYvM…æè_ÈC„ñdÃdXg›¡Ó1›ƒ\ÄÀ(,cü¸Ì¶÷_¼ØÝ>ÚÛuØßÞz±ûjgëõÞþ»­íýÝÃeE—1¸¦Ç‘†º3ànüíŽ* ˜9ù÷/_&B¶Çø÷¾
þßlº‚®-°›ºœy[Öî€ÚÞÔ¤ã'¯lÁVqöñ„Z^»‡çÌ›(Ã€ì` H§ó¸)µÝ‡ÆºqDyÀ­òã?1haÉ7n-¸æÊ^4I©Õ+Ô(¶Ú0tÄW5“¾Øªž øªžf1¶²`KáV=AƒåÅ…]
µ¡×†›~«‘o!Á7ŒŠU/ (#ÆÊ>fQ<wÕ2AñÓÜûý¦’*)éXË‚…G¤\±¾ØÜ*{UHxî.N«'(V×‘Ô 8M´¨Xm÷dQ²­jÚz'QZ=·ERENV•`à)&í¨j=½¥tÝº¸MSÐ´CÅÓŠz¶E®Öâ¨ÊÇÀJùXUCo'bŠÇEò›ðöëÆyî‚ ‡SÒçV!.tžGIKú§äÐ)T(Õ»É£b÷´Q3¿TéJ6ûíáæÐ¬:¥€ ki)`k¥›-Ì½;‡E›çÉd™5Š5ºŽ&™&œñ¶Íóxtñ8„Ÿ‰Hcwè¡æ/P¨á(K:Ù Ól©-õ+{¦¿b0”(MþÍaÝ ÙmœXX¨½›^¿JgOþÍ¢<b…m@g3v¢WXq¤ËÊ¯?|>É0ÐgŒƒ>IàcŽqÆ†?þ#üDÖÍ°Pý;WvÙ_öyþ|â}ÞBãöÛ¼Bh`%É‚cÑ0J.3«aó‡%±©ÒPs4¬•ŒþÜÌ€»ûä,uÏ²O,ãÞ^¹›Ó¿7–J¹L&œ¨q‰ZÒl^v+µÌË\gk›¥	…3Zåª<¯RØ¥-³cnÑ+ÀÛ±S;.x Ñç|MK’@m,Ÿðö:ð…š;‹ˆÉ:»ÔÑŠÂ¶ùÖþñOcùMGqŠÔô´Ëkâý“=!ÈCº³í¦ÚvŸn%R%A*xCSês½0Y4Í·:Õ¥5œ²»4¥¦dRlÇÜÂ‹ßh/V~s¯A5&î¶ˆ±LB®‰kâl;ei¦üT@üÉm~IÝÚ‚Î‡GÑ×¼eÞ­z;ã¨8fQ>æó›ÂÖud½ä›[ýy_p—K%¯ýýjW>¢õBµ"â.õ¢9r±Í§¨¦·€Â»±ÉºÂ®Ü>1í¦|½¤ô–©UY>j›4`k­vŸe¡¾e´awÚÏ§‹µk@7/ÈŒ¿RoªJ&î”wjYGÓ€0	Ò"O½Ï¤	b1pWÿ²¢d‚“ˆ¨È"•3vU0ñ¿ž>‰ËHB@þB3¨óìM‹QžÔåš¡=c)S«ù°šÁ°*eÕ'òžGi–›Z*Zã¹öóQ4Î*óXŒç ”üÊš×Þ-Ü?+VP‘¯Lf™>D†FÄ†Î{ªC_HŠ¤T“ÛWh¯OvÏùêz×ÎÏPÌrC.må)nèJ7³¸“àq}r'•æR°ì35­:’¸Ç™¼Þ¹¼?®úøîõ›-cù°dêÕ–Ì2 ¸Î#›“7KD½KºÞ­—{¯Þ½Ü:<Ú}-5¡lÏ÷^m½ÚÞÝ{½o¤äŠ¬.ƒ¹pó‹—3¬‘Ô\ò:Ôi€e÷¶$tŸ4ƒ®lGL7Æ ú»
€~Cöóo‹Gü¢ók>±UH âiˆ¡_@
°°Ã4ù~üøñ Îêƒ\Ú2ÍÐm“xnÛ%),K·Qš×¤‡JðMÖÍ‰n^¹ÿÒœ<ÉÍG7fšÙQÊµÛ¶\"­ŸÊíº…‘‡B^ëïH^ŽxÉdCÇµøfca+»í?ÉæÚtHç¾]—tñšNÑ²©K$ïâ"…kºcïÅM²r/Þ%«xM§<ni)XŸ“Väòš†Wq$É÷S¹Í6Õiå^dVñQ.´U—DÞ[vH®éŽÃðk²ò¸$Êc›ªx Ï>ÝTÏæNxˆÏn®ˆgt«rXñ`-V>¤-70x€;Våås!äåÛT·æØƒÐc]È8²#š;´¸ÍæÞŸîîT+{GnKºÌŽI/,—
ÕË»·!Œ@£ˆa>	ß°ºŠÁè’ dç “4)¯ÌÅ¦‚ÀÚäÆ=^ßGÛ•eø7öx’LWO¹*ú/‡¿CÞ‹Gxuá#±w½5³ÎB¶¬_f/²‹8ßŽŠ¸Ûë'È»Œã¢k®ŠÛ9z†‘çup®jáJ“b5BOj÷N0Zˆò±Èáæ¦­õx±Þã ì&¨ªVx±àÊØMÎ³ÿÝt3T;Ñ¬ê&<,iÇ¢MÍ1—_V—èºEÝ%'¾ÁfV¼º{A€)û;-ßQÝ¯5m7ø®HòÊÊñ:²¬ZÞàVÒ‡XŽÂ—ƒ¥6¸oÕbšHO6ìòz})‘áésÚ”$P¸rô©¢WÛf\ã1ôÚœ\ØÉ5¯1Ã^	n2oÖM¤xí~{CÂvõˆ)°<'|O2K¢TÖê%÷FL·©¯¥ôa$BäòûáVVµí('¸š@ôµÌ vÂ…hvÅÔûØíÓé=…KºL&ÉÌ;ƒ¿ó‘v³ÐF£e&;9±¯YWÞˆ¨œ„•­i'kÖÕ Òî3KŒ¸uË–8Ñ¢m	-Ø¶_Óe'€)}ˆwÍT­Š/=ö[_[¾,{ŽÑÝ»ëèà¯³Ö±Îðêèj¼Óëx”åã'ê†¾Ðj>…1^ßˆ®ØÓŸå»ÑèLÍ¿}A„_4¨æÕ~)Z>æ§%]óÂ¡^Däë9ÝU{	;–õ\3<Â})*:èJiŽV,,˜’ruøðt%žÄy”ŽÑiGøŠ±j: „å÷ÛL=‡­Z"„‘U‡º•-iÂ5©e$Õ#ÅÏÈªifŽá~Š@+ÝÂI§‰1ËÄ-aFeF˜'°'‰dsóLx>FXÈá¿¾…éÙ2ÿJÞy>š0‹s:ÏÝ<¸˜ŠgMMXjêºé‘¾Y^Ú-³!ïÓ°/j[á4~õÌ`¥¶Hýã*Ü±ôûý‚Hà•Ý.ú‹uº-~Þ¨³kº„–…³ø{¥]Äl7W°øâDËx¸È}E’5£.Óvâ“hž–J” †cÊ×Êz%ÄM‡nNâ,õS@¦sZ*0Î‹ ï¶£—.˜³g7ÔoÙR«¹—4S©ï0Em%G}3V^·_ƒá”j‹Mˆœ¤
Í›)FÕmºN'ÔÝôZ(ÓLA[ƒÖÐŠ½ÁyJX“‘'hfâ„#öIÈÒX–%ÞH/ §›
‚Â9“ö‘ã- Ü¨KŽ¹1	pE“YW×wã\àKŠù,Å[±®ÒŒ}•´XsÌ«ˆí›-h	z±êBŸ-3QÏÊÑ*M’C¥’1‡•u ¿IÌu-•&•h¥«¹nçï>þ‘Å—hjGo~à-õÓøqB,Àƒjté|ÎX
¹·.IÓViã®Q+b;ûVÏöÀÚ‰Mé­¶4“ˆÃÇ[ÌÑò0[2£óM¬®²-å‘2ô¹Pr¥
êSØ&zMðà;”¦û|ç‡ƒüKGTIŠÂšpf¬Å@íWJRaF­©>q„aÿ	$³²¤·?BMTúÿà>¨x[›FaâŽBdPßMñklKv1¢M
Þ3sq2ÓVôOé›-½ŽOp§åø‡ä6òacqÄÖ„ƒÆÙFx°Ëz(ˆŽ,lÀÎ#q¥z\®¾Ìƒ§³K9ÈÖ@Å£,Â_f5æà7­ÖðQÓ±rsA]BÜ
¶É³!‹J14¦G›@]C­c/‹iEo^bÒËâØ”kúUhªÏü•~§Ç©>É¦9õ¦²Éš‡‹ªl&Íæ¹­ÃŠÆ[eC¤
†¿SFÜÖuÚÊoÛe±kÈg{(®=½®:©žGž«Î˜?õIrÍ™ný©¡Ó`µéW¯[yÂVwÖU{ŠU:ÕtèTs’ÔêŒ¨ñä§ê ºæÈ8|äŸðPójî[Öì>B^{žäÂh×¾qe]ŽÚ¹’$µû¿}½w´k>Öò6.§),6‡€
áñ]0j>Of”(EÖéiÈ«‹[e‰z:ŠÀÉ‹)7ˆ¶Í-µEŒó7Æ;"~ b©/~Žæ9º²ÛS7X¹ûv½ Äé£Ò;i7!ñeù×Y.yu“½µ‹iI€´k	À^ef^¬jªUòƒ¤¢Âÿ0mÄßÀï–_+²B²Iß9–ìõP•$%Ž	ÇH *Ù,»k°6‡eö´på>Ãc×m³XÝ›r£{ç×ë¼Ëˆ¸T€ßÚé7ôõÆç,5ïW)j,|‘ÎÌ¯ÍÜy+Õlob6IOA€F­Î<å—ÙðîŠ¾šÄ"©ËŒÅú2Ó8'Rõ”:ó8ÓÛÞãÌ¼.uœ.Õ2ªÀö(/XÍB4ÓbEgTˆ‚°w#6O6¢ 	îëƒ¿!V¸èïT(²Û’§Û­'ûb½©£f…*4Á”Ñ)oEÕ¬&<ìN-Û_»†¶5`åµ‚.àÔyÓsrðŠ‰t¯Ð!oáux—{¬“4p êÛnh¬ßáöR¸ü¢Sy•P- ø 5 ·[@c¯ò¶«G0ÈúÄ£í’™ÆJ÷Ý¸;ßç<(¦Z·d§õW¡E$ÖÕô^–0“ÂoØ×ºº›÷wš%Ñùæ™LßNŒG8fNqrŒwe€yý0°ôÇˆÅ—ŽujJÇ¶‰î¾õ‰‘uñ[LÒ˜·Ù~’`vj$sµKô©z«pÐ©!T1ž¾‡<
ÀÝËYÆ]’ùgÓ1ÿ$œ¬cG~€£-´à(d mV£Ð¨Ã«Ø›~üÈL{¨oga’fA×òûiÎ¬-nÔ‡}![Šïâ»/§«…¸€YM2«‘GIjFŒŽ²ÝËQœvÄ–ñz¢¢ïø…·ŽçÓ_ÌÀÞDÍ€>aÙøÓ“_½|±‡!ßvÓ‰ÖSožPÏDã{±Ù?^{Kî`¢ã|0á?:-zÏÃ±è…4QC±N–íõ C\ ›g_—qoú_ÆÏ¹FG”…¿‰yyŠ¥°ž*§Æðï,c¶Ð%[” ;!Ô	UÊÊƒÁ„»x>Î
 n;Õfq+ÁüÖ¿jóË/±ÑcõþÖÞ™õ+­ŽÞ'–yuiÂL?v5¨CN/6˜™5g Iz‘<ñNµ	bOÌ.¯	ZA#ÖP§çù €äÁÑãê6qrÜöm˜û…“b[/Ó«‹Ç‘Y³Ð5½œy? ÀT›âßìCç­1sB÷É™q£Ó„ÀaÆEJáálE<Ö#vH‰«µ§ãƒö¬p{™íÙ©è6BöqÚ‚¶»rêé%ëžW›u»M}#});4ù1jƒ®Ä‚©cÍ!¡«”˜kOîG8ç˜E!ªBN{†©z3´Üe`íhËx9Y•WoïñXf…øàºÔÖ=x.p[ÔMŸÀïÚ•êÚ ù¦‹óLõ@¤:Þ6C û6_šº<y_b~)N90†&kÎÎ£”Ëo¢Ÿþˆ*QèyAî¹UàÚcëKñ!™ÍÜÊM´'
a4¤Ã¸¤Ô–lµh—Ä7[¾»„5ù¶CŒöâÑ6ŠüèšßTÇÄ+lœ‘{ê.>‡p$QVcéŸEE÷ý×¦FKåV %ñÈ;•0—“°Î!:|vYÕK¨qá^2=Ë_m±
î8>rÚœü_žNŒaÄ&ÕÒ“!KeÛ—t	d<…UÂë½N¸aÉéÏ¢¢/<œß¼K.&×È¦sÕ—É†pÖ–°O·˜E£xåjå;K÷zõ+vÙ«è<9gš_­*0¹ÅOÒø’Å
Ú¦Å9†¬[YÇè·2Þ-¼¨p·/SVœÁv±RL‚qŠ××ÖØÅÊIR’È¾~,_¼¦’Œ>¨€¥¶ûŒ.q‹aÅd5½¾Mí•X1Œ‡`ù;`¹ÿˆí•¡”Ið^;T¯C•‰³±é›:a“o†âK¸¥§êÏôpÍ¨F¢Ú?‚Zt`a•d³¦FÔvèÉ³(™+/¿aEò€Ùú¯nì¸ÎOx û57°½wá‰ÞN>×I½û|'ðE‚æöus‡9"å}ÈIÓÊ§øöÿ  ÿÿì}ÙnI’à{E(QÛHV“ÉKT©Ø<ÀÒQ­©¤ÕQ³€Z(™Af”23²#3E²9ü‚}˜‡}X`_vûi°°/öU2?0û	ëæ÷a~DdR%UË»«Š~™››™›ÛAÊU ­­,ÅQšƒ}uV—ýþµFØÍ”‘éhWýÜÊ†gÚÏíìb¨ý¼C—þ¶•|DŸ{d|ÄgålH:|Imj8ì…áYä©CÚÚ±SK"ì_	-‡ý–šgïKüØ±—~½Ñhu“òÐ@¥M·g°ë^YŸ‹-ÿ€…Qú‚cVÆë¡s£sÏx‰ÔýÎ?Pf÷¾à0"P«ö8ÐÏð±B>ç•ëd^“•Yt°Ì+åû.áR‡Ëìüƒ~6œÆ]’å9«·°ÒN¹áM€ý!9W¼$´Ó?öºšÆ nR9ú$F™R´EiÏ›öXmhÒC@ò‰ERpê8ÉèÛzË”y€}§ÎÙC¹KÚ¯ÝÎ<ìÒižtÀb¬hÌÄéN4	 p%!T‚²ÚXl;“òðŒmd|KPg
çcÞô»â’`±ˆõ³‰ôƒ‹BñçùtVžB´âÙyQŒ¸/¦l¦Å¨„Eð4È™±R¶Z DxdìÞŽ·E·M–:ç7än§sr2àd¬\c‹§`Ì`å›ƒ%:¤ÏðQ"j`FWÜÉåúª“u|Ãwu[t5¦±X‘£ù0"ápsýŸVü=†§x+SÆ#¯,4;§‹•I(‚½”ˆ‰§CB$eŸðY?Â‘Æ½Ó	HØféX¯Æ‚"´Y8Úù?Yñ¯òtvIxÈÕUv^ö!;œ»åz›]ûj{Ü»ž+Îð]^ò™s‚Çùe5ŸÉC_[–ð²œ0»÷e5Évì=¶4¾@-§ä¨Yt7VI?ŒCÌ8Äìsà1Î@†g™ÏþÖÉ° ³Ú¹1Hs›ÍÙ'L4½ŒŸ+×~}ÓÐXºˆÊnUøHòÙ¨¿Kÿ®«søÛ ö~FÕQ])#%ñ²ëNÖ¶3b	Ñ®"J&„–='â[·]Ç=tm÷[Î†‡1¡¼ç[tHYöl˜+ÅÜNÖ[H×©Ù B˜%š,ö¤,hD%®ß/2êµü)ÏŠ?ÏËIc’ )ì‘ãÎë|’Â„ãXeP']ÿNcTã|ÝQÍŠâªhU0i/#	bdg¢QN‚²I²£Šl-îÏ]-º(Lk¨_§£LI'w\å«(¦ö¶TÂZŒ#
®ô`#§šGëaªpEáR×÷ vð(tù‡ŽÎ]¾i¹À40ø/ºÀtŸã?Æî|ÐÕóO<K‹hÏj»kmÆ‡s'fžS0ö#nW¶u2£h¾¼» u¿ºxXÖ&ª÷ürË+/Ða
Î¶ƒ”Š„,4å6¸è…ü0?.†„v6LÝ¦7PŒüÓ³ ¶éBÍ½¼:™×Óª^›Tý«Æhü«	²F3
GüžÙ'5ï•`}‰írÐ",©Ø¾ËONŠÉl¿Ó»N/V3øöÁRjì¹¥[ƒ†iP‡ËÚÎGˆDC×­ñ–ÑmƒcûD±_K•×X[	˜°7îWçãT\`3öá¾9||@ôFµ ÛIŽ:žëúÍ ¡ßñŠê	ûîžÀ_onL.Þà²-0“ùV zÃˆ†¨ºàY1%òå4¯qÔ‚Y]þÊ÷¶ÙY·ðžÉ§ÕpN`>,NgP³ŠH‹ë[ÙE1:—KúÀœ¶g%"Å7`”|@oX“!Ùì¢Þ§†ŒY5g–z½^« îd¸IŒ†(jÓ-W|!kt¥	9m,"		ýÀ„,7Ùì5²Ù1$ÞR§‚Q'óé.X
’zÚyôöSŽpžƒÉ/° ·ï=²”õIix²Qâ»€¥Çušâ«í„±t—[r'R|å‘uÇÈû¯8²½ª	µÚb ¡Ñ8^VàèEþO >Ý[g_ S_¦1û5¦¼Oª4¥Z®Å‚°_¸þˆ¼ôÝ*ù‰)‘N|ªX®‡ÄÔÇ×¿ªi¸m¤çúÕá5¨n¦Ä!Ü>ªëü²Á¥1†¿f0"íe’™å»ªùxeÅ=%bzj }Ù]öþR{Ë³¯`gáª\?|>ÇÅ.ü—²¯ôÈÎiÛJÏâgí*¡ý³ß\Œi°™úQÇªl:õ?Ó[1ülô_ÉMhŽ?š½\z}éts þô×^l¯¤\gÄÐ¶XÎÎC5æº›¢¤Bw¡FNoB_­p×Æ5Ö¦Wséè³ìp¢¸)&ErHÏ×FRIÿ·žÛ2û*‚œþ]ÔˆÒQÁÆ¯‡åÕ›à4þq9š8'G(~-pÊí›÷xÆgàŒïŽQÑPÊÉdÔXšÏD>±­?ž(Üà¹bëÅ6¶oÚR›ûêW²Ð®në
½¤‚HÍ~ÍJ=–cðŽY™’¾Ý¨*6,B`÷F‹&ƒÚ"éÅò‚Ê¡2ÔFÛû¡t¥_ÒÕ‘ÈÛ.cx(¼ø®AêêÒìiN¯êù.ªz
KzP\†õë›»*sb§ñÈSQï˜YÆón…™| SO·F†éx¯ÔÞšwÊm¯ƒ}Ò‹9[Ã›z‡^­¥V”{Q‹¸€e—gä‰V“ÆèEÈ>2—õo—„Ït qœ†rh¤,âAàþÊx­:WxM»f?#/‚Ûªg‰ÛL²ƒ_Ñ~~ë÷zÆeµ<¦°^Ë2Û´¿ ¥Ÿ¦¿Ñ­#®¦|ÖaJ¢ÐÅ‰ÔÊY™Otçu°ø­Xaq´(é!se3"‚=^•Æ1Ðêiq|ðÂƒð*Nh_·èÑx-<:¯[xˆ€î6 IPZÐ¥ð(µ0vÓáŠ"vÄ‡Ëµøþ\Í'¢åÒˆ¹ÜP82·p‡Y¶þ 2SO¢¬änæJ¬"MD+í:±'°) ©4"“HIž6¶ÝDÔ©*<,ž,o&–@+²§”4ªXñ¦	ô–š`+‡h#KXO2“Ø4¢IYã“@R³¶›‚ÖZÒ°@ÒÖH3?è,,’º5Ò˜í¶…&p6ÆãkcÍ¡i\±bäæm%$sÅ
¼›7MéšÐ]ˆ4´LHûŠO®@Gi¹®±rˆ4±ü÷&ó	L"5?vpþ,Ùí&"ƒ²çÁ\?n¹ŽÉ®xÂ^¤!¿ïá6`ïØe›š7‰k¢ëéÃï±³òú–0ºp)õk©ùûfÇ=ö©ñÊ~À…ƒ=Ã@â:‹¥‚èeOK€QÌù&U¢Ü¸_1è@Èøu9~·†ûipëÀµ!¸Ö{†“Ëµ-^Õþ*ý7Yƒ“w Ã=/•ê7°ñPÎAÏü?–d-3mUt5ÃŽ­eÈíÚÖÆÆú¨vK‹ ÀZã?µæøÞÞ’´W1þé³ð»Î]Xšç>6‰z.ÔŽ5”2Ö}Ì€ëó®ùÖ7/‡@£åö«W"@jãN®z`Ë#>Á„Šù¼²èÕ’gòb"cÏÈ1]Z¯ë•y]´¿C¿ôˆ~ï„½Jù×ø÷o¼zø×Ø+‘W˜Æ€ób}Ø×EM³¬Å…“0ãtß`ke…ßE ¡ïéàõ‰eÿÚÆò*n‰!¬4ÅVb‹èª!Îˆ¾PÌa…[™QdVö/f`¬„X/+ôzŽ4å Ä\WîrÿAJ‡5	¨ƒ¥ÂšC9CŒ/ääMÖßGRä”ƒØ‡™®©š~·ïNb“Ýy™{ÇÜÈÞ»¥'ùäY9óÅåR@RLµÂ·ÿtQP/—M4o.páÜ‰53\jaê&~ÏŠúÃ¿T}ÿ–ô•ë=,"3ØCÑÑ«SzãOÃƒ•÷ëÜ¬C'\iß?,GôÀëf`5R­šåíZöU|(¤é•ë·>ÏqnÖ ˆñ÷õ¯³'Å,§Ó<¢‡
j|)¬üêjßÜzŽÝs¡÷ãÑ(û¡ˆÖ…Þ¼VjQkŽ-Œ’\	ìµ`´ˆ€Lš&Q{wµB
Fâ€Îb·12³Ínƒr…à9ÎlÂ’D‚„`üIÀ?~­ñË­€Ú)7´!ví¡JÏ‹é|8cIŽ"¤©áþÕw†i75ûòq€‘&ËøìÒl/à‹"Ô¶‰-äâ•ÞÝ f]ýauŠ)²©Ö†¥Nkmi°w³.ýP%¡fyß‡TöàýK>­”‹E2?¸´9ãG¬2~9<döSÅó¶ )®¤^l!ÔÀÆ–_ëì“Á•°À¦e,2™­qêÜðïtq(v¼ðyw7Œ‰Öã†U+™¤S"Lç­0þJÙxnñÃr(^Mh Ï6ãâ8¨ÈŽ³€SÎòaydø?ãÁ|¤z*Ëñ0dV`+ùëzbý…VLGUjT×2yE¶dÌúÌÎžùôøgÈèeøby´ì¢üKF ÷ˆuPè–P¥aÛv–…,9¡è+V&ÊÃÞë‰žš8Ôa?1)i32 è2ø	ËBê$Ü•)v„_X1’^øK0Ÿ(T›C«Ó ˜·ŠfÂ_ü
;„ÑëŸ6'?m‘ê³ã¼»±Jÿ×ÛØZy“©=KXá&æVhYzœà¶KoU©É¬H‰,¤%¶»Iú°AD[³tÎ×îdÍ1^=ùÄZ¬‡3$y™3TEnïäQ{ì»qód½°ËJ4˜ÍNÀXÙ.Kx½¤­;¡ï=NžÌ¡NÅiÌp âÙÉ ¯fÝ˜gU›U¯@à`ùû’«¸b«ê%¨²°>^@¥-uZmªO€¼€„[Ìp÷R¼´Zòæ0çRÎs–ÖwCî–˜Æúe¿•^£VÜÜQÑ/‰¶  -ð@:6–§êæ˜Œ¦þ3‹óeÄ>H/¦­PâüfÕÙÙPf‰?šÍÀ£b|"¬‹V©“¶©Ó@ÚŠOÇ¸¨#2Çä’Zñh7vj3úÙ1TlJ7r¢QÚ§v|¹Tâßœ‹ExØ7ºy–xˆ¤Qò—0[CB®%N6‰Ï¥aß•’íC;³ìÑ\>/þ<ÏëBÜ#n-Sô>ŠY=µî$‘«5¡ž¬:ÏòzVž”“jNÝKŽæð¢Hæ¦©Ç¬õ/Ó(š6€å”Ñ²ÅÐG3qóüzs“’xÀ3ñ$é*úHV¤Ãµ6Þõ¯³ìx–Ñ60J\yúrŽ´eÔFDØ·EØtÎù,?Ë§Ù$îB!ÒLhçkç–«îkqVî<öMç„¢¡ÇwffÒ@ê’ÓwÔ™JÆ¬
8Àª69£@¡çs2ØTÉE”Vb*ó	ÄGWƒ÷™’¤ÌXá‚E‹šq^ÐÚTãŠ jö>mÞ@³CTªø¤JË¿(ÌîtÔ÷&Š4ú-Å$Ul¹ê‹¥ýMZ¬R4åtÔt§A	ˆOwŸTI”Diºs›“Štn¯JÄÁÓø4ýÈ•~ÚO>H1VþØÊ4ñÂlÎ2Œÿ-pu9ë\Ìú#²wüÖŒ×ß^Íî¬fwW³Í›ÓwucîÝ†w·çÜ”oØA¡²Ó^2ËnË°Û±k›3äqO6·0á0û&6Ä âåì–fÜ¾2,ÄéÏo2uÉ/ÀéŠrùAÜKÚ.‡º;Ñò˜üM³øF¾Ùzò°KsÖ¾‚››yš÷…:÷FMã6+v•ãýÎf£ù©ÑìFˆ‡~3q—:½6YQ'4\øÞ.ì2š%zïm+H+5j¾;³;êº¤vÐ»ðµ»]–Ê¾Ú1¯6¬Ëe\²¡÷±H8viç€kµ[Œ…›4ÃÇ¦à(¾À„§q„â·ÈêöL°@†Õ|F}qÆÕ¸Ðcnê?|	»ÃÅH·ðt>«£a±TIT·9Ñ,íÃ(Jh'ˆW×‚m{ê5µ6ýö·™þ»•Ë¢e½­™Dû½bŒ^é¹þLqt‹˜x¤ø¸éÉN¦rš]›²›Í,Ç6B1$ïæ©W8ªùÄ¿1Bè³WŽÎ‚Ø5­Oö¯ ZaLÍ‡³ý«·ÉwÙW ìwÙ¦Ï·­Ê³ŒVÇ?˜N *áý^§E]6T‘óØ%‘Yª5ñ(T1Hö|–¬H)éóüF,ðñO
,½Ÿ'gÐ0þ¼ßùéx˜ßÅ`0„‰W“bL°d\¥  OÁSŽ§Ålmƒ%ùÌOÞ­ßÞÈªI~RÎ/aX¶ÆänñÔÊ
ÆÇ”Ô!<J‘Ñ.ÿ©$2Fµwñ´“Nbò
ëyÀ%H?—K@U½tR¤‰ˆÃá’ü!³ì),ËeêëwðóÍ´†óxJ€\Dð¢f‹v@è·p÷Œ©á¢ù²<@á§Ïn.x0ÚçhÄD~Î’¥aö'	©©5$2¼6¤q¤1¶Ñ\Ž" ¶.¨æ1óìb˜jjE2C>M
âl/‹ÌªÁ7v²{Ë¬«„¹Z~m”iwb)S‹“oBÂôBmÍx3ø(’¢Ü.ãV‹pëCZÀÛe©’PôöPŒôl$¡l8NCm’ ÍàIIš5ä%	5€§&aoTüwû­ËFÍàò†ÿ	Êm®HŽkpm"©ÍÖ†NLPFb·q•ð/4†Yš‘öèb-Ÿy™zi(¹í½=²C*ßö‰/~w,_’ò¨ësºÉ3~/WŒ	5#´”°ž	!½ Ö…ìGBI+ÔÍÂbÑáZï|ãÚŠÙ«$UÐRmÑ(µ§›ü|–äôÂÂ§ÒÌ‰ Ú0Ÿ Ilöf"À K4«}>k¼Ùc×p*hð¼M¨CÇcú(ö; :Üsy=£!öÖgƒÔ¦€ÆhAqšUU5šÔÙ¢Òë…œ{­æþ¯þ‘“ç5†ÖžEÞ›WýK¬¡¦™3Xk¾;›x‚(:²¡F0|2Ûå†¹ƒ˜êâÙ+ñÏúÎ
U2XSIQÄX‰Åãý…UpöpóÊN0‘xxŒVLž­ØµQdž„	$FÉ}·Ð²5‹Îµh|®¶ºZÆèj¥+§«E¤®,!VWj´®†ñºtµ‚òë1¥y4LWŠ]Œ?T×R·C˜/‡›¤‡rŠj>ZÇqÊ²·ÙÂ¡œ–ICŠVÂêVeX–ª…•ÞÚj—à°ÒU2-GÓÀª¥Ã}}È–ßÆ~Ä6¸E+íÕ5¬,¤´aeq·×S©A éˆ’ünKI¤F€¤%Jôûëj‚h „A‘iYª´¹¶
´±6j¼¡†ÊÞˆW…#Þß¤¬Ò¥Hwþ°,`YÚ 3”6Ù¡x½ÔQPf‰‚²@¦(Q½q¶((3FAY k”´ÌQPf‚âH 	u2KJHªÒ4”ÖÙ¨ ˆŒT	£ó%­J¨Ëc±ì¯Dò¬Õl;a€²\AY8Ó”VÙ® Ä³‚@i”õ
Ê2_Ai™ý
JÚÄ–”‹5Õ&”…³aAi”Këu‘¬XPZfÆ‚’¶D2dAY8K”V™² ¤MiI³¬¦Íše5Ç2g¥CÉ­Õˆj“O\#Dö¨S†±²›ÆkÜ_-"²ƒµI‚v•Ö:BVÁ²ÔfPÜ,fé`n•çé4%×™¿Û&›¬AÎ3­ÃöyÏ ´Ê}%¢5Êe	yÐ ´Ì…%mb­s¢A	çEƒÒ 7m0ÙþL˜Quú35š¼yS¹›ÓnÞDá•pfÍÈ•‘žb~àßÔâlÞ8ÝèuÑW¹šŠ¾äðnˆ_À©l§§l£ŸGaÞòž¿ëö]©4µÖ¡=à°W}àò/&ùxÿjçÚXÍ$KÞÌ§aÏÃ“jÕÃ?\G·¶ðaeIv>b¬~#ßææâù³¤ ÁDÇ|ì¶»W%'ì.^:ÚÈò¢èëÆòÒä Âðç'ïúu5]ƒ]Ñ_€ºD±Sˆî)»8gêå˜ |¸u%Lôw³ÕB!o£÷íÆ~òq9"c5jmÊZ›XóÂ—Pqƒ§Q~±væ¹ˆAÕâÕbF{„Úƒµ×ßn¼¼±É"b3c-ÂdíNÀ*ž(˜YËW‡X!‹e÷¥ª+tÍ0hÛå9 ØÕaõ^ä„@ÝW}éìBì2Q…­3§ùpê—f%÷˜¬¾çkÌã§I®1ÑÄÈô…qLáø/1·?SôPÕ=¸ùFªÆ/æÇ£r¶Å„ö›º…²"ÂÙmµ9.Û ûqžöÙ£z`@Ž‘ÅÐØ´4Æ|Ô@0h,EK*yÉzÆ2–f_{8„g€!gôºø3‘‘¼‡	~¹?'AÃ­äëtó*4Ý²ƒU~ñbùl_ûî‚Lì„—88€ÈÎ½%ï@½1'Ù»ÙwÃrö2ôì¡uu•Us±`Æ„$ÎæT[éI2(ö b4eM(q›W@¸LÎ²—å¤b1†„îÇâÇQl54.þSJ;·ˆZG,†ºÈÂ}ð›üëdêâpºqïË¤ÌØìÝSÝ«)àC²ÍÛtn!pnVç0<`ÿU_Â·MÚsC›Mçü±ŸÌÁÈ ë>0êãÃŠeï:Gá˜¾ó‚·ÿ,¤Ù±vIH-$ëØdûB­9ã›&¤óA1É{<ŸHÐ:&4¨!$®š8>J\Ï+ÀuÈÿ Œ?A%ÈmX¥ˆnO¶MÎâE]œP¥¡B#ÚT¤‰]ûû0Yô/H@q¸lÂh…öÐb»/L;B™Ð¥vï¨WâòÊÙHÊäWb[„¸<†£»¥Ö±ù=™¿®¡ÿ
¶36ä™e@VAóc*Á\¤_DC÷ ­Ú5ÎRo<#‘IBKJmÜÓbÿ
¬xÁÕE°þ3jS^pš³+Ú…’@·bÜþŸ›ä*`rùð/`uÞVèÇ2c‘ËÀ4Ú¯h2BÙÍ}ºR»nîg
#Ÿ¨lî[ _ÊñOK°¥¸>Ÿ“1ÚHLN¿`°zwCü‡ªþð×ºôJ-gå(€¿zòËúO=¥uÁßz©…¦þg¿¸Î™ùÔÞ¨Ò™[?3­ó³:ÿðO9Ó9·Q.#/>2‚ÅãV…ÑúÊëÂç à R<"UK|
¬©î‘‘Jî~!…¸Ãp“¼¹ŽÆIøú4û3ÙQd†þš‚Úá\†ˆ¾CØ`±k#[ÏÞÿqš=/¦rxÿð×÷EÙµŸcù"¤zeGãéþ•¡a¸_B¿%ôÊÎô"1dBô•ÙA¹èÿø6¸ózô©Aƒ*îâ£†¿‡ÙkoÅ°Ío *@ñ ,›FtdRAMk¼¢™¦÷3ïªrå[`ð|Í1(µG?U	h‘Ì¡‚vnß¸‰só²¿HÙÌZ;-ëétúAéRö¨±=Ô_¼ÞxC5Ô±ÑÓþ &I=2à¡}ÒWªWocÃbF°÷bÆïfèõEàºÄ?~«!j$h
Þ†jYcX+	×/Åw(ìEÜŸ]·¾G‰ê–aÀÁ€YvKÇ.¹MPõkH…_æ @S­ïùJ‡kø7ÔKc,ß¾¹k°»'¹Ak1O$Wd>+æŽUFnØ’ñ½g³9$êùÖÞ{we`Âg® ‹Rde}"\>>e¿ó²_;»ª¯ž+fŠë¯¿M˜'Å,áÀ*-»nPË&Bs‡"D8‡'ÍuöÓÕ5p.n—ä0å#_âÏLq…èv>o¬ØˆÊéú³ÝŠr
ŸöfLv‹ÿ²[lÇ-/Ö´T:·æ¶™iP®¢™ny&ŒEzU.•0Ü‰†Á‘=»_ÂMž=#'Y0™+ÿþñÝç_­\u²@¶+<Så‚¸‰ôqßÊd?ÊŸ}4ƒ¼AA¡îª8;É‡Ô¨¯Ê¨ûd=»=XÍîîÀæxscàÇÚpŒ¾OÌMgÅ„|ÐÛd4sH¡VáÓ%…ÂO|!…ŽÙ\š¸¹Óó›bJ25zZø'Øì/«!#@…~•›—…ùÄ·nZ€•/w‘ûü«Œˆð«ËÛ¼~™’_¿(ÈÒ–"ÄQÇçO~mKBKyBâ_½m'¡]hßŸŠ“wÇÕ…èE_ß·Vð§Owï:ñ«äîá³ºÉíÃÑÁ
.¡¥häÑ'6²s‚<òµ°Í¶A˜U!BªØª(—ºq¢Æ «ïšbzØRúôÎŒê.EæàÀŽh†>ê;ŸÏêòx^ÖZì±l,<í[Ê~IA°aô­ûø3¸$‡n¢2s¨ïúBˆf„B¯×PxN/È½Ê&Ü¨¬d³:?yˆ|^­#BM8¨¡ƒVX“= wøp4¸þ\©%ÖïJø‡Á«¢Td!Q„vØGÎoý|
±ªl\LgÔâ}Ù‡7ÕˆÀ´Œ8¼#Þ„Ž“r’ý1 âÀr3¹Xõm¤£ 7¡åB-Ðl“Œ+’LÀ»DôÂxÂ³'…€e7‘%8+Â7Ç¬è{‹"˜ì”E®·ã’¢1Ž¸ÇH¼ç”ˆ=¬0ÖÎ>O	…J=9ôi¥$Cn¸—mí’y/uñXi¶„¬$¦“f3dáâáš—Î5±‹Ãø\ãÖvé–}º@eŸk3õ¸”(iÁmEÑ9Ï„›¤ün7µ7iëÚ0€4+é’ª*ž@©¥Òš‰‡V£_¥ìó‰	D'G"Æ=b*NLïqšçË	ÆE ÄdòŠg$oH-éÊú]•žÈã.ýÌirŸ¹ƒ:‘t\OéùE‰¦9§}ŸÞl1J-ð
Á½r˜ÚiLÌzYD´‚Òµ hèe¦MÅ$(;‰è%e ¤-³”ö1kCrÜÎë£Ywc¥7«^Áâ9@tûŒD`3>eé¸ŒÎ£ù¸Zô„Ÿxæ%üÄ {z§MFM£®†WY1þ´™¦ˆ„‹·\À0~»Þzh°šï)5“Î‡ìŽY4wAIöþÉ4ËOaë“§iÙ¿9µŸ
CåS n'Å¿1Rƒ›‡p/f=/À’‘FLWq¥"±AöÖ·ý—"­-°¢Šª2þÈ²ÇÜ7‹ÁÜ¦Dõª©7*ñ;ÌJ…'^Zj«F’m‰¶Q çvw'‰·'íîO"1Ò²yÝð-Šy^=#<bf‡ÙÛ£ù¬}ø+¡”ÕnöÕh|èºGÆ?»~i€ˆ%ô66ˆß~²ZaÚS(M8µ4Èˆé²Ø•,¯4u\eSÂTŠQ¾›ÙZDS×‚GþräU˜·~ÎÖ”þé“Ø†å¿ÙÏ€ÈòmÔˆÌzír¡üz©¬ªLg[šR.ß3Õ¯{yH3‡uóÙ‡ÿmg¯ž?žúý+ª%ëÇï5rJs—±;ø3%r”ú§dN6öë[2s8õ 8È#U³ße›þô¹P8c€A7á‘kv9 N?Ù~öÚPSHEtÑPñ5þ›F#ÁxR::¨°Ú8¨.¾iÛg/©ò k_#pÁÆDŽ);uzBt{ ’m$jo›ÙF¾1s_äCë}x&þiIÒƒv»AÒ_äÃ÷$.–x,*ýèŒeIFK÷Ç÷†HÞöó½uaß|‡ePïÄþP~LÐõú7¿9ÙdØ~©?–Å9•· ÐH!ôH
{¦ŒËOªBýªÆ/«|:c~_Œ	4DÆgxFsÔ×sî§-ê¾ ÀúûA>›M&ðä¸šÝ«Æ§åýx\þy^¼Ó¿ªÖ¤®NËa±ú›ëÝLŒrWÌBLâ5%
tà‡»T…ÁžÐÁ“'"›EYä{Ç§²›uG»Ù”f_Ífäc‚Á''ÅtJc	u]ÕJ¶ÞWeÿ÷ÈÄwiìSŽ3.£Œþ÷ûC6
+ò\ÄŸ±9‹qê`#_t'ƒj\¨1ÈXó3ù€ö	–#å´Øƒ¾~¯ƒš4ðøû÷ØE+¬o|sXŸ˜¯ðt÷_Ãâ¬ÎG?Ì=-OrŠiûY>½Ÿp66ãŸüÆ¦J=_ùè{kðPcàÃ|Ëmb…[0ŠÌÆÃHþ«zH fÝ•ÉäáuUÖ«å“òïŠK¼Î{§Õaã=)·fh‹È[½auÖí¨ ÀA² ²Ù¾ mÏkš­CuÖëH¶¢O‰íåY}iô@Z!è@=¤]˜ôüF]Ùû¾f¡6ÀC:?ÏËYvZÌNÝÎ:™ö:ÄYé‹ËÎªÆ›GÅlP‘Eê<{úâ¥vE2 »‹ ê®ÁÆ!¯ì¬ 4ú%áR)ŸÐL,€ë?O«±ÖÀµú¥ìf|ñô‡[öòôÒ8ûIëÇ#ßÈ<c#£»X#ó\#B[h½ÖµþƒÉ¬Ï mžÐFÅð•ÐŽ¹ž¯åÉîÚ]±O,‰X£ÀK- E5ù®zhx‹×íqR¥ûÕ¤”K!;s²{‰t\Ñ=›Iô[Ï ¥è6CQÿÕøÝ¸:gŒ$ªoåH¯³bHPÌ
ì†·vw0ˆYF¥Mßë·ª%†üA “!¯uín16¿Î:¾Âù!OŽìDâ4™ŒxEí§kð­ÁëiAHõàeQVA|!‚üHhò‹¡óÝÛ¢MÔüZ™°ZÚ½ÞCêƒ.ø†çS£‘9£¸Z+¯ô'©ÍL‹³yIS·ëã±z«ÌÆ¨xF	8á’Z{÷Üço¸iRÊOûïµÆŠ_‘Q©úåô	isøtRŒWY†ùÛX%•0†WäbgÊ«ZÎþÈ‚Å»	÷‡”2lÁŸ«¿áÈ.aÏ£ºpæÉao>DfùúÍA÷õ¬¥{*˜†Ñ˜ö‡Øúzö„,UÖ‡üsÀƒUiÐ «#Á´zÜhØOò‰
ï ·§¿0Z¤0eÍr8“'"‡›	HWï‹ìe~¼›KÒßû)Èúp5ù>Ÿ5©júÕ˜rZ‹T¢9¿ŒÞ;Ð\Nå;Õ ýIšbÏ+hœïÑi‡#¯gãÁ{"A?GÆíóGùÓìô$¯û´mh…4Éè˜Îd1µ{§¬þ“bTuuBnyT×ùeï”ˆz‚Ì“#kö¢˜uAHff×Öï):±"ìã¾«ÍÇ+>C¬0gV+½)™>ãA×«Ùkhï\¶}áN˜[ä®?¼ÿ#`åT`Ž˜4ˆÍ„K3½©ÜßeÇ–Ì]ÆQŸªÎAÉWçd,P‹Z ð4½YE_0
“ï–Ó{Â”øEQ¿§µç=igL¾ÒØ¸cJešNFv®l+;„Ò‘³‚ûHµýTØ¹™_:'?Îãøùˆ5àš1št¨ïfÐÅªÓG|ÙÉßNOÉZË²‰IdŸFh¡õmOì/™‰C¬úïÝ64h7ã† â€0ål@ñ¾Tk*6SÑ¸“Ú"Dt*2"’g èŠá€KNûdQF%ärP;…È·/ºŸ_Ò@Î“øYR;ÑàKuöŽG3qÊ6ô3M³Ÿ¾“oúâŠZÕÒQwÑÏ²¯³Í¡Ýõ|¾¢÷Þ/OO_–"þûz¥Ó%	­éãSÏí&8Lžä³AïtX™O6¼žuaTdpwä¿¶n¯˜ø/ÛØ£X¿A/ýÞ’õ¥8õ,q- L(¹eÀŠ¸ó`¤øUG®ð8cŽl~6ÐDê°I¿:ØÏ6wÔã*y{X¼/†pÛA¨ÃÙ‡
¡ÉéÐ„úEµÙäÛçÁó1Í>…W£F¤äCrbÕžŸ1›ººšR#®¬¤?¥®’þºkØ•dÒ÷x‹êÝfóZ?ùêSÿ&aæGäjgÇç})¦ZÕ ¯FfË^hóåäŒùïÈœùWI³ÞN˜5_ÃÔµFf|Y‡Õ92ãî*µ	³ßr¾ìgdºì#ïl­i©IQ³SYGŸ’šù˜¨|åC1‘ÈZ·OÞ4ì“õy8³žˆü‹kE<½ÿ%ñ”Â÷Ëš0‘ŒÑn‘ž–Í—Z%Åa™_ÑØ)R©ÍJ¥!Ê^Z>2];Ä4	¬ÓûÕI·ÿQøÞÓÇÜ{ùèé/zž<{þàÅÑ‹Ÿž=¿÷àÑó£på´¢ë~b¶«zW‹Ì-íf»ÔC«¨ “j£IQxp]k·ÃçHþÅ#‡ÜRÌZöw	ë¼Tx/kÃY?Ë+Y½æí‘3¾ÐÔâç|=ã30•ñ=\‰XÔ#·S¨0’}:‚a7ŸžäÒ[Gdn1+)¹Z×$:ÆCÇ“ŸQáÕªæt-ÐFˆ1š.‚É²ì@m»|ŠféÒj8­sII6o¨-¬öõéxE;-ÈªÞ2/ybQs³ô²ÿJ›ŸV¦ 3«8ÝºÙ¢	Á'K—Îl‚|†Tw:~ØO0LêF°¹è‚-¼–ó¥Ö™DÃDÎSGó¡XYëS>ó…	'ó™
;6Rš(…à{,!÷óéà¸"Ç]ÐÝÖ`V|Lp»Ov¶Ü™`{Ìî´ŒÞMJ°m( "ùÚÕñÐ—ìï‹q	¹ü¬¯u_0^ÕÝÕ;Þ£Ý<I­m?†Lèïé^>>)†‹ÌG4èä‡¼béÛwÃs¨±F<]ÑÍðtNHwJ'îfê@ÝPÛÏjFÕ[5N+‡Zÿ®®Æ)Ú6Ïj[í[ìÿEþ¾Pl¿ØÍžùÉ¬:¤)g{xùä1ý5,`khìª Oyøè~qšZaò&¸_æ,Ž`ùÏ.ê¼®I­—ô~Å%¾,T5¢ŒQ!ŒõÖ"Û££ÄTdÁÑõð€½zw‡ú/ê—¢Çfõ'ùå°ÊûTD Õ…ŒÀr"‹ùÂi²ÛgM½%ä~™’S9Ó¬¢½ÂjÂmÛ)½|5«‰ç¡:Ïý½"Ÿ`-£¼têÒ‡è×<e¤S?Çêç%‘¢ìì)öý	ÏÇk~ÏžbßËñ»'ùdj×ÏÑ>ôMåt¥¿Äj+þoWUoÐ^‰8ætFž¡}˜Ò²ÕþÒ“ù°@B_`µ\ùJ…h–÷òT$7ä®—fæÖÜ5Ê~£rÿ5ß®LIåQò]Ì4GL¾€}A—€ eâ¼FÊéý9»\.„(?%#”t‰Ëóò¨‚Kã”d‰a ¯µ€è<£-ètÌùöJ“ÄÁœµá!G˜?~økV\€fIH¹áˆ
{‘ƒZ=£‘/è œ“m` æ‡êüÑX“Q,hÇä5i{å¨Djj]£
/Žúñ4¤/v¯ò9ü'žc­1Â±Vk…Ã¹ºKgÐW/ÖP{”eoŸ‚‘ÜÁWdû–Y	éiøIp™ËÇäÈŸ+ ÐÅU«þÕ•Ž€×,"k¡†-’Ó&r¸ÐƒFúEï­i… m•ôÇ†"ÌKç±¿ý­Ád©-<4ÍŒì@ý\4˜ž‘÷Ñ§lÚ[2,1‡Ûu¦´
ƒ5 qaTƒC¼}D/;J€ÄÞñŒ½õãƒÖ ñ™au±‰­òIX Ò~ã›h¥Î¾/€&Ù=$Çù3öC¯ïŒ3ÉÑxD»¾¢ôüéã/zß?xñòéóŸî=}òàù½GGÕ_nºãheïö ]×˜CVÊéØÀð†¾KÝ9Rï¿4>­L‰×ú’z;üiü5\§jÉ‘†»_“}c‹ÆˆÛ–ÚÅØ¹_‘³ÕõÿP„Tœ¢û=a9ä,·òõŸÆ?~øëjö5ôÉæJûûúÖŸÆ¯ÏCU…ÕæG6{ï«+…kèåýiþ¾ªW5Ê¡Ül4÷î|m9Ü]o{vè+èÛæ-–ŒlfgÑèv^pÝ­!=DzTëG	Œ±‚ð¾ñ"âtj9k Zæ¡Ë\[õ·c‡¦«ù¬¤:¿ÕÑ8M°ñæý>&'Õ›5$%†bz®9bH–ÉÔ€:ÝŒXóE^ù"¯|‘Wdù"¯|‘W¾È+_ä•_«¼¢D¸¼"Õºu´fÍ_šæÏÂÞ–Â{ÿÎn)–¤
X>0È÷ŠZ‘^k
ðéýýóG/¨—åÿÎ5×þSæ„Ç:~éÏÆ}l¡µûeVaYq—ã~uÞ£:õ¨Û¹_L‹Ÿ	^]œ!@2ˆc¢ïÃÎŠ®àˆB°›Z`hëê¬]Œª÷e?ïé«ZnC¤Øâ=ü!Í ²rIþ ú:½ËÀò#+!aâ[
éÂñÍHÆ4Mg«S[«õÝ|øÎ^1k\v#†±é¦%¿´ÖWÎî­4»•ñ5'ÒøpDE1S‚×dÃ#öœÈ7_]yGr-Ó2};P™O¥4Ì)}8ILõ©Û>Å†ZuM(ûÀýÝa˜b9(™CúÝï\"¤åíWWúÇ8(<?üØƒ Ù&ÇGo½bµf©Û WÇ|ÿNážš ýˆ `žAïbD)¸«#èƒêv°Ÿ¹fì}u¹Œ¦¦øÂÎH:¸±kq=!PãÞÏþÈž*“ ñÎà×ôíÂR|ôR^RÂânQ¼íˆ×Fcócí"ÒÀv	ßÒ‹Gùœ^*~ø§Š¿ãwŒâõwü^Þ±ÛD³Ù{ü‘Îš™¥ï_h·x–‘ÙŽ¼¤ûy1Ï1Ûå6u®a”ñYçq9~—ÑJö­¸¯Ä>cWvê;úÛúP^¾eâþ®â5Ìëq³–ôV—~Ý˜¹ÕÏU9îŽ‹ë9±§¶¾¬\œÃ®ÂÝÕL0éOÒIœíw7<‰Ý ÍX°e¸ðBU –Ò0¯Ð#5ÀïéaïõÆE¿oÁCEÖÙsÚñÃº±ñSqÉâGb2†ÓˆId¥É÷9¸
wëêœ“UPÑ¼+.§ÊÚ‰Z!ÜOÎÿ® ÆÛO&ôŠÖ‚v|'6òè5|†ÕO!âsCæ£Y£»ïèˆÞ!7š¤mó©©±äQ€®ì“I§¯Åï7TF€d…tµ6¿ñ),ø·²¶Æ,´tÛ<ÔÍWþ…ÀÊëÉ¬û>"²šú-ò—e‹ô{g¹†ÜSˆTâ>Ò¸—hÏÜ¤u˜êO3T-ŒV²)‡–vcjFžý·ª¡AJ“/mŒò™]ÑèJkeœk@´5úw´ÑµIÊõ)âÓõ/#ÿ‚¬_Ž©l°Ÿ)"ÀØ¯”ÝÌ>Õl"¸mÒŠz&¨Á°w$4nGðRTÞ	0Ü9Ixƒ¼þI}r?Âø‹ELÞBY‘|¸JàSÔ4ŒÚŒÆ:Æöà•VlaüôºaB‹MÄèñûÌPŒ¯4›4½+1XÁ¦Ÿçç‹ãhÌ]cÎò½/4YÀûÄ¡ÜßŒúiI@4DŒÙÇàj-Zì(½édX’cÁjÇïÅºuÁ‡|eùÍZ8~“.±2¹<³Rä(Á±€.;!{t·gÂhQxÐ±˜e¦ih#pv€F¸¥Z ÃÃQ‰¥¼iã´°š–’ê'{¬¹Zuó©²¨L„8»@šÍ%Ú[]Ð gÝõ?Ý_?[@à=?o
	ü ¥È4ãl58nê‰.2=ƒÑS+7ý,³2M™™<ÅÑÖ…¨ü^4XZaLŠN‚¡mi]By~>Ä‘žÓÁLžšÒBS`¢ŽžÐƒa4š	Ë„5¥[ë$KggØºØ7+îES–8BÛáÉí£«ÁjPX–¯ò\k\æOe)œ-u>_ÍtcâäÒŒp“ûc}Þ¡0ÔMÇ	×þÖzAÕC)‘»ì]LbÃ…ieoÏeÅÖ0“,¤«'ýsEã ã³1õ›ðfú®œLì\Ì“þX%îB“$´Ó6Ù¥5=oÓÁ
&÷ƒ©®•cKÙK·Ž1=€õ/´HÒ…#éòiW{gµnÑ~ìÿB%´7û0/†CÚôGa®ÀE¯ 2±Ö—¿qÿB,MÞ~u%zÒ´¿ì Ó-PŸóEfá(ÞÊ×YyF(|Ö{KqóZ3#iyeÁ»®S¿š²Ëp±ó…y½#Ã¼¢Ù€ì„NzŽ-@lz&¡m#¬¬]o²¶í«©§áÀãdn99M÷¾›—Ã>!‰[Ù”p„ý«­Û×f¤Q'.¥Ãwo°åD>Þ2âzj1{¿Eòª
•d&U’fûëƒ-«Ç‰›ÇUF‘Tí^øÅ4ÜYO Ÿº„› ï?üó¸¬ìN'0Ø?[æ‰Š×32MÖ6õ¸©£Ú‰AŒGpµ2aŒD¤";ºi$ûTÇN\fÄiE’.©8JB¡Ezu¾)¡FsÜŽÌÂX5áÈ¯¢l³X²‘ÜaÎ´gålH‚”âÊ=2Ð©Ù¥Õõq~YÍgßC¶6ï^Û{ñP±mŒ†˜úèF{ýì1ãY09Öh©\:ˆ¬\$%hÐFÃøL·ó˜Õs›Ï[¨ÍÜLhÈå‹µ;7wI,§H6HÀÏ†ó©à8.üi°zjXÄåÆ¯G×È"¯"ƒ˜YÈÑØÜ¹X ÚtAäf¦©ß{5#¦ Ò?â’ÀÏâtÃ™½ëò““b2Ûïô.†Ó‹Õþc¢¢¹ë7€~JÓ”ý~a¥ØµÄ'²~pãè÷ð!$GÓý-¶°Z,öÈÊû~u>NYW6Cl]]Ä6H™\¼½Ì³ç%œþj-ƒ*[ˆð8Nüy]TuÖÇ—NÙeS2$c—‡S4èu€Eu&Ç„8Q2e‹#{=Ä$…1Ùƒ©8Vn£8°+-FÐ0¯Ú•ßÓQÉ‘F¶5EN‡u}c±.úZÿ¿ÿùßþ+ãW,Dƒµ…wPa9kC#aÞìÒCpÙèà>íUûïÿ÷?þíA'ÅÌ>ÞÂiM?ÕåÓ†ø)/"#ÃÿþÿGúGäÌæˆlDiÍ<Íº4T`>]q	õ•5TNš '1wÍÑô„A`L„‰môì4¾’Ù"­=“yº£7ÒKØØz%r(W†ÎQTŽÁÌd{žO÷!xÅÍà¼+>! 9äÜ8¦;¸nŽÕ9å(„/Çýò¬R(¯~»µ–ƒöÖQÉ’AX(t!¤Ü	HàÓì©>ÉyEvlÉ*]”?ƒ²ã +&Ž!IhùÑÆÊ@«Î Ñ\´ÙHWïa½ÔÅvSaï~~¹8?žVÃ9dm'H9«&k›ë[[R:’KúÀ<ÛÈ …H¿€<@ëv›zf¨(Vgd…éõ5$‰®ÕœZ«öz=»žJÅÏ³NR(#ø¾/p¼‰”†dÝ²I-${=±/Ÿ’qÎ=™Owáâv»ú!‚>Bt²a9.ÖÆpñ‹ÿÁ#Žî,’x3õwÕÏ4Y0–*-˜$³Téz;1Ù—Ö’ÝõÙ#À¥í1+uW‡Ä1EBèfQB±E«ÃOê­íV¸"Ø…Ž6þ#Éž*šd›Ï—Ç< ñþöÖÙ»h%Íâð@ý\Ý0$< ¿¤$>iOx ÿL®lZöÐŸ" ¯ÂL(ÞD2a}›Áº±_Ê®@âH¦m$cÇç¼Gòô=BƒÀ¿“«°€ô?É•x`¿ößÏ}E8Pyß¿6b¼¦¡®‘ç×Š´WfÚ5f€8÷fý}Ð´žók	ÝkOþÂ9žˆÐ77=äçƒ·Ê>k9rˆ/8Q±j}–h[Åå‘+'CÃÜiæN%æN}˜‹æÿubîÃû?.eeTè4\•ù·~µHjçjF]ÁÜ_QXúËGeYxƒ°!ë’
å
×Ÿ˜j’6Ã:‡n™çÐÛÙðLûy—KÁï gÓ%ébHWk Q
êáì³qØ,Ê¹ëU˜ŠlÔ¸ºbbHƒˆmÒÈVßQ:3*úåÜµS‚Â2¿;÷
j8“¤Ž•Ö0”Îo†åJ†fwQé#½Hö÷›A‚®ò×væOËŽ Æ—Q>DêxA¤žÄUl3ÂvÛ’ú&d}0?ÔZ±ï#h‡Å¨¨óaŸn73,O¦{¥¦@É¤k±àºßx0”_u¨¾qwîI}~“ •	r VWß||Ðzo‘>9ØŠüRMpVh¶~”]âPEÙT¬†¾ïãƒVËv±0p¿ð´¨º³)s³n<¶ó:yiNÀµ¾–Ð†1+Å~ø´¿¡þF ½ãÅTÁÏƒ„ ’\ëiO0K\c£ò‰kýe– *º¡ÝúÝ›»2ÝkvªÕ½¸³aÝÀàf ÏóÞ Â«|†-@ÏTã7{>°Ò³!Ü ½whOÀòâ™ÑLÖÕ8{–‹¡f~I¿ñG3Ãm…\,0Sp:™8m^câÅt´Kÿ®«ó l@Éñv–Ë ýiÞ/]–ƒoß‘™4r´}¡ELÛÍ ùx5>;[#KÉ¿édw­¹µ()i°Z7 XÏUÉ ¸ÊSÊH˜kÈ@Ï×òù¬’à,Æ®)Væñê€â7ŠbÅ5ÂÃ¥aX¯ëC¸*o³·“IÁÄ§É³, w6,Á+sÅ¿#‰3.lß/¦£¼Ëz•‚ìSÌWfIµcÆAy[¥<rH¨Ñ'Lš1Ÿbš—Œikâ×€‘ÎÓ »GŽnÓÐembº,ú7BgñõMé'ø’¤PÏ•k›rþç9Y
Û0›UggÃÂ¢q×8‹¢Ù:J”Z‘%±$cýˆÃþ®Æ•³¥ŸTŸuÜ¯²+$_á5^õõ¦ºC8‚„Ç%ÂücÄ/	
sš&XøÏš`¤C–8Éq^ÒÚª‹£áP/ˆøi4Þ+Èþ¹ìvIn$!V® £NYÙDAÇ~ÚæhÜ(Ë¢Ç`ÐU´Nf…b¥óÉnY³ÕŽŽÜ3vÔy{±áÕu~Ùƒ¬Å]ádÿ2Z4>ž9ò7Í†ë<»^=»Rjì5ëV‘»÷¾‚\!{”K\Î·Æ§%Š‡ð£¿IˆšýZ.#„@ZW¸+ä[I
Y÷+”ä¬¼uïšPÓQõÍ%ÓK§á3þ–y›ÄŽüöE’½TqQ^vÐhJÂ¹h*h]QÏxLõ³b&užÔâŸŽÂ­i%åµÁƒîzóÈÆ‚‰QÏ]VtýÒj«„Û˜–É<
Øâ?—^4wÄñíŒt0É¤4N2mA<8®ÀxÕuÝúê!ÌŠ²†—šÏ@P™Øs~$"È‘0Ñ>'ïŽ«^dd„±6EUì­ Í½ÁÎ{‡}ßÎþBë9Àf¸Õ¶*ŒŸ‹¡ùÁC¿(úûWjð!CÑ,¼r€(,MŒ_Š‚3¢`”½[LT$fdÙ‡¬ ìH” Oâud„aº	Udõ&<#ï;„áŠ¢{60Úa1^Ì:þœ ò€üc¹3{:A¸]”+7³£ÑZòÑYs‚R
EmÙ5Ø³ÛÙZ]ž¨«†cDS^ 82ÓÃSË¦Ç^ê	O4ÏÌ]ƒ8Â‚/HÕ¸®2™‰SNö`Ä 
Q[K£Ai6Ô/û5ö`;6F÷u%hf«,yƒüÉ–—-@AYuÆÒ×†G¥Ëûõò/cH¿Ù›Îòz:î9Qko°í×¹ºÁyÓ_³%ð;¦t@DŠªž‚KÎÝ@×Š>C*¾Áv`ô¡ãñ&Ùe„V½[ó³YÚ†Wg¤JL{¤JÂYZ/nÌ\²tjEB€Ø%@õ¡¶R¶Û™­ÜÔŒ¦4|ª„V¬ bËïth}íÕjßÜäÊêfÄY¢9©vð¶£/Ü¹›Ú¿ÅA mažîÙ¹ô<r^çz•*¡^®yŸÖ
hÌSHfø7äúz}­øv†$— SsæÂ”ÖÀR,¡|¿5Ö)…Y]`Pžtà¡a	Ou]°ø,ÚAI:œ¥Fìµ±´±ˆ;±ÈPºž±üyï–[N;+i“À.YîØRlBeÇn.sî!Ví»ÚÕ›Wé[R<mUÑ;4í'´yCan&¢ö1XVèJižR ”Èr¢ôORÀoÀÐ­Td¯eJÇÖ;>‹ßCÀŽp½{Ôƒ3½k(j å	áï ¶F˜¢¸‹ÌLk™úu\gÝ+]‹x?¿¤È•ë>¦#Öz‹ Ò¨fÐ1@ÙD CmÌ<ã¶ø’u˜[Q²ÇDÜ¨ã«†yì¨£.äåñ:,ù½@µ!Ž¸*fù÷ÿñ¿þãßþqµra6,rp?X5ëE@‹tÇ1Î+WÃN2Óô›ì¨’p&¤2HÆG–¤8¸`	7òÚw¼wYrnSw”ÆvÓF@Ê¸Çhgié»º,N©ÆK'pã=ˆ\’DTüËöˆe1„$Äi¿æVä?ÔP{+¾Ê˜C`¤e¨YÂM.sÝI[Ïh6ée¬º`¦W§ôz‘6ÜÕ“®\Ç—+º`²W˜ü6Í«3
Î!oñ8Šc‰#Ûþ†°j¹x%0+úY¦p/á[/vfZ9»$¬>¦möÞ[£O‘¡x±][Ñˆ*‰ö—°3’öÆï–¼4ÄfSçMˆ8ÅÝåáòÞ+šmqúÈi£•ê5Jcó¼rs.OjÐ.¿Ùˆ;²ZP{D† œ¸sL“&î‘©ÕDìî"sN`,q9àæqéI^—…JFíz>>!u;¿h:©TÌú“ž<+Cˆ°´ió<W‹ÎüJKnüñ$p¦ØUQ"«\œM6•ðÕZD†&2HkPþåDü+-Oõ—¥ö·õR‹dà”©¥NÐŽ!)+:~¼FÐãÇ‘­N{”	üK„u}G#÷g´Ù4={jÄ5âÞôv&óvA²‰.¸ðiwŒÙ(¿X¬mÝÎàötX“:õáI: yð€‡—ÞÀ3ÖH£š|VX˜šqÙæàc‹óÇg"ÎðÎÆú'ò°¥˜å/Àßð6K5D/„–vD„or
Ñs‚@;€Añ£•HÒ’6¯¤VRKiç'7hÝJäDÒŽþ„¨Oš–i+#ô!rD¥÷[¡&%‰)&ERÆ(}ëOuqºo4^àj¿óÓñ0¿ãa]÷;ãªšãb“®Šºö›1²’¤Oq2ŠˆpZv$-ëîÍÄ#nB‹4‚¹{ÃÙþíûau\Ä<Ç´ïéî‚eJáÑy{aL¦à\*~Ñ‹zB/_6#»döÀöÈõà‚ÌaœivÕ†8F—ê¦,lâ10K3.³Ëªñ÷€$dÃ2ßqvMì­ê»ŒgDò/Ž–›çË‡ž>8ß÷ÈûyíÏšd|Ná{Á:;¢JÂ¸6só®¿ç…c‡ym­ëþšíÆýÍÕoúiY.Z–†Ký¾ rë<:Ë‡‚PÌ(¸8•Â¤¬%„;ÏíÍEÞG6«I¡¹¦¤ÏëÍÍ–çÅR×·}¢¥iŽy›ùL%KB©)<­V˜_µ/iñ©Ÿîì­Ï©Óê° tÙº-·Yzª{ÒÌ=²"yÓ6dÎêfÕW'Š<ç»8½*Þø=:uÜè %øW`x#äy óº›÷fÇUÿRï”¬A¿µËŒÿaÛ¤áúïFþP–åã¨Újãç%àëÈ T{·AÔå‘Ž.ÙWÇ	Pq\o„¨£"²DšÔD-YªÉÞ¬’ Ì“TÆªís1äÎ…¾Œ21]PØí•TçCV¹ ²ÒÄ‘•fîˆ¬´rJU}®‰IÕ³õ`ŒYK²ñsäÅ<H}YiçùÈJ+ÿGVÒ¡ö…d%†t™âÉJÄ)#¦o
i‡šª‚½n`Â,‘æDò97K3»¾à7Éænèœõ¶²2S÷Šq5c¢¹YÚŒ"JÆf.°²ã$¥r#q(rLÃ¸SÄ"'„ršÆM.¶ßõ:ÆNhœ"Ë”<Í†•Ön°¢6sZIjs1¯XsdÉ¾‰­N²;¦¬6ã$YQ¢†I©;'ÅwV”…·PÔŒw†Yù[GRknÀÙÐ4mó6Üºí<ÜØOÀŒeøºÑa.Ûß-0¸†>obxKò{A­ï›Õ2üß ,ËJ+?81¡¥øÂA‰­$Ñ¥½_”d±&Ý?Ž6n´ðÜ2rwªx“æë–:¿˜“ì÷F;¼!áfÙ>pÉä³‰7”%0bæ—r ¾s	ßF¼ë¢-,ÕÒzY^x´×D{ë_¨’æx“®ß„ƒ"ÙBq¯Ã‚zk	&®g‘Cý)ôâXc/RÜW ¹!©7Ýy–ìàÎ£Œ–“xÅÍAÏ2Ñ^|ÓIËª*‰ú)Ëî¥·èÊ|2´qs°CV:=¤ã	¼žÉ‹ŽÏÃí|‘¦®)ž5Î55ÚÀºX‚e->  ÏÇå€ü&Ú:ƒðíHNñí¸47¼1Ê©¶Ö–ÓÂÛ¯tûúÕÌ	 ‹Öÿ(6ÚMWÜ½ýmŠZü~_l‹È*%„žj|jQ+!V0ñˆõaó°áYØŒÝ­ÆÆÎ|;ºMN¸J”ÔÙ&A	¶b	Ã>ÂºÅïÛÚ…„k®Ù¥×Dq½å‡†KXr#<Ü§¸Þÿ  ÿÿì}ÛnÜH²àû|Ev¡wPê–XºØnwµdY’{tŽly,¹g>†M)‰îªb5É’­Ñê3,v±`ƒ9À<öå¼êOöö6"oÌ$3“É’J–»Mt[ÅdÞ32"22.ì¹¶›8Û¤ß°«8÷¡ðñÉ7;ª7ëèàcö5o˜g¨wjšo=ªú¦'›âT”ßÄ%þnïò,œ„í€DlVòSPEWƒôŽ @5F/V²¶­4,¢ElhêBêèð:>Ó?jz8HcH§Ôtq%¿mÈà:Ð¬¼¡Wó¦cvù9¥€=³N–6.ª)jnåœ¸¡Ë<–y–œœt©¢R˜/—ñ$ÍŠ¦—òþ¼:ï§,ºWzê[ÌÉÓ$ŠÉV˜EyCtŸxwU>ÕãÎá3ÉïÉä!©npX™¼ZºiªøþN×wxYÕ|o­._3tòé=‡n‚Œ„çT¸Û‹GG btþó2†Ò¸˜o½wzÏÐº1¤beVð"ÜØî&B:bš\ý#—­ÖH”Ð˜:ìÃ:Ð|^èÞ/ÓØ¤D7`˜Ü©Ñy8ÊØ`HžÂøHJŽãÁiÈî¢0Êº†Ù\^4ëðYÞ/Ï@-. [€(ÝÍ[I6Æ·¨eø:'¤2´†pºYÄcS Í U™– úTí=$d£?²«¿fIÈ 5„¤“ —–À@lÀ"1P3€Î-ïZÀI«– )¯dï(`Òþ5%Gš—;£8;¹ú÷1,åõ!SNNK¸\¹_ÌtJFa’ÎÙƒìêïE29t&Þ%ph‚r<Ä°0È¸ä5¡³–Pá*ö’¼ é1a®ú’8¯8ëss7c"ÒÀ‚`h»ÙFcŒi¢«h×7‘îŠÞ¨{øP€ç‘ÙñÐKTž¾úG–ÐÅÛän!ŒÆ±¶nu?ó>Aâ*q°]:L ÒZÛ–ó¾Ù{¡ mÜš¥4´œˆ³ðkÚ.¶OîÀt 2À¿¹«ÆHv}1[Ï…Ïø¤U·lð×ax…»^ç@Yh *€ã»aŸÇãÓé(‡1ê<“õ´ªFÄé"ü[.ZC&±zµÏM†[äfL·ˆÃx‹Ü)ó-b1à‚£š4®²™Ê*JÒÿDJºO·Zh[÷cÑÍ[—Ä3=œ&OÕº0Ð@²ÒŽ•©ÕÌÉV\‹z„Ù‹µ« …ÀÛa1VðÙE¸¬E»ÐÇfõeY¶È™mqÌeÃ 7†Š=Âð¹ž!WYCžç•ÂRCæ»³3V˜‰o¶ñªzƒtazâŸÐóZ8>÷
&ö/± Ë:K—ôzc‹xÿè}Ÿ`«®ÌØ ¯?Èa“Œ£¼ÉFˆ×cÁ ˜Ø@µò8Y¾~`1{S3†øöPÏÃç]^y€‹
|mwaÁw†y.QA‘RkÌûwP •K·3)–ž¼ì8zyé ›B|\F„øxâ£ìx?kÁk¸20ÚõÕC[ÖÊUï!-ì¯¯’R³©—>GKk›ú8TTcpeïPøl=ñÏ}³ÏÑ5^Æ¼nÉ+ƒ ªÀÉªâ·]uyÎ–&Šï¤á‰Q ‹¾y§¤úÓÊé{cmÔ)¼ärÅY¤©ò¤$¼±º™=¯Ó¹÷ÐŒsêÐí4Bð™Awˆ¶äï$ ½æ¡¢.ä‘—”·v^¹mjG¥€Ê£ý6·>¦×4˜–Çg~*n¾h£QŠVÕ «Ã›ÞúbøÄ#"‡NÎÄaÅcT^ÞÎš4yæµÙÚ¹RdO³CÅ™7­ÉI ¯¶—sTÂ1 §_@ÚAOß€øpÿ€~{Þæ"PuXó
h¶âT¼
Ú¼®yš-ú¡ êäÏo˜¾û¾Ù×«íÖ°C´°tw@UsÒ;žGy¢¤Š3ƒ,†ŸÑfñ)”Ÿeüw-™Ú<qß„	70‘ÑÐÃVŸàJ7o²‰ñÓ+žÉhe*æB[›‹9jš\ik1ÅümÆÚZyxüP°·íX;ë±ëÛµaC=ÙMƒÝ¸FcmØÎO…ŸÛhKÓ*¼cÉ ú´×>Ö4^ú˜ø´ÑÉÄG×ËôZx¦»ùjQ‘Êp¹ç¢fÖî-6«è£Á^Ê³DâJÍw³b*¯éŠr’ 2ÝÝc8NFH?ÐÉ‹¿ç“êS¿2K7‚uþßÿúßÿSVê–âã§Ls~6p'ÌìÌtõã:´I}š[5917iÿ÷üWQåoÎtçó‡5c_Ü¤»›oë[8:C?‰º Û=§Ùg	o^~äf·máe[Û·àÓhã‚Ïv.,Éâµ½¥){8Œám¥£É0.<|[yñ»-ÍghO0»1½Ý¶ë:›l<¬V;³ZŽ¦l§×^u»„ç)jwí£Â´hâ´GHÕÌÔ$a’¥ÇÉ0~dé0f™6£Q2&ÏÂVÌh›Z“´E¥ÿ¨«°A5ÔQ¥pRÇÉÉÆ…ü©›r¦Ð ZpÐú·Í( h‡™Ú2 ƒ1ôÆ€ê3’ÒRæE˜â$+KÝ´é•Ím0YßdXþ´-/×ð")k­Êjç›/ð’1ôei™ü9¢Fwçx®¢&“áàg¼`>‚¿Q–Nðb%««•¬RD´\Íd³7..Dt¦>Y^$9^ÅÃ¯àûûuÇi›VfE–Y©ˆ?&E«š4O¹†)üâzhL€ŽZÕ,ÞÎëï—ÏNßÝ9¥ÖH³Šõõ”Q]š×º"*ã%êó!­/bÜÑÃÇë<¥ƒîùpïŠ"¶€I%Ô ê"îMÄ¼B¼+»ì8æ±³Úå‹&jXoÙ¨”õg¡Ýn¦J6*d‘cÕäQ(°™$ ë(n6t/‰6:ü6ü)ä0ÏæÁôh„[ˆqÔá™ñF²½ÂÀì¾ç<µ2*[5•©uYÕû¨KÛâÔˆKD‰°J¾ÜÀÒkô(lˆ•¼Þ£Í[;çvP=¦ýE=;×ÅÇátXü§ tú6}Ü 6“Å¿L“,Ž¬Ñ8>¢ˆí\Ñ÷°hæ¯ê.„Wþ„ÓiA¯ºŽ-ÎÎQà€ªóÌPEhVp«M€¦Ç§0à8Ûè,/Ëìÿü¿¼ärÎú9Â†=YßÓÂY5Ò‘˜Å.Vmÿ¶a¶ASÙµÌÎºl5UáYá·îªÒáòšP)Ù¿OûÔ¯§äÉr§óˆZ¢CJÁz}õ.®È¥•¿[W£iq?‚7©7Ý¾GR|ùHþl]‰.´{¤I¨š*[ï1 ý<ˆ‡ôm3;e¾©9ªž™¢hþÁ¥;’[Hµ†ZêÀýþÛº(s²þˆþi]øI–Žÿw±¿¿ªýQÆzš#‹Uzž™PyøÖØ­}²Mù¶_$?…YüÞÑü¯›åúÜN€Z$±	8&òÈöãb*®¥Š³ÖYÁÔËßë—£áÁÉú,#HPŸ¨v``Â³†Lnˆiðºú|Aá{&/`„ƒdIW˜#„‡Å³®¿(o%ÆYÜx,Öp@9²A4ËÚG—1Þä¥~›7ÄF©FðnZ_^ŸÔªú\`øsÑ›âõîž¨l‡ú4¾ú·y@„ßä™©´ð»ü«‚©O³ÞO¨ßè9.6sL=ëRs·Ö_úú¢7ªé1Ï›ÚÀÌ’¶‡äŸãBÿz%{Éøg2NÑoî¬×*> +ò361+P‰òžRªÓ¢˜äý^ï$Mƒ“ao%{AüºÄTŸ`€³>‡1õ#ðÓÕßÑ¶vž”WŸ¶wð¡´„ŸËüê ÇÍ,•Íœ¸YvF¨¡
7©ßJGqFýêöÈÓíŸ,cs@U'7Æsoº‡b×ÒO·ìê®ÃéMíh¦øîV*¨¼ZÑÍáÒÊ2ÍûjJ®²_í7VŽã*»ÌR\-äü¾™z†C/góØm#šo¤.†I^<M3¶4¨Êü4ºCaÿÁ4‰.œ°7·:6dBÜvÙ`;{Ñd;e:Ê¡@§AggIóØ˜·#w [ó}!(ÒWèFp+Ìãnƒa¤Ë€¥y™ìŽÜ·‚udåép
; }2Â±½H'K+½U²D­èF9§	ÞñÈkA¤:tû]ÔÜñ©î¶×ìÐÉê»©°Aè¾âkÜ[zÌL‘ÌaOÄ£Ðå_è¬ÄÜq^”îŠ&\…ÖìXÈà®Wù­¤Uë½*m„Ãöôçb:N~™Æ¢.ÅYªÓ#Š³[Åg®KÛ“»Kå–ŠSTôø½…˜<'Û	J¯þ~'¹Í”Ë.ÚågÃÎ{˜úÚRÕnRZÚ6]«è¹¹vfF\¡&G¬kìY³»¢A“¡›ÈÉúÊÍýâÃîÂ§ñàç£ô£!Ó\q´Ñd
&„a×YŸï·Qãƒž%SÄGÔä+’JòW¢÷]}0ÍÁ˜|œEÚÚzý¦¹~÷<z{PNôˆa&.îð†§ð¿ã†Ó*ÍÃgƒ 2@Sô"Õ5£ÃK7Càn¨áPYGÁÜé}k„BæN5É·ð‘ú†ê0œr£uzøñlˆÝHGÔ @Ü»”GÈ?â‹õ4ôeÊ[ExÜóTÍÞä{—ê5>nÀñÂ†ø´Áˆt"V”+Úlª`-oãZ6c˜€Ù–GALAÂ¸c¯âøð#Ë”|Å8Öâç¿ÆÍ›—OŸ¼‚ ÚÑEÑØS`âÄó°#ö°"¾%‰o,X±Yý<55`5‡®Í—Õkå¢-Ö»æØ%=¼0žÍL±,tŽjEqqÙÀÖ5YÛåÔÌ´rØñF#²º€E„Ü”NÓ˜ìàšÊ–œŸ©ó;=R':=àf™Ãñ‹f\õµÔ3Ø4„Ã3Œé‰Ø"d®Ý©a#÷c¯˜;šŒ†}zRi «~©d•ð´Þ3ZþÊì È—¿ûÝñtL….d+¢$çe<šb¨ºA˜¢	u—ñ}±¾ÎC8o&é"E"ì÷À´ÓœºÎˆÑQ.LÍ£n‡ùÃæ¥FÓa‘L†¬Ø3þb/'B·!3ˆiÖgñ(íª´€å:ãLpòÍã§Ã4,º¼§AS©r·³ØY$ ³°€šçË?¨ˆ.é5ˆ^Û« uôz¡ª}ÈcŒ§”<Á ßgWÿÁk†ßK&Fåœˆìõ7eûØ©ËEe–EÞ¨³ÂTFàð˜Ášž£Ÿù3ô1?ž¢ÃdevxCèN}w\ƒç4ÃSZZx-_”T5/ÎÑÒ¹3àõJÇk"¾=y¹Ç“/Öl›.5sr®±Y5§Ñ¥ù3‹á!?JLctœ­1Lçšè`ýtµ–}Í×Lyëê¯äY<öº)ÅzïtUi¬.’ño´ŠaÃÁ”†$eŸáOx&Sæ4e@wÜdj³¤À(9JãRR£cx¯¸G¥õ¹ˆt$NÏ=‹a’{,¤:9y(WDwì>Ë}lÕ5¼MpQÍ*„27mûÖð±‘1hqÏQ“-ébÐ{¾bÐ†HK„¼üÚ@eÌ¬•ÝgÔ6‘iqÃð”‰1^½°ì>×-ºyé¢ùTc¾Y•·1k­M1[ÞÈè|…º,µÞú¬ó	àw+ûÀD¶Ônšw‚”\€ÄÁHP¹: ˜ HËf0ªÍ ÷ÅZ.ûíñÝŸžPj	äþ¡ûoŽº+yp•Ä 	Pxr#­v°Û52
ßiv%‘È¶øYÆx*eÑèvòØã¨z½bÆ•FÆäò2ÀÜ—\º»1táBgÜº’»­®9˜eùSÓ{Xº6†…6]­<½ú‡‘G½úo»}â|Æ×F¾ÅÍy¥²½ÐºÞ+©5´:k¢I%¬ƒ9²¨W°ãJˆmêÀ yÁd^ýÐÐyêh4KßˆwAcà‰¥wIÁ$Jó@èc¨¬H˜Ëx’ÁŸ³ä·ÕÿÉ§åã0Ñ‘TÃÀ=Rê?NÆ¡zÏTg2µŸ•3"uÁE…xnà>ºðH Ç³,Ç¨æCp‡Y4#²ã0ý,þ=†]‰?e\š/m'!þ‚Éˆ¼ØìI
cÁßNÃ"'“gqž‡'¬RæSìOüËa<š ~£Ÿ€nlÚ¿ŠÐòÂ÷ïwÆ_˜ÛYÝ=Xú!$mN|ÅEÈ€?ž‰_ìÊß~9ÙKNhHzZ\¦ëBs8Ý½~CÇ6‡çy’NãQÌ']ìljXúv<ŒÃäÁNÓ»ã÷ñ  žŽwÙ'ÊRö	^¤¿`/ä¿À¡r8üA,¯öñ5•5ÒÕîS…´WE2d‰|åû¤;êv²_$Åc8@æÓ†èé@Í8ËÒ¬C©úYšD?8Úà«‰ÂKŸbªq‘³T ž>ù1œ(Iúäûqæ?³/®úd‹ÿäuSƒÊéß­pR Èƒ}P×Ç£=ý¥í±_úç*pö	»M™Â²YàµÄj hp]LŽüà„c,x›ƒÓuºYÚALa5Þ¤›xÝ\=îvÚFm7ô«îéÄò‚sÈþPƒàÇd!±mä8iOyoèïR&9EBì:	|ÝJ£øqõÛ/™%ùU6¬¥²ù‡ä£4Ð³ôK1ú&o1¿¼zutÕª£´´i´ŽÛGm¹=>ŒH˜(Ñ›Ð5€I8íkcbúC¸àˆëNÓ·ü;‡t»¿Ø?8ìü@Ð%åc‘\b1ºÿ¥Œ’<^‡TÚGŸ}òŒý` Ã0)K<xV¦1œÊÒéoö©D°}òGö[ìÏ/As3õ˜Ž½*˜¶Ot<*P\÷ÂÌäülôW¸®gx¤‘Þ£¬øjK"1Áæ‚ut^R#|òE•\jÒÝ
B
'<Ò$·té;§I^¤Y2HéZS´Mª¦“8¬\Rfxlpz&Š<ê*•¨Âcé)TéŒ±Pï!ÀÇ\G¢È]¾
¼.8Ý(ß:—ï”¾É7×ËêÓŸ–¯E'¨Q(A@‡eù^‚¥Ò«áIX¾Ó“²8r^ô&_•šxðL& G˜'‰ÍñTh²a#ï¯ÕÁv“öÊ6[ƒûÕc\¡žÕÁ‹9ŸxÛ!'_[=d_Æ@ò"g+¨¦h«x“ABéèSœ’×o	^Û0¥Ìöf­Ò=ÃµnîÍQ©ŠaÜ§ô†—]g(	?DzÏ0ZÌçM_6a¢Î€[§/€ñ§y	xfE˜è¾×#Ó£|%GpZM	;ðÃ ©Õ1vÝKöš±\;ÇÇ°³µ+ÔÌ)}çâ=³ Ñ”_KX&EuËfâ+ÆÃfª·ÝöÙÖþÞÞÎÖáîþóƒà`owk÷psksÿíÓý½7E\Vz
'1¬D‹ãp’Ÿ¦…¸4ÿEÐ»nÎ¿TîòY5¨Œ:<O¥®0¿˜°J.3ÂtÒ®¤–ÐÅow¡üp¹°€Ç¦:t)aYqyR8LÃ2€å…Y#GçD[’ù°ŸA9»Ýp‘ÕôØhà€oÂpÂ2\Óc0·¼}ªyEŽ|Šð[Vf‰5§p5LquÃuqÊ¶LÊýB«”«H¹vÃ–è·ngÿ¾U	E Ôó~g‘°
fì€™¸fbê%vÄšÜ#ÚÕóLº’(æ-á2åçãgkønì×Àe‘s©ŒO³ì{ÉÑ°Éã‰œ&Þ9œ§»P+ð,UîY”¯T}ªCÝv½HD‡”1´•¡æÊ8ôTž¢p½âc(ì€jø›\¦Ç"…ªÌËä~½oÔëõîE´#QÜéÒ	~c+é á,Š2ÒçðC˜ ö¡Q¶¡–²ÚE2Ùr•¨ ä«öÍ>ú•ù0fC­}unÊ\¬qØ¨À¥ÅÙ:aB‹£I‰}”½ÈOÅ¥6Ñ;b\H“¯/ÄHsþ	ÈìIcÚ~}ÁÆÃ½¡KBÌ{Ìü)gÃÎeð®ì®<|‹$>——ÀÜ«H·6T~¡¬XmcÓ¹À-¶
ÛÔreÄ8iv¦".oÈŠæ0Ö ïÑùù¿<t]*õu‚ì"ƒ–^DÇ‹Ì;³–æ¦ïIþjÂqÐ³8JB^ƒžh¬¢‚- %V¨Ä±à³ÙíÄÎ ÑúŸííâÈÎ0FìQMQÊT6Jøž?^/sµ-Š"0QG˜Œ©^¥0’<(ÀGˆ»7‰Ž;å’ÉÙØ%ÇáYJ™Åæ Ì~™&g)y±ý´¾DkøRr•™W™Ç±aB‡Þç{ÉÎÝ¥É½ãw;ôÃR¨lOd~¦GáðOiösœí³“dð¾d¨ê°Dã÷y0¦ÓèxÚ8¤£^ø>üØ&G9NGð>ï}}!k†™Cu—ô«2 ÄŒÞçï*È)Ì²ðüÉØ®LöžN½ò¡[Á›P«Ì+Û„å¬4E8è^ÐÐ}­òË…`ÂŽs¢28RÑ2ÔÕ)q&0Ø¤‹ŸH_ùþ¬o`+Áx:z¢%HúöÛ…?5oj¯°C˜¿›(¤—S›Ì¸(sC6Ì~ÙbßºµB˜éõ–—¨Lžñlø“£ØøÀ9v!xŸ&ãn‡t”êè¿…åý×ñÒÒyqõ×8BªK.	$üëøù¶lí[ÒùW¤(¹ê‹Ç±9í3í™PE&c*Eaß¨ïI}ÇñY²¬^å·òC·òå1vº†èvc™ñâjšëê?"”Ä“M¶û0F
ZÌ²Ñ±j#Á€­LNª$ &êùdÿðíÖþó§»?ÂÞJ<î,˜%N•)é+ òÔ>DNá	•ý vëÄå‡:ì ’´æ%V1¶…@6¬3)Å´C»ÕiK´L”é]2A_pA w0b"ÚËw5EïN†Ãß0£Î¶ÖÜý¨«Ñ;<ô¿Èâ38ûñ­|7(ÊqÉü#ÎPM2Æ²â˜…r¤f²ˆÙªTÓlDÑ~ 5u`A%e4ef%¯˜`Íô£RG£rÀä›Ú5ðû˜ÌFÈ×ª¤äìÎã)¥¹Ú¾’7 ×ŒùŸ’â´ÛIF =Ô tÜx=.2Ã¹“\zOG‹ÚÐ]meÜtö	JéóÝCþÝc>²¾rQ¿~|Sôy¾.ÞM§Ã³Yo.ZÎ¹pRe‡,‹M½¤	*¹ <S€6óm „¯^îuMYÒ1g³º1rT¥n2:ámÒ¡êÍü¬Te4àŽ1P=†PRÍÏÞœëvX†N¥~QüÙæŸßþiwûðPÃ½åe{¦?ììþø‡Cs.¤í’¿hÇéoS¦ÓÃ’ò\ìå‡ªÆ7.«ìÏo2ÚRsÉ1˜­»x£ßl(ƒícñ‘™ë™ªªVs/ì%oþ‘2‹æ~²†y7ùt÷ˆ˜%ë¸Ôü]­.0…Ž@Ú8+<lÏÜ%ŽÅGd§X	 œw;«Qqz €i: 9ˆ²ðÛ .‹Éþ£=\Q0—àÍ†eWŠTìZŽþ&ãÀ`ËÁÃû†J8æèòjj9¬.
Ò½Ïé&õ:;(¿ã7!µú´÷K* ¨–¢•Œ‰®RÓxƒ2Nî^¨¥.˜Ì¯ÍËRÀU9‰àÉƒ{•3þOCð”tÉâU[–^˜')©n_íð-ñØn‘L€“+¦á0ùòóã‘¦‡(ÖdÎÊ-êô\åk,!8Ô« ,Úsö[cÍ.È˜
§:°§Ùý¦ƒþä°Åëâ6+²¾ò]c/…"‡d0%×UvŠ«Xˆ~ñW½ktDð‹Ý” ÜŒv±« )zJžïÿ)Ž†$tÚ>ŽÂ¥c¿_VÛgwt²}öjh_NãêÆxºØ£¤x¦å‡ãddHeÙŸ…ÅÓd¬å­&ÉŒ›ƒ0ªæTÓªã`Š'bìÍ0Š")¦C¥cxS÷AïÙè8T7â½b?Y³ËbûÈûtßá8ÝD5Î$ÍüîóîSu.„š˜ñÞ¸ªHfvé
Ð,ÙÆŒèe÷`ÿ€âjxË'ÃÆagáõò›²ÜÓdÔ®ŸÇÿ|¸ƒ’ƒ¦¨G¦Ðõ‰§ì˜'_ª§±Qî¬îF|×vc]«¾!(‚ÄÃg˜çÒ,’•n>¼á²V-à–ãuP×^rã©‰ÚBéÖ_xä„éVº"Oz²[ ;á¹¨šÖðBMq’«#âó¯lÀ5E«JWÓÞ¦6¯[è=­¶
U%³¦Žsµµnž¤u§¹V­é¹|5ì:&Âl³YÂ›&ˆøæ8/"CªÜD†oax’‰±*ËYå{c<A.j]Õ$YIõCÆÀ8©Zºš$KW?À¬¤ùYMCJn”y1¥—3SÿjÄ)RÓI™¨‡LmjXw½âçJ‚V©Š£U^9JpÃq¨«ÂQc}#ö.s_>2£ú²Š×Ædá’Âªxg3ŠT4»«$øÜü@n±…êW?hG¹SSÄ@¦m3OˆB2a¸#	Ï¹€¢d§åãREåRRÛOÌ˜³«äV¿+¬ueÇÕËi”‚ÆMY/nÈ¶ vººM×óèo._ËdDs5–¬Ú`$.0‚Ô»ï(Q~5uØQ°šGë¢D8¦.òT¡Dï§£XùµZ®ŽåµÔQEˆFÀVrúÞTA%‹uMõ3êµ•ÇxêBI£Íª¬¦ª¾ œ˜Mòg;‡›o·w7‰^%jD¨§mŽVÊ$EÖ ÏÍX–l'h–gçÐxvÆ§Æd¨¢xƒ@‡-Œ"˜C³LjaQ“)AP)U±Ür„Æñ*ØÌã•—ŠGÔ5 ioNˆƒyðkþââŠX•þ†WëøÚÄ!±ê«œ¯×”lå–XMU.ˆ×dJ¶rN|M\’˜GÛ7¥‚CRT“åÌç‡
\Í2¡ßé"Éf”¬Øm0Œ­Ry:ð¹œ/…ÓÒrÉô2«äÀ´Œ<Õ‰eîkàÅ?þøvo÷ÇÍ­ý‰Ë^Ì‚ÿø#ÁâÌ‰K“d—}CøQÜ¼Q¤>jÅ©Ž,^…±NÇT¯ß ¨Ì{í¨dÆýÈ¬1[²b¡\HÆQú!àV—ÝÎ6`Ô÷@\'á	t*F3Ð,>I 0UÑzùo“Õë_u3X‰h\7À|]ï¥h›v¨½˜¼6Û|`71ÛTã‘¬œèi9Ó‹œÞ
‹E`õÈ¢i[Å!¦9{u°ófkŠ(BÙZ¨A®º|u<+þ*Ÿ2ë(ÛG—øiŠÍk Lqp¶Np”yk€–ó! ;þ8N“!aK£Œ¥1õ)ýc4€†]Àô[ÇøÏ$EkåUAóã8… TG¥Ë«Kê·”E5!9¹(´óW×÷A  7û*>©O“ŒZÃÄTö¨’)(£"å}˜!ª§yˆª$Û;{;‡;ê÷w__Ô†pÙûú½—¿s“*y3Å×BL»I½äÒÍÐ€7Çz(TkvxÓ:îí>ÿ—’žc””~!fš: ãfS·æº!âLûèÓ'©ÓöUu5ù¤½òºÌ›‚€õgBîù@ «»-<Ù¦xY-A_SŒ]^¥)éž÷Î¦uƒÊÈ¶z4¶kõ† ŠØsx³B¯Å#Êii†h>1å@ªw˜øÌp)ŠµºË¤‹0Ë}&>mî4Õ™óØ§
ëÊç·Ù6BîZqm4Ÿ}+jo»sŸî¿ÜÙÚ<8”[WTdÛ¼Oâ¹Æö•³Ò¸EÎÚÁbàÞƒ¼Æf”ƒd‰k'Ê!:Of3Þ¡‹²­ïÑñ1Ü¥ãS¿O™kwêêäÌ°ß$Xúo8®êrƒûÍ¶^ìm>¾ó²zXQ´qd²TŽ,blFô®iú(°dÑðÁ§®åÃÀÅ¥ésÍåýk³ LwèS,èÎËÝýíý·[›/¨©£ie¹žÓµW6Î®þŽ£t--kË‰&¬
Tø˜•¨ÊbU­)^Æ”lQ¨*KTÓ¯2|n|aæ	uøPÑ+/õf1«ôy%6ƒ|á¾¯ð™è›WúÎÂGzÔ¢M¨Þ±hKx\¡.ª§,|LÞ²ð1:¾’*­\u˜quû'ð³ûúŸöŸLxŸSEßE¦5@VÞ”pŒÆ}ÝVñ}žŽ¥aµ‘)U8~õr«àï¡¶,êcóZÖÐ¡°ŠÕƒÓŒZCÅ"%J?«€wÌqÇÛp|’…oéR½ýúÂ„]8Šw¢Öîg	Jf™ÁÉIl‘¢zxK–¾‘iPœU«‡dêOºíÕ+cl	{¨öFÚy@`ß¦:
3ôú«º®½x­°+ÔíBéæÎ§è’„dÓÐ_HÞ'¯™õ7yÓg¸»ö®þíc¤+|ãÎDÊ(g4-£B¤2ø`Wþ\è¼QùrÙÍ?JÙ£§,™õŠt5MGi¾ðiûwðÌÚCúÉÚ5”ßBC[ÌˆR7M 1yzõ7àçº,‹5Ç0æNrO<e'Åy’*½Qá'Xá¿F;™ôÉçLø*{Ç¯ÃÛŸ*î/¨Ò7ÜÎ¯'ÜUQ¥72õ–{¤øt*{¤]ŽÜ~Ÿ„g«²C‚Ã`ÙèEÜ'èw¯¥tJ0§Ñ§ÅÒµWÙ5¬ï,„ïÆÛï”êO¬ì×a™ŠŽo¿[Ì)ZÙ!`çäê¿qrË¤¦™6RgjJW…Í@Ü/}š#±idRçÏEÍ2‡5¿re··ÊOD|ûäd“{åSXÆÇEY8¾ú7;îÓª|#B€xÚûª Þð¨?.þ³B§;Ì£&˜;Œ"<ª…26ÅQ¢‘å sT¢£¥ã-<eˆ@hªÄ.Ë¦‰xàÕË•§„‹wÆèu—êÔ“;;=ŒÓY8©z©ÿúBkBøEotèê‚u§’ÝGu,œ¸SøQÔ©OÊH'J¬P}jr×¯×zùN0ÝÏ9]	
Ajžz¸¦%Þ’êìüBŸ˜òÜD~ÿ{Ös&!òç«Õ>v‹3ØQ½&z¦|?Í8à/ÅÅ‡8kÁ«N×ê®Ò=ƒðàC÷ý í+åQR¯¿wºViÑ‡¬¾QTæ>¡fÆmŽs Æ3mk“óì

™WBAÔ‚P}˜á4wÅ½faðž§g©œ^SŒS$²º³ÿ
UU¸€=¤{àç!6èà0ìM=Dq‡‘!RM‘U–Jnu=“!p‹8´ÚFù;ÿ^çzb[ï§íJí ?¼öÅ¶^<m_è0†9MÇ3ô’SÔúf'iûbÌÛjûr›ì<e.©Y=ÎŽ ÖôD­6€0±tNø[p|.˜—EÑ¶FÛË62Cªˆ ‹^ÍÑ"²=Û¢vê$OA;ƒt˜fygÑRxÐØ@q„´÷R–æXEïþ²¥Œ1J“%*|YÏÒ£DÓˆ×‘¼«µ|XzHN1èÍI%„L-”ÅHõFýã¯bö1wF‡õÅ3Z08³Í¢»lêJ?W£âxSUC\½pGuô6˜]Ó/@l§#‚^¥ïÔ©¨¹ó·Òþ`rLyú%CÏ[ëÄä}8Þl7ì›’¶ÂÝhŒÇ¡G;|UCy9`µªÜ¥¸R×ª¨?¼ýŠÓ5
ÄÇÿÜ>CH rGrNš†÷ÄC1ò$ÄC­l¹‚0®ÈïØXÌ¾`Qâñét®÷Ø‡ÖÁÙ­´¬Ú"‚~)á`|é¬Œ9\õ6öxÁ[™ÁêMá§ëï	”|Â­€ÍW÷J°†—éð³Øí7Ã»4fÃÍ»/÷÷v¸&k×™ë¬	ª³_PÝ÷XY/Š¶´Š-­è|]Õ×+Ë“oL'2+‹ŠdSkð¸ä`%GßØ¡ˆ=L â,Fea‘€åíÝ³®œéJ¾:ýd˜¦¼Ñ¡|‘œ¥VúéâÎn¼9qSŒAY‹5F¼x*r76³ø²ƒ×Måª^j„gg!Í­Œ®jz.]@“¸	™ÎÏçTNQàh‚Ü6Ã“Š ÉÄER¡rz>#@ÿC{~2[ÿ—øü%vAˆ’DIJvKÜz%G[0¨ØEMg&ËÌÊÉ•Ã±Hr–éÅn1m±ãÄ‘ß9½XÇê§š['Ì[™Â±UŸ|Ub.ƒàbZ$&Š°µ¾`.šáK5(ÝÐÐ½¶'ËD7	¡D¤VTV"’œvB9ƒ?nlÇù#¥OÄ‹½¢É@Y¹{YÖ_iHi/ F[¾~S…{mªs	/nU'Ù¦h’¸`M
žr=|ÄMþ„8F”t˜…ùéá¤$D&q­‰%†|(¬­ÞVÐ»wìÅÓž~íFóW¹ä#l‚dtŽ#’x|¿ŒrQ‚·ÊQ–NP~˜¡àæ/¸ŠbÃ‰ÝZ¥¸Ø´_~2NÐÐyãâ‚¤Àk%Å956Ê!žé–ƒïï›8ŽpŒ1…c­ÔŠ,µb*bºoä·3£ðã½’5Ü?bæêEc¥nƒ`~®·¼kÞA²G£ôF˜­ÞDòt~q³9”Álºè¸ù Dbš‚esþÙüZ²mç[öü:ª¸*JÇÓ£QR¹‡…ï·™YTTÚÚfª÷Ô=íÆ"(Õ!‚8ëàm,JËø\s–¢?cjÑ"ÚD/ÝÝ¦uJçÐ6bOÅ·ÕÂ4Ñ§4•€WKÓDŸÒƒÉqµ,$ù”Dõþç¡Pò©V¢õêÉix¿H>ÖºÃÓ}ê(€*Ÿdá¨Z‡H÷©ã(-˜‡¯j%òƒO-Rè¬×Á“}j`²:½8¦5—½4ƒ4šp– ]Æï¤ò°àÅËýgû‡û/ß¾|µiö¥O‡U–Ge?@üÎPï¤ø vÔÜ#3r«:ZÍÎe/Œ»×xð®aDq¯Ž¨ÎÎÃPM hRV¸.î¬h#ŒŽ–V¬²Ô(@e·	p™©…w¢[:•`¼1ªÚ`ÈÉzñ/ÀDZÏIC¿?1I¨º:®›Äº®‡Q}Æ¢žeÖz…<Yí¶“*zÌeš·!ÃóIS–™Úqíú[^é­OIwB{†fqÈu—©sëå·ÔVI€÷ÍS·¹
‘XÀÈ+Ÿlâ [&C@É§0åq¶ÑY^^øÿKviýoÒ(gÅ¤ä±æxŽÎ‰t<·‡R½7ÀþvW^¨ò¡ü¸‹t>ËÍÎ­×›–yµ_ÏÌt*X¬»¼¼@–—³¡±ÏÜ¢ìn’î3X÷di;>Kñ| ¯<gµ†>YôÓ!
ìÀyöþý••ïÅCº)s‰7¾úQœ¡ñ±?XÀNjÒ~‰^‘
KŠîø`Æ®þJRÑ$'Â¤Ð°ådS’'y’I˜…6‹ãB/ciˆþ0RgX€dr§Ážë?‘î5Ñ­‹H<Ýþ©´Eše8†Ø>²‚Ö»ÀKñîfö bs/­wUõ´%u(„'MZ2U5úÖ ×SÑV£E•5öî¾ã™ìM4ôÞ¦áãÒï¹3Ûƒ*÷Ïì©Œ«5Ì»ëæ‰î­°iÒ9³ÂŠ°&äúh.µé4/-!7¹	sy Ø‹GïÞ\Þí › $J³·lAâ¨³€²@gM´Èt©½ ˜Qéž]–Èîêäµ34®ÛŒ6ån°m*Žè½UŒ›Â$÷²ièQ}2ÇèW'mÎ+Mpß”ö¦ÕÐ•úc²©>|M¾j\”;<Íþ3býæPhùLµG?Ú²ÇîHHº Kd“ÅËFþ“¨0'¾KÞ_´&BòJä“>¼.fÐ{Ž~,rIŽ®¸š.´nk¤bÐeFqNØŸVo¦ Þní?Ûy¹µ»¹SÛº´CU·aƒGgí¶8 æ{b£‹÷†íŽÙèÝ	éÒßL¼îÒ§û•ã€-¼ô%/vÿ<wqº¸]n½ÙEÉ; dØzñt‘Ý2-–òštÊgqs‡ÅÕ?²Ä¢ÿ{÷åOxuOºÿ„ØW­ »Ûó†©9Ð>DÉ;#Š2Íß@„CÃ‰]äT»ÇÜV}lÎÎ—ÖnÚÏ‚¿6ÔsÊÔé™©¹qŽìQ(Á®éûõJ…¿ªc%iáRÓ_4E÷%¿ÿ½VG×lD!Nì,ÅÎUõ“Ÿî¶†:*ÑŽ`&B,ËPúÅ
Âlæ°’èHôŸP068ò¯bÂöŽÒñ´È’:Áô Ž¢µC¨æ4S›ËO¥ÎIw¨tâ3›Z'>3¨vâ3?õN|,ˆöfÕ<‰¢ßè4Yrh-”B?ÌfB¯í%'§÷˜è9èn¬YWŸ>Š€Ú ¼˜ÊºäfÅTþµ~ ¤ºJ ]W	­gàÝ6Ž"¤’ÞqzÂùGCw_ãß5ÌÏ»ñY‚d/Qäap-6e,ÒPs®i4^X°OPk²/¨ë mÚÊÀøx+Sÿ¾VîÄÒŒŠu…Ù<¹,êJúÊ7þ|ìžtç2bYwBwï²Ïy+}²dpz"bu»Ò&ßôœ"‹æsù”:ÐÂã¼‡¤¿Ntù´7dÄG	f¾WÌ*›NëøTËÃx|Rœ’uò ¡l“háx»1#Š˜9&Š`B	¥éxý9ŠÇiN ¤ÆyÐ(¨¤Õ1¯ñÍYçžhQ íÏ¥yó”u{FkôÂœŸ5Õ­]³=Œµ„)šÀTÛ`QŽãbpêµ½p’ À¶4ù°Ä}½¦»¹søŒââ4EW/ö½ê%½DÅYÞ÷lFŠ9°%—0´[Çîé¾é¹ôË†&P}Rq¾ïÛWðÏHQÓÚ³e÷kûÙ·q%Õ÷Ìo¹Úò.ëT¿óí,"<×O`¾É$l2=-ŸiÑ³ÙÐ;ž@±à“Ït~hBÕý|œ…ÜÓ"1@~£™ê`ÜZhTs^?lCx}Ð'
ýÀÞåqWí­G·dlZ(¢Ua3ã ®†y´?çÀaaGK	x.NÎâ¨ï	>êØ|JxŒ_ÆƒÎ€QþS©¶áÕ©ÇZ‘ Ÿ14Ô…cìÊêòB o£®K˜]>€,
ÿçŸŽ§YúwÕ‚mÚŸwûÓ YM³q:¥G:ÔZJÆgW&QHº8|¶]úú‚ÍÎe…€ìæyJCx™&W:H³,¦÷(9ã‰¨ÏÒ$g›#ìî?WX¡!Á‰I4IC Æ9º…Íâ)ÒBTM­¢oâŠCO¡´ œp"ŒÑ~–Ï¸Ö,¯…@5€^hhÌ&bžÑfò
‡gi¶H¦E I'¬S¦Îµ#Q/ÕcT‡„Gaò1]$¿ ‡wõ7göÿDÆtˆ$&"°Æš¡‘\Ù@ØÐGGÀKÁbqþÕ»äFî‡1±e ”pnü6{& ŒÎcçI“CAj1`¼æ	ë½ÖðÂÎc§õÐ*-éÇ¾ ‘zõU:Á¥ÕêÑ¥ýnâ¨]½G¨^õèS`+Ãns™=¿Ôãmßêiå’ tg8ô¡àæÉqÀÔÜŽó»ÓùŽ&/qX–OÃýíD]ïº?úSh¬¡PHGÝE9&Ê»|ùNñe¬\ˆˆ¬÷i1tž–ÅÃð#ðµîþÓi‚@*ÍãQRÊB±å‡’Pñ˜á²¹v›­£GYþLc<ÒWáÒ«2½L©2ÔI0zaRœi:ü9):>«J»WYÙQÞÏŠ2àV¡®«Ù‹pó°+rj¬ê{‡,¼úl2Bÿ<F*‚dÅ¡ð™ÓzoQ~ÜˆhGÉæ‹Ý‹qÈ·+³•ŽOý˜¦'Ã˜ìòe$‡l×{ü{›:ù$yHÃÓS.'BÁ5¼N‘ÅzL$aÍyÖ·5L§cˆ€G%Eõªla›£¤ä.‘ã û,7kÜWÿµ4ù~ÃP\LòQŸþFV~—nê:Z-“Ï—è]Âjpß Öm&õûmtN‹b’÷{=q@CaÞ0 WîºHxwˆb¦œÊš*ûœçÁÏ4
ÏÏ’øÃcºÐƒbãï/Sx[úîÁJ¼Òäª|˜ÉÒFçíÑ0ÿì_0øFgœÂ0DÓXæc Õvã˜ú£,]2&ãx©ñ2S¬â
j#¬Qç@TÔ‹yÃ{yU_b'×5ƒð ¹2ùøÆw4¾`CxlƒÒ ô³â“B¨Cº•åö;ÄÒæ^³ë¹ÃÏ{CD1•‰Ý[]kôV>¿Ž}q’ÅñXß,é×µ3¤®v‹Ò]3„|ò­Ôà]¿uF§8Ïz„nžç:jx^”ù9^ÚS)GãˆÊwJ6—^ù0]2ìtóžhplQ>š†Üvr‚›TSKèŽ®þ>NF©vh1ÙT® \½iæ9[8É¡mÅHúÚ¢§ª" Á6ƒŠsïUõ…²‰é´ øxì0AÓ3¤[«@æit‰O³*cùDp\8ÆÑÆERo4Mt…ÈÊÄßHè)ÑwYzMM`$UôÙ^`5é‡8jV+Ï¼ô$M¨·>çL4¾ajý”âÃÈÓÂGi~¦Ùi)­Ý‹¢´ó¨\íû6×+ñ0UJá’v_r¹ÞØbEkö­Zª½º²TŽ—(²Â+WTºÑÅ£½6z6=ÁÃiUü¢„qò Ìõ`;Xq~
ˆëg2ú(Â¶˜Mñg8¦>é´a-š}>SÓ¤È
eîýTÞ#Šãä
aÉ’Æ°×‡1éª"']m%%uÊßXƒß›÷5Ák~vÒ† ÜŒJÿeÂ•0kýoåRãì8Ò#T3Åß“ôãFg™,“Õ{ð_S	”²ýŒæ5ÌÕæk®‡î×'aqÚ|¡@ÛÝæaNÐ––½™W+K½ÊÔ²ØŸ’¨8Ý¸ RÜÌBEg+«äû³ÕÑ2¹w,¯Œ–ß¯=„—•µàáý`Áïá,÷—Wðà;²|·¶º´¶9à¹7X
¾û¾­­­ðàû{ü÷Zpï”Ý[Öî‘•4¡Ÿ‚•ïWÉ«ˆ¬ýåÚ< ­;ËË˜’zø‹º ½PÍœÅ)—\oZâÍ6ªÜŸ¥9“ˆŠÀ¾$&ÈuœdÌ9Þ§Sÿ$ù"	›¥ÁB®Ü0s†»}!–­ˆk›fØWL}õ7la
s£ïïŽŠáêd»äÁ×œí­NFåÀ"T(8KWÿN&)Nr‚}ÂŒ¤€š¦QÁY®®M½Ê ‹¦¨tãe9ZÙ²uM]>iØƒµ …ˆ¯þ-…ó,:í%ÎÐ”«Î£\ã(¥Ý™^ÃÌÓå$…ö£½¢^¼Úgf|¥v1l ¸p(Wé¤ nÎ€ˆ>:ÿôÕƒ„µ3~ø‰ÂpªGW1á¾ƒšÔö›2h4 *ÍÞÜLÆ>ÇMvÔdÙ\ƒ’¶\eÜCcÜCûC{îC@52…ÊwÛ £ïC…”¨°Ñƒt-_¹eôãF¥·Ä£YÃO‘éBÂ "ÑKºaé&ö¹-§Å©z†}.nD¿¼–üŽeqéòÃþPõ™üåýµ™<†r-M¦™ô}fQØá{…m”O­°ãüîúêáª"Çº¯ÆÛý®v™ˆB(—l«"ÉÒè—8ùÊ™\„Í­‰	nøµâ¢ØL­­öÂÖb•å¿\èV,‰“|3Š¾„Bù
ÅÛêõyzVZž¡ô&B¡ì*pÉUcÁP<o?$
·ËoŠ7âª‚ú ¨ÖP–âª€¹$°ÆÆÑl]UxD¯xé³d`Æ>‹ÅŒ‹UìõÈVôô¾˜‹lË/P-@9t&
¡x8™”’œ¥	^
œŒñs:-ØÙÚR7QõnQ›ƒÂíñªÚœ¢?%Å)eT7KîÕÅÄÈÞn60­|igeuíÞý®ª6¦…‚Ù¾`æS]}|uBF¡Ä0<N#ö ØÙºqií‡ãøÃ+X®bè šÏY¬„Åí¾Ot!n§Ç´Ïšk0ËB#u¡€ø6°×Q‰	ä ®o¸ Æ63gÜj«B™Nsf/cÀëÅzº™(Ijœ¤ns ¤Æ±ŽºÁŽë’E)Å·çëÃô$JËS4dvÌE|ÑfÑ§@g‡ lyŽ&]‡I%Ã4íJYâ/qû#{Ø%ä¾ùq<Ä¡Fa?m­[%â5#™zšdÔéµM`+$Åv:°S>FG‹dkoogëpwÿùAðê`çåÁ"GfŽ©W†iËd·am<”w€Q“[ð[õIepÖƒÃ›^Yé„™µå†u8 ÒOéþq†Ý)ùd€Aãp@é^¨dZt5‚miµ…ÄC7yíP>†þ†³p¤%']dþSXèw\$x"º@™þrìr&aYù‹×püdyéÏöt½j”~4øKîà:
sè<³
êâö(^
‡°ß¢ó¥d¼ÌÂ‚Srý˜t¨·~Æx¿¿ú+šxâ¿#¼õqí—'òLÊ®[HœäL6°2¤~	Ë6ß°l¿Ý8ìÎ¬ËŒpæãrÎ‘Ö~»k÷iC¦}‰{V{î6¸|¾qÏ~»kö‰‚…|	ûñ%ìÇÝ‰Gñ%ì‡,þ%ì‡åùöã‡ý¸Ca;¾Äé(Çò%N‡ö¡¢D€ÚÎÿAËâ õ$´9*w[—ÎgÏßDœÅF3S÷æ÷Œ¶xKN–o$hâµÂ&’ëNô	xmãÞ8nµ«O ÿ"\ËŒ í¶œöˆäÜÊÞøFÂ2»X¯ùFC™%8³=ïÝ‚#0iŽ€Ô+iÞë÷y¬DCÀ¢9®OS¬¢™ãás·–ö·Œh‹^æ950ç‡¨ü–™U¨t©Y|Ð†û0<âQ€â"<xÖÑõ‹×ŸÑT¼/&,ÃÆû{)nb7.øEëRöÀÔÈj~ç¦†ØÙ{e²ß³7º„Õ-fY5Ä±ÎçÏ¨gŒ×¡Ýbd!{u‚™Jò=È;ú¸N‹´‚¥ŒŽÜœqv ÕÕÃk\_‰XÄI{ÆæÎ Àxéy ÊQ†Æ(ôÇüdÛÇ®²Â„^dLCÙX²Ê{˜t•ó„§UÃXì§ÉŠ«¢òl–Í1¹/ÞA¦;'i†ÞÎ9Ì1Î}Äxv›h©Ž‚È¡FÓÙ6É‹=IS9/æÃ2ÁZŽNk¥àyy”|SwÔQZ-ž‰•`·âQ2©v]OtwßPÁ³†r†jäP<+ÃáPDrüÍ= µô¼©LÙeµ ìkSqìäáÕßãdò^ÆôUû–Ý]Õ
B_½J–ÖŠË{UB!üêùÒYM™…¶€ó4g‰ƒÐgÆZè½«Q`ßTW¹üj¼\°ºcŠ/oâaWÅ!‹l,o‰¼¥7,fgÙM;@«âºÁ(c,Y‘XešŸyªû»Ã]7Bì»â©ÊÈn§ÆÃ4”îvî!a'l*Âº/÷T?3øúÄ¼öàu›6Ã\‡þjRøîÈ2pç9ƒ‘3ÔÈ4óÎá`OŠNðq˜\$ø§‘É¶Ç¬Ë%œz‰¡z«D
!ð=¼^~ãÒÆý
³-4ÄWJè|=ÍÒÛIXf‘pëÜV··ù„Õ£µÀ²ýZêõ…Ë’V½ëÞ®ÙP™ê{ÝAf óf:¾HÆS‹…Œx8ï‘¸ÍP“)lDY¥¡¦‹=•Çè“çSôcÓ¥å+É›'FÆzèV9ÏÜˆ^¯‘ñ¬¼Ê§Ô§ ÂTøOƒ«b×Óv:\õ[y"ïiá<H}>Çâ?Æª4¶©íÐ5Ö™*ïÁrö¥>XÉíøÖX—Îyµ®±JKæ=^•«1lú*GÔbë[+6óm­Q€µ~CçÙÈlÆ)øXMØÃ,Â(B3`†LâP3Êx¶s¸ùv{wsaQ ü¿‚ }û­[ï¤4/hR	Òí°ÀŒúŽ«¦Þ/‰ñ·¤ëËÄ9,cþUg±tKaí†ÝŽ vOÒéü€æˆÍÜ•‘Õ6HE-rb£<´"f1G°°Ë/¼äèìDkÅåhƒù©mÐÀØ4P0Œ»A‘îìKÜ Ÿ“¢Û9ì, ogÛ:gaÝ~†ÁšÏÈXs×éº«Þy­Ä×ÑI¬\=hÎT'~Ž6%ÅpµÙœ©Nmê¸Ù¹Þ~9-XßzŸåølôV $>XV}&óó¡ß‘w‡è¾$JÌ!El‡ÜÊµºÉ2õ\ ø)€cZ4Œ7£ˆïìK“äÛ Gvëª¬QÔ÷Q]±]òÕëõ£uéšØçÚÏ$ÿ¥-m…Ã-í”svu…” c½ëb¾}ñ¨õÝ«ãž°ñìn×NwZ§”žÌùúRY²™Êz*T	‚ ¬~‘S€
µ¨Y5ú©ÜV®ùÜV¶òPîºÛ4Rp³;Ó¥ƒêºÄè@'¤ƒQ‡D#‰™/ˆzÍ]ªM7»þ«yCV®Î¹dã>B4'] [ïÐ1ü%Ó€¹m‡«ãZ¿4ô	¿³°Ú¥ïÑ¨.8“¯€¾ÊX,o­¬¦©¤N·Û–æ¤·m1U:n×4´bO=CÕù»õm{°Nx…· Qÿùp{¾%q|€ÒXÐv‹[c[Ÿ	¦*÷6·TŸjäá~¾°£Ü™´ÃÉ¤ÝjZN3V2@oòn
¬ŒÉæ¨ uÃÓõçø¼6)R®i:P.²=K£pÈ´™JÐ[(!|¾	Wˆ¥×ÄÞÊ²±dÌRG%LÆ\°m!›†nŒùäj@ntÊ‚ÍÛç­"öœ;R•ùK3äCÅò·Æ9ôhGÎ£G^ã\6–k5ŸBÞì9‘®ì•LÆQr’Šäo3èj@N+“qÎìZMG0ÎÉ¢Ê`¥\Ûc¢êæ3U^‰7L•«7rª\™ŒSe/ÐjªT¹‡s¾j"o›L³ÉPnNþÖŒàšºT¢¸¦œf$ç.U›¾7\w—ÅpbÝä¾”0ë	Z.€Ÿ~QîiñKlêÊa±gŠÀhvzÏäÞ*±Ä¯AáˆÍªˆ“‚‹i¤ë½Ó{3˜2UÄCkó±lòf"›´Âñn«»9NÉ&²³Iš-olº–*8õŒ)“e×ænŒWú¼fË†èÌ€6=¦MÁŸ2Ï×Õ‹·•‰ß-{tÁ‘3«ÀyyW±Xq•wùÍÛœåfâåÝ¢ñÄm8°“®PÝýL7P£Ï|ÇÀ¾lå¹Ó[¦TE£.JøYlÕö“äÖ\»jÀ»3ÛMðË¦Sž¶–`æXõ L6õ·­˜ÛNÌ©\nqzE'Z™cï0­-¢®Zî§“f“­5ìP“ƒpˆÂ7E´ÿ[T€û2¯—‘ŽMìg¾Ã®›¸)†jJ*ÆÝ:Ä¨¶Â#x%Ì–Ó:«eV5„Û.ë¦#?˜¶ ¢ZlýÖèêÊfù:å°Ph}VÝŒªšŽÉ ©# ´éÈÐ'¨œÕ'Â jOŠÓºM×¤
%†Ê¤Ë%2šÎAþ‚‚J}ç³éŒeôTÔ²'Æ{…õâ4##+²
œÉ‹_;(˜YåŒ®Ï°WÈYèbEøÐ¡J	ë½âÔ·„‹‡þ.€¨QÜ£¿ï/÷VíO‚WŠ¸ÎpzîÑã@r86@×ébãåÝmwˆË¬]]ÒEÉ·Ó-.vukgsûvúâsi¦Ëso§_\¢é„¦ŠØtîSÅƒ÷íýb—žsîLçÑ&cl@zfbê,øs½8J£sµAÀî€ä–Î	ÿQAš:ù±­S…ÎPƒïekùuS—Å'ûÈÅð6ßy1è“¥‡ux+q<Þ’$ƒöÎ…˜c1©­°c”ÊÀ0@`þÚÂ£Ââ)›m1]À_Ã±iä>ËyšÝn¸HŽ¨ìýˆj¥ÔïeŒ.ºÃ,î2U5«+&¿GÞ©Á‚‹YBi±¹éô»‰Ui*BŽ{p™»‚Ä=JðÛ0gÖ.Ï«Ý©{•¨™:ôðåÄ ¦hœíò¹y µ;V¯"þ“SUdñ+2Óô 
mžÃu{ó-÷Ê-
úÏ˜YS£MAÓìYK÷zäétÌaEJNâ‚]$2|ˆ;`—'9|-Å^g¹UÏ¶"PI¿EëÙ ]ì_Ÿ0yÖ"»Gd/ŽyiL\ÝGXÌu–;XíTy—ŠBP£Ÿ]ÚÞ:mO¯3Kó¸m…ZR†á]Á¥cýxÕNr…+<¿)¥ ”Dî@¥ªYq5î_omÙÇx¼|\â?JÎ«Ìñ0ý&’Š˜,fœ~ÈÂ‰3š6>Òô¥!#ßÇHâÐ’épy¹Oÿkt¨‹&5{”Pb3Ü²¦3)–ž¼ìXÝêñ¡[©¼ÌQŸœ
êq ”•™Uè»ÓXã›Öi²™òºîVõsôxô=”Ý^Ãn÷f¬¬bB¡.‡QÐMê;2l!¶wu:Y6v1ûDeÖ¨OÉµ½?1†ü–Ò§Â3Sš¥	$|pÖ‰sMÚM€ëu‘Û'AgwuŸÎmÝÕQßiì4+bùðÇolû¤z(àóùD×á1>vâ7Íp)Î]ÝóÀw|Èw	4xw4mö(©>ŸdÛß |Fû]ÁÞÕ0Mÿ9Œûî|/©†òžNZÕç“à€›†ŠÏÔ®Tîê®˜‹DàsüD	Þ‹•´qx¬>ŸFL0 ¹‰›y?µº¬xxº6ê¸ š
—¿†}ÎT~%›Zß±\-`‘«OØZ‰Çoº°Î;ßío>>¼tÜ=@ÉF¼VŸº#DªWäv¢©ö¶()æ&KÞûaïªoÄÚ}‚o5ß‰³Iñ19Wô,z]Q·ÉW£QÂÚbRŒõÍ$CÃÇæôqÆ	š­ŠI]ØÔbZôÚIð©» ¼†dCsSi8O·—^ÍÇ%mp¢¶—w–¶«Î3m§Š6xA­°ÝPfå+ñ1ûìœ•ÌÖ»O`ýéõÈÁ KÑ~$¥QI>JÓâtxîUüC2ŽÒAN«8L½IF‘NNO«ÏQ|ž%Ôqë]£f{ü&ÁèµúhüÅ
šË­÷uƒŠ½Ýýe5¾F50[ŽZOK+Ë÷«–v\-ÜoŸfNƒ»ª™%±©ålÏ¼ðø-XT-ô„ƒt|œd#ßR„t¶ö¼Iüq0œ&™ÄÙ(#37¢åãmë,Jè=ÁÏ/ZËå9/>16Ô‡ù\âa\ÄèvÝ» Aßëm²C#ß9aÉç{»¸2f›B‡~™Õ}â#¼¹·èHÇ`nÊí ¯þ¥7µè¬ê¿OÂñ¹?èÍ6ÎÆ„Kì3±¥ÚŒª¡®øç2!~tä¨ˆT™–T„¦Ü	*ÂïÆÈˆÃqE™Å}è·™ò°Ç)ÁðÁ ï€êÑ«nçMoF½jB%i%=:é/'{ÉI8HãZàPKOÝ‹íÍ†ñ¼½ ðº{“5sÍûê‰æ¯fæ0ÑýIŒ²¥5³CjPœ;S_]¾4»Ñ6‹°„«‚?²5Mn[ÑEÿø#Áüv§Žø¢zYŸ£µNÎ96‡l§9:G™ÕF¡Æé(†Ù°’…(AáI8ûäµ ãÿ  ÿÿì}ËrG–è¾¿"…Ñø‚n|èa7l‰‘”Ì¶(²	ª»'tR(e(¸ª QCs5‹ù‡»sÌ¢ÃáÕÄÝôòòÇî9ù¨Ê¬ÊW DÉÊè¶ ßyò¼ò<Ôr'AÅæ~I©06ÒÑ&YÈsó:y„ß‚ Üä¤ù]KÚWOµÛ?í‚g^uClÁÄ–yáÀ¢ndIÀY-”%!?6Þ”%äI®$'”³ÊoãÌ• gõM—PÊP§KH‰ik³>×Ë¯ðÁÁu'
R\Ã5WÇàZø¶bÌOEØš•F/<…ZÃ0J#Fo‡ÉÕ?œµþ2’Ì£V4q×êÁf»+]ýrÜ”¹ÂN<†s1Uàñ[Q`ˆÞÊ
ƒ#ÓK=¡/ï}[šjÌm]ÙRnØ’ÿãÆzïG3wï	'ðgÿ,ìÿxŸÛd Z'”ÐeÎ8´£	H>ƒ0¥gcð¼“Ä²‚®UòI¸Eq_Æ¨(e,íl óL¥ÆÅ¶¬b5gRU—òõ’„#@Z7kÑ8FJÌGóÑÐã[ Ð\¡ö°8•RÎ´üjÕJHø€ßáª¨d ¤
B@sã0†8zXôF‚Ö	ëß$9XìøÏnm`D©7Š¨óï†Pþïâ:it!f5/–Ò«¸u\"[.¿j„	od‘Š*‹kjð±r¾¾Œ¬þËå±”qPÎFˆ/—ÊRV}h–…öÅ»RUú¨éyO±)?ÇžÌËÓ(Í|¯9Êõ¶T•›‡Âm	(i`´¹(”Å5ÃOÚ¶ýsôÉ|Š•è“¨¨h®ñã^çRv½¦‚Nê‘í*P|ÍÁíîÞ¤àvý`°`²%]Á%ï+º]±ÿ/KQèŒjK¤$¿(IK‹d-ü##Ù
º³þ”Že}«¬ÎÄ>`7I‚wí(¥ÿ² e¹j7'Ù"jíöq4iby{»N©ÝòV*S2ÿ@©tè”¬¢À\šß.ÜÇò©¦Á¶î]ÌËX»¾N„içs`®¥™"Y)Ã¥G“2Hz4Áq^¨­ÌéÔ±Ì¯ƒEÉÅ?0ÕÜhì'…e¤ÆæÑ­òL÷^Lñ]}sþàvNÛEs“9n-{ÐÜ¡Vsê½¢æð¹Iö=ŽÝuÚô¸·×†}Í:Ú(¸7ÇgNj“·lVÑ9ßÙŸ%	Ð’çi˜<ä™é¬-}»ÕžEƒo*ífðÓa#¥ÐOi{˜ ÙœÑ‘fØŒNGAQÊúbªéppþ4Ó£ˆƒq4Ÿ…îº™„?i¨¾DÜ’æ“ò£4ü¦RWªÚœ²Mëâˆd˜Y<ò.²dæÝÃÓ«ß@[{¼ó×Fõ!Ä¿ÖÔNc¸½€*€¦åÝõ‚•Õ÷ƒ,¹ú­?òóÏ•öMcÒ|dP0½Ç4þòâàË ^Bà—•Í)Õ×ïÕ8×8!Û@Èšºµº4SÃ„Z7ç-I3ÿsEÅ%Ë8ºyš]·50ÏÈ±{&°½Ô]ši8 Ûµ˜7ƒS¾CÚ÷ž&‹Á›´Ó,Èf)›ñ!t„[ª¼ì®pQ¶zñƒé4Ä>X|à.vT~s6œ„? >\ÆÀGa–š–¾ÔU8EƒŠ¤ZÛ¶8²nøYaÂœïÂ µŒ=˜YZÊ—“c3.JÇÞÆE^æ”&­g¡5?¹ë%ÂªFô*ÆŒleØ§û’‘ØæùÈ¤ÓÚ¦ÏJÜšñ®ÖšÑbœkÎ)gÇÓY¿±.Óû•9W¿meM«˜¹ùyO?EÌbËæ8
Äƒ›[/p!cåf‡ÿàï4	ÃIì“/ü¡¡Ãv”ôGáæÇ†÷~`0TˆÉïxåPˆ{Âàßü}Ìà'(ð?…¥X.øi¿,û¾^&é`õ_'©E$pÈiÐ¿“ø­áÉØ«$ÃzÊ­0ç·À æ•3?]¯WNVzœ;go(H=¦Œ?®~ÕdÛ'U}ñä¿¸Þ=ÑFaœµÖÛ÷Œ“*D2à£“i¤*êÁ,ã>ÐZ;
SÚÕÏ„‰S°LÖë˜u‹¿­Ñÿæ‰aÍS€k~Aîõ1å¶	Me¨†¥}è,m]Ö™Ô¿eC6€(ÿyj2Ï5=ù¼hÇ€oð)¡xdD^b>\*tèµ?ì•ihbrhíè3ÓÐšCWÖè¥aÖ£ÛvÜçmêâµÁÈµ¤Š»mÓ¾§Ò¸T~Zµr[¤!'p•l\$ú©„.²v×!Ò«¸|]	†dìñòµiÛÌ*Bó)ÙÕ‹óØÛU¾¿Å2ê]YfxY¶º·qŒ¬ g„‹ÍE71GáðËÙö[Ã^pfäkÓ^Ió
&ÑGK§Ñ¤¤»7«ÛM¹ŒËì€ÎîRg‚ÂÊv$ái0Ä$UèE»­½®ÍàÓuESáa&P%×â¨ôïùö•²Wû@^äQÇmª¼ RC†Œ&2ÂžG'0µñ	Igú—+ †)`Èà$Eƒ z‚ÿ1âCûÖíže‡|vÀe×Ã[-ÃºGte²ñ¿-}J(«ÊFí^,™ýe·dHsÌÜnà½0é yjˆÌ  ü©Å ÇÕÛq4] õaÜPáù{ø3ÅfQ?È¢7k$W7Œ”ÎßÞa—Ä›[ü»6Jbä…,•ì&æetdÙ/¯¦cµVÑpœÓ)×Õ¢ènånlO×¶G}Æ¾%lõVLý¥]®¬ o¬WgÛ_¸EoG\ö4ô–3£~½2}ž¡³œãµ$˜Ã²¶F¾GS¤.,’Rad¼¼à.ìì¬íï¯ý)±ÖÌŠ†Y8;4éö2+¡OzRF>Ç.äo" „âëá NGQÖl´Nó>m*æ;ŽvÈé¾¾}AÛ½Ø|y¹&þÞþ^yùÚÙO‡FŽcóÈˆÅ•K¹Òb±ÂÝ«™*‹‘Tm¢
¼ìlj`’µª<}<wo[_“3ø¿d¢2¹B£l‰ÓZ¿"ùºtå’§ô^‰vÀ[íþYt³æúŠ	¥ñ\c·¯Ýwˆ”¼ªwMÇ¯‹»Au“ÞÛdÞ(¿åÚÕ–ëmÀÍÖôKY>W_Çª¼Ü³"³Î¹f„à|‹¬W‰&Ô‰©¿¹¹.©Y
©£bØoÕ·ÈEœ]|7·/`rš;*’(L3%Ó‘Ä0òÏ_ÙÜË¥Shvò~è'ï^Ìê¹øÁÚ…¼7>Ýúò‚@göütjç=k1ƒšNØ3æ7–™l/r¼7A,g¬l¶)¾XÀÜ/#– »úõzçv];û8/u[…ÉqpÞzK¥Äd6é#KÎÓF‰ÚÜ+\C|ö ¿„?È‚¹×æ÷Ü¸­oÙp3ÍX !µ<˜‘OEª< >œm\#])„Ùšt„KäèŸ}¬IE
Ù×»‘_ÉóÁñC­¡	cGy?ìã‡!b’Vá‹/¸y	×¥oÐ`m~lÚ…¤OÀ~þ>_B‹@§"YZÕêFU@,—¸ä\Ùr'4²‘6æ›v¬J ½"PÁ‹?ixo&m9Zõ&gœ±¯å»3äæ}àA	ˆLî•ñ¼%rŸÔ¼OïÏ¬¨¯­5Šð´èGƒëSëæàQ/n°ÞtÖUü"9cñ•I	P8µÚ(Œìo‘œZÈ¯ÂZ×…¼¨£ú†íKüWƒŒýÛÒkËN)©©ðdÎëâ ®5Öv•÷êÂv¯p‹
+>n.˜ÿ}g§s­ î­T2{‰WºœO÷æM~è%‰ºEšÉ,ŸÆ;ëXOÓlï—ÆiÜÝò«Noâ'½a@5HÅ#NÛ\šþåJÓì'ww¢À3Nõ××§šËçw1³Öºá#W×³­ñ~Ñ­“¢?oÉ ¡‰ìL¿¯Ç™÷|½AœÙ îÎ¼žI¬Î¢l6²^&d¢2hB£Ýh’ö©cô8ÈàKñg·àorië¦±HÁ$î"\Poìù»Á”>ó·_$µ=êt%Võ‡‰2­|e
2Í`ªF„iµ`%Üž-ª3Ÿ1ÖÞ·Ç	æžO“p÷Ò¨?mü(y;¾úï
0”@¾4‘=c?c ?Kx>%òîy‡lw»WÿyõäÑsrx´ÛÛ}¶½×}J6ïµ7tÌUÞ[›ÝyÌzöSñ
ÝÇ>®æøÄ'µãÖ,3ÆöÀW½q«&’²Ïþ¾›ãQ-¦¥µhs5j€€*@¶ð8ÁÓ¢+ÒÍ?Ñð=šÖ”fŸ-rÎ¹š*•fZ‘fhÑyJTÁ:×î$&ÖŠïi¾H~¬=2U(MB4Ÿ‡nÏéñSœý3¸  ÖÖ/âÁoÙ†ö+ŒYniÀÉ%ŒßÙÆ*ÝÐ…i8r-&>»ëàá+FLZ¿|tçõ†îgw}|á|r&˜fa³±†¤­¼œ”˜ŒÂ<ö`Ã"÷í=ëm“/ò%¯… ±Â"ŽNfø:aû­±i±pfèì2sNžLïE‘ ýgàŸxHÿäXû ~7ÁÄ–¶q¥¬§±‚ÕŸÆ?º9³&„#È3z M•$[d/Ñ…ü[k¾¶÷eÇi@	7½OngÇOÕÚ–DýÔœ‰o\©fâýÑr–hRfViÓ‡U²&Þ0{«˜UY•íwÉã½gŸ±-ï	›1ùù3>û4ð™Ð†|Æhä¦`´îvwç3J£åý¡4Ô~ÆiŸNcjÝÏH­VÈ•âÝ)¯·¼Ü:ÿzjé)+Ê¡8E»¯Ô$VÚŒÂÌÛ‚g{ Z] -R…ÂO1Ã5¼­á¦ð_PÝŒÁ#.Â®bÙ4¨X.Øó(Ówõans(»°™Á¬ÕZöž®*¼æIŸá§3³/J¸ó1²5òWC$•rosƒ{¶XEÌÖüiÜFhnÞ£Î”ô\¨Å¹áQÞÚE¿‹ãÙí‘¼Ã‹WÞj)šiˆ7Ao,Ò±Ùè–ô€K[äŸ]èƒmÞÄÁ;(Ï³´‘ô­!{¥-ðOsUý]Ro#W®ú¦ä8toÝ®ÛÇöÉÈ¼#øˆ}_m›ìÆHžÐ¤ï&}â‚é,yge
ƒ·A”xƒ.×;qßÎö Âàd•l<}º»}¼wð¬×~t°ÿjg¯»J8Þu0w.5îÏF\•[´;öiÐÊÇ“ã8pY¾¦LÆíi>h‚BÃ•bj~B¿EsÁKL[“Ä€A‚öë9§u	™õÏH3L{nK1ýÆn’Ä˜Þ* É¯ò&J##¥Q¦ «8±¹¥›nQ­ëU—dHÆíi]hrdQKç±#“I@ŽèÐ“9 '“à±ðoM­uÏKÀðíî»ð`8Tî=ŒŽ_+ß™Èå{Câ€#^e‡Q2n6vÂ4üÃïÐÀö$LÃÅåm5VìàçÂ,·hJ 'nY
vqD^Èo½ÓI8ŽßÀnÛ6È²ƒšWÆÙÊ½{K×àum‹Eîj‘RA¹¦"¥ÂT’p%…ÇIž-HëÊ¾ê7B˜P"_émüN$”Ú	I^o¶|V(­eLyÞ*<v)gU5L•†avÙz-’¨!	û0•OÖµ*U.Ïõ1ß<Éõq¾Ÿ@o¨ùiþMÔ¼çë5BÃ¸ÍPóš&dÇr#™ÍPooššP‡å•vïõ¸$¹"bæ7V^¬Ó	qïæE»á2Ñ¿ï œa¨„Môï ¸–:Tûg M“Ì\ñ÷gzZj¥1DPUÞ„ÚÆ9J{«`ÐÉ‘º.ÄÎÜ©¤AÔAtPàÍN/Ñä´0ãsïÕ0-0 _¬ò»¿|“ÑEìUÛÛ`*"On®oÞmk"}šLGK_zE/
8€Þ²7¹ú­¯I6^h˜š…{ÝvÕ©¡/°¶nÆ'»ø‡ÏiT­xáÈýõa¯œŸ¸iìfˆ°'M`™ª[3/˜êIÁÎÝt$³—zã€ÏhÉêŽ~ã .n¼I’Êgp[ÉÂÜ:UÖ¼W–?Eð[ÔlF9-Ö3wôG\Æ3n™í$’´S•5¶ŽñPû‰ÝKI7‹Þ,Ç8Æ•RÀ#™€99}€â wÂ c„rìú»<ˆjCÌëÍíÜÁûÍm#MÁÂ®ÌÓCÎ‰ÌÓø_	L`ž†9AjÏ7ìÜMŸÓù} yZv©vÆÔR7Ä˜Ê F-_¼ 
(^H^jWÒNã$k6ƒUrB5îUµµGh$nÇãi„Íú¥6e·È0´¾ÕŠ|»Iáø¡!CI"Éº±"F3„A&–¬)WÈ—vë£ópÐÜ°Å3êÆº1>~1w'çäNbÎ,}§Á”­È·*ƒå³nQ…)ƒ¹€Iú[œ BÝ	ÞÁ/ã šàCÆ0äêØ[zÛ¬;[ÌöáƒÊt­›Ü°Æ¤t‡;Ú²³vp’6ùæ­ÀAÑïÆÁyscUÝ{Ð,ù [ìÌD`ÏAÀ’¢9òØ2ÔI7`ö!šˆw¥àÎuÒ²‹2=äC:"7§äi@šO¤PÕ:€G}Éó¬Üj¸b¯Áu€SKÃ¦³âq4i6Ö+—âôÛrW¹È¯ôÇ³ðë ˜çá[œc×ò-{^á8L‚ÑÀojV}Ð¢3sïû¹”“q“ÅT3Haœ¹ÏyŽ¤(¡± Èbcò ÆòNÙBãï®èÅö¥ÒÌátN—ÿjßgP»e‰OÊ+:ñ2á½IwK0hÖ]õ˜”u”Vf|ûB°XY^Â†º
»57<9–ÍÇ{'áM©?w¹6°:Ó]*ñv¿°ù½ò¾<Ãz»dEcÃ4t&Úâ}TŸjaŒƒò+¿FËkú$™ó~íäG'™Xúµ–ž$ræ?²¬.£:Ÿ^l"E©¸ŠxÍ»“xÕ­ºœ,kö7SQ®ÕS¥(®«[ö^¹ï™×7¨nm´àgÎ^·-º(µü]Š¢³M|p´»Ýí¯’¡Û2½(¾ù2$_˜¡—#LQ¬.1Eñº}žn2Ey›²Þ¾.Íc¦æ¬ý½hŠRs.·ßn˜3À2×~]Xi¹>Eq¹úEÀ“×fi½‚¼ZÖñ*ŠkSÝX¾/>`Xr<òAÙþ!Ëe÷%Ïž];pcè:ByABÉ[Ê÷NëœªÄ³Â–ï]÷¨å‹Úüécm¯­¢,‘FzR&É×*':ž^]E©íßUÝ{‹*.a(üéEMŠá½—>I)–MY–ïŽV~WuQ[ÃëÌa—äm¹‡©Ë`ÈŒ Í‰°Dÿ²é(˜LÂä£p/«g ¢w.£._‡lÑ¥=­X‡øùÈðÎ®ßEæ}:Ât3ž^§×â	Ãw±ˆq=ÚÏñ>æs¼–1-‚<¡ÏžsÀÔ#4I
àL(ü$&¨*N¯®N`Ë©ðÑÿg/‡K ÖÇÁ¤¤ÀàöuPP=êAðî`ø·0üñºÎ:àæ¶ÞJJ»ÈF/<…‚Ö0Œ}Ž¨t&WÿpÔùË,H2ghâªÓƒU¹ª\ýrbÊôÆãî›‚aù6žR†…Z‘.ð.MÜ…ÁÎäÛ5Ö“®]5´cÐ€'ÖÌE†§@Ïm}¶iö|xä›G¶Q¥q:Kªž6ŸÍšËƒžUôÌeñ*Ày#ïói˜ÙrAÑÏÊÖn'ôÍ*Íñl½H­èAžh=Åë9ÉŸ ‹T±jrÅ«ìÄc is•—&}Ó·¹ìž›AÁ­–úÛ‰ô[]E†(!žZ±ÝÌ“’ç©ÕÈs9&ž­ÎRF`1V®”i×3yXwºlœ`N{YsqYšÂ\lBÙ"Ó°ŽëóZ1×K…o¸64Öq†kceå¾N­ø´ûìÙîÑ*™úiõ=tÐ¹&VSG‹o×8;ÒïAíÐ×SB»Ð6 7k–õ"Æ‡Q c Z´ø])¹§\M)fMi‹çÕ'³[Ÿ¼¹hÈ%]ØæëÓ4‹­|o
A¶½Ëˆ¸Ä»*DìóGoiñ K‹iˆAG³k&#l<Ü›¤ý$â¢“.· #³ »:†D·æ0ðï!S)Ä;KŠk‚{ ,ŽÔ4ÕkQ+ZvýÀhAïuË°”nš…ÖªwÐR±4E¿7,iÇÌÓhoòõ&sb½ÉçáEŸHÙ´ÿÃA¯ÉŸêFAî0ßh°•ç÷;€Y=Åù(ÈÐ~a¤¶Ï$ˆÿøAHÐ>M´w/²<¿ßÁEþÐpû±Ÿ›°ùä~Ðú‘“ÌÃ÷™îð?ÝÁC¸‰÷X™àïà.hÈýˆ(ÏMÙbv¿x­A{–i!´ù!,„xJPƒö÷³qŸqß½¥§ý=šÍõPR:“æN¤sE/ä:ƒ{@D¹@7Ëêˆ¿,¦nk³-‰Ó’dv$Æ÷Û%ÚÀ"láÁ¬ÏÈõ­$4¹h=Lô9h§åw•Kw8.c_²&ÔâÄjÌ´kXaN‹<=Ö¥ªa±üVvÊ>Øm1Ô¥Z3T:ãŸa?rõÏ,ŒÊ¢(–ûH Œé)….ÑË-u‘«D™ëçó—„Û%  ëææB Î¯ 6Ûœ†xs˜áy„
ð±Ñ›rÎzË²ÑmÄWÛ„OkÀ·{´w°sðj»{xÜÝîxZò9ñr+¾Ü ÉÛ’ÏgV+¾k·á³äÕ±ßsYï™ïƒÉrïÙí½=²4˜N=÷îûî}JöÞ$Ož’îQÖÑ“ƒÿÕsÊØ
wO²Ñª(ü|Ðc0TxŽ÷ã1|Œ`.£èôê¿'ý( A’EC •|`î!ïOKç««CâPâ6†úg­ðhÓ·Zž5lÖçGO©¯útJŽ‚hô6xGšÝÃ=üŽJ‡¸-šWírÍo"^ŠZ\Š	O)¶xgY6M;kp™g-¸“­IÜJØŠÚxE]Âa0ee:Ý“8cÀÓž%#³.÷Ñh–ä5tPT Z˜Åõy2"Jú×6°Kã¦…H8(1ïö‹/,un±JíÎÒ¿EÙY“nìTce¾¦)mklh§çùV¼]Ý¾`ß]¾¶QË’%èÞ‘¹‰™~â®Ši W}…û™ûubÌJfNNEŸýýøÕöÁ³Ç{OVIƒä7Š‹w‚5uø^¹XFÝ:Ò†°oÈÏ?“a0JaAYŒÁA7ë`¤ JÇÑÃÞ§MÇ*áû.È8„ÃïÀæ£°UŠ·QÎ«q´‡Ô(Èf,
hpËÆ¯ÕŒ™ä'éµˆ”2M>ÌªCn_ÀmØ·48/=#ÙÖdøÅÈ{šÍªïø˜Uãþ,í ˆ\©ô!ýÏ²Q4	wOB›¶áMÈ%Õ”…OÜ*aŸ3oìd6(}–£äUq`V(ua 3fsá&á;š€@¢É}× ;šµ`­Á}Íf©}†~ÐÝ bPÕñTo‘íA§Ðž  !ni|ßpz:ë÷áNÌêùõ]Úå}ŒQXÐ{%	á?Ííƒ£ÞÚÁpˆK]éƒœŸšÆð3¢ÁjRÝC Â3‰ÏÄó? ÂÕ/ä§Yx‚/Pm‚¨"ƒ›
Äf?ÑÔbõáss]óF@¨Ää¦Iü6	¦>	à±˜äÌcv Û\0HœfYÕlQ•Q¸éCÉàaœ™pŽ/¶@‡N¤Miq{ ~€«•À$‰†šáyGÛ£à­ÆáZµgÓ¶Ä¯´µÔè¬øªµŠ"ûdh¸ÅÐ}æ–)i× ÌlŸÀ­Ànq	æ`NŽ€¦ -jî>;:x~¼{ô
~}õýî¿ÕnÌ²
“T¦°’·°t=`*²Júc+NZo6Zív[_Ý$—Äù’ c¾£øÓ °¬%¨0Ò Œ"vZSPá5m	uÓ5Ý—m§c¾öøk;w]^]‡ïŸQÏË%[Ù^oÙB<È^ƒ_±@sÅ
Fš*[€`áŠoY(’q.µ˜mi—í¾qµøo'É5/Pû½ØÞ®[o‡µú·ÐÏÏäêW ƒ“øM \s8®©›@cø.F â¹ó Ñïuµ6“Æ½úŸ$ŠS’h’ÎÆqº šo®~MÕ3ÿTÉá>0ž#¾M@ óõÖ%|Æ¸q¹-¤¤1ˆè¨v²$l!ë&Öq™4}ä$„ïW'_àGN;
Ðå$!¶èV®Ù¥}FïŸˆ—&ž7¨VxÜ?vå‡=Dv:‹Mnƒä«yŽAÖ%›íuòx¶b÷0(ñçû;»O{Ìdml€$Ç"£‰šÀHôƒÅlLµfS0Û\u/Äžþ¶"	²ñY@~‚+%¶p Hƒ³“žR¤p®~Ei˜ŠÇ.T”~ª4NRäÂ?Çp§I0&M®|Y¶Œg~2¾EUåoªf'o¢ io¢4®'f|‘°f|ÂXº 8ÇKUéE¥iŸì§"þ©kó{e¹ùÜp£¼žI®›”ó¹‰I}¦æ@XS”• èOG!Hop*ád0æ §…Ø²\ ©áÑLPäÒv'›êöÐiú ÕOVËÙ'g¹æ¢BøÎ[°-`ˆ> z8¿³àUÒÿÎ¦a÷å¯66ïÌGì>
m§n²Ÿ±û¤4\i¸K7‚ä)3üLø> áëà¥cLÂø·Ÿ0×UtõËÀõ@ô‚ñI„R{Æ£ôMÛci“X¼ß¢¥3ãûa–fÑð]ë„™Ñ«é’µQygªý¨–à}mqXx‚/žÿEy‰@ûÕÿÁöohw‹&‹Í} Qk'|õC½EšÎ®Ô²L›µ¢ÛÜ¾¶±½ÎÄÄeþáa _/‡_ã8“~˜dá¿£z_P©5ÿà£ùñÁN·‡ª…T@?‰¦Y ;DŸÙ-Ò¤|nâLƒÓ õ¨¶§šýÐxÁM	IÃE­@õü¹….L"£Ž¸ð¿ÉÕ?ËÇéŠ#ç›Í³ÀåWà“!Ð`#Cw¦±ê}xfgñ C‡½cg;w~eïü{.fñÈþÕ/§QŸÚ©á¬·F@Ù¢7prmrPe$í0¦¤Ï
¹y
Ó4	ûômp‡@ö²ÙÁâ\5ÌöŸ™…ióÂµOs$8ôÞØœ@óëQÓb‹g:ÃE¼GŒ¿YüJT“žèu)˜ˆT«í ‰l€]¨ê‘?âH(GþÍCŠT¶s¤bÂN×_dÎÐ8L£w¿Ý›üö³}¸7£&r ÿ6‰Ô‚ª	QGqùŽ¢ÃýŸ{Ï®}})"—ÌfcË¦xÂŽ+ÙØ‰Nq+cA3™bn‰1¹wocãO¢tæµã"ßlìÂ"lî€.ŸÍÆ'À=Àu´“
ÇÍµÿ½³vºŠ"¡ý^ž€uâÃ”ù”<ë®bñ',y®0
!Y¹¾'è¶”Åb†¿  §=Tpl !èÍÖQ·€%g‚ûeò=Î¤1†ŽjÕ@&ÝÓY |3:MgÀ3È°‘¿BöÉïœäÚy'}šíÒ{Ò?û?’`ˆ‚Æö}`ö½e5ê@flÖK”í›6zY–Ô¢à¦½}ñíµ'ŸúäÃ®0GEk÷¸ŸÂ•Çør•ÜY__ÿ¬WaQÎì¬Ñ|Z²©n·ÛŸ(3vš€Ì½(÷À:Y”}ø#KßðŒQÛ¹Ø³1µA©ö'Èwµß¥¨qàŽ¤Äö(œœfgT!»N¶L1n4þ¤ãrX$@{€üŒ*€gáä˜Ál ôÚÏðý@‘@HéûÐ8Â¤ƒ†a¨÷„P*:÷£q Sc
¬Ži±ö`•†Œ¦e‹ý'Y…Ú#@{#Ô4_ä¤l•D“aüÒ“„ ¨›}@?¤ía4Ø‘Esf²ÇÊ¬]PW«$ÙBj³¢aìþœ‡weíhnN•³ÏpâÊ¦±Ö©mI©qŒµ¶ì%ˆ+úéV»8ä—ìõãÝ]Å™}‚±cK©õO>²‹ŒÉˆÕ ‹–ÔÕ‘¶êµ­I÷!äi5ŠBµ¤mµÆ9±ô)&íÑ^Óƒém’¨;õê0‡÷ÉÐI˜Hƒs¶•=´¨ò+aGš'ÜYDÌ“vÉ'v^µ…éTÿà]™åNÓs×%¼i"‡š´ûPÍMTè2’Ùø½çz±q·Ì{íÊ<™EivgYìÇùÖÛ3©ÔôkLP-îPÕbÒèÀIÈìm"³·‘³o[w”³íóûNóˆFG€³ôŸ¹òÖ!ý›ªÏü™òA¬ÔŒ¾P-ùÓ?î½·9¹®øÊŠEa‚"{÷ÌE(—Aý&ÐèÄ-•‹ÁŠ v?Ìê n;ìŠj©wX^¼LKû¶D ^¿ôÃÄjá°TwÖNŽZ¼Ôyß’î!ÚK}«Zs6Å¢/Š'Ñ÷­æä˜cñ¹å¡WéIŒ×b‹A/^xZp“Ep_… Þ¾@qh«Íµ]4°s˜o å.„ÿÜ»:ã@Ðªm§A„˜_4Žôµ)ZóÏ¼¹xèÉæO6Ë×>çÉ‹È“Üjgñsäž¶{j®€˜à	pÝà¡ïv·÷v¬!úDñg|òÙŠ2W^[Q<Bë¥Þ+¾\^^ôÙÓ<ªVùÛ×ÝË8pËCg&Š/›æ“§W”ºTuá'z¹ÔÕÓËe½\Üú{¹xëLåòZ<ùÉç-ÞúÛ5Î½&5š[/—úÜ–ª¥¯MØ— Õ—‹IMÏÕgõº«¥¸W®’·_?”§:_.ó©°°àèŠ#«ZÑPÊÅSË/—:{äYÓ¯ž'V	
¤»ˆ*SsIhäô^Ê­¤Dõ“ä²(Á¨]åx–p#FŸ^<3õr…çÍœ9Ã‹š<¥ùYB.>.çÅ”ø ÆaJŒ«¬˜JV´–<*ãš=gu5ºK…žK3ø¶µ±IÎð?fó¾¤G¥œ¯YÉÉ?“hŒë˜ÎFiè«¸;
‡@éÏ¶ßrÙ¼ë"¾ŠÆê£‘)Y·+ðvQØK:}ý‰&?„ì†ùawéM„Á”³p‚ŸôBååXbi[cÐr2}À”¿>ï@˜	Æ<XðžÁI4Š2×êy’À«»|ýäp·÷=yM~(¡n®ü¡M¹¡²dŸñ‚†"°©XD¹¹0€óa€]ø)þ§žwedOÍo¡.ôôY·ŽÂéèiîu}ä…RÛ’ýzrà‚’ `Nû³$ÉvY¤Í:\jy|¿[ÙjG)î5ÝjÞÙÖÖü½stAc„ú3M¹"V¬ÜR6Ãúv¨–µ5r0Í¢q—£Ïõ«ÞKŒysš„oXÔ:œn»ÝÆ†uLéµ´SS´á£I=Ôc‡E¨×|÷ë5¯ÀÚœ=Õ/WVj Ä’”Y|z:
?„¡¦pYWÔl¸  ,jêÃß‡¶d¯Ë¸U[2Â‚N_ü§m‘ÆÁ3ªT=xüØÐT-‰<N×žÝÆÜ`éœÒ‡£¤b`Ž‚þõ4WK sQ‡…èÃâ¢@·–‰æí¬&r¸ÝoEùˆÆ¡CäQ1 EYÔ³Ö=ò¶õ'•UU_‰*q5Ï÷·/š, ‘±cìÚ
&âQ(W)x¾ù*tê<§•7“ï!‹ pÖºƒ–l«ÐßCÝÁ<“Š´•EÕåmí“îÕyëÝ<ù›ïÝó~PõWmÑÚþ†MÞF;^{çÙ×‚R;o`bÚ»º äžçâ-¶.ñµÛñ¼îìÃÜ„qcÇÒækòËÒlËÕ${#S)I[33ÐxÐ¡'ñ[üÛªªEhá„NXöó¸><âªÿ…_¾nTƒþœ)ž°<	“pÒB2¾úç(‹¦£8åqñR5òSÒ¸ÄÐóiœD:¤-Ê‚no+§WJøL¦'­M<=øwkÞÓ~F'A‚à¥×q_¼Ð.ö‚D(PžEióïÇ “ÒP>ðÝwð†º¥_Â'òíw´Þ;åÝÄt£y×w>•º}Ä?ów‚, ¡ékt	‹–:|Ò=¤É•¯þ{¬Ìõ	ˆS³ ïïv0­Óý0
ÓC¼Þ‘4Êcør-ÿ–ñ8…Ç Lµ÷ã€Xü&èqigÈÓ«ß~šEƒ ¤üaå%‹n	[jKÉlÕFR«wèÀ’˜™™¥²^Ld¯êåÌhøßxn»ãà¤ÉFdÚâ¢oÉá«QÓ²Ë!ÂÅlÊö÷5´’ÂP“)"Gòy2(¯‘*ëÙ:-äRå•e_,>¨Èƒac2ƒ-¹på:uã³^%­¶©Ãá½zèŠ|P)SC«Qè¿´±96ÖTŸµqN.À/ùž“ê„è2^ÕTÑaZMµ1£aq©t?ëQó×ˆ$ªÕ5XÃ8yŽ¹-d¬á7ç2ò_Â¤)m°L¹L-Ìs7¦&³ÑÃ
Ig'Oq
i‡è)qƒ3<ìÒØ²äê·þlóÈÁ÷¦ú{“TÈ_€»ãèTß·FSgÃejiÙw…~Î¿éeŠëµßÅÖZ¶qÁm¨uÇµ‘ýÜ×§ÂøÞŸ—•*ía„JX`¢iL)6þÔ	X…¬UecÎxˆÖº}bs\•›’lºT-o‘wŸ·ÉgÉÂwÛØ#±l¶z™½CèÄÚ9L mÜ‹—Ìµ´IO¸­Á¹ã¡•Í:Aþgo€¹&o_ð/_Ý¾€l¹&ùòÓSh¨Y˜Ý÷´ºâ|sE*ä>’Öb Ö0Û-Ø/ñ=†ŽgU¸Å1þÍ<;mM&ñ8ì°[d«¥Œ±Ø!· {sÕKc€Bý›  Ž=°Ÿ¥«Ž§ÈJ¥/8\¼Ô¡ö³O!b¬@;níìŠKd,n¨<©Ž&ÅºÐ/ÔêAKy$&%pØw'¥®´_ŠŠâ$÷C[£¦l÷r³2Ï$g÷G÷±þÀîXÕšÃ¶u4lS\€õéí]wš´¢ˆ]‘ÇOÈÂˆÓÙº´y&•TãìnEßrRrcUâ2Jž®5ý[Ùš˜ÔàPRžÝuÌZk‚WS}}‘ß\+â’+‹šé
Dé4ž\ýö&5#GØ‡=™ã\¦¤žPÍî!áÀ,õÆ; [røˆÂ3gÈÊnY”"›óÅtœBU¿ 0q½ÇUïÕÁwJžžîxJv8°oKéá¨rõFûžO<Ü’Š€‡
‡ƒh6®,·Ðz«º›’Âq!=rÁÀ	ÙÓ¨ˆR\L–ñ¥`ñ\f
|N(ÒÃ8«òKP%Þ‚X»ºp?ì`î™˜e¾}Í³%Qž”ü‘l¸ŸýöTŸ¢¦èC¤ª±T’÷Ài¼êg89§ÁdÍ˜\qä­<*W0x{Z_3b/I¾2þÎÕ²Kõß¾ë÷º‡‡¯öw{½î“Ýž_'ž~Ñ¾–"LàèÑw\°óÿra²Àk3*‡ö)fïÛc?	çÌ2/Š—1‰—	É²È–Z6BYc ïy†ž¦1y€|zn<ÛTÀBZÙ³¹ICÕ§[ñaE&7ß4ê|.2_'^¤û×Äd-^¦e*o"ÁŒKÏ€zÿØr†à~Wù†H·jÁˆD‡£Y*\hÖ.4>F(
då‘í²ÅGjíÎºE®,žåuoE¹Å}{äœò\Ü"µ¾žÇ·É¾æ~¼BY³Òpã ­@xW1G<£,a 4`àêž£9UM©¢o¨žzAzJÒ@CÎµ¹D"V€J&#à<oœ”OÃšùQÕBãùEe¡qêXÞÏ§f œz!p–À©±"ÃŽ?‡U+þÍç¯Ä™y›ÉzÈÎGÆËNÖÓV@.å¬2›>Rt5Ï–u€\n9?ªeK2 Íu†…QS–¤q’y†@ü64dÔªƒáÀ×u³çx­Ü+p	lçê¿Qú§X2@ßÏ°Ð¸—è4õ²‡¬%´Ò^½=þðö«õ¥ó¥±{ÞÍ¢DÞÎ­†ŸÅ«¿gÃÃÐwøKÅ{{x‰‹u=H'øÌþ H ÍÏeÕÂÎœva}	Óþü„­}ùûm•3Š'á8ÆàßmßÐ+×@¬ã¬uKw,ÀFœ†RÐŒ®Á¾QlØ¼µ®¢pAO\~	‹—õ¹QÃCÖ£<üÌÑíu„1rßDYœ8Æ¥
^õš»vœžîê6Š&pŒÁ[Ô¦T-æjMš¯qIú]æ´ò?o·28ž<§ÌsN”&ñ›gg6î(-h(µQÞ¡óUù‹¾¬?Ú”ÅlnXÚÎ|I×Âßyt¾õèØR¨˜7%ªfjµÇ€?ù/k„s(òoÉD—òðo¿¸"ØÛ¯Ð‹ä\’p”†žg0§ì[‹ãzïp:/¯uð=7XÎ›µV.K«:JsÎ,å©kß¯ªžå¬}¯
z+÷â]ÁíP6ÃÔO¬—5§B?ìÒìj¬’î£UÒýŠ‹k¡dp™"¹•Éœßêý4aè}gÓã)Àmt¤M»"Û3HËœÕ˜`•q8™±ü)"DÔ8¼ú;uô"76yƒ1ÈÈ4¹úõMdHÂR,z #—ligö®¬Þ°ËÒ 6/Jm"ÐÕ3›2³-Á1N²¸Òêú|0y,²»ÐvÌ\ë©å“Yöç\Ðò¸Øž@¼´ñe§Ç¹Ü÷@zˆŽM#i|’Ðôh³€„cøSòI8ŒðUø8ørÍJÉ ÿ{Ývi’S6ÊÀ¢Ù_ÓK—Î ÑúØ¤¨fKiFÆ'Ú¸]ÛPŸîÃ  »lùåõOGßb3ÀmU„ ¨1ùn?Í"];£sfe#^ý#NW™"þÆ£«ß2šc7?­ÊóCdH'ãhÒ:k½¸ÃùUæ¤&d(rúdI Âîä”ï¶éÙË;Cþ´$:,¿.ÁˆÑ¸©å¨*µDdœ¨MZ5Ë§ŒOÃÌÂ‰#çµ)7Ïu`äÍüTii—Ì\ä<Lª‘)½ ê	'ÁüXLÕ¬\ÎVq¡›õæ;q‹^«(‘È”wzfk"ýì«”·üU…yrš2ßåQ\5˜opL‡üeà´¥Ø¬;$ñÉ!¸9ÄÜˆŒN¬Â¡ŽÎâWˆ^U•þJt¢B4|ŽNƒg[ë•aë>ñg¬mœ¢™o6›ð×Q$IÇ¼_DèY‡©Îá†»‘J0ñ6FRa>©	­ÞßÖ$ºTÆ(‘×óQãááÎcc·e Öûèª^Fß—1º…ùTF­ºø8pVÉžþ¤-GNj6ÅÁ ý-Aˆ õpy`Œ‘‚>'ç7©zq€…ƒƒëˆ‹á“ðüê—Öˆ"¡¾Ä¡~Àu±t|)žÆ>ƒç9ð÷+ïá~¹Ãyp)?ˆnð/x¢UëHÈb€raó4üÑ8‡ã£‘6yTÐï‡ÓìA£=u¿K›Äã´h*v4ðÏ`Â2žÓc×‰_ƒ(NFáàAyåÕÊUœ®ã75B—òEE3]£øô³´eq<~"!OaK¨ëÜ`ø;š”ö~)rWYÐõi8ÆT¯ã é³Ô* Œ°ö2	 ¯Q€“á¤àR—YLóe§d–Î®~I¢8]\óIìúµ#‘ÙtQïS¶l¡Žç|“cfÔ*¸:²¸lÕ_ô™Bår@ó`Ùƒ =Ê…½š!%t]2ô½­gl£”-.ÙÀ›¢I‘äy‡ù>•¦§áNÞ®KDêî®o„ýƒd<ÙI‚Óh/|ÀL|HÍ`³v˜Ð©—ö°Üó)`aÌ/A_æôì·i2ñÔËtrÉÓ2.Ò4*Ýb†—cr0LÐ!ž&È]7˜W Œ‹uÌŠw&£2R† Âà†Ó*ÖVßµn›{§„ÒOÄàr=mÜwG!þùèÝÞ I©T§Ñ¢4¿±²ÕÆ×?zž6<˜Ù¸×ÎA [üƒÊœôÂ?g6(r_ûR_É‹ã¨m»[‚®qÁŠ´, /¦'-Ìh_¬ùB8Ö2Èah¸?¯†å¢„P-&üâ2!‹íé[NÔœ©I´/F•¿¯ø•« OŒ&úåí2Ó‰ûçlÅ¬èN7I‚”ÊFi<‚ã\ÌBÍªÖ‚Þäî£31¬85@BÈ¦QÁn¼	1eÑ$&i8Ã7"À
À ëì‹o<^T„‡?rÅ“Ø0¿.ÛTÄàƒÕd “I;äðÙ“UòçÃ'8ÿÞ_Ÿæ~pÞ&÷öé,?ýêí¯.K)ŒÂ·z£Ù
×©×OÙ­ikÙÏZ-fýó"6vXÞÃ$FLT§V„!Šö#!BLCàÛ&\Q‰ï€ ¯øŠ“«ß0Ñ€J´e±›6YäÚtf{êík¢þÕòdî2°òQc³² 2›—‰ •p]nQ‡½„—‚š«ÅlÍÓØ›“Ùhd­Ÿ+·sySX¡"î éQ¢ÝÄ¢¦i…W€þ\Ý- YÀ²·²›°ûÖ•ásãhd‡I”³¨Í–ŽgKtm²¬fSKæ¼Åw÷u&¾ªF…~Kýå÷¼ÍŠýoINÔ‘P=­8â‡½OU¹2¤òÀ‹Ål°«%zƒ«Æ öáÅ¿¬olÜ½³þÒ!@WÂi4QT],?UÊ*Î<Š~YWÅ¬Í˜S‹ë5 8Fd†®31FM¢Ó³ŒE½“§FO©kåÓÍÙ”9×»‚Ëzñ§Üg©½+X|ù™AC9{{ûÏŸv¯þóê?ÈÎ.yzðdï™žv{®1šÀ½kQ-Á‹n^0jâ¿ ýÍÙï_¯6¦ç«4Q5ÉðóÊKò3dœÒù¶oílÜŸž¿Âÿ¼$ ÷£ìÆGâÐÛ¢tÊN ¹[ÇÇïXld›)”)Š‘Â§˜ØúÊ€ *PH4‡¹ŠÆ§fôÈãšmÑ‚ˆbvˆæÌæúòô‚ó}¡bZWŠ§$5>¾ë|ãƒB€êÄ@—ktÇí=újÇØ¢ðLohp§¯yø%	‡@Âä0îìÝƒÆ$n‰¯LBŒÙDÉ("YE°­,ne	&ñ¸•þøŽ©lãªèè#9V#öŽNmâ$õ— ¶®òó©[Ž«ß®§OPaÊsY0¬’!ò2†EO3Œ%z@]Æ¦‰N¦«]û:m'…ãÖäÎQNùè[Âz))Zl0f{¥š’ŒS((eS‡›rÔµ©—«€H¯âÐ²\'úõúº++& õU‰,u ßMÊ‡ýS1ìWê°ÖÝ»‡§W3ýÊÃWX°_krÈhÂÆ®.þ`·w¼K¾{þìx÷ˆü¿_Éæúæ}7õ\æ]4ù1­óB‡Ï6õ^é}Øô|•£x°Ìv)xQƒój?ÏaFÖ”\ýŸ,ŒÒ*F¹@ÃZdù¡–N­`Òxh¡·Ù-:7‹qPïYø–ÖŽò ˆ’Y2bŒâ~@ÿÔº%jýâ]U`(Â¹k¤ŒÊÛ{i½c0éÿ›X2]m«¯œ*¨Ó,Hñ¤7;G™x ï¸W—HÊ9ùt\ýdAûŽ*F :½bb‰¡@ÊOÂŸf‡Ò×<Ç„#u*K9¡zìív›7^å  ÚV¬¥­±û®ÍNSÝŽµº;úüèé|
waþý¤éSÜÎ§ˆHó`J[F+å½-m"Å#zYÿÍäÈè#ÜN-"gÏs)E0åí3„ÔRšªK²š2Ñ`Z9žu
{©ó*£^…€m‘o{Á›Pq›±¦7«¬I©"ço×ûÚÈwÁÊ2ùðÅ]¦ÉÕ\Šž&=üåExÀtmM­™u i= k¥#ŸÂ±ÉöŸ€¤qž¡ÎÔM­¾Xzž^›ûÎbœbz#iÉÂø0V0F!1f\/q²Ÿ”{Ë,„ƒ›µ™&í›£Åpšœ>BpàÇÜbW¸Ð#)W˜=ÿË’ýM©Ä5š5ÿ¶‘ §x#¤š9ÂÉèßßÂ­¥63”9jV UólÊ¢›;Ð„Amû¢ä¦¨¥R^â»_9îêéÞ³ï{°ý˜FÂx†ÅËbogìEŽÇdYÇ£„ì˜ã€|-­Ø¤†òðp6aVÁQœöö¯U0VÅàz,‚³Ù2sLÒÛ/‹²6ÛÑÆCê©sõO|9†I¢AY[QcÍ¹måMëÆÑä±Tá¯QøVš$‡÷üù g´€x>‰Á L\¨Ÿ‹škî)ö“0íGÔt¿;™£ò4·K¿ÓiKSIûgáç [ý.ÒûâRY
rZ¸ü·«TTëíP<5Ù_¦ºC–l6Ònýãâ'Üt2KÑÒqšÄhCuë²ºõ>ãá+°=^W£@'êå0¼?:TE^w„vV\K<¢S'!é…§3@iWÿ¬á¾=ciûl²“{Ê¾î<cŽ¨]ÎÍ¨©½Ã©÷Áa•ŽÒ|’Ôl;:MX®­¶2ç©·ë@*j4--Ì© Sµ°=t‹vj¼¿|Ã4'¿)XEzˆ•7h•6mhd'g.*”œ»¤¦owâ·ÙêMaŠ¥„3N:¦TOÐøþê7+ß!I–ôÞDé¿+•5ªÆæm8B¿ÍnÅB³ïNxg¥p>€ÖÖH«Õ¢NÔI<
ñæ08eÎX)¡ø~ÿ}ÂÒòšE-¬„ÖÐ)åâ€õE³5ÀŽÝ~¦²-}Ä/¨+¢Løé9üsÈð&ûžcÌiŽ;$ÍP±J²­i¤³>ÚÖ4ÈÏy(äßÄÑ rì»¥ô	uQL×žOã$#HR)FÕ/£©Yy@^ÐstÒù'>ÕÕbðÕ?\vlÛÃ}˜‰ï‹4’þÙq˜ŒWQ¾èå_Â0R/Ž¢Ù`ü)oÃ}ÇÓ\j´Ù®üÜòÛr”¸YOŽ6îÑ?Ž¿j<d´…þ¸únÍ#fíf©ÔyOúBé»èû=„ÍBØgƒ„ýYþÑˆ!6iyÄ$<…£F<’¾P÷7ÍÛž†ÙNGà4aê	L7)`æ$}–DYÔGí«*üÄÉxª/è}W`'‘·ú#i¯¯wèÿ„ Áê‡r}Ö±¥ö ÑMíÙY;8I›!ÚÑã— µHZ|ª´Ü	Þ¥¢e?ŒFÍ¼»5Ò„¾N¾$÷óÿlÞ]YÁ´
¬¾Öðå³¦èl…l‘ÒÉ;Çª—ÒžOƒwÑGñÛ”Í>HQMY~cõ "g£ŒÚ¶½xIïŸºz¯ÚÀbíý³f3 WR‘é1±ïÑ+£¥Ü¡qœÅìÐòŸØwaš{Æ,ûü[+Ó˜>RF;v§á ¬hL¼ÐR˜¹yc‚¡:ÓpÒXFC´ÙÌ{(-„V;ÌûÚj¿ÀQ_3Îò=§UÃäàä¶ç‰H47£Sµg<‘ý7¥æ4"ZìÑšÏ×ß¬Tf2³ ¥$¯KY­Cã’ð˜d©ž#¯ÇH²áYòl‡Ü-¾“"–Oj:<ÍÈ©‡›6ÿU \ÍTàü"øYÆXß(Ç$‚y÷”š_@h¤_”Öå-Ê1RžšùneÁÞó4¬é«ÇÑ¸rho‚QœäÐžï·ú5¬c½Üîc7›Á¤qüÄÚôklpWÝUÚ!LÓoÜ!u Ø¤»+r/÷¥ê4ò¶Í}¥ÍŸ¼Ú|­´äæÓhc]muo½rb}4š<Ž³`„ûÅïË¼Uµ~<Àeã%¡”GÜ›˜}=@°Û(yJñ†=”/ßL½z¨µùQÅÀ:“
HoŽ02Ã$	çªLDîv«MÙ8
ÁwþJšÛñ8Lú>Zá¬”
hW‘µaù·§³ôLI\Ã–¾7èèQÃû	ÌõÓ‘â5\Õ!£"nÑÏ?kðƒè
™ûŽŒ·Ú6€6z†&N}<ûAÛþð,ž”;˜âw´‡^8&ýp„’½¾ut^jÛ?ƒÓ€¯óæø·cÏ'QVêãµDeŸ“-¹N^‡LºYs}¥ÅÏQù¿¤È*ü±$Ih§ q†ÍÕ7¡ƒ®<Ñ¤Mƒ‘2c~4ˆZ¾iòÏ)šb(Û,T­¹TðÍOM©£½Jh hœe…*-q‹[´ˆ£I£–ý0:šFØdÅc:{å§ó-/’*w°ÂÖ^™µX©L¡ Þ®[­¯Y¹Þ0c¹&=Ãàžq(~pW¿ÀÉ)ðþ(Ñå[ZÖZ¼©RÁRE „a”«Œ,W ïJÍƒ®a1¬ô¥j
ÐªÜC§ ö;’ãEñ—ø—3ÒGR†y•¼(1º«L^2ž
XJsh ±"Íh”¢äñ‘Xâwê£'hF6`ŠZ&%Ùyp¦¸(Ò³oóóPxÞJŒ|Áý&þ;iËHkûlƒAù{í† ·ß»6šè6¡ÙŠØiäb;¨–ƒ-ÿÊWËhb8°|4eA<Óx¿UVÄ³x£ï¶%^ê¶e"ÈõiüV ×6`ÇÑl¦ÍBhWk(·;wcáŽ$”Z»¯œ{‘–ÎÒÚ¥ÃT²ßb¾U:Ažñg¥Á7šÝ¥÷GAV,˜(ÝBy\j§‡]ù%íiri¹Q¾a¬d`ùâuÅg¾Zñ‘MJ¾%°G=G¡ù)ktŒ<µ—²ò#Jx.)¾z¡ý*pÿÑÁÓÝ^ð÷«íƒýÝ£í½îSZ¦Ê½îÓî«ýîñÑÞöó§]q9÷¯~äÑÇPàý>{ ÍÏ'2ñ7ÓåD#CôL]“ çÚóï…~H÷—¾ôÍ¨Ž‰ÿD“Q…„ÒJ(@šÌPp(HÈ72â3 ¤¤Ð’öJk1v}YÈ*Ærá²2@¾!®þ+ã»en—C«ŒS«›¶ú¥[¾^å[1ÉRU6ƒÕüÏ&À+gÆ8<VãR\¹†rX°ÃCv«¾!Üž€öQ'\#(@¿IäØYtü”™TqZÀ•ì„Ë¦ÇG)­¬ñ¢Œb@m…OÜ+”7’Ê¨¬ÉRk’(wèA¨*æRíƒÇ‹Ìû( ¦Ýn+lÐiÿ^
FIüœ÷_d¯3ÙØî>Ý}¶Ó=Ú;xÕÝ>Øí­»°*MÅ¸¢Ž:ÿ|ùææ|ôÚ7­C©zí®Ï‚«QÓ•;Ï3’ò>ce¸jôs¼p²b•=îÒŠÝó>ð›V£¡dbÖ*+W‰šú%åD³»BâN1/Îµ!R1Šn`5™!*ê‡£pHÅè¢•«‹*‡{W~Î‹ßv"4«@!2iWÌÆwTS›×Œ0Ìäõí‹„é°.	­ 91÷$‰z”¸%©WÊnÿW\îWec¥úÇ<ŽX·¨)ø9©ÖÖÚþäZM'Þ°ŠÑ3:¯£7¿¢tBPIÕ<º½B+UÅ£[LW•×Ó‰@¼¿ƒäêájŸL’*÷#¿f´véê7¬ÌV¾“Ó³ÊÚû{{–¡Bç‡4ž¼ÊâWô§¦¯+ßTÚŸÄñjsüæˆ/E° òÁt
¸•÷Ÿ¯Gô¶ª|C+É{©}íÌ¥[uÌ·I”…À……¶Q^‹._ý½º}‘¿¡¾f¯wÐ£4>¥ÓQ„1r+/Ö_^¶ÏGéùëÊð9ž9Bç«ÿ:BØ;&,ºXã]'Ì½šgÃ]œO`»âŽw±ö%ùØVàÍ¿\†vËŠñ °¬€¿U¥A§lk\6¾(û{¶Y1=ÙTM]J{øA=ÉÃ­mz¾HæcÊÂu3$³&på †Y˜ôCô¢0|J	NAïQ•Àb”+óQ¬ÐTcŠJÉ>E"N²‘Š—-xÙ2í2Âq˜£GA|ù•Ñ^ã³­•6Œ¹´Å§áhØ¢¯·6ô5ƒÊ¢SÉ^¥b^ÉŒ4vÅe‘î&Ýž²¹FÙTCuŒÄ[QˆXkäûÃ=²$ƒÔrI*ñÓ±ßqt*}¼«½jÈÈ{…Y‘É(_oƒ™û%4*¤3d÷H!]rð¤A$K—É9ðîqJvGá)Sž)`^6—±Ø¦ŸÀ˜èÜb4ÛÑ¯åŽÁ+;wÛ­ÿ\pQ¹M…Ÿ²OˆÆÀÇÏ ñ«áªk^¶R]ro ãã^]@`¼Øàòéó†@[Ö}4Ý. B‘›áÓãè<47WÚIH]ÐšJ„W•˜,ëq4?çš–5$¼ïñ´ÿ?   ÿÿ äì„xœìZÝn7¾ÏS0ƒ Zd)ñ&ue^g³Û‹m'-°
„š¡,63Ã	É±¥
z˜b/Š½èå>_lÉùÿ—dïU‰¢ñŒæž?~ç’¬¤ýbå¡¤Ï™ç"©ÞŸpì¹ö_Ž­ó'¨0®Ÿ¡O$§ŽI&±wM°GÅ.ƒÇ·tEÜÁt8â$ô°CÖÈ:BÖ‘5ÜØÌÆ"ÄAžõlìÒÛìEü˜=Ã#r<,ÄØ'gVhŸ ù}·¤’ 9ã.áñ?¶ð°$öäøq.qµšb	ÞÙÂG0³Cìµ=-¨6SågÐ†X	cŸÌ4†û	pÂp‚ÀvÎg€8TËá1çsFW²àe$$C.Aoèýoœb>(+LQ1NIû…GV4÷…=<X­UI­&Åš<oÔ{ÕÇïZ‘Gðy] lÆ_¡·Ô“œ	ôÕ8á[6ËN¡"¨/sz×[>®	°Yìù¦À(˜4ÑáŠˆ/8gˆ¢¥ÊzÞpê"õ?Ûaž°'ÈwO³Ç—è‡Eª8™é-)ÇÊ{‚¹³,y>G‡ç‚yXØ#¢Ñ	’,´'ã)²A×ÀXb­_äLóÂ©ÄSÐ_ÉÙfòº"ã’84#Y¢”ë0vE™§Ž¾%¸…ð3ë¯‘p0G!ãèŠ3ŸIÆÐÅýï÷ÿfˆEè=ó¨C%$FeN·Ø‹@<¡­ñp[ú€—KÜÀ72DgçHù>ýz@Fó"GšÑ°L3è½ˆ<Fâ…œBféä–¤šÚpžæÂ¢9M'§éB0\…à4ÿ`x¼ ÚmTRØ‰çdI•YôƒÕì§rª6ÑVNBÄ#NÙ•±™‰r"°Z„÷±ôßò%c#H¨8X÷±ùê0k'{ƒ >­KÒ ë:Æ‚æäD\0n‡ŒB4ó¢§ÊY|ÆBÅ06½|­óÌUF(¶«@ƒ¿_ONÐ]_½ÎÆ†¤ƒ“¢°Îc§H=ö¤„Ir„ðTOIQÇÐƒ†žXF¢ä½Ï}ÿgà=@àÁÿ³ê½ß€©ÉzÏÓÀìÖ¹áqŠ’7}£€°‹3É›Þü-pˆ_ý#™“˜¨$_ç¾ßc¿ú3„K¹SÙóþ¿Dñ…Œ°FM!´‰ú%"Æbäãp00Ôî”>Ï¦þLÖ‰«·EÏoËsèyâßªüV1{S¨Ö>VQû<¬YXÚá3øÎn	_xð÷’º°ú‘æa0}À¸½ÐÝ,tÐ÷šÝ‰‘G‚¹Dgggè}W0}µæœÄ°ÖQ	†§EÃ‹^%WŠ…ïÆ”•Bë,#)×	(”éÂÄtèK©eAUnwTW=å€¥hÊ¯ì;Ûw‘¿²q$YE°KÂ%ÌëV‰©ìÆø‚Ÿ© 6ìía•Üµ:`-è•ÁJ¯<hýb\N@®™"‡ù!æÄ¡Fm¢P¨—ê—`AÁ¥në2&"‘5ì²N)Š‡è´Õ÷i€ÕÚf&ñÜ#5y2_¨Ë<ª‚P{j§Äö#Ïä’`·´$’ü:~•¤X{^]šåÇÉq¸ú¹¾D}ÙÒµ¨¢DáuI$,¬ŠÒ¡®œc7Ëån„I…¶;åÐ•ìN	[â-$%ºÚ•4Ÿ¬ó¤‹s—0¾;ëürl ÷¿ÝÞ¾ùi¸Ø¸ÔVµ+áE¼ÅíJ÷Ë}ÍÅéÍRZç?aÊú7tçHËq)À–ÚÕR×ü4Ã´ðÊm»tòÃU{Çï‡wHL!ŸÎ?ú‚|Ð HŒ»Ó¹·6ð¶’Ô—ºª	mÎÜu^^H­qì5Šÿ(e°ÔúÆ->qiä×(PÜÇ†bwGˆº+¢65ÚÀæ!$rÑ
È‚H$ƒá·ß­ƒÂ§î ñ[Hçÿ‚ç¶ï} ˜U¡Ñ¦û§z7¢¯Ñd8
±&çr0=BÖ±ÕÂÊÅë##ì.<$`Ó5° FŸžm”2[ûÙF©þ€·Ÿš©©x{ž»ñYmì€­G.,å·ÔG³l®çÏ¿6ÅvlseÜ·Oj9‘j ±°?Õ¾G&ƒÎjVì( û½«TVÏ­þ#Õ¯ Z¶Ÿª 9¶D´'¨Bk©àÁi~£>)” °ÿ3.¬£F™±Ÿ?GjÇçLh>Ó“FªJ™fFýJ7`<Þ sø»<fÒ­O-MÇb²ímN5ÔÌª¸„¹')æ×*rÀL‹¡ÕPQ÷MêÞU+	‚Ò@¹%ŒxèårUüüªÚà.SMý9@¶”Ü<*ê6‰­F³/Tñ—éÙÌ¡®nÈÿ*Ýóº…©'€ˆJÛÞ!U×xÝ¸ó¦z%«Xñlœ¸Kø<î-<äÒj™Ô éÕfÚ[Ô­¶ËÚ(J»ÈTùºU"y8à…¸ü8™j ÑÈYRé©.D^Uºj’±¯±š9tš*=WÝÃ«y Õ'ãi¡]ŠÕ|{¬ÁÿÒZçZ”¥bµ]vN™;þÚcâx2‘1ÙvÎ—{ÍÖ¾á”ˆ¦Šè8ß—¬A–ÉŽâ»Å,Ô2W)B¨|´¬÷Cz©7u[+Kþ¤÷’—4df{®÷.òîÅ®ŠÉƒòwP´…ZóŒB7peOF'­aTÇÉž´’ÔÔØ#äß°.ÙH„•Ë¶ÔÕ€—ê‡Ñ/Œk\¾œQ™oÜ9¡©C¢ Ÿ¦?ùñå‰v§FŠ™;¿ÔèÍTÙVç*þ.Ü¦Æ)*uN;¥ë°n—'±¼ÿOæN¨°×—l
õI»#»Š 5Ê-éÔ-8 ¾²qy êÅÊßÏv'¶6ùJI~ÿsÏÈœ¿ âËií<º2»ð——”;Iî®œlË÷U*T­»€-žkuJg^L[d{ã
Ó%ËÒd8¤îˆ©¨½U³Ãä;ÜÛA·\ûíaµËîÂå›ƒÅÂ$¹iÓÃ
Ž’S‹ù8v(ô
Ê}½ûŽ;»]ùÜA±šë Ãã6ö_Àt±.ïóå‹–ÄvR<R¿>F¡~Yè‚4ÞÙ*;°Ä¦k'þ˜uåÀ“é	ýwrÙáguè­;…­Ì†¦1+dcK6?’Îã{]ŽÕ§mj
Y×ºÌH4—Á¥òAC3?ÒÚ¡¶ßôMñP­ÐxÊ_4¨ocæ] A¦á°Ó&†Ì¢[U™†=€Kªþ×Ýúk@àåî~ÝÍy|zÖnuÙ!¢6ï¼Ôi8=–Ð¦‘vr˜È$[
nömk¬Ù%»‰V¾
`:Ø¥*¹ÔÖž«¾v·ÄýV D²û.ðÖ
Q>ÝmAäTEb8©\«­:¾Í0ÝÊbwÄ=\ÚÖ#ˆü˜Í#)YôUCOˆ–öA6hJÏg›J ½bîiÁ¶½H¢êr…×*¿šÝ·[_3òG,Ý5[F“;†éK&d¿/û$™>¶w©P×:Ü³MfÑ>t¹ª7‰Ôn²>¾ÝTw^êWmžý>>E•¬ÝÓ?ßíNršƒÝ6šÍë²Uí™h6¶­õòÃÃ»ôHÚzJ(ÛHPV2«æ]ÇAY\²ê&C/E©yN-Ãu©ŒO“&&`:Ç‹îÿpñQî¦Téa.O¾:Ÿnó–MUè*ÙãþÅùØJè+]%|ˆ9¬Êoø¥önVMäW£ˆÕ•Œò¥d}‡­íŠ\zg3{Ÿþ	oŸü  ÿÿ øC