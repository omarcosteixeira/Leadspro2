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
  firebaseConfigRegional,
} from "./firebase";
import { PainelPrincipalRegionalView } from "./components/PainelPrincipalRegionalView";
import { AdminRegionalView } from "./components/AdminRegionalView";
import { RelatoriosRegionalView } from "./components/RelatoriosRegionalView";
import { CheckListRegionalView } from "./components/CheckListRegionalView";
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
  ADMINISTRADOR: "Administrador",
  GERENTE_REGIONAL_SM: "Gerente Regional SM",
};

const VIEW_PERMISSIONS: Record<string, UserRole[]> = {
  dashboard: [
    ROLES.ADMIN_MASTER,
    ROLES.ADMINISTRADOR,
    ROLES.GERENTE_REGIONAL_SM,
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
  checklist: [
    ROLES.ADMIN_MASTER,
    ROLES.ADMINISTRADOR,
    ROLES.GERENTE_REGIONAL_SM,
    ROLES.GESTOR_UNIDADE,
    ROLES.SALA_MATRICULA,
  ],
  formularios: [
    ROLES.ADMIN_MASTER,
    ROLES.ADMINISTRADOR,
    ROLES.GERENTE_REGIONAL_SM,
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
    ROLES.ADMINISTRADOR,
    ROLES.GERENTE_REGIONAL_SM,
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
    ROLES.ADMINISTRADOR,
    ROLES.GERENTE_REGIONAL_SM,
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
  const [unidadeFilter, setUnidadeFilter] = useState("");
  const [periodoFilter, setPeriodoFilter] = useState("");
  const [tipoFilter, setTipoFilter] = useState("");
  const [listaFilter, setListaFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [bolsaFilter, setBolsaFilter] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [situacaoFilter, setSituacaoFilter] = useState("");
  const [vagasUnidadeFilter, setVagasUnidadeFilter] = useState("");
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
    const matchesUnidade = !unidadeFilter || item.unidade === unidadeFilter;
    const matchesPeriodo = !periodoFilter || item.periodo === periodoFilter;
    const matchesTipo = !tipoFilter || item.tipo === tipoFilter;
    const matchesLista = !listaFilter || item.lista === listaFilter;
    const matchesStatus = !statusFilter || item.status === statusFilter;
    const matchesBolsa = !bolsaFilter || item.bolsa === bolsaFilter;
    const matchesCurso = !cursoFilter || item.curso === cursoFilter;
    const matchesSituacao = !situacaoFilter || item.situacao === situacaoFilter;
    return (
      matchesSearch &&
      matchesUnidade &&
      matchesPeriodo &&
      matchesTipo &&
      matchesLista &&
      matchesStatus &&
      matchesBolsa &&
      matchesCurso &&
      matchesSituacao
    );
  });

  const uniqueUnidades = Array.from(
    new Set(data.map((i) => i.unidade).filter(Boolean)),
  ).sort();

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
    const matchesVagasUnidade =
      !vagasUnidadeFilter || item.unidade === vagasUnidadeFilter;
    const matchesPeriodo =
      !vagasPeriodoFilter || item.periodo === vagasPeriodoFilter;
    const matchesMetodologia =
      !vagasMetodologiaFilter || item.metodologia === vagasMetodologiaFilter;
    const matchesBolsa = !vagasBolsaFilter || item.bolsa === vagasBolsaFilter;

    return matchesVagasUnidade && matchesPeriodo && matchesMetodologia && matchesBolsa;
  });

  const uniqueVagasUnidades = Array.from(
    new Set(safeVagas.map((i) => i.unidade).filter(Boolean)),
  ).sort();

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
            unidade: String(getVal(row, "Unidade", "unidade", "Polo", "polo", "Campus", "campus") || profile.unidade || "Matriz"),
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
              value={unidadeFilter}
              onChange={(e) => setUnidadeFilter(e.target.value)}
            >
              <option value="">Todas as Unidades</option>
              {uniqueUnidades.map((u) => (
                <option key={u} value={u}>
                  {u}
                </option>
              ))}
            </select>
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
                Unidade
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasUnidadeFilter}
                onChange={(e) => setVagasUnidadeFilter(e.target.value)}
              >
                <option value="">Todas as Unidades</option>
                {uniqueVagasUnidades.map((u) => (
                  <option key={u} value={u}>
                    {u}
                  </option>
                ))}
              </select>
            </div>
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
      profile.role === "Admin Master" ||
      profile.role === "Administrador" ||
      profile.role === "Gerente Regional SM"
    ) {
      return true;
    }
    const isRegional = localStorage.getItem("servidor_selected") === "regional";
    if (isRegional) {
      return ["dashboard", "admin", "relatorios", "checklist", "formularios"].includes(view);
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
      const fpQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI),
        orderBy("createdAt", "desc"),
      );
      const fpvQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI_VAGAS),
        orderBy("createdAt", "desc"),
      );

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
            {(localStorage.getItem("servidor_selected") === "regional" ? [
              { id: "dashboard", label: "Painel Principal", icon: LayoutDashboard },
              { id: "admin", label: "AdministraÃ§Ã£o", icon: Settings },
              { id: "relatorios", label: "RelatÃ³rios", icon: BarChart3 },
              { id: "checklist", label: "CheckList", icon: CheckSquare },
              { id: "formularios", label: "FormulÃ¡rios", icon: FileText },
            ] : [
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
            ]).map(
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
              {localStorage.getItem("servidor_selected") === "regional"
                ? "Regional"
                : localStorage.getItem("servidor_selected") === "comercial"
                  ? "Comercial"
                  : "Principal (SM)"}
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
              {localStorage.getItem("servidor_selected") === "regional" && currentView === "dashboard" && (
                <PainelPrincipalRegionalView
                  profile={profile!}
                  onToast={showToast}
                  links={links}
                  onNavigateTab={(tab) => setCurrentView(tab)}
                />
              )}
              {localStorage.getItem("servidor_selected") === "regional" && currentView === "admin" && (
                <AdminRegionalView
                  profile={profile!}
                  users={users}
                  links={links}
                  onToast={showToast}
                  uniqueUnidades={uniqueUnidades}
                />
              )}
              {localStorage.getItem("servidor_selected") === "regional" && currentView === "relatorios" && (
                <RelatoriosRegionalView
                  profile={profile!}
                  onToast={showToast}
                  uniqueUnidades={uniqueUnidades}
                />
              )}
              {localStorage.getItem("servidor_selected") === "regional" && currentView === "checklist" && (
                <CheckListRegionalView
                  profile={profile!}
                  onToast={showToast}
                  uniqueUnidades={uniqueUnidades}
                />
              )}
              {localStorage.getItem("servidor_selected") === "regional" && currentView === "formularios" && (
                <FormulariosView
                  user={profile!}
                  onToast={showToast}
                />
              )}
              {localStorage.getItem("servidor_selected") !== "regional" && currentView === "dashboard" && (
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
              {localStorage.getItem("servidor_selected") !== "regional" && currentView === "relatorios" && (
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
              {localStorage.getItem("servidor_selected") !== "regional" && currentView === "admin" && (
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
  const [servidor, setServidor] = useState<"principal" | "comercial" | "regional">(
    (localStorage.getItem("servidor_selected") as "principal" | "comercial" | "regional") ||
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
            const allServers: Array<"principal" | "comercial" | "regional"> = ["principal", "comercial", "regional"];
            const otherServers = allServers.filter((s) => s !== currentSelected);

            let loggedInServer: "principal" | "comercial" | "regional" | null = null;

            for (const otherServer of otherServers) {
              const targetConfig =
                otherServer === "regional"
                  ? firebaseConfigRegional
                  : otherServer === "comercial"
                    ? firebaseConfigComercial
                    : firebaseConfigPrincipal;

              let testApp;
              try {
                testApp = getApp("test_login_" + otherServer);
              } catch {
                testApp = initializeApp(targetConfig, "test_login_" + otherServer);
              }
              const testAuth = getAuth(testApp);

              try {
                await signInWithEmailAndPassword(testAuth, email, password);
                loggedInServer = otherServer;
                break;
              } catch {
                // Continue checking next server
              }
            }

            if (loggedInServer) {
              localStorage.setItem("servidor_selected", loggedInServer);
              const serverNames: Record<string, string> = {
                principal: "Principal (SM)",
                comercial: "Comercial",
                regional: "Regional",
              };
              onToast(
                `Login autenticado! Redirecionando para o Servidor ${serverNames[loggedInServer]}...`,
                "success",
              );
              setTimeout(() => {
                window.location.reload();
              }, 1000);
              return;
            }

            throw err;
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

          {/* Servidor Selector (Principal, Comercial e Regional) */}
          <div className="flex bg-[#032554] p-1.5 rounded-2xl border border-[#0b3c7c] shadow-inner gap-1">
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
            <button
              type="button"
              onClick={() => {
                if (servidor !== "regional") {
                  localStorage.setItem("servidor_selected", "regional");
                  window.location.reload();
                }
              }}
              className={`flex-1 py-3 text-xs font-extrabold rounded-xl transition-all cursor-pointer ${servidor === "regional" ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20" : "text-slate-400 hover:text-white"}`}
            >
              Regional
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
                xœì½ÛrÉ‘(øÎ¯ÔöôTu€—–T$ˆ)`#’À µ»ÌªL²™•YÊÌÂE˜íëÙ‡ó°çyí¬vÖlLc&³c6v^æqø'ûçÖ=î·¼ R¤ÔiÝDed\=<<Ü=<Ü	!¤¼œE2º(;÷ˆçÉ£ßÏã<
½Ï‚dm\Ea\>ÏòévPý4›F×ÞÜYºu¤(Ðzdã©7!ETîhv¯H¿ß×›X&ØÆ€Dý2È'QÙ§Ý ×=o…þ¾Œ“ (^SûùÊÉ<IÈìbå!™]®Üï?"y6OÃ(\¹HÈ(ËÃ(çVŠ$(£•ûkk!¶RLÉI6žƒl^&q­¤Yñ$^b][y´¶æƒîª‚'«a|öôž›N“Ý*ž$Á(JôáŒ’lüuï¢€¾¤åÊ(KBÞaÚè™ŽVÖ;þ)8Š’èò¯ÁUÚ¢·/q:›—Þ*ïËJÞÇ»Å´Š|ÄÁÀÊŒ¢_€|>ÊýSxéš˜Ú«*þ3ßï¬Lƒ8!Ý½Ù8ÎÒ ñÁõ6ˆaõ~Lö!+Íþ¥a*í”MDÿz0“YŒ£S@œ(‡»ˆ¦³$û‡	Žº?Î¦_6oí?ÿl„Ø‡¾ãÙÉ—†¼Ð%Aa:?ÓWòy1òÕü;Ù›ey9Oã0£¿(z¦ói”gzw¾4lu{ø×Kw¿PŒÝšçEö%3´cìà—†¸´S?ã*ù¼¸ºŸgá¼\[=Æ~tõaÜŒ5ó¥áï–#Þ	ÒË¯ùÈh²r~—‘ý°½‚â÷óˆãBÑŸ³nw†SAº0x’ÍJwÈ‡èrãjv-&}v]5{ÐÊÌ!D4V›¿l¯ç+÷d•¡à—¹¬£iT”ù­Ö(x¿´E)úõó^@>ó^åÿœ…Ù_”ÃžEy}øÒ’wëgœ$Ÿ'_F%€=É&qðEË©êÇ—†šZ×~FOòyÑg' aDvÓIÅ_–vR%•èÉ—†¦Fç~FTòùµh/ƒ6øñ<ùË’R¦¡ÂÎÄØ—/O­îýŒ©ä3‹NePÎfd+HÃ8>½v¢ M~iˆÈzõ·®›Š6oýú”BOÅ¯zåU|
à„Yç©ö²P%[YzåeŒu¨ßUñêã?g ½)‰y`@nÚBUBÉXÁ<…ni/ªw%.®³8 ‡At4fR]50ÞD+õjõ;ÊD‘Ñ¼,³´ ß­Ú¸ŠÄGGÕ“$º ?Í‹2>¹\$ “`¶ò€ÌJ@]Žo¥‰ªë@ =½dMzFÆv5öÙ‡©@’xüHÕ‘ù×:P„ÝiÂË,’½Y”vO‚¤ˆz«²cæ8²uSX’Þì×¾õ¬H.cm…%.¢ilSãØª”€"§ ô@&á2×(A™i#J¬ŒAæÈ>>\1£$È=¨ÂÀ¼øôóÑ4ö²a\£$
7®’,¢w-J¶R”–ý Š¦üÂ„Sq
Tå|eº#¡¶i±2F
–S^’+z>ÈfÁ8./f¶ƒ¯)Ù$X¸g€Bl42àï9&ÐRðñ_>þwXé.Dª&Â³–Ÿ¬"/¯§Y™Œ×ÞµXóO~{”Åp6{‰DwRò‘)¤ˆ\&Wÿ€oª§¸ê²"«Ì¡–/.•=‰‚Î/¯.
wÒ2¿Ü¤FƒäŸþ‰t:ffzlçä¦ÇR¾ì’Ct09H»ð”¾Ø¸:GxEÿ$ÆIêv§t„Ó~Ï2²±±A:£ €iëéÐ`£pL‹‰C‡âÒ5ºÕ³ÈÔy(rÞÏtÀ\ˆ©y'f9iý¶L°eÒy7J‚ôCÇ Lª_ie%l('ñdãJþ4†Ï2>‚e2ÆÍg\âœ->Y›}`Q1+×0ÏƒË~\Ð¿ ›8Þãµ·°„°{V­³8¿f+ˆaf&`­>]x‘L~¶\(/­¼Õ‹¥#™n²:Œ½)¶¡…lO…ƒ)01E‰lûN6Ü¹Õ
ÓI–ï0ÕÇúûÛÇNÅb¢_ ÀöF?aýt Æ¬wG|rÄ5NÇÉ<ŒŠî¨‡=òí·d)N)‹ÃgPÕnXôOþQÇ¢žÛ11ûÁ%n Ð»ì,/Y&qxáaVLhAy	¸c(@þN½÷“(”§oMvxÌyž:²LM*—ïû ê Ös¥ÞyŒ{`ÑM/“Ä±²bFÙ‰g¥]+Í€Y–²uÀè„;#gõ’±ršSÛ=~«åÐ6Ìq­î¥ÉåÆU™Ï#ÏÊ¦ñÆ»¼ì’¶ŸrØÄ¤¥N9I Ã—ÅDrœ×Æèmª°kgwÈÌ$¬ç¾–ÊŽ² €ÅÈèE‚³ Np
åTn\sd£Ë‡èø”cœÌ·{ª©”Ÿ_á¶f1ÛÉ4Ë£³8:ƒ’ä}‰-“Ãlþ‡ÀÊzm¾^‰¶é6Ëæ(7 ïVÍÃp§qQRNêŸ³ŽS¡hÜ«zt¯Y†¦ñldøÆ<Hçï÷àÝ˜IzÕBÝDCoÒèœl£œÒë—Ù‹l#¾ÂRH'ÝdÏ:½vð/‚yŒc
1Ù™‡½äã-à¿UH(U[OV½‹J2¨°Z==™§c*#ý-¢4;Ã¾¼ö¨"‡¢?6Å±ÎòH>%>ö¦ˆircÅ8Ña/Z¾w= ²‰í å!Ž)¡åíHw: …ë2)7aüÅ|<†æ:xË(Ï³¼CWõY‡Uïäuåûìå±§«b±¬]Ù{è‘øùXô¨Ä9}‚]S½h¼Ñ¦H5zg€`ÀmL¦²KWÚþÁJ?V»…H@b!k'×Ço¶Z+×ŒCdÜ)05IÄµoâ€ê6`·ŠËtLº¼ÃªX#YO¢©Ñq¶’`Zäš
Îƒ¸$Angc(Ÿà6 ÔGËdkïÅ‹­£Ý½W‡ýÃá‹Ãw[{¯Ž†[G‡À>«EÉ[Ýâ'ìø
¡Ù'¾˜‚øîî²ÜžNd ¯BãEyVãR¨*NéIM˜i«ŸD«;`M…C àˆÁG1ÚlÓYW®wµ÷r,îº'Øû'”˜]gSX§CK°Þ%†ój®É8(Ç§¤ØN'JI 8ÇYõé:ÀïnË;ð‰™l/'N_úØ(_J¼É{œxt|"¯…Áþ)T¸o“æE„Šëˆ+~qºp-ÃJå`dÓ)ÞÔŠ~EB¼R£sùÂ¦R¼J«µÎ ¹¹Ø‰‚ÂJ€=úø§qÙËþÇ+V&ZÞ3Å°hJÙâÈÆ…!„H0ŽŸE¢uØ×1šÒÎVØ†*Ä\“A!ü‚ý6 ¬ø0^ ˆ‚||zåÓeÆ’‰W£X§£—Arûœrì´Ì3ùª—yÂ(ÀñÛ§‚¯-RÍ¼VþPK¨i•OœVp_O©)I1@+·¥ÞkJ‰ÙÓ{j$Õ–U¼ï²­lišM^¶ÙäKÄrÃEuçS¥ôäµÙrî²;¯ïÃ/{ÙßúJ€ÛÌrñ¿’¯æò'Ó,ŒÜû(vSå!]ËeuL\Œ¨Äç‘\nd¤du•üS–åD0,`&ýŽª¾Ý÷s ƒLú6‹t±ÄìK"?_ïHüGòe	Ê[tmù(|½æôòž Ç š"µŽ
¶q€tãù)¿xå[0ê®8±Rr±ZÑfFN—€dƒW¡–5W)0ÖèøÔ'MþtÚW3?ˆÜ Kúzg°*ø7¨]ÿè©EXC5ÆògõÌÄWsõÙS»V †º¤‘VM0¦†ªån†©´^•™”Úã „âùö[¥Ë‘©Uý1+´ºÄûÄ1MèTLŒ’­éx`'ò)´“ÅœØéÂN%¼«4•vÍXÉÌvZ¬c¦l:É³iÅ£Ã¨ì²¥Mõ0l]kh×ëõØw»¡1m±¡ÊccW_Ö6õeÏžþÖ®‰ª}×(¢T÷K ÃªíîÕµŠéî	÷,R¤=Ê.ªÆ,Žýžñ(FhêA„<+²^;g k„2ê£ŒiÛÑI0OD}¢Æ1¶·5;Ú‡·É¦ñÚçÚ¨îêï¶W	MÏP:íšäÅu­:ÁÕy*1ºÛóY™Ê« ”èÊUeXte×¿ýgÇ ôB¤öÔr„¼Kzf³§´t)û-ªß9ç.ºŠ…ÖQµH¦ú€1ÓùéãIt¿#’t<¤BA|%ÑŽcÿùªlÄe³ÅBL7g<ó×EmÏz	ìÙð$°ƒW{o†[Ã=Cë÷û'Ž†°»P¶·|$¤Š 9Ë$`³Í©².Ý%M¸Ð,I\¡B¤š)®pA—‡)`ÐÑèBM0|üÂ†Ö|ôÀ†bK´V[ŠÀÇ+IàS!M¨æ‰BŸ‡ÉQL|ésƒ#»•I¢OCM]'Ì¥E?G…ÃCZ´a‘ÂgóäÃ6”¦$ƒ“CøÙd±ÔØ3¨g,2ÌÍGÃÏÇ¨ãÉ§êüâývTD?°®aÇŽsòÍ•¿úk!SÃCs „0+6ß‹y`üº¾ˆéÜR’Åd'öi‰yÐÇ}Há€ë>Äÿë<}ôäÊÒOßÄì¿o3¶<šfg1¬ÿ^«¬þ ó<º	iÙ¦T^{q­Ì&“„ztãPÓÆFã$Œø–k"ÿÜÓÑÛÐI›ãð­\",ƒº
lð
9‡ßTäP³T7ºa’à ï`P†fr+ZW‡&§Ûî5ç`èfÂL5>%^MÛ»,7JÇùƒZc?m‹ž%jÈÇš«Ù±8çÿIüXÐK-qmÈ©JŒ°1QNÀHá°5¹ÅêX„*B6D¬!²­v³Ó«"@w@RÜ}^P‹~çŽˆA{ZÀ ºs;2.z,žÿ ×Z…ÿ’ãxE9‹‘¡{>’¼ÅÈÑ9Ÿ7àü)5y‘éŒ«à²…HÝŒ…’‚±”ËèØ> :´[ça6{(9MH5rì.d$®™ß_êÉH¿dæë8×@h¶Å*aÚãkã±Ì†DÆ÷)ÉQ’Ó¡’`é×;#ŠGÙÎÅ8Jº!U^uèŒ¿“Ç\¡\Ú
’b£cTðzòë£—/vñbÀNÀMË§¢¡	ŠKûe|/6ûÇko²„‰Š9aé´áç W²nÓs6Ap†!ƒ¹šYËÐÚ›€nyvÎu`û]JtÛo:%~ù ðÞè' õ´Öcœýkl|G¾³ÙVIhuB•¼r EÚ–JÝ´GLeeß n3µÆºˆÁU¶	B4z,ÞßÒ½Í¡¸Cr%‹‘Ço¬$2²´ÆT=¾gMAŠÈ[G¤`Ý³ ñÐtÕé%È {dP×ô%A`@¥P¤5Nù!´„Ðc«nš_éŸ:Lg‚ÛBS{¾Òe„˜ŠÞÄùÆÔ3Pµ1+¢©½ijc–þfÿ­¨ªÝÑ=vç@Öjæ4ò·øìÝÕ”1u6Ú|¨ºw†Û~è±Ì3vÖ‰Vf_%ÖE›_oñCóCMhÝ«|Ž¯5ÙOãQŽ7	äDóäféª¦_‹¤šÚÂx—FÇ·yÊ-ð€í1œ;[¤åNÑa&jmhÈ*6{«×—Zª]ØXcZMI<Á­“D'PI]çÉK×bCnm…c‘Ž¾Àý‹ý…¬yGÔ\ÛÚua‘9ª¨^+‘6(ÕŸù°ì®át¾žÍäFö=ýZ$ñ8ê®÷ê0oDåÊ©2Î­åMæ¦úâöFlÏ§9—™µ²Ì˜…QÑN§ÇÑLß<•®Ì[`w±*‘UU'Ô¼¬˜Î×HÒV0—úUÁy›n[…nƒJ¯lÔÅølmÜœ­C3ÕAD*üìÜ¦·žÛáÕÈ–*!àS òÊyKÓD¶Ò\qÁÉá6êæÉª¿öLæœ&y@©tŸ^h	…Z»ˆßu3#Ä•ðˆ‰`YéD?©ƒNËþ
Y[Ð%´ÜØ­ÊfÔ°ô¾ÍÃvå0H4jÁ„XüîÙtØYª 8ç;@VM¬@UÙž›aZÿÉ[oKP'Í°Ç|‚âÒß¬&Eâ1Ð:êh^9\¬õð -pŽe¦ËŠýjÄ´6ç:\¬–oITrQ5BËûµÇÆ—âC{Ÿõë¹Sh
mÍNPÌÇ¦½ª|°	YùdFM¥z>DéîÙóH¶ô4ž€ d
Â1=14>tµ“Iu>ÌOe)Ï"ÒHï2¨ÌžK´[0Fo·JuxÙÜ7ãlÓ,_ÓK¯ÖKYB‡3gm%éÞo[è¿ùÁ$í€G˜×`m€ÚÐ¾U•S 0F¬
ËV·¿ÿÞÐ˜qöpd·rÞs	=§–íý7W¢%ý¼%t£0,CñŠ¯¬§dlbqžpMâ	ÐZÊ’]Óí»ÿžâ_¿Wkù†zYÞ-¨Û©ÑÊV(Õ†a¸5/ÊlÊí­oaÛ ±RY¦õ©µ§§æ‘ßÝ¨ÿö×Ã£Ãáþþ»—;‡‡ÃwL«f¼02à·¿ÞåRÇ¨ŸNSƒ8»Óö¶ö]ZÔ!—Ê§[Ê¿ŒRœ©uOÏèó0^û5øÉ"U—kQLkf‰x»Å;Öz<€^ø¬žßW”À"ÒŽh`UÜÕ™7WS2Ã1Uª™Å•ïñµŽ‡|;C{4íßp.-8wíªnzðO‰ÒGÞ†8)Šçä
•-ÀRÕQ9>”60Sè´°P«OøY@’ž“çh´­rÏ‚0ÇëÙxÛa
3WñtÓä•ý2H°¾ÒÅÄ°¿PC6'ÔªLæ¥wi»}…áN`á ýý<¢ÅB’ Æ×ÈIó…AŸ¼¡%&A¤eœÓbgÁ$€ñäÉ@¢üøo1H9ú˜:{±lh™RšŸ†1»¬ÎÎÂŠø,(XX¿¾`O³Ÿ¢%Þ	œxr£‰a¥,+ki‰ìA[Á2VŸ^ô€Ø5…€—è	½’Ó8[’•¿•<‚Æ‰QŠ€œ˜ ·>WwCe[ÐYNioÂåš[™K2uT¥<l²æ6½Àa±ÚZiM¹g=T^Í¸§LÊ &ó”Z:šG›lq?–y9ÞzrêtE™.*S_Ýó³[´ëNƒì@«^ïÖŠj˜vÜ°nµ}œ³`­\®üRú-ñzAÁÐ«™†ú;ÏÎ¥k”QTžGQjzš˜+çè bz±à±ºžx¨yGyrú@o†yrJüžœ~µ¶F¨_ÖÝ4;Ïƒ™ái…*_|üóïç€"ªÕÓZ‹•#Ãêäp4W÷6|ŽB,-Þý˜²A¦#Ý‘Õd¨U$áÊ#îèH¼£óå3Dsúá:ö`€¹€lÒ[ˆ¨}­h¾A`»5Ý °õ+·–Ž} ‰¤ Éxãjý—×¶3¯'Ðpút—¯(¹lŸ¬Òtx·µ m³Å×Uø Óñè!èƒ›T¹®ñóx}mvñÖR¤¢eBÝBY€"ûœôÔÌ€pÐðûyGb*ZSá…¯ë5m"\d±NK37G3ZEcjã|e–1Õ¿5Ê×3êY 	Ó¸Xä"X…û@æÃOöm7ÁxÍÊNÿ").–	þ±³(Ÿlº¥‚íÓFƒïi†‘åÊÉš&Ç™\‹UÀŒlš±}}mÍv¹tó‰Õ¼¯Ýœzlgçi›ye#ôÍ«‹Ø†×ÓÛ—Mâ'À1ü7¯Fû—z}ènJVyáFŽÌV~ | Ü.|.ëpÌ}ÓÚ\ØSx²ßæ~cuÆ]Q–³ÃùŒ²°ÂŒ?€À±rcç¦#k+#ä94ÇÓ ·å†ÉÛ‹r§Ná†‘õã—Ž´+9¹®nÉ~½íLèÛßg›}ç^d6$/w'3!zñw.$køg"Ë‘†~Þ‰àm.:T»ÛÉà=ñÌ…ÎÃÆéÈ³â“â¬óø,"¸Oœ$Ðßôê ŒŠ,™C'Qö*W¨'@Ú`D4tOÁ,[ÿ0g¢6'ÑChó‹â!?3³5Ûº¶¨A»à"†I2¹úÙødíúF½¸(´–qËP@™–O¹Ïƒ¤ö«qùÙ]´>èr·ktzU—R{ä;8ˆ¦Ïã‹(ì®»þdQòwL‘à‰qPºUTÏ4Ø¢ë'áÌå² ,˜ùÐ¦\ÙR.nMõSP*ºDïxhîVèHÁo²·kŸ¢Àn×ËtQ*™?%ÆhîÎÅ@ï¹¿vÿaÿ~Ïæ¼+‚Ìp§Êò"¤8¶ñ¤Ñ™²qñÍ¸ž§E7Åb®l7Çmœ3OÆxÌùµéÔX÷{\#c´`ƒm¾·f.=³éº
­Œ	ä›0×'hß×Ö&p¶Ûî_nÂ<^ÑÛ@_š‰Ý|ªƒL/:žï7]õçÆÁ¤¿¶m±'ßáºD+¾Æ¹÷½½Ýâ«ZûÕMZK"øÿOÜ4kª[PÄ¦(·$ŽõQ&¾º	ú¼«JZNÞ|~«£ßÝnb«¢×}u3ZFÃeepÏÛA²>8ç—Ow¸ÆÙ\0Æ¦?¾fu7®¦?¢ÃâKÛXÛ,jÝRøæëº6PÛí²&ÐÚ—-èu£ÈÈ.ÜL$ôÆœ¾<èý5	ƒÞøu±-ãZ¸cÓÎ›š*Zo¡4 q?ð¬©1P…£:«	P1”ÖPÿoÃJÜV¦yŽR­P¨“¤Žh±ó£jõªÝ êžDü£¦s(Ý~~Z^ÛžÃe›ÔSYÔ™¹.ÈQŒ‚ä¦¡+,6z¨-U½Ù†k§Q­5v·æä­Ž€JÕ5ˆ­%àj™ÍVÖWï“Š©t¨—4Á´3pw~úƒ½˜ÚJ u!ðêùl^Œ¸hšŒªÄþºÎ’•_$h“n:QEœôõÊÕÊž˜\5Äë|åá/+ÅUå•±yc5±Úâg¨}{ÓËyRÆ‡>~”±7ÅÆ•æˆÏn@øŸ¡ÞÃ!«rY½#naíl¦ö'Có7øï™í& —/Ð¶ >Ÿsúîª³èÁçC{L‚Õ=V6ííŽ»Ûz¬pVŸ¼¬ó ™Á0
áH¤2VÙB\÷B<·ã®ê…ÍmWðÚ_Å¼kÎDÛÌºæ¬ø®æœ9µl˜q–‰Í÷¸Í|å|ýó=þÛœoÓYk›)7ÝLßÕ¬Kç£/ó±¹/ÚÌ}!ç¾ðÏ½'>èßÄÜkî[Í¼–ÿNæUXê‡½e`Ð[…½Ó` ¹0 >´^ÌHF
SÎ‰>ä-QÂôÈ´»‘ë·¥‹ò4
B<eÞÂù‘ƒ ÅU&(¾`¢å©µtÀ¥ûXêõû¦ÕÔÿôdµ<]°o¬äÆ¸ÜÔâ(»ðÇææN7ü†ÄvÄ©»õ]°°oèªÇ›Ûg›‚O»(Ì¶OHEi„ëG_=Þˆ×‹Ì¬KÜ-ƒã^´Œ zaŠ?°¨{Ä‚O…ËSjÖôí·^}t•ÎJ<–u´òXëŸSâè,¤í—Z‚Lc…;[Ž[™“ýúp"(ê~iýãªŒBKã—º‰Ð“ñ“úó<¹58«Ã8z®°¨Ç†®¼TP	]Ä¢uÔ‡Þ'ÚÁX$“ªÈµ‹ÍÊ“gYi\— ;S‚Cîbn|KÒr{ëòlOÊQ^êpƒí
èýÊ%á?ªÃI_yÜiFÌµkÐ»¨bAK7$2a©ðÛÚ¼Øºðhmõ‘}3 ]á'eè§*º•©W¥øPmž^RÐ¡0 ß“uÿUélÛu=ªh¤f;l·!j[beÈNéM·jÕÕîT²ø2iµay·¬»XÝ•ÂŠ2„]1ŠµS3ëGn¿Õ™£{“¨¦†­äéš»à=–jÙø*nïx­xÎhÎN*vŒºÑó›nÖÂT·ª›÷t¡
$Í`©Ø1V´˜1·¨T]kGWp@ûµþ#ß6æ1Moèå§……ð<|kPÌæù,i †ÌspðŽ~hxTþŸ¥‚ò‚œy¯jËR$ÊgÉ|“T‘á
3öp5ÇH*´Ýz*§ÁuM/xKaSQC3Çy5N«‘¹#r]ç-]ñÛP_q}=V÷[°â£ƒ‰]Ù—¡jù”ö©æ²æBíéŠ¤æ&Õ-]C_¨AMõÔÜžqùÜº·P«UsëòŠ“!í-Ô®®Þªo5©]r³. Õ4[fKMÖ¦yNoeóü½®ù
YÉ…ÜTZQÁ¢ŠÐŠjSˆVTâSŒV‚ÜÍ\¹×,Ô‰Ï¨PU=ôˆoîz5liÕ†´ˆZ¢J{èK¹ìúÜÛyèÚªì×mÕ®ý“}×‘j§5@[TÙÑÊK„Q‚²P",wGU§Yp½²%S`pBè„íï™Íf«õ÷­ÕPÅ‡KKÕèúÌa8k¸c@Ë3âå'¦	 x•6€¹ý8¨Ž½à,ÓÛ2{ór’¡[ëYæ´à3N±Š´Ó¬ÊðªgšËHªlgº
q‡XgïTRøTë(ŠÓûÊT«
®uàòNW5è3Œ¾òa°Ø²^¹í‰[+Kv˜ñgÉ!àÀÆÕ#ÿ\x»yªdKýèò?û ìÜ«(=OUØ¤(Åèð2.{VÐCôËÐxèûòv0¶OÈ©îÕ<I¤€íl;¹Ô–o#²l\àŽ"ôzlÑ½’n%€±š¦ÝØ‡˜ÇC•=‰‚I9F¤ëMJ"™§a335DqrSÕ“/»¼ä6à¤
s‡ŽÅÆU•ËÄž{:\ 0ÓÂíÁè í	—;LÌˆ TŠIz'f9IŠ`;Ë¤ón”é‡Ž±ûkþdÕOyÒ²q%šHÃgÁ2zi+ƒq‰³·øpdmöz¯Œ•‹…Ëú—tÇ{¼ö–ð‡]@ë,Î´ÙŠ”‹=MÀZ]|¹Ø±×ë–Œs8U¹l:[°Å!ì·Ú6û9p÷¦xsMgHZ‹ŠQ¥ø.ƒ"k³¬¦Ó-ßaÒõw#ì«XLù ÝÞè§Š@ñTž\Ð°††/Q·=1	ûÁ%õµáô€l%ïÂÃÿ›ð@ˆbtÇP€üzç›è[“¿çNm¡B]NN*—ïû€ð }ÞyŒûHÑM/“Ä¹ÈÌîª'žUum¬*fYÊpžÑvFÎêåaåª	&j½“,GÐê^š\n\¡ äYåÃ4žÂxyä#}'å°¨G}"ê”/ÿØö™xmŒÞ¦ U.õR¸v7œukTKÝ‚n\ñzáà,ˆœL9©WÇÆlÙˆó!ºóúêh\¨Ë·áž02B] 8¹ÐñvSÓñFÊQFÈæ°c*\›¯W¢qºÏCë²=‹ÅªyNã”Ê#ìö‹S¡l<BIõð^W8±šçÙÈÊAÒ08_0`Ï»1“¥ª;„÷·zƒ± [ÄcBANÂWí¡3+Wžtzíà_ó±LŸ‚CLã¡(l@ ;døo®*Ô»¾$—
÷úÞ½“yJýû¢ƒ¢7ÐuÏ0	fX;Çpü9Ë3å~ÚÛ¦É-Šã$‡½h4TiHšÇÝ‘Žß.ß»Þð {BYˆcJ{y'¤;U1ƒËM€C1£3ßù'é|úYÓÀ¼Ëòºˆò}öòØÓÿ±¸Ö®Ò€<?ë£ƒ•Q¢ú[¥ÔX±F›F42Ñ;..¡Me—®´-…•~¬6‘€TCÖN®Žèw[kÅ6dÞÒR‚‡zÍØH\\ŠY6ïÈè¡À»Æùò2ãÙ	³§¥E¶Ä[M	ãbÆ²só¢®-e«¿lÙî×µÇŽõöô”š’°€ P¡Vò¥žRSpY+õ£x«)rgZfû•|7J©™ŒÂÝÄ‹PÍQ¹£%1ÜÅÚ‚5“Î“äiÿÕ«â× —5úõmß2Š¼®Ãùè(	p±7£l'ŠÓQä!]ô	lFAç)[;Ú·å{œyäU³@PÁnÈú·­%F%ïàS ˆgc÷4;—ÅdoØˆd/Ôô¹Ö™->ÝzR{¨_>Ïòßr¢æ€_ÿø–G°iÄÙNÊ,Ç'ñX å_FÓ¬«+Ø9clW¾ÕˆHÐõÇÒðH+/‹ßÆåi?½S‘þúE–—Ýn°LFc‚qà©ÒJ{·¦‚*®¬?6Š<ÅGÕÅµÒ<…µÕO(+°•MgAñZ•»zse™Û`zÛs‚T(ëTo@
[`@¯™ÁU´HUA"Þ[aÊ¿©0!Õcàèv«È þ8Üß”!lØ¼UE5×ÂÄ!Æ"°Z3umb CãvÜs7®Ïû6ÃJéHdŒô¢ÿ^«³^º]uÄöÁÔuã=~ª‚Ò/cÁ°2Y©…=Â®O¡âŒ¬ßï;>Åx­XKž
l`:À4ö¬æ¬ntÃ„ÆË¾ƒA	]9P?¦-`œÙ¤¯u¤q`rÎyŸÍÅkT©ÌùBò;ËcîÆ‰V{“°G‡Ã;‡ï¶ö^·ŽÌh¼Õ] ÿ‰&žò3S.ˆ/¦ŠA±¡â»«ßàž.Dú*£°"%5¾P¥0=Sê ,ÔeZ’[ÇNrÎËÄ'2È²$}ä‹i$×v^}vW.„µÃë›¤ƒ)`b“'B%ªé)ÔÑšx‘.‹Œ`âÜcËâw*ëb-^•+)‘bx ‰V”L‘ìx“aDE
Çæb2{¿ÖwNØèUö}…Ž³<|"h!gê |œ]Y­sð£#Ð–¾‡ñ,É¦Ø¥mö‹§G)®ÖÎÎ«—<åæPvŒýû5üî‹1$Aší¤Eœâ—}|mÑs+:ˆ¹(1¾[ÜÂ¿lhÚÈÐÒ§ŠõÒ"Ì@àT´¨/Rq”ÏãTD|±¾ 8³÷—5cW1Ú(Wã¤Ï³S&ÕzzØBókG¹¹#?F¹VÚÊ¼÷‡ž§÷N¦ Â€eEu]Ï;¦Æ<PÈªáÜó‰l¨4÷.ìñ+Œž_Öxi¡öåm/«¦—õÎ-Û­_† kðxt>Ÿ]R©—îÜÕ;É³ùŒªŽ©Âˆ¡òÛ. Q”¹ÚôI[Õ²&èãÑÇ?SŠ£2©³ÿñßŠ•‰FäG®P¦¸å;ÁøTìf ´"Œð­F‰ºNSô0flÇ àv/äõˆúðý†à©1ç ßýû]?âÛ.+ÑcûöqÀVÝ™§å[3ú6ý"»ßÅ^˜„ÊP¢PQ.‰ ´ ‡úT‰!j½^Ã\rc½;šê½ßhÓ)lé´¤ád²0º1ÒÇÇpzÎ¶Ó]háO?)ÐØþM;a`çW£$Î57õÌæ'§V&ÛBa³îã@³æõàQ…ÿÅ_7þküv½ò@Û~,\]ÜÁ–ˆS –V‡g®ñíçÀ¢±‰4uìà‰K¢çtŠMäN–•ÁÔPåLchÃ)ò”QÁª”t"ç&KÑc„ó(ßBË'e×®ŽG•bÖÌ*9P£E™,IÍ,¼nªåG\#‹ÕêZV• ûHê³¿;”ñWdISãšÔ€~0G¨õX•j®ãÕÛ5ÁfËüÓfUãFQ§y£ý¸@vêóK1rŒéÍÇýRd¡mé”ÈÐfWµ×sëÁ‰z¨d\Q™$
ƒíº´õ³¤FCqÊ]UÏ‡ÄÖu¨v¨îäBÛL›~h›FûVµB‚…ŒžôáˆQi™©«ksœ#º|¦¢èJ!©Ç£“C*%¶K]Z-{È¼õ¯+Vç’<àÀ¦»êÝƒˆ§ØIxd‰ýòWHu;6X8¥Ô.1h´ÍMäÊ“ŸO	y±­jFÜ/01NýÕôý÷AÁW	aY¬	™“é’¹ž¤ué‚±ždO‰Õ«±3u*J¥ëaö£¸NêÀ4Ôò©Ú<áé9‚!ÓP×ªQŒˆ•Å	Ýz™±½˜)Î¹¨G]ªrC`Õ{Êô¥+JÜeCyK·@6
‘•iÓæ³0hÖ$/CûªÕì—Gß´-²`òƒ$þCÞ,È¬(ž«–›£Ë²áp¬•rŸbZB›gúCaùäññ?‰i–‡ÈUªÈöÀãØ9PÍù`Ç¿ QÕc9ßUÍpÔMÐ‘s–‰¥o¶NS:Ö!
°Ÿ‘2Zæ'$^Õ¡È9äèeqtÑ×Înu2ÑþDbBo-“
ÔÜÓU¬¬;á—ïiþÅnh¬â¢PÚ K@R»ï¥”(Ø†VµÑ‰ÜL°zÇ-qX~úSðx¬Ál–\0C>ì'ö¤‹6î<‘[EµáßE¹Ìh¿Ñ¢Q¬˜kÞ…€ÏÈ>ˆ oZ†áT…ªõÏà@H«yP÷{-êÉ.ùP\´ío¿×rL5oÏÚ7¦Å„ª=Fµ¸‰xÏ‹ß­wŒ½„Ö¡£³8“µ'‰åìÓ¸ãNñÊBï÷‡ÃgÿÓ«CžüÿúŸÿ2œdy@Î²ñÇ%ÿD4Î#;ž}gåÿüÏ¨Î¿÷ñ?Ñó†I´¦-ŸÌ“SX_d:Ë Ö:åÎ¢ZaP{zh’bX®q<NhL.×Pœø8Çq°ü¿ÿz}ïÞÿø¯ÿåÿ"Ã4Jƒù4ZDùÇ?’˜…¾.«*³ã aÎŸ²<…™Å`ž¢Û}jðÂfê-ö~8¦Ý)¢9¡µ$ ¿S6Ï*é±o0
â‹lpïÞ!ä‰V¦Aœ )Ê7ªhÒªü
¥¢'8ØÜ('X0=EH O£	=›a1Û ½œÝñàµ=yŽ™ºòñÏ“¸dç§Ø×­ýçä{òð r(!Á¿<½
Õ}ü÷‚Â
sÑdž†ÞŒq*2Þ»7df1Ìtz%$õeŒÃpïÞ~ÙéeÏ4ºpZ–³b°ºZÄ6~|…Ï«/²I–¾—;”%p1}Â5³ÐJá™ß£"ùF÷‘²è9ð#þZH ‘øø¯S€âÌ>,‘}˜ÙÑÇ?¥ÅÒ{±Ú6C‚Hc@UÂ…oy˜F-¹2òA;Pq’\â˜`ëŒ. 3:¡ò°Ü}:ãŒ`^@Ù—Å¤ˆ‚±l²WÑ­Cøù…"„²5ªG.Èevì.h»‚¦¯óUî"XEÓM´ÆÈzaQ$f%'A€ì€õÜF	[DÈAGèýÞš\ xìLêøÆðXÓ'»(¥£åPP…1.¶=ÉbpgÒd^°=<\ùø¯)P<Òe?™ãŠ*¾ü›+2?eqÚEv¥ ê{‚k³ñ<ÂÅJC)â’ŸX¤¾åài”øLàDÎ(p0˜è[	ÅaäA
‹?`l?tD8J·ú&NãÌy·BÜ#—øœÙœì]Óäïa–—þ.}õÿ®)#jÒ­øPn§–ê‹óHÄiž‚'Õ·vvvÈöA[ƒƒáÑÞ2VŒ4ûP+9(¥g1…Ò,Àä@ü”ƒ¶g¬nú¿?<<Ü@Å‡‡Ã²Gö÷Ž†/hÃWÛ»ÛÐ†¢Ÿcy9§Ï¶$šc4½JO£˜±¹ñÐ–îÿÇCNzï~½Cw^ÓM½úõXìááüþÇ!’uÈøfç[–¹wáß­]òrò²}Zå¬òÅ.ÍóRÂƒìl¿Þn»ƒ _w~|ÃÀŸ|0Êš}÷Õëá95B¯ÀÀþØ+¢0Fu-BB‰ðÈm,ÐT`çp@ã@£>šCÚvtÌ“²k2³cû˜ÆXùF6×>g_»«¿Û^ÀúêÐ “°bM"Ø ^°KðTbt%.¤%W&P• ÇöitR-½nÿº‰‰šâ¦Û•ÃƒõFuà”h‹Ôž®^…ìKzvs<´¼ô%+ß{=¹h¼°+­ÊfÕ…3òÐmØ1nªÂMÑºÀçSR
I€¶«²É¾f:Ãå3ignÛŽ)ã^Ý‰«ÕA´R{bC™™a8Š™[N½È¨1zM=Gþm¡_Ð+è3eƒ6‡±$Y½õ¦Dæ*eN$çGé%L3"MÎ6LÇÚÛoÁÀÚCjU¨Š\ÿ¢oŠWÆiÊã9ñêw]Ú\5BÉÁ`@:&h«K{ÃeÝ™)@»ãù$1MGS™"ò÷?³½QxüÒ¼Û4FæóS`„º†ÐÈallðÛ9¼/·‹Tô„áÄ$*‘céò%Ú,…­á‹WÛÃƒÝ½wÃ­½C³/Ö¨QqV<bŠ¼/}tY w[ŸX]eÔs ë·ƒóœâ<+Ê s©^ôYn7“¯¢óO6“
©ZL$ëˆ;‘~©u¥y¡ÉÏ1æë~¨uéDÝ+á=áMËE:íŽõ"í¤nÁˆßŠQ³g­1a¤šfŒøxñ©0^TÍxim#FŠ†!£HQÆŒl²¸^rq¥wqËrš›ÉV‰ZSÚþ¯O£_¾…6Õêrj:í‹ÎíÓ Ûùì"~%Çºyƒ‡‡7Ð#dÖ°W¼”¸ÀÓí=¶ò×Ô˜/œÏ*òè•f	{4‘G‹—ªë)¼Ú ‰E_n®he€FÇÖRyb·×Š‚ûnYì5“É
GøoÚ{«µ®Ž¼¾è£:Á¼Çq_ÜÑ`Ç{xMCÝVJÜè„k-_ú È)´O)~ç±]¦FÕ½Ìš)þJøTøPv™OÝ$1”÷jÓ‡¤®ì^Ï0âµèÖ×feØ_^b•VY)¥Æµ*i½âa˜Få©’ÁTãë©Ñ©¢8ºÇvfètOõ^e–½wy8 ±Ú9ÜqÄ]#‡$1=aWï«)}ö°“8Ëê¾JN16ÌóûéŠ&ÍÈ+?;ÒWPA¾¹’}×oüLƒ‰nOàe–E|%Sªôâ—é5 »×.­Ü¹À]Ç"‘‚”H*nŸÈ}ñÝ'Æ5Œ#¹/Nœë ˜X+AÅö(­´¤Õ/*‹ô}±C[r¶KS*ØbÛ%·Í©‡Òø¢Ü%å^ê{êD|kæë o@ Ä>K»eì¼f‰}±¹J«-«eµ'Oü>‰´œT–’ëã3š¶¢‰Œ–2`·bd×¶5a®é“\jµÑãNýÓ`"ßsE¿\£¾°k˜qîc³OÏE¹	®aÚzý^ñ/Z”ÞÕ°Å(ˆZGÙÎÅ8Jº!½ÓÒô‡°‰¦ñ8«ZM»S±š¤z‹¹¦
®'¿>zùb}êï$T…þÔYs(êâ¾"¼ã{C[{«´ôK˜¨QX:møyžMY§©ï¡jÃ[¼)­Ècð7ô>F7ÏÎùu/àÐK¿­p×%~ù,»-H5mw†ï¸7c6Û'‡ú¹ƒ:¡J^¹qÌ!‘æíÑÓº‘¢/Ôm¦Öøvbp•mÂ>	‹wÇ
ZœŽëyôêt+<¯,­‰+RÃ$=¢Œó%¼ŒÔ¢ïÆ
˜ó¥z TÇÝxCbJ«X¦+³#ºOÕOHí|éóïðô«Ósµ,¦çjÎ×Xi:´]4öšTc Hv
4Ä3TœnžìCc–Kt(Ü©¥“É„› VO’WmÚ*Ü+19‘Ï‰Yn–jÌ"ÀÜãÐ4”
‘Þ:èµ/ÒªWKmÖ¦UªØƒÛ76#Jú/?5µ,e_ïXøÖñŒPa[».ÜP‡fª¯GÝôV%xÊö‰ßµòxkÌ]{Úk¤}ÍZÒ~›†]æz0!¿{öúÀ§]ç|jÚ:íÐ.»­;\¡'Û¸/WV÷}yÚ“63À¹AŠôÆÏÆß4ëhtî¼²¾j¸Öf1Ox}9ê¿Šò•L+?. ;y*!7&ó{õÙDÛ3Ã×TRàCEÊ²E¨]_{l|)>Ä³™ýo¢)´…L’¾ÝšùLRWïƒOYB9WãÊõT]bÀAWºò$TšU‹sL%ä»”Ež­ïðqÐ)ªˆñtJSÊ4õÌ8!µÔ
Õ}¬T£eØ%ôÐ€BôÚf|;vó¹~W†º0«Ô+f95lc”5ê^Çåï¿7ø`[ñƒGn+ç=÷—sèL¥@Ñ’®áƒaV,¡¾zØÍÀ÷2ášÄ“4cZÈ¯kzúï©a€!'ÞÌ™w%g¶†5FÈš ÉyH6P;sØ~¹òƒ-pµú9
FäUpO˜Ùw«lÍ!´h|Atõ~~
éd¦ùx¿‘â:~ŽÎã}a‘1ìÆùÊI¬GŠu}¦»q¥w¬®æõÊpcYé¦b÷‰ðÜ}¿ÿÈ	ÖŒÞï•?|ÓUºIË¥/&¨þYë~SÅ—ùADa0ä0K&âÍtß = +D®lÐº„GÞÏ‚|ë4ÈËÂû/-îÌþ¶Šíßõê¾à$2Gf_æ²¾}¹“÷".ÊúyÃ”C§þSpöì™žÛ¯*ñÙt¬ßLh|¹&yü#@;Šé@½Þ'ÉD{}è–û¯ÏlA‡,øòGÔïUL<S´XŒsg:µPêè@]\Ø~þcà6®Ð•¦Þ:—åQU¨öÕ…:,u(ÜhØT›úzËqø;LÏ.·â|œD÷«û-b­Ü²ëx Ø¥w6„®´èÕv^sRÑý$;#SýÐú¶½>
.ôQtÆÀ^ÑaÜøiŸM'×÷Þßó£œíõ¬ºû* `MÿÍåHSš‘±jîÓUóƒ½j¬:´üc¯Ú¶åÆîÄ(zrúÀmŒqi¼qH	uMG•±±}!žQe¶Fô¼‘”¬¾8$ÛÈæÅ£yÌÌž»ÌÚÂ³GµzúÀiÍ{àéô•á…©c‹Š Ã¼fe¸è£+‹k_3UK½LÛOó¢ŒO.WFQyE©Í¯ˆ¦qmRpK€¯*îtÿQM<0Zgå×ZÀ~:çÜþ7‚fº’¾þ0ÀÊ£)ÜÍtRø°! YÄ(Ê´X%š{QEäÅjÑ.Í‹ZD\ ÚZj¡VIB´føY…77¬èB›Ð¸WH‘çiy}Õ!ê!Ô5%£ð¦¨µ¬¶Ú½Â]K¸¹þ»ªœC¯)UkÕZÚçÓ‰Ó^ÙÊ“ˆõi–V/r¨¼r8­Wcç”µjt¡–á6Ÿ[®Ë;X•w±&o¿"kÖcQ^ãruEÎã°<P¥„•ï«£ïUæªâëýà†àòdtY—¯œñ`æ–ˆÆ[ÀêÞÿº#~lõ‰Xîdñ³<Çìkà9:OE??Ã&3ûÌ›Ììo}“»ˆ&T×î(‹ÐÊÙL++Œ×Þu•z†i«,ÕLÆ»PujhŽ§÷Gr]Ec;Oñþ–«2 j÷ëEQÚ£ê¸U15[„½­µÑ·2zìõõÇµÝ×¿¿þ86ýúÓdßot¤ÁÖßhÔµû×ŸÊ; úSsÀíVÅÝ c¬¾{ê±n˜u!£"±gy›ü‡O;» Î UÉº"¹&R³RÞÒÚ0ü­I¡Ú+†=|Ç~2/8»qÍ\•iz·øµ³ÜOù«¢¯>¡¶>¨<²ßß.ë.\\í<EÓ|e–ÅúCûûzF£ßUh½uˆìò´ª­ðIŒö“Ü*/gH‹âÄñrŠO0G³r£Ó¿HŠ‹e‚|Ù€&QsÍ+ÝžÓ·Ýè;Û@LÎŸ¬Òy[˜@zoµ¸
ãÝó„W<&NÝ:B¡°³0r4¢‚èÃ ƒ2qyIEˆ{Ý†Èp…ª¹–©ñ®ú¸.>j™%’ò4\•Î¦íc°}y²síœØ	]Àd#Í§Ã?2TÀC‘8Mâ4òrUþ^t^Äƒ…<§aQË¹-ù)DÞðÕ9ÿÀl–<ZŒ”´C>vM ‰0+vÓR7Þœ
±:î· C-Èòvvž¶%DlÄ‹Qæ*ž°†ÓÄé‡Ä;ÇÓP;Fx`ž*üà?‹óQSjÏy
òj°Úï÷«N…¤H5%íF*´ÈÝ•î´"géšá¹9êh­^˜Ü×Dí9›—¸AŽKQ³0ž´‘ƒ
µ—vÐÚÚ^,é@³¬dPº¯T,ŽŽ†ëÐ6À0bí}% q(F6£¶:ÎS>¨'«ìCC~]°xª½´,."*tžŠ_-:ažÚ)þŠ€lRXxQwÚÕkH¹˜mµŠTö¯m&9ûÜl¸f†ÞVËM/ð•Â­Žþ>ˆÛ Äˆ ù• ¤‘þXv+-é{ƒð©“Ô²ªZå.[V—ÝáÎÚÑ6qÜôTüjYP; zª~/DëðS:Änƒ›2Në_^J£¤^ËÙÀk©O·²)5bb¾÷Z–¤×WŸFÍE}³¸ s/@ø c_¯š·‘¹/VÐÃ¤«-Q"÷œÐÉL¢“Ò§±)O£ ôêÌËÜ+õ=2„¾GB<^_›]¼Õ4ó*óñ”ºñD”ó8ôªxG,Ì¥£ÉùÊºïÈ—ªÒ±‡éŠhèÅQvQ}Áƒ3nøUÊì©‰iªßœçÉÖeóñäJÖ*TŸk9tÂŠK©H…?¹ð‰JyºØ|Á
Î¶nT˜:í\%’i^¼
q=]p­‹× nË.^ºâÔÚÈÕSö•^k0N‹œThtÔcévTàÞjŒ"ÎµˆøA?åJ¤t9U~UÖVw–*¹ë¡z«ÇZ¡ÕO%^7@w”•[è!ÒŸçÉ»ú K=ÂùëÝÞè§Âshy”ô?öà¤¯V2¶Þ4Í{pc÷o\¢û¦wÕë¤ê)k|„T=ÜÇèÀï´Znª õ8+ñ?×½HdéKX‡À8<ËÊ®€HC¡ú`ÌÞ~,²(å±Oå¢Db´ŽúØûÚQ½äO’I•eèÍó P¶>¤zõ)A°Ù»VäfKºŠ"CzîaoV+ø›'å(/u¸·ìÉÊ%á?$›æ×í;Þr";ÂOC€Ÿª)5íÞ”ª°@?|°¯£¬>²µæþY¬ä¢Ê°b'¬ž½zÆ«=ë¥1_•ÔLzšª["6/T»<u>IV¿Lv©n½V~«´Êlô²ÚôÛæÀgŸÿàI@­“sžå5ùUî«Gwã^×\­1”·CÂ"Í5Úug<;ù"zC-2ÊýSØïì[µ·é_Ø—ŒbÆ9Þ4
ãùT‡ç/ZÏ.õâ»ùm¼<[måÔf”\6j’ãã‡&u#l3FUŸ&WÃÜÞýhùÿþ÷ÿÇêCý˜ÚªaéÔ,|ü!šaS‹x
Þ¬DÐþ ¢G·ý(}äã<füŠú4€×•ãûkÌZ ]Uz›ãÑ±,ÏrM“Ú(ÑáC¹3ø§	yMé¯quÈ‡M‹É¡´È‹0h‘miÉ^hzØk®¥Ž=Â§	d\$Ã§C¥@¹5uœ+e­Ãö·ZØÓ‹ Ý4nÃHÛ‘ðƒkOè{˜E”×¾¦ù¶Yu¬áðá¾ˆO½ÜF»ñ¥Ðïé‹þð“îNÉ ·´Ñ í¡Þ~¸xó›t£[ZÎmÜÇ<ö©W?Pß×z{zèâ¦ek´Ãˆ/Ðõ»×s*ÔûOÜPML3ê»¥Jó=VAzÖuJÓÁ‡~zÖ°t›"!V=^Ä­qzØ¦åï±­«q$ØÏíiæÀ×pís#Ö\ˆÈIž®±´·CÚ9rë„êÝe„¾ýÂ®TÊ
Øë7¿k\+P™Ç­ŒÌ4¶×Ôà±>¨¸“¿VHX¸;Z0èÏÖ·ìä"Æ
fS>Ãýûâ{rÍoÝ~}·g,–þòy–ÿ–UìúüªyÊï¬†,¯»É¢ž›iñ™FW¥ýÂÐìÍ‰¦ÇŸiÖÕº|`“-cjçÒG=ánXöfÿåãhlqsûS‚¼&ÙZ^*âŠöšubMLéÝ£–qc­%:5\^S;¿RêUŸ»K÷QGm–oÄV¥éÕ7åé±Uy#ÎP^Ñ²ÍWä´áº—å7:]qR# r°ìÊ¦SlWN9fEýªã&gª¢T=ÖE?VKML‚ªÇ¹	è×µCqYÐÔ9¶*Ëb±‚ø»})ŒO¤Êí†mJ^·¡ùWµú&ø=ŸÞq.å™¶m/VKÂ*ÈoHª±ðývÄú_f°ÂŒU~fíN ½–+´¥5‹mšz[h3{–‚û(ŠÓ;‚w5Ãæ?$¯Ò×\yìÔhXj}ùýì“íæM=ã,9„ý}ãêQõüzØÌõûŒ7â<åÔPªj>ªÁË"I·³<Ndÿî€í»ÊGÍÜËgÔ¤sñ»gÃ4ù3Ú‡Ý JÇVW±¢o¾™t”öñER?ÁåÊù®—ßŽ
ø+^r£T,@ë„üól†t-÷¹ˆf§®¶¾ßEœÂZôÁoœÈÚ2)ÆA[ñZÿWü”8`0Ê­ËrëþBÑE\.ÜRKs`øÅÍrÙ¹^´í€Ú½/Ö4¡›na<rMõ&![ñ¸î©r*Qg/p¥ÆEÛœˆ4¥W_å5y¾è*\ºøüïð/µëêÖÙÂþÆ¢3Úî¢|àxk©$ÿ‹ð!ð°’ð×ýj¿'Èz«ËÒÃùhŠhnÆíößB³ú©÷¹Á«§¦Š¡Y˜µAE+(ñÕh<¾Ð'©éËÈFÑGÌ·T•³$|ð",ÙÊ¦3ŒfZIô½üåçzã¨<ú=0
QõvÈ%d êzûšŸ	Ôˆâ,NãŒ˜åÊÊZmÔœÒyo…WÝBáÚŽª}{Q×?½>7mí?ÿ2P§Öê¦˜S+@UÄwªn`6R .ûVS]ÙŸÑò†h)â=}¸)´S_*‚*mšaÉ÷3žªoŸOýnUgëÑ´áÐßFÃFƒ»¿,*«¹73•¯ýj”·sƒPYQ{‡•UÜÔ5‚VaíÛ¶øð˜R×§’ÍÂaq5þ—º}~ýÌ¥ù•­É¯cK C¿¾¥Áýnä1q>õ³@v7÷žòYˆY£¹ûÝ ˆ:tûâ‘Äòe³s1 ÷×î?ìW^7ü:°J„ý,X%Nš?5Zi'Ü_<^}h¢… þ,˜¢™|©\”iñ3û$¿}"t‚AD4¬R¾TT´íi~FFùí“I˜ÿÞlû)äË÷ë¾áR·ÉúÛÃÆ/ç(pÈÁ¼‰SœŒ0 Ý½uU“TÁºi¯Ðˆ àÛ‡ã,*ÐüÂJÒÔlÖ¹£iŠ5è=“­µAj¸;ë¹«Y¿#"ÈØ9àfÏ¡ð”“0#[:ÿà©ïZ‹«ÃÍ·Áð±)‡nZØâè"·Âða~v Q²±Ð/|ðÖ-‰wL¦A:’ŽmÜÿ,bŠÌžÖ$L<Â@St¬}©ê˜æÓdÊIóøB“û6S…0öÂ<‰Ó°Íe]Bºü ³û?0GmÆÓb8ŸuæZv[Lš€Ü¦øÅ×ït6rÔãâs7ûŒª^«‰Ò91™åa@Ø¡9ûN›Ë@øØÛ»ŒË—áõÑ+ïPÝ	zÈ_›/¶C70/µY!]ú;ÔÛMc ê5LŠìDËÁÖÝ16‡ÃáËéPóöæhöÙÝŽ'ÔD•ÃkáQSïÛõ½éÂËÓäÉ¢kïPú.Ðt_sÉ*Üká£‘=øLêc3CMµ€-c|Úê^O"¿¹›ÖÅŒB£t!æä
Ñ|/S3¸¬…>qM8íçOšK-ä‹=Ÿi§YDÆO2Ò’O­ñ£ù~ï£,û°Löƒô$‰0ÜÝt™DåØñìn?µWæZP›úeèÁ6^#‹4­–Eqv!l­Qeûð«yZÞ/ùDã ÊÎz#!ÿ8Oã4[Ö0§Y*Ñ¤Ò©â]Kä~á,¨	rõ0Zº7*¦7~ÜåÊ#8‘å@³&Pœ/æUE‡ª1V¥6Iç0HÎðÎ¸'Øƒþ nkfõµ«DTš“!^ê	ØÞzzcXä+Kýê˜âu÷’j‚sâ‚õ\YU7BÜ›1N¢ÚOV+¯À<‘‹CÖ×¼‘**.Ž[<-Ï´¢­+qUnXMèÅk	Yù/»wÐÙ=…Çv®Ê›îN0MCJâžvàp1öMaö’µ ÝœOO™üŒ~”7Ðm¯€Ìë<éVSù&X§9¦s·³ó~†c†Z—IçÝ(	ÒêÁ$Q2‡.•-Ú¶ôX„Ú–Ô”q>°TÒÖ5V­u–ž§7®äOÜÔã0ƒ÷2‚·a
‡ç>ª‚¨f'IÆR2C¬ÛÝÒ›gÊ)‚½sµÞ¢ê¬v:ìÒ\ûg[_ÆøG’(ûdžŽ©ô·eÊÄoâèœ²³!çR³ô(‡?gy†q%ñ'¿o¶-3mÁœç»fê¼ˆòDÓÐ™ ØòqçúË>¦Ô@"O ÐÃYŒoˆ7ê,8ù›M‰X3Ë÷®Dtr`!ÈŽß>VÝ¶l*|I-“rHv1aý‚H:Qžgy‡âËY‡Õ ä5ô~Ÿ½<6GµÀHðí‰ÙôS¨ÉÉc/H X×hŠÂÇhˆõÝ”Üai"‰eTÐÛüç,À`Ñ|À4…‰ÜXÇ+Ÿf!BáÇ#
ý½Ã£Îc‚×+!9H/É5£ý„NMã"z©O‹é\/à/ë L¼ÿÌ(â
ÈëS…›R+q1Òy9K0šBÐ@[×lQ2’y¬¢ë-›áóÞ~ ÑVÔí áƒB««ä×Q2Sà8žpäƒ‚À–¥—’ÂŠ¡$i&ÙçÙ˜ÎYÄüGA3Ýp@0~¦Ñ9ýÙíi„†½„®AŽ°u<Dø_á½Ë×+×ŸÂÂ?…,Ü1ÍùÓ 5¾'ë½þŒ:ôÊËî} èk£p\ZEy?*JåQ9ÏSŽ»^ùæŠ6? ªë÷˜ëšÕ„ávpYÐá²Á«Yƒ€«@±GQÞ“îÙ¬ñÏ Xœ—î³$.»Ñ!$Ã48ÆYBü ':Úé˜£5ÁÍéÙ+ÚVÍñÚ[yÃÀH_Û#+dÝûí¾(Ã;…íôKú"S‚ã6Áéâ-Õ³¡ÉnE<Ëƒ?ÄI¤{d7%þ»ÀD0-ñ|À¹%¤y1M€®WÅïuí÷Ú[eäÂÆy”‰eÅîÐEï¤6BšÅú£Ç?zÊ@óôóˆŠ~ÝÕßm¯.vôQò¼ýñ¿ømK	W@O«Df™X/ºëž
´+üëk¸+û>¬+vAÕÿþÑ£o®ø+‘Ø–ùØø7
<„]P\¦cA`ÁÃ!ý)`ØÍÓœÅ Ýî7OO“ðª~çA\ZdÙª_ÖÛS¿V˜MM&Q¾;E¿Ô€9¯tº	ãgÃAM:$D—BÑIV'»A©§$¸ŠÌ$¥
7ÓO³7Y31ÁÎy†JŒMóHàÁJ„Y¾V‹Š]ºŸ±Ïl÷¢ VøUµ†™íÚ!ýqz’!*ü.ý]úÓ¸¾#ß\1 0¾¾PÊ¡ @À/È"ÐºÁA«ô°È¯³œýêõqHÑYŒëE;† ˜Bt\ÆgFGo©÷™ëzŸ ‘2#Ï·ßx8ÊÇq$‰ÓýÂjÕAC™:Þ˜¾¼8²nVåãßžï»Ê6}“û¾+µ‚H_‘)èŠ ŸÚêf4<<Cöé6åÖ@Ço{ì°³;§X0ïÏùAfJÞ¤("1ÿÙç$OIœ+@™„¼ÿnx4æõ«ì, ÃCm Ë¦­<Âà;Ä­½äã—Éw€&¼Êdˆëï–~—¾ž°ÈÎäùØI“1-ˆî¨Ã;È©s3 +:þø¯ƒo®pº¯±ú} ÊIp–å¨Ì¦³ =H†ÎÍÊ 9
­ªþ{%ˆ0ã¡cÆÀ…«„D©u‰S?F,<òš;ªR(UB' áþ¾ sà§eqÈ}~ÙÄ”æ¸Ìe^)·¼ðÌ$'dˆTlb-±kÞÏ³„9n=Ø{±sØÿqçðhïàÝëW»ÛÃíœå¹è{föLãNÃÃ)ë%"£Þi[DÅqó¬.Fxõ²b‘îá<`ï™h&Úiƒf³(¡„ÁEO<†* q%ˆcÜûÇV·+ñÈî2X|ÏþuÍ÷5AÉî+J¶/	>édQòxOn®b/eÄÅ¿G¨¨vµÙì˜vjò4L`EaÏ¥”©¶>—œˆÂ.ÕžVÊDêuŽëï>ä»9A3jnƒn@ŸC+äOaWœ±<ôì$þ=Þ-é‘ìuã<K³ILƒ%+qÐº­—«Â¶}ÆåxTnì|b´ÏH×ÅÆväË"\U$kkïåÎÁÖîð…ú¥Ÿˆ73Ä'é¨]½“6làm°ÕW…:¶H%ŒÑü.ãÐDG6‘Wû©ÕATÌ ÓÇ?žEŒ53h×{j¢#7-?µžz(u•^d¡.¢ ö-ˆ³Ú¸t„Žë¦ß¿Û/@§k—‡%a…<Æ&ôÜ‰¼¢p^_AGMcÛÆÖ	K›eE¹Â°€Z t
¦QÛ99Á UúY§©³FükzkƒÄ%‰sô‡ñ]¨Ø´Ñ>ˆŽÓð€wBIC†š•J‚¢<˜§4Z`Q@ßaÉ¸Pá°@îv0Ç;6®wºvêmP©öY}iv~OQ«€BÈç]ƒíÙYW„.„ë‘'äÁ>:vŠ›ÓfuºhÙée¡±â}è)‹—L)äa#¶n¥ç[z|êaåá|ê]\ZYYAIELY§Š³Qå#Q’j˜Š	²[û' îvpùŒ j5œU¦s%±e²²ní™ÚBq«2Bßª¡ª5ä]’T¿y°>ó®‰Ðýs~ˆð"šŽ`B#ØëlÒ"t:frl³8…%ÄÿGËdkïÅ‹­£Ý½W‡ý­á‹WÛÃƒÝ½wÃ­½ÃeA—1¸¦Ã‘úº3 nüÍŽk]@Psò×/_k¬	ÙãïÝû*øµé2º¶ÀnjsæmZ³b{“Ž;¿60„­âôã?7µ´ uO™7V† ÙÁ4@‘&ó¸)±Ý‡Fº0ÐyÀaùñOZXrÄ®ž¹²„W M\ju
5Š­&-ñUÌ¤+¶ŠÇ+¾Š§YŒ­,ØR¸×`yqaW‡ZƒÐkÂM
¿ÕÈ·àëGÅ*P”hÆÊ.fz‘=·Åã3AqÓìóýº’*	éXÊ‚…C¤l±¾ØÜ*{UHxn/N‹Ç+V×‘T¯8­hQ±ÚìÉ¢dÕ/ZjÚz+QZ<7ERANV…`íá)¦í¨j=½¡tÝº¸IS-Ð´CÙÓŠz¶E®Öâ¨ÈœEÀJ¹XUCo&b²ÇFí7ãíÒòÜ¦¤O­Bl$è<â$âôOÈ¡+\¨"ªs“G2Äöh£f~uý¥-ÙÜï´‡›C³\èäNp‚®¥¹€]¬•n†˜{'ý¢ÍóxºLÅYG“LãÏxSæy4>hÂ/D¤1;ô³Pó7(ÔPÈ9li†bKýÎœéï%Hâ?dsX7È@v'jïº×¯ÒãEÆÙ“‡³ €XaÐÙŒœHÄeVÉ²ðëŸO2ôá Obø˜ãBœQƒáÿ?‘uS,TÿÖÂ•Yöç}ž>ŸyŸ7Ð¸ý6/XImÁ‘`Ä™ƒÕ°ùÃ’H*4Ôk%£¿43`ïþc¾sÝsGÛ'–qo¯ÜÍõ¿×†J¹Œ§”¨Q‰ZÒl^v+µÌËTgkš¥1…3Zå¢<­’Ù¥-“cjÑËÀÛžf§v\Ð ¢Ïéš`–Z‚ncù„¶×Á€/DÑ|ØYXLÒÙ™¢ŽžQš°E·öùkŽ£7¨éi—ÖDûÇ{¢!Ö-;Õ´û´+á*	­‚×zJ}aªF"k‚COs­Nei@«ìŽžRS2.†aH-¼èöbäW÷Dcìn[ËZÈ5–bLœi§ÌÍ”Ÿ2ˆ?©Í¯Vw€¶à‡óÑQ0¢5Õ»Qo'ŠÓQä!ß¶Ž ÃëÕ¾ÙÕŸÅÑ9u¹TÒÚßÈW³ò±^/TËtÜE£^4G.¶²yŠjz(´›¤ËìÊÍÓnB×K¢ß25
#Ë§Û&ÈÚc£Ýg†@¨omØ­v™Áód±vèÒè\›ñWâÍ@FÉ˜Áð. [ÖéiÀ@¨n‘'^§Ò±Ø«YP2ÆÉXTd–J‰9†*ûŸOŸFeÀ¡ˆ ÁŒ‰
T†yvÓbœÇu¹fhÏXÂ”ÃhÞ¯æAC0¬JXõ±¼gA’åª–ŠÖh®½|„Yeƒñø’_ócÛ»ùûgäÂ
*ò•ñ,“‡ÈÐÛÐiOåoèë›¸ˆK1y‘yÅ@ïµõÉì9]]Ï£Úù±YnÈ%­<YÂµ¾ÒÕ,nÇx\ŸEq¥9—ûÌBL«Œ$îpfû{/÷€7£ÁUß¼*Ë‡%U¯´dæÁeÞ¿YÂê]’õ·_î¾z÷rxx´sÀ5¾lÏw__míìì))¹"«Í`.\@ý¢åkÄ5—´q`Ø½-1Ý§žÁ=W6#¦+c ù]@¿ÖöóoŠGô¢óØ*$ ñÔÇÐ/ Ø¡šü?>~Ü«³úÐ.m©fômSóÜÀ¶K­„7°¬¾êyUº¯Ýdíìhçåû¯ž“&ÙùôYÏl)e´µÛ6_"ž­ŸÊí}×DòR§å¥ˆ€—ÌA6´\‹o66²›þ“Ln¡M‡dî›uI¯é”¾A6uIË»x‡´Â5Ý1÷â¦¹ï’Q¼¦S·´ä­ÏÊ§Wàåòš†Wq$I÷U¹É6Õiä^dFñé\h«.±¼7ì+\Ó‹áõÖdä±I”Ã7UñÊCŸ]6º©šÍžpŸÝ\ÍhWe±âÞZŒ<tHC;0¸‡;÷Våä³!äååÛTå·äØ½èÇ
²rd§i:Ì@Òì6›}ºç¹S-ì©-é29Özaè¸D¨^Ú½f*Eó©ñ««hŒ.	JròÑ(NâòR]l*È	¬MjÜSàõ}´]Y†“¨€a‡Ó8]PUôß‡¼OãÕ™Ä‚Ýõ–Ì:mÙ²~™½ÈÎ£|+(¢n¯#ïFEW]7sô#Oë \ÕÂ•<Ö:ˆÕ0=©Ù;Æh!ÊSÄÒ77Mu¨Ã‹õû `6¡«j™ªŒÝ¤,1¸ßU7}µkšUÙ„£ƒÕÚ1hRsÌå–•%ºvFwµ_o5+NÝ=/À„ý”ïtÝ¯1m7è®¨åå•ãud^5¿Á-¤¶,…1.C¬pß¨Åµ&=<Ù0ËËõ%D†§Äj““@æÊÕÓ§Š.,\m›q5ŽGÑkuraVÄ×¼Äs%ØÉ´Y;QÇkûÛk-lWO“"áâç„/£ifH”ÂZ½¤Þˆõ-…ëkuú0f!réýp#«Øv„\I zžZf ;æ€‚5»¢ê}l÷é÷žB¥ O]*gæ­ŽÁßùXºYèãñ2áÉ÷¤Ëoºœ„•¬I'kÆÕ ­Ýg†qã–q¢EÛZ°mèËŽ“ûïª©Ze_zä;²¾¶|Yö£»w×ÑÁ_g­cœÑåÑåx§ƒhœåáqCŸi5ŸÂ¯®YWÌéÏò`|*æß¼ B/Tsƒb¿d-ÓÓ’®z¡Ð@/"|õ¬îŠ½„óz®á´})(:èJiŽV,,˜’Puøh²M£<HBôGÚa¾bŒšöaéý6UÇa£– adÔ!nesšp¥ÏÔ²’è‘àgxÕzfŠán
C#ÝÀI«Ž1ËÄ-aÆeF˜Ç°'±duóŒy>FXðá¿e¾…õ³eúU{§ùô„Y”cÐyêæÁÆT<kjÂRU×uOë[‘å% Ý2Ñ>ú¬¶JƒàWOök‹Ä/6®ÂK¿ß/4	¼²ÛE±ŽC·ÙÏkqv­/¡eFáþžEig1ÛÕ,º8ÑÃ2.R_‘Úš‹P—€iÛÑI0OJ!Jè†cÂ×Ê%z%Ä­]Äê'6Læ4T`”AÞmFÏ)œ7gÏlÀ«ß2*¥Vs#2¯ÖL¥¾C5•õÍyíf\†U«-Z4ÁrjThÞT1]Ý&èt|-è›^eš*hjÐZ178G	«2ÒÉLœPÄ>éYšË²DéyätUW8'Ü>2áF]rDI€+šÎº²¾kë_\lÏg	Þú‹„p]€¦ì«¸Åše^¥Ù¾™‚3¡g«Î÷Ù0u¬ÒZ²¯TRXð›š¹®¡’Ð†©K´ÜÕ\·óÿH¢4µÓo~à-õÓèqBÄÀƒjtî|NY
Ù·.µ¦êÛF­ÚÛÙ·:¶ÆNÄ ¨’´ÞÊasc1Ž8t¼Å-³%5:'ÐÄê*
l ¡Ï…’*UPŸBF0Ñ!4AS€o`ì<Ršîóí7=?è—<
t%)
kÌ™±¥_)J5¶¦òÄrøý'h™…Ý ~ûÃ×D¥ÿêƒŠ¶µ©&ö(Xñ]¿6Æ¶dÓ´IÞ{fv#Vf½ù“ûfËÆÑ	î´ÿÐ‚ÜD>ìoÄŽØšpÐÀ8ÓvYñ‘‘…˜y8n ¢TŽË¶Ãçyðtv`(Ç<Ùh"{„EøË¬ÆüºÕz>*…›}AéKˆZÁ6y6$AÉ†€Æôh(k¨uìe0­èÍ‹Mºu™›êÄZÿÊ4Ågúª×S]’­ç”;˜ÈÆô<ú±¨È¦ÒLžÛ8¬h\°U6D¢ ÿ»Îˆ›ºNSùmº,¶ùLÅµ§×U'ÕÞóhÏÁsÕóç>I®9Ó­?5´¬6ýªâu+OØêÎºjO±êO§šjN’Z5žüTT×ûO€ÜÝ¼šú–U»“×žÇ93Úß1o\—£ö@®¤	EíþovvÔÇZÞÆæá$…ÅæÐPÁ<¾3FÍåÉ”ƒ¡ÈšL’hŸ_]–%èt)“nM›[Ý1Ê_+ïˆøAK]ñs<ÏÑeÙž¸ÁJÝ·ËÀN…ÞIº	‰.Ê7Tg¹äÔu¬õè­YLJZ»† ìT¦æÅ¨v zP%?p*Êüë¸ã¸ÝrkEVˆ7é:Çâ½Þª'šcÂIÃ~@9›eÖaÖä°Ôžæ¯|ÃexÌºMöÂ«}³ƒ/ptï|°þÈ¹ŒˆK øÐL¿Ö_¯]ÎRò~•¢ÆÂéÔüšÌ³RÕöÆfSë)Ð¨Õ™'ô2Þ]‘W“H 2u™‘H^f
sÒÊ¢žRf3¹Mà=ÎÉëRÇêR)s¨
lü‚ÕÌ×H3a tF„hÐövÄFâÉ¦@4Á=Øÿ…f…‹þN™"»-yºÙz2/Ö«:j˜¯NTyrVTÍjBÁÃìÔ²ùµËahZÓ V^	èÎ@×=» ¯§Kw
]käÍ¿osµb’D]Û‰õÛÔ^
—_0áW	Å€Rp³:•·]=ŒA–'m—LÝwãî|—ó ˜jÙ¾—–_™Q³&¨ž ÷¼„šzt@¾¹’Õ]¿¿Õ,±Î7Ïcú¶#<ÂQs¢ˆ“e¼Ëœðë§€i€¥?$º@pä¨S:¶Mt÷-OŒŒ‹ßl’BÚfûI‚Ù©‘ÌÅ.Ñ×Õ[•€ó€N¡Šñt=äé Ü¹˜eÔ%™{6ÑOÜÈÉ8v¤8ÒÒI
ŽLf5:´ŠÝôãŸAfêìC}3“ztÀ¿¿àæÌÒâF|Øc²%»ðÎ¾»rºXˆf5ÿ?   ÿÿì}ÙrW–à{}E
á© ]$¸‰²Ls	Y‹KÕ’¥‘d÷D¨VHiHTI³ùóÐó0ó2SOÓ1/Ñ1¯þ“þžO˜{î¾œ»d”%—owYDfÞýÜsÎ=«nV“cf5\•$v”­Ñ«êáE¿uÕŠ­ƒ{¢À?P‡·ŽÓŸíÀã±Ø©ab	Ù¨Žéà¯ž>y)ßŽ
@ZGÎ>œIOh¿gÇ½×[o4ßxh,iÇžf£§éXäA‹©šeó<ÈìBÐ­«s
¯ë@›Þ—3çÜI£Ãê’WÐ$\
µ _Pcò|–á3ÌÉ%JHÚ$MòÆÑdÂÝwtDï]iÛ|j²¸–K0õú}þþ÷ÐékñûI™¥Ï†ñÞœîOÌ¿•µ5fú+[‚zBñÅa¦vÍ:£ÈÊç_C&Oð©VÉ@Œù±Ý¥- fU Ç°3ÔY³U>d¸âè+Ÿ°9vÿæš»•ËÙ}yL!¾îôÕ™%C“Ç™Ž“ òˆMáßê]ç2³€BŽÉÚq˜£¹ª3ÌŒK«ÊBgÅ@ÎØB%¶Ô^ŸéO[+ Ï}EžÍ”Šv'×{vWV»…îd½æ´fø`§´×—NÙØ$øË<6Ùˆ±¦–5+LVÉ!×Ü$ G°çð‰ TœæëâMìØÇÀ hëàœ,ê‹_?”ì'r>Ô	qAe©É#xÄ`›W”]Ÿ’¿ôËÅµèR¢ÓÂ‹¨ËxS¦ Ómz4e}í{Û`6‚ùÕaÊRÀ(X;sæ7B¸üÆªúáÕ@ºE×× ‚aèžZn}e¼™½+§Sû…Dúc• ÒËb®c[Ô‚]%¶”º`JMªÃ¶1¢°gEÚ(RÕ5õT‡ØOB8s[ëÎ^ã
Ž”Å\zÃ|Ö}ûÙ•jÒRÙÈ›	+Ü§’LÁæ$=d#å#´Í"ˆQ’2“»ü‡?¬‚­‡Â·Ýúòwî_ŽL,ƒŒM¢§ëŒ§,å}C^Ò[äŽ' Š™b½•®³òlº†|ÖcÎ¯ß"jÉf÷Þu-Æ¸ÙhœµalçÓMó~±q¹qGæÒ½Úü<{•ŸdßæïË3¦Óü|S,“]ýtT\ÐÄllÓŠRÖmlCö[žï–üénw.FÙlHhØùÆlŒæ)ÞÞÚÊÎ7NË¹–Ù×Íån*eÿHXj†Ïèja1Œœ¬jÔWý‰y=ÓØËwh‚åÞžž´—§RÖ’÷š©z-¬¬û`D_ÐBÇÔxËÌ×mt&þ¢íl9)€	¨iYí÷H+2±°xd²¦jtÛ¡ƒ¯óšÜ¹êùn6+"k¶}÷ÚÌë|@Ù?S±Û»Ùxo$‹vò±n"ÝÇ»OJ0·í|‘‹è“Ä¶§8x" ÖV–â(ÎÁ¾:«ËAÿÙ äfFÐÈl¼¯~îd£3íçnv1Ò~Þ¡[ÛÊ~ ¬Ï}2>kÅçå|D:|EmjÕrØÃ³ÈS†5´µc§–„8¼Rû-5Ï>”ð±goýf£Ñê&å¡J›nÏ`!Öý²îŠÿ€…Qú’cVÆë¡s£sÏx	×ýÎ?Pf÷¾ä0 P«ö8ÐÏð±B>ç•ëtQ“Yv°Ì+å›”p©ÃevþA?-ã.ÉöœUËŽ[Xi'Üð&ÀÇþˆÜ+^Üé{]Íb+nb9ú$†™T´CqÏ÷XmhÒƒAò±E’qê8Éè‡»zË”x€}§NÙA¹KÚŸlÜÎ<äÒižtÀb¬hÄÄéN 4 p!AT³ÚØî:“òÐŒ]d|JPg
çcÞô»â’@±ˆõ³ôƒ‹®â‹Ù¼<…hÅóó¢˜°å¾˜±˜ã6ÁÓ 'ÆÎ’²Ý&ÂÃc÷ö¼-ºm·Ô9'°9$ÿàvº .'AÎÇÚ5¶yjÙZùæ`±é3x”€˜Ñwr¹¾êdß0Â]Ý]Mh,Ö@gäj~„H8Ü\ÿ§5á©ÞÊ”ñÈ+ÌÎéfe’Œ`oelâéˆ ‰a9 tÖp¤qït¶Y:CÖ«1„ mvþOÖü»<›_ru•—ÈÇ÷nc»Þf×¾Ú ÷n†çÅš3<äC—–|â”àI~Y-æòÒ×–$¼*§Ìî@Â}UM³=ûŒ­Œ.PËÆ¹jÝ­uÒ£ó
1ÿ(DçHŒó áùFÃó¿u4,ð¬voâÜ&Hsþ#MïãçÚµ_Á$4–,"$r Gþœ|6ìÓ¿ëêþ6¤}žQqCTVÊPÉ‚¼ìºÓÝÌ#XB¤«ˆ	ÁåBNÁ‘øÎm„D÷ñ ÝÛƒáŽsàaL(íùR–=å“J·“ÍáÒõÔƒjöÄR!Ä
MÛ/Q‰Ë÷‹Œz-}Ê³â/‹rZ Ã˜&¬öÈqçu>M!Âq¨2°“.'›1®qºîˆfEqE´*˜´ˆ!6²3Õ0'Ù$ÞQE¶úsWŠ.
Æâ×Ù8SÜÉWø*Š)„½-…°áÅÆˆ.z°‘S Éˆ#õ0E¸¢p®ë;xºüCGæ.ß´Ü`üÝ`:‚OqƒŸ`úßdõüÏÖâÚ³Ûî^›ñáÜ‰™÷ŒüíÊ®ŽfÎ—ºZ÷¨k@‡dm£rA/wìeEÖd˜‚²máKJYB–šRœIô®ü(?)FØ
ím™0kt›j Œ5òOÏZµmwÕ\åUQÏªzcZAö¯ÃñßMi”0˜ÑxÌõÌ>®ù ëKì‹!Iå¨ÀÎ]ÞïÓùa§w1š]¬gðöRjìyx¥[ƒ†qP‡óÚÎGGC÷­ñ‘ÑmƒcçD‘_K”×X;	p6Tç“TX`3öÁ~8||€õF¥ »IŽ8ž0ëºfÐÎïyYõ„óº'¶à¯··¦opÞˆÉb
+½ŒO`D#T\ð¼˜þr–×‰0j­Y]þÊ÷vØY·ð™ÉOfÕhAÖ|TœÎÉBÍ+Â-nîdÄè\.ésrØ™•€?€Qô½a}LGäTÉbõ!5dÌª³ÔëõzX¡C’á&1¢°M·Tñ¥¬Ñ•&ä´±'$äS²Ýä°×ÈaÇ€xGÝRDõ³}°$õ´òê»í§\á<“_à@nßûd+ë~ix²Qâ§€¥Çušâ»í„±t·Ûr'R|ç‘}ÇÐûï8r¼ª)µÚbB£q½ªÀÑ‹ü?YðÙÁ&û‘˜ú²@0‰éÂ/1å}R¡)Ôr-6„ýÂåGä¥O«ä$&Db0ñ±B¹S_ø6b¨¦Á¶‘žëW× º™‰%Áö½ºÎ/{\ƒ`aøk#ÒÀ^&Ù)P¾®ªQ‘OÖÖÜ["&§fÑ7‘ÓeŸ/u¶<ç
N.Êõ¯Ï§x°˜Â%çJìœv¬ô,~Ö©Ú?ùÃÅˆ›©t¬Ê¦SÏ‘ñ3½ÃÏæHÿ•Ü„æøs¤ÙË¥×—N7GâOíåÎJŠÒÐ¸#†ŽÅjN*1×Ý|!zò6rÂxüj…»6Ô4Z›^É¥#Ï²Ãuˆâ¦˜É!=_I%ýßz´e¶*‚/8ý».¨¥# ƒƒ_Ê	*7Áqü“r<unŽPüRàí›÷zÆ.gàŒïŽQáP}•“Ñ¨±5Ÿb[;~44Q¸Á%RÅÖ›mß´­6ÏÕ¯d£]ÙÖª¤‚HÍ|ÍJ9–cðŽY™’¾Ý¨*6,‚`Æ‹&ƒÚ"Œ¨â	yAùPj£­~(]è—¤:yÛe,†ß5HB]š=ÍéU=Ÿ€¢Š§°¤Á¥aX¿¹ÐU™£0e°GžŠzÇÌ2žw+Ìäzº52LÇ{¥öÖ¼Sn{ì“*æl	oªU­¥3V”zQ‹¸€e—gä‰V“†Fô"d™
Ëú·+‚g:Ð8LC96RH…‡xÐ¿BY®Uç
®i×ìg¤ãe`[õ,a›qvð+Ú/ƒo]¯g(«}ðà1…õZ–Ù¦=¸”~šnüFŽPMù¬Ã;”D¦Š©•“2ëÎë`ñ[±Ââ$hQÒCæÊfD'(z¼*c ÕÓâøà…!àUœÐ¾nÑ£ðZxt^·ðØm€“ ´ÀKáQja
ì¦ÃEìˆ—kñý¹˜ODË¥s¹¡pdn[á³l7üAd¦žDXÉÝÌþ”X9FšˆVÚwbO`S@SiD&‘’06:<ml»‰ ©SUxX<Y*ÞL,*VdO)iT±âM5è-5Á*VŽÑFV°žd&±iD“²Æ'¤fm7+¬µÄa¤­‘f¾ÕIX$uk¤)0ÛmMàmŒÇ×ÆšCÓ¸bÅ
ÈÍÛJHæŠ$x7o0šÒ5¡9ºi`™ö+ž\;ŽÒr]cåibðïMæ˜Dj~ìà4üY²ÛMDeÏƒ¹~Ürã]ñ„½HC~ßÃmÀ6Þ±-Ê¶5o×D×Ó‡ßcç øõat'àRê—Ró÷Í®zìSÁâ•ƒ€!¶Jôg,—Y¬t‰^Õùl¸‚5Š9ß¤JC”û÷+ß°.'ï6p?n¸¤4—zOÁðqz¹±¥Åk¡Ò_%ÿ&{Ð2ÜóR‰~¥ôÎÿ}Iö1ÓVE3ìÙR†Ì®ílmmÞ‰J·´
¬5þSkŽ?áí­H*p£Ÿ>‹¿ëÜU€¤yôy´IÔ;p€ N¬!”±tô1[\Ÿw}È·þëE9-Ï°_¼Y¤6îôè®Ž<âL°˜Ï+‹ª–<“™xFŽùëÒz]/xÌóè¢ýû¹Gô{#ìÊ¿Æ¿ã•Ã¿Æ^‰¼Â4œêÃ¾†,jše-.œ„§û[;+ü. }__^[ñ¯M ,!¯â–ÂJS8a¥!´ˆ®ÂŒè…V¸•@fåà"`ÆJˆô²BÕs¤)?û%æºr—ûR<¬q@å,6Ì"|!'o²ÿ>”"§„>ÌtMÕô¸}¯p›ìÎkðÜ{æAöê–žæÓçåDPÌ—HB b*¸ý§³‚Zx¹lªysoD'ÖÌp©…©Cù=/êŸÿ¥xü·Œ…@_¹ÞÃ"2ƒ=´½:¥¬x¼_GS°p¥}ÿ¨Ó¯›ÕHµj–·Ùgñ¡¦×®ßú>Ç¹Y[AŒ¾o~ž=-æ9õ˜æ=TÈPãKaàWûæ†Ô;‚tì-„ÞG¢ì_E´.ôæµR‹Zsìð`”4àJà¬£ExÒ4ŽÚ{Ê¨R0lpÓÄÐ4>Ì6§Êä:³K	Xãbýãj_nÔI¹¡=‘kVzQÌ£9KrAMÏ¯f¸ƒLóº©ÙG„4™Ç@—f{Ñ ¾XÉ |´Ml!®Äðî!ë
ì«SLM¥6,uZXjÓH‚½Ÿué‡*	5Ëû>¢Ìø¨ôÇ,ù´.FÉüâÒæŽ±ÊøåàD ÙRLmASXIUh!ØÀ†–_ëì£•0Ã¦e,2™­qæcÜpw:;»^ø¼»ÆDkHqÃ¢•L_¤‡3ÂLç­ þJÙx´øa>¯&¤N g[Žqq|©È‰³§œç£²$øß“áb¬ztUÊñðÊ¬ÁQò×õÄúí˜ªèªQYËô;r$cÖgvöÌg'?BF/Ã?Ë£eå?X2¬ ½G¬ƒBZB•j„Ûy²ä„¢¬X™({¯§zjâP‡ƒÄ¤¤ÍÈ€ Ëà',©“pW¦Ø5~aÅHzá/Á|¢Pi9,NƒbjÍ„¾øv£×[?lO/~Ø!ÿ«ÏNòîÖ:ý¿ÞÖÎÚ›LÈXÂ71÷°@Ë’ã]z«JLfEJ„èd!)±ÝMÒ‡"Úš¥s¾q'jŽ	Tõä_$þÓ"=œ IejÌlPy¼“kDí±ïÆÍ“õÂ”•h0›½€±²]Jx½¤í;Áï>NžÌ±ŽÅiÌë@Ø³þ0¯ïÍ»[1>Îª6¯¾†ƒåïK®JÖý8[U/A‘…õiP•´Ôiµ©<òj1ÇÝKñÒjË—ÜtšÃœs9/XBZŸ†Ü-1)Œõ5ÎÊ~)½F­¸¹ãbP&lÉ…´–Ò±±<U70Çd0õßYœ/#öAz1m…ç7¯ÎÎF2Kü½ù<*&}a]´N™˜´C¶¤­è”qíEá9¦—ÔŠGÓØ©Ãè'ÇPE)ÝÈ‰FiŸÙñåR‘s*¡a_èæYâ!’FÉ_Âd	¹–8Ù$:—}WŠ·]ìÌr@sù¼üË"¯¡GÜZ¦è%|³zjÝI"Uk‚=µê<ÏëyÙ/§Õ‚º—Ü[À‹"™š¦v³^Ô¿LÃhÚ4–SFËAÏ…æùõö6Eð€_7æâHÒU¬I—km¼›ŸgÙõ,£l ”¸ð(>ôÕ\=hË¨9Œˆ°o³°é”óy~–Ï²iÜ…B¤ÐÎ×Î-WÝ×â®Üyñü‹ÎEC¯ïÌÌ¤Ô%·ï¨-
2•„Yp€UmrGBïçd°©œ‹(­ØUSˆ .)0!I³5c…3-jÆ%Jx!‹Ö¦Q³÷Yóš]¢RÙ'UZ^øEé`v§ã7Q¤qÑoÉ&©"Ö–‹¾XÚßñ´Å.EsQÎÆMO” ûtgöI•D.A”¦'·9ªH§öªD<OÓ¯\é·ýä‹#å„£L/Ì,ÃøßU—³ÎÅ¬? yÇoaÍhýíõìÎzvw=ÛÞ²)}wX7¦Þmhw{ÊMéö‰*Kí“ì¶»¹¶)C÷dsc³`ÜaóA#^ÎniFí›ÃR”èü6—ü”Þà¡(•Æ½¤ír¬»­ŽÈß4‰oDà›qÀ'!_»4'ík¸¹™§y_¨s_aØ4n³b—q99ìl7ª‘_Í4B<ô›	»ÔéµÉŽ:¡áÂZx»0e4JTïm+H+5j~:Ž³­'ê[º¥vÐ»°ÚÝ.+%_íˆWÒå®%ÉÐûX$»4Zç€kµ[Œ…Û4ÃÇ¶ (¾À„¥q„âZd¥=“Æ,aµ˜S_œI5)ôØ†Ûú_Âîp1Ò-<[ÌëhX,UÅÁmn4+û0J„Ú	ÂÕu ‡`ÛžzMíŸ@G¿ÿ}¦ÿnå²h™Dïj&Ñ~¯£Wz/?SÝ"&)>nz²“Ù”Üf7fLs¡™åØF(çã=Â<õ
YŽj1õŒø”ã³ tÍêþá¬VRóÑüðêí#ò]ö,Hö‡lÛçÛƒ,­Ê³ŒV'?Â2õaUÂç½.N‹º&d¨"÷±KÂ³TâQ¨b5ø,Y’RVdÀó‘ePËÒûqzYFŸ;?œŒòÉ»ØŒ`âÕ´˜(™T)€§à)'³b¾±Å’|æýw›··²jš÷Ë9¡%Ê6ß-žZYÁøã¸’º"„G)2Úå?‘J"cT{O;éT &!¯°™ÜX‚øsµTÅÐKG•Aœˆ8®È2Ëž±Á²\¦¾þq?ÏÐL;a¸ÿ§ð€/Š`¶hg	ýîž15Ü4_–(üöâÁÍF[ÂÂ¹1–Ÿ…³di¤ýIBjªD	‰¯iiŒm4—£¨­3ª9Ì¼»ÆƒšR‘ÌàO“‚8ÛÁ"³jë;Ù½eÖULÜ-¿4Ê´;±„)ŽÅÉ!fz©À¶f¼|IQn—‰q«E¸õ!-àíªDI(x{0Fz6’P6§¡6	Ifð¤$Íš@ò’„ÀS“°7*þ»ýÖ%£fpyÃÿ¥Œ6U$×†P›Hl³³¥#”ØmœC¥!üa–f¤=¾ØÈ„_¦žGHîzµGvHåÛ>öÅïŽåKRu½aN7yÆõrÅ„`3‚K	é™Ô	z`_Èy$Ø˜´BÝ\Ð‘!$®µñÎ7®­˜½KòRÅZŠ-¥öt“ŸÏó“QÜ^XX£âTš9PåS4‰ÍÁ|Ht‹æµÏg7{âNž÷¢	uèxL¿E~ç#X¨Î÷\ÞÌhH‡ƒÍù0µ)À1ZPœfUU`&µD¶¨ô:G!ç^«¹ÿ«ääyµg“æ'Õàk¨iæÖšOgOE6Àä‘Oæûü¡C0÷S]œ {9þùÀÙÁ H&k*)Š+±Xb¼¿°Îr_Ù&ÑŠÉ³»°6ŠÌ“0AÄ(š¢ï–Ú¶fÑ¹–ÏÕ6BWË]­£tEâtµˆÔ•%ÄêJÖÕ0^—.VPž af=&4†éJ±‹ñ‡êZéqÓåp`“ôPNQÉGë8NYö6[:”Ó*qH@°ÂJX¼ÂJ£«µ°‚[[ìVºH¦åhXµt¸¯9ò»¸Ã87 ¨a¥½¸†•¥„6¬,ïöºd*"54Q’ßír)‰Ô´D‰~¿aYM˜0È2­JTƒ6×V`ƒ6ÖFlƒ7ÔPxÃñŠpÄû›äC²éÎæ¥,Kd†‚Ò&;¯—š!
JÃ,QP–È%ª7Î¥qÆ((Kd‚’–9
JÃìQP ¡Nfq	IUšf¢‚Ò:‘‘*at¾¤U	Ucy,Vý•Hžµµží&°A–+(Kgº‚Ò*Û”xV(²^AYAæ+(-³_AI›ØŠ²`±¦ÚdÂ‚²t6,(2bi½.“JËÌXPÒ¶¨A†,(KgÉ‚Ò*S”´)­(c–ÕÔ²Y³¬æXæ¬ôGrk5¢…Úä÷‡=Ä„a¬ì§Ñ$÷WËÙÁÚ-I‚t•Ö2BV—e¥Ì ¸YÌÒ—¹Už3¤Ó”\gþn›²9Ï´Ûç=ƒÒ*÷”4ŒÖ(”äAƒÒ2”´‰µÎ‰%œJƒÜh´ÁdûW0aFÅéwÌÔhRó¦r7§iÞDá•pgÍ—È•‘žb~àßÔânÞ8ÝwÑëb r5ï’Ã»%VüneË,yzÊ6úytÍ[ê9p]·O¥ÒÔZ‡ö€wÀ^€Ê¿œæ“Ã«½kc7“,yx3‡=LªUÿ<¢ŽnmáÃÊŠì|ÄXýF¾ÍÌ…/ò)fIAƒ‰ŽùØ5 2lw¯JŽØ]¸t¤‘åE1Ðå¥É„á?Éûïu5]ƒ]ÑO€]¢Ðƒ	D”]œ3õrBPP>:¼º&úûÙÖz¡	·Õûr#?ù¤“±µ¶e­m¬Š©ðeTÜàiœ_lœƒy.bPµƒxµ˜Ñ¡öpãõ—[ï‡ol´ˆØÌX›0Ý¸°‡Šg 
f–ÄòÕ!VÈbÛ}©ê
]2AFvy¦:¬Þ‹œ¨ûª/]ˆ\&Š°ufæ4ÍüÜŒ#äžÝ÷|íáyü8É5&š™¾0ŠéAÿ%æÖá'Š¬z šo¤¡jòrq2.ç‡WŒ	a¿°©[ +"œÝV‡ã’‘z.çiŸ=ªgÈ5²‡–Æ˜¥èuI%/ÙÌXÆÒìs§ƒð0äŒ^!<’÷2Á•û¡8	n%«ÓMUHºeë\ñbùl_ûtÁ	&vÂË€e‰çÞ‘:PoÃIöáÅ~öõ¨œÿD†žÝ'¸®®²j!6ìñ„ Äù‚J+=iƒCÅPŒ¦¬	%nó27ÉYöªœV,Æ}àPã¸1Š­†ÄÅKi`Ò@‹h€uäÂbˆ‹,Ø¿yA¿þA¦.§÷¾LÊŒ½Ä‰Ñ=Õ=±š>4&ÛÔ¦óõàÒ	zt³:_Ã#öo¬zX	ß6iÏ6œóûrÒ_€‘AÖ}6eØÇ«>uŽÀ1ýäµÿ,¤Ù±NIH,$ë‘µÉ…Xs&Æ7KHç!ƒb’{<ŸHÐ:&4¨A$®˜8>J\Î+–ë˜ÿA	‚HÛ"°JÙžl›ÜÅ‹ºèS¡¡#ÚT¤‰}ûû0ZôoH@p¸jÄh…öÐb»/;GB˜Ð¥vï¨×âòÊ9•”É¯Ä±pyGwKícs=™¿®!ÿ
¶36ä™e@Aóc*Á\dPDC÷ ®ÜµÍS5ž‘È$¡-¥6îi±Ô‰¼àê"È ÿµ)/8N‰Ù•ÇíBI [nÿ‹Oó0¹üü/`uÞ–éÇ2c‘ËÀ4Ú/h\2BÙÍ}¼\»nîg2#)oîÛ _Êñ‹±¥°>Ÿ“1Ü@LNƒ`õî† øUýó_ëÒg(µ
ø—ã ü:àÉ•õ/xJë‚¿)ð2 R
MýÏ~q™3ó©½Q¡3¿¶~bRççuþó?åLæÜF¸Œ¼øÀ [3è+¯Ÿƒ€JñˆT-á)t³JÄvºGF*ºû… âƒ&y{“àõYör¢Èþkv
FhÇ7¢}—°b#ÖF¶™=zðý,{QÌ¦äòþó_ßeÐ~J®ýåËè•]g‡W†„áA	ý–Ð+»Ó‹<Ä	Ñ{Ufåbð=ÀÛ,àÎë‘k¤86¨¸WŒþg¯½Ã6¿¨ 	ÈC,(@Ù,"#“jZã;šiú0óî*¾Ï÷Üƒ;qPðc•€É*HçÍ›0·(k”Í¬µÓ²žÍA¦ä.ezÛcýÅë­7TB=íÖ$©G¶xhŸô•êÕÛØ¨˜è½˜sÝU_Ô%þñ[Q#@Sð>0,ËÃZKP¿KèP˜EèOŒ®[ëQ¢²ep0àB–ÝÒ¡KTü¡†—/s@ ©Ö÷|§Ã5ü×›@ê‚•³1–ï?Ü‰5˜îIÐÄZÌGGÉ™ÏŠy¢C•[2£z6k“C¬žoï½º+ƒ>sX”k@vÖÇÂå£ÑàQ;¯*ðµ³«úê¹l¦PýmšÀ<-æ9˜ÀViÙuƒ² X6š;$!Â¹<i®³¯Ä¨sq»¼ Ç)ù|b‚+Dæ°÷iË`ÅATN×ŸìQ”Sø¸c²[üoÇ±ÅqÜñBMK¡skj›™å*šéŽgÂX¤WÕáJÃ8bøÙ³%hhòì9¹É‚É\ùøÇw_|¶vÕÉÙ®ðL•â&ÒÇ})“ý(K|6ôñò}„º/©à¬Ÿ¨Q_•Qöéfv{¸žÝ!ÿ»;›ãí­¡jÃ1ú>47›SòAo+ÑÌA…FX…6
?ñ*tÌæàÄí½žßÄø#`S’±Ñ³úç‚Ãþªš4XèWyxYøüè¦Xùíà.sp_|–~}u‡×ÏÓ@òë—ÙÚR„8ªáZ ãüàÉ¯mNh…,ƒHüá«wí$´Kûþ°è¿;©.üûD¿(ú¹µ‚?}¼g×‰_%OŸÕMVp	-E#>±•à’ÿYÛì„IÂ¤Š£ŠR©»'úçleõSSÌŽ[rŸ¾Á™QÝ%ËØ=š¡úÎçóº<Y”µ{,›OûÖŒ²ŸSd}ë?>þÀ.Ê!—…Û©Ìì»y‡ ¢9ÁÂë5žÓ»ä^á
cnTV²y÷ß Ÿ—€ëSjÈ ÔdÁ>.?Wj‹u]©ÿâ8¨*J…@Ò!Eh	}äþ6ÈgÛ¡Ê&ÅlNí!Þ—xSÉš–‡rÅ›ÄÑ/§ùÈ"¾Xn&«¾t4p ä&´\¨˜m“qE’	x·ˆ*Œ§<»qRX¦‰,YÀyÖ³¢Ÿ=,Š`²S1XPoÇ#$Ecq‘xÏ){Xa¤}ž
•zrèÓJI†Ü0p/+ÚÞ%'ò^éæ±ÒlYIL'ÍfÈÂÅƒš—Î5±‹ãø\ãÖvé–ºAå€k3õ¸”(iÁmEÑ9Ï„›¤ün7µ7iûÚ0€4+éœª*ž@©¤Òš‰‡V£_¥œó‰	D7G"Æ=b
NLïqšçË	ÆY ÄdòŠf$H-éÊæ]•žÈë.ýÌirŸ¹ƒ:‘t\OéùE‰¦9§}ŸÞl1q•ZÀ{{ä2µ×0˜˜Åô²:‰`¥1hAÑÀËL=š
IP	öÁJÈ@IÛfÉídÖ†äº×÷æÝ­µÞ¼ú.÷É¢›Øg$›ñ)KÇetÍÇÕ¢'ü¶À3/á7ÝÓ;m2ju5Š¸ÊŠñ§Í4%@$”X¼ÍàF3à·ë­‡Ö	‹ùžQ3é|ÉnàšEs”äì÷gY~
GŸ<BHËÁÍ‰ýT*Ÿ px;ù*þ…‘Ü¼„{!ëE–Œ4bºŠ+‰r°9¼íWŠ´¶ÀŠ
ªnÈø#Ëž@pß,7¢M‰ÊUS5*q
æ¥Â¯,µU#Î¶	GÛ(s;ÝI¢ö¤þ$3!-›×kQL=Jx÷Œðˆ‘5;ÎÞÞ[Ì«ñÏ%˜²ÚÏ>»‰=@÷Éøç×o#¶„jcƒðí_Â@@V+Ìc@z
¥	¥–1YSÉòø*€S'U6#D¥çû™½RËHêZÐÈ_½
óÖOÁÒ?~Û0°üoHö@²ü5B³^»\(¿^,k,ÕÆ³-M)Wï™ê—½<¢™Ãºùüçÿífß½x2ó1ú!WT‹×ë5tJs—1ü™’@9ŠýS2'çõ-™9Üz`9È#U³?dÛþô¹P8a€A7¡5S€ÓOv˜½6Äßt¥"²h¨øšÿM£ ˆ‘à<)TXlß´í³Uù€µˆ¯P°1–cFÃNãƒDžÙ,É®¾$êl›ÙF¾0s_ä#ë}t&þiIÒƒv»AÒ_æ£÷$K<•~ïŒeIFK÷Ç÷†HÞöóƒMaß|‡ePïÄþP~LÀõúw¿;]LØdÚ~©ß—Å9å· Ðp!ôJ
g¦ŒËûU¡~U“WU>›³?¿)&d5DÆgxFsÔ×î§-ê¾$‹õ÷Ã|>»7Â““j~¿šœ–gôãIù—EñÝ„†øUµ¦uuZŽŠõß]ïgb”ûbb¯)R ?Þ§"ö„ž<Ù,Êê^?¯Ø;>•ý¬;ÞÏf4{øz6'î÷‹ÙŒÆ.êºª;m½¯ÊÁWÈÄ÷iìSŽ3.£[ü!õû#6ºVä%¸ˆ?gsãÔ—|Ñ«I¡Æ<&cÍÏäÚ'XŽ”³â ú>úJ_jÒÀ×âï¯œe?­°¾ùâ›Ã"ðÄüx…§ûdðªgu>þ–@îiÙÏ)¤fùìrÒçdlÎ?ù#MÿŒz¾òÐ÷Ö,à¡F,À‡ù–ÛÄ·`™‡¡üïê€šuOT&“‡wÔUY¯–OË¿+.ñ:÷Ø;­èI¹5C[„ßêª³nG¬
,8p„7›Ã¤íEM³u¨ÎzIVô)±³<¯/H+¨‡´»&=²~ã®l}_³Pà!Ÿçå<;-æýa·³I¦½I±Avúâ²³®Ñæq1Vd“:ÏŸ½|¥©H†ätPÝ7È8ä•G¿"4 C*åSš‰ aóÇY5Ñ¸VB¢”ýìO/Ÿ}ÛcÛ^ž^w?¹ÄúõÈ72ÏØÈè.6È<7SD†Æ6Z¯u­ÿ`C2ë³Õ6ohãbðJp‡\O‹×òfwíî	Ø'¶DìQÖKm 5ù®z`x‹×íqT¥ûÕ¤˜Kg!;rrz	w\Ñ3›IðGÏXJÑm†‚þw“w“ê|’1”¨¾•#½ÎŠ1w(pÞÚÝÁ æå6Ùú^¿U-1àÏ õ‡½Öµ{ÄØü:éøú„òCžÙ‰„i2+ñš:O×4àÛƒ×³‚ êá«¢¯øRþþ‘àä—s‚ç»v:Eš¨ùµ2aµ´z½ÔG]ðÏgF#†qµV¾ÓŸ¤63+Î%MÝ®Çzèm¬2£ìEà„JjíÝwŸ¿á¦I)c<¼×{$~EF¥ê—³§¤ÍÑ³i1Ygbäoc—TÂ^‘³qœ(¯k9cø#cí&ÔRÊ±#®ÿŽ»\{Õ…O¾öæCd–¯ßu_¿ÁZº¯‚iiÏñUã¶¹™=%[• ÿÐ` U4ÈêHFc0­7öÓ|ªÂ;èíé/Œéš²fù:“'G"‡›	pWï‹ìU~²ŸJÒßûðú š|ŸÏ‚šV5}H®ÎjL9­E*ÑÜ¿ŒÞ;Ð\Nù;Õ ýIšbÏ+hœŸÑY‡¯gÃÁ{ÂA?GÆ9íó{ùÓì´Ÿ×Ú6´BšdtHg¼˜:½3Vÿi1®ºº!7‡¼W×ùeï”°zÍ“+kö²˜wIff×6è)<±&ìã¾®Í'k>Cì0'Vk½™>£A×ëÙkhïÜv|A'Ì-r7=øž1°r*0G²7ˆÍ„s3¼)ßße×–Ì]ÆQ_¢NAÉžTçd,P‹Z ð4½yE_0
“î–³ûÂ”øeQ¿§µ=igL¾ÒØ¸crešLFv®l+;Ó‘»‚ûHµýTØ¹™_:7?Nãø¿|ÄÚâš1šôUßÏ´E»Nñm'?<=%{ilð&&’UtÁ…Ö·=q¾d&±ë_¹mh(ÐnÆAÄÂäG°ÅûR­ñU±‰ŠFÔ!¬S‘–<FgXŒ¦°¸ä¶O6e\B.uRûJ ¡ù%=ä>‰ß%µÛ¾T`ïxo..BÙ–~§á¯é`3£wòÍ@h ¨U-uý,û<ÛÞÚÚ¢Ñ}ÏçkzïƒòôôU)âo±¯×z0]ò@Ñ†>>õÜn‚¯ÉÓ|>ìŽ*ÂóÉ†7³.ŒŠîŽüÏÎí5þeê·ÈàÅ£¯,nPßŠ{£¢ž'î¬c
En°"î<+zÕ‘;4Î˜#›Ÿ4Ñ€ºlÒ¯Ž³í=uÂxƒŠßï‹h;v8ûùŸ'Ch|:4¡~Qi6ùöEAà|B³OáÕ¨)ùÜXµ§'gÌ¦®®fÔÈ‡+éO)«¤¿îv%™ô}Þ¢úF·Ù¼Öo¾úÔ¿H˜ù=r	µ3†ãóÆ¾S­jW#³e/´ùòrÆüwdÎü«¤Yï&Ìšïaê^#3¾,F£ê™qºJmÂì·œ/û™.ûÈ;[kZjRÔìTÖÑ§¤&d~#¦"_ùPLE$²Öí“·ûd}Î,„'"ÿâZ!‚ïÿHQ<Åðƒ²&D$cø€[¤gÀ¤eóaÁ¹V‰qX¦ÆïèGì©Äf¥’e¯)[2]:Ä$	¬ÓU¿;€ÿVøþ³'OÞõøÙ·/{Ÿ>ñðå½—?<¿÷âþÃÇ/î½•Óš.û1Ù¾ê]m2·´¿7ß§ZE™TO»ˆÀƒËZ»>Gò9ä–"Öê²¿OHç¥‚{YîúY^Éê5hÜñ…¤¿çëŸÉ
SþÏÐÃ…ˆE=¦q;…8 c(Ù§cö×‹Y?—Þ:"s‹YIñuÐºÆÑ1:™þˆ2¯V5§k6‚ÑdŒ—erhÛ¥S4K—VÃisJ²yClaµ¯OÇËÚiAVõ¦yÉ‹š›%×ýWÚü´j4™YÅéÆmÈžI>Yºufä3¤ºÓñ£Á{aR6‚ÍEgláµü˜‡,µî$$Êpž:ˆ˜ÅÎZŸªå3_˜ëd¾#Sa×FŠ%|ŸÅâ!èïA>žTäº²ÛÌŠOlÈÉ–'l™N›­Ñ»i	¶% D$_»º~ø†ý}1)!—Ÿõµî&Ã«ºg £Z`×{´›‡cÉ¢µíÇà	ý=ÝÏ'ýb´Ì|DN¾Í+–Þ±}7<‡kÄÓ=Ïu§tâ¦Ôµý¼fX½Uã´r¨õ¯ëjòSÑ¶yVÛjß"ÿ/ó÷…"ûÅ~ö¢ÈûóÈ œüñÕÓ'ô×¨€£y¤‘«<åá£ÅiNp…I›@¿ÌI\ÁñŸ]2ÔE]“Z¯¨~ÅE¾,T5"ŒQ!ŒõÖ"Û£#ÄThÁ‘õð€½zwÇú/ê—¢Çfõ§ùå¨Ê”E ÕÀr"‹ùÂm²ÛgM¼%ø®LÉ)ŸiVÑ^a5AÛvJ•¯f5ñ<Tç…¿Wä¬¥bœ—N]úýš§Œt*ðçX“¼$\”]ƒ=Å¾ïó|¼æ÷ì)öý¨œ¼{šOgvñíC?TNWúK¬¶¢ÿvUõí•°cNgäÚ‡É-[Ýè/}kò2è¢ÐX-—¿R!š¥>BÞŠäÜ÷ŸÒÌ<šûæOÙo”ï¿æÇ•	©<ÂA~Š™äˆñ7°o(ð €cÀ,‚×0@9{°`ÊåB°ò32B‰—8?/¯*87NQ–òZˆÎ3Ú‚LÇœÏq¯4QÌYr…ùÓÏÍŠ0ÐÌ 	)7Qa/r«g4ò”s³±íŒ…ù¶:<Ñxkµcüˆš´½s”#5¥®Q‡ýz’»—Wù‚\þï±ÖáZ«µÂ×Q]‹­3ÐW/öP{”eoŸ‘èà+r|Ë¬„t‡4ü$¸ÌåråÏÕ¢ÐÍU»þÙ•€×,"k¡†#ßÏi9(ô ‘AÑ{kZ!H[%ý±a óÒiìïoYjKM3#;P?gfgäÄ}ð)›ö–JÌáv)­Ã`…¸N_aTƒ¯yû4¸½ì^ÂJœY‹q°yrÔzA|fX]lbë|Öi?ÈõM´Rgß€“‹ì>‚G0“üŒû¡ê;cÀŒs4ÂÂÔw€”^<{òðeï›‡/_={ñÃýgO¾¸ÿøÞõ—›î8ZÙ{<@b×5æU§r:6 0¸¡ïROŽdÅïON+“ãµ¾¤Þž|êT-9ÒhÿsrnlÖq{ÀR»'÷ó{ä®FeýßV ¥è~CH¹Ë­}þçÉŸ'ÏF?ÿu=ûúds¥ý}~ëÏ“ïÍš@UÕäG{ï³+¾
×ÐËs²ô§ùûª^×0‡r_°ÁÜ{òµípO½íÙ¡ï ï˜·Ø2r˜M£ÇyÉ}S¸†ôÝ<èQíE0ÆÂûÆ›ˆã©Õìe i™[„",soÕßŽš.Bæ³’èüVGkT4â4ÁÆ›À˜ô«7kH`JÁ:õ\sØ,“Ô€8ÝŒXó¿ò¿ò¿"ËoüÊoüÊoüÊoüÊ¯•_Q"œ_‘âNÝ:Z³ˆæ/MógaoËáÕ¿3-Å£’TkÁ‡¦cù^a+ÒkMWü@zÿâñ«‡êeDxÃ¿s„ÚÆœðøFÇ•þlÜÀZÓ¯(³
ËŠë¼œªóuÐ©ÇÝÎƒbVüHàê¢?‚ ÉÀŽ‰¾;kjyt— j€ÝÔCÛWgÿëb\½/yOß}Ôr"ÅïéÜàiæ ‘ÝXKò•×é]¶Ù	¹&¾­.¿ÓŒdLÓt¶;µµ[_/Fïì³öÉµa7b›>aZòKkåìÞJS°[ßsÂÆ”…3%pMÎ8<bÏ	óÙ•w$×2ý'“·Ó •ùìXrÃÌÒCÄ”AÐ€ºí“¥ØR»®± å ¨¿;“B¬$3cHøƒ‹„ ¼ýìJÿX[ºž?ÿØƒ Ú&×`Go½Bµf©Û V'üüÎ@OMÀ~L@0Ï w1¢ØÕôáu;8Ì\3ö‚¾z€(£©)¾°3’nLÍ"Ô4îûüOì©2	ïzM¿Ñ–â£WRI	Ý¢xÛ¯Æ:æÇš"ÒÀC¦‚„o©âQ>§JÅŸÿ©âï¸ŽQ¼þšëáÓ&šÍÞçzD:kf–b¼©iñ,#³©¤çy1Ï1Ûå6u®a”ñYçI9y—Q%ûVè+±Ï˜ÊN}G[Jå[&ôw¯aªÇÍZÒ7Z)ý:û1s««rÒê9q¦´¾ª^ô‹QWÁîz&ˆÎìé$ÎÎ†{Åif,Øƒ2\x¡*Ëi˜Wè‘à÷ì¸÷zëÂß·à¡Bëì9íøQ]Ùø©¸$ñc1ÃiÄD²Òäûû\…»uuN‘É:ˆhÞ—3å
íD­î'çWPãíg'?|EkA;¾yø>Ã¯ê§ñ€¹!óÆÑ¬ÑÝwtDï&iÛ|jÊ_,~VWöI®†¤Ó×â÷Ê#@2„BºZ›ßøü[Y[#Zºmêƒæ€+â°R=™ußç#„WSC¿E>ã²l‘¾r¶kÄ=…H%î#{	‰Öù÷ÌHZ‡©þ4“±@ÕÂP°’C9²||4©yFôchUCƒ”&_Úå3»¢Ñ•ÖÊ$×Ñ–\èß«~¨MR~¨OŸ®ùdÿBàtByƒÃL!F~%ÿëdö©fÁm“^RÐ3—Š!ûDBãv/…å Ã~Â+Dàõ2è“ûF_,D`ÒúÈŠ Ö‡‹óeŠš†±â›ÑXÇØù¸²—Û?~C†n˜Ðb1z<Æ¾ 3ã+Í&MïJVéùùò0$s÷˜“|ïð~#a(÷7£¾AÚÁX’% ac±uµ6-ö”Þl:*Éµ`½ãÆ÷b‚ÝºàƒC¾²üf-˜ ¿IY™TžY)ràP@· =zÚ3a´(<èXÌe¦ih#pw€F¸¥Z ÃƒQ	¥¼iã¶°ž–’ê'{ì¹Úuó©²¨L„¸»@šÍ%Ú[]Ð gÝÍ??Ø<[‡…À{~Ñt%ð”6"ÓŒ³Õà¸©'ºÉôFo­Üô3°ÍÊ4efòG[v¢òoxÑ`k…1):	v#„¶¥uiäùýzŽ3exhJ7MYuõ„£Ñä•°LXSºµn²tv†­«ñ€}³æ*š²ÄÚ_(KnÝVƒ®•aù*Èí±ÆeþT–Â)«¥îçë™nLœ¼Cšnrì¢Ï;†ºé0áÚßZ/¨x(…!r·½‹ql8s#­ì­ñ¹¤Øf’Å‚tõ¤®idœb6¦|ÞÌÞ•Ó©ý‚³YÒ«$Â]hœ„vÛ&§´¦÷mz#XÃø~0Õµrl){béÖ1¡·°þ…éOºÑb$½a>ëjïì«Ö-Ú}ã_Ê¢„öf_æÅpH›¾á(Ò¸¨
*›`}ù;÷/ÄÒäígW¢'MúË0Ýñ9ßdŽâ­|p•gÁg½·6¯53’–*Þu"øÕ„]†‹/ÌëæÍd'tÒsìhbÓ3	íaeízÓ]_XM='sÇÉizðõ¢JÜÉf„"^íÜ¾6#:q)Ý¾Ã'òñŽ×S‹Ùû%’WUˆ$3)’4ÛßîX=NÝ<®2"°ˆ¤j÷ÂÓ ³žB?u	šÀïþçIYÙNË`ÿl™'*^ÏÈl4ÝØÖã¦Žk'1ÁÕÊTh„1‘Šìè¦‘ìS;q™§Iº¤â(	éÕùJ¤„bÍaw66ScÕ„#¿Š²ÍbÉFr‡9Óž—óÙ6R
+÷É@gf—NT×'ùeµ˜I`Ø1Ú¾{mìÅCÅ¶Ù0bêƒoíõSØ°'4ŽWdÃ äX£­rñ ²s‘” Aã3ÝÎc^/l:o¢6s3¡!—/6îÜHÜe$±œBuÚ ­~>ZÌÅÙr×Ÿ«§†E|¡Üøõè¹AäUd39›;0k‰¶Ý%r³‰ÓÔ|7#¦ Ò?æœÀÏâtƒÆÌ>uy¿_Lç‡ÞÅhv±žÁ?ö'*š»®ôcšÎ°
+Å®Å>8‘õƒG×Ã‡€M÷·ÜÆj±Ø#; ìÕù$e_Ù±}uÛ@=frHðö*?É^ž—pû«µn(o!Âã8ñçuVÕÙ_:e—LÉŒ]NÑÀ×Õ™žäDÑ”Í6Žíý“ Æ$2<¤¢X¹â@®´Aw¶¶Â´j_~OGA8GÙÖd9Òõ…Eº´eÐ÷úÿýÏÿö_½b!º¬m¼
«Ù	óf·&H‚ÛF÷qïÚÿ¿ÿñoÿ2)fnôá6N‹hú±nŸ6Äyþ÷ü?Ò?"g6Gä JkæYÖ¥¡óÙš‹¨¯ì¨1 êpÒ9‰´¸k‹¦'ša"Dl«g§ñ•$ÈfiíÑ˜ÄÓ½‘^Â†Ö+‘@¹2tîÆåÌ<—±ç‰Ñt€G@ÜÎ»æcKÎCºëæX[Žør2(Ï*òê·[k5`o]•,„…BLÊ >Ëžé“LáWdÇ¯ÒEé3[1
°fÂ’„–_m¬´êÍE›uñÖK]Œà46óÎÖÏÏç'³j´ £ƒ¬í(çÕtc{s'c[JGrI˜w`$ég°ƒ× hÝnSÏÕ¡ÅêŒì0U_C’(ájQ-¨µj¯×³[à©¡Tü<û‚à$…2‚ïÛÉñ×‘HiDö-›Ö‚³×ûò)÷Üþb¶Š8íê‡úÑÉFå¤Ø˜€âgþƒW;ÝY$;ðv6ì«Ÿ{h²`,UZ0I'¨Ðõv c²/­%ÓõÙ#À¥0+uW†Ä!EBèfQB¡E«ƒOê­ÝV°"È…6ü#Éž*šd›Ï—Ç<¢ñþ6Ù»h%ÍâðHý\Ý0$<"¿$'>iOx$ÿL®lZöÑŸ" ¯BL(ÜD2a}‡ÁÒØ¯äT q$Ó’±ãS>#yú¡Áà¿ÉUXÀ#úOr%Øïˆýûéƒ¯*õý+`#Ækèip~­@{e¦]cˆoÖOÑMë¹¸–«{íÉ_¸ÀúFã¦‡ütàVÙg­†±â'ò"V­Ol«8?råd¨a;KƒÜ™„Ü™rÑ|â¿NÈ}ôàû•€¬Œ
«2ÿÖ¯HíÜCÍ°+˜û+Kù°,/ða 6d]²&C¡\áòSL’fX÷Ðóz;i?ï²k) øänº"Yéj$JA9œ}7›E9º^©ÈA[@¡;&4ˆØ&mñÅ3ãbP.\;%(,ó»£WPÃ™&u¬¬°F¹Àt~3,(W24»úH¯êE²¿ßÀp•¿v3Zvda|åC¨Y/ˆ4â‘“¸K;ŒpÜv¤¼	Ùß%ÌµVl}í°u>Ðã¢ÁfÆË“é^‰)P4éZ,x†€nÅåªÕ7Ž£ñÎ=©Ïorie‚œ«‹o>üÒzµHÝÚŠüRM`VH¶~]ä«Š’©È²ò¾¿´Z¶‹¥÷7šw6%n–Æc·1­“Js²\›{`	m³Rè‡ÿ@ûË!ê/Ø«e¼˜©õó !ˆ$—ƒzÚLÅb×Ø¨|ì™å°Šnh·y÷f–]™î5_v*Õ½™ugÃºYw:€›YxžGð^å3l±ôL4~3kÏz6„Ä÷Îí	X^B<3šÉºšdÏóI1ÒÌ/é7þhf¸­f
N'§MkL¸˜÷éßuuä(:ÞÍòI9 ?ÍÅã‰Krðã;6ÓF®¶/µˆiû$¯&gG`kd+ù7¬ãî5·V !%Vëë¹"6P~J	s	YÐó|1¯är×+óxu@ñE±âšFááÒ0¨×å!\”·ÝÛË$câ“äYÐ{[ã•¹ì…ß‘Äv‚³q^ƒå	U¥ çó•YÑÊÚ±ãKy[[Jyå«FŸÈeÒŒù|+¦yÉ˜¶&~	é<meÈÕm6²¬mL–E?ãFè,þ¡~(ý¨ß’ì¹vmcÎÿ¼ [ÁcfóêìlTX¸3îga4[F‰b+²%g¬_qØßÕ¤rŽôÓ
â³NUv…ä+¼£Ñ«þ ÀDw¨Gñ¸HØÿaèø%AaNÓ
ñùDŒtÈ'9ÎKZ[uqo4è"~÷
r~.»]Ò†Iˆ„*È¨SVöQÐñ€Ÿ¶97Ê²è1t­“Y¡Xé|²[Ölõ£#÷ŒuÞ^n¸÷ê:¿ìAÖâ®p²™H-Ïù›fÃuž]UOÆ©”{ÍºUänÆ½o  J"äŒrŽËùÖø”ƒ¤ ñüqðw#	Q³_‹ÀeqÀ’Öî
ùVb`§BÖýE9ko]]ê`šÀÊ¢¾¹daî4|Çß1µIìÊo+’ì­2Œò"°ƒ†S¶ÈSëŠzÎcªŸs)ó¤ÿtnM+)¯½<è9¡šG6ŒÕˆzî²¢Ë—îP‹Xu!ÜÅ¤LæUÀfÿ9÷¢¹c Žog¤ƒi&m q4iâùÀqý Â«Ôu›;¨‡0+Ê^J"<AybÏý‘° ÷„‰vXôßT4¼ÈØcm²$ªØGAš{ƒ÷3ú¾ýDë¹Àf¸Õ¶*Œž‹¡ù—‡~Q¯Ô&àC†¢Yxù QX
š½'DÁ*(y·ˆ¨HÌÈ²þXAÉ‘(A*žÄëÈ,Át'ªÈ:þêMxFÞwÁE÷l`¸Ã"¼˜uü9å!ùŸåÎìé½nàtQ®Ü@ÎŽhDkÉ‡gÍ	P)ud7àÌîfuy6¤®Ž9Ly€àjÈLgO-{±'<Ñ<3wÂ¾ EãºÈx&Ž9Ùƒ<ƒ*XDm-v¤ÙPP¿ì/ÔPØƒÝØPÞ×… ™-²äò';^² %P|ØJ_‘.ï×K?¼„!]³7›çõtÜ{¢ÖÞp×/suƒó0¢¿asàwLî€°U=—œ»®ù*ú©Äòw£]·É)#¸êÝ†ŸÌÒ6¼2#UbÒ#UîÒzqc~àœ¥S+Ä.¬Å°•²ÝÎl	à¶f4• áS%´d`!vüN‡Ö×^ öÍMî¬žaFÜ%š/2•Þvä…{7³Â¦ôoù%à¶0M÷œ\z9¯ó)½‚H•`/×¼Ok$f)$üŒÇr}½¾Ö	t;C’KPŠ©9saBk`(„–¾‹ßéÂ¤.0(O:ðÐ°„§ºÎX|1íVI:œ¥Fœu°´±Xd(]ÏX¾†¼wË-§µ´I`J–;6ƒ›G EÉ±›Åœ{ˆTûT»zó*}K€‹§Í¡"zç€F¢ý„o(ÌMÂBØ>¶ºR’§”J$9Qü'1à—`H+9k™’±õNÎ"×÷ÐbG¨Þ}êÇÁ‰ÞÖµ¾(j eŸÐw`[#DQ(Ç"3ÓZ¦~×Y÷J—">È/©ríz€ÉˆµÞ"€´¨…	4‚P2€P›Æ1Ï¸-¾$æQ”$Ãq Z u}Õ ]u”B^^¯ÃœÀÐDâúˆ«b–ÿÿë?þí—Ø+wÍFEî f½(bh‘î8F™açqã
cØIfš~“Uî„´C¶’ñ‘%$)nX€Ã¼ö]ïÝM†œ»ÔF¥1ƒ]Æ¤2.F1ZÆFZúº.‹S*ñR×	Üx"—$!U~@Ù±,F!	pÚïy€ù/5ÔÞŠï2æi#dÙ b–p“«ÜwÒÖsšMz».ˆéÕ)U/Ò†»z2ÂµëøvE7Löã2“_¦yuF×„SÈ›<âX¢ÆÈ±¿!¨Z-\	ÈŠ~–)ØKøÖYE†VÎ/	©I›½zcô)|!/´k;%ÑþNFÒÙ¸ñÓÁ’—†Èlê|¢³	!ç »»:X>øŽæF[?rÜh¥zbÄØ<¯Ü\£«ã4…á[qGVkÕ“aNèÓ¸‰£ûdj5aûFûÈœKœ¸yXzš—£U’Q»^Lú¤n‡ÁM'•
Y`ÒÓçeV6mžçjÙ™_iÉ?Î–)¦*J$•Ë“É¦¾Ú‹ÈÐDim•9ÿJËSýÛV{†Ûz«E2pm•©­NŽ!)+2~¼FÐãÇ‘-N{”)üG„u}G#ú3Úlšœ=5â‹šqoz{“y» ÙD—Üø4c6Î/6†;·3PÀžŽªsrC§><I ð°3“ÒxÆiT’Ï
S3)Û\|lvþäLÄÞÛÚ¼ãD¶³üøÞf©†¨b@HiÇ@ø!§+zN h (~µIZÒæ•tÁJj)íþäí±[‰ÜHÚáŸöI“2íd?0@¢²õ~+Ô¤ $1Á¤HÊÅoaùÉ°.NÃ;Ã\v~8å“wa8¬‹ÑagRUÓbR@¬cÒUQ×~3FV’ä)NFNËŽ¤eé.ÑL<B‚Xä Ì]g{5Ú7£ê¤ˆyŽißÓÓÛ”B£óöÌ˜LÁ¹Rø¢-þ* ,õ*^¾lF&tÉìíëá™Ã$ÑìªaŒnÕMYØÄ)b`–f\f–U“o HÈe¾ãLMì­ê»ŒgDò/ŽW›çËž¾u\¾ï“žòÚŸ5ÉøœÂ7dëìP†°™›wÿ=/;Ìkkï\÷×l?îèoîxôxÓHËrÉÐ²4\ê÷å\çÑy~2*€bFÁÅ©d e-AìØ}î`>,òº`óÚ˜škJñ¼ÞÞnyÿP$us×ÇÚQšæ˜·™ÏT2'”šÂÓj…ùUû’ÖŸúñÎÁæ|˜:­Î@—mjÑr›5 §º'ÍÜ';’7mCæ¬nVíIÕ‡Tä9?ÅéUáòÆõèÔq£s˜à_Làç5Ì›h>˜ŸTƒK½S²[ü6.3þ‡m“†Ë¿ù7BY•£j«Ÿ#”€¯#[¡Ú{¢.ttÉ¾:N€
ˆˆãz#D¡h+K¸I-A$à’•zÌ(
Ê<Ie¬Ú>Cî\èË(“…ÝYIu>d¥‘"+MYiæŽÈJ+§DQÕçš˜T=ËPÆ˜µ$+?GþQÌƒ$Ñç‘•vž¬´òd%}UÂ¾¬Ä€.â!óA¼#Y‰8eÄäM!éPSQ°×L˜%ÒœHž çfif×ü&ÙÜ“#ÞVVfJ¯3&š›¥Í("dlæ+;N*72‡Ò)'4Œ;,rC(ç iÜælû]¯cì”Æ)²LÉÓÜhXií+Êq3§•¤6—óŠ5G–ìû‘Øªá$»gúÈj3Nò‘%j˜”zrR|gEYúEÍxçè••¿u8µæœ@ÓoÃ£ÛÎÃñÈX…¯æªýÝƒkèó&†·"¿·Ðª5ð}£Z…ÿ”UùÀAiå'&´_8(¤•Äº´÷‹ƒ’ÌÖ¤ûÇÑfãÌž[FîNeoÒ|ÝRçcr’ýÞh‡7ÄÜ¤ùÀý   ÿÿì}ÛnG–à{E¸àmm2y“d¹LJC‘”›3’¨!)w4‚”¬J’)WU–3³(±9üŒÅ‹],X`ÐôSc_æ•²_°Ÿ°çÄ-#2ã–Å‹([	[¬ŒŒ{œ8çÄ‰s	·FŸm¬áð¹BÌ,ãB€Âv. ¯ÇºÎ[ÃµjZ_—m5PßúWÃª„Þ„KÆ—ñ [ÈoU iP¯\ƒ…IÓ²¨í!Ô§¡}àb¾Òb’[¢zÝœç:&ûÌy*¥å Zqs³WSÑ¾âô9T'kZUAØ¯Òì¾6äæ]™;ƒÛfu—•Âá«¥C°ÅbÃæáÞl4{‘¶¦!–5kjcûíb9-þ9íxn—ò›èÚ„oG8}ø·ãµ™!˜+«Ír¨®uÍháÝ×ª~ý<i85–¿í¶+Þ¼ýmŠÿ~›oÏ*¸žjã|êªZBìÑ¦‰{¬w«‡ÝÊ`ìnÕ×w¦àÛQurÜE¼¨®® +à0ìÖÍß6›K¸YœÂµ»ôšTTïú]Ã,¹æî³Yïönâl“~Í®âf<,X„R Ä'ßì¨Þ¬£ƒÙ×¼až¡
TÜih¾-RÕ7=Ù§¢ú&>(ñw¿!ÏãILÐHÄf%?¥ÉÅQt=Hï
Ôcôb%«[JÃ"ZÄº¦.¤Ž¯cá3ý£¦Çý,túGMWò;ƒu\š•7ôjÞlÌ.ÿ1§Ô°gÖÉÒúy=EÍ­œ×ÕCc•Ç2ÏR‚S.UTŠ‹9²—L²¼ôÍ¸”÷õy·8eÑ½ÒSßàØbA~œ¦ƒ„lÆù ðD'	‰wWçS­1î>“ŒñžL’š‡ÍÉ+•›¦šïàïtýq‡—UÍ÷ÖÊÒC'ŸÜsè&ÈHxN…»gÉè0TŒÎö8 ›ùÖOîZ7†T¬Í
^„ÛÝ@ˆCGL“Ë¿2°Õ*¤4¦{Å°}4ƒ×º÷Ë41)Ñõ™&wj4dŽr6Rd0>’‘£¤³{„AL˜eÓÃµÆl®.šuø¬î—o @-. [€(ÝÍ›iÞ&·¨Uø:'¤2´†pºQ&cS Í U™– úTí=$dœ ?²Ë¿äiÌ 5†¤ã —–À@lÀ"1P½a %ZÞ•€“V-AS^ÉÞQÀ¤ýó%Gš—Û£$?¾ü1,åÕ!SNNK¸\¾_ÌlJFqZDÎÙýüòoeÚ9t¦%Þ%phŠr<Ä°0È¤ä¡³‘Pã*ž¥EI²#Â\õ¥I^qÖçæ0®ÇDÄÃ‚`h»Ù†7Æ4ÑU´››HwEoÔ=|(ÀóÐìxh•§/ÿž§tñ6¸ÛCA…ql¬[ÓÏ|H¸ZlC—²ˆ´Ö6ƒå¢gö^(@Û ·fé-'b£$,üšWÛÅváÉ˜Tø7wÕÉÎ€£Ïgë¹ð™ Ÿ´ê–þ&/s×ë(k@p|7,ã‹d|2Åâ0Fg²žÖÕˆ8]„›ÁEÈÁ$V¯÷Ùg¸E®Çt‹8Œ·È2ß".8ªIã*›±©¬b $ýOd¤ûtë§¹¶q?Ý¢uI<ÓÃi’ñT­$[I?ËáX™YÍœlÅå±h‘0{±v´¸s;,Æ
>Ÿ¢—µhúØ¬¾,â «Ñ¹6³-Ž¹lôÚPq€A>W3äªjÄgE­°ÔùÁîìŒfâ›-¼ª^']˜žä'ô¼Ï‚‚‰}ÅKÌ	Ã²ÎBÇ%½&Øß2Ù=|ß#Øª+36Àë
Ø$ãAá³âuÃX0&6P¯|C N–®XÌÞÔŒ!¾„=Ô‹øE—Wá¢_Û›ažKTPfÔ3Áþí—håÒíLÊ…'{G//`ã1 ÄÇeDˆO!!>ÊŽ³¼‚+£]_3´e£\ýÒÂþ†*)ùMm‚ô9ZZÛ4Ç¡j¤ƒ+‡x‡ÂgóÅËì…›ÝŽÎ{cðº%¯d@ª '«šßvÕå=:[š(¾“†ÇF.úæ1œ’šO+§ïÞÚ¨SxÉåŠ³HRÕII¸·º™=¯Ó¹ÐŒsêÐí4Bð™Awˆ¶î$ ½æ¡¢.—T·vA¹mjG•€* š ý6·>fÐ4˜–'da*n¡hÃ«E«òèê°Ãf°¾>ûÉˆÈ¡“SqX	U·3Ÿ&ÏMm¶v®Ùãw¨8ó¦59	ÕörŽJ8ôH;èî0lÏÛ\ªŽ^ÍVœŠWA›×ÀÕ@³Å0@ü…3tßû}ý±Ún;´AwÔ5'ƒãyT'Jª8ÓÏø9Ø(?’
³Œÿ®%S{Ž'îë0!&2Øê¼Sé>›˜0½â› ­LÅÂQhks1§Q-C“Ëm-Æ X¸ÍX[«±€O
¶kg=vuû±6lh »i°"£7Àh¬Ûù©ðsmiZEp,TŸVà:Äš&HŸ6:™øèz™AÏt7_MTä…2\®Á9¯™µ‡@AK‡Í*úð˜ÃKy–H\nønVLå5]QNR¦»{ŒÇéé:y	÷ãÜaR}*àWféZ°Îÿû_ÿûÊJÝ²A|Ât€iÎÏî¤c€›;³C]ý¸	mRŸæ–AMNÌµ@ÚÿýÿUTù›„3Ý¹ÁÍÃš1†¯nÒÝƒÍ·õ-¡ŸÄ]Ðí³ü³„· ?r³Û¶ð²­í[ðñÚ¸à3ƒK²xmoiŠÁc8C d›Ùh2LÊ ßVAünKóZ&Ì®ALo·íÀÇºÎ&«•‡ÅÎÃ¬–£)ÛéßµWÝ.áE†Ú]»h‡0­š8í2535I˜äÙQ:LGy6LX¦Á(“çqÎkf´M­IÚ¢ÒÔUX§ê¨³8©£ôxý\þÔM92h-8èýÛÆ íhÇ¹Ú2 ƒ1ôÆ€š3’ÊRæeœ÷“4kKíÛôÊæ6˜Š¬m0,ÿZH€–WkxžVµÖeóMŠxéú²°DþŒ‘×Ý9ž«¨ÉdÜÿ/˜áï Ï&x±’7ÕJÖF"ŠZ®Áf:†Y‹‡ëçç":S,Í“¯âáWôýý&‚ã´M+³,Ë,7$Ó²U>ÍS®a
¿¸ £V5‹·óúû¥Ó“7DwN©5âW±¾š2ªKóZWDe¼DÓa¾3¤õy¢;zøãx§tÐ=î]‘@Ä0©„´S]ÄÝGÌkÄ»¶ËŽâa‘1«]¾h¢†Í–JYÚífªd£Bæ9VÝIÕ‰›Iº†âfC÷ÒÁz‡ß†?…æÙÜŸŽp1Žz?>5ÞHÖ W˜Ýœ§VFe+F£2µ.«zri[œq‰(VÉ·Ç,½FÄžXÉk‹´ykçÜªÇ´¿¨§`çšÉQ<–?ÅÃ) ¾M{Ôfòä—iš'kC4Ž(b;Sô=,šù+ºá£?álZRç«®c‹…³sEøà€ê<7¨"4+¸Õ‹&@Ó“p’¯w––¢%öÿ"ü¿´àrÎú9Â†=Y“ßÓÂYùéÈÌb«¶Û0ëÑTvC-³³ö€­¦*<+ü6]U:¼S^*%ûw-àiŸúµl‚29¥CîtQKtHI¢(Z[d_ƒ‹+réGÕïÖÕhZÜàMêM·ï‘_>’?[W¢íi*_ek‹@?â!½GÛÌNY§¯aGjŽªg¦(šð_éŽäR­¡–:p„ÿ¶.Êœ¬?¢Z~’gã?'Gìï¯jT±žnÅª¼ÏL¨ükìÖöÇÙLù¶Ÿ'?ÅyòÞÑü¯›åúÜN€Z$±	8&òÈöãb*®¥Š³ÖYÁ4Èßë—£áµÁÉ6ú,½A >QíÀÀ„gžLnˆñx]ý
¡ pÀ=“—0Â~:‰‡¤+	Ü „‡Å³®¿(o%ÆYâ½@M"k8¢œY'šƒeí£ËïòR¿Ý4Ä2àÝ´î]Ôªú\`øsÑëâõîž¨l›ú4¾ü÷›<€¿É3SiáwùWSŸf½ŸP¿Ñ7¸ØÌ1õ¬KÍÝZYè«‹Þ¨¦ÇMÞÜÐf–´y’Žýë•<KÇ?“q†~sg½V	)X‘Ÿ±‰YJ””R”å¤è-.gYt<\AÉÅ(Š~]bªO0ÀYïÇÃ„úøéòoh[{“”WŸ¶wð¡´„ŸËüê ÇÍ,•Ý8q³ìœPCnR¿™’œúÕ]$O·~²ŒÍUuœìçî»‡b×ÒO·ìêŽÃéMãh¦øîV*¨¼ZÑÍáÂòÝôÕ”\å° Ú-n¬ÇUv™¥¸Z(ø}3õ‡^N9æ±ÛFøo¤Î‡iQ>Ír¶4¨ÊüxÝˆ¡0ƒÿhš.œ°7·:6dBÜvá±=÷ÙÀN™Ž²7è4*’ü4À<zóvätk¾ÏEeö
ÝnÆEÒõFºXüËdwà¾l"ƒø°È†SØè“Žíe6YX^\!ÔZ€n”3š¼Dê¡C·ßEÍŸšn{Í¬¾{ñ‘
„î+¾Æ‹€c+™2ƒì‰d»üV˜;)ÊÊ]Ñ„«ÐšÜõ*#¿5‚´b¢W•Í€pÒžþœOÇé/ÓDÔ¥8KuzD	qv«ø¬¡Ãui{rw©Ü2@qŠŠ¿7“d+E)ãåßN“´°ù‚rÙE»ülØÂ9 qÓA_}@êÚÁ>¥5 mØµŠž›kg6`ÄåjrÈºÆþ˜5Û±+:4º©œÜ¡¯ÜÜ/>ì.¬’ô>Ì>º2Í•Ö}¦`‚‘1A¨võ¹ñ~+5>è9Q2E|D>_‘lP’¿½ïêƒñc
qikëõýîyö (œþèÃL\ÜàOàÿ0Æ§UZ€ÏAd€4dèEªkF‡n†ÀÝçPÙDÁÜé}óB!s§š›øÈzsÃ uN9„Ñ:üxzb7Òy ˆ{—
ùGB±ž†¾ŒAyë{žjØ›|ïR½ÆÇ8AØŸ6‘N„ÀŠrEýö¨
Ö
6®e`3†	ˆ‘m$À4 Œ»0*Ž?²LÉW‰c-aþkÜ¼yõôÈë(Šê½	16!Nœ± ;â +â[B‘ø„Æ‚›5ÌS“«9ìpm–¸¬^+m±Þ5Ç.YÄãÙŒÀ[ÁRç¨–—¶ÎgmWP0ÓÊaÇ½FdM‹¹)¦1Ù	 ÀU•!­8?=Rçwz¤NtzÀÍ2‡ÇâÍ¸j©g°iÜ‡§Ó±EÌ\»SÃFîÇ^1w45:ûô¤Ê@VýRË*áimÑhù+³ _üîwGÓ1ºÍxˆ’œ½d4ÅPuý8Cê.ãûb}]ÄpÞL³yŠDØï7€i§u‘ £\˜šGÝó‡ÍK¦Ã2Y±çüÅ^N„nCf<Ó¬Ï“QÖUiËuÊ;™àä[$O‡Y\vyO£<¡Rång¾3O:Qgn5Ï—~P+]Òk½¶WAëX\$ÍpPí¡@c<eä	uø†<¿üO^3œøö˜•s"²×ßTíc§.æ•Y]x£Î
SÃckz†~æOÑÇüxŠ“•Ùá¡;õq9Œ^ÐOiiáµ|^RÕ¢<CKçNŸ×+¯‰øödïO¾˜‹X7°mºÔÌÉ¹Æf5œFWæÏ,†‡4þ|(1Ñq¶Æ0i¢ƒµ“•FöÕP3åÍË¿PgqðØë¦k‹'+JcMŒ£U¦ì$&søâÃ8ý˜1§)}ºÓ¨à&W›%%FÉQ—’ÃÅ=ª¬ÏE¤#qz^´¦!°ÇBj’“‡rEtÇî³ÜÇÖ]ÃÛM±Ñ¬B(qÓ¶oƒ÷Ù’.½*õDZ"dïk•1³>VvŸQslØD¦ÅeÃS&þÅxõÂ²‡\·èæ¥óæSù6fEÞÆ¬¶6Åly#£óê²4z°.$€ß­ìÙR»iÞ	fPr#Aåš€`‚"A,ý`Ô0šAï‹\öÛã»?‹F@i$8p0û‡vîßuWòà*‰A ðäFZí`·d¾ÓìJ,"mñ³Œ!ñT:Ë¢Ñm%°ÇƒúõŠW“{ÈË˜ æ¾ä:ÐÝ¡ç:ãÖ•ÜmtÍÁ,›Ì‡˜šÅ‡•kcXhÓÕÊÓË¿yÔËÿæÝí“0à3¾zùg4çåÚöBëú x¤ZÔÐú¬‰&•°æÈ¢AÁŽk!N´© ƒæ“yùWBCçQ¨£Ñ,1|#Þ?$–å4Ü%“AVDBCeEâXÆãþœ¦Ç¸Õ¨þO1­î˜‡‰Ž¬ î¥þ£t«÷LM&SûY;#R\ôXˆçî£p<Ëüj>ôw˜EÓ0";þÀÓÀòáÀÐcØ•øSÆE¡ù²ÑVã/˜Œ¤Ï«Ížf0üýá$.‹x2yžE|Ì*e>ÅþÈ¿$£	â7ú	èÆ†ý«-/|ñ>pw`ü…¹=Õ=ÑS€¥BÒÆ$Å·QRÆ| øsÿ¹øÅ®üðí—ãgé1IO‹Ët]h§»×oèØÆñð¬H‹ýþI2Jø£‹íKßJ†Iiø‚<ØIöagü>é—ÔÓÑüï.zDYÊÁ‹ô—ì…ü*‡ÃÄòj_SY#]íUH{U¦C–ÈW¾Gº£a'ûyR>†d1ícˆžÔÜIò<Ë;”ªŸféà'Püa5QxéQLµ=.ó3–
ÀÓ#?Æ%‰R¼d?ââgöEÂUlòŸ¼n
cP9ý»OJy°oêzx´§¿ôï¡=öKÿ\Î¡`·!SX6¼ö€Xõ®‰©ãÂ‘œpŒ¥ Âo“ã²²F÷1KÛO(¬&t¯™«ÇÝNÛhì†^Ý=˜@¾ApùÏüX€€,$¶'í)ïý]É$0§¨@ˆ]'qŠ¯›Ù y\ÿöKnI~•©lþ!ù0Ë ôÆ,ýBÌƒ¾É[Ì/¯^]}„ê(í#õÖ1bû¨í#£Ç‡‘	zº0	'=mLLüqÝI6À-ÿãöÝî/w÷:?tIù˜D$XŒî@)£´HÖ •¶ÆÑg<g?è0LÊ÷ŸWi§²tú›}ªlü3û-ögŸ— ¹™úLÇ^5LÛ#:(®‰{af
þ6
ú+\Ó3<Ò‹HïÑV|µ¥1Áæ‚MtÞ R#|òE•\hÒÝBŠç'>Ô$·té;'iQfyÚÏèZS´Mª¦Ò¤x‰¬\Zexlpv*Š<ê*•¨Âcé)TéŒ±Pï!ÀÇ\G¢È]¾
¼.8Ý¨Þ:—ï”¾É7×«êÓŸV¯ež¤¨Q(A@‡eù^¥Ò«áq\½³ãª8r^ô&_UqŸ¹ŒØO…Š 2ùZf¶m´W¶cØ¬ü¬þ˜à:=¬ÞpÈ‰Ãk9ËÚ2!g¸— æ/Ê‚-•š¢-×~6Lû)%˜Oqì¯ß<ª¼¶aF¹êF¥ÏÔº¹ÛF¥*†ZŸÒ«\vo¡$@ù 	;CYh,6}Ù€‰:¶œ¾ jŸñ@ .+Âdô‹‹dzXôóôŽ¥a'{$5/ÆîbD{ÉG3Þjûè¶°v·*8•“\¼Pp"B—òûË¤¨^bÙLücÅÀ×LÇ¶;€µ¹ûìÙöæÁÎî‹ýh÷ÙÎæÎÁÆæÆîÛ§»Ï~ÜÑnYé)EÄ°ÖÿãIq’•âvüAØºÿR»´gÕ ª/*sð<Ñ ësÍøÅ~Uºˆ(> 7v%5Š°\„¾|»sÕ‡‹¹9<5¡K‰¿ŠË“Á©8“>,/Ì9<#2ª’Ì‡ýŒ
ÈÙíÆóä°¡€ÀF'¹d†Wq™ËÈ¸Õ5S½È(rR„_§°2¬9…V«ñˆë®‹#P®ªm`RíZ¥\EÊž–°DD¿u;Ûø‡ð­J(` ^ô:ó„U0cÄÈÄ}Ó«¨@°;'îÊäÑîÈ˜Ò­¤ŸÞ.Sq6îsþ…ïÆ^\æ9;Ê2Ë¾—¬›8‡Èi¢qœãiy±›³MÕžUaAÉð*E§z1Tb×‹èr†¶rTQ„ª<Õy§(\¯xÎXJ5 þ&—é±H¡ºñ2¹×ìuïAÝxÑíH`—·t‚÷’#l%ëûÎ¼()}Ž?Ä)`žaj©ªW “-W…
b¾j/Ñ¾£W›c6tÌÚSç¦ÊÅ‡
ìX’ŸÂéö!´8šTØGÙ‹üø[©½ Æ¥1	ùú\Œ´àŸ€ÌÐž4xí×çl<Üí¹„@ Ä¼ÇÌqrÎÀ1î\DïªîÊS¶Hâsy\<ð„tkóäçÊŠ566ÜÙb«°M-WFŒ“f'q&ðÆ¬hcðÂœô«ÓÕ…²Q_§Èb1héåàhž¹aÖÒÜô=-^M8zžÒ˜× '«¨ah‰ª0E"jv±}
@´ö‡ƒçÏvð®c{˜ ö¨¡‘@…'ë•³|/G¯—¸~E˜¨£LÆÔïLÉfà#F‚½8uª%“³°KŽâÓŒ27ŠqAœÿ2MO3òrëisˆÖð…ä$j3¯2*ŽcÃ„½/ž¥‡8_t—¦#tƒßíÐ @U{"sôã0;Œ‡ÌòŸ“|—£ôm?ïCUï„á_0~_Dýa6O›Dýl´¿?.ÓÃ§#z_,~}.k†Y@uô«2ÄÞïjÈ)ÎóøìÉØ®\öžN½ò¡[Ã›P«Ì+Û„å¬4E8èžÓ=­ò‹¹hÂÎm¢28êQ2ÔÕ©p&0Ø¤‹ŸRH_þþ¬­c+Ñx:z‰2$HúöÛ¹?5oj¯°C˜¿›*¤—S›Ì¸¬rC6Ì~ ÙdßºB˜é€õ–—¨ðñlø“£ØøÁu.zŸ¥ãn‡t”êè¿…åý×ñÂÂyyù—c8+ªK/$üëøù¶jí[ÒùW¤(¹ê‹Ç±9í3í™”DEe§c*.aß¨“I}ÇÉY²¬^å·òC·öå1vº†èvc•ñâjšëò?(r'lö`Œ´˜­d£cÕ&F‚[™‚UE@LÔóÉîÁÛÍÝOw~„½”xÜ™3‹–jSÒS&@å©CˆœÂ*äìÖ#ˆ#ªM:Ø$iÌK¬b*ŒlôYg2thÆpÿªÓ–h™(Ó»&e‚¾à‚@îhÄd±ï4Š^’‡¾3`FmmøõP× wxº™'§pöãZùnÐˆã"øG$œ?Ž štŒeÄ1åHÍd³Õ©"¦Ùˆ¢ý@jêÀœ:KÊhªÌJ^1ÁšG­ŽFå€É7µkà÷1™¯UEÉÙåÆSJsµ!}%)nDï‹?¦åI·“Ž 4QÕÏEp“1ô¸ÈdçNNpé…->jCwµ•qÓ]Ø'(Ž/vFùwøÈzRFÈeúúñMQäùºx	Od½QŠh9çÂI•:°,6µGTrAx¦ÿl[@	_í=ë™²dcÎfuä¨ÚýÜüctÌÛ¤CÕ›#øY©Êh#À',cD"z¡¤šŸ½9×í°Zý¢øó?½ýãÎÖÁ †{KKöLØÞùñæ\HÛ?¤´ð¢§¿M™NŒ?Ês±—êªÝ¸¬²G<¿É:KÍ%Ç`6ãâ~³®v‘ûˆˆÌÜÌT×i°Úua/yó”Y4÷“5Ì»É§{‘ˆY²ŽKÍÐÕúSèˆÄ ³ÂóÈöÌ]âàX~DvŠ• ¢Á©q·³2h‚ N0M$Gƒ<þÀ6€Ë<†,‚ÿhçd4Ìå ¸EóaÕ•2»–£¿Éø0ØRôð¾¡Ž9º¼šFë‚‹‚tïsºIÝËN‡%Êïø•G£>íýÂ„
(ª¥h%g¢+†Ô4Þ Ê_Ðã†»j©#&kó¢pÕN"xòàž@åŒÿÓ<%]²xÆV¥çn’‰”T·§vø–øGl·L'ÀÉ•Óx˜þùÆ›ã‘¦Ç(ÖdÎÊ-êô\å,!8T  ,Úö[cÍÎÉ˜
§:°§ùýfý˜þä°ÅëâÆ)²¾íê]c/…Æ†d0%×UuŠëRˆ~ñW½ktDð‹Ý” ÜŒv±« )zÊ >Û=úc’üIè}<ˆŽh´/ü~QoŸ]ÆÉöÙ«¡}9=Œ«ãébgŒ’â~œUŽÒ‘!•e—OÓ±–·ž$3nôãA=§šVÓ0Ã`o†Q”i9*Ã›º§Zx¡žBÇiDºïûÉš]ÛG^œ‡ˆÇÙêk¦Y^p_pŸês!ôiÄlˆwïª"™Ù¡+@O°d3¢÷“ýÝ}Š«á­˜S@¹×KoªrOÓQ»B|ÿå`%MQL¡ë“ôOØ…/O¾PO%b7¢ÜYÝø®íÆ¦ÒUsCöQ‰‡Ï¸(>dù@Vºiøð†Ë6Zµ€[Ž×A}xÉ§&j¥›yá‘¦[éŠ<uêÉnî„ç¢RhZÃK5ÅuJ®ˆÏ¿²·Õ­*]L?x›Ø¼n ·õ´Æ*ÔµÉ|çú/jÝ<Ië:Os-¬ZÓùjØuL„Ùf³Ä1·/L-ðÍqV©r¾Åñ~2J'Æª,_du–ï1Œñ]µ¨uÕ“d%õq|ôãdjéz’,]ÿ ³’?æñ`Sr£Ì‹)½š™æW#N‘*M*pÈD<dª¨aÝõŠ_(	Z¥*ŽVyåAŠëc]ïŽbÔ±w•ûâ‘ÕWU¼æ0&WVÅ;ƒŠfw”„›È-¶Póê&·kbŠ$hÃ´-æòPH&w$ñPTì´Üb¼QªÈ \Jjû‰Ymv•Üêw…µ®í¸f9-ƒRÐ¸)›ÅÙæÔN×÷ ©ãz½óþòL¦Aø«±dÕ#qiü£Þ}G‰ê«©ÃŽ‚õ<Z%Â1u‘¤
%z?Åª¯õrM,çè¬¥Ž:B4¶’ÃÐw_µ,ÖQøê1fÔk«ŽñÔW’F›UYM]}A91›äÏ·6ÞnílÌ½JÔˆPOÛ­TIŠ¬Až›±,ÙJÑÚ'®ÎÎ±ñìŒOƒÉPEñ[<àÀšebPsóšL!Š¢Æ0H¥Šå–#xÇ«¸Z3W^*jQ×€¦ƒ9!æMFDÀ¯ù‹‹+bUZø^­ã«CbÕ×9^¯)ÙÊ-±šê\¯É”låœø,š¸$1¶o..J‡¤¨!*˜s¸ü2¡ßé"ÉfT¬Øm0Œ­Ry:Ù¹š/…ÓÒrÉô*«äÀ´Œ<Õ‰eî+àÅþñí³76w·÷%j¬z1rüç	gÞZ|’E|\ö5áGep7"õQ+Gƒpd5ð:Œu:F zýÆ @UÞ+lGm 3îGfvaØ’5Ãå²øC:d"n^ÙílF}ÄuC§´÷Ì“ã
S­—ñ6Y½þU71ƒ•í‚‹w Ì75ñöDÛ´CíÅäÙæ»ŽÙ¦B dÕDO«™žçôV˜Î(«GMÛ:1ÍÙ«ýí=˜­)¢ek¡¹vè
Õñ¬Mø«bÊŒ©+¢l]â§)6ø×@™â6àlà(óÖj -çC@vò±?œ¦9Z<Â–F?J;êS>úÇhé»€é·ŽñŸI†fÉ	ª‚G	p
Q¥ŽJÿVV—4l#(‹jBrrQhç/ÿ¦ïƒHAoöU |RŸ¦9µ†I¨ìQ%SPFEÊ»0CTOó UI¶¶Ÿml«ßß}}ÞÂÅâ×çè¦ü›TÉ›)¾bÚMê%nþƒFŽ¸>ÖC¡òX³‹À›ÖñÙÎ‹ªè9Ö@I9áb¦¹ 0n6u»a®k"Î´!}’:m_ÕWCOÚ« Ë¼›A°þLÈ}3ÀênOvŸk^VKTÃ×cWWiJzà½³iÝ 2²¥màÄZ½&€âÞl‡Ð+ñˆrZüÍ'¦HýŸî1E±Vw™tf¹ÏÄ§Í¦:sû´Sc]ùüúm#ä®×F7³oEímwîÓÝ½íÍý¹uEE¶ÍûTñå!ž+l_9+Þý+r^Óä6£d?O];QÑy2›ñ]”m}Žá.Ÿæ}ºÈÜ¸SW'g†ý&Á2|ÃqU—kÜo¡°õòÙÆ‹Û{õÃŠ¢#“C Êpdc3¢wMÓG%‹†>M-..MŸ+®(ï_›eºCŸbA·÷vv·vßnn¼¤¦Ž¦•åzNW^Ù$¿üŽÒµ´¬-'š°*PácV¢ªŠÕµ¦xS²E¡ª*QO¿
Èð¹	…™'Ô³CM¯¼Ò›UÄ¬Ò¹•Øò…;¹Âg¢o^é$é:‹6¡ºÁ¢5j,áQº¨.±ð1¹ÅÂÇèáJª´rMÔavÈÕíŸÀÏîëÜß}1áAztF}ç™Ö Y™{SÁ1Oôt[Å÷E6–®‚ÕF¦TáøÕÞ3®‚¿{ˆÚ²¨sŒÍkYc‡Â~,V/ŽNrja‹”AöAX¼c:ÞÆãã<~K—êí×ça$ì"ÂQ¼µöap?LP1ËNŽ‹5À-²t‚L£ß¬X]!SÏwÒ?¯æAƒHØƒå@µ÷0¤Îû0Õaœ£{_ÕGíùk…]9§n*7p>E‡Ÿ$$+˜†þBŠyÍ¬¿És˜Ž$ÇÝõìòoÐ>†´Â7îL¤
gFÓr*Dª¢våÏ¹Î•/—½Ñ¡T=zÊ’Y¯HPÓt”sŸ Ì9‡Ò3š@òôò¯À»Ýè”Í7œ¶˜;ÉÝáTg=ªFÅwŸbæ¸ouîdÒ'Ÿ3áÀ§ê¿
Œoª¸/ŸZWhÊ'è	w#TëL½å)Ž•ªi·ß'á^ªêà¾°lô’ìtŠû¸R:%ÇÁ§ÅÒ¿VÕ5¬ï„ÅÛï”êÔ«ê×A•ŠÞo¿[Ì3YÕ!.äò¿—Iz'Hõb¦tOØð ”=Cc¢OBg¤s4Ô@êÍs5¡óÖðíVuu³úDÄ·OÞ_îOa)Gœ<öçñøòßí8N«òˆ·Uâ‰ë«’z¤£>±øÏ*â›î´Ž:y˜›c.)Êø°7Ø´ˆ†qƒÌQZE–7‘ÓQÇTŸƒ]–ÍñÐ©—«8õówÆøãMÿåÔm:ãàÇÙ‡<žÔ]Â}®5!w¢G8t7ÁºSË‚.œ:–¨IÜûhà}Ô#UX%` V(#,ù|ãëµ^¼S'Lw*NW‚Bš§iN	n¤z?×'¦:»ßÿ^…µ‚IiBƒ§Àjµ”âŒ¼uØ¬‰žëÞO‹Ù‡Iù!IÆZ¤Àºwþ“Õ¦_òÀˆ7øÐ}ßGGyœÓë_<Y­µhúÕÜ(ªÞ=saÐ0¥6P{™¶Î•£{ÉyvE`,jqø©NÊpZ¸‚L³˜s/²ÓLN¯) €)ìWÓ³~Žj‡ÿ:\ÀÒÝÝóxtpc¦q¡<Iâ!,L™×–Jnu=ê‘!òn‹ ´ÚFù;ÿ^çzC[[,OÚ•ÚFŸtí‹m¾|Ú¾ÐAsšgè%§¨3ô3Î³öÅ˜ÇÓöå6Ø¹É\RófP#@­•èöYm `báŒð¶H$øœ3O‡,|lƒ¶Wmä†T®=#šC3*d»?¶…úì4Iž‚vúÙ0Ë‹Î¼¥ð4¢x’Ò"ÜKyV`‹÷—,eŒ!‘,!ØËeñ,½1JXqÕ6»^Ë‡…‡ä#Ì×âµ4âÆPŒÔlTÐ?þZ#fgoX_<‹Eý“8ß(»KÖ(œÎ8œôs=M0U5ôÇÕW$RGaƒÐ5ýbÉv:ÂÕÕúN{š;+í÷'G”§_0„Ë¼µNLNÐâõvÃ¾)YD_+¼ð@Y€Æxzä±ÃW=n–VëÊ¹]Š+uÍ†æÃûÑ«9>³¡@|ÁÆí1Äß©v$ç¤i,-@<#Ob<Ôª±·–jñÎãŠlÁ…¸8à‹Ñ%ŸLGñÚ"ûÐ:º•–Õ[ä!Ç/$Œ/ÜQŒ1‡«^oçŒ¬^~ºúž@À'Ü
Ø|} „ kØË†ŸÅžh¿ÎÙÅ-nÑÝÛ}¶½Ï%0y;¸Î%Xç>¨Î@t?`e¼(>ØÒ
¶´¬óUtU_//M>¾1È¬,*>’Mu®ÁãŠƒ•<}{`‡"ö0R2Jòx8¨
‹,oïžuå\HWòÝÐé'Ã00õ¢NÝËô4³ÒOwv°Ì‰›úÉZ¬ÙÅS“Ó¸±™ÅŸ¼~ð•«{ŠÞaœ…4×.º~¦é¹pa Mâ&d:?ŸQ9Eu€£	rÛk$—i9„ÊYœ÷œìýíù]ÈlíŸ’³=ì‚%=0ˆ’”ì– ñJŽ¶`P³Mš:ÎL–™•ñˆk‡c5jã,Ó‹Ýb[G©#¿sz±Ž•O5·N˜·2„c«ùªÂ\ÁÅ"´ %HLaksÁ\4#”jPº¡¡{mOV‰nB‰H-†§¬D$9;ì„r:ÝØJŠCF:rJ9žˆ{EÉ@U¹{YÖ^iüæ  F{ºž¯Âgmªs	/®ÕÙ¦h’¸`M
žr5|ÄÍî„8F”tÇÅÉ5á¤$D&q­‰%†|(¬­ßVÐ»wÈésÅÛ~íFó×¹ô#l‚ttŽ#’x|¿„rQ‚·Êƒ<› ü0GÁÍŸq½bÃ‰ÝZe¸Ø´ß~:NÑØxýüœdÀk¥å5ø)ú1žé–¢ïï›8ŽxŒ|­Ô²,µl*bºoä·3£øã½’5Ü?bÐãúEc­nƒ`þFoyW¼ƒdFé0[¿‰äéüâfs(ƒé»è¸ù Dbš‚esþÉiZ²mç[öüªµ*ÊÆûÓÃQZ
¹‡…ï·™:ÔØÚÚbêïÔ<í&"0Ô‚8ëàFe,RÊøLsX¢?cjU"ÚDOÙÝ¦uæ*Í6bOÅ·õÂ41¤4•€×KÓÄÒýÉQ½,$…”Dû±Pò©W¢êÉI|š¼L?6ºÃÓCê(*çñ¨^‡H©ã0+™—­z%òCH-Rè¬×Á“Cj`²:½8¦ùË^˜AÍ(+€®bhRyXôro÷ùîÁîÞÛ½WfötXUyTêÄ!õNŠjGÍ=2#·Æ©C¡ÕìüQõÂ¸{ï‘†÷šÑˆêì<ÕÒˆ&e…›âÎš6ÂèpaÙ*+ATv› —™Yx'Ú°¥S)Æür ªu†œl —üL¤õœ4`è÷'&	UWÇu“ØÔõ0ªÏXÔ³ÌZ/â±'Ë¢ÝöbREYD†æmÈðlâË2óB;®]Ë+½ùò)éîN¨ñÏÐ,¹ê’#un½\â–Ú*	¾yê6¢G!òKyíÓœMdëÃd(ù¦<É×;KKKÿÁ.­ÿíBål	©x¬›¼GçD:HžÛC©Þ‚`»+/TùP~\Æev3ËÍÎ­×»1,ók¿*ž¹IÐ©a±îÒÒYZbHÌ†Æ>pÄdgƒtŸÃº§[ÉiÚOnòªsVkè“E?ò¨AÀ6œgïß_^þ^<¤»Ÿ1·tãËÿ%9 ‡ƒLà¤!ýá—è5IÐ¨t±¤èfìò/$Ar"L×N&É0#E
™G1™Äyl“A°8D0.´äcXAt²‡Ñ2ãÈ$“;ö\ÿ‰t_Â¨‰n]DÉÓ­Ÿ*[¤Y¶Saˆí!+h½‚ï®g!!6÷Òz'PWOÛW¢AÇBxâÓ’±¨ªÑ7^OM[QTÖØ»ûrŒg²7áé½MÃÇ¥ßsg¶Uî¿°§2®Ö0ïR¬»Ito…M“Î™V„5!×Gs©½ðh£Ee	¹îÈM˜»Æ}À^<‚öàòn]õ¤ƒ,Ë$tæPè¬©Ó—™.µ3*Ý³ËÙÃÝ¼vf‚Æu›Q_n½¨¯8¢÷V0Fm“¼˜Oã€êÓq<Fß6™'ï›ÊŒ´R¬,Õ‡OõWÞ¹¾Ã³>#Öo=•ÏT)ô³ ÏØµ ‰Idl°PÔÈVnˆ’×­i‹¼éød‡Š`>Š©ù/¾@7FûF/W~>
Ö¼ÄÉºÌÖÍ	ûÓú… ÁÛÍÝçÛ{›;Ï`j[—vhàz6øÑà´Ý‡Ô*OltñîÙî˜^‰.ýÍ¤æ.5¹_9ØÄ»\òrçO7.%—Æ­7»(yd›/ŸÎ³Ë£ùJ“Mù,n“¸¼ü{žZÔzï¾X	oäI÷ûàªñ¡•dgkî¦áC*´†QòÎH˜Lów-áP\bRþ‚*í˜ÛjN€Í‡ÂÙÂêu»OWrÚg7™–<³ 7Î‘]Ñ	Ó5¾ÅJ¯®œXKš»ÐÔMsÉï¯ÕÑ5ÛFˆƒ!;"±ãRó@§{£¡þG´“•‰Ë2”~±B£8ïgì†ô#úèü”÷Á†Ã«èÇ°½ÙxZæé?cz'ÌÆÙRóG©Íå§ÒÒ¤;45ñ™M[Ÿ46ñ¹9­M|,ˆözµ7‰¢¶è´Dr(#t=?Ìf¯í%'§÷˜è†6è9Ì¯‚Otö 6 /¦².ñ¢Yß”mÞ ©®h×Aëx·…£ˆ© wœÆpþÀÑÐãÝ×FÄÆxãwžùy·=>MQ.L’Š<·ŠMâãoÂ5Æ{öÉj>³¦j¯ÑTÍ¡ã‹O°ž/uk5áN?,Í8 È«ÌæÉe ÐÔ½W¾ð‡PÌ{`wx Ó8—mÊšº¿!»ì·Ü#[i§'"V‡°›jòÍ¢SdcQh®žJµY8s0lWu®žöö‰ø(ÑµìÏWãšµ¤ï´ŽO½L4LÆÇå	Y#<e}c|„OkoF”sL4€	%”¦ã­æ(gy@ú€<úš‘WPI«cÙýY½sO´ {öçÂ¼yª‡z3£ú²
za>Í|uëñÌlc-aŠ&ðµ1X@£¤ìŸ-Áb<Ià[˜|X`ú»AÓíï>£¤<ÉÐÃãËÝýƒ z	AçOI^ô›€Q ¾lÉŒšÖ±;‘÷=aÙÐ²©Gj~íCûJcé	"*PV¢lá^c?‡V"nšzù-7VÁ¥aš—[¡Eäƒçú	Ìw?Ä>‹Òê‰‘m3Sý°s!ù@gîº¨ïçà,äž‰ò~ªƒ!%h! Qþ¼aØ†ðú Oú½+’®ÚÛ€nÉ°´4PÄªÂfÆQR h:/€ÃÂŽVð\’ž&ƒ^ ø¨c)0~j9Fù#L¥ÚFP§kE¢bzÈÐPŽ±Ë+Ks¼º.avõ ²ü)þ3œgB:^žäÙj³ÕˆciÞíJLƒd5ËÇÙ”Eè P)Ÿ^þe˜bÒýÃÁógÐ¥¯ÏÙì\DQÔ™‹ÈNQd4:Æ)heý,ÏzR`œ›uEšl³cðZÀý§â
+2$8‘ ©ƒ&itÑ¤@o¯y2"eVŠªiSt9<€âÐSh#« 'œ†??Í†§\–×B @/4jfáÄh³y	…âÓ,Ÿ'Ó2…ƒdÖ©Î6ÓÒÚ–Ç¨=õÕ!ñaœ~ÌæÉ/Àá]þ•À™ý¿1"Iˆˆia\hT66ôÑ!ðRðŸAR|õîš ÙËý0&–¡€Ò>ÎMØfoÁD€ñÃyâ<ir(ÄZx• ¹@ÂÅz„5‚p€óØ‚D´
dKúñƒ/E¤A}Õâ9cÜfµ†fàfg¿}µ+–nð•˜¦}jlUDk.S¢ç—f(ë[=­\”î‡!Ü| 9Š˜üí8¿;}êhò‡_õxîo'êz7Ý4Ÿ¦@c…š@:šžÇ1QÞõàËwŠ‹båBDd9ºOãw¡O´<Æ¯u÷ŸöHR9h‘ŒÒJŠ-?ôHBÅc†Kì26[Gó$þ™ú¼x¤¯ÂEPez™Je¨“b`À´<+³løsZvæBV•v¯¶²£š•U,«R]W³s`ÿ°krj¬ê{‡,¼þl0Bÿ"A*‚ä€Å¡™ÓfoQ~ÜˆhGÉÆËóqÈ·k³•ý˜eÇÃ„ìðe$l×ù÷6uó×O‹˜F~§\Î ×ð:Eë=0‘„5Xßæ0›C„èÓ2…zU¶°ÍQRr—„Æ¦Ç({fEúú¾–&—nøŠ‹I1êÑßÈŠÃïÊû¼@G+UòÙ½KX‰îÀšÍ’ ù`Ì¼õÎIYNŠÞâ¢8 ¡0o˜MÐÙö1]$¼;D1SAeMµ}Îóàgš…ç§iòá1]è~¹~Œ÷—¼-|÷`9Yö¹vªf‰´Þy{8ŒÇ?‡—¾Þg0Œ°4–ùHµÝæ¥ù(K—Ž‡é8Yð^fŠU\Fm„Uê“ˆŠz1o¸s¯®ê+ìäºfŽ—'ß„Ž&lYÐB„îcœaòRuH·¶Üa‡XÚ¼Óv3wüyoˆABÃPâDî­¬z]žUÏ¯c_çI2Ö7Kúuí©«Ý¢t×!Ÿ|+yœæ·ÎètìXÏœGˆáÖiày®¢†D™_à¥=•rxGìP¾S²¹Ôðª‡é’a§ý{Âã¯¢z4¹­ô7©¦–Ð]þmœŽ2íÐb‰©>\¸~âç9[øöÉOmÍöùÊ¢§ª" Á¦€ŠÏîõ…²‰Ù´¤øxì°,Óá£[«@æñzvÄÇ¯ÊX=8.“ÁúyÚoø&ºBdmâ¯%¢”è»,½ª&ö1*ºb/±šìC2ð«ÆUg^z’¥Ô	Ÿs&|€o˜Z¯ûQ|ùcZXãAEQñû"¥µQ”vŽ’ë}ßâZ`¦J)\ÒJ.×¼-z(šßej¥öêÊR;^¢È
¯0\QEh¯çF{môlzŒ‡ÓºøE‰Î@™›1t°ââ×ÏdôQDc1[ØÏGL}²©g-ü‹~3Sã‹=d…2÷~j
HïÅrÆ°dIcØëCƒ˜tE‘“®´’’:åo¬ÁïMñø|ðZœ·!÷ ãŸÐG¹ðÌZÀ[ç@¹à=ƒ¥Cz„òÓX<ð=É>®w–ÈY¹ÿùJ ”íg4¯a47ñXs5t¿6‰Ëÿ…m÷0ýx‚&â°ì~^­*õ(SËbLåÉú9b?5Xï<_^!ßŸ®Œ–È½“hiy´ð ú~õ!¼,¯Fï?è~g9º¿´üøƒß‘åè»Õ•…Õg~{ý…è»ïàÛêê*p ¾¿Ç¯F÷@Ùg«Ñê=²ü€f#ôS´üý
Ye‘Õ?_™§ ugÙK(©‡¿¨PÒË Õ|À‰PœrÉ5Ôo¶Qåþ4+˜DTÄë%	A®ã8g>/ð>º)æIì—¹²gæwûB,[×úf8TL}ùWla
s KïŽŠáêd»äÁWœíÍNFÕÀ¨Ppšõ/ÿƒL2œäãôÄ9É 5M?¢&‚³:]]›:‹MQéÆËr´²eëš¹\Í°k
‘\þ{ç7XtÚKœ¡)WG¹ÆaFã±3½†™§ËI
íG{E½x¥ÇÌø*íbØ@IéP.®ÓIAÜœqCtþék 	kgüð…àT/ÿ:b0Â]ùÔöýy(¤€
B³77“±@Èq“5Y6× ¤í WÐÐãþÂžûQÍ£<@¡òÝÀèû8`!%*ôb®å+·Œ~ìUjpK<üza*ˆL5½¤–nâÛâxZž„¨gØçâZôËÛhÉo[—.?ìUŸ)LQ>\›)`(WÒdšIßg…¾WØFùÔ
;Îï®¯¡«jr¬ûjÝï—‰(„rÉ¶j’,~q‰S¨œÉEØÜJ˜à_)Ü‰ÍÔÚj/l-V[þ‹¹nÍ’8UâÛ‰p¢?_"œ4ë ÷?•å©J¯#ÂÉŽ—\%ð7ã$Àðö#p»|_WÔE½†zWÌ%5ä…Rëª" (Åã¨Ìö™%3öéX,f\¬ââ"ÙÌ€>‚Þ7 s‘`ùàª(‡Îb(O&¥$§YŠ—ÇcüœMKv¶¶ÔÍCÏ@½›Ôæ t;ùc¼jŸö§èiyBÕŠ{u11²·¦ÕË—v–WVïÝà
jcZØ ˜]àKf>ÕÕg ÂWW|cJã³4‚­Ö^p8N>¼‚åZ'† ùœÕÀJXÜÐî‡âvz¬Aû¬¹ã]à§ëŒðÃ·½ŽZ¨tC£ yÛÌá¨C¬
Ã=_:Í™ƒŒ¯Âéz‚©áºþøGÞ!ŒºÞFÞºd¼PJñíùFÀú0=‰Êò™3Gß`£ìQ Hò ¶¢ŒG“®Ã¤’ašv¥,a•¸ý‘=šrß‚ü¸ž
âPƒ+‰Ÿ¶Ö­ÈŠFñš‘L=MsêËÚ&°’r+ëÛ©Æ >çÉæî³gÛ›;»/ö£WûÛ{ûó™9¦^¦-“Ý†Õ{(ï £!·à·ê“xÃYnze¥f6Ö–ÖaH?¥ûG9v§âi ¡?Lâ1 ¥{Q ’ÝiÙÕ¶¥ÕÝäµCùúÎÂ%–‚tqÅOiòa®×q‘à~†èeøË±Ë™D„eå/\_Âñ“å¥?ÛÓu/Ô(ýð¸Aîà:Œè<³
êâ.R"¼a¿ÎÒñ0sNÉõcÒ¡Nøãýþò/hâ‰ÿŽðÖÇY´W	œÈs)»n!qòÈ™l`eHýmíf£­ýv#Ý°;³.3Â¹?7@í·»vŸ6Ú—pfçnƒËçÎì·»fŸ(È—h_¢yÜx_¢y4
|‰æA+ýu{ò¿ëÑ<îP4Ž/á7ª±|	¿¡}¨é 3ÁÐ`¸ H=ŽmþÇÝF£7³ç¯#*¢×zÔ½ùc#ÞRxC§ïäk	qx¥ ‡äêaC^eÛ¸7Ž[›êÀ¿ˆÂ2#h»¢â.·2#¾– Ê.ÖëfƒœÌJÙ5÷nÁ‘Œƒtƒ€ätÓë÷y¬„'Ñ®/ÑÌa~ð¹[KûÛŽ1´IïèœŠ•7^¨úV ™U¨t¥0|÷Ñ4û >äÁ}’2ÞÞÑÕ†×žÓT¼&,Ãú9û{!.X×ÏùE™RöÀÔÈ&*t¦†ØÙ{e²ß³7º•Æõ-ÖV5Ä±ÆçÏ¨>Œ·œ•ad!{U}™¦ñ=È;ú¸OË¬†¥ŒþÙœás Õ5£f\]7X„?{ÎæÎ Àx—y Ê÷’c´¤zô˜Ÿl™ÃÓõP˜Rù~Î%ë¼‡IÙ8OxpZ1ŒÅŽp|ÆY5Mf³lŽ‰sñjá Ûþ8ÉrtbÎaŽqî#Æ³ÛD{LkpÚ1hƒ¼ÌÑA4ßbþ8®¬åè´Ö
ž•ƒ€’{	õ2=ÈêÅsñ! ìö~2J'õ®ë‰îî*€!ÖPÃPJ`e8êYHŽƒ¿¹ žûÊT]VÊ¾úŠc'.ÿÚ§ýŒ÷ò ¡o¨±·äîªVúT²ê°V\ö8¨
á—/~ÌãÁ”^8Ï
–ØCÆ`¬>¸öMuU; ¬Æ‹9‹ ;¡øñF?vU2ÏÆòÈ[zqböíSÄZ•ŒÐ»ý€2Æ’õ‰u¦Iq…§zµk0ÜM; Ä¾Ë!¨Œ<áVöa<ÌbéEçvÂ¦"nºppOÍ3C¨«Ë+^7U3Ì…qè¯&ÍïŒ,wž39CEK3ï÷ûÉ¤\ïD‡ÅÇy‚¼L¶=<`S.áT7,	UG]'RïÅãèõÒ—’íW˜mÎ6)¥óõ4ÏFl'a™yÂnp[y.e1 	«#A#€%ûµÔªÿ	Ó•$;"¬z×½ßþ˜ê{ÝAf ófŽú³HÇS‹á‹x8ï‘õ¹)ÏÂu@Y¥!ßÅžÊcôÈ‹)º§éÒò5†äÍœ£ŠGc=ô
ëœJ`nD¯×ÈÇV^çSšSPc*Â§ÁU±‰ëi;®ú­<Qð´p¤9‚c	ŸcUÛÔvèÆ›LUð`9ûÒ¬ävÂGk¬Kç¼Ú×X¥%¯ÊÕ6}#j±õ­›ù¶Ö(ÀZ¿ƒ¡ld6›|¬6$ìaŠþñ`€ÖÀ(™Ä¡akñ|û`ãíÖÎÆÜ¼@øw ‚ }û­[ï¤²ðiúèfX`FýÇUÓÚ—Äø[Òõeb‚‚	–±øª3_y›°vÃnÐ¸'ét~@«Ä…fîÊÈj¤¢9±QZ³˜SØåAò	ôa¢µâòŸÁÜÏŠŒ6h`lÚ&Ý¹¨Ìvöw…mTL†iÙítæ€·³í³°n¿Ã`Ígd¬¹›tÝÕà¼Vâëèˆ$V®ø35‰Ÿ£MI1\mú35)£M7;×;,§ë[à³ŸN”ÄKª+d~>;òn¢þà½’Rs¤Û!·v­nr‹L(îà˜6&ƒßÙ&É·AŽìÖUY&£A/DuÅvÉ×<¬7Ö•Çák?“ü—¶´4 SÎÙõuR‚Žõ®‹¹ìÅ£Ö7v/­Ž{BïÙÝ®tî4:©”óõ¥²d3•T&¨EQUý<§ 5"jQ³òºú¨ÝV®†ÜV¶r<îºÛ4Rp³—Ó¥ƒê‘Äè'¦ƒQ‡D„™/ˆšµÔ›ö{ôk89V®Î¹gcã>B4'=[ïÐ0ª%Ó€¹m‡ãF¿4ô	¿ó¸Þ¥ïA¦Î9“¯€¾ÊXÌ‘o­¬¦©¤N·Û–æ¤·m1U:n×4´b›@=CÕ§»Ím{°Ix…· Qÿr°u³%q|€ÒXÐv‹Û`[Ÿ	¦j÷6·TŸjäáþfaG¹3i>†“I»Õ´œ*f¬d&€2Þä]X“ÍÁ>Î›†/¦3êÏÉYmR¤\Ót& \d{žâ!Ófö” ·8PB¸r+gˆ‹ËKÆ’10K•0sÁ¶…lº1æ“«¹Ðe(cˆ!oŸ·šL8pîJÕæ/Ë‘3Èß¼sÐŽœÇ€¼Æ¹ô–k5ŸBÞ8‘®ìµLÇƒô83Èß¼3èj@N+“qÎìZMG0ÎÉ¢Ê`•\;`¢êæ3U]‰{¦ÊÕ9U®LÆ©²h5UªÜÃ9_‘wÀŒM¦ùd(7'ó#8_—*çËiFrîRé{Ã°pwY'ÖL^I	³ž å"øéuwrOKbSÿS‹‹¦ÀŠf?Ÿ÷L.Þ­BKXŽØ¬Šø8)¸˜Fº¶xroS¦šxhõf,›‚™HŸV8Þmu7ÆÙ@v6Íò9ò­CWRg¢ž1e²ìÚÜÞ°w¡Ïk¶ŒqŒ>Ú h³#Úü©ò¼q]½[™„Ý°G9³Ê¼‘—w5‹Wy‡‘ßM›³\O¼[4ž¸¶MvÒª»Ÿé†jô™ïÁ—-£<wzËTªcÔE‰?‹Mã ¢Æ~’ÜškWÍ xwf»É~ÙtÊÓÖÌâ [É¦þà¶sÛ‰†¼§—Qt¢•9Ž¾Ú"˜ªå~Ê0i6Ùšß`‡šììÇCŒÿ½!b›ýŸÄ¢Ü“yƒŒtlb?óvÓÄM1TSR1œÖ«Ž¾kÑ³œÖYí,³ê1ÜvY×ÐÁd°ÕÂ`ë°F—ÏQ6Ë×©€…BëÃ¸îfÄhPå‹#&cŸŽ Ð¦#C8œ rVœ¨a2>.Oš6]“:D”j“.—Èh:ùK
*ÍÏ¦39’AQQË.žïÖÊ“$ÑX™×àL^üÚAÁÌ:(gt}†ƒ"ÉBkÂ‡UJX[,OBKh!vèïˆÅ00úûþÒâŠý¢ãIôJ—ÃÁNÏ‹ô8°H9› «tÑ{ywÛâ2kW—tQòít‹‚]ÝÚÞØº¾„\šéòÜÛé—h:¡©&6½ñŽ©âÁûö~±KÏîLçÑcl@znbê,øs­<Ìgjƒ€ÝÉ-œþ£†4uòc[§¡ßKÖØñk¦.‹Oö3‹}àmÖÏ¿bÐ'›ðVáx¼%Iûí1ÇbRZaÇ(•'€a€À ýµE=…Å1R6Ûbº{¿†cÓÈ}–‹Š,/»ÝxžRÙû!ÕJ‹¨;Ë=oÇyÒeªjVW4L~¼“Ç‚‹YB•Y¹±át§‰Ui*BŽ{p™»†ÄJðÛ0gÖ.Ï«Ý©•h˜:ôðåÄ ¦ðÎŒvùìhãŽ5¨HøäÔYÂŠÌ4=ˆBýóc¸n÷Ùr¯Ü¢`øŒ™55Ú4Ížµôâ"y:³CX™‘ã¤d‰b Øåi_+±×iaÕóÂ‡­ÔD#ÍoÒzÖIû×#Lž5ÏîÙ‹×ß.Õ‚«û‹¹ƒÉr«:ïRSòºÏ¥í­Ñöô:ó¬HÚV¨U eÁ\8ÖWíô'WºB¼ó›R
BéÀT5+®‡ó[\]
1¯—ø’ó:sc<L?„‰¤â&‹gòxâ’Ï¹4}ñdäûIZ2,-õè^‡ºhRóŒJl†[Öt&åÂ“½ŽÕ­º•ÊËÍÉ©±¡@Y™Y…N°;Þò¶´I“Í”×=p·ªŸ£Ç£è¹ êö*v{qÆÊjÖ!ê
Ý¤¡“Q#Ãb{W§Ã;Èª±óþ8$Ø²FB
ÌH®í„9ˆ1Ô¶”!zÌLiH„à¬çš´ë ×«"·O‚Îîê>½Y´uWG}§±Ó¬ˆå3À¿±íÿyê›@ŸÇÈï ¸
ñ™°¿iŽ€KqîêŽ¸	\pÇ‡|‘€Ç»£¡¨ß£¤ú|’m`ðíwU{WwÀMlúÏaÜwpç‡xI5”tÒª>Ÿ\7T|Fˆ q¥rWwÅH>—ÁßA”ì±ØPI‡ÇêóiÄ7$×q"ï§V–OWFçTSáâ×°Ï™jÁ¯dSë;–«Ìsõ	{@+ñ„MÖy'à»ýÍ§ÑRŽ{ (¹"ÃèOãÃúÓt„HõŠÜN4•ÂÁ%ÕÃÜdÉ;à0ì]÷Ø¸O­¦æ;q6©1>&çŠE¯*ê6ùj4JX[LŠ±¾™dhøØœ>Î8A³uBñ#©›ZL‹^C;©>M”Wlhn*çéZþ?   ÿÿì}ËrG–è¾¿"…ÑxÀn|èÑ2Ú2")™Ý¢È&©î™àUHI H–TÁU‰š«YÌ?Ücž¯nÜÍ,/ìž“™U•Y•¯@‰’•Ñm@¾óäyåyˆu©ÝÌ .)‹Ë{›ñxg»<Ó$U4Ár‡Í–2+_‰E³sVT2Û,Ì1Õ²²BIŒþ#1Ë¢HÒqgç£÷^Íß…Ñ0~×MYG±7ÉÈâ‰%èiµœçômÈwðÙ9-‹xñÛmxÔjQø‹5t—[ï>P6*þvVåüÕÄl)Z=uÖVT=í„Y¸Í(Œ7§AÏ]åfIŒ‰Fj5›3/"?EM=!CÀß ŽNÃdìÛŠÖÐžï)	.£i˜IŒi„ÌÜ˜e”Rô­3¡ox‚Ÿ-_´RËs_|rlÈ…Ç\£ 0ìºwC‚±×›T‡'¾{Â‹.æ{³„1f“FBQÕ}bÉ£¹7˜HKãn*ü ¯Æ]o k0Y9~Ðè½?èÍ¶ÎíÖ„+¿gù•j²:è†…â¿‘ñ£#€Š&ÓaßÜ
*ÂoadÄ¸¢¬búM®<¼X2%h~Ðè; {täªûy3ç[­S¯üE%i¥Ž9;égÏÃ3:ˆƒZâPCO5ŠíbÓx~¸¤ðzx7s-úê™¯ææ0QãIŒ“Î=}@j–P\S__½Ò‡ÑÖ«°òPågJua[5ÙEÿúŒ`}sÐK~Qµ­O†ÑÚ$o87G1Ž;;GYÕD¡¢xÀnÉÂ0Då	htœÇ	MÂØÜÏ<)Ææ@:Ú$En¾R'ð[€Ûœt¡8°I»Àãêé¢vû§]ðÌ«nˆ-˜Ø2/¼ XÔÍaŽ,	8«¹²$ÒÅ›²€<	Ò•ñ„Šo–Åm¼…¹äÌ ¾é*êñtÉ‚”˜¦±6ës³ü
\·Bšâ‚8~l¸:F ÿÓÁ·c~j,šÀÖ¼´ƒ38`Ú9Â„1zë(H®ÿá¬õ×)M2Zaä®u›í®týó	pSæ
[ñÎÅTAÄo†Ô½•!F¦—y(B^Þû¶4õ˜/Ú
º,²•Ü°ÿÇµÕÞfîÞ.ÇŸƒó`ðÃI|a“X`XA—ãÐ#|†AÊÎÆ&ày'‰å5\ª“p‹â¾ŒQYªXÚÙ@æ™*ËmYÆjÎ¤ª.åë	F€´n×¢ppb(1ÍG{ÈŽHî€@7t…ÚÃâTJ97Òò«U+!áq‡ë¢’*ÍŒÃâèaÑiä´.·ÎðM’ƒÅŽÿìÖÖ@”z£ˆ&1ÿn	åÿ.NPÑ¨“FçbVÃñ|)½Ê['$²Åò«F˜ðF… ¨²¸¦Ÿ*çëËÈê¿\\KlDþåbCYÊª¢Í²Ô¾bWªJŠšž›òKìÉ¢<ÓŒâ{ÍA‘¨·£ªÜ<nˆ@É£õÈe©,n~Ò¶í_¢OS¬EŸD-@“@s­o1îu!e7kšÓéÏ=²]Šo8¸ÝýÛÜn@‡”“Å(é
.ùPÑíÊýU‰BgTCX"%ùEIZX„$‹háÉ.P°Å˜õ§l,ë[e}&öûIBßwÃ”ýËƒ–r¨Ýœdƒ¨µ»ßÇaÔÆòöv½J»Å­T¦dþRÙ6°)YE™&4»]¸åSCƒmÝ»˜—±vs×ÎÀÜH'2;DòR…K&Uôh‚ã«­ÌéÔ±Ì¯ƒe)Ä?0çÕÜhì'sËHÍ£[ä˜îƒ˜0â»úúìÁíœ¶‹>æ&3ÜZþ ¹Å¬æÔ{	DÍás›ì{»ë´éqo¯ûšt´QpoÎi<:£Ìþ¦hÙ®£s¾s0M %/Ó yÉc2ÓyWúv£;‡ªµ›ÂOûI|‚(õ˜}J»§€	Úí)iŠÍØtÔÔ¥¼/®š†ÁÓ Í0Š(ŽÃ(ÿœë®ÛIð£†*áKÄi>E ÏS:Jƒ?ÕêJU»€³¾i}‘ì3’GÑE–L½{x~ý+Hc+O·þÖª?„ãwÂšºi·PÐôƒ¢»C:b²ú.Í’ë_Óm‘Ÿ~ªµo;æ#ƒ‚qè®ñ—gXõ
¿ªmN¥¾~¯žÁ¹Æ	ÙBž°Ô­õ¥™&Ìº¹hIÚÅŸK*.YÄ!°ÍÓìº­yFŽÝ3í•îÒL‚hl×&`Þ.Lõiß{Ú<oÒM3šMS>ã}è·TyÙ]¢lýâÓÉ$Ä>œà>vT}s6œß>\ÄÀÁ`šš–¾ÔS8…ÃÇ-†¤:Ûv²nùYaÂœïŠZÆC˜YZÉ—Sc3.JÇÞÆE^æ”&­g©5?¹ë%ÂºFô*Æ)FƒF¶± üÓCÉHlýbdÒim²g%aÍx_kÍh1Î5ç”³¿ãé¬ßø—éýÊœ«BÜ¶ª¦5Ÿ¹ùyO?EÌbËç8¢ùƒ›[/p)cÅf‡ÿàï,	‚¨€?þÉþÐÐa3L£`ýSÃ÷~d0TˆÉoxPˆ{Âà¿røû”Á/§Àü–b±à§ý²úí‡z™dƒ5d‘À] §ÁþNâw†'K`¯’ë)·ÂœGÜƒšWÎât½^9y9Ü9‹@Aê)cÔüqý‹&Û^~RõOñ‹ëÝmÆYgµûÀ8©R$>10™Ð„’º¨³Œ§ø@kí(HYW?.NÁ2y¯cÞ-þ¶Âþ[$†5O, ®ù¹×§ŒÛ&,•0 –õ¡³´uYg®1ÿ–5Ù ¢rü©É<×ôäsÜ:ŠßàSB)ðÈ2ˆ$¼Â|¸LèÐkø+Ó©ý‰É¡µcÏL§Öº²F/²C&Áðm‡Á}Þ¦.ßŒ\+ª¸»6í{*Ëä§S«Vnƒ´ä®’‹D?•ÐEÖîz¤U¹b5·“Gµ`HÆ¯Þ˜¶Í¬"4Ÿ’]½8‹½]íûËQL‘QïËú0ÃË²Õ½M`d=#\¬W(º‰é8N¿œo¾3ì…`F™öJšÂ1Ž–NÂ¨¢»7«ÛM¹Œ«ì€ÎîRg‚ÂË&M’àŒFÃ˜¤
½èvõ¯×Ù|º®i*<Ìêä:?*ý{¾}¥üÕžÊ‹ü/æ¸Í”@jÈ)§‰œpáILm|BÒ©þå
€a’ž„£pH¡'ø'>¬oÝîYvÈg\v=¢Õ"¬{ò®L6>ùïFKŸ
Êª³ÑF»W#KfÙ­Ò\ s·x/Lúhžâ 3  f1ÈqõvNæh½ wTxöþÌ°Y8 YøÖbäê†“ÒÙÛ;ì’Ds‹·ÑF)y.K%»‰ùeY¶ÀËk€ëX­U4\„àÇtÊuµ(º[¹ÛÓµíQŸ³o	_½‡•W?‡i_(+Èc"ëÕÙöî¼·!{z+˜Q¿Þö¹>ÏÐYÁñZÌaYY!ß£	RI‰f/ ïB¸[[+»»+ÿÅcJ¼õ“³¢aVÀaÎMº½ÌJØ“ž”‘Ï±Å›¡øz8ì¦“Q˜µ[–Ó|AÁšÊ„ùž£rºoî^²vÇë¯®Vò¿×¤¿W_]½qöÓ#C‡Q„ãØ<rbqåÄÂÄ’„\i±Xá4LˆÅ…Hê6Q%^v650ÉZUž>ž‚»ÇwGäþ/Y ¨Ln.‡1¶Äi­_“|]:†j)ŽSz¯D;àîàœ&ý¬½ºäFÂei½ÔØík÷Å"¥¨ê]ÓÄñëân0Ý¤÷6™7Êo¹vµ¥Ç:dp³5ýBV…ÏÕ7±*ï#÷¬È­sn!xÆ"ëUÂˆ¹"qõ·°‘‚!W%5K)uÔû­ú¹äg—ß-ì¸œæŽŠ”®™L“ÉHbÅç?ÚÜ«¥WjvŠ~Ø'ï^Ìê¹øÁÚ¥¼7>Ýúòœ@göütj=k1ƒšNØ3æ7–K™ËÙ^åx'B,g¬l¾)¾XÀÜ/#š]ÿr³s»©}Žº­óÂä˜^tÞ1)1™FdÉbZ«P›¥kˆÏ—ð{Y0÷Ú<àž[wõí¯Zn¦ ¤Ž3rë©HÄ‡³µ¤+¥0ÛŽpc‰ýó©H)ûz7Ê‡ÇWòbpüÐhèœ„q‹£¢þñã1I«ðÕWÂ¼DèÒ×X°6?6íRÒ'`?ÿ:[¹ME²´jÔª€X,q):§BÙ²',²‘6æ›v¬Z ½2PÁñ×Þ›K[‡V}Égì+CùîŒ¹ùxP"“ûcm<o‰Ü'5…èÓ;Æ3/êkk„’{ZÂa ô©Í‚ðh7Xo:ë*~‘œ±øÊ†¤(‚Z­•FvŠ·HA-äWa­ëˆBÞÔQ}Ãö¥þ‡«AÆþmÙµå§”4Ôx2õ?sP×k»Ê‡õÜv¯t‹
j>n/˜ÿë< ÎOçFAÜ[©dö¯u9›îÍ›ü°KuŠ4“Y<wÖ±žfÙÞ#®ŒÓ0¸»?ÖÞòŸô†õh 58msiúWKm³ŸÜI<Þ
©gœêG7§ZÈç÷1³Öºá#W7³­Ìãý<a['EÞ¤É°¥‰ìÌ¾oÇYô|³Aœù îÎ¢žI¬ÎÂl:²^&d 2(bÑnH¥æ=¦|™ÿÙÐ!üM®lÀ4æé‚Fqá‚ycÏÞ¦ô™½ý<q¨íQ§k±ª?N”iå+SiS"Lk¬káölÁPAøŒ±ö¾9J¸0÷r’3	÷¯ŒúÓ–Á‘·£ë_ñ®!G	ä÷&²gágàg	Ï§„@Þ¾è‘ÍþþQÿú?¯ÿc<yIö¶·_lîôŸ“õÝ5sU†÷ãçÖåw^³^ýT¼"b÷ñË>ñ	Aí¸5‹Â'ž±}#ðÕ @oÜª‰¤¬Ã³?ïÁxT‹iÄ@-–†Æ\  
/<Nð´ØÊt‹O,|¦µ¥Ùg‹œ…s®¦J•™–G¤†šwžU°ÎµÅÄZñÍÉu¢¦
•	2ˆóÐí9=bJ° ÄÚúE<¾-»ðÁ~…1Ë-ø#¹„‰;ÛZ&y7laŽ\‹‰Ïï{'xø#§&­_±¶ózC÷óû>¾p>9L³°ÙXCÒÖ^N*LFiž{°f‘ûv^n’ßå!K8^Aã…G¦ø:aû­±i±fèì2sAžLïó"6úÏÀ?ñ)[üStbíƒùÝÐÈ–¶q©‹¬§±‚ÕŸÆ?º9·&„#Év m•$[d/Ñ…ü[kKÛûªç4 „ŽÛÞ'·ˆ³§ÇkmKÂAêÎÅˆ·®T³Fq„ýh9K4)3«´îÃ*Yo˜½UÌª¬ÊvûGäéÎ‹/ØŒ•„Í¸üüŸ}ø,×†|Áhä¶`´þfëJcåÃ¡4Ô~ÁiŸNãjÝ/H­QÈ¥òÝ©¨·¸Ü:ÿzfé)+ª¡8óvT“Xi3JäfÞîÜ"ÛÓêh‘:~Žùnàmw0…ÿæÕÝÉ<²!à"ì*–uƒŠå’?r}× æ6ƒ²›ÜÈ*Q­eïéºÂk–ô~*13û¢„{1#_£x0DR©öf07x`‹UÄmÍŸÇ:BsóCæLÉÎ…Yœåí¡]ôû1Ÿ1žÝÉ;¼xí­–A ™†xôüE:6Ý’pY‹â³«{°-š8xåy–5’¾±5ä²¬þi®j£¿Šàm$ãÊU_—‡¬ÚÃuûØ>™wŸ<FßÚ#£m“ÝÉšiú>LgÉ{+SHßÑ0ñ]®·âíB…áÉ2ÙÜ{þ|{óhgïÅa÷ÉÞîë­þ2x×ÁÜ¹Ôx0epUî°îø§9X@+CGG1uY¾aLÆÝKi>h‚ÂÂ•cf~Â¾EsÁ+L[“Ä€Ah÷ÍŒÓº‚Ìç¤$‰=·e>ýÖv’Ä˜ÞŠ²äW	y¦a#¥‹Q¦ «8±¹¥›nQ£ëU—dHÆíi]hrdQ+ç±#“ˆ’v ìdöØÉ$x,â[ÓBÝó
0|³ý>Ø;=Uî=ŒŽ_+ß™ÈåCâ@ ^eOÃdÜnmið=†ßaíI‡‹ËÛh-ÙÁÏ…YrÜ2d)œ¸e!ØÅy¡¸1ìN'Á8~w¸kØ _Èj"\gk÷2ß[¶¯ûhÓXÌsWË”
Ê5ÍS*L$	w^Rx”Ðô|NZØTöU¿É…	%ò•ÞÆOá”ABé äõfËg…2ð#-c*òVá±K9«êaª4³ËÖkžDI0€©|¶&¨u©rq&¨OÅæIF¨O‹ýzÃÌO‹oš˜ =ß¬j>ŒÛµ¨iBv<7’Ùµôöfù¡	sX^êfñÎáž$—ò˜9G­¥ãUc:!áÝ<o7B&ú·£-”3•P¢	ÿ ×R‡ÉcÁà )ÊÌ{¦§•VCÔªª›ÐØ8Gi¯³`UÂ{R×…Ø™9•´ñ1ˆ9ˆK\ Ùéšœ–£ùxÆÜ{FK Þ#ò/–ÅÝ_¼Éè¼
öºí-ä‘'×W×ïw5‘>M¦£•/=¢ç@o”ìD×¿4ÉÆ›Wó£p¯Û®f U"ô…–ÁVÀÍƒÀødÿèðy!©o!œ¹¿9ìUð·Ý. É­ÅIX¦úÖÌ
$fã…f`R²s·Éì¥Þ8àZò€ºƒœ¿u Wˆ·Þ$Iå¸Í‡äJaîÖ*kÞÈ«Ê¿Ÿ#øÍk6£ˆœë™{úŒ#.ã·‰ÌfJÚ©Úš¿XÇx¨ýòÝKI?ß.Æ8Æ•RÀ#™€99}€â wÂ c„jìúû"ˆjCÌëÍíÜÁûÍm#íœ…]š¥‡‚™¥ñ?˜À,‚ÔmØ™›>£“ÙïÓ€Ç,-ûL;cj©bLeÐ ‰ˆ/æ/$¯´+é¦q’µÛt™œ0;eª¶î¤‚Íx<¡IÐ>a_jSvç¹ N­oµy¾Ýd€püØP‹SI"ß’UcEŒfƒ–L,YQ/‘ßØT¬?/‚a{ÍÏ¨GZ«ÆøøåÜœœ“/9‰Pv°ðEœÑ	_‘îÔ6ËçÝCš)‹ƒ¹€Iú{œ BÝ¢ïá—1#|È8-¹:ö–Ý6ëÎ–³ýöqmºÖMnYcRºÃíÒì¼KOÒ¶Ø¼%8(öÝ˜^´×–ÕÍ°Í’Ò¸ÅÎLö<)š#ï€-ã@“tf¿¢‰x_
îÜ$!/°(Ó#@1¤#r“qJž¤ÅDJU­xÔ—<ÏÊ–+ö\8µ4h;+~‡Q»µÒZººlG ßŽ»Êeq¥?…ßÀT`¸ßâ„»–oÑó
ÆABGC¿©YõAóÎÌ1¼OìçJNÆu;PÍ"…qR¼à9–n „Æƒ çS0–wÊâwE/¶/•egsºúgûž8ƒÚ-
L|R^±‰W	ïmº[9ƒfÝUIYG©À`mÆw/së[ž—°¥…®ÒnÍOŽe‹ñ>ÄIxSjÎÏ]­­Ît7…J¼Ý/l~¯¢/ÏÀ°Þ.eÑØ0:m‰†>ªOµpÎAù•_£e†Ç¿5{’,ø¿vò£“L,ýZKO9óYVWQO/6‘²Ô\E¼æ'ÜI¼êÖ]N5{‹›i^nÔS¥,®«[õ^yè™×7¨nc´àgÎ^·-z^ù»”Eg›þtï`{³x´LNÝ–éeñÍ—!ùÂœz9Â”ÅêS¯Ûçé&S–7…)ëÝËÓ…yÌ4œµ¿MY®ÓåvãÛwXäÚo
+-ÖÁ§,.WŸ²äðäµYZ¯ ¯–M<‡ÊâÚT7–ï‹V|P¶ÈrÙ}É³g×ÜzŽP^Pñ–ò½Ó:§ªüYaÃ÷®{ÔòEmþô±±×VYH#=)“äkUO¯®²4öïªî½E5—°(üéECŠá½—>I)MYïŽV~WuQ[ÃëÌa—äm¹‡©Ë`ÈŒ Í‰°@ÿ²ÉˆFQ|îeÍDôÎeÌåkŸ/º²§5ë?ÑÙÍ»È|HG˜~&òÁkàôF<aÄ.vi>0®gNû9Ñ§Ã|NÔ2¦E'ôÅsa˜z‚&IÎ$!€ÂOè0A]qzÃpu“X$HU€‡äýñRp¸±„…´Ã`L£&@ƒ;ÐAAý¨‡ôýÞéßƒà‡›:ëb€ÛwØz+)í"[‡ÁD;§A˜ès@¥£ ¹þ‡£Î_§4ÉœuÂÈUçVåªrýó‰)Ó›ˆ»o
B†å›xÂfE2¼Êgxebà.v&ß¬ðžtíê¡8ƒ† ü±f.3<Q=·õÅ¦ÙóáQlÙD•ÆÙ4©{Ú|1k®nxVÒ3“ÅkÎ ÎÒyŸÏÃÌVŠ~V¶v;Y ŸhViŽgëEšXEòÄêy(QÏI¦Ä]¤ŠW³+Qe+H›«¼2éƒœ†”¸ÍÝ`÷bï´MK¶h‰tÔßN¤ßš*2ôàÀñÄBˆífžŒ<O¬Fž‹1ñ´huj4‹±r¥bL»žÉÃºÓeãóp
Ø‹š‹ËÒæbÊæ™†u\Ÿ×Š™^*|Ãµ¡±Ž3\/3(÷ujýýçý/¶–ÉÄO«ï¡ƒ.4ù;°š&Z|»ÆÙ¡þ Êh‡®¸™Ú­€¶¸Y³¬1>ŽÐ¢ÀoJ,<ÝàjJ1k*[<«>™oèÌúäõyC.éÂ6ßœ¦9ßÊ¦äÛ»ˆˆK¢«RAÄ?Bñ–æ²4oö1èhvÃd„­ow¢t„BtÒåtdäWÇèÖþdª1…˜âgcIQcMPc„Å‘š¦~­BfEËb@ô^·Kå¦Yh­z-+Sô{ÓÁbvÌ<ö&ßSo20'Ö›|a^ô‰”mAû?ôšü©näž†ã[¶òü~0«§8ŸÚ¥FjûB‚Ä…í²D{·ñ"Ëóû\ä·Ÿ
ñ¹½ [Lî7 ­Ÿ8ÙÁ<|_èŽøñcÑ<„Ûx•	þîòÇ†ÜOˆòÜV-g÷€×´g‘BëÃBH¤5h¿ù‰Ý[xJÐß¢yÐ¬Q%¥3io…4)z¡ÐÌÙ"Ê9º¸]VGâe1u[“˜mIœ–$‹°#1¾ß.Ð†af}Fnn%¡ÉEëi`¢ÏA;©¾,]¹Ãqû’5¡'Vc¦]Ã
Xæé±.Uø„ç·²›hÔöÁn‹¡.Õš¡Òÿãø‘ëÿÉ‚°J!ÊbÙ±OÀ¸žb^èÊ{¹e ¥.r™(sýrþ’p»  àÝÜ^Àù• Àgû‰Á€Óo3<P>6zÁYo8B6ºø›ðiø¶vö¶ö^oö÷ú›ý=OK>§1^aÅW4y[òÙàÌjÅwã6|Vƒ¼&ö{.ë=ó}0YîÝ"»½wç4Kédâi¸÷Ðßpïs’°w¢,8KDJº'1HXÏöþåÐ)c[(ÜÉF«>`îçƒ³€¡‚xács…g×ÿBJh’…§@jùÀjÜCÞŸTÎWW‡Äÿ ,$lõÏZáÑ¦oµ<kØ4¬/ž3_õÉ„ÐpôŽ¾'íþþ~G¥CÜÍ«v¹æ7/E-.Å„§[¼ó,›¤½¸ÌÓÜÉNw¾¢.^QCÃà”NGÙß¸N÷$Î8ðt§ÉÈ¬Ë}2š&Aˆfq}™ŒÈãŠþµìÒ¸m!J,ºýê+K;¼R7…³ôïavÞf[;ÕZš­iÊÚÚéy±oò®î^òï®ÞØ¨eÅô±èÈÜÄL?qWói W}Iø™ûuâÌJdNNEŸýýèõæÞ‹§;Ï–Iƒä·Š‹w‚5õÄ^¹XNÝzÒ†ðoÈO?‘S:JaAyŒÁa?ëa¤ JGáÃÞ'mÇ*áû.É8€ÃïÁæ£°UŠ·QÁ«	´‡ÔˆfS”Þ±ñkc&yÅIz“ÇFJ¹&fÕ#w/a„.ì[JÏ‚+G„@ÀH¶5~1òžf³ê{>fÕ§ñ`šöPD®TúPÄþŒ§Ù(ŒÀÝQ`3Â6¼	¹ä¡†²â‰;ìàsæ­à±J¦cÊèë°|& ,çf…R2c6nÊ}G3
”#*ÜGpú“°ÝZ¡“pîk6Mí3ôƒîCÀ %POõ9d]#ˆãº â–ÖÞ_ZN@O§ƒÜ‰™A½¸¾»¼O1
z¯$ü§½¹wp¸²wzŠK]ê‘½‚ŸšÄð3¢Áj2Ýá™ÄˆgâÅPáúgòã48Á¨.AÔÀÁíGùf¿xÑÔbõás}UóF@˜Äå¦(~—Ð‰Ox,&9óˆÀ¦§YV5›GÔeaúP1xgF&\àŽ-Ð¡iSZÞ€àj%0Iâ!ƒ¡vpÑÓö˜óVã`
­ºÓIWâ…—ºÚjtVbÕZ	E‘}24Üâè>sË”¬‹f6Ï)p+°[B‚Ù›ÑÐ Eí½ýí{/¶^Ã¯¯ÿ²ýoM…³¬Â%•	¬ä,]˜Š¬’þÐ‰“ÎÛµN·ÛÕW7É%q±$ÀØ	Þ3üiPX6T8i F;m(¨†šµ†ºíšî’‹Ë¶Ó1_{üµ»®®®'öÏŽ¨gå’­l¯7‡l¡d¯%®Õ\±’‘fÊ Ø@¸â;ŠdœK#fÛcÚUF{`\E#þÛIrÍÔ~o ¶·…ëÖÛa-€þmô‹3¹þÈ`¿¥À5§A"4u
ô1†ïb*‘;ýN_Ûa;	`Üëÿ“„qJFét§K ¢	}{ýKªžùçJwñ‰mX¬·)á3Æ+l!%AÈFµ“¥Ü²aâWIÓ'NBÄ~õŠ~â´£¥ ]Ab‹nå¦‘ýXšÑô.ðIþÒ¤ÃóÕŠˆûÇ¯|«õí!¢…;ŽóMnïÓaòU<Æ ë’õî*y
<_²…û”øóÝ½­íç‡Üdml€$Ç"£‰ZŽ‘Ø‹)Ø˜iÌ¦`¶¹ê^ ‰1<!ûm!DdãsJ~„+•o7à §)&=e$’¦p®…iæ]¨(ý\iœ¤È…ŽàÎ:&m¡|Y´Œg~2¾EÕåo¦f¢·!M:ôm˜ÆÍDÀL,ÖŒO gx©ª¼¨´í“ý\Ä?um~¯,·Ÿ€n”×3ÉM“r1·|R_¨ùGÖe%ú“Q ÒœJ'1Ðä´»1–$5<šE.mw¹b³DÝ:M¤úÙj9ƒè¼Ð\ÔßE6¢ÑGToçwN_§ûït$ð÷ P>}½¶~o6b÷Ih;u“ýÜˆÝg¥éêLÃ]º$O™áÂ÷	_ /cÄ¿ƒ"˜8Å\{pTáõÏo×Ñ£ã“¥.þŒÇè›¶ÇÒ¢8-¾DKgÆ÷ý4ÍÂÓ÷nF¯¦KÖFy©ö£Z‚÷Èâ°ð_<ÿ‹ñ‡ í×ÿÛ¿£Ý-š,¶wf„­àm8ôi:»RË2mÖŠnsûÆÆö:—ù‡‡~³~­£`LA’ÿŽê|AeÖü	üƒæG{[ýCT-¤ù’p’Å º§è3»AÚLƒ/L|é„žQÔK Úžiöc@@ãI7% i;f¬Óóº0‰Œ9âÂÿ¢ëÿ,§KŽœo6Ï—_O†@ƒÛ™Ö²GôáqÇÃiíï9Ø¹ó+{çßkõ1‹'%»×?Ÿ…f§r€³ÜeßÂÉuÉQ<D•‘t´§1#}x†øS ÌS¸¦)·_Aßwd/›,ÎUÃ¼aÿ¹ùP¶/]û4C‚Cï-´¸-ö°x¦3œÇ{Äø›Å¯D5éI€^W‚ä‘jµ¡4‘°5°Aó ò	È¿½ÏÊfTL8Àézâ‹Ì9‡ižÇïv¢ïƒA¶÷fÔFÔâß¦1‘šsC5¡"šï(® ØQt¸ÿóáÞ‹ßD_Š($³éÄ²	ž°ãJ¶¶Â3ÜÊ8§ÈLq·Ä˜<x°¶öu^z³ÚÆ	‘o:va>w@—/¦ã`‹ã:ºIÀ„ãöÊÿÚZ9[F‘Ð~¯sž€wâÃTù”<›®bñ',E¡r
!Y¹~ è¶”Åb†? §*¸6€ôfë¨[Â’3Áý"y‡CÁ¤q†ŽiÕ@&ý³)MP¾M#ä0Ã2lä¯€}rÀ;'¹²BžÆÉ€e;dôžÎƒÁ„ž¢ q…}š}oydºÛÍeû¦^”%u^ð³Þ¾úŠõÚ…“O}òa×˜£²µ{Ü’OáÊc|µLî­®®~LÖ«´(çvÖh>-ÙTw»ÝÏ”;K@æž—{àÌË>ü§oxÁ©íLlƒÙ˜Ú Ô@û€ä»º?ïSÔ¸æàŽ¤Äî(ˆÎ²s¦]%¦7Òq5, =@~FÀ‹ :æ$g6Pzdøþ (> $æéûÐ8Â¤ƒ†a˜wÄ(á˜êT@Æ˜K¤gZ¬=X¥!ã…iÙùþQ–„öÐ^Ã5íã‚”-“0:_yr‚0Bu³Ù‡´{FC;²hO­Aöx™vKêjõƒÄ""[Hm–4ŒÂŸáˆ®¬ÍÌ©
öN\Ù4þÁ:µ©"3Ž±Ö–½q%P?Ýè–‡üŠ¿~üIw—óâÌ¾†ÁØ±¥Ìú§ÙEÆdÄjÐEKêjŒH[÷ÚÖ¤û‹ò´šGE¡ZÑ¶ZãœXúÌ'íÑ^Óƒém’¨;óê0‡÷É°I˜Hƒs¶µ=´¨òkaG–'ÜYFÌ“vÉ'v^½éDÿà]›å
NÓs×%¼i"‡š´ûPÏMTê2’iü^ òF¯=À-ó^»2On‘GÚýiûq¾Í¶ÄL*5ýÚÔ‹;TA½‚4:p2{ëÈì­ìß»Î=åÅlMû|Æ¿Ó<¢±à,ýg®¼õcÈGÿ¦ê3I¦|P/£/ÔKñô{ïmN®+¾²bY¸ ÈŸÅ=sÊeØ¼	4:qË@Õb°"hÜ·:hÚÎ#»¢Zš–ã79ÓÒ½+¨7¯ü0±Z,5µS££/5BÑ·$…{ˆöRCßªÖœe±¨ÇËâIô}«99îX|ayèUúCãµØrÐË7%žÎ¹É2¸¯BPï^¢8´ÑÚ.XŠ;Ì·€r—Âá]]ˆñ- hõ¶"æÏÇ	úÚ”­ÅgÑ<è)~O6­«7>çÉ‹È“ÜèfñKäž6{j/˜à	p­­íCÀCßmoîlíYCôåÅŸ9ðÉg›—™òÚæÅ#´^Yš½âËåáEŸ?Í£jU¼ýÝ½Œ7<tfyñeÓ|òôæ¥)Uû‰^.Mõôr™Ag/·þ^.Þ:S¹¼ÉŸüäóÎßú»Î½!5šY/—æÜ–ª¥oLØ Õ—‹IM/ÔgÍºk¤¸W.“5·_?”§:_.³ina!Ð•@V¢¡T‹§–_.MöÈ³¦_=O¬HwT¦æ’Ð(è½”[I‰
ê'Éea6‚QûÊñ -FŒ>½x*fšå
/š9s†—5=xJó³„\4|\Á‹)ñAÃTWY-0‘¬h-yTÆ4;zÁêjt—
/<“fð]gmœãÌæCIÊ8_³’S|¤Q8ÆuL¦£4ðUÜ§@éÏ7ß	Y¿ï"¾ŠÆú£‘)Y·+ðvYøK:{ý	£ïþÎÎÂüð»ô6Ä`ÊYá'ýƒPu9–XÚÖÅ´œ\Ÿ0å¯ÏÛËÍ H„†±V¼'=	GaÆâZ½LB½º+ÖOö·ÿBž‡Ñ %ÌÍU<´i#7Ô–ì³3^ÐP6‹(7p>ð?Áÿ4³ñ®ì©ù-Õ…ž€>ÍâÎA0½'í¾¼Ð@j[° ßLœSÌ™ÓÁ4I`²}i³	—ÚF`ãßï—6ºaŠ{Í¶Zt¶±1{ot†.XŒP¦©PÄæk'w”Í°¾ªee…ìM²pÂåýªwã
cÞž$Á[u 	§Ûív±aSz-í5mÄhRÍØá<Ôk±ûÍš×`mÆžÈWKK bAJ„,>;C‹ÐP¸,+6œ
õáB[²ÓçÜ†ª-aA§/ñS‰¿6HkïSªî=}êhª–ODgkOˆncn±tÎèÃA R10GtðC3ÍÕhÃLÔa.ú0?…(QÃEâ†Y;kˆnBwã[Q~bqè9„Lè0õ¼ó€¼ë|­²ªê+Q- ®æùþîeÛƒ42vœ][ÏD¹EGþ(T¨<_‚|:MžÓª›)öGP8ïÜCË¾Uèï¡î`‘IEÚÊ²êâ¶ŽõÉöê¢ó€mžüÍš÷îy?¨ú«¶XmÃ&o£¯½óìkN©ƒŸƒ70q
í]=§äžçâ-¶.ðµÛñ¼îìÃÜ„qcÇÒækòËÒlËÕ${+S)IÛ03ÐxØc'ñ;üÛªjœEhî„NXv‹¸>"âªÿs¿|Ý¨7üS<ay$A42¾þŸQNFq*ââ¥jä!®2dq‰¡ç³8	u
H[”ÝÞÖN¯’ð™LN:ëxzðïªÐ¼§ƒ$Nh‚à¥×q_k{IB(ÏÃ4ƒùbIY(øî;øCÝ²/á¿Q|ó«÷^y71ÝhÑ5ÆO¥nŸˆÏ¢Ç-šQš¾A—°h©Ãgý}–\ùú¿ÇÊ\Ÿ85¥x7é¤I÷§aîãõ¥QžÂ—+Å·bˆ§á(8`j¼¯Äâ·t@ãÊÎç×¿þ8‡´¤úaån	[jKÉlÕF2«wèÀ’˜™›¥ò^Ld¯îåÌiøßEn»#zÒæ£ädÚâr`Éá«QÓòË‘‡‹Y—íï5jh%…¡&SDä-ò$­®‘)ëù:-äRå•e_,1hžÃÆdp[rá*têÆg½ZZmSÿ†Ã1zõ°ø¡2¦†UcÐecsl¬©>kãŒ\€_ò=&Õ	ÑU¼ª©¢Ã´šjcFÃãRé~Ö£æGˆ$êÕ5XÃ8y¹-b¬á7ç*ò_À¤m°L¹J-Ìs7¦&³ÑÃ
I§'Ïq
iè)qK0<ìÒÚ¥Yrýë`:¢åü)Ùû‹©þN4„
ÅpžéûÖ(pšl¸L--û®ÐÏÙ7½Jq½ö»ÜZË6Î¹5²î¸6¡ŸùúÔ8ßûóªV¥{¢X‡p3Š uFV#kuÙX0ykÝ>ñ‡9¡ÊMÉc>[ª„7HžwW´)fÉÃw»Ø'±|¶z™½GØÄºL mÜñ+îZÚf§ÜÖðÂñÐÊç ÿ³3Ä\“w/ÅÀW¯ï^B¶\“bùé4Ô,Ìî{Z_q1™"
Ik1 k˜íì—Çx„CÏ³*ÜâXŒÀþæž¶&Q<züÙª…éßÂc,öÈèÞ\õÊ Pÿ&( €…c†<Ägåjã	²Ré±€‹Wº!Ô~vDŒHàÇ­]yià‚Œó*Oª§€Iùƒn#ôµzÐ2‰K	öÝI©kí¢¢8)üÐV˜)Ûƒ‚Ã_¯E„,2ÉÙýÑ}¬?°‡{Vµ¦Æ°mÛ`}z;d×&m¥(bWä‰²0âl¶.mžIå#Õ8¿_Ó·œTÜX•¸Œ’§kCÿV¾&.58””ç÷³Öšà5T__7—ÇŠ¸Ê¢vºD†a:‰£ë_ß£vè»â°'s’Ë”Ôªù=d1¸¥ÞxhtKŸ¼ˆÌ2†²„[ÎK™Í‡ûb:N¡®_È1q³ÇUïÕÁ÷*žžîxJv8°oKåá¨
õF÷O<ÜŠŠ@„
Ãp:n,·Ôz«º›’Âq!=rÁÀ	ÙÓ¨ä¥¼˜<ãKÉâ¹ÌÄ4œ0P¦‡qV— N¼sbíêÂý°ƒ¹gbžùöÈ–ÄxRò²æ~.ôÛS}Šš²<U¥’¼NãU?ÃÉ&ÆäŠ oQ¹èpèíi|Íˆ¿$ùÈø;WË.Õÿ®tØßß½»}xØ¶}è×‰§_´¯¥—zô]†ìü¿\¸,ð†ÅŒ* ½BŠ9äûö8H‚³ÌçÅË˜ÄË„dQäK#!‰¬q€÷<COÓ˜"@>;7‘mŠòVölnÒPÍ#Á)CVgbØ<“›ou1™o/HÒýkb²–¯“*•7‘`ÎÇ¥ç@½è8Cpä~WÅ†H·jÎˆDû£iš»Ð¬:]h|ŒP,ÈÊ#Û9d‹ÔÊ½U‹\Y>ËëÞ
Êr)‹ûöÈ9Õ¹¸Ej};=o“}Íýx…<²f¥ÆAZð¾"bŽ2yFYÂ@1h$ÀÁÕ=GsªšJEßP=Í‚ôT¤!€†‚ks‰D¼ •îE#à<ï”œ”OÃ†ùQÕÂ—ãùEÎË"Bã4±¼Ÿ9NÃ@8ÍBà,€SãE†«Qü›KÁ_ågæm&ëi ;C/;YO[¹T³Ê¬ûHÑõ<[>Ör¹ãxü¨—É€¶Ð–FMƒi’Æ	Hæ2 ñ»ÀQ«^†šfÏñZ¹Wà%ØÖõ£ôÏ°$EßÏ Ô¸—è4õ²‡l$´²^½=þæðökõ¥ó¥µ}1MÃDÞÎ–ŸÅ«¿gÇÃÐwðKÅ{{x‰‹M=H#|f¤@æç²zágÎº°¾„é‹x~ÂÖ¾üý¶ªÅ“`cðï®oè• V­qÖy ¤;`#N)èF×àß(6lÞZ×¼A/¿ü>/ës%¢†‡¬Çy:ø™£ÛëäFÄâÉ}{fqâ—)xÕkî:Øqz¶3lÚ(Œàé[Ô¦L-æjÍšÇ—opIú]æ´ò…8o·28Ž^2æ¹ JQü6ÆÙ™;*:•Ú(ïÐÅªüE_ÖmËb¶0,í¾¤knï,:ßftl!TÌ›5³Æ´¿ÚcÀŸü(4Âù·ä¢Kõø·Ÿ_ìíWèEr®H0JÏ3˜QömÄq}p8•×º	øž,gÍZ+—…‚U¥¹`–ŠÔµVUÏsÖ~P½•{qÈŒ®ˆàv(›áê'V«šÓ\?ìÒìj¬’¢UÒÃš‹k©dp™"¹•É‚ß:üqJ“ÜÐûÞºÇS€ÛèH›vE¶g"–9'ª15À*ã šòü)yˆ¨qpývêè/@nlòc‘IrýËÛÐ„¥\ô<6@F.ÙÒÎì]Y¿aW•l^”Ú:e «Kn6™;È´2Ì¶Çeq%¤ÕÍù`ŠXd÷¡í˜»Ö)0ÒÈ'³êÏ9§/äQ¹"xeã«N3¹;î€ôsšFÒø$aéÑ¦”cøSòIpâ« ðqðå$š•’aþwúÝÊ$'*lTE³¿¦—.¢õ±IQÍVÒŒŒO´q»6¡>Û‡!%Û|ùÕõOGß`3Àmu„ ¨1ÅnÓ§!‹®ŒÑ93ò‡¯ÿ§Ë\‘ÿNâÑõ¯Ë±[ˆVåù12¤“quÎ;Ç÷x ¿ÚœÔ„eNŸ,¡ ìFgb·MÏ^ÞÙŠ§¥¼ÃêëŒŽÛZ~	©RK@DÆ‰Ú¤U³|Ê™ñ4È,œ8rÞÀP›r#ˆ\FÑÌOUVÐ“vÉÌEÎÂ¤™ÒK žpÜÅTÍÀÊ<`º¹Po¾Ó·èÊ€‰Ly§g¶Æ ÒÏ¾Ny«_Õ˜'§)ó}ÅUƒù¶ ÇôÈ_§· -Å®`Ý‰O¾XÄÍSÌÈéÄ2êè<NqèUUë¯B'j”AÃçè4x¶°µ^¶ÆÚÆ)šùf³	³µ@’tÌûeˆžu˜ê®ñþðÔ`7RÆ’G¼M'aDjÌGë[fB«÷·5‰.µ1*äõbÔúvë©±Û* ë}tU/£ï€‚Ë˜ÝÂ|j£Ö]|œ «ˆdORO–#'5Œb:DK"h´.Ìƒ3RÐ'àäâ&Õ/ð pppq\@×?wF	$%ç\KÇ'‰—âI2à3xžñ°öî—;\—òƒè–øB$Ze±ŽrYP.lž†?0çpcœÓp¤MEƒ`’=nu'ÃSÝïÒ&‰8-šJ¥ü3°Œ—ìØuâ×0LéÉ(>®®¼^¹ŽÓuü¦FèR¾¨É`f¡kŸ}‘¶ l2ŽÇÀO$ä9l	sƒ‡Qeï"wU‘%€®Ï‚1¦zÓdÀS«P@Á (ìÏ$¢Ð×ˆâ¤F8)ø£Òe³|Ù)™¦ÓëŸ“0NçÆ|»>r22›.ê}Ê-ÔòRlrÌZó®œ_¶ú/úL¡r9 E°ì!MÏƒ¡ra`¯fcGH]W}ïêÛ0ådƒK6ð¦hR$¹GÞãþA¥O¥éi¸W´«Å‘º»ïaDÿ G[	=Ûƒö¹˜‰˜ylÖ:õÒÖƒ{>,Œù%ØËœžý6M&žx™N.x:¹Œ‹4I·˜áåˆ…tˆg	rWæ(ãb³âË¨œ”!¨p¸´Š·Õw­Û&ÃÞ)¡€ô1xŠ\O$÷íQ€>y¿3l3*ÕÁitÍo-mt1ÃõÞ‡§f6îµsèÿ¸6'½ð/˜†ÜW~¯¯äÅq4¶Ý­@ŒÐ‡¸`ÅZæ€Ó“Œf´¯Ö|!kä04ÜŸUÃrYA¨~q™‹†Åö†ô jÎÔ$Ú£ÊßWü*T€'Fýêv™éÄý6ÈbÖ	t§Ÿ$4e²Qàß¸³P³ªµ 79‚ûèLëN²iL°‹èÛ SE1Iƒ)¾V Yg_dxãñÊ ’+qÄ#WÅ†ùõùž "¬¢!L&í‘ýÏ–ÉŸ÷ŸáüÿöŒ´wéE—<Ø}¢³üô\¨·¿ºD,¹¤0Þéfk\§^?e·¦md?kµ˜õÏ‹ØÚây“€Ž¸¨Î¬íG¹1	€o‹„¢ß^ñ%&×¿b¢ •hÃb7m²Èµ!è4ÈvÔÛ×Fý«åÉÜe`å£ÆæeNe6/.&áºÜþ¢{	/5/V‹Øšç%°·£éhd­_(·y3·BEÜAÒ)¢D»‰ECÓ
¯ ý…º;‡æ–½•Ý¼x„Ý·®ŸG#;Lj œGm¶t<?[¢{l“e5ƒœZ1ç-¿{¨3ñU5*ì[æ_(¿ç­×ì+r¢Ž„êiÅ8ì]¦z(”!µ^,fƒ]-aÐX5°Çÿ´º¶vÿÞê+‡ ]¤ÑD1u±üT)«8‹(úU]·6ãN-®× zŒÈ]gbŒ6š„gç:z¯H˜?¥®¬UO·øe]æ\ïç\Öñ×…Ï$R{5V°øò3ƒ†rîì¾|Þ¿þÏëÿØ#[ÛäùÞ³zÚí¹Æ0‚{×aZ‚ãnuÎð_€þö?­÷è#úzmr±Ì@XQ†Ÿ—^‘cÌqÆæÛA¾µ·öprñÿóŠ€<³÷I@o‡	Ð)?än¿ç±‘m¦P¦(F
Ÿbbëk‚¨À Ñæ*Ÿ™it2x,k¶E£#!óÙ!š3›ëËÓ£öBÅµ®OIj||×øÆŠêÄ@WhtÇCí=úãªˆ±Åà™ÝPzo yøÏKœ’ý¸³÷[QÜÉ¿2521f%£ˆds€ídq'KÈi;éï¹Ê6®‹Ž>’c=bïèÌ&N2	`Ëá*¿œ¸å¸&1ðíúpöU¦"—Ç*"/cXô4ÃX¢{Ìel’èdºÑØµ¯Ó&qòI0î¼EîUàŒ¾“[/¥Sš¢Åöc¶Ó0í5”dœBIŽRÖu¸©@]ëzI±ˆì*ž-+t¢VW])X1¨¯Jd¡ønR1ì×å°T‡µîÞ<ýûš)è'P¾Æ‚?Òä5Ñ„]_üÞöáÑ6ùîå‹£íòÿ~!ë«ëÝÔs‘taôCÚä…Ÿmš½Òqú°îù*Çð`•íRð¢ç5~žÃŒ¬)¹þßY¦uŒr‰†µÈòC-ZÁ¤+ðÐBo³]vnã Þ‹à«“;Êƒ J¦Éˆÿ1Š”ý©uKÔú7ÄûªÀP†s×Hµ·÷Ê{G€ÿM,™E‰®¶ŽÕWNÔY¤8:œžŒÃ, ï‡¸WW5H*8ùt\ýäAûŽjF :½bb‰¡@ªO‚§!‡Ê×"GÄ‘9U¥P=ön·+/PíkÖÒÖØ}7f§©nÇJÓ}yð|¶…»0û~²‹ô9nçsD¤½7aÆ-£¥êÞV6‘á½‰¬ÿf
dô	n§‘óç¹”!˜êöBj©FMõ%YM™X0­Ï:†½ÔyUQ¯BÀ6È7‡ôm ¸MƒXS›UÕ¤Ô‘ó7+ˆ}mä»de¹|x|Ÿ«D
õ‡¢'IMÉx0]DSk&Á¨GZèFéÈ'pl²ý§ Yœg¨3q“A«/–ž§×æ¾³§˜ÞÈFZ²Ã†0>B#ŽŒQHŒ×+œ…ì'%ÅÞ2aåà¦Dm¦Iûæh1œ†&§O.8ˆcîð+\ê‘”+,¾žýeÉþ¦TáÍšÛÈA@P¼RÍadôoŒoáÖR›Ê5«©yÖeÑÍhÂ ¶À}QrS4R)/ðÝ¯‰F÷Nõ|çÅ_aû1„ñË—ÄÞÎØ!ó9È¢ŽG	Ù1Ãù&ZZ²IäáÓiÄ­‚Ã¸&÷‡ã0z*Uø[¼“F÷øRü!ÏháSþË(Òa>¾T?—5WÜS$A:™z?šÒQuš›•ßÙ´¥©¤ƒó`Œs À¾OÃôq¥,Y\	þÛW*ªõ¶ÀcMþ—©®ÇÂ·˜Ž´[ÿ´ü	7LS4Ù›$1Ý¹ªo½Ïxø\ ôÛKùQÞUýaxHsè<¼Ô¬³¯¹áiUáñ„Mä08›ÂÝ¼þUÀ_Õhè½˜Õ'B4ŒY*e'ƒ?î½à¨&-ì™áZ¿Â©åJGi1Ifž%<iTW™óÄÆWÚ…ùš>H‹ÔtÎu|Ý2Š¸®Ø0ÍÉ¯ç<;ÄÚcªŠd×4B€3)ã®/\ìÿ7[ñ»H6ßÒ…JâùñàŒi"ÀIÇ]é1³Ø_ýf;$	EÞ»¢üw¥¶FÕŠÍ¼è€8Å­ØBhöÝ	:£TØhe…t:æœÄ£ öé÷*J	Ãoðûï˜.û–VÔ,ka%4ëM;<Ú_vìâ í‘Mé3cLD™ðÓKøgŸãMþ½À˜=Ò÷Hš¡D½L²	x€F"-òS ·q8Ä†ûn(}B]Ôi²u“8É’T†QõËhkV@“c–,Mºø$¦º\¾ü»«žm{¸Å
·U=NšÎ‚d¼ŒŒòaññŒ #fÀ"µ[œÑm„ÔS–Œ5Û–¿‘[~ÓBÖ7ëÙÁÚöÇÁþ[ßrÚÂ~\þ0KÉgí¦©Ôù¡ô…ÒwÙö»›…°Ï	Ó”©øÀ­Á(bˆoÛ¬…<bœÁQH#H_¨û šmÏ‚l+„#¢x m˜zÓMJ˜9IŸ%^™öÍë£°q‡WÍžÉçlEÎ*è=‚+°…“(Zý´ŽVW{ì9'Ìër}Þ±¥ö0<=Ech´K³ó.=IÛ„ã—ÀàwHZ~ªµÜ¢ïÓ¼å Gí¢»Ò„¾J~OÿY¿¿´„ùx/b­!àËí¼³%²AÖH¯è«^I{>¡ï¢âw)?›]Ú² ÂëDNG3Ò:~Åî‹˜ºz¯ºÀbmÓÁy»MÙ•TvLü{t/„)…;4Ž³˜Zñÿ.H:À˜Uçukår“— 	©R`Çþ$ˆ ÊÊÆÌ[	M^¹¿Œ9™Ñ€òÔ|h|XôPY«¶_ôµÑ=ÆQ_•3.Š=gUƒdïä{¾çIž1mÊ¦:íNEF46û?Uš³Ð^±Gk1CQg¸T›É0Èh8JIQ—±ZûÆ%á1Éâ©@^=‘dc¹ó³¹_~'…Þ>žpÔ´–‘ÇR37]ñkŽp5SóËÁO2Æú“rLÒ(È‘÷Ï˜1B#ý¢´®nQ‘JôÔ.v+£;&X&ÒWOÃqíÐÞÒQœÐ^ì·ú5¬cµÚíc?›Â¤qâÄºìklp_ÝUÖ!LÓoÜ!u Ø¤ûKr/¥ê,„›¶ÍC¥Í×^m)m ¹ù4Z[U[=X­Ø ­ÿŽâŒŽp¿Äñý¾hU¯F¸l¼$Œòä÷&æ_ìv†JÂM¼aßÊ—oª^=T¿ä‰>ŸÄÀ:Ð¨§Ã·b JAÆ¹*‘»Ýè26ŽAðÓ­¿‘öf<’Aˆ¯/8+¥Ú¸ÀAUdµNñGçÝÉ4=W2°ð¥ï{9ô¨qêrÌõójâ5\Ö!ƒFó <?ý¤ÁyWÈÜ÷d”¸ÑEÿwÖèÚêðì3<mûýó8ªv0ÁïX‡Á˜‚JöúÖáE¥íàN¾.šãßŽ9¼ŒÂ¬Ò	«Îr¢
™ô³öêR7‹_¢{“¦È*ü¡"Ih§ qí5ÕÈ¾‡>)a4't¤ÌX\¢–ošüsŠ¶o“5T³$|‹SSêh¯¾pçé*KÜ ÆÝïã0jcø¼ªCAOÓ›,yLg§ú¼áERå–øÚk“b¦µ)”ÔÛu«õ5k×f,×dÇ`œÁ3%îúg89>ÐP%ºbK«QGK’×#u*X©„°W!Œr•#Ëß•È»Vs/kX+}Ç¨š´*÷Ð+‰½ÄŽx1ÿ+ÿW0ÒG2†y™WÝe&¯8O,%Š9,V™/3LQòÆ@?<ƒ9s³ŠÑ¥1c–‡œpE-—’ì<8SB9²o8óómîB*1ò%÷›Ôøï¤+#­%ì³K‡Ãê÷ÚA·µ÷]´5mC³¥|g¤‘Ëí`Z¾Tü«X-§‰ÁÐ"pˆÑ”‰”ÙIüNY‘HGN$\Ø–xI¨Û•‰ ×çñ»¹v;Ž¦Ã m—B»ZC¹=ØY~7æîHB©û*¸iéBa ­]Q*p¼Á$ûî$ÔcÓÈÉ3þ¬4ø“fwÙý‘F¦*·‡Ð–ÚiÆá×@>EI{`G†\VGnTl˜ +X¾úJÝÀü³Xmþ‘OJ¾°G=G©ù©jtŒ<µW²ò#LDR$±ú\ûUâþƒ½çÛ‡]Àß¯7÷v·6wúÏsÐ2U>ì?ï¿Þíìl¾|ÞÏ/çîõ/€<Óz0Åg ùÅDÆþfºœ(cdˆž™T\{ñ}®Òý&¤/}3¦c?±ßdT!¡´

&ó‡ÇJò'ñ RRhÉÞg•µ»¾*eã…pY ØWÿµ‰‰Ý2·+ UÆ©õM[þÒ­X¯òm>ÉJU>ƒåâO#à•3ã¯q•_¹†rxÔ¾}~«r|CÄÃ8ë£N„F0ýR&‘àgÑóSf2Åi	W²7)Ÿ¥´ªÆ‹1Š”5¦]…OÜ)•7’Ê¨ªÉÊ¥Ö$QîÐƒ0UÌ•Ú‡|XôQBM·ÛU(Ù ciÿ^åŒRþsÑ5"»Îö`³ÿ|ûÅVÿ`gïusoûp¹Ü…ei*ÆõÔùs(ö°°[¢ûq~ÓÊxŠªûérùÌs5jÞ.Þùÿ  ÿÿ ygzxœì]o7ò=¿‚YÔz%[‰›Ö±}Èå£îÚ¶¯è!(z—–Ø¬–[’²¤
ú1Å=}ÈÓ¡¿Àì†ä~p¿W²}ÈáJÎrÅg†Ã™á,bá9ÃBöœ—œ3„áŸœá€þ‚9žò	ŠðOI(ÙÀÙA‘Üé?½‡ ­áïúé=øë±PH4Á¡—‹ˆqùrá‘ ¡^£•/ù2~J |,ñ930ø’’pâŸ²¹LqÔëqÞK z98ø ñ1;éoÎwlJÏÐ	gS&P5,ÒÂ¿ƒedcÏI@.YHì'x‘9yýCîWºÈ~{A¯å5À§X=Zt|Ã8ÖÌ{½þíú_(yÿ`ÅõËéï3ˆ3PJJMÈºÖS2¦€ªˆ7£ð!•Öøs™ÁÙH	¯žy˜Y£^€Ô¨×áõGšaJ,¯C ¡4ðqŽyE§Ö€ïqÀ8p£Þéƒ¾t¥^ÁkÔó™3‰ƒlœ§^êw%|oøõodšÇù†{ðÒ|–*ìI¢°z´Qäô]±î÷µâfª8güƒ˜¢ôð‡¿Ÿý0˜IˆÁO‚…ï${§êÙúoþ‚±ypõæ]Hæ½ttñGE$ôcüézl;¹7zÍKJÎ’[¶p’ù9çœJòŠ¤i–÷	Êw¾wV°ò'½þ@²×goÎ$§áz"
(s§ÿv÷Çõ`ˆÅûÒô©9%–×ÿæ”!¢y¨aŽ‘1›"1óˆì¾²6ð¬:©½A–ÞõÀ
õSKR²b1jŽx:]“íâDÎxˆW}z…¼ ¡ÌÅ‘#"ìwé~áÇÓ­†Ÿ¡oö	GŸ×ñË"Øe@Hýq= © Ÿ9›«gÃT¸ðp@ÿ'P~z¹t/ˆœ¢1ŽÜÇé|yÖƒþddÏ%ÉBº£E€ÀœI÷‚>Zì} ¹’Ž'éxAÜ¯vwBÕ:dO7œŒrÓG¥ÙÅÔžc¿4Ç×DHe¹
èÏ3êÛæ>Ä¡G(Èß7öT8±ñ&z7«¡<=‘Å›aŽ9‡3)Yhfás°¥ŽV¥Ãim*Š.'¢D>q×¨ÄÂ¡‹±K¦„ãÀw¿ØÝEvEøõò	¼Ô|™O ¡%Ÿháî£hé>BœÍBŸøZ~bÛa®™Éq(¨¤,tq „Ç‚—.48+AmtÏÀÎ¥É‰é›‡Ã>ôr´Úûr†¹°ðøe²Y¬½©Ùs8Ô¿[œ6¼MÞÄœ·vÅ·×¿ƒYðàä¢¿¼FÏ1÷EÃ&sê#õGmáî!1=Èº#Œ­îãÊma£‹€§À|ÃëÆÕ5ÿÅz¹âHøý(Çïd—œƒJº¾öFË´þ1`Ÿ-‡ƒIlÎ©¦#`Þ‡Â¦PÇ¹€­'ÐË€Œ¯?^šWó<ÿkŒŒÑÔ˜3 !É”³0YõZ%V:Wôeiƒœ§DÉX<`¤\çç(QÜÀÃ<çœcl<Øê/lòœæ}BŠ€§0Çþ¦Š`|¡ð”ÍólÏ@Ò'¢fY_ThÃéƒL!¤ZEB9ô^Ññ{£þ€“( 2zŽ>„wœ~›Öü¯ˆ;1ó[
ü”èX< ¡:x?i['Z»¼Ó5ü_<uoZÄÁ}-û€W“„F>É›å=é"w½;y•(×Â}RÖ;)„Ò ²Räœ‹JÎÁàJ•PŽž‘|bäXš¬FyÜD€C,°Åˆ‚ÛÕæ4'Øä%•1è¨ˆ^‘¢®œÌ½IAò¾,˜‡r	J4ØG’EîÞp„\í¯jN,õ‹Âá^À™ú£¹÷Ã94Œf² )—Q,Š"N­}áGÎ_gÂÛ4´gA›ÙÉ–Á`PÄt…ƒ'47Î	Ÿ® ¢€XbczDg§‘géè€7=&r õ‹ÐCçîå¼û( ý„¨„!Áh [bj*Õyd©ó".™7éF0]§ƒ/muŽG*6É‡ñ!‹ÂfRYÝqêåT4Õå`Um%¯(Ê˜Íqï•NüuáôK Àl‡Ë.<_ÜŒÛÉÙ È”V•ÔÁ\·†w%äÍ¸`ÜU±i^RE+~È"…0f½xãsæ+¦‰$¯*PïëÓ½}pINOžô‡¤“‚pŽÓÌ¬êv„„I,@èUÃQÔ:t«ªg2ƒÝ5ïÌÿ§âÝ‚âÁ	ÿ*·•'‰˜œcƒã ¡ºj`g’7÷7‡!~õ_ÔdNÆ0QwM>µÆoq^ý©ÂÛ©/fþ ùŸ:ZÍBúóŒ$·\ºc®º
Ã³©?e"êu^òëâzžø·2¾ÂúeŸ½NU+»e¯ý_€¯™O,ß–Ÿ¹ï*ëzÀó„ú°û‘Æa|úñ),Gw•»aH8–ttt„vÑ_r¬/Çœ{±[›Ë/ÝGB®,E¨»Á¸2/€Ðw$œÌ¦ˆÇ©ÁÜý+úY]Î^Ruñ5Ø2ñV ¶¢	[¦xáÎÝ©¦“R.öœp	sƒâÂ¾UdÁ“²ÞV²>Éñƒï`eÜCµ;LNå’Yª8´n'¸Fªn"Ì‰GÍ²}ÐB¡^ª_ÂK
"-aë9Ïc 2súmÜ)hq4Ê>U°JÞJ|
;«/ÄeŽT@¨%±SÂû“å„`¿°%û:|’˜X÷¢¼54Ê·{»ÑâÇÍÓÓ%B)¼Ê¨ …¹]ßc<N¸Ã¡œl˜Dh›C&Wø›CÂ‘xF‰.6µM€sœdqn†E×l¾†S°qp€\ÿzEÔ{õâûþ68tEÁæ€Ïâ#nS8UC°%»¸ºÊtŽÓ"ƒ`É¹-•»¥*ùišU¾P5ùÍ—×;ÜD§Ð”†`Îß‚÷ö f!±ßžÌWoK¶AÓU6hÌ_Úô‚i‹ã.QüP°`)÷X¦Ä§³iÅV•Bl¾ƒ¨¿°êò-®8BG(«fxZ;n	>(õà¿sþOè7Ÿ½àJ ¸2BÃ}«Þõúès´×DØ?SW·½ÑrvT>^b7Á!Á7]
@ôþÁJ-fí>Xi"Õü¸~_MÅ8ó|8*uxÔÿ Ãl®‡kGÊp´Ï•pOïU æê2ŠÎ§Ê÷È¸ÉªÖ
fÅžrt_ûjÉªŸJQý
´eý¾ì ÇœH•vå…ÕD¨æ¤·þéA½ŸAàüg<«É)·ŒÙ"uâs&4žÑ~-T)L3­z§g<> -ÿ»Ø¥_mZê®3rÀÊÉªýµ3;US3«àæÞK}~íRÙ^N3ÓÀhÕ”Ä9G£ƒ:wÕ‚ 4Pb‰f<
,[÷Ÿ”ÜÅv áÍ5j
nº
º‰âA«V/üeë¬ÇP7Ø¿Jÿ¸jcê	@£Ïmk•ªÊ TÖetÖNÜF|â<nM<ØÒr˜TC©.,ÝšÔÔ[m¦µ–”æ5 å›B#>=B¾Ýi¢³¤2PY{©tQGcWfÕcheUz¯º…Tmª‹ÅÓDÇµ¿[l*…oF­s¬I1¥Å“Ö)­ë¯-&Ž'³*”×­3&…ËÛÌÖ|à€F
h×ÎKVx–É‰2õóV¨a®‚†ª:ë;³ªq>ä¶·z]¶µ´å÷;où¤˜|ËýÞÞ¾ÙU0y#ûm;MªV_‚Q¨	Ýì7ªQ&w¯¤ê6 ‚‰KnÊö“ºl×Q¥à^
Uºý£aÏ‹3Jó['T$µP´Ã§šÉO¾ýJI¢Ù…S-õ™[GjïÍDÙÆ­¶"þ6¿MµTÈœ¶R×ÂÝ6Ibyý{&Nˆ°îV–­V¹ø¤YmA€jÅ”t*Ò©âq4À©?¸­¾µ±GJòëÌg<#«ºÜ§5ãhÈá³€pùœr/ IíÊþºX¯R‚j<Tk\£PZíböÎ¶~…É’ef²‹;¤jÄ”Ö&ŸülP¶ÁÚì¯‡nuuY-œÌ&I¥M.d4ÝòßFÝ'ºJ×||u7«(¸±{|ƒƒ½îÛOùNh÷òWêÃ/wQ¤_æ² u×£s÷’Ê_bÕv¿Í²r;(«ÐÏI±ÃêÒ[g
‘õMbVÈÚ”¬Ý’Ìã™G‰ÊÓÖ%…¬J]Ú- ]Èð¹’AMÓniìP™oú*©–K<Ù…ÕiL»ÑKÔËVØoå‰ ¶èTU&€~Ç%]þçíë×<(|Íd}¸”Þµ;m|Gˆ¨ÃÛ¦:U§»":ûêâ$ßm©s³ikÃÍv-ÙŒ´b)€É`¢äBZ{´«òÚíwÛ Éþ›0X*òþfÂZ*rÃIåRÕq5OÈt*‹Í‰sj¯ ìVþî°®éë	Ñ>ÈZúé¢þ†¾“ÎÝÏñ¶È,‚8ƒœà¥²¯æôm_¯iöK{Ì–ÁX×0]Á„ì6²‹‘éÂ{Ÿ
UÖá­2Žv³¢ÞDSÛÁºÈvU>:1¢_uxv|€JV»£|þ²9Èå^´ó¨øùi]«¼ÍÚº1^¾Ý`x“ISN	e	ÊBf•¼k¹(Ë—,ºÉ¼—¼©qŽn†ëP$ILðé¼`výÑÇ;V¦!DéQ¬.÷o?:­sŸð'KhÙãüÅñ3à•Ð%m!|¹Y”_óKemV…æ—µ€UIF±(Y×°5•È¥5›Ùûô&^ßû   ÿÿ ’…<–