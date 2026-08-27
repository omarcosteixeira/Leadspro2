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
                    type="text"
                    required
                    value={editFormData.nome}
                    onChange={(e) =>
                      setEditFormData({ ...editFormData, nome: e.target.value })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
       xœì½ÛrÉ‘(øÎ¯ÔöôTu
 /-©SÀF$%@jw!™U™ ²™•YÊÌÂE˜íëÙ‡ó°çyí¬vÖlLc&³c6v^æqø'ûçÖ=î·¼ R¤ÔiÝDed\=<<Ü=<Ü	áÏúJŸ=¹wXÏ:M¶S!=	ÆQB&IP/ƒi´Ñ'Ùä)£‹rù¢ ÇYZ.³$d)E”Ñò£ÕU2/¯u|5r%”‹Èw¾Wh‹Þ¾Äél^z«,/gÐ5ìBÇû=~?ó(ô~<’y´q…qù,Ë§ÛAJÞÇko‰,Ý:Ò(ÔzdÃ?LBŠ¨ÜÑ*í^Uä#d0èÍ÷+3Š~òù4(÷Oá¥Ê ?‰ÊJ¯ªøuÏûÁ?FmÊÏ—çIBfËÉìrùþàÉ³yFáòEBÆYF9ÿÃ1à>` Cˆ)tt2/†Ù¼Lâ4ZNqâY/1†.#ÊøfnÅí—€Á;ËÓ NHwo6‰³4H|p½"GX½“}ÈJ³i˜J;5$&jþõ`&!³$˜D§€8Q3vMgIö'8êÁ$›~EØ¼µÿì³búNfÇ_òB—…èüL_ÉçÅÈ—ÿñïdo–åå<Ã Œþ¢è™Î§QžéÝùÒ°Õíá_/ÝýB1vkžÙ—ÌÐN°ƒ_âÒNýŒ«äóâê~ž…órql-@ö˜øÑÕ‡q3ÖÌ—†s¼[6Ö‘  Azùu"Ÿ,ŸŸÆeäÃB?l¯`§øý<â¸P¦Á¬ÛáTnÖ³Y	âù]n\Í®Å¤Ï®«fZ™ù!„ˆÆjó—íõ|åÖW
~™Ëê šFE™Ñj‚÷ñK[”¢_?ïä3ïQþñÏY˜ýE9ìY”ÇÐ‡/)y·~ÆIòyqòETØ“ì$þ¢h9UýøÒPSëÚÏèI>/zâì$ŒÈnz’GEñ—¥TI%zò¥¡©Ñ¹Ÿ•|~-Ú‹ „~2Oþ²¤”i¨°31öåKÃS«{?c*ùÌ¢S”ó‚„Ù
Ò0ƒO¯(h“_"²^ý­ë&„¢Í[g?‚>¥PÁñ«^y`ß…¸aÖy¢½,TÉV–žEycê÷BU¼üøÏ@oJbÞ›¶P•P2V0O¡[Ú‹ªÄF‰‹ë,ÈAD™T×ƒFŒ7ÑJ½ZùŽ2Qd</Ë,-Èw+6®"ñÑQõ8‰.ÈOó¢Œ/—	ÈI0[~@f% .Ç·ÒDÕ5 €ž^²&=#c»ûìÃT 	I<ù $êÈük(Ân4áEÉÞ,J»ÇARD½ÇUÙ1sœž²uSX’Þì×¾õ¬H.cm…%.¢ilSãØª”€"§ ôP&á2×(A™i#J,O@æÈ>>\1M¢$È=¨ÂÀ¼øôóñ4ö²a\ã$
7®’,¢w-J¶R”–ý Š¦üÂ„Sq
Tå|yº#¡¶i±<A
–S^’+z>ÌfÁ$./f¶ƒ¯)Ù$X¸g€Bl42äï9%ÐRðñ_>þwXé.Dª&Â³–×W—×Ó¬LÆkïZ¬ùõßže1šÍ^ Ñ=‰(ùÈRÄ.“«‚À7ÕS\uY‰Uæ@ËÇ—ÊžDAHç—W…;i™_nÒl‘ú'Òé˜™é±“›Kù²KÑmÀä íÂS6úbãêáHxƒã'©ÛÒNe<ËÈÆÆéŒƒ¦­§CƒÀ1-N:“®Ñ­žE¦Îc@‘óA† æBLÍë<1ËIë·>Ávú¤ónœé‡ŽA˜T¿4Ò4ÎJØPŽã“+ùÓ@>Íøúd‚›Ï¤Ä9[|8²6ûÀ¢b0V®Qž—ƒ¸ @7q¼G«oa	a÷¬Zgq~ÍVÃÌMÀZy²ð"™ý:h¹P^Xy«KG2Ýeÿt{Sl9D+Øž
S`bŠÙ(öl¸s«¦“,ßaªô÷·ŠÅD?€íÂúé@ ŒYïŽù(äˆ!kœN’yÝñ {äÛoÉRœS‡O¡ªÝ°œü£ŽE=·7bböƒKÜ  3vÿØY^Ò'qxáaVLhAy	¸#(@þN½’(=)Oßšìð˜ó<u8 e™šT.9Þ÷!ÔA­ç8J½	ò÷À¢+šî“Ä±²Â†•»+íÚXiÌ²”­F',Ø9«—Œ•ëÀœÚîÑ[-‡†Ü°aN"hu/M.7®Ê|yVþ(§0Þ}àý#`—´ý”ÃF &å(uêÈIàx¾(N$ÇymŒÞ¦
»vv‡,ÀLÂzÞàkÙ¨ì0
XŒü‡^$8â§PNåÆÕ‘1G6º|ˆ.O9ÂÉ|Û±§šJùðù%nÛaF³L³<:‹£ó7(Ù@ÞØ29Èæ¬¬×æë•h›nóÐ¸lŽrðnÕ<
§q%å¤þ9ë8ú‡ÆM±ªG÷šehÏF¶€oÌƒ¤apþ¾ qÞM˜¤WÝ!ÔM4ô&ÎÉ6Ê)½A™=Ï& 0âë,…ô¤Ûìé«N¯ü‹`“À˜‚Lc`væa/ùøGøoJÕÖúŠwQIV+°§ÇótB%c¤¿Å«(ÍÎ°/o =ªÈ¡èMqlÇŸ³<’O‰½)bšÜÂX1NtØ‹F…ú÷®‡D61¤ <Ä%´¼½!éN‡¤ pí“rÆ_Ì'h®C€·Œò<Ë;tUŸeqøXõnH^Q¾Ï^{º:$»ÁÚ•½‡‰Ÿõ@Jœ#Ñ'Ø5Õ‹ÆmŠT£w†ÜÆd*»t¥í¬ôcµ[ˆ$²vr}ôa«µrÍ8D¶ÁS“D\»ñ&¨nv» ¸L'¤Ë;<¤Š5’åñI45:ÎVL‹\SÁy—$ÃílåÜ ƒºá¸O¶öž?ßÙ:ÜÝ{y08=ß9x·µ÷òp´ux ì³Z”¼ÕÝp(~ÂŽ¯šmqâ‹¹Ñ)øˆïî.ËíéDú*4^”g5¾P!…j â”žÔ„™¶úH´ºóÖT8~|£ÍF0uåzW{/Çâ®«q‚½ÿ„S ë“lJ «ñth	Ö»Äp^Í5™åä”tÛéD)	 ç8K¢]øÝmy>‘ “íåÄéË åK‰7y³ŽŽ…Ö÷O¡Â}˜4/"T\G\ñ‹Ó…kV*#›Nñ¦&P¤ð+â•Ë6•âUZ­u~ÈÍÅNVÌèáÇ?MÒxÂ^ö?þ[±|¢åà=óX‹¦”-Žl\BˆãøY$Z‡}£)íl…}`¨ÂAÌ5ÂÏÙoÀŠãŠ(È'§‡Q>í3–L¼Å:½’Ûg”c§ežÊW½Ì:£ GoŸ¾N´H5óZù-¡¦U>qZÁ}=¥¦$Å ­Ü–z¯)%fOï©‘T[Vñ¾}[ÙÒ4!š0Ü·ÙäKÄrÃEuç¥ôäµÙrnßÇ×wÈá‹û^ö·¾àv ³\ü/å«¹üÉ4£÷>ŠÝTyHW ÅrY#*ñ9B$—)YY!?Â”e9ÌË˜I¿£j„o÷ƒÈ “¾Í"E,1û’ÈÏ×;ÿ±|Y‚òV]ÛÂE>
Ag¯9½¼'È1ˆ¦H­£‚­F ÝÂx~Ê/žGùŒº+N¬”\¬V´™‘Óe£ ÙàU¨eÍÅU
ŒU:>õI“¿öÕÌ"7È’¾Þ¬
þj×?zjÖP±üY=3ñ•Á\}öÔÄ®ˆ¡.i$ÕEL€©¡j¹›a*­Wec&%öø ¡øD¾ýVérdjUÌ
­.ñ>qL:£dk:Ø‰|
íd1'v:ƒ°S	ï*M¥]3V2³ë˜)›ŽólÚEñè *»liS=[×ÚõzƒöÝ®FhL[l¨òÈØÕûÚ¦Þ÷ìéoíšè Úw"Ju¿0¬
Ùî^]«˜îž qO3 UAÚ£ì¢jÌâØèbÔ¦¾ŠgEÖkç„`PF”!m;:æ‰¨OÔ8Áö¶fÇP›àâð¶!Ù4^\Õ]ùÝö
0!°)âJç±]“¼¸®U'¸:O%FWâb{>Kâ	S¹1p€]¹ªKƒ®ìú·ßâá€^ˆÔžZŽwIÏlö”–.e¿Eâ;çÜEWq£Ð:ª¶ ÉT¿bÌtF~úøG]Àïˆ¤é‚PCI´ãØ¶"qÙl±ÐÓÍOÁüuQ[Æ³ÞD{:: 	ìÕÎË½7£­Ñž!‚ƒcÇCØ]([ƒ[>REœeH°ÙæTÙ@—î’ˆ&\h–$®P!RÍW¸ ËÃ0èht!ƒ&˜‚>~aCk¾Fz`C±%Z«-Eàã•$ð©&TóŽD¡ÏCä(&¾¸†ÁÝJƒ$Qˆ§¡¦®æÒ¢ŸŽ£Âa„!­NÚ°HáÓyòaJS’ÁÉ¡Füì²XjìÔ3ææ£áç€ÔñäSu~ñ~;*¢ŸX×°cÇ9ùæÊ_ýµ©a¡9Ð	B˜›ïÅ<0þN]_Ätn©ÉbH²cû´Ä<èc‹>¤pÀuâÿõ>zreé§oböß·[M³³6x¯UVPƒyÝ„€´lÓG*¯½¸Vf''	?*ôèÆ¡¦;&"Hó-×Dþ¹§£·3$ “6Çá[¹D"XuØàr;¾+¨È¡f©nt£$ÁÞÁ 	ÌäV´®4MN·ÝkÎÁÐÍ„™j|J¼š¶·/7JÇùƒZc?m‹ž%jÈÇš«Ù±8çÿIüXÐK-qmÈ©JŒ°1QNÀHá°5¹ÅêX„*B6D¬!²­v³Ó«"@w@RÜ}^P‹AçŽˆA{ZÀ ºs;2.z,žÿ ×Z…ÿ’ãxI9‹±¡{>”¼ÅØÑ9Ÿ7äü)5y‘éŒ«à²…HÝŒ…’‚±”ËèØ> :´[ça6{ 9MH5rì.d,®™ß_èÉX¿dæë8×@h¶Å*aÚãkã±Ì†DÆ÷)ÉQ’Ó¡’`é×;#Š‡ÙÎÅ$Jº!U^uèŒ¿“Ç\¡\Ú
’b£cTðZÿõá‹ç»x1`'à¦åÑPƒÅ¥ý2¾›ƒ£Õ·ŠYÂDÅœ°tÚð3+Y·é9› 8ÃÁ\Í¬ehíM@7Š<;çÇ:°‡}ˆ.%‰ºí7¿|PxoüzZ
ë1Îþ5¶¾#ß‚Ùl«$4„:¡J^9€"mK¥îÚ£¦2„²oP·™Zc]Äà*Û¡=ïoé^‹fŠÐÜ!¹’ÅÈã7V’
YZcªß³¦ Eä‡­#R
°îYxhºêôd=2(kú’ 0 R(2€§üZBè±U7Í¯ôO¦3Ám¡©=_é2BLEoâ|cê¨Ú˜ÑŒÔÞ4µ1Ë
³ÿVTÕîèƒ»ó k5sù[|öîjÊ˜:m>TÝ;£m?ôXæ;ëƒD+³¯kŠ¢Í¯·øù¡¦
´îÕ
>Ã×šì§ñ8Ç›r¢yòÇ?³tUÓ¯ERMma|—FÇ·yÊ-ð€í1œ;[¤åNÑa&jmhÈ*6{«×Zª]ØXcZMI|‚['‰Ž¡’ :ºÎ“—®Å†ÜÚ
Ç"}ûûsYóŽ¨¹·µë&>Â"sTQ¼V"m*Pj09òQÙ]Åé|=›Éì{úµHâIÔ]ëÕaÞ˜2Ê”Seœ)ZË›Ì)>LõÅíØžOsö™µ²Ì˜…QÑN§ÇÑLß<•®Ì[`w±*‘UU'Ô¼¬˜Î×HÒV0—úUÁy›n[…nƒJ¯lÔÅølmÜœ­C3ÕAD*üìÜ¦·žÛáÕÈ–*!àS òÊyKÓD¶Ò\qÁÉá6êæÉª¿öLæœ&y@©tŸ^h	…Z»ˆßu3#Ä•ðˆ‰`YéD?©ƒNËþ
Y[Ð%´ÜØ­ÊfÔ°ô¾ÍÃvå0H4jÁ„XüîÙtØYª 8ç;@VM¬@UÙž›aZÿÉ[oKP'Í°Ç|‚âÒß¬&Eâ1Ð+uÔÑ¼r¸XëáAZàËL—ûÕˆimÎu¸X-ß’¨ä¢j„–÷«/Å‡ö>ë×s§ÐÚš£˜)ŽM{Uù`²òÉŒšJõ|(ˆÒÝ³ç‘léi=5@È„zbh|èj'“ê|˜Ÿ>ÊRžE¤ÞeP™=—h·`ŒÞn•êð²¹oÆÙ¦Y¾¦—^­—²„gÎÚJÒ;¼ß¶ÐóƒIÚ0¯ÁÚ -´¡}«*§@`ŒX–¬8nÿ½¡05âìáÈnå¼çþzN-Ûûo®DKúyK,èFaX†â9_YOÈ*ÙÄâ<ášÄ'@/h(KBztM·ïÁ{zˆý^­åêey· vl§F+[¡T…áÖ¼(³)··¾…mƒÄJe™6 ÖfœžšG~ww þÛ_Fûûï^ìŒ~Ü1­šñÂÈßþz—K£~:MâìNÛÚÚwiQ‡T\*Ÿn}(ÿ"Jq¦Ö1<=£ÌÃxeì×dà'‹T]®i<D1­™%âìïTtZëñzá³z~_Qv‹H;¢UqWgÞ\qLÉÇT©fW¼Ç×:òííIÐ´Ã¹´àÜµ«º}èÁ?	$Jyâ¤ü	ž“+T¶ KUGåøtþQÚÀL ÓÂB­ágU IzNž£Ñ¶Ê=Â¯gãm‡)Ì\QÄÓM“WRôË ÁúJÃþ@ÙœP«2™—Þ¥!ìöUV,  ¿ŸG@t¢XHÀø’â9Îb>°07´ÄIiç´ØYpÀxò€d Q~ü·‹$„}L½X¶4‹L)ÍOÃ˜]VggaE|¬F¬__°§ÙOÑïN<¹ÑÄ¨ŒR–•µ´Dö ­à«O¯úÀ ìšÂ
ÀKô„^Š ÉiÈ-ÉÊßJAãÄ(E@NL€[Ÿ«»¡²-è,§´7árÍ­Ì%™:ªRž6Ys›^à°Xm­´¦Ü³*¯fÜS&eŒ’yJ-Í£M¶¸Ë¼o=9uº¢L•©¯îù…Ù-Úu§AöJ«^ïÖ²j˜vÜ°nµ}œ³`-_.ÿRú-ñzAÁÐ«™†Cú;ÏÎ¥k”qTžGQjzš˜Ëçè bz±à±ºžx¨yGY?} 7Ã<9%~ON¿Z]%Ô/ënšçÁÌð´B•ÆÏ?þù÷s@ÕÆÊé­ÅÊ‘aur8š«ŒûF>G!–‡ï~LÙ Ó†îÈêd¨U$áò#îèH¼£óå3Dsúá:ö`€¹€lÒ[ˆ¨}­h¾A`»5Ý °õ+·–Ž} ‰¤ Éxãjí—×¶3¯uh8}²ËW”\¶ë+4]žÇíG-@Ûlñu@>Àt<zÈ úàf U®k|à<Z[]¼õƒ©h™P·P È>'=53 4ü~ä‘˜Š‡ÖTxáëzM;.²X§¥™›£­¢Ž1µq¾<Ë˜êßåëõ,Ð„i\,r¬Â} óáƒ'û¶ƒ›`2‰fåFgp‘}‚ì,Ê'›n©`û´Ñà{‡ad¹r²¦Éq&×b0#›fl_[]µ].Ý|b5ïk7§ÛÙyÚf^Ù}óê"¶áõÇôöe“øàþƒ›W£ƒýK½>t7%«¼p#GfË?H >Pn— 
>—u8æ¾im.ì)<Ùos¿±:ã®(ËÙá|FÙGØ	aÆ&@àX>±sÓ±µ•rˆˆãÎi€ÛrÃäíÅ¹S'pÃÈúñKÇÚ•Æœ\W·d¿Þv&ôíï³Í…¾s/2’—»“™½ø…;’5üŒ3‘åHC?ïDð6*ƒÝídðžxæBçá?ãtäYq‹É qÖy|Ü'ŽhozuÐ	ÆE–Ì¡“({•ËÔ í0"šº'`–­ÿ†@˜3Q›“h‡!´ùEñƒƒ™ÙšÇm][Ô ]pÃ$™\ýl|C²z}£^\ZË¸e( LK§ÜgARûU‡8Žüì.ZŸ	t¹Û5:½¢K©=òDÓgñEv×\²Î(	ù;¦H	ðÄ8¨‡@Ý*ªglÑõ“ðærYPÌ|hS®ì)7¦ú)¨]"†w<´Fw+ô¤à7ÙÛµOQ`·ëeº(•ÌŸc4Hwçb¨wÜ_½ÿpp¿gsÞAf¸SeyRœÛxÒèLÙ¸øf\ÏÓÎ¢-‡ÈŽËãAÀvsÜÆÉ1ódŒÇÌ_{‘Nu¿Ç52F6Øæ{kæÒ3›®«ÐÊ˜@¾	s}‚¶ð}Ý0aÍ“åóMý—›0Wô6Ð—fb7Ÿê Ó‹Î‚ç{ÅMWý¹q0é¯mF[ìÉw¸.ÑŠ¯qî½Aoo·øêƒÖ~u“Ö’þÇÃ7Íšê±)JÅ-‰c}”‰¯n‚>ïª’–“7Ÿßêèw·›Øªèu_ÝŒV‡Ñ°AYÜóv¬ÎùåÃÓî‚q6Œ±é¯YÁÁ«éè°øÒvÖ6‹Z·¾ùº®Ôv;„¬	´öåcczÝ(2²K7	½1§o#zãEMÂ 7¾E]l‹Æ¸îØ´óf#d…¦ŠÖÃ[(ÍÀãAEÜ<kjTá¨ÎjTŒ¤5ÔÇÿÛ°·•iž£T+Åê$©#Zìü¨Z½j7€º'ÿ¨éJ·ß€ß†–×¶çpÙ¦õTuf®Ër˜ã 9…iè
‹žjKEUo¶áÚiTkÝí‚9y«# Rubk	¸Zf³åµ•ûd™b*ê%M0íÜ…Ÿ€þ`/¦¶H]<ƒz> .š&£j±¿n€³dùWÀ	Ú¤›NT'}½rµ²'&Wñ:_~øËJqUyelÞXMG¬¶øÙ jßÞôbž”ñeìM±q¥9â³þg¨÷pÈªœAVïH†[X;›©ýÉÐüþ{j»ÉÇÈås´-¨Ïçœ¾û†jÆ,zðùÃ“àDu•M{»ãî¶+œÕg/ë<hf0ŒB8©ŒU¶×½Ïíã¸«zasÛ¼öW1ïš3Ñ6³®9+¾«9gN-fœebó=i3ß9ßÿ|Oþ6çÛtÖÚfÊM7Ów5ëÒùhÃÄË|lî‹6s_È¹/üsï‰ú71÷š{ãV3¯å¿“ygVúÆáAoôV!Aï4è_.¨­3’‘Â”s¢yK”0=2-Ånäúmé¢<‚ÐO™·°@~$ä`#hq•	Š/˜hyj-Ýpé>–zí¾i5õ?­¯”§Öà•Ü—›ºAgþØÜÜIâ†ÿÀØŽ8uw¢¾ö]õxsûlSði…Ùö	©(pýè«Çñz‘¹€u	’Û¢epÜ‹–To!LñuXð©pyJÍš¾ýÖ«®ÒY‰Ç²ŽVkýsJ…´ýRKi¬pgËq+óc²_N¤EÝ/­\•QhiüR7±z2~Ò`ž'·guGÏõØÐ•—
*¡‹X´†úÐûD»‘ ‹ä¤*ríb³²þ4+ëdgJphÁ]ÌoÉ@Zno]žÍb½gá¥7Ø®€Þ/_þ£:œô•ÇfÄ\»Æ ½‹*´tC"Æ
¿­Á‹­{VWÙ7Ú…^/C?UÑ­L½*Å‡jóô’€…ùž¬ù¨JgÛ®ëQE#5Ûa»QÛ+CvJoºU«®v§’Åû¤Õ†åÝ²î`uW
+Êv1Ä(ÖNÍ¬?¹ýVgbŒîM¢šNf´’§kî‚÷Xªe|à«¸½ãµâu:£9;©Ø1êFÏoºYSÝªnÞÓ…*4ƒ¥bÇ4ZÑbÆÜr Ru­]!ÀíW|Û˜Ç4½¡—ŸÂóð­A1›ç³¤2ÏÀÁ;úI áQù7~Z”
ÊræU¼ª-K‘(Ÿ%óM:PE†+ÌlØÃÕ0#©TÐvë©œ×5½àU,…MEÍçÕ$­FæŽ@Ê5·tÅoC}Åõ]ôXÝoÁŠ&ve_j„ªåSÚ§šËšµ§+’š›T·tu}¡5ÕSs{ÆåsëfÜB­ztVÍ­Ë+N†´·P»ºz«¾AÖ¤vÉÍº€VÓl˜-5Y›æ9½•Íó÷ºæ+hd%rShE‹*B+ªYL!ZQ‰O1Z	r7så^³P'>£BUõÐw^ ¾-¸ëÕ°¥UÒ"j‰*í¡,å²ësoç- k7ª²_·U{¸öOöe\GªÖ mQeG+/F	ÊB‰°ÜuUfdÁYôÈ–LÁ	¡¶¿g6Cš­Öß·VC.--T ë3‡à¬áŽý-Ïˆ—Ÿ˜&€àUÚ æöãUu4èg™Þ–Ù›—'Ú¸µžeN>ã«H;Íª¯zö¡¹Œ¤Êv¦û¨wˆuö®A%…O%°ó 8½¯LµªàZ.?átUƒ>Ãè+ï‹-ë•ÛÖÝZY²kÀŒÏ$K 6®ùçÂCØÍS%[êG—ÿñÄe?à^Féé|ªÂ&E)F‡—qÙ³‚â _Î€ÆƒDß—·ƒ±}BNu¯æI"= lgÛ)È¥ ¶|‘eãw¡×cû‹î•t+ŒÕ4íÆ>Ä<ªìI„LÊ1"]oRÉ<›™©!Š“›ªž|Ùå­ ·_ˆ U˜;t,6®ª\&¶ð”ØÓáÂà €™þhFmO¸Üab†@¡RLÒë<1Ë©HRÛé“Î»q¤:Æî¯ù“U?åIËÆ•üi Ÿf|}ôÒV“goñáÈÚìõ^1+—ô/è&Ž÷hõ-â»€ÖYœi³=({š€µ²ør±c¯×-çpªrÙt¶`+ŠCØoµmösàîMñæ0šÎ´!£Jñ]EÖfY+L§[¾Ã¤éïFØ%V±˜òç º½ñOâ,¨<¹ a_¢n{böƒKêjÃé;ÙJ0Þ…‡ÿ7áÅèŽ  ù;õÎ7Ñ·&Ï"ÚB…ºœœT./Þ÷!áA28ú¼	ò÷‘¢+šî“Ä¹ÈÌîª'žUum¬*fYÊpžÑvFÎêåaåª	&j½ã,ŸDÐê^š\n\¡ äYå£4žÂxyä#}'å°¨G}"ê”/ÿØö™xmŒÞ¦ U.õR¸v7œukTKÝ‚n\ñzáà,ˆœL9©WGÆlÙˆó!º’óúêh\¨Ë·!áž02B] 8¹ÐñvSÓñ!FÊQFÈæ°c*\›¯W¢qºÏCë²=‹ÅªyNã”Ê#ìö‹S¡l<BIõð^W8±šçÙÈÊAÒ08_0`Ï»	“¥ª;„÷·zƒ± [ÄcBM@NÂWí¡3+—Ÿ¾êôÚÁ¿æ!b™>˜ÆCQØ€@wÈðß*,\Q¨w}I.îõ½{Çó”ú÷EEo êžá$˜aíÃñç,Ïh”Cøio˜&·(VŒ“ö¢Ñ L@P¥! i6šdtG:zÛ¿w=$¼á!ö„²G”öòNIwªb—› ‡b>Ag¾òOÒù..ô³,¦ÿx—‡äuåûìå±§ÿCbq¬]9¤!y*~>ÖG=*£Dõ	¶J-¨±b6hd¢w\†\B›Ê.]i[
+ýXm "©†¬\#Ñï¶ÖŠlÈ¼¥¥õš±‘2(¸¸Ó7ïÈè¡À»Æùò2“Ù1³§¥E¶Ä[M	ãbFß¹yQ×–²Õï[¶ûuí±£A½==¥¦$,  T¨•|¡§Ô”\ÖJý(ÞjJÄ…Ü™úl¿’ïF)µ3ó‚Q£›xª9*w´#†»X[°fÒy’<éâ¿zUüd_ó¨_ßö-£Èëà:˜ƒ± {3ÊvÂ 8gAÒEŸÀftž°µ£}ëßãÌ#¯šÚ€
vCÖ¿m-Áhà *yŸ hD<c¸§Ù¹,^ {ÃFì${¡¦ÏµÎlñéÖ“ÚCýòY–ÿ–5üúÇ·<‚M#6ÈvpRfÑ$>Ž'Zx(ÿ"šf]]ÁÎc›¸ò½¨FD‚&h¬?–†GZyYü6.O»øéŠô7(²¼ìvƒ>[Œ=
Æ§J+íÝª
ª¸¼öØ(>öW×JóÖÖ ¡¬ÀV6yÄkUîêYÌ•>9²Áô¶ç©PÖ©Þ€¶2À2€^5ƒ«h‘"ª‚D¼·Â”SaBªÇÀÑíV‘Aýq´¿)CØ°y«Šj®…#ˆCŒE`µfê6ÚÄ@‡Æí¸çn\Ÿ÷m†•Ò‘ÈéÅà½Vg½t» êˆíƒ©ëÆ{üT¥ïcÁ!°2Y©…=Â®O¡âŒl08>Åx­XKž
l`:À4ö¬æ¬nt£„ÆË¾ƒA	]9P?¦-`œÙÉ@ëHãÀäœó>›‹×:¨R+˜ó…<äw–Ç4Ü­ö&aFÏwÞmí½<mšÑ8x«» 6þM<åg¦\_LƒbCÅwW¿Á=]ˆôUFaEJj|¡Jaz§ÔX¨Ë´$·Žäœ—‰5&Nd ¾$}ä‹i$×v^}vW.„µÃë›¤ƒ)`b“'B%ªé)ÔÑšx‘.‹Œ`âÜcËâw*ëb-^•+)‘bx ‰V”L‘ìx“aDE
Çæb2{¿ÖwNØèUö}…N²<\´3u >Î®€¬Öyõ£#Ð–¾‡ñ,É¦Ø¥mö‹§G)®ÖÎÎË<åæPv‚ýû5üˆ1$Aší¤Eœâ—}|mÑs+:ˆ¹(1¾[ÜÂ¿lhÚÈÐÒ§ŠõÒ"Ì@àT´¨/Rq”ÏâTD|±¾ 8³÷—5cW1Ú(WãÉ€g§Lªõô°…æ×ŽrsG~Œr;¬´•yï7=Oï1œ6L@„Ë²êºžwBy
 ¯¨†Wp<Ì'2°¡rÐÜ»°Ç¯0z~Yå1¤…Ú—·ÝWM÷õÎõíÖ¯Ã5x<:ŸO/©ÔKwîê‰=É³ùŒªŽ¨Âˆ¡òÛ!. q”¹ÚôI[ÕC²*èãáÇ?MRŠ£2©³ÿñßŠåÈ\¡Lq!Ëw‚É©Ø!Í h3Dá[u¦èaÌØŽ@þÀí^ÈëõáûÁSbÎA¾7þ	ö»AÄ·]V¢Çöí£4€­
º3OË·fômúEv¾‹¼0	•¡E¡¢\2Ah@õ©C8Ôz½†¹>àÆzw4Õ{¿Ñ¦SØÒiI£“9ÈÂèÆHK#Àé9ÛNw¡…?ý¤d@cû7í„¡_’48×ÜÔ3›ŸœZ˜l…ÍºCÍš×ƒÿEþÝø¯ñÛõÊmû±PpepwZ"NX>Xœ¹Æw‹Æ&Ò,Ô±ƒ'.‰œÓA*v"_p²¬¾ †„*gCN‘§Œ
¦P¥¤97YŠ#œGùZ>)»vu<ª³fVÉÍÐ(ÊdIjfÙÈàuSÕ(?ú«àY¬ÆP×²ªÝß`@RŸýÝ¡Œ¿" Kš×¤ôƒ9B­ÇªT p¯Þ®¡6[æŸ6«7Š:ÍíÇ²SŸ‡\Š‘cLo>î"mK§D†6»ª½ž[HÔC%ãŠÊ$ùSl×¥­Ÿ%5ŠSÆÐèªêx>8$¶®CµûCu'ÚfÚôCÛ4Ú·ª²,dœðÉ ~€•–Yºº6×ÁiP0¢Ëw`*Š®’z<:9¤Rb»Ô¥Õ2a°‡È[ÿºbu.Élº«ÞØ=ˆxÚ„÷`H–Ø/…T·cƒ…SJíƒFÛÜTA®<ù)ññ”ÛªfÄýC#àÔ_Aß¿q|•–åÁš9ùñ.™ëIúQ—.ëIæñ”hQ½1S§ ¢ôQº¾fÿ7Šë¤Ls@-ŸªÍ™ž#2u­ÅˆXYœÐ­—Û[yÒèœ‹zÔ¥*7¶Q½§L_º¢„À]6”·td£Y™6m>ƒfMrÚW­e¿<ú¦m‘ “$ñ‚ðfAfEñ\µÜ]–‡c­”3øÓÚ<Ów
Ë'ÿIL³<D®RE¶ÇÎ¡jÎ;þˆªËù®À¨êl†£n‚Îˆœ³L,}³ušÒ±Q€ýŒ”Ñ2?!éôªEnxÌ!G/ãøˆ£‹v~p«“‰ö''ôðÖ2©@Í=]ÅÊºÃ~ùžæ_ì†Æ*.
¥²$µ»ñ^J‰‚mhUÈÍ«wÜ‡å§?ÇúÌfÉå+fÈ‡ýÄžtÑÆÝƒ'r«¨6ü»(ûŒÖñ-úÅŠ¹æ YøŒì‰ú¦eNU¨Zÿ”´šu¿×r¡žì’Å5AÛþö[q-ÇTóö¬}cZœPÕ£Ç¨7ïyñ»µŽ±—Ð:ttg²ö$±œwÜ	"^YèýþèÕèéÇÿôò€‡g'ÿã¿þçÿƒŒN²< gÙäã¿’"ç‘Ï>‰3ÈòþgTçßûøŸèyÃI´¦-?™'§°¾Èt—¬uÊEµÂ öôÐ$Å°\“x’Ð˜\4®; 8ñqOâ`	øÿôúÞ½ÿñ_ÿËÿEFi”ói4´ˆò$1}\$VUf!ÆAÃœ?ey
3‹Á<E·Ôà…ÍÔ[ìýhB»SDsBkI@~§l0žUÒcß`ÄÙðÞ½È-Oƒ8AR”oTÑ¤UùJÅ Op°¹qN°`zŠ (žF1 z6Ãb¶A{9»ãÁkzò2?tåãŸOâ’Ÿb_·öŸ‘ïÉ?ÀÿÈ „ÿòô*T÷ñß
+ÌYD'ó4ôfŒS‘ñÞ½3óˆa† Ó“(d ©/b†D€{÷ö#ÈN/{¦Ð…Ó²œÃ••"´ñã+|^yždé{¹ÓIYÓ'\3­žù= ’ot)}@Ï)€ñ×BÂ ˆÄÇÄ' `öa‰ìÃÌŽ?þ)-–Þ‹XÐ¶Q<@š ª.|ËÃ4j©È•¡èÚŠ“äÇ[gtÙ˜Ñ	•‡åîÓégóÊ¾(N*(Ë&{Ý:„Ÿ_(B(›Q¨zä‚\f7À®á‚¶+hú:_åÞ)‚™‘q4ÝDë`Œ¬‡fEâ`VrÈXÏmd‘°åÑq„t„Þï­ÉŠÇÎt¡Žo5²‹R:ZU˜ãbK±Ñã,& wÆñ!-@æÛÃÃ•ÿšÅ#!]ö's\Q…Á—s¥Cæ§,N»È®ô @}qMcV ^€G¸X)cˆÀ"E\ò‚Ô7 <m€Ÿø‘3
œ Ìô­„â0ò …Å0¶º"¥[§qæ¼[!î‘K|ÎlNö®éNò÷0Ë¿K—¾ü×”C5éV|(·SKÈÄy$â4OÁÈêG[;»‡;d{ ­Á«Ñá^Ÿ+Fš}¨•œ”Ò³˜Bi`r þÊÁÛ3V·ýßì¡âƒƒÑ+²Gö÷^ŽžÓ&F/·w·¡E?'òrÎ€mI:5&hz•žFË0cËrã¡-Ýÿÿ†œôÞÿüz‡ì¼¦›þzùë°Ø£ƒ=øý#2"kñÍÎÿÖ—¹wàß­]òbò²Zå¬òù.ÍóBÂƒìl¿Þm»ƒ _v~|ÃÀŸ|0Êš}÷åëÑ+sj„^ý±WDaŒ>êZ„„á‘7ÚX ©ÀÎà€ÆF4!‡´íè8˜'e×df'(ö1±òl¯Î¾vW~·½rë«C LÂŠ5‰`ƒzuÂ.ÁS‰Ñ•¸V”\™@uT‚Û§ÑqHµôºýë&&jŠ›nWÖÕS¢-R{ºz²/éÙÍñÐòÒ–¬D|ïõä ñÂr<®´*›UÎÈO@·aÇ¸©
7EëŸOI)$Ú®È&šé—Ï¤ý¹m8¦Œ{u'®VÑfDí-ˆef†á(ndn9õ"£Æ|è5õù·…~A¯`À”ÚVÄ’dõÖ›™W¨”9‘œ¥—0Íˆ49Û0ko¿k7©U¡*rý‹¾)^D¤)çÄ«cÜuicpÕ%ƒu é˜ !¬.í—ug¤ íŽç“Ä4}LeŠ@ÊßÿÌöDIàAòKónlÓ|™ÏOêvB#‡±±Áoçð¾4Ü^,RÑ†'Q‰K—÷Ð(Ñf)lžï¼Ü½ÚÝ{7ÚÚÛ90ûbÝgÀgÅ³(¦èÁû2@—z·õ‰ÕUF=²~;8Ïù'Î³¢0—êEŸåv3ù2:ÿd3©ªÅD²Ž¸	éw0‘ZWšçšüóhÞH±î‡Z—NÔ½þÑþ×´\¤ÓîX/ÒNêŒøø­5{ÖFZ¡iÆˆ×pŸ
ãEÕŒÇ€‘öÑ6b¤¸`2ŠeÌÈ&‹ë%Wz·,÷¨¹™l•¨5¥íÿú4úõá[hS­.· ¦Ó¾ˆàÜ~1°Ï.òèWr¬›7øxxq=Bf{ÅK‰<ÝÞc+pMùÂù¬"^`–°gAy´x©ºžÂ«XäæŠVxaxl-•'v{ (¸ïv‘Å^5™¬pŒ_ð¦½·ºqëê¸ Áë[†>ªÌ{÷Åv¼‡×4ÄÐm¥ÔiÀN¸€1Ðòe €œBû”âwÛEajdQÝË¬™â¯„ÿA…e·‘ùÔMrCyï F0HêÊîõÓ ^‹n=qmV†ýå•!Vi••Rh¬Q«’fÐK €iT^*L5¾ž*Š£{lg†N÷TïUfÙ{÷‡+ ÃGÜ5rHó×võŽ±šÒg;‰³¬î«äƒ0`Ã<¿Ÿ®hÒŒ¼ò³#}ä›+ÙwýÆÏ4ˆ‘è–ñ	¼Ì²¢ˆÏ¢dJ•^üò ½d÷Ú¥•;¸ëX$R€‰@Åí¹/¾¤»â‰qãPî‹'Îõ‡L¬•Æ b{”VZÒê•Eú¾Ø¡-9Û¥)l±í’Ûf‰Ôiü¢Ü%å^è{ê‰øÖÌ×AÞ€ ˆ}–vËØyÍûbs•V[VËjO>ñû$ÒrRYJ®ÏhÚŠ&2ZÊÝŠ‘]ÛÖ„¹¦CN~p©ÕF;õSLƒ‰|ÿÍýrúÂ®aFÄ¹Í=å&¸†iëõ{Å¿hPzWÃ"£\ jf;“(é†ôNKÐÿÂ&šÆ“¬j5íNÅj’ê-æ˜*¸Ö}øâù.úÔßI¨
ý‰³æPÔÅ}ExÆ÷†¶úVié—0Q¢°tÚð³<›²NSßBÕ†#0¶xSZ‘Çàoè}Œnžóë^À) –8">z[á®JüòYv[jÚîßqoÆl¶O(ôsuB•¼rã˜C"ÍÚ£¦u#E_¨ÛL­ñíÄà*Û„}=ïŽ´8×óèÕéVx^YZW¤†IzLçJx©Eß0çKõU€Rwãa‰M(­¢OW>fGtŸªŸ: ø2àÞáéW§çjYLÏÕœ¯±Òth»2hì5©$Æ@ìhˆg¨8Ý<Ù‡Æ,—èP¸!RK'“	7¬ž$¯Ú´U¸Wbr"Ÿ³:Ü,Ô˜E€¹Ç¡i"("½uÐk_¤T®–Ú¬M	ª>T±-¶olF”ô!^~jjYÊ¾Þ±ð­â Â¶v?\¸ ÍT_º	è­Jð
”í¿k+4äñ60Ö˜ºö´×
Hûšµ¤ý6»Ì$õ`B,~÷ìõO»ÎùÔµ=tÚ¡]v[w¸BO·q_®¬îûó µ'mf€sƒéŸÿŒ¿hÖÑèÜ	xe}Õp­ÍbžðúrÔå+˜V~\ vòTBnLæ÷ê³‰¶g†¯;©¤À‡Šþ”e‹P»¾úØøR|ˆg3ûßDSh
™$}»5ó™¤®ÞŸ²„r ®Æ•ë)¨ºÄ€ƒ®:1>tåI¨4«ç˜JÈw)‹><[ßáã( STãé”¦”iê™qBj©ªûX©F1Ê°Kè¡…èµÍø.vìæsý®7(t`V©WÌrjØÆ(kÔ-¼ŽËßoðÁ¶âŽÜVÎ{î/çÐ™J¢%]ÂÃ¬XB/0|õ°›ïeÂ5‰OÒŒiU ¿®é¼§††œx33dÞ•œÙÖ!k$ç!Ù@íXÌaûåò2´ÀÕÊwä0“—ÁY|ÂÌ¾[`k¡Eã¢«÷óSH'3ÍÇûý‹„§Ðñstï‹Œa7Î—c=R¬ë3Ý³(½cu5¯W†ËšH7ƒ¸O„çîûƒGN°fô~¯üá›®ÒMZ.ýx1™@õÏZ÷›*¾Ì"úƒ!‡Yr"~ÑL÷ÝÐCÒ±Bä:ÁM¡KAÈpäý4È·Nƒ¼| ¸ÿÒràÎœáo‹¡Øñ]¯îN"sdöeN ëÛ—;yÏã¢¬Ÿ7ÌA9tê?Õ	g`Ïžé¹ýªŸMÇúÍ„Æ—ë$C‚ÿ`DÃhG1ª×û$9Ñ^z£å®ãõ™-è_Bâú½ Š‰gqŠ‹qîL§J(£‹ÛÏ<ÂÆºÒtÂ{Aç²|£3®
Õ¾²P‡¥…›jS_o™#‡éÙåVœO’è~u¿E¬•[v»ôÎ†Ð•½ÚÎk®C*ºŸdgdªßBß¶×‡ÁE€>ŠÎØ+:Œ?í³éáäúïÞû{~˜³±½žUw_¬é¿¹iJÓ"2VÍ}ºj~°WU‡¶ÿ`làUÛ¶ÜØEë§ÜÐÆ—Æ‡ô—P×t\Ûiý*³5¢ç¤Œ`õÅ!ÙF6/ÏcfölÜeÖž=ª•ÓÎH+hÞO§¯/(L[Tæ5Ó(ÃÅ ]Y\ûš©
XêeÚ~še||¹<ŽÊó(JÍh~E4kCúƒ[Ê |Uq§jâÑ:+¿Ör öÓ9Çà6ð¿4Ó•ôõ‡VMán¦“Â‡èÌ
ô †DQ¦Å*ÑÜ³ˆj("/V‹vi^Ô"âÕÖRµ2H¢5ÃÏ*¼¹A`EÚ„Æ½BŠ<OËë«éT¡®)…7E­e}°ÕîîZÂ…ÌõßUEàlzmL©êX«ÖÒ>§˜NŒØ˜ö: ÈV'@¬Oã°´z‘Cå•Ãi½;§¬U£µ·ùÜr]ÞÁª¼‹5yûY³‹ò—«+r‡åé*%¬|_}¯20We_ï7—'£Ëº|åŒ3·D4ÞV÷nøÔñc«OÄrp'{Œç˜µà9f_ÏÑy"úù6™ÙgÞdfë›ŒØE4¡ºvGY„VÎ¾`ZY™`¼ö®«Ô3L[e©fê40Þ…ªSC[p<½ï`8’ë*Ûy‚÷·¤X•Q»_/ŠÒU‡À­Š©Ù"ìm­¾•Ñc¯¯?®í¾þøíøõÇ±é×Ÿ&û~£#¶þF£®Ý¿þTÞÐŸšû n·*îcõÝPuÀ|¬áˆ=ËÛä?|jÜÙp¨JÖÉ5‘š•ò–Ö†áoM
Õ^1ìá;ö“yÁÙû«ÞàªLÓ»Å¯å~Ê_}uÚBú òÈ
|[¸¬¹pqµóMóåY{èíïë~W¡õÖ!²ËÐª¶Âõí'=¸U^ÎÅ‰ãåŸ`2‰fåFgp‘}‚|Ù€&QsÍ+ÝžÓ·Ýè;Û@LÎ×Wè¼-L ½·ZÜN…qîyÂ+§n¡PØY9QAôa˜A™¸¼¤‡"Ä½nÃd¸BÕ\ËÔxW}\µÌIy®JgÓö1Ø¾Ž¬ï\F{ÇÇvB0ÙHsÁéððP$N“8¼\•¿ç1Æ`!ÏhX`ÔrnKC~
‘7üFuÎ?0›%O§#%í]h"ÌŠÝ´Ô7§B¬Žû-ÈP²¼§m	ñb”¹Š'¬aç4qúá‚ñÎ1Â4ÔŽ˜§
?øÏâ|Ô”ÚsžD£|ƒ¬ƒªS!é#ÒFEI»‘Š-rw¥;­ÈYºfxnŽ:ÚA«&÷5‘C;EÎæ%.EãRÔ,LæÅmä Bí¥ÝA`#´¶öŸUK:Ðl+”î+‹£ã„á:´0ŒX{_	@Š‘Í¨­ƒA§ó„j}…}hÈ¯O´—–ÅED…Îñ«eA'ìÂ;Å_MŠ¯ êN»z)³­V‘Êþ• ÍÂ$gŸ›WÃÌðÂÛj¹é¾R¸ÕÑÃq€4¿€4ÒËn¥%=po>q’ZVUë£ÜeËê²;ÜY;Ú&Ž›žˆ_-jDOÔï…a~J‡ØmpSÆiýkÁKi”Ôk9x-õÉV6¥FLÌ÷^Ë’ôúê“ƒ¨¹¨o`î´bìëUóv#2÷Å2z˜tõ¯%Jäžó:™It\ú46åi„^y™{¥¾G†Ð÷HÈ€Gk«³‹·š†`>Ce>žò@7ù€ˆr‡^µïˆ…¹Ôbô!9_^óùñRUZ"ö0]½8Î.ªÏ"xpÆ¿J™=51Mõ›ó<ÙºŒ`>žüOÈjeês-‡NXq)©á'>Q)O›/X¡ÂÙÖ
S§+D2Í‹W!®§®uñÔmÙÅKWœZÛ¹z*Â¾ÒkÆi‘Óƒ
Žz,ÝŽ
Ü[QÄ9C¢?èç±\iƒ”.§Ê¯ÊÚêÎR…#w=ToõX+´Zâ©ÄëèŽ³r=äŸæyrgà®>èRpþÃz·7þ©0ÃZ%ý=8é«•†Œ­7MóÜØý›—¨GÃ¾é]õú©zÊ!U÷1:ô;í‡–›*H=ÎJüÏu¯YúVÄ0O³²+ ÒP¨>³·‹,JyìS¹(‘­¡>ö¾vT/ù“ä¤Ê2ôf‹y@#”­©^}gJl6Ç®¹Ù’®¢Èž{Ø›•
þf½gá¥Wà¶€=Y¾$ü‡dÓüº}Ç[N$cGøiðS#¥¦=Â›Rè‡öu”•G¶ÖÜ?‹•\TVì„Õ³WÏxµg½4æ«’šIOSuKÄæ…j—§Î'ÉêûÄa—êÖkå·J«\ÀF/«M¿Ý`|öøžÔÚ19çY^#‘_Uà¾zt7îuÍÕCy;$,ÒLQ£]w&³ã/¢7Ô"£Ü?…ýÎ¾U{›þÕˆ}É(fœãM£0žOuxþ¢õìR/þ·›ßÆË³ÕVNmFÉeaÓÈ°¡&9>~hR7Â6cTõÉ`r5ÌíÝ†ÿïÿ¬>Ô©Ý¨–NÍÒÀÇ‚ 6µ±ˆ§àÍAí+ztÛÙ¯Ò—A>Îó`Æ¯¨Ox]>º¿ŠÀ¬ÚU¥·9Ëò,×4©>”;ƒš×”þW‡yØ´(ñ‘J‹¼ƒÙ––Üè…¦‡½æZêØ#|š@¶ÀE2|:Tš”[U×È¹òñWÖ:l«…=M°hÐMã6Œ´	?¸ö„¾‡YDyíkšo›5PÇzî{±øÔËm´_
ýñž¾è?‰áî”Œ rIÚêí‡‹7o°I7º¥åÜÆ}ÌcŸzõõ}­·§‡.nZ¶VA;ŒøíQ¿{=§BÝÙ°ÿñÄÕÄ4£¾»Pª4ßc¤gMW¡4|è§gK·)bÕãEÜ§‡mZnðÛº×A‚ýÜžf}×87aÍ…ˆì‘äéK{;¤#·ÞI¨Þ]v@háÛï!ìJ¥¬€½þpó»Æµ•yÜÊÈLãa{MëƒÚ;ùk…„…»³ ƒþ,`ÍpËN.b¬`6å3\Ð¿/¾'×ÙüÖí×w{Æb)á/ŸeùoyPÅ®Ï¯š§ünÁjÈòº›,ê¹™ŸitUÚ/Í¾Ðœhzü™fa]­Ë6Ù2¦v.}ÔÑî†eoöñ_>þ÷ˆÆÁ1·?%øÁk’­å¥"®h¯Y'ÖÄ”Þ=j7ÖZ¢SÃå5õ°ó+¥^õ¹»tuÔfùFlUš^}Sž[•‘7âå-Û|EN®{Y®q£Ó'Õ1*Ë®Üi:Åvå”CaVÔa ê1nòqF¡*ú@Õc]ôcµÔÄ$¨zœ›€~ÝX;4—Mc«²,ö+ˆ¿Û—ÂøDªÜnØ¦äušßxuQ«oa‚ÿÐCðéÝçRžy`Ûöbµ$ÜH ‚ü†¤ßoG¬?ñe+ÌXåÇ`ÖîÐk¹bA[Z³Ð¦©·…6³g¹!¸ó 8½#xW3lþCò*}Í•ÇN†Õ©Ö—¯ûëgŸl7oê™dÉìïWªç×Ãf®Ýg¼çy,§€ü€RUóQ^!Hºåq"wlßU>jŽà^>£&‹ß=¥1ÈŸÑ>ìQ:±*¸Š}óÍ¤£´/¢ú	.—WÉp¸ü¦pTÀ_ñ’¥bÁäZ'Œáo˜g3¤k¹ÏE4;Ípµün(âÖb ¿~ãtHVû¤˜	lÅ«ƒ_=òSâ€AÂ(·&Ë­ùEq¹pK-Íá7ËeçxÐ¶vj÷¼XÓ„nz¸…ñØ55^Ô›„lÅãº§Ê©D½À•5ls"ÒX”^}•×äù¢«péâó¿Ã¿Ôn4®«[7ldû‹Îh»‹òã­¥’4þ/Â‡ÀÃJÂ_Gô«ýž [è­.Kæã)¢¹·Û?|Íê/¤Þç¯žš*†fa2Ô-£ÄW£-ðøF@Ÿ¤¦/#E1ßRUÎ’ðÁ‹°d+›Î0ši%Ñ÷^ð—Ÿë£òè÷À(DÕÛ!?”ªëík0|2$P#Š³83`–++kµQsJç½^u…k;ªöíE]ÿ0ôúÜH´µÿìË@Z[¨›bN­ UÜ}¨ºÙH¸ì[MueFË¢¥ˆ÷ôeà¦ÐN}©ª´i†%ßÏxª¾}"<õ»]T­GÓ†Cîþ²X¨4¬æÞÌT¾2ô«QjÜÎBeEí"TVqS×Z…µGl_ØâÃcJ]ŸþI6‡ÅuÔø_êzôùõ3—æW¶&¿Ž-^ý2ø–ô»‘ÇÄùÔÏÙÝpÜ{Êg!fæîwƒ"êÐí‹GË—ÍÎÅÜ_½ÿpPyÝðëÀ*>ö³`•8iþÔh¥pñxõu ‰ø³`ŠfRð¥rQ¦!ÄÏì“üö‰PÐ	ýYÑ°JùRQÑ¶§ùå·O&aþ{³ì§/[Ü¯û„KÝ&ëo¿œ£Àó&Nq2Â€t÷fÔUMRë¤½B#‚4€oM²¨@ó+IwP³YçŽ¦)Ö ÷L¶Ö©áî¬ç¬fýŒˆ c[ä4€›=ÂPNÂŒlQè|ü#€§¾k-®7ßÃÇ¦ºia‹; ‹Ü
Ã‡ùÙFÉÆB7¼ðÁ[W´$Þ1™é<H:v´qÿ³ˆ)2{Z“0ñMÑ±ö¥ªcR˜O“)'ÍãMîØLL1xÂlØó8NÃ6—u	étòƒAÌîÿÀµO‹á|Ö™kÙm1ir›â3\¿ÓÙlÈQkŒ‹ÏÝìw2ªzým¬&JçxÄd–‡a‡
äLì;m.áco/ìV|0)[\†×;D¯¼C1tG$è!m¾ØÝÀ¼Ôf…téïPo7ª×0)²-[wÇØ‡/§CÍ#Ø›£Ùgw;>¡&j¬^šzß®ïM®Xž&O]{‡Òw¦ûšKVá^ìÁgb0P›jªlãÓVöòø$ò›»i]lÁ(4JâaN®pÍ÷2Å1ƒËZè×„Ó~þ¤¹ÔBÞ±Øó™všEdñ´ #Í ùÔoñ8šïgð>Î²}²¤ÇI„áî¦}•Ç³»ýÔ^™kA5nê—¡Ûx,Ò´ZÅÙ…°µF•íÃ¯zäiy¿ä!Œƒ*;gè„üã<Ó¬¯aN=²T¢I¥SÅ»–È.ü0ÂYPäêa´t%oTLoü¸ËåFp"ËfM 8_Ì«ŠUc¬Jl’ÎAœáqO°ýÜÖÌêkW‰¨4'#¼Ô°;½õôÆ°ÈW–úÕ1Åëî%ÕçÄë¹²¢n„¸7cœDµ×W*¯À¬ËÅ¡ëŒkÞÈÇ­ž–gZQ‹Ö•¸À*7¬&ôÎâµ„¬ˆüÝ;èìžÂc;WåMw'˜¦!%qO»Fp0‹&{Œ¦0{ÉZ€nN‹OO™üŒ~”7Ðm¯€Ìë<éVSù&X§9¦s·³óA†c†Zû¤ónœé‡Nõ`’(™C—Êm[z,BmKjÊ8XH*ië«Ö:KÏÓWò§nêq˜Á»à-A˜Âáy€ª $ªÙ1B’±”ÌÃG#$ëv·ôæ™rŠ`ïœD­·¨:«»´…×þÙÖ—1þ‘$ŠÃÂ>ž§*ým™2ñ›8:§ìlÈ¹Ô,=Ì€ÄáÏYža\IüÉï›mËL[0çù®™:/¢¼ÀÑtt&(öƒ|Åy€þ²¨5È ôhãâúNþfS"ÖLÿÞõˆN!ÙÑÛÇªÛÀ–M…/©>)7dóÉÖ/È€¤åy–w(¾œeqøXrH^Cï÷ÙËcsÄPŒßÖÍ¦Ÿ@HN{AÅºFS>FC¬ï¸ äKI,£‚Þ¦à?g‹æ¦)LäÆ:®ðXù4
?îRìïv¼^	ÉAzI®±í'tjÑ:¤>y,¦r=‡¿¬0Eðþc0£ˆ+ ¯OvlvJ­ÄÅ,Hçå,Áh
Am]³EÉHæ‘Š®×7Ãç½ü  ¢¬¨ÛAÂ…VVÈ¯£d,¦Àq<áÈ	-/žH/$…CIÒL²= ÎÏ³	³ˆù‚fºá`üL£sú³ÛÓ+z	]ƒá êxˆð¿Â{—¯W®?……
Y¸c*šó¦j|OÖzƒuè•—Ýû@ÑW;Fá0¸´Šò~T”Ê£rž§:;v½üÍm@U×ï1×5«?Ãíà² ÃeƒW³ Wb£¼'Ý³YãŸA°8/=(fI\v;Ë¢CH†iqŒ³„øAOt´Ó1Gk‚›Ó³—´¬š£Õ·ò†‘¾ö¶G–Éš÷Û}Q†w
Û–ôE¦Çm‚ÓÅZªgC“ÝŠxšˆ“8H9öÈnJüw‰`Zâù*€sKHób6š ]¯ˆßkÚïÕ·:ÊÈ„ó0ËŠÝ¡‹Þ3Hm„4‹3>ô8Aô”æäýº+¿Û^\ìè£äyâñÛ–®€žV‰Ì2³^t×<hWø×VqWö}XSì‚ªÿý£Gß\ñW"±-ó±ñox» ¸L'
‚À‚‡#úSÀ°›§9‹AºÝo ž
ž&áUýÎƒ¸´È²U¿¬·§:­0š:9‰òÝ)ú¥Ìy©ÓM?jÒ!a(ºŠN²:ÙJ=%ÁUd&)U¸™~šå¸Éš‰ÙvÎ3Tblš@ÆP"ÌòÝ°ú[TìÒýŒ}f»e @°Â¯ò¨5Ììl×éÓãQáwéïÒï˜ÆuøùæŠ…ñÕð…Rý~A~ðÐZí¤ ‡E~åôèW¯CŠÈb\/Ú1¨0ÁŠ ‹à2>3Š(8zK½Ìä@ØX×±ˆ”y¶ý¦À{ÀQ>‰ƒ¸ Iœ~ˆèV«ÊÔñÆøóåÅ‘u³*ÿ>ô|ßU¶é›äÈ÷]9¨œ@úŠLAWô€d ýÔV7£áá²gH·)·† :zÛc‡Ý9Å‚ù`Î2ãPòæ¸ Ea‰ùÏ'yJâà\Ê$äýw£³¸ 1¯_fg•j]6måqß!ní%ÿØ'ßšð*S!®¿[ú]úzÀ";“çcÇYL&´ º ï G¤ÎÍH ¬èäã¿¿¹Âé¾Æê÷(ÇÁY–£v0›Î‚ô4":7+ƒä4*´ªï• ÂhŒ‡Žb¬R¥Ö$Ný°ðÈkî¨J¡T	 „ûû‚ÌŸ–Å5 øeSšã27”y¤@ÞòÂ3“4ž!R±‰µÄ®ù Ïæ¸õÕÞóƒÁ;‡{¯Þ½~¹»=ÚÞÁYž‹^°°gfÏt1Îá4<œ²^"2ê¶ET7Ïêb”W/Z!éÌöž‰f¢6h6‹J\ôÄc¨ÚW‚8Æ½lu»ŒáöéÀêäà{ö¯k¾¯	Jv_Q²}IðI ‹2Ç{rs{)#.þ=BEµ«ÍfÇ´S“§a+
{.¥Lµõ¹äD^€t©ö4°R&R¯p\÷!ßÍ	šQstúDZ!
»âŒå© gÇñïñ^hI`¯›äYšäÁ4X2ð°Ý¡Ûz¹*l{0Ðh\ŽGåÆÞÀ'FûŒt]llG¾,ÂUE²¶ö^ì¼ÚÚ=W¿ôñÆb†ø¤#µ«wÒ†¼¶úªPÇVC©„1šßešèÈ&òjß!µz3ÈôñgcÍÚõžšèÈMËO­§JÝ@¥Y(¤‹è¨}â¬6î!!‚ãºiç÷ïöÐéÚå!DIX!OƒÉ‡zîÄ^Q8¯-£Æ#¦1†mcë„¥Í²¢\fX@- :Ó¨ícÐ*ý¬@ˆÓÔY#þ5½µAâ’Ä9úÃø.TŽlÚh_‰ŽÒðï„’†-4+•EùjžÒhE}?€%àB…Ã. ¹ÛÁïØ¸ÞéÚ©w´A¥Úgõ¥Ùùa<E­
	 cœwT´d[d]ºþ­GÖÉƒU|tì#6§ÍêtÑ²Ó}¡±â}è)‹—L)äa#¶n¥ç[z|êaåá|ê]\Z^^FIELY£Š³qå#Q’j˜Š	²[û' îvpù” j5œU¦s%±>Y^³öLm¡¸U¡oÕÐÕò.Iª‡ß€<XŸy×DhŽÁ9?DxMÇ0¡ì@Šu6i:39H¶ÀYœŒÂâÿã>ÙÚ{þ|gëpwïåÁ`kô|çåöèÕîÞ»ÑÖÞÎA_Ðe®ép¤¾î©ó€ãZÔœüõË×kB¶Ãø{÷¾
þ_mºŒ®-°›Úœy[Öì€ØÞÄ¤ãÇ¯Œ`«8ýøÏB--HÝÃSæ•!@v0P¤“yÜ”Øî€C£]è<à¨üø'-,9âÆWÏ\YÂ+&.µ:…ÅV†–ø*fÒ[Åã_ÅÓ,ÆVl)ÜŠÇk°¼¸°«C­Aè5á&…ßjä[Hðõ£b•À(J4ce3=ÈÈž[ÈÀâñ™ ¸ivŠù~]I•„t,eÁÂ!R¶XßNln•½*$¼‹·§Åã«ëHªWœÖ´¨XmödQ²ê­?5m½•(-ž›"© '+B°öðÓvTµ‚žÞPºn]Ü¤©hÚ‹¡ìiE=Û"WkqTä?Î"`¥\¬ª!†71Ùc#‰ö›LðöéFynƒ ‡S2 V!6tžqqú'äÐe.TÕ¹É#bû´Q3¿ºþÒ–lîÚÃÍ¡Y.tr'8F×Ò\À.ÖJ7#Ì½“†~ÑæY<í“F±FÖÑ$Óø3ÞT yMN‡ðiÌý,Ôü
5rN'dš‘ØR¿3gú;C	’øÙÖ2ÝÆ‰……Ú»îªôxQqöäáß,È VØt6#Çq™GÒ~ýáóq†>#ôqs\ˆ3j0üñ_à'²nŠ…ÜZ¸2Ëþ¼ÏÓç3ïó·ßæB+©-8Œƒø"s°6X²I…†š¢a­dô—fìÝÂw`®{îhûD÷öÊÝ\ÿ{m¨”ËxJ‰Z•¨%Íæe·RËÜ§:[Ó,)œÑÒ(åi•Ì.­OŽ¨E/[ o{šÚQAˆ>£k‚YXj	ºå:m¯ƒ_ˆ¢ù°³°˜¤³3E=£(4a‹níÿò× D	nPÓ“.­‰ö÷DC­;[vªi÷iWÂUZ¯õ”úÂT/ŒDÖ‡žæZÊÒ€VÙ=¥¦d\ŒÂZxÑìÅÈ¯î5ˆÆØÝ¶ úZÈ5–bLœi§ÌÍ”Ÿ0ˆ?©Í¯Vw€¶àóña0¦5Ô»Qo'ŠÓqä!ß¶Ž ÃëÕ¾ÙÕŸÅÑ9u¹TÒÚßÈW³ò‰^/TËtÜE£^4G.¶²yŠjz(´›¤ËìÊÍÓnB×K¢ß25
#Ë§Û&Éêc£Ý§†@¨omØ­v™ÁóÉbí*Ð¥Ñ¹6ã/Å›*Œ’1ƒ;á]@·¬ÓÓ€P	Ü"O¼*O¥1b1´W_P2ÆÉYTd–J‰!9‚*ûŸOŸFeÀ¡ˆ ÁŒ‰
T†yvÓb’Çu¹fhÏXÂ”ÃhÞ¯æAC0¬JXõ±¼gA’åª–ŠÖh®½|„Yeƒñú’_ócÛ»ùûgäÂ
*ò•ñ,“‡ÈÐÛÐiOåoèë›¸ˆK1y‘yÅ@ïµõÉì9]]Ï¢Úù³YnÈ%­<YÂµ¾ÒÕ,nÇx\ŸEq¥9—ûÌBL«Œ$îpfû¯ö^ìoF‚«>¾{õz¤,–T½Ò’™—yxsüf	«wIÖ;Ú~±ûòÝ‹ÑÁáÎ+®	ôe{¶ûrôrkg÷Õž’’+²ÚæÂÔ/ZN±F\sIë§†ÝÛÓ}êÜse3bº2ßE ôkmÏ00ÿ¦(pH/:¿¢[… žúú¤ ;T“?ãÇ§À{uVÚ¥-ÕŒ¾mjžØv©•ð–Õ·Q=¯J÷• ›¬í¼|ÿÕsÒ$;Ÿ¾1ë™-¥ŒV¢vÛæKÄ³uãS¹}ã£oášÈ£C^êï´¼ð’9È†–kñÍÆÂFvÓ’É-´éÌ}³.Éâ5Ò7È¦.iyïV¸¦;æ^ÜÔ!#÷â]2Š×tÊaà––¼õYùô
¼\^Óð*Ž$é^¢*7ÙÂ¦:Ü‹ƒÌ(^2mÕ%–÷†b…kºc1¼ÞšŒ<6‰rXã¦*^zè³ËF7ÕB³Ùîã³›+¢íª,VÜ[‹‘‡id÷pçÞªœ|6„¼¼|›ª|à–»·ýXARŽì4M‡HšÝf³ïO÷<wª…½#µ%í“#­†ŽK„ê¥ÝÛ`F ‚QÄ0Ÿß°²‚Áè’ $g ã$./ÕÅ¦‚ÃÚ¤Æ=^ßGÛ•>ü›D;œÆéÊ	UEÿíðwÁ»`ñ4^ùH,Ø]oÉ¬Ó†-”Ùóì<Ê·‚"êö1ò.aTtÕUq3GO1ò´ÊU-\Éc­ƒXÓ“š½cŒ¢<E,ípsÓT‡:¼Xï± fºª–y± ÊØMÊ“¡û]uÓW»¦Y•M8:X­ˆ&5Ç\nYY¢katW;ñõfP³âÔÝóLØßIùN×ýC1ÑvƒîŠZ^^9^GæUóÜBú`ËÁRãr0´À
÷ZÌQkÒÃú†Y^®/!2<Ù V›œ2W®ž>UtaájÛŒ«q<Š^«“³"¾æ%f˜+ÁN¦ÍÚ‰:^Ûß^ka»zšÉ°?'|M3C¢Öê%õF¬o)\_«Ó‡	‘Kï‡YÅ¶#œàJÑóÔ2Ø1¬ÙeUïc»OÏ¹÷*xêR™83ouþÎ'ÒÍB7˜Lú„wr2!ß“.¿9èrV²*0¬Wƒ´vŸbÄ[6Ä‰mshÁ¶ýJ_v˜Ü‡xWMÕ
ûÒ#ß‘µÕUàË²gÝ½»†þ:«ƒàŒ//gÀ;½Š&Y®‹úL«ùÆxuÍºbN–ï“S1ÿæzÑ šû%kùˆž–tÕ…zák¬guWì%äˆ×sEðw¨íKAÑAWJs´b‘`Á”„ªÃÇ'ËÑ4Êƒ$D¤æ+Æ¨i–ÞoSõp6j	FFâV6§	WúLõ%D?Ã«Ö3SwSÊéNZmpŒé´„™”aÃžÄ’ÕÍ3æQøaÁ‡ÿ–ùÖÏ–éWíæÓfQŽAç©›Sñ¬©	KU]×=­oE–—€v}2¦}XmË”Á¯žì×‰_l\…=–Á`Phxe·‹Áb‡n³Ÿ×âìZ_B}FáþžEig1ÛÕ,º8ÑÃ2.R_‘Úš‹P—€iÛÑq0OJ!Jè†cÂ×Ê%z%Ä­]Äê'6Læ4T`”AÞmFÏ)œ7gÏlÀ«ß2*¥Vs#2¯ÖL¥¾C5•õÍyíf\†U«-Z4ÁrjThÞT1]Ý&èt|-è›^eš*hjÐZ178G	«2ÒÉLSÄ> YšË²DéyätUW8'Ü>2áF]rDI€+šÎº²¾kë_\lÏg	Þú‹„p]€¦ì«¸Åše^¥Ù¾™‚3¡g«Î÷Ù0u¬ÒZ²¯TRXð›š¹®¡’Ð†©K´ÜÕ\·óÿH¢4µÓo~à-õÓèqBÄÀƒjtî|NY
Ù·.µ¦êÛF­ÚÛÙ·:¶ÆNÄ ¨’´ÞÊasc1Ž8t¼Å-³%5:'ÐÄÊ
	l ¡Ï…’*UPŸBÆ0Ñ!4AS€o`ì<Ršî³í7=?è—<
t%)
kÌ™±¥_)J5¶¦òÄrøý'h™…Ý ~ûÃ×D¥ÿêƒŠ¶µ©&ö(Xñ]¿6Æ¶dÓ´IÞ{fv#Vf½ù“ûfË&¯¢cÜi)þ¡¹‰|Øßˆ±5á q¦ì²
â###0ópÜ@E©—m‡ÏóàéìÐPŽy²5ÐDö‹ðY9øu«õ|T
7ú‚Ò—µ‚mòlH‚’éÑ&PÖPëØË`ZÑ››të2;6Õ‰µþ•hŠÏôUÿ®§º$[Ï)w0‘'èyôcQ‘M¥™<·qXÑ¸`«lˆDAÿw7u¦òÛtYlò™ŠkO¯«Nª½çÑžƒçª3æÏ}’\s¦[jh5XmúUÅëVž°ÕuÕžbÕŸN5:Õœ$µ:#j<ù©:¨®92öŸ ¹'<ºy5õ-«v&¯=‹sf´¿cÞ¸2.Gí\I	ŠÚƒß¾Ú=ÜQky›‡“›C@óøÎ5—'SJ„"ëä$‰öùÕÅQY"N'‘8i1áÑ´¹Õm£üµòŽˆ4±Ô?'ó]™í‰¬Ô}»\ ìôQè¤›è¢|Cu–KN]GZÞšÅ¤$ µkÀNej^Œj‡¢Uò§¢Ìÿ°Þˆ;¾¡Û-·Vd…x“®s,Þëm *q¢9&‘4ì7 ”³YfÆ`MKíiþÊ7\†Ç¬Ûd/,°Ú7;øG÷Î¯Ö9—q© šé×úëµËYJÞ¯RÔXø"š_“¹sVªÚÞØlj=µ:ó„^fÃ»+òj	@¦.3ÉËL!pNZYÔSÊÌa&·	¼Ç™!y]êX]ª!eUí‘_°šùi¦3Ä‚ÎˆÂÞŽØH<Ùˆ‚&¸¯ö¡Yá¢¿S¦ÈnKžn¶žÌ‹õªŽšæ«€ÓUÞG§œU³šPð0;Õ7¿v9MkÀÊ+]À¨óºg¤àõcéN¡k¼ù×ámî±VLÒÐ‚¨k»!±~›ÚKáòNøUB± àƒÔ Ül…NåmWcå‰GÛ%“FçB÷Ý¸;ßå<¦Z¶ïe§åW¦EÔ¬	ª'è=/¡&…Þ’o®du×ïo5K¬óÍ3Ã˜¾ípÔœ(âdïò 'üú)``éO‰.9êÔ„ŽmÝ}Ë#ãâ7›¤¶Ù~’`vj$s±KtõV%à< C¨b<]y: w.fuIæžMGô7r2ŽéŽ´t’‚#“¤Y@£­b7ýøg™:CûPßÌÂ¤ž]ðïÏ¹9³´¸ö˜lÉ.¼³ï®œ.âf5Ï¬†%‰e0:Ìv.&QÒUëãõDAÞÑoÇ§?›Ý©˜yÂÄ²Ñ3¦õ_¾x¾‹!ßv’‰ÖgžPÏ¤4Æ÷bsp´úV»;€‰–óÁ˜6üè4ë=Ç"ÒTÅ8Y6×ƒqÁ‚nžS|íãÞô!º,”Ÿs'Œ+%~Ñ ü0Ka=½ÿ  ÿÿì}ÙrG’à{E²LÛVP…‹ (‡Q<Ôì!E.IiÖŒMU	TŠU•ÕYU ¾`æaÖl_vûilÇlÌöeÌÆöU2?0û	÷áqdV"ÕŠ™Q™—‡‡»‡‡ºRSó[&ïÁg>Ãœl!Qá„¤MÒ$oM&Ü}GGô¹+ m›OM×r	¦^ÿ¢Ïßÿ:}-~¿19³ôÙ0¾Ñ›Óý‰ù·²¶&LekPO(½8ÌÔªY{|Îé5dòŸj•Ä˜[]ÚAjVhÛC5ûÊ‡ ˆ_}åïÇîß„¹[¹œÝ—ÛôÐâëN_íY24¹é8	A!ÿ€Úþ­ÞuÞ(30)ä˜¬‡9špP]0œaf\Z-¸< ¬qVäŒ-Rbkíõù‘þ4X{î+öl¦T´;Ñø¸Þƒ´»²Ú-t'ë5§5Ã;¥½¾tÊÆ&Á_æ±9ÈF˜ZÖ¬0]%Ç\s‘€ÁšÃ'Qrš+¬«7±M`oƒ£­ƒs²¨/~ýP²ŸÈþP;ÄÕ¥&àÃm^Qv}JþnÐ/W×¢ ¤L¦…¨ëxS¦ ómº5e}í{Û`6BøÕqÊº€Q2°¶çÌo„rùUõÃ_éJý¾#÷Ô*pë+ãÍì]9Ú/D˜ Ò«Ù^sÚj¬ì’(³¥Ü»Ô¤wØ1F.ìY‘6ŠôêšzªCì'aœ¹}ëÎ^ã8œ(‹¹ô†ù¬ûö³+Õ"¤¥²'V¸O%™‚-I÷.Ù mAŒ’´Øx”™\å?üÁì;p(|Ù­/çþåèÄ2ÈØ$zºÎxÊRÞ7ä%½EÎx«˜)Ö[ùà:+Ï&p×Ïz,Âùõ[äZ²Ù¹†w]‹±N6šdmXÛùtgÓ¼_l\nÜ‘¹t¯6?Ï^å'Ù·ùûòŒÝi~¾)ÀdW?ô 1Û Û´¢†”uÛý–ç»%?DºÛ‹Q6v¾1£yŠ··¶²óÓr®eöusù‚›JÙ'–šá3ºZX#'«õUbîDÏ4v Áòš`¹·§'íå©”µä½fª^‹*kÁ>ÓW#´Â15ÞÄòs¸ÎÄ_ô£-'0A5-«ýiE&LÑTÁH·:ø:¯É™«žïf³ò'³í»×f^çšÈþ˜ŠØÞÍÆÛx!Y´“uÙè>Þ|R‚¹}híà‹\DŸ$ä$¶„<Å© ´¶²GiöÕY]2øÏa73BFfã}õs'i?w³‹‘öó]úÛV~ð}î“ñYŸ—óéðµ©Uà°†g‘§=jhkÇN-‰px%´ö[jž}(ñcÏ^úÍF£ÕMÊC•6ÝžÁB¬ûeÝ;þ£ô%Ç¬Œ×C#æFçžñ©û Ìî}Éa2D Víq Ÿác…|Î+5Öé¢&+³ì`™WÊ7\Â¥—Ùøý|´Œ»$ËsV-;na¥<rÃ› û#r®xEh§ìu5‹AÜ¤rôIŒ2¤h‡Òž;6í±ÚÐ¤;†€ä‹¤àÔq’Ñwõ–)ó ûN³)†r—´5>Ù¸yØ¥Ó<é€ÅXÑ˜‰Ó h6@áJB¨eµ'°9Üu&åá»Èø8– ÎÎÇ¼éwÅ%Á.bëgé…â‹Ù¼<…hÅóó¢˜0p_ÌØÌŠq	‹ài3c¤lµ@ˆðÈØ½=o‹n› -uÎ	nÉÿº.@ÊÉ€Gý±v-ž‚1ƒ•o–è>CÀG‰¨]q'—ë«NÖñ#ÜÕmÑÕ„ÆbtFŽæWÀˆ„ÃÍõZó÷žzà­L¼²Ðìœ.V&e@ öRf &žŽ‘–ÂgýG÷N' a›¥3d½CŠÐfáhçÿdÍ¿Ê³ù%á!WWÙy9€|ìpî6–ëmví«íArïbx^¬9ÃC>tyÉ'Î	žä—Õb.}mYÂ«rÊìD ÜWÕ4Û³÷ØÊøµlœ‘£fÑÝZ'ý01OàóOCtŽÄ8? ž`2<ÿ['Ã‚ÎjçÆ ÍmB4ç1Ñô>0~®]ûõLCcé"B*ºUá? ÉgãÁ>ý»®Îáoƒ@ÚûU7Du¥Œ”\ ÄË®;ÝØÍ<Š%D»Š(™Z.ôœˆïÜF@tÐµ=î8Æ„òž/Ñ!eÙóQ>©cp;Ùî ]O=¤fO€
a–Ph²Ø~YÐˆJ\¿_dÔkøSžY”ÓÆ4RØ#/Æ×ù4…	Ç±Ê Nºþ,Æ¸Æùº£šÅUÑª`Ò"^4F4ÄÈÎT£œe“dGÙZÜŸ»ZtQ˜2ÖP¿ÎÆ™’Nî¸ÊWQL%ìm©„µ/6F\éÁFN4#ŽÖÃTáŠÂ¥®o@íàQèò»|Óri`ð_té>Å~‚Ýø8 «çŸx–'ÐžÕv×ÚŒçNÌ<§`ìGÜ®ìêdFÑ|ywAë~u,ð°¬mTï!øåŽV^ Ãœm)	Y>hÊmp!ÑùQ~RŒ0ím™80ºMo ù§gAmÛ…š{yÕ_Ô³ªÞ˜Vý«ÆhüwSe#Œfù=³Oj>(ÁúÛä EXR9*°}—÷ûÅt~Øé]Œfëüƒ}F°”{^éÖ aÔá²¶ó"ÑÐuk¼etÛàØ>Qì×RåµGÖÆN&$ìÕù$ØŒ}¸€oŸ½Q-Èn’Á€£Ž'Âº~3hhç÷¼¢zÂþ†»'ð×Û[Ó‹7¸lÌd1…ƒ¨^Æ'0¢ª.x^Ìˆ|9ËëDµ`V¤¿ò}mvä-¼gò“Y5Z˜ŠÓ9Ô¼"ÒâæN¶AQŒÎå’>0'‡íY‰Hñ%ÐÖÇtDvÅ »¨©!cV-˜¥^¯×Ã*ˆ;$n£!ŠÚtÁ_Ê]iBN‹HBB?0%ËM6{lv‰wÔ)…`T1ÛKAROû!þ±Ó~ÊÎs0ùv äö½O–²î—†÷!%¾Xz\§)¾ÚNKwÉ°%wr ÅWYwŒ¼ßðŠ#Û«šR«-óèUŽ^äÿ	Àg›ìDcêËÁ4¦¿Æ”÷I•¦4PËµXö×‘—¾[%ÿ 1%Ã‰Ëõ˜úøÚà·C5·ô\¿:¼ÕÍL€8„Û÷ê:¿ìApiƒ…á¯ŒHC{™dC¦@ùºªFE>Y[sO‰˜žšDßDv—½¿ÔÞòì+ØY¸*×ŸOqc±ÿ•ì+=²sÚ¶Ò³øY»JDhÿä7cl¦~Ô±*›N=GÆÏôV?›#ýWršãÏ‘f/—^_:Ý‰?ýµ—Û+)—†Æ1´-V³óP¹îæƒ(©ÐG¨‘Æ›ÐW+ÜµqM£µéÕ\:ú,;\‡(nŠI‘Òóµ‘TÒÿ­ç¶Ì¾Šà §×5¢td°ñëQ9Aõ&8RŽ§ÎÉŠ_œrûæ=ž±Ã8ã»cT4T‡r25–æ‘OlkÇ†'
7¸D®Øz±í›¶Ôæ¾ú•,´«ÛºB/© R³ƒA³RCå¼cV¦¤o7*Š‹ØƒqÅ¢É ¶#zñ„¼ r¨µÑö~(]é—tu$ò¶ËX
/¾k:„º4{šÓ«z>‹*†žÂ’—†aýævà®ÊœÅ)C<òTÔ;f–ñ¼[a&èÔÓ­‘a:Þ+µ·ærÛë`ŸôbÎÖð¦ÞÂ¡Wké‚å^Ô".`Ùåy¢Õ¤q#z²LÅeýÛá3h§¡)ä…‡x¸…²^«Î^Ó®ÙÏHÇËà¶êYâ6“ìàW´_†ßú½žqYíÃ)¬×²Ì6íÁ/@é§éÆotëˆ«)ŸõcX`‡’(´Cq"µrVæÝy,~+VXœ-JzÈ\ÙŒˆàEW¥q´zZ¼ð ¼ŠÚ×-z´^Îë" ‚»h”t)<J-LÝt¸¢ˆ] ñár-¾?Wó‰h¹4b.7ŽÌm+Üa–í†?ˆÌÔ“è+¹›ùÂŸ+ÇHÑJûNì	l
h*È$RÆF§§m7$uª
‹'KÅ›‰%PÅŠì)%*V¼©F½¥&XÅÊ1ÚÈ
–Á“Ì$6hRÖø$Ô¬í¦`…µ–4,´5ÒÌ·:‹¤n4Åf»m¡	\£ñøÚXshW¬X¹y[	É\±‚ïæFSº&4G"-Ò¾bÅ“k'ÐQZ®k¬#M¬ ÿ½É|“HÍœ†?Kv»‰È ìy0×[®c²+ž°iÈï{c¸ØÆ;¶EÙ¶æMâšèzúð{ì€¼¾#Ìî\JýZjþ¾Ù1B}*D¼rp!Ä DÏp¸Îb¥ zUç³á
`s¾IÕ†(÷îW:2¾a]NÞmà~Ü:pIm®õž‚áãôrcK‹×Bµ¿JÿMÖ ÿt¸ç¥Rý6Ê9è™ÿû’¬%b¦­Š®fØ³µ™¡]ÛÙÚÚ¼ÕniXkü§ÖÂÛ[‘Và*Æ?}~×¹« KóÜçÑ&QïÀU ‚Ú±†RÆº# p}Þõ!ßú¯åh´ÜÃ~õJHmÜéÑUlyÄ'˜P1ŸW½ZòL^Ldâ9æ¯Këu½Rà1Ï£‹öwì—Ñï}‚°W)ÿÿþWÿ{%ò
Óp^¬û²¨i–µ¸prœîl­¬ð» ô}¼>±,â_›ÀXB^Å-1„•¦xÂJCl]5ÄÑŠ9¬p+3ª€ÌÊÁEÀŒ•ëe…^Ï‘¦üâ”˜ëÊ]î?Hé°&•°TØp(gˆñ…œ¼ÉúûHŠœrû0Ó5UÓ¯àö½ÂIl²;¯!sï™Ù{·ô4Ÿ>/'‚£`¾¸üC
AŠ©V8àöŸ.
jáå²©æÍ.¼‘;±f†K-LÂÄïèyQÿü/ÕÀã¿e }åz‹ÈöÐBtôê”ÞøÓðÀ`åÀãý:7kÁÐ	WÚ÷Ê1=ðºXT«fy»‘}
izíú­Oàsœ›5bü}óóìi1Ï©Ç4è¡B†_
k¿ºÚ77¤Þ¤c÷ÜBèýx4Ê~(¢u¡7¯•ZÔšc‡£¤W{--" “¦IÔÞ]F­‚‘8`³Øí@ŒLãÃl³Û \!8AŽ3[°$‘ !ð_kür+ vÊ­Aˆ]{¨Ò‹b¶ÍY’£ij¸5Ã„ašÇMÍ>‚|`¤É2~ »4Û‹øÅ¢Hõ£mb¹x%†w7ˆYW`XbŠlªµa©ÓÂZ›Fìý¬K?TI¨YÞ÷ÆG=¸AÌ’O+åbD‘Ì.mÎø«Œ_O™ýX1Å¼-hŠ+©×[5°q…å×:ûhp%,°i‹Lfkœù7üÆ;]Š/|ÞÝc¢5ä¸aÕJ¦éáŒÓy+Œ¿Röž[ü°ŠWZ'À³-Ç¸8*²ã,à”ó|TöƒÿÛb2\ŒUBÏ‚B%p9†Ìl%]O¬¿ÐŠé¨ŠBêZ¦ß‘-³>³³g>;ù2zþX-»(ÿÁ’Qè=bº%T©FØ¶g!KN(ºÁŠ•‰ò¸÷zª§&u8HLAJÚŒº~Â²:	weŠ]#áVŒ¤þŒÁ'
Õ¦ÁÃê4(æ­¢™pÀ¿Âaôzë‡íéÅ;äõÙIÞÝZ§ÿ×ÛÚY{“©=KXá&æVhYzœà¶KoU©É¬H‰,¤%¶»Iú°AD[³tÎ7îdCÍ1^=ùÄZ¬‡3$y™3TEnïäQ{ì»qód½°ËJ4˜Í^ÀXÙ.Kx½¤­;¡ï=NžÌ±NÅiÌp âY˜×÷æÝ­˜gU›WßÀÁò÷%W%p?ÄVÕKPea}¼€JZê´ÚTŸ y	·˜ãî¥xiµäK.:ÍaÎ¥œ,!­ï†Ü-1-Œõ5.Ê~)½F­¸¹ãbP!lI@Zàtl,OÕÌ1MýgçËˆ}^L[¡ÄùÍ«³³‘Ìo>ŠI_X­S!&mS§´Ÿ2Ž=pQGdŽé%µâÑnìÔfô³c¨"Ø”näD£´Ïìør©Ä¿9‹ð°/tó,ñI£ä/a¶†„\KœlŸKÃ¾+%Û‡vf9 ¹|^þe‘×…¸GÜZ¦è%|³zjÝI"WkB=Xužçõ¼ì—ÓjAÝKî-àE‘ÌMS;ŽY/ê_¦Q4mË)£e‹¡çâæùõö6%ð€7æâHÒU¬I‡km¼›ŸgÙñ,£l`”¸ò(>ôÕ=hË¨9Œˆ°o‹°éœóy~–Ï²iÜ…B¤™ÐÎ×Î-WÝ×â¬Üyñü‹ÎECïÌÌ¤Ô%§ï¨-
2•ŒYp€UmrFBÏçd°©’‹(­ÄUSˆ ®)0%I3˜±Â‹5ã%¼ µ©ÆAÔì}Ö¼f‡¨TñI•–~Q:˜ÝéxàMiô[ŠIªØrÕKû;ž¶X¥h.ÊÙ¸éNƒŸî,#>©’(%ˆÒtç6'éÜ^•ˆƒ§ñiú‘+ý´Ÿ|b¬ü°•iâ…ù‚eÿ[àêrÖ¹˜õdïø)¬¯¿½žÝYÏî®gÛ[6§ïëÆÜ»ïnÏ¹)ß"±ƒBe)¦½b–Ý–a·c×6gÈãžlnaÂaöL:l>ˆaÄËÙ-Í¸}3dXŠÓŸßfê’_€Ó2åòÃ¸—´]Žuw¢Õ1ù›fñ|³=ô$ä+`—æ¬}77ó4ïuî+ŒšÆmVì2.'‡íF5òR£Ùýfâ.uzm²¢Nh¸ð-¼]Øe4JôÞÛ.VVjÔ|wg[-vÔ·tIí wákw»¬”}µc^mX—Ë¸–dCïc‘pìÒÎ×j·1
·i†mÁQ|	5NãÅo‘Õí™4`«ÅœúâLªI¡Ç6ÜÖøv‡‹‘náÙb^GÃb©’¨ns¢YÙ‡Q&”ÐN¯®=ÛöÔkjÿl:úýï3ýw+—EË$zW3‰ö{Å½Òs	ü™âè1ñHñqÓ“Ì¦ä4»1c7šYŽm„bH>Þ-ÌS¯pT‹©c„Ðç Ÿ±kV÷¯ ZaLÍGóÃ«·ÈwÙg ìÙ¶Ï·­Ê³ŒV'?˜ú •ð~¯‹Ó¢®	ªÈyì’È,Õ†xª$>KV†¤ƒˆx~#øø–ÞÓ³h>ìüp2Ê'ïb0ÁÄ«i1!X2©R €§à)'³b¾±Å’|æýw›··²jš÷Ë9á%Ë6˜Ü-žZYÁøã¸’:"„G)2Úå?‘J"cT{O;éT &!¯°™ÜX‚ôsµTÅÐK'•Ašˆ8®È2Ëž±Á²\¦¾þq?ÏÐL;a8ÿ§È@/Š`¶h„~wÏ˜.š/Ë~zñŒàæ‚£-á€pŽFLägá,Yi’š*QC"ÃkCGcÍå(jë‚jN3¯Á.†É ¦V$3äÓ¤ ÎöB°È¬|a'»·ÌºJè‘«å×F™v'–2Å±8ù"$L/ØÖŒ7ƒ")Êí21nµ·¾1¤¼]•*	EoÅHÏFÊ†ã4Ô&!	Òž”¤YH^’PxjöFÅ·ßºlÔ.oøŸ œÑæŠäØ°×&’ÚìléÄe$vçPiÿAc˜¥i/6ò‘—©ç‘†’»ÞÛ#;¤òmŸøâwÇò%)ºÞ0§›<ã÷rÅ„P3BK	ë™Ò	z`]È~$Ô˜´BÝ\Ð‘!,®µðÎ7®­˜½JòPÅ-ÕR{ºÉÏçùÉ¨@N/,¬Qq*Íœªò)šÄæ`>$ºDóÚç³Æ›=q§‚Ï{Ñ„:t<¦ßˆb¿ó ªsÄ=—73Òá`s>Lm
hŒ§YUX£I-‘-*½Ž!QÈ¹×jîÿê9y^chíYäƒùI5¸Äjš9ƒµæ»³‰'È€¢#Š`óÈ'ó}þÐa˜{ˆ©.Î½ÿ|à¬`P%ˆ5•EŒ•X,1Þ_Xg9¯ì‰‡ÇhÅäÙ‰XEæI˜ Ab”LÑwK-[³è\ËÆçj¡«eŒ®ÖQº"qºZDêÊbu¥Fëj¯KW+(O€°°SšGÃt¥ØÅøCu­t;„ùr8°Iz(§¨æ£u§,{›-Êi•4$ Xa%¬^a¥Q‰U©ZX	á­­v	+]%Ór4¬Z:Ü×‡lù]ÜáGlƒPÔ°Ò^]ÃÊRJV–w{]2‘šŽ(Éïv¹”DjHZ¢D¿ß°®&ˆJ™V¥ªA›k«°Ak£¶Áj¨¼áxU8âýMÊŠ!]ŠtçËÒ	–¥2CAi“Š×KÍ¥a–((KdŠÕg‹‚Ò8c”%²FAIË¥aö((ŽP'³¤„¤*M3QAi
ŠÈH•0:_Òª„ª±<«þJ$ÏÚZÏvØ Ë”¥3]Ai•í
J<+”FY¯ ¬ ó”–Ù¯ ¤MlEY°XSm2aAY:”F±´^—ÉŠ¥ef,(iKÔ C”¥³dAi•)JÚ”V”1ËjjÙ¬YVs,sV:Ä‘ÜZx¡6ùÄ5BÔa1e+ûi¼ÉýÕ ";X;$hYi­#d%,+Í`ÅÍb–æVyÎNSrù»m²Éä<Ó:lŸ÷J«ÜgPÒ(Z£hPVJË\hPÒ&Ö:'”p^4(r£Ñ“í_Á„U§ß1S£É›7•»9íæMnP	gÖ|‰\é)Öè×	þM-ÎæÓ­Aq^•«©xAï–€øœÊ–yzÊ6úyæ-ï9ð»nß•JSkÚÞ{5 .ÿršO¯ö®ÕL²äáÍ|ö<|0©V=üóÈutkVVdç#Æê7òmŽ`.~‘O1K
òLtÌÇ®‘a»{UrÂîâ¥£,/Šn,/M ÿIÞ7¨«)èìŠ~êÅL!z ìâœ©—B‚òÑáÕ•0ÑßÏ¶Ö3…Ld¼­Þ—{ûÉ'å˜ŒÕ¨µ-kmcUÌ_f@ÅžÆùÅÆ9˜ç"U;ˆW‹íj7^¹õ~øÆ&‹ˆÍŒµÓ;{¨x `fI,_b…,–Ý—ª®Ð5Ã dl—ç€`W‡Õ{‘u_õ¥³±ËD¶.Ìœæ£™_šq”Ü²úž¯=2Ÿ&¹ÆDS#ÓÆ1=„ã¿ÄÜ:üLÑCUàæi¨š¼\œŒËùáBØ/lêÊŠg·Õæ¸dlƒîÄyÚgê9F#cÓÒóQÁ ±=.©ä%›ËXš}î!àtž†œÑëâ/DFò&øåþD\œ·’¯ÓÍ«ÐtËÖùÅ‹å³}í»N0±^àà "K,8÷Ž¼õÆP0œd^ìg_ÊùOdèÙ}Bëê*«bÁOIœ/¨¶Ò“68dPìAÅhÊšPâ6¯€p#˜œe¯ÊiÅb	ÝŽÅ<Ž£8Øjh\ü§”v n´ŽXu‘…ûà7/ø×?ÈÔÅátãÞ—I™±—Ø1º§º'VSÀ‡& d›·éÜB:áÝ¬ÎaxÄþU_Â·MÚsC›Mçü¾œô`duŸMõñaÅªw£pLßyÁÛ„ŒÒìX»$¤’õl²C¡Öœ‰ñÍÒyÈ ˜äßÏçA´Ž	êF‰«&Ž×ó
pó?(ãOP	r[V)¢Û“m“³xQ}ª4ThD›Š4±o&‹þ	(WM­ÐZl÷¥ÉbçH(ºÔ®á=#õZœB^9Û" I™üJl‹0—§Ñpt·Ô:6¿'ó×5ô_ÁVbÆ†<³è Ã*hþa`L%˜‹Šhè 5@»£yêg$2IhI©{Zì_u"/¸º6ÀFmÊNSbvåqC»PèVŒÛÿâS“¼ALN ?ÿX·ºÃ±ÌXä20ö+Ú#‡LD„Pvs¯Ô®›û™ÂÈG*›ûÈ—rüãl).ƒOÅ§†ÇdÌ7S£Óß0X½»!þcUÿü×ºôJ­çå8€¿zòËú=¥uÁßz©…¦þg¿¸Î™ùÔÞ¨Ò™[?1­óó:ÿùŸr¦sn£\F^|`‹Ç­
£ô•×…ÏAÀA¥xDª–ø:Y%R;Ý##•ÜýB
q‡á&y{“ðõYö²£Èþkv
FhÇ7r""4úa7‚ÅF¬l3{ôàûYö¢˜MÉáýç¿¾/Ê6¨ý”ûË—!Õ+;Ï¯Ãƒú-¡Wv¦yˆ!¢÷¨ÌÊÅà{À·YÀ×£×H
pl Pq¯5ü=Î^{+†m~Qˆ‡ (`Ù,¢#“
jZã;šiú0ó®*W¾Ï×ÜƒR;qTðS•€É*hçÍ›8·(k”Í¬µÓ²žÍA§”.ezÛcýÅë­7TC=í`’Ô#Ú'}¥zõ66*æ{/æün†^_®Küã·¢F¦à}`X –5†µ–pýR,q‡ÂnPÄý‰Ñuë{”¨n¸e·tì’ÛU¿†T¨aðe
„0Õúž¯t¸†ÿxH]°r1ÆòÝá›;±»{’4±óÄÑIArEæ³bîèPeä†-Ñ{6k‘C¢žoí½wW† &|æ
°(5`@VÖ'Âå£ÑQ;¯*ðµ³«úê¹b¦¸þúÛ4yZÌs0¬Ò²ëuA±l"4wH B„sxÒ\g?^QçâvyAŽS>ò%þøÄWˆÎaïÓÖÁŠ¨œ®?Ù­(§ðqoÆd·øß¶c‹í¸ãÅš–JçÖÜ63ÊU4ÓÏ„±H¯ªÃ•†;qÂð=8²gJ¸¡É³çä$&såOàß}ñÙÚU'd»Â3UN ˆ›H÷¥Lö£,ñÙÐÇsÈôê¾¤Š³~>¢F}UFØ§›Ùíázv‡üïîlŽ··†~¬ÇèûÈÜl^LÉ½­@F3‡a>^RØ(üÄo¤Ð1›k@·÷z~ã@LI¦FÏêŸÿ	6û«jNÈP¡_åæeáC>ò­›`å·»ÌÆ}ñYFDøõÕm^¿LÉ¯_diKâ¨†cŒóƒ'¿¶%¡Š<!ñ‡¯Þµ“Ð.µïûÃ¢ÿî¤ºð¯ý¢èûÖ
þôñî]'~•Ü=|V7¹}8:XÁ%´<úÄVvNgHþg-l³mfUˆ*¶*Ê¥îFœèŸ3Èê»¦˜·”>}ƒ3£ºK‘98°{4CõÏçuy²(k-öX6žö­e¿¤ Ø0úÖ¿}ü\’C·Q™;Ôwó!DsB!„×k(<§ä^å
nTV²y÷ß"Ÿ—@ëˆPjè ÖdÁ>.†?Wj‰õ»R#þÅqðª(YH‡D¡$ö‘óÛ ŸAl‡*›³9µ‡x_àM5&0-#/äˆ7%„£_Nó‘?DXn&«¾t4p ä&´\¨šm“qE’	x—ˆ^Oyvã¤°ì&²da çEøæ˜}ïaQ“2 ˆÁÂõv<BR4Æ÷‰÷œ±‡ÆÚÙç)¡P©'‡>­”dÈ÷²¢­]r"ï•.+Í–•ÄtÒl†,\<\óÒ¹&vqŸkÜúÃ.Ýr@¨Ð bm¦÷%-¸­(º ç™p“”ßí¦ö&m]f%]RUÅ(U TZ3ñÐjô«”}Þ 1(âäèÂ#BÄ˜¡GLÅ‰é=Nó|9¡À¸”˜L^ñŒä©%]Ù¼«2‚ÁyÜ¥¿‚¹!ÍÂBî3wP'’’ë)="¿(Ñ4çñ´/PâaáÓâÀ›-&B©ÞA!¸·GS{ƒ‰YB/Ë¡“ˆVP£½ÌÔ£©˜%Àa/] $ ”´e–Ò!Æ`mHŽÛy}oÞÝZëÍ«ïà qŸ º‰}F"°Ÿ²t\FçÑ|\-zÂO<ó~bÐ=½Ó&£¦QW£ˆ«¬ÚLSDB‰ÅÛ.`˜¿]o=´NXÍ÷ŒšIç#HvÇ,š» ${¿?ËòSØúäéBZnNí§ÂPù€ÃÛÉGñ/ŒÔàæ!Ü‹Y/
°d¤ÓU\©HlƒÍámÿ¥Hk¬¨¢ê†Œ?²ì	÷Í"Ap#·)Q½jêJüNs†Rá‰W–Úª‘dÛD¢mÈ¹ÝÝIâíI»û“HÌ„´l^7|‹bÞ£„WÏÙqööÞb^þ+¡”Õ~öÙh|èºOÆ?¿~i€ˆ%ô66ˆß~²ZaÚS(M8µ4Èˆé²Ø•,¯4uRe3ÂTŠq¾ŸÙZFS×‚GþräU˜·~ÊÖ”þñ“Ø†å#²Ÿ ‘åÛ¨™õÚåBùõRYT˜Î¶4¥\½gª_÷òˆfëæóŸÿw¶›}÷âÉÌ'è‡\Q-Y?~·¨‘Sš»ŒÝñÀŸ)	”£Ô?%s²±_ß’™Ã©ÀAö©šý!Ûö§Ï…Âº	ˆ\ƒ°ËpúÉ³×†ú›B*¢‹†Š¯Éðß4
€	~À“ÒÑA…ÕÆAuñMÛ>{I•/ X‹ø6&rÌhØi|¨Ó¢Ûìê Q{ÛÌ6ò…™û"ÉXï£3ñ—HK’´Û’þ2½‡ ép±Äc1ÐPé÷ÎX–„`´t|oˆäm??ØTöÍwX¶ õN¼àåÇ]¯÷»ÓÅ„Mæ¡í—ú}YœSyB¤°÷hÊ¸¼_êW5yUå³9ûó›bB !2>Ã3jœ£¾^p?mQ÷%ÖßóùìÞt
ONªùýjrZžÑ'å_ÅwâWÕšÖÕi9*Öw½Ÿ‰Qî‹YˆI¼¦DüxŸª0Ø:xòDd³(«{ý¼bïøTö³îx?›ÑìáëÙœ|L0¸ß/f3K¸¨ëªîP²õ¾*_!ß§±Oa8Î¸Œz80øCÞïØp(¬ÈKpÎæ,Æ©ƒ|Ñ«I¡Æ<&cÍÏäÚ'XŽ”³â ú>úJ5iàkñ÷WØE+¬o|sXŸ˜¯ðtŸ^£â¬ÎÇßÌ=-û9Å´Ã,Ÿ]NúœÍù'¤±éŸQÏW>úÞš<Ô˜ø0ßr›Xã–Œ"³ñ0’ÿ]="P³î‰ÊdòðŽº*ëÕòiùwÅ%^ç{§Õaã=)·fh‹È[½QuÖí¨ ÀA² ²Ù¾ m/jš­CuÖëH¶¢O‰íåy}iô@Z!è@=¤]˜ôüÆ]Ùû¾f¡6ÀC:?ÏËyvZÌûÃng“L{“bƒ¬ôÅeg]ãÍãb>¬È"už?{ùJ»"’ÝEPuß`ãWv^ýŠð€©”Oi&@„ÍgÕDkàZý	‰Rö³?½|öm-{yziœý$ˆõã‘odž±‘Ñ]lyn¡ˆ-´^ëZÿÁ†dÖgÐ6OhãbøJhÇ\O‹×òdwí®	Ø'–D¬Qà¥¢š|W½4¼Åëö8©Òýê
RÊ¥‹G9Ù½D:®èžÍ$zŠ­g€Rô@›¡¨ÿÝäÝ¤:ŸdŒ$ªoåH¯³bDPÌ
ì†·vw0ˆyF¥Mßë·ª%†üA þ×ºv·›_ç!_Ÿp~È“#;‘8Mf#^Sûéš|`kðzVR=|UÔãu_ÊŸ ?šürNè|·Ãv§¨C5?¢V&¬–ö@¯wÀú¨¾áùÌhdÁ(®ÖÊwú“ÔffÅÙ¢¤©ÛõñX½UfcT<£œpI­½ûîó7Ü4)eŒ§ƒ÷ZcÄ¯È¨Týrö”´9z6-&ë,CŒüm¬’JÃ+r1Ž3åu-gdÁâÝ„ûCJ™#¶…àÏõßqd—°çQ]8óä°7"³|ýæ¨ûúÖÒ}LÃhL{ŽCls3{J–*@þ9àÁ€ª4hÕ‘ŒÆ`Z=n4ì§ùT…wÐÛÓ_-R˜²f9œÉ“#‘Ã‡Í¤«÷Eö*?ÙÏF%éïýd}¸š|ŸÏšV5}HŽÎjL9­E*ÑÜ¿ŒÞ;Ð\Nå;Õ ýIšbÏ+hœïÑY‡#¯gãÁ{"A?GÆ9íó{ùÓì´Ÿ×Ú6´BšdtLg²˜Ú½3Vÿi1®ºº!7‡¼W×ùeï”ˆz‚Ì“#kö²˜wAHff×6è):±&ìã¾®Í'k>C¬0gVk½™>ãA×ëÙkhï\¶}áN˜[än>zð=`åT`Ž˜4ˆÍ„K3½©ÜßeÇ–Ì]ÆQ_ªÎAÉžTçd,P‹Z ð4½yE_0
“ï–³ûÂ”øeQ¿§µ=igL¾ÒØ¸cJešNFv®l+;„Ò‘³‚ûHµýTØ¹™_:'?Îãø¿|ÄpÍM:Ô÷3èbÕé#¾ìäï‡§§d-åÙÄ$²ŠO#´Ðú¶'ö—ÌÄ!Vý+·ÚÍ¸!ˆ8 LyP¼/Õ‡ŠÍT4î¤¶ŠŒˆä:Ãb4à’Ó>Y”q	¹ÔN!òí+A„ä—tó$~–ÔNC4øR]€½ã½¹8e[ú™†¿¦ƒ9ÌŒ
ßÉ7qE­jé¨»ègÙçÙöÖÖ–Pî{>_Ó{”§§¯J‹}½Öƒé’‡‹6ôñ©çv&Oóù°w:ªˆÌ'ÞÌº0*2¸;ò?;·×Lü—mP¬ß"ƒ¾²¤A})îŠzž¸ &ŠÜ2`EÜy8Vüª#Wxœ1G6?h¢uØ¤_fÛ{j‡ñ•¼=*Þ#¸í Ôáìçž
¡ÉéÐ„úEµÙäÛÁó	Í>…W£F¤äCrbÕžžœ1›ººšQ#®¬¤?¥®’þºkØ•dÒ÷y‹êÝfóZ?ùêSÿ"aæ÷È!ÔÎŽÏûRLµªA_Ì–½ÐæËÈóß‘9ó¯’f½›0k¾†©kÌø²ªsdÆ9ÜUjf¿å|ÙÏÈtÙGÞÙZÓR“¢f§²Ž>%5!ó1PùÊ‡b*"‘µnŸ¼mØ'ëópf!<ù×ŠxzÿGJâ)…”5a"£Ü"=!-›.µJŠÃ25~G?b§H¥6+•†({mhùÈtíÓ$°NTýî þGDáûÏž<yxÿÕãgß¾ì=|úüÅÃ—÷^þðüÞ‹û¿¸÷®œÖtÝAÌöUïj‘¹¥ý½ù>õÐ*j È¤ÚxÚE\×Úíð9’ÿðÈ!·³V‡ý}Â:/ÞËÚpÖÏòJV¯ù@{äŒ/4µø9_ÏøL Låc<CW"õ˜Æíê L dŸŽaØ_/fý\zëˆÌ-f%%×AëšDÇxèdú#*¼ZÕœ®Ú1FÓE0Y–È¡m—OÑ,]Z§u.)Éæµ…Õ¾>¯h§YÕ›Bæ%O,jn–^Cö_ióÓªÑdf§C·!{B4!ødéÒ™MÏêNÇï	†IÝ6]°…×òc²Ô:“h˜(Ãyê(b>+k}ªÀg¾0ád¾#SaÇFJ¥|ŸÅâ!äïA>žTä¸ºÛÌŠOnÈÎ–;lÙ6ƒÑ»i	¶%D$_»:~úý}1)!—Ÿõµî&Ã«º{ £Z`Ç{´›‡c)¢µíÇ	ý=ÝÏ'ýb´Ì|DN¾Í+–Þ±}7<‡kÄÓÝÏ„t§tân¦Ôµý¼fT½Uã´r¨õ¯ëjòSÑ¶yVÛjßbÿ/ó÷…bûÅ~ö¢Èûóè¤œüñÕÓ'ô×¨€­y¤±«<åá£ÅiNh…É›à~™³<8‚=â?»d¨‹º&µ^Ñû—ø²PÕˆ2F…0Ö#X‹lŽS‘G×ÃöêÝë¿¨_Š7šÕŸæ—£*PT2Ë‰,æ§Énžu4õ–KøeJNåL³Šö
«	·m§ôòÕ¬&ž‡ê¼ð÷Š|‚µTŒóÒ©K¢_ó”‘Nþ«s’—DŠ²k°§Ø÷}ž×üž=Å¾•“wOóéÌ®!ž£}è›ÊéJ‰ÕVüß®ªÞ ½qÌéŒ<Cû0¥e«ý¥&/óQ…¾Àj¹ò•
Ñ,ï#ä©HnÈ}ÿ.ÍÌ­¹oþ”ýFåþk¾]™’Ê£ä»˜iŽ˜|û†‚.8Ê"Äy”³v¹\Q~FF(é—çåQ—Æ)ÉÃ@^kÑyF[Ðé˜ó9î•&‰ƒ9kÃCŽ0úù¯Yqš$!å†#*ìEjõŒF¾ ƒrN6¶}€˜o«óÇMF± “GÔ¤í•£©©u*¼8èÇÓ¾Ø=¼ÊäðŸxŽµÆÇZ­äêZ,	@3\½XCíQ–½}F2p_‘í[f%¤;¤á'Áe.Ÿ#®€BW­úgW:^³ˆ<¬…¶|?§Mäp¡ŠÞ[Ó
AÚ*éE˜—Îcÿ{ƒÉR["xhšÙú¹h0;#;îƒOÙ´·dXb·ëLik â:(Â¨‡yû4ˆ^v/'G06OŽZÄg†ÕÅ&¶Î'aHûAŽo¢•:û¦ š\d÷‘<BhœägÌØ½¾3Ì$Gãa,ìúˆÒ‹gO¾ì}óðå«g/~¸ÿìéÃ÷ß{¢þrÓG+{·hìºÆ²êTNÇF †7ô]êÎ‘¢øàýãÉieJ¼Ö—ÔÛáÏ“Ïá:UKŽ4Úÿœì[4FÜ°Ô.ÆÎýü9«Q]ÿ·AHÅ)ºß–CÎrkŸÿyòçÉ³ÑÏ]Ï>‡>Ù\iŸßúóä;Â³&PUaµ†ù‘ÍÞûìŠCázyN@š¿¯êur(÷Í½;_[w×Ûžú
ú¶y‹%#›ÙY4º—\7EkHÑÅƒÕúQc¬ ¼o¼ˆ8ZÍZˆ–¹D(Á2×VýíØ¡é*d>+©Îou´FE#Nl¼ù` ‚I¿q³†¡ÄP¬SÏ5GÉ2yƒP§›k~“W~“W~“WdùM^ùM^ùM^ùM^ùµÊ+JA„Ë+RÝ©[GkÑü¥iþ,ìmÙ ¼÷ïì–âQIª€µàCÓ±€|¯¨éµ¦?Þß¿xüê¡zQÞðï\£qí?cNx|¡ã—þlÜÀZ»_Qf–×y9Tç=ê S»Å¬ø‘àÕE’A}wÖxt—@j€ÝÔC[Wgýëb\½/yO_}Ôr"ÅïéÜàiæ ‘ÝXKòÕ×é]–Y		ßRHŽßiF2¦i:[ÚZ­¯£wöŠYëäÚ°1ŒMŸ0-ù¥µ¾rvo¥)Ø­Œ¯9‘ÆGc*Bˆ™¼&{±çD¾ùìÊ;’k™þ“éÛi€Ê|v,¥afHéÃ!HbÊ0h@Ýö	(¶Ôªk"@9 îïÃä«AÉÌÒþà!…(o?»Ò?ÖÀAáùó¿€=mr qôÖÛ(Vk–ºpuÂ÷ïî©	Ú	
æô.F”‚»:‚>¼ n‡™kÆ^ÐWËhjŠ/ìŒ¤ƒ»f×5îûüOì©2	ï~M¿Ñ.,ÅG¯ä%%| îÅÛŽxm4Ö1?Ö."Í<dWð-½x”Ïé¥âÏÿTñwüŽQ¼þšß+Â;v›h6{Ÿß#ÒY3³ãýKíÏ2ò1Û‘×t?ï#æ9f»Ü¦Î5Œ2>ë<)'ï2zAÉ¾÷•ØgìÊN}G[ÊË·LÜßU¼†y=nÖ’¾ÑêÒ¯³3·ú±*']Àqq='ö4ÃÖWÕÃ‹~1ê*Ü]ÏÓ™ý ÄÙÞpwÃã±ØÒl{P†!Tb¹ ó
=Rüž÷^o½Qôû<Td=§?ª«1?õ—,~,&c8˜DVš|Ÿƒ«p·®Î)1YÍ»âr¦\¡¨Âýäüï
j¼ýìäGB¯h-hÇwb#ï^ÃgøQý"07dÞ8š5ºûŽŽèr£IÚ6ŸšúKèÊ>ÉÑtúZü~CeˆÀ@†PHWkóŸÂ‚+kkÌBK·ÍC}ÐpåOÜ V^OfÝ÷ù‘ÕÔÐo‘ä¸,[¤¯œåqO!R‰ûHã^B¢uþ=s’Öaª?Íd,Pµ0.XÉ¦Y>>Ú©yFôcÜª†)M¾´1ÊgvE£+­•I®ÑÖ\èßÐF?Ô&)?Ô§ˆO×¿Œü²~!t:¡²Áa¦ˆ c¿Rþu72ûT³‰à¶I/)ê™ bÃÞ‘Ð¸ÁKQy'Àp§Ÿð
xýƒúä~„ñ‹˜¼…>²"(øp•Àc¦¨i+µuŒí?À+¬ØÂøé2tÃ„›ˆÑã1ö˜¡_i6izWb°‚M¿ÈÏ—Ç!Ñ˜»Æœå{_h²€÷‰C¹¿õÒ&’,hˆsˆÁÕZ´ØPz³é¨$Ç‚õŽß‹të‚ùÊò›µpü&]beryf¥ÈQ‚c]vBöènÏ„Ñ¢ð c1!ÊLÓÐFàì pKµ@#†$†£KyÓÆia=3,%ÕOöYsµêæSeQ™2qv5›K´·º Îº›~°y¶€À{~ÑøJ‘iÆÙjpÜÔ]dz£§VnúXfeš23yŠ£­;Qù7¼h°´Â˜;BÛÒº4„òü|ˆ#=§ƒ™2<4¥„¦ÀD=¡Ãh4–	kJ·ÖI–ÎÎ°u5°oÖÜ‹¦,q„¶Ã*’Û#FWƒÕ °2,_å¹<Ö¸ÌŸÊR8Zê|¾žéÆÄÉ+¤á&÷Çú¼Ca¨›Ž®ý­õ‚ª‡R"wÙ»˜Ä†7ÒÊÞŸËŠ­a&Y,HWOúçšÆ!AÇ)fcê7áÍì]9Ú/¸˜5!ý±J"Ü…&Ih§m²KkzÞ¦'‚5LîS]+Ç–²'–nz ë_h‘þ¤-FÒæ³®öÎ>jÝ¢ýØ'þ¥,Jhoöa^‡´éŽÂ ]‹^Aeb¬/çþ…Xš¼ýìJô¤iÙ#@¦[ >ç‹ÌÂQ¼•®³òŒP ø¬÷–âæµfFÒòÊ‚w]§(~5e—ábçózG†yE³Ù	ô;Z€ØôLB»FXY»Þtc×VSOÃÇÉÜqrš|½(GBw²á‡W;·¯ÍH£N\J7†ïÁpÇ‰|¼cÄõÔbö~‰äU*ÉLª$Íö7‡;VS7«Œ,"©Ú½ð‹i¸³žB?u	7@-ÞÿüÏ“²²;À`ÿl™'*^ÏÈl4ÝØÖã¦Žk'1ÁÕÊTh„1‘Šìè¦‘ìS;q™§Iº¤â(	…éÕùJ¤„bÍqw66ScÕ„#¿Š²ÍbÉFr‡9Óž—óY6RŠ+÷É@gf—NT×'ùeµ˜I`Ø6Ú¾{mìÅCÅ¶Y0bêƒ/íõSX°'4ŽWdÁ äX£¥ré ²r‘” Aã3ÝÎc^/l>o¢6s3¡!—/6îÜHÜe$±œ"uÚ - ?-f‚ãl¹ð§Áê©a”¿]#7ˆ¼Šbf!GcsæbhÛ‘›M<˜¦þà»)1Îø1—l° þh x§nÌì]—÷ûÅt~Øé]Œfëüc¢¢¹ë7€~JÓ–ƒAa¥ØµÄ'²~pãè÷ð!$GÓý-·°Z,öÈÊûAu>IYW6Cl]]Ä6H™\¼½ÊO²—ç%œþj-ƒ*[ˆð8Nüy]TuÖÇ—NÙeS2$c—‡S4èu€Eu¦'„8Q2e‹c{=Ä$…1Ùƒ©8Vn£8°+-FÐ­­0¯Ú—ßÓQÉ‘F¶5EN‡u}a±.úZÿ¿ÿùßþ+ãW,Dƒµ…wPa5kC#aÞìÒCpÙèà>îUûïÿ÷?þíA'ÅÌ>ÜÂiM?ÖåÓ†ø1/"#ÃÿþÿGúGäÌæˆlDiÍ<Ëº4T`>[s	õ•5TNš '1wÍÑô„AS`L„‰mõì4¾’Ù"­=“yº£7ÒKØØz%r(W†Î½Á¸œ€™'È2ö<1žîCðŠ›Áy×|B rÈ¹qLwpÝ«sÊQ_NåY¥P^ývk­í­£’%ƒ°PèBH¹ÀgÙ3}’)òŠìØ’Uº(e+ÆÖLC’Ðò£•VA£¹h³±®ÞÃz©‹ì¦ÂÞüürq~2«F2:ÈÚNr^M7¶7w2¶¤t$—ôy¶‘A
‘~;x€Öí6õÌPP¬ÎÈ
ÓëkH%\-ªµVíõzv<5”ŠŸgœ¤PFð};9^àx!)ÈºeÓZHözb_>%ãœÛ_Ìöáâv»ú!‚>Bt²Q9)6&pñ‹ÿÁ#Žî,’x;öÕÏ=4Y0–*-˜$³Téz;1Ù—Ö’ÝõÙ#À¥0+uW‡Ä1EBèfQB±E«ÃOê­ÝV¸"Ø…Ž6þ#Éž*šd›Ï—Ç<¢ñþ6Ù»h%ÍâðHý\Ý0$<"¿¤$>iOx$ÿL®lZöÑŸ" ¯ÂL(ÞD2a}›Áº±_É®@âH¦m$cÇ§¼Gòô=BƒÁ“«°€GôŸäJ<°ßû÷ÓG_TÞ÷¯¯i¨k¤Áùµ"í•™v .¼Y?E4­çâZB÷Ú“¿p'"ôÆMùéà­²ÏZbÅN”E¬ZŸ$ÚVqyäÊÉPÃ0w–†¹3‰¹3æ¢ùÄ˜ûèÁ÷+AY:Weþ­_-’Ú¹‡šQW0÷W–þòQY^àÃ lÈºdM†B¹Âõ'¦š$¤Í°Î¡;æ9ôv6:Ó~ÞeÇR@ð;ÈÙtEºÒÕh”‚z8ûl6‹rîz¦"5n…®˜Ò b›4¶Õw”ÎŒ‹A¹pí” °ÌïÎ½‚Î4©ce…5Ê¥ó›aA¹’¡Ù]ÔGzÅH/’ýýf† «üµ›ùÓ²#€ñe”‘z^iÄ£'qAÛŒ°Ýv¤¾	Yß%ÌµVìûÚa1.ê|4 ÛEÃÍŒË“é^©)P2éZ,x†€.ÅåWªoœFã{RŸß$he‚œ€ÕÕ7´Þ[¤¶"¿Tœš­_ eE×9TQ6«¡ïûð Õ²],ÜßxZTÝÙ”¹Y7»y¼4'àÚÜKhÃ˜•b?üÚ_ŽP!Ð^ñb¦àçABPI.‡õ´'˜Š%®±QùÄ5‚þ2ËPÝÐnóîÍ€]™î5;ÕêÞÜÙ°îFàNp3€çyoáU>Ã gªñ›=XéÙnÞ;´'`y	ñÌh&ëj’=Ï'ÅH3¿¤ßø£™á¶B.˜)8Lœ6¯1ñb6Þ§×ÕyP6 äx7Ë'åþ4'.ËÁ·ïØL9Ú¾Ô"¦íg|¼šœ‚­‘¥äßt²Ž»ÖÜZ””4X­¬çªdÜ@å)e$Ì5d çùb^Ip×+óxu@ñE±âšFááÒ0¬×õ!\•·ÝÛË¤`âÓäYÐ{[–à•¹â…ß‘Ä¶ƒ³q^ƒå	½JAö)æ+³"ÈÚ±ã ¼­R9$Ôè	&Í˜Ï1ÍKÆ´5ñkÀHçi= G·ÙPè²¶1]ýŒ¡³ø‡ú¦ô“N|IR¨çÚµM9ÿó‚,m˜Í«³³QaÑÎ¸kœEÑl%J­È’X’±~ÄaW“ÊÙÒO+ˆÏ:TÙ’¯ðŒF¯úƒSÝ¡NAÂãaþ‡±?â—…9M,|ÄçGM0Ò!Kœä8/imÕÅ½ÑH/ˆøi4Þ+Èþ¹ìvIn$!V® £NYÙDAÇ~ÚæhÜ(Ë¢Ç`ÐU´Nf…b¥óÉnY³ÕŽŽÜ3vÔy{¹áÞ«ëü²Y‹»ÂÉþ5d. µh|<säoš×yv¼z2v¥ÔØkÖ­"w3î}¹$Bö(—¸œoO9J
áG7’5ûµ\F€´®pWÈ·’;²îg(ÉY{ëÞ5¡¦	¢,ê›K¦–NÃgüó6‰ùí‹${©ã¢¼ì Ñ”„%rÑTÐº¢žó˜êgÅ\ê<©Å?…[ÓJÊkƒÝ'ôæ‘5¢ž»¬èú¥;Ô"Vw1-“y°Å.½hîˆãÛé`šIhœdÚ‚x>p\?€ñªëºÍÔC˜e/5ž 2±çüHD{ÂD»?,úïNª^dl„±6EUì­ Í½ÁÎ{}ßÎ~"ƒõ`3Üj[ÆÏÅÐüà¡_ƒÃ+µø¡h^9@–‚&Æ/EÁQ°
ÊÞ-&*3²ìÃ?VPv$J‹†'ñ:2Â0Ý‰„‡*²N†¿zž‘÷ÂpEÑ=í°/fNPyHþg¹3{:A¸]”+7³£ÑZòÑYs‚R
EmÙØ³»ÙF]ž©«†cDS^ 82ÓÃSË¦Ç^ê	O4ÏÌ]ƒ8Â‚/HÕ¸®2™‰SNö`Ä 
Q[K£Ai6Ô/û5ö`76F÷u%hf«,yƒüÉŽ—-@AYuÆÒ×†G¥Ëûõò/cH¿Ù›Íóz:î9Qko¸ë×¹ºÁyÓß°%ð;¦t@DŠªžKÎÝ@×Š>C*¾án`ô¡ãñ6Ùe„V½Ûð³YÚ†Wg¤JL{¤JÂYZ/nÌ\²tjEB€Ø%@õ¡¶R¶Û™­ÜÖŒ¦4|ª„V¬ bÇïth}íÕjßÜäÊêfÄY¢9©vð¶£/Ü»›Ú¿åA mažîÙ¹ô<r^çSz•*¡^®yŸÖ
hÌSHfø7äúz}­øv†$— SsæÂ”ÖÀR,¡|¿5Ö)…Y]`Pžtà¡a	Ou]°ø,ÚAI:œ¥Fìµ±´±ˆ;±ÈPºž±|yï–[N;ki“À.YîØRlBeÇn.sî!Ví»ÚÕ›Wé[R<mUÑ;4í'´yCan&¢ö1XVèJižR ”Èr¢ôORÀ/ÀÐ­Td¯eJÇÖ;9‹ßCÀŽp½ûÔƒ3½­k(j eŸðw[#LQ\ŽEf¦µLý:®³î•®E|_RäÚõ Ók½Ei	T3h„ l"€¡6fžq[|É:Ì­(Y†ã@"nÔñUÃ<vÔQòòx–ü†^ ÚÇ‡@\³üûÿø_ÿñoÿ¸ÄZ¹09¸l€šõ¢ˆ Eºãg†•Ç+Œa'™iúMvTI8Ò$ã#KHR\°€„yí;Þ»‹,9w©;
Jc»Œi# e\Œc´Œ3Œ´ôu]§Tã¥Ž¸ñD.I"ªþÀe{Ä²Bâ´_ó +òj¨½_eÌ!0ÒFÈ²Ô,á&W¹î¤­ç4›ô*V]0Ó«Sz½HîêÉ×®ãË]0Ù+L~™æÕ…	ç7‹xÅ±D‘mCXµZ¼˜ý,S¸—ð­;³Š­œ_VÓ6{ï-ŒÑ§È…P¼Ø®­hD•DûKØI{ãÆwK^b³©ó‰Î&Dœƒâîêpùà;šmyúÈi£•ê5Jcó¼rs®NjÐ.¿ØŠ;²ZP{L† œ¸sL“&Žî“©ÕDìí#sN`,q9àæqéi^ŽV…JFíz1é“º†_4T*f}€IOŸ—!DXÙ´yž«eg~¥%7þp8Sìª(‘U.Ï&›Jøj-"C¤5(ÿr"þ•–§ú·¥ö·õR‹dà”©¥NÐŽ!)+:~¼FÐãÇ‘­N{”)üG„u}G#÷g´Ù4={jÄ5âÞôö&óvA²‰.¹ðiwŒÙ8¿ØnìÜÎàötT“:õáI: yð€‡™”ÞÀ3ÖH£š|VX˜šIÙæàc‹ó'g"ÎðÞÖæ'ò°¥˜å/Àßð6K5D/„–vL„or
Ñs‚@{€Añ£•HÒ’6¯¤VRKiç'7hÝJäDÒŽþ„¨Oš–i'#ô!rD¥÷[¡&%‰)&ERÆ(}ëO†uqzh4^àê°óÃÉ(Ÿ¼ãa]Œ;“ªš“b“®Šºö›1²’¤Oq2ŠˆpZv$-ëîÍÄ#nB‹4‚¹{ÃÙþí›QuRÄ<Ç´ïéî‚eJáÑy{aL¦à\)~Ñ‹zB/_6#»döÀöÈõð‚Ìa’hvÕ†8F—ê¦,lâ10K3.³ËªÉ7€$dÃ2ßqvMì­ê»ŒgDò/ŽW›çË‡ž>8ß÷ÉOyíÏšd|NáÁ:»G”„q#lææ]ÏÇóÚZ;×ý5Û;ú›«=ÞôÒ²\
´,—ú}Aå×ytžŸŒ
¡˜Qpq*…HYK;vž;˜‹|€l^“BsMI#ž×ÛÛ-ÏŠ¥nîúD2JÓó6ó™J–„RSxZ­0¿j_Ò:ãS?Ý9ØœS§Õ9bè²M-Zn³ôT÷¤™ûdEò¦mÈœÕÍª=©úŠ<ç»8½*Þø=:uÜè%øW`x#äy ó¦›æ'ÕàRï”¬A¿ËŒÿaÛ¤áúïFþPVåã¨Újãç%àëÈ T{·AÔå‘Ž.ÙWÇ	Pq\o„¨£"²DšÔD-Y©ÉÁ|€’ Ì“TÆªís1äÎ…¾Œ21]PØí•TçCV¹ ²ÒÄ‘•fîˆ¬´rJU}®‰IÕ³õ`ŒYK²ñsäÅ<H}YiçùÈJ+ÿGVÒ¡ö…d%†t™âÉJÄ)#¦o
i‡šª‚½n`Â,‘æDò97K3»¾à7Éænèœõ¶²2S÷Šq5c¢¹YÚŒ"JÆf.°²ã$¥r#q(rBÃ¸SÄ"'„ršÆm.¶ßõ:ÆNiœ"Ë”<Í†•Ön°¢7sZIjs9¯XsdÉ¾‰­N²{¦¬6ã$YQ¢†I©;'ÅwV”¥·PÔŒwŽYù[GRknÀÙÐ4mó6Üºí<ÜØOÀŒUøºÑa®Úß-0¸†>obx+ò{A­ï›Õ*üß ¬ÊJ+?81¡•øÂA‰­$Ñ¥½_”d±&Ý?Ž6n´ðÜ2rwªx“æë–:¿˜“ì÷F;¼!áfÕ>pÉä³‰7”0bæ—r ¾s	ßF¼ë¢-4´´þÿ   ÿÿì}ÛnÜH–à{ET¢¶‘ª’˜ºØ.W–d,ÉÕš‘-$W÷ÀcØT&%ÑÎLf‘LÙj>c±Àb‹ô ýÔØ—yÕŸìì'ì9qc¦”’ì2Qe%ƒqçœ8q.ÎçÆ¬ðh«žúÖ_«âgxã/_Âƒ"l¡z«Mƒzù,Lª–E´h! >mì ó•“ÜÕëæ<71Ù·`ÎS(-{ÑŠÙÍ^IEûšÓçP,iUya¿B³ûÆ[íÊÜÜf0s(»¬¬´à'ÀX-¼-+6®aó Ù‹45uð±¬©\SÛwhËiY¨ïÓŽçv) ¿‰.Aøv„ÓGýv¼13se¥YöÕµ.-¼ûVÕ¯Ÿ'²Æò·¢£ÝtÅ«·¿M!@ñßoómQ³J®§š8Ÿº®–{´iâëÝêaƒ·2»[­ë;Sðm©:9î"µ¨®¬ä¡+æá0ìÖ­þ¾m:—pÓ8…kvé5.¨ÞÍ»†óXrÍ=Üg³ÞÍÝÄÙ&ý†]ÅyÌ¸_°¥€O¾éQ½YG³¯yÃ<C¨¸SÑ|ëPÕ7=Ù§¢ø&>(ñw;ß‘á8$h$b³’_âè£â(º¤wÊ1z±’•M¥a-bMSRG‡×±ð™þQÓÃ^A:ý£¦‹+ùíþš®ÍÊz5o2b—ÿ˜SêØ3ëdií¢œ¢æVÎ‰kê¡±Èc™g)ÁÉH›**…ÙÙÆIš×Í¸”÷gåy·8eÑ½ÒSßàØbF~žÄýˆl„i?«‰Nâï®Ì§ZcÜ9|&ã=™<$U«!“—7M%ßÁ?èúã/«šï­åÅk†N>}àÐM‘ðœ
w;Ñð(TŒÎö#8 ªùV;§­C*–f/Âí®#Ä¡#¦ñÕß3Øj…ôcS‡½bX‡šÁk Ýûu™”èzL“;50G)ÉIÈqÔ;Ù=B?$Lƒ²êáFc6Í:|÷Ë3 P‹è JwóFœöÑ­ j¾Î	©­!œ®çÑÈhs
HUf§!¨þàU{	Eèìê/i2H!é$È¥%0°HTg ÓDË»pÒª%hÊ+Ù{
˜´u`É‘&ÂåÖ0JO®þcKy}È”“Ó.—– 3™ag3Dv/½ú[÷BqNƒw	£1,2*yMè¬$”¸Š8ËIrL˜«¾8Ê Á+ÎúÜÆÍ˜ˆÔ° ZÆn¶Qcšè*ÚÕM¤»¢7ê>àydv<´ÊÓWOcºxëÜí¡ „‘Â8VÖ­êgÞ'H\)¶¡K‡IDZk›ÁrÖ5{/ m€[³ô€†–±Q"~­VÛÅváÉ˜Tø7wÕÉÎ€£/¦ë¹ð™ Ÿ´êþ*/q×ë(K@p|7,ãËht:†â0Fg²ž–Õˆ8]„«ÁE+ÈÁ$V/÷¹Îp‹ÜŒéqo‘{e¾E,\pT“ÆU6c!SYÅ@IúŸHHûùæ/sM+â~,ÚYã’x¦‡Ó$ã©H6£^’Â±2±š9ÙŠËcQ‡0{±f4¸s;,Æ
¾˜ —µhúØ¬¾,â «Ñ¹1³-Ž¹lôÆP±‡A>×3ä*jè‡çY©°ÔùÉîìŒfâ›M¼ª^#m˜žèô¼ŽÎ½‚‰}ÃKÌ	Ã²ÖBË%½DØß<Ú=zß%Øª+36Àë2Ø$£~Vg#Äë†±`Ll \ùŽ œ,^?°˜½©)C|	{¨—áË6¯<ÀE¾¶=7ç;Ã<—¨ O¨5f„ý;ÈÑÊ¥ÝçÏö[Ž^^:À¦Æ€—!>^†„ø(;ÞÏZð®Œv}ÕÐ–•rå{Hûë«¤Tojã¥ÏÑÐÚ¦:U#Õ\ÙÇ;>/÷þ±ëovã9ºÚËƒ×-y%c Rµ 8Y•ü¶«.ïÑÙÒXñ481
tÑ7á”T}9}¯­:…—\®8‹ Uœ´‘„×V7µçu:÷šqNÝ š£™F>SèÑ–ü4×<TÔ…<ò’âÖÎ+·Mí¨PyTã¡ßæÖÇôšÓòøÒOÅÍmÔjCÑªjtuØaÓ[_ŸƒhHäÐÉ™8¬xŒÊËÛY&Ï¬6[3WŠì©w¨8õ¦59	ôÕörŽJ8ôôH;èéîÐoÏÛ\ªŽ+^ÍVœŠWA›×ÀO³E?@üùÓwß×ûúcµÝvh‚î/(kNzÇó(N”Tq¦—Fð³¿žß’ò³Œÿ¡!S{'î›0!ð&2xØê¼Sigu61~zÅ³ LÅüQhcs1§Q-C“KM-Æ ˜¿ÍXS«1
ö¶kf=v}û±&l¨'»i°"£×Ãh¬	ÛyWø¹‰¶4­Â;–ªO+pícMã¥‰OL|t½L¯…gº›¯Æ}*òB.×àœ×ÌÚ}  ¡Ãf}Ô˜ÃKy–H\ªønVLå5]QNR¦»{Gñé:yñ÷ãÜbR}*àWféF°Îÿû_ÿûÊJÝ²A|üt€iÎÏî¤c€ÙƒÙ¡€®~\…6©OsË &'æF íÿþÿ*ªüMÂ™îÜ`ö°fŒá«›t÷`óm}Ë Ggè1C7tÛ'£$ý,áÍËÜô¶-¼lcû|jm\ð™ÂÎ…%Y¼¶74Å`‡1œ! ²d8D¹‡o+/~·¡ù-ã	f7 ¦·Ûvàc]g“‡ÕÊÃbçaVËÑ”íôïÚ«n—ð2Aí®]´C˜Mœö‰š™š$ŒÓä8DOƒ4D,ÓzÈ‹0Cç%³Ú¦Ö$mQé?ê*¬QuTGIœÔq|²v!ê¦‡	4ˆô‡þm½÷ ´ÃTmÁzc@ÕIa)³¦½(NÃÒR×mzesLEV×–ßƒ" åÅ^ÄE­eBå|ã^<‚¾,,’?#GTëîÏUÔd2ì}Àæ#øÛO“1^¬¤Uµ’Õa‚ˆ"€–K°`ÖÂÁÚÅ…ˆÎÔ%‹ó$Ã«xøüø°Šà8mÓÊ,É2KÕÑ§8oÔBæ)×0…_\	ÐQ«šÅÛyýãâÙé¢;§Ô©W±¾ž2ªKóZWDe¼DÕa¾3¤õE¤;zøãx§´Ð=î]‘@Ä0©„´S]Ä½Ž˜—ˆwi—‡ƒ,2bV»|ÑD«-•²þ$´ÛÍTÉF…Ì!r¬º“<ª6“tÅÍ†îÅýµ¿9Ì³y09âbõAxf¼‘,A¯00{è9OŒÊ–Fej]Võ>*äÒ¶85âQ"¬’ï7°ô½ÖÄJ^íÐæ­s;¨Ñþ¢ž‚kêGÇádÿ& tú6}Z£6“F¿Nâ4ê[3¢q|BÛ¹¢ïaÑÌ_Ö]/ý	'“œ:_u[,œ+¢À€ªóÌPEhZp«¦G§0à(]k-.‹ìÿü¿¸àrÎú9Â†=Y•?ÐÂYÕÒ‘˜Æ.Vmÿ¶a¶FSÙµÌÎºl5Uáiá·êªÒáòšP)Ù¿OûÔ¯&cäÉr«õ„Z¢CJÁj‡}õ.®È¥Ÿ¿W£iq?7©7Ý¼GR|ùDþl\‰.´{¢I¨ê*[í0 ý<ˆ‡ôm3;e¾©9ªžš¢hþÁ¿ÐÉ-¤C-uàþÿm\”9YBÿ4.ü,MFŽZOØß/j±žfÈbÞ§&Tþ…5vkëS—lõ'|ÛÏ“_Â4zïhþËf¹>· Il$Ž‰<Rƒý¸˜ŠëC©â¬uZ0õò÷úõhxcp²…>KgÔ'ª˜ð¬&“bj¼®~_P8äžƒÉŒ°Ãi‹@3„á°xÚõåí  Ä8‹j/P£€Å(çBÖˆæ`Yûè2ÆûŠ¼Ôo³†Ø~¢¼[€Öýë“:CUŸŽ zS¼Þý•mQŸÆWÿ>Ëˆð›<5•~—¿(˜º›õ~FýFÏp±™cêi—š»µþºÐ×½QMYÞÜÐ¦–´Õ8$ÿúË•ìÄ£d” ßÜi¯U|@
Vä61-P‰òžRªÓ<gÝNç$I‚“Ag%;A|Ybª;à¬ÂADýürõ7´­%åÁÕ§í]|há'Ã2_à¸±™¥²™€7ËN	5Tá&õÉ0J©_Ýy¾ù‹el¨*ãäÚxîu÷PìºBúé–]Ýv8½©ÍßÝJåƒW#º9XXZDã±Y_MÉUö¢ÝàÆÊq\e—YŠ«…Œß7SÏpè…áŒc»mDýÔÅ ÎòçIÊ–5@™€Z÷ b(Ìà?˜ÄýK'ìÍ­Ž™·]ÖØÎ^ÔÙÀN˜Žrm(ÐIEéYÜ‡y¬ÍÛ’;Ð­ù>äÉ+t#¸fQ»Æ0ÒeÀR¿LvÇ î[Á*2²d0€>áØž'ã…¥Î2Y Öt£œÓïxä• Rºý.jîøTuÛkvèdõÝ‹TØ t_ñ5î Ž-<f&Èf°'¢aèò/tV`î(ËwEc®Bkv,dp×«ŒüÖÒ²uŠ^6ÂaHsús1Å¿N"Q—â,ÕéÅÇÙ­â³†×¥íÉÝ¥rË Å)*züÞ@Lž‘Í¥ŒW;‹âÌæÊeíò³aç€Æ=L}å)k×)­mÃ®UôÜ\;³#.‡Pã#Ö5öÇ¬ÙŽ]Ñ‰ ÉÐMää}åæ~ñawa½Ó¨÷á(ùäFÈ4WÔ_«3ŒŒ	B°ë¬Ï÷‰¨ñAÏ‰’)â#ªóÉ%ù+Ñû¶>˜ú`L>Î"mm½~S_¿{½=(
§?zÄ0÷xÃSøßqÃÇi•æá³A ‡	z‘j›Ñá¥›!p7Ts¨¬¢`îô‰¾ÕB!s§gøÈZuÃ uL8„Ñ:=üxÖÄn¤#ª î]Ê#äñÅzú2å-#<îyªboò£Kõ7àxaC|š`D:+Ê­·GU°–·q-›L@ŒlË£ ¦ aÜ…±Wq|ø‘eB¾aHkñó_ãæÍ‹§K^APîè¼hì)°	qâŒyØ{XßŠÄÇ7¬Ø¬~žšj°šÃ×f‰ËêµrÑë]sì’^Og¦Ø
æ:Gµ¤¸¸¬aëê¬í2jfZ9ìx­YUÀ"BnJ§iLvpEeHÎOÔùƒ©p³ÌÁ‰øE3.ûZêlÂÁÆôDl2×îÔ°‘û±WÌMFNÃ>=©0U¿”²JxZí-ev äËßýîx2¢B²P’³'ª®&hBÝf|C¬¯³Î›q2O‘ûý0í$£®3"t”Só¤Ýbþ°y©ádÇã+ö‚¿ØËÉ‚ÐmÈŒbšõE4LÚ*-`¹ÎxG œ|³èù 	ó6ïiFTªÜnÍ·æI+hÍÍ¡æùâOj¢Kz¢×ö*hAh†ƒjòã)!Ï0¨ÃwäÅÕòšáÄ·ÏÄ¨œ‘½þ®h;u9¯Ì²èÂuV˜ÊSXÓsô3†>æGt˜¬ÌoÝ©oòAð’fxNK¯åó’ªfù9Z:·z¼^éxM$À·gû;<ùr.`ÝÀ¶éR3'ç›Uq]˜?³Òøó±Ä4FÇÙÃt®‰VO—+ÙW|Í”7®þBAžÅÁc`¯›R¬vN—•Æª" ÿF«6LÙiHæðþ„Gaü)aNSzt§QÁMª6KrŒ’£4.%5:†÷Š{TXŸ‹HGâôÜ±¦!°ÇBª’“ÇrEtÇîÓÜÇ–]ÃÛU±Ñ´B(qÓ¶oƒ÷Ù’.}à+­‰´DÈþ·*cf}¬ì>£æØ°‰L‹Ë†§Lü‹ñê…e÷¹nÑÍKçÍ§ómÌ²¼YilŠÙðFFç+Ôe©ôÖ'`O ¿[Ù&²¥vÓ¼Ì ä$F‚ÊUÁE‚XÖƒQÅh½/VrÙoïütŒ€RIpà` ÷íÜ}Ô]Éƒ«$I€Â“iµƒÝ®@QøN³+±ˆ@¶ÅÏ2†ÄSé,‹F·•åÀ÷Ë×+f\idL /Óg˜‡’ë@w7†.\èŒ[[r·Ð5³¬2bj:×Æ°Ð¦«•çW7ò¨Wÿ­v·ý€ÏøZË·8£9/•¶Z×{Å#Õ¢†–gM4©„u0Gõ
v\
q¢M= 4/˜Ì«¿:Bf‰áñ.hü! ±$¥á.)˜ô“,ú*+fÀ2ž¤ðç,>Á­Fõ²IqÇÄ8LtdÕ0pï+õÇ£P½gª2™ÚÏÒ‘ºà¢ÇB<7p]x$€ãYšáTó¡?¸Ã,š†Ùñ˜þ –ÿ †Á®ÄŸ2.
Í—7ãÁdD=^lö8±àï§až…ãñ‹(ËÂV)ó)öGþå0Ž¿ÑO@7Öí_Ehyá{ˆ÷»ã/Ìí¬î™ž,ý ’ÖÇ1¾£<äÀŸ/Ä/vå‡o¿žìÄ'4$=-.Óu¡9œî^¿¡c…ƒó,Îz§Ñ0âŒ.vÖµ,}3D¹áò`§ÉÇíÑû¨—SOGó¿»ìe)»/Ò÷Øù78T?‰åÕ>¾¦²FºÚ]ªö*,‘¯|—´‡]ÂNöó$
ÈlÒÃ=-¨¹¥i’¶(U?KâþON 6øÃj¢ðÒ¥˜jk”§ç,€§K~ÇJ¤.Ùc?Ãìû"áªK6øO^7…1¨œþÝÇ9Š<Ø7u]<ÚÓ_úw‰Ðû¥.g—P°[—),›^»@¬z€WÅÔqáÈON8ÆRPa€·ÉaÞ;]¥û˜¥DV£uº‰WÍÕãn§mTvC·ìžNL ß 8‡üçO~*@@ÛFŽ“ö”÷†þ.d˜ST Ä®ã0Æ×¤=-û5µ$¿J•T6ÿ|”$ z#–~)æAßäæ—W¯Ž®<Bu”ö‘ÖÖ1bû¨í#£Ç‡‘	zº0	§]mLLüqÝiÒÇ-ÿóÖ!Ýî{»‡­Ÿº¤|J"’K,F÷? ”aœE«J[ãè³K^°t&e‰/Š4†SY:ýÍ>¶Kþ™ýû³ÇKÐÜLý¦ƒc¯¦í
WÅ½03ÿý®êžhƒE¤÷ä'+¾†Úâ¾˜`sÁ*:¯ ©>ù¢?J.5én	!…Gó„i’[ºô­Ó8Ë“4î%t­)Ú¦¿ UÓ¿Çq”í!+Þœœ	‡"OÚJ%ªðXz
U:cìÔ;Fð1×QÆ†(r—¯¯ËN7Šw†Îå;¥oòMàõ¢zàô'ÅkžF1jJgÆaY¾`©ôjp¯ƒä¤(Žœ½ÉWGÕdÜc.#ÖG¡"ÈÆ‡L¾–™mí•í6ë?«?G¸B«…7râðZCÎ²¶LÈîG€ù³<cK¥¦hËuâ^L	æsûë7O„ê¯mP®z½RéŽáƒZ7wÛ¨TÅPësz•Ëî-”("ag(Mã€Å¦/ë0QgÀ–Ó@í“,ìÀeE˜Œ¾Ó!“£¬—ÆGp,M;ÙÃ ©y1v#ÚK>šñV[ÇÇ°…µ»TÁ)œäâ…²€º”ß?X&EõËfâW+¾f:¶í>l¨Ý­ÃíÝ—ÁÁîÎöÆöáúÆúîÛç»;?¯‹h·¬ôŽ"bXkˆÿFá8;Mrq;þ« líŒ)]Ú³jPÕ•9xž Ÿô¸füb
¿*]DÐ»’X.@_¾í¹âÃåÜžªÐ¥Ä_ÅåIàÔœI–fUIæÃ~äl·ÃyrTQ@`£“\´Ã	‹¸LOedÜâš©\ä9ò)Â¯SX™ÖœB«ÕxÄå×Æ(WÕ60)ö­R®"eÏKX" ßÚ­-üCøV%0PÏº­yÂ*˜²bdâ>‰éU ØžwerhwdÌéfÔ‹ûo	—);õ8ÿÂwc·.óœe™eßKÖ…MœCä4Ñ8Îá$?ØÍYŽ‡¦bÏª° dx£ŽS¹*±ëEútH)C[)ª¨ŒúBUžê¼S®W<g¬¥P“ËôT¤PÝx™Ü­öº÷ n¼è‚¶$0ŠË[:ÁûÑ1¶’ôêÎ¼((}?†1`žaj)ªW “-W
B¾j{hßÑ-Í‡1:fíªsSäbÃFv,JÏàt	ûZŽì£ìE~ü-Ô†Þ	ãÒ˜ˆ|{!FšñO@fhO¼öÛ6îö\B bÞcæ89eà¶.ƒwEwå)[$ñ¹¼.xBºµyHòeÅ*›Îîl±UØ¦–+#ÆI³“0xCV4ƒ±xaÎúÅéêRÙ¨¯cä±´´×?žgn˜µ47}³WcŽƒ^Dý8ä5è‰Æ*JØZb…
L	†š]Cl­þáðÅÎ6Þul"äÁžT4¨ðd­pö€ïÙÓàõ"×Ï¢(u´€É˜à	#ùÀl|„H°;ãþq«X29û »ä8<K(s£„é¯“ø,!{›Ï«Ë@´†/%'Qšy•‰Pq&tè}¶á|Ñ]Ñ~»E?,ô í‰ÌÁÏƒä(ü1I?Dé.;2éÛAÚƒªÞ	Ã¿^ô>zƒdÒ?ž6
zÉ°¾?uñQ†Ó¼Ï:ß^ÈšagfPÝ%ýÄª 1Ã÷Ù»r
Ó4<6¶+•½§S¯|h—ð&Ô*óÊ6ay+MÚ4Æ@W«ür.³s›¨ŽzT†uµ
œ	6iã§Ò—~‚?«kØJ0š÷P†Iß?Wá§ÆðMívó·c…ôrj“ ’åEnÈ†Ù¡3ì[»R3²Þòò¾3žrT[ ß8°Îï“xÔn‘–Rò÷°¼ÿ:ZXX {W9³ ºø’@Â¿ŽÞ‘ï‹Ö¾'­EŠÂ‘«¾x›Ó>ÓžIIT§a<¢âö:™Ô—q}”%‹ÁêU~/?´K_žbÇ kˆ~a7ß!®¦¹®þ³"w²Î6`ÆHA‹ÙšA6:Vmb$°•É(PÄD=Ÿí¾ÝØ}ù|ûgØÛ@‰G­9³h©4%]eTžÚ‡È)ü!¡BÀn]‚8¢øP¥ƒ-À@’6À¼DÀ*ÆÂˆÈFu&¡A×va÷oZM‰–‰2½«R&è.ä†L{ù®B£è%É`Pà;fÔÙÖŠ_uz‡§û½4:ƒ³?ÐÊwƒFÁ?a° áüi ÕÄ#,‹ ŽY(Gj&‹˜­L1ÍFíRSæÔYRFSdVòŠ	Öl<Jåp4*· L¾©]¿Él„|­
JÎ.7žSš«éIqzŸ˜ý1ÎOÛ­x ÑAU?ÁFÐcà"£œ;9Á¥r´ø°	ÝÕVÆMwaŸ 8>Û2äß>æ#ëJ!—éëÇ7E!çkã%t28‹õF)¢åœ'UvèÀ²ØÔ>MPÉá™ü³žm%|µ¿Ó>6eIFœÍjGÈQ´û¹ùÇð„·I‡ª7Gð³R•ÑF€OX8ÂˆDôBI5?{s®ÝbZ¥úEñëzûÇíÍÃ?@í™þ°µýóÍ¹¶ŒûháE;N›2F”çb/?•U»pYeOx~“u–šKŽÁlÆÅýnMl‡ûˆˆÌ\ÍTÖi°Úua/yóO”Y4÷“5Ì»É§»CÄ,YÇ¥æ÷èjy)tbÐÆYáyd{æ.qpÌ?!;ÅJ ÑàÔ¸ÝZîWA§
˜¦’ƒ~~d›ÀeCÁ´‡ó2ªær Ü¢é èJžˆ]ËÑßxtl1xüÐP	Çm^M%‡uÁEAº÷9Ý¤îe'ƒåwüÊ£RŸö~iBÕR´’2ÑCjoPäÏèqÃÝµÔ¥“ùµyY¸J'¼
yô@ rÆÿiž’.Y¼Lc‹Òs³d"%Õíª¾%þÛÍã1prù$ÄF¾qv¼"ÒôÅš¬Ñi¹Ež«Üb…%‡
”E{É~k¬ÙQáT6ã$°ƒ¤ÒŸ¶x]Ü8EÖ·U¼kì¥ÐØ¦äºŠNq]
Ñ/þªwŽ~±›”›Ñî1v EOé‡ç»ÇŒ¢„ÞÙGýpá˜FûÂï—åöÙeœlŸ½Ú—ÓÃ¸ºž.¶G()î…Iñá8RYöaþ<iyËI2ãz/ì—sªiåq01öfEç“Ò1¼©{Ž …ê1tœF¤ò^±Ÿ¬ÙE±}äÅ¹op”¬£¾fœ¤þ… ÷y÷©<BŸFÌ†x¯]U$3Ûtè	–lbFô~²}°{@q5¼eãAã°5÷zñMQîy<lVˆÏã¿nâ ä )ê‘)t}¢Þ)»ðåÉ—ê©DìF”;«»ßµÝXUºªnÈŠ ñðfÙÇ$íËJ7ÞpÙF£pËñ:¨/¹ñÔDm¡t3/<rÂt+]‘§N=Ù-Ðó\T
MkØSS\§äòˆøü+pKMÑªÒõÀôƒ·é€ÍëÖ zKO«¬BY›¬®ã\ÿE­›'i]çi®…Ukz)_»Ž‰0›l–0äö…1¢¾9Îó¾!Un"Ã·0<ˆ†ñØX•å‹¬Îò=„1ž «µ®r’¬¤ü!£`œD-]N’¥Ë`V’ìç4ìOBJn”y1¥3SýjÄ)R¥I™¨‡L­jXw½â—J‚V©Š£U^¹ã:†£P×ûƒ£õFì]ä¾|bFõE¯9ŒÉÂ…UñÎz¿¯¢Ùm%Áçær‹-T½úAƒÉ­’˜"
 Ú0m“¹<’	ÃIxÎ;-·o”*2(—’Ú~bV›m%·ú]a­K;®ZNË 4nÊjqC¶9µÓå=hê¸žGï|}ùJ&Ó ê«±dÕ#qiü£Þ}G‰â«©ÃŽ‚å<Z%Â1u‘¤
%z?ÅŠ¯årU,çè¬¥Ž2B4¶’ÃÐ÷º
JY¬£¨«Ç˜Q¯­8ÆS_ImVe5eõåÄl:¿Ø:\»¹½>Oô*Q#B=ms´R$)²ynÆ²d3FkŸ°8;‡Æ³3>&CÅ:lla¿sh–‰AÍÍk2… *Ã …*–[ŽP;^ÅÕšy¼òRQãˆÚ4íÍ	q0¯2"~Í_\\«ÒÂßðj_ë8$V}™óáõš’­Ü«©ÌñšLÉVÎ‰Ï¢‰Kóhûæâ¢T°qHŠ*ò Œ9÷P«^&ô;]Ä#ÙŒ‚»Fƒ±U*OG#;ó¥pZZ.™^d•˜–‘§:±¢Ì}¼øÏ?¿ÝÙþy}cwë@¢Æ¢Ó Çþ™`qæ­¥N²ˆË¾!ü¨nÖ(RµâqÔG/ÃX«eª×o Tä½ÆvÔ2å~df†-Y2¼P.‹?Æ£~ò1àæ•íÖ&`Ô÷@\Çá	t*B{Ï4:‰¡0UÑzùo“Õë_u3XéÓ!¸Ôn€ùª&Þ¾h›v¨¹˜¼2Û|`71ÛTá‘¬˜èI1ÓóœÞ
ÓE`õÄ¢i[Æ!¦9{u°µ³5A¡l-Ô ×]¾:ž¥	•M˜1u@”í£Kü4Å†ú5P¦¸	8['¸Ê¼µ@Ëù}ê&qŠ°¥QÆÒÎˆú„þ)Z:Ã.`ú­#ügœ Yr„ª ÙqœBP¨£Ò¿Å€Õ%õÛÊ¢šœ\Úù«¿éû PÐ›}ŸÔçqJ­a"*{TÉ”Q‘ò.ÌÕÓ<DU’Í­­Ã-õû»o/*C¸ì|{nÊß¹I•¼™âk!¦Ý¤^réæ?häˆ›c=*5»¼iw¶_þSAÏ±JÊ	¿3Í€q³©ÛsÝq¦}ôé“Ôiû¦¼‚|Ò^y]æÍAÀú3!÷l €ÕÝží¾Ð¼¬– „¯)Æ.®Ò”tÏ{gÓºAedS=ÛÀ‰µzC Åì9¼é¡×âå´ÔC4Ÿ˜b å;L|¦¸ÇÅÝeÒE˜æ>Ÿ&wšêÌyìÓV‰uåó[o!w­¸6šÍ¾µ7Ý¹Ïw÷·6ÖåÖÙ6ïsÅ—‡x®±}å¬Ôî_‘ó†v°¸÷ ¯±å {iìÚ‰rˆÎ“Ù”wè¢lã{t|wéøTïÓEæÊº:9Sì7	–þŽ«ºÜà~ó…­½õ—/·öË‡EG&û@•áÈ"ÆfDïš¦K|ªZ>\\š>×\QÞ¿&Êt‡îbA·ö·w7wßn¬ïQSGÓÊr=§k¯l”^ýGéZZÖ–MX¨ð1+QÅÊZS¼Œ)Ù¢PU”(§_døÜøÂÌ3êÙ¡¤W^èÍ*bVéÜJlùÂ\á3Ö7¯t’…tE›PÝ`Ñ5–ð¸D]T—Xø˜Übácôp%UZ¹&ê 9âêöÏàgûõ?ì¾˜ð >>§Š¾óLk€,Ï½)à'ºº­âû,IWÁj#ªpüj‡«àï¡¶,êcóZÖÐ¡°ŠÕƒÓ”ZCÅ"¥Ÿ|Vï˜‡Ž·áè$ßÒ¥zûí…	»pïD­=Ü	
f™ÁÉId‘¢z¸E–Niô›e«+dêùNúçÕ<(c	{°¨ö†ÔyD`ß¦:
Stï«ú¨½x­°+ÔíBáæÎ§èð’„dÓÐ_HÖ%¯™õ7yÓ¥¸»v®þícH+|ãÎDŠpf4-¥B¤"Ê`[þœk½QùrÙÍJÑ£ç,™õŠ´5M†I6wýcÎ9”žÑ‘çWÞm¦S6_qÚbî$w‡StRœõ¨BßÝÅÌqß2êÜÉ¤;Ÿ3áÀ§è¿
oª¸/ŸRWhÊô„»*õF¦ÞrÇJE´‹‹Ûï“p/UtHpß}X6zIvâ>®”N	Æ±·˜@ú×*º†Uã}‚p xûRzý:,RÑûáíw‹y&+:„ÂåŒ\ý÷<Šï©¡^Ì”î	€²4&º:#£i¤RgÏÕøÎ[Å·[ÑÕâßî¼¿ÜžÂR0Ž8"xìOÃÑÕ¿ÛqœVåo+Ç×79õHG}bñŸEÄ7Ýiuò07Ç\RäáQ%n°)hã™ƒ¸Š,m §/¢Ž©>Û,?šâ¡S/WpêïŒñÇ«þË©ÛtÆÁ’i8.»„ÿöBkB8îDpèn‚u§”]8µ,Q“¸öa¿>ôQ—aE”€X¡Œ°Tç_¯õò:aºSqº‚Ô<ÕØHsJp#Õ³ø…>1ÅÙ…üþ÷
(¬fLJã<V«y gä­£jMô\÷~’åpÈ^8ŠòQ4Ò"–½óŸ®Tý’{F¼Á‡îûÚ8Êãœ^çt¥Ô¢1èWu£¨z÷Ì…AÅ”ÚT@ìeÚ:×Žî%çÙ1+Å]¨xà§:)ƒIæ
2ÍbÎ½LÎ9½¦€¦°_UÏú%8*þËp{HwwÏãYÐÁaŒ™jÄ…ü4
û†°0yZZ*¹Õõ¨G†È»‚2ÐhåìXü­'èmµ“Ÿ6+µ…>éšÛØ{Þ¼ÐasšŒ¦è%§¨Sô3LO’æÅ˜ÇÓæåÖÙ¹É\RÓjP#@­æèöYm `báœð¶H$ø\0O‡,|l…¶m¤†T®=#šC3*d»7²…úlUIž‚vzÉ I³Ö¼¥ð$ x¢>Ò"ÜKi’a‡‹–2ÆH–ìyß²x–Þ¥ŠF¬¸â›]®åãÂcrŠfNJñZ*qc(Fª6*è-³O™3„7¬/žÅ‚Þi˜®çíEkNgNú¹‚Æ›ªúãê…+©£°Áèš~± ‚d;áêJ}§Ž=Í¿•ö{ãcÊÓ/ÂeÞZ'Æ§èGñf»aß”,¢¯^x ,@c<È=òØá«7Ë«eåÜ6Å•ºfCõáýè–ŸÙP >Ž`ãöâï;’sÒ4– Š‘Ç!jÕØ[‹¥xg„qE¶àÂÜÀðÅh‹N'ÃpµÃ>4Ž„n¥eåyÈñK	£KwcÌáª·¶ÇsŽ@ÆVo
?]O à·6_Þ(!Àö“Ág±'šo†vqË†›µ÷ww¶¸&m×©ë´ªÓ/ ¨î{¬¬ƒÅ[ZÆ––t¾Š®êë¥Åñ§7¦™•EÅG²©Î5xZp°’Ç£oìPÄ&PŠ†QúEa‘€åíÝ³®œéJ¾:ýl ¦^´Ñ©{Ÿ%VúéâÎn¼9qS@?Y‹5 »xJr76³ø“ƒ×OuåÊžb„wg!Íµ‹®Ÿiz.]@“¸	™Î‡s*§(p4An›ÁII€äâ<ÎP9‹óž’ ÿ¡=¿™­þSt¾]¢¤GQ’’Ý$^ÉÑJ¶IÇ™É2³2qép¬Fmœfz±[Lcë8väwN/Ö±|Wsë„y+“A8¶ê’o
Ìå`\,BR‚ÄD¶VÌE3|©¥º×öd‘è&!”ˆ”bxÊJD’³ÃN(gð§ÓÍ(;b¤#¥”ã™x±WT#(*w/Ëê«ßìÄhO×­«p§Iu.áÅ#£²#ÛQ¬)@ÁS®‡¸ÙçO‰’Ó0;½!œ4…„È$®5±Ä…µåÛ
z·à9}¡x»Ó¯Ýhþ2·‚M€ÁqDQ.JðV¹Ÿ&c”¦(¸ù3®b­Øpl`·V‡	.v íW†b46^»¸ 	ðZq~N~²^ˆgºÅàÇ‡&Ž#a ßH+µ$K-™Š˜îùíÌ0ü´@¯d÷ô¸|ÑXªÛ ˜Ÿéí#oãšwìÑ(½fË7‘<Ý_Üle0ë.Znþ Q§˜&„`Ùœ2GšV‡lÛù–=¿Šj-†Š’ÑÁähçBîaáûm¦å‡¶¶6™ú;õÏ_Û‘uH… Î:¸Q‹”2:×–èÏˆZ•ˆ6ÑSv»…i­¹ÂA³ØSñm¹0Mô)M%àåÒ4Ñ§to|\.I>%QÅþe(”|Ê•è_½zržE{ñ§JwxºO9På“4–ëé>u%9ó²U®D~ð©E
õ:x²OLV§Ç´ú²—fF3Ê ‹šTìíï¾Ø=ÜÝ»ÿjÝìÏž«(J}€ø1¤ ÞIñAí¨¹GfäV9u(´š?Š^w¯ñà]"2ÂáA!Q‡¡š@Ñ¤¬pUÜYÒF-,Ye%¨Q€Êncà2ïD¶t*Æ˜_TµÆ“ô¢_‰´ž“úýþÂ$¡êê¸n«ºFõ‹z–YëE<òdY´Û^Lªè1“EdhÞ†ÏÇuY¦^hÇµëoy¥7öž“öî˜ÿÌâë.9RçÆË%n©­’ ï›§v%z"ÿ0‡‘—>ÍÙÄA¶>Œ€’OaÊ£t­µ¸¸ðÿìÒúß.¤QÎ¶‘‚Çš5à•8:'ÒAòÜJõÜ ûÛ]y¡Ê‡ð£<Ì“Ù,7;4^CZlfXfk¿.ž™%è”°X{qqŽ,.2$fCcŸ	¸õC²½NÚ/`Ýã…Íè,îE³¼âœÕúdÑ»C%Ø‚óìÃ‡KK?Š‡´æ–ntõŸÃ(E`°€	W¤?ü½$	æ.–]âÁŒ]ý…$¢#HN„é¡¡ÃÉ8$$‹!ó0$ã0m2‡Æ…–ü#+ˆNö0ZfX€d|¯Ážë?‘öŒšèÖE¤CžoþRØ"M³œ
ClYAã]à¥xw3{€	±¹—Ö;²zÚ:Â“:-‹ª}«Ñë)i«Ñ"ŠÊ{w_ŽñLö&jzoÓðqé÷Ü›íA•ûgöTÆÕæ]Šu³D÷VØ4éœYaEXr}4—Ú6š–kŽÜ„¹k< ìÅ#ho.o·ÐUOÜOÒ·lA¢~keÎšZ=i‘éR{A1£Ò=»,‘=ÜÝÈkg&h\·­Ë]c/ZWÑ{££6Iî¤“Ð£úxŽÐ·MR“÷MaFZŽ
©?Ö –êÃ§ú›Ú¹¾Ç³ç?#Öo=•ÏT)ô³ ;ìZ€„¤ ²@ÖY(jd+‰
3b§äµDcÚ"o:îìPáÍG15ÿÎKtC‘a´oôrUÏG¡ÑZ-qr€.³usÂþ¤|á@ðvc÷ÅÖþÆöúLmãÒÜš~Ü?k¶Å¡ µÊ]¼×lwÌF¯DH›þfRs—šÜŽ6ð.—ìmÿiæRrqiÜx³‹’÷@v°±÷|ž]Íb˜dÂgq}…ùÕßÓØ¢Ö{ÿÅJx#OÚÿ€ØW-'Û›s³†©Ð>DÉ{#a2Íß@„Cq‰Iù3ª´cn«:6
ç+7í>Á_Éé€yÜdZòÌ‚Ü8GvE'LWÔø:…_Y9±”4w©©%šç’ßÿ^«£m¶CvDbÇ¥êN÷FCýh'+!–e(ýb…†aÚK2Øñ'ôÑù(ïƒùWÑa{÷“Ñ$Oã8Áô N˜•³¥æR›Ë»ÒÒ¤;45ñ™N[Ÿ)46ñ™Ö&>D{³Ú›DQ[tZ"9”‘ºž€¦³Œ×ö’“Ó{JtCôV¯‚O—´ö# 6 /¦².ñ¢Yß”­Þ ©.h×Aãx·‰£© w”œ…pþÀÑÐãÝ·FÄÆxãw5óónkt£\˜Dyô1n›2Ä§¾	×4ï!Ø'¨Õ™TU{¦j_|¼õ|©ë\«iwúaiÆEµ*Àlž\ UÝ{å›Å¼Gv‡:sÙ¦¬:¡»óÙe‡¼¥.ÙŒS8=±:„ÝT“ï:N‘E¡¹x
ÕfáÌÝÃ°Ñ_Õ¹xšÛ'â£D×²?ßŒJÖ’u§u|Êe‚A4:ÉOÉ*yTS¶NbŒði]›%ÇõaB	¥éx«9ŒFIF‘ ¤FYP+¨¤Õ1‡ìõYkçžhöìÏ¥yóõfFõeôÂ|šÕÕ­Ç3³=Œµ„)ÃÔÆ`Ž£¼wêµpÀ¶0þ¸Àôw½¦»¾sø£ü4A{»‡^õ‚ÎŸ¢4ëz6£@}Ø’5­ew"_÷\úeCË¦.)ùµ÷í+¥g$ˆ¨@íY‰²…»•ýì[‰¸iêzæ·ÜXy—†uª^nùv‘žëÇ0ß½xÖY”Oˆ´h‹™bè‡O ˜óÉç:s?Õ¡‹ò~>ÎBîi‘ ¿QOu0¤-4ª>¯¶!¼>è…~`ï²¨­öÖ£[2ì-Ñ‡ª°™ƒq•#(ÚŸÖKà°°£‚<ÅgQ¿ë	>êØ|JxŒ_†ZNQþS©¶áÕ©§Z‘ ›14Ô†cìÒòâ\ oÃ¶K˜]<€,	ÿçŸŽç§iò‘ÚlUâXÚŸw»Ó YMÒQ2¡G:TFŠGgWÄý´ÿpøbºôí›Ë ZsÙÎ²„FÇ0­¤—¤iDïQ2ŒsÓ§®HãŒmv^¸ÿL\a…B†'$uÐ$.eèí5†$OrQ5aŠ.‡ûPz
m$Å à„ÓÇðçgÉàŒ+ÃòZTè…FíÀl"œm6 {Pø8<KÒy2Écø 9H2fjm1-­-yŒÚWQ-…ñ§džü
ÞÕ_	œÙÿÑ!’ˆˆ˜Æ…IeaC/ÿ‰DÙ7ïnk¹ÆÄ2”PÚÃ¹ñÛì˜H 0~8œ'MýP¯â5H¸Xï½°†p;}¨‡VlI?~ð¡ˆÔ«¯Z<gŒÛ¬ÖPÜììwGíŠ¥ë=B%¦©GŸ [ÑšË”èù¥ÊúVO+—¥;ƒ7HŽC ¦úvœß>u4y‰ÃÀ¯xjîoÇêzWÝTŸª@c…š@:ªžÇ1QÞõàËŠ‹båBDd)xHãw¡O´4„Ÿ€¯u÷ŸöHR9hãBŠ-?®‘„ŠÇ—õ}°ËØl=J£ðõyñD_…K¯Êô2…ÊP+ÆÀ€q~ž'ÉàCœ·æ|V•v¯´²ÃšåE,«\]W³sàúa—äÔXÕYxùYg„þe„TÈ‹C;á3§ÕÞ¢ü¸9ÑŽ’õ½í‹qÈ·K³•ŒNžüœ$'ƒˆlóe$‡lW;ü{“:ùëÅYH#¿S.§‚kx ‹õ˜HÂšó¬ocLúŒ!ÂôqC½*[Øä()¹KBcÓc”=³"}ùñ_K“K7üÅÅ$véodÅáwá}^ £å"ù|Þ%,½`ÕfIP}0fÞZë4ÏÇY·Ó4æ’1:Û>¡‹„w‡(fÊ¨¬©´ÏyüLó¡ðü,Ž>>¥ÝË×Nðþ2·…-EKu®Š‡Y"­µÞÂÑÿr€Á×Z£F€–FÀ2©¶Û¼TeéâÑ Eµ—™b—Pa…úd¢¢^ÌîÜ‹«ú;¹®„cÇ¥ñ§7¾£ñÂC4¡ûÄg˜ì	¡i—–ÛïK›w:Ã®æ?ïÑhJœ¨þƒå•Z—gÅóeì‹“4ŠFúÆ`I_ÖÎºÚJ·Írç[©Æi~ãŒNÇnžõÌÕ1Ü:<ÏuÔð¼(óK¼´§RŽÚ;”ï”l.5¼âaºdØéú=Qã¯¢x4¹Íø7©¦–Ð^ýmíÐb‰©>\¸|RÏs6ðí-“ŸÚ’íóµE!NUE@‚LŸÝËêe%“INññÈaY.§ÃG·VÌSëÙŸzUÆâéÃqáhõ×.âŠx£n¢›(D–&þF"J‰¾ËÒ+jb¡¢+ö«I>FýzÕ¸âLÀK“˜:ásÎDà¦¶Öý(>Œü1-¬Q?	‚ÀÀÔû"¥µ{Q”fŽ’Ë}ßäZ`¦J)\ÒîK.Wk[¬¡hõ.SµWW–ÒñEVx…ÁàŠ*B×zn´×FÏ¦'x8-‹_”èL”¹C+ÎNq} ÃO"‹ÙÂ~Ê8bê“LjÖ¢~Ñg35u±‡¬PæÞOUé¢øC.Ñ–,i{}l“.+rÒåFRR§ü5ø£)_¼fg'MÂÀø§ôßa*<³ÖñÖ9P.ÔžÁŽã=BÕÓX<ð=K>­µÉ"Y~ ÿÕ•@)Û4¯a47ðXs=t¿:óÓúÚî0½pŒ&â°ìõ¼ZQê=P¦†Åþ÷óÓµ Åõ,T­õbi™üx¶<\$NƒÅ¥áÂ£àÇ•Çð²´<~ø¨,øœåàáâ2ðãý@–‚V–Vv ü zÁ?À·••à ýø€ÿ^	<‚²;+ÁÊ²ôˆf#ôS°ôã2Ya‘•?_›§ ugÙ(©‡¿¨ÓË Õ|À‰PœrÉÕ©%Þl£ÊýY’1‰¨ˆ×K"‚\ÇIÊ|^à}:u;’Í“°^,äÊ53g¸ÛbÙ’¸¶n†}ÅÔWÅ&0×ºôÎ`à¨®N¶K|ÍÙÞHàdT¬
gIïê?È8ÁIŽ1NO˜’PÓäj"8«ÓÕµ©³`Ñ•n¼,G+[¶®‰ËÕ{° ÑÕ¿'p~ƒE§½ÄšpÕy”k%4;Ók˜zºœ¤Ð~´WÔ‹—»ÌŒ¯Ð.†ååâ2ÄÍçÐGçŸ¾z°fÆ¿PØ Nõèê¯C#Ü%PÚ~ýBµÒ@@¡é››ÊXÀç¸ÉŽš,›kPÒv€«Œ{hŒ{èqŒaÏ}¨æQê¡Pùn`ô}è±Öb¶å+·Œ~Z«Ôà–xÔëqø© 2]HT_ô’nXº‰}n‹ÃI~ê£žaŸ‹Ñ/o¢%¿eY\ºü°?T}&?Eym&¡\K“i*}Ÿivø^aå®vœß]_}W•äXÕ0º?T.Qå’m•$Yýâ'_9“‹°¹• 1Á-¿V¸›©µÕ^ØZ¬´ü—sí’%q¬Ä·ÿáD¾F8©ÖAï
ËS#”ÞD„“m.¹Jào,Æ‰‡!àíG:ávùuaD\UPåÊ1P\0—Ö"H­«
 Oƒ<9`–ÌØ§e±˜q±ŠÙH#€>‚Þ7 s‘u`ùàª(‡ÎôC(ŽÇ¥$gIŒ—'#üœLrv¶¶ÔÍCÏ@½Ôæ w;ùc¼jö§èq~JÕõ‚{u11²·ë5Lk-_ÚZZ^yðð‘+4ªiaƒ`v{Ì|ª­Ï@€¯®øÆ(”„ç/i ;[7.­½àp}|ËµF@ó9«•°¸¡Ý÷	ÄíôXƒöYs¦vYh€Ÿ¶3Âßö:J¡~ ÐöTÛfêGícUèïùÒiÎìex½N7üHÔ®T»2„Q»6†Qm]2^(¥øö|C`}˜žDayŠ†ÌŽ™£ˆ¯¿žw)PDé! [–‡ÃqÛaRÉ0M³R–°JÜþÈM	¹oA~\Oq¨Á•ÄO[ëVdA£xÍH¦žÇ)õemØ
ÉG¾™ôìT£ûGódcwggkãp{÷åAðê`kÿ`ž#3ÇÔ+Ã´e²Û°ÖÊ[‚À¨È-ø­ú8ì§pÖƒÃ›^Yé„™µå†u8 ÒOéþqŠÝ)ød€Aè¢p@é^¨dw’·5‚miµÄC7ymQ>†þ†³p¤%#mdöK}œë¶\$¸— º@™þrìr&aYù‹×çpüdyéÏæt½j”~Ô¸Aná:
3è<³
jãv(^°ßúçñh˜…9§äú)iQ'üŒñ~õ4ñÄ‡xëã,Ú-Nä©”]78ÕÈ™l`eHýmm¶ÑÖ~»‘nØY›áÌÆäŒ¨ýv×în#¡}gVyî7¸|¾áÌ~»kvG1@¾FóøÍãþÄ£øÍ£Ràk4Zé—íÉÿ¾Gó¸GÑ8¾†ß(Æò5ü†ö¡¤€JÌÿAƒá õ$´ùwÎfÏßDTÄZëQ÷æ÷ŒxKá¾“o$Äáµ‚’ë‡9ô	txmãÞ8nmª;€…eJÐvD{Ä]ndF|#A”]¬×lƒœLJÙ5÷~Á‘Œƒ4C@ª4ëõû<V¢&Ñ×§.ÑÔa~ð¹_KûÛŽ1´AïèœŠ•³/T|Ë€Ì*TºP¾{hš}ñà>Q¼héjÃ«/h*^–aí‚ý½¬kü‡¢L){`jdº3SCì‹lŒ½²ÙïéÝŒÃr‹k«ÇâXåógTÆ[NƒÊ0²†=†ª¾LÓøä~Z'yRÂRFÿlÎð9€êªQ3®¯,ÂŸ½`sg `¼Ë†< åûÑ	ZR=zÌO6Íáéº(Œ©|?eŠÇÆ’eÞÃ¤‚lœ'<8-ÆbG8uÆY%Mf³lŽ‰sñjá0Ùú4NRtbÎaŽqîCÆ³ÛD{LkpôÚ1­õ`ì¥è šŠo1	ÖrtZKÏó¾GÉýˆz™î'åâ©øàQ	vû Æãr×õDw÷À<k(†a¨FÅ³2õ,$ÇÁßÜP‹@ÏëÊ]VÊ¾ÖÇN^ýµ7Š{	ïåaDßPcoÑÝU­ ôÕ«dÑa­¸ì±W%Â¯þž-üœ†ý	3¼pžd,±úŒÁX½w5
ì›ê*v€_—sAwDñâ^4h«8džå- ‘·ôâÄì»NhU4Dïö}ÊKÖG$–™&ÅžêÕ®ÂpWí€û.ù 2ò„›ÉÇÑ 	¥HØ	›Š°êbÀÁ=UÏ¾®.¯=xÝTÍ0Æ¡¿W¾=´ÜyÎ`ä-Í¼sØëEã|­|dŸæ	þ©e²íá«r	§ºaN¨:ê‘B|Ïž¯ß¸”l¿Áls5a“b:_ÏÓdÈv–™'Üè·UÍ¥,4auDh°h¿–:Fõ?aºò‘$Ç„Uïº·«·?¦…ú^·h½™£þ,âÑÄbø"Î{$=n
TgáÚ§¬†ÒPÝÅžÊctÉË	º§iÓò%†äÍœ£ŠGc=ô
ËœŠgnD¯×ÈÇxV^æSªSPb*ü§ÁU±‰ëi:®ú­<‘÷´p¤:‚cñŸcUÛÔtèÆ«L•÷`9ûR¬ävüGk¬Kç¼š×X¥%ó¯ÊÕ6}™#j°õ­›ù¶Æ(ÀZ¿ƒ¡óld:›|¬6$ìaŠþa¿ÖÀ(˜Ä¡bkñbëpýíæöúÜ¼@ø5î Aûþ{·ÞIa5P§é£›`)õ7WLk_ãïIÖ—‰	2>$XÆì›Ö|ámÂÚ»y@åž¤Õú	­6š¹+#«mŠZäÄFyhIÌbLa—_xÉ'Ð‡‰ÖŠËs?+2Ú ±hw`µç‚<Ù>Ø¶A6Äy»uØšÞÎ¶7tÎÂºýJƒ5Ÿ‘°æ®ÒuW¼óZ‰¯£#’X¹zPŸ©JümJŠáj³>S•9ÚÔq³s½ýrZ°¾õ>ÍñÙè„@I|´¨ºBæçC¿#ïêÐ+I?6G
±rK×ê&·ÈÔ!â~ ŽiýA´Þïó}i’|äÈn]•%2ìw}TWl—|ÕÃzõh]xö¹ö3ÉiKá B:åœ]^g!%hYïº˜Ë^<j}g÷Òê¸'¬=»Û•ÎF'…ƒr¾¾T–l¦²žÊ%B@‚ (ªŸç DD-jVµ®>J·•+>·•»î6ÜìåÃté z$1úÅ	é`Ô!Ñ aæ¢Îc³£–rÓõý*NŽ•«sîÙØ¸ÍIÏÆÖ;ôCŒjIà4`nÛáÁ¸Ò/}Âï4,wéG©Îä+ ¯2sä{+«i*©Óí¦¥9émZL•ŽÛ5­ØÆSÏPõ©ÅÁîSu[Á¬’^á-@Ô¿nÎ¢$Žo PÚlq+ÌcãâSÁTéÞæV€ê¡Fîg;ÊIð1œLš­¦åT1e%S”ñ&ï¦ÀÊ˜löqQ5|1Q?Dç]´I‘rMÓ™€riíEÒL›¹¦½ÅÂ•›ðpX8Cì,-K†À,µTÂdÌÛ²ièÆ˜O®ä6@—¡Œ!†¼}ÞJ2aÏ¹ó(Uš¿$E>TÌ «Cvä<zä5Îem¹Fó)äÍžéÊ^šÁxÔO1ƒü­v]È©se2Î™½@£ÉâÆ9YT¬k{L”@Ý|¦Š+ñš©rõFN•+“qªìM•*÷pÎWEäí1cãI:ÈÍÉßê\]—
W—ÓŒäÜ¥*Ó÷†+`áî²N¬š¼’f=AËð³ÖÝÉ-,‰MýO9,vLÍ~>˜\¼[… –°4(±YñqRp1tµsú`
S¦’xhe6–MÞLdV8Þmµ×G	YGv6NÒ9ò­C×Rg¢že²ìÚÜµaï*BŸ×lÃ}´Ð&Ç´)øSäyãºzñ¶2ñ»%`.8rf•x#/ïJ+®ò#¿Y›³ÜL¼[4ž¸±MvÒª»Ÿé†jô™ïÁ×-£<÷zËª#ÔE	?‹Mã ¢Ê~’ÜškWMx÷f»É~ÝtÊÓÔÌâ [É¦þà¶sÛ‰y†¼§—Qt¢•9öŽ¾Ú ˜ªå~Ê0i6ÙZ½Á5Ù9ÿ{]Ä6û?‘E¸+ózéØÄ~æ;ìª‰›b¨¦¤b8­CV+}—¢g9­³šYf•c<¸í²n: ƒÉ` ªÁÖ`®þž¢l–¯S…Ö‡aÙÍˆÑ ª.Ž˜Œ}:@›=àp‚ÊY]r! Ñè$?­ÚtËaPb(Mº\"£éäÏ)¨Tw>›ÎèXEE-»pl¼WXÍO£°oDcyZ‚3yñk3ë œÑõöŠ$],	ZT)aµ“Ÿú–ÐBìÐß95Ša`ô÷ÃÅÎ²ý¢ãYðJ—ÃÁNÏzèPÇæè:]¬½¼»íq™µ«Kº(ùvºÅÁ®nm­oÞN_|.Ítyîíô‹K4ÐT›Î¼cªxð¡½_ìÒsÆi=Ygì­HOML®æGIÿ\m°; ¹…sÂ”¦N~lëT¢3Ôà{Ñ;~ÕÔeñÉ~rq ¼ÍÚÅ^úxáqÞ
·$q¯¹s!æXL*C+ì¥ò0¤¿¶¨§°8FÊf[LW`ï×plºÏrA–¤y»Î“#*{?¢Ziug¡çí0ÚLUÍêŠ†Éï‘wª±àb–Py’¯¯;ÝibUšŠã\æ.!qü6Ì™µÍójwê^%*f§=|91ˆ)jgF»|®håŽÕ«ˆÿä”YüŠL5=ˆBëçÇpÝ^?dË½rƒ‚þ3fÖÔhRÐ4{ÖÒy>±CXž“(g‰b Øåq_±×YfÕóÂ‡­ÔD#ÍoÐzÖHû×%Lž5ÏîÙK­¿]ªW÷	s“åV[eÞ¥¤Të>—¶·JÛÓëL“,jZ¡V”axWpéX?^µÓŸ\î
ñÎoJ)Å}wüQÕ¬¸Î¯³²èc<^<.ñ%çeæÆx˜~IÅL3J>¦áØ$ŸiúR“‘ïc$qhÉt¸¸Ø¥ÿÕ:ÔE“šJ(±nYÓçÏö[V·z|èV*/sT'§Ä†z eef:ÁîÔ–Ç°¥Ušl¦¼î»Uý=~BÏE·W°Û)++Y‡P¨Ë`t“úNF‰[ˆí}ŽÚA]ôF>Á–5jàS`Jrm¯ÀÏAŒ¡¿¥ô©°ÆÌ”f©	_ œvâ\“vàz]äv'èì¾îÓÙ¢­û:ê{¦E,ŸþømÿÏƒTÏ|#¿‡hà:<ÆgÂNü¦9.Å¹¯;b¸àžù"ïŽ†¢õ%ÕçN¶ý‚Ág´ßUì}Ý³ØôŸÃ¸ïáÎ÷ñ’j(ïé¤U}îÜ4T|Fˆ r¥r_wÅL$ŸËàï!JðöXl¨¤‰Ãcõ¹1Á,€ä&nBäýÔò¢âáéÚ¨ã‚j*\~	ûœ©|!›Zß±\-`ž«OØZ‰Çoº°Î{ßÍo>>¼tÜ=@ÉF¼–Ÿª#DªWäv¢©ö¶()æ&KÞûaï²oÄÊ}‚o5%ß‰ÓIñ19Wô,z]Q·ÉW£QÂÚ`RŒõM%CÃÇæôqÊ	š®ŠI]ØÔ`ZôšIð©º ¼†dCsSi8O7—^ÍÇ%mp¢¶)—wš¶ËÎ3m§Š&xA­°ÙP¦á+ÿ?   ÿÿì}ÛrG–à{E
£õ‚Ýx‘ä¶Ñ–IÉìE6Iu÷„V!Y6PW$jh>ÍÃüÃ¾9æ¡Ãá§}™Çåí9y©Ê¬Ê[ DÉÊè¶ ïyòÜò\DÑÇìœ•Ì7sL`µ¬­‘ãAš ÿHB³(’l’$ùùøWó·Q<LÞv3ÚÅIâM2òdj	zZ-§áyð&¢;Øìœ–E¬øm‚6<jµ(üÅºËmv¨»ër~jb¶­ž:ëªžvÜ,ÜfÆŠ›Ó „å®r³$ÆD#µšÍ™žŸ¢ƒ¦…žÁáoÄ£(ø¶"¤µ´çû€„ƒñ,JÉ4L'AŒÌÜ„f”3ô­3¡oy‚Ÿ-_´RËs_|rlÈ…Å\†ã01ìºwC‚±×›T‡§¾{ÂŠ.æ{³¸1f“FBQÕ}bÑÜL¤¥q7åþ×¿“®75˜¬¿G‚ø?èÍ·ÎÝÖ„KÜ3q¥š¬º¡¡øodCüèÈ{ "…ÉtAEè7·‚Š°Ã[±®(«Ø…~“++–L	š4úè¹ê~ÞÔùVëÔ+QIGZ©cÎNúãÙÓè,$a-q¨!§Åv¹i<ß_ÒN‹x=¼€Î›¹}õL‰…Wss˜ªñ$&içž> 5M(Îƒ©o®_éÃhëUX"TÁ_Ù™º°­šì¢}B°¾9h%¿¨ÚÖ'Ãhm’7œ›£Ç£¬j¢Pq2	a7Œda¡ò$ˆƒyaLÐqž¤A%æ~I©01ÒÑ&Y(ró•:y„ß’ Üæ¤ÅÝHÚWOµÛ?í‚g^uClÁÔ–yáÀ¢ndIÀY-”%¡8.Þ”%äI®$'T|³Êoã-Ì• gõM—PÉP§K†H‰Y–h³>7Ë¯ðÁÁu'
2\ÃWÇà:ø¶bÌOEØš•ÖqxtFa”FŒÞ:	Óë:kýu¤¹G­(v×:†ÍvWºþù¸)s…dçbªÀã·£À½•!F¦—z(B^Þû¶4õ˜/Ú
º,²•Ü°ÿÇõÞfîÞNàÏÁy8øá4¹°É@´N8¬ Ë‚qèF1H>Ã0£gcð¼“Ä²‚®UòI¸Eq_Æ¨,U,íl óL•Æå¶¬b5gRU—òõŠ„c@Z·kÑ8EJÌGóÑÒã’; Ð]¡ö°8•RÎ´üjÕJHø€ßáº¨d ¤
B@sã0†8zXôF‚Ö	ëß$9XìøÏnm`D©7Šhóï–Pþï’:it!f5š,–Ò«¼u\"[.¿j„	odQŠ*‹kjð±r¾¾Œ¬þËå±”qPÁFˆ/—ÊRV}h–¥öÅ»RUú¨éyO±)?Çž,ÊÓ(Ë|¯9*õvT•›‡Âm	(i`´¹,•ÅÃOÚ¶ýsôÉbŠµè“¨hh®õ-Æ½.¤ìfMþÔ#ÛÕ ø†ƒÛÝ¿MÁíÁ0`Ád1Jº‚KÞWt»rÿ_V¢ÐÕ–HI~Q’–!É"ZøGF²t1fý)ËúVYŸ‰}À~šïºQFÿeAË
9ÔnN²EÔÚÝï“(ncy{»^¥ÝòV*S2ÿ@©tè”¬¢À\šß.ÜÇò©¡Á¶î]ÌËX»¹N„iç`n¤™"Y©Â¥G“*Hz4Áq^¨­ÌéÔ±Ì¯ƒe)Ä?0ÕÜhì'…e¤ÆæÑ­òL÷^Lñ]}sþàvNÛEs“9n-{ÐÜ¡Vsê½¢æð¹Mö=ŽÝuÚô¸·×†}Í:Ú(¸·Çg”ŒÏjS´l×Ñ9ß9˜¥)Ð’çY˜>†ä!™é¼+}»ÕEÃ?ÕÚÍà§Ã4E J=¤Ÿ²î0A»=£#Í°Ž:‚£”õÅTÓáð(üqf9F†“(Ÿ…îº†?j¨¾DÜ‘æSòã,üS­®Tµ8+d›ÖÇÉ>0³ y]äéÌ»‡§×¿‚4¶öxço­úCˆ1~'¬©›%p{U M?*º;ÆTVßòôú×Ál´ÈO?ÕÚ·Hó‘AÁ8ôÓøË‹ƒ3¬‚z_Õ6§R_¿WOà\“”l!OiêÖúÒLSjÝ\´$íâÏ—,ãèæivÝÖÀ<#Çî™ÀöJwi¦a<¶k0o¦z‡´ï=mƒ7ífyÏ26ãCè·TyÙ]á¢lýâÓi
ˆ}¸øÀ}ì¨úæl8¿|¸ŒÂÁ,3,}©ªpŠ†[Iu¶ípdÝò³Â„9ß…jafY%_N1ŒÍ¸(›xy™Sš´ž¥^Ôüä®—ëÑ{¨04²”`Ÿ¾”ŒÄ6/Æ&Ö6}VâÖŒ÷µÖŒã\sN9û;žÎú=p™Þ¯Ì¹*øm«jZÅÌÍÏ{ú)b[6Çq ÜÜzK{,7;üG giÆü±O¾ð‡†ÛQ:‡›3
ÜûÁP!&¿98âU@!þí	ƒÿ`ð÷1ƒŸ Àü–b¹à§ý²úíûz™¤ƒ5¤‘À] §AÿN“·†'K`¯Òë)·ÂœGÜƒšWÎât½^9Y9æÜ9{‹@Aê1eÔüqý‹&Ûž8©ú‹'ÿÅõî‰6
“¼³Þ}`œT)’˜Lƒ4 uQf™ÌðÖÚQ˜Ñ®~"Lœ‚e²^'¬[ümþ·HkžZ \ór¯)·Mh*#`@5,íCgië²ÎÜ þ-²Dåø/2“y®éÉçEë$|ƒO	¥À#Ë ’XðóáR¡C¯ýa¯L#û“CkGŸ™FÖº²F/óc*Á°m‡Á}Þ¦._Œ\+ª¸»6í{&Kå§‘U+·EZrWÉÆE¢ŸJè"kw=Òª\±šÛÉWµ`HÆ¯^›¶Í¬"4Ÿ’]½8½]íûËq £Þ—õa†—e«{ÇÈ
zF¸Ø¬PtÓqŽ ¿œo¿5ìgF¾2í•4¯ Ž&8Z6âŠîÞ¬n7å2®²:»K	
+ÛAš†gA<LH¦Ð‹nWÿzÝ˜À§ëš¦ÂÃL N®ÅQéßóí+e¯ö¼Èÿ¤ŽÛTy¤†ŒMd„#¼ˆN#`j“S’Íô/W SÀÁi4Ž†ôÿcÄ‡ö­Û=Ëùì€Ë®‡·Z†uèÊdã#~7ZúTPV6Ú½Y2ûËnÅæ˜;¸ÝÀ{aÒ/@óÔ˜ ø3‹AŽ«·“hº@ëÃ¸ Âó÷ðgŠÍ¢AGo,ÖH®n)¿½Ã.‰7·øwm”ÄÈY*ÙMÌ/«èÈ²^^LÇj­¢á"8?¦S®«EÑÝÊÝØž®múŒ}KÙê=¬˜ú9Êú\YAÞX¯Î¶¿p‹ÞŽ¸ìiè­`Fýz;dú<CgÇkI0‡em|Ž§H]X$¥ ÇÈx!yÁ]ØÙYÛß_ûW(Sb­¥˜³{pvhÒíeVBŸô¤Œ|Ž](ÞD@Å×Ãa7›Ž£¼Ýê´œæ|ÚT&Ì÷íÓ}}÷’¶{±ùòjMü½!ý½þòêµ³Ÿ:Œ"Çæ‘‹+? *–r¥	Äb5„{Ð0U "©ÛD•xÙÙÔÀ$kUyúx
îßv¾"çðÉEer…FÙ§µ~Mòuéª¥8Né½í€·ºƒó íçíõ7.Kë¹Æn_»/î)EUïš&Ž_wƒê&½·É¼Q~Ëµ«-=Ö!Û€›­é—²*|®¾‰Uy¹gEfsÃÁ+ø0Y¯ÅÔ‰©¿¹¹.©YJ©£fØoÕ·ÈEœ]|7·/`rš;*’(L3¥Ó±Ä0òÏ´¹=VK¯ÔìýÐOÞ½˜Õ/rñƒµKyo|ºõäÎìùé2Ô.zÖb5°gÌo,—2+&Ø^åx/F,g¬l¶)¾XÀÜ/#– ¿þåfçvS;û8š,u[…ÉIpÑyK¥Ät%çˆi£Bm”®!>{P\ÂïeÁÜkó€{nÝÕ·¿j¹™f,€:ÌÈ­§"uÎ6n®”ÂlC:ÂŒ%
ôÏ>6¤"¥ìëÝH¯äÅàø¡ÑÐ‚„1‹£¢öñÃ1I«ðÅÜ¼„ëÒ7h°6?6íRÒ'`?ÿ˜¯¡E S‘,­u£* –K\ŠÎ®l9LRÙHóM;V-Ð^¨àÅ×Þ›I[‡V}Égì+CùîŒ¹yxP"“ûcm<o‰Ü'5ïÓ;Æ3+êkk„"<-Ñ0äúÔfA‡9x4‹¬7u¿HÎX|eCRN­6J#;Å[¤ ò«°ÖuD!oê¨¾aûRÿÃÕ cÿ¶ôÚ²SJê¼™óúŸ8¨kµ]åý‚º°Ý+Ý¢ÂZ€ÛæÿXÄÙéÜ(ˆ{+•Ì^âµ.çÓ½y“zI¢¡.B‘f2Ë§ñÎ:–ÀÓt Û{Ä•qw·âÇºÓ›øIoPRóˆÓ6—¦µÒ6ûÉ&“(ðŒSýÕÆ©æòù}ä¬µn¸ÁÈÕÍl+E¼ŸGtë¤èÏÛA:li";Óï›Äqæ=ßlg6ˆ;‚3¯g«ó(Ÿ-—	™„¨Ši´ÅÙ€:FO‚¾öÁþ&W¶N`‹tÄIá‚zcÏß¦ô™¿ý"q¨íQ§k±ª?L”iå+SiS"Lk¬káölÁPAøŒ±ö¾9I™0÷|*˜„ûWFýiËà‡@ÉÛÉõ¯xWÈ¡ò{Ù3†ð3ð³„çSB ï^ôÈvÿð¤ý×ÿ~@='‡G»Ç»Ï¶÷úOÉæƒî†Ž¹*Ãû±së²;¯ƒY¯À~*^á±ûØÇÕŸø„ vÜšeáãÏØ¾øjP 7nÕDRÖáÙÂw€`<ªÅ´b MCc®FPÈž¤xZtå@ºù'¾GÓZƒÒì³EÎÂ9WS¥ÊLË#RÃ-:O‰*XçÚb­øžæ‹äÇ:Ñ#S…Ê)Dóùqèöœ1ÅÙ?ƒ bmý"^ò–]ø`¿Â˜å–ü‘\Âøm­Ñ]˜†#×bâóûÞ	þÈè‚IëW,ƒî¼ÞÐýü¾/œOÎÓ,lv#Ö´µ—“
“QšçÂlXä¾½gÇÛä÷FyÈŽ×BÐXaGã¾N˜Å~klZ,œ€Y:»Ì\'Sà{Q$ÀFÿø'ÑEÀ?E'Ö>¨ßMÛÒ6®t‘õ4V°úÓøG7gÖ„°sä!yF ­’d«€ì%:°¢kÍÒö¾ì9(¡ã¶÷É-ãìøé±š@ÛÒh9Ã‚31â+Õ¬Q¡?ZÎÂMÊÌ*mú°JÖÄfo³*ëÃ£²ýþ	y¼÷ì36£å=a3&?ÆgŸ>ÚÏÜŒÖßîï|Fi´¼?”†zÀÏ8í“ÁiL­û©5
ù±R¾;õ–—[Bç_O-=eeA5§h÷G5‰•6£„0óvç†àÙ¨V@‹Ô¡ðSÌßpok¸ƒüWTw'cðÈ†€‹°«X6*–Kö<Êô]˜ÛÊ.lfp#«Dµ–½§ë
¯yÒgø©ÄÌì‹îÅ|ŒlüUÀI¥Ú›ÁÜà-V³5š‚1š›SgJz.ÔâÜð(oí¢ßÅŒñìöHÞáÅkoµÍ4Ä› ‹7éØltKzÀ¥-ŠÏ®FôÁ¶hâà”çYÚHúÆÖ=ÈÒø§¹ªþ.)‚·‘Œ+W}Srz°n×ícûddÞ|DŒ¾¯Œ¶Mvc$Oh²wñ€¸`:OßY™Âàmå Þ ËõN2°³ýC¨0<]%ÛOŸînŸì<;î>:Øµ³×_%ï:˜;ƒšfã®ÊÚû´ heÈ“ø$	\…¯)“q÷Ršš Ðp%Ã„šŸÐoÑ\ð
ÓÖ¤	` ûzÎi]ÁFæƒsÒÓÔžÛRL¿µ›¦	¦·
hò«”¼‰²HÄHéb”)è*Imné¦[ÔèzUÃ%’q{ZšÛDÔÊyì'ÈÀ¤$È= z2ôdR<þ­i¡îy¾Ù}ŒFÊ½‡Ññkå;¹|oÈ CpÄ¢ì(J'íÖN˜…ßcøØž„p¸¸¼­ÖŠü\˜Eà–!M	àÄ-KÁ.ŽÈÅ¡w:'É¸Ã]{ÀùBöPáÊ8[»—boé¼î£Mc±È]-S*(×T¤T˜Jî¢¤ð$²óiaSÙWýFJä+½ŸÂ)ƒ„ÒA;!ÉëÍ–Ï
eà¯´Œ)Ï[…Ç.å¬ª‡©Ò0Ì.[¯E5¤á ¦òÉš Ö¥Êå™ >æ›'¡>.öè5?-¾ib‚Zô|³F¨b·jQÓ„ìXn$³jéíMóCê°¼ÒÍ“½ã.I®ˆ˜9'­•ëÆtBÜ»yÑn¸Lô¯';(g*¡Dý ®¥•ÇÂÁ9@Sœ›+þöLO+­4†¨ªª›ÐØ8Gi¯³`UÂ{R×…Ø™;•´ñ1ˆ:ˆK\ Ùé%šœ–£b<cî½£%àïâ‹U~÷—o2º¨‚½n{LEäÉÍõÍû]M¤O“éhåKO£èEÐ[@öâë_šdãÍ†©ùQ¸×mW3*úÒ Ë€`+àæA`|²‹pø¼…FÕŠ·Î€Üßö*ø‰Û†Àn€kqÒ–©¾5ó‰Ùx¡˜”ìÜí@G2{©7øŒ–< îH0à·à
ÑàvÀ›$©|·Å\)ÌÝ: SeÍÛyUù÷S¿EÍf‘Ób=sOŸqÄe<ã6‘ÙN#I;U[ógëµŸØ½ŒôóèÍrŒc\)<’	˜ÓÑ˜Óø'p'0¦A¨Æ®¿Ïƒ8¡6Ä¸ÞÜÎ¼ßÜø0Ò,ìÊ<=œÈ<ÿ	ÌÓ° HÝù†»é“`:ãÃ` à1OË>ÕÎ˜Zêã†S4Hb âË@ÅÉKíJºY’æív°JN©Æ= ª¶î¤Âíd2Ò°}J¿Ô¦ì¹ FÖ·Z‘o7 ?4Ô"d$‰Aä[²n¬ˆÑaÐ’‰%kJãò{»‚ŠõÇÑE8loØâõHkÝ¿œ;‚“sò%'1çÊ–¾ˆ³`ÊÖ@ä;µÁòY7Ã(ÈŽÂŒÆÁÆ\À$ý=I¡îïà—IÅø1*¹:ö–Þ6ëÎ–³ýöamºÖMnYcRºÃíùy78ÍÚ|óVà èw“à¢½±ªn†=h–|Æ-vf"°ç `IÑylš¤0ûMÄûRpç&iYE™Š!‘›ŒSò4 -&RªjÀ£¾äyVî´\±×à:À©eaÛYñû$ŠÛ­µÖÊÕe‹8ývÜU.‹+ýñ,ü& ¦ÃEø'äØµ|ËžW8	Ó`<ô›šU´èÌÃûÄ~®ädÜd±ÕL Rg.ÅsžcéJh,²Ø˜"€±¼S¶Çø»+z±}©4s8ÓÕÿ°ï‰3¨Ý²ÀÄ'åx•ðÞ¦»%4ë®zLÊ:Jk3¾{)X¬oY^Â–ºJ»57<9–ÍÇ{'áM©?wµ6´:ÓÝ*ñv¿°ù½ò¾<Ãz»d”EcÃ4r&Úâ}TŸjaŒƒò+¿FËkú$Yð~íäG'™Xúµ–ž$ræ?²¬®¢:Ÿ^l"e©¹ŠxÍ»“xÕ­»œ,kö7SQnÔS¥,®«[õ^ùÒ#2¯oPÝÆhÁÏœ½,n[tQù»”Eg›þøàhw»|²JFnËô²øæË|aF^Ž0e±ºÄ”ÅëöyºÉ”åuaÊz÷r´4™†³ö÷¢)KÃuºÜn|»aÎ Ë\ûMa¥å:ø”ÅåêSO^›¥õ
òjÙÄs¨,®Muc`ù¾ø€aÅñÈeû‡,—Ý—<{víÀ­¡èå	o)ß;­sªÏ
[¾wÝ£–/jó§½¶Ê²DéI™$_«‚èxzu•¥±W}tï-ª¹„	 ð§)†÷^ú$¥X6eY¾;ZY<ø]ÕEmy¯3{„]’·åJ0¦.?‚!3‚6'ÂýË¦ã ŽÃô£p/kf ¢w.£._‡lÑ•=­Y‡øùÈðÎnÞEæ}:Âôsž^§7â	Ãw±ˆq=ÚÏñ>æs¼–1-‚<¡ÏžsÀÔ#4I
àLR(ü4&¨+No®NaË©
ðÑÿg/‡K vN‚¸	dÀàtPP?êaðî`ô÷0üá¦Îºàö¶ÞJJ»ÈÖqxQ¥úPé$L¯ÿé¨ó×YæÎ:Qìªs«rU¹þùÔ”éÇÝ7!ÃòM2¥µ"^	à^™¸KƒÉ7k¬']»zhÆ ! ÿF¬™ËOžÛúlÓìùðÈ7l£Jãl–Ö=m>›5W7<«é™ËâU0€ó4FÞçÓ0³å‚¢Ÿ•­ÝNè'šUšãÙz‘&ZÑƒ<Ñz$Š×s’)>A©bÕ,äŠWÙI& Òæ*/Mú §!%ns7Ø½8µƒ’-Z!õ·Sé·¦Š=8PB<µb»™'%ÏS«‘çrL<-Z¥Àb¬\)Ó®gò°îtÙ8Á<œö²æâ²4…¹Ø„²E¦a×çµb®—
ßpmh¬ã×ÆÊÊ}ZÿðiÿÙ³Ý£U2õÓê{è Mþ¬¦‰ß®qv(¤ßƒ2Ú¡+n¦„v+ m nÖ,ëEŒ£@Æ ´h'ð›R sO7¸šRÌšÊÏ«Of:·>ysÑKº°Í7§i[ùÞ‚l{—q‰wU*ˆØç(ÞÒâA–Ò.ƒŽf7LFØúv/ÎiÄE']nAGfAvu‰nÍaàßC¦Sˆ)v6–5Ö5ö0@X©iê×*¢V´ì Ñ‚Þë–a©Ü4­Uï ¥beŠ~o:XÒŽ™§ÑÞä{êMæÄz“/Â‹>‘²-hÿ‡ƒ^“?Õ­‚ÜQ4¹Õ`+Ïï7 ³zŠóQ¡ý ÇHmŸIÿñƒ }šhï6^dy~¿‹ü¡áöc!>·`‹Éý õ#';˜‡ï3Ýá?~(ºƒ‡pï±2ÁßÀ]þÐûQžÛ
²åì~ðÚ€ö,ÓBhóCXñ” íïgã ?ã ¾{KO	ú[4š7ê¡¤t&í(ÈæŠ^Èuö€ˆr.n—ÕYÌÜÖ$f[§%É2ìHŒï·K´!EØÂƒYŸ‘›[IhrÑz˜èsÐN«ï+Wîp\Æ¾dM¨Å‰Õ˜i×°Â‚–yz¬KUÃ>bù­ì&µ}°Ûb¨Kµf¨tÆ?Ã8~äú¿ó0ªRˆ²Xvì#0¦§XºD/·´ÔE®e®ŸÏ_n—  ¬›Û8¿Øl?2pâÍa†ç*ÀÇFoÊ9ë-GÈF·_c>­ßîÑÞÁÎÁ«íþáI»àiÉç4Æ+¬ø
ƒ&oK>œY­ønÜ†Ïj×Ä~Ïe½g¾&Ë½[d·÷ö<È³`:õ4ÜûÒßpïS’°÷â<<KyJºG	HXGOþç±SÆ¶P¸’V}@áçƒ³€¡ÂxLàcsGg×ÿ¢€i€ÔòÕ¸‡¼?­œ¯®4ˆÿAYˆÛêŸ´Â£MßjyÖ°iXŸ=¥¾êÓ)9
¢ñÛài÷÷ð;8*â¶h^µË5¿‰x)jq)&<¥Øâçù4ë­ÁežuàNvâ¤“²uñŠº†£`6ÎÿÆtº§IÎ€§;KÇf]î£ñ,-jè ¨@´0‹ëótLVô¯]`—&m‘pPbÞí_XêÜa•ºœgòó6Ý*Ø©ÖÊ|M3ÚÖØÐNÏ‹­x-ºº{É¾»zm£–KÐ‡¼#s3ýÄ]Ó 4®ú
÷	2÷ë
,Ä˜•,ÌœŠ>ûûÉ«íƒg÷ž¬’Éo98ïkêñ½r±:Œºõ¤aßŸ~"£`œ9Ã‚²ƒÃ~ÞÃH@”N¢	†½ŸLÛŽU8Â÷]’I‡ßƒÍG+`«o¢‚Wãh©QÏXÐàŽ_k3É+NÒk)cš|˜UÜ½„º°oYp^9"zF²­Éð‹‘÷4›Ußó1«%ƒYÖC¸RéCû3™åã(wÇ¡ÍÛð&ä’‡ÊBˆ'îT°ƒÏ™·ö€ÇN2›”¾NË‡Q
òÀª80+”º0³¹p“ðÍ@ Qä>†kÐŸFíÖZ0Öà¾æ³Ì>C?è> bP
ÕñTïcÚ5‚8N¡ !niü¥åôl6À˜Ô‹ë»´Ëû£° ÷JÂÚÛGÇk£.u¥G
~jšÀÏˆS¨IuÏ$A<Ïÿ€
×?“gá)¾@u	¢Šn?*o˜EüâES‹Ô‡ÏÍuÍ¡R“›âämL}Àc1É™'ì ¶¹8`8Í²ªÙ<¢.£pÓ‡ŠÁÃ$72á_0lH›²òö ü W+Iš)µÃ‹ž¶GÁ[MÂ´êÎ¦]‰^éj¨ÑYñUk%EöÉÑp‹¡ûÜ-SÒ.n@˜Ù>€[ÝâÌÁ4Œ€¦ -jî>;:x~²{ô
~}õ—Ým*Ü˜e&©La%oaézÀTd•ì‡N’vÞltºÝ®¾ºI.IŠ%ÆþKøŽâOƒÂ²‘ ÂH0ŠØiCA…3Ô´%0Ôm×tW\\¶ŽùrØKà¯íÜuuu=¾vD=/—le{½9dð {-~ÅÍ+iªl‚„+¹c¡HÆ¹4b¶=¦]e´ÆU4â¿$×¼@í÷b{[¸n½Öèß.@??“ë_€ÆÉ› ¸æp¦\S7€>&ð]‚@Åsç¢ßëk;l§!Œ{ýÒ(ÉHH¢8›M’l@4Þ\ÿ’©gþ©’Ã}`<Ç|›€ ëmJøŒqã
[HIcÑQídIØB6#L¬ã*iúÈIß¯^±Àœv´ +HBbÑ­Ü4²ŸH3úŒÞ9>/M:<oP­ð¸ìÊ·Zß#Zˆ°ÓÙDlrû0¦ _õÈ“p².Ùì®“Ç°Àó[¸?€A‰?ß?ØÙ}zÌLÖ&ö Hr,Â	š¨	ŒD?XLÁ&Tû`6³ÍU÷IŒá	éoK!’ ŸäG¸Rb»€48Ë0é)%‘Aèúçq”…™xìÂ@EÙ§Jã$E.üs'p–ÒæÊ—eËxæ×(ã[T]þ¦j¶0~i'xeI30ç‹„5ãÆÒÀ9^ª*/*mûd?ñO]›ß+Ëí'à†åõLrÓ¤œÏMLê35ÿ Âš¢¬A:AzƒS	ãá4šƒœb7Êr¤†G£È¥íN(6KÔí¡ÓôAªŸ¬–ó8ŒÏÍEð]t`#:À}@õ&p~çÁ«,¤ÿMÃþ Ê^mlÞ›Ø}ÚNÝd?5b÷Ii:¹:Óp—nÉSfø™ð}@Â×ÀË&˜„ñï &`®=8ªèúç7€ëè“Ó¥.öŒGé›¶ÇÒâD¼ß
¢¥3ãû~–åÑè]ç”™Ñ«é’µQygªý¨–à}eqXx‚/žÿIy‰c€öëÿ‚íßÑîMÛû@3¢ÎNø&„z‹4]©e™6kE·¹}cc{‰‰ËüÃÃ@¿Y¿ÖI8!ƒ0ÍÃCõ¾ RkþþÁGó“ƒþ1ª2q ƒ4šæ	€î}f·H›jð¹‰o0ÎÔK Úžjö@@“i7%$Y;å´Õóº0‰œ:âÂÿâëÿ,Ÿd+Žœo6Ï—_O†@ƒÝ™ÖªGôáI˜Ÿ'ÃiŸ8Ø¹ó+{çßkõ1‹g@ö¯>‹ÔNåg¸5Ê½“ë’“dˆ*#éhG	%}x†øSÈÍS˜¦IØ¯ oƒ;²—ÍçªaÞ°ÿÌ|(ÌÚ—®}š#Á¡÷Æš_†{X<Ó.â=büÍâW¢šô¤@¯+ÁD¤ZmhMdìBlÐ<ˆüGBòoR¤²] pºžø"s†ÆaÇçÉÛ½øûpïÃ½·‘µø·iL¤ÜPM¨ˆæ;Š+(vîÿ||ðìÆ7Ñ—"rÉl6±lŠ'ì¸’­è·24™)æ–˜66¾¥7¯mùfastùl69¶è!®£›†T8n¯ý¯µ³U	í÷Zð¬~ ÊG äÙÄp‹?a)ÒèpÐQÉÊõ=!@·¥,{4ü˜ 8íPÁ°„ 7[DÝ–œ	î—É;s&1tT«B0éŸÍ‚å›ñÙ,Fž3Ì ÃFþzÙ'¼s’kkäq’h¶CJïÉà<ü@‚
÷`PØ÷¡Ù÷–@Ö¨k ™±Ý,Q¶oÚèeYR‹‚˜ööÅ´×.œ|æ“»Æ•­Ýã–| Wã«Uro}}ýC²^¥E9³³FóiÉ¦ºÛí~¢ÌØY
2÷¢ÜëdQöá,}Ã3FmçbÌÆÔ¥Úœ"ßÕý!|—¡ÆU€;’:ºã0>ËÏ©Bvl™bÜhüI'Õ°H€ö ùU ÏÂø˜Ál ô:Èñý@‘@HéûÐ8Â¤ƒ†a¨wL	(Ñ$Ð©€Œ1VHÏ´X{°JCÆÓ²Åþ‡qžF¡öÐ^Ã5í)[%Q<J^zr‚0B
u³é‡¬;Šâ¡Y´gÖ {¬Ìº%uµúAbá‘-¤6+ÆáÏ‰pxWÖŽææT9û'®lû`Ú–T‘ÇXkË^‚¸¨ŸmuËC~É^?þ¤»Ë¢8³¯aA0vl)µþ)Fv‘1±tÑ’º#ÒÖ½¶5é>äâ‚<­æQQ¨V´­Ö8'–>Å¤=Úkz0½M² ug^æð>:	ipÎ¶¶‡U~­3âèñB@³ñ¤€;ËˆyÒ.ùÄÎ«—£0›ê¼k³\Ãizîº„7MäP3‚vê¹‰J]F:‹ßyÞ¨pË¼×®Ì“Yä‘v–'~œo³-1“JM¿¶ÀõâUP/† œ„ÌÞ&2{û÷¶sOy1ÛÐ>Ÿ±ï4ht8Kÿ™+oýòÑ¿©úÌ_’)ÔÀJÃèõR<ýãÞ{›“ëŠ¯¬X&(²gqÏ\„r6oNÝ2Pµ¬÷Ã¬š¶óÈ®¨–f‡€åÅkÁ´tïJêõK?L¬KMgíÔè¨ÅKPô-Iá¢½ÔÐ·ª5gcY,êñ²x}ßjNÎ€9_Xz•þÄx-¶ôòu‰§7Y÷UêÝK‡¶º\ÛEK1‡ùPîRø/¼«1¾­ÞvDˆùEã$E_›²5ÿÌ›‹‡žâgþdÓºzíspž¼ˆ<É­nž<Gîi¸§ö
ˆ	ž ×ÚÙ=<ôÝîöÞÎ5DŸ(þÌO>[QæÊk+ŠGh½²4{Å—ËkÃ‹>{šGÕ*ûáº{nyèÌDñeÓ|òôŠÒ”ª.üD/—¦zz¹Ì¡³—‹[/o©\^‹'?ù¼Å[·Á¹7¤FskãåÒœÛRµô	û´úr1©é¹ú¬Yw÷JÃU²áVâë‡òTçËe>0]qdÕ(JµxjùåÒd<kúÕóäÁjA@t—1@ej.	œÞK¹•”¨ ~’\åcµ¯ÒnÄèÓ‹§b¦Y®ð¢™3gxYÓƒ§4?KÈEÃÇ¼˜Ô8L…q•ÕSÉŠÖ’Ge’ÃAÓ£ç¬®Fw©ðÂsißv66É9þÇÌ`~)éQ)çkVròAMpÓÙ8}wGá(ýùö["›÷½AÄWÑX42%ëvÞ.{I§¯?Qü}ÈÞÙi˜v—ÞDL9cü¤ª.ÇKÛºƒ–“éó ¦üõyÂ€Äh`˜ð`EÀ{§Ñ8Êi\«çYH¯îŠõ“ÃÝã¿§Qü@	usåmÚÈµ%ûìŒ4”€MÅ"ÊÍ…œìÂOñ?Íl¼k#{j~Ku¡' Ïò¤sNÇïH{¯ï#/4Ú–,è7“”s:˜¥)L¶Ï"m6áRÛlÌãûÝÊV7Êp¯éVóÎ¶¶æï-˜£#ÔŸi*±bíäŽ²Ö·Cµ¬­‘ƒiM"¸®_õn\aÌÛÓ4|Ã¢4át»Ý.6lÂ`J¯¥½†¢Mê¡;,B½»ß¬yÖæì©€|µ²Ò  –¤DÈ“³³qø!´…ËÒ¸¢aÃay ÐPþ>´%{}Æm¨Ú’1túâ?•øk‹´žQ¥êÁãÇÎ€¦jùHäqºö”è6æKç”>… s~h¦¹Zm˜‹:,D§%j¸³LÜ0og‘ÃMèn|+Ê¯@4"‡ˆŠÊ¢žw·¯UVU}%ªÄÕ<ßß½l{°€FÆŽ±k+ü™HXtˆG¡B¥àùä«ÐiòœVÝL¾‡,‚ÂyçZ>°­Bu‹L*ÒV–U—·u´OºWtóäo6¼wÏûAÕ_µEkû6yíxíg_Jì¼‰Qhïê‚’{ž‹·ØºÄ×nÇóº³[DpÆ5ŒiK›¯É/K³-W“üí­L¥D$ýmÃÌ@“aþ&oño«¨q¡…:aÙ/âúðˆ¨þ~ùºQnøs¦xÂò$LÃx…drýßã<šŽ“ŒÇÅËÔÈCLeHãCÏgIé¶(º½­^%á3™žv6ñôàßu®yÏi2Ÿ)‚—^Ç}ùB»ØK¡@ye9Ì€LJCùÀwßÁwê–~	ÿ{ä›ïh½wÊ»‰éFó®1î|&uûˆæ=îy@CÓ7è-uø¤H“+_ÿ×D™ë§fÞßí`Ú¤ûQf‡x½#i”ÇðåZñ-âq4O ˜ïÇ+ ±äM0’ÊÎ§×¿þ8‹†A9Hõ)Â6ÊKÝ¶Ô–’Ùª¤VïÐ%133Ke½˜È^ÝË™Ñð¿óÜv'Ái›"È´;#ÄåÀ’ÃW£¦e—C„‹Ù”íï5jh%…¡&SDä-òdP]#UÖ³uZÈ¥Ê+Ë¾X|P‘ÃÆd0[rá*têÆg½ZZmSÿ†Ã1zõÐø¡R¦†V£Ðecsl¬©>kãœ\€_ò=&Õ	ÑU¼ª©¢Ã´šjFÃâRé~Ö£æ¯IÔ«k°†qòs[&þÈXÃoÎUä¿„ISÚ`™r•Z˜çnLMf£/†’ÍNŸâ²ÑSâgxØ¥µäéõ¯ƒÙ8(çƒ¿˜êïÅC¨P¼ ÷'Ñ™¾o§É†ËÔÒ²ï
ýœÓ«×k¿Ë­µlã‚ÛP#ëŽk#ú¹¯O3ð½?/kUº£•°À:DÓ„Rlü©3°Y«ËÆœñ­uûÄæ¸*7#Ùt©ZÞ""î>oSÌ’…ïžt±FbÙlõ2{Ð‰u˜@Û¸/™ki›ž
p[ÃÇC+›?t‚üÏÞsMÞ½ä_½º{	=ØrMòåggÐP³0»ïi}ÅÅ<æŠTÈ}$­1Ä ¬a¶[°_ãz=Ïªp‹>ý›yvÚšÄÉ$ì±[d«e‹RŒ±Ø#w {sÕ+c€Bý›  Ž=²Ÿ•«N¦ÈJe/8\¼Ô¡ö³O!b¢@;níìÊKd"n¨<©ž&åºÐ/ÔêAKy$&%pØw'¥®µ_ŠŠâ´ðC[£¦l
³²È$g÷G÷±þÀîYÕšÃ¶u4lS\€õéí]wš´•¢ˆ]‘ÇOÈÂˆÓÙº´y&•Tãü~MßrZqcUâ2Jž®ý[Ùš˜ÔàPRžßwÌZk‚×P}}YÜ\+âŠ+‹ÚÙ
FÙ4‰¯}ŽÛ‘#ìŠÃžÌqH.SRO¨f÷Æp`–z“¡Ð-9|Dá™3de	·,J™Í‡ùb:N¡®_˜¸Ùãª÷ƒjà{OOw<%;Ø·¥r„pT…z£ûÀ'nEEÀC…OÂa4›4–[j½UÝ‰MIá¸¹`à„ìiTD)/&ËøR²x.3>'”éaœUù%¨oA¬]]¸v0÷LÂ2ß¾æÙ’(OJþ@6ÜÏ…~{ªOQSö!RÕX*É{à4^õ3œœÓ`²aL®d òV•+½=­¯³—$_çjÙ¥úïßõOŽû‡‡¯öwûOvý:ñô‹öµa2 G¾ËàÒ€ÿ—“^Ó˜Q´WH1ƒ|ßi8g–yQ¼ŒI¼LH–EÎ°4²’ÈxÏ3ô4)äÓsãÙ¦ÒÊžÍMªy$8eèÂêŒ+2¹ù¦Qç³p‘ù&ñ‚$Ý¿&&kù0­Ry	f|\vÔû‡Ž3‡ð»*6DºUF$:Ï2áB³ît¡ñ1B±P +lç->Rk÷Ö-reù,¯{+(Ë¥,îÛ#çTçâ©õíô<¾Mö5÷ãòÈš•†iÂûŠ@ˆ9jÈdìe	Å ‘ W÷Í©j*}Cõ4ÒS‘† 
®Í%±Tjxó¼SrR>æGU\Žç=X”e„Æiby?w8œ†pš…ÀY§ÆŠ;þV£ø7—œ¿gæm&ëi ;G/;YO[¹T³ÊlúHÑõ<[>Ör¹ãxü¨—-É€¶Ð–FMƒYš%)Hæ92 ÉÛÐQ«^†_5Ížãµr¯À=J$°ëÿBéŸbÉ }?ÃRà^¢ÓüÕË²‘ÐJ{õöø[ÀÛ¯AÔ—JÌ—ÖîÅ`<‹Ry;·Z~¯þž!Cßág<,o<ìímà%.6õ ñ™ý!.@šŸËê…9íÂú¦/üù	[û6ò÷ÛªfOÃI‚Á¿»¾¡Wn€Xµ&yç—îX€$¥ ]ƒ}£Ø°yk]Eá‚ž¸ü>/ës%¢†‡¬Gy:ø™£Ûë#bþä¾;Œò$uŒK¼ê5wì$;Û6mÅpŒÁ[Ô¦T-æjMš‡—¯qIú]æ´ò?o·28‰ŸSæ¹ Jqò&ÁÙ™;*Im”wèbUþ"„/ëƒ¶e1›–vƒ _Òµ°Æwo3:¶*æM‰„Y£Ú_í1àOþäáŠü[2Ñ¥züÛ/®öö+ô"9W$g¡çÌ)û6â¸Þ;œÎËkÝ|Ï–óf­•ËRÁª‰Òœ3KEêÚ÷«ªg9kß«‚ÞÊ½8dFWDp»”ÍðõëUÍ©Ð»4»«¤/Ñ*éËš‹k©dp™"¹•Éœß:þq¤ÂÐûÞ¦ÇS€ÛèH›vE¶g"–9'ª15À*“0ž±ü)"DÔ$¼þ';uô"76yƒ1ÈÈ4½þåMdHÂR.z #—ligö®¬ß°«Ê 6/Jm2ÐÕ%3›2­³-Á1ÆyR	ius>˜<Ù}h;a®u
Œ4òÉ¬ús.èyRnO ^ÙøªÓã\îŽ{ =D'Œ¦‘,9Miz´Y@Â	Hüù4Eø*||9M€fedX…ÿ½~·2É©
U`Ñì¯é¥Kg€h}lRT³•4#“SmÜ®m¨O÷a]¶üêˆú§£o°à¶:BPÔ˜|·ƒg®NÐ93²‡¯ÿ™d«L‘ÿN“ñõ¯9Í±[ˆVåù!2¤“IwÎ;/î±@~µ9©	Êœ>y€°ŸñÝ6={ygc(ž–D‡Õ×%1š´µüR¥–€ˆŒµI«fù”1ãY˜[8qä¼¡6åFà¹Œ¢™Ÿª¬ 'í’™‹œ‡I52¥—@=á$˜‹©š•+xÀ:.ts¡Þ|§#nÑk•%™òNÏlA¤Ÿ}òV¿ª1ONSæû<Š«óí Žé‘¿ÎÜb€¶»‚u‡$9ý>¤7G˜‘Ñ‰U8Ôñy’"â
Ñ«ªÖ_…NÔ(ƒ†ÏÑiðlak½2l}Iük§hæ›Í&üÍBÔIÒ1ï—zÖaªs¸Æ‡Ã‘Án¤ŒED¼Í¦QLjÌGë[jB«÷·5‰.µ1*äõbÜúöpç±±Û* ë}tU/£ï€‚Ë˜ÝÂ|j£Ö]|œ 8«ˆdOR–#'5›Ž“`ˆþ– D)Ðz¸<0ÆHAŸ€“‹›T¿8ÀƒÂÁÁuÄE„pƒ8¼¸þ¹3¦Hh q(‚p],ŸÄ_Š§IDÏày<Ä—µ÷p¿Üá<¸”D·ø<Ñ*u$d1@¹°yþÀhœÃŒqFÑX›<*Âiþ°ÕGºß¥MâqZ4•J;øg8aÏé±ëÄ¯a”§ãpø°ºòzå:N×ñ›¡Kù¢&ƒ™…®qröYÚ‚²È8™ ?‘’§°%Ôun0üÅ•½_ŠÜUEt}N0Õë$H,µJ (# …ý™Äô5pRcœüQé2Oh¾ìŒÌ²ÙõÏi”d‹c>‰]¿r22›.ê}Ê–-ÔòœorÂŒZÅ×G—­þ‹>S¨AŽh,{dçáP¹0°Wó±#¤‚®+†¾wõŒm”±…²Á%xS4)’Ü#ï1ÿ Ò§Òô4Ü+ÚÕb‰HÝÝ÷0¢Lâ48;€öÂÌÀÄ‡Ô<6k‡	ziëÁ=ŸÆüôeNÏ~›&“L½L'—<!ã"M£Ò-fx9¡!Ãâi‚ÜuƒyÊ¸XÇ¬xg2*#e*n8­bmõ]ë¶É°wJ( ýD@ž!×ÓÉ}wâŸÞíÛ”JupJó[+[]Ìpýƒ÷áiÃƒ™{íºÅ?¬ÍI/üsfƒ"÷µßë+yqmw+Ãõ!.Xñ–àÅô$£…íë€5_ÇZ9÷çÃ°\VªÅ„ßCC¦À"„a±½!}Ã‰š35‰öEÀ¨ò÷¿
à©ÑD¿º]fú#qÿœ²˜uÝé§iQÙ(KÆðoRˆY¨YÕZÐ›Á}t&†õ§HÙ4*ØÅÁ›SÅ	ÉÂ¾V Yg_dxãñÊ "”8ü‘+‰ÃüúlOPƒVñ&“õÈá³'«äÏ‡OpþÇ{BÚûÁE—<Ø¤³üô\¨·¿ºD,y¤0
ßêfk\§^?e·¦md?kµ˜õÏ‹ØÚayÓ03QZ†(Ú…1o‹¹¢ß^ñ%!×¿b¢ •hËb7m²Èµ!è,Ì÷ÔÛ×Fý«åÉÜe`å£ÆfeAe6+.*áºÜþ¢{	/5+V‹Øš§%°·ãÙxl­_(·ySX¡"î ÙQ¢ÝÄ¢¡i…W€þBÝ- YÀ²·²›°ûÖ•ásãxl‡I”³¨Í–ŽgKtm²¬fS+æ¼åw_êL|U
ý–úÊïy›5ûßŠœ¨#¡zZqÄ{Ÿª
eHí‹Ù`WKôVìÃ‹YßØ¸oý¥C€®…Òh¢¨ºX~ª”UœEýª®ŠY›1§×k@p
ŒÈ]gŒ6šFgç9‹:z¯H(žR×6ª§[ü²)s®÷—õâëÂg©½+X|ù™AC9÷öŸ?í_ÿÇõ¿]òôàÉÞ3=íö\cÃ½ëP-Á‹n^0îœá¿ ýíÙï_¯6¦«4Qçøyå%y2Îè|;È·ö6¾œ^¼Âÿ¼$ ¢üÆGâÐÛ¡tÆN ¹[Ç'ïXld›)”)Š‘Â§˜ØúÚ€ *PH4‡¹Š&gfÊãšmÑ‚1ˆbvˆæÌæúòô‚‹}¡bZWŠ§$5>¾ë|ãƒB€êÄ@—kt'Cí=úã:±Eá™ÞÐàÞ@óð/JŽ€
„éaÜÙ»‡­8éˆ¯LBŒÙDÉ("YE°<éä)¥É¤“ýðŽ©l“ºèè#9Ö#öŽÏlâ$õ— ¶®òó©[Žkß®§OP%a*rY0¬’#ò2†EÏrŒ%z@]Æ¦©N¦k]û:m'…“ÎäÎQNùè;Âz)›Zl1f{e½†’ŒS((eS‡›
Ôµ©—ë€H¯â9Ð²B'úÕúº++& õU‰,u ßM*†ýºöê°ÖÝ{€§_3ýªÃ×X°_irÈhÊÆ®/þ`÷ød—|÷üÙÉîù¿ÍõÍ/ÝÔs™tQüCÖä…Ÿmš½Ò1ú°éù*Gñ`•íRð¢ç5~žÃŒ¬¹þßyeuŒr‰†µÈòC-ZÁ¤+ðÐBo³[vnã Þ³ð-­#åA%³tÌþ'ƒ€þ©uKÔú7ÄûªÀP†s×Hµ·÷Ê{Ç  ÿ›X2]m«¯œ:¨Ó,HI|<;D¹x ï‡¸WW5H*8ùlRýdAûŽjF :½bb‰¡@ªOÃg‡Ê×<GÌŽ‘:U¥P=ön·Ë¯rPíkÖÒÖØ}7f§©nÇZÓ}~ôt¾…»0ÿ~Ò‹ô)nçSD¤}0¥Æ-ã•êÞV6‘â½‰¬ÿfrdôn§‘³ç¹Œ"˜êöBj©FMõ%YM™h0­Ï:Š½ÔyUQ¯BÀ¶È7ÇÁ›Pq›±¦7«ªI©#çoÖûÚÈwÉÊ2ùðÅ}¦)Ô\Šž¦5=ü%åExÀtmM­™u k= ¥#ŸÂ±ÉöŸ€¤qž¡ÎÔM­¾Xzž^›ûÎbœbz#kÉÂø0V0F!1f\¯p²Ÿ”{Ë,„•ƒ›µ™&í›£Åpšœ>BpàÇÜaW¸Ô#)W˜=ÿË’ýM©Â5š5ÿ¶‘ §xc¤šÂËèßßÂ­¥63”jV UólÊ¢›;Ð„Amû¢ä¦h¤R^â»_5îêéÞ³¿Ãöc	ã–/7ˆ½±C9se²cŽòM´´b“ÈÃ£YÌ¬‚£¤&÷‡“(~,Uø[¾•Få÷ð’ÿ!ÏhãSþó8Ã0{x©~.k®¹§8HÃlQô~<ÆÕinW~§Ó–¦’ÎÃ	Î! þð]eÇì‹+e)È2àJðß¾RQ­·Ck²¿Lu=†¼Ål¬ÝúÇåO¸éd–¡ÉÞ4MÐèÎU}ë}ÆÃç ß^Êò^¨êÃCšCçá¥® }ÍO«
Gtê$$ÇáÙîæõ?üU†Þ‹Y}"ÔHÃ˜¥Rv2øóñÁ3æ!€jÒÂ˜~ õ+œ:QæÑ©t”“¤öÇÑYÊ’Fu•9Om|¥]˜¯éƒ´H½@çLgÀöÐ-£¨ëŠÓœü¦àyè!ÖSU$»¡œI9(w}ábÿ¿ÙIÞÆ²ù–.TËg¤œtÜ•3óýÕoV±C’Pä½+:ÑÈWjkT­ØÌÛp„ˆ3ÜŠ„fßð 3J5 á€ÖÖH§Ó¡ÞÀi2ñÁá08c^E¡ø~ÿÕe`iEÍ²VB³ÞŒ²#ÀÃ¡ý`Çþ 	³Ù–¾’ÔQ&üôþ9dx“}Ï1f´'=’å(Q¯’|«G@ ‘H‹üTÄÄ@èM±!Ç¾[JŸPuštáÅ4Is‚$•bTý2Úš‡äM–F']|âS]-_ýÝUÏ¶=Ìb…Ùª¾ÈÂ œŸ„édåãâãKF:ÎEj·£ÅÛp'¨Ç4)m¶+#·ü¦…¬nÖ“£ô£Ã?¶¾e´…þ¸ú;n–"fíf™Ôù±ô…ÒwÙö{›…°Ï	³,ü³C|Û¦-äÓðŽBñHúBÝÞ´h{æ;Q€Ò†©§0Ý´„˜“ôYâ•i`_Q…;¬ªpx&Œ³å9« ÷®ÀN¢hõÒ:Y_ïÑÿ	N˜Õåú¬cKía4¡¿14ÚòónpšµC4Ç/Áï¬üTk¹¼ËDËAÛEwk¤}üž|YügóþÊ
æ`½ðµF€/ŸµEg+d‹l^Ñ9V½’ö|¼Cˆ>JÞfìlöAhË‚«9çÔHëÅKz_øÔÕ{Õk7œ·Û½’Š@C‰}îƒ(àM’<a‡VüÄ¾³Â…0fÕyÝZ¹œÆô9HBªTØ±?c€²²1õVB“Wæ¯LÆœÌÂx°Ô|h|XôPY­vXôµÕ}£¾,g\4ä{N«†éÁé÷lÏS‘1mF§:ëÎxF4:û?UšÓÐ^‰Gk>C^o¸R›É0Ìƒhœ‘¢.eµKÂc’ÅSŽ¼z#ÉÆrç	f=r¿üN
½	|<a¨éð,'¥fnºüWp5Sóˆà'cýI9&iäÈûgÔŽ8 B#ý¢´®nQ‘JôÔ.v+öbL8°J¤¯G“Ú¡½	ÆIZ@{±ßê×°ŽõjCºý|“ÆIðëÒ¯±Á}uWi‡0M¿!p‡Ô`“î¯È½|)U§!Ü´m¾TÚ|íÕæ+¥ 7ŸFëj«ëµ õßI’cÜ/~|¿/ZÕë'ãq \6^JyÄ½IØ×C»½¡’poØ·òå›©WÕ/"Ñç£X‡ ®ÄhøæCÄHÀ8We"r·[]ÊÆQ~¼ó7ÒÞN&a:ˆðõg¥T@8è¢Š¬Ö)þ`è¼;eçJ¶ô½aO@§N`®ç˜W¯áªx~úIƒDWÈÜ÷d”¸ÕEÿwÚèÚêðìs<mûÃó$®v0ÅïhÇá„Â1JöúÖÑE¥íàN¾.šãßŽ9<£¼Ò‰ªÎ[r¢
™öóöúJ7Ož£{;ÈUøCE’Ð6Î@âÛª‘}}R¢xMƒ±2c~4ˆZ¾iòÏÚ¾aLÖ<TÍ’Tð-NM©£½JøÂhœ¥7ª,q‹[t¿O¢¸áóª=M#l²â1½êð–I•;Xak¯MŠš^Ô¦PRo×­Ö×¬]o˜±\“ƒap
Ï8?¸ëŸáäø@Cy”èŠ-­F-I^Ô©`¥"Â^…0ÊUF–¿+‘w­æA
×°VúŽR5hUî¡W{‰)ð¢øKüËi†#)Ã¼J^TÝU&/O,%Š94V™/3ÊPòÆ@?,ƒ9u³JÐ¥1§–‡Œ0E-“’ì<8S\9óoóó­p!•ù’ûMküwÚ•‘Ö
öÙ†Ãê÷ÚA·µw]´5mC³±3ÒÈåvP-[*þU¬–ÑÄph8øhÊ‚xÊì4y«¬ˆ§£F'&lK¼$ÔíÊDëÓä­@®]ÀŽãÙ0ÌÚ¥Ð®ÖPnv&îÆÂI(µq_÷"-+¤µ+J†7¨d¿Åœ„zt‚<ãÏJƒ?iv—ÞiY±` rYm©fväS”´¦qdÈ¥uäFÅ†q°’å‹/ÔŸùjÅG6)ùJTÀõ¥æ§ªÑQT0òÔ^ÊÊ(;âI‘øê…ö«ÄýGOw»€¿_mìïmïõŸ
Ð2U>î?í¿ÚïŸím?Ú—sÿú@Œi=˜á³Ðüb"“3]N”1rDÏÔÇ*q®½ø^è‡t¿qéKßŒê˜øOô7UH(­‚¤Éüá!‡’„üIF|€”Z²÷Ye-Æ®¯JYÅ8B!\Ö(6ÄÕmb|·Ìí
h•qj}ÓV§tË×«|+&Y©Êf°ZÜáY<°rfŒÃc5®Äµ‘k(€Eí;d·JàÂÆi¥pÂ5‚ôK™DþEÏO™I§%\ÉÞ¤lzlp”Òª/Ê(´qÐUøÄ½Ry#©Œªš,!µ¦©²p‡„ªb®Ô>xàÃ¢jºÝ®2@É½öï¥`”ÄÏEÿÕˆì:ÛƒíþÓÝg;ý£½ƒWýíƒÝãÕrV¥©WÔSç_Ì¡ØÃÂn/pˆîÇâ¦•ñU÷ÓÕò™Ap5jÞ®jçEjMÞaÊg¬WãNV¬²'À]úB±{1 ~Û
a4&JÂZ@eå*Q›µ´š1uWHÜ© æå¹¶DNA±Á-¬&3DeÝ“pŽ¨]Ö ruYåpïÊ¯ÑEùÛN„ášPˆL»5³õÕTÁæõ#3y}÷2e:¬+B+”Þ¸ÿ  ÿÿ 
û-˜xœì]o7ò=¿‚YÔz%[‰›Öµ]¤ùhƒ»6†í+zŠ„Þ¥%6«å–äZRý˜âŠ>äéÐ_à?vCr?¸_Ú•lr¸³Kq†Ã™ápf8‹ÎX@=*q(Éâ‘¿îÜCIsNÉ˜^ÿ‹!Ÿ §2Æ×¿Ã‹£†GœM™$ü!•Öøs™ÁùH	]O=Ì¬QÏ±ÄjÔ«ðúƒGÍ0ú^…@BeàKâ ó’N­?à€qôœbÔ;}Ð×ƒ®TôX£žÅB2tÎ$òqžêÔ}|¯ùõïØg%œ¯¹Öà3‰e,ÔBNðOI(ÍR„îÏúRˆU¿ÿå½äÙc¡hÆø{1!D¢#ôãßÏ~Ä’bð³`á[ÉÞêŸzjÑçìÅ<b\‚2ücï‹àªçmHf½ltùGE$ôüÙzRl;…=Èæ%%g),[8éˆâœ3N%yI²n–w)Ê·9¾·–°ò'½þ@²Wg¯Ï$§áÞDPÙsÎþ›ÝŸVƒy æï*Ó³ðœaÃNI€åõ¿9eˆh*ÁF˜côbî‘ 89E"öˆì¾³ƒxV/N‚i…<,½	êÎûhYÆþ‚s†p†š#žM7PØ Šñü]iàDÆ<D†+‡>½B^€…øVäˆ{Ä]¸Ÿ9ÇÉtËá'è[‚}ÂÑ'ÃUÒY»È©?®Ç4õô3g3õb˜
×Þxÿ”Ÿ^.Ü"g„„hŒ#÷q6ŸAž¿ÁûddÏ%É\º£y€.Atî|$9öÞƒ€\IÇ‰ô¼ î»»Ž¡z²§NF…é£ÊìbjÏ±_™ã"¤²\ô—˜úÆv¡K0&¡G(Èß§×¿qª‘±fŒ½›ÕÐ?I‘žÈâÍ°ÀœÃ‹XJZ£Yøléû£å‡~@ÌÖÕÚ¶²F•EWQ*ŸäÕ¨ÄÜ¡‹±K¦„ãÀw?ÛÝEvEøÕù:5_f@hÉ'š»û(Z¸gqè_ËOL`;Ì439•”….”òXàÒƒÆ¥Rý†c°sÙB
bzÎfaÀ°ý•-÷>_¡aa ,$<~‘nkojöõï§oÓž„óÖ®øîú0hˆþvò
=ÃÜk6É˜S©?j“w‰éAþ:BÁØz}\»-ltð˜ox}Á¸Ú¡æ¿D/÷@)¿øîò‘SaPE×çÂÞh¹Ö?ì1Ør8˜ÉwàŒj:æ½/m
uœØz½ÈøúÃ¡E5/ò¿ÁÈM½€9’\9K“Õ¯åQj5à…ãòŠ>¯ìcó”(‹Œ”«âŠ×ð°È9ç¶zÍÂK›¼ y‘"àéÌ±¿©"_è¼ e``ó<Ã3ô‘hƒYÖg5Úpú W©V‘Ro/éœø½QÀI =GÂ;N¿MkþWÄšù-~Jp@UÐP¼‹´­­]ÞÙþ/ž¹7-âÎáÊ¾–Ž}À«ynÜaTá#‘¼YÞ“.r×¹™×)€r- lÑ'e³3±‘B(ª*EÁ¹¨å®U	åèÉ7)F¥éj”ÇM8Ä[Œ(¹]mNx‚ë¼¤*Ñ+RÖ•3‚¹7)IÞ‚Ã‚1p8 — Dƒ}$YäîGÈÕþªæÄBw”÷ÎÌ-ôKäÐ0Še	R.¢DeœZû& Âœ¯cáGŽ-:I‚‰$‚XŒ¬dË`0(cºÂAä	ÍsÂ§«Ò ˆ( –Ã˜é££cpËåY6ºGàM‰hDý2´ÅÐ™{ƒw Ÿõ€ð $´qKMM­:,užçáÁ%óbqmóªâtð¥­ƒã‘ŠMŠaF2GÈÀ¢°X*Ë¢_œf9•Mu5XU[‰Ä+‹2a3™Fèaµ#ïÂé6@‰Ù*]x>¿·Ó³A)­3*©ƒ¹
n;ïJ4È‹¹`ÜU±iQRe+~È"…0a½xãsæ+¦Abø*Pï›Ó½}pINOžô‡¤“‚pŽH½v„„I,@x«‡£¨uèVUÏd»kÞ™5þ/Å»ÅƒþÕ$n;*O19ÇÇACuÕ>À>Î¤=÷·€!éú/j2'c˜¨»&ŸZã·8¯þRá’íÔ3š|ev3Ó¤BË8¤¿ÄÄÈ@¦8êõŒ µxz¥áùÔïÉ"õª(ùUy=Oò[_aýªÏÞ¤ªµ¯U¯ý_€¯YL,ß–Ÿ»ï*ëzÀó„ú°û‘Æa|úñ),Gwy©•žø§l&	Çr‚ŽŽŽÐ.úªÀújÌ¹—¸µ…,ðÂ}Ô)äÊS„ú5×æúž„“xŠx’Ôéð”uè—LË%U_ƒ-o%:`+š°eŠçîÌúh:7)å2aÏ—07(.ì[E<)ëm%ëÓ?øÞVÆ=T»ÃäôP!™¥ÚI€Cëv k¤ê(ÂœxÔ,Û-ªSý^Ri[Ïy– ‘Øé·q§¤Å}t°Vö™‚ÕòæPâ‹€ÔØÉD}!.p¤B-yˆRÞW˜|('û¥-‘Ú×á“ÔÄºÕ­¡Q¾ÙÛæ?mžž®¢HáuF(,ìŠäãqæÄåd3À4BÛòÖJ6‡„#ñ
Œo
j› ç8ÍâÜË·ŒoŽÂ9>ÈõoW$@½—ÏèoƒCWlø49â6…S5[²‹««Lç8+2¸–‚ÛR»[ê’Ÿ¦Yåu“ß|iI½ÃMt
Miæüx_`’øÝÙÉÜy5Ð[±j$˜®ªA»`þÂ¦L+Xw’‡’Ë¸oÄ2%>§5(žãÆ‡b³Dý¹ö¢–5«1e>:By5Ã—ãàƒÂP ®ñK0çÿ„÷uã§@/¸(©ŒÐpß©¾^}Šöúƒûgêê¶7ÚAÎ®³•%D†ØMpHðM€½{°T‹Y¹–šHõ ?®Þ5CSñÎ<Nã£ZÝ^§õ?è0ŸëáÃÆÑ¥2ísåÜ—÷j uåçSm?2n2¬YÍŠ=åè¾òÕ’Õ{V(EuhËê]ÕAN8‘)íÒë‰PÍÉný³ƒz¿‚ÀùÏx^“Sm9³>DêÄçLh<£ýF¨J˜fZýN7Îxr@[þw¹J¿Þ´4]g€•“Õøkgvª¦fVÁ%Ì½—ùüÚ¥²½œgf£USZäêÜÕZ¡Kó(°lUòþ¤šà.·o®Q3póª ×QÜ hÕše¡‚¿|ÍêâûWé×mL=hTê¹m­Ru€º¢ƒÒºÒ]¬p6NÜF|ê<nM<ØÒj˜Ô@éÉfÚšÔÌ[]Ok#)ë×€L”o
xz …$|³7ÒD#fIe ²öRé¼‰Æ®ÌjÆÐÊªì^u©ÚT‹§‰ö)VómM°vþoH­s¬I™(T«Ië”Öõ×'“YÊ«ÖÓÂåmf[à”€F
h×ÎKÖx–é‰2õ‹VhÍ\%UuÖwfU“|Èmoõ¦lkeËïwÞòi1ù–û½¼}³«`òFöÛv(Ö©Zs	F©&to°¿Vê0¹{kAênjø‘ºä¦l?­ËvUšî¥P¥Û?3öœa¹8£2ß°uBERE8|ª™üä›/”$Ö»pªe>sëHí½™(Û¸ÕVÄßæ·©v€J™ÓVêZ¸Û&I,¯ÿÈÅ	ÖÝÊ²uÀ²Ÿ¬d[ Z9%‰‡tªxÅ8õbáï{·Õ·6öBI~ýù‚gdU—›â´õ8Úrø4 \>£ÜHZ»²¿*×«T Öžª­‘ÜZ¡´ÚÅü;œmý
“%ËÍdwHÕˆ)­M?ùÙ :lƒµÙ_ÝêêòZ8;9XLÒJ›\È?hº>¿º%Nt•®ùøênVSpc÷ø{Ó·ŸòÐî¯Ô‡Ÿï¢Hw² M×£3÷’Ê_bÙv¿É³r;(¯ÐÏi±ÃOêÒ[g
×"ë›Ä¬)Y»¥™Ç3Ž•§mJ
Y—º´[@$ºá3%ƒ†¦Ý²Ø¡6ßôEñR­x²êÓ˜v£—¨—¯°ßÊlÑ©ª\ ýŽK¶üOÛ×¯yPúšÉúp)»kwÚøŽQ‡·Mu¦NwEtþÕÅH¾#Ú2çf+ÒV†›íZ²iåR “Á.EÉ¥´öhWåµÛ)î¶@“ý×a°PåýÍ6„µTä †“Ê…:ª“jžéT›ÿæÔ®½‚°[õ»Ã¦¦¯'ÄšôAÞ²O{Ê€vÒ¹ûÞv‰#ˆ3È	^(ûjNßöõšf_±´Çl9ŒuÓLÈn#»™.¼÷©PeþÑ2çh8+êM5µ¬‹l—ÕS C úU‡g·Á¨bµ;Êç«ÍA,÷¢GåÏO›ZíhÞVkãåÛ†7É‘¬Ë)¡ü AyÈ¬’w-e¹ã’G7¹÷Rt#5ÎÑMÃp*ãƒ4‰	>Ä×|¼cÕ±aB”%êrÿö£óÑªð	º„¶=É_?^	]‚ÑÂwÐ˜›Eù¿ÔÖfÕh~UÛX•d”‹’uÛº¹¬f3ïÏaâÕ½ÿ   ÿÿ `¤Ä˜