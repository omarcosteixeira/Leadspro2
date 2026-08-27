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
  const [serverDropdownOpen, setServerDropdownOpen] = useState(false);
  const activeServidor = (localStorage.getItem("servidor_selected") as "principal" | "comercial" | "regional") || "principal";

  const switchServer = (newServer: "principal" | "comercial" | "regional") => {
    if (localStorage.getItem("servidor_selected") !== newServer) {
      localStorage.setItem("servidor_selected", newServer);
      window.location.reload();
    }
  };
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
          <div className="p-6 pb-4 flex items-center space-x-3">
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

          {/* Seletor de Servidor Ativo na Barra Lateral */}
          <div className="px-4 pb-3">
            <div className="flex items-center justify-between text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
              <span className="flex items-center gap-1.5">
                <Globe size={12} className="text-sky-400" />
                Servidor Conectado
              </span>
              <span className="flex items-center gap-1 text-[9px] text-emerald-400 font-extrabold bg-emerald-500/10 px-1.5 py-0.5 rounded">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                Online
              </span>
            </div>
            <div className="grid grid-cols-3 bg-[#032554] p-1 rounded-xl border border-[#0b3c7c] gap-1">
              <button
                type="button"
                onClick={() => switchServer("principal")}
                className={`py-2 px-1 text-[11px] font-extrabold rounded-lg transition-all text-center cursor-pointer ${
                  activeServidor === "principal"
                    ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-blue-500/30"
                    : "text-slate-400 hover:text-white hover:bg-[#082a5c]"
                }`}
                title="Servidor Principal (SM) - gestaopro-761e1"
              >
                Principal
              </button>
              <button
                type="button"
                onClick={() => switchServer("comercial")}
                className={`py-2 px-1 text-[11px] font-extrabold rounded-lg transition-all text-center cursor-pointer ${
                  activeServidor === "comercial"
                    ? "bg-gradient-to-r from-emerald-500 to-teal-600 text-white shadow shadow-emerald-500/30"
                    : "text-slate-400 hover:text-white hover:bg-[#082a5c]"
                }`}
                title="Servidor Comercial - gestaodeleadspro-d4230"
              >
                Comercial
              </button>
              <button
                type="button"
                onClick={() => switchServer("regional")}
                className={`py-2 px-1 text-[11px] font-extrabold rounded-lg transition-all text-center cursor-pointer ${
                  activeServidor === "regional"
                    ? "bg-gradient-to-r from-indigo-500 to-purple-600 text-white shadow shadow-indigo-500/30"
                    : "text-slate-400 hover:text-white hover:bg-[#082a5c]"
                }`}
                title="Servidor Regional - gen-lang-client-0111023338"
              >
                Regional
              </button>
            </div>
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
            {/* Interactive Server Switcher Dropdown in Header */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setServerDropdownOpen(!serverDropdownOpen)}
                className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-xl text-[11px] font-extrabold shadow-sm uppercase tracking-wider transition-all cursor-pointer border ${
                  activeServidor === "regional"
                    ? "bg-gradient-to-r from-indigo-600 to-purple-600 text-white border-indigo-400/30 shadow-indigo-500/20"
                    : activeServidor === "comercial"
                      ? "bg-gradient-to-r from-emerald-600 to-teal-600 text-white border-emerald-400/30 shadow-emerald-500/20"
                      : "bg-gradient-to-r from-blue-600 to-sky-600 text-white border-sky-400/30 shadow-blue-500/20"
                }`}
              >
                <Globe size={13} className="shrink-0" />
                <span>
                  Servidor:{" "}
                  {activeServidor === "regional"
                    ? "Regional"
                    : activeServidor === "comercial"
                      ? "Comercial"
                      : "Principal (SM)"}
                </span>
                <ChevronDown size={13} className="shrink-0 ml-0.5" />
              </button>

              {serverDropdownOpen && (
                <>
                  <div
                    className="fixed inset-0 z-40"
                    onClick={() => setServerDropdownOpen(false)}
                  />
                  <div className="absolute left-0 mt-2 w-64 bg-[#011a3c] border border-[#092e5c] rounded-2xl shadow-2xl z-50 p-2 space-y-1">
                    <div className="px-3 py-2 text-[10px] font-bold text-slate-400 uppercase tracking-wider border-b border-[#092e5c]">
                      Alternar Servidor Ativo
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        setServerDropdownOpen(false);
                        switchServer("principal");
                      }}
                      className={`w-full flex items-center justify-between p-2.5 rounded-xl text-left transition-all cursor-pointer ${
                        activeServidor === "principal"
                          ? "bg-blue-600 text-white font-bold"
                          : "text-slate-300 hover:bg-[#082a5c] hover:text-white"
                      }`}
                    >
                      <div>
                        <div className="text-xs font-bold flex items-center gap-1.5">
                          <Building2 size={14} />
                          Principal (SM)
                        </div>
                        <div className="text-[10px] text-slate-300/80">Sala de MatrÃ­cula</div>
                      </div>
                      {activeServidor === "principal" && <Check size={14} />}
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setServerDropdownOpen(false);
                        switchServer("comercial");
                      }}
                      className={`w-full flex items-center justify-between p-2.5 rounded-xl text-left transition-all cursor-pointer ${
                        activeServidor === "comercial"
                          ? "bg-emerald-600 text-white font-bold"
                          : "text-slate-300 hover:bg-[#082a5c] hover:text-white"
                      }`}
                    >
                      <div>
                        <div className="text-xs font-bold flex items-center gap-1.5">
                          <Briefcase size={14} />
                          Comercial
                        </div>
                        <div className="text-[10px] text-slate-300/80">GestÃ£o Leads e Vendas</div>
                      </div>
                      {activeServidor === "comercial" && <Check size={14} />}
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setServerDropdownOpen(false);
                        switchServer("regional");
                      }}
                      className={`w-full flex items-center justify-between p-2.5 rounded-xl text-left transition-all cursor-pointer ${
                        activeServidor === "regional"
                          ? "bg-indigo-600 text-white font-bold"
                          : "text-slate-300 hover:bg-[#082a5c] hover:text-white"
                      }`}
                    >
                      <div>
                        <div className="text-xs font-bold flex items-center gap-1.5">
                          <Globe size={14} />
                          Regional
                        </div>
                        <div className="text-[10px] text-slate-300/80">CoordenaÃ§Ã£o Regional</div>
                      </div>
                      {activeServidor === "regional" && <Check size={14} />}
                    </button>
                  </div>
                </>
              )}
            </div>
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
                            <span className="text-[10px] text-blue-500 font-bold px-2 py-0.5 bg-blue-50 roundxœì½ËrI’ xçWb³«"²€À#]Dƒ ˜….’À`Öì²(¤G¸ð¤‡{”»‰†È\gsØ9¯ÌÖÎŠ´T‹”Èˆ´Ì¥Í?Ù/˜OXU{?Ý= 	V¥K&nnO555U55Õ$^:žeYïÑÒø\%y]^«d’Tu™\7fßX®¦QÞTå ©ÞÖ4)Ó".È/~AúÍÍakdœEUõ"š$›½:¹¨—^¯®L/Þú{:+§Y²ôÍÊ
9.òziTd1™^,­‘éåÒÊð2:QyHYÌò8™0¼³.Ëqz*Ûðqc¹Ž_êX‡Àã[„Ç×fÕ„ÿÍ¢:Yúve%
‚¼˜$£*€ów#ØÞF•dÉ¸Bê,ÊfÉ¦ÀØ:ªgU¬E¾så'¿ŸÈfÓä@¾8Ki…¬PŸ¶±H’a•'I=¤M‚u„»¡ p5Î›Ð½'ÐvUâivÂfê¢Ò°»˜ÕYš'Ky‘'dT”qRÒß½Å†Êu‘ÍÍMÒ;H ‰¼Nz°q»Kˆ¡Ê*¬2sæksÚ+“ªŠâ¢S³#€¹j•¾ÍßèN‘Ÿ%evlè`e±jV$Ìßò‹ÿ­€:&$åãîì²¨´QÓ·ùÛ>„fË¤.Ê¼Û°‹Ñ^5Ëßonx+ÈY‘Ã(KªNÍs-›çïÍÍijx­oÓ:-rNIÔ"x$~m,³«Ðqú‘ö2wEž>R¿ç®ÆƒtM¤ÏÍÞ°OÍÙi/ª"î<r’Úz[1ÝDîlÏ<Î’’ˆª¥1bHI`¯'K@­ÃçBšÃ˜Ò·Í½¸žFÛR†i<hf|6F³º.ò†9ƒ-KÇïacÃ}\5®®*©)@’x;Àº1xØ­PQîO“¼_—°ù5¹nâlN£¨Ú¾&x“0¼WÉiËbÝ¨äïDÛnÂüç¸PO’Ã?Î¢2!Uú#lÑ«__“åÆR”ÑûýiTWÛÓiß·±Ì&3”#HÁÚÀBÆÀðµò}Ñ•"yƒ‰#ÿôOä—ø£úeOjOYõþòvÓ…ÀT!?Ž?÷²;“VU°WuZgÐ£—ÉI
²JT—`Ì¿#œó´?«OŠ4?é†	89jÂ€¶ùŸsš›V:,ÙÝ8­a;€iZ‡¥ÎK=-ÊÉ“¨ŽúÍ´DàÐº§z|+«—ëR¦'Ë‰÷.e“I”f¢ }éRj<=eàg§³²*d|éRjZñ¬–åø+-ù]Å³èÃ?ÃÆÜÒÐÀbZ”õ,‡€/]:%t ¢ñÞi@LL–â"~‡’`
â"+NÒH”Ö’ºÔpØíå'È¾È‰”ÂìyT—éx–E&(er'8RvxÝdŽ±œd7›Ê_·-Ø½
—ìó"Ž².›sÃÖlSy*a¡.E#Ü4)÷È5K’™¦QAI$†*EÖl)Ê²V"Žc@
Î‰UqÊU¬	ŠýmÅ¾c2ÌvÛ'@ŒjJþg×ÔT¬ûÚµõLPÓÔ›€ºrGeTÞèBì4¤—núÀŸ«ã4CÁƒrÊÕ0Kò“ú”J’+!6yÃW7ûF1.²CØ¦7¯¾Í‡õ_]cLgk4ÍÎ`/à„ö!`¾HòÓŠA”)€@Œz_@'¤¨B&œDS¨Ò«á]ÀÝ4dñ¥™£Q–è‰–òÑx…Jù¯«å/	g=ärN¥OäËeÑöUj’-œÝÄbXÌß€Öôy9N/’äT K+äG\.óùÃ¬ªÓãKñ*uW¿YYYþØÏhü>.‹)°Ù×)Hx¸ÒŽ³â|ér)‚]ØêìN@…ç§Ð¢\Žk©NaªÎ—à×9Õ-“It±tN?ÉºOÓh?™\.ýšDy:ÁN¥99†½ÿþXüÏÊˆ.ì5G/‹pfPým¥J¸€!t@FI}ž@hMÓ¹
q$~H]Ÿ#Ýnœ~åÐ1 F¶L@ËÿzÅ7;'ÑÔ+5¤|måÚ¿3¡6ÈG©ì­ÄÁöÓ¯œ¡(~†Û³GYåß‚}\yâÉìÙ§)LR]²ÊX˜MC–’;ˆ¦½UÄ¢tä%ÏÔümZàû·N(Ø‹Â†‹»È¯øg£IZo^±=aÇR”@´eR¦ï4Á^'eüÇ\Œ&ñºz]£Hé;“°ëbPb	å½€êg#‹FIf‹¬¿÷(ñµ™C~k2ZZ(“^ ¬G„ÊY_zi>mÕÛŸ4ŸÎüÇ)õå”c•+“?ÎRØ–½Å9Œ&9¶u<Œé.Ž‡C½ù0C­É¬ÆYN¨ÄµÿÇ?,m¦9åæákdÖ†ßÈµÛ £­&‰]“Ç)ÀŒ1HÆÑKÔ™sæþÉšf°N¯’ë‹uò³ü´ k+k_û
xH¨oÅÒôÀéãÇÃôBÑñûŽîw‹êJ•‰¶Ÿ!~ÞSt;Ê¤{ŒeBáußˆªRÌQ5KMu¨}ë°ü¯‡ÂÞSÞ]¢jÍþþt]”ùàzD¦ZS?&û•f¿o˜*ÔÀ{r‘L¦Yñ'8êá¸˜|FØ¼sðô“búŽ§Ç÷yéi£° Ÿé+ùÄé¿ÿÑÏP~Rôttî¶z£þJéî=ÅØzþxZz@zßWáþŒ«ŸWø™÷¼ØÚ`#ìÃ8~´~ßpN ˜â}T‘(¿ü<‘ˆÃúa{;Åg	Ç…j8‰¦ýþ”êÞCrÂfð}r¹y5½“>½ŸwÂÇPeÖ„¾Ë&Ãû°¬…ÑÆ=Þš/›üt‹R3€ùy/ø¤{AR~øKÛ'w¬{ŸŠÃn¼éóîÒ²êgœü”8ù\¡ý¤h©ÃÝ7Ô4M÷~FÏO‰ž8;‰"lR$5¬-ïšÚö¡?#ê§Ö¢¡ùì‡¿ ýì=Ð¡I[Þû†§®òÏ˜úIE'f©}+„9µM—˜B9H°ÿMë&nyWôNî‰ÞÁÑ÷CÝ´¹ªüä÷<[50ÞD+©)"³`¬“ÄƒZ@j:ø™Ö€ºßêv3Ú¦ûlWcŸ}˜ÚéBØ\ªsÙ¨n“˜Æólk+(q•LR›Ë¯JëTË`—¹F	ZmWýëÆ4N²¨ô JèfCÛôTÔ"Õ7=qZ¡}¼y•Q½h1{sqéæ[PòrŽ'nŒ>‰]ˆ,²WäŠž¯ÓhœÖxC³|ÅHÉéÁÂ=*`£é‘uþ^’m¼á7Úþ¬t"¡‰ð¬åeäåç¿¢°!.ìŠ;¿ür³,—V¸L6¯*íÖ³ê)®º¢JÄª³oG³Å¥²gI3•~ózK»÷hf¦ÇvNní2¡™]rˆn¾Ëjªð„¾Ú¼:GxDÕ]Ãé÷'t„“aNæÐa„szlô ŽIuâÐ¡ô˜ôn,2užŠœ0bj^•™YNZ¿-lg‘ôÞŽ²(ß3“ê—FšFEÊqz²y%šÈãÇÁ"Á[8Ñ¸Æ9›8²6ûÀ"0+×vYF—Ã´¢@·p¼¯WÞÀÂîY´Îâüš­ †™Ú€µühîE2úuØq¡<·ò†KO2Ýeÿx{Sl9B+Øž*S`bªÙ(ölºs«¦“,ßaª_ëïo:‹‰~ Ûý€õÓ
@³Þñ	PÈ‘BÖ4g³8©ú#á‡Âç¨‚~Ô±hàöFLÌAt‰ tÆî;ËËI_x˜ZP^î5 §Þù¥À7&;<æ¬ÌHY¦fÁ%Çû¾uPë9ŽRßGeŠ{`ÕM/’Ì±²bFÙ™g¥]+Í€Y‘³uÀè„;#gxÉX¹Í©í¿~£åÐ6Ìq­îçÙåæÞö¬ümv!í xÿØ%m?å°ˆI9J:r’ 8ÇÏ«Éq^£·©ÂžÝ!xOxZlòµlTvTD,FþC/Ei†S(§róêµ1G6º¼O.Oy“ùÆ¹•M¥|ø,.& f;™¦er–&çß£dyŸcËä°˜ýYY¯Í×+Ñ6Ýæ¡qÙåàÝªy;ž¤9õ{Á|8ú‡ÆM±Â£{Å2´Œg#;xÙ9ÊZçï÷èí˜Izá¡n¢¥7yrNž œ2ÖÅ³b#¾ÂRÈOú=È¿ìºÁ¿Šfq4ŽŒ)8Ä4î‚ÁÄ~öáOðß($”ª­eï¢’*¬V`Ogù˜JÆôúóË$/Î°/ßC{T‘CÑ›âØŽ?§e$Ÿ{SÄ4¹…±bœè°
->¸^'²‰uÚÊC¼¦„–··Nú“uRQ¸.’zÆ_ÍÆch®G€·LÊ²({tUŸiüPõn¼ª’ò€½<ôtuXìkWöz$~>Ô=ªqŽDŸ`×T/o´%RÞ Xç6&Ù¥+mÿ`¥ªÝB$ ±µ“ë×o¶Z+×ŒCdœ×3ìvQu™IŸwx*ÖHQ¦'ÉÄè8[I0-rMEçQZ“(ŽŸc(Ÿá6 ÔG‹dgÿÙ³Ý£½ý‡ÃÃíg»‡owö_mïû¬%ou/^?aÇWÍ¶8ñÅÜè|Äww—åöt"}/Ê³_¨B5PiNOjbÝï	‰Vw™ÀšŠ·d€_ ¥h³M¦}¹ÞÕÞË±¸ïjœÄ%|q÷°O‡`½KçÕ\“qTOI°N”’ pŽ‹,Òu€ßÝ–wá‰
ÙžÇiÑåK‰7ù€³^Ëka°
îÀ¤YE=W&\ñ«î;
7%l:Å›š@‘Â¯HˆWjt._ØTŠWiµfx©AÅ°™ 3zôáÏã<³—ƒÿZ-8Žm|Ã¢)e‹#†"Á8~‰Öa_ÏhJ;[aªpsM…ð3öÛ °âÃx*‰ÊñéQRNK&^b½ž^ÉíSÊ±Ó2å«^fƒQ€×o	¾N´H5óZùC-¡¡U>qZÁ=¥¡$Å ­ÜŽzo(%fOï©‘ÔXVñ¾‹¶²¥mB4axÑf“/ËÕ”Ò“×fË¹‹>î¼¹C_¼èe›+n2ËÅÿB¾šËŸLŠ8Épï£ØM•‡tP,—ÕWl!’ËŒ”,/“ï`ÊŠ’æåÌ¤ßQ5Â·ûa	dIßf‘ž"–˜}Aä×ÜRäË”·2èÚ.òQø:{ÍéåAŽA4EjTl5â éÆóS~ñ<)wÐo8±Rr±ZÑfFN—èüM^…ZÖ¦Ÿú¤Éß‚Nûjæ‘›dA_ïVš¯Uý£§a`ÕËŸÕ#\Žm2˜«ÏžšØµ1Ô$°ºh‚	05T-w;L¥õªlÌ¤$Ðÿ  ”þÉ4³*5Ô³B«K¼OÓ„NÅÄ(ÙšŽv"ŸB;YÌ‰Î ìTÂ»JSi×Œ•Ìl§Å:fÊ¦ã²˜ôQ<:Lê>[ÚTÃÖµ†vƒÁ°‚}·¯Óª|m:£Ó6õEÏžþÆ®‰ª{×(¢„û%€aUÈv÷p­bº‚Ä=.€TEù€²‹ª1‹cg^G£4õe‚<+²^»g k„2¢ŒiO’ãh–‰úDclogzµ	.ŽzÜ2^‡\Õ_þÃ“e`B`SÄ3”ÞC»&yq]«NpužJŒ®¤Õ“u-EUn\ D_®*ÃÒ /»þ‹_àá€^ˆÔZŽwAÏlö”––Þeâ;çÜEWq£Ð:ª¶ ÉT¿´~øð'’\Àï„äéƒPSïYx§tY6â²Ùb¡¦›3ž‚ùÓ=ëÝD{¼}ØËÝûßoïlï"Øp8<vì0„Ý…ÏQà	©*ÊÎ
	$ØlKªl KwAD.4KW¨©fŠ+\Ðåa
t4ºALA¿°¡5ß =°¡Ø­Õ–"ðñJø¤	Õ¼#QèóÐ 9Š	‚/C®ap„@`·ò(Ëâi¨©ë„¹´ègƒÓ¤raHk’6,Røx–½gÎ9ÔˆŸ}@–ZnüÆ"ÃÜ|4üpŒ:žr¢Î/Þ=Iªä‡Ö5ìØiI¾¸òW-éUìð ÄEµõNÌãïÐõELç–*`,Æ¤8¶OKÌƒ>¶èc
\÷1þß¼àñèc W–~ú&fÿ]—±•É¤8KÑà;­²æ#€Ìóè&¤e›>RyíÅµº89ÉøQ! G?5mÜi2~Ÿ@Âˆo¹&ÒðÏ½!t°9ßÈ%’À2hªÀ¯sØñ]EE5KM£ÛÎ2àÊtyip+ZWZ‡&§ÛîµÎDãcPâÕ´½‹rÓé tœMÑ8¨3öK·¨úÆÇškØ±8çÿDYú#,è…Š¸.äÔ%åÅUÁHá°5¹ÅêY„*A6D:÷ŒÙV»Õ„ÐwŸÔbØ»#bÐpŸ}¸#ãRÑ Ç2àùr­!ü—ÇÊYŒÝó‘ä-FŽÎø¼uÎŸþ™wWÁe‘z ‹Q«ãíží¢G»Õâ^ûPr£€íÁ…Œü®²ŸëÉ(ì»ç\¡lp~Ý3íñµñ|\3¾OIŽrœn•K¿Ø™P48*v/ÆIÖ©òªGgü­<æbå¢ÐÞD »£‚×Æož?ÛÃ‹» 7¯9ˆ†$(.í—ñ½Ú¾^y£¸LTÌ	K§?¹’u›ž³	²€c0ÌÕÌZ†Ö¾èFQçüXö°÷É¥$±@·ý¦#Pâw
ï~ ROKa=ÆÙ¿Æ¶Àwä[0›m•Äá€fP'TÉ+Pä±m©ÔO{ôÞT†Pöê6S¬‹\e› ôA£¯Åûº×¢™"twH®d1òø•¤BF–Ö˜ª‡¬)ÈùaëH”¬eš®:½ dJàš¾d¨Š¡Æ	?„–zhÕMó+ýSéLp[hkÏWºNPÓCÑ›8ß˜zª6fE4#µ7mmL‹ÊßÀôÃ¿V¡Ú}ÐCwþ dfN#óÏÞ]MSg£Í‡ª{wû‰z,ó”õ§Q¦•9P‰EÑæ×[üÐüÐPZ÷jŸâkCöÓtTÒ¨hb¢yò‡¿°tUÓoERCmqz‚.ÑµROxÊ-ð€í1œ;›¤åNÑa*jmiÈ*6{«×çZª]ØXcZMYz‚['IŽ¡’(ŽzºÎ“—®Å–ÜÚ
Ç"=}ûû3Yó®¨¹·õxÂ"s„¨^+‘6”ŽO£r»î¯àt¾šNåFö+úµÊÒqÒ_4aÞˆ2Ê›”Seœ)ZË›Ì)>LõÅíØžOs.2k1d™1£¢½Þ€£™¾y*]™·ÁîbU"«ªN¨yX1¯‘d†*1Ú§*8oÓÀMc«ÁmPé•ºŸí¯›³õh¦&ˆH…Ÿ½ƒ›Ãô¶Âsb;¼ÙR>"¯Üh·d1M´a+Íœn£nž"üuà ñRé>½Ð
µv¿›fFˆ+àÁ²Ò‰0~R½Žý7²>¶ Jh¹±Z•í¨aé}Û‡íÊahÔ‚	©ø=°é°³T=@q4Îw€¬š$@UÙž›aÒüÉ[oGP'í°Ç|‚âÒß¬&Eâ1ÐKuÔÑ¾r¸XëáA:àËL—ûÕŠi]Îu¸X-ß²¤æ¢j‚–÷+/Õûö>ë×sçÐÚ™£˜)ŽM¡|°	YùdFM¥z>Déî9ðH¶ô4ž€ TÖ¬±}ídRóÓGYÊ³ˆôjA71³çíŒÑÛ­Z^¶÷Í8Û4Ë7ôÁ«õR–ÐáÌY[	Azgƒ÷Ûúo~0é‹ˆíj°6@mhßBåŒ«Â2ƒUÇí_ýÊÐ˜qöpd·r>p	=§–íÝW¢%ý¼%t£2,CñŠ¯¬Gd…laqžpMÒ ´”%1=º¦Û÷ð=Ä¿~§Öòõ²¼[P;ÖÊ”jÛq¼3«êbÂí­oaÛ ±RY¦©µ§§æ‘ßÝ¨ÿþ·ÛG‡ÛoŸïn·kZ5ã…‘u~ûëm)uŒúé45ˆ³;mhkß¥ERq©|ºõ¡üó$Ç˜XÇðôŒ>2ã•±_›Ÿ,º\ÓzˆbZ3KÄ«Ø-Þ‰è´Öãuè…Ïêù] ì:,"íˆVÅ]ysÅ1%3S¥šY\9ð_ëxÈ·3´'AÓþMçÒ‚s×.tûÐƒH”>ò6ÄIù#<'W¨l) BGåøôþQÚÀL óÊB­!ágU IzN^¢Ñ¶Ê=â¯gãm‡	Ì\U¥“-“WRôË ÁúJÃþ@ÙœP«2™—Þ¥!ìö†;…ô³ˆnDI"Dr\#ÇEÊGCò=-q•Q^§%-vD0ž2"H”þõ"…	!GSo?•mÃ"Jóó8e—ÕÙYX•žE«ë×ìiñC²À;“Oi4±]'9ËÊZZ ûÐVôcÁêÓë‚þ0 »&°ð=‹Í’Ó8’[•¿‘<‚Æ‰QŠ€œ˜ ·>WwCe;ÐYNioÂåš[™K2uT¥<l²æ6=Ça±ÚZiM¥g=¯f<P&u”mg³œZ:ú¢Y>”y9ÞzrêtE™.*S_Ýó³[´ëÎ£â¥V½Þ­%Õ0í¸aÝjû8qÕ~-ý–x½ à?èUƒ©ÑßeqÞkÅD¼€4Ì¢wÍ^²{=9ýfe…P¿¬»yq^FSÃÓ
U?ûð—?Î ETF,ÀðÈ°:;–¤¼Ðç(ÄòÐâÝ)d:Â0CLŠàößpGGâFù“>C4§á(ìZ$YQ‹	¶[Óý[¿rkùà8 ’(BÇþÚ‰	Èb£ïñ%—­'Ýçö£ ]¶ø& 
`v$W èW7¨7°¢çëÕ•éÅ?HedPä€“ž†þ8‹Ê$²Þ_×kšŠålÚ¼šÑ*ºàS—KÓ‚©þ­Q¾šRÏm˜ÆÅ"Áî™<Ù·ÜDãq2­7{Ã‹¬ºX$øÇÎ¢|²é–
¶O¾,ì«Y5MŽ3¹«€Ù´cûêÊŠíréæ«y_»9õxRœç]æ•Ð7¯.b^Lo_-Á@×Ì` _»›R(þ/Æ  üJ¹](ø\Ö­Šp¸rß´6öZØéÆ`ÄîŠ²œÎ¦”}„flüŽ¥ó;79!Lc@qÜ9p[n˜¼½øJîÔY$Ü0ÊHÀv c9¹·d¿Þv&ôíï“Í…¾sÏ3’—»“™½ø{w.$kø	g¢(‘†~Ú‰àmÎ;T»ÛÉà=ñÌ…ÎÃÂé(‹ê“â¬óô,±c7ÎR4ªŠlñÜ¡×´Àˆh.èµ‚Y¶þ#Ç_·9‰nB›Ÿ?ù.šš­yÜÖuEÚ1L’ÉÕÏ†À·NV®oÔ‹‹Jk·”Iíñ”û4Ê@j¿êÇ‘ŸÝEë3.÷ûF§—u)u@¾$€ƒhºñ4½HâþªëOÖ%!Ç)žGÍhZEÍLƒ-º~þÀ\.sÊÂ€™vóÅR.nžLôSP*ºDLØ=AØCØ½L·¥ÒiÞÇxÕZ×hÔêáÚÀæ¼Af¸Sey2†¼Õ™²qñÍ¸ž!îº<nl7Ç]œ3OÆxÌùµéÔX÷{Ü ct`ƒm¾·a.=³éº
ÆòM˜ë´ƒïë–	ë8Ûm÷§›0Wô.Ð—fb7ŸpéygÁó=pÓUnLús›Ñ{ò®K´âk{oÐÛÛ-¾æ µŸÝ¤u$‚ÿþßñÄM³¦ºEl‹RqKâØeâ³› O»ª¤åäÍç7ýîvŠ^÷ÙÍh8Œ†Ê`pÏÛA²98çý‡§;Ü9ãlÎcÓ_3ÁÁ«éè0ÿÒvÖ.‹Z·¾ùºnÔv;„l´vÿ±±½nÙ%‚›‰„Þ˜Ó·‘½ñ¢?'aÐß¢)¶Ek\wlÚy³²BSEëá-”fàñU îž5µªpTg*¶¥5Ô‡ÿÛ°·•iž£T+Åê$©#šïü(¬^µ@Ý“ˆÔv¥ÛoÀoCËkÛs¸lÓœz*‹:3×e9*£Q”Â4ô…ÅÆÀ µ¥¢j6Ûpí4ÂZcw»`NÞš¨T]ƒØZ®ÖÅtiuy,QL¥C½¤	¦»³ðÐoíÅÔUi
gPÏÇ³jÀEÓdT ö7pš-ýx A›tÓ‰qÒ×+W+{br5¯ó¥¯W•WÆöÕtÄj‹Ÿ- öíMÏgYúøQÆÞT›Wš#>»á†z‡¬ÊdxG2ÜÂÚÙLíOæoðßcÛM>>@.Ÿ¡mAs>çôÝ7T3fÑWŸ1ì1	NT÷XÙ¶·;în›±ÂY}fð²Þ#€fÃ¨„#‘`¬²¹¸î¹xnÇê…ÍmxíÏbÞ5g¢]f]sV|WsÎœZ¶Ì8ËÄæ{Üe¾Çr¾ÇþùÿmÎ·é¬µË”›n¦ïjÖ¥óÑ–‰—ùØÜW]æ¾’s_ùçÞôobî5÷Æf^Ë'óÎ*úÆáAoôV!Aï4èOÔ‡ÖóÉHaÊ9Ñ‡¼5J˜™–b7rý¶tQŸ&Qì€§.;X #ä`#hqÈÅL´>µ–î·¸t¿–zuÍ´šú_6–ëÓ9kðÆJnËMÝ ŽŠlnî$qÓ`HlGœº;Qßû†®z¼¹}¶)øt‹Âlû„T”F¸~ôÕãx=Ï\ÀºÉmÞ28îyËª7¦ø‹ºG,ø\žR³¦_üÂ«é¬ÄcYG+µþ9%ŽÎBÚ~©%È4V¸³•¸•ù1Ù¯'Ò‰¢î—Ö?®`Z¿ÔMlžŒŸ4œ•Ù­Áãè¹Â¢ºòRAºˆE«¨]#ÚÁXd'¡ÈµóÍÊÆã¢6®KÝ	Á¡Ew17¾%i¥½uy6‹zTÄ—:Ü`»z¿tIøp8é+;Í„¹vMz!´vC"Æ
¿­-Á‹­{ß¬,cßèBx£ŽýTE·2õª¿V›§—Ô t(È¯Èªÿˆªv¶í¦iØ»mˆÚ–Ù)½é†V]ãN%‹/’N–wËº;€5])”!ìbˆQ¬›šY8rû­ÎÄÝ›D0Ìh%O×Þï±TÇ>øÀ¸½ãµâu:£9;	ìM£ç7Ý¬…©n…›÷t!’v°vL£-fÌ-*U×ÚÑÐ~eøoó˜¦·ôòãÂBx¾5(¦³ršµ Cæ¹8xG?
4<*ÿÖOóRAyAÎ¼ŠÚ²‰òY2ß¤!20³aWÃpŒ¤RA×­'8®kzÁ«X
›@íçÕ8#sO åªÎ[ºâ·¡¾âú.z¬î·`ÅG»²/5Baù”ö©á²æ\íéŠ¤ö&Õ-]CŸ«AMõÔÞžqùÜº7W«U{ëòŠ“!íÍÕ®®Þjn5©]r³. 54 ³¥&ëÒ<§·²yþÞÔ|€F¹›*@Ì«T3ŸB4P‰O1¹›9¸×ÌÕ‰O¨PU=ôˆosîzlihCšG-Ò:ÁR.û>÷vÞºv#”ýº«ÚÃµ²/ã:Rí¤hó*;:y‰0JPJ„ånâ¨š4#sÎ¢7@¶d
N°ý’Ùi¶Z¿ì¬†ªÞ_ZZ¨@×g+ÀYÃû=Zž/?1É ÁCÚ æöãe8ôœ³LoËìÏê“mÜ:Ï2§ŸpŠU¤vU†W=ûµ¹Œ¤Êvªû¨wˆuö®E%…OXGeT®)S­\›Àå'œ®jÐg}å=Â`±e½rÛ†[+Kv˜ñÙ!àÀæÕ7þ¹ðvóTÉ–úÑå:öAÙ¸I~:›¨°IIŽÑáe\ö¢¢‡8è—3¢ñ Ñ÷åí`lŸSÝ«y’H »Ùv
r)¨-ßFdÙ´ÂEèõØþ¢{%ÝÉ c5M»±1‡*{–D1“rŒH×[”D2OÃffjˆâä¦ª'_vy+ÈmÀ"Hæ«Í«ËÄž:\ 0“ÊíÁè í	—;L,ˆ TŠIzUff9IŠ`;‹¤÷v”Eùûž±ûkþdÕOyÒ²y%šÈãÇÁ"zi«£q³7ÿpdmözÆÊÅÂå¦ýË º…ã}½ò†¬ã»€ÖYœi³=({Ú€µ<ÿr±c¯7-çp*¸lz;°¥1ì·Ú6û)p÷¦xs”L¦HZ«@È¨Z|—A‘µYÖ
Óé–ï0é¯õw#ì«XLù3 Ýþè‡@ x‚'4¬¡áKÔmOLÂAtI½Cm:=`'[FÀ»ððÿ&<Ð¢Ýk(@þN½óMôÉßs§ˆ¶P¡.'gÁåÅû¾NxŽ>ßGeŠûHÕM/’Ì¹ÈÌîªgžUum¬*fEÎpžÑvFÎðò°r55ÈÞqQŽhu?Ï.7¯PPò¬òí<Àxyä#}'å°¨G}"ê”/ÿÔö™xmŒÞ¦ !—Çz)\»›Îº5ª¥nA7¯ø½pt¥N¦œÔÍ«×ÆlÙˆó>¹\'=æõÕÑ¸P—oë„{Â(uâäBÇÛirNMÇ×1ZP‰2B1ûÑŽ©pm¾^‰Æé>­ËöD,«æíx’æTa·_œ
ýcãJÂÃ{pb5Ï³‘”¢¬epþ¾`Àž·c&K…;„÷·Zzƒ± [ÂcBANÂWí¡7­—¿ìºÁ¿Šf1b™>‡˜ÆCQØ€@wÈðß(,\V¨w}I.îõƒÇ³œú÷EEßC#Ô=ÃI4ÅÚ9†ãÏiYÐ(‡ðÓÞþ0MnQ¬'9ìE£A˜€ Êc@Òb{\Ðéõ›Å×ë„7¼Ž=¡,ÄkJ{y'ÖI¢b×[ ‡j6Fg¾=òOÒù..ô³"¥ÿx—×É«*)ØËCOÿ×‰Åm°våÖÉcñó¡>:èQdªO°UjAk´eD#½3à²Î%´‰ìÒ•¶¥°ÒÕ"jÈÚÉ5Âýnk­XÀ†Ì;ZJTàP¯)ƒ‚‹K1‹æ=x¿Ç8_^f<=fö´´ÈŽxk(a\ÌXtn^4µ¥lõ-Ûý¦öØÑ ÞžžÒP *ÖJ>×SJ.k¥¾o%ÒJîL‹l¿’ïF)µ3ó‚Iœ¢›xª9©wµ#†»X[°fòY–=êã¿zUüä¢æQ¿¹í[F‘×Áu8E#.öf”íÅQu:*¢2¦‹>ƒÍ(ê=bkGû¶ø€3¼jh*Ø‹Yÿžh	F‡IÍ;ø@#âÙÃ=-Îeñ
Ù6b'Ù5}®uf‹O·žÔê—O‹ò÷œ¨9à×?¾álZ±A¶ƒ“2MÆéq:ÖÂ#@ùçÉ¤èë
vÎÛÄ•ïE"4Acý±4<Ò*ëê÷i}ÚÇOoU¤¿aU”u¿-’‘ÅØ£`yª´ÒÞ®¨ ŠK«â#OñQ¸¸Vš§°¶†evŠÉ4*^«rWÏb®,’×6˜Þœ Ê:ÕÂVXÐ+fp-RD(HÄ;+LùR=Žn·ŠêwÛ[2„›·PTs-Ac,«5S·Ñ%:4nÇ=wãú¼ë2¬œŽDÆH¯†ï´:›… ÛQ@ìL]7ÞÃà§*(ý¢0\V¦È’(·°GØõé1Tœ‘‡C‡À§±¯kÉSlA˜¦¢¢ÁžÕœ5n;£ñ²ï`PBWÔigv2Ô:Ò:09ç¼ÏæâµªÔ
æ|!ù]”)wãD«½IØ£Ãíg»‡owö_mï™Ñ8x«{ 6þM<åg¦\_LƒbCÅwW¿Á=]ˆôUFaEJj|¡JazžæÔX¬Ë´$·Žäœ—‰5&Nd E%HÌûÈÓH®í2|vBˆÚáuMÒÇˆÁ°N±É¡Õt‰êhM¼H—EF0qî±ˆeñ…;•u±‰¯Ê•”H1<ÐˆD+J¦Hv<ŒÉNŠ0¢"…có1™½_ë;'lô*{…¾BÇEoZÈ™: gW@Vë½üÎ€hKßãtšìÒö‹§'9®ÖÞî‹Ýç<åæPvŒýû-üŠ1dQ^ìæUšã—|mÑs+:ˆ¥(1¾[ÜÁ¿lhÚÈÐÒ'Äzif p*ZÔ©8ê§i."¾X_PœÙÏËš±«m”«ñdÈ³S&ÕzzØBókO¹¹#ß%¥VÚÊ¼ÿ;‡§÷N¦ Â€eIu]Ï;¦Æ<PÈ—TÃ+¸
æØP9hî]ØãW=¿¬ðÒBíËÛ^TM/ê[´[¿@Öàñè|>¾¤R/Ý¹Ã{R³)U	¼¦
#†ÊoÖq’ÈÕ¦¤OÚª^'+‚>}øó8§8*“zþµZ:Ñˆ€üÈÊŠr7ŸŠÒ€6E„¾Õ(Q×iŠÆŒà5È¸ÝË y¢>üjSðÔ‚˜sï~€ýn˜ðm—•°}ûuÁVÝ™åõ3ú6ý"»ßÅ^˜„ÊP¢PQ.‰ ´ ‡úÄµÁ e®¹±ÞMõþï´é¶tZÒöÉdatc¤%àôœ]§»ÒÂŸ~T2 ±ý[vÂº_’´8×ÜÒ3›ŸœZ˜l…­¦ëš5¯ÿ«þWÝø¯ñÛÍÊmû±PpypwZ"NX>Xœ¹ÆwX‹Æ&Ò,Ô³ƒ'.ˆœÓA*v"_p²¬¾ †„*gZCN§L*¦P¥¤97YŠ#œ'åZ>)»vu<ª³fVÉÍÐ(ÊdAjfÙÈàuKÕ(?ú«àY¬ÆP×²ªÝßd@RŸýÝ¡Œ¿" š×¤ôƒ9B­ÇªT p¯Þ®¡6[æŸ¶BEæöÓ
Ù©OC.ÅÈ1¦7÷s‘…¶¥S"C›joàÖƒõPÉ8P™$
ƒíº´õ³ FCqÊ]U=Ï‡Ä6u¨qwr®m¦K?´M£{«Z!ÁbÆ	ŸáˆQy]T©«ksœF#º|¦¢êK!iÀ£“C*%¶}Z-È¼ñ¯+Vç‚<àÀ¦ûêÝƒH'=ØIxÖÉûå¯êvl°pJ©]bÐh››*È•'?%>žòb[hFÜ/01NýÕôý÷AÁW	aY¬	™“é’¹ž¤ué‚±ždO‰Õ«±3u*J¥ëaö§¸NêÀ´Ôò©Ú<áé9‚!óX×ªQŒH•Å	Ýz™±½˜)MÎ¹¨G]ªrC`ÕÊô¥/JÜeCyC·@6
‘•iÓfÓ8j×$/BûªÕuÙ/¾é‰ÈF€É²ôÇ(¾YYQ¼T-·G—eÃáX+å>Å´„6Ïô†ÂòÉããÓ,‘CªÈîÀãØ¹®šóÁŽ¢ªÇr¾+0ª:Ûá¨› 3"ç,Kßl¦ô¬C`?e´ÌOHzƒÐ¡È9äèeqt1ÔÎnu2ÑýDâ„ÞZ&¨¹§«XYw8Â/ßÓü‹ÝÐX¥U¥´A–€¤v7ÞK)Q°-´Ñ‰ÜL°zË-qX~úSðx¬Ñtš]¾d†|ØOìImÜ=x"·Š°áßE½Èh¿Ñ¢Q¬˜kÞ…€Ï(Þ“„ oZ†áT…ªõÏà@HÃ<¨û½‘õd—|(®	Úö/~!®å˜jÞµoLªªzôÕâ&â=/~»Ú3öZ‡ŽÎâLÖž$–sHãŽ;AÄƒ…Þl¿Ü~üá?½8äáÙÉÿü¯ÿùÿ Û'E‘³büá_È‡?“óÈ.†g§dù?ÿ3ªó|øOô¼á¤ÚÓVžÌ²SX_d2KëÖ:åÎ¢ZaP{zh’cX®q:ÎhL.×Pœø4OÇi´ ü¿ÿzýàÁÿü¯ÿåÿ"Ûy”ói4´JÊ")}\$VU1ÆAÃœ?e3‹Á<E·‡Ôà…ÍÔìýö˜v§Jf„Ö’üNÙ`<«¤Ç¾Ñ(J/Šõ!O²4‰ÒI9P¾YQE“Vå?P(UC<ÁÁæF%Á‚ù)Bx t’¤H€èÙ‹Ùí•ìŽ¯}Ð“ç´’ù¡+þr’Öìüûºsð”üŠüü‰\JHð/ÏA¯B@uþ­¢°ÂœUr2ËcoÆ4<Øff)Ìtzœd$õyŠÃðàÁAÙéeÏ<ºpZ×Ój}y¹J#müø
Ÿ—Ÿ'EþNîtR–ÀÅô×Ì\+…g~‡ˆäÝGÊ" çÀøk!aDâÃ¿L bˆP°x¿@`fGþœWïÄN,hÛv<@ª.|ËÃ4j©È•¡èÚŠ³ìÇ[grÙ˜Ñ	•‡åîÓëg³
Ê>¯N‚@Œe›½ŠnÂÏ/!”Í¨	T=rA.³`×pAÛ4}¯rïÁ‚,È(™l¡u0FÖÃÀÓ„"q4­9	d¬ç6²HØÊä8A:Aï÷ÖäÅcgºPÇ†Çš!ÙC)-‡¢f¤¸Ørlô¸H	Àq|HyÁöðpåÃ¿ä@ñHL—ýÉWTeðå_\éù¡Hó>²+èPßc\Ó˜ˆà.VÊ"°H•ÖüÄ‚ õ(O ÄçþA”ŒG' 3}«¡8Œ<ÊañGŒí‡î‚GéÖÐÄiœ9ïVˆ{äŸ3Û“½kº“ü+˜å?äÈ_üû¿iÊˆu5éV|(·SÈÅy$â4OÁÈpƒÛ;»{G»äÉ>A[ƒ—ÛGû‹X)Òì3@­ì ”Ÿ¥JÓ“#ñP^ØžA°ºUèÿÁöááþ:T|x¸ý’ì“ƒý—GÛÏhÛ/žì=6ýËË9C¶%éÔ4£éU~š,ÁŒ-É‡¶´öïÿ9éýÿõÕ.9Ü}E7	üõâ·ÛÀboîÃïÜ&Ûd2~¿û¿/ÊÜ»‡ðïÎy¾yY‰!­ò+¬òÙÍó\Âƒì>yµ³½ì:‚~9ÜýîRðÁ(wiö½¯¶_šS#ô
ì½"
côQ×"$”¼ÑÆMvÏ 44¢	9¤=IŽ£YV÷MfvŒbÓë!ßÈ–ñ:äìkùO–O`}õè€IX±&lP¯NØ%x*1º’VÒŠ’+¨ŽJcû4:©–^·ÝÂDMqÓïËáÁz£:pJ´Eê@W¯Bö=»9Z^úÀ’•ˆïƒÜ4^XŽÇ•Ve³êÂùè6ì˜7Uá¦h}àó))…$@ÛeÙäP3áò™´¿3·mÇ”q¯îÄÀÕê ÚlS{bC™™a8Š™[N½È¨1zMGþí _Ð+2eƒ6‡X’¬ÞfS"ó
•2'’ó£ô¦‘&g¦cÝí·``ÝÆ!µ*TE®Ñ7Å+ã‚ˆ4åñœxõŒ».]®Z¡ä`° 4ƒÕ¥½á²îM¢ Ýó|’˜¦£©LHùÇï¢é~†(	<HyiÞm›#óù)0BýChä067ùíÞ—–Û‹U.zÂpâ$©‘céó%º,…íg»/žl¿ÜÛ»½³¿{höÅº5*€#.ª§IJÑƒ÷eˆ.ônë««ŒdývpžóOœgE`.Õ‹>ËÝfòErþÑfR!U‡‰dq'Òï`"µ®´Ï#4ù)æÑ¼‘bÝµ.¨{%ü£'ü¯i¹H§Ý±^¤Ô-ññ[1jö¬&Œ´BÓŒ¯á">ãEÕŒÇ€‘öÑ6b¤¸`2ŠeÌÈ&‹ë%çWzW·¬ô¨¹™l•©5¥íÿú4úõá;hS­.· ¦Ó¾ˆàÜ~1°Ï.òèWr¬›7øxxq=Af{ÅK‰<ýÁC+pM­ùâÙ4G¯0KØ³ ‰<Z¼„®§ðj#$C¹¹¢•^[Kå‰ýÁ(
î»}d±WL&+á¼iï­nÔ¹:.hðú– êóÇ}qGƒïá51t[)uq£®`´| 'Ð>¥ø½‡vQ˜YT÷2k¦ø+áPáCÙmd>u“ÄPÞ;¨Q Ì’ú²{Ã4ˆ×¢[O\›•ayeˆUZeµ”ZkÔª¤ôˆ‡C`•¤ ƒ©Æ7P£SEqtíÌÐéê½Ê,{ïòp@c´s¸ãˆ»FIbþz"Â®Þ1VSúìa'q–Õ}HN16ÌóûéŠ&ÍÈ+?»ÒWPE¾¸’}×oüL¢‰nžÀË´¨ªô,É&TéÅ/Òk@v¯]Z¹{»ŽE"	¨‘nŸÈ}ñÝOŒkGr_<q®_€8¸ÎÄZi*¶Gi¥%­~QY¤ï‹=Ú’³]šRÁÛ.¹m–H=”Æÿ'Ê]²Qî¹¾§žèoÍ|=äè€Øgi·Œ×,q 6Wiµeµ¬öä¿O"-'•¥äúø„¦­h"£¥¬³[1²kO4a®í“\jµÑãNýÓ`"ß}qE¿\£¾°o˜qîckHÏE¹	®aÚzýNñ/Z”ÞÕ°Å(ˆZGÅîÅ8Éú1½ÓÒô‹°I&é¸­¦½‰XMR½ÅÜS×Æož?ÛCŸú»U¡?rÖŠº¸¯¯Áø^ÁÐVÞ(-ý&*C–N~ZÖiêû@¨ÚpÆoJ+òü{z£_çüºp
h%Žˆ_¿	¸ë¿ƒ|–Ý¤‚…¶;ÃwÜ›1›íŠÃýÜAP%¯Ü8æHóžöè½iÝHÑê6S|;1¸Ê6aŸ„F_‹wÇ
ZœŽëyôêt+<¯,­‰+RÃ$=¢Œó&%¼ŒÔ¢ïÆ ÌùR}¡TÇÝxCbJ«X¤+³#ºOÔOH|òoñô«7pµ,¦çjÎ×Xi:´]4õšTc Hv*4Ä3Tœnžâ}k–Kt(Ü’©¥“É„› Ö@’WmÚî•˜œÈçÄ¬7K5f`phšJ…HoôÚéAUƒ«¥6kS‚ªUìA‹í›%}H—ŸÚZ–²¯w,|ëÇ†xÆ¨°­Ýî¨G35×£nz«¼eûÄïÆ
y¼Œ5fƒ®=í5 i_³–´ß¥a—yD£LHÅï½>ðéÖ9ŸZ£±‡N;´ËnëWèÉã6îËU4}Ÿc¤ö¤Ëpn"½ñóÃ_ð÷Í:;¯¬/×Æ,æ	¯/GóWQ>ØÀ$øqØÉS	¹1™ßÃg]Ï_wRIý)Ë– v}å¡ñ¥zŸN§ö¾‰æÐ25Húvkæ35H}½>e	å@\+×SPu‰]ub|èË“PiV-Î1•ïR}x¶¾ÃÇQ@§¨"ÆÓ)M)ÓÖ3ã„ÔR+„ûT£eØ%ôÐ€BôÚf|ç;vó¹~W†º0Cê³œ¶1Êu¯ãò¯~eðÁ¶âŽÜVÎî/çÐ™J¢%]Âã¢Z@/0|õ°›ïdÂ5IOò‚iU ¿®é¾£††œx33dÞ•’Ù6!k$ç!Ù@íXÌaûåÒ·2´ÀÕò—ä(‘ÑYzÂÌ¾\`k¡Eã¢«÷óSH'SÍÇûÚEFªSèø9:÷…EÆ°çKÇ©)Öõ™îÆY”Þ±úš×+ÃeC¤›À ÖˆðÜ½6üÆ	ÖŒÞï•?|ÓUºIË¥/&¨þYë~KÅ—ùVDa0ä0ËNÄ/šiÍ ½NzVˆ\'Ø )t)Ž¼GåÎiTÖ_	î¿¶¸3gøOÄPl‡ø®W÷9'‘92»ŸÈúv'ïYZÕÍó†9(‡Ný§:áìÙ3=·_ñÙt¬ßNh|¹NÊ4&øF4¬€vT“uõºF²íõko´Ü¼>³²àËCHQ¿T1ñ4ÍÑb1-éÔB©£etqaûùOGØ¼BWšNx/è\QnöF¡PíËsuXêP¸Ñ°©6õõ–9âðw˜ž]î¤å8KÖÂý±VnÙu<ìÓ;BWZ;¯¹	t?+ÆÎÈT¿#4„¾m¯¢‹}1°:Œ?í³éáäúïÞù{~T²±½š†»¯6ôß\Ž4¥m«f®šoíUcÕ¡màßxhÛ–»£hãô+7´1Æ¥ñÆ!ý5Ô5ccûB mQe¶Fô¼‘”¬¾8$OÍKG³”™=w™µ…gjùô+g¤š÷•§ÓW†¦Ž­A‡yÍ4Êp5DW×¾fBK½LÛ³ªN/—FI}ž$¹Í¯J&icRpK€/wzøMC<0Zgðk#`?½snÿA3]I_`åÑîf:)üº% YÄ(Ê4_%š{QEäùjÑ.Í‹ZD\ ÆZ¡ƒ$!Z3üáÍ‚+ºÐ%4îRäY^__õH/<„¦¦dÞµ–ÍÁVûW¸k	2×ŠÀÙ:ôÆ˜RáX«ÖÒ>§˜NŒØ˜ö: ÈVg@¬OÓ°4¼È¡òàp:¯ÆÞ)kÕèB#Ãm>·\—w°*ïbMÞ~E6¬Çª¾ÆåêŠœ§q}ºN•V¾Gßæ
ñõ~pCpy2º¬ËgÎx0sKDã`uï†ÿ@Ý?¶úH,w²ÇxŽižcú9ð½G¢ŸŸ`“™~âMfú·¾Éˆ]Dªw”yhåôÓÊ`‚ñ:¸©g˜¶ÊRÍ4i`¼U§†¶àxºæ`8’ëí=Âû[R¬*€¨­5‹¢´Gá¸¡˜šÂÞ6Úè[=öúúãÚîëßŽ_›~ýi³ï7:Òbëo4êÚýëOð€þ4Üp»¸`ŒÕwO@=Öm ó±.dÂ{–·ÉøÔ¸ÓàP•¬+’"5+å-­Ãßšª»bØÃwd³Š³k+ÞàªLÓ»Ã¯•~ÊŠ¾ºAm!}PùÆ
|[¸¬ºpqµóMË¥i‘zèíï«)~ÐzëÙãh¡­p#EûInÕ—S¤Eiæx9Å'“i½Ù^dÕÅ"Á?¾l@“¨¹ææ•nÏéÛnôm N&gŒËtÞæ&Þ[-n§â´B÷<ñæ‰Ó´ŽP(ìÍ­¨ ú°^@™´¾¤‡"Ä½nÃd¸BÕ\ËÔxW}\µÌIy®JgÓö1Ø¾Žlì^&ûÇÇvB0ÙHsÁéððP$Í³4O¼\•¿½g)Æ`!OiX`Ôr>‘†ü"ßóÕ%ÿÀl–<š”tC>vM 0+vÓR7Þœ
±:Ö:¡dùIqžw%DlÄóQæOØÀÎiâô×sÄ;Ç“X;FøÊ<UøÖç£¦Ôžó š”›Ô`u8†N…¤H5%í'*´ÈÝ—î´"géšá¹9êh­^˜¬i"‡vŠ\Ìj\Š Çå¨YÏªu´‘ƒ
µ—n­ÐÚ9x–t ÙV2(Ýg
*GÇÃuh`±ö>€8£˜R[ƒ^ïÔÆ2ûÐ’_,i/‹‹ˆ
½GâWÇ‚NØ…GvŠ¿" ›æ^AÔvx)³V‘Êþ™ ÍÜ$ç€›‡afxáí´ÜôŸ)Üšèáƒ¸@ŒšŸ	@Zée·Ò‘¸79I«jôQî²eMÙî¬mÇMÄ¯Žµ¢Gê÷\„°	?¥Cì.¸)ã´þµà¥4Jtœ¼–úh§˜P#&æ{¯cIz}õÑaÒ^Ô7‹s0÷„_ubì›Uóv#2÷Åz˜tõ¯5Jäžó:™Yr\û46õiÅ^y]z¥¾o¡ï!¾^]™^¼Ñ4³)*óñ”ºñ{D”ó4öªxG,Ì¥£_“ó¥Uß‘/Ò±‡éŠhèÅQq>‹àÁ7ý*eö4Ä4ÕoÎódë2‚ùxò?"+Áás-‡NXq)©á'ç>Q©Oç›/X¡ÂÙÖ
S§ËD2ÍóW!®§®uþÔmÙùKN­m\=°¯ôZƒqZäô  ÑQ¥ÛQ{ÃEœ3$jñ­~Ë•6HéJªü
ÖÖt–*¹ë¡zÃchµÄÄëèŽŠz=äŸgevgàt©G8ÿa½ÛýP™a-’þÇœôÕJCÆ6›¦ynìþM£KÔ£aßô®zýƒ„žºÁGHèá>F×ýNû¡å¶
r³ÿs=hD‘?‡qŒÃã¢îˆ´jÆìíÇ<‹Rû%£UÔÇ®iGõ’?ÉNB–¡7[Ì ¡lýšêÕw'ÁfsìZ‘›-éE†ôÒÃÞ,ø›zTÄ—:\Ûödé’ð’Móëöo9‰Œá§!ÀOFJM{„7¥è‡öu”åol­¹ƒ\TvÂðì53^ÝY/ù
R3éiªi‰Ø¼PãòÔù$Yý"qØ¥¦õü´Êlô²ÚôÛæÀgŸÿàI@£“sžå5ùM ÷Õ£»qoj®ÑÊÛ!a‘fŠÝº3žß‹ÞP‹Œúàö;ûVímú×` vŸQÌ8Ç›$q:›èðüûÎ³K½øßn~[/Ï†­œºŒ’ËÂ¦‘aKMr|üÐ¤i„]Æ¨ê“Áä˜Û»!ÿßü¬>4©Û¨Z–NÃÒÀÇ‚ 6±ˆ§àÍZAí¯ztÛÙo€Òý çe4åWÔ'¼.½^[A`6í*èmŽGÇ²<ËµMj«D‡åÎàŸ6ä5¥¿ÖÕ!C¶-J|$‡Ò!/Â C¶…7z¡éa¯½–&öŸ6Íq‘Ÿ•f åVÔ5r®|üµ»ßjaO,: tË¸#mcDÂ·®=¡ïaQ^ûšöÛf-Ô±Y„Ã‡û^l!>ÍríÆ}¡?ÞÓýá'1Ü’@n.i£EÛC½ýpñæ{lÒni9·qóØ§Yý@}_ëíé¡‹Û–­UÐ#>G{ÔïÞÀ©Pw6ì<qC51Í¨ï.”*í÷XéYÕU(múéYËÒm‹„z¼ˆÛàô°KË-Þc;Wã:H°ŸÛÓÌu_ÃÎ­DXs!"{$yºÖÒÞiçÈwªw—Zøî{»R)+`¯ßÞü®q£@e·22ÓzØÞPƒÇú qàNþF!aîîÌiÁ ?sX3Ü²“ó+˜MùôïóïÉM6¿MûõÝž±XJøË§Eù{T±ïó«æ)¿W±Š²é&‹zn¦Åg]•ö÷†f_hN4=þT³°ëòM¶Œ©KMô„»aÙŸ~øçÿ#¡q°ÅAÌíO	¾õšdky©ˆ+Úk×‰µ1¥wZÆµŽèÔryM=ìüJ©W}î.ÝGµY¾;•¦Wß”§ÇNeä8CyEË¶_‘Ó†ë^–kÝètÅI8F@p°ìÊ¦SìVN9fEýBq“3
¡è¡ÇºèÇjiˆIzœ›€~ÝX74—Mc§²,ö+ˆ¿»—ÂøDªÜ^Ü¥äušßzuQ«on‚ÿµ‡àÓ»3Î¥<óÀ¶ëÅjI¸‘@EåI5^ëF¬?ñe+ÌXå»hÚíÐk¹bA[Z³Ð¦©·…6³g¹!¸Ê¨:½#x‡6ÿ!yH_så±S£auÂúòýì“íæM=ã";„ý}óê›ðüzØÌÕ5ÆqžÇrê È(š0xY„ év–Ç‰Þ°}Wù¨9‚{ùŒštÎ÷l;OAþL`7Hò±UÁUªè›o&¥}z‘ÄÔOp½´B~Äuâò›ÂQÅKn”ŠEã÷h0‚¿qYL‘®•>1Ðì¤ÀÕ6ô»¡HsX‹Q†þ"øÓu²²Hªq”ÁV¼2üÍ7~J1HåVe¹U¡ä"­çn©£90üâf¹ìÜ¯ ÚvÀNí^ƒkšÐM·0¹¦Æóz“­x\÷„œJ4Ù\éqQÃÁ6'"EéÕWyMž/º€KŸÿþ¥q£q]Ýºa#;ØßXtFÛ]”o-AÒø„¯ƒ„¿‰è‡ýž [è­®Èg£	¢¹·Û?|Íš/¤®qƒWOM¡Y˜µAEK(ñ5h<¾Ð'©éËÈFÑo˜o©³$|ð",Ù)&SŒf$úÞþòs³qT™ü…$¼òC	¨ºÙ¾fÃ'C5¢8Ó8c f9XY§šS:ï­ðÐ-®ííÛóºþaèõ©‘hçàéý@F[¨›bN£ ˆî>TÝÀl¤ \ö­¦¦²?£åÑRÄ{º¸)´S÷A•6Í°äûOÕ·„§~·‹ª³ÍhÚrèo£a«ÁÝO‹…JÃjîÍÌ@å3C¿¥ÆíÜ +êî!XÅM]#h6±Ý³Å‡Ç”º>ý£l‹ë¨ñïëzôùõ3—æg¶&?-^½|K‹úÝÈcâ|êgìn8
î=å“³Vs÷»Auèvï‘Äòe³{±NÖVÖ¾¯~X%ÂÇ~¬'Í­´î{WŸšh!€?	¦h&÷•‹2!~fŸä·„‚N0èO‚ˆ†UÊ}EEÛžægd”ß>š„ùoíV°C¾ìp¿î>—ºMÖß6ÞŸ£Àmæû4ÇÉˆ#ÒßŸRW5YÖ-H{…Fyß¾=.’
Í/¬$ÝAÍV“;š¶XƒÞ3ÙF¤–»³ž[°˜õk0"‚Œm‘Ónö
@%‰²C¡óáO žæ®u¸:Ü~›rè¦…î€Îs+æg%›sÝðÂo]Ñ’xÇdå³(ëÙÑÆýÏ<¦ÈìéLÂÄ#4EÇº—
Ç¤0Ÿ6SNšÇšÜÿ°™˜`ð$†Ù°æqšÇ].ëÒèäGÃ”Ýÿ9ê2žÃù¤3×±ÛbÒä¶Ä/f¸~§³Ù’£ÑŸ»ÙïdTõæÛXm”ÎñˆÉ,#ÂÈ™Øwº\ÂÇÞ^Ø­øh\w¸¯wˆ^y‡bèŽHÐCþÚ~±ºy©Í
éÓß1 Þ^žUo`Rd':¶éŽ±9_N‡ÚG°?C³Ïþ“ô„š¨±rx-<ië}·¾·]¸byÚ<YôíJßÚîk.X…\`´²ŸˆAÀ@}lf¨©°eŒO[Þ/Ó“Äoî¦u±£Ð*]ˆ‡9¹ÂA´ßËÇ.k¡O\Nûù“öRsyÇbÏ'Úiæ‘qÄÓŒ´ƒäck¼Åãh¾ŸÂû¨(Þ/’ƒ(?Îw7Y$I=v<»ÛOã•¹Tã¦~°7È"m«e^œ[TÙ>üjFžŽ÷K>Â8¨²{†ÞHÈ?Îò4/5ÌiF– š*ÞµDÞrá‡ÎŠš ‡‡Ñ!Ð•x¼Q1½ñã.—¾2‚Y4Åùb^:ÆX1”,Ø"½Ã(;Ã;ãž`ú¸­™Õ7®QiI¶ñROÄîô6ÓÃ"_Yê‡cŠ7ÝKjÎ‰Ös!dYÝqoÆ8‰jo,¯ÀlÈÅ¡ëLkÞÈU‹ãVOË3­¨EëJZa•›VzgñZBQ%þ‹ÇîtvOá¡+xÓÝ	¦iHIÜÓnE£N“1Æã£©Ì^² ›“êÄÓS&?£åMtÛ+ óªÌú!`*Àë4Çtòvq>,pÌPë"é½eQþ¾L–D1sèlÑ¶¥Ç"Ô¶¤¡Œcð…¤’¶©±°ÖYzžÞ¼’?mpSÃÞ‹Þ„)žø¨
B¢Z#$KÉ<y4B²nwKoŸ)§öÎIÔz‹ª³Æé°K[xpíŸm}ãI 8,ìãY>¦ÒßŽ)Ÿ&ç”9—ZäG8ü9-Œ+‰?ù}³'2ÓÌy¹g¦Îª¤¬ðG2™‰ªƒ¨'i¡¿ì×ÔHä z{šââúNþfS"ÖÌâƒëu":¹n!*^¿y¨ºlÙDø’Z$õìj6Ãúô’²,ÊÅ—³"ªA®“WÐûöòÐ1Ô#Á·³éGP'’“‡^@±¾Ñ…Ñë».(¹ËÒDË¨ ·%øÏi„Á¢ù€i
¹±Ž+<V>-b„Âw»Gû‡G½‡¯WBr”_’k,Fû	š¤U²©Šé\Ïà/ë L¼M)â
ÈëS…›žR+q1Òy9K0šBÐ@[×lQ2’ùZE×[4Ãç½ü  ¢¬¤ßCÂ…–—Éo“l
,¦Àq<á(	-/K/$‡CIÒL²= ÎÏŠ1³„ù‚fúñ:Áø™'çôg Vôº9â!Ôñáƒ÷>_¯\
ÿ²pÇT4çsLÔøY§Ô¡WY÷×€¢¯ôŒÂqtiåý”*“zVæ:;v½ôÅm@U×ï0×5«?Šã'ÑeE‡Ë¯f
 ®Å%å@ºg³Æ?….`q^zXM³´î÷–D‡Ó<âgñWÑÑ^Ï­	nNÏ^Ð.°j^¯¼‘7ŒôÕ7²DV½ßÖDÞ)lgXqXÒ	X˜·	N?h©Mv+âqý˜fi”sìÝ’øïÁ´Àó€sKHób6š ]/‹ß«Úï•7:ÊÈ„ó¨ËŠÝ¡‹Þ3Hm„4‹3>ô8Fô”æ–	ýúËx²¸ØÓGÉó+Äÿê÷),%\­™e6b½è¯z*Ð®ð¯®à®ìû°ªØUÿ»o¾ùâŠ¿2‰m™£ÀCØEÕe>V<Þ¦?û%pšÓ¤Ûƒâ©ài^ÕÏè<Jk‹,[õËzªó×
“¡©““¤Ü› _jÀœ:Ý„ñ³á &ÖE—bÑIV'»A©§d¸ŠÌ$¥
7ÓO‹7Y3±ÁÎy†JŒ-óHàÑJÄE¹‡¿%ÕÝÏØg¶{Q +ü*ZãÂÎvíþ4?.þÿ!ÿ’i\×¿$_\1 0¾¾PÊ¡ @À/È"ÐºÁA«°Èo‹’ýêõqHÑYŒëE;Æ ˜Bt\§gFGo©w™ëê ‘º OŸ|_á=à¤§QZ‘,Íß'ô«UeêxcüùòâÈºÊÇ¿¯{¾ï)Ûô-òÚ÷]9¨]'8ô™‚¾è!É ú©­nFÃã3dÏnSnôúÍ€vögfÃ?ÈLcÉ›ã‚…A$æ?‡œä)‰ƒs(“w_nŸ¥yý¢8‹Èv¡6ÐeÓN™Fqô%âÖ~öáO‹äK@^e2Äõ—È_M"Xdgò|ì¸HÉ˜DwÔáäHÔ¹‰€ø—õ/®pº¯±ú ÊqtV”¨,&Ó(?MHÎÍê(;M*­ªá;%ˆ0ã¡cÆÀ…«„D©U‰Sß%,<òŠ;ªR(UC' á~Y‘ðÓ²¸ä!¿lbJs\æ†2Ï£È[Yyf’Æ2D*6±–Ø5–EÆ·¾Ü¶{8ün÷ðhÿåÛW/öžl?ÙÅYž‰^°°gfÏt1Îá4<œ°^"2ê¶ET7Ïêb”WÏ;!éÎ"ö‰f¢.h6M2J\ôÄc¨
ÚW‚8Æ½{hu;ˆGÆpéÀšäàö¯k¾¯	J¶¦(Ù$ø¤EHÈã¹¹Š½”ÿ¡¢Ú5f³cÚ©ÉÓ0…=—R¦Æú\r"
ÏAºT{X)©×8®¿ûïæÍ¨¹º}"­?…]qÊòèÙqúG¼ZÓ#ØëÆe‘'e4‰<â ;t[/Â¶¯†+ñ¨ÜØ;#øÄhŸ‘®‹ÝÈ—E¸B$kgÿùîË½ígê—~"ÞZÌŸ4b¤£vx'mÙÀ»`«¯
ulµ.•0Fó{ŒCÙB^íK¤V/“j
™>üé,a¬™A»ÞQ¹iù©õÄC©[¨ô<…ô= µoAœÕÆ½NGˆà¸nÛùý»ýtºqyQVÈãhüþ„ž;±€WÎ«K¨ñ(“IŠaÛØ:aiÓ¢ª—PË€NÅ4j»ÇÇ´J?+â4uÖˆMom¸ qŽþ0¾•c ›6ÚÇ—¢cÛyü’wBIC†š•Ê¢ª~9Ëi´Àª‚¾Â’p¡Âa€ÜïaŽ·l\ouíÔ[Ú Rí³úòâü( V…1Îû*Ú²-².	]ÿ6 ä«|tì#6§ÍêtÕ±Ó‹BcÅû0P /˜RÈÃFlÝÊÀ·ôøÔÃÊÃùÔ»¸´´´„’Š˜²Jg£Ê'¢$Õ0d·öO@Ý'ÑåcZ€ªÕpV™nÌ•ÄÉÒªµgjÅ­Ê}«†¨Ö’wAR=üäÁúÌ»&BsÏù!Â³d2‚	M`R¬³I‹Ðé˜ÉA²Îâ`–ÿ-’ýgÏvwŽöö_w¶Ÿí¾x²ýroÿíöÎþîá¢ Ë\ÓáH}ÝY§nüÍŽk]@Psò×/_k¬	ÙãïÝûü¿Út]›c7µ9ó®­Ù±½‰IÇŽ_Ø†­âôÃkjiAêž2o¬²ƒ1h€"ÌRà¦ÄveèâHç·ë&ÐÂ‚#nÜXpõÌ•%¼iâR«S¨Ul5ah‰¯b&]±U<^ñU<íbl°`GáV<^ƒåù…]j-B¯	7)ü†‘o.Á×Š!P”hÆÊ.fz‘=·Åã3AqÓìóý:H•„t,eÁÊ!R¶XßMlî”=ÞE‡Û‹ÓâñŠÕM$Õ+Në šW¬6{2/Yõ‹Ö›¶ÞJ”ÏM‘T“e!X{xŠI7ª §7”®;7iªšîb({:QÏ®ÈÕYùO£³X)«ˆáÍDLöØH¢ý&c¼}CúIYÚ ÀáÃ”©Uˆ½§Qš%œþ	9t‰BDunòH†Øþ m4Ì¯®¿´%›µ!A{¸4Ë…N.àDÇèZšØÅFéfsïæ±_´yšNI«X#ëh“iüo*Ð<MÆ§CxOD³C?5ƒBE’ÓÉ™f[l©_š3ý%¡DYúc1ƒuƒd¿uba¡®Ã/©0Îž<ü›FeÄ
Û€ÎäX".³âÈ…_ø|\` Ï}œÂÇâ”øgø‰¬›b¡†·®Ì²?ïóôùÄû¼ÆÝ·yÐÀJjŽD£(½(¬†Í–lDr¡¡¦hØ(ýÔÌ€½ûùÌuÏ=mŸXÄ½=¸›ë¯•rN(Q«’µ¤Å¬îµÌ‹Tgkš¥1…3Z•¢<­’Ù¥-’×Ô¢—-€7ÍNíuEˆ>¥k‚YXj	ºåm¯‡_ˆ¢ù°³°˜¤·;A=£(4a‡níþó×('nPÓ£>­‰ö÷DC­;;vªi÷iWÂUZ¯ô”æÂT/ŒDÖ‡žæZÊÒ€VÙ]=¥¡dZmÇ1µð¢7Ø‹‘_Ýk±»-lA,j!×XŠ1q¦27S~Ä þ¤6¿ZÝÚ‚ÎFGÑˆÖ¼­Þz{qTŽŠ¨Œéüf°uD=^¯öÍ®þ,MÎ©Ë¥šÖþ½|5+ëõBµ,AÇ]4êEsäj§˜å¨¦7€B»±EúÌ®Ü<1ígt½dú-S£0²|ºmÒ:Yyh´û¸ÀÍ-£»Õ.3x>™¯]º<9×fü…x3P…Q2fp'¼è–uz0*[ä‰WÅã©4F,ÖíÕ¿((ãä×YTd–J‰uòª ì<}’Ô‡"6€ü30&*PæÙË«q™6åš¢=c	S£y¿šÁ°*aÕÇòžEYQªZ­Ñ\ûå8Š‹`ƒñ\÷%¿0æÇ¶wó÷ÏÈ…òÕé´‡ÈÐÛÐiOåoèë÷i•ÖbòóŠÞkë“Ùsººž&ó3b³Ü’KZy²„k}¥«Y|’âq}z–¤As.ö™•˜VIÜáÌ^î?ßÞŒ‡>¾}ùj[Y>,¨z¥%3.óðæøÍVï‚¬wûÉó½oŸoí¾äš@_¶§{/¶_ììî½ÜWRr «Í`Î]@ý¢åkÄ5—´q`Ø½-0Ý§žÁ=W6#¦+c ù]@¿ÖöóoŠGô¢óK:±!$ ñÔÇÐÏ!Ø¡šü?>~<h²úÐ.m©fômSóÜÀ¶K­„7°¬¾êyUº¯Ýdíìhçåû¯ž“&ÙùôYÏl)e´Û6_"ž­Ÿàö¾…k"y©¿ÓòRDÀKæ Z®Å·ZÙMÿI&·Ð¥C2÷Íº$‹7tJß Ûº¤å¿CZá†î˜{q[‡ŒÜówÉ(ÞÐ)‡[XðÖgåÓ+ðrymÃIÒ½DUn²…mu¹ç™Q¼d:Ú©K,ï;Ä
7tÇbx½5ylå°ÆmU¼ðÐg—n«…f³'ÜÇg·WD3ÚUY¬¸·#Ò¶ÜÃ{«ròÙòòò]ªò[rìÞ
ôcYH9²Ó4f iv›Í¾?=ðÜ©öŽÔ–t‘¼Özaè¸D¨^Ú½Mf*Eó©ñËËhŒ.	jròÑ(ÍÒúR]lªÈ1¬MjÜSáõ}´]Y„³¤‚aÇ“4_>¡ªè¿þ!x,žÆ«3‰»ë-™uÚ²eÃºxVœ'åNT%ýÁ0EÞ%Nª¾º*næ(FžÖA¹ª¹+y¨u«azR³wŒÑB”§ˆ¥nn™êP‡<ôÀlBWÕ2/T»EYb²î~WÝôÕ®iVeŽVkÇ¢IÍ1—[V–èÛEÝÕN|½Ô¬8u¼ öwR¾Óu¿ÆPL´Ý¤»¢–—WŽ×‘yÕü·>Ør°Æ¸-°Â}£sÔšô°±i–—ëKˆ6‰Õ&'Ì•«§O.Ì]m—qµŽGÑkuraVÄ×¼Äs%ØÉ´Y;QÇkûÛ+-l×@“"áâç„Ï“IaH”ÂZ½¦Þˆõ-…ëkuú0f!réýp#«Øv„\I žZ¦ ;æ€‚5»$êýÿ  ÿÿì}ËrW–Ø¾¿â²Bî(¨Âƒ%A jö"MRG°b¢*¤XUYYE Âà¼˜…ŽðÆîÕ„'¢#¼qÄ„·ü“ùû|Ï}?Î}d HµîL‹¨Ì¼ïsÏ9÷<¹3¹5¦Ç"z
» mé3ïŒþ»ª0ýb8\%bÃ!ùéÏÂ¼g b%*Ã†ådôûuXºgë:‘Ñ·X-J¶Ÿ›ÇŽ/¦ˆ!Þ×[µÎß¬ÏÉæÆåËê‡Ý½¿	þz=áœ½<›QÞéy9¬›Ñ®ôÐçRÍ=:Çó>{ûëæA1<–ûo;ˆ0Gƒ07(é%ïùÓ–ôõ¶EDœ±g¸’–W¢s*Üƒ.mB)-ÀŠE-<3qøÁÑZ9)›b<‚x¤=+ÆjéXæß¦Û0lµRÀYmH¯lÎÍZU‹$G$ùÑ´ù1ƒpÿ	ë¹“NbV­ß!`	3œè›ŠÒ$þX{žñˆÂ¯`-Äô_óØÂ¦n™½5~³ïÌ³²¤ó,Ìƒ© kJA©nëbÅ[[7s
v«ä€é`À[[c8ˆþµ¢'ûBH‹ä_|^­;—Á`Ð7ðà°ÛA·Óaó?/¤îÚ<B«ÃYü=ÏÒÎs¶k,v8!Â2(Y¬HãÌ• K€g÷ËÃb1žË«„i8&c­œATB8ÑæÔµ&Î?!Ò0õ¥%c¼ðntöÃ¡_®Ø ò-»¢j¥;QßÝåºª-äˆwc}ëvãK0œŠBl‘ÑÿÒè  yÓÕLq›ê ×Ãz0‰^†0MW´%h‰^lç	aõ‡ìb&`(ZšP–åëd¹§ëÐË9ö‘£»qƒ,¹dÆ$”+šÌúª½Ç¯jï/fcðú+ååº¥—4m_%,Öó*ÃöÍ¾hqz~ê°×–™¨gåhÕ6cµª[+K¿o˜ëZ"	cšæV„šë÷þôþ¯¤<S;Óó¼ÔKxÆÔ	%_£‹àsÚRÈõº4º¶hvîµSÌ³oõl,JÄWQ?2F«¦-ŒÅà°ù¶°<¬oèÙy‰&Ö×É]‘^È æÂœ	U@žBèFhì	å8;˜¦ÿðþ+„þÁÞ4ea
Iá²Æƒ«k Š+%—RBFÄÖTiéxüãci7hz`]ã°T¬¯}-0qgÁ?ïuõkn7ìj†4	õ3s;q>6{QŠØlõðyy”–ÁXÛÀã-¹Š-ƒÄÙFx”Êz Ee¶f`#`¥j^®¾ø´³;–pù,y‘áOêˆ9øEÖy¢|Ôt$Ã\˜Ê<BÌ
6Ùs>0¦›@ÕB4°—Å´B4/¾éŽ3W›šÈÚ|+šò5ûi¾7Õ©>Ê6¿TL~&˜ß˜jQù™~fóÜ–²"y`C6D²"þÞdÄmY§-ü¶C»†|v„â¨ö:¤©FõÑˆâ9¤cþÐšäˆN7®5t:›~…xÝ †-¦ëŠj±âÚ©”Ò)¢IÊÒ%5?!EuDeŒk€|i^ÍbËjêÃïk«†í?°=®,ç¨§ô^É\µÿøüÑËúe”·qy8…a¡;ˆ ÔòˆïœQóy2 D
²ŽŽÆå3áºxw>=–úÂÉªÉ0ˆ¶Í­i‹X6ßëèˆðÂ¸–ú×Ïá¢AvÒƒ•…oW€k¥ÜI…	)Oç?0™å¯­WÆˆ^ÛÕÔMÀè×º {é}±šÝ‘#Ýåñ‡ÍNüùíøÃò[VHtéÇ£¾O±J56Ž 5<K,¨`³ì6¬ÉÚ–¦ixãw|†ÇnÛf/œeu=;Ä‡ðÎÏ7·=gD8ªtoÙÏ/ÌŸ>g©x¿àU£³#Þ_›¹óNª&o|7‘Ò4HucæÌ¾+Ê5‰ôN=¯I©œ™F”s2ê‚œR}<ª™ ?ÎÐëž3¤*ó°
%ÂÁj†u’Æ3|‰%ž‘)€½²Qp²/LpŸ?ûÂ°Â…x§\‹ž–;O¶c½n#rÀ°NÐõ1<å¨Èi‚‹‡=¨Uûm_¬¡mMC¡ò\®.…ÚæÅŠ[‘-/R?÷*]è?‡—ñclÒŽ³¢¾í†‚úûÌ^
Ž_q$\	å /”`¹4òÏ==œAVÜ#3-O¤ì;I¯r$S­úGÙiõ–Kk‚ð½5ô¦0/ÐòÙ¹jîâÍ¥v‰>½3œé»_‚
Gï‰FNŽñ®Hp"ÜO)¤Q(ý© å),G25)cÛ‡pßJcd9~óM±>ó7‰îNäf.©ÄÀoY:9…ãéGÈ3ðÁé¬f!É|ÝtÉ^	#'KíÈ8ÊÒI]ù@™ÕH0ê±&MßÿÞ™z;®RßþŒ^&ÍO t€xÿX˜3+‹ùâ)¿[r‡wþÞ¿§ËƒØÁ¬¦ÀÌj„*Iî(_£—õƒÓa9îë[÷D‰~do=/¦?ßG¹JÃÄ²1Óî_>yüR¾=—€´ö¼}9“™Ð~·ûƒW¯ßxè¬XÇ)žæ£géXÔAšÈ©Xšeû<¨üBÐoê¯«@›Þ–g­Žsî¥Ñáui(Y’¡…ZÐN(¨1}>Ëðæd‰(%¤mÒ&Eãh2áþ[6¢·ˆ®€¶m?µY\Ç%˜yýË>ÿ{èô•üýÚ¦ÌÊgÃúÆlÎô'ßªÚ3ýµ+A=`øâÑ»æœQdå¯!“'øTëd Öüøî²–(Pó*€cøê­¸*º@Bqôu¸OØ·{ÍýÊU{OÓ;Ö_õ†úÌÒ¡©ãÌÆI
ýÄ¦ðoý¶÷Z›9@¡Æäì8ÌÑ^Ý‡nÆeÔå%m9R3vP‰+µ7çGû3Ö
ÈóP“g;¥¢Û‰AÇÍ”Ý•Óni:Y¯x­Y>Ø9í•S66	ñ²HÍA5b­©cÍÁ—U
Èµ7	èì9|"U§½Ã¦x;î1°(Ú*8'Ëúò×ÿ‰œ}BüA0YjörØU×‡ôïý
q-º„ŒèBÃì£ø"š2Þœ)˜t›MUßøÅßv˜d~M˜r0š6Îœý.¿vª~x5)D1õ5€`8ºgV_[oÚ·Õlæ¾a‚h¼dCzQÎMlkZ°KbÄ–QL©ÉtØ2Fö¼(E¦ºfžêÐÿI	gájÝùk\#²œËà¸hûo>;×-BZ*·u3áEøTÒ)¸œ„¥‡ì¤|äƒvY9JÚbçQµËøƒÅ*¸:p(bÛ/çÿåÉÄdl’=]‘²TôyIoÐ;ž„*nŠõF=¸ ÕÑtE;àÎ/Þ jÉn÷Ñu#Ç¹ÙœµeìæÓmgÅ°\;[»­réž¯N^ä»â]uÄušŸ¯Ëer«ŽËSvh×À6­l eÝÚ&d¿ùné™îvëtLÚcJÃNÖÚ	š§xscƒœ¬Vs#³¯ŸËÜTªá[™°ÔŸÑ7ÂbX9Yõ¨Ï‡Sû$¦±	–o³Ëƒm3i¯H¥l$ïµSõ:XÙöÁ‰¾¡ƒö™ñ&–7X¬ÛøHþÅ>ÚÚðR SP3²ÚoÓVTbaùÈfMõ™¶C»ß½s5ó›¤­~¦k¶ùå…×y—%²¿/§â&¶÷³ñvÞHíäcÝD>ºwW`nÛ;ø¢Ñ'):Im¡Hq*ñD¬,ÅIœƒ}uÔT#ÿY£ä¦¥h¤ìèŸ[d|dü¼INÇÆÏÛlëo9ùÁwõ¹GÇç¬ø¼ši‡/™M­^wcDyæÁ0`†¶nìÔŠò wÎ¥”Ã}ËÌ³ï(øØv·~½ÓhM“òØ@•Mw`°ë^ÕÇåVxÀÒ(ý’cÖÆë±£óÀx)×ý6<Pn÷~Éar@`Víi`Ÿác…|Þ+=ÖÙ¢¡;sÙÁr¯”oKPÂå—Û„ýl¼ˆŒ»¢ÛsT_vÜÒJ;{ä–7>ö‡ô^ñ’âÎðØ›ºM­¸åØ“f²PÑÃ=·]Üã´a0H·-)Ä)Æ©ç%£?¾i¶ÌˆØwš”M”/i[“ƒµ[$@.½æi<ÆŠAL¼î$Bs ŠbPD%1«;õã›Þ¤4ã&2>%¨3…÷±húmyF¡‹‡XÄúÙDúÁÆÅVñ§E;¯!Zñü¤,§|¹O[¾m9©`
bì-)ß-`"<ö`;Ø¢ß&pK½
›ÇôÜÀå ô|¬\`›§×˜¯Uhë?C€G¨‘'—‹óé…†ïê–ìjÊb±F:£Wós DÒáæâ?¬„{ŒO=òV¥ŒG^9`vÂ6‹(0‚»•ØÄÃ1EÇÕˆÒÙ0ÀÑÆƒÓ‰pØvéó^­!DYh»°²Þåv~FiÈù99©FîÝÖv½!¡Ú nFàÅŠ7<äCŸ–|â”àqqV/æêÒ·,IxYÍ¸Ý„û²ž‘m÷Œ]]`–-½j–ýUÚ§ó
1ÿ(DoOŽó áùFÃó¿w4,ñ¬qoŒâÜ.Hsþ#ÍàëçÊEXÁ%4Ž,"&r`Gþœ<™ŒvØßM}[Ò=Ï¨¸!)+å¨äA^nÝÙÚM,!ÒUDÈ„àr)§H|ëÂ ¢û¸‹îíîñ–wàaL(íù
!ÏÆÅ´Ö„Áïdýxéz@5Ûr©b	…%‹V%‹¨$äû%a^Ë@Ÿ
RþeQÍJd³Œ•Â!î¤)f9D8Uv2åït3&N×=Ñ¬,¾ˆV“–ñ¢1¤‘ÁFöfæ¤ ›Å;êÈÖRîKÑeáÂXKüÚNˆæNnûÂWYl!ì-%„u/6Ft9„Ðƒœ8HF<©‡-Â•Ep]ß‚Ø! Ðz2wõfÉfÁÑf#ø7ø1¦mpDV/>	l-Ž »íïµÎŸ˜}OÁÈÔ®Ü4ÑŒÆùJwÁê~u-(¬MTî!éå–»¬ÈzSR¶|IKÈóA3jƒ3‰Á•å[¡ían1”µFáé9«¶é¯š¯¼.š¶nÖf5dÿj0ÿýŒEÙˆƒ[GBÏâšw+°¾ÄN½hQ’TKìÜÃa9›ßéNÇíé*°Ï(”2cÏ;ç¦5hõ¯í}„p4lß:Ó68uN4ùuDyËoc+2ÎÆýúdš|Æ!XÀGˆƒ°Þ¨äf–Á€'Ž§Ìº©´¤óÛAV=ã|ƒî‰/ø«ÍÙékœ·b²˜ÁÅ
D/“Ñ<+[Ê_¶E“	£Îš5%í¯zWb‡y‹Ÿ™â ­ÇºæãòpNj^Snq}‹¬1cs9cìÉagVRú &Ñô†õ1ÓSqL»lî0CFR/¸¥Þ`0À*H’
7‰ámú¥¤Š/T¾2!g%8!)˜Ñí¦‡½A;Ä[ú–B!j¸hwÀRÖ3~¨«ê¶Ÿs…\L~ ¹}ïÑ­l†•å}ÈG‰Ÿž×kJì¶ÆÒß2lË½HéGöCï×¼ãÈñªgÌj‹/‹Æ¹÷²G/úÿtÁÛÝuþ"1eàÓEXb*údBS¨åBnÿ…ËèËV)<HLˆÄaâc…r3$¦9¾eàÛŠ¡šÛVz®_\ƒè¦•Kƒí»MSœ ¸4ÁÒð×Fd€½J²¡R |S×ã²˜®¬ø·DLNÍ¢o§Ë=_úlÎœ,\”^ŸOñ`q…ÿ•œ+3²sÞ±2³ø9§JFhÿä'|¦aÐq*ÛN={ÖÏüV,?›=óWv†ãÏža/—__9ÝìÉ?Ãµ/wVr”†Ö1v,®æä¡sÓÍR¡'b#/Œ7Å¯N¸kKMc´”\zò,7\‡,~ŠI™2ðµ•T2üm@[æª"Ä‚³¿›’Qz28øÍ¸š¢rÇ?®&3ïæ%,ÎÑ¾¯gürÎøþ55W9Z[ó‰ð'®µãGC¥\&U\z³­ã›·Õö¹ú•l´/Û:G•T©Ùƒ¨Y©%ÇòÞ1+SÚ·•@Ç†Eìî¤æÑdP[„1S<!/ªBm,«Êúe©ŽdÞvË €áåwR‡0—æ@sfÕÀ' ¨âà)-éApiÖ¯oFtUöÄLYìQ ¢Ù1·ŒÝJ3ùH§n­Óé^™½µèTØ^GûdŠ9WÂ›«…CUkùŒ£^Ì".bÙy¦Õ¤¥=ÙGæÂ²ùíÁ3h¦¡ì[)”ÂC>ˆè_¡\®uç®Y×üg¢ãËÀ¶îYÁ6çìàW²_ß¦^ÏRV‡à!`
´,sM{p(û4ßø©š
Y?Æv(™L;/R« e!Ö]ÔÁâ·b…ÇI0¢¤ÇÌ•íˆ^PôtUÇÀ¨gÄñÁ‹B ªx¡}ýbFµðè¼~!°Û'AY/ÅGi„)p›ŽW”± >\aÄ÷b>-—EÌ†Â‰¹mÄ;$äfüƒÄL‰.°Rø™/Â(±²4‘¬´ãÅžÀ¦€¦ÒHL"'alrxÚØå&‚¤NÕáañd©x3©ªXQ=å¤QÅJ0ÕH¤·Ü«XÙG¹‚m$3IM#™”5=	$5ërSpÂZ+IÚšhæ;“„%R·&š³ý¶Ð®ÉÆD|m¬94+Vœ€Ü¢­Œd®XA‚w‹“)]3šc‘–i_±Èµé(/×5Vö‘&® þƒÉ|"“ÈÍF8KörQAÙ‹h®¿\¤xW<a/ÒPØ÷Ærpw\‹²MÃ›Ä7ÑôöØÙ~}KšÝŽ¸”†¥Ôâ}·k„ûT²xÕ(âBˆ­»ÃY‹$dWºD/›¢=¾‚5J9ßäJC´ûð+ßqSMß®á~Â:ð’Ò\ê=ÃÇÙÙÚ†¯…Iµü›îÁð-ÈpO*-ú<”r°;ÿÝKÄL[SÌ°íJˆ%]ÛÚØX¿”nxkâ§Ñœx"Ú»"©ÀyŠ~†,Â®sç’Ðç±&QïÀ« }b-¡Œ£#`ùâ†¼ëc¾õß,ª1àhu†Ãâ•Ä"-ãNîzäÈ#>Á‹…¼²˜j)0y9‘i`ä˜¿.«×rû".Úß~˜{D¿1ÂA¡ü+üû×A9ü+ì•Ì+ÌbÀ¡>îkÈ£¦9ÖâÒIÐq¾o°³³Òï"Ð÷Ìå±e	ÿÚÂó*^Bxé
'¼t„ÙUG˜‘}¡Ã‹°2cHRN#f`¼ÄH//L=G›
³?PR®+_
ÿA†‡¨š‚¥Âš‡9c„/æäM÷?„RÔ”£Ð‡™®éšawèŽb³Ýy-ž{Û>ÈAÝÒ“bö¬šJŠ‚ùâŠÙBHTÌ¤Â·ÿ|VÐ/Gf†7¸ð&tbÝ—–0uˆ#¿½geóþoõ(à¿e-úÊ÷–‘Ü¡Åðèù!Óø³ðÀ`å âýzš‚•hè„sãû‡Õ„]xý¬VªU»¼Y#Ÿ¥‡B›^¹xbø<çfc1ú¾þ9yRÎæ1-"zè¡Ö—Ò!,®Í©·éØZ³Ÿ€D9¼Šh]è-h¥–´æØÁ(YÀ•ÈY‹F‹ˆð¤yuð”1+¤h$Ø`’Ò¤Ð4>ÌeN”s&èuf#–$$$²ÆÅú§Õ¿Üè“rM{#×¬ô¼lã9Or”@MÏ¯a¸ƒLûºiØGÐ#„4›Ç@—a{Ñ¾xÉ¨|t™ØB>\Éá}…¬s°?¬1A6“ÚðÔiq©M'	öé³ujž÷}Ì˜ññ 4èxòi-\L’ÅÅe™;~Â*ã—ƒ‰f?VH±µ]a%WM@¡…bVx~­£Vâ›‘1°$*[cbÜpw>;”º^„¼»;ÆDëHqã¢b.Òƒ–2ÓÅR®íZü8ŠW“R'€³Ï¸8½TôÄ9‹SÍ‹q5ŒüïÊéñb¢z–lU"ÊñøÊ¬ÀQ
×Äú‹í˜	ªèª1YËì{z$SÖgnöÌ§?AF/Ë?Ë£åí?Xq¬ ½'¬ƒbZBj„Û9‰YrB1VœL”ûƒW335q¬ÃQf
RÚfb@Ðeôž…ÔK¸«RìZ	¿°b%½—h>Y˜4†§A±µŠvÂPü
7„Ñ«7g§?nÑÿ5GEc•ýß`ckå5Ñg 1–¸ÀMÎ=.Ðrä8Ñc—ßª“9‘!:YLJìv“õa‡ˆ¶vé¬Ý&Ç†cS=…IütH HJ™š2ÔEïìI{ì/ÓæÉfáÊJ4˜ÍvÄXÙ-%¼YòöâwgOfßÄâ,æ@‡u ìÙð¸hîÎû)>Î©6¯¿†ƒçïË®J×}?[Õ,Q‘…óiT•´Ôkµ«<òRj1ÇÝKñ²Ô–_rÓYsÁå<ç	iCr¿¤¤0Î×8+û•òuâæNÊQE™°K.¤³<Žç©º†9fƒiøÎâ}™°2‹m+”9¿y}t4VYâïÎçàQ1Jë¢UÆÄäê¼%]ŠNY×PÔQžcvÆ¬x>ŒarU$™2œX”öÖ/—‹ü»S±ûÂ4Ï’‘4Já'kHÈµÌÉfÑ¹<è;×¼}ìbg—]–ËçÅ_ESJ=âÍ¨eŠYâW1§§¥;É¤j]°§µV½gE3¯†Õ¬^0÷’»xQfSÓÜŽSÖ‹æ—yÍ˜FÄrÊjÙ!è“¹Ô<¿ÚÜdè ˆëÆ\þIºÊ’U ërmŒwýsò€_Ï»`¡Ä…Gé¡_ÍÕƒµŒšÃÈû.›O9ŸGEKfi
5n,Àr¾v~9ï¿’wåÞóg_ô^C(v}çf&8 >½}'mQ©t Ìºð€¼j—;
v?§ƒÍå\dYŠ-Ðe1ƒø’Abò>’t[3^c±DÍ´D	/tÑ–©&AÌì½íÞ@·KT.û¤Ë’~Yz˜ÝédLi]ô—d“t‘k+D_<íïd¶Ä.%sQ¶“®'J„}º}öI—L.A–®'·;ªÈ§öº$<­Oó¯\ù·ýì‹'å„£Ì/Ì<ÃøßUW³.ä¬? yÇoaÝhý­Ur{•|¹J67\Jß?n:Sïeh÷ò”›Ñíc$vP¬\Šh_1É^–`/G®]ÊP¤=ÙüÂ™CòOœ;ì>ˆã„—³_ºQûnÀp)Jt~“‹K~JoñPŒÊ§½¤Ý²oº]‘¿nß‰Àw;ã€Ob¾néNÚWps³@ó¡Pç¡Â±iÚfÅ-“jz§·Ù©FqJktÓ‰Ðo6ì2§×.;ê…†‹káÝÂ•Ñt(I½·[œ ­Ì<¨ûéØ'Kœ¨ïØ–ºAïâjw·\)ùZŽx-Cº|ÂuI2ô.	Ç-Ö9âZí+Fá&Ëð±))J(0¡Ai¼ ¡¸YkÏ”± dX/æÌgZOK3¶á¦ù#”°;^¬tOó&K—Lqð27š+û0I„2Ú‰ÂÕE¤‡hÛz]íŸ­@G¿ÿ=1/å²è˜Dß4L¢Ã^1V¯ì^æ8º%L<r|ÜÌd'íŒÞf×Z®¹0Ìr\#‹ó	a‘z….G½˜…F|v«ÉQºÚfxçV+©Åx~çüÍCúù„ül†|{…¡UE–Ñúà'X¦!¬Jü¼7åaÙ4”Õô>vFy–zM>ŠUŒ¢†Ý%/Ç´ƒœ‰üFtYàãõ²~š%–†Óç;½ÆÅômjÆ0ñzVN)”LëœÀSðTÓ¶œ¯mð$ŸÅðíú­RÏŠa5§´„CÙç»åS'+˜xœ6PÒW„ø(eF»âgZIfŒZÞÅÓM:‰I(*¬7–(þ¼Zªcèå£Ê(ND¯È’§|°<—i¨ÜÁ/04ÛNîà)| ÁÓ2š-Ú[Â°…{`L7-”åŠ¸½Fp}ÁƒÑ–ð…ð®Fœåçá,ye’‘š*SB¢ÂkCGcÍå(j›ŒjA³hÀ.†ó ¶T„XüiVgw#xdVc}#a'û7ìºšéQ»–FÙv'Ž0Å³8ù"ÆL_*°­oEV”ÛËÄ¸5"Ü†ÆðöªDI(x0F~6’X6¯¡e’ ÍàIIº5ä%‰5€§&áotüw÷­OFíàò–ÿ	J]ªH¯k 6QØfkÃD&(!qÛ8JÇð4†Yž‘öät­XP~™y y3¨=rC*ß
±/aw¬P’ò¤ëwº)ˆÐË•SŠÍ(.¥¤gFQ/$è}¡ç‘bcÚ
ssAG†Xt¸ÎÆ{ßø¶bî.©K•Xh%¶è”ÚÓO~>/Æ%r{áaÊCeæDAm\ÌÐ$6»ócÊÀ [4oB>k¢Ùßp*jð¼L¨ÃÆcûhò;ÃBõö„çò:a!v×çÇ¹MŽ1‚ât«ªkt©%³Eå×±8
5÷FÏý‡GNŸ7X6yw~PÎ°†ºfÎà­…t6éPL`CŒBýd¾#zs1ÕÅ	rãŸ¼ŒŠd"±¦²¢ˆñ’Š%&ú‹‹àÜá ÷•íh"ñø˜<[©k§È<¤@Œ¢)öîRÛÖ-:×eãs-¡kÉ]KGéJÄéZ"RÉˆÕ•­«c¼.S¬ =âÌzJhžÓ•cÕu¥Ç!N—ãMòC9%%KÇq"ä¹t(§«Ä!Á
/qñ
/2H\•¨…—Üºb—è°òE2KŽ¦ƒUKOøúÐ#wø‘Çà5¼,/®áåRB^.ïözÉTDzh:¢,¿ÛË¥$Ò#@ÒeúýÆe5Q0ˆ`Â(ËtU¢´¹e6hcËˆmð†:
oD#AŽ|¼bL–¢Üùã¼t†ei‡ÌPP–É%êåfˆ‚Ò1K”KdŠ’Õ;g‹‚Ò9c”Kd‚’—9
JÇìQP< £q¸„¬*]3QAY:™‘*ct¡¤UUSy,®ú+™<kc•ÜÌ`‡,WP.é
ÊRÙ® ¤³‚@é”õ
Êd¾‚²dö+(y»¢,X¼©e2aA¹t6,(2b½^&+”%3cAÉÛ¢² \:K”¥2eAÉ›ÒeÌršºlÖ,§9ž9+Å‘ÜZh¡1ùÌ=BÄaF)a/;y´Éýµä‚Èì`Ë-I†t—¥e„¼d.Ë•f0ƒâg1Ë_æ¥òœ!æä:wÛåuÈyft¸|Þ3(Kå>ƒ’‡Ñ:å@ƒryÐ ,™JÞÄ–Î‰%žJ‡Üh¬ÁlûW0aFÅé·íÔhJó¦s7çiÞd•pg-.‘+#?Åû:Ã¿i‰»yçtkPüEoÊ‘ÎÕTŽ‚Kï.±â§p+»Ì’ç§lcŸ'×|I=®ë©TºZë°ðø«Pù³bzç|ûÂÚÍ,KÑÌÇaÏ#“kÕ#>O¨£—¶ðáåŠì|äXÃF¾ÝÌ‡/ú)fIAƒ‰ŽýØ7 ²lwÏ+Ø}¸ô¤‘Õi92å•É„á?(†oGM=Ý€]ÑÏ€]’Ðƒ	Dwµ]œ7õjJQP1¾s~.MôwÈÆ*PÈ”ÇÛ|µ‘ŸbZMèX­Z›ªÖ&VÅVør*að4)N×NÀ<1¨ÚB¼ZìhPûxíÕWïŽ_»h±™q6a¶v;b•Î Í,‰å«C¬å¶‡RÕ•¦d$‚œìŠ\uX¿“9!P÷ÕP:»¹Ìa›ÌÌa1nÃÜŒ'äžÒÝ|àyÂ8É7&šY™¾0Š@ÿ)åÖ&Š¬ºšo¤¡zúbq0©æwÎ9ÂaSw@VF8»¥Ç'ì\ ÎÓ!{ÔÀÐkd9¶-‹1Ÿ4ŒK±ë’N^²NxÆRòy ³AsFoÊ¿P)x™Êý©TœD·²Õé¶*$ÝªƒU¡xq|¶/Bºà;éå À²¤‚so)h0†‚å$ûàt‡|3®æ?Ó¡“{×55©rÃM)Jœ/˜´268fP ÅdÊšXâ¶ ƒp-LÈËjVóCRöCqŽÓÆ(´Z—ð-¥ƒH-¢Ö‰‹%.r`üæ%ýú'•º8žn<ø2+3ö%NŒé©ˆÕñ¡‰0Ù¶6]¬‡°ÎÐ£ÛÕÅîñSÕãJøe“ö\Óa“Á9¨¦ÃþÓÇ>!¨¸êSç	óO^TûÏ#€ÐÑ@šç”ÄÄBª]rGŠ5[9¾6#‡
ŠIÿˆ|l Që˜Ø ®‘øbâô(q9¯\®}ñ#ü"Aa‹À+%d{ªmz/›rÈ„†ŒXS‰&vÜïãh1¼!ÁáU#F'´‡ÛýÒh±·'…	}f×ðŽ#f%!Ï½cYI•üJ‹8 W‡Épt7ô>v×“…ëZò¯h+)cC‘YdÐq´ø02¦
ÌEFe2tàÀ]‹ñ<Wã™ˆLÛRfãžûWBŒÀ®.’ˆŸI›òRà””]yÚÐ.–z)Â~ñ©qÞ ¦7÷«óe™îx,3¹L£Ã‚öÄ%a!´ÝÜÇËµ›æ~63ò‘òæ¡
¥ÿ¸[ËàSñ©Á1óÇÄÌèô7Öï®	‚ÿX7ïÿÚT!C©«€ßy5‰À¯žBYÿñ‚§².ø»Ï(ò!¥ÐÌÿì—9sŸÚk:‹kë'&u~Öïÿ¥à2çe„ËÈ‹`é¸Uq0ƒ¾Š¦9x ”ŽHµ$<ÅnV™ØÎôÈÈEw¿„BÜaÐ$on€£q¼>%¡'ŠÎðý_É!¡í_‹2DFh]Â®Š­Xd<¼ÿCKž—íŒ^Þßÿõ]Y-ÚOèµ¿z½ò«q{çÜ’0Ü¯ ß
zåwz™‡2!¯Êü¢\Ž~ xk#î¼¹FnP€ýh‘Š;xÅ¤áï>y¬·ùDÈ@rAÊÚ„ŒL	¨YïY¦é;$¸«Bø¼ØskZì$@!ŒU"R${¨ »cÜ†¹E5Z‰¤læ­VM;™~”»T=šAl÷Í¯6^3	ujô¬?X“¬ùâ¡}²Wº×`cãrN¡÷t.t3L}Q—„Çï4ÄŒ"MÁûÈ°@,kk%CýR^B‡Â5(Rbu½´%)[†G.rÃ„.uLPñkL„_>â@RïÅNÇk„¯7‘ÔWÎÆ8¾;âpgÖàº'u@3kqOdWä>+ö‰ŽUF4lÙpŒêÙœMŽ±z¡½ê®,LúÌ•`Qj­ÝÙWŒÇG¹Ó{Yƒ¯[5TÏg3¥úëïÓæI9/ÀF:°*Ë®k”¥²‰°Ü!‘ÞåÉpýx%Fœ‹—Ë²ŸóQ(ñÇ'&¸BdÛŸ¶VDítýÉE5…û0f»Åÿv—8Ž[A¨YRè¼4µ%¶A¹Žfº˜0éUwx¥ˆáv1ü Žìä~š‚<£7Y0™«~ÿøþóÏVÎ{$’í
ÏT9… n2}ÜW*Ù¶ÄçCŸÌ!oPÔG¨ÿ‚	Î†Å˜õÕ„°ÏÖÉ­ãUr›þïËc°9ÞÜ8Cm<FßÇ€æÚy9£6"Í<Th…UøxQa§ð¿¡BÏl®NÜÜ„MŒ?6%=mÞÿö—õœ¢ÀB¿ÊÃËÃ‡|äG7/ÀÊo÷2÷ùg„²ð«WwxÃ<$¿~QÒ­­dˆ£®*ÎžüÚå„®åñI8|õM7	í¥Îýð¸¾=¨OÃûÄ¾(Gæ¹u‚?}¼g×‹_¥N˜ÕuNp	#E£ˆ>±AN(ðÓÿ9ÛíÄIÂ¤Ê£ŠR©/NôÏøÊš§¦l÷—ä>Cƒ³£º+–9:°»,Có/æMu°¨#ö™JOû¥å0§ É0ú6||Â|”C/·$R™{Øwý6EDsŠ!¤×k,<gpÉƒÂÎÜè¬dó¦¾@>© ×Q¦&Ô’Ak¨!À>.?çz‹M]©ÿb?ª*Ê…@Ò!eh}ôþ6*ZˆíP“iÙÎ™=Ä»joê	]Ó*áðB¯x3Š8†Õ¬‡c@¤ËÏäâÔwŽ€Ü„Žµ³M:®D2à1…ñLd7Î
Ë5‘8/ãšc^Ì³‡EÌvÊ€"êít„¤dŒ#á1’î9'b/œ´óÏsB¡2OsZ9É;îåÅØ»ìDÞWºy¼tÛB^2ÓIóòpñ æesÍìb?=×´õ‡[úÕˆmP5bÄ–™zÚ?J–¼à¶²˜Œ\`Â]R~/7µ×yûÚ1€4/ùœª.@©¤òšI‡Vc_åœó‰	d‘7Çp"Æ=bNlïq–çË&X ÌdòšfdH#éÊú—:#<Q×]ö+šÒ.<ä>wõ"é ¹žò#òË’LsžNû%>/¼Ýbæ*-wP(ìmÓËÔvÇ`bÓËsèd‚”Î Å /;õh.$A‰$pØÎ( %o›·G‘1XÒëvÑÜ÷7Vóú{¸@Ü£ˆ~fŸ‰lÖ§<—Õy2×=á·‘y	¿1˜žÞy“ÑÓhêqÂUVŽ?o¦9"¡¤âmF70±˜¿Ý`=´N\Ì÷”™IcHv×,–» ¢gØ’âŽ>}:ƒ–£ëûé0T!àñ­ì«øVjpû„¬ç%X2²ˆé:®T"6Èîúñ­°Rdi¬¤ êšŒ?yÁ}I"nB›’”«æjTÒ:ÌJ‡'¾²ÔV8Û.m§@ÎËéN2µ'ËéO1ò²y]³ÅÖ£ÄwÏ
˜X³}òæîb^OÞÿ•bÊz‡|vv€îÑñÏ/Þ$ l	ÓÆFá;¼„‘€¬N˜ÇˆôJJ­2R²,®’ñU §NkÒR¢RNŠâ®Ôe$uKÐÈ_½JóÖOÁZÒ?~Û1°üoHö@²âuB³A»\(¿^,k-ÕÆ³KšR^½gjXöòeëó÷ÿ“Ü$ß?Ü†ý˜+ªÃë§u‹:e¹Ë¸ŽþÌI œÄþ9™“­óú†În=°ôŒÑªäd3œ>Š 0è.T ¡áÊpú!wÈ+KüÍV*!‹†Š¯èð_w
€˜~ ’Ò±AÅÅÆQqñuÛ>QU( Øñ5"
6Îr´,ì4>HÔé	‘íÁ’Ü4—DŸm;ÛÈvî‹b¬b½ä_2-I~Ðn?Hú‹bü‚¤ƒbIÄb`¡Òïñ,	ÑhéáøÞÉÛ}¾»®#ìÛï°lú|!ª)¸^üîw‡‹)ŸÌ×/õ‡ª<aü:.„]Iáì±”qÅ°.õ¯zú².Ú9ÿóÛrJWCf|†gÌ8G½~Ú²îºXÿx\ÌÛ»³<9¨ç÷êéauÄ>žVY”ßOYˆ_]kÖÔ‡Õ¸\ýÝÅ‘£Ü‘³“xÅøþað'lðô‰ÌfQÕw‡EÍß‰©ìþd‡´,{ø*™Ó)‡eÛ²XÂeÓÔM¡­wu5ú™ø‹}
ÃñÆeÕã‹£?”~Ì‡ÃÖŠ¾ñg|Îrœæ²Ñ/ú³ãzZê1OèX‹#õ€õ	–#U[îBß{_›KMøFþýµ·ìû²Þ·X|{Xž¸¯ôtŸŽ^–ãò¨)&ßQÈ=¬†ƒ´;¤hÏ¦CAÆæâ“?²ØôO™ç«˜{ïÌÄ|˜oøM¬Ë Ž‘ùx8Êÿ¾ÓèYde:yxÇ\•ÍjÅ¬ú‡ò¯s—¿3êðñÈž´[3´Eù­Á¸>ê÷äªÀ‚gAy³9|AÛ^4,[‡îlÐSdÅœ?ËóæÌê¶BÁyHûk2 ë7é«Öø÷µÒÅIQÍÉa9÷{ëtÚëtkt§OÏz«mž”óãšnRïÙÓ/É1=]Tw,2yeç%ÅÑ/)èÑJÅŒeb@Xÿ©­§FúOH”²Cþôâéw¾íÕá™u÷SKl^B#ŒŽîtÎs2Eth|£ÍZæ>$»>_mû†6)§ ¯wÈ´x¡nvþž°ˆ}rKä`½ô2PSïê· †7DÝ@U¦_½A†¹L²÷° §—rÇ5;³D§<zÖRÊX3ô¿Ÿ¾Ö'SÂQ¢þVô‚”c
bþPà4¼q»ƒAÌ	ã6ùú^¼Ñ-qà'€†Ç½6ÄøüzØø†”òCžÕ‰‚i:+ñŠ>O,àßƒWmIQõñË²™¬øBýþ‘âäsŠçû=~:e–¨ù!³2áµŒf½]Ô{}ð/Z«‘Ç¸F+ß›Or›iË£EÅR·›ãq«íÆ{Æ8¥’F{÷üç¯…iRÎGïŒÆÊ_‰QéúUû„¶9~:+§«<CŒúmí’N#*
6NåU#gŒxdÁ¡Ý”úCJ™=~„àÏÕß	`Wk/¢ºâ)ÖÞ~ˆÌòÕë½þ«×XK÷t0«1ã9¾jbÁÖ×ÉºUdùç€¨² ANG*ƒ=hý¸Ó°Ÿ3ÞÁlÏ|aµÈÖ”7+Ö™>Ù“9|øL€»zW’—ÅÁW´¿w-ðú š|W´AÍê†=¤Wg=¦‚Õ¢•Ø@îÊ_Vï=h®`üný¤Mñç54.ÎhÛ@ ê¹pðŽrÐÁ‘qÎúüAý´;Íˆµ­Ð&ùÒ9/¦OoËë?)'uß´ æw›¦8RVO¢yze%/Êy˜d.Qàvm£Æ+Ò>î›š¢Ñbºé3äbµ2héô9ºX%¯ ½×jcøñ°°È]xÿÎÀª©Àé"p@àÜ 6ÁÍHðf|Ÿ_x2wG}a‰¥x\ŸÐ±@-f bÐæ5{ÁM(lº[µ÷¤)ñ‹²yÇj/ÊÎ˜	|•±qÏæÊ™Œê\ÛVö(¦£w#;ö‘n;ù©´s³¿ô& o~‚Æ‰ÅˆÅµc4™«¾CŒE—»Î‰m§?8<¤{imð&6’ÕtÁ…Î·y¾T&¹ë_ûm(ÐmÆA$ÂæG°¥ûÒ­‰Uq‰ŠAô¡¬SI(KN€Ñ9.Ç3X\zÛ§›2© —ƒ>)”¿})‘ÐýâŒ zŸÄï’Æmˆ_jJ°w¼;—!²aÞiÄk6˜;Äª@á~3’(fUËFÝG?#Ÿ“Í)Ý	|¾bö>ª_V2þÿze Ó¥)­™ãÓÏÝ&Äš<)æÇƒÃqMy>Õð:éÃ¨èàn«ÿlÝZ±á_µ±Ë ~ƒ^>úÚáÍ­¸;.›yæ^Àp¦Pæ–+âÞƒ‰¦W=µC@ã¬9òù¹`ÀèË&ûjïÙÜÖ'L4¨ùíqù®ƒ¶ƒb‡£÷ÿ:¥ÂàÓ¡	ý‹I³é·ÏK
çS–}
¯ÆŒHé‡ôÆj<=8â6uMÝ2#!¬d?•¬’ýúÒ2°«è¤ï‰õ7¦Íæ…yó5§þEÆÌïÒK¨›1Ÿ7ö¥œjÝ€¼™-aÌW<P3¿s_eÍúfÆ¬Åæî52ã³r<®O «4&Ì«ùòŸ‰éò‚³u¦¥'ÅÌNUsJzBö7r: òUåTd"kÓ>yÓ²O6çáÍBz"Š/.4ò øþÅ3?ªJDÇÂ" “FæÇ¥àZÆá™¿gñ[¤›UZB”¿¶¤||ÉLé—$ðNï×ÃþþGYá{O?~pïå£§ß½<xòìùƒw_üøìîó{=¿ûTN+¦ìÇBf;ºw½ÉÂÒþî|‡yh• dZm2ë#!kí÷ÄéDäšXëËþ%gîUm¸ë“¢VÕ1Ð½ãKI-~Ï73>Ófü1ž¡GËfÂâvJq ÆPòO'0ìoí°PÞ:2s‹]IóuÐºÁÑq:ý„2¯N5¯k	6’1dœ—århÛ§S,K—QÃk]pJªyKlá´oN'ÈÚAVÍ¦y©‹ž›#×Pý×ÆüŒj,™]ÅëÆ’m¨žI>Y¶uvô3¤º×ñÃÑ;
aJ6‚ÍÅdláµúX„,uî$$ªpž&ˆØåÎ:Ÿêå³_Øëd¿£Sá×F†|Çâ¡èï~ÑÔôº²ÛÌŠ(lèÉV'l¹N›¯ÑÛY¶ D$_»¾~øR†ý]9­ —Ÿóµé¦Â«úg §[à×{´›Å¢-ÛÅ†{ºWL‡åø2ó‘D:ù®¨yzÇå»9Ôx#®Øaxº ¨;§ÿ0õ n¬ígÇêK5Î*ÇZÿ¦©§?—Ë6Ïk;í;äÿEñ®Ôd¿Ü!ÏËb8€éÊùî_>yÌ~K8š{¹*ÁS>º_WØ´	ôË‚äÁì¡øÙ§C]4­õ’éW|äËCU#ÂÂØŒ`-³=zBL<YØkv·oþb~)fÜh^VœëbÄXZ]ò<'²œ/Ü&û=xÖ3Ä[’/Ê”‚ñ™vãV´m‡LùjW“Ïcuž‡{E>ÁZ*'EåÕeÑ¯EÊH¯‚xŽÕ9(*ÊE¹5øSìû¡ÈÇkÏŸbß«éÛ'Å¬ukÈçhæ¡òº2_bµ5ýw«ê7h¯”ó:£ÏÐ>lnÙéÆ|Z“Å¸D…½Àjùü•Ñ¬ôêV¤äNø”ûhîØ?U¿I¾ÿBW.¤
Å)æ’#Îß@À¾c‰— ìf‘ì¼ªöþ‚+—KÉÊ·t„
/	~^]Upnœ¡,9äµ]d´™Ž=ŸýAe£8˜³1<ä
ó§÷%å)hHB*GtØ‹Äê„E¾`ƒòn6®}€µ0ßÕ'¦â¬vŠÑ“vwŽq¤¶Ô5)ðp`^Ocòbÿòª^ÐËæ=Ö#\kVÄ: ªk¹uöÚáêåyóŒd@_Óã[‘
Ò²ð“à2WLé•¿Ð‹Â6Wïúgç& ^ðˆ<¼…Žü°`M ÐƒFFåàm… l•ÌÇ–"ÌË¤±¿ÿ½Ed™-<´ÍŒÜ@ý‚5hè‰ûàS¶í-9”ØÃí{SZ…ÁZq‘¿(Ò¨_úöIt!änÆJìì9‹±»~°·ô‚„Ì°úØÄVÅ$œ2~Ðë›l¥!ß–€“Kr	Á#™ÆiqÄýPõ5`Î9Z(aáê;@JÏŸ>~ðbðíƒ/Ÿ>ÿñÞÓ'žß{t÷±þËOwœ¬< ±ë[s õ¡šŽ nØ»Ü“£XñÑ»GÓÃÚæx/™·ÃŸ§Ÿƒ:ÕHŽ4Þùœž—5FÜ°Ô.ÖÉýü.½«1Yÿw5HM)úßR’Cïr+ŸÿyúçéÓñû¿®’Ï¡O>WÖßç7þ<ýžÒ¬)TÕPm@~â°>;«p½<£KX¼«›Ush÷Ìƒ'ßØÿÔ»žæ†Žù[F³·iì8_rß4®¡=$7zÔûÇŒµƒð¾ó&âxêjö2‚´ì-B–½·úoÏÍ!‹Y)	tq£g4*ñšàã-F#`L†õX˜5d0%–`y®yl!Jƒ§Ûk~ãW~ãW~ãWTù_ù_ù_ù_ùµò+Z@„ó+JÜiZGÑâ¥mþ,ímù ‚úw®¥xXÑ*`-øÀv, ßklE{mØŠ€ÈàŸ?zù@¿LoÄw¾Ñ€Tû·Ü	OltZéÏÇ}l¡ýŠ6«p¬¸Nªé¨>0fÒïÝ/Ûò'
W§Ã1HvLö½ß[ÑËc
¸$ 0ì®Æ¾zûß”“ú]5*æî£–Û)¶|Çæ(3ˆìÆƒ\Ò?˜¼Îì2²ýÈN¨5	m…ráøa$c›¦óÝiœÝúf1~ëî˜³O¾»ÃØö	3’_:û«f÷F™‚Ý bÏ)7>ž0BÎ”Â5=ãðˆ?§üÍgçÁ‘\¨ôŸ\ÞÎTí¾â†¹!e† ‰)‡ sÛ§K±¡wÝ`ªP6…¸$Öþð	i@yóÙ¹ù±±l=ßÿìA mÓë °£7Þ$¡Ú°Ôí «Sq~[ÐSS°ŸP,ô.G”»&€>8enwˆoÆ^²W÷e43Å—vFÊÁ«Y¤zB‚Æ½ïžý‰?Õ&AòE¯Ù7†ÂR~ôR))á©[”o{òµÕXÏþØPDÚxÀUð-S<ªçL©øþ_jñNèåëo„^Þqm¢Ýì=¡Gd³æf)Öû†Ï1ò±ÛQê@vžwó»]aSçFYŸõWÓ·„)(ù·R_‰}ÆUvú;öÛùP)ßˆÔßÕ¢†­·k)ßh­ôëí¤Ì­~ª«i`\ªçä™æÐú²~p:,Ç}»«DöGå$ÎÏ†MäiPf<Øƒ6\x¡*Ëe˜Wš‘àw»?xµñZãïðP£uþœuü°©'|üÌG\‘ø‰œŒå4b#YeòýC®Âý¦>aÈdD4oË³V»B{Q+¤ûÉÉ?”ÌxûéÁO_±ZÐNèÆFß¾†Ïð«ú!D<ànÈ¢q4ktÿ-Ñ[D£IÛ¶ŸÚò‡…ÕU}Ò«!íô•üýšñ¡T®Öö7!…øVÕ6ˆ…‘n[„ú`9àªŸ…¬RO’þ»bŒðjzè7èj\Ž-Ò×Þv…§­$|¤q/!Ùºøž»)ë0ÝŸa2©ZZ
Vz(ÇŽ¡1µ#ÏÈ~,­jlÊäË£zæV´º2Z™Æ"º’ó;kµÑIªÍ)âÓo£ø‚î_œop‡h$ÀÉ¯âýƒÌ?5l"„mÒzöRq„ážHhÜà¥±¼`¸7Ìxƒˆ¼þQ}ò?Âè‹ƒlÚÂ9‘ôú‘À#±LIÓ0^b3ë; Wî²bÆoÈÐ-Zl"VûØ`†b}eØ¤™]ÉÁJ2ý¼8¹<ÉÆü=$?øÂà‚ß(*ÂÍèov0–ä€†°1w°uu6-õ”A;WôZ°Úóã{qÁƒnŠÁ!_9~³L€ß¤¬l*Ï­H(`ÛNÑ;íD-JO 6{E¹iÚÜ a©iÄ²Ä`TA©hÚº-¬ËRRÿä¯‘=×»n?Õ•9ƒwèÑ°¹D{kJà¬¿þçûëG«°xÏÏ»®~2Fd›q.58aê‰n2»ƒ±[«0ýŒl³6Í™™ºÅ±Ö¥¨ú^tØZiLŠN‚ß¡me]yq?Ä^àA¢O#M™¡9k¢¯žÐƒe4š½Ž	kN·ÎM–ÍÎ²uµðoV|EÉ¡ëð…²äîˆÑÝà5ØZY–¯êÚg\öOm)œ³Zú~¾JLcâì2Œp³ûã}Ñ¡4ÔÍ‡	ßþÖyÁÄC9‘¿í}ŒcÃ™eeïŒÏ'ÅÎ0³,”«'ûsÅ  ã”³±å›ð¦}[ÍfîÁfMi¼’wapÆm›žÒ†Ý·Ù`ãûÁT×É±¥í‰•[Ç”ÝÀúZd?ÙFË‘Ž‹¶o¼s¯Z7X?îÿR%¬7÷2/‡CÛGC)ÀEUPDn‚óåïü¿K“7ŸËžé/ÀtÄçb“y8Š7êÁ©Ž(‚Ïol^f$Kª,D×MŽà×vY.v¡0¯·U˜W4›ÐÉÌ±cˆÍÏ$tÓ
+ëÖ›­Ý…Õ4Ópàq2·¼œ¦»ß,ªñˆ¢Ä-ÒRŠpç|ëÖ…iÔ‹KéÇðÝ=Þò"oYq=˜½_!yU¥H’(‘¤Ýþúñ–ÓãÌÏãª"ËHªn/B1:ëôÓT ) lñîý¿N«ÚítY÷ç’y¢Òõ¬ÌF³µM3nê¤ñbã\L…V#©ÈnšÈ>Õs—YqZ‘¤K:Ž’hÑ^½¯dJ(ÑvÛ‰B«&ùu”mK6‘;Ì›ö¼šéFð‘2X¹GÚÚ]zQ]gõbþ-$áÇhóË7`/*v™c!¦>ø†±^?…{Ìâx%6BŽuÚ*";—H	µÑ°>3í<æÍÂ¥óN j;7rùtíöµÄ]FËiTgÒYàgãE+)Î†¿þ,X=3,åÇ¯G÷È"¯#ƒØYÈÑØÜ‘¹8K´é/‘ŸM<š¦~÷û1E€Íø‘àl° þh x§4fî©+†Ãr6¿ÓœŽÛÓUÿ¸Ÿèhî¦0ŒizÇÕhT:)vöÁ‹¬=8¦>ähº¿Ëm¬‹=±³À¾_ŸLsö•ÏÛW°-Ôc'W€o/‹òâ¤‚Û_cdpCyÇ‹?o²ªÞþ„Ò)ûdJ…dì‹pŠ¾Ž¨Þì€"'†¦\¶qâî‡œÅ06yPá!5Å*\reÄº½±§U;ê{6
Ê9²È¶6Ëé‘®/Òe,ƒ¹×ÿï¿ÿ—ÿÌéÑÅÙ`cã=P¸š½a‘0¯wk¢„!ºmlp÷®ý×ÿóÿíŸA&ÅÍ>ÜÆM?Öí3†ø1o"GÃÿþÏÿKùGÜæˆDeÍÜ’>X´+>¢>w£Æ€¨ÃKä%2â®Y,š™0h„‰±›ÆW‘ —¥uGcOôVz	ZÏen íÊÐ»;šTS0ó^Æ'FÓC ž q;8ïJˆ	@.9×é¬Ûcõn9à«é¨:ª5Èëß~­«{çªäð <ºdRnG8ð–<5'™Ã¯¨Ž^¥Òg¶b`Å†1$	­¸Ú8hõ4™‹–LLñÖKSŽá4•.óÎ×/Ìm=^ÐÑAÖv
”óz¶¶¹¾Eø–²‘œ±öØÅD†ìè5 ZwÛ43CõX@±†ÐfêkH%]-ê³Vn"5”ŽŸç^¼¤PVð}79^äz!)é¾‘Y#9{3q(Ÿ’uÏ.ÚPüÀi×?dÐGˆN6®¦åÚ¿8ó½â¸éÎÙ7Éd´£n£É‚±TiÑ$i‚t B×[‘ŒÉ¡´–\×çŽ O”¶Ë­Ô}’€#¡ŸE	…£F
^©·n.+’\˜ áÂ?’ì©fI¶Å|EÁ=ïow¿KV2,÷ôßÙÕ-CÂ=úKqù#Pö„{êÏìÊ¶eßû)Ð…¡Ä„ÁM"ÖGpý•œ
$ŽdÞá@2v|Êg¤È?#,Xàü7»
¸ÇþÉ®$ûíñ?}ð•á@•¾ÿJ ØŠñšºVœ_+ÐžÛi×¸â"˜õSöÁÒz..Ôê^ò.ðD„¡Ñøé!?¸ÕöYWÃ‡8ñ‚3y§Ö'	¶uš9÷2ÔpÈmó ·UÛ† Í'þë„Ü‡÷¸UQ¡ó`UåßúÕ©›{¨vsaÙ¯–åá>ÀÆ¬KVT(”s\~b‹IbÒçºeßCo‘ñ‘ñóK~- ¿ÜM¯HC»Z‰RTçÞãfQž®WC*rPÓPèŽÉA "¶IW|ÇðÌ¤UßN	
ÏüîéôpfYk+¬q!1]ØÊ¹
Íîƒ>Ò+†z‘ìï×0\Õ¯›$œ–Y˜PFùªGÖ"ä$þR¥#·-%oBö÷æ‡F+®>‚uXNÊ¦Øq1`“ˆÅ
dº×b
Mú! [ñE B…ªC÷ãh¼ó@êóë\Z• §ÃÂšâ›¿´A-ÒG·¶2¿T˜•’­_ de×ùª¢d*±¬–¼ïÃ/­‘íâÒ‹ûMKŠ;»7Gãq³3­SJsº\ëÛ`	m³2è‡ÿ@û—CÔ_H°×ËxÚêõ !ˆ$/õ¬'˜ŠÃ®ñQ…Ø5
þ*Ë5`ÓÐnýËëYvmº×}Ù™T÷zÖëËÄº³\ÏÂ‹<‚×ð:ŸáKÏEã×³öb`) çC¸F|ï=0ž€å%Ä3c™¬ë)yVLË±a~É¾	G3Ãm…|(°Spz™8]ZcÃE;Ùa7õI”7`èø&)¦Õ€þ°•¦>ÉÁïÄNš¸Ú¾0"¦íH>^Oö"ÁÖèVŠoz¤çïµ°V !%Vëø"6P~J		]Ð“µb1¯Õr–Sß‹¼: „¢xñM£ðpiÔ›ò!ÊÛlÅ˜„$yŽôö†ÃxŸ½;’xãÂNðý²Xž0U
rN1_™+ZY7az)oK©®jÕØµL†1_hÅ/ÛÖ$,£ç­ì.½ºµÇR–µ‰É²ØgÂÇ?4euâ[’ƒ=W.\Ìùt+DlC2¯ŽÆ¥ƒ;Ó®qFse”(¶¢[âpÆæ‡ÿ]OkïH?©!>ëtT“s$_ážÌ%&ºC8¢ˆÇGÂøCÂ/	
wš¦PøPÌ"šh¤Cž8És^2ÚjÊ»ã±D_ñÓj|PÒósÖïÓ6üHB¼ TAEr²ÿÈ‚Žü´íÑøQ–eÑ «hâ„beó!7œÙšGG;ê¼}¹áÞmšâl Y‹ûÒÉþd. µX|<{ä¯»×{vU=Y§RIìëV™»÷¾‚(‰3*8.ï[ëS’Äcð'Àß$ÄÌ~G(â€%mjÜòÂÀ^ÒÿE9+o|]ê`šÁÊ¢¾¹tqî4~Çß²µIüÊï*’Ü­²Œò2°ƒS2¶ÈS‰ëÊf.bª•s%ódÿl~M')¯»<è9ašG>ŒÕHzîòbÊ—n3‹X}!¼‰I™ì«€ËþîÅpÇ@ßŽh3¢l q4@Œ	|à¹~ áÕêºõ-ÔC˜m¯$ <qàþHY»ÒD{x\ßÔ§,¼ÈÄ
cm³$º¸GA™{ƒ÷67ú¾E~¦ƒ\`	nµ­§çrháåa_”£;çzð!C1,‚|€,<MŠ^Ê‚¢h”¼;DT&fä9¼ äH–(OâUb”`ú‰UfŒõ:>£à;„àÊbz6pÜá^Ì:þ„‚ò1ýŸãÎè½nàtYÎý@ÎžhÄh)„gí	FP)}d×àÌÞ$kMutÌ\5<s &˜
 ÁÕ›ÎXžZ.>bO(x¢ynî…|A‰ÆM‘1ðLsò[x ]°ˆÚFíHó¡ ~Ù_è¡ð7SCáxß‚Wd)O¶‚d
J „°3”¡6"]Ño~	C¾f¯Ítü{¢ÑÞñÍ°ÌÕÎÃ‰þšËß¶¹ÊRÔM.9_Fº«2¤’Ëw|32úØõx“ž2Š«Þ®…É,k#(3Ò%%=Ò%ã.m?æÎYzµ!@ÜÁúP,[)×íÌ• nFS>]b»@÷b+ìtè|”ß\çÎšfä]¢û"3éà-O^¸}=+lKÿ.¿ÄÜ§é“Ëî#'M1c×A©Rìå›÷­€Ä,2…l‚OD¼!ÏÐ7èkA·	’\‚QLÃ™:8£@!°¤ð]þ6H§|'u‘AÒÇ†%=ÕMÆâŒ±Xn•”ÃYÞhäÉÑË‹Ô‰%†ÒŒå»hÈ{¿ÜðÚYÉ›¦d¹í2H©yDQ’û¹Pì¹ÇHuHµk6¯Ó·D¸xÖ*¢÷h"ÚOìðÆÂÜdL †íSkÙ¡s-yÊY¡L’“Ä
~Õ	Æ´R‰³F´Œmpp”¸¾Ç;Aõî1?Aô6.ÌEÑ¨†”¾Ûš ŠR9–˜™Ñ2óë¸ ýsSŠx¿8cÈ•‹&#6zK Ò%@-N d€’‰„º4ˆ9¶øŠtØGQ‘ÏDjôõÕ€<~ÕÑ
yu½ŽsþC/mÈëC$®Š]þý¿ýÿûoÿ|‰½ò×l\à~°bÖÓ25€%Ò§(3ì<n\a;ËL3l²£KÆuÈW2=²Œ$ÅÑ‹p¸‰×¡ë½¿ÉÒó&s§ÑAiì`—)i¤ŒKQŒ%ã#-}ÓTå!“xéën¼‘K²ª? lOX#‹8Ëïy„…/5ÌÞJì2æ˜h#fÙ b–x“W¹ï´­g,›ôUìº$¦ç‡L½Èî›ÉW.ÒÛ•Ü0ÕÏL~•çÕ™\A!¯ðˆc‰Çþš êjáJBVò3¢a/ãÛ t’š­šŸQRŸ’6õÖèsøB(Ah7v4!JbýeœŒ¬³qí§ƒ'/‘ÙÜù$gCÎQv÷ê`y÷{–íòøQàF'Õk#¦æyîç½:®ÁP~±‘vduVíVdá¤Î1›Ø»G§ÖP¶o¼ƒÌ9ƒ°¤ù€ë‡¥'E5¾*P²j7‹éÖíqøbé¤r!ëLzö¬ŠÂ•M[ä¹ºìÌÏäÆŽçË”Re’ÊË“É®¾Þ‹ÄÐdic•9ÿÜÈSýÛV†»ôVËdàÆ*ÿR[!CRVFdüx¨Ç3"Wœö(3øëûŽ&ôg¬Ù<9{nÄ5âÞ¶#&ónA²‰^rãótŒdRœ®¯mÝ" €=×'ô†Î|x².@8ag¦U0ðŒ3Ò¤$Ÿ¦fZ-sñqÙùƒ#gx{cý¶yØÌÂ‹ÿ  ÿÿì}ÙnI’à{…W¢¶‘¬"3yH*U)ERjÎP"‡¤ª{ ¤`f)3#+"’›ËÏX,°ØÅ¢=@?5æe^ù'ûû	kæW¸GøÉCR•UbF„ßnnfnn|@{Ã{,Ô½RÚ ßätF?  ÝGò­D–°q°‚j
;?ÕöTkñœHfÃ?.ì&eZ&€ ;Q,½]5È)‰O0)‚2zñ›[~ršÅÇkZ…î•a®ÖZoŽ†Ñø½³x¸Ö§é$ÇèëšŠ³Ì®ÆÈR<¥QD¸ÓªzÒªÜ]#ñˆ›ƒFŽÑƒyý†sök´gÃô(öYŽ)ùéîÂe
¡ÑÑìÌ˜Áy£ðEküU ˜Ïë•¼lÑŒtè’Ñg®­0†q4¤ÑUÂ]ªÛ2·Š“GÁ,L¹¬¢X–ŽŸ!À†e¶ãìšØZÚà¨ïÜÒqt³q¾lài›Çá{x<ˆ2{Ô$-;„g0ƒY§JÜ°áVs³®¿åCMó²²vuóWÒóúë«‡=ÖðR³\2´,Wùü‘òuãÑ":Æ€bJÁñ±d0d- vÓynµ8£qÂŠL”1Ö”Tâyµ´4ãù£$©Ýk½ÔÕ1ï1›©`N(4„g¥fWmZ§eµãÕnq:¬Ö#æ€Žto¹Í*PCÝC5°"QÓ:dÌêfÅvÒ>†"ø./Š‡7~N7Zü’	s%ð>3 s×Í«ÅQ:8W…Õð[8'üGU'Í,ÿndßˆé¦lËºf±sÄä°ud3”Y·×ä‘ö.ØV§æ =âÔ­¼†Š˜”™nR	‰¸äF-HV‹KP™Ji›‰!7.´E”ñÉ‚Üf‡,…²ÔÈ‘¥&†ˆ,53Gdi&£DQÔfšTœ££O[’%#Ïä³ 	´ydi6ËG–f²d)|VÜ¶,ù€Îc!s'Ö‘,yŒ2|ò&—t¨©(Øj&ÔiL$‹“s=5Óësæ	Vw3Ž©&Þ.µÌÊ{E¿˜1PÝ,lD!c3XÙpP¹‘Š8¦V2¦nÜ)`Á	!)PÒ¸ÄÙö‡VÃØ	õSTQ%3£aif3X‘73Z	ªózV±zÏ‚m?kÕŒdïë6²ÊˆƒldEò*&…îœÛY‘®½…¼j¼…ñÈÊ¿Ö8µæ
œ•@Ã6oÃ­;›…Ûâq¶n´›7mïæè\C›7Ñ½²{sÍZÛ7Ñ«›°ÃtS6p˜f²ƒº[8L¤ÄºÌn‡)˜­	·£Õú™Å=·ôÜÊÞ„Ùº…ŽÏÇäÛ½Ño‰¹¹i¸`ôÙÄÓbfr ¶sy=ÖuÞnTÓú¦¬ðh«úÖ¿V%Ìð&\2¾„EØB~«Mƒzù,Lê–E5´h! ¦š6ö€CˆùJƒInˆêusž›˜ì;0ç)•–ƒhÅíÍ^EEûšÓçP¬hUa¿R³ûÆ›we>Üf0s¨º¬¬µ'ÀX-‚-k6÷®aó Ù‹45u±¬©]SÛwhËiYðwÈiÇs·ßDWÎ |;ÂéÃ¿oÌÁ\Ye–Cu­+Fo¿UõëçIÍ¬±üèh7]ñúíoSPü÷Û|[xV)ÀõTçS×ÕbI›&î±Þ­6<q+ƒ±»U_ß™‚oKÕÉqñ¢ºªJP€®X€Ã°;X7ÿ}Ûl.áfq
×ìÒkRR½›w°äš{¸/f½›»‰³Mú»Š˜ñ°`JŸ|³£z³Ž&³¯yÃ<C¨¸SÓ|ëRÕ7ýµ)NEùM|Pâïv¿#Ï£IDÐHÄf%?'ñÅQt5Hï
Tcôb%+›JÃ"ZÄš¦.¤Ž¯cá3ý£¾úiïéõ½¸’ß¬Éà:Ð¬¼¡Wó¦cvù9¥€=³N–Ö.ªoÔÜÊ9qM=4–y,ó,%89iSE¥(Ÿ#ûñ$Í
ßŒKy^w‹SÝ+=õŽ-æäÙ4Äd#Ê¹':IH¼»*Ÿjqçð™dŒ÷dòT78¬‡L^.Ý4U|ÿ ë;¼¬j¾·–¯:ùôžC7AFÂs*ÜíÄ££P1:ÿÙá€4®æ[ížÞ3´n©X™¼7¶»Ž‡Ž˜&WÏe`«2HhLöˆaúh>èÞ/ÓØ¤D×g˜Ü©Ñy8ÊØ`HžÂøHJŽãþiÄîa”u7³¹¼hÖá³¼_¾ µ¸€n ¢t7o$Yß	 –áëœÊÐÂézM6g€Tev‚ê.Pµ÷qŒþÈ®þ’%ƒÔ^¤ ¹´b‰ê-è,Ñò®œ´j	šòJö3LÚ?Xr¤‰p¹5Š³“«ÃR^2åä4„Ë¥ûÀL§d%yÇ"»Ÿ]ý­Hú‡Î¤ Á»MPŽ‡—ƒ¼&tÖ^T¸Š$/HzL˜«¾$ÎÁ+ÎúÜÆÍ˜ˆxX-c7ÛðÆ˜&ºŠv}é®èº‡x™í£òôÕß³„.Þ:w{(a¬0Žµu«û™	W‰ƒmèÒaZ ‘ÖÚf°œ÷ÌÞhàÖ,= ¡åDl”˜…_ój»Ø.<¹CÓÊ ÿæ®#ÙpôÅl=>à“VÝ¢AÂ_‡á%îzeE ¨ Žï†e|O§£HÆ¨óLÖÓª§‹ðo=¸h9˜ÄêÕ>û·ÈÍ˜n‡ñù¬Ì·ˆÅ€ŽjÒ¸Êf,d*«(Iÿ)i?Ýüy®iEÜE;o\Ïôpšd<UãÂ@ÉfÜO38V¦V3'[qy,êf/Ö¬‚wn‡ÅXÁçSá²íB›Õ—Ed5Ú"7f¶Å1—ƒÞ*0Ãt=C®²†AtžW
K™ŸìÎÎXa&¾ÙÄ«ê5Ò†é‰FÏkÑø<(˜Ø7¼Äœ0,k-´\ÒëaŒý-âÝ£w=‚­º2c¼þN›d<È}6B¼nÁÄª5ïÀÉâõ‹Ù›š1Ä—°‡z½hóÊ;¸¨À×¶çæBg˜ç)µÆŒ±Z¹´[“báÉ~ËÑËKØx1¹Œ1bRv|˜µà5\íúê¡-kåª÷ö7TIÉoj¤ÏÑÐÚ¦>U#Õ\9Ä;¦{ÿØ7»	÷2ÆàuK^É€T- NV¿íªË{t¶4Q|'OŒ]ôÍc8%ÕS#§ïÞÚ¨SxÉåŠ³H	RåII¸·º™=¯Ó¹ÐŒsêÐÍ4B0Í ;D[
wÐ\óPQ
ÈKÊ[» Ü6µ£R@PM€~›[3hLË2È0·P´áÕ†¢UytuØa3X_ÓA<"rèäLVFäíÌ§És[›­™+E–ügÞ´&'¡Ú^ÎQ	Ç€~i}bâþÃö¼ÍE ê°æÐlÅ©x´y\	4[CÔÉ_Ø0C÷½ß×«íÎ°C´°ðù"€ªædp<òDIgúY?ëÅ§@Ra–ñ?4dj/ðÄ}&ÁÀDFÃ [}‚w*íÜg¦W|$£‘©X8
ml.æ4ªehr©©Å·kj5pà	CÁÁ¶cÍ¬Ç®o?Ö„d7Vdtã5a;?~n¢-M«Ž%ƒêÓ
\‡XÓécbj¢“‰I×ËZx¦»ùr2 "/”árÎyÍ¬=
:lVÑ‡Ç^Ê³ÄË¥šïfÅT^Óå$A!eº»ÇhœŒ~ “—p?Î-&Õ§~e–nëü¿ÿýþ—¬Ô-Ä¦Ls~1p'Ü>Ø™
èêÇuh“ú4wjrbnÒþïÿüo¢Êß$œéÎnÖŒ1|5p“îl¾­ïàèý,fèF€nûdœf_$¼ù‘›Ý¶…—mlß‚Ékã‚i;öÊâµ½¡)KÆp† È6ÒÑd¾­‚øÝ†æ3´L ˜Ý€˜ÞnÛÉºÎ&«•‡ÅÎÃ¬–£)ÛéßµGÝ.áEŠÚ]»h‡0-š8íR535I˜déq2Œw²t³LëƒQ2&Ï£VÌh›Z“´E¥ÿ¨«°F5ÔQ¥pRÇÉÉÚ…ü©›r¦Ð ZpÐú·õAÒÐŽ2µe@cè!Ô7>f$¥¥Ì^”õã$‹*KíÛôÊæ6˜Š¬®3,¿-Ä@ËË5¼HÊZ«2„Úù&Á¼d}YX$FŽÈëîÏUÔd2ê¿Çæ#ø;ÈÒ	^¬duµ’ÕQŠˆ¢-W`3Ã¬EÃµ‹©GçIŽWñð«óãý:‚ã´M+³$Ë,ÕÄ“¢Q>ÍS®a
¿¸ £V5‹·óêÇÅ³Ó×DwN©5âW±¾ž2ªKóZWDe¼DÝa¾3¤õE¬;zøãx¿i¡{>Ü»â[À¤jÐNuw1¯ïÊ.;Ž†ylÄ¬vù¢‰Ö[6*eýIh·›©’
™CäXu'yT'
l&	è*Š›ÝKk-~þr˜gó`z4Â-Ä8êƒèÌx#Y^a`v?pž•-ÊÔº¬ê}TÈ¥mqjÄ%¢DX%ß7°ô}yb%¯vióÖÎ¹TiQOÁÎ5âãh:,~Ž†S :}›>ö¨Ídñ/Ó$‹Ö†hQÄv®è{X4ó—uÂËFÂé´ ÎW]ÇgçŠ(ð	ÀÕynPEhVp«M€¦Ç§0à8[k-.vÙÿ]øqÁåœõK„{²:=¾§…³òÒ‘˜Å.Vmÿ®aÖ£©ì†Zfgí[MUxVø­»ªtx§¼&TJöïFÀÓ>õ«é9drF‡Üj=¢–èð&ît:«]ö5¸¸"—~Tþn\¦Åýž¤ÞtóIñå#ù³q%ºÐî‘&¡òU¶Úe úeé=ÚfvÊ:};RsT=3EÑüƒÿJw$·jµÔû#ü·qQædýýÓ¸ð“,ÿ9n=bUû£Œõt‹,VéxfBà_Xc·¶>öÈÖ`Ê·ý<ù9ÊâwŽæÝ,×—vÔ"‰M€Ä1‘Gf°Sq}(UœµÎ
¦Aþ^¿oN¶Ðgé-‚õ‰j&<ódrCŒÇëêWP…Cî9˜ìÁûÉ$’¶$p‹"Ïºþ¢¼”g±÷5î°XÃÊ¹5¢9XÖ>ºŒñ¾"/õÛmCì ÕÞ@ëþõI¡ª/†¿D½)^ïó•mQŸÆWÿv›á7yf*-ü.ÿª`êÓ¬÷ê7ú›9¦žu©¹[ë¯}}ÑÕô¸Í›ÚÀÌ’6Cò/q¡½düžŒSô›;ëµJHÁŠ¼Ç&f*Q>PJuZ“¼×íž¤içdØAÉn§Óùu‰©>!À g}cêGàç«¿¡mímR\}ÚÞuÀ‡VÐ~r,ó«76³Tv+àÄÍ²3BU¸IýF:Š3êW·Kžnþl›ªª8ÙÏÝwÅ®+¤ŸnÙÕm‡Ó›ÚÑLñÝ­TP=x5¢›Ã…¥E4»í«)¹ÊaA´ÜX9Ž«ì2Kqµóûfê½0œqÌc·ðßH]“¼xšfliP”ùðºCaÿi2¸pÂžÜêØ	qÛ¥ÇvöÂg;e:ÊÞP ÓNggÉ æÑ›·%w [ó}®S¤/ÑàF”Çma¤Ë€Å¿LvÇ î[Á:2ˆŽòt8…€>áØ^¤“…¥î2Y Öt£œÓÁñÈkA¤:tû]ÔÜñ©î¶×ìÐÉê»“TØ t_ñ5îŽ-=f¦Èæ°'âQäò/tVbî8/JwE®Bkv,dp×«ŒüÎÒ²uŠ^–6ÂaHsús1'¿LcQ—â,Õé%ÄÙ­â³†×¥íÉÝ¥rË Å)*züÞ@Lž“Í¥ŒW;‹“ÜæÊeíò³aç€Æ=L}å©jû”Ö€¶M`×*zn®Yƒ—C¨ÉëûcÖlÇ®èDÐdè¦rr‡¾rs¿˜Ø]Xÿ4î¿?J?º2ÍÖ|¦`‚‘1A¨võ¹ñ~#5&ôœ(™">"Ÿ¯H6(É_‰Þ·õÁøƒ1…8‹´µõêµ¿~÷<{PNôˆa&.îð†§ðã†Éi•à³A ‡)z‘j›Ñá¥›!p7ä9TÖQ0wúDŸ¼PÈÜ©&ù>²Vß0@†Sa´Î ?žžØtD âÞ¥Bþ‘P¬§¡/cPÞ*Âãž§jö&?ºT¯1¹'bj‚éD¬(WÔoª`­`ãZ6c˜€ÙV@ALAÂ¸ã â˜ø‘eJ¾aHk	ó_ãæÍËÔ#¯:Nµ£ó¢±×!¦À&Ä‰3`G`E|G(Sh,X±YÃ<5y°šÃ×f‰ËêµrÑë]sì’.^Ïf¦Ø
:Gµ¤¸¸ô°u>k»œÚ€™V;î5"«XDÈMé4ÉN ®¨iÉùé‘:Ð#u¢Ón–9<¿hÆåPK=ƒMãA4<Ã˜žˆ-"æÚ6r?öŠ¹£É¨ÑiØ§¿*dÕ/•¬žV»FË_™ ùòw¿;žŽ©Ð…lDC”äìÇ£)†ªëG)šP·ßÇë«<‚óf’ÎS$Â~¿L;Í©ëŒåÂÔ<j·˜?l^j4ÉdÈŠ=çör² t2ã˜f}Ò¶JX®3ÞÈ'ß<~:L£¢Í{ÚÉb*Un·æ[ó¤ÕiÍÍ¡æùâOj¢Kz¢×ö*hÝ.Ah†ƒjòã)%O0¨ÃwäùÕòšáÄ·ÏÄ¨œ‘½þ®l;u9¯Ì²èÂkuV˜Ê3XÓsô3†>æÇSt˜¬ÌoÝ©o‹açÍð”–^Ëç%UÍ‹s´tnõy½Òñšxßžìïð×—sÖl›.5sr®±Y5§Ñ¥ù3‹á!?JLctœ­1Lçšè`õt¹–}%ÔLyãê/äY<öº)Åj÷tYi¬.’ño´ŠaÃÁ”F$eŸáOt%Sæ4¥OwÜdj³¤À(9JãRR£cø ¸G¥õ¹ˆt$NÏ]‹a’{,¤:9y(WDwì>Ë}lÕ5¼MpQÍ*„27mûÖð±‘1hpÏQ“-ébÐ{¡bPO¤%Bö¿5P3ëce÷5Ç†MdZ\Æ0<eâ_ŒW/,{Èu‹n^:o>Õ˜oc–åmÌJcSÌ†72:_¡.K­·!ëBøÝÉ>0‘-µ›æ`% q0T®&(ÄÒF5£ô¾XËe¿=þ|à§k”ÚrÿÐÎýû£îJ\%1HžÜH«ìv‚ŒÂwš]‰E¤ ²-~–1$žJgY4º­¼ öxP½^1ãJ#cry™ÀÜ—\º»1táBgÜÚ’»­®9˜eùSÓ}Xº6†…6]­<½ú»‘G½úïÞÝ>	>ã£—oqFs^ªl/´®ŠGªE­ÎšhR	ë`Ž,ì¸âD›z 0h^0™W%4t…:ÍÃ7â]ÐøC@biFÃ]R0¤yGèc¨¬H”Ëx’ÁŸ³ä·ÕÿÉ§åã0Ñ‘TÃÀ} ÔœŒ#õž©Îdj?+gDê‚‹ñÜÀ}tá‘ ŽgYŽ?PÍ‡þà³è;ŒÈŽ?pÀô°|ø0ôv%þ”qQh¾t´™Dø&#îój`³')Œ8Š<šLžÇy°J™O±?ò/‡ñh‚ø~º±nÿ*BËßC¼Ü`nduOô7ÀÒáÕú$Á§Q\D| øóà¹øÅ®üðé—“ä„†¤§Åå{]h§»W¯éØÆÑð<Oòƒþi<Šù£‹uí{¿ãÂðy°ÓôÃöø]Ü/¨§£ùß]öˆ²”=‚é{ìüW8T‡?‰åÕ>¾¢²FºÚ=ªö²H†ì%_ùiz„ìçIñù´!zZPs+Î²4kQª~–&ƒŸœ@mð‡ÕDá¥G1ÕÖ¸ÈÎÙ[ žyM”Wzdý8Œò÷ì‹„«Ùà?yÝÆ rúw#š(ò`ßÔõðhOéß$B{ì—þ¹
œ=BÁn]¾aÙ,ðÚbÕ4¸*¦ŽG~rÂ1–‚
;x›ýÓUºÙ»ƒ˜Âj¼N7ñª¹zÜí´ÚnèUÝÓ‰	äçÿü©ÁÈBbÛÈqÒžòÞÐß¥LsŠ
„Øu%ø¸‘âÇÕo¿d–×/³aí-›x}”¦ zcöþRÌƒ¾ÉÌ/¯^]u„ê(í#õÖ1bû¨í#£ÇÄÈ€„‰½	]˜„Óž6&¦?„~¸î4à–¶uH·ûÞîÁaë'‚.)Ó€ˆä‹Ñý(e”äñ*¼¥­qôÙ#ÏÙ:“²—ÏËw§²÷ô7ûT"Øùgö[ìÏ>/As3õ˜Ž½*˜¶Gt<*P\÷ÂÌäülôW¸ªgx¤‘Þ£Ÿ¬øjKb‚Íëè¼¤Føä‹þ*¹Ô¤»„ÍNt¤InéÒ·N“¼H³¤ŸÒµ¦h›þTMÿ'q¾‡¬\Rfxlpz&Š<j+•¨Âcé)TéŒ±Pï!ÀÇ\G¢È]>
¼._pºQ>3t.Ÿ)}“O¯—Õ§?-‹,NP£P:ƒ€w–ås	–J¯†'Qù8LOÊâÈyÑ›|uT}@Æ}æ2b}<*‚l|Èäk™Ù¶ÑÙŽa³nð³ú,ÆuzX-¼á‡×r–µeBÎp?ÌŸ9[*õ¶\é0é'”`>Å±¿zýH¨nðÚ†)åª×k•î>¨us·JUµ>¥W¹ìÞBya åC$ìe¡i°Øôa&êØrú ¨}šG¸¬“Ñw»ä`z”÷³äŽ¥)a'{$5/ÆîbD{ÉG3Þjëø¶°v·*8¥“\¼Pp"B—òûË¤¨^bÙLücÅÀ×LÇ¶=€µ±»³³µq¸½ûâ s°»³½±}¸¾±¾ûæéîÎ³uí–•žÂQDkñß8šä§i!nÇ„­ó/•K{Vªú¢2ÏÓ¤}®¿˜Â¯JÂôÆ®¼ít°\}ù¶çÊ—ssx>ªC——'…S#p&}X^˜5rtNdT%™ûÙÉ!g»Í“£šœäâuNTÆez,#ã–×LÕ"O ÈQH~ÂÊ,°æZ­Æ#®n¸6Ž@¹ª¶I¹_h•r){nXBÀú­ÝÚÂ?„oUB õ¼×š'¬‚; F&î“˜^E	‚í9qW&÷ˆvGÆ\nÆýdó–p™òóqŸó/|7öjà2ÏÙQÆYö½d]ØäÀ9DNãM‹Ó»9+ðÐTîY”/ÔqªC%v½È€)ch+C•ñ@¨ÊSwŠÂõŠçŒU TªáOr™‹7T7^¾îÕûFÝ{P7^tA[Åå-àýø[Iû~„3/Êt”>G¢°Ï°	µ”ÕÎ+É–«D_µ=´ïèUæÃ˜³öÔ¹)s±Æa£;ggpº„}-Ž&%öQö"?þ–jCoˆqiLL¾½#Íù' 34„'^ûíw{.!1ï1sœœ1pŒZ—·ewå)[¼âsy	\<ð„tkóäÊŠÕ66ÜÙb«°M-WFŒ“f'Q*ðF¬hcíà…9?è—§«Ke£¾J/ÄbÐÒÞàxž¹aÖÞ¹é{’¿œpô<$¯Ai¬¢‚- %V¨Ä±`¨Ù5ÄÖ ÑêŸïlã]ÇÖ0FìQM#
OÖJgøœ?î¼ZäúYEàK-àk|ÛÁ;FòÙøˆ`w'ƒãV¹drövÉqt–RæF1.ˆ²_¦ÉYJö6ŸÖ—h_JN¢2ó*¡â86LèÐ»|'9Âù¢»4¡üv‹~X *Û™;Ï†éQ4ücš½³]vdì| OYªz+ÿúƒñ»¼Ó¦ÓÁ1ð´q§ŸŽºÑ»ècw˜å8wy÷ÛY3ìÌª»¤ŸX•@LÑ»üm9EY?™Û•ÉÞÓ©W>´+xj•ye›°¼€•¦íc §U~9×™°s›¨ŽzT†uµJœ	6iã§Þ/ýV×°•Îx:ÚC¼úþû¹?5oj¯°C˜¿(¤—S›Ì¸(sC6Ì~Ù`ßÚµB˜éõ–—ïPá;ãÙð'G%°ð©Ö¹Î»4·[¤¥TG‡ü=,ï¿ŽÈÞÕ_Nà¬¨.¹$ðâ_ÇoÉ÷ekß“Ö¿"EáÈU_<ŽÍiŸiÏ¤$ªSdQ2¦âö:™Ô—q%ËÁêU~/?´+_cÇ kˆ~a7–ß"®¦¹®þs€"w²Î6`ÆHA‹ÙšA6:Vmb$°•É)P•ÄD=Ÿì¾ÙØ}ñtûìm ÄãÖœY´T™’ž2*OBäþP!`·AQ~¨ÓÁ` I`^b`aDd£Ï:“Ò ë@»0†û7­¦DËD™ÞÖ)ôrwFL{ù¶F£è%ÉpXâ;fÔÙÖš_u5z‡§û½,>ƒ³?ÐÊgƒFÁ?b° áüqªIÆXA³PŽÔL1[•*â;Q´HM˜SgIM™YÉ+&X³ñ¨”ÃÑ¨Ü0ù¦vü>¾f#äkURrv¹ñ”Ò\mHßHŠÛ¡÷‰ù“â´ÝJF ]TõsÜx=.2Ã¹“\z!G‹šÐ]meÜtö	ŠãóíCþíc>²ž”r™¾~|Sy¾6^B§Ã³Yo”"ZÎ¹pRe‡,‹MíÓ*¹ <Sÿ¬ç›@	_îï´MYÒ1g³Ú1rTí~nþ1:ámÒ¡êÍü¬Te´à1"=†PRÍÏÞœk·X†V¥~QüùúŸÞüq{óðPÃ½ÅE{¦?lm?ûÃ¡9ÒöÉ -¼hÇéoS¦Óãò\ìá§ªj7.«ìÏo²ÎRsÉ1˜Í¸x£ß­)ƒíc1‰ÈÌõLU«]ö’7ÿH™Es?YÃ¼›|º»DÌ’u\jþ€®V˜BGGÚ8+<lÏÜ%ŽÅGd§X	 œ·[Ëƒ:âô@ÓtÀëÎ ‹>°Íà2!‹à?ÚÃyõs9 nÑlXv¥HÅ®åèo2>¶ØyxßP	Çm^M-‡uÁEAº÷9Ý¤îe§ÃåwüÊ£VŸö|iBÕR´’1ÑCjoPæÏéqÃÝµÔ¥“…µyY
¸*'¼
ypO rÆÿiž’.Y¼JcËÒs·ÉDJªÛS;|Gü#¶[$àäŠi4LþŒ|ãíñŠHÓ#k²Fgåuz®r‹5–*PPíû­±fdL…S-ØŒÓlÈ~Ó~DrØâuqãYßVù¬±—BcC2˜’ë*;Åu)D¿ø£Þ5:"øÅnJPnF»ÇØU€ýÍ :ß=þc¿‡Wè}<ˆŽi´/ü~YmŸ]ÆÉöÙ£¡}9=Œ«ãéb{Œ’â~”–Ž“‘á-Ëþ<*ž&c-oõ•Ì¸ÞÕœê»ê8˜†‰{2Œ¢HŠéPéÞÔ=EÐÂõ:N#Òx¯ØOÖì¢Ø>òâ<´@4N×Q_3I³ðB€û‚ûT¡O#fC<{WÉÌ6]z‚%›˜½ŸlìP\Oùd˜ Â8lÍ½Z|]–{šŒšâóø/‡›8(9hŠzäº>qÿ”]øò×—ê©DìF”;«»ŸµÝXWºªoÈ>Š ñðåù‡4ÈJ7^sÙF£pËñ:¨/¹ñÔ—ÚBéf^xä„éVº"Oúk·@wÂsQ)4­aO}ã:%WGÄç_Ù€[ê­*]L?x›Ø¼n ·ôwµU¨j“ù:Îõ_Ôºù+­ëükaÕš^ÈGÃ®c"Ì&›%Š¸}a‚hoŽób`x+7‘á[Ä£db¬ÊòEVgùÁOÐU‹ZWõ•¬¤ú!Šã>`œT-]}%KW?À¬¤ù³,L#Jn”y1½/g¦þÕˆS¤J“
ò¥ò­¨aÝõŠ_(/´JU­òÊƒ×1GºÞÅ0¨7bï2÷å#3ª/«xÅaL.)¬ŠwÖÍn+/Bn~ ·ØBõ«4˜Üªˆ)â@¾Ûd.…dÂpGsEÉNË-Æ¥ŠÊ¥¤¶Ÿ˜Õf[É­~WXëÊŽ«—Ó2(›²^ÜmNítuš:®çÑ;ï/_Ëd„¿KVm0˜FÁ?êÝw”(¿š:ì(XÍ£uQ"SùGªP¢÷ÓQ¬üZ-WÇrŽÎZê¨"D#`+9}÷UPÉb…¯cF½¶òO}%i´Y•ÕTÕ”³é@þ|ëpýÍæöú<Ñ«Dõ´ÍÑJùJ‘5Ès3–%›	ZûDåÙ92ž1Õ˜Uoè°±EƒÌ¡Y&57¯É:Nm¤TÅrË¼ãU\­™Ç+/5Ž¨m@ÓÁœó:#"à×üÅÅ±*-ü¯ÖñÕÇ!±ê«œ¯×ôÚÊ-±šª\¯ÉôÚÊ9ñY4qIbmß\\”
6IQM”3ç*pùeB¿ÓE<’Í(Y±»`4[¥òt4²s9_
§¥å’ïË¬’Ó2ò·N¬(s_/þó³7;ÛÏÖ7v·$j,{1rüçg‹3o->É"&—}CøQÜm£H}ÔŠÇÑ Y¼
c­–¨^½6 P™÷ÛQÈŒû‘™]¶dÅðB¹,þŒé‡7¯l·6£¾â:‰N S1Ú{fñI…©
ˆÖËÇx›¬^ÿª›˜ÁÊ€öÁÅ»	 æëšxû¢mÚ¡æbòÚlóÝÄlS!PŒG²r¢§åLÏsz+LgÕ#‹¦m‡˜æìåÁÖ>ÌÖQ„²µPƒ\;t…êxV&üe>eÆÔ%Q¶.ñÓük Lqp¶Np”yk€–ó! ;þØN“-aK£Œ¥1õ)ýc´t†]Àô[ÇøÏ$E³äUAóã8…N©ŽJÿ–V—4l#(‹jBrrQhç¯þ¦ïƒŽ‚Þì«@ø¤>M2jSÙ£J¦ ŒŠ”wa†¨žæ!ª’lníln©ßß~{QÂe÷ÛtSþÖMªäÍ_1í&õ’K7ÿA#GÜë¡Py¬ÙEàMë¸³ýâŸJzŽ5PRNø…˜i.è Œ›MÝn˜ë†ˆ3ícHŸ¤NÛ7ÕÕä“ö*è2ïv¬?rß°º›ÂÀ“ÝçÚ—ÕÒ©àkŠ±Ë«4å}à½³iÝ 2²©màÄZ½!€âÞl‡ÐkñˆrZüÍ'¦HõÓ÷˜¢X£»Lº³Ügbjr§©Î\À>mUXW>¿~Û¹kÅµÑíì[Q{ÓûtwkcýàPn]Q‘mó>U|yˆtí+gÅ»EÎÚÁbàÁƒ¼Æf”ƒìg‰k'Ê!:Of3Þ¡‹²ïÑ1îÒ1ÕïÓEæÚº:93ì7	–áŽ«ºÜà~…­½õ/¶ö«‡EG¾*Ã‘EŒÍˆÞ5M–,>˜êZ>\\š>×\QÞ¿&Êt‡>Å‚níoïnî¾ÙXß£¦Ž¦•åzN×^Ù8»úŽÒµ´¬-'š°*Pa2+Q•ÅªZS¼ŒéµE¡ª,Q}ás
3O¨g‡Š^y©7«ˆY¥s+±äwr…i¢o^é$“tE›PÝ`Ñ5–ð¸B]T—X˜Ln±0=\I•V®‰:L¸ºýøÙ~õ»/:LxŸSEßy¦5@–ç^—pŒÆ=ÝVñ]žŽ¥«`µ‘)U8~¹¿ÃUðwP[uŽ±y-käPØÄêEÓŒZCÅâÍ ý ¬Þ2o¢ñI½¡KõæÛ‹0vÙÁQ¼µöapï&(™e''±EŠàY:A¦Ño–­®©ç;éŸWó ŒA$ìÁr Ú{Rç}˜ê(ÊÐ½¯ê£öâ•Â®\P·¥›8Ÿ¢ÃOx%$+øý…ä=òŠY“ç0q†»kçêoÐ>†´Â'îL¤gFßeTˆTFlËŸs­×*_.{£9B){ô”½f½"m@MÓQšÏ}‚þ1çJÏè“§WÞíV§l¾æ´ÅÜIî§ì¤8ëQ…4*¾û3Ç}Ë¨s'_}ò9|ÊÞñ«Àèî§Šûò©t…¾ù=án„*½‘oï¸GŠc¥²GÚÅÅÝ÷I¸—*;$¸ï,½$ûâ>®”N	Æqði1ô¯Uv«Æûá@ñî;¥:õ*ûuX¾Eï‡wß-æ™¬ì
—srõ?Š8ù,Hõb¦tOØð ”í 1Ñ'¡3Ò9šFjàíís5¡óVóíVvu£üDÄ·OÞ_îOa)G<ögÑøêßì8N«òµˆ·Uà‰ë›‚z¤£>±øÏ2â›î´Ž:y˜›c.)Šè¨7Ø´ˆ†qƒÌ¤Š,o §/¢Ž©>Û,?šâ¡S/WrêoñÇëþË©ÛtÆÁÓY4©º„ÿöBkB8îDpèn‚u§’]8µ,Q“¸öÑÀú¨GÊ°"JÀ ¬PFXòùÆ×k½|«N˜îTœ®… 5O=6ÒœÜHõ,~¡OLyv!¿ÿ½
«9“Ò„OÕj(Åyë¨^=×½›æ²ŽââCµHUïü§+u¿äo0Ñ}ßGGyœÓëïž®TZ4ýªoUïž¹0¨™R›ƒ
¨½L[çÚÑ½ä<»"0æ•¸5üT'e8Í]A¦YÌ¹éY*§×PÀö«îY¿G•Ã.`éîîy<:8Œ1S¸PœÆÑÀ¦È*K%·ºõÈy·APÚm£|„‹¯õ½¡­v‹Óf¥¶Ð']ób{O›:ŒaNÓñ½äu†~FÙIÚ¼óxÚ¼Ü:;7™ÂÛ¬ÔÆP«º}VÀ˜X8'ü‡-	¦æé…­Ñö²ÌðV„kEÏˆæÐŒ
Ùîm¡>[u’§ ~:L³¼5o)<íÐ@<ñ iî¥,Í±ŠîýEKcH$Köb`Y<KoŒRE#V\q„Í®Öòaá!9Å3'•x-µ¸1#Õô?VˆÙÇÜÂÖÏbþi”­íEkNgNú¹‚&˜ªúãê…+©£°Áèš~± ‚d;áê*}§Ž=Í¿“öû“cÊÓ/ÂeÞY'&§èGñf»aß”,¢¯^x ,@c<È=òØá«7Ë«UåÜ6Å•ºfC=ñ~ô*ŽÏl(“#Ø¸} †ø;åŽäœ4¥ˆ‡bäI„‡Z5öÖb%Þa\‘-¸°7° |1ZÀ¢ÄãÓé(Zí²#¡[iYµErüRÂÁøÒÅs¸êõöxÎÈ˜ÁêMá§ëï	”|Â­€ÍW÷J°†ýtøEì‰æ›á‚]Ü²áæíýÝ­.ÉšÁu&Á:óAuö+ j û+ëàE1aKËØÒ’ÎWÑU}µ´8ùøÚt"³²¨˜$›ê\ƒÇ%+y<úôÀE,1R<Š³h8(‹XÞÞ=ëÊ¹®ä»¡ÓO†)``êEºÉYj¥Ÿ.îì&a!˜7ô“µX²‹T‘Ó¸±™ÅŸ¼~ò•«zŠÞaœ…4×.º~¦)]º0€&q2÷çTNQàè¹m†'’ˆ‹¤Bå,Î{F€þGöü.d¶úOñù>vAˆ’DIJvKx%GS0¨Ø&Mg&ËÌÊxÄ•Ã±µq–éÅn1­ãÄ‘ß9½XÇò§š['Ì[™Â±U|Sb.ƒàb$&Š°µ¾`.šJ5(ÝÐÐ½¶'Ë—nB‰H%†§¬D¼rvØ	åþtº±çGŒtd”r<öŠ<’²r÷²¬¾ÓøÍA@Œöt=_…;Mªs	/nUÙ¦hŒ’¸`M
þæzøˆ›Ý	qþŒ(é0‹òÓÂI3HˆLâZKùPX[½­ wîÓŠ·;ýÚæ¯r;ÉGØÉèG$ñøqå¢o•Y:Aùa†‚›?ã*zÅ†»µ:Jq±;Ð~møÉ8Acãµ‹’¯•çÔà'ïGx¦[ìüxßÄqDcàk¥–d©%SÓ}#¿Eè•¬áþƒW/+uó·zûÈÛ¸æ$K¥7Âlõ&’¿wà7›CLßE@ËÍß  êÓ„,›óOæHÓêm;ß²çWQ­ÅPQ:>˜’BÈ=,|¿ÍÔ¡šx`hk“©¿SOðü±‹ÀP‡Tâ¬ƒ•±H)ãsÍa‰žÆÔªD´‰ž²Û-|×š+4Ûˆ=ßVÓ—!¥©¼Zš¾)ÝŸWËÂ«’¨bÿ"J>ÕJô¯A=9Îâ½äc­;ü}HPå“,UëïCê8Jæe«Z‰üR‹:ëuð×!50Y^ßùË^šAÍ(K€.chRyXgo÷ùîáîþ›ý—ëfötXeyTêÄ!õNŠjGÍ=2#·Ú©C¡ÕìüQöÂ¸{ï
‘†÷êÑˆêì<ÕÒˆ&e…ëâÎŠ6ÂèhaÉ*+ATv› —™Zx'Ú°¥S	Æür ª5†œl ÿL¤õœ4`è÷g&	UWÇu“X×õ0ªÏXÔ³ÌZ/"YÈ“eÑîz1©¢Ç­,"Có6dx>ñe™y¡×®¿å•ÞØ{JÚ»jü34‹C®»äH/—¸¥¶J‚ožÚµèQˆü£F^ù4gÙú0J>…)³µÖââb‡ÿ¿`—Öÿv!r¶ƒ˜”<Öm^…£s"$ÏÍ¡ToÁ°¿Ý•ª|¨?.¢"½åfçÆkH‹Ý–ÙÃÚ¯‹gnt*X¬½¸8G³¡±/ÜÙ^'íç°îÉÂf|–ôãÛ¼òœÕúdÑO‡<*°çÙû÷—–~‰´Ræ–n|õŸ£8Càp°€	œÔ¤?ü½"	.–]âÁŒ]ý…¤¢#HN„é¡¡ÃÉ$¦$O ó("“(‹l2‡Æ…–üc+ˆNö0ZfÔ± Éä³{®ÿDÚ{0j¢[‘.yºùsi‹4Ë6p*±} dwAâÝÍì"$Äæ^ZïªêiJ4èHO|Z2U5úäÑë©h«Ñ"ŠÊ{v_ŽñLö&<½·iø¸ô{>›íA•ûoì©Œ«1Ì»ënÝ[aÓ¤sf…aMÈõÑ\j/<Úh^ZB®9ræ®ñ ° ½¸¼ÝBW=É ÍÞ°‰­9”:kjõ¥E¦KíÅŒJ÷ì²D–¸»‘WÎLÐ¸n3êËí±õGôÞ¨ Æ¨Ma’»Ù4
¨>Gcôm“zò¾.ÍH«Q!õd`©&>Õßxçú3ž½ð±~sè©|¡J¡_ÉØa×$"m²ÎBQ#[IT¸%vJ^K4¦-ò¦ã“*‚ù(¦æß}n(rŒö^®ü|­y‰“t™­›ö§Õ' ‚7»Ï·ö7¶×w`j—vhàz6øñà¬Ù‡Ô*OltñìÙî˜^‰6ýÍ¤æ.5¹_9ØÀ»\²·ý§[—’‹KãÆ›]”üd{OçÙåÑ|)†I§|×‡qT\ý=K,j½Ÿ¿X	oäIûûàªñ¡d{sî¶áC*4†Qò³‘0™æïF Â¡¸Ä¤ü9UÚ1·UŸ ›…ó…•›vŸ®ätÀ<n2-yfAnœ#»¢
¦kj|ÝR¯ªœXy5w©©%šç’ßÿ^«£m¶CvDbÇ¥úN÷FCýh'+!–e(ýb…FQÖOsØÉGôÑù(ïƒÛ9
¯¢Áö¤ãi‘%ÿp‚ï;pÂ¬-5”Ú\~*-MÚ¸CSÓlÚš˜fÐØÄt{Z›˜,ˆöfµ7‰¢¶è´Dr(#t=?Ìf¯í%'§÷˜è†6è9Ì¯‚©GZû1P€SY—xÑ¬oÊ¿Öo	€TW	´ë† ñ¼ÝÄQDT€;NÏ"8àhèñî[#bc¼ñ[Ïü¼ÝŸ%(&ñEŒ[Å¦Œññ7ášFã=ûä 5ŸÙ@]µ×hªæÐñÅ¬çK]çZMc¸ÓK3(òª ³yr Ôuï•oü!óØè4Îe›²ê„îîwd—ò–zd3ÉàôDÄêvSM¾ë:E6…æ2•ªÍÂ™{€ac¸ªs™šÛ'bR¢kÙÓ7ãŠµ¤ï´Ž©Z¦3ŒÇ'Å)Y%<e}cLÂ§µ7#JŽ9&À„JÓñVsÓœ< }@}xç¯ ’VÇ²û³zçžhöìéÒ¼yÊD½™Q}Y½0Ÿf¾ºõxf¶ÄXK˜¢	ü@mPà8.ú§AKÐ&	àÜa“L7hºýÃ4Š‹Ó=<îíÕK:Š³¼ØŒõm`K.`Ô´–Ý‰¼/]†eCË¦©øµí+¥g$ˆ¨@X‰²…{µýZ‰¸iêæ·ÜX—†uª_n…v‘žë'0ßýdù,JË!-Úb¦úa'(æBò€ÎÜO>tQÝÏ‡ÀYÈ=-^vßðS)AòçÃ6„×}¢Ðì]·ÕÞtK†Ý ¥"†P63p0îÄÕŠöÔzv´D€çâä,ôÁG[H‰€ñËPË0Êa*Õ6‚:õX+ÒÉ§Gµá»´¼8×§QÛ%Ì. ËŸ£?Ãy&¤ãÅi–~ 6[µ8–öôvWb$«i6N§ô(BÊHÉøìê/Ãd‘öŸï@—¾½`³sÙétZs²ç)Ž1d
Zi?Í²˜Þ£äçf@]‘&9Ûì¼pÿ™¸ÂŠ„	N$Hê I]4ÎÑÛkH‘¢jÃ] 8ôÚHËÀ	g€áÏÏÒáW†åµ¨ÐÚÙD81Úl‡ìAáãè,ÍæÉ´Hàä é„uªµÅ´´¶ä1j_=FµHt%ÓyòpxW%pfÿ/dL‡Hb"bZ`$•„}t¼ü'Fçß¼½!@ör?Œ‰e( ´s¶Ù0‘ `üp;Oš
‘^%h.p±Þa à<v† Ñ ­Ù’~üàBiP_µxÎ·Y­¡¸ÙÙoGíŠ¥<B%¦i@Ÿ [ÑšË”èù¥ÊúNO+—¥;Ãa7HŽ# &;ÎïNŸ:š¼ÄaàW&ÏýíD]ïº›€zª4VP¨	¤£îy_Ê»|øAqQ¬\ˆˆ,uîÓø]è-‹‡ÑGàkÝý§=ÒTšÇ£¤”…bË=’P‘Ìpéïƒ]ÆfëèQGï©Ï‹Gú*\U¦—)U†Z	LŠó"M‡ï“¢5²ª´{••èÔ¬(cYêºšû‡]‘ScU?:dáÕ´Îý‹©
‡v"dNë½Eùqs ¢%ë{Û-âoWf+Ÿ<z–¦'Ã˜lóe$‡lW»ü{“:ùë'yD#¿S.g€‚kxœ"‹õ˜HÂš¬oc˜NŒ!ÂôI‘@½*[Øä()¹KBcÓc”=³"}5…¯¥É¥þƒâb’zô7²âð»ô>/ÐÑrùú|Þ%,wîÀªÍ’ ž0fÞZë´(&y¯Û4æÓ	:Û>¡‹„w‡(fÊ©¬©²ÏyüLó¡ðü,‰?<¦Ý/ÖNðþ2…§…,ÅK>×Neb–Hk­7GÃhü>¼`ðµÖ8…`„¥1°ÌÇ@ªí6/õ¤,]2&ãxÁ{™)Vq	µV¨OF *êÅ¼áÎ½¼ª/±“ëšA8v\š||:šP°!<dAiºOŒp†Éžêve¹Ã±´y§3ìzîèËÞƒ˜†¡Ä‰Ü[^ñº<+Ó¯c_œdq<Ö7{õëÚRW»Aé¶B>ùVò8ÍoœÑéØ-°ž9Ã­ÓÀó\G/ˆ2¿ÀK{*åðŽØ¡|§ds©á•‰é’a§ý{Âã¯¢Lš†Üfr‚›TSKh®þ6NF©vh±ÄTW ®Þ†øyÎ¾½E2ù©­Ø>_[âTU$ØÀPñÙ½¬>PV!1–å"9>ºµ
d¯gGL~UÆ2à¸p4ŒkIM¼á›è&
‘•‰¿‘ˆR¢ï²ôŠú²PÑ{Õ¤â_5®<ðÒ“4¡Nøœ3á|ÃÔzÝbbäiai§Ó	#0~_¤´ö ŠÒÌQrµï›\¬ÄÃT)…KÚCÉåª·EEó»L-Õ^]Y*ÇKYáƒ+ªíõÜh¯žMOðpZ¿(Ñ™(s=†VœŸâzOFE4³…ýŒqÄÔ”N=ká_ôÛ™_ì!+”¹÷S]@z(þ+4†½–4†=>4ˆI—9ér#)©SþÆüÑÏ¯ùÙI‚p0þ)ýw”	Á¬õ¼u”Þ3Øq2¤G(?Åß“ôãZk‘,’å{ðŸ¯JÙÞ£yó ¹Çšë¡ûÕITœú/h»;À<ô£	šˆÃ²ûyµ²Ô; L‹ý1§k@Šý,Ô`­õ|i™üx¶<Z$÷N;‹K£…WÂÃÒJçáý}`Áïá,wî/.?þàÁd©óÃÊòÂÊä€ä^¡óÃðmee8€?Þã¿W:÷@Ù•ÎÊ=²ô€f#ôSgéÇe²Â*"+¾6O@ëÎ²SRQ  —ªù€¡8å’«©%Þl£ÊýYš3‰¨ˆ×Kb‚\ÇIÆ|^à}:u;’Ï“È/reÏÌîö…X¶"®õÍp¨˜úê¯ØÂæ:F—Þ9ÃÕÉvÉƒ¯9Û)œŒÊP¡à,í_ý;™¤8É	Æé‰2’jš~DMguºº6u,š¢Ò—åheËÖ5u¹ša	k
_ý[
ç7XtÚKœ¡)WG¹ÆQJã±3½†™§ËI
íG{E½x¹ÇÌøJíbØ@qáP.®ÒIAÜœqCtþéc 	kfüð3…àT®þ:b0Â]ùÔöýÙñPH„fon&cã&;j²l®AIÛ®2 1 Çý!=÷¡C5² …Ê·› £ï¢€…”¨ÐkˆAÚ–¯Ü2ú±W©Á-ñðëq„© 2]HÔ@ô’nXº‰Cn‹£iq¢žaŸ‹Ñ/o¢%¿eY\ºü°?T}¦0Eùpm¦€¡\K“i&}ŸYvø^aåS+ì8¿»¾†®ªÈ±î«at¨]&¢Ê%ÛªH²4úÅ%N¡r&as+Aâ·üZáNl¦ÖV{ak±Êò_Îµ+–Ä‰ßþk„=}pR¯ƒÞÿ”–§F(½‰'Û
\r•ÀßXŒ“ CÀ»tÂíò}aD\UPÕª1P\0—Ö"H­«Š€ ;EzÀ,˜±OËb1ãb»]²‘Å }½o æ"ëÀòÁTPDP<šLJIÎÒ/NÆø9ìlm©›‡žz7¨ÍAávòÇxÕ>íNÑ“â”2ªë%÷êbbdo×=L«—/m--¯Ü»ÿÀÕÆ´°A0»À=f>ÕÖg ƒ®øÆ(”Fç/h ;[7.­½àpx	ËµF@ó9«•°¸¡Ý	ÄíôXƒöYsÆ»,4ÀOÛá‡o{•P?h‡Fò¶™9ÃQ‡X†{¾tš3^/„ÓÍ?RÃµýñ¼+ Cµ½1Œ¼uÉx¡”âÛó€õaz¥å)2;fŽ"¾ÁzÑ£@g‡ ly&m‡I%Ã4ÍJYÂ*qû#{4%ä¾ùq<Ä¡W?m­[%â5#™zšdÔ—µM`+$ÅfÚ·S|Í“Ý­ÃíÝ—[ûó™9¦^¦-“Ý†Õ{(o	 £&·à·ê“hÁYnze¥f6Ö–Öá H?¥ûÇv§äi ¡?Œ£1 ¥{Q ’ÝiÑÖ¶¥ÕÝäµEùúÎÂ–œ´qùÏIüa®×r‘à~ŠèeøË±Ë™D„eå\_Àñ“å¥?›Óu/Ô(ýð¸Aná:Šrè<³
jãv)^ˆ†°ßçÉx˜…9§äú1iQ'üŒñ~wõ4ñÄGxëã,Ú+Nä™”]78yäL6°2¼ýmív£­ýv#Ý°;³63Â¹?·@í·»vŸ6Ú×pfµôyƒË—Îì·»fŸ(È×h_£y|>ñ(¾Fó¨øÍƒVúëöäÿ¹GóøŒ¢q|¿QŽåkøíCE7 •˜	þƒÃ9@êIdó?î6½=Q½Ö£îÍñŽÂ:}'ßHˆÃk9$×sèð:ÛÆ½qÜÚTŸ þE–AÛmw¹‘ñQv±^·äd–PÊö¨¹ŸÉ8H·H¾H·½~_ÆJxâÝâúøBÍæÓçµ´¿íCôŽÎ©Xy›á…Êo9Y…J—
ÃQM³£#Ü'.¢ƒç-]mxõ9}‹×À„eX»`/ÅëÚÿ¡(SÊ˜Ù@…îÜÔû"c¬Aö{öF7“¨Ú¢ÅÚê¡†8VùüÕ‡ñ–Ó 2Œ,¤a¡ª/Ó4¾yG¢i‘V°”Ñ?›3| ºzÔŒëë‹ðgÏÙÜ ï²!@ù~|‚ƒTó“MsxºŠ *ßÏ˜â±±d•÷0© ç	NË†±ØŽÏ8«¢Él–Í1q.^-¦['i†NÌ9Ì1Î}Äxv›hiŽ:‡vLk½³Nö2tMÅ·˜?ŠÊÖrtZ+Ï‹A@Éý˜z™¤Õâ™øP	vû %“j×õ—îî*€!ÖPÃPJ`e8êYHŽƒ?¹ žûÊ”]VÊ¾úŠc'¯þÚ'ý”÷ò0¦O¨±·èîªVúT²ì°V\ö8¨
áWÏžeÑ`Ê¯œ§9{ÙBÆ`¬>¸öMu•; ¬ÆË9‹ ;¦øñF?¶U2ÏÆòÈzqböíSÄZÐ»ý€2Æ’õ/«L“â
OõjWc¸ëv@ˆ}—BPyÂÍôÃx˜FÒ‹Î=$ì„METw1ààžêg†PW—×¼nªf˜ãÐ_NêßYî<g0r†Š–fÞ9ê÷ãI±Öê|æç	þñ2Ùöð€u¹„SÝ° TuH!>ç;¯_»”l¿Álsž°I	¯§Y:b;	ËÌntƒÛÊs)‹MX1,Ú¯¥ŽQýO˜®| é1aÕ»îíüöÇt Pß«2­×sÔŸE2žZ_Dâ¼GÚç¦@>×e5”†|{*Ñ#/¦èž¦MËW’×sNŒ*’ÆzèV9•ÀÜˆ^¯‘	¬¼Ê§Ô§ ÂT„Oƒ«b×Ót:\õ[y¢àiá<H}>Ç>Æª4¶©éÐ5Ö™ªàÁrö¥>XÉí„ÖX—Îy5®±JK<^•«1lú*GÔ`ë[+6ómQ€µ~CØÈl6'˜¬6$,1Eÿh0@ë`†LâP³µx¾u¸þfs{}n^ |; AÐ¾ÿÞ­wRZø4}t³,0£þ†ãªƒiíKbü=iÁú21AÎ‡Ë˜Óš/½MX»a7¨Ý“´Z?¡ÕâB3wedµRQ‹œØ(­ˆYÌ)ìò‹ ùú0ÑZqùÏ`îgEF406íŽãö\§H·v…m'Ÿ“¢Ý:lÍogÛ:gaÝ~†ÁšÏÈXs×éº«Áy­Ä×ÑI¬\=ðgª?G›’b¸ÚôgªS G›:nv®wXNÖ·Ág9>(/,ª®ùù0ìÈ»úƒCôJ2HÌ‘Bl‡ÜÊµºÉ-2uH ¸€cÚ`¯|g_š$ß9²[We‰Œ½ÕÛ%_ý°^?Z—‡C®ýLò_ÚÒF4ŒÑ€N9gW×YH	ZÖ».æ²ZßÙ½´:î	½gw»Ò¹Óè¤tPÎ×—Ê’ÍT6P™ BH§Ó)«Ÿç BD-jV^W•ÛÊ•ÛÊFŽÇ]w›F
nöòaºtP=’ýâDt0êh€0óQ÷¡ÙQKµi¿G¿š“cåêœ{66î#DsÒ³±õý£Z8˜Ûvx0®õKCŸð;‹ª]úÁdê‚3ù
è«ŒÅùÞÊjšJêt»iiNz›S¥ãvMC+¶	Ô3T}jq°ûXßV°ëd€Wxõ/‡›·QÇ7 (m¶¸5æ±qñ™`ªros'@õ	¡Fîov”;“&àc8™4[MË©bÆJf(ãMÞM•ñµ9ØÇEÝðÅtF}Ÿ÷Ð&EÊ5MgÊ¥A¶çé 2mfO	z‹%„+7áá°t†Ø]Z4–Œ€Yj©„É˜¶-dÓÐ1Ÿ\Èm€.CCyû¼UdÂsPª2i†|¨˜AþäÃ€vä<ä5Î¥·\£ùòæÀ‰te¯Ì`2$'©˜AþäAWrê\™Œsf/Ðh²8‚qNU+åÚ%P7Ÿ©òJÜ3U®ÞÈ©re2N•½@£©RåÎùª‰¼fl2Í&C¹9ù“ÁùºT¢8_N3’s—ªMßk®€…»Ëb8±jòJJ˜õ-×Ÿ^w'÷´°$6õ?å°Ø5V4ûù¼grñn‚XÂÒ pÄfUÄÇIÁÅ4ÒÕîé½L™*â¡•Û±l
f"}Záx·Õ^§dÙÙ$ÍæÈw¶]Kœ‰zÆ”É²ks{ÃÞÕ„>¯Ø2Fúh MiSð§ÌóÚuõlevKÀ’.8rf•x-/ï*+®ò#¿Û6g¹™0xwh<qc›ì¤-Tw¿ÐÔèß10‚¯[FIŸõ–)UÆ¨‹}›ÆDµý$¹5×®šð>›í&øuÓ)©©%˜9ÄA=¶’MýÁm+æ¶yO/£èD+s}µA0UËý”aÒl²5¿Á5Ù9ˆ†ÿ{]Ä6ûØ¢Ü“yƒŒtlb?óvÝÄM1TSÞb8­CV+}W¢g9­³šYfUc<¸í²n: ƒÉ` ªÁÖ`®þž¡l–¯S…Ö‡QÕÍˆÑ ÊGLÆ> MG†p8Aå¬¹PÃx|RœÖmº&Uˆ0(1T&].‘ÑtòTê;ŸMg|,ƒ¢¢–]41Þ+¬§q40¢±"«À™¼øµƒ‚™uPÎèúE’….V„-ª”°Ú-NCKh!vèïˆÅ00úûþbwÙ~Ññ¤óR—ÃÁNÏ]zèRÇæè:]ô^ÞÝu‡¸ÌÚÕ%]”|7Ýâ‚`W·¶Ö7ï¦/!—fº<÷núÅ%šNhªˆMo½cªxð¾½_ìÒó–;Óz´ÎØ[ð>31uü¹Z¥ƒsµAÀî€äÎ	ÿQAš:ù±­S…ÎPƒïEkìøUS—Å'ûÈÅð6k?1è“…‡ux+q<Þ’$ýæÎ…˜c1©­°c”ÊÀ0@`þÚ¢žÂâ)›m1]½_Á±iä>Ëuò4+ÚíhžQÙûÕJëPw–1zÞŽ²¸ÍTÕ¬®h˜üy'³„*Òb}ÝéN«ÒT„÷à2w‰”à·aÎ¬mžW»S*Q3;uèáË‰ALáíòÙ?ÐÚkP‘ðÉ©*²„™iz…úçÇpÝî²å^¹AÁð3kj4)hš=kén—<ŽÙ!¬HÉI\°‹D†1ìò$‡¯¥Øë,·êyab+5ÑHó´ž5ÒÆþõ“gÍ³{Döàõ·KCµàê>Âbî`²ÜÁj«Ê»T‚¼îsi{«´=½Î,Íã¦jHFp—ŽõãU;ýÉ®ïü¦”‚P2pÇUÍŠ«áüº+‹!Æãer‰ÿ(9¯27ÆÃôC˜H*.`²˜qú!‹&Î Ù˜.¤é‹'#ßÇHâÐ’épq±Gÿó:ÔE“šJ(±nYÓšOö[V·z|èV*/sÔ'§Â† eef:ÁîxËcØÒ:M6S^÷ÀÝª~Ž>¢ç‚²Û+ØíîŒ•U¬C(Ôå0
ºIC'£B†-ÄÖ<ÿ  ÿÿì}KsÛH–î_Is<=TH=lW»5åRÈ’ìR—m©$¹«;|eˆ%”I€€¶<j­f1ÿáî*fÑQQ«wÓË«?vÏÉL 	 $eQ¶3º]"‰LäãäyåÉïÜüt8Y¼ì¢R’-—¤¥Â”âÚÜ  FÓm)):®™òG\$A%Ài'Î6ió ×Y™Û°³EÝ§×Ë¶uÔÍ¦e,·€|fÛÿvˆêë`·cäÈfÑ1n‰:ñYkÒ‹³¨;â:xÁ‚y™€ÝQSÕ(©–Ùös$ƒ[´ßUì¢î€ëØô·aÜ¸ó)(©šúDVµÜ˜7UÜ"FP;RYÔ]q-Û2ød	dÄbM#M Õr3n‚ë ’yœ„äçSë«
ÂÓÌ¬ã‚G*\~
û\„|"›º¼ceXÀ²Ÿ0'´Ê
mº°Í… ïæ'ŸZ$RŒ;”l™aÊ…|X-u DWdÑT*“o”EÀdågÀ4î]ÅF¬'P›©`'Nç5Æ¢W$VÕÕ­ÃjÔzXLŠ¶½©|hXL SNÐtPp$ËÎ¦ÓRn¡™WK‚rÏF	¦RcO7W¹™)Ì¥Òà²Ö¦\ÞiÞ]Ï4YMø‚Ú`³¡L«WbÑcvNËJ¦ë…¸\VVØQ/ŽðþHÄ³(²dEéÙð©úû ìGï»	oâ8"‹Œ4[@O«åÄ?óÞ¸CôÎY$
m´ð¨ÕRÒ/ÖðºÜz÷AùÂFå¾ÝƒU5¿F51[‚QOµÕÕ›v2,Ü&Š[Ó`Lä®r«$ÆD#µ'›+/2?EC‰”!é¯…ƒ Qk1ÖÚÙó“ÇüóÞpÄlìÇ#/DenÄ3Êû	Þ­3¡oÉÏ–/ºôq^(96Ô"0×ûþÐO}„]'Wdˆ½Þäq¨pBQt˜ïÍZÁ˜M*ŒþüQûÄ’¡¹7èHKsÝTÞo@‚¼ú­uÉDÖ ³*þóÂtÒ›nœ»1ŒW¶Ï²-ÕdtÐ‡â¿–	¡É‘ Eòé\ŠðoBŠˆÅ››± WØ~ÓUQ,™4?hüÐ<^äªßóæ—oµ—zÕ/*éH+Ï˜³“þ|ú,8õz‘_KjHãYF±oÏ—´Ór¼/ »Í\C_=-aáÕ®9ŒËx£¸sOHÍŠK0õõÕK=Œ¶Þ…•A|/ÖÔÓÁ¶j²‹~ÿ”áófÐK~Ñr]J†ÑZ'¯97GþwvŽâQ“„
£‘³aý '^èm°WÆgQìÅAdng–”
#3Ž6ÉBž›¯ðÉ#ý`‘“.äv-i®žµ›žv˜WÝ€-Û2/¼ ZÔõa†,	Ø«™²$äÒÅ2‡<	Ê–”xBù7Ër7.`®53(5]B%C=®.ë{`%&I¤ÍúÜ,¿Â“ëNà%8 Á®NÀ:x¶bÌOEl-JëÈ?…ö:?ˆ=#GoûñÕßO}?ñâ”ðTºŸ:‚Év?tõË	hSæv¢¬‹é‰ßÚ<z«(‚BŒJ/¿¡mnïÛ2ÐÔ1_´tYd+¹a+÷×VÜ~4k÷H¸ŒöÎüÞÛ“èÜfñgü~…]æŠC7Áòéû	_›GN+
z¸rV);á6Å©ŠQQª\ÚYAÕ™*•‹iYÆÇœIU]Î×Kæi-Ö t0JŒâùh÷ùò÷Ù0èú.¨=,N§”s"-¿Z½
?{¸n*i‰!`¸‚ñ5=,ú LÖeÑÔ$9XìüÏm`¢ÔE4Áü[Éÿm££QgÎ¤¬£ÙRz»NZdóÕW4Af¹¡XVqMn«æKUdõ_ÎÄRåA¹‘}9_(KÕUt#h–…÷Å€]YvúxèéùHØ”_°'óò,HRÏkóD½²Ëàp›%FÛ`…³¸!ü¤mÚ¿ Oæ]¬¡O¢ 	Ð\ëÄ½Î­ìfU39ý©#ÛÕ¨øšÁíî/¸]Ïë{LQÒK¼äc¡Ûóÿº‚BgtCX’h(IsCH²˜td$»AÁcöŸòwYÏ*ë=±¿p+Ž½Ý áÿ e¹j'Ùdå§»?EAØÆòöz•zó©*Éè@©|x—¬¦ÀTš>.œùÔ0`[w.F
ÖnîÞùœ˜ùD¦§HQªtI¨R%IB|Ï«r-sú…ò»Ì§ƒEÉÍ?•0gõÜhâ'³ÈHMÌ£ÛD¦û(!Œx®¾>=¸3v‘n2Å®š;<j®¼/A¨9îÆ,R|cv1=îéµq_s€ŽwqbpÑðÔãñ7yÍvøÎÞ$ŽA–¼LüøeÐgØLg]åÛÍî$èÿG­Þ~:ˆ£A ¦Ô#þ)é€´Ûþ¦	VãÝ)¿¡ŒQ*Ú®i¿èÿ<ñ“QD½þ(³Ï™ïºû?k¤žDÜQú“y¼aâÿGíYåÑ.ð,_LÚ¾‘=e,¼‰4ž[xvõXc+OvþÜª„ñ;aLÝ$‚Ý¬dúaÞÜ‘7ä¶ús/¯~ëM†^‹ýíoµúmcJTR0¾zOxüÕÁÁVI½ÂÀ/k“Sy^?WOa]£˜mƒ yêÖúÐLcÝœ×díüÏ¥2/™Ç"ðÉÓÌº­‚¹GŽÙ3‘í¥nÓŒý°j×6pÞ6LuiÏ{Úƒ7î&©—NÑãh§´t²»$MÙúÆ÷Æã{öoaCÕ3gó‹cÿ'à‡óxñ¡ß›$æ+_j	†;œ‚þ£gRÁm;’Y·h‘C˜0ç[ßC/ãô,©äËÉ_c.JFäà"R8¥ÉëYøEÍGîz‹°î½‡ŽqÑ Qmà*€øô•$¶~>4ù´¶ù±’Œf¼¯f´çšsÊÙÏñtÑoâ€Ët~eÎU!w[ÕÓšõÜ|¼§ï"f±}zÙ›Û/p¡rùf‡¿ôwû~˜ÓŸøD¥?tØâÞÐ_¿Íd˜ñÞ&Ã’0ùìè„WN…ø7‘ÿ"èï6“_&o˜üJ*Å|ÉOûeõÛu2É_Öüt’GD‚všÿ;ŽÞŽ,A½ŠS|®´+ÌyÄ-4¨9åÌW—tÊ)Ê‘ÔÎÅYRO¸¢Æà«_5Ùö²•ªŸxÊ_\çž£0J;«ÝÆN&èÑÈÙØ‹=V7õ —Ñh­ù	oêoL˜S0LÑêH4‹¿­ðsƒÄ0æ±…À5¿ öú„kÛŒ§2T£Àò6t‘¶®èÌ5~¿eM€¨,ÿyb
Ï5ù¼jGÀoð(¡0xTD1^c>\ntè½?â”i`?brxíø1ÓÀšCWõè%~zÄ-1íðrÊÙÔÅCkÅw×æ}O”÷rûi`õÊm²–šÀU‰qQäg	ºÈÚÜkU¶XíÚÉÃ’±ÅË7¦i3»Í«dw/NoWûþby¨¨o©þ0ÃÉ²õz›äÈ%öŒt±^‘è&¥ãÐ 9Û~o˜©Œ<4Í•Ò//Fø¶d„ß½ÙÝnÊe\Utq—ºQ¶½8öO½°±¤$/º]ýéucu ®kž
B˜@]\gK¥?Ï·TœÚ{ê ÿ‡_ÜæÎ5l d¢þyp€R°d¢?¹b‡ôN‚aÐ÷ %øŸ>¼mÝìYfˆ2®¸YkÑ=YS¦Ÿìwc¤O…eÕÕhcÜ«Q%³ŸìViÎA¹ƒÝº&ý6Ïq@ ‚?µä¸Z;Æ3Ô>ðA»)<}âÜ,èyiðÎäjFˆÒéë;â’duËýncŒRöæ™"•ì!æUvd™Ò­ácµ>¢Ñ"¤>¦s®—KÉw«6c;º¶êõ-£'D9÷slIg{Äde½;Û~Âµv(mOCk¹2Jkí@øóå¯%Á–•ö­?£tHJ^ŠÈx>{À^ØÙYyþ|å¯P]µÇ˜³ö7`í0¤›VÂô”Œ|ŽYÈÏDÀÅÓÃ~7ƒ´Ýê´œáò¼ª*˜ï9ê¡¦ûæî¯÷jýõåJö÷šò÷êëË7Îv6XßáX6Bn@,®ü€X¸Y’‘+M k Üƒ†©±¸I=&ªàËÎª%YëÊÓã)¸[|ßyÈÎàÿJJYÉÍì0®–8£õk–¯ËÇP-ùr*ç•¼ÙíyñVÚ^]r3á¢´^jâöµóâ†HÉ%?iÒøu¸Ü7Iž&óDÑ†kw[Æ¡Æ€›£éç2*<®¾ŽQ‘—œø ˆÎ¹f†@Æ¢úU‚_Eîo#¯\UÜ,…ÕQì·ú[Ô’­]
z·Œ/vš)+Â3žÄã¡¢0ÊÏ°]{¬–Â³“·Ã?‘[1»_ÔB£µun(ÍR	yF¢3ßütjç-k9C90óË…ªŠej/G9Þñ>–+[L
•8‘Ái°xéÕ¯×Û·ëšÙ'Áh®Ó:+MŽ¼óÎ{n%Æ“°‡*¹dLkió ¸B™ƒ|þ¤æ¤Éí¹uW_ÿ²åVš± Cê”‘…—"uÎÖ®Q®ÆlC9"‚%rö/>6”"…íK®”½OÉó—ã‡F¯ÎD˜ˆ8ÊÛoFˆ)^…ßýN†—H_úk£©iŠ?ÛùËtmd^Þ%ÒªQ3eÄ|…KÞ¸'-QÌ‘´˜oÚwÕ€ö
 ‚WÔèÞÂÚr\hÕèœ±ÇTŠ:3qó1ø BD¦ëµ÷‘-rJj
Ù&ãY”òik†’Ý´è}_úS›Kòh†¬u’3ªmÈ*„"¥ÕZdWº-’KõTX{u¤$ÞÒ±|†M•ôÅÕ0cz]¾mÅ*Åýd sùü'NêÚ`mWù¸¤žÅî×¢üÀÇâ’ù_f!q±:×Jâd§’ù–x­Éé|odñÃ7IÐ×!i:3ï|Æ<Í_d;¸4vÃpÝ-ÿ±~é-ûIPG©ÝˆÓVWº¹Ô6ß“;‰F;GÄ©~x8ÕÒ>¿@ÎÚè†kD®n[™áý<æS§ ?o{q¿¥Avæß7Áq–-_/ˆ³x‰ÁY>g2«Ó -ÀËŒ|t…í†aÒã£G^
_fnõ¼>üÍ.m@7fiÂ£-¤~{úf0¥ÏôõgÁ¡¶£N×°ªoeºô•	dZÐT„iM´`nÏ†êá3bí}}cîå8Sî_ý§-Ã=.ÞŽ¯~Ã½Âú‚%°ß›ÄžÂÏàgç+A ïžo°í­ƒã­«ÿ¾ú¯}öø%;8Ü=Ú}±½·õŒ­?è®é”«ÞO¬[WìyÍ’€ýÊ|Eb÷‰Ë9?¡@P;vÍ<Aøä16¯FúàV’²ŽÏ¾õ? ƒÍù¨–ÓrŠ§xóc< ]€bàQŒ«ÅG¢[~âð=šÚ–fï-jÎ¾šªô´X¢2ÌÐ¬ýT¤‚µ¯[aÄ¬~¤þ¢ø±vôÐô@¥ƒœ¢eÿ$u»'!¦¤úg¸€\[?ˆG²f>Ø·0f¹å€?Ê•0¹g[Ë,k†L£‘k9ñÙ}r‚‡?¹`òúåÃà3¯t?»O¹GÉ™`ê…-nÄ
I[;9©(Ex.ÌÁšÅîÛ{q´Í~o´‡,p¼&Š@'x:a6û­Ø´X¤ ³tv›9O&àû¬(„÷gà?Ñ€þ“7bmƒß»ñB[ÚÆ¥.ªžÆ¬÷ièèæ"šfŽ=b/ø´Ë"Ùj “LQJâßúä+ez_o8(¡á6yåæ±vrõÄ“ Ûâ —8aÁ…ñÎ•jÖhŽð-kaA“2«JëUÉšxÃ|[ÅìÊºyVö|ë˜=Ù{ñ…›ñò‘¸™°Ÿ¿ð³OƒŸeÞ/-
GÛÚÞÚùÂÒxùx,ý€_xÚ'ÃÓ„[÷Skù±Tœ;åÏÍ/·„î~=ôTU(Î¬ÞÊI¬´%²0own™í{u´X
?Åü×p¶†3˜À¿ º;!ÂîbY7¸X.Äñ¨ðwõ oS8»°šáYÕZ½=]wxM“>ƒæ3«/%¸ó2Š1ÊS’Jµ5C¸ÁV‘ˆ5õ¼!†›ñË”|]xÄ¹áPÞí¢ŸÙ‚ñìñHdxñÚY-§@³!ôìŒEY6›ÜRpyü³«?°Í«8t‡Òñ,¯¤|c«(dyüÓü¨MþÎ	ÁÛ(ÆK[}]¹8ô`Õ×M‰}2*ïH>FßÚCcl“=‰HÍ^ò!ì1M§ñ«Rè½÷‚Ì¼r½õìjèŸ,³íýgÏv·÷ö_uï?ÿqgok™I¾ëPî\
jÔ›SØ*wxsâÓ* U!ÂãÈsE¾áJÆÝ¥?‚ÂáJú?áßb¸à%¦­‰#à ^÷Í”Ýº„‰L{g¬íÇ±=·eÖýÖnG˜ÞÊãÉ¯bö.H‚#¥‹(SÐTÛ®¥›vQ£íU…K2$ã&F›ÚLÔÊz<P‰Yè±C¾ |eöùÊÄ¸,ò[Ó@íó
1|½ûÁßJûÞŽ_—¾3‰ËÆâ@2^0eA<j·vüÄÿ	áw8°=óÐpqx›­%;ù¹8KÆ[ú<%€“·Ì…»8òÃ÷tì¢w°‡»vÀuCn 'Â•q¶¶/³¹åc íG›Çb–½Z¤T(mÓ,¥ÂX±pg…Ç±—œÍ(›Ú¾åo2c¢„|¥ñ+iÊ`¡t0NH¹õfËg…6ðC­b*óVá²+9«ê0U…Ùë5K¢†ØïAW>ÙÔºU9¿Ô'rò” Ô'ù|‚¼áá§ù7MBPó–¯75{;5ÒÄìDn$sjqÛ›ç‡füÂòR7öŽö¥%¹”aæ·–^­Ó	ÉÛÍ³6#m¢¿ï ax-šà?p-Ïp{Ìï5…©ùÁÏ/ô´RKˆšQUuç”êë"XKp ýœ©ë v¦N%m<âDû/ÐÌôCN‹€Ñì}ÆÜ{F Ï#²/–åÞŸÈè¬özì­7Î'×W×ïw5HŸ¦ÐÑÊ—Ä èY	Ø›ÇöÂ«ßzšdãÍ‰F¸ùÑ¸×MW3’*úÜËÀ`+äF0”ìâ7NŸHiÜ­¸€tâþú¸W®O,[,É¢ÅYT¦úÔLK$æà…fdR¨s‹ÁŽTõRð…-¨î0SÀŽàrÓ`1èM±T¾ÛlL®0æŽèÊ¶æbP^ÕþýÉoÖ°™’Éi‰ž¹§Ï8â
žq‡ÈlÇâªùKtÁí—Í^Â¶ÒàÝ|‚c\)ÉÌéhÌéè‰Ü	ŒiªØõ÷%ˆzCÌÀõæznð~s]ÐÃX;Sa—¦i!×D¦©ü¯:0MÅ\ u§{íÔUŸzãé+x= ijnqïŒ©¦7Ä˜Ê A-_¼	˜¼ÖŽ¤›DqÚn{Ëì„{Ü=îjë1HÊßŽFc/öÛ'üKmÊî,ÀÀzV›åÛ{HÇO16PÌ ö[5>ˆh†ðÒB‰e+¥ÊKì÷fëO‚s¿ß^³ám°Öª¿è;’“³ó…&1å Šæ>ˆSo,ÆÀÔwj/6‚å‹fú—ú	ÇÁÆ\ $ýÅÈPw¼ðËÈB<Èä@®Ž¹å»Í:³Eo¿yTë®u’[VLJ7ÜÑs/=ëz'I[NÞ,ÿnä·×–Ë“aÍRÒ8ÅÎLö")š#ï€-ã@“tæ{#Lƒx_wn’†P”é ¥¹ÉØ%b iÞ‘ÂUë žòIñáNË…½ÛV-ñÛÎŠ‚°ÝZi-]^´˜è·ã~ä"ßÒ·gà×A0Îá[œ”c÷òÍ»_þÈ½aŸÖ5«?hÖž9^OÁ~®äd\ØåL 
Œ³´â¥Î±„r-4‚œML`¬Î”âw¡Û‡Ê3‡ó>]þ«}Nœ vó"JÊ+Þñªà]¤½•)hÖY%tÊú–
Öz|÷"S±¾y	[Zê*âÖÜôä¶|ßÇX	²¤úÜåJßz™îºX	ùú…íÞ«l‹K¾’QMÓÀ™hKV¤¸>ËE8Š«žF«
½6?’ÌõZ=õÐI–´ÚÊ‚"ÎèoVÀUVGiÅvA¤(µ«"¤þÉë$¤gëWNæÕ{Ë5Ó¬\ëM•¢¸¶nõöÊWd^*¨nc¶@g/Š;=+î»E›þdÿpw{ëèx™Ü‘éE¡æËPîÂHaŠb½SÒî#^“)Ê›<”õîÅ`n7föš~‹¦(ÇéºvCmF\˜çØ¯‹+Í÷‚OQ\W}Š’Ñi²´·‚H5›Ü*ŠkRÝXÝ/2¬\<¢°l:d¹z}‰Ø²kF^àE(%TnKQ÷´îRUv¬°IÝë„§¨¬.ßÚ*Êe$Q2)w­r¡C¼ÕU”Æ÷»êo'OQíJXFtyÑPbç’’”bÞ’eþ×ÑŠBÐwËWÔæ§ð:³GØ-y[®Ãbêò#2#hs"Ìñ~Ùxè…¡ßŠëeÍDô—Ëø•¯1èÊœÖ¢Chwddc×Eæc^„ÙJe>x^ËM9‹]/{1ŽgÆø9Ù¦#|N>eL‹ vèËÍ…)hê1†$y°&1~âÔ§×LW'Ð‰y’T…xXÖþ—[
Žk,‡FÚ‘?òÂ&D€‚ÛÓQA}©ûÞ‡ýÁ¾ÿöºÖ:Áâ-¶>JJ;ÈÖ‘
/ò:?ˆõ9à¡c?¾ú»ã™ï'^œ:Ÿ	B×3G0*×#W¿œ˜2½IÜ}–¯£1WXxIÿ2#žþ¥I»0Ä™|½"ZÒÕ«C	ø3‰f.2<yzmëKL3ñàQNÛF—Æé$®ß´ùÖ\}¹áXDÏT¯™8MeÔ}>0[i(Ò¢líq² ?1¬ÒŒgKMüA‚xâÏD”|Î)¦d]¢J<fWò‘h$m~äµÉä¤Äiî@»çûƒ¶W¨EK¬SþíDù­©#CO\-‚ØæÉÅóØä9ŸO‹Wg®A#0«V*ßi÷3¢;]1NÐ§=¯¾¸"M¡/6£l–nXßK9­˜ê¤‚
×†Á:N¸6Q¦pîëÜúÏ¶^¼Ø=\fcšWŸàƒÎ=ù{0š&^|»ÇÙáþÎh‡¯¸™Úí€¶¸Ù³¬71nÆŒ ´'ðY9åM7Øš
fMeŠ§õ'‹	ÚŸ¼>+ä’¶ùú<ÍÙT~4‡ ˜Þy .É¦
‘ø|‹ð–fYš¤=c:™Ý0aë›½0éÅ4t¹™ÅÖ1$º5ÃÀ„L5&ˆ)±6–5Ö5v ,ŽÔ4õmð(Z±=ÏAOÚeX*;Í"kË{Ðò`¥‹´3,kÇ¬Óhwò½òNåÄº“ÏÆ‹>‘²´ÿæ¨×tŸj¡(wŒšlÕþ}4«—8·B=÷RDjû"‚ä7"‚žóD{‹¸‘Õþ}ù¦éö¶ŸÅ%Ø¼sŸµÞr±ƒyø¾ÈùãMÉ\„EÜÇ¥~{ù¦)÷IžE%Ù¢wŸ½6=óŒZ¿‰!™ÔàýýD’³7÷” ŸcxÐ´¨‡ŠÓ™µw/™
½PúflåM,VÔ‘<YLÜÑ$æXg$É<âHŒç·sŒ!AØàÁ¬ÇÈÍ£$4¹h‰&ú´ãê¹ÀÒ¥ŽËØ–ê	µ\b5fÚ5Œ0EžëPË°€E~+{ˆFmì±å¡Z3T:ñÏÇ]ý3õƒª„(ŠeÆn		?Å¬Ô•µ²`¤Uä2+õõËú+Æí@4³¸€ý+H@ôö–Ñ€3oŠ0<T %Fo,5ëMd£;ˆ¯qŸ6€o÷pogÿÇí­ƒã­í­}b$Ÿ3/âËšÈ‘|6:³Fñ]{Ÿ5 ¯Iüž+zÏ¼L‘{·÷þÌKo<&î}EÜû”,ì½0õOc™’îqÖáÓý;rÚØ	÷@‰Ñª¿0»çƒ7fCùçøâ^4‚ôeœ^ý#ìóâ4€,¨å«1p‚½?®¬¯î4ˆÿ -$cõÇZãÑæoµkØ<¬/Ÿñ»êã1;ô‚á{ïkoìáw°T:Æmñ¼j‡k>!9jq(&>UŠÅ;KÓq²±›yÒ=Ù	£N,FÔÅ-jh¢ï¼É0ý³ðéžD© žî$š}¹‡“8¨¾C¢‚ÐÂ,®/ã!{Tñ¿vA]µ-BÂ!‰e³¿ûå™;â¡n/N“‚ô¬Í§
fªµ4]Õ„×5V´Ëó|*ÞdMÝ½ß]¾±IËJ$è#Ù¹ŠY~â¬fÝ 6^Zô%y'ÈÜ®XH(+‰Ÿ:5}ö÷ã·÷_<Ù{ºÌZ’ßrh(.Ý	Æ´!çÊ¥êé¶¡Lˆø†ýíolà',¨Àìo¥ˆT Bé8!ìýhÜvŒÂßwÁF>,þL>F[­xåºšd{(¼t"P@½;6}­!f	'éM†”O>ôjƒÝ½€7taÞïÔ¿t €‘lc2übÔ=ÍaÕ÷(aÕƒ¨7I6ÐD­TùcF“t„>ðîÐ·aÎ„\öPC[ùÄ
w ¬yktìØc“‘Çåë¸¼Ä`,gf¥R2s6oÊîŽ¦0PÁ¨ÀrÂ6ØíÖŠ7V`¿¦“ÄÞCu 3(†ÇqUï°#Þ4’8v¡!!oií×rz2éõ`OLMêùöÛæ}‚(,x{%öáŸööþáÑÊþ`€C]Ú`û¹>5Žàgdƒ1<É}Ï,B>/ÿ€®~a?Oü<ê2dœ,>+ÈÎ0süüDSËÊŸë«š3Æ­ a7…ÑûØSÀc1Ù™Çb¶¥9`°8Í¶ª9<¢n£ÈÐ‡JÀÃ(5*á’_n:Q6%Åîú­V!“8êsjûçÚ3ÝjäO Vw2î*ºðRW;@ÏJŽZk¡”lŸ·»OÝ6%oâŒ™í3´˜-iÁìýðd
È¢öþÁî‹Ãý—Ç»‡?Â¯?~·û×¦ÆÙV–ÊFò†®'Ì’­’¼íDqçÝZ§Ûíê7Ù%Q>$àØßù8ÿ48,*B4€¢ˆ64T¤BÍk‚BÝvuwÉ¥eÛåUÃžƒ~m×®«£ÛóggÔÓjÉVµ—¬![d Aìµäó4[¬P¤¹³6®èŽE"ûÒHÙ&t»ªh÷Œ£h¤;E®y€ÚïÂvQ´n}Öäß.P¿\“«_A†Ñ;´fàÇÒS7ö@>Fð]„D%sç£ßÛÒ6ØŽ}xïÕÿ‰ƒ(a>Âd2Š’% ÑØ{wõkR^óOU>Ås(§	`>Þ¦‚Ïˆ—ÇB*ƒ€¿Õ.–²XÈf‚I4\M·\„ÈùÚÈxËeG«Dt¹Hˆ,¾•ëfö#¥G_Ø»ä'ÙI“ŽÏ\+÷OlùVë›#d6:e“Ü>ðú1ØWì©?[—­wWÙàÙ’îhPÑÏŸïïì>;!k#; ’ŠE8ÂµŒ#ñ–P°÷>˜CÁl}Õ@2#<!ÿm.Blã3ý[*›nà`NLzÊE¤—Àºúe$~’v!PQò©Ê8Å‘ÿ9†8½kKçË¼m<ói”ñ,ªns7›¾¼¸ã½’¨™	˜ÊAÂ˜ñcîà'U••¶½³ŸŠùWí”eñ¸aG‘ŽI®[”Ë¾eú"ÍoÀX+9+ÁÐ}°Þ`Uü°?Ž@æ ¦…Ü«\`©áÒ„hri›Ë›ë&ø4)Lõ“õrùáYî¹¨	¾óLD¢to‚æwæý˜øüßÉØáï°|ïÇµõ{Ó	»[áíÔuöSvŸ”§Sº3{i!D^©‡_ß
¾- ¼d„I8ÿör0qsíÁRW¿¼^BÏhu‰c<.ß´-æ”FÙiñB-]ßO“$:'"Œ¾œ.Y‹ò(+ÇjÞCË……§xâù?\—8j¿ú¿`Øþ€q·²Ø~2#èìøï‚ž¯HÓÅ•Z†i‹Vt‡Û7¶×…˜¸Â?úÍrøµŽýëùqêÿ'ºð•GóÇð<4?ÞßÙ:B×B’-@/Æi¤;À;³›¬Í=ø2Ä×{§ú%ÐmÏ=û0 Ñ8€â³$‚
Rþ ÷óçºÐ‰”_Ä…ÿ…Wÿ.%KŽœo¶›®{”†>3­eúðÈOÏ¢þkì;Ø¹ó+“óïµ¶0‹§Çž_ýrôxœÊ!öxk ’-x+×eÇQ]FÊÒ".úpñ'_†§OS¿‚wÜÈ¤˜,ÎQC¿aþEøŸ´/\ó4E‚CòÄæZn†{Xˆég¹=büÍr¯¤Òƒ¼®€dHµZh²6Q6h"(™PÎüÛœ©lçLÅÄœWO¨Ì\°qèÆÑYô~/üÉï¥ÏaßÛ¨Zî·iB¤fœPTDóÅä3Šîÿt´ÿâÚ'‘*¥e6Y6ÆvlÉÖNpŠSe2•)q-1b¬­ý1+ÓÆÆI“o2rqÑw`—/&£P‹á8º±ÏãöÊÿÚY9]F“Ð¾¯3@4BÑªzZžMW±ÐKžFGú€úBB(Q®‰º#e±ØÑðgP`µC`× 0†·Ù6ÀÔ-hÉ™à~žºÃ‘TÒ„BÇ½j`³­Ó‰£}3<„¨3`†TØØ÷‡À)9à\YaO¢¸Ç³ryÏzg~ï-óhhÜƒ—Â¼÷ÍwoEf¾°ÛÍeSÓFÏ+’:+¸yk¿ûoµ+ŸPòa×”£¢¶û½…žÂ•Çør™Ý[]]½IÕ«ˆ(qÖ>­ÄTw»ÝOT;ÁæžU{Ìª>ü»HßðBHÛ©Ôs0µÁ©ñ'¨wußúô¸fäŽ¢ÌîÐOÓ3î]e›&ŒÍ}ÒQ	Ø0?£à…žr’)h½öR<R”³ô}aòAÃkxwÈ(7{ÁÈÓ¹€Œ˜KlÃ4X;X¥!ã…iØÙüûa¾v	0^ÃH5íW¹([fA8ˆ^5AxC}³ø‡¤;Â¾Y´'V=Q&ÝBºZïAb‘ÈJ%b‡ôçd8²)kCSkªR}†/Mšø`íÚ¦ò Ž±>­ÞÄ‘ÀóÉf·Xä×âôã?t{9+ÎìkXŒSÊ£ò7»Ä˜ÊX¾hÅ]ˆ´õ[ÛštjqQžÖóXr¨V¼­VœK›Y§	õ5-˜Î&@Ý)©Qàä•á0‰gokshqå×CGÂ	ÏÆï,ó”Y¢`çÕË¡ŸŒõÞµ^®`7‰³®ðM“8Ô¼A;õÜD…/#ž„ ïù2oÔ«µ8eä±—ú)"òX{k’F4Í·Ù”˜E¥¦]0A½¸¡
êÅ ÒèàI¨ì­£²·–«ï;÷J'fkÚã3ñæ¿Ö’ÞóÒY?B>Ò«–ù1Ea¢4D_¨—üèçžN®+T[±(ÂPÇâÄ\„jé7¯•NÜ6Pµ¢·#¢šÖ#dW,—f‹€åÕ›LiéÞUÔ›×4N\.’–šöÚéÑ)’!o[±Â	¦½R‘ú¨5gcQ,îñ¢…>õ1§f .Ÿ[zKí¡ˆ!¶xéÅ›‚OgÚdî[¨w/ÐÚìJo–æ[ ¹ã?¿]›ñ-hõºc/@ÎŸUŽb¼kSÔ–Ÿeõì 'ÿYÙ´.ßPŽ¨‹¨Üì¦ÑKÔž¶A{j/™@$¸ÖÎîð¡ow·÷vö­}Y¡+”|¶Y™*¯mVÐzEivŠ¯–7†}q4®Uyö#}÷*Ü$øÌ²BUÓ(yz³ÒTªÎ|D¯–¦~zµLá³W‹Û¯²ÏT-o²#?u½³³þnƒuo(¦öÆ«¥¹¶UöÒ7ìsðê«Åä¦—î³fÍ5rÜ—*.³5·_ÿ*¢;_-Ó‘ia!Ù•dVÐPª…èåWK“9">I{Ž¨ƒÕ@@At ª4WŒF)ï•ÜJ%TPš%—éÞºUZ”%2ˆ‘Ò
Ñ1Ó,Wx^Í™3¼x’ Sš%Ô¢Ñãr]¬„j|MEqUÝc%ŠÖ’Ge”ÂBó¥—ª®ÆwYÒ…§ò¾ï¬­³3üÇ¬`~¥øQ¹ækvrÊ^ŒpãÉ0ñ©Ž»C ’þlû½$‘õûd¡:ë‡F¦dÝ.àí¢ˆ“t~ú„?ùâœÃüˆ½ô.@0åÔñ“þ@¨:–¶u0/§ðçMÑýyûY 1À0’`E {z'Á0H9®ÕËÄg©¹|üì`÷è;ö,ß•ðk®ò M‹ÜP2efHÔP ›ŠÅ”›Š8Ä†ã?Íb¼ko&z~w!‘Ð'iÔ9ôÇÃ¬½·E±Xms6ô›Ù3Z‚™rÚ›Ä1tvK m6ÑRÛHlâÆ÷‡¥Ínà\ó©–mnNßš7E#”®4åŽØlììNi2¬g‡å²²ÂöÇi0
`sô¤•\¹¢˜·Ç±ÿN 4Ñt»Ý.Vl¢`*§¥Mù6¥…fêpõšÏ~³ê5Z›²¥òåÒR‚˜“!NO‡þMx—EpEÃŠ3ÂüH¡¡?ücxKö¶„¶Qö–A°à¥/ùSÁ¿6Ykÿwªî?yâ4-—[bó±ÇL71lsùpèƒUÊ‘×{ÛÌs5Ù0•t˜I>Ì.!
Öpgž¼aÚÆ2‡ëðÝPTO82‡€›®¢žu°÷?–UÕò)QWs|÷¢MPŠP×–ä1QÑ‘
å.âIÕ¡Óä8­:™r‚ÂYçF>ˆ©ÂûåÌ3©(SY<:¿©ãmò¹:ï<à“§~³Fž=ò*ÝµÅŸ¦6‘ƒvHsGlkF«C¬™˜„„&?žIrâºÍÖ9žv;Ž×mØÁM×ðNã»´ùšhYšm¹šÔo2•Sü·3úüï8z[=@³ÍœÐ	Ëó×G"þ¡û?»—¯{«áþ”)ž°<õc?ì>]ýs˜ãa”H\¼¤Œ<$\†—Z>â@ç€´¡,èæ¶¶z•„Ïl|ÒYÇÕƒÿ®JÏ{Ò‹£áðÄ‹‘¼ô>î‹WÚÁ^° Ê³ I¡ÿ½lRåß}ß!Ô-ÿþ7Ø×ßòç>”ÎML;Z6¸ó‰ÒìcùY¶¸ã¥‡¦oÐ$ZiðéÖO®|õQ©¯OÁœšx¸·½q“æŸàö”·</Wòoå+žCÿˆ©ñ|ü$½óz^T™öìê·Ÿ'Aß+^R=Š°½åµ@·„)µ¥d¶z#yÔ;4`IÌ,ÂRE+&±W¿å,dø2·Ý±wÒoÉÄ´;#ÄEÏ’ÃWã¦›#ƒ‹YWãï5nèR
CM¦ˆœÉ[ìI¯:Fî¬ã´ˆË²®¬ÞÅ’/Íò`Ø”¡`+W¸rŸºñX¯–VÛÔ¾aqŒ·zøˆ|@¨\©áqê¿´©96ÕTŸµqJ-€–|OÃIuFt•¯jÑqZÍc#£¸TºŸõ¬ù!2‰úã®aì¼äÜ–Ž?6>Aës•ùÏ¡Ó\6Xº\•æ¾S“Ùä‹aŒ%““gØ…dƒé%qK*<îÒzî¥ñÕo½ÉÐ+úï±ýïLÏï…}x ?Þ§ú¶5œ&®JKË¼—äçô“^•¸¤ù.¦Ö23NCM¬;¶"è§Þ>5Í€º^×étÂ‚êŒ#.±ñ”Î(Àjb­nKÅ#«­›'q0']¹	{$^Á‡ª%áM–%Â}.ëä½ðÝ£.¶ D¬è­Þfß`¼cÝœ&06îÕkqµ´ÍW´­þ¹ã UôAýg¯¹&ï^È_þx÷Z°åš”ÃON¡¢f`ö»§õçý˜
©PÞ‘´bˆYCo7‘`¾ïc|6ˆÂ.Žäøßâf§­Jü±‹lÉŸƒ17ØhÞüè¥ P&(	€Ã±û}ñYÙþhŒªTòJÒÅkÝ+Êí<ç1*Q‚XnmïŠMd”íPµS%2)~ÐM„~ Ö´\GV‚¤}wRêZý¹¸(Nò{h+<”íA®á¯×!óLröûè”èlážÕ­©	l[ÅÀ¶Ò`}z;T×!m…)bwäÉ²(â¼·.ožÉå£<qv¿æo9©\c-á2*7]ÞocVƒÃIyvßÑkm^C÷õE¾sVÄ¥tµ“%Ö’q^ýöÎ¶ìŠ#žÌ±H®PR"U‹}È1D¤Þ¨o'tKŸ¬ÈÌ*‡²À-g¥Èæ#îb:V¡î_È8q³ÃUòjÎà7*7=ÝxJv:°OKe	a©r÷F÷·â"Pá#¿LFÁr¯wÙwbsR86$!¬=JVŠ)2¾*ž+LAvÃIEzç£rÔ…w&¬]M¸v0÷L$2ß¾‘Ù’¸NÊþ­¹isªOQS´‘¥ª±<¤Î3x•89eÀdCL®¨öVŽÊåõûä›Ö ×ÅI5@†~¹Z½RýÃ·[ÇG[?>ß=:Úzº{Dk„x/š)"l É©ÃÖ€]ÿW‹°ÞpÌ¨œÚ+¢XP>µÅ^ìO™e>+¤`RÉ¼Ä–F1BŠXO\CbhLÏ×Mf›ò¤•=››òªæHp¥WçQgòµY&7juÙ—˜o‚¤øþ5˜¬ÅiÀ¸*åM"XèqÉHï·'Gvï*ŸeWÍˆHt0œ$ÙšUçJŠEYud»†l¹#µroÕbWÇòº³‚¢\¨æ¾9§Ú·I­¯§×ñm¶¯¹ä‘5+Ò„÷K!æ¨a£!e	b0H@«»æT5•©P=Í@z*ÖPC®µ¹L"Q@Jõ÷Ã!hžw
MŠR±a~Ôrá@ÀÅûhèÁY™4N“Èû©ápá4ƒÀ™ƒ¦&ŠJ;t«þÍ…Ô¯²5#‡Éd§À‘!ÅÉcÔRÍ*³N±¢ëy¶(Ñj¹ã8ü¨—M%€6÷AM½IœD1Xæ)* Ñ{ßQ«^›fÏ!œÜSBÛ¹úZÿœKzx÷Ó/<î!:Ã_IñŒVÞ*ùÆß·ý ¾T0_Z»ç½á$ˆÕéÜlÑ"^é7C†¶ý/|X)d>L¾m@2›Þ ñ˜ýˆi@šËêE¬9oÂz¦/òø	kS+ÑïmU3ŠÇþ(Bðï.zå„Uk”vHëN lD‰¯€n º†ø¦ÃFöºfEzÙæ§XxXHÑç%D‚­Ç+/ÐÂÑíÏdAÄòÈ}·¤Qìx/wð–·¹kaGÉé^¿i¥ „eô†ØR®ÊÝb®ÚÜ ytñ‡„ÐxeNk_Èõv;ƒ£ð%Wžs¡Fï"ì9¸£2 R§tŠnBPU|i[5³e`i7 Š®™=¾Óø|›É±¹H1²$j ³Æ½¿ÚeÀŸè”áœŠè5…éRÝôú³;‚É÷
I"ç’ùÃÄ'®Á”¶o#ë£Óé´ºÖuÐ÷Ôd9mÖZµÌ•¬š8Í¥²”§®ý¸®z‘³ö£:è­Ú‹Ãft!‚Û# l?èŸX­zN3ÿ°Ë³«‰Jú
£’¾ª]q-œ®P$·3Yê[G?O¼8ô¾·N8
piÓ®¨ñ
b™³£šPÌ 2òÃ‰ÈŸ’ADü«¿#Ø©£=µ10°Ù;Ä cãøê×w!	K1èYb€ŒZ²¥žùve}‡]V^`»E©}¦ ººa“Ù™VŠÙ–`Ã4ª@Z]ßL‰EvêŽÄÕº4º“Y½Ï9ã]Èãb:dñÊÄW/=NuÝq¬G èHÈ4–D'1O6ñ˜?‹?‘$ûƒ OAƒ/ÇÈ¬„õ«ô¿·Õ­tr\¦*±hæ×tÒ¥@´6•\³•4#£-n×6<Ïç¡ï±]1üêõGG_c5àmu†PrcÊÙö~ž];áåÌ<*&Þxõ÷(YŽ\øï8^ý–ò»ù‚h]ž7‘!‚°sÖyuO ùÕúTNÈPäôIcŒÝðTÎ¶éØ‹œ!?ZÊ¬ž.ÁƒQ[«/p#U©	ŒÈØQ›µj¶O…2žø©EGÍjSn™ëÀ¨!šõ©Ê6”Y2k‘Ó(©F¥ô¤'¬„¸ÇbzÌ Êå:`ºµP²ÞéÀ-zSV@™"¦Èé™­DúÞ×%oõ«šòäe¾/Q\5œoxÌû~âáµ%ØŒÛgÑÉO>GÜ`nD!'–aQ‡gQŒŒËÇ[Uµö*r¢&4zŽÎƒgƒ­%eØúŠÑk›¦hÖ›Í!üÍ jA$é”÷‹ oÖaªsØÆý!n¤Æ’!Þ&ã d5å£õ¡Õß·5™.µwTÄëù°õÍÁÎc³UÖßÑ-ß"0Þ0Hp•S Ð-ô§öÖú• Ê% ©*¢ØÓ¯Ôå¨IMÆÃÈëã}K0"¼d=lè‡P¤ MàÉùNªoÐAaá`;â |Ø€^èŸ_ýÒr&ÔS4”Lpm,ž$OŠÇQÀ‰Ïpótˆ¯jçá´Üá\ŠFÑ-ù…L´Ê±Ž2[X.LžF?0çˆ`œA0Ô&òz=œ>juÇýîwe’$N‹æ¡"ŽþÓú0Œ—|ÙuæW?H¼“¡ßTyýá:O×é›£«ôEÍ3]Ãèô‹µe˜q4}"fÏ`JøÕ9ØÁðwVæ~.vW•YxØõ©?ÂT¯#/î‰Ô*°¿özÐÖÐÃN±SðG¥É4âù²6I&W¿ÄA”ÌnŒQ»>t ™CõwÊæmÔ1öRNr$‚Z³®¿9Ûlõ_ô™BU9I 9XvßKÎü~iÃÀ\M§Ž°
»®úÞÕ+¶A"Ì‡lÐM1¤H¹yOÜ*îTšŽ†7òz5,¥¹ûT„ýdîÄÞé>ÔÏî€”xŸ‡'Àdí£Soíás°ÏÇÀ…1¿?™Ó«ß¦ÎDcRèäœ»“Ù¸(Ó¸u‹^Ž9ä ã…xž wÕ^6.>cv¼Uˆ2$A7RV‰ºú¦uÓd˜»¾#`OPëé‚å¾;ôñÏÇöúm.¥:Ø—ù­¥Í.f¸~K^<-<˜9¸×®AàµøGµ>é©lpæ¾ò{ýC$£qìn…b¤?ÄE+j™^LG2ZšÑžXó…H®e°Ã0pZ3ËE…¡ZBø	æ*#‹íék)Ôœ©I´'F—?ÕüÊ]€'Æýêt™å¢ýK5ÈÖ	rg+Ž½„ÛFI4„ÿF¹™…žUm½é"8ÅgbhjÀ„PMã†]è½ó1eQ±ÄŸàpPuñE†3R•Ì‰#¹¢02ôoKÌ	:bðÀ*ìCg’vðâé2ûÓÁSìÿÑŸŸ²ösï¼Ë<¬‹ü¤êã¯.K€(ü÷ú ÙšÖ©÷OÙ£iÅÏZ#féy[;"ïaì{Caªó(BMûafBŒ}ÐÛBé¨Äs@ W<E‰ØÕo˜hH‡£mZâ¦M¹6øé^y÷µÑÿj92wXQÜØ¢ÌèÌÅ"Àm —û¾¨#^‚ä Åq Só¬ öv8­ÏçÎíÜÞÌ¢P‘w°d‚,ÑbÑ0´‚ÐŸ»»3jÎh™ìì… »o7‡všÔP¹@m¶4<»Z¢;lSm5ƒZ	ç-¾ûJâ[ö¨ðoùýBõ<o½ÿ[±u"T/+åb?ç®‡ÜR;àÅbØÕ
}À€Õc óðê_V×Öîß[}í0 kp@Ow«G•ª‹3GÑ¯úªD´™¸Ôâ:ðN@™àÕ™ÑFãàô,¨£÷òÔˆÙQêÊZuuó_ÖUÍõ~¦e½úc~g¥}+T|õ˜A#9öž¿|¶uõßWÿµÏvvÙ³ý§{/ô²›8Æ „}×á^‚W1ì<oØ9Åÿõ·ÿeÍ¿ç=ô~\Ÿ/sÒDV˜âç¥×ìfÈ8åýí Þº±öÕøüGüç5¸¤IRo‡Ð‰XÔn}ØÈ¶P(ŠQIO1©õµ‚©À)ÑsŒNÍ2:î=RßkŽEó†`Bf½C6g×W»çwø	•ðºr>¥¸ññ\è<t'Æ@ºÒ£;êk÷ÑV%Æ§g¾C½{=ÍÁVb RÀ"ÐÎ><j…Q'ûÊTÉ`Ä˜C”Œ&’ÕPÌ¶“F4fƒ8u’·„Ë6ª›ŽË±ŽØ;<µ™“ü¾¨å°•_ŽÝv\|»?œA‚)Ïe!¸JŠÌË‹ž¤ˆ%ºÏ¯ŒcM× ]{:m2'û£Î;ÔÎÑÎõè;YôR2ñŒØî#f»$-§Q’±”uoÊY×ºÞR¬"ßŠg ËrŸèÃÕUW
VL Ju‰ÌõÔIÊ_ûÇâµ(¿Ö:{põïkº ï@õõ5ìÕCMîQƒÅ»ëƒßß=:Þeß¾|q¼{Èþß¯l}uý+·ôœç]¾MšœÐá±M³S:!Ö‰§rœVÕ®_Ôð¼ÆÇs˜‘5aWÿ;õƒ¤ÎQ.0°U~xJçV0ù
^hÐmv‹ÆÍf<÷ÂÏŸÉ.Êƒ!Ê&ñPü1ŒzÿS{-QcXèÏï—†Î]ceÔÎÞ+ïÐÄ;zaôß˜Á9Jt´uª~IàÔIgAŠÂ£ÉÉ(H³ò­~çê²FI¹&ŸŒ
ÐOÚtTÕùìK!–Rxìÿ<	@<T¾–0¡XF~A¨º,UÂò²w»]YyY’@9î°-mÅî»¶8Íòt¬4Ñ—‡Ï¦›PØÓÏ'ßHŸât>CÆÀÚûcÜ2\ªÎme9Ñ‡ÈÒ'S2£[8ZF.ŽçÎ`ªÓg€Ô*5Õ‡deâ`Z9Ÿuœ{•ûUe½%¶É¾>òÞù¥kÓ`ÖÔq³ªž”:sþz¹¯M|ª¬°_Ý.‘Üý!­èq\óÃ_p]D¦kA4µaüõP{ºQ:ò1,›ÿé HŽóÏŒÝbÐzK¯ÓksßY‚SLgdC­Øá¯0B%ÁŒ($ÆŒëÍB½'¥`o™°âå¦Dm¦NSs´VC“Ó'3ä2wÄ.üH¥-,¿žþdÉ~¦TÑÍžÚ(H@J¼!JÍœaUöoÄ·p{©Í
eÎšË»yÖUÓÍ4ap[à¼”rS4r)ÏñÜ¯ŠD£;§z¶÷â»#˜~L#a\Ãâä¹·;d–å1c€ÌkyJS,5ÑÒ’Íjh`&¡ˆ
¢šY¼Õáå?þ{å­rá]È?ÔMB<Ê}¯ï'.ÊŸ‹'WÔ.þ   ÿÿ ö,Sxœì<ÙnÜH’ïþŠt¡a°º«J‡-»»tA-É=šµ-A’³;UL©Øf‘l’jÔú˜Á>4æ¡Ÿóú±È‹™I²Š:¼èÅá£HFDFFFFFDF’¼nè(.Ù1=#ëëë¤3JY6
&,Êã­¨ a‡¼xA¼gD^kÛÎû-Dú5!ÙhÌ&,[¿¡§Y‰·LÑK¶~#þß² m¸²œCŠ_M°òg÷ö™üåvì<N'EHÓ ÎÜ>½-_ý°+Rd,]¿IÒø<Ùó[àã8¦YŽlð·mÚ;££/Eâ6•1€Š#2
i–} ¶Þ9»è§qÆú+‹$é¿&i\D>óû/¯Cr§>Kåjiq‘Lèuÿª¿ “ë>-ò¸³aHmmüÒ$Ÿ³ë¼ çq”÷ÏâÐ'ü	'ö;ë¿²ð	ù‘³N9bEJ£»ß©IaüÒj/©4—MŒF^‹F^;üÄRFŠ	¡é¯Ep“¿í #`’E~LòØ3|Š?à	ŒzÎ&”$4¥¡L3Iâ‚L‚‹”Þý~÷_ñÀâ9±XöƒK“éó]ü§?ŠC’M†üw_‘šTÄ³vVäyYÏPY·Ã`ôeýfL#?dB†·Óf‰Àè_ƒœ9«ùex\b2í¿Ôº‚ª0Ž/Y:TÊ„j’ƒT² Õ­OÃ÷@K“¬?‚éÔ)²<8ŸªÛ,¡#Ö¿î/w,®7œ>¬íÄWQS†äï07—ù €X´cLS©Nkü‘-É!ÊZùÖKKˆËKˆ®µTÞ<J*•>~LÚ‰áe9-P;¨Ím%±¶ jZ> 4a<ÛcuWŸ5ZX ý~ŸlƒLÒ8dÄgä€^Pn±3Âí¼`çÎ¡k²„B ƒ4N2rdG4„)‰Öqk³lH¶4>9] 4™ðê#üw ì¦x.-æx“!Éò4ˆ.z$ß’NVŒF,[üé°4ÓN—¬oË8ðQZßM‹&ÀFEò~²ë$Nsr^DÂ¢ÖwÃ«éY''§=Å´¾“¬öÊÆ{Ïn‡³ÄÓò‰£,''£éh|ÌÒILR~¤oO¡hé(§9ó:îj‰Ã&	,§ômÂ`p´]ó‰‰¹ÖmåÂúépi…ÿ8<xÓÙkÙ:PrxEf?2X´KH÷ „…º/a£"£>•74þŽw€±áq³Å”]ÀP-l9HT{Áò †ˆâ€xÀz
ì¦¥Î OÆ=W•ÎvpN¼ç
žüöy.@»$ey‘Fdi•Ã‰fzS`™ÐXß‘ÎñââÿRðÌ„„g@ûÁùù18G€ôžæã=Ë<6€®áC¯Kú$+ï*˜;tš)ÌBO“[ ôEò-y­ÿY~Õí+²w²¯ØËž"Ö%›d‰5q½5džÐ)jôa|•‰±yÏ&±çÂp ‘Es˜FÓ“S>_$ëö¼€‹µKGcÏ£|Jtä0‰çPñQQ˜C“8Å éWâËŽÀáÁ#´˜j,W%±™À%ÉÇÀ·˜ Gql%,-+‘É¯ƒžNÀ³`#X±&ÄÃIÈ¢å®\ž¬«)8á`šÖæà[=-9ÖˆRæ”¥ûg¿™§ÀpùžWpV‹Aø¼YÎýªƒî9…†æcK%üžß­pâ³œ!Ž¼„å®ÖAc—p˜JA"×PZ¤žñj§VƒWå³Û²#àÇaš.r²n I½È·ÊàÖ°ã§Áo¦ÅZµ†Éh=ò­ÔV0^]³}Û‘¶H¥yò´´rº u=b<zL*ƒvIÃ8ÕÚ®åm?†~,ºˆ\Ž[yL#rÄü1"¼²¥Ê	›íš@	Ù€^uM*¯pf¬çµ…óC+œï-0nm–m¬•ÅÊˆÁb•ÇÇqNC”—¾o5V>C
^6N¾ò¨y‹Ç>ªÝžÏ¥}rÚLhâáÛ0'_aO=˜¶]xkž÷c®**qî_‚W˜Àoˆ€‘W‹“ìæ€»q\ƒßîüL¼íxÂÒQ@Ãn¹² ~OZƒtŒ†õaÎI‘=sB‹®ïùC¥=oNje¹`¹qêÕ¥£òŸžû H¡s?4Mâæ ‚Gé„p°ÂÀØç8µøã8r	$øŒS8b2b!FöõØÁµƒ;ÃhÀcŽ¿çðð1
r‡hÆe šcyø›&ŒÁ&Ó­Ü[ìòøc’°t›fè*|çDµÈDœÌ[êZ°`À4Ž‚„†Çr:Ôjs¦™¯³r
Æqhµa«¯5¦v*} Í3~ÃE>vº¸I1¿ÄAäuz¤Óu†5HˆÒmÁÎž_á Í’jèŠ¾W˜Â0¥ÊB¹zÏ›Õõ•é›|çúŒMÉ»ûŒœ¥yðˆN‹4‚v,ÒKÞTWAÂ¡³0š Ê"›ÏJã]ÜOa–ÍÏøªf)­í=ËÅÞpG´]T¿ÔÿÒ‘6’;Ì=râ8º=¡&§Â§—Ãœ»'é(Û»?.Yaä‹dO[‘˜à²ÆÚ/¢ œO%ÍöÁ!˜’¡ÈË×„ó³áI¦G¾ô~ÓŠÿL£ÕEšêûîóZl¥)ÎÎ´®’ŒÑr)žå]Å_º·bMdþŒ€C¶fuH¬¤^_Y=4'4E°mø’ ;00®ïâ+e\`ÃÂg™Wí6„5{˜š&d˜Ô{ÓÒÞ‹Ñu™00ún%„Ýà‘ý&dÀ< jyÆ×Âjtùü1Z0M8³ƒ™x5íˆi`Ž¢‘=hjÇÔ\c"iIµ2•åÅ[€ê^öVÝ
¦Ì)á¨=æ9ÊÌ›Ñ±R0&k§fò#Èõ÷£p*{¯²_¥í?Ü·{4 ûýi{ÿýîáöÞÖ;¥ZMÀG[ï¶>½ß:>ÜÛþønKMÎ÷wÿã1·˜Ž
Üö€5_32aø®irbŒ‘£yÞFC@Òk×ÏU~¨îŒ¾êÑxŽI¾âïLSa˜4ÇÌ|·ÎÕ¡\BVMÃ× FB«kÄ¶N_Iß–±Jc:¸¬4 2~…1)­f<­­¦M­
­÷Ì"+ûk=UL: ‚ƒžžÃE>°5fÂÃ·jÚ˜Ö(pØ˜UÊÞ@”?Fr÷¬NdFP©~“˜/ÄXÛ%3yâ´Ô«<:‹Œh£47ãÅEÊ‘éÀò÷Êä‘2r3Y*jMS«ãsò <skÓ"ô5RkƒÕ@éò;UŽ’z­éÓ+(ú;ñÈóñïYlï¿{·»}¼·ÿáh°½õn÷ÃÎÖáÞþ§­íýÝ£^)…žÁJc†6ÿš-C™÷:²ƒà\%z¦¡›%'ë(ž¬À…ø9Æj›Ay50„`Õ‰ÇÒ´œ‘šønšµ’`*9¶š ]¹e!©>³«bp—ïPì^x€ïÍÔ0ôcÀÖTÂœƒôæŒx}WEÜ©ZÌËqí`„›¦JÀ3¢ö˜…ìœ‡Ñ%«Kƒ½¿YoƒëòÝNp÷Œ* `vþÂ3U ¼-å'Ÿ¿¹IEë–p€Ï%Æ‘qEÃ[2¨rw|j‡®íÆðÇà¼qàRùsÔBíÿñ«oàÛ ‚øKÁXÁMçgŒN&©¼Ãoº¨îtÄº%rU®.’ôöÓ»ßQÃmš"2€Ëù¡§‡v¦^9Ã\·ò*N¿dcÆã¿½;úÛ È1¡óKGŸòøå™úÚ]­àŸÅñŸ|‚ðÅÓÐîKš$`[%}ÝE­g=á@¦,kw;utk·y•9/ŒÍjå³"ù©¤÷é›½„ùš½£ý#¾æÀ]–„ãN÷dñôvpf×Ÿ+Ík;sÈBšßý7¬#DìcâÀòQ˜‹§±b’t
kŽjn–í’~‚Š["!¶Á§FÇÍÂ·ä/à¶‚oþí‚*t˜]Y1ñËÊ
ømm»Ã½Úy?cùcQ¥øbÍÚÇJ—åJéÉ²]ê’ÒÑ ~\ŒsQ3,xé‹³¡_¯Cfsãå–…/¢•J˜ÀAËÅHàËä ¸4± Æßö”/8åz© ý³ùIÌê »X R£âÔ§‹“Y¤â]«Ê¬Ë`–ÒÐç¥ºèB=|ÓX¯qÝ_©mdc˜W\˜v±†’qÆÂó>ß½Eµáw¼JwÄ&§^eé{§PCiìªÉbÌM.·\Ã-Õ’7fEb-ÿ8Ø#Û4õ³“ä"ÏÿÁI’õ—°©¼]&á…qûªvZ˜ä©.+²‰„^bqYfVÊ[Ír»Ø…¨ZS–¹5e‚ú+ ^`&ÀÊxp>ÂxôÅ™¸œƒïgd7d"yf©¹[.Skd„¦žA›a±Æ²ú¾¼TVnRêöèûÊ<†q–¡ò€?vWmO“mÉu6dÊ¦zMÇÝŠ SóþDŠ@'gÐÆÊ}AøB*bƒÉ£¢Ï?‰6ˆn½®Ñ†ÃoJ…°âf¸{\3ß[îR–„À†×á‹p¯Ó§5ÿW†[™ùx™iYÀ…÷Ï2ÚÆŠ6¼uþ_¸voæw‰çúZ<ö¯F…Bþ$#/º÷¦Í¸óŽ|…1¯S t-Þòý£YÎÄ½5¨ª–sQ+9 ®U	tôÄÈ7)†%RÕô¸YqFA8n×<§	<ÁY^R•Š‚KæêŠØ pFÞÀ£gY áƒVH'ý¥…eÒçþ*—Ä”?pw‡¦öG­çNñZ%Eî`æÓD…K“kß†…¥ë‹l-8¶:Ô“Y¬•7’-ƒÁÀ¥tIÃØ+w7Übvˆ( –¸ ñì”Uöê±xÓ,pBÝ¥ðWýó¼û$ý„¨B‚å7nÊÔÔªó²¡Î×expŠl¨'‚¸Å8|iãFÐxY­	—mDXH9Z~Ói'×TWƒU~ø#d#w(¥˜­£6’¶*…ac}¦md~ý8i«µ!c“ Î`R‡¦Ü
q~­¡Áz§ý$æ…ísN0Ä	ÏçÑóMÅã«Ìà”kF<,[—äðàMwmA Ì¡Ä6tfo[bòZnwõxx uèIUÏÜ¤l5Çø+Þ(ž8éT“¸m©<zkãHn{©'mµOí‹iêIkj/Í  ý/j²¹ÁÞF“Íã	X¯þ­ÂŽí,ë Œ™&º±ŠŸä.¿[]xÙô6UC}kü­ÛoG¾«Òk`¬[õÙ›Tµö¶êµÓ3ð5£XOäÂ—î;f]ÏCø=|˜ý„Ó>}§ŽîMÍf=ß¼^$›–è«1ç’tk­,ð´ÿ²UÈU¦ùmxQ› ä‹ÆÅ„WÌ`jÐÞîý7gÏÜø¸#Õ2ñæðSQ„-âhíÄ¯=Y‹×6Kshæ-²…‡Gbë8‰ÊñƒïòJ8<n"szÄJfáuÒÈØ$xH9Q}&E·˜ÀC|0¤j^g[d):ÝyÒq´¸K†3Ç^+X­lÖrz²;)Õ—W¸'™<éŠ±“’}õde>fÔw¦„²¯o”‰íŸU§'y²´˜\ŸÞ?=]aYIëŒ
phÍ
¹ñJ;qkùø~ˆ*B»?¦ÚÂ¿?æ6¼c•û}QMÐÙPYœÇQáµ ÷ïƒYÏL¼·;?wBƒWÜqK.q÷ÅÃ‚Š+Å­ÌÎ†.2xËm©-uÉOqåu?¾k²Þá1:E&Aæü¼/°‘~·^™[÷žVlB‚éª´³ØŸšü‚i‹ÓŸùÃ±`ZúbX&ÌŠIMnj*…â«	ükçäcyÉŠ#ód«qLÈ…›‚Šg©ðàê[0çÿ	÷³à'À/¸DVFp¼÷øÌãÇU»ƒ„b‘Wš{Ë=ÒYìÌ åÓ©CH0{9ø¦S<µ»ŽuGØ™Ûþ77œIü/o?7cÙ>¬y>¬Æëµºƒõ—²þ‡¬•m½xÑ=»HtõY¢U—á^°>Õ>'ÂMÆZ+hUãa—­ªé€?m¹ý\u¥$´ÒÞŒ¢z&ðêè]½P¯X!¬ÿqšY'Nì«ö‹Äø,ÊÂòJ#V%LWýLÎ¸*ø.ýo÷ZËýzÓÒ´a!£“Õø¶µ8ñÂ–1¸„¶—´ÏÏ]*ÓËipff/÷€È]ÍDÁóQ8,I‘&¡a«äý›j‚Û½†_l£jtq‹Ø³8nh¼šÇƒ¿²ŸÍêâómîoÔMLÞ h”>`ûP•ªË Ô8ý2:žÇ¼rÌ<ØÒj˜ÔÀ)/,}0«Ú[Ík#+³û@D”/
Ò"Á(Èxðdi™;”ó 1av5¸nâ±­°š)Ì•ÞW}À¨šT‹Ç™–µ¿fXT
?ŽÛÎgE”ç6il= aÙ˜Q¡|;·EU¸üÖf/8Ò2"-šyÉÏR­(ß¶B3Úr4ë¬¿šU•ù§žêMÙÖÊ”_i=åU1ùçû<ôù“ƒÉGÙoÓ¡˜¥jÍ%NMèÒ`e¦ÕQê/ÍD©Û¨‘‡rÉEÙ¾ªËîw°4ÜK~ÔRœ&_p‹3*í-ÌmYšÃÑ=>¼D~òä‰Ù.^ÚgžÉ½·ò#s+æ×äæùmx‰“9ËÝéÎIšßý³Nˆ°¾îXÎ¸±â“Ù9/ÀËMIëa¡Q0A'žþËÆþ~éÏõ­…=‚P2½û#öcž‰Q].ŠÓfÓ˜7 k[!Kóí …LÕ®¬Ô|ôÎÁš¹
à5cäfÊ\»XžÃy¨_!²d¥™lãaj­:òsê°{ôÍ<=ô¤½+káÌä ˜¨J›R(4}9Øg£žHmGW¾ú:«)x´{üˆ…½õW1Ÿ„ì’½¥¾ð=~ZY¦íÑ«þyÏñ%næ­Ä'eV®gœî™‡ù‡Èx¦p&1ñÙ(/ËS²æ¥2ü›v9Ã<móêR—æ…gìÏòhÇ !‡i^:v¨Í7ý`oªY‰'³Ð >i^x"ºìaw®LÔ÷Ü*§ö[8.ºûßÍï?—sšÉ8¸¤÷Ú;óän|ÆLq]~	à+1]žºxË_‰7íÜ<ˆµÊ‡ž†5·@d°(ÙIk//b^{>Çíf€þìx”Ïï7!Œ®’‰Ápù—jYÍÅ<•_1ÿñÜÎÜ‚0¯úoc×]|{"›‘>(/}t‘Ÿ¡o¥sÏ-Ù¶B©ùÅüþŠËÜb™³•8Æ6L[´,oÙÆÈ´‘½dXÖá¯ß”mƒgD½JSç£µÛ›ê*ÐJ ýââÙ¿…áXí–ã³y”¡á^Ì—Qý7Ó«WížhyÝÎŒ—Ÿ6¾OŽdVN‰”‰ñ¹|LÞÍÙ(+—2º)½—†oí?*ç¡2ª$&øt£°¸ûÃ§=ó³%AQz"ÕåùÓGçË·Ö~Õ…y!»Ì_ll¬2^‚1/„o¡1‹òÞÔÖfÕh~UÛK2Ü¢d^Ã6«DN×l–Ïíðÿ   ÿÿ À¿œ