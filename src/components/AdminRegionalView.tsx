import React, { useState, useEffect, useMemo } from "react";
import {
  Users,
  Link2,
  UserPlus,
  Plus,
  Trash2,
  Edit2,
  ExternalLink,
  Search,
  Download,
  Upload,
  ShieldCheck,
  Building,
  Briefcase,
  Calendar,
  Phone,
  Mail,
  CheckCircle2,
  XCircle,
  KeyRound,
  Lock,
  Unlock,
  Copy,
  FileSpreadsheet,
} from "lucide-react";
import { db, COLLECTIONS, secondaryAuth } from "../firebase";
import {
  collection,
  onSnapshot,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  serverTimestamp,
  setDoc,
  writeBatch,
} from "firebase/firestore";
import { createUserWithEmailAndPassword, sendPasswordResetEmail } from "firebase/auth";
import { Funcionario, LinkUtil, UserProfile, UserRole } from "../types";
import * as XLSX from "xlsx";

interface Props {
  profile: UserProfile;
  users: UserProfile[];
  links: LinkUtil[];
  onToast: (msg: string, type?: "success" | "error") => void;
  uniqueUnidades?: string[];
}

export function AdminRegionalView({
  profile,
  users,
  links: propLinks,
  onToast,
  uniqueUnidades = [],
}: Props) {
  const isAdmin =
    profile.role === "Administrador" ||
    profile.role === "Admin Master" ||
    profile.email === "marcos.teixeira@estacio.br" ||
    profile.email === "canaldonutri@gmail.com";

  const [activeTab, setActiveTab] = useState<"funcionarios" | "links" | "usuarios">("funcionarios");

  // ==========================================
  // 1. FUNCIONÁRIOS SM STATE
  // ==========================================
  const [funcionarios, setFuncionarios] = useState<Funcionario[]>([]);
  const [loadingFunc, setLoadingFunc] = useState(true);
  const [searchFunc, setSearchFunc] = useState("");
  const [unitFilter, setUnitFilter] = useState("Todas");
  const [isAddingFunc, setIsAddingFunc] = useState(false);
  const [editingFunc, setEditingFunc] = useState<Funcionario | null>(null);

  // Form State for Funcionario SM
  const [funcForm, setFuncForm] = useState({
    nome: "",
    email: "",
    matricula: "",
    telefone: "",
    cargo: "Consultor SM",
    unidade: "",
    dataNascimento: "",
    status: "Ativo" as "Ativo" | "Inativo",
    observacao: "",
  });

  // ==========================================
  // 2. LINKS ÚTEIS STATE
  // ==========================================
  const [linksList, setLinksList] = useState<LinkUtil[]>(propLinks || []);
  const [isAddingLink, setIsAddingLink] = useState(false);
  const [editingLink, setEditingLink] = useState<LinkUtil | null>(null);
  const [linkForm, setLinkForm] = useState({
    nome: "",
    url: "",
    local: "Geral",
  });

  // ==========================================
  // 3. USUÁRIOS DO SISTEMA (Admin only)
  // ==========================================
  const [isAddingUser, setIsAddingUser] = useState(false);
  const [userForm, setUserForm] = useState({
    name: "",
    email: "",
    role: "Gerente Regional SM" as UserRole,
    unidade: "",
    phone: "",
    password: "",
  });
  const [creatingUser, setCreatingUser] = useState(false);

  // Subscribe to Funcionários & Links in real time
  useEffect(() => {
    const unsubFunc = onSnapshot(
      collection(db, COLLECTIONS.FUNCIONARIOS),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as Funcionario[];
        list.sort((a, b) => (a.nome || "").localeCompare(b.nome || ""));
        setFuncionarios(list);
        setLoadingFunc(false);
      },
      (err) => {
        console.error("Erro ao carregar funcionários:", err);
        setLoadingFunc(false);
      }
    );

    const unsubLinks = onSnapshot(
      collection(db, COLLECTIONS.LINKS),
      (snap) => {
        const list = snap.docs.map((d) => ({
          id: d.id,
          ...d.data(),
        })) as LinkUtil[];
        setLinksList(list);
      },
      (err) => {
        console.error("Erro ao carregar links:", err);
      }
    );

    return () => {
      unsubFunc();
      unsubLinks();
    };
  }, []);

  // Filtered Funcionários
  const filteredFuncionarios = useMemo(() => {
    return funcionarios.filter((f) => {
      const matchUnit = unitFilter === "Todas" || f.unidade === unitFilter;
      if (!matchUnit) return false;
      if (searchFunc.trim()) {
        const term = searchFunc.toLowerCase();
        return (
          f.nome?.toLowerCase().includes(term) ||
          f.email?.toLowerCase().includes(term) ||
          f.matricula?.toLowerCase().includes(term) ||
          f.cargo?.toLowerCase().includes(term) ||
          f.telefone?.includes(term)
        );
      }
      return true;
    });
  }, [funcionarios, unitFilter, searchFunc]);

  // Handle Save / Edit Funcionario SM
  const handleSaveFuncionario = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!funcForm.nome.trim()) {
      onToast("O nome do funcionário é obrigatório.", "error");
      return;
    }

    try {
      const payload: any = {
        nome: funcForm.nome.trim(),
        email: funcForm.email.trim(),
        matricula: funcForm.matricula.trim() || `SM-${Date.now().toString().slice(-4)}`,
        telefone: funcForm.telefone.trim(),
        cargo: funcForm.cargo.trim(),
        unidade: funcForm.unidade.trim(),
        dataNascimento: funcForm.dataNascimento.trim(),
        status: funcForm.status,
        observacao: funcForm.observacao.trim(),
        tipo: "sm",
        updatedAt: serverTimestamp(),
      };

      if (editingFunc) {
        await updateDoc(doc(db, COLLECTIONS.FUNCIONARIOS, editingFunc.id), payload);
        onToast("Funcionário SM atualizado com sucesso!", "success");
      } else {
        payload.createdAt = serverTimestamp();
        await addDoc(collection(db, COLLECTIONS.FUNCIONARIOS), payload);
        onToast("Funcionário SM cadastrado com sucesso!", "success");
      }

      setIsAddingFunc(false);
      setEditingFunc(null);
      setFuncForm({
        nome: "",
        email: "",
        matricula: "",
        telefone: "",
        cargo: "Consultor SM",
        unidade: "",
        dataNascimento: "",
        status: "Ativo",
        observacao: "",
      });
    } catch (err: any) {
      console.error(err);
      onToast(`Erro ao salvar funcionário: ${err.message}`, "error");
    }
  };

  const handleEditFuncionario = (f: Funcionario) => {
    setEditingFunc(f);
    setFuncForm({
      nome: f.nome || "",
      email: f.email || "",
      matricula: f.matricula || "",
      telefone: f.telefone || "",
      cargo: f.cargo || "Consultor SM",
      unidade: f.unidade || "",
      dataNascimento: f.dataNascimento || "",
      status: (f.status as any) || "Ativo",
      observacao: f.observacao || "",
    });
    setIsAddingFunc(true);
  };

  const handleDeleteFuncionario = async (id: string, nome: string) => {
    if (!window.confirm(`Deseja realmente excluir o funcionário SM "${nome}"?`)) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.FUNCIONARIOS, id));
      onToast("Funcionário removido com sucesso.", "success");
    } catch (err: any) {
      onToast(`Erro ao excluir: ${err.message}`, "error");
    }
  };

  // Export to Excel
  const handleExportExcel = () => {
    const dataToExport = funcionarios.map((f) => ({
      Nome: f.nome,
      Matricula: f.matricula,
      Cargo: f.cargo || "SM",
      Unidade: f.unidade || "",
      Telefone: f.telefone || "",
      Email: f.email || "",
      "Data Nascimento": f.dataNascimento || "",
      Status: f.status || "Ativo",
      Observação: f.observacao || "",
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Funcionarios_SM");
    XLSX.writeFile(wb, `Funcionarios_Regional_SM_${new Date().toISOString().split("T")[0]}.xlsx`);
    onToast("Relatório exportado com sucesso!", "success");
  };

  // Import from Excel
  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const rawData: any[] = XLSX.utils.sheet_to_json(ws);

        if (!rawData || rawData.length === 0) {
          onToast("A planilha importada está vazia.", "error");
          return;
        }

        const batch = writeBatch(db);
        let count = 0;

        for (const row of rawData) {
          const nome = row.Nome || row.nome || row["Nome Completo"];
          if (!nome) continue;

          const docRef = doc(collection(db, COLLECTIONS.FUNCIONARIOS));
          batch.set(docRef, {
            nome: String(nome).trim(),
            matricula: String(row.Matricula || row.matricula || `SM-${Date.now().toString().slice(-4)}`).trim(),
            cargo: String(row.Cargo || row.cargo || "Consultor SM").trim(),
            unidade: String(row.Unidade || row.unidade || "").trim(),
            telefone: String(row.Telefone || row.telefone || "").trim(),
            email: String(row.Email || row.email || "").trim(),
            dataNascimento: String(row["Data Nascimento"] || row.dataNascimento || row.Aniversario || "").trim(),
            status: row.Status || "Ativo",
            observacao: row["Observação"] || row.observacao || "",
            tipo: "sm",
            createdAt: serverTimestamp(),
          });
          count++;
        }

        await batch.commit();
        onToast(`${count} funcionários importados com sucesso!`, "success");
      } catch (err: any) {
        console.error(err);
        onToast(`Erro ao importar planilha: ${err.message}`, "error");
      }
    };
    reader.readAsBinaryString(file);
  };

  // ==========================================
  // LINKS ÚTEIS HANDLERS
  // ==========================================
  const handleSaveLink = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!linkForm.nome.trim() || !linkForm.url.trim()) {
      onToast("Nome e URL são obrigatórios.", "error");
      return;
    }

    try {
      const payload = {
        nome: linkForm.nome.trim(),
        url: linkForm.url.trim(),
        local: linkForm.local.trim() || "Geral",
        updatedAt: serverTimestamp(),
      };

      if (editingLink) {
        await updateDoc(doc(db, COLLECTIONS.LINKS, editingLink.id), payload);
        onToast("Link atualizado com sucesso!", "success");
      } else {
        await addDoc(collection(db, COLLECTIONS.LINKS), {
          ...payload,
          createdAt: serverTimestamp(),
        });
        onToast("Link adicionado com sucesso!", "success");
      }

      setIsAddingLink(false);
      setEditingLink(null);
      setLinkForm({ nome: "", url: "", local: "Geral" });
    } catch (err: any) {
      onToast(`Erro ao salvar link: ${err.message}`, "error");
    }
  };

  const handleDeleteLink = async (id: string, nome: string) => {
    if (!window.confirm(`Excluir o link "${nome}"?`)) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.LINKS, id));
      onToast("Link removido com sucesso.", "success");
    } catch (err: any) {
      onToast(`Erro ao excluir link: ${err.message}`, "error");
    }
  };

  // ==========================================
  // USUÁRIOS DO SISTEMA (Admin only)
  // ==========================================
  const handleCreateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!userForm.name || !userForm.email) {
      onToast("Nome e Email são obrigatórios.", "error");
      return;
    }

    setCreatingUser(true);
    try {
      const password = userForm.password || "123456";
      const userCred = await createUserWithEmailAndPassword(
        secondaryAuth,
        userForm.email.trim(),
        password
      );

      const uid = userCred.user.uid;
      const newUserDoc: UserProfile = {
        uid,
        name: userForm.name.trim(),
        email: userForm.email.trim(),
        role: userForm.role,
        unidade: userForm.unidade.trim(),
        phone: userForm.phone.trim(),
        servidor: "regional",
        mustChangePassword: true,
        createdAt: serverTimestamp(),
      };

      await setDoc(doc(db, COLLECTIONS.USERS, uid), newUserDoc);
      onToast(
        `Usuário ${userForm.name} criado com sucesso! Senha inicial: ${password}`,
        "success"
      );

      setIsAddingUser(false);
      setUserForm({
        name: "",
        email: "",
        role: "Gerente Regional SM",
        unidade: "",
        phone: "",
        password: "",
      });
    } catch (err: any) {
      console.error(err);
      onToast(`Erro ao criar usuário: ${err.message}`, "error");
    } finally {
      setCreatingUser(false);
    }
  };

  const handleUpdateUserRole = async (userId: string, newRole: UserRole) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.USERS, userId), {
        role: newRole,
        updatedAt: serverTimestamp(),
      });
      onToast("Função do usuário atualizada com sucesso.", "success");
    } catch (err: any) {
      onToast(`Erro ao atualizar: ${err.message}`, "error");
    }
  };

  const handleToggleUserBlock = async (userId: string, currentBlocked?: boolean) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.USERS, userId), {
        blocked: !currentBlocked,
        updatedAt: serverTimestamp(),
      });
      onToast(currentBlocked ? "Usuário desbloqueado." : "Usuário bloqueado.", "success");
    } catch (err: any) {
      onToast(`Erro ao alterar status: ${err.message}`, "error");
    }
  };

  const handleSendResetPassword = async (email: string) => {
    try {
      await sendPasswordResetEmail(secondaryAuth, email);
      onToast(`Link de redefinição de senha enviado para ${email}.`, "success");
    } catch (err: any) {
      onToast(`Erro ao enviar e-mail: ${err.message}`, "error");
    }
  };

  const handleDeleteUser = async (userId: string, userName: string) => {
    if (!window.confirm(`Deseja realmente remover o usuário "${userName}" do sistema?`)) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.USERS, userId));
      onToast("Usuário removido da base de dados.", "success");
    } catch (err: any) {
      onToast(`Erro ao remover: ${err.message}`, "error");
    }
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto pb-16">
      {/* Header */}
      <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="px-3 py-1 bg-blue-50 text-blue-700 text-xs font-extrabold rounded-full border border-blue-200">
              Administração Regional
            </span>
          </div>
          <h1 className="text-2xl font-black text-slate-900">Gestão Administrativa</h1>
          <p className="text-xs text-slate-500 mt-1">
            Gerencie o time de Funcionários SM, cadastre Links Úteis e configure acessos de usuários.
          </p>
        </div>

        {/* Sub-Tabs Nav */}
        <div className="flex flex-wrap items-center gap-2 bg-slate-100 p-1.5 rounded-2xl border border-slate-200">
          <button
            onClick={() => setActiveTab("funcionarios")}
            className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
              activeTab === "funcionarios"
                ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <Users size={16} />
            <span>Funcionários SM</span>
            <span className="px-2 py-0.5 bg-blue-50 text-blue-700 text-[10px] font-extrabold rounded-full">
              {funcionarios.length}
            </span>
          </button>

          <button
            onClick={() => setActiveTab("links")}
            className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
              activeTab === "links"
                ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <Link2 size={16} />
            <span>Links Úteis</span>
            <span className="px-2 py-0.5 bg-slate-200 text-slate-700 text-[10px] font-extrabold rounded-full">
              {linksList.length}
            </span>
          </button>

          {/* Sub-aba Usuários: only for Administrador */}
          {isAdmin && (
            <button
              onClick={() => setActiveTab("usuarios")}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer ${
                activeTab === "usuarios"
                  ? "bg-white text-blue-600 shadow-sm border border-slate-200/60"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              <ShieldCheck size={16} />
              <span>Usuários do Sistema</span>
              <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-[10px] font-extrabold rounded-full">
                Admin
              </span>
            </button>
          )}
        </div>
      </div>

      {/* ========================================================================= */}
      {/* 1. SUB-ABA: FUNCIONÁRIOS SM */}
      {/* ========================================================================= */}
      {activeTab === "funcionarios" && (
        <div className="space-y-6">
          {/* Action Bar */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex flex-col sm:flex-row items-center gap-3 flex-1">
              <div className="relative w-full sm:w-80">
                <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  value={searchFunc}
                  onChange={(e) => setSearchFunc(e.target.value)}
                  placeholder="Buscar por nome, cargo, matrícula..."
                  className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>

              {uniqueUnidades.length > 0 && (
                <select
                  value={unitFilter}
                  onChange={(e) => setUnitFilter(e.target.value)}
                  className="w-full sm:w-auto px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                >
                  <option value="Todas">Todas as Unidades</option>
                  {uniqueUnidades.map((u) => (
                    <option key={u} value={u}>
                      {u}
                    </option>
                  ))}
                </select>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <label className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer">
                <Upload size={15} />
                <span>Importar Planilha</span>
                <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportExcel} className="hidden" />
              </label>

              <button
                onClick={handleExportExcel}
                className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition flex items-center gap-2 cursor-pointer"
              >
                <Download size={15} />
                <span>Exportar Excel</span>
              </button>

              <button
                onClick={() => {
                  setEditingFunc(null);
                  setFuncForm({
                    nome: "",
                    email: "",
                    matricula: "",
                    telefone: "",
                    cargo: "Consultor SM",
                    unidade: "",
                    dataNascimento: "",
                    status: "Ativo",
                    observacao: "",
                  });
                  setIsAddingFunc(true);
                }}
                className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition flex items-center gap-2 cursor-pointer"
              >
                <Plus size={16} />
                <span>Novo Funcionário SM</span>
              </button>
            </div>
          </div>

          {/* Modal / Form Add & Edit Funcionario */}
          {isAddingFunc && (
            <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
              <div className="bg-white rounded-3xl p-6 md:p-8 max-w-2xl w-full shadow-2xl border border-slate-100 space-y-6 max-h-[90vh] overflow-y-auto">
                <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                      <Users size={20} />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-900">
                        {editingFunc ? "Editar Funcionário SM" : "Novo Funcionário SM"}
                      </h3>
                      <p className="text-xs text-slate-500">Preencha os dados do colaborador regional</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setIsAddingFunc(false)}
                    className="p-2 text-slate-400 hover:text-slate-600 rounded-xl cursor-pointer"
                  >
                    ✕
                  </button>
                </div>

                <form onSubmit={handleSaveFuncionario} className="space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="sm:col-span-2">
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Nome Completo *
                      </label>
                      <input
                        type="text"
                        required
                        value={funcForm.nome}
                        onChange={(e) => setFuncForm({ ...funcForm, nome: e.target.value })}
                        placeholder="Ex: João da Silva"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Cargo / Função
                      </label>
                      <input
                        type="text"
                        value={funcForm.cargo}
                        onChange={(e) => setFuncForm({ ...funcForm, cargo: e.target.value })}
                        placeholder="Ex: Consultor SM, Gerente SM"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Matrícula
                      </label>
                      <input
                        type="text"
                        value={funcForm.matricula}
                        onChange={(e) => setFuncForm({ ...funcForm, matricula: e.target.value })}
                        placeholder="Ex: SM-1029"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Data de Nascimento (Aniversário)
                      </label>
                      <input
                        type="date"
                        value={funcForm.dataNascimento}
                        onChange={(e) => setFuncForm({ ...funcForm, dataNascimento: e.target.value })}
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        WhatsApp / Telefone
                      </label>
                      <input
                        type="text"
                        value={funcForm.telefone}
                        onChange={(e) => setFuncForm({ ...funcForm, telefone: e.target.value })}
                        placeholder="(21) 99999-9999"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        E-mail
                      </label>
                      <input
                        type="email"
                        value={funcForm.email}
                        onChange={(e) => setFuncForm({ ...funcForm, email: e.target.value })}
                        placeholder="colaborador@estacio.br"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Unidade / Polo
                      </label>
                      <input
                        type="text"
                        value={funcForm.unidade}
                        onChange={(e) => setFuncForm({ ...funcForm, unidade: e.target.value })}
                        placeholder="Ex: Polo Centro, Norte Shopping"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Status
                      </label>
                      <select
                        value={funcForm.status}
                        onChange={(e) => setFuncForm({ ...funcForm, status: e.target.value as any })}
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                      >
                        <option value="Ativo">Ativo</option>
                        <option value="Inativo">Inativo</option>
                      </select>
                    </div>

                    <div className="sm:col-span-2">
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Observações
                      </label>
                      <textarea
                        rows={2}
                        value={funcForm.observacao}
                        onChange={(e) => setFuncForm({ ...funcForm, observacao: e.target.value })}
                        placeholder="Anotações sobre horários, escala ou particularidades..."
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none resize-none"
                      />
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setIsAddingFunc(false)}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                    >
                      Cancelar
                    </button>
                    <button
                      type="submit"
                      className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition cursor-pointer"
                    >
                      {editingFunc ? "Salvar Alterações" : "Cadastrar Funcionário"}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* Table / List */}
          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50/80 text-slate-500 text-[11px] font-extrabold uppercase tracking-wider border-b border-slate-100">
                    <th className="px-6 py-4">Nome / Colaborador</th>
                    <th className="px-6 py-4">Cargo</th>
                    <th className="px-6 py-4">Unidade</th>
                    <th className="px-6 py-4">Aniversário</th>
                    <th className="px-6 py-4">Contato</th>
                    <th className="px-6 py-4">Status</th>
                    <th className="px-6 py-4 text-right">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs">
                  {filteredFuncionarios.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="px-6 py-12 text-center text-slate-400">
                        <Users size={32} className="mx-auto mb-2 text-slate-300" />
                        <p className="font-bold text-slate-600">Nenhum funcionário SM encontrado.</p>
                        <p className="text-[11px] text-slate-400 mt-0.5">Clique em "Novo Funcionário SM" ou importe uma planilha.</p>
                      </td>
                    </tr>
                  ) : (
                    filteredFuncionarios.map((f) => (
                      <tr key={f.id} className="hover:bg-slate-50/60 transition-colors">
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-xs shrink-0">
                              {f.nome.charAt(0)}
                            </div>
                            <div>
                              <p className="font-bold text-slate-900">{f.nome}</p>
                              <p className="text-[11px] text-slate-400">{f.matricula || "Sem matrícula"}</p>
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-4 font-medium text-slate-700">
                          {f.cargo || "Consultor SM"}
                        </td>
                        <td className="px-6 py-4 text-slate-600">
                          {f.unidade || "Geral"}
                        </td>
                        <td className="px-6 py-4">
                          {f.dataNascimento ? (
                            <span className="px-2.5 py-1 bg-amber-50 text-amber-800 text-[11px] font-bold rounded-lg border border-amber-200">
                              {f.dataNascimento.includes("-")
                                ? `${f.dataNascimento.split("-")[2]}/${f.dataNascimento.split("-")[1]}`
                                : f.dataNascimento}
                            </span>
                          ) : (
                            <span className="text-slate-300">-</span>
                          )}
                        </td>
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-2">
                            {f.telefone ? (
                              <a
                                href={`https://wa.me/55${f.telefone.replace(/\D/g, "")}`}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-bold rounded-lg flex items-center gap-1 transition"
                              >
                                <Phone size={11} />
                                <span>{f.telefone}</span>
                              </a>
                            ) : (
                              <span className="text-slate-300">-</span>
                            )}
                          </div>
                        </td>
                        <td className="px-6 py-4">
                          <span
                            className={`px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase ${
                              f.status === "Inativo"
                                ? "bg-rose-50 text-rose-600 border border-rose-200"
                                : "bg-emerald-50 text-emerald-700 border border-emerald-200"
                            }`}
                          >
                            {f.status || "Ativo"}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-1.5">
                            <button
                              onClick={() => handleEditFuncionario(f)}
                              className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition cursor-pointer"
                              title="Editar"
                            >
                              <Edit2 size={15} />
                            </button>
                            <button
                              onClick={() => handleDeleteFuncionario(f.id, f.nome)}
                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition cursor-pointer"
                              title="Excluir"
                            >
                              <Trash2 size={15} />
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

      {/* ========================================================================= */}
      {/* 2. SUB-ABA: LINKS ÚTEIS */}
      {/* ========================================================================= */}
      {activeTab === "links" && (
        <div className="space-y-6">
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <h2 className="text-lg font-bold text-slate-900">Gerenciar Links Úteis</h2>
              <p className="text-xs text-slate-500">Cadastre atalhos para portais, manuais e ferramentas</p>
            </div>

            <button
              onClick={() => {
                setEditingLink(null);
                setLinkForm({ nome: "", url: "", local: "Geral" });
                setIsAddingLink(true);
              }}
              className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition flex items-center gap-2 cursor-pointer"
            >
              <Plus size={16} />
              <span>Adicionar Novo Link</span>
            </button>
          </div>

          {/* Modal Add/Edit Link */}
          {isAddingLink && (
            <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
              <div className="bg-white rounded-3xl p-6 md:p-8 max-w-lg w-full shadow-2xl border border-slate-100 space-y-6">
                <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center">
                      <Link2 size={20} />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-900">
                        {editingLink ? "Editar Link Útil" : "Novo Link Útil"}
                      </h3>
                      <p className="text-xs text-slate-500">Informe o título e endereço do link</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setIsAddingLink(false)}
                    className="p-2 text-slate-400 hover:text-slate-600 rounded-xl cursor-pointer"
                  >
                    ✕
                  </button>
                </div>

                <form onSubmit={handleSaveLink} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                      Título do Link *
                    </label>
                    <input
                      type="text"
                      required
                      value={linkForm.nome}
                      onChange={(e) => setLinkForm({ ...linkForm, nome: e.target.value })}
                      placeholder="Ex: Portal do Aluno SIA"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                      URL / Link *
                    </label>
                    <input
                      type="text"
                      required
                      value={linkForm.url}
                      onChange={(e) => setLinkForm({ ...linkForm, url: e.target.value })}
                      placeholder="https://exemplo.com.br"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                      Categoria / Local
                    </label>
                    <input
                      type="text"
                      value={linkForm.local}
                      onChange={(e) => setLinkForm({ ...linkForm, local: e.target.value })}
                      placeholder="Ex: Portais, Acadêmico, Geral, Suporte"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setIsAddingLink(false)}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                    >
                      Cancelar
                    </button>
                    <button
                      type="submit"
                      className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition cursor-pointer"
                    >
                      {editingLink ? "Atualizar Link" : "Salvar Link"}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* Links Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {linksList.map((l) => (
              <div
                key={l.id}
                className="p-5 bg-white rounded-3xl border border-slate-100 shadow-sm flex flex-col justify-between gap-4"
              >
                <div>
                  <div className="flex items-start justify-between gap-2">
                    <span className="px-2.5 py-0.5 bg-slate-100 text-slate-600 text-[10px] font-extrabold rounded-md uppercase">
                      {l.local || "Geral"}
                    </span>
                    <div className="flex items-center gap-1">
                      <button
                        onClick={() => {
                          setEditingLink(l);
                          setLinkForm({
                            nome: l.nome,
                            url: l.url,
                            local: l.local || "Geral",
                          });
                          setIsAddingLink(true);
                        }}
                        className="p-1 text-slate-400 hover:text-blue-600 rounded cursor-pointer"
                      >
                        <Edit2 size={14} />
                      </button>
                      <button
                        onClick={() => handleDeleteLink(l.id, l.nome)}
                        className="p-1 text-slate-400 hover:text-rose-600 rounded cursor-pointer"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </div>

                  <h3 className="font-bold text-slate-900 text-sm mt-2">{l.nome}</h3>
                  <p className="text-xs text-slate-400 truncate mt-0.5">{l.url}</p>
                </div>

                <a
                  href={l.url.startsWith("http") ? l.url : `https://${l.url}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="w-full py-2 bg-slate-50 hover:bg-blue-50 text-blue-600 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 border border-slate-200/60 cursor-pointer"
                >
                  <span>Testar e Abrir Link</span>
                  <ExternalLink size={13} />
                </a>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 3. SUB-ABA: USUÁRIOS DO SISTEMA (Apenas Administrador) */}
      {/* ========================================================================= */}
      {activeTab === "usuarios" && isAdmin && (
        <div className="space-y-6">
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 bg-emerald-100 text-emerald-800 text-[10px] font-black uppercase rounded-md">
                  Controle de Acessos
                </span>
              </div>
              <h2 className="text-lg font-bold text-slate-900 mt-1">Usuários do Servidor Regional</h2>
              <p className="text-xs text-slate-500">
                Crie e configure contas de acesso para Administradores e Gerentes Regionais SM
              </p>
            </div>

            <button
              onClick={() => {
                setUserForm({
                  name: "",
                  email: "",
                  role: "Gerente Regional SM",
                  unidade: "",
                  phone: "",
                  password: "",
                });
                setIsAddingUser(true);
              }}
              className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition flex items-center gap-2 cursor-pointer"
            >
              <UserPlus size={16} />
              <span>Novo Usuário Regional</span>
            </button>
          </div>

          {/* Modal Create User */}
          {isAddingUser && (
            <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
              <div className="bg-white rounded-3xl p-6 md:p-8 max-w-lg w-full shadow-2xl border border-slate-100 space-y-6">
                <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                      <UserPlus size={20} />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-900">Novo Usuário do Sistema</h3>
                      <p className="text-xs text-slate-500">Cadastre o login e defina o perfil de permissão</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setIsAddingUser(false)}
                    className="p-2 text-slate-400 hover:text-slate-600 rounded-xl cursor-pointer"
                  >
                    ✕
                  </button>
                </div>

                <form onSubmit={handleCreateUser} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                      Nome Completo *
                    </label>
                    <input
                      type="text"
                      required
                      value={userForm.name}
                      onChange={(e) => setUserForm({ ...userForm, name: e.target.value })}
                      placeholder="Ex: Maria Pereira"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                      E-mail de Login *
                    </label>
                    <input
                      type="email"
                      required
                      value={userForm.email}
                      onChange={(e) => setUserForm({ ...userForm, email: e.target.value })}
                      placeholder="usuario@estacio.br"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Perfil de Acesso *
                      </label>
                      <select
                        value={userForm.role}
                        onChange={(e) => setUserForm({ ...userForm, role: e.target.value as UserRole })}
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none"
                      >
                        <option value="Administrador">Administrador (Total)</option>
                        <option value="Gerente Regional SM">Gerente Regional SM</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Senha Provisória
                      </label>
                      <input
                        type="text"
                        value={userForm.password}
                        onChange={(e) => setUserForm({ ...userForm, password: e.target.value })}
                        placeholder="Padrão: 123456"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Unidade / Polo
                      </label>
                      <input
                        type="text"
                        value={userForm.unidade}
                        onChange={(e) => setUserForm({ ...userForm, unidade: e.target.value })}
                        placeholder="Ex: Polo Regional Norte"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Telefone / WhatsApp
                      </label>
                      <input
                        type="text"
                        value={userForm.phone}
                        onChange={(e) => setUserForm({ ...userForm, phone: e.target.value })}
                        placeholder="(21) 99999-9999"
                        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setIsAddingUser(false)}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer"
                    >
                      Cancelar
                    </button>
                    <button
                      type="submit"
                      disabled={creatingUser}
                      className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl shadow-md shadow-blue-600/20 transition cursor-pointer disabled:opacity-50"
                    >
                      {creatingUser ? "Criando Usuário..." : "Criar Usuário"}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* Users Table */}
          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50/80 text-slate-500 text-[11px] font-extrabold uppercase tracking-wider border-b border-slate-100">
                    <th className="px-6 py-4">Usuário</th>
                    <th className="px-6 py-4">Email</th>
                    <th className="px-6 py-4">Unidade</th>
                    <th className="px-6 py-4">Perfil</th>
                    <th className="px-6 py-4">Status</th>
                    <th className="px-6 py-4 text-right">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs">
                  {users.map((u) => (
                    <tr key={u.uid} className="hover:bg-slate-50/60 transition-colors">
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-xs shrink-0">
                            {u.name?.charAt(0) || "U"}
                          </div>
                          <div>
                            <p className="font-bold text-slate-900">{u.name}</p>
                            {u.phone && <p className="text-[11px] text-slate-400">{u.phone}</p>}
                          </div>
                        </div>
                      </td>
                      <td className="px-6 py-4 text-slate-600 font-mono text-[11px]">{u.email}</td>
                      <td className="px-6 py-4 text-slate-600">{u.unidade || "Todas"}</td>
                      <td className="px-6 py-4">
                        <select
                          value={u.role}
                          onChange={(e) => handleUpdateUserRole(u.uid, e.target.value as UserRole)}
                          className="px-2.5 py-1 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-700 outline-none"
                        >
                          <option value="Administrador">Administrador</option>
                          <option value="Gerente Regional SM">Gerente Regional SM</option>
                          <option value="Admin Master">Admin Master</option>
                        </select>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase ${
                            u.blocked
                              ? "bg-rose-50 text-rose-600 border border-rose-200"
                              : "bg-emerald-50 text-emerald-700 border border-emerald-200"
                          }`}
                        >
                          {u.blocked ? "Bloqueado" : "Ativo"}
                        </span>
                      </td>
                      <td className="px-6 py-4 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          <button
                            onClick={() => handleSendResetPassword(u.email)}
                            className="p-1.5 text-slate-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg transition cursor-pointer"
                            title="Enviar Redefinição de Senha"
                          >
                            <KeyRound size={15} />
                          </button>
                          <button
                            onClick={() => handleToggleUserBlock(u.uid, u.blocked)}
                            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-lg transition cursor-pointer"
                            title={u.blocked ? "Desbloquear" : "Bloquear"}
                          >
                            {u.blocked ? <Unlock size={15} /> : <Lock size={15} />}
                          </button>
                          <button
                            onClick={() => handleDeleteUser(u.uid, u.name)}
                            className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition cursor-pointer"
                            title="Excluir Usuário"
                          >
                            <Trash2 size={15} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default AdminRegionalView;
