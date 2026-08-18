const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

// 1. Add MetaCurso to imports
const importTarget = `  MetaSM,`;
const importReplace = `  MetaSM,
  MetaCurso,`;
code = code.replace(importTarget, importReplace);

// 2. Add State
const stateTarget = `  const [metaSM, setMetaSM] = useState<MetaSM[]>([]);`;
const stateReplace = `  const [metaSM, setMetaSM] = useState<MetaSM[]>([]);
  const [metaCursos, setMetaCursos] = useState<MetaCurso[]>([]);`;
code = code.replace(stateTarget, stateReplace);

// 3. Add onSnapshot
const snapshotTarget = `    let unsubMetaSM = () => {};
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
    }`;
const snapshotReplace = `    let unsubMetaSM = () => {};
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
    }`;
code = code.replace(snapshotTarget, snapshotReplace);

// 4. Add cleanup
const cleanupTarget = `      unsubMetaSM();
      unsubQgLigacoes();`;
const cleanupReplace = `      unsubMetaSM();
      unsubMetaCursos();
      unsubQgLigacoes();`;
code = code.replace(cleanupTarget, cleanupReplace);

// 5. Update component props (DashboardView, RelatoriosView, AdminView)
// DashboardView props
const dbPropTarget = `                  metaSM={metaSM}
                  qgLigacoes={qgLigacoes}`;
const dbPropReplace = `                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}`;
code = code.replace(dbPropTarget, dbPropReplace);

// RelatoriosView props
const relPropTarget = `                  metaDia={metaDia}
                  metaSM={metaSM}
                  ligacoes={ligacoes}`;
const relPropReplace = `                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  ligacoes={ligacoes}`;
code = code.replace(relPropTarget, relPropReplace);

// AdminView props
const adminPropTarget = `                  metaDia={metaDia}
                  metaSM={metaSM}
                  qgLigacoes={qgLigacoes}`;
const adminPropReplace = `                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}`;
code = code.replace(adminPropTarget, adminPropReplace);

// AdminView definition
const adminDefTarget = `  metaDia,
  metaSM,
  qgLigacoes,`;
const adminDefReplace = `  metaDia,
  metaSM,
  metaCursos,
  qgLigacoes,`;
code = code.replace(adminDefTarget, adminDefReplace);

const adminTypeTarget = `  metaDia: MetaDia[];
  metaSM: MetaSM[];
  qgLigacoes: QgLigacao[];`;
const adminTypeReplace = `  metaDia: MetaDia[];
  metaSM: MetaSM[];
  metaCursos: MetaCurso[];
  qgLigacoes: QgLigacao[];`;
code = code.replace(adminTypeTarget, adminTypeReplace);

// Add to AdminView tabs
const adminTabTarget = `          { id: "metaSM", label: "Meta SM", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },`;
const adminTabReplace = `          { id: "metaSM", label: "Meta SM", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },
          { id: "metaCursos", label: "Meta Cursos", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },`;
code = code.replace(adminTabTarget, adminTabReplace);

const adminStateTarget = `    | "metaSM"
  >(profile?.role === "Gestor Unidade" ? "forecast" : "usuarios");`;
const adminStateReplace = `    | "metaSM"
    | "metaCursos"
  >(profile?.role === "Gestor Unidade" ? "forecast" : "usuarios");`;
code = code.replace(adminStateTarget, adminStateReplace);

// Import MetaCursosView and render it in AdminView
const adminImportTarget = `import MetaSMView from "./components/MetaSMView";`;
const adminImportReplace = `import MetaSMView from "./components/MetaSMView";
import MetaCursosView from "./components/MetaCursosView";`;
code = code.replace(adminImportTarget, adminImportReplace);

const adminRenderTarget = `      {activeTab === "metaSM" && (
        <MetaSMView metaSM={metaSM} onToast={onToast} />
      )}`;
const adminRenderReplace = `      {activeTab === "metaSM" && (
        <MetaSMView metaSM={metaSM} onToast={onToast} />
      )}
      {activeTab === "metaCursos" && (
        <MetaCursosView metaCursos={metaCursos} onToast={onToast} />
      )}`;
code = code.replace(adminRenderTarget, adminRenderReplace);

// DashboardView props inside App.tsx (definition)
const dbDefTarget = `  metaSM,
  qgLigacoes,`;
const dbDefReplace = `  metaSM,
  metaCursos,
  qgLigacoes,`;
code = code.replace(dbDefTarget, dbDefReplace);

const dbTypeTarget = `  metaSM: MetaSM[];
  qgLigacoes: QgLigacao[];`;
const dbTypeReplace = `  metaSM: MetaSM[];
  metaCursos: MetaCurso[];
  qgLigacoes: QgLigacao[];`;
code = code.replace(dbTypeTarget, dbTypeReplace);

// Add to defaultWidgets in DashboardView
const widgetTarget = `    metaSM: true,
    qgLigacoes: true,`;
const widgetReplace = `    metaSM: true,
    metaCursos: true,
    qgLigacoes: true,`;
code = code.replace(widgetTarget, widgetReplace);

fs.writeFileSync('src/App.tsx', code);
