const fs = require('fs');

// 1. Update src/types.ts
let types = fs.readFileSync('src/types.ts', 'utf8');
types = types.replace(
  `'Gerente Comercial (Comercial)' | 'FDV (Comercial)' | 'Financeiro' | 'Técnico';`,
  `'Gerente Comercial (Comercial)' | 'FDV (Comercial)' | 'Financeiro' | 'Técnico' | 'Regional';`
);
fs.writeFileSync('src/types.ts', types);

// 2. Update src/App.tsx
let app = fs.readFileSync('src/App.tsx', 'utf8');

// Add REGIONAL to ROLES
app = app.replace(
  `TECNICO: "Técnico",\n};`,
  `TECNICO: "Técnico",\n  REGIONAL: "Regional",\n};`
);

// Add checklist to VIEW_PERMISSIONS
app = app.replace(
  `  admin: [`,
  `  checklist: [\n    ROLES.ADMIN_MASTER,\n    ROLES.REGIONAL\n  ],\n  admin: [`
);

// Give REGIONAL access to dashboard, admin, relatorios, formularios
const addRoleToView = (viewName) => {
  const regex = new RegExp(`(${viewName}: \\[[\\s\\S]*?)(  \\],)`);
  app = app.replace(regex, `$1    ROLES.REGIONAL,\n$2`);
};
addRoleToView('dashboard');
addRoleToView('relatorios');
addRoleToView('admin');
addRoleToView('formularios');

// Add CheckList to main nav
app = app.replace(
  `{ id: "admin", label: "Administração", icon: Settings },`,
  `{ id: "checklist", label: "CheckList", icon: CheckSquare },\n              { id: "admin", label: "Administração", icon: Settings },`
);

// Add CheckList main view
app = app.replace(
  `{currentView === "admin" && (`,
  `{currentView === "checklist" && (\n                <div className="p-8">\n                  <h2 className="text-2xl font-bold mb-4 text-slate-800">CheckList</h2>\n                  <p className="text-slate-500">Conteúdo em breve.</p>\n                </div>\n              )}\n              {currentView === "admin" && (`
);

// Add Funcionários SM to Admin subtabs
app = app.replace(
  `{ id: "funcionarios", label: "Funcionários (Insumos)", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },`,
  `{ id: "funcionarios", label: "Funcionários (Insumos)", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },\n          { id: "funcionariosSM", label: "Funcionários SM", roles: ["Admin Master", "Regional"] },`
);

// Give Regional access to Links Uteis inside admin
app = app.replace(
  `{ id: "links", label: "Links Úteis", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)", "Gestor Unidade"] },`,
  `{ id: "links", label: "Links Úteis", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)", "Gestor Unidade", "Regional"] },`
);

// Add the view for Funcionários SM inside AdminView
app = app.replace(
  `{activeTab === "funcionarios" && (`,
  `{activeTab === "funcionariosSM" && (\n        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">\n          <h3 className="text-xl font-bold text-slate-900 mb-4">Funcionários SM</h3>\n          <p className="text-slate-500">Conteúdo em breve.</p>\n        </section>\n      )}\n\n      {activeTab === "funcionarios" && (`
);

// Modify Dashboard widgets for Regional role
const widgetModifier = `const isRegional = profile?.role === ROLES.REGIONAL;
  
  let widgets = profile?.dashboardWidgets
    ? { ...defaultWidgets, ...profile.dashboardWidgets }
    : defaultWidgets;
    
  if (isRegional) {
    widgets = {
      ...defaultWidgets,
      stats: false,
      planner: false,
      qg: false,
      periodos: false,
      bomDia: false,
      forecast: false,
      metaDia: false,
      metaCursos: false,
      metaSM: false,
      aniversarios: true,
      links: true
    };
  }`;
app = app.replace(
  `const widgets = profile?.dashboardWidgets\n    ? { ...defaultWidgets, ...profile.dashboardWidgets }\n    : defaultWidgets;`,
  widgetModifier
);

fs.writeFileSync('src/App.tsx', app);
console.log("Patched successfully!");
