const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

// Add "metaSM" to activeTab state
const stateTarget = `    | "crescimentoAnual"
    | "formularios"
  >(profile?.role === "Gestor Unidade" ? "forecast" : "usuarios");`;
const stateReplacement = `    | "crescimentoAnual"
    | "formularios"
    | "metaSM"
  >(profile?.role === "Gestor Unidade" ? "forecast" : "usuarios");`;
code = code.replace(stateTarget, stateReplacement);

// Add button
const tabsTarget = `          { id: "metaDia", label: "Meta Dia", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },`;
const tabsReplacement = `          { id: "metaDia", label: "Meta Dia", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },
          { id: "metaSM", label: "Meta SM", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)"] },`;
code = code.replace(tabsTarget, tabsReplacement);

fs.writeFileSync('src/App.tsx', code);
