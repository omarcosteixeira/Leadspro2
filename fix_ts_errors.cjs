const fs = require('fs');

let app = fs.readFileSync('src/App.tsx', 'utf8');

// Fix 1: qg: false -> qgLigacoes: false (or just remove it if qg is wrong, let's change to qgLigacoes: false)
app = app.replace(
  `      planner: false,
      qg: false,
      periodos: false,`,
  `      planner: false,
      qgLigacoes: false,
      periodos: false,`
);

// Fix 2: add "funcionariosSM" to activeTab type union
app = app.replace(
  `    | "funcionarios"
    | "crescimentoAnual"
    | "formularios"
    | "metaSM"
    | "metaCursos"
  >`,
  `    | "funcionarios"
    | "funcionariosSM"
    | "crescimentoAnual"
    | "formularios"
    | "metaSM"
    | "metaCursos"
  >`
);

fs.writeFileSync('src/App.tsx', app);
console.log("Fixed TypeScript errors!");
