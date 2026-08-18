const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

const tabTypeTarget = `    "historico" | "bases" | "fiesProuni" | "planoAcao" | "empresas" | "insumos" | "isencoes" | "pedidos_cursos" | "metaDia" | "ligacoes" | "crescimento" | "manutencao" | "sales"
  >("historico");`;
const tabTypeReplace = `    "historico" | "bases" | "fiesProuni" | "planoAcao" | "empresas" | "insumos" | "isencoes" | "pedidos_cursos" | "metaDia" | "ligacoes" | "crescimento" | "manutencao" | "sales" | "metaSM"
  >("historico");`;
code = code.replace(tabTypeTarget, tabTypeReplace);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
