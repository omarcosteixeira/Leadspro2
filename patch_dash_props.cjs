const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const target1 = `  metaDia,
  qgLigacoes,`;
const replace1 = `  metaDia,
  metaSM,
  qgLigacoes,`;
code = code.replace(target1, replace1);

const target2 = `  metaDia: MetaDia[];
  qgLigacoes: QgLigacao[];`;
const replace2 = `  metaDia: MetaDia[];
  metaSM: MetaSM[];
  qgLigacoes: QgLigacao[];`;
code = code.replace(target2, replace2);

fs.writeFileSync('src/App.tsx', code);
