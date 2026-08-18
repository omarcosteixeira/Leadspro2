const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

const importTarget = `  MetaDia,
  QgLigacao,
  Ligacao,`;
const importReplace = `  MetaDia,
  MetaSM,
  QgLigacao,
  Ligacao,`;
code = code.replace(importTarget, importReplace);

const propsTarget = `  pedidosCursos?: PedidoCursoEntry[];
  metaDia?: MetaDia[];
  ligacoes?: Ligacao[];`;
const propsReplace = `  pedidosCursos?: PedidoCursoEntry[];
  metaDia?: MetaDia[];
  metaSM?: MetaSM[];
  ligacoes?: Ligacao[];`;
code = code.replace(propsTarget, propsReplace);

const fnTarget = `  pedidosCursos = [],
  metaDia = [],
  ligacoes = [],`;
const fnReplace = `  pedidosCursos = [],
  metaDia = [],
  metaSM = [],
  ligacoes = [],`;
code = code.replace(fnTarget, fnReplace);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
