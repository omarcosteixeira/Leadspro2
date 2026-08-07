import fs from 'fs';
let code = fs.readFileSync('src/types.ts', 'utf8');

code = code.replace(
  "status: 'Pendente' | 'Sem retorno' | 'Interessado' | 'Não Interessado' | 'Convertido';",
  "status: 'Pendente' | 'Sem retorno' | 'Interessado' | 'Não Interessado' | 'Convertido' | 'Contato via Sales';"
);

code = code.replace(
  "status: 'Pendente' | 'Interessado' | 'Convertido' | 'Não tem interesse' | 'Sem retorno';",
  "status: 'Pendente' | 'Interessado' | 'Convertido' | 'Não tem interesse' | 'Sem retorno' | 'Contato via Sales';"
);

fs.writeFileSync('src/types.ts', code);
