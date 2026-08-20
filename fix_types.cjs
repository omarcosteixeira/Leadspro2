const fs = require('fs');

let code = fs.readFileSync('src/types.ts', 'utf8');
code = code.replace(
  'aniversarios?: boolean;',
  'aniversarios?: boolean;\n    metaSM?: boolean;\n    metaCursos?: boolean;'
);
fs.writeFileSync('src/types.ts', code);

