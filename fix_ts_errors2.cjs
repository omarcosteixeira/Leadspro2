const fs = require('fs');

let app = fs.readFileSync('src/App.tsx', 'utf8');

// Fix 3: periodos: false -> periodo: false
app = app.replace(
  `      qgLigacoes: false,
      periodos: false,
      bomDia: false,`,
  `      qgLigacoes: false,
      periodo: false,
      bomDia: false,`
);

fs.writeFileSync('src/App.tsx', app);
console.log("Fixed TypeScript error 2!");
