const fs = require('fs');

let app = fs.readFileSync('src/App.tsx', 'utf8');

// Fix 5: Add metaDia to defaultWidgets
app = app.replace(
  `    metaCursos: true,
    aniversarios: true,`,
  `    metaCursos: true,
    metaDia: true,
    aniversarios: true,`
);

fs.writeFileSync('src/App.tsx', app);
console.log("Fixed TypeScript error 4 (defaultWidgets)!");
