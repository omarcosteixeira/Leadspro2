const fs = require('fs');

let types = fs.readFileSync('src/types.ts', 'utf8');

// Fix 4: Add metaDia to dashboardWidgets type
types = types.replace(
  `    metaCursos?: boolean;`,
  `    metaCursos?: boolean;
    metaDia?: boolean;`
);

fs.writeFileSync('src/types.ts', types);
console.log("Fixed TypeScript error 3 (in types.ts)!");
