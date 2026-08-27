const fs = require('fs');
let app = fs.readFileSync('src/App.tsx', 'utf8');

// Give Regional access to Formulários inside admin
app = app.replace(
  `{ id: "formularios", label: "Formulários", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)", "Gestor Unidade"] },`,
  `{ id: "formularios", label: "Formulários", roles: ["Admin Master", "Líder/FDV", "Gestor Comercial", "Gerente Comercial (Comercial)", "Gestor Unidade", "Regional"] },`
);

fs.writeFileSync('src/App.tsx', app);
console.log("Patched formularios successfully!");
