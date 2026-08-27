const fs = require('fs');
let rules = fs.readFileSync('firestore.rules', 'utf8');

const oldFiesRules = `    match /artifacts/gestaodeleadspro-d4230/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isFinanceiro() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;

const newFiesRules = `    match /artifacts/gestaodeleadspro-d4230/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Financeiro", "Sala de Matrícula", "Líder/FDV", "SSA", "Gestor Unidade", "Regional"]);
      allow create, update: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV", "Sala de Matrícula", "SSA", "Gestor Unidade", "Regional"]);
      allow delete: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV"]);
    }`;

rules = rules.replace(oldFiesRules, newFiesRules);

const oldFiesVagasRules = `    match /artifacts/gestaodeleadspro-d4230/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isFinanceiro() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;

const newFiesVagasRules = `    match /artifacts/gestaodeleadspro-d4230/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Financeiro", "Sala de Matrícula", "Líder/FDV", "SSA", "Gestor Unidade", "Regional"]);
      allow create, update: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV", "Sala de Matrícula", "SSA", "Gestor Unidade", "Regional"]);
      allow delete: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV"]);
    }`;

rules = rules.replace(oldFiesVagasRules, newFiesVagasRules);

// also for gestaopro-761e1
const oldFiesRules1 = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isFinanceiro() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;
const newFiesRules1 = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Financeiro", "Sala de Matrícula", "Líder/FDV", "SSA", "Gestor Unidade", "Regional"]);
      allow create, update: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV", "Sala de Matrícula", "SSA", "Gestor Unidade", "Regional"]);
      allow delete: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV"]);
    }`;
rules = rules.replace(oldFiesRules1, newFiesRules1);

const oldFiesVagasRules1 = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isFinanceiro() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;
const newFiesVagasRules1 = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Financeiro", "Sala de Matrícula", "Líder/FDV", "SSA", "Gestor Unidade", "Regional"]);
      allow create, update: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV", "Sala de Matrícula", "SSA", "Gestor Unidade", "Regional"]);
      allow delete: if isAuthenticated() && hasAnyRole(["Admin Master", "Gestor Comercial", "Gerente Comercial (Comercial)", "Líder/FDV"]);
    }`;
rules = rules.replace(oldFiesVagasRules1, newFiesVagasRules1);

fs.writeFileSync('firestore.rules', rules);
console.log("Patched firestore.rules for fies_prouni successfully!");
