const fs = require('fs');
let code = fs.readFileSync('firestore.rules', 'utf8');

const fpAnchor = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }
    match /artifacts/gestaopro-761e1/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;

const fpReplacement = `    match /artifacts/gestaopro-761e1/public/data/fies_prouni/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isSSA() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || isSSA() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }
    match /artifacts/gestaopro-761e1/public/data/fies_prouni_vagas/{entryId} {
      allow get, list: if isAuthenticated() && (isMasterUser() || isComercial() || isSSA() || canAccessUnit(resource.data.unidade));
      allow create, update: if (isMasterUser() || isComercial() || isLider() || isSalaMatricula() || isSSA()) && (isMasterUser() || isComercial() || isSSA() || canAccessUnit(request.resource.data.unidade));
      allow delete: if isMasterUser() || isComercial() || isLider();
    }`;

if(code.includes(fpAnchor)) {
  code = code.replace(fpAnchor, fpReplacement);
  fs.writeFileSync('firestore.rules', code, 'utf8');
  console.log("Patched firestore.rules");
} else {
  console.log("Anchor not found in firestore.rules");
}
