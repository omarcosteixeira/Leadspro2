const fs = require('fs');
let code = fs.readFileSync('firestore.rules', 'utf8');

const target = `    match /artifacts/gestaopro-761e1/public/data/meta_dia/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }`;

const replace = target + `
    match /artifacts/gestaopro-761e1/public/data/meta_sm/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }
    match /artifacts/gestaopro-761e1/public/data/meta_cursos/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }`;

code = code.replace(target, replace);
fs.writeFileSync('firestore.rules', code);
