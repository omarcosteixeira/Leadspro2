const fs = require('fs');
let code = fs.readFileSync('firestore.rules', 'utf8');

const target = `    match /artifacts/gestaodeleadspro-d4230/public/data/meta_dia/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }`;
    
const replacement = target + `
    match /artifacts/gestaodeleadspro-d4230/public/data/meta_sm/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }
    match /artifacts/gestaodeleadspro-d4230/public/data/meta_cursos/{id} {
      allow read: if isAuthenticated();
      allow write: if isPrincipal() || isComercial();
    }`;

code = code.replace(target, replacement);

fs.writeFileSync('firestore.rules', code);
