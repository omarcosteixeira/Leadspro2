import fs from 'fs';
let code = fs.readFileSync('firestore.rules', 'utf8');

const target1 = `    match /artifacts/gestaopro-761e1/public/data/email_campaign_logs/{id} {
      allow read, write: if true;
    }`;
const replacement1 = target1 + `
    match /artifacts/gestaopro-761e1/public/data/sales_contacts/{id} {
      allow read, write: if true;
    }`;

const target2 = `    match /artifacts/gestaodeleadspro-d4230/public/data/email_campaign_logs/{id} {
      allow read, write: if true;
    }`;
const replacement2 = target2 + `
    match /artifacts/gestaodeleadspro-d4230/public/data/sales_contacts/{id} {
      allow read, write: if true;
    }`;

code = code.replace(target1, replacement1);
code = code.replace(target2, replacement2);
fs.writeFileSync('firestore.rules', code);
