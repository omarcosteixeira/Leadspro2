import fs from 'fs';
let code = fs.readFileSync('src/firebase.ts', 'utf8');

const target = `export const firebaseConfigPrincipal = {
  ...firebaseConfigPrincipalRaw,
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY_PRINCIPAL || firebaseConfigPrincipalRaw.apiKey,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN_PRINCIPAL || firebaseConfigPrincipalRaw.authDomain,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID_PRINCIPAL || firebaseConfigPrincipalRaw.projectId,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET_PRINCIPAL || firebaseConfigPrincipalRaw.storageBucket,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID_PRINCIPAL || firebaseConfigPrincipalRaw.messagingSenderId,
  appId: import.meta.env.VITE_FIREBASE_APP_ID_PRINCIPAL || firebaseConfigPrincipalRaw.appId,
};`;

const replacement = `export const firebaseConfigPrincipal = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY_PRINCIPAL || "AIzaSyDSlDqp7Nn7UCjjF2jkcEwcaMyoRZV4yBo",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN_PRINCIPAL || "gestaopro-761e1.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID_PRINCIPAL || "gestaopro-761e1",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET_PRINCIPAL || "gestaopro-761e1.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID_PRINCIPAL || "371631944013",
  appId: import.meta.env.VITE_FIREBASE_APP_ID_PRINCIPAL || "1:371631944013:web:1883a10f7be4e5a5738704",
};`;

code = code.replace(target, replacement);
fs.writeFileSync('src/firebase.ts', code);
