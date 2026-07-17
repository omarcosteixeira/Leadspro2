const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

const anchor = `{[
              { title: "Personalizado", stats: metaDiaStats.custom },
              { title: "Geral (Todo o Período)", stats: metaDiaStats.allTime },
              { title: "Mensal (Últimos 30 Dias)", stats: metaDiaStats.monthly },
              { title: "Semanal (Últimos 7 Dias)", stats: metaDiaStats.weekly }
            ].map(section => (`;

const replacement = `{[
              ...(metaDiaDataInicio || metaDiaDataFim ? [{ title: "Personalizado", stats: metaDiaStats.custom }] : []),
              { title: "Geral (Todo o Período)", stats: metaDiaStats.allTime },
              { title: "Mensal (Últimos 30 Dias)", stats: metaDiaStats.monthly },
              { title: "Semanal (Últimos 7 Dias)", stats: metaDiaStats.weekly }
            ].map(section => (`;

if (code.includes(anchor)) {
    code = code.replace(anchor, replacement);
    fs.writeFileSync('src/components/RelatoriosView.tsx', code, 'utf8');
    console.log("Successfully patched RelatoriosView.tsx.");
} else {
    console.log("Anchor not found.");
}
