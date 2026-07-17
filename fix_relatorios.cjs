const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

code = code.replace(
  '  const [metaDiaDataInicio, setMetaDiaDataInicio] = useState("");\n  const [metaDiaDataFim, setMetaDiaDataFim] = useState("");\n',
  ''
);

const anchor = '  const [ligacoesSearchTerm, setLigacoesSearchTerm] = useState("");';
const replacement = '  const [ligacoesSearchTerm, setLigacoesSearchTerm] = useState("");\n  const [metaDiaDataInicio, setMetaDiaDataInicio] = useState("");\n  const [metaDiaDataFim, setMetaDiaDataFim] = useState("");';

if (code.includes(anchor)) {
    code = code.replace(anchor, replacement);
} else {
    // maybe it was replaced differently? Let's just put it at the top of the component
    const compAnchor = 'export function RelatoriosView({';
    code = code.replace(compAnchor, compAnchor + '\n  const [metaDiaDataInicio, setMetaDiaDataInicio] = useState("");\n  const [metaDiaDataFim, setMetaDiaDataFim] = useState("");\n');
}

fs.writeFileSync('src/components/RelatoriosView.tsx', code, 'utf8');
console.log("Fixed RelatoriosView.");
