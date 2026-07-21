const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

// 1. Add state
const stateAnchor = `  const [ligacoesFiltroAtendente, setLigacoesFiltroAtendente] = useState("");
  const [ligacoesFiltroOrigem, setLigacoesFiltroOrigem] = useState("");
  const [ligacoesSearchTerm, setLigacoesSearchTerm] = useState("");`;
const stateReplacement = `  const [ligacoesFiltroAtendente, setLigacoesFiltroAtendente] = useState("");
  const [ligacoesFiltroOrigem, setLigacoesFiltroOrigem] = useState("");
  const [ligacoesFiltroStatus, setLigacoesFiltroStatus] = useState("");
  const [ligacoesSearchTerm, setLigacoesSearchTerm] = useState("");`;

if (code.includes(stateAnchor)) {
  code = code.replace(stateAnchor, stateReplacement);
} else {
  console.log("Could not find state anchor");
}

// 2. Add filter logic
const filterAnchor = `      if (ligacoesFiltroAtendente && l.atendenteId !== ligacoesFiltroAtendente) return false;
      if (ligacoesFiltroOrigem && l.origemId !== ligacoesFiltroOrigem) return false;
      
      if (ligacoesSearchTerm) {`;
const filterReplacement = `      if (ligacoesFiltroAtendente && l.atendenteId !== ligacoesFiltroAtendente) return false;
      if (ligacoesFiltroOrigem && l.origemId !== ligacoesFiltroOrigem) return false;
      if (ligacoesFiltroStatus && l.status !== ligacoesFiltroStatus) return false;
      
      if (ligacoesSearchTerm) {`;
if (code.includes(filterAnchor)) {
  code = code.replace(filterAnchor, filterReplacement);
} else {
  console.log("Could not find filter anchor");
}

// 3. Update dependencies
const depAnchor = `  }, [ligacoes, ligacoesDataInicio, ligacoesDataFim, ligacoesFiltroAtendente, ligacoesFiltroOrigem, ligacoesSearchTerm]);`;
const depReplacement = `  }, [ligacoes, ligacoesDataInicio, ligacoesDataFim, ligacoesFiltroAtendente, ligacoesFiltroOrigem, ligacoesFiltroStatus, ligacoesSearchTerm]);`;
if (code.includes(depAnchor)) {
  code = code.replace(depAnchor, depReplacement);
} else {
  console.log("Could not find dep anchor");
}

// 4. Update UI
const uiAnchor = `              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">Base / Ação</label>
                <select value={ligacoesFiltroOrigem} onChange={e => setLigacoesFiltroOrigem(e.target.value)} className="w-full text-sm border-slate-200 rounded-lg p-2">
                  <option value="">Todas</option>
                  {origensUnicas.map(o => (
                    <option key={o.id} value={o.id}>{o.name}</option>
                  ))}
                </select>
              </div>
            </div>`;
const uiReplacement = `              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">Base / Ação</label>
                <select value={ligacoesFiltroOrigem} onChange={e => setLigacoesFiltroOrigem(e.target.value)} className="w-full text-sm border-slate-200 rounded-lg p-2">
                  <option value="">Todas</option>
                  {origensUnicas.map(o => (
                    <option key={o.id} value={o.id}>{o.name}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">Status</label>
                <select value={ligacoesFiltroStatus} onChange={e => setLigacoesFiltroStatus(e.target.value)} className="w-full text-sm border-slate-200 rounded-lg p-2">
                  <option value="">Todos</option>
                  <option value="Convertido">Convertido</option>
                  <option value="Interesse">Interesse</option>
                  <option value="Não atendeu">Não atendeu</option>
                  <option value="Sem interesse">Sem interesse</option>
                </select>
              </div>
            </div>`;

if (code.includes(uiAnchor)) {
  code = code.replace(uiAnchor, uiReplacement);
} else {
  console.log("Could not find ui anchor");
}

fs.writeFileSync('src/components/RelatoriosView.tsx', code, 'utf8');
console.log("Done");
