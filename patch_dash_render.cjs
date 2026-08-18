const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const target = `      {/* Bom Dia Captação (Complete - All cards) */}`;
const replacement = `      {/* Meta SM Dashboard Card */}
      {widgets.metaSM && metaSM.length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta SM
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            {metaSM.sort((a,b) => b.semestre.localeCompare(a.semestre)).map(m => (
              <React.Fragment key={m.id}>
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                  <p className="text-xs font-bold text-slate-500 uppercase">Semestre {m.semestre}</p>
                  <p className="text-lg font-black text-slate-900">{m.realizado}</p>
                  <p className="text-xs text-slate-400">Realizado SM</p>
                </div>
                <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Meta Dia</p>
                  <p className="text-lg font-black text-blue-700">{m.metaDia - m.realizado}</p>
                  <p className="text-xs text-blue-400">Meta: {m.metaDia}</p>
                </div>
                <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP A.A</p>
                  <p className="text-lg font-black text-emerald-700">{m.metaAA - m.realizado}</p>
                  <p className="text-xs text-emerald-400">Ano Anterior: {m.metaAA}</p>
                </div>
                <div className="bg-purple-50/50 p-4 rounded-2xl border border-purple-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Final</p>
                  <p className="text-lg font-black text-purple-700">{m.metaFinal - m.realizado}</p>
                  <div className="mt-2 w-full bg-purple-100 rounded-full h-1.5">
                    <div className="bg-purple-500 h-1.5 rounded-full" style={{ width: \`\${Math.min(100, (m.realizado / (m.metaFinal || 1)) * 100)}%\` }}></div>
                  </div>
                  <p className="text-xs text-purple-500 mt-1">{((m.realizado / (m.metaFinal || 1)) * 100).toFixed(1)}% Alcançado</p>
                </div>
              </React.Fragment>
            ))}
          </div>
        </section>
      )}

      {/* Bom Dia Captação (Complete - All cards) */}`;

code = code.replace(target, replacement);

fs.writeFileSync('src/App.tsx', code);
