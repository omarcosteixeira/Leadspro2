const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const target = `      {/* Bom Dia Captação (Complete - All cards) */}`;
const replacement = `      {/* Meta Cursos Dashboard Card */}
      {widgets.metaCursos && metaCursos.length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta Cursos
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {metaCursos.sort((a,b) => b.semestre.localeCompare(a.semestre) || a.curso.localeCompare(b.curso)).map(m => (
              <div key={m.id} className="bg-slate-50 border border-slate-200 rounded-2xl p-5 shadow-sm">
                <div className="mb-4">
                  <span className="bg-slate-200 text-slate-700 text-[10px] font-bold px-2 py-1 rounded-full uppercase tracking-wider">{m.semestre}</span>
                  <h4 className="text-lg font-black text-slate-900 mt-2">{m.curso}</h4>
                </div>
                
                <div className="grid grid-cols-2 gap-3 mb-4">
                  <div className="bg-white p-3 rounded-xl border border-slate-100 shadow-sm">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Realizado</p>
                    <p className="text-xl font-black text-slate-900">{m.realizado}</p>
                  </div>
                  <div className="bg-white p-3 rounded-xl border border-slate-100 shadow-sm">
                    <p className="text-[10px] font-bold text-slate-400 uppercase">Meta Dia</p>
                    <p className="text-xl font-black text-slate-900">{m.metaDia}</p>
                  </div>
                </div>

                <div className="space-y-3">
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-slate-500 font-medium">GAP Meta Dia</span>
                    <span className="font-bold text-blue-600">{m.metaDia - m.realizado}</span>
                  </div>
                  <div className="flex justify-between items-center text-sm border-t border-slate-200 pt-2">
                    <span className="text-slate-500 font-medium">GAP A.A (Meta: {m.metaAA})</span>
                    <span className="font-bold text-emerald-600">{m.metaAA - m.realizado}</span>
                  </div>
                  <div className="flex justify-between items-center text-sm border-t border-slate-200 pt-2">
                    <span className="text-slate-500 font-medium">GAP Final (Meta: {m.metaFinal})</span>
                    <span className="font-bold text-purple-600">{m.metaFinal - m.realizado}</span>
                  </div>
                  
                  <div className="mt-4">
                    <div className="flex justify-between items-center mb-1">
                      <span className="text-xs font-bold text-slate-500 uppercase">Alcance</span>
                      <span className="text-xs font-black text-purple-600">{((m.realizado / (m.metaFinal || 1)) * 100).toFixed(1)}%</span>
                    </div>
                    <div className="w-full bg-slate-200 rounded-full h-2">
                      <div className="bg-purple-500 h-2 rounded-full transition-all" style={{ width: \`\${Math.min(100, (m.realizado / (m.metaFinal || 1)) * 100)}%\` }}></div>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      {/* Bom Dia Captação (Complete - All cards) */}`;

code = code.replace(target, replacement);

fs.writeFileSync('src/App.tsx', code);
