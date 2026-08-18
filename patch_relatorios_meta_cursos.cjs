const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

// 1. Add MetaCurso import
const importTarget = `  MetaSM,
  QgLigacao,`;
const importReplace = `  MetaSM,
  MetaCurso,
  QgLigacao,`;
code = code.replace(importTarget, importReplace);

// 2. Add to Props
const propsTarget = `  metaSM?: MetaSM[];
  ligacoes?: Ligacao[];`;
const propsReplace = `  metaSM?: MetaSM[];
  metaCursos?: MetaCurso[];
  ligacoes?: Ligacao[];`;
code = code.replace(propsTarget, propsReplace);

// 3. Add to destructured params
const fnTarget = `  metaSM = [],
  ligacoes = [],`;
const fnReplace = `  metaSM = [],
  metaCursos = [],
  ligacoes = [],`;
code = code.replace(fnTarget, fnReplace);

// 4. Add to activeTab types
const tabTypeTarget = `| "sales" | "metaSM"
  >("historico");`;
const tabTypeReplace = `| "sales" | "metaSM" | "metaCursos"
  >("historico");`;
code = code.replace(tabTypeTarget, tabTypeReplace);

// 5. Add to Tab bar
const tabListTarget = `          { id: "metaSM", label: "Meta SM", icon: Target },`;
const tabListReplace = `          { id: "metaSM", label: "Meta SM", icon: Target },
          { id: "metaCursos", label: "Meta Cursos", icon: Target },`;
code = code.replace(tabListTarget, tabListReplace);

// 6. Add render block
const renderTarget = `        {activeTab === "crescimento" && (`;
const renderReplace = `        {activeTab === "metaCursos" && (
          <div className="space-y-12">
            <div className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm space-y-8">
              <div>
                <h3 className="text-2xl font-bold text-slate-900">Acompanhamento de Meta Cursos</h3>
                <p className="text-slate-500 text-sm">Resumo e projeção de crescimento por curso</p>
              </div>

              {metaCursos.length === 0 ? (
                <div className="text-center py-12 text-slate-400">
                  <Target size={48} className="mx-auto mb-4 opacity-50" />
                  <p>Nenhum dado cadastrado.</p>
                </div>
              ) : (
                <>
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

                  {/* Gráficos de crescimento por curso */}
                  <div className="mt-12 space-y-8">
                    <h4 className="text-xl font-bold text-slate-800 border-b border-slate-200 pb-2">
                      Crescimento Anual por Curso
                    </h4>
                    {Array.from(new Set(metaCursos.map(m => m.curso))).sort().map(cursoNome => {
                      const dadosCurso = metaCursos
                        .filter(m => m.curso === cursoNome)
                        .sort((a,b) => a.semestre.localeCompare(b.semestre));
                      
                      return (
                        <div key={cursoNome} className="border border-slate-100 rounded-2xl p-6 bg-slate-50/50">
                          <h5 className="text-lg font-bold text-slate-700 mb-6">{cursoNome}</h5>
                          <div className="h-[250px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <LineChart data={dadosCurso}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                <XAxis dataKey="semestre" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: "#64748b" }} />
                                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: "#64748b" }} />
                                <Tooltip 
                                  contentStyle={{ borderRadius: "16px", border: "none", boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)" }}
                                />
                                <Legend wrapperStyle={{ paddingTop: "20px" }} />
                                <Line type="monotone" name="Meta Final" dataKey="metaFinal" stroke="#94a3b8" strokeWidth={3} dot={{ r: 4, fill: "white" }} />
                                <Line type="monotone" name="Realizado" dataKey="realizado" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: "white" }} />
                              </LineChart>
                            </ResponsiveContainer>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {activeTab === "crescimento" && (`;
code = code.replace(renderTarget, renderReplace);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
