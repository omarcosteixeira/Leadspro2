const fs = require('fs');
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

const target = `        {activeTab === "crescimento" && (`;
const replacement = `        {activeTab === "metaSM" && (
          <div className="space-y-12">
            <div className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm space-y-8">
              <div>
                <h3 className="text-2xl font-bold text-slate-900">Acompanhamento de Meta SM</h3>
                <p className="text-slate-500 text-sm">Resumo e projeção de crescimento do canal SM</p>
              </div>

              {metaSM.length === 0 ? (
                <div className="text-center py-12 text-slate-400">
                  <Target size={48} className="mx-auto mb-4 opacity-50" />
                  <p>Nenhum dado cadastrado.</p>
                </div>
              ) : (
                <>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                    {metaSM.sort((a,b) => b.semestre.localeCompare(a.semestre)).slice(0, 1).map(m => (
                      <React.Fragment key={m.id}>
                        <div className="bg-slate-50 p-6 rounded-2xl border border-slate-100">
                          <p className="text-sm font-bold text-slate-500 uppercase">Semestre {m.semestre}</p>
                          <p className="text-3xl font-black text-slate-900 mt-2">{m.realizado}</p>
                          <p className="text-sm text-slate-400">Realizado SM</p>
                        </div>
                        <div className="bg-blue-50/50 p-6 rounded-2xl border border-blue-100/50">
                          <p className="text-sm font-bold text-slate-500 uppercase">GAP Meta Dia</p>
                          <p className="text-3xl font-black text-blue-700 mt-2">{m.metaDia - m.realizado}</p>
                          <p className="text-sm text-blue-400">Meta: {m.metaDia}</p>
                        </div>
                        <div className="bg-emerald-50/50 p-6 rounded-2xl border border-emerald-100/50">
                          <p className="text-sm font-bold text-slate-500 uppercase">GAP A.A</p>
                          <p className="text-3xl font-black text-emerald-700 mt-2">{m.metaAA - m.realizado}</p>
                          <p className="text-sm text-emerald-400">Ano Anterior: {m.metaAA}</p>
                        </div>
                        <div className="bg-purple-50/50 p-6 rounded-2xl border border-purple-100/50">
                          <p className="text-sm font-bold text-slate-500 uppercase">GAP Final</p>
                          <p className="text-3xl font-black text-purple-700 mt-2">{m.metaFinal - m.realizado}</p>
                          <div className="mt-4 w-full bg-purple-100 rounded-full h-2">
                            <div className="bg-purple-500 h-2 rounded-full transition-all" style={{ width: \`\${Math.min(100, (m.realizado / (m.metaFinal || 1)) * 100)}%\` }}></div>
                          </div>
                          <p className="text-sm text-purple-500 mt-2 font-bold">{((m.realizado / (m.metaFinal || 1)) * 100).toFixed(1)}% Alcançado</p>
                        </div>
                      </React.Fragment>
                    ))}
                  </div>
                  
                  {/* Gráfico de crescimento */}
                  <div className="mt-8 border-t border-slate-100 pt-8">
                    <h4 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                      <TrendingUp size={20} className="text-blue-600" />
                      Crescimento Histórico (Realizado x Final)
                    </h4>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={[...metaSM].sort((a,b) => a.semestre.localeCompare(b.semestre))}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="semestre" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: "#64748b" }} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: "#64748b" }} />
                          <Tooltip 
                            contentStyle={{ borderRadius: "16px", border: "none", boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)" }}
                          />
                          <Legend wrapperStyle={{ paddingTop: "20px" }} />
                          <Line type="monotone" name="Meta Final" dataKey="metaFinal" stroke="#94a3b8" strokeWidth={3} dot={{ r: 4, fill: "white" }} />
                          <Line type="monotone" name="Realizado SM" dataKey="realizado" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: "white" }} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {activeTab === "crescimento" && (`;
code = code.replace(target, replacement);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
