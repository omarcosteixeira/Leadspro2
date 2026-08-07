import fs from 'fs';
let code = fs.readFileSync('src/components/EvasaoView.tsx', 'utf8');

const fields = `
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">Parcela Leve</label>
                    <select
                      value={formData.parcelaLeve || ""}
                      onChange={e => setFormData({...formData, parcelaLeve: e.target.value})}
                      className="w-full px-3 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione...</option>
                      {[0,1,2,3].map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">Mensalidades</label>
                    <select
                      value={formData.mensalidades || ""}
                      onChange={e => setFormData({...formData, mensalidades: e.target.value})}
                      className="w-full px-3 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione...</option>
                      {[0,1,2,3,4,5].map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">Parcelamento</label>
                    <select
                      value={formData.parcelamento || ""}
                      onChange={e => setFormData({...formData, parcelamento: e.target.value})}
                      className="w-full px-3 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione...</option>
                      {[0,1,2,3,4,5,6,7,8,9].map(v => <option key={v} value={v}>{v}</option>)}
                    </select>
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-bold text-slate-700 mb-1">Observação (Opcional)</label>
`;

code = code.replace(
  /<div className="md:col-span-2">\s*<label className="block text-sm font-bold text-slate-700 mb-1">Observação \(Opcional\)<\/label>/,
  fields.trim()
);

fs.writeFileSync('src/components/EvasaoView.tsx', code);
