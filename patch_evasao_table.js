import fs from 'fs';
let code = fs.readFileSync('src/components/EvasaoView.tsx', 'utf8');

const fields = `
                    {item.observacao && (
                      <div className="text-xs text-slate-500 mt-1 line-clamp-2" title={item.observacao}>
                        Obs: {item.observacao}
                      </div>
                    )}
                    {item.parcelaLeve && (
                      <div className="text-xs text-blue-600 mt-1 font-medium">
                        PL: {item.parcelaLeve}
                      </div>
                    )}
                    {item.mensalidades && (
                      <div className="text-xs text-indigo-600 mt-1 font-medium">
                        Mensalidades: {item.mensalidades}
                      </div>
                    )}
                    {item.parcelamento && (
                      <div className="text-xs text-purple-600 mt-1 font-medium">
                        Parcelamento: {item.parcelamento}
                      </div>
                    )}
`;

code = code.replace(
  /\{\s*item\.observacao && \(\s*<div className="text-xs text-slate-500 mt-1 line-clamp-2" title=\{item\.observacao\}>\s*Obs: \{item\.observacao\}\s*<\/div>\s*\)\s*\}/,
  fields.trim()
);

fs.writeFileSync('src/components/EvasaoView.tsx', code);
