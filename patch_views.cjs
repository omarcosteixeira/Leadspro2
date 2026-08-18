const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const dbTarget = `<DashboardView
                  leads={leads}
                  planner={planner}
                  links={links}
                  profile={profile!}
                  onToast={showToast}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  metaDia={metaDia}
                  qgLigacoes={qgLigacoes}
                  users={users}
                />`;
const dbReplacement = `<DashboardView
                  leads={leads}
                  planner={planner}
                  links={links}
                  profile={profile!}
                  onToast={showToast}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  qgLigacoes={qgLigacoes}
                  users={users}
                />`;
code = code.replace(dbTarget, dbReplacement);

const relTarget = `<RelatoriosView
                  leads={leads}
                  bases={bases}
                  fiesProuni={fiesProuni}
                  calendarioAcoes={calendarioAcoes}
                  empresas={empresasParceiras}
                  periodos={periodos}
                  campanhas={campanhas}
                  onToast={showToast}
                />`;
const relReplacement = `<RelatoriosView
                  leads={leads}
                  bases={bases}
                  fiesProuni={fiesProuni}
                  calendarioAcoes={calendarioAcoes}
                  empresas={empresasParceiras}
                  periodos={periodos}
                  campanhas={campanhas}
                  onToast={showToast}
                  metaSM={metaSM}
                />`;
code = code.replace(relTarget, relReplacement);

fs.writeFileSync('src/App.tsx', code);
