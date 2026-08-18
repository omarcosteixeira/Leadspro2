const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const adminCallTarget = `                  callBotApi={callBotApi}
                  metaDia={metaDia}`;
const adminCallReplacement = `                  callBotApi={callBotApi}
                  metaDia={metaDia}
                  metaSM={metaSM}`;
code = code.replace(adminCallTarget, adminCallReplacement);

const adminDefTarget = `  callBotApi,
  metaDia,`;
const adminDefReplacement = `  callBotApi,
  metaDia,
  metaSM,`;
code = code.replace(adminDefTarget, adminDefReplacement);

const adminPropsTarget = `  callBotApi: (
    path: string,
    options?: { method?: "GET" | "POST"; body?: any },
  ) => Promise<any>;
  metaDia: MetaDia[];`;
const adminPropsReplacement = `  callBotApi: (
    path: string,
    options?: { method?: "GET" | "POST"; body?: any },
  ) => Promise<any>;
  metaDia: MetaDia[];
  metaSM: MetaSM[];`;
code = code.replace(adminPropsTarget, adminPropsReplacement);

fs.writeFileSync('src/App.tsx', code);
