const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const importMetaSMViewTarget = `import GrowthYearOverYearView from "./components/GrowthYearOverYearView";`;
const importMetaSMViewReplacement = `import GrowthYearOverYearView from "./components/GrowthYearOverYearView";
import MetaSMView from "./components/MetaSMView";`;
code = code.replace(importMetaSMViewTarget, importMetaSMViewReplacement);

const adminRenderTarget = `      {activeTab === "metaDia" && (
        <MetaDiaView metaDia={metaDia} onToast={onToast} />
      )}`;
const adminRenderReplacement = `      {activeTab === "metaDia" && (
        <MetaDiaView metaDia={metaDia} onToast={onToast} />
      )}
      {activeTab === "metaSM" && (
        <MetaSMView metaSM={metaSM} onToast={onToast} />
      )}`;
code = code.replace(adminRenderTarget, adminRenderReplacement);

fs.writeFileSync('src/App.tsx', code);
