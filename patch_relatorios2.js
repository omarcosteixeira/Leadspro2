import fs from 'fs';
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

code = code.replace(
  'const filteredEmpresasParceiras = useMemo(() => isPrivileged ? empresasParceiras : empresasParceiras.filter(e => e.unidadesVinculadas?.includes(profile.unidade || "") || e.unidade === profile.unidade), [empresasParceiras, profile, isPrivileged]);',
  'const filteredEmpresasParceiras = useMemo(() => isPrivileged ? empresasParceiras : empresasParceiras.filter(e => e.unidadesVinculadas?.includes(profile.unidade || "") || (e as any).unidade === profile.unidade), [empresasParceiras, profile, isPrivileged]);'
);

code = code.replace(
  '  }, [filteredMetaDia, filteredMetaDiaDataInicio, filteredMetaDiaDataFim]);',
  '  }, [filteredMetaDia, metaDiaDataInicio, metaDiaDataFim]);'
);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
