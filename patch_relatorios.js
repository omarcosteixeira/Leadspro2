import fs from 'fs';
let code = fs.readFileSync('src/components/RelatoriosView.tsx', 'utf8');

code = code.replace(
  'const filteredEmpresasParceiras = useMemo(() => isPrivileged ? empresasParceiras : filteredEmpresasParceiras.filter(e => e.unidadesVinculadas?.includes(profile.unidade || "") || e.unidade === profile.unidade), [empresasParceiras, profile, isPrivileged]);',
  'const filteredEmpresasParceiras = useMemo(() => isPrivileged ? empresasParceiras : empresasParceiras.filter(e => e.unidadesVinculadas?.includes(profile.unidade || "") || e.unidade === profile.unidade), [empresasParceiras, profile, isPrivileged]);'
);

code = code.replace(
  'const filteredPedidosCursos = useMemo(() => isPrivileged ? (pedidosCursos || []) : (pedidosCursos || []), [filteredPedidosCursos, isPrivileged]);',
  'const filteredPedidosCursos = useMemo(() => isPrivileged ? (pedidosCursos || []) : (pedidosCursos || []), [pedidosCursos, isPrivileged]);'
);

code = code.replace(
  'const filteredInsumosPedidos = useMemo(() => isPrivileged ? insumosPedidos : filteredInsumosPedidos.filter(i => i.unidade === profile.unidade), [filteredInsumosPedidos, profile, isPrivileged]);',
  'const filteredInsumosPedidos = useMemo(() => isPrivileged ? insumosPedidos : insumosPedidos, [insumosPedidos, profile, isPrivileged]);'
);

code = code.replace(
  'const filteredInsumosBaixas = useMemo(() => isPrivileged ? insumosBaixas : filteredInsumosBaixas.filter(i => i.unidade === profile.unidade), [filteredInsumosBaixas, profile, isPrivileged]);',
  'const filteredInsumosBaixas = useMemo(() => isPrivileged ? insumosBaixas : insumosBaixas, [insumosBaixas, profile, isPrivileged]);'
);

code = code.replace(
  'const filteredMetaDia = useMemo(() => isPrivileged ? (metaDia || []) : (metaDia || []).filter(m => m.unidade === profile.unidade), [filteredMetaDia, profile, isPrivileged]);',
  'const filteredMetaDia = useMemo(() => isPrivileged ? (metaDia || []) : (metaDia || []), [metaDia, profile, isPrivileged]);'
);

code = code.replace(
  'const filteredSolicitacoesManutencao = useMemo(() => isPrivileged ? (solicitacoesManutencao || []) : (solicitacoesManutencao || []).filter(s => s.unidade === profile.unidade), [filteredSolicitacoesManutencao, profile, isPrivileged]);',
  'const filteredSolicitacoesManutencao = useMemo(() => isPrivileged ? (solicitacoesManutencao || []) : (solicitacoesManutencao || []), [solicitacoesManutencao, profile, isPrivileged]);'
);

code = code.replace(
  'const filteredAnalysisSchemes = useMemo(() => isPrivileged ? (analysisSchemes || []) : (analysisSchemes || []).filter(a => a.unidade === profile.unidade), [filteredAnalysisSchemes, profile, isPrivileged]);',
  'const filteredAnalysisSchemes = useMemo(() => isPrivileged ? (analysisSchemes || []) : (analysisSchemes || []), [analysisSchemes, profile, isPrivileged]);'
);

code = code.replace(
  'if (filteredMetaDiaDataInicio && ',
  'if (metaDiaDataInicio && '
);
code = code.replace(
  'if (filteredMetaDiaDataFim && ',
  'if (metaDiaDataFim && '
);

fs.writeFileSync('src/components/RelatoriosView.tsx', code);
