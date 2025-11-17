// Dados dos registros de consultas médicas carregados do Excel
let allPatients = [];

// Função para mostrar/ocultar indicador de carregamento
function showLoading(show) {
    const loadingIndicator = document.getElementById('loadingIndicator');
    if (loadingIndicator) {
        loadingIndicator.style.display = show ? 'flex' : 'none';
    }
}

// Função para mostrar mensagem de erro
function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    if (errorDiv) {
        errorDiv.innerHTML = `<i class="fas fa-exclamation-triangle"></i> ${message}`;
        errorDiv.style.display = 'flex';
        setTimeout(() => {
            errorDiv.style.display = 'none';
        }, 5000);
    }
    console.error(message);
}

// Função para mostrar mensagem de sucesso
function showSuccess(message) {
    const errorDiv = document.getElementById('errorMessage');
    if (errorDiv) {
        errorDiv.className = 'success-message';
        errorDiv.innerHTML = `<i class="fas fa-check-circle"></i> ${message}`;
        errorDiv.style.display = 'flex';
        setTimeout(() => {
            errorDiv.style.display = 'none';
            errorDiv.className = 'error-message';
        }, 3000);
    }
}

// Função para processar arquivo Excel carregado via input
function handleFileUpload(file) {
    showLoading(true);
    const errorDiv = document.getElementById('errorMessage');
    if (errorDiv) {
        errorDiv.style.display = 'none';
    }
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (processExcelData(workbook)) {
                showLoading(false);
                showSuccess(`Arquivo carregado com sucesso! ${allPatients.length} registros processados.`);
                initializeCharts();
            } else {
                throw new Error('Não foi possível processar os dados do Excel. Verifique se o arquivo tem o formato correto.');
            }
        } catch (error) {
            showLoading(false);
            showError(`Erro ao processar arquivo: ${error.message}`);
            console.error('Erro ao processar arquivo:', error);
        }
    };
    
    reader.onerror = function() {
        showLoading(false);
        showError('Erro ao ler o arquivo. Por favor, tente novamente.');
    };
    
    reader.readAsArrayBuffer(file);
}

// Função auxiliar para converter número (trata ponto como separador de milhar)
function parseNumber(value) {
    if (value === null || value === undefined || value === '') return 0;
    
    // Se já é um número, retornar diretamente
    if (typeof value === 'number') {
        return isNaN(value) ? 0 : value;
    }
    
    const str = String(value).trim();
    
    // Se a string está vazia, retornar 0
    if (str === '' || str === '-') return 0;
    
    // Remover pontos (separadores de milhar) e substituir vírgula por ponto (decimal)
    const cleaned = str.replace(/\./g, '').replace(',', '.');
    const parsed = parseFloat(cleaned);
    
    // Log para debug se o número for muito grande ou inválido
    if (isNaN(parsed)) {
        console.warn(`parseNumber: valor inválido "${value}" -> "${str}" -> "${cleaned}"`);
        return 0;
    }
    
    if (parsed > 1000000000) {
        console.warn(`parseNumber: número muito grande "${value}" -> ${parsed}`);
    }
    
    return parsed;
}

// Função auxiliar para extrair faixa etária média
function getAgeFromRange(ageRange) {
    if (!ageRange) return null;
    const str = String(ageRange).trim().toLowerCase();
    
    // Padrões: "0 a 4 anos" -> 2, "5 a 19 anos" -> 12, "20 a 39 anos" -> 29.5, etc.
    const match = str.match(/(\d+)\s*a\s*(\d+)/);
    if (match) {
        const min = parseInt(match[1]);
        const max = parseInt(match[2]);
        return Math.round((min + max) / 2);
    }
    
    // "65 anos ou mais" -> 70
    if (str.includes('65') && (str.includes('mais') || str.includes('ou'))) {
        return 70;
    }
    
    // "Idade ignorada" -> null
    if (str.includes('ignorada') || str.includes('ignorado')) {
        return null;
    }
    
    return null;
}

// Armazenar dados agregados de cada tabela (estrutura completa)
let aggregatedData = {
    idade: [],
    sexo: [],
    domicilio: [],
    rendimento: []
};

// Função para processar dados do Excel e converter para o formato esperado
function processExcelData(workbook) {
    // Limpar dados anteriores
    aggregatedData = {
        idade: [],
        sexo: [],
        domicilio: [],
        rendimento: []
    };
    
    console.log('Planilhas disponíveis:', workbook.SheetNames);
    
    const sheetName = workbook.SheetNames[0];
    console.log(`Processando planilha: ${sheetName}`);
    const worksheet = workbook.Sheets[sheetName];
    
    // Converter para JSON - array de arrays
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length < 1) {
        console.error('Planilha está vazia');
        return false;
    }
    
    console.log(`Total de linhas na planilha: ${jsonData.length}`);
    console.log('Primeiras 15 linhas:', jsonData.slice(0, 15));
    
    // Processar cada tabela individualmente e completamente
    let currentTable = null;
    let tableStartRow = -1;
    let headerRow = null;
    
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) {
            // Linha vazia pode indicar fim de tabela
            if (currentTable) {
                console.log(`  Fim da tabela ${currentTable} detectado na linha ${i + 1} (linha vazia)`);
                currentTable = null;
                tableStartRow = -1;
            }
            continue;
        }
        
        const firstCell = String(row[0] || '').trim();
        const firstCellLower = firstCell.toLowerCase();
        const secondCell = String(row[1] || '').trim().toLowerCase();
        
        // Detectar início de nova tabela (isso reseta a tabela anterior)
        if (firstCellLower === 'idade' || (firstCellLower.includes('idade') && firstCellLower.length < 15 && !firstCellLower.includes('ignorada'))) {
            if (currentTable) {
                console.log(`  Fim da tabela ${currentTable} na linha ${i} (nova tabela detectada)`);
            }
            currentTable = 'idade';
            tableStartRow = i;
            headerRow = null;
            console.log(`\n=== TABELA IDADE encontrada na linha ${i + 1} ===`);
            continue;
        } else if (firstCellLower === 'sexo' || firstCellLower === 'genero' || 
                   (firstCellLower.includes('sexo') && firstCellLower.length < 10) ||
                   (firstCellLower.includes('genero') && firstCellLower.length < 10) ||
                   (firstCellLower.includes('gênero') && firstCellLower.length < 10)) {
            if (currentTable) {
                console.log(`  Fim da tabela ${currentTable} na linha ${i} (nova tabela detectada)`);
            }
            currentTable = 'sexo';
            tableStartRow = i;
            headerRow = null;
            console.log(`\n=== TABELA SEXO/GENERO encontrada na linha ${i + 1} ===`);
            continue;
        } else if (firstCellLower.includes('domicílio') || firstCellLower.includes('domicilio') ||
                   firstCellLower.includes('situação') || firstCellLower.includes('situacao')) {
            if (currentTable) {
                console.log(`  Fim da tabela ${currentTable} na linha ${i} (nova tabela detectada)`);
            }
            currentTable = 'domicilio';
            tableStartRow = i;
            headerRow = null;
            console.log(`\n=== TABELA DOMICÍLIO encontrada na linha ${i + 1} ===`);
            continue;
        } else if (firstCellLower.includes('rendimento') || firstCellLower.includes('renda') ||
                   (firstCellLower.includes('classe') && firstCellLower.length > 10)) {
            if (currentTable) {
                console.log(`  Fim da tabela ${currentTable} na linha ${i} (nova tabela detectada)`);
            }
            currentTable = 'rendimento';
            tableStartRow = i;
            headerRow = null;
            console.log(`\n=== TABELA RENDIMENTO encontrada na linha ${i + 1} ===`);
            continue;
        }
        
        // Se estamos dentro de uma tabela, processar os dados
        if (currentTable && i > tableStartRow) {
            // Verificar se é linha de cabeçalho
            if (secondCell.includes('não consultou') || secondCell.includes('nao consultou') || 
                secondCell.includes('consultou') || secondCell === 'total') {
                headerRow = row;
                console.log(`  Cabeçalho: ${row.join(' | ')}`);
                continue;
            }
            
            const categoria = firstCell;
            
            // Ignorar linhas de total geral (não são categorias)
            if (categoria && (categoria.toLowerCase().includes('total') || categoria.toLowerCase() === 'total geral')) {
                console.log(`  Ignorando linha de total: ${categoria}`);
                continue;
            }
            
            // Ler valores brutos antes de parsear
            const rawNaoConsultou = row[1];
            const rawConsultou = row[2];
            const rawTotal = row[3];
            
            const naoConsultou = parseNumber(rawNaoConsultou);
            const consultou = parseNumber(rawConsultou);
            const total = parseNumber(rawTotal);
            
            // Log detalhado para debug
            if (total > 0 || consultou > 0 || naoConsultou > 0) {
                console.log(`  Linha ${i + 1}: "${categoria}" | Raw: [${rawNaoConsultou}, ${rawConsultou}, ${rawTotal}] | Parsed: [${naoConsultou}, ${consultou}, ${total}]`);
            }
            
            // Verificar se a linha tem dados válidos
            if (total === 0 && naoConsultou === 0 && consultou === 0) {
                // Pode ser fim da tabela ou linha vazia
                if (categoria.length === 0) {
                    continue;
                }
            }
            
            if (!categoria || categoria.length === 0) continue;
            
            // Validar que os números fazem sentido (total deve ser >= consultou + naoConsultou)
            if (total > 0 && Math.abs(total - (consultou + naoConsultou)) > 1) {
                console.warn(`  AVISO: Linha ${i + 1} - Total (${total}) não corresponde à soma (${consultou} + ${naoConsultou} = ${consultou + naoConsultou})`);
            }
            
            // Processar cada tabela e armazenar dados agregados
            if (currentTable === 'idade') {
                const age = getAgeFromRange(categoria);
                if (age !== null) {
                    aggregatedData.idade.push({ 
                        categoria, 
                        age, 
                        naoConsultou, 
                        consultou, 
                        total 
                    });
                    console.log(`  [${aggregatedData.idade.length}] ${categoria} (${age} anos) - Não: ${naoConsultou}, Sim: ${consultou}, Total: ${total}`);
                }
            } else if (currentTable === 'sexo') {
                // Detectar gênero: Homens/Masculino = M, Mulheres/Feminino = F
                const categoriaLower = categoria.toLowerCase();
                let gender = 'F'; // Default para feminino
                if (categoriaLower.includes('homem') || 
                    categoriaLower.includes('masculino') ||
                    categoriaLower.includes('homens') ||
                    categoriaLower === 'm' ||
                    categoriaLower.startsWith('hom')) {
                    gender = 'M';
                } else if (categoriaLower.includes('mulher') || 
                          categoriaLower.includes('feminino') ||
                          categoriaLower.includes('mulheres') ||
                          categoriaLower === 'f' ||
                          categoriaLower.startsWith('mulh')) {
                    gender = 'F';
                }
                
                aggregatedData.sexo.push({ 
                    categoria, 
                    gender, 
                    naoConsultou, 
                    consultou, 
                    total 
                });
                console.log(`  [${aggregatedData.sexo.length}] ${categoria} -> gender: ${gender} - Não: ${naoConsultou}, Sim: ${consultou}, Total: ${total}`);
            } else if (currentTable === 'domicilio') {
                const domicilio = categoria.toLowerCase().includes('rural') ? 'rural' : 'urbano';
                aggregatedData.domicilio.push({ 
                    categoria, 
                    domicilio, 
                    naoConsultou, 
                    consultou, 
                    total 
                });
                console.log(`  [${aggregatedData.domicilio.length}] ${categoria} (${domicilio}) - Não: ${naoConsultou}, Sim: ${consultou}, Total: ${total}`);
            } else if (currentTable === 'rendimento') {
                aggregatedData.rendimento.push({ 
                    categoria, 
                    rendimento: categoria, 
                    naoConsultou, 
                    consultou, 
                    total 
                });
                console.log(`  [${aggregatedData.rendimento.length}] ${categoria} - Não: ${naoConsultou}, Sim: ${consultou}, Total: ${total}`);
            }
        }
    }
    
    // Calcular totais de cada tabela para validação
    const totalIdade = aggregatedData.idade.reduce((sum, item) => sum + item.total, 0);
    const totalSexo = aggregatedData.sexo.reduce((sum, item) => sum + item.total, 0);
    const totalDomicilio = aggregatedData.domicilio.reduce((sum, item) => sum + item.total, 0);
    const totalRendimento = aggregatedData.rendimento.reduce((sum, item) => sum + item.total, 0);
    
    console.log(`\n=== RESUMO FINAL ===`);
    console.log(`Dados agregados por tabela:`, {
        idade: `${aggregatedData.idade.length} categorias - Total: ${totalIdade.toLocaleString('pt-BR')}`,
        sexo: `${aggregatedData.sexo.length} categorias - Total: ${totalSexo.toLocaleString('pt-BR')}`,
        domicilio: `${aggregatedData.domicilio.length} categorias - Total: ${totalDomicilio.toLocaleString('pt-BR')}`,
        rendimento: `${aggregatedData.rendimento.length} categorias - Total: ${totalRendimento.toLocaleString('pt-BR')}`
    });
    
    // Validar consistência entre tabelas
    const totals = [totalIdade, totalSexo, totalDomicilio, totalRendimento].filter(t => t > 0);
    if (totals.length > 1) {
        const maxTotal = Math.max(...totals);
        const minTotal = Math.min(...totals);
        if (Math.abs(maxTotal - minTotal) > 100) {
            console.warn(`Aviso: Totais das tabelas não coincidem! Diferença: ${maxTotal - minTotal}`);
        } else {
            console.log(`✓ Totais das tabelas são consistentes: ${maxTotal.toLocaleString('pt-BR')}`);
        }
    }
    
    // Criar registros individuais apenas para compatibilidade com filtros (amostra pequena)
    createPatientsFromAggregatedData();
    
    return aggregatedData.idade.length > 0 || aggregatedData.sexo.length > 0 || 
           aggregatedData.domicilio.length > 0 || aggregatedData.rendimento.length > 0;
}

// Função para criar registros individuais a partir dos dados agregados (para compatibilidade com filtros)
// Esta função cria uma amostra representativa, mas os gráficos usam dados agregados diretamente
function createPatientsFromAggregatedData() {
    allPatients = [];
    
    // Calcular total geral (usar a tabela que tiver o maior total)
    const totalIdade = aggregatedData.idade.reduce((sum, item) => sum + item.total, 0);
    const totalSexo = aggregatedData.sexo.reduce((sum, item) => sum + item.total, 0);
    const totalDomicilio = aggregatedData.domicilio.reduce((sum, item) => sum + item.total, 0);
    const totalRendimento = aggregatedData.rendimento.reduce((sum, item) => sum + item.total, 0);
    
    // Usar o total real da tabela (não o maior, mas o que representa melhor os dados)
    // Normalmente todas as tabelas devem ter o mesmo total, mas vamos usar a de idade como referência
    const totalGeral = totalIdade > 0 ? totalIdade : Math.max(totalSexo, totalDomicilio, totalRendimento);
    
    // Limitar a criação de registros para não gerar números enormes
    // Criar apenas uma amostra representativa (máximo 10.000 registros)
    const maxSamples = 10000;
    const sampleFactor = totalGeral > maxSamples ? maxSamples / totalGeral : 1;
    
    // Criar registros baseados na tabela de idade (mais detalhada)
    // IMPORTANTE: Criar apenas uma amostra pequena para filtros, não todos os registros
    if (aggregatedData.idade.length > 0) {
        for (const idadeItem of aggregatedData.idade) {
            const age = idadeItem.age;
            
            // Criar amostra proporcional para "não consultou" (limitado)
            const naoConsultouSamples = Math.max(1, Math.round(idadeItem.naoConsultou * sampleFactor));
            for (let i = 0; i < naoConsultouSamples && allPatients.length < maxSamples; i++) {
                allPatients.push({
                    age: age,
                    gender: 'M',
                    hasConsulta: false,
                    numConsultas: 0,
                    domicilio: 'urbano',
                    rendimento: 'sem rendimento'
                });
            }
            
            // Criar amostra proporcional para "consultou" (limitado)
            const consultouSamples = Math.max(1, Math.round(idadeItem.consultou * sampleFactor));
            for (let i = 0; i < consultouSamples && allPatients.length < maxSamples; i++) {
                allPatients.push({
                    age: age,
                    gender: 'M',
                    hasConsulta: true,
                    numConsultas: 1,
                    domicilio: 'urbano',
                    rendimento: 'sem rendimento'
                });
            }
        }
        
        // Distribuir por sexo baseado na tabela de sexo
        if (aggregatedData.sexo.length > 0 && allPatients.length > 0) {
            let index = 0;
            for (const sexoItem of aggregatedData.sexo) {
                const gender = sexoItem.gender;
                const proporcaoSexo = totalGeral > 0 ? sexoItem.total / totalGeral : 0;
                const count = Math.round(allPatients.length * proporcaoSexo);
                
                for (let i = 0; i < count && index < allPatients.length; i++) {
                    allPatients[index].gender = gender;
                    index++;
                }
            }
        }
        
        // Distribuir por domicílio baseado na tabela de domicílio
        if (aggregatedData.domicilio.length > 0 && allPatients.length > 0) {
            let index = 0;
            for (const domicilioItem of aggregatedData.domicilio) {
                const domicilio = domicilioItem.domicilio;
                const proporcaoDomicilio = totalGeral > 0 ? domicilioItem.total / totalGeral : 0;
                const count = Math.round(allPatients.length * proporcaoDomicilio);
                
                for (let i = 0; i < count && index < allPatients.length; i++) {
                    allPatients[index].domicilio = domicilio;
                    index++;
                }
            }
        }
        
        // Distribuir por rendimento baseado na tabela de rendimento
        if (aggregatedData.rendimento.length > 0 && allPatients.length > 0) {
            let index = 0;
            for (const rendimentoItem of aggregatedData.rendimento) {
                const rendimento = rendimentoItem.rendimento;
                const proporcaoRendimento = totalGeral > 0 ? rendimentoItem.total / totalGeral : 0;
                const count = Math.round(allPatients.length * proporcaoRendimento);
                
                for (let i = 0; i < count && index < allPatients.length; i++) {
                    allPatients[index].rendimento = rendimento;
                    index++;
                }
            }
        }
    }
    
    console.log(`Amostra criada para filtros: ${allPatients.length} registros (de ${totalGeral} totais)`);
}

// Função para carregar arquivo Excel
async function loadExcelFile() {
    showLoading(true);
    
    // Detectar se está rodando localmente (file://) ou em servidor web
    const isLocalFile = window.location.protocol === 'file:';
    
    // Tentar ambas as extensões (.xlsx e .xls)
    // Usar caminho relativo que funciona tanto localmente quanto em servidor web (Vercel)
    const possiblePaths = [
        'Analise/Tabela 6.1 - 2-2.xlsx',
        './Analise/Tabela 6.1 - 2-2.xlsx',
        '/Analise/Tabela 6.1 - 2-2.xlsx',
        'Analise/Tabela 6.1 - 2-2.xls',
        './Analise/Tabela 6.1 - 2-2.xls',
        '/Analise/Tabela 6.1 - 2-2.xls'
    ];
    
    let lastError = null;
    
    for (const filePath of possiblePaths) {
        try {
            console.log(`Tentando carregar: ${filePath}`);
            const response = await fetch(filePath);
            
            if (!response.ok) {
                throw new Error(`Erro ao carregar arquivo: ${response.status}`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            
            if (processExcelData(workbook)) {
                showLoading(false);
                console.log('Arquivo Excel carregado com sucesso!');
                // Mostrar mensagem de sucesso apenas se não estiver em modo silencioso
                if (!isLocalFile) {
                    showSuccess(`Arquivo carregado automaticamente`);
                }
                // Inicializar gráficos com os dados carregados
                initializeCharts();
                return; // Sucesso, sair da função
            } else {
                throw new Error('Não foi possível processar os dados do Excel');
            }
        } catch (error) {
            console.warn(`Falha ao carregar ${filePath}:`, error.message);
            lastError = error;
            // Continuar para tentar o próximo caminho
            continue;
        }
    }
    
    // Se chegou aqui, nenhum arquivo foi carregado com sucesso
    showLoading(false);
    console.error('Erro ao carregar arquivo Excel:', lastError);
    
    // Mensagem de erro adaptada ao ambiente
    let errorMsg;
    if (isLocalFile) {
        errorMsg = `Para carregar o arquivo automaticamente, é necessário usar um servidor web. ` +
                  `Use o botão "Selecionar Arquivo Excel" acima para fazer upload manual do arquivo.`;
    } else {
        errorMsg = `Não foi possível carregar o arquivo automaticamente. ` +
                  `Por favor, use o botão "Selecionar Arquivo Excel" acima para fazer upload manual do arquivo. ` +
                  `Erro: ${lastError ? lastError.message : 'Arquivo não encontrado'}`;
    }
    showError(errorMsg);
}

// Função para gerar dados simulados (fallback)
function generatePatientData() {
    const ageGroups = [
        { range: [20, 30], count: 30 },
        { range: [31, 40], count: 45 },
        { range: [41, 50], count: 60 },
        { range: [51, 60], count: 55 },
        { range: [61, 70], count: 40 },
        { range: [71, 100], count: 25 }
    ];
    
    allPatients = [];
    
    ageGroups.forEach(group => {
        for (let i = 0; i < group.count; i++) {
            const age = Math.floor(Math.random() * (group.range[1] - group.range[0] + 1)) + group.range[0];
            const gender = Math.random() > 0.45 ? 'M' : 'F';
            const hasConsulta = Math.random() > 0.3; // 70% tem consulta
            
            allPatients.push({
                age,
                gender,
                hasConsulta
            });
        }
    });
}

// Filtrar pacientes baseado nos filtros ativos
function getFilteredPatients() {
    const ageFilter = document.getElementById('ageFilter').value;
    const diseaseFilter = document.getElementById('diseaseFilter').value;
    
    return allPatients.filter(patient => {
        // Filtro de idade
        if (ageFilter !== 'all') {
            if (ageFilter === 'ignorada') {
                // Idade ignorada - verificar se age é null
                if (patient.age !== null) return false;
            } else if (ageFilter === '65+') {
                // 65 anos ou mais
                if (patient.age === null || patient.age < 65) return false;
            } else {
                const [min, max] = ageFilter.split('-').map(Number);
                if (patient.age === null || patient.age < min || patient.age > max) return false;
            }
        }
        
        // Filtro de consulta
        if (diseaseFilter !== 'all') {
            const hasConsultaFilter = diseaseFilter === 'yes';
            if (patient.hasConsulta !== hasConsultaFilter) return false;
        }
        
        return true;
    });
}

// Função auxiliar para verificar se uma faixa etária corresponde ao filtro
function ageMatchesFilter(age, ageFilter) {
    if (ageFilter === 'all') return true;
    if (ageFilter === 'ignorada') return age === null;
    if (ageFilter === '65+') return age !== null && age >= 65;
    
    const [min, max] = ageFilter.split('-').map(Number);
    return age !== null && age >= min && age <= max;
}

// Função auxiliar para obter faixa etária do filtro
function getAgeRangeFromFilter(ageFilter) {
    if (ageFilter === 'all') return null;
    if (ageFilter === 'ignorada') return { min: null, max: null, ignorada: true };
    if (ageFilter === '65+') return { min: 65, max: 150 }; // 65 anos ou mais
    const [min, max] = ageFilter.split('-').map(Number);
    return { min, max };
}

// Calcular dados dos gráficos usando dados agregados diretamente, aplicando filtros
function calculateChartData(patients) {
    // Obter valores dos filtros
    const ageFilter = document.getElementById('ageFilter')?.value || 'all';
    const diseaseFilter = document.getElementById('diseaseFilter')?.value || 'all';
    
    // SEMPRE usar dados agregados diretamente - não usar pacientes individuais para cálculos
    let total = 0;
    let withConsulta = 0;
    let withoutConsulta = 0;
    let totalNumConsultas = 0;
    
    // IMPORTANTE: Usar apenas UMA tabela para calcular o total geral
    // Todas as tabelas representam a mesma população, apenas segmentada de forma diferente
    // Usar a tabela de idade como referência principal (mais detalhada)
    
    // Aplicar filtros aos dados agregados de idade
    let filteredIdadeData = aggregatedData.idade;
    if (ageFilter !== 'all') {
        const ageRange = getAgeRangeFromFilter(ageFilter);
        if (ageRange) {
            if (ageRange.ignorada) {
                // Filtrar apenas itens com idade ignorada (age === null)
                filteredIdadeData = aggregatedData.idade.filter(item => item.age === null);
            } else {
                filteredIdadeData = aggregatedData.idade.filter(item => {
                    return item.age !== null && item.age >= ageRange.min && item.age <= ageRange.max;
                });
            }
        }
    }
    
    // Resetar variáveis antes de calcular
    total = 0;
    withConsulta = 0;
    withoutConsulta = 0;
    totalNumConsultas = 0;
    
    if (filteredIdadeData.length > 0) {
        filteredIdadeData.forEach((item, index) => {
            const itemTotal = Number(item.total) || 0;
            const itemConsultou = Number(item.consultou) || 0;
            const itemNaoConsultou = Number(item.naoConsultou) || 0;
            
            total += itemTotal;
            withConsulta += itemConsultou;
            withoutConsulta += itemNaoConsultou;
            totalNumConsultas += itemConsultou; // Cada pessoa que consultou = 1 consulta
        });
    } else {
        // Fallback apenas se não houver dados agregados (não deveria acontecer)
        console.warn('⚠ Usando dados individuais como fallback - dados agregados não disponíveis');
        total = patients.length;
        withConsulta = patients.filter(p => p.hasConsulta).length;
        withoutConsulta = total - withConsulta;
        totalNumConsultas = patients.reduce((sum, p) => sum + (p.numConsultas || 0), 0);
    }
    
    // Garantir que são números
    total = Number(total) || 0;
    withConsulta = Number(withConsulta) || 0;
    withoutConsulta = Number(withoutConsulta) || 0;
    totalNumConsultas = Number(totalNumConsultas) || 0;
    
    console.log('\n=== TOTAIS CALCULADOS ===');
    console.log('Total de registros:', total.toLocaleString('pt-BR'), `(${total})`);
    console.log('Com consulta:', withConsulta.toLocaleString('pt-BR'), `(${withConsulta})`);
    console.log('Sem consulta:', withoutConsulta.toLocaleString('pt-BR'), `(${withoutConsulta})`);
    console.log('Total de consultas:', totalNumConsultas.toLocaleString('pt-BR'), `(${totalNumConsultas})`);
    console.log('Verificação:', total, '=', withConsulta, '+', withoutConsulta, '?', (Math.abs(total - (withConsulta + withoutConsulta)) < 1));
    
    if (total > 100000000) {
        console.error('❌ ERRO: Total muito grande! Algo está errado na leitura dos dados.');
        console.error('Dados agregados completos:', JSON.stringify(aggregatedData, null, 2));
    }
    
    // Validar que os totais fazem sentido
    if (Math.abs(total - (withConsulta + withoutConsulta)) > 1) {
        console.warn('Aviso: Total não corresponde à soma. Verificando dados...');
        console.log('Dados agregados completos:', aggregatedData);
    }
    
    // Calcular dados por sexo a partir da tabela de sexo (genero)
    let menTotal = 0;
    let womenTotal = 0;
    let menWithConsulta = 0;
    let womenWithConsulta = 0;
    
    // SEMPRE usar dados da tabela genero (sexo)
    if (aggregatedData.sexo.length > 0) {
        console.log('\n=== PROCESSANDO DADOS DE SEXO/GENERO ===');
        console.log('Total de itens na tabela sexo:', aggregatedData.sexo.length);
        aggregatedData.sexo.forEach((item, index) => {
            console.log(`  Item ${index + 1}: ${item.categoria} -> gender: ${item.gender}, Total: ${item.total}, Consultou: ${item.consultou}`);
            if (item.gender === 'M') {
                menTotal += item.total;
                menWithConsulta += item.consultou;
            } else if (item.gender === 'F') {
                womenTotal += item.total;
                womenWithConsulta += item.consultou;
            } else {
                console.warn(`  AVISO: Gênero desconhecido "${item.gender}" para categoria "${item.categoria}"`);
            }
        });
        console.log(`Resultado: Homens (M) - Total: ${menTotal}, Com Consulta: ${menWithConsulta}`);
        console.log(`Resultado: Mulheres (F) - Total: ${womenTotal}, Com Consulta: ${womenWithConsulta}`);
    } else {
        console.warn('⚠ AVISO: Tabela sexo/genero está vazia! Usando fallback.');
        // Fallback apenas se não houver dados agregados
        const men = patients.filter(p => p.gender === 'M');
        const women = patients.filter(p => p.gender === 'F');
        menTotal = men.length;
        womenTotal = women.length;
        menWithConsulta = men.filter(p => p.hasConsulta).length;
        womenWithConsulta = women.filter(p => p.hasConsulta).length;
    }
    
    // Calcular dados por domicílio a partir da tabela de domicílio
    let urbano = 0;
    let rural = 0;
    
    if (aggregatedData.domicilio.length > 0) {
        aggregatedData.domicilio.forEach(item => {
            if (item.domicilio === 'urbano') {
                urbano += item.total;
            } else {
                rural += item.total;
            }
        });
    } else {
        // Fallback
        urbano = patients.filter(p => p.domicilio === 'urbano').length;
        rural = patients.filter(p => p.domicilio === 'rural').length;
    }
    
    // Agrupar por classes de rendimento a partir da tabela de rendimento
    const rendimentoGroups = {};
    if (aggregatedData.rendimento.length > 0) {
        aggregatedData.rendimento.forEach(item => {
            rendimentoGroups[item.rendimento] = item.total;
        });
    } else {
        // Fallback
        patients.forEach(p => {
            const rend = p.rendimento || 'sem rendimento';
            if (!rendimentoGroups[rend]) {
                rendimentoGroups[rend] = 0;
            }
            rendimentoGroups[rend]++;
        });
    }
    
    // Calcular idade média a partir da tabela de idade
    let avgAge = 0;
    if (aggregatedData.idade.length > 0) {
        let totalAge = 0;
        let totalCount = 0;
        aggregatedData.idade.forEach(item => {
            totalAge += item.age * item.total;
            totalCount += item.total;
        });
        avgAge = totalCount > 0 ? Math.round(totalAge / totalCount) : 0;
    } else {
        // Fallback
        avgAge = patients.length > 0 
            ? Math.round(patients.reduce((sum, p) => sum + p.age, 0) / patients.length)
            : 0;
    }
    
    // Distribuição
    const distributionData = {
        labels: ['Sem Consulta Médica', 'Com Consulta Médica'],
        datasets: [{
            data: [withoutConsulta, withConsulta],
            backgroundColor: ['#3498db', '#2ecc71'],
            borderWidth: 0
        }]
    };
    
    // Distribuição por sexo - dados vêm diretamente da tabela sexo/genero da planilha
    // Os valores menTotal e womenTotal são calculados a partir de aggregatedData.sexo
    const riskFactorsData = {
        labels: ['Masculino', 'Feminino'],
        datasets: [{
            label: 'Total de Registros',
            data: [menTotal, womenTotal],
            backgroundColor: [
                '#1a5490',
                '#e91e63'
            ],
            borderWidth: 0
        }]
    };
    
    // Distribuição por situação do domicílio
    const domicilioData = {
        labels: ['Urbano', 'Rural'],
        datasets: [{
            data: [urbano, rural],
            backgroundColor: ['#2ecc71', '#f39c12'],
            borderWidth: 0
        }]
    };
    
    // Distribuição por classes de rendimento
    const rendimentoLabels = Object.keys(rendimentoGroups).sort();
    const rendimentoValues = rendimentoLabels.map(label => rendimentoGroups[label]);
    const rendimentoData = {
        labels: rendimentoLabels.map(l => l.length > 30 ? l.substring(0, 30) + '...' : l),
        datasets: [{
            label: 'Número de Registros',
            data: rendimentoValues,
            backgroundColor: [
                '#3498db', '#2ecc71', '#f39c12', '#e74c3c', 
                '#9b59b6', '#1abc9c', '#34495e'
            ],
            borderWidth: 0
        }]
    };
    
    // Segmentação por idade - aplicar filtros aos dados agregados
    // Determinar quais datasets mostrar baseado no filtro de consulta médica
    let ageSegmentationData = {
        labels: [],
        datasets: []
    };
    
    // Se o filtro for "all", mostrar ambos os datasets
    if (diseaseFilter === 'all') {
        ageSegmentationData.datasets = [
            {
                label: 'Com Consulta Médica',
                data: [],
                backgroundColor: '#2ecc71',
                borderWidth: 0
            },
            {
                label: 'Sem Consulta Médica',
                data: [],
                backgroundColor: '#3498db',
                borderWidth: 0
            }
        ];
    } else if (diseaseFilter === 'yes') {
        // Apenas "Com Consulta Médica"
        ageSegmentationData.datasets = [
            {
                label: 'Com Consulta Médica',
                data: [],
                backgroundColor: '#2ecc71',
                borderWidth: 0
            }
        ];
    } else if (diseaseFilter === 'no') {
        // Apenas "Sem Consulta Médica"
        ageSegmentationData.datasets = [
            {
                label: 'Sem Consulta Médica',
                data: [],
                backgroundColor: '#3498db',
                borderWidth: 0
            }
        ];
    }
    
    let ageGroupStats = [];
    
    if (aggregatedData.idade.length > 0) {
        // Aplicar filtros aos dados de idade para o gráfico
        let filteredIdadeForChart = aggregatedData.idade;
        
        // Filtro de idade (se aplicado, mostra apenas a faixa selecionada)
        if (ageFilter !== 'all') {
            const ageRange = getAgeRangeFromFilter(ageFilter);
            if (ageRange) {
                if (ageRange.ignorada) {
                    // Filtrar apenas itens com idade ignorada (age === null)
                    filteredIdadeForChart = aggregatedData.idade.filter(item => item.age === null);
                } else {
                    filteredIdadeForChart = aggregatedData.idade.filter(item => {
                        return item.age !== null && item.age >= ageRange.min && item.age <= ageRange.max;
                    });
                }
            }
        }
        
        // Usar dados filtrados da tabela de idade
        ageSegmentationData.labels = filteredIdadeForChart.map(item => item.categoria);
        
        // Preencher datasets baseado no filtro de consulta médica
        if (diseaseFilter === 'all') {
            // Mostrar ambos os datasets
            ageSegmentationData.datasets[0].data = filteredIdadeForChart.map(item => item.consultou);
            ageSegmentationData.datasets[1].data = filteredIdadeForChart.map(item => item.naoConsultou);
        } else if (diseaseFilter === 'yes') {
            // Apenas "Com Consulta Médica"
            ageSegmentationData.datasets[0].data = filteredIdadeForChart.map(item => item.consultou);
        } else if (diseaseFilter === 'no') {
            // Apenas "Sem Consulta Médica"
            ageSegmentationData.datasets[0].data = filteredIdadeForChart.map(item => item.naoConsultou);
        }
        
        ageGroupStats = filteredIdadeForChart.map(item => {
            const percentage = item.total > 0 ? ((item.consultou / item.total) * 100).toFixed(1) : 0;
            return {
                label: item.categoria,
                total: item.total,
                withConsulta: item.consultou,
                percentage
            };
        });
    } else {
        // Fallback para dados individuais
        const ageGroups = [
            { label: '0 a 4 anos', range: [0, 4] },
            { label: '5 a 19 anos', range: [5, 19] },
            { label: '20 a 39 anos', range: [20, 39] },
            { label: '40 a 49 anos', range: [40, 49] },
            { label: '50 a 64 anos', range: [50, 64] },
            { label: '65 anos ou mais', range: [65, 150] },
            { label: 'Idade ignorada', range: null }
        ];
        
        ageSegmentationData.labels = ageGroups.map(g => g.label);
        
        // Preencher datasets baseado no filtro de consulta médica (fallback)
        if (diseaseFilter === 'all') {
            // Mostrar ambos os datasets
            ageSegmentationData.datasets[0].data = ageGroups.map(group => {
                if (group.range === null) {
                    // Idade ignorada
                    return patients.filter(p => p.age === null && p.hasConsulta).length;
                }
                return patients.filter(p => 
                    p.age !== null &&
                    p.age >= group.range[0] && 
                    p.age <= group.range[1] && 
                    p.hasConsulta
                ).length;
            });
            ageSegmentationData.datasets[1].data = ageGroups.map(group => {
                if (group.range === null) {
                    // Idade ignorada
                    return patients.filter(p => p.age === null && !p.hasConsulta).length;
                }
                return patients.filter(p => 
                    p.age !== null &&
                    p.age >= group.range[0] && 
                    p.age <= group.range[1] && 
                    !p.hasConsulta
                ).length;
            });
        } else if (diseaseFilter === 'yes') {
            // Apenas "Com Consulta Médica"
            ageSegmentationData.datasets[0].data = ageGroups.map(group => {
                if (group.range === null) {
                    return patients.filter(p => p.age === null && p.hasConsulta).length;
                }
                return patients.filter(p => 
                    p.age !== null &&
                    p.age >= group.range[0] && 
                    p.age <= group.range[1] && 
                    p.hasConsulta
                ).length;
            });
        } else if (diseaseFilter === 'no') {
            // Apenas "Sem Consulta Médica"
            ageSegmentationData.datasets[0].data = ageGroups.map(group => {
                if (group.range === null) {
                    return patients.filter(p => p.age === null && !p.hasConsulta).length;
                }
                return patients.filter(p => 
                    p.age !== null &&
                    p.age >= group.range[0] && 
                    p.age <= group.range[1] && 
                    !p.hasConsulta
                ).length;
            });
        }
        
        ageGroupStats = ageGroups.map(group => {
            let groupPatients;
            if (group.range === null) {
                // Idade ignorada
                groupPatients = patients.filter(p => p.age === null);
            } else {
                groupPatients = patients.filter(p => 
                    p.age !== null && p.age >= group.range[0] && p.age <= group.range[1]
                );
            }
            const groupWithConsulta = groupPatients.filter(p => p.hasConsulta).length;
            const groupTotal = groupPatients.length;
            const percentage = groupTotal > 0 ? ((groupWithConsulta / groupTotal) * 100).toFixed(1) : 0;
            return {
                label: group.label,
                total: groupTotal,
                withConsulta: groupWithConsulta,
                percentage
            };
        });
    }
    
    return {
        distributionData,
        riskFactorsData,
        ageSegmentationData,
        domicilioData,
        rendimentoData,
        ageGroupStats,
        stats: {
            total,
            withConsulta,
            withoutConsulta,
            totalNumConsultas,
            menWithConsulta,
            womenWithConsulta,
            menTotal,
            womenTotal,
            urbano,
            rural,
            avgAge
        }
    };
}

// Referências dos gráficos
let distributionChart, riskFactorsChart, ageSegmentationChart, domicilioChart, rendimentoChart;

// Configurações dos gráficos
const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: {
            position: 'bottom',
            labels: {
                padding: 20,
                font: {
                    size: 14
                }
            }
        },
        tooltip: {
            enabled: true,
            callbacks: {
                label: function(context) {
                    if (context.dataset.label) {
                        return context.dataset.label + ': ' + context.parsed.y;
                    }
                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                    const percentage = ((context.parsed / total) * 100).toFixed(1);
                    return context.label + ': ' + context.parsed + ' (' + percentage + '%)';
                }
            }
        },
        onClick: function(evt, elements) {
            if (elements.length > 0) {
                const element = elements[0];
                const chart = this;
                // Permitir cliques para filtrar (implementação futura)
            }
        }
    },
    interaction: {
        intersect: false,
        mode: 'index'
    }
};

// Função para atualizar todos os gráficos
function updateCharts() {
    // SEMPRE usar dados agregados diretamente, não pacientes filtrados
    // Os filtros serão aplicados aos dados agregados, não aos pacientes individuais
    const chartData = calculateChartData([]); // Passar array vazio pois não usamos pacientes
    
    // Formatar números grandes com separador de milhar
    const formatNumber = (num) => {
        if (num === null || num === undefined || isNaN(num)) return '0';
        return Math.round(num).toLocaleString('pt-BR');
    };
    
    // Atualizar estatísticas - usar valores exatos dos dados agregados
    const totalRecordsEl = document.getElementById('totalRecords');
    const totalConsultasEl = document.getElementById('totalConsultas');
    const numConsultasEl = document.getElementById('numConsultas');
    
    if (totalRecordsEl) totalRecordsEl.textContent = formatNumber(chartData.stats.total);
    if (totalConsultasEl) totalConsultasEl.textContent = formatNumber(chartData.stats.withConsulta);
    if (numConsultasEl) numConsultasEl.textContent = formatNumber(chartData.stats.withoutConsulta);
    
    // Atualizar percentuais por sexo
    const malePercentageEl = document.getElementById('malePercentage');
    const femalePercentageEl = document.getElementById('femalePercentage');
    
    if (malePercentageEl && chartData.stats.total > 0) {
        const malePct = ((chartData.stats.menTotal / chartData.stats.total) * 100).toFixed(1);
        malePercentageEl.textContent = malePct + '%';
    }
    if (femalePercentageEl && chartData.stats.total > 0) {
        const femalePct = ((chartData.stats.womenTotal / chartData.stats.total) * 100).toFixed(1);
        femalePercentageEl.textContent = femalePct + '%';
    }
    
    // Atualizar insights
    const insightsElement = document.getElementById('insightsList');
    if (insightsElement) {
        const totalPct = chartData.stats.total > 0 ? ((chartData.stats.withConsulta / chartData.stats.total) * 100).toFixed(1) : 0;
        const urbanoPct = chartData.stats.total > 0 ? ((chartData.stats.urbano / chartData.stats.total) * 100).toFixed(1) : 0;
        const ruralPct = chartData.stats.total > 0 ? ((chartData.stats.rural / chartData.stats.total) * 100).toFixed(1) : 0;
        const avgConsultas = chartData.stats.withConsulta > 0 ? (chartData.stats.totalNumConsultas / chartData.stats.withConsulta).toFixed(1) : 0;
        
        const insightsHTML = `
            <li>${totalPct}% da população realizou consulta médica nos últimos 12 meses</li>
            <li>${chartData.stats.menTotal > 0 ? ((chartData.stats.menWithConsulta / chartData.stats.menTotal) * 100).toFixed(1) : 0}% dos homens realizaram consulta</li>
            <li>${chartData.stats.womenTotal > 0 ? ((chartData.stats.womenWithConsulta / chartData.stats.womenTotal) * 100).toFixed(1) : 0}% das mulheres realizaram consulta</li>
            <li>${urbanoPct}% da população reside em área urbana</li>
            <li>${ruralPct}% da população reside em área rural</li>
            <li>Média de ${avgConsultas} consultas por pessoa que realizou consulta</li>
        `;
        insightsElement.innerHTML = insightsHTML;
    }
    
    // Atualizar estatísticas por grupo de idade
    const ageGroupStatsEl = document.getElementById('ageGroupStats');
    if (ageGroupStatsEl && chartData.ageGroupStats) {
        const ageStatsHTML = chartData.ageGroupStats.map(stat => 
            `<li>Faixa ${stat.label} anos: ${stat.withConsulta} pessoas com consulta médica (${stat.percentage}% do total da faixa)</li>`
        ).join('');
        ageGroupStatsEl.innerHTML = ageStatsHTML || '<li>Nenhum dado disponível</li>';
    }
    
    // Atualizar gráfico de distribuição
    if (distributionChart) {
        distributionChart.data = chartData.distributionData;
        distributionChart.update('active');
    }
    
    // Atualizar gráfico de fatores de risco
    if (riskFactorsChart) {
        riskFactorsChart.data = chartData.riskFactorsData;
        const maxValue = Math.max(...chartData.riskFactorsData.datasets[0].data, 1);
        riskFactorsChart.options.scales.y.max = Math.ceil(maxValue * 1.2);
        riskFactorsChart.update('active');
    }
    
    // Atualizar gráfico de segmentação por idade
    if (ageSegmentationChart) {
        ageSegmentationChart.data = chartData.ageSegmentationData;
        // Calcular valor máximo considerando todos os datasets disponíveis
        let maxAgeValue = 1;
        chartData.ageSegmentationData.datasets.forEach(dataset => {
            const datasetMax = Math.max(...dataset.data, 0);
            if (datasetMax > maxAgeValue) maxAgeValue = datasetMax;
        });
        ageSegmentationChart.options.scales.y.max = Math.ceil(maxAgeValue * 1.2);
        ageSegmentationChart.update('active');
    }
    
    // Atualizar gráfico de situação do domicílio
    if (domicilioChart) {
        domicilioChart.data = chartData.domicilioData;
        domicilioChart.update('active');
    }
    
    // Atualizar gráfico de rendimento
    if (rendimentoChart) {
        rendimentoChart.data = chartData.rendimentoData;
        const maxRendValue = Math.max(...chartData.rendimentoData.datasets[0].data, 1);
        rendimentoChart.options.scales.y.max = Math.ceil(maxRendValue * 1.2);
        rendimentoChart.update('active');
    }
}

// Função para resetar filtros
function resetFilters() {
    document.getElementById('ageFilter').value = 'all';
    document.getElementById('diseaseFilter').value = 'all';
    updateCharts();
}

// Função para inicializar gráficos
function initializeCharts() {
    const initialData = calculateChartData(allPatients);
    
    // Gráfico de distribuição
    distributionChart = new Chart(document.getElementById('distributionChart'), {
        type: 'doughnut',
        data: initialData.distributionData,
        options: {
            ...chartOptions,
            plugins: {
                ...chartOptions.plugins,
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return context.label + ': ' + context.parsed + ' (' + percentage + '%)';
                        }
                    }
                },
                onClick: function(evt, elements) {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const consultaFilter = document.getElementById('diseaseFilter');
                        if (index === 0) {
                            consultaFilter.value = 'no';
                        } else if (index === 1) {
                            consultaFilter.value = 'yes';
                        }
                        updateCharts();
                    }
                }
            }
        }
    });

    // Gráfico de fatores de risco
    riskFactorsChart = new Chart(document.getElementById('riskFactorsChart'), {
        type: 'bar',
        data: initialData.riskFactorsData,
        options: {
            ...chartOptions,
            scales: {
                y: {
                    beginAtZero: true,
                    max: 300,
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                },
                x: {
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                }
            },
            plugins: {
                ...chartOptions.plugins
            }
        }
    });
    
    // Gráfico de segmentação por idade
    ageSegmentationChart = new Chart(document.getElementById('ageSegmentationChart'), {
        type: 'bar',
        data: initialData.ageSegmentationData,
        options: {
            ...chartOptions,
            scales: {
                y: {
                    beginAtZero: true,
                    stacked: false,
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                },
                x: {
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                }
            },
            plugins: {
                ...chartOptions.plugins,
                onClick: function(evt, elements) {
                    if (elements.length > 0) {
                        const element = elements[0];
                        const label = initialData.ageSegmentationData.labels[element.index];
                        const ageFilter = document.getElementById('ageFilter');
                        
                        // Mapear labels da tabela para valores do filtro
                        let filterValue = 'all';
                        if (label.includes('0 a 4')) filterValue = '0-4';
                        else if (label.includes('5 a 19')) filterValue = '5-19';
                        else if (label.includes('20 a 39')) filterValue = '20-39';
                        else if (label.includes('40 a 49')) filterValue = '40-49';
                        else if (label.includes('50 a 64')) filterValue = '50-64';
                        else if (label.includes('65 anos ou mais') || label.includes('65 ou mais')) filterValue = '65+';
                        else if (label.includes('ignorada') || label.includes('Ignorada')) filterValue = 'ignorada';
                        
                        ageFilter.value = filterValue;
                        updateCharts();
                    }
                }
            }
        }
    });
    
    // Gráfico de situação do domicílio
    const domicilioCanvas = document.getElementById('domicilioChart');
    if (domicilioCanvas) {
        domicilioChart = new Chart(domicilioCanvas, {
            type: 'doughnut',
            data: initialData.domicilioData,
            options: {
                ...chartOptions,
                plugins: {
                    ...chartOptions.plugins,
                    onClick: function(evt, elements) {
                        // Removido filtro de domicílio
                    }
                }
            }
        });
    }
    
    // Gráfico de classes de rendimento
    const rendimentoCanvas = document.getElementById('rendimentoChart');
    if (rendimentoCanvas) {
        rendimentoChart = new Chart(rendimentoCanvas, {
            type: 'bar',
            data: initialData.rendimentoData,
            options: {
                ...chartOptions,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            font: {
                                size: 12
                            }
                        }
                    },
                    x: {
                        ticks: {
                            font: {
                                size: 10
                            },
                            maxRotation: 45,
                            minRotation: 45
                        }
                    }
                },
                plugins: {
                    ...chartOptions.plugins,
                    onClick: function(evt, elements) {
                        // Removido filtro de rendimento
                    }
                }
            }
        });
    }
    
    // Adicionar event listeners aos filtros
    document.getElementById('ageFilter').addEventListener('change', updateCharts);
    document.getElementById('diseaseFilter').addEventListener('change', updateCharts);
    
    // Atualizar insights iniciais
    updateCharts();
}

// Carregar dados quando a página carregar
window.onload = function() {
    // Tentar carregar o arquivo Excel automaticamente
    loadExcelFile();
    
    // Configurar upload manual de arquivo
    const fileInput = document.getElementById('excelFileInput');
    const fileNameSpan = document.getElementById('fileName');
    
    if (fileInput) {
        fileInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                fileNameSpan.textContent = `Arquivo selecionado: ${file.name}`;
                handleFileUpload(file);
            }
        });
    }
};

// Adicionar event listener para o botão reset
document.addEventListener('DOMContentLoaded', function() {
    const resetBtn = document.getElementById('resetFiltersBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', resetFilters);
    }

    // Adicionar animações suaves
    const cards = document.querySelectorAll('.card');
    cards.forEach((card, index) => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(20px)';
        setTimeout(() => {
            card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
            card.style.opacity = '1';
            card.style.transform = 'translateY(0)';
        }, index * 200);
    });
});

