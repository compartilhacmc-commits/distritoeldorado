// Configuração da Planilha do Google Sheets
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_GID = '278071504';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SHEET_GID}`;

// Variáveis globais
let allData = [];
let filteredData = [];
let chartUnidades = null;
let chartEspecialidades = null;
let chartPrazos = null;

// Inicialização
document.addEventListener('DOMContentLoaded', function() {
    loadData();
});

// Carregar dados da planilha
async function loadData() {
    showLoading(true);
    try {
        const response = await fetch(SHEET_URL);
        const csvText = await response.text();
        
        // Parse CSV
        const rows = csvText.split('\n').map(row => {
            // Parse CSV considerando campos com vírgulas entre aspas
            const matches = row.match(/(".*?"|[^,]+)(?=\s*,|\s*$)/g);
            return matches ? matches.map(field => field.replace(/^"|"$/g, '').trim()) : [];
        });
        
        // Cabeçalhos
        const headers = rows[0];
        
        // Converter para objetos
        allData = rows.slice(1)
            .filter(row => row.length > 1 && row[0]) // Filtrar linhas vazias
            .map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = row[index] || '';
                });
                return obj;
            });
        
        filteredData = [...allData];
        
        // Inicializar interface
        populateFilters();
        updateDashboard();
        
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert('Erro ao carregar dados da planilha. Verifique a URL e as permissões.');
    } finally {
        showLoading(false);
    }
}

// Mostrar/Ocultar loading
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (show) {
        overlay.classList.add('active');
    } else {
        overlay.classList.remove('active');
    }
}

// Popular filtros
function populateFilters() {
    // Unidades
    const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
    const selectUnidade = document.getElementById('filterUnidade');
    selectUnidade.innerHTML = '<option value="">Todas as Unidades</option>';
    unidades.forEach(unidade => {
        const option = document.createElement('option');
        option.value = unidade;
        option.textContent = unidade;
        selectUnidade.appendChild(option);
    });
    
    // Especialidades
    const especialidades = [...new Set(allData.map(item => item['Cbo Especialidade']))].filter(Boolean).sort();
    const selectEspecialidade = document.getElementById('filterEspecialidade');
    selectEspecialidade.innerHTML = '<option value="">Todas as Especialidades</option>';
    especialidades.forEach(especialidade => {
        const option = document.createElement('option');
        option.value = especialidade;
        option.textContent = especialidade;
        selectEspecialidade.appendChild(option);
    });
    
    // Status
    const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
    const selectStatus = document.getElementById('filterStatus');
    selectStatus.innerHTML = '<option value="">Todos os Status</option>';
    statusList.forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        selectStatus.appendChild(option);
    });
}

// Aplicar filtros
function applyFilters() {
    const unidade = document.getElementById('filterUnidade').value;
    const especialidade = document.getElementById('filterEspecialidade').value;
    const status = document.getElementById('filterStatus').value;
    
    filteredData = allData.filter(item => {
        return (!unidade || item['Unidade Solicitante'] === unidade) &&
               (!especialidade || item['Cbo Especialidade'] === especialidade) &&
               (!status || item['Status'] === status);
    });
    
    updateDashboard();
}

// Limpar filtros
function clearFilters() {
    document.getElementById('filterUnidade').value = '';
    document.getElementById('filterEspecialidade').value = '';
    document.getElementById('filterStatus').value = '';
    
    filteredData = [...allData];
    updateDashboard();
}

// Atualizar dashboard
function updateDashboard() {
    updateCards();
    updateCharts();
    updateTable();
}

// Atualizar cards
function updateCards() {
    const total = allData.length;
    const filtrado = filteredData.length;
    
    // Calcular pendências com 15 e 30 dias
    const hoje = new Date();
    let pendencias15 = 0;
    let pendencias30 = 0;
    
    filteredData.forEach(item => {
        const dataSolicitacao = parseDate(item['Data da Solicitação']);
        if (dataSolicitacao) {
            const diasDecorridos = Math.floor((hoje - dataSolicitacao) / (1000 * 60 * 60 * 24));
            
            if (diasDecorridos >= 15 && diasDecorridos < 30) {
                pendencias15++;
            }
            if (diasDecorridos >= 30) {
                pendencias30++;
            }
        }
    });
    
    // Atualizar valores
    document.getElementById('totalPendencias').textContent = total;
    document.getElementById('percentTotal').textContent = '100%';
    
    document.getElementById('pendencias15').textContent = pendencias15;
    document.getElementById('percent15').textContent = total > 0 ? ((pendencias15 / total) * 100).toFixed(1) + '%' : '0%';
    
    document.getElementById('pendencias30').textContent = pendencias30;
    document.getElementById('percent30').textContent = total > 0 ? ((pendencias30 / total) * 100).toFixed(1) + '%' : '0%';
    
    document.getElementById('totalFiltrados').textContent = filtrado;
    document.getElementById('percentFiltrados').textContent = total > 0 ? ((filtrado / total) * 100).toFixed(1) + '%' : '100%';
}

// Atualizar gráficos
function updateCharts() {
    // Gráfico de Unidades
    const unidadesCount = {};
    filteredData.forEach(item => {
        const unidade = item['Unidade Solicitante'] || 'Não informado';
        unidadesCount[unidade] = (unidadesCount[unidade] || 0) + 1;
    });
    
    const unidadesLabels = Object.keys(unidadesCount).sort((a, b) => unidadesCount[b] - unidadesCount[a]);
    const unidadesValues = unidadesLabels.map(label => unidadesCount[label]);
    
    createBarChart('chartUnidades', unidadesLabels, unidadesValues, '#4A90E2', chartUnidades);
    
    // Gráfico de Especialidades
    const especialidadesCount = {};
    filteredData.forEach(item => {
        const especialidade = item['Cbo Especialidade'] || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });
    
    const especialidadesLabels = Object.keys(especialidadesCount).sort((a, b) => especialidadesCount[b] - especialidadesCount[a]);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);
    
    createBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#28a745', chartEspecialidades);
    
    // Gráfico de Prazos
    const hoje = new Date();
    let menos15 = 0;
    let entre15e30 = 0;
    let mais30 = 0;
    
    filteredData.forEach(item => {
        const dataSolicitacao = parseDate(item['Data da Solicitação']);
        if (dataSolicitacao) {
            const diasDecorridos = Math.floor((hoje - dataSolicitacao) / (1000 * 60 * 60 * 24));
            
            if (diasDecorridos < 15) {
                menos15++;
            } else if (diasDecorridos < 30) {
                entre15e30++;
            } else {
                mais30++;
            }
        }
    });
    
    createPieChart('chartPrazos', 
        ['Menos de 15 dias', '15 a 30 dias', 'Mais de 30 dias'],
        [menos15, entre15e30, mais30],
        chartPrazos
    );
}

// Criar gráfico de barras
function createBarChart(canvasId, labels, data, color, chartInstance) {
    const ctx = document.getElementById(canvasId);
    
    // Destruir gráfico anterior se existir
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    // Criar novo gráfico
    const newChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Quantidade',
                data: data,
                backgroundColor: color,
                borderColor: color,
                borderWidth: 2,
                borderRadius: 8,
                barThickness: 40
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    cornerRadius: 8
                },
                datalabels: {
                    display: true,
                    color: '#FFFFFF',
                    font: {
                        weight: 'bold',
                        size: 14
                    },
                    formatter: (value) => value
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12 },
                        color: '#6c757d'
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                },
                x: {
                    ticks: {
                        font: { size: 11 },
                        color: '#6c757d',
                        maxRotation: 45,
                        minRotation: 45
                    },
                    grid: {
                        display: false
                    }
                }
            }
        },
        plugins: [{
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach(function(dataset, i) {
                    const meta = chart.getDatasetMeta(i);
                    if (!meta.hidden) {
                        meta.data.forEach(function(element, index) {
                            ctx.fillStyle = '#FFFFFF';
                            ctx.font = Chart.helpers.fontString(14, 'bold', 'Arial');
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'middle';
                            const dataString = dataset.data[index].toString();
                            ctx.fillText(dataString, element.x, element.y);
                        });
                    }
                });
            }
        }]
    });
    
    // Atualizar referência
    if (canvasId === 'chartUnidades') {
        chartUnidades = newChart;
    } else if (canvasId === 'chartEspecialidades') {
        chartEspecialidades = newChart;
    }
}

// Criar gráfico de pizza
function createPieChart(canvasId, labels, data, chartInstance) {
    const ctx = document.getElementById(canvasId);
    
    // Destruir gráfico anterior se existir
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    // Criar novo gráfico
    chartPrazos = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: ['#28a745', '#ffc107', '#dc3545'],
                borderColor: '#FFFFFF',
                borderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 20,
                        font: { size: 13, weight: '600' },
                        color: '#212529'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${label}: ${value} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// Atualizar tabela
function updateTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="loading-cell"><i class="fas fa-inbox"></i> Nenhum registro encontrado</td></tr>';
        return;
    }
    
    filteredData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item['N° Solicitação'] || '-'}</td>
            <td>${formatDate(item['Data da Solicitação'])}</td>
            <td>${item['Nº Prontuário'] || '-'}</td>
            <td>${item['Telefone'] || '-'}</td>
            <td>${item['Unidade Solicitante'] || '-'}</td>
            <td>${item['Cbo Especialidade'] || '-'}</td>
            <td>${formatDate(item['Data Início da Pendência'])}</td>
            <td>${item['Status'] || '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

// Funções auxiliares
function parseDate(dateString) {
    if (!dateString) return null;
    
    // Tentar vários formatos de data
    const formats = [
        /(\d{2})\/(\d{2})\/(\d{4})/, // DD/MM/YYYY
        /(\d{4})-(\d{2})-(\d{2})/, // YYYY-MM-DD
    ];
    
    for (let format of formats) {
        const match = dateString.match(format);
        if (match) {
            if (format === formats[0]) {
                // DD/MM/YYYY
                return new Date(match[3], match[2] - 1, match[1]);
            } else {
                // YYYY-MM-DD
                return new Date(match[1], match[2] - 1, match[3]);
            }
        }
    }
    
    return null;
}

function formatDate(dateString) {
    if (!dateString) return '-';
    const date = parseDate(dateString);
    if (!date || isNaN(date.getTime())) return dateString;
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
}

// Atualizar página
function refreshData() {
    loadData();
}

// Download Excel
function downloadExcel() {
    if (filteredData.length === 0) {
        alert('Não há dados para exportar.');
        return;
    }
    
    // Preparar dados para exportação
    const exportData = filteredData.map(item => ({
        'Nº Solicitação': item['N° Solicitação'] || '',
        'Data Solicitação': item['Data da Solicitação'] || '',
        'Nº Prontuário': item['Nº Prontuário'] || '',
        'Telefone': item['Telefone'] || '',
        'Unidade Solicitante': item['Unidade Solicitante'] || '',
        'CBO Especialidade': item['Cbo Especialidade'] || '',
        'Data Início Pendência': item['Data Início da Pendência'] || '',
        'Status': item['Status'] || ''
    }));
    
    // Criar workbook
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    
    // Definir largura das colunas
    const colWidths = [
        { wch: 15 }, // Nº Solicitação
        { wch: 15 }, // Data Solicitação
        { wch: 15 }, // Nº Prontuário
        { wch: 15 }, // Telefone
        { wch: 30 }, // Unidade Solicitante
        { wch: 30 }, // CBO Especialidade
        { wch: 18 }, // Data Início Pendência
        { wch: 20 }  // Status
    ];
    ws['!cols'] = colWidths;
    
    // Baixar arquivo
    const hoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Pendencias_Vivver_${hoje}.xlsx`);
}