// ==========================================================
// 1. VARIÁVEIS GLOBAIS E CONSTANTES
// ==========================================================
let isAdmin = false;
const SENHA_CORRETA = "123";

let activeChartFilter = null;
let activeChartFilterType = null;
let lastDisplayData = [];

// Base de Dados
let rawData = []; 
let ntpRawData = []; 
let dbRawData = []; 
let dbColumns = [];

// Gráficos e Filtros Globais
let charts = {}; 
let currentYear = 'Todos'; 
let slaTimeView = 'monthly'; 
let ntpTimeView = 'monthly'; 
let ntpFilterState = {}; 
let slaFilterState = {}; 
let dbFilterState = {}; 
let ordersFilterState = {};

// Histórico (Ctrl+Z)
let historyStack = []; 
const MAX_HISTORY = 5;

// Tooltip Customizado para Gráficos
const customTooltipConfig = {
    backgroundColor: '#1e293b', titleColor: '#f8fafc', bodyColor: '#cbd5e1',
    borderColor: '#334155', borderWidth: 1, padding: 12, cornerRadius: 8,
    callbacks: {
        label: function(context) {
            let label = context.dataset.label || '';
            let val = context.raw || 0;
            if (label.includes('SLA')) {
                let abs = context.chart.data.datasets[0].extraAbs ? context.chart.data.datasets[0].extraAbs[context.dataIndex] : '-';
                let fec = context.chart.data.datasets[0].extraFec ? context.chart.data.datasets[0].extraFec[context.dataIndex] : '-';
                return [`${label}: ${val.toFixed(1)}%`, `Aberturas: ${abs}`, `Fechamentos: ${fec}`];
            }
            return `${label}: ${val}`;
        }
    }
};

// ==========================================================
// FUNÇÃO DE TMA (Conversão BR)
// ==========================================================
function converterDataBR(dataStr) {
    if (!dataStr) return null;
    let str = String(dataStr).trim();
    if(str === "") return null;
    let partes = str.split(/[\s/:]+/);
    if (partes.length >= 3) {
        let dia = parseInt(partes[0], 10);
        let mes = parseInt(partes[1], 10) - 1; 
        let ano = parseInt(partes[2], 10);
        if(ano > 31) { 
            if (ano < 100) ano += 2000;
            let hora = partes[3] ? parseInt(partes[3], 10) : 0;
            let min = partes[4] ? parseInt(partes[4], 10) : 0;
            let seg = partes[5] ? parseInt(partes[5], 10) : 0;
            return new Date(ano, mes, dia, hora, min, seg);
        }
    }
    let fallback = new Date(str);
    return isNaN(fallback.getTime()) ? null : fallback;
}

// ==========================================================
// 2. INICIALIZAÇÃO DA APLICAÇÃO (ON LOAD)
// ==========================================================
document.addEventListener('DOMContentLoaded', () => {
    if(window.lucide) lucide.createIcons();

    const storedData = localStorage.getItem('sla_dashboard_data_v10'); 
    if (storedData) try { rawData = JSON.parse(storedData); } catch(e){ rawData = []; }
    
    const storedNtp = localStorage.getItem('ntp_dashboard_data_v10'); 
    if (storedNtp) try { ntpRawData = JSON.parse(storedNtp); } catch(e){ ntpRawData = []; }
    
    const storedDb = localStorage.getItem('db_dashboard_data_v10'); 
    if (storedDb) { try { const parsed = JSON.parse(storedDb); dbRawData = parsed.rows || []; dbColumns = parsed.cols || []; } catch (e) { dbRawData = []; dbColumns = []; } }
    
    const yearSelect = document.getElementById('globalYearFilter'); 
    if (yearSelect) { currentYear = (yearSelect.value === 'Todos') ? 'Todos' : parseInt(yearSelect.value); }
    
    init();
    
    const slaSearch = document.getElementById('slaTableSearch'); 
    if(slaSearch) slaSearch.addEventListener('keyup', debounce(() => filterSlaTable(lastDisplayData, true), 300));
    
    const ntpSearch = document.getElementById('ntpSearch'); 
    if(ntpSearch) ntpSearch.addEventListener('keyup', debounce(() => filterNtpTable(), 300));
    
    window.addEventListener('scroll', (e) => {
        if (e.target.closest('.overflow-x-auto') || e.target.closest('.db-select') || e.target.closest('.custom-scrollbar')) return; 
        if (typeof closeSlaFilters === 'function') closeSlaFilters();
        if (typeof closeNtpFilters === 'function') closeNtpFilters();
        if (typeof closeDbFilters === 'function') closeDbFilters();
        if (typeof closeOrdersFilters === 'function') closeOrdersFilters();
    }, true);
    
    document.addEventListener('keydown', function(event) { if ((event.ctrlKey || event.metaKey) && event.key === 'z') { event.preventDefault(); undoLastAction(); } });
    
    const calcInput = document.getElementById('fileUploadCalc'); 
    if(calcInput) calcInput.addEventListener('change', handleCalculatorFile);

    const dateInputProd = document.getElementById('dataBaseCalculo');
    if(dateInputProd) dateInputProd.value = new Date().toISOString().split('T')[0];
});

function init() { 
    populateUnidadeFilter(); 
    updateDashboard(); 
    if(ntpRawData.length > 0) processNTPData(); 
    if(dbRawData.length > 0) renderDatabase(); 
    updateUndoButton(); 
}

function debounce(func, wait) {
    let timeout;
    return function(...args) { clearTimeout(timeout); timeout = setTimeout(() => func.apply(this, args), wait); };
}

// ==========================================================
// 3. CONTROLE DE AUTENTICAÇÃO E PERMISSÕES
// ==========================================================
function toggleUserMenu(event) {
    if (event) event.stopPropagation();
    const dropdown = document.getElementById('userDropdown');
    dropdown.classList.toggle('hidden');
    if(window.lucide) lucide.createIcons(); 
}

document.addEventListener('click', function(event) {
    const dropdown = document.getElementById('userDropdown');
    const btn = document.getElementById('userAvatarBtn');
    if (!btn.contains(event.target) && !dropdown.contains(event.target)) {
        dropdown.classList.add('hidden');
    }
});

function openLoginModal() {
    document.getElementById('userDropdown').classList.add('hidden');
    document.getElementById('loginModal').classList.remove('hidden');
    document.getElementById('adminPasswordInput').value = '';
    document.getElementById('loginErrorMsg').classList.add('hidden');
    document.getElementById('adminPasswordInput').focus();
}

function closeLoginModal() { document.getElementById('loginModal').classList.add('hidden'); }

function checkLogin() {
    const input = document.getElementById('adminPasswordInput').value;
    if (input === SENHA_CORRETA) {
        isAdmin = true; updateInterfaceForAdmin(); closeLoginModal();
    } else {
        document.getElementById('loginErrorMsg').classList.remove('hidden');
        const inputField = document.getElementById('adminPasswordInput');
        inputField.classList.add('border-red-500', 'ring-1', 'ring-red-500');
        setTimeout(() => { inputField.classList.remove('border-red-500', 'ring-1', 'ring-red-500'); }, 500);
    }
}

function performLogout() {
    isAdmin = false;
    document.querySelectorAll('.admin-only').forEach(el => el.classList.add('hidden'));
    document.getElementById('userStatusLabel').innerText = "Visitante";
    document.getElementById('userStatusLabel').classList.remove('text-blue-600');
    document.getElementById('btnLoginOption').classList.remove('hidden');
    document.getElementById('btnLogoutOption').classList.add('hidden');
    const btn = document.getElementById('userAvatarBtn');
    btn.innerHTML = '<i data-lucide="user" class="w-5 h-5 text-gray-600"></i>';
    btn.classList.remove('ring-2', 'ring-blue-500', 'ring-offset-2');
    document.getElementById('userDropdown').classList.add('hidden');
    document.getElementById('adminToolsMenu').classList.add('hidden');
    if(window.lucide) lucide.createIcons();
}

function updateInterfaceForAdmin() {
    document.getElementById('userStatusLabel').innerText = "Administrador";
    document.getElementById('userStatusLabel').classList.add('text-blue-600');
    document.getElementById('btnLoginOption').classList.add('hidden');
    document.getElementById('btnLogoutOption').classList.remove('hidden');
    const btn = document.getElementById('userAvatarBtn');
    btn.innerHTML = '<i data-lucide="user-check" class="w-5 h-5 text-blue-600"></i>';
    btn.classList.add('ring-2', 'ring-blue-500', 'ring-offset-2');
    document.querySelectorAll('.admin-only').forEach(el => el.classList.remove('hidden'));
    if(window.lucide) lucide.createIcons();
}

function checkPermission() {
    if (!isAdmin) { alert("Ação permitida apenas para administradores."); return false; }
    return true;
}

// ==========================================================
// 4. PREFERÊNCIAS GERAIS E NAVEGAÇÃO ENTRE ABAS
// ==========================================================
function toggleDarkMode() {
    document.body.classList.toggle('dark');
    const icon = document.getElementById('darkModeIcon');
    if (document.body.classList.contains('dark')) icon.setAttribute('data-lucide', 'sun');
    else icon.setAttribute('data-lucide', 'moon');
    lucide.createIcons();
    if(typeof updateDashboard === 'function') updateDashboard();
    if(typeof processNTPData === 'function') processNTPData();
}

function switchTab(tab) {
    const slaBtn = document.getElementById('btn-sla'); const ntpBtn = document.getElementById('btn-ntp');
    if (['calc', 'repair', 'db', 'prod'].includes(tab)) { 
        slaBtn.classList.remove('active'); slaBtn.classList.add('inactive'); ntpBtn.classList.remove('active'); ntpBtn.classList.add('inactive'); 
    } else { 
        document.querySelectorAll('.nav-btn').forEach(btn => { btn.classList.remove('active', 'inactive'); btn.classList.add('inactive'); }); 
        const btn = document.getElementById('btn-' + tab); 
        if(btn) { btn.classList.remove('inactive'); btn.classList.add('active'); } 
    }
    
    ['sla','ntp','db','calc', 'repair', 'prod'].forEach(v => { document.getElementById('view-' + v)?.classList.add('hidden'); }); 
    document.getElementById('view-' + tab)?.classList.remove('hidden');
    
    if(tab === 'ntp') processNTPData(); 
    if(tab === 'db') renderDatabase(); 
    lucide.createIcons();
}

function toggleAllDetails() {
    const details = document.querySelectorAll('.js-details');
    if(details.length === 0) return;
    const isHidden = details[0].classList.contains('hidden');

    details.forEach(el => {
        if (isHidden) { el.classList.remove('hidden'); el.classList.add('flex'); } 
        else { el.classList.add('hidden'); el.classList.remove('flex'); }
    });
}

function toggleAllChartsSize() {
    const ids = ['slaChart', 'volumeChart'];
    const firstCard = document.getElementById('card-slaChart');
    if(!firstCard) return;
    const isExpanding = firstCard.classList.contains('lg:col-span-1');

    ids.forEach(id => {
        const card = document.getElementById('card-' + id);
        if (!card) return;
        if (isExpanding) {
            card.classList.remove('lg:col-span-1');
            card.classList.add('lg:col-span-2');
        } else {
            card.classList.remove('lg:col-span-2');
            card.classList.add('lg:col-span-1');
        }
        setTimeout(() => {
            if (charts[id]) charts[id].resize();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 350);
    });
}

// ==========================================================
// LÓGICA DO HEATMAP
// ==========================================================
function toggleHeatmap(view) {
    const volContainer = document.getElementById('volChartContainer');
    const heatContainer = document.getElementById('heatmapContainer');
    const btnBar = document.getElementById('btnViewBar');
    const btnHeat = document.getElementById('btnViewHeat');
    const title = document.getElementById('volChartTitle');

    if(view === 'heat') {
        volContainer.classList.add('hidden'); heatContainer.classList.remove('hidden');
        btnBar.classList.replace('bg-white', 'transparent'); btnBar.classList.replace('shadow-sm', 'shadow-none'); btnBar.classList.replace('text-blue-600', 'text-slate-400');
        btnHeat.classList.replace('text-slate-400', 'text-orange-500'); btnHeat.classList.add('bg-white', 'shadow-sm');
        title.innerText = "Mapa de Calor (Pico de Aberturas)";
        renderHeatmap();
    } else {
        heatContainer.classList.add('hidden'); volContainer.classList.remove('hidden');
        btnHeat.classList.replace('bg-white', 'transparent'); btnHeat.classList.replace('shadow-sm', 'shadow-none'); btnHeat.classList.replace('text-orange-500', 'text-slate-400');
        btnBar.classList.replace('text-slate-400', 'text-blue-600'); btnBar.classList.add('bg-white', 'shadow-sm');
        title.innerText = "Comparativo Aberturas x Fechamentos";
    }
}

function renderHeatmap() {
    const grid = document.getElementById('heatmapGrid');
    grid.innerHTML = '';
    
    const hours = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
    const hourCounts = new Array(11).fill(0);
    
    lastDisplayData.forEach(obj => {
        if(obj.ordersListFiltered) {
            obj.ordersListFiltered.forEach(o => {
                if(o.data && o.data.includes(' ')) {
                    const timePart = o.data.split(' ')[1];
                    const hour = parseInt(timePart.split(':')[0]);
                    const idx = hours.indexOf(hour);
                    if(idx !== -1) hourCounts[idx]++;
                }
            });
        }
    });

    const maxCount = Math.max(...hourCounts) || 1;

    hourCounts.forEach((count, idx) => {
        const intensity = count / maxCount;
        let colorClass = 'bg-slate-100'; 
        if(intensity > 0) colorClass = 'bg-orange-200';
        if(intensity > 0.4) colorClass = 'bg-orange-400';
        if(intensity > 0.7) colorClass = 'bg-orange-600';
        
        grid.innerHTML += `
            <div class="flex flex-col justify-end h-full group relative">
                <div class="w-full rounded-t-sm transition-all duration-500 ${colorClass}" style="height: ${Math.max((intensity * 100), 10)}%;"></div>
                <div class="absolute bottom-full mb-2 left-1/2 -translate-x-1/2 opacity-0 group-hover:opacity-100 transition-opacity bg-slate-800 text-white text-[10px] px-2 py-1 rounded pointer-events-none whitespace-nowrap z-20">
                    ${hours[idx]}h: ${count} chamados
                </div>
            </div>
        `;
    });
}

// ==========================================================
// 5. HISTÓRICO, BACKUP E RESET
// ==========================================================
function saveStateToHistory() {
    if (historyStack.length >= MAX_HISTORY) historyStack.shift(); 
    historyStack.push({ rawData: JSON.parse(JSON.stringify(rawData)), ntpRawData: JSON.parse(JSON.stringify(ntpRawData)), dbRawData: JSON.parse(JSON.stringify(dbRawData)), dbColumns: JSON.parse(JSON.stringify(dbColumns)) });
    updateUndoButton();
}

function undoLastAction() {
    if (historyStack.length === 0) return;
    const btn = document.getElementById('btnUndoMinimal'); btn.style.transform = "scale(0.9)"; setTimeout(()=>btn.style.transform = "scale(1)", 150);
    if(!confirm("Desfazer a última ação?")) return;
    const lastState = historyStack.pop();
    rawData = lastState.rawData; ntpRawData = lastState.ntpRawData; dbRawData = lastState.dbRawData; dbColumns = lastState.dbColumns;
    saveAllData(); init(); updateUndoButton();
}

function fullReset() {
    if(!checkPermission()) return;
    if(!confirm("TEM CERTEZA? Isso apagará todos os dados importados e gráficos.")) return;
    localStorage.removeItem('sla_dashboard_data_v10'); localStorage.removeItem('ntp_dashboard_data_v10'); localStorage.removeItem('db_dashboard_data_v10');
    rawData = []; ntpRawData = []; dbRawData = []; dbColumns = []; historyStack = [];
    init(); clearChartDateFilter(); updateUndoButton();
    alert("Dashboard limpo com sucesso!");
}

function updateUndoButton() { 
    const btn = document.getElementById('btnUndoMinimal'); 
    if(historyStack.length > 0) btn.classList.remove('hidden'); else btn.classList.add('hidden'); 
}

function saveAllData() { 
    try { 
        localStorage.setItem('sla_dashboard_data_v10', JSON.stringify(rawData)); 
        localStorage.setItem('ntp_dashboard_data_v10', JSON.stringify(ntpRawData)); 
        localStorage.setItem('db_dashboard_data_v10', JSON.stringify({ cols: dbColumns, rows: dbRawData.slice(0, 300) })); 
    } catch(e) { console.warn("Memória local cheia."); } 
}

// ==========================================================
// 6. LÓGICA DE DATAS E STATUS (FERIADOS, PRAZOS E STATUS)
// ==========================================================
const feriadosSP = new Set([
    "2025-01-01", "2025-01-25", "2025-03-03", "2025-03-04", "2025-03-05", "2025-04-18", "2025-04-21", "2025-05-01", "2025-06-19", "2025-07-09", "2025-09-07", "2025-10-12", "2025-11-02", "2025-11-15", "2025-11-20", "2025-12-25",
    "2026-01-01", "2026-01-25", "2026-02-02", "2026-02-16", "2026-02-17", "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04", "2026-07-09", "2026-09-07", "2026-10-12", "2026-11-02", "2026-11-15", "2026-11-20", "2026-12-25"
]);

function addBusinessDays(startDate, daysToAdd) {
    let currentDate = new Date(startDate);
    let addedDays = 0; let safeGuard = 0; 
    while (addedDays < daysToAdd && safeGuard < 200) {
        currentDate.setDate(currentDate.getDate() + 1);
        const dayOfWeek = currentDate.getDay(); 
        const isoDate = currentDate.toISOString().split('T')[0];
        if (dayOfWeek !== 0 && dayOfWeek !== 6 && !feriadosSP.has(isoDate)) addedDays++;
        safeGuard++;
    }
    return currentDate;
}

function calcularVencimento(dataAberturaIso, sla, estadoIgnorado) {
    if (!dataAberturaIso) return null;
    const parts = dataAberturaIso.split('-');
    const startDate = new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
    if (isNaN(startDate.getTime())) return null;

    let daysToAdd = 5; 
    const slaUpper = sla ? sla.toUpperCase().trim() : "BAIXO";

    if (slaUpper === "MUITO ELEVADO") daysToAdd = 1;
    else if (slaUpper === "ELEVADO") daysToAdd = 2;
    else if (slaUpper === "MEDIO" || slaUpper === "MÉDIO") daysToAdd = 3;

    const dueDate = addBusinessDays(startDate, daysToAdd);
    return dueDate.toISOString().split('T')[0];
}

function calculateSmartStatus(originalStatus, vencimentoIso, dataFechamentoIso = null, statusPrazoCsv = null) { 
    const statusUpper = (originalStatus || '').toUpperCase(); 
    if (dataFechamentoIso || statusUpper.includes('FINAL') || statusUpper.includes('CONCLU') || statusUpper.includes('ENCERRAD') || statusUpper.includes('FECHAD') || statusUpper.includes('ATENDID')) { 
        return { text: 'Finalizado', class: 'text-emerald-700 font-bold bg-emerald-100 px-2 py-0.5 rounded border border-emerald-200 uppercase tracking-wide text-[9px]' }; 
    } 
    if (!vencimentoIso) return { text: 'S/ Data', class: 'text-slate-500 italic bg-slate-100 px-2 py-0.5 rounded uppercase tracking-wide text-[9px]' }; 
    
    const today = new Date(); today.setHours(0,0,0,0); 
    const vParts = vencimentoIso.split('-');
    const vencDate = new Date(parseInt(vParts[0]), parseInt(vParts[1])-1, parseInt(vParts[2]));
    vencDate.setHours(0,0,0,0);
    
    if (vencDate < today) return { text: 'Vencido', class: 'text-rose-700 font-black bg-rose-100 px-2 py-0.5 rounded border border-rose-200 uppercase tracking-wide text-[9px]' }; 
    else if (vencDate.getTime() === today.getTime()) return { text: 'Vence Hoje', class: 'text-orange-700 font-black bg-orange-100 px-2 py-0.5 rounded border border-orange-200 uppercase tracking-wide text-[9px]' }; 
    else { 
        const amanha = new Date(today); amanha.setDate(today.getDate()+1);
        if(vencDate.getTime() === amanha.getTime()) return { text: 'Vence Amanhã', class: 'text-yellow-700 font-bold bg-yellow-100 px-2 py-0.5 rounded border border-yellow-200 uppercase tracking-wide text-[9px]' };
        return { text: 'No Prazo', class: 'text-blue-700 font-bold bg-blue-100 px-2 py-0.5 rounded border border-blue-200 uppercase tracking-wide text-[9px]' }; 
    } 
}

function parseDateToIso(dateStr) { 
    if (!dateStr || typeof dateStr !== 'string') return null; 
    let cleanDate = dateStr.trim(); if (cleanDate.includes(' ')) cleanDate = cleanDate.split(' ')[0]; 
    const br = /^([0-9]{1,2})[\/\-]([0-9]{1,2})[\/\-]([0-9]{2,4})/.exec(cleanDate); 
    if (br) { let y = br[3]; let m = br[2]; let d = br[1]; if (y.length === 2) y = '20' + y; return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`; } 
    const iso = /^([0-9]{4})[-\/]([0-9]{1,2})[-\/]([0-9]{1,2})/.exec(cleanDate); 
    if (iso) return `${iso[1]}-${iso[2].padStart(2,'0')}-${iso[3].padStart(2,'0')}`; 
    return null; 
}

function getMonthNameFromIso(isoDate) { 
    if (!isoDate) return null; 
    const months = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO']; 
    const parts = isoDate.split('-'); 
    if (parts.length > 1) { const mIdx = parseInt(parts[1], 10) - 1; if (mIdx >= 0 && mIdx <= 11) return months[mIdx]; } 
    return null; 
}

// ==========================================================
// 7. IMPORTAÇÃO E PROCESSAMENTO DE DADOS PRINCIPAIS (CSV)
// ==========================================================
function toggleUploadMenu(event) { event.stopPropagation(); document.getElementById('slaUploadMenu').classList.toggle('show'); document.addEventListener('click', closeUploadMenu); }
function closeUploadMenu() { document.getElementById('slaUploadMenu')?.classList.remove('show'); document.removeEventListener('click', closeUploadMenu); }

function showLoading() { document.getElementById('loadingOverlay').style.display = 'flex'; }
function hideLoading() { document.getElementById('loadingOverlay').style.display = 'none'; }

function downloadSlaTemplate() { 
    const headers = ["Unidade", "Data Abertura", "AB", "FC", "Status", "UF", "OS", "SLA DE ATENDIMENTO"]; 
    const exampleRow1 = ["CD São Paulo", "01/01/2026", "01/01/2026", "", "Aberto", "SP", "505500", "MEDIO"]; 
    const csvContent = "data:text/csv;charset=utf-8," + headers.join(";") + "\n" + exampleRow1.join(";"); 
    const encodedUri = encodeURI(csvContent); 
    const link = document.createElement("a"); link.setAttribute("href", encodedUri); link.setAttribute("download", "modelo.csv"); 
    document.body.appendChild(link); link.click(); document.body.removeChild(link); closeUploadMenu(); 
}

function findHeader(headers, patterns) { 
    if(!headers) return null; 
    for (const p of patterns) { const found = headers.find(h => h.trim().toUpperCase() === p.toUpperCase()); if (found) return found; } 
    for (const p of patterns) { const rx = new RegExp(p, 'i'); const found = headers.find(h => rx.test(h)); if (found) return found; } 
    return null; 
}

function cleanUnitName(name) { if(!name) return "Indefinido"; return name.replace(/^[0-9]+\s*[-–]\s*/, '').trim(); }

async function handleFileSelect(event, type) {
    if (!checkPermission()) { event.target.value = ''; return; }
    const files = event.target.files; if (!files || files.length === 0) return;
    
    closeUploadMenu(); showLoading(); saveStateToHistory(); 
    let processedCount = 0; let totalFiles = files.length; let combinedHeaders = null;

    for (let i = 0; i < totalFiles; i++) {
        const file = files[i];
        await new Promise((resolve) => { 
            Papa.parse(file, { 
                header: true, skipEmptyLines: true, worker: false, encoding: "ISO-8859-1", 
                complete: function(results) { 
                    if(results.data && results.data.length > 0) { 
                        if (!combinedHeaders) combinedHeaders = results.meta.fields; 
                        if (type === 'ntp') parseNTPCSV(results.data, results.meta.fields, false); 
                        else if (type === 'sla_abertura') mergeSlaData(results.data, results.meta.fields, 'abertura', false); 
                        else if (type === 'sla_fechamento') mergeSlaData(results.data, results.meta.fields, 'fechamento', false); 
                    } 
                    resolve(); 
                }, 
                error: function(err) { console.error(err); resolve(); } 
            }); 
        });
        processedCount++; document.getElementById('loadingText').innerText = `Processando arquivo ${processedCount} de ${totalFiles}...`;
    }

    try { 
        if (type === 'ntp') { 
            localStorage.setItem('ntp_dashboard_data_v10', JSON.stringify(ntpRawData)); 
            if(ntpRawData.length > 0 && (!dbColumns || dbColumns.length === 0)) dbColumns = combinedHeaders; 
            localStorage.setItem('db_dashboard_data_v10', JSON.stringify({ cols: dbColumns, rows: dbRawData.slice(0, 300) })); 
        } else { 
            localStorage.setItem('sla_dashboard_data_v10', JSON.stringify(rawData)); 
        } 
    } catch (e) { alert("Memória Local cheia!"); }
    
    if (type === 'ntp') { processNTPData(); renderDatabase(); } else { init(); }
    hideLoading(); event.target.value = ''; alert(`Processamento concluído!`);
}

function mergeSlaData(rows, headers, forcedType, autoSave = true) {
    if (!rows || !rows.length) return;
    
    const colAb = findHeader(headers, ['AB', 'Data Abertura', 'date', 'emissao']);
    const colFc = findHeader(headers, ['FC', 'Fechamento', 'Conclusao', 'Data Fim']);
    let dataTargetHdr = forcedType === 'abertura' ? colAb : colFc;

    const unidadeHdr = headers[0]; const statusHdr = findHeader(headers, ['status', 'situacao', 'fase', 'estado']); const osHdr = findHeader(headers, ['os', 'pedido', 'chamado', 'order', 'protocolo']); 
    const colClassificacaoSla = findHeader(headers, ['SLA DE ATENDIMENTO', 'Classificação', 'Classificacao', 'Prioridade', 'SLA', 'Nível']);
    const estadoHdr = findHeader(headers, ['UF', 'Estado', 'ESTADO']); 
    const colLancPeca = findHeader(headers, ['DATA LANÇAMENTO DA PEÇA', 'DATA LANCAMENTO DA PECA', 'DATA LANÇAMENTO DA PECA']);
    const colSituacao = findHeader(headers, ['Situação', 'Situacao', 'SITUACAO', 'SITUAÇÃO']);
    const colSolucao = findHeader(headers, ['Solução', 'Solucao', 'SOLUCAO', 'SOLUÇÃO', 'AM', 'Solução Final', 'Solucao Final']);
    const colVencimentoSla = findHeader(headers, ['VENCIMENTO OS', 'Vencimento', 'Data Vencimento', 'Prazo limite', 'Prazo']);
    
    const dataMap = {}; rawData.forEach(item => { dataMap[item.idKey] = item; }); 
    
    rows.forEach(r => {
        let unidade = unidadeHdr ? (r[unidadeHdr] || '') : ''; unidade = unidade ? cleanUnitName(unidade) : 'Indefinido';
        let dataOriginal = dataTargetHdr && r[dataTargetHdr] ? r[dataTargetHdr] : '';
        let dtIso = parseDateToIso(dataOriginal); let rowYear = currentYear !== 'Todos' ? currentYear : 2026; let mes = "INDEFINIDO"; 
        if (dtIso) { rowYear = parseInt(dtIso.split('-')[0]); const mName = getMonthNameFromIso(dtIso); if (mName) mes = mName; }
        let estado = estadoHdr && r[estadoHdr] ? r[estadoHdr] : "SP";
        const idKey = `${unidade}-${mes}-${rowYear}`;
        
        if (!dataMap[idKey]) dataMap[idKey] = { idKey, unidade, estado: estado, mes, ano: rowYear, aberturas: 0, fechamentos: 0, ntps: 0, reincidencias: 0, trocas: 0, ordersList: [], fechamentosList: [] };
        
        if (forcedType === 'fechamento') { 
            if(!dataMap[idKey].fechamentosList) dataMap[idKey].fechamentosList = []; 
            dataMap[idKey].fechamentosList.push({ data: dataOriginal, dateIso: dtIso }); 
            dataMap[idKey].fechamentos++; 
        } else if (forcedType === 'abertura') { 
            const st = statusHdr ? (r[statusHdr] || 'Indefinido') : 'Indefinido'; 
            const os = osHdr ? (r[osHdr] || 'N/A') : 'N/A';
            let classificacao = colClassificacaoSla ? (r[colClassificacaoSla] || 'BAIXO') : 'BAIXO'; 
            let finalVencIso = null; let finalVencStr = "";
            
            if (colVencimentoSla && r[colVencimentoSla]) {
                let vStr = String(r[colVencimentoSla]).trim();
                finalVencIso = parseDateToIso(vStr);
                if (finalVencIso) finalVencStr = finalVencIso.split('-').reverse().join('/');
            }

            // Captura o fechamento se estiver na mesma linha de abertura
            let dataFcOriginal = colFc && r[colFc] ? r[colFc] : '';
            let fcIso = parseDateToIso(dataFcOriginal);

            let isNtp = false;
            if (colLancPeca && r[colLancPeca] && String(r[colLancPeca]).trim() !== '') { dataMap[idKey].ntps++; isNtp = true; }
            
            let isReinc = false;
            if (colSituacao && r[colSituacao]) {
                const valSit = String(r[colSituacao]).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
                if (valSit.includes('reincidencia')) { dataMap[idKey].reincidencias++; isReinc = true; }
            }
            
            let isTroca = false;
            if (colSolucao && r[colSolucao]) {
                const valSolOriginal = String(r[colSolucao]).toLowerCase();
                if (valSolOriginal.includes('solicitar troca de equipamento') || valSolOriginal.includes('troca técnica') || valSolOriginal.includes('troca tecnica')) { dataMap[idKey].trocas++; isTroca = true; }
            }

            if(!dataMap[idKey].ordersList) dataMap[idKey].ordersList = []; 
            dataMap[idKey].ordersList.push({ 
                os, statusOriginal: st, data: dataOriginal, dateIso: dtIso, 
                vencimento: finalVencStr, vencimentoIso: finalVencIso, 
                dataFechamento: dataFcOriginal, dataFechamentoIso: fcIso,
                classificacaoSla: classificacao, isNtp: isNtp, isReinc: isReinc, isTroca: isTroca, rawRow: r 
            }); 
            dataMap[idKey].aberturas++; 
        }
    });
    rawData = Object.values(dataMap); if(autoSave) { try { localStorage.setItem('sla_dashboard_data_v10', JSON.stringify(rawData)); init(); } catch(e) {} }
}

function parseNTPCSV(rows, headers, autoSave = true) {
    if (!rows || !rows.length) return;
    const colDataOriginal = findHeader(headers, ['Data Abertura', 'Abertura', 'Criado em', 'Data']); 
    const colDataLancPeca = findHeader(headers, ['DATA LANÇAMENTO DA PEÇA', 'DATA LANCAMENTO DA PECA', 'DATA LANÇAMENTO DA PECA']);
    const colFinalizado = findHeader(headers, ['Data Fechamento', 'Conclusão', 'Encerrado', 'Fim']); 
    const colStatus = findHeader(headers, ['Status', 'Situacao', 'Status do Pedido']); 
    const colId = findHeader(headers, ['Chamado', 'ID', 'OS', 'Ticket']); 
    const colUnidade = findHeader(headers, ['Unidade', 'Cliente', 'Loja']); 
    const colTecnico = findHeader(headers, ['Técnico', 'Atendente', 'Responsável']); 
    const colPosto = findHeader(headers, ['Posto', 'Nome Posto', 'Autorizada', 'Prestador']); 
    const colVencimento = findHeader(headers, ['VENCIMENTO OS', 'Vencimento', 'Data Vencimento', 'Prazo limite', 'Prazo', 'Data Vencimento Calculada']);
    
    if (!dbColumns.includes('Status Prazo Calc')) dbColumns.push('Status Prazo Calc');
    let data = ntpRawData; 
    
    rows.forEach(row => {
        const rowStr = JSON.stringify(row).toUpperCase(); const statusVal = colStatus ? (row[colStatus] || '') : '';
        let rawDate = "";
        if (colDataLancPeca && row[colDataLancPeca] && String(row[colDataLancPeca]).trim() !== '') rawDate = row[colDataLancPeca];
        else if (colDataOriginal && row[colDataOriginal]) rawDate = row[colDataOriginal];
        
        let dateIso = "N/A"; let year = null; let dtParsed = parseDateToIso(rawDate); if(dtParsed) { dateIso = dtParsed; year = parseInt(dtParsed.split('-')[0]); }
        let rawFim = colFinalizado ? (row[colFinalizado] || '') : ''; let fimIso = parseDateToIso(rawFim);
        let vencIso = null; let vencStr = "";
        
        if (colVencimento && row[colVencimento]) {
            vencStr = String(row[colVencimento]).trim(); vencIso = parseDateToIso(vencStr);
            if (vencIso) vencStr = vencIso.split('-').reverse().join('/');
        }

        let smart = calculateSmartStatus(statusVal, vencIso, fimIso);
        row['Status Prazo Calc'] = smart.text; 
        dbRawData.push(row);
        
        let isExplicitNTP = rowStr.includes('NTP') || statusVal.toUpperCase().includes('NTP');
        if (!isExplicitNTP && !row._forceNTP) return;
        
        let rawUnit = colUnidade ? (row[colUnidade] || '') : ''; let unitClean = cleanUnitName(rawUnit.replace(/"/g, '')); let rawPosto = colPosto ? (row[colPosto]||"") : "";
        data.push({ id: colId ? (row[colId]||'') : '', unidade: unitClean||'Indefinido', tecnico: colTecnico ? (row[colTecnico]||'').trim() : '', data: dateIso, ano: year, originalStatus: statusVal, finalizadoIso: fimIso, finalizado: rawFim, vencimento: vencStr, vencimentoIso: vencIso, nomePosto: rawPosto, extras: row });
    });
    ntpRawData = data; if(autoSave) { try { localStorage.setItem('ntp_dashboard_data_v10', JSON.stringify(ntpRawData)); localStorage.setItem('db_dashboard_data_v10', JSON.stringify({ cols: dbColumns, rows: dbRawData.slice(0, 300) })); processNTPData(); renderDatabase(); } catch(e) { } }
}

// ==========================================================
// 8. DASHBOARD SLA (LÓGICA PRINCIPAL E GRÁFICOS)
// ==========================================================
Chart.register(ChartDataLabels);
Chart.defaults.set('plugins.datalabels', { font: { weight: 'bold', size: 11 }, anchor: 'end', align: 'top', offset: -2, clamp: false, clip: false, display: (context) => context.dataset.data[context.dataIndex] > 0 });
Chart.defaults.layout.padding = { top: 20, right: 20, left: 10, bottom: 10 };

function populateUnidadeFilter() { const select = document.getElementById('unidadeFilter'); if(!select) return; const unidades = [...new Set(rawData.map(i => i.unidade))].sort(); select.innerHTML = '<option value="Todas">Todas Unidades</option>'; unidades.forEach(u => { const op = document.createElement('option'); op.value = u; op.innerText = u; select.appendChild(op); }); }
function changeUnidade(val) { activeUnidade = val; updateDashboard(); }
function changeEstado(val) { activeEstado = val; updateDashboard(); }

// NOVO: FILTRO DE ANO GERAL
function changeYear(year) { currentYear = (year === 'Todos') ? 'Todos' : parseInt(year); updateDashboard(); if(ntpRawData.length > 0) processNTPData(); }
function checkDateFilter(obj) { 
    if (currentYear === 'Todos') return true; 
    if (obj.ano !== undefined) return obj.ano === currentYear; 
    if (obj.data && obj.data !== 'N/A') { let y = parseInt(obj.data.split('-')[0]); return y === currentYear; }
    return true;
}

function toggleSlaTimeView() { const checkBox = document.getElementById('slaViewToggle'); slaTimeView = checkBox.checked ? 'daily' : 'monthly'; clearChartDateFilter(); }

function toggleChartDateFilter(label) {
    if (activeChartFilter === label) { clearChartDateFilter(); } else {
        activeChartFilter = label; activeChartFilterType = slaTimeView;
        document.getElementById('activeChartFilterContainer').classList.remove('hidden');
        document.getElementById('activeChartFilterLabel').innerText = label;
        updateDashboard(); if(ntpRawData.length > 0) filterNtpTable();
    }
}

function clearChartDateFilter() {
    activeChartFilter = null; activeChartFilterType = null;
    const container = document.getElementById('activeChartFilterContainer'); if(container) container.classList.add('hidden');
    updateDashboard(); if(ntpRawData.length > 0) filterNtpTable();
}

function updateDashboard() {
    // 1. Filtragem Inicial por Unidade / Estado / Ano
    let filteredObjects = rawData.filter(d => checkDateFilter(d));
    if (typeof activeUnidade !== 'undefined' && activeUnidade !== "Todas") filteredObjects = filteredObjects.filter(d => d.unidade === activeUnidade);
    const estSelect = document.getElementById('estadoFilter'); if(estSelect){ const uniqueEst = [...new Set(filteredObjects.map(d => d.estado))].sort(); const prev = estSelect.value; estSelect.innerHTML = '<option value="Todos">Todos Estados</option>'; uniqueEst.forEach(uf => { estSelect.innerHTML += `<option value="${uf}">${uf}</option>`; }); estSelect.value = uniqueEst.includes(prev) ? prev : "Todos"; }
    if(typeof activeEstado !== 'undefined' && activeEstado !== "Todos") filteredObjects = filteredObjects.filter(d => d.estado === activeEstado);

    // Dados Brutos (Sem o filtro do clique no próprio gráfico) - Usado para o gráfico de linha não quebrar
    let chartData = filteredObjects.map(obj => {
        return {
            ...obj,
            calcAberturas: obj.aberturas,
            calcFechamentos: obj.fechamentos,
            calcNtp: obj.ntps || 0,
            calcReinc: obj.reincidencias || 0,
            calcTroca: obj.trocas || 0,
            ordersListFiltered: obj.ordersList || [],
            fechamentosListFiltered: obj.fechamentosList || []
        };
    });

    // 2. Aplicação do filtro de Clique no Gráfico para Cards e Tabelas
    let displayData = filteredObjects.map(obj => {
        let calcAb = obj.aberturas; let calcFec = obj.fechamentos;
        let calcNtp = obj.ntps || 0; let calcReinc = obj.reincidencias || 0; let calcTroca = obj.trocas || 0;
        let oList = obj.ordersList || []; let fList = obj.fechamentosList || [];

        if (activeChartFilter) {
            if (activeChartFilterType === 'monthly') {
                const monthMap = { "Jan":"JANEIRO", "Fev":"FEVEREIRO", "Mar":"MARÇO", "Abr":"ABRIL", "Mai":"MAIO", "Jun":"JUNHO", "Jul":"JULHO", "Ago":"AGOSTO", "Set":"SETEMBRO", "Out":"OUTUBRO", "Nov":"NOVEMBRO", "Dez":"DEZEMBRO" };
                const fullMonth = monthMap[activeChartFilter];
                if (obj.mes !== fullMonth) { calcAb = 0; calcFec = 0; calcNtp = 0; calcReinc = 0; calcTroca = 0; oList = []; fList = []; }
            } else if (activeChartFilterType === 'daily') {
                oList = oList.filter(o => o.dateIso && `${o.dateIso.split('-')[2]}/${o.dateIso.split('-')[1]}` === activeChartFilter);
                fList = fList.filter(f => f.dateIso && `${f.dateIso.split('-')[2]}/${f.dateIso.split('-')[1]}` === activeChartFilter);
                calcAb = oList.length; calcFec = fList.length; calcNtp = oList.filter(o => o.isNtp).length; calcReinc = oList.filter(o => o.isReinc).length; calcTroca = oList.filter(o => o.isTroca).length;
            }
        }
        
        return { ...obj, calcAberturas: calcAb, calcFechamentos: calcFec, calcNtp: calcNtp, calcReinc: calcReinc, calcTroca: calcTroca, ordersListFiltered: oList, fechamentosListFiltered: fList };
    });

    lastDisplayData = displayData;

    // 3. Atualizar Indicadores Globais e Calcular TMA / Críticos
    let totalAbs = 0; let totalFec = 0; let totalNtp = 0; let totalReinc = 0; let totalTroca = 0;
    let tempoTotalMs = 0; let chamadosValidos = 0; let totalCriticos = 0;

    displayData.forEach(obj => { 
        totalAbs += obj.calcAberturas; totalFec += obj.calcFechamentos; totalNtp += obj.calcNtp; totalReinc += obj.calcReinc; totalTroca += obj.calcTroca; 
        
        if(obj.ordersListFiltered) {
            obj.ordersListFiltered.forEach(o => {
                const dyn = calculateSmartStatus(o.statusOriginal, o.vencimentoIso, o.dataFechamentoIso);
                if(dyn.text === 'Vencido') totalCriticos++;

                // Cálculo preciso do TMA lendo as colunas brutas AB e FC na mesma linha
                let dataABStr = o.rawRow ? (o.rawRow['AB'] || o.rawRow['Data Abertura'] || o.data) : o.data;
                let dataFCStr = o.rawRow ? (o.rawRow['FC'] || o.rawRow['Fechamento'] || o.dataFechamento) : o.dataFechamento;

                if (dataABStr && dataFCStr) {
                    let dtAB = converterDataBR(dataABStr);
                    let dtFC = converterDataBR(dataFCStr);
                    if (dtAB && dtFC && dtFC >= dtAB) {
                        tempoTotalMs += (dtFC.getTime() - dtAB.getTime());
                        chamadosValidos++;
                    }
                }
            });
        }
    });

    let slaRate = totalAbs > 0 ? (totalFec / totalAbs) * 100 : 0; if(slaRate > 100) slaRate = 100;
    
    document.getElementById('statAberturas').innerText = totalAbs.toLocaleString(); 
    document.getElementById('statFechamentos').innerText = totalFec.toLocaleString(); 
    document.getElementById('statSLA').innerText = slaRate.toFixed(1) + '%';
    const box = document.getElementById('slaColorBox'); if(box) box.className = `p-2.5 rounded-lg transition-colors shrink-0 ${slaRate>=90?'bg-emerald-100 text-emerald-600':(slaRate>=75?'bg-yellow-100 text-yellow-600':'bg-rose-100 text-rose-600')}`;
    
    document.getElementById('statNTP').innerText = totalNtp.toLocaleString();
    document.getElementById('percNTP').innerText = totalAbs > 0 ? ((totalNtp/totalAbs)*100).toFixed(1) + '%' : '0%';
    document.getElementById('statReincidencia').innerText = totalReinc.toLocaleString();
    document.getElementById('percReincidencia').innerText = totalAbs > 0 ? ((totalReinc/totalAbs)*100).toFixed(1) + '%' : '0%';
    document.getElementById('statTroca').innerText = totalTroca.toLocaleString();
    document.getElementById('percTroca').innerText = totalAbs > 0 ? ((totalTroca/totalAbs)*100).toFixed(1) + '%' : '0%';
    
    // Formatação inteligente do TMA Médio
    let tmaFormatted = "-";
    if (chamadosValidos > 0) {
        let tmaMedioMs = tempoTotalMs / chamadosValidos;
        let totalMinutos = Math.floor(tmaMedioMs / (1000 * 60));
        let horas = Math.floor(totalMinutos / 60);
        let minutos = totalMinutos % 60;
        let dias = Math.floor(horas / 24);
        let horasRestantes = horas % 24;
        
        if (dias > 0) {
            tmaFormatted = `${dias}d ${horasRestantes}h ${minutos}m`;
        } else {
            tmaFormatted = `${horas}h ${minutos}m`;
        }
    }

    document.getElementById('statTMA').innerText = tmaFormatted;
    document.getElementById('statCriticos').innerText = totalCriticos.toLocaleString();

    // 4. Passar Dados Corretos aos Gráficos
    updateSlaCharts(chartData, displayData); 
    
    // 5. Atualizar Tabela SLA
    filterSlaTable(displayData, true);
}

// Gráfico SLA e Gráfico de Volumes
function updateSlaCharts(fullData, filteredData) {
    let labels = []; let absData = []; let fecData = []; let slaData = [];
    
    if(slaTimeView === 'monthly') {
        const months = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"];
        
        // SLA Linha (Baseado em fullData para não quebrar o visual da linha)
        const fullMData = months.map(m => { 
            const sub = fullData.filter(d => d.mes === m); 
            const abs = sub.reduce((a,c) => a+(c.calcAberturas||0), 0); 
            const fec = sub.reduce((a,c) => a+(c.calcFechamentos||0), 0); 
            let sla = abs > 0 ? (fec/abs)*100 : 0; 
            if(sla > 100) sla = 100; 
            return { sla }; 
        });

        // Volumes Barras (Baseado em filteredData para que as barras respondam ao clique)
        const filtMData = months.map(m => { 
            const sub = filteredData.filter(d => d.mes === m); 
            const abs = sub.reduce((a,c) => a+(c.calcAberturas||0), 0); 
            const fec = sub.reduce((a,c) => a+(c.calcFechamentos||0), 0); 
            return { abs, fec }; 
        });

        labels = months.map(m=>m.substring(0,3)); 
        slaData = fullMData.map(d=>d.sla);
        absData = filtMData.map(d=>d.abs); 
        fecData = filtMData.map(d=>d.fec); 
    } else {
        let dailyFullMap = {}; 
        fullData.forEach(item => { 
            let oList = item.ordersListFiltered || item.ordersList || [];
            let fList = item.fechamentosListFiltered || item.fechamentosList || [];
            
            oList.forEach(o => { 
                if(o.dateIso) { if(!dailyFullMap[o.dateIso]) dailyFullMap[o.dateIso] = {abs:0, fec:0}; dailyFullMap[o.dateIso].abs++; } 
            }); 
            fList.forEach(f => { 
                if(f.dateIso) { if(!dailyFullMap[f.dateIso]) dailyFullMap[f.dateIso] = {abs:0, fec:0}; dailyFullMap[f.dateIso].fec++; } 
            }); 
        });

        let dailyFiltMap = {}; 
        filteredData.forEach(item => { 
            let oList = item.ordersListFiltered || item.ordersList || [];
            let fList = item.fechamentosListFiltered || item.fechamentosList || [];
            
            oList.forEach(o => { 
                if(o.dateIso) { if(!dailyFiltMap[o.dateIso]) dailyFiltMap[o.dateIso] = {abs:0, fec:0}; dailyFiltMap[o.dateIso].abs++; } 
            }); 
            fList.forEach(f => { 
                if(f.dateIso) { if(!dailyFiltMap[f.dateIso]) dailyFiltMap[f.dateIso] = {abs:0, fec:0}; dailyFiltMap[f.dateIso].fec++; } 
            }); 
        });
        
        let allDates = Object.keys(dailyFullMap).sort(); 
        
        if (allDates.length === 1) {
            const dIso = allDates[0]; const dObj = new Date(dIso + "T12:00:00");
            const prev = new Date(dObj); prev.setDate(dObj.getDate() - 1); const next = new Date(dObj); next.setDate(dObj.getDate() + 1);
            dailyFullMap[prev.toISOString().split('T')[0]] = {abs:0, fec:0}; dailyFullMap[next.toISOString().split('T')[0]] = {abs:0, fec:0};
            dailyFiltMap[prev.toISOString().split('T')[0]] = {abs:0, fec:0}; dailyFiltMap[next.toISOString().split('T')[0]] = {abs:0, fec:0};
            allDates = Object.keys(dailyFullMap).sort();
        }

        labels = allDates.map(d => { const parts = d.split('-'); return `${parts[2]}/${parts[1]}`; }); 
        slaData = allDates.map(d => { let abs = dailyFullMap[d].abs; let fec = dailyFullMap[d].fec; let sla = abs > 0 ? (fec/abs)*100 : 0; if (sla > 100) sla = 100; return sla; });
        absData = allDates.map(d => dailyFiltMap[d] ? dailyFiltMap[d].abs : 0); 
        fecData = allDates.map(d => dailyFiltMap[d] ? dailyFiltMap[d].fec : 0); 
    }

    const clickHandler = (e, elements) => { if (elements.length > 0) { const chart = e.chart; const index = elements[0].index; const label = chart.data.labels[index]; toggleChartDateFilter(label); } };

    if(charts['sla']) charts['sla'].destroy(); 
    const ctxSla = document.getElementById('slaChart').getContext('2d'); 
    const gradientSla = ctxSla.createLinearGradient(0, 0, 0, 300); gradientSla.addColorStop(0, 'rgba(37, 99, 235, 0.2)'); gradientSla.addColorStop(1, 'rgba(37, 99, 235, 0)');
    const isDark = document.body.classList.contains('dark'); const chartTextColor = isDark ? '#f8fafc' : '#64748b'; const gridColor = isDark ? 'rgba(148, 163, 184, 0.1)' : 'rgba(100, 116, 139, 0.1)'; 
    Chart.defaults.color = chartTextColor;

    charts['sla'] = new Chart(ctxSla, { 
        type: 'line', 
        data: { labels: labels, datasets: [{ label:'SLA %', data: slaData, extraAbs: absData, extraFec: fecData, borderColor:'#2563eb', borderWidth:3, tension:0.4, fill:true, backgroundColor:gradientSla, pointRadius: slaTimeView === 'daily' ? 2 : 0, clip: false, datalabels:{ color: chartTextColor, display: slaTimeView === 'daily' ? 'auto' : (context) => context.dataset.data[context.dataIndex] > 0, formatter:(v)=>v>0?v.toFixed(0)+'%':'', offset: slaTimeView === 'daily' ? 4 : 8, align:'top', clip: false, font: { size: slaTimeView === 'daily' ? 9 : 11 } } }] }, 
        options: { responsive:true, maintainAspectRatio:false, onClick: clickHandler, onHover: (event, chartElement) => { event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default'; }, layout: { padding: { top: 35, right: 20, left: 10, bottom: 10 } }, scales:{ y:{ min: 0, max: 110, ticks: { stepSize: 20, color: chartTextColor, callback: function(value) { return value <= 100 ? value + '%' : ''; } }, grid: { color: function(context) { if (context.tick.value > 100) return 'transparent'; return gridColor; }, drawBorder: false } }, x:{ grid:{display:false}, ticks: { maxTicksLimit: slaTimeView === 'daily' ? 12 : 15, color: chartTextColor, font: { size: slaTimeView === 'daily' ? 9 : 11 } } } }, plugins:{ legend:{display:false}, tooltip: customTooltipConfig } } 
    });

    if(charts['vol']) charts['vol'].destroy(); 
    charts['vol'] = new Chart(document.getElementById('volumeChart'), { 
        type: 'bar', 
        data: { labels: labels, datasets: [ { label:'Abert', data: absData, backgroundColor:'#fca5a5', borderRadius:3, barPercentage: 0.6, categoryPercentage: 0.8 }, { label:'Fech', data: fecData, backgroundColor:'#86efac', borderRadius:3, barPercentage: 0.6, categoryPercentage: 0.8 } ] }, 
        options: { responsive:true, maintainAspectRatio:false, onClick: clickHandler, onHover: (event, chartElement) => { event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default'; }, layout: { padding: { top: 30, right: 10, left: 10, bottom: 0 } }, plugins:{ legend:{position:'bottom', labels: { color: chartTextColor }}, datalabels: { color: chartTextColor, display: (context) => context.dataset.data[context.dataIndex] > 0, anchor: 'end', align: 'top', offset: 0, clip: false, font: { size: slaTimeView === 'daily' ? 10 : 10, weight: 'bold' }, formatter: function(value) { return value; } } }, scales:{ x:{ grid:{display:false}, ticks: { maxTicksLimit: 15, color: chartTextColor } }, y: { grace: '20%', ticks: { color: chartTextColor } } } } 
    });
}

// ==========================================================
// 9. TABELAS E MODAIS (SLA & PEDIDOS)
// ==========================================================
function openOrdersModal(idKey) { 
    const item = lastDisplayData.find(i => i.idKey === idKey); if(!item) return; 
    let list = item.ordersListFiltered || item.ordersList || []; 
    if(list.length === 0 && item.calcAberturas > 0) list = [{ os: 'Memória Otimizada', data: '-', vencimento: '-', statusOriginal: 'Detalhes ocultos para performance' }]; 
    currentOrdersList = list; document.getElementById('ordersModalTitle').innerText = `${item.unidade} - ${item.mes} (${item.ano})`; 
    ordersFilterState = {}; renderOrdersList(currentOrdersList); document.getElementById('ordersModal').classList.remove('hidden'); 
}
let currentOrdersList = [];

function renderOrdersList(list) { 
    const tbody = document.getElementById('ordersModalBody'); if(!tbody) return; 
    if(!list || list.length === 0) { tbody.innerHTML = '<tr><td colspan="5" class="p-4 text-center text-gray-400">Nenhum pedido neste período.</td></tr>'; return; } 
    let html = ''; 
    list.slice(0, 500).forEach(order => { 
        const osLink = (order.os && order.os !== 'N/A' && order.os !== 'Memória Otimizada') ? `<a href="https://imbera.telecontrol.com.br/assist/admin/os_press.php?os=${order.os}" target="_blank" class="text-blue-600 font-bold hover:underline">${order.os}</a>` : (order.os || 'N/A'); 
        const dynStatus = calculateSmartStatus(order.statusOriginal, order.vencimentoIso, order.dataFechamentoIso); 
        const colCombined = `${order.classificacaoSla || '-'} / ${order.statusOriginal || '-'}`;
        html += `<tr class="hover:bg-blue-50 transition-colors"><td class="px-4 py-3 border-b border-slate-100">${osLink}</td><td class="px-4 py-3 border-b border-slate-100 text-slate-500 text-[10px]">${order.data || ''}</td><td class="px-4 py-3 border-b border-slate-100 font-bold text-slate-700 text-[10px]">${order.vencimento || 'N/A'}</td><td class="px-4 py-3 border-b border-slate-100 text-[10px]"><span class="${dynStatus.class}">${dynStatus.text}</span></td><td class="px-4 py-3 border-b border-slate-100 text-slate-400 text-[10px] italic">${colCombined}</td></tr>`; 
    }); 
    tbody.innerHTML = html; 
}

function filterSlaTable(dataSet, isCalculated = false) {
    let data = dataSet || rawData; if (!isCalculated) data = rawData.filter(d => checkDateFilter(d));
    const text = document.getElementById('slaTableSearch')?.value?.toLowerCase() || ''; if(text) data = data.filter(r => (r.unidade||'').toLowerCase().includes(text));
    for(const k in slaFilterState){ const v = slaFilterState[k]; if(!v) continue; data = data.filter(r => (''+ (r[k]||'')).toString() === v); }
    const tbody = document.getElementById('tableBody'); if(!tbody) return; if(data.length === 0) { tbody.innerHTML = '<tr><td colspan="6" class="p-4 text-center text-sm text-slate-400 italic">Nenhum registro encontrado.</td></tr>'; return; }
    const htmlRows = [];
    data.slice(0, 300).forEach(row => {
        const abs = (row.calcAberturas !== undefined) ? row.calcAberturas : row.aberturas; const fec = (row.calcFechamentos !== undefined) ? row.calcFechamentos : row.fechamentos;
        let sla = abs > 0 ? (fec/abs)*100 : 0; if (sla > 100) sla = 100; 
        const color = sla >= 90 ? 'text-emerald-600' : (sla>=75?'text-yellow-600':'text-rose-600'); const btnHtml = `<button onclick="openOrdersModal('${row.idKey}')" class="bg-blue-50 hover:bg-blue-100 text-blue-600 p-1.5 rounded transition-colors"><i data-lucide="search" class="w-3.5 h-3.5"></i></button>`;
        
        const alertPing = sla < 80 ? `<span class="relative flex h-2 w-2 mr-2"><span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span><span class="relative inline-flex rounded-full h-2 w-2 bg-rose-500"></span></span>` : '';

        htmlRows.push(`<tr class="hover:bg-slate-50 transition-colors"><td class="px-6 py-3 font-normal text-slate-600 border-b border-slate-50 text-[11px]">${row.unidade}</td><td class="px-6 py-3 text-slate-500 font-normal border-b border-slate-50 text-[11px]">${row.mes}/${row.ano}</td><td class="px-6 py-3 border-b border-slate-50 text-[11px] text-center">${btnHtml}</td><td class="px-6 py-3 text-right text-slate-600 border-b border-slate-50 text-[11px] font-normal">${abs}</td><td class="px-6 py-3 text-right text-slate-600 border-b border-slate-50 text-[11px] font-normal">${fec}</td><td class="px-6 py-3 border-b border-slate-50"><div class="flex items-center justify-end font-black ${color} text-[11px]">${alertPing} ${sla.toFixed(1)}%</div></td></tr>`);
    });
    tbody.innerHTML = htmlRows.join(''); lucide.createIcons();
}

function toggleOrdersFilter(e, key) { e.stopPropagation(); ['os','data','vencimento','statusDyn','classificacao'].forEach(k => { if(k !== key) document.getElementById(`orders-filter-${k}`)?.classList.remove('show'); }); const dropdown = document.getElementById(`orders-filter-${key}`); if(!dropdown) return; if(dropdown.classList.contains('show')) { dropdown.classList.remove('show'); return; } let uniqueValues = new Set(); currentOrdersList.forEach(o => { let val = ''; if(key === 'os') val = o.os || 'N/A'; if(key === 'data') val = o.data || ''; if(key === 'vencimento') val = o.vencimento || 'N/A'; if(key === 'statusDyn') { const dynStatus = calculateSmartStatus(o.statusOriginal, o.vencimentoIso, o.dataFechamentoIso); val = dynStatus.text || ''; } if(key === 'classificacao') val = `${o.classificacaoSla || '-'} / ${o.statusOriginal || '-'}`; if(val) uniqueValues.add(val); }); const sortedValues = [...uniqueValues].sort(); let html = `<div onclick="applyOrdersFilter('${key}', 'ALL')" class="text-blue-600 font-bold">Limpar Filtro</div>`; sortedValues.forEach(v => { html += `<div onclick="applyOrdersFilter('${key}', '${(''+v).replace(/'/g,"\\'")}')" class="truncate">${v}</div>`; }); dropdown.innerHTML = html; dropdown.classList.add('show'); document.addEventListener('click', closeOrdersFilters); }
function closeOrdersFilters() { document.querySelectorAll('#ordersModal .db-select').forEach(el => el.classList.remove('show')); document.removeEventListener('click', closeOrdersFilters); }
function applyOrdersFilter(key, value) { if(value === 'ALL') delete ordersFilterState[key]; else ordersFilterState[key] = value; filterOrdersGrid(); }
function filterOrdersGrid() { const filtered = currentOrdersList.filter(o => { const dynStatus = calculateSmartStatus(o.statusOriginal, o.vencimentoIso, o.dataFechamentoIso); const combined = `${o.classificacaoSla || '-'} / ${o.statusOriginal || '-'}`; let match = true; if(ordersFilterState['os'] && o.os !== ordersFilterState['os']) match = false; if(ordersFilterState['data'] && o.data !== ordersFilterState['data']) match = false; if(ordersFilterState['vencimento'] && o.vencimento !== ordersFilterState['vencimento']) match = false; if(ordersFilterState['statusDyn'] && dynStatus.text !== ordersFilterState['statusDyn']) match = false; if(ordersFilterState['classificacao'] && combined !== ordersFilterState['classificacao']) match = false; return match; }); renderOrdersList(filtered); }

function toggleSlaFilter(e, key, triggerId) { e.stopPropagation(); ['sla-filter-unidade','sla-filter-mes'].forEach(id => { if(id !== `sla-filter-${key}`) document.getElementById(id)?.classList.remove('show'); }); const dropdown = document.getElementById(`sla-filter-${key}`); if(!dropdown) return; if(dropdown.classList.contains('show')) { dropdown.classList.remove('show'); return; } let filtered = lastDisplayData.filter(d => checkDateFilter(d)); if (typeof activeUnidade !== 'undefined' && activeUnidade !== "Todas") filtered = filtered.filter(d => d.unidade === activeUnidade); const uniqueValues = [...new Set(filtered.map(r => (''+(r[key]||''))))].filter(v=>v!=='').sort((a,b)=> (isNaN(a)? (''+a).localeCompare(b) : Number(a)-Number(b))); let html = `<div onclick="applySlaFilter('${key}', 'ALL')" class="text-blue-600 font-bold">Limpar Filtro</div>`; uniqueValues.forEach(v => { html += `<div onclick="applySlaFilter('${key}', '${(''+v).replace(/'/g,"\\'")}')" class="truncate">${v}</div>`; }); dropdown.innerHTML = html; dropdown.classList.add('show'); document.addEventListener('click', closeSlaFilters); }
function toggleNtpFilter(e, key, triggerId) { e.stopPropagation(); ['ntp-filter-data','ntp-filter-unidade','ntp-filter-status','ntp-filter-id'].forEach(id => { if(id !== `ntp-filter-${key}`) document.getElementById(id)?.classList.remove('show'); }); const dropdown = document.getElementById(`ntp-filter-${key}`); if(dropdown.classList.contains('show')) { dropdown.classList.remove('show'); return; } let filtered = ntpRawData.filter(d => checkDateFilter(d)); const uniqueValues = [...new Set(filtered.map(r => (r[key]||"") ))].filter(v=>v).sort(); let html = `<div onclick="applyNtpFilter('${key}', 'ALL')" class="text-blue-600 font-bold">Limpar Filtro</div>`; uniqueValues.forEach(v => { html += `<div onclick="applyNtpFilter('${key}', '${(v+"").replace(/'/g, "\\'")}')" class="truncate">${v}</div>`; }); dropdown.innerHTML = html; dropdown.classList.add('show'); document.addEventListener('click', closeNtpFilters); }
function closeSlaFilters(){ document.querySelectorAll('.db-select').forEach(el => el.classList.remove('show')); document.removeEventListener('click', closeSlaFilters); }
function closeNtpFilters() { document.querySelectorAll('.db-select').forEach(el => el.classList.remove('show')); document.removeEventListener('click', closeNtpFilters); }
function applySlaFilter(key, value){ if(value === 'ALL') delete slaFilterState[key]; else slaFilterState[key] = value; filterSlaTable(lastDisplayData, true); }
function applyNtpFilter(key, value) { if(value === 'ALL') delete ntpFilterState[key]; else ntpFilterState[key] = value; filterNtpTable(); }

// ==========================================================
// 10. DASHBOARD NTP (LÓGICA E GRÁFICOS)
// ==========================================================
function sendNtpToTab() {
    let ntpRows = []; let filteredObjects = rawData.filter(d => checkDateFilter(d));
    if (typeof activeUnidade !== 'undefined' && activeUnidade !== "Todas") filteredObjects = filteredObjects.filter(d => d.unidade === activeUnidade);
    if (typeof activeEstado !== 'undefined' && activeEstado !== "Todos") filteredObjects = filteredObjects.filter(d => d.estado === activeEstado);
    filteredObjects.forEach(item => { if(item.ordersList) { item.ordersList.forEach(order => { if(order.isNtp && order.rawRow) { ntpRows.push({ ...order.rawRow, _forceNTP: true }); } }); } });
    if (ntpRows.length === 0) { alert("Nenhuma NTP encontrada nos dados filtrados."); return; }
    let headers = Object.keys(ntpRows[0]); let existingOs = new Set(ntpRawData.map(n => n.id));
    let newRows = ntpRows.filter(r => { let osKey = r['OS'] || r['os'] || r['Chamado'] || r['ID']; if (!osKey) return true; if (existingOs.has(osKey)) return false; return true; });
    if (newRows.length > 0) { parseNTPCSV(newRows, headers, true); alert(`${newRows.length} novas NTPs enviadas para a aba NTP!`); } else { alert("As NTPs atuais já foram enviadas ou cadastradas na aba NTP."); }
    switchTab('ntp');
}

function processNTPData() { let baseData = ntpRawData.filter(d => d.ano !== null && checkDateFilter(d)); const units = ["Todos", ...new Set(baseData.map(d => d.unidade))].sort(); const uSelect = document.getElementById('ntpUnitFilter'); if(uSelect){ uSelect.innerHTML = ''; units.forEach(u => { const op = document.createElement('option'); op.value = u; op.innerText = u; uSelect.appendChild(op); }); uSelect.value = "Todos"; } updateNtpFilters('unit'); }

function updateNtpFilters(changedLevel) { let filtered = ntpRawData.filter(d => checkDateFilter(d)); const uVal = document.getElementById('ntpUnitFilter')?.value || 'Todos'; const pSelect = document.getElementById('ntpPostoFilter'); if(uVal !== "Todos") filtered = filtered.filter(d => d.unidade === uVal); if(changedLevel === 'unit' && pSelect) { const postos = ["Todos", ...new Set(filtered.map(d => d.nomePosto))].sort(); pSelect.innerHTML = ''; postos.forEach(p => { const op = document.createElement('option'); op.value = p; op.innerText = p; pSelect.appendChild(op); }); pSelect.value = "Todos"; } filterNtpTable(); }
function toggleNtpTimeView() { const checkBox = document.getElementById('ntpViewToggle'); ntpTimeView = checkBox.checked ? 'daily' : 'monthly'; filterNtpTable(); }
function clearAllNtpFilters() { ntpFilterState = {}; document.getElementById('ntpUnitFilter').value = 'Todos'; document.getElementById('ntpPostoFilter').value = 'Todos'; document.getElementById('ntpSearch').value = ''; updateNtpFilters('unit'); }

function filterNtpTable() { 
    let data = ntpRawData.filter(d => checkDateFilter(d)); 
    const uVal = document.getElementById('ntpUnitFilter')?.value || 'Todos'; 
    const pVal = document.getElementById('ntpPostoFilter')?.value || 'Todos'; 
    const search = (document.getElementById('ntpSearch')?.value || '').toLowerCase(); 
    
    if(uVal !== "Todos") data = data.filter(d => d.unidade === uVal); 
    if(pVal !== "Todos") data = data.filter(d => d.nomePosto === pVal); 
    if(search) data = data.filter(d => (d.id||"").toLowerCase().includes(search) || (d.unidade||"").toLowerCase().includes(search) || (d.status||"").toLowerCase().includes(search) || (d.data||"").toLowerCase().includes(search)); 
    
    for(const k in ntpFilterState) { const v = ntpFilterState[k]; if(!v) continue; data = data.filter(r => ((r[k]||"") + "").toString() === v); } 
    
    if (activeChartFilter) {
        if (activeChartFilterType === 'monthly') {
            const months = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
            const mIdx = months.indexOf(activeChartFilter);
            data = data.filter(d => { if(!d.data || d.data === 'N/A') return false; const parts = d.data.split('-'); return parseInt(parts[1]) - 1 === mIdx; });
        } else if (activeChartFilterType === 'daily') {
            data = data.filter(d => { if(!d.data || d.data === 'N/A') return false; const parts = d.data.split('-'); return `${parts[2]}/${parts[1]}` === activeChartFilter; });
        }
    }
    updateNtpVisuals(data); 
}

// Gráficos NTP
function updateNtpVisuals(data) {
    const isDark = document.body.classList.contains('dark'); const chartTextColor = isDark ? '#f8fafc' : '#64748b'; Chart.defaults.color = chartTextColor;
    
    let labelsLine = []; let ntpChartData = [];
    if (ntpTimeView === 'monthly') {
        labelsLine = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
        const monthCounts = Array(12).fill(0);
        data.forEach(d => { if(d.data && d.data !== 'N/A') { const parts = d.data.split('-'); const monthIndex = parseInt(parts[1]) - 1; if (monthIndex >= 0 && monthIndex < 12) monthCounts[monthIndex]++; } });
        ntpChartData = monthCounts;
    } else {
        let dailyMap = {};
        data.forEach(d => { if(d.data && d.data !== 'N/A') { if(!dailyMap[d.data]) dailyMap[d.data] = 0; dailyMap[d.data]++; } });
        let allDates = Object.keys(dailyMap).sort();
        if (allDates.length === 1) { const dIso = allDates[0]; const dObj = new Date(dIso + "T12:00:00"); const prev = new Date(dObj); prev.setDate(dObj.getDate() - 1); const next = new Date(dObj); next.setDate(dObj.getDate() + 1); dailyMap[prev.toISOString().split('T')[0]] = 0; dailyMap[next.toISOString().split('T')[0]] = 0; allDates = Object.keys(dailyMap).sort(); }
        labelsLine = allDates.map(d => { const parts = d.split('-'); return `${parts[2]}/${parts[1]}`; });
        ntpChartData = allDates.map(d => dailyMap[d]);
    }

    const statusCounts = { 'Finalizados': Array(12).fill(0), 'Pendentes': Array(12).fill(0) }; 
    data.forEach(d => { if(d.data && d.data !== 'N/A') { const parts = d.data.split('-'); const monthIndex = parseInt(parts[1]) - 1; if (monthIndex >= 0 && monthIndex < 12) { if((d.originalStatus || d.status || "").toUpperCase().includes("FINAL") || d.finalizado) statusCounts['Finalizados'][monthIndex]++; else statusCounts['Pendentes'][monthIndex]++; } } });

    if(charts['ntpLine']) charts['ntpLine'].destroy(); 
    const ctxLine = document.getElementById('ntpChart').getContext('2d'); 
    const gradLine = ctxLine.createLinearGradient(0, 0, 0, 300); gradLine.addColorStop(0, 'rgba(244, 63, 94, 0.4)'); gradLine.addColorStop(1, 'rgba(244, 63, 94, 0)'); 
    charts['ntpLine'] = new Chart(ctxLine, { type: 'line', data: { labels: labelsLine, datasets: [{ label: 'NTPs', data: ntpChartData, borderColor: '#f43f5e', borderWidth: 3, backgroundColor: gradLine, fill: true, tension: 0.4, pointRadius: ntpTimeView === 'daily' ? 2 : 0 }] }, options: { responsive:true, maintainAspectRatio:false, layout: { padding: { top: 20, right: 10, left: 10, bottom: 0 } }, plugins:{legend:{display:false}}, scales:{ x:{grid:{display:false}, ticks: {color: chartTextColor, maxTicksLimit: ntpTimeView === 'daily' ? 12 : 12}}, y:{beginAtZero:true, grace: '5%', ticks: {color: chartTextColor}} } }});

    const counts = {}; data.forEach(d => { counts[d.unidade] = (counts[d.unidade]||0) + 1; }); const sorted = Object.entries(counts).sort((a,b) => b[1]-a[1]).slice(0, 10); 
    if(charts['ntpBar']) charts['ntpBar'].destroy(); charts['ntpBar'] = new Chart(document.getElementById('ntpBarChart'), { type: 'bar', data: { labels: sorted.map(i => i[0]), datasets: [{ label: 'Qtd', data: sorted.map(i => i[1]), backgroundColor: ['#1e293b', '#334155', '#475569', '#64748b', '#94a3b8'], borderRadius:4 }] }, options: { indexAxis:'y', responsive:true, maintainAspectRatio:false, layout: { padding: { right: 30, left: 0 } }, plugins:{legend:{display:false}}, scales:{x:{display:false, ticks: {color: chartTextColor}}, y:{grid:{display:false}, ticks: {color: chartTextColor}}} }});
    
    if(charts['ntpYear']) charts['ntpYear'].destroy(); charts['ntpYear'] = new Chart(document.getElementById('ntpYearChart'), { type: 'bar', data: { labels: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"], datasets: [{ label: 'Finalizados', data: statusCounts['Finalizados'], backgroundColor: '#10b981' }, { label: 'Pendentes', data: statusCounts['Pendentes'], backgroundColor: '#f43f5e' }] }, options: { responsive:true, maintainAspectRatio:false, layout: { padding: { top: 20, right: 0 } }, plugins:{ legend:{ position:'top', labels:{boxWidth:10, font:{size:9}, color: chartTextColor} }, datalabels: { display: false } }, scales:{ x:{ stacked:true, grid:{display:false}, ticks: {color: chartTextColor} }, y:{ stacked:true, beginAtZero:true, grace: '5%', ticks: {color: chartTextColor} } } }});
    
    const totalAll = ntpRawData.length || 1; document.getElementById('ntpPercentBadge').innerText = `${data.length} NTPs (${((data.length/totalAll)*100).toFixed(1)}% do DB)`;
    
    const tbody = document.getElementById('ntpTableBody'); if(!tbody) return; if(data.length === 0) { tbody.innerHTML = '<tr><td colspan="4" class="p-4 text-center text-sm text-slate-400 italic">Nenhum registro encontrado.</td></tr>'; return; }
    const htmlRows = []; 
    data.slice(0, 300).forEach((row) => { 
        const dynSmart = calculateSmartStatus(row.originalStatus || row.status, row.vencimentoIso, row.finalizadoIso);
        row.status = dynSmart.text; row.statusClass = dynSmart.class;
        const rowJson = JSON.stringify(row).replace(/'/g, "&apos;"); 
        htmlRows.push(`<tr class="hover:bg-rose-50 cursor-pointer transition-colors" onclick='openModal("infoModal", ${rowJson})'>
            <td class="px-6 py-3 text-slate-500 border-b border-slate-50 text-[10px]">${row.data}</td>
            <td class="px-6 py-3 font-bold text-slate-700 border-b border-slate-50 text-[10px]">${row.unidade}</td>
            <td class="px-6 py-3 border-b border-slate-50 text-[10px]"><span class="${dynSmart.class}">${dynSmart.text}</span></td>
            <td class="px-6 py-3 border-b border-slate-50"><a href="https://imbera.telecontrol.com.br/assist/admin/os_press.php?os=${row.id}" target="_blank" onclick="event.stopPropagation()" class="bg-rose-100 text-rose-700 px-2 py-0.5 rounded text-[9px] font-bold hover:bg-rose-200 hover:text-rose-900 transition-colors cursor-pointer" title="Abrir OS">${row.id}</a></td>
        </tr>`); 
    }); 
    tbody.innerHTML = htmlRows.join('');
}

function openModal(modalId, rowData = null) {
    if(modalId === 'infoModal' && rowData) { 
        const content = document.getElementById('modalContent'); 
        let vencHTML = `<p class="font-bold text-slate-400 italic">Não informado</p>`; 
        if(rowData.vencimento) { 
            const isExpired = rowData.statusClass.includes('red') || rowData.statusClass.includes('rose'); 
            const color = isExpired ? 'text-rose-600' : (rowData.statusClass.includes('orange')?'text-orange-600':'text-blue-600'); 
            vencHTML = `<p class="font-black ${color} text-base">${rowData.vencimento}</p>`; 
        } 
        let html = `<div class="space-y-4">
            <div class="grid grid-cols-2 gap-4">
                <div class="bg-slate-50 p-3 rounded-lg border border-slate-100"><p class="text-[10px] uppercase text-slate-400 font-bold">Unidade</p><p class="font-bold text-slate-700 text-xs">${rowData.unidade}</p></div>
                <div class="bg-slate-50 p-3 rounded-lg border border-slate-100"><p class="text-[10px] uppercase text-slate-400 font-bold">Técnico</p><p class="font-bold text-slate-700 text-xs">${rowData.tecnico}</p></div>
                <div class="bg-blue-50 p-3 rounded-lg border border-blue-100"><p class="text-[10px] uppercase text-blue-400 font-bold">Posto</p><p class="font-bold text-blue-700 text-xs">${rowData.nomePosto}</p></div>
                <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100"><p class="text-[10px] uppercase text-emerald-400 font-bold">Finalizado Em</p><p class="font-bold text-emerald-700 text-xs">${rowData.finalizado || 'Em andamento'}</p></div>
            </div>
            <div class="bg-yellow-50 p-4 rounded-xl border border-yellow-100 text-center">
                <p class="text-[10px] uppercase text-yellow-600 font-black tracking-widest mb-1">DATA DE VENCIMENTO</p>${vencHTML}
                <div class="mt-2 text-xs font-bold text-slate-500 border-t pt-2 border-yellow-200">Status Atual: <span class="${rowData.statusClass}">${rowData.status}</span></div>
            </div>
            <h4 class="font-bold text-slate-800 text-xs border-b pb-2 pt-2 uppercase tracking-wide">Dados Completos</h4>
            <div class="grid grid-cols-1 md:grid-cols-1 gap-x-4 gap-y-1 max-h-90 overflow-y-auto custom-scrollbar pr-2">`; 
        
        if(rowData.extras) { 
            const skipKeys = ['_forceNTP', 'Vencimento Calculado (Auto)', 'Status Prazo Calc'];
            Object.entries(rowData.extras).forEach(([k,v]) => { 
                if(!skipKeys.includes(k) && v && (''+v).trim() !== "") {
                    html += `<div class="flex justify-between border-b border-slate-50 py-1 text-[10px]"><span class="text-slate-500 font-medium">${k}:</span><span class="font-bold text-slate-700 text-right max-w-[300px]" title="${v}">${v}</span></div>`; 
                }
            }); 
        } 
        html += `</div></div>`; 
        content.innerHTML = html; 
    } 
    document.getElementById(modalId).classList.remove('hidden');
}

function closeModal(modalId) { document.getElementById(modalId).classList.add('hidden'); }

// ==========================================================
// 11. CALCULADORA DE VENCIMENTOS (EXCEL PARA CSV)
// ==========================================================
function handleCalculatorFile(event) {
    const file = event.target.files[0]; if (!file) return;
    const statusBox = document.getElementById('calcStatusBox'); const statusTitle = document.getElementById('calcStatusTitle'); const statusMsg = document.getElementById('calcStatusMsg'); const statusIcon = document.getElementById('calcStatusIcon'); const previewSection = document.getElementById('calcPreviewSection');
    statusBox.classList.remove('hidden', 'bg-red-50', 'bg-green-50', 'text-red-800', 'text-green-800', 'border-red-200', 'border-green-200'); statusBox.classList.add('bg-blue-50', 'text-blue-800', 'border-blue-200'); statusTitle.innerText = "Lendo Excel..."; statusMsg.innerText = "Analisando formato das colunas..."; previewSection.classList.add('hidden');
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result); const workbook = XLSX.read(data, { type: 'array' }); const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { defval: "" }); let countCalculated = 0;
            calcProcessedData = jsonSheet.map(row => {
                const getVal = (key) => { const foundKey = Object.keys(row).find(k => k.trim().toUpperCase() === key.toUpperCase()); return foundKey ? row[foundKey] : ''; };
                const rawClassificacao = getVal('CLASSIFICACAO CHAMADO') || getVal('SLA DE ATENDIMENTO') || getVal('PRIORIDADE'); const rawAb = getVal('ab') || getVal('Data Abertura'); 
                let dateIso = null; let dataAbCorrigida = rawAb; 
                if (typeof rawAb === 'number') { const jsDate = XLSX.SSF.parse_date_code(rawAb); dateIso = `${jsDate.y}-${String(jsDate.m).padStart(2,'0')}-${String(jsDate.d).padStart(2,'0')}`; dataAbCorrigida = `${String(jsDate.d).padStart(2,'0')}/${String(jsDate.m).padStart(2,'0')}/${jsDate.y}`; } else { dateIso = parseDateToIso(rawAb); }
                let vencimentoStr = ""; const classificacao = rawClassificacao || "BAIXO";
                if (dateIso) { const vencIso = calcularVencimento(dateIso, classificacao, 'SP'); if (vencIso) { vencimentoStr = vencIso.split('-').reverse().join('/'); countCalculated++; } }
                let novaLinha = { ...row }; const abKey = Object.keys(novaLinha).find(k => k.trim().toUpperCase() === 'AB'); if (abKey) { novaLinha[abKey] = dataAbCorrigida; }
                novaLinha["DATA VENCIMENTO CALCULADA"] = vencimentoStr; return novaLinha;
            });
            statusBox.classList.remove('bg-blue-50', 'text-blue-800', 'border-blue-200'); statusBox.classList.add('bg-green-50', 'text-green-800', 'border-green-200'); statusTitle.innerText = "Leitura Concluída!"; statusMsg.innerText = `${countCalculated} prazos calculados de ${jsonSheet.length} linhas importadas.`; statusIcon.classList.remove('text-blue-800'); statusIcon.classList.add('text-green-600');
            renderCalcPreview(calcProcessedData);
        } catch (err) {
            console.error(err); statusBox.classList.remove('bg-blue-50', 'text-blue-800', 'border-blue-200'); statusBox.classList.add('bg-red-50', 'text-red-800', 'border-red-200'); statusTitle.innerText = "Erro"; statusMsg.innerText = "Falha ao ler o arquivo Excel. Verifique se é um arquivo .xlsx válido.";
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderCalcPreview(data) {
    const previewSection = document.getElementById('calcPreviewSection'); const thead = document.getElementById('calcPreviewHead'); const tbody = document.getElementById('calcPreviewBody');
    if (!data || data.length === 0) return;
    const keys = Object.keys(data[0]); let headHtml = '<tr class="bg-slate-100/50 text-slate-500">';
    keys.forEach(k => { const isCalc = k === "DATA VENCIMENTO CALCULADA" || k.toUpperCase() === 'AB'; const colorClass = isCalc ? 'text-blue-600 font-black' : 'font-semibold'; headHtml += `<th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap ${colorClass}">${k}</th>`; });
    headHtml += '</tr>'; thead.innerHTML = headHtml;
    let bodyHtml = '';
    data.slice(0, 30).forEach(row => {
        bodyHtml += '<tr class="hover:bg-slate-50">';
        keys.forEach(k => { const val = row[k] || ''; const isCalc = k === "DATA VENCIMENTO CALCULADA" || k.toUpperCase() === 'AB'; const bgClass = isCalc ? 'bg-blue-50/20 font-bold text-blue-800' : 'text-slate-600'; bodyHtml += `<td class="px-4 py-2 border-b border-slate-50 truncate max-w-[150px] ${bgClass}">${val}</td>`; });
        bodyHtml += '</tr>';
    });
    if(data.length > 30) { bodyHtml += `<tr><td colspan="${keys.length}" class="px-4 py-3 text-center text-[10px] text-slate-400 italic bg-slate-50/50">+ ${data.length - 30} linhas processadas e ocultas nesta prévia</td></tr>`; }
    tbody.innerHTML = bodyHtml; previewSection.classList.remove('hidden'); lucide.createIcons();
}

function downloadCalculatedCSV() {
    if (!calcProcessedData || calcProcessedData.length === 0) return;
    const csv = Papa.unparse(calcProcessedData, { delimiter: ";", quotes: true, header: false }); const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.setAttribute("href", url); link.setAttribute("download", "dados_sistema_imbera.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
}

// ==========================================================
// 12. TRACKING OUTBOUND
// ==========================================================
let repairProcessedData = []; let repairFullData = []; 
function handleRepairFile(event) {
    const file = event.target.files[0]; if (!file) return;
    const osInput = document.getElementById('repairOsInput').value; const targetOS = osInput.split(/[\n,;]+/).map(s => s.trim().toUpperCase()).filter(s => s);
    if (targetOS.length === 0) { alert("Por favor, cole as OS Cliente antes de importar a planilha."); event.target.value = ''; return; }
    document.getElementById('loadingOverlay').style.display = 'flex'; document.getElementById('loadingText').innerText = "Filtrando e cruzando diagnósticos inteligentemente...";

    Papa.parse(file, {
        header: false, skipEmptyLines: true, encoding: "ISO-8859-1",
        complete: function(results) {
            let data = results.data; if (!data || data.length === 0) { hideLoading(); return; }
            let repairColumns = data[0]; let filteredData = [repairColumns]; let fullData = [repairColumns]; let count = 0;
            let idxAL = 37; let idxAM = 38; let idxAU = 46;
            let hAU = repairColumns.findIndex(h => typeof h === 'string' && (h.toUpperCase().includes('OS CLIENTE') || h.toUpperCase() === 'AU'));
            let hAL = repairColumns.findIndex(h => typeof h === 'string' && h.toUpperCase().includes('DEFEITO CONSTATADO'));
            let hAM = repairColumns.findIndex(h => typeof h === 'string' && h.toUpperCase().includes('SOLU'));
            if(hAU !== -1) idxAU = hAU; if(hAL !== -1) idxAL = hAL; if(hAM !== -1) idxAM = hAM;

            for (let i = 1; i < data.length; i++) {
                let row = data[i]; let osCliente = (row[idxAU] || "").toString().trim().toUpperCase();
                if (targetOS.includes(osCliente)) {
                    let defeitosRaw = row[idxAL] || ""; let solucoesRaw = row[idxAM] || "";
                    let matched = matchIntelligentDefect(defeitosRaw, solucoesRaw);
                    row[idxAL] = matched.d; row[idxAM] = matched.s;
                    filteredData.push(row); count++;
                }
                fullData.push(row); 
            }
            repairProcessedData = filteredData; repairFullData = fullData; 
            renderRepairPreview(idxAL, idxAM, idxAU); document.getElementById('loadingOverlay').style.display = 'none';
            if(count === 0) { alert("Nenhuma das OS informadas foi encontrada no arquivo CSV."); } else { alert(`Processamento concluído! ${count} OS foram localizadas, filtradas e corrigidas na base.`); }
            event.target.value = '';
        },
        error: function(err) { console.error(err); document.getElementById('loadingOverlay').style.display = 'none'; alert("Erro ao ler o arquivo CSV."); event.target.value = ''; }
    });
}

function matchIntelligentDefect(defRaw, solRaw) {
    if (!defRaw && !solRaw) return { d: "", s: "" };
    let defs = defRaw.split(',').map(s => s.trim()).filter(s => s); let sols = solRaw.split(',').map(s => s.trim()).filter(s => s);
    if (defs.length <= 1 && sols.length <= 1) { return { d: defs[0] || defRaw, s: sols[0] || solRaw }; }
    const getKeywords = (str) => str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").split(/[\s/]+/);
    let bestMatch = { d: defs[0] || defRaw, s: sols[0] || solRaw, score: -1 };
    const categoryMaps = [
        ['gas', 'refrigeracao', 'vazamento', 'carga', 'refrigerante', 'fluido', 'filtro', 'solda', 'obstrucao', 'capilar'],
        ['motor', 'micromotor', 'condensador', 'ventilador', 'helice', 'compressor', 'queimado', 'rele', 'protetor', 'ptc'],
        ['placa', 'eletronica', 'display', 'modulo', 'sensor', 'potenciometro', 'chicote', 'cabo', 'fiacao', 'painel', 'controlador'],
        ['porta', 'borracha', 'gaxeta', 'dobradica', 'mola', 'puxador', 'vidro'],
        ['termostato', 'temperatura', 'gelando', 'congelando', 'regulagem', 'calibracao'],
        ['limpeza', 'sujo', 'sujeira', 'higienizacao', 'preventiva', 'dreno']
    ];
    for (let d of defs) { let dKeys = getKeywords(d); for (let s of sols) { let sKeys = getKeywords(s); let score = 0; score += dKeys.filter(k => k.length > 3 && sKeys.includes(k)).length * 2; for (let cat of categoryMaps) { let dHasCat = dKeys.some(k => cat.includes(k)); let sHasCat = sKeys.some(k => cat.includes(k)); if (dHasCat && sHasCat) score += 5; } if (score > bestMatch.score) { bestMatch = { d: d, s: s, score: score }; } } }
    return bestMatch;
}

function renderRepairPreview(idxAL = 37, idxAM = 38, idxAU = 46) {
    const previewSection = document.getElementById('repairPreviewSection'); const thead = document.getElementById('repairPreviewHead'); const tbody = document.getElementById('repairPreviewBody');
    if (!repairProcessedData || repairProcessedData.length <= 1) return;
    let headHtml = '<tr class="bg-slate-100/50 text-slate-500"><th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap">Status</th><th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap font-bold text-orange-600">OS Cliente (AU)</th><th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap font-bold text-blue-600">Defeito Final (AL)</th><th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap font-bold text-emerald-600">Solução Final (AM)</th></tr>';
    thead.innerHTML = headHtml; let bodyHtml = '';
    for (let i = 1; i < Math.min(repairProcessedData.length, 100); i++) {
        let row = repairProcessedData[i];
        bodyHtml += `<tr class="hover:bg-slate-50"><td class="px-4 py-2 border-b border-slate-50"><span class="bg-emerald-100 text-emerald-600 px-2 rounded text-[9px] font-bold border border-emerald-200"><i data-lucide="check" class="w-3 h-3 inline"></i> Otimizado</span></td><td class="px-4 py-2 border-b border-slate-50 font-bold text-slate-700">${row[idxAU] || ''}</td><td class="px-4 py-2 border-b border-slate-50 text-slate-600 truncate max-w-[250px] bg-blue-50/20" title="${row[idxAL] || ''}">${row[idxAL] || ''}</td><td class="px-4 py-2 border-b border-slate-50 text-slate-600 truncate max-w-[250px] bg-emerald-50/20" title="${row[idxAM] || ''}">${row[idxAM] || ''}</td></tr>`;
    }
    if(repairProcessedData.length > 100) { bodyHtml += `<tr><td colspan="4" class="px-4 py-3 text-center text-[10px] text-slate-400 italic bg-slate-50/50">+ ${repairProcessedData.length - 100} linhas processadas e ocultas nesta prévia</td></tr>`; }
    tbody.innerHTML = bodyHtml; previewSection.classList.remove('hidden'); lucide.createIcons();
}

function downloadRepairCSV() {
    if (!repairProcessedData || repairProcessedData.length === 0) return;
    const csv = Papa.unparse(repairProcessedData, { delimiter: ";", quotes: true }); const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.setAttribute("href", url); link.setAttribute("download", "os_corrigidas_filtradas.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
}
function downloadFullRepairCSV() {
    if (!repairFullData || repairFullData.length === 0) return;
    const csv = Papa.unparse(repairFullData, { delimiter: ";", quotes: true }); const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.setAttribute("href", url); link.setAttribute("download", "base_completa_corrigida.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
}

// ==========================================================
// 13. BASE DE DADOS (VISUALIZAÇÃO COMPLETA)
// ==========================================================
function renderDatabase() { 
    if(!dbColumns.length) return; 
    const thead = document.getElementById('dbHeader'); 
    let hHtml = '<tr class="bg-slate-100 text-slate-600 sticky top-0 z-10">'; 
    dbColumns.forEach((col, idx) => { hHtml += `<th class="px-4 py-3 font-bold text-[10px] border-b border-slate-200 group min-w-[120px]"><div class="th-content" id="th-db-${idx}">${col}<span onclick="toggleDbFilter(event, ${idx}, 'th-db-${idx}')" class="filter-icon">▼</span></div><div id="filter-dropdown-${idx}" class="db-select custom-scrollbar"></div></th>`; }); 
    hHtml += '</tr>'; thead.innerHTML = hHtml; 
    filterDbTable(); 
}

function toggleDbFilter(e, colIdx, triggerId) { 
    e.stopPropagation(); document.querySelectorAll('.db-select').forEach(el => { if(el.id !== `filter-dropdown-${colIdx}`) el.classList.remove('show'); }); 
    const dropdown = document.getElementById(`filter-dropdown-${colIdx}`); if(dropdown.classList.contains('show')) { dropdown.classList.remove('show'); return; } 
    const filteredForDropdown = dbRawData; const uniqueValues = [...new Set(filteredForDropdown.map(r => r[dbColumns[colIdx]]))].filter(v=>v).sort(); 
    let html = `<div onclick="applyDbFilter(${colIdx}, 'ALL')" class="text-blue-600 font-bold">Limpar Filtro</div>`; 
    uniqueValues.forEach(v => { html += `<div onclick="applyDbFilter(${colIdx}, '${(v+"").replace(/'/g, "\\'")}')" class="truncate">${v}</div>`; }); 
    dropdown.innerHTML = html; dropdown.classList.add('show'); document.addEventListener('click', closeDbFilters); 
}

function closeDbFilters() { document.querySelectorAll('.db-select').forEach(el => el.classList.remove('show')); document.removeEventListener('click', closeDbFilters); }
function applyDbFilter(colIdx, value) { if(value === 'ALL') delete dbFilterState[colIdx]; else dbFilterState[colIdx] = value; filterDbTable(); }

function filterDbTable() { 
    const tbody = document.getElementById('dbBody'); if(!tbody) return; 
    const dbStart = document.getElementById('dbDateStart').value; const dbEnd = document.getElementById('dbDateEnd').value; 
    let dateColName = findHeader(dbColumns, ['AB', 'Data Abertura', 'date', 'emissao', 'data']); 
    
    let filtered = dbRawData.filter(row => { 
        for(let idx in dbFilterState) { if(row[dbColumns[idx]] !== dbFilterState[idx]) return false; } 
        if((dbStart || dbEnd) && dateColName) { 
            const rowDate = parseDateToIso(row[dateColName]); 
            if(rowDate && dbStart && rowDate < dbStart) return false; 
            if(rowDate && dbEnd && rowDate > dbEnd) return false; 
        } 
        return true; 
    }); 
    
    if(filtered.length === 0) { tbody.innerHTML = `<tr><td colspan="${dbColumns.length}" class="p-4 text-center text-sm text-slate-400 italic">Nenhum registro encontrado.</td></tr>`; return; } 
    
    let htmlRows = []; 
    filtered.slice(0, 500).forEach(row => { 
        let tr = '<tr class="hover:bg-slate-50">'; 
        dbColumns.forEach(col => { tr += `<td class="px-4 py-2 border-b border-slate-100 text-[10px] truncate max-w-[150px]">${row[col] || ''}</td>`; }); 
        tr += '</tr>'; htmlRows.push(tr); 
    }); 
    if(filtered.length > 500) htmlRows.push(`<tr><td colspan="${dbColumns.length}" class="p-2 text-center text-xs text-slate-400 italic">Mostrando 500 de ${filtered.length} registros</td></tr>`); 
    tbody.innerHTML = htmlRows.join(''); 
}

function exportarDBCsv() {
    if (!dbRawData || dbRawData.length === 0) return alert("Não há dados para exportar.");
    let headerRow = dbColumns.join(";"); let csvContent = headerRow + "\n";
    dbRawData.forEach(row => { let rowArray = dbColumns.map(col => { let val = row[col] || ""; return `"${String(val).replace(/"/g, '""')}"`; }); csvContent += rowArray.join(";") + "\n"; });
    let blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' }); let url = URL.createObjectURL(blob); let a = document.createElement("a"); a.href = url; a.download = "imbera_base_dados_calculada.csv"; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ==========================================================
// 14. PRODUTIVIDADE E SLA (EXCEL PARA DASHBOARD)
// ==========================================================

// Feriados enviados combinados até 2026
const feriadosProdutividade = new Set([
    "2022-01-01", "2022-03-01", "2022-04-15", "2022-04-17", "2022-04-21", "2022-05-01", "2022-06-03", "2022-09-07", "2022-10-12", "2022-11-02", "2022-11-15", "2022-11-20", "2022-12-25",
    "2023-01-01", "2023-02-21", "2023-04-07", "2023-04-09", "2023-04-21", "2023-05-01", "2023-06-08", "2023-09-07", "2023-10-12", "2023-11-02", "2023-11-15", "2023-11-20", "2023-12-25",
    "2024-01-01", "2024-03-29", "2024-04-21", "2024-05-01", "2024-05-30", "2024-09-07", "2024-10-12", "2024-11-02", "2024-11-15", "2024-11-20", "2024-12-25",
    "2025-01-01", "2025-04-18", "2025-04-21", "2025-05-01", "2025-06-19", "2025-09-07", "2025-10-12", "2025-11-02", "2025-11-15", "2025-11-20", "2025-12-25",
    "2026-01-01", "2026-02-16", "2026-02-17", "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04", "2026-09-07", "2026-10-12", "2026-11-02", "2026-11-15", "2026-11-20", "2026-12-25"
]);

let prodRawData = [];
let prodProcessedData = [];

// Função idêntica ao DIATRABALHOTOTAL.INTL
function calcularDiasUteisProd(startIso, endIso) {
    if (!startIso || !endIso) return null;
    let s = new Date(startIso + "T00:00:00");
    let e = new Date(endIso + "T00:00:00");
    
    // Se a data de abertura for maior que o fechamento
    let isNegative = false;
    if (s > e) { let temp = s; s = e; e = temp; isNegative = true; }
    
    let days = 0;
    let cur = new Date(s);
    while(cur <= e) {
        let dayOfWeek = cur.getDay();
        let iso = cur.toISOString().split('T')[0];
        if (dayOfWeek !== 0 && dayOfWeek !== 6 && !feriadosProdutividade.has(iso)) {
            days++;
        }
        cur.setDate(cur.getDate() + 1);
    }
    return isNegative ? -(days) : days;
}

function handleProdutividadeFile(event) {
    const file = event.target.files[0]; if (!file) return;
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.getElementById('loadingText').innerText = "Processando lógicas de SLA e Produtividade...";

    Papa.parse(file, {
        header: true, skipEmptyLines: true, encoding: "ISO-8859-1",
        complete: function(results) {
            prodRawData = results.data;
            reprocessarDatasBase();
            document.getElementById('loadingOverlay').style.display = 'none';
            event.target.value = '';
        },
        error: function(err) {
            console.error(err);
            document.getElementById('loadingOverlay').style.display = 'none';
            alert("Erro ao ler o arquivo CSV.");
        }
    });
}

function reprocessarDatasBase() {
    if(!prodRawData.length) return;
    
    let dtBaseInput = document.getElementById('dataBaseCalculo').value;
    let dtBaseIso = dtBaseInput || new Date().toISOString().split('T')[0]; // Representa CA1 / HOJE()

    prodProcessedData = prodRawData.map(row => {
        let newRow = { ...row };
        
        // Mapeamento de Colunas Padrão do Sistema
        let dtAberturaRaw = row['AB'] || row['Data Abertura'] || row['M2'] || "";
        let dtFechamentoRaw = row['FC'] || row['Data Fechamento'] || row['O2'] || "";
        let responsabilidadeRaw = row['Responsabilidade KOF'] || row['AM2'] || row['Responsável'] || "";
        
        let isoAbertura = parseDateToIso(dtAberturaRaw);
        let isoFechamento = parseDateToIso(dtFechamentoRaw);

        // --- 1. Dias Trabalhados (BT2) e SLA ---
        let diasTrab = "Aberto";
        if (isoFechamento) {
            let d = calcularDiasUteisProd(isoAbertura, isoFechamento);
            diasTrab = d !== null ? d - 1 : "Aberto";
        }
        newRow['Dias Trabalhados'] = diasTrab;

        let sla = "Aberto";
        if (diasTrab !== "Aberto") {
            if (diasTrab <= 2) sla = "D+1";
            else if (diasTrab === 3) sla = "D+2";
            else if (diasTrab === 4) sla = "D+3";
            else if (diasTrab > 4) sla = ">D+4";
        }
        newRow['SLA'] = sla;

        // --- 2. Saldo e Idade (Até Hoje / Data Base CA1) ---
        let diasAbertoHoje = calcularDiasUteisProd(isoAbertura, new Date().toISOString().split('T')[0]) - 1;
        newRow['Dias em Aberto (Hoje)'] = diasAbertoHoje;

        // Dias Abertos Ajust
        let faixaIdade = "";
        if (diasAbertoHoje <= 3) faixaIdade = "Entre 0 a 3 dias";
        else if (diasAbertoHoje <= 7) faixaIdade = "Entre 4 a 7 dias";
        else if (diasAbertoHoje <= 14) faixaIdade = "Entre 8 a 14 dias";
        else if (diasAbertoHoje <= 30) faixaIdade = "Entre 15 a 30 dias";
        else faixaIdade = "Acima de 31 dias";
        newRow['Idade Saldo Ajust (Hoje)'] = faixaIdade;

        // N e m (Idade baseada na Célula CA1)
        let diasAbertoBase = isoFechamento && isoFechamento <= dtBaseIso ? "Finalizado" : calcularDiasUteisProd(isoAbertura, dtBaseIso) - 1;
        newRow['Idade (Ref. Data Base)'] = diasAbertoBase;
        
        if(diasAbertoBase === "Finalizado") {
            newRow['Faixa Saldo (Ref. Data Base)'] = "Fechado na Data";
        } else {
            if (diasAbertoBase <= 3) newRow['Faixa Saldo (Ref. Data Base)'] = "Entre 0 a 3 dias";
            else if (diasAbertoBase <= 7) newRow['Faixa Saldo (Ref. Data Base)'] = "Entre 4 a 7 dias";
            else if (diasAbertoBase <= 14) newRow['Faixa Saldo (Ref. Data Base)'] = "Entre 8 a 14 dias";
            else if (diasAbertoBase <= 30) newRow['Faixa Saldo (Ref. Data Base)'] = "Entre 15 a 30 dias";
            else newRow['Faixa Saldo (Ref. Data Base)'] = "Acima de 31 dias";
        }

        // --- 3. Status OS e Responsabilidade ---
        newRow['Status OS'] = isoFechamento ? "Finalizada" : "Aberta";
        
        // Lógica Simplificada de Responsabilidade
        let respUpper = responsabilidadeRaw.toUpperCase();
        if(respUpper === "" || respUpper.includes("IMBERA")) newRow['Responsabilidade Final'] = "RESPONSABILIDADE IMBERA";
        else newRow['Responsabilidade Final'] = "RESPONSABILIDADE KOF";

        return newRow;
    });

    renderProdPreview();
}

function renderProdPreview() {
    const previewSection = document.getElementById('prodPreviewSection');
    const thead = document.getElementById('prodPreviewHead');
    const tbody = document.getElementById('prodPreviewBody');
    
    if (!prodProcessedData || prodProcessedData.length === 0) return;
    
    // Focando nas colunas mais importantes geradas
    const chavesImportantes = ["Data Abertura", "Data Fechamento", "Dias Trabalhados", "SLA", "Dias em Aberto (Hoje)", "Idade Saldo Ajust (Hoje)", "Faixa Saldo (Ref. Data Base)", "Responsabilidade Final"];
    const colunas = Object.keys(prodProcessedData[0]).filter(k => chavesImportantes.includes(k) || k === 'OS' || k === 'Chamado');

    let headHtml = '<tr class="bg-slate-100/50 text-slate-600">';
    colunas.forEach(k => { headHtml += `<th class="px-4 py-3 border-b border-slate-200 whitespace-nowrap font-bold">${k}</th>`; });
    headHtml += '</tr>';
    thead.innerHTML = headHtml;

    let bodyHtml = '';
    prodProcessedData.slice(0, 50).forEach(row => {
        bodyHtml += '<tr class="hover:bg-purple-50 transition-colors">';
        colunas.forEach(k => { 
            let v = row[k] || ''; 
            // Colorir SLA
            let corTexto = "text-slate-600";
            if(v === "D+1" || v === "D+2") corTexto = "text-emerald-600 font-bold";
            if(v === "D+3") corTexto = "text-yellow-600 font-bold";
            if(v === ">D+4" || v === "ACIMA D+3") corTexto = "text-rose-600 font-black";

            bodyHtml += `<td class="px-4 py-2 border-b border-slate-50 truncate max-w-[150px] ${corTexto}">${v}</td>`; 
        });
        bodyHtml += '</tr>';
    });
    
    if(prodProcessedData.length > 50) {
        bodyHtml += `<tr><td colspan="${colunas.length}" class="px-4 py-3 text-center text-[10px] text-slate-400 italic bg-slate-50/50">+ ${prodProcessedData.length - 50} linhas processadas e ocultas nesta prévia</td></tr>`;
    }

    tbody.innerHTML = bodyHtml;
    previewSection.classList.remove('hidden');
    lucide.createIcons();
}

function downloadProdCSV() {
    if (!prodProcessedData || prodProcessedData.length === 0) return;
    const csv = Papa.unparse(prodProcessedData, { delimiter: ";", quotes: true }); 
    const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' }); 
    const url = URL.createObjectURL(blob); 
    const link = document.createElement("a"); 
    link.setAttribute("href", url); 
    link.setAttribute("download", "base_produtividade_calculada.csv"); 
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}