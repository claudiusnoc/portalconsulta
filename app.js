/* ===================================================
   CLARO INFRA MG - BI DASHBOARD
   EQS Engenharia | Controle de Compras 2026
=================================================== */

// Global State
const PEDIDOS = [];
let filteredData = [];
let charts = {};

// Configurações de Estilo (Cores do Tema Roxo)
const CHART_COLORS = {
  purple: '#C0392B',
  purpleLight: 'rgba(192, 57, 43, 0.08)',
  blue: '#3a7cc0',
  green: '#2d8a5e',
  yellow: '#c5920a',
  orange: '#c97a2a',
  red: '#C0392B',
  text: '#1a1714',
  muted: '#5c554d',
  border: 'rgba(0,0,0,0.04)'
};

const CHART_DEFAULTS = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: {
      display: false
    },
    tooltip: {
      backgroundColor: 'rgba(255, 255, 255, 0.9)',
      titleColor: '#1f2937',
      bodyColor: '#6b7280',
      borderColor: 'rgba(0,0,0,0.05)',
      borderWidth: 1,
      padding: 12,
      displayColors: false,
      callbacks: {
        label: (context) => ` ${context.raw} pedidos`
      }
    }
  },
  scales: {
    x: { grid: { display: false }, ticks: { color: '#6b7280', font: { family: 'Inter', size: 11 } } },
    y: { grid: { color: 'rgba(0,0,0,0.03)', drawBorder: false }, ticks: { color: '#6b7280', font: { family: 'Inter', size: 11 } } }
  }
};

const STATUS_CONFIG = {
  'FINALIZADO': { cls: 'badge-finalizado', label: 'Finalizado' },
  'EM PROCESSO/PARCIAL': { cls: 'badge-parcial', label: 'Em Processo/Parcial' },
  'AGUARD. ENTREGA/COLETA': { cls: 'badge-aguardando', label: 'Aguard. Entrega/Coleta' },
  'EM ANALISE COMPRAS': { cls: 'badge-analise', label: 'Em Análise Compras' },
  'ENVIAR REPARO AGST': { cls: 'badge-reparo', label: 'Enviar Reparo AGST' },
  'ELIMINADO POR RESIDUO': { cls: 'badge-eliminado', label: 'Eliminado por Resíduo' },
  'EM ANDAMENTO': { cls: 'badge-parcial', label: 'Em Andamento' },
  'HUGO ALEXANDRE': { cls: 'badge-analise', label: 'Análise Hugo' },
  'ELOI JOSE': { cls: 'badge-analise', label: 'Análise Eloi' },
  'REJEITADO': { cls: 'badge-eliminado', label: 'Rejeitado' },
};

const APROV_CONFIG = {
  'APROVADO': { cls: 'badge-aprovado', label: 'Aprovado' },
  'HUGO ALEXANDRE': { cls: 'badge-hugo', label: 'Hugo Alexandre' },
  'ELOI JOSE': { cls: 'badge-hugo', label: 'Eloi Jose' },
  'FINALIZADO': { cls: 'badge-finalizado', label: 'Finalizado' },
  'REJEITADO': { cls: 'badge-eliminado', label: 'Rejeitado' },
  'PENDENTE': { cls: 'badge-pendente', label: 'Pendente' },
  '': { cls: 'badge-pendente', label: 'Pendente' },
};

// ===================================================
// INITIALIZATION
// ===================================================
document.addEventListener('DOMContentLoaded', () => {
  setCurrentDate();
  initTheme();
  fetchFromGoogleSheets();
});

function setCurrentDate() {
  const now = new Date();
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  const dateStr = now.toLocaleDateString('pt-BR', options);
  document.getElementById('currentDate').textContent = dateStr.charAt(0).toUpperCase() + dateStr.slice(1);
}

// ===================================================
// DARK MODE
// ===================================================
function initTheme() {
  const saved = localStorage.getItem('eqs-bi-theme');
  if (saved === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
    const toggle = document.getElementById('darkModeToggle');
    if (toggle) toggle.checked = true;
  }
}

function toggleDarkMode() {
  const toggle = document.getElementById('darkModeToggle');
  const isDark = toggle.checked;

  if (isDark) {
    document.documentElement.setAttribute('data-theme', 'dark');
    localStorage.setItem('eqs-bi-theme', 'dark');
  } else {
    document.documentElement.removeAttribute('data-theme');
    localStorage.setItem('eqs-bi-theme', 'light');
  }

  // Re-renderizar ícones Lucide após mudança de tema
  if (typeof lucide !== 'undefined') lucide.createIcons();
}

// ===================================================
// HELPER: CONVERSOR UNIVERSAL DE DATAS
// ===================================================
// Converte qualquer formato de data recebido do Google Sheets:
//   - Serial numérico do Excel: 46111 → 30/03/2026
//   - Formato ISO:              2026-03-30 → 30/03/2026
//   - Formato já correto:       30/03/2026 → 30/03/2026 (não muda)
function parseSheetDate(val) {
  if (!val || val === '') return '';
  const str = String(val).trim();

  // CASO 1: Serial numérico do Excel (ex: 46111)
  // O Excel conta dias desde 1 de janeiro de 1900
  if (/^\d{4,6}$/.test(str)) {
    const serial = parseInt(str, 10);
    // Apenas converte se for um número plausível de data (após 2000)
    if (serial > 36526 && serial < 60000) {
      // Base: 1 de janeiro de 1900 = serial 1
      const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      const d = date.getUTCDate().toString().padStart(2, '0');
      const m = (date.getUTCMonth() + 1).toString().padStart(2, '0');
      const y = date.getUTCFullYear();
      return `${d}/${m}/${y}`;
    }
  }

  // CASO 2: Formato ISO (ex: 2026-03-30 ou 2026-03-30T00:00:00)
  if (str.includes('-') && !str.includes('/')) {
    const parts = str.split('T')[0].split('-'); // remove parte de hora se houver
    if (parts.length === 3) {
      return `${parts[2].padStart(2,'0')}/${parts[1].padStart(2,'0')}/${parts[0]}`;
    }
  }

  // CASO 3: Já está no formato correto dd/mm/aaaa
  return str;
}

// ===================================================
// GOOGLE SHEETS API SYNC
// ===================================================
// CONFIGURAÇÃO DA PLANILHA (sem API Key — usa export público CSV)
// A planilha deve estar compartilhada como "Qualquer pessoa com o link pode ver"
const G_SHEETS_CONFIG = {
  sheetId: '1NCznd6Rwmopf14n5VpwHpGRQcY8e4PIh9rRNCmLtk4g'
};

// Parser CSV robusto (lida com campos entre aspas e vírgulas dentro de células)
function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  return lines.map(line => {
    const cells = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
        else { inQuotes = !inQuotes; }
      } else if (ch === ',' && !inQuotes) {
        cells.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    cells.push(current.trim());
    return cells;
  }).filter(row => row.some(cell => cell !== ''));
}

async function fetchFromGoogleSheets() {
  const statusEl = document.getElementById('spStatus');
  const btnSync = document.getElementById('btnSyncHeader');
  const syncBadge = document.getElementById('syncBadge');

  if (!G_SHEETS_CONFIG.sheetId || G_SHEETS_CONFIG.sheetId.includes('COLE_AQUI')) {
    if (statusEl) {
      statusEl.className = 'sp-status error';
      statusEl.textContent = '⚠️ Sheet ID não configurado no app.js';
    }
    updateDashboard(PEDIDOS);
    return;
  }

  try {
    if (statusEl) {
      statusEl.className = 'sp-status loading';
      statusEl.textContent = '⏳ Lendo Google Sheets...';
    }
    if (btnSync) btnSync.innerHTML = '<i data-lucide="loader" class="spin" style="width:16px;height:16px"></i> Atualizando...';

    // Export público CSV — não requer API Key nem autenticação
    const url = `https://docs.google.com/spreadsheets/d/${G_SHEETS_CONFIG.sheetId}/export?format=csv&t=${Date.now()}`;

    let response;
    try {
      response = await fetch(url);
    } catch (networkErr) {
      // fetch() rejeita com TypeError quando não há conexão
      throw new Error('SEM_INTERNET');
    }
    if (response.status === 403 || response.status === 401) {
      throw new Error('PLANILHA_PRIVADA');
    }
    if (response.status === 404) {
      throw new Error('PLANILHA_NAO_ENCONTRADA');
    }
    if (!response.ok) {
      throw new Error(`HTTP_${response.status}`);
    }

    const csvText = await response.text();
    const rows = parseCSV(csvText);

    if (!rows || rows.length < 2) {
      if (statusEl) {
        statusEl.className = 'sp-status';
        statusEl.textContent = '📊 Planilha conectada, mas vazia.';
      }
      updateDashboard(PEDIDOS);
      return;
    }

    // Procura a linha que contém os cabeçalhos reais (procurando pela coluna SM ou STATUS)
    let headerRowIndex = 0;
    for (let j = 0; j < Math.min(5, rows.length); j++) {
        if (rows[j].some(val => typeof val === 'string' && (val.toUpperCase() === 'SM' || val.toUpperCase() === 'STATUS'))) {
            headerRowIndex = j;
            break;
        }
    }

    const headers = rows[headerRowIndex].map(h => h.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\.ºº\s]/g, "").trim());
    
    const consolidatedData = [];
    
    // Varre as linhas de dados (começando logo após o cabeçalho)
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue; // Pula linha vazia

      // Mapeia os valores usando o índice dos cabeçalhos
      const rowObj = {};
      headers.forEach((h, index) => {
        rowObj[h] = row[index] || '';
      });

      const smVal = rowObj['SM'] || rowObj['NSM'] || rowObj['NUMEROSM'] || '';
      const descricaoVal = rowObj['DESCRICAO'] || '';

      // ============================================================
      // FILTRO RIGOROSO: Só processa linhas com SM numérica válida
      // (ignora linhas de legenda, títulos e linhas sem SM real).
      // SM válida = string que contém apenas dígitos, ex: "325995"
      // ============================================================
      const smString = String(smVal).trim();
      const isValidSM = smString.length >= 4 && /^\d+$/.test(smString);
      if (!isValidSM) continue;

      let statusVal = (rowObj['STATUS'] || '').toString().toUpperCase().trim();
      let aprovacaoVal = (rowObj['STATUSAPROVACAO'] || rowObj['APROVACAO'] || '').toString().toUpperCase().trim();

      // Correção: Itens finalizados, rejeitados ou eliminados NÃO podem ter aprovação 'Pendente'
      if (['FINALIZADO', 'ELIMINADO POR RESIDUO', 'REJEITADO'].includes(statusVal)) {
        if (aprovacaoVal === '' || aprovacaoVal === 'PENDENTE') {
          aprovacaoVal = 'FINALIZADO';
        }
      }

      // Converte ambas as datas usando o conversor universal
      const dataEmissao = parseSheetDate(rowObj['EMISSAO'] || rowObj['DATAEMISSAO'] || rowObj['DATA'] || '');
      const dataAtualizacao = parseSheetDate(rowObj['DATAATUALIZACAO'] || rowObj['ATUALIZACAO'] || '');

      consolidatedData.push({
        sm: smString,
        pedido: rowObj['PEDIDO'] || '',
        emissao: dataEmissao,
        status: statusVal,
        aprovacao: aprovacaoVal,
        descricao: descricaoVal,
        atualizacao: dataAtualizacao,
        fonte: 'Google Sheets'
      });
    }

    // ====================================================
    // DEDUPLICAÇÃO por SM: A ÚLTIMA ocorrência de cada SM
    // na planilha vence (a mais recente enviada pelo robô).
    // Como todas as linhas têm SM válida, nunca haverá
    // chaves aleatórias que causam duplicação.
    // ====================================================
    const smMap = new Map();
    consolidatedData.forEach(p => {
      smMap.set(p.sm, p);  // Última entrada da mesma SM sobrescreve a anterior
    });
    const deduplicated = Array.from(smMap.values());

    PEDIDOS.length = 0;
    PEDIDOS.push(...deduplicated);

    updateDashboard(PEDIDOS);

    // Atualiza badges da UI
    if (syncBadge) {
      syncBadge.style.display = 'flex';
      const now = new Date();
      document.getElementById('syncTime').textContent = `Sincronizado ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    }

    if (statusEl) {
      statusEl.className = 'sp-status success';
      statusEl.textContent = `✅ Banco de Dados conectado!`;
    }

  } catch (err) {
    console.error('Google Sheets Sync Error:', err);
    if (statusEl) {
      statusEl.className = 'sp-status error';
      const msgs = {
        'SEM_INTERNET':           '❌ Sem conexão. Verifique sua internet.',
        'PLANILHA_PRIVADA':       '❌ Planilha privada. Compartilhe como "qualquer pessoa pode ver".',
        'PLANILHA_NAO_ENCONTRADA':'❌ Planilha não encontrada. Verifique o ID no app.js.',
      };
      statusEl.textContent = msgs[err.message] || `❌ Erro ao sincronizar (${err.message}).`;
    }
    updateDashboard(PEDIDOS);
  } finally {
    if (btnSync) btnSync.innerHTML = '<i data-lucide="refresh-cw" style="width:16px;height:16px"></i> Sincronizar';
    if (typeof lucide !== 'undefined') lucide.createIcons();
  }
}

// ===================================================
// DASHBOARD LOGIC (FILTERING & RENDERING)
// ===================================================
function updateDashboard(data) {
  const now = new Date();
  const curMonth = (now.getMonth() + 1).toString().padStart(2, '0');
  const curYear = now.getFullYear().toString();

  // Filtrar apenas o mês atual para o Dashboard (Home)
  const monthData = data.filter(p => {
    if (!p.emissao || !p.emissao.includes('/')) return false;
    const parts = p.emissao.split('/');
    return parts[1] === curMonth && (parts[2] === curYear || parts[2] === curYear.slice(-2));
  });

  const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  const hasCurrentMonth = monthData.length > 0;

  // Sem dados no mês atual → exibe zeros (não mostra histórico acumulado)
  const displayData = hasCurrentMonth ? monthData : [];

  document.getElementById('dashboard-subtitle').textContent = hasCurrentMonth
    ? `${monthNames[now.getMonth()]} ${curYear} · BH Metro ADM · EQS Engenharia`
    : `Nenhum pedido em ${monthNames[now.getMonth()]} ainda · BH Metro ADM · EQS Engenharia`;

  renderKPIs(displayData);
  renderCharts(displayData);
  renderPreviewTable(displayData);

  // Define o filtro do dropdown para o mês atual
  document.getElementById('filterMes').value = curMonth;
  applyFilters();
}

function renderKPIs(data) {
  const total = data.length;
  const finalizado = data.filter(p => p.status === 'FINALIZADO' || p.aprovacao === 'FINALIZADO').length;
  const andamento = data.filter(p => p.status === 'EM PROCESSO/PARCIAL' || p.status === 'EM ANDAMENTO').length;
  const aguardando = data.filter(p => p.status === 'AGUARD. ENTREGA/COLETA').length;
  
  const analise = data.filter(p =>
    ['EM ANALISE COMPRAS', 'HUGO ALEXANDRE', 'ELOI JOSE', 'ENVIAR REPARO AGST'].includes(p.status) ||
    ['HUGO ALEXANDRE', 'ELOI JOSE', 'PENDENTE', ''].includes(p.aprovacao)
  ).length;
  const aprovado = data.filter(p => p.aprovacao === 'APROVADO').length;
  
  // Rejeitado só conta REJEITADO (desconectado do Eliminado por Resíduo de acordo com feedback do usuário)
  const rejeitado = data.filter(p => p.status === 'REJEITADO').length;

  animateValue('kpi-total', total);
  animateValue('kpi-finalizado', finalizado);
  animateValue('kpi-andamento', andamento);
  animateValue('kpi-aguardando', aguardando);
  animateValue('kpi-analise', analise);
  animateValue('kpi-aprovado', aprovado);
  animateValue('kpi-rejeitado', rejeitado);

  // Update Progress Bars
  if (total > 0) {
    document.getElementById('bar-finalizado').style.width = `${(finalizado / total) * 100}%`;
    document.getElementById('bar-andamento').style.width = `${(andamento / total) * 100}%`;
    const barAg = document.getElementById('bar-aguardando');
    if (barAg) barAg.style.width = `${(aguardando / total) * 100}%`;
    document.getElementById('bar-analise').style.width = `${(analise / total) * 100}%`;
    document.getElementById('bar-aprovado').style.width = `${(aprovado / total) * 100}%`;
    const barRej = document.getElementById('bar-rejeitado');
    if (barRej) barRej.style.width = `${(rejeitado / total) * 100}%`;
  }
}

function animateValue(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  const start = parseInt(el.textContent) || 0;
  const duration = 1000;
  let startTime = null;

  function step(timestamp) {
    if (!startTime) startTime = timestamp;
    const progress = Math.min((timestamp - startTime) / duration, 1);
    el.textContent = Math.floor(progress * (value - start) + start);
    if (progress < 1) window.requestAnimationFrame(step);
  }
  window.requestAnimationFrame(step);
}

// ===================================================
// CHARTS (Chart.js)
// ===================================================
function renderCharts(data) {
  renderStatusChart(data);
  renderAprovChart(data);
  renderMonthlyChart(PEDIDOS); // Gráfico mensal sempre mostra o histórico completo
}

function renderStatusChart(data) {
  const ctx = document.getElementById('chartStatus');
  if (!ctx) return;

  const statusGroups = {};
  data.forEach(p => {
    const label = (STATUS_CONFIG[p.status] || { label: p.status || 'Outro' }).label;
    statusGroups[label] = (statusGroups[label] || 0) + 1;
  });

  if (charts.status) {
    charts.status.data.labels = Object.keys(statusGroups);
    charts.status.data.datasets[0].data = Object.values(statusGroups);
    charts.status.update('none');
  } else {
    charts.status = new Chart(ctx.getContext('2d'), {
      type: 'bar',
      data: {
        labels: Object.keys(statusGroups),
        datasets: [{
          label: 'Pedidos',
          data: Object.values(statusGroups),
          backgroundColor: CHART_COLORS.purple,
          borderRadius: 4,
          borderSkipped: false,
          barThickness: 24
        }]
      },
      options: { ...CHART_DEFAULTS, indexAxis: 'y' }
    });
  }
}

function renderAprovChart(data) {
  const ctx = document.getElementById('chartAprov');
  if (!ctx) return;

  const hugoCount = data.filter(p => p.aprovacao === 'HUGO ALEXANDRE').length;
  const eloiCount = data.filter(p => p.aprovacao === 'ELOI JOSE').length;
  // Agrupa os aprovados totais + os finalizados automáticos
  const aprovadoCount = data.filter(p => p.aprovacao === 'APROVADO' || p.aprovacao === 'FINALIZADO').length;

  if (charts.aprov) {
    charts.aprov.data.datasets[0].data = [hugoCount, eloiCount, aprovadoCount];
    charts.aprov.update('none');
  } else {
    charts.aprov = new Chart(ctx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Hugo Alexandre', 'Eloi Jose', 'Aprovados'],
        datasets: [{
          data: [hugoCount, eloiCount, aprovadoCount],
          backgroundColor: [CHART_COLORS.purple, CHART_COLORS.blue, CHART_COLORS.green],
          borderWidth: 0,
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '75%',
        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {
              boxWidth: 10,
              padding: 15,
              font: { family: 'Inter', size: 11, weight: 600 },
              color: CHART_COLORS.muted
            }
          }
        }
      }
    });
  }
}

function renderMonthlyChart(data) {
  const ctx = document.getElementById('chartMes');
  if (!ctx) return;

  const monthCounts = {};
  const monthLabels = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

  // Inicializa meses com zero
  monthLabels.forEach(m => monthCounts[m] = 0);

  const curYear = new Date().getFullYear().toString();
  const curYearShort = curYear.slice(-2);

  data.forEach(p => {
    if (!p.emissao || !p.emissao.includes('/')) return;
    const parts = p.emissao.split('/');
    
    // Ignora dados de outros anos para o gráfico mensal do ano corrente não somar 2024, 2025, 2026 juntos
    if (parts[2] !== curYear && parts[2] !== curYearShort) return;

    const monthIdx = parseInt(parts[1], 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12) {
      monthCounts[monthLabels[monthIdx]]++;
    }
  });

  if (charts.mes) {
    charts.mes.data.datasets[0].data = Object.values(monthCounts);
    charts.mes.update('none');
  } else {
    charts.mes = new Chart(ctx.getContext('2d'), {
      type: 'line',
      data: {
        labels: Object.keys(monthCounts),
        datasets: [{
          label: 'Volume',
          data: Object.values(monthCounts),
          borderColor: CHART_COLORS.purple,
          backgroundColor: (context) => {
            const chart = context.chart;
            const { ctx, chartArea } = chart;
            if (!chartArea) return null;
            const gradient = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
            gradient.addColorStop(0, 'rgba(192, 57, 43, 0)');
            gradient.addColorStop(1, 'rgba(192, 57, 43, 0.12)');
            return gradient;
          },
          fill: true,
          tension: 0.4,
          pointRadius: 0,
          pointHitRadius: 10,
          borderWidth: 3
        }]
      },
      options: {
        ...CHART_DEFAULTS,
        interaction: { intersect: false, mode: 'index' }
      }
    });
  }
}

// ===================================================
// TABLES
// ===================================================
function buildRow(p) {
  const statusCfg = STATUS_CONFIG[p.status] || { cls: 'badge-pendente', label: p.status || 'Pendente' };
  const aprovCfg = APROV_CONFIG[p.aprovacao] || APROV_CONFIG[''];

  return `<tr>
    <td><strong>${p.sm || '-'}</strong></td>
    <td style="font-family:monospace; color:var(--primary); font-weight:700">${p.pedido || '-'}</td>
    <td>${p.emissao || '-'}</td>
    <td><span class="badge ${statusCfg.cls}">${statusCfg.label}</span></td>
    <td><span class="badge ${aprovCfg.cls}">${aprovCfg.label}</span></td>
    <td>${p.descricao}</td>
    <td style="font-size:11px;color:var(--text-faint)">${p.fonte || '-'}</td>
    <td style="font-size:11px;color:var(--text-muted)">${p.atualizacao || '-'}</td>
  </tr>`;
}

function renderPreviewTable(data) {
  const latest = [...data].reverse().slice(0, 8);
  document.getElementById('tbody-preview').innerHTML = latest.map(buildRow).join('');
}

function renderPedidosTable(data) {
  const tbody = document.getElementById('tbody-pedidos');
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:48px;color:var(--text-muted)">Nenhum pedido encontrado.</td></tr>`;
    document.getElementById('filterCount').textContent = '0 pedidos';
    return;
  }
  tbody.innerHTML = data.map(buildRow).join('');
  document.getElementById('filterCount').textContent = `${data.length} pedidos encontrados`;
}

// ===================================================
// FILTERS & NAVIGATION
// ===================================================
function applyFilters() {
  const searchInput = document.getElementById('globalSearch');
  const search = (searchInput?.value || '').toLowerCase().trim();
  const status = document.getElementById('filterStatus')?.value || '';
  const aprov = document.getElementById('filterAprov')?.value || '';
  const mes = document.getElementById('filterMes')?.value || '';

  filteredData = PEDIDOS.filter(p => {
    // Busca abrangente
    const matchSearch = !search || [p.sm, p.pedido, p.descricao, p.status, p.aprovacao].some(f =>
      (f || '').toString().toLowerCase().includes(search)
    );
    const matchDirect = !search || (p.pedido || '').toString().toLowerCase() === search || (p.sm || '').toString().toLowerCase() === search;

    // Lógica inteligente de Grupos para vincular o Clique do KPI com o filtro
    let matchStatus = false;
    if (!status) {
      matchStatus = true;
    } else if (status === 'GRP_ANALISE') {
       matchStatus = ['EM ANALISE COMPRAS', 'HUGO ALEXANDRE', 'ELOI JOSE', 'ENVIAR REPARO AGST'].includes(p.status) || 
                     ['HUGO ALEXANDRE', 'ELOI JOSE', 'PENDENTE', ''].includes(p.aprovacao);
    } else if (status === 'GRP_ANDAMENTO') {
       matchStatus = ['EM PROCESSO/PARCIAL', 'EM ANDAMENTO'].includes(p.status);
    } else if (status === 'GRP_FINALIZADO') {
       matchStatus = ['FINALIZADO'].includes(p.status) || p.aprovacao === 'FINALIZADO';
    } else {
       matchStatus = p.status === status;
    }

    const matchAprov = !aprov || p.aprovacao === aprov;

    // Filtro de mês: normaliza o ano da data para 4 dígitos antes de comparar
    const curYear = new Date().getFullYear().toString(); // ex: "2026"
    const matchMes = !mes || (() => {
      if (!p.emissao || !p.emissao.includes('/')) return false;
      const parts = p.emissao.split('/');
      if (parts.length < 3) return false;
      const monthPart = parts[1];
      const rawYear   = parts[2].trim();
      // Normaliza: "26" → "2026", "2026" → "2026"
      const yearFull  = rawYear.length === 2 ? '20' + rawYear : rawYear;
      return monthPart === mes && yearFull === curYear;
    })();

    return (matchSearch || matchDirect) && matchStatus && matchAprov && matchMes;
  });

  renderPedidosTable(filteredData);

  // Auto-navegar para a aba de pedidos se o usuário digitar na busca global e estiver focada
  if (search.length > 0 && document.activeElement === searchInput) {
    if (!document.getElementById('view-pedidos').classList.contains('active')) {
      showView('pedidos');
    }
  }
}

function filterByKPI(type) {
  // Reset all filters EXCEPT the Month filter
  // The KPI counts reflect the currently selected month, so clicking them
  // should keep the month filter intact to match the count.
  document.getElementById('globalSearch').value = "";
  document.getElementById('filterStatus').value = "";
  document.getElementById('filterAprov').value = "";

  // 'total' não aplica filtro extra — mostra todos os pedidos do mês
  if (type === 'finalizado') document.getElementById('filterStatus').value = "GRP_FINALIZADO";
  if (type === 'andamento')  document.getElementById('filterStatus').value = "GRP_ANDAMENTO";
  if (type === 'aguardando') document.getElementById('filterStatus').value = "AGUARD. ENTREGA/COLETA";
  if (type === 'analise')    document.getElementById('filterStatus').value = "GRP_ANALISE";
  if (type === 'aprovado')   document.getElementById('filterAprov').value  = "APROVADO";
  if (type === 'rejeitado')  document.getElementById('filterStatus').value = "REJEITADO";

  applyFilters();
  showView('pedidos');
}

function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(`view-${id}`).classList.add('active');

  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById(`nav-${id}`).classList.add('active');

  // Ao voltar para o Dashboard limpa a busca global para evitar estado residual
  if (id === 'dashboard') {
    const search = document.getElementById('globalSearch');
    if (search) { search.value = ''; applyFilters(); }
  }

  if (id === 'status' && !charts.statusFull) renderStatusFullChart();
}

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const main = document.querySelector('.main-content');
  sidebar.classList.toggle('collapsed');
  main.classList.toggle('full');
}

function renderStatusFullChart() {
  const ctx = document.getElementById('chartStatusFull');
  if (!ctx) return;

  const statusGroups = {};
  PEDIDOS.forEach(p => {
    const label = (STATUS_CONFIG[p.status] || { label: p.status || 'Outro' }).label;
    statusGroups[label] = (statusGroups[label] || 0) + 1;
  });

  if (charts.statusFull) {
    charts.statusFull.data.labels = Object.keys(statusGroups);
    charts.statusFull.data.datasets[0].data = Object.values(statusGroups);
    charts.statusFull.update('none');
    return;
  }
  charts.statusFull = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: Object.keys(statusGroups),
      datasets: [{
        label: 'Total Acumulado',
        data: Object.values(statusGroups),
        backgroundColor: CHART_COLORS.purple,
        borderRadius: 8
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y'
    }
  });
}
