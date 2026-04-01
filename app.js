/* ===================================================
   CLARO INFRA MG - BI DASHBOARD
   EQS Engenharia | Controle de Compras 2026
=================================================== */

// Global State
const PEDIDOS = [];
let filteredData = [];
let charts = {};
let sortState = { col: null, dir: 'asc' };
let pageState = { current: 1, pageSize: 50 };

// Debounce utilitário — evita reconstrução da tabela a cada tecla
let _searchTimer = null;
function debouncedSearch() {
  clearTimeout(_searchTimer);
  _searchTimer = setTimeout(applyFilters, 200);
}

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
  'ELOI JOSE': { cls: 'badge-eloi', label: 'Eloi José' },
  'ELOI JOSÉ': { cls: 'badge-eloi', label: 'Eloi José' },  // alias com acento
  'ADILSON RODRIGUES': { cls: 'badge-adilson', label: 'Adilson Rodrigues' },
  'MARCOS CARIAS': { cls: 'badge-marcos', label: 'Marcos Carias' },
  'FINALIZADO': { cls: 'badge-finalizado', label: 'Finalizado' },
  'REJEITADO': { cls: 'badge-eliminado', label: 'Rejeitado' },
  'PENDENTE': { cls: 'badge-pendente', label: 'Pendente' },
  '': { cls: 'badge-pendente', label: 'Pendente' },
};

// ===================================================
// INITIALIZATION
// ===================================================
// AUTO-SYNC (agendado: 10:00 e 16:00 — alinhado com Power Automate)
// ===================================================
const SYNC_SCHEDULE = [{ h: 10, m: 0 }, { h: 16, m: 0 }];
let autoSyncCountdownTimer = null;

function startAutoSync() {
  if (autoSyncCountdownTimer) clearInterval(autoSyncCountdownTimer);

  autoSyncCountdownTimer = setInterval(() => {
    const now = new Date();
    // Disparar quando exatamente no horário (janela de 1 segundo)
    const hh = now.getHours();
    const mm = now.getMinutes();
    const ss = now.getSeconds();

    if (ss === 0 && SYNC_SCHEDULE.some(t => t.h === hh && t.m === mm)) {
      fetchFromGoogleSheets();
    }
  }, 1000);
}

function updateNextSyncLabel() {
  const labelEl = document.getElementById('nextSyncLabel');
  if (!labelEl) return;

  const now = new Date();
  const todayBase = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const candidates = SYNC_SCHEDULE.map(t => {
    const dt = new Date(todayBase);
    dt.setHours(t.h, t.m, 0, 0);
    if (dt <= now) dt.setDate(dt.getDate() + 1); // já passou hoje → amanhã
    return dt;
  });

  const next = candidates.reduce((a, b) => a < b ? a : b);
  const timeStr = next.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  labelEl.textContent = `Próxima sincronização: ${timeStr}`;
}

// Atualiza o label a cada minuto
setInterval(updateNextSyncLabel, 60000);

document.addEventListener('DOMContentLoaded', () => {
  setCurrentDate();
  initTheme();
  initSortableHeaders();
  fetchFromGoogleSheets();
  updateNextSyncLabel();
  startAutoSync();
});

function initSortableHeaders() {
  document.querySelectorAll('#tablePedidos .th-sortable').forEach(th => {
    th.addEventListener('click', () => setSortCol(th.dataset.sort));
  });
}

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
      return `${parts[2].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[0]}`;
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
      const hh = now.getHours().toString().padStart(2, '0');
      const mm = now.getMinutes().toString().padStart(2, '0');
      document.getElementById('syncTime').textContent = `${deduplicated.length} pedidos · ${hh}:${mm}`;
    }

    if (statusEl) {
      statusEl.className = 'sp-status success';
      statusEl.textContent = `✅ ${deduplicated.length} registros carregados`;
    }

  } catch (err) {
    console.error('Google Sheets Sync Error:', err);
    if (statusEl) {
      statusEl.className = 'sp-status error';
      const msgs = {
        'SEM_INTERNET': '❌ Sem conexão. Verifique sua internet.',
        'PLANILHA_PRIVADA': '❌ Planilha privada. Compartilhe como "qualquer pessoa pode ver".',
        'PLANILHA_NAO_ENCONTRADA': '❌ Planilha não encontrada. Verifique o ID no app.js.',
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
  updateFilterBadge(hasCurrentMonth ? `${monthNames[now.getMonth()]} ${curYear}` : null);

  // Define o filtro do dropdown para o mês atual
  document.getElementById('filterMes').value = curMonth;
  applyFilters();
}

function updateFilterBadge(label) {
  const badge = document.getElementById('kpi-filter-badge');
  const text = document.getElementById('kpi-filter-text');
  if (!badge) return;
  if (label) {
    text.textContent = `Exibindo: ${label}`;
    badge.style.display = 'inline-flex';
  } else {
    badge.style.display = 'none';
  }
}

function clearMonthFilter() {
  document.getElementById('filterMes').value = '';
  document.getElementById('kpi-filter-badge').style.display = 'none';
  renderKPIs(PEDIDOS);
  renderCharts(PEDIDOS);
  renderPreviewTable(PEDIDOS);
  applyFilters();
}

function renderKPIs(data) {
  const total = data.length;
  const finalizado = data.filter(p => p.status === 'FINALIZADO' || p.aprovacao === 'FINALIZADO').length;
  const andamento = data.filter(p => p.status === 'EM PROCESSO/PARCIAL' || p.status === 'EM ANDAMENTO').length;
  const aguardando = data.filter(p => p.status === 'AGUARD. ENTREGA/COLETA').length;

  const APROVADORES_PENDENTES = ['HUGO ALEXANDRE', 'ELOI JOSE', 'ELOI JOSÉ', 'ADILSON RODRIGUES', 'MARCOS CARIAS', 'PENDENTE', ''];
  const analise = data.filter(p =>
    ['EM ANALISE COMPRAS', 'HUGO ALEXANDRE', 'ELOI JOSE', 'ENVIAR REPARO AGST'].includes(p.status) ||
    APROVADORES_PENDENTES.includes(p.aprovacao)
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
  const eloiCount = data.filter(p => p.aprovacao === 'ELOI JOSE' || p.aprovacao === 'ELOI JOSÉ').length;
  const adilsonCount = data.filter(p => p.aprovacao === 'ADILSON RODRIGUES').length;
  const marcosCount = data.filter(p => p.aprovacao === 'MARCOS CARIAS').length;
  // Aprovados totais (APROVADO explícito + FINALIZADO automático)
  const aprovadoCount = data.filter(p => p.aprovacao === 'APROVADO' || p.aprovacao === 'FINALIZADO').length;

  const chartData = [hugoCount, eloiCount, adilsonCount, marcosCount, aprovadoCount];
  const chartLabels = ['Hugo Alexandre', 'Eloi José', 'Adilson Rodrigues', 'Marcos Carias', 'Aprovados'];
  const chartColors = [CHART_COLORS.orange, CHART_COLORS.blue, '#2d8a7e', '#7c5cbf', CHART_COLORS.green];

  if (charts.aprov) {
    charts.aprov.data.labels = chartLabels;
    charts.aprov.data.datasets[0].data = chartData;
    charts.aprov.data.datasets[0].backgroundColor = chartColors;
    charts.aprov.update('none');
  } else {
    charts.aprov = new Chart(ctx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: chartLabels,
        datasets: [{
          data: chartData,
          backgroundColor: chartColors,
          borderWidth: 0,
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '72%',
        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {
              boxWidth: 10,
              padding: 12,
              font: { family: 'Inter', size: 10, weight: 600 },
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

  // Destaque visual no mês atual: ponto maior e vermelho EQS
  const curMonthIdx = new Date().getMonth(); // 0–11
  const pointRadii = monthLabels.map((_, i) => i === curMonthIdx ? 7 : 0);
  const pointColors = monthLabels.map((_, i) => i === curMonthIdx ? CHART_COLORS.purple : 'transparent');
  const pointBorders = monthLabels.map((_, i) => i === curMonthIdx ? '#fff' : 'transparent');

  if (charts.mes) {
    charts.mes.data.datasets[0].data = Object.values(monthCounts);
    charts.mes.data.datasets[0].pointRadius = pointRadii;
    charts.mes.data.datasets[0].pointBackgroundColor = pointColors;
    charts.mes.data.datasets[0].pointBorderColor = pointBorders;
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
          pointRadius: pointRadii,
          pointBackgroundColor: pointColors,
          pointBorderColor: pointBorders,
          pointBorderWidth: 2,
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

// Retorna dias desde a emissão (null se data inválida)
const DIAS_PARADO_ALERTA = 15; // dias sem atualização para acionar alerta

function calcDiasDesdeAtualizacao(atualizacao) {
  if (!atualizacao || !atualizacao.includes('/')) return null;
  const parts = atualizacao.split('/');
  if (parts.length < 3) return null;
  const y = parseInt(parts[2].length === 2 ? '20' + parts[2] : parts[2], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[0], 10);
  const dt = new Date(y, m, d);
  if (isNaN(dt.getTime())) return null;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return Math.floor((today - dt) / 86400000);
}

function calcDiasAberto(emissao) {
  if (!emissao || !emissao.includes('/')) return null;
  const parts = emissao.split('/');
  if (parts.length < 3) return null;
  const y = parseInt(parts[2].length === 2 ? '20' + parts[2] : parts[2], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[0], 10);
  const dt = new Date(y, m, d);
  if (isNaN(dt.getTime())) return null;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return Math.floor((today - dt) / 86400000);
}

const STATUS_FINAL = new Set(['FINALIZADO', 'REJEITADO', 'ELIMINADO POR RESIDUO']);

function buildDiasCell(p) {
  const dias = calcDiasAberto(p.emissao);
  if (dias === null) return `<td class="dias-cell">—</td>`;
  const isFinal = STATUS_FINAL.has(p.status) || p.aprovacao === 'FINALIZADO';
  if (isFinal) return `<td class="dias-cell dias-ok" title="${dias} dias no total">${dias}d</td>`;
  if (dias > 30) return `<td class="dias-cell dias-alert" title="Pedido aberto há ${dias} dias">⚠ ${dias}d</td>`;
  if (dias > 15) return `<td class="dias-cell dias-warn" title="Pedido aberto há ${dias} dias">${dias}d</td>`;
  return `<td class="dias-cell">${dias}d</td>`;
}

function buildRow(p) {
  const statusCfg = STATUS_CONFIG[p.status] || { cls: 'badge-pendente', label: p.status || 'Pendente' };
  const aprovCfg = APROV_CONFIG[p.aprovacao] || APROV_CONFIG[''];

  // Stale alert: open orders not updated for >= DIAS_PARADO_ALERTA days
  const isFinal = STATUS_FINAL.has(p.status) || p.aprovacao === 'FINALIZADO';
  const diasAtualiz = calcDiasDesdeAtualizacao(p.atualizacao);
  const isParado = !isFinal && diasAtualiz !== null && diasAtualiz >= DIAS_PARADO_ALERTA;
  const paradoTitle = isParado ? `Sem atualização há ${diasAtualiz} dia${diasAtualiz !== 1 ? 's' : ''}` : '';
  const alertIcon = isParado
    ? `<span class="parado-icon" title="${paradoTitle}">!</span>`
    : '';

  return `<tr${isParado ? ' class="row-parado"' : ''}>
    <td><strong>${p.sm || '-'}</strong>${alertIcon}</td>
    <td style="font-family:monospace; color:var(--primary); font-weight:700">${p.pedido || '-'}</td>
    <td>${p.emissao || '-'}</td>
    <td><span class="badge ${statusCfg.cls}">${statusCfg.label}</span></td>
    <td><span class="badge ${aprovCfg.cls}">${aprovCfg.label}</span></td>
    <td>${p.descricao}</td>
    ${buildDiasCell(p)}
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
    const hasFilters = (document.getElementById('filterStatus')?.value || '') ||
      (document.getElementById('filterAprov')?.value || '') ||
      (document.getElementById('filterMes')?.value || '') ||
      (document.getElementById('globalSearch')?.value || '').trim();
    const noData = !PEDIDOS.length;
    const icon = noData ? 'cloud-off' : 'search-x';
    const title = noData
      ? 'Nenhum dado carregado'
      : 'Nenhum pedido encontrado';
    const subtitle = noData
      ? 'Clique em <strong>Sincronizar</strong> para carregar os dados da planilha.'
      : hasFilters
        ? 'Tente ajustar ou limpar os filtros aplicados.'
        : 'Não há pedidos para exibir.';
    tbody.innerHTML = `<tr class="empty-state-row"><td colspan="9">
      <div class="empty-state">
        <i data-lucide="${icon}" class="empty-state-icon"></i>
        <p class="empty-state-title">${title}</p>
        <p class="empty-state-sub">${subtitle}</p>
      </div>
    </td></tr>`;
    lucide.createIcons();
    document.getElementById('filterCount').textContent = '0 pedidos';
    renderPagination(0, 0);
    return;
  }

  // Pagination
  const total = data.length;
  const { pageSize } = pageState;
  const totalPages = Math.ceil(total / pageSize);
  if (pageState.current > totalPages) pageState.current = totalPages;
  if (pageState.current < 1) pageState.current = 1;
  const start = (pageState.current - 1) * pageSize;
  const pageData = data.slice(start, start + pageSize);

  tbody.innerHTML = pageData.map(buildRow).join('');
  const showing = `${start + 1}–${Math.min(start + pageSize, total)} de ${total}`;
  document.getElementById('filterCount').textContent = `${total} pedido${total !== 1 ? 's' : ''} · mostrando ${showing}`;
  renderPagination(totalPages, pageState.current);
}

function renderPagination(totalPages, current) {
  let wrap = document.getElementById('pagination-wrap');
  if (!wrap) return;
  if (totalPages <= 1) { wrap.innerHTML = ''; return; }

  const prev = current > 1;
  const next = current < totalPages;

  // Show max 5 page buttons around current
  const pages = [];
  const range = 2;
  for (let i = Math.max(1, current - range); i <= Math.min(totalPages, current + range); i++) {
    pages.push(i);
  }

  wrap.innerHTML = `
    <button class="page-btn" onclick="goToPage(${current - 1})" ${prev ? '' : 'disabled'}>
      <i data-lucide="chevron-left" style="width:14px;height:14px"></i>
    </button>
    ${pages[0] > 1 ? `<button class="page-btn" onclick="goToPage(1)">1</button>${pages[0] > 2 ? '<span class="page-ellipsis">…</span>' : ''}` : ''}
    ${pages.map(p => `<button class="page-btn ${p === current ? 'active' : ''}" onclick="goToPage(${p})">${p}</button>`).join('')}
    ${pages[pages.length - 1] < totalPages ? `${pages[pages.length - 1] < totalPages - 1 ? '<span class="page-ellipsis">…</span>' : ''}<button class="page-btn" onclick="goToPage(${totalPages})">${totalPages}</button>` : ''}
    <button class="page-btn" onclick="goToPage(${current + 1})" ${next ? '' : 'disabled'}>
      <i data-lucide="chevron-right" style="width:14px;height:14px"></i>
    </button>
  `;
  lucide.createIcons();
}

function goToPage(n) {
  pageState.current = n;
  renderPedidosTable(filteredData.length || !PEDIDOS.length ? filteredData : sortData(PEDIDOS));
}

// ===================================================
// FILTERS & NAVIGATION
// ===================================================
function applyFilters() {
  pageState.current = 1; // reset to first page on any filter change
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
        ['HUGO ALEXANDRE', 'ELOI JOSE', 'ELOI JOSÉ', 'ADILSON RODRIGUES', 'MARCOS CARIAS', 'PENDENTE', ''].includes(p.aprovacao);
    } else if (status === 'GRP_ANDAMENTO') {
      matchStatus = ['EM PROCESSO/PARCIAL', 'EM ANDAMENTO'].includes(p.status);
    } else if (status === 'GRP_FINALIZADO') {
      matchStatus = ['FINALIZADO'].includes(p.status) || p.aprovacao === 'FINALIZADO';
    } else {
      matchStatus = p.status === status;
    }

    // Normaliza acento para comparação (ex: "ELOI JOSÉ" filtra tanto "ELOI JOSÉ" quanto "ELOI JOSE")
    const normAprov = (v) => v.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const matchAprov = !aprov || normAprov(p.aprovacao) === normAprov(aprov);

    // Filtro de mês: normaliza o ano da data para 4 dígitos antes de comparar
    const curYear = new Date().getFullYear().toString(); // ex: "2026"
    const matchMes = !mes || (() => {
      if (!p.emissao || !p.emissao.includes('/')) return false;
      const parts = p.emissao.split('/');
      if (parts.length < 3) return false;
      const monthPart = parts[1];
      const rawYear = parts[2].trim();
      // Normaliza: "26" → "2026", "2026" → "2026"
      const yearFull = rawYear.length === 2 ? '20' + rawYear : rawYear;
      return monthPart === mes && yearFull === curYear;
    })();

    return (matchSearch || matchDirect) && matchStatus && matchAprov && matchMes;
  });

  renderPedidosTable(sortData(filteredData));

  // Auto-navegar para a aba de pedidos se o usuário digitar na busca global e estiver focada
  if (search.length > 0 && document.activeElement === searchInput) {
    if (!document.getElementById('view-pedidos').classList.contains('active')) {
      showView('pedidos');
    }
  }
}

// Ordena array de pedidos pelo sortState global
function sortData(data) {
  const { col, dir } = sortState;
  if (!col) return data;
  return [...data].sort((a, b) => {
    let va, vb;
    if (col === 'dias') {
      va = calcDiasAberto(a.emissao) ?? -1;
      vb = calcDiasAberto(b.emissao) ?? -1;
    } else if (col === 'emissao' || col === 'atualizacao') {
      // Converte dd/mm/aaaa para número comparável aaaammdd
      const toNum = s => {
        if (!s || !s.includes('/')) return 0;
        const [d, m, y] = s.split('/');
        const yf = y && y.length === 2 ? '20' + y : y || '0';
        return parseInt(yf + (m || '00').padStart(2, '0') + (d || '00').padStart(2, '0'), 10);
      };
      va = toNum(a[col]); vb = toNum(b[col]);
    } else {
      va = (a[col] || '').toString().toLowerCase();
      vb = (b[col] || '').toString().toLowerCase();
    }
    if (va < vb) return dir === 'asc' ? -1 : 1;
    if (va > vb) return dir === 'asc' ? 1 : -1;
    return 0;
  });
}

// Handler de clique nos cabeçalhos — chamado pelo onclick nos <th>
function setSortCol(col) {
  if (sortState.col === col) {
    sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
  } else {
    sortState.col = col;
    sortState.dir = 'asc';
  }
  // Atualiza classes visuais nos th
  document.querySelectorAll('#tablePedidos .th-sortable').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.dataset.sort === col) {
      th.classList.add(sortState.dir === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  });
  renderPedidosTable(sortData(filteredData));
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
  if (type === 'andamento') document.getElementById('filterStatus').value = "GRP_ANDAMENTO";
  if (type === 'aguardando') document.getElementById('filterStatus').value = "AGUARD. ENTREGA/COLETA";
  if (type === 'analise') document.getElementById('filterStatus').value = "GRP_ANALISE";
  if (type === 'aprovado') document.getElementById('filterAprov').value = "APROVADO";
  if (type === 'rejeitado') document.getElementById('filterStatus').value = "REJEITADO";

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

// ===================================================
// EXPORT CSV
// ===================================================
function exportCSV() {
  const source = filteredData.length ? filteredData : PEDIDOS;
  if (!source.length) return;

  const headers = ['SM', 'Pedido', 'Emissão', 'Status', 'Aprovação', 'Descrição', 'Dias em Aberto', 'Fonte', 'Atualização'];
  const rows = source.map(p => {
    const dias = calcDiasAberto(p.emissao);
    return [
      p.sm, p.pedido, p.emissao, p.status, p.aprovacao,
      p.descricao, dias !== null ? dias : '',
      p.fonte, p.atualizacao
    ].map(v => {
      const s = (v ?? '').toString().replace(/"/g, '""');
      return /[,"\n\r]/.test(s) ? `"${s}"` : s;
    });
  });

  const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
  const bom = '\uFEFF'; // UTF-8 BOM for Excel
  const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const date = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
  a.href = url;
  a.download = `pedidos-eqs-${date}.csv`;
  a.click();
  URL.revokeObjectURL(url);
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
