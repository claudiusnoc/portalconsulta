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
  purple: '#7c3aed',
  purpleLight: 'rgba(124, 58, 237, 0.1)',
  blue: '#3b82f6',
  green: '#10b981',
  yellow: '#f59e0b',
  orange: '#f97316',
  red: '#ef4444',
  text: '#1f2937',
  muted: '#6b7280',
  border: 'rgba(0,0,0,0.05)'
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
        label: (context) => ` ${context.parsed.y || context.parsed} pedidos`
      }
    }
  },
  scales: {
    x: { grid: { display: false }, ticks: { color: '#6b7280', font: { family: 'Inter', size: 11 } } },
    y: { grid: { color: 'rgba(0,0,0,0.03)', drawBorder: false }, ticks: { color: '#6b7280', font: { family: 'Inter', size: 11 } } }
  }
};

const STATUS_CONFIG = {
  'FINALIZADO':             { cls: 'badge-finalizado',  label: 'Finalizado' },
  'EM PROCESSO/PARCIAL':    { cls: 'badge-parcial',     label: 'Em Processo/Parcial' },
  'AGUARD. ENTREGA/COLETA': { cls: 'badge-aguardando',  label: 'Aguard. Entrega/Coleta' },
  'EM ANALISE COMPRAS':     { cls: 'badge-analise',     label: 'Em Análise Compras' },
  'ENVIAR REPARO AGST':     { cls: 'badge-reparo',      label: 'Enviar Reparo AGST' },
  'ELIMINADO POR RESIDUO':  { cls: 'badge-eliminado',   label: 'Eliminado por Resíduo' },
  'EM ANDAMENTO':           { cls: 'badge-parcial',     label: 'Em Andamento' },
  'HUGO ALEXANDRE':         { cls: 'badge-analise',     label: 'Análise Hugo' },
  'ELOI JOSE':              { cls: 'badge-analise',     label: 'Análise Eloi' },
  'REJEITADO':              { cls: 'badge-eliminado',   label: 'Rejeitado' },
};

const APROV_CONFIG = {
  'APROVADO':        { cls: 'badge-aprovado', label: 'Aprovado' },
  'HUGO ALEXANDRE':  { cls: 'badge-hugo',     label: 'Hugo Alexandre' },
  'ELOI JOSE':       { cls: 'badge-hugo',     label: 'Eloi Jose' },
  'FINALIZADO':      { cls: 'badge-finalizado',label: 'Finalizado' },
  'REJEITADO':       { cls: 'badge-eliminado', label: 'Rejeitado' },
  'PENDENTE':        { cls: 'badge-pendente', label: 'Pendente' },
  '':                { cls: 'badge-pendente', label: 'Pendente' },
};

// ===================================================
// INITIALIZATION
// ===================================================
document.addEventListener('DOMContentLoaded', () => {
  setCurrentDate();
  initTheme();
  autoLoadCloud();
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
// CLOUD DATA (npoint.io - gratuito, sem autenticação)
// ===================================================
const NPOINT_KEY = 'eqs-bi-npoint-id';

async function autoLoadCloud() {
  const statusEl = document.getElementById('spStatus');
  const npointId = localStorage.getItem(NPOINT_KEY);
  
  if (!npointId) {
    // Sem dados publicados ainda
    updateDashboard(PEDIDOS);
    if (statusEl) {
      statusEl.className = 'sp-status';
      statusEl.textContent = 'Importe a planilha para começar.';
    }
    if (typeof lucide !== 'undefined') lucide.createIcons();
    return;
  }
  
  try {
    if (statusEl) {
      statusEl.className = 'sp-status loading';
      statusEl.textContent = '⏳ Carregando dados da nuvem...';
    }
    
    const response = await fetch(`https://api.npoint.io/${npointId}`);
    
    if (response.ok) {
      const cloudData = await response.json();
      
      if (cloudData && cloudData.length > 0) {
        PEDIDOS.length = 0;
        PEDIDOS.push(...cloudData);
        updateDashboard(PEDIDOS);
        showSyncBadge();
        
        if (statusEl) {
          statusEl.className = 'sp-status success';
          statusEl.textContent = `✅ ${cloudData.length} pedidos carregados!`;
        }
        
        // Mostra botão Publicar no header
        const headerBtn = document.getElementById('btnSyncHeader');
        if (headerBtn) headerBtn.style.display = 'flex';
      }
    }
  } catch (e) {
    console.log('Erro ao carregar da nuvem:', e);
    updateDashboard(PEDIDOS);
    if (statusEl) {
      statusEl.className = 'sp-status error';
      statusEl.textContent = '⚠️ Erro ao carregar. Importe manualmente.';
    }
  }
  
  if (typeof lucide !== 'undefined') lucide.createIcons();
}

async function publishToCloud() {
  if (PEDIDOS.length === 0) {
    alert('Nenhum dado para publicar. Importe a planilha primeiro.');
    return;
  }
  
  const statusEl = document.getElementById('spStatus');
  const npointId = localStorage.getItem(NPOINT_KEY);
  
  if (statusEl) {
    statusEl.className = 'sp-status loading';
    statusEl.textContent = '⏳ Publicando para equipe...';
  }
  
  try {
    let response;
    
    if (npointId) {
      // Atualiza o bin existente
      response = await fetch(`https://api.npoint.io/${npointId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(PEDIDOS)
      });
    } else {
      // Cria um novo bin
      response = await fetch('https://api.npoint.io/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(PEDIDOS)
      });
    }
    
    if (response.ok) {
      const result = await response.json();
      
      // Salva o ID para uso futuro (na URL retornada)
      if (!npointId && result) {
        // Extrai o ID da URL retornada ou do campo 'id'
        const newId = result.id || (typeof result === 'string' ? result : null);
        if (newId) {
          localStorage.setItem(NPOINT_KEY, newId);
        }
      }
      
      if (statusEl) {
        statusEl.className = 'sp-status success';
        statusEl.textContent = `✅ ${PEDIDOS.length} pedidos publicados para a equipe!`;
      }
      
      showSyncBadge();
    } else {
      throw new Error(`HTTP ${response.status}`);
    }
    
  } catch (err) {
    console.error('Publish error:', err);
    if (statusEl) {
      statusEl.className = 'sp-status error';
      statusEl.textContent = `❌ Erro ao publicar: ${err.message}`;
    }
  }
}

// ===================================================
// EXCEL IMPORT (Manual)
// ===================================================
function handleFileInput(e) {
  const file = e.target.files[0];
  if (file) handleFile(file);
}

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    
    const consolidatedData = [];
    
    workbook.SheetNames.forEach(name => {
      // Aceita a aba "COMPRAS", "COMPRAS 2026" ou "COMPRAS 26"
      // Rejeita explicitamente abas antigas como "COMPRAS 23", "COMPRAS 24", "COMPRAS 25"
      const upperName = name.toUpperCase();
      const isCompras = upperName.includes('COMPRAS');
      const isOldYear = upperName.includes('23') || upperName.includes('24') || upperName.includes('25') || upperName.includes('2023') || upperName.includes('2024') || upperName.includes('2025');

      if (isCompras && !isOldYear) {
        const worksheet = workbook.Sheets[name];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { range: 1, defval: '' });
        
        const mapped = rawData.map(row => {
          // Normaliza chaves para aceitar variações entre abas (ex: S.M. e SM, Data Emissão e Emissão)
          const normalizedRow = {};
          for (let k in row) {
            if (!k) continue;
            // Remove acentos, pontos, e garante letras maíusculas e sem espaços (ex: STATUS APROVAÇÃO -> STATUSAPROVACAO)
            const safeKey = k.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\.ºº\s]/g, "").trim();
            normalizedRow[safeKey] = row[k];
          }

          const smVal = normalizedRow['SM'] || normalizedRow['NSM'] || normalizedRow['NUMEROSM'] || 0;
          const dtEmissao = parseExcelDate(normalizedRow['EMISSAO'] || normalizedRow['DATAEMISSAO'] || normalizedRow['DATA']);
          
          let statusVal = (normalizedRow['STATUS'] || '').toString().toUpperCase().trim();
          let aprovacaoVal = (normalizedRow['STATUSAPROVACAO'] || normalizedRow['APROVACAO'] || '').toString().toUpperCase().trim();
          
          // Correção: Itens finalizados, rejeitados ou eliminados NÃO podem ter aprovação 'Pendente'
          if (['FINALIZADO', 'ELIMINADO POR RESIDUO', 'REJEITADO'].includes(statusVal)) {
            if (aprovacaoVal === '' || aprovacaoVal === 'PENDENTE') {
              aprovacaoVal = 'FINALIZADO';
            }
          }

          return {
            sm:          smVal,
            pedido:      normalizedRow['PEDIDO'] || '',
            emissao:     dtEmissao,
            status:      statusVal,
            aprovacao:   aprovacaoVal,
            descricao:   normalizedRow['DESCRICAO'] || '',
            atualizacao: parseExcelDate(normalizedRow['DATAATUALIZACAO'] || normalizedRow['ATUALIZACAO']),
            fonte:       name
          };
        }).filter(p => p.sm || p.descricao);
        
        consolidatedData.push(...mapped);
      }
    });

    if (consolidatedData.length > 0) {
      PEDIDOS.length = 0; 
      PEDIDOS.push(...consolidatedData);
      
      updateDashboard(PEDIDOS);
      showSyncBadge();
      // Re-inicializa ícones Lucide após mudança de conteúdo
      if (typeof lucide !== 'undefined') lucide.createIcons();
      
      // Mostra botão de salvar para equipe no header
      const headerBtn = document.getElementById('btnSyncHeader');
      if (headerBtn) headerBtn.style.display = 'flex';
      
      // Atualiza status na sidebar
      const statusEl = document.getElementById('spStatus');
      if (statusEl) {
        statusEl.className = 'sp-status success';
        statusEl.textContent = `✅ ${consolidatedData.length} pedidos importados! Clique em "Publicar".`;
      }
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcelDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const d = val.getDate().toString().padStart(2, '0');
    const m = (val.getMonth() + 1).toString().padStart(2, '0');
    const y = val.getFullYear();
    return `${d}/${m}/${y}`;
  }
  return val.toString().trim();
}

function showSyncBadge() {
  const badge = document.getElementById('syncBadge');
  const time = document.getElementById('syncTime');
  const now = new Date();
  badge.style.display = 'flex';
  time.textContent = `Sincronizado: ${now.getHours().toString().padStart(2,'0')}:${now.getMinutes().toString().padStart(2,'0')}`;
  document.getElementById('importBtn').innerHTML = '📥 Atualizar';
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

  // Se o mês atual estiver vazio, mostramos tudo mas avisamos no subtítulo
  const displayData = monthData.length > 0 ? monthData : data;
  const isFiltered = monthData.length > 0;
  
  const monthNames = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  document.getElementById('dashboard-subtitle').textContent = isFiltered 
    ? `Dados de ${monthNames[now.getMonth()]} ${curYear} • BH Metro ADM • EQS Engenharia`
    : `Total Acumulado (Sem dados para ${monthNames[now.getMonth()]}) • EQS Engenharia`;

  renderKPIs(displayData);
  renderCharts(displayData);
  renderPreviewTable(isFiltered ? displayData : data); // Preview focado no mês atual

  // Define o filtro do dropdown explicitamente para o mês atual e aplica à tabela de pedidos!
  // Isso garante que tabelas não carreguem "Todos os Meses" (que traz itens velhos vazios) no início.
  document.getElementById('filterMes').value = curMonth;
  applyFilters(); 
}

function renderKPIs(data) {
  const total = data.length;
  // A categoria finalizado agora agrega Eliminados e finalizações reais
  const finalizado = data.filter(p => 
    p.status === 'FINALIZADO' || 
    p.status === 'ELIMINADO POR RESIDUO' || 
    p.status === 'REJEITADO' ||
    p.aprovacao === 'FINALIZADO'
  ).length;
  
  const andamento  = data.filter(p => p.status === 'EM PROCESSO/PARCIAL' || p.status === 'EM ANDAMENTO').length;
  
  // Analise agora inclui TUDO que não está finalizado/andamento nem aprovado
  const analise    = data.filter(p => 
    p.status === 'EM ANALISE COMPRAS' || 
    p.status === 'AGUARD. ENTREGA/COLETA' || 
    p.status === 'HUGO ALEXANDRE' || 
    p.status === 'ELOI JOSE' ||
    p.aprovacao === 'HUGO ALEXANDRE' ||
    p.aprovacao === 'ELOI JOSE' ||
    p.aprovacao === 'PENDENTE' ||
    p.aprovacao === ''
  ).length;
  
  const aprovado   = data.filter(p => p.aprovacao === 'APROVADO').length;

  animateValue('kpi-total', total);
  animateValue('kpi-finalizado', finalizado);
  animateValue('kpi-andamento', andamento);
  animateValue('kpi-analise', analise);
  animateValue('kpi-aprovado', aprovado);

  // Rejeitados conta os itens com status REJEITADO ou ELIMINADO POR RESIDUO
  const rejeitado = data.filter(p => 
    p.status === 'REJEITADO' || 
    p.status === 'ELIMINADO POR RESIDUO'
  ).length;
  animateValue('kpi-rejeitado', rejeitado);

  // Update Progress Bars
  if (total > 0) {
    document.getElementById('bar-finalizado').style.width = `${(finalizado / total) * 100}%`;
    document.getElementById('bar-andamento').style.width = `${(andamento / total) * 100}%`;
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

  if (charts.status) charts.status.destroy();
  charts.status = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: Object.keys(statusGroups),
      datasets: [{
        label: 'Pedidos',
        data: Object.values(statusGroups),
        backgroundColor: CHART_COLORS.purple,
        borderRadius: 12,
        borderSkipped: false,
        barThickness: 24
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y'
    }
  });
}

function renderAprovChart(data) {
  const ctx = document.getElementById('chartAprov');
  if (!ctx) return;

  const hugoCount     = data.filter(p => p.aprovacao === 'HUGO ALEXANDRE').length;
  const eloiCount     = data.filter(p => p.aprovacao === 'ELOI JOSE').length;
  // Agrupa os aprovados totais + os finalizados automáticos
  const aprovadoCount = data.filter(p => p.aprovacao === 'APROVADO' || p.aprovacao === 'FINALIZADO').length;
  
  if (charts.aprov) charts.aprov.destroy();
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

function renderMonthlyChart(data) {
  const ctx = document.getElementById('chartMes');
  if (!ctx) return;

  const monthCounts = {};
  const monthLabels = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  
  // Inicializa meses com zero
  monthLabels.forEach(m => monthCounts[m] = 0);

  data.forEach(p => {
    if (!p.emissao || !p.emissao.includes('/')) return;
    const parts = p.emissao.split('/');
    const monthIdx = parseInt(parts[1], 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12) {
      monthCounts[monthLabels[monthIdx]]++;
    }
  });

  if (charts.mes) charts.mes.destroy();
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
          const {ctx, chartArea} = chart;
          if (!chartArea) return null;
          const gradient = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
          gradient.addColorStop(0, 'rgba(124, 58, 237, 0)');
          gradient.addColorStop(1, 'rgba(124, 58, 237, 0.15)');
          return gradient;
        },
        fill: true,
        tension: 0.4, // Curva Spline
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

// ===================================================
// TABLES
// ===================================================
function buildRow(p) {
  const statusCfg = STATUS_CONFIG[p.status] || { cls: 'badge-pendente', label: p.status || 'Pendente' };
  const aprovCfg = APROV_CONFIG[p.aprovacao] || APROV_CONFIG[''];
  
  return `<tr>
    <td><strong>${p.sm || '-'}</strong></td>
    <td style="font-family:monospace; color:var(--purple); font-weight:700">${p.pedido || '-'}</td>
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
  const search  = (searchInput?.value || '').toLowerCase().trim();
  const status  = document.getElementById('filterStatus')?.value || '';
  const aprov   = document.getElementById('filterAprov')?.value || '';
  const mes     = document.getElementById('filterMes')?.value || '';

  filteredData = PEDIDOS.filter(p => {
    // Busca abrangente
    const matchSearch = !search || [p.sm, p.pedido, p.descricao, p.status, p.aprovacao].some(f => 
      (f || '').toString().toLowerCase().includes(search)
    );
    
    // Suporte a busca direta pelo ID do pedido ou SM
    const matchDirect = !search || (p.pedido || '').toString().toLowerCase() === search || (p.sm || '').toString().toLowerCase() === search;

    // Lógica simplificada e consolidada para "Em Análise Compras"
    const isAnaliseStatus = status === 'EM ANALISE COMPRAS' && (
      p.status === 'EM ANALISE COMPRAS' ||
      p.status === 'AGUARD. ENTREGA/COLETA' ||
      p.status === 'HUGO ALEXANDRE' || 
      p.status === 'ELOI JOSE' || 
      p.aprovacao === 'HUGO ALEXANDRE' || 
      p.aprovacao === 'ELOI JOSE' ||
      p.aprovacao === 'PENDENTE' ||
      p.aprovacao === ''
    );

    const matchStatus = !status || p.status === status || isAnaliseStatus;
    const matchAprov  = !aprov  || p.aprovacao === aprov;
    // Filtro de mês agora garante que o ano seja sempre do ano corrente pra evitar mistura de anos passados (Mar 2024, Mar 2025...)
    const curYear = new Date().getFullYear().toString();
    const matchMes    = !mes || (p.emissao && p.emissao.split('/')[1] === mes && (p.emissao.split('/')[2] === curYear || p.emissao.split('/')[2] === curYear.slice(-2)));
    
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
  // Reset all filters
  document.getElementById('globalSearch').value = "";
  document.getElementById('filterStatus').value = "";
  document.getElementById('filterAprov').value = "";
  document.getElementById('filterMes').value = ""; // Remove o filtro de mês para KPIs globais!

  if (type === 'finalizado') document.getElementById('filterStatus').value = "FINALIZADO";
  if (type === 'andamento')  document.getElementById('filterStatus').value = "EM PROCESSO/PARCIAL";
  if (type === 'analise')    document.getElementById('filterStatus').value = "EM ANALISE COMPRAS";
  if (type === 'aprovado')   document.getElementById('filterAprov').value = "APROVADO";
  if (type === 'rejeitado')  document.getElementById('filterStatus').value = "ELIMINADO POR RESIDUO";

  applyFilters();
  showView('pedidos'); // Volta para a tela de resultados
}

function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(`view-${id}`).classList.add('active');
  
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById(`nav-${id}`).classList.add('active');
  
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

  if (charts.statusFull) charts.statusFull.destroy();
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
