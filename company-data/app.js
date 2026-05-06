// === State ===
const state = {
  selected: null,
  sector: '',
  search: '',
  inputs: {
    multiple: 7,
    growth: 0,
    uplift: 0,
    risk: 20,
    eps: null,
    epsOverridden: false,
    assetValue: null,        // total asset value in millions; when set, drives valuation in asset-mode
    proposedAction: '',
    researchGrade: '',
    probability: 50,
  },
  charts: {},
  chartBuilders: {},
  chartTitles: {},
  modalChart: null,
};

// === Utility ===
const $ = (id) => document.getElementById(id);

const fmt = (v, dp = 2) => {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return Number(v).toLocaleString('en-GB', { minimumFractionDigits: dp, maximumFractionDigits: dp });
};

const fmtPct = (v, dp = 1) => (v === null || v === undefined || isNaN(v)) ? '—' : `${Number(v).toFixed(dp)}%`;

const currencySym = (cur) => ({ GBP: '£', EUR: '€', USD: '$', GBX: 'p' }[cur] || '');

// === Company list render ===
function renderCompanyList() {
  const q = state.search.toLowerCase().trim();
  const filtered = COMPANIES.filter((c) => {
    if (state.sector && c.sector !== state.sector) return false;
    if (!q) return true;
    return (c.name && c.name.toLowerCase().includes(q)) || (c.tidm && c.tidm.toLowerCase().includes(q));
  });

  const list = $('companyList');
  list.innerHTML = filtered.slice(0, 300).map((c) => {
    const isActive = state.selected && state.selected.name === c.name;
    return `<div class="company-item ${isActive ? 'active' : ''}" data-name="${encodeURIComponent(c.name)}">
      <div class="company-item-name">${c.name}</div>
      <div class="company-item-tidm">${c.tidm || '—'}</div>
    </div>`;
  }).join('');

  $('searchCount').textContent = `${filtered.length} of ${COMPANIES.length} companies${filtered.length > 300 ? ' (showing 300)' : ''}`;

  list.querySelectorAll('.company-item').forEach((el) => {
    el.addEventListener('click', () => {
      const name = decodeURIComponent(el.dataset.name);
      selectCompany(name);
    });
  });
}

function populateSectorFilter() {
  const sectors = [...new Set(COMPANIES.map((c) => c.sector).filter(Boolean))].sort();
  const sel = $('sectorFilter');
  sectors.forEach((s) => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    sel.appendChild(opt);
  });
}

// === Valuation math (mirrors spreadsheet) ===
function computeGrowth(c) {
  if (!c.eps || !c.eps3yAgo || c.eps3yAgo <= 0) return 0.01;
  const g = (Math.pow(c.eps / c.eps3yAgo, 1 / 3) - 1) * 100;
  return Math.max(0.01, g);
}

function defaultMultiple(marketCap) {
  if (marketCap === null || marketCap === undefined) return 6;
  if (marketCap >= 500) return 8;
  if (marketCap >= 100) return 7;
  return 6;
}

function valuation(c, inputs) {
  const { multiple, growth, uplift, risk, assetValue } = inputs;
  let formula, mode;
  let assetsPerShare = null;

  // Asset mode: if user has entered a total asset value, formula = AssetValue × 100 / Shares (per-share in price units)
  if (assetValue !== null && assetValue !== undefined && !isNaN(assetValue) && assetValue !== '' && c.sharesInIssue && c.sharesInIssue > 0) {
    assetsPerShare = (Number(assetValue) * 100) / c.sharesInIssue;
    formula = assetsPerShare;
    mode = 'asset';
  } else {
    const eps = (inputs.eps !== null && inputs.eps !== undefined && !isNaN(inputs.eps)) ? inputs.eps : (c.eps || 0);
    formula = ((eps * (growth / 100) * multiple) + eps) * multiple;
    mode = 'eps';
  }

  const actual = formula * ((100 + uplift) / 100);
  const valuedCap = (c.price && c.price > 0) ? (c.marketCap / c.price) * actual : null;
  const opportunity = (c.price && c.price > 0) ? (((100 / c.price) * actual) - 100) / 100 : null;
  const riskFactor = risk / 100;
  const netOpportunity = opportunity !== null ? opportunity - riskFactor : null;
  return { formula, actual, valuedCap, opportunity, riskFactor, netOpportunity, mode, assetsPerShare };
}

// === Selection ===
function selectCompany(name) {
  const c = COMPANIES.find((x) => x.name === name);
  if (!c) return;
  state.selected = c;

  // Reset inputs to defaults for this company
  state.inputs.multiple = defaultMultiple(c.marketCap);
  state.inputs.growth = computeGrowth(c);
  state.inputs.uplift = 0;
  state.inputs.risk = 20;
  state.inputs.eps = c.eps;
  state.inputs.epsOverridden = false;
  state.inputs.assetValue = null;
  state.inputs.proposedAction = '';
  state.inputs.researchGrade = '';
  state.inputs.probability = 50;

  syncInputsUI();
  renderCompanyList();
  render();
}

function syncInputsUI() {
  $('inMultiple').value = state.inputs.multiple;
  $('inGrowth').value = state.inputs.growth;
  $('inUplift').value = state.inputs.uplift;
  $('inRisk').value = state.inputs.risk;
  $('inProbability').value = state.inputs.probability;
  $('inProposedAction').value = state.inputs.proposedAction || '';
  $('inResearchGrade').value = state.inputs.researchGrade || '';
  $('inAssetValue').value = state.inputs.assetValue !== null ? state.inputs.assetValue : '';
  const epsInput = $('inEPS');
  epsInput.value = state.inputs.eps !== null ? state.inputs.eps : '';
  epsInput.classList.toggle('overridden', state.inputs.epsOverridden);
  updateInputLabels();
}

function setNumField(id, value, dp) {
  const el = $(id);
  if (!el) return;
  // Don't clobber while the user is editing that field
  if (document.activeElement === el) return;
  const n = Number(value);
  el.value = isNaN(n) ? '' : n.toFixed(dp);
}

function updateInputLabels() {
  setNumField('outMultiple', state.inputs.multiple, 2);
  setNumField('outGrowth', state.inputs.growth, 2);
  setNumField('outUplift', state.inputs.uplift, 2);
  setNumField('outRisk', state.inputs.risk, 2);
  setNumField('outProbability', state.inputs.probability, 0);

  if (state.selected) {
    const c = state.selected;
    const raw = computeGrowth(c);
    $('growthHint').textContent = `Auto (3y EPS): ${fmtPct(raw, 1)}`;

    if (state.inputs.epsOverridden) {
      $('epsHint').textContent = `Override active — reported EPS: ${fmt(c.eps, 2)}`;
    } else {
      $('epsHint').textContent = `Reported EPS: ${fmt(c.eps, 2)}`;
    }

    // Asset hint — show computed per-share when asset mode is active
    const av = state.inputs.assetValue;
    if (av !== null && av !== '' && !isNaN(av) && c.sharesInIssue > 0) {
      const aps = (Number(av) * 100) / c.sharesInIssue;
      $('assetHint').textContent = `Asset mode active — Assets/share = ${fmt(aps, 1)} (${currencySym(c.currency)}${fmt(av, 1)}m ÷ ${fmt(c.sharesInIssue, 1)}m × 100). Growth & Multiple ignored.`;
    } else {
      $('assetHint').textContent = 'When set, valuation = Asset Value × 100 / Shares. Growth & Multiple are ignored.';
    }
  }
}

// === Render ===
function render() {
  const c = state.selected;
  if (!c) return;
  const cur = currencySym(c.currency);

  // Header
  $('companyTidm').textContent = `${c.tidm || ''} · ${c.sector || ''}`;
  $('companyName').textContent = c.name;
  $('companyMeta').textContent = [c.industry, c.supersector, c.subsector].filter(Boolean).join(' › ');

  // KPIs
  $('kpiPrice').textContent = c.price !== null ? `${cur}${fmt(c.price, 2)}` : '—';
  const ch = c.changePct;
  const chEl = $('kpiChange');
  chEl.textContent = ch !== null ? `${ch > 0 ? '+' : ''}${fmt(ch, 2)}% today` : '—';
  chEl.className = 'kpi-delta ' + (ch > 0 ? 'up' : ch < 0 ? 'down' : '');

  $('kpiMcap').textContent = c.marketCap !== null ? `${cur}${fmt(c.marketCap, 1)}m` : '—';
  $('kpiMcapSub').textContent = c.marketCap !== null ? (c.marketCap >= 1000 ? `${cur}${fmt(c.marketCap / 1000, 2)}bn` : 'small cap') : '';
  $('kpiPE').textContent = fmt(c.pe, 1);
  $('kpiPEG').textContent = c.peg !== null ? `PEG ${fmt(c.peg, 2)}` : '';
  $('kpiEPS').textContent = fmt(c.eps, 2);
  $('kpiEPS3y').textContent = c.eps3yAgo !== null ? `3y ago: ${fmt(c.eps3yAgo, 2)}` : '';
  $('kpiYield').textContent = c.yield !== null ? fmtPct(c.yield, 1) : '—';
  $('kpiShares').textContent = fmt(c.sharesInIssue, 1);
  $('kpiCurrency').textContent = c.currency || '';

  // Valuation
  const v = valuation(c, state.inputs);
  $('rFormula').textContent = `${cur}${fmt(v.formula, 2)}`;
  $('rActual').textContent = `${cur}${fmt(v.actual, 2)}`;
  $('rValuedCap').textContent = v.valuedCap !== null ? `${cur}${fmt(v.valuedCap, 1)}m` : '—';
  const oppPct = v.opportunity !== null ? v.opportunity * 100 : null;
  $('rOpp').textContent = oppPct !== null ? `${oppPct > 0 ? '+' : ''}${fmt(oppPct, 1)}%` : '—';

  const oppItem = $('rOppItem');
  oppItem.classList.remove('positive', 'negative');
  if (oppPct !== null) oppItem.classList.add(oppPct > 0 ? 'positive' : 'negative');

  // Formula hint swaps depending on mode
  const fh = $('formulaHint');
  if (fh) {
    fh.innerHTML = v.mode === 'asset'
      ? 'Formula (Asset Mode): <code>(Asset Value × 100 / Shares) × ((100 + Uplift) / 100)</code>'
      : 'Formula: <code>((EPS × (Growth/100) × Multiple) + EPS) × Multiple</code>';
  }

  // Charts
  renderPriceVsTarget(c, v);
  renderGauge(v);
  renderEPS(c);
  renderPeers(c);
  renderSensitivity(c);

  // Tear sheet
  renderTearsheet(c);
}

// === Tear sheet ===
const TEARSHEET_YEARS = 6; // Show the last N years across all charts/table/CAGR grid

function normaliseTidm(t) {
  return (t || '').toString().toUpperCase().replace(/[._]/g, '').trim();
}

function trimTearsheet(ts, n) {
  if (!ts || !ts.years || ts.years.length <= n) return ts;
  const start = ts.years.length - n;
  const metrics = {};
  for (const k of Object.keys(ts.metrics || {})) {
    metrics[k] = (ts.metrics[k] || []).slice(start);
  }
  return {
    years: ts.years.slice(start),
    periodEnding: (ts.periodEnding || []).slice(start),
    metrics,
  };
}

function getTearsheet(c) {
  if (typeof TEARSHEETS === 'undefined') return null;
  const key = normaliseTidm(c.tidm);
  return TEARSHEETS[key] || null;
}

function cagr(start, end, years) {
  if (start === null || end === null || years <= 0) return null;
  if (start === 0) return null;
  // If sign flip, CAGR is not meaningful
  if ((start < 0) !== (end < 0)) return null;
  const ratio = end / start;
  if (ratio <= 0) return null;
  return (Math.pow(ratio, 1 / years) - 1) * 100;
}

function yoy(arr) {
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    if (i === 0 || arr[i] === null || arr[i - 1] === null || arr[i - 1] === 0) {
      out.push(null);
    } else {
      // Use abs of prior to preserve sign of change even across sign flips
      out.push(((arr[i] - arr[i - 1]) / Math.abs(arr[i - 1])) * 100);
    }
  }
  return out;
}

function movingAverage(arr, window) {
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    if (i < window - 1) { out.push(null); continue; }
    const slice = arr.slice(i - window + 1, i + 1);
    if (slice.some((v) => v === null)) { out.push(null); continue; }
    out.push(slice.reduce((a, b) => a + b, 0) / window);
  }
  return out;
}

function fmtMillions(v, cur) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  const abs = Math.abs(v);
  if (abs >= 1000) return `${cur}${(v / 1000).toLocaleString('en-GB', { maximumFractionDigits: 2 })}bn`;
  return `${cur}${v.toLocaleString('en-GB', { maximumFractionDigits: 0 })}m`;
}

function renderTearsheet(c) {
  const rawTs = getTearsheet(c);
  const body = $('tearsheetBody');
  const empty = $('tearsheetEmpty');
  const badges = $('tearsheetBadges');
  const cur = currencySym(c.currency);

  if (!rawTs || !rawTs.years || rawTs.years.length === 0) {
    body.style.display = 'none';
    empty.style.display = 'block';
    $('tearsheetTitle').textContent = 'Financial History';
    $('tearsheetSub').textContent = `${c.name} · ${c.tidm || ''}`;
    badges.innerHTML = '';
    ['chartTurnover','chartTurnoverGrowth','chartPostTax','chartEPSGrowth'].forEach((k) => destroyChart(k));
    return;
  }

  body.style.display = 'block';
  empty.style.display = 'none';

  // Window everything to the last N years
  const ts = trimTearsheet(rawTs, TEARSHEET_YEARS);
  const years = ts.years;
  const firstYr = years[0];
  const lastYr = years[years.length - 1];
  const turnover = ts.metrics.turnover || [];
  const postTax = ts.metrics.postTaxProfit || [];
  const reportedEPS = ts.metrics.reportedEPS || [];
  const turnoverPctChg = ts.metrics.turnoverPctChg || [];

  // Header
  $('tearsheetTitle').textContent = `${c.name} — Financial History`;
  $('tearsheetSub').textContent = `${c.tidm || ''} · ${c.sector || ''} · ${c.subsector || ''}`;
  badges.innerHTML = `
    <div class="ts-badge">${firstYr}–${lastYr}</div>
    <div class="ts-badge">${years.length} years</div>
  `;

  // KPIs
  $('tsPeriod').textContent = `${firstYr}–${lastYr}`;

  const firstTurnover = turnover.find((v) => v !== null);
  const lastTurnover = [...turnover].reverse().find((v) => v !== null);
  const firstTurnoverIdx = turnover.findIndex((v) => v !== null);
  const lastTurnoverIdx = turnover.length - 1 - [...turnover].reverse().findIndex((v) => v !== null);
  const turnoverYears = lastTurnoverIdx - firstTurnoverIdx;
  const turnoverCagr = cagr(firstTurnover, lastTurnover, turnoverYears);
  $('tsTurnoverCagr').textContent = turnoverCagr !== null ? `${turnoverCagr > 0 ? '+' : ''}${turnoverCagr.toFixed(1)}%` : '—';
  $('tsTurnoverCagrSub').textContent = turnoverYears > 0 ? `${years[firstTurnoverIdx]}→${years[lastTurnoverIdx]}` : '—';

  const firstProfit = postTax.find((v) => v !== null);
  const lastProfit = [...postTax].reverse().find((v) => v !== null);
  const firstProfitIdx = postTax.findIndex((v) => v !== null);
  const lastProfitIdx = postTax.length - 1 - [...postTax].reverse().findIndex((v) => v !== null);
  const profitYears = lastProfitIdx - firstProfitIdx;
  const profitCagr = cagr(firstProfit, lastProfit, profitYears);
  $('tsProfitCagr').textContent = profitCagr !== null ? `${profitCagr > 0 ? '+' : ''}${profitCagr.toFixed(1)}%` : 'n/m';
  $('tsProfitCagrSub').textContent = profitYears > 0 ? `${years[firstProfitIdx]}→${years[lastProfitIdx]}` : '—';

  $('tsLatestProfit').textContent = fmtMillions(lastProfit, cur);
  $('tsLatestProfitSub').textContent = lastProfit !== null ? `FY${years[lastProfitIdx]}` : '—';

  // EPS CAGR by look-back
  renderEpsCagrGrid(reportedEPS, years);

  // Charts
  renderChartTurnover(years, turnover, cur);
  renderChartTurnoverGrowth(years, turnoverPctChg, turnover);
  renderChartPostTax(years, postTax, cur);
  renderChartEPSGrowth(years, reportedEPS);

  // Data table
  renderTearsheetTable(ts, cur);
}

function renderEpsCagrGrid(eps, years) {
  const grid = $('tsEpsCagrGrid');
  const n = eps.length;
  const endIdx = n - 1 - [...eps].reverse().findIndex((v) => v !== null);
  if (endIdx < 0) { grid.innerHTML = '<div class="ts-kpi-sub">No reported EPS data</div>'; return; }
  const endVal = eps[endIdx];
  const cells = [];
  const maxPeriod = Math.min(10, eps.length - 1);
  for (let period = 1; period <= maxPeriod; period++) {
    const startIdx = endIdx - period;
    if (startIdx < 0) {
      cells.push(`<div class="ts-eps-cell na"><div class="ts-eps-period">${period}y</div><div class="ts-eps-val">—</div></div>`);
      continue;
    }
    const startVal = eps[startIdx];
    const g = cagr(startVal, endVal, period);
    if (g === null) {
      cells.push(`<div class="ts-eps-cell na"><div class="ts-eps-period">${period}y</div><div class="ts-eps-val">n/m</div></div>`);
    } else {
      const cls = g >= 0 ? 'positive' : 'negative';
      cells.push(`<div class="ts-eps-cell ${cls}"><div class="ts-eps-period">${period}y</div><div class="ts-eps-val">${g > 0 ? '+' : ''}${g.toFixed(1)}%</div></div>`);
    }
  }
  grid.innerHTML = cells.join('');
}

function renderChartTurnover(years, turnover, cur) {
  const ctx = $('chartTurnover').getContext('2d');
  const builder = (targetCtx) => {
    const h = targetCtx.canvas.height || 280;
    const grad = targetCtx.createLinearGradient(0, 0, 0, h);
    grad.addColorStop(0, 'rgba(165, 50, 28,0.3)');
    grad.addColorStop(1, 'rgba(165, 50, 28,0.0)');
    return {
      type: 'line',
      data: {
        labels: years,
        datasets: [{
          label: 'Turnover',
          data: turnover,
          borderColor: chartColors.accent,
          backgroundColor: grad,
          fill: true,
          tension: 0.3,
          pointRadius: 3,
          pointHoverRadius: 6,
          borderWidth: 2.5,
          spanGaps: true,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: chartColors.tooltipBg, borderColor: chartColors.tooltipBorder, borderWidth: 1,
            titleColor: chartColors.tooltipTitle, bodyColor: chartColors.tooltipBody, padding: 12, cornerRadius: 8,
            callbacks: { label: (c) => `${cur}${Number(c.parsed.y).toLocaleString('en-GB', { maximumFractionDigits: 0 })}m` },
          },
        },
        scales: {
          x: baseScales.x,
          y: { ...baseScales.y, ticks: { ...baseScales.y.ticks, callback: (v) => `${cur}${v >= 1000 ? (v/1000).toFixed(1)+'bn' : v+'m'}` } },
        },
      },
    };
  };
  const name = (state.selected && state.selected.name) || '';
  renderExpandable('chartTurnover', ctx, `${name} — Turnover (Full History)`, builder);
}

function renderChartTurnoverGrowth(years, pctChg, turnover) {
  const ctx = $('chartTurnoverGrowth').getContext('2d');
  const hasProvided = pctChg && pctChg.some((v) => v !== null);
  const data = hasProvided ? pctChg : yoy(turnover);
  const colors = data.map((v) => v === null ? 'rgba(148,163,184,0.4)' : (v >= 0 ? 'rgba(47, 122, 57,0.75)' : 'rgba(165, 50, 28,0.75)'));
  const builder = () => ({
    type: 'bar',
    data: {
      labels: years,
      datasets: [{
        label: 'YoY Growth (%)',
        data,
        backgroundColor: colors,
        borderRadius: 4,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: chartColors.tooltipBg, borderColor: chartColors.tooltipBorder, borderWidth: 1,
          titleColor: chartColors.tooltipTitle, bodyColor: chartColors.tooltipBody, padding: 12, cornerRadius: 8,
          callbacks: { label: (c) => c.parsed.y === null ? '—' : `${c.parsed.y > 0 ? '+' : ''}${c.parsed.y.toFixed(1)}%` },
        },
      },
      scales: {
        x: baseScales.x,
        y: { ...baseScales.y, ticks: { ...baseScales.y.ticks, callback: (v) => `${v}%` }, beginAtZero: false },
      },
    },
  });
  const name = (state.selected && state.selected.name) || '';
  renderExpandable('chartTurnoverGrowth', ctx, `${name} — Turnover YoY Growth`, builder);
}

function renderChartPostTax(years, postTax, cur) {
  const ctx = $('chartPostTax').getContext('2d');
  const ma3 = movingAverage(postTax, 3);
  const builder = () => ({
    type: 'line',
    data: {
      labels: years,
      datasets: [
        {
          label: 'Post-tax Profit',
          data: postTax,
          borderColor: chartColors.accent,
          backgroundColor: 'rgba(165, 50, 28,0.08)',
          fill: true,
          tension: 0.3,
          pointRadius: 3,
          pointHoverRadius: 6,
          borderWidth: 2.5,
          spanGaps: true,
        },
        {
          label: '3-Year Moving Average',
          data: ma3,
          borderColor: chartColors.amber,
          borderDash: [6, 4],
          backgroundColor: 'transparent',
          fill: false,
          tension: 0.3,
          pointRadius: 0,
          pointHoverRadius: 5,
          borderWidth: 2,
          spanGaps: true,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: chartColors.text, font: { family: 'Inter Tight' } } },
        tooltip: {
          backgroundColor: chartColors.tooltipBg, borderColor: chartColors.tooltipBorder, borderWidth: 1,
          titleColor: chartColors.tooltipTitle, bodyColor: chartColors.tooltipBody, padding: 12, cornerRadius: 8,
          callbacks: { label: (c) => `${c.dataset.label}: ${c.parsed.y === null ? '—' : cur + Number(c.parsed.y).toLocaleString('en-GB', { maximumFractionDigits: 0 }) + 'm'}` },
        },
      },
      scales: {
        x: baseScales.x,
        y: { ...baseScales.y, ticks: { ...baseScales.y.ticks, callback: (v) => `${cur}${v >= 1000 ? (v/1000).toFixed(1)+'bn' : v+'m'}` }, beginAtZero: false },
      },
    },
  });
  const name = (state.selected && state.selected.name) || '';
  renderExpandable('chartPostTax', ctx, `${name} — Reported Post-tax Profit + 3-Year Moving Average`, builder);
}

function renderChartEPSGrowth(years, eps) {
  const ctx = $('chartEPSGrowth').getContext('2d');
  const growth = yoy(eps);
  const colors = growth.map((v) => v === null ? 'rgba(148,163,184,0.4)' : (v >= 0 ? 'rgba(47, 122, 57,0.75)' : 'rgba(165, 50, 28,0.75)'));
  const builder = () => ({
    type: 'bar',
    data: {
      labels: years,
      datasets: [{
        label: 'EPS YoY Growth (%)',
        data: growth,
        backgroundColor: colors,
        borderRadius: 4,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: chartColors.tooltipBg, borderColor: chartColors.tooltipBorder, borderWidth: 1,
          titleColor: chartColors.tooltipTitle, bodyColor: chartColors.tooltipBody, padding: 12, cornerRadius: 8,
          callbacks: { label: (c) => c.parsed.y === null ? '—' : `${c.parsed.y > 0 ? '+' : ''}${c.parsed.y.toFixed(1)}%` },
        },
      },
      scales: {
        x: baseScales.x,
        y: { ...baseScales.y, ticks: { ...baseScales.y.ticks, callback: (v) => `${v}%` }, beginAtZero: false },
      },
    },
  });
  const name = (state.selected && state.selected.name) || '';
  renderExpandable('chartEPSGrowth', ctx, `${name} — Reported EPS YoY Growth`, builder);
}

function renderTearsheetTable(ts, cur) {
  const metricOrder = [
    ['turnover', 'Turnover (m)', 'm', true],
    ['totalExpenses', 'Total Expenses (m)', 'm', false],
    ['turnoverPctChg', 'Turnover % Change', '%', false],
    ['operatingProfit', 'Operating Profit (m)', 'm', false],
    ['grossMargin', 'Gross Margin (%)', '%', false],
    ['preTaxProfit', 'Pre-tax Profit (m)', 'm', false],
    ['postTaxProfit', 'Reported Post-tax Profit (m)', 'm', true],
    ['sharesInIssue', 'Shares in Issue (m)', 'num', false],
    ['adjustedEPS', 'Adjusted EPS', 'num', false],
    ['reportedEPS', 'Reported EPS', 'num', true],
    ['dividendPerShare', 'Dividend per Share (adj.)', 'num', false],
    ['currentAssets', 'Current Assets (m)', 'm', false],
    ['totalAssets', 'Total Assets (m)', 'm', false],
    ['currentLiabilities', 'Current Liabilities (m)', 'm', false],
    ['totalLiabilities', 'Total Liabilities (m)', 'm', false],
    ['netBorrowing', 'Net Borrowing (m)', 'm', false],
    ['nav', 'NAV (m)', 'm', false],
    ['totalEquity', 'Total Equity (m)', 'm', false],
    ['profitOnTurnover', '% Profit on Turnover', '%', false],
  ];

  const years = ts.years;
  let html = '<thead><tr><th>Metric</th>';
  years.forEach((y) => { html += `<th>${y}</th>`; });
  html += '</tr></thead><tbody>';

  metricOrder.forEach(([key, label, unit, emph]) => {
    const arr = ts.metrics[key] || [];
    html += `<tr class="${emph ? 'ts-row-emphasis' : ''}"><th>${label}</th>`;
    for (let i = 0; i < years.length; i++) {
      const v = arr[i];
      if (v === null || v === undefined || isNaN(v)) {
        html += '<td class="ts-cell-null">—</td>';
      } else if (unit === '%') {
        html += `<td>${v.toFixed(1)}%</td>`;
      } else if (unit === 'm') {
        html += `<td>${cur}${v.toLocaleString('en-GB', { maximumFractionDigits: 0 })}</td>`;
      } else {
        html += `<td>${v.toLocaleString('en-GB', { maximumFractionDigits: 2 })}</td>`;
      }
    }
    html += '</tr>';
  });
  html += '</tbody>';
  $('tsTable').innerHTML = html;
}

// === Charts ===
const chartColors = {
  grid: 'rgba(15,23,42,0.06)',
  text: '#5a6370',
  accent: '#a5321c',
  accent2: '#b8923a',
  green: '#2f7a39',
  red: '#a5321c',
  amber: '#b8923a',
  tooltipBg: '#ffffff',
  tooltipBorder: 'rgba(15,23,42,0.1)',
  tooltipTitle: '#0f1419',
  tooltipBody: '#5a6370',
};

function destroyChart(key) {
  if (state.charts[key]) {
    state.charts[key].destroy();
    state.charts[key] = null;
  }
}

// Register an expandable chart by storing its builder so it can be replayed
// into the modal canvas later.
function renderExpandable(key, targetCtx, title, builder) {
  destroyChart(key);
  state.chartBuilders[key] = builder;
  state.chartTitles[key] = title;
  state.charts[key] = new Chart(targetCtx, builder(targetCtx));
}

function openChartModal(key) {
  const builder = state.chartBuilders[key];
  if (!builder) return;
  const modal = $('chartModal');
  const canvas = $('chartModalCanvas');
  $('chartModalTitle').textContent = state.chartTitles[key] || 'Chart';

  // Destroy any previous modal chart
  if (state.modalChart) { state.modalChart.destroy(); state.modalChart = null; }

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');

  // Wait a frame for the modal to lay out so the canvas gets correct dimensions
  requestAnimationFrame(() => {
    const ctx = canvas.getContext('2d');
    state.modalChart = new Chart(ctx, builder(ctx));
  });
}

function closeChartModal() {
  const modal = $('chartModal');
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
  if (state.modalChart) { state.modalChart.destroy(); state.modalChart = null; }
}

const baseScales = {
  x: {
    grid: { color: chartColors.grid, drawBorder: false },
    ticks: { color: chartColors.text, font: { size: 11, family: 'Inter Tight' } },
  },
  y: {
    grid: { color: chartColors.grid, drawBorder: false },
    ticks: { color: chartColors.text, font: { size: 11, family: 'Inter Tight' } },
    beginAtZero: true,
  },
};

function renderPriceVsTarget(c, v) {
  destroyChart('priceVsTarget');
  const ctx = $('chartPriceVsTarget').getContext('2d');
  const priceVal = c.price || 0;
  const targetVal = v.actual || 0;
  const gain = targetVal > priceVal ? targetVal - priceVal : 0;

  const gradCur = ctx.createLinearGradient(0, 0, 0, 240);
  gradCur.addColorStop(0, 'rgba(156,163,175,0.9)');
  gradCur.addColorStop(1, 'rgba(75,85,99,0.5)');

  const gradTarg = ctx.createLinearGradient(0, 0, 0, 240);
  gradTarg.addColorStop(0, targetVal > priceVal ? 'rgba(34,197,94,0.9)' : 'rgba(239,68,68,0.9)');
  gradTarg.addColorStop(1, targetVal > priceVal ? 'rgba(34,197,94,0.3)' : 'rgba(239,68,68,0.3)');

  state.charts.priceVsTarget = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Current Price', 'Target Valuation'],
      datasets: [{
        data: [priceVal, targetVal],
        backgroundColor: [gradCur, gradTarg],
        borderRadius: 8,
        borderSkipped: false,
        barPercentage: 0.6,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: chartColors.tooltipBg,
          borderColor: chartColors.tooltipBorder,
          borderWidth: 1,
          titleColor: chartColors.tooltipTitle,
          bodyColor: chartColors.tooltipBody,
          padding: 12,
          cornerRadius: 8,
        },
      },
      scales: baseScales,
    },
  });
}

function renderGauge(v) {
  destroyChart('gauge');
  const ctx = $('chartGauge').getContext('2d');
  const oppPct = v.opportunity !== null ? Math.max(-100, Math.min(200, v.opportunity * 100)) : 0;
  const riskPct = v.riskFactor * 100;

  // Build a half-donut: show opportunity vs risk
  const dataPos = Math.max(0, oppPct);
  const dataNeg = Math.max(0, -oppPct);

  state.charts.gauge = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Value Opportunity', 'Risk Factor', 'Headroom'],
      datasets: [{
        data: [Math.max(0.1, Math.abs(oppPct)), riskPct, Math.max(10, 100 - Math.abs(oppPct) - riskPct)],
        backgroundColor: [
          oppPct >= 0 ? chartColors.green : chartColors.red,
          chartColors.amber,
          'rgba(15,23,42,0.06)',
        ],
        borderColor: 'transparent',
        borderWidth: 0,
        circumference: 180,
        rotation: 270,
        cutout: '70%',
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: chartColors.text, font: { size: 11, family: 'Inter Tight' }, boxWidth: 10, padding: 10 },
        },
        tooltip: {
          backgroundColor: chartColors.tooltipBg,
          borderColor: chartColors.tooltipBorder,
          borderWidth: 1,
          titleColor: chartColors.tooltipTitle,
          bodyColor: chartColors.tooltipBody,
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            label: (ctx) => `${ctx.label}: ${ctx.parsed.toFixed(1)}%`,
          },
        },
      },
    },
  });

  $('gaugeCaption').textContent = `Upside ${oppPct > 0 ? '+' : ''}${oppPct.toFixed(1)}% · Risk ${riskPct.toFixed(0)}% · Net ${(oppPct - riskPct).toFixed(1)}%`;
}

function renderEPS(c) {
  const ctx = $('chartEPS').getContext('2d');

  let labels, data, titleSuffix;
  if (c.epsHistory) {
    const allLabels = ['10y ago','9y ago','8y ago','7y ago','6y ago','5y ago','4y ago','3y ago','2y ago','1y ago','Current'];
    let firstIdx = 0;
    while (firstIdx < c.epsHistory.length && (c.epsHistory[firstIdx] === null || c.epsHistory[firstIdx] === undefined)) firstIdx++;
    labels = allLabels.slice(firstIdx);
    data = c.epsHistory.slice(firstIdx);
    const years = data.length - 1;
    titleSuffix = `${years}-Year History`;
  } else {
    const start = c.eps3yAgo || 0;
    const end = c.eps || 0;
    labels = ['3y ago', '2y ago', '1y ago', 'Current'];
    data = labels.map((_, i) => start + (end - start) * (i / 3));
    titleSuffix = '3-Year Trajectory (interpolated)';
  }

  $('epsChartTitle').textContent = `EPS — ${titleSuffix}`;

  const builder = (targetCtx) => {
    const h = targetCtx.canvas.height || 240;
    const grad = targetCtx.createLinearGradient(0, 0, 0, h);
    grad.addColorStop(0, 'rgba(165, 50, 28,0.3)');
    grad.addColorStop(1, 'rgba(165, 50, 28,0.0)');
    return {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: 'Reported EPS',
          data,
          borderColor: chartColors.accent,
          backgroundColor: grad,
          fill: true,
          tension: 0.35,
          pointBackgroundColor: chartColors.accent,
          pointRadius: 5,
          pointHoverRadius: 7,
          borderWidth: 2.5,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { backgroundColor: chartColors.tooltipBg, borderColor: chartColors.tooltipBorder, borderWidth: 1, titleColor: chartColors.tooltipTitle, bodyColor: chartColors.tooltipBody, padding: 12, cornerRadius: 8 },
        },
        scales: baseScales,
      },
    };
  };

  renderExpandable('eps', ctx, `${c.name} — EPS ${titleSuffix}`, builder);
}

function renderPeers(c) {
  destroyChart('peers');
  const ctx = $('chartPeers').getContext('2d');

  const peers = COMPANIES.filter((p) => p.subsector && p.subsector === c.subsector && p.pe !== null && p.yield !== null && p.name !== c.name);
  const peerData = peers.map((p) => ({ x: p.pe, y: p.yield, name: p.name }));
  const selfData = (c.pe !== null && c.yield !== null) ? [{ x: c.pe, y: c.yield, name: c.name }] : [];

  state.charts.peers = new Chart(ctx, {
    type: 'scatter',
    data: {
      datasets: [
        {
          label: `Subsector Peers (${c.subsector || '—'})`,
          data: peerData,
          backgroundColor: 'rgba(139,92,246,0.5)',
          borderColor: 'rgba(139,92,246,0.8)',
          pointRadius: 6,
          pointHoverRadius: 8,
        },
        {
          label: c.name,
          data: selfData,
          backgroundColor: chartColors.green,
          borderColor: '#0f1419',
          borderWidth: 2,
          pointRadius: 10,
          pointHoverRadius: 12,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: chartColors.text, font: { family: 'Inter Tight' } } },
        tooltip: {
          backgroundColor: chartColors.tooltipBg,
          borderColor: chartColors.tooltipBorder,
          borderWidth: 1,
          titleColor: chartColors.tooltipTitle,
          bodyColor: chartColors.tooltipBody,
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            label: (ctx) => `${ctx.raw.name}: P/E ${ctx.parsed.x.toFixed(1)}, Yield ${ctx.parsed.y.toFixed(1)}%`,
          },
        },
      },
      scales: {
        x: { ...baseScales.x, title: { display: true, text: 'P/E Ratio', color: chartColors.text } },
        y: { ...baseScales.y, title: { display: true, text: 'Dividend Yield (%)', color: chartColors.text } },
      },
    },
  });
}

function renderSensitivity(c) {
  const ctx = $('chartSensitivity').getContext('2d');

  const multiples = [];
  for (let m = 3; m <= 15; m += 1) multiples.push(m);

  // Compute target valuation at each multiple for a given growth override
  const valAt = (growthDelta) =>
    multiples.map((m) => valuation(c, { ...state.inputs, multiple: m, growth: state.inputs.growth + growthDelta }).actual);

  const vMinus10 = valAt(-10);
  const vMinus5  = valAt(-5);
  const vMain    = valAt(0);
  const vPlus5   = valAt(+5);
  const vPlus10  = valAt(+10);
  const price = c.price || 0;

  const builder = () => ({
    type: 'line',
    data: {
      labels: multiples.map((m) => `${m}x`),
      datasets: [
        // Pessimistic band: -10% → -5% (red tint)
        {
          label: '-10% growth',
          data: vMinus10,
          borderColor: 'rgba(165, 50, 28,0.55)',
          backgroundColor: 'transparent',
          borderDash: [2, 4],
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false,
          tension: 0.35,
          order: 5,
        },
        {
          label: '-5% growth',
          data: vMinus5,
          borderColor: 'rgba(165, 50, 28,0.8)',
          backgroundColor: 'rgba(165, 50, 28,0.08)',
          borderDash: [4, 4],
          borderWidth: 1.75,
          pointRadius: 0,
          fill: '-1', // fill to previous dataset (-10%)
          tension: 0.35,
          order: 4,
        },
        // Optimistic band: +5% → +10% (green tint). Put +10% first so +5% can fill to it.
        {
          label: '+10% growth',
          data: vPlus10,
          borderColor: 'rgba(47, 122, 57,0.55)',
          backgroundColor: 'transparent',
          borderDash: [2, 4],
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false,
          tension: 0.35,
          order: 5,
        },
        {
          label: '+5% growth',
          data: vPlus5,
          borderColor: 'rgba(47, 122, 57,0.8)',
          backgroundColor: 'rgba(47, 122, 57,0.1)',
          borderDash: [4, 4],
          borderWidth: 1.75,
          pointRadius: 0,
          fill: '-1', // fill to previous dataset (+10%)
          tension: 0.35,
          order: 4,
        },
        // Main target line on top
        {
          label: 'Target Valuation',
          data: vMain,
          borderColor: chartColors.green,
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 6,
          fill: false,
          tension: 0.35,
          order: 1,
        },
        // Current price reference
        {
          label: 'Current Price',
          data: multiples.map(() => price),
          borderColor: chartColors.amber,
          borderDash: [6, 6],
          pointRadius: 0,
          fill: false,
          borderWidth: 2,
          tension: 0,
          order: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            color: chartColors.text,
            font: { family: 'Inter Tight', size: 11 },
            boxWidth: 14,
            padding: 8,
          },
        },
        tooltip: {
          backgroundColor: chartColors.tooltipBg,
          borderColor: chartColors.tooltipBorder,
          borderWidth: 1,
          titleColor: chartColors.tooltipTitle,
          bodyColor: chartColors.tooltipBody,
          padding: 12,
          cornerRadius: 8,
          mode: 'index',
          intersect: false,
        },
      },
      scales: {
        x: { ...baseScales.x, title: { display: true, text: 'Valuation Multiple', color: chartColors.text } },
        y: { ...baseScales.y, title: { display: true, text: 'Valuation per Share', color: chartColors.text } },
      },
      interaction: { mode: 'index', intersect: false },
    },
  });

  renderExpandable('sensitivity', ctx, `${c.name} — Valuation Sensitivity (Multiple)`, builder);
}

// === Events ===
function bindEvents() {
  $('search').addEventListener('input', (e) => {
    state.search = e.target.value;
    renderCompanyList();
  });

  $('sectorFilter').addEventListener('change', (e) => {
    state.sector = e.target.value;
    renderCompanyList();
  });

  const sliderIds = ['inMultiple', 'inGrowth', 'inUplift', 'inRisk'];
  const numIds = ['outMultiple', 'outGrowth', 'outUplift', 'outRisk'];
  const keys = ['multiple', 'growth', 'uplift', 'risk'];

  sliderIds.forEach((id, i) => {
    $(id).addEventListener('input', (e) => {
      state.inputs[keys[i]] = parseFloat(e.target.value);
      updateInputLabels();
      if (state.selected) render();
    });
  });

  // Editable number boxes — type any value, slider clamps/snaps to its own range
  numIds.forEach((id, i) => {
    const el = $(id);
    const slider = $(sliderIds[i]);
    const commit = () => {
      const v = parseFloat(el.value);
      if (isNaN(v)) return;
      state.inputs[keys[i]] = v;
      // Mirror to slider (it clamps to min/max and snaps to step automatically)
      slider.value = v;
      if (state.selected) render();
    };
    el.addEventListener('input', commit);
    el.addEventListener('change', commit);
    // On blur, re-format to consistent dp
    el.addEventListener('blur', () => updateInputLabels());
    // Enter key blurs (and thus formats)
    el.addEventListener('keydown', (e) => { if (e.key === 'Enter') el.blur(); });
  });

  $('inEPS').addEventListener('input', (e) => {
    const v = parseFloat(e.target.value);
    if (isNaN(v)) return;
    state.inputs.eps = v;
    state.inputs.epsOverridden = state.selected ? (v !== state.selected.eps) : true;
    $('inEPS').classList.toggle('overridden', state.inputs.epsOverridden);
    updateInputLabels();
    if (state.selected) render();
  });

  $('resetEPSBtn').addEventListener('click', () => {
    if (!state.selected) return;
    state.inputs.eps = state.selected.eps;
    state.inputs.epsOverridden = false;
    $('inEPS').value = state.inputs.eps !== null ? state.inputs.eps : '';
    $('inEPS').classList.remove('overridden');
    updateInputLabels();
    render();
  });

  // Asset value input — empty string = EPS mode
  $('inAssetValue').addEventListener('input', (e) => {
    const v = e.target.value.trim();
    state.inputs.assetValue = (v === '' || isNaN(parseFloat(v))) ? null : parseFloat(v);
    updateInputLabels();
    if (state.selected) render();
  });

  $('clearAssetBtn').addEventListener('click', () => {
    state.inputs.assetValue = null;
    $('inAssetValue').value = '';
    updateInputLabels();
    if (state.selected) render();
  });

  // Proposed Action / Research Grade dropdowns
  $('inProposedAction').addEventListener('change', (e) => {
    state.inputs.proposedAction = e.target.value;
  });
  $('inResearchGrade').addEventListener('change', (e) => {
    state.inputs.researchGrade = e.target.value;
  });

  // Probability slider + number
  $('inProbability').addEventListener('input', (e) => {
    state.inputs.probability = parseFloat(e.target.value);
    updateInputLabels();
  });
  $('outProbability').addEventListener('input', (e) => {
    const v = parseFloat(e.target.value);
    if (isNaN(v)) return;
    state.inputs.probability = v;
    $('inProbability').value = v;
  });
  $('outProbability').addEventListener('blur', () => updateInputLabels());

  $('resetBtn').addEventListener('click', () => {
    if (state.selected) selectCompany(state.selected.name);
  });

  // Chart expand buttons — delegated so dynamically-added buttons work too
  document.addEventListener('click', (e) => {
    const btn = e.target.closest && e.target.closest('.expand-btn');
    if (btn) {
      const key = btn.getAttribute('data-chart');
      if (key) openChartModal(key);
      return;
    }
    // Close modal on backdrop / close button
    if (e.target.dataset && e.target.dataset.close === '1') {
      closeChartModal();
    }
  });

  // ESC closes the modal
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeChartModal();
  });

  // Download Word tear sheet
  $('downloadBtn').addEventListener('click', downloadDocx);
}

// === Word document export ===
const DOCX_CDN = 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.min.js';

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      if (existing.dataset.loaded === '1') return resolve();
      existing.addEventListener('load', () => resolve());
      existing.addEventListener('error', reject);
      return;
    }
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => { s.dataset.loaded = '1'; resolve(); };
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

function dataURLToUint8Array(dataURL) {
  const base64 = dataURL.split(',')[1];
  const bin = atob(base64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

function chartImage(key) {
  const ch = state.charts[key];
  if (!ch) return null;
  try {
    const dataUrl = ch.toBase64Image('image/png', 1.0);
    const canvas = ch.canvas;
    const w = canvas.clientWidth || canvas.width;
    const h = canvas.clientHeight || canvas.height;
    return { data: dataURLToUint8Array(dataUrl), w, h };
  } catch (e) {
    console.warn('Could not capture chart', key, e);
    return null;
  }
}

async function downloadDocx() {
  const c = state.selected;
  if (!c) return;
  const btn = $('downloadBtn');
  const sub = $('downloadBtnSub');
  const origSub = sub.textContent;
  btn.disabled = true;
  sub.textContent = 'Building document…';

  try {
    await loadScript(DOCX_CDN);
    if (typeof docx === 'undefined') throw new Error('docx library failed to load');

    const D = docx;
    const cur = currencySym(c.currency);
    const v = valuation(c, state.inputs);

    // ---------- Capture chart images BEFORE building (DOM-bound work) ----------
    const charts = {
      chartTurnover: chartImage('chartTurnover'),
      chartTurnoverGrowth: chartImage('chartTurnoverGrowth'),
      chartPostTax: chartImage('chartPostTax'),
      chartEPSGrowth: chartImage('chartEPSGrowth'),
    };

    // ---------- Helpers ----------
    const text = (s, opts = {}) => new D.TextRun({ text: String(s), ...opts });
    const para = (children, opts = {}) => new D.Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
    const heading = (lvl, s) => new D.Paragraph({ heading: lvl, children: [text(s, { bold: true })], spacing: { before: 240, after: 120 } });

    const imageBlock = (img, targetWidth = 560) => {
      if (!img) return para(text('(chart not available)', { italics: true, color: '888888' }));
      const ratio = img.h && img.w ? img.h / img.w : 0.5;
      const height = Math.max(160, Math.round(targetWidth * ratio));
      return new D.Paragraph({
        alignment: D.AlignmentType.CENTER,
        spacing: { before: 120, after: 240 },
        children: [new D.ImageRun({
          type: 'png',
          data: img.data,
          transformation: { width: targetWidth, height },
          altText: { title: 'Chart', description: 'Chart from Equity and Markets Insight Company Data app', name: 'chart' },
        })],
      });
    };

    const fmtN = (n, dp = 2) => (n === null || n === undefined || isNaN(n)) ? '—' : Number(n).toLocaleString('en-GB', { minimumFractionDigits: dp, maximumFractionDigits: dp });
    const fmtMcap = (n) => (n === null || n === undefined || isNaN(n)) ? '—' : `${cur}${Number(n).toLocaleString('en-GB', { maximumFractionDigits: 1 })}m`;

    // Reverse date format YYMMDD (e.g., 260428)
    const t = new Date();
    const yyMmDd = String(t.getFullYear()).slice(-2)
      + String(t.getMonth() + 1).padStart(2, '0')
      + String(t.getDate()).padStart(2, '0');

    // ---------- Build content ----------
    const children = [];

    // Title block
    children.push(new D.Paragraph({
      heading: D.HeadingLevel.HEADING_1,
      children: [text(`${c.name}`, { bold: true })],
      spacing: { after: 60 },
    }));
    children.push(para([
      text(`${c.tidm || '—'}`, { color: '4F46E5', bold: true }),
      text(`  ·  ${c.industry || ''}${c.subsector ? ' › ' + c.subsector : ''}`, { color: '475569' }),
    ], { spacing: { after: 120 } }));
    children.push(para([text('Date: ', { bold: true }), text(yyMmDd)], { spacing: { after: 60 } }));
    children.push(para([text('Reviewed by: ', { bold: true }), text('')], { spacing: { after: 240 } }));

    // ---------- TEAR SHEET (table + charts) ----------
    const tsRaw = getTearsheet(c);
    if (tsRaw && tsRaw.years && tsRaw.years.length) {
      const ts = trimTearsheet(tsRaw, TEARSHEET_YEARS);

      // Tear sheet table
      const metricOrder = [
        ['turnover', 'Turnover (m)', 'm'],
        ['totalExpenses', 'Total Expenses (m)', 'm'],
        ['turnoverPctChg', 'Turnover % chg', '%'],
        ['operatingProfit', 'Operating Profit (m)', 'm'],
        ['grossMargin', 'Gross Margin (%)', '%'],
        ['preTaxProfit', 'Pre-tax Profit (m)', 'm'],
        ['postTaxProfit', 'Post-tax Profit (m)', 'm'],
        ['sharesInIssue', 'Shares in Issue (m)', 'num2'],
        ['adjustedEPS', 'Adjusted EPS', 'num1'],
        ['reportedEPS', 'Reported EPS', 'num1'],
        ['dividendPerShare', 'Dividend per Share', 'num'],
        ['currentAssets', 'Current Assets (m)', 'm'],
        ['totalAssets', 'Total Assets (m)', 'm'],
        ['totalLiabilities', 'Total Liabilities (m)', 'm'],
        ['netBorrowing', 'Net Borrowing (m)', 'm'],
        ['nav', 'NAV (m)', 'm'],
        ['totalEquity', 'Total Equity (m)', 'm'],
        ['profitOnTurnover', '% Profit on Turnover', '%'],
      ];
      const yearCount = ts.years.length;
      const metricColW = 2600;
      const yearColW = Math.floor((9360 - metricColW) / yearCount);
      const colWidths = [metricColW, ...Array(yearCount).fill(yearColW)];

      const headerCells = [
        new D.TableCell({
          width: { size: metricColW, type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          shading: { fill: 'EEF1F7', type: D.ShadingType.CLEAR },
          children: [para(text('Metric', { bold: true, size: 18 }))],
        }),
        ...ts.years.map((y) => new D.TableCell({
          width: { size: yearColW, type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 60, right: 60 },
          shading: { fill: 'EEF1F7', type: D.ShadingType.CLEAR },
          children: [new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [text(String(y), { bold: true, size: 18 })] })],
        })),
      ];

      const fmtCell = (val, unit) => {
        if (val === null || val === undefined || isNaN(val)) return '—';
        const n = Number(val);
        if (unit === '%') return `${n.toFixed(1)}%`;
        if (unit === 'm') return `${cur}${n.toLocaleString('en-GB', { maximumFractionDigits: 0 })}`;
        if (unit === 'num1') return n.toLocaleString('en-GB', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
        if (unit === 'num2') return n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        return n.toLocaleString('en-GB', { maximumFractionDigits: 2 });
      };

      const dataRows = metricOrder.map(([key, label, unit]) => {
        const arr = ts.metrics[key] || [];
        const cells = [
          new D.TableCell({
            width: { size: metricColW, type: D.WidthType.DXA },
            margins: { top: 50, bottom: 50, left: 80, right: 80 },
            children: [para(text(label, { size: 16 }))],
          }),
          ...arr.map((val) => new D.TableCell({
            width: { size: yearColW, type: D.WidthType.DXA },
            margins: { top: 50, bottom: 50, left: 60, right: 60 },
            children: [new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [text(fmtCell(val, unit), { size: 16 })] })],
          })),
        ];
        return new D.TableRow({ children: cells });
      });

      children.push(new D.Table({
        width: { size: 9360, type: D.WidthType.DXA },
        columnWidths: colWidths,
        rows: [new D.TableRow({ children: headerCells, tableHeader: true }), ...dataRows],
      }));
      children.push(para(text(''), { spacing: { after: 240 } }));

      // 4 charts in 2-up rows (or stacked)
      children.push(heading(D.HeadingLevel.HEADING_3, 'Turnover — Full History'));
      children.push(imageBlock(charts.chartTurnover, 580));

      children.push(heading(D.HeadingLevel.HEADING_3, 'Turnover — Year-on-Year Growth (%)'));
      children.push(imageBlock(charts.chartTurnoverGrowth, 580));

      children.push(heading(D.HeadingLevel.HEADING_3, 'Reported Post-tax Profit + 3-Year Moving Average'));
      children.push(imageBlock(charts.chartPostTax, 580));

      children.push(heading(D.HeadingLevel.HEADING_3, 'Reported EPS — Year-on-Year Growth (%)'));
      children.push(imageBlock(charts.chartEPSGrowth, 580));
    } else {
      children.push(para(text('No tear sheet data available for this company.', { italics: true, color: '94A3B8' })));
    }

    // ---------- STATEMENT (blank section for free-text notes) ----------
    children.push(heading(D.HeadingLevel.HEADING_2, 'Statement'));
    children.push(para(text(' '))); // blank line
    children.push(para(text(' ')));
    children.push(para(text(' ')));

    // ---------- VALUATION TABLE ----------
    children.push(heading(D.HeadingLevel.HEADING_2, 'Valuation'));

    const epsUsed = state.inputs.eps !== null && state.inputs.eps !== undefined ? state.inputs.eps : c.eps;
    const oppPct = v.opportunity !== null ? v.opportunity * 100 : null;
    const fmtSigned = (n, dp, suffix = '') => {
      if (n === null || n === undefined || isNaN(n)) return '—';
      return `${n > 0 ? '+' : ''}${fmtN(n, dp)}${suffix}`;
    };

    const valRows = [
      ['Shares in Issue', fmtN(c.sharesInIssue, 2), 'm'],
      ['EPS', fmtN(epsUsed, 2), state.inputs.epsOverridden ? 'override' : ''],
      ['PE', c.pe !== null ? fmtN(c.pe, 1) : '—', ''],
      ['Market Capitalisation', fmtN(c.marketCap, 1), 'm'],
      ['Growth rate', `${fmtN(state.inputs.growth, 2)}%`, v.mode === 'asset' ? 'ignored (asset mode)' : ''],
      ['Valuation multiple', `${fmtN(state.inputs.multiple, 2)}`, v.mode === 'asset' ? 'ignored (asset mode)' : 'x'],
      ['Dividend percentage', c.yield !== null ? `${fmtN(c.yield, 2)}%` : '—', ''],
      ['Assets per share', v.mode === 'asset' && v.assetsPerShare !== null ? fmtN(v.assetsPerShare, 2) : '', v.mode === 'asset' ? `${cur}${fmtN(state.inputs.assetValue, 1)}m ÷ ${fmtN(c.sharesInIssue, 1)}m × 100` : ''],
      ['Valuation uplift / reduction', `${fmtN(state.inputs.uplift, 2)}%`, ''],
      ['Formula valuation', fmtN(v.formula, 2), ''],
      ['Actual valuation', fmtN(v.actual, 2), ''],
      ['Valued Capitalisation', fmtN(v.valuedCap, 1), 'm'],
      ['Current share price', c.price !== null ? fmtN(c.price, 2) : '—', ''],
      ['Risk Factor', `${fmtN(state.inputs.risk, 0)}%`, ''],
      ['Value Opportunity', oppPct !== null ? fmtSigned(oppPct, 1, '%') : '—', ''],
      ['Proposed action', state.inputs.proposedAction || '', ''],
      ['Research Grade', state.inputs.researchGrade || '', ''],
      ['Probability', state.inputs.probability !== null && state.inputs.probability !== '' ? `${fmtN(state.inputs.probability, 0)}%` : '', ''],
    ];

    // 3-column valuation table: Factor | Figure | Explanation
    const valColWidths = [3200, 2200, 3960];
    const valHeader = new D.TableRow({
      tableHeader: true,
      children: [
        new D.TableCell({
          width: { size: valColWidths[0], type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          shading: { fill: 'EEF1F7', type: D.ShadingType.CLEAR },
          children: [para(text('Factor', { bold: true }))],
        }),
        new D.TableCell({
          width: { size: valColWidths[1], type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          shading: { fill: 'EEF1F7', type: D.ShadingType.CLEAR },
          children: [new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [text('Figure', { bold: true })] })],
        }),
        new D.TableCell({
          width: { size: valColWidths[2], type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          shading: { fill: 'EEF1F7', type: D.ShadingType.CLEAR },
          children: [para(text('Explanation', { bold: true }))],
        }),
      ],
    });
    const valDataRows = valRows.map(([factor, figure, explanation]) => new D.TableRow({
      children: [
        new D.TableCell({
          width: { size: valColWidths[0], type: D.WidthType.DXA },
          margins: { top: 50, bottom: 50, left: 100, right: 100 },
          children: [para(text(factor))],
        }),
        new D.TableCell({
          width: { size: valColWidths[1], type: D.WidthType.DXA },
          margins: { top: 50, bottom: 50, left: 100, right: 100 },
          children: [new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [text(figure, { bold: !!figure })] })],
        }),
        new D.TableCell({
          width: { size: valColWidths[2], type: D.WidthType.DXA },
          margins: { top: 50, bottom: 50, left: 100, right: 100 },
          children: [para(text(explanation, { color: '6B7489' }))],
        }),
      ],
    }));

    children.push(new D.Table({
      width: { size: 9360, type: D.WidthType.DXA },
      columnWidths: valColWidths,
      rows: [valHeader, ...valDataRows],
    }));

    // ---------- Build document ----------
    const doc = new D.Document({
      styles: {
        default: { document: { run: { font: 'Arial', size: 20 } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 36, bold: true, font: 'Arial', color: '0F172A' },
            paragraph: { spacing: { before: 0, after: 120 }, outlineLevel: 0 } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 28, bold: true, font: 'Arial', color: '4F46E5' },
            paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 } },
          { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 22, bold: true, font: 'Arial', color: '0F172A' },
            paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 2 } },
        ],
      },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
        children,
      }],
    });

    const blob = await D.Packer.toBlob(doc);
    const safeName = c.name.replace(/[^A-Za-z0-9 _.-]/g, '').slice(0, 80);
    const filename = `${yyMmDd} ${safeName} — Tear Sheet.docx`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    sub.textContent = 'Downloaded ✓';
    setTimeout(() => { sub.textContent = origSub; }, 2000);
  } catch (err) {
    console.error('Download failed:', err);
    sub.textContent = `Failed: ${err.message || 'unknown error'}`;
    setTimeout(() => { sub.textContent = origSub; }, 4000);
  } finally {
    btn.disabled = false;
  }
}

// === Init ===
function init() {
  populateSectorFilter();
  renderCompanyList();
  bindEvents();
  // Auto-select first for instant visual feedback
  if (COMPANIES.length) selectCompany(COMPANIES[0].name);
}

init();
