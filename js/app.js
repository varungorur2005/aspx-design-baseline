const state = { page: 'growth', growth: 'copilot', mci: 'copilot', referral: 'fasttrack', search: '' };

const tenants = [
  ['Fabrikam 123-aee-6e9', 'Korea', 'SME&C SMB', 'BP', '32', '32', 'Acquire', '$1,600'],
  ['Fabrikam 123-85f-501', 'Luxembourg', 'SME&C SMB', 'BS', '28', '28', 'Acquire', '$1,400'],
  ['Fabrikam 123-453-4e5', 'United Kingdom', 'SME&C SMB', 'BS', '28', '28', 'Grow', '$1,400'],
  ['Fabrikam 123-62f-ae4', 'Germany', 'SME&C SMB', 'BP', '18', '18', 'Acquire', '$900'],
  ['Fabrikam 123-7bc-114', 'Hong Kong', 'SMME Corporate', 'E3', '540', '390', 'Monetize', '$7,500'],
  ['Fabrikam 123-c41-2df', 'United States', 'SMME Corporate', 'E5', '1,220', '810', 'Monetize', '$18,200'],
  ['Fabrikam 123-4c8-0f3', 'Canada', 'SME&C SMB', 'BP', '210', '165', 'Acquire', '$3,900'],
  ['Fabrikam 123-8ef-6c2', 'Netherlands', 'SMME Corporate', 'E3', '780', '505', 'Grow', '$11,250']
];

const referrals = [
  ['FR-10482', 'Contoso Retail', 'Microsoft field', 'Accepted', 'In progress', 'Outreach pending', '2026-05-27'],
  ['FR-10477', 'Fabrikam Finance', 'Partner Center', 'Pending review', 'Not started', 'Pending', '2026-05-25'],
  ['FR-10451', 'Northwind Health', 'FastTrack', 'Accepted', 'Engaged', 'Active', '2026-05-18'],
  ['FR-10412', 'Adventure Works', 'Microsoft field', 'Declined', 'Closed', 'Complete', '2026-05-08'],
  ['FR-10398', 'Tailspin Toys', 'Partner Center', 'Accepted', 'In progress', 'Active', '2026-05-02']
];

const campaigns = [
  ['SDC-2048', 'Contoso Retail', 'Americas', 'Fabrikam Partner', '998311', 'New', 'Adele Vance', 'Alex Wilber', 'Defender, Sentinel', 'Active', 'Security Sales', '2026-05-29'],
  ['SDC-2037', 'Northwind Health', 'EMEA', 'Contoso Services', '871204', 'Outreach', 'Megan Bowen', 'Miriam Graham', 'Purview', 'Active', 'Usage signal', '2026-05-24'],
  ['SDC-2019', 'Adventure Works', 'Americas', 'Fabrikam Partner', '998311', 'Funding (Blocked)', 'Adele Vance', 'Alex Wilber', 'Defender for Cloud', 'On hold', 'Security Sales', '2026-05-19'],
  ['SDC-1996', 'Tailspin Toys', 'APAC', 'Wide World Importers', '661920', 'Engaged', 'Nestor Wilke', 'Patti Fernandez', 'Sentinel', 'Active', 'Partner nominated', '2026-05-11']
];

const growthConfig = {
  copilot: {
    kpis: [['Acquire', '1,284'], ['Grow', '746'], ['Monetize', '392'], ['Expansion ready', '628'], ['Adoption ready', '481'], ['Potential earnings', '$2.8M']],
    charts: [['Tenant Penetration', ['No licenses', 'Licensed', 'Free chat'], [68, 44, 59]], ['Opportunities', ['Acquire', 'Grow', 'Monetize'], [72, 51, 34]], ['Adoption', ['Healthy', 'Starting', 'At risk'], [61, 48, 22]], ['Usage', ['Teams', 'Outlook', 'Word', 'Excel'], [76, 63, 58, 41]]]
  },
  e3: { kpis: [['Tenants', '2,118'], ['Healthy', '1,326'], ['At risk', '424'], ['Usage opportunities', '869'], ['Potential earnings', '$1.9M']], charts: [['Utilization', ['High', 'Medium', 'Low'], [64, 48, 27]], ['Whitespace', ['Teams', 'SharePoint', 'Exchange'], [43, 69, 35]], ['Tenant health', ['Healthy', 'At risk', 'Unhealthy'], [72, 31, 18]], ['Recommendations', ['Adoption', 'Expansion', 'Renewal'], [59, 46, 33]]] },
  e5: { kpis: [['Tenants', '1,740'], ['Healthy', '1,021'], ['At risk', '388'], ['Usage opportunities', '711'], ['Potential earnings', '$2.4M']], charts: [['Utilization', ['High', 'Medium', 'Low'], [58, 49, 29]], ['Whitespace', ['Security', 'Compliance', 'Voice'], [71, 53, 38]], ['Tenant health', ['Healthy', 'At risk', 'Unhealthy'], [65, 34, 21]], ['Recommendations', ['Adoption', 'Expansion', 'Renewal'], [62, 44, 30]]] },
  e5exp: { kpis: [['Tenants', '936'], ['Eligible', '614'], ['At risk', '179'], ['Expansion opportunities', '438'], ['Potential earnings', '$1.6M']], charts: [['Utilization', ['High', 'Medium', 'Low'], [56, 45, 24]], ['Whitespace', ['Security', 'Compliance', 'Analytics'], [67, 51, 36]], ['Tenant health', ['Healthy', 'At risk', 'Unhealthy'], [69, 29, 17]], ['Potential earnings', ['Security', 'Compliance', 'Voice'], [78, 54, 41]]] },
  cloud: { kpis: [['Tenants', '1,052'], ['Windows 365 whitespace', '18,420'], ['CPOR eligible', '482'], ['MCI eligible', '291'], ['Potential earnings', '$984K']], charts: [['Cloud PC utilization', ['High', 'Medium', 'Low'], [69, 42, 19]], ['Seat whitespace', ['Enterprise', 'Frontline', 'Business'], [77, 58, 36]], ['Tenant health', ['Healthy', 'At risk', 'Unhealthy'], [63, 32, 15]], ['Recommendations', ['Trial', 'Expand', 'Optimize'], [52, 46, 31]]] },
  tenants: { kpis: [['Total tenants', '5,846'], ['Enterprise', '712'], ['SMME Corporate', '1,536'], ['SME&C SMB', '3,598']], charts: [['Tenant segments', ['Enterprise', 'SMME', 'SMB'], [22, 48, 81]], ['Dominant SKU', ['E5', 'E3', 'BP', 'BS'], [44, 67, 74, 39]]] }
};

const growthHeaders = ['Tenant Name', 'Country/Region', 'Tenant Segment', 'Dominant SKU Group', 'Eligible M365 Seats', 'Seat Whitespace', 'Copilot Opportunity', 'Potential Earnings'];
const referralHeaders = ['Referral ID', 'Customer Name', 'Referral Source', 'Review Status', 'Engagement Status', 'Outreach Status', 'Created Date'];
const campaignHeaders = ['Campaign ID', 'Customer Name', 'Customer Area', 'Partner Name', 'Partner ID PLA', 'State', 'Partner DRI', 'Microsoft DRI', 'Targeted Workloads', 'Campaign Status', 'Campaign Source', 'Created Date'];

function kpiGrid(items, clinic = false) {
  return `<div class="kpi-grid${clinic ? ' clinic' : ''}">${items.map(([label, value], i) => `
    <article class="kpi-card"><span class="kpi-label">${label}</span>
    <button class="kpi-info" title="${label} information">ⓘ</button>
    ${clinic || i < 3 ? `<button class="kpi-value link" data-kpi="${label}">${value}</button>` : `<span class="kpi-value">${value}</span>`}</article>`).join('')}</div>`;
}

function barChart(title, labels, values) {
  const max = Math.max(...values);
  return `<article class="chart-card"><div class="chart-head"><div><h3>${title}</h3><p>As of 05/29/2026, 01:42 PM</p></div><button class="chart-menu" aria-label="Chart options">•••</button></div>
    <div class="bar-chart">${labels.map((label, i) => `<div class="bar-item"><span class="bar-value">${values[i]}</span><div class="bar" style="height:${Math.max(4, values[i] / max * 75)}%"></div><span class="bar-label" title="${label}">${label}</span></div>`).join('')}</div></article>`;
}

function donutChart(title, labels) {
  return `<article class="chart-card"><div class="chart-head"><div><h3>${title}</h3><p>As of 05/29/2026, 01:42 PM</p></div><button class="chart-menu">•••</button></div>
    <div class="donut-layout"><div class="donut"></div><div class="legend">${labels.map(x => `<span>${x}</span>`).join('')}</div></div></article>`;
}

function toolbar(type) {
  return `<div class="command-bar"><div class="command-group"><button class="command">⇩ Download</button><button class="command">↗ Data Dictionary</button></div>
    <div class="command-group"><button class="command change-columns">▥ Change columns</button><button class="command filter">▽ Filter</button><input class="search" type="search" placeholder="Search" aria-label="Search"></div></div>`;
}

function table(headers, rows, type) {
  const query = state.search.toLowerCase();
  const visible = rows.filter(row => row.join(' ').toLowerCase().includes(query));
  return `<div class="result-status">${visible.length} results</div>${toolbar(type)}<div class="grid-wrap"><table><thead><tr>${headers.map(h => `<th>${h} ↕</th>`).join('')}</tr></thead>
    <tbody>${visible.map(row => `<tr>${row.map((cell, i) => `<td>${i === 0 ? `<button class="cell-link" data-detail="${cell}">${cell}</button>` : statusCell(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
    <div class="pagination"><span>Showing 1 - ${visible.length} of ${visible.length}</span><button>‹</button><button class="active">1</button><button>›</button><button>25 results⌄</button></div>`;
}

function statusCell(value) {
  if (['Active', 'Accepted', 'Healthy', 'Complete'].includes(value)) return `<span class="badge success">${value}</span>`;
  if (/Pending|Blocked|On hold|At risk/.test(value)) return `<span class="badge warning">${value}</span>`;
  return value;
}

function renderGrowth() {
  const cfg = growthConfig[state.growth];
  document.getElementById('growthContent').innerHTML = `${kpiGrid(cfg.kpis)}<div class="charts-grid four">${cfg.charts.map(c => barChart(...c)).join('')}</div>${table(growthHeaders, tenants, 'growth')}`;
  wireContent();
}

function concurrentCaps() {
  return `<section class="concurrent"><h2>Concurrent Cap Status</h2><div class="mci-card-grid">
    <article class="mci-card"><span>Copilot CSP</span><strong>72%</strong><span class="badge success">On track</span></article>
    <article class="mci-card"><span>Copilot EA</span><strong>86%</strong><span class="badge warning">Near cap</span></article>
    <article class="mci-card"><span>Security</span><strong>48%</strong><span class="badge success">On track</span></article>
  </div></section>`;
}

function renderMci() {
  const content = document.getElementById('mciContent');
  if (state.mci !== 'copilot') {
    content.innerHTML = `${concurrentCaps()}<div class="empty-state"><div><div class="empty-icon">ⓘ</div><h2>Data unavailable</h2><p>Performance metrics are coming soon for this solution area.</p></div></div>`;
    return;
  }
  content.innerHTML = `${kpiGrid([['Eligible engagements', '684'], ['Active engagements', '291'], ['Completed', '176'], ['Potential earnings', '$1.3M']])}
    <div class="charts-grid">${barChart('Engagement status', ['Eligible', 'Claimed', 'Active', 'Complete'], [78, 54, 43, 31])}${barChart('Performance by month', ['Feb', 'Mar', 'Apr', 'May'], [41, 49, 64, 73])}</div>
    ${concurrentCaps()}${table(['Customer Name', 'Engagement Name', 'Status', 'Workshop Date', 'Partner Center ID', 'Potential Earnings'], [
      ['Contoso Retail', 'Copilot Value Discovery', 'Active', '2026-06-12', '998311', '$18,000'],
      ['Fabrikam Finance', 'Copilot Studio Workshop', 'Pending', '2026-06-18', '998311', '$12,500'],
      ['Northwind Health', 'Copilot Adoption Accelerator', 'Complete', '2026-05-14', '871204', '$9,800']
    ], 'mci')}`;
  wireContent();
}

function renderReferrals() {
  if (state.referral === 'fasttrack') {
    document.getElementById('referralContent').innerHTML = `${kpiGrid([['Total', '248'], ['Accepted', '164'], ['Acceptance Rate', '66%'], ['Pending Review', '38'], ['Pending Outreach', '29'], ['Pending Engagement', '17']])}
      <div class="charts-grid">${donutChart('Referral Sources', ['Microsoft field', 'Partner Center', 'FastTrack', 'Other'])}${donutChart('Referral Review Status', ['Accepted', 'Pending', 'Declined', 'Expired'])}${donutChart('Referral Outreach Status', ['Pending', 'Contacted', 'Complete', 'Overdue'])}${donutChart('Referral Engagement Status', ['Not started', 'In progress', 'Engaged', 'Closed'])}</div>
      ${table(referralHeaders, referrals, 'referrals')}`;
  } else {
    document.getElementById('referralContent').innerHTML = `${kpiGrid([['New', '14'], ['In Progress', '37'], ['⚠ Blocked', '6']], true)}
      <div class="charts-grid">${donutChart('Campaigns by State', ['New', 'Outreach', 'Engaged', 'Funding'])}${donutChart('Customers by Area', ['Americas', 'EMEA', 'APAC', 'Other'])}</div>
      ${table(campaignHeaders, campaigns, 'clinic')}<div class="as-of">Last refreshed: 05/29/2026, 01:42 PM</div>`;
  }
  wireContent();
}

function renderPage() {
  document.querySelectorAll('.page').forEach(x => x.classList.remove('active'));
  document.getElementById(`${state.page}Page`).classList.add('active');
  document.querySelectorAll('.rail-link').forEach(x => x.classList.toggle('active', x.dataset.page === state.page));
  const titles = { growth: ['Growth opportunities', 'AI Business Solutions & Security'], mci: ['MCI performance', 'AI Business Solutions & Security'], referrals: ['Referrals and Deal Clinics', 'FastTrack Referrals and Security Deal Clinic'] };
  document.getElementById('pageTitle').textContent = titles[state.page][0];
  document.getElementById('pageSubtitle').textContent = titles[state.page][1];
  if (state.page === 'growth') renderGrowth();
  if (state.page === 'mci') renderMci();
  if (state.page === 'referrals') renderReferrals();
}

function openPanel(title, body, showFooter = true) {
  document.getElementById('panelTitle').textContent = title;
  document.getElementById('panelBody').innerHTML = body;
  document.getElementById('panel').classList.remove('hidden');
  document.getElementById('scrim').classList.remove('hidden');
  document.querySelector('#panel footer').classList.toggle('hidden', !showFooter);
}

function closePanel() {
  document.getElementById('panel').classList.add('hidden');
  document.getElementById('scrim').classList.add('hidden');
}

function wireContent() {
  const search = document.querySelector('.page.active .search');
  if (search) {
    search.value = state.search;
    search.addEventListener('input', e => { state.search = e.target.value; state.page === 'growth' ? renderGrowth() : state.page === 'mci' ? renderMci() : renderReferrals(); });
    search.focus();
  }
  document.querySelectorAll('.page.active .filter').forEach(btn => btn.addEventListener('click', () => openPanel('Filter', `<div class="filter-list"><label>Status<select><option>All</option><option>Active</option><option>Pending</option></select></label><label>Country/Region<select><option>All</option><option>United States</option><option>Germany</option><option>United Kingdom</option></select></label><label>Customer segment<select><option>All</option><option>Enterprise</option><option>SMME Corporate</option><option>SME&C SMB</option></select></label></div>`)));
  document.querySelectorAll('.page.active .change-columns').forEach(btn => btn.addEventListener('click', () => openPanel('Change columns', `<div class="filter-list">${growthHeaders.map((h, i) => `<label><span><input type="checkbox" ${i < 6 ? 'checked' : ''}> ${h}</span></label>`).join('')}</div>`)));
  document.querySelectorAll('.page.active [data-detail]').forEach(btn => btn.addEventListener('click', () => openPanel(btn.dataset.detail, `<dl class="detail-list"><dt>Record</dt><dd>${btn.dataset.detail}</dd><dt>Status</dt><dd>Active</dd><dt>Last updated</dt><dd>05/29/2026</dd><dt>Owner</dt><dd>Fabrikam Partner</dd></dl>`, false)));
  document.querySelectorAll('.page.active [data-kpi]').forEach(btn => btn.addEventListener('click', () => openPanel(`${btn.dataset.kpi} details`, `<p>The grid has been filtered to records represented by this KPI.</p>`, false)));
}

document.querySelectorAll('.rail-link').forEach(btn => btn.addEventListener('click', () => { state.page = btn.dataset.page; state.search = ''; renderPage(); }));
document.querySelectorAll('[data-growth]').forEach(btn => btn.addEventListener('click', () => { state.growth = btn.dataset.growth; document.querySelectorAll('[data-growth]').forEach(x => x.classList.toggle('active', x === btn)); state.search = ''; renderGrowth(); }));
document.querySelectorAll('[data-mci]').forEach(btn => btn.addEventListener('click', () => { state.mci = btn.dataset.mci; document.querySelectorAll('[data-mci]').forEach(x => x.classList.toggle('active', x === btn)); state.search = ''; renderMci(); }));
document.querySelectorAll('[data-referral]').forEach(btn => btn.addEventListener('click', () => { state.referral = btn.dataset.referral; document.querySelectorAll('[data-referral]').forEach(x => x.classList.toggle('active', x === btn)); state.search = ''; renderReferrals(); }));
document.getElementById('insightsButton').addEventListener('click', () => openPanel('Insights', '<p>The AI Business Solutions & Security Insights experience equips marketing, sales, and customer success partner teams with actionable leads, data insights, and incentive information across the Microsoft customer lifecycle.</p>', false));
['panelClose', 'panelCancel', 'panelApply', 'scrim'].forEach(id => document.getElementById(id).addEventListener('click', closePanel));
renderPage();
