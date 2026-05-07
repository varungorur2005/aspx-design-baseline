document.addEventListener('DOMContentLoaded', () => {
  initOppTabs();
  initTableSelection();
  initFeedbackButton();
  initFlyout();
  initOcvModal();
  initGlobalFeedback();
  renderOppView('copilot');
});

// ═══════════════════════════════════════════════
// DATA (realistic ASPX columns from Excel export)
// ═══════════════════════════════════════════════

const copilotColumns = [
  'Tenant Name', 'Tenant ID', 'Tenant Country/Region', 'Tenant Segment',
  'Customer Sub-Segment', 'Industry / Vertical', 'Customer TPID', 'Customer Name',
  'EA Renewal', 'Largest Seat CSP Renewal', 'Dominant SKU Group',
  'Copilot Eligible M365 Seats', 'Copilot Seats Whitespace',
  'Copilot MAU (Licensed)', 'Copilot PAU', 'Copilot Utilization',
  'Adoption Status', 'Free Copilot Chat MAU (Unlicensed)', 'All Copilot MAU',
  'Free to Paid Whitespace', 'All Agents MAU', 'Copilot Opportunity', 'CSP Promos'
];

const e3Columns = [
  'Tenant Name', 'Tenant ID', 'Tenant Country/Region', 'Tenant Segment',
  'Customer Sub-Segment', 'Industry / Vertical', 'TPID', 'Customer Name',
  'EA Renewal', 'Largest Seat CSP Renewal', 'Dominant SKU Group',
  'Total M365 Seat Count', 'Tenant Health', 'Tenant Recommendation',
  'Workload Name', 'MAU', 'PAU', 'Utilization', 'Usage Whitespace',
  'Usage Threshold', 'Workload Health', 'Workload Recommendation'
];

const copilotData = [
  {
    'Tenant Name': 'Fabrikam 123-BE8-8F2', 'Tenant ID': 'D12E2A30-AC07-442C-9A46-EBEEE85BD4D7',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-1271114', 'Customer Name': 'Contoso D4D75557',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '131,185', 'Copilot Seats Whitespace': '130,895',
    'Copilot MAU (Licensed)': '245', 'Copilot PAU': '290', 'Copilot Utilization': '84%',
    'Adoption Status': 'Failure to Thrive/Adopt',
    'Free Copilot Chat MAU (Unlicensed)': '12,650', 'All Copilot MAU': '12,895',
    'Free to Paid Whitespace': '12360', 'All Agents MAU': '414',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-BFE-B06', 'Tenant ID': 'D12E2A30-D256-4F8E-BD03-8D3687987063',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-1762082', 'Customer Name': 'Contoso 70631041',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '122,221', 'Copilot Seats Whitespace': '121,971',
    'Copilot MAU (Licensed)': '235', 'Copilot PAU': '250', 'Copilot Utilization': '94%',
    'Adoption Status': 'Failure to Thrive/Adopt',
    'Free Copilot Chat MAU (Unlicensed)': '3,233', 'All Copilot MAU': '3,468',
    'Free to Paid Whitespace': '2983', 'All Agents MAU': 'Not Available',
    'Copilot Opportunity': 'Acquire', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-CF7-65A', 'Tenant ID': 'D12E2A30-CED6-400C-9B56-FCAC58FF39E8',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-1290446', 'Customer Name': 'Contoso 39E85223',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '106,580', 'Copilot Seats Whitespace': '103,380',
    'Copilot MAU (Licensed)': '3,150', 'Copilot PAU': '3,200', 'Copilot Utilization': '98%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '9,747', 'All Copilot MAU': '12,897',
    'Free to Paid Whitespace': '6547', 'All Agents MAU': '105',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-4E9-8F8', 'Tenant ID': 'D12E2A30-394A-45B3-8B10-53F81F819E3B',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-1255140', 'Customer Name': 'Contoso 9E3B7570',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '98,971', 'Copilot Seats Whitespace': '96,638',
    'Copilot MAU (Licensed)': '2,189', 'Copilot PAU': '2,333', 'Copilot Utilization': '93%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '6,923', 'All Copilot MAU': '9,112',
    'Free to Paid Whitespace': '4590', 'All Agents MAU': '25',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-CDD-C65', 'Tenant ID': 'D12E2A30-309C-4671-A6F8-5B28C58782FF',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-5608676', 'Customer Name': 'Contoso 82FF4338',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '89,724', 'Copilot Seats Whitespace': '87,822',
    'Copilot MAU (Licensed)': '1,980', 'Copilot PAU': '1,902', 'Copilot Utilization': '100%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '8,380', 'All Copilot MAU': '10,360',
    'Free to Paid Whitespace': '6478', 'All Agents MAU': '152',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-9DF-E5E', 'Tenant ID': 'D12E2A30-7A82-4E90-B345-1234ABCD5678',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-1284106', 'Customer Name': 'Contoso DDA52053',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '85,077', 'Copilot Seats Whitespace': '85,052',
    'Copilot MAU (Licensed)': '16', 'Copilot PAU': '25', 'Copilot Utilization': '64%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '5,806', 'All Copilot MAU': '5,822',
    'Free to Paid Whitespace': '5781', 'All Agents MAU': '1',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-A1B-2C3', 'Tenant ID': 'D12E2A30-5F6E-4D7C-8B9A-0A1B2C3D4E5F',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-2045873', 'Customer Name': 'Contoso 7E8F9012',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '74,320', 'Copilot Seats Whitespace': '73,890',
    'Copilot MAU (Licensed)': '410', 'Copilot PAU': '430', 'Copilot Utilization': '95%',
    'Adoption Status': 'Failure to Thrive/Adopt',
    'Free Copilot Chat MAU (Unlicensed)': '4,215', 'All Copilot MAU': '4,625',
    'Free to Paid Whitespace': '3805', 'All Agents MAU': '88',
    'Copilot Opportunity': 'Acquire', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-D4E-5F6', 'Tenant ID': 'D12E2A30-1A2B-4C3D-5E6F-7A8B9C0D1E2F',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-3156984', 'Customer Name': 'Contoso 3A4B5C6D',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '65,890', 'Copilot Seats Whitespace': '64,120',
    'Copilot MAU (Licensed)': '1,720', 'Copilot PAU': '1,770', 'Copilot Utilization': '97%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '7,450', 'All Copilot MAU': '9,170',
    'Free to Paid Whitespace': '5730', 'All Agents MAU': '234',
    'Copilot Opportunity': 'Grow', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-7G8-H9I', 'Tenant ID': 'D12E2A30-2B3C-4D5E-6F7G-8H9I0J1K2L3M',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-4267095', 'Customer Name': 'Contoso 5E6F7G8H',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '52,340', 'Copilot Seats Whitespace': '51,890',
    'Copilot MAU (Licensed)': '425', 'Copilot PAU': '450', 'Copilot Utilization': '94%',
    'Adoption Status': 'Failure to Thrive/Adopt',
    'Free Copilot Chat MAU (Unlicensed)': '3,120', 'All Copilot MAU': '3,545',
    'Free to Paid Whitespace': '2695', 'All Agents MAU': '47',
    'Copilot Opportunity': 'Acquire', 'CSP Promos': 'View'
  },
  {
    'Tenant Name': 'Fabrikam 123-J0K-L1M', 'Tenant ID': 'D12E2A30-3C4D-4E5F-6G7H-8I9J0K1L2M3N',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Not Available', 'Industry / Vertical': 'Not Available',
    'Customer TPID': '-5378106', 'Customer Name': 'Contoso 9A0B1C2D',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'Not Available',
    'Copilot Eligible M365 Seats': '43,210', 'Copilot Seats Whitespace': '42,780',
    'Copilot MAU (Licensed)': '390', 'Copilot PAU': '430', 'Copilot Utilization': '91%',
    'Adoption Status': 'Healthy',
    'Free Copilot Chat MAU (Unlicensed)': '2,890', 'All Copilot MAU': '3,280',
    'Free to Paid Whitespace': '2460', 'All Agents MAU': '19',
    'Copilot Opportunity': 'Monetize', 'CSP Promos': 'View'
  }
];

const e3Data = [
  {
    'Tenant Name': 'Fourth Coffee Ltd.', 'Tenant ID': 'A8F21B30-CC01-4D5E-91AA-12345678ABCD',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Retail',
    'TPID': '-1045231', 'Customer Name': 'Fourth Coffee Inc.',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '03/2027',
    'Dominant SKU Group': 'OE3',
    'Total M365 Seat Count': '485', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Microsoft 365 E3',
    'Workload Name': 'Exchange Online', 'MAU': '412', 'PAU': '485',
    'Utilization': '85%', 'Usage Whitespace': '73',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Maintain usage'
  },
  {
    'Tenant Name': 'Graphic Design Institute', 'Tenant ID': 'B2C33A10-DD02-4E6F-82BB-98765432DCBA',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Education',
    'TPID': '-1098776', 'Customer Name': 'Graphic Design Corp.',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '06/2026',
    'Dominant SKU Group': 'OE1',
    'Total M365 Seat Count': '320', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Microsoft Teams', 'MAU': '89', 'PAU': '320',
    'Utilization': '28%', 'Usage Whitespace': '231',
    'Usage Threshold': 'Below', 'Workload Health': 'Unhealthy',
    'Workload Recommendation': 'Drive Teams adoption'
  },
  {
    'Tenant Name': 'Nod Publishers', 'Tenant ID': 'C3D44B20-EE03-4F7G-73CC-11223344EEFF',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Small Business', 'Industry / Vertical': 'Media',
    'TPID': '-1156432', 'Customer Name': 'Nod Publishing Group',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '09/2026',
    'Dominant SKU Group': 'BB',
    'Total M365 Seat Count': '145', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Business Premium',
    'Workload Name': 'SharePoint Online', 'MAU': '118', 'PAU': '145',
    'Utilization': '81%', 'Usage Whitespace': '27',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Upsell ready'
  },
  {
    'Tenant Name': 'Wide World Importers', 'Tenant ID': 'E4F55C40-FF04-4H8I-94DD-55667788AABB',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Manufacturing',
    'TPID': '-1267543', 'Customer Name': 'Wide World Importers LLC',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '12/2026',
    'Dominant SKU Group': 'OE3',
    'Total M365 Seat Count': '680', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Microsoft 365 E5',
    'Workload Name': 'Microsoft Teams', 'MAU': '620', 'PAU': '680',
    'Utilization': '91%', 'Usage Whitespace': '60',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Upsell ready'
  },
  {
    'Tenant Name': 'Tailspin Toys', 'Tenant ID': 'F5G66D50-GG05-4I9J-A5EE-66778899BBCC',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Small Business', 'Industry / Vertical': 'Retail',
    'TPID': '-1378654', 'Customer Name': 'Tailspin Toys Inc.',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '03/2027',
    'Dominant SKU Group': 'BB',
    'Total M365 Seat Count': '210', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage',
    'Workload Name': 'Exchange Online', 'MAU': '98', 'PAU': '210',
    'Utilization': '47%', 'Usage Whitespace': '112',
    'Usage Threshold': 'Below', 'Workload Health': 'Unhealthy',
    'Workload Recommendation': 'Drive Exchange adoption'
  },
  {
    'Tenant Name': 'Wingtip Toys', 'Tenant ID': 'G6H77E60-HH06-4J0K-B6FF-778899AABBDD',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Entertainment',
    'TPID': '-1489765', 'Customer Name': 'Wingtip Toys Corp.',
    'EA Renewal': '06/2027', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'OE3',
    'Total M365 Seat Count': '950', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Microsoft 365 E5 Security',
    'Workload Name': 'SharePoint Online', 'MAU': '780', 'PAU': '950',
    'Utilization': '82%', 'Usage Whitespace': '170',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Cross-sell security'
  },
  {
    'Tenant Name': 'Litware Inc.', 'Tenant ID': 'H7I88F70-II07-4K1L-C7GG-8899AABBCCEE',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Technology',
    'TPID': '-1590876', 'Customer Name': 'Litware Inc.',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '09/2026',
    'Dominant SKU Group': 'OE1',
    'Total M365 Seat Count': '410', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Microsoft Teams', 'MAU': '156', 'PAU': '410',
    'Utilization': '38%', 'Usage Whitespace': '254',
    'Usage Threshold': 'Below', 'Workload Health': 'Unhealthy',
    'Workload Recommendation': 'Drive Teams adoption'
  },
  {
    'Tenant Name': 'Adventure Works', 'Tenant ID': 'I8J99G80-JJ08-4L2M-D8HH-99AABBCCDDFF',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Corporate', 'Industry / Vertical': 'Outdoor Recreation',
    'TPID': '-1601987', 'Customer Name': 'Adventure Works Cycles',
    'EA Renewal': '03/2028', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'OE3',
    'Total M365 Seat Count': '1,250', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Microsoft 365 E5',
    'Workload Name': 'Exchange Online', 'MAU': '1,100', 'PAU': '1,250',
    'Utilization': '88%', 'Usage Whitespace': '150',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Upsell ready'
  },
  {
    'Tenant Name': 'Proseware Inc.', 'Tenant ID': 'J9K00H90-KK09-4M3N-E9II-AABBCCDDEEFF',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMC',
    'Customer Sub-Segment': 'Small Business', 'Industry / Vertical': 'Software',
    'TPID': '-1712098', 'Customer Name': 'Proseware Inc.',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '06/2027',
    'Dominant SKU Group': 'BB',
    'Total M365 Seat Count': '175', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Business Premium',
    'Workload Name': 'Microsoft Teams', 'MAU': '160', 'PAU': '175',
    'Utilization': '91%', 'Usage Whitespace': '15',
    'Usage Threshold': 'Above', 'Workload Health': 'Healthy',
    'Workload Recommendation': 'Cross-sell Copilot'
  }
];

const oppData = {
  copilot: {
    kpis: [
      { value: '6,625', label: 'Tenants' },
      { value: '3%', label: 'Seat Penetration' },
      { value: '83%', label: 'Paid Utilization' },
      { value: '1,781', label: 'Acquire' },
      { value: '1,710', label: 'Monetize' },
      { value: '868', label: 'Grow' },
      { value: '0', label: 'E7 Upsell-ready' },
      { value: '$1,430,000', label: 'Potential Earnings' },
    ],
    columns: copilotColumns,
    rows: copilotData
  },
  e3: {
    kpis: [
      { value: '1,240', label: 'Tenants' },
      { value: '54%', label: 'Paid Utilization' },
      { value: '28,400', label: 'Usage Whitespace' },
      { value: '312', label: 'Upsell Ready' },
      { value: '680', label: 'Unhealthy' },
      { value: '92', label: 'Incentive Eligible' },
      { value: '$420,000', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e3Data
  },
  e5: {
    kpis: [
      { value: '890', label: 'Tenants' },
      { value: '61%', label: 'Paid Utilization' },
      { value: '19,200', label: 'Usage Whitespace' },
      { value: '198', label: 'Upsell Ready' },
      { value: '412', label: 'Unhealthy' },
      { value: '78', label: 'At Risk' },
      { value: '$580,000', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e3Data
  },
  e5exp: {
    kpis: [
      { value: '520', label: 'Tenants' },
      { value: '78%', label: 'Paid Utilization' },
      { value: '10,400', label: 'Usage Whitespace' },
      { value: '245', label: 'Healthy' },
      { value: '180', label: 'Unhealthy' },
      { value: '42', label: 'Incentive Eligible' },
      { value: '$310,000', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e3Data
  }
};

let currentTab = 'copilot';
let selectedRow = null;
let selectedRowData = null;

// ═══════════════════════════════════════════════
// OPPORTUNITY TABS
// ═══════════════════════════════════════════════

function initOppTabs() {
  document.querySelectorAll('.opp-tabs-bar .opp-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.opp-tabs-bar .opp-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentTab = tab.dataset.opp;
      if (oppData[currentTab]) {
        renderOppView(currentTab);
      }
    });
  });
}

function renderOppView(opp) {
  const data = oppData[opp];
  if (!data) return;

  // KPIs
  const kpiRow = document.getElementById('kpiRow');
  kpiRow.innerHTML = data.kpis.map(k =>
    `<div class="kpi-tile">
      <svg class="kpi-info-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1Zm.75 10.5h-1.5V7h1.5v4.5ZM8 6a.75.75 0 1 1 0-1.5A.75.75 0 0 1 8 6Z"/></svg>
      <span class="kpi-value">${k.value}</span>
      <span class="kpi-label">${k.label}</span>
    </div>`
  ).join('');

  // Table header — "Give Feedback" as first column
  const head = document.getElementById('tableHead');
  head.innerHTML = '<th class="col-feedback">Give Feedback</th>' + data.columns.map(c => `<th>${c}</th>`).join('');

  // Table body — each row gets a Give Feedback dropdown
  const body = document.getElementById('tableBody');
  body.innerHTML = data.rows.map((row, idx) => {
    const cells = data.columns.map(col => {
      let val = row[col] || '\u2014';
      if (col === 'Tenant Name') {
        val = `<a href="#" class="tenant-link" data-row="${idx}" data-flyout="tenant-details">${val}</a>`;
      } else if (col === 'Copilot Opportunity') {
        const cls = val.toLowerCase();
        val = `<a href="#" class="tenant-link" data-row="${idx}" data-flyout="opportunity-details"><span class="status-badge ${cls}">${val}</span></a>`;
      } else if (col === 'Adoption Status') {
        const cls = val.includes('Healthy') ? 'healthy' : val.includes('Failure') ? 'unhealthy' : '';
        val = cls ? `<span class="status-badge ${cls}">${val}</span>` : val;
      } else if (col === 'Tenant Health' || col === 'Workload Health') {
        const cls = val.toLowerCase().replace(/\s/g, '-');
        val = `<span class="status-badge ${cls}">${val}</span>`;
      } else if (col === 'All Copilot MAU' && val !== 'Not Available' && val !== '\u2014') {
        val = `<a href="#" class="tenant-link" data-row="${idx}" data-flyout="copilot-mau">${val}</a>`;
      } else if (col === 'All Agents MAU' && val !== 'Not Available' && val !== '\u2014') {
        val = `<a href="#" class="tenant-link" data-row="${idx}" data-flyout="agents-mau">${val}</a>`;
      } else if (col === 'CSP Promos' && val === 'View') {
        val = `<a href="#" class="tenant-link" data-row="${idx}" data-flyout="csp-promos">View</a>`;
      }
      return `<td>${val}</td>`;
    }).join('');
    const feedbackCell = `<td class="col-feedback"><div class="row-feedback-wrapper"><button class="row-feedback-btn" data-row="${idx}"><svg width="14" height="14" viewBox="0 0 20 20" fill="currentColor"><path d="M7 5.5a2.5 2.5 0 1 1 5 0 2.5 2.5 0 0 1-5 0Zm2.5-1.5a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3ZM5.5 10A1.5 1.5 0 0 0 4 11.5v.5c0 2.39 2.33 4 5.5 4 .88 0 1.69-.1 2.4-.3a3.48 3.48 0 0 1-.4-.74c-.6.16-1.28.24-2 .24-2.78 0-4.5-1.36-4.5-3.2v-.5a.5.5 0 0 1 .5-.5h6.14c.16-.36.37-.7.62-1H5.5Z"/><path d="M14.5 11a2.5 2.5 0 0 0-2.08 3.88l-.37 1.42a.4.4 0 0 0 .54.47l1.56-.65A2.5 2.5 0 1 0 14.5 11Zm-1.5 2.5a1.5 1.5 0 1 1 1.65 1.49.5.5 0 0 0-.38.13l-.69.29.17-.65a.5.5 0 0 0-.06-.38A1.49 1.49 0 0 1 13 13.5Z"/></svg><svg class="row-feedback-chevron" width="10" height="10" viewBox="0 0 10 10" fill="currentColor"><path d="M2.5 3.5L5 6l2.5-2.5"/></svg></button><div class="row-feedback-dropdown hidden"><div class="feedback-dropdown-header">What feedback do you have?</div><button class="feedback-option" data-reason="incorrect" data-row="${idx}">Incorrect data</button><button class="feedback-option" data-reason="missing" data-row="${idx}">Missing data</button><button class="feedback-option" data-reason="outdated" data-row="${idx}">Outdated data</button><button class="feedback-option" data-reason="mismatch" data-row="${idx}">Data doesn't match my records</button><button class="feedback-option" data-reason="other" data-row="${idx}">Other</button></div></div></td>`;
    return `<tr data-row="${idx}">${feedbackCell}${cells}</tr>`;
  }).join('');

  renderCharts(opp);
}

// ═══════════════════════════════════════════════
// CHARTS (Chart.js)
// ═══════════════════════════════════════════════

function renderCharts(opp) {
  const container = document.getElementById('chartsRow');
  if (!container) return;

  if (opp !== 'copilot') {
    container.innerHTML = '';
    return;
  }

  // Destroy existing charts to prevent memory leaks
  Object.keys(Chart.instances).forEach(key => {
    Chart.instances[key].destroy();
  });

  container.innerHTML = `
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">Copilot Tenant Penetration</div><div class="chart-card-subtitle">As of 2026-05-07, 01:42 p.m.</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chartTenantPen"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">Copilot Opportunities</div><div class="chart-card-subtitle">As of 2026-05-07, 01:42 p.m.</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chartOpportunities"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">Copilot Adoption by State</div><div class="chart-card-subtitle">As of 2026-05-07, 01:42 p.m.</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chartAdoption"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">Copilot Usage by License</div><div class="chart-card-subtitle">As of 2026-05-07, 01:42 p.m.</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chartUsage"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">MCI Potential Earnings</div><div class="chart-card-subtitle">As of 2026-05-07, 01:42 p.m.</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chartMci"></canvas></div>
    </div>
  `;

  new Chart(document.getElementById('chartTenantPen'), {
    type: 'doughnut',
    data: {
      labels: ['Has Copilot Licenses', 'No Copilot Licenses'],
      datasets: [{ data: [3487, 3138], backgroundColor: ['#0078d4', '#f7941d'], borderWidth: 0 }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } } }
  });

  new Chart(document.getElementById('chartOpportunities'), {
    type: 'doughnut',
    data: {
      labels: ['Acquire', 'Monetize', 'Grow', 'Other'],
      datasets: [{ data: [2235, 1781, 868, 1710], backgroundColor: ['#1b3a5c', '#0078d4', '#107c10', '#f7941d'], borderWidth: 0 }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } } }
  });

  new Chart(document.getElementById('chartAdoption'), {
    type: 'bar',
    data: {
      labels: ['Other', 'Failure to\nThrive/Adopt', 'Healthy', 'Starting'],
      datasets: [{ data: [3200, 2800, 2500, 1200], backgroundColor: '#0078d4', borderRadius: 2, barThickness: 40 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'x',
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, ticks: { font: { size: 9 } }, grid: { color: '#f0f0f0' } }, x: { ticks: { font: { size: 9 }, maxRotation: 45, minRotation: 45 }, grid: { display: false } } }
    }
  });

  new Chart(document.getElementById('chartUsage'), {
    type: 'line',
    data: {
      labels: ['Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May'],
      datasets: [
        { label: 'Unlicensed MAU', data: [2800, 2900, 3000, 3100, 3200, 3300, 3350, 3400, 3450, 3480, 3490, 3500], borderColor: '#c62828', backgroundColor: '#c62828', borderWidth: 1.5, pointRadius: 2, tension: 0.3 },
        { label: 'AI MAU', data: [600, 650, 700, 750, 800, 850, 900, 920, 950, 980, 1000, 1050], borderColor: '#0078d4', backgroundColor: '#0078d4', borderWidth: 1.5, pointRadius: 2, tension: 0.3 },
        { label: 'Licensed MAU', data: [300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520], borderColor: '#107c10', backgroundColor: '#107c10', borderWidth: 1.5, pointRadius: 2, tension: 0.3 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } } },
      scales: { y: { beginAtZero: true, ticks: { font: { size: 9 } }, grid: { color: '#f0f0f0' } }, x: { ticks: { font: { size: 9 } }, grid: { display: false } } }
    }
  });

  new Chart(document.getElementById('chartMci'), {
    type: 'doughnut',
    data: {
      labels: ['Copilot + Power Platform', 'Copilot + Deployment Accelerator (S)', 'Copilot + Deployment Accelerator (XS)', 'Copilot + PoC', 'Other'],
      datasets: [{ data: [35, 19, 14, 13, 19], backgroundColor: ['#1b3a5c', '#0078d4', '#4fc3f7', '#107c10', '#f7941d'], borderWidth: 0 }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 8 } } } } }
  });
}

// IN-ROW GIVE FEEDBACK (dropdown per row)
// ═══════════════════════════════════════════════

function initFeedbackButton() {
  const table = document.getElementById('dataTable');
  if (!table) return;

  // Toggle dropdown on feedback button click (event delegation)
  table.addEventListener('click', (e) => {
    const btn = e.target.closest('.row-feedback-btn');
    const opt = e.target.closest('.feedback-option');

    if (btn) {
      e.stopPropagation();
      const wrapper = btn.closest('.row-feedback-wrapper');
      const dd = wrapper.querySelector('.row-feedback-dropdown');
      const wasOpen = !dd.classList.contains('hidden');

      // Close all other dropdowns first
      document.querySelectorAll('.row-feedback-dropdown').forEach(d => d.classList.add('hidden'));
      document.querySelectorAll('.row-feedback-btn').forEach(b => b.classList.remove('active'));

      if (!wasOpen) {
        dd.classList.remove('hidden');
        btn.classList.add('active');
      }
    } else if (opt) {
      e.stopPropagation();
      const reason = opt.dataset.reason;
      const rowIdx = parseInt(opt.dataset.row);
      const rowData = oppData[currentTab].rows[rowIdx];

      // Close dropdown
      document.querySelectorAll('.row-feedback-dropdown').forEach(d => d.classList.add('hidden'));
      document.querySelectorAll('.row-feedback-btn').forEach(b => b.classList.remove('active'));

      openOcvModal(reason, rowData, 'table');
    }
  });

  // Close dropdowns on outside click
  document.addEventListener('click', () => {
    document.querySelectorAll('.row-feedback-dropdown').forEach(d => d.classList.add('hidden'));
    document.querySelectorAll('.row-feedback-btn').forEach(b => b.classList.remove('active'));
  });
}

function initTableSelection() {}

// ═══════════════════════════════════════════════
// FLYOUT PANEL (5 types from screenshots 2-6)
// ═══════════════════════════════════════════════

function initFlyout() {
  const overlay = document.getElementById('flyoutOverlay');
  const closeBtn = document.getElementById('flyoutClose');

  // Click on links in the table (using event delegation on document)
  document.addEventListener('click', (e) => {
    const tenantLink = e.target.closest('.tenant-link');
    if (!tenantLink) return;
    e.preventDefault();
    const rowIdx = parseInt(tenantLink.dataset.row);
    const flyoutType = tenantLink.dataset.flyout;
    const rowData = oppData[currentTab].rows[rowIdx];
    openFlyout(flyoutType, rowData);
  });

  closeBtn.addEventListener('click', closeFlyout);
  overlay.addEventListener('click', closeFlyout);

  // Thumbs up/down in the flyout
  document.getElementById('flyoutPanel').querySelectorAll('.flyout-thumb').forEach(thumb => {
    thumb.addEventListener('click', () => {
      const vote = thumb.dataset.vote;
      // Visual feedback on the thumb
      document.querySelectorAll('.flyout-thumb').forEach(t => t.classList.remove('voted'));
      thumb.classList.add('voted');
      // Open OCV modal after a brief delay for visual feedback
      setTimeout(() => {
        const reason = vote === 'up' ? 'positive' : 'negative';
        closeFlyout();
        openOcvModal(reason, currentFlyoutData, 'flyout');
      }, 300);
    });
  });
}

let currentFlyoutData = null;

function openFlyout(type, rowData) {
  currentFlyoutData = rowData;
  const overlay = document.getElementById('flyoutOverlay');
  const panel = document.getElementById('flyoutPanel');
  const title = document.getElementById('flyoutTitle');
  const body = document.getElementById('flyoutBody');

  // Reset thumb state
  panel.querySelectorAll('.flyout-thumb').forEach(t => t.classList.remove('voted'));

  overlay.classList.remove('hidden');
  panel.classList.remove('hidden');

  if (type === 'tenant-details') {
    title.textContent = rowData['Tenant Name'];
    body.innerHTML = renderTenantDetailsFlyout(rowData);
  } else if (type === 'copilot-mau') {
    title.textContent = 'All Copilot MAU';
    body.innerHTML = renderCopilotMauFlyout(rowData);
  } else if (type === 'agents-mau') {
    title.textContent = 'All Agents MAU';
    body.innerHTML = renderAgentsMauFlyout(rowData);
  } else if (type === 'opportunity-details') {
    title.textContent = 'Opportunity Details';
    body.innerHTML = renderOpportunityDetailsFlyout(rowData);
  } else if (type === 'csp-promos') {
    title.textContent = (rowData['Customer Name'] || rowData['Tenant Name']) + ' | CSP Promotions';
    body.innerHTML = renderCspPromosFlyout(rowData);
  }
}

function closeFlyout() {
  document.getElementById('flyoutOverlay').classList.add('hidden');
  document.getElementById('flyoutPanel').classList.add('hidden');
  document.querySelectorAll('.flyout-thumb').forEach(t => t.classList.remove('voted'));
}

// Screenshot 2: Tenant Details flyout
function renderTenantDetailsFlyout(data) {
  return `
    <table class="flyout-info-table">
      <tr><td>Tenant ID:</td><td>${data['Tenant ID'] || '—'}</td></tr>
      <tr><td>TPID:</td><td>${data['Customer TPID'] || data['TPID'] || '—'}</td></tr>
      <tr><td>Claims:</td><td>—</td></tr>
      <tr><td>Claimed workloads:</td><td>—</td></tr>
    </table>
    <h4 class="flyout-section-title">Subscriptions</h4>
    <div class="flyout-toolbar">
      <button class="flyout-toolbar-btn">↓ Download</button>
      <span class="flyout-toolbar-spacer"></span>
      <button class="flyout-toolbar-btn">⊞ Filter</button>
      <input class="flyout-search" type="text" placeholder="Search" />
    </div>
    <table class="flyout-info-table flyout-subs-table">
      <thead>
        <tr><td><strong>Sku Name</strong></td><td><strong>Status</strong></td><td><strong>Current State End Date</strong></td><td><strong>Included Quantity</strong></td></tr>
      </thead>
      <tbody>
        <tr><td>WINDOWS 10/11 ENTERPRISE E3</td><td>Active</td><td>10/31/2027</td><td>4994</td></tr>
        <tr><td>PROJECT PLAN 5</td><td>Active</td><td>10/31/2027</td><td>3</td></tr>
        <tr><td>PROJECT PLAN 3</td><td>Active</td><td>10/31/2027</td><td>91</td></tr>
        <tr><td>MICROSOFT TEAMS ROOMS PRO</td><td>Active</td><td>10/31/2027</td><td>50</td></tr>
        <tr><td>MICROSOFT TEAMS SHARED DEVICES</td><td>Active</td><td>10/31/2027</td><td>25</td></tr>
        <tr><td>POWER AUTOMATE PROCESS</td><td>Active</td><td>10/31/2027</td><td>1</td></tr>
        <tr><td>MICROSOFT VIVA EMPLOYEE COMMUNICATIONS AND COMMUNITIES</td><td>Active</td><td>10/31/2027</td><td>1</td></tr>
        <tr><td>MICROSOFT TEAMS PHONE STANDARD</td><td>Active</td><td>10/31/2027</td><td>13584</td></tr>
      </tbody>
    </table>
  `;
}

// Screenshot 3: All Copilot MAU flyout
function renderCopilotMauFlyout(data) {
  return `
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">The total Copilot MAU encompassing licensed and unlicensed users. Additionally, included below is a breakdown of Copilot usage. Unless specified in the metric, MAU can include licensed and unlicensed users. <a href="#" style="color:#0078d4;">Learn more ↗</a></p>
    <table class="flyout-info-table">
      <tr><td><strong>Metric Name</strong></td><td style="text-align:right;"><strong>MAU</strong></td></tr>
      <tr><td>All Copilot MAU</td><td style="text-align:right;">${data['All Copilot MAU'] || '—'}</td></tr>
      <tr><td>Free Copilot MAU (Unlicensed)</td><td style="text-align:right;">${data['Free Copilot Chat MAU (Unlicensed)'] || '—'}</td></tr>
      <tr><td>Copilot MAU (Licensed)</td><td style="text-align:right;">${data['Copilot MAU (Licensed)'] || '—'}</td></tr>
      <tr><td>Copilot in Teams</td><td style="text-align:right;">248</td></tr>
      <tr><td>Copilot in Outlook</td><td style="text-align:right;">4,850</td></tr>
      <tr><td>Copilot in Word</td><td style="text-align:right;">88</td></tr>
      <tr><td>Copilot in Excel</td><td style="text-align:right;">43</td></tr>
      <tr><td>Copilot in Powerpoint</td><td style="text-align:right;">67</td></tr>
      <tr><td>Copilot in Engage</td><td style="text-align:right;color:#999;">Not available</td></tr>
      <tr><td>Copilot in OneNote</td><td style="text-align:right;">4</td></tr>
      <tr><td>Sales Copilot (CRM)</td><td style="text-align:right;color:#999;">Not available</td></tr>
    </table>
  `;
}

// Screenshot 4: All Agents MAU flyout
function renderAgentsMauFlyout(data) {
  return `
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">All Agents MAU provides the total usage of agents all up across any application and user. The table below gives a more granular breakdown of licensed, unlicensed, and agent MAU by extension/application type. <a href="#" style="color:#0078d4;">Learn more ↗</a></p>
    <table class="flyout-info-table">
      <tr><td><strong>Agent Extension Type</strong></td><td style="text-align:right;"><strong>Licensed Agent MAU</strong></td><td style="text-align:right;"><strong>Unlicensed Agent MAU</strong></td><td style="text-align:right;"><strong>Total</strong></td></tr>
      <tr><td>All Up</td><td style="text-align:right;">13</td><td style="text-align:right;">401</td><td style="text-align:right;">414</td></tr>
      <tr><td>Agents in Copilot Chat</td><td style="text-align:right;">13</td><td style="text-align:right;">401</td><td style="text-align:right;">414</td></tr>
      <tr><td>Connectors</td><td style="text-align:right;">0</td><td style="text-align:right;">401</td><td style="text-align:right;">0</td></tr>
      <tr><td>Plugin (Actions)</td><td style="text-align:right;">0</td><td style="text-align:right;">0</td><td style="text-align:right;">0</td></tr>
      <tr><td>Custom Engine Copilot</td><td style="text-align:right;">4</td><td style="text-align:right;">278</td><td style="text-align:right;">282</td></tr>
    </table>
    <h4 class="flyout-section-title">Data Definitions</h4>
    <div style="font-size:12px;color:#424242;line-height:1.7;">
      <p><strong>Agents in Copilot Chat:</strong> Agents that appear within the Copilot Chat experience. These can be pre-built Microsoft agents (analyst, researcher, etc.) as well as custom built and/or connectors.</p>
      <p style="margin-top:10px;"><strong>Connectors:</strong> Graph connectors enable users to extend Microsoft Copilot by integrating external data sources into Microsoft Graph, enhancing Copilot's ability to reason over your enterprise content and provide more comprehensive responses. Connectors can be 3rd party as well (e.g., ServiceNow).</p>
      <p style="margin-top:10px;"><strong>Plugin (Actions):</strong> Actions for an agent to interact with other systems to read or write information in near real-time (e.g., create a ticket).</p>
      <p style="margin-top:10px;"><strong>Custom Engine Copilot:</strong> These are fully custom Copilots built using Copilot Studio or other frameworks.</p>
    </div>
  `;
}

// Screenshot 5: Opportunity Details flyout
function renderOpportunityDetailsFlyout(data) {
  const opportunity = data['Copilot Opportunity'] || '—';
  const eligibleSeats = data['Copilot Eligible M365 Seats'] || '—';
  const whitespace = data['Copilot Seats Whitespace'] || '—';
  const freeMau = data['Free Copilot Chat MAU (Unlicensed)'] || '0';
  const pau = data['Copilot PAU'] || '0';
  const allMau = data['All Copilot MAU'] || '0';

  // Calculate penetration percentages
  const copilotPen = pau !== '0' && eligibleSeats !== '—' ?
    Math.round((parseInt(pau.replace(/,/g, '')) / parseInt(eligibleSeats.replace(/,/g, ''))) * 100) + '%' : '0%';
  const freePen = freeMau !== '0' && eligibleSeats !== '—' ?
    Math.round((parseInt(freeMau.replace(/,/g, '')) / parseInt(eligibleSeats.replace(/,/g, ''))) * 100) + '%' : '0%';
  const freeWhitespace = parseInt((freeMau || '0').replace(/,/g, '')) + parseInt(whitespace.replace(/,/g, ''));

  let reasoning = '';
  if (opportunity === 'Monetize') {
    reasoning = 'Penetrated with Copilot Chat MAU but not Copilot paid licenses';
  } else if (opportunity === 'Acquire') {
    reasoning = 'Low Chat MAU penetration and low M365 paid license penetration';
  } else if (opportunity === 'Grow') {
    reasoning = 'Notable penetration in Copilot Chat and paid Copilot licenses';
  }

  let recommendation = '';
  if (opportunity === 'Monetize') {
    recommendation = 'Upsell free Chat to paid (Agents and/or M365 Copilot)';
  } else if (opportunity === 'Acquire') {
    recommendation = 'Grow use of Copilot Free Chat to build engagement';
  } else if (opportunity === 'Grow') {
    recommendation = 'Increase paid Copilot licenses further';
  }

  return `
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">The Copilot Opportunity value provides direction on what next steps to take with the customer. Included below are key penetration metrics and AI/ML insights that explain why the customer is categorized as acquire, monetize, or grow. <a href="#" style="color:#0078d4;">Learn more ↗</a></p>
    <table class="flyout-info-table">
      <tr><td><strong>Opportunity:</strong></td><td><span class="status-badge ${opportunity.toLowerCase()}">${opportunity}</span></td></tr>
      <tr><td><strong>M365 Copilot Penetration:</strong></td><td>${copilotPen}</td></tr>
      <tr><td><strong>Free Copilot Chat Penetration:</strong></td><td>${freePen}</td></tr>
      <tr><td><strong>Copilot Eligible Seats:</strong></td><td>${eligibleSeats}</td></tr>
      <tr><td><strong>M365 Copilot Whitespace:</strong></td><td>${whitespace}</td></tr>
      <tr><td><strong>Free Copilot Chat Whitespace:</strong></td><td>${freeWhitespace.toLocaleString()}</td></tr>
      <tr><td><strong>Opportunity Reasoning:</strong></td><td>${reasoning}</td></tr>
    </table>
    <div class="flyout-recommendation-box">
      <div class="flyout-rec-header">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="#9a6700"><path d="M8 1a1.5 1.5 0 0 1 1.5 1.5v.7a5 5 0 0 1 2.8 2.8h.7a1.5 1.5 0 0 1 0 3h-.7a5 5 0 0 1-2.8 2.8v.7a1.5 1.5 0 0 1-3 0v-.7a5 5 0 0 1-2.8-2.8H3a1.5 1.5 0 0 1 0-3h.7A5 5 0 0 1 6.5 3.2v-.7A1.5 1.5 0 0 1 8 1z"/></svg>
        <strong>Recommendation</strong>
      </div>
      <p>${recommendation}</p>
    </div>
    <h4 class="flyout-section-title">AI-Insights Hub <span class="preview-badge">Preview</span></h4>
    <div style="font-size:13px;color:#242424;line-height:1.7;">
      <p><strong style="color:#0078d4;">Expansion Ready Propensity:</strong> High</p>
      <p style="margin-top:8px;"><strong>Expansion Ready Reasoning</strong></p>
      <p style="color:#616161;">Account's recent earnings call shows plans to invest in AI and automation, indicating an opportunity for M365 Copilot to help streamline workflows and reduce manual effort. Account has over 90% of Teams monthly usage from Meetings. M365 Copilot can automate meeting recaps and capture next steps.</p>
      <p style="margin-top:12px;"><strong style="color:#0078d4;">Adoption Ready Propensity:</strong> Medium</p>
      <p style="margin-top:8px;"><strong>Adoption Ready Reasoning</strong></p>
      <p style="color:#616161;">The tenant has a number of files viewed of 19,522, and a total number of files created of 4,281. 38% of users within the tenant have a high likelihood to become engaged.</p>
    </div>
  `;
}

// Screenshot 6: CSP Promotions flyout
function renderCspPromosFlyout(data) {
  return `
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">List of all active New Commerce Cloud Solution Provider (CSP) Promotions related to this solution area. Active promotions do not necessarily meet the eligibility criteria for this tenant. Visit the Pricing workspace for a complete list of eligible promotions across products. <a href="#" style="color:#0078d4;">Learn More ↗</a></p>
    <table class="flyout-info-table flyout-promos-table">
      <thead>
        <tr><td><strong>Promotion Name</strong></td><td><strong>Discount Type</strong></td><td><strong>Discount Rate</strong></td><td><strong>End Date</strong></td></tr>
      </thead>
      <tbody>
        <tr><td>Scale Up with Microsoft 365 Copilot 20% offer</td><td>Percent off</td><td>20%</td><td>06/30/2026</td></tr>
        <tr><td>Getting Started with Microsoft 365 Copilot 15% offer</td><td>Percent off</td><td>15%</td><td>06/30/2026</td></tr>
        <tr><td>Updated: M365 Purview Suite promo offer for M365 Copilot</td><td>Percent off</td><td>50%</td><td>07/01/2026</td></tr>
        <tr><td>Introductory offer: Up to 15% off Microsoft 365 Copilot Business</td><td>Percent off</td><td>15%</td><td>06/30/2026</td></tr>
        <tr><td>Bundle and save: Up to 35% off Microsoft 365 Business Standard and Microsoft 365 Copilot Business</td><td>Percent off</td><td>35%</td><td>06/30/2026</td></tr>
        <tr><td>Limited time offer: Microsoft 365 Copilot for All—40% offer</td><td>Percent off</td><td>40%</td><td>06/30/2026</td></tr>
        <tr><td>15% off Microsoft 365 E7 triennial subscription, 300-9,999 licenses</td><td>Percent off</td><td>15%</td><td>12/31/2026</td></tr>
      </tbody>
    </table>
  `;
}

// ═══════════════════════════════════════════════
// OCV FEEDBACK MODAL (Two-step: classify → form)
// ═══════════════════════════════════════════════

function initOcvModal() {
  // Close buttons
  document.getElementById('ocvClose').addEventListener('click', closeOcvModal);
  document.getElementById('ocvClose2').addEventListener('click', closeOcvModal);
  document.getElementById('ocvCancelStep1').addEventListener('click', closeOcvModal);
  document.getElementById('ocvCancelStep2').addEventListener('click', closeOcvModal);
  document.getElementById('ocvOverlay').addEventListener('click', closeOcvModal);
  document.getElementById('ocvSubmit').addEventListener('click', submitFeedback);

  // Back button
  document.getElementById('ocvBackBtn').addEventListener('click', () => {
    document.getElementById('ocvStep2').classList.add('hidden');
    document.getElementById('ocvStep1').classList.remove('hidden');
  });

  // Classification buttons
  document.querySelectorAll('.ocv-classify-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const type = btn.getAttribute('data-type');
      showOcvStep2(type);
    });
  });
}

let ocvFeedbackContext = null; // Stores context for table/flyout feedback

function openOcvModal(reason, rowData, source) {
  const modal = document.getElementById('ocvModal');
  const overlay = document.getElementById('ocvOverlay');

  overlay.classList.remove('hidden');
  modal.classList.remove('hidden');

  // Store context for table/flyout sources
  ocvFeedbackContext = { reason, rowData, source };

  if (source === 'general') {
    // Show step 1 (classification)
    document.getElementById('ocvStep1').classList.remove('hidden');
    document.getElementById('ocvStep2').classList.add('hidden');
  } else {
    // For table/flyout feedback, go directly to step 2
    const typeLabel = reason === 'positive' ? 'compliment' : 'problem';
    showOcvStep2(typeLabel, rowData, source);
  }
}

function showOcvStep2(type, rowData, source) {
  document.getElementById('ocvStep1').classList.add('hidden');
  document.getElementById('ocvStep2').classList.remove('hidden');

  const questionLabel = document.getElementById('ocvQuestionLabel');
  const contextSection = document.getElementById('ocvContext');
  const contextDetails = document.getElementById('ocvContextDetails');
  const columnsSection = document.getElementById('ocvColumnsSection');
  const columnsGrid = document.getElementById('ocvColumnsGrid');

  // Set question based on type
  if (type === 'compliment') {
    questionLabel.innerHTML = 'What did you like? <span class="required">*</span>';
  } else if (type === 'problem') {
    questionLabel.innerHTML = 'What went wrong? <span class="required">*</span>';
  } else if (type === 'suggestion') {
    questionLabel.innerHTML = 'What would you like to see improved? <span class="required">*</span>';
  }

  // Use stored context if not passed directly
  const ctx = ocvFeedbackContext || {};
  const rd = rowData || ctx.rowData;
  const src = source || ctx.source;

  // Show context for table/flyout sources
  if (src && src !== 'general' && rd) {
    contextSection.classList.remove('hidden');
    const tenantName = rd['Tenant Name'] || '—';
    contextDetails.innerHTML = `
      <div class="ctx-item"><span class="ctx-key">Source:</span><span class="ctx-val">${src === 'flyout' ? 'Flyout Panel' : 'Data Table'}</span></div>
      <div class="ctx-item"><span class="ctx-key">Tab:</span><span class="ctx-val">${currentTab.charAt(0).toUpperCase() + currentTab.slice(1)} Opportunities</span></div>
      <div class="ctx-item"><span class="ctx-key">Tenant:</span><span class="ctx-val">${tenantName}</span></div>
    `;

    // Show columns grid
    columnsSection.classList.remove('hidden');
    const columns = oppData[currentTab].columns;
    const displayCols = columns.slice(0, 12);
    columnsGrid.innerHTML = displayCols.map(col => {
      const val = rd[col] || '—';
      return `<div class="ocv-col-item"><div class="col-label">${col}</div><div class="col-value">${val}</div></div>`;
    }).join('');
  } else {
    contextSection.classList.add('hidden');
    columnsSection.classList.add('hidden');
  }

  // Reset form
  document.getElementById('ocvComments').value = '';
  document.querySelectorAll('input[name="ocvContact"]').forEach(r => r.checked = false);
}

function closeOcvModal() {
  document.getElementById('ocvModal').classList.add('hidden');
  document.getElementById('ocvOverlay').classList.add('hidden');
  // Reset to step 1
  document.getElementById('ocvStep1').classList.remove('hidden');
  document.getElementById('ocvStep2').classList.add('hidden');
  ocvFeedbackContext = null;
}

function submitFeedback() {
  closeOcvModal();
  showToast();
}

function showToast() {
  const toast = document.getElementById('ocvToast');
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 4000);
}

// ═══════════════════════════════════════════════
// GLOBAL FEEDBACK (top-right nav button)
// ═══════════════════════════════════════════════

function initGlobalFeedback() {
  document.getElementById('globalFeedbackBtn').addEventListener('click', () => {
    openOcvModal('general', null, 'general');
  });
}
