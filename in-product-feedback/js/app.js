document.addEventListener('DOMContentLoaded', () => {
  initSidebarNavigation();
  initOppTabs();
  initMciTabs();
  initToolbarControls();
  initTableSelection();
  initFeedbackButton();
  initFlyout();
  initOcvModal();
  initGlobalFeedback();
  initFasttrackControls();
  initGlobalFeedbackDelegation();
  renderCurrentPage();
});

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// DATA (realistic ASPX columns from Excel export)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

const copilotColumns = [
  'Tenant Name', 'Tenant ID', 'Tenant Country/Region', 'Tenant Segment',
  'Customer Sub-Segment', 'Industry Vertical', 'Customer TPID', 'Customer Name',
  'EA Renewal', 'Largest Seat CSP Renewal', 'Dominant SKU Group',
  'Copilot Eligible M365 Seats', 'Copilot Seats Whitespace', 'Copilot MAU (Licensed)',
  'Copilot PAU', 'Copilot Utilization', 'Adoption Status',
  'Free Copilot Chat MAU (Unlicensed)', 'All Copilot MAU', 'Free to Paid Whitespace',
  'All Agents MAU', 'Copilot Opportunity', 'Potential Earnings', 'CSP Promos'
];

const e3Columns = [
  'Tenant Name', 'Tenant ID', 'Tenant Country/Region', 'Tenant Segment',
  'Customer Sub-Segment', 'Industry Vertical', 'Customer TPID', 'Customer Name',
  'EA Renewal', 'Largest Seat CSP Renewal', 'Dominant SKU Group',
  'Total M365 Seat Count', 'Tenant Health', 'Tenant Recommendation',
  'Workload Name', 'MAU', 'PAU', 'Utilization', 'Whitespace',
  'Usage Threshold', 'Tipping Point', 'Workload Health', 'Usage Opportunity'
];

const allCopilotColumns = [
  'Tenant Name', 'Tenant ID', 'Tenant Domain', 'Tenant Country/Region', 'Tenant Segment',
  'Customer TPID', 'Customer Name', 'EA Renewal', 'Largest Seat CSP Renewal', 'Earliest CSP Renewal',
  'Copilot Eligible M365 Seats', 'Copilot Seats Whitespace', 'Copilot MAU (Licensed)', 'Copilot PAU',
  'Copilot Utilization', 'Adoption Status', 'Free Copilot Chat MAU (Unlicensed)', 'All Copilot MAU',
  'Free to Paid Whitespace', 'Copilot Opportunity', 'Potential Earnings', 'MCI Potential Earnings',
  'Starting HWM', 'Prior HWM', 'Opportunity Units', 'Rate Per Unit', 'Max Comp Units',
  'CPOR Potential Earnings', 'CPOR Association', 'CSP Association', 'UPOR Association', 'Is Charity Offer',
  'Dominant SKU Group', 'CSP Tier', 'Partner Name', 'Partner Center ID (MPN)', 'CSP Partner Name',
  'Reseller Name', 'Customer Sub-Segment', 'Industry Vertical', 'MCI Engagement Name',
  'Data Security Maturity', 'Utilization MIP P2', 'Utilization IRM', 'Copilot in Teams MAU',
  'Copilot in Outlook MAU', 'Copilot in Word MAU', 'Copilot in Excel MAU', 'Copilot in Powerpoint MAU',
  'Copilot in Engage MAU', 'Copilot in OneNote MAU', 'Sales Copilot (CRM) MAU', 'All Agents MAU',
  'Copilot Chat Agents MAU', 'Connector Agents MAU', 'Plugins Agents MAU', 'Custom Engine Agents MAU',
  'Advanced Deployment Guide Name', 'Advanced Deployment Guide Link', 'Resource Name', 'Resource Link',
  'CSP Promos', 'CSP Promotion Names', 'MCI Eligibility'
];

const allE3Columns = [
  'Tenant Name', 'Tenant ID', 'Tenant Domain', 'Tenant Country/Region', 'Tenant Segment',
  'Customer TPID', 'Customer Name', 'EA Renewal', 'Largest Seat CSP Renewal', 'Earliest CSP Renewal',
  'Dominant SKU Group', 'Total M365 Seat Count', 'Tenant Health', 'Tenant Recommendation',
  'Workload Name', 'MAU', 'PAU', 'Utilization', 'Whitespace', 'Usage Threshold', 'Tipping Point',
  'Workload Health', 'Usage Opportunity', 'Customer Country/Region', 'Customer Segment', 'Starting HWM',
  'Prior HWM', 'Opportunity Units', 'Rate Per Unit', 'Max Comp Units', 'MCI Potential Earnings',
  'Potential Earnings', 'CPOR Association', 'CSP Association', 'Claim Status', 'UPOR Association',
  'Is Charity Offer', 'CSP Tier', 'Reseller Name', 'Partner Name', 'Partner Center ID (MPN)',
  'CSP Partner Name', 'Customer Sub-Segment', 'Industry Vertical', 'MCI Engagement Name',
  'Claim ID', 'Claimed By Others', 'MCI Eligibility'
];

const columnAliases = {
  'Customer TPID': ['Customer TPID', 'TPID'],
  'Whitespace': ['Whitespace', 'Usage Whitespace']
};

const tabDisplayNames = {
  copilot: 'Copilot Opportunities',
  e3: 'E3 Opportunities',
  e5: 'E5 Opportunities',
  e5exp: 'E5 Expansion'
};

const mciTabDisplayNames = {
  copilotPower: 'Copilot + Power',
  secureProductivity: 'Secure Productivity',
  cloudEndpoints: 'Cloud Endpoints'
};

const chartPalette = ['#0078d4', '#1b3a5c', '#107c10', '#f7941d', '#4fc3f7', '#9c27b0'];
const chartSubtitle = 'As of 2026-05-29, 01:42 p.m.';

const filterConfig = {
  copilot: [
    { field: 'Tenant Country/Region' },
    { field: 'Tenant Segment' },
    { field: 'Copilot Opportunity', options: ['Acquire', 'Monetize', 'Grow', 'Other'] },
    { field: 'Adoption Status' },
    { field: 'Dominant SKU Group' }
  ],
  e3: [
    { field: 'Tenant Country/Region' },
    { field: 'Tenant Segment' },
    { field: 'Tenant Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Name' }
  ],
  e5: [
    { field: 'Tenant Country/Region' },
    { field: 'Tenant Segment' },
    { field: 'Tenant Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Name' }
  ],
  e5exp: [
    { field: 'Tenant Country/Region' },
    { field: 'Tenant Segment' },
    { field: 'Tenant Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Health', options: ['Healthy', 'Unhealthy', 'At Risk'] },
    { field: 'Workload Name' }
  ]
};

const copilotData = [
  {
    'Tenant Name': 'Fabrikam 123-aee-6e9', 'Tenant ID': 'D12E2A30-0002-4355-a055-ed07f88bc72a',
    'Tenant Country/Region': 'Korea', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'N/A',
    'Customer TPID': '-186553416', 'Customer Name': 'Contoso aeecc72a',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '32', 'Copilot Seats Whitespace': '32',
    'Copilot MAU (Licensed)': 'Not Available', 'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '4', 'All Copilot MAU': '4',
    'Free to Paid Whitespace': '4', 'All Agents MAU': '2', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': 'Not Available', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-85f-501', 'Tenant ID': 'D12E2A30-0005-414e-a92e-fbbc96ca4137',
    'Tenant Country/Region': 'Luxembourg', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'N/A',
    'Customer TPID': '-166826878', 'Customer Name': 'Contoso 85fd4137',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Copilot Eligible M365 Seats': '28', 'Copilot Seats Whitespace': '28',
    'Copilot MAU (Licensed)': 'Not Available', 'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '3', 'All Copilot MAU': '3',
    'Free to Paid Whitespace': '3', 'All Agents MAU': 'Not Available', 'Copilot Opportunity': 'Other',
    'Potential Earnings': 'Not Available', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-453-4e5', 'Tenant ID': 'D12E2A30-0008-4c40-b3c3-1947eecba3f6',
    'Tenant Country/Region': 'United Kingdom', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'N/A',
    'Customer TPID': '-101316272', 'Customer Name': 'Contoso 4539a3f6',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Copilot Eligible M365 Seats': '28', 'Copilot Seats Whitespace': '28',
    'Copilot MAU (Licensed)': 'Not Available', 'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '1', 'All Copilot MAU': '1',
    'Free to Paid Whitespace': '1', 'All Agents MAU': '0', 'Copilot Opportunity': 'Other',
    'Potential Earnings': 'Not Available', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-62f-ae4', 'Tenant ID': 'D12E2A30-000a-4f46-bb1f-b39d6b6a4dab',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Process Manufacturing',
    'Customer TPID': '-27195194', 'Customer Name': 'Contoso 62fa4dab',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '18', 'Copilot Seats Whitespace': '18',
    'Copilot MAU (Licensed)': 'Not Available', 'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '1', 'All Copilot MAU': '1',
    'Free to Paid Whitespace': '1', 'All Agents MAU': 'Not Available', 'Copilot Opportunity': 'Other',
    'Potential Earnings': 'Not Available', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-7bc-114', 'Tenant ID': 'D12E2A30-0011-41bc-a6f0-0c7eac021145',
    'Tenant Country/Region': 'Hong Kong', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Financial Services',
    'Customer TPID': '-95442018', 'Customer Name': 'Contoso 7bc01145',
    'EA Renewal': '2027-03-31', 'Largest Seat CSP Renewal': '2026-11-15',
    'Dominant SKU Group': 'E3', 'Copilot Eligible M365 Seats': '540', 'Copilot Seats Whitespace': '390',
    'Copilot MAU (Licensed)': '150', 'Copilot PAU': '182', 'Copilot Utilization': '82%',
    'Adoption Status': 'Starting', 'Free Copilot Chat MAU (Unlicensed)': '96', 'All Copilot MAU': '246',
    'Free to Paid Whitespace': '96', 'All Agents MAU': '34', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$7,500', 'CSP Promos': 'Available (8)'
  },
  {
    'Tenant Name': 'Fabrikam 123-9d2-7aa', 'Tenant ID': 'D12E2A30-0016-45de-b6b2-9b6d4f277aa1',
    'Tenant Country/Region': 'Romania', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Retail',
    'Customer TPID': '-84522731', 'Customer Name': 'Contoso 9d277aa1',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-10-01',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '86', 'Copilot Seats Whitespace': '86',
    'Copilot MAU (Licensed)': 'Not Available', 'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '12', 'All Copilot MAU': '12',
    'Free to Paid Whitespace': '12', 'All Agents MAU': '5', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': '$2,400', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-c41-2df', 'Tenant ID': 'D12E2A30-0019-4f21-b4af-3314df2c41ff',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Healthcare',
    'Customer TPID': '-44211876', 'Customer Name': 'Contoso c412c41f',
    'EA Renewal': '2026-12-31', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Copilot Eligible M365 Seats': '1,220', 'Copilot Seats Whitespace': '810',
    'Copilot MAU (Licensed)': '340', 'Copilot PAU': '410', 'Copilot Utilization': '83%',
    'Adoption Status': 'Healthy', 'Free Copilot Chat MAU (Unlicensed)': '215', 'All Copilot MAU': '555',
    'Free to Paid Whitespace': '215', 'All Agents MAU': '148', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$18,200', 'CSP Promos': 'Available (5)'
  },
  {
    'Tenant Name': 'Fabrikam 123-5a7-b91', 'Tenant ID': 'D12E2A30-0022-4a7e-9175-22bc5a7b9100',
    'Tenant Country/Region': 'Belgium', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Business Services',
    'Customer TPID': '-31884509', 'Customer Name': 'Contoso 5a7b9100',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Copilot Eligible M365 Seats': '64', 'Copilot Seats Whitespace': '48',
    'Copilot MAU (Licensed)': '16', 'Copilot PAU': '19', 'Copilot Utilization': '84%',
    'Adoption Status': 'Failure to Thrive/Adopt', 'Free Copilot Chat MAU (Unlicensed)': '22', 'All Copilot MAU': '38',
    'Free to Paid Whitespace': '22', 'All Agents MAU': 'Not Available', 'Copilot Opportunity': 'Other',
    'Potential Earnings': '$1,350', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-4c8-0f3', 'Tenant ID': 'D12E2A30-0025-44c8-a0f3-9c294c80f311',
    'Tenant Country/Region': 'Canada', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Education',
    'Customer TPID': '-22177364', 'Customer Name': 'Contoso 4c80f311',
    'EA Renewal': '2027-06-30', 'Largest Seat CSP Renewal': '2026-09-30',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '210', 'Copilot Seats Whitespace': '165',
    'Copilot MAU (Licensed)': '45', 'Copilot PAU': '63', 'Copilot Utilization': '71%',
    'Adoption Status': 'Starting', 'Free Copilot Chat MAU (Unlicensed)': '58', 'All Copilot MAU': '103',
    'Free to Paid Whitespace': '58', 'All Agents MAU': '17', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': '$3,900', 'CSP Promos': 'Available (6)'
  },
  {
    'Tenant Name': 'Fabrikam 123-8ef-6c2', 'Tenant ID': 'D12E2A30-0029-48ef-b6c2-6aa28ef6c233',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Information Technology',
    'Customer TPID': '-11863294', 'Customer Name': 'Contoso 8ef6c233',
    'EA Renewal': '2028-01-31', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Copilot Eligible M365 Seats': '780', 'Copilot Seats Whitespace': '505',
    'Copilot MAU (Licensed)': '220', 'Copilot PAU': '275', 'Copilot Utilization': '80%',
    'Adoption Status': 'Healthy', 'Free Copilot Chat MAU (Unlicensed)': '144', 'All Copilot MAU': '364',
    'Free to Paid Whitespace': '144', 'All Agents MAU': '92', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$11,250', 'CSP Promos': 'Available (4)'
  },
  {
    'Tenant Name': 'Fabrikam 123-1af-3c2', 'Tenant ID': 'D12E2A30-0031-41AF-B3C2-1AF3C2D4E501',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Managed', 'Industry Vertical': 'Banking',
    'Customer TPID': '-10452111', 'Customer Name': 'Contoso 1af3c201',
    'EA Renewal': '2027-12-31', 'Largest Seat CSP Renewal': '2027-08-31',
    'Dominant SKU Group': 'E5', 'Copilot Eligible M365 Seats': '4,500',
    'Copilot Seats Whitespace': '3,100', 'Copilot MAU (Licensed)': '1,220',
    'Copilot PAU': '1,386', 'Copilot Utilization': '89%',
    'Adoption Status': 'Healthy', 'Free Copilot Chat MAU (Unlicensed)': '640',
    'All Copilot MAU': '1,860', 'Free to Paid Whitespace': '640',
    'All Agents MAU': '315', 'Copilot Opportunity': 'Grow',
    'Potential Earnings': '$42,500', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-2b4-5d6', 'Tenant ID': 'D12E2A30-0032-42B4-A5D6-2B45D6E7F602',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Retail',
    'Customer TPID': '-98345102', 'Customer Name': 'Contoso 2b45d602',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-15',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '75',
    'Copilot Seats Whitespace': '59', 'Copilot MAU (Licensed)': '16',
    'Copilot PAU': '21', 'Copilot Utilization': '76%',
    'Adoption Status': 'Starting', 'Free Copilot Chat MAU (Unlicensed)': '18',
    'All Copilot MAU': '34', 'Free to Paid Whitespace': '18',
    'All Agents MAU': '6', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': '$1,950', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-3c7-8e1', 'Tenant ID': 'D12E2A30-0033-43C7-B8E1-3C78E1F2A703',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Information Technology',
    'Customer TPID': '-87456213', 'Customer Name': 'Contoso 3c78e103',
    'EA Renewal': '2027-05-31', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Copilot Eligible M365 Seats': '860',
    'Copilot Seats Whitespace': '510', 'Copilot MAU (Licensed)': '280',
    'Copilot PAU': '350', 'Copilot Utilization': '80%',
    'Adoption Status': 'Healthy', 'Free Copilot Chat MAU (Unlicensed)': '132',
    'All Copilot MAU': '412', 'Free to Paid Whitespace': '132',
    'All Agents MAU': '74', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$12,750', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-4d9-1f2', 'Tenant ID': 'D12E2A30-0034-44D9-A1F2-4D91F2A3B804',
    'Tenant Country/Region': 'Korea', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Electronics',
    'Customer TPID': '-76567324', 'Customer Name': 'Contoso 4d91f204',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Copilot Eligible M365 Seats': '42',
    'Copilot Seats Whitespace': '42', 'Copilot MAU (Licensed)': 'Not Available',
    'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '7',
    'All Copilot MAU': '7', 'Free to Paid Whitespace': '7',
    'All Agents MAU': '1', 'Copilot Opportunity': 'Other',
    'Potential Earnings': 'Not Available', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-5ea-7b4', 'Tenant ID': 'D12E2A30-0035-45EA-B7B4-5EA7B4C5D905',
    'Tenant Country/Region': 'United Kingdom', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Strategic', 'Industry Vertical': 'Media',
    'Customer TPID': '-65678435', 'Customer Name': 'Contoso 5ea7b405',
    'EA Renewal': '2028-03-31', 'Largest Seat CSP Renewal': '2027-11-30',
    'Dominant SKU Group': 'E5', 'Copilot Eligible M365 Seats': '3,200',
    'Copilot Seats Whitespace': '1,450', 'Copilot MAU (Licensed)': '1,480',
    'Copilot PAU': '1,750', 'Copilot Utilization': '85%',
    'Adoption Status': 'Failure to Thrive/Adopt', 'Free Copilot Chat MAU (Unlicensed)': '420',
    'All Copilot MAU': '1,900', 'Free to Paid Whitespace': '420',
    'All Agents MAU': '205', 'Copilot Opportunity': 'Grow',
    'Potential Earnings': '$31,800', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-6fb-2c5', 'Tenant ID': 'D12E2A30-0036-46FB-A2C5-6FB2C5D6EA06',
    'Tenant Country/Region': 'France', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Consumer Goods',
    'Customer TPID': '-54789546', 'Customer Name': 'Contoso 6fb2c506',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': '2027-02-28',
    'Dominant SKU Group': 'E3', 'Copilot Eligible M365 Seats': '1,150',
    'Copilot Seats Whitespace': '820', 'Copilot MAU (Licensed)': '205',
    'Copilot PAU': '330', 'Copilot Utilization': '62%',
    'Adoption Status': 'Starting', 'Free Copilot Chat MAU (Unlicensed)': '188',
    'All Copilot MAU': '393', 'Free to Paid Whitespace': '188',
    'All Agents MAU': '54', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': '$9,600', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-7ac-4d8', 'Tenant ID': 'D12E2A30-0037-47AC-B4D8-7AC4D8E7FB07',
    'Tenant Country/Region': 'Australia', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Mining',
    'Customer TPID': '-43990657', 'Customer Name': 'Contoso 7ac4d807',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-10-31',
    'Dominant SKU Group': 'BP', 'Copilot Eligible M365 Seats': '210',
    'Copilot Seats Whitespace': '210', 'Copilot MAU (Licensed)': 'Not Available',
    'Copilot PAU': '0', 'Copilot Utilization': 'Not Available',
    'Adoption Status': 'Not Available', 'Free Copilot Chat MAU (Unlicensed)': '41',
    'All Copilot MAU': '41', 'Free to Paid Whitespace': '41',
    'All Agents MAU': '8', 'Copilot Opportunity': 'Acquire',
    'Potential Earnings': '$4,200', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-8bd-6e9', 'Tenant ID': 'D12E2A30-0038-48BD-A6E9-8BD6E9F8AC08',
    'Tenant Country/Region': 'Japan', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Managed', 'Industry Vertical': 'Automotive',
    'Customer TPID': '-32101768', 'Customer Name': 'Contoso 8bd6e908',
    'EA Renewal': '2028-06-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Copilot Eligible M365 Seats': '5,000',
    'Copilot Seats Whitespace': '2,640', 'Copilot MAU (Licensed)': '2,040',
    'Copilot PAU': '2,360', 'Copilot Utilization': '86%',
    'Adoption Status': 'Healthy', 'Free Copilot Chat MAU (Unlicensed)': '710',
    'All Copilot MAU': '2,750', 'Free to Paid Whitespace': '710',
    'All Agents MAU': '402', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$49,900', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-9ce-1a4', 'Tenant ID': 'D12E2A30-0039-49CE-B1A4-9CE1A4B9CD09',
    'Tenant Country/Region': 'India', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Business Services',
    'Customer TPID': '-21212879', 'Customer Name': 'Contoso 9ce1a409',
    'EA Renewal': '2027-01-31', 'Largest Seat CSP Renewal': '2026-09-15',
    'Dominant SKU Group': 'BS', 'Copilot Eligible M365 Seats': '155',
    'Copilot Seats Whitespace': '118', 'Copilot MAU (Licensed)': '37',
    'Copilot PAU': '52', 'Copilot Utilization': '71%',
    'Adoption Status': 'Starting', 'Free Copilot Chat MAU (Unlicensed)': '29',
    'All Copilot MAU': '66', 'Free to Paid Whitespace': '29',
    'All Agents MAU': '12', 'Copilot Opportunity': 'Other',
    'Potential Earnings': '$2,850', 'CSP Promos': 'Available (11)'
  },
  {
    'Tenant Name': 'Fabrikam 123-afd-5b7', 'Tenant ID': 'D12E2A30-0040-4AFD-A5B7-AFD5B7C0DE10',
    'Tenant Country/Region': 'Brazil', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Agriculture',
    'Customer TPID': '-10323980', 'Customer Name': 'Contoso afd5b710',
    'EA Renewal': '2027-07-31', 'Largest Seat CSP Renewal': '2027-04-30',
    'Dominant SKU Group': 'E3', 'Copilot Eligible M365 Seats': '640',
    'Copilot Seats Whitespace': '455', 'Copilot MAU (Licensed)': '118',
    'Copilot PAU': '185', 'Copilot Utilization': '64%',
    'Adoption Status': 'Failure to Thrive/Adopt', 'Free Copilot Chat MAU (Unlicensed)': '96',
    'All Copilot MAU': '214', 'Free to Paid Whitespace': '96',
    'All Agents MAU': '33', 'Copilot Opportunity': 'Monetize',
    'Potential Earnings': '$7,800', 'CSP Promos': 'Available (11)'
  }
];

const e3Data = [
  {
    'Tenant Name': 'Fabrikam 123-e1d-298', 'Tenant ID': 'd12e2a30-66b2-4419-b795-7967cefc7f09',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Other Professional Services',
    'TPID': '-6374392', 'Customer Name': 'Contoso 7f094392',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BB, BS', 'Total M365 Seat Count': '10,338', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'DLM P1 + P2', 'MAU': '0', 'PAU': '0',
    'Utilization': 'Not Available', 'Usage Whitespace': '0', 'Usage Threshold': 'Below', 'Tipping Point': '0%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e1d-298', 'Tenant ID': 'd12e2a30-66b2-4419-b795-7967cefc7f09',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Other Professional Services',
    'TPID': '-6374392', 'Customer Name': 'Contoso 7f094392',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BB, BS', 'Total M365 Seat Count': '10,338', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'DNR P1 + P2', 'MAU': '82', 'PAU': '87',
    'Utilization': '94%', 'Usage Whitespace': '5', 'Usage Threshold': 'Above', 'Tipping Point': '80%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e1d-298', 'Tenant ID': 'd12e2a30-66b2-4419-b795-7967cefc7f09',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Other Professional Services',
    'TPID': '-6374392', 'Customer Name': 'Contoso 7f094392',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BB, BS', 'Total M365 Seat Count': '10,338', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'Entra ID P1+P2', 'MAU': '79', 'PAU': '151',
    'Utilization': '52%', 'Usage Whitespace': '72', 'Usage Threshold': 'Above', 'Tipping Point': '10%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Not Available'
  },
  {
    'Tenant Name': 'Fabrikam 123-e1d-298', 'Tenant ID': 'd12e2a30-66b2-4419-b795-7967cefc7f09',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Other Professional Services',
    'TPID': '-6374392', 'Customer Name': 'Contoso 7f094392',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BB, BS', 'Total M365 Seat Count': '10,338', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'M365 Apps', 'MAU': '86', 'PAU': '98',
    'Utilization': '87%', 'Usage Whitespace': '12', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-fcb-571', 'Tenant ID': 'd12e2a30-77c4-4f5b-86a1-3f4d9b7b5710',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Process Manufacturing',
    'TPID': '-5517820', 'Customer Name': 'Contoso b5717820',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '438', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'Exchange Online', 'MAU': '411', 'PAU': '438',
    'Utilization': '94%', 'Usage Whitespace': '27', 'Usage Threshold': 'Above', 'Tipping Point': '80%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-fcb-571', 'Tenant ID': 'd12e2a30-77c4-4f5b-86a1-3f4d9b7b5710',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Process Manufacturing',
    'TPID': '-5517820', 'Customer Name': 'Contoso b5717820',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '438', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell', 'Workload Name': 'Teams', 'MAU': '7', 'PAU': '422',
    'Utilization': '2%', 'Usage Whitespace': '415', 'Usage Threshold': 'Below', 'Tipping Point': '60%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-4ab-882', 'Tenant ID': 'd12e2a30-88d5-4ab1-a909-4cf120ab8821',
    'Tenant Country/Region': 'Hong Kong', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Wholesale Distribution',
    'TPID': '-4488291', 'Customer Name': 'Contoso ab882921',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-15',
    'Dominant SKU Group': 'BP', 'Total M365 Seat Count': '182', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'Intune', 'MAU': '96', 'PAU': '158',
    'Utilization': '61%', 'Usage Whitespace': '62', 'Usage Threshold': 'Above', 'Tipping Point': '40%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-7de-410', 'Tenant ID': 'd12e2a30-91e6-47de-b104-8b937de410aa',
    'Tenant Country/Region': 'Romania', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Construction',
    'TPID': '-3331048', 'Customer Name': 'Contoso 7de410aa',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Total M365 Seat Count': '94', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell', 'Workload Name': 'SharePoint', 'MAU': '1', 'PAU': '78',
    'Utilization': '1%', 'Usage Whitespace': '77', 'Usage Threshold': 'Below', 'Tipping Point': '35%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-55a-c90', 'Tenant ID': 'd12e2a30-a2f7-455a-9c90-507455ac9001',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Healthcare',
    'TPID': '-2209450', 'Customer Name': 'Contoso 55ac9001',
    'EA Renewal': '2028-02-29', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '1,480', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME3', 'Workload Name': 'M365 Apps', 'MAU': '1,210', 'PAU': '1,360',
    'Utilization': '89%', 'Usage Whitespace': '150', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-1ce-204', 'Tenant ID': 'd12e2a30-b308-41ce-b204-7a3f11ce2048',
    'Tenant Country/Region': 'Belgium', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Professional Services',
    'TPID': '-1187204', 'Customer Name': 'Contoso 1ce2048a',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2027-01-01',
    'Dominant SKU Group': 'BP', 'Total M365 Seat Count': '126', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell', 'Workload Name': 'Exchange Online', 'MAU': '38', 'PAU': '112',
    'Utilization': '34%', 'Usage Whitespace': '74', 'Usage Threshold': 'Below', 'Tipping Point': '50%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-b11-c22', 'Tenant ID': 'd12e2a30-c111-4b11-ac22-1b11c22d3301',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Strategic', 'Industry Vertical': 'Banking',
    'TPID': '-9103301', 'Customer Name': 'Contoso b11c2201',
    'EA Renewal': '2027-12-31', 'Largest Seat CSP Renewal': '2027-08-31',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '2,840',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'Teams', 'MAU': '2,330',
    'PAU': '2,710', 'Utilization': '86%',
    'Usage Whitespace': '380', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-c33-d44', 'Tenant ID': 'd12e2a30-c222-4c33-ad44-2c33d44e4402',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Manufacturing',
    'TPID': '-8204402', 'Customer Name': 'Contoso c33d4402',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '720',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'SharePoint', 'MAU': '205',
    'PAU': '610', 'Utilization': '34%',
    'Usage Whitespace': '405', 'Usage Threshold': 'Below',
    'Tipping Point': '35%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-d55-e66', 'Tenant ID': 'd12e2a30-c333-4d55-ae66-3d55e66f5503',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Technology',
    'TPID': '-7305503', 'Customer Name': 'Contoso d55e6603',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-01',
    'Dominant SKU Group': 'BP', 'Total M365 Seat Count': '188',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'Exchange Online', 'MAU': '167',
    'PAU': '181', 'Utilization': '92%',
    'Usage Whitespace': '14', 'Usage Threshold': 'Above',
    'Tipping Point': '80%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e77-f88', 'Tenant ID': 'd12e2a30-c444-4e77-af88-4e77f88a6604',
    'Tenant Country/Region': 'Korea', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Electronics',
    'TPID': '-6406604', 'Customer Name': 'Contoso e77f8804',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'BS', 'Total M365 Seat Count': '96',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Intune', 'MAU': '29',
    'PAU': '82', 'Utilization': '35%',
    'Usage Whitespace': '53', 'Usage Threshold': 'Below',
    'Tipping Point': '40%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-f99-a10', 'Tenant ID': 'd12e2a30-c555-4f99-aa10-5f99a10b7705',
    'Tenant Country/Region': 'United Kingdom', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Managed', 'Industry Vertical': 'Insurance',
    'TPID': '-5507705', 'Customer Name': 'Contoso f99a1005',
    'EA Renewal': '2028-03-31', 'Largest Seat CSP Renewal': '2027-11-30',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '3,640',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'Entra ID P1+P2', 'MAU': '2,548',
    'PAU': '3,186', 'Utilization': '80%',
    'Usage Whitespace': '638', 'Usage Threshold': 'Above',
    'Tipping Point': '10%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Not Available'
  },
  {
    'Tenant Name': 'Fabrikam 123-a21-b32', 'Tenant ID': 'd12e2a30-c666-4a21-ab32-6a21b32c8806',
    'Tenant Country/Region': 'France', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Hospitality',
    'TPID': '-4608806', 'Customer Name': 'Contoso a21b3206',
    'EA Renewal': '2027-10-31', 'Largest Seat CSP Renewal': '2027-03-15',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '1,120',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'M365 Apps', 'MAU': '918',
    'PAU': '1,048', 'Utilization': '88%',
    'Usage Whitespace': '130', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-b43-c54', 'Tenant ID': 'd12e2a30-c777-4b43-ac54-7b43c54d9907',
    'Tenant Country/Region': 'Australia', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Mining',
    'TPID': '-3719907', 'Customer Name': 'Contoso b43c5407',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-11-30',
    'Dominant SKU Group': 'BP', 'Total M365 Seat Count': '244',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'DLM P1 + P2', 'MAU': '51',
    'PAU': '172', 'Utilization': '30%',
    'Usage Whitespace': '121', 'Usage Threshold': 'Below',
    'Tipping Point': '25%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-c65-d76', 'Tenant ID': 'd12e2a30-c888-4c65-ad76-8c65d76e1108',
    'Tenant Country/Region': 'Japan', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Strategic', 'Industry Vertical': 'Automotive',
    'TPID': '-2821108', 'Customer Name': 'Contoso c65d7608',
    'EA Renewal': '2028-06-30', 'Largest Seat CSP Renewal': '2028-01-31',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '4,280',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'Teams', 'MAU': '3,596',
    'PAU': '4,052', 'Utilization': '89%',
    'Usage Whitespace': '456', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-d87-e98', 'Tenant ID': 'd12e2a30-c999-4d87-ae98-9d87e98f2209',
    'Tenant Country/Region': 'India', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Business Services',
    'TPID': '-1932209', 'Customer Name': 'Contoso d87e9809',
    'EA Renewal': '2027-04-30', 'Largest Seat CSP Renewal': '2026-10-31',
    'Dominant SKU Group': 'BS', 'Total M365 Seat Count': '132',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME3',
    'Workload Name': 'M365 Apps', 'MAU': '74',
    'PAU': '116', 'Utilization': '64%',
    'Usage Whitespace': '42', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e09-f10', 'Tenant ID': 'd12e2a30-d000-4e09-af10-ae09f10a3310',
    'Tenant Country/Region': 'Brazil', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Agriculture',
    'TPID': '-1043310', 'Customer Name': 'Contoso e09f1010',
    'EA Renewal': '2027-07-31', 'Largest Seat CSP Renewal': '2027-05-31',
    'Dominant SKU Group': 'E3', 'Total M365 Seat Count': '860',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Exchange Online', 'MAU': '163',
    'PAU': '612', 'Utilization': '27%',
    'Usage Whitespace': '449', 'Usage Threshold': 'Below',
    'Tipping Point': '50%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  }
];

const e5Data = [
  {
    'Tenant Name': 'Fabrikam 123-fcb-571', 'Tenant ID': 'd12e2a30-44b1-4fcb-9571-3bf14fcb5710',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Process Manufacturing',
    'TPID': '-5517820', 'Customer Name': 'Contoso fcb57100',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '612', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME5', 'Workload Name': 'Exchange Online', 'MAU': '544', 'PAU': '602',
    'Utilization': '90%', 'Usage Whitespace': '58', 'Usage Threshold': 'Above', 'Tipping Point': '80%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-fcb-571', 'Tenant ID': 'd12e2a30-44b1-4fcb-9571-3bf14fcb5710',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Process Manufacturing',
    'TPID': '-5517820', 'Customer Name': 'Contoso fcb57100',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '612', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME5', 'Workload Name': 'Teams', 'MAU': '498', 'PAU': '566',
    'Utilization': '88%', 'Usage Whitespace': '68', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-2e4-990', 'Tenant ID': 'd12e2a30-55c2-42e4-a990-6e7f42e49900',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Healthcare',
    'TPID': '-2281990', 'Customer Name': 'Contoso 2e499000',
    'EA Renewal': '2026-12-31', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '1,124', 'Tenant Health': 'At Risk',
    'Tenant Recommendation': 'Recover usage and upsell to ME5', 'Workload Name': 'SharePoint', 'MAU': '404', 'PAU': '922',
    'Utilization': '44%', 'Usage Whitespace': '518', 'Usage Threshold': 'Below', 'Tipping Point': '55%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-6d7-441', 'Tenant ID': 'd12e2a30-66d3-46d7-a441-1c2946d74410',
    'Tenant Country/Region': 'Canada', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Education',
    'TPID': '-1944410', 'Customer Name': 'Contoso 6d744410',
    'EA Renewal': '2027-06-30', 'Largest Seat CSP Renewal': '2026-09-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '284', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before upsell', 'Workload Name': 'Intune', 'MAU': '76', 'PAU': '208',
    'Utilization': '37%', 'Usage Whitespace': '132', 'Usage Threshold': 'Below', 'Tipping Point': '40%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-8aa-650', 'Tenant ID': 'd12e2a30-77e4-48aa-b650-6ca248aa6500',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Technology',
    'TPID': '-1366500', 'Customer Name': 'Contoso 8aa65000',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2027-03-01',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '356', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Drive upsell to ME5', 'Workload Name': 'M365 Apps', 'MAU': '318', 'PAU': '352',
    'Utilization': '90%', 'Usage Whitespace': '34', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-f12-a23', 'Tenant ID': 'd12e2a30-d111-4f12-aa23-bf12a23b4401',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Strategic', 'Industry Vertical': 'Healthcare',
    'TPID': '-9154401', 'Customer Name': 'Contoso f12a2301',
    'EA Renewal': '2027-12-31', 'Largest Seat CSP Renewal': '2027-09-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '2,480',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'Exchange Online', 'MAU': '2,136',
    'PAU': '2,372', 'Utilization': '90%',
    'Usage Whitespace': '236', 'Usage Threshold': 'Above',
    'Tipping Point': '80%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-a34-b45', 'Tenant ID': 'd12e2a30-d222-4a34-ab45-ca34b45c5502',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Manufacturing',
    'TPID': '-8265502', 'Customer Name': 'Contoso a34b4502',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '940',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Teams', 'MAU': '282',
    'PAU': '742', 'Utilization': '38%',
    'Usage Whitespace': '460', 'Usage Threshold': 'Below',
    'Tipping Point': '60%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-b56-c67', 'Tenant ID': 'd12e2a30-d333-4b56-ac67-db56c67d6603',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Technology',
    'TPID': '-7376603', 'Customer Name': 'Contoso b56c6703',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2027-02-28',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '310',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'M365 Apps', 'MAU': '273',
    'PAU': '302', 'Utilization': '90%',
    'Usage Whitespace': '29', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-c78-d89', 'Tenant ID': 'd12e2a30-d444-4c78-ad89-ec78d89e7704',
    'Tenant Country/Region': 'Korea', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Electronics',
    'TPID': '-6487704', 'Customer Name': 'Contoso c78d8904',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '188',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Intune', 'MAU': '48',
    'PAU': '151', 'Utilization': '32%',
    'Usage Whitespace': '103', 'Usage Threshold': 'Below',
    'Tipping Point': '40%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-d90-e01', 'Tenant ID': 'd12e2a30-d555-4d90-ae01-fd90e01f8805',
    'Tenant Country/Region': 'United Kingdom', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Managed', 'Industry Vertical': 'Legal',
    'TPID': '-5598805', 'Customer Name': 'Contoso d90e0105',
    'EA Renewal': '2028-05-31', 'Largest Seat CSP Renewal': '2027-12-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '3,120',
    'Tenant Health': 'At Risk', 'Tenant Recommendation': 'Recover usage and upsell to ME5',
    'Workload Name': 'SharePoint', 'MAU': '1,450',
    'PAU': '2,520', 'Utilization': '58%',
    'Usage Whitespace': '1,070', 'Usage Threshold': 'Below',
    'Tipping Point': '55%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e12-f23', 'Tenant ID': 'd12e2a30-d666-4e12-af23-ae12f23a9906',
    'Tenant Country/Region': 'France', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Retail',
    'TPID': '-4609906', 'Customer Name': 'Contoso e12f2306',
    'EA Renewal': '2027-11-30', 'Largest Seat CSP Renewal': '2027-04-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '1,280',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'Entra ID P1+P2', 'MAU': '1,088',
    'PAU': '1,246', 'Utilization': '87%',
    'Usage Whitespace': '158', 'Usage Threshold': 'Above',
    'Tipping Point': '10%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Not Available'
  },
  {
    'Tenant Name': 'Fabrikam 123-f34-a45', 'Tenant ID': 'd12e2a30-d777-4f34-aa45-bf34a45b1107',
    'Tenant Country/Region': 'Australia', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Energy',
    'TPID': '-3721107', 'Customer Name': 'Contoso f34a4507',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '276',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'Teams', 'MAU': '201',
    'PAU': '248', 'Utilization': '81%',
    'Usage Whitespace': '47', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-a56-b67', 'Tenant ID': 'd12e2a30-d888-4a56-ab67-ca56b67c2208',
    'Tenant Country/Region': 'Japan', 'Tenant Segment': 'Enterprise',
    'Customer Sub-Segment': 'Enterprise Strategic', 'Industry Vertical': 'Automotive',
    'TPID': '-2832208', 'Customer Name': 'Contoso a56b6708',
    'EA Renewal': '2028-07-31', 'Largest Seat CSP Renewal': '2028-02-29',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '4,420',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'DLM P1 + P2', 'MAU': '3,730',
    'PAU': '4,118', 'Utilization': '91%',
    'Usage Whitespace': '388', 'Usage Threshold': 'Above',
    'Tipping Point': '25%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-b78-c89', 'Tenant ID': 'd12e2a30-d999-4b78-ac89-db78c89d3309',
    'Tenant Country/Region': 'India', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Professional Services',
    'TPID': '-1943309', 'Customer Name': 'Contoso b78c8909',
    'EA Renewal': '2027-05-31', 'Largest Seat CSP Renewal': '2026-10-15',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '224',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before upsell',
    'Workload Name': 'Exchange Online', 'MAU': '41',
    'PAU': '186', 'Utilization': '22%',
    'Usage Whitespace': '145', 'Usage Threshold': 'Below',
    'Tipping Point': '50%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-c90-d12', 'Tenant ID': 'd12e2a30-e000-4c90-ad12-ec90d12e4410',
    'Tenant Country/Region': 'Brazil', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Agriculture',
    'TPID': '-1054410', 'Customer Name': 'Contoso c90d1210',
    'EA Renewal': '2027-08-31', 'Largest Seat CSP Renewal': '2027-06-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '780',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Drive upsell to ME5',
    'Workload Name': 'M365 Apps', 'MAU': '608',
    'PAU': '716', 'Utilization': '85%',
    'Usage Whitespace': '108', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  }
];

const e5expData = [
  {
    'Tenant Name': 'Fabrikam 123-3af-218', 'Tenant ID': 'd12e2a30-88f5-43af-a218-88b743af2180',
    'Tenant Country/Region': 'Belgium', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Business Services',
    'TPID': '-742180', 'Customer Name': 'Contoso 3af21800',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-08-01',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '92', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Expand existing E5 footprint', 'Workload Name': 'Teams', 'MAU': '74', 'PAU': '91',
    'Utilization': '81%', 'Usage Whitespace': '17', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-3af-218', 'Tenant ID': 'd12e2a30-88f5-43af-a218-88b743af2180',
    'Tenant Country/Region': 'Belgium', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Business Services',
    'TPID': '-742180', 'Customer Name': 'Contoso 3af21800',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-08-01',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '92', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Recover workload health before expansion', 'Workload Name': 'SharePoint', 'MAU': '21', 'PAU': '79',
    'Utilization': '27%', 'Usage Whitespace': '58', 'Usage Threshold': 'Below', 'Tipping Point': '45%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-0bd-305', 'Tenant ID': 'd12e2a30-99a6-40bd-b305-90a140bd3050',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Legal',
    'TPID': '-1933050', 'Customer Name': 'Contoso 0bd30500',
    'EA Renewal': '2027-11-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '126', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Expand existing E5 footprint', 'Workload Name': 'Exchange Online', 'MAU': '102', 'PAU': '121',
    'Utilization': '84%', 'Usage Whitespace': '19', 'Usage Threshold': 'Above', 'Tipping Point': '80%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-6ce-144', 'Tenant ID': 'd12e2a30-aab7-46ce-b144-77f246ce1440',
    'Tenant Country/Region': 'Canada', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Architecture',
    'TPID': '-861440', 'Customer Name': 'Contoso 6ce14400',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2027-02-14',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '58', 'Tenant Health': 'Unhealthy',
    'Tenant Recommendation': 'Drive usage before expansion', 'Workload Name': 'Intune', 'MAU': '9', 'PAU': '44',
    'Utilization': '20%', 'Usage Whitespace': '35', 'Usage Threshold': 'Below', 'Tipping Point': '40%',
    'Workload Health': 'Unhealthy', 'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-4fe-088', 'Tenant ID': 'd12e2a30-bbc8-44fe-a088-18d944fe0880',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Professional Services',
    'TPID': '-520880', 'Customer Name': 'Contoso 4fe08800',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-01',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '71', 'Tenant Health': 'Healthy',
    'Tenant Recommendation': 'Expand existing E5 footprint', 'Workload Name': 'M365 Apps', 'MAU': '58', 'PAU': '69',
    'Utilization': '84%', 'Usage Whitespace': '11', 'Usage Threshold': 'Above', 'Tipping Point': '60%',
    'Workload Health': 'Healthy', 'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-d13-e24', 'Tenant ID': 'd12e2a30-e111-4d13-ae24-fd13e24f5501',
    'Tenant Country/Region': 'United States', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Healthcare',
    'TPID': '-9165501', 'Customer Name': 'Contoso d13e2401',
    'EA Renewal': '2027-10-31', 'Largest Seat CSP Renewal': '2027-06-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '184',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'Exchange Online', 'MAU': '152',
    'PAU': '178', 'Utilization': '85%',
    'Usage Whitespace': '26', 'Usage Threshold': 'Above',
    'Tipping Point': '80%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e35-f46', 'Tenant ID': 'd12e2a30-e222-4e35-af46-ae35f46a6602',
    'Tenant Country/Region': 'Germany', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Manufacturing',
    'TPID': '-8276602', 'Customer Name': 'Contoso e35f4602',
    'EA Renewal': '2027-09-30', 'Largest Seat CSP Renewal': 'Not Available',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '240',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before expansion',
    'Workload Name': 'Teams', 'MAU': '88',
    'PAU': '206', 'Utilization': '43%',
    'Usage Whitespace': '118', 'Usage Threshold': 'Below',
    'Tipping Point': '60%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-f57-a68', 'Tenant ID': 'd12e2a30-e333-4f57-aa68-bf57a68b7703',
    'Tenant Country/Region': 'Netherlands', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Technology',
    'TPID': '-7387703', 'Customer Name': 'Contoso f57a6803',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2027-03-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '156',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'M365 Apps', 'MAU': '126',
    'PAU': '149', 'Utilization': '85%',
    'Usage Whitespace': '23', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-a79-b80', 'Tenant ID': 'd12e2a30-e444-4a79-ab80-ca79b80c8804',
    'Tenant Country/Region': 'Korea', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Electronics',
    'TPID': '-6498804', 'Customer Name': 'Contoso a79b8004',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-12-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '98',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'Intune', 'MAU': '69',
    'PAU': '92', 'Utilization': '75%',
    'Usage Whitespace': '23', 'Usage Threshold': 'Above',
    'Tipping Point': '40%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-b91-c02', 'Tenant ID': 'd12e2a30-e555-4b91-ac02-db91c02d9905',
    'Tenant Country/Region': 'United Kingdom', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Strategic', 'Industry Vertical': 'Legal',
    'TPID': '-5509905', 'Customer Name': 'Contoso b91c0205',
    'EA Renewal': '2028-02-29', 'Largest Seat CSP Renewal': '2027-12-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '320',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Recover workload health before expansion',
    'Workload Name': 'SharePoint', 'MAU': '96',
    'PAU': '282', 'Utilization': '34%',
    'Usage Whitespace': '186', 'Usage Threshold': 'Below',
    'Tipping Point': '45%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-c13-d24', 'Tenant ID': 'd12e2a30-e666-4c13-ad24-ec13d24e1106',
    'Tenant Country/Region': 'France', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Retail',
    'TPID': '-4611106', 'Customer Name': 'Contoso c13d2406',
    'EA Renewal': '2027-11-30', 'Largest Seat CSP Renewal': '2027-05-15',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '212',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'Entra ID P1+P2', 'MAU': '168',
    'PAU': '201', 'Utilization': '84%',
    'Usage Whitespace': '33', 'Usage Threshold': 'Above',
    'Tipping Point': '10%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Not Available'
  },
  {
    'Tenant Name': 'Fabrikam 123-d35-e46', 'Tenant ID': 'd12e2a30-e777-4d35-ae46-fd35e46f2207',
    'Tenant Country/Region': 'Australia', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C - SMB Commercial', 'Industry Vertical': 'Education',
    'TPID': '-3722207', 'Customer Name': 'Contoso d35e4607',
    'EA Renewal': 'Not Available', 'Largest Seat CSP Renewal': '2026-11-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '140',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before expansion',
    'Workload Name': 'DLM P1 + P2', 'MAU': '34',
    'PAU': '118', 'Utilization': '29%',
    'Usage Whitespace': '84', 'Usage Threshold': 'Below',
    'Tipping Point': '25%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-e57-f68', 'Tenant ID': 'd12e2a30-e888-4e57-af68-ae57f68a3308',
    'Tenant Country/Region': 'Japan', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Automotive',
    'TPID': '-2833308', 'Customer Name': 'Contoso e57f6808',
    'EA Renewal': '2028-06-30', 'Largest Seat CSP Renewal': '2028-01-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '410',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'Exchange Online', 'MAU': '356',
    'PAU': '398', 'Utilization': '89%',
    'Usage Whitespace': '42', 'Usage Threshold': 'Above',
    'Tipping Point': '80%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Licensing and Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-f79-a80', 'Tenant ID': 'd12e2a30-e999-4f79-aa80-bf79a80b4409',
    'Tenant Country/Region': 'India', 'Tenant Segment': 'SME&C SMB',
    'Customer Sub-Segment': 'SME&C Commercial - SMB Default', 'Industry Vertical': 'Business Services',
    'TPID': '-1944409', 'Customer Name': 'Contoso f79a8009',
    'EA Renewal': '2027-04-30', 'Largest Seat CSP Renewal': '2026-10-31',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '126',
    'Tenant Health': 'Healthy', 'Tenant Recommendation': 'Expand existing E5 footprint',
    'Workload Name': 'Teams', 'MAU': '95',
    'PAU': '119', 'Utilization': '80%',
    'Usage Whitespace': '24', 'Usage Threshold': 'Above',
    'Tipping Point': '60%', 'Workload Health': 'Healthy',
    'Usage Opportunity': 'Usage'
  },
  {
    'Tenant Name': 'Fabrikam 123-a02-b13', 'Tenant ID': 'd12e2a30-f000-4a02-ab13-ca02b13c5510',
    'Tenant Country/Region': 'Brazil', 'Tenant Segment': 'SMME Corporate',
    'Customer Sub-Segment': 'SMME Corporate Managed', 'Industry Vertical': 'Agriculture',
    'TPID': '-1055510', 'Customer Name': 'Contoso a02b1310',
    'EA Renewal': '2027-08-31', 'Largest Seat CSP Renewal': '2027-06-30',
    'Dominant SKU Group': 'E5', 'Total M365 Seat Count': '276',
    'Tenant Health': 'Unhealthy', 'Tenant Recommendation': 'Drive usage before expansion',
    'Workload Name': 'M365 Apps', 'MAU': '71',
    'PAU': '234', 'Utilization': '30%',
    'Usage Whitespace': '163', 'Usage Threshold': 'Below',
    'Tipping Point': '60%', 'Workload Health': 'Unhealthy',
    'Usage Opportunity': 'Usage'
  }
];
const oppData = {
  copilot: {
    kpis: [
      { value: '60,560', label: 'Tenants' },
      { value: '3%', label: 'Seat Penetration' },
      { value: '78%', label: 'Paid Utilization' },
      { value: '2,076', label: 'Acquires', color: '#107c10' },
      { value: '10,470', label: 'Monetize', color: '#5c2d91' },
      { value: '3,700', label: 'New', color: '#0078d4' },
      { value: '0', label: '# of E7 Upsell-ready' },
      { value: '$95,000', label: 'Potential Earnings' },
    ],
    columns: copilotColumns,
    rows: copilotData
  },
  e3: {
    kpis: [
      { value: '15,953', label: 'Tenants' },
      { value: '88%', label: 'Paid Utilization' },
      { value: '3,296,190', label: 'Usage Whitespace' },
      { value: '10,549', label: 'Upsell ready', color: '#107c10' },
      { value: '5,404', label: 'Unhealthy', color: '#d83b01' },
      { value: '0', label: 'At Risk' },
      { value: '0', label: 'Incentive Display' },
      { value: '$50', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e3Data
  },
  e5: {
    kpis: [
      { value: '11,206', label: 'Tenants' },
      { value: '90%', label: 'Paid Utilization' },
      { value: '1,546,009', label: 'Usage Whitespace' },
      { value: '7,353', label: 'Upsell ready', color: '#107c10' },
      { value: '3,440', label: 'Unhealthy', color: '#d83b01' },
      { value: '413', label: 'At Risk', color: '#9c27b0' },
      { value: '0', label: '# of E7 Upsell-ready' },
      { value: '$50', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e5Data
  },
  e5exp: {
    kpis: [
      { value: '394', label: 'Tenants' },
      { value: '81%', label: 'Paid Utilization' },
      { value: '206,673', label: 'Usage Whitespace' },
      { value: '174', label: 'Upsell ready', color: '#107c10' },
      { value: '212', label: 'Unhealthy', color: '#d83b01' },
      { value: '0', label: 'At Risk' },
      { value: '0', label: '# of E7 Upsell-ready' },
      { value: '$50', label: 'Potential Earnings' },
    ],
    columns: e3Columns,
    rows: e5expData
  }
};

const mciColumns = {
  copilotPower: [
    'Engagement Name', 'Claim ID', 'TPID', 'Tenant Name', 'Tenant ID', 'Partner Area',
    'Date Initiated', 'Date POE Completed', 'Months Since Completion', 'Performance Status',
    'Milestone', 'Goal Type', 'Attainment Goal', 'Total Attainment', 'Repeat Customer',
    'M365 Copilot Revenue Goal', 'M365 Copilot Revenue Attainment', 'Copilot Chat MAU Goal', 'Copilot Chat MAU Attained'
  ],
  secureProductivity: [
    'Engagement Name', 'Claim ID', 'TPID', 'Tenant Name', 'Tenant ID', 'Partner Area',
    'Date Initiated', 'Date POE Completed', 'Months Since Completion', 'Performance Status',
    'Milestone', 'Goal Type', 'Attainment Goal', 'Total Attainment', 'Revenue Goal', 'Revenue Attained'
  ],
  cloudEndpoints: [
    'Engagement Name', 'Claim ID', 'TPID', 'Tenant Name', 'Tenant ID', 'Partner Area',
    'Date Initiated', 'Date POE Completed', 'Months Since Completion', 'Performance Status',
    'Milestone', 'Goal Type', 'Attainment Goal', 'Total Attainment', 'Revenue Goal', 'Revenue Attained'
  ]
};

const mciFilterConfig = {
  copilotPower: [
    { field: 'Partner Area' },
    { field: 'Performance Status', options: ['Passing', 'Failing', 'Pre-Milestone'] },
    { field: 'Milestone' },
    { field: 'Goal Type' },
    { field: 'Repeat Customer', options: ['Yes', 'No'] }
  ],
  secureProductivity: [
    { field: 'Partner Area' },
    { field: 'Performance Status', options: ['Passing', 'Failing', 'Pre-Milestone'] },
    { field: 'Milestone' },
    { field: 'Goal Type' }
  ],
  cloudEndpoints: [
    { field: 'Partner Area' },
    { field: 'Performance Status', options: ['Passing', 'Failing', 'Pre-Milestone'] },
    { field: 'Milestone' },
    { field: 'Goal Type' }
  ]
};

const mciData = {
  copilotPower: {
    kpis: [
      { title: 'Copilot + Power Envisioning', value: '57%', badge: 'Eligible', tone: 'eligible' },
      { title: 'Copilot + Power Deployment', value: '68%', badge: 'Eligible', tone: 'eligible' }
    ],
    columns: mciColumns.copilotPower,
    rows: [
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (XS)', 'Claim ID': '3321266', 'TPID': '11393510', 'Tenant Name': 'Fabrikam 123-aee-6e9', 'Tenant ID': 'D12E2A30-0002-4355-a055-ed07f88bc72a', 'Partner Area': 'ASEAN', 'Date Initiated': '7/7/2025', 'Date POE Completed': '9/29/2025', 'Months Since Completion': '8', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '193%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '113%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '80%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (XS)', 'Claim ID': '3323065', 'TPID': '13649469', 'Tenant Name': 'Fabrikam 123-85f-501', 'Tenant ID': 'D12E2A30-0005-414e-a92e-fbbc96ca4137', 'Partner Area': 'ASEAN', 'Date Initiated': '7/10/2025', 'Date POE Completed': '9/23/2025', 'Months Since Completion': '8', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '93%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '55%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '38%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (XS)', 'Claim ID': '3323372', 'TPID': '8452682', 'Tenant Name': 'Fabrikam 123-453-4e5', 'Tenant ID': 'D12E2A30-0008-4c40-b3c3-1947eecba3f6', 'Partner Area': 'ASEAN', 'Date Initiated': '7/11/2025', 'Date POE Completed': '9/17/2025', 'Months Since Completion': '8', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '23%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '0%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '23%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (XS)', 'Claim ID': '3325469', 'TPID': '7898511', 'Tenant Name': 'Fabrikam 123-62f-ae4', 'Tenant ID': 'D12E2A30-000a-4f46-bb1f-b39d6b6a4dab', 'Partner Area': 'ASEAN', 'Date Initiated': '7/17/2025', 'Date POE Completed': '9/29/2025', 'Months Since Completion': '8', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '600%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '375%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '127%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (S)', 'Claim ID': '3325479', 'TPID': '3152920', 'Tenant Name': 'Fabrikam 123-c9a-7b2', 'Tenant ID': 'D12E2A30-1122-4c9a-87b2-4d5e6f7a8b9c', 'Partner Area': 'ASEAN', 'Date Initiated': '7/17/2025', 'Date POE Completed': '8/27/2025', 'Months Since Completion': '9', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '325%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '20%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '294%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (M)', 'Claim ID': '3326681', 'TPID': '6355563', 'Tenant Name': 'Fabrikam 123-d4e-3f5', 'Tenant ID': 'D12E2A30-2233-4d4e-93f5-5e6f7a8b9c0d', 'Partner Area': 'MEA', 'Date Initiated': '7/21/2025', 'Date POE Completed': '10/20/2025', 'Months Since Completion': '7', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '69%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$125,000', 'M365 Copilot Revenue Attainment': '0%', 'Copilot Chat MAU Goal': '1250', 'Copilot Chat MAU Attained': '42%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (S)', 'Claim ID': '3328181', 'TPID': '2006870', 'Tenant Name': 'Fabrikam 123-e5f-4a6', 'Tenant ID': 'D12E2A30-3344-4e5f-a4a6-6f7a8b9c0d1e', 'Partner Area': 'ASEAN', 'Date Initiated': '7/24/2025', 'Date POE Completed': '9/2/2025', 'Months Since Completion': '8', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '177%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '0%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '177%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (S)', 'Claim ID': '3328421', 'TPID': '12218363', 'Tenant Name': 'Fabrikam 123-f6a-5b7', 'Tenant ID': 'D12E2A30-4455-4f6a-b5b7-7a8b9c0d1e2f', 'Partner Area': 'South Europe', 'Date Initiated': '7/24/2025', 'Date POE Completed': '10/16/2025', 'Months Since Completion': '7', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '441%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '70%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '312%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (XS)', 'Claim ID': '3329730', 'TPID': '940707', 'Tenant Name': 'Fabrikam 123-a7b-6c8', 'Tenant ID': 'D12E2A30-5566-4a7b-c6c8-8b9c0d1e2f3a', 'Partner Area': 'ASEAN', 'Date Initiated': '7/29/2025', 'Date POE Completed': '9/25/2025', 'Months Since Completion': '8', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '249%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '71%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '86%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (M)', 'Claim ID': '3332611', 'TPID': '1293510', 'Tenant Name': 'Fabrikam 123-b8c-7d9', 'Tenant ID': 'D12E2A30-6677-4b8c-d7d9-9c0d1e2f3a4b', 'Partner Area': 'ASEAN', 'Date Initiated': '8/5/2025', 'Date POE Completed': '10/30/2025', 'Months Since Completion': '7', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '33%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$125,000', 'M365 Copilot Revenue Attainment': '2%', 'Copilot Chat MAU Goal': '1250', 'Copilot Chat MAU Attained': '31%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (S)', 'Claim ID': '3334184', 'TPID': '4185201', 'Tenant Name': 'Fabrikam 123-1af-3c2', 'Tenant ID': 'D12E2A30-0031-41AF-B3C2-1AF3C2D4E501', 'Partner Area': 'North Europe', 'Date Initiated': '8/8/2025', 'Date POE Completed': '10/29/2025', 'Months Since Completion': '7', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '212%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '96%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '116%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (M)', 'Claim ID': '3336420', 'TPID': '5296312', 'Tenant Name': 'Fabrikam 123-3c7-8e1', 'Tenant ID': 'D12E2A30-0033-43C7-B8E1-3C78E1F2A703', 'Partner Area': 'North Europe', 'Date Initiated': '8/12/2025', 'Date POE Completed': '11/6/2025', 'Months Since Completion': '6', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '188%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$125,000', 'M365 Copilot Revenue Attainment': '44%', 'Copilot Chat MAU Goal': '1250', 'Copilot Chat MAU Attained': '144%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (XS)', 'Claim ID': '3338188', 'TPID': '6407423', 'Tenant Name': 'Fabrikam 123-7ac-4d8', 'Tenant ID': 'D12E2A30-0037-47AC-B4D8-7AC4D8E7FB07', 'Partner Area': 'ANZ', 'Date Initiated': '8/15/2025', 'Date POE Completed': '10/8/2025', 'Months Since Completion': '7', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '72%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '18%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '54%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (S)', 'Claim ID': '3340426', 'TPID': '7518534', 'Tenant Name': 'Fabrikam 123-8bd-6e9', 'Tenant ID': 'D12E2A30-0038-48BD-A6E9-8BD6E9F8AC08', 'Partner Area': 'Japan', 'Date Initiated': '8/20/2025', 'Date POE Completed': '11/19/2025', 'Months Since Completion': '6', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '356%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '122%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '234%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (M)', 'Claim ID': '3342871', 'TPID': '8629645', 'Tenant Name': 'Fabrikam 123-afd-5b7', 'Tenant ID': 'D12E2A30-0040-4AFD-A5B7-AFD5B7C0DE10', 'Partner Area': 'LATAM', 'Date Initiated': '8/26/2025', 'Date POE Completed': '11/28/2025', 'Months Since Completion': '6', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '141%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$125,000', 'M365 Copilot Revenue Attainment': '31%', 'Copilot Chat MAU Goal': '1250', 'Copilot Chat MAU Attained': '110%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (S)', 'Claim ID': '3345100', 'TPID': '9740756', 'Tenant Name': 'Fabrikam 123-c9d-8e0', 'Tenant ID': 'D12E2A30-7788-4c9d-e8e0-0d1e2f3a4b5c', 'Partner Area': 'North Europe', 'Date Initiated': '9/1/2025', 'Date POE Completed': '11/15/2025', 'Months Since Completion': '6', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '275%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '150%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '125%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (XS)', 'Claim ID': '3347200', 'TPID': '1085867', 'Tenant Name': 'Fabrikam 123-d0e-9f1', 'Tenant ID': 'D12E2A30-8899-4d0e-f9f1-1e2f3a4b5c6d', 'Partner Area': 'UKI', 'Date Initiated': '9/5/2025', 'Date POE Completed': '11/22/2025', 'Months Since Completion': '6', 'Performance Status': 'Failing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '45%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '12%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '33%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (M)', 'Claim ID': '3349300', 'TPID': '2196978', 'Tenant Name': 'Fabrikam 123-e1f-0a2', 'Tenant ID': 'D12E2A30-9900-4e1f-a0a2-2f3a4b5c6d7e', 'Partner Area': 'India', 'Date Initiated': '9/10/2025', 'Date POE Completed': '12/1/2025', 'Months Since Completion': '5', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$125,000', 'M365 Copilot Revenue Attainment': '0%', 'Copilot Chat MAU Goal': '1250', 'Copilot Chat MAU Attained': '0%' },
      { 'Engagement Name': 'Copilot + Power Deployment Accelerator (S)', 'Claim ID': '3351400', 'TPID': '3308089', 'Tenant Name': 'Fabrikam 123-f2a-1b3', 'Tenant ID': 'D12E2A30-0011-4f2a-b1b3-3a4b5c6d7e8f', 'Partner Area': 'LATAM', 'Date Initiated': '9/15/2025', 'Date POE Completed': '12/10/2025', 'Months Since Completion': '5', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '156%', 'Repeat Customer': 'Yes', 'M365 Copilot Revenue Goal': '$50,000', 'M365 Copilot Revenue Attainment': '88%', 'Copilot Chat MAU Goal': '500', 'Copilot Chat MAU Attained': '68%' },
      { 'Engagement Name': 'Copilot + Power Envisioning & PoC (XS)', 'Claim ID': '3353500', 'TPID': '4419190', 'Tenant Name': 'Fabrikam 123-a3b-2c4', 'Tenant ID': 'D12E2A30-0022-4a3b-c2c4-4b5c6d7e8f9a', 'Partner Area': 'Korea', 'Date Initiated': '9/20/2025', 'Date POE Completed': '12/18/2025', 'Months Since Completion': '5', 'Performance Status': 'Passing', 'Milestone': 'M2', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '410%', 'Repeat Customer': 'No', 'M365 Copilot Revenue Goal': '$25,000', 'M365 Copilot Revenue Attainment': '210%', 'Copilot Chat MAU Goal': '250', 'Copilot Chat MAU Attained': '200%' }
    ]
  },
  secureProductivity: {
    kpis: [
      { title: 'Secure Productivity Envisioning', value: '50%', badge: 'Eligible', tone: 'eligible' }
    ],
    columns: mciColumns.secureProductivity,
    rows: [
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3351167', 'TPID': '5006376', 'Tenant Name': 'Fabrikam 123-aee-6e9', 'Tenant ID': 'D12E2A30-0002-4355-a055-ed07f88bc72a', 'Partner Area': 'India', 'Date Initiated': '9/19/2025', 'Date POE Completed': '11/21/2025', 'Months Since Completion': '6', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '90%', 'Revenue Goal': '$12,500', 'Revenue Attained': '90%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3351747', 'TPID': '2844077', 'Tenant Name': 'Fabrikam 123-85f-501', 'Tenant ID': 'D12E2A30-0005-414e-a92e-fbbc96ca4137', 'Partner Area': 'North Europe', 'Date Initiated': '9/22/2025', 'Date POE Completed': '12/16/2025', 'Months Since Completion': '5', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3352434', 'TPID': '93142289', 'Tenant Name': 'Fabrikam 123-453-4e5', 'Tenant ID': 'D12E2A30-0008-4c40-b3c3-1947eecba3f6', 'Partner Area': 'ASEAN', 'Date Initiated': '9/23/2025', 'Date POE Completed': '12/3/2025', 'Months Since Completion': '5', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3355300', 'TPID': '5329443', 'Tenant Name': 'Fabrikam 123-62f-ae4', 'Tenant ID': 'D12E2A30-000a-4f46-bb1f-b39d6b6a4dab', 'Partner Area': 'ASEAN', 'Date Initiated': '9/29/2025', 'Date POE Completed': '11/13/2025', 'Months Since Completion': '6', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '1754%', 'Revenue Goal': '$25,000', 'Revenue Attained': '1754%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3362603', 'TPID': '5526668', 'Tenant Name': 'Fabrikam 123-c9a-7b2', 'Tenant ID': 'D12E2A30-1122-4c9a-87b2-4d5e6f7a8b9c', 'Partner Area': 'North Europe', 'Date Initiated': '10/15/2025', 'Date POE Completed': '12/9/2025', 'Months Since Completion': '5', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3375933', 'TPID': '881935', 'Tenant Name': 'Fabrikam 123-d4e-3f5', 'Tenant ID': 'D12E2A30-2233-4d4e-93f5-5e6f7a8b9c0d', 'Partner Area': 'North Europe', 'Date Initiated': '11/12/2025', 'Date POE Completed': '2/25/2026', 'Months Since Completion': '3', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3376735', 'TPID': '99526230', 'Tenant Name': 'Fabrikam 123-e5f-4a6', 'Tenant ID': 'D12E2A30-3344-4e5f-a4a6-6f7a8b9c0d1e', 'Partner Area': 'ASEAN', 'Date Initiated': '11/13/2025', 'Date POE Completed': '3/3/2026', 'Months Since Completion': '2', 'Performance Status': 'Passing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': 'Not available', 'Revenue Goal': 'Not available', 'Revenue Attained': 'Not available' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3382526', 'TPID': '39758914', 'Tenant Name': 'Fabrikam 123-f6a-5b7', 'Tenant ID': 'D12E2A30-4455-4f6a-b5b7-7a8b9c0d1e2f', 'Partner Area': 'ASEAN', 'Date Initiated': '11/25/2025', 'Date POE Completed': '12/1/2025', 'Months Since Completion': '5', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '3979%', 'Revenue Goal': '$25,000', 'Revenue Attained': '3979%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3383974', 'TPID': '3152920', 'Tenant Name': 'Fabrikam 123-a7b-6c8', 'Tenant ID': 'D12E2A30-5566-4a7b-c6c8-8b9c0d1e2f3a', 'Partner Area': 'ASEAN', 'Date Initiated': '11/27/2025', 'Date POE Completed': '12/9/2025', 'Months Since Completion': '5', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '294%', 'Revenue Goal': '$12,500', 'Revenue Attained': '294%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3406204', 'TPID': '8452682', 'Tenant Name': 'Fabrikam 123-b8c-7d9', 'Tenant ID': 'D12E2A30-6677-4b8c-d7d9-9c0d1e2f3a4b', 'Partner Area': 'ASEAN', 'Date Initiated': '1/19/2026', 'Date POE Completed': '4/8/2026', 'Months Since Completion': '1', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3408815', 'TPID': '4625104', 'Tenant Name': 'Fabrikam 123-2b4-5d6', 'Tenant ID': 'D12E2A30-0032-42B4-A5D6-2B45D6E7F602', 'Partner Area': 'North Europe', 'Date Initiated': '1/26/2026', 'Date POE Completed': '4/22/2026', 'Months Since Completion': '1', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '126%', 'Revenue Goal': '$12,500', 'Revenue Attained': '126%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3410162', 'TPID': '5736215', 'Tenant Name': 'Fabrikam 123-5ea-7b4', 'Tenant ID': 'D12E2A30-0035-45EA-B7B4-5EA7B4C5D905', 'Partner Area': 'UKI', 'Date Initiated': '1/29/2026', 'Date POE Completed': '4/30/2026', 'Months Since Completion': '1', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '214%', 'Revenue Goal': '$25,000', 'Revenue Attained': '214%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3412054', 'TPID': '6847326', 'Tenant Name': 'Fabrikam 123-6fb-2c5', 'Tenant ID': 'D12E2A30-0036-46FB-A2C5-6FB2C5D6EA06', 'Partner Area': 'South Europe', 'Date Initiated': '2/3/2026', 'Date POE Completed': '5/2/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3413848', 'TPID': '7958437', 'Tenant Name': 'Fabrikam 123-9ce-1a4', 'Tenant ID': 'D12E2A30-0039-49CE-B1A4-9CE1A4B9CD09', 'Partner Area': 'India', 'Date Initiated': '2/6/2026', 'Date POE Completed': '5/8/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': 'Not available', 'Revenue Goal': 'Not available', 'Revenue Attained': 'Not available' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3415579', 'TPID': '8069548', 'Tenant Name': 'Fabrikam 123-1af-3c2', 'Tenant ID': 'D12E2A30-0031-41AF-B3C2-1AF3C2D4E501', 'Partner Area': 'United States', 'Date Initiated': '2/10/2026', 'Date POE Completed': '5/14/2026', 'Months Since Completion': '0', 'Performance Status': 'Failing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$25,000', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3420100', 'TPID': '4063201', 'Tenant Name': 'Fabrikam 123-b4c-3d5', 'Tenant ID': 'D12E2A30-AA01-4b4c-a3d5-b4c3d5e6f701', 'Partner Area': 'Korea', 'Date Initiated': '2/15/2026', 'Date POE Completed': '5/14/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3422200', 'TPID': '5174312', 'Tenant Name': 'Fabrikam 123-c5d-4e6', 'Tenant ID': 'D12E2A30-BB02-4c5d-b4e6-c5d4e6f7a802', 'Partner Area': 'United States', 'Date Initiated': '2/20/2026', 'Date POE Completed': '5/20/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '185%', 'Revenue Goal': '$25,000', 'Revenue Attained': '185%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3424300', 'TPID': '6285423', 'Tenant Name': 'Fabrikam 123-d6e-5f7', 'Tenant ID': 'D12E2A30-CC03-4d6e-c5f7-d6e5f7a8b903', 'Partner Area': 'ANZ', 'Date Initiated': '2/25/2026', 'Date POE Completed': '5/25/2026', 'Months Since Completion': '0', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '42%', 'Revenue Goal': '$12,500', 'Revenue Attained': '42%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (M)', 'Claim ID': '3426400', 'TPID': '7396534', 'Tenant Name': 'Fabrikam 123-e7f-6a8', 'Tenant ID': 'D12E2A30-DD04-4e7f-d6a8-e7f6a8b9c004', 'Partner Area': 'MEA', 'Date Initiated': '3/1/2026', 'Date POE Completed': '5/28/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '320%', 'Revenue Goal': '$25,000', 'Revenue Attained': '320%' },
      { 'Engagement Name': 'Secure AI Productivity Envisioning & PoC (S)', 'Claim ID': '3428500', 'TPID': '8407645', 'Tenant Name': 'Fabrikam 123-f8a-7b9', 'Tenant ID': 'D12E2A30-EE05-4f8a-e7b9-f8a7b9c0d105', 'Partner Area': 'Japan', 'Date Initiated': '3/5/2026', 'Date POE Completed': '5/30/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$12,500', 'Revenue Attained': '0%' }
    ]
  },
  cloudEndpoints: {
    kpis: [
      { title: 'Cloud Endpoints Envisioning', value: '0%', badge: 'Pending', tone: 'pending' }
    ],
    columns: mciColumns.cloudEndpoints,
    rows: [
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3338702', 'TPID': '2511110', 'Tenant Name': 'Fabrikam 123-aee-6e9', 'Tenant ID': 'D12E2A30-0002-4355-a055-ed07f88bc72a', 'Partner Area': 'North Europe', 'Date Initiated': '8/21/2025', 'Date POE Completed': '11/18/2025', 'Months Since Completion': '6', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3362967', 'TPID': '29095691', 'Tenant Name': 'Fabrikam 123-85f-501', 'Tenant ID': 'D12E2A30-0005-414e-a92e-fbbc96ca4137', 'Partner Area': 'ASEAN', 'Date Initiated': '10/16/2025', 'Date POE Completed': '1/21/2026', 'Months Since Completion': '4', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3377759', 'TPID': '11424807', 'Tenant Name': 'Fabrikam 123-453-4e5', 'Tenant ID': 'D12E2A30-0008-4c40-b3c3-1947eecba3f6', 'Partner Area': 'North Europe', 'Date Initiated': '11/14/2025', 'Date POE Completed': '2/13/2026', 'Months Since Completion': '3', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3383900', 'TPID': '47916855', 'Tenant Name': 'Fabrikam 123-62f-ae4', 'Tenant ID': 'D12E2A30-000a-4f46-bb1f-b39d6b6a4dab', 'Partner Area': 'North Europe', 'Date Initiated': '11/27/2025', 'Date POE Completed': '3/12/2026', 'Months Since Completion': '2', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3402078', 'TPID': '21500966', 'Tenant Name': 'Fabrikam 123-c9a-7b2', 'Tenant ID': 'D12E2A30-1122-4c9a-87b2-4d5e6f7a8b9c', 'Partner Area': 'North Europe', 'Date Initiated': '1/9/2026', 'Date POE Completed': '4/17/2026', 'Months Since Completion': '1', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3404192', 'TPID': '3164201', 'Tenant Name': 'Fabrikam 123-3c7-8e1', 'Tenant ID': 'D12E2A30-0033-43C7-B8E1-3C78E1F2A703', 'Partner Area': 'North Europe', 'Date Initiated': '1/12/2026', 'Date POE Completed': '4/16/2026', 'Months Since Completion': '1', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3405984', 'TPID': '4275312', 'Tenant Name': 'Fabrikam 123-5ea-7b4', 'Tenant ID': 'D12E2A30-0035-45EA-B7B4-5EA7B4C5D905', 'Partner Area': 'UKI', 'Date Initiated': '1/16/2026', 'Date POE Completed': '4/24/2026', 'Months Since Completion': '1', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3407426', 'TPID': '5386423', 'Tenant Name': 'Fabrikam 123-6fb-2c5', 'Tenant ID': 'D12E2A30-0036-46FB-A2C5-6FB2C5D6EA06', 'Partner Area': 'South Europe', 'Date Initiated': '1/20/2026', 'Date POE Completed': '4/28/2026', 'Months Since Completion': '1', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '68%', 'Revenue Goal': '$18,750', 'Revenue Attained': '68%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3409335', 'TPID': '6497534', 'Tenant Name': 'Fabrikam 123-7ac-4d8', 'Tenant ID': 'D12E2A30-0037-47AC-B4D8-7AC4D8E7FB07', 'Partner Area': 'ANZ', 'Date Initiated': '1/23/2026', 'Date POE Completed': '5/1/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '118%', 'Revenue Goal': '$18,750', 'Revenue Attained': '118%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3411278', 'TPID': '7508645', 'Tenant Name': 'Fabrikam 123-8bd-6e9', 'Tenant ID': 'D12E2A30-0038-48BD-A6E9-8BD6E9F8AC08', 'Partner Area': 'Japan', 'Date Initiated': '1/30/2026', 'Date POE Completed': '5/6/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': 'Not available', 'Revenue Goal': '$37,500', 'Revenue Attained': 'Not available' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3413166', 'TPID': '8619756', 'Tenant Name': 'Fabrikam 123-9ce-1a4', 'Tenant ID': 'D12E2A30-0039-49CE-B1A4-9CE1A4B9CD09', 'Partner Area': 'India', 'Date Initiated': '2/4/2026', 'Date POE Completed': '5/11/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3415021', 'TPID': '9730867', 'Tenant Name': 'Fabrikam 123-afd-5b7', 'Tenant ID': 'D12E2A30-0040-4AFD-A5B7-AFD5B7C0DE10', 'Partner Area': 'LATAM', 'Date Initiated': '2/9/2026', 'Date POE Completed': '5/18/2026', 'Months Since Completion': '0', 'Performance Status': 'Failing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3416884', 'TPID': '1841978', 'Tenant Name': 'Fabrikam 123-1af-3c2', 'Tenant ID': 'D12E2A30-0031-41AF-B3C2-1AF3C2D4E501', 'Partner Area': 'United States', 'Date Initiated': '2/13/2026', 'Date POE Completed': '5/20/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': 'Not available', 'Revenue Goal': '$18,750', 'Revenue Attained': 'Not available' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3418709', 'TPID': '2952089', 'Tenant Name': 'Fabrikam 123-2b4-5d6', 'Tenant ID': 'D12E2A30-0032-42B4-A5D6-2B45D6E7F602', 'Partner Area': 'North Europe', 'Date Initiated': '2/18/2026', 'Date POE Completed': '5/23/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '142%', 'Revenue Goal': '$18,750', 'Revenue Attained': '142%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3420583', 'TPID': '3063190', 'Tenant Name': 'Fabrikam 123-4d9-1f2', 'Tenant ID': 'D12E2A30-0034-44D9-A1F2-4D91F2A3B804', 'Partner Area': 'Korea', 'Date Initiated': '2/24/2026', 'Date POE Completed': '5/28/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3422100', 'TPID': '1074301', 'Tenant Name': 'Fabrikam 123-a9b-8c0', 'Tenant ID': 'D12E2A30-FF06-4a9b-f8c0-a9b8c0d1e206', 'Partner Area': 'LATAM', 'Date Initiated': '2/28/2026', 'Date POE Completed': '5/26/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '167%', 'Revenue Goal': '$18,750', 'Revenue Attained': '167%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3424200', 'TPID': '2185412', 'Tenant Name': 'Fabrikam 123-b0c-9d1', 'Tenant ID': 'D12E2A30-0107-4b0c-a9d1-b0c9d1e2f307', 'Partner Area': 'United States', 'Date Initiated': '3/3/2026', 'Date POE Completed': '5/29/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$37,500', 'Revenue Attained': '0%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3426300', 'TPID': '3296523', 'Tenant Name': 'Fabrikam 123-c1d-0e2', 'Tenant ID': 'D12E2A30-0208-4c1d-b0e2-c1d0e2f3a408', 'Partner Area': 'North Europe', 'Date Initiated': '3/7/2026', 'Date POE Completed': '5/31/2026', 'Months Since Completion': '0', 'Performance Status': 'Failing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '55%', 'Revenue Goal': '$18,750', 'Revenue Attained': '55%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (M)', 'Claim ID': '3428400', 'TPID': '4407634', 'Tenant Name': 'Fabrikam 123-d2e-1f3', 'Tenant ID': 'D12E2A30-0309-4d2e-c1f3-d2e1f3a4b509', 'Partner Area': 'India', 'Date Initiated': '3/10/2026', 'Date POE Completed': '6/2/2026', 'Months Since Completion': '0', 'Performance Status': 'Passing', 'Milestone': 'M1', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '203%', 'Revenue Goal': '$37,500', 'Revenue Attained': '203%' },
      { 'Engagement Name': 'Cloud Endpoints Envisioning & PoC (S)', 'Claim ID': '3430500', 'TPID': '5518745', 'Tenant Name': 'Fabrikam 123-e3f-2a4', 'Tenant ID': 'D12E2A30-0410-4e3f-d2a4-e3f2a4b5c610', 'Partner Area': 'ASEAN', 'Date Initiated': '3/14/2026', 'Date POE Completed': '6/5/2026', 'Months Since Completion': '0', 'Performance Status': 'Pre-Milestone', 'Milestone': 'Pre-Milestone', 'Goal Type': 'Individual', 'Attainment Goal': '100%', 'Total Attainment': '0%', 'Revenue Goal': '$18,750', 'Revenue Attained': '0%' }
    ]
  }
};

let currentPage = 'growth';
let currentTab = 'copilot';
let currentMciTab = 'copilotPower';
let selectedRow = null;
let selectedRowData = null;
let currentUtilityPanel = null;

const activeColumnsByTab = {
  copilot: [...copilotColumns],
  e3: [...e3Columns],
  e5: [...e3Columns],
  e5exp: [...e3Columns]
};

const activeColumnsByMciTab = {
  copilotPower: [...mciColumns.copilotPower],
  secureProductivity: [...mciColumns.secureProductivity],
  cloudEndpoints: [...mciColumns.cloudEndpoints]
};

const activeFiltersByTab = {
  copilot: {},
  e3: {},
  e5: {},
  e5exp: {}
};

const activeFiltersByMciTab = {
  copilotPower: {},
  secureProductivity: {},
  cloudEndpoints: {}
};

const searchTermsByTab = {
  copilot: '',
  e3: '',
  e5: '',
  e5exp: ''
};

const searchTermsByMciTab = {
  copilotPower: '',
  secureProductivity: '',
  cloudEndpoints: ''
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// OPPORTUNITY TABS
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

function initSidebarNavigation() {
  const navGrowth = document.getElementById('navGrowth');
  const navMci = document.getElementById('navMci');
  const navFasttrack = document.getElementById('navFasttrack');

  if (navGrowth) {
    navGrowth.addEventListener('click', event => {
      event.preventDefault();
      switchPage('growth');
    });
  }

  if (navMci) {
    navMci.addEventListener('click', event => {
      event.preventDefault();
      switchPage('mci');
    });
  }

  if (navFasttrack) {
    navFasttrack.addEventListener('click', event => {
      event.preventDefault();
      switchPage('fasttrack');
    });
  }
}

function initOppTabs() {
  document.querySelectorAll('.opp-tabs-bar .opp-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.opp-tabs-bar .opp-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentTab = tab.dataset.opp;
      closeUtilityFlyout();
      if (oppData[currentTab]) {
        renderOppView(currentTab);
      }
    });
  });
}

function initMciTabs() {
  document.querySelectorAll('.mci-tabs-bar .mci-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.mci-tabs-bar .mci-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentMciTab = tab.dataset.mci;
      closeUtilityFlyout();
      if (mciData[currentMciTab]) {
        renderMciView(currentMciTab);
      }
    });
  });
}

function canonicalizeColumn(column) {
  return ({ TPID: 'Customer TPID', 'Usage Whitespace': 'Whitespace' })[column] || column;
}

function getAllColumnsForTab(tab) {
  return tab === 'copilot' ? allCopilotColumns : allE3Columns;
}

function getVisibleColumns(tab) {
  return (activeColumnsByTab[tab] || oppData[tab].columns).map(canonicalizeColumn);
}

function getRowValue(row, column) {
  const aliases = columnAliases[column] || [column];
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias) && row[alias] !== undefined && row[alias] !== null && row[alias] !== '') {
      return row[alias];
    }
  }
  return '';
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getDisplayRows(opp) {
  const search = (searchTermsByTab[opp] || '').trim().toLowerCase();
  const filters = activeFiltersByTab[opp] || {};

  return oppData[opp].rows
    .map((row, originalIndex) => ({ row, originalIndex }))
    .filter(({ row }) => {
      const matchesFilters = Object.entries(filters).every(([field, values]) => {
        if (!values || !values.length) return true;
        return values.includes(String(getRowValue(row, field) || '').trim());
      });

      if (!matchesFilters) return false;
      if (!search) return true;

      return getAllColumnsForTab(opp).some(column => String(getRowValue(row, column) || '').toLowerCase().includes(search));
    });
}

function getMciVisibleColumns(tab) {
  return activeColumnsByMciTab[tab] || mciData[tab].columns;
}

function getMciDisplayRows(tab) {
  const search = (searchTermsByMciTab[tab] || '').trim().toLowerCase();
  const filters = activeFiltersByMciTab[tab] || {};

  return mciData[tab].rows
    .map((row, originalIndex) => ({ row, originalIndex }))
    .filter(({ row }) => {
      const matchesFilters = Object.entries(filters).every(([field, values]) => {
        if (!values || !values.length) return true;
        return values.includes(String(getRowValue(row, field) || '').trim());
      });

      if (!matchesFilters) return false;
      if (!search) return true;

      return mciData[tab].columns.some(column => String(getRowValue(row, column) || '').toLowerCase().includes(search));
    });
}

function getActiveFilterCount(tab) {
  return Object.values(activeFiltersByTab[tab] || {}).reduce((count, values) => count + (values && values.length ? 1 : 0), 0);
}

function getActiveMciFilterCount(tab) {
  return Object.values(activeFiltersByMciTab[tab] || {}).reduce((count, values) => count + (values && values.length ? 1 : 0), 0);
}

function updateToolbarState() {
  const isGrowth = currentPage === 'growth';
  const filterBtn = document.getElementById(isGrowth ? 'filterBtn' : 'mciFilterBtn');
  const filterLabel = document.getElementById(isGrowth ? 'filterBtnLabel' : 'mciFilterBtnLabel');
  const searchInput = document.getElementById(isGrowth ? 'tenantSearch' : 'mciSearch');
  const activeCount = isGrowth ? getActiveFilterCount(currentTab) : getActiveMciFilterCount(currentMciTab);

  if (filterLabel) {
    filterLabel.textContent = activeCount ? `Filter (${activeCount})` : 'Filter';
  }

  if (filterBtn) {
    filterBtn.classList.toggle('is-active', activeCount > 0);
  }

  if (searchInput && searchInput !== document.activeElement) {
    searchInput.value = isGrowth ? (searchTermsByTab[currentTab] || '') : (searchTermsByMciTab[currentMciTab] || '');
  }
}

function switchPage(page) {
  currentPage = page;
  closeUtilityFlyout();

  const aspxView = document.getElementById('aspxView');
  const mciView = document.getElementById('mciView');
  const fasttrackView = document.getElementById('fasttrackView');
  const navGrowth = document.getElementById('navGrowth');
  const navMci = document.getElementById('navMci');
  const navFasttrack = document.getElementById('navFasttrack');
  const navPageName = document.querySelector('.nav-page-name');

  if (aspxView) aspxView.classList.toggle('hidden', page !== 'growth');
  if (mciView) mciView.classList.toggle('hidden', page !== 'mci');
  if (fasttrackView) fasttrackView.classList.toggle('hidden', page !== 'fasttrack');
  if (navGrowth) navGrowth.classList.toggle('active', page === 'growth');
  if (navMci) navMci.classList.toggle('active', page === 'mci');
  if (navFasttrack) navFasttrack.classList.toggle('active', page === 'fasttrack');
  if (navPageName) {
    if (page === 'growth') navPageName.textContent = 'Growth Opportunities';
    else if (page === 'mci') navPageName.textContent = 'MCI Performance';
    else navPageName.textContent = 'FastTrack Referrals';
  }

  renderCurrentPage();
}

function renderCurrentPage() {
  if (currentPage === 'mci') {
    renderMciView(currentMciTab);
  } else if (currentPage === 'fasttrack') {
    renderFasttrackView();
  } else {
    renderOppView(currentTab);
  }
}

function renderOppView(opp) {
  const data = oppData[opp];
  if (!data) return;

  updateToolbarState();

  const kpiRow = document.getElementById('kpiRow');
  kpiRow.innerHTML = data.kpis.map(k =>
    `<div class="kpi-tile">
      <svg class="kpi-info-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1Zm.75 10.5h-1.5V7h1.5v4.5ZM8 6a.75.75 0 1 1 0-1.5A.75.75 0 0 1 8 6Z"/></svg>
      <span class="kpi-value"${k.color ? ` style="color:${k.color}"` : ''}>${k.value}</span>
      <span class="kpi-label">${k.label}</span>
    </div>`
  ).join('');

  const visibleColumns = getVisibleColumns(opp);
  const visibleRows = getDisplayRows(opp);

  const head = document.getElementById('tableHead');
  head.innerHTML = visibleColumns.map(c => `<th>${c}</th>`).join('');

  const body = document.getElementById('tableBody');
  if (!visibleRows.length) {
    body.innerHTML = `<tr><td colspan="${visibleColumns.length}" style="padding: 20px; text-align: center; color: #616161;">No rows match the current search and filters.</td></tr>`;
    renderCharts(opp);
    return;
  }

  body.innerHTML = visibleRows.map(({ row, originalIndex }) => {
    const cells = visibleColumns.map(col => {
      let val = getRowValue(row, col) || '—';
      if (col === 'Tenant Name') {
        val = `<a href="#" class="tenant-link" data-row="${originalIndex}" data-flyout="tenant-details">${val}</a>`;
      } else if (col === 'Copilot Opportunity') {
        const cls = String(val).toLowerCase();
        val = `<a href="#" class="tenant-link" data-row="${originalIndex}" data-flyout="opportunity-details"><span class="status-badge ${cls}">${val}</span></a>`;
      } else if (col === 'Adoption Status') {
        const cls = String(val).includes('Healthy') ? 'healthy' : String(val).includes('Failure') ? 'unhealthy' : '';
        val = cls ? `<span class="status-badge ${cls}">${val}</span>` : val;
      } else if (col === 'Tenant Health' || col === 'Workload Health') {
        const cls = String(val).toLowerCase().replace(/\s/g, '-');
        val = `<span class="status-badge ${cls}">${val}</span>`;
      } else if (col === 'All Copilot MAU' && val !== 'Not Available' && val !== '—') {
        val = `<a href="#" class="tenant-link" data-row="${originalIndex}" data-flyout="copilot-mau">${val}</a>`;
      } else if (col === 'All Agents MAU' && val !== 'Not Available' && val !== '—') {
        val = `<a href="#" class="tenant-link" data-row="${originalIndex}" data-flyout="agents-mau">${val}</a>`;
      } else if (col === 'CSP Promos' && String(val).includes('Available')) {
        val = `<a href="#" class="tenant-link" data-row="${originalIndex}" data-flyout="csp-promos">${val}</a>`;
      }
      return `<td>${val}</td>`;
    }).join('');

    return `<tr data-row="${originalIndex}">${cells}</tr>`;
  }).join('');

  renderCharts(opp);
}

function getMciPerformanceStatusMarkup(value) {
  const label = String(value || '—');
  const statusClass = label.toLowerCase().replace(/\s+/g, '-');
  return `<span class="performance-status ${statusClass}"><span class="performance-status-dot"></span>${escapeHtml(label)}</span>`;
}

function renderMciView(tab) {
  const data = mciData[tab];
  if (!data) return;

  updateToolbarState();

  // Enhanced KPIs - compute from data
  const rows = data.rows;
  const totalEngagements = rows.length;
  const passingCount = rows.filter(r => r['Performance Status'] === 'Passing').length;
  const failingCount = rows.filter(r => r['Performance Status'] === 'Failing').length;
  const preMilestoneCount = rows.filter(r => r['Performance Status'] === 'Pre-Milestone').length;
  const passRate = totalEngagements ? Math.round((passingCount / totalEngagements) * 100) : 0;
  const repeatCustomers = rows.filter(r => r['Repeat Customer'] === 'Yes').length;

  const enhancedKpis = [
    ...data.kpis,
    { title: 'Total Engagements', value: String(totalEngagements), badge: 'All', tone: 'neutral' },
    { title: 'Passing', value: String(passingCount), badge: passRate + '% pass rate', tone: 'passing' },
    { title: 'Failing', value: String(failingCount), badge: 'Needs attention', tone: 'failing' },
    { title: 'Pre-Milestone', value: String(preMilestoneCount), badge: 'Early stage', tone: 'pending' },
  ];

  const kpiRow = document.getElementById('mciKpiRow');
  if (kpiRow) {
    kpiRow.innerHTML = enhancedKpis.map(kpi => `
      <div class="mci-kpi-tile">
        <div class="mci-kpi-actions">
          <button class="icon-btn" type="button" aria-label="Filter KPI">
            <svg viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"></path></svg>
          </button>
          <button class="icon-btn" type="button" aria-label="KPI info">
            <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1Zm.75 10.5h-1.5V7h1.5v4.5ZM8 6a.75.75 0 1 1 0-1.5A.75.75 0 0 1 8 6Z"></path></svg>
          </button>
        </div>
        <div class="mci-kpi-title">${escapeHtml(kpi.title)}</div>
        <div class="mci-kpi-value">${escapeHtml(kpi.value)}</div>
        <div class="mci-kpi-badge ${escapeHtml(kpi.tone)}">${escapeHtml(kpi.badge)}</div>
      </div>
    `).join('');
  }

  // Charts
  renderMciCharts(tab, rows);

  const visibleColumns = getMciVisibleColumns(tab);
  const visibleRows = getMciDisplayRows(tab);
  const head = document.getElementById('mciTableHead');
  const body = document.getElementById('mciTableBody');
  const summary = document.getElementById('mciResultsSummary');

  if (head) {
    head.innerHTML = visibleColumns.map(column => `<th>${escapeHtml(column)}</th>`).join('');
  }

  if (!body) return;

  if (!visibleRows.length) {
    body.innerHTML = `<tr><td colspan="${visibleColumns.length}" style="padding: 20px; text-align: center; color: #616161;">No rows match the current search and filters.</td></tr>`;
  } else {
    body.innerHTML = visibleRows.map(({ row, originalIndex }) => {
      const cells = visibleColumns.map(column => {
        let value = getRowValue(row, column) || '—';
        if (column === 'Performance Status') {
          value = getMciPerformanceStatusMarkup(value);
        } else {
          value = escapeHtml(value);
        }
        return `<td>${value}</td>`;
      }).join('');
      return `<tr data-row="${originalIndex}">${cells}</tr>`;
    }).join('');
  }

  if (summary) {
    const total = visibleRows.length;
    summary.textContent = total ? `Showing 1 - ${Math.min(total, 10)} of ${total} results` : 'Showing 0 - 0 of 0 results';
  }

  // Render MCI charts
  renderMciCharts(tab, visibleRows.map(r => r.row));
}
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// CHARTS (Chart.js)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•


let mciChartInstances = [];

function renderMciCharts(tab, rows) {
  const container = document.getElementById('mciChartsRow');
  if (!container) return;

  mciChartInstances.forEach(c => c.destroy());
  mciChartInstances = [];

  const statusCounts = {};
  const areaCounts = {};
  rows.forEach(row => {
    const status = row['Performance Status'] || 'Unknown';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
    const area = row['Partner Area'] || 'Unknown';
    areaCounts[area] = (areaCounts[area] || 0) + 1;
  });

  const statusColors = { 'Passing': '#107c10', 'Failing': '#d13438', 'Pre-Milestone': '#f7941d' };

  container.innerHTML = `
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">By Performance Status</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="mciChartStatus"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">By Partner Area</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="mciChartArea"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">Attainment Distribution</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="mciChartAttainment"></canvas></div>
    </div>
  `;

  const statusLabels = Object.keys(statusCounts);
  mciChartInstances.push(new Chart(document.getElementById('mciChartStatus'), {
    type: 'doughnut',
    data: {
      labels: statusLabels,
      datasets: [{ data: statusLabels.map(l => statusCounts[l]), backgroundColor: statusLabels.map(l => statusColors[l] || '#999') }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 12 } } }
    }
  }));

  const areaLabels = Object.keys(areaCounts).sort((a, b) => areaCounts[b] - areaCounts[a]);
  mciChartInstances.push(new Chart(document.getElementById('mciChartArea'), {
    type: 'bar',
    data: {
      labels: areaLabels,
      datasets: [{ data: areaLabels.map(l => areaCounts[l]), backgroundColor: '#1b3a5c', borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true, grid: { color: '#e8e8e8', drawBorder: false }, ticks: { font: { size: 10 } } },
        y: { grid: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  }));

  const attainmentBuckets = { '0%': 0, '1-50%': 0, '51-100%': 0, '101-200%': 0, '200%+': 0 };
  rows.forEach(row => {
    const val = parseInt(row['Total Attainment']);
    if (isNaN(val)) return;
    if (val === 0) attainmentBuckets['0%']++;
    else if (val <= 50) attainmentBuckets['1-50%']++;
    else if (val <= 100) attainmentBuckets['51-100%']++;
    else if (val <= 200) attainmentBuckets['101-200%']++;
    else attainmentBuckets['200%+']++;
  });
  const attLabels = Object.keys(attainmentBuckets);
  const attColors = ['#d13438', '#f7941d', '#fff4ce', '#107c10', '#0078d4'];
  mciChartInstances.push(new Chart(document.getElementById('mciChartAttainment'), {
    type: 'bar',
    data: {
      labels: attLabels,
      datasets: [{ data: attLabels.map(l => attainmentBuckets[l]), backgroundColor: attColors, borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: '#e8e8e8', drawBorder: false }, ticks: { font: { size: 10 }, stepSize: 1 } },
        x: { grid: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  }));
}
function chartCardMarkup(title, id) {
  return `
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">${title}</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
        <button class="chart-card-filter"><svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 2h13l-5 6v5l-3 2V8L1.5 2z"/></svg> Filter</button>
      </div>
      <div class="chart-canvas-wrap"><canvas id="${id}"></canvas></div>
    </div>
  `;
}

function chartLegendOptions(fontSize = 9) {
  return {
    position: 'bottom',
    labels: {
      boxWidth: 10,
      boxHeight: 10,
      usePointStyle: true,
      pointStyle: 'circle',
      font: { size: fontSize }
    }
  };
}

function chartGridAxis() {
  return {
    ticks: { font: { size: 9 }, color: '#616161' },
    grid: { color: '#f0f0f0' }
  };
}

function chartCategoryAxis() {
  return {
    ticks: { font: { size: 9 }, color: '#616161' },
    grid: { display: false }
  };
}

function destroyCharts() {
  if (typeof Chart === 'undefined' || !Chart.instances) return;
  Object.keys(Chart.instances).forEach(key => {
    Chart.instances[key].destroy();
  });
}

function getChartDefinitions(opp) {
  const definitions = {
    copilot: [
      {
        title: 'Copilot Tenant Penetration',
        id: 'chartTenantPen',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Has Copilot Licenses', 'No Copilot Licenses'],
            datasets: [{ data: [1817, 58743], backgroundColor: ['#0078d4', '#f7941d'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(10) } }
        }
      },
      {
        title: 'Copilot Opportunities',
        id: 'chartOpportunities',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Acquire', 'Monetize', 'Grow', 'Other'],
            datasets: [{ data: [2076, 10470, 3700, 44314], backgroundColor: ['#1b3a5c', '#0078d4', '#107c10', '#f7941d'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(10) } }
        }
      },
      {
        title: 'Copilot Adoption by State',
        id: 'chartAdoption',
        config: {
          type: 'bar',
          data: {
            labels: ['Other', 'Failure to Thrive/Adopt', 'Healthy', 'Starting'],
            datasets: [{ data: [29000, 14000, 12500, 5060], backgroundColor: '#0078d4', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, ...chartGridAxis() }, y: chartCategoryAxis() }
          }
        }
      },
      {
        title: 'Copilot Usage by License',
        id: 'chartUsage',
        config: {
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
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: chartLegendOptions(9) },
            scales: { y: { beginAtZero: true, ...chartGridAxis() }, x: chartCategoryAxis() }
          }
        }
      },
      {
        title: 'MCI Potential Earnings',
        id: 'chartMci',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Copilot + Power Platform', 'Copilot + Deployment Accelerator (S)', 'Copilot + Deployment Accelerator (XS)', 'Copilot + PoC', 'Other'],
            datasets: [{ data: [35, 19, 14, 13, 19], backgroundColor: ['#1b3a5c', '#0078d4', '#4fc3f7', '#107c10', '#f7941d'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(8) } }
        }
      }
    ],
    e3: [
      {
        title: 'Utilization by Workload',
        id: 'chartE3Utilization',
        config: {
          type: 'bar',
          data: {
            labels: ['Exchange Online', 'Teams', 'SharePoint', 'M365 Apps', 'Intune', 'Entra ID'],
            datasets: [{ data: [92, 88, 75, 91, 45, 62], backgroundColor: '#0078d4', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: {
              x: { beginAtZero: true, max: 100, ...chartGridAxis(), ticks: { font: { size: 9 }, color: '#616161', callback: value => `${value}%` } },
              y: chartCategoryAxis()
            }
          }
        }
      },
      {
        title: 'Opportunities by Usage Whitespace',
        id: 'chartE3Whitespace',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'DLM', 'Other'],
            datasets: [{ data: [890000, 1200000, 680000, 320000, 206190], backgroundColor: chartPalette.slice(0, 5), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      },
      {
        title: 'Workload Health',
        id: 'chartE3Health',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Healthy', 'Unhealthy'],
            datasets: [{ data: [65, 35], backgroundColor: ['#107c10', '#f7941d'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(10) } }
        }
      },
      {
        title: 'Workload Recommendations',
        id: 'chartE3Recommendations',
        config: {
          type: 'bar',
          data: {
            labels: ['Drive upsell', 'Drive usage', 'Improve health'],
            datasets: [{ data: [10549, 5404, 2754], backgroundColor: '#1b3a5c', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, ...chartGridAxis() }, y: chartCategoryAxis() }
          }
        }
      },
      {
        title: 'Potential Earnings by Workloads',
        id: 'chartE3Earnings',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'Entra ID', 'M365 Apps'],
            datasets: [{ data: [18, 14, 8, 5, 5], backgroundColor: chartPalette.slice(0, 5), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      }
    ],
    e5: [
      {
        title: 'Utilization by Workload',
        id: 'chartE5Utilization',
        config: {
          type: 'bar',
          data: {
            labels: ['Exchange Online', 'Teams', 'SharePoint', 'M365 Apps', 'Intune', 'Entra ID'],
            datasets: [{ data: [94, 90, 78, 92, 52, 71], backgroundColor: '#0078d4', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: {
              x: { beginAtZero: true, max: 100, ...chartGridAxis(), ticks: { font: { size: 9 }, color: '#616161', callback: value => `${value}%` } },
              y: chartCategoryAxis()
            }
          }
        }
      },
      {
        title: 'Opportunities by Usage Whitespace',
        id: 'chartE5Whitespace',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'Entra ID', 'Other'],
            datasets: [{ data: [420000, 540000, 310000, 180000, 96009], backgroundColor: chartPalette.slice(0, 5), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      },
      {
        title: 'Workload Health',
        id: 'chartE5Health',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Healthy', 'Unhealthy', 'At Risk'],
            datasets: [{ data: [68, 25, 7], backgroundColor: ['#107c10', '#f7941d', '#9c27b0'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(10) } }
        }
      },
      {
        title: 'Workload Recommendations',
        id: 'chartE5Recommendations',
        config: {
          type: 'bar',
          data: {
            labels: ['Drive upsell', 'Drive usage', 'Secure identity'],
            datasets: [{ data: [7353, 3440, 1653], backgroundColor: '#1b3a5c', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, ...chartGridAxis() }, y: chartCategoryAxis() }
          }
        }
      },
      {
        title: 'Potential Earnings by Workloads',
        id: 'chartE5Earnings',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'Entra ID', 'Exchange Online'],
            datasets: [{ data: [16, 13, 9, 7, 5], backgroundColor: chartPalette.slice(0, 5), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      }
    ],
    e5exp: [
      {
        title: 'Utilization by Workload',
        id: 'chartE5ExpUtilization',
        config: {
          type: 'bar',
          data: {
            labels: ['Exchange Online', 'Teams', 'SharePoint', 'M365 Apps'],
            datasets: [{ data: [88, 81, 65, 84], backgroundColor: '#0078d4', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: {
              x: { beginAtZero: true, max: 100, ...chartGridAxis(), ticks: { font: { size: 9 }, color: '#616161', callback: value => `${value}%` } },
              y: chartCategoryAxis()
            }
          }
        }
      },
      {
        title: 'Opportunities by Usage Whitespace',
        id: 'chartE5ExpWhitespace',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'Other'],
            datasets: [{ data: [62000, 78000, 44000, 22673], backgroundColor: chartPalette.slice(0, 4), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      },
      {
        title: 'Workload Health',
        id: 'chartE5ExpHealth',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Healthy', 'Unhealthy'],
            datasets: [{ data: [45, 55], backgroundColor: ['#107c10', '#f7941d'], borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(10) } }
        }
      },
      {
        title: 'Workload Recommendations',
        id: 'chartE5ExpRecommendations',
        config: {
          type: 'bar',
          data: {
            labels: ['Drive upsell', 'Drive usage', 'Remediate health'],
            datasets: [{ data: [174, 212, 42], backgroundColor: '#1b3a5c', borderRadius: 4, barThickness: 14 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, ...chartGridAxis() }, y: chartCategoryAxis() }
          }
        }
      },
      {
        title: 'Potential Earnings by Workloads',
        id: 'chartE5ExpEarnings',
        config: {
          type: 'doughnut',
          data: {
            labels: ['Teams', 'SharePoint', 'Intune', 'Exchange Online'],
            datasets: [{ data: [15, 12, 8, 5], backgroundColor: chartPalette.slice(0, 4), borderWidth: 0 }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: chartLegendOptions(9) } }
        }
      }
    ]
  };

  return definitions[opp] || [];
}

function renderCharts(opp) {
  const container = document.getElementById('chartsRow');
  if (!container) return;

  const chartDefinitions = getChartDefinitions(opp);
  destroyCharts();

  if (!chartDefinitions.length) {
    container.innerHTML = '';
    return;
  }

  container.innerHTML = chartDefinitions.map(chart => chartCardMarkup(chart.title, chart.id)).join('');
  chartDefinitions.forEach(chart => {
    new Chart(document.getElementById(chart.id), chart.config);
  });
}

function initToolbarControls() {
  const changeColumnsBtn = document.getElementById('changeColumnsBtn');
  const filterBtn = document.getElementById('filterBtn');
  const searchInput = document.getElementById('tenantSearch');
  const mciChangeColumnsBtn = document.getElementById('mciChangeColumnsBtn');
  const mciFilterBtn = document.getElementById('mciFilterBtn');
  const mciSearchInput = document.getElementById('mciSearch');
  const closeBtn = document.getElementById('utilityFlyoutClose');
  const overlay = document.getElementById('utilityFlyoutOverlay');

  if (changeColumnsBtn) {
    changeColumnsBtn.addEventListener('click', () => openUtilityFlyout('columns'));
  }

  if (filterBtn) {
    filterBtn.addEventListener('click', () => openUtilityFlyout('filters'));
  }

  if (searchInput) {
    searchInput.addEventListener('input', event => {
      searchTermsByTab[currentTab] = event.target.value;
      renderOppView(currentTab);
    });
  }

  if (mciChangeColumnsBtn) {
    mciChangeColumnsBtn.addEventListener('click', () => openUtilityFlyout('columns'));
  }

  if (mciFilterBtn) {
    mciFilterBtn.addEventListener('click', () => openUtilityFlyout('filters'));
  }

  if (mciSearchInput) {
    mciSearchInput.addEventListener('input', event => {
      searchTermsByMciTab[currentMciTab] = event.target.value;
      renderMciView(currentMciTab);
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener('click', closeUtilityFlyout);
  }

  if (overlay) {
    overlay.addEventListener('click', closeUtilityFlyout);
  }

  updateToolbarState();
}

function openUtilityFlyout(type) {
  closeFlyout();
  currentUtilityPanel = type;
  document.getElementById('utilityFlyoutOverlay').classList.remove('hidden');
  document.getElementById('utilityFlyout').classList.remove('hidden');
  renderUtilityFlyout();
}

function closeUtilityFlyout() {
  currentUtilityPanel = null;
  const overlay = document.getElementById('utilityFlyoutOverlay');
  const panel = document.getElementById('utilityFlyout');
  if (overlay) overlay.classList.add('hidden');
  if (panel) panel.classList.add('hidden');
}

function renderUtilityFlyout() {
  if (currentUtilityPanel === 'columns') {
    renderColumnFlyout();
  } else if (currentUtilityPanel === 'filters') {
    renderFilterFlyout();
  }
}

function renderColumnFlyout() {
  const title = document.getElementById('utilityFlyoutTitle');
  const body = document.getElementById('utilityFlyoutBody');
  const footer = document.getElementById('utilityFlyoutFooter');
  const isGrowth = currentPage === 'growth';
  const allColumns = isGrowth ? getAllColumnsForTab(currentTab) : mciData[currentMciTab].columns;
  const selectedColumns = new Set(isGrowth ? getVisibleColumns(currentTab) : getMciVisibleColumns(currentMciTab));
  const contextLabel = isGrowth ? tabDisplayNames[currentTab] : mciTabDisplayNames[currentMciTab];

  title.textContent = 'Customize Columns';
  body.innerHTML = `
    <p class="utility-flyout-intro">Choose which columns appear for ${contextLabel}.</p>
    <p class="utility-selection-count">${selectedColumns.size} of ${allColumns.length} columns selected</p>
    <div class="utility-checkbox-list">
      ${allColumns.map((column, index) => `
        <label class="utility-checkbox-item">
          <input type="checkbox" class="utility-column-checkbox" data-column-index="${index}" ${selectedColumns.has(column) ? 'checked' : ''}>
          <span class="utility-checkbox-text">${escapeHtml(column)}</span>
        </label>
      `).join('')}
    </div>
  `;
  footer.innerHTML = `
    <div class="utility-actions">
      <button class="ocv-btn-secondary" id="utilityResetBtn" type="button">Reset</button>
      <button class="ocv-btn-primary" id="utilityApplyBtn" type="button">Apply</button>
    </div>
  `;

  body.querySelectorAll('.utility-column-checkbox').forEach(checkbox => {
    checkbox.addEventListener('change', updateColumnSelectionCount);
  });

  document.getElementById('utilityResetBtn').addEventListener('click', () => {
    const defaultColumns = new Set((isGrowth ? oppData[currentTab].columns.map(canonicalizeColumn) : mciData[currentMciTab].columns));
    body.querySelectorAll('.utility-column-checkbox').forEach(checkbox => {
      const column = allColumns[Number(checkbox.dataset.columnIndex)];
      checkbox.checked = defaultColumns.has(column);
    });
    updateColumnSelectionCount();
  });

  document.getElementById('utilityApplyBtn').addEventListener('click', () => {
    const nextColumns = [...body.querySelectorAll('.utility-column-checkbox:checked')].map(checkbox => allColumns[Number(checkbox.dataset.columnIndex)]);
    if (isGrowth) {
      activeColumnsByTab[currentTab] = nextColumns;
    } else {
      activeColumnsByMciTab[currentMciTab] = nextColumns;
    }
    closeUtilityFlyout();
    renderCurrentPage();
  });
}

function updateColumnSelectionCount() {
  const body = document.getElementById('utilityFlyoutBody');
  const checked = body.querySelectorAll('.utility-column-checkbox:checked').length;
  const total = body.querySelectorAll('.utility-column-checkbox').length;
  const counter = body.querySelector('.utility-selection-count');
  if (counter) {
    counter.textContent = `${checked} of ${total} columns selected`;
  }
}

function getFilterDefinitions(tab) {
  return currentPage === 'growth' ? (filterConfig[tab] || []) : (mciFilterConfig[tab] || []);
}

function getFilterOptions(tab, field, staticOptions) {
  if (staticOptions) return staticOptions;
  const rows = currentPage === 'growth' ? oppData[tab].rows : mciData[tab].rows;
  return [...new Set(rows
    .map(row => String(getRowValue(row, field) || '').trim())
    .filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function renderFilterFlyout() {
  const title = document.getElementById('utilityFlyoutTitle');
  const body = document.getElementById('utilityFlyoutBody');
  const footer = document.getElementById('utilityFlyoutFooter');
  const activeKey = currentPage === 'growth' ? currentTab : currentMciTab;
  const filters = currentPage === 'growth' ? (activeFiltersByTab[currentTab] || {}) : (activeFiltersByMciTab[currentMciTab] || {});
  const definitions = getFilterDefinitions(activeKey);
  const contextLabel = currentPage === 'growth' ? tabDisplayNames[currentTab] : mciTabDisplayNames[currentMciTab];

  title.textContent = 'Filter';
  body.innerHTML = `
    <p class="utility-filter-hint">Apply filters for ${contextLabel}. Rows must match all selected criteria.</p>
    <div class="utility-filter-groups">
      ${definitions.map(definition => {
        const options = getFilterOptions(activeKey, definition.field, definition.options);
        const selected = new Set(filters[definition.field] || []);
        return `
          <section class="utility-filter-group" data-field="${escapeHtml(definition.field)}">
            <div class="utility-filter-group-title">${escapeHtml(definition.field)}</div>
            <div class="utility-filter-options">
              ${options.length ? options.map(option => `
                <label class="utility-checkbox-item">
                  <input type="checkbox" class="utility-filter-checkbox" value="${escapeHtml(option)}" ${selected.has(option) ? 'checked' : ''}>
                  <span class="utility-checkbox-text">${escapeHtml(option)}</span>
                </label>
              `).join('') : '<div class="utility-filter-empty">No options available</div>'}
            </div>
          </section>
        `;
      }).join('')}
    </div>
  `;
  footer.innerHTML = `
    <div class="utility-actions">
      <button class="ocv-btn-secondary" id="utilityClearFiltersBtn" type="button">Clear All</button>
      <button class="ocv-btn-primary" id="utilityApplyFiltersBtn" type="button">Apply</button>
    </div>
  `;

  document.getElementById('utilityClearFiltersBtn').addEventListener('click', () => {
    body.querySelectorAll('.utility-filter-checkbox').forEach(checkbox => {
      checkbox.checked = false;
    });
  });

  document.getElementById('utilityApplyFiltersBtn').addEventListener('click', () => {
    const nextFilters = {};
    body.querySelectorAll('.utility-filter-group').forEach(group => {
      const field = group.dataset.field;
      const values = [...group.querySelectorAll('.utility-filter-checkbox:checked')].map(checkbox => checkbox.value);
      if (values.length) {
        nextFilters[field] = values;
      }
    });

    if (currentPage === 'growth') {
      activeFiltersByTab[currentTab] = nextFilters;
    } else {
      activeFiltersByMciTab[currentMciTab] = nextFilters;
    }
    closeUtilityFlyout();
    renderCurrentPage();
  });
}
// IN-ROW GIVE FEEDBACK (dropdown per row)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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

// Event delegation for Give Feedback column on MCI and FastTrack tables
function initGlobalFeedbackDelegation() {
  // Handle feedback buttons on MCI and FastTrack tables (same pattern as Growth)
  ['mciTable', 'ftTable'].forEach(tableId => {
    const table = document.getElementById(tableId);
    if (!table) return;

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

        // Close dropdown
        document.querySelectorAll('.row-feedback-dropdown').forEach(d => d.classList.add('hidden'));
        document.querySelectorAll('.row-feedback-btn').forEach(b => b.classList.remove('active'));

        openOcvModal(reason, null, 'table');
      }
    });
  });
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// FLYOUT PANEL (5 types from screenshots 2-6)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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

  // Thumbs up/down in the flyout — handled by inline variant script
}

let currentFlyoutData = null;

function openFlyout(type, rowData) {
  currentFlyoutData = rowData;
  closeUtilityFlyout();
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
      <tr><td>Tenant ID:</td><td>${data['Tenant ID'] || 'â€”'}</td></tr>
      <tr><td>TPID:</td><td>${data['Customer TPID'] || data['TPID'] || 'â€”'}</td></tr>
      <tr><td>Claims:</td><td>â€”</td></tr>
      <tr><td>Claimed workloads:</td><td>â€”</td></tr>
    </table>
    <h4 class="flyout-section-title">Subscriptions</h4>
    <div class="flyout-toolbar">
      <button class="flyout-toolbar-btn">â†“ Download</button>
      <span class="flyout-toolbar-spacer"></span>
      <button class="flyout-toolbar-btn">âŠž Filter</button>
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
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">The total Copilot MAU encompassing licensed and unlicensed users. Additionally, included below is a breakdown of Copilot usage. Unless specified in the metric, MAU can include licensed and unlicensed users. <a href="#" style="color:#0078d4;">Learn more â†—</a></p>
    <table class="flyout-info-table">
      <tr><td><strong>Metric Name</strong></td><td style="text-align:right;"><strong>MAU</strong></td></tr>
      <tr><td>All Copilot MAU</td><td style="text-align:right;">${data['All Copilot MAU'] || 'â€”'}</td></tr>
      <tr><td>Free Copilot MAU (Unlicensed)</td><td style="text-align:right;">${data['Free Copilot Chat MAU (Unlicensed)'] || 'â€”'}</td></tr>
      <tr><td>Copilot MAU (Licensed)</td><td style="text-align:right;">${data['Copilot MAU (Licensed)'] || 'â€”'}</td></tr>
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
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">All Agents MAU provides the total usage of agents all up across any application and user. The table below gives a more granular breakdown of licensed, unlicensed, and agent MAU by extension/application type. <a href="#" style="color:#0078d4;">Learn more â†—</a></p>
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
  const opportunity = data['Copilot Opportunity'] || 'â€”';
  const eligibleSeats = data['Copilot Eligible M365 Seats'] || 'â€”';
  const whitespace = data['Copilot Seats Whitespace'] || 'â€”';
  const freeMau = data['Free Copilot Chat MAU (Unlicensed)'] || '0';
  const pau = data['Copilot PAU'] || '0';
  const allMau = data['All Copilot MAU'] || '0';

  // Calculate penetration percentages
  const copilotPen = pau !== '0' && eligibleSeats !== 'â€”' ?
    Math.round((parseInt(pau.replace(/,/g, '')) / parseInt(eligibleSeats.replace(/,/g, ''))) * 100) + '%' : '0%';
  const freePen = freeMau !== '0' && eligibleSeats !== 'â€”' ?
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
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">The Copilot Opportunity value provides direction on what next steps to take with the customer. Included below are key penetration metrics and AI/ML insights that explain why the customer is categorized as acquire, monetize, or grow. <a href="#" style="color:#0078d4;">Learn more â†—</a></p>
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
    <p style="font-size:13px;color:#616161;margin-bottom:16px;">List of all active New Commerce Cloud Solution Provider (CSP) Promotions related to this solution area. Active promotions do not necessarily meet the eligibility criteria for this tenant. Visit the Pricing workspace for a complete list of eligible promotions across products. <a href="#" style="color:#0078d4;">Learn More â†—</a></p>
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
        <tr><td>Limited time offer: Microsoft 365 Copilot for Allâ€”40% offer</td><td>Percent off</td><td>40%</td><td>06/30/2026</td></tr>
        <tr><td>15% off Microsoft 365 E7 triennial subscription, 300-9,999 licenses</td><td>Percent off</td><td>15%</td><td>12/31/2026</td></tr>
      </tbody>
    </table>
  `;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// OCV FEEDBACK MODAL (Two-step: classify â†’ form)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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
    // For table/flyout feedback, go directly to step 2 (hide back button)
    document.getElementById('ocvBackBtn').classList.add('hidden');
    const typeLabel = reason === 'positive' ? 'compliment' : 'problem';
    showOcvStep2(typeLabel, rowData, source);
  }
}

function showOcvStep2(type, rowData, source) {
  document.getElementById('ocvStep1').classList.add('hidden');
  document.getElementById('ocvStep2').classList.remove('hidden');

  const questionLabel = document.getElementById('ocvQuestionLabel');
  const templateHint = document.getElementById('ocvTemplateHint');
  const contextSection = document.getElementById('ocvContext');
  const contextDetails = document.getElementById('ocvContextDetails');
  const columnsSection = document.getElementById('ocvColumnsSection');
  const columnsGrid = document.getElementById('ocvColumnsGrid');

  // Set question and template based on type
  if (type === 'compliment') {
    questionLabel.innerHTML = 'What did you like? <span class="required">*</span>';
    templateHint.classList.add('hidden');
    templateHint.innerHTML = '';
  } else if (type === 'problem') {
    questionLabel.innerHTML = 'What went wrong? <span class="required">*</span>';
    templateHint.classList.remove('hidden');
    templateHint.innerHTML = `
      <div class="template-title">Please include the following to help us triage faster:</div>
      <ul class="template-list">
        <li><strong>Page/Tab:</strong> Which page or tab were you on?</li>
        <li><strong>What happened:</strong> Describe the issue (e.g., incorrect data, missing info, error)</li>
        <li><strong>Expected behavior:</strong> What did you expect to see instead?</li>
        <li><strong>Frequency:</strong> Does this happen every time or intermittently?</li>
        <li><strong>Impact:</strong> How does this affect your workflow? (blocking, slowing down, minor)</li>
      </ul>
    `;
  } else if (type === 'suggestion') {
    questionLabel.innerHTML = 'What would you like to see improved? <span class="required">*</span>';
    templateHint.classList.remove('hidden');
    templateHint.innerHTML = `
      <div class="template-title">Help us understand your suggestion:</div>
      <ul class="template-list">
        <li><strong>Feature/Area:</strong> Which part of the tool does this relate to?</li>
        <li><strong>Current experience:</strong> What are you doing today and what is the pain point?</li>
        <li><strong>Desired experience:</strong> What would the ideal outcome look like?</li>
        <li><strong>Business impact:</strong> How would this improvement help you or your team?</li>
        <li><strong>Priority:</strong> Nice-to-have or critical for your workflow?</li>
      </ul>
    `;
  }

  // Use stored context if not passed directly
  const ctx = ocvFeedbackContext || {};
  const rd = rowData || ctx.rowData;
  const src = source || ctx.source;

  // Show context for table/flyout sources
  if (src && src !== 'general' && rd) {
    contextSection.classList.remove('hidden');
    const tenantName = rd['Tenant Name'] || 'â€”';
    contextDetails.innerHTML = `
      <div class="ctx-item"><span class="ctx-key">Source:</span><span class="ctx-val">${src === 'flyout' ? 'Flyout Panel' : 'Data Table'}</span></div>
      <div class="ctx-item"><span class="ctx-key">Tab:</span><span class="ctx-val">${currentTab.charAt(0).toUpperCase() + currentTab.slice(1)} Opportunities</span></div>
      <div class="ctx-item"><span class="ctx-key">Tenant:</span><span class="ctx-val">${tenantName}</span></div>
    `;

    // Show columns grid
    columnsSection.classList.remove('hidden');
    const columns = getVisibleColumns(currentTab);
    const displayCols = columns.slice(0, 12);
    columnsGrid.innerHTML = displayCols.map(col => {
      const val = getRowValue(rd, col) || 'â€”';
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
  document.getElementById('ocvBackBtn').classList.remove('hidden');
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// GLOBAL FEEDBACK (top-right nav button)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

function initGlobalFeedback() {
  document.getElementById('globalFeedbackBtn').addEventListener('click', () => {
    openOcvModal('general', null, 'general');
  });
  const mciFbBtn = document.getElementById('mciFeedbackBtn');
  if (mciFbBtn) mciFbBtn.addEventListener('click', () => { openOcvModal('general', null, 'general'); });
  const ftFbBtn = document.getElementById('ftFeedbackBtn');
  if (ftFbBtn) ftFbBtn.addEventListener('click', () => { openOcvModal('general', null, 'general'); });
}

// ═══════════════════════════════════════════════════
// FASTTRACK REFERRALS PAGE
// ═══════════════════════════════════════════════════

const fasttrackColumns = ['Name', 'Referral Source', 'Referral ID', 'Acceptance', 'Referral Status', 'Date'];

const fasttrackData = [
  { 'Name': 'Fabrikam Holdings CEE', 'Referral Source': 'Partner Campaign', 'Referral ID': '002123', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Tarnyba LT', 'Referral Source': 'Partner Campaign', 'Referral ID': '002124', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Legal Services', 'Referral Source': 'Partner Campaign', 'Referral ID': '002121', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Bank BG', 'Referral Source': 'Partner Campaign', 'Referral ID': '002122', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Property Mgmt', 'Referral Source': 'Partner Campaign', 'Referral ID': '002120', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Telco Group', 'Referral Source': 'Partner Campaign', 'Referral ID': '002116', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Energy BG', 'Referral Source': 'Partner Campaign', 'Referral ID': '002117', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Forestry OYJ', 'Referral Source': 'Partner Campaign', 'Referral ID': '002114', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Transit ZRT', 'Referral Source': 'Partner Campaign', 'Referral ID': '002115', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Retail Baltic', 'Referral Source': 'Partner Campaign', 'Referral ID': '002112', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Ventures Group', 'Referral Source': 'Partner Campaign', 'Referral ID': '002113', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Software SE', 'Referral Source': 'Partner Campaign', 'Referral ID': '002111', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Bank D.D.', 'Referral Source': 'Partner Campaign', 'Referral ID': '002109', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Pharmaceuticals PLC', 'Referral Source': 'Partner Campaign', 'Referral ID': '002107', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam University HU', 'Referral Source': 'Partner Campaign', 'Referral ID': '002108', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Finance Bulgaria', 'Referral Source': 'Partner Campaign', 'Referral ID': '002105', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam IT Services ZRT', 'Referral Source': 'Partner Campaign', 'Referral ID': '002106', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Gov Services HU', 'Referral Source': 'Partner Campaign', 'Referral ID': '002103', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Financial Romania S.A.', 'Referral Source': 'Partner Campaign', 'Referral ID': '001946', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 24, 2026' },
  { 'Name': 'Fabrikam Retail SP Z.O.O.', 'Referral Source': 'Partner Campaign', 'Referral ID': '001945', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam Logistics SA', 'Referral Source': 'Partner Campaign', 'Referral ID': '001944', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam Transilvania S.A.', 'Referral Source': 'Partner Campaign', 'Referral ID': '001890', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam Politehnica University', 'Referral Source': 'Partner Campaign', 'Referral ID': '001875', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam Medical University', 'Referral Source': 'Partner Campaign', 'Referral ID': '001873', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam University RO', 'Referral Source': 'Partner Campaign', 'Referral ID': '001874', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 12, 2026' },
  { 'Name': 'Fabrikam Oil & Gas Group', 'Referral Source': 'Partner Campaign', 'Referral ID': '001860', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 5, 2026' },
  { 'Name': 'Fabrikam Airlines PLC', 'Referral Source': 'Partner Campaign', 'Referral ID': '001855', 'Acceptance': 'Pending', 'Referral Status': 'Pending', 'Date': 'March 5, 2026' },
  { 'Name': 'Fabrikam Marketplace PL', 'Referral Source': 'Partner Campaign', 'Referral ID': '001850', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'March 5, 2026' },
  { 'Name': 'Fabrikam Aerospace Corp', 'Referral Source': 'Direct Referral', 'Referral ID': '001845', 'Acceptance': 'Decline', 'Referral Status': 'Expired', 'Date': 'February 28, 2026' },
  { 'Name': 'Fabrikam Telecom AB', 'Referral Source': 'Direct Referral', 'Referral ID': '001840', 'Acceptance': 'Accept', 'Referral Status': 'Completed', 'Date': 'February 28, 2026' },
  { 'Name': 'Fabrikam Nordic Bank', 'Referral Source': 'Partner Campaign', 'Referral ID': '001835', 'Acceptance': 'Accept', 'Referral Status': 'Completed', 'Date': 'February 20, 2026' },
  { 'Name': 'Fabrikam Appliances AB', 'Referral Source': 'Direct Referral', 'Referral ID': '001830', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'February 20, 2026' },
  { 'Name': 'Fabrikam Vehicles AB', 'Referral Source': 'Partner Campaign', 'Referral ID': '001825', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'February 15, 2026' },
  { 'Name': 'Fabrikam Broadcasting FI', 'Referral Source': 'Partner Campaign', 'Referral ID': '001820', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'February 15, 2026' },
  { 'Name': 'Fabrikam Petrochemicals S.A.', 'Referral Source': 'Direct Referral', 'Referral ID': '001815', 'Acceptance': 'Pending', 'Referral Status': 'Pending', 'Date': 'February 10, 2026' },
  { 'Name': 'Fabrikam Telekom HU', 'Referral Source': 'Partner Campaign', 'Referral ID': '001810', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'February 10, 2026' },
  { 'Name': 'Fabrikam Discount Retail', 'Referral Source': 'Partner Campaign', 'Referral ID': '001805', 'Acceptance': 'Accept', 'Referral Status': 'Completed', 'Date': 'February 5, 2026' },
  { 'Name': 'Fabrikam International Banking', 'Referral Source': 'Direct Referral', 'Referral ID': '001800', 'Acceptance': 'Decline', 'Referral Status': 'Expired', 'Date': 'January 28, 2026' },
  { 'Name': 'Fabrikam Communications AB', 'Referral Source': 'Partner Campaign', 'Referral ID': '001795', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'January 28, 2026' },
  { 'Name': 'Fabrikam Savings Bank CZ', 'Referral Source': 'Partner Campaign', 'Referral ID': '001790', 'Acceptance': 'Accept', 'Referral Status': 'Active', 'Date': 'January 20, 2026' },
];

let ftSearchTerm = '';
let ftCurrentPage = 1;
let ftPageSize = 25;
let ftFilterStatus = 'all';
let ftFilterAcceptance = 'all';

function getFilteredFtData() {
  return fasttrackData.filter(row => {
    if (ftSearchTerm) {
      const term = ftSearchTerm.toLowerCase();
      const match = row['Name'].toLowerCase().includes(term) ||
                    row['Referral ID'].toLowerCase().includes(term);
      if (!match) return false;
    }
    if (ftFilterStatus !== 'all' && row['Referral Status'] !== ftFilterStatus) return false;
    if (ftFilterAcceptance !== 'all' && row['Acceptance'] !== ftFilterAcceptance) return false;
    return true;
  });
}

function renderFasttrackView() {
  const filtered = getFilteredFtData();
  const total = filtered.length;
  const totalPages = Math.ceil(total / ftPageSize);
  if (ftCurrentPage > totalPages) ftCurrentPage = 1;
  const start = (ftCurrentPage - 1) * ftPageSize;
  const pageRows = filtered.slice(start, start + ftPageSize);

  // KPIs
  const kpiRow = document.getElementById('ftKpiRow');
  const totalReferrals = fasttrackData.length;
  const activeCount = fasttrackData.filter(r => r['Referral Status'] === 'Active').length;
  const pendingCount = fasttrackData.filter(r => r['Referral Status'] === 'Pending').length;
  const completedCount = fasttrackData.filter(r => r['Referral Status'] === 'Completed').length;
  const expiredCount = fasttrackData.filter(r => r['Referral Status'] === 'Expired').length;
  const acceptRate = Math.round((fasttrackData.filter(r => r['Acceptance'] === 'Accept').length / totalReferrals) * 100);

  kpiRow.innerHTML = [
    { value: totalReferrals, label: 'Total Referrals' },
    { value: activeCount, label: 'Active', color: '#107c10' },
    { value: pendingCount, label: 'Pending', color: '#f7941d' },
    { value: completedCount, label: 'Completed', color: '#0078d4' },
    { value: expiredCount, label: 'Expired', color: '#d13438' },
    { value: acceptRate + '%', label: 'Acceptance Rate' },
  ].map(k => `<div class="kpi-tile ft-kpi-tile" data-filter="${k.label}">
    <span class="kpi-value"${k.color ? ` style="color:${k.color}"` : ''}>${k.value}</span>
    <span class="kpi-label">${k.label}</span>
  </div>`).join('');

  // Table head
  const head = document.getElementById('ftTableHead');
  head.innerHTML = fasttrackColumns.map(c => {
    const sortable = c === 'Referral ID' ? ' class="sortable"' : '';
    return `<th${sortable}>${c}${c === 'Referral ID' ? ' <span class="sort-arrow">&#8595;</span>' : ''}</th>`;
  }).join('');

  // Table body
  const body = document.getElementById('ftTableBody');
  if (!pageRows.length) {
    body.innerHTML = `<tr><td colspan="${fasttrackColumns.length}" style="padding:20px;text-align:center;color:#616161;">No referrals match the current search and filters.</td></tr>`;
  } else {
    body.innerHTML = pageRows.map((row, idx) => {
      const cells = fasttrackColumns.map(col => {
        let val = row[col] || '';
        if (col === 'Name') {
          val = `<a href="#" class="ft-name-link">${escapeHtml(val)}</a>`;
        } else if (col === 'Acceptance') {
          const cls = val === 'Accept' ? 'ft-accept' : val === 'Decline' ? 'ft-decline' : 'ft-pending';
          val = `<span class="ft-badge ${cls}">${val}</span>`;
        } else if (col === 'Referral Status') {
          const cls = val === 'Active' ? 'ft-status-active' : val === 'Completed' ? 'ft-status-completed' : val === 'Expired' ? 'ft-status-expired' : 'ft-status-pending';
          val = `<span class="ft-badge ${cls}">${val}</span>`;
        }
        return `<td>${val}</td>`;
      }).join('');
      return `<tr>${cells}</tr>`;
    }).join('');
  }

  // Pagination
  const summary = document.getElementById('ftResultsSummary');
  const showing = Math.min(start + ftPageSize, total);
  summary.textContent = `${start + 1} - ${showing} of ${total} results`;

  const pageIndicator = document.getElementById('ftPageIndicator');
  pageIndicator.textContent = ftCurrentPage;

  document.getElementById('ftPrevPage').disabled = ftCurrentPage <= 1;
  document.getElementById('ftNextPage').disabled = ftCurrentPage >= totalPages;

  // Charts
  renderFtCharts();
}

let ftChartInstances = [];

function renderFtCharts() {
  const container = document.getElementById('ftChartsRow');
  if (!container) return;

  // Destroy existing charts
  ftChartInstances.forEach(c => c.destroy());
  ftChartInstances = [];

  // Calculate data for charts
  const statusCounts = {};
  const sourceCounts = {};
  const monthCounts = {};

  fasttrackData.forEach(row => {
    const status = row['Referral Status'];
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    const source = row['Referral Source'];
    sourceCounts[source] = (sourceCounts[source] || 0) + 1;

    const dateStr = row['Date'];
    const month = dateStr.split(' ').slice(0, 2).join(' ').replace(',', '');
    monthCounts[month] = (monthCounts[month] || 0) + 1;
  });

  const statusColors = { 'Active': '#107c10', 'Pending': '#f7941d', 'Completed': '#0078d4', 'Expired': '#d13438' };
  const sourceColors = ['#0078d4', '#1b3a5c', '#4fc3f7'];

  container.innerHTML = `
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">By Status</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="ftChartStatus"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">By Source</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="ftChartSource"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-card-header">
        <div><div class="chart-card-title">By Month</div><div class="chart-card-subtitle">${chartSubtitle}</div></div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="ftChartMonth"></canvas></div>
    </div>
  `;

  // Status pie chart
  const statusLabels = Object.keys(statusCounts);
  const statusData = Object.values(statusCounts);
  ftChartInstances.push(new Chart(document.getElementById('ftChartStatus'), {
    type: 'doughnut',
    data: {
      labels: statusLabels,
      datasets: [{ data: statusData, backgroundColor: statusLabels.map(l => statusColors[l] || '#999') }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 12 } } }
    }
  }));

  // Source bar chart
  const sourceLabels = Object.keys(sourceCounts);
  const sourceData = Object.values(sourceCounts);
  ftChartInstances.push(new Chart(document.getElementById('ftChartSource'), {
    type: 'bar',
    data: {
      labels: sourceLabels,
      datasets: [{ data: sourceData, backgroundColor: sourceColors.slice(0, sourceLabels.length), borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'x',
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: '#e8e8e8', drawBorder: false }, ticks: { font: { size: 10 } } },
        x: { grid: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  }));

  // Monthly bar chart
  const monthLabels = Object.keys(monthCounts).reverse();
  const monthData = monthLabels.map(m => monthCounts[m]);
  ftChartInstances.push(new Chart(document.getElementById('ftChartMonth'), {
    type: 'bar',
    data: {
      labels: monthLabels,
      datasets: [{ data: monthData, backgroundColor: '#0078d4', borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: '#e8e8e8', drawBorder: false }, ticks: { font: { size: 10 } } },
        x: { grid: { display: false }, ticks: { font: { size: 9 }, maxRotation: 45 } }
      }
    }
  }));
}

function initFasttrackControls() {
  const searchInput = document.getElementById('ftSearch');
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      ftSearchTerm = searchInput.value.trim();
      ftCurrentPage = 1;
      renderFasttrackView();
    });
  }

  const statusFilter = document.getElementById('ftFilterStatus');
  if (statusFilter) {
    statusFilter.addEventListener('change', () => {
      ftFilterStatus = statusFilter.value;
      ftCurrentPage = 1;
      renderFasttrackView();
    });
  }

  const acceptFilter = document.getElementById('ftFilterAcceptance');
  if (acceptFilter) {
    acceptFilter.addEventListener('change', () => {
      ftFilterAcceptance = acceptFilter.value;
      ftCurrentPage = 1;
      renderFasttrackView();
    });
  }

  document.getElementById('ftPrevPage').addEventListener('click', () => {
    if (ftCurrentPage > 1) { ftCurrentPage--; renderFasttrackView(); }
  });
  document.getElementById('ftNextPage').addEventListener('click', () => {
    const total = getFilteredFtData().length;
    if (ftCurrentPage < Math.ceil(total / ftPageSize)) { ftCurrentPage++; renderFasttrackView(); }
  });

  // KPI tile click filtering
  document.getElementById('ftKpiRow').addEventListener('click', (e) => {
    const tile = e.target.closest('.ft-kpi-tile');
    if (!tile) return;
    const filter = tile.dataset.filter;
    if (filter === 'Total Referrals' || filter === 'Acceptance Rate') {
      ftFilterStatus = 'all';
      ftFilterAcceptance = 'all';
    } else if (filter === 'Active' || filter === 'Completed' || filter === 'Pending' || filter === 'Expired') {
      if (ftFilterStatus === filter) { ftFilterStatus = 'all'; }
      else { ftFilterStatus = filter; ftFilterAcceptance = 'all'; }
    }
    // Sync dropdowns
    document.getElementById('ftFilterStatus').value = ftFilterStatus;
    document.getElementById('ftFilterAcceptance').value = ftFilterAcceptance;
    ftCurrentPage = 1;
    renderFasttrackView();
  });
}

