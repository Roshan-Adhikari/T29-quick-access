// LiveAccess - Google Sheets Live Quick Access Search Engine
// Handles GIS OAuth2, dynamic two-step search, caching, history and rendering.

// State variables
let tokenClient;
let accessToken = null;
let sheetHeaders = null;
let emailIndex = null;
let recentLookups = [];

// DOM Elements
const btnSettings = document.getElementById('btnSettings');
const btnCloseSettings = document.getElementById('btnCloseSettings');
const modalSettings = document.getElementById('modalSettings');
const formSettings = document.getElementById('formSettings');
const btnSaveSettings = document.getElementById('btnSaveSettings');
const btnUseDefaults = document.getElementById('btnUseDefaults');

const btnConnect = document.getElementById('btnConnect');
const userProfile = document.getElementById('userProfile');
const userAvatar = document.getElementById('userAvatar');
const userName = document.getElementById('userName');
const userEmail = document.getElementById('userEmail');
const btnSignOut = document.getElementById('btnSignOut');

const sectionSetup = document.getElementById('sectionSetup');
const btnStartSetup = document.getElementById('btnStartSetup');
const sectionSearch = document.getElementById('sectionSearch');

const txtSearchEmail = document.getElementById('txtSearchEmail');
const btnSearch = document.getElementById('btnSearch');
const btnClearCache = document.getElementById('btnClearCache');
const indexStatus = document.getElementById('indexStatus');

const loaderContainer = document.getElementById('loaderContainer');
const loaderHeadline = document.getElementById('loaderHeadline');
const loaderSubtext = document.getElementById('loaderSubtext');

const dashboardResults = document.getElementById('dashboardResults');
const codeRawJSON = document.getElementById('codeRawJSON');
const historyList = document.getElementById('historyList');

// Normalizes header keys to make lookup resilient to spaces, case, and newlines
function normalizeHeaderKey(key) {
  if (!key) return '';
  return key.toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Check configuration on load
document.addEventListener('DOMContentLoaded', () => {
  loadConfig();
  checkCachedSession();
  updateHistoryUI();
  
  // Register basic event listeners
  btnSettings.addEventListener('click', () => showSettingsModal(true));
  btnCloseSettings.addEventListener('click', () => showSettingsModal(false));
  btnStartSetup.addEventListener('click', () => showSettingsModal(true));
  
  btnSaveSettings.addEventListener('click', saveConfig);
  btnUseDefaults.addEventListener('click', loadSampleConfig);
  
  btnConnect.addEventListener('click', connectGoogleAccount);
  btnSignOut.addEventListener('click', signOutGoogle);
  
  btnSearch.addEventListener('click', triggerSearch);
  txtSearchEmail.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') triggerSearch();
  });
  btnClearCache.addEventListener('click', () => {
    fetchSpreadsheetIndex(true);
  });
});

// Extract spreadsheet ID from Google Sheet URL if necessary
function extractSpreadsheetId(input) {
  if (!input) return '';
  input = input.trim();
  // Match standard spreadsheet URL pattern: /spreadsheets/d/[ID]/
  const match = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return input;
}

// Load Settings from LocalStorage
function loadConfig() {
  const defaultClientId = '837920518571-n8mv0vkh69iubhi0p8nfgb1f7cf2v3dm.apps.googleusercontent.com';
  const defaultSpreadsheetId = '1yvC466_OqOeUotT8okfbKFE3QR028VzxCmsZL5GLGR8';

  const clientId = localStorage.getItem('cfg_client_id') || defaultClientId;
  const spreadsheetId = localStorage.getItem('cfg_spreadsheet_id') || defaultSpreadsheetId;
  const sheetName = localStorage.getItem('cfg_sheet_name') || 'Master Data';
  const searchColumn = localStorage.getItem('cfg_search_column') || 'Email id';

  document.getElementById('cfgClientId').value = clientId;
  document.getElementById('cfgSpreadsheetId').value = spreadsheetId;
  document.getElementById('cfgSheetName').value = sheetName;
  document.getElementById('cfgSearchColumn').value = searchColumn;

  // Save to localStorage if not already present
  if (!localStorage.getItem('cfg_client_id')) localStorage.setItem('cfg_client_id', clientId);
  if (!localStorage.getItem('cfg_spreadsheet_id')) localStorage.setItem('cfg_spreadsheet_id', spreadsheetId);
  if (!localStorage.getItem('cfg_sheet_name')) localStorage.setItem('cfg_sheet_name', sheetName);
  if (!localStorage.getItem('cfg_search_column')) localStorage.setItem('cfg_search_column', searchColumn);

  if (clientId && spreadsheetId) {
    // Show search area, hide setup welcome card
    sectionSetup.classList.add('hidden');
    sectionSearch.classList.remove('hidden');
    initGoogleAuth();
  } else {
    // Show setup welcome card
    sectionSetup.classList.remove('hidden');
    sectionSearch.classList.add('hidden');
  }
}

// Load default placeholders for testing
function loadSampleConfig() {
  document.getElementById('cfgClientId').value = '837920518571-n8mv0vkh69iubhi0p8nfgb1f7cf2v3dm.apps.googleusercontent.com';
  document.getElementById('cfgSpreadsheetId').value = '1yvC466_OqOeUotT8okfbKFE3QR028VzxCmsZL5GLGR8';
  document.getElementById('cfgSheetName').value = 'Master Data';
  document.getElementById('cfgSearchColumn').value = 'Email id';
}

// Save Settings to LocalStorage
function saveConfig() {
  const clientId = document.getElementById('cfgClientId').value.trim();
  let spreadsheetId = document.getElementById('cfgSpreadsheetId').value.trim();
  const sheetName = document.getElementById('cfgSheetName').value.trim();
  const searchColumn = document.getElementById('cfgSearchColumn').value.trim();

  if (!clientId || !spreadsheetId || !sheetName || !searchColumn) {
    alert('Please fill out all fields in the configuration.');
    return;
  }

  // Clean spreadsheet ID if it was pasted as a full URL
  spreadsheetId = extractSpreadsheetId(spreadsheetId);

  localStorage.setItem('cfg_client_id', clientId);
  localStorage.setItem('cfg_spreadsheet_id', spreadsheetId);
  localStorage.setItem('cfg_sheet_name', sheetName);
  localStorage.setItem('cfg_search_column', searchColumn);

  // Clear in-memory index on sheet configuration changes
  sheetHeaders = null;
  emailIndex = null;
  updateIndexStatus('Config changed. Reconnecting...', 'yellow');

  showSettingsModal(false);
  loadConfig();
}

function showSettingsModal(show) {
  if (show) {
    modalSettings.classList.remove('hidden');
  } else {
    modalSettings.classList.add('hidden');
  }
}

// Initialize Google OAuth2 Token Client
function initGoogleAuth() {
  const clientId = localStorage.getItem('cfg_client_id');
  if (!clientId || clientId.startsWith('YOUR_')) return;

  try {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
      callback: (tokenResponse) => {
        if (tokenResponse && tokenResponse.access_token) {
          accessToken = tokenResponse.access_token;
          
          // Cache access token for 1 hour
          const expiryTime = Date.now() + (tokenResponse.expires_in * 1000);
          localStorage.setItem('oauth_access_token', accessToken);
          localStorage.setItem('oauth_token_expiry', expiryTime);
          
          // Request user profile info
          fetchUserProfile(accessToken);
          
          showAuthenticatedState(true);
          updateIndexStatus('Connected. Caching index...', 'yellow');
          
          // Prime spreadsheet index download
          fetchSpreadsheetIndex();
        }
      },
    });
  } catch (err) {
    console.error('Failed to initialize GIS SDK:', err);
    updateIndexStatus('Google SDK Load Error', 'red');
  }
}

// Check if cached token is still valid
function checkCachedSession() {
  const token = localStorage.getItem('oauth_access_token');
  const expiry = localStorage.getItem('oauth_token_expiry');
  
  if (token && expiry && Date.now() < parseInt(expiry)) {
    accessToken = token;
    showAuthenticatedState(true);
    updateIndexStatus('Session active. Loading index...', 'yellow');
    
    // Fetch profile
    fetchUserProfile(accessToken);
    
    // Load sheet index
    setTimeout(() => {
      fetchSpreadsheetIndex();
    }, 500);
    return true;
  }
  showAuthenticatedState(false);
  return false;
}

// Connect to Google Account
function connectGoogleAccount() {
  if (!tokenClient) {
    const clientId = localStorage.getItem('cfg_client_id');
    if (!clientId) {
      showSettingsModal(true);
      alert('Please configure your Google OAuth Client ID first.');
      return;
    }
    initGoogleAuth();
  }
  
  if (tokenClient) {
    // Use prompt: 'consent' to guarantee a new refresh/access token if necessary, or empty string to do standard token refresh
    tokenClient.requestAccessToken({ prompt: '' });
  } else {
    alert('Google Identity Services SDK could not be initialized. Check console or verify your Client ID.');
  }
}

// Sign out from Google Account
function signOutGoogle() {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {
      console.log('Access token revoked');
    });
  }
  
  accessToken = null;
  sheetHeaders = null;
  emailIndex = null;
  localStorage.removeItem('oauth_access_token');
  localStorage.removeItem('oauth_token_expiry');
  
  showAuthenticatedState(false);
  updateIndexStatus('Disconnected.', 'red');
  
  // Clear results
  dashboardResults.classList.add('hidden');
}

// Show/Hide authenticated profile UI
function showAuthenticatedState(isAuthenticated) {
  if (isAuthenticated) {
    btnConnect.classList.add('hidden');
    userProfile.classList.remove('hidden');
    
    txtSearchEmail.removeAttribute('disabled');
    btnSearch.removeAttribute('disabled');
  } else {
    btnConnect.classList.remove('hidden');
    userProfile.classList.add('hidden');
    
    txtSearchEmail.setAttribute('disabled', 'true');
    btnSearch.setAttribute('disabled', 'true');
    
    userName.textContent = 'User Account';
    userEmail.textContent = 'email@domain.com';
    userAvatar.textContent = 'U';
  }
}

// Fetch basic user profile info from Google API
async function fetchUserProfile(token) {
  try {
    const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (res.ok) {
      const profile = await res.json();
      userName.textContent = profile.name || 'User Account';
      userEmail.textContent = profile.email || '';
      userAvatar.textContent = (profile.given_name || profile.name || 'U').charAt(0).toUpperCase();
    }
  } catch (err) {
    console.error('Failed to load user profile info:', err);
  }
}

// Sheets API call wrapper
async function callSheetsAPI(spreadsheetId, range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;
  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    }
  });

  if (!response.ok) {
    if (response.status === 401) {
      signOutGoogle();
      throw new Error('Your session expired. Please sign in again.');
    }
    const errObj = await response.json().catch(() => ({}));
    const errMsg = errObj.error ? errObj.error.message : response.statusText;
    throw new Error(`Sheets API error: ${response.status} - ${errMsg}`);
  }
  return await response.json();
}

// Convert index column index to Excel column letter
function getColumnLetter(colIndex) {
  let temp;
  let letter = '';
  colIndex = colIndex + 1;
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

// Helper to update spreadsheet loading and connection status in UI
function updateIndexStatus(text, colorClass) {
  const dot = indexStatus.querySelector('.status-dot');
  const txt = indexStatus.querySelector('.status-text');
  
  dot.className = 'status-dot';
  dot.classList.add(`dot-${colorClass}`);
  txt.textContent = text;
}

// Helper to show loading bar
function showLoader(show, headline = '', subtext = '') {
  if (show) {
    loaderContainer.classList.remove('hidden');
    loaderHeadline.textContent = headline;
    loaderSubtext.textContent = subtext;
  } else {
    loaderContainer.classList.add('hidden');
  }
}

// Fetch Spreadsheet Index (Headers & Email Column)
async function fetchSpreadsheetIndex(forceRefresh = false) {
  if (!accessToken) return;

  const rawSpreadsheetId = localStorage.getItem('cfg_spreadsheet_id');
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  const sheetName = localStorage.getItem('cfg_sheet_name') || 'Master Data';
  const searchColumn = localStorage.getItem('cfg_search_column') || 'Email id';

  if (!spreadsheetId) {
    updateIndexStatus('Spreadsheet not configured.', 'red');
    return;
  }

  // Use cached values if available
  if (!forceRefresh && sheetHeaders && emailIndex) {
    return { headers: sheetHeaders, emails: emailIndex };
  }

  try {
    updateIndexStatus('Loading metadata...', 'yellow');
    showLoader(true, 'Reading Header structure...', 'Fetching column headers to locate email indices...');

    // 1. Fetch headers row A1:ZZ1
    const headersData = await callSheetsAPI(spreadsheetId, `'${sheetName}'!A1:ZZ1`);
    if (!headersData.values || headersData.values.length === 0) {
      throw new Error('Spreadsheet exists but returned no data cells in the first row.');
    }

    sheetHeaders = headersData.values[0];

    // Find the email column
    const normSearchColumn = normalizeHeaderKey(searchColumn);
    const emailColIndex = sheetHeaders.findIndex(h => normalizeHeaderKey(h) === normSearchColumn);

    if (emailColIndex === -1) {
      throw new Error(`Search Column "${searchColumn}" not found in headers. Please check spelling in Connection Settings.`);
    }

    const colLetter = getColumnLetter(emailColIndex);

    // 2. Fetch the entire email column (e.g. E2:E)
    showLoader(true, 'Downloading Email Index...', `Downloading column ${colLetter} for over 50,000+ entries. This matches records quickly.`);
    
    // Fetch values from row 2 onwards
    const emailData = await callSheetsAPI(spreadsheetId, `'${sheetName}'!${colLetter}2:${colLetter}`);
    
    // Parse emails, trim and convert to array
    emailIndex = emailData.values ? emailData.values.map(row => (row[0] || '').trim().toLowerCase()) : [];

    showLoader(false);
    updateIndexStatus(`Index cached: ${emailIndex.length.toLocaleString()} records.`, 'green');
    
    return { headers: sheetHeaders, emails: emailIndex };
  } catch (err) {
    showLoader(false);
    updateIndexStatus('Indexing failed.', 'red');
    console.error('Indexing error:', err);
    
    let adviceText = '';
    if (err.message === 'Failed to fetch' || err.message.includes('Failed to fetch')) {
      adviceText = '\n\nTroubleshooting tips:\n1. Verify your Spreadsheet ID (ensure it is just the ID, not the full URL).\n2. Disable any adblockers or privacy extensions (like Brave Shields) for this page, as they may block Google API requests.\n3. Make sure you are connected to the internet.';
    }
    
    alert(`Spreadsheet Indexing failed: ${err.message}${adviceText}\n\nPlease verify Spreadsheet ID, Sheet Name, and that your Google Account has permissions.`);
  }
}

// Trigger query for email
function triggerSearch() {
  const emailInput = txtSearchEmail.value.trim();
  if (!emailInput) {
    alert('Please enter a student email to search.');
    return;
  }
  
  if (!emailInput.includes('@')) {
    alert('Please enter a valid email address.');
    return;
  }
  
  searchStudent(emailInput);
}

// Search Student email and display row
async function searchStudent(emailAddress) {
  if (!accessToken) {
    alert('Please connect your Google account first.');
    return;
  }

  const queryEmail = emailAddress.trim().toLowerCase();

  try {
    // Ensure index is loaded
    const index = await fetchSpreadsheetIndex();
    if (!index) return;

    const { headers, emails } = index;

    showLoader(true, 'Searching Database...', `Locating record matching "${queryEmail}"...`);

    // Find index of email
    const matchIndex = emails.indexOf(queryEmail);
    if (matchIndex === -1) {
      showLoader(false);
      alert(`Student with email ID "${emailAddress}" not found in the live sheet index.\n\nDouble check the email or click "Refresh Index" if the record was recently added.`);
      return;
    }

    // Row number is matchIndex + 2 (since column retrieval started at row 2)
    const rowNum = matchIndex + 2;

    showLoader(true, 'Retrieving Profile Data...', `Downloading row ${rowNum} details from sheet...`);

    const rawSpreadsheetId = localStorage.getItem('cfg_spreadsheet_id');
    const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
    const sheetName = localStorage.getItem('cfg_sheet_name') || 'Master Data';
    const lastColLetter = getColumnLetter(headers.length - 1);

    // Fetch details of this row
    const rowRange = `'${sheetName}'!A${rowNum}:${lastColLetter}${rowNum}`;
    const rowDataResponse = await callSheetsAPI(spreadsheetId, rowRange);

    if (!rowDataResponse.values || rowDataResponse.values.length === 0) {
      throw new Error(`Successfully located row ${rowNum} but no cells were returned.`);
    }

    const rowData = rowDataResponse.values[0];

    // Build normalized map
    const student = {};
    headers.forEach((header, idx) => {
      const normKey = normalizeHeaderKey(header);
      student[normKey] = rowData[idx] !== undefined ? rowData[idx].trim() : '';
    });

    // Helper fields
    student._name = getVal(student, ['Student Name'], 'Student Record');
    student._email = queryEmail;
    student._row = rowNum;

    // Display student data
    displayStudentDetails(student, headers, rowData);
    
    // Add to lookups list
    addToHistory(student._name, queryEmail);
    
    showLoader(false);
  } catch (err) {
    showLoader(false);
    console.error('Search query error:', err);
    alert(`Search query failed: ${err.message}`);
  }
}

// Retrieve values by checking multiple possible field variations
function getVal(student, possibleKeys, defaultValue = '-') {
  for (const key of possibleKeys) {
    const norm = normalizeHeaderKey(key);
    if (student[norm] !== undefined && student[norm] !== '') {
      return student[norm];
    }
  }
  return defaultValue;
}

// Format status classes based on string values
function applyStatusBadge(element, statusText) {
  if (!element) return;
  element.textContent = statusText || '-';
  element.className = 'val badge-container';
  
  if (!statusText || statusText === '-') {
    return;
  }
  
  const text = statusText.toLowerCase();
  const badge = document.createElement('span');
  badge.textContent = statusText;
  badge.className = 'badge-status';
  
  if (text.includes('disbursed') || text.includes('active') || text.includes('done') || text.includes('yes') || text.includes('closed')) {
    badge.classList.add('nbfc-disbursed');
  } else if (text.includes('dropout') || text.includes('reject') || text.includes('failed') || text.includes('no') || text.includes('foreclosed')) {
    badge.classList.add('nbfc-dropout');
  } else if (text.includes('pending') || text.includes('progress') || text.includes('submitted') || text.includes('process')) {
    badge.classList.add('nbfc-pending');
  } else {
    badge.className = 'badge';
  }
  
  element.innerHTML = '';
  element.appendChild(badge);
}

// Update UI dashboard labels with matched record data
function displayStudentDetails(student, headers, rowData) {
  // Update banner profile card
  document.getElementById('lblStudentName').textContent = getVal(student, ['Student Name']);
  document.getElementById('badgeCohort').textContent = getVal(student, ['Cohort ID', 'Cohort']);
  document.getElementById('badgeCategory').textContent = getVal(student, ['Category', 'Test Category']);
  
  document.getElementById('lblStudentEmail').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M21.75 6.75v10.5a2.25 2.25 0 01-2.25 2.25h-15a2.25 2.25 0 01-2.25-2.25V6.75m19.5 0A2.25 2.25 0 0019.5 4.5h-15a2.25 2.25 0 00-2.25 2.25m19.5 0v.243a2.25 2.25 0 01-1.07 1.916l-7.5 4.615a2.25 2.25 0 01-2.36 0L3.32 8.91a2.25 2.25 0 01-1.07-1.916V6.75"/></svg>
    ${getVal(student, ['Email id', 'Email'])}`;
    
  document.getElementById('lblStudentPhone').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M2.25 6.622M2.25 9c0 5.108 4.027 9.24 9 9.24 2.894 0 5.5-.327 7.747-.94a.75.75 0 00.563-.647c.187-1.393.308-2.8-.073-4.148a.75.75 0 00-.702-.513H15.02a.75.75 0 00-.69.44l-.79 1.58A13.23 13.23 0 017.5 9.79l1.58-.79a.75.75 0 00.44-.69v-3.72a.75.75 0 00-.513-.702A24.819 24.819 0 004.1 3.81a.75.75 0 00-.647.563C2.827 6.6 2.25 9 2.25 9z" stroke-linecap="round" stroke-linejoin="round"/></svg>
    ${getVal(student, ['Mobile'])}`;
    
  document.getElementById('lblStudentAltPhone').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M2.25 6.75c0 8.284 6.716 15 15 15h2.25a2.25 2.25 0 002.25-2.25v-1.372c0-.516-.351-.966-.852-1.091l-4.423-1.106c-.44-.11-.902.055-1.173.417l-.97 1.293c-2.824-1.802-5.14-4.118-6.942-6.942l1.293-.97c.362-.271.527-.834.417-1.173L6.963 3.102a1.125 1.125 0 00-1.091-.852H4.5A2.25 2.25 0 002.25 4.5v2.25z" stroke-linecap="round" stroke-linejoin="round"/></svg>
    Alt: ${getVal(student, ['Alternate Number'])}`;

  // Status banners
  const nbfcStatusText = getVal(student, ['NBFC Status']);
  const closedStatusText = getVal(student, ['Closed/Not closed', 'Closed']);
  document.getElementById('lblBannerNBFCStatus').textContent = nbfcStatusText;
  document.getElementById('lblBannerClosedStatus').textContent = closedStatusText;

  // Banner status styling color overrides
  const bannerNbfc = document.getElementById('lblBannerNBFCStatus');
  if (nbfcStatusText.toLowerCase().includes('disbursed')) {
    bannerNbfc.style.color = 'var(--color-green)';
  } else if (nbfcStatusText.toLowerCase().includes('dropout') || nbfcStatusText.toLowerCase().includes('reject')) {
    bannerNbfc.style.color = 'var(--color-red)';
  } else {
    bannerNbfc.style.color = 'var(--color-yellow)';
  }

  // Section 1: Course & Fees
  document.getElementById('valCourseName').textContent = getVal(student, ['Course Name']);
  document.getElementById('valCommonName').textContent = getVal(student, ['Common Name']);
  document.getElementById('valBatchStart').textContent = getVal(student, ['Batch Start Date']);
  document.getElementById('valFeeNoGST').textContent = getVal(student, ['Course Fees(Exclusive of GST)', 'Course Fees (Exclusive of GST)']);
  document.getElementById('valFeeWithGST').textContent = getVal(student, ['Course Fees ( Inclusive of GST)', 'Course Fees (Inclusive of GST)']);
  document.getElementById('valGSTStatus').textContent = getVal(student, ['GST Yes/NO', 'GST YesNo']);
  document.getElementById('valGSTTaken').textContent = getVal(student, ['GST Taken']);

  // Section 2: Scholarship & Net Fees
  document.getElementById('valScholarshipCat').textContent = getVal(student, ['Scholarship Category']);
  document.getElementById('valScholarshipPct').textContent = getVal(student, ['Scholarship %']);
  document.getElementById('valScholarshipAmt').textContent = getVal(student, ['Scholarship Amount on (Column J)', 'Scholarship Amount on\n(Column J)', 'Scholarship Amount']);
  document.getElementById('valEffectiveFee').textContent = getVal(student, ['Effective fee Column (J-P-M-AL) Without GST', 'Effective fee\nColumn (J-P-M-AL) Without GST', 'Effective Fee']);
  document.getElementById('valActualFullFee').textContent = getVal(student, ['Actual Full Fee (W/O GST) - Vineet', 'Actual Full Fee']);
  document.getElementById('valActualRegFee').textContent = getVal(student, ['Actual Registration Fee']);
  document.getElementById('valActualRemFee').textContent = getVal(student, ['Actual Remaining Fee']);

  // Section 3: Test & Registration Fees
  document.getElementById('valTestFeePaid').textContent = getVal(student, ['Test Fee Paid', 'Test Fee Paid ']);
  document.getElementById('valTestFeeDate').textContent = getVal(student, ['Test fee Date']);
  document.getElementById('valTestFeeUTR').textContent = getVal(student, ['Test Fee UTR']);
  document.getElementById('valRegFeePaid').textContent = getVal(student, ['Reg. fee paid']);
  document.getElementById('valRegFeeDate').textContent = getVal(student, ['Date of Reg. Fee', 'Date of Reg Fee']);
  document.getElementById('valRegFeeUTR').textContent = getVal(student, ['UTR for Registration Fees']);
  document.getElementById('valStudentCat').textContent = getVal(student, ['Student Category', 'Category']);

  // Section 4: Loan & NBFC Status
  document.getElementById('valPaymentMode').textContent = getVal(student, ['Payment Mode']);
  document.getElementById('valNBFCName').textContent = getVal(student, ['NBFC NAME', 'NBFC']);
  document.getElementById('valLenderName').textContent = getVal(student, ['Propelld - Lender Name', 'Propelld Lender Name']);
  document.getElementById('valNBFCAppID').textContent = getVal(student, ['App ID', 'Application ID']);
  applyStatusBadge(document.getElementById('valNBFCStatus'), nbfcStatusText);
  document.getElementById('valForeclosureStatus').textContent = getVal(student, ['Foreclosure/Refund Status']);
  
  document.getElementById('valLoanAmount').textContent = getVal(student, ['Loan Amount']);
  document.getElementById('valDisbursedAmount').textContent = getVal(student, ['Disbursed Amount/Direct Payment', 'Disbursed Amount']);
  document.getElementById('valDisbursedDate').textContent = getVal(student, ['Disbursed/Paid Date', 'Disbursed Date']);
  document.getElementById('valDisbursalUTR').textContent = getVal(student, ['UTR']);
  document.getElementById('valOverallAmountPaid').textContent = getVal(student, ['Over all Amount Paid']);
  document.getElementById('valBajajAdjusted').textContent = getVal(student, ['Bajaj Adjusted Amount']);
  
  document.getElementById('valEMIAmount').textContent = getVal(student, ['EMI Amount']);
  document.getElementById('valEMITenure').textContent = getVal(student, ['EMI Tenure']);
  document.getElementById('valEMIStartDate').textContent = getVal(student, ['EMI Start Date']);
  document.getElementById('valSubventionAmt').textContent = getVal(student, ['Subvention Amount']);
  document.getElementById('valClosedStatus').textContent = closedStatusText;
  document.getElementById('valUpsellStatus').textContent = getVal(student, ['Upsell']);

  // Section 5: Cohort Timeline & Admin Details
  document.getElementById('valRetentionWeek').textContent = getVal(student, ['Retention Week']);
  document.getElementById('valPrateekWeek').textContent = getVal(student, ['Prateek sir Week']);
  document.getElementById('valDeadline').textContent = getVal(student, ['Deadline']);
  document.getElementById('valSOPDate').textContent = getVal(student, ['Weekly date/SOP date']);
  document.getElementById('valReferralAmount').textContent = getVal(student, ['Referral amount']);
  document.getElementById('valMasaiBitSom').textContent = getVal(student, ['BITSoM/Masai']);
  document.getElementById('valBucket').textContent = getVal(student, ['Bucket']);

  // Set avatar initials
  const name = getVal(student, ['Student Name'], 'N A');
  const initials = name.split(' ').map(n => n[0]).slice(0,2).join('').toUpperCase();
  document.getElementById('studentAvatar').textContent = initials;

  // View raw JSON view (excluding helper underscore properties)
  const cleanJSON = {};
  Object.keys(student).forEach(k => {
    if (!k.startsWith('_')) {
      // Re-map to readable names from headers
      const origHeader = headers.find(h => normalizeHeaderKey(h) === k) || k;
      cleanJSON[origHeader] = student[k];
    }
  });
  
  codeRawJSON.textContent = JSON.stringify(cleanJSON, null, 2);

  // Reveal results
  dashboardResults.classList.remove('hidden');
  
  // Smooth scroll to results
  dashboardResults.scrollIntoView({ behavior: 'smooth' });
}

// Session based recent search history
function addToHistory(name, email) {
  // Remove if exists to bubble up
  recentLookups = recentLookups.filter(item => item.email !== email);
  recentLookups.unshift({ name, email });
  
  // Limit to 6 items
  if (recentLookups.length > 6) {
    recentLookups.pop();
  }
  
  updateHistoryUI();
}

function updateHistoryUI() {
  if (recentLookups.length === 0) {
    historyList.innerHTML = '<div class="no-history">No recent searches in this session</div>';
    return;
  }
  
  historyList.innerHTML = '';
  recentLookups.forEach(item => {
    const chip = document.createElement('button');
    chip.className = 'history-chip';
    chip.innerHTML = `
      <span><strong>${item.name}</strong> (${item.email})</span>
    `;
    chip.addEventListener('click', () => {
      txtSearchEmail.value = item.email;
      searchStudent(item.email);
    });
    historyList.appendChild(chip);
  });
}
