// LiveAccess - Google Sheets Live Quick Access Search Engine
// Handles GIS OAuth2, dynamic two-step search, caching, history and rendering.

// State variables
let tokenClient;
let accessToken = null;
let sheetHeaders = null;
let emailIndex = null;
let mobileIndex = null;
let nameIndex = null;
let recentLookups = [];
let isPrivacyMode = false;
let currentStudent = null;
let currentStudentRawRow = null;

// IndexedDB Caching Utility Functions
const DB_NAME = 'LiveAccessCache';
const DB_VERSION = 1;
const STORE_NAME = 'cache';

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = (e) => resolve(e.target.result);
    request.onerror = (e) => reject(e.target.error);
  });
}

async function getCacheItem(key) {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(key);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error('IndexedDB get error:', err);
    return null;
  }
}

async function setCacheItem(key, value) {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.put(value, key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error('IndexedDB put error:', err);
  }
}

async function clearDBCache() {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.clear();
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error('IndexedDB clear error:', err);
  }
}


// DOM Elements
const btnPrivacyToggle = document.getElementById('btnPrivacyToggle');
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
const autocompleteSuggestions = document.getElementById('autocompleteSuggestions');

// Autocomplete State
let activeSuggestionIndex = -1;
let currentSuggestions = [];

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
  btnPrivacyToggle.addEventListener('click', togglePrivacyMode);
  btnSettings.addEventListener('click', () => showSettingsModal(true));
  btnCloseSettings.addEventListener('click', () => showSettingsModal(false));
  btnStartSetup.addEventListener('click', () => showSettingsModal(true));
  
  btnSaveSettings.addEventListener('click', saveConfig);
  btnUseDefaults.addEventListener('click', loadSampleConfig);
  
  btnConnect.addEventListener('click', connectGoogleAccount);
  btnSignOut.addEventListener('click', signOutGoogle);
  
  btnSearch.addEventListener('click', triggerSearch);
  
  // Event listeners for autocomplete search suggestions
  txtSearchEmail.addEventListener('input', handleSearchInput);
  txtSearchEmail.addEventListener('keydown', handleSearchKeydown);
  document.addEventListener('click', handleOutsideClick);
  
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

function togglePrivacyMode() {
  isPrivacyMode = !isPrivacyMode;
  
  const txtSpan = btnPrivacyToggle.querySelector('span');
  const eyeOpen = btnPrivacyToggle.querySelector('.icon-eye-open');
  const eyeClosed = btnPrivacyToggle.querySelector('.icon-eye-closed');
  
  if (isPrivacyMode) {
    txtSpan.textContent = 'Privacy On';
    btnPrivacyToggle.classList.add('btn-primary');
    btnPrivacyToggle.classList.remove('btn-secondary');
    eyeOpen.classList.add('hidden');
    eyeClosed.classList.remove('hidden');
  } else {
    txtSpan.textContent = 'Privacy Off';
    btnPrivacyToggle.classList.add('btn-secondary');
    btnPrivacyToggle.classList.remove('btn-primary');
    eyeOpen.classList.remove('hidden');
    eyeClosed.classList.add('hidden');
  }
  
  if (currentStudent && sheetHeaders && currentStudentRawRow) {
    displayStudentDetails(currentStudent, sheetHeaders, currentStudentRawRow);
  }
  updateHistoryUI();
}

function maskEmail(email) {
  if (!isPrivacyMode || !email || !email.includes('@')) return email;
  const [local, domain] = email.split('@');
  if (local.length <= 2) return `${local[0]}*@${domain}`;
  return `${local[0]}${'*'.repeat(local.length - 2)}${local[local.length - 1]}@${domain}`;
}

function maskPhone(phone) {
  if (!isPrivacyMode || !phone) return phone;
  const clean = phone.trim();
  if (clean.length <= 4) return '******';
  const last4 = clean.slice(-4);
  const leading = clean.slice(0, -4);
  if (leading.startsWith('+')) {
    const countryCode = leading.slice(0, 3);
    const middle = leading.slice(3);
    return `${countryCode}${'*'.repeat(Math.max(1, middle.length))}${last4}`;
  } else {
    return `${'*'.repeat(Math.max(1, leading.length))}${last4}`;
  }
}

function formatINR(val) {
  if (!val || val === '-') return '-';
  const clean = val.replace(/[^0-9.]/g, '');
  if (!clean || isNaN(parseFloat(clean))) return val;
  const num = parseFloat(clean);
  const formatter = new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  });
  return formatter.format(num);
}

function displayVal(val, type = 'text') {
  if (!val || val === '-') return '-';
  if (isPrivacyMode) {
    if (type === 'email') return maskEmail(val);
    if (type === 'phone') return maskPhone(val);
    if (type === 'currency') return '₹ *,**,***';
    if (type === 'utr' || type === 'appid') return '********';
    return val;
  }
  if (type === 'currency') {
    return formatINR(val);
  }
  return val;
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
  if (!accessToken) return null;

  const rawSpreadsheetId = localStorage.getItem('cfg_spreadsheet_id');
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  const sheetName = localStorage.getItem('cfg_sheet_name') || 'Master Data';
  const searchColumn = localStorage.getItem('cfg_search_column') || 'Email id';

  if (!spreadsheetId) {
    updateIndexStatus('Spreadsheet not configured.', 'red');
    return null;
  }

  // 1. Try loading from memory first
  if (!forceRefresh && sheetHeaders && emailIndex) {
    return { headers: sheetHeaders, emails: emailIndex, names: nameIndex, mobiles: mobileIndex };
  }

  // 2. Try loading from IndexedDB cache if not force refreshing
  if (!forceRefresh) {
    try {
      updateIndexStatus('Loading local cache...', 'yellow');
      const cached = await getCacheItem('sheet_index_cache');
      if (cached && 
          cached.spreadsheetId === spreadsheetId && 
          cached.sheetName === sheetName && 
          cached.searchColumn === searchColumn &&
          cached.headers && cached.emails) {
        
        sheetHeaders = cached.headers;
        emailIndex = cached.emails;
        nameIndex = cached.names || [];
        mobileIndex = cached.mobiles || [];
        
        updateIndexStatus(`Offline Cache: ${emailIndex.length.toLocaleString()} records.`, 'green');
        return { headers: sheetHeaders, emails: emailIndex, names: nameIndex, mobiles: mobileIndex };
      }
    } catch (err) {
      console.warn('Failed to retrieve index from IndexedDB cache:', err);
    }
  }

  try {
    updateIndexStatus('Loading metadata...', 'yellow');
    showLoader(true, 'Reading Header structure...', 'Fetching column headers to locate database indices...');

    // 1. Fetch headers row A1:ZZ1
    const headersData = await callSheetsAPI(spreadsheetId, `'${sheetName}'!A1:ZZ1`);
    if (!headersData.values || headersData.values.length === 0) {
      throw new Error('Spreadsheet exists but returned no data cells in the first row.');
    }

    sheetHeaders = headersData.values[0];

    // Find the columns
    const normSearchColumn = normalizeHeaderKey(searchColumn);
    const emailColIndex = sheetHeaders.findIndex(h => normalizeHeaderKey(h) === normSearchColumn);

    if (emailColIndex === -1) {
      throw new Error(`Search Column "${searchColumn}" not found in headers. Please check spelling in Connection Settings.`);
    }

    const normNameColumn = normalizeHeaderKey("Student Name");
    const nameColIndex = sheetHeaders.findIndex(h => normalizeHeaderKey(h) === normNameColumn);

    const normMobileColumn = normalizeHeaderKey("Mobile");
    const mobileColIndex = sheetHeaders.findIndex(h => normalizeHeaderKey(h) === normMobileColumn);

    // 2. Fetch the indices in parallel
    showLoader(true, 'Downloading Database Indices...', 'Downloading columns for Name, Email, and Mobile fields to optimize queries...');

    const promises = [];
    const keys = [];

    // Email index (configured search column)
    const emailColLetter = getColumnLetter(emailColIndex);
    promises.push(callSheetsAPI(spreadsheetId, `'${sheetName}'!${emailColLetter}2:${emailColLetter}`));
    keys.push('email');

    // Name index
    if (nameColIndex !== -1) {
      const nameColLetter = getColumnLetter(nameColIndex);
      promises.push(callSheetsAPI(spreadsheetId, `'${sheetName}'!${nameColLetter}2:${nameColLetter}`));
      keys.push('name');
    }

    // Mobile index
    if (mobileColIndex !== -1) {
      const mobileColLetter = getColumnLetter(mobileColIndex);
      promises.push(callSheetsAPI(spreadsheetId, `'${sheetName}'!${mobileColLetter}2:${mobileColLetter}`));
      keys.push('mobile');
    }

    const results = await Promise.all(promises);

    results.forEach((resData, idx) => {
      const key = keys[idx];
      const rawValues = resData.values ? resData.values.map(row => (row[0] || '').trim()) : [];
      
      if (key === 'email') {
        emailIndex = rawValues.map(v => v.toLowerCase());
      } else if (key === 'name') {
        nameIndex = rawValues;
      } else if (key === 'mobile') {
        // Clean mobile number - preserve digits and plus sign
        mobileIndex = rawValues.map(v => v.replace(/[^0-9+]/g, ''));
      }
    });

    // Handle missing indices gracefully
    if (!nameIndex) nameIndex = new Array(emailIndex.length).fill('');
    if (!mobileIndex) mobileIndex = new Array(emailIndex.length).fill('');

    // Save back to IndexedDB
    try {
      const cacheObj = {
        spreadsheetId,
        sheetName,
        searchColumn,
        headers: sheetHeaders,
        emails: emailIndex,
        names: nameIndex,
        mobiles: mobileIndex,
        timestamp: Date.now()
      };
      await setCacheItem('sheet_index_cache', cacheObj);
    } catch (err) {
      console.warn('Failed to save index to IndexedDB:', err);
    }

    showLoader(false);
    updateIndexStatus(`Index cached: ${emailIndex.length.toLocaleString()} records.`, 'green');
    
    return { headers: sheetHeaders, emails: emailIndex, names: nameIndex, mobiles: mobileIndex };
  } catch (err) {
    showLoader(false);
    updateIndexStatus('Indexing failed.', 'red');
    console.error('Indexing error:', err);
    
    let adviceText = '';
    if (err.message === 'Failed to fetch' || err.message.includes('Failed to fetch')) {
      adviceText = '\n\nTroubleshooting tips:\n1. Verify your Spreadsheet ID (ensure it is just the ID, not the full URL).\n2. Disable any adblockers or privacy extensions (like Brave Shields) for this page, as they may block Google API requests.\n3. Make sure you are connected to the internet.';
    }
    
    alert(`Spreadsheet Indexing failed: ${err.message}${adviceText}\n\nPlease verify Spreadsheet ID, Sheet Name, and that your Google Account has permissions.`);
    return null;
  }
}

// Autocomplete Suggestions Handlers
function handleSearchInput(e) {
  const query = e.target.value.trim().toLowerCase();
  
  if (!emailIndex || query.length < 2) {
    hideSuggestions();
    return;
  }

  // Filter indices for matches
  const suggestions = [];
  const maxSuggestions = 10;

  for (let i = 0; i < emailIndex.length; i++) {
    const email = emailIndex[i] || '';
    const name = nameIndex[i] || '';
    const mobile = mobileIndex[i] || '';

    const emailMatch = email.includes(query);
    const nameMatch = name.toLowerCase().includes(query);
    const mobileMatch = mobile.includes(query);

    if (emailMatch || nameMatch || mobileMatch) {
      suggestions.push({
        name: name || 'Student Record',
        email: email,
        mobile: mobile,
        index: i
      });
      if (suggestions.length >= maxSuggestions) {
        break;
      }
    }
  }

  currentSuggestions = suggestions;
  renderSuggestions(query);
}

function renderSuggestions(query) {
  if (currentSuggestions.length === 0) {
    hideSuggestions();
    return;
  }

  autocompleteSuggestions.innerHTML = '';
  activeSuggestionIndex = -1;

  currentSuggestions.forEach((item, idx) => {
    const suggestionDiv = document.createElement('div');
    suggestionDiv.className = 'autocomplete-suggestion';
    
    // Mask the email/mobile if privacy mode is enabled
    const maskedEmail = displayVal(item.email, 'email');
    const maskedMobile = item.mobile ? displayVal(item.mobile, 'phone') : '';

    // Highlight matched segments
    const displayName = highlightMatch(item.name, query);
    const displayEmail = highlightMatch(maskedEmail, query);
    const displayMobile = maskedMobile ? highlightMatch(maskedMobile, query) : '';

    let subText = displayEmail;
    if (displayMobile) {
      subText += ` | Mobile: ${displayMobile}`;
    }

    suggestionDiv.innerHTML = `
      <span class="suggestion-name">${displayName}</span>
      <span class="suggestion-sub">${subText}</span>
    `;

    // Click behavior
    suggestionDiv.addEventListener('click', () => {
      // Set to email as the query value, load profile
      txtSearchEmail.value = item.email;
      hideSuggestions();
      searchStudent(item.email);
    });

    autocompleteSuggestions.appendChild(suggestionDiv);
  });

  autocompleteSuggestions.classList.remove('hidden');
}

function highlightMatch(text, query) {
  if (!text) return '';
  const index = text.toLowerCase().indexOf(query);
  if (index === -1) return text;
  
  const before = text.substring(0, index);
  const match = text.substring(index, index + query.length);
  const after = text.substring(index + query.length);
  
  return `${before}<mark>${match}</mark>${after}`;
}

function handleSearchKeydown(e) {
  if (autocompleteSuggestions.classList.contains('hidden')) {
    if (e.key === 'Enter') {
      triggerSearch();
    }
    return;
  }

  const items = autocompleteSuggestions.querySelectorAll('.autocomplete-suggestion');

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    activeSuggestionIndex = (activeSuggestionIndex + 1) % items.length;
    updateActiveSuggestion(items);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    activeSuggestionIndex = (activeSuggestionIndex - 1 + items.length) % items.length;
    updateActiveSuggestion(items);
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (activeSuggestionIndex >= 0 && activeSuggestionIndex < currentSuggestions.length) {
      const selected = currentSuggestions[activeSuggestionIndex];
      txtSearchEmail.value = selected.email;
      hideSuggestions();
      searchStudent(selected.email);
    } else {
      triggerSearch();
    }
  } else if (e.key === 'Escape') {
    hideSuggestions();
  }
}

function updateActiveSuggestion(items) {
  items.forEach((item, idx) => {
    if (idx === activeSuggestionIndex) {
      item.classList.add('active');
      // Scroll suggestion into view if needed
      item.scrollIntoView({ block: 'nearest' });
    } else {
      item.classList.remove('active');
    }
  });
}

function hideSuggestions() {
  autocompleteSuggestions.classList.add('hidden');
  autocompleteSuggestions.innerHTML = '';
  activeSuggestionIndex = -1;
  currentSuggestions = [];
}

function handleOutsideClick(e) {
  if (!txtSearchEmail.contains(e.target) && !autocompleteSuggestions.contains(e.target)) {
    hideSuggestions();
  }
}

// Trigger query for student (supports Name, Email, or Mobile)
function triggerSearch() {
  const searchInput = txtSearchEmail.value.trim();
  if (!searchInput) {
    alert('Please enter a student name, email, or mobile number to search.');
    return;
  }
  
  searchStudent(searchInput);
}

// Search Student and display row (by Email, Mobile, or Name)
async function searchStudent(searchQuery) {
  if (!accessToken) {
    alert('Please connect your Google account first.');
    return;
  }

  const query = searchQuery.trim();

  try {
    // Ensure index is loaded
    const index = await fetchSpreadsheetIndex();
    if (!index) return;

    const { headers, emails, names, mobiles } = index;

    showLoader(true, 'Searching Database...', `Locating record matching "${query}"...`);

    let matchIndex = -1;
    let queryType = '';

    // 1. Detect query type and find matching row index
    if (query.includes('@')) {
      // Search by Email
      queryType = 'email';
      const cleanEmail = query.toLowerCase();
      matchIndex = emails.indexOf(cleanEmail);
    } else if (/^[0-9+\-\s()]+$/.test(query) && query.replace(/[^0-9]/g, '').length >= 5) {
      // Search by Mobile (if it looks like a phone number)
      queryType = 'mobile';
      const cleanMobile = query.replace(/[^0-9+]/g, '');
      
      // Try exact match first
      matchIndex = mobiles.indexOf(cleanMobile);
      if (matchIndex === -1) {
        // Try partial match
        matchIndex = mobiles.findIndex(m => m && m.includes(cleanMobile));
      }
    } else {
      // Search by Name
      queryType = 'name';
      
      // Try exact case-insensitive match first
      matchIndex = names.findIndex(n => n && n.toLowerCase() === query.toLowerCase());
      if (matchIndex === -1) {
        // Try partial match
        matchIndex = names.findIndex(n => n && n.toLowerCase().includes(query.toLowerCase()));
      }
    }

    if (matchIndex === -1) {
      showLoader(false);
      let alertMsg = `Student with query "${query}" not found.`;
      if (queryType === 'email') {
        alertMsg = `Student with email ID "${query}" not found in the live sheet index.`;
      } else if (queryType === 'mobile') {
        alertMsg = `Student with mobile number "${query}" not found in the live sheet index.`;
      } else if (queryType === 'name') {
        alertMsg = `Student named "${query}" not found in the live sheet index.`;
      }
      alert(`${alertMsg}\n\nDouble check details or click "Refresh Index" if the record was recently added.`);
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
    student._email = emails[matchIndex] || getVal(student, ['Email id', 'Email'], 'email@domain.com');
    student._row = rowNum;

    // Display student data
    displayStudentDetails(student, headers, rowData);
    
    // Add to lookups list
    addToHistory(student._name, student._email);
    
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
  // Save current student and row data for re-rendering
  currentStudent = student;
  currentStudentRawRow = rowData;

  // Update banner profile card
  document.getElementById('lblStudentName').textContent = displayVal(getVal(student, ['Student Name']));
  document.getElementById('badgeCohort').textContent = displayVal(getVal(student, ['Cohort ID', 'Cohort']));
  document.getElementById('badgeCategory').textContent = displayVal(getVal(student, ['Category', 'Test Category']));
  
  document.getElementById('lblStudentEmail').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M21.75 6.75v10.5a2.25 2.25 0 01-2.25 2.25h-15a2.25 2.25 0 01-2.25-2.25V6.75m19.5 0A2.25 2.25 0 0019.5 4.5h-15a2.25 2.25 0 00-2.25 2.25m19.5 0v.243a2.25 2.25 0 01-1.07 1.916l-7.5 4.615a2.25 2.25 0 01-2.36 0L3.32 8.91a2.25 2.25 0 01-1.07-1.916V6.75"/></svg>
    ${displayVal(getVal(student, ['Email id', 'Email']), 'email')}`;
    
  document.getElementById('lblStudentPhone').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M2.25 6.622M2.25 9c0 5.108 4.027 9.24 9 9.24 2.894 0 5.5-.327 7.747-.94a.75.75 0 00.563-.647c.187-1.393.308-2.8-.073-4.148a.75.75 0 00-.702-.513H15.02a.75.75 0 00-.69.44l-.79 1.58A13.23 13.23 0 017.5 9.79l1.58-.79a.75.75 0 00.44-.69v-3.72a.75.75 0 00-.513-.702A24.819 24.819 0 004.1 3.81a.75.75 0 00-.647.563C2.827 6.6 2.25 9 2.25 9z" stroke-linecap="round" stroke-linejoin="round"/></svg>
    ${displayVal(getVal(student, ['Mobile']), 'phone')}`;
    
  document.getElementById('lblStudentAltPhone').innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16"><path d="M2.25 6.75c0 8.284 6.716 15 15 15h2.25a2.25 2.25 0 002.25-2.25v-1.372c0-.516-.351-.966-.852-1.091l-4.423-1.106c-.44-.11-.902.055-1.173.417l-.97 1.293c-2.824-1.802-5.14-4.118-6.942-6.942l1.293-.97c.362-.271.527-.834.417-1.173L6.963 3.102a1.125 1.125 0 00-1.091-.852H4.5A2.25 2.25 0 002.25 4.5v2.25z" stroke-linecap="round" stroke-linejoin="round"/></svg>
    Alt: ${displayVal(getVal(student, ['Alternate Number']), 'phone')}`;
 
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
  document.getElementById('valCourseName').textContent = displayVal(getVal(student, ['Course Name']));
  document.getElementById('valCommonName').textContent = displayVal(getVal(student, ['Common Name']));
  document.getElementById('valBatchStart').textContent = displayVal(getVal(student, ['Batch Start Date']));
  document.getElementById('valFeeNoGST').textContent = displayVal(getVal(student, ['Course Fees(Exclusive of GST)', 'Course Fees (Exclusive of GST)']), 'currency');
  document.getElementById('valFeeWithGST').textContent = displayVal(getVal(student, ['Course Fees ( Inclusive of GST)', 'Course Fees (Inclusive of GST)']), 'currency');
  document.getElementById('valGSTStatus').textContent = displayVal(getVal(student, ['GST Yes/NO', 'GST YesNo']));
  document.getElementById('valGSTTaken').textContent = displayVal(getVal(student, ['GST Taken']), 'currency');
 
  // Section 2: Scholarship & Net Fees
  document.getElementById('valScholarshipCat').textContent = displayVal(getVal(student, ['Scholarship Category']));
  document.getElementById('valScholarshipPct').textContent = displayVal(getVal(student, ['Scholarship %']));
  document.getElementById('valScholarshipAmt').textContent = displayVal(getVal(student, ['Scholarship Amount on (Column J)', 'Scholarship Amount on\n(Column J)', 'Scholarship Amount']), 'currency');
  document.getElementById('valEffectiveFee').textContent = displayVal(getVal(student, ['Effective fee Column (J-P-M-AL) Without GST', 'Effective fee\nColumn (J-P-M-AL) Without GST', 'Effective Fee']), 'currency');
  document.getElementById('valActualFullFee').textContent = displayVal(getVal(student, ['Actual Full Fee (W/O GST) - Vineet', 'Actual Full Fee']), 'currency');
  document.getElementById('valActualRegFee').textContent = displayVal(getVal(student, ['Actual Registration Fee']), 'currency');
  document.getElementById('valActualRemFee').textContent = displayVal(getVal(student, ['Actual Remaining Fee']), 'currency');
 
  // Section 3: Test & Registration Fees
  document.getElementById('valTestFeePaid').textContent = displayVal(getVal(student, ['Test Fee Paid', 'Test Fee Paid ']), 'currency');
  document.getElementById('valTestFeeDate').textContent = displayVal(getVal(student, ['Test fee Date']));
  document.getElementById('valTestFeeUTR').textContent = displayVal(getVal(student, ['Test Fee UTR']), 'utr');
  document.getElementById('valRegFeePaid').textContent = displayVal(getVal(student, ['Reg. fee paid']), 'currency');
  document.getElementById('valRegFeeDate').textContent = displayVal(getVal(student, ['Date of Reg. Fee', 'Date of Reg Fee']));
  document.getElementById('valRegFeeUTR').textContent = displayVal(getVal(student, ['UTR for Registration Fees']), 'utr');
  document.getElementById('valStudentCat').textContent = displayVal(getVal(student, ['Student Category', 'Category']));
 
  // Section 4: Loan & NBFC Status
  document.getElementById('valPaymentMode').textContent = displayVal(getVal(student, ['Payment Mode']));
  document.getElementById('valNBFCName').textContent = displayVal(getVal(student, ['NBFC NAME', 'NBFC']));
  document.getElementById('valLenderName').textContent = displayVal(getVal(student, ['Propelld - Lender Name', 'Propelld Lender Name']));
  document.getElementById('valNBFCAppID').textContent = displayVal(getVal(student, ['App ID', 'Application ID']), 'appid');
  applyStatusBadge(document.getElementById('valNBFCStatus'), nbfcStatusText);
  document.getElementById('valForeclosureStatus').textContent = displayVal(getVal(student, ['Foreclosure/Refund Status']));
  
  document.getElementById('valLoanAmount').textContent = displayVal(getVal(student, ['Loan Amount']), 'currency');
  document.getElementById('valDisbursedAmount').textContent = displayVal(getVal(student, ['Disbursed Amount/Direct Payment', 'Disbursed Amount']), 'currency');
  document.getElementById('valDisbursedDate').textContent = displayVal(getVal(student, ['Disbursed/Paid Date', 'Disbursed Date']));
  document.getElementById('valDisbursalUTR').textContent = displayVal(getVal(student, ['UTR']), 'utr');
  document.getElementById('valOverallAmountPaid').textContent = displayVal(getVal(student, ['Over all Amount Paid']), 'currency');
  document.getElementById('valBajajAdjusted').textContent = displayVal(getVal(student, ['Bajaj Adjusted Amount']), 'currency');
  
  document.getElementById('valEMIAmount').textContent = displayVal(getVal(student, ['EMI Amount']), 'currency');
  document.getElementById('valEMITenure').textContent = displayVal(getVal(student, ['EMI Tenure']));
  document.getElementById('valEMIStartDate').textContent = displayVal(getVal(student, ['EMI Start Date']));
  document.getElementById('valSubventionAmt').textContent = displayVal(getVal(student, ['Subvention Amount']), 'currency');
  document.getElementById('valClosedStatus').textContent = displayVal(closedStatusText);
  document.getElementById('valUpsellStatus').textContent = displayVal(getVal(student, ['Upsell']));
 
  // Section 5: Cohort Timeline & Admin Details
  document.getElementById('valRetentionWeek').textContent = displayVal(getVal(student, ['Retention Week']));
  document.getElementById('valPrateekWeek').textContent = displayVal(getVal(student, ['Prateek sir Week']));
  document.getElementById('valDeadline').textContent = displayVal(getVal(student, ['Deadline']));
  document.getElementById('valSOPDate').textContent = displayVal(getVal(student, ['Weekly date/SOP date']));
  document.getElementById('valReferralAmount').textContent = displayVal(getVal(student, ['Referral amount']), 'currency');
  document.getElementById('valMasaiBitSom').textContent = displayVal(getVal(student, ['BITSoM/Masai']));
  document.getElementById('valBucket').textContent = displayVal(getVal(student, ['Bucket']));
 
  // Set avatar initials
  const name = getVal(student, ['Student Name'], 'N A');
  const initials = name.split(' ').map(n => n[0]).slice(0,2).join('').toUpperCase();
  document.getElementById('studentAvatar').textContent = initials;
 
  // View raw JSON view (excluding helper underscore properties) or mask it
  if (isPrivacyMode) {
    codeRawJSON.textContent = "{\n  \"message\": \"Raw JSON details are hidden when Privacy Mode is enabled.\"\n}";
  } else {
    const cleanJSON = {};
    Object.keys(student).forEach(k => {
      if (!k.startsWith('_')) {
        const origHeader = headers.find(h => normalizeHeaderKey(h) === k) || k;
        cleanJSON[origHeader] = student[k];
      }
    });
    codeRawJSON.textContent = JSON.stringify(cleanJSON, null, 2);
  }
 
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
      <span><strong>${item.name}</strong> (${displayVal(item.email, 'email')})</span>
    `;
    chip.addEventListener('click', () => {
      txtSearchEmail.value = item.email;
      searchStudent(item.email);
    });
    historyList.appendChild(chip);
  });
}
