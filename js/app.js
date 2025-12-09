/**
 * Overtime Summary - Manager Dashboard
 * Features:
 * - Per-user capacity and multiplier overrides (localStorage persistence)
 * - Cost calculation using entry's inherent hourlyRate with Base/OT Premium breakdown
 * - Native Clockify "Grouped Table" style UI with parent/child rows
 */

// ============================================================================
// Global State
// ============================================================================

const state = {
  token: null,
  claims: null,
  isLoading: false,
  currentData: null,
  rawEntries: null,
  users: [],                    // Fetched workspace users
  userOverrides: {},            // { [userId]: { multiplier, capacity } } - persisted in localStorage
  overridesCollapsed: true,
  configCollapsed: true,
  activeTab: 'summary',         // 'summary' or 'detailed'
  summaryGroupBy: 'user'        // 'user' | 'client' | 'date' | 'week'
};

// Density preference (persisted)
const density = {
  mode: 'compact',
  selectEl: null,
  setMode(mode) {
    const allowed = ['compact', 'comfortable', 'spacious'];
    if (!allowed.includes(mode)) {
      mode = 'compact';
    }
    this.mode = mode;
    document.body.dataset.density = mode;
    document.body.classList.remove('density-compact', 'density-comfortable', 'density-spacious');
    document.body.classList.add(`density-${mode}`);

    if (this.selectEl && this.selectEl.value !== mode) {
      this.selectEl.value = mode;
    }

    try {
      localStorage.setItem('overtime_density', mode);
    } catch (err) {
      console.warn('Unable to persist density preference', err);
    }
  },
  init(selectEl) {
    this.selectEl = selectEl || this.selectEl;
    let saved = 'compact';
    try {
      saved = localStorage.getItem('overtime_density') || 'compact';
    } catch (err) {
      console.warn('Unable to read density preference', err);
    }
    this.setMode(saved);
  }
};

// ============================================================================
// DOM Elements (lazy-loaded after DOMContentLoaded)
// ============================================================================

let elements = {};

function initElements() {
  console.log('Initializing DOM elements v2...');
  elements = {
    // Density + quick filters
    densitySelect: document.getElementById('densitySelect'),
    quickMonth: document.getElementById('quickMonth'),
    quickWeek: document.getElementById('quickWeek'),
    quickToday: document.getElementById('quickToday'),
    // Config inputs
    configDaily: document.getElementById('configDaily'),
    configWeekly: document.getElementById('configWeekly'),
    configMultiplier: document.getElementById('configMultiplier'),
    configToggle: document.getElementById('configToggle'),
    configContent: document.getElementById('configContent'),
    configApplyBtn: document.getElementById('configApplyBtn'),
    configSummary: document.getElementById('configSummary'),
    // Date pickers
    startDate: document.getElementById('startDate'),
    endDate: document.getElementById('endDate'),
    generateBtn: document.getElementById('generateBtn'),
    // State containers
    loadingState: document.getElementById('loadingState'),
    errorState: document.getElementById('errorState'),
    errorMessage: document.getElementById('errorMessage'),
    resultsContainer: document.getElementById('resultsContainer'),
    emptyState: document.getElementById('emptyState'),
    // Summary strip elements
    summaryUsers: document.getElementById('summaryUsers'),
    summaryTotalHours: document.getElementById('summaryTotalHours'),
    summaryRegularHours: document.getElementById('summaryRegularHours'),
    summaryOvertimeHours: document.getElementById('summaryOvertimeHours'),
    summaryTotalCost: document.getElementById('summaryTotalCost'),
    summaryOtPremium: document.getElementById('summaryOtPremium'),
    userCountLabel: document.getElementById('userCountLabel'),
    // User overrides
    userOverridesCard: document.getElementById('userOverridesCard'),
    userOverridesBody: document.getElementById('userOverridesBody'),
    userOverridesContent: document.getElementById('userOverridesContent'),
    overridesToggle: document.getElementById('overridesToggle'),
    resetOverridesBtn: document.getElementById('resetOverridesBtn'),
    overrideSearch: document.getElementById('overrideSearch'),
    // Report table
    reportTableBody: document.getElementById('reportTableBody'),
    // Summary table
    summaryCard: document.getElementById('summaryCard'),
    summaryTableBody: document.getElementById('summaryTableBody'),
    summaryUserCount: document.getElementById('summaryUserCount'),
    reportCard: document.getElementById('reportCard'),
    otChart: document.getElementById('otChart'),
    otChartContainer: document.getElementById('otChartContainer'),
    // Group by dropdown
    groupByBtn: document.getElementById('groupByBtn'),
    groupByMenu: document.getElementById('groupByMenu'),
    groupByLabel: document.getElementById('groupByLabel'),
    groupDropdown: document.querySelector('.group-dropdown'),
    // Tab navigation
    tabBtns: document.querySelectorAll('.tab-btn'),
    // Export button
    exportBtn: document.getElementById('exportBtn')
  };

  // Debug: Log any null elements
  const nullElements = Object.entries(elements).filter(([k, v]) => v === null).map(([k]) => k);
  if (nullElements.length > 0) {
    console.error('Missing DOM elements:', nullElements);
  } else {
    console.log('All DOM elements found successfully');
  }
}

// ============================================================================
// Configuration (read from UI inputs)
// ============================================================================

function getConfig() {
  return {
    dailyThreshold: parseFloat(elements.configDaily.value) || 8,
    weeklyThreshold: parseFloat(elements.configWeekly.value) || 40,
    overtimeMultiplier: parseFloat(elements.configMultiplier.value) || 1.5
  };
}

function updateConfigSummary() {
  const daily = elements.configDaily.value || '8';
  const weekly = elements.configWeekly.value || '40';
  const mult = elements.configMultiplier.value || '1.5';
  if (elements.configSummary) {
    elements.configSummary.textContent = `(${daily}h / ${weekly}h / ${mult}x)`;
  }
}

// ============================================================================
// localStorage Persistence for User Overrides
// ============================================================================

function getStorageKey() {
  if (!state.claims || !state.claims.workspaceId) return null;
  return `overtime_overrides_${state.claims.workspaceId}`;
}

function loadUserOverrides() {
  const key = getStorageKey();
  if (!key) return {};

  try {
    const saved = localStorage.getItem(key);
    if (saved) {
      return JSON.parse(saved);
    }

    // Migration: check for old userMultipliers format
    const oldKey = `overtime_multipliers_${state.claims.workspaceId}`;
    const oldSaved = localStorage.getItem(oldKey);
    if (oldSaved) {
      const oldData = JSON.parse(oldSaved);
      // Migrate to new format
      const migrated = {};
      for (const [userId, multiplier] of Object.entries(oldData)) {
        migrated[userId] = { multiplier };
      }
      // Save in new format and remove old
      localStorage.setItem(key, JSON.stringify(migrated));
      localStorage.removeItem(oldKey);
      return migrated;
    }

    return {};
  } catch (err) {
    console.warn('Failed to load user overrides:', err);
    return {};
  }
}

function saveUserOverrides() {
  const key = getStorageKey();
  if (!key) return;

  try {
    localStorage.setItem(key, JSON.stringify(state.userOverrides));
  } catch (err) {
    console.warn('Failed to save user overrides:', err);
  }
}

function resetUserOverrides() {
  state.userOverrides = {};
  saveUserOverrides();
  renderUserOverridesTable();
  recalculateAndRender();
}

// ============================================================================
// Initialization
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
  initElements();  // Initialize DOM references first
  initializeApp();
  updateConfigSummary();  // Show initial config values
});

function initializeApp() {
  // Extract token from URL
  extractToken();

  // Apply theme
  applyTheme();
  density.init(elements.densitySelect);

  // Load saved overrides from localStorage
  state.userOverrides = loadUserOverrides();

  // Set default date range (last 30 days)
  setDefaultDateRange();

  // Initialize event listeners
  initEventListeners();

  // Show empty state initially
  showState('empty');
}

function extractToken() {
  const params = new URLSearchParams(window.location.search);
  state.token = params.get('auth_token');

  if (!state.token) {
    console.warn('No auth_token found in URL');
    showError('Authentication token not found. Please reload the add-on from Clockify.');
    return;
  }

  // Decode token to get claims
  try {
    const parts = state.token.split('.');
    if (parts.length === 3) {
      const payload = atob(parts[1].replace(/-/g, '+').replace(/_/g, '/'));
      state.claims = JSON.parse(payload);
      console.log('Token claims:', state.claims);
    }
  } catch (err) {
    console.error('Failed to decode token:', err);
  }
}

function applyTheme() {
  if (state.claims && state.claims.theme === 'DARK') {
    document.body.classList.add('cl-theme-dark');
  }
}

function setDefaultDateRange() {
  const today = new Date();
  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  elements.startDate.value = formatDateForInput(thirtyDaysAgo);
  elements.endDate.value = formatDateForInput(today);
}

function formatDateForInput(date) {
  return date.toISOString().split('T')[0];
}

function applyQuickRange(preset) {
  const today = new Date();
  let start = new Date(today);
  let end = new Date(today);

  if (preset === 'lastMonth') {
    start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    end = new Date(today.getFullYear(), today.getMonth(), 0);
  } else if (preset === 'thisWeek') {
    const day = today.getDay(); // 0 (Sun) - 6 (Sat)
    const diffToMonday = (day + 6) % 7;
    start.setDate(today.getDate() - diffToMonday);
  } else if (preset === 'today') {
    // start and end already set to today
  } else {
    return;
  }

  elements.startDate.value = formatDateForInput(start);
  elements.endDate.value = formatDateForInput(end);
  handleGenerateReport();
}

// ============================================================================
// Event Listeners
// ============================================================================

function initEventListeners() {
  // Generate button
  elements.generateBtn.addEventListener('click', handleGenerateReport);

  // Density toggle
  if (elements.densitySelect) {
    elements.densitySelect.addEventListener('change', (e) => density.setMode(e.target.value));
  }

  // Quick date actions
  if (elements.quickMonth) elements.quickMonth.addEventListener('click', () => applyQuickRange('lastMonth'));
  if (elements.quickWeek) elements.quickWeek.addEventListener('click', () => applyQuickRange('thisWeek'));
  if (elements.quickToday) elements.quickToday.addEventListener('click', () => applyQuickRange('today'));

  // Export CSV button
  if (elements.exportBtn) {
    elements.exportBtn.addEventListener('click', exportToCSV);
  }

  // Config inputs - recalculate when changed (debounced)
  let recalcTimeout;
  const configInputs = [
    elements.configDaily,
    elements.configWeekly,
    elements.configMultiplier
  ];

  configInputs.forEach(input => {
    input.addEventListener('input', () => {
      clearTimeout(recalcTimeout);
      recalcTimeout = setTimeout(() => {
        if (state.rawEntries) {
          recalculateAndRender();
        }
      }, 300);
    });
  });

  if (elements.configApplyBtn) {
    elements.configApplyBtn.addEventListener('click', () => {
      if (state.rawEntries) {
        recalculateAndRender();
      }
      renderUserOverridesTable();
      updateConfigSummary();
    });
  }

  // Config toggle
  elements.configToggle.addEventListener('click', () => {
    state.configCollapsed = !state.configCollapsed;
    elements.configToggle.classList.toggle('collapsed', state.configCollapsed);
    elements.configContent.classList.toggle('hidden', state.configCollapsed);
  });

  // Overrides toggle
  elements.overridesToggle.addEventListener('click', () => {
    state.overridesCollapsed = !state.overridesCollapsed;
    elements.overridesToggle.classList.toggle('collapsed', state.overridesCollapsed);
    elements.userOverridesContent.classList.toggle('hidden', state.overridesCollapsed);
  });

  // Reset overrides button (double-tap confirmation for sandbox compatibility)
  let resetConfirmTimeout = null;
  let resetPendingConfirm = false;
  const resetDefaultLabel = elements.resetOverridesBtn.textContent || 'Reset All';

  elements.resetOverridesBtn.addEventListener('click', () => {
    if (resetPendingConfirm) {
      // Second click - execute reset
      clearTimeout(resetConfirmTimeout);
      resetPendingConfirm = false;
      elements.resetOverridesBtn.textContent = resetDefaultLabel;
      elements.resetOverridesBtn.classList.remove('btn-danger');
      resetUserOverrides();
    } else {
      // First click - ask for confirmation
      resetPendingConfirm = true;
      elements.resetOverridesBtn.textContent = 'Click again to confirm';
      elements.resetOverridesBtn.classList.add('btn-danger');

      resetConfirmTimeout = setTimeout(() => {
        resetPendingConfirm = false;
        elements.resetOverridesBtn.textContent = resetDefaultLabel;
        elements.resetOverridesBtn.classList.remove('btn-danger');
      }, 3000);
    }
  });

  if (elements.overrideSearch) {
    elements.overrideSearch.addEventListener('input', () => renderUserOverridesTable());
  }

  // Tab navigation
  if (elements.tabBtns) {
    elements.tabBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const tab = btn.dataset.tab;
        switchTab(tab);
      });
    });
  }

  // Group by dropdown
  if (elements.groupByBtn) {
    elements.groupByBtn.addEventListener('click', toggleGroupDropdown);
    elements.groupByMenu.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => changeGroupBy(btn.dataset.group));
    });
    document.addEventListener('click', closeGroupDropdown);
  }
}

// ============================================================================
// Tab Switching
// ============================================================================

function switchTab(tab) {
  state.activeTab = tab;

  // Update tab button states
  elements.tabBtns.forEach(btn => {
    if (btn.dataset.tab === tab) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  // Show/hide appropriate containers
  if (tab === 'summary') {
    elements.summaryCard.classList.remove('hidden');
    elements.reportCard.classList.add('hidden');
  } else {
    elements.summaryCard.classList.add('hidden');
    elements.reportCard.classList.remove('hidden');
  }
}

// ============================================================================
// Group By Dropdown
// ============================================================================

function toggleGroupDropdown(e) {
  e.stopPropagation();
  elements.groupByMenu.classList.toggle('hidden');
  elements.groupDropdown.classList.toggle('open');
}

function closeGroupDropdown(e) {
  if (!e.target.closest('.group-dropdown')) {
    elements.groupByMenu.classList.add('hidden');
    elements.groupDropdown.classList.remove('open');
  }
}

function changeGroupBy(group) {
  console.log('[DEBUG] changeGroupBy called with:', group);
  state.summaryGroupBy = group;
  console.log('[DEBUG] state.summaryGroupBy is now:', state.summaryGroupBy);

  // Update label
  const labels = { user: 'User', client: 'Client', project: 'Project', task: 'Task', date: 'Date', week: 'Week' };
  elements.groupByLabel.textContent = labels[group] || 'User';

  // Update active state
  elements.groupByMenu.querySelectorAll('button').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.group === group);
  });

  // Close dropdown
  elements.groupByMenu.classList.add('hidden');
  elements.groupDropdown.classList.remove('open');

  // Re-render summary table
  if (state.currentData && state.currentData.users) {
    renderSummaryTable(state.currentData.users);
  }
}

function recalculateAndRender() {
  if (!state.rawEntries) return;

  const config = getConfig();
  const analysis = calculateUserOvertime(state.rawEntries, config);
  state.currentData = analysis;
  renderResults(analysis);
}

// ============================================================================
// API Calls - Per-User Time Entries (X-Addon-Token compatible)
// ============================================================================

async function handleGenerateReport() {
  const startDate = elements.startDate.value;
  const endDate = elements.endDate.value;

  // Validate dates
  if (!startDate || !endDate) {
    showError('Please select both start and end dates.');
    return;
  }

  if (new Date(startDate) > new Date(endDate)) {
    showError('Start date must be before end date.');
    return;
  }

  if (!state.claims) {
    showError('Invalid authentication token.');
    return;
  }

  // Show loading state
  showState('loading');

  try {
    // Fetch all users' time entries (works with X-Addon-Token)
    const entries = await fetchAllUsersTimeEntries(startDate, endDate);
    state.rawEntries = entries;

    // Calculate per-user overtime
    const config = getConfig();
    const analysis = calculateUserOvertime(entries, config);

    state.currentData = analysis;
    renderResults(analysis);
    renderUserOverridesTable();
    showState('results');
    elements.exportBtn.disabled = false;
  } catch (err) {
    console.error('Report generation error:', err);
    showError(err.message || 'Failed to generate report. Please try again.');
  }
}

async function fetchAllUsersTimeEntries(startDate, endDate) {
  const { backendUrl, workspaceId } = state.claims;

  if (!backendUrl || !workspaceId) {
    throw new Error('Missing backendUrl or workspaceId in token claims');
  }

  // Step 1: Fetch all users in workspace
  console.log('Fetching workspace users...');
  const usersUrl = `${backendUrl}/v1/workspaces/${workspaceId}/users`;

  const usersResponse = await fetch(usersUrl, {
    headers: {
      'X-Addon-Token': state.token,
      'Content-Type': 'application/json'
    }
  });

  if (!usersResponse.ok) {
    const errorText = await usersResponse.text();
    console.error('Users API error:', usersResponse.status, errorText);
    throw new Error(`Failed to fetch users: ${usersResponse.status}`);
  }

  const users = await usersResponse.json();
  state.users = users; // Store for overrides table
  console.log(`Found ${users.length} users in workspace`);

  // Step 2: Fetch time entries for each user (in parallel batches)
  const allEntries = [];
  const batchSize = 5; // Fetch 5 users at a time to stay under rate limits

  for (let i = 0; i < users.length; i += batchSize) {
    const batch = users.slice(i, i + batchSize);
    console.log(`Fetching entries for users ${i + 1}-${Math.min(i + batchSize, users.length)} of ${users.length}...`);

    const batchResults = await Promise.all(
      batch.map(user => fetchUserTimeEntriesWithMeta(user, startDate, endDate))
    );

    for (const entries of batchResults) {
      allEntries.push(...entries);
    }
  }

  console.log(`Total entries fetched: ${allEntries.length}`);
  return allEntries;
}

async function fetchUserTimeEntriesWithMeta(user, startDate, endDate) {
  const { backendUrl, workspaceId } = state.claims;

  // Format dates for API (ISO 8601)
  const start = new Date(startDate);
  start.setHours(0, 0, 0, 0);
  const end = new Date(endDate);
  end.setHours(23, 59, 59, 999);

  const url = `${backendUrl}/v1/workspaces/${workspaceId}/user/${user.id}/time-entries?start=${encodeURIComponent(start.toISOString())}&end=${encodeURIComponent(end.toISOString())}&page-size=500&hydrated=true`;

  try {
    const response = await fetch(url, {
      headers: {
        'X-Addon-Token': state.token,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      console.warn(`Failed to fetch entries for user ${user.name} (${user.id}): ${response.status}`);
      return [];
    }

    const entries = await response.json();

    // Attach user info to each entry
    return entries.map(entry => ({
      ...entry,
      userId: user.id,
      userName: user.name,
      userEmail: user.email || ''
    }));
  } catch (err) {
    console.warn(`Error fetching entries for user ${user.name}:`, err);
    return [];
  }
}

// ============================================================================
// Per-User Overtime Calculation (Using Entry's Inherent hourlyRate)
// ============================================================================

function calculateUserOvertime(entries, config) {
  if (!entries || entries.length === 0) {
    return emptyResult();
  }

  // Group entries by user
  const userMap = new Map();

  for (const entry of entries) {
    const userId = entry.userId || entry._id || 'unknown';
    const userName = entry.userName || 'Unknown User';
    const userEmail = entry.userEmail || '';

    if (!userMap.has(userId)) {
      userMap.set(userId, {
        userId,
        userName,
        userEmail,
        entries: []
      });
    }

    userMap.get(userId).entries.push(entry);
  }

  // Calculate overtime for each user
  const users = [];

  for (const [userId, userData] of userMap) {
    const userAnalysis = calculateSingleUserOvertime(userData, config);
    users.push(userAnalysis);
  }

  // Sort by overtime hours (descending)
  users.sort((a, b) => b.overtimeHours - a.overtimeHours);

  // Calculate team totals
  const teamSummary = calculateTeamSummary(users);

  return {
    users,
    summary: teamSummary,
    config
  };
}

function calculateSingleUserOvertime(userData, config) {
  const { userId, userName, userEmail, entries } = userData;

  // Get user-specific overrides or fall back to globals
  const userOverride = state.userOverrides[userId] || {};
  const userMultiplier = userOverride.multiplier !== undefined
    ? parseFloat(userOverride.multiplier)
    : config.overtimeMultiplier;
  const userCapacity = userOverride.capacity !== undefined
    ? parseFloat(userOverride.capacity)
    : config.dailyThreshold;

  // Group entries by day
  const byDay = new Map();

  for (const entry of entries) {
    // Extract date from timeInterval
    let dateKey = null;
    if (entry.timeInterval && entry.timeInterval.start) {
      dateKey = entry.timeInterval.start.substring(0, 10);
    }

    if (!dateKey) continue;

    if (!byDay.has(dateKey)) {
      byDay.set(dateKey, []);
    }
    byDay.get(dateKey).push(entry);
  }

  const entryDetails = [];
  let totalHours = 0;
  let overtimeHours = 0;
  let totalBaseCost = 0;
  let totalOtPremium = 0;

  for (const [dateKey, dayEntries] of byDay) {
    // Sort entries by start time within the day to preserve OT ordering
    dayEntries.sort((a, b) => {
      const aStart = a.timeInterval?.start ? new Date(a.timeInterval.start).getTime() : 0;
      const bStart = b.timeInterval?.start ? new Date(b.timeInterval.start).getTime() : 0;
      return aStart - bStart;
    });

    let dayAccumulatedHours = 0;

    for (const entry of dayEntries) {
      const entryHours = calculateDuration(entry);
      const entryRate = getEntryHourlyRate(entry);
      if (!entryHours) continue;

      // Split entry into regular vs overtime based on daily threshold and accumulated hours
      const remainingRegularCapacity = Math.max(0, userCapacity - dayAccumulatedHours);
      const entryRegularHours = Math.min(entryHours, remainingRegularCapacity);
      const entryOvertimeHours = Math.max(0, entryHours - entryRegularHours);
      dayAccumulatedHours += entryHours;

      const entryBaseAmount = (entryRegularHours + entryOvertimeHours) * entryRate;
      const entryPremiumAmount = entryOvertimeHours * entryRate * (userMultiplier - 1);
      const entryTotalAmount = entryBaseAmount + entryPremiumAmount;

      totalHours += entryHours;
      overtimeHours += entryOvertimeHours;
      totalBaseCost += entryBaseAmount;
      totalOtPremium += entryPremiumAmount;

      const startIso = entry.timeInterval?.start || null;
      let endIso = entry.timeInterval?.end || null;
      if (!endIso && startIso && entryHours) {
        const end = new Date(startIso);
        end.setHours(end.getHours() + entryHours);
        endIso = end.toISOString();
      }

      entryDetails.push({
        entryId: entry.id || entry._id || `${userId}-${dateKey}-${entryDetails.length}`,
        date: dateKey,
        start: startIso,
        end: endIso,
        description: entry.description || '',
        project: {
          name: entry.project?.name || 'No Project',
          color: entry.project?.color || '#999999'
        },
        clientName: entry.project?.clientName || entry.clientName || entry.client?.name || '',
        taskName: entry.task?.name || entry.taskName || '',
        tags: Array.isArray(entry.tags) ? entry.tags.map(t => t.name).filter(Boolean) : [],
        billable: !!entry.billable,
        durationHours: round(entryHours),
        regularHours: round(entryRegularHours),
        overtimeHours: round(entryOvertimeHours),
        hourlyRate: round(entryRate),
        otRate: round(entryRate * userMultiplier),
        baseAmount: round(entryBaseAmount),
        premiumAmount: round(entryPremiumAmount),
        totalAmount: round(entryTotalAmount)
      });
    }
  }

  // Sort entries by start time descending for display
  entryDetails.sort((a, b) => {
    const aTime = a.start ? new Date(a.start).getTime() : new Date(a.date).getTime();
    const bTime = b.start ? new Date(b.start).getTime() : new Date(b.date).getTime();
    return bTime - aTime;
  });

  const regularHours = totalHours - overtimeHours;
  const totalCost = totalBaseCost + totalOtPremium;

  return {
    userId,
    userName,
    userEmail,
    multiplier: userMultiplier,
    capacity: userCapacity,
    totalHours: round(totalHours),
    regularHours: round(regularHours),
    overtimeHours: round(overtimeHours),
    baseCost: round(totalBaseCost),
    otPremium: round(totalOtPremium),
    totalCost: round(totalCost),
    entries: entryDetails,
    daysWorked: byDay.size
  };
}

function getEntryHourlyRate(entry) {
  // Clockify stores hourlyRate.amount in cents (or smallest currency unit)
  if (entry.hourlyRate && typeof entry.hourlyRate.amount === 'number') {
    return entry.hourlyRate.amount / 100;
  }
  // Fallback: check if there's a rate at project level
  if (entry.project && entry.project.hourlyRate && entry.project.hourlyRate.amount) {
    return entry.project.hourlyRate.amount / 100;
  }
  return 0; // No rate available
}

function calculateTotalHours(entries) {
  return entries.reduce((total, entry) => {
    return total + calculateDuration(entry);
  }, 0);
}

function calculateDuration(entry) {
  // Reports API returns duration in seconds as a number
  if (typeof entry.duration === 'number') {
    return entry.duration / 3600; // Convert seconds to hours
  }

  // Try timeInterval
  const timeInterval = entry.timeInterval;
  if (!timeInterval) return 0;

  // If duration is provided as ISO 8601 duration string
  if (timeInterval.duration) {
    const match = timeInterval.duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
    if (match) {
      const hours = parseInt(match[1] || 0, 10);
      const minutes = parseInt(match[2] || 0, 10);
      const seconds = parseInt(match[3] || 0, 10);
      return hours + minutes / 60 + seconds / 3600;
    }
  }

  // Calculate from start/end
  if (timeInterval.start && timeInterval.end) {
    const start = new Date(timeInterval.start);
    const end = new Date(timeInterval.end);
    const diffMs = end - start;
    return diffMs / (1000 * 60 * 60);
  }

  return 0;
}

/**
 * Extract and aggregate projects from a day's entries
 * Returns array of { name, color, hours } sorted by hours descending
 */
function extractDayProjects(entries) {
  const projectMap = new Map();

  for (const entry of entries) {
    const name = entry.project?.name || 'No Project';
    const color = entry.project?.color || '#999999';
    const hours = calculateDuration(entry);

    if (!projectMap.has(name)) {
      projectMap.set(name, { name, color, hours: 0 });
    }
    projectMap.get(name).hours += hours;
  }

  // Convert to array and sort by hours (most worked first)
  return Array.from(projectMap.values())
    .map(p => ({ ...p, hours: round(p.hours) }))
    .sort((a, b) => b.hours - a.hours);
}

function calculateTeamSummary(users) {
  const totalHours = users.reduce((sum, u) => sum + u.totalHours, 0);
  const overtimeHours = users.reduce((sum, u) => sum + u.overtimeHours, 0);
  const regularHours = totalHours - overtimeHours;

  const baseCost = users.reduce((sum, u) => sum + u.baseCost, 0);
  const otPremium = users.reduce((sum, u) => sum + u.otPremium, 0);
  const totalCost = baseCost + otPremium;

  return {
    totalHours: round(totalHours),
    regularHours: round(regularHours),
    overtimeHours: round(overtimeHours),
    userCount: users.length,
    costs: {
      baseCost: round(baseCost),
      otPremium: round(otPremium),
      totalCost: round(totalCost),
      currency: 'USD'
    }
  };
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function emptyResult() {
  return {
    users: [],
    summary: {
      totalHours: 0,
      regularHours: 0,
      overtimeHours: 0,
      userCount: 0,
      costs: {
        baseCost: 0,
        otPremium: 0,
        totalCost: 0,
        currency: 'USD'
      }
    },
    config: getConfig()
  };
}

// ============================================================================
// UI State Management
// ============================================================================

function showState(stateName) {
  state.isLoading = stateName === 'loading';

  // Hide all states (with defensive checks)
  if (elements.loadingState) elements.loadingState.classList.add('hidden');
  if (elements.errorState) elements.errorState.classList.add('hidden');
  if (elements.resultsContainer) elements.resultsContainer.classList.add('hidden');
  if (elements.emptyState) elements.emptyState.classList.add('hidden');

  // Disable button during loading
  if (elements.generateBtn) elements.generateBtn.disabled = state.isLoading;

  // Show requested state
  switch (stateName) {
    case 'loading':
      if (elements.loadingState) elements.loadingState.classList.remove('hidden');
      break;
    case 'error':
      if (elements.errorState) elements.errorState.classList.remove('hidden');
      break;
    case 'results':
      if (elements.resultsContainer) elements.resultsContainer.classList.remove('hidden');
      break;
    case 'empty':
    default:
      if (elements.emptyState) elements.emptyState.classList.remove('hidden');
      break;
  }
}

function showError(message) {
  if (elements.errorMessage) elements.errorMessage.textContent = message;
  showState('error');
}

// ============================================================================
// Rendering - User Overrides Table
// ============================================================================

function renderUserOverridesTable() {
  if (!elements.userOverridesBody) {
    console.error('userOverridesBody element not found');
    return;
  }

  const config = getConfig();
  const defaultMultiplier = config.overtimeMultiplier;
  const defaultCapacity = config.dailyThreshold;
  const searchTerm = (elements.overrideSearch?.value || '').trim().toLowerCase();

  if (!state.users || state.users.length === 0) {
    elements.userOverridesBody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 16px;">
          No users found
        </td>
      </tr>
    `;
    return;
  }

  const filteredUsers = searchTerm
    ? state.users.filter(user => {
        const haystack = `${user.name || ''} ${user.email || ''}`.toLowerCase();
        return haystack.includes(searchTerm);
      })
    : state.users;

  if (filteredUsers.length === 0) {
    elements.userOverridesBody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 12px;">
          No users match your search
        </td>
      </tr>
    `;
    return;
  }

  let html = '';

  for (const user of filteredUsers) {
    const userOverride = state.userOverrides[user.id] || {};
    const currentCapacity = userOverride.capacity !== undefined ? userOverride.capacity : defaultCapacity;
    const currentMultiplier = userOverride.multiplier !== undefined ? userOverride.multiplier : defaultMultiplier;

    const hasCapacityOverride = userOverride.capacity !== undefined;
    const hasMultiplierOverride = userOverride.multiplier !== undefined;
    const isOverridden = hasCapacityOverride || hasMultiplierOverride;

    html += `
      <tr>
        <td>
          <span style="font-weight: 500;">${escapeHtml(user.name)}</span>
          ${isOverridden ? '<span class="custom-badge">CUSTOM</span>' : ''}
        </td>
        <td style="text-align: right;">
          <input
            type="number"
            value="${currentCapacity}"
            min="1"
            max="24"
            step="0.5"
            data-user-id="${user.id}"
            data-field="capacity"
            class="override-input"
          />
        </td>
        <td style="text-align: right;">
          <input
            type="number"
            value="${currentMultiplier}"
            min="1"
            max="5"
            step="0.1"
            data-user-id="${user.id}"
            data-field="multiplier"
            class="override-input"
          />
        </td>
      </tr>
    `;
  }

  elements.userOverridesBody.innerHTML = html;

  // Add event listeners to override inputs
  const inputs = elements.userOverridesBody.querySelectorAll('.override-input');
  inputs.forEach(input => {
    input.addEventListener('change', handleOverrideChange);
  });
}

function handleOverrideChange(event) {
  const input = event.target;
  const userId = input.dataset.userId;
  const field = input.dataset.field;
  const value = parseFloat(input.value);

  const config = getConfig();
  const defaultValue = field === 'capacity' ? config.dailyThreshold : config.overtimeMultiplier;

  if (isNaN(value) || value < 1) {
    input.value = defaultValue;
    return;
  }

  // Initialize override object if needed
  if (!state.userOverrides[userId]) {
    state.userOverrides[userId] = {};
  }

  // Check if value is different from default
  if (Math.abs(value - defaultValue) < 0.01) {
    // Same as default, remove this field from override
    delete state.userOverrides[userId][field];

    // If no overrides left, remove the user entry
    if (Object.keys(state.userOverrides[userId]).length === 0) {
      delete state.userOverrides[userId];
    }
  } else {
    // Custom value
    state.userOverrides[userId][field] = value;
  }

  saveUserOverrides();
  recalculateAndRender();
  renderUserOverridesTable(); // Refresh to show/hide CUSTOM badge
}

// ============================================================================
// Rendering - Results Summary and Grouped Table
// ============================================================================

function renderResults(data) {
  if (!data || !data.summary) {
    showError('Invalid response data');
    return;
  }

  const { summary, users } = data;

  // Render summary strip (with defensive checks)
  if (elements.summaryUsers) elements.summaryUsers.textContent = summary.userCount;
  if (elements.summaryTotalHours) elements.summaryTotalHours.textContent = formatHours(summary.totalHours);
  if (elements.summaryRegularHours) elements.summaryRegularHours.textContent = formatHours(summary.regularHours);
  if (elements.summaryOvertimeHours) elements.summaryOvertimeHours.textContent = formatHours(summary.overtimeHours);
  if (elements.summaryTotalCost) elements.summaryTotalCost.textContent = formatCurrency(summary.costs?.totalCost || 0, 'USD');
  if (elements.summaryOtPremium) elements.summaryOtPremium.textContent = formatCurrency(summary.costs?.otPremium || 0, 'USD');
  // Count total entries
  const totalEntries = users.reduce((sum, user) => sum + (Array.isArray(user.entries) ? user.entries.length : 0), 0);
  if (elements.userCountLabel) elements.userCountLabel.textContent = `(${totalEntries} entries)`;

  // Render both tables
  renderSummaryTable(users);
  renderUserTable(users);
}

function renderUserTable(users) {
  if (!elements.reportTableBody) {
    console.error('reportTableBody element not found');
    return;
  }

  if (!users || users.length === 0) {
    elements.reportTableBody.innerHTML = `<tr><td colspan="5" class="empty-state-cell">No time entries found</td></tr>`;
    return;
  }

  // Flatten all entries with user info attached
  const allEntries = [];
  for (const user of users) {
    const userEntries = Array.isArray(user.entries) ? user.entries : [];
    for (const entry of userEntries) {
      allEntries.push({
        ...entry,
        userName: user.userName,
        userId: user.userId
      });
    }
  }

  // Sort by date (most recent first)
  allEntries.sort((a, b) => new Date(b.date) - new Date(a.date));

  let html = '';

  for (const entry of allEntries) {
    const initials = entry.userName.split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();

    // Build project/task/client line like native Clockify: • PROJECT: TASK - CLIENT
    let projectLine = '';
    const projectColor = entry.project?.color || '#00bcd4';
    const projectName = entry.project?.name;

    if (projectName) {
      let projectText = projectName;
      if (entry.taskName) {
        projectText += `: ${entry.taskName}`;
      }
      if (entry.clientName) {
        projectText += ` - ${entry.clientName}`;
      }
      projectLine = `
        <div class="entry-project" style="color: ${projectColor};">
          <span class="project-dot" style="background-color: ${projectColor};"></span>
          ${escapeHtml(projectText)}
        </div>
      `;
    }

    // Only show tags if they exist
    let tagsHtml = '';
    if (entry.tags && entry.tags.length > 0) {
      tagsHtml = `<div class="entry-tags">`;
      entry.tags.forEach(tag => {
        tagsHtml += `<span class="tag-pill">${escapeHtml(tag)}</span>`;
      });
      tagsHtml += `</div>`;
    }

    // Amount display
    const hasRate = entry.hourlyRate && entry.hourlyRate > 0;
    const totalAmountDisplay = hasRate
      ? formatCurrency(entry.totalAmount, 'USD')
      : '';

    // Description - show nothing if empty
    const descriptionHtml = entry.description
      ? `<div class="entry-desc">${escapeHtml(entry.description)}</div>`
      : '';

    // Rate display
    const rateDisplay = hasRate ? `<div class="cell-sub">@${formatCurrency(entry.hourlyRate, 'USD')}/hr</div>` : '';
    const otRateDisplay = hasRate && entry.overtimeHours > 0 ? `<div class="cell-sub">@${formatCurrency(entry.otRate, 'USD')}/hr</div>` : '';

    html += `
      <tr class="entry-row">
        <td class="text-left">
          ${descriptionHtml}
          ${projectLine}
          ${tagsHtml}
        </td>
        <td class="text-left col-user">
          <div class="user-cell">
            <span class="user-avatar">${initials}</span>
            <span class="user-name">${escapeHtml(entry.userName)}</span>
          </div>
        </td>
        <td class="text-right col-time">
          <div class="cell-main">${formatTimeRange(entry.start, entry.end)}</div>
          <div class="cell-sub">${formatDate(entry.date)}</div>
        </td>
        <td class="text-right col-duration">
          <div class="cell-main">${formatHours(entry.durationHours)}</div>
        </td>
        <td class="text-right col-hours">
          <div class="cell-main">${formatHours(entry.regularHours)}</div>
          ${rateDisplay}
        </td>
        <td class="text-right col-hours">
          <div class="cell-main ${entry.overtimeHours > 0 ? 'text-danger' : ''}">${formatHours(entry.overtimeHours)}</div>
          ${otRateDisplay}
        </td>
        <td class="text-right col-amount">
          <div class="cell-main">${totalAmountDisplay}</div>
        </td>
      </tr>
    `;
  }

  elements.reportTableBody.innerHTML = html;
}

// ============================================================================
// Data Grouping Functions
// ============================================================================

function groupDataByClient(users) {
  const byClient = {};
  users.forEach(user => {
    user.entries.forEach(entry => {
      const client = entry.clientName || '(No Client)';
      if (!byClient[client]) {
        byClient[client] = { name: client, regularHours: 0, overtimeHours: 0, totalAmount: 0 };
      }
      byClient[client].regularHours += entry.regularHours || 0;
      byClient[client].overtimeHours += entry.overtimeHours || 0;
      byClient[client].totalAmount += entry.totalAmount || 0;
    });
  });
  return Object.values(byClient).sort((a, b) => b.totalAmount - a.totalAmount);
}

function groupDataByDate(users) {
  const byDate = {};
  users.forEach(user => {
    user.entries.forEach(entry => {
      const date = entry.date;
      if (!byDate[date]) {
        byDate[date] = { date, regularHours: 0, overtimeHours: 0, totalAmount: 0 };
      }
      byDate[date].regularHours += entry.regularHours || 0;
      byDate[date].overtimeHours += entry.overtimeHours || 0;
      byDate[date].totalAmount += entry.totalAmount || 0;
    });
  });
  return Object.values(byDate).sort((a, b) => b.date.localeCompare(a.date));
}

function groupDataByWeek(users) {
  const byWeek = {};
  users.forEach(user => {
    user.entries.forEach(entry => {
      const weekKey = getISOWeekKey(entry.date);
      if (!byWeek[weekKey]) {
        byWeek[weekKey] = { week: weekKey, regularHours: 0, overtimeHours: 0, totalAmount: 0 };
      }
      byWeek[weekKey].regularHours += entry.regularHours || 0;
      byWeek[weekKey].overtimeHours += entry.overtimeHours || 0;
      byWeek[weekKey].totalAmount += entry.totalAmount || 0;
    });
  });
  return Object.values(byWeek).sort((a, b) => b.week.localeCompare(a.week));
}

function groupDataByProject(users) {
  const byProject = {};
  users.forEach(user => {
    user.entries.forEach(entry => {
      const project = entry.project?.name || '(No Project)';
      if (!byProject[project]) {
        byProject[project] = { name: project, regularHours: 0, overtimeHours: 0, totalAmount: 0 };
      }
      byProject[project].regularHours += entry.regularHours || 0;
      byProject[project].overtimeHours += entry.overtimeHours || 0;
      byProject[project].totalAmount += entry.totalAmount || 0;
    });
  });
  return Object.values(byProject).sort((a, b) => b.totalAmount - a.totalAmount);
}

function groupDataByTask(users) {
  const byTask = {};
  users.forEach(user => {
    user.entries.forEach(entry => {
      const task = entry.taskName || '(No Task)';
      if (!byTask[task]) {
        byTask[task] = { name: task, regularHours: 0, overtimeHours: 0, totalAmount: 0 };
      }
      byTask[task].regularHours += entry.regularHours || 0;
      byTask[task].overtimeHours += entry.overtimeHours || 0;
      byTask[task].totalAmount += entry.totalAmount || 0;
    });
  });
  return Object.values(byTask).sort((a, b) => b.totalAmount - a.totalAmount);
}

function getISOWeekKey(dateStr) {
  const d = new Date(dateStr);
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function formatWeekRange(weekKey) {
  // weekKey is like "2025-W49"
  const [year, weekPart] = weekKey.split('-W');
  const weekNum = parseInt(weekPart, 10);

  // Get the first day of the year
  const jan1 = new Date(Date.UTC(parseInt(year), 0, 1));
  const dayOfWeek = jan1.getUTCDay() || 7;

  // Find the first Monday of week 1
  const firstMonday = new Date(jan1);
  firstMonday.setUTCDate(jan1.getUTCDate() + (1 - dayOfWeek) + (dayOfWeek > 4 ? 7 : 0));

  // Calculate the Monday of the requested week
  const weekStart = new Date(firstMonday);
  weekStart.setUTCDate(firstMonday.getUTCDate() + (weekNum - 1) * 7);

  // Calculate Sunday of that week
  const weekEnd = new Date(weekStart);
  weekEnd.setUTCDate(weekStart.getUTCDate() + 6);

  const formatOpts = { month: 'short', day: 'numeric' };
  return `${weekStart.toLocaleDateString('en-US', formatOpts)} - ${weekEnd.toLocaleDateString('en-US', formatOpts)}`;
}

// ============================================================================
// Summary Table (with grouping support)
// ============================================================================

function renderSummaryTable(users) {
  if (!elements.summaryTableBody) {
    console.error('summaryTableBody element not found');
    return;
  }

  if (!users || users.length === 0) {
    elements.summaryTableBody.innerHTML = `<tr><td colspan="7" class="empty-state-cell">No data available</td></tr>`;
    elements.otChartContainer.classList.add('hidden');
    return;
  }

  const groupBy = state.summaryGroupBy || 'user';
  console.log('[DEBUG] renderSummaryTable - groupBy:', groupBy);

  // Dispatch to appropriate render function
  if (groupBy === 'user') {
    console.log('[DEBUG] Rendering by USER');
    renderSummaryByUser(users);
  } else if (groupBy === 'client') {
    console.log('[DEBUG] Rendering by CLIENT');
    renderSummaryByClient(groupDataByClient(users));
  } else if (groupBy === 'project') {
    console.log('[DEBUG] Rendering by PROJECT');
    renderSummaryByProject(groupDataByProject(users));
  } else if (groupBy === 'task') {
    console.log('[DEBUG] Rendering by TASK');
    renderSummaryByTask(groupDataByTask(users));
  } else if (groupBy === 'date') {
    console.log('[DEBUG] Rendering by DATE');
    renderSummaryByDate(groupDataByDate(users));
  } else if (groupBy === 'week') {
    console.log('[DEBUG] Rendering by WEEK');
    renderSummaryByWeek(groupDataByWeek(users));
  } else {
    console.log('[DEBUG] Unknown groupBy value:', groupBy);
  }
}

function renderSummaryByUser(users) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${users.length} users)`;
  }

  // Show OT chart for user grouping
  renderOTChart(users);

  // Sort users by total hours descending
  const sortedUsers = [...users].sort((a, b) => b.totalHours - a.totalHours);

  let html = '';

  // Track totals
  let totalCapacity = 0;
  let totalRegular = 0;
  let totalOvertime = 0;
  let totalHoursAll = 0;
  let totalAmount = 0;

  for (const user of sortedUsers) {
    const initials = user.userName.split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();
    const capacity = user.capacity * user.daysWorked;
    const totalHours = user.totalHours || 0;
    const regularHours = user.regularHours || 0;
    const overtimeHours = user.overtimeHours || 0;

    // Accumulate totals
    totalCapacity += capacity;
    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalHoursAll += totalHours;
    totalAmount += user.totalCost || 0;

    // Calculate utilization percentages
    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    // Overtime styling
    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';

    // High OT row class (>30% overtime)
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <div class="user-cell">
            <span class="user-avatar">${initials}</span>
            <span class="user-name">${escapeHtml(user.userName)}</span>
          </div>
        </td>
        <td class="text-right">${formatHours(capacity)}</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(user.totalCost || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${sortedUsers.length} users)</strong></td>
      <td class="text-right">${formatHours(totalCapacity)}</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderSummaryByClient(clients) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${clients.length} clients)`;
  }

  // Hide OT chart for non-user grouping
  elements.otChartContainer.classList.add('hidden');

  let html = '';
  let totalRegular = 0, totalOvertime = 0, totalAmount = 0;

  for (const client of clients) {
    const regularHours = client.regularHours || 0;
    const overtimeHours = client.overtimeHours || 0;
    const totalHours = regularHours + overtimeHours;

    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalAmount += client.totalAmount || 0;

    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <span class="client-name">${escapeHtml(client.name)}</span>
        </td>
        <td class="text-right">-</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(client.totalAmount || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalHoursAll = totalRegular + totalOvertime;
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${clients.length} clients)</strong></td>
      <td class="text-right">-</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderSummaryByDate(dates) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${dates.length} dates)`;
  }

  // Hide OT chart for non-user grouping
  elements.otChartContainer.classList.add('hidden');

  let html = '';
  let totalRegular = 0, totalOvertime = 0, totalAmount = 0;

  for (const dateItem of dates) {
    const regularHours = dateItem.regularHours || 0;
    const overtimeHours = dateItem.overtimeHours || 0;
    const totalHours = regularHours + overtimeHours;

    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalAmount += dateItem.totalAmount || 0;

    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    // Format date nicely
    const dateObj = new Date(dateItem.date + 'T00:00:00');
    const formattedDate = dateObj.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <span class="date-label">${formattedDate}</span>
        </td>
        <td class="text-right">-</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(dateItem.totalAmount || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalHoursAll = totalRegular + totalOvertime;
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${dates.length} days)</strong></td>
      <td class="text-right">-</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderSummaryByWeek(weeks) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${weeks.length} weeks)`;
  }

  // Hide OT chart for non-user grouping
  elements.otChartContainer.classList.add('hidden');

  let html = '';
  let totalRegular = 0, totalOvertime = 0, totalAmount = 0;

  for (const weekItem of weeks) {
    const regularHours = weekItem.regularHours || 0;
    const overtimeHours = weekItem.overtimeHours || 0;
    const totalHours = regularHours + overtimeHours;

    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalAmount += weekItem.totalAmount || 0;

    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    // Format week nicely
    const weekRange = formatWeekRange(weekItem.week);

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <div class="week-label">
            <span class="week-key">${weekItem.week}</span>
            <span class="week-range">${weekRange}</span>
          </div>
        </td>
        <td class="text-right">-</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(weekItem.totalAmount || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalHoursAll = totalRegular + totalOvertime;
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${weeks.length} weeks)</strong></td>
      <td class="text-right">-</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderSummaryByProject(projects) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${projects.length} projects)`;
  }

  // Hide OT chart
  elements.otChartContainer.classList.add('hidden');

  let html = '';
  let totalRegular = 0, totalOvertime = 0, totalAmount = 0;

  for (const project of projects) {
    const regularHours = project.regularHours || 0;
    const overtimeHours = project.overtimeHours || 0;
    const totalHours = regularHours + overtimeHours;

    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalAmount += project.totalAmount || 0;

    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <span class="project-name">${escapeHtml(project.name)}</span>
        </td>
        <td class="text-right">-</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(project.totalAmount || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalHoursAll = totalRegular + totalOvertime;
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${projects.length} projects)</strong></td>
      <td class="text-right">-</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderSummaryByTask(tasks) {
  // Update count label
  if (elements.summaryUserCount) {
    elements.summaryUserCount.textContent = `(${tasks.length} tasks)`;
  }

  // Hide OT chart
  elements.otChartContainer.classList.add('hidden');

  let html = '';
  let totalRegular = 0, totalOvertime = 0, totalAmount = 0;

  for (const task of tasks) {
    const regularHours = task.regularHours || 0;
    const overtimeHours = task.overtimeHours || 0;
    const totalHours = regularHours + overtimeHours;

    totalRegular += regularHours;
    totalOvertime += overtimeHours;
    totalAmount += task.totalAmount || 0;

    const total = regularHours + overtimeHours;
    const regularPct = total > 0 ? (regularHours / total) * 100 : 0;
    const overtimePct = total > 0 ? (overtimeHours / total) * 100 : 0;

    const otClass = overtimeHours > 0 ? 'text-danger font-bold' : '';
    const highOT = totalHours > 0 && (overtimeHours / totalHours) > 0.3;
    const rowClass = highOT ? 'summary-row high-ot' : 'summary-row';

    html += `
      <tr class="${rowClass}">
        <td class="text-left">
          <span class="task-name">${escapeHtml(task.name)}</span>
        </td>
        <td class="text-right">-</td>
        <td class="text-right">${formatHours(regularHours)}</td>
        <td class="text-right ${otClass}">${formatHours(overtimeHours)}</td>
        <td class="text-right font-bold">${formatHours(totalHours)}</td>
        <td class="text-center">
          <div class="utilization-bar">
            <div class="bar-regular" style="width: ${regularPct}%"></div>
            <div class="bar-overtime" style="width: ${overtimePct}%"></div>
          </div>
          <div class="utilization-label">${regularPct.toFixed(0)}% reg / ${overtimePct.toFixed(0)}% OT</div>
        </td>
        <td class="text-right font-bold">${formatCurrency(task.totalAmount || 0, 'USD')}</td>
      </tr>
    `;
  }

  // Add totals row
  const totalHoursAll = totalRegular + totalOvertime;
  const totalRegPct = totalHoursAll > 0 ? (totalRegular / totalHoursAll) * 100 : 0;
  const totalOtPct = totalHoursAll > 0 ? (totalOvertime / totalHoursAll) * 100 : 0;

  html += `
    <tr class="summary-row totals-row">
      <td class="text-left"><strong>TOTAL (${tasks.length} tasks)</strong></td>
      <td class="text-right">-</td>
      <td class="text-right">${formatHours(totalRegular)}</td>
      <td class="text-right text-danger">${formatHours(totalOvertime)}</td>
      <td class="text-right">${formatHours(totalHoursAll)}</td>
      <td class="text-center">
        <div class="utilization-bar">
          <div class="bar-regular" style="width: ${totalRegPct}%"></div>
          <div class="bar-overtime" style="width: ${totalOtPct}%"></div>
        </div>
        <div class="utilization-label">${totalRegPct.toFixed(0)}% reg / ${totalOtPct.toFixed(0)}% OT</div>
      </td>
      <td class="text-right">${formatCurrency(totalAmount, 'USD')}</td>
    </tr>
  `;

  elements.summaryTableBody.innerHTML = html;
}

function renderOTChart(users) {
  if (!elements.otChart) return;

  // Get top 5 users with most overtime
  const usersWithOT = users
    .filter(u => u.overtimeHours > 0)
    .sort((a, b) => b.overtimeHours - a.overtimeHours)
    .slice(0, 5);

  if (usersWithOT.length === 0) {
    elements.otChartContainer.classList.add('hidden');
    return;
  }

  elements.otChartContainer.classList.remove('hidden');

  const maxOT = Math.max(...usersWithOT.map(u => u.overtimeHours));

  let html = '';
  for (const user of usersWithOT) {
    const pct = maxOT > 0 ? (user.overtimeHours / maxOT) * 100 : 0;
    const otStr = formatHours(user.overtimeHours);

    html += `
      <div class="chart-bar-row">
        <div class="chart-bar-label">${escapeHtml(user.userName)}</div>
        <div class="chart-bar-wrapper">
          <div class="chart-bar-fill" style="width: ${pct}%">
            ${pct > 30 ? `<span class="chart-bar-value">${otStr}</span>` : ''}
          </div>
        </div>
        ${pct <= 30 ? `<span class="chart-bar-value-outside">${otStr}</span>` : ''}
      </div>
    `;
  }

  elements.otChart.innerHTML = html;
}

// ============================================================================
// Formatting Utilities
// ============================================================================

function formatDurationHMS(hours) {
  if (hours === null || hours === undefined || isNaN(hours)) return '0:00:00';
  const totalSeconds = Math.round(hours * 3600);
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return `${h}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

function formatTimeRange(start, end) {
  if (!start && !end) return '--';
  const startStr = start ? formatTime(start) : '--';
  const endStr = end ? formatTime(end) : '--';
  return `${startStr} - ${endStr}`;
}

function formatTime(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' });
  } catch (err) {
    return '--';
  }
}

function formatDateNumeric(dateStr) {
  if (!dateStr) return '';
  try {
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-GB'); // dd/mm/yyyy
  } catch (err) {
    return dateStr;
  }
}

function formatHours(hours) {
  if (hours === null || hours === undefined) return '0h';
  const h = parseFloat(hours);
  if (isNaN(h)) return '0h';

  const wholeHours = Math.floor(h);
  const minutes = Math.round((h - wholeHours) * 60);

  if (minutes === 0) {
    return `${wholeHours}h`;
  }
  return `${wholeHours}h ${minutes}m`;
}

function formatCurrency(amount, currency = 'USD') {
  if (amount === null || amount === undefined) return '$0.00';

  try {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: currency
    }).format(amount);
  } catch (err) {
    return `$${parseFloat(amount).toFixed(2)}`;
  }
}

function formatDate(dateStr) {
  if (!dateStr) return '';

  try {
    const date = new Date(dateStr);
    return new Intl.DateTimeFormat('en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric'
    }).format(date);
  } catch (err) {
    return dateStr;
  }
}

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ============================================================================
// CSV Export
// ============================================================================

function exportToCSV() {
  if (!state.currentData || !state.currentData.users) return;

  const rows = [];

  if (state.activeTab === 'summary') {
    const groupBy = state.summaryGroupBy || 'user';

    if (groupBy === 'user') {
      // Summary by User
      const headers = [
        'User Name', 'Capacity (hrs)', 'Regular Hours', 'Overtime Hours',
        'Total Hours', 'Utilization %', 'Base Amount', 'OT Premium', 'Total Amount'
      ];
      rows.push(headers.join(','));

      for (const user of state.currentData.users) {
        const capacityHrs = user.capacity || 8;
        const totalRegular = user.entries.reduce((sum, e) => sum + (e.regularHours || 0), 0);
        const totalOT = user.entries.reduce((sum, e) => sum + (e.overtimeHours || 0), 0);
        const totalHours = totalRegular + totalOT;
        const totalBase = user.entries.reduce((sum, e) => sum + (e.baseAmount || 0), 0);
        const totalPremium = user.entries.reduce((sum, e) => sum + (e.premiumAmount || 0), 0);
        const totalAmount = user.entries.reduce((sum, e) => sum + (e.totalAmount || 0), 0);
        const utilization = capacityHrs > 0 ? Math.round((totalHours / capacityHrs) * 100) : 0;

        const row = [
          escapeCSV(user.userName),
          capacityHrs,
          totalRegular.toFixed(2),
          totalOT.toFixed(2),
          totalHours.toFixed(2),
          utilization,
          totalBase.toFixed(2),
          totalPremium.toFixed(2),
          totalAmount.toFixed(2)
        ];
        rows.push(row.join(','));
      }
    } else if (groupBy === 'client') {
      // Summary by Client
      const headers = ['Client', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Total Amount'];
      rows.push(headers.join(','));

      const clients = groupDataByClient(state.currentData.users);
      for (const client of clients) {
        const totalHours = (client.regularHours || 0) + (client.overtimeHours || 0);
        const row = [
          escapeCSV(client.name),
          (client.regularHours || 0).toFixed(2),
          (client.overtimeHours || 0).toFixed(2),
          totalHours.toFixed(2),
          (client.totalAmount || 0).toFixed(2)
        ];
        rows.push(row.join(','));
      }
    } else if (groupBy === 'project') {
      // Summary by Project
      const headers = ['Project', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Total Amount'];
      rows.push(headers.join(','));

      const projects = groupDataByProject(state.currentData.users);
      for (const project of projects) {
        const totalHours = (project.regularHours || 0) + (project.overtimeHours || 0);
        const row = [
          escapeCSV(project.name),
          (project.regularHours || 0).toFixed(2),
          (project.overtimeHours || 0).toFixed(2),
          totalHours.toFixed(2),
          (project.totalAmount || 0).toFixed(2)
        ];
        rows.push(row.join(','));
      }
    } else if (groupBy === 'task') {
      // Summary by Task
      const headers = ['Task', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Total Amount'];
      rows.push(headers.join(','));

      const tasks = groupDataByTask(state.currentData.users);
      for (const task of tasks) {
        const totalHours = (task.regularHours || 0) + (task.overtimeHours || 0);
        const row = [
          escapeCSV(task.name),
          (task.regularHours || 0).toFixed(2),
          (task.overtimeHours || 0).toFixed(2),
          totalHours.toFixed(2),
          (task.totalAmount || 0).toFixed(2)
        ];
        rows.push(row.join(','));
      }
    } else if (groupBy === 'date') {
      // Summary by Date
      const headers = ['Date', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Total Amount'];
      rows.push(headers.join(','));

      const dates = groupDataByDate(state.currentData.users);
      for (const dateItem of dates) {
        const totalHours = (dateItem.regularHours || 0) + (dateItem.overtimeHours || 0);
        const row = [
          dateItem.date,
          (dateItem.regularHours || 0).toFixed(2),
          (dateItem.overtimeHours || 0).toFixed(2),
          totalHours.toFixed(2),
          (dateItem.totalAmount || 0).toFixed(2)
        ];
        rows.push(row.join(','));
      }
    } else if (groupBy === 'week') {
      // Summary by Week
      const headers = ['Week', 'Date Range', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Total Amount'];
      rows.push(headers.join(','));

      const weeks = groupDataByWeek(state.currentData.users);
      for (const weekItem of weeks) {
        const totalHours = (weekItem.regularHours || 0) + (weekItem.overtimeHours || 0);
        const row = [
          weekItem.week,
          escapeCSV(formatWeekRange(weekItem.week)),
          (weekItem.regularHours || 0).toFixed(2),
          (weekItem.overtimeHours || 0).toFixed(2),
          totalHours.toFixed(2),
          (weekItem.totalAmount || 0).toFixed(2)
        ];
        rows.push(row.join(','));
      }
    }
  } else {
    // Detailed export - one row per entry
    const headers = [
      'User Name', 'Date', 'Start Time', 'End Time', 'Project', 'Client', 'Task', 'Tags',
      'Description', 'Duration (hrs)', 'Regular Hours', 'Overtime Hours',
      'Hourly Rate', 'OT Rate', 'Base Amount', 'OT Premium', 'Total Amount'
    ];
    rows.push(headers.join(','));

    for (const user of state.currentData.users) {
      for (const entry of user.entries) {
        const row = [
          escapeCSV(user.userName),
          entry.date,
          formatTimeForCSV(entry.start),
          formatTimeForCSV(entry.end),
          escapeCSV(entry.project?.name || ''),
          escapeCSV(entry.clientName || ''),
          escapeCSV(entry.taskName || ''),
          escapeCSV((entry.tags || []).join('; ')),
          escapeCSV(entry.description || ''),
          entry.durationHours,
          entry.regularHours,
          entry.overtimeHours,
          entry.hourlyRate,
          entry.otRate,
          entry.baseAmount,
          entry.premiumAmount,
          entry.totalAmount
        ];
        rows.push(row.join(','));
      }
    }
  }

  const csvContent = rows.join('\n');
  downloadCSV(csvContent);
}

function escapeCSV(str) {
  if (!str) return '';
  str = String(str);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function formatTimeForCSV(isoString) {
  if (!isoString) return '';
  try {
    return new Date(isoString).toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' });
  } catch {
    return '';
  }
}

function downloadCSV(csvContent) {
  const startDate = elements.startDate.value;
  const endDate = elements.endDate.value;
  const filename = `Overtime_Report_${startDate}_to_${endDate}.csv`;

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
