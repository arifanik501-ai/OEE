/**
 * ══════════════════════════════════════════════════════════════════════════════
 * MEP FAN LTD. - Multi-Tab Excel Web Spreadsheet Engine
 * Tabs:
 *   - "Fan Lathe" (3 Machines: Rotor, Bottom Cover, Top Cover)
 *   - "Fan Auto Powder Coating" (4 Machines: APC (Blade), APC (Downpipe), APC (Body, Cover), Blade Rivet)
 * Dynamic Month/Year Selection + Strict Numeric Validation + Real-Time OEE
 * ══════════════════════════════════════════════════════════════════════════════
 */

'use strict';

// ─── 39 COLUMNS CONFIGURATION (A to AM - COMPACT FIT FOR 100% ZOOM) ───────────
const EXCEL_COLUMNS = [
  { col: 'A', label: 'Date', width: 66, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'B', label: 'Day', width: 34, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'C', label: 'Shift', width: 46, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'D', label: 'Machine Name', width: 105, isReadOnly: true, zone: 'sky', align: 'left' },
  { col: 'E', label: 'Machine Capacity (Pcs)', width: 40, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'F', label: 'Actual Prd. (Pcs)', width: 40, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'G', label: 'Rejection (Pcs)', width: 32, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'H', label: 'Planned Prd. Time (Min)', width: 40, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'I', label: 'Expected DownTime (Min)', width: 32, isFormula: true, formula: '=IF(E{r}<>"",30,0)', zone: 'sky', align: 'right' },
  { col: 'J', label: 'Total Prd. Run Time (Min)', width: 40, isFormula: true, formula: '=H{r}-AH{r}', zone: 'sky', align: 'right' },
  
  // 23 Downtime columns K to AG (Purple Zone)
  { col: 'K', label: 'Heater/Coil  Problem', code: 10, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'L', label: 'Power Shutdown', code: 11, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'M', label: 'Machine Breakdown', code: 12, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'N', label: 'Die/ Mold Problem', code: 13, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'O', label: 'Model/ Die Change', code: 14, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'P', label: 'Air Presser Problem', code: 15, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Q', label: 'Water line Problem', code: 16, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'R', label: 'Lift Problem', code: 17, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'S', label: 'D Coil Insert', code: 18, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'T', label: 'RM Shortage', code: 19, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'U', label: 'Crean Problem', code: 20, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'V', label: 'Worker Absent', code: 21, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'W', label: 'Printer M/C Problem ', code: 22, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'X', label: 'UPS shutdown', code: 23, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Y', label: 'Load Problem', code: 24, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Z', label: 'Namaz', code: 25, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AA', label: 'Conveyor Belt Problem', code: 26, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AB', label: 'Fitting Problem', code: 27, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AC', label: 'Gas Presser Problem', code: 28, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AD', label: 'Mold polish & Clean', code: 29, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AE', label: 'Alu. Ash Extraction', code: 30, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AF', label: 'Robot Problems', code: 31, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AG', label: 'Alu. Recipe Problem', code: 32, width: 32, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  
  // KPI Columns AH to AL (Cyan Zone)
  { col: 'AH', label: 'Total Down Time (Mins)', width: 40, isFormula: true, formula: '=SUM(K{r}:AG{r})', zone: 'cyan', align: 'right' },
  { col: 'AI', label: 'Availability (%)', width: 34, isFormula: true, formula: '=IFERROR(J{r}/H{r},"0")', isPercent: true, zone: 'cyan', align: 'right' },
  { col: 'AJ', label: 'Performance (%)', width: 34, isFormula: true, formula: '=IFERROR(F{r}/E{r},"0")', isPercent: true, zone: 'cyan', align: 'right' },
  { col: 'AK', label: 'Quality (%)', width: 34, isFormula: true, formula: '=IFERROR(F{r}/(F{r}+G{r}),"0")', isPercent: true, zone: 'cyan', align: 'right' },
  { col: 'AL', label: 'OEE (%)', width: 34, isFormula: true, formula: '=IFERROR(AK{r}*AJ{r}*AI{r},"0")', isPercent: true, zone: 'cyan', align: 'right' },
  
  // Column AM (Remarks) (Sky Blue Zone) - Only Column that allows text!
  { col: 'AM', label: 'Remarks', width: 110, zone: 'sky', align: 'left' }
];

// ─── MULTI-TAB WORKBOOK CONFIGURATION (SUMMARY + 9 PRODUCTION TABS) ───────────
const SHEET_TABS = {
  summary_oee_yearly: {
    id: 'summary_oee_yearly',
    name: 'Summary of OEE',
    colorArgb: 'FFEA580C',
    title: 'MEP FAN LTD.',
    subtitle: 'Summary of OEE',
    isSummary: true,
    machines: []
  },
  summary_downtime: {
    id: 'summary_downtime',
    name: 'Total Downtime Report',
    colorArgb: 'FFF59E0B',
    title: 'MEP FAN LTD.',
    subtitle: 'Total Downtime Report',
    isSummary: true,
    machines: []
  },
  summary_production: {
    id: 'summary_production',
    name: 'Production Output',
    colorArgb: 'FF3B82F6',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Output',
    isSummary: true,
    machines: []
  },
  summary_status: {
    id: 'summary_status',
    name: 'Downtime & Running Status',
    colorArgb: 'FF10B981',
    title: 'MEP FAN LTD.',
    subtitle: 'MEP Fan Limited- Down Time & Running Time Status',
    isSummary: true,
    machines: []
  },
  summary_oee: {
    id: 'summary_oee',
    name: 'OEE Report',
    colorArgb: 'FF8B5CF6',
    title: 'MEP FAN LTD.',
    subtitle: 'OEE Report',
    isSummary: true,
    machines: []
  },
  fan_lathe: {
    id: 'fan_lathe',
    name: 'Fan Lathe',
    colorArgb: 'FF0054A6',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Rotor (CNC)',
      'Bottom Cover (CNC)',
      'Top Cover (CNC)'
    ]
  },
  fan_powder: {
    id: 'fan_powder',
    name: 'Fan Auto Powder Coating',
    colorArgb: 'FF8B5CF6',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'APC (Blade)',
      'APC (Downpipe)',
      'APC (Body, Cover)',
      'Blade Rivet'
    ],
    timeGroups: [
      { startIdx: 0, count: 3, label: 'APC (Powder Coating)' },
      { startIdx: 3, count: 1, label: 'Blade Rivet' }
    ]
  },
  fan_die_casting: {
    id: 'fan_die_casting',
    name: 'Fan Auto Die Casting',
    colorArgb: 'FF10B981',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Auto Die Casting (Cover)',
      'Auto Die Casting (Body)'
    ]
  },
  fan_power_press: {
    id: 'fan_power_press',
    name: 'Fan Power Press & Stamping',
    colorArgb: 'FFF59E0B',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Shank',
      'Clamp',
      'Rotor & Stator',
      'Blade',
      'Downpipe'
    ]
  },
  fan_assemble: {
    id: 'fan_assemble',
    name: 'Fan Assemble',
    colorArgb: 'FFEF4444',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Conveyor Belt'
    ]
  },
  fan_armature: {
    id: 'fan_armature',
    name: 'Fan Armature',
    colorArgb: 'FFDC2626',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Winding Machine'
    ]
  },
  fan_rechargeable: {
    id: 'fan_rechargeable',
    name: 'Fan Rechargeable',
    colorArgb: 'FF06B6D4',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Conveyor Belt'
    ]
  },
  exhaust_fan: {
    id: 'exhaust_fan',
    name: 'Exhaust Fan',
    colorArgb: 'FF3B82F6',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Conveyor Belt'
    ]
  },
  capacitor: {
    id: 'capacitor',
    name: 'Capacitor',
    colorArgb: 'FFF43F5E',
    title: 'MEP FAN LTD.',
    subtitle: 'Production Performance Analysis Report',
    machines: [
      'Conveyor Belt'
    ]
  }
};

const SUMMARY_DEPTS = [
  { id: 'fan_die_casting', name: 'Fan Auto Die Casting' },
  { id: 'fan_power_press', name: 'Fan Power Press & Stamping' },
  { id: 'fan_powder', name: 'Fan Auto Powder Coating' },
  { id: 'fan_lathe', name: 'Fan Lathe' },
  { id: 'fan_assemble', name: 'Fan Assemble' },
  { id: 'fan_armature', name: 'Fan Armature' },
  { id: 'fan_rechargeable', name: 'Rechargeable Fan' },
  { id: 'exhaust_fan', name: 'Exhaust Fan' },
  { id: 'capacitor', name: 'Capacitor' }
];

// ─── USER ROLES & ACCESS CONTROL CONFIGURATION ───────────────────────────────
const USERS_CONFIG = {
  admin: {
    id: 'admin',
    name: 'Admin',
    role: 'System Administrator',
    pin: '8250',
    icon: '👑',
    svgIcon: '<svg class="w-7 h-7 text-amber-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/></svg>',
    color: '#D97706',
    description: 'Master Administrator • Full System Access',
    allowedDepts: ['fan_lathe', 'fan_powder', 'fan_die_casting', 'fan_power_press', 'fan_assemble', 'fan_armature', 'fan_rechargeable', 'exhaust_fan', 'capacitor']
  },
  takbir: {
    id: 'takbir',
    name: 'Takbir',
    role: 'Production Incharge',
    pin: '9696',
    icon: '⚡',
    svgIcon: '<svg class="w-7 h-7 text-rose-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>',
    color: '#E11D48',
    description: 'Fan Armature & Winding Department',
    allowedDepts: ['fan_armature']
  },
  monir: {
    id: 'monir',
    name: 'Monir',
    role: 'Production Incharge',
    pin: '2222',
    icon: '🏭',
    svgIcon: '<svg class="w-7 h-7 text-emerald-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 8h20M2 16h20M6 4v16M18 4v16M10 8v8M14 8v8"/></svg>',
    color: '#059669',
    description: 'Fan Power Press & Auto Die Casting',
    allowedDepts: ['fan_power_press', 'fan_die_casting']
  },
  anwar: {
    id: 'anwar',
    name: 'Anwar',
    role: 'Production Incharge',
    pin: '1111',
    icon: '⚙️',
    svgIcon: '<svg class="w-7 h-7 text-sky-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/></svg>',
    color: '#0284C7',
    description: 'Fan Lathe & Auto Powder Coating',
    allowedDepts: ['fan_lathe', 'fan_powder']
  },
  bikash: {
    id: 'bikash',
    name: 'Bikash',
    role: 'Production Incharge',
    pin: '0011',
    icon: '📦',
    svgIcon: '<svg class="w-7 h-7 text-purple-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="14" x="2" y="3" rx="2"/><line x1="8" x2="16" y1="21" y2="21"/><line x1="12" x2="12" y1="17" y2="21"/><path d="m7 8 5 3 5-3"/></svg>',
    color: '#7C3AED',
    description: 'Rechargeable, Capacitor & Exhaust Fan',
    allowedDepts: ['fan_rechargeable', 'capacitor', 'exhaust_fan']
  },
  anik: {
    id: 'anik',
    name: 'Anik',
    role: 'Assemble & Fan Incharge',
    pin: '1234',
    icon: '🔧',
    svgIcon: '<svg class="w-7 h-7 text-teal-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/></svg>',
    color: '#0D9488',
    description: 'Fan Assemble, Rechargeable, Exhaust & Capacitor',
    allowedDepts: ['fan_assemble', 'fan_rechargeable', 'exhaust_fan', 'capacitor']
  },
  viewer: {
    id: 'viewer',
    name: 'Viewer',
    role: 'Read-Only Mode',
    pin: null,
    isReadOnly: true,
    icon: '👁️',
    svgIcon: '<svg class="w-7 h-7 text-indigo-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>',
    color: '#4F46E5',
    description: 'Full View Mode • Read Only (Editing Disabled)',
    allowedDepts: ['fan_lathe', 'fan_powder', 'fan_die_casting', 'fan_power_press', 'fan_assemble', 'fan_armature', 'fan_rechargeable', 'exhaust_fan', 'capacitor']
  }
};

let CurrentUser = null;
let SelectedLoginUser = null;
let EnteredPin = '';

function getActiveSummaryDepts() {
  if (!CurrentUser || CurrentUser.id === 'admin' || CurrentUser.id === 'viewer') {
    return SUMMARY_DEPTS;
  }
  const allowed = CurrentUser.allowedDepts || [];
  return SUMMARY_DEPTS.filter(dept => allowed.includes(dept.id));
}

let ACTIVE_TAB = 'summary_oee_yearly';

function getMachinesForTab(tabId) {
  return SHEET_TABS[tabId]?.machines || SHEET_TABS.fan_lathe.machines;
}

function getRowsPerDay(tabId = ACTIVE_TAB) {
  return getMachinesForTab(tabId).length;
}

const TIME_COLUMNS = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI'];

function getTimeGroupInfo(tabId, machineIdx) {
  const tabInfo = SHEET_TABS[tabId];
  if (!tabInfo || !tabInfo.timeGroups) {
    return { isMaster: true, span: 1, masterIdx: machineIdx, isSlave: false, count: 1 };
  }
  for (const g of tabInfo.timeGroups) {
    if (machineIdx >= g.startIdx && machineIdx < g.startIdx + g.count) {
      const isMaster = (machineIdx === g.startIdx);
      return {
        isMaster,
        span: isMaster ? g.count : 0,
        masterIdx: g.startIdx,
        isSlave: !isMaster,
        count: g.count
      };
    }
  }
  return { isMaster: true, span: 1, masterIdx: machineIdx, isSlave: false, count: 1 };
}

// ─── MONTH & YEAR CONTROLLER ──────────────────────────────────────────────────
const MonthYearState = {
  year: 2026,
  monthIndex: 7, // August
  monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  monthShortNames: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  dayNames: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
};

function getDaysInSelectedMonth() {
  return new Date(MonthYearState.year, MonthYearState.monthIndex + 1, 0).getDate();
}

function getMaxActiveRow(tabId = ACTIVE_TAB) {
  return 5 + (getDaysInSelectedMonth() * getRowsPerDay(tabId));
}

function getStorageKey(tabId, year, monthIndex) {
  return `mep_oee_v23_${tabId}_${year}_${monthIndex}`;
}

// ─── DYNAMIC LOCK & TODAY ENGINE ──────────────────────────────────────────────
function isDateInCurrentMonth(dayNum) {
  const now = new Date();
  return (
    MonthYearState.year === now.getFullYear() &&
    MonthYearState.monthIndex === now.getMonth() &&
    dayNum === now.getDate()
  );
}

function getDayFromRow(rowNum, tabId = ACTIVE_TAB) {
  if (rowNum < 6) return 0;
  const rowsPerDay = getRowsPerDay(tabId);
  return Math.floor((rowNum - 6) / rowsPerDay) + 1;
}

function isRowLocked(rowNum) {
  if (rowNum === 5) return false;
  const dayNum = getDayFromRow(rowNum);

  const now = new Date();
  const curYear = now.getFullYear();
  const curMonth = now.getMonth();
  const curDay = now.getDate();

  // Past years/months: fully unlocked
  if (MonthYearState.year < curYear) return false;
  if (MonthYearState.year === curYear && MonthYearState.monthIndex < curMonth) return false;

  // Future years/months: fully locked
  if (MonthYearState.year > curYear) return true;
  if (MonthYearState.year === curYear && MonthYearState.monthIndex > curMonth) return true;

  // Current month: progressive daily unlock
  return dayNum > curDay;
}

// ─── APPLICATION STATE ────────────────────────────────────────────────────────
const SheetState = {
  rows: [],
  totals: {},
  selected: { row: 6, colLetter: 'E', element: null },
  rangeSelection: { start: null, end: null, isSelecting: false },
  isEditing: false,
  activeInput: null,
  undoStack: [],
  redoStack: [],
  searchResults: [],
  searchIndex: 0
};

// ─── INITIALIZATION & STATE RESTORATION ───────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  restoreAppState();
  initMonthYearSelectors();
  initSheetTabButtons();
  initModalDtInputs();
  bindExcelEvents();
  initAuthSystem();
  initFirebase();
});

// ─── ROLE-BASED AUTHENTICATION & LOGIN SYSTEM ────────────────────────────────
function initAuthSystem() {
  updatePortalClock();
  setInterval(updatePortalClock, 1000);

  // Bind profile cards
  document.querySelectorAll('.user-portal-card').forEach(card => {
    card.addEventListener('click', () => {
      const uid = card.getAttribute('data-user-id');
      if (uid && USERS_CONFIG[uid]) {
        openPinModal(uid);
      }
    });
  });

  // Modal close
  document.getElementById('btnClosePinModal')?.addEventListener('click', closePinModal);

  // Numpad clicks
  document.querySelectorAll('.numpad-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const key = btn.getAttribute('data-key');
      handlePinKey(key);
    });
  });

  // Submit button
  document.getElementById('btnSubmitPin')?.addEventListener('click', () => {
    verifyAndLogin();
  });

  // View Only Mode Buttons
  document.getElementById('btnEnterViewOnlyMode')?.addEventListener('click', () => {
    applyUserLogin(USERS_CONFIG.viewer, true);
  });
  document.getElementById('btnBannerViewOnly')?.addEventListener('click', () => {
    applyUserLogin(USERS_CONFIG.viewer, true);
  });

  // Hidden input for keyboard typing
  const pinInput = document.getElementById('pinHiddenInput');
  pinInput?.addEventListener('input', (e) => {
    const val = e.target.value;
    EnteredPin = val.slice(0, 4);
    updatePinDots();
    if (EnteredPin.length === 4) {
      verifyAndLogin();
    }
  });

  // Modal keyboard keydown
  window.addEventListener('keydown', (e) => {
    const modal = document.getElementById('pinModal');
    if (!modal || modal.classList.contains('hidden')) return;

    if (e.key === 'Escape') {
      closePinModal();
    } else if (e.key === 'Enter') {
      verifyAndLogin();
    } else if (e.key === 'Backspace') {
      handlePinKey('backspace');
    } else if (/^[0-9]$/.test(e.key)) {
      handlePinKey(e.key);
    }
  });

  // Switch User button in top action bar
  document.getElementById('btnSwitchUser')?.addEventListener('click', () => {
    logoutUser();
  });

  // Check existing session
  const savedUid = sessionStorage.getItem('mep_auth_user');
  if (savedUid && USERS_CONFIG[savedUid]) {
    applyUserLogin(USERS_CONFIG[savedUid], false);
  } else {
    showLoginPortal();
  }
}

function updatePortalClock() {
  const now = new Date();
  const dateStr = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  const timeStr = now.toLocaleTimeString('en-GB');

  const elDate = document.getElementById('portalDateStr');
  const elTime = document.getElementById('portalTimeStr');
  if (elDate) elDate.textContent = dateStr;
  if (elTime) elTime.textContent = timeStr;

  const legacyClock = document.getElementById('portalClock');
  if (legacyClock) {
    legacyClock.textContent = `🗓️ ${dateStr} • ⏰ ${timeStr} • 🔒 Role-Based Secure Access`;
  }
}

function showLoginPortal() {
  const portal = document.getElementById('loginPortal');
  const appContainer = document.getElementById('appMainContainer');
  if (portal) {
    portal.classList.remove('hidden');
    portal.style.display = 'flex';
  }
  if (appContainer) {
    appContainer.classList.add('hidden');
    appContainer.style.display = 'none';
  }
}

function hideLoginPortal() {
  const portal = document.getElementById('loginPortal');
  const appContainer = document.getElementById('appMainContainer');
  if (portal) {
    portal.classList.add('hidden');
    portal.style.display = 'none';
  }
  if (appContainer) {
    appContainer.classList.remove('hidden');
    appContainer.style.display = 'flex';
  }
}

function openPinModal(userId) {
  SelectedLoginUser = USERS_CONFIG[userId];
  if (!SelectedLoginUser) return;
  EnteredPin = '';
  
  const iconEl = document.getElementById('pinModalIcon');
  const nameEl = document.getElementById('pinModalUserName');
  const roleEl = document.getElementById('pinModalUserRole');
  const errEl = document.getElementById('pinErrorMsg');
  const hiddenInp = document.getElementById('pinHiddenInput');

  if (iconEl) {
    iconEl.innerHTML = SelectedLoginUser.svgIcon || SelectedLoginUser.icon;
  }
  if (nameEl) nameEl.textContent = SelectedLoginUser.name.toUpperCase();
  if (roleEl) roleEl.textContent = SelectedLoginUser.description;
  if (errEl) errEl.textContent = '';
  if (hiddenInp) {
    hiddenInp.value = '';
    hiddenInp.focus();
  }

  updatePinDots();

  const modal = document.getElementById('pinModal');
  if (modal) {
    modal.classList.remove('hidden');
  }
}

function closePinModal() {
  SelectedLoginUser = null;
  EnteredPin = '';
  const modal = document.getElementById('pinModal');
  if (modal) {
    modal.classList.add('hidden');
  }
}

function updatePinDots() {
  for (let i = 0; i < 4; i++) {
    const dot = document.getElementById(`dot${i}`);
    if (dot) {
      if (i < EnteredPin.length) {
        dot.classList.add('active');
      } else {
        dot.classList.remove('active');
      }
    }
  }
}

function handlePinKey(key) {
  if (key === 'clear') {
    EnteredPin = '';
  } else if (key === 'backspace') {
    EnteredPin = EnteredPin.slice(0, -1);
  } else if (/^[0-9]$/.test(key) && EnteredPin.length < 4) {
    EnteredPin += key;
  }

  updatePinDots();
  const hiddenInp = document.getElementById('pinHiddenInput');
  if (hiddenInp) hiddenInp.value = EnteredPin;

  if (EnteredPin.length === 4) {
    verifyAndLogin();
  }
}

function verifyAndLogin() {
  if (!SelectedLoginUser) return;

  if (EnteredPin === SelectedLoginUser.pin) {
    applyUserLogin(SelectedLoginUser, true);
  } else {
    const errEl = document.getElementById('pinErrorMsg');
    if (errEl) {
      errEl.textContent = '❌ ভুল পাসওয়ার্ড! সঠিক পাসওয়ার্ড দিন।';
    }

    EnteredPin = '';
    updatePinDots();
    const hiddenInp = document.getElementById('pinHiddenInput');
    if (hiddenInp) {
      hiddenInp.value = '';
      hiddenInp.focus();
    }
  }
}

function applyUserLogin(user, showWelcome = true) {
  CurrentUser = user;
  sessionStorage.setItem('mep_auth_user', CurrentUser.id);
  closePinModal();

  // 1. Determine Target Sheet Tab for this User immediately
  let targetTab = ACTIVE_TAB;
  if (CurrentUser.id === 'admin' || CurrentUser.id === 'viewer') {
    if (!targetTab || !isTabAllowedForUser(targetTab)) {
      targetTab = 'summary_oee_yearly';
    }
  } else if (CurrentUser.allowedDepts && CurrentUser.allowedDepts.length > 0) {
    // For Incharge users: go directly to their assigned production department tab
    if (!CurrentUser.allowedDepts.includes(targetTab)) {
      targetTab = CurrentUser.allowedDepts[0];
    }
  } else {
    targetTab = 'summary_oee_yearly';
  }

  ACTIVE_TAB = targetTab;
  localStorage.setItem('mep_oee_active_tab', targetTab);

  // 2. Update Top Bar User Badge & Controls
  const iconEl = document.getElementById('currentUserIcon');
  const nameEl = document.getElementById('currentUserName');
  const saveBtn = document.getElementById('btnSaveGrid');
  const excelBtn = document.getElementById('btnExportExcel');
  const syncBtn = document.getElementById('btnCloudSync');
  const formulaInput = document.getElementById('formulaBarInput');

  if (syncBtn) {
    if (CurrentUser.id === 'admin') {
      syncBtn.style.display = 'flex';
      syncBtn.className = 'h-[44px] px-[20px] py-[10px] rounded-lg bg-[#16A34A] hover:bg-[#15803D] text-white border border-emerald-500 text-xs font-bold shadow-sm flex items-center gap-1.5 cursor-pointer transition';
      syncBtn.innerHTML = '<span id="syncSpinIcon" class="text-sm inline-block">🟢</span><span>Publish</span>';
      syncBtn.title = 'সব ডিপার্টমেন্টের ডাটা একত্রিত করে লাইভ পাবলিশ করুন';
    } else {
      syncBtn.style.display = 'none'; // Admin ছাড়া সব এন্ট্রি শিট ও ইউজারের জন্য বাটনটি সম্পূর্ণ বাদ
    }
  }

  if (CurrentUser.isReadOnly) {
    if (iconEl) iconEl.innerHTML = '👁️';
    if (nameEl) nameEl.innerHTML = `<span class="text-amber-300 font-extrabold text-xs block tracking-wide">VIEW ONLY</span><span class="text-slate-300 text-[10px] font-medium block">Read-Only Mode</span>`;
    if (saveBtn) saveBtn.style.display = 'none';
    if (excelBtn) excelBtn.style.display = 'none';
    if (formulaInput) formulaInput.readOnly = true;
  } else {
    if (iconEl) {
      iconEl.innerHTML = CurrentUser.id === 'admin' ? '🛡️' : (CurrentUser.svgIcon || CurrentUser.icon);
    }
    if (nameEl) {
      nameEl.innerHTML = `<span class="text-white font-black text-xs block tracking-wide">${CurrentUser.name.toUpperCase()}</span><span class="text-blue-200 text-[10px] font-medium block">${CurrentUser.role}</span>`;
    }
    if (saveBtn) saveBtn.style.display = 'flex';
    if (excelBtn) excelBtn.style.display = 'flex';
    if (formulaInput) formulaInput.readOnly = false;
  }

  // 3. Update Tab Bar Visibility for User & Visual Active Class
  updateTabBarForUser();
  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    if (btn.dataset.tab === targetTab) {
      btn.classList.add('active-tab');
      btn.scrollIntoView({ inline: 'nearest', behavior: 'auto' });
    } else {
      btn.classList.remove('active-tab');
    }
  });

  // 4. Pre-render the EXACT target sheet BEFORE unhiding container
  loadSheetData();
  recalculateAllFormulas();
  renderExcelTable();
  selectInitialCell();

  // 5. Seamlessly reveal the main container with zero blink/flicker
  document.documentElement.classList.remove('user-unauthenticated');
  document.documentElement.classList.add('user-authenticated');
  hideLoginPortal();

  if (showWelcome) {
    if (CurrentUser.isReadOnly) {
      showToast(`👁️ ভিউ-অনলি মোড: সমস্ত রিপোর্ট ও ড্যাশবোর্ড দেখা যাবে (এডিট বন্ধ)।`, 'info');
    } else {
      showToast(`🎉 স্বাগতম ${CurrentUser.name}! এক্সেস আনলক হয়েছে।`, 'success');
    }
  }
}

function logoutUser() {
  CurrentUser = null;
  sessionStorage.removeItem('mep_auth_user');
  document.documentElement.classList.remove('user-authenticated');
  document.documentElement.classList.add('user-unauthenticated');
  showLoginPortal();
  showToast('🔒 সফলভাবে লগআউট হয়েছে।', 'info');
}

function isTabAllowedForUser(tabId) {
  if (!CurrentUser || CurrentUser.id === 'admin') return true;
  // All 5 summary tabs are allowed for everyone
  if (SHEET_TABS[tabId]?.isSummary) return true;
  // Production tab allowed only if in user's permitted depts
  return CurrentUser.allowedDepts.includes(tabId);
}

function updateTabBarForUser() {
  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    const tabId = btn.getAttribute('data-tab');
    if (isTabAllowedForUser(tabId)) {
      btn.style.display = '';
    } else {
      btn.style.display = 'none';
    }
  });
}

function restoreAppState() {
  const savedTab = localStorage.getItem('mep_oee_active_tab');
  if (savedTab && SHEET_TABS[savedTab]) {
    ACTIVE_TAB = savedTab;
  }

  const savedMonth = localStorage.getItem('mep_oee_selected_month');
  if (savedMonth !== null && !isNaN(savedMonth)) {
    MonthYearState.monthIndex = Number(savedMonth);
  }

  const savedYear = localStorage.getItem('mep_oee_selected_year');
  if (savedYear !== null && !isNaN(savedYear)) {
    MonthYearState.year = Number(savedYear);
  }
}

function initMonthYearSelectors() {
  const selectMonth = document.getElementById('selectMonth');
  const selectYear = document.getElementById('selectYear');

  if (selectMonth) {
    selectMonth.value = MonthYearState.monthIndex;
    selectMonth.addEventListener('change', (e) => {
      saveSheetData(false);
      MonthYearState.monthIndex = Number(e.target.value);
      localStorage.setItem('mep_oee_selected_month', MonthYearState.monthIndex);
      onMonthYearChange();
    });
  }

  if (selectYear) {
    selectYear.value = MonthYearState.year;
    selectYear.addEventListener('change', (e) => {
      saveSheetData(false);
      MonthYearState.year = Number(e.target.value);
      localStorage.setItem('mep_oee_selected_year', MonthYearState.year);
      onMonthYearChange();
    });
  }
}

function initSheetTabButtons() {
  const container = document.getElementById('sheetTabsContainer');
  const btnLeft = document.getElementById('btnScrollTabsLeft');
  const btnRight = document.getElementById('btnScrollTabsRight');

  btnLeft?.addEventListener('click', () => {
    container?.scrollBy({ left: -140, behavior: 'auto' });
  });

  btnRight?.addEventListener('click', () => {
    container?.scrollBy({ left: 140, behavior: 'auto' });
  });

  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tabId = btn.dataset.tab;
      if (tabId && SHEET_TABS[tabId]) {
        switchSheetTab(tabId);
      }
    });
  });

  // Ensure active tab class matches restored ACTIVE_TAB & scroll active tab into view
  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    if (btn.dataset.tab === ACTIVE_TAB) {
      btn.classList.add('active-tab');
      btn.scrollIntoView({ inline: 'nearest', behavior: 'auto' });
    } else {
      btn.classList.remove('active-tab');
    }
  });

  updateTabBarForUser();
}

function switchSheetTab(tabId) {
  if (!isTabAllowedForUser(tabId)) {
    showToast('⚠️ এই ট্যাবে আপনার এক্সেস নেই!', 'warning');
    return;
  }
  if (ACTIVE_TAB === tabId) return;

  saveSheetData(false);
  ACTIVE_TAB = tabId;
  localStorage.setItem('mep_oee_active_tab', tabId);

  // Update button visual active states
  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    if (btn.dataset.tab === tabId) {
      btn.classList.add('active-tab');
    } else {
      btn.classList.remove('active-tab');
    }
  });

  loadSheetData();
  recalculateAllFormulas();
  renderExcelTable();
  selectInitialCell();
}

function onMonthYearChange() {
  loadSheetData();
  recalculateAllFormulas();
  renderExcelTable();
  selectInitialCell();
  setupRealtimeBroadcastListener();
  syncAllCloudData(false);
}

function selectInitialCell() {
  const now = new Date();
  const rowsPerDay = getRowsPerDay();
  let targetRow = 6;

  if (MonthYearState.year === now.getFullYear() && MonthYearState.monthIndex === now.getMonth()) {
    targetRow = 6 + (now.getDate() - 1) * rowsPerDay;
  }

  const maxR = getMaxActiveRow();
  if (targetRow > maxR) targetRow = 6;
  selectCell('E', targetRow);
}

// ─── DATA GENERATION & PERSISTENCE ────────────────────────────────────────────
function generateBlankMonthRows(tabId, year, monthIndex) {
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const yr2Digit = String(year).slice(-2);
  const monthShort = MonthYearState.monthShortNames[monthIndex];
  const machines = getMachinesForTab(tabId);

  const rows = [];
  let rowCounter = 6;

  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(year, monthIndex, day);
    const dayName = MonthYearState.dayNames[d.getDay()];
    const dateStr = `${day}-${monthShort}-${yr2Digit}`;

    for (let mIdx = 0; mIdx < machines.length; mIdx++) {
      const rowObj = { row: rowCounter };
      EXCEL_COLUMNS.forEach(c => rowObj[c.col] = { val: null, formula: null });

      rowObj.A.val = dateStr;
      rowObj.B.val = dayName;
      rowObj.C.val = 'Morning';
      rowObj.D.val = machines[mIdx];

      rows.push(rowObj);
      rowCounter++;
    }
  }

  return rows;
}

function getStoredLocalData(tabId, year, monthIndex) {
  const key = getStorageKey(tabId, year, monthIndex);
  const saved = localStorage.getItem(key);
  if (!saved) return null;
  try {
    const parsed = JSON.parse(saved);
    if (Array.isArray(parsed)) {
      let hasInputs = false;
      for (const r of parsed) {
        if (r.E?.val || r.F?.val || r.G?.val || r.H?.val) {
          hasInputs = true;
          break;
        }
      }
      return { rows: parsed, updatedAt: hasInputs ? 1 : 0 };
    } else if (parsed && Array.isArray(parsed.rows)) {
      return parsed;
    }
  } catch (e) {
    console.error('Error parsing stored excel data:', e);
  }
  return null;
}

function loadSheetData() {
  const expectedRowsCount = getDaysInSelectedMonth() * getRowsPerDay(ACTIVE_TAB);
  const expectedMachines = getMachinesForTab(ACTIVE_TAB);
  const localData = getStoredLocalData(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);

  if (localData && Array.isArray(localData.rows)) {
    const parsed = localData.rows;
    if (parsed.length === expectedRowsCount) {
      let matches = true;
      const rowsPerDay = expectedMachines.length;
      for (let i = 0; i < parsed.length; i++) {
        const mIdx = i % rowsPerDay;
        if (parsed[i].D?.val !== expectedMachines[mIdx]) {
          matches = false;
          break;
        }
      }
      if (matches) {
        SheetState.rows = parsed;
        return;
      }
    }
  }

  // If August 2026 for Fan Lathe and initial template exists
  if (ACTIVE_TAB === 'fan_lathe' && MonthYearState.year === 2026 && MonthYearState.monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
    SheetState.rows = JSON.parse(JSON.stringify(INITIAL_EXCEL_ROWS));
  } else {
    SheetState.rows = generateBlankMonthRows(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
  }
}

function pushHistoryState() {
  if (SheetState.undoStack.length > 50) SheetState.undoStack.shift();
  SheetState.undoStack.push(JSON.stringify(SheetState.rows));
  SheetState.redoStack = [];
}

function undoAction() {
  if (SheetState.undoStack.length === 0) {
    showToast('Nothing to undo', 'info');
    return;
  }
  SheetState.redoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(SheetState.undoStack.pop());
  recalculateAllFormulas();
  renderExcelTable();
  saveSheetData(false);
  showToast('↩️ Undone', 'info');
}

function redoAction() {
  if (SheetState.redoStack.length === 0) {
    showToast('Nothing to redo', 'info');
    return;
  }
  SheetState.undoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(SheetState.redoStack.pop());
  recalculateAllFormulas();
  renderExcelTable();
  saveSheetData(false);
  showToast('↪️ Redone', 'info');
}

// ─── FIREBASE REALTIME DATABASE CONFIGURATION & SYNC ENGINE ───────────────────
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyBcjbR7Qu7M-RnHUtLJ9zeehILqQHYLw4E",
  authDomain: "whatsapp-c10ef.firebaseapp.com",
  databaseURL: "https://whatsapp-c10ef-default-rtdb.firebaseio.com",
  projectId: "whatsapp-c10ef",
  storageBucket: "whatsapp-c10ef.firebasestorage.app",
  messagingSenderId: "675053106773",
  appId: "1:675053106773:web:b7078468691a07ecfec6dc",
  measurementId: "G-89Z8WBJ3R0"
};

const FIREBASE_RTDB_BASE_URL = "https://whatsapp-c10ef-default-rtdb.firebaseio.com";
let firebaseDb = null;
let isFirebaseInitialized = false;
const cloudSyncDebounceTimers = {};
let lastBroadcastTimestamp = 0;

function cleanDataForFirebase(data) {
  if (data === undefined) return null;
  return JSON.parse(JSON.stringify(data, (key, value) => {
    return value === undefined ? null : value;
  }));
}

function initFirebase() {
  try {
    if (typeof firebase !== 'undefined') {
      if (!firebase.apps.length) {
        firebase.initializeApp(FIREBASE_CONFIG);
      }
      firebaseDb = firebase.database();
      isFirebaseInitialized = true;
      console.log('✅ Firebase Realtime Database SDK initialized: whatsapp-c10ef');
      updateCloudStatusUI('online', 'Cloud Live');

      if (firebase.auth) {
        firebase.auth().signInAnonymously().catch((err) => {
          console.log('Auth note:', err.message);
        });
      }

      firebaseDb.ref('.info/connected').on('value', (snap) => {
        if (snap.val() === true) {
          updateCloudStatusUI('online', 'Cloud Live');
        } else {
          updateCloudStatusUI('offline', 'Local Saved (Offline)');
        }
      });

      setupRealtimeBroadcastListener();
    } else {
      updateCloudStatusUI('online', 'Cloud Live (REST)');
    }

    // Initial background sync
    syncAllCloudData(false);
  } catch (e) {
    console.error('Firebase initialization error:', e);
    updateCloudStatusUI('online', 'Cloud Live (REST)');
  }
}

function updateCloudStatusUI(status, text) {
  const dot = document.getElementById('cloudStatusDot');
  const txt = document.getElementById('cloudStatusText');
  const autoSave = document.getElementById('autoSaveIndicator');

  if (dot) {
    if (status === 'online' || status === 'synced') {
      dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
      if (txt) txt.textContent = '🟢 LIVE';
      if (autoSave) autoSave.textContent = '🟢 Local + Cloud Live';
    } else if (status === 'syncing') {
      dot.className = 'w-2 h-2 rounded-full bg-amber-400 animate-ping';
      if (txt) txt.textContent = '⏳ SYNCING';
      if (autoSave) autoSave.textContent = '⏳ Syncing Cloud...';
    } else {
      dot.className = 'w-2 h-2 rounded-full bg-slate-400';
      if (txt) txt.textContent = '⚫ OFFLINE';
      if (autoSave) autoSave.textContent = '💾 Local Saved';
    }
  }
}

// ─── INCHARGE & USER: INSTANT DUAL-SAVE (REST API + SDK) TO CLOUD ─────────────
function pushSheetDataToCloud(tabId, year, monthIndex, rowsData, timestamp = Date.now()) {
  if (!tabId || SHEET_TABS[tabId]?.isSummary) return;

  const payload = {
    rows: cleanDataForFirebase(rowsData),
    tabId: tabId,
    year: year,
    monthIndex: monthIndex,
    updatedAt: timestamp,
    updatedBy: CurrentUser ? CurrentUser.name : 'Incharge',
    userId: CurrentUser ? CurrentUser.id : 'unknown'
  };

  // 1. Firebase SDK Write (Instant WebSockets)
  if (firebaseDb) {
    firebaseDb.ref(`mep_oee_v2/data/${year}/${monthIndex}/${tabId}`).set(payload).then(() => {
      updateCloudStatusUI('synced', 'Local + Cloud Saved');
    }).catch(() => {});
  }

  // 2. Guaranteed REST API Write (Standard HTTP PUT without 64KB quota issue)
  const restUrl1 = `${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/data/${year}/${monthIndex}/${tabId}.json`;
  fetch(restUrl1, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  }).then(() => {
    updateCloudStatusUI('synced', 'Local + Cloud Saved');
  }).catch(() => {});
}

// ─── ADMIN: CONSOLIDATE & BROADCAST ALL DATA LIVE ─────────────────────────────
async function adminBroadcastLiveData() {
  const spinIcon = document.getElementById('syncSpinIcon');
  if (spinIcon) spinIcon.classList.add('animate-spin');
  updateCloudStatusUI('syncing', 'Broadcasting Live...');

  try {
    const year = MonthYearState.year;
    const monthIndex = MonthYearState.monthIndex;

    // 1. Save current active tab locally first
    saveSheetData(false);

    // 2. Fetch existing cloud departments from Firebase REST API (with SDK fallback)
    let cloudDepts = {};
    try {
      if (firebaseDb) {
        const snap = await firebaseDb.ref(`mep_oee_v2/data/${year}/${monthIndex}`).once('value');
        cloudDepts = snap.val() || {};
      }
    } catch(e) {}

    if (!cloudDepts || Object.keys(cloudDepts).length === 0) {
      try {
        const restUrl = `${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/data/${year}/${monthIndex}.json`;
        const res = await fetch(restUrl);
        if (res.ok) {
          cloudDepts = (await res.json()) || {};
        }
      } catch(e) {}
    }

    const consolidated = {};

    // 3. For all 9 production departments: build complete dataset
    Object.keys(SHEET_TABS).forEach(tabId => {
      if (!SHEET_TABS[tabId].isSummary) {
        if (tabId === ACTIVE_TAB && SheetState.rows && SheetState.rows.length > 0) {
          consolidated[tabId] = {
            rows: cleanDataForFirebase(SheetState.rows),
            tabId: tabId,
            year: year,
            monthIndex: monthIndex,
            updatedAt: Date.now(),
            updatedBy: CurrentUser ? CurrentUser.name : 'Admin'
          };
          return;
        }

        const cloudItem = cloudDepts[tabId];
        const localData = getStoredLocalData(tabId, year, monthIndex);
        const localUpdated = localData ? (localData.updatedAt || 0) : 0;
        const cloudUpdated = cloudItem ? (cloudItem.updatedAt || 0) : 0;

        if (localData && Array.isArray(localData.rows) && localData.rows.length > 0 && localUpdated >= cloudUpdated) {
          consolidated[tabId] = {
            rows: cleanDataForFirebase(localData.rows),
            tabId: tabId,
            year: year,
            monthIndex: monthIndex,
            updatedAt: localUpdated || Date.now(),
            updatedBy: CurrentUser ? CurrentUser.name : 'Admin'
          };
        } else if (cloudItem && Array.isArray(cloudItem.rows) && cloudItem.rows.length > 0) {
          consolidated[tabId] = cloudItem;
          const storageKey = getStorageKey(tabId, year, monthIndex);
          localStorage.setItem(storageKey, JSON.stringify(cloudItem));
        } else if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
          consolidated[tabId] = {
            rows: cleanDataForFirebase(localData.rows),
            tabId: tabId,
            year: year,
            monthIndex: monthIndex,
            updatedAt: Date.now(),
            updatedBy: CurrentUser ? CurrentUser.name : 'Admin'
          };
        } else {
          const rows = (tabId === 'fan_lathe' && year === 2026 && monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined')
            ? INITIAL_EXCEL_ROWS
            : generateBlankMonthRows(tabId, year, monthIndex);
          consolidated[tabId] = {
            rows: cleanDataForFirebase(rows),
            tabId: tabId,
            year: year,
            monthIndex: monthIndex,
            updatedAt: Date.now(),
            updatedBy: CurrentUser ? CurrentUser.name : 'Admin'
          };
        }
      }
    });

    const deptKeys = Object.keys(consolidated);

    // 4. Publish Master Live Broadcast via Firebase SDK (Instant WebSockets)
    const cleanedConsolidated = cleanDataForFirebase(consolidated);
    if (firebaseDb) {
      await firebaseDb.ref(`mep_oee_v2/live_broadcast/${year}/${monthIndex}`).set(cleanedConsolidated);
    }

    // 5. REST PUT Backup
    try {
      const liveBroadcastUrl = `${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/live_broadcast/${year}/${monthIndex}.json`;
      await fetch(liveBroadcastUrl, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(cleanedConsolidated)
      });
    } catch(e) {}

    // 6. Fire Global Broadcast Signal
    const timestamp = Date.now();
    const signalPayload = {
      timestamp: timestamp,
      publishedBy: CurrentUser ? CurrentUser.name : 'Admin',
      year: year,
      monthIndex: monthIndex
    };

    if (firebaseDb) {
      await firebaseDb.ref('mep_oee_v2/broadcast_signal').set(signalPayload);
    }

    try {
      await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/broadcast_signal.json`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(signalPayload)
      });
    } catch(e) {}

    // 7. Recalculate & re-render admin view
    loadSheetData();
    recalculateAllFormulas();
    renderExcelTable();
    updateCloudStatusUI('synced', 'Live Broadcasted');

    showToast(`🎉 সফলভাবে ${deptKeys.length}টি ডিপার্টমেন্টের ডাটা লাইভ ব্রডকাস্ট করা হয়েছে! সকল ডিভাইসে লাইভ আপডেট সম্পন্ন।`, 'success');
  } catch (err) {
    console.error('Error broadcasting live data:', err);
    showToast('❌ লাইভ ব্রডকাস্ট করতে সমস্যা হয়েছে: ' + (err.message || ''), 'error');
    updateCloudStatusUI('error', 'Broadcast Failed');
  } finally {
    if (spinIcon) spinIcon.classList.remove('animate-spin');
  }
}

// ─── CLIENTS: REALTIME BROADCAST & INSTANT WEBSOCKET LISTENERS ───────────────
let activeBroadcastRef = null;

function setupRealtimeBroadcastListener() {
  const year = MonthYearState.year;
  const monthIndex = MonthYearState.monthIndex;

  if (firebaseDb) {
    // 1. Instant Global Broadcast Signal Listener via WebSocket (Triggered ONLY on Admin Publish)
    firebaseDb.ref('mep_oee_v2/broadcast_signal').off();
    firebaseDb.ref('mep_oee_v2/broadcast_signal').on('value', async (snap) => {
      const signal = snap.val();
      if (signal) processIncomingBroadcastSignal(signal);
    });

    // 2. Detach previous month/year live broadcast listeners
    if (activeBroadcastRef) activeBroadcastRef.off();

    // 3. Instant Realtime Live Broadcast WebSocket Listener (Receives ONLY Admin-Published Snapshots)
    activeBroadcastRef = firebaseDb.ref(`mep_oee_v2/live_broadcast/${year}/${monthIndex}`);
    activeBroadcastRef.on('value', (snap) => {
      const liveData = snap.val();
      if (liveData && typeof liveData === 'object') {
        applyIncomingBroadcastData(liveData, false);
      }
    });
  }

  // Periodic fallback check every 3 seconds for non-WebSocket connections
  if (!window._mepSyncInterval) {
    window._mepSyncInterval = setInterval(async () => {
      try {
        const res = await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/broadcast_signal.json`);
        if (res.ok) {
          const signal = await res.json();
          if (signal) processIncomingBroadcastSignal(signal);
        }
      } catch(e) {}
    }, 3000);
  }
}

function applyIncomingBroadcastData(liveData, isSilent = true) {
  if (!liveData || typeof liveData !== 'object') return;
  const year = MonthYearState.year;
  const monthIndex = MonthYearState.monthIndex;
  let validCount = 0;

  Object.keys(liveData).forEach(tabId => {
    const cloudItem = liveData[tabId];
    if (cloudItem && Array.isArray(cloudItem.rows) && cloudItem.rows.length > 0) {
      // Incharge local draft protection: if user is working on their permitted tab and has newer draft, keep draft
      if (CurrentUser && CurrentUser.id !== 'admin' && CurrentUser.allowedDepts && CurrentUser.allowedDepts.includes(tabId)) {
        const localData = getStoredLocalData(tabId, year, monthIndex);
        const localUpdated = localData ? (localData.updatedAt || 0) : 0;
        const cloudUpdated = cloudItem.updatedAt || 0;
        if (localData && localUpdated > cloudUpdated && localUpdated > 0) {
          return;
        }
      }

      const storageKey = getStorageKey(tabId, year, monthIndex);
      localStorage.setItem(storageKey, JSON.stringify(cloudItem));
      validCount++;
    }
  });

  if (validCount > 0) {
    if (!SheetState.isEditing) {
      loadSheetData();
      recalculateAllFormulas();
      renderExcelTable();
    }
    updateCloudStatusUI('synced', 'Live Synced');
  }
}

async function processIncomingBroadcastSignal(signal) {
  if (!signal || !signal.timestamp) return;

  if (signal.timestamp > lastBroadcastTimestamp) {
    const isInitial = (lastBroadcastTimestamp === 0);
    lastBroadcastTimestamp = signal.timestamp;

    const year = MonthYearState.year;
    const monthIndex = MonthYearState.monthIndex;

    if (signal.year === year && signal.monthIndex === monthIndex) {
      try {
        let liveData = null;
        if (firebaseDb) {
          const snap = await firebaseDb.ref(`mep_oee_v2/live_broadcast/${year}/${monthIndex}`).once('value');
          liveData = snap.val();
        }
        if (!liveData) {
          try {
            const res = await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/live_broadcast/${year}/${monthIndex}.json`);
            if (res.ok) liveData = await res.json();
          } catch(e) {}
        }

        if (liveData && typeof liveData === 'object' && Object.keys(liveData).length > 0) {
          applyIncomingBroadcastData(liveData);

          if (!isInitial && CurrentUser && CurrentUser.id !== 'admin') {
            showToast('📢 লাইভ ডাটা পাবলিশ হয়েছে! স্ক্রিনে সর্বশেষ ডাটা চলে এসেছে।', 'info');
          }
        }
      } catch(e) {
        console.error('Error processing broadcast signal:', e);
      }
    }
  }
}

// ─── MASTER SYNC / PUBLISH ACTION DISPATCHER ─────────────────────────────────
function handleSyncOrPublish() {
  if (CurrentUser && CurrentUser.id === 'admin') {
    adminBroadcastLiveData();
  } else {
    syncAllCloudData(true);
  }
}

async function syncAllCloudData(showFeedback = true) {
  const spinIcon = document.getElementById('syncSpinIcon');
  if (spinIcon) spinIcon.classList.add('animate-spin');
  updateCloudStatusUI('syncing', 'Syncing Data...');

  try {
    const year = MonthYearState.year;
    const monthIndex = MonthYearState.monthIndex;

    // 1. Fetch live broadcast first
    let monthData = null;
    try {
      const res = await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/live_broadcast/${year}/${monthIndex}.json`);
      if (res.ok) monthData = await res.json();
    } catch(e) {}

    // 2. Fallback to raw depts data
    if (!monthData || Object.keys(monthData).length === 0) {
      try {
        const res = await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/data/${year}/${monthIndex}.json`);
        if (res.ok) monthData = await res.json();
      } catch(e) {}
    }

    let updatedCount = 0;

    if (monthData && typeof monthData === 'object') {
      Object.keys(monthData).forEach(tabId => {
        const cloudItem = monthData[tabId];
        if (cloudItem && Array.isArray(cloudItem.rows) && cloudItem.rows.length > 0) {
          const localData = getStoredLocalData(tabId, year, monthIndex);
          const localUpdated = localData ? (localData.updatedAt || 0) : 0;
          const cloudUpdated = cloudItem.updatedAt || 0;

          // CRITICAL: If local data is newer or has unsaved edits, PRESERVE IT and push to cloud!
          if (localData && localUpdated >= cloudUpdated && localUpdated > 0) {
            pushSheetDataToCloud(tabId, year, monthIndex, localData.rows, localUpdated);
            return;
          }

          // Otherwise update local storage with cloud data
          const storageKey = getStorageKey(tabId, year, monthIndex);
          localStorage.setItem(storageKey, JSON.stringify(cloudItem));
          updatedCount++;
        }
      });
    }

    if (updatedCount > 0) {
      loadSheetData();
      recalculateAllFormulas();
      renderExcelTable();
      updateCloudStatusUI('synced', 'Cloud Synced');
    } else {
      updateCloudStatusUI('synced', 'Local + Cloud Synced');
    }

    if (showFeedback) {
      if (updatedCount > 0) {
        showToast(`☁️ ফায়ারবেস ক্লাউড থেকে সর্বশেষ ডাটা সফলভাবে সিঙ্ক হয়েছে! (${updatedCount}টি শিট)`, 'success');
      } else {
        showToast('ℹ️ লোকাল ডাটা সুরক্ষিত রয়েছে।', 'info');
      }
    }
  } catch (err) {
    console.error('Error syncing cloud data:', err);
    if (showFeedback) showToast('❌ ক্লাউড ডাটা সিঙ্ক করতে সমস্যা হয়েছে।', 'error');
    updateCloudStatusUI('error', 'Sync Failed');
  } finally {
    if (spinIcon) spinIcon.classList.remove('animate-spin');
  }
}

function saveSheetData(pushHistory = true) {
  if (pushHistory) pushHistoryState();
  try {
    const now = Date.now();
    const key = getStorageKey(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
    const dataObj = {
      rows: SheetState.rows,
      updatedAt: now,
      tabId: ACTIVE_TAB,
      year: MonthYearState.year,
      monthIndex: MonthYearState.monthIndex,
      updatedBy: CurrentUser ? CurrentUser.name : 'User'
    };
    localStorage.setItem(key, JSON.stringify(dataObj));
    triggerSaveIndicator();

    // Instant Non-blocking Dual-Save to Firebase Cloud with keepalive
    pushSheetDataToCloud(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, SheetState.rows, now);
  } catch (e) {
    console.error('Failed to save to localStorage:', e);
  }
}

function triggerSaveIndicator() {
  const el = document.getElementById('autoSaveIndicator');
  if (el) {
    el.textContent = '🟢 Saved';
    el.style.opacity = '1';
  }
}

function resetToOriginalData() {
  const tabName = SHEET_TABS[ACTIVE_TAB].name;
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  if (confirm(`Reset entire sheet for [${tabName}] (${monthName} ${MonthYearState.year})?`)) {
    const key = getStorageKey(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
    localStorage.removeItem(key);
    loadSheetData();
    recalculateAllFormulas();
    renderExcelTable();
    showToast(`🔄 Reset [${tabName}] for ${monthName} ${MonthYearState.year}`, 'info');
  }
}

// ─── AUTO-CALCULATION ENGINE ──────────────────────────────────────────────────
function recalculateRow(rowObj) {
  const r = rowObj.row;
  const rowsPerDay = getRowsPerDay();
  const mIdx = (r - 6) % rowsPerDay;
  const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);

  const cap = Number(rowObj.E?.val) || 0;
  const act = Number(rowObj.F?.val) || 0;
  const rej = Number(rowObj.G?.val) || 0;
  let plan = Number(rowObj.H?.val) || 0;

  // I: Expected DownTime = IF(E<>"", 30, 0)
  const hasCap = (rowObj.E?.val !== null && rowObj.E?.val !== '' && rowObj.E?.val !== undefined && rowObj.E?.val != 0);
  const expDt = hasCap ? 30 : 0;
  if (!rowObj.I) rowObj.I = {};
  rowObj.I.val = expDt;
  rowObj.I.formula = `=IF(E${r}<>"",30,0)`;

  // AH: Total Down Time = SUM(K:AG)
  let sumDt = 0;
  let hasAnyDt = false;
  EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => {
    const val = Number(rowObj[c.col]?.val) || 0;
    if (rowObj[c.col]?.val !== null && rowObj[c.col]?.val !== '' && rowObj[c.col]?.val !== undefined) {
      hasAnyDt = true;
    }
    sumDt += val;
  });
  if (!rowObj.AH) rowObj.AH = {};
  rowObj.AH.val = sumDt;
  rowObj.AH.formula = `=SUM(K${r}:AG${r})`;

  // J: Total Prd. Run Time = H - AH
  if (!rowObj.J) rowObj.J = {};
  if (plan > 0 || hasAnyDt) {
    rowObj.J.val = Math.max(0, plan - sumDt);
  } else {
    rowObj.J.val = '-';
  }
  rowObj.J.formula = `=H${r}-AH${r}`;

  const numRunTime = typeof rowObj.J.val === 'number' ? rowObj.J.val : 0;

  // AI: Availability (%) = IFERROR(J/H, "0")
  let avail = plan > 0 ? (numRunTime / plan) : 0;
  if (groupInfo.isSlave) {
    const dayStartRow = r - mIdx;
    const masterRow = SheetState.rows.find(row => row.row === (dayStartRow + groupInfo.masterIdx));
    if (masterRow && masterRow.AI && typeof masterRow.AI.val === 'number') {
      avail = masterRow.AI.val;
    }
  }

  if (!rowObj.AI) rowObj.AI = {};
  rowObj.AI.val = avail;
  rowObj.AI.formula = `=IFERROR(J${r}/H${r},"0")`;

  // AJ: Performance (%) = IFERROR(F/E, "0")
  const perf = cap > 0 ? (act / cap) : 0;
  if (!rowObj.AJ) rowObj.AJ = {};
  rowObj.AJ.val = perf;
  rowObj.AJ.formula = `=IFERROR(F${r}/E${r},"0")`;

  // AK: Quality (%) = IFERROR(F/(F+G), "0")
  const totalPrd = act + rej;
  const qual = totalPrd > 0 ? (act / totalPrd) : 0;
  if (!rowObj.AK) rowObj.AK = {};
  rowObj.AK.val = qual;
  rowObj.AK.formula = `=IFERROR(F${r}/(F${r}+G${r}),"0")`;

  // AL: OEE (%) = IFERROR(AK*AJ*AI, "0")
  const oee = avail * perf * qual;
  if (!rowObj.AL) rowObj.AL = {};
  rowObj.AL.val = oee;
  rowObj.AL.formula = `=IFERROR(AK${r}*AJ${r}*AI${r},"0")`;

  // If this is master row, propagate availability to slave rows
  if (groupInfo.isMaster && groupInfo.count > 1) {
    for (let offset = 1; offset < groupInfo.count; offset++) {
      const slaveRow = SheetState.rows.find(row => row.row === (r + offset));
      if (slaveRow) {
        if (!slaveRow.AI) slaveRow.AI = {};
        slaveRow.AI.val = avail;
        const sPerf = slaveRow.AJ?.val || 0;
        const sQual = slaveRow.AK?.val || 0;
        if (!slaveRow.AL) slaveRow.AL = {};
        slaveRow.AL.val = avail * sPerf * sQual;
      }
    }
  }
}

function recalculateAllFormulas() {
  SheetState.rows.forEach(r => recalculateRow(r));
  recalculateTotalRow();
}

function recalculateTotalRow() {
  const totals = { E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, AH: 0 };
  EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => totals[c.col] = 0);
  const rowsPerDay = getRowsPerDay();

  SheetState.rows.forEach(r => {
    // Piece-count columns are summed across all rows
    totals.E += Number(r.E?.val) || 0;
    totals.F += Number(r.F?.val) || 0;
    totals.G += Number(r.G?.val) || 0;

    const mIdx = (r.row - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);

    // Time-related columns are summed ONLY from master rows (no duplication!)
    if (groupInfo.isMaster) {
      totals.H += Number(r.H?.val) || 0;
      totals.I += Number(r.I?.val) || 0;
      totals.J += (typeof r.J?.val === 'number' ? r.J.val : 0);
      totals.AH += Number(r.AH?.val) || 0;

      EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => {
        totals[c.col] += Number(r[c.col]?.val) || 0;
      });
    }
  });

  totals.AI = totals.H > 0 ? (totals.J / totals.H) : 0;
  totals.AJ = totals.E > 0 ? (totals.F / totals.E) : 0;
  totals.AK = (totals.F + totals.G) > 0 ? (totals.F / (totals.F + totals.G)) : 0;
  totals.AL = totals.AI * totals.AJ * totals.AK;

  SheetState.totals = totals;
}

// ─── TAB MONTHLY SUMMARY CALCULATOR (FOR EXECUTIVE DOWNTIME REPORT) ───────────
function getTabMonthlySummary(tabId, year, monthIndex) {
  let rowsData = [];
  if (tabId === ACTIVE_TAB && monthIndex === MonthYearState.monthIndex && year === MonthYearState.year) {
    rowsData = SheetState.rows;
  } else {
    const localData = getStoredLocalData(tabId, year, monthIndex);
    if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
      rowsData = localData.rows;
    } else if (tabId === 'fan_lathe' && year === 2026 && monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
      rowsData = INITIAL_EXCEL_ROWS;
    } else {
      rowsData = generateBlankMonthRows(tabId, year, monthIndex);
    }
  }

  const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);
  const dtSums = {};
  dtCols.forEach(c => dtSums[c.col] = 0);

  let totalPlannedMins = 0;
  let totalDownMins = 0;
  const rowsPerDay = getRowsPerDay(tabId);

  rowsData.forEach(r => {
    const mIdx = (r.row - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(tabId, mIdx);

    if (groupInfo.isMaster) {
      const planned = Number(r.H?.val) || 0;
      totalPlannedMins += planned;

      let rowDt = 0;
      dtCols.forEach(c => {
        const v = Number(r[c.col]?.val) || 0;
        dtSums[c.col] += v;
        rowDt += v;
      });
      totalDownMins += (typeof r.AH?.val === 'number' && r.AH.val > 0 ? r.AH.val : rowDt);
    }
  });

  const totalRunMins = totalPlannedMins - totalDownMins;
  const dtPercent = totalPlannedMins > 0 ? (totalDownMins / totalPlannedMins) : 0;

  return {
    tabId,
    name: SHEET_TABS[tabId]?.name || tabId,
    dtSums,
    totalDownMins,
    totalPlannedMins,
    totalRunMins,
    dtPercent
  };
}

// ─── EXECUTIVE TOTAL DOWNTIME SUMMARY TABLE RENDERER ──────────────────────────
function renderSummaryDowntimeReport(table) {
  const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);
  const totalColsCount = 1 + dtCols.length + 4; // Machine No + 23 DT cols + 4 summary cols
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;

  // 1. Column Widths (% based so table spans 100% full screen width)
  const colgroup = document.createElement('colgroup');
  
  // Machine No column (16%)
  const col0 = document.createElement('col');
  col0.style.width = '16%';
  colgroup.appendChild(col0);

  // 23 DT columns (2.95% each = 67.85%)
  dtCols.forEach(() => {
    const col = document.createElement('col');
    col.style.width = '2.95%';
    colgroup.appendChild(col);
  });

  // 4 Summary columns (4.03% each = 16.15%)
  const summaryWidthPcts = ['3.8%', '4.75%', '3.8%', '3.8%'];
  for (let i = 0; i < 4; i++) {
    const col = document.createElement('col');
    col.style.width = summaryWidthPcts[i];
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);

  // 2. Thead
  const thead = document.createElement('thead');

  // Row 1: MEP FAN LTD.
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = totalColsCount;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: Total Downtime Report (Month Year)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = totalColsCount;
  th2.textContent = `Total Downtime Report (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Vertical headers row (Clean and readable without any number row)
  const trHeaders = document.createElement('tr');

  const thMachineNo = document.createElement('th');
  thMachineNo.className = 'text-center align-middle font-bold text-xs border border-slate-600 bg-[#7EC8E3] text-black px-2';
  thMachineNo.textContent = 'Machine No.';
  thMachineNo.style.verticalAlign = 'middle';
  trHeaders.appendChild(thMachineNo);

  dtCols.forEach(c => {
    const th = document.createElement('th');
    th.className = 'v-header-cell zone-purple';
    const span = document.createElement('span');
    span.className = 'v-header-text';
    span.textContent = c.label;
    th.appendChild(span);
    trHeaders.appendChild(th);
  });

  const summaryHeaders = [
    'Total Down Time\n(Mins)',
    'Production Running\nTime (Mins)',
    'Total Run Time\n(Mins)',
    'Total Down Time\n(%)'
  ];

  summaryHeaders.forEach((lbl, idx) => {
    const th = document.createElement('th');
    th.className = 'v-header-cell ' + (idx < 2 ? 'zone-purple' : 'zone-cyan');
    const span = document.createElement('span');
    span.className = 'v-header-text';
    span.innerHTML = lbl.replace('\n', '<br>');
    th.appendChild(span);
    trHeaders.appendChild(th);
  });

  thead.appendChild(trHeaders);
  table.appendChild(thead);

  // 3. Tbody
  const tbody = document.createElement('tbody');

  // Accumulators for Total Row
  const grandDtSums = {};
  dtCols.forEach(c => grandDtSums[c.col] = 0);
  let grandTotalDown = 0;
  let grandPlanned = 0;
  let grandRunTime = 0;

  getActiveSummaryDepts().forEach(dept => {
    const tabSummary = getTabMonthlySummary(dept.id, year, MonthYearState.monthIndex);

    // Accumulate
    dtCols.forEach(c => grandDtSums[c.col] += tabSummary.dtSums[c.col]);
    grandTotalDown += tabSummary.totalDownMins;
    grandPlanned += tabSummary.totalPlannedMins;
    grandRunTime += tabSummary.totalRunMins;

    const tr = document.createElement('tr');

    // Col 1: Department Name
    const tdDept = document.createElement('td');
    tdDept.className = 'summary-dept-cell';
    tdDept.textContent = dept.name;
    tr.appendChild(tdDept);

    // Downtime columns
    dtCols.forEach(c => {
      const td = document.createElement('td');
      td.className = 'summary-data-cell';
      const val = tabSummary.dtSums[c.col];
      td.textContent = (val > 0 ? val.toLocaleString() : '-');
      tr.appendChild(td);
    });

    // Summary columns
    const tdDown = document.createElement('td');
    tdDown.className = 'summary-highlight-cell';
    tdDown.textContent = (tabSummary.totalDownMins > 0 ? tabSummary.totalDownMins.toLocaleString() : '-');
    tr.appendChild(tdDown);

    const tdPlanned = document.createElement('td');
    tdPlanned.className = 'summary-data-cell';
    tdPlanned.textContent = (tabSummary.totalPlannedMins > 0 ? tabSummary.totalPlannedMins.toLocaleString() : '-');
    tr.appendChild(tdPlanned);

    const tdRun = document.createElement('td');
    tdRun.className = 'summary-data-cell font-bold';
    tdRun.textContent = tabSummary.totalRunMins.toLocaleString();
    tr.appendChild(tdRun);

    const tdPct = document.createElement('td');
    tdPct.className = 'summary-kpi-cell font-bold';
    tdPct.textContent = `${(tabSummary.dtPercent * 100).toFixed(2)}%`;
    tr.appendChild(tdPct);

    tbody.appendChild(tr);
  });

  // 4. Total (Mins) Row
  const trTotal = document.createElement('tr');
  trTotal.className = 'summary-total-row';

  const thTotTitle = document.createElement('th');
  thTotTitle.className = 'summary-total-title';
  thTotTitle.textContent = 'Total (Mins)';
  trTotal.appendChild(thTotTitle);

  dtCols.forEach(c => {
    const td = document.createElement('td');
    const v = grandDtSums[c.col];
    td.textContent = (v > 0 ? v.toLocaleString() : '-');
    trTotal.appendChild(td);
  });

  const tdTotDown = document.createElement('td');
  tdTotDown.textContent = grandTotalDown.toLocaleString();
  trTotal.appendChild(tdTotDown);

  const tdTotPlanned = document.createElement('td');
  tdTotPlanned.textContent = grandPlanned.toLocaleString();
  trTotal.appendChild(tdTotPlanned);

  const tdTotRun = document.createElement('td');
  tdTotRun.textContent = grandRunTime.toLocaleString();
  trTotal.appendChild(tdTotRun);

  const tdTotPct = document.createElement('td');
  const grandPct = grandPlanned > 0 ? (grandTotalDown / grandPlanned) * 100 : 0;
  tdTotPct.textContent = `${grandPct.toFixed(2)}%`;
  trTotal.appendChild(tdTotPct);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ─── TAB PRODUCTION OUTPUT SUMMARY CALCULATOR ─────────────────────────────────
function getTabProductionOutputSummary(tabId, year, monthIndex) {
  let rowsData = [];
  if (tabId === ACTIVE_TAB && monthIndex === MonthYearState.monthIndex && year === MonthYearState.year) {
    rowsData = SheetState.rows;
  } else {
    const localData = getStoredLocalData(tabId, year, monthIndex);
    if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
      rowsData = localData.rows;
    } else if (tabId === 'fan_lathe' && year === 2026 && monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
      rowsData = INITIAL_EXCEL_ROWS;
    } else {
      rowsData = generateBlankMonthRows(tabId, year, monthIndex);
    }
  }

  let runningMins = 0;
  let capacityPcs = 0;
  let actualPrdPcs = 0;
  let rejectionPcs = 0;
  const rowsPerDay = getRowsPerDay(tabId);

  rowsData.forEach(r => {
    const mIdx = (r.row - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(tabId, mIdx);

    if (groupInfo.isMaster) {
      runningMins += Number(r.H?.val) || 0;
    }
    capacityPcs += Number(r.E?.val) || 0;
    actualPrdPcs += Number(r.F?.val) || 0;
    rejectionPcs += Number(r.G?.val) || 0;
  });

  const outputPct = capacityPcs > 0 ? (actualPrdPcs / capacityPcs) : 0;

  return {
    tabId,
    name: SHEET_TABS[tabId]?.name || tabId,
    runningMins,
    capacityPcs,
    actualPrdPcs,
    rejectionPcs,
    outputPct
  };
}

// ─── EXECUTIVE PRODUCTION OUTPUT REPORT TABLE RENDERER ────────────────────────
function renderProductionOutputReport(table) {
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;

  // Column Widths
  const colgroup = document.createElement('colgroup');
  const widths = ['22%', '13%', '13%', '13%', '13%', '13%', '13%'];
  widths.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // Thead
  const thead = document.createElement('thead');

  // Row 1: MEP FAN LTD.
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = 7;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: Production Output (Month Year)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = 7;
  th2.textContent = `Production Output (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Header Row 1 (Merged top headers)
  const tr3 = document.createElement('tr');

  const thSec = document.createElement('th');
  thSec.rowSpan = 2;
  thSec.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thSec.textContent = 'Section Name';
  thSec.style.verticalAlign = 'middle';
  tr3.appendChild(thSec);

  const thRun = document.createElement('th');
  thRun.rowSpan = 2;
  thRun.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thRun.innerHTML = 'Production<br>Running (hr)';
  thRun.style.verticalAlign = 'middle';
  tr3.appendChild(thRun);

  const thCap = document.createElement('th');
  thCap.rowSpan = 2;
  thCap.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thCap.innerHTML = 'Machine<br>Capacity (Pcs)';
  thCap.style.verticalAlign = 'middle';
  tr3.appendChild(thCap);

  const thQty = document.createElement('th');
  thQty.colSpan = 2;
  thQty.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1';
  thQty.textContent = 'Prodcution Qty.';
  thQty.style.verticalAlign = 'middle';
  tr3.appendChild(thQty);

  const thOut = document.createElement('th');
  thOut.rowSpan = 2;
  thOut.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thOut.innerHTML = 'Standrad Wise<br>Production Output<br>(%)';
  thOut.style.verticalAlign = 'middle';
  tr3.appendChild(thOut);

  const thRem = document.createElement('th');
  thRem.rowSpan = 2;
  thRem.className = 'prod-output-remarks-header border border-red-800';
  thRem.textContent = 'Remarks';
  thRem.style.verticalAlign = 'middle';
  tr3.appendChild(thRem);

  thead.appendChild(tr3);

  // Row 4: Sub-headers for Production (pcs) and Rejection (Pcs)
  const tr4 = document.createElement('tr');

  const thProdPcs = document.createElement('th');
  thProdPcs.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1.5';
  thProdPcs.textContent = 'Production (pcs)';
  thProdPcs.style.verticalAlign = 'middle';
  tr4.appendChild(thProdPcs);

  const thRejPcs = document.createElement('th');
  thRejPcs.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1.5';
  thRejPcs.textContent = 'Rejection (Pcs)';
  thRejPcs.style.verticalAlign = 'middle';
  tr4.appendChild(thRejPcs);

  thead.appendChild(tr4);
  table.appendChild(thead);

  // Tbody
  const tbody = document.createElement('tbody');

  let grandRunning = 0;
  let grandCapacity = 0;
  let grandProduction = 0;
  let grandRejection = 0;

  getActiveSummaryDepts().forEach(dept => {
    const tabSummary = getTabProductionOutputSummary(dept.id, year, MonthYearState.monthIndex);

    grandRunning += tabSummary.runningMins;
    grandCapacity += tabSummary.capacityPcs;
    grandProduction += tabSummary.actualPrdPcs;
    grandRejection += tabSummary.rejectionPcs;

    const tr = document.createElement('tr');

    // Col 1: Section Name
    const tdSec = document.createElement('td');
    tdSec.className = 'prod-output-section-cell';
    tdSec.textContent = dept.name;
    tr.appendChild(tdSec);

    // Col 2: Production Running (hr)
    const tdRun = document.createElement('td');
    tdRun.className = 'summary-data-cell';
    tdRun.textContent = tabSummary.runningMins > 0 ? tabSummary.runningMins.toLocaleString() : '-';
    tr.appendChild(tdRun);

    // Col 3: Machine Capacity (Pcs)
    const tdCap = document.createElement('td');
    tdCap.className = 'summary-data-cell';
    tdCap.textContent = tabSummary.capacityPcs > 0 ? tabSummary.capacityPcs.toLocaleString() : '-';
    tr.appendChild(tdCap);

    // Col 4: Production (pcs)
    const tdProd = document.createElement('td');
    tdProd.className = 'summary-data-cell';
    tdProd.textContent = tabSummary.actualPrdPcs > 0 ? tabSummary.actualPrdPcs.toLocaleString() : '-';
    tr.appendChild(tdProd);

    // Col 5: Rejection (Pcs)
    const tdRej = document.createElement('td');
    tdRej.className = 'summary-data-cell';
    tdRej.textContent = tabSummary.rejectionPcs > 0 ? tabSummary.rejectionPcs.toLocaleString() : '0';
    tr.appendChild(tdRej);

    // Col 6: Standrad Wise Production Output (%)
    const tdPct = document.createElement('td');
    tdPct.className = 'prod-output-kpi-cell';
    tdPct.textContent = `${(tabSummary.outputPct * 100).toFixed(2)}%`;
    tr.appendChild(tdPct);

    // Col 7: Remarks
    const tdRem = document.createElement('td');
    tdRem.className = 'summary-data-cell';
    tdRem.textContent = '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row (Yellow/Gold)
  const trTotal = document.createElement('tr');

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'prod-output-total-title';
  tdTotTitle.textContent = 'Total';
  trTotal.appendChild(tdTotTitle);

  const tdTotRun = document.createElement('td');
  tdTotRun.className = 'prod-output-total-cell';
  tdTotRun.textContent = grandRunning.toLocaleString();
  trTotal.appendChild(tdTotRun);

  const tdTotCap = document.createElement('td');
  tdTotCap.className = 'prod-output-total-cell';
  tdTotCap.textContent = grandCapacity.toLocaleString();
  trTotal.appendChild(tdTotCap);

  const tdTotProd = document.createElement('td');
  tdTotProd.className = 'prod-output-total-cell';
  tdTotProd.textContent = grandProduction.toLocaleString();
  trTotal.appendChild(tdTotProd);

  const tdTotRej = document.createElement('td');
  tdTotRej.className = 'prod-output-total-cell';
  tdTotRej.textContent = grandRejection.toLocaleString();
  trTotal.appendChild(tdTotRej);

  const tdTotPct = document.createElement('td');
  tdTotPct.className = 'prod-output-total-cell';
  const grandOutputPct = grandCapacity > 0 ? (grandProduction / grandCapacity) * 100 : 0;
  tdTotPct.textContent = `${grandOutputPct.toFixed(2)}%`;
  trTotal.appendChild(tdTotPct);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'prod-output-total-cell';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ─── EXECUTIVE PRODUCTION OUTPUT SUMMARY TABLE RENDERER ───────────────────────
function renderSummaryProductionReport(table) {
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;
  const headers = [
    'Machine No.',
    'Total Running (Hrs)',
    'Total Capacity (Pcs)',
    'Target / Hr (Pcs)',
    'Actual Prd. (Pcs)',
    'Actual Prd. / Hr',
    'Total Achievement (%)',
    'Total Rejection (Pcs)'
  ];

  // 1. Colgroup with exact % widths to guarantee 100% full table coverage
  const colgroup = document.createElement('colgroup');
  const widths = ['23%', '11%', '11%', '11%', '11%', '11%', '11%', '11%'];
  widths.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // 2. Thead
  const thead = document.createElement('thead');

  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = headers.length;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = headers.length;
  th2.textContent = `Production Output Summary Report (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  const trH = document.createElement('tr');
  headers.forEach((h, idx) => {
    const th = document.createElement('th');
    th.textContent = h;
    th.style.fontWeight = 'bold';
    th.style.padding = '8px 6px';
    th.style.textAlign = idx === 0 ? 'left' : (idx === 6 ? 'center' : 'right');
    th.style.border = '1px solid #000000';
    th.style.backgroundColor = (idx === 0) ? '#7EC8E3' : ((idx === 6) ? '#38B6FF' : '#56C5D0');
    th.style.color = '#000000';
    trH.appendChild(th);
  });
  thead.appendChild(trH);
  table.appendChild(thead);

  // 3. Tbody
  const tbody = document.createElement('tbody');
  const grandTotals = { runMins: 0, cap: 0, act: 0, rej: 0 };

  Object.keys(SHEET_TABS).forEach(tabId => {
    if (SHEET_TABS[tabId].isSummary) return;
    const s = getTabProductionOutputSummary(tabId, year, MonthYearState.monthIndex);
    grandTotals.runMins += s.runningMins;
    grandTotals.cap += s.capacityPcs;
    grandTotals.act += s.actualPrdPcs;
    grandTotals.rej += s.rejectionPcs;

    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';
    tr.title = `Click to view ${s.name}`;
    tr.addEventListener('click', () => switchTab(tabId));

    const isHighAch = s.achievement >= 0.95;

    tr.innerHTML = `
      <td style="font-weight:bold; color:#1E40AF; text-align:left; border:1px solid #CBD5E1; padding:6px 8px; background:#F8FAFC;">
        ${s.name}
      </td>
      <td style="text-align:right; font-weight:bold; border:1px solid #CBD5E1; padding:6px 8px;">
        ${s.runHours > 0 ? s.runHours.toFixed(1) : '-'}
      </td>
      <td style="text-align:right; font-weight:bold; border:1px solid #CBD5E1; padding:6px 8px;">
        ${s.capacityPcs > 0 ? s.capacityPcs.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; border:1px solid #CBD5E1; padding:6px 8px;">
        ${s.targetPerHr > 0 ? s.targetPerHr.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; font-weight:bold; color:#0F766E; border:1px solid #CBD5E1; padding:6px 8px; background:#F0FDFA;">
        ${s.actualPrdPcs > 0 ? s.actualPrdPcs.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; border:1px solid #CBD5E1; padding:6px 8px;">
        ${s.actualPerHr > 0 ? s.actualPerHr.toLocaleString() : '-'}
      </td>
      <td style="text-align:center; font-weight:bold; border:1px solid #CBD5E1; padding:6px 8px; ${isHighAch ? 'background:#E0F2FE; color:#0369A1;' : ''}">
        ${s.capacityPcs > 0 ? (s.achievement * 100).toFixed(1) + '%' : '-'}
      </td>
      <td style="text-align:right; font-weight:bold; color:#B91C1C; border:1px solid #CBD5E1; padding:6px 8px; background:#FEF2F2;">
        ${s.rejectionPcs > 0 ? s.rejectionPcs.toLocaleString() : '0'}
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Grand Total Row
  const totalRunHrs = Math.round((grandTotals.runMins / 60) * 10) / 10;
  const totalTargetPerHr = totalRunHrs > 0 ? Math.round(grandTotals.cap / totalRunHrs) : 0;
  const totalActualPerHr = totalRunHrs > 0 ? Math.round(grandTotals.act / totalRunHrs) : 0;
  const totalAch = grandTotals.cap > 0 ? (grandTotals.act / grandTotals.cap) : 0;

  const trTotal = document.createElement('tr');
  trTotal.style.backgroundColor = '#EBF4FF';
  trTotal.style.fontWeight = 'bold';
  trTotal.style.borderTop = '2px solid #1E40AF';
  trTotal.style.borderBottom = '2px solid #1E40AF';

  trTotal.innerHTML = `
    <td style="text-align:left; border:1px solid #CBD5E1; padding:8px 8px; color:#1E3A8A;">Total:</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalRunHrs > 0 ? totalRunHrs.toFixed(1) : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${grandTotals.cap > 0 ? grandTotals.cap.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalTargetPerHr > 0 ? totalTargetPerHr.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px; color:#0F766E;">${grandTotals.act > 0 ? grandTotals.act.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalActualPerHr > 0 ? totalActualPerHr.toLocaleString() : '-'}</td>
    <td style="text-align:center; border:1px solid #CBD5E1; padding:8px 8px; color:#0369A1;">${grandTotals.cap > 0 ? (totalAch * 100).toFixed(1) + '%' : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px; color:#B91C1C;">${grandTotals.rej.toLocaleString()}</td>
  `;
  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ─── TAB DOWNTIME & RUNNING TIME STATUS CALCULATOR ────────────────────────────
function getTabDowntimeRunningStatus(tabId, year, monthIndex) {
  let rowsData = [];
  if (tabId === ACTIVE_TAB && monthIndex === MonthYearState.monthIndex && year === MonthYearState.year) {
    rowsData = SheetState.rows;
  } else {
    const localData = getStoredLocalData(tabId, year, monthIndex);
    if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
      rowsData = localData.rows;
    } else if (tabId === 'fan_lathe' && year === 2026 && monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
      rowsData = INITIAL_EXCEL_ROWS;
    } else {
      rowsData = generateBlankMonthRows(tabId, year, monthIndex);
    }
  }

  let plannedTimeMins = 0;
  let downTimeMins = 0;
  const rowsPerDay = getRowsPerDay(tabId);

  const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);

  rowsData.forEach(r => {
    const mIdx = (r.row - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(tabId, mIdx);

    if (groupInfo.isMaster) {
      plannedTimeMins += Number(r.H?.val) || 0;

      let rowDt = 0;
      dtCols.forEach(c => {
        rowDt += Number(r[c.col]?.val) || 0;
      });
      downTimeMins += (typeof r.AH?.val === 'number' && r.AH.val > 0 ? r.AH.val : rowDt);
    }
  });

  const runTimeMins = plannedTimeMins - downTimeMins;
  const runTimePct = plannedTimeMins > 0 ? (runTimeMins / plannedTimeMins) : 0;
  const downTimePct = plannedTimeMins > 0 ? (downTimeMins / plannedTimeMins) : 0;

  return {
    tabId,
    name: SHEET_TABS[tabId]?.name || tabId,
    plannedTimeMins,
    runTimeMins,
    downTimeMins,
    runTimePct,
    downTimePct
  };
}

// ─── EXECUTIVE DOWNTIME & RUNNING STATUS REPORT TABLE RENDERER ────────────────
function renderDowntimeRunningStatusReport(table) {
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;

  // Column Widths (% based)
  const colgroup = document.createElement('colgroup');
  const widths = ['20%', '13.5%', '13.5%', '13.5%', '13%', '13%', '13.5%'];
  widths.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // Thead
  const thead = document.createElement('thead');

  // Row 1: MEP FAN LTD.
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = 7;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: MEP Fan Limited- Down Time & Running Time Status (Month Year)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = 7;
  th2.textContent = `MEP Fan Limited- Down Time & Running Time Status (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Header Row (Sage Green Background + Red Remark's)
  const tr3 = document.createElement('tr');

  const headers = [
    { label: 'Machine No.', isRed: false },
    { label: 'Planned\nProduction Time\n(mins)', isRed: false },
    { label: 'Production Run\nTime (Mins)', isRed: false },
    { label: 'Machine Down\nTime (Mins)', isRed: false },
    { label: 'Production\nRunning Time (%)', isRed: false },
    { label: 'Production Down\nTime (%)', isRed: false },
    { label: "Remark's", isRed: true }
  ];

  headers.forEach(h => {
    const th = document.createElement('th');
    th.className = h.isRed ? 'status-remarks-header' : 'status-header-cell';
    th.innerHTML = h.label.replace(/\n/g, '<br>');
    tr3.appendChild(th);
  });

  thead.appendChild(tr3);
  table.appendChild(thead);

  // Tbody
  const tbody = document.createElement('tbody');

  let grandPlanned = 0;
  let grandRun = 0;
  let grandDown = 0;

  getActiveSummaryDepts().forEach(dept => {
    const status = getTabDowntimeRunningStatus(dept.id, year, MonthYearState.monthIndex);

    grandPlanned += status.plannedTimeMins;
    grandRun += status.runTimeMins;
    grandDown += status.downTimeMins;

    const tr = document.createElement('tr');

    // Col 1: Machine No.
    const tdName = document.createElement('td');
    tdName.className = 'status-name-cell';
    tdName.textContent = dept.name;
    tr.appendChild(tdName);

    // Col 2: Planned Production Time (mins)
    const tdPlanned = document.createElement('td');
    tdPlanned.className = 'summary-data-cell';
    tdPlanned.textContent = status.plannedTimeMins > 0 ? status.plannedTimeMins.toLocaleString() : '-';
    tr.appendChild(tdPlanned);

    // Col 3: Production Run Time (Mins)
    const tdRun = document.createElement('td');
    tdRun.className = 'summary-data-cell font-bold';
    tdRun.textContent = status.runTimeMins !== 0 ? status.runTimeMins.toLocaleString() : '0';
    tr.appendChild(tdRun);

    // Col 4: Machine Down Time (Mins)
    const tdDown = document.createElement('td');
    tdDown.className = 'summary-data-cell font-bold';
    tdDown.textContent = status.downTimeMins > 0 ? status.downTimeMins.toLocaleString() : '-';
    tr.appendChild(tdDown);

    // Col 5: Production Running Time (%)
    const tdRunPct = document.createElement('td');
    tdRunPct.className = 'summary-data-cell font-bold';
    tdRunPct.textContent = `${(status.runTimePct * 100).toFixed(1)}%`;
    tr.appendChild(tdRunPct);

    // Col 6: Production Down Time (%)
    const tdDownPct = document.createElement('td');
    tdDownPct.className = 'summary-data-cell font-bold';
    tdDownPct.textContent = `${(status.downTimePct * 100).toFixed(1)}%`;
    tr.appendChild(tdDownPct);

    // Col 7: Remark's
    const tdRem = document.createElement('td');
    tdRem.className = 'summary-data-cell';
    tdRem.textContent = '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row (SubTotal)
  const trTotal = document.createElement('tr');

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'status-subtotal-title font-bold';
  tdTotTitle.textContent = 'SubTotal';
  trTotal.appendChild(tdTotTitle);

  const tdTotPlanned = document.createElement('td');
  tdTotPlanned.className = 'status-subtotal-cell';
  tdTotPlanned.textContent = grandPlanned.toLocaleString();
  trTotal.appendChild(tdTotPlanned);

  const tdTotRun = document.createElement('td');
  tdTotRun.className = 'status-subtotal-cell';
  tdTotRun.textContent = grandRun.toLocaleString();
  trTotal.appendChild(tdTotRun);

  const tdTotDown = document.createElement('td');
  tdTotDown.className = 'status-subtotal-cell';
  tdTotDown.textContent = grandDown.toLocaleString();
  trTotal.appendChild(tdTotDown);

  const tdTotRunPct = document.createElement('td');
  tdTotRunPct.className = 'status-subtotal-cell';
  const grandRunPct = grandPlanned > 0 ? (grandRun / grandPlanned) * 100 : 0;
  tdTotRunPct.textContent = `${grandRunPct.toFixed(1)}%`;
  trTotal.appendChild(tdTotRunPct);

  const tdTotDownPct = document.createElement('td');
  tdTotDownPct.className = 'status-subtotal-cell';
  const grandDownPct = grandPlanned > 0 ? (grandDown / grandPlanned) * 100 : 0;
  tdTotDownPct.textContent = `${grandDownPct.toFixed(1)}%`;
  trTotal.appendChild(tdTotDownPct);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'status-subtotal-cell';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ─── TAB OEE SUMMARY CALCULATOR ───────────────────────────────────────────────
function getTabOEESummary(tabId, year, monthIndex) {
  let rowsData = [];
  if (tabId === ACTIVE_TAB && monthIndex === MonthYearState.monthIndex && year === MonthYearState.year) {
    rowsData = SheetState.rows;
  } else {
    const localData = getStoredLocalData(tabId, year, monthIndex);
    if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
      rowsData = localData.rows;
    } else if (tabId === 'fan_lathe' && year === 2026 && monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
      rowsData = INITIAL_EXCEL_ROWS;
    } else {
      rowsData = generateBlankMonthRows(tabId, year, monthIndex);
    }
  }

  let capacityPcs = 0;
  let totalProduction = 0;
  let rejectionPcs = 0;
  let plannedTimeMins = 0;
  let downTimeMins = 0;
  const rowsPerDay = getRowsPerDay(tabId);

  const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);

  rowsData.forEach(r => {
    const mIdx = (r.row - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(tabId, mIdx);

    capacityPcs += Number(r.E?.val) || 0;
    totalProduction += Number(r.F?.val) || 0;
    rejectionPcs += Number(r.G?.val) || 0;

    if (groupInfo.isMaster) {
      plannedTimeMins += Number(r.H?.val) || 0;

      let rowDt = 0;
      dtCols.forEach(c => {
        rowDt += Number(r[c.col]?.val) || 0;
      });
      downTimeMins += (typeof r.AH?.val === 'number' && r.AH.val > 0 ? r.AH.val : rowDt);
    }
  });

  const runTimeMins = plannedTimeMins - downTimeMins;
  const availability = plannedTimeMins > 0 ? (runTimeMins / plannedTimeMins) : 0;
  const performance = capacityPcs > 0 ? (totalProduction / capacityPcs) : 0;
  const quality = (totalProduction + rejectionPcs) > 0 ? (totalProduction / (totalProduction + rejectionPcs)) : 0;
  const oee = availability * performance * quality;

  return {
    tabId,
    name: SHEET_TABS[tabId]?.name || tabId,
    capacityPcs,
    totalProduction,
    rejectionPcs,
    plannedTimeMins,
    runTimeMins,
    downTimeMins,
    availability,
    performance,
    quality,
    oee
  };
}

// ─── EXECUTIVE OEE REPORT TABLE RENDERER ───────────────────────────────────────
function renderOEEReport(table) {
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;

  // Column Widths (% based)
  const colgroup = document.createElement('colgroup');
  const widths = ['20%', '11%', '11%', '10%', '10%', '10%', '10%', '9%', '9%'];
  widths.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // Thead
  const thead = document.createElement('thead');

  // Row 1: MEP FAN LTD.
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = 9;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: OEE Report (Month Year)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = 9;
  th2.textContent = `OEE Report (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Header Top Row
  const tr3 = document.createElement('tr');

  const thSec = document.createElement('th');
  thSec.rowSpan = 2;
  thSec.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thSec.textContent = 'Section';
  thSec.style.verticalAlign = 'middle';
  tr3.appendChild(thSec);

  const thCap = document.createElement('th');
  thCap.rowSpan = 2;
  thCap.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thCap.innerHTML = 'Machine<br>Capacity (Pcs)';
  thCap.style.verticalAlign = 'middle';
  tr3.appendChild(thCap);

  const thQty = document.createElement('th');
  thQty.colSpan = 2;
  thQty.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1';
  thQty.textContent = 'Production Qty.';
  thQty.style.verticalAlign = 'middle';
  tr3.appendChild(thQty);

  const thAvail = document.createElement('th');
  thAvail.rowSpan = 2;
  thAvail.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thAvail.innerHTML = 'Availability<br>(%)';
  thAvail.style.verticalAlign = 'middle';
  tr3.appendChild(thAvail);

  const thPerf = document.createElement('th');
  thPerf.rowSpan = 2;
  thPerf.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thPerf.innerHTML = 'Performance<br>(%)';
  thPerf.style.verticalAlign = 'middle';
  tr3.appendChild(thPerf);

  const thQual = document.createElement('th');
  thQual.rowSpan = 2;
  thQual.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thQual.innerHTML = 'Quality<br>(%)';
  thQual.style.verticalAlign = 'middle';
  tr3.appendChild(thQual);

  const thOEE = document.createElement('th');
  thOEE.rowSpan = 2;
  thOEE.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-2';
  thOEE.innerHTML = 'OEE (%)';
  thOEE.style.verticalAlign = 'middle';
  tr3.appendChild(thOEE);

  const thRem = document.createElement('th');
  thRem.rowSpan = 2;
  thRem.className = 'status-remarks-header border border-red-800';
  thRem.textContent = "Remark's";
  thRem.style.verticalAlign = 'middle';
  tr3.appendChild(thRem);

  thead.appendChild(tr3);

  // Row 4: Sub-headers
  const tr4 = document.createElement('tr');

  const thTotProd = document.createElement('th');
  thTotProd.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1.5';
  thTotProd.innerHTML = 'Total Production<br>(pcs)';
  thTotProd.style.verticalAlign = 'middle';
  tr4.appendChild(thTotProd);

  const thRej = document.createElement('th');
  thRej.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-2 py-1.5';
  thRej.innerHTML = 'Rejection<br>(Pcs)';
  thRej.style.verticalAlign = 'middle';
  tr4.appendChild(thRej);

  thead.appendChild(tr4);
  table.appendChild(thead);

  // Tbody
  const tbody = document.createElement('tbody');

  let grandCap = 0;
  let grandProd = 0;
  let grandRej = 0;
  let grandPlanned = 0;
  let grandRun = 0;

  getActiveSummaryDepts().forEach(dept => {
    const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);

    grandCap += oeeData.capacityPcs;
    grandProd += oeeData.totalProduction;
    grandRej += oeeData.rejectionPcs;
    grandPlanned += oeeData.plannedTimeMins;
    grandRun += oeeData.runTimeMins;

    const tr = document.createElement('tr');

    // Col 1: Section
    const tdSec = document.createElement('td');
    tdSec.className = 'oee-section-cell';
    tdSec.textContent = dept.name;
    tr.appendChild(tdSec);

    // Col 2: Machine Capacity (Pcs)
    const tdCap = document.createElement('td');
    tdCap.className = 'summary-data-cell';
    tdCap.textContent = oeeData.capacityPcs > 0 ? oeeData.capacityPcs.toLocaleString() : '-';
    tr.appendChild(tdCap);

    // Col 3: Total Production (pcs)
    const tdProd = document.createElement('td');
    tdProd.className = 'summary-data-cell';
    tdProd.textContent = oeeData.totalProduction > 0 ? oeeData.totalProduction.toLocaleString() : '-';
    tr.appendChild(tdProd);

    // Col 4: Rejection (Pcs)
    const tdRej = document.createElement('td');
    tdRej.className = 'summary-data-cell';
    tdRej.textContent = oeeData.rejectionPcs > 0 ? (oeeData.rejectionPcs % 1 !== 0 ? oeeData.rejectionPcs.toFixed(1) : oeeData.rejectionPcs.toLocaleString()) : '0';
    tr.appendChild(tdRej);

    // Col 5: Availability (%)
    const tdAvail = document.createElement('td');
    tdAvail.className = 'summary-data-cell font-bold';
    tdAvail.textContent = `${(oeeData.availability * 100).toFixed(0)}%`;
    tr.appendChild(tdAvail);

    // Col 6: Performance (%)
    const tdPerf = document.createElement('td');
    tdPerf.className = 'summary-data-cell font-bold';
    tdPerf.textContent = `${(oeeData.performance * 100).toFixed(0)}%`;
    tr.appendChild(tdPerf);

    // Col 7: Quality (%)
    const tdQual = document.createElement('td');
    tdQual.className = 'summary-data-cell font-bold';
    tdQual.textContent = `${(oeeData.quality * 100).toFixed(1)}%`;
    tr.appendChild(tdQual);

    // Col 8: OEE (%)
    const tdOEE = document.createElement('td');
    tdOEE.className = 'summary-data-cell font-bold';
    tdOEE.textContent = `${(oeeData.oee * 100).toFixed(1)}%`;
    tr.appendChild(tdOEE);

    // Col 9: Remark's
    const tdRem = document.createElement('td');
    tdRem.className = 'summary-data-cell';
    tdRem.textContent = '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row
  const trTotal = document.createElement('tr');

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'oee-total-title font-bold';
  tdTotTitle.textContent = 'Total';
  trTotal.appendChild(tdTotTitle);

  const tdTotCap = document.createElement('td');
  tdTotCap.className = 'oee-total-cell font-bold';
  tdTotCap.textContent = grandCap.toLocaleString();
  trTotal.appendChild(tdTotCap);

  const tdTotProd = document.createElement('td');
  tdTotProd.className = 'oee-total-cell font-bold';
  tdTotProd.textContent = grandProd.toLocaleString();
  trTotal.appendChild(tdTotProd);

  const tdTotRej = document.createElement('td');
  tdTotRej.className = 'oee-total-cell font-bold';
  tdTotRej.textContent = grandRej > 0 ? (grandRej % 1 !== 0 ? grandRej.toFixed(1) : grandRej.toLocaleString()) : '0';
  trTotal.appendChild(tdTotRej);

  const grandAvail = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
  const grandPerf = grandCap > 0 ? (grandProd / grandCap) : 0;
  const grandQual = (grandProd + grandRej) > 0 ? (grandProd / (grandProd + grandRej)) : 0;
  const grandOEE = grandAvail * grandPerf * grandQual;

  const tdTotAvail = document.createElement('td');
  tdTotAvail.className = 'oee-total-cell font-bold';
  tdTotAvail.textContent = `${(grandAvail * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotAvail);

  const tdTotPerf = document.createElement('td');
  tdTotPerf.className = 'oee-total-cell font-bold';
  tdTotPerf.textContent = `${(grandPerf * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotPerf);

  const tdTotQual = document.createElement('td');
  tdTotQual.className = 'oee-total-cell font-bold';
  tdTotQual.textContent = `${(grandQual * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotQual);

  const tdTotOEE = document.createElement('td');
  tdTotOEE.className = 'oee-total-cell font-bold';
  tdTotOEE.textContent = `${(grandOEE * 100).toFixed(1)}%`;
  trTotal.appendChild(tdTotOEE);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'oee-total-cell';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ─── YEARLY OEE SUMMARY CALCULATOR (ALL 12 MONTHS) ───────────────────────────
// ─── FIXED PERMANENT HISTORICAL 2026 OEE SUMMARY DATA (JANUARY TO JULY 2026) ─
const FIXED_YEARLY_OEE_2026 = {
  0: { // January
    monthName: 'January',
    capacityPcs: 1124393,
    totalProduction: 611910,
    rejectionPcs: 2815.9,
    achievement: 0.544,
    availability: 0.74,
    performance: 0.75,
    quality: 0.99,
    oee: 0.60
  },
  1: { // February
    monthName: 'February',
    capacityPcs: 618970,
    totalProduction: 370274,
    rejectionPcs: 1779,
    achievement: 0.598,
    availability: 0.73,
    performance: 0.79,
    quality: 0.99,
    oee: 0.61
  },
  2: { // March
    monthName: 'March',
    capacityPcs: 609528,
    totalProduction: 281886,
    rejectionPcs: 1055,
    achievement: 0.462,
    availability: 0.57,
    performance: 0.66,
    quality: 1.00,
    oee: 0.42
  },
  3: { // April
    monthName: 'April',
    capacityPcs: 1111210,
    totalProduction: 677516,
    rejectionPcs: 2648,
    achievement: 0.610,
    availability: 0.76,
    performance: 0.73,
    quality: 1.00,
    oee: 0.58
  },
  4: { // May
    monthName: 'May',
    capacityPcs: 30000,
    totalProduction: 30088,
    rejectionPcs: 579,
    achievement: 1.003,
    availability: 0.89,
    performance: 0.99,
    quality: 0.987,
    oee: 0.87
  },
  5: { // Jun
    monthName: 'Jun',
    capacityPcs: 211140,
    totalProduction: 207640,
    rejectionPcs: 802,
    achievement: 0.983,
    availability: 0.94,
    performance: 0.98,
    quality: 0.995,
    oee: 0.92
  },
  6: { // July
    monthName: 'July',
    capacityPcs: 146537,
    totalProduction: 144720,
    rejectionPcs: 476,
    achievement: 0.988,
    availability: 0.93,
    performance: 0.97,
    quality: 0.990,
    oee: 0.90
  }
};

// ─── YEARLY OEE SUMMARY CALCULATOR (ALL 12 MONTHS) ───────────────────────────
function getYearlyOEESummary(year) {
  const monthNames = MonthYearState.monthNames;
  const yearlyData = [];

  for (let m = 0; m < 12; m++) {
    // If year 2026 and month is Jan-Jul (m <= 6), use the fixed permanent historical company data!
    if (year === 2026 && FIXED_YEARLY_OEE_2026[m]) {
      const fixed = FIXED_YEARLY_OEE_2026[m];
      yearlyData.push({
        monthIndex: m,
        monthName: fixed.monthName || monthNames[m],
        capacityPcs: fixed.capacityPcs,
        totalProduction: fixed.totalProduction,
        rejectionPcs: fixed.rejectionPcs,
        plannedTimeMins: 0,
        runTimeMins: 0,
        availability: fixed.availability,
        performance: fixed.performance,
        quality: fixed.quality,
        oee: fixed.oee,
        achievement: fixed.achievement,
        isFixed: true
      });
      continue;
    }

    let grandCap = 0;
    let grandProd = 0;
    let grandRej = 0;
    let grandPlanned = 0;
    let grandRun = 0;

    getActiveSummaryDepts().forEach(dept => {
      const oeeData = getTabOEESummary(dept.id, year, m);
      grandCap += oeeData.capacityPcs;
      grandProd += oeeData.totalProduction;
      grandRej += oeeData.rejectionPcs;
      grandPlanned += oeeData.plannedTimeMins;
      grandRun += oeeData.runTimeMins;
    });

    const availability = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
    const performance = grandCap > 0 ? (grandProd / grandCap) : 0;
    const quality = (grandProd + grandRej) > 0 ? (grandProd / (grandProd + grandRej)) : 0;
    const oee = availability * performance * quality;
    const achievement = grandCap > 0 ? (grandProd / grandCap) : 0;

    yearlyData.push({
      monthIndex: m,
      monthName: monthNames[m],
      capacityPcs: grandCap,
      totalProduction: grandProd,
      rejectionPcs: grandRej,
      plannedTimeMins: grandPlanned,
      runTimeMins: grandRun,
      availability,
      performance,
      quality,
      oee,
      achievement,
      isFixed: false
    });
  }

  return yearlyData;
}

// ─── EXECUTIVE SUMMARY OF OEE (12 MONTH ANNUAL REPORT) TABLE RENDERER ─────────
function renderYearlyOEESummaryReport(table) {
  const year = MonthYearState.year;
  const yearlyData = getYearlyOEESummary(year);

  // Column Widths (% based)
  const colgroup = document.createElement('colgroup');
  const widths = ['10%', '13%', '11%', '11%', '9%', '9%', '9%', '9%', '9%', '10%'];
  widths.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // Thead
  const thead = document.createElement('thead');

  // Row 1: MEP FAN LTD.
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = 10;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: Summary of OEE (Year)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = 10;
  th2.textContent = `Summary of OEE (${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Header Top Row
  const tr3 = document.createElement('tr');

  const headers = [
    { label: 'Month', rowSpan: 2 },
    { label: 'Detail', rowSpan: 2 },
    { label: 'Machine\nCapacity\n(Pcs)', rowSpan: 2 },
    { label: 'Production Qty.', colSpan: 2 },
    { label: 'Availability (%)', rowSpan: 2 },
    { label: 'Performance\n(%)', rowSpan: 2 },
    { label: 'Quality (%)', rowSpan: 2 },
    { label: 'OEE (%)', rowSpan: 2 },
    { label: "Remark's", rowSpan: 2 }
  ];

  headers.forEach(h => {
    const th = document.createElement('th');
    if (h.rowSpan) th.rowSpan = h.rowSpan;
    if (h.colSpan) th.colSpan = h.colSpan;
    th.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-1 py-2';
    th.innerHTML = h.label.replace(/\n/g, '<br>');
    th.style.verticalAlign = 'middle';
    tr3.appendChild(th);
  });

  thead.appendChild(tr3);

  // Row 4: Sub-headers for Production (pcs) and Rejection (Pcs)
  const tr4 = document.createElement('tr');

  const thTotProd = document.createElement('th');
  thTotProd.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-1 py-1.5';
  thTotProd.innerHTML = 'Total<br>Production<br>(pcs)';
  thTotProd.style.verticalAlign = 'middle';
  tr4.appendChild(thTotProd);

  const thRej = document.createElement('th');
  thRej.className = 'border border-slate-600 bg-[#7EC8E3] text-black font-bold text-xs text-center align-middle px-1 py-1.5';
  thRej.innerHTML = 'Rejection<br>(Pcs)';
  thRej.style.verticalAlign = 'middle';
  tr4.appendChild(thRej);

  thead.appendChild(tr4);
  table.appendChild(thead);

  // Tbody
  const tbody = document.createElement('tbody');

  yearlyData.forEach((m, idx) => {
    const isRunningMonth = (idx === MonthYearState.monthIndex);

    // Row 1: Total
    const trA = document.createElement('tr');
    if (isRunningMonth) {
      trA.className = 'yearly-running-month-row';
    }

    // Col 1: Month (rowSpan 2)
    const tdMonth = document.createElement('td');
    tdMonth.rowSpan = 2;
    tdMonth.className = isRunningMonth ? 'yearly-month-cell running-month-active' : 'yearly-month-cell';
    tdMonth.textContent = m.monthName;
    trA.appendChild(tdMonth);

    // Col 2: Detail
    const tdDetailA = document.createElement('td');
    tdDetailA.className = isRunningMonth ? 'yearly-detail-cell running-month-detail' : 'yearly-detail-cell';
    tdDetailA.textContent = 'Total';
    trA.appendChild(tdDetailA);

    // Col 3: Machine Capacity
    const tdCap = document.createElement('td');
    tdCap.className = isRunningMonth ? 'summary-data-cell running-month-data font-mono font-bold' : 'summary-data-cell font-mono';
    tdCap.textContent = m.capacityPcs > 0 ? m.capacityPcs.toLocaleString() : '-';
    trA.appendChild(tdCap);

    // Col 4: Total Production
    const tdProd = document.createElement('td');
    tdProd.className = isRunningMonth ? 'summary-data-cell running-month-data font-mono font-bold' : 'summary-data-cell font-mono';
    tdProd.textContent = m.totalProduction > 0 ? m.totalProduction.toLocaleString() : '-';
    trA.appendChild(tdProd);

    // Col 5: Rejection
    const tdRej = document.createElement('td');
    tdRej.className = isRunningMonth ? 'summary-data-cell running-month-data font-mono font-bold' : 'summary-data-cell font-mono';
    tdRej.textContent = m.rejectionPcs > 0 ? (m.rejectionPcs % 1 !== 0 ? m.rejectionPcs.toFixed(1) : m.rejectionPcs.toLocaleString()) : '0';
    trA.appendChild(tdRej);

    // Col 6: Availability (rowSpan 2)
    const tdAvail = document.createElement('td');
    tdAvail.rowSpan = 2;
    tdAvail.className = isRunningMonth ? 'yearly-kpi-cell running-month-kpi font-bold font-mono' : 'yearly-kpi-cell font-bold font-mono';
    tdAvail.textContent = m.capacityPcs > 0 || m.plannedTimeMins > 0 || m.isFixed ? `${(m.availability * 100).toFixed(0)}%` : '-';
    trA.appendChild(tdAvail);

    // Col 7: Performance (rowSpan 2)
    const tdPerf = document.createElement('td');
    tdPerf.rowSpan = 2;
    tdPerf.className = isRunningMonth ? 'yearly-kpi-cell running-month-kpi font-bold font-mono' : 'yearly-kpi-cell font-bold font-mono';
    tdPerf.textContent = m.capacityPcs > 0 || m.isFixed ? `${(m.performance * 100).toFixed(0)}%` : '-';
    trA.appendChild(tdPerf);

    // Col 8: Quality (rowSpan 2)
    const tdQual = document.createElement('td');
    tdQual.rowSpan = 2;
    tdQual.className = isRunningMonth ? 'yearly-kpi-cell running-month-kpi font-bold font-mono' : 'yearly-kpi-cell font-bold font-mono';
    if (m.isFixed && m.quality !== undefined) {
      const qVal = m.quality * 100;
      tdQual.textContent = `${qVal.toFixed(qVal % 1 !== 0 ? 1 : 0)}%`;
    } else {
      tdQual.textContent = m.totalProduction > 0 ? `${(m.quality * 100).toFixed(m.quality >= 0.999 ? 0 : 1)}%` : '-';
    }
    trA.appendChild(tdQual);

    // Col 9: OEE (rowSpan 2)
    const tdOEE = document.createElement('td');
    tdOEE.rowSpan = 2;
    tdOEE.className = isRunningMonth ? 'yearly-kpi-cell running-month-kpi font-bold font-mono' : 'yearly-kpi-cell font-bold font-mono';
    tdOEE.textContent = m.capacityPcs > 0 || m.isFixed ? `${(m.oee * 100).toFixed(0)}%` : '-';
    trA.appendChild(tdOEE);

    // Col 10: Remark's (rowSpan 2)
    const tdRem = document.createElement('td');
    tdRem.rowSpan = 2;
    tdRem.className = isRunningMonth ? 'yearly-remarks-cell running-month-data' : 'yearly-remarks-cell';
    tdRem.textContent = '';
    trA.appendChild(tdRem);

    tbody.appendChild(trA);

    // Row 2: Total Acheivement (%)
    const trB = document.createElement('tr');
    trB.className = isRunningMonth ? 'yearly-month-divider yearly-running-month-row-b' : 'yearly-month-divider';

    // Col 2: Detail
    const tdDetailB = document.createElement('td');
    tdDetailB.className = isRunningMonth ? 'yearly-detail-cell running-month-detail font-bold' : 'yearly-detail-cell';
    tdDetailB.textContent = 'Total Acheivement (%)';
    trB.appendChild(tdDetailB);

    // Col 3, 4, 5: Merged Total Acheivement Value (colSpan 3)
    const tdAch = document.createElement('td');
    tdAch.colSpan = 3;
    const isHighlight = m.achievement >= 0.95 && m.capacityPcs > 0;
    if (isRunningMonth) {
      tdAch.className = 'yearly-ach-cell running-month-ach';
    } else {
      tdAch.className = isHighlight ? 'yearly-ach-highlight font-mono' : 'yearly-ach-cell font-mono';
    }
    tdAch.textContent = m.capacityPcs > 0 || m.isFixed ? `${(m.achievement * 100).toFixed(1)}%` : '-';
    trB.appendChild(tdAch);

    tbody.appendChild(trB);

    // Gap / Separator row after each month (matching screenshot gap)
    const trGap = document.createElement('tr');
    trGap.className = 'yearly-gap-row';
    const tdGap = document.createElement('td');
    tdGap.colSpan = 10;
    tdGap.className = 'yearly-gap-cell';
    trGap.appendChild(tdGap);
    tbody.appendChild(trGap);
  });

  table.appendChild(tbody);
}

// ─── RENDERING THE SPREADSHEET ────────────────────────────────────────────────
function renderExcelTable() {
  const table = document.getElementById('excelMainTable');
  if (!table) return;

  table.replaceChildren();

  // If viewing Executive Summary Reports
  if (ACTIVE_TAB === 'summary_downtime') {
    renderSummaryDowntimeReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_production') {
    renderProductionOutputReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_status') {
    renderDowntimeRunningStatusReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_oee') {
    renderOEEReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_oee_yearly') {
    renderYearlyOEESummaryReport(table);
    return;
  }

  // Column Widths (% based to span 100% full screen width)
  // Perfectly balanced so even the longest names like "Auto Die Casting (Cover)" NEVER get clipped in any tab!
  const colgroup = document.createElement('colgroup');
  EXCEL_COLUMNS.forEach(c => {
    const col = document.createElement('col');
    if (c.col === 'A') col.style.width = '4.5%'; // Fits '1-Aug-26' cleanly
    else if (c.col === 'B') col.style.width = '2.0%'; // Fits 'Sat', 'Sun' cleanly
    else if (c.col === 'C') col.style.width = '3.0%'; // Fits 'Morning', 'Night' completely
    else if (c.col === 'D') col.style.width = '10.5%'; // Fits 'Auto Die Casting (Cover)' completely with margins
    else if (['E', 'F', 'H', 'J', 'AH'].includes(c.col)) col.style.width = '3.6%'; // Fits 5-digit numbers (11111)
    else if (c.col === 'G') col.style.width = '2.5%'; // Fits rejection numbers
    else if (c.col === 'I') col.style.width = '2.2%'; // Fits Expected DT (30/0)
    else if (['AI', 'AJ', 'AK', 'AL'].includes(c.col)) col.style.width = '2.1%'; // Fits 100% KPI
    else if (c.col === 'AM') col.style.width = '4.5%'; // Remarks
    else col.style.width = '2.0%'; // 23 Downtime columns (Fits 3 digits comfortably like 120, 250, 480)
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  // Table Headers
  const thead = document.createElement('thead');
  const tabInfo = SHEET_TABS[ACTIVE_TAB];
  const rowsPerDay = getRowsPerDay();

  // Row 1: MEP FAN LTD. (colSpan 39)
  const tr1 = document.createElement('tr');
  tr1.className = 'mep-banner-row1';
  const th1 = document.createElement('th');
  th1.colSpan = EXCEL_COLUMNS.length;
  th1.textContent = 'MEP FAN LTD.';
  tr1.appendChild(th1);
  thead.appendChild(tr1);

  // Row 2: Production Performance Analysis Report (colSpan 39)
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = EXCEL_COLUMNS.length;
  th2.textContent = 'Production Performance Analysis Report';
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Dark Navy Blue Bar with Department/Tab Name (colSpan 39)
  const tr3 = document.createElement('tr');
  tr3.className = 'mep-banner-row3';
  const th3 = document.createElement('th');
  th3.colSpan = EXCEL_COLUMNS.length;
  th3.className = 'th-month-bar';
  th3.textContent = tabInfo.name;
  tr3.appendChild(th3);
  thead.appendChild(tr3);

  // Row 4: Column Headers
  const trHeaders = document.createElement('tr');

  EXCEL_COLUMNS.forEach(c => {
    const th = document.createElement('th');
    th.title = c.label;

    // Horizontal / Upright headers up to Machine Name (Cols A, B, C, D)
    if (['A', 'B', 'C', 'D'].includes(c.col)) {
      th.className = `h-header-cell zone-${c.zone}`;
      th.textContent = c.label;
    } else {
      th.className = `v-header-cell zone-${c.zone}`;
      const span = document.createElement('span');
      span.className = 'v-header-text';
      span.textContent = c.label;
      th.appendChild(span);
    }

    trHeaders.appendChild(th);
  });
  thead.appendChild(trHeaders);

  // Row 5: RED TOTAL ROW
  const trTotal = document.createElement('tr');
  trTotal.className = 'mep-total-row';
  trTotal.dataset.row = 5;

  // A5:D5 Merged: Total:
  const tdTotalTitle = document.createElement('td');
  tdTotalTitle.colSpan = 4;
  tdTotalTitle.className = 'total-title-cell text-center';
  tdTotalTitle.textContent = 'Total:';
  tdTotalTitle.dataset.row = 5;
  tdTotalTitle.dataset.col = 'A';
  tdTotalTitle.addEventListener('click', (e) => {
    e.stopPropagation();
    selectCell('A', 5);
  });
  trTotal.appendChild(tdTotalTitle);

  // E5 to AM5 Total Cells
  const maxR = getMaxActiveRow();
  EXCEL_COLUMNS.slice(4).forEach(colDef => {
    const colLetter = colDef.col;
    const td = document.createElement('td');
    td.dataset.row = 5;
    td.dataset.col = colLetter;
    td.className = 'text-center font-mono';

    const val = SheetState.totals[colLetter];
    td.textContent = formatCellValue(val, colDef, true);

    if (colDef.isPercent) {
      if (colLetter === 'AI') td.dataset.formula = '=IFERROR(J5/H5,0)';
      if (colLetter === 'AJ') td.dataset.formula = '=IFERROR(F5/E5,0)';
      if (colLetter === 'AK') td.dataset.formula = '=IFERROR(F5/(F5+G5),0)';
      if (colLetter === 'AL') td.dataset.formula = '=IFERROR(AK5*AJ5*AI5,0)';
    } else if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
      td.dataset.formula = `=SUM(${colLetter}6:${colLetter}${maxR})`;
    }

    td.addEventListener('click', (e) => {
      e.stopPropagation();
      selectCell(colLetter, 5);
    });

    trTotal.appendChild(td);
  });

  thead.appendChild(trTotal);
  table.appendChild(thead);

  // Body Rows
  const tbody = document.createElement('tbody');

  SheetState.rows.forEach(rowObj => {
    const r = rowObj.row;
    const dayNum = getDayFromRow(r);
    const locked = isRowLocked(r);
    const isToday = isDateInCurrentMonth(dayNum);
    const dayName = rowObj.B?.val;
    const isFriday = (dayName === 'Fri');
    const isDayEnd = ((r - 6) % rowsPerDay === (rowsPerDay - 1)); // Last machine of each day
    const mIdx = (r - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);

    const tr = document.createElement('tr');
    tr.dataset.row = r;
    tr.dataset.day = dayNum;

    if (locked) tr.classList.add('row-locked');
    if (isFriday) tr.classList.add('row-friday');
    if (isToday) tr.classList.add('row-today');
    if (isDayEnd) tr.classList.add('row-day-end');

    EXCEL_COLUMNS.forEach((colDef, cIdx) => {
      const colLetter = colDef.col;
      const isTimeCol = TIME_COLUMNS.includes(colLetter);
      const isDateDayShiftCol = ['A', 'B', 'C'].includes(colLetter);

      // If slave row in a merged time group, skip rendering the time columns!
      if (isTimeCol && groupInfo.isSlave) {
        return;
      }

      // Merge Date (A), Day (B), Shift (C) for all machines on the same day
      if (isDateDayShiftCol && mIdx > 0) {
        return; // Skip rendering in secondary machine rows of the same day
      }

      const td = document.createElement('td');
      td.dataset.row = r;
      td.dataset.col = colLetter;
      td.dataset.cidx = cIdx;

      if (isDateDayShiftCol && mIdx === 0 && rowsPerDay > 1) {
        td.rowSpan = rowsPerDay;
        td.style.verticalAlign = 'middle';
      }

      if (isTimeCol && groupInfo.isMaster && groupInfo.count > 1) {
        td.rowSpan = groupInfo.count;
        td.style.verticalAlign = 'middle';
      }

      if (colDef.isReadOnly) {
        td.className = 'cell-readonly-fixed';
      } else if (colLetter === 'J') {
        td.className = 'cell-formula-runtime';
      } else if (colDef.isDt) {
        td.className = 'cell-downtime';
      } else if (colLetter === 'AH') {
        td.className = 'cell-total-dt';
      } else if (colDef.isFormula || colDef.isPercent) {
        td.className = 'cell-kpi';
      } else {
        td.className = 'cell-white';
      }

      if (isDateDayShiftCol || isDayEnd || (isTimeCol && groupInfo.isMaster && (mIdx + groupInfo.count >= rowsPerDay))) {
        td.classList.add('cell-day-end-border');
      }

      const cellVal = rowObj[colLetter]?.val;
      td.textContent = formatCellValue(cellVal, colDef);

      if (colDef.align === 'right') td.classList.add('text-right', 'font-mono');
      if (colDef.align === 'center') td.classList.add('text-center');

      if (rowObj[colLetter]?.formula) {
        td.dataset.formula = rowObj[colLetter].formula;
      } else if (colDef.formula) {
        td.dataset.formula = colDef.formula.replace(/{r}/g, r);
      }

      // Mouse Events strictly for unlocked rows
      if (!locked) {
        td.addEventListener('mousedown', (e) => {
          if (e.button === 0) {
            onCellMouseDown(colLetter, r, e);
          }
        });

        td.addEventListener('mouseenter', () => {
          onCellMouseEnter(colLetter, r);
        });

        td.addEventListener('dblclick', (e) => {
          e.stopPropagation();
          if (colDef.isReadOnly) {
            return;
          }
          startCellEdit(td, colLetter, r, colDef);
        });
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);

  // Default select cell E6 if no active cell exists so Arrow keys work immediately!
  if (!SheetState.selected || !document.querySelector(`.mep-excel-table [data-col="${SheetState.selected.colLetter}"][data-row="${SheetState.selected.row}"]`)) {
    setTimeout(() => selectCell('E', 6), 10);
  }
}

function formatCellValue(val, colDef, isTotalRow = false) {
  if (val === null || val === undefined || val === '') {
    if (isTotalRow && (colDef.isNumeric || colDef.isDt || colDef.isFormula)) {
      return colDef.isPercent ? '0%' : '0';
    }
    return '';
  }

  if (colDef.isPercent) {
    const num = Number(val) || 0;
    return `${(num * 100).toFixed(0)}%`;
  }

  if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
    if (typeof val === 'number') {
      return val !== 0 ? String(val) : '0';
    }
    const n = Number(val);
    if (!isNaN(n)) {
      return n !== 0 ? String(n) : '0';
    }
  }

  return String(val);
}

// ─── CELL SELECTION, RANGE SELECTION & STATS ──────────────────────────────────
function selectCell(colLetter, rowNum, resetRange = true) {
  if (ACTIVE_TAB === 'summary_downtime' || ACTIVE_TAB === 'summary_production') return;
  if (isRowLocked(rowNum)) return;

  if (SheetState.isEditing && SheetState.activeInput) {
    SheetState.activeInput.blur();
  }

  clearSelectionStyles();

  let targetTd = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="${rowNum}"]`);
  if (!targetTd && ['A', 'B', 'C'].includes(colLetter)) {
    const rowsPerDay = getRowsPerDay();
    const masterRow = rowNum - ((rowNum - 6) % rowsPerDay);
    targetTd = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="${masterRow}"]`);
  }

  // Handle merged time group cells (e.g. APC in Fan Auto Powder Coating)
  if (!targetTd && TIME_COLUMNS.includes(colLetter)) {
    const tabInfo = SHEET_TABS[ACTIVE_TAB];
    if (tabInfo && tabInfo.timeGroups) {
      const rowsPerDay = getRowsPerDay();
      const mIdx = (rowNum - 6) % rowsPerDay;
      const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);
      if (groupInfo.isSlave) {
        const dayStartRow = rowNum - mIdx;
        const masterRow = dayStartRow + groupInfo.masterIdx;
        targetTd = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="${masterRow}"]`);
      }
    }
  }

  if (!targetTd) return;

  targetTd.classList.add('selected-cell');

  SheetState.selected = { row: rowNum, colLetter: colLetter, element: targetTd };
  if (resetRange) {
    SheetState.rangeSelection = { start: { col: colLetter, row: rowNum }, end: { col: colLetter, row: rowNum } };
  }

  attachFillHandle(targetTd);

  // Update Name Box
  const nameBox = document.getElementById('activeCellAddress');
  if (nameBox) nameBox.textContent = `${colLetter}${rowNum}`;

  // Update Formula Bar Input
  const colDef = EXCEL_COLUMNS.find(c => c.col === colLetter);
  const formulaInput = document.getElementById('formulaBarInput');
  if (formulaInput) {
    if (colDef?.isReadOnly) {
      const rowObj = SheetState.rows.find(r => r.row === rowNum);
      formulaInput.value = rowObj ? (rowObj[colLetter]?.val || '') : (rowNum === 5 ? 'Total:' : '');
      formulaInput.readOnly = true;
    } else if (targetTd.dataset.formula) {
      formulaInput.value = targetTd.dataset.formula;
      formulaInput.readOnly = true;
    } else {
      const rowObj = SheetState.rows.find(r => r.row === rowNum);
      const rawVal = rowObj ? rowObj[colLetter]?.val : targetTd.textContent.trim();
      formulaInput.value = (rawVal !== null && rawVal !== undefined) ? rawVal : '';
      formulaInput.readOnly = rowNum === 5;
    }
  }

  updateRangeStats();
}

function attachFillHandle(targetTd) {
  if (!targetTd) return;
  document.querySelectorAll('.excel-fill-handle').forEach(fh => fh.remove());

  const colLetter = targetTd.dataset.col;
  const rowNum = Number(targetTd.dataset.row);
  const colDef = EXCEL_COLUMNS.find(c => c.col === colLetter);

  if (rowNum === 5 || isRowLocked(rowNum) || colDef?.isReadOnly) return;

  const fh = document.createElement('div');
  fh.className = 'excel-fill-handle';
  fh.title = 'Drag to AutoFill (Hold Ctrl for Series)';
  fh.addEventListener('mousedown', onFillHandleMouseDown);
  targetTd.appendChild(fh);
}

function clearSelectionStyles() {
  document.querySelectorAll('.mep-excel-table td.selected-cell, .mep-excel-table th.selected-cell').forEach(td => {
    td.classList.remove('selected-cell');
  });
  document.querySelectorAll('.mep-excel-table td.in-range-selection').forEach(td => {
    td.classList.remove('in-range-selection');
  });
  document.querySelectorAll('.excel-fill-handle').forEach(fh => fh.remove());
}

function onCellMouseDown(colLetter, rowNum, e) {
  if (isRowLocked(rowNum) || rowNum === 5) return;

  if (e && e.shiftKey && SheetState.selected) {
    if (e.preventDefault) e.preventDefault();
    SheetState.rangeSelection.start = { col: SheetState.selected.colLetter, row: SheetState.selected.row };
    SheetState.rangeSelection.end = { col: colLetter, row: rowNum };
    highlightSelectedRange();
    return;
  }

  if (e && e.preventDefault) e.preventDefault();

  // Remove copied border animation on new selection
  document.querySelectorAll('.mep-excel-table td.excel-copied-cell').forEach(td => td.classList.remove('excel-copied-cell'));

  SheetState.isSelecting = true;
  SheetState.rangeSelection = { start: { col: colLetter, row: rowNum }, end: { col: colLetter, row: rowNum } };
  selectCell(colLetter, rowNum, false);
}

function onCellMouseEnter(colLetter, rowNum) {
  if (!SheetState.isSelecting || isRowLocked(rowNum) || rowNum === 5) return;
  SheetState.rangeSelection.end = { col: colLetter, row: rowNum };
  highlightSelectedRange();
}

document.addEventListener('mouseup', () => {
  if (SheetState.isSelecting) {
    SheetState.isSelecting = false;
    updateRangeStats();
  }
});

function highlightSelectedRange() {
  const start = SheetState.rangeSelection?.start;
  const end = SheetState.rangeSelection?.end;
  if (!start || !end) return;

  const maxActive = getMaxActiveRow();
  const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === start.col);
  const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === end.col);
  if (startCIdx === -1 || endCIdx === -1) return;

  const minC = Math.min(startCIdx, endCIdx);
  const maxC = Math.max(startCIdx, endCIdx);
  const minR = Math.min(start.row, end.row);
  const maxR = Math.min(Math.max(start.row, end.row), maxActive);

  document.querySelectorAll('.mep-excel-table td.in-range-selection').forEach(td => td.classList.remove('in-range-selection'));

  for (let r = minR; r <= maxR; r++) {
    if (isRowLocked(r) || r === 5) continue;
    for (let c = minC; c <= maxC; c++) {
      const colLetter = EXCEL_COLUMNS[c].col;
      const td = document.querySelector(`.mep-excel-table td[data-col="${colLetter}"][data-row="${r}"]`);
      if (td && !td.classList.contains('selected-cell')) {
        td.classList.add('in-range-selection');
      }
    }
  }

  // Attach fill handle to bottom-right cell of the selection range
  const brCol = EXCEL_COLUMNS[maxC].col;
  const brTd = document.querySelector(`.mep-excel-table td[data-col="${brCol}"][data-row="${maxR}"]`);
  if (brTd) {
    attachFillHandle(brTd);
  }

  updateRangeStats(minR, maxR, minC, maxC);
}

function updateRangeStats(minR, maxR, minC, maxC) {
  if (minR === undefined) {
    const start = SheetState.rangeSelection.start || SheetState.selected;
    const end = SheetState.rangeSelection.end || SheetState.selected;
    if (!start || !end) return;

    const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === start.col);
    const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === end.col);
    minC = Math.min(startCIdx, endCIdx);
    maxC = Math.max(startCIdx, endCIdx);
    minR = Math.min(start.row, end.row);
    maxR = Math.min(Math.max(start.row, end.row), getMaxActiveRow());
  }

  let count = 0;
  let numCount = 0;
  let sum = 0;

  for (let r = minR; r <= maxR; r++) {
    if (isRowLocked(r)) continue;
    const rowObj = SheetState.rows.find(row => row.row === r);
    for (let c = minC; c <= maxC; c++) {
      count++;
      const colLetter = EXCEL_COLUMNS[c].col;
      const val = rowObj ? rowObj[colLetter]?.val : (r === 5 ? SheetState.totals[colLetter] : null);
      if (val !== null && val !== undefined && val !== '' && !isNaN(val)) {
        sum += Number(val);
        numCount++;
      }
    }
  }

  const avg = numCount > 0 ? (sum / numCount).toFixed(numCount > 1 ? 1 : 0) : 0;

  setText('statCount', `COUNT: ${count}`);
  setText('statSum', `SUM: ${sum.toLocaleString()}`);
  setText('statAverage', `AVERAGE: ${avg}`);
}

// ─── EXCEL FILL HANDLE & AUTOFILL ENGINE ──────────────────────────────────────
function computeAutoFillVal(targetR, targetCIdx, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, isCtrl) {
  if (fillDir === 'down' || fillDir === 'up') {
    const colLetter = EXCEL_COLUMNS[targetCIdx].col;
    const sourceVals = [];
    for (let r = srcMinR; r <= srcMaxR; r++) {
      const rowObj = SheetState.rows.find(row => row.row === r);
      sourceVals.push(rowObj ? (rowObj[colLetter]?.val ?? '') : '');
    }

    const K = sourceVals.length;
    if (K === 0) return '';

    // 1. Single cell source
    if (K === 1) {
      const v = sourceVals[0];
      if (v === '' || v === null || v === undefined) return '';

      const num = Number(v);
      if (!isNaN(num) && typeof v !== 'boolean' && v !== '') {
        if (isCtrl) {
          // Increment series (e.g. 1 -> 2, 3, 4...)
          const offset = (fillDir === 'down') ? (targetR - srcMaxR) : (srcMinR - targetR);
          return fillDir === 'down' ? (num + offset) : (num - offset);
        } else {
          // Copy cell (e.g. 5 -> 5, 5, 5...)
          return num;
        }
      }
      return v; // Copy text
    }

    // 2. Multiple cells source (e.g. [1, 2] -> 3, 4, 5...)
    const nums = sourceVals.map(v => (v !== '' && v !== null && !isNaN(Number(v))) ? Number(v) : null);
    const allNumeric = nums.every(n => n !== null);

    if (allNumeric) {
      let step = (K === 2) ? (nums[1] - nums[0]) : ((nums[K - 1] - nums[0]) / (K - 1));
      if (fillDir === 'down') {
        const offset = targetR - srcMaxR;
        const res = nums[K - 1] + (step * offset);
        return Math.round(res * 100) / 100;
      } else if (fillDir === 'up') {
        const offset = srcMinR - targetR;
        const res = nums[0] - (step * offset);
        return Math.round(res * 100) / 100;
      }
    }

    // Pattern cycle for non-numeric or irregular text
    if (fillDir === 'down') {
      const offset = targetR - srcMaxR;
      const idx = (offset - 1) % K;
      return sourceVals[idx];
    } else {
      const offset = srcMinR - targetR;
      const idx = (K - 1 - ((offset - 1) % K) + K) % K;
      return sourceVals[idx];
    }
  } else if (fillDir === 'right' || fillDir === 'left') {
    const rowObj = SheetState.rows.find(row => row.row === targetR);
    const sourceVals = [];
    for (let c = srcMinC; c <= srcMaxC; c++) {
      const colLetter = EXCEL_COLUMNS[c].col;
      sourceVals.push(rowObj ? (rowObj[colLetter]?.val ?? '') : '');
    }

    const K = sourceVals.length;
    if (K === 0) return '';

    if (K === 1) {
      const v = sourceVals[0];
      if (v === '' || v === null || v === undefined) return '';
      const num = Number(v);
      if (!isNaN(num) && typeof v !== 'boolean' && v !== '') {
        if (isCtrl) {
          const offset = (fillDir === 'right') ? (targetCIdx - srcMaxC) : (srcMinC - targetCIdx);
          return fillDir === 'right' ? (num + offset) : (num - offset);
        } else {
          return num;
        }
      }
      return v;
    }

    const nums = sourceVals.map(v => (v !== '' && v !== null && !isNaN(Number(v))) ? Number(v) : null);
    const allNumeric = nums.every(n => n !== null);

    if (allNumeric) {
      let step = (K === 2) ? (nums[1] - nums[0]) : ((nums[K - 1] - nums[0]) / (K - 1));
      if (fillDir === 'right') {
        const offset = targetCIdx - srcMaxC;
        const res = nums[K - 1] + (step * offset);
        return Math.round(res * 100) / 100;
      } else {
        const offset = srcMinC - targetCIdx;
        const res = nums[0] - (step * offset);
        return Math.round(res * 100) / 100;
      }
    }

    if (fillDir === 'right') {
      const offset = targetCIdx - srcMaxC;
      const idx = (offset - 1) % K;
      return sourceVals[idx];
    } else {
      const offset = srcMinC - targetCIdx;
      const idx = (K - 1 - ((offset - 1) % K) + K) % K;
      return sourceVals[idx];
    }
  }

  return '';
}

function onFillHandleMouseDown(e) {
  e.stopPropagation();
  e.preventDefault();

  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('🔒 ভিউ মোড: ডাটা পরিবর্তন করা সম্ভব নয়।', 'warning');
    return;
  }

  const start = SheetState.rangeSelection?.start || SheetState.selected;
  const end = SheetState.rangeSelection?.end || SheetState.selected;
  if (!start || !end) return;

  const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === start.col);
  const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === end.col);
  if (startCIdx === -1 || endCIdx === -1) return;

  const srcMinC = Math.min(startCIdx, endCIdx);
  const srcMaxC = Math.max(startCIdx, endCIdx);
  const srcMinR = Math.min(start.row, end.row);
  const srcMaxR = Math.min(Math.max(start.row, end.row), getMaxActiveRow());

  let isCtrl = e.ctrlKey;
  let currentHover = { row: srcMaxR, cIdx: srcMaxC };

  // Create or get floating preview tooltip
  let tooltip = document.querySelector('.excel-autofill-tooltip');
  if (!tooltip) {
    tooltip = document.createElement('div');
    tooltip.className = 'excel-autofill-tooltip';
    document.body.appendChild(tooltip);
  }

  const updatePreview = (targetRow, targetCIdx, ctrlPressed, mouseX, mouseY) => {
    document.querySelectorAll('.mep-excel-table td.in-autofill-preview').forEach(td => td.classList.remove('in-autofill-preview'));

    const maxActive = getMaxActiveRow();
    targetRow = Math.max(6, Math.min(targetRow, maxActive));
    targetCIdx = Math.max(0, Math.min(targetCIdx, EXCEL_COLUMNS.length - 1));

    let fillDir = 'none';
    let pMinR = srcMinR, pMaxR = srcMaxR, pMinC = srcMinC, pMaxC = srcMaxC;

    const dRowDown = targetRow - srcMaxR;
    const dRowUp = srcMinR - targetRow;
    const dColRight = targetCIdx - srcMaxC;
    const dColLeft = srcMinC - targetCIdx;

    const maxDelta = Math.max(dRowDown, dRowUp, dColRight, dColLeft);

    if (maxDelta > 0) {
      if (maxDelta === dRowDown) {
        fillDir = 'down';
        pMaxR = targetRow;
      } else if (maxDelta === dRowUp) {
        fillDir = 'up';
        pMinR = targetRow;
      } else if (maxDelta === dColRight) {
        fillDir = 'right';
        pMaxC = targetCIdx;
      } else if (maxDelta === dColLeft) {
        fillDir = 'left';
        pMinC = targetCIdx;
      }
    }

    if (fillDir !== 'none') {
      for (let r = pMinR; r <= pMaxR; r++) {
        if (isRowLocked(r) || r === 5) continue;
        for (let c = pMinC; c <= pMaxC; c++) {
          const isInsideSource = (r >= srcMinR && r <= srcMaxR && c >= srcMinC && c <= srcMaxC);
          if (!isInsideSource) {
            const colLetter = EXCEL_COLUMNS[c].col;
            const td = document.querySelector(`.mep-excel-table td[data-col="${colLetter}"][data-row="${r}"]`);
            if (td) td.classList.add('in-autofill-preview');
          }
        }
      }

      const previewVal = computeAutoFillVal(targetRow, targetCIdx, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, ctrlPressed);
      if (mouseX && mouseY) {
        tooltip.textContent = `${previewVal !== null && previewVal !== undefined ? previewVal : ''}`;
        tooltip.style.left = `${mouseX + 16}px`;
        tooltip.style.top = `${mouseY + 12}px`;
        tooltip.style.display = 'block';
      }
    } else {
      tooltip.style.display = 'none';
    }

    return { fillDir, pMinR, pMaxR, pMinC, pMaxC };
  };

  const onMouseMove = (moveEvent) => {
    isCtrl = moveEvent.ctrlKey;
    const targetTd = moveEvent.target.closest('td');
    if (targetTd && targetTd.dataset.row && targetTd.dataset.col) {
      const r = Number(targetTd.dataset.row);
      const cIdx = EXCEL_COLUMNS.findIndex(c => c.col === targetTd.dataset.col);
      if (r && cIdx !== -1) {
        currentHover = { row: r, cIdx };
      }
    }
    updatePreview(currentHover.row, currentHover.cIdx, isCtrl, moveEvent.clientX, moveEvent.clientY);
  };

  const onKeyDown = (keyEvent) => {
    if (keyEvent.key === 'Control') {
      isCtrl = true;
      updatePreview(currentHover.row, currentHover.cIdx, isCtrl, tooltip.offsetLeft - 16, tooltip.offsetTop - 12);
    }
  };

  const onKeyUp = (keyEvent) => {
    if (keyEvent.key === 'Control') {
      isCtrl = false;
      updatePreview(currentHover.row, currentHover.cIdx, isCtrl, tooltip.offsetLeft - 16, tooltip.offsetTop - 12);
    }
  };

  const onMouseUp = (upEvent) => {
    document.removeEventListener('mousemove', onMouseMove);
    document.removeEventListener('mouseup', onMouseUp);
    document.removeEventListener('keydown', onKeyDown);
    document.removeEventListener('keyup', onKeyUp);

    if (tooltip) tooltip.style.display = 'none';
    document.querySelectorAll('.mep-excel-table td.in-autofill-preview').forEach(td => td.classList.remove('in-autofill-preview'));

    isCtrl = upEvent.ctrlKey || isCtrl;
    const { fillDir, pMinR, pMaxR, pMinC, pMaxC } = updatePreview(currentHover.row, currentHover.cIdx, isCtrl, 0, 0);

    if (fillDir === 'none') return;

    pushHistoryState();

    let updatedCount = 0;
    const affectedRows = new Set();

    if (fillDir === 'down' || fillDir === 'up') {
      for (let c = srcMinC; c <= srcMaxC; c++) {
        const colDef = EXCEL_COLUMNS[c];
        if (colDef.isReadOnly || colDef.isFormula) continue;

        if (fillDir === 'down') {
          for (let r = srcMaxR + 1; r <= pMaxR; r++) {
            if (isRowLocked(r) || r === 5) continue;
            const targetRow = SheetState.rows.find(row => row.row === r);
            if (!targetRow) continue;

            const val = computeAutoFillVal(r, c, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, isCtrl);
            if (!targetRow[colDef.col]) targetRow[colDef.col] = {};
            targetRow[colDef.col].val = val;
            affectedRows.add(targetRow);
            updatedCount++;
          }
        } else if (fillDir === 'up') {
          for (let r = srcMinR - 1; r >= pMinR; r--) {
            if (isRowLocked(r) || r === 5) continue;
            const targetRow = SheetState.rows.find(row => row.row === r);
            if (!targetRow) continue;

            const val = computeAutoFillVal(r, c, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, isCtrl);
            if (!targetRow[colDef.col]) targetRow[colDef.col] = {};
            targetRow[colDef.col].val = val;
            affectedRows.add(targetRow);
            updatedCount++;
          }
        }
      }
    } else if (fillDir === 'right' || fillDir === 'left') {
      for (let r = srcMinR; r <= srcMaxR; r++) {
        if (isRowLocked(r) || r === 5) continue;
        const targetRow = SheetState.rows.find(row => row.row === r);
        if (!targetRow) continue;

        if (fillDir === 'right') {
          for (let c = srcMaxC + 1; c <= pMaxC; c++) {
            const colDef = EXCEL_COLUMNS[c];
            if (colDef.isReadOnly || colDef.isFormula) continue;

            const val = computeAutoFillVal(r, c, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, isCtrl);
            if (!targetRow[colDef.col]) targetRow[colDef.col] = {};
            targetRow[colDef.col].val = val;
            affectedRows.add(targetRow);
            updatedCount++;
          }
        } else if (fillDir === 'left') {
          for (let c = srcMinC - 1; c >= pMinC; c--) {
            const colDef = EXCEL_COLUMNS[c];
            if (colDef.isReadOnly || colDef.isFormula) continue;

            const val = computeAutoFillVal(r, c, srcMinR, srcMaxR, srcMinC, srcMaxC, fillDir, isCtrl);
            if (!targetRow[colDef.col]) targetRow[colDef.col] = {};
            targetRow[colDef.col].val = val;
            affectedRows.add(targetRow);
            updatedCount++;
          }
        }
      }
    }

    affectedRows.forEach(row => recalculateRow(row));
    recalculateTotalRow();
    saveSheetData(false);
    renderExcelTable();

    // Select the newly filled expanded range
    const newStartCol = EXCEL_COLUMNS[pMinC].col;
    const newEndCol = EXCEL_COLUMNS[pMaxC].col;
    SheetState.rangeSelection = {
      start: { col: newStartCol, row: pMinR },
      end: { col: newEndCol, row: pMaxR }
    };
    highlightSelectedRange();

    const isSeries = isCtrl || (srcMaxR - srcMinR >= 1) || (srcMaxC - srcMinC >= 1);
    showToast(`✨ AutoFill: ${updatedCount} cells ${isSeries ? 'series completed' : 'copied'} successfully`, 'success');
  };

  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', onMouseUp);
  document.addEventListener('keydown', onKeyDown);
  document.addEventListener('keyup', onKeyUp);
}

// ─── STRICT NUMERIC SANITIZER ────────────────────────────────────────────────
function sanitizeNumericValue(val) {
  if (typeof val !== 'string') return val;
  const trimmed = val.trim();
  if (trimmed.startsWith('=')) {
    try {
      const sanitized = trimmed.substring(1).replace(/[^0-9+\-*/(). ]/g, '');
      if (sanitized) {
        const result = Function(`'use strict'; return (${sanitized})`)();
        if (!isNaN(result) && isFinite(result)) return Math.round(result * 100) / 100;
      }
    } catch (e) {
      // Fallback
    }
  }
  const cleaned = trimmed.replace(/[^0-9.]/g, '');
  return cleaned === '' ? null : (Number(cleaned) || 0);
}

// ─── IN-PLACE CELL EDITING ────────────────────────────────────────────────────
function startCellEdit(td, colLetter, rowNum, colDef, initialChar = null) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('🔒 ভিউ মোড: কোনো ডাটা পরিবর্তন বা এডিট করা যাবে না।', 'warning');
    return;
  }

  if (ACTIVE_TAB === 'summary_downtime' || ACTIVE_TAB === 'summary_production' || rowNum === 5 || colDef.isFormula || colDef.isReadOnly || isRowLocked(rowNum)) {
    return;
  }

  if (colLetter !== 'AM' && initialChar !== null) {
    if (!/^[0-9.=\-+/*()]$/.test(initialChar)) {
      return;
    }
  }

  SheetState.isEditing = true;

  const rowObj = SheetState.rows.find(r => r.row === rowNum);
  const currentVal = rowObj ? (rowObj[colLetter]?.val ?? '') : td.textContent.trim();

  let input;
  if (colLetter === 'AM') {
    input = document.createElement('textarea');
    input.rows = 1;
    input.value = initialChar !== null ? initialChar : currentVal;
    input.className = 'cell-edit-input-remarks';

    const autoGrow = () => {
      input.style.height = 'auto';
      input.style.height = Math.max(26, input.scrollHeight) + 'px';
    };
    input.addEventListener('input', autoGrow);
    setTimeout(autoGrow, 0);
  } else {
    input = document.createElement('input');
    input.type = 'text';
    input.inputMode = 'numeric';
    input.value = initialChar !== null ? initialChar : currentVal;
    input.className = 'cell-edit-input';

    input.addEventListener('keydown', (e) => {
      if (['Backspace', 'Delete', 'Tab', 'Enter', 'Escape', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key) || e.ctrlKey || e.metaKey) {
        return;
      }
      if (!/^[0-9.=\-+/*() ]$/.test(e.key)) {
        e.preventDefault();
      }
    });

    input.addEventListener('input', () => {
      if (!input.value.trim().startsWith('=')) {
        input.value = input.value.replace(/[^0-9.]/g, '');
      }
    });
  }

  td.textContent = '';
  td.appendChild(input);
  SheetState.activeInput = input;

  input.focus();
  if (initialChar === null && input.select) input.select();

  const commit = () => {
    if (!SheetState.isEditing) return;
    SheetState.isEditing = false;
    SheetState.activeInput = null;
    const newVal = input.value.trim();
    saveCellUpdate(colLetter, rowNum, newVal, colDef);
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      commit();
      navigateSelection(0, 1);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      commit();
      navigateSelection(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === 'ArrowUp' && colLetter !== 'AM') {
      e.preventDefault();
      e.stopPropagation();
      commit();
      navigateSelection(0, -1);
    } else if (e.key === 'ArrowDown' && colLetter !== 'AM') {
      e.preventDefault();
      e.stopPropagation();
      commit();
      navigateSelection(0, 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      SheetState.isEditing = false;
      SheetState.activeInput = null;
      updateSingleRowDisplay(rowNum);
      selectCell(colLetter, rowNum);
    }
  });
}

function saveCellUpdate(colLetter, rowNum, newVal, colDef) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    return;
  }

  if (colDef.isReadOnly || colDef.isFormula || isRowLocked(rowNum) || rowNum === 5) {
    return;
  }

  const rowObj = SheetState.rows.find(r => r.row === rowNum);
  if (!rowObj) return;

  pushHistoryState();

  if (!rowObj[colLetter]) rowObj[colLetter] = {};

  if (colLetter === 'AM') {
    rowObj[colLetter].val = newVal === '' ? null : newVal;
  } else {
    const cleanNum = sanitizeNumericValue(newVal);
    rowObj[colLetter].val = cleanNum;
  }

  recalculateRow(rowObj);
  recalculateTotalRow();
  saveSheetData(false);

  updateSingleRowDisplay(rowNum);
  updateTotalRowDisplay();
  selectCell(colLetter, rowNum);
}

function updateSingleRowDisplay(rowNum) {
  const tabInfo = SHEET_TABS[ACTIVE_TAB];
  if (tabInfo && tabInfo.timeGroups) {
    const rowsPerDay = getRowsPerDay();
    const mIdx = (rowNum - 6) % rowsPerDay;
    const dayStart = rowNum - mIdx;
    for (let i = 0; i < rowsPerDay; i++) {
      updateDirectRowDisplay(dayStart + i);
    }
    return;
  }
  updateDirectRowDisplay(rowNum);
}

function updateDirectRowDisplay(rowNum) {
  const rowObj = SheetState.rows.find(r => r.row === rowNum);
  if (!rowObj) return;

  EXCEL_COLUMNS.forEach(colDef => {
    const colLetter = colDef.col;
    const td = document.querySelector(`.mep-excel-table td[data-col="${colLetter}"][data-row="${rowNum}"]`);
    if (!td) return;

    const cellVal = rowObj[colLetter]?.val;
    td.textContent = formatCellValue(cellVal, colDef);

    if (rowObj[colLetter]?.formula) {
      td.dataset.formula = rowObj[colLetter].formula;
    }
  });
}

function updateTotalRowDisplay() {
  EXCEL_COLUMNS.slice(4).forEach(colDef => {
    const colLetter = colDef.col;
    const td = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="5"]`);
    if (!td) return;

    const val = SheetState.totals[colLetter];
    td.textContent = formatCellValue(val, colDef, true);
  });
}

// ─── KEYBOARD NAVIGATION ENGINE (WITH EXCEL-GRADE MERGED CELL AWARENESS) ──────
function getCellSpanInfo(colLetter, rowNum) {
  const rowsPerDay = getRowsPerDay();
  const mIdx = (rowNum - 6) % rowsPerDay;

  if (['A', 'B', 'C'].includes(colLetter)) {
    const masterRow = rowNum - mIdx;
    return { masterRow, span: rowsPerDay, isMerged: rowsPerDay > 1 };
  }

  if (TIME_COLUMNS.includes(colLetter)) {
    const tabInfo = SHEET_TABS[ACTIVE_TAB];
    if (tabInfo && tabInfo.timeGroups) {
      const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);
      const dayStart = rowNum - mIdx;
      const masterRow = dayStart + groupInfo.masterIdx;
      return { masterRow, span: groupInfo.count, isMerged: groupInfo.count > 1 };
    }
  }

  return { masterRow: rowNum, span: 1, isMerged: false };
}

function navigateSelection(dCol, dRow) {
  if (SHEET_TABS[ACTIVE_TAB]?.isSummary) return;

  let cur = SheetState.selected;
  const maxActive = getMaxActiveRow();

  if (!cur || !cur.colLetter || !cur.row) {
    selectCell('E', 6);
    return;
  }

  const colIdx = EXCEL_COLUMNS.findIndex(c => c.col === cur.colLetter);
  let newColIdx = (colIdx >= 0) ? colIdx : 4;
  let newRowNum = cur.row;

  // Horizontal Navigation (Left / Right)
  if (dCol !== 0) {
    if (Math.abs(dCol) >= 50) {
      newColIdx = dCol > 0 ? (EXCEL_COLUMNS.length - 1) : 0;
    } else {
      newColIdx = Math.max(0, Math.min(EXCEL_COLUMNS.length - 1, newColIdx + dCol));
    }
  }

  // Vertical Navigation (Up / Down) with Merged Cell awareness
  if (dRow !== 0) {
    const spanInfo = getCellSpanInfo(cur.colLetter, cur.row);
    
    if (Math.abs(dRow) >= 50) {
      newRowNum = dRow > 0 ? maxActive : 6;
    } else if (dRow > 0) {
      // Moving Down: jump past the merged cell's entire span in 1 single step!
      newRowNum = spanInfo.masterRow + spanInfo.span;
      if (newRowNum > maxActive) newRowNum = maxActive;
    } else if (dRow < 0) {
      // Moving Up: jump to the top master cell above the merged block in 1 single step!
      const targetAbove = spanInfo.masterRow - 1;
      if (targetAbove >= 6) {
        const aboveInfo = getCellSpanInfo(cur.colLetter, targetAbove);
        newRowNum = aboveInfo.masterRow;
      } else {
        newRowNum = 6;
      }
    }
  }

  const newColLetter = EXCEL_COLUMNS[newColIdx].col;
  selectCell(newColLetter, newRowNum);

  // Smooth auto-scroll into view
  const target = document.querySelector(`.mep-excel-table [data-col="${newColLetter}"][data-row="${newRowNum}"]`);
  if (target) {
    target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
}

// ─── ADVANCED QUICK ROW ENTRY MODAL ───────────────────────────────────────────
let modalActiveRow = 6;

function initModalDtInputs() {
  const container = document.getElementById('modalDtContainer');
  if (!container) return;

  container.replaceChildren();

  EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => {
    const card = document.createElement('div');
    card.className = 'entry-dt-card';
    
    const label = document.createElement('label');
    label.className = 'text-[11.5px] text-slate-700 truncate pr-1 font-bold';
    label.textContent = `${c.code}. ${c.label}`;
    label.title = c.label;

    const input = document.createElement('input');
    input.type = 'text';
    input.inputMode = 'numeric';
    input.placeholder = '0';
    input.dataset.dtCol = c.col;
    input.className = 'w-14 text-right border border-slate-300 rounded px-1 py-0.5 text-xs font-bold text-purple-900 outline-none focus:border-purple-600';
    
    input.addEventListener('input', (e) => {
      e.target.value = e.target.value.replace(/[^0-9.]/g, '');
      updateModalLiveKpi();
    });

    card.appendChild(label);
    card.appendChild(input);
    container.appendChild(card);
  });
}

function openQuickEntryModal(rowNum) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('🔒 ভিউ মোড: ডাটা এন্ট্রি বন্ধ রয়েছে। শুধুমাত্র দেখতে পারবেন।', 'warning');
    return;
  }

  if (rowNum === 5) rowNum = 6;
  if (isRowLocked(rowNum)) {
    return;
  }

  modalActiveRow = rowNum;
  const rowObj = SheetState.rows.find(r => r.row === rowNum);
  if (!rowObj) return;

  const modal = document.getElementById('quickEntryModal');
  if (!modal) return;

  setText('modalMachineTitle', `${rowObj.D?.val || 'Machine'} — ${rowObj.A?.val || ''} (${rowObj.B?.val || ''})`);

  const mCap = document.getElementById('mCap');
  const mAct = document.getElementById('mAct');
  const mRej = document.getElementById('mRej');
  const mPlan = document.getElementById('mPlan');
  const mRemarks = document.getElementById('mRemarks');

  if (mCap) mCap.value = rowObj.E?.val ?? '';
  if (mAct) mAct.value = rowObj.F?.val ?? '';
  if (mRej) mRej.value = rowObj.G?.val ?? '';
  if (mPlan) mPlan.value = rowObj.H?.val ?? '';
  if (mRemarks) mRemarks.value = rowObj.AM?.val ?? '';

  document.querySelectorAll('#modalDtContainer input').forEach(input => {
    const col = input.dataset.dtCol;
    input.value = rowObj[col]?.val ?? '';
  });

  updateModalLiveKpi();
  modal.classList.remove('hidden');
  mAct?.focus();
}

function closeQuickEntryModal() {
  const modal = document.getElementById('quickEntryModal');
  if (modal) modal.classList.add('hidden');
}

function updateModalLiveKpi() {
  const cap = Number(document.getElementById('mCap')?.value) || 0;
  const act = Number(document.getElementById('mAct')?.value) || 0;
  const rej = Number(document.getElementById('mRej')?.value) || 0;
  const plan = Number(document.getElementById('mPlan')?.value) || 0;

  let totalDt = 0;
  document.querySelectorAll('#modalDtContainer input').forEach(i => {
    totalDt += Number(i.value) || 0;
  });

  setText('mTotalDtBadge', `Total DT: ${totalDt} min`);

  const runTime = plan > 0 ? Math.max(0, plan - totalDt) : 0;
  const avail = plan > 0 ? (runTime / plan) : 0;
  const perf = cap > 0 ? (act / cap) : 0;
  const qual = (act + rej) > 0 ? (act / (act + rej)) : 0;
  const oee = avail * perf * qual;

  setText('mRunTime', `${runTime} min`);
  setText('mAvail', `${(avail * 100).toFixed(0)}%`);
  setText('mPerf', `${(perf * 100).toFixed(0)}%`);
  setText('mQual', `${(qual * 100).toFixed(0)}%`);
  setText('mOEE', `${(oee * 100).toFixed(0)}%`);
}

function saveModalEntry(andNext = false) {
  if (isRowLocked(modalActiveRow) || modalActiveRow === 5) return;

  const rowObj = SheetState.rows.find(r => r.row === modalActiveRow);
  if (!rowObj) return;

  pushHistoryState();

  const cap = document.getElementById('mCap')?.value.trim();
  const act = document.getElementById('mAct')?.value.trim();
  const rej = document.getElementById('mRej')?.value.trim();
  const plan = document.getElementById('mPlan')?.value.trim();
  const rem = document.getElementById('mRemarks')?.value.trim();

  rowObj.E.val = sanitizeNumericValue(cap);
  rowObj.F.val = sanitizeNumericValue(act);
  rowObj.G.val = sanitizeNumericValue(rej);
  rowObj.H.val = sanitizeNumericValue(plan);
  rowObj.AM.val = rem || '';

  document.querySelectorAll('#modalDtContainer input').forEach(input => {
    const col = input.dataset.dtCol;
    const v = input.value.trim();
    if (!rowObj[col]) rowObj[col] = {};
    rowObj[col].val = sanitizeNumericValue(v);
  });

  recalculateRow(rowObj);
  recalculateTotalRow();
  saveSheetData(false);
  updateSingleRowDisplay(modalActiveRow);
  updateTotalRowDisplay();

  showToast(`✅ Saved ${rowObj.D?.val}`, 'success');

  const nextRow = modalActiveRow + 1;
  const maxActive = getMaxActiveRow();
  if (andNext && nextRow <= maxActive && !isRowLocked(nextRow)) {
    openQuickEntryModal(nextRow);
    selectCell('E', nextRow);
  } else {
    closeQuickEntryModal();
    selectCell('E', modalActiveRow);
  }
}

// ─── SEARCH IN SHEET (Ctrl+F) ─────────────────────────────────────────────────
function openSearchBar() {
  const bar = document.getElementById('floatingSearchBar');
  const input = document.getElementById('searchInput');
  if (!bar || !input) return;

  bar.classList.remove('hidden');
  input.focus();
  input.select();
  performSearch();
}

function closeSearchBar() {
  const bar = document.getElementById('floatingSearchBar');
  if (bar) bar.classList.add('hidden');
}

function performSearch() {
  const q = document.getElementById('searchInput')?.value.trim().toLowerCase();
  SheetState.searchResults = [];
  SheetState.searchIndex = 0;

  if (!q) {
    setText('searchResultCount', '0 of 0');
    return;
  }

  SheetState.rows.forEach(r => {
    if (isRowLocked(r.row)) return;
    EXCEL_COLUMNS.forEach(c => {
      const v = String(r[c.col]?.val ?? '').toLowerCase();
      if (v.includes(q)) {
        SheetState.searchResults.push({ row: r.row, col: c.col });
      }
    });
  });

  const count = SheetState.searchResults.length;
  if (count > 0) {
    setText('searchResultCount', `1 of ${count}`);
    const match = SheetState.searchResults[0];
    selectCell(match.col, match.row);
  } else {
    setText('searchResultCount', '0 of 0');
  }
}

function cycleSearch(next = true) {
  const total = SheetState.searchResults.length;
  if (total === 0) return;

  if (next) {
    SheetState.searchIndex = (SheetState.searchIndex + 1) % total;
  } else {
    SheetState.searchIndex = (SheetState.searchIndex - 1 + total) % total;
  }

  setText('searchResultCount', `${SheetState.searchIndex + 1} of ${total}`);
  const match = SheetState.searchResults[SheetState.searchIndex];
  selectCell(match.col, match.row);
}

// ─── CLIPBOARD COPY & PASTE (EXACT EXCEL BEHAVIOR) ───────────────────────────
function handleClipboardCopy() {
  const start = SheetState.rangeSelection.start || SheetState.selected;
  const end = SheetState.rangeSelection.end || SheetState.selected;
  if (!start || !end) return;

  const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === start.col);
  const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === end.col);
  const minC = Math.min(startCIdx, endCIdx);
  const maxC = Math.max(startCIdx, endCIdx);
  const minR = Math.min(start.row, end.row);
  const maxR = Math.max(start.row, end.row);

  const lines = [];
  document.querySelectorAll('.mep-excel-table td.excel-copied-cell').forEach(td => td.classList.remove('excel-copied-cell'));

  for (let r = minR; r <= maxR; r++) {
    const rowObj = SheetState.rows.find(row => row.row === r);
    const lineVals = [];
    for (let c = minC; c <= maxC; c++) {
      const colLetter = EXCEL_COLUMNS[c].col;
      const v = rowObj ? (rowObj[colLetter]?.val ?? '') : (r === 5 ? SheetState.totals[colLetter] : '');
      lineVals.push(v !== null && v !== undefined ? v : '');

      const td = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="${r}"]`);
      if (td) td.classList.add('excel-copied-cell');
    }
    lines.push(lineVals.join('\t'));
  }

  const tsv = lines.join('\n');
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).then(() => {
      showToast(`📋 Copied ${maxR - minR + 1} × ${maxC - minC + 1} cells`, 'info');
    }).catch(() => {
      fallbackCopyText(tsv);
    });
  } else {
    fallbackCopyText(tsv);
  }
}

function fallbackCopyText(text) {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  try {
    document.execCommand('copy');
    showToast('📋 Copied cells to clipboard', 'info');
  } catch (err) {}
  document.body.removeChild(textarea);
}

function handleClipboardPaste(text) {
  if (!text) return;
  const cur = SheetState.selected;
  if (!cur || cur.row === 5 || isRowLocked(cur.row)) return;

  pushHistoryState();

  const lines = text.split(/\r\n|\n|\r/).filter(l => l.length > 0);
  if (lines.length === 0) return;

  const start = SheetState.rangeSelection.start || cur;
  const end = SheetState.rangeSelection.end || cur;
  const minR = Math.min(start.row, end.row);
  const maxR = Math.min(Math.max(start.row, end.row), getMaxActiveRow());
  const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === (start.col || cur.colLetter));
  const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === (end.col || cur.colLetter));
  const minC = Math.min(startCIdx, endCIdx);
  const maxC = Math.max(startCIdx, endCIdx);

  const isSingleCellCopied = (lines.length === 1 && !lines[0].includes('\t'));
  const isMultiCellRangeSelected = (minR !== maxR || minC !== maxC);

  // If copying 1 single value and multi-cells are selected -> Fill the entire range (Same as Excel!)
  if (isSingleCellCopied && isMultiCellRangeSelected) {
    const singleVal = lines[0].trim();
    let modifiedCount = 0;
    const affectedRows = new Set();

    for (let r = minR; r <= maxR; r++) {
      if (r === 5 || isRowLocked(r)) continue;
      const rowObj = SheetState.rows.find(row => row.row === r);
      if (!rowObj) continue;

      let rowModified = false;
      for (let c = minC; c <= maxC; c++) {
        const colDef = EXCEL_COLUMNS[c];
        if (colDef.isReadOnly || colDef.isFormula) continue;
        if (!rowObj[colDef.col]) rowObj[colDef.col] = {};

        if (colDef.col === 'AM') {
          rowObj[colDef.col].val = singleVal;
        } else {
          rowObj[colDef.col].val = sanitizeNumericValue(singleVal);
        }
        rowModified = true;
        modifiedCount++;
      }
      if (rowModified) {
        recalculateRow(rowObj);
        affectedRows.add(r);
      }
    }

    if (modifiedCount > 0) {
      recalculateTotalRow();
      saveSheetData(false);
      affectedRows.forEach(r => updateSingleRowDisplay(r));
      updateTotalRowDisplay();
      highlightSelectedRange();
      showToast(`📄 Pasted into ${modifiedCount} cells`, 'success');
    }
    return;
  }

  // Multi-cell grid paste
  const maxActive = getMaxActiveRow();
  const affectedRows = new Set();
  let modifiedCount = 0;

  lines.forEach((line, rOffset) => {
    const targetRow = (isMultiCellRangeSelected ? minR : cur.row) + rOffset;
    if (targetRow > maxActive || isRowLocked(targetRow) || targetRow === 5) return;
    const rowObj = SheetState.rows.find(r => r.row === targetRow);
    if (!rowObj) return;

    const values = line.split('\t');
    let rowModified = false;
    values.forEach((val, cOffset) => {
      const targetCIdx = (isMultiCellRangeSelected ? minC : startCIdx) + cOffset;
      if (targetCIdx >= EXCEL_COLUMNS.length) return;
      const colDef = EXCEL_COLUMNS[targetCIdx];
      if (colDef.isFormula || colDef.isReadOnly) return;

      const trimmed = val.trim();
      if (!rowObj[colDef.col]) rowObj[colDef.col] = {};

      if (colDef.col === 'AM') {
        rowObj[colDef.col].val = trimmed;
      } else {
        rowObj[colDef.col].val = sanitizeNumericValue(trimmed);
      }
      rowModified = true;
      modifiedCount++;
    });

    if (rowModified) {
      recalculateRow(rowObj);
      affectedRows.add(targetRow);
    }
  });

  if (modifiedCount > 0) {
    recalculateTotalRow();
    saveSheetData(false);
    affectedRows.forEach(r => updateSingleRowDisplay(r));
    updateTotalRowDisplay();
    highlightSelectedRange();
    showToast(`📄 Pasted ${modifiedCount} cells`, 'success');
  }
}

// ─── MULTI-CELL RANGE DELETE / CLEAR (EXCEL-LIKE BULK DELETE) ──────────────────
function deleteSelectedRange(setZero = false) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('🔒 ভিউ মোড: কোনো ডাটা ডিলিট বা পরিবর্তন করা যাবে না।', 'warning');
    return;
  }

  if (ACTIVE_TAB.startsWith('summary_')) {
    return;
  }

  const start = SheetState.rangeSelection.start || SheetState.selected;
  const end = SheetState.rangeSelection.end || SheetState.selected;
  if (!start || !end) return;

  const startCIdx = EXCEL_COLUMNS.findIndex(c => c.col === start.col);
  const endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === end.col);
  const minC = Math.min(startCIdx, endCIdx);
  const maxC = Math.max(startCIdx, endCIdx);
  const minR = Math.min(start.row, end.row);
  const maxR = Math.min(Math.max(start.row, end.row), getMaxActiveRow());

  pushHistoryState();

  let modifiedCount = 0;
  const affectedRows = new Set();

  for (let r = minR; r <= maxR; r++) {
    if (r === 5 || isRowLocked(r)) continue;
    const rowObj = SheetState.rows.find(row => row.row === r);
    if (!rowObj) continue;

    let rowModified = false;
    for (let c = minC; c <= maxC; c++) {
      const colDef = EXCEL_COLUMNS[c];
      if (colDef.isReadOnly || colDef.isFormula) continue;

      const colLetter = colDef.col;
      if (!rowObj[colLetter]) rowObj[colLetter] = {};

      const newVal = setZero ? 0 : null;
      if (rowObj[colLetter].val !== newVal) {
        rowObj[colLetter].val = newVal;
        rowModified = true;
        modifiedCount++;
      }
    }

    if (rowModified) {
      recalculateRow(rowObj);
      affectedRows.add(r);
    }
  }

  if (modifiedCount > 0) {
    recalculateTotalRow();
    saveSheetData(false);

    affectedRows.forEach(r => updateSingleRowDisplay(r));
    updateTotalRowDisplay();

    // Re-highlight the range selection
    highlightSelectedRange();

    const actionText = setZero ? '0 সেট' : 'মুছে ফেলা';
    showToast(`🧹 ${modifiedCount}টি সেলের ডাটা সফলভাবে ${actionText} হয়েছে!`, 'info');
  }
}

// ─── EXPORT TO EXCEL (PREMIUM STYLED WORKBOOK WITH TIMES NEW ROMAN) ────────────
async function exportExcelFile() {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('🔒 ভিউ মোড: এক্সেল ফাইল ডাউনলোড করার অনুমতি নেই।', 'warning');
    return;
  }

  try {
    saveSheetData(false);
    const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
    const year = MonthYearState.year;

    // 1. High-End Styled Export via ExcelJS
    if (typeof ExcelJS !== 'undefined') {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'MEP FAN LTD.';
      workbook.created = new Date();

      // All users get the 5 Executive Summary Sheets + their assigned Production Entry Sheet(s)
      const summaryTabKeys = ['summary_oee_yearly', 'summary_downtime', 'summary_production', 'summary_status', 'summary_oee'];
      let userProductionTabs = [];

      if (!CurrentUser || CurrentUser.id === 'admin' || CurrentUser.id === 'viewer') {
        userProductionTabs = Object.keys(SHEET_TABS).filter(tabId => !SHEET_TABS[tabId].isSummary);
      } else {
        const allowed = CurrentUser.allowedDepts || [];
        userProductionTabs = allowed.filter(tabId => SHEET_TABS[tabId] && !SHEET_TABS[tabId].isSummary);
      }

      const tabsToExport = [...summaryTabKeys, ...userProductionTabs];

      for (const tabId of tabsToExport) {
        const tabInfo = SHEET_TABS[tabId];
        if (!tabInfo) continue;

        // ─── IF SUMMARY DOWNTIME REPORT WORKSHEET ───
        if (tabId === 'summary_downtime') {
          const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);
          const totalColsCount = 1 + dtCols.length + 4;

          const wsSummary = workbook.addWorksheet(tabInfo.name, {
            views: [{ showGridLines: true, state: 'frozen', ySplit: 4, xSplit: 1 }],
            properties: { tabColor: { argb: tabInfo.colorArgb } }
          });

          // Set column widths
          wsSummary.getColumn(1).width = 25; // Machine No
          for (let c = 2; c <= 1 + dtCols.length; c++) {
            wsSummary.getColumn(c).width = 6.5;
          }
          for (let c = 2 + dtCols.length; c <= totalColsCount; c++) {
            wsSummary.getColumn(c).width = (c === totalColsCount ? 10 : 13);
          }

          // Row 1: MEP FAN LTD. (Banner)
          wsSummary.mergeCells(1, 1, 1, totalColsCount);
          const r1 = wsSummary.getCell('A1');
          r1.value = 'MEP FAN LTD.';
          r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
          r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
          r1.alignment = { horizontal: 'center', vertical: 'middle' };
          wsSummary.getRow(1).height = 32;

          // Row 2: Total Downtime Report (Month Year)
          wsSummary.mergeCells(2, 1, 2, totalColsCount);
          const r2 = wsSummary.getCell('A2');
          r2.value = `Total Downtime Report (${monthName} ${year})`;
          r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
          r2.alignment = { horizontal: 'center', vertical: 'middle' };
          wsSummary.getRow(2).height = 20;

          // Row 3: Vertical headers (Clean layout without number row)
          wsSummary.getRow(3).height = 110;
          const r3Machine = wsSummary.getCell('A3');
          r3Machine.value = 'Machine No.';
          r3Machine.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          r3Machine.alignment = { horizontal: 'center', vertical: 'middle' };
          r3Machine.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7EC8E3' } };
          r3Machine.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

          dtCols.forEach((c, idx) => {
            const cell = wsSummary.getRow(3).getCell(idx + 2);
            cell.value = c.label;
            cell.font = { name: 'Times New Roman', size: 9, bold: true, color: { argb: 'FF000000' } };
            cell.alignment = { textRotation: 90, vertical: 'bottom', horizontal: 'center', wrapText: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F72DE' } };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          });

          const summaryHeaders = [
            'Total Down Time (Mins)',
            'Production Running Time (Mins)',
            'Total Run Time (Mins)',
            'Total Down Time (%)'
          ];
          summaryHeaders.forEach((lbl, idx) => {
            const cell = wsSummary.getRow(3).getCell(2 + dtCols.length + idx);
            cell.value = lbl;
            cell.font = { name: 'Times New Roman', size: 9, bold: true, color: { argb: 'FF000000' } };
            cell.alignment = { textRotation: 90, vertical: 'bottom', horizontal: 'center', wrapText: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF56C5D0' } };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          });

          // Department Rows (Row 4 onwards)
          let currentExcelRow = 4;
          const grandDtSums = {};
          dtCols.forEach(c => grandDtSums[c.col] = 0);
          let grandTotalDown = 0;
          let grandPlanned = 0;
          let grandRunTime = 0;

          getActiveSummaryDepts().forEach(dept => {
            const tabSummary = getTabMonthlySummary(dept.id, year, MonthYearState.monthIndex);
            dtCols.forEach(c => grandDtSums[c.col] += tabSummary.dtSums[c.col]);
            grandTotalDown += tabSummary.totalDownMins;
            grandPlanned += tabSummary.totalPlannedMins;
            grandRunTime += tabSummary.totalRunMins;

            const row = wsSummary.getRow(currentExcelRow);
            row.height = 20;

            const cDept = row.getCell(1);
            cDept.value = dept.name;
            cDept.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDept.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00AEEF' } };
            cDept.alignment = { horizontal: 'left', vertical: 'middle' };
            cDept.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

            dtCols.forEach((c, idx) => {
              const cell = row.getCell(idx + 2);
              const val = tabSummary.dtSums[c.col];
              if (val > 0) {
                cell.value = val;
                cell.numFmt = '#,##0';
              } else {
                cell.value = '-';
              }
              cell.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
              cell.alignment = { horizontal: 'right', vertical: 'middle' };
              cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            // Summary columns
            const cDown = row.getCell(2 + dtCols.length);
            cDown.value = tabSummary.totalDownMins;
            cDown.numFmt = '#,##0';
            cDown.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDown.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE8D0' } };
            cDown.alignment = { horizontal: 'right', vertical: 'middle' };
            cDown.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

            const cPlanned = row.getCell(3 + dtCols.length);
            cPlanned.value = tabSummary.totalPlannedMins;
            cPlanned.numFmt = '#,##0';
            cPlanned.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cPlanned.alignment = { horizontal: 'right', vertical: 'middle' };
            cPlanned.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

            const cRun = row.getCell(4 + dtCols.length);
            cRun.value = tabSummary.totalRunMins;
            cRun.numFmt = '#,##0';
            cRun.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRun.alignment = { horizontal: 'right', vertical: 'middle' };
            cRun.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

            const cPct = row.getCell(5 + dtCols.length);
            cPct.value = tabSummary.dtPercent;
            cPct.numFmt = '0.00%';
            cPct.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF065F46' } };
            cPct.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            cPct.alignment = { horizontal: 'right', vertical: 'middle' };
            cPct.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

            currentExcelRow++;
          });

          // Total (Mins) Row
          const rTotal = wsSummary.getRow(currentExcelRow);
          rTotal.height = 22;

          const cTotTitle = rTotal.getCell(1);
          cTotTitle.value = 'Total (Mins)';
          cTotTitle.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cTotTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
          cTotTitle.alignment = { horizontal: 'center', vertical: 'middle' };
          cTotTitle.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };

          dtCols.forEach((c, idx) => {
            const cell = rTotal.getCell(idx + 2);
            cell.value = grandDtSums[c.col];
            cell.numFmt = '#,##0';
            cell.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
          });

          const cTotDown = rTotal.getCell(2 + dtCols.length);
          cTotDown.value = grandTotalDown;
          cTotDown.numFmt = '#,##0';
          cTotDown.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cTotDown.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
          cTotDown.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotDown.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };

          const cTotPlanned = rTotal.getCell(3 + dtCols.length);
          cTotPlanned.value = grandPlanned;
          cTotPlanned.numFmt = '#,##0';
          cTotPlanned.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cTotPlanned.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
          cTotPlanned.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotPlanned.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };

          const cTotRun = rTotal.getCell(4 + dtCols.length);
          cTotRun.value = grandRunTime;
          cTotRun.numFmt = '#,##0';
          cTotRun.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cTotRun.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
          cTotRun.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRun.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };

          const cTotPct = rTotal.getCell(5 + dtCols.length);
          cTotPct.value = grandPlanned > 0 ? (grandTotalDown / grandPlanned) : 0;
          cTotPct.numFmt = '0.00%';
          cTotPct.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cTotPct.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5128' } };
          cTotPct.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotPct.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };

          continue; // done with summary_downtime tab!
        }

        // ─── IF PRODUCTION OUTPUT SUMMARY REPORT WORKSHEET ───
        if (tabId === 'summary_production') {
          const wsProd = workbook.addWorksheet(tabInfo.name, {
            views: [{ showGridLines: true, state: 'frozen', ySplit: 4, xSplit: 1 }],
            properties: { tabColor: { argb: tabInfo.colorArgb } }
          });

          // Set column widths
          wsProd.getColumn(1).width = 28; // Section Name
          wsProd.getColumn(2).width = 16; // Production Running (hr)
          wsProd.getColumn(3).width = 16; // Machine Capacity (Pcs)
          wsProd.getColumn(4).width = 16; // Production (pcs)
          wsProd.getColumn(5).width = 16; // Rejection (Pcs)
          wsProd.getColumn(6).width = 20; // Standrad Wise Production Output (%)
          wsProd.getColumn(7).width = 16; // Remarks

          // Row 1: MEP FAN LTD. (Banner)
          wsProd.mergeCells('A1:G1');
          const r1 = wsProd.getCell('A1');
          r1.value = 'MEP FAN LTD.';
          r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
          r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
          r1.alignment = { horizontal: 'center', vertical: 'middle' };
          wsProd.getRow(1).height = 32;

          // Row 2: Production Output (Month Year)
          wsProd.mergeCells('A2:G2');
          const r2 = wsProd.getCell('A2');
          r2.value = `Production Output (${monthName} ${year})`;
          r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
          r2.alignment = { horizontal: 'center', vertical: 'middle' };
          wsProd.getRow(2).height = 20;

          // Header Borders & Fill Styles
          const skyHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7EC8E3' } };
          const borderThin = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          // Row 3 & 4: Merged Header Layout
          wsProd.getRow(3).height = 24;
          wsProd.getRow(4).height = 24;

          wsProd.mergeCells('A3:A4');
          const hSec = wsProd.getCell('A3');
          hSec.value = 'Section Name';
          hSec.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hSec.fill = skyHeaderFill;
          hSec.alignment = { horizontal: 'center', vertical: 'middle' };
          hSec.border = borderThin;

          wsProd.mergeCells('B3:B4');
          const hRun = wsProd.getCell('B3');
          hRun.value = 'Production Running (hr)';
          hRun.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hRun.fill = skyHeaderFill;
          hRun.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hRun.border = borderThin;

          wsProd.mergeCells('C3:C4');
          const hCap = wsProd.getCell('C3');
          hCap.value = 'Machine Capacity (Pcs)';
          hCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hCap.fill = skyHeaderFill;
          hCap.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hCap.border = borderThin;

          wsProd.mergeCells('D3:E3');
          const hQty = wsProd.getCell('D3');
          hQty.value = 'Prodcution Qty.';
          hQty.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hQty.fill = skyHeaderFill;
          hQty.alignment = { horizontal: 'center', vertical: 'middle' };
          hQty.border = borderThin;

          wsProd.mergeCells('F3:F4');
          const hOut = wsProd.getCell('F3');
          hOut.value = 'Standrad Wise Production Output (%)';
          hOut.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hOut.fill = skyHeaderFill;
          hOut.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hOut.border = borderThin;

          wsProd.mergeCells('G3:G4');
          const hRem = wsProd.getCell('G3');
          hRem.value = 'Remarks';
          hRem.font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
          hRem.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
          hRem.alignment = { horizontal: 'center', vertical: 'middle' };
          hRem.border = borderThin;

          // Row 4 Sub-headers
          const hProdPcs = wsProd.getCell('D4');
          hProdPcs.value = 'Production (pcs)';
          hProdPcs.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hProdPcs.fill = skyHeaderFill;
          hProdPcs.alignment = { horizontal: 'center', vertical: 'middle' };
          hProdPcs.border = borderThin;

          const hRejPcs = wsProd.getCell('E4');
          hRejPcs.value = 'Rejection (Pcs)';
          hRejPcs.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hRejPcs.fill = skyHeaderFill;
          hRejPcs.alignment = { horizontal: 'center', vertical: 'middle' };
          hRejPcs.border = borderThin;

          // Apply border styling to merged companions
          ['A4', 'B4', 'C4', 'E3', 'F4', 'G4'].forEach(addr => {
            wsProd.getCell(addr).border = borderThin;
          });

          // Department Rows (Row 5 onwards)
          let currentExcelRow = 5;
          let grandRunning = 0;
          let grandCapacity = 0;
          let grandProduction = 0;
          let grandRejection = 0;

          getActiveSummaryDepts().forEach(dept => {
            const tabSummary = getTabProductionOutputSummary(dept.id, year, MonthYearState.monthIndex);
            grandRunning += tabSummary.runningMins;
            grandCapacity += tabSummary.capacityPcs;
            grandProduction += tabSummary.actualPrdPcs;
            grandRejection += tabSummary.rejectionPcs;

            const row = wsProd.getRow(currentExcelRow);
            row.height = 20;

            const cSec = row.getCell(1);
            cSec.value = dept.name;
            cSec.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8DDFE' } };
            cSec.alignment = { horizontal: 'left', vertical: 'middle' };
            cSec.border = borderThin;

            const cRun = row.getCell(2);
            cRun.value = tabSummary.runningMins;
            cRun.numFmt = '#,##0';
            cRun.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRun.alignment = { horizontal: 'right', vertical: 'middle' };
            cRun.border = borderThin;

            const cCap = row.getCell(3);
            cCap.value = tabSummary.capacityPcs;
            cCap.numFmt = '#,##0';
            cCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cCap.alignment = { horizontal: 'right', vertical: 'middle' };
            cCap.border = borderThin;

            const cProd = row.getCell(4);
            cProd.value = tabSummary.actualPrdPcs;
            cProd.numFmt = '#,##0';
            cProd.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cProd.alignment = { horizontal: 'right', vertical: 'middle' };
            cProd.border = borderThin;

            const cRej = row.getCell(5);
            cRej.value = tabSummary.rejectionPcs;
            cRej.numFmt = '#,##0';
            cRej.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRej.alignment = { horizontal: 'right', vertical: 'middle' };
            cRej.border = borderThin;

            const cPct = row.getCell(6);
            cPct.value = tabSummary.outputPct;
            cPct.numFmt = '0.00%';
            cPct.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cPct.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F72DE' } };
            cPct.alignment = { horizontal: 'right', vertical: 'middle' };
            cPct.border = borderThin;

            const cRem = row.getCell(7);
            cRem.value = '';
            cRem.border = borderThin;

            currentExcelRow++;
          });

          // Total Row (Yellow/Gold)
          const rTotal = wsProd.getRow(currentExcelRow);
          rTotal.height = 22;
          const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
          const borderThickBottom = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          const cTotTitle = rTotal.getCell(1);
          cTotTitle.value = 'Total';
          cTotTitle.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotTitle.fill = yellowFill;
          cTotTitle.alignment = { horizontal: 'center', vertical: 'middle' };
          cTotTitle.border = borderThickBottom;

          const cTotRun = rTotal.getCell(2);
          cTotRun.value = grandRunning;
          cTotRun.numFmt = '#,##0';
          cTotRun.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotRun.fill = yellowFill;
          cTotRun.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRun.border = borderThickBottom;

          const cTotCap = rTotal.getCell(3);
          cTotCap.value = grandCapacity;
          cTotCap.numFmt = '#,##0';
          cTotCap.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotCap.fill = yellowFill;
          cTotCap.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotCap.border = borderThickBottom;

          const cTotProd = rTotal.getCell(4);
          cTotProd.value = grandProduction;
          cTotProd.numFmt = '#,##0';
          cTotProd.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotProd.fill = yellowFill;
          cTotProd.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotProd.border = borderThickBottom;

          const cTotRej = rTotal.getCell(5);
          cTotRej.value = grandRejection;
          cTotRej.numFmt = '#,##0';
          cTotRej.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotRej.fill = yellowFill;
          cTotRej.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRej.border = borderThickBottom;

          const cTotPct = rTotal.getCell(6);
          cTotPct.value = grandCapacity > 0 ? (grandProduction / grandCapacity) : 0;
          cTotPct.numFmt = '0.00%';
          cTotPct.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotPct.fill = yellowFill;
          cTotPct.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotPct.border = borderThickBottom;

          const cTotRem = rTotal.getCell(7);
          cTotRem.value = '';
          cTotRem.fill = yellowFill;
          cTotRem.border = borderThickBottom;

          continue; // done with summary_production tab!
        }

        // ─── IF DOWNTIME & RUNNING STATUS REPORT WORKSHEET ───
        if (tabId === 'summary_status') {
          const wsStatus = workbook.addWorksheet(tabInfo.name, {
            views: [{ showGridLines: true, state: 'frozen', ySplit: 3, xSplit: 1 }],
            properties: { tabColor: { argb: tabInfo.colorArgb } }
          });

          // Set column widths
          wsStatus.getColumn(1).width = 28; // Machine No.
          wsStatus.getColumn(2).width = 18; // Planned Production Time (mins)
          wsStatus.getColumn(3).width = 18; // Production Run Time (Mins)
          wsStatus.getColumn(4).width = 18; // Machine Down Time (Mins)
          wsStatus.getColumn(5).width = 18; // Production Running Time (%)
          wsStatus.getColumn(6).width = 18; // Production Down Time (%)
          wsStatus.getColumn(7).width = 16; // Remark's

          // Row 1: MEP FAN LTD. (Banner)
          wsStatus.mergeCells('A1:G1');
          const r1 = wsStatus.getCell('A1');
          r1.value = 'MEP FAN LTD.';
          r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
          r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
          r1.alignment = { horizontal: 'center', vertical: 'middle' };
          wsStatus.getRow(1).height = 32;

          // Row 2: Sub-banner
          wsStatus.mergeCells('A2:G2');
          const r2 = wsStatus.getCell('A2');
          r2.value = `MEP Fan Limited- Down Time & Running Time Status (${monthName} ${year})`;
          r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
          r2.alignment = { horizontal: 'center', vertical: 'middle' };
          wsStatus.getRow(2).height = 20;

          // Header Styles (Sage Green)
          const sageHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD5E8D4' } };
          const borderThin = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          // Row 3: Headers
          wsStatus.getRow(3).height = 35;
          const statusHeadersDef = [
            { col: 1, label: 'Machine No.', isRed: false },
            { col: 2, label: 'Planned Production Time (mins)', isRed: false },
            { col: 3, label: 'Production Run Time (Mins)', isRed: false },
            { col: 4, label: 'Machine Down Time (Mins)', isRed: false },
            { col: 5, label: 'Production Running Time (%)', isRed: false },
            { col: 6, label: 'Production Down Time (%)', isRed: false },
            { col: 7, label: "Remark's", isRed: true }
          ];

          statusHeadersDef.forEach(h => {
            const cell = wsStatus.getRow(3).getCell(h.col);
            cell.value = h.label;
            cell.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: h.isRed ? 'FFFFFFFF' : 'FF000000' } };
            cell.fill = h.isRed ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } } : sageHeaderFill;
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = borderThin;
          });

          // Department Rows (Row 4 onwards)
          let currentExcelRow = 4;
          let grandPlanned = 0;
          let grandRun = 0;
          let grandDown = 0;

          getActiveSummaryDepts().forEach(dept => {
            const status = getTabDowntimeRunningStatus(dept.id, year, MonthYearState.monthIndex);
            grandPlanned += status.plannedTimeMins;
            grandRun += status.runTimeMins;
            grandDown += status.downTimeMins;

            const row = wsStatus.getRow(currentExcelRow);
            row.height = 20;

            const cName = row.getCell(1);
            cName.value = dept.name;
            cName.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cName.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
            cName.alignment = { horizontal: 'left', vertical: 'middle' };
            cName.border = borderThin;

            const cPlanned = row.getCell(2);
            cPlanned.value = status.plannedTimeMins;
            cPlanned.numFmt = '#,##0';
            cPlanned.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cPlanned.alignment = { horizontal: 'right', vertical: 'middle' };
            cPlanned.border = borderThin;

            const cRun = row.getCell(3);
            cRun.value = status.runTimeMins;
            cRun.numFmt = '#,##0';
            cRun.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRun.alignment = { horizontal: 'right', vertical: 'middle' };
            cRun.border = borderThin;

            const cDown = row.getCell(4);
            cDown.value = status.downTimeMins;
            cDown.numFmt = '#,##0';
            cDown.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDown.alignment = { horizontal: 'right', vertical: 'middle' };
            cDown.border = borderThin;

            const cRunPct = row.getCell(5);
            cRunPct.value = status.runTimePct;
            cRunPct.numFmt = '0.0%';
            cRunPct.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRunPct.alignment = { horizontal: 'right', vertical: 'middle' };
            cRunPct.border = borderThin;

            const cDownPct = row.getCell(6);
            cDownPct.value = status.downTimePct;
            cDownPct.numFmt = '0.0%';
            cDownPct.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDownPct.alignment = { horizontal: 'right', vertical: 'middle' };
            cDownPct.border = borderThin;

            const cRem = row.getCell(7);
            cRem.value = '';
            cRem.border = borderThin;

            currentExcelRow++;
          });

          // SubTotal Row
          const rTotal = wsStatus.getRow(currentExcelRow);
          rTotal.height = 22;
          const subtotalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
          const borderThickBottom = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          const cTotTitle = rTotal.getCell(1);
          cTotTitle.value = 'SubTotal';
          cTotTitle.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotTitle.fill = subtotalFill;
          cTotTitle.alignment = { horizontal: 'left', vertical: 'middle' };
          cTotTitle.border = borderThickBottom;

          const cTotPlanned = rTotal.getCell(2);
          cTotPlanned.value = grandPlanned;
          cTotPlanned.numFmt = '#,##0';
          cTotPlanned.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotPlanned.fill = subtotalFill;
          cTotPlanned.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotPlanned.border = borderThickBottom;

          const cTotRun = rTotal.getCell(3);
          cTotRun.value = grandRun;
          cTotRun.numFmt = '#,##0';
          cTotRun.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotRun.fill = subtotalFill;
          cTotRun.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRun.border = borderThickBottom;

          const cTotDown = rTotal.getCell(4);
          cTotDown.value = grandDown;
          cTotDown.numFmt = '#,##0';
          cTotDown.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotDown.fill = subtotalFill;
          cTotDown.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotDown.border = borderThickBottom;

          const cTotRunPct = rTotal.getCell(5);
          cTotRunPct.value = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
          cTotRunPct.numFmt = '0.0%';
          cTotRunPct.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotRunPct.fill = subtotalFill;
          cTotRunPct.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRunPct.border = borderThickBottom;

          const cTotDownPct = rTotal.getCell(6);
          cTotDownPct.value = grandPlanned > 0 ? (grandDown / grandPlanned) : 0;
          cTotDownPct.numFmt = '0.0%';
          cTotDownPct.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotDownPct.fill = subtotalFill;
          cTotDownPct.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotDownPct.border = borderThickBottom;

          const cTotRem = rTotal.getCell(7);
          cTotRem.value = '';
          cTotRem.fill = subtotalFill;
          cTotRem.border = borderThickBottom;

          continue; // done with summary_status tab!
        }

        // ─── IF OEE SUMMARY REPORT WORKSHEET ───
        if (tabId === 'summary_oee') {
          const wsOEE = workbook.addWorksheet(tabInfo.name, {
            views: [{ showGridLines: true, state: 'frozen', ySplit: 4, xSplit: 1 }],
            properties: { tabColor: { argb: tabInfo.colorArgb } }
          });

          // Set column widths
          wsOEE.getColumn(1).width = 28; // Section
          wsOEE.getColumn(2).width = 18; // Machine Capacity (Pcs)
          wsOEE.getColumn(3).width = 18; // Total Production (pcs)
          wsOEE.getColumn(4).width = 16; // Rejection (Pcs)
          wsOEE.getColumn(5).width = 16; // Availability (%)
          wsOEE.getColumn(6).width = 16; // Performance (%)
          wsOEE.getColumn(7).width = 16; // Quality (%)
          wsOEE.getColumn(8).width = 16; // OEE (%)
          wsOEE.getColumn(9).width = 16; // Remark's

          // Row 1: MEP FAN LTD. (Banner)
          wsOEE.mergeCells('A1:I1');
          const r1 = wsOEE.getCell('A1');
          r1.value = 'MEP FAN LTD.';
          r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
          r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
          r1.alignment = { horizontal: 'center', vertical: 'middle' };
          wsOEE.getRow(1).height = 32;

          // Row 2: Sub-banner
          wsOEE.mergeCells('A2:I2');
          const r2 = wsOEE.getCell('A2');
          r2.value = `OEE Report (${monthName} ${year})`;
          r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
          r2.alignment = { horizontal: 'center', vertical: 'middle' };
          wsOEE.getRow(2).height = 20;

          // Header Styles
          const skyHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7EC8E3' } };
          const borderThin = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          // Row 3 & 4: Merged Header Layout
          wsOEE.getRow(3).height = 24;
          wsOEE.getRow(4).height = 24;

          wsOEE.mergeCells('A3:A4');
          const hSec = wsOEE.getCell('A3');
          hSec.value = 'Section';
          hSec.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hSec.fill = skyHeaderFill;
          hSec.alignment = { horizontal: 'center', vertical: 'middle' };
          hSec.border = borderThin;

          wsOEE.mergeCells('B3:B4');
          const hCap = wsOEE.getCell('B3');
          hCap.value = 'Machine Capacity (Pcs)';
          hCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hCap.fill = skyHeaderFill;
          hCap.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hCap.border = borderThin;

          wsOEE.mergeCells('C3:D3');
          const hQty = wsOEE.getCell('C3');
          hQty.value = 'Production Qty.';
          hQty.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hQty.fill = skyHeaderFill;
          hQty.alignment = { horizontal: 'center', vertical: 'middle' };
          hQty.border = borderThin;

          wsOEE.mergeCells('E3:E4');
          const hAvail = wsOEE.getCell('E3');
          hAvail.value = 'Availability (%)';
          hAvail.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hAvail.fill = skyHeaderFill;
          hAvail.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hAvail.border = borderThin;

          wsOEE.mergeCells('F3:F4');
          const hPerf = wsOEE.getCell('F3');
          hPerf.value = 'Performance (%)';
          hPerf.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hPerf.fill = skyHeaderFill;
          hPerf.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hPerf.border = borderThin;

          wsOEE.mergeCells('G3:G4');
          const hQual = wsOEE.getCell('G3');
          hQual.value = 'Quality (%)';
          hQual.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hQual.fill = skyHeaderFill;
          hQual.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hQual.border = borderThin;

          wsOEE.mergeCells('H3:H4');
          const hOEE = wsOEE.getCell('H3');
          hOEE.value = 'OEE (%)';
          hOEE.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hOEE.fill = skyHeaderFill;
          hOEE.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hOEE.border = borderThin;

          wsOEE.mergeCells('I3:I4');
          const hRem = wsOEE.getCell('I3');
          hRem.value = "Remark's";
          hRem.font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
          hRem.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
          hRem.alignment = { horizontal: 'center', vertical: 'middle' };
          hRem.border = borderThin;

          // Row 4 Sub-headers
          const hTotProd = wsOEE.getCell('C4');
          hTotProd.value = 'Total Production (pcs)';
          hTotProd.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hTotProd.fill = skyHeaderFill;
          hTotProd.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hTotProd.border = borderThin;

          const hRej = wsOEE.getCell('D4');
          hRej.value = 'Rejection (Pcs)';
          hRej.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hRej.fill = skyHeaderFill;
          hRej.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hRej.border = borderThin;

          ['A4', 'B4', 'D3', 'E4', 'F4', 'G4', 'H4', 'I4'].forEach(addr => {
            wsOEE.getCell(addr).border = borderThin;
          });

          // Department Rows (Row 5 onwards)
          let currentExcelRow = 5;
          let grandCap = 0;
          let grandProd = 0;
          let grandRej = 0;
          let grandPlanned = 0;
          let grandRun = 0;

          getActiveSummaryDepts().forEach(dept => {
            const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);
            grandCap += oeeData.capacityPcs;
            grandProd += oeeData.totalProduction;
            grandRej += oeeData.rejectionPcs;
            grandPlanned += oeeData.plannedTimeMins;
            grandRun += oeeData.runTimeMins;

            const row = wsOEE.getRow(currentExcelRow);
            row.height = 20;

            const cSec = row.getCell(1);
            cSec.value = dept.name;
            cSec.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8A9EE' } };
            cSec.alignment = { horizontal: 'left', vertical: 'middle' };
            cSec.border = borderThin;

            const cCap = row.getCell(2);
            cCap.value = oeeData.capacityPcs;
            cCap.numFmt = '#,##0';
            cCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cCap.alignment = { horizontal: 'right', vertical: 'middle' };
            cCap.border = borderThin;

            const cProd = row.getCell(3);
            cProd.value = oeeData.totalProduction;
            cProd.numFmt = '#,##0';
            cProd.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cProd.alignment = { horizontal: 'right', vertical: 'middle' };
            cProd.border = borderThin;

            const cRej = row.getCell(4);
            cRej.value = oeeData.rejectionPcs;
            cRej.numFmt = '#,##0';
            cRej.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRej.alignment = { horizontal: 'right', vertical: 'middle' };
            cRej.border = borderThin;

            const cAvail = row.getCell(5);
            cAvail.value = oeeData.availability;
            cAvail.numFmt = '0%';
            cAvail.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cAvail.alignment = { horizontal: 'right', vertical: 'middle' };
            cAvail.border = borderThin;

            const cPerf = row.getCell(6);
            cPerf.value = oeeData.performance;
            cPerf.numFmt = '0%';
            cPerf.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cPerf.alignment = { horizontal: 'right', vertical: 'middle' };
            cPerf.border = borderThin;

            const cQual = row.getCell(7);
            cQual.value = oeeData.quality;
            cQual.numFmt = '0.0%';
            cQual.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cQual.alignment = { horizontal: 'right', vertical: 'middle' };
            cQual.border = borderThin;

            const cOEE = row.getCell(8);
            cOEE.value = oeeData.oee;
            cOEE.numFmt = '0.0%';
            cOEE.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cOEE.alignment = { horizontal: 'right', vertical: 'middle' };
            cOEE.border = borderThin;

            const cRem = row.getCell(9);
            cRem.value = '';
            cRem.border = borderThin;

            currentExcelRow++;
          });

          // Total Row (Peach / Soft Gold)
          const rTotal = wsOEE.getRow(currentExcelRow);
          rTotal.height = 22;
          const peachFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8CBAD' } };
          const borderThickBottom = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          const cTotTitle = rTotal.getCell(1);
          cTotTitle.value = 'Total';
          cTotTitle.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotTitle.fill = peachFill;
          cTotTitle.alignment = { horizontal: 'left', vertical: 'middle' };
          cTotTitle.border = borderThickBottom;

          const cTotCap = rTotal.getCell(2);
          cTotCap.value = grandCap;
          cTotCap.numFmt = '#,##0';
          cTotCap.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotCap.fill = peachFill;
          cTotCap.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotCap.border = borderThickBottom;

          const cTotProd = rTotal.getCell(3);
          cTotProd.value = grandProd;
          cTotProd.numFmt = '#,##0';
          cTotProd.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotProd.fill = peachFill;
          cTotProd.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotProd.border = borderThickBottom;

          const cTotRej = rTotal.getCell(4);
          cTotRej.value = grandRej;
          cTotRej.numFmt = '#,##0';
          cTotRej.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotRej.fill = peachFill;
          cTotRej.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotRej.border = borderThickBottom;

          const grandAvail = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
          const grandPerf = grandCap > 0 ? (grandProd / grandCap) : 0;
          const grandQual = (grandProd + grandRej) > 0 ? (grandProd / (grandProd + grandRej)) : 0;
          const grandOEE = grandAvail * grandPerf * grandQual;

          const cTotAvail = rTotal.getCell(5);
          cTotAvail.value = grandAvail;
          cTotAvail.numFmt = '0%';
          cTotAvail.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotAvail.fill = peachFill;
          cTotAvail.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotAvail.border = borderThickBottom;

          const cTotPerf = rTotal.getCell(6);
          cTotPerf.value = grandPerf;
          cTotPerf.numFmt = '0%';
          cTotPerf.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotPerf.fill = peachFill;
          cTotPerf.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotPerf.border = borderThickBottom;

          const cTotQual = rTotal.getCell(7);
          cTotQual.value = grandQual;
          cTotQual.numFmt = '0.0%';
          cTotQual.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotQual.fill = peachFill;
          cTotQual.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotQual.border = borderThickBottom;

          const cTotOEE = rTotal.getCell(8);
          cTotOEE.value = grandOEE;
          cTotOEE.numFmt = '0.0%';
          cTotOEE.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          cTotOEE.fill = peachFill;
          cTotOEE.alignment = { horizontal: 'right', vertical: 'middle' };
          cTotOEE.border = borderThickBottom;

          const cTotRem = rTotal.getCell(9);
          cTotRem.value = '';
          cTotRem.fill = peachFill;
          cTotRem.border = borderThickBottom;

          continue; // done with summary_oee tab!
        }

        // ─── IF SUMMARY OF OEE (YEARLY 12-MONTH REPORT) WORKSHEET ───
        if (tabId === 'summary_oee_yearly') {
          const wsYearly = workbook.addWorksheet(tabInfo.name, {
            views: [{ showGridLines: true, state: 'frozen', ySplit: 4, xSplit: 2 }],
            properties: { tabColor: { argb: tabInfo.colorArgb } }
          });

          // Set column widths
          wsYearly.getColumn(1).width = 16; // Month
          wsYearly.getColumn(2).width = 24; // Detail
          wsYearly.getColumn(3).width = 18; // Machine Capacity (Pcs)
          wsYearly.getColumn(4).width = 18; // Total Production (pcs)
          wsYearly.getColumn(5).width = 15; // Rejection (Pcs)
          wsYearly.getColumn(6).width = 15; // Availability (%)
          wsYearly.getColumn(7).width = 15; // Performance (%)
          wsYearly.getColumn(8).width = 15; // Quality (%)
          wsYearly.getColumn(9).width = 15; // OEE (%)
          wsYearly.getColumn(10).width = 16; // Remark's

          // Row 1: MEP FAN LTD. (Banner)
          wsYearly.mergeCells('A1:J1');
          const r1 = wsYearly.getCell('A1');
          r1.value = 'MEP FAN LTD.';
          r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
          r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
          r1.alignment = { horizontal: 'center', vertical: 'middle' };
          wsYearly.getRow(1).height = 32;

          // Row 2: Sub-banner
          wsYearly.mergeCells('A2:J2');
          const r2 = wsYearly.getCell('A2');
          r2.value = `Summary of OEE (${year})`;
          r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
          r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
          r2.alignment = { horizontal: 'center', vertical: 'middle' };
          wsYearly.getRow(2).height = 20;

          // Header Styles
          const skyHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7EC8E3' } };
          const borderThin = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          // Row 3 & 4: Merged Header Layout
          wsYearly.getRow(3).height = 24;
          wsYearly.getRow(4).height = 24;

          wsYearly.mergeCells('A3:A4');
          const hMonth = wsYearly.getCell('A3');
          hMonth.value = 'Month';
          hMonth.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hMonth.fill = skyHeaderFill;
          hMonth.alignment = { horizontal: 'center', vertical: 'middle' };
          hMonth.border = borderThin;

          wsYearly.mergeCells('B3:B4');
          const hDetail = wsYearly.getCell('B3');
          hDetail.value = 'Detail';
          hDetail.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hDetail.fill = skyHeaderFill;
          hDetail.alignment = { horizontal: 'center', vertical: 'middle' };
          hDetail.border = borderThin;

          wsYearly.mergeCells('C3:C4');
          const hCap = wsYearly.getCell('C3');
          hCap.value = 'Machine Capacity (Pcs)';
          hCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hCap.fill = skyHeaderFill;
          hCap.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hCap.border = borderThin;

          wsYearly.mergeCells('D3:E3');
          const hQty = wsYearly.getCell('D3');
          hQty.value = 'Production Qty.';
          hQty.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hQty.fill = skyHeaderFill;
          hQty.alignment = { horizontal: 'center', vertical: 'middle' };
          hQty.border = borderThin;

          wsYearly.mergeCells('F3:F4');
          const hAvail = wsYearly.getCell('F3');
          hAvail.value = 'Availability (%)';
          hAvail.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hAvail.fill = skyHeaderFill;
          hAvail.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hAvail.border = borderThin;

          wsYearly.mergeCells('G3:G4');
          const hPerf = wsYearly.getCell('G3');
          hPerf.value = 'Performance (%)';
          hPerf.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hPerf.fill = skyHeaderFill;
          hPerf.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hPerf.border = borderThin;

          wsYearly.mergeCells('H3:H4');
          const hQual = wsYearly.getCell('H3');
          hQual.value = 'Quality (%)';
          hQual.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hQual.fill = skyHeaderFill;
          hQual.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hQual.border = borderThin;

          wsYearly.mergeCells('I3:I4');
          const hOEE = wsYearly.getCell('I3');
          hOEE.value = 'OEE (%)';
          hOEE.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hOEE.fill = skyHeaderFill;
          hOEE.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hOEE.border = borderThin;

          wsYearly.mergeCells('J3:J4');
          const hRem = wsYearly.getCell('J3');
          hRem.value = "Remark's";
          hRem.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
          hRem.fill = skyHeaderFill;
          hRem.alignment = { horizontal: 'center', vertical: 'middle' };
          hRem.border = borderThin;

          // Row 4 Sub-headers
          const hTotProd = wsYearly.getCell('D4');
          hTotProd.value = 'Total Production (pcs)';
          hTotProd.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hTotProd.fill = skyHeaderFill;
          hTotProd.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hTotProd.border = borderThin;

          const hRej = wsYearly.getCell('E4');
          hRej.value = 'Rejection (Pcs)';
          hRej.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
          hRej.fill = skyHeaderFill;
          hRej.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          hRej.border = borderThin;

          ['A4', 'B4', 'C4', 'E3', 'F4', 'G4', 'H4', 'I4', 'J4'].forEach(addr => {
            wsYearly.getCell(addr).border = borderThin;
          });

          // 12 Months Rows (2 rows per month)
          const yearlyData = getYearlyOEESummary(year);
          let currentExcelRow = 5;

          yearlyData.forEach(m => {
            const row1Num = currentExcelRow;
            const row2Num = currentExcelRow + 1;
            const row1 = wsYearly.getRow(row1Num);
            const row2 = wsYearly.getRow(row2Num);
            row1.height = 20;
            row2.height = 20;

            const borderMonthDivider = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'medium', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } }
            };

            // Col 1: Month (merge row1:row2)
            wsYearly.mergeCells(`A${row1Num}:A${row2Num}`);
            const cMonth = wsYearly.getCell(`A${row1Num}`);
            cMonth.value = m.monthName;
            cMonth.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
            cMonth.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
            cMonth.alignment = { horizontal: 'center', vertical: 'middle' };
            cMonth.border = borderMonthDivider;
            wsYearly.getCell(`A${row2Num}`).border = borderMonthDivider;

            const hasData = (m.capacityPcs > 0 || m.totalProduction > 0 || m.isFixed);

            // Row 1: Total
            const cDet1 = row1.getCell(2);
            cDet1.value = 'Total';
            cDet1.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDet1.alignment = { horizontal: 'center', vertical: 'middle' };
            cDet1.border = borderThin;

            const cCap = row1.getCell(3);
            if (hasData && m.capacityPcs > 0) {
              cCap.value = m.capacityPcs;
              cCap.numFmt = '#,##0';
              cCap.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
              cCap.value = '-';
              cCap.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            cCap.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cCap.border = borderThin;

            const cProd = row1.getCell(4);
            if (hasData && m.totalProduction > 0) {
              cProd.value = m.totalProduction;
              cProd.numFmt = '#,##0';
              cProd.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
              cProd.value = '-';
              cProd.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            cProd.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cProd.border = borderThin;

            const cRej = row1.getCell(5);
            if (hasData && m.rejectionPcs > 0) {
              cRej.value = m.rejectionPcs;
              cRej.numFmt = '#,##0';
              cRej.alignment = { horizontal: 'right', vertical: 'middle' };
            } else if (hasData) {
              cRej.value = 0;
              cRej.numFmt = '#,##0';
              cRej.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
              cRej.value = 0;
              cRej.numFmt = '#,##0';
              cRej.alignment = { horizontal: 'right', vertical: 'middle' };
            }
            cRej.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cRej.border = borderThin;

            // Merged KPI Columns (Cols 6 to 10)
            wsYearly.mergeCells(`F${row1Num}:F${row2Num}`);
            const cAvail = wsYearly.getCell(`F${row1Num}`);
            if (hasData || m.plannedTimeMins > 0) {
              cAvail.value = m.availability;
              cAvail.numFmt = '0%';
            } else {
              cAvail.value = '-';
            }
            cAvail.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cAvail.alignment = { horizontal: 'center', vertical: 'middle' };
            cAvail.border = borderMonthDivider;
            wsYearly.getCell(`F${row2Num}`).border = borderMonthDivider;

            wsYearly.mergeCells(`G${row1Num}:G${row2Num}`);
            const cPerf = wsYearly.getCell(`G${row1Num}`);
            if (hasData) {
              cPerf.value = m.performance;
              cPerf.numFmt = '0%';
            } else {
              cPerf.value = '-';
            }
            cPerf.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cPerf.alignment = { horizontal: 'center', vertical: 'middle' };
            cPerf.border = borderMonthDivider;
            wsYearly.getCell(`G${row2Num}`).border = borderMonthDivider;

            wsYearly.mergeCells(`H${row1Num}:H${row2Num}`);
            const cQual = wsYearly.getCell(`H${row1Num}`);
            if (hasData) {
              cQual.value = m.quality;
              cQual.numFmt = '0.0%';
            } else {
              cQual.value = '-';
            }
            cQual.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cQual.alignment = { horizontal: 'center', vertical: 'middle' };
            cQual.border = borderMonthDivider;
            wsYearly.getCell(`H${row2Num}`).border = borderMonthDivider;

            wsYearly.mergeCells(`I${row1Num}:I${row2Num}`);
            const cOEE = wsYearly.getCell(`I${row1Num}`);
            if (hasData) {
              cOEE.value = m.oee;
              cOEE.numFmt = '0%';
            } else {
              cOEE.value = '-';
            }
            cOEE.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cOEE.alignment = { horizontal: 'center', vertical: 'middle' };
            cOEE.border = borderMonthDivider;
            wsYearly.getCell(`I${row2Num}`).border = borderMonthDivider;

            wsYearly.mergeCells(`J${row1Num}:J${row2Num}`);
            const cRem = wsYearly.getCell(`J${row1Num}`);
            cRem.value = '';
            cRem.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
            cRem.border = borderMonthDivider;
            wsYearly.getCell(`J${row2Num}`).border = borderMonthDivider;

            // Row 2: Total Acheivement (%)
            const cDet2 = row2.getCell(2);
            cDet2.value = 'Total Acheivement (%)';
            cDet2.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };
            cDet2.alignment = { horizontal: 'center', vertical: 'middle' };
            cDet2.border = borderMonthDivider;

            wsYearly.mergeCells(`C${row2Num}:E${row2Num}`);
            const cAch = wsYearly.getCell(`C${row2Num}`);
            if (hasData && m.capacityPcs > 0) {
              cAch.value = m.achievement;
              cAch.numFmt = '0.0%';
              const isHighlight = m.achievement >= 0.95;
              if (isHighlight) {
                cAch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF38B6FF' } };
              }
            } else {
              cAch.value = '-';
            }
            cAch.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FF000000' } };
            cAch.alignment = { horizontal: 'center', vertical: 'middle' };
            cAch.border = borderMonthDivider;
            wsYearly.getCell(`D${row2Num}`).border = borderMonthDivider;
            currentExcelRow += 2;

            // Gap / Separator row after each month
            const rGap = wsYearly.getRow(currentExcelRow);
            rGap.height = 8;
            wsYearly.mergeCells(`A${currentExcelRow}:J${currentExcelRow}`);
            const cGap = wsYearly.getCell(`A${currentExcelRow}`);
            cGap.value = '';
            cGap.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFE2F7' } };
            cGap.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'thin', color: { argb: 'FF000000' } }
            };

            currentExcelRow++;
          });

          continue; // done with summary_oee_yearly tab!
        }

        // ─── STANDARD PRODUCTION WORKSHEET ───
        const ws = workbook.addWorksheet(tabInfo.name, {
          views: [{ showGridLines: true, state: 'frozen', ySplit: 5, xSplit: 4 }],
          properties: { tabColor: { argb: tabInfo.colorArgb } }
        });

        // Set column widths
        ws.columns = EXCEL_COLUMNS.map(c => ({
          key: c.col,
          width: c.col === 'AM' ? 28 : (c.col === 'D' ? 22 : (c.col === 'A' ? 12 : (c.col === 'B' || c.col === 'C' ? 8 : 6.5)))
        }));

        // Row 1: MEP FAN LTD. (Banner)
        ws.mergeCells('A1:AM1');
        const r1 = ws.getCell('A1');
        r1.value = 'MEP FAN LTD.';
        r1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
        r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0054A6' } };
        r1.alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getRow(1).height = 32;

        // Row 2: Production Performance Analysis Report
        ws.mergeCells('A2:AM2');
        const r2 = ws.getCell('A2');
        r2.value = 'Production Performance Analysis Report';
        r2.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003366' } };
        r2.alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getRow(2).height = 20;

        // Row 3: Section / Department Name (Fan Lathe / Fan Auto Powder Coating)
        ws.mergeCells('A3:AM3');
        const r3 = ws.getCell('A3');
        r3.value = tabInfo.name;
        r3.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        r3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00334E' } };
        r3.alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getRow(3).height = 20;

        // Row 4: 39 Column Headers
        ws.getRow(4).height = 105;
        const borderThin = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        EXCEL_COLUMNS.forEach((c, idx) => {
          const cell = ws.getRow(4).getCell(idx + 1);
          cell.value = c.label;
          cell.font = { name: 'Times New Roman', size: 9, bold: true, color: { argb: 'FF000000' } };
          if (idx < 4) {
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          } else {
            cell.alignment = { textRotation: 90, vertical: 'bottom', horizontal: 'center', wrapText: true };
          }
          cell.border = borderThin;

          let bgArgb = 'FF7EC8E3';
          if (c.zone === 'purple') bgArgb = 'FF9F72DE';
          if (c.zone === 'cyan') bgArgb = 'FF56C5D0';

          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
        });

        // Fetch Tab Data
        let rowsData = [];
        if (tabId === ACTIVE_TAB) {
          rowsData = SheetState.rows;
        } else {
          const localData = getStoredLocalData(tabId, year, MonthYearState.monthIndex);
          if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
            rowsData = localData.rows;
          } else if (tabId === 'fan_lathe' && year === 2026 && MonthYearState.monthIndex === 7 && typeof INITIAL_EXCEL_ROWS !== 'undefined') {
            rowsData = INITIAL_EXCEL_ROWS;
          } else {
            rowsData = generateBlankMonthRows(tabId, year, MonthYearState.monthIndex);
          }
        }

        // Compute Tab Totals
        const tabTotals = { E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, AH: 0 };
        EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => tabTotals[c.col] = 0);
        const rowsPerDay = getRowsPerDay(tabId);
        rowsData.forEach(r => {
          tabTotals.E += Number(r.E?.val) || 0;
          tabTotals.F += Number(r.F?.val) || 0;
          tabTotals.G += Number(r.G?.val) || 0;

          const mIdx = (r.row - 6) % rowsPerDay;
          const groupInfo = getTimeGroupInfo(tabId, mIdx);

          if (groupInfo.isMaster) {
            tabTotals.H += Number(r.H?.val) || 0;
            tabTotals.I += Number(r.I?.val) || 0;
            tabTotals.J += (typeof r.J?.val === 'number' ? r.J.val : 0);
            tabTotals.AH += Number(r.AH?.val) || 0;
            EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => tabTotals[c.col] += Number(r[c.col]?.val) || 0);
          }
        });
        tabTotals.AI = tabTotals.H > 0 ? (tabTotals.J / tabTotals.H) : 0;
        tabTotals.AJ = tabTotals.E > 0 ? (tabTotals.F / tabTotals.E) : 0;
        tabTotals.AK = (tabTotals.F + tabTotals.G) > 0 ? (tabTotals.F / (tabTotals.F + tabTotals.G)) : 0;
        tabTotals.AL = tabTotals.AI * tabTotals.AJ * tabTotals.AK;

        // Row 5: Total Summary Row
        ws.mergeCells('A5:D5');
        const tTitle = ws.getCell('A5');
        tTitle.value = 'Total:';
        ws.getRow(5).height = 22;

        for (let colIdx = 1; colIdx <= EXCEL_COLUMNS.length; colIdx++) {
          const cell = ws.getRow(5).getCell(colIdx);
          cell.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };

          if (colIdx >= 5) {
            const colDef = EXCEL_COLUMNS[colIdx - 1];
            const val = tabTotals[colDef.col];
            if (colDef.isPercent) {
              cell.value = val || 0;
              cell.numFmt = '0%';
            } else if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
              cell.value = val || 0;
              cell.numFmt = '#,##0';
            }
          }
        }

        // Body Rows (Row 6 onwards)
        const daysInMonth = getDaysInSelectedMonth();
        rowsData.forEach((rowObj, rIdx) => {
          const excelRowIdx = 6 + rIdx;
          const excelRow = ws.getRow(excelRowIdx);
          excelRow.height = 20;

          const isFriday = (rowObj.B?.val === 'Fri');
          const isDayEnd = (rIdx % rowsPerDay === (rowsPerDay - 1));
          const mIdx = rIdx % rowsPerDay;
          const groupInfo = getTimeGroupInfo(tabId, mIdx);

          EXCEL_COLUMNS.forEach((colDef, cIdx) => {
            const colIdx = cIdx + 1;
            const colLetter = colDef.col;
            const isTimeCol = TIME_COLUMNS.includes(colLetter);

            const cell = excelRow.getCell(colIdx);
            const val = rowObj[colDef.col]?.val;

            if (isTimeCol && groupInfo.isSlave) {
              // Slave cell in time group will be spanned by merged master cell
            } else {
              if (colDef.isPercent) {
                cell.value = (val !== null && val !== undefined && val !== '') ? Number(val) : 0;
                cell.numFmt = '0%';
              } else if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
                if (val !== null && val !== undefined && val !== '' && !isNaN(val)) {
                  cell.value = Number(val);
                  cell.numFmt = '#,##0';
                } else {
                  cell.value = val ?? '';
                }
              } else {
                cell.value = val ?? '';
              }
            }

            // Cell Font: Times New Roman Bold
            cell.font = { name: 'Times New Roman', size: 9.5, bold: true, color: { argb: 'FF000000' } };

            // Alignment
            cell.alignment = {
              vertical: 'middle',
              horizontal: colDef.align || (colDef.isNumeric || colDef.isPercent ? 'right' : 'center')
            };

            // Background Fill
            let fillArgb = 'FFFFFFFF';
            if (isFriday) {
              fillArgb = 'FFE3F8E9'; // Light Green
            } else if (colDef.isDt) {
              fillArgb = 'FFE8DDFE'; // Soft Lilac Downtime
            } else if (colDef.col === 'AH') {
              fillArgb = 'FFC8F0FA'; // Soft Cyan Total DT
            }

            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };

            // Border (Solid Black Thin Grid Lines)
            cell.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: isDayEnd ? { style: 'medium', color: { argb: 'FF000000' } } : { style: 'thin', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } }
            };
          });
        });

        // Merging for each day: Date (A), Day (B), Shift (C) + Time Columns for time groups
        for (let d = 0; d < daysInMonth; d++) {
          const dayStart = 6 + (d * rowsPerDay);
          const dayEnd = dayStart + rowsPerDay - 1;

          if (rowsPerDay > 1) {
            // Merge Date, Day, Shift across all machines of the day
            try {
              ws.mergeCells(dayStart, 1, dayEnd, 1);
              ws.getCell(dayStart, 1).alignment = { vertical: 'middle', horizontal: 'center' };
              ws.getCell(dayStart, 1).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'medium', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
              };

              ws.mergeCells(dayStart, 2, dayEnd, 2);
              ws.getCell(dayStart, 2).alignment = { vertical: 'middle', horizontal: 'center' };
              ws.getCell(dayStart, 2).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'medium', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
              };

              ws.mergeCells(dayStart, 3, dayEnd, 3);
              ws.getCell(dayStart, 3).alignment = { vertical: 'middle', horizontal: 'center' };
              ws.getCell(dayStart, 3).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'medium', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
              };
            } catch (e) {}
          }

          // If tab has timeGroups, merge time columns (H to AI) across group rows
          if (tabInfo.timeGroups) {
            for (const group of tabInfo.timeGroups) {
              if (group.count > 1) {
                const rStart = dayStart + group.startIdx;
                const rEnd = rStart + group.count - 1;
                TIME_COLUMNS.forEach(colLetter => {
                  const colIdx = EXCEL_COLUMNS.findIndex(c => c.col === colLetter) + 1;
                  try {
                    ws.mergeCells(rStart, colIdx, rEnd, colIdx);
                    ws.getCell(rStart, colIdx).alignment = { vertical: 'middle', horizontal: 'right' };
                    if (rEnd === dayEnd) {
                      ws.getCell(rStart, colIdx).border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'medium', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                      };
                    }
                  } catch (e) {}
                });
              }
            }
          }
        }
      }

      // Write styled buffer & download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (typeof saveAs !== 'undefined') {
        saveAs(blob, `MEP_FAN_Report_${monthName}_${year}.xlsx`);
      } else {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `MEP_FAN_Report_${monthName}_${year}.xlsx`;
        a.click();
        window.URL.revokeObjectURL(url);
      }

      showToast(`✨ Premium styled Excel workbook exported (${Object.keys(SHEET_TABS).length} Tabs)`, 'success');
      return;
    }

    // 2. SheetJS Fallback
    if (typeof XLSX !== 'undefined') {
      const wb = XLSX.utils.book_new();
      Object.keys(SHEET_TABS).forEach(tabId => {
        if (!isTabAllowedForUser(tabId)) return;
        const tabInfo = SHEET_TABS[tabId];
        if (tabInfo.isSummary) return; // Summary is rendered in ExcelJS
        let rowsData = (tabId === ACTIVE_TAB) ? SheetState.rows : generateBlankMonthRows(tabId, year, MonthYearState.monthIndex);
        const dataMatrix = [
          ['MEP FAN LTD.'],
          ['Production Performance Analysis Report'],
          [tabInfo.name, '', '', ''],
          EXCEL_COLUMNS.map(c => c.label)
        ];
        const ws = XLSX.utils.aoa_to_sheet(dataMatrix);
        XLSX.utils.book_append_sheet(wb, ws, tabInfo.name);
      });
      XLSX.writeFile(wb, `MEP_FAN_Report_${monthName}_${year}.xlsx`);
      showToast('📊 Workbook exported successfully', 'success');
    }
  } catch (err) {
    console.error('Export error:', err);
    showToast('⚠️ Excel export failed', 'error');
  }
}

// ─── PDF EXPORT ───────────────────────────────────────────────────────────────
function exportPDFReport() {
  try {
    const { jsPDF } = window.jspdf;
    if (!jsPDF) {
      window.print();
      return;
    }

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
    const year = MonthYearState.year;
    const tabInfo = SHEET_TABS[ACTIVE_TAB] || { name: 'Report' };

    doc.setFillColor(0, 51, 102);
    doc.rect(0, 0, 297, 18, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(13);
    doc.setFont('times', 'bold');
    doc.text('MEP FAN LTD.', 14, 8);

    doc.setFontSize(9);
    doc.setFont('times', 'normal');
    doc.text(`${tabInfo.name} — (${monthName} ${year})`, 14, 14);

    if (tabInfo.isSummary) {
      const table = document.getElementById('excelMainTable');
      if (table) {
        doc.autoTable({
          html: table,
          startY: 22,
          theme: 'grid',
          headStyles: { fillColor: [0, 51, 102], textColor: 255, fontStyle: 'bold', fontSize: 7.5 },
          styles: { fontSize: 7, cellPadding: 1.5, font: 'times' },
          alternateRowStyles: { fillColor: [248, 250, 252] }
        });
      }
    } else {
      const headers = [['Date', 'Day', 'Shift', 'Machine', 'Cap', 'Actual', 'Rej', 'Plan', 'Run', 'Total DT', 'Avail %', 'Perf %', 'Qual %', 'OEE %', 'Remarks']];

      const t = SheetState.totals;
      const bodyData = [
        [
          'Total:', '', '', '',
          t.E || 0, t.F || 0, t.G || 0, t.H || 0, t.J || 0, t.AH || 0,
          `${((t.AI || 0) * 100).toFixed(0)}%`,
          `${((t.AJ || 0) * 100).toFixed(0)}%`,
          `${((t.AK || 0) * 100).toFixed(0)}%`,
          `${((t.AL || 0) * 100).toFixed(0)}%`,
          ''
        ]
      ];

      SheetState.rows.forEach(r => {
        bodyData.push([
          r.A?.val || '',
          r.B?.val || '',
          r.C?.val || 'Morning',
          r.D?.val || '',
          r.E?.val || '',
          r.F?.val || '',
          r.G?.val || '',
          r.H?.val || '',
          r.J?.val || '-',
          r.AH?.val || 0,
          `${((r.AI?.val || 0) * 100).toFixed(0)}%`,
          `${((r.AJ?.val || 0) * 100).toFixed(0)}%`,
          `${((r.AK?.val || 0) * 100).toFixed(0)}%`,
          `${((r.AL?.val || 0) * 100).toFixed(0)}%`,
          r.AM?.val || ''
        ]);
      });

      doc.autoTable({
        head: headers,
        body: bodyData,
        startY: 22,
        theme: 'grid',
        headStyles: { fillColor: [0, 51, 102], textColor: 255, fontStyle: 'bold', fontSize: 7 },
        styles: { fontSize: 6.5, cellPadding: 1.5, font: 'times' },
        alternateRowStyles: { fillColor: [248, 250, 252] }
      });
    }

    doc.save(`MEP_FAN_${ACTIVE_TAB}_${monthName}_${year}_Report.pdf`);
    showToast('📄 PDF Report downloaded', 'success');
  } catch (e) {
    console.error('PDF export failed:', e);
    window.print();
  }
}

// ─── CONTEXT MENU & SHORTCUTS ─────────────────────────────────────────────────
function initContextMenu() {
  const menu = document.getElementById('gridContextMenu');
  const table = document.getElementById('excelMainTable');

  if (!menu || !table) return;

  table.addEventListener('contextmenu', (e) => {
    const td = e.target.closest('td, th');
    if (!td || !td.dataset.col) return;
    const row = Number(td.dataset.row);
    if (isRowLocked(row)) return;

    e.preventDefault();
    selectCell(td.dataset.col, row);

    menu.style.left = `${e.clientX}px`;
    menu.style.top = `${e.clientY}px`;
    menu.classList.remove('hidden');
  });

  document.addEventListener('click', (e) => {
    if (!menu.contains(e.target)) {
      menu.classList.add('hidden');
    }
  });

  menu.querySelectorAll('.context-menu-row').forEach(item => {
    item.addEventListener('click', () => {
      const action = item.getAttribute('data-action');
      const cur = SheetState.selected;
      if (!cur || isRowLocked(cur.row)) return;

      const colDef = EXCEL_COLUMNS.find(c => c.col === cur.colLetter);

      if (CurrentUser && CurrentUser.isReadOnly && ['quick-form', 'paste', 'clear', 'set-zero'].includes(action)) {
        showToast('🔒 ভিউ মোড: এডিটিং বন্ধ রয়েছে।', 'warning');
        menu.classList.add('hidden');
        return;
      }

      if (action === 'quick-form') {
        openQuickEntryModal(cur.row);
      } else if (action === 'copy') {
        handleClipboardCopy();
      } else if (action === 'paste') {
        navigator.clipboard?.readText().then(handleClipboardPaste);
      } else if (action === 'clear') {
        deleteSelectedRange(false);
      } else if (action === 'set-zero') {
        deleteSelectedRange(true);
      }
      menu.classList.add('hidden');
    });
  });
}

function bindExcelEvents() {
  document.getElementById('btnCloudSync')?.addEventListener('click', () => handleSyncOrPublish());
  document.getElementById('btnSaveGrid')?.addEventListener('click', () => saveSheetData(true));
  document.getElementById('btnExportExcel')?.addEventListener('click', exportExcelFile);
  document.getElementById('btnExportPDF')?.addEventListener('click', exportPDFReport);
  document.getElementById('btnPrint')?.addEventListener('click', () => window.print());
  document.getElementById('btnResetOriginal')?.addEventListener('click', resetToOriginalData);
  document.getElementById('btnUndo')?.addEventListener('click', undoAction);
  document.getElementById('btnRedo')?.addEventListener('click', redoAction);
  document.getElementById('btnQuickEntryModal')?.addEventListener('click', () => {
    const cur = SheetState.selected;
    openQuickEntryModal(cur ? cur.row : 6);
  });

  // Modal events
  document.getElementById('btnModalClose')?.addEventListener('click', closeQuickEntryModal);
  document.getElementById('btnModalCancel')?.addEventListener('click', closeQuickEntryModal);
  document.getElementById('btnModalSave')?.addEventListener('click', () => saveModalEntry(false));
  document.getElementById('btnModalSaveNext')?.addEventListener('click', () => saveModalEntry(true));
  ['mCap', 'mAct', 'mRej', 'mPlan'].forEach(id => {
    const el = document.getElementById(id);
    el?.addEventListener('input', (e) => {
      e.target.value = e.target.value.replace(/[^0-9.]/g, '');
      updateModalLiveKpi();
    });
  });

  // Search events
  document.getElementById('searchInput')?.addEventListener('input', performSearch);
  document.getElementById('btnSearchNext')?.addEventListener('click', () => cycleSearch(true));
  document.getElementById('btnSearchPrev')?.addEventListener('click', () => cycleSearch(false));
  document.getElementById('btnSearchClose')?.addEventListener('click', closeSearchBar);
  document.getElementById('searchInput')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      cycleSearch(!e.shiftKey);
    } else if (e.key === 'Escape') {
      closeSearchBar();
    }
  });

  // Formula bar input & sanitize
  const formulaInput = document.getElementById('formulaBarInput');
  formulaInput?.addEventListener('input', () => {
    const cur = SheetState.selected;
    if (cur && cur.colLetter !== 'AM' && !formulaInput.readOnly) {
      if (!formulaInput.value.trim().startsWith('=')) {
        formulaInput.value = formulaInput.value.replace(/[^0-9.]/g, '');
      }
    }
  });

  formulaInput?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const cur = SheetState.selected;
      if (cur && cur.row !== 5 && !isRowLocked(cur.row)) {
        const colDef = EXCEL_COLUMNS.find(c => c.col === cur.colLetter);
        if (colDef && !colDef.isFormula && !colDef.isReadOnly) {
          saveCellUpdate(cur.colLetter, cur.row, formulaInput.value.trim(), colDef);
        }
      }
      formulaInput.blur();
    }
  });

  // Fast and smooth mouse drag range selection across cells
  const mainTable = document.getElementById('excelMainTable');
  mainTable?.addEventListener('mouseover', (e) => {
    if (!SheetState.isSelecting) return;
    const td = e.target.closest('td');
    if (!td || !td.dataset.col || !td.dataset.row) return;
    const r = Number(td.dataset.row);
    if (r === 5 || isRowLocked(r)) return;
    SheetState.rangeSelection.end = { col: td.dataset.col, row: r };
    highlightSelectedRange();
  });

  // Native Clipboard copy event
  window.addEventListener('copy', (e) => {
    if (SheetState.isEditing) return;
    if (document.activeElement && ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
    const cur = SheetState.selected;
    if (!cur) return;

    e.preventDefault();
    handleClipboardCopy();
  });

  // Global Keyboard Navigation
  window.addEventListener('keydown', (e) => {
    if (SheetState.isEditing) return;
    if (document.activeElement === formulaInput) return;
    if (document.activeElement === document.getElementById('searchInput')) return;
    if (!document.getElementById('quickEntryModal')?.classList.contains('hidden')) {
      if (e.key === 'Escape') closeQuickEntryModal();
      return;
    }

    let cur = SheetState.selected;
    if (!cur || !cur.element || !document.body.contains(cur.element)) {
      if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter'].includes(e.key)) {
        e.preventDefault();
        selectCell('E', 6);
        return;
      }
      return;
    }

    const colDef = EXCEL_COLUMNS.find(c => c.col === cur.colLetter);

    // Shortcuts
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      const maxActive = getMaxActiveRow();
      SheetState.rangeSelection.start = { col: 'E', row: 6 };
      SheetState.rangeSelection.end = { col: 'AM', row: maxActive };
      selectCell('E', 6, false);
      highlightSelectedRange();
      showToast('🔲 All editable cells selected', 'info');
      return;
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: এডিটিং বন্ধ রয়েছে।', 'warning');
        return;
      }
      saveSheetData(true);
      showToast('💾 Saved!', 'success');
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) return;
      undoAction();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) return;
      redoAction();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      openSearchBar();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      handleClipboardCopy();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: পেস্ট বা পরিবর্তন করা যাবে না।', 'warning');
        return;
      }
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText().then(text => {
          if (text) handleClipboardPaste(text);
        }).catch(() => {
          showToast('⚠️ Please use Right Click -> Paste or browser paste', 'warning');
        });
      }
    } else if (e.key === ' ' || (e.altKey && e.key.toLowerCase() === 'e')) {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: ডাটা এন্ট্রি বন্ধ রয়েছে।', 'warning');
        return;
      }
      if (!isRowLocked(cur.row)) openQuickEntryModal(cur.row);
    } else if (e.shiftKey && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      // Excel-like Shift + Arrow Multi-cell Range Selection
      e.preventDefault();
      if (!SheetState.rangeSelection.start) {
        SheetState.rangeSelection.start = { col: cur.colLetter, row: cur.row };
      }
      if (!SheetState.rangeSelection.end) {
        SheetState.rangeSelection.end = { col: cur.colLetter, row: cur.row };
      }

      let endR = SheetState.rangeSelection.end.row;
      let endCIdx = EXCEL_COLUMNS.findIndex(c => c.col === SheetState.rangeSelection.end.col);

      if (e.key === 'ArrowUp') endR = Math.max(6, endR - 1);
      if (e.key === 'ArrowDown') endR = Math.min(getMaxActiveRow(), endR + 1);
      if (e.key === 'ArrowLeft') endCIdx = Math.max(0, endCIdx - 1);
      if (e.key === 'ArrowRight') endCIdx = Math.min(EXCEL_COLUMNS.length - 1, endCIdx + 1);

      SheetState.rangeSelection.end = { col: EXCEL_COLUMNS[endCIdx].col, row: endR };
      highlightSelectedRange();
      return;
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      navigateSelection(0, e.ctrlKey ? -100 : -1);
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      navigateSelection(0, e.ctrlKey ? 100 : 1);
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      navigateSelection(e.ctrlKey ? -100 : -1, 0);
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      navigateSelection(e.ctrlKey ? 100 : 1, 0);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      navigateSelection(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: কোনো ডাটা এডিট করা যাবে না।', 'warning');
        return;
      }
      if (colDef && !colDef.isFormula && !colDef.isReadOnly && cur.row !== 5 && !isRowLocked(cur.row)) {
        startCellEdit(cur.element, cur.colLetter, cur.row, colDef);
      }
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: কোনো ডাটা ডিলিট বা এডিট করা যাবে না।', 'warning');
        return;
      }
      e.preventDefault();
      deleteSelectedRange(false);
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('🔒 ভিউ মোড: কোনো ডাটা টাইপ বা এডিট করা যাবে না।', 'warning');
        return;
      }
      if (cur.row !== 5 && colDef && !colDef.isFormula && !colDef.isReadOnly && !isRowLocked(cur.row)) {
        if (cur.colLetter !== 'AM' && !/^[0-9.=\-+/*()]$/.test(e.key)) {
          return;
        }
        e.preventDefault();
        startCellEdit(cur.element, cur.colLetter, cur.row, colDef, e.key);
      }
    }
  });

  // Automatically flush and persist any active cell edit on browser refresh or navigation
  const flushAndSaveOnUnload = () => {
    if (SheetState.isEditing && SheetState.activeInput) {
      const cur = SheetState.selected;
      if (cur && cur.row !== 5 && !isRowLocked(cur.row)) {
        const colDef = EXCEL_COLUMNS.find(c => c.col === cur.colLetter);
        if (colDef && !colDef.isReadOnly && !colDef.isFormula) {
          const rowObj = SheetState.rows.find(r => r.row === cur.row);
          if (rowObj) {
            const rawVal = SheetState.activeInput.value.trim();
            const cleanVal = (cur.colLetter === 'AM') ? (rawVal === '' ? null : rawVal) : sanitizeNumericValue(rawVal);
            if (!rowObj[cur.colLetter]) rowObj[cur.colLetter] = {};
            rowObj[cur.colLetter].val = cleanVal;
            recalculateRow(rowObj);
            recalculateTotalRow();
          }
        }
      }
    }
    saveSheetData(false);
  };

  window.addEventListener('beforeunload', flushAndSaveOnUnload);
  window.addEventListener('pagehide', flushAndSaveOnUnload);

  initContextMenu();
}

function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = 'toast-msg';

  const iconSpan = document.createElement('span');
  iconSpan.textContent = type === 'success' ? '✅' : (type === 'error' ? '❌' : (type === 'warning' ? '⚠️' : 'ℹ️'));

  const textSpan = document.createElement('span');
  textSpan.textContent = message;

  toast.appendChild(iconSpan);
  toast.appendChild(textSpan);
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}
