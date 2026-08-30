let lastBroadcastTimestamp = 0;
/**
 * 
 * MEP FAN LTD. - Multi-Tab Excel Web Spreadsheet Engine
 * Tabs:
 *   - "Fan Lathe" (3 Machines: Rotor, Bottom Cover, Top Cover)
 *   - "Fan Auto Powder Coating" (4 Machines: APC (Blade), APC (Downpipe), APC (Body, Cover), Blade Rivet)
/**
 * 
 * MEP FAN LTD. - Multi-Tab Excel Web Spreadsheet Engine
 * Tabs:
 *   - "Fan Lathe" (3 Machines: Rotor, Bottom Cover, Top Cover)
 *   - "Fan Auto Powder Coating" (4 Machines: APC (Blade), APC (Downpipe), APC (Body, Cover), Blade Rivet)
 * Dynamic Month/Year Selection + Strict Numeric Validation + Real-Time OEE
 * 
 */

'use strict';

// ──────────────────────────────────────────────────────────────────────────
const EXCEL_COLUMNS = [
  { col: 'A', label: 'Date', width: 76, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'B', label: 'Day', width: 42, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'C', label: 'Shift', width: 54, isReadOnly: true, zone: 'sky', align: 'center' },
  { col: 'D', label: 'Machine Name', width: 160, isReadOnly: true, zone: 'sky', align: 'left' },
  { col: 'E', label: 'Machine Capacity (Pcs)', width: 68, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'F', label: 'Actual Prd. (Pcs)', width: 68, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'G', label: 'Rejection (Pcs)', width: 50, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'H', label: 'Planned Prd. Time (Min)', width: 64, isNumeric: true, zone: 'sky', align: 'right' },
  { col: 'I', label: 'Expected DownTime (Min)', width: 50, isFormula: true, formula: '=IF(E{r}<>"",30,0)', zone: 'sky', align: 'right' },
  { col: 'J', label: 'Total Prd. Run Time (Min)', width: 64, isFormula: true, formula: '=H{r}-AH{r}', zone: 'sky', align: 'right' },
  
  // 26 Downtime columns (Purple Zone)
  { col: 'K', label: 'Heater/Coil Problem', code: 10, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'L', label: 'Power Shutdown', code: 11, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'M', label: 'Machine Breakdown', code: 12, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'N', label: 'Die/ Mold Problem', code: 13, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'O', label: 'Model/ Die Change', code: 14, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'P', label: 'Air Presser Problem', code: 15, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Q', label: 'Water line Problem', code: 16, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'R', label: 'QC Test', code: 17, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'S', label: 'D Coil Insert', code: 18, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'T', label: 'RM Shortage', code: 19, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'U', label: 'Crean Problem', code: 20, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'V', label: 'Worker Absent', code: 21, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'W', label: 'Printer M/C Problem', code: 22, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'X', label: 'UPS shutdown', code: 23, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'DT_BOX_MOVING', label: 'Box Moving', code: 24, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'DT_UP_CASTING', label: 'Up Casting Meeting', code: 25, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'DT_SHEET_CUTTING', label: 'Sheet Cutting & Pipe Carry', code: 26, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Y', label: 'Load Problem', code: 27, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'Z', label: 'Namaz', code: 28, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AA', label: 'Conveyor Belt Problem', code: 29, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AB', label: 'Fitting Problem', code: 30, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AC', label: 'Gas Presser Problem', code: 31, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AD', label: 'Mold polish & Clean', code: 32, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AE', label: 'Alu. Ash Extraction', code: 33, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AF', label: 'Robot Problems', code: 34, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AG', label: 'Alu. Recipe Problem', code: 35, width: 52, isNumeric: true, isDt: true, zone: 'purple', align: 'right' },
  { col: 'AG_LOSS', label: 'Speed/Cap. Loss (Min)', width: 64, isFormula: true, formula: '=IF(AND(E{r}>0,H{r}>0),ROUND(MAX(0,H{r}*(1-F{r}/E{r})),0),0)', zone: 'purple-loss', isLossDt: true, webOnly: true, align: 'right' },
  
  // KPI Columns AH to AL (Preserved Original Keys!)
  { col: 'AH', label: 'Total Down Time (Mins)', width: 68, isFormula: true, formula: '=SUM(K{r}:AG{r})', zone: 'total-dt', isTotalDt: true, align: 'right' },
  { col: 'AI', label: 'Availability (%)', width: 56, isFormula: true, formula: '=IFERROR(J{r}/H{r},"0")', isPercent: true, zone: 'cyan', isAvail: true, align: 'right' },
  { col: 'AJ', label: 'Performance (%)', width: 56, isFormula: true, formula: '=IFERROR(F{r}/E{r},"0")', isPercent: true, zone: 'cyan', isPerf: true, align: 'right' },
  { col: 'AK', label: 'Quality (%)', width: 56, isFormula: true, formula: '=IFERROR(F{r}/(F{r}+G{r}),"0")', isPercent: true, zone: 'cyan', isQual: true, align: 'right' },
  { col: 'AL', label: 'OEE (%)', width: 56, isFormula: true, formula: '=IFERROR(AK{r}*AJ{r}*AI{r},"0")', isPercent: true, zone: 'cyan', isOee: true, align: 'right' },
  
  // Column AM (Remarks) (Sky Blue Zone) - Only Column that allows text!
  { col: 'AM', label: 'Remarks', width: 140, zone: 'sky', isRemarks: true, align: 'left' }
];

// ──────────────────────────────────────────────────────────────────────────
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

const USERS_CONFIG = {
  admin: {
    id: 'admin',
    name: 'Admin',
    role: 'System Administrator',
    pin: '8250',
    icon: '🛡️',
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
    pin: '2645',
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

const TIME_COLUMNS = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'DT_BOX_MOVING', 'DT_UP_CASTING', 'DT_SHEET_CUTTING', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AG_LOSS', 'AH', 'AI'];

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

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
// App startup initialized at bottom of script

// ──────────────────────────────────────────────────────────────────────────
function bindExcelEvents() {
  // Save Grid Button
  document.getElementById('btnSaveGrid')?.addEventListener('click', () => {
    if (CurrentUser && CurrentUser.isReadOnly) {
      showToast('Permission Denied: Read-only access.', 'warning');
      return;
    }
    saveSheetData(true);
    showToast('Changes saved to cloud successfully!', 'success');
  });

  // Export Excel Button
  document.getElementById('btnExportExcel')?.addEventListener('click', () => {
    exportExcelFile();
  });

  // Export PDF Button
  document.getElementById('btnExportPdf')?.addEventListener('click', () => {
    exportPDFReport();
  });

  // Reset Month Button
  document.getElementById('btnResetMonth')?.addEventListener('click', () => {
    if (CurrentUser && CurrentUser.isReadOnly) {
      showToast('Permission Denied: Read-only access.', 'warning');
      return;
    }
    if (confirm('Are you sure you want to reset this month\'s data?')) {
      resetCurrentMonthData();
    }
  });

  // Quick Entry Modal Close & Save Buttons
  document.getElementById('btnModalClose')?.addEventListener('click', closeQuickEntryModal);
  document.getElementById('btnModalSave')?.addEventListener('click', saveQuickEntryModal);

  // Search Bar Buttons
  document.getElementById('btnOpenSearch')?.addEventListener('click', openSearchBar);
  document.getElementById('btnCloseSearch')?.addEventListener('click', closeSearchBar);

  // Keyboard Shortcuts & Context Menu
  initKeyboardShortcuts();
  initContextMenu();
}

function initAuthSystem() {
  updatePortalClock();
  setInterval(updatePortalClock, 1000);

  // Bind mouse move for dynamic radial spotlight lighting on real glass cards
  document.querySelectorAll('.user-portal-card').forEach(card => {
    card.addEventListener('mousemove', (e) => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      card.style.setProperty('--mouse-x', `${x}px`);
      card.style.setProperty('--mouse-y', `${y}px`);
    });
  });

  // Bind profile cards click
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

  // Modal keyboard keydown handler (Single source of truth to prevent double-typing)
  window.addEventListener('keydown', (e) => {
    const modal = document.getElementById('pinModal');
    if (!modal || modal.classList.contains('hidden')) return;

    if (e.key === 'Escape') {
      e.preventDefault();
      closePinModal();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      verifyAndLogin();
    } else if (e.key === 'Backspace') {
      e.preventDefault();
      handlePinKey('backspace');
    } else if (/^[0-9]$/.test(e.key)) {
      e.preventDefault();
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
    legacyClock.textContent = `${dateStr} • ${timeStr} • MEP FAN LTD. OEE`;
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
  const len = EnteredPin.length;
  for (let i = 0; i < 4; i++) {
    const dot = document.getElementById('dot' + i);
    if (dot) {
      if (i < len) {
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

  if (EnteredPin.length === 4) {
    // Immediate zero-delay verification
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
      errEl.textContent = 'Incorrect PIN! Please try again.';
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
    syncBtn.style.display = 'none';
  }

  if (CurrentUser.isReadOnly) {
    if (iconEl) iconEl.innerHTML = '<svg class="w-4 h-4 text-indigo-300 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>';
    if (nameEl) nameEl.innerHTML = `<span class="text-amber-300 font-bold text-xs block tracking-wide">VIEW ONLY</span><span class="text-slate-300 text-[10px] font-medium block mt-0.5">Read-Only Mode</span>`;
    if (saveBtn) saveBtn.style.display = 'none';
    if (excelBtn) excelBtn.style.display = 'none';
    if (formulaInput) formulaInput.readOnly = true;
  } else {
    if (iconEl) {
      if (CurrentUser.id === 'admin') {
        iconEl.innerHTML = '<svg class="w-4 h-4 text-amber-300 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>';
      } else {
        iconEl.innerHTML = CurrentUser.svgIcon ? CurrentUser.svgIcon.replace('w-7 h-7', 'w-4 h-4 text-white') : '<svg class="w-4 h-4 text-slate-200" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>';
      }
    }
    if (nameEl) {
      nameEl.innerHTML = `<span class="text-white font-bold text-xs block tracking-wide">${CurrentUser.name.toUpperCase()}</span><span class="text-slate-300 text-[10px] font-medium block mt-0.5">${CurrentUser.role}</span>`;
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
  GoogleSheetsRealtimeEngine.listenToActiveSheet(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);

  // 5. Seamlessly reveal the main container with zero blink/flicker
  document.documentElement.classList.remove('user-unauthenticated');
  document.documentElement.classList.add('user-authenticated');
  hideLoginPortal();

  if (showWelcome) {
    if (CurrentUser.isReadOnly) {
    showToast('Information updated.', 'info');
    } else {
    showToast('Operation completed successfully.', 'success');
    }
  }
}

function logoutUser() {
  CurrentUser = null;
  sessionStorage.removeItem('mep_auth_user');
  document.documentElement.classList.remove('user-authenticated');
  document.documentElement.classList.add('user-unauthenticated');
  showLoginPortal();
    showToast('Information updated.', 'info');
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

  // Enable Smooth Horizontal Mouse Wheel Scrolling on Tab Bar
  container?.addEventListener('wheel', (e) => {
    if (e.deltaY !== 0) {
      e.preventDefault();
      container.scrollLeft += e.deltaY;
    }
  }, { passive: false });

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
    showToast('Permission Denied: This operation is restricted.', 'warning');
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
  GoogleSheetsRealtimeEngine.listenToActiveSheet(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
}

function onMonthYearChange() {
  loadSheetData();
  recalculateAllFormulas();
  renderExcelTable();
  selectInitialCell();
  GoogleSheetsRealtimeEngine.listenToActiveSheet(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
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

// ──────────────────────────────────────────────────────────────────────────
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
  return LocalStorageEngine.load(tabId, year, monthIndex);
}

function loadSheetData() {
  const expectedRowsCount = getDaysInSelectedMonth() * getRowsPerDay(ACTIVE_TAB);
  const expectedMachines = getMachinesForTab(ACTIVE_TAB);
  const localData = getStoredLocalData(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);

  if (localData && Array.isArray(localData.rows) && localData.rows.length > 0) {
    const parsed = localData.rows;

    // Best case: row count and machine names match perfectly
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

    // Fallback: row count matches but machine names might differ slightly
    // Still use the data - it came from Firebase and has valid entered values
    if (parsed.length === expectedRowsCount) {
      SheetState.rows = parsed;
      return;
    }

    // Last resort: data has different row count but has actual entered data
    // Use it anyway rather than showing blank - data preservation is critical
    let hasAnyData = false;
    for (const r of parsed) {
      if (r.E?.val || r.F?.val || r.G?.val || r.H?.val) {
        hasAnyData = true;
        break;
      }
    }
    if (hasAnyData) {
      SheetState.rows = parsed;
      return;
    }
  }

  // Only generate blank rows if there's truly no saved data at all
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
    showToast('Information updated.', 'info');
    return;
  }
  SheetState.redoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(SheetState.undoStack.pop());
  recalculateAllFormulas();
  renderExcelTable();
  saveSheetData(false);
    showToast('Information updated.', 'info');
}

function redoAction() {
  if (SheetState.redoStack.length === 0) {
    showToast('Information updated.', 'info');
    return;
  }
  SheetState.undoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(SheetState.redoStack.pop());
  recalculateAllFormulas();
  renderExcelTable();
  saveSheetData(false);
    showToast('Information updated.', 'info');
}

// ──────────────────────────────────────────────────────────────────────────
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDummyKeyForRestApiCompatibility",
  authDomain: "whatsapp-c10ef.firebaseapp.com",
  databaseURL: "https://whatsapp-c10ef-default-rtdb.firebaseio.com",
  projectId: "whatsapp-c10ef",
  storageBucket: "whatsapp-c10ef.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef123456"
};

const FIREBASE_RTDB_BASE_URL = "https://whatsapp-c10ef-default-rtdb.firebaseio.com";
let firebaseDb = null;
let isFirebaseInitialized = false;

function cleanDataForFirebase(data) {
  return JSON.parse(JSON.stringify(data, (key, value) => {
    if (value === undefined) return null;
    return value;
  }));
}

// ──────────────────────────────────────────────────────────────────────────
// MODULE 1: LOCAL STORAGE CACHE (0ms Instant Display & Offline Fallback)
// ──────────────────────────────────────────────────────────────────────────
const LocalStorageEngine = {
  getKey(tabId, year, monthIndex) {
    return getStorageKey(tabId, year, monthIndex);
  },

  save(tabId, year, monthIndex, rows, updatedBy = 'User') {
    try {
      const key = this.getKey(tabId, year, monthIndex);
      const payload = {
        rows: cleanDataForFirebase(rows),
        tabId: tabId,
        year: year,
        monthIndex: monthIndex,
        updatedAt: Date.now(),
        updatedBy: CurrentUser ? CurrentUser.name : updatedBy
      };
      localStorage.setItem(key, JSON.stringify(payload));
      triggerSaveIndicator();
      return true;
    } catch (e) {
      console.error('LocalStorageEngine.save error:', e);
      return false;
    }
  },

  load(tabId, year, monthIndex) {
    try {
      const key = this.getKey(tabId, year, monthIndex);
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.rows)) return parsed;
      if (Array.isArray(parsed)) return { rows: parsed, updatedAt: 0 };
      return null;
    } catch (e) {
      console.error('LocalStorageEngine.load error:', e);
      return null;
    }
  },

  clear(tabId, year, monthIndex) {
    try {
      localStorage.removeItem(this.getKey(tabId, year, monthIndex));
      return true;
    } catch (e) {
      return false;
    }
  }
};

function getStoredLocalData(tabId, year, monthIndex) {
  return LocalStorageEngine.load(tabId, year, monthIndex);
}

// ──────────────────────────────────────────────────────────────────────────
// MODULE 2: PURE GOOGLE SHEETS REAL-TIME ENGINE (Cell/Row Granular Cloud Sync)
// ──────────────────────────────────────────────────────────────────────────
let activeSheetRef = null;
let activeSheetListenerTab = null;

const GoogleSheetsRealtimeEngine = {
  // Push a single modified row immediately to Firebase Realtime Database
  async pushRowChange(tabId, year, monthIndex, rowObj) {
    if (!rowObj || typeof rowObj.row !== 'number') return;
    if (CurrentUser && CurrentUser.isReadOnly) return;

    const rIdx = rowObj.row - 6;
    if (rIdx < 0) return;

    const cleanRow = cleanDataForFirebase(rowObj);
    const path = `mep_oee_v2/sheets/${year}/${monthIndex}/${tabId}/rows/${rIdx}`;

    updateCloudStatusUI('syncing', 'Saving...');

    try {
      if (firebaseDb) {
        firebaseDb.ref(path).set(cleanRow);
      }
      // REST API call
      fetch(`${FIREBASE_RTDB_BASE_URL}/${path}.json`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(cleanRow)
      }).then(() => {
        updateCloudStatusUI('synced', 'All changes saved to cloud');
      }).catch(() => {});
    } catch (err) {
      console.error('pushRowChange error:', err);
      updateCloudStatusUI('offline', 'Saved locally (offline)');
    }
  },

  // Push batch of modified rows (e.g. from paste or modal entry)
  async pushRowsBatch(tabId, year, monthIndex, rowsArray) {
    if (!Array.isArray(rowsArray) || rowsArray.length === 0) return;
    if (CurrentUser && CurrentUser.isReadOnly) return;

    updateCloudStatusUI('syncing', 'Saving...');

    const promises = rowsArray.map(rObj => {
      const rIdx = rObj.row - 6;
      if (rIdx < 0) return Promise.resolve();
      const cleanRow = cleanDataForFirebase(rObj);
      const path = `mep_oee_v2/sheets/${year}/${monthIndex}/${tabId}/rows/${rIdx}`;

      if (firebaseDb) {
        firebaseDb.ref(path).set(cleanRow);
      }
      return fetch(`${FIREBASE_RTDB_BASE_URL}/${path}.json`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(cleanRow)
      });
    });

    try {
      await Promise.all(promises);
      updateCloudStatusUI('synced', 'All changes saved to cloud');
    } catch (e) {
      updateCloudStatusUI('synced', 'All changes saved to cloud');
    }
  },

  // Listen to the active sheet in real-time via WebSockets
  listenToActiveSheet(tabId, year, monthIndex) {
    if (activeSheetRef) {
      try { activeSheetRef.off(); } catch (e) {}
      activeSheetRef = null;
    }

    activeSheetListenerTab = `${tabId}_${year}_${monthIndex}`;
    const path = `mep_oee_v2/sheets/${year}/${monthIndex}/${tabId}/rows`;

    if (firebaseDb) {
      activeSheetRef = firebaseDb.ref(path);

      // Listen for individual row updates from other collaborators
      activeSheetRef.on('child_changed', (snap) => {
        const rIdx = Number(snap.key);
        const incomingRow = snap.val();
        if (!incomingRow || isNaN(rIdx)) return;

        // If local user is currently editing THIS exact row, don't interrupt cursor
        if (SheetState.isEditing && SheetState.selected && SheetState.selected.row === (rIdx + 6)) {
          return;
        }

        if (Array.isArray(SheetState.rows) && SheetState.rows[rIdx]) {
          SheetState.rows[rIdx] = incomingRow;
          recalculateRow(SheetState.rows[rIdx]);
          recalculateTotalRow();
          updateSingleRowDisplay(rIdx + 6);
          updateTotalRowDisplay();
          LocalStorageEngine.save(tabId, year, monthIndex, SheetState.rows, 'RealtimeCloud');
        }
      });
    }

    // Fallback polling every 3 seconds
    if (!window._mepGoogleSheetsPoll) {
      window._mepGoogleSheetsPoll = setInterval(async () => {
        if (SheetState.isEditing) return; // Don't poll while user is actively typing
        try {
          const res = await fetch(`${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/sheets/${MonthYearState.year}/${MonthYearState.monthIndex}/${ACTIVE_TAB}/rows.json`);
          if (res.ok) {
            const cloudRows = await res.json();
            if (Array.isArray(cloudRows) && cloudRows.length > 0 && !SheetState.isEditing) {
              let hasChanges = false;
              cloudRows.forEach((r, idx) => {
                if (!r || idx >= SheetState.rows.length) return;
                // Quick compare input fields
                if (JSON.stringify(r.E) !== JSON.stringify(SheetState.rows[idx]?.E) ||
                    JSON.stringify(r.F) !== JSON.stringify(SheetState.rows[idx]?.F) ||
                    JSON.stringify(r.G) !== JSON.stringify(SheetState.rows[idx]?.G) ||
                    JSON.stringify(r.H) !== JSON.stringify(SheetState.rows[idx]?.H) ||
                    JSON.stringify(r.AM) !== JSON.stringify(SheetState.rows[idx]?.AM)) {
                  SheetState.rows[idx] = r;
                  recalculateRow(SheetState.rows[idx]);
                  updateSingleRowDisplay(idx + 6);
                  hasChanges = true;
                }
              });
              if (hasChanges) {
                recalculateTotalRow();
                updateTotalRowDisplay();
                LocalStorageEngine.save(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, SheetState.rows, 'CloudPoll');
              }
            }
          }
        } catch(e) {}
      }, 3000);
    }
  },

  // Fetch full sheet from Firebase
  async fetchSheetFromCloud(tabId, year, monthIndex) {
    try {
      const url = `${FIREBASE_RTDB_BASE_URL}/mep_oee_v2/sheets/${year}/${monthIndex}/${tabId}.json?nocache=${Date.now()}`;
      const res = await fetch(url, { cache: 'no-store' });
      if (res.ok) {
        const data = await res.json();
        if (data && Array.isArray(data.rows) && data.rows.length > 0) {
          return data.rows;
        }
      }
    } catch (e) {
      console.error('fetchSheetFromCloud error:', e);
    }
    return null;
  }
};

// ──────────────────────────────────────────────────────────────────────────
// MODULE 3: INITIALIZATION & USER ACTIONS (Google Sheets Model)
// ──────────────────────────────────────────────────────────────────────────

function initFirebase() {
  try {
    if (typeof firebase !== 'undefined') {
      if (!firebase.apps.length) {
        firebase.initializeApp(FIREBASE_CONFIG);
      }
      firebaseDb = firebase.database();
      isFirebaseInitialized = true;
      console.log(' Google Sheets Realtime Engine Connected: whatsapp-c10ef');
      updateCloudStatusUI('online', 'All changes saved to cloud');

      if (firebase.auth) {
        firebase.auth().signInAnonymously().catch(() => {});
      }

      firebaseDb.ref('.info/connected').on('value', (snap) => {
        if (snap.val() === true) {
          updateCloudStatusUI('online', 'All changes saved to cloud');
        } else {
          updateCloudStatusUI('offline', 'Saved locally (offline)');
        }
      });

      // Start listening to active sheet
      GoogleSheetsRealtimeEngine.listenToActiveSheet(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
    } else {
      updateCloudStatusUI('online', 'All changes saved to cloud (REST)');
    }

    // Initial background sync from cloud
    syncAllCloudData(false);
  } catch (e) {
    console.error('Firebase init error:', e);
    updateCloudStatusUI('online', 'All changes saved to cloud');
  }
}

function updateCloudStatusUI(status, text) {
  const dot = document.getElementById('cloudStatusDot');
  const txt = document.getElementById('cloudStatusText');
  const subTxt = document.getElementById('cloudStatusSubText');
  const autoSave = document.getElementById('autoSaveIndicator');

  if (dot) {
    if (status === 'online' || status === 'synced') {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse shrink-0';
      if (txt) txt.textContent = 'LIVE';
      if (subTxt) subTxt.textContent = text || 'All changes saved to cloud';
      if (autoSave) autoSave.textContent = '● ' + (text || 'All changes saved to cloud');
    } else if (status === 'syncing') {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-amber-400 animate-ping shrink-0';
      if (txt) txt.textContent = 'SAVING';
      if (subTxt) subTxt.textContent = text || 'Saving to cloud...';
      if (autoSave) autoSave.textContent = ' Saving...';
    } else {
      dot.className = 'w-2.5 h-2.5 rounded-full bg-slate-400 shrink-0';
      if (txt) txt.textContent = 'OFFLINE';
      if (subTxt) subTxt.textContent = text || 'Saved locally';
      if (autoSave) autoSave.textContent = '💾 Saved locally';
    }
  }
}

// Background sync on startup / tab change
async function syncAllCloudData(showFeedback = false) {
  try {
    const year = MonthYearState.year;
    const monthIndex = MonthYearState.monthIndex;
    const cloudRows = await GoogleSheetsRealtimeEngine.fetchSheetFromCloud(ACTIVE_TAB, year, monthIndex);

    if (cloudRows && Array.isArray(cloudRows) && cloudRows.length > 0) {
      cloudRows.forEach(r => recalculateRow(r));
      SheetState.rows = cloudRows;
      LocalStorageEngine.save(ACTIVE_TAB, year, monthIndex, cloudRows, 'CloudSync');
      recalculateAllFormulas();
      renderExcelTable();
      updateTotalRowDisplay();
      updateCloudStatusUI('synced', 'All changes saved to cloud');
    showToast('Operation completed successfully.', 'success');
    } else {
      updateCloudStatusUI('synced', 'All changes saved to cloud');
    }
  } catch (err) {
    console.error('syncAllCloudData error:', err);
  }
}

function handleInchargeSave() {
  // Optional manual save trigger -> pushes full active sheet to cloud
  GoogleSheetsRealtimeEngine.pushRowsBatch(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, SheetState.rows);
    showToast('Operation completed successfully.', 'success');
}

function adminBroadcastLiveData() {
  handleInchargeSave();
}

function handleSyncOrPublish() {
  syncAllCloudData(true);
}

function saveSheetData(pushHistory = true) {
  if (pushHistory) pushHistoryState();
  LocalStorageEngine.save(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, SheetState.rows);
}

function triggerSaveIndicator() {
  const el = document.getElementById('autoSaveIndicator');
  if (el) {
    el.textContent = '● Saved to cloud';
    el.style.opacity = '1';
  }
}

function resetToOriginalData() {
  const tabName = SHEET_TABS[ACTIVE_TAB].name;
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  if (confirm(`Reset entire sheet for [${tabName}] (${monthName} ${MonthYearState.year})?`)) {
    LocalStorageEngine.clear(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex);
    loadSheetData();
    recalculateAllFormulas();
    renderExcelTable();
    GoogleSheetsRealtimeEngine.pushRowsBatch(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, SheetState.rows);
    showToast('Information updated.', 'info');
  }
}

// ──────────────────────────────────────────────────────────────────────────
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

  // AG_LOSS: Daily Speed / Capacity Loss Downtime (Min)
  let effectivePlan = plan;
  if (groupInfo.isSlave) {
    const dayStartRow = r - mIdx;
    const masterRow = SheetState.rows.find(row => row.row === (dayStartRow + groupInfo.masterIdx));
    if (masterRow) {
      if (masterRow.H) effectivePlan = Number(masterRow.H.val) || 0;
      // Copy all time columns and downtimes from master row
      TIME_COLUMNS.forEach(cLetter => {
        if (masterRow[cLetter]) {
          if (!rowObj[cLetter]) rowObj[cLetter] = {};
          rowObj[cLetter].val = masterRow[cLetter].val;
          rowObj[cLetter].formula = masterRow[cLetter].formula;
        }
      });
    }
  }

  let lossDt = 0;
  let hasValidLossData = false;

  if (groupInfo.count > 1) {
    const dayStartRow = r - mIdx;
    let groupPerfSum = 0;
    let activeProducts = 0;

    for (let i = 0; i < groupInfo.count; i++) {
      const gRow = SheetState.rows.find(row => row.row === (dayStartRow + groupInfo.masterIdx + i));
      if (gRow) {
        const gCap = Number(gRow.E?.val) || 0;
        const gAct = Number(gRow.F?.val) || 0;
        if (gCap > 0) {
          groupPerfSum += (gAct / gCap);
          activeProducts++;
        }
      }
    }

    if (effectivePlan > 0 && activeProducts > 0) {
      hasValidLossData = true;
      const avgPerf = groupPerfSum / activeProducts;
      if (avgPerf < 1) {
        lossDt = Math.round(effectivePlan * (1 - avgPerf));
      }
    }
  } else {
    if (cap > 0 && plan > 0) {
      hasValidLossData = true;
      const achRate = act / cap;
      if (achRate < 1) {
        lossDt = Math.round(plan * (1 - achRate));
      }
    }
  }

  if (!rowObj.AG_LOSS) rowObj.AG_LOSS = {};
  rowObj.AG_LOSS.val = hasValidLossData ? lossDt : '-';
  rowObj.AG_LOSS.formula = `=IF(AND(E${r}>0,H${r}>0),ROUND(MAX(0,H${r}*(1-F${r}/E${r})),0),0)`;

  // AH: Total Down Time = SUM of all dtCols
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
  rowObj.AH.val = Math.round(sumDt);
  rowObj.AH.formula = `=SUM(K${r}:AG${r})`;

  // J: Total Prd. Run Time = H - AH
  if (!rowObj.J) rowObj.J = {};
  if (effectivePlan > 0 || hasAnyDt) {
    rowObj.J.val = Math.max(0, Math.round(effectivePlan - sumDt));
  } else {
    rowObj.J.val = '-';
  }
  rowObj.J.formula = `=H${r}-AH${r}`;

  const numRunTime = typeof rowObj.J.val === 'number' ? rowObj.J.val : 0;

  // AI: Availability (%) = IFERROR(J/H, "0")
  let avail = effectivePlan > 0 ? (numRunTime / effectivePlan) : 0;
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
  const totals = { E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, AH: 0, AG_LOSS: 0 };
  EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => totals[c.col] = 0);
  const rowsPerDay = getRowsPerDay();

  SheetState.rows.forEach(r => {
    totals.E += Number(r.E?.val) || 0;
    totals.F += Number(r.F?.val) || 0;
    totals.G += Number(r.G?.val) || 0;

    const rowIdx = r.row - 6;
    const mIdx = rowIdx % rowsPerDay;
    const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);

    if (groupInfo.isMaster || groupInfo.count === 1) {
      totals.H += Number(r.H?.val) || 0;
      totals.I += Number(r.I?.val) || 0;
      totals.J += (typeof r.J?.val === 'number' ? r.J.val : 0);
      totals.AH += Number(r.AH?.val) || 0;
      totals.AG_LOSS += (typeof r.AG_LOSS?.val === 'number' ? r.AG_LOSS.val : 0);

      EXCEL_COLUMNS.filter(c => c.isDt).forEach(c => {
        totals[c.col] += Number(r[c.col]?.val) || 0;
      });
    }
  });

  totals.J = Math.round(totals.J);
  totals.AH = Math.round(totals.AH);
  totals.AI = totals.H > 0 ? (totals.J / totals.H) : 0;
  totals.AJ = totals.E > 0 ? (totals.F / totals.E) : 0;
  totals.AK = (totals.F + totals.G) > 0 ? (totals.F / (totals.F + totals.G)) : 0;
  totals.AL = totals.AI * totals.AJ * totals.AK;

  SheetState.totals = totals;
}

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
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
  thMachineNo.className = 'text-center align-middle font-bold text-xs border border-slate-300 bg-slate-100 text-slate-900 px-2';
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

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
function renderProductionOutputReport(table) {
  const monthName = MonthYearState.monthNames[MonthYearState.monthIndex];
  const year = MonthYearState.year;

  // Column Widths (% based)
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
  thSec.className = 'sum-header-section';
  thSec.textContent = 'Section Name';
  tr3.appendChild(thSec);

  const thRun = document.createElement('th');
  thRun.rowSpan = 2;
  thRun.className = 'sum-header-avail';
  thRun.innerHTML = 'Production<br>Running (hr)';
  tr3.appendChild(thRun);

  const thCap = document.createElement('th');
  thCap.rowSpan = 2;
  thCap.className = 'sum-header-cap';
  thCap.innerHTML = 'Machine<br>Capacity (Pcs)';
  tr3.appendChild(thCap);

  const thQty = document.createElement('th');
  thQty.colSpan = 2;
  thQty.className = 'sum-header-prod';
  thQty.textContent = 'Production Qty.';
  tr3.appendChild(thQty);

  const thOut = document.createElement('th');
  thOut.rowSpan = 2;
  thOut.className = 'sum-header-perf';
  thOut.innerHTML = 'Standard Wise<br>Production Output<br>(%)';
  tr3.appendChild(thOut);

  const thRem = document.createElement('th');
  thRem.rowSpan = 2;
  thRem.className = 'sum-header-rem';
  thRem.textContent = 'Remarks';
  tr3.appendChild(thRem);

  thead.appendChild(tr3);

  // Row 4: Sub-headers for Production (pcs) and Rejection (Pcs)
  const tr4 = document.createElement('tr');

  const thProdPcs = document.createElement('th');
  thProdPcs.className = 'sum-header-prod';
  thProdPcs.textContent = 'Production (pcs)';
  tr4.appendChild(thProdPcs);

  const thRejPcs = document.createElement('th');
  thRejPcs.className = 'sum-header-rej';
  thRejPcs.textContent = 'Rejection (Pcs)';
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
    tdSec.className = 'sum-cell-sec';
    tdSec.textContent = dept.name;
    tr.appendChild(tdSec);

    // Col 2: Production Running (hr)
    const tdRun = document.createElement('td');
    tdRun.className = 'sum-cell-avail font-mono';
    tdRun.textContent = tabSummary.runningMins > 0 ? (tabSummary.runningMins / 60).toFixed(1) : '-';
    tr.appendChild(tdRun);

    // Col 3: Machine Capacity (Pcs)
    const tdCap = document.createElement('td');
    tdCap.className = 'sum-cell-cap font-mono';
    tdCap.textContent = tabSummary.capacityPcs > 0 ? tabSummary.capacityPcs.toLocaleString() : '-';
    tr.appendChild(tdCap);

    // Col 4: Production (pcs)
    const tdProd = document.createElement('td');
    tdProd.className = 'sum-cell-prod font-mono';
    tdProd.textContent = tabSummary.actualPrdPcs > 0 ? tabSummary.actualPrdPcs.toLocaleString() : '-';
    tr.appendChild(tdProd);

    // Col 5: Rejection (Pcs)
    const tdRej = document.createElement('td');
    tdRej.className = 'sum-cell-rej font-mono';
    tdRej.textContent = tabSummary.rejectionPcs > 0 ? (tabSummary.rejectionPcs % 1 !== 0 ? tabSummary.rejectionPcs.toFixed(1) : tabSummary.rejectionPcs.toLocaleString()) : '0';
    tr.appendChild(tdRej);

    // Col 6: Standard Wise Production Output (%)
    const tdOut = document.createElement('td');
    tdOut.className = 'sum-cell-perf font-mono';
    tdOut.textContent = `${(tabSummary.outputPct * 100).toFixed(1)}%`;
    tr.appendChild(tdOut);

    // Col 7: Remarks
    const tdRem = document.createElement('td');
    tdRem.className = 'sum-cell-rem';
    tdRem.textContent = tabSummary.remarks || '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row
  const trTotal = document.createElement('tr');
  trTotal.className = 'sum-total-row';

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'sum-total-title font-mono';
  tdTotTitle.textContent = 'Total';
  trTotal.appendChild(tdTotTitle);

  const tdTotRun = document.createElement('td');
  tdTotRun.className = 'sum-total-run font-mono';
  tdTotRun.textContent = grandRunning > 0 ? (grandRunning / 60).toFixed(1) : '0';
  trTotal.appendChild(tdTotRun);

  const tdTotCap = document.createElement('td');
  tdTotCap.className = 'sum-total-cap font-mono';
  tdTotCap.textContent = grandCapacity.toLocaleString();
  trTotal.appendChild(tdTotCap);

  const tdTotProd = document.createElement('td');
  tdTotProd.className = 'sum-total-prod font-mono';
  tdTotProd.textContent = grandProduction.toLocaleString();
  trTotal.appendChild(tdTotProd);

  const tdTotRej = document.createElement('td');
  tdTotRej.className = 'sum-total-rej font-mono';
  tdTotRej.textContent = grandRejection > 0 ? (grandRejection % 1 !== 0 ? grandRejection.toFixed(1) : grandRejection.toLocaleString()) : '0';
  trTotal.appendChild(tdTotRej);

  const overallOutputPct = grandCapacity > 0 ? (grandProduction / grandCapacity) : 0;
  const tdTotOut = document.createElement('td');
  tdTotOut.className = 'sum-total-perf font-mono';
  tdTotOut.textContent = `${(overallOutputPct * 100).toFixed(1)}%`;
  trTotal.appendChild(tdTotOut);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'sum-cell-rem';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

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
    th.style.border = '1px solid #CBD5E1';
    th.style.backgroundColor = '#F1F5F9';
    th.style.color = '#0F172A';
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

    tr.innerHTML = `
      <td style="font-weight:700; color:#0F172A; text-align:left; border:1px solid #E2E8F0; padding:6px 8px; background:#F8FAFC;">
        ${s.name}
      </td>
      <td style="text-align:right; font-weight:600; border:1px solid #E2E8F0; padding:6px 8px;">
        ${s.runHours > 0 ? s.runHours.toFixed(1) : '-'}
      </td>
      <td style="text-align:right; font-weight:600; border:1px solid #E2E8F0; padding:6px 8px;">
        ${s.capacityPcs > 0 ? s.capacityPcs.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; border:1px solid #E2E8F0; padding:6px 8px;">
        ${s.targetPerHr > 0 ? s.targetPerHr.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; font-weight:bold; color:#0F172A; border:1px solid #CBD5E1; padding:6px 8px; background:#FFFFFF;">
        ${s.actualPrdPcs > 0 ? s.actualPrdPcs.toLocaleString() : '-'}
      </td>
      <td style="text-align:right; border:1px solid #CBD5E1; padding:6px 8px;">
        ${s.actualPerHr > 0 ? s.actualPerHr.toLocaleString() : '-'}
      </td>
      <td style="text-align:center; font-weight:bold; border:1px solid #CBD5E1; padding:6px 8px; ${isHighAch ? 'background:#FFFFFF; color:#0F172A;' : ''}">
        ${s.capacityPcs > 0 ? (s.achievement * 100).toFixed(1) + '%' : '-'}
      </td>
      <td style="text-align:right; font-weight:bold; color:#0F172A; border:1px solid #CBD5E1; padding:6px 8px; background:#FFFFFF;">
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
  trTotal.style.backgroundColor = '#F1F5F9';
  trTotal.style.fontWeight = 'bold';
  trTotal.style.borderTop = '1.5px solid #64748B';
  trTotal.style.borderBottom = '3px double #0F172A';

  trTotal.innerHTML = `
    <td style="text-align:left; border:1px solid #CBD5E1; padding:8px 8px; color:#0F172A;">Total:</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalRunHrs > 0 ? totalRunHrs.toFixed(1) : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${grandTotals.cap > 0 ? grandTotals.cap.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalTargetPerHr > 0 ? totalTargetPerHr.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px; color:#0F172A;">${grandTotals.act > 0 ? grandTotals.act.toLocaleString() : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px;">${totalActualPerHr > 0 ? totalActualPerHr.toLocaleString() : '-'}</td>
    <td style="text-align:center; border:1px solid #CBD5E1; padding:8px 8px; color:#0F172A;">${grandTotals.cap > 0 ? (totalAch * 100).toFixed(1) + '%' : '-'}</td>
    <td style="text-align:right; border:1px solid #CBD5E1; padding:8px 8px; color:#0F172A;">${grandTotals.rej.toLocaleString()}</td>
  `;
  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
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

  // Row 2: Header Title
  const tr2 = document.createElement('tr');
  tr2.className = 'mep-banner-row2';
  const th2 = document.createElement('th');
  th2.colSpan = 7;
  th2.textContent = `MEP Fan Limited- Down Time & Running Time Status (${monthName} ${year})`;
  tr2.appendChild(th2);
  thead.appendChild(tr2);

  // Row 3: Header Row
  const tr3 = document.createElement('tr');

  const headers = [
    { label: 'Machine No.', cls: 'sum-header-section' },
    { label: 'Planned\nProduction Time\n(mins)', cls: 'sum-header-cap' },
    { label: 'Production Run\nTime (Mins)', cls: 'sum-header-avail' },
    { label: 'Machine Down\nTime (Mins)', cls: 'sum-header-down' },
    { label: 'Production\nRunning Time (%)', cls: 'sum-header-perf' },
    { label: 'Production Down\nTime (%)', cls: 'sum-header-rej' },
    { label: "Remark's", cls: 'sum-header-rem' }
  ];

  headers.forEach(h => {
    const th = document.createElement('th');
    th.className = h.cls;
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
    tdName.className = 'sum-cell-sec';
    tdName.textContent = dept.name;
    tr.appendChild(tdName);

    // Col 2: Planned Production Time (mins)
    const tdPlanned = document.createElement('td');
    tdPlanned.className = 'sum-cell-cap font-mono';
    tdPlanned.textContent = status.plannedTimeMins > 0 ? status.plannedTimeMins.toLocaleString() : '-';
    tr.appendChild(tdPlanned);

    // Col 3: Production Run Time (Mins)
    const tdRun = document.createElement('td');
    tdRun.className = 'sum-cell-avail font-mono';
    tdRun.textContent = status.runTimeMins !== 0 ? status.runTimeMins.toLocaleString() : '0';
    tr.appendChild(tdRun);

    // Col 4: Machine Down Time (Mins)
    const tdDown = document.createElement('td');
    tdDown.className = 'sum-cell-down font-mono';
    tdDown.textContent = status.downTimeMins > 0 ? status.downTimeMins.toLocaleString() : '-';
    tr.appendChild(tdDown);

    // Col 5: Production Running Time (%)
    const tdRunPct = document.createElement('td');
    tdRunPct.className = 'sum-cell-perf font-mono';
    tdRunPct.textContent = `${(status.runTimePct * 100).toFixed(1)}%`;
    tr.appendChild(tdRunPct);

    // Col 6: Production Down Time (%)
    const tdDownPct = document.createElement('td');
    tdDownPct.className = 'sum-cell-rej font-mono';
    tdDownPct.textContent = `${(status.downTimePct * 100).toFixed(1)}%`;
    tr.appendChild(tdDownPct);

    // Col 7: Remark's
    const tdRem = document.createElement('td');
    tdRem.className = 'sum-cell-rem';
    tdRem.textContent = '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row
  const trTotal = document.createElement('tr');
  trTotal.className = 'sum-total-row';

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'sum-total-title font-mono';
  tdTotTitle.textContent = 'SubTotal';
  trTotal.appendChild(tdTotTitle);

  const tdTotPlanned = document.createElement('td');
  tdTotPlanned.className = 'sum-total-planned font-mono';
  tdTotPlanned.textContent = grandPlanned.toLocaleString();
  trTotal.appendChild(tdTotPlanned);

  const tdTotRun = document.createElement('td');
  tdTotRun.className = 'sum-total-run font-mono';
  tdTotRun.textContent = grandRun.toLocaleString();
  trTotal.appendChild(tdTotRun);

  const tdTotDown = document.createElement('td');
  tdTotDown.className = 'sum-total-down font-mono';
  tdTotDown.textContent = grandDown.toLocaleString();
  trTotal.appendChild(tdTotDown);

  const overallRunPct = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
  const overallDownPct = grandPlanned > 0 ? (grandDown / grandPlanned) : 0;

  const tdTotRunPct = document.createElement('td');
  tdTotRunPct.className = 'sum-total-perf font-mono';
  tdTotRunPct.textContent = `${(overallRunPct * 100).toFixed(1)}%`;
  trTotal.appendChild(tdTotRunPct);

  const tdTotDownPct = document.createElement('td');
  tdTotDownPct.className = 'sum-total-rej font-mono';
  tdTotDownPct.textContent = `${(overallDownPct * 100).toFixed(1)}%`;
  trTotal.appendChild(tdTotDownPct);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'sum-cell-rem';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

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

// ──────────────────────────────────────────────────────────────────────────
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
  thSec.className = 'sum-header-section';
  thSec.textContent = 'Section';
  tr3.appendChild(thSec);

  const thCap = document.createElement('th');
  thCap.rowSpan = 2;
  thCap.className = 'sum-header-cap';
  thCap.innerHTML = 'Machine<br>Capacity (Pcs)';
  tr3.appendChild(thCap);

  const thQty = document.createElement('th');
  thQty.colSpan = 2;
  thQty.className = 'sum-header-prod';
  thQty.textContent = 'Production Qty.';
  tr3.appendChild(thQty);

  const thAvail = document.createElement('th');
  thAvail.rowSpan = 2;
  thAvail.className = 'sum-header-avail';
  thAvail.innerHTML = 'Availability<br>(%)';
  tr3.appendChild(thAvail);

  const thPerf = document.createElement('th');
  thPerf.rowSpan = 2;
  thPerf.className = 'sum-header-perf';
  thPerf.innerHTML = 'Performance<br>(%)';
  tr3.appendChild(thPerf);

  const thQual = document.createElement('th');
  thQual.rowSpan = 2;
  thQual.className = 'sum-header-qual';
  thQual.innerHTML = 'Quality<br>(%)';
  tr3.appendChild(thQual);

  const thOEE = document.createElement('th');
  thOEE.rowSpan = 2;
  thOEE.className = 'sum-header-oee';
  thOEE.innerHTML = 'OEE (%)';
  tr3.appendChild(thOEE);

  const thRem = document.createElement('th');
  thRem.rowSpan = 2;
  thRem.className = 'sum-header-rem';
  thRem.textContent = "Remark's";
  tr3.appendChild(thRem);

  thead.appendChild(tr3);

  // Row 4: Sub-headers
  const tr4 = document.createElement('tr');

  const thTotProd = document.createElement('th');
  thTotProd.className = 'sum-header-prod';
  thTotProd.innerHTML = 'Total Production<br>(pcs)';
  tr4.appendChild(thTotProd);

  const thRej = document.createElement('th');
  thRej.className = 'sum-header-rej';
  thRej.innerHTML = 'Rejection<br>(Pcs)';
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
    tdSec.className = 'sum-cell-sec';
    tdSec.textContent = dept.name;
    tr.appendChild(tdSec);

    // Col 2: Machine Capacity (Pcs)
    const tdCap = document.createElement('td');
    tdCap.className = 'sum-cell-cap font-mono';
    tdCap.textContent = oeeData.capacityPcs > 0 ? oeeData.capacityPcs.toLocaleString() : '-';
    tr.appendChild(tdCap);

    // Col 3: Total Production (pcs)
    const tdProd = document.createElement('td');
    tdProd.className = 'sum-cell-prod font-mono';
    tdProd.textContent = oeeData.totalProduction > 0 ? oeeData.totalProduction.toLocaleString() : '-';
    tr.appendChild(tdProd);

    // Col 4: Rejection (Pcs)
    const tdRej = document.createElement('td');
    tdRej.className = 'sum-cell-rej font-mono';
    tdRej.textContent = oeeData.rejectionPcs > 0 ? (oeeData.rejectionPcs % 1 !== 0 ? oeeData.rejectionPcs.toFixed(1) : oeeData.rejectionPcs.toLocaleString()) : '0';
    tr.appendChild(tdRej);

    // Col 5: Availability (%)
    const tdAvail = document.createElement('td');
    tdAvail.className = 'sum-cell-avail font-mono';
    tdAvail.textContent = `${(oeeData.availability * 100).toFixed(0)}%`;
    tr.appendChild(tdAvail);

    // Col 6: Performance (%)
    const tdPerf = document.createElement('td');
    tdPerf.className = 'sum-cell-perf font-mono';
    tdPerf.textContent = `${(oeeData.performance * 100).toFixed(0)}%`;
    tr.appendChild(tdPerf);

    // Col 7: Quality (%)
    const tdQual = document.createElement('td');
    tdQual.className = 'sum-cell-qual font-mono';
    tdQual.textContent = `${(oeeData.quality * 100).toFixed(1)}%`;
    tr.appendChild(tdQual);

    // Col 8: OEE (%)
    const tdOEE = document.createElement('td');
    tdOEE.className = 'sum-cell-oee font-mono';
    tdOEE.textContent = `${(oeeData.oee * 100).toFixed(1)}%`;
    tr.appendChild(tdOEE);

    // Col 9: Remark's
    const tdRem = document.createElement('td');
    tdRem.className = 'sum-cell-rem';
    tdRem.textContent = oeeData.remarks || '';
    tr.appendChild(tdRem);

    tbody.appendChild(tr);
  });

  // Total Row
  const trTotal = document.createElement('tr');
  trTotal.className = 'sum-total-row';

  const tdTotTitle = document.createElement('td');
  tdTotTitle.className = 'sum-total-title font-mono';
  tdTotTitle.textContent = 'Total';
  trTotal.appendChild(tdTotTitle);

  const tdTotCap = document.createElement('td');
  tdTotCap.className = 'sum-total-cap font-mono';
  tdTotCap.textContent = grandCap.toLocaleString();
  trTotal.appendChild(tdTotCap);

  const tdTotProd = document.createElement('td');
  tdTotProd.className = 'sum-total-prod font-mono';
  tdTotProd.textContent = grandProd.toLocaleString();
  trTotal.appendChild(tdTotProd);

  const tdTotRej = document.createElement('td');
  tdTotRej.className = 'sum-total-rej font-mono';
  tdTotRej.textContent = grandRej > 0 ? (grandRej % 1 !== 0 ? grandRej.toFixed(1) : grandRej.toLocaleString()) : '0';
  trTotal.appendChild(tdTotRej);

  // Overall KPIs
  const overallAvail = grandPlanned > 0 ? (grandRun / grandPlanned) : 0;
  const overallPerf = grandCap > 0 ? (grandProd / grandCap) : 0;
  const overallQual = (grandProd + grandRej) > 0 ? (grandProd / (grandProd + grandRej)) : 1.0;
  const overallOEE = overallAvail * overallPerf * overallQual;

  const tdTotAvail = document.createElement('td');
  tdTotAvail.className = 'sum-total-avail font-mono';
  tdTotAvail.textContent = `${(overallAvail * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotAvail);

  const tdTotPerf = document.createElement('td');
  tdTotPerf.className = 'sum-total-perf font-mono';
  tdTotPerf.textContent = `${(overallPerf * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotPerf);

  const tdTotQual = document.createElement('td');
  tdTotQual.className = 'sum-total-qual font-mono';
  tdTotQual.textContent = `${(overallQual * 100).toFixed(0)}%`;
  trTotal.appendChild(tdTotQual);

  const tdTotOEE = document.createElement('td');
  tdTotOEE.className = 'sum-total-oee font-mono';
  tdTotOEE.textContent = `${(overallOEE * 100).toFixed(1)}%`;
  trTotal.appendChild(tdTotOEE);

  const tdTotRem = document.createElement('td');
  tdTotRem.className = 'sum-cell-rem';
  tdTotRem.textContent = '';
  trTotal.appendChild(tdTotRem);

  tbody.appendChild(trTotal);
  table.appendChild(tbody);
}

// ──────────────────────────────────────────────────────────────────────────
const YEARLY_OEE_PRESET_2026 = {
  0: { name: 'January', capacityPcs: 1124393, totalProduction: 611910, rejectionPcs: 2815.9, achievement: 0.544, availability: 0.74, performance: 0.75, quality: 0.99, oee: 0.60, isFixed: true },
  1: { name: 'February', capacityPcs: 618970, totalProduction: 370274, rejectionPcs: 1779, achievement: 0.598, availability: 0.73, performance: 0.79, quality: 0.99, oee: 0.61, isFixed: true },
  2: { name: 'March', capacityPcs: 609528, totalProduction: 281886, rejectionPcs: 1055, achievement: 0.462, availability: 0.57, performance: 0.66, quality: 1.00, oee: 0.42, isFixed: true },
  3: { name: 'April', capacityPcs: 1111210, totalProduction: 677516, rejectionPcs: 2648, achievement: 0.610, availability: 0.76, performance: 0.73, quality: 1.00, oee: 0.58, isFixed: true },
  4: { name: 'May', capacityPcs: 30000, totalProduction: 30088, rejectionPcs: 579, achievement: 1.003, availability: 0.89, performance: 0.99, quality: 0.987, oee: 0.87, isFixed: true },
  5: { name: 'Jun', capacityPcs: 211140, totalProduction: 207640, rejectionPcs: 802, achievement: 0.983, availability: 0.94, performance: 0.98, quality: 0.995, oee: 0.92, isFixed: true },
  6: { name: 'July', capacityPcs: 146537, totalProduction: 144720, rejectionPcs: 476, achievement: 0.988, availability: 0.93, performance: 0.97, quality: 0.99, oee: 0.90, isFixed: true }
};

function getYearlyOEESummary(year) {
  return MonthYearState.monthNames.map((name, mIdx) => {
    // If year 2026 and we have historical preset for past months (0 to 6 = Jan to Jul)
    if (year === 2026 && YEARLY_OEE_PRESET_2026[mIdx] && mIdx < MonthYearState.monthIndex) {
      const p = YEARLY_OEE_PRESET_2026[mIdx];
      return {
        monthName: p.name || name,
        monthIndex: mIdx,
        capacityPcs: p.capacityPcs,
        totalProduction: p.totalProduction,
        rejectionPcs: p.rejectionPcs,
        plannedTimeMins: 0,
        runTimeMins: 0,
        availability: p.availability,
        performance: p.performance,
        quality: p.quality,
        oee: p.oee,
        achievement: p.achievement,
        isFixed: true,
        remarks: ''
      };
    }

    // Live calculation from department sheets
    let capacityPcs = 0;
    let totalProduction = 0;
    let rejectionPcs = 0;
    let plannedTimeMins = 0;
    let runTimeMins = 0;
    let hasData = false;

    getActiveSummaryDepts().forEach(dept => {
      const summary = getTabOEESummary(dept.id, year, mIdx);
      if (summary.capacityPcs > 0 || summary.totalProduction > 0 || summary.plannedTimeMins > 0) {
        hasData = true;
      }
      capacityPcs += summary.capacityPcs;
      totalProduction += summary.totalProduction;
      rejectionPcs += summary.rejectionPcs;
      plannedTimeMins += summary.plannedTimeMins;
      runTimeMins += summary.runTimeMins;
    });

    const availability = plannedTimeMins > 0 ? (runTimeMins / plannedTimeMins) : 0;
    const performance = capacityPcs > 0 ? (totalProduction / capacityPcs) : 0;
    const quality = (totalProduction + rejectionPcs) > 0 ? (totalProduction / (totalProduction + rejectionPcs)) : (totalProduction > 0 ? 1.0 : 0);
    const oee = availability * performance * quality;
    const achievement = capacityPcs > 0 ? (totalProduction / capacityPcs) : 0;

    return {
      monthName: (mIdx === 5 ? 'Jun' : name),
      monthIndex: mIdx,
      capacityPcs,
      totalProduction,
      rejectionPcs,
      plannedTimeMins,
      runTimeMins,
      availability,
      performance,
      quality,
      oee,
      achievement,
      isFixed: hasData,
      remarks: ''
    };
  });
}

function renderYearlyOEESummaryReport(table) {
  const year = MonthYearState.year;
  const yearlyData = getYearlyOEESummary(year);

  // Column Widths (% based to fit 100% full screen)
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

  // Row 1: MEP FAN LTD. (Royal Deep Navy)
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

  // Row 3: Header Top Row (Cyan/Sky Blue matching template)
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
    th.className = 'yearly-template-header';
    th.innerHTML = h.label.replace(/\n/g, '<br>');
    tr3.appendChild(th);
  });

  thead.appendChild(tr3);

  // Row 4: Sub-headers for Production (pcs) and Rejection (Pcs)
  const tr4 = document.createElement('tr');

  const thTotProd = document.createElement('th');
  thTotProd.className = 'yearly-template-header';
  thTotProd.innerHTML = 'Total<br>Production<br>(pcs)';
  tr4.appendChild(thTotProd);

  const thRej = document.createElement('th');
  thRej.className = 'yearly-template-header';
  thRej.innerHTML = 'Rejection<br>(Pcs)';
  tr4.appendChild(thRej);

  thead.appendChild(tr4);
  table.appendChild(thead);

  // Tbody
  const tbody = document.createElement('tbody');

  yearlyData.forEach((m, idx) => {
    const isRunningMonth = (idx === MonthYearState.monthIndex);

    // Row 1: Total
    const trA = document.createElement('tr');
    trA.className = isRunningMonth ? 'yearly-row-active-a' : 'yearly-row-normal-a';

    // Col 1: Month (rowSpan 2)
    const tdMonth = document.createElement('td');
    tdMonth.rowSpan = 2;
    tdMonth.className = isRunningMonth ? 'yearly-month-cell running-month-active' : 'yearly-month-cell';
    tdMonth.textContent = m.monthName;
    trA.appendChild(tdMonth);

    // Col 2: Detail
    const tdDetailA = document.createElement('td');
    tdDetailA.className = isRunningMonth ? 'yearly-detail-cell running-detail-active' : 'yearly-detail-cell';
    tdDetailA.textContent = 'Total';
    trA.appendChild(tdDetailA);

    // Col 3: Machine Capacity
    const tdCap = document.createElement('td');
    tdCap.className = isRunningMonth ? 'yearly-data-cell running-data-active font-mono' : 'yearly-data-cell font-mono';
    tdCap.textContent = m.capacityPcs > 0 ? m.capacityPcs.toLocaleString() : '-';
    trA.appendChild(tdCap);

    // Col 4: Total Production
    const tdProd = document.createElement('td');
    tdProd.className = isRunningMonth ? 'yearly-data-cell running-data-active font-mono' : 'yearly-data-cell font-mono';
    tdProd.textContent = m.totalProduction > 0 ? m.totalProduction.toLocaleString() : '-';
    trA.appendChild(tdProd);

    // Col 5: Rejection
    const tdRej = document.createElement('td');
    tdRej.className = isRunningMonth ? 'yearly-data-cell running-data-active font-mono' : 'yearly-data-cell font-mono';
    tdRej.textContent = m.rejectionPcs > 0 ? (m.rejectionPcs % 1 !== 0 ? m.rejectionPcs.toFixed(1) : m.rejectionPcs.toLocaleString()) : '0';
    trA.appendChild(tdRej);

    // Col 6: Availability (rowSpan 2)
    const tdAvail = document.createElement('td');
    tdAvail.rowSpan = 2;
    tdAvail.className = isRunningMonth ? 'yearly-kpi-cell running-kpi-active font-mono' : 'yearly-kpi-cell font-mono';
    tdAvail.textContent = (m.capacityPcs > 0 || m.isFixed) ? `${Math.round(m.availability * 100)}%` : '-';
    trA.appendChild(tdAvail);

    // Col 7: Performance (rowSpan 2)
    const tdPerf = document.createElement('td');
    tdPerf.rowSpan = 2;
    tdPerf.className = isRunningMonth ? 'yearly-kpi-cell running-kpi-active font-mono' : 'yearly-kpi-cell font-mono';
    tdPerf.textContent = (m.capacityPcs > 0 || m.isFixed) ? `${Math.round(m.performance * 100)}%` : '-';
    trA.appendChild(tdPerf);

    // Col 8: Quality (rowSpan 2)
    const tdQual = document.createElement('td');
    tdQual.rowSpan = 2;
    tdQual.className = isRunningMonth ? 'yearly-kpi-cell running-kpi-active font-mono' : 'yearly-kpi-cell font-mono';
    if (m.capacityPcs > 0 || m.isFixed) {
      const qVal = m.quality * 100;
      tdQual.textContent = (qVal % 1 !== 0) ? `${qVal.toFixed(1)}%` : `${Math.round(qVal)}%`;
    } else {
      tdQual.textContent = '-';
    }
    trA.appendChild(tdQual);

    // Col 9: OEE (rowSpan 2)
    const tdOEE = document.createElement('td');
    tdOEE.rowSpan = 2;
    tdOEE.className = isRunningMonth ? 'yearly-kpi-cell running-kpi-active font-mono' : 'yearly-kpi-cell font-mono';
    tdOEE.textContent = (m.capacityPcs > 0 || m.isFixed) ? `${Math.round(m.oee * 100)}%` : '-';
    trA.appendChild(tdOEE);

    // Col 10: Remarks (rowSpan 2)
    const tdRem = document.createElement('td');
    tdRem.rowSpan = 2;
    tdRem.className = isRunningMonth ? 'yearly-remarks-cell running-data-active' : 'yearly-remarks-cell';
    tdRem.textContent = m.remarks || '';
    trA.appendChild(tdRem);

    tbody.appendChild(trA);

    // Row 2: Total Acheivement (%)
    const trB = document.createElement('tr');
    trB.className = isRunningMonth ? 'yearly-row-active-b' : 'yearly-row-normal-b';

    // Col 2: Detail
    const tdDetailB = document.createElement('td');
    tdDetailB.className = isRunningMonth ? 'yearly-detail-cell running-detail-active' : 'yearly-detail-cell';
    tdDetailB.textContent = 'Total Acheivement (%)';
    trB.appendChild(tdDetailB);

    // Col 3, 4, 5: Merged Total Acheivement Value (colSpan 3)
    const tdAch = document.createElement('td');
    tdAch.colSpan = 3;
    const achVal = m.achievement;
    const isOver95 = (achVal >= 0.95 && m.capacityPcs > 0);

    if (isRunningMonth) {
      tdAch.className = 'yearly-ach-running font-mono';
      tdAch.textContent = `${(achVal * 100).toFixed(1)}%`;
    } else if (isOver95) {
      tdAch.className = 'yearly-ach-highlight font-mono';
      tdAch.textContent = `${(achVal * 100).toFixed(1)}%`;
    } else if (m.capacityPcs > 0 || m.isFixed) {
      tdAch.className = 'yearly-ach-cell font-mono';
      tdAch.textContent = `${(achVal * 100).toFixed(1)}%`;
    } else {
      tdAch.className = 'yearly-ach-cell font-mono';
      tdAch.textContent = '-';
    }
    trB.appendChild(tdAch);

    tbody.appendChild(trB);
  });

  table.appendChild(tbody);
}

function renderExcelTable() {
  const table = document.getElementById('excelMainTable');
  if (!table) return;

  table.replaceChildren();

  // If viewing Executive Summary Reports
  if (ACTIVE_TAB === 'summary_downtime') {
    table.className = 'mep-excel-table mep-summary-table mep-dt-summary-table';
    renderSummaryDowntimeReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_production') {
    table.className = 'mep-excel-table mep-summary-table';
    renderProductionOutputReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_status') {
    table.className = 'mep-excel-table mep-summary-table';
    renderDowntimeRunningStatusReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_oee') {
    table.className = 'mep-excel-table mep-summary-table';
    renderOEEReport(table);
    return;
  }
  if (ACTIVE_TAB === 'summary_oee_yearly') {
    table.className = 'mep-excel-table mep-summary-table';
    renderYearlyOEESummaryReport(table);
    return;
  }

  table.className = 'mep-excel-table mep-entry-table';

  // Column Widths from EXCEL_COLUMNS definition
  const colgroup = document.createElement('colgroup');
  EXCEL_COLUMNS.forEach(c => {
    const col = document.createElement('col');
    col.style.width = `${c.width}px`;
    col.style.minWidth = `${c.width}px`;
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
      if (colLetter === 'AM') td.dataset.formula = '=IFERROR(F5/E5,0)';
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

    const tr = document.createElement('tr');
    tr.dataset.row = r;
    tr.dataset.day = dayNum;

    if (locked) tr.classList.add('row-locked');
    if (isFriday) tr.classList.add('row-friday');
    if (isToday) tr.classList.add('row-today');
    if (isDayEnd) tr.classList.add('row-day-end');

    EXCEL_COLUMNS.forEach((colDef, cIdx) => {
      const colLetter = colDef.col;
      const isDateDayShiftCol = ['A', 'B', 'C'].includes(colLetter);

      // 1. Date, Day, Shift (Cols A, B, C) - Merged per Day
      if (isDateDayShiftCol && mIdx > 0) {
        return; // Skip slave rows
      }

      // 2. Time-Related Columns (Planned Time, Downtimes, Total DT, Runtime, Availability)
      const tabInfo = SHEET_TABS[ACTIVE_TAB];
      const isTimeCol = TIME_COLUMNS.includes(colLetter);
      let groupInfo = { isMaster: true, span: 1, count: 1, isSlave: false };
      if (tabInfo && tabInfo.timeGroups && isTimeCol) {
        groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);
        if (groupInfo.isSlave) {
          return; // Skip slave rows for merged continuous line time columns
        }
      }

      const td = document.createElement('td');
      if (isDateDayShiftCol && rowsPerDay > 1) {
        td.rowSpan = rowsPerDay;
      } else if (tabInfo && tabInfo.timeGroups && isTimeCol && groupInfo.count > 1 && groupInfo.isMaster) {
        td.rowSpan = groupInfo.count;
      }
      td.dataset.row = r;
      td.dataset.col = colLetter;
      td.dataset.cidx = cIdx;

      if (colDef.isReadOnly) {
        td.className = 'cell-readonly-fixed';
        if (isDateDayShiftCol && mIdx > 0) {
          td.classList.add('cell-date-secondary');
        }
      } else if (colLetter === 'J') {
        td.className = 'cell-formula-runtime';
      } else if (colDef.isLossDt) {
        td.className = 'cell-loss-dt';
      } else if (colDef.isDt) {
        td.className = 'cell-downtime';
      } else if (colLetter === 'AH') {
        td.className = 'cell-total-dt';
      } else if (colDef.isFormula || colDef.isPercent) {
        td.className = 'cell-kpi';
      } else {
        td.className = 'cell-white';
      }

      if (isDayEnd) {
        td.classList.add('cell-day-end-border');
      }

      const cellVal = rowObj[colLetter]?.val;
      td.textContent = formatCellValue(cellVal, colDef);

      const hasEnteredVal = (cellVal !== null && cellVal !== undefined && cellVal !== '' && cellVal !== '-' && cellVal != 0);
      if (hasEnteredVal || (colLetter === 'AM' && cellVal && String(cellVal).trim().length > 0)) {
        td.classList.add('cell-has-data');
      } else {
        td.classList.remove('cell-has-data');
      }

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
    return `${Math.round(num * 100)}%`;
  }

  if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
    const n = Number(val);
    if (!isNaN(n)) {
      const rounded = Math.round(n);
      return rounded !== 0 ? String(rounded) : '0';
    }
  }

  return String(val);
}

// ──────────────────────────────────────────────────────────────────────────
function selectCell(colLetter, rowNum, resetRange = true) {
  if (ACTIVE_TAB === 'summary_downtime' || ACTIVE_TAB === 'summary_production') return;
  if (isRowLocked(rowNum)) return;

  if (SheetState.isEditing && SheetState.activeInput) {
    SheetState.activeInput.blur();
  }

  clearSelectionStyles();

  const rowsPerDay = getRowsPerDay();
  const mIdx = (rowNum - 6) % rowsPerDay;
  const tabInfo = SHEET_TABS[ACTIVE_TAB];

  let targetRow = rowNum;
  if (['A', 'B', 'C'].includes(colLetter)) {
    targetRow = rowNum - mIdx;
  } else if (tabInfo && tabInfo.timeGroups && TIME_COLUMNS.includes(colLetter)) {
    const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);
    if (groupInfo.isSlave) {
      targetRow = (rowNum - mIdx) + groupInfo.masterIdx;
    }
  }

  let targetTd = document.querySelector(`.mep-excel-table [data-col="${colLetter}"][data-row="${targetRow}"]`);
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

// ──────────────────────────────────────────────────────────────────────────
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
    showToast('Permission Denied: This operation is restricted.', 'warning');
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
    showToast('Operation completed successfully.', 'success');
  };

  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', onMouseUp);
  document.addEventListener('keydown', onKeyDown);
  document.addEventListener('keyup', onKeyUp);
}

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
function startCellEdit(td, colLetter, rowNum, colDef, initialChar = null) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('Permission Denied: This operation is restricted.', 'warning');
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
      // Prevent bubbling so window shortcuts don't intercept editing
      e.stopPropagation();
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
  if (initialChar !== null) {
    input.setSelectionRange(input.value.length, input.value.length);
  } else if (input.select) {
    input.select();
  }

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

  const affectedRows = [];
  const tabInfo = SHEET_TABS[ACTIVE_TAB];
  if (tabInfo && tabInfo.timeGroups) {
    const rowsPerDay = getRowsPerDay();
    const mIdx = (rowNum - 6) % rowsPerDay;
    const groupInfo = getTimeGroupInfo(ACTIVE_TAB, mIdx);
    const dayStart = rowNum - mIdx;

    if (TIME_COLUMNS.includes(colLetter)) {
      for (let i = 0; i < groupInfo.count; i++) {
        const gRowObj = SheetState.rows.find(r => r.row === (dayStart + groupInfo.masterIdx + i));
        if (gRowObj) {
          if (!gRowObj[colLetter]) gRowObj[colLetter] = {};
          gRowObj[colLetter].val = cleanNum;
        }
      }
    }

    for (let i = 0; i < rowsPerDay; i++) {
      const targetRowObj = SheetState.rows.find(r => r.row === (dayStart + i));
      if (targetRowObj) {
        recalculateRow(targetRowObj);
        affectedRows.push(targetRowObj);
      }
    }
  } else {
    recalculateRow(rowObj);
    affectedRows.push(rowObj);
  }

  recalculateTotalRow();
  saveSheetData(false);
  GoogleSheetsRealtimeEngine.pushRowsBatch(ACTIVE_TAB, MonthYearState.year, MonthYearState.monthIndex, affectedRows);
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

    const hasEnteredVal = (cellVal !== null && cellVal !== undefined && cellVal !== '' && cellVal !== '-' && cellVal != 0);
    if (hasEnteredVal || (colLetter === 'AM' && cellVal && String(cellVal).trim().length > 0)) {
      td.classList.add('cell-has-data');
    } else {
      td.classList.remove('cell-has-data');
    }

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

// ──────────────────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────────────────
// NAVIGATION & QUICK ENTRY
// ──────────────────────────────────────────────────────────────────────────
function navigateSelection(dCol, dRow) {
  if (SHEET_TABS[ACTIVE_TAB]?.isSummary) return;

  let cur = SheetState.selected;
  if (!cur) {
    selectCell('E', 6);
    return;
  }

  const spanInfo = getCellSpanInfo(cur.colLetter, cur.row);
  const curMasterRow = spanInfo.isMerged ? spanInfo.masterRow : cur.row;
  let nextRow = curMasterRow;

  if (dRow > 0) {
    // Moving Down: jump over merged span in 1 keypress
    nextRow = curMasterRow + (spanInfo.isMerged ? spanInfo.span : dRow);
  } else if (dRow < 0) {
    // Moving Up: jump to previous cell's master row in 1 keypress
    if (spanInfo.isMerged) {
      const prevTarget = curMasterRow - 1;
      if (prevTarget >= 6) {
        const prevSpanInfo = getCellSpanInfo(cur.colLetter, prevTarget);
        nextRow = prevSpanInfo.isMerged ? prevSpanInfo.masterRow : prevTarget;
      } else {
        nextRow = 6;
      }
    } else {
      const targetR = curMasterRow + dRow;
      if (targetR >= 6) {
        const targetSpanInfo = getCellSpanInfo(cur.colLetter, targetR);
        nextRow = targetSpanInfo.isMerged ? targetSpanInfo.masterRow : targetR;
      } else {
        nextRow = 6;
      }
    }
  }

  const curCIdx = EXCEL_COLUMNS.findIndex(c => c.col === cur.colLetter);
  let nextCIdx = curCIdx + dCol;

  const maxR = getMaxActiveRow();
  nextCIdx = Math.max(0, Math.min(EXCEL_COLUMNS.length - 1, nextCIdx));
  nextRow = Math.max(6, Math.min(maxR, nextRow));

  const nextCol = EXCEL_COLUMNS[nextCIdx].col;
  selectCell(nextCol, nextRow);
}

let modalActiveRow = 6;

function openQuickEntryModal(rowNum) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('Permission Denied: Read-only access.', 'warning');
    return;
  }
  const modal = document.getElementById('quickEntryModal');
  if (!modal) return;

  modalActiveRow = rowNum;
  const rowObj = SheetState.rows.find(r => r.row === rowNum);
  if (!rowObj) return;

  const day = getDayFromRow(rowNum);
  const machineName = rowObj.D?.val || 'Machine';
  const titleEl = document.getElementById('modalMachineTitle');
  if (titleEl) {
    titleEl.textContent = `${machineName} — ${day}-${MonthYearState.monthIndex + 1}-${MonthYearState.year}`;
  }

  // Populate Inputs
  const mCap = document.getElementById('mCap');
  const mAct = document.getElementById('mAct');
  const mRej = document.getElementById('mRej');
  const mPlan = document.getElementById('mPlan');
  const mRemarks = document.getElementById('mRemarks');

  if (mCap) mCap.value = rowObj.E?.val || '';
  if (mAct) mAct.value = rowObj.F?.val || '';
  if (mRej) mRej.value = rowObj.G?.val || '';
  if (mPlan) mPlan.value = rowObj.H?.val || '';
  if (mRemarks) mRemarks.value = rowObj.AM?.val || '';

  // Generate 23 Downtime Inputs
  const dtContainer = document.getElementById('modalDtContainer');
  if (dtContainer) {
    dtContainer.innerHTML = '';
    EXCEL_COLUMNS.filter(c => c.isDt).forEach(colDef => {
      const div = document.createElement('div');
      div.className = 'flex flex-col';
      const label = document.createElement('label');
      label.className = 'text-[10px] font-bold text-slate-600 truncate mb-0.5';
      label.title = colDef.label;
      label.textContent = colDef.label;
      const inp = document.createElement('input');
      inp.type = 'number';
      inp.dataset.col = colDef.col;
      inp.className = 'entry-input-field text-xs text-center font-mono py-1';
      inp.value = rowObj[colDef.col]?.val || '';
      inp.placeholder = '0';
      inp.addEventListener('input', updateModalLiveKPI);
      div.appendChild(label);
      div.appendChild(inp);
      dtContainer.appendChild(div);
    });
  }

  [mCap, mAct, mRej, mPlan].forEach(inp => {
    if (inp) {
      inp.removeEventListener('input', updateModalLiveKPI);
      inp.addEventListener('input', updateModalLiveKPI);
    }
  });

  updateModalLiveKPI();
  modal.classList.remove('hidden');
  mAct?.focus();
}

function closeQuickEntryModal() {
  const modal = document.getElementById('quickEntryModal');
  if (modal) modal.classList.add('hidden');
}

function updateModalLiveKPI() {
  const mCap = parseFloat(document.getElementById('mCap')?.value) || 0;
  const mAct = parseFloat(document.getElementById('mAct')?.value) || 0;
  const mRej = parseFloat(document.getElementById('mRej')?.value) || 0;
  const mPlan = parseFloat(document.getElementById('mPlan')?.value) || 0;

  let totalDt = 0;
  document.querySelectorAll('#modalDtContainer input').forEach(inp => {
    totalDt += parseFloat(inp.value) || 0;
  });

  const mTotalDtBadge = document.getElementById('mTotalDtBadge');
  if (mTotalDtBadge) mTotalDtBadge.textContent = `Total DT: ${totalDt} min`;

  const runTime = Math.max(0, mPlan - totalDt);
  const avail = mPlan > 0 ? (runTime / mPlan) * 100 : 0;
  const perf = mCap > 0 ? (mAct / mCap) * 100 : 0;
  const qual = (mAct + mRej) > 0 ? (mAct / (mAct + mRej)) * 100 : (mAct > 0 ? 100 : 0);
  const oee = (avail / 100) * (perf / 100) * (qual / 100) * 100;

  setText('mRunTime', `${runTime} min`);
  setText('mAvail', `${avail.toFixed(0)}%`);
  setText('mPerf', `${perf.toFixed(0)}%`);
  setText('mQual', `${qual.toFixed(0)}%`);
  setText('mOEE', `${oee.toFixed(0)}%`);
}

function saveQuickEntryModal() {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('Permission Denied: Read-only access.', 'warning');
    return;
  }
  const rowObj = SheetState.rows.find(r => r.row === modalActiveRow);
  if (!rowObj) return;

  const mCap = parseFloat(document.getElementById('mCap')?.value);
  const mAct = parseFloat(document.getElementById('mAct')?.value);
  const mRej = parseFloat(document.getElementById('mRej')?.value);
  const mPlan = parseFloat(document.getElementById('mPlan')?.value);
  const mRemarks = document.getElementById('mRemarks')?.value || '';

  if (!isNaN(mCap)) rowObj.E = { val: mCap };
  if (!isNaN(mAct)) rowObj.F = { val: mAct };
  if (!isNaN(mRej)) rowObj.G = { val: mRej };
  if (!isNaN(mPlan)) rowObj.H = { val: mPlan };
  rowObj.AM = { val: mRemarks };

  document.querySelectorAll('#modalDtContainer input').forEach(inp => {
    const col = inp.dataset.col;
    const v = parseFloat(inp.value);
    rowObj[col] = { val: isNaN(v) ? 0 : v };
  });

  recalculateRow(rowObj);
  updateSingleRowDisplay(modalActiveRow);
  recalculateTotalRow();
  updateTotalRowDisplay();
  closeQuickEntryModal();
  saveSheetData(true);
  showToast('Row updated and saved successfully.', 'success');
}

// ──────────────────────────────────────────────────────────────────────────
// SEARCH BAR
// ──────────────────────────────────────────────────────────────────────────
function openSearchBar() {
  const bar = document.getElementById('searchBarContainer');
  if (bar) {
    bar.classList.remove('hidden');
    const inp = document.getElementById('searchInput');
    if (inp) {
      inp.focus();
      inp.select();
    }
  }
}

function closeSearchBar() {
  const bar = document.getElementById('searchBarContainer');
  if (bar) bar.classList.add('hidden');
}

// ──────────────────────────────────────────────────────────────────────────
// CLIPBOARD & RANGE OPERATIONS
// ──────────────────────────────────────────────────────────────────────────
function handleClipboardCopy() {
  if (!SheetState.rangeSelection.start || !SheetState.rangeSelection.end) {
    const cur = SheetState.selected;
    if (cur) {
      const rowObj = SheetState.rows.find(r => r.row === cur.row);
      const val = rowObj ? (rowObj[cur.colLetter]?.val ?? '') : '';
      navigator.clipboard.writeText(String(val)).then(() => {
        showToast('Copied cell to clipboard.', 'info');
      }).catch(() => {});
    }
    return;
  }

  const startC = EXCEL_COLUMNS.findIndex(c => c.col === SheetState.rangeSelection.start.col);
  const endC = EXCEL_COLUMNS.findIndex(c => c.col === SheetState.rangeSelection.end.col);
  const minC = Math.min(startC, endC);
  const maxC = Math.max(startC, endC);

  const minR = Math.min(SheetState.rangeSelection.start.row, SheetState.rangeSelection.end.row);
  const maxR = Math.max(SheetState.rangeSelection.start.row, SheetState.rangeSelection.end.row);

  const lines = [];
  for (let r = minR; r <= maxR; r++) {
    const rowObj = SheetState.rows.find(ro => ro.row === r);
    const rowVals = [];
    for (let c = minC; c <= maxC; c++) {
      const colLetter = EXCEL_COLUMNS[c].col;
      rowVals.push(rowObj ? (rowObj[colLetter]?.val ?? '') : '');
    }
    lines.push(rowVals.join('\t'));
  }

  const tsv = lines.join('\n');
  navigator.clipboard.writeText(tsv).then(() => {
    showToast(`Copied ${maxR - minR + 1} x ${maxC - minC + 1} cells.`, 'info');
  }).catch(() => {});
}

function handleClipboardPaste(text) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('Permission Denied: Read-only access.', 'warning');
    return;
  }
  if (!text) return;

  const cur = SheetState.selected;
  if (!cur) return;

  const startC = EXCEL_COLUMNS.findIndex(c => c.col === cur.colLetter);
  if (startC === -1) return;

  const rows = text.split(/\r?\n/).filter(r => r.length > 0);
  let modifiedCount = 0;

  pushHistoryState();

  rows.forEach((rowStr, rOffset) => {
    const targetRow = cur.row + rOffset;
    if (isRowLocked(targetRow)) return;

    const rowObj = SheetState.rows.find(r => r.row === targetRow);
    if (!rowObj) return;

    const cells = rowStr.split('\t');
    cells.forEach((valStr, cOffset) => {
      const targetCIdx = startC + cOffset;
      if (targetCIdx >= EXCEL_COLUMNS.length) return;

      const colDef = EXCEL_COLUMNS[targetCIdx];
      if (colDef.isReadOnly || colDef.isFormula) return;

      let val = valStr.trim();
      if (colDef.col !== 'AM' && val !== '') {
        const num = parseFloat(val);
        if (!isNaN(num)) val = num;
      }

      rowObj[colDef.col] = { val };
      modifiedCount++;
    });

    recalculateRow(rowObj);
    updateSingleRowDisplay(targetRow);
  });

  recalculateTotalRow();
  updateTotalRowDisplay();
  saveSheetData(false);
  showToast(`Pasted into ${modifiedCount} cells successfully.`, 'success');
}

function deleteSelectedRange(isCut = false) {
  if (CurrentUser && CurrentUser.isReadOnly) {
    showToast('Permission Denied: Read-only access.', 'warning');
    return;
  }

  pushHistoryState();

  if (!SheetState.rangeSelection.start || !SheetState.rangeSelection.end) {
    const cur = SheetState.selected;
    if (cur && !isRowLocked(cur.row)) {
      const colDef = EXCEL_COLUMNS.find(c => c.col === cur.colLetter);
      if (colDef && !colDef.isReadOnly && !colDef.isFormula) {
        const rowObj = SheetState.rows.find(r => r.row === cur.row);
        if (rowObj) {
          rowObj[cur.colLetter] = { val: null };
          recalculateRow(rowObj);
          updateSingleRowDisplay(cur.row);
          recalculateTotalRow();
          updateTotalRowDisplay();
          saveSheetData(false);
        }
      }
    }
    return;
  }

  const startC = EXCEL_COLUMNS.findIndex(c => c.col === SheetState.rangeSelection.start.col);
  const endC = EXCEL_COLUMNS.findIndex(c => c.col === SheetState.rangeSelection.end.col);
  const minC = Math.min(startC, endC);
  const maxC = Math.max(startC, endC);

  const minR = Math.min(SheetState.rangeSelection.start.row, SheetState.rangeSelection.end.row);
  const maxR = Math.max(SheetState.rangeSelection.start.row, SheetState.rangeSelection.end.row);

  let count = 0;
  for (let r = minR; r <= maxR; r++) {
    if (isRowLocked(r)) continue;
    const rowObj = SheetState.rows.find(ro => ro.row === r);
    if (!rowObj) continue;

    for (let c = minC; c <= maxC; c++) {
      const colDef = EXCEL_COLUMNS[c];
      if (colDef.isReadOnly || colDef.isFormula) continue;
      rowObj[colDef.col] = { val: null };
      count++;
    }
    recalculateRow(rowObj);
    updateSingleRowDisplay(r);
  }

  recalculateTotalRow();
  updateTotalRowDisplay();
  saveSheetData(false);
  showToast(`Cleared ${count} cells.`, 'info');
}

// ──────────────────────────────────────────────────────────────────────────
// UNDO & REDO
// ──────────────────────────────────────────────────────────────────────────
function undoAction() {
  if (SheetState.undoStack.length === 0) {
    showToast('Nothing to undo.', 'info');
    return;
  }
  const snap = SheetState.undoStack.pop();
  SheetState.redoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(snap);
  renderExcelTable();
  recalculateTotalRow();
  updateTotalRowDisplay();
  saveSheetData(false);
  showToast('Action undone.', 'info');
}

function redoAction() {
  if (SheetState.redoStack.length === 0) {
    showToast('Nothing to redo.', 'info');
    return;
  }
  const snap = SheetState.redoStack.pop();
  SheetState.undoStack.push(JSON.stringify(SheetState.rows));
  SheetState.rows = JSON.parse(snap);
  renderExcelTable();
  recalculateTotalRow();
  updateTotalRowDisplay();
  saveSheetData(false);
  showToast('Action redone.', 'info');
}

// ──────────────────────────────────────────────────────────────────────────
// CORPORATE LUXURY EXECUTIVE-CLASS EXCEL WORKBOOK EXPORT ENGINE
// Multi-Tab Enterprise Suite with Frozen Panes, Dynamic KPI Tinting, Zebra Striping,
// Generous Padding, Zero-Clipping Headers, and Deep Executive Navy Banners
// ──────────────────────────────────────────────────────────────────────────
async function exportExcelFile() {
  try {
    showToast('Generating executive luxury Excel workbook...', 'info');

    const monthNames = MonthYearState.monthNames || ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[MonthYearState.monthIndex] || 'Month';
    const year = MonthYearState.year || 2026;
    const userRole = CurrentUser?.name ? CurrentUser.name.toUpperCase() : 'USER';
    const filename = `MEP_OEE_ENTERPRISE_REPORT_${userRole}_${monthName}_${year}.xlsx`;

    if (typeof ExcelJS === 'undefined') {
      showToast('ExcelJS library is loading. Please try again.', 'warning');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'MEP FAN LTD. - OEE SCADA MES';
    workbook.lastModifiedBy = CurrentUser?.name || 'Admin';
    workbook.created = new Date();
    workbook.modified = new Date();

    const allTabIds = Object.keys(SHEET_TABS);
    const allowedTabIds = allTabIds.filter(tId => isTabAllowedForUser(tId));

    for (const tabId of allowedTabIds) {
      if (tabId === 'summary_oee_yearly') {
        buildYearlySummaryExcelSheet(workbook, monthName, year);
      } else if (tabId === 'summary_oee') {
        buildOEESummaryExcelSheet(workbook, monthName, year);
      } else if (tabId === 'summary_production') {
        buildProductionOutputExcelSheet(workbook, monthName, year);
      } else if (tabId === 'summary_status') {
        buildRunningStatusExcelSheet(workbook, monthName, year);
      } else if (tabId === 'summary_downtime') {
        buildTotalDowntimeReportExcelSheet(workbook, monthName, year);
      } else {
        const tabRows = (tabId === ACTIVE_TAB && SheetState.rows) ? SheetState.rows : getTabData(tabId, year, MonthYearState.monthIndex);
        buildDepartmentExcelSheet(workbook, tabId, monthName, year, tabRows);
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    if (typeof saveAs !== 'undefined') {
      saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 300);
    }

    showToast(`Executive Excel exported successfully! (${allowedTabIds.length} tabs included)`, 'success');
  } catch (err) {
    console.error('Export Excel Error:', err);
    showToast('Excel export failed: ' + (err.message || 'Unknown error'), 'error');
  }
}

function getTabData(tabId, year, monthIndex) {
  const stored = LocalStorageEngine.load(tabId, year, monthIndex);
  if (stored && Array.isArray(stored.rows) && stored.rows.length > 0) {
    return stored.rows;
  }
  if (Array.isArray(stored) && stored.length > 0) {
    return stored;
  }
  return generateBlankMonthRows(tabId, year, monthIndex);
}

function openpyxlLetter(c) {
  let s = '';
  while (c > 0) {
    let m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - m) / 26);
  }
  return s;
}

// ──────────────────────────────────────────────────────────────────────────
// 1. DEPARTMENT EXCEL SHEET BUILDER (Enterprise Grade with Freeze Panes)
// ──────────────────────────────────────────────────────────────────────────
function buildDepartmentExcelSheet(workbook, tabId, monthName, year, rowsData) {
  const tabInfo = SHEET_TABS[tabId] || {};
  const sheetName = tabInfo.name || 'Department';
  const ws = workbook.addWorksheet(sheetName.substring(0, 31), {
    views: [{ state: 'frozen', xSplit: 10, ySplit: 5, topLeftCell: 'K6', showGridLines: true }]
  });

  const totalCols = EXCEL_COLUMNS.length; // 40 columns

  // ── ROW 1: Executive Corporate Banner (Deep Royal Navy #0F294D) ──
  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells(1, 1, 1, totalCols);
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // ── ROW 2: Report Subtitle (Steel Navy #1E3A8A) ──
  const r2 = ws.addRow(['Production Performance & OEE Analysis Report']);
  ws.mergeCells(2, 1, 2, totalCols);
  r2.height = 22;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // ── ROW 3: Department Title (Dark Slate #334155) ──
  const r3 = ws.addRow([`${sheetName}  •  Period: ${monthName} ${year}  •  SCADA MES Live Data`]);
  ws.mergeCells(3, 1, 3, totalCols);
  r3.height = 20;
  r3.getCell(1).font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  r3.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  r3.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // ── ROW 4: Downtime Code Numbers (Royal Purple #7C3AED with White text) ──
  const row4Vals = new Array(totalCols).fill('');
  EXCEL_COLUMNS.forEach((c, idx) => {
    if (c.code) row4Vals[idx] = c.code;
  });
  const r4 = ws.addRow(row4Vals);
  r4.height = 20;
  r4.eachCell((cell, colIdx) => {
    const colDef = EXCEL_COLUMNS[colIdx - 1];
    if (colDef?.isDt || colDef?.code) {
      cell.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } }; // Royal Purple
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF4C1D95' } },
        left: { style: 'thin', color: { argb: 'FF4C1D95' } },
        bottom: { style: 'thin', color: { argb: 'FF4C1D95' } },
        right: { style: 'thin', color: { argb: 'FF4C1D95' } }
      };
    } else {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    }
  });

  // ── ROW 5: Column Headers (Tall 115pt Height, Distinct Corporate Fills & Text Wrap) ──
  const colHeaders = EXCEL_COLUMNS.map(c => c.label);
  const r5 = ws.addRow(colHeaders);
  r5.height = 115;
  r5.eachCell((cell, colIdx) => {
    const colDef = EXCEL_COLUMNS[colIdx - 1];
    let bgArgb = 'FFD9E1F2'; // Royal Sky (Metadata)
    let txtArgb = 'FF0F172A';

    if (colDef?.col === 'E' || colDef?.col === 'F') {
      bgArgb = 'FFE2EFDA'; // Soft Mint (Capacity & Actual)
      txtArgb = 'FF14532D';
    } else if (colDef?.col === 'G') {
      bgArgb = 'FFFCE4D6'; // Soft Rose (Rejection)
      txtArgb = 'FF7F1D1D';
    } else if (colDef?.col === 'H' || colDef?.col === 'I' || colDef?.col === 'J') {
      bgArgb = 'FFDDEBF7'; // Soft Cyan (Planned & Run time)
      txtArgb = 'FF0C4A6E';
    } else if (colDef?.isDt) {
      bgArgb = 'FFEDE9FE'; // Soft Lilac (Downtimes)
      txtArgb = 'FF4C1D95';
    } else if (colDef?.isLossDt || colDef?.isTotalDt) {
      bgArgb = 'FFFEF3C7'; // Soft Amber (Loss & Total DT)
      txtArgb = 'FF78350F';
    } else if (colDef?.col === 'AI') {
      bgArgb = 'FFCFFAFE'; // Cyan (Availability)
      txtArgb = 'FF155E75';
    } else if (colDef?.col === 'AJ') {
      bgArgb = 'FFDBEAFE'; // Blue (Performance)
      txtArgb = 'FF1E40AF';
    } else if (colDef?.col === 'AK') {
      bgArgb = 'FFD1FAE5'; // Emerald (Quality)
      txtArgb = 'FF065F46';
    } else if (colDef?.col === 'AL') {
      bgArgb = 'FFFEF08A'; // Golden Trophy (OEE %)
      txtArgb = 'FF713F12';
    } else if (colDef?.isRemarks) {
      bgArgb = 'FFF1F5F9';
      txtArgb = 'FF334155';
    }

    const fSize = colDef?.isRemarks ? 12 : (colDef?.col === 'AL' ? 10 : 8.5);
    cell.font = { name: 'Times New Roman', size: fSize, bold: true, color: { argb: txtArgb } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF334155' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } }
    };
  });

  // ── DATA ROWS (Row 6 onwards with Zebra Striping & KPI Tinting) ──
  const rowsPerDay = getRowsPerDay(tabId);
  const startRowIdx = 6;
  let currentExcelRow = startRowIdx;
  let rowsToIterate = [];
  if (Array.isArray(rowsData)) {
    rowsToIterate = rowsData;
  } else if (rowsData && Array.isArray(rowsData.rows)) {
    rowsToIterate = rowsData.rows;
  } else {
    rowsToIterate = generateBlankMonthRows(tabId, year, MonthYearState.monthIndex);
  }

  rowsToIterate.forEach((rowObj, rIdx) => {
    const r = rowObj.row;
    const mIdx = (r - 6) % rowsPerDay;
    const dayIdx = Math.floor(rIdx / rowsPerDay);
    const isEvenDay = (dayIdx % 2 === 0);
    const zebraBg = isEvenDay ? 'FFFFFFFF' : 'FFF8FAFC'; // Ultra-clean alternating day tint

    const rowVals = EXCEL_COLUMNS.map(c => {
      const val = rowObj[c.col]?.val;
      if (val === null || val === undefined || val === '') return '';
      if (c.col === 'A' || c.col === 'B' || c.col === 'C' || c.col === 'D' || c.col === 'AM') return val;
      if (c.isPercent && typeof val === 'number') return val;
      if ((c.isNumeric || c.isDt || c.isFormula) && typeof val === 'number') return Math.round(val);
      return val;
    });

    const dataRow = ws.addRow(rowVals);
    dataRow.height = 18;

    dataRow.eachCell((cell, colIdx) => {
      const colDef = EXCEL_COLUMNS[colIdx - 1];
      const val = cell.value;
      const isRejCol = (colDef?.col === 'G');
      const isDtCol = colDef?.isDt;
      const isOEECol = (colDef?.col === 'AL');

      let cellBg = zebraBg;
      let cellTxt = 'FF0F172A';
      let isBold = colDef?.isFormula || colDef?.isPercent || false;

      // Rejection Red Highlight
      if (isRejCol && typeof val === 'number' && val > 0) {
        cellBg = 'FFFEE2E2'; // Soft Light Red
        cellTxt = 'FF991B1B'; // Deep Crimson
        isBold = true;
      }
      // Downtime Purple Highlight
      if (isDtCol && typeof val === 'number' && val > 0) {
        cellBg = 'FFEDE9FE'; // Soft Lilac
        cellTxt = 'FF5B21B6'; // Deep Violet
        isBold = true;
      }
      // OEE Golden / Green / Amber Tint
      if (isOEECol && typeof val === 'number') {
        if (val >= 0.85) {
          cellBg = 'FFDCFCE7'; // Soft Emerald
          cellTxt = 'FF166534';
        } else if (val > 0 && val < 0.50) {
          cellBg = 'FFFEF3C7'; // Soft Amber
          cellTxt = 'FF92400E';
        }
        isBold = true;
      }

      cell.font = { name: 'Times New Roman', size: 10, bold: isBold, color: { argb: cellTxt } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cellBg } };
      cell.alignment = {
        horizontal: colDef?.align || 'center',
        vertical: 'middle',
        wrapText: true
      };

      if (colDef?.isPercent) {
        cell.numFmt = '0%';
      } else if (colDef?.isNumeric || colDef?.isDt || colDef?.isFormula) {
        if (typeof val === 'number') cell.numFmt = '#,##0';
      }

      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });

    currentExcelRow++;
  });

  // Apply Merges & Dark Navy Day Separators
  const totalDays = Math.floor(rowsToIterate.length / rowsPerDay);
  for (let d = 0; d < totalDays; d++) {
    const dayStartExcelRow = startRowIdx + (d * rowsPerDay);
    const dayEndExcelRow = dayStartExcelRow + rowsPerDay - 1;

    // Merge Cols A, B, C (Date, Day, Shift)
    if (rowsPerDay > 1) {
      ws.mergeCells(dayStartExcelRow, 1, dayEndExcelRow, 1);
      ws.mergeCells(dayStartExcelRow, 2, dayEndExcelRow, 2);
      ws.mergeCells(dayStartExcelRow, 3, dayEndExcelRow, 3);
    }

    // Merge Time Groups (e.g. APC 3 machines in Auto Powder Coating)
    if (tabInfo.timeGroups) {
      tabInfo.timeGroups.forEach(g => {
        if (g.count > 1) {
          const gStart = dayStartExcelRow + g.startIdx;
          const gEnd = gStart + g.count - 1;
          EXCEL_COLUMNS.forEach((colDef, cIdx) => {
            if (TIME_COLUMNS.includes(colDef.col)) {
              const colNum = cIdx + 1;
              try {
                ws.mergeCells(gStart, colNum, gEnd, colNum);
              } catch (e) {}
            }
          });
        }
      });
    }

    // Day End Thick Corporate Navy Border
    for (let c = 1; c <= totalCols; c++) {
      const cell = ws.getCell(dayEndExcelRow, c);
      const b = cell.border || {};
      cell.border = {
        ...b,
        bottom: { style: 'medium', color: { argb: 'FF0F294D' } }
      };
    }
  }

  // ── TOTAL ROW AT BOTTOM (Deep Corporate Navy #0F294D with Crisp White Text) ──
  const lastDataRow = currentExcelRow - 1;
  const totalExcelRow = currentExcelRow;
  const totRowVals = new Array(totalCols).fill('');
  totRowVals[0] = 'Total:';

  EXCEL_COLUMNS.forEach((colDef, idx) => {
    const colNum = idx + 1;
    const colLet = colDef.col;
    if (['A', 'B', 'C', 'D'].includes(colLet)) return;

    if (colDef.isPercent) {
      if (colLet === 'AI') totRowVals[idx] = { formula: `IFERROR(J${totalExcelRow}/H${totalExcelRow}, 0)` };
      else if (colLet === 'AJ') totRowVals[idx] = { formula: `IFERROR(F${totalExcelRow}/E${totalExcelRow}, 0)` };
      else if (colLet === 'AK') totRowVals[idx] = { formula: `IFERROR(F${totalExcelRow}/(F${totalExcelRow}+G${totalExcelRow}), 0)` };
      else if (colLet === 'AL') totRowVals[idx] = { formula: `IFERROR(AK${totalExcelRow}*AJ${totalExcelRow}*AI${totalExcelRow}, 0)` };
    } else if (colDef.isNumeric || colDef.isDt || colDef.isFormula) {
      totRowVals[idx] = { formula: `SUM(${openpyxlLetter(colNum)}6:${openpyxlLetter(colNum)}${lastDataRow})` };
    }
  });

  const totRow = ws.addRow(totRowVals);
  totRow.height = 26;
  ws.mergeCells(totalExcelRow, 1, totalExcelRow, 4);

  totRow.eachCell((cell, colIdx) => {
    const colDef = EXCEL_COLUMNS[colIdx - 1];
    cell.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };

    if (colDef?.isPercent) {
      cell.numFmt = '0%';
    } else if (colDef?.isNumeric || colDef?.isDt || colDef?.isFormula) {
      cell.numFmt = '#,##0';
    }

    cell.border = {
      top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      left: { style: 'thin', color: { argb: 'FF334155' } },
      bottom: { style: 'double', color: { argb: 'FFFFFFFF' } },
      right: { style: 'thin', color: { argb: 'FF334155' } }
    };
  });

  // Generous column widths (Zero Clipping)
  const colWidthMap = {
    A: 11.0, B: 7.0, C: 10.0, D: 32.0, E: 13.0, F: 13.0, G: 9.0,
    H: 11.0, I: 10.0, J: 11.0, AH: 12.0, AI: 11.0, AJ: 11.0, AK: 11.0, AL: 12.0, AM: 38.0
  };
  EXCEL_COLUMNS.forEach((c, idx) => {
    const w = colWidthMap[c.col] || (c.isDt ? 8.0 : 10.0);
    ws.getColumn(idx + 1).width = w;
  });
}

// ──────────────────────────────────────────────────────────────────────────
// 2. PRODUCTION OUTPUT SUMMARY SHEET (Enterprise Executive Edition)
// ──────────────────────────────────────────────────────────────────────────
function buildProductionOutputExcelSheet(workbook, monthName, year) {
  const ws = workbook.addWorksheet('Production Status', {
    views: [{ state: 'frozen', ySplit: 4, showGridLines: true }]
  });

  // Title Row 1
  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells('A1:G1');
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // Subtitle Row 2
  const r2 = ws.addRow([`Production Output & Efficiency Analysis (${monthName} ${year})`]);
  ws.mergeCells('A2:G2');
  r2.height = 22;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // Header Rows 3 & 4 (2-Level Wrapped Header with Generous Widths)
  const r3 = ws.addRow(['Section Name', 'Production Running (hr)', 'Machine Capacity (Pcs)', 'Production Qty.', '', 'Standard Wise Production Output (%)', 'Remarks']);
  const r4 = ws.addRow(['', '', '', 'Total Production (pcs)', 'Rejection (Pcs)', '', '']);
  r3.height = 26;
  r4.height = 30;

  ws.mergeCells('A3:A4');
  ws.mergeCells('B3:B4');
  ws.mergeCells('C3:C4');
  ws.mergeCells('D3:E3');
  ws.mergeCells('F3:F4');
  ws.mergeCells('G3:G4');

  [r3, r4].forEach(r => {
    r.eachCell((cell, colIdx) => {
      let bgArgb = 'FFD9E1F2';
      let txtArgb = 'FF0F172A';
      if (colIdx === 4) { bgArgb = 'FFE2EFDA'; txtArgb = 'FF14532D'; }
      else if (colIdx === 5) { bgArgb = 'FFFCE4D6'; txtArgb = 'FF7F1D1D'; }
      else if (colIdx === 6) { bgArgb = 'FFDDEBF7'; txtArgb = 'FF0C4A6E'; }

      cell.font = { name: 'Times New Roman', size: 10.5, bold: true, color: { argb: txtArgb } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF334155' } },
        left: { style: 'thin', color: { argb: 'FF94A3B8' } },
        bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
        right: { style: 'thin', color: { argb: 'FF94A3B8' } }
      };
    });
  });

  const depts = getActiveSummaryDepts();
  let startDataRow = 5;
  let curRow = startDataRow;

  depts.forEach((dept, idx) => {
    const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);
    const runHours = Math.round(oeeData.runTimeMins / 60);
    const zebraBg = (idx % 2 === 0) ? 'FFFFFFFF' : 'FFF8FAFC';

    const row = ws.addRow([dept.name, runHours, oeeData.capacityPcs, oeeData.totalProduction, oeeData.rejectionPcs, oeeData.performance, oeeData.remarks || '']);
    row.height = 22;

    row.eachCell((cell, colIdx) => {
      let cellBg = zebraBg;
      let cellTxt = 'FF0F172A';
      let isBold = (colIdx === 6);

      if (colIdx === 5 && typeof cell.value === 'number' && cell.value > 0) {
        cellBg = 'FFFEE2E2';
        cellTxt = 'FF991B1B';
        isBold = true;
      }

      cell.font = { name: 'Times New Roman', size: 10.5, bold: isBold, color: { argb: cellTxt } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cellBg } };
      cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle', wrapText: true };
      if (colIdx === 6) cell.numFmt = '0%';
      else if ([2, 3, 4, 5].includes(colIdx)) cell.numFmt = '#,##0';

      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });
    curRow++;
  });

  // Total Row (Deep Navy #0F294D)
  const lastDataRow = curRow - 1;
  const totRowVals = [
    'Total:',
    { formula: `SUM(B${startDataRow}:B${lastDataRow})` },
    { formula: `SUM(C${startDataRow}:C${lastDataRow})` },
    { formula: `SUM(D${startDataRow}:D${lastDataRow})` },
    { formula: `SUM(E${startDataRow}:E${lastDataRow})` },
    { formula: `IFERROR(D${curRow}/C${curRow}, 0)` },
    ''
  ];
  const totRow = ws.addRow(totRowVals);
  totRow.height = 26;

  totRow.eachCell((cell, colIdx) => {
    cell.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
    cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle' };
    if (colIdx === 6) cell.numFmt = '0%';
    else if ([2, 3, 4, 5].includes(colIdx)) cell.numFmt = '#,##0';
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      left: { style: 'thin', color: { argb: 'FF334155' } },
      bottom: { style: 'double', color: { argb: 'FFFFFFFF' } },
      right: { style: 'thin', color: { argb: 'FF334155' } }
    };
  });

  // Generous column widths (Zero Clipping)
  ws.columns = [
    { width: 32.0 }, // Section Name
    { width: 24.0 }, // Production Running (hr)
    { width: 24.0 }, // Machine Capacity (Pcs)
    { width: 24.0 }, // Total Production (pcs)
    { width: 18.0 }, // Rejection (Pcs)
    { width: 34.0 }, // Standard Wise Production Output (%)
    { width: 28.0 }  // Remarks
  ];
}

// ──────────────────────────────────────────────────────────────────────────
// 3. RUNNING & DOWNTIME STATUS SHEET (Enterprise Executive Edition)
// ──────────────────────────────────────────────────────────────────────────
function buildRunningStatusExcelSheet(workbook, monthName, year) {
  const ws = workbook.addWorksheet('Down Time & Running Status', {
    views: [{ state: 'frozen', ySplit: 3, showGridLines: true }]
  });

  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells('A1:G1');
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r2 = ws.addRow([`Down Time & Running Time Status (${monthName} ${year})`]);
  ws.mergeCells('A2:G2');
  r2.height = 22;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r3 = ws.addRow([
    'Machine No.',
    'Planned Production Time (mins)',
    'Production Run Time (Mins)',
    'Machine Down Time (Mins)',
    'Production Running Time (%)',
    'Production Down Time (%)',
    'Remarks'
  ]);
  r3.height = 48;

  r3.eachCell((cell, colIdx) => {
    let bgArgb = 'FFD9E1F2';
    let txtArgb = 'FF0F172A';
    if (colIdx === 3) { bgArgb = 'FFE2EFDA'; txtArgb = 'FF14532D'; }
    else if (colIdx === 4) { bgArgb = 'FFFCE4D6'; txtArgb = 'FF7F1D1D'; }
    else if (colIdx === 5) { bgArgb = 'FFDDEBF7'; txtArgb = 'FF0C4A6E'; }
    else if (colIdx === 6) { bgArgb = 'FFFEF3C7'; txtArgb = 'FF78350F'; }

    cell.font = { name: 'Times New Roman', size: 10.5, bold: true, color: { argb: txtArgb } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF334155' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } }
    };
  });

  const depts = getActiveSummaryDepts();
  let startDataRow = 4;
  let curRow = startDataRow;

  depts.forEach((dept, idx) => {
    const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);
    const runPct = oeeData.plannedTimeMins > 0 ? (oeeData.runTimeMins / oeeData.plannedTimeMins) : 0;
    const dtPct = oeeData.plannedTimeMins > 0 ? (oeeData.downTimeMins / oeeData.plannedTimeMins) : 0;
    const zebraBg = (idx % 2 === 0) ? 'FFFFFFFF' : 'FFF8FAFC';

    const row = ws.addRow([dept.name, oeeData.plannedTimeMins, oeeData.runTimeMins, oeeData.downTimeMins, runPct, dtPct, oeeData.remarks || '']);
    row.height = 22;

    row.eachCell((cell, colIdx) => {
      let isBold = [5, 6].includes(colIdx);
      cell.font = { name: 'Times New Roman', size: 10.5, bold: isBold, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: zebraBg } };
      cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle', wrapText: true };
      if ([5, 6].includes(colIdx)) cell.numFmt = '0%';
      else if ([2, 3, 4].includes(colIdx)) cell.numFmt = '#,##0';

      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });
    curRow++;
  });

  // Total Row
  const lastDataRow = curRow - 1;
  const totRowVals = [
    'Total:',
    { formula: `SUM(B${startDataRow}:B${lastDataRow})` },
    { formula: `SUM(C${startDataRow}:C${lastDataRow})` },
    { formula: `SUM(D${startDataRow}:D${lastDataRow})` },
    { formula: `IFERROR(C${curRow}/B${curRow}, 0)` },
    { formula: `IFERROR(D${curRow}/B${curRow}, 0)` },
    ''
  ];
  const totRow = ws.addRow(totRowVals);
  totRow.height = 26;

  totRow.eachCell((cell, colIdx) => {
    cell.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
    cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle' };
    if ([5, 6].includes(colIdx)) cell.numFmt = '0%';
    else if ([2, 3, 4].includes(colIdx)) cell.numFmt = '#,##0';
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      left: { style: 'thin', color: { argb: 'FF334155' } },
      bottom: { style: 'double', color: { argb: 'FFFFFFFF' } },
      right: { style: 'thin', color: { argb: 'FF334155' } }
    };
  });

  ws.columns = [
    { width: 32.0 }, // Machine No.
    { width: 30.0 }, // Planned Production Time (mins)
    { width: 28.0 }, // Production Run Time (Mins)
    { width: 28.0 }, // Machine Down Time (Mins)
    { width: 28.0 }, // Production Running Time (%)
    { width: 26.0 }, // Production Down Time (%)
    { width: 28.0 }  // Remarks
  ];
}

// ──────────────────────────────────────────────────────────────────────────
// 4. TOTAL DOWNTIME MATRIX REPORT (Enterprise Executive Edition)
// ──────────────────────────────────────────────────────────────────────────
function buildTotalDowntimeReportExcelSheet(workbook, monthName, year) {
  const ws = workbook.addWorksheet('Total Downtime Report', {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 4, showGridLines: true }]
  });
  const dtCols = EXCEL_COLUMNS.filter(c => c.isDt);
  const totalCols = 1 + dtCols.length + 4;

  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells(1, 1, 1, totalCols);
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r2 = ws.addRow([`Total Downtime Comprehensive Report (${monthName} ${year})`]);
  ws.mergeCells(2, 1, 2, totalCols);
  r2.height = 22;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 3: Codes (10 to 35) in Purple
  const codeVals = new Array(totalCols).fill('');
  dtCols.forEach((c, idx) => {
    codeVals[idx + 1] = c.code || '';
  });
  const r3 = ws.addRow(codeVals);
  r3.height = 20;
  r3.eachCell((cell, colIdx) => {
    if (colIdx > 1 && colIdx <= dtCols.length + 1) {
      cell.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    }
  });

  // Row 4: Column Labels
  const labelVals = ['Machine No.', ...dtCols.map(c => c.label), 'Total Down Time (mins)', 'Planned Production Time (mins)', 'Total Production Run Time (mins)', 'Total Down Time (%)'];
  const r4 = ws.addRow(labelVals);
  r4.height = 110;
  r4.eachCell((cell, colIdx) => {
    let bgArgb = 'FFD9E1F2';
    let txtArgb = 'FF0F172A';
    if (colIdx > 1 && colIdx <= dtCols.length + 1) { bgArgb = 'FFEDE9FE'; txtArgb = 'FF4C1D95'; }
    else if (colIdx === dtCols.length + 2) { bgArgb = 'FFFEF3C7'; txtArgb = 'FF78350F'; }
    else if (colIdx === dtCols.length + 4) { bgArgb = 'FFE2EFDA'; txtArgb = 'FF14532D'; }
    else if (colIdx === dtCols.length + 5) { bgArgb = 'FFCFFAFE'; txtArgb = 'FF155E75'; }

    cell.font = { name: 'Times New Roman', size: 8.5, bold: true, color: { argb: txtArgb } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
  });

  const depts = getActiveSummaryDepts();
  let startDataRow = 5;
  let curRow = startDataRow;

  depts.forEach((dept, idx) => {
    const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);
    const dtSumMap = {};
    dtCols.forEach(c => { dtSumMap[c.col] = 0; });

    const rows = getTabData(dept.id, year, MonthYearState.monthIndex);
    (rows || []).forEach(r => {
      dtCols.forEach(c => {
        const val = parseFloat(r[c.col]?.val) || 0;
        dtSumMap[c.col] += val;
      });
    });

    const dtVals = dtCols.map(c => dtSumMap[c.col]);
    const dtPct = oeeData.plannedTimeMins > 0 ? (oeeData.downTimeMins / oeeData.plannedTimeMins) : 0;
    const rowVals = [dept.name, ...dtVals, oeeData.downTimeMins, oeeData.plannedTimeMins, oeeData.runTimeMins, dtPct];
    const row = ws.addRow(rowVals);
    row.height = 20;

    const zebraBg = (idx % 2 === 0) ? 'FFFFFFFF' : 'FFF8FAFC';

    row.eachCell((cell, colIdx) => {
      let cellBg = zebraBg;
      let cellTxt = 'FF0F172A';
      let isBold = (colIdx > dtCols.length + 1);

      if (colIdx > 1 && colIdx <= dtCols.length + 1 && typeof cell.value === 'number' && cell.value > 0) {
        cellBg = 'FFEDE9FE';
        cellTxt = 'FF5B21B6';
        isBold = true;
      }

      cell.font = { name: 'Times New Roman', size: 10, bold: isBold, color: { argb: cellTxt } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cellBg } };
      cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle' };
      if (colIdx === totalCols) cell.numFmt = '0%';
      else if (colIdx > 1) cell.numFmt = '#,##0';

      cell.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, left: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }, right: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
    });
    curRow++;
  });

  // Total Row (Deep Navy #0F294D)
  const lastDataRow = curRow - 1;
  const totVals = new Array(totalCols).fill('');
  totVals[0] = 'Total:';
  for (let c = 2; c < totalCols; c++) {
    const colLet = openpyxlLetter(c);
    totVals[c - 1] = { formula: `SUM(${colLet}${startDataRow}:${colLet}${lastDataRow})` };
  }
  const totDtLet = openpyxlLetter(dtCols.length + 2);
  const planLet = openpyxlLetter(dtCols.length + 3);
  totVals[totalCols - 1] = { formula: `IFERROR(${totDtLet}${curRow}/${planLet}${curRow}, 0)` };

  const totRow = ws.addRow(totVals);
  totRow.height = 26;
  totRow.eachCell((cell, colIdx) => {
    cell.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
    cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle' };
    if (colIdx === totalCols) cell.numFmt = '0%';
    else if (colIdx > 1) cell.numFmt = '#,##0';
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'double', color: { argb: 'FFFFFFFF' } }, right: { style: 'thin' } };
  });

  const colWidths = [{ width: 30.0 }];
  dtCols.forEach(() => colWidths.push({ width: 13.0 }));
  colWidths.push({ width: 20.0 }, { width: 26.0 }, { width: 26.0 }, { width: 20.0 });
  ws.columns = colWidths;
}

// ──────────────────────────────────────────────────────────────────────────
// 5. OEE DEPARTMENT SUMMARY SHEET (Enterprise Executive Edition)
// ──────────────────────────────────────────────────────────────────────────
function buildOEESummaryExcelSheet(workbook, monthName, year) {
  const ws = workbook.addWorksheet('OEE', {
    views: [{ state: 'frozen', ySplit: 4, showGridLines: true }]
  });

  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells('A1:I1');
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r2 = ws.addRow([`Overall Equipment Effectiveness (OEE) Report (${monthName} ${year})`]);
  ws.mergeCells('A2:I2');
  r2.height = 22;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r3 = ws.addRow(['Section', 'Machine Capacity (Pcs)', 'Production Qty.', '', 'Availability (%)', 'Performance (%)', 'Quality (%)', 'OEE (%)', "Remark's"]);
  const r4 = ws.addRow(['', '', 'Total Production (pcs)', 'Rejection (Pcs)', '', '', '', '', '']);
  r3.height = 24;
  r4.height = 30;

  ws.mergeCells('A3:A4');
  ws.mergeCells('B3:B4');
  ws.mergeCells('C3:D3');
  ws.mergeCells('E3:E4');
  ws.mergeCells('F3:F4');
  ws.mergeCells('G3:G4');
  ws.mergeCells('H3:H4');
  ws.mergeCells('I3:I4');

  [r3, r4].forEach(r => {
    r.eachCell((cell, colIdx) => {
      let bgArgb = 'FFD9E1F2';
      let txtArgb = 'FF0F172A';
      if (colIdx === 3) { bgArgb = 'FFE2EFDA'; txtArgb = 'FF14532D'; }
      else if (colIdx === 4) { bgArgb = 'FFFCE4D6'; txtArgb = 'FF7F1D1D'; }
      else if (colIdx === 5) { bgArgb = 'FFCFFAFE'; txtArgb = 'FF155E75'; }
      else if (colIdx === 6) { bgArgb = 'FFDBEAFE'; txtArgb = 'FF1E40AF'; }
      else if (colIdx === 7) { bgArgb = 'FFD1FAE5'; txtArgb = 'FF065F46'; }
      else if (colIdx === 8) { bgArgb = 'FFFEF08A'; txtArgb = 'FF713F12'; }

      cell.font = { name: 'Times New Roman', size: 10.5, bold: true, color: { argb: txtArgb } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { top: { style: 'medium' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
    });
  });

  const depts = getActiveSummaryDepts();
  let startDataRow = 5;
  let curRow = startDataRow;

  depts.forEach((dept, idx) => {
    const oeeData = getTabOEESummary(dept.id, year, MonthYearState.monthIndex);
    const zebraBg = (idx % 2 === 0) ? 'FFFFFFFF' : 'FFF8FAFC';

    const row = ws.addRow([dept.name, oeeData.capacityPcs, oeeData.totalProduction, oeeData.rejectionPcs, oeeData.availability, oeeData.performance, oeeData.quality, oeeData.oee, oeeData.remarks || '']);
    row.height = 22;

    row.eachCell((cell, colIdx) => {
      let cellBg = zebraBg;
      let cellTxt = 'FF0F172A';
      let isBold = [5, 6, 7, 8].includes(colIdx);

      if (colIdx === 4 && typeof cell.value === 'number' && cell.value > 0) {
        cellBg = 'FFFEE2E2';
        cellTxt = 'FF991B1B';
      }
      if (colIdx === 8 && typeof cell.value === 'number') {
        if (cell.value >= 0.85) { cellBg = 'FFDCFCE7'; cellTxt = 'FF166534'; }
        else if (cell.value > 0 && cell.value < 0.50) { cellBg = 'FFFEF3C7'; cellTxt = 'FF92400E'; }
      }

      cell.font = { name: 'Times New Roman', size: 10.5, bold: isBold, color: { argb: cellTxt } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cellBg } };
      cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle', wrapText: true };
      if ([5, 6, 7, 8].includes(colIdx)) cell.numFmt = '0%';
      else if ([2, 3, 4].includes(colIdx)) cell.numFmt = '#,##0';

      cell.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, left: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }, right: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
    });
    curRow++;
  });

  // Total Row
  const lastDataRow = curRow - 1;
  const totRowVals = [
    'Total:',
    { formula: `SUM(B${startDataRow}:B${lastDataRow})` },
    { formula: `SUM(C${startDataRow}:C${lastDataRow})` },
    { formula: `SUM(D${startDataRow}:D${lastDataRow})` },
    { formula: `AVERAGE(E${startDataRow}:E${lastDataRow})` },
    { formula: `IFERROR(C${curRow}/B${curRow}, 0)` },
    { formula: `IFERROR(C${curRow}/(C${curRow}+D${curRow}), 0)` },
    { formula: `IFERROR(E${curRow}*F${curRow}*G${curRow}, 0)` },
    ''
  ];
  const totRow = ws.addRow(totRowVals);
  totRow.height = 26;

  totRow.eachCell((cell, colIdx) => {
    cell.font = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
    cell.alignment = { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'middle' };
    if ([5, 6, 7, 8].includes(colIdx)) cell.numFmt = '0%';
    else if ([2, 3, 4].includes(colIdx)) cell.numFmt = '#,##0';
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'double', color: { argb: 'FFFFFFFF' } }, right: { style: 'thin' } };
  });

  ws.columns = [
    { width: 32.0 }, // Section
    { width: 24.0 }, // Machine Capacity
    { width: 24.0 }, // Total Production
    { width: 18.0 }, // Rejection
    { width: 18.0 }, // Availability
    { width: 18.0 }, // Performance
    { width: 16.0 }, // Quality
    { width: 18.0 }, // OEE
    { width: 28.0 }  // Remarks
  ];
}

// ──────────────────────────────────────────────────────────────────────────
// 6. YEARLY SUMMARY OF OEE (Enterprise Executive Edition)
// ──────────────────────────────────────────────────────────────────────────
function buildYearlySummaryExcelSheet(workbook, monthName, year) {
  const ws = workbook.addWorksheet('Summary of OEE', {
    views: [{ state: 'frozen', ySplit: 5, showGridLines: true }]
  });

  const r1 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells('A1:J1');
  r1.height = 36;
  r1.getCell(1).font = { name: 'Times New Roman', size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  r1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F294D' } };
  r1.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r2 = ws.addRow(['MEP FAN LTD.']);
  ws.mergeCells('A2:J2');
  r2.height = 20;
  r2.getCell(1).font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FFDBEAFE' } };
  r2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  r2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r3 = ws.addRow([`Annual Executive Summary of OEE (${year})`]);
  ws.mergeCells('A3:J3');
  r3.height = 20;
  r3.getCell(1).font = { name: 'Times New Roman', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  r3.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  r3.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  const r4 = ws.addRow(['Month', 'Detail', 'Machine Capacity (Pcs)', 'Production Qty.', '', 'Availability (%)', 'Performance (%)', 'Quality (%)', 'OEE (%)', "Remark's"]);
  const r5 = ws.addRow(['', '', '', 'Total Production (pcs)', 'Rejection (Pcs)', '', '', '', '', '']);
  r4.height = 22;
  r5.height = 38;

  ws.mergeCells('A4:A5');
  ws.mergeCells('B4:B5');
  ws.mergeCells('C4:C5');
  ws.mergeCells('D4:E4');
  ws.mergeCells('F4:F5');
  ws.mergeCells('G4:G5');
  ws.mergeCells('H4:H5');
  ws.mergeCells('I4:I5');
  ws.mergeCells('J4:J5');

  [r4, r5].forEach(r => {
    r.eachCell((cell, colIdx) => {
      let bgArgb = 'FFD9E1F2';
      let txtArgb = 'FF0F172A';
      if (colIdx === 4) { bgArgb = 'FFE2EFDA'; txtArgb = 'FF14532D'; }
      else if (colIdx === 5) { bgArgb = 'FFFCE4D6'; txtArgb = 'FF7F1D1D'; }
      else if (colIdx === 6) { bgArgb = 'FFCFFAFE'; txtArgb = 'FF155E75'; }
      else if (colIdx === 7) { bgArgb = 'FFDBEAFE'; txtArgb = 'FF1E40AF'; }
      else if (colIdx === 8) { bgArgb = 'FFD1FAE5'; txtArgb = 'FF065F46'; }
      else if (colIdx === 9) { bgArgb = 'FFFEF08A'; txtArgb = 'FF713F12'; }

      cell.font = { name: 'Times New Roman', size: 10.5, bold: true, color: { argb: txtArgb } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { top: { style: 'medium' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
    });
  });

  const yearlyData = getYearlyOEESummary(year);
  let curRow = 6;
  yearlyData.forEach(m => {
    const rowA = ws.addRow([m.monthName, 'Total', m.capacityPcs || 0, m.totalProduction || 0, m.rejectionPcs || 0, m.availability, m.performance, m.quality, m.oee, m.remarks || '']);
    const rowB = ws.addRow(['', 'Total Acheivement  (%)', m.capacityPcs || 0, '', '', m.availability, m.performance, m.quality, m.oee, '']);
    rowA.height = 20;
    rowB.height = 20;

    ws.mergeCells(curRow, 1, curRow + 1, 1);

    rowA.eachCell((cell, colIdx) => {
      cell.font = { name: 'Times New Roman', size: 10.5 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } }; // Soft Lilac
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if ([6, 7, 8, 9].includes(colIdx)) cell.numFmt = '0%';
      else if ([3, 4, 5].includes(colIdx)) cell.numFmt = '#,##0';
      cell.border = { top: { style: 'thin', color: { argb: 'FFCBD5E1' } }, left: { style: 'thin', color: { argb: 'FFCBD5E1' } }, bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } }, right: { style: 'thin', color: { argb: 'FFCBD5E1' } } };
    });

    rowB.eachCell((cell, colIdx) => {
      cell.font = { name: 'Times New Roman', size: 10.5, bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }; // Sky Blue
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if ([6, 7, 8, 9].includes(colIdx)) cell.numFmt = '0%';
      else if ([3, 4, 5].includes(colIdx)) cell.numFmt = '#,##0';
      cell.border = { top: { style: 'thin', color: { argb: 'FFCBD5E1' } }, left: { style: 'thin', color: { argb: 'FFCBD5E1' } }, bottom: { style: 'medium', color: { argb: 'FF0F294D' } }, right: { style: 'thin', color: { argb: 'FFCBD5E1' } } };
    });

    curRow += 2;
  });

  ws.columns = [
    { width: 18.0 }, // Month
    { width: 28.0 }, // Detail
    { width: 24.0 }, // Capacity
    { width: 24.0 }, // Total Production
    { width: 18.0 }, // Rejection
    { width: 18.0 }, // Availability
    { width: 18.0 }, // Performance
    { width: 18.0 }, // Quality
    { width: 18.0 }, // OEE
    { width: 28.0 }  // Remarks
  ];
}

function exportPDFReport() {
  window.print();
}

// ──────────────────────────────────────────────────────────────────────────
// KEYBOARD SHORTCUTS & CONTEXT MENU
// ──────────────────────────────────────────────────────────────────────────
function initKeyboardShortcuts() {
  const formulaInput = document.getElementById('formulaInput');

  window.addEventListener('keydown', (e) => {
    // If currently editing a cell or focused in an input/textarea, let the input handle all keys!
    if (SheetState.isEditing || SheetState.activeInput || document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA') {
      return;
    }
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
      SheetState.rangeSelection.end = { col: 'AP', row: maxActive };
      selectCell('E', 6, false);
      highlightSelectedRange();
      return;
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('Permission Denied: Read-only access.', 'warning');
        return;
      }
      saveSheetData(true);
      showToast('Changes saved to cloud successfully!', 'success');
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
        showToast('Permission Denied: Read-only access.', 'warning');
        return;
      }
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText().then(text => {
          if (text) handleClipboardPaste(text);
        }).catch(() => {
          showToast('Please use Ctrl+V or browser paste.', 'warning');
        });
      }
    } else if (e.key === ' ' || (e.altKey && e.key.toLowerCase() === 'e')) {
      e.preventDefault();
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('Permission Denied: Read-only access.', 'warning');
        return;
      }
      if (!isRowLocked(cur.row)) openQuickEntryModal(cur.row);
    } else if (e.shiftKey && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
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
        showToast('Permission Denied: Read-only access.', 'warning');
        return;
      }
      if (colDef && !colDef.isFormula && !colDef.isReadOnly && cur.row !== 5 && !isRowLocked(cur.row)) {
        startCellEdit(cur.element, cur.colLetter, cur.row, colDef);
      }
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('Permission Denied: Read-only access.', 'warning');
        return;
      }
      e.preventDefault();
      deleteSelectedRange(false);
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      if (CurrentUser && CurrentUser.isReadOnly) {
        showToast('Permission Denied: Read-only access.', 'warning');
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

  const flushAndSaveOnUnload = () => {
    if (EditingState.isEditing && EditingState.cell) {
      const input = EditingState.cell.querySelector('input');
      if (input) {
        commitCellEdit(EditingState.cell, EditingState.colLetter, EditingState.row, input.value);
      }
    }
    saveSheetData(false);
  };

  window.addEventListener('beforeunload', flushAndSaveOnUnload);
  window.addEventListener('pagehide', flushAndSaveOnUnload);

  initContextMenu();
}

function initContextMenu() {
  const menu = document.getElementById('contextMenu');
  if (!menu) return;

  window.addEventListener('click', () => {
    menu.classList.add('hidden');
  });

  document.getElementById('ctxCopy')?.addEventListener('click', () => {
    handleClipboardCopy();
    menu.classList.add('hidden');
  });

  document.getElementById('ctxPaste')?.addEventListener('click', () => {
    if (navigator.clipboard && navigator.clipboard.readText) {
      navigator.clipboard.readText().then(text => {
        if (text) handleClipboardPaste(text);
      });
    }
    menu.classList.add('hidden');
  });

  document.getElementById('ctxClear')?.addEventListener('click', () => {
    deleteSelectedRange(false);
    menu.classList.add('hidden');
  });

  document.getElementById('ctxUndo')?.addEventListener('click', () => {
    undoAction();
    menu.classList.add('hidden');
  });

  document.getElementById('ctxRedo')?.addEventListener('click', () => {
    redoAction();
    menu.classList.add('hidden');
  });

  document.getElementById('ctxExport')?.addEventListener('click', () => {
    exportExcelFile();
    menu.classList.add('hidden');
  });
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
  iconSpan.innerHTML = type === 'success' 
    ? '<svg class="w-4 h-4 text-emerald-400 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>' 
    : (type === 'error' 
      ? '<svg class="w-4 h-4 text-red-400 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>' 
      : (type === 'warning'
        ? '<svg class="w-4 h-4 text-amber-400 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
        : '<svg class="w-4 h-4 text-sky-400 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'));

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

// ══════════════════════════════════════════════════════════════════════════════
// APPLICATION ENTRY POINT
// ══════════════════════════════════════════════════════════════════════════════
function startApp() {
  try {
    restoreAppState();
    initMonthYearSelectors();
    initSheetTabButtons();
    bindExcelEvents();
    initAuthSystem();
    initFirebase();
  } catch (err) {
    console.error('Error starting app:', err);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', startApp);
} else {
  startApp();
}
