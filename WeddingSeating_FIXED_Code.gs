/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * WEDDING SEATING SYSTEM - PRODUCTION READY
 * ═══════════════════════════════════════════════════════════════════════════════
 * CORRECTED VERSION - START_ROW: 12 (Fixed for actual sheet structure)
 * ═══════════════════════════════════════════════════════════════════════════════
 */

const CONFIG = {
  SHEETS: {
    GUEST_LIST_1: 'Guest List 1',
    GUEST_LIST_2: 'Guest List 2',
    GUEST_LIST_3: 'Guest List 3',
    MASTER_GUESTS: 'Master_Guests',
    SEATING_PLAN: 'SEATING_PLAN',
    CONFIG_SHEET: 'CONFIG',
    DASHBOARD: 'DASHBOARD',
    DEBUG: '_DEBUG',
    SCRIPT_LOG: '_SCRIPT_LOG'
  },
  
  COLUMNS: {
    ID: 0,           // A
    NAME: 1,         // B - FullName
    SOURCE: 2,       // C - SourceList
    RSVP: 3,         // D
    GROUP: 4,        // E
    TABLE: 5,        // F - AssignedTable
    SEAT: 6,         // G - AssignedSeat
    COLOUR: 7,       // H - SourceColour
    ASSIGNED: 8,     // I - IsAssigned
    STATUS: 9,       // J - ValidationStatus
    DIETARY: 10,     // K
    ACCESS: 11,      // L
    VIP: 12,         // M
    PRIORITY: 13     // N
  },
  
  SEATING: {
    START_ROW: 12,         // FIXED: Table 1 header starts at row 12
    TABLES: 20,
    SEATS_PER_TABLE: 10,
    ROWS_PER_TABLE: 12     // 1 header + 10 seats + 1 spacer
  },
  
  COLOURS: {
    EMPTY: '#FFFFFF',
    PARTIAL: '#FBBC05',
    FULL: '#34A853'
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// ON EDIT TRIGGER - Main Entry Point
// ═══════════════════════════════════════════════════════════════════════════════

function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Only process SEATING_PLAN sheet
  if (sheetName !== CONFIG.SHEETS.SEATING_PLAN) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Only process column B (guest name column)
  if (col !== 2) return;
  
  // Check if valid seat row
  const seatInfo = getTableAndSeatFromRow(row);
  if (!seatInfo) return;
  
  const { tableNum, seatNum } = seatInfo;
  const guestName = e.value ? String(e.value).trim() : '';
  const oldGuestName = e.oldValue ? String(e.oldValue).trim() : '';
  
  try {
    if (guestName !== '') {
      // ASSIGN guest
      handleGuestAssignment(guestName, tableNum, seatNum, row, e.range);
    } else if (oldGuestName !== '') {
      // UNASSIGN guest
      handleGuestUnassignment(oldGuestName, tableNum, seatNum);
    }
    
    // Update table header
    updateTableHeader(tableNum);
    updateDashboard();
    
  } catch (error) {
    console.error('Error:', error);
    e.range.setValue(oldGuestName || '');
    SpreadsheetApp.getUi().alert('⚠️ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ASSIGNMENT HANDLERS
// ═══════════════════════════════════════════════════════════════════════════════

function handleGuestAssignment(guestName, tableNum, seatNum, seatRow, seatCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  // Find guest in Master_Guests
  let guestRowIndex = -1;
  let guestRow = -1;
  let currentTable = null;
  let currentSeat = null;
  let isCurrentlyAssigned = false;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.NAME] === guestName) {
      guestRowIndex = i;
      guestRow = i + 1;
      currentTable = masterData[i][CONFIG.COLUMNS.TABLE];
      currentSeat = masterData[i][CONFIG.COLUMNS.SEAT];
      isCurrentlyAssigned = masterData[i][CONFIG.COLUMNS.ASSIGNED];
      break;
    }
  }
  
  // Validation: Guest must exist
  if (guestRowIndex === -1) {
    throw new Error(`Guest "${guestName}" not found in Master_Guests!\n\nMake sure the guest name matches exactly.`);
  }
  
  // Get guest colour
  const guestColour = masterData[guestRowIndex][CONFIG.COLUMNS.COLOUR];
  const guestVIP = masterData[guestRowIndex][CONFIG.COLUMNS.VIP];
  
  // Check if already assigned elsewhere - clear old seat
  if (isCurrentlyAssigned === true || isCurrentlyAssigned === 'TRUE') {
    if (currentTable && currentSeat) {
      // Clear old seat in SEATING_PLAN
      const oldSeatRow = findSeatRow(currentTable, currentSeat);
      if (oldSeatRow) {
        const oldCell = seatingSheet.getRange(oldSeatRow, 2);
        oldCell.clearContent();
        oldCell.setBackground(null);
        oldCell.setFontColor('#000000');
      }
      updateTableHeader(currentTable);
    }
  }
  
  // Update Master_Guests
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.TABLE + 1).setValue(tableNum);
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.SEAT + 1).setValue(seatNum);
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.ASSIGNED + 1).setValue(true);
  
  // Apply colour to seat cell
  if (guestColour) {
    seatCell.setBackground(guestColour);
    seatCell.setFontColor('#FFFFFF');
    seatCell.setFontWeight('bold');
  }
  
  // Log action
  logAction('ASSIGN', guestName, tableNum, seatNum, `VIP: ${guestVIP}`);
}

function handleGuestUnassignment(guestName, tableNum, seatNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const masterData = masterSheet.getDataRange().getValues();
  let guestRow = -1;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.NAME] === guestName) {
      guestRow = i + 1;
      break;
    }
  }
  
  if (guestRow === -1) return;
  
  // Clear assignment in Master_Guests
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.TABLE + 1).clearContent();
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.SEAT + 1).clearContent();
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.ASSIGNED + 1).setValue(false);
  
  // Clear seat formatting
  const seatRow = findSeatRow(tableNum, seatNum);
  if (seatRow) {
    const seatCell = seatingSheet.getRange(seatRow, 2);
    seatCell.setBackground(null);
    seatCell.setFontColor('#000000');
    seatCell.setFontWeight('normal');
  }
  
  logAction('UNASSIGN', guestName, tableNum, seatNum, 'Seat cleared');
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function getTableAndSeatFromRow(row) {
  if (row < CONFIG.SEATING.START_ROW) return null;
  
  const relativeRow = row - CONFIG.SEATING.START_ROW;
  const tableIndex = Math.floor(relativeRow / CONFIG.SEATING.ROWS_PER_TABLE);
  const rowWithinTable = relativeRow % CONFIG.SEATING.ROWS_PER_TABLE;
  
  // Row 0 is header, rows 1-10 are seats, row 11 is spacer
  if (rowWithinTable === 0 || rowWithinTable === 11) return null;
  
  const tableNum = tableIndex + 1;
  const seatNum = rowWithinTable;
  
  if (tableNum > CONFIG.SEATING.TABLES) return null;
  
  return { tableNum, seatNum };
}

function findSeatRow(tableNum, seatNum) {
  if (tableNum < 1 || tableNum > CONFIG.SEATING.TABLES) return null;
  if (seatNum < 1 || seatNum > CONFIG.SEATING.SEATS_PER_TABLE) return null;
  
  return CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE + seatNum;
}

function updateTableHeader(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const masterData = masterSheet.getDataRange().getValues();
  let seatedCount = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.TABLE] == tableNum) {
      seatedCount++;
    }
  }
  
  const headerRow = CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE;
  const countCell = seatingSheet.getRange(headerRow, 2);
  countCell.setValue(`(${seatedCount}/10)`);
  
  if (seatedCount === 0) {
    countCell.setBackground(CONFIG.COLOURS.EMPTY);
    countCell.setFontColor('#666666');
  } else if (seatedCount === 10) {
    countCell.setBackground(CONFIG.COLOURS.FULL);
    countCell.setFontColor('#FFFFFF');
  } else {
    countCell.setBackground(CONFIG.COLOURS.PARTIAL);
    countCell.setFontColor('#000000');
  }
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  
  if (!dashboardSheet) return;
  
  const masterData = masterSheet.getDataRange().getValues();
  let total = 0, seated = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.NAME]) {
      total++;
      if (masterData[i][CONFIG.COLUMNS.ASSIGNED] === true || masterData[i][CONFIG.COLUMNS.ASSIGNED] === 'TRUE') {
        seated++;
      }
    }
  }
  
  // Update dashboard (adjust cell references as needed)
  dashboardSheet.getRange('B4').setValue(total);
  dashboardSheet.getRange('B5').setValue(seated);
  dashboardSheet.getRange('B6').setValue(total - seated);
}

function logAction(actionType, guestName, table, seat, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.SCRIPT_LOG);
  
  if (!logSheet) return;
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([timestamp, actionType, guestName, table, seat, '', '', details]);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ON OPEN - Custom Menu
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💒 Wedding Seating')
    .addItem('✅ Validate Data', 'validateData')
    .addItem('🔄 Refresh All Tables', 'refreshAllTables')
    .addItem('📊 Update Dashboard', 'updateDashboard')
    .addSeparator()
    .addItem('🪑 Auto-Seat All Guests', 'autoSeatAll')
    .addItem('🗑️ Clear All Assignments', 'clearAllAssignments')
    .addSeparator()
    .addItem('🧪 Run System Test', 'runSystemTest')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function validateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const debugSheet = ss.getSheetByName(CONFIG.SHEETS.DEBUG);
  const ui = SpreadsheetApp.getUi();
  
  const masterData = masterSheet.getDataRange().getValues();
  const issues = [];
  const seenNames = {};
  const assignedSeats = {};
  
  for (let i = 1; i < masterData.length; i++) {
    const name = masterData[i][CONFIG.COLUMNS.NAME];
    const table = masterData[i][CONFIG.COLUMNS.TABLE];
    const seat = masterData[i][CONFIG.COLUMNS.SEAT];
    const isAssigned = masterData[i][CONFIG.COLUMNS.ASSIGNED];
    
    if (!name) continue;
    
    // Check duplicates
    if (seenNames[name]) {
      issues.push(['DUPLICATE', name, 'Name appears multiple times', i + 1, new Date()]);
    }
    seenNames[name] = true;
    
    // Check double-booking
    if (isAssigned === true || isAssigned === 'TRUE') {
      if (table && seat) {
        const seatKey = `${table}-${seat}`;
        if (assignedSeats[seatKey]) {
          issues.push(['DOUBLE_BOOK', name, `Seat ${seatKey} already taken by ${assignedSeats[seatKey]}`, i + 1, new Date()]);
        }
        assignedSeats[seatKey] = name;
      }
    }
  }
  
  // Clear and update debug sheet
  const lastRow = debugSheet.getLastRow();
  if (lastRow > 1) {
    debugSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  
  if (issues.length > 0) {
    debugSheet.getRange(2, 1, issues.length, 5).setValues(issues);
    ui.alert('⚠️ Issues Found', `Found ${issues.length} issue(s).\nCheck _DEBUG sheet for details.`, ui.ButtonSet.OK);
  } else {
    ui.alert('✅ Validation Complete', 'No issues found! All data is valid.', ui.ButtonSet.OK);
  }
}

function refreshAllTables() {
  for (let t = 1; t <= CONFIG.SEATING.TABLES; t++) {
    updateTableHeader(t);
  }
  updateDashboard();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All tables refreshed!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function autoSeatAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🪑 Auto-Seating',
    'This will automatically seat all unassigned guests.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const masterData = masterSheet.getDataRange().getValues();
  const unassigned = [];
  
  for (let i = 1; i < masterData.length; i++) {
    const isAssigned = masterData[i][CONFIG.COLUMNS.ASSIGNED];
    const name = masterData[i][CONFIG.COLUMNS.NAME];
    if (name && isAssigned !== true && isAssigned !== 'TRUE') {
      unassigned.push({
        row: i + 1,
        name: name,
        colour: masterData[i][CONFIG.COLUMNS.COLOUR]
      });
    }
  }
  
  if (unassigned.length === 0) {
    ui.alert('ℹ️ Info', 'All guests are already seated.', ui.ButtonSet.OK);
    return;
  }
  
  // Find available seats
  const assignedSeats = new Set();
  for (let i = 1; i < masterData.length; i++) {
    const table = masterData[i][CONFIG.COLUMNS.TABLE];
    const seat = masterData[i][CONFIG.COLUMNS.SEAT];
    if (table && seat) {
      assignedSeats.add(`${table}-${seat}`);
    }
  }
  
  const availableSeats = [];
  for (let t = 1; t <= CONFIG.SEATING.TABLES; t++) {
    for (let s = 1; s <= CONFIG.SEATING.SEATS_PER_TABLE; s++) {
      if (!assignedSeats.has(`${t}-${s}`)) {
        availableSeats.push({ table: t, seat: s });
      }
    }
  }
  
  // Assign guests
  let assigned = 0;
  for (const guest of unassigned) {
    if (availableSeats.length === 0) break;
    
    const seat = availableSeats.shift();
    
    // Update Master_Guests
    masterSheet.getRange(guest.row, CONFIG.COLUMNS.TABLE + 1).setValue(seat.table);
    masterSheet.getRange(guest.row, CONFIG.COLUMNS.SEAT + 1).setValue(seat.seat);
    masterSheet.getRange(guest.row, CONFIG.COLUMNS.ASSIGNED + 1).setValue(true);
    
    // Update SEATING_PLAN
    const seatRow = findSeatRow(seat.table, seat.seat);
    const seatCell = seatingSheet.getRange(seatRow, 2);
    seatCell.setValue(guest.name);
    
    if (guest.colour) {
      seatCell.setBackground(guest.colour);
      seatCell.setFontColor('#FFFFFF');
      seatCell.setFontWeight('bold');
    }
    
    updateTableHeader(seat.table);
    assigned++;
  }
  
  updateDashboard();
  logAction('AUTO_SEAT', '', '', '', `Auto-seated ${assigned} guests`);
  
  ui.alert('✅ Complete', `Successfully seated ${assigned} guests.`, ui.ButtonSet.OK);
}

function clearAllAssignments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Clear',
    'This will remove ALL seating assignments.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 1) {
    masterSheet.getRange(2, CONFIG.COLUMNS.TABLE + 1, lastRow - 1, 3).clearContent();
    masterSheet.getRange(2, CONFIG.COLUMNS.ASSIGNED + 1, lastRow - 1, 1).setValue(false);
  }
  
  // Clear all seats
  for (let t = 1; t <= CONFIG.SEATING.TABLES; t++) {
    for (let s = 1; s <= CONFIG.SEATING.SEATS_PER_TABLE; s++) {
      const seatRow = findSeatRow(t, s);
      const cell = seatingSheet.getRange(seatRow, 2);
      cell.clearContent();
      cell.setBackground(null);
      cell.setFontColor('#000000');
      cell.setFontWeight('normal');
    }
    updateTableHeader(t);
  }
  
  updateDashboard();
  logAction('CLEAR_ALL', '', '', '', 'All assignments cleared');
  
  ui.alert('✅ Complete', 'All seating assignments cleared.', ui.ButtonSet.OK);
}

function runSystemTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tests = [];
  
  // Test 1: Sheets exist
  const requiredSheets = Object.values(CONFIG.SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());
  const missingSheets = requiredSheets.filter(s => !existingSheets.includes(s));
  
  tests.push({
    name: 'Sheet Structure',
    status: missingSheets.length === 0 ? 'PASS' : 'FAIL',
    detail: missingSheets.length === 0 ? 'All sheets exist' : `Missing: ${missingSheets.join(', ')}`
  });
  
  // Test 2: Master_Guests has data
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const masterData = masterSheet.getDataRange().getValues();
  
  tests.push({
    name: 'Guest Data',
    status: masterData.length > 1 ? 'PASS' : 'FAIL',
    detail: `${masterData.length - 1} guests found`
  });
  
  // Test 3: SEATING_PLAN structure
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  const table1Cell = seatingSheet.getRange(CONFIG.SEATING.START_ROW, 1).getValue();
  
  tests.push({
    name: 'Seating Plan',
    status: table1Cell.includes('TABLE 1') ? 'PASS' : 'WARN',
    detail: table1Cell.includes('TABLE 1') ? 'Structure correct' : 'Check table structure'
  });
  
  // Build report
  let report = '═══════════════════════════════════════════════════════════════\n';
  report += '              SYSTEM TEST RESULTS\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  
  for (const test of tests) {
    const icon = test.status === 'PASS' ? '✅' : test.status === 'WARN' ? '⚠️' : '❌';
    report += `${icon} ${test.name}: ${test.status}\n   ${test.detail}\n\n`;
  }
  
  ui.alert('🧪 System Test', report, ui.ButtonSet.OK);
}
