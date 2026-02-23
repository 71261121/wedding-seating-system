/**
 * Wedding Seating System - Apps Script
 * Production-Ready Implementation
 * 
 * INSTALLATION:
 * 1. Open the Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code
 * 4. Paste this entire code
 * 5. Save (Ctrl+S)
 * 6. Return to sheet and refresh
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  SHEETS: {
    GUEST_LIST_1: 'Guest List 1',
    GUEST_LIST_2: 'Guest List 2',
    GUEST_LIST_3: 'Guest List 3',
    MASTER_GUESTS: 'Master_Guests',
    SEATING_PLAN: 'SEATING_PLAN',
    CONFIG: 'CONFIG',
    DASHBOARD: 'DASHBOARD',
    DEBUG: '_DEBUG',
    SCRIPT_LOG: '_SCRIPT_LOG'
  },
  COLUMNS: {
    MASTER: {
      GUEST_ID: 0,        // A
      GUEST_NAME: 1,      // B
      SOURCE_LIST: 2,     // C
      RSVP: 3,            // D
      GROUP: 4,           // E
      ASSIGNED_TABLE: 5,  // F
      ASSIGNED_SEAT: 6,   // G
      SOURCE_COLOUR: 7,   // H
      IS_ASSIGNED: 8,     // I
      VALIDATION_STATUS: 9 // J
    }
  },
  SEATING: {
    START_ROW: 7,         // First table header row
    TABLES_COUNT: 20,
    SEATS_PER_TABLE: 10,
    ROWS_PER_TABLE: 12    // 1 header + 10 seats + 1 spacer
  }
};

// ============================================
// ON EDIT TRIGGER - Main Logic
// ============================================

function onEdit(e) {
  const sheet = e.range.getSheet();
  
  // Only process SEATING_PLAN sheet
  if (sheet.getName() !== CONFIG.SHEETS.SEATING_PLAN) {
    return;
  }
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Only process column B (guest name column)
  if (col !== 2) {
    return;
  }
  
  // Check if this is a seat row (not a table header or spacer)
  const seatInfo = getTableAndSeatFromCell(row);
  if (!seatInfo) {
    return;
  }
  
  const { tableNum, seatNum } = seatInfo;
  const guestName = e.value;  // New value
  const oldGuestName = e.oldValue;  // Previous value
  
  try {
    if (guestName && guestName.trim() !== '') {
      // ASSIGN: Guest name entered
      handleGuestAssignment(guestName, tableNum, seatNum, row);
    } else if (oldGuestName && oldGuestName.trim() !== '') {
      // UNASSIGN: Guest name cleared
      handleGuestUnassignment(oldGuestName, tableNum, seatNum);
    }
    
    // Update table header
    updateTableHeaderFormatting(tableNum);
    
  } catch (error) {
    // Revert the change on error
    e.range.setValue(oldGuestName || '');
    logAction('ERROR', guestName || oldGuestName, tableNum, seatNum, error.message);
  }
}

// ============================================
// ASSIGNMENT HANDLERS
// ============================================

function handleGuestAssignment(guestName, tableNum, seatNum, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Find guest in Master_Guests
  const masterData = masterSheet.getDataRange().getValues();
  let guestRow = -1;
  let currentAssignment = null;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.MASTER.GUEST_NAME] === guestName) {
      guestRow = i + 1; // 1-indexed
      currentAssignment = {
        table: masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE],
        seat: masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT]
      };
      break;
    }
  }
  
  if (guestRow === -1) {
    throw new Error('Guest not found in Master_Guests');
  }
  
  // Check if guest is already assigned elsewhere
  if (currentAssignment.table && currentAssignment.table !== '') {
    // Guest is already seated - reject assignment
    const seatingPlanSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
    
    // Find and clear the old seat
    const oldSeatRow = findSeatRow(currentAssignment.table, currentAssignment.seat);
    if (oldSeatRow) {
      seatingPlanSheet.getRange(oldSeatRow, 2).clearContent();
    }
    
    // Clear old assignment in Master_Guests
    masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE + 1).clearContent();
    masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT + 1).clearContent();
    masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.IS_ASSIGNED + 1).setValue(false);
  }
  
  // Update Master_Guests with new assignment
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE + 1).setValue(tableNum);
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT + 1).setValue(seatNum);
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.IS_ASSIGNED + 1).setValue(true);
  
  // Apply source colour to the seat cell
  const sourceColour = masterData[guestRow - 1][CONFIG.COLUMNS.MASTER.SOURCE_COLOUR];
  if (sourceColour) {
    const seatCell = seatingSheet.getRange(row, 2);
    seatCell.setBackground(sourceColour);
    seatCell.setFontColor('#FFFFFF');
  }
  
  // Log the action
  logAction('ASSIGN', guestName, tableNum, seatNum, `Assigned to Table ${tableNum}, Seat ${seatNum}`);
  
  // Update old table header if guest was reassigned
  if (currentAssignment.table && currentAssignment.table !== tableNum) {
    updateTableHeaderFormatting(currentAssignment.table);
  }
}

function handleGuestUnassignment(guestName, tableNum, seatNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Find guest in Master_Guests
  const masterData = masterSheet.getDataRange().getValues();
  let guestRow = -1;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.MASTER.GUEST_NAME] === guestName) {
      guestRow = i + 1;
      break;
    }
  }
  
  if (guestRow === -1) {
    return; // Guest not found, nothing to do
  }
  
  // Clear assignment in Master_Guests
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE + 1).clearContent();
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT + 1).clearContent();
  masterSheet.getRange(guestRow, CONFIG.COLUMNS.MASTER.IS_ASSIGNED + 1).setValue(false);
  
  // Clear formatting from seat cell
  const seatRow = findSeatRow(tableNum, seatNum);
  if (seatRow) {
    const seatCell = seatingSheet.getRange(seatRow, 2);
    seatCell.setBackground(null);
    seatCell.setFontColor('#000000');
  }
  
  // Log the action
  logAction('UNASSIGN', guestName, tableNum, seatNum, `Removed from Table ${tableNum}, Seat ${seatNum}`);
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getTableAndSeatFromCell(row) {
  // Calculate which table and seat based on row number
  // Tables start at row 7, each table takes 12 rows (1 header + 10 seats + 1 spacer)
  
  if (row < CONFIG.SEATING.START_ROW) {
    return null;
  }
  
  const relativeRow = row - CONFIG.SEATING.START_ROW;
  const tableIndex = Math.floor(relativeRow / CONFIG.SEATING.ROWS_PER_TABLE);
  const rowWithinTable = relativeRow % CONFIG.SEATING.ROWS_PER_TABLE;
  
  // Row 0 is header, rows 1-10 are seats, row 11 is spacer
  if (rowWithinTable === 0 || rowWithinTable === 11) {
    return null; // Header or spacer row
  }
  
  const seatNum = rowWithinTable; // Seats 1-10
  const tableNum = tableIndex + 1;
  
  return { tableNum, seatNum };
}

function findSeatRow(tableNum, seatNum) {
  // Calculate the row number for a given table and seat
  const tableHeaderRow = CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE;
  return tableHeaderRow + seatNum;
}

function updateTableHeaderFormatting(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Count assigned guests for this table
  const masterData = masterSheet.getDataRange().getValues();
  let seatedCount = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE] == tableNum) {
      seatedCount++;
    }
  }
  
  // Update table header text
  const headerRow = CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE;
  const headerCell = seatingSheet.getRange(headerRow, 2);
  headerCell.setValue(`(${seatedCount}/10 seated)`);
  
  // Apply colour based on fill status
  if (seatedCount === 0) {
    headerCell.setBackground('#FFFFFF');
    headerCell.setFontColor('#666666');
  } else if (seatedCount === 10) {
    headerCell.setBackground('#34A853');
    headerCell.setFontColor('#FFFFFF');
  } else {
    headerCell.setBackground('#FBBC05');
    headerCell.setFontColor('#000000');
  }
}

function logAction(actionType, guestName, table, seat, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.SCRIPT_LOG);
  
  if (!logSheet) return;
  
  const timestamp = new Date();
  logSheet.appendRow([timestamp, actionType, guestName, table, seat, details]);
}

// ============================================
// ON OPEN - Custom Menu
// ============================================

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💒 Wedding Seating Tools')
    .addItem('✅ Validate Data', 'buildDebugReport')
    .addItem('🗑️ Clear All Assignments', 'clearAllAssignments')
    .addItem('🔄 Refresh Table Headers', 'refreshAllTableHeaders')
    .addSeparator()
    .addItem('📊 Update Dashboard', 'updateDashboard')
    .addItem('📝 Test System', 'runSystemTest')
    .addToUi();
}

// ============================================
// MENU FUNCTIONS
// ============================================

function buildDebugReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debugSheet = ss.getSheetByName(CONFIG.SHEETS.DEBUG);
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const guestListSheets = [
    ss.getSheetByName(CONFIG.SHEETS.GUEST_LIST_1),
    ss.getSheetByName(CONFIG.SHEETS.GUEST_LIST_2),
    ss.getSheetByName(CONFIG.SHEETS.GUEST_LIST_3)
  ];
  
  // Clear debug sheet (keep header)
  const lastRow = debugSheet.getLastRow();
  if (lastRow > 1) {
    debugSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  
  const issues = [];
  const masterData = masterSheet.getDataRange().getValues();
  const seenNames = {};
  
  // Check for issues
  for (let i = 1; i < masterData.length; i++) {
    const guestId = masterData[i][CONFIG.COLUMNS.MASTER.GUEST_ID];
    const guestName = masterData[i][CONFIG.COLUMNS.MASTER.GUEST_NAME];
    const isAssigned = masterData[i][CONFIG.COLUMNS.MASTER.IS_ASSIGNED];
    const assignedTable = masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE];
    const assignedSeat = masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT];
    
    // Check for duplicates
    if (guestName) {
      if (seenNames[guestName]) {
        issues.push(['DUPLICATE', guestName, `Appears multiple times in Master_Guests`, i + 1, new Date()]);
      }
      seenNames[guestName] = true;
    }
    
    // Check for invalid assignments
    if (isAssigned === true) {
      if (!assignedTable || !assignedSeat) {
        issues.push(['INVALID_ASSIGNMENT', guestName, 'Marked as assigned but missing table/seat', i + 1, new Date()]);
      }
      if (assignedTable < 1 || assignedTable > 20) {
        issues.push(['INVALID_TABLE', guestName, `Invalid table number: ${assignedTable}`, i + 1, new Date()]);
      }
      if (assignedSeat < 1 || assignedSeat > 10) {
        issues.push(['INVALID_SEAT', guestName, `Invalid seat number: ${assignedSeat}`, i + 1, new Date()]);
      }
    }
  }
  
  // Write issues to debug sheet
  if (issues.length > 0) {
    debugSheet.getRange(2, 1, issues.length, 5).setValues(issues);
  }
  
  // Show result
  const ui = SpreadsheetApp.getUi();
  if (issues.length === 0) {
    ui.alert('✅ Validation Complete', 'No issues found! All data is valid.', ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Issues Found', `Found ${issues.length} issue(s). Check _DEBUG sheet for details.`, ui.ButtonSet.OK);
  }
}

function clearAllAssignments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Clear',
    'This will remove ALL seating assignments. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Clear Master_Guests assignments
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 1) {
    masterSheet.getRange(2, CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE + 1, lastRow - 1, 3).clearContent();
    masterSheet.getRange(2, CONFIG.COLUMNS.MASTER.IS_ASSIGNED + 1, lastRow - 1, 1).setValue(false);
  }
  
  // Clear SEATING_PLAN guest cells
  for (let tableNum = 1; tableNum <= 20; tableNum++) {
    for (let seat = 1; seat <= 10; seat++) {
      const row = findSeatRow(tableNum, seat);
      const cell = seatingSheet.getRange(row, 2);
      cell.clearContent();
      cell.setBackground(null);
      cell.setFontColor('#000000');
    }
    updateTableHeaderFormatting(tableNum);
  }
  
  // Log action
  logAction('CLEAR_ALL', '', '', '', 'All assignments cleared');
  
  ui.alert('✅ Complete', 'All seating assignments have been cleared.', ui.ButtonSet.OK);
}

function refreshAllTableHeaders() {
  for (let tableNum = 1; tableNum <= 20; tableNum++) {
    updateTableHeaderFormatting(tableNum);
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Complete', 'All table headers have been refreshed.', ui.ButtonSet.OK);
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let totalGuests = 0;
  let seatedGuests = 0;
  const groupStats = { Family: { seated: 0, total: 0 }, Friends: { seated: 0, total: 0 }, Colleagues: { seated: 0, total: 0 } };
  const tableStats = {};
  
  for (let i = 1; i < masterData.length; i++) {
    const guestName = masterData[i][CONFIG.COLUMNS.MASTER.GUEST_NAME];
    if (!guestName) continue;
    
    totalGuests++;
    const group = masterData[i][CONFIG.COLUMNS.MASTER.GROUP];
    const isAssigned = masterData[i][CONFIG.COLUMNS.MASTER.IS_ASSIGNED];
    const assignedTable = masterData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE];
    
    if (isAssigned === true) {
      seatedGuests++;
      if (group && groupStats[group]) {
        groupStats[group].seated++;
      }
      if (assignedTable) {
        tableStats[assignedTable] = (tableStats[assignedTable] || 0) + 1;
      }
    }
    
    if (group && groupStats[group]) {
      groupStats[group].total++;
    }
  }
  
  // Update dashboard values
  dashboardSheet.getRange('B4').setValue(totalGuests);
  dashboardSheet.getRange('B5').setValue(seatedGuests);
  dashboardSheet.getRange('B6').setValue(totalGuests - seatedGuests);
  
  // Update group stats
  dashboardSheet.getRange('B15').setValue(groupStats.Family.seated);
  dashboardSheet.getRange('C15').setValue(groupStats.Family.total);
  dashboardSheet.getRange('B16').setValue(groupStats.Friends.seated);
  dashboardSheet.getRange('C16').setValue(groupStats.Friends.total);
  dashboardSheet.getRange('B17').setValue(groupStats.Colleagues.seated);
  dashboardSheet.getRange('C17').setValue(groupStats.Colleagues.total);
  
  // Update table stats
  for (let t = 1; t <= 20; t++) {
    dashboardSheet.getRange(20 + t, 2).setValue(tableStats[t] || 0);
  }
  
  SpreadsheetApp.getUi().alert('✅ Dashboard Updated', `Total: ${totalGuests}, Seated: ${seatedGuests}, Waiting: ${totalGuests - seatedGuests}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSystemTest() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Test 1: Check all sheets exist
  const requiredSheets = Object.values(CONFIG.SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());
  const missingSheets = requiredSheets.filter(s => !existingSheets.includes(s));
  
  if (missingSheets.length > 0) {
    ui.alert('❌ Test Failed', `Missing sheets: ${missingSheets.join(', ')}`, ui.ButtonSet.OK);
    return;
  }
  
  // Test 2: Check Master_Guests has data
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const masterData = masterSheet.getDataRange().getValues();
  if (masterData.length < 2) {
    ui.alert('❌ Test Failed', 'Master_Guests has no data', ui.ButtonSet.OK);
    return;
  }
  
  // Test 3: Check SEATING_PLAN structure
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  if (seatingSheet.getRange('A7').getValue() !== 'TABLE 1') {
    ui.alert('❌ Test Failed', 'SEATING_PLAN structure incorrect', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('✅ All Tests Passed', 'System is ready for use!', ui.ButtonSet.OK);
}

// ============================================
// UTILITY: Sync Guest Lists to Master
// ============================================

function syncGuestListsToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  
  const guestListNames = [
    CONFIG.SHEETS.GUEST_LIST_1,
    CONFIG.SHEETS.GUEST_LIST_2,
    CONFIG.SHEETS.GUEST_LIST_3
  ];
  
  const sourceColours = {
    'Guest List 1': '#4285F4',
    'Guest List 2': '#EA4335',
    'Guest List 3': '#34A853'
  };
  
  const allGuests = [];
  let guestId = 1001;
  
  // Get existing assignments
  const existingData = masterSheet.getDataRange().getValues();
  const existingAssignments = {};
  
  for (let i = 1; i < existingData.length; i++) {
    const name = existingData[i][CONFIG.COLUMNS.MASTER.GUEST_NAME];
    if (name) {
      existingAssignments[name] = {
        table: existingData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_TABLE],
        seat: existingData[i][CONFIG.COLUMNS.MASTER.ASSIGNED_SEAT],
        isAssigned: existingData[i][CONFIG.COLUMNS.MASTER.IS_ASSIGNED]
      };
    }
  }
  
  // Collect guests from all lists
  for (const listName of guestListNames) {
    const sheet = ss.getSheetByName(listName);
    if (!sheet) continue;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const name = data[i][1]; // Column B: Guest Name
      const rsvp = data[i][2]; // Column C: RSVP
      const group = data[i][3]; // Column D: Group
      
      if (name && rsvp === 'Y') {
        const existing = existingAssignments[name] || {};
        allGuests.push([
          guestId++,
          name,
          listName,
          'Y',
          group || '',
          existing.table || '',
          existing.seat || '',
          sourceColours[listName],
          existing.isAssigned || false,
          'OK'
        ]);
      }
    }
  }
  
  // Clear and update Master_Guests
  masterSheet.getRange(2, 1, masterSheet.getLastRow(), 10).clearContent();
  masterSheet.getRange(2, 1, allGuests.length, 10).setValues(allGuests);
  
  SpreadsheetApp.getUi().alert('✅ Sync Complete', `Synchronized ${allGuests.length} guests to Master_Guests`, SpreadsheetApp.getUi().ButtonSet.OK);
}
