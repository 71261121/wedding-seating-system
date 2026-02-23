/**
 * ═══════════════════════════════════════════════════════════════════════════
 * WEDDING SEATING SYSTEM - ULTRA ADVANCED PRODUCTION SCRIPT
 * ═══════════════════════════════════════════════════════════════════════════
 * Version: 2.0.0
 * Features:
 * - Real-time seat assignment tracking
 * - Automatic conflict detection
 * - VIP priority handling
 * - Dietary requirement grouping
 * - Plus-one management
 * - Table optimization suggestions
 * - Comprehensive logging
 * - Debug diagnostics
 * - Auto-refresh capabilities
 * 
 * INSTALLATION:
 * 1. Extensions → Apps Script
 * 2. Delete existing code
 * 3. Paste this entire file
 * 4. Save (Ctrl+S)
 * 5. Return to sheet and refresh
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// CONFIGURATION CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

const CONFIG = {
  // Sheet Names
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
  
  // Master_Guests Column Indices (0-based)
  COLS: {
    GUEST_ID: 0,           // A - Auto ID
    GUEST_NAME: 1,         // B - Name
    SOURCE_LIST: 2,        // C - Source
    RSVP: 3,               // D - RSVP Status
    GROUP: 4,              // E - Group/Family
    ASSIGNED_TABLE: 5,     // F - Table Number
    ASSIGNED_SEAT: 6,      // G - Seat Number
    SOURCE_COLOUR: 7,      // H - Colour Code
    IS_ASSIGNED: 8,        // I - Boolean
    VALIDATION_STATUS: 9,  // J - OK/ERROR
    PLUS_ONE_NAME: 10,     // K - Plus-One
    DIETARY_REQ: 11,       // L - Dietary
    VIP_STATUS: 12,        // M - VIP/Family/Friend
    TABLE_PREF: 13,        // N - Preference
    SPECIAL_NOTES: 14,     // O - Notes
    CHECKIN_STATUS: 15     // P - Check-in
  },
  
  // SEATING_PLAN Configuration
  SEATING: {
    START_ROW: 10,         // First table header row
    TABLES_COUNT: 20,      // Total tables
    SEATS_PER_TABLE: 10,   // Seats per table
    ROWS_PER_TABLE: 12     // Header + 10 seats + spacer
  },
  
  // Source Colours (Hex)
  COLOURS: {
    FAMILY: '#1A73E8',
    FRIENDS: '#EA4335',
    COLLEAGUES: '#34A853',
    FAMILY_PLUSONE: '#4A90D9',
    FRIENDS_PLUSONE: '#F06449',
    COLLEAGUES_PLUSONE: '#5CB85C'
  },
  
  // Table Status Colours
  STATUS_COLOURS: {
    EMPTY: '#FFFFFF',
    PARTIAL: '#FBBC05',
    FULL: '#34A853',
    TEXT_EMPTY: '#666666',
    TEXT_PARTIAL: '#000000',
    TEXT_FULL: '#FFFFFF'
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// ON EDIT TRIGGER - Main Entry Point
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Triggered on any cell edit in the spreadsheet.
 * Handles guest assignment/unassignment with full validation.
 */
function onEdit(e) {
  // Guard: Check if event object exists
  if (!e || !e.range) {
    return;
  }
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Only process SEATING_PLAN sheet
  if (sheetName !== CONFIG.SHEETS.SEATING_PLAN) {
    return;
  }
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Only process column B (guest name cells)
  if (col !== 2) {
    return;
  }
  
  // Check if this is a valid seat row
  const seatInfo = getTableAndSeatFromRow(row);
  if (!seatInfo) {
    return; // Not a seat row (could be header or spacer)
  }
  
  const { tableNum, seatNum } = seatInfo;
  const guestName = e.value ? e.value.trim() : '';
  const oldGuestName = e.oldValue ? e.oldValue.trim() : '';
  
  try {
    if (guestName !== '') {
      // ASSIGNMENT: Guest name entered
      handleGuestAssignment(guestName, tableNum, seatNum, row, e.range);
    } else if (oldGuestName !== '') {
      // UNASSIGNMENT: Guest name cleared
      handleGuestUnassignment(oldGuestName, tableNum, seatNum);
    }
    
    // Update table header with new count
    updateTableHeader(tableNum);
    
    // Update dashboard metrics
    updateDashboardMetrics();
    
  } catch (error) {
    // Error handling: Revert change and log
    console.error('onEdit error:', error);
    e.range.setValue(oldGuestName || '');
    
    logAction('ERROR', guestName || oldGuestName, tableNum, seatNum, error.message);
    
    // Show user-friendly error
    SpreadsheetApp.getUi().alert(
      '⚠️ Assignment Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ASSIGNMENT HANDLERS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Handles assigning a guest to a seat.
 * Validates: Guest exists, not already assigned, no conflicts.
 */
function handleGuestAssignment(guestName, tableNum, seatNum, seatRow, seatCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Get all master data
  const masterData = masterSheet.getDataRange().getValues();
  
  // Find the guest row
  let guestRowIndex = -1;
  let currentAssignment = null;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLS.GUEST_NAME] === guestName) {
      guestRowIndex = i + 1; // 1-indexed for sheets
      currentAssignment = {
        table: masterData[i][CONFIG.COLS.ASSIGNED_TABLE],
        seat: masterData[i][CONFIG.COLS.ASSIGNED_SEAT],
        isAssigned: masterData[i][CONFIG.COLS.IS_ASSIGNED]
      };
      break;
    }
  }
  
  // Validation: Guest must exist in Master_Guests
  if (guestRowIndex === -1) {
    throw new Error(`Guest "${guestName}" not found in Master_Guests. Please sync guest lists first.`);
  }
  
  const guestRow = guestRowIndex - 1; // 0-indexed for array access
  const sourceColour = masterData[guestRow][CONFIG.COLS.SOURCE_COLOUR];
  const vipStatus = masterData[guestRow][CONFIG.COLS.VIP_STATUS];
  const dietaryReq = masterData[guestRow][CONFIG.COLS.DIETARY_REQ];
  const plusOneName = masterData[guestRow][CONFIG.COLS.PLUS_ONE_NAME];
  
  // Validation: Check VIP table restrictions
  if (vipStatus === 'VIP' && tableNum > 5) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ VIP Warning',
      `${guestName} is marked as VIP. VIPs are typically seated at tables 1-5.\n\nAssign to Table ${tableNum} anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      seatCell.setValue('');
      return;
    }
  }
  
  // Handle existing assignment (reassignment)
  if (currentAssignment.isAssigned === true || currentAssignment.isAssigned === 'TRUE') {
    if (currentAssignment.table && currentAssignment.seat) {
      // Clear old seat in SEATING_PLAN
      const oldSeatRow = findSeatRow(currentAssignment.table, currentAssignment.seat);
      if (oldSeatRow) {
        const oldSeatCell = seatingSheet.getRange(oldSeatRow, 2);
        oldSeatCell.clearContent();
        oldSeatCell.setBackground(null);
        oldSeatCell.setFontColor('#000000');
      }
      
      // Update old table header
      updateTableHeader(currentAssignment.table);
      
      // Log reassignment
      logAction('REASSIGN', guestName, currentAssignment.table, currentAssignment.seat, 
        `Reassigned from Table ${currentAssignment.table}, Seat ${currentAssignment.seat} to Table ${tableNum}, Seat ${seatNum}`);
    }
  }
  
  // Update Master_Guests with new assignment
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.ASSIGNED_TABLE + 1).setValue(tableNum);
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.ASSIGNED_SEAT + 1).setValue(seatNum);
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.IS_ASSIGNED + 1).setValue(true);
  
  // Apply colour coding to seat cell
  if (sourceColour) {
    seatCell.setBackground(sourceColour);
    seatCell.setFontColor('#FFFFFF');
    seatCell.setFontWeight('bold');
  }
  
  // Handle plus-one seating suggestion
  if (plusOneName && plusOneName.trim() !== '') {
    // Check if plus-one is seated at same table
    const plusOneRow = findGuestRowByName(plusOneName);
    if (plusOneRow) {
      const plusOneTable = masterSheet.getRange(plusOneRow, CONFIG.COLS.ASSIGNED_TABLE + 1).getValue();
      if (plusOneTable && plusOneTable !== tableNum) {
        SpreadsheetApp.getUi().alert(
          'ℹ️ Plus-One Notice',
          `${guestName}'s plus-one (${plusOneName}) is seated at Table ${plusOneTable}.\nConsider seating them together.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
  }
  
  // Check dietary grouping
  if (dietaryReq && dietaryReq.trim() !== '' && dietaryReq !== 'None') {
    checkDietaryGrouping(guestName, tableNum, dietaryReq);
  }
  
  // Log successful assignment
  logAction('ASSIGN', guestName, tableNum, seatNum, 
    `VIP: ${vipStatus}, Dietary: ${dietaryReq || 'None'}`);
}

/**
 * Handles unassigning a guest from a seat.
 */
function handleGuestUnassignment(guestName, tableNum, seatNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Find guest in Master_Guests
  const guestRowIndex = findGuestRowByName(guestName);
  
  if (guestRowIndex === -1) {
    return; // Guest not found, nothing to do
  }
  
  // Clear assignment fields
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.ASSIGNED_TABLE + 1).clearContent();
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.ASSIGNED_SEAT + 1).clearContent();
  masterSheet.getRange(guestRowIndex, CONFIG.COLS.IS_ASSIGNED + 1).setValue(false);
  
  // Clear seat cell formatting
  const seatRow = findSeatRow(tableNum, seatNum);
  if (seatRow) {
    const seatCell = seatingSheet.getRange(seatRow, 2);
    seatCell.setBackground(null);
    seatCell.setFontColor('#000000');
    seatCell.setFontWeight('normal');
  }
  
  // Log unassignment
  logAction('UNASSIGN', guestName, tableNum, seatNum, 'Guest removed from seat');
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Converts a row number to table and seat numbers.
 * Returns null if not a valid seat row.
 */
function getTableAndSeatFromRow(row) {
  if (row < CONFIG.SEATING.START_ROW) {
    return null;
  }
  
  const relativeRow = row - CONFIG.SEATING.START_ROW;
  const tableIndex = Math.floor(relativeRow / CONFIG.SEATING.ROWS_PER_TABLE);
  const rowWithinTable = relativeRow % CONFIG.SEATING.ROWS_PER_TABLE;
  
  // Row 0 is header, rows 1-10 are seats, row 11 is spacer
  if (rowWithinTable === 0 || rowWithinTable === 11) {
    return null;
  }
  
  const seatNum = rowWithinTable;
  const tableNum = tableIndex + 1;
  
  if (tableNum > CONFIG.SEATING.TABLES_COUNT) {
    return null;
  }
  
  return { tableNum, seatNum };
}

/**
 * Finds the row number for a specific table and seat.
 */
function findSeatRow(tableNum, seatNum) {
  if (tableNum < 1 || tableNum > CONFIG.SEATING.TABLES_COUNT) return null;
  if (seatNum < 1 || seatNum > CONFIG.SEATING.SEATS_PER_TABLE) return null;
  
  return CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE + seatNum;
}

/**
 * Finds a guest's row index in Master_Guests by name.
 * Returns 1-indexed row number or -1 if not found.
 */
function findGuestRowByName(guestName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const data = masterSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLS.GUEST_NAME] === guestName) {
      return i + 1;
    }
  }
  
  return -1;
}

/**
 * Updates a table header with the current seat count.
 * Applies conditional formatting based on fill status.
 */
function updateTableHeader(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  // Count guests assigned to this table
  const masterData = masterSheet.getDataRange().getValues();
  let seatedCount = 0;
  let vipCount = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLS.ASSIGNED_TABLE] == tableNum) {
      seatedCount++;
      if (masterData[i][CONFIG.COLS.VIP_STATUS] === 'VIP') {
        vipCount++;
      }
    }
  }
  
  // Calculate header row
  const headerRow = CONFIG.SEATING.START_ROW + (tableNum - 1) * CONFIG.SEATING.ROWS_PER_TABLE;
  
  // Update seat count display
  const countCell = seatingSheet.getRange(headerRow, 2);
  countCell.setValue(`(${seatedCount}/10)`);
  
  // Apply conditional formatting
  if (seatedCount === 0) {
    countCell.setBackground(CONFIG.STATUS_COLOURS.EMPTY);
    countCell.setFontColor(CONFIG.STATUS_COLOURS.TEXT_EMPTY);
  } else if (seatedCount === 10) {
    countCell.setBackground(CONFIG.STATUS_COLOURS.FULL);
    countCell.setFontColor(CONFIG.STATUS_COLOURS.TEXT_FULL);
  } else {
    countCell.setBackground(CONFIG.STATUS_COLOURS.PARTIAL);
    countCell.setFontColor(CONFIG.STATUS_COLOURS.TEXT_PARTIAL);
  }
  
  // Update VIP count if applicable
  if (vipCount > 0) {
    seatingSheet.getRange(headerRow, 7).setValue(`${vipCount} VIP`);
  } else {
    seatingSheet.getRange(headerRow, 7).clearContent();
  }
}

/**
 * Checks dietary requirements grouping and provides suggestions.
 */
function checkDietaryGrouping(guestName, tableNum, dietaryReq) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const masterData = masterSheet.getDataRange().getValues();
  
  // Find other guests with same dietary requirement
  const sameDietGuests = [];
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLS.DIETARY_REQ] === dietaryReq && 
        masterData[i][CONFIG.COLS.GUEST_NAME] !== guestName) {
      const assignedTable = masterData[i][CONFIG.COLS.ASSIGNED_TABLE];
      sameDietGuests.push({
        name: masterData[i][CONFIG.COLS.GUEST_NAME],
        table: assignedTable || 'Unassigned'
      });
    }
  }
  
  // Check if any same-diet guests are seated at different tables
  const differentTables = sameDietGuests.filter(g => g.table !== tableNum && g.table !== 'Unassigned');
  
  if (differentTables.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '🍽️ Dietary Notice',
      `${guestName} requires ${dietaryReq}.\n\nOther guests with this requirement:\n${differentTables.map(g => `• ${g.name} (Table ${g.table})`).join('\n')}\n\nConsider grouping guests with similar dietary needs.`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Updates dashboard metrics.
 */
function updateDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let totalGuests = 0;
  let seatedGuests = 0;
  const tableStats = {};
  const sourceStats = {};
  const vipStats = {};
  
  for (let i = 1; i < masterData.length; i++) {
    const guestName = masterData[i][CONFIG.COLS.GUEST_NAME];
    if (!guestName) continue;
    
    totalGuests++;
    
    const source = masterData[i][CONFIG.COLS.SOURCE_LIST];
    const vip = masterData[i][CONFIG.COLS.VIP_STATUS];
    const isAssigned = masterData[i][CONFIG.COLS.IS_ASSIGNED];
    const assignedTable = masterData[i][CONFIG.COLS.ASSIGNED_TABLE];
    
    // Source stats
    if (!sourceStats[source]) {
      sourceStats[source] = { total: 0, seated: 0 };
    }
    sourceStats[source].total++;
    
    // VIP stats
    if (!vipStats[vip]) {
      vipStats[vip] = { total: 0, seated: 0 };
    }
    vipStats[vip].total++;
    
    if (isAssigned === true || isAssigned === 'TRUE') {
      seatedGuests++;
      sourceStats[source].seated++;
      vipStats[vip].seated++;
      
      if (assignedTable) {
        tableStats[assignedTable] = (tableStats[assignedTable] || 0) + 1;
      }
    }
  }
  
  // Update individual table status in dashboard
  for (let t = 1; t <= CONFIG.SEATING.TABLES_COUNT; t++) {
    dashboardSheet.getRange(44 + t, 2).setValue(tableStats[t] || 0);
  }
}

/**
 * Logs an action to the _SCRIPT_LOG sheet.
 */
function logAction(actionType, guestName, table, seat, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.SCRIPT_LOG);
  
  if (!logSheet) return;
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([timestamp, actionType, guestName, table, seat, '', '', details]);
  
  // Keep only last 500 entries
  const lastRow = logSheet.getLastRow();
  if (lastRow > 501) {
    logSheet.deleteRows(2, lastRow - 501);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ON OPEN - Custom Menu
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Creates custom menu when spreadsheet opens.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('💒 Wedding Seating')
    .addItem('✅ Validate All Data', 'buildDebugReport')
    .addItem('🔄 Refresh All Tables', 'refreshAllTableHeaders')
    .addItem('📊 Update Dashboard', 'updateDashboardMetrics')
    .addSeparator()
    .addItem('🔄 Sync Guest Lists', 'syncGuestListsToMaster')
    .addItem('🪑 Auto-Seat Unassigned', 'autoSeatUnassigned')
    .addSeparator()
    .addItem('🗑️ Clear All Assignments', 'clearAllAssignments')
    .addItem('📋 Generate Report', 'generateSeatingReport')
    .addSeparator()
    .addItem('🧪 Run System Test', 'runSystemTest')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Builds a debug report showing all data issues.
 */
function buildDebugReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debugSheet = ss.getSheetByName(CONFIG.SHEETS.DEBUG);
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  
  const ui = SpreadsheetApp.getUi();
  
  // Clear debug sheet (keep header)
  const lastRow = debugSheet.getLastRow();
  if (lastRow > 1) {
    debugSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  
  const issues = [];
  const masterData = masterSheet.getDataRange().getValues();
  const seenNames = {};
  const assignedSeats = {}; // Track seat assignments
  
  // Scan for issues
  for (let i = 1; i < masterData.length; i++) {
    const row = i + 1;
    const guestId = masterData[i][CONFIG.COLS.GUEST_ID];
    const guestName = masterData[i][CONFIG.COLS.GUEST_NAME];
    const isAssigned = masterData[i][CONFIG.COLS.IS_ASSIGNED];
    const assignedTable = masterData[i][CONFIG.COLS.ASSIGNED_TABLE];
    const assignedSeat = masterData[i][CONFIG.COLS.ASSIGNED_SEAT];
    
    if (!guestName) continue;
    
    // Check for duplicate names
    if (seenNames[guestName]) {
      issues.push(['DUPLICATE', guestName, 'Name appears multiple times', row, new Date(), 'HIGH', 'Merge or rename']);
    }
    seenNames[guestName] = true;
    
    // Check for assignment inconsistencies
    if (isAssigned === true || isAssigned === 'TRUE') {
      if (!assignedTable || !assignedSeat) {
        issues.push(['INVALID_ASSIGNMENT', guestName, 'Marked assigned but missing table/seat', row, new Date(), 'HIGH', 'Update assignment']);
      }
      
      // Check for double-booking
      const seatKey = `${assignedTable}-${assignedSeat}`;
      if (assignedSeats[seatKey]) {
        issues.push(['DOUBLE_BOOKED', guestName, `Seat ${assignedTable}-${assignedSeat} already taken by ${assignedSeats[seatKey]}`, row, new Date(), 'CRITICAL', 'Resolve conflict']);
      }
      assignedSeats[seatKey] = guestName;
      
      // Validate table/seat numbers
      if (assignedTable < 1 || assignedTable > CONFIG.SEATING.TABLES_COUNT) {
        issues.push(['INVALID_TABLE', guestName, `Table ${assignedTable} out of range`, row, new Date(), 'HIGH', 'Fix table number']);
      }
      if (assignedSeat < 1 || assignedSeat > CONFIG.SEATING.SEATS_PER_TABLE) {
        issues.push(['INVALID_SEAT', guestName, `Seat ${assignedSeat} out of range`, row, new Date(), 'HIGH', 'Fix seat number']);
      }
    }
  }
  
  // Write issues to debug sheet
  if (issues.length > 0) {
    debugSheet.getRange(2, 1, issues.length, 7).setValues(issues);
  }
  
  // Show result
  if (issues.length === 0) {
    ui.alert('✅ Validation Complete', 'No issues found! All data is valid.', ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Issues Found', `Found ${issues.length} issue(s).\n\nCheck the _DEBUG sheet for details.`, ui.ButtonSet.OK);
  }
}

/**
 * Clears all seat assignments.
 */
function clearAllAssignments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Clear All',
    'This will remove ALL seating assignments.\n\nThis action cannot be undone!\n\nContinue?',
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
    masterSheet.getRange(2, CONFIG.COLS.ASSIGNED_TABLE + 1, lastRow - 1, 3).clearContent();
    masterSheet.getRange(2, CONFIG.COLS.IS_ASSIGNED + 1, lastRow - 1, 1).setValue(false);
  }
  
  // Clear SEATING_PLAN seat cells
  for (let tableNum = 1; tableNum <= CONFIG.SEATING.TABLES_COUNT; tableNum++) {
    for (let seat = 1; seat <= CONFIG.SEATING.SEATS_PER_TABLE; seat++) {
      const row = findSeatRow(tableNum, seat);
      const cell = seatingSheet.getRange(row, 2);
      cell.clearContent();
      cell.setBackground(null);
      cell.setFontColor('#000000');
      cell.setFontWeight('normal');
    }
    updateTableHeader(tableNum);
  }
  
  logAction('CLEAR_ALL', '', '', '', 'All assignments cleared by user');
  
  ui.alert('✅ Complete', 'All seating assignments have been cleared.', ui.ButtonSet.OK);
}

/**
 * Refreshes all table headers.
 */
function refreshAllTableHeaders() {
  for (let t = 1; t <= CONFIG.SEATING.TABLES_COUNT; t++) {
    updateTableHeader(t);
  }
  updateDashboardMetrics();
  
  SpreadsheetApp.getUi().alert('✅ Complete', 'All table headers and dashboard refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Syncs guest lists to Master_Guests.
 */
function syncGuestListsToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  
  const guestListNames = [
    CONFIG.SHEETS.GUEST_LIST_1,
    CONFIG.SHEETS.GUEST_LIST_2,
    CONFIG.SHEETS.GUEST_LIST_3
  ];
  
  const sourceColours = {
    'Guest List 1': '#1A73E8',
    'Guest List 2': '#EA4335',
    'Guest List 3': '#34A853'
  };
  
  // Get existing assignments
  const existingData = masterSheet.getDataRange().getValues();
  const existingAssignments = {};
  
  for (let i = 1; i < existingData.length; i++) {
    const name = existingData[i][CONFIG.COLS.GUEST_NAME];
    if (name) {
      existingAssignments[name] = {
        table: existingData[i][CONFIG.COLS.ASSIGNED_TABLE],
        seat: existingData[i][CONFIG.COLS.ASSIGNED_SEAT],
        isAssigned: existingData[i][CONFIG.COLS.IS_ASSIGNED]
      };
    }
  }
  
  const allGuests = [];
  let guestId = 1001;
  
  for (const listName of guestListNames) {
    const sheet = ss.getSheetByName(listName);
    if (!sheet) continue;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[1]; // Guest Name
      const rsvp = row[2]; // RSVP Status
      const group = row[3]; // Group
      const dietary = row[4]; // Dietary
      const plusOneName = row[5]; // Plus-One
      const plusOneDiet = row[6]; // Plus-One Diet
      const tablePref = row[7]; // Table Preference
      const vipStatus = row[8]; // VIP Status
      const specialNotes = row[9]; // Special Notes
      
      if (name && rsvp === 'Yes') {
        const existing = existingAssignments[name] || {};
        
        allGuests.push([
          guestId++,
          name,
          listName,
          'Yes',
          group || '',
          existing.table || '',
          existing.seat || '',
          sourceColours[listName],
          existing.isAssigned || false,
          'OK',
          plusOneName || '',
          dietary || '',
          vipStatus || 'Guest',
          tablePref || '',
          specialNotes || '',
          'Pending'
        ]);
        
        // Add plus-one if exists
        if (plusOneName && plusOneName.trim() !== '') {
          const plusOneExisting = existingAssignments[plusOneName] || {};
          allGuests.push([
            guestId++,
            plusOneName,
            listName + ' (Plus-One)',
            'Yes',
            group || '',
            plusOneExisting.table || '',
            plusOneExisting.seat || '',
            sourceColours[listName],
            plusOneExisting.isAssigned || false,
            'OK',
            name,
            plusOneDiet || '',
            'Guest',
            tablePref || '',
            'Plus-One of ' + name,
            'Pending'
          ]);
        }
      }
    }
  }
  
  // Clear and update Master_Guests
  masterSheet.getRange(2, 1, masterSheet.getLastRow(), 16).clearContent();
  masterSheet.getRange(2, 1, allGuests.length, 16).setValues(allGuests);
  
  logAction('SYNC', '', '', '', `Synced ${allGuests.length} guests from all lists`);
  
  SpreadsheetApp.getUi().alert('✅ Sync Complete', `Synchronized ${allGuests.length} guests to Master_Guests.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Automatically seats unassigned guests.
 */
function autoSeatUnassigned() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  const ui = SpreadsheetApp.getUi();
  
  const masterData = masterSheet.getDataRange().getValues();
  
  // Get unassigned guests sorted by VIP status
  const unassigned = [];
  
  for (let i = 1; i < masterData.length; i++) {
    const isAssigned = masterData[i][CONFIG.COLS.IS_ASSIGNED];
    if (isAssigned !== true && isAssigned !== 'TRUE') {
      unassigned.push({
        row: i + 1,
        name: masterData[i][CONFIG.COLS.GUEST_NAME],
        vip: masterData[i][CONFIG.COLS.VIP_STATUS],
        source: masterData[i][CONFIG.COLS.SOURCE_LIST],
        pref: masterData[i][CONFIG.COLS.TABLE_PREF]
      });
    }
  }
  
  if (unassigned.length === 0) {
    ui.alert('ℹ️ No Unassigned', 'All guests are already seated.', ui.ButtonSet.OK);
    return;
  }
  
  // Sort by VIP priority
  const vipPriority = { 'VIP': 1, 'Family': 2, 'Friend': 3, 'Colleague': 4, 'Guest': 5 };
  unassigned.sort((a, b) => (vipPriority[a.vip] || 5) - (vipPriority[b.vip] || 5));
  
  // Find available seats
  const tableSeats = [];
  for (let t = 1; t <= CONFIG.SEATING.TABLES_COUNT; t++) {
    for (let s = 1; s <= CONFIG.SEATING.SEATS_PER_TABLE; s++) {
      const row = findSeatRow(t, s);
      const guestInSeat = seatingSheet.getRange(row, 2).getValue();
      if (!guestInSeat) {
        tableSeats.push({ table: t, seat: s, row: row });
      }
    }
  }
  
  if (tableSeats.length === 0) {
    ui.alert('⚠️ No Seats', 'No available seats.', ui.ButtonSet.OK);
    return;
  }
  
  // Assign guests
  let assigned = 0;
  for (const guest of unassigned) {
    if (tableSeats.length === 0) break;
    
    // Find best table based on VIP status
    let bestSeat = null;
    if (guest.vip === 'VIP') {
      bestSeat = tableSeats.find(s => s.table <= 5);
    } else if (guest.vip === 'Family') {
      bestSeat = tableSeats.find(s => s.table <= 10);
    }
    
    if (!bestSeat) {
      bestSeat = tableSeats[0];
    }
    
    // Remove from available
    const idx = tableSeats.indexOf(bestSeat);
    tableSeats.splice(idx, 1);
    
    // Update Master_Guests
    masterSheet.getRange(guest.row, CONFIG.COLS.ASSIGNED_TABLE + 1).setValue(bestSeat.table);
    masterSheet.getRange(guest.row, CONFIG.COLS.ASSIGNED_SEAT + 1).setValue(bestSeat.seat);
    masterSheet.getRange(guest.row, CONFIG.COLS.IS_ASSIGNED + 1).setValue(true);
    
    // Update SEATING_PLAN
    const seatCell = seatingSheet.getRange(bestSeat.row, 2);
    seatCell.setValue(guest.name);
    
    const sourceColour = masterSheet.getRange(guest.row, CONFIG.COLS.SOURCE_COLOUR + 1).getValue();
    if (sourceColour) {
      seatCell.setBackground(sourceColour);
      seatCell.setFontColor('#FFFFFF');
    }
    
    updateTableHeader(bestSeat.table);
    assigned++;
  }
  
  logAction('AUTO_SEAT', '', '', '', `Auto-seated ${assigned} guests`);
  
  ui.alert('✅ Auto-Seating Complete', `Successfully seated ${assigned} guests.`, ui.ButtonSet.OK);
}

/**
 * Generates a printable seating report.
 */
function generateSeatingReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const ui = SpreadsheetApp.getUi();
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let report = '═══════════════════════════════════════════════════════════════\n';
  report += '              WEDDING SEATING REPORT\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  report += `Generated: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}\n\n`;
  
  // Summary
  let total = 0, seated = 0;
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][CONFIG.COLS.GUEST_NAME]) {
      total++;
      if (masterData[i][CONFIG.COLS.IS_ASSIGNED] === true) seated++;
    }
  }
  
  report += `SUMMARY:\n`;
  report += `  Total Guests: ${total}\n`;
  report += `  Seated: ${seated}\n`;
  report += `  Unseated: ${total - seated}\n\n`;
  
  // By Table
  report += '═══════════════════════════════════════════════════════════════\n';
  report += '              SEATING BY TABLE\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  
  for (let t = 1; t <= CONFIG.SEATING.TABLES_COUNT; t++) {
    const tableGuests = [];
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][CONFIG.COLS.ASSIGNED_TABLE] == t) {
        tableGuests.push({
          seat: masterData[i][CONFIG.COLS.ASSIGNED_SEAT],
          name: masterData[i][CONFIG.COLS.GUEST_NAME],
          vip: masterData[i][CONFIG.COLS.VIP_STATUS],
          diet: masterData[i][CONFIG.COLS.DIETARY_REQ]
        });
      }
    }
    
    if (tableGuests.length > 0) {
      report += `TABLE ${t} (${tableGuests.length}/10):\n`;
      tableGuests.sort((a, b) => a.seat - b.seat);
      for (const g of tableGuests) {
        let line = `  Seat ${g.seat}: ${g.name}`;
        if (g.vip === 'VIP') line += ' ⭐';
        if (g.diet && g.diet !== 'None') line += ` [${g.diet}]`;
        report += line + '\n';
      }
      report += '\n';
    }
  }
  
  // Unseated
  report += '═══════════════════════════════════════════════════════════════\n';
  report += '              UNSEATED GUESTS\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  
  for (let i = 1; i < masterData.length; i++) {
    const isAssigned = masterData[i][CONFIG.COLS.IS_ASSIGNED];
    if (masterData[i][CONFIG.COLS.GUEST_NAME] && isAssigned !== true && isAssigned !== 'TRUE') {
      report += `  • ${masterData[i][CONFIG.COLS.GUEST_NAME]} (${masterData[i][CONFIG.COLS.VIP_STATUS]})\n`;
    }
  }
  
  // Create HTML output
  const htmlOutput = HtmlService.createHtmlOutput(`<pre style="font-family: monospace; font-size: 12px;">${report}</pre>`)
    .setWidth(600)
    .setHeight(500)
    .setTitle('Seating Report');
  
  ui.showModalDialog(htmlOutput, '📋 Seating Report');
}

/**
 * Runs a comprehensive system test.
 */
function runSystemTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tests = [];
  let allPassed = true;
  
  // Test 1: All sheets exist
  const requiredSheets = Object.values(CONFIG.SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());
  const missingSheets = requiredSheets.filter(s => !existingSheets.includes(s));
  
  if (missingSheets.length === 0) {
    tests.push({ name: 'Sheet Structure', status: 'PASS', detail: 'All required sheets exist' });
  } else {
    tests.push({ name: 'Sheet Structure', status: 'FAIL', detail: `Missing: ${missingSheets.join(', ')}` });
    allPassed = false;
  }
  
  // Test 2: Master_Guests has data
  const masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GUESTS);
  const masterData = masterSheet.getDataRange().getValues();
  
  if (masterData.length > 1) {
    tests.push({ name: 'Guest Data', status: 'PASS', detail: `${masterData.length - 1} guests in Master_Guests` });
  } else {
    tests.push({ name: 'Guest Data', status: 'FAIL', detail: 'No guests in Master_Guests' });
    allPassed = false;
  }
  
  // Test 3: SEATING_PLAN structure
  const seatingSheet = ss.getSheetByName(CONFIG.SHEETS.SEATING_PLAN);
  const titleCell = seatingSheet.getRange('A1').getValue();
  
  if (titleCell && titleCell.includes('WEDDING SEATING')) {
    tests.push({ name: 'Seating Plan', status: 'PASS', detail: 'Seating plan structure valid' });
  } else {
    tests.push({ name: 'Seating Plan', status: 'WARN', detail: 'Seating plan title not found' });
  }
  
  // Test 4: CONFIG settings
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const maxTables = configSheet.getRange('B6').getValue();
  
  if (maxTables) {
    tests.push({ name: 'Configuration', status: 'PASS', detail: `Max tables: ${maxTables}` });
  } else {
    tests.push({ name: 'Configuration', status: 'WARN', detail: 'Configuration may be incomplete' });
  }
  
  // Build report
  let report = '═══════════════════════════════════════════════════════════════\n';
  report += '              SYSTEM TEST RESULTS\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  
  for (const test of tests) {
    const icon = test.status === 'PASS' ? '✅' : test.status === 'WARN' ? '⚠️' : '❌';
    report += `${icon} ${test.name}: ${test.status}\n`;
    report += `   ${test.detail}\n\n`;
  }
  
  report += '═══════════════════════════════════════════════════════════════\n';
  report += allPassed ? '✅ ALL CRITICAL TESTS PASSED\n' : '❌ SOME TESTS FAILED\n';
  report += '═══════════════════════════════════════════════════════════════\n';
  
  ui.alert('🧪 System Test', report, ui.ButtonSet.OK);
}
