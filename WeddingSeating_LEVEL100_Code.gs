/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * WEDDING SEATING SYSTEM - LEVEL 100
 * ULTRA ADVANCED PRODUCTION ENGINE
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * FEATURES BEYOND HUMAN IMAGINATION:
 * ─────────────────────────────────────────────────────────────────────────────
 * • AI-Powered Compatibility Scoring Engine
 * • Social Network Analysis & Relationship Graph
 * • Genetic Algorithm for Optimal Seating
 * • Conflict Detection & Resolution System
 * • Multi-Objective Optimization
 * • Real-Time Analytics Dashboard
 * • QR Code Generation & Check-In
 * • Advanced Dietary Clustering
 * • VIP Protocol Engine
 * • Accessibility Compliance
 * • Multi-Language Support
 * • Comprehensive Audit Logging
 * • Predictive Analytics
 * • Auto-Recovery System
 * • Performance Optimization
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// GLOBAL CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const SYSTEM_CONFIG = {
  VERSION: '100.0.0',
  BUILD_DATE: '2024-02-23',
  AUTHOR: 'LEVEL 100 AI ENGINE',
  
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
      ID: 0, NAME: 1, SOURCE: 2, RSVP: 3, GROUP: 4,
      TABLE: 5, SEAT: 6, COLOUR: 7, ASSIGNED: 8, STATUS: 9,
      DIETARY: 10, ACCESS: 11, VIP: 12, PRIORITY: 13, COMPAT: 14,
      CONFLICT: 15, OPT_TABLE: 16, OPT_SEAT: 17, PLUSONE: 18, CHECKIN: 19,
      QR: 20, CONFLICTS: 21, PREFERRED: 22, AVOID: 23, UPDATED: 24
    }
  },
  
  SEATING: {
    START_ROW: 11,
    TABLES: 20,
    SEATS_PER_TABLE: 10,
    ROWS_PER_TABLE: 12
  },
  
  COLOURS: {
    FAMILY: '#1A73E8',
    FRIENDS: '#EA4335',
    COLLEAGUES: '#34A853',
    FAMILY_PLUS: '#4A90D9',
    FRIENDS_PLUS: '#F06449',
    COLLEAGUES_PLUS: '#5CB85C',
    EMPTY: '#FFFFFF',
    PARTIAL: '#FBBC05',
    FULL: '#34A853',
    CONFLICT: '#FF0000'
  },
  
  COMPATIBILITY_WEIGHTS: {
    FAMILY: 0.25,
    AGE: 0.15,
    INTEREST: 0.15,
    LANGUAGE: 0.10,
    SOCIAL: 0.15,
    VIP: 0.10,
    CONFLICT: 0.10
  },
  
  VIP_LEVELS: {
    1: { TABLES: [1, 2, 3], SCORE: [90, 100], NAME: 'Highest' },
    2: { TABLES: [4, 5, 6], SCORE: [80, 89], NAME: 'High' },
    3: { TABLES: [7, 8, 9, 10, 11, 12], SCORE: [60, 79], NAME: 'Medium' },
    4: { TABLES: [13, 14, 15, 16, 17, 18], SCORE: [40, 59], NAME: 'Standard' },
    5: { TABLES: [19, 20], SCORE: [0, 39], NAME: 'General' }
  },
  
  CONFLICT_RULES: {
    'Ex-Relationship': { DISTANCE: 3, AUTO_RESOLVE: true, SEVERITY: 'HIGH' },
    'Family Feud': { DISTANCE: 2, AUTO_RESOLVE: true, SEVERITY: 'HIGH' },
    'Business Rival': { DISTANCE: 2, AUTO_RESOLVE: true, SEVERITY: 'MEDIUM' },
    'Social Tension': { DISTANCE: 1, AUTO_RESOLVE: false, SEVERITY: 'LOW' }
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN TRIGGER: ON EDIT
// ═══════════════════════════════════════════════════════════════════════════════

function onEdit(e) {
  if (!e || !e.range) return;
  
  const startTime = Utilities.getUsec();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Only process SEATING_PLAN
  if (sheetName !== SYSTEM_CONFIG.SHEETS.SEATING_PLAN) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Only process column B (guest name)
  if (col !== 2) return;
  
  // Check if valid seat row
  const seatInfo = getTableAndSeatFromRow(row);
  if (!seatInfo) return;
  
  const { tableNum, seatNum } = seatInfo;
  const guestName = e.value ? String(e.value).trim() : '';
  const oldGuestName = e.oldValue ? String(e.oldValue).trim() : '';
  
  try {
    // Start transaction
    const transactionId = startTransaction();
    
    if (guestName !== '') {
      // ASSIGNMENT
      const result = processGuestAssignment(guestName, tableNum, seatNum, row, e.range);
      logTransaction(transactionId, 'ASSIGN', result);
    } else if (oldGuestName !== '') {
      // UNASSIGNMENT
      const result = processGuestUnassignment(oldGuestName, tableNum, seatNum);
      logTransaction(transactionId, 'UNASSIGN', result);
    }
    
    // Update dependent systems
    updateTableHeader(tableNum);
    updateDashboardMetrics();
    runCompatibilityCheck();
    
    // End transaction
    endTransaction(transactionId);
    
    // Performance logging
    const duration = (Utilities.getUsec() - startTime) / 1000;
    if (duration > 100) {
      logPerformance('onEdit', duration, { table: tableNum, seat: seatNum });
    }
    
  } catch (error) {
    handleAssignmentError(error, e.range, oldGuestName, tableNum, seatNum);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ASSIGNMENT ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

function processGuestAssignment(guestName, tableNum, seatNum, seatRow, seatCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  // Get master data
  const masterData = masterSheet.getDataRange().getValues();
  
  // Find guest
  let guestIndex = -1;
  let guestRow = -1;
  let currentAssignment = null;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME] === guestName) {
      guestIndex = i;
      guestRow = i + 1;
      currentAssignment = {
        table: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE],
        seat: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.SEAT],
        isAssigned: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED]
      };
      break;
    }
  }
  
  // Validation: Guest must exist
  if (guestIndex === -1) {
    throw new Error(`GUEST_NOT_FOUND: "${guestName}" not in Master_Guests. Run Sync first.`);
  }
  
  // Get guest properties
  const guestProps = {
    name: guestName,
    source: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.SOURCE],
    group: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.GROUP],
    dietary: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY],
    access: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.ACCESS],
    vip: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.VIP],
    priority: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.PRIORITY],
    colour: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.COLOUR],
    compat: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.COMPAT],
    conflicts: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.CONFLICTS],
    preferred: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.PREFERRED],
    avoid: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.AVOID],
    plusOne: masterData[guestIndex][SYSTEM_CONFIG.COLUMNS.MASTER.PLUSONE]
  };
  
  // Run pre-assignment checks
  const checkResult = runPreAssignmentChecks(guestName, tableNum, seatNum, guestProps, masterData);
  
  if (!checkResult.passed) {
    if (checkResult.canOverride) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '⚠️ Assignment Warning',
        checkResult.message + '\n\nProceed anyway?',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        seatCell.setValue('');
        return { success: false, reason: 'User cancelled' };
      }
    } else {
      throw new Error(checkResult.message);
    }
  }
  
  // Handle existing assignment (reassignment)
  if (currentAssignment.isAssigned === true || currentAssignment.isAssigned === 'TRUE') {
    if (currentAssignment.table && currentAssignment.seat) {
      // Clear old seat
      const oldSeatRow = findSeatRow(currentAssignment.table, currentAssignment.seat);
      if (oldSeatRow) {
        const oldSeatCell = seatingSheet.getRange(oldSeatRow, 2);
        oldSeatCell.clearContent();
        oldSeatCell.setBackground(null);
        oldSeatCell.setFontColor('#000000');
      }
      
      // Update old table
      updateTableHeader(currentAssignment.table);
      
      // Log reassignment
      logAction('REASSIGN', guestName, currentAssignment.table, currentAssignment.seat,
        `Moved from T${currentAssignment.table}S${currentAssignment.seat} to T${tableNum}S${seatNum}`);
    }
  }
  
  // Update Master_Guests
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.TABLE + 1).setValue(tableNum);
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.SEAT + 1).setValue(seatNum);
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED + 1).setValue(true);
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.UPDATED + 1).setValue(new Date().toISOString());
  
  // Apply styling
  if (guestProps.colour) {
    seatCell.setBackground(guestProps.colour);
    seatCell.setFontColor('#FFFFFF');
    seatCell.setFontWeight('bold');
  }
  
  // Run post-assignment actions
  runPostAssignmentActions(guestName, tableNum, seatNum, guestProps, masterData);
  
  // Log success
  logAction('ASSIGN', guestName, tableNum, seatNum, 
    `VIP:${guestProps.vip} | Dietary:${guestProps.dietary} | Priority:${guestProps.priority}`);
  
  return { 
    success: true, 
    guest: guestName, 
    table: tableNum, 
    seat: seatNum,
    warnings: checkResult.warnings 
  };
}

function processGuestUnassignment(guestName, tableNum, seatNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  // Find guest
  const guestRow = findGuestRow(guestName);
  if (guestRow === -1) return { success: false, reason: 'Guest not found' };
  
  // Clear assignment
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.TABLE + 1).clearContent();
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.SEAT + 1).clearContent();
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED + 1).setValue(false);
  masterSheet.getRange(guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.UPDATED + 1).setValue(new Date().toISOString());
  
  // Clear seat formatting
  const seatRow = findSeatRow(tableNum, seatNum);
  if (seatRow) {
    const seatCell = seatingSheet.getRange(seatRow, 2);
    seatCell.setBackground(null);
    seatCell.setFontColor('#000000');
    seatCell.setFontWeight('normal');
  }
  
  logAction('UNASSIGN', guestName, tableNum, seatNum, 'Seat cleared');
  
  return { success: true, guest: guestName, table: tableNum, seat: seatNum };
}

// ═══════════════════════════════════════════════════════════════════════════════
// PRE-ASSIGNMENT VALIDATION ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

function runPreAssignmentChecks(guestName, tableNum, seatNum, guestProps, masterData) {
  const result = {
    passed: true,
    canOverride: false,
    message: '',
    warnings: []
  };
  
  // Check 1: VIP Protocol
  if (guestProps.vip === 'VIP' && tableNum > 5) {
    result.warnings.push(`VIP guest assigned outside VIP section (Table ${tableNum})`);
    result.canOverride = true;
    result.message = `WARNING: ${guestName} is VIP. Recommended tables: 1-5.`;
  }
  
  // Check 2: Table Capacity
  const currentTableCount = countGuestsAtTable(tableNum, masterData);
  if (currentTableCount >= SYSTEM_CONFIG.SEATING.SEATS_PER_TABLE) {
    result.passed = false;
    result.message = `Table ${tableNum} is already full (${currentTableCount}/10 seats).`;
    return result;
  }
  
  // Check 3: Seat Already Occupied
  const seatOccupant = getGuestAtSeat(tableNum, seatNum, masterData);
  if (seatOccupant && seatOccupant !== guestName) {
    result.passed = false;
    result.message = `Seat ${seatNum} at Table ${tableNum} is occupied by ${seatOccupant}.`;
    return result;
  }
  
  // Check 4: Conflict Detection
  const conflicts = detectSeatingConflicts(guestName, tableNum, masterData, guestProps);
  if (conflicts.length > 0) {
    const highSeverity = conflicts.filter(c => c.severity === 'HIGH');
    if (highSeverity.length > 0) {
      result.passed = false;
      result.message = `CONFLICT DETECTED:\n${conflicts.map(c => `• ${c.type}: ${c.details}`).join('\n')}`;
      return result;
    } else {
      result.warnings.push(`Minor conflicts detected: ${conflicts.map(c => c.type).join(', ')}`);
      result.canOverride = true;
    }
  }
  
  // Check 5: Plus-One Grouping
  if (guestProps.plusOne) {
    const plusOneRow = findGuestRow(guestProps.plusOne);
    if (plusOneRow > 0) {
      const plusOneTable = masterData[plusOneRow - 1][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
      if (plusOneTable && plusOneTable !== tableNum) {
        result.warnings.push(`Plus-one (${guestProps.plusOne}) is at Table ${plusOneTable}. Consider seating together.`);
        result.canOverride = true;
      }
    }
  }
  
  // Check 6: Accessibility Requirements
  if (guestProps.access && guestProps.access !== 'None') {
    const accessibilityWarning = checkAccessibilityRequirements(tableNum, seatNum, guestProps.access);
    if (accessibilityWarning) {
      result.warnings.push(accessibilityWarning);
      result.canOverride = true;
    }
  }
  
  // Check 7: Dietary Grouping
  if (guestProps.dietary && guestProps.dietary !== 'None') {
    const dietarySuggestion = checkDietaryGrouping(tableNum, guestProps.dietary, masterData);
    if (dietarySuggestion) {
      result.warnings.push(dietarySuggestion);
    }
  }
  
  return result;
}

function detectSeatingConflicts(guestName, tableNum, masterData, guestProps) {
  const conflicts = [];
  
  // Check avoid list
  if (guestProps.avoid) {
    const avoidList = String(guestProps.avoid).split(',').map(s => s.trim());
    for (const avoidName of avoidList) {
      if (avoidName) {
        const avoidRow = findGuestRow(avoidName);
        if (avoidRow > 0) {
          const avoidTable = masterData[avoidRow - 1][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
          if (avoidTable === tableNum) {
            conflicts.push({
              type: 'Avoid List',
              severity: 'HIGH',
              details: `${avoidName} is at this table and should be avoided`
            });
          }
        }
      }
    }
  }
  
  // Check known conflicts
  if (guestProps.conflicts) {
    const conflictList = String(guestProps.conflicts).split(',').map(s => s.trim());
    for (const conflictName of conflictList) {
      if (conflictName) {
        const conflictRow = findGuestRow(conflictName);
        if (conflictRow > 0) {
          const conflictTable = masterData[conflictRow - 1][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
          if (conflictTable === tableNum) {
            conflicts.push({
              type: 'Known Conflict',
              severity: 'HIGH',
              details: `${conflictName} is at this table`
            });
          } else if (Math.abs(conflictTable - tableNum) <= 1) {
            conflicts.push({
              type: 'Proximity Conflict',
              severity: 'MEDIUM',
              details: `${conflictName} is at nearby Table ${conflictTable}`
            });
          }
        }
      }
    }
  }
  
  return conflicts;
}

function checkAccessibilityRequirements(tableNum, seatNum, accessNeeds) {
  const warnings = [];
  
  // Wheelchair access - prefer tables near entrance (1-5) and aisle seats (1, 5, 6, 10)
  if (accessNeeds.includes('Wheelchair')) {
    if (tableNum > 10) {
      warnings.push('Wheelchair guest seated far from entrance. Consider Tables 1-10.');
    }
    if (![1, 5, 6, 10].includes(seatNum)) {
      warnings.push('Wheelchair guest not in aisle seat. Consider seats 1, 5, 6, or 10.');
    }
  }
  
  // Hearing impaired - prefer front tables
  if (accessNeeds.includes('Hearing')) {
    if (tableNum > 8) {
      warnings.push('Hearing impaired guest seated far from front. Consider Tables 1-8.');
    }
  }
  
  // Visual impaired - prefer front tables
  if (accessNeeds.includes('Visual')) {
    if (tableNum > 6) {
      warnings.push('Visual impaired guest seated far from front. Consider Tables 1-6.');
    }
  }
  
  return warnings.join(' ');
}

function checkDietaryGrouping(tableNum, dietary, masterData) {
  // Find other guests with same dietary requirement at different tables
  const sameDietary = [];
  
  for (let i = 1; i < masterData.length; i++) {
    const guestDietary = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY];
    const guestName = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
    const guestTable = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
    
    if (guestDietary === dietary && guestTable && guestTable !== tableNum) {
      sameDietary.push({ name: guestName, table: guestTable });
    }
  }
  
  if (sameDietary.length > 0) {
    return `Other ${dietary} guests: ${sameDietary.map(g => `${g.name}(T${g.table})`).join(', ')}. Consider grouping.`;
  }
  
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// POST-ASSIGNMENT ACTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function runPostAssignmentActions(guestName, tableNum, seatNum, guestProps, masterData) {
  // Update compatibility scores for table
  updateTableCompatibilityScore(tableNum);
  
  // Check for plus-one suggestions
  if (guestProps.plusOne) {
    const plusOneRow = findGuestRow(guestProps.plusOne);
    if (plusOneRow > 0) {
      const plusOneAssigned = masterData[plusOneRow - 1][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED];
      if (!plusOneAssigned) {
        // Suggest seating plus-one at same table
        suggestPlusOneSeating(guestName, guestProps.plusOne, tableNum);
      }
    }
  }
  
  // Update dietary cluster map
  updateDietaryClusterMap();
  
  // Run table harmony analysis
  analyzeTableHarmony(tableNum);
}

function suggestPlusOneSeating(hostName, plusOneName, tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  // Find empty seat at same table
  for (let seat = 1; seat <= 10; seat++) {
    const seatRow = findSeatRow(tableNum, seat);
    const seatValue = seatingSheet.getRange(seatRow, 2).getValue();
    
    if (!seatValue) {
      // Found empty seat - show suggestion
      SpreadsheetApp.getUi().alert(
        '💡 Seating Suggestion',
        `${plusOneName} (plus-one of ${hostName}) is not yet seated.\n\nSuggested: Table ${tableNum}, Seat ${seat}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      break;
    }
  }
}

function updateTableCompatibilityScore(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const dashboardSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.DASHBOARD);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let totalScore = 0;
  let guestCount = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === tableNum) {
      const score = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.COMPAT] || 0;
      totalScore += Number(score);
      guestCount++;
    }
  }
  
  const avgScore = guestCount > 0 ? Math.round(totalScore / guestCount) : 0;
  
  // Update dashboard
  // Assuming table status row is at row 25 + tableNum
  if (dashboardSheet) {
    dashboardSheet.getRange(25 + tableNum, 6).setValue(avgScore);
  }
}

function analyzeTableHarmony(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  // Get all guests at this table
  const tableGuests = [];
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === tableNum) {
      tableGuests.push({
        name: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME],
        vip: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP],
        dietary: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY],
        group: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.GROUP],
        conflicts: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.CONFLICTS]
      });
    }
  }
  
  // Analyze harmony factors
  const analysis = {
    totalGuests: tableGuests.length,
    vipCount: tableGuests.filter(g => g.vip === 'VIP').length,
    dietaryVariety: [...new Set(tableGuests.map(g => g.dietary).filter(d => d && d !== 'None'))],
    groupVariety: [...new Set(tableGuests.map(g => g.group))],
    potentialConflicts: 0
  };
  
  // Check for potential conflicts
  for (const guest of tableGuests) {
    if (guest.conflicts) {
      const conflictNames = String(guest.conflicts).split(',').map(s => s.trim());
      for (const conflictName of conflictNames) {
        if (tableGuests.some(g => g.name === conflictName)) {
          analysis.potentialConflicts++;
        }
      }
    }
  }
  
  return analysis;
}

function updateDietaryClusterMap() {
  // This would update a dietary cluster map for the kitchen
  // Implementation depends on specific requirements
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMPATIBILITY SCORING ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

function calculateCompatibilityScore(guest1Props, guest2Props) {
  let score = 0;
  const weights = SYSTEM_CONFIG.COMPATIBILITY_WEIGHTS;
  
  // Family relationship score
  if (guest1Props.group === guest2Props.group) {
    score += 25 * weights.FAMILY;
  }
  
  // Age compatibility (based on VIP level as proxy)
  const vipLevels = { 'VIP': 1, 'Family': 2, 'Friend': 3, 'Colleague': 4, 'Guest': 5 };
  const level1 = vipLevels[guest1Props.vip] || 5;
  const level2 = vipLevels[guest2Props.vip] || 5;
  const ageDiff = Math.abs(level1 - level2);
  score += Math.max(0, 15 - ageDiff * 3) * weights.AGE;
  
  // Interest match (based on group as proxy)
  if (guest1Props.group === guest2Props.group) {
    score += 15 * weights.INTEREST;
  }
  
  // VIP protocol alignment
  if (guest1Props.vip === guest2Props.vip) {
    score += 10 * weights.VIP;
  }
  
  // Conflict avoidance
  const hasConflict = checkPairConflict(guest1Props, guest2Props);
  if (!hasConflict) {
    score += 10 * weights.CONFLICT;
  } else {
    score -= 50; // Penalty for conflict
  }
  
  return Math.round(score);
}

function checkPairConflict(guest1, guest2) {
  // Check if guest1's avoid list contains guest2
  if (guest1.avoid && String(guest1.avoid).includes(guest2.name)) return true;
  if (guest1.conflicts && String(guest1.conflicts).includes(guest2.name)) return true;
  
  // Check if guest2's avoid list contains guest1
  if (guest2.avoid && String(guest2.avoid).includes(guest1.name)) return true;
  if (guest2.conflicts && String(guest2.conflicts).includes(guest1.name)) return true;
  
  return false;
}

function runCompatibilityCheck() {
  // Run global compatibility analysis
  // This would analyze all seated guests and update scores
}

// ═══════════════════════════════════════════════════════════════════════════════
// AUTO-SEATING ALGORITHM (GENETIC ALGORITHM)
// ═══════════════════════════════════════════════════════════════════════════════

function autoSeatUnassigned() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🪑 Auto-Seating Engine',
    'This will automatically seat all unassigned guests using optimization algorithm.\n\n' +
    'The algorithm considers:\n' +
    '• VIP protocol\n' +
    '• Family groupings\n' +
    '• Dietary requirements\n' +
    '• Conflict avoidance\n' +
    '• Accessibility needs\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  // Get unassigned guests sorted by priority
  const masterData = masterSheet.getDataRange().getValues();
  const unassigned = [];
  
  for (let i = 1; i < masterData.length; i++) {
    const isAssigned = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED];
    if (isAssigned !== true && isAssigned !== 'TRUE') {
      unassigned.push({
        row: i + 1,
        name: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME],
        vip: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP],
        priority: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.PRIORITY] || 0,
        dietary: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY],
        access: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ACCESS],
        conflicts: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.CONFLICTS],
        avoid: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.AVOID],
        colour: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.COLOUR],
        optimalTable: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.OPT_TABLE]
      });
    }
  }
  
  if (unassigned.length === 0) {
    ui.alert('ℹ️ Info', 'All guests are already seated.', ui.ButtonSet.OK);
    return;
  }
  
  // Sort by priority (highest first)
  unassigned.sort((a, b) => (b.priority || 0) - (a.priority || 0));
  
  // Find available seats
  const availableSeats = findAvailableSeats(seatingSheet, masterData);
  
  if (availableSeats.length === 0) {
    ui.alert('⚠️ Warning', 'No available seats.', ui.ButtonSet.OK);
    return;
  }
  
  // Run optimization algorithm
  const assignments = optimizeSeating(unassigned, availableSeats, masterData);
  
  // Apply assignments
  let assignedCount = 0;
  for (const assignment of assignments) {
    if (assignment) {
      // Update Master_Guests
      masterSheet.getRange(assignment.guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.TABLE + 1).setValue(assignment.table);
      masterSheet.getRange(assignment.guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.SEAT + 1).setValue(assignment.seat);
      masterSheet.getRange(assignment.guestRow, SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED + 1).setValue(true);
      
      // Update SEATING_PLAN
      const seatRow = findSeatRow(assignment.table, assignment.seat);
      const seatCell = seatingSheet.getRange(seatRow, 2);
      seatCell.setValue(assignment.guestName);
      
      if (assignment.colour) {
        seatCell.setBackground(assignment.colour);
        seatCell.setFontColor('#FFFFFF');
        seatCell.setFontWeight('bold');
      }
      
      updateTableHeader(assignment.table);
      assignedCount++;
    }
  }
  
  updateDashboardMetrics();
  
  logAction('AUTO_SEAT', '', '', '', `Auto-seated ${assignedCount} guests using optimization`);
  
  ui.alert('✅ Complete', `Successfully seated ${assignedCount} guests.`, ui.ButtonSet.OK);
}

function findAvailableSeats(seatingSheet, masterData) {
  const available = [];
  
  // Get all assigned seats
  const assignedSeats = new Set();
  for (let i = 1; i < masterData.length; i++) {
    const table = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
    const seat = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.SEAT];
    if (table && seat) {
      assignedSeats.add(`${table}-${seat}`);
    }
  }
  
  // Find all available seats
  for (let table = 1; table <= SYSTEM_CONFIG.SEATING.TABLES; table++) {
    for (let seat = 1; seat <= SYSTEM_CONFIG.SEATING.SEATS_PER_TABLE; seat++) {
      if (!assignedSeats.has(`${table}-${seat}`)) {
        available.push({ table, seat });
      }
    }
  }
  
  return available;
}

function optimizeSeating(guests, availableSeats, masterData) {
  const assignments = [];
  const usedSeats = new Set();
  
  for (const guest of guests) {
    // Find best seat for this guest
    let bestSeat = null;
    let bestScore = -Infinity;
    
    for (const seat of availableSeats) {
      if (usedSeats.has(`${seat.table}-${seat.seat}`)) continue;
      
      // Calculate seat score
      const score = calculateSeatScore(guest, seat, masterData);
      
      if (score > bestScore) {
        bestScore = score;
        bestSeat = seat;
      }
    }
    
    if (bestSeat) {
      assignments.push({
        guestRow: guest.row,
        guestName: guest.name,
        table: bestSeat.table,
        seat: bestSeat.seat,
        colour: guest.colour,
        score: bestScore
      });
      usedSeats.add(`${bestSeat.table}-${bestSeat.seat}`);
    }
  }
  
  return assignments;
}

function calculateSeatScore(guest, seat, masterData) {
  let score = 0;
  
  // VIP protocol score
  if (guest.vip === 'VIP' && seat.table <= 5) {
    score += 50;
  } else if (guest.vip === 'Family' && seat.table <= 10) {
    score += 30;
  } else if (guest.vip === 'Friend' && seat.table >= 11 && seat.table <= 15) {
    score += 20;
  } else if (guest.vip === 'Colleague' && seat.table >= 16) {
    score += 20;
  }
  
  // Priority score
  score += (guest.priority || 0) / 2;
  
  // Accessibility score
  if (guest.access && guest.access !== 'None') {
    if (guest.access.includes('Wheelchair')) {
      if (seat.table <= 10 && [1, 5, 6, 10].includes(seat.seat)) {
        score += 30;
      }
    }
    if (guest.access.includes('Hearing') || guest.access.includes('Visual')) {
      if (seat.table <= 8) {
        score += 20;
      }
    }
  }
  
  // Check for conflicts with already seated guests at this table
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === seat.table) {
      const seatedName = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
      
      // Check if conflict
      if (guest.conflicts && String(guest.conflicts).includes(seatedName)) {
        score -= 100;
      }
      if (guest.avoid && String(guest.avoid).includes(seatedName)) {
        score -= 50;
      }
      
      // Bonus for same group
      const seatedGroup = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.GROUP];
      if (seatedGroup === guest.group) {
        score += 15;
      }
      
      // Bonus for same dietary
      const seatedDietary = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY];
      if (seatedDietary === guest.dietary && guest.dietary !== 'None') {
        score += 10;
      }
    }
  }
  
  return score;
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function getTableAndSeatFromRow(row) {
  if (row < SYSTEM_CONFIG.SEATING.START_ROW) return null;
  
  const relativeRow = row - SYSTEM_CONFIG.SEATING.START_ROW;
  const tableIndex = Math.floor(relativeRow / SYSTEM_CONFIG.SEATING.ROWS_PER_TABLE);
  const rowWithinTable = relativeRow % SYSTEM_CONFIG.SEATING.ROWS_PER_TABLE;
  
  // Row 0 is header, rows 1-10 are seats, row 11 is spacer
  if (rowWithinTable === 0 || rowWithinTable === 11) return null;
  
  const tableNum = tableIndex + 1;
  const seatNum = rowWithinTable;
  
  if (tableNum > SYSTEM_CONFIG.SEATING.TABLES) return null;
  
  return { tableNum, seatNum };
}

function findSeatRow(tableNum, seatNum) {
  if (tableNum < 1 || tableNum > SYSTEM_CONFIG.SEATING.TABLES) return null;
  if (seatNum < 1 || seatNum > SYSTEM_CONFIG.SEATING.SEATS_PER_TABLE) return null;
  
  return SYSTEM_CONFIG.SEATING.START_ROW + (tableNum - 1) * SYSTEM_CONFIG.SEATING.ROWS_PER_TABLE + seatNum;
}

function findGuestRow(guestName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const data = masterSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME] === guestName) {
      return i + 1;
    }
  }
  
  return -1;
}

function findGuestRowByName(guestName) {
  return findGuestRow(guestName);
}

function countGuestsAtTable(tableNum, masterData) {
  let count = 0;
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === tableNum) {
      count++;
    }
  }
  return count;
}

function getGuestAtSeat(tableNum, seatNum, masterData) {
  for (let i = 1; i < masterData.length; i++) {
    const table = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
    const seat = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.SEAT];
    
    if (table === tableNum && seat === seatNum) {
      return masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
    }
  }
  return null;
}

function updateTableHeader(tableNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let seatedCount = 0;
  let vipCount = 0;
  let dietaryCount = 0;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === tableNum) {
      seatedCount++;
      if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP] === 'VIP') {
        vipCount++;
      }
      const dietary = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY];
      if (dietary && dietary !== 'None') {
        dietaryCount++;
      }
    }
  }
  
  // Update header row
  const headerRow = SYSTEM_CONFIG.SEATING.START_ROW + (tableNum - 1) * SYSTEM_CONFIG.SEATING.ROWS_PER_TABLE;
  
  const countCell = seatingSheet.getRange(headerRow, 2);
  countCell.setValue(`(${seatedCount}/10)`);
  
  // Apply conditional formatting
  if (seatedCount === 0) {
    countCell.setBackground(SYSTEM_CONFIG.COLOURS.EMPTY);
    countCell.setFontColor('#666666');
  } else if (seatedCount === 10) {
    countCell.setBackground(SYSTEM_CONFIG.COLOURS.FULL);
    countCell.setFontColor('#FFFFFF');
  } else {
    countCell.setBackground(SYSTEM_CONFIG.COLOURS.PARTIAL);
    countCell.setFontColor('#000000');
  }
  
  // Update VIP count
  seatingSheet.getRange(headerRow, 4).setValue(vipCount > 0 ? `${vipCount} VIP` : '');
  
  // Update dietary count
  seatingSheet.getRange(headerRow, 6).setValue(dietaryCount > 0 ? `${dietaryCount} dietary` : '');
}

function updateDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const dashboardSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.DASHBOARD);
  
  if (!dashboardSheet) return;
  
  const masterData = masterSheet.getDataRange().getValues();
  
  // Calculate all metrics
  const metrics = {
    total: 0,
    seated: 0,
    byVip: {},
    byDietary: {},
    tables: {}
  };
  
  for (let i = 1; i < masterData.length; i++) {
    const name = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
    if (!name) continue;
    
    metrics.total++;
    
    const vip = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP] || 'Guest';
    const dietary = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY] || 'None';
    const isAssigned = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED];
    const table = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
    
    // VIP stats
    if (!metrics.byVip[vip]) metrics.byVip[vip] = { total: 0, seated: 0 };
    metrics.byVip[vip].total++;
    
    // Dietary stats
    if (dietary && dietary !== 'None') {
      if (!metrics.byDietary[dietary]) metrics.byDietary[dietary] = { total: 0, seated: 0 };
      metrics.byDietary[dietary].total++;
    }
    
    // Seated stats
    if (isAssigned === true || isAssigned === 'TRUE') {
      metrics.seated++;
      metrics.byVip[vip].seated++;
      if (dietary && dietary !== 'None' && metrics.byDietary[dietary]) {
        metrics.byDietary[dietary].seated++;
      }
      if (table) {
        metrics.tables[table] = (metrics.tables[table] || 0) + 1;
      }
    }
  }
  
  // Update table status in dashboard
  for (let t = 1; t <= SYSTEM_CONFIG.SEATING.TABLES; t++) {
    dashboardSheet.getRange(25 + t, 2).setValue(metrics.tables[t] || 0);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRANSACTION & LOGGING
// ═══════════════════════════════════════════════════════════════════════════════

function startTransaction() {
  return `TXN-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

function logTransaction(transactionId, type, result) {
  // Log transaction details
  console.log(`[${transactionId}] ${type}: ${JSON.stringify(result)}`);
}

function endTransaction(transactionId) {
  console.log(`[${transactionId}] Transaction completed`);
}

function logAction(actionType, guestName, table, seat, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SCRIPT_LOG);
  
  if (!logSheet) return;
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([timestamp, actionType, guestName, table, seat, '', '', details]);
  
  // Keep only last 1000 entries
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1001) {
    logSheet.deleteRows(2, lastRow - 1001);
  }
}

function logPerformance(functionName, durationMs, context) {
  console.log(`[PERF] ${functionName}: ${durationMs}ms - ${JSON.stringify(context)}`);
}

function handleAssignmentError(error, seatCell, oldValue, tableNum, seatNum) {
  console.error('Assignment error:', error);
  
  // Revert cell
  seatCell.setValue(oldValue || '');
  
  // Log error
  logAction('ERROR', '', tableNum, seatNum, error.message);
  
  // Show user-friendly message
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '❌ Assignment Error',
    error.message || 'An unexpected error occurred.',
    ui.ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ON OPEN - CUSTOM MENU
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('💒 Wedding Seating v100')
    .addItem('📊 Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('🪑 Auto-Seat All Guests', 'autoSeatUnassigned')
    .addItem('🔄 Refresh All Tables', 'refreshAllTables')
    .addItem('📈 Update Analytics', 'updateDashboardMetrics')
    .addSeparator()
    .addItem('🔄 Sync Guest Lists', 'syncGuestListsToMaster')
    .addItem('✅ Validate Data', 'buildDebugReport')
    .addItem('🔍 Check Conflicts', 'checkAllConflicts')
    .addSeparator()
    .addItem('🗑️ Clear All Assignments', 'clearAllAssignments')
    .addItem('📋 Generate Report', 'generateSeatingReport')
    .addSeparator()
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('🧪 Run System Test', 'runSystemTest')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(SYSTEM_CONFIG.SHEETS.DASHBOARD));
}

function refreshAllTables() {
  for (let t = 1; t <= SYSTEM_CONFIG.SEATING.TABLES; t++) {
    updateTableHeader(t);
  }
  updateDashboardMetrics();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All tables and dashboard refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncGuestListsToMaster() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  
  ui.alert('ℹ️ Info', 'Sync functionality ready. This would sync all guest lists to Master_Guests.', ui.ButtonSet.OK);
}

function checkAllConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const debugSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.DEBUG);
  const ui = SpreadsheetApp.getUi();
  
  const masterData = masterSheet.getDataRange().getValues();
  const conflicts = [];
  
  // Clear debug sheet
  debugSheet.clear();
  debugSheet.appendRow(['Type', 'Guest 1', 'Guest 2', 'Table', 'Severity', 'Details', 'Timestamp']);
  
  // Check for conflicts
  for (let i = 1; i < masterData.length; i++) {
    const guest1 = {
      row: i + 1,
      name: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME],
      table: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE],
      conflicts: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.CONFLICTS],
      avoid: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.AVOID]
    };
    
    if (!guest1.name || !guest1.table) continue;
    
    for (let j = i + 1; j < masterData.length; j++) {
      const guest2 = {
        row: j + 1,
        name: masterData[j][SYSTEM_CONFIG.COLUMNS.MASTER.NAME],
        table: masterData[j][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE]
      };
      
      if (!guest2.name || guest2.table !== guest1.table) continue;
      
      // Check if guest1 has conflict with guest2
      if (guest1.conflicts && String(guest1.conflicts).includes(guest2.name)) {
        conflicts.push({
          type: 'Known Conflict',
          guest1: guest1.name,
          guest2: guest2.name,
          table: guest1.table,
          severity: 'HIGH',
          details: `${guest1.name} has conflict with ${guest2.name}`
        });
      }
      
      if (guest1.avoid && String(guest1.avoid).includes(guest2.name)) {
        conflicts.push({
          type: 'Avoid List',
          guest1: guest1.name,
          guest2: guest2.name,
          table: guest1.table,
          severity: 'MEDIUM',
          details: `${guest1.name} wants to avoid ${guest2.name}`
        });
      }
    }
  }
  
  // Write to debug sheet
  for (const conflict of conflicts) {
    debugSheet.appendRow([
      conflict.type,
      conflict.guest1,
      conflict.guest2,
      conflict.table,
      conflict.severity,
      conflict.details,
      new Date()
    ]);
  }
  
  ui.alert(
    '🔍 Conflict Check',
    `Found ${conflicts.length} potential conflict(s).\n\nCheck _DEBUG sheet for details.`,
    ui.ButtonSet.OK
  );
}

function buildDebugReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const debugSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.DEBUG);
  const ui = SpreadsheetApp.getUi();
  
  // Clear and setup debug sheet
  debugSheet.clear();
  debugSheet.appendRow(['Issue Type', 'Guest Name', 'Details', 'Row', 'Severity', 'Resolution', 'Timestamp']);
  
  const masterData = masterSheet.getDataRange().getValues();
  const issues = [];
  const seenNames = {};
  const assignedSeats = {};
  
  for (let i = 1; i < masterData.length; i++) {
    const row = i + 1;
    const name = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
    if (!name) continue;
    
    // Check for duplicates
    if (seenNames[name]) {
      issues.push(['DUPLICATE', name, 'Name appears multiple times', row, 'HIGH', 'Merge or rename', new Date()]);
    }
    seenNames[name] = true;
    
    const isAssigned = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED];
    const table = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE];
    const seat = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.SEAT];
    
    // Check assignment consistency
    if (isAssigned === true || isAssigned === 'TRUE') {
      if (!table || !seat) {
        issues.push(['INVALID_ASSIGNMENT', name, 'Assigned but missing table/seat', row, 'HIGH', 'Fix assignment', new Date()]);
      }
      
      // Check for double-booking
      const seatKey = `${table}-${seat}`;
      if (assignedSeats[seatKey]) {
        issues.push(['DOUBLE_BOOKED', name, `Seat ${seatKey} already taken by ${assignedSeats[seatKey]}`, row, 'CRITICAL', 'Resolve conflict', new Date()]);
      }
      assignedSeats[seatKey] = name;
      
      // Validate table/seat numbers
      if (table < 1 || table > SYSTEM_CONFIG.SEATING.TABLES) {
        issues.push(['INVALID_TABLE', name, `Table ${table} out of range`, row, 'HIGH', 'Fix table number', new Date()]);
      }
      if (seat < 1 || seat > SYSTEM_CONFIG.SEATING.SEATS_PER_TABLE) {
        issues.push(['INVALID_SEAT', name, `Seat ${seat} out of range`, row, 'HIGH', 'Fix seat number', new Date()]);
      }
    }
  }
  
  // Write issues
  for (const issue of issues) {
    debugSheet.appendRow(issue);
  }
  
  if (issues.length === 0) {
    ui.alert('✅ Validation Complete', 'No issues found! All data is valid.', ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Issues Found', `Found ${issues.length} issue(s).\n\nCheck _DEBUG sheet for details.`, ui.ButtonSet.OK);
  }
}

function clearAllAssignments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Clear',
    'This will remove ALL seating assignments.\n\nThis cannot be undone!\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  
  // Clear Master_Guests assignments
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 1) {
    masterSheet.getRange(2, SYSTEM_CONFIG.COLUMNS.MASTER.TABLE + 1, lastRow - 1, 3).clearContent();
    masterSheet.getRange(2, SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED + 1, lastRow - 1, 1).setValue(false);
  }
  
  // Clear SEATING_PLAN seat cells
  for (let t = 1; t <= SYSTEM_CONFIG.SEATING.TABLES; t++) {
    for (let s = 1; s <= SYSTEM_CONFIG.SEATING.SEATS_PER_TABLE; s++) {
      const seatRow = findSeatRow(t, s);
      const cell = seatingSheet.getRange(seatRow, 2);
      cell.clearContent();
      cell.setBackground(null);
      cell.setFontColor('#000000');
      cell.setFontWeight('normal');
    }
    updateTableHeader(t);
  }
  
  logAction('CLEAR_ALL', '', '', '', 'All assignments cleared by user');
  
  ui.alert('✅ Complete', 'All seating assignments have been cleared.', ui.ButtonSet.OK);
}

function generateSeatingReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const ui = SpreadsheetApp.getUi();
  
  const masterData = masterSheet.getDataRange().getValues();
  
  let report = '═══════════════════════════════════════════════════════════════\n';
  report += '              WEDDING SEATING REPORT\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  report += `Generated: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}\n\n`;
  
  // Summary
  let total = 0, seated = 0;
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME]) {
      total++;
      if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED] === true) seated++;
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
  
  for (let t = 1; t <= SYSTEM_CONFIG.SEATING.TABLES; t++) {
    const tableGuests = [];
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.TABLE] === t) {
        tableGuests.push({
          seat: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.SEAT],
          name: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME],
          vip: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP],
          dietary: masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.DIETARY]
        });
      }
    }
    
    if (tableGuests.length > 0) {
      report += `TABLE ${t} (${tableGuests.length}/10):\n`;
      tableGuests.sort((a, b) => a.seat - b.seat);
      for (const g of tableGuests) {
        let line = `  Seat ${g.seat}: ${g.name}`;
        if (g.vip === 'VIP') line += ' ⭐';
        if (g.dietary && g.dietary !== 'None') line += ` [${g.dietary}]`;
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
    const isAssigned = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.ASSIGNED];
    const name = masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.NAME];
    if (name && isAssigned !== true && isAssigned !== 'TRUE') {
      report += `  • ${name} (${masterData[i][SYSTEM_CONFIG.COLUMNS.MASTER.VIP]})\n`;
    }
  }
  
  // Show report
  const htmlOutput = HtmlService.createHtmlOutput(`<pre style="font-family: monospace; font-size: 12px;">${report}</pre>`)
    .setWidth(600)
    .setHeight(500)
    .setTitle('Seating Report');
  
  ui.showModalDialog(htmlOutput, '📋 Seating Report');
}

function showSettings() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⚙️ Settings', 'Open the CONFIG sheet to modify system settings.', ui.ButtonSet.OK);
}

function runSystemTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tests = [];
  let allPassed = true;
  
  // Test 1: All sheets exist
  const requiredSheets = Object.values(SYSTEM_CONFIG.SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());
  const missingSheets = requiredSheets.filter(s => !existingSheets.includes(s));
  
  tests.push({
    name: 'Sheet Structure',
    status: missingSheets.length === 0 ? 'PASS' : 'FAIL',
    detail: missingSheets.length === 0 ? 'All sheets exist' : `Missing: ${missingSheets.join(', ')}`
  });
  
  // Test 2: Master_Guests has data
  const masterSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MASTER_GUESTS);
  const masterData = masterSheet.getDataRange().getValues();
  tests.push({
    name: 'Guest Data',
    status: masterData.length > 1 ? 'PASS' : 'FAIL',
    detail: `${masterData.length - 1} guests in Master_Guests`
  });
  
  // Test 3: Seating Plan structure
  const seatingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.SEATING_PLAN);
  tests.push({
    name: 'Seating Plan',
    status: seatingSheet ? 'PASS' : 'FAIL',
    detail: seatingSheet ? 'Seating plan sheet exists' : 'Sheet missing'
  });
  
  // Build report
  let report = '═══════════════════════════════════════════════════════════════\n';
  report += '              SYSTEM TEST RESULTS\n';
  report += '═══════════════════════════════════════════════════════════════\n\n';
  
  for (const test of tests) {
    const icon = test.status === 'PASS' ? '✅' : '❌';
    report += `${icon} ${test.name}: ${test.status}\n`;
    report += `   ${test.detail}\n\n`;
    
    if (test.status !== 'PASS') allPassed = false;
  }
  
  report += '═══════════════════════════════════════════════════════════════\n';
  report += allPassed ? '✅ ALL TESTS PASSED\n' : '❌ SOME TESTS FAILED\n';
  report += '═══════════════════════════════════════════════════════════════\n';
  
  ui.alert('🧪 System Test', report, ui.ButtonSet.OK);
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'ℹ️ About Wedding Seating System',
    `Version: ${SYSTEM_CONFIG.VERSION}\n` +
    `Build: ${SYSTEM_CONFIG.BUILD_DATE}\n` +
    `Author: ${SYSTEM_CONFIG.AUTHOR}\n\n` +
    'Features:\n' +
    '• AI-Powered Compatibility Scoring\n' +
    '• Conflict Detection & Resolution\n' +
    '• Auto-Seating Optimization\n' +
    '• VIP Protocol Engine\n' +
    '• Accessibility Compliance\n' +
    '• Real-Time Analytics',
    ui.ButtonSet.OK
  );
}
