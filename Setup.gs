/**
 * Roy-Hart Survivor Week — Setup
 * 
 * Run setupSheets() ONCE from the Apps Script editor to create all data tabs
 * with the correct headers. Safe to re-run; it skips tabs that already exist.
 *
 * Run loadTestData() to populate fake teachers across all 6 tribes for testing.
 * Run wipeTestData() to clear it back out.
 */

// ============================================================================
// SHEET SCHEMA — header row for each tab
// ============================================================================
const SHEET_SCHEMA = {
  Members:              ['id', 'name', 'email', 'tribe_id', 'status', 'competed_days', 'original_tribe_id', 'joined_at'],
  Tribes:               ['id', 'name', 'color_label', 'color_hex', 'status', 'points', 'absorbed_into', 'absorbed_at'],
  Challenges:           ['day', 'date', 'type', 'name', 'status', 'opens_at', 'closes_at', 'results_locked'],
  Selections:           ['tribe_id', 'day', 'voter_email', 'candidate_email', 'timestamp'],
  Competitors:          ['tribe_id', 'day', 'member_id', 'member_name', 'locked_at'],
  ChallengeResults:     ['day', 'tribe_id', 'placement', 'raw_time_or_score', 'entered_by', 'entered_at'],
  Eliminations:         ['tribe_id', 'day', 'voter_email', 'target_email', 'timestamp'],
  EliminationResults:   ['day', 'tribe_id', 'eliminated_member_ids', 'resolved_at'],
  EscapeRoomClues:      ['clue_number', 'release_time', 'title', 'riddle_text', 'cipher_answer', 'image_url'],
  EscapeRoomProgress:   ['tribe_id', 'clue_number', 'solved_at', 'attempt_count'],
  PuzzlePhotos:         ['tribe_id', 'photo_url', 'submitted_by', 'timestamp', 'verified_by', 'verified_at'],
  TribeChat:            ['tribe_id', 'member_email', 'member_name', 'message', 'timestamp'],
  JuryChat:             ['member_email', 'member_name', 'message', 'timestamp'],
  Confessionals:        ['member_email', 'member_name', 'tribe_id_at_time', 'text', 'timestamp', 'hidden'],
  IdolState:            ['idol_id', 'idol_name', 'secret_word', 'hint_log', 'found_by_email', 'claimed_at', 'played_for_day', 'played_at'],
  JuryVotes:            ['voter_email', 'target_email', 'timestamp'],
  AdminLog:             ['timestamp', 'admin_email', 'action', 'details']
};

// ============================================================================
// MAIN SETUP — run this once after pasting code into the Apps Script editor
// ============================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let created = 0, existed = 0;

  // Create each tab with headers
  Object.keys(SHEET_SCHEMA).forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (sheet) {
      existed++;
      return;
    }
    sheet = ss.insertSheet(tabName);
    const headers = SHEET_SCHEMA[tabName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#2D1B3D')
      .setFontColor('#F3F3EE');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    created++;
  });

  // Seed the Tribes tab if empty
  const tribesSheet = ss.getSheetByName('Tribes');
  if (tribesSheet && tribesSheet.getLastRow() < 2) {
    const rows = TRIBES.map(function(t) {
      return [t.id, t.name, t.color_label, t.color_hex, 'active', 0, '', ''];
    });
    tribesSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Seed the Challenges tab if empty
  const challengesSheet = ss.getSheetByName('Challenges');
  if (challengesSheet && challengesSheet.getLastRow() < 2) {
    const rows = CHALLENGES.map(function(c) {
      return [c.day, c.date, c.type, c.name, 'pending', '', '', false];
    });
    challengesSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Seed the EscapeRoomClues tab with our 4 clues
  const cluesSheet = ss.getSheetByName('EscapeRoomClues');
  if (cluesSheet && cluesSheet.getLastRow() < 2) {
    const clues = [
      [1, '2026-05-07T08:00:00', 'Out of Office',
        "The vacation is finally here. The bell has rung for the last time, the lockers are empty, and the calendar shows that one beautiful word. Type the season that follows June 21.",
        'SUMMER', ''],
      [2, '2026-05-07T09:30:00', 'Faculty Lounge Logic Grid',
        "Four teachers (Miller, Davis, Garcia, and Chen) are planning their summer trips to four different spots (Beach, Mountains, City, Lake).\n\n• Miller is NOT going to the Lake or the Beach.\n• The teacher going to the City has a shorter name than Garcia.\n• Chen is going to the Beach.\n• Davis is NOT going to the City.\n\nType the LAST NAME of the teacher going to the Mountains.",
        'MILLER', ''],
      [3, '2026-05-07T11:00:00', 'The Scavenger Desk',
        "Look closely at the desk. The 3-digit code is hidden in the clutter.\n\n• Digit 1: Number of red pens in the cup\n• Digit 2: Number of coffee mugs on the desk\n• Digit 3: The number on the calculator screen",
        '635', ''],
      [4, '2026-05-07T12:30:00', 'Acronym Soup',
        "In Middle School, we speak in acronyms. Take the first letter of each:\n\n1. Social Emotional Learning\n2. Teacher Evaluation System\n3. Augmented Reality Lab\n4. Year-End Party\n\nType the 4-letter code.",
        'STAY', '']
    ];
    cluesSheet.getRange(2, 1, clues.length, clues[0].length).setValues(clues);
  }

  // Initialize / repair IdolState — always ensure correct schema and both rows.
  // This is idempotent: safe to run repeatedly. It preserves any existing
  // secret_word / hint_log / found_by_email / claimed_at data when migrating
  // from the old single-row schema.
  const idolSheet = ss.getSheetByName('IdolState');
  if (idolSheet) {
    const expectedHeaders = SHEET_SCHEMA.IdolState; // ['idol_id', 'idol_name', 'secret_word', 'hint_log', 'found_by_email', 'claimed_at', 'played_for_day', 'played_at']

    // Read whatever's currently there
    const existingData = idolSheet.getDataRange().getValues();
    const existingHeaders = existingData.length > 0 ? existingData[0] : [];
    const headersMatch = expectedHeaders.length === existingHeaders.length &&
      expectedHeaders.every(function(h, i) { return existingHeaders[i] === h; });

    // Capture existing row data, mapped by header name (so column moves don't matter)
    const existingRowsByIdolId = {};   // 'rammy' / 'bear' -> {column_name: value}
    let legacyRow = null;              // single-row pre-migration data, mapped by old header name

    if (existingData.length >= 2) {
      // Detect old schema (no idol_id column) vs new
      const idCol = existingHeaders.indexOf('idol_id');
      if (idCol === -1) {
        // OLD schema. The single data row had columns:
        // secret_word, hint_log, found_by_email, claimed_at, played_for_day, played_at
        legacyRow = {};
        for (let c = 0; c < existingHeaders.length; c++) {
          legacyRow[existingHeaders[c]] = existingData[1][c];
        }
      } else {
        // NEW schema. Capture each row by its idol_id.
        for (let r = 1; r < existingData.length; r++) {
          const id = (existingData[r][idCol] || '').toString();
          if (id !== 'rammy' && id !== 'bear') continue;
          const rec = {};
          for (let c = 0; c < existingHeaders.length; c++) {
            rec[existingHeaders[c]] = existingData[r][c];
          }
          existingRowsByIdolId[id] = rec;
        }
      }
    }

    // Rewrite headers if they don't match
    if (!headersMatch) {
      // Clear the entire sheet's content first, then rewrite header
      idolSheet.clear();
      idolSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      idolSheet.getRange(1, 1, 1, expectedHeaders.length)
        .setFontWeight('bold')
        .setBackground('#2D1B3D')
        .setFontColor('#F3F3EE');
      idolSheet.setFrozenRows(1);
    }

    // Build the two canonical rows, preserving any captured data
    function buildRow(idolId, idolName) {
      const rec = existingRowsByIdolId[idolId] || (idolId === 'rammy' && legacyRow) || {};
      return [
        idolId,
        idolName,
        rec.secret_word || '',
        rec.hint_log || '',
        rec.found_by_email || '',
        rec.claimed_at || '',
        rec.played_for_day || '',
        rec.played_at || ''
      ];
    }

    // Always write both rows in canonical positions (rows 2 and 3)
    idolSheet.getRange(2, 1, 2, expectedHeaders.length).setValues([
      buildRow('rammy', 'Rammy'),
      buildRow('bear',  'The Bear')
    ]);

    // Truncate any extra rows beyond row 3
    if (idolSheet.getLastRow() > 3) {
      idolSheet.getRange(4, 1, idolSheet.getLastRow() - 3, idolSheet.getLastColumn()).clearContent();
    }
  }

  // Log the setup action
  logAdminAction('SYSTEM', 'setupSheets', 'Created ' + created + ' new tabs, ' + existed + ' already existed.');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Setup complete: ' + created + ' tabs created, ' + existed + ' already existed.',
    'Roy-Hart Survivor', 5
  );
}

// ============================================================================
// TEST DATA — load fake teachers for development
// ============================================================================
function loadTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = ss.getSheetByName('Members');
  if (!members) {
    SpreadsheetApp.getUi().alert('Run setupSheets() first.');
    return;
  }
  if (members.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert('Members tab is not empty. Replace existing data with test data?',
                          ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    members.getRange(2, 1, members.getLastRow() - 1, members.getLastColumn()).clearContent();
  }

  const firstNames = ['Avery','Bailey','Casey','Drew','Emery','Finley','Gray','Harper','Indigo','Jordan',
                      'Kai','Logan','Morgan','Nico','Oakley','Parker','Quinn','Riley','Sage','Taylor',
                      'Reese','Skyler','Rowan','Hayden','Marlowe','Spencer','Phoenix','Devon','Robin','Lane',
                      'Blake','Cameron','Dakota','Ellis','Frankie','Gentry'];
  const lastNames  = ['Aster','Briggs','Calder','Driscoll','Ellsworth','Forrester','Greaves','Holloway','Ives','Jansen'];

  const rows = [];
  let idx = 0;
  // 6 teachers per tribe = 36 total
  TRIBES.forEach(function(tribe) {
    for (let i = 0; i < 6; i++) {
      const fn = firstNames[idx % firstNames.length];
      const ln = lastNames[(idx + i) % lastNames.length];
      const id = 'm_' + (idx + 1).toString().padStart(3, '0');
      // For testing purposes only - these are not real emails
      const email = 'test_' + fn.toLowerCase() + '.' + ln.toLowerCase() + '@royhart.org';
      rows.push([id, fn + ' ' + ln, email, tribe.id, 'active', '', tribe.id, new Date().toISOString()]);
      idx++;
    }
  });

  members.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  logAdminAction('SYSTEM', 'loadTestData', 'Loaded ' + rows.length + ' test members.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Loaded ' + rows.length + ' test members.', 'Roy-Hart Survivor', 5);
}

function wipeTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = ss.getSheetByName('Members');
  if (!members) return;
  if (members.getLastRow() > 1) {
    members.getRange(2, 1, members.getLastRow() - 1, members.getLastColumn()).clearContent();
  }
  logAdminAction('SYSTEM', 'wipeTestData', 'Members tab cleared.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Members tab cleared.', 'Roy-Hart Survivor', 3);
}

// ============================================================================
// ADMIN LOG — internal helper
// ============================================================================
function logAdminAction(adminEmail, action, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AdminLog');
  if (!sheet) return;
  sheet.appendRow([new Date().toISOString(), adminEmail || 'SYSTEM', action, details || '']);
}

/**
 * One-shot repair function — fixes the IdolState tab if it's in a bad state.
 * Run this from the Apps Script editor: Run menu → repairIdolState.
 *
 * Safe to run any time. Preserves any existing secret words, hints, holders.
 */
function repairIdolState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idolSheet = ss.getSheetByName('IdolState');
  if (!idolSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast('IdolState tab not found. Run setupSheets() first.', 'Error', 5);
    return;
  }

  const expectedHeaders = SHEET_SCHEMA.IdolState;
  const existingData = idolSheet.getDataRange().getValues();
  const existingHeaders = existingData.length > 0 ? existingData[0] : [];

  const existingRowsByIdolId = {};
  let legacyRow = null;

  if (existingData.length >= 2) {
    const idCol = existingHeaders.indexOf('idol_id');
    if (idCol === -1) {
      legacyRow = {};
      for (let c = 0; c < existingHeaders.length; c++) {
        legacyRow[existingHeaders[c]] = existingData[1][c];
      }
    } else {
      for (let r = 1; r < existingData.length; r++) {
        const id = (existingData[r][idCol] || '').toString();
        if (id !== 'rammy' && id !== 'bear') continue;
        const rec = {};
        for (let c = 0; c < existingHeaders.length; c++) {
          rec[existingHeaders[c]] = existingData[r][c];
        }
        existingRowsByIdolId[id] = rec;
      }
    }
  }

  // Hard reset: clear, rewrite headers, write both rows
  idolSheet.clear();
  idolSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  idolSheet.getRange(1, 1, 1, expectedHeaders.length)
    .setFontWeight('bold')
    .setBackground('#2D1B3D')
    .setFontColor('#F3F3EE');
  idolSheet.setFrozenRows(1);

  function buildRow(idolId, idolName) {
    const rec = existingRowsByIdolId[idolId] || (idolId === 'rammy' && legacyRow) || {};
    return [
      idolId, idolName,
      rec.secret_word || '',
      rec.hint_log || '',
      rec.found_by_email || '',
      rec.claimed_at || '',
      rec.played_for_day || '',
      rec.played_at || ''
    ];
  }

  idolSheet.getRange(2, 1, 2, expectedHeaders.length).setValues([
    buildRow('rammy', 'Rammy'),
    buildRow('bear',  'The Bear')
  ]);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'IdolState repaired. Rammy and Bear rows reset with correct schema. Existing data preserved.',
    'Done', 5
  );
}
