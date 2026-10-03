/**
 * Roy-Hart Survivor Week — Main Server Code
 * 
 * Contains:
 *   - doGet(e): web app entry point + routing
 *   - getUserRole(): admin / active / jury / unknown
 *   - findMemberByEmail(): looks up roster
 *   - include(): partial template helper
 *   - getCurrentUser(): client-callable user info
 */

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================
function doGet(e) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const params = (e && e.parameter) || {};
  const view = params.view || null;

  const role   = getUserRole(email);
  const member = (role === 'active' || role === 'jury') ? findMemberByEmail(email) : null;

  // Routing logic:
  //  - ?view=contest     → public contest dashboard (any logged-in user)
  //  - admin user        → Admin dashboard
  //  - jury user         → Jury dashboard
  //  - active user       → Tribe dashboard
  //  - anything else     → NotOnRoster page
  let templateName, pageTitle;

  if (view === 'contest') {
    templateName = 'Contest';
    pageTitle    = 'Roy-Hart Survivor — Contest';
  } else if (view === 'tribe' && member) {
    // Allows admins-who-are-also-members AND jury members to view their tribe page
    templateName = 'Tribe';
    pageTitle    = 'Roy-Hart Survivor — Tribe';
  } else if (view === 'jury' && (role === 'jury' || role === 'admin')) {
    templateName = 'Jury';
    pageTitle    = 'Roy-Hart Survivor — Jury';
  } else if (role === 'admin') {
    templateName = 'Admin';
    pageTitle    = 'Roy-Hart Survivor — Admin';
  } else if (role === 'jury') {
    templateName = 'Jury';
    pageTitle    = 'Roy-Hart Survivor — Jury';
  } else if (role === 'active') {
    templateName = 'Tribe';
    pageTitle    = 'Roy-Hart Survivor — Tribe';
  } else {
    templateName = 'NotOnRoster';
    pageTitle    = 'Roy-Hart Survivor';
  }

  const tpl = HtmlService.createTemplateFromFile(templateName);
  tpl.userEmail = email;
  tpl.userRole  = role;
  tpl.member    = member; // may be null
  tpl.tribe     = member ? getTribeById(member.tribe_id) : null;
  tpl.scriptUrl = ScriptApp.getService().getUrl();
  tpl.pageView  = ({ Contest: 'contest', Tribe: 'tribe', Jury: 'jury', Admin: 'admin' })[templateName] || 'other';
  tpl.tribesJson = JSON.stringify(TRIBES);

  return tpl.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// ROLE DETECTION
// ============================================================================
function getUserRole(email) {
  if (!email) return 'unknown';
  if (ADMIN_EMAILS.indexOf(email) !== -1) return 'admin';

  const member = findMemberByEmail(email);
  if (!member) return 'unknown';
  if (member.status === 'jury')   return 'jury';
  if (member.status === 'active') return 'active';
  return 'unknown';
}

function findMemberByEmail(email) {
  if (!email) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const headers   = data[0];
  const emailCol  = headers.indexOf('email');
  const nameCol   = headers.indexOf('name');
  const tribeCol  = headers.indexOf('tribe_id');
  const statusCol = headers.indexOf('status');
  const idCol     = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    const rowEmail = (data[i][emailCol] || '').toString().toLowerCase();
    if (rowEmail === email.toLowerCase()) {
      return {
        id:       data[i][idCol],
        name:     data[i][nameCol],
        email:    data[i][emailCol],
        tribe_id: data[i][tribeCol],
        status:   data[i][statusCol]
      };
    }
  }
  return null;
}

function getTribeById(tribeId) {
  for (let i = 0; i < TRIBES.length; i++) {
    if (TRIBES[i].id === tribeId) return TRIBES[i];
  }
  return null;
}

// ============================================================================
// CLIENT-CALLABLE — used by google.script.run from front-end
// ============================================================================
function getCurrentUser() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const role  = getUserRole(email);
  const member = (role === 'active' || role === 'jury') ? findMemberByEmail(email) : null;
  return {
    email:  email,
    role:   role,
    member: member,
    tribe:  member ? getTribeById(member.tribe_id) : null
  };
}

// ============================================================================
// TEMPLATE HELPER — lets HTML files include other HTML files (with scriptlets evaluated)
// Pass an optional params object to thread variables into the included template.
// ============================================================================
function include(filename, params) {
  const tpl = HtmlService.createTemplateFromFile(filename);
  if (params && typeof params === 'object') {
    Object.keys(params).forEach(function(k) { tpl[k] = params[k]; });
  }
  return tpl.evaluate().getContent();
}

// ============================================================================
// CONTEST URL — share this with everyone for the public dashboard view
// ============================================================================
function getContestDashboardUrl() {
  return ScriptApp.getService().getUrl() + '?view=contest';
}

// ============================================================================
// ROSTER MANAGEMENT (Phase 1.5)
// All functions below are admin-only and write to the Members tab.
// ============================================================================

/**
 * Throws if the active user is not an admin. Used to guard write operations.
 */
function requireAdmin_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email || ADMIN_EMAILS.indexOf(email) === -1) {
    throw new Error('Unauthorized: admin access required.');
  }
  return email;
}

/**
 * Returns all members grouped by tribe. Shape:
 *   {
 *     tribes: [{id, name, color_hex, members: [...]}, ...],
 *     totalCount: number,
 *     activeCount: number,
 *     juryCount: number
 *   }
 */
function listMembers() {
  requireAdmin_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol     = headers.indexOf('id');
  const nameCol   = headers.indexOf('name');
  const emailCol  = headers.indexOf('email');
  const tribeCol  = headers.indexOf('tribe_id');
  const statusCol = headers.indexOf('status');

  // Initialize a bucket per tribe
  const grouped = {};
  TRIBES.forEach(function(t) {
    grouped[t.id] = {
      id: t.id, name: t.name, color_label: t.color_label, color_hex: t.color_hex,
      members: []
    };
  });

  let activeCount = 0, juryCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[idCol]) continue; // skip empty rows
    const member = {
      id:       row[idCol],
      name:     row[nameCol],
      email:    row[emailCol],
      tribe_id: row[tribeCol],
      status:   row[statusCol] || 'active'
    };
    if (grouped[member.tribe_id]) {
      grouped[member.tribe_id].members.push(member);
    }
    if (member.status === 'jury') juryCount++; else activeCount++;
  }

  // Sort each tribe's members alphabetically by name
  Object.keys(grouped).forEach(function(tid) {
    grouped[tid].members.sort(function(a, b) {
      return (a.name || '').localeCompare(b.name || '');
    });
  });

  return {
    tribes: TRIBES.map(function(t) { return grouped[t.id]; }),
    totalCount: activeCount + juryCount,
    activeCount: activeCount,
    juryCount: juryCount
  };
}

/**
 * Adds a new member. Validates email format, prevents duplicates.
 * Returns the created member object.
 */
function addMember(name, email, tribeId) {
  const adminEmail = requireAdmin_();
  name = (name || '').toString().trim();
  email = (email || '').toString().trim().toLowerCase();

  if (!name) throw new Error('Name is required.');
  if (!email) throw new Error('Email is required.');
  if (!/^[^\s@]+@royhart\.org$/i.test(email)) {
    throw new Error('Email must end with @royhart.org');
  }
  if (!getTribeById(tribeId)) throw new Error('Invalid tribe.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('email');

  for (let i = 1; i < data.length; i++) {
    if ((data[i][emailCol] || '').toString().toLowerCase() === email) {
      throw new Error('A member with this email already exists.');
    }
  }

  const id = 'm_' + Utilities.getUuid().split('-')[0];
  const newRow = [
    id, name, email, tribeId, 'active', '', tribeId, new Date().toISOString()
  ];
  sheet.appendRow(newRow);
  logAdminAction(adminEmail, 'addMember', name + ' (' + email + ') → ' + tribeId);

  return {
    id: id, name: name, email: email, tribe_id: tribeId, status: 'active'
  };
}

/**
 * Updates an existing member's name, email, and/or tribe.
 * Pass null for a field to skip changing it.
 */
function updateMember(id, name, email, tribeId) {
  const adminEmail = requireAdmin_();
  if (!id) throw new Error('Member id is required.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol    = headers.indexOf('id');
  const nameCol  = headers.indexOf('name');
  const emailCol = headers.indexOf('email');
  const tribeCol = headers.indexOf('tribe_id');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('Member not found.');

  // Validate inputs
  if (name !== null && name !== undefined) {
    name = name.toString().trim();
    if (!name) throw new Error('Name cannot be blank.');
  }
  if (email !== null && email !== undefined) {
    email = email.toString().trim().toLowerCase();
    if (!/^[^\s@]+@royhart\.org$/i.test(email)) {
      throw new Error('Email must end with @royhart.org');
    }
    // Check duplicate
    for (let i = 1; i < data.length; i++) {
      if (i === rowIndex) continue;
      if ((data[i][emailCol] || '').toString().toLowerCase() === email) {
        throw new Error('Another member already has this email.');
      }
    }
  }
  if (tribeId !== null && tribeId !== undefined) {
    if (!getTribeById(tribeId)) throw new Error('Invalid tribe.');
  }

  // Apply updates (sheet rows are 1-indexed; +1 converts our 0-indexed array idx)
  const sheetRow = rowIndex + 1;
  if (name)    sheet.getRange(sheetRow, nameCol  + 1).setValue(name);
  if (email)   sheet.getRange(sheetRow, emailCol + 1).setValue(email);
  if (tribeId) sheet.getRange(sheetRow, tribeCol + 1).setValue(tribeId);

  logAdminAction(adminEmail, 'updateMember', id + ' updated');
  return { id: id, name: name, email: email, tribe_id: tribeId };
}

/**
 * Removes a member by id (deletes the entire row).
 */
function removeMember(id) {
  const adminEmail = requireAdmin_();
  if (!id) throw new Error('Member id is required.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      sheet.deleteRow(i + 1);
      logAdminAction(adminEmail, 'removeMember', id);
      return { ok: true };
    }
  }
  throw new Error('Member not found.');
}

/**
 * Wipes the entire roster (after confirmation client-side).
 * Use this to clear test data before entering the real Friday-draw roster.
 */
function wipeRoster() {
  const adminEmail = requireAdmin_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  logAdminAction(adminEmail, 'wipeRoster', 'Cleared all members.');
  return { ok: true };
}

/**
 * Returns the list of tribes (for the picker UI). Public — no admin guard.
 */
function getTribes() {
  return TRIBES.slice();
}

/**
 * Returns the admin email allowlist (for the diagnostic display in Admin.html).
 * Admin-only — no point exposing this list to non-admins.
 */
function getAdminEmails() {
  requireAdmin_();
  return ADMIN_EMAILS.slice();
}

// ============================================================================
// CONTEST STATE (Phase 2)
// ============================================================================

const DAYS_ORDER = ['mon', 'tue', 'wed', 'thu', 'fri'];

const DAY_TO_CHALLENGE = {
  'mon': { name: 'Cornhole',          type: 'cornhole',    icon: '◈' },
  'tue': { name: 'Foosball',          type: 'foosball',    icon: '◇' },
  'wed': { name: 'Puzzle Race',       type: 'puzzle',      icon: '◆' },
  'thu': { name: 'Survivor Escape',   type: 'escape_room', icon: '✦' },
  'fri': { name: 'Obstacle Course',   type: 'obstacle',    icon: '★' }
};

const CHALLENGE_DATE_TO_DAY = {
  '2026-05-04': 'mon',
  '2026-05-05': 'tue',
  '2026-05-06': 'wed',
  '2026-05-07': 'thu',
  '2026-05-08': 'fri'
};

/**
 * Determines current contest day. Returns 'mon'|'tue'|...|'fri', or null
 * (pre-contest) or 'post' (post-contest). Honors a SIM_DAY script property
 * for dry-run testing.
 */
function getCurrentDay_() {
  const sim = PropertiesService.getScriptProperties().getProperty('SIM_DAY');
  if (sim) {
    if (sim === 'pre') return null;
    if (sim === 'post') return 'post';
    if (DAYS_ORDER.indexOf(sim) !== -1) return sim;
  }
  const now = new Date();
  const dStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (CHALLENGE_DATE_TO_DAY[dStr]) return CHALLENGE_DATE_TO_DAY[dStr];
  if (dStr < '2026-05-04') return null;
  if (dStr > '2026-05-08') return 'post';
  return null;
}

function getPreviousDay_(day) {
  if (day === 'post') return 'fri';
  const idx = DAYS_ORDER.indexOf(day);
  if (idx <= 0) return null;
  return DAYS_ORDER[idx - 1];
}

// ----------------------------------------------------------------------------
// SHEET READERS — these centralize parsing and are used by getContestState()
// ----------------------------------------------------------------------------
function readMembers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const idCol = h.indexOf('id');
  const nameCol = h.indexOf('name');
  const emailCol = h.indexOf('email');
  const tribeCol = h.indexOf('tribe_id');
  const statusCol = h.indexOf('status');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][idCol]) continue;
    out.push({
      id: data[i][idCol],
      name: data[i][nameCol],
      email: data[i][emailCol],
      tribe_id: data[i][tribeCol],
      status: data[i][statusCol] || 'active'
    });
  }
  return out;
}

function readChallengeResults_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.CHALLENGE_RESULTS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const dayCol = h.indexOf('day');
  const tribeCol = h.indexOf('tribe_id');
  const placementCol = h.indexOf('placement');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][dayCol] || !data[i][tribeCol]) continue;
    out.push({
      day: data[i][dayCol],
      tribe_id: data[i][tribeCol],
      placement: parseInt(data[i][placementCol], 10) || null
    });
  }
  return out;
}

function readEliminationResults_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATION_RESULTS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const dayCol = h.indexOf('day');
  const tribeCol = h.indexOf('tribe_id');
  const idsCol = h.indexOf('eliminated_member_ids');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][dayCol]) continue;
    const ids = (data[i][idsCol] || '').toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    out.push({
      day: data[i][dayCol],
      tribe_id: data[i][tribeCol],
      member_ids: ids
    });
  }
  return out;
}

function readChallengeStatuses_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.CHALLENGES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  const h = data[0];
  const dayCol = h.indexOf('day');
  const statusCol = h.indexOf('status');
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][dayCol]) continue;
    map[data[i][dayCol]] = data[i][statusCol] || 'pending';
  }
  return map;
}

// ----------------------------------------------------------------------------
// SCORING
// ----------------------------------------------------------------------------
function pointsForPlacement_(placement, day) {
  if (!placement || placement < 1 || placement > 6) return 0;
  const table = (day === 'fri') ? SCORING_FRIDAY : SCORING_DAILY;
  return table[placement - 1] || 0;
}

// ----------------------------------------------------------------------------
// MAIN STATE FUNCTION — what the dashboard calls
// ----------------------------------------------------------------------------
function getContestState() {
  const members         = readMembers_();
  const results         = readChallengeResults_();
  const eliminations    = readEliminationResults_();
  const challengeStatus = readChallengeStatuses_();
  const currentDay      = getCurrentDay_();

  // Build tribe-keyed lookup of eliminated member ids (from any day)
  const eliminatedIds = {};
  eliminations.forEach(function(e) {
    e.member_ids.forEach(function(id) { eliminatedIds[id] = e.day; });
  });

  // Compute per-tribe data
  const tribeStats = {};
  TRIBES.forEach(function(t) {
    tribeStats[t.id] = {
      id: t.id, name: t.name, color_label: t.color_label, color_hex: t.color_hex,
      points: 0, members_active: 0, members_jury: 0, members_total: 0,
      placements: {}, // { day: placement }
      points_today: 0
    };
  });

  // Members
  members.forEach(function(m) {
    const t = tribeStats[m.tribe_id];
    if (!t) return;
    t.members_total++;
    if (m.status === 'jury' || eliminatedIds[m.id]) t.members_jury++;
    else t.members_active++;
  });

  // Results → points
  results.forEach(function(r) {
    const t = tribeStats[r.tribe_id];
    if (!t) return;
    const pts = pointsForPlacement_(r.placement, r.day);
    t.points += pts;
    t.placements[r.day] = r.placement;
    if (r.day === currentDay) t.points_today = pts;
  });

  // To array, sorted by points desc, with stable rank
  const leaderboard = TRIBES.map(function(t) { return tribeStats[t.id]; });
  leaderboard.sort(function(a, b) {
    if (b.points !== a.points) return b.points - a.points;
    return a.name.localeCompare(b.name);
  });
  leaderboard.forEach(function(t, idx) { t.rank = idx + 1; });

  // Hero block — state-dependent headline copy
  const hero = buildHero_(currentDay, leaderboard, members);

  // Today panel
  const today = buildToday_(currentDay, challengeStatus, leaderboard);

  // Yesterday panel
  const prevDay = (currentDay && currentDay !== 'post') ? getPreviousDay_(currentDay)
                                                        : (currentDay === 'post' ? 'fri' : null);
  const yesterday = prevDay ? buildYesterday_(prevDay, results, eliminations, members) : null;

  return {
    now: new Date().toISOString(),
    state: currentDay === null ? 'pre_contest'
         : currentDay === 'post' ? 'post_contest'
         : 'active',
    currentDay: currentDay,
    hero: hero,
    leaderboard: leaderboard,
    today: today,
    yesterday: yesterday,
    idol: { state: 'hidden' },              // Phase 4/7
    confessionals: [],                      // Phase 7
    rosterCount: members.length
  };
}

function buildHero_(currentDay, leaderboard, members) {
  if (!currentDay) {
    if (members.length === 0) {
      return { eyebrow: 'Pre-Contest', headline: 'Awaiting Castaways',
               subline: 'The tickets have not yet been drawn.', accent: 'parchment' };
    }
    return { eyebrow: 'The Tribes Have Been Drawn',
             headline: 'Stand By',
             subline: 'Day 1 begins Monday, May 4',
             accent: 'ember' };
  }
  if (currentDay === 'post') {
    const winner = leaderboard[0];
    return { eyebrow: 'The Tribe Has Spoken',
             headline: winner ? winner.name.toUpperCase() : 'WINNERS',
             subline: 'Champion crowned · Pending Jury MVP',
             accent: winner ? winner.color_hex : 'ember' };
  }
  const ch = DAY_TO_CHALLENGE[currentDay];
  const dayLabel = { mon: 'DAY 1 · MONDAY', tue: 'DAY 2 · TUESDAY', wed: 'DAY 3 · WEDNESDAY',
                     thu: 'DAY 4 · THURSDAY', fri: 'DAY 5 · FRIDAY (FINALE)' }[currentDay];
  return { eyebrow: dayLabel,
           headline: ch.name.toUpperCase(),
           subline: '',  // filled by today panel status
           accent: 'ember',
           dayIcon: ch.icon };
}

function buildToday_(currentDay, statusMap, leaderboard) {
  if (!currentDay || currentDay === 'post') return null;
  const ch = DAY_TO_CHALLENGE[currentDay];
  const status = statusMap[currentDay] || 'pending';
  const isFriday = currentDay === 'fri';

  const statusCopy = {
    pending:     'Awaiting kickoff',
    voting:      'Tribes are picking competitors',
    in_progress: isFriday ? 'Obstacle course in progress' : 'Challenge in progress',
    complete:    'Challenge complete'
  };

  return {
    day: currentDay,
    name: ch.name,
    icon: ch.icon,
    status: status,
    statusLabel: statusCopy[status] || statusCopy.pending,
    isFinale: isFriday
  };
}

function buildYesterday_(day, results, eliminations, members) {
  const ch = DAY_TO_CHALLENGE[day];
  if (!ch) return null;

  // Find winner (placement = 1)
  const winnerResult = results.filter(function(r) { return r.day === day && r.placement === 1; })[0];
  let winner = null;
  if (winnerResult) {
    const tribe = getTribeById(winnerResult.tribe_id);
    if (tribe) winner = { id: tribe.id, name: tribe.name, color_hex: tribe.color_hex };
  }

  // Find eliminated members
  const elim = eliminations.filter(function(e) { return e.day === day; })[0];
  let eliminated = [];
  if (elim) {
    elim.member_ids.forEach(function(id) {
      const m = members.filter(function(mm) { return mm.id === id; })[0];
      if (m) {
        const tribe = getTribeById(m.tribe_id);
        eliminated.push({
          name: m.name,
          tribe_id: m.tribe_id,
          tribe_name: tribe ? tribe.name : '',
          tribe_color_hex: tribe ? tribe.color_hex : '#888'
        });
      }
    });
  }

  // If no data exists yet (challenge hasn't been resolved), return null
  if (!winner && eliminated.length === 0) return null;

  return {
    day: day,
    name: ch.name,
    icon: ch.icon,
    winner: winner,
    eliminated: eliminated
  };
}

// ----------------------------------------------------------------------------
// ADMIN: SIMULATION CONTROLS (Phase 2)
// Used during dry runs to manually advance the contest state.
// ----------------------------------------------------------------------------

/** Set the simulated current day. Pass 'pre','mon','tue','wed','thu','fri','post', or null to clear. */
function setSimDay(day) {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  if (!day || day === 'auto') {
    props.deleteProperty('SIM_DAY');
  } else if (['pre','mon','tue','wed','thu','fri','post'].indexOf(day) !== -1) {
    props.setProperty('SIM_DAY', day);
  } else {
    throw new Error('Invalid sim day: ' + day);
  }
  logAdminAction(adminEmail, 'setSimDay', String(day));
  return { ok: true, simDay: day || null };
}

function getSimDay() {
  return PropertiesService.getScriptProperties().getProperty('SIM_DAY') || null;
}

/**
 * Set placements for a given day. Pass an array like:
 *   [{tribe_id: 'marlin', placement: 1}, {tribe_id: 'bone', placement: 2}, ...]
 * Replaces any existing results for that day.
 */
function setChallengeResults(day, placements) {
  const adminEmail = requireAdmin_();
  if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
  if (!Array.isArray(placements)) throw new Error('Placements must be an array.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB.CHALLENGE_RESULTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dayCol = headers.indexOf('day');

  // Remove existing rows for this day (delete from bottom up to keep indices stable)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][dayCol] === day) sheet.deleteRow(i + 1);
  }

  // Append new rows
  placements.forEach(function(p) {
    if (!p.tribe_id || !p.placement) return;
    sheet.appendRow([day, p.tribe_id, parseInt(p.placement, 10), '', adminEmail, new Date().toISOString()]);
  });

  // Mark the challenge as complete
  setChallengeStatus_(day, 'complete');

  logAdminAction(adminEmail, 'setChallengeResults', day + ': ' + placements.length + ' placements');
  return { ok: true };
}

/**
 * Set eliminations for a given day. memberIds is an array of member ids.
 * Updates Members.status to 'jury' for those members.
 */
function setEliminations(day, memberIds) {
  const adminEmail = requireAdmin_();
  if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
  if (!Array.isArray(memberIds)) throw new Error('memberIds must be an array.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Update EliminationResults — replace any existing for this day
  const erSheet = ss.getSheetByName(TAB.ELIMINATION_RESULTS);
  const erData = erSheet.getDataRange().getValues();
  const erHeaders = erData[0];
  const erDayCol = erHeaders.indexOf('day');

  for (let i = erData.length - 1; i >= 1; i--) {
    if (erData[i][erDayCol] === day) erSheet.deleteRow(i + 1);
  }

  if (memberIds.length > 0) {
    // Group by tribe
    const members = readMembers_();
    const byTribe = {};
    memberIds.forEach(function(id) {
      const m = members.filter(function(mm) { return mm.id === id; })[0];
      if (!m) return;
      if (!byTribe[m.tribe_id]) byTribe[m.tribe_id] = [];
      byTribe[m.tribe_id].push(id);
    });
    Object.keys(byTribe).forEach(function(tribeId) {
      erSheet.appendRow([day, tribeId, byTribe[tribeId].join(','), new Date().toISOString()]);
    });
  }

  // Update each member's status to 'jury'
  const memSheet = ss.getSheetByName(TAB.MEMBERS);
  const memData = memSheet.getDataRange().getValues();
  const memHeaders = memData[0];
  const idCol = memHeaders.indexOf('id');
  const statusCol = memHeaders.indexOf('status');
  for (let i = 1; i < memData.length; i++) {
    if (memberIds.indexOf(memData[i][idCol]) !== -1) {
      memSheet.getRange(i + 1, statusCol + 1).setValue('jury');
    }
  }

  logAdminAction(adminEmail, 'setEliminations', day + ': ' + memberIds.length + ' eliminated');
  return { ok: true };
}

function setChallengeStatus_(day, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.CHALLENGES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dayCol = headers.indexOf('day');
  const statusCol = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (data[i][dayCol] === day) {
      sheet.getRange(i + 1, statusCol + 1).setValue(status);
      return;
    }
  }
}

function setChallengeStatus(day, status) {
  requireAdmin_();
  if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
  const validStatuses = ['pending','voting','in_progress','complete'];
  if (validStatuses.indexOf(status) === -1) throw new Error('Invalid status: ' + status);
  setChallengeStatus_(day, status);
  return { ok: true };
}

/**
 * Resets ALL contest progress — clears results, eliminations, restores all
 * members to active. Does NOT touch the roster itself.
 */
function resetContestProgress() {
  const adminEmail = requireAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clear ChallengeResults
  const cr = ss.getSheetByName(TAB.CHALLENGE_RESULTS);
  if (cr.getLastRow() > 1) cr.getRange(2, 1, cr.getLastRow() - 1, cr.getLastColumn()).clearContent();

  // Clear EliminationResults
  const er = ss.getSheetByName(TAB.ELIMINATION_RESULTS);
  if (er.getLastRow() > 1) er.getRange(2, 1, er.getLastRow() - 1, er.getLastColumn()).clearContent();

  // Reset all members to active
  const mem = ss.getSheetByName(TAB.MEMBERS);
  const memData = mem.getDataRange().getValues();
  const memHeaders = memData[0];
  const statusCol = memHeaders.indexOf('status');
  for (let i = 1; i < memData.length; i++) {
    if (memData[i][0]) mem.getRange(i + 1, statusCol + 1).setValue('active');
  }

  // Reset all challenge statuses to pending
  const ch = ss.getSheetByName(TAB.CHALLENGES);
  const chData = ch.getDataRange().getValues();
  const chHeaders = chData[0];
  const chStatusCol = chHeaders.indexOf('status');
  for (let i = 1; i < chData.length; i++) {
    if (chData[i][0]) ch.getRange(i + 1, chStatusCol + 1).setValue('pending');
  }

  // Clear sim day
  PropertiesService.getScriptProperties().deleteProperty('SIM_DAY');

  logAdminAction(adminEmail, 'resetContestProgress', 'All progress cleared.');
  return { ok: true };
}

/**
 * Returns active members for a given tribe — used by admin elimination picker.
 */
function getActiveMembersByTribe(tribeId) {
  requireAdmin_();
  return readMembers_().filter(function(m) {
    return m.tribe_id === tribeId && m.status === 'active';
  });
}

// ============================================================================
// PHASE 3 — TRIBE DASHBOARD STATE
// ============================================================================

/**
 * Returns true if the active user is an admin OR an active member of the tribe.
 * Used to gate tribe-specific reads (chat, election state, etc.).
 */
function requireTribeMember_(tribeId) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Not signed in.');
  if (ADMIN_EMAILS.indexOf(email) !== -1) return { email: email, isAdmin: true };
  const member = findMemberByEmail(email);
  if (!member) throw new Error('You are not on the roster.');
  if (member.status !== 'active') throw new Error('You are no longer an active member.');
  if (tribeId && member.tribe_id !== tribeId) throw new Error('Not a member of this tribe.');
  return { email: email, member: member, isAdmin: false };
}

/**
 * Master endpoint for the Tribe Dashboard.
 * Returns everything one tribe member needs: identity, roster, election state,
 * council state, chat messages, and tribe stats.
 */
function getTribeState() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = findMemberByEmail(email);
  if (!me) throw new Error('You are not on the roster.');

  const tribe = getTribeById(me.tribe_id);
  if (!tribe) throw new Error('Tribe not found.');

  const allMembers = readMembers_();
  const tribeMembers = allMembers.filter(function(m) { return m.tribe_id === tribe.id; });

  // All active members are eligible candidates. Already-competed members are
  // tagged in the UI via previousCompetitions (e.g. "✓ competed Mon") so tribes
  // can self-prefer fresh members but use repeats when needed.
  const competedIds = readCompetedMemberIds_();
  const competitionDaysByMember = readCompetitionDaysByMember_();
  const eligibleCandidates = tribeMembers.filter(function(m) {
    return m.status === 'active';
  });

  const props = PropertiesService.getScriptProperties();
  const electionDay = props.getProperty('ELECTION_OPEN_FOR_DAY') || '';
  const councilDay = props.getProperty('COUNCIL_OPEN_FOR_DAY') || '';
  const councilTribe = props.getProperty('COUNCIL_TARGET_TRIBE') || '';

  // Election state
  let election = { open: false, day: '', myVotes: [], candidates: [], lockedCompetitors: [], totalVoters: 0 };
  if (electionDay) {
    election.open = true;
    election.day = electionDay;
    election.candidates = eligibleCandidates.map(function(m) {
      return {
        id: m.id,
        name: m.name,
        previousCompetitions: competitionDaysByMember[m.id] || []
      };
    });
    election.myVotes = readMyElectionVotes_(tribe.id, electionDay, email);
    election.totalVoters = tribeMembers.filter(function(m) { return m.status === 'active'; }).length;
  }
  // Locked competitors (always show if any are locked, regardless of election open state)
  election.lockedCompetitors = readLockedCompetitors_(tribe.id);

  // Council state
  let council = {
    open: false,
    isOurTribe: false,
    day: '',
    candidates: [],
    myVotes: [],
    hasVoted: false,
    eliminatedNames: []
  };
  if (councilDay && councilTribe === tribe.id) {
    council.open = true;
    council.isOurTribe = true;
    council.day = councilDay;
    council.candidates = tribeMembers
      .filter(function(m) { return m.status === 'active'; })
      .map(function(m) { return { id: m.id, name: m.name }; });
    council.myVotes = readMyCouncilVotes_(tribe.id, councilDay, email);
    council.hasVoted = council.myVotes.length > 0;
  }
  // Recent eliminations from our tribe (always show)
  council.eliminatedNames = readRecentEliminationsForTribe_(tribe.id, allMembers);

  // Tribe stats — compute the same way Contest dashboard does
  const results = readChallengeResults_();
  let totalPoints = 0;
  results.forEach(function(r) {
    if (r.tribe_id !== tribe.id) return;
    totalPoints += pointsForPlacement_(r.placement, r.day);
  });

  // Compute tribe rank
  const tribeStats = TRIBES.map(function(t) {
    let pts = 0;
    results.forEach(function(r) {
      if (r.tribe_id === t.id) pts += pointsForPlacement_(r.placement, r.day);
    });
    return { id: t.id, points: pts };
  }).sort(function(a, b) { return b.points - a.points; });
  let myRank = 1;
  for (let i = 0; i < tribeStats.length; i++) {
    if (tribeStats[i].id === tribe.id) { myRank = i + 1; break; }
  }

  // Chat messages (last 50)
  const chat = readTribeChat_(tribe.id, 50);

  return {
    now: new Date().toISOString(),
    me: {
      id: me.id,
      name: me.name,
      email: me.email,
      tribe_id: me.tribe_id,
      status: me.status
    },
    tribe: {
      id: tribe.id,
      name: tribe.name,
      color_hex: tribe.color_hex,
      color_label: tribe.color_label,
      members: tribeMembers.map(function(m) {
        return {
          id: m.id, name: m.name, email: m.email,
          status: m.status,
          competed: competedIds.indexOf(m.id) !== -1
        };
      }),
      activeCount: tribeMembers.filter(function(m) { return m.status === 'active'; }).length,
      juryCount: tribeMembers.filter(function(m) { return m.status === 'jury'; }).length,
      points: totalPoints,
      rank: myRank
    },
    currentDay: getCurrentDay_(),
    election: election,
    council: council,
    chat: chat
  };
}

// ----------------------------------------------------------------------------
// HELPERS — sheet readers specific to Phase 3
// ----------------------------------------------------------------------------
function readCompetedMemberIds_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.COMPETITORS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const idCol = h.indexOf('member_id');
  const ids = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol]) ids.push(data[i][idCol]);
  }
  return ids;
}

// Returns { member_id: ['mon', 'wed', ...] } — which days each member already competed.
// Used by the candidate tag UI so tribes can see who's already played.
function readCompetitionDaysByMember_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.COMPETITORS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  const h = data[0];
  const idCol = h.indexOf('member_id');
  const dayCol = h.indexOf('day');
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const id = data[i][idCol];
    const day = data[i][dayCol];
    if (!id || !day) continue;
    if (!out[id]) out[id] = [];
    if (out[id].indexOf(day) === -1) out[id].push(day);
  }
  return out;
}

function readLockedCompetitors_(tribeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.COMPETITORS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const idCol = h.indexOf('member_id');
  const nameCol = h.indexOf('member_name');
  const lockedAtCol = h.indexOf('locked_at');
  // Group by day, return only rows for this tribe
  const byDay = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    const day = data[i][dayCol];
    if (!byDay[day]) byDay[day] = [];
    byDay[day].push({ id: data[i][idCol], name: data[i][nameCol], lockedAt: data[i][lockedAtCol] });
  }
  // Return as array sorted by day
  const out = [];
  DAYS_ORDER.forEach(function(d) {
    if (byDay[d]) out.push({ day: d, members: byDay[d] });
  });
  return out;
}

function readMyElectionVotes_(tribeId, day, voterEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.SELECTIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');
  const candCol = h.indexOf('candidate_email');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    if (data[i][dayCol] !== day) continue;
    if ((data[i][voterCol] || '').toLowerCase() !== voterEmail.toLowerCase()) continue;
    out.push(data[i][candCol]);
  }
  return out;
}

function readMyCouncilVotes_(tribeId, day, voterEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');
  const targetCol = h.indexOf('target_email');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    if (data[i][dayCol] !== day) continue;
    if ((data[i][voterCol] || '').toLowerCase() !== voterEmail.toLowerCase()) continue;
    out.push(data[i][targetCol]);
  }
  return out;
}

function readRecentEliminationsForTribe_(tribeId, allMembers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATION_RESULTS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const dayCol = h.indexOf('day');
  const tribeCol = h.indexOf('tribe_id');
  const idsCol = h.indexOf('eliminated_member_ids');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    const ids = (data[i][idsCol] || '').toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    ids.forEach(function(id) {
      const m = allMembers.filter(function(mm) { return mm.id === id; })[0];
      if (m) out.push({ name: m.name, day: data[i][dayCol] });
    });
  }
  return out;
}

function readTribeChat_(tribeId, limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.TRIBE_CHAT);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const emailCol = h.indexOf('member_email');
  const nameCol = h.indexOf('member_name');
  const msgCol = h.indexOf('message');
  const tsCol = h.indexOf('timestamp');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    out.push({
      email: data[i][emailCol],
      name: data[i][nameCol],
      message: data[i][msgCol],
      timestamp: data[i][tsCol]
    });
  }
  // Take last N
  return out.slice(-limit);
}

// ============================================================================
// PHASE 3 — ELECTION (vote for tomorrow's competitors)
// ============================================================================

/**
 * Submit (or replace) the caller's election votes for the given day.
 * candidateIds is an array of 1-2 member ids (caller can vote for fewer than 2).
 */
function submitElectionVotes(day, candidateIds) {
  const ctx = requireTribeMember_();
  if (ctx.isAdmin && !ctx.member) {
    throw new Error('Admins must have a member entry to vote in elections.');
  }
  const me = ctx.member;
  if (!me) throw new Error('Not a tribe member.');

  const props = PropertiesService.getScriptProperties();
  const openDay = props.getProperty('ELECTION_OPEN_FOR_DAY');
  if (!openDay || openDay !== day) throw new Error('Election is not open for this day.');

  if (!Array.isArray(candidateIds)) throw new Error('candidateIds must be an array.');
  if (candidateIds.length > 2) throw new Error('You may cast at most 2 votes.');

  // Resolve member ids → emails (Selections stores candidate_email)
  const allMembers = readMembers_();
  const memberById = {};
  allMembers.forEach(function(m) { memberById[m.id] = m; });

  // Validate each candidate is an active tribemate. Already-competed members
  // ARE eligible — tribes can re-elect anyone if they're short on fresh members.
  const candEmails = [];
  candidateIds.forEach(function(id) {
    const cand = memberById[id];
    if (!cand) throw new Error('Unknown candidate.');
    if (cand.tribe_id !== me.tribe_id) throw new Error('Candidate is not in your tribe.');
    if (cand.status !== 'active') throw new Error('Candidate is no longer active.');
    candEmails.push(cand.email);
  });
  // No duplicates
  if (candEmails.length === 2 && candEmails[0] === candEmails[1]) {
    throw new Error('Your two votes must be for different members.');
  }

  // Replace any existing votes from this voter for this tribe/day
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.SELECTIONS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][tribeCol] === me.tribe_id &&
        data[i][dayCol] === day &&
        (data[i][voterCol] || '').toLowerCase() === me.email.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  candEmails.forEach(function(ce) {
    sheet.appendRow([me.tribe_id, day, me.email, ce, new Date().toISOString()]);
  });

  return { ok: true, votes: candEmails.length };
}

/**
 * Admin: closes the election window AND locks in the top 2 vote-getters
 * per tribe as the day's competitors. If a tribe has ≤2 active candidates,
 * those candidates are auto-locked without needing votes.
 * Ties for the second slot are broken randomly.
 *
 * NOTE: All active members are eligible. Already-competed members can be
 * re-elected if a tribe is short on fresh participants.
 */
function closeElectionAndLock(day) {
  const adminEmail = requireAdmin_();
  if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allMembers = readMembers_();

  // Tally votes per tribe
  const selSheet = ss.getSheetByName(TAB.SELECTIONS);
  const selData = selSheet.getDataRange().getValues();
  const selH = selData[0];
  const sTribeCol = selH.indexOf('tribe_id');
  const sDayCol = selH.indexOf('day');
  const sCandCol = selH.indexOf('candidate_email');

  // tallies[tribeId][candidateEmail] = count
  const tallies = {};
  for (let i = 1; i < selData.length; i++) {
    if (selData[i][sDayCol] !== day) continue;
    const tid = selData[i][sTribeCol];
    const cemail = selData[i][sCandCol];
    if (!tallies[tid]) tallies[tid] = {};
    tallies[tid][cemail] = (tallies[tid][cemail] || 0) + 1;
  }

  // For each active tribe, pick top 2
  const compSheet = ss.getSheetByName(TAB.COMPETITORS);
  const tribesPicked = [];

  TRIBES.forEach(function(t) {
    // All active members are eligible (already-competed members can be re-elected)
    const tribeMembers = allMembers.filter(function(m) {
      return m.tribe_id === t.id && m.status === 'active';
    });
    if (tribeMembers.length === 0) return;

    let chosen = [];

    if (tribeMembers.length <= 2) {
      // Auto-lock all eligible
      chosen = tribeMembers.slice();
    } else {
      // Tally for this tribe's candidates
      const tallyForTribe = tallies[t.id] || {};
      // Score each member
      const scored = tribeMembers.map(function(m) {
        return { member: m, votes: tallyForTribe[m.email] || 0, rand: Math.random() };
      });
      // Sort: votes desc, then random for tie-break
      scored.sort(function(a, b) {
        if (b.votes !== a.votes) return b.votes - a.votes;
        return a.rand - b.rand;
      });
      chosen = [scored[0].member, scored[1].member];
    }

    // Remove any existing competitors for this tribe/day
    const compData = compSheet.getDataRange().getValues();
    const compH = compData[0];
    const cTribeCol = compH.indexOf('tribe_id');
    const cDayCol = compH.indexOf('day');
    for (let i = compData.length - 1; i >= 1; i--) {
      if (compData[i][cTribeCol] === t.id && compData[i][cDayCol] === day) {
        compSheet.deleteRow(i + 1);
      }
    }
    // Write 2 rows
    const ts = new Date().toISOString();
    chosen.forEach(function(m) {
      compSheet.appendRow([t.id, day, m.id, m.name, ts]);
    });
    tribesPicked.push({ tribe: t.name, names: chosen.map(function(m) { return m.name; }) });
  });

  // Close the election window
  PropertiesService.getScriptProperties().deleteProperty('ELECTION_OPEN_FOR_DAY');

  logAdminAction(adminEmail, 'closeElectionAndLock', day + ': ' + JSON.stringify(tribesPicked));
  return { ok: true, tribesPicked: tribesPicked };
}

// ============================================================================
// PHASE 3 — TRIBAL COUNCIL (vote out)
// ============================================================================

/**
 * Submit (or replace) the caller's council votes for the given day.
 * targetIds is an array of 1-2 member ids (caller can cast 1 or 2 votes).
 */
function submitCouncilVotes(day, targetIds) {
  const ctx = requireTribeMember_();
  const me = ctx.member;
  if (!me) throw new Error('Not a tribe member.');

  const props = PropertiesService.getScriptProperties();
  const openDay = props.getProperty('COUNCIL_OPEN_FOR_DAY');
  const targetTribe = props.getProperty('COUNCIL_TARGET_TRIBE');
  if (!openDay || openDay !== day) throw new Error('Council is not open for this day.');
  if (targetTribe !== me.tribe_id) throw new Error('Your tribe is not at council tonight.');

  if (!Array.isArray(targetIds)) throw new Error('targetIds must be an array.');
  if (targetIds.length > 2) throw new Error('You may cast at most 2 votes.');
  if (targetIds.length === 2 && targetIds[0] === targetIds[1]) {
    throw new Error('Your two votes must be for different members.');
  }

  // Validate targets are active members of our tribe
  const allMembers = readMembers_();
  const memberById = {};
  allMembers.forEach(function(m) { memberById[m.id] = m; });
  const targetEmails = [];
  targetIds.forEach(function(id) {
    const t = memberById[id];
    if (!t) throw new Error('Unknown target.');
    if (t.tribe_id !== me.tribe_id) throw new Error('Target is not in your tribe.');
    if (t.status !== 'active') throw new Error(t.name + ' has already been eliminated.');
    targetEmails.push(t.email);
  });

  // Replace any existing votes from this voter
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATIONS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][tribeCol] === me.tribe_id &&
        data[i][dayCol] === day &&
        (data[i][voterCol] || '').toLowerCase() === me.email.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  targetEmails.forEach(function(te) {
    sheet.appendRow([me.tribe_id, day, me.email, te, new Date().toISOString()]);
  });

  return { ok: true, votes: targetEmails.length };
}

/**
 * Admin: closes council, tallies votes, eliminates the top 2 vote-getters.
 * Tie for second slot → random among tied.
 * Updates Members.status to 'jury' for those eliminated.
 */
function closeCouncilAndEliminate(day, tribeId) {
  const adminEmail = requireAdmin_();
  if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
  if (!getTribeById(tribeId)) throw new Error('Invalid tribe: ' + tribeId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allMembers = readMembers_();
  const tribeMembers = allMembers.filter(function(m) {
    return m.tribe_id === tribeId && m.status === 'active';
  });

  // Tally votes
  const elimSheet = ss.getSheetByName(TAB.ELIMINATIONS);
  const elimData = elimSheet.getDataRange().getValues();
  const eH = elimData[0];
  const eTribeCol = eH.indexOf('tribe_id');
  const eDayCol = eH.indexOf('day');
  const eTargetCol = eH.indexOf('target_email');

  const tally = {}; // email -> count
  for (let i = 1; i < elimData.length; i++) {
    if (elimData[i][eTribeCol] === tribeId && elimData[i][eDayCol] === day) {
      const t = elimData[i][eTargetCol];
      tally[t] = (tally[t] || 0) + 1;
    }
  }

  let toEliminate = [];
  if (tribeMembers.length <= 2) {
    // Degenerate case — auto-eliminate all active members (will trigger absorption)
    toEliminate = tribeMembers.slice();
  } else {
    const scored = tribeMembers.map(function(m) {
      return { member: m, votes: tally[m.email] || 0, rand: Math.random() };
    });
    scored.sort(function(a, b) {
      if (b.votes !== a.votes) return b.votes - a.votes;
      return a.rand - b.rand;
    });
    toEliminate = [scored[0].member, scored[1].member];
  }

  // Write EliminationResults
  const erSheet = ss.getSheetByName(TAB.ELIMINATION_RESULTS);
  const erData = erSheet.getDataRange().getValues();
  const erH = erData[0];
  const erDayCol = erH.indexOf('day');
  const erTribeCol = erH.indexOf('tribe_id');
  for (let i = erData.length - 1; i >= 1; i--) {
    if (erData[i][erDayCol] === day && erData[i][erTribeCol] === tribeId) {
      erSheet.deleteRow(i + 1);
    }
  }
  erSheet.appendRow([day, tribeId, toEliminate.map(function(m) { return m.id; }).join(','), new Date().toISOString()]);

  // Update Members.status → jury
  const memSheet = ss.getSheetByName(TAB.MEMBERS);
  const memData = memSheet.getDataRange().getValues();
  const mH = memData[0];
  const idCol = mH.indexOf('id');
  const statusCol = mH.indexOf('status');
  const eliminatedIds = toEliminate.map(function(m) { return m.id; });
  for (let i = 1; i < memData.length; i++) {
    if (eliminatedIds.indexOf(memData[i][idCol]) !== -1) {
      memSheet.getRange(i + 1, statusCol + 1).setValue('jury');
    }
  }

  // Close council window
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('COUNCIL_OPEN_FOR_DAY');
  props.deleteProperty('COUNCIL_TARGET_TRIBE');

  logAdminAction(adminEmail, 'closeCouncilAndEliminate',
    day + '/' + tribeId + ': ' + toEliminate.map(function(m) { return m.name; }).join(', '));

  // Phase 7: enqueue ceremony
  try {
    const tribe = getTribeById(tribeId);
    _enqueueCeremony_('council_reveal', {
      tribe_id: tribeId,
      tribe_name: tribe ? tribe.name : '',
      tribe_color_hex: tribe ? tribe.color_hex : '#888',
      eliminated: toEliminate.map(function(m) {
        return { name: m.name, tribe_color_hex: tribe ? tribe.color_hex : '#888' };
      }),
      day: day
    });
  } catch (e) { /* ceremony enqueue is best-effort */ }

  return { ok: true, eliminated: toEliminate.map(function(m) { return { id: m.id, name: m.name }; }) };
}

/**
 * Admin / jury only — get the full council vote tally for a given day/tribe.
 * Tribe members do NOT see this; only admin and jury.
 */
function getCouncilTally(day, tribeId) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const isAdmin = ADMIN_EMAILS.indexOf(email) !== -1;
  const member = findMemberByEmail(email);
  const isJury = member && member.status === 'jury';
  if (!isAdmin && !isJury) throw new Error('Not authorized to view council tally.');

  const allMembers = readMembers_();
  const tribeMembers = allMembers.filter(function(m) { return m.tribe_id === tribeId; });
  const emailToName = {};
  tribeMembers.forEach(function(m) { emailToName[m.email.toLowerCase()] = m.name; });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATIONS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const targetCol = h.indexOf('target_email');

  const tally = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] === tribeId && data[i][dayCol] === day) {
      const t = (data[i][targetCol] || '').toLowerCase();
      tally[t] = (tally[t] || 0) + 1;
    }
  }
  // Convert to array sorted by votes
  const out = [];
  Object.keys(tally).forEach(function(em) {
    out.push({ name: emailToName[em] || em, votes: tally[em] });
  });
  out.sort(function(a, b) { return b.votes - a.votes; });
  return out;
}

// ============================================================================
// PHASE 3 — CHAT & CONFESSIONALS
// ============================================================================

function postTribeChatMessage(text) {
  const ctx = requireTribeMember_();
  const me = ctx.member;
  if (!me) throw new Error('Only tribe members can post in tribe chat.');
  text = (text || '').toString().trim();
  if (!text) throw new Error('Message cannot be empty.');
  if (text.length > 500) throw new Error('Message is too long (max 500 chars).');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.TRIBE_CHAT);
  sheet.appendRow([me.tribe_id, me.email, me.name, text, new Date().toISOString()]);
  return { ok: true };
}

function submitConfessional(text) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = findMemberByEmail(email);
  if (!me) throw new Error('Only roster members can submit confessionals.');
  if (me.status !== 'active') throw new Error('Confessionals are for active castaways.');
  text = (text || '').toString().trim();
  if (!text) throw new Error('Confessional cannot be empty.');
  if (text.length > 200) throw new Error('Confessional is too long (max 200 chars).');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.CONFESSIONALS);
  sheet.appendRow([me.email, me.name, me.tribe_id, text, new Date().toISOString(), false]);
  return { ok: true };
}

function getRecentConfessionals(limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.CONFESSIONALS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const emailCol = h.indexOf('member_email');
  const nameCol = h.indexOf('member_name');
  const tribeCol = h.indexOf('tribe_id_at_time');
  const textCol = h.indexOf('text');
  const tsCol = h.indexOf('timestamp');
  const hiddenCol = h.indexOf('hidden');

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][hiddenCol] === true) continue;
    if (!data[i][textCol]) continue;
    const tribe = getTribeById(data[i][tribeCol]);
    out.push({
      name: data[i][nameCol],
      tribe_id: data[i][tribeCol],
      tribe_name: tribe ? tribe.name : '',
      tribe_color_hex: tribe ? tribe.color_hex : '#888',
      text: data[i][textCol],
      timestamp: data[i][tsCol]
    });
  }
  // Newest last; return last N
  return out.slice(-(limit || 30));
}

// ============================================================================
// PHASE 3 — WINDOW CONTROLS (admin-driven for Phase 3; auto in Phase 4)
// ============================================================================

function getWindowState() {
  const props = PropertiesService.getScriptProperties();
  return {
    election_day: props.getProperty('ELECTION_OPEN_FOR_DAY') || '',
    council_day: props.getProperty('COUNCIL_OPEN_FOR_DAY') || '',
    council_tribe: props.getProperty('COUNCIL_TARGET_TRIBE') || ''
  };
}

function setElectionWindow(day) {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  if (!day) {
    props.deleteProperty('ELECTION_OPEN_FOR_DAY');
  } else {
    if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
    props.setProperty('ELECTION_OPEN_FOR_DAY', day);
  }
  logAdminAction(adminEmail, 'setElectionWindow', String(day));
  return { ok: true };
}

function setCouncilWindow(day, tribeId) {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  if (!day || !tribeId) {
    props.deleteProperty('COUNCIL_OPEN_FOR_DAY');
    props.deleteProperty('COUNCIL_TARGET_TRIBE');
  } else {
    if (DAYS_ORDER.indexOf(day) === -1) throw new Error('Invalid day: ' + day);
    if (!getTribeById(tribeId)) throw new Error('Invalid tribe: ' + tribeId);
    props.setProperty('COUNCIL_OPEN_FOR_DAY', day);
    props.setProperty('COUNCIL_TARGET_TRIBE', tribeId);
  }
  logAdminAction(adminEmail, 'setCouncilWindow', day + '/' + tribeId);
  return { ok: true };
}

// ============================================================================
// PHASE 4 — ESCAPE ROOM EDITOR (admin)
// ============================================================================

function getEscapeRoomCluesAdmin() {
  requireAdmin_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_CLUES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const numCol = h.indexOf('clue_number');
  const releaseCol = h.indexOf('release_time');
  const titleCol = h.indexOf('title');
  const riddleCol = h.indexOf('riddle_text');
  const answerCol = h.indexOf('cipher_answer');
  const imageCol = h.indexOf('image_url');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][numCol]) continue;
    out.push({
      clue_number: parseInt(data[i][numCol], 10),
      release_time: (data[i][releaseCol] || '').toString(),
      title: data[i][titleCol] || '',
      riddle_text: data[i][riddleCol] || '',
      cipher_answer: data[i][answerCol] || '',
      image_url: data[i][imageCol] || ''
    });
  }
  return out.sort(function(a, b) { return a.clue_number - b.clue_number; });
}

function updateClue(clueNumber, fields) {
  const adminEmail = requireAdmin_();
  clueNumber = parseInt(clueNumber, 10);
  if (!clueNumber || clueNumber < 1) throw new Error('Invalid clue number');
  if (!fields || typeof fields !== 'object') throw new Error('Fields object required');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_CLUES);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const numCol = h.indexOf('clue_number');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][numCol], 10) === clueNumber) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('Clue ' + clueNumber + ' not found.');

  const sheetRow = rowIndex + 1;
  const allowed = ['title', 'release_time', 'riddle_text', 'cipher_answer', 'image_url'];
  const updated = [];
  allowed.forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(fields, key)) {
      const colIdx = h.indexOf(key);
      if (colIdx === -1) return;
      let val = fields[key];
      if (val === null || val === undefined) val = '';
      sheet.getRange(sheetRow, colIdx + 1).setValue(val);
      updated.push(key);
    }
  });

  logAdminAction(adminEmail, 'updateClue', 'Clue ' + clueNumber + ': ' + updated.join(','));
  return { ok: true, updated: updated };
}

// ============================================================================
// PHASE 4 — IDOL (Rammy & the Bear)
// ============================================================================

// Helper: read a single idol's state by id ('rammy' or 'bear')
function _readIdolRow_(idolId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const h = data[0];
  const idCol = h.indexOf('idol_id');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][idCol] || '').toString() === idolId) {
      return { row: data[i], rowIndex: i + 1, headers: h };
    }
  }
  return null;
}

function _idolStateFromRow_(idolId, idolRow, isAdmin) {
  if (!idolRow) {
    return { idol_id: idolId, idol_name: idolId, state: 'unset', hint_log: [], holder: null, secret_word: null };
  }
  const h = idolRow.headers;
  const row = idolRow.row;
  const nameCol = h.indexOf('idol_name');
  const wordCol = h.indexOf('secret_word');
  const hintCol = h.indexOf('hint_log');
  const foundByCol = h.indexOf('found_by_email');
  const claimedAtCol = h.indexOf('claimed_at');
  const playedAtCol = h.indexOf('played_at');

  const idolName = (row[nameCol] || idolId).toString();
  const secretWord = (row[wordCol] || '').toString().trim();
  const foundBy = (row[foundByCol] || '').toString().trim();
  const playedAt = (row[playedAtCol] || '').toString().trim();

  let state = 'unset';
  if (!secretWord) state = 'unset';
  else if (!foundBy) state = 'hidden';
  else if (!playedAt) state = 'claimed';
  else state = 'played';

  let holder = null;
  if (foundBy) {
    const member = findMemberByEmail(foundBy);
    if (member) {
      const tribe = getTribeById(member.tribe_id);
      holder = {
        member_name: member.name,
        member_email: isAdmin ? member.email : null,
        tribe_id: tribe ? tribe.id : null,
        tribe_name: tribe ? tribe.name : null,
        tribe_color_hex: tribe ? tribe.color_hex : null,
        claimed_at: row[claimedAtCol] || null
      };
    }
  }

  const hintLog = (row[hintCol] || '').toString().split('\n').filter(Boolean);

  return {
    idol_id: idolId,
    idol_name: idolName,
    state: state,
    secret_word: isAdmin ? secretWord : null,
    holder: holder,
    hint_log: hintLog,
    played_at: playedAt || null
  };
}

// Returns BOTH idols (Rammy and Bear) — the new Admin UI shows them side-by-side
function getIdolState() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const isAdmin = ADMIN_EMAILS.indexOf(email) !== -1;

  const rammy = _readIdolRow_('rammy');
  const bear = _readIdolRow_('bear');

  return {
    idols: [
      _idolStateFromRow_('rammy', rammy, isAdmin),
      _idolStateFromRow_('bear', bear, isAdmin)
    ]
  };
}

function setIdolSecret(idolId, word) {
  const adminEmail = requireAdmin_();
  if (idolId !== 'rammy' && idolId !== 'bear') throw new Error('Invalid idol id.');
  word = (word || '').toString().trim();
  if (!word) throw new Error('Secret word cannot be empty.');
  if (word.length < 3) throw new Error('Secret word must be at least 3 characters.');

  const idolRow = _readIdolRow_(idolId);
  if (!idolRow) throw new Error('Idol row not found. Re-run setupSheets() to seed.');
  const wordCol = idolRow.headers.indexOf('secret_word') + 1;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);
  sheet.getRange(idolRow.rowIndex, wordCol).setValue(word);
  logAdminAction(adminEmail, 'setIdolSecret', idolId + ' (set)');
  return { ok: true };
}

function addIdolHint(idolId, text) {
  const adminEmail = requireAdmin_();
  if (idolId !== 'rammy' && idolId !== 'bear') throw new Error('Invalid idol id.');
  text = (text || '').toString().trim();
  if (!text) throw new Error('Hint cannot be empty.');

  const idolRow = _readIdolRow_(idolId);
  if (!idolRow) throw new Error('Idol row not found.');
  const hintCol = idolRow.headers.indexOf('hint_log') + 1;
  const existing = (idolRow.row[hintCol - 1] || '').toString();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d h:mm a');
  const idolName = (idolRow.row[idolRow.headers.indexOf('idol_name')] || idolId).toString();
  const newLog = (existing ? existing + '\n' : '') + '[' + ts + ' · ' + idolName + '] ' + text;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);
  sheet.getRange(idolRow.rowIndex, hintCol).setValue(newLog);
  logAdminAction(adminEmail, 'addIdolHint', idolId + ': ' + text);
  return { ok: true };
}

// claimIdol now tries BOTH idols; whichever matches gets claimed.
// If both have the same word (admin error), Rammy is checked first.
function claimIdol(attemptWord) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('You must be signed in.');
  const member = findMemberByEmail(email);
  if (!member) throw new Error('Only roster members can claim an idol.');
  if (member.status !== 'active') throw new Error('Only active players can claim an idol.');

  attemptWord = (attemptWord || '').toString().trim();
  if (!attemptWord) throw new Error('Enter a word.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);

  // Try Rammy then Bear
  const tryIds = ['rammy', 'bear'];
  for (let i = 0; i < tryIds.length; i++) {
    const idolId = tryIds[i];
    const idolRow = _readIdolRow_(idolId);
    if (!idolRow) continue;
    const h = idolRow.headers;
    const wordCol = h.indexOf('secret_word');
    const foundByCol = h.indexOf('found_by_email');
    const claimedAtCol = h.indexOf('claimed_at');
    const nameCol = h.indexOf('idol_name');

    const secretWord = (idolRow.row[wordCol] || '').toString().trim();
    const currentHolder = (idolRow.row[foundByCol] || '').toString().trim();
    if (!secretWord) continue;
    if (currentHolder) continue;
    if (attemptWord.toLowerCase() !== secretWord.toLowerCase()) continue;

    // Match! Claim it.
    sheet.getRange(idolRow.rowIndex, foundByCol + 1).setValue(member.email);
    sheet.getRange(idolRow.rowIndex, claimedAtCol + 1).setValue(new Date().toISOString());

    const idolName = (idolRow.row[nameCol] || idolId).toString();
    logAdminAction('SYSTEM', 'idolClaimed', idolName + ' by ' + member.name + ' (' + member.email + ')');
    const tribe = getTribeById(member.tribe_id);

    try {
      _enqueueCeremony_('idol_claimed', {
        idol_name: idolName,
        member_name: member.name,
        tribe_name: tribe ? tribe.name : '',
        tribe_color_hex: tribe ? tribe.color_hex : '#888'
      });
    } catch (e) { /* best-effort */ }

    return {
      ok: true,
      idol_name: idolName,
      message: idolName + ' is yours.',
      tribe_name: tribe ? tribe.name : '',
      tribe_color_hex: tribe ? tribe.color_hex : '#888'
    };
  }

  throw new Error('That is not a correct idol word.');
}

function markIdolAsPlayed(idolId, day) {
  const adminEmail = requireAdmin_();
  if (idolId !== 'rammy' && idolId !== 'bear') throw new Error('Invalid idol id.');
  const idolRow = _readIdolRow_(idolId);
  if (!idolRow) throw new Error('Idol not found.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);
  const playedDayCol = idolRow.headers.indexOf('played_for_day') + 1;
  const playedAtCol = idolRow.headers.indexOf('played_at') + 1;
  sheet.getRange(idolRow.rowIndex, playedDayCol).setValue(day || '');
  sheet.getRange(idolRow.rowIndex, playedAtCol).setValue(new Date().toISOString());
  logAdminAction(adminEmail, 'markIdolAsPlayed', idolId + ' / ' + (day || ''));
  return { ok: true };
}

function resetIdol(idolId) {
  const adminEmail = requireAdmin_();
  if (idolId !== 'rammy' && idolId !== 'bear') throw new Error('Invalid idol id.');
  const idolRow = _readIdolRow_(idolId);
  if (!idolRow) throw new Error('Idol not found.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.IDOL_STATE);
  // Clear word, hint, foundBy, claimedAt, playedDay, playedAt — preserve idol_id and idol_name
  const h = idolRow.headers;
  const cols = ['secret_word', 'hint_log', 'found_by_email', 'claimed_at', 'played_for_day', 'played_at'];
  cols.forEach(function(name) {
    const idx = h.indexOf(name);
    if (idx >= 0) sheet.getRange(idolRow.rowIndex, idx + 1).setValue('');
  });
  logAdminAction(adminEmail, 'resetIdol', idolId);
  return { ok: true };
}

// ============================================================================
// PHASE 4 — ALL-TRIBES MIRROR (admin)
// ============================================================================

function computeElectionTallies_(tribeId, day, allMembers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.SELECTIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { voters: 0, results: [] };
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');
  const candCol = h.indexOf('candidate_email');

  const tally = {};
  const voters = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId || data[i][dayCol] !== day) continue;
    const ce = (data[i][candCol] || '').toString().toLowerCase();
    tally[ce] = (tally[ce] || 0) + 1;
    voters[(data[i][voterCol] || '').toString().toLowerCase()] = true;
  }

  const emailToName = {};
  allMembers.forEach(function(m) { emailToName[m.email.toLowerCase()] = m.name; });

  const results = [];
  Object.keys(tally).forEach(function(em) {
    results.push({ name: emailToName[em] || em, votes: tally[em] });
  });
  results.sort(function(a, b) { return b.votes - a.votes; });

  return { voters: Object.keys(voters).length, results: results };
}

function computeCouncilTallies_(tribeId, day, allMembers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ELIMINATIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { voters: 0, results: [] };
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const dayCol = h.indexOf('day');
  const voterCol = h.indexOf('voter_email');
  const targetCol = h.indexOf('target_email');

  const tally = {};
  const voters = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId || data[i][dayCol] !== day) continue;
    const te = (data[i][targetCol] || '').toString().toLowerCase();
    tally[te] = (tally[te] || 0) + 1;
    voters[(data[i][voterCol] || '').toString().toLowerCase()] = true;
  }

  const emailToName = {};
  allMembers.forEach(function(m) { emailToName[m.email.toLowerCase()] = m.name; });

  const results = [];
  Object.keys(tally).forEach(function(em) {
    results.push({ name: emailToName[em] || em, votes: tally[em] });
  });
  results.sort(function(a, b) { return b.votes - a.votes; });

  return { voters: Object.keys(voters).length, results: results };
}

function getAllTribesAdminView() {
  requireAdmin_();

  const allMembers = readMembers_();
  const competedIds = readCompetedMemberIds_();
  const results = readChallengeResults_();

  const tribePoints = {};
  results.forEach(function(r) {
    tribePoints[r.tribe_id] = (tribePoints[r.tribe_id] || 0) + pointsForPlacement_(r.placement, r.day);
  });

  const ranked = TRIBES.map(function(t) {
    return { id: t.id, points: tribePoints[t.id] || 0 };
  }).sort(function(a, b) { return b.points - a.points; });
  const rankById = {};
  ranked.forEach(function(t, idx) { rankById[t.id] = idx + 1; });

  const props = PropertiesService.getScriptProperties();
  const electionDay = props.getProperty('ELECTION_OPEN_FOR_DAY') || '';
  const councilDay = props.getProperty('COUNCIL_OPEN_FOR_DAY') || '';
  const councilTribe = props.getProperty('COUNCIL_TARGET_TRIBE') || '';

  const tribesData = TRIBES.map(function(tribe) {
    const tribeMembers = allMembers.filter(function(m) { return m.tribe_id === tribe.id; });

    const electionTallies = electionDay ? computeElectionTallies_(tribe.id, electionDay, allMembers) : null;
    const councilTallies = (councilDay && councilTribe === tribe.id)
      ? computeCouncilTallies_(tribe.id, councilDay, allMembers)
      : null;
    const chat = readTribeChat_(tribe.id, 25);
    const lockedCompetitors = readLockedCompetitors_(tribe.id);

    return {
      id: tribe.id,
      name: tribe.name,
      color_hex: tribe.color_hex,
      color_label: tribe.color_label,
      points: tribePoints[tribe.id] || 0,
      rank: rankById[tribe.id],
      members: tribeMembers.map(function(m) {
        return {
          id: m.id, name: m.name, email: m.email,
          status: m.status,
          competed: competedIds.indexOf(m.id) !== -1
        };
      }),
      activeCount: tribeMembers.filter(function(m) { return m.status === 'active'; }).length,
      juryCount: tribeMembers.filter(function(m) { return m.status === 'jury'; }).length,
      electionTallies: electionTallies,
      councilTallies: councilTallies,
      lockedCompetitors: lockedCompetitors,
      chat: chat
    };
  });

  return {
    tribes: tribesData,
    electionDay: electionDay,
    councilDay: councilDay,
    councilTribe: councilTribe
  };
}

// ============================================================================
// PHASE 5 — ESCAPE ROOM PLAYER FLOW
// ============================================================================

const ESCAPE_ROOM_COOLDOWN_MS = 30 * 1000;

/**
 * Read all clues (no cipher answers — those stay server-side).
 */
function readClues_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_CLUES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const numCol = h.indexOf('clue_number');
  const releaseCol = h.indexOf('release_time');
  const titleCol = h.indexOf('title');
  const riddleCol = h.indexOf('riddle_text');
  const answerCol = h.indexOf('cipher_answer');
  const imageCol = h.indexOf('image_url');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][numCol]) continue;
    out.push({
      clue_number: parseInt(data[i][numCol], 10),
      release_time: (data[i][releaseCol] || '').toString(),
      title: data[i][titleCol] || '',
      riddle_text: data[i][riddleCol] || '',
      cipher_answer: data[i][answerCol] || '',
      image_url: data[i][imageCol] || ''
    });
  }
  return out.sort(function(a, b) { return a.clue_number - b.clue_number; });
}

/**
 * Read this tribe's progress: solved clues with timestamps and attempt counts.
 */
function readTribeProgress_(tribeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_PROGRESS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const numCol = h.indexOf('clue_number');
  const solvedCol = h.indexOf('solved_at');
  const attemptCol = h.indexOf('attempt_count');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] !== tribeId) continue;
    out.push({
      clue_number: parseInt(data[i][numCol], 10),
      solved_at: data[i][solvedCol] || '',
      attempt_count: parseInt(data[i][attemptCol], 10) || 0,
      _row: i + 1
    });
  }
  return out;
}

function findOrCreateProgressRow_(tribeId, clueNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_PROGRESS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const tribeCol = h.indexOf('tribe_id');
  const numCol = h.indexOf('clue_number');
  for (let i = 1; i < data.length; i++) {
    if (data[i][tribeCol] === tribeId && parseInt(data[i][numCol], 10) === clueNumber) {
      return { sheet: sheet, headers: h, rowIndex: i + 1 };
    }
  }
  // Create new row
  sheet.appendRow([tribeId, clueNumber, '', 0]);
  return { sheet: sheet, headers: h, rowIndex: sheet.getLastRow() };
}

/**
 * Get escape room state for the calling user's tribe.
 * Returns: clues array (no answers), progress, currentClue (if any), unlockTime, finishedAt
 */
function getEscapeRoomState() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Not signed in.');

  // Identify tribe — admins can pass tribeId; everyone else gets their own
  const member = findMemberByEmail(email);
  const isAdmin = ADMIN_EMAILS.indexOf(email) !== -1;
  if (!member && !isAdmin) throw new Error('You are not on the roster.');
  if (!member) throw new Error('No tribe associated with this user.');

  return _buildEscapeRoomStateForTribe(member.tribe_id, email);
}

function _buildEscapeRoomStateForTribe(tribeId, email) {
  const allClues = readClues_();
  const progress = readTribeProgress_(tribeId);
  const now = new Date();

  // Build progress lookup
  const progressByClue = {};
  progress.forEach(function(p) { progressByClue[p.clue_number] = p; });

  // Determine the highest solved clue
  let highestSolved = 0;
  progress.forEach(function(p) {
    if (p.solved_at && p.clue_number > highestSolved) highestSolved = p.clue_number;
  });

  // Are all clues solved?
  const finishedAt = (highestSolved >= allClues.length && allClues.length > 0)
    ? progressByClue[allClues.length].solved_at
    : null;

  // Public clue list (no answers exposed)
  const publicClues = allClues.map(function(c) {
    const p = progressByClue[c.clue_number];
    const releaseTime = c.release_time ? new Date(c.release_time) : null;
    const released = releaseTime ? now >= releaseTime : true;
    return {
      clue_number: c.clue_number,
      title: c.title,
      release_time: c.release_time,
      released: released,
      solved: !!(p && p.solved_at),
      solved_at: p ? p.solved_at : null
    };
  });

  // Current clue: lowest unsolved clue that is also released AND whose predecessor is solved
  let currentClue = null;
  let nextUnlockTime = null;
  for (let i = 0; i < allClues.length; i++) {
    const c = allClues[i];
    const p = progressByClue[c.clue_number];
    const isSolved = p && p.solved_at;
    if (isSolved) continue;

    const prereqMet = (i === 0) || (progressByClue[allClues[i - 1].clue_number] && progressByClue[allClues[i - 1].clue_number].solved_at);
    if (!prereqMet) break;

    const releaseTime = c.release_time ? new Date(c.release_time) : null;
    if (releaseTime && now < releaseTime) {
      // Time-gated: prerequisite solved but next clue not yet released
      nextUnlockTime = c.release_time;
      break;
    }

    // This is the active clue
    currentClue = {
      clue_number: c.clue_number,
      title: c.title,
      riddle_text: c.riddle_text,
      image_url: c.image_url,
      total: allClues.length
    };
    break;
  }

  // Cooldown check
  let cooldownUntil = null;
  if (currentClue) {
    const cooldownKey = 'ER_COOLDOWN__' + tribeId + '__' + currentClue.clue_number + '__' + (email || 'anon');
    const stored = PropertiesService.getScriptProperties().getProperty(cooldownKey);
    if (stored) {
      const t = parseInt(stored, 10);
      if (!isNaN(t) && t > Date.now()) {
        cooldownUntil = new Date(t).toISOString();
      }
    }
  }

  return {
    now: now.toISOString(),
    tribe_id: tribeId,
    total_clues: allClues.length,
    solved_count: progress.filter(function(p) { return !!p.solved_at; }).length,
    clues: publicClues,
    currentClue: currentClue,
    nextUnlockTime: nextUnlockTime,
    finishedAt: finishedAt,
    cooldownUntil: cooldownUntil
  };
}

/**
 * Submit an attempted answer. Server-side validation, no answer leak.
 * Returns: { correct, message, cooldownUntil?, finished? }
 */
function submitClueAttempt(clueNumber, attempt) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Not signed in.');
  const member = findMemberByEmail(email);
  if (!member) throw new Error('Only tribe members can submit clue answers.');
  if (member.status !== 'active') throw new Error('Jury members cannot submit answers.');

  clueNumber = parseInt(clueNumber, 10);
  attempt = (attempt || '').toString().trim();
  if (!attempt) throw new Error('Enter an answer.');
  if (!clueNumber || clueNumber < 1) throw new Error('Invalid clue number.');

  const tribeId = member.tribe_id;

  // Cooldown check (per-tribe-per-clue, not per-user — anyone in the tribe shares)
  const sharedCooldownKey = 'ER_COOLDOWN__' + tribeId + '__' + clueNumber + '__shared';
  const stored = PropertiesService.getScriptProperties().getProperty(sharedCooldownKey);
  if (stored) {
    const t = parseInt(stored, 10);
    if (!isNaN(t) && t > Date.now()) {
      const remaining = Math.ceil((t - Date.now()) / 1000);
      throw new Error('Slow down — try again in ' + remaining + ' seconds.');
    }
  }

  // Verify this clue is the active one for the tribe (prevents skipping ahead)
  const state = _buildEscapeRoomStateForTribe(tribeId, email);
  if (!state.currentClue || state.currentClue.clue_number !== clueNumber) {
    if (state.finishedAt) {
      return { correct: false, message: 'Your tribe has already finished the escape room.', alreadyDone: true };
    }
    if (state.nextUnlockTime) {
      return { correct: false, message: 'Next clue unlocks at ' + _formatLocalTime_(state.nextUnlockTime) + '.' };
    }
    throw new Error('That clue is not currently active for your tribe.');
  }

  // Look up the actual answer
  const allClues = readClues_();
  const clue = allClues.filter(function(c) { return c.clue_number === clueNumber; })[0];
  if (!clue) throw new Error('Clue not found.');

  const expected = (clue.cipher_answer || '').toString().trim();
  const isCorrect = expected.toLowerCase() === attempt.toLowerCase();

  // Update progress row
  const ref = findOrCreateProgressRow_(tribeId, clueNumber);
  const attemptColIdx = ref.headers.indexOf('attempt_count') + 1;
  const solvedColIdx = ref.headers.indexOf('solved_at') + 1;
  const currentAttempts = parseInt(ref.sheet.getRange(ref.rowIndex, attemptColIdx).getValue(), 10) || 0;
  ref.sheet.getRange(ref.rowIndex, attemptColIdx).setValue(currentAttempts + 1);

  if (isCorrect) {
    const ts = new Date().toISOString();
    ref.sheet.getRange(ref.rowIndex, solvedColIdx).setValue(ts);
    // Clear any cooldown
    PropertiesService.getScriptProperties().deleteProperty(sharedCooldownKey);

    // Check if this was the last clue
    const isFinal = clueNumber >= allClues.length;
    const newState = _buildEscapeRoomStateForTribe(tribeId, email);
    return {
      correct: true,
      message: isFinal ? 'You finished the escape room!' : 'Correct! Loading next clue…',
      finished: isFinal,
      finishedAt: ts,
      currentClue: newState.currentClue,
      nextUnlockTime: newState.nextUnlockTime
    };
  } else {
    // Apply 30-second cooldown
    const until = Date.now() + ESCAPE_ROOM_COOLDOWN_MS;
    PropertiesService.getScriptProperties().setProperty(sharedCooldownKey, String(until));
    return {
      correct: false,
      message: 'Not the right answer. Try again in 30 seconds.',
      cooldownUntil: new Date(until).toISOString()
    };
  }
}

function _formatLocalTime_(iso) {
  try {
    const d = new Date(iso);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a');
  } catch (e) {
    return iso;
  }
}

// ============================================================================
// ADMIN — escape room oversight + auto-placement
// ============================================================================

/**
 * Admin view: per-tribe escape room progress.
 * Returns array of { tribe, current_clue_number, solved_count, finishedAt, attempts_total }
 */
function getEscapeRoomAdminProgress() {
  requireAdmin_();
  const allClues = readClues_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_PROGRESS);
  const data = sheet.getDataRange().getValues();
  const byTribe = {};
  if (data.length >= 2) {
    const h = data[0];
    const tribeCol = h.indexOf('tribe_id');
    const numCol = h.indexOf('clue_number');
    const solvedCol = h.indexOf('solved_at');
    const attemptCol = h.indexOf('attempt_count');
    for (let i = 1; i < data.length; i++) {
      const tid = data[i][tribeCol];
      if (!tid) continue;
      if (!byTribe[tid]) byTribe[tid] = { progress: [], attempts_total: 0 };
      const p = {
        clue_number: parseInt(data[i][numCol], 10),
        solved_at: data[i][solvedCol] || '',
        attempt_count: parseInt(data[i][attemptCol], 10) || 0
      };
      byTribe[tid].progress.push(p);
      byTribe[tid].attempts_total += p.attempt_count;
    }
  }

  return TRIBES.map(function(t) {
    const tp = byTribe[t.id] || { progress: [], attempts_total: 0 };
    const solved = tp.progress.filter(function(p) { return !!p.solved_at; });
    const highest = solved.reduce(function(max, p) {
      return Math.max(max, p.clue_number);
    }, 0);
    const isFinished = highest >= allClues.length && allClues.length > 0;
    let finishedAt = null;
    if (isFinished) {
      const finalP = tp.progress.filter(function(p) { return p.clue_number === allClues.length; })[0];
      finishedAt = finalP ? finalP.solved_at : null;
    }
    return {
      tribe_id: t.id,
      tribe_name: t.name,
      tribe_color_hex: t.color_hex,
      total_clues: allClues.length,
      solved_count: solved.length,
      highest_clue: highest,
      finished: isFinished,
      finished_at: finishedAt,
      attempts_total: tp.attempts_total
    };
  });
}

/**
 * Auto-compute Thursday placements based on finishing order.
 * Tribes that didn't finish are placed by (highest clue solved DESC, then attempts ASC).
 */
function computeEscapeRoomPlacements() {
  const adminEmail = requireAdmin_();
  const allClues = readClues_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.ESCAPE_ROOM_PROGRESS);
  const data = sheet.getDataRange().getValues();

  const byTribe = {};
  if (data.length >= 2) {
    const h = data[0];
    const tribeCol = h.indexOf('tribe_id');
    const numCol = h.indexOf('clue_number');
    const solvedCol = h.indexOf('solved_at');
    const attemptCol = h.indexOf('attempt_count');
    for (let i = 1; i < data.length; i++) {
      const tid = data[i][tribeCol];
      if (!tid) continue;
      if (!byTribe[tid]) byTribe[tid] = { progress: [], attempts_total: 0 };
      const p = {
        clue_number: parseInt(data[i][numCol], 10),
        solved_at: data[i][solvedCol] || '',
        attempt_count: parseInt(data[i][attemptCol], 10) || 0
      };
      byTribe[tid].progress.push(p);
      byTribe[tid].attempts_total += p.attempt_count;
    }
  }

  // Score each tribe
  const scored = TRIBES.map(function(t) {
    const tp = byTribe[t.id] || { progress: [], attempts_total: 0 };
    const solved = tp.progress.filter(function(p) { return !!p.solved_at; });
    const highest = solved.reduce(function(max, p) { return Math.max(max, p.clue_number); }, 0);
    const finalP = tp.progress.filter(function(p) { return p.clue_number === allClues.length; })[0];
    const finishedAt = (highest >= allClues.length && finalP) ? finalP.solved_at : null;
    return {
      tribe_id: t.id,
      finished: !!finishedAt,
      finished_at: finishedAt,
      highest_clue: highest,
      attempts_total: tp.attempts_total
    };
  });

  // Sort: finished first (by finish time ASC), then by highest clue DESC, then attempts ASC
  scored.sort(function(a, b) {
    if (a.finished && !b.finished) return -1;
    if (!a.finished && b.finished) return 1;
    if (a.finished && b.finished) {
      return new Date(a.finished_at) - new Date(b.finished_at);
    }
    if (a.highest_clue !== b.highest_clue) return b.highest_clue - a.highest_clue;
    return a.attempts_total - b.attempts_total;
  });

  // Assign placements 1-6
  const placements = scored.map(function(s, idx) {
    return { tribe_id: s.tribe_id, placement: idx + 1 };
  });

  // Use existing setChallengeResults to write
  setChallengeResults('thu', placements);

  logAdminAction(adminEmail, 'computeEscapeRoomPlacements', JSON.stringify(placements));
  return { ok: true, placements: placements };
}

// ============================================================================
// PHASE 6 — JURY DASHBOARD
// ============================================================================

function requireJuryOrAdmin_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Not signed in.');
  if (ADMIN_EMAILS.indexOf(email) !== -1) return { email: email, isAdmin: true };
  const member = findMemberByEmail(email);
  if (!member) throw new Error('You are not on the roster.');
  if (member.status !== 'jury') throw new Error('This is reserved for the Jury.');
  return { email: email, member: member, isAdmin: false };
}

/**
 * All-tribes view for jury — same as admin but with one minor difference:
 * jury members aren't marked as admin, so the all-tribes endpoint
 * accepts both. Most fields are identical.
 */
function getAllTribesJuryView() {
  // Allow admin or jury to call
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const isAdmin = ADMIN_EMAILS.indexOf(email) !== -1;
  const member = findMemberByEmail(email);
  const isJury = member && member.status === 'jury';
  if (!isAdmin && !isJury) throw new Error('Reserved for jury and admins.');

  // Reuse admin pipeline by temporarily allowing jury
  return _buildAllTribesView_();
}

function _buildAllTribesView_() {
  const allMembers = readMembers_();
  const competedIds = readCompetedMemberIds_();
  const results = readChallengeResults_();

  const tribePoints = {};
  results.forEach(function(r) {
    tribePoints[r.tribe_id] = (tribePoints[r.tribe_id] || 0) + pointsForPlacement_(r.placement, r.day);
  });

  const ranked = TRIBES.map(function(t) {
    return { id: t.id, points: tribePoints[t.id] || 0 };
  }).sort(function(a, b) { return b.points - a.points; });
  const rankById = {};
  ranked.forEach(function(t, idx) { rankById[t.id] = idx + 1; });

  const props = PropertiesService.getScriptProperties();
  const electionDay = props.getProperty('ELECTION_OPEN_FOR_DAY') || '';
  const councilDay = props.getProperty('COUNCIL_OPEN_FOR_DAY') || '';
  const councilTribe = props.getProperty('COUNCIL_TARGET_TRIBE') || '';

  const tribesData = TRIBES.map(function(tribe) {
    const tribeMembers = allMembers.filter(function(m) { return m.tribe_id === tribe.id; });

    const electionTallies = electionDay ? computeElectionTallies_(tribe.id, electionDay, allMembers) : null;
    const councilTallies = (councilDay && councilTribe === tribe.id)
      ? computeCouncilTallies_(tribe.id, councilDay, allMembers)
      : null;
    const chat = readTribeChat_(tribe.id, 25);
    const lockedCompetitors = readLockedCompetitors_(tribe.id);

    return {
      id: tribe.id,
      name: tribe.name,
      color_hex: tribe.color_hex,
      color_label: tribe.color_label,
      points: tribePoints[tribe.id] || 0,
      rank: rankById[tribe.id],
      members: tribeMembers.map(function(m) {
        return {
          id: m.id, name: m.name, email: m.email,
          status: m.status,
          competed: competedIds.indexOf(m.id) !== -1
        };
      }),
      activeCount: tribeMembers.filter(function(m) { return m.status === 'active'; }).length,
      juryCount: tribeMembers.filter(function(m) { return m.status === 'jury'; }).length,
      electionTallies: electionTallies,
      councilTallies: councilTallies,
      lockedCompetitors: lockedCompetitors,
      chat: chat
    };
  });

  return {
    tribes: tribesData,
    electionDay: electionDay,
    councilDay: councilDay,
    councilTribe: councilTribe
  };
}

// Refactor the admin function to use the same internal builder
function getAllTribesAdminView_v2_unused() {
  // (keep original admin function intact for backward compat)
}

// ============================================================================
// JURY STATE — main endpoint for the Jury Dashboard
// ============================================================================
function getJuryState() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Not signed in.');
  const isAdmin = ADMIN_EMAILS.indexOf(email) !== -1;
  const member = findMemberByEmail(email);
  const isJury = member && member.status === 'jury';
  if (!isAdmin && !isJury) throw new Error('This is reserved for the Jury.');

  const allMembers = readMembers_();

  // Identify the jury body
  const juryMembers = allMembers
    .filter(function(m) { return m.status === 'jury'; })
    .map(function(m) {
      const tribe = getTribeById(m.tribe_id);
      return {
        id: m.id,
        name: m.name,
        email: m.email,
        original_tribe_id: m.tribe_id,
        original_tribe_name: tribe ? tribe.name : '',
        original_tribe_color_hex: tribe ? tribe.color_hex : '#888'
      };
    });

  // MVP candidates = active members on any tribe
  const mvpCandidates = allMembers
    .filter(function(m) { return m.status === 'active'; })
    .map(function(m) {
      const tribe = getTribeById(m.tribe_id);
      return {
        id: m.id,
        name: m.name,
        tribe_id: m.tribe_id,
        tribe_name: tribe ? tribe.name : '',
        tribe_color_hex: tribe ? tribe.color_hex : '#888'
      };
    })
    .sort(function(a, b) {
      if (a.tribe_id !== b.tribe_id) return a.tribe_id.localeCompare(b.tribe_id);
      return a.name.localeCompare(b.name);
    });

  // Read MVP voting state
  const mvpState = _readMvpState_(allMembers);
  // Filter "my vote" to current user
  let myVote = null;
  if (member) {
    const myVoteRow = mvpState.votes.filter(function(v) { return v.voter_email === email; })[0];
    if (myVoteRow) myVote = myVoteRow.target_id;
  }

  // Jury chat
  const juryChat = _readJuryChat_(50);

  // My identity
  const me = member ? {
    id: member.id,
    name: member.name,
    email: member.email,
    original_tribe_id: member.tribe_id,
    original_tribe_name: (getTribeById(member.tribe_id) || {}).name || '',
    original_tribe_color_hex: (getTribeById(member.tribe_id) || {}).color_hex || '#888'
  } : null;

  return {
    now: new Date().toISOString(),
    me: me,
    isAdmin: isAdmin,
    juryMembers: juryMembers,
    juryCount: juryMembers.length,
    currentDay: getCurrentDay_(),
    mvpVotingOpen: mvpState.open,
    mvpClosed: mvpState.closed,
    mvpCandidates: mvpCandidates,
    mvpMyVote: myVote,
    mvpVoteCount: mvpState.votes.length,
    mvpResults: mvpState.closed ? mvpState.results : null,
    mvpWinner: mvpState.closed ? mvpState.winner : null,
    juryChat: juryChat
  };
}

function _readJuryChat_(limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.JURY_CHAT);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0];
  const emailCol = h.indexOf('member_email');
  const nameCol = h.indexOf('member_name');
  const msgCol = h.indexOf('message');
  const tsCol = h.indexOf('timestamp');
  const allMembers = readMembers_();

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][msgCol]) continue;
    const memberEmail = (data[i][emailCol] || '').toString().toLowerCase();
    const m = allMembers.filter(function(mm) { return (mm.email || '').toLowerCase() === memberEmail; })[0];
    const tribe = m ? getTribeById(m.tribe_id) : null;
    out.push({
      email: data[i][emailCol],
      name: data[i][nameCol],
      message: data[i][msgCol],
      timestamp: data[i][tsCol],
      original_tribe_color_hex: tribe ? tribe.color_hex : '#888',
      original_tribe_name: tribe ? tribe.name : ''
    });
  }
  return out.slice(-limit);
}

function postJuryChatMessage(text) {
  const ctx = requireJuryOrAdmin_();
  text = (text || '').toString().trim();
  if (!text) throw new Error('Message cannot be empty.');
  if (text.length > 500) throw new Error('Message is too long (max 500 chars).');

  const memberInfo = ctx.member || { email: ctx.email, name: '(Admin)' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.JURY_CHAT);
  sheet.appendRow([memberInfo.email, memberInfo.name, text, new Date().toISOString()]);
  return { ok: true };
}

// ============================================================================
// MVP VOTING
// ============================================================================

/**
 * Read the MVP voting state.
 * Uses ScriptProperties for open/closed flags and the JuryVotes tab for ballots.
 */
function _readMvpState_(allMembers) {
  const props = PropertiesService.getScriptProperties();
  const open = props.getProperty('MVP_VOTING_OPEN') === 'true';
  const closed = props.getProperty('MVP_VOTING_CLOSED') === 'true';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.JURY_VOTES);
  const data = sheet.getDataRange().getValues();
  const votes = [];
  if (data.length >= 2) {
    const h = data[0];
    const voterCol = h.indexOf('voter_email');
    const targetCol = h.indexOf('target_email');
    const tsCol = h.indexOf('timestamp');
    const emailToMember = {};
    allMembers.forEach(function(m) { emailToMember[(m.email || '').toLowerCase()] = m; });
    for (let i = 1; i < data.length; i++) {
      const ve = (data[i][voterCol] || '').toString().toLowerCase();
      const te = (data[i][targetCol] || '').toString().toLowerCase();
      const target = emailToMember[te];
      if (!target) continue;
      votes.push({
        voter_email: ve,
        target_email: te,
        target_id: target.id,
        target_name: target.name,
        target_tribe_id: target.tribe_id,
        timestamp: data[i][tsCol]
      });
    }
  }

  // Compute results (only meaningful if closed, but always available)
  const tally = {};
  votes.forEach(function(v) {
    if (!tally[v.target_email]) {
      tally[v.target_email] = {
        member_id: v.target_id,
        name: v.target_name,
        tribe_id: v.target_tribe_id,
        votes: 0
      };
    }
    tally[v.target_email].votes++;
  });
  const results = Object.keys(tally).map(function(em) {
    const r = tally[em];
    const tribe = getTribeById(r.tribe_id);
    return Object.assign({}, r, {
      tribe_name: tribe ? tribe.name : '',
      tribe_color_hex: tribe ? tribe.color_hex : '#888'
    });
  });
  results.sort(function(a, b) { return b.votes - a.votes; });

  // Winner: top of results, with random tiebreak among ties for first
  let winner = null;
  if (results.length > 0) {
    const topVotes = results[0].votes;
    const tied = results.filter(function(r) { return r.votes === topVotes; });
    winner = tied.length === 1 ? tied[0] : tied[Math.floor(Math.random() * tied.length)];
  }

  return { open: open, closed: closed, votes: votes, results: results, winner: winner };
}

function submitMvpVote(targetMemberId) {
  const ctx = requireJuryOrAdmin_();
  if (!ctx.member) throw new Error('Only jury members can submit MVP votes.');
  if (ctx.member.status !== 'jury') throw new Error('Only jury members can submit MVP votes.');

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MVP_VOTING_OPEN') !== 'true') {
    throw new Error('MVP voting is not open.');
  }
  if (props.getProperty('MVP_VOTING_CLOSED') === 'true') {
    throw new Error('MVP voting has already closed.');
  }

  if (!targetMemberId) throw new Error('Pick a candidate.');
  const allMembers = readMembers_();
  const target = allMembers.filter(function(m) { return m.id === targetMemberId; })[0];
  if (!target) throw new Error('Candidate not found.');
  if (target.status !== 'active') throw new Error('Candidate must be an active player.');

  // Replace any existing vote from this voter
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.JURY_VOTES);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const voterCol = h.indexOf('voter_email');

  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][voterCol] || '').toString().toLowerCase() === ctx.email.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  sheet.appendRow([ctx.email, target.email, new Date().toISOString()]);

  return { ok: true };
}

// ============================================================================
// ADMIN — MVP voting controls
// ============================================================================
function setMvpVotingOpen(open) {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  if (open) {
    props.setProperty('MVP_VOTING_OPEN', 'true');
    props.deleteProperty('MVP_VOTING_CLOSED');
    logAdminAction(adminEmail, 'mvpVotingOpen', 'opened');
  } else {
    props.deleteProperty('MVP_VOTING_OPEN');
    logAdminAction(adminEmail, 'mvpVotingOpen', 'closed (paused, not finalized)');
  }
  return { ok: true };
}

function closeMvpVotingAndReveal() {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('MVP_VOTING_OPEN');
  props.setProperty('MVP_VOTING_CLOSED', 'true');
  logAdminAction(adminEmail, 'closeMvpVotingAndReveal', 'reveal triggered');
  return { ok: true };
}

function resetMvpVoting() {
  const adminEmail = requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('MVP_VOTING_OPEN');
  props.deleteProperty('MVP_VOTING_CLOSED');
  // Clear the JuryVotes tab
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.JURY_VOTES);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  logAdminAction(adminEmail, 'resetMvpVoting', 'cleared');
  return { ok: true };
}

function getMvpAdminState() {
  requireAdmin_();
  const allMembers = readMembers_();
  return _readMvpState_(allMembers);
}

// ============================================================================
// PHASE 7 — CEREMONY EVENT SYSTEM
// ============================================================================
//
// Ceremonies are short-lived "event" markers stored in ScriptProperties.
// The Contest dashboard polls for active events and plays them once.
// Each ceremony has a unique key; once dismissed (by viewer or by timeout),
// the marker stays but with a "consumed" flag so it doesn't replay.
//
// Event types: 'idol_claimed', 'council_reveal', 'tribe_merge'
//
// Storage: ScriptProperties['CEREMONY_QUEUE'] = JSON array of events
//   [{ id: 'uuid', type: 'idol_claimed', payload: {...}, ts: '...' }, ...]
//
// Each ceremony lives for ~30 seconds in the queue, then gets pruned.

const CEREMONY_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours — covers same-day catch-up

function _readCeremonyQueue_() {
  const raw = PropertiesService.getScriptProperties().getProperty('CEREMONY_QUEUE');
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}

function _writeCeremonyQueue_(arr) {
  // Prune events older than TTL
  const now = Date.now();
  const filtered = arr.filter(function(e) {
    if (!e.ts) return false;
    return (now - new Date(e.ts).getTime()) < CEREMONY_TTL_MS;
  });
  PropertiesService.getScriptProperties().setProperty('CEREMONY_QUEUE', JSON.stringify(filtered));
}

function _enqueueCeremony_(type, payload) {
  const queue = _readCeremonyQueue_();
  const id = Utilities.getUuid();
  queue.push({
    id: id,
    type: type,
    payload: payload || {},
    ts: new Date().toISOString()
  });
  _writeCeremonyQueue_(queue);
  return id;
}

/**
 * Public read for the Contest dashboard — returns active (unexpired) ceremonies.
 */
function getActiveCeremonies() {
  const queue = _readCeremonyQueue_();
  // Re-prune in case nothing has triggered a write recently
  _writeCeremonyQueue_(queue);
  return _readCeremonyQueue_();
}

/**
 * Admin: re-trigger a ceremony manually for testing.
 */
function triggerCeremonyForTesting(type, customPayload) {
  requireAdmin_();
  let payload = customPayload || {};
  // If admin doesn't supply payload, build a sensible default
  if (type === 'idol_claimed' && !payload.member_name) {
    payload = {
      idol_name: 'Rammy',
      member_name: 'Test Member',
      tribe_name: 'Tribe Test',
      tribe_color_hex: '#E2723F'
    };
  } else if (type === 'council_reveal' && !payload.eliminated) {
    payload = {
      tribe_id: 'tribe_ember',
      tribe_name: 'Tribe Ember',
      tribe_color_hex: '#C8332B',
      eliminated: [
        { name: 'Test One', tribe_color_hex: '#C8332B' },
        { name: 'Test Two', tribe_color_hex: '#C8332B' }
      ],
      day: 'mon'
    };
  } else if (type === 'tribe_merge' && !payload.dissolving) {
    payload = {
      dissolving: { name: 'Tribe Bone', color_hex: '#F3F3EE' },
      receiving: [
        { name: 'Tribe Ember', color_hex: '#C8332B' },
        { name: 'Tribe Marlin', color_hex: '#1F6FB8' }
      ]
    };
  }
  return { ok: true, id: _enqueueCeremony_(type, payload) };
}

function clearAllCeremonies() {
  requireAdmin_();
  PropertiesService.getScriptProperties().deleteProperty('CEREMONY_QUEUE');
  return { ok: true };
}

// ============================================================================
// HOOK INTO EXISTING WRITES — auto-enqueue ceremonies
// ============================================================================
// These are NOT used directly. Instead, the original claimIdol() and
// closeCouncilAndEliminate() functions in earlier sections are modified
// in-place to call _enqueueCeremony_() on success. See those functions.

// ============================================================================
// TRIBE MERGE / ABSORPTION — explicit admin action
// ============================================================================
//
// Triggered when admin recognizes a tribe has dropped to ≤2 active members
// and decides to redistribute them. Spec says "smallest-first" redistribution.

function dissolveTribeAndRedistribute(tribeIdToDissolve) {
  const adminEmail = requireAdmin_();
  const dissolving = getTribeById(tribeIdToDissolve);
  if (!dissolving) throw new Error('Tribe not found.');

  const allMembers = readMembers_();
  const activeFromTribe = allMembers.filter(function(m) {
    return m.tribe_id === tribeIdToDissolve && m.status === 'active';
  });
  if (activeFromTribe.length === 0) throw new Error('No active members in that tribe.');

  // Compute size of each OTHER tribe (active members), pick smallest first.
  // Exclude any tribe that has no active members — those are effectively already
  // dissolved and shouldn't receive new members. Also exclude the tribe currently
  // being dissolved.
  const otherTribeSizes = TRIBES
    .filter(function(t) { return t.id !== tribeIdToDissolve; })
    .map(function(t) {
      const count = allMembers.filter(function(m) {
        return m.tribe_id === t.id && m.status === 'active';
      }).length;
      return { tribe: t, count: count };
    })
    .filter(function(x) { return x.count > 0; });

  if (otherTribeSizes.length === 0) {
    throw new Error('No remaining tribes with active members to receive redistributed castaways.');
  }

  // Round-robin: assign smallest-first, recompute after each assignment
  const reassignments = [];
  activeFromTribe.forEach(function(member) {
    otherTribeSizes.sort(function(a, b) { return a.count - b.count; });
    const target = otherTribeSizes[0];
    reassignments.push({ member: member, newTribeId: target.tribe.id });
    target.count++;
  });

  // Write to Members sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const idCol = h.indexOf('id');
  const tribeCol = h.indexOf('tribe_id');

  reassignments.forEach(function(r) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === r.member.id) {
        sheet.getRange(i + 1, tribeCol + 1).setValue(r.newTribeId);
        break;
      }
    }
  });

  // Build merge ceremony payload
  const receivingTribes = {};
  reassignments.forEach(function(r) {
    if (!receivingTribes[r.newTribeId]) {
      const t = getTribeById(r.newTribeId);
      receivingTribes[r.newTribeId] = { name: t.name, color_hex: t.color_hex };
    }
  });
  _enqueueCeremony_('tribe_merge', {
    dissolving: { name: dissolving.name, color_hex: dissolving.color_hex },
    receiving: Object.keys(receivingTribes).map(function(k) { return receivingTribes[k]; })
  });

  logAdminAction(adminEmail, 'dissolveTribeAndRedistribute',
    dissolving.name + ' → ' + reassignments.map(function(r) {
      const t = getTribeById(r.newTribeId);
      return r.member.name + ' to ' + t.name;
    }).join('; '));

  return {
    ok: true,
    reassignments: reassignments.map(function(r) {
      const t = getTribeById(r.newTribeId);
      return { name: r.member.name, newTribe: t.name };
    })
  };
}

// ============================================================================
// CONFESSIONAL TICKER — already collected, just exposed for ticker reads
// ============================================================================
// getRecentConfessionals(limit) already exists; add a shuffle for variety.

function getConfessionalsForTicker(limit) {
  const list = getRecentConfessionals(limit || 50);
  // Shuffle for varied display
  for (let i = list.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = list[i]; list[i] = list[j]; list[j] = tmp;
  }
  return list;
}

// ============================================================================
// DAY CONTROLS — server endpoints for the Admin "Day Controls" panel
// ============================================================================
// These mirror the Sim Panel actions but are meant to be admin's daily-rhythm
// buttons: 8am close election + lock; 2pm open council; 6pm close council +
// open next election. They wrap existing functions and surface a clean state
// summary that the Admin panel can use to highlight what to do next.

function getDayControlsState() {
  requireAdmin_();
  const props = PropertiesService.getScriptProperties();
  const electionDay = props.getProperty('ELECTION_OPEN_FOR_DAY') || '';
  const councilDay = props.getProperty('COUNCIL_OPEN_FOR_DAY') || '';
  const councilTribe = props.getProperty('COUNCIL_TARGET_TRIBE') || '';
  const simDay = props.getProperty('SIM_DAY') || '';

  // Compute "what's the live current day" — sim override if set, else infer
  // from real date. Roy-Hart Survivor runs May 4-8, 2026.
  let currentDay = '';
  if (simDay && simDay !== 'auto') {
    currentDay = simDay;
  } else {
    currentDay = getCurrentDay_() || '';
  }

  // Find next-day-after-currentDay for "open next election" suggestions
  const nextDayMap = { mon: 'tue', tue: 'wed', wed: 'thu', thu: 'fri', fri: '' };
  const nextDay = nextDayMap[currentDay] || '';

  // Today's placements (if any) per tribe
  const todaysPlacements = {}; // tribeId -> placement number
  if (currentDay) {
    const results = readChallengeResults_();
    results.filter(function(r) { return r.day === currentDay; }).forEach(function(r) {
      todaysPlacements[r.tribe_id] = parseInt(r.placement, 10) || null;
    });
  }
  const placementCount = Object.keys(todaysPlacements).length;

  // Compute losing tribe of current day (the tribe that placed 6th)
  let suggestedLoserTribeId = '';
  if (currentDay) {
    Object.keys(todaysPlacements).forEach(function(tid) {
      if (todaysPlacements[tid] === 6) suggestedLoserTribeId = tid;
    });
  }

  // Locked competitors per tribe (so we can render "competitors locked" status)
  const lockedByTribe = {};
  if (currentDay) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compSheet = ss.getSheetByName(TAB.COMPETITORS);
    const compData = compSheet.getDataRange().getValues();
    if (compData.length > 1) {
      const h = compData[0];
      const tCol = h.indexOf('tribe_id');
      const dCol = h.indexOf('day');
      const nCol = h.indexOf('member_name');
      for (let i = 1; i < compData.length; i++) {
        if (compData[i][dCol] !== currentDay) continue;
        const tid = compData[i][tCol];
        if (!lockedByTribe[tid]) lockedByTribe[tid] = [];
        lockedByTribe[tid].push(compData[i][nCol]);
      }
    }
  }

  // MVP voting state (used Friday for the special MVP steps in Day Controls)
  const mvpOpen = props.getProperty('MVP_VOTING_OPEN') === 'true';
  const mvpClosed = props.getProperty('MVP_VOTING_CLOSED') === 'true';
  let mvpState = 'pending'; // not yet opened
  if (mvpClosed) mvpState = 'revealed';
  else if (mvpOpen) mvpState = 'open';

  return {
    currentDay: currentDay,
    nextDay: nextDay,
    electionOpenForDay: electionDay,
    councilOpenForDay: councilDay,
    councilOpenForTribe: councilTribe,
    suggestedLoserTribeId: suggestedLoserTribeId,
    todaysPlacements: todaysPlacements,
    placementCount: placementCount,
    lockedCompetitorsByTribe: lockedByTribe,
    mvpState: mvpState,
    days: ['mon', 'tue', 'wed', 'thu', 'fri'],
    dayLabels: { mon: 'Monday', tue: 'Tuesday', wed: 'Wednesday', thu: 'Thursday', fri: 'Friday' },
    challengeLabels: {
      mon: 'Cornhole', tue: 'Foosball', wed: 'Puzzle', thu: 'Escape Room', fri: 'Obstacle Course'
    }
  };
}

// ============================================================================
// EMERGENCY ENDPOINTS — used only by the Admin Emergency tab
// ============================================================================

/**
 * Returns active+jury members of a tribe for the emergency elim picker.
 * Admin-only.
 */
function getMembersForTribeAdmin(tribeId) {
  requireAdmin_();
  if (!tribeId) throw new Error('Tribe id required.');
  const tribe = getTribeById(tribeId);
  if (!tribe) throw new Error('Tribe not found.');
  const all = readMembers_();
  return all
    .filter(function(m) { return m.tribe_id === tribeId; })
    .map(function(m) {
      return {
        id: m.id,
        name: m.name,
        email: m.email,
        status: m.status
      };
    });
}

/**
 * Manually eliminate one or more members from a tribe (bypassing council).
 * Use only when someone quits or needs to be removed outside normal flow.
 * Admin-only.
 */
function manuallyEliminateMembers(tribeId, memberIds) {
  const adminEmail = requireAdmin_();
  if (!tribeId) throw new Error('Tribe id required.');
  if (!Array.isArray(memberIds) || memberIds.length === 0) {
    throw new Error('At least one member must be selected.');
  }

  const tribe = getTribeById(tribeId);
  if (!tribe) throw new Error('Tribe not found.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const idCol = h.indexOf('id');
  const tribeCol = h.indexOf('tribe_id');
  const statusCol = h.indexOf('status');
  const nameCol = h.indexOf('name');

  const eliminated = [];

  memberIds.forEach(function(mid) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] !== mid) continue;
      if (data[i][tribeCol] !== tribeId) {
        throw new Error('Member ' + data[i][nameCol] + ' is not in the selected tribe.');
      }
      if (data[i][statusCol] !== 'active') {
        // skip already-jury members silently
        continue;
      }
      sheet.getRange(i + 1, statusCol + 1).setValue('jury');
      eliminated.push({ id: mid, name: data[i][nameCol] });
      break;
    }
  });

  if (eliminated.length === 0) {
    throw new Error('No active members were eliminated (all selected may already be on jury).');
  }

  logAdminAction(adminEmail, 'manuallyEliminateMembers',
    tribe.name + ': ' + eliminated.map(function(e) { return e.name; }).join(', '));

  return { ok: true, eliminated: eliminated };
}
