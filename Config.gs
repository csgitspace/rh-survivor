/**
 * Roy-Hart Survivor Week — Configuration
 * 
 * All app-wide constants live here. Edit this file to update admin emails,
 * tribe definitions, or scoring rules without touching the rest of the code.
 */

// ============================================================================
// ADMIN EMAILS — these users get the Admin Dashboard on login
// ============================================================================
const ADMIN_EMAILS = [
  'cschaus@royhart.org',
  'rdodge@royhart.org',
  'cdimayo@royhart.org',
  'msweeney@royhart.org',
  'jheck@royhart.org'
];

// ============================================================================
// TRIBES — the six teams. Edit names freely; ids and hex colors are used by code.
// ============================================================================
const TRIBES = [
  { id: 'marlin',   name: 'Tribe Marlin',   color_label: 'Blue',   color_hex: '#1F6FB8' },
  { id: 'bone',     name: 'Tribe Bone',     color_label: 'White',  color_hex: '#F3F3EE' },
  { id: 'orchid',   name: 'Tribe Orchid',   color_label: 'Purple', color_hex: '#6B3FA0' },
  { id: 'ember',    name: 'Tribe Ember',    color_label: 'Red',    color_hex: '#C8332B' },
  { id: 'hibiscus', name: 'Tribe Hibiscus', color_label: 'Pink',   color_hex: '#E5559C' },
  { id: 'fern',     name: 'Tribe Fern',     color_label: 'Green',  color_hex: '#3B8C4F' }
];

// ============================================================================
// CHALLENGE SCHEDULE — Mon-Fri of contest week
// ============================================================================
const CHALLENGES = [
  { day: 'mon', date: '2026-05-04', type: 'cornhole',    name: 'Cornhole',          format: 'bracket'    },
  { day: 'tue', date: '2026-05-05', type: 'foosball',    name: 'Foosball',          format: 'bracket'    },
  { day: 'wed', date: '2026-05-06', type: 'puzzle',      name: 'Puzzle Race',       format: 'auto_timed' },
  { day: 'thu', date: '2026-05-07', type: 'escape_room', name: 'Survivor Escape',   format: 'auto_timed' },
  { day: 'fri', date: '2026-05-08', type: 'obstacle',    name: 'Obstacle Course',   format: 'manual'     }
];

// ============================================================================
// SCORING — points awarded per placement
// ============================================================================
const SCORING_DAILY  = [10, 7, 5, 3, 1, 0]; // 1st through 6th place, Mon-Thu
const SCORING_FRIDAY = [20, 14, 10, 6, 2, 0]; // double-weighted finale

// ============================================================================
// SHEET TAB NAMES — change these only if you also change setupSheets()
// ============================================================================
const TAB = {
  MEMBERS:              'Members',
  TRIBES:               'Tribes',
  CHALLENGES:           'Challenges',
  SELECTIONS:           'Selections',
  COMPETITORS:          'Competitors',
  CHALLENGE_RESULTS:    'ChallengeResults',
  ELIMINATIONS:         'Eliminations',
  ELIMINATION_RESULTS:  'EliminationResults',
  ESCAPE_ROOM_CLUES:    'EscapeRoomClues',
  ESCAPE_ROOM_PROGRESS: 'EscapeRoomProgress',
  PUZZLE_PHOTOS:        'PuzzlePhotos',
  TRIBE_CHAT:           'TribeChat',
  JURY_CHAT:            'JuryChat',
  CONFESSIONALS:        'Confessionals',
  IDOL_STATE:           'IdolState',
  JURY_VOTES:           'JuryVotes',
  ADMIN_LOG:            'AdminLog'
};
