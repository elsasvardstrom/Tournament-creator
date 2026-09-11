/**
 * Americano Tournament Generator for Google Sheets
 *
 * Features:
 * - User-defined number of players, courts, match length and tournament length
 * - Automatic 5-minute breaks between rounds
 * - Player names + unique initials
 * - Fair rotation when not everyone can play at the same time
 * - Tries to avoid repeated partners and opponents
 * - Tournament sheet using initials
 * - Manual score entry
 * - Automatic individual points table
 *
 * No player names or spreadsheet IDs are hard-coded.
 */

const CONFIG = {
  BREAK_MINUTES: 5,
  SHEETS: {
    SETTINGS: "Settings",
    PLAYERS: "Players",
    TOURNAMENT: "Tournament"
  },
  COLORS: {
    HEADER: "#d9ead3",
    SCORE_INPUT: "#e2f0d9",
    TOTAL_SCORE: "#fff2cc",
    INFO: "#fff2cc"
  }
};


// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎾 Americano")
    .addItem("1. Set up tournament", "setupTournament")
    .addItem("2. Generate schedule", "generateTournament")
    .addSeparator()
    .addItem("Generate new schedule", "generateTournament")
    .addToUi();
}


// ============================================================
// STEP 1 — TOURNAMENT SETUP
// ============================================================

function setupTournament() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const numberOfPlayers = promptPositiveInteger_(
    ui,
    "Number of players",
    "How many players are participating?",
    4
  );
  if (numberOfPlayers === null) return;

  const requestedCourts = promptPositiveInteger_(
    ui,
    "Number of courts",
    "How many courts are available?",
    1
  );
  if (requestedCourts === null) return;

  const maxPossibleCourts = Math.floor(numberOfPlayers / 4);
  const numberOfCourts = Math.min(requestedCourts, maxPossibleCourts);

  const matchMinutes = promptPositiveInteger_(
    ui,
    "Match length",
    "How many minutes should each match last?\n\nA 5-minute break is automatically added between rounds.",
    1
  );
  if (matchMinutes === null) return;

  const totalMinutes = promptPositiveInteger_(
    ui,
    "Tournament length",
    "How many minutes should the entire tournament last?",
    matchMinutes
  );
  if (totalMinutes === null) return;

  const numberOfRounds = Math.floor(
    (totalMinutes + CONFIG.BREAK_MINUTES) /
    (matchMinutes + CONFIG.BREAK_MINUTES)
  );

  if (numberOfRounds < 1) {
    ui.alert("No round fits within the selected tournament length.");
    return;
  }

  createSettingsSheet_(
    ss,
    numberOfPlayers,
    numberOfCourts,
    matchMinutes,
    totalMinutes,
    numberOfRounds
  );

  const playerSheet = createPlayersSheet_(ss, numberOfPlayers);

  let message =
    "Tournament setup created!\n\n" +
    `${numberOfPlayers} players\n` +
    `${numberOfCourts} courts\n` +
    `${matchMinutes} minutes per match\n` +
    `${CONFIG.BREAK_MINUTES} minutes between rounds\n` +
    `${numberOfRounds} rounds\n\n` +
    "Next:\n" +
    "1. Enter initials and names in the Players sheet.\n" +
    "2. Choose 🎾 Americano → 2. Generate schedule.";

  if (requestedCourts > numberOfCourts) {
    message +=
      `\n\nNote: ${requestedCourts} courts were requested, but only ` +
      `${numberOfCourts} can be used simultaneously with ${numberOfPlayers} players.`;
  }

  ui.alert(message);
  playerSheet.activate();
}


// ============================================================
// STEP 2 — GENERATE SCHEDULE
// ============================================================

function generateTournament() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  const playerSheet = ss.getSheetByName(CONFIG.SHEETS.PLAYERS);

  if (!settingsSheet || !playerSheet) {
    ui.alert(
      "Set up the tournament first:\n\n" +
      "🎾 Americano → 1. Set up tournament"
    );
    return;
  }

  const settings = readSettings_(settingsSheet);
  const players = readPlayers_(playerSheet, settings.numberOfPlayers);

  const validationError = validatePlayers_(players);
  if (validationError) {
    ui.alert(validationError);
    return;
  }

  const schedule = createAmericanoSchedule_(
    players.length,
    settings.numberOfRounds,
    settings.numberOfCourts
  );

  if (!schedule) {
    ui.alert(
      "A valid schedule could not be created.\n\n" +
      "Try generating the schedule again."
    );
    return;
  }

  const tournamentSheet = createTournamentSheet_(
    ss,
    players,
    schedule,
    settings.matchMinutes,
    settings.numberOfCourts
  );

  updateStandings_();

  ui.alert(
    "Schedule created!\n\n" +
    "Enter each team's score manually in the green score columns.\n" +
    "The individual points table updates automatically."
  );

  tournamentSheet.activate();
}


// ============================================================
// AUTOMATIC STANDINGS UPDATE
// ============================================================

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.SHEETS.TOURNAMENT) return;

  const firstColumn = e.range.getColumn();
  const lastColumn = e.range.getLastColumn();

  // Score columns G:H only
  if (lastColumn < 7 || firstColumn > 8) return;

  updateStandings_();
}


function updateStandings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  const playerSheet = ss.getSheetByName(CONFIG.SHEETS.PLAYERS);
  const tournamentSheet = ss.getSheetByName(CONFIG.SHEETS.TOURNAMENT);

  if (!settingsSheet || !playerSheet || !tournamentSheet) return;

  const settings = readSettings_(settingsSheet);
  const players = readPlayers_(playerSheet, settings.numberOfPlayers);

  const scoreStartCol = 10; // J
  const lastMatchRow = 1 + settings.numberOfRounds * settings.numberOfCourts;

  if (lastMatchRow < 2) return;

  const matchData = tournamentSheet
    .getRange(2, 3, lastMatchRow - 1, 6) // C:H
    .getValues();

  const stats = {};

  players.forEach(player => {
    stats[player.initials] = {
      matches: 0,
      points: 0
    };
  });

  matchData.forEach(row => {
    const team1 = [String(row[0]).trim(), String(row[1]).trim()];
    const team2 = [String(row[2]).trim(), String(row[3]).trim()];
    const score1 = row[4];
    const score2 = row[5];

    team1.forEach(initials => {
      if (stats[initials]) {
        stats[initials].matches++;

        if (typeof score1 === "number" && !isNaN(score1)) {
          stats[initials].points += score1;
        }
      }
    });

    team2.forEach(initials => {
      if (stats[initials]) {
        stats[initials].matches++;

        if (typeof score2 === "number" && !isNaN(score2)) {
          stats[initials].points += score2;
        }
      }
    });
  });

  const standings = players.map(player => [
    player.initials,
    player.name,
    stats[player.initials].matches,
    stats[player.initials].points
  ]);

  tournamentSheet
    .getRange(2, scoreStartCol, standings.length, 4)
    .setValues(standings);

  tournamentSheet
    .getRange(2, scoreStartCol + 3, standings.length, 1)
    .setBackground(CONFIG.COLORS.TOTAL_SCORE)
    .setFontWeight("bold");
}


// ============================================================
// SHEET CREATION
// ============================================================

function createSettingsSheet_(
  ss,
  numberOfPlayers,
  numberOfCourts,
  matchMinutes,
  totalMinutes,
  numberOfRounds
) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
  }

  sheet.clear();

  const values = [
    ["AMERICANO SETTINGS", ""],
    ["Number of players", numberOfPlayers],
    ["Number of courts", numberOfCourts],
    ["Match length (minutes)", matchMinutes],
    ["Break between rounds (minutes)", CONFIG.BREAK_MINUTES],
    ["Tournament length (minutes)", totalMinutes],
    ["Number of rounds", numberOfRounds]
  ];

  sheet.getRange(1, 1, values.length, 2).setValues(values);

  sheet
    .getRange("A1:B1")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);

  sheet.getRange("A2:A7").setFontWeight("bold");
  sheet.autoResizeColumns(1, 2);
}


function createPlayersSheet_(ss, numberOfPlayers) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PLAYERS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PLAYERS);
  }

  sheet.clear();

  sheet
    .getRange("A1:C1")
    .setValues([["No.", "Initials", "Player"]])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);

  const rows = [];

  for (let i = 1; i <= numberOfPlayers; i++) {
    rows.push([i, "", ""]);
  }

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  const infoRow = numberOfPlayers + 3;

  sheet
    .getRange(infoRow, 1, 3, 3)
    .setValues([
      ["INFO", "", ""],
      ["", "Enter unique initials, e.g. JS", "Enter the player's full name eg Jannik Sinner"],
      ["", "If two players have the same initials, use e.g. JS1 and JS2", ""]
    ]);

  sheet
    .getRange(infoRow, 1, 1, 3)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.INFO);

  sheet
    .getRange(infoRow + 1, 1, 2, 3)
    .setFontStyle("italic");

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 240);
  sheet.setFrozenRows(1);

  return sheet;
}


function createTournamentSheet_(
  ss,
  players,
  schedule,
  matchMinutes,
  numberOfCourts
) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.TOURNAMENT);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.TOURNAMENT);
  }

  // Remove old merges before rebuilding the sheet.
  sheet.getDataRange().breakApart();
  sheet.clear();

  const headers = [
    "Round",
    "Court",
    "Team 1",
    "",
    "Team 2",
    "",
    "Score T1",
    "Score T2"
  ];

  sheet.getRange(1, 1, 1, 8).setValues([headers]);

  sheet
    .getRange("A1:H1")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(CONFIG.COLORS.HEADER);

  sheet.getRange("C1:D1").merge();
  sheet.getRange("E1:F1").merge();

  const scheduleRows = [];

  for (let round = 0; round < schedule.length; round++) {
    for (let court = 0; court < numberOfCourts; court++) {
      const match = schedule[round][court];

      scheduleRows.push([
        round + 1,
        court + 1,
        players[match[0][0]].initials,
        players[match[0][1]].initials,
        players[match[1][0]].initials,
        players[match[1][1]].initials,
        "",
        ""
      ]);
    }
  }

  sheet
    .getRange(2, 1, scheduleRows.length, 8)
    .setValues(scheduleRows);

  let firstRoundRow = 2;

  for (let round = 0; round < schedule.length; round++) {
    const startMinute =
      round * (matchMinutes + CONFIG.BREAK_MINUTES);

    const endMinute =
      startMinute + matchMinutes;

    const timeText =
      `${formatMinutes_(startMinute)}-${formatMinutes_(endMinute)}`;

    const roundRange = sheet.getRange(
      firstRoundRow,
      1,
      numberOfCourts,
      1
    );

    roundRange.merge();

    roundRange
      .setValue(`${round + 1}\n${timeText}`)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet
      .getRange(
        firstRoundRow + numberOfCourts - 1,
        1,
        1,
        8
      )
      .setBorder(
        null,
        null,
        true,
        null,
        null,
        null,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

    firstRoundRow += numberOfCourts;
  }

  sheet
    .getRange(2, 7, scheduleRows.length, 2)
    .setBackground(CONFIG.COLORS.SCORE_INPUT)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  const scoreStartCol = 10; // J

  sheet
    .getRange(1, scoreStartCol, 1, 4)
    .setValues([[
      "Initials",
      "Player",
      "Matches",
      "POINTS"
    ]])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setHorizontalAlignment("center");

  const standings = players.map(player => [
    player.initials,
    player.name,
    0,
    0
  ]);

  sheet
    .getRange(2, scoreStartCol, standings.length, 4)
    .setValues(standings);

  sheet
    .getRange(2, scoreStartCol + 3, standings.length, 1)
    .setBackground(CONFIG.COLORS.TOTAL_SCORE)
    .setFontWeight("bold");

  sheet
    .getRange(1, 1, scheduleRows.length + 1, 8)
    .setBorder(true, true, true, true, true, true);

  sheet
    .getRange(1, scoreStartCol, players.length + 1, 4)
    .setBorder(true, true, true, true, true, true);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 65);
  sheet.setColumnWidth(3, 75);
  sheet.setColumnWidth(4, 75);
  sheet.setColumnWidth(5, 75);
  sheet.setColumnWidth(6, 75);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(10, 85);
  sheet.setColumnWidth(11, 200);
  sheet.setColumnWidth(12, 90);
  sheet.setColumnWidth(13, 90);

  sheet
    .getRange(2, 2, scheduleRows.length, 7)
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);

  return sheet;
}


// ============================================================
// SCHEDULING ALGORITHM
// ============================================================

function createAmericanoSchedule_(
  numberOfPlayers,
  numberOfRounds,
  numberOfCourts
) {
  const playersPerRound = numberOfCourts * 4;

  const matchesPlayed = Array(numberOfPlayers).fill(0);

  const partnerCount = Array.from(
    { length: numberOfPlayers },
    () => Array(numberOfPlayers).fill(0)
  );

  const opponentCount = Array.from(
    { length: numberOfPlayers },
    () => Array(numberOfPlayers).fill(0)
  );

  const schedule = [];

  for (let round = 0; round < numberOfRounds; round++) {
    let bestRound = null;
    let bestScore = Infinity;

    for (let attempt = 0; attempt < 2500; attempt++) {
      const candidates =
        [...Array(numberOfPlayers).keys()];

      // Players with fewer matches get priority.
      candidates.sort((a, b) => {
        const difference =
          matchesPlayed[a] - matchesPlayed[b];

        if (difference !== 0) {
          return difference;
        }

        return Math.random() - 0.5;
      });

      const selectedPlayers =
        candidates.slice(0, playersPerRound);

      shuffleInPlace_(selectedPlayers);

      const matches = [];
      let roundScore = 0;

      for (
        let i = 0;
        i < selectedPlayers.length;
        i += 4
      ) {
        const four = [
          selectedPlayers[i],
          selectedPlayers[i + 1],
          selectedPlayers[i + 2],
          selectedPlayers[i + 3]
        ];

        const combinations = [
          [[four[0], four[1]], [four[2], four[3]]],
          [[four[0], four[2]], [four[1], four[3]]],
          [[four[0], four[3]], [four[1], four[2]]]
        ];

        let bestCombination = null;
        let bestCombinationScore = Infinity;

        combinations.forEach(combination => {
          const a = combination[0][0];
          const b = combination[0][1];
          const c = combination[1][0];
          const d = combination[1][1];

          let score = 0;

          // Repeated partners get a large penalty.
          score += partnerCount[a][b] * 100;
          score += partnerCount[c][d] * 100;

          // Repeated opponents get a smaller penalty.
          score += opponentCount[a][c] * 10;
          score += opponentCount[a][d] * 10;
          score += opponentCount[b][c] * 10;
          score += opponentCount[b][d] * 10;

          if (score < bestCombinationScore) {
            bestCombinationScore = score;
            bestCombination = combination;
          }
        });

        matches.push(bestCombination);
        roundScore += bestCombinationScore;
      }

      // Slight penalty for selecting players who have already
      // played more often than others.
      selectedPlayers.forEach(player => {
        roundScore += matchesPlayed[player] * 3;
      });

      // Random tie-breaker.
      roundScore += Math.random();

      if (roundScore < bestScore) {
        bestScore = roundScore;
        bestRound = matches;
      }
    }

    if (!bestRound) {
      return null;
    }

    schedule.push(bestRound);

    bestRound.forEach(match => {
      const a = match[0][0];
      const b = match[0][1];
      const c = match[1][0];
      const d = match[1][1];

      matchesPlayed[a]++;
      matchesPlayed[b]++;
      matchesPlayed[c]++;
      matchesPlayed[d]++;

      partnerCount[a][b]++;
      partnerCount[b][a]++;
      partnerCount[c][d]++;
      partnerCount[d][c]++;

      addOpponent_(opponentCount, a, c);
      addOpponent_(opponentCount, a, d);
      addOpponent_(opponentCount, b, c);
      addOpponent_(opponentCount, b, d);
    });
  }

  return schedule;
}


// ============================================================
// READ + VALIDATE DATA
// ============================================================

function readSettings_(sheet) {
  return {
    numberOfPlayers: Number(sheet.getRange("B2").getValue()),
    numberOfCourts: Number(sheet.getRange("B3").getValue()),
    matchMinutes: Number(sheet.getRange("B4").getValue()),
    numberOfRounds: Number(sheet.getRange("B7").getValue())
  };
}


function readPlayers_(sheet, numberOfPlayers) {
  return sheet
    .getRange(2, 2, numberOfPlayers, 2)
    .getValues()
    .map(row => ({
      initials: String(row[0]).trim().toUpperCase(),
      name: String(row[1]).trim()
    }));
}


function validatePlayers_(players) {
  for (let i = 0; i < players.length; i++) {
    if (!players[i].initials) {
      return `Player ${i + 1} is missing initials.`;
    }

    if (!players[i].name) {
      return `Player ${i + 1} is missing a name.`;
    }
  }

  const initials = players.map(player => player.initials);
  const uniqueInitials = new Set(initials);

  if (uniqueInitials.size !== initials.length) {
    return (
      "Two or more players have the same initials.\n\n" +
      "Use unique initials such as JS1 and JS2."
    );
  }

  return null;
}


// ============================================================
// HELPERS
// ============================================================

function promptPositiveInteger_(
  ui,
  title,
  message,
  minimum
) {
  const response = ui.prompt(
    title,
    message,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const value = Number(response.getResponseText().trim());

  if (
    !Number.isInteger(value) ||
    value < minimum
  ) {
    ui.alert(
      `Please enter a whole number of at least ${minimum}.`
    );
    return null;
  }

  return value;
}


function addOpponent_(opponentCount, a, b) {
  opponentCount[a][b]++;
  opponentCount[b][a]++;
}


function shuffleInPlace_(array) {
  for (
    let i = array.length - 1;
    i > 0;
    i--
  ) {
    const j =
      Math.floor(Math.random() * (i + 1));

    [array[i], array[j]] =
      [array[j], array[i]];
  }

  return array;
}


function formatMinutes_(totalMinutes) {
  const hours =
    Math.floor(totalMinutes / 60);

  const minutes =
    totalMinutes % 60;

  return (
    String(hours).padStart(2, "0") +
    ":" +
    String(minutes).padStart(2, "0")
  );
}
