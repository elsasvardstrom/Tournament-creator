// ============================================================
// 🎾 AMERICANO TOURNAMENT GENERATOR
// ============================================================

const BREAK_MINUTES = 5;


// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎾 Americano")
    .addItem("1. Skapa turnering", "setupTournament")
    .addItem("2. Generera spelschema", "generateTournament")
    .addSeparator()
    .addItem("Generera på nytt", "generateTournament")
    .addToUi();
}


// ============================================================
// STEP 1 — CREATE TOURNAMENT
// ============================================================

function setupTournament() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // -------------------------------
  // NUMBER OF PLAYERS
  // -------------------------------

  const playerResponse = ui.prompt(
    "Antal deltagare",
    "Hur många spelare deltar i turneringen?",
    ui.ButtonSet.OK_CANCEL
  );

  if (playerResponse.getSelectedButton() !== ui.Button.OK) return;

  const numberOfPlayers = parseInt(playerResponse.getResponseText());

  if (isNaN(numberOfPlayers) || numberOfPlayers < 4) {
    ui.alert("❌ Du måste ha minst 4 spelare.");
    return;
  }


  // -------------------------------
  // NUMBER OF COURTS
  // -------------------------------

  const courtResponse = ui.prompt(
    "Antal banor",
    "Hur många tennisbanor finns tillgängliga?",
    ui.ButtonSet.OK_CANCEL
  );

  if (courtResponse.getSelectedButton() !== ui.Button.OK) return;

  const requestedCourts = parseInt(courtResponse.getResponseText());

  if (isNaN(requestedCourts) || requestedCourts < 1) {
    ui.alert("❌ Ange minst 1 bana.");
    return;
  }

  const maxPossibleCourts = Math.floor(numberOfPlayers / 4);

  const numberOfCourts = Math.min(
    requestedCourts,
    maxPossibleCourts
  );


  // -------------------------------
  // MATCH LENGTH
  // -------------------------------

  const matchResponse = ui.prompt(
    "Matchlängd",
    "Hur många minuter ska varje match vara?\n\n5 minuters paus läggs automatiskt mellan matcherna.",
    ui.ButtonSet.OK_CANCEL
  );

  if (matchResponse.getSelectedButton() !== ui.Button.OK) return;

  const matchMinutes = parseInt(matchResponse.getResponseText());

  if (isNaN(matchMinutes) || matchMinutes < 1) {
    ui.alert("❌ Ange en giltig matchlängd.");
    return;
  }


  // -------------------------------
  // TOTAL TOURNAMENT TIME
  // -------------------------------

  const totalResponse = ui.prompt(
    "Turneringstid",
    "Hur många minuter ska hela turneringen pågå?",
    ui.ButtonSet.OK_CANCEL
  );

  if (totalResponse.getSelectedButton() !== ui.Button.OK) return;

  const totalMinutes = parseInt(totalResponse.getResponseText());

  if (isNaN(totalMinutes) || totalMinutes < matchMinutes) {
    ui.alert("❌ Turneringen måste vara minst lika lång som en match.");
    return;
  }


  const numberOfRounds = Math.floor(
    (totalMinutes + BREAK_MINUTES) /
    (matchMinutes + BREAK_MINUTES)
  );


  // ============================================================
  // SETTINGS SHEET
  // ============================================================

  let settingsSheet = ss.getSheetByName("Inställningar");

  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Inställningar");
  }

  settingsSheet.clear();

  const settings = [
    ["AMERICANO – INSTÄLLNINGAR", ""],
    ["Antal spelare", numberOfPlayers],
    ["Antal banor", numberOfCourts],
    ["Matchlängd", matchMinutes],
    ["Paus mellan matcher", BREAK_MINUTES],
    ["Total turneringstid", totalMinutes],
    ["Antal omgångar", numberOfRounds]
  ];

  settingsSheet
    .getRange(1, 1, settings.length, 2)
    .setValues(settings);

  settingsSheet
    .getRange("A1:B1")
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  settingsSheet
    .getRange("A2:A7")
    .setFontWeight("bold");

  settingsSheet.autoResizeColumns(1, 2);


  // ============================================================
  // PLAYER SHEET
  // ============================================================

  let playerSheet = ss.getSheetByName("Spelare");

  if (!playerSheet) {
    playerSheet = ss.insertSheet("Spelare");
  }

  playerSheet.clear();

  playerSheet
    .getRange("A1:C1")
    .setValues([
      ["Nr", "Initialer", "Spelare"]
    ]);

  playerSheet
    .getRange("A1:C1")
    .setFontWeight("bold")
    .setBackground("#d9ead3");


  const playerRows = [];

  for (let i = 1; i <= numberOfPlayers; i++) {
    playerRows.push([
      i,
      "",
      ""
    ]);
  }

  playerSheet
    .getRange(2, 1, numberOfPlayers, 3)
    .setValues(playerRows);


  // Instructions
  const instructionRow = numberOfPlayers + 3;

  playerSheet
    .getRange(instructionRow, 1, 3, 3)
    .setValues([
      ["INFO", "", ""],
      [
        "",
        "Skriv spelarens initialer, t.ex. ES",
        "Skriv spelarens fullständiga namn"
      ],
      [
        "",
        "Om två spelare har samma initialer, använd t.ex. JS1 och JS2",
        ""
      ]
    ]);

  playerSheet
    .getRange(instructionRow, 1, 1, 3)
    .setFontWeight("bold")
    .setBackground("#fff2cc");

  playerSheet
    .getRange(instructionRow + 1, 1, 2, 3)
    .setFontStyle("italic");


  playerSheet.setColumnWidth(1, 60);
  playerSheet.setColumnWidth(2, 130);
  playerSheet.setColumnWidth(3, 220);


  ui.alert(
    "✅ Turneringen är skapad!\n\n" +
    numberOfPlayers + " spelare\n" +
    numberOfCourts + " banor\n" +
    matchMinutes + " min per match\n" +
    numberOfRounds + " omgångar\n\n" +
    "Fyll nu i initialer och namn i bladet 'Spelare'.\n\n" +
    "Exempel:\nJS | Jannik Sinner\n\n" +
    "Om två personer har samma initialer:\nJS11 och JS2"
  );

  playerSheet.activate();
}


// ============================================================
// STEP 2 — GENERATE TOURNAMENT
// ============================================================

function generateTournament() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const settingsSheet = ss.getSheetByName("Inställningar");
  const playerSheet = ss.getSheetByName("Spelare");

  if (!settingsSheet || !playerSheet) {

    ui.alert(
      "❌ Skapa först turneringen via:\n\n" +
      "🎾 Americano → 1. Skapa turnering"
    );

    return;
  }


  // ============================================================
  // READ SETTINGS
  // ============================================================

  const numberOfPlayers =
    Number(settingsSheet.getRange("B2").getValue());

  const numberOfCourts =
    Number(settingsSheet.getRange("B3").getValue());

  const matchMinutes =
    Number(settingsSheet.getRange("B4").getValue());

  const numberOfRounds =
    Number(settingsSheet.getRange("B7").getValue());


  // ============================================================
  // READ PLAYERS
  // ============================================================

  const playerData = playerSheet
    .getRange(2, 2, numberOfPlayers, 2)
    .getValues();


  const players = playerData.map(row => ({
    initials: String(row[0]).trim().toUpperCase(),
    name: String(row[1]).trim()
  }));


  // Check missing values

  for (let i = 0; i < players.length; i++) {

    if (players[i].initials === "") {

      ui.alert(
        "❌ Spelare " +
        (i + 1) +
        " saknar initialer."
      );

      return;
    }

    if (players[i].name === "") {

      ui.alert(
        "❌ Spelare " +
        (i + 1) +
        " saknar namn."
      );

      return;
    }
  }


  // Check duplicate initials

  const initials =
    players.map(p => p.initials);

  const uniqueInitials =
    new Set(initials);

  if (uniqueInitials.size !== initials.length) {

    ui.alert(
      "❌ Två spelare har samma initialer.\n\n" +
      "Använd exempelvis JS1 och JS2."
    );

    return;
  }


  // ============================================================
  // GENERATE SCHEDULE
  // ============================================================

  const schedule =
    createAmericanoSchedule(
      players,
      numberOfRounds,
      numberOfCourts
    );


  if (!schedule) {

    ui.alert(
      "❌ Kunde inte skapa schema.\n\n" +
      "Försök generera igen."
    );

    return;
  }


  // ============================================================
  // CREATE TOURNAMENT SHEET
  // ============================================================

  createTournamentSheet(
    ss,
    players,
    schedule,
    matchMinutes,
    numberOfCourts
  );


  ui.alert(
    "✅ Spelschemat är klart!\n\n" +
    "Du hittar det i bladet 'Tunering'.\n\n" +
    "Skriv resultaten manuellt i de gröna kolumnerna.\n" +
    "Poängen räknas automatiskt ihop per spelare."
  );

  ss.getSheetByName("Tunering").activate();
}


// ============================================================
// CREATE TOURNAMENT SHEET
// ============================================================

function createTournamentSheet(
  ss,
  players,
  schedule,
  matchMinutes,
  numberOfCourts
) {

  let sheet = ss.getSheetByName("Tunering");

  if (!sheet) {
    sheet = ss.insertSheet("Tunering");
  }

  sheet.clear();


  // ============================================================
  // HEADERS
  // ============================================================

  const headers = [
    "Omgång",
    "Bana",
    "Dubbelpar 1",
    "",
    "Dubbelpar 2",
    "",
    "Resultat DP1",
    "Resultat DP2"
  ];


  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange("A1:H1")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#d9ead3");


  // Merge pair headers

  sheet.getRange("C1:D1").merge();
  sheet.getRange("E1:F1").merge();


  // ============================================================
  // CREATE MATCH ROWS
  // ============================================================

  let currentRow = 2;

  const scheduleRows = [];


  for (let round = 0; round < schedule.length; round++) {

    const roundStartRow = currentRow;

    const startMinute =
      round * (matchMinutes + BREAK_MINUTES);

    const endMinute =
      startMinute + matchMinutes;

    const timeText =
      formatMinutes(startMinute) +
      "-" +
      formatMinutes(endMinute);


    for (
      let court = 0;
      court < schedule[round].length;
      court++
    ) {

      const match =
        schedule[round][court];


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


      currentRow++;
    }


    // Empty court rows if needed

    const usedCourts =
      schedule[round].length;

    for (
      let empty = usedCourts;
      empty < numberOfCourts;
      empty++
    ) {

      scheduleRows.push([
        round + 1,
        empty + 1,
        "",
        "",
        "",
        "",
        "",
        ""
      ]);

      currentRow++;
    }
  }


  sheet
    .getRange(
      2,
      1,
      scheduleRows.length,
      8
    )
    .setValues(scheduleRows);


  // ============================================================
  // FORMAT ROUND CELLS
  // ============================================================

  let row = 2;


  for (
    let round = 0;
    round < schedule.length;
    round++
  ) {

    const startMinute =
      round * (matchMinutes + BREAK_MINUTES);

    const endMinute =
      startMinute + matchMinutes;


    const timeText =
      formatMinutes(startMinute) +
      "-" +
      formatMinutes(endMinute);


    const range =
      sheet.getRange(
        row,
        1,
        numberOfCourts,
        1
      );


    range.merge();


    range
      .setValue(
        (round + 1) +
        "\n" +
        timeText
      )
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);


    // Bottom border after each round

    sheet
      .getRange(
        row + numberOfCourts - 1,
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


    row += numberOfCourts;
  }


  // ============================================================
  // RESULT INPUT
  // ============================================================

  if (scheduleRows.length > 0) {

    sheet
      .getRange(
        2,
        7,
        scheduleRows.length,
        2
      )
      .setBackground("#e2f0d9")
      .setHorizontalAlignment("center")
      .setFontWeight("bold");
  }


    // ============================================================
  // PLAYER SCORE TABLE
  // ============================================================

  const scoreStartCol = 10;

  sheet
    .getRange(
      1,
      scoreStartCol,
      1,
      4
    )
    .setValues([[
      "Initialer",
      "Spelare",
      "Antal matcher",
      "POÄNG"
    ]]);


  sheet
    .getRange(
      1,
      scoreStartCol,
      1,
      4
    )
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");


  const lastMatchRow = scheduleRows.length + 1;


  for (
    let i = 0;
    i < players.length;
    i++
  ) {

    const scoreRow = i + 2;

    const initials = players[i].initials;


    // Initials
    sheet
      .getRange(
        scoreRow,
        scoreStartCol
      )
      .setValue(initials);


    // Player name
    sheet
      .getRange(
        scoreRow,
        scoreStartCol + 1
      )
      .setValue(players[i].name);


    // ============================================================
    // NUMBER OF MATCHES
    // ============================================================

    const matchFormula =
      `=COUNTIF(C$2:C$${lastMatchRow};"${initials}")+` +
      `COUNTIF(D$2:D$${lastMatchRow};"${initials}")+` +
      `COUNTIF(E$2:E$${lastMatchRow};"${initials}")+` +
      `COUNTIF(F$2:F$${lastMatchRow};"${initials}")`;

    sheet
      .getRange(
        scoreRow,
        scoreStartCol + 2
      )
      .setFormula(matchFormula);


    // ============================================================
    // TOTAL SCORE
    // ============================================================

    const scoreFormula =
      `=SUMPRODUCT(` +
      `((C$2:C$${lastMatchRow}="${initials}")+(D$2:D$${lastMatchRow}="${initials}"));` +
      `G$2:G$${lastMatchRow}` +
      `)+` +
      `SUMPRODUCT(` +
      `((E$2:E$${lastMatchRow}="${initials}")+(F$2:F$${lastMatchRow}="${initials}"));` +
      `H$2:H$${lastMatchRow}` +
      `)`;

    sheet
      .getRange(
        scoreRow,
        scoreStartCol + 3
      )
      .setFormula(scoreFormula)
      .setBackground("#fff2cc")
      .setFontWeight("bold");
  }
  // ============================================================
  // FORMATTING
  // ============================================================

  sheet
    .getRange(
      1,
      1,
      scheduleRows.length + 1,
      8
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );


  sheet
    .getRange(
      1,
      scoreStartCol,
      players.length + 1,
      4
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );


  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 60);

  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 70);

  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 70);

  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);

  sheet.setColumnWidth(10, 80);
  sheet.setColumnWidth(11, 180);
  sheet.setColumnWidth(12, 100);
  sheet.setColumnWidth(13, 100);


  sheet
    .getRange(
      2,
      2,
      scheduleRows.length,
      7
    )
    .setHorizontalAlignment("center");


  sheet.setFrozenRows(1);
}


// ============================================================
// MAIN SCHEDULING ALGORITHM
// ============================================================

function createAmericanoSchedule(
  players,
  numberOfRounds,
  numberOfCourts
) {

  const n = players.length;

  const playersPerRound =
    Math.min(
      n,
      numberOfCourts * 4
    );


  const matchesPlayed =
    Array(n).fill(0);


  const partnerCount =
    Array.from(
      { length: n },
      () => Array(n).fill(0)
    );


  const opponentCount =
    Array.from(
      { length: n },
      () => Array(n).fill(0)
    );


  const schedule = [];


  for (
    let round = 0;
    round < numberOfRounds;
    round++
  ) {

    let bestRound = null;
    let bestScore = Infinity;


    for (
      let attempt = 0;
      attempt < 2000;
      attempt++
    ) {

      // Players with fewer matches get priority

      const candidates =
        [...Array(n).keys()];


      candidates.sort((a, b) => {

        const diff =
          matchesPlayed[a] -
          matchesPlayed[b];

        if (diff !== 0) {
          return diff;
        }

        return Math.random() - 0.5;
      });


      const selected =
        candidates.slice(
          0,
          playersPerRound
        );


      shuffleInPlace(selected);


      const matches = [];

      let score = 0;


      for (
        let i = 0;
        i < selected.length;
        i += 4
      ) {

        if (i + 3 >= selected.length) {
          break;
        }


        const four = [
          selected[i],
          selected[i + 1],
          selected[i + 2],
          selected[i + 3]
        ];


        const combinations = [

          [
            [four[0], four[1]],
            [four[2], four[3]]
          ],

          [
            [four[0], four[2]],
            [four[1], four[3]]
          ],

          [
            [four[0], four[3]],
            [four[1], four[2]]
          ]

        ];


        let bestCombination = null;
        let bestCombinationScore = Infinity;


        for (
          const combination of combinations
        ) {

          const a = combination[0][0];
          const b = combination[0][1];

          const c = combination[1][0];
          const d = combination[1][1];


          let combinationScore = 0;


          // Strongly avoid repeated partners

          combinationScore +=
            partnerCount[a][b] * 100;

          combinationScore +=
            partnerCount[c][d] * 100;


          // Also avoid repeated opponents

          combinationScore +=
            opponentCount[a][c] * 10;

          combinationScore +=
            opponentCount[a][d] * 10;

          combinationScore +=
            opponentCount[b][c] * 10;

          combinationScore +=
            opponentCount[b][d] * 10;


          if (
            combinationScore <
            bestCombinationScore
          ) {

            bestCombinationScore =
              combinationScore;

            bestCombination =
              combination;
          }
        }


        matches.push(
          bestCombination
        );

        score +=
          bestCombinationScore;
      }


      // Prefer equal number of matches

      for (
        const player of selected
      ) {

        score +=
          matchesPlayed[player] * 3;
      }


      score += Math.random();


      if (
        score <
        bestScore
      ) {

        bestScore =
          score;

        bestRound =
          matches;
      }
    }


    if (!bestRound) {
      return null;
    }


    schedule.push(
      bestRound
    );


    // ============================================================
    // UPDATE STATISTICS
    // ============================================================

    for (
      const match of bestRound
    ) {

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


      addOpponent(
        opponentCount,
        a,
        c
      );

      addOpponent(
        opponentCount,
        a,
        d
      );

      addOpponent(
        opponentCount,
        b,
        c
      );

      addOpponent(
        opponentCount,
        b,
        d
      );
    }
  }


  return schedule;
}


// ============================================================
// HELPERS
// ============================================================

function addOpponent(
  opponentCount,
  a,
  b
) {

  opponentCount[a][b]++;
  opponentCount[b][a]++;
}


function shuffleInPlace(array) {

  for (
    let i = array.length - 1;
    i > 0;
    i--
  ) {

    const j =
      Math.floor(
        Math.random() * (i + 1)
      );

    const temp =
      array[i];

    array[i] =
      array[j];

    array[j] =
      temp;
  }

  return array;
}


function formatMinutes(totalMinutes) {

  const hours =
    Math.floor(
      totalMinutes / 60
    );

  const minutes =
    totalMinutes % 60;


  return (
    String(hours).padStart(2, "0") +
    ":" +
    String(minutes).padStart(2, "0")
  );
}
