/** @OnlyCurrentDoc */
/**
 * Cheers Monday workbook - ONE-TIME SETUP for the Google Sheet.
 * Adds (1) the hover notes and (2) the dropdowns, which do not survive the
 * Excel import. Shows a small toast message, not a pop-up box.
 *
 * HOW TO RUN: in the Google Sheet choose Extensions > Apps Script, delete any
 * code in the editor, paste this whole file, click Save, choose "setupSheet",
 * click Run, and approve the permission prompt (it only edits this spreadsheet).
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NOTES = {
 "Bar Card": {
  "B4": "Home team = the team listed first on the DartConnect schedule. Choose from the list. Team rosters come from the Teams tab.",
  "B9": "Player 1 and Player 2 are set by the captains each night and decide the singles matchups (game 4 = Home P1 v Away P2, game 6 = P2 v P1, game 8 = P1 v P1, game 10 = P2 v P2). Leave blank = captain is Player 1 and the other regular is Player 2. Pick a different order or a sub from the list; a brand-new sub can be typed in, then enter their 3DA and MPR on the Handicap tab (yellow Entered cells). If a sub is Player 1 and Player 2 is blank, Player 2 defaults to the non-captain regular, so set Player 2 yourself if the sub replaces that regular.",
  "B14": "Pre-Match worksheet, Part 2. Combined MPR = the two players' MPR added (not averaged). The difference is looked up in the marks chart (Settings). The lower team receives the marks. Full steps: Handicap tab, Part 2.",
  "B15": "Pre-Match worksheet, Part 1. Team average 3DA = the two players' 3DA added and divided by 2. The difference between the two team averages is looked up in the spot chart (Settings). The lower team starts at 501 minus the spot. Full steps: Handicap tab, Part 1.",
  "B16": "Game 4 = Home Player 1 v Away Player 2. Singles percentage handicapping. Lower 3DA ÷ higher 3DA = %. Lower player starts at 501 × % (rounded); higher player starts at 501. ASSUMPTION: the worksheet does not say how to convert the percentage to a start score. Full steps: Handicap tab, Part 3.",
  "B20": "Game 11 is 701 doubles, scratch, per the Match Night worksheet."
 },
 "Settings": {
  "B4": "League-average 3DA used to scale the Cheers Rating (3DA ÷ this). Source: Division Average on the DCTV standings page, 24 Nov 2026. Update when you refresh stats.",
  "B5": "League-average MPR used to scale the Cheers Rating. Source: Division Average on the DCTV standings page, 24 Nov 2026.",
  "B6": "ASSUMPTION. How much of the rating comes from 3DA vs MPR. 50% = equal weight. Cricket is 4 of 11 games; 501 doubles and singles are 6 of 11 games; you could weight by that.",
  "B8": "ASSUMPTION. Player of the Week shrink constant K. A week's score is pulled toward 100 by darts ÷ (darts + K). With K = 60, a 60-dart week counts half. Raise it to be more cautious about small samples.",
  "B9": "ASSUMPTION. Minimum darts thrown in a week to be eligible for Player of the Week.",
  "B11": "ASSUMPTION. Players with fewer legs than this are flagged Provisional on the ranking. 18 = two full nights (9 legs per player per night).",
  "A14": "Source: your \"Pre-Match Handicap Worksheet (Doubles 501 & Cricket)\", Part 1 Step 3. Each band applies from its lower bound up to the next one. Spot is subtracted from 501 for the lower team.",
  "D14": "Source: your \"Pre-Match Handicap Worksheet\", Part 2 Step 3. Marks are given to the lower team and count toward closing numbers only."
 },
 "Players": {
  "C4": "3DA = three-dart average (points per 3 darts) from the DartConnect League Leaderboard, 'All 01' view: singles and doubles 501/701 legs. Pulled 24 Nov 2026.",
  "D4": "MPR = marks per round from the DartConnect League Leaderboard, Doubles Cricket view. Free handicap marks are not thrown darts and are not counted.",
  "E4": "Cheers Rating = 100 × [w × (3DA ÷ league 3DA) + (1−w) × (MPR ÷ league MPR)]. 100 = league average. w and the league averages are on Settings.",
  "E5": "Example: Bieber J, 3DA 54.56, MPR 2.03 → 100 × [0.5 × (54.56 ÷ 39.90) + 0.5 × (2.03 ÷ 1.35)] = 143.6.",
  "H4": "Legs played from the DCTV standings (LP). Each player throws about 9 legs a night: 7 doubles + 2 singles.",
  "I4": "Provisional = fewer legs than the minimum on Settings B11. Numbers are still settling.",
  "J4": "Optional. 3DA for a player with no league stats yet, e.g. from another league or their DartConnect profile. Used only until they have league stats. Averages from other formats may not match Cheers exactly."
 },
 "Handicap": {
  "C7": "Entered 3DA for a new player with no league stats. Must be between 10 and 100.",
  "D7": "Entered MPR for a new player with no league stats. Must be between 0.10 and 6.00.",
  "E7": "3DA used for this player. Order: 1) Entered 3DA typed in column C (new player), 2) their own league 3DA if they have any league stats, 3) Entered 3DA on the Players tab. If none, the Bar Card asks for the numbers instead of showing handicaps.",
  "G7": "Basis: Entered = you typed both numbers here. Own stats = the player's league stats. Entered (Players tab) = numbers kept on the Players tab. Needs 3DA & MPR = new player with nothing entered yet.",
  "B16": "Home team average 3DA = combined 3DA of the two players ÷ 2. Pre-Match worksheet Part 1, Step 1.",
  "B20": "Team average difference = higher team average − lower team average. Pre-Match worksheet Part 1, Step 2.",
  "B21": "Spot = looked up in the chart on Settings (A14:B21). 0.0–2.9 → 0; 3.0–5.9 → 25; 6.0–8.9 → 50; 9.0–11.9 → 75; 12.0–14.9 → 100; 15.0–17.9 → 125; 18.0–20.9 → 150; 21.0+ → 175 (cap). The lower team starts at 501 minus the spot.",
  "B23": "Lower team starts at 501 minus the spot. The alternative on the worksheet is the higher team starting at 501 plus the spot (next row). Same difference either way.",
  "B27": "Combined MPR = the two players' MPR added (not averaged). Pre-Match worksheet Part 2, Step 1.",
  "B31": "Combined MPR difference = higher team combined MPR − lower team combined MPR. Pre-Match worksheet Part 2, Step 2.",
  "B32": "Starting marks = looked up in the chart on Settings (D14:E19). 0.00–0.49 → 0; 0.50–0.99 → 1; 1.00–1.49 → 2; 1.50–1.99 → 3; 2.00–2.49 → 4; 2.50+ → 5 (cap). Given to the lower team; they count toward closing numbers only.",
  "G38": "Lower player's 3DA ÷ higher player's 3DA. Example: Cordrey 41.25 ÷ Zeitz 50.10 = 82.3%.",
  "I38": "Start score for the lower player = 501 × the percentage, rounded. The higher player starts at 501. ASSUMPTION: the Match Night worksheet gives the percentage but not this conversion.",
  "B45": "For comparison only. The Match Night worksheet, Step 2, shows \"lower ÷ higher = %\" for doubles. The Pre-Match worksheet uses difference charts instead. This workbook uses the charts."
 },
 "Season Ranking": {
  "F4": "Cheers Rating (see Players tab). 100 = league average. Combines 3DA and MPR; uses every dart a player threw (singles and doubles).",
  "I4": "Provisional = fewer legs than the minimum (Settings B11)."
 },
 "Weekly POTW": {
  "B6": "Baseline = the player's season 3DA on the Players tab. Snapshot (paste as values) before refreshing Players each week.",
  "D6": "Enter the player's 3DA for the night from DartConnect (all 01 legs).",
  "E6": "Enter the player's cricket MPR for the night.",
  "F6": "Total darts the player threw that night (used to pull a short sample toward 100).",
  "G6": "Raw score = 100 × [w × (week 3DA ÷ baseline 3DA) + (1−w) × (week MPR ÷ baseline MPR)]. 100 = playing exactly to their own average.",
  "H6": "Adjusted = 100 + (raw − 100) × darts ÷ (darts + K). Fewer darts pull the score toward 100. K is on Settings B8.",
  "I6": "Eligible = darts thrown at least the minimum on Settings B9."
 }
};
  var count = 0;
  Object.keys(NOTES).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    Object.keys(NOTES[name]).forEach(function (cell) {
      sh.getRange(cell).setNote(NOTES[name][cell]);
      count++;
    });
  });

  var bc = ss.getSheetByName('Bar Card');
  var teams = ss.getSheetByName('Teams').getRange('A5:A20');
  bc.getRange('B4:B5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(teams, true)
      .setAllowInvalid(false).setHelpText('Pick a team from the list.').build());
  bc.getRange('B9:B10').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(bc.getRange('D9:D20'), true)
      .setAllowInvalid(true).setHelpText('Optional. Blank = captain is Player 1. Pick a regular or a sub, or type a new sub.').build());
  bc.getRange('B11:B12').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(bc.getRange('E9:E20'), true)
      .setAllowInvalid(true).setHelpText('Optional. Blank = captain is Player 1. Pick a regular or a sub, or type a new sub.').build());

  // limits for entered 3DA / MPR (new players)
  var v3 = SpreadsheetApp.newDataValidation().requireNumberBetween(10, 100).setAllowInvalid(false)
    .setHelpText('3DA between 10 and 100, for example 38.5.').build();
  var vm = SpreadsheetApp.newDataValidation().requireNumberBetween(0.1, 6).setAllowInvalid(false)
    .setHelpText('MPR between 0.10 and 6.00, for example 1.40.').build();
  var hc = ss.getSheetByName('Handicap'), pl = ss.getSheetByName('Players');
  hc.getRange('C7:C10').setDataValidation(v3); hc.getRange('D7:D10').setDataValidation(vm);
  pl.getRange('J5:J40').setDataValidation(v3); pl.getRange('K5:K40').setDataValidation(vm);

  ss.toast('Added ' + count + ' notes and the dropdowns. All done.', 'Setup complete', 10);
}
