/** @OnlyCurrentDoc */
/**
 * Cheers Monday Sheet - UPDATE 28 Nov 2026 (run once on the existing Sheet).
 * Adds limits for entered new-player numbers, matching the phone card:
 *   3DA 10 to 100, MPR 0.10 to 6.00.
 * Applies to Handicap C7:D10 (Entered 3DA / MPR) and Players J5:K40.
 * Also adds hover notes on Handicap C7 and D7 explaining the limits.
 *
 * HOW TO RUN: Extensions > Apps Script, click + next to Files > Script,
 * name it "update", paste this file, Save, choose "update20261128", Run.
 */
function update20261128() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var v3 = SpreadsheetApp.newDataValidation().requireNumberBetween(10, 100).setAllowInvalid(false)
    .setHelpText('3DA between 10 and 100, for example 38.5.').build();
  var vm = SpreadsheetApp.newDataValidation().requireNumberBetween(0.1, 6).setAllowInvalid(false)
    .setHelpText('MPR between 0.10 and 6.00, for example 1.40.').build();
  var hc = ss.getSheetByName('Handicap'), pl = ss.getSheetByName('Players');
  hc.getRange('C7:C10').setDataValidation(v3); hc.getRange('D7:D10').setDataValidation(vm);
  pl.getRange('J5:J40').setDataValidation(v3); pl.getRange('K5:K40').setDataValidation(vm);
  hc.getRange('C7').setNote('Entered 3DA for a new player with no league stats. Must be between 10 and 100.');
  hc.getRange('D7').setNote('Entered MPR for a new player with no league stats. Must be between 0.10 and 6.00.');
  ss.toast('Limits added: 3DA 10-100, MPR 0.10-6.00.', 'Update 28 Nov done', 10);
}
