/** @OnlyCurrentDoc */
/**
 * Cheers Monday Sheet - UPDATE 28 Nov 2026 (b). Run once on the existing Sheet, then delete.
 * Throwing-order wording ("option" / cork rule) on Bar Card A21 and the How It Works tab,
 * to match the phone card, and shows Cricket (game 1) above 501 doubles on the Bar Card.
 *
 * HOW TO RUN: Extensions > Apps Script, + next to Files > Script, name it "update",
 * paste this file, Save, choose "update20261128b", Run.
 */
function update20261128b() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Bar Card').getRange('A21').setValue("Home team has the first ‘option’. Loser of the prior leg has the ‘option’ thereafter. (See How It Works.)");
  ss.getSheetByName('How It Works').createTextFinder("Who throws first in a leg is decided at the board and does not change handicaps.").replaceAllWith("Throwing order: the home team has the first ‘option’; the loser of the prior leg has the ‘option’ thereafter. ‘Option’ = the choice of who corks (throws one dart at the bull) first; closest to the bull throws first; on a tie, cork again in the reverse order. Throwing order does not change handicaps.");
  // Cricket is game 1: show cricket (row 15) above 501 doubles (row 14) on the Bar Card.
  var bc = ss.getSheetByName('Bar Card');
  if (String(bc.getRange('A14').getValue()).indexOf('501') === 0) {
    bc.getRange('A14:B14').moveTo(bc.getRange('A40:B40'));
    bc.getRange('A15:B15').moveTo(bc.getRange('A14:B14'));
    bc.getRange('A40:B40').moveTo(bc.getRange('A15:B15'));
  }
  ss.toast('Throwing-order wording updated; Cricket now above 501 doubles.', 'Update 28 Nov (b) done', 10);
}
