# Source files

Everything needed to rebuild the Bar Card and the Google Sheet workbook.

| Folder | What it is |
|---|---|
| `card/bar_card_template.html` | The Bar Card page (layout, rules, calculations). |
| `card/data.json` | Stats snapshots, schedule, rosters, captains, chart settings. Updated each week. |
| `card/build_card.py` | Combines the two into `index.html` at the top of the repo (the live site). |
| `card/test_calc.js` | Quick check that the math still matches the workbook (Rensch 433, Dirks 485, Cordrey 413, Dirks 461). |
| `workbook/build_workbook.py` | Builds `Cheers_Monday_Handicap_and_Ranking.xlsx` (needs Python + openpyxl). |
| `workbook/setup_sheet.gs` | Apps Script that adds hover notes and dropdowns after importing into Google Sheets. |
| `data/` | DartConnect League Portal exports (stats, schedule, rosters/captains; no contact info). |

## Rules (card and workbook must agree)
- 501 doubles: team average 3DA difference -> spot chart; lower team starts at 501 minus spot.
- Cricket: combined MPR difference -> marks chart; lower team gets free marks.
- Singles: lower 3DA / higher 3DA x 501, rounded.
- Players with any league stats use them. New players need 3DA and MPR entered.
- Makeup matches use the stats snapshot from before their scheduled night.
- League format: SISO.
