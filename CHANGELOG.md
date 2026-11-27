# Changelog

Every release of the Bar Card, newest first. The version is shown at the bottom of the card, and each weekly handicap audit records which version the league was using.

To release: bump `VERSION` in `source/card/bar_card_template.html`, add an entry here, and use the version in the commit message.

## v0.9.0 beta – 27 Nov 2026
First version for league use (Monday 30 Nov 2026).

- **Game order:** Cricket first, then 501 doubles, 501 singles and 701, each with "How it's calculated".
- **Lookup charts:** spot and marks charts with tonight's row highlighted.
- **Singles rule confirmed:** 501 × (lower 3DA ÷ higher 3DA), per the DartConnect worksheet.
- **New players:** enter name, 3DA (10–100) and MPR (0.10–6.00) in place; the old "assumed strong" handicap is removed.
- **Layout:** HOME/AWAY labels, captain marked (C), full names, uniform result tiles, collapsible rules, phone layout fixes, light/dark header.
- **Throwing order:** 'option' note in the Cricket panel, defined in the rules (cork; ties re-cork in reverse order).
- **Version and beta notice** shown on the card.
- Stats through Mon 23 Nov 2026 (2 nights).

## v0.1 beta – 26 Nov 2026 (unversioned)
First public test on GitHub Pages.

- Night and matchup from the DartConnect schedule; lineup with captain as Player 1.
- 501 doubles spot, cricket marks, singles starts; "How it's worked out" math.
- New players and subs treated as strong players (singles scratch).
- Dartboard header, light/dark mode, beta tag, QR code.

## Planned for v1.0
- Rules committee sign-off on the open items in DECISIONS.md.
- A few clean weekly handicap audits.
