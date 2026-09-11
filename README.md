# 🎾 Americano Tournament Generator

A Google Sheets tournament generator for Americano tennis tournaments, built with Google Apps Script.

## Features

- Choose number of players
- Choose number of courts
- Choose match duration
- Choose total tournament duration
- Automatic 5-minute breaks between matches
- Enter player names and initials
- Automatically generates randomized doubles matches
- Tries to distribute matches fairly between players
- Avoids repeated partners and opponents where possible
- Automatically calculates each player's total score

## How to use

1. Create a new Google Sheet.
2. Go to Extensions → Apps Script.
3. Copy the contents of `Code.gs` into the Apps Script editor.
4. Save the project.
5. Reload the Google Sheet.
6. Open the `🎾 Americano` menu.
7. Select `1. Skapa turnering`.
8. Enter the tournament settings.
9. Enter player names and initials in the `Spelare` sheet.
10. Select `2. Generera spelschema`.
11. Enter match scores in the generated tournament sheet.

## Player initials

Each player must have unique initials.

Example:

JS – Jannik Sinner
CA – Carlos Alcaraz

If two players have the same initials, use numbers:

JS1  
JS2

## Scoring

Enter each team's score manually in the tournament sheet.

The total score for each individual player is calculated automatically.
