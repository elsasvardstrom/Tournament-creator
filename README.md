# 🎾 Americano Tournament Generator

A universal **Americano tennis tournament generator for Google Sheets**, built with Google Apps Script.

It creates a randomized doubles schedule based on the number of players, courts, match length and total tournament length.

## Features

- Choose the number of players
- Choose the number of available courts
- Choose match length
- Choose total tournament length
- Automatic 5-minute breaks between rounds
- Enter player names and unique initials
- Automatically generates doubles matches
- Tries to distribute matches fairly
- Tries to avoid repeated partners
- Tries to avoid repeated opponents
- Automatically rotates resting players when not everyone can play at once
- Enter match scores manually
- Automatically calculates each player's total points
- No player names or Google Sheet IDs are hard-coded

## How the tournament works

Each doubles match contains four players.

If there are more players than available court spots, some players rest during each round.

Example:

| Players | Courts | Playing each round | Resting |
|---:|---:|---:|---:|
| 16 | 4 | 16 | 0 |
| 18 | 4 | 16 | 2 |
| 21 | 5 | 20 | 1 |
| 23 | 5 | 20 | 3 |
| 24 | 6 | 24 | 0 |

The generator prioritizes players who have played fewer matches, which helps distribute rest rounds fairly.

## Installation

1. Create a new Google Sheet.
2. Open **Extensions → Apps Script**.
3. Delete the default code in `Code.gs`.
4. Copy the contents of [`Code.gs`](Code.gs) from this repository into the Apps Script editor.
5. Save the project.
6. Return to the Google Sheet and reload the page.
7. A new menu called **🎾 Americano** will appear.

The first time you run the script, Google will ask you to authorize the Apps Script project.

## Usage

### 1. Set up the tournament

Choose:

**🎾 Americano → 1. Set up tournament**

You will be asked for:

- Number of players
- Number of courts
- Match length in minutes
- Total tournament length in minutes

A 5-minute break is automatically added between rounds.

The script creates:

- `Settings`
- `Players`

### 2. Enter players

In the `Players` sheet, enter a unique set of initials and the player's name.

Example:

| No. | Initials | Player |
|---:|---|---|
| 1 | JS | Jannik Sinner |
| 2 | CA | Carlos Alcaraz |
| 3 | RF | Roger Federer |
| 4 | RN | Rafael Nadal |

If two players would normally have the same initials, use unique versions such as:

```text
JS1
JS2
```

Initials are used in the tournament schedule to keep it compact.

### 3. Generate the schedule

Choose:

**🎾 Americano → 2. Generate schedule**

The script creates a `Tournament` sheet.

Example:

| Round | Court | Team 1 |  | Team 2 |  | Score T1 | Score T2 |
|---|---:|---|---|---|---|---:|---:|
| 1 | 1 | JS | RF | CA | RN | 21 | 14 |
|  | 2 | ND | BS | AM | HR | 17 | 21 |

### 4. Enter scores

Enter each team's score manually in the green columns:

- `Score T1`
- `Score T2`

The standings table on the right updates automatically.

Each player receives the score earned by their team in that match.

For example, if:

```text
JS + RF = 21
CA + RN = 14
```

then:

- JS receives 21 points
- RF receives 21 points
- CA receives 14 points
- RN receives 14 points

## Scheduling logic

The generator uses a scoring-based randomized search.

It strongly penalizes:

1. Repeated partners
2. Repeated opponents

It also prioritizes players with fewer completed/scheduled matches when deciding who plays in each round.

Because the schedule is randomized, generating the tournament again may produce a different valid schedule.

## Files

```text
americano-tournament-generator/
├── Code.gs
├── appsscript.json
├── .gitignore
├── LICENSE
└── README.md
```

## Notes

- Minimum number of players: **4**
- Every active court requires exactly **4 players**
- The script automatically limits the number of usable courts if there are not enough players
- Match times are currently shown as elapsed tournament time, starting at `00:00`
- The 5-minute break between rounds is fixed in `CONFIG.BREAK_MINUTES`

## Customization

You can change the break length near the top of `Code.gs`:

```javascript
const CONFIG = {
  BREAK_MINUTES: 5,
  ...
};
```

You can also change sheet names and formatting colors in the same configuration object.

## Privacy

This repository does not contain:

- Real participant lists
- Google Sheet IDs
- API keys
- Account credentials

Player data stays in the user's own Google Sheet.

## License

MIT License.
