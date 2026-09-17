# TCXC Rate Deck: a Claude skill

Turn any carrier's voice rate deck into a [TelecomsXchange (TCXC)](https://www.telecomsxchange.com) rate upload file, with Claude.

Carrier price lists arrive in every shape: notice blocks above the table, codes packed into lists and ranges, billing increments hidden in footnotes, removed codes on a separate sheet. This skill teaches Claude the exact TCXC upload format and gives it a tested converter and validator, so the file it hands back can be uploaded as it is, together with a short report of anything that needs your decision.

## What it does

- Converts `.xlsx`, `.xls` and `.csv` rate decks into the 11-column TCXC upload CSV.
- Finds the sheet, header row and columns on its own, and lets you override any of them.
- Handles code lists and ranges, separate country and area code columns, `+` and `00` prefixes, day-first and US dates, and decimal commas.
- Splits billing increments such as `60/1` into Interval 1 and Interval N, or applies increment rules written in a deck's notes.
- Carries removed and blocked codes through as Discontinued and Forbidden.
- Checks and repairs existing TCXC files, such as ones edited in Excel.
- Stops and asks when a deck is ambiguous (two price columns, codes that could be ranges or country-area pairs) instead of guessing.
- Flags what needs a human decision: rates far below the rest of their country, decks that look out of date, future-dated codes whose calls fall back to a cheaper prefix, and assumed default increments.
- Never rounds a price down. A markup is applied only when you ask for one, and rounds up.

## The TCXC upload format

```
Country,Description,Prefix,Effective from,Rate Id,Forbidden,Discontinued,Price 1,Price N,Interval 1,Interval N
,,447700,ASAP,,0,0,0.0125,0.0125,1,1
,,4915,2027-01-01 00:00:00,,0,0,0.0421,0.0421,60,1
,,5511,ASAP,,0,0,0.0098,0.0098,30,6
```

One row per dial code. Prefix is digits only; Effective from is `ASAP` or `YYYY-MM-DD HH:MM:SS`; Forbidden and Discontinued are `0` or `1`; prices are plain decimals per minute; intervals are whole seconds. [SKILL.md](skills/tcxc-rate-deck/SKILL.md) explains every column and how TCXC applies effective dates.

## Install

The skill is the folder [`skills/tcxc-rate-deck`](skills/tcxc-rate-deck).

### Claude Code

Run these two commands inside Claude Code:

```
/plugin marketplace add TelecomsXChangeAPi/tcxc-rate-deck-skill
/plugin install tcxc-rate-deck@tcxc
```

The scripts need Python with pandas and openpyxl (`pip install pandas openpyxl`); Claude will tell you if they are missing.

Prefer a plain skill folder? Copy `skills/tcxc-rate-deck` into `~/.claude/skills/` (or a project's `.claude/skills/`).

### Claude apps (claude.ai and desktop)

Download `tcxc-rate-deck.zip` from the [latest release](https://github.com/TelecomsXChangeAPi/tcxc-rate-deck-skill/releases/latest) and upload it under **Settings > Features**. Custom skills are available on Pro, Max, Team and Enterprise plans with code execution enabled.

### Claude API

Upload the `skills/tcxc-rate-deck` folder through the Skills API and use it with the code execution tool. See [Using Agent Skills with the API](https://platform.claude.com/docs/en/build-with-claude/skills-guide).

## Use it

Ask Claude in plain language and give it the file:

```
Convert ~/Downloads/supplier_rates.xlsx into a TCXC upload file. Keep the carrier's exact prices, no markup. Save it to ~/Downloads and tell me anything I should check before uploading.
```

Or call the skill by name: `/tcxc-rate-deck:tcxc-rate-deck` when installed as a plugin, `/tcxc-rate-deck` when installed as a folder.

```
/tcxc-rate-deck:tcxc-rate-deck ~/Downloads/supplier_rates.xlsx
```

Claude reads the deck, runs the converter, checks the result against the deck and reports what it found and what needs your decision.

### Example prompts

- "Convert ~/Downloads/amendment.xlsx to TCXC format with an 11% markup, and use the billing increments from the notes in the sheet"
- "Check ~/Downloads/my_upload.csv against the carrier's original ~/Downloads/price_list.xlsx and give me a fixed file if anything is wrong"
- "Our supplier sent a new A-Z price list, prepare it for upload to TCXC: ~/Downloads/new_supplier_rates.xlsx"

### Follow-ups that work well

- "block 24997 until the carrier confirms"
- "add an 11% markup"
- "use the peak rate column"
- "start the future-dated codes ASAP"
- "the Brazil rows should be 60/60, redo it"

Prompts work best with the full file path, whether you want exact prices or a markup, and where to save the result.

The scripts also work on their own:

```bash
python skills/tcxc-rate-deck/scripts/convert_to_tcxc.py deck.xlsx --out tcxc_upload.csv
python skills/tcxc-rate-deck/scripts/validate_tcxc.py tcxc_upload.csv
```

Run `convert_to_tcxc.py --help` for every option.

## Example

[`examples/acme_rate_notice.xlsx`](examples/acme_rate_notice.xlsx) is a fictional supplier deck with a notice block, code lists, ranges, text dates, a billing column and a closed code. [`examples/acme_rate_notice_tcxc.csv`](examples/acme_rate_notice_tcxc.csv) is the converter's output for it, produced with `--now 2026-09-17` so the dates stay reproducible.

## Requirements

Python 3.8 or newer with pandas and openpyxl (plus xlrd for old `.xls` files). Tested with Python 3.14, pandas 3.0 and openpyxl 3.1. The scripts only read the deck you give them and write a CSV; they make no network calls.

## Before you upload

The skill catches a lot, but you are responsible for the rates you publish. Read the report, look at any flagged rows, and keep the carrier's original deck.

## Development

```bash
pip install pandas openpyxl pytest
pytest
```

## Contributing

Issues and pull requests are welcome, especially rate deck layouts the converter does not handle yet. Please do not attach real carrier price lists: describe the layout, or share a copy with the codes and prices replaced.

## License

MIT. See [LICENSE](LICENSE).
