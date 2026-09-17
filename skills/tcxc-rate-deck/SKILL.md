---
name: tcxc-rate-deck
description: Convert any carrier or vendor voice rate deck (price list, rate sheet, A-Z file or rate amendment in xlsx, xls or csv, from any wholesale carrier) into a TelecomsXchange (TCXC) rate upload CSV that can be uploaded directly, and check or repair existing TCXC upload files. Use this whenever someone mentions TCXC or TelecomsXchange rates or routes, the TCXC format, columns like Price 1, Price N, Interval 1, Interval N, Forbidden or Discontinued, or wants a carrier price list turned into an upload file, even if they only say "convert this rate sheet", "prepare this deck for upload" or "check my rates file before I upload it".
license: MIT
compatibility: Needs Python 3.8+ with pandas and openpyxl (xlrd for old .xls files).
---

# TCXC rate deck conversion

Sellers on TelecomsXchange (TCXC) publish voice routes by uploading a rate CSV. Carriers send price lists in every shape: preambles above the table, codes split across columns or packed into lists, billing increments hidden in footnotes. The job is to turn whatever arrives into a file the seller can upload as-is, and to catch the mistakes that cost money once it is live.

## The upload format

Exactly these 11 columns, in this order, one row per dial code, comma-separated UTF-8:

```
Country,Description,Prefix,Effective from,Rate Id,Forbidden,Discontinued,Price 1,Price N,Interval 1,Interval N
,,447700,ASAP,,0,0,0.0125,0.0125,1,1
,,4915,2027-01-01 00:00:00,,0,0,0.0421,0.0421,60,1
,,5511,ASAP,,0,0,0.0098,0.0098,30,6
```

| Column | Value | Why it matters |
|---|---|---|
| Country, Description | empty | Accepted uploads leave them blank. |
| Prefix | full dial code, digits only: country code plus area or network code, no `+`, `00`, spaces or hyphens | Calls route on the longest matching prefix, so a truncated or mangled code sells the wrong destination at this price. |
| Effective from | `ASAP`, or a future date as `YYYY-MM-DD HH:MM:SS` | See "How TCXC applies Effective from". |
| Rate Id | empty | |
| Forbidden | `0`, or `1` to block the route | Blocked routes leave Market View immediately. Never blank. |
| Discontinued | `0`, or `1` to withdraw the code | Also immediate with ASAP. Never blank. |
| Price 1 | price per minute for the first interval, plain decimal such as `0.0123` | No currency symbols, thousands separators, decimal commas or scientific notation (`1e-05`). |
| Price N | price per minute for later intervals | Same as Price 1 unless the carrier prices later intervals differently. |
| Interval 1 | first billing increment in whole seconds | A carrier increment of `60/1` means Interval 1 = 60, Interval N = 1. |
| Interval N | every later increment in whole seconds | Never write `60/1` into one cell: the upload rejects it or bills wrong. |

### How TCXC applies Effective from

- `ASAP` applies on upload. Decreases take effect immediately. Increases are detected by TCXC's upload logic and moved to 7 days from upload; the current rate keeps selling until then, then the new rate replaces it.
- A future date starts the rate at that time.
- If the carrier's date is today or already past, write `ASAP`: a start time that has passed means "now", and ASAP says so unambiguously.

TCXC enforces notice periods itself, so keep the carrier's dates and never add notice periods yourself.

## Workflow

### 1. Look at the deck first

Open the file and read it the way a person would before running anything. These change the conversion:

- **Sheet and header row.** Decks often have a logo or notice block above the table and several sheets (rates, code changes, time bands).
- **What the deck is.** Product or route name, issue date, validity dates, currency. Carriers send near-identical files for different products or routes (e.g. a "Standard" and a "Premium" list); converting the wrong one uploads the wrong prices. Always tell the user which product and date you converted.
- **Currency.** Prices must be in the currency the seller's TCXC account sells in. If the deck is in another currency, stop and ask; never convert currencies silently.
- **Billing increments.** Usually a column (`Pulse`, `Billing`, `Increment`, values like `60/1`). Some carriers only state them in notes, e.g. "the following countries will be billed 60/60: ...". Read the notes: they override the default for the destinations they name, and the names must match the deck's destination names.
- **Removed or blocked codes.** A status column (Closed, Deleted, Removed, Blocked) or a separate code-changes sheet. These are part of the deck: removed codes go out as `Discontinued=1` (dated when the deck gives a removal date) and blocked ones as `Forbidden=1`, so the upload carries the carrier's removals instead of silently dropping them.
- **Time bands.** If a code has peak and off-peak rates, TCXC takes one price per code: ask which band to use (usually the flat or standard one).

`references/deck-layouts.md` shows how common layouts map onto the converter's options; read it when a deck has notes-based increments, code-change sheets, split code columns or anything else unusual.

### 2. Convert with the bundled script

```bash
python scripts/convert_to_tcxc.py DECK.xlsx --out tcxc_<carrier>_<product>_<yyyymmdd>.csv
```

(paths are relative to this skill's directory; needs Python 3 with pandas and openpyxl)

It detects the sheet, header row and columns, then prints the mapping it used, a summary and a "Check before upload" list, writes the CSV and validates it. Compare the printed mapping with what you saw in step 1 and override anything wrong:

| Option | Use when |
|---|---|
| `--sheet NAME`, `--header-row N` | the wrong sheet or header row was picked |
| `--code COL [COL]` | the code column was misdetected; give two columns to join them, e.g. `--code "Country Code" "Area Code"` |
| `--rate COL`, `--rate-n COL` | several price columns (current vs new, peak vs off-peak), or a separate later-interval price |
| `--effective COL`, `--date-format FMT` | the date column or its format was misread, e.g. US dates `--date-format '%m/%d/%Y'` |
| `--increment COL`, or `--interval1 COL --interval-n COL` | the increment column(s) were misdetected |
| `--increment-rule "Brazil=30/6"` (repeat), `--default-increment 1/1` | increments come from notes: a rule applies to destinations whose name starts with the given text (case and punctuation ignored, longest match wins) |
| `--hyphen range` or `--hyphen concat` | codes like `447700-447709` (a range) vs `82-01` (country and area, joined to 8201); the script decides per column and stops if unsure |
| `--discontinue CODE...` / `--forbid CODE...` (or `@file` with one code per line) | codes removed or blocked outside the rate table, e.g. on a code-changes sheet |
| `--on-duplicate first\|last` | the same code appears twice with different values; the default stops and shows them so you can decide |
| `--skip-bad-rows` | only after reading why rows failed and confirming they are junk (repeated headers, section titles) |
| `--markup-percent P`, `--decimals N` | only when the user asks for a margin or rounding; prices are rounded up so they never fall below cost |

Exit code 0 means written and valid, 1 means written but invalid, 2 means nothing was written because the deck needs a decision; the message says what.

When a layout is beyond the options (a PDF table, codes in free text, one column per country), reshape it into a simple table first (a code or code-list column, rate, effective date, increment, destination, status) and run the converter on that, so the output format stays identical.

### 3. Check the result against the deck

The built-in validation proves the file is well-formed, not that it matches the carrier's prices. Before handing it over:

- Row count: codes in the deck (after expanding lists and ranges) versus rows written.
- A handful of rows compared by eye with the deck, covering each increment type, a dated row and any flagged row.
- Every "Check before upload" line. Rules that matched nothing usually mean a spelling mismatch with the deck's destination names, which silently leaves those codes on the default increment.
- Codes with a future start date. Until that date, calls to them match the longest prefix that is already live, often a cheap country-wide code (a new mobile range falling back to the fixed-line catch-all). The converter lists these; ask the user whether to start them ASAP, especially on a first upload from a new supplier.

### 4. Report to the user

Keep it short and lead with what they need to act on:

- Output file path and row count.
- Which deck it came from: carrier, product, issue or validity date, currency. Say so if the deck looks old (the converter warns when its newest date is over 90 days back): uploading a stale deck can overwrite newer prices.
- Increments: counts per value and where they came from (column, the deck's notes, or an assumed default).
- Effective from: how many ASAP and how many dated.
- Anything set to Forbidden or Discontinued, and why.
- Decisions for the user: rates far below the rest of their country (a possible carrier typo; suggest blocking with Forbidden=1 until the carrier confirms), zero prices, assumed default increments, ambiguous dates, future-dated codes that fall back to a cheaper live prefix.

## Checking or repairing an existing TCXC file

Run `python scripts/validate_tcxc.py FILE.csv`. Errors break the upload or bill wrong; warnings need a judgement call.

If the carrier's original deck is available, convert from it rather than patching the file: hand edits drift in ways a format check cannot see, such as `60/1` typed as `60/60` or prices rounded down below cost. Compare the two and tell the user what differed.

Without the original, run the converter on the TCXC file itself. It recognises the TCXC columns, splits `30/6` typed into one cell, fills blank flags with 0, drops extra columns (such as ASR/ACD statistics from a portal export) and rewrites prices as plain decimals. Tell the user what could not be verified without the carrier's deck.

## Pitfalls

- **Excel.** Opening the CSV in Excel is harmless; saving it is not. Excel rewrites prices to the displayed decimals, can turn long codes into `4.47E+11`, and switches to Windows line endings. If someone must edit in Excel, validate again afterwards.
- **Rounding.** Round only when asked, and round up: a price rounded down sells below cost on every minute.
- **Dates.** `03/04/2026` is 3 April in most carrier decks and 4 March in US ones. The converter picks one format for the whole column and warns when every date is ambiguous.
- **Default increments.** If a deck states no increment anywhere, 1/1 is a common default but still an assumption; say so in the report.
