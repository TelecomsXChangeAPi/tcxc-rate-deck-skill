# Common rate deck layouts

How the shapes carriers actually send map onto `convert_to_tcxc.py`. Find the closest match, then check the converter's printed mapping against the deck. All names, codes and prices below are illustrative.

## Contents
1. One code per row with an increment column
2. Increments only in the deck's notes
3. Removals on a separate code-changes sheet
4. Code lists and ranges in one cell
5. Country code and area code in separate columns
6. Several price columns or time bands
7. An existing TCXC file or portal export
8. Anything else (PDF, free text, wide tables)
9. Full A-Z amendments with a BI Type column and a code-changes sheet

## 1. One code per row with an increment column

Typical of large A-Z price lists: a notice block with customer, product name, validity dates and currency, then a table such as

`Destination | Country Code | Area Code | Complete Code | Rate (USD / Min) | Valid From | Rate Change | Pulse | Notes`

Usually no options are needed: the converter prefers a full-code column over split country/area columns, reads the rate column even when its header is a leftover spreadsheet formula, and splits `Pulse` values such as `60/1`. Report the product name from the notice block, because carriers send near-identical lists for different products. Some workbooks carry a numbering-plan sheet with a change indicator per code (New, Deleted...): convert deleted codes as in section 3.

## 2. Increments only in the deck's notes

Some amendment files have no increment column. The billing rules sit in a note, sometimes inside the destination column below the table, for example:

> Calls to the following destinations are billed 60/60: Maldives, Samoa. Billed 60/1: Vietnam, Sri Lanka Dialog Mobile. Billed 30/6: Brazil. All other destinations are billed per second.

Turn each named destination into a rule and let everything else take the default:

```bash
python scripts/convert_to_tcxc.py amendment.xlsx \
  --increment-rule "Maldives=60/60" --increment-rule "Samoa=60/60" \
  --increment-rule "Vietnam=60/1" --increment-rule "Sri Lanka Dialog Mobile=60/1" \
  --increment-rule "Brazil=30/6" --default-increment 1/1
```

Rules match destination names that start with the rule text, ignoring case and punctuation, so `Sri Lanka Dialog Mobile` matches only `Sri Lanka-Dialog Mobile` while `Sri Lanka` would match every Sri Lanka destination. Use the deck's wording: when a note names a sub-network, make the rule just as specific, and spell names the way the destination column does (a rule `Samoa` does not match a destination called `Western Samoa`). The converter lists rules that matched nothing (usually a spelling difference such as `Korea South` vs `South Korea`); fix them before uploading. If the notes give no default for everything else, say in the report that 1/1 was assumed, and point out destinations the notes may have missed (satellite and international networks are common gaps).

Codes in these files are often written country-area with a hyphen (`93-70`); the converter joins them (9370) when the column is clearly country-area pairs.

## 3. Removals on a separate code-changes sheet

Amendments often include a sheet like `Destination | Effective Date | Code(s) | Comments` where comments say something like `Code Addition`, `Code Removal` or `No Change`. Additions normally already appear in the rate table. Removals must become Discontinued rows:

```python
import openpyxl
rows = list(openpyxl.load_workbook('amendment.xlsx', read_only=True, data_only=True)['Code Changes'].iter_rows(values_only=True))
header = next(i for i, r in enumerate(rows) if r and 'Code(s)' in r)
col = {name: j for j, name in enumerate(rows[header]) if name}
removed = [str(r[col['Code(s)']]).replace('-', '') for r in rows[header + 1:]
           if r and 'removal' in str(r[col['Comments']]).lower()]
open('removed_codes.txt', 'w').write('\n'.join(removed))
```

then add `--discontinue @removed_codes.txt`. The option writes them as ASAP; if a removal date is still in the future, add those codes to the rate table instead (code, rate, the removal date, status `Closed`) so the date carries through.

## 4. Code lists and ranges in one cell

Smaller carriers put several codes per row: `49152, 49162, 49172-49174` or `234803; 234806`. The converter splits on commas, semicolons, pipes and line breaks, removes spaces and `+`, strips a leading `00`, and expands `447400-447409` into ten codes when every hyphenated code in the column looks like a range. A column mixing ranges and country-area pairs stops the conversion: pass `--hyphen range` or `--hyphen concat` after looking at the deck.

## 5. Country code and area code in separate columns

With no full-code column, the converter joins `Country Code` and `Area Code` automatically. For other names pass both in order: `--code "CC" "City Code"`. Area cells may hold lists or ranges; each is prefixed with the country code.

## 6. Several price columns or time bands

`Current Rate | New Rate` picks `New Rate` automatically. Peak/off-peak, CLI/non-CLI or several product columns stop the conversion; ask the user which price they sell, then pass `--rate`. If a code appears once per time band, keep only the band the user sells (filter the deck, or `--on-duplicate`) and say so.

## 7. An existing TCXC file or portal export

Files already in TCXC columns convert as they are: `Prefix`, `Price 1`, `Price N`, `Effective from`, `Interval 1`/`Interval N` and the flag columns are recognised. The converter splits `30/6` typed into one interval cell, turns blank flags into 0, drops extra columns such as `Last Day ASR ... Statistics Time` from a portal export, and normalises prices. It cannot tell whether a hand-edited increment or rounded price still matches the carrier, so compare with the carrier's deck whenever you can get it.

## 8. Anything else

For a PDF, free text or a wide table with one column per country, build a plain table first (one row per code or code list with rate, effective date, increment, destination and status), save it as CSV, and run the converter on that. Checking your intermediate table against the source is part of the job: spot-check rows before converting.

## 9. Full A-Z amendments with a BI Type column and a code-changes sheet

Some large carriers send every amendment as the full A-Z (100,000 rows or more) even when only a handful of rows change. Typical shape: sheets `Rate Changes` and `Code Changes` plus several empty sheets, the header 20 or so rows down after a notice block naming the product, customer and `Date:`, and columns like:

`Destination | City Code(s) | Price($) | Effective Date | Comments | BI Type | BI Effective Date | Service Level`

Conversion: `--hyphen concat` when codes are written as country plus area (`93-76`; `229-0145` keeps its leading zero), increments from `BI Type` (detected automatically; values such as `1/1`, `60/60`, `60/1`, `30/6`), and `Comments` is the status column. Legal text after the table is skipped as rows without a code. Removals often sit only on the `Code Changes` sheet (`Code Removed`), see section 3.

### Patterns that look like errors but are usually normal on these decks

Mention them in one line at most, under "normal for this carrier", not in the decisions list, unless the user says otherwise:

- **Blocking-level prices on networks the product doesn't carry.** A few mobile networks priced at 1 to 3 per minute while the country's fixed or "Other" rate is a fraction of a cent. The within-country outlier check will list the cheap codes; they are the real routes, the expensive ones are a deterrent. Do not suggest blocking the cheap ones.
- **The code-changes sheet lists whole destinations.** When a few codes are added, every code of that destination appears with `No Change`, and the matching rate rows say something like `Code Change ... No Rate Change` with a future effective date. Only rows marked `Code Added` or `Code Removed` are actual changes.
- **Effective dates that go backwards.** A destination announced as `Pending Code Change On <date>` in one amendment shows its original, older effective date again once the change has taken effect. That is the change completing, not a reversal.
- **Codes appear and disappear between amendments without a code-changes entry**, sometimes backdated, and destinations are renamed at the same price. A code removed from a specific destination is then billed at the shorter prefix's rate.
- **Billing increment changes with `No Change` in Comments**; the change shows only in `BI Type` and `BI Effective Date`, sometimes at short notice.
- **Codes added under a network name price differently from the country's `Other` row** while the comment says `No Rate Change`.
- **Satellite and international network codes** (870, 881, 882, 883) at several units per minute, stray spaces or blanks in text columns, and 10-digit national numbering that makes codes look long.

### What is still worth a line in the report

Rows whose `Comments` say `Rate Increase` or `Rate Decrease`, rows marked `Code Added` or `Code Removed`, and increment changes visible in `BI Effective Date` after the amendment date. When the previous amendment is available, diff the two on code, price and `BI Type` and report only those differences; that is a far shorter and more useful list than the outlier checks on a deck this shape.
