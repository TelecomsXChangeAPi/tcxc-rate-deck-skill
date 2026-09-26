#!/usr/bin/env python3
"""
Convert a carrier rate deck (xlsx, xls or csv) into a TCXC rate upload CSV.

    python convert_to_tcxc.py DECK [--out FILE.csv] [options]

The script finds the sheet, header row and columns on its own, prints the
mapping it used so you can check it, writes the CSV and validates it.
Override anything it gets wrong with the options (see --help).

Exit codes: 0 = written and valid, 1 = written but validation found errors,
2 = nothing written because the deck needs a decision (the message says what).
"""

import argparse
import csv
import math
import re
import statistics
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_tcxc import COLUMNS, print_report, validate  # noqa: E402

HEADER_SCAN_ROWS = 200
ASAP_WORDS = {'asap', 'immediate', 'immediately', 'now'}
DISCONTINUE_RE = re.compile(r'\b(closed|removed|removal|remove|deleted|delete|discontinued|discontinue|withdrawn|terminated|ceased)\b')
FORBID_RE = re.compile(r'\b(blocked|barred|forbidden|suspended)\b')
TRUE_WORDS = {'1', 'yes', 'y', 'true', 'x'}
CURRENCY_RE = re.compile(r'\b(EUR|GBP|INR|AED|SAR|CNY|RMB|JPY|CAD|AUD|CHF|ZAR|NGN|BRL|MXN|PKR)\b|[€£₹]')
NUMBER_RE = re.compile(r'\d[\d,]*(?:\.\d+)?(?:[eE][-+]?\d+)?|\.\d+')
# Two-digit years come before four-digit ones, and day-first before month-first:
# the first format that reads every date in the column wins.
DATE_BASES = ['%Y-%m-%d', '%d/%m/%y', '%d/%m/%Y', '%m/%d/%y', '%m/%d/%Y', '%d-%m-%y', '%d-%m-%Y',
              '%d.%m.%y', '%d.%m.%Y', '%Y/%m/%d', '%d-%b-%y', '%d-%b-%Y', '%d %b %y', '%d %b %Y',
              '%d %B %Y', '%b %d %Y', '%B %d %Y', '%Y%m%d']
DATE_FORMATS = [base + time for base in DATE_BASES for time in ('', ' %H:%M', ' %H:%M:%S')]


class DeckError(Exception):
    """The deck needs a human decision; nothing is written."""


@dataclass
class Rate:
    code: str
    effective: str
    forbidden: int
    discontinued: int
    price1: float
    price_n: float
    interval1: int
    interval_n: int
    row: str
    destination: str

    def values(self):
        return (self.effective, self.forbidden, self.discontinued, self.price1, self.price_n,
                self.interval1, self.interval_n)


# ---------------------------------------------------------------- cell helpers

def is_blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ''


def text(value):
    if is_blank(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def norm(value):
    return re.sub(r'[^a-z0-9]+', ' ', text(value).lower()).strip()


def parse_price(value):
    """Return a float, None for a blank cell, or raise ValueError."""
    if is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = NUMBER_RE.search(str(value).replace(' ', ''))
    if not match:
        raise ValueError(value)
    number = match.group(0)
    if ',' in number and '.' not in number:
        head, _, tail = number.rpartition(',')
        # 0,0123 is a decimal comma; 1,000 is a thousands separator
        number = number.replace(',', '.') if len(tail) != 3 or head == '0' else number.replace(',', '')
    else:
        number = number.replace(',', '')
    return float(number)


def parse_increment(value):
    """'60/1' -> (60, 1, 2); '60' -> (60, 60, 1); blank -> None."""
    if is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return int(value), int(value), 1
    numbers = [int(n) for n in re.findall(r'\d+', str(value))]
    if not numbers:
        raise ValueError(value)
    if len(numbers) == 1:
        return numbers[0], numbers[0], 1
    return numbers[0], numbers[1], 2


def clean_date_text(value):
    s = re.sub(r'\s+', ' ', str(value).replace(',', ' ')).strip()
    return re.sub(r'^(\d{4}-\d{2}-\d{2})T', r'\1 ', s)


def parse_date(value, fmt):
    """Return a datetime, or None for blank/ASAP. Raises ValueError if unreadable."""
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.to_pydatetime() if hasattr(value, 'to_pydatetime') else value
    if isinstance(value, (int, float)) and 20000 <= value <= 80000:  # Excel serial date
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    s = clean_date_text(text(value))
    if s.lower() in ASAP_WORDS:
        return None
    if not fmt:
        raise ValueError(value)
    parsed = datetime.strptime(s, fmt)
    if not 1990 <= parsed.year <= 2100:
        raise ValueError(value)
    return parsed


def infer_date_format(values, forced=None):
    strings = sorted({clean_date_text(text(v)) for v in values
                      if isinstance(v, str) and text(v) and clean_date_text(text(v)).lower() not in ASAP_WORDS})
    if not strings:
        return forced, None

    def reads_all(fmt):
        try:
            return all(1990 <= datetime.strptime(s, fmt).year <= 2100 for s in strings)
        except ValueError:
            return False

    if forced:
        if not reads_all(forced):
            bad = [s for s in strings if not _reads(s, forced)][:3]
            raise DeckError(f'--date-format {forced!r} does not read dates like {bad}')
        return forced, None
    matching = [fmt for fmt in DATE_FORMATS if reads_all(fmt)]
    if not matching:
        raise DeckError(f"Can't read the effective dates (e.g. {strings[:3]}); pass --date-format, "
                        f"e.g. --date-format '%d/%m/%Y'")
    chosen = matching[0]
    swapped = chosen.replace('%d', '#').replace('%m', '%d').replace('#', '%m')
    note = None
    if swapped != chosen and swapped in matching:
        note = (f'every date could be day-first or month-first (e.g. {strings[0]!r}); read as {chosen!r}. '
                f"If the deck is US-style, rerun with --date-format '{swapped}'")
    return chosen, note


def _reads(s, fmt):
    try:
        datetime.strptime(s, fmt)
        return True
    except ValueError:
        return False


# ------------------------------------------------------------- column mapping

def code_score(h):
    exact = {'complete code', 'full code', 'dial code', 'dial codes', 'dialcode', 'dialcodes', 'dialing code',
             'dialling code', 'dialing codes', 'dialling codes', 'destination code', 'destination codes', 'prefix',
             'prefixes', 'code', 'codes', 'code s', 'city code s', 'city codes', 'breakout code', 'breakout codes',
             'e164', 'e 164', 'number prefix'}
    if h in exact:
        return 3
    words = set(h.split())
    if not words or words & {'country', 'area', 'currency', 'zip', 'postal', 'rate', 'change', 'time', 'type', 'plan'}:
        return 0
    if words & {'prefix', 'prefixes', 'dial', 'dialcode', 'dialcodes'}:
        return 2
    return 1 if h.endswith((' code', ' codes', ' code s')) else 0


def rate_score(h):
    words = set(h.split())
    if h == 'price n' or not words & {'rate', 'rates', 'price', 'prices', 'cost', 'tariff', 'usd', 'eur', 'gbp', 'charge'}:
        return 0
    if words & {'change', 'changes', 'type', 'id', 'effective', 'date', 'increment', 'status', 'old', 'previous',
                'prev', 'valid', 'plan'}:
        return 0
    return 3 if h == 'price 1' else 2 + ('new' in words)


def rate_n_score(h):
    return 3 if h in ('price n', 'next price', 'subsequent price', 'additional price') else 0


def date_score(h):
    words = set(h.split())
    if words & {'end', 'expiry', 'expires', 'until', 'to', 'issued', 'issue', 'change', 'type'}:
        return 0
    if 'effective' in words or h.startswith('valid from') or h in ('start date', 'eff date', 'from date'):
        return 2
    return 1 if 'date' in words else 0


def increment_kind(h):
    words = set(h.split())
    # 'bi' is the "BI Type" (billing increment) column some carriers use; its "BI Effective Date" is excluded by 'date' below
    if not words & {'pulse', 'increment', 'increments', 'billing', 'rounding', 'interval', 'intervals', 'inc', 'bi'}:
        return None
    if words & {'price', 'rate', 'rates', 'date'}:
        return None
    if words & {'1', 'initial', 'first', 'minimum'}:
        return 'first'
    if words & {'n', 'next', 'subsequent', 'additional'}:
        return 'next'
    return 'single'


def destination_score(h):
    words = set(h.split())
    if words & {'code', 'codes', 'currency', 'change'}:
        return 0
    if words & {'destination', 'breakout'}:
        return 3
    if 'country' in words:
        return 2
    return 1 if words & {'network', 'operator', 'description', 'zone', 'carrier', 'region'} else 0


def is_status(h):
    return bool(set(h.split()) & {'change', 'changes', 'status', 'comment', 'comments', 'note', 'notes',
                                   'remark', 'remarks', 'action', 'indicator'})


def column_index(name, headers, raw_headers):
    target = norm(name)
    for i, h in enumerate(headers):
        if h == target:
            return i
    if re.fullmatch(r'[A-Za-z]{1,2}', name):  # Excel column letter
        index = 0
        for ch in name.upper():
            index = index * 26 + ord(ch) - 64
        return index - 1
    raise DeckError(f'No column named {name!r}. Columns are: {[h for h in raw_headers if h]}')


def best_column(headers, raw_headers, scorer, role, allow_tie=False):
    scored = [(scorer(h), i) for i, h in enumerate(headers) if h]
    best = max((s for s, _ in scored), default=0)
    if best == 0:
        return None
    top = [i for s, i in scored if s == best]
    if len(top) > 1 and not allow_tie:
        names = [raw_headers[i] for i in top]
        raise DeckError(f'Several columns could be the {role}: {names}. Say which with --{role.replace(" ", "-")}')
    return top[0]


def map_columns(headers, raw_headers, args):
    def given(name):
        return column_index(name, headers, raw_headers)

    mapping = {}
    if args.code:
        mapping['code'] = [given(n) for n in args.code]
    else:
        best = best_column(headers, raw_headers, code_score, 'code')
        if best is not None:
            mapping['code'] = [best]
        elif 'country code' in headers and 'area code' in headers:
            mapping['code'] = [headers.index('country code'), headers.index('area code')]
        else:
            raise DeckError(f'No code column found. Columns are: {[h for h in raw_headers if h]}; pass --code')

    mapping['rate'] = given(args.rate) if args.rate else best_column(headers, raw_headers, rate_score, 'rate')
    if mapping['rate'] is None:
        raise DeckError(f'No rate column found. Columns are: {[h for h in raw_headers if h]}; pass --rate')
    mapping['rate_n'] = given(args.rate_n) if args.rate_n else best_column(headers, raw_headers, rate_n_score, 'rate n')
    mapping['effective'] = (given(args.effective) if args.effective
                            else best_column(headers, raw_headers, date_score, 'effective', allow_tie=True))

    mapping['increment'] = mapping['interval1'] = mapping['interval_n'] = None
    if args.increment:
        mapping['increment'] = given(args.increment)
    elif args.interval1 or args.interval_n:
        if not (args.interval1 and args.interval_n):
            raise DeckError('Pass both --interval1 and --interval-n')
        mapping['interval1'], mapping['interval_n'] = given(args.interval1), given(args.interval_n)
    elif not args.increment_from_rules_only:
        kinds = defaultdict(list)
        for i, h in enumerate(headers):
            kind = increment_kind(h)
            if kind:
                kinds[kind].append(i)
        if len(kinds['first']) == 1 and len(kinds['next']) == 1:
            mapping['interval1'], mapping['interval_n'] = kinds['first'][0], kinds['next'][0]
        elif len(kinds['single']) == 1:
            mapping['increment'] = kinds['single'][0]
        elif kinds['single'] or kinds['first'] or kinds['next']:
            names = [raw_headers[i] for k in kinds.values() for i in k]
            raise DeckError(f'Not sure which column holds billing increments: {names}. '
                            f'Pass --increment, or --interval1 and --interval-n')

    if args.destination:
        mapping['destination'] = [given(n) for n in args.destination]
    else:
        scored = [(destination_score(h), i) for i, h in enumerate(headers) if h]
        best = max((s for s, _ in scored), default=0)
        if best == 3:
            mapping['destination'] = [next(i for s, i in scored if s == 3)]
        elif best == 2:  # Country plus network/description columns
            mapping['destination'] = [i for s, i in scored if s in (2, 1)]
        else:
            mapping['destination'] = [i for s, i in scored if s == 1][:1]

    used = set(mapping['code']) | {mapping['rate'], mapping['rate_n'], mapping['effective']}
    mapping['status'] = [i for i, h in enumerate(headers) if h and is_status(h) and i not in used]
    mapping['forbidden'] = next((i for i, h in enumerate(headers) if h in ('forbidden', 'blocked')), None)
    mapping['discontinued'] = next((i for i, h in enumerate(headers) if h == 'discontinued'), None)
    return mapping


# ------------------------------------------------------------------- codes

def code_tokens(cell):
    return [t for t in (p.replace(' ', '').lstrip('+') for p in re.split(r'[,;|\n]+', cell)) if t]


def infer_hyphen_mode(pairs, forced=None):
    if forced:
        return forced
    if not pairs:
        return None

    def looks_like_range(a, b):
        return len(a) == len(b) >= 4 and int(b) >= int(a) and int(b) - int(a) <= 10000

    # A column is either all ranges (447700-447709) or all joined parts (82-01), so decide once
    # for the whole column and stop when it is mixed rather than guess row by row.
    range_like = [p for p in pairs if looks_like_range(*p)]
    if len(range_like) == len(pairs):
        return 'range'
    if len(range_like) <= 0.05 * len(pairs):
        return 'concat'
    examples = ['-'.join(p) for p in range_like[:2]] + ['-'.join(p) for p in pairs if not looks_like_range(*p)][:2]
    raise DeckError(f"Hyphenated codes are mixed: some look like ranges, some like country-area joins "
                    f"(e.g. {examples}); pass --hyphen range or --hyphen concat")


def expand_token(token, hyphen_mode):
    if '-' not in token:
        return [token]
    parts = token.split('-')
    if len(parts) == 2 and all(p.isdigit() for p in parts) and hyphen_mode == 'range':
        a, b = parts
        return [str(n).zfill(len(a)) for n in range(int(a), int(b) + 1)]
    return [''.join(parts)]


def finish_code(code):
    if code.startswith('00'):
        code = code[2:]
    if not code.isdigit() or code.startswith('0'):
        raise ValueError(code)
    return code


# ----------------------------------------------------------------- the work

def locate(args):
    suffix = Path(args.deck).suffix.lower()
    found = []
    if suffix in ('.csv', '.txt', '.tsv'):
        sheets = [(Path(args.deck).name, read_csv_cells(args.deck))]
    else:
        try:
            book = pd.ExcelFile(args.deck)
        except ImportError as exc:
            raise DeckError(f'Cannot open {args.deck}: {exc}. For .xls files install xlrd or save as .xlsx')
        names = [args.sheet] if args.sheet else book.sheet_names
        for name in names:
            if name not in book.sheet_names:
                raise DeckError(f'No sheet named {name!r}; sheets are {book.sheet_names}')
        sheets = ((name, book.parse(name, header=None, dtype=object)) for name in names)
    for name, cells in sheets:
        index = find_header(cells, args.header_row)
        if index is not None:
            found.append((name, cells, index))
    if not found:
        raise DeckError("Couldn't find a header row with both a code column and a rate column. "
                        "Pass --sheet and --header-row, or --code and --rate.")
    return found[0], [f[0] for f in found[1:]]


def read_csv_cells(path):
    data = open(path, 'rb').read()
    for encoding in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            content = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    try:
        delimiter = csv.Sniffer().sniff(content[:20000], delimiters=',;\t|').delimiter
    except csv.Error:
        delimiter = ','
    rows = list(csv.reader(content.splitlines(), delimiter=delimiter))
    width = max((len(r) for r in rows), default=0)
    return pd.DataFrame([r + [''] * (width - len(r)) for r in rows], dtype=object)


def find_header(cells, header_row=None):
    candidates = [header_row - 1] if header_row else range(min(len(cells), HEADER_SCAN_ROWS))
    for i in candidates:
        if i >= len(cells):
            break
        headers = [norm(v) for v in cells.iloc[i].tolist()]
        has_code = any(code_score(h) for h in headers) or {'country code', 'area code'} <= set(headers)
        if header_row or (has_code and any(rate_score(h) for h in headers)):
            return i
    return None


def convert(args):
    now = args.now or datetime.now()
    (sheet, cells, header_index), other_sheets = locate(args)
    raw_headers = [text(v) for v in cells.iloc[header_index].tolist()]
    headers = [norm(v) for v in raw_headers]
    m = map_columns(headers, raw_headers, args)
    body = cells.iloc[header_index + 1:]

    about = []
    for i in range(header_index):
        parts = [text(v) for v in cells.iloc[i].tolist() if text(v)]
        if parts:
            about.append(' '.join(parts)[:100])
    currency_hits = {hit.group(0) for hit in CURRENCY_RE.finditer(' '.join(about + raw_headers))}

    date_format, date_note = (None, None)
    if m['effective'] is not None:
        date_format, date_note = infer_date_format(body.iloc[:, m['effective']].tolist(), args.date_format)

    hyphen_pairs = []
    for i in m['code']:
        for cell in body.iloc[:, i].tolist():
            for token in code_tokens(text(cell)):
                parts = token.split('-')
                if len(parts) == 2 and all(p.isdigit() for p in parts):
                    hyphen_pairs.append(tuple(parts))
    hyphen_mode = infer_hyphen_mode(hyphen_pairs, args.hyphen)

    rules = sorted(((norm(k), v) for k, v in args.increment_rule), key=lambda r: -len(r[0]))
    default_increment = args.default_increment
    rule_hits = Counter()
    increment_sources = Counter()
    problems = defaultdict(list)
    skipped, skipped_examples = 0, []
    rates = []
    newest_date = None

    for position, values in enumerate(body.itertuples(index=False, name=None)):
        row = f'row {header_index + 2 + position}'
        code_cells = [text(values[i]) for i in m['code']]
        if not any(code_cells):
            if any(text(v) for v in values):
                skipped += 1
                if len(skipped_examples) < 3:
                    skipped_examples.append(' | '.join(text(v) for v in values if text(v))[:90])
            continue

        destination = ' - '.join(text(values[i]) for i in m['destination'] if text(values[i]))
        status = ' '.join(text(values[i]) for i in m['status']).lower()
        rate_text = text(values[m['rate']]).lower()
        forbidden = bool(FORBID_RE.search(status) or FORBID_RE.search(rate_text)
                         or (m['forbidden'] is not None and norm(values[m['forbidden']]) in TRUE_WORDS))
        discontinued = bool(DISCONTINUE_RE.search(status) or DISCONTINUE_RE.search(rate_text)
                            or (m['discontinued'] is not None and norm(values[m['discontinued']]) in TRUE_WORDS))

        try:
            if len(m['code']) == 1:
                codes = [c for t in code_tokens(code_cells[0]) for c in expand_token(t, hyphen_mode)]
            else:
                lead = re.sub(r'\D', '', code_cells[0])
                tails = [c for t in code_tokens(code_cells[1]) for c in expand_token(t, hyphen_mode)] or ['']
                codes = [lead + re.sub(r'\D', '', t) for t in tails]
            codes = [finish_code(c) for c in codes]
        except ValueError as exc:
            problems['code is not a dial code'].append(f'{row}: {exc.args[0]!r} in {code_cells}')
            continue

        try:
            price1 = parse_price(values[m['rate']])
            price_n = parse_price(values[m['rate_n']]) if m['rate_n'] is not None else None
        except ValueError as exc:
            if forbidden or discontinued:
                price1 = price_n = None
            else:
                problems['rate is not a number'].append(f'{row}: {exc.args[0]!r}')
                continue
        if price1 is None:
            if not (forbidden or discontinued):
                problems['rate is missing'].append(f'{row}: codes {codes[:3]}')
                continue
            price1 = 0.0
        if price_n is None:
            price_n = price1

        try:
            when = parse_date(values[m['effective']], date_format) if m['effective'] is not None else None
        except ValueError:
            problems['effective date is unreadable'].append(f'{row}: {text(values[m["effective"]])!r}')
            continue
        effective = 'ASAP' if when is None or when <= now else when.strftime('%Y-%m-%d %H:%M:%S')
        if when is not None and (newest_date is None or when > newest_date):
            newest_date = when

        pair, source = None, None
        try:
            if m['increment'] is not None:
                found = parse_increment(values[m['increment']])
                pair = found[:2] if found else None
            elif m['interval1'] is not None:
                first = parse_increment(values[m['interval1']])
                later = parse_increment(values[m['interval_n']])
                if first and first[2] == 2:  # a '30/6' typed into one cell
                    pair = first[:2]
                elif later and later[2] == 2:
                    pair = later[:2]
                elif first and later:
                    pair = (first[0], later[0])
        except ValueError as exc:
            problems['billing increment is unreadable'].append(f'{row}: {exc.args[0]!r}')
            continue
        if pair:
            source = 'increment column'
        else:
            key = norm(destination)
            rule = next(((k, v) for k, v in rules if key == k or key.startswith(k + ' ')), None)
            if rule:
                pair, source = rule[1], 'rules'
                rule_hits[rule[0]] += len(codes)
            else:
                pair, source = default_increment, 'default'
        if pair[0] <= 0 or pair[1] <= 0:
            problems['billing increment must be positive'].append(f'{row}: {pair}')
            continue

        for code in codes:
            increment_sources[source] += 1
            rates.append(Rate(code, effective, int(forbidden), int(discontinued), price1, price_n,
                              pair[0], pair[1], row, destination))

    if problems and not args.skip_bad_rows:
        lines = [f'  - {msg}: {len(ex)} rows, e.g. ' + '; '.join(ex[:3]) for msg, ex in problems.items()]
        raise DeckError('Some rows could not be converted (nothing written):\n' + '\n'.join(lines) +
                        '\nFix the deck or the options, or pass --skip-bad-rows to leave them out.')

    final, conflicts, exact_duplicates = {}, [], 0
    for rate in rates:
        earlier = final.get(rate.code)
        if earlier is None:
            final[rate.code] = rate
        elif earlier.values() == rate.values():
            exact_duplicates += 1
        else:
            conflicts.append((earlier, rate))
            if args.on_duplicate == 'last':
                final[rate.code] = rate
    if conflicts and args.on_duplicate == 'error':
        shown = '; '.join(f'{a.code} ({a.row}: {a.price1:g} {a.interval1}/{a.interval_n} vs '
                          f'{b.row}: {b.price1:g} {b.interval1}/{b.interval_n})' for a, b in conflicts[:5])
        raise DeckError(f'{len(conflicts)} codes appear more than once with different values, e.g. {shown}. '
                        f'This is often peak/off-peak bands or overlapping ranges; decide which to keep '
                        f'(filter the deck, or pass --on-duplicate first/last).')

    for flag, codes in (('forbidden', args.forbid), ('discontinued', args.discontinue)):
        for code in codes:
            if code in final:
                setattr(final[code], flag, 1)
            else:
                final[code] = Rate(code, 'ASAP', int(flag == 'forbidden'), int(flag == 'discontinued'), 0.0, 0.0,
                                   default_increment[0], default_increment[1], 'added by option', '')

    if args.markup_percent or args.decimals is not None:
        decimals = 6 if args.decimals is None else args.decimals
        factor = 1 + (args.markup_percent or 0) / 100
        for rate in final.values():
            rate.price1 = round_up(rate.price1 * factor, decimals)
            rate.price_n = round_up(rate.price_n * factor, decimals)

    out = Path(args.out or f'tcxc_{re.sub(r"[^A-Za-z0-9]+", "_", Path(args.deck).stem).strip("_")}.csv')
    with open(out, 'w', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle, lineterminator='\n')
        writer.writerow(COLUMNS)
        for r in final.values():
            writer.writerow(['', '', r.code, r.effective, '', r.forbidden, r.discontinued,
                             format_price(r.price1), format_price(r.price_n), r.interval1, r.interval_n])

    warnings = []
    if date_note:
        warnings.append(date_note)
    if currency_hits:
        warnings.append(f'the deck mentions {sorted(currency_hits)}; check it is priced in the currency your '
                        f'TCXC account sells in (nothing was converted)')
    for key, _ in rules:
        if not rule_hits[key]:
            warnings.append(f'increment rule {key!r} matched no destination (check spelling against the deck)')
    if increment_sources['default'] and (m['increment'] is not None or m['interval1'] is not None or rules):
        warnings.append(f'{increment_sources["default"]} codes had no increment in the deck and got the default '
                        f'{default_increment[0]}/{default_increment[1]}')
    elif increment_sources['default'] == len(rates) and rates:
        warnings.append(f'the deck has no billing increments; every code got the default '
                        f'{default_increment[0]}/{default_increment[1]} - confirm with the carrier or its notes')
    if conflicts:
        warnings.append(f'{len(conflicts)} codes were listed more than once with different values; '
                        f'kept the {args.on_duplicate} one')
    if problems:
        warnings.append('left out rows that could not be converted: ' +
                        '; '.join(f'{msg} ({len(ex)}, e.g. {ex[0]})' for msg, ex in problems.items()))
    if newest_date is not None and newest_date < now - timedelta(days=90):
        warnings.append(f'the newest effective date in the deck is {newest_date:%Y-%m-%d} '
                        f'({(now - newest_date).days} days ago); make sure this is the carrier\'s current deck')
    warnings.extend(future_fallbacks(final.values()))
    warnings.extend(suspicious_rates(final.values()))

    report(args, sheet, header_index, other_sheets, about, raw_headers, m, date_format, hyphen_mode, hyphen_pairs,
           len(rates), final, exact_duplicates, skipped, skipped_examples, increment_sources, rule_hits, warnings, out)
    result = validate(out, now=now)
    print()
    print_report(result, out)
    return 1 if result['errors'] else 0


def future_fallbacks(rates):
    """Dated codes whose calls match a cheaper, already-live shorter prefix until their start date."""
    rates = list(rates)
    live = {r.code: r for r in rates if r.effective == 'ASAP' and not (r.forbidden or r.discontinued)}
    groups = defaultdict(list)
    for r in rates:
        if r.effective == 'ASAP' or r.forbidden or r.discontinued:
            continue
        parent = next((live[r.code[:k]] for k in range(len(r.code) - 1, 0, -1) if r.code[:k] in live), None)
        if parent and parent.price1 < r.price1:
            groups[(parent.code, r.effective)].append(r)
    lines = []
    for (parent_code, effective), members in list(groups.items())[:10]:
        parent = live[parent_code]
        codes = ', '.join(m.code for m in members[:3]) + (f' +{len(members) - 3} more' if len(members) > 3 else '')
        lines.append(f'{len(members)} codes ({codes}) only start {effective}; if they are not already live on TCXC, '
                     f'their calls match {parent_code} at {parent.price1:g} until then '
                     f'(they are priced from {min(m.price1 for m in members):g}) - ask the user whether to start them ASAP')
    return lines


def suspicious_rates(rates):
    groups = defaultdict(list)
    for r in rates:
        if r.destination and not (r.forbidden or r.discontinued) and r.price1 > 0:
            groups[norm(re.split(r'\s*-\s*', r.destination)[0])].append(r)
    found = []
    for country, members in groups.items():
        if len(members) < 4:
            continue
        # Compare with the country's lower quartile, not its median: decks often price whole
        # networks at blocking levels, which drags the median far above normal rates.
        prices = sorted(r.price1 for r in members)
        lower_quartile = statistics.quantiles(prices, n=4)[0]
        for r in members:
            # Only a lone outlier looks like a typo; a group of cheap codes is usually real fixed-line pricing.
            if r.price1 < lower_quartile * 0.05 and sum(p <= r.price1 * 2 for p in prices) <= 3:
                found.append((r, country, lower_quartile))
    lines = [f'{r.code} ({r.destination}) at {r.price1:g} is far below the rest of {country} '
             f'(lower quartile {quartile:g}); normal for some fixed lines, but check it is not a carrier error '
             f'before selling (block with --forbid {r.code} until confirmed)'
             for r, country, quartile in found[:10]]
    if len(found) > 10:
        lines.append(f'...and {len(found) - 10} more rates far below the rest of their country')
    zero = [r.code for r in rates if r.price1 == 0 and not (r.forbidden or r.discontinued)]
    if zero:
        lines.append(f'{len(zero)} active codes have a 0 price, e.g. {zero[:5]}')
    return lines


def round_up(value, decimals):
    scale = 10 ** decimals
    return math.ceil(round(value * scale, 6)) / scale


def format_price(value):
    s = f'{value:.10f}'.rstrip('0').rstrip('.')
    return s or '0'


def report(args, sheet, header_index, other_sheets, about, raw_headers, m, date_format, hyphen_mode, hyphen_pairs,
           rate_count, final, exact_duplicates, skipped, skipped_examples, increment_sources, rule_hits, warnings, out):
    def names(indices):
        return ', '.join(repr(raw_headers[i]) for i in indices) or '-'

    print(f'Deck:     {args.deck} (sheet {sheet!r}, header on row {header_index + 1})')
    if other_sheets:
        print(f'          other sheets that also look like rate tables: {other_sheets} (use --sheet to pick one)')
    for line in about[:12]:
        print(f'  about:  {line}')
    increments = ('-' if m['increment'] is None and m['interval1'] is None else
                  names([m['increment']]) if m['increment'] is not None else names([m['interval1'], m['interval_n']]))
    print(f'Columns:  code <- {names(m["code"])} | rate <- {names([m["rate"]])}'
          + (f' | next-interval rate <- {names([m["rate_n"]])}' if m['rate_n'] is not None else '')
          + f' | effective <- {names([m["effective"]]) if m["effective"] is not None else "- (all ASAP)"}'
          + (f' ({date_format})' if date_format else '')
          + f' | increments <- {increments} | destination <- {names(m["destination"])}'
          + (f' | status <- {names(m["status"])}' if m['status'] else ''))
    if hyphen_mode and hyphen_pairs:
        a, b = hyphen_pairs[0]
        shown = f'{a}-{b} -> {a}..{b}' if hyphen_mode == 'range' else f'{a}-{b} -> {a}{b}'
        print(f'Codes:    hyphenated codes read as {"ranges" if hyphen_mode == "range" else "joined parts"} (e.g. {shown})')
    rows = list(final.values())
    print(f'Rows:     {rate_count} codes from the deck -> {len(rows)} TCXC rows'
          + (f'; {exact_duplicates} exact duplicates dropped' if exact_duplicates else '')
          + (f'; {skipped} row{"s" if skipped != 1 else ""} without a code skipped, e.g. {skipped_examples}' if skipped else ''))
    effective = Counter('ASAP' if r.effective == 'ASAP' else 'dated' for r in rows)
    print(f'Effective from: ' + ', '.join(f'{k} {v}' for k, v in effective.most_common()))
    intervals = Counter(f'{r.interval1}/{r.interval_n}' for r in rows)
    print(f'Increments: ' + ', '.join(f'{k}: {v}' for k, v in intervals.most_common())
          + ' (from ' + ', '.join(f'{k} {v}' for k, v in increment_sources.most_common()) + ')')
    if rule_hits:
        top = ', '.join(f'{k} {v}' for k, v in rule_hits.most_common(12))
        more = f', +{len(rule_hits) - 12} more' if len(rule_hits) > 12 else ''
        print(f'Rules:    {top}{more}')
    print(f'Flags:    Forbidden {sum(r.forbidden for r in rows)}, Discontinued {sum(r.discontinued for r in rows)}')
    if args.markup_percent:
        print(f'Markup:   +{args.markup_percent}% on every price, rounded up')
    if warnings:
        print('Check before upload:')
        for w in warnings:
            print(f'  - {w}')
    print(f'Wrote:    {out}')


def parse_rule(value):
    name, sep, increment = value.rpartition('=')
    if not sep or not name.strip():
        raise argparse.ArgumentTypeError(f'expected DESTINATION=FIRST/NEXT, e.g. "Brazil=30/6", got {value!r}')
    return name.strip(), parse_pair(increment)


def parse_pair(value):
    try:
        found = parse_increment(value)
    except ValueError:
        found = None
    if not found:
        raise argparse.ArgumentTypeError(f'expected an increment like 60/1, got {value!r}')
    return found[:2]


def parse_codes(values):
    codes = []
    for value in values or []:
        chunk = open(value[1:]).read() if value.startswith('@') else value
        codes.extend(finish_code(re.sub(r'[\s+-]', '', c)) for c in re.split(r'[,;\n]+', chunk) if c.strip())
    return codes


def parse_now(value):
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    raise argparse.ArgumentTypeError('expected YYYY-MM-DD or YYYY-MM-DD HH:MM:SS')


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('deck', help='carrier rate deck: .xlsx, .xls or .csv (a TCXC file to repair also works)')
    parser.add_argument('--out', help='output CSV path (default: tcxc_<deck name>.csv)')
    where = parser.add_argument_group('where the table is')
    where.add_argument('--sheet', help='sheet name (default: first sheet with a code and a rate column)')
    where.add_argument('--header-row', type=int, help='1-based row number of the column headers')
    cols = parser.add_argument_group('columns (header text or Excel letter; auto-detected when omitted)')
    cols.add_argument('--code', nargs='+', metavar='COL', help='code column; give two to join them, e.g. "Country Code" "Area Code"')
    cols.add_argument('--rate', metavar='COL', help='price per minute column')
    cols.add_argument('--rate-n', metavar='COL', help='price for later intervals, if the deck has a separate one')
    cols.add_argument('--effective', metavar='COL', help='effective date column (none: every rate is ASAP)')
    cols.add_argument('--date-format', help="strptime format when dates are misread, e.g. '%%m/%%d/%%Y'")
    cols.add_argument('--increment', metavar='COL', help='billing increment column holding values like 60/1')
    cols.add_argument('--interval1', metavar='COL', help='first-increment column (use with --interval-n)')
    cols.add_argument('--interval-n', metavar='COL', help='next-increment column (use with --interval1)')
    cols.add_argument('--destination', nargs='+', metavar='COL', help='destination name column(s), used for rules and warnings')
    inc = parser.add_argument_group('billing increments not given per row')
    inc.add_argument('--increment-rule', action='append', type=parse_rule, default=[], metavar='DEST=FIRST/NEXT',
                     help='increment for destinations starting with DEST, e.g. "Brazil=30/6" (repeatable; longest match wins)')
    inc.add_argument('--default-increment', type=parse_pair, default=(1, 1), metavar='FIRST/NEXT',
                     help='increment for codes with no column value or rule (default 1/1)')
    inc.add_argument('--increment-from-rules-only', action='store_true',
                     help='ignore any increment column and use only rules and the default')
    codes = parser.add_argument_group('codes and flags')
    codes.add_argument('--hyphen', choices=['range', 'concat'],
                       help='read 447700-447709 as a range, or 82-01 as 8201 (default: decided from the whole column)')
    codes.add_argument('--forbid', nargs='+', default=[], metavar='CODE', help='codes to block (Forbidden=1); @file reads a list')
    codes.add_argument('--discontinue', nargs='+', default=[], metavar='CODE',
                       help='codes to withdraw (Discontinued=1); @file reads a list')
    codes.add_argument('--on-duplicate', choices=['error', 'first', 'last'], default='error',
                       help='a code listed twice with different values: stop (default) or keep the first/last')
    codes.add_argument('--skip-bad-rows', action='store_true', help='leave out rows that cannot be converted instead of stopping')
    price = parser.add_argument_group('prices (only when the user asks)')
    price.add_argument('--markup-percent', type=float, help='add a margin, e.g. 11 for +11%%; prices are rounded up')
    price.add_argument('--decimals', type=int, help='round prices up to this many decimals (default with markup: 6)')
    parser.add_argument('--now', type=parse_now, metavar='DATE',
                        help='treat this date (YYYY-MM-DD[ HH:MM:SS]) as today, for reproducible runs')
    args = parser.parse_args()
    try:
        args.forbid = parse_codes(args.forbid)
        args.discontinue = parse_codes(args.discontinue)
    except OSError as exc:
        parser.error(f'cannot read code list: {exc}')
    except ValueError as exc:
        parser.error(f'not a dial code: {exc.args[0]!r}')

    try:
        sys.exit(convert(args))
    except DeckError as exc:
        print(f'NOT CONVERTED: {exc}', file=sys.stderr)
        sys.exit(2)


if __name__ == '__main__':
    main()
