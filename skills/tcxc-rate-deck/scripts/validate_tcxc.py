#!/usr/bin/env python3
"""
Check a CSV against the TCXC rate upload format.

    python validate_tcxc.py FILE.csv

Exit code 0 means the file is valid (warnings may still need a look),
1 means it has errors that would break the upload or bill wrongly.
"""

import csv
import re
import statistics
import sys
from collections import Counter, defaultdict
from datetime import datetime

COLUMNS = ['Country', 'Description', 'Prefix', 'Effective from', 'Rate Id', 'Forbidden',
           'Discontinued', 'Price 1', 'Price N', 'Interval 1', 'Interval N']
# Origin-based rates (OBR): the same with the caller's origin after Prefix.
OBR_COLUMNS = COLUMNS[:3] + ['CLI Prefix'] + COLUMNS[3:]
# What the switch falls back to when a caller's origin can't be used.
OBR_FALLBACKS = ('OBR_INVALID_CLI_RATE', 'OBR_UNKNOWN_COUNTRY_RATE', 'OBR_GLOBAL_DEFAULT_RATE')

DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$')
PRICE_RE = re.compile(r'^\d+(\.\d+)?$')
INTERVAL_RE = re.compile(r'^[1-9]\d*$')
MAX_EXAMPLES = 3


def validate(path, now=None):
    """Return a dict with errors, warnings (message -> examples) and stats."""
    now = now or datetime.now()
    errors = defaultdict(list)
    warnings = defaultdict(list)
    stats = {'rows': 0, 'effective': Counter(), 'intervals': Counter(), 'forbidden': 0,
             'discontinued': 0, 'prices': []}
    result = {'errors': errors, 'warnings': warnings, 'stats': stats}

    raw = open(path, 'rb').read()
    if raw.startswith(b'\xef\xbb\xbf'):
        warnings['File starts with a byte-order mark (usually added by Excel); save it without one'].append('line 1')
    if b'\r\n' in raw:
        warnings['Windows line endings, usually a sign the file was saved from Excel; '
                 'make sure prices and codes were not rewritten'].append('whole file')
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        errors['File is not UTF-8 text'].append('whole file')
        return result

    rows = list(csv.reader(text.splitlines()))
    if not rows:
        errors['File is empty'].append('whole file')
        return result

    header = rows[0]
    obr = header[:len(OBR_COLUMNS)] == OBR_COLUMNS
    columns = OBR_COLUMNS if obr else COLUMNS
    stats['obr'] = obr
    if header[:len(columns)] != columns:
        errors['Header must start with exactly: ' + ','.join(COLUMNS) +
               ' (or, for origin-based rates: ' + ','.join(OBR_COLUMNS) + ')'].append('line 1: ' + ','.join(header))
        return result
    if len(header) > len(columns):
        warnings['Extra columns after Interval N are not part of the upload format; remove them'].append(
            'line 1: ' + ', '.join(header[len(columns):]))

    first_line = {}
    origin_prices = defaultdict(dict)  # prefix -> {CLI prefix: price 1}, origin-based rates only
    for n, row in enumerate(rows[1:], start=2):
        if not any(cell.strip() for cell in row):
            warnings['Blank line'].append(f'line {n}')
            continue
        if len(row) < len(columns):
            errors[f'Row has fewer than {len(columns)} fields'].append(f'line {n}')
            continue
        cli = ''
        if obr:
            cli = row[3]
            row = row[:3] + row[4:]
        _, _, prefix, effective, rate_id, forbidden, discontinued, price1, price_n, interval1, interval_n = row[:11]
        stats['rows'] += 1

        def bad(message, value):
            errors[message].append(f'line {n}: {value!r}')

        key = (prefix, cli) if obr else prefix
        if obr and not (cli in OBR_FALLBACKS or (cli.isdigit() and len(cli) <= 10)):
            bad('CLI Prefix must be up to 10 digits (the caller\'s country or range) or one of ' +
                ', '.join(OBR_FALLBACKS), cli)
        if not prefix.isdigit():
            bad('Prefix must be digits only (country code + area code, no +, 00, spaces or hyphens)', prefix)
        elif prefix.startswith('0'):
            bad('Prefix must not start with 0 (strip the international or trunk prefix)', prefix)
        elif key in first_line:
            if obr:
                bad('Prefix and CLI Prefix appear together more than once', f'{prefix}/{cli} (first on line {first_line[key]})')
            else:
                bad(f'Prefix appears more than once', f'{prefix} (first on line {first_line[key]})')
        else:
            first_line[key] = n

        if effective == 'ASAP':
            stats['effective']['ASAP'] += 1
        elif DATE_RE.match(effective):
            try:
                when = datetime.strptime(effective, '%Y-%m-%d %H:%M:%S')
                stats['effective']['dated'] += 1
                if when <= now:
                    warnings['Effective from is already in the past; use ASAP'].append(f'line {n}: {effective!r}')
            except ValueError:
                bad('Effective from is not a real date', effective)
        else:
            bad('Effective from must be ASAP or YYYY-MM-DD HH:MM:SS', effective)

        if rate_id and not rate_id.isdigit():
            bad('Rate Id must be empty (or a TCXC rate id)', rate_id)

        for name, value in (('Forbidden', forbidden), ('Discontinued', discontinued)):
            if value not in ('0', '1'):
                bad(f'{name} must be 0 or 1 (never blank)', value)
        stats['forbidden'] += forbidden == '1'
        stats['discontinued'] += discontinued == '1'
        active = forbidden != '1' and discontinued != '1'

        prices_ok = True
        for name, value in (('Price 1', price1), ('Price N', price_n)):
            if not PRICE_RE.match(value):
                prices_ok = False
                bad(f'{name} must be a plain decimal like 0.0123 (no symbols, commas or scientific notation)', value)
            elif len(value.partition('.')[2]) > 8:
                warnings[f'{name} has more than 8 decimals (float noise?)'].append(f'line {n}: {value!r}')
        if prices_ok:
            stats['prices'].append(float(price1))
            if obr and active:
                origin_prices[prefix][cli] = float(price1)
            if active and float(price1) == 0:
                warnings['Price is 0 on a route that is not Forbidden or Discontinued'].append(f'line {n}: {prefix}')
            if float(price1) != float(price_n):
                warnings['Price N differs from Price 1 (fine only if the carrier prices later intervals differently)'].append(f'line {n}: {prefix}')

        intervals_ok = True
        for name, value in (('Interval 1', interval1), ('Interval N', interval_n)):
            if not INTERVAL_RE.match(value):
                intervals_ok = False
                bad(f'{name} must be a whole number of seconds (e.g. 60, never 60/1)', value)
        if intervals_ok:
            stats['intervals'][f'{interval1}/{interval_n}'] += 1

    # A fallback cheaper than a real origin lets a caller pay less by hiding or
    # garbling their caller ID.
    for prefix, prices in origin_prices.items():
        real = {cli: p for cli, p in prices.items() if cli not in OBR_FALLBACKS}
        if not real:
            continue
        top_cli = max(real, key=real.get)
        for fallback in OBR_FALLBACKS:
            if fallback in prices and prices[fallback] < real[top_cli]:
                warnings['Fallback rate cheaper than a real origin (callers could hide their origin to pay less)'].append(
                    f'{prefix}: {fallback} {prices[fallback]:g} < origin {top_cli} {real[top_cli]:g}')

    if stats['rows'] == 0:
        errors['File has no rate rows'].append('whole file')
    return result


def print_report(result, path, out=sys.stdout):
    stats, errors, warnings = result['stats'], result['errors'], result['warnings']
    print(f'TCXC format check: {path}', file=out)
    if stats['rows']:
        effective = ', '.join(f'{k} {v}' for k, v in stats['effective'].most_common())
        intervals = ', '.join(f'{k}: {v}' for k, v in stats['intervals'].most_common())
        kind = 'origin-based rates, ' if stats.get('obr') else ''
        print(f'  {kind}rows {stats["rows"]} | Effective from: {effective} | Forbidden {stats["forbidden"]} '
              f'| Discontinued {stats["discontinued"]}', file=out)
        print(f'  intervals {intervals}', file=out)
        if stats['prices']:
            prices = stats['prices']
            print(f'  prices min {min(prices):g}, median {statistics.median(prices):g}, max {max(prices):g}', file=out)
    for title, found in (('ERRORS (the upload would fail or bill wrong)', errors), ('WARNINGS', warnings)):
        if found:
            print(title + ':', file=out)
            for message, examples in found.items():
                shown = '; '.join(examples[:MAX_EXAMPLES])
                more = f' (+{len(examples) - MAX_EXAMPLES} more)' if len(examples) > MAX_EXAMPLES else ''
                print(f'  - {message}: {len(examples)} x, e.g. {shown}{more}', file=out)
    print('RESULT: ' + ('INVALID' if errors else 'VALID'), file=out)


def main():
    if len(sys.argv) != 2:
        print(__doc__.strip())
        sys.exit(2)
    result = validate(sys.argv[1])
    print_report(result, sys.argv[1])
    sys.exit(1 if result['errors'] else 0)


if __name__ == '__main__':
    main()
