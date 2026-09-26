"""Tests for the tcxc-rate-deck skill scripts. Run with: pytest"""

import csv
import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills' / 'tcxc-rate-deck' / 'scripts'
EXAMPLES = ROOT / 'examples'
CONVERT = SCRIPTS / 'convert_to_tcxc.py'
VALIDATE = SCRIPTS / 'validate_tcxc.py'
HEADER = 'Country,Description,Prefix,Effective from,Rate Id,Forbidden,Discontinued,Price 1,Price N,Interval 1,Interval N'


def run(*args, cwd):
    return subprocess.run([sys.executable, *map(str, args)], cwd=cwd, capture_output=True, text=True)


def rows(path):
    return {r['Prefix']: r for r in csv.DictReader(open(path, newline=''))}


def write_csv(path, text):
    path.write_text(text)
    return path


def test_example_deck_converts_to_expected_file(tmp_path):
    out = tmp_path / 'out.csv'
    result = run(CONVERT, EXAMPLES / 'acme_rate_notice.xlsx', '--out', out, '--now', '2026-09-17', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert out.read_text() == (EXAMPLES / 'acme_rate_notice_tcxc.csv').read_text()
    assert 'their calls match 49 at 0.0078' in result.stdout  # future-dated codes fall back to a cheaper prefix


def test_example_output_passes_validation(tmp_path):
    result = run(VALIDATE, EXAMPLES / 'acme_rate_notice_tcxc.csv', cwd=tmp_path)
    assert result.returncode == 0, result.stdout
    assert 'RESULT: VALID' in result.stdout


def test_code_lists_ranges_dates_and_closed_codes(tmp_path):
    out = rows(EXAMPLES / 'acme_rate_notice_tcxc.csv')
    assert len(out) == 25
    assert {'49172', '49173', '49174', '447400', '447409', '33607', '521'} <= set(out)
    assert (out['49152']['Interval 1'], out['49152']['Interval N']) == ('60', '1')
    assert out['49152']['Effective from'] == '2026-09-23 00:00:00'  # future date kept, read day-first
    assert out['49']['Effective from'] == 'ASAP'  # past date
    assert (out['234805']['Discontinued'], out['234805']['Effective from']) == ('1', '2026-09-20 00:00:00')


def test_joined_country_and_area_columns_with_decimal_commas(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv',
                     'Destination;Country Code;Area Code;Current Rate;New Rate;Effective Date;Billing Increment\n'
                     'Spain Mobile;34;6, 7;0,0100;0,0120;03/04/2027;60/60\n')
    out = tmp_path / 'out.csv'
    result = run(CONVERT, deck, '--out', out, '--now', '2026-09-17', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    got = rows(out)
    assert set(got) == {'346', '347'}
    assert got['346']['Price 1'] == '0.012'  # New Rate chosen over Current Rate
    assert 'day-first or month-first' in result.stdout  # ambiguous dates are called out


def test_stops_when_several_price_columns_could_apply(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv', 'Code,Rate Peak,Rate Off-peak\n44,0.01,0.005\n')
    result = run(CONVERT, deck, '--out', tmp_path / 'out.csv', cwd=tmp_path)
    assert result.returncode == 2
    assert '--rate' in result.stderr
    assert not (tmp_path / 'out.csv').exists()


def test_stops_on_mixed_hyphenated_codes(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv', 'Code,Rate\n4912-4919,0.01\n52-334217,0.02\n')
    result = run(CONVERT, deck, '--out', tmp_path / 'out.csv', cwd=tmp_path)
    assert result.returncode == 2
    assert '--hyphen' in result.stderr


def test_country_area_hyphens_are_joined(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv', 'Destination,City Code(s),Price($)\nKorea-Mobile,82-10,0.03\nPeru-Lima,51-1,0.004\n')
    out = tmp_path / 'out.csv'
    result = run(CONVERT, deck, '--out', out, cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert set(rows(out)) == {'8210', '511'}


def test_increment_rules_and_unmatched_rule_warning(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv',
                     'Destination,Code,Rate\nBrazil-Sao Paulo,5511,0.01\nNepal-Ncell Mobile,97798,0.1\nNepal-Other,977,0.08\n')
    out = tmp_path / 'out.csv'
    result = run(CONVERT, deck, '--out', out, '--increment-rule', 'Brazil=30/6',
                 '--increment-rule', 'Nepal Ncell Mobile=60/1', '--increment-rule', 'Korea South=60/1', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    got = rows(out)
    assert (got['5511']['Interval 1'], got['5511']['Interval N']) == ('30', '6')
    assert (got['97798']['Interval 1'], got['97798']['Interval N']) == ('60', '1')
    assert (got['977']['Interval 1'], got['977']['Interval N']) == ('1', '1')
    assert "'korea south' matched no destination" in result.stdout


def test_markup_rounds_up_without_float_noise(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv', 'Code,Rate\n44,0.0415\n33,0.1\n')
    out = tmp_path / 'out.csv'
    result = run(CONVERT, deck, '--out', out, '--markup-percent', '11', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    got = rows(out)
    assert got['44']['Price 1'] == '0.046065'
    assert got['33']['Price 1'] == '0.111'


def test_repairs_a_hand_edited_tcxc_file(tmp_path):
    broken = write_csv(tmp_path / 'broken.csv',
                       HEADER + ',Last Day ASR\r\n'
                       ',,5531,ASAP,,,,0.0170,0.0170,30/6,30/6,\r\n'
                       ',,4930,ASAP,,,,0.0120,0.0120,1,1,\r\n')
    check = run(VALIDATE, broken, cwd=tmp_path)
    assert check.returncode == 1
    assert 'Interval 1 must be a whole number' in check.stdout
    assert 'Forbidden must be 0 or 1' in check.stdout

    fixed = tmp_path / 'fixed.csv'
    result = run(CONVERT, broken, '--out', fixed, cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert fixed.read_text().splitlines()[0] == HEADER
    got = rows(fixed)
    assert (got['5531']['Interval 1'], got['5531']['Interval N']) == ('30', '6')
    assert got['4930']['Forbidden'] == got['4930']['Discontinued'] == '0'


def test_flags_a_rate_far_below_the_rest_of_its_country(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Destination', 'Code', 'Rate', 'Pulse'])
    for code, dest, rate in [('25470', 'Kenya-Mobile A', 0.21), ('25471', 'Kenya-Mobile A', 0.21),
                             ('25472', 'Kenya-Mobile B', 0.24), ('254', 'Kenya-Fixed', 0.25),
                             ('25479', 'Kenya-Mobile C', 0.0009)]:
        ws.append([dest, code, rate, '60/60'])
    wb.save(tmp_path / 'deck.xlsx')
    result = run(CONVERT, tmp_path / 'deck.xlsx', '--out', tmp_path / 'out.csv', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert '25479 (Kenya-Mobile C) at 0.0009 is far below' in result.stdout


def test_old_deck_warning(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv', 'Code,Rate,Effective Date\n44,0.01,2025-01-01\n')
    result = run(CONVERT, deck, '--out', tmp_path / 'out.csv', '--now', '2026-09-17', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'newest effective date in the deck is 2025-01-01' in result.stdout


def test_bi_type_column_is_the_increment(tmp_path):
    deck = write_csv(tmp_path / 'deck.csv',
                     'Destination,City Code(s),Price($),Effective Date,Comments,BI Type,BI Effective Date,Service Level\n'
                     'Afghanistan-Mobile A,93-76,0.1500,21-Sep-18,No Change,1/1,19-Oct-07,Standard\n'
                     'Afghanistan-Mobile B,93-70,0.1400,06-Feb-25,No Change,60/1,25-Dec-25,Standard\n'
                     'Vietnam-Mobile,84-35,0.0500,10-May-22,No Change,60/60,01-Jan-22,Standard\n')
    out = tmp_path / 'out.csv'
    result = run(CONVERT, deck, '--out', out, '--hyphen', 'concat', '--now', '2026-09-19', cwd=tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "increments <- 'BI Type'" in result.stdout
    got = rows(out)
    assert (got['9376']['Interval 1'], got['9376']['Interval N']) == ('1', '1')
    assert (got['9370']['Interval 1'], got['9370']['Interval N']) == ('60', '1')
    assert (got['8435']['Interval 1'], got['8435']['Interval N']) == ('60', '60')
    assert "effective <- 'Effective Date'" in result.stdout
    assert 'default' not in result.stdout.split('Increments:')[1].split('\n')[0]
