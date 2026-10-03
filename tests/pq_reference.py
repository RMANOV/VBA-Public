"""Narrow source-bound reference checks, NOT a Power Query M evaluator.

Only the documented baseline/fixed pagination guards are modeled. Unknown
forms fail closed. Grammar must be checked separately with an M parser;
native Excel/Power BI evaluation remains a separate verification step.
"""
import os
from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def source(name):
    ref = os.environ.get('VBA_PQ_REF')
    if ref:
        return subprocess.check_output(
            ['git', '-C', str(ROOT), 'show', ref + ':power-query/' + name],
            text=True,
        )
    return (ROOT / 'power-query' / name).read_text()


def pagination_bindings():
    text = source('16_API_PAGINATION.pq')
    if not re.search(r'PageSize\s*=\s*100\s*,', text):
        raise ValueError('Unsupported fixture page size')
    offset = text.split('// VARIANT: CURSOR-BASED PAGINATION', 1)[0]
    cursor = text.split('// CursorPages = List.Generate(', 1)[1].split('// )', 1)[0]
    link = text.split('// LinkPages = List.Generate(', 1)[1].split('// )', 1)[0]
    bindings = {}
    for name, block in [('offset', offset), ('cursor', cursor), ('link', link)]:
        block = re.sub(r'^\s*// ?', '', block, flags=re.M) if name != 'offset' else block
        conditions = re.findall(r'each\s+(\[hasMore\]|\[receivedCount\]\s*>\s*0|not\s+\[isTerminal\])\s*,', block)
        if len(conditions) != 1:
            raise ValueError('Unsupported emission condition: ' + name)
        fixed = conditions[0] != '[hasMore]'
        expected = '[receivedCount] > 0' if name == 'offset' else 'not [isTerminal]'
        if fixed and re.sub(r'\s+', ' ', conditions[0]) != expected:
            raise ValueError('Wrong variant emission guard')
        if name == 'offset':
            if 'List.Count(page[data]) >= PageSize' not in block:
                raise ValueError('Unsupported page-size predicate')
            if fixed and not re.search(r'each\s+if\s+\[hasMore\]\s+then', block):
                raise ValueError('Missing guarded next fetch')
            if fixed and not re.search(r'else\s*\[\s*pageData\s*=\s*\{\},\s*currentOffset\s*=\s*\[currentOffset\],\s*receivedCount\s*=\s*0,\s*hasMore\s*=\s*false\s*\]', block, re.S):
                raise ValueError('Missing empty terminal sentinel')
            if 'page       = FetchPage(nextOffset)' not in block:
                raise ValueError('Unsupported offset fetch')
        else:
            token = 'nextCursor' if name == 'cursor' else 'nextUrl'
            if fixed and not re.search(r'each\s+if\s+\[' + token + r'\]\s*<>\s*null\s+then', block):
                raise ValueError('Missing token guard')
            if fixed and ('isTerminal = false' not in block or 'isTerminal = true' not in block):
                raise ValueError('Missing fetched/terminal state distinction')
            if fixed and not re.search(r'else\s*\[pageData = \{\},\s*' + token + r' = null,\s*isTerminal = true\]', block):
                raise ValueError('Unsupported token terminal record')
        if not re.search(r'each\s+\[pageData\]', block):
            raise ValueError('Unsupported page selector')
        bindings[name] = fixed
    bindings['typed_empty'] = bool(re.search(
        r'Result\s*=\s*if\s+List.IsEmpty\(CombinedRecords\)\s+then\s+#table\(\s*type table\s*\[id = Int64.Type, name = text, created_at = datetime\],\s*\{\}\s*\)\s*else\s*Table.FromRecords\(CombinedRecords\)', text, re.S))
    return bindings


def paginate(kind, pages, initial=0):
    fixed = pagination_bindings()[kind]
    requests = []
    internal_transitions = 0
    def fetch(key):
        requests.append(key)
        if key not in pages:
            raise AssertionError('Unexpected request: ' + repr(key))
        return pages[key]
    value = fetch(initial)
    out = []
    key = initial
    terminal = False
    while True:
        data, token = (value, None) if kind == 'offset' else value
        more = len(data) >= 100 if kind == 'offset' else token is not None
        accepted = (len(data) > 0 if kind == 'offset' else not terminal) if fixed else more
        if not accepted:
            break
        out.extend(data)
        internal_transitions += 1
        if internal_transitions > 10:
            raise AssertionError('Non-finite reference transitions')
        if fixed and not more:
            value = [] if kind == 'offset' else ([], None)
            terminal = True
        else:
            key = key + 100 if kind == 'offset' else token
            value = fetch(key)
    # Public transitions count request advances; a no-HTTP sentinel transition
    # is reported separately and never changes expected request counts.
    return dict(rows=out, requests=requests, transitions=len(requests)-1), internal_transitions


def empty_schema():
    columns = ['id', 'name', 'created_at'] if pagination_bindings()['typed_empty'] else []
    return dict(rows=[], columns=columns, missing=[x for x in ['id', 'name', 'created_at'] if x not in columns])


def currency_beta():
    text = source('06_CUSTOM_M_FUNCTIONS.pq')
    raw = re.search(r'\[Product = "Beta",\s*Amount = "([^"]+)"', text).group(1)
    sep = re.search(r'each fnCleanCurrency\(_, "([.,])"\)', text).group(1)
    if 'allowedChars = {"0","1","2","3","4","5","6","7","8","9","-", sep}' not in text:
        raise ValueError('Unsupported currency normalization')
    cleaned = ''.join(ch for ch in raw.strip() if ch in '0123456789-' + sep)
    return float(cleaned.replace(',', '.') if sep == ',' else cleaned)


def currency_eu():
    text = source('06_CUSTOM_M_FUNCTIONS.pq')
    if 'Text.Replace(cleaned, ",", ".")' not in text:
        raise ValueError('Missing explicit EU separator support')
    return float('850,00'.replace(',', '.'))
