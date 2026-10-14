"""Build embeds/network.html from the comment workbook.

Usage: python3 build_network.py ["CiB Network BGB.xlsx"]
Standard library only.
"""
import collections
import datetime as dt
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# sheet name -> (post code, name col, ISO time col, Bangla col, English col, rank col, profile URL col)
SHEETS = {
    'BGB Top 3 Comment': ('BGB', 0, 1, 2, 3, 5, None),
    'JRA Top 3 Comment': ('JRA', 0, 1, 2, 3, 6, 4),
}
# A profile URL identifies an account; a display name does not (two different accounts are both
# called "Aminul Islam"). Rows without a URL are merged by name only if this is switched on.
MERGE_SAME_NAME_WITHOUT_URL = False
RANKS = {'topfirst': 1, 'first': 1, 'second': 2, 'third': 3}


def col_index(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.parse(z.open('xl/sharedStrings.xml')).getroot().findall(M + 'si'):
            strings.append(''.join(t.text or '' for t in si.iter(M + 't')))
    rels = {r.get('Id'): r.get('Target') for r in ET.parse(z.open('xl/_rels/workbook.xml.rels')).getroot()}
    sheets = {}
    for sh in ET.parse(z.open('xl/workbook.xml')).getroot().find(M + 'sheets'):
        target = rels[sh.get(R + 'id')]
        part = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        rows = []
        for row in ET.parse(z.open(part)).getroot().iter(M + 'row'):
            cells = {}
            for c in row.findall(M + 'c'):
                v = c.find(M + 'v')
                if c.get('t') == 's' and v is not None:
                    val = strings[int(v.text)]
                elif c.get('t') == 'inlineStr':
                    val = ''.join(t.text or '' for t in c.iter(M + 't'))
                else:
                    val = v.text if v is not None else None
                cells[col_index(c.get('r'))] = val
            if cells:
                rows.append([cells.get(i) for i in range(max(cells) + 1)])
        sheets[sh.get('name')] = rows
    return sheets


def dhaka(iso):
    return dt.datetime.strptime(iso[:19], '%Y-%m-%dT%H:%M:%S') + dt.timedelta(hours=6)


def fmt(t):
    return f"{t.day} {t.strftime('%b')}, {t.hour % 12 or 12}:{t.strftime('%M:%S')} {'AM' if t.hour < 12 else 'PM'}"


def shorten(text, limit=30):
    """Opening words of a comment, for the caption under its cluster."""
    out = ''
    for word in text.split():
        if out and len(out) + len(word) + 1 > limit:
            break
        out = (out + ' ' + word).strip()
    return out.rstrip(',;') + '…'


def build(xlsx):
    sheets = read_xlsx(xlsx)
    texts = {}                                   # id -> text record
    by_profile = collections.OrderedDict()       # account key -> {'name', 'url', 'links': {text id -> [times]}}
    all_times = []
    for sheet, (post, c_name, c_time, c_bn, c_en, c_rank, c_url) in SHEETS.items():
        variants = collections.defaultdict(collections.Counter)
        for n, row in enumerate(sheets[sheet][1:]):
            row = row + [None] * 8
            if not row[c_name] or not row[c_time]:
                continue
            rank = RANKS[row[c_rank].strip().lower()]
            tid = post[0] + str(rank)
            when = dhaka(row[c_time])
            all_times.append(when)
            variants[tid][(row[c_bn].strip(), (row[c_en] or '').strip())] += 1
            name = row[c_name].strip()
            url = (row[c_url] or '').strip() if c_url is not None else ''
            key = url or (('name', name) if MERGE_SAME_NAME_WITHOUT_URL else (post, n))
            entry = by_profile.setdefault(key, {'name': name, 'url': url, 'links': collections.defaultdict(list)})
            entry['links'][tid].append(when)
            texts.setdefault(tid, {'id': tid, 'post': post, 'rank': rank, 'kind': 'hub'})
        for tid, counts in variants.items():
            (bn, en), top = counts.most_common(1)[0]
            total = sum(counts.values())
            texts[tid].update(bn=bn, en=en, short=shorten(en), comments=total, variants=total - top)

    profiles = []
    for entry in by_profile.values():
        name, links = entry['name'], entry['links']
        out = [{'t': tid, 'n': len(ts), 'first': fmt(min(ts)), 'last': fmt(max(ts))} for tid, ts in sorted(links.items())]
        profiles.append({
            'id': 'p' + str(len(profiles)), 'kind': 'profile', 'name': name, 'url': entry['url'], 'links': out,
            'total': sum(l['n'] for l in out),
            'posts': sorted({'BGB' if l['t'][0] == 'B' else 'JRA' for l in out}),
        })
    for t in texts.values():
        t['profiles'] = sum(1 for p in profiles if any(l['t'] == t['id'] for l in p['links']))

    first, last = min(all_times), max(all_times)
    return {
        'texts': sorted(texts.values(), key=lambda t: t['id']),
        'profiles': profiles,
        'source': (f"Data: Facebook comments carrying the three most-copied texts under each post, "
                   f"{first.day} to {last.day} {last.strftime('%B %Y')}, Dhaka time. "
                   "Accounts are identified by profile link where the data has one."),
    }


if __name__ == '__main__':
    xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'CiB Network BGB.xlsx')
    data = build(xlsx)
    template = open(os.path.join(HERE, 'src', 'network.template.html'), encoding='utf-8').read()
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    os.makedirs(os.path.join(HERE, 'embeds'), exist_ok=True)
    out = os.path.join(HERE, 'embeds', 'network.html')
    open(out, 'w', encoding='utf-8').write(template.replace('/*DATA*/', payload))
    print(f"{out}: {len(data['profiles'])} profiles, {len(data['texts'])} texts, "
          f"{sum(t['comments'] for t in data['texts'])} comments")
