"""Build bars/index.html: comments per hour under both posts, on one time axis.

Usage: python3 build_bars.py ["BGB Chief .xlsx"]
Standard library only. Each hour's bar is split into copy-paste comments (a text that appears more than once
under the same post, the report's definition) and the rest; a gold mark under it counts comments that landed
in the same second as another comment.
"""
import collections
import datetime as dt
import json
import os
import sys

from build_full import HERE, SHEETS, load_comments, load_post_times


def build(xlsx):
    comments = load_comments(xlsx)
    uses = {p: collections.Counter(c['text'] for c in comments if c['post'] == p) for p in SHEETS.values()}
    same_second = collections.Counter(c['when'].replace(microsecond=0) for c in comments)

    start = min(c['when'] for c in comments).replace(minute=0, second=0, microsecond=0)
    end = max(c['when'] for c in comments).replace(minute=0, second=0, microsecond=0)
    hours = int((end - start).total_seconds() // 3600) + 1
    # per hour: [BGB total, BGB copied, JRA total, JRA copied, comments in a shared second]
    rows = [[0, 0, 0, 0, 0] for _ in range(hours)]
    for c in comments:
        h = int((c['when'] - start).total_seconds() // 3600)
        k = 0 if c['post'] == 'BGB' else 2
        rows[h][k] += 1
        if uses[c['post']][c['text']] > 1:
            rows[h][k + 1] += 1
        if same_second[c['when'].replace(microsecond=0)] > 1:
            rows[h][4] += 1

    posted = load_post_times()
    totals = {p: sum(1 for c in comments if c['post'] == p) for p in SHEETS.values()}
    copied = {p: sum(n for n in uses[p].values() if n > 1) for p in SHEETS.values()}
    return {
        'headline': 'Two waves of comments, one night apart',
        'dek': 'Comments per hour under the official BGB post and the JRA post',
        't0': int(start.replace(tzinfo=dt.timezone.utc).timestamp() * 1000),   # Dhaka wall-clock, read back with UTC getters
        'hours': rows,
        'posted': {p: round((w - start).total_seconds() / 3600, 3) for p, w in posted.items()},
        'totals': totals, 'copied': copied,
        'sync': sum(n for n in same_second.values() if n > 1),
        'source': 'Data: public comments under the official BGB post and the JRA post, Dhaka time. '
                  'Copy-paste: a text that appears more than once under the same post.',
    }


if __name__ == '__main__':
    xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'BGB Chief .xlsx')
    data = build(xlsx)
    template = open(os.path.join(HERE, 'src', 'bars.template.html'), encoding='utf-8').read()
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    # its own folder, so it can be deployed as a separate Vercel project (Root Directory: bars)
    os.makedirs(os.path.join(HERE, 'bars'), exist_ok=True)
    out = os.path.join(HERE, 'bars', 'index.html')
    open(out, 'w', encoding='utf-8').write(template.replace('/*DATA*/', payload))
    peak = {p: max(r[k] for r in data['hours']) for p, k in (('BGB', 0), ('JRA', 2))}
    print(f"{out}: {len(data['hours'])} hours, totals {data['totals']}, copied {data['copied']}, "
          f"peaks {peak}, same-second comments {data['sync']}")
