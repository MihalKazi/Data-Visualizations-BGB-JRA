"""Build embeds/network-full.html: every account that posted a copied comment under either post.

Usage: python3 build_full.py ["BGB Chief .xlsx"]
Standard library only. The layout is computed here, so the embed only has to draw it.
"""
import collections
import datetime as dt
import json
import math
import os
import random
import re
import sys

from build_network import HERE, SHEETS as TOP3_SHEETS, dhaka, fmt, read_xlsx, shorten

# sheet name -> post code. Columns are looked up by header name.
SHEETS = {'BGB post': 'BGB', 'JRA post': 'JRA'}
MIN_USES = 5        # a text counts as copied once it appears in this many comments
MIN_LENGTH = 40     # shorter phrases ("Good job") repeat by chance, so they are left out
LABEL_FROM = 13     # accounts with at least this many copied comments get their name printed
TRANSLATIONS = os.path.join(HERE, 'CiB Network BGB.xlsx')
# the report's narratives, matched by keyword in this order; a text that matches none is general praise for BGB.
# Keyword sorting is rough: check the result (python3 build_full.py --narratives) before publishing.
NARRATIVES = [
    ('story',  ['পিলখানা', 'সেলাই', 'রিকশা', '২২৫৮', '২,২৫৮', 'ড্রোন', 'মিয়ানমার', '২৭১', 'স্কোয়াড', 'পদ সৃষ্টি', 'আহত', 'মাথায়', 'গল্প', 'অভিজ্ঞতা']),
    ('july',   ['জুলাই', 'অভ্যুত্থান', 'আন্দোলন', 'সংযম']),
    ('attack', ['উদ্দেশ্য', 'প্রণোদিত', 'প্রনোদিত', 'ভিত্তিহীন', 'বানোয়াট', 'অপপ্রচার', 'মিথ্যা', 'গুজব', 'নিন্দা', 'ষড়যন্ত্র', 'অসত্য',
                'বিভ্রান্ত', 'নেত্র', 'হেয়', 'কালিমা', 'ক্ষুণ্ন', 'অপচেষ্টা', 'প্রোপাগান্ডা']),
    ('dg',     ['মহাপরিচালক', 'মহা পরিচালক', 'ডিজি', 'স্যার', 'সর্বশ্রেষ্ঠ', 'নেতৃত্ব']),
    ('july',   ['মানবিক', 'মানবতা', 'মানব সেবা', 'মানবসেবা', 'বিপদ', 'দুঃসময়', 'দুঃসময়', 'সংকট', 'বিপন্ন']),
]


def narrative(text):
    return next((name for name, words in NARRATIVES if any(w in text for w in words)), 'praise')


def norm(text):
    return re.sub(r'\s+', ' ', text or '').strip()


def load_comments(xlsx):
    sheets = read_xlsx(xlsx)
    comments = []
    for sheet, post in SHEETS.items():
        rows = sheets[sheet]
        col = {name: rows[0].index(name) for name in ('profileId', 'profileName', 'date', 'text', 'profileUrl')}
        for row in rows[1:]:
            row = row + [None] * (len(rows[0]) - len(row))
            if not row[col['date']]:
                continue
            url = (row[col['profileUrl']] or '').strip()
            pid = str(row[col['profileId']])
            # Facebook's pfbid ids change from post to post, so they only identify an account within one post
            key = url or ((post, pid) if pid.startswith('pfbid') else pid)
            comments.append({'post': post, 'key': key, 'url': url, 'name': (row[col['profileName']] or '').strip(),
                             'when': dhaka(row[col['date']]), 'text': norm(row[col['text']])})
    return comments


def load_translations():
    """English versions of the comment texts that the first workbook already translated."""
    if not os.path.exists(TRANSLATIONS):
        return {}
    sheets, out = read_xlsx(TRANSLATIONS), {}
    for sheet, (_, _, _, c_bn, c_en, _, _) in TOP3_SHEETS.items():
        for row in sheets.get(sheet, [])[1:]:
            if len(row) > c_en and row[c_bn] and row[c_en]:
                out.setdefault(norm(row[c_bn]), norm(row[c_en]))
    return out


def load_post_times():
    """When each post went up (Dhaka time), from the first workbook's Timeline sheet."""
    out = {}
    if os.path.exists(TRANSLATIONS):
        for row in read_xlsx(TRANSLATIONS).get('Timeline', [])[1:]:
            post = next((p for p in SHEETS.values() if row and row[0] and p in row[0]), None)
            if post and post not in out:
                out[post] = dt.datetime.strptime(row[2], '%I:%M %p, %d %B %Y')
    return out


def separate(items, gap, passes):
    """Push overlapping circles apart. items: dicts with x, y, r, group. gap(a, b) -> extra spacing."""
    cell = 2 * max(i['r'] for i in items) + 10
    for _ in range(passes):
        grid = collections.defaultdict(list)
        for i in items:
            grid[(int(i['x'] // cell), int(i['y'] // cell))].append(i)
        for (gx, gy), bucket in list(grid.items()):
            for dx, dy in ((0, 0), (1, 0), (0, 1), (1, 1), (-1, 1)):
                other = grid.get((gx + dx, gy + dy))
                if not other:
                    continue
                for ai, a in enumerate(bucket):
                    for b in (bucket[ai + 1:] if (dx, dy) == (0, 0) else other):
                        ox, oy = b['x'] - a['x'], b['y'] - a['y']
                        dist = math.hypot(ox, oy) or 0.01
                        need = a['r'] + b['r'] + gap(a, b)
                        if dist < need:
                            k = (need - dist) / dist / 2
                            a['x'] -= ox * k; a['y'] -= oy * k
                            b['x'] += ox * k; b['y'] += oy * k


def layout(accounts, texts):
    """Clusters first (one disc per text, BGB texts left, JRA texts right), then the accounts inside them."""
    rnd = random.Random(7)
    for t in texts:
        home = [a for a in accounts if a['home'] == t['i']]
        t['r'] = math.sqrt(sum((a['r'] + 1.2) ** 2 for a in home) / 0.86) + 1 if home else 0
        t['gx'] = (t['jra_share'] - 0.5) * 400
        t['x'], t['y'] = t['gx'] + rnd.uniform(-70, 70), rnd.uniform(-160, 160)
    discs = [t for t in texts if t['r']]
    shared = collections.Counter()
    for a in accounts:
        ids = sorted(l['t'] for l in a['links'])
        for i, x in enumerate(ids):
            for y in ids[i + 1:]:
                shared[(x, y)] += 1
    for step in range(900):
        if step < 840:
            for t in discs:
                t['x'] += (t['gx'] - t['x']) * 0.06
                t['y'] += (0 - t['y']) * 0.06
            for (x, y), n in shared.items():   # texts posted by the same accounts drift together
                a, b = texts[x], texts[y]
                k = 0.003 * min(n, 5)
                dx, dy = b['x'] - a['x'], b['y'] - a['y']
                a['x'] += dx * k; a['y'] += dy * k; b['x'] -= dx * k; b['y'] -= dy * k
        for i, a in enumerate(discs):
            for b in discs[i + 1:]:
                dx, dy = b['x'] - a['x'], b['y'] - a['y']
                dist = math.hypot(dx, dy) or 0.01
                need = a['r'] + b['r'] + 7
                if dist < need:
                    k = (need - dist) / dist / 2
                    a['x'] -= dx * k; a['y'] -= dy * k; b['x'] += dx * k; b['y'] += dy * k

    for a in accounts:
        t = texts[a['home']]
        ang, d = rnd.uniform(0, math.tau), rnd.uniform(0, t['r'])
        a['x'], a['y'], a['group'] = t['x'] + math.cos(ang) * d, t['y'] + math.sin(ang) * d, a['home']
    gap = lambda a, b: 1.3 if a['group'] == b['group'] else 7
    for step in range(320):
        if step < 280:
            for a in accounts:
                t = texts[a['home']]
                a['x'] += (t['x'] - a['x']) * 0.08
                a['y'] += (t['y'] - a['y']) * 0.08
        separate(accounts, gap, 2)

    pad = 24
    minx = min(a['x'] - a['r'] for a in accounts) - pad
    miny = min(a['y'] - a['r'] for a in accounts) - pad
    for a in accounts:
        a['x'], a['y'] = round(a['x'] - minx, 1), round(a['y'] - miny, 1)
        del a['group']
    for t in texts:
        del t['r'], t['gx'], t['x'], t['y']
        home = [a for a in accounts if a['home'] == t['i']]
        if home:   # where the cluster ended up, for the callouts
            t['cx'] = round(sum(a['x'] for a in home) / len(home), 1)
            t['cy'] = round(sum(a['y'] for a in home) / len(home), 1)
            t['cr'] = round(max(math.hypot(a['x'] - t['cx'], a['y'] - t['cy']) + a['r'] for a in home), 1)
    return (round(max(a['x'] + a['r'] for a in accounts) + pad), round(max(a['y'] + a['r'] for a in accounts) + pad))


def busiest_minute(times):
    """Most comments an account posted inside any 60 seconds."""
    times, best, start = sorted(times), 0, 0
    for end, t in enumerate(times):
        while (t - times[start]).total_seconds() > 60:
            start += 1
        best = max(best, end - start + 1)
    return best


def quote(text, limit=62):
    return ' “' + shorten(text['en'], limit) + '”' if text['en'] else ''


def annotate(accounts, texts):
    """The callouts printed on the graph. Each one points at the strongest example of a pattern."""
    notes = []
    at_text = lambda t: {'x': t['cx'], 'y': t['cy'], 'r': t['cr']}
    at_account = lambda a: {'x': a['x'], 'y': a['y'], 'r': a['r']}
    placed = [t for t in texts if 'cx' in t]

    bgb = max((t for t in placed if t['jra_share'] == 0), key=lambda t: t['profiles'], default=None)
    if bgb:
        notes.append({**at_text(bgb), 'text': bgb['i'], 'title': f"{bgb['profiles']} profiles, one identical comment",
                      'body': ('Each posted it once.' if bgb['comments'] == bgb['profiles'] else f"{bgb['comments']} comments in all.") + quote(bgb)})
    mixed = max((t for t in placed if 0 < t['jra_share'] < 1), key=lambda t: t['comments'], default=None)
    if mixed:
        shared = sum(1 for t in texts if 0 < t['jra_share'] < 1)
        notes.append({**at_text(mixed), 'text': mixed['i'], 'tour': True, 'title': f"{shared} of the texts turn up under both posts",
                      'body': f"This one was pasted by {mixed['on']['BGB']} accounts under the BGB post and {mixed['on']['JRA']} under the JRA post."})
    both = [a for a in accounts if len(a['posts']) > 1]
    if both:
        notes.append({**at_account(max(both, key=lambda a: (len(a['links']), a['total']))), 'color': 'both',
                      'tour': True, 'title': f"{len(both)} accounts appear under both posts",
                      'body': "The black circles pasted copied comments under the BGB post and again under the JRA post."})
    jra = max((t for t in placed if t['jra_share'] == 1), key=lambda t: t['comments'], default=None)
    if jra:
        notes.append({**at_text(jra), 'text': jra['i'], 'title': f"{jra['comments']} comments from {jra['profiles']} profiles",
                      'body': 'The same text, posted again and again.' + quote(jra)})
    fast = max(accounts, key=lambda a: a['minute'])
    if fast['minute'] >= 5:
        notes.append({**at_account(fast), 'account': accounts.index(fast), 'title': f"One account: {fast['minute']} comments in a minute",
                      'body': f"“{fast['name']}” posted {fast['total']} copied comments, using {len(fast['links'])} different texts."})
    return notes


def when(t, seconds=False):
    return f"{t.day} {t.strftime('%b')}, {t.hour % 12 or 12}:{t.strftime('%M:%S' if seconds else '%M')} {'AM' if t.hour < 12 else 'PM'}"


def burst(events):
    """An account's busiest 60 seconds: (comments, first, last, text used most, different texts used)."""
    best, start = (0, None, None, None, 0), 0
    for end, e in enumerate(events):
        while (e[0] - events[start][0]).total_seconds() > 60:
            start += 1
        if end - start + 1 > best[0]:
            window = events[start:end + 1]
            used = collections.Counter(w[2] for w in window)
            best = (len(window), window[0][0], window[-1][0], used.most_common(1)[0][0], len(used))
    return best


def tell(accounts, texts, posted, minutes):
    """The story the animation tells. The replay stops at each chapter and a panel opens beside the accounts involved.

    Every number, time and name comes from the data. The sentences around them are the editorial part: edit them here.
    A panel has: tone (which post), eyebrow, figure + unit, title, body, quote + cite, and the spot it points at (x, y, r).
    """
    events = sorted((e[0], e[1], e[2], i) for i, a in enumerate(accounts) for e in a['events'])
    of = lambda post: [e for e in events if e[1] == post]
    secs = lambda a, b: max(1, round((b - a).total_seconds()))
    stamp = lambda t: when(t).replace(', ', ' · ')
    name = lambda i: accounts[i]['name'].strip()

    def around(group):   # the circle that encloses a set of accounts
        x = sum(a['x'] for a in group) / len(group)
        y = sum(a['y'] for a in group) / len(group)
        return {'x': round(x, 1), 'y': round(y, 1), 'r': round(max(math.hypot(a['x'] - x, a['y'] - y) + a['r'] for a in group), 1)}
    middle = lambda post: {**around([a for a in accounts if a['posts'] == [post]]), 'r': 26}
    cluster = lambda t: {'x': t['cx'], 'y': t['cy'], 'r': t['cr']}

    chapters = []
    # 1. the first post
    if 'BGB' in posted:
        chapters.append({'at': minutes(posted['BGB']) + 50, 'cards': [{
            'when': minutes(posted['BGB']), 'r': 0, 'tone': 'BGB', 'eyebrow': stamp(posted['BGB']),   # points at that moment on the timeline
            'title': 'The official BGB page publishes a post',
            'body': 'Comments start to arrive. Many will turn out to be the same sentences, pasted by different accounts.'}]})
    # 2. its busiest hour
    hours = collections.defaultdict(list)
    for e in of('BGB'):
        hours[e[0].replace(minute=0, second=0)].append(e)
    top = max((t for t in texts if t['jra_share'] == 0 and 'cx' in t), key=lambda t: t['profiles'])
    if hours:
        hour, rows = max(hours.items(), key=lambda kv: len(kv[1]))
        who = len({e[3] for e in rows})
        span = f"{hour.day} {hour.strftime('%b')} · {hour.hour % 12 or 12} to {(hour.hour + 1) % 12 or 12} {'AM' if hour.hour + 1 < 12 else 'PM'}"
        # the panel points at one cluster, so it leads with what that cluster is; the hour's total goes in the body
        card = {**cluster(top), 'tone': 'BGB', 'eyebrow': span, 'figure': str(top['profiles']), 'unit': 'accounts paste this one sentence',
                'body': f"It is the busiest hour: {len(rows)} copied comments arrive, from {who} different accounts."}
        if top['en']:
            card.update(quote=shorten(top['en'], 90), cite='The sentence this cluster shares')
        chapters.append({'at': minutes(hour) + 60, 'cards': [card]})
    # 3-5. the second post: a long quiet, then accounts firing one after another
    bursts = sorted((b for b in ((burst([e for e in a['events'] if e[1] == 'JRA']), i) for i, a in enumerate(accounts)
                                 if any(e[1] == 'JRA' for e in a['events'])) if b[0][0] >= 8), key=lambda b: b[0][1])
    if bursts and 'JRA' in posted:
        (n, first, last, text, kinds), lead = bursts[0]
        before = sum(1 for e in of('JRA') if e[0] < first)
        chapters.append({'at': minutes(first) - 0.5, 'cards': [{
            'when': minutes(posted['JRA']), 'r': 0, 'tone': 'JRA', 'eyebrow': stamp(posted['JRA']),
            'title': 'A second post goes up, on the JRA page',
            'body': f"For {round((first - posted['JRA']).total_seconds() / 3600)} hours it draws only {before} copied comments. "
                    f"Then, at {when(first).split(', ')[1]}, that changes."}]})
        relay = [b for b in bursts[1:] if (b[0][1] - first).total_seconds() <= 1800][:3]
        group = [lead] + [i for _, i in relay]
        card = {**around([accounts[i] for i in group]), 'tone': 'JRA', 'eyebrow': stamp(first), 'figure': str(n),
                'unit': f"comments in {secs(first, last)} seconds, from one account",
                'body': f"“{name(lead)}” goes first." + (' Then ' + ', '.join(f"“{name(i)}” {b[0]} times" for b, i in relay) + '.' if relay else '')}
        if kinds == 1 and all(b[3] == text and b[4] == 1 for b, _ in relay) and texts[text]['en']:
            card.update(quote=shorten(texts[text]['en'], 90), cite='All of them paste this same line')
        chapters.append({'at': minutes(max([last] + [b[2] for b, _ in relay])) + 0.3, 'focus': group, 'cards': [card]})
        (n, first, last, _, _), fast = max(bursts, key=lambda b: b[0][0])
        if fast not in group:
            a = accounts[fast]
            chapters.append({'at': minutes(last) + 0.3, 'focus': [fast], 'cards': [{
                'x': a['x'], 'y': a['y'], 'r': a['r'], 'tone': 'JRA', 'eyebrow': stamp(first), 'figure': str(n),
                'unit': f"comments in {secs(first, last)} seconds",
                'body': f"“{name(fast)}” switches between {len(a['links'])} different texts, {a['total']} copied comments in all."}]})
    for c in chapters:
        if len(c.get('focus', [])) <= 6:
            for i in c.get('focus', []):
                accounts[i]['label'] = True   # the accounts a chapter names get their name printed
    chapters.sort(key=lambda c: c['at'])

    # 6. after the replay: the two methods side by side
    bgb = [a for a in accounts if a['posts'] == ['BGB']]
    heavy = [a for a in accounts if 'JRA' in a['posts'] and sum(1 for e in a['events'] if e[1] == 'JRA') >= 10]
    share = round(100 * sum(sum(1 for e in a['events'] if e[1] == 'JRA') for a in heavy) / max(1, len(of('JRA'))))
    once = sum(1 for a in bgb if a['total'] == 1)
    chapters.append({'cards': [
        {**middle('BGB'), 'over': True, 'tone': 'BGB', 'eyebrow': 'BGB post', 'figure': f"{once} of {len(bgb)}",
         'unit': 'accounts pasted a comment once', 'body': 'Many accounts, one comment each.'},
        {**middle('JRA'), 'over': True, 'tone': 'JRA', 'eyebrow': 'JRA post', 'figure': f"{share}%",
         'unit': f"of the copied comments came from {len(heavy)} accounts", 'body': 'Few accounts, many comments each.'}]})
    # 7. the lines
    chapters.append({'reveal': 'edges', 'cards': [{
        **cluster(top), 'tone': 'both', 'eyebrow': 'The network', 'figure': str(len(texts)), 'unit': 'texts posted word for word',
        'body': 'Each cluster, like this one, is one text. A line joins two accounts that posted the same one.'}]})
    # 8-9. what ties the two posts together
    both = [a for a in accounts if len(a['posts']) > 1]
    if both:
        a = max(both, key=lambda a: (len(a['links']), a['total']))
        chapters.append({'spot': {'color': 'both'}, 'cards': [{
            'x': a['x'], 'y': a['y'], 'r': a['r'], 'marks': [[b['x'], b['y'], b['r']] for b in both],   # every one of them gets a ring
            'tone': 'both', 'eyebrow': 'Both posts', 'figure': str(len(both)),
            'unit': 'accounts appear under both posts',
            'body': 'The black circles pasted copied comments under the BGB post and again under the JRA post.'}]})
    mixed = [t for t in texts if 0 < t['jra_share'] < 1 and 'cx' in t]
    if mixed:
        t = max(mixed, key=lambda t: t['comments'])
        chapters.append({'spot': {'text': t['i']}, 'cards': [{
            **cluster(t), 'tone': 'both', 'eyebrow': 'Shared text', 'figure': str(len(mixed)), 'unit': 'texts turn up under both posts',
            'body': f"This one was pasted by {t['on']['BGB']} accounts under the BGB post and {t['on']['JRA']} under the JRA post."}]})
    return {
        'headline': "Tracing the Coordinated Network Hour by Hour",
        'dek': f"Under the BGB post, {once} of {len(bgb)} accounts pasted a comment once. "
               f"Under the JRA post, {len(heavy)} accounts wrote {share}% of the copied comments.",
        'chapters': chapters,
    }


def build(xlsx):
    comments = load_comments(xlsx)
    english = load_translations()
    uses = collections.Counter(c['text'] for c in comments)
    copied = [t for t, n in uses.most_common() if n >= MIN_USES and len(t) >= MIN_LENGTH]
    index = {t: i for i, t in enumerate(copied)}

    by_account = collections.OrderedDict()
    for c in comments:
        if c['text'] not in index:
            continue
        a = by_account.setdefault(c['key'], {'name': c['name'], 'url': c['url'], 'links': collections.defaultdict(list), 'posts': set(), 'events': []})
        a['links'][index[c['text']]].append(c['when'])
        a['events'].append((c['when'], c['post'], index[c['text']]))
        a['posts'].add(c['post'])

    start = min(e[0] for a in by_account.values() for e in a['events']).replace(minute=0, second=0)
    minutes = lambda when: round((when - start).total_seconds() / 60, 2)
    accounts = []
    for a in by_account.values():
        links = [{'t': t, 'n': len(ts), 'first': fmt(min(ts)), 'last': fmt(max(ts))}
                 for t, ts in sorted(a['links'].items(), key=lambda kv: -len(kv[1]))]
        total = sum(l['n'] for l in links)
        accounts.append({'name': a['name'], 'url': a['url'], 'links': links, 'total': total,
                         'posts': sorted(a['posts']), 'home': links[0]['t'], 'r': round(2.6 + 1.8 * math.sqrt(total), 2),
                         'label': total >= LABEL_FROM,
                         # every copied comment as [minutes since the first hour, 1 if on the JRA post], for the replay
                         'ev': [[minutes(w), int(p == 'JRA')] for w, p, _ in sorted(a['events'])],
                         'events': sorted(a['events']),
                         'minute': busiest_minute([t for ts in a['links'].values() for t in ts])})

    texts = []
    for t, i in index.items():
        rows = [c for c in comments if c['text'] == t]
        texts.append({'i': i, 'bn': t, 'en': english.get(t, ''), 'comments': len(rows), 'nar': narrative(t),
                      'profiles': len({c['key'] for c in rows}),
                      'on': {p: len({c['key'] for c in rows if c['post'] == p}) for p in SHEETS.values()},
                      'jra_share': sum(c['post'] == 'JRA' for c in rows) / len(rows)})
    # copied comments that landed in the same second: [minute, [account, ...]], one entry per comment
    stamps = collections.defaultdict(list)
    for i, a in enumerate(accounts):
        for w, _, _ in a['events']:
            stamps[w.replace(microsecond=0)].append(i)
    sync = [[minutes(w), who] for w, who in sorted(stamps.items()) if len(who) > 1]

    width, height = layout(accounts, texts)
    notes = annotate(accounts, texts)
    story = tell(accounts, texts, load_post_times(), minutes)
    for t in texts:
        t['jra_share'] = round(t['jra_share'], 2)
    for a in accounts:
        del a['minute'], a['events']

    posts = collections.Counter(p for a in accounts for p in (['both'] if len(a['posts']) > 1 else a['posts']))
    return {
        'w': width, 'h': height, 'texts': texts, 'accounts': accounts, 'notes': notes, 'story': story, 'sync': sync,
        't0': int(start.replace(tzinfo=dt.timezone.utc).timestamp() * 1000),   # Dhaka wall-clock, read back with UTC getters
        'posted': {p: minutes(w) for p, w in load_post_times().items()},
        'counts': {'BGB': posts['BGB'], 'JRA': posts['JRA'], 'both': posts['both'],
                   'bgb_once': sum(1 for a in accounts if a['posts'] == ['BGB'] and a['total'] == 1),
                   'jra_heavy': sum(1 for a in accounts if a['posts'] == ['JRA'] and a['total'] >= 10),
                   'comments': sum(a['total'] for a in accounts), 'all_comments': len(comments)},
    }


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    xlsx = args[0] if args else os.path.join(HERE, 'BGB Chief .xlsx')
    data = build(xlsx)
    if '--narratives' in sys.argv:   # list every text under its narrative, to check the keyword sorting
        for name in ('attack', 'dg', 'july', 'story', 'praise'):
            print(f"\n== {name}")
            for t in data['texts']:
                if t['nar'] == name:
                    print(f"  {t['comments']:>3}  {t['bn'][:100]}")
        sys.exit()
    template = open(os.path.join(HERE, 'src', 'full.template.html'), encoding='utf-8').read()
    os.makedirs(os.path.join(HERE, 'embeds'), exist_ok=True)
    # the same graph twice, differing only in how the lines between clusters are drawn, so the two can be compared
    for name, lines in (('network-full.html', 'single'), ('index.html', 'all')):
        data['lines'] = lines
        payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
        out = os.path.join(HERE, 'embeds', name)
        open(out, 'w', encoding='utf-8').write(template.replace('/*DATA*/', payload))
        print(out)
    c = data['counts']
    print(f"{len(data['accounts'])} accounts, {len(data['texts'])} copied texts, "
          f"{c['comments']} of {c['all_comments']} comments, canvas {data['w']}x{data['h']}, "
          f"{len(data['sync'])} same-second moments ({sum(len(s[1]) for s in data['sync'])} comments)")
