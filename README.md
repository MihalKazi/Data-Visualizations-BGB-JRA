# BGB / JRA comment network

An embeddable, animated network chart of copy-pasted Facebook comments under two posts about the BGB chief: one official BGB post and one JRA post.

Each circle is a profile. Profiles that posted the same copied text sit in one cluster. Lines join profiles that posted the same text. The chart replays the comments in time order, then draws the lines, then walks through the patterns one at a time.

## What it shows

- Under the BGB post, almost every account commented once.
- Under the JRA post, a few accounts wrote many of the copied comments.
- Some accounts appear under both posts.

## Files

| File | Purpose |
| --- | --- |
| `build_full.py` | Reads the workbook, finds copied texts, lays out the clusters and writes the embeds |
| `build_network.py` | Workbook reader used by `build_full.py`, and the first top-3 chart |
| `src/full.template.html` | The embed: layout, animation, story panels, hover box |
| `src/network.template.html` | Template for the first top-3 chart |

## Build

Python 3 standard library only, no packages to install.

```
python3 build_full.py
```

The script expects `BGB Chief .xlsx` and `CiB Network BGB.xlsx` next to it. Neither file is in this repo. The output goes to `embeds/`:

- `index.html`: every account's lines to every cluster it posted in
- `network-full.html`: one line per account, to its main cluster

## Embed

Each output is a single self-contained HTML file.

```html
<iframe src="index.html" width="100%" height="1500"
        style="border:0" loading="lazy" title="BGB comment network"></iframe>
```

The embed posts its height to the parent page, so the host can resize the iframe:

```js
addEventListener('message', e => {
  if (e.data && e.data.type === 'embed-height') iframe.style.height = e.data.height + 'px';
});
```

Query options:

| Option | Effect |
| --- | --- |
| `?theme=dark` or `?theme=auto` | Dark theme, or follow the system |
| `?static=1` | Finished picture, no animation |
| `?lines=all` or `?lines=single` | Lines per account |

The layout switches to a vertical one below 620px width. On phones, tapping a circle opens a solid box; on wider screens hovering opens a see-through box so lines can be followed across it.

## Data and privacy

The generated pages in `embeds/` are the verified report and are included. The source workbooks (`*.xlsx`) carry real people's names and profile links and are kept out of this repo (see `.gitignore`).

How a copied text is defined: used by at least 5 comments and at least 40 characters long. Accounts are identified by profile link, not by display name.
