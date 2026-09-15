# html2pptx

**English** | [简体中文](README.zh-CN.md)

Convert a convention-based HTML slide deck into an editable PowerPoint (`.pptx`) presentation — no browser, no rendering engine, just Python.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)

## Features

- 🎞️ 16:9 slides (13.333 × 7.5 in), fully editable in PowerPoint / WPS
- 🧩 Automatic layouts: cards, stat blocks, comparison panels, flow diagrams, numbered plans, pull quotes, highlight panels, cover & closing pages
- ✍️ Preserves `<b>/<strong>/<em>` bold and `<br>` line breaks
- 📝 Auto page numbers, footers, and per-slide speaker notes
- 😀 Emoji supported
- 🌏 UTF-8 / GB18030 auto-detect

## Requirements

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io/) — `pip install python-pptx`
- Microsoft YaHei font for best CJK rendering (Windows default; change the `FONT` constant in the script if needed)

## Usage

```bash
python html2ppt.py your-deck.html
python html2ppt.py your-deck.html -o my-deck.pptx
```

The output defaults to the input filename with a `.pptx` extension.

## How it works

The converter is **convention-based**, not a general HTML renderer: it parses the document tree and maps each element to a layout by its `class` attribute, then draws it with a built-in light/dark design system.

| HTML | Rendered as |
|---|---|
| `<section class="slide">` | one slide |
| `class="cover"` / `class="final"` | dark cover / closing pages |
| `.kicker` | small cyan eyebrow label |
| `h1` / `h2` | slide title (`<br>` splits lines) |
| `.lead` | intro paragraph |
| `.sub` | gray footnote |
| `.badge` / `.foot` | cover badge / cover footer |
| `.quote` + `.who` | pull quote with attribution |
| `.mirror` | full-width highlight panel |
| `.cards` + `.card` (`.ic`, `h3`, `p`) | card grid (add `c2` / `c3` for 2 / 3 columns) |
| `.bullets` + `.b` (+ `<small>`) | bullet items; `<small>` becomes a gray sub-line |
| `.stat` + `.box` (`.num`, `.lbl`) | big-number stat boxes (`.num up` / `.num down` → green / pink) |
| `.vs` + `.side` | two-column comparison (`side a` ✔ cyan / `side b` ✖ pink) |
| `.chain` + `.link` (`.tag`, `.tt`) | flow diagram rows with ▼ arrows |
| `.num3` + `.row` (`.n`, `h4`, `p`) | numbered plan rows |

### Minimal example

```html
<section class="slide">
  <div class="kicker">01 · Introduction</div>
  <h2>Why it matters</h2>
  <p class="lead">A short lead paragraph.</p>

  <div class="cards c2">
    <div class="card">
      <div class="ic">💡</div>
      <h3>Idea</h3>
      <p>Card description goes here.</p>
    </div>
    <div class="card">
      <div class="ic">🚀</div>
      <h3>Action</h3>
      <p>Another card description.</p>
    </div>
  </div>

  <p class="sub">Footnote text.</p>
</section>
```

## Output

- 16:9 (13.333 × 7.5 in)
- Microsoft YaHei typeface
- Footer with deck name plus an `NN / TT` page number on every content slide
- An auto-generated speaker note attached to each slide

## Limitations

- Not a general HTML → PPTX renderer: colors, fonts and layout come from the built-in template, not your CSS.
- Images, charts, tables and animations are not supported.
- Unrecognized elements are skipped with a `[warn]` message on stderr.
- Text extraction is heuristic — heavily nested markup may need small manual adjustments.

## License

MIT