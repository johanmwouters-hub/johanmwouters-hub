# cv-layout-check

Small Python tools to tailor a two-column Word CV and check its layout before anyone sees it.

I tailor my CV for every application. Each version changes a few lines: a different objective, a reordered skills list, one more bullet under a job. In a two-column Word layout, one extra line is enough to push part of the first page onto page 2, leave page 2 almost empty, or wrap a sidebar entry onto a second line. Word does not warn you. You only notice when you open the PDF, or worse, when the recruiter does.

These tools make that check automatic.

## What it does

- **Edits the CV without Word.** `cvlib.py` changes the document XML directly (lxml): replace text in a paragraph while keeping its formatting, grow or shrink a list in place, move blocks between cells, add a spacer that controls where page 2 starts.
- **Refuses entries that would wrap.** `assert_one_line` measures each sidebar entry with the real font metrics (fontTools) and fails before rendering if an item is wider than the column.
- **Renders and reads back.** `check_page1` and `sections_intact` render the document with LibreOffice, read the PDF with poppler and verify that page 1 ends where it should and that no sidebar section is split across pages.
- **Compares font families.** `fontswap.py` swaps a font family throughout a .docx and renders each candidate, so you can pick the most compact one that still reads well.

## Quick start

Requirements: Python 3.10+, LibreOffice (`soffice`), poppler (`pdfinfo`, `pdftotext`) and the Lato font.

```
pip install -r requirements.txt
python demo.py
```

`demo.py` builds a synthetic CV for "Jane Example" (no personal data), grows the skills list, saves a tailored copy and checks both:

```
all sidebar entries fit on one line
sample.docx: pages=1 split sections: none
tailored.docx: pages=1 split sections: none
```

Compare fonts on your own CV:

```
python fontswap.py my-cv.docx "Calibri" "Lato" "Figtree" "Open Sans"
```

## Lessons learned

- **Check the font is really installed.** If the CV font is missing, LibreOffice silently falls back to another font with different widths and every page-break check is wrong. `find_font` looks in the usual font folders and fails loudly instead.
- **A spacer paragraph beats page-break settings.** "Keep with next" and "page break before" behave differently in Word and LibreOffice when a table cell spans rows. A spacer of a known height gives the same result in both.
- **Measure before you render.** Rendering takes seconds; measuring text width takes microseconds. Catching a wrapping skill before rendering keeps the edit loop fast.
- **Balance the columns.** When the right column is shorter than the sidebar on page 1, the sidebar overflow creates a nearly empty page 2. Checking where each column ends catches this.
- **Unescape what you read back.** `pdftotext` returns `&amp;` for `&`, which made "Pricing & packaging" look like a split section until the text was unescaped.

## Built with Claude

I built these tools together with Claude (Anthropic) while running my own job search, as part of an agent workflow that finds roles, tailors CVs and tracks applications. Claude wrote most of the code; I defined the problem, set the layout rules and reviewed every CV it produced.

## License

MIT, see [LICENSE](LICENSE).
