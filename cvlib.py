# -*- coding: utf-8 -*-
"""Tailor a two-column Word CV and check its layout before anyone sees it.

Built for a CV laid out as one Word table: a narrow left sidebar (objective,
skills, languages, education, awards) and a wide right column (profile and
experience). Typical structure:

  rows 0-1  header: photo, name, headline
  rows 2-5  block A: left cell vertically merged over rows 2-5 (sidebar top),
            right cells hold the profile and the first job
  row 6     block B: left cell = rest of the sidebar, right cell = rest of
            the experience

The recurring defect: block A no longer fits on page 1, Word or LibreOffice
pushes part of it to page 2, and page 2 ends up almost empty. The functions
below edit the document XML directly (lxml, no Word needed) and verify the
result by rendering it with LibreOffice and reading the PDF back.
"""
import subprocess, zipfile, os, re, copy, html
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def q(n):
    return W + n


def load(src):
    return etree.fromstring(zipfile.ZipFile(src).read('word/document.xml'))


def paras(root):
    return list(root.iter(q('p')))


def text_of(p):
    return ''.join(t.text or '' for t in p.iter(q('t')))


def cell(root, row, col):
    tbl = root.find(q('body') + '/' + q('tbl'))
    return tbl.findall(q('tr'))[row].findall(q('tc'))[col]


def set_run_text(run, text):
    ts = run.findall('.//' + q('t'))
    ts[0].text = text
    ts[0].set(XML_SPACE, 'preserve')
    for t in ts[1:]:
        t.getparent().remove(t)


def apply_edits(root, edits, check):
    """Replace whole-paragraph text by global paragraph index."""
    ps = paras(root)
    for idx, new in sorted(edits.items()):
        p = ps[idx]
        runs = [r for r in p.findall(q('r')) if r.findall('.//' + q('t'))]
        old = text_of(p)
        assert old.startswith(check[idx]), (idx, old[:60])
        set_run_text(runs[0], new)
        for r in runs[1:]:
            p.remove(r)


def grab(root, idx_list):
    """Capture paragraph elements by index before any structural change.

    Indices shift once paragraphs are added or removed, so take references
    first and mutate afterwards.
    """
    ps = paras(root)
    return [ps[i] for i in sorted(idx_list)]


def set_list(root, idx_list, items):
    """Rewrite a sidebar list to an arbitrary length, addressing it by index."""
    return fill_list(grab(root, idx_list), items)


def fill_list(block, items):
    """Rewrite a captured sidebar list to an arbitrary length.

    Extra paragraphs are cloned from the last existing one and surplus ones
    removed, so SKILLS can grow or shrink per role. Takes elements rather than
    indices, because indices shift the moment anything is added or moved.
    """
    parent = block[0].getparent()
    for p, txt in zip(block, items):
        runs = [r for r in p.findall(q('r')) if r.findall('.//' + q('t'))]
        set_run_text(runs[0], txt)
        for r in runs[1:]:
            p.remove(r)
    made = []
    if len(items) > len(block):
        anchor = block[-1]
        for txt in items[len(block):]:
            new = copy.deepcopy(block[-1])
            runs = [r for r in new.findall(q('r')) if r.findall('.//' + q('t'))]
            set_run_text(runs[0], txt)
            for r in runs[1:]:
                new.remove(r)
            anchor.addnext(new)
            anchor = new
            made.append(new)
        return block + made
    for p in block[len(items):]:
        parent.remove(p)
    return block[:len(items)]


# CT_PPr requires its children in this order; Word ignores or rejects them otherwise.
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr']


def set_ppr_flag(p, tag):
    """Add a boolean paragraph property, respecting the schema's element order."""
    pPr = p.find(q('pPr'))
    if pPr is None:
        pPr = etree.Element(q('pPr'))
        p.insert(0, pPr)
    if pPr.find(q(tag)) is not None:
        return
    rank = PPR_ORDER.index(tag)
    el = etree.Element(q(tag))
    for child in pPr:
        name = etree.QName(child).localname
        if name not in PPR_ORDER or PPR_ORDER.index(name) > rank:
            child.addprevious(el)
            return
    pPr.append(el)


def keep_together(elems, break_before=False):
    """Stop a sidebar block from splitting across a page boundary.

    keepLines on every paragraph and keepNext on all but the last. Word honors
    these inside a table cell; LibreOffice does not, so break_before also puts a
    hard page break in front of the block, which both engines obey.
    """
    for i, p in enumerate(elems):
        set_ppr_flag(p, 'keepLines')
        if i < len(elems) - 1:
            set_ppr_flag(p, 'keepNext')
    if break_before:
        set_ppr_flag(elems[0], 'pageBreakBefore')


def set_spacing(elems, line=None, after=None, before=None):
    """Set paragraph spacing.

    The sidebar shipped with line=360 (1.5) and no space after, so a wrapped
    entry looked like two entries. Single line spacing inside a paragraph plus
    space after it separates entries instead of separating their lines.
    """
    for p in elems:
        pPr = p.find(q('pPr'))
        if pPr is None:
            pPr = etree.Element(q('pPr'))
            p.insert(0, pPr)
        sp = pPr.find(q('spacing'))
        if sp is None:
            sp = etree.Element(q('spacing'))
            ref = pPr.find(q('ind'))
            (ref.addprevious(sp) if ref is not None else pPr.append(sp))
        if line is not None:
            sp.set(q('line'), str(line)); sp.set(q('lineRule'), 'auto')
        if after is not None:
            sp.set(q('after'), str(after))
        if before is not None:
            sp.set(q('before'), str(before))


def clone_block(heading_tpl, item_tpl, heading, items, anchor, where='before'):
    """Build a new sidebar section from an existing heading and item paragraph."""
    made = []
    for tpl, txt in [(heading_tpl, heading)] + [(item_tpl, t) for t in items]:
        p = copy.deepcopy(tpl)
        runs = [r for r in p.findall(q('r')) if r.findall('.//' + q('t'))]
        set_run_text(runs[0], txt)
        for r in runs[1:]:
            p.remove(r)
        made.append(p)
    ref = anchor
    for p in made:
        if where == 'before':
            ref.addprevious(p)
        else:
            ref.addnext(p); ref = p
    return made


def spacer(after_elem, twips):
    """An empty paragraph of exact height, placed after a block.

    LibreOffice ignores page breaks and keepNext inside a table row that splits
    across pages, and a large space-after saturates instead of pushing. An empty
    paragraph with an exact line height is the one lever that moves content to
    the next page by a controllable amount.
    """
    p = etree.Element(q('p'))
    pPr = etree.SubElement(p, q('pPr'))
    sp = etree.SubElement(pPr, q('spacing'))
    sp.set(q('before'), '0'); sp.set(q('after'), '0')
    sp.set(q('line'), str(twips)); sp.set(q('lineRule'), 'exact')
    after_elem.addnext(p)
    return p


def move_before(elems, anchor):
    """Reorder sidebar blocks within a cell.

    LibreOffice ignores keepNext, keepLines and page breaks inside a table row
    that splits across pages, so reordering is the only page-break control that
    both engines respect.
    """
    for el in elems:
        el.getparent().remove(el)
    for el in elems:
        anchor.addprevious(el)


def reorder_blocks(blocks, root, row, col):
    """Lay the sidebar blocks out in the given order inside one cell."""
    flat = [p for b in blocks for p in b]
    for p in flat:
        p.getparent().remove(p)
    dest = cell(root, row, col)
    anchor = dest.find(q('tcPr'))
    pos = list(dest).index(anchor) + 1 if anchor is not None else 0
    for offset, p in enumerate(flat):
        dest.insert(pos + offset, p)


def move_elems(elems, root, row, col, where='start'):
    """Move captured paragraph elements into another table cell."""
    for p in elems:
        p.getparent().remove(p)
    dest = cell(root, row, col)
    if where == 'start':
        anchor = dest.find(q('tcPr'))
        pos = list(dest).index(anchor) + 1 if anchor is not None else 0
        for offset, p in enumerate(elems):
            dest.insert(pos + offset, p)
    else:
        for p in elems:
            dest.append(p)


def save(src, dst, root):
    zin = zipfile.ZipFile(src)
    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == 'word/document.xml'
                          else zin.read(item.filename))


# ------------------------------------------------------- sidebar line widths

# Usable text width of the sidebar. Example: a 2701-twip cell with 165 twips of
# cell margin and a -20 twip paragraph indent on each side leaves 2411 twips,
# which is 120.6pt. Set this to your own cell.
SIDEBAR_PT = 120.6
_metrics = {}


FONT_DIRS = ['/usr/share/fonts', '/usr/local/share/fonts', os.path.expanduser('~/.fonts'),
             os.path.expanduser('~/.local/share/fonts'), os.path.expanduser('~/Library/Fonts'),
             '/Library/Fonts', '/System/Library/Fonts']


def find_font(font):
    """Locate a .ttf file by its file stem, e.g. 'Lato-Regular'."""
    import glob
    for d in FONT_DIRS:
        hits = glob.glob(os.path.join(d, '**', font + '.ttf'), recursive=True)
        if hits:
            return hits[0]
    raise FileNotFoundError(f'{font}.ttf not found in {FONT_DIRS}')


def text_width(s, font='Lato-Regular', pt=11):
    """Rendered width in points, measured with the real font metrics."""
    from fontTools.ttLib import TTFont
    if font not in _metrics:
        path = find_font(font)
        f = TTFont(path)
        _metrics[font] = (f['head'].unitsPerEm, f.getBestCmap(), f['hmtx'])
    upm, cmap, hmtx = _metrics[font]
    return sum(hmtx[cmap[ord(c)]][0] for c in s if ord(c) in cmap) / upm * pt


def assert_one_line(items, limit=SIDEBAR_PT):
    """Refuse any sidebar entry that would wrap."""
    bad = [(t, round(text_width(t), 1)) for t in items if text_width(t) > limit]
    assert not bad, f'these sidebar entries wrap (limit {limit}pt): {bad}'


# ---------------------------------------------------------------- fit check

def page_texts(docx):
    """Render with LibreOffice and return the text of each page."""
    d = os.path.dirname(os.path.abspath(docx)) or '.'
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', d, docx],
                   capture_output=True, timeout=180)
    pdf = os.path.join(d, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
    n = int(re.search(r'Pages:\s+(\d+)',
                      subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
    out = []
    for i in range(1, n + 1):
        out.append(subprocess.run(['pdftotext', '-layout', '-f', str(i), '-l', str(i), pdf, '-'],
                                  capture_output=True, text=True).stdout)
    return pdf, out


def norm(s):
    return re.sub(r'\s+', ' ', s)


def sidebar_text(pdf, x_limit=210.0):
    """Per-page text of the left sidebar only.

    Searching whole pages gives false hits, because the summary mentions the
    same things the sidebar lists (patents, papers, degrees).
    """
    xml = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True).stdout
    out = []
    for page in re.findall(r'<page .*?</page>', xml, re.S):
        words = re.findall(r'<word xMin="([\d.]+)"[^>]*>([^<]*)</word>', page)
        out.append(' '.join(html.unescape(w) for x, w in words if float(x) < x_limit))
    return out


def sections_intact(pdf, sections, x_limit=210.0):
    """sections: {heading: last words of that section}. Returns the split ones."""
    pages = sidebar_text(pdf, x_limit)
    bad = []
    for head, tail in sections.items():
        a = next((i for i, p in enumerate(pages) if head in p), None)
        b = next((i for i, p in enumerate(pages) if norm(tail) in p), None)
        if a is None or b is None or a != b:
            bad.append(f'{head} heading p{None if a is None else a + 1} tail p{None if b is None else b + 1}')
    return bad


def check_page1(docx, must_end_sidebar, must_end_firstjob):
    """BLOCK A has to be complete on page 1.

    must_end_sidebar   last few words of the last sidebar entry on page 1
    must_end_firstjob  last few words of the first job's paragraph
    """
    pdf, pages = page_texts(docx)
    p1 = norm(pages[0])
    ok_side = norm(must_end_sidebar) in p1
    ok_job = norm(must_end_firstjob) in p1
    print(f'pages={len(pages)}  sidebar_complete={ok_side}  first_job_complete={ok_job}')
    if not (ok_side and ok_job):
        print('--- page 2 starts with ---')
        print(norm(pages[1])[:400] if len(pages) > 1 else '(no page 2)')
    return pdf, ok_side and ok_job, pages
