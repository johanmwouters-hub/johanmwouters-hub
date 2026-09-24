# -*- coding: utf-8 -*-
"""Swap a font family throughout a .docx, then render it to compare layouts.

Usage:  python fontswap.py cv.docx "Old Family" "Lato" "Figtree" "Open Sans"
Writes cmp/<Family>.docx and .pdf plus a page-1 PNG for each candidate and
prints the page count, so you can pick the most compact family that still
reads well. The family name is usually only referenced in styles.xml,
fontTable.xml and the theme, but body and header parts are rewritten too.
"""
import zipfile, sys, os, subprocess
import cvlib as C

PARTS = ('word/styles.xml', 'word/fontTable.xml', 'word/theme/theme1.xml',
         'word/document.xml', 'word/header1.xml', 'word/footer1.xml')


def swap(src, dst, old, new, old_light=None, new_light=None):
    old_light = old_light or old + ' Light'
    new_light = new_light or new + ' Light'
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in PARTS:
                s = data.decode('utf8').replace(old_light, new_light).replace(old, new)
                data = s.encode('utf8')
            zout.writestr(item, data)


if __name__ == '__main__':
    src, old, fams = sys.argv[1], sys.argv[2], sys.argv[3:]
    os.makedirs('cmp', exist_ok=True)
    for fam in fams:
        out = os.path.join('cmp', fam.replace(' ', '') + '.docx')
        swap(src, out, old, fam)
        pdf, pages = C.page_texts(out)
        subprocess.run(['pdftoppm', '-r', '70', '-png', '-f', '1', '-l', '1', pdf,
                        out[:-5]], capture_output=True)
        print(f'{fam:16s} pages={len(pages)}')
