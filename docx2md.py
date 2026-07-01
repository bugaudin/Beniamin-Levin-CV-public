#!/usr/bin/env python3
"""Minimal DOCX -> Markdown converter using only the standard library."""
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
W = NS['w']


def tag(name):
    return f'{{{W}}}{name}'


def run_text(run):
    """Return markdown text for a run, applying bold/italic."""
    parts = []
    for node in run.iter():
        if node.tag == tag('t'):
            parts.append(node.text or '')
        elif node.tag == tag('tab'):
            parts.append('\t')
        elif node.tag == tag('br'):
            parts.append('\n')
    text = ''.join(parts)
    if not text.strip():
        return text
    rpr = run.find(tag('rPr'))
    bold = italic = False
    if rpr is not None:
        b = rpr.find(tag('b'))
        i = rpr.find(tag('i'))
        bold = b is not None and b.get(tag('val')) not in ('0', 'false')
        italic = i is not None and i.get(tag('val')) not in ('0', 'false')
    lead = len(text) - len(text.lstrip())
    trail = len(text) - len(text.rstrip())
    core = text.strip()
    if bold and italic:
        core = f'***{core}***'
    elif bold:
        core = f'**{core}**'
    elif italic:
        core = f'*{core}*'
    return text[:lead] + core + text[len(text) - trail:] if trail else text[:lead] + core


def para_to_md(p):
    text = ''.join(run_text(r) for r in p.findall(tag('r'))).strip()
    if not text:
        return ''

    style = ''
    ppr = p.find(tag('pPr'))
    is_list = False
    if ppr is not None:
        ps = ppr.find(tag('pStyle'))
        if ps is not None:
            style = (ps.get(tag('val')) or '').lower()
        if ppr.find(tag('numPr')) is not None:
            is_list = True

    if 'title' in style:
        return f'# {text}'
    if 'heading1' in style or style == 'heading 1':
        return f'## {text}'
    if 'heading2' in style:
        return f'### {text}'
    if 'heading3' in style:
        return f'#### {text}'
    if is_list:
        return f'- {text}'
    return text


def main():
    src, dst = sys.argv[1], sys.argv[2]
    with zipfile.ZipFile(src) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    body = root.find(tag('body'))

    lines = []
    for el in body:
        if el.tag == tag('p'):
            md = para_to_md(el)
            lines.append(md)
        elif el.tag == tag('tbl'):
            for row in el.findall(tag('tr')):
                cells = []
                for cell in row.findall(tag('tc')):
                    txt = ' '.join(
                        ''.join(run_text(r) for r in p.findall(tag('r'))).strip()
                        for p in cell.findall(tag('p'))
                    ).strip()
                    cells.append(txt)
                lines.append('| ' + ' | '.join(cells) + ' |')
            lines.append('')

    # Collapse multiple blank lines.
    out = []
    blank = False
    for ln in lines:
        if ln == '':
            if not blank:
                out.append('')
            blank = True
        else:
            out.append(ln)
            blank = False

    with open(dst, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out).strip() + '\n')
    print(f'Wrote {dst}')


if __name__ == '__main__':
    main()
