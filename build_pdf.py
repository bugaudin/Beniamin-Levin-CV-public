#!/usr/bin/env python3
"""Generate styled A4 CV PDF from README.md via LibreOffice."""
import os
import re
import shutil
import subprocess
import tempfile

from docx import Document
from docx.shared import Pt, RGBColor, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x44, 0x44, 0x44)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
PAGE_W_IN = 8.27
MARGIN_LR = Inches(0.28)
MARGIN_TOP = Inches(0.15)
MARGIN_BOTTOM = Inches(0.22)
LINE_SPACING = 0.93

# Base sizes at scale 1.0; scaled at build time for page-break tuning.
BASE_FONT_BODY = 9.99
BASE_FONT_SECTION = 11.59
BASE_FONT_JOB = 10.58
BASE_FONT_JOB_DATE = 9.99
BASE_FONT_ROLE = 9.99
BASE_FONT_SUBHEAD = 9.99
BASE_FONT_NAME = 24.04
BASE_FONT_CONTACT = 9.99
BASE_FONT_EDU = 9.99
BASE_FONT_EDU_DETAIL = 9.69

SECTION_TITLES = {
    'PROFESSIONAL SUMMARY',
    'CORE COMPETENCIES & TECHNICAL SKILLS',
    'PROFESSIONAL EXPERIENCE',
    'FREELANCE & PERSONAL PROJECTS — PART-TIME / OUTSIDE FULL-TIME EMPLOYMENT',
    'EDUCATION & ACADEMIC TRAINING',
    'LANGUAGES',
}

FREELANCE_SECTION = 'Freelance & Personal Projects — Part-Time / Outside Full-Time Employment'
FREELANCE_MARKER = 'freelance'
README_NAME = 'README.md'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOFFICE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'


def load_secrets():
    secrets_path = os.path.join(SCRIPT_DIR, 'secrets.txt')
    secrets = {}
    try:
        with open(secrets_path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                key, _, value = line.partition('=')
                key = key.strip()
                value = value.strip()
                if key:
                    secrets[key] = value
    except FileNotFoundError:
        pass
    if not secrets:
        raise SystemExit('secrets.txt not found or empty — add key=value entries (e.g. PHONE=...)')
    return secrets


def strip_md_bold(text):
    return re.sub(r'\*\*(.+?)\*\*', r'\1', text).strip()


def strip_md_italic(text):
    text = text.strip()
    if text.startswith('*') and text.endswith('*') and not text.startswith('**'):
        return text[1:-1].strip()
    return strip_md_bold(text)


def is_section_header(text):
    return strip_md_bold(text).upper() in SECTION_TITLES


def is_italic_line(text):
    return (
        text.startswith('*')
        and text.endswith('*')
        and not text.startswith('**')
        and not text.startswith('* ')
    )


def parse_bullet_segments(text):
    text = text.strip()
    if text.startswith('* '):
        text = text[2:]
    match = re.match(r'\*\*(.+?):\*\*\s*(.*)', text, re.DOTALL)
    if match:
        return [(match.group(1) + ': ', True), (match.group(2), False)]
    return [(strip_md_bold(text), False)]


def parse_skill_line(text):
    match = re.match(r'\*\*(.+?):\*\*\s*(.*)', text.strip())
    if not match:
        return None
    return match.group(1), match.group(2)


def parse_job_header(text):
    inner = strip_md_bold(text)
    if '\t' in inner:
        title, dates = inner.split('\t', 1)
        return title.strip(), dates.strip()
    return inner, None


def parse_readme(path):
    with open(path, encoding='utf-8') as f:
        lines = [line.rstrip() for line in f]

    data = {
        'name': None,
        'contact_bits': [],
        'linkedin': None,
        'sections': [],
    }
    current = None

    def start_section(title):
        nonlocal current
        current = {'title': title, 'blocks': []}
        data['sections'].append(current)

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped.startswith('**') and stripped.endswith('**') and data['name'] is None:
            data['name'] = strip_md_bold(stripped)
            i += 1
            continue

        if data['name'] and not data['contact_bits'] and not stripped.startswith('**'):
            if stripped.startswith('[') and '](' in stripped:
                m = re.match(r'\[(.+?)\]\((.+?)\)', stripped)
                if m:
                    data['linkedin'] = (m.group(1), m.group(2))
            else:
                data['contact_bits'].append(stripped)
            i += 1
            continue

        if stripped.startswith('[') and '](' in stripped and data['linkedin'] is None:
            m = re.match(r'\[(.+?)\]\((.+?)\)', stripped)
            if m:
                data['linkedin'] = (m.group(1), m.group(2))
            i += 1
            continue

        skill = parse_skill_line(stripped)
        if skill and current and current['title'].upper() == 'CORE COMPETENCIES & TECHNICAL SKILLS':
            current['blocks'].append({'type': 'skill', 'label': skill[0], 'rest': skill[1]})
            i += 1
            continue

        if stripped.startswith('**') and stripped.endswith('**'):
            title = strip_md_bold(stripped)
            if is_section_header(stripped):
                start_section(title)
                i += 1
                continue

            if current and current['title'].upper() in {
                'PROFESSIONAL EXPERIENCE',
                FREELANCE_SECTION.upper(),
            }:
                company, dates = parse_job_header(stripped)
                role = None
                subtitle = None
                bullets = []
                j = i + 1
                while j < len(lines):
                    nxt = lines[j].strip()
                    if not nxt:
                        j += 1
                        continue
                    if is_section_header(nxt):
                        break
                    if nxt.startswith('* ') and not nxt.startswith('**'):
                        bullets.append(parse_bullet_segments(nxt))
                        j += 1
                        continue
                    if is_italic_line(nxt):
                        if role is None:
                            role = strip_md_italic(nxt)
                        else:
                            subtitle = strip_md_italic(nxt)
                        j += 1
                        continue
                    if nxt.startswith('**'):
                        break
                    break
                current['blocks'].append({
                    'type': 'job',
                    'company': company,
                    'dates': dates,
                    'role': role,
                    'subtitle': subtitle,
                    'bullets': bullets,
                })
                i = j
                continue

            if current and current['title'].upper() == 'EDUCATION & ACADEMIC TRAINING':
                title = strip_md_bold(stripped)
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                next_line = lines[j].strip() if j < len(lines) else ''
                if '\t' in title or is_italic_line(next_line):
                    company, dates = parse_job_header(stripped)
                    role = None
                    if is_italic_line(next_line):
                        role = strip_md_italic(next_line)
                        j += 1
                    current['blocks'].append({
                        'type': 'job',
                        'company': company,
                        'dates': dates,
                        'role': role,
                        'subtitle': None,
                        'bullets': [],
                    })
                    i = j
                    continue

                current['blocks'].append({'type': 'subhead', 'text': title})
                i += 1
                continue

        if stripped.startswith('* '):
            if current:
                current['blocks'].append({
                    'type': 'bullet',
                    'segments': parse_bullet_segments(stripped),
                })
            i += 1
            continue

        if current and current['title'].upper() == 'PROFESSIONAL SUMMARY':
            current['blocks'].append({'type': 'para', 'text': stripped})
        i += 1

    return data


def scaled_pt(base, scale):
    return Pt(base * scale)


class CvBuilder:
    def __init__(self, scale, line_spacing=LINE_SPACING):
        self.scale = scale
        self.line_spacing = line_spacing
        self.font_body = scaled_pt(BASE_FONT_BODY, scale)
        self.font_section = scaled_pt(BASE_FONT_SECTION, scale)
        self.font_job = scaled_pt(BASE_FONT_JOB, scale)
        self.font_job_date = scaled_pt(BASE_FONT_JOB_DATE, scale)
        self.font_role = scaled_pt(BASE_FONT_ROLE, scale)
        self.font_subhead = scaled_pt(BASE_FONT_SUBHEAD, scale)
        self.font_name = scaled_pt(BASE_FONT_NAME, scale)
        self.font_contact = scaled_pt(BASE_FONT_CONTACT, scale)
        self.font_edu = scaled_pt(BASE_FONT_EDU, scale)
        self.font_edu_detail = scaled_pt(BASE_FONT_EDU_DETAIL, scale)
        self.content_w_in = PAGE_W_IN - 2 * MARGIN_LR.inches
        self.doc = Document()
        self._init_styles()

    def _init_styles(self):
        normal = self.doc.styles['Normal']
        normal.font.name = 'Calibri'
        normal.font.size = self.font_body
        normal.font.color.rgb = BLACK
        pf = normal.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = self.line_spacing

        sec = self.doc.sections[0]
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)
        sec.top_margin = MARGIN_TOP
        sec.bottom_margin = MARGIN_BOTTOM
        sec.left_margin = MARGIN_LR
        sec.right_margin = MARGIN_LR

    def set_space(self, p, before=0, after=0):
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        return p

    def justify(self, p):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return p

    def add_bottom_border(self, p):
        ppr = p._p.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '2')
        bottom.set(qn('w:color'), '1F3A5F')
        pbdr.append(bottom)
        ppr.append(pbdr)

    def name_heading(self, text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.set_space(p, 0, 2)
        r = p.add_run(text.upper())
        r.bold = True
        r.font.size = self.font_name
        r.font.color.rgb = ACCENT

    def contact_line(self, text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.set_space(p, 0, 1)
        r = p.add_run(text)
        r.font.size = self.font_contact
        r.font.color.rgb = GREY

    def contact_link(self, text, url):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.set_space(p, 0, 1)
        r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        run = OxmlElement('w:r')
        r_pr = OxmlElement('w:rPr')
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '444444')
        r_pr.append(color)
        size = OxmlElement('w:sz')
        size.set(qn('w:val'), str(int(self.font_contact.pt * 2)))
        r_pr.append(size)
        run.append(r_pr)
        text_el = OxmlElement('w:t')
        text_el.text = text
        run.append(text_el)
        hyperlink.append(run)
        p._p.append(hyperlink)

    def section(self, title, page_break_before=False):
        p = self.doc.add_paragraph()
        if page_break_before:
            p.paragraph_format.page_break_before = True
        self.set_space(p, 4, 2)
        r = p.add_run(title.upper())
        r.bold = True
        r.font.size = self.font_section
        r.font.color.rgb = ACCENT
        self.add_bottom_border(p)

    def skill_line(self, label, rest):
        p = self.justify(self.doc.add_paragraph())
        self.set_space(p, 0.85, 0.85)
        r = p.add_run(label + ': ')
        r.bold = True
        r.font.color.rgb = ACCENT
        r.font.size = self.font_body
        r2 = p.add_run(rest)
        r2.font.size = self.font_body

    def dated_line(self, title, dates, title_font=None, date_bold=True):
        if title_font is None:
            title_font = self.font_job
        p = self.doc.add_paragraph()
        self.set_space(p, 2.8, 0)
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(self.content_w_in), WD_TAB_ALIGNMENT.RIGHT
        )
        r = p.add_run(title)
        r.bold = True
        r.font.size = title_font
        r.font.color.rgb = BLACK
        if dates:
            rd = p.add_run('\t' + dates)
            rd.bold = date_bold
            rd.font.size = self.font_job_date
            rd.font.color.rgb = GREY

    def job(self, company, role, dates, subtitle=None):
        self.dated_line(company, dates)
        if role:
            pr = self.doc.add_paragraph()
            self.set_space(pr, 0, 1.5)
            ri = pr.add_run(role)
            ri.italic = True
            ri.font.size = self.font_role
            ri.font.color.rgb = ACCENT
        if subtitle:
            ps = self.doc.add_paragraph()
            self.set_space(ps, 0, 1.5)
            rs = ps.add_run(subtitle)
            rs.italic = True
            rs.font.size = self.font_role
            rs.font.color.rgb = GREY

    def bullet(self, segments):
        p = self.justify(self.doc.add_paragraph(style='List Bullet'))
        self.set_space(p, 0.2, 0.2)
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.first_line_indent = Inches(-0.13)
        for text, bold in segments:
            r = p.add_run(text)
            r.bold = bold
            r.font.size = self.font_body

    def edu(self, inst, dates, detail):
        self.dated_line(inst, dates, title_font=self.font_edu, date_bold=True)
        pd = self.justify(self.doc.add_paragraph())
        self.set_space(pd, 0, 1)
        ri = pd.add_run(detail)
        ri.italic = True
        ri.font.size = self.font_edu_detail
        ri.font.color.rgb = GREY

    def subhead(self, text):
        p = self.doc.add_paragraph()
        self.set_space(p, 5, 1)
        r = p.add_run(text)
        r.bold = True
        r.font.size = self.font_subhead
        r.font.color.rgb = BLACK

    def para(self, text, after=1.5):
        p = self.justify(self.doc.add_paragraph())
        self.set_space(p, 0, after)
        r = p.add_run(text)
        r.font.size = self.font_body

    def build_from_data(self, data, phone):
        self.name_heading(data['name'])
        contact = data['contact_bits'][0] if data['contact_bits'] else ''
        contact = re.sub(
            r'\[([^\]]+)\]\([^)]+\)',
            r'\1',
            contact,
        )
        if '|' in contact:
            parts = [part.strip() for part in contact.split('|')]
            contact = '   |   '.join([parts[0], phone] + parts[1:])
        else:
            contact = f'{contact}   |   {phone}'
        self.contact_line(contact)
        if data['linkedin']:
            self.contact_link(data['linkedin'][0], data['linkedin'][1])

        summary_blocks = None
        for section in data['sections']:
            title_upper = section['title'].upper()
            page_break = title_upper == FREELANCE_SECTION.upper()
            self.section(section['title'], page_break_before=page_break)

            if title_upper == 'PROFESSIONAL SUMMARY':
                paras = [b for b in section['blocks'] if b['type'] == 'para']
                for idx, block in enumerate(paras):
                    after = 0 if idx == len(paras) - 1 else 1.5
                    self.para(block['text'], after=after)
                continue

            if title_upper == 'CORE COMPETENCIES & TECHNICAL SKILLS':
                for block in section['blocks']:
                    self.skill_line(block['label'], block['rest'])
                continue

            if title_upper in {'PROFESSIONAL EXPERIENCE', FREELANCE_SECTION.upper()}:
                for block in section['blocks']:
                    if block['type'] != 'job':
                        continue
                    self.job(
                        block['company'],
                        block['role'],
                        block['dates'],
                        subtitle=block.get('subtitle'),
                    )
                    for segments in block['bullets']:
                        self.bullet(segments)
                continue

            if title_upper == 'EDUCATION & ACADEMIC TRAINING':
                for block in section['blocks']:
                    if block['type'] == 'job':
                        self.edu(block['company'], block['dates'], block['role'])
                    elif block['type'] == 'subhead':
                        self.subhead(block['text'])
                    elif block['type'] == 'bullet':
                        self.bullet(block['segments'])
                continue

            if title_upper == 'LANGUAGES':
                for block in section['blocks']:
                    if block['type'] == 'bullet':
                        self.bullet(block['segments'])


def convert_docx_to_pdf(docx_path, pdf_path):
    outdir = os.path.dirname(pdf_path)
    subprocess.run(
        [SOFFICE, '--headless', '--convert-to', 'pdf', '--outdir', outdir, docx_path],
        check=True,
        capture_output=True,
    )
    generated = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
    if generated != pdf_path:
        shutil.move(generated, pdf_path)


def page_text(pdf_path, page_num):
    result = subprocess.run(
        ['pdftotext', '-f', str(page_num), '-l', str(page_num), pdf_path, '-'],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


def freelance_starts_on_page_2(pdf_path):
    page1 = page_text(pdf_path, 1).lower()
    page2 = page_text(pdf_path, 2).lower()
    if FREELANCE_MARKER in page1:
        return False
    return FREELANCE_MARKER in page2


def layout_ok(pdf_path):
    return freelance_starts_on_page_2(pdf_path)


def build_trial(data, phone, tmp_dir, tag, scale, line_spacing):
    docx_path = os.path.join(tmp_dir, f'cv-{tag}.docx')
    pdf_path = os.path.join(tmp_dir, f'cv-{tag}.pdf')
    builder = CvBuilder(scale, line_spacing=line_spacing)
    builder.build_from_data(data, phone)
    builder.doc.save(docx_path)
    convert_docx_to_pdf(docx_path, pdf_path)
    return pdf_path


def find_best_layout(data, phone, tmp_dir):
    lo, hi = 0.90, 1.20
    best_scale = lo
    for i in range(14):
        mid = (lo + hi) / 2
        pdf_path = build_trial(data, phone, tmp_dir, f'scale-{i}', mid, LINE_SPACING)
        if layout_ok(pdf_path):
            best_scale = mid
            lo = mid
        else:
            hi = mid

    lo, hi = LINE_SPACING, 0.97
    best_line_spacing = LINE_SPACING
    for i in range(12):
        mid = (lo + hi) / 2
        pdf_path = build_trial(
            data, phone, tmp_dir, f'ls-{i}', best_scale, mid
        )
        if layout_ok(pdf_path):
            best_line_spacing = mid
            lo = mid
        else:
            hi = mid

    return best_scale, best_line_spacing


def main():
    secrets = load_secrets()
    if 'PHONE' not in secrets:
        raise SystemExit('secrets.txt must contain PHONE=...')
    phone = secrets['PHONE']

    readme_path = os.path.join(SCRIPT_DIR, README_NAME)
    data = parse_readme(readme_path)

    out_pdf = os.path.join(SCRIPT_DIR, 'Beniamin-Levin-CV.pdf')
    legacy = os.path.join(SCRIPT_DIR, 'beniamin-levin-cv.pdf')

    with tempfile.TemporaryDirectory() as tmp:
        best_scale, best_line_spacing = find_best_layout(data, phone, tmp)
        final_docx = os.path.join(tmp, 'cv-final.docx')
        final_pdf = os.path.join(tmp, 'cv-final.pdf')
        builder = CvBuilder(best_scale, line_spacing=best_line_spacing)
        builder.build_from_data(data, phone)
        builder.doc.save(final_docx)
        convert_docx_to_pdf(final_docx, final_pdf)

        if os.path.lexists(legacy):
            os.remove(legacy)
        if os.path.lexists(out_pdf):
            os.remove(out_pdf)
        shutil.copy2(final_pdf, out_pdf)

    body_pt = BASE_FONT_BODY * best_scale
    print(
        f'saved {out_pdf} (scale={best_scale:.3f}, body={body_pt:.2f}pt, '
        f'line_spacing={best_line_spacing:.3f}, bottom_margin={MARGIN_BOTTOM.inches:.2f}in)'
    )


if __name__ == '__main__':
    main()
