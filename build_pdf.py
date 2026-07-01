#!/usr/bin/env python3
"""Generate styled A4 .docx CV with justified body text."""
import os
import shutil
import subprocess
import tempfile

from docx import Document
from docx.shared import Pt, RGBColor, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x44, 0x44, 0x44)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
PAGE_W_IN = 8.27
MARGIN_LR = Inches(0.30)
MARGIN_TOP = Inches(0.18)
MARGIN_BOTTOM = Inches(0.12)
CONTENT_W_IN = PAGE_W_IN - 2 * MARGIN_LR.inches

# Tuned so page 1 ends after Professional Experience and Freelance starts at top of page 2.
FONT_BODY = Pt(9.99)
FONT_SECTION = Pt(11.59)
FONT_JOB = Pt(10.58)
FONT_JOB_DATE = Pt(9.99)
FONT_ROLE = Pt(9.99)
FONT_SUBHEAD = Pt(9.99)
FONT_NAME = Pt(24.04)
FONT_CONTACT = Pt(9.99)
FONT_EDU = Pt(9.99)
FONT_EDU_DETAIL = Pt(9.69)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


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

doc = Document()

normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = FONT_BODY
normal.font.color.rgb = BLACK
pf = normal.paragraph_format
pf.space_before = Pt(0)
pf.space_after = Pt(0)
pf.line_spacing = 0.93

sec = doc.sections[0]
sec.page_width = Mm(210)
sec.page_height = Mm(297)
sec.top_margin = MARGIN_TOP
sec.bottom_margin = MARGIN_BOTTOM
sec.left_margin = MARGIN_LR
sec.right_margin = MARGIN_LR


def set_space(p, before=0, after=0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    return p


def justify(p):
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_bottom_border(p):
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '2')
    bottom.set(qn('w:color'), '1F3A5F')
    pbdr.append(bottom)
    ppr.append(pbdr)


def name_heading(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_space(p, 0, 2)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = FONT_NAME
    r.font.color.rgb = ACCENT
    return p


def contact_line(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_space(p, 0, 1)
    r = p.add_run(text)
    r.font.size = FONT_CONTACT
    r.font.color.rgb = GREY
    return p


def contact_link(text, url):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_space(p, 0, 1)
    r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run = OxmlElement('w:r')
    r_pr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '444444')
    r_pr.append(color)
    size = OxmlElement('w:sz')
    size.set(qn('w:val'), str(int(FONT_CONTACT.pt * 2)))
    r_pr.append(size)
    run.append(r_pr)
    text_el = OxmlElement('w:t')
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    p._p.append(hyperlink)
    return p


def section(title, page_break_before=False):
    p = doc.add_paragraph()
    if page_break_before:
        p.paragraph_format.page_break_before = True
    set_space(p, 4, 2)
    r = p.add_run(title.upper())
    r.bold = True
    r.font.size = FONT_SECTION
    r.font.color.rgb = ACCENT
    add_bottom_border(p)
    return p


def skill_line(label, rest):
    p = justify(doc.add_paragraph())
    set_space(p, 0.85, 0.85)
    r = p.add_run(label + ': ')
    r.bold = True
    r.font.color.rgb = ACCENT
    r.font.size = FONT_BODY
    r2 = p.add_run(rest)
    r2.font.size = FONT_BODY
    return p


def dated_line(title, dates, title_font=FONT_JOB, date_bold=True):
    p = doc.add_paragraph()
    set_space(p, 2.8, 0)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(CONTENT_W_IN), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(title)
    r.bold = True
    r.font.size = title_font
    r.font.color.rgb = BLACK
    rd = p.add_run('\t' + dates)
    rd.bold = date_bold
    rd.font.size = FONT_JOB_DATE
    rd.font.color.rgb = GREY
    return p


def job(company, role, dates):
    dated_line(company, dates)
    pr = doc.add_paragraph()
    set_space(pr, 0, 1.5)
    ri = pr.add_run(role)
    ri.italic = True
    ri.font.size = FONT_ROLE
    ri.font.color.rgb = ACCENT


def bullet(segments):
    p = justify(doc.add_paragraph(style='List Bullet'))
    set_space(p, 0.2, 0.2)
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.first_line_indent = Inches(-0.13)
    if isinstance(segments, str):
        segments = [(segments, False)]
    for text, bold in segments:
        r = p.add_run(text)
        r.bold = bold
        r.font.size = FONT_BODY
    return p


def edu(inst, dates, detail):
    dated_line(inst, dates, title_font=FONT_EDU, date_bold=True)
    pd = justify(doc.add_paragraph())
    set_space(pd, 0, 1)
    ri = pd.add_run(detail)
    ri.italic = True
    ri.font.size = FONT_EDU_DETAIL
    ri.font.color.rgb = GREY


def subhead(text):
    p = doc.add_paragraph()
    set_space(p, 5, 1)
    r = p.add_run(text)
    r.bold = True
    r.font.size = FONT_SUBHEAD
    r.font.color.rgb = BLACK


def para(text, after=1.5):
    p = justify(doc.add_paragraph())
    set_space(p, 0, after)
    r = p.add_run(text)
    r.font.size = FONT_BODY
    return p
name_heading('Beniamin Levin')
_secrets = load_secrets()
if 'PHONE' not in _secrets:
    raise SystemExit('secrets.txt must contain PHONE=...')
contact_line(f'Kiryat Yam, Haifa District, Israel   |   {_secrets["PHONE"]}   |   benlevin0213@gmail.com')
contact_link('linkedin.com/in/beniaminlevin', 'https://www.linkedin.com/in/beniaminlevin/')

section('Professional Summary')
para('Tech Lead with 19+ years of experience building, scaling, and supporting enterprise telecom OSS/BSS, automotive, and mobile platforms. Deep hands-on background in Java/Spring engineering, system design, distributed delivery, and Kubernetes/OpenShift-based cloud delivery.')
para('Led distributed engineering teams of up to 11 developers across European development centres and India. Delivered production systems for major European telecom programmes including Telenet, Telefonica UK/O2, Vodafone Hungary, Nuuday/TDC, M-net, T-Mobile Netherlands/Odido, and Andorra Telecom.')
para('Delivered measurable engineering and business impact: reduced regression defects by 60% through CI and code-review improvements, reduced page load times by 30–80% through profiling, caching, and SQL tuning, improved accessibility compliance, and established a Pune engineering pod by running 70+ interviews, hiring 5 engineers, and leading on-site architecture mentoring.')
para('Hands-on with AI-assisted engineering tooling including LLM-based defect triage and repository-aware code analysis integrated into team SDLC.')
para('Seeking relocation to the Netherlands for a Tech Lead role with an IND-recognised employer under Highly Skilled Migrant sponsorship.', after=0)

section('Core Competencies & Technical Skills')
skill_line('Technical Leadership & Delivery', 'Engineering Leadership, System Design, Distributed Systems, Team Mentoring, Agile/Scrum, Sprint Planning, Story Mapping, L3 Production Support')
skill_line('Backend', 'Java, Spring Boot, Spring Cloud, Quarkus, REST, GraphQL, Microservices')
skill_line('Messaging & Data', 'Kafka, RabbitMQ, Redis, Cassandra, MySQL, SQL')
skill_line('Cloud & DevOps', 'Docker, Kubernetes, OpenShift, GCP, AWS, Azure, CI/CD')
skill_line('Frontend & Mobile', 'Angular, JavaScript, jQuery, HTML5, CSS3, PHP, Android')
skill_line('AI-Assisted Engineering', 'Cursor, Claude Sonnet, OpenAI GPT, Grok (xAI) API, natural-language-to-SQL, LLM prompt engineering, AI-assisted defect triage, repository-aware code analysis')

section('Professional Experience')
job('Netcracker Technology — Haifa, Israel / Nizhny Novgorod', 'Lead Software Engineer', 'July 2015 – Present')
bullet('Reduced regression defects by 60% by introducing code-review guidelines, improving CI pipelines, and strengthening release-quality practices across the team.')
bullet('Optimised business-critical order-entry flows to meet customer non-functional requirements, reducing page load times by 30–80% through profiling, caching improvements, and SQL tuning.')
bullet('Introduced AI-assisted development workflows using Cursor across the team for coding, refactoring, and unit-test generation, spanning enterprise Java/Spring and Angular codebases.')
bullet('Configured and operationalised a Claude Sonnet-based AI defect-triage assistant by enabling repository and documentation access and defining project-specific rules and skills to surface likely root causes and suggested fixes.')
bullet('Established and scaled an engineering pod in Pune, India: ran 70+ interviews, hired 5 engineers, and completed two multi-month on-site trips for architecture mentoring and knowledge transfer.')
bullet('Led technical design, feature development, and scaling of enterprise telecom OSS/BSS product suites across Residential/Business Order Entry, CPQ, Cloud BSS, Product Catalog, Sales Order Engine/Validation, BAIP asset ordering, and Kubernetes-based delivery streams.')
bullet('Directed a distributed engineering team of up to 11 developers across European development centres, owning sprint planning, story mapping, delivery timelines, code quality, and Level 3 production incident triage.')
bullet('Delivered product enhancements, defect resolution, production support, release integration, backports, security fixes, and customer specific improvements across multiple Jira streams and 20+ customers.')
bullet('Supported major European telecom programmes including Telenet, Telefonica UK/O2, Vodafone Hungary, Nuuday/TDC, M-net, T-Mobile Netherlands/Odido, Andorra Telecom, and other regional operators.')
bullet('Owned complex merge, rebase, dependency, and backport delivery across ROE, BOE, SOE/SOV, POC, OEF, CPQ, and Cloud BSS release streams.')
bullet('Improved ADA/accessibility compliance for Charter Communications by resolving keyboard-only navigation defects, JAWS screen-reader behaviour, and inaccessible error-message tooltip issues in CPQ/Quotes UI flows.')

job('Symphony Teleca Corporation — Nizhny Novgorod', 'Principal Software Engineer', 'February 2011 – July 2015')
bullet([('Samsung Messaging — Suwon, South Korea: ', True), ('Selected for an on-site engagement at Samsung HQ to optimise and customise core messaging infrastructure for flagship Android devices.', False)])
bullet([('SiriusXM Connected Vehicle Platform: ', True), ('Engineered the Java server-side architecture and Android client for a vehicle infotainment platform demonstrated live at CES 2012.', False)])
bullet([('Android UI & Widgets — Nokia, Savelli: ', True), ('Developed low-level Android UI adaptations replicating native Nokia experiences, and built weather/alert widgets focused on responsiveness, asynchronous loading, and low power consumption.', False)])

job('Onlyoffice — Nizhny Novgorod', 'Software Engineer', 'November 2009 – January 2011')
bullet('Designed and implemented operational subsystems for TeamLab, now Onlyoffice, an enterprise collaboration and document management platform.')
bullet('Delivered modules for bookmark systems, corporate planners/organisers, project templates, real-time analytics aggregators, and automated bulk LDAP/employee data ingestion pipelines.')

job('MERA NN — Nizhny Novgorod', 'Software Engineer', 'November 2006 – November 2009')
bullet('Built high-availability web applications for operating, remotely administering, and monitoring Nortel ICP communication infrastructure.')
bullet('Authored architecture specifications and deployment plans; maintained JUnit/Cobertura test suites to raise coverage and reliability.')

section('Freelance & Personal Projects — Part-Time / Outside Full-Time Employment', page_break_before=True)
bullet('Delivered end-to-end PHP/MySQL/JavaScript web applications, customer portals, CRM integrations, and back-office tooling for telecom and business clients, including the flagship dealer/consumer sales platform for Morgan Motor Company, covering requirements analysis, deployment, support, and iterative enhancement.')
bullet('Led technical architecture, backend automation, performance optimisation, and technical oversight for enterprise platforms CRMs; built file-rendering services, cron task runners, email to comment synchronisation, database tuning, an OpenAI ChatGPT-powered email-reply generator (context-aware drafting with greeting personalisation, thread trimming, and chase-email detection), and an AI ticket comment generator that drafts support responses from full ticket context and conversation history.')
bullet('Designed and built "SurAI", a natural-language business-intelligence copilot embedded in the CRM that converts plain-English questions into schema-aware SQL via the Grok (xAI) API, injecting table/column metadata and comments to ground the model and prevent hallucinated fields; enforced safety through a read-only database connection, SELECT-only guardrails, and column validation against the live schema, then auto-rendered results as hyperlinked record tables and Chart.js visualisations (bar/pie/line) for non-technical staff.')
bullet('Designed, built, deployed, and maintained a bespoke PHP/MySQL CRM for client, contact, and sales-workflow management, owning architecture, cloud provisioning, production migration, and ongoing support as sole developer.')
bullet('Initially deployed on Microsoft Azure before migrating production infrastructure to AWS EC2/RDS for improved cost and operational control.')

section('Education & Academic Training')
edu('N. I. Lobachevsky State University of Nizhny Novgorod', '2003 – 2008', "Master's, Faculty of Computational Mathematics and Cybernetics — Graduated with Honours")
edu('Nortel Networks (New York, USA)', 'March 2009 – June 2009', 'Advanced Corporate Training: Common OA&M Architecture')

subhead('Engineering Leadership & Management')
bullet('Team Management Programme: advanced leadership, delegation, and delivery frameworks')
bullet('Situational Management & Leadership Dynamics: interactive corporate residency')
bullet('Core Leadership Skills & Multi-Site Team Orchestration')
bullet('People Development, Performance Assessment & Engineering Accreditation')

subhead('Advanced Software Engineering & Frameworks')
bullet('Angular 2+: advanced architecture & state management')
bullet("Hands-On Microservices with Spring Boot and Spring Cloud, 40h, O'Reilly")
bullet('Introduction to Containers with Docker, Kubernetes & OpenShift, Coursera')
bullet('Quarkus Super-Heroes Workshop')
bullet('Redis for Java Developers & Redis Data Structures, Redis University')
bullet('DS101: Introduction to Apache Cassandra, DataStax Academy')
bullet("Apache Kafka and RabbitMQ messaging with Java and Spring, O'Reilly")
bullet('GraphQL for Java / full-stack development')
bullet('Spring Framework Ecosystem & Application Architecture')
bullet('Post-Production High-Availability Performance Support & Best Practices')

section('Languages')
bullet([('English: ', True), ('Full professional working proficiency — 19+ years of daily use in international engineering teams, technical documentation, code review, and customer-facing delivery; includes on-site work in India and South Korea, plus corporate training in New York, USA.', False)])
bullet([('Hebrew: ', True), ('Upper-intermediate — professional working proficiency', False)])
bullet([('Russian: ', True), ('Native', False)])

out_dir = SCRIPT_DIR
out_pdf = os.path.join(out_dir, 'Beniamin-Levin-CV.pdf')
soffice = '/Applications/LibreOffice.app/Contents/MacOS/soffice'

with tempfile.TemporaryDirectory() as tmp:
    tmp_docx = os.path.join(tmp, 'cv.docx')
    doc.save(tmp_docx)
    subprocess.run(
        [soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp, tmp_docx],
        check=True,
        capture_output=True,
    )
    staged = os.path.join(tmp, 'cv-staged.pdf')
    shutil.copy2(os.path.join(tmp, 'cv.pdf'), staged)
    legacy = os.path.join(out_dir, 'beniamin-levin-cv.pdf')
    if os.path.lexists(legacy):
        os.remove(legacy)
    if os.path.lexists(out_pdf):
        os.remove(out_pdf)
    shutil.move(staged, out_pdf)

print('saved', out_pdf)
