#!/usr/bin/env python3
"""Generate nl-job-application-tracker.html from nl-job-application-tracker.md."""

import json
import re
from pathlib import Path
from typing import Optional

ROOT = Path(__file__).parent
MD_PATH = ROOT / "nl-job-application-tracker.md"
HTML_PATH = ROOT / "nl-job-application-tracker.html"

# Stale job URLs → stable careers search pages (verified 2026-06-29).
LINK_REPLACEMENTS = {
    "https://www.backbase.com/careers/jobs/7981900-lead-engineering-team-lead-etl": "https://www.backbase.com/careers/jobs?office=amsterdam",
    "https://careers.vodafoneziggo.com/vacatures/chapter-lead-software-engineering-6399": "https://careers.vodafoneziggo.com/werkgebieden/tech-data/network-cloud",
    "https://jobs.kpn.com/vacature/1754/lead-engineer-platform-development": "https://jobs.kpn.com/vacatures",
    "https://www.libertyglobal.com/careers/search-jobs/": "https://www.libertyglobal.com/careers/",
    "https://www.werkenbijabnamro.nl/en/vacancy/9316/java-developer-11": "https://www.werkenbijabnamro.nl/en/vacancies",
    "https://www.payconiq.com/careers": "https://careers.payconiq.com/en",
    "https://www.elastic.co/careers/jobs": "https://jobs.elastic.co/",
    "https://irdeto.com/careers/current-openings/": "https://careers.irdeto.com/",
    "https://www.mendix.com/careers/open-positions/": "https://www.mendix.com/company/careers/",
    "https://planonsoftware.com/about-planon/careers/": "https://planonsoftware.com/careers/",
    "https://www.hcltech.com/careers/Careers-in-Netherlands": "https://www.hcltech.com/careers",
    "https://www.sogeti.com/nl-nl/carriere/": "https://www.sogeti.com/careers",
    "https://www.cgi.com/nl/nl/karriere": "https://www.cgi.com/en/careers",
    "https://careers.bol.com/en/vacancies/": "https://careers.bol.com/en/jobs/",
    "https://www.abnamro.com/en/careers/job-openings": "https://www.werkenbijabnamro.nl/en/vacancies",
    "https://www.capgemini.com/nl-nl/careers/jobs/": "https://careers.capgemini.com/",
    "https://www.uber.com/us/en/careers/": "https://www.uber.com/global/en/careers/list/?location=Amsterdam%2C-North-Holland%2C-Netherlands",
    "https://www.leaseweb.com/careers/jobs": "https://www.leaseweb.com/en/career/jobs",
    "https://schubergphilis.com/careers/software-engineer-java": "https://schubergphilis.com/en/careers/",
    "https://schubergphilis.com/careers": "https://schubergphilis.com/en/careers/",
    "https://www.unit4.com/careers/vacancies": "https://www.unit4.com/about-us/careers",
}


def _load_table_column(path: Path, column_index: int) -> dict[int, str]:
    if not path.exists():
        return {}
    values: dict[int, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|") or line.startswith("| #") or line.startswith("|---"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) <= column_index or not parts[1].isdigit():
            continue
        values[int(parts[1])] = parts[column_index]
    return values


def _load_salary_forks() -> dict[int, str]:
    return _load_table_column(ROOT / "nl-job-salary-forks.md", 3)


def _load_work_arrangements() -> dict[int, dict[str, str]]:
    path = ROOT / "nl-job-work-arrangements.md"
    if not path.exists():
        return {}
    rows: dict[int, dict[str, str]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|") or line.startswith("| #") or line.startswith("|---"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 6 or not parts[1].isdigit():
            continue
        rows[int(parts[1])] = {
            "model": parts[3],
            "days": parts[4],
            "note": parts[5] if len(parts) > 5 else "",
        }
    return rows


def _work_line(work: dict[str, str]) -> str:
    if not work:
        return ""
    return f"- **Work model:** {work['model']} · **Min office:** {work['days']}"


def dedupe_work_model_lines(text: str) -> str:
    return re.sub(
        r"(- \*\*Work model:\*\* [^\n]+\n)(?:- \*\*Work model:\*\* [^\n]+\n)+",
        r"\1",
        text,
    )


def fix_tracker_links(text: str) -> str:
    for old, new in LINK_REPLACEMENTS.items():
        text = text.replace(old, new)
    return text


def inject_work_into_tracker(md_path: Path, work: dict[int, dict[str, str]]) -> None:
    text = dedupe_work_model_lines(md_path.read_text(encoding="utf-8"))
    for company_id, info in work.items():
        line = _work_line(info)
        block_pattern = rf"(### {company_id}\. [^\n]+\n\n- \[.\] Applied\n\n)"
        if re.search(
            rf"### {company_id}\. [^\n]+\n\n- \[.\] Applied\n\n- \*\*Work model:\*\*",
            text,
        ):
            text = re.sub(
                rf"(### {company_id}\. [^\n]+\n\n- \[.\] Applied\n\n)- \*\*Work model:\*\* [^\n]+\n",
                rf"\1{line}\n",
                text,
                count=1,
            )
            continue
        text = re.sub(
            block_pattern,
            rf"\1{line}\n",
            text,
            count=1,
        )
    md_path.write_text(text, encoding="utf-8")


def _extract_url(value: str) -> str:
    """Pull href from markdown [text](url) or return plain URL/text."""
    if not value:
        return ""
    value = value.strip()
    m = re.match(r"\[([^\]]*)\]\(([^)]+)\)", value)
    if m:
        return m.group(2).strip()
    return value


def parse_markdown(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8")
    companies: list[dict] = []
    current_tier = ""
    blocks = re.split(r"\n(?=### \d+\.)", text)

    for block in blocks:
        header = re.match(r"### (\d+)\.\s+(.+?)\s*\n", block)
        if not header:
            tier_match = re.match(r"## (Tier \d+[^\n]+)", block)
            if tier_match:
                current_tier = tier_match.group(1).strip()
            continue

        num, name = header.group(1), header.group(2).strip()
        name = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", name)

        def field(label: str) -> str:
            m = re.search(rf"- \*\*{re.escape(label)}:\*\* (.+)", block)
            return m.group(1).strip() if m else ""

        applied = "- [x]" in block.split("Applied date:")[0]

        companies.append(
            {
                "id": int(num),
                "name": name,
                "tier": current_tier,
                "roles": field("Best-fit roles"),
                "official": _extract_url(field("Official")),
                "allVacancies": _extract_url(
                    field("All vacancies") or field("Job board") or field("Job search")
                ),
                "linkedin": _extract_url(field("LinkedIn")),
                "note": field("Note"),
                "sources": field("Sources"),
                "applied": applied,
            }
        )

    salaries = _load_salary_forks()
    work = _load_work_arrangements()
    for company in companies:
        company["salary"] = salaries.get(company["id"], "")
        w = work.get(company["id"], {})
        company["workModel"] = w.get("model", "")
        company["workDays"] = w.get("days", "")
        company["workNote"] = w.get("note", "")

    return companies


def _migrate_state_to_names(state: dict, companies: list[dict]) -> dict:
    """Map legacy numeric company IDs to stable company names."""
    if not state:
        return {}
    id_to_name = {c["id"]: c["name"] for c in companies}
    id_to_name.update({str(c["id"]): c["name"] for c in companies})
    migrated: dict = {}
    for key, value in state.items():
        if not isinstance(value, dict):
            continue
        name = id_to_name.get(key)
        if name is None and str(key).isdigit():
            name = id_to_name.get(int(key))
        target = name if name else key
        migrated[target] = value
    return migrated


def _load_embedded_state(html_path: Path) -> dict:
    """Preserve tracker progress embedded in the HTML file across rebuilds."""
    if not html_path.exists():
        return {}
    text = html_path.read_text(encoding="utf-8")
    match = re.search(
        r'<script type="application/json" id="embedded-state">\s*([\s\S]*?)\s*</script>',
        text,
    )
    if not match:
        return {}
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError:
        return {}


def render_html(companies: list[dict], embedded_state: Optional[dict] = None) -> str:
    data = json.dumps(companies, ensure_ascii=False)
    embedded = json.dumps(embedded_state or {}, ensure_ascii=False, indent=2)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>NL Job Application Tracker</title>
  <style>
    :root {{
      --bg: #f6f7f9;
      --card: #fff;
      --text: #1a1a1a;
      --muted: #5c6570;
      --border: #e2e6ea;
      --accent: #0b57d0;
      --accent-soft: #e8f0fe;
      --done: #137333;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.45;
    }}
    header {{
      position: sticky;
      top: 0;
      z-index: 10;
      background: rgba(246, 247, 249, 0.95);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid var(--border);
      padding: 1rem 1.25rem;
    }}
    h1 {{ margin: 0 0 0.35rem; font-size: 1.35rem; }}
    .sub {{ color: var(--muted); font-size: 0.9rem; margin-bottom: 0.75rem; }}
    .stats {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.75rem;
      font-size: 0.9rem;
    }}
    .stat {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 0.4rem 0.7rem;
    }}
    .stat strong {{ color: var(--accent); }}
    .toolbar {{
      margin-top: 0.75rem;
      display: flex;
      flex-wrap: wrap;
      gap: 0.5rem;
      align-items: center;
    }}
    input[type="search"] {{
      flex: 1;
      min-width: 200px;
      padding: 0.5rem 0.65rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      font-size: 0.95rem;
    }}
    button {{
      border: 1px solid var(--border);
      background: var(--card);
      border-radius: 8px;
      padding: 0.45rem 0.7rem;
      cursor: pointer;
      font-size: 0.9rem;
    }}
    button:hover {{ border-color: var(--accent); }}
    main {{ max-width: 920px; margin: 0 auto; padding: 1rem 1.25rem 2rem; }}
    .tier {{
      margin: 1.25rem 0 0.5rem;
      font-size: 0.85rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      color: var(--muted);
    }}
    .card {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 0.9rem 1rem;
      margin-bottom: 0.65rem;
      transition: border-color 0.15s, box-shadow 0.15s;
    }}
    .card.applied {{
      border-color: #b7dfc8;
      background: #f8fdf9;
    }}
    .card-head {{
      display: flex;
      gap: 0.75rem;
      align-items: flex-start;
    }}
    .card-head input[type="checkbox"] {{
      width: 1.15rem;
      height: 1.15rem;
      margin-top: 0.15rem;
      cursor: pointer;
      flex-shrink: 0;
    }}
    .title-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.5rem;
      align-items: baseline;
      width: 100%;
    }}
    .num {{ color: var(--muted); font-size: 0.85rem; }}
    .name {{ font-size: 1.05rem; font-weight: 650; }}
    .saved {{
      margin-left: auto;
      font-size: 0.75rem;
      color: var(--done);
      opacity: 0;
      transition: opacity 0.2s;
    }}
    .saved.show {{ opacity: 1; }}
    .roles {{ color: var(--muted); font-size: 0.9rem; margin: 0.35rem 0 0.5rem; }}
    .links {{ display: flex; flex-wrap: wrap; gap: 0.45rem; margin-bottom: 0.55rem; }}
    .links a {{
      font-size: 0.85rem;
      color: var(--accent);
      text-decoration: none;
      background: var(--accent-soft);
      padding: 0.2rem 0.5rem;
      border-radius: 6px;
    }}
    .links a:hover {{ text-decoration: underline; }}
    .meta {{ font-size: 0.8rem; color: var(--muted); margin-bottom: 0.5rem; }}
    .fields {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: 0.45rem;
    }}
    .fields label {{
      display: block;
      font-size: 0.75rem;
      color: var(--muted);
      margin-bottom: 0.15rem;
    }}
    .fields input {{
      width: 100%;
      padding: 0.35rem 0.45rem;
      border: 1px solid var(--border);
      border-radius: 6px;
      font-size: 0.85rem;
    }}
    .note {{
      margin-top: 0.45rem;
      font-size: 0.82rem;
      color: #8a4b00;
      background: #fff8e8;
      border-radius: 6px;
      padding: 0.35rem 0.5rem;
    }}
    .hidden {{ display: none !important; }}
    .save-status {{
      font-size: 0.8rem;
      color: var(--done);
      margin-top: 0.35rem;
      min-height: 1.1rem;
    }}
    footer {{
      text-align: center;
      color: var(--muted);
      font-size: 0.8rem;
      padding: 1rem;
    }}
  </style>
</head>
<body>
  <header>
    <h1>NL Job Application Tracker</h1>
    <div class="sub">Auto-saves by company name to localStorage (+ optional .html file link in Chrome/Edge)</div>
    <div class="save-status" id="save-status"></div>
    <div class="stats">
      <div class="stat">Total: <strong id="stat-total">0</strong></div>
      <div class="stat">Applied: <strong id="stat-applied">0</strong></div>
      <div class="stat">Remaining: <strong id="stat-remaining">0</strong></div>
    </div>
    <div class="toolbar">
      <input type="search" id="search" placeholder="Search company or role…" />
      <button type="button" id="filter-applied">Show applied only</button>
      <button type="button" id="filter-pending">Show pending only</button>
      <button type="button" id="filter-all">Show all</button>
      <button type="button" id="export-json">Export backup</button>
      <button type="button" id="link-file">Link HTML file</button>
      <button type="button" id="clear-storage">Clear saved data</button>
    </div>
  </header>
  <main id="list"></main>
  <footer>Regenerate from markdown: <code>python3 build_tracker_html.py</code></footer>
  <script type="application/json" id="embedded-state">
{embedded}
  </script>
  <script>
    const COMPANIES = {data};
    const STORAGE_KEY = "nl-job-application-tracker-v2";
    const LEGACY_STORAGE_KEY = "nl-job-application-tracker-v1";
    const FILE_DB = "nl-job-tracker";
    const FILE_STORE = "handles";
    const FILE_KEY = "html-file";
    const EMBEDDED_RE = /(<script type="application\\/json" id="embedded-state">)[\\s\\S]*?(<\\/script>)/;

    function companyKey(c) {{
      return c.name.trim();
    }}

    function migrateLegacyState(raw) {{
      if (!raw || typeof raw !== "object") return {{}};
      const migrated = {{}};
      for (const c of COMPANIES) {{
        const legacy = raw[c.id] ?? raw[String(c.id)];
        if (legacy) migrated[companyKey(c)] = legacy;
      }}
      for (const [key, val] of Object.entries(raw)) {{
        if (val && typeof val === "object" && "applied" in val && !migrated[key]) {{
          migrated[key] = val;
        }}
      }}
      return migrated;
    }}

    function loadEmbeddedState() {{
      const el = document.getElementById("embedded-state");
      if (!el) return {{}};
      try {{
        return migrateLegacyState(JSON.parse(el.textContent || "{{}}"));
      }} catch {{
        return {{}};
      }}
    }}

    function loadState() {{
      try {{
        const current = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{{}}");
        if (Object.keys(current).length) return migrateLegacyState(current);
        const legacy = JSON.parse(localStorage.getItem(LEGACY_STORAGE_KEY) || "{{}}");
        if (!Object.keys(legacy).length) return {{}};
        const migrated = migrateLegacyState(legacy);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated));
        return migrated;
      }} catch {{
        return {{}};
      }}
    }}

    function syncEmbeddedDom() {{
      const el = document.getElementById("embedded-state");
      if (el) el.textContent = JSON.stringify(state, null, 2);
    }}

    function setSaveStatus(msg) {{
      const el = document.getElementById("save-status");
      el.textContent = msg;
    }}

    let fileHandle = null;
    let fileWriteTimer = null;

    function openFileDb() {{
      return new Promise((resolve, reject) => {{
        const req = indexedDB.open(FILE_DB, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(FILE_STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      }});
    }}

    async function storeFileHandle(handle) {{
      const db = await openFileDb();
      await new Promise((resolve, reject) => {{
        const tx = db.transaction(FILE_STORE, "readwrite");
        tx.objectStore(FILE_STORE).put(handle, FILE_KEY);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      }});
    }}

    async function readStoredFileHandle() {{
      const db = await openFileDb();
      return new Promise((resolve, reject) => {{
        const tx = db.transaction(FILE_STORE, "readonly");
        const req = tx.objectStore(FILE_STORE).get(FILE_KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      }});
    }}

    async function ensureFilePermission(handle) {{
      if (!handle) return false;
      if ((await handle.queryPermission({{ mode: "readwrite" }})) === "granted") return true;
      return (await handle.requestPermission({{ mode: "readwrite" }})) === "granted";
    }}

    async function initFileHandle() {{
      if (!window.showOpenFilePicker) return;
      try {{
        const stored = await readStoredFileHandle();
        if (stored && await ensureFilePermission(stored)) {{
          fileHandle = stored;
          setSaveStatus("Linked to HTML file — saves go to browser + file");
        }}
      }} catch (err) {{
        console.warn("Could not restore file handle", err);
      }}
    }}

    async function linkHtmlFile() {{
      if (!window.showOpenFilePicker) {{
        alert("Linking the HTML file needs Chrome or Edge. localStorage still works in all browsers.");
        return;
      }}
      try {{
        const [handle] = await window.showOpenFilePicker({{
          types: [{{ description: "HTML", accept: {{ "text/html": [".html"] }} }}],
          multiple: false,
        }});
        if (!(await ensureFilePermission(handle))) return;
        fileHandle = handle;
        await storeFileHandle(handle);
        await writeStateToFile();
        setSaveStatus("Linked — saving to browser + file");
      }} catch (err) {{
        if (err.name !== "AbortError") console.error(err);
      }}
    }}

    function scheduleFileWrite() {{
      if (!fileHandle) return;
      clearTimeout(fileWriteTimer);
      fileWriteTimer = setTimeout(() => writeStateToFile(), 400);
    }}

    async function writeStateToFile() {{
      if (!fileHandle) return;
      try {{
        syncEmbeddedDom();
        const json = JSON.stringify(state, null, 2);
        const file = await fileHandle.getFile();
        let html = await file.text();
        const updated = html.replace(EMBEDDED_RE, `$1\\n${{json}}\\n$2`);
        if (updated === html) {{
          console.warn("embedded-state block not found in linked file");
          return;
        }}
        const writable = await fileHandle.createWritable();
        await writable.write(updated);
        await writable.close();
      }} catch (err) {{
        console.error("File save failed", err);
        setSaveStatus("Saved to browser (file write failed — click Link HTML file)");
      }}
    }}

    function saveState() {{
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      syncEmbeddedDom();
      scheduleFileWrite();
      setSaveStatus(fileHandle ? "Saved to browser + file" : "Saved to browser (link file to sync .html)");
      setTimeout(() => {{
        if (document.getElementById("save-status").textContent.startsWith("Saved")) {{
          setSaveStatus(fileHandle ? "Synced to file" : "");
        }}
      }}, 1500);
    }}

    let state = {{ ...loadEmbeddedState(), ...loadState() }};
    let filterMode = "all";

    function companyState(c) {{
      const key = companyKey(c);
      if (!state[key]) state[key] = {{ applied: false, date: "", role: "", status: "" }};
      return state[key];
    }}

    function todayIso() {{
      const d = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      return `${{d.getFullYear()}}-${{pad(d.getMonth() + 1)}}-${{pad(d.getDate())}}`;
    }}

    function linkify(url, label) {{
      if (!url) return "";
      const clean = url.replace(/^https?:\\/\\//, "");
      return `<a href="${{url}}" target="_blank" rel="noopener">${{label || clean}}</a>`;
    }}

    function escapeHtml(text) {{
      return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }}

    function formatMd(text) {{
      if (!text) return "";
      let s = escapeHtml(text);
      s = s.replace(/\\*\\*([^*]+)\\*\\*/g, "<strong>$1</strong>");
      s = s.replace(/\\[([^\\]]+)\\]\\(([^)]+)\\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
      return s;
    }}

    function render() {{
      const q = document.getElementById("search").value.trim().toLowerCase();
      const list = document.getElementById("list");
      list.innerHTML = "";
      let appliedCount = 0;
      let lastTier = "";

      COMPANIES.forEach((c) => {{
        const s = companyState(c);
        const stateKey = companyKey(c);
        if (c.applied && !s.applied && !state[stateKey]) {{
          s.applied = true;
        }}
        if (s.applied) appliedCount++;

        const hay = `${{c.name}} ${{c.roles}} ${{c.tier}}`.toLowerCase();
        if (q && !hay.includes(q)) return;
        if (filterMode === "applied" && !s.applied) return;
        if (filterMode === "pending" && s.applied) return;

        if (c.tier && c.tier !== lastTier) {{
          lastTier = c.tier;
          const h = document.createElement("div");
          h.className = "tier";
          h.textContent = c.tier;
          list.appendChild(h);
        }}

        const card = document.createElement("article");
        card.className = "card" + (s.applied ? " applied" : "");
        card.dataset.id = c.id;

        const links = [
          linkify(c.official, "Official"),
          linkify(c.allVacancies, "All jobs"),
          linkify(c.linkedin, "LinkedIn"),
        ].filter(Boolean).join("");

        card.innerHTML = `
          <div class="card-head">
            <input type="checkbox" id="cb-${{c.id}}" ${{s.applied ? "checked" : ""}} aria-label="Applied to ${{c.name}}" />
            <div style="flex:1">
              <div class="title-row">
                <span class="num">#${{c.id}}</span>
                <span class="name">${{c.name}}</span>
                <span class="saved" id="saved-${{c.id}}">Saved</span>
              </div>
              <div class="roles">${{c.roles || ""}}</div>
              ${{c.salary ? `<div class="meta"><strong>Salary fork:</strong> ${{formatMd(c.salary)}}</div>` : ""}}
              ${{c.workModel ? `<div class="meta"><strong>Work:</strong> ${{formatMd(c.workModel)}} · <strong>Min office:</strong> ${{formatMd(c.workDays)}}/week</div>` : ""}}
              <div class="links">${{links}}</div>
              <div class="meta">${{formatMd(c.sources || "")}}</div>
              ${{c.note ? `<div class="note">${{formatMd(c.note)}}</div>` : ""}}
              <div class="fields">
                <div>
                  <label for="date-${{c.id}}">Applied date</label>
                  <input type="date" id="date-${{c.id}}" value="${{s.date || ""}}" />
                </div>
                <div>
                  <label for="role-${{c.id}}">Role</label>
                  <input type="text" id="role-${{c.id}}" placeholder="Job title" value="${{s.role || ""}}" />
                </div>
                <div>
                  <label for="status-${{c.id}}">Status</label>
                  <input type="text" id="status-${{c.id}}" placeholder="e.g. No response" value="${{s.status || ""}}" />
                </div>
              </div>
            </div>
          </div>
        `;

        list.appendChild(card);

        const flash = () => {{
          const el = document.getElementById(`saved-${{c.id}}`);
          el.classList.add("show");
          setTimeout(() => el.classList.remove("show"), 900);
        }};

        const persist = () => {{
          saveState();
          flash();
          updateStats();
          card.classList.toggle("applied", s.applied);
        }};

        card.querySelector(`#cb-${{c.id}}`).addEventListener("change", (e) => {{
          s.applied = e.target.checked;
          if (e.target.checked) {{
            s.date = todayIso();
            const dateInput = card.querySelector(`#date-${{c.id}}`);
            if (dateInput) dateInput.value = s.date;
          }}
          persist();
        }});

        ["date", "role", "status"].forEach((field) => {{
          card.querySelector(`#${{field}}-${{c.id}}`).addEventListener("input", (e) => {{
            s[field] = e.target.value;
            persist();
          }});
        }});
      }});

      document.getElementById("stat-total").textContent = COMPANIES.length;
      document.getElementById("stat-applied").textContent = appliedCount;
      document.getElementById("stat-remaining").textContent = COMPANIES.length - appliedCount;
    }}

    function updateStats() {{
      const applied = COMPANIES.filter((c) => companyState(c).applied).length;
      document.getElementById("stat-applied").textContent = applied;
      document.getElementById("stat-remaining").textContent = COMPANIES.length - applied;
    }}

    document.getElementById("search").addEventListener("input", render);
    document.getElementById("filter-applied").addEventListener("click", () => {{ filterMode = "applied"; render(); }});
    document.getElementById("filter-pending").addEventListener("click", () => {{ filterMode = "pending"; render(); }});
    document.getElementById("filter-all").addEventListener("click", () => {{ filterMode = "all"; render(); }});
    function buildNameKeyedExport() {{
      const companies = {{}};
      for (const c of COMPANIES) {{
        companies[companyKey(c)] = {{ ...companyState(c) }};
      }}
      for (const [key, val] of Object.entries(state)) {{
        if (/^\\d+$/.test(key)) continue;
        if (!(key in companies) && val && typeof val === "object" && "applied" in val) {{
          companies[key] = {{ ...val }};
        }}
      }}
      return {{
        version: 2,
        keyBy: "companyName",
        exportedAt: new Date().toISOString(),
        companies,
      }};
    }}

    document.getElementById("export-json").addEventListener("click", () => {{
      const payload = buildNameKeyedExport();
      const blob = new Blob([JSON.stringify(payload, null, 2)], {{ type: "application/json" }});
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `nl-job-tracker-backup-${{payload.exportedAt.slice(0, 10)}}.json`;
      a.click();
      URL.revokeObjectURL(a.href);
    }});
    document.getElementById("clear-storage").addEventListener("click", () => {{
      if (confirm("Clear all saved checkboxes and notes from this browser?")) {{
        localStorage.removeItem(STORAGE_KEY);
        state = {{ ...loadEmbeddedState() }};
        syncEmbeddedDom();
        scheduleFileWrite();
        render();
      }}
    }});
    document.getElementById("link-file").addEventListener("click", linkHtmlFile);

    initFileHandle().then(() => {{
      syncEmbeddedDom();
      render();
    }});
  </script>
</body>
</html>
"""


def main() -> None:
    work = _load_work_arrangements()
    md_text = fix_tracker_links(MD_PATH.read_text(encoding="utf-8"))
    MD_PATH.write_text(md_text, encoding="utf-8")
    inject_work_into_tracker(MD_PATH, work)
    companies = parse_markdown(MD_PATH)
    embedded_state = _migrate_state_to_names(_load_embedded_state(HTML_PATH), companies)
    HTML_PATH.write_text(render_html(companies, embedded_state), encoding="utf-8")
    print(f"Wrote {HTML_PATH} ({len(companies)} companies)")


if __name__ == "__main__":
    main()
