#!/usr/bin/env python3
"""Generate companies-transfer-analysis.html from companies-transfer-analysis.md."""

import html
import re
from pathlib import Path

ROOT = Path(__file__).parent
MD_PATH = ROOT / "companies-transfer-analysis.md"
HTML_PATH = ROOT / "companies-transfer-analysis.html"

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{title}</title>
  <style>
    :root {{
      color-scheme: light dark;
      --bg: #ffffff;
      --text: #1f2328;
      --border: #d0d7de;
      --head: #f6f8fa;
      --muted: #57606a;
    }}
    @media (prefers-color-scheme: dark) {{
      :root {{
        --bg: #0d1117;
        --text: #e6edf3;
        --border: #30363d;
        --head: #161b22;
        --muted: #8b949e;
      }}
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      padding: 1.5rem 2rem 3rem;
      font: 14px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
      color: var(--text);
      background: var(--bg);
    }}
    h1 {{
      margin: 0 0 1rem;
      font-size: 1.75rem;
      line-height: 1.25;
    }}
    p {{
      margin: 0 0 1rem;
      max-width: 72rem;
    }}
    em {{ color: var(--muted); }}
    .table-wrap {{
      width: 100%;
      overflow-x: auto;
      margin-top: 1.25rem;
    }}
    table {{
      width: 100%;
      min-width: 1100px;
      border-collapse: collapse;
      font-size: 13px;
    }}
    th, td {{
      border: 1px solid var(--border);
      padding: 10px 12px;
      vertical-align: top;
      text-align: left;
    }}
    th {{
      background: var(--head);
      font-weight: 600;
      white-space: nowrap;
    }}
    tr:nth-child(even) td {{ background: color-mix(in srgb, var(--head) 35%, transparent); }}
    footer {{
      margin-top: 2rem;
      color: var(--muted);
      font-size: 12px;
    }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  {intro}
  {table}
  <footer>Regenerate: <code>python3 build_companies_transfer_html.py</code></footer>
</body>
</html>
"""


def inline_md(text: str) -> str:
    text = html.escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    return text.replace("&lt;br&gt;", "<br>")


def parse_markdown(path: Path) -> tuple[str, str, str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    title = ""
    intro_parts: list[str] = []
    table_rows: list[list[str]] = []

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith("# "):
            title = line[2:].strip()
            i += 1
            continue
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if all(set(c) <= {"-", ":", " "} for c in cells):
                i += 1
                continue
            table_rows.append(cells)
            i += 1
            continue
        intro_parts.append(line)
        i += 1

    intro_html = ""
    for part in intro_parts:
        if part.startswith("*") and part.endswith("*"):
            intro_html += f"<p><em>{inline_md(part.strip('*'))}</em></p>\n"
        else:
            intro_html += f"<p>{inline_md(part)}</p>\n"

    if not table_rows:
        table_html = ""
    else:
        header, *body = table_rows
        thead = "<thead><tr>" + "".join(f"<th>{inline_md(cell)}</th>" for cell in header) + "</tr></thead>"
        tbody_rows = []
        for row in body:
            tbody_rows.append("<tr>" + "".join(f"<td>{inline_md(cell)}</td>" for cell in row) + "</tr>")
        tbody = "<tbody>" + "".join(tbody_rows) + "</tbody>"
        table_html = f'<div class="table-wrap"><table>{thead}{tbody}</table></div>'

    return title, intro_html, table_html


def main() -> None:
    title, intro, table = parse_markdown(MD_PATH)
    HTML_PATH.write_text(
        HTML_TEMPLATE.format(title=html.escape(title), intro=intro, table=table),
        encoding="utf-8",
    )
    print(f"Wrote {HTML_PATH}")


if __name__ == "__main__":
    main()
