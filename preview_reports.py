from email.parser import BytesParser
from email.policy import default
from html import escape
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import pandas as pd

from compare_employees import REPORT_FILENAMES, process_reports


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "web_uploads"
REPORT_FILES = list(REPORT_FILENAMES.values())
LAST_RUN_FILE = BASE_DIR / ".last_preview_run.txt"
APP_TABS = {
    "compare-employees": "Compare Employees",
    "trace-duplicate-stock-items": "Trace Duplicate Stock Items",
}


def get_active_report_dir() -> Path:
    if LAST_RUN_FILE.exists():
        saved = LAST_RUN_FILE.read_text(encoding="utf-8").strip()
        if saved:
            path = Path(saved)
            if path.exists():
                return path
    return BASE_DIR


def set_active_report_dir(path: Path) -> None:
    LAST_RUN_FILE.write_text(str(path), encoding="utf-8")


def latest_report_path(filename: str, report_dir: Path) -> Path | None:
    primary = report_dir / filename
    if primary.exists():
        return primary

    pattern = f"{primary.stem}_*{primary.suffix}"
    matches = sorted(
        report_dir.glob(pattern),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return matches[0] if matches else None


def load_report(filename: str, report_dir: Path) -> pd.DataFrame:
    path = latest_report_path(filename, report_dir)
    if path is None:
        return pd.DataFrame()
    df = pd.read_excel(path, dtype=str).fillna("")
    if "id" in df.columns:
        df["id"] = df["id"].astype(str).str.strip()
    return df


def render_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "<p>No rows found.</p>"

    headers = "".join(f"<th>{escape(str(col))}</th>" for col in df.columns)
    rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{escape(str(value))}</td>" for value in row.tolist())
        rows.append(f"<tr>{cells}</tr>")
    body = "".join(rows)
    return f'<table id="report-table" class="report-table"><thead><tr>{headers}</tr></thead><tbody>{body}</tbody></table>'


def render_summary(df: pd.DataFrame) -> str:
    total_rows = len(df.index)
    active_count = 0
    inactive_count = 0

    if "hr_status" in df.columns:
        statuses = df["hr_status"].astype(str).str.strip().str.lower()
        active_count = int((statuses == "active").sum())
        inactive_count = int((statuses != "active").sum())

    cards = [
        ("Total Rows", str(total_rows)),
        ("HR Active", str(active_count)),
        ("HR Not Active", str(inactive_count)),
    ]

    return "".join(
        f'<div class="card"><div class="card-label">{escape(label)}</div>'
        f'<div class="card-value">{escape(value)}</div></div>'
        for label, value in cards
    )


def parse_multipart(handler: BaseHTTPRequestHandler) -> dict[str, tuple[str, bytes]]:
    content_type = handler.headers.get("Content-Type", "")
    length = int(handler.headers.get("Content-Length", "0"))
    body = handler.rfile.read(length)

    raw_message = (
        f"Content-Type: {content_type}\r\n"
        f"MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8") + body
    message = BytesParser(policy=default).parsebytes(raw_message)

    files: dict[str, tuple[str, bytes]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        if not name or not filename:
            continue
        files[name] = (Path(filename).name, part.get_payload(decode=True) or b"")
    return files


def save_uploaded_file(upload_name: str, file_info: tuple[str, bytes], run_dir: Path) -> Path:
    filename, content = file_info
    suffix = Path(filename).suffix.lower()
    target_name = f"{upload_name}{suffix}"
    target = run_dir / target_name
    target.write_bytes(content)
    return target


def render_app_tabs(active_tab: str, selected_report: str) -> str:
    links = []
    for slug, label in APP_TABS.items():
        query = f"?tab={quote(slug)}&report={quote(selected_report)}"
        current = " app-tab-active" if slug == active_tab else ""
        links.append(f'<a class="app-tab{current}" href="{query}">{escape(label)}</a>')
    return "".join(links)


def render_report_tabs(selected_report: str) -> str:
    links = []
    for filename in REPORT_FILES:
        label = filename.replace(".xlsx", "").replace("_", " ").title()
        current = ' class="report-tab active"' if filename == selected_report else ' class="report-tab"'
        href = f"/?tab=compare-employees&report={quote(filename)}"
        links.append(f'<a{current} href="{href}">{escape(label)}</a>')
    return "".join(links)


def render_compare_panel(selected_report: str, message: str = "", error: str = "") -> str:
    report_dir = get_active_report_dir()
    df = load_report(selected_report, report_dir)
    path = latest_report_path(selected_report, report_dir)
    report_name = path.name if path else selected_report
    data_source = report_dir.name if report_dir != BASE_DIR else "project folder"

    download_links = []
    for filename in REPORT_FILES:
        report_path = latest_report_path(filename, report_dir)
        if report_path is None:
            continue
        label = filename.replace(".xlsx", "").replace("_", " ").title()
        href = f"/download?file={quote(report_path.name)}"
        download_links.append(
            f'<a class="download-link" href="{href}">Download {escape(label)}</a>'
        )

    message_html = f'<div class="notice success">{escape(message)}</div>' if message else ""
    error_html = f'<div class="notice error">{escape(error)}</div>' if error else ""

    return f"""
      <section class="hero">
        <div class="eyebrow">Local Workflow</div>
        <h1>Compare Employees</h1>
        <p class="lede">Upload a fresh system file and HR file, rerun the comparison, review the results, and download the generated reports from one place.</p>
        <div class="hero-pills">
          <span>Browser Upload</span>
          <span>Excel Preview</span>
          <span>Download Ready</span>
        </div>
      </section>
      <section class="panel upload-panel">
        <div class="panel-title-row">
          <div>
            <div class="panel-kicker">Run Comparison</div>
            <h2>Upload Source Files</h2>
          </div>
          <div class="source-chip">Source: {escape(data_source)}</div>
        </div>
        <form method="post" enctype="multipart/form-data">
          <input type="hidden" name="tab" value="compare-employees">
          <input type="hidden" name="report" value="{escape(selected_report)}">
          <div class="upload-grid">
            <div>
              <label for="system_file">System file</label>
              <input id="system_file" name="system_file" type="file" accept=".xls,.xlsx" required>
            </div>
            <div>
              <label for="hr_file">HR file</label>
              <input id="hr_file" name="hr_file" type="file" accept=".xls,.xlsx" required>
            </div>
            <div class="button-cell">
              <button type="submit">Upload And Compare</button>
            </div>
          </div>
        </form>
      </section>
      {message_html}
      {error_html}
      <section class="report-shell">
        <div class="panel-title-row">
          <div>
            <div class="panel-kicker">Results</div>
            <h2>{escape(report_name)}</h2>
          </div>
        </div>
        <div class="report-tabs">{render_report_tabs(selected_report)}</div>
        <div class="downloads">{''.join(download_links)}</div>
        <div class="toolbar">
          <div class="search-box">
            <input id="report-search" type="search" placeholder="Search employee ID, name, address, or status">
          </div>
          <div id="search-meta" class="search-meta">Showing all rows</div>
        </div>
        <div class="summary">{render_summary(df)}</div>
        <div class="table-wrap">{render_table(df)}</div>
      </section>
    """


def render_stock_placeholder(selected_report: str) -> str:
    return f"""
      <section class="hero">
        <div class="eyebrow">Soon Development</div>
        <h1>Trace Duplicate Stock Items</h1>
        <p class="lede">This tab is reserved for the next module. The plan is to upload stock files, detect repeated item codes or descriptions, and show duplicate clusters with source references.</p>
        <div class="hero-pills">
          <span>Duplicate Detection</span>
          <span>Stock Audit</span>
          <span>Next Module</span>
        </div>
      </section>
      <section class="panel roadmap-grid">
        <div class="roadmap-card">
          <div class="panel-kicker">Planned Scope</div>
          <h2>What This Module Will Do</h2>
          <ul>
            <li>Upload stock inventory files from local folders.</li>
            <li>Detect duplicate item codes, names, and near-matches.</li>
            <li>Group duplicates into a reviewable dashboard.</li>
            <li>Export duplicate findings into Excel reports.</li>
          </ul>
        </div>
        <div class="roadmap-card accent-card">
          <div class="panel-kicker">Status</div>
          <h2>Not Yet Active</h2>
          <p>The dashboard tab is ready, but the duplicate stock comparison logic has not been built yet.</p>
          <a class="ghost-link" href="/?tab=compare-employees&report={quote(selected_report)}">Go Back To Compare Employees</a>
        </div>
      </section>
    """


def render_page(
    active_tab: str,
    selected_report: str,
    message: str = "",
    error: str = "",
) -> str:
    if active_tab not in APP_TABS:
        active_tab = "compare-employees"
    if selected_report not in REPORT_FILES:
        selected_report = REPORT_FILES[0]

    if active_tab == "trace-duplicate-stock-items":
        content = render_stock_placeholder(selected_report)
    else:
        content = render_compare_panel(selected_report, message=message, error=error)

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Operations Dashboard</title>
  <style>
    :root {{
      --bg: #f6efe4;
      --paper: #fffaf2;
      --ink: #16202a;
      --muted: #6c7380;
      --line: #d8cab9;
      --accent: #0f766e;
      --accent-2: #c96f2d;
      --accent-soft: #ddf3ef;
      --sand: #f0e4d2;
      --good: #eaf8ef;
      --good-ink: #166534;
      --bad: #feeeee;
      --bad-ink: #9f1239;
      --shadow: 0 24px 70px rgba(73, 48, 18, 0.12);
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(201, 111, 45, 0.18), transparent 32%),
        radial-gradient(circle at top right, rgba(15, 118, 110, 0.18), transparent 28%),
        linear-gradient(180deg, #efe3d0 0%, var(--bg) 55%, #f8f3eb 100%);
      min-height: 100vh;
    }}
    .page-shell {{
      max-width: 1240px;
      margin: 0 auto;
      padding: 28px 20px 48px;
      animation: fade-up 500ms ease;
    }}
    .masthead {{
      display: grid;
      grid-template-columns: 1.2fr 0.8fr;
      gap: 18px;
      margin-bottom: 20px;
    }}
    .mast-card {{
      background: rgba(255, 250, 242, 0.8);
      backdrop-filter: blur(6px);
      border: 1px solid rgba(216, 202, 185, 0.9);
      border-radius: 22px;
      padding: 22px;
      box-shadow: var(--shadow);
    }}
    .brand-line {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      margin-bottom: 18px;
    }}
    .brand {{
      letter-spacing: 0.12em;
      text-transform: uppercase;
      font-size: 12px;
      color: var(--muted);
    }}
    .pulse {{
      width: 12px;
      height: 12px;
      border-radius: 999px;
      background: var(--accent);
      box-shadow: 0 0 0 rgba(15, 118, 110, 0.4);
      animation: pulse 1.8s infinite;
    }}
    .mast-card h1 {{
      margin: 0 0 10px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: clamp(32px, 4vw, 54px);
      line-height: 0.95;
    }}
    .mast-card p {{
      margin: 0;
      color: var(--muted);
      font-size: 16px;
      line-height: 1.5;
    }}
    .stat-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      height: 100%;
    }}
    .stat-chip {{
      background: linear-gradient(180deg, #fff, #f8f0e4);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      display: flex;
      flex-direction: column;
      justify-content: end;
      animation: float-in 700ms ease both;
    }}
    .stat-chip:nth-child(2) {{ animation-delay: 90ms; }}
    .stat-chip:nth-child(3) {{ animation-delay: 160ms; }}
    .stat-chip:nth-child(4) {{ animation-delay: 230ms; }}
    .stat-label {{
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .stat-value {{
      font-size: 22px;
      margin-top: 10px;
    }}
    .app-tabs {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin-bottom: 18px;
    }}
    .app-tab {{
      text-decoration: none;
      color: var(--ink);
      background: rgba(255, 250, 242, 0.84);
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 12px 18px;
      font-size: 15px;
      transition: transform 180ms ease, background 180ms ease;
    }}
    .app-tab:hover {{
      transform: translateY(-1px);
    }}
    .app-tab-active {{
      background: var(--accent);
      color: white;
      border-color: var(--accent);
    }}
    .hero, .panel, .report-shell {{
      background: rgba(255, 250, 242, 0.82);
      border: 1px solid rgba(216, 202, 185, 0.95);
      border-radius: 22px;
      padding: 22px;
      box-shadow: var(--shadow);
      margin-bottom: 18px;
      animation: fade-up 600ms ease both;
    }}
    .hero {{
      position: relative;
      overflow: hidden;
    }}
    .hero::after {{
      content: "";
      position: absolute;
      inset: auto -60px -60px auto;
      width: 220px;
      height: 220px;
      border-radius: 999px;
      background: radial-gradient(circle, rgba(15, 118, 110, 0.16), transparent 65%);
      pointer-events: none;
    }}
    .eyebrow, .panel-kicker {{
      color: var(--accent-2);
      text-transform: uppercase;
      letter-spacing: 0.12em;
      font-size: 12px;
      margin-bottom: 10px;
    }}
    .hero h1, .panel h2, .report-shell h2 {{
      margin: 0 0 10px;
      font-family: Georgia, "Times New Roman", serif;
    }}
    .lede {{
      margin: 0;
      max-width: 760px;
      color: var(--muted);
      line-height: 1.6;
      font-size: 16px;
    }}
    .hero-pills {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-top: 18px;
    }}
    .hero-pills span {{
      border: 1px solid var(--line);
      background: white;
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 13px;
    }}
    .panel-title-row {{
      display: flex;
      justify-content: space-between;
      align-items: start;
      gap: 14px;
      margin-bottom: 18px;
    }}
    .source-chip {{
      background: var(--sand);
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 12px;
      font-size: 13px;
      color: var(--muted);
    }}
    .upload-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
      align-items: end;
    }}
    label {{
      display: block;
      margin-bottom: 8px;
      color: var(--muted);
      font-size: 14px;
    }}
    input[type="file"] {{
      width: 100%;
      border: 1px solid var(--line);
      background: #fffdf8;
      border-radius: 12px;
      padding: 10px;
      font-family: inherit;
    }}
    button {{
      width: 100%;
      border: 0;
      border-radius: 14px;
      background: linear-gradient(135deg, var(--accent), #0b5c56);
      color: white;
      padding: 14px 16px;
      font-size: 15px;
      cursor: pointer;
      box-shadow: 0 14px 30px rgba(15, 118, 110, 0.18);
    }}
    .button-cell {{
      align-self: stretch;
    }}
    .report-tabs, .downloads {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 16px;
    }}
    .toolbar {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      align-items: center;
      margin-bottom: 16px;
    }}
    .search-box {{
      min-width: 260px;
      flex: 1 1 320px;
    }}
    .search-box input {{
      width: 100%;
      border: 1px solid var(--line);
      background: #fffdf8;
      border-radius: 12px;
      padding: 11px 14px;
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
      font-size: 14px;
    }}
    .search-meta {{
      color: var(--muted);
      font-size: 13px;
    }}
    .report-tab, .download-link, .ghost-link {{
      text-decoration: none;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 14px;
      font-size: 14px;
      background: white;
      color: var(--ink);
    }}
    .report-tab.active {{
      background: var(--accent-soft);
      border-color: rgba(15, 118, 110, 0.25);
    }}
    .download-link {{
      color: var(--accent);
    }}
    .ghost-link {{
      display: inline-block;
      margin-top: 14px;
    }}
    .notice {{
      border-radius: 16px;
      padding: 14px 16px;
      margin-bottom: 14px;
      font-size: 14px;
      box-shadow: var(--shadow);
    }}
    .notice.success {{
      background: var(--good);
      color: var(--good-ink);
    }}
    .notice.error {{
      background: var(--bad);
      color: var(--bad-ink);
    }}
    .summary {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }}
    .card {{
      background: linear-gradient(180deg, #fff, #faf3e9);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      animation: float-in 700ms ease both;
    }}
    .card-label {{
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 8px;
    }}
    .card-value {{
      font-size: 30px;
      font-weight: 700;
      line-height: 1;
    }}
    .table-wrap {{
      overflow: auto;
      background: white;
      border: 1px solid var(--line);
      border-radius: 18px;
      max-height: 520px;
    }}
    .report-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
    }}
    .report-table th, .report-table td {{
      border-bottom: 1px solid #eee2cf;
      padding: 10px 12px;
      text-align: left;
      vertical-align: top;
      white-space: nowrap;
    }}
    .report-table th {{
      position: sticky;
      top: 0;
      background: var(--accent-soft);
      font-weight: 600;
      letter-spacing: 0.01em;
    }}
    .report-table td:first-child, .report-table th:first-child {{
      font-family: "Consolas", "Courier New", monospace;
    }}
    p {{
      margin: 0;
    }}
    .roadmap-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 18px;
    }}
    .roadmap-card {{
      background: white;
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 20px;
    }}
    .accent-card {{
      background: linear-gradient(180deg, #fff5ea, #fff);
    }}
    .roadmap-card ul {{
      margin: 0;
      padding-left: 18px;
      color: var(--muted);
      line-height: 1.7;
    }}
    @keyframes fade-up {{
      from {{ opacity: 0; transform: translateY(18px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    @keyframes float-in {{
      from {{ opacity: 0; transform: translateY(14px) scale(0.98); }}
      to {{ opacity: 1; transform: translateY(0) scale(1); }}
    }}
    @keyframes pulse {{
      0% {{ box-shadow: 0 0 0 0 rgba(15, 118, 110, 0.4); }}
      70% {{ box-shadow: 0 0 0 14px rgba(15, 118, 110, 0); }}
      100% {{ box-shadow: 0 0 0 0 rgba(15, 118, 110, 0); }}
    }}
    @media (max-width: 900px) {{
      .masthead {{
        grid-template-columns: 1fr;
      }}
      .panel-title-row {{
        flex-direction: column;
      }}
    }}
  </style>
</head>
<body>
  <main class="page-shell">
    <section class="masthead">
      <div class="mast-card">
        <div class="brand-line">
          <div class="brand">Operations Dashboard</div>
          <div class="pulse"></div>
        </div>
        <h1>Flexible local tools for operations review.</h1>
        <p>Start with employee comparison today, then extend the same dashboard into stock audit and duplicate tracing next.</p>
      </div>
      <div class="stat-grid">
        <div class="stat-chip">
          <div class="stat-label">Current Module</div>
          <div class="stat-value">Employee Comparison</div>
        </div>
        <div class="stat-chip">
          <div class="stat-label">Preview Mode</div>
          <div class="stat-value">Local Browser App</div>
        </div>
        <div class="stat-chip">
          <div class="stat-label">Next Module</div>
          <div class="stat-value">Duplicate Stock Trace</div>
        </div>
        <div class="stat-chip">
          <div class="stat-label">File Support</div>
          <div class="stat-value">XLS And XLSX</div>
        </div>
      </div>
    </section>
    <nav class="app-tabs">{render_app_tabs(active_tab, selected_report)}</nav>
    {content}
  </main>
<script>
  (function () {{
    const input = document.getElementById("report-search");
    const table = document.getElementById("report-table");
    const meta = document.getElementById("search-meta");
    if (!input || !table || !meta) {{
      return;
    }}

    const rows = Array.from(table.querySelectorAll("tbody tr"));
    const update = () => {{
      const query = input.value.trim().toLowerCase();
      let visible = 0;

      rows.forEach((row) => {{
        const text = row.textContent.toLowerCase();
        const match = !query || text.includes(query);
        row.style.display = match ? "" : "none";
        if (match) {{
          visible += 1;
        }}
      }});

      meta.textContent = query
        ? "Showing " + visible + " matching row(s)"
        : "Showing all " + rows.length + " row(s)";
    }};

    input.addEventListener("input", update);
    update();
  }})();
</script>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/download":
            self.handle_download(parsed)
            return

        query = parse_qs(parsed.query)
        selected_report = query.get("report", [REPORT_FILES[0]])[0]
        active_tab = query.get("tab", ["compare-employees"])[0]

        page = render_page(active_tab, selected_report).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(page)))
        self.end_headers()
        self.wfile.write(page)

    def handle_download(self, parsed) -> None:
        report_dir = get_active_report_dir()
        query = parse_qs(parsed.query)
        requested = query.get("file", [""])[0]

        file_path = report_dir / Path(requested).name
        if not file_path.exists():
            self.send_response(404)
            self.end_headers()
            return

        content = file_path.read_bytes()
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{file_path.name}"',
        )
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self) -> None:
        active_tab = "compare-employees"
        selected_report = REPORT_FILES[0]
        try:
            files = parse_multipart(self)
            if "system_file" not in files or "hr_file" not in files:
                raise ValueError("Upload both the system file and the HR file.")

            run_dir = UPLOAD_DIR / f"run_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}"
            run_dir.mkdir(parents=True, exist_ok=True)

            system_path = save_uploaded_file("system_clean", files["system_file"], run_dir)
            hr_path = save_uploaded_file("hr_clean", files["hr_file"], run_dir)
            result = process_reports(system_path=system_path, hr_path=hr_path, output_dir=run_dir)
            set_active_report_dir(run_dir)

            message = (
                f"Compared {result['system_file']} and {result['hr_file']}. "
                f"Matched: {result['matched_count']}, "
                f"Inactive updates: {result['inactive_count']}, "
                f"Active updates: {result['active_count']}, "
                f"New active: {result['new_active_count']}."
            )
            page = render_page(active_tab, selected_report, message=message).encode("utf-8")
            self.send_response(200)
        except Exception as exc:
            page = render_page(active_tab, selected_report, error=str(exc)).encode("utf-8")
            self.send_response(400)

        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(page)))
        self.end_headers()
        self.wfile.write(page)

    def log_message(self, format: str, *args) -> None:
        return


if __name__ == "__main__":
    host = "127.0.0.1"
    port = 8000
    server = HTTPServer((host, port), ReportHandler)
    print(f"Preview server running at http://{host}:{port}")
    server.serve_forever()
