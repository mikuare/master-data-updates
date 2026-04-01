from email.parser import BytesParser
from email.policy import default
from html import escape
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import quote
from urllib.parse import parse_qs, urlparse

import pandas as pd

from compare_employees import REPORT_FILENAMES, process_reports


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "web_uploads"
REPORT_FILES = list(REPORT_FILENAMES.values())
LAST_RUN_FILE = BASE_DIR / ".last_preview_run.txt"


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
    return pd.read_excel(path).fillna("")


def render_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "<p>No rows found.</p>"

    headers = "".join(f"<th>{escape(str(col))}</th>" for col in df.columns)
    rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{escape(str(value))}</td>" for value in row.tolist())
        rows.append(f"<tr>{cells}</tr>")
    body = "".join(rows)
    return f"<table><thead><tr>{headers}</tr></thead><tbody>{body}</tbody></table>"


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


def render_page(selected: str, message: str = "", error: str = "") -> str:
    report_dir = get_active_report_dir()
    links = []
    for filename in REPORT_FILES:
        label = filename.replace(".xlsx", "").replace("_", " ").title()
        current = ' class="active"' if filename == selected else ""
        links.append(f'<a{current} href="/?report={escape(filename)}">{escape(label)}</a>')

    df = load_report(selected, report_dir)
    path = latest_report_path(selected, report_dir)
    report_name = path.name if path else selected
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

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Employee Reports</title>
  <style>
    :root {{
      --bg: #f4efe6;
      --panel: #fffaf2;
      --ink: #1f2937;
      --muted: #6b7280;
      --line: #d6c7b2;
      --accent: #0f766e;
      --accent-soft: #d9f3ef;
      --good: #e8f7ef;
      --good-ink: #146c43;
      --bad: #fdecec;
      --bad-ink: #9f1239;
    }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: linear-gradient(180deg, #efe7d7 0%, var(--bg) 100%);
      color: var(--ink);
    }}
    .wrap {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 24px;
    }}
    .hero {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 20px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
    }}
    .panel {{
      background: white;
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 18px;
      margin-bottom: 18px;
    }}
    .upload-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 12px;
      align-items: end;
    }}
    label {{
      display: block;
      font-size: 14px;
      color: var(--muted);
      margin-bottom: 6px;
    }}
    input[type="file"] {{
      width: 100%;
      border: 1px solid var(--line);
      background: #fffdf8;
      border-radius: 10px;
      padding: 8px;
      box-sizing: border-box;
    }}
    button {{
      border: 0;
      border-radius: 10px;
      background: var(--accent);
      color: white;
      padding: 12px 16px;
      font-size: 14px;
      cursor: pointer;
    }}
    .nav {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin: 16px 0 20px;
    }}
    .nav a {{
      text-decoration: none;
      color: var(--ink);
      background: #f7f0e4;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 14px;
      font-size: 14px;
    }}
    .nav a.active {{
      background: var(--accent);
      color: white;
      border-color: var(--accent);
    }}
    .downloads {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 16px;
    }}
    .download-link {{
      text-decoration: none;
      color: var(--accent);
      background: white;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 14px;
      font-size: 14px;
    }}
    .meta {{
      color: var(--muted);
      font-size: 14px;
      margin-bottom: 16px;
    }}
    .notice {{
      border-radius: 12px;
      padding: 12px 14px;
      margin-bottom: 14px;
      font-size: 14px;
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
      background: white;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 16px;
    }}
    .card-label {{
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 6px;
    }}
    .card-value {{
      font-size: 28px;
      font-weight: 700;
      line-height: 1;
    }}
    .table-wrap {{
      overflow: auto;
      background: white;
      border: 1px solid var(--line);
      border-radius: 16px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }}
    th, td {{
      border-bottom: 1px solid #eee2cf;
      padding: 10px 12px;
      text-align: left;
      vertical-align: top;
      white-space: nowrap;
    }}
    th {{
      position: sticky;
      top: 0;
      background: var(--accent-soft);
    }}
    p {{
      padding: 16px;
      margin: 0;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <h1>Employee Report Preview</h1>
      <div class="panel">
        <form method="post" enctype="multipart/form-data">
          <div class="upload-grid">
            <div>
              <label for="system_file">System file</label>
              <input id="system_file" name="system_file" type="file" accept=".xls,.xlsx" required>
            </div>
            <div>
              <label for="hr_file">HR file</label>
              <input id="hr_file" name="hr_file" type="file" accept=".xls,.xlsx" required>
            </div>
            <div>
              <button type="submit">Upload And Compare</button>
            </div>
          </div>
        </form>
      </div>
      {message_html}
      {error_html}
      <div class="meta">Showing: {escape(report_name)} | Source: {escape(data_source)}</div>
      <div class="nav">{''.join(links)}</div>
      <div class="downloads">{''.join(download_links)}</div>
      <div class="summary">{render_summary(df)}</div>
      <div class="table-wrap">{render_table(df)}</div>
    </div>
  </div>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/download":
            self.handle_download(parsed)
            return

        query = parse_qs(parsed.query)
        selected = query.get("report", [REPORT_FILES[0]])[0]
        if selected not in REPORT_FILES:
            selected = REPORT_FILES[0]

        page = render_page(selected).encode("utf-8")
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
        selected = REPORT_FILES[0]
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
            page = render_page(selected, message=message).encode("utf-8")
            self.send_response(200)
        except Exception as exc:
            page = render_page(selected, error=str(exc)).encode("utf-8")
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
