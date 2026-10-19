"""
JobPulse - Universal Desktop Excel Sync Bridge
Supports: macOS, Windows, Linux
Author: Muhammad Irfan Firdaus (irfnfrdaus)
License: MIT
"""

import json
import os
import platform
import subprocess
import sys
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

PORT = 5252
OS_SYSTEM = platform.system()

# Determine user's default Documents directory cross-platform
if OS_SYSTEM == "Windows":
    DOCS_DIR = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Documents", "Job_Tracker")
else:
    DOCS_DIR = os.path.join(os.path.expanduser("~"), "Documents", "Job_Tracker")

os.makedirs(DOCS_DIR, exist_ok=True)
PRIMARY_EXCEL_PATH = os.path.join(DOCS_DIR, "Job_Applications_Tracker.xlsx")

STATUS_COLORS = {
    "Applied": "E0F2FE",        # Light Sky
    "Under Review": "FEF3C7",   # Light Yellow
    "Assessment / Test": "EDE9FE", # Light Purple
    "Interviewing": "FDE68A",   # Warm Amber
    "Shortlisted": "D1FAE5",    # Soft Green
    "Offer Received": "86EFAC", # Vibrant Mint
    "Rejected": "FEE2E2",       # Soft Red
    "Wishlist / Target": "F1F5F9" # Light Slate
}

THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

def init_excel_if_missing():
    """Initializes a new Excel workbook if one does not already exist."""
    if os.path.exists(PRIMARY_EXCEL_PATH):
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Applications"

    headers = [
        "App ID", "Date Applied", "Company Name", "Job Title", "Job Type",
        "Location", "Platform / Source", "Application Status", "Job Posting URL",
        "HR / Contact Email", "Expected Salary", "Follow-Up Date", "Notes / Tech Stack"
    ]

    ws.append(headers)

    header_fill = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 26

    widths = {
        "A": 12, "B": 14, "C": 26, "D": 32, "E": 20, "F": 22,
        "G": 18, "H": 20, "I": 35, "J": 26, "K": 18, "L": 14, "M": 40
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(PRIMARY_EXCEL_PATH)
    wb.close()
    print(f"Created new master Excel spreadsheet at: {PRIMARY_EXCEL_PATH}")

def launch_excel():
    """Cross-platform launcher for opening the master Excel file."""
    try:
        if OS_SYSTEM == "Darwin":
            # macOS
            subprocess.Popen(["open", PRIMARY_EXCEL_PATH])
            return True, "Launched file via macOS open"
        elif OS_SYSTEM == "Windows":
            # Windows Explorer shell
            subprocess.Popen(["explorer.exe", PRIMARY_EXCEL_PATH])
            return True, "Launched file via Windows Shell"
        else:
            # Linux (xdg-open)
            subprocess.Popen(["xdg-open", PRIMARY_EXCEL_PATH])
            return True, "Launched file via Linux xdg-open"
    except Exception as e:
        return False, f"Launch failed: {e}"

def get_tracker_html():
    """Generates the live web view table from the master Excel spreadsheet."""
    rows = []
    if os.path.exists(PRIMARY_EXCEL_PATH):
        try:
            wb = openpyxl.load_workbook(PRIMARY_EXCEL_PATH, data_only=True)
            ws = wb.active
            for r in range(1, ws.max_row + 1):
                row_vals = [str(ws.cell(row=r, column=c).value or '') for c in range(1, ws.max_column + 1)]
                rows.append(row_vals)
            wb.close()
        except Exception as e:
            rows = [[f"Error loading Excel: {e}"]]

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>JobPulse | Master Applications Tracker</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', sans-serif; }}
    body {{ background: #0B1120; color: #F1F5F9; padding: 24px; font-size: 13px; }}
    .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid #1E293B; padding-bottom: 16px; }}
    h1 {{ font-size: 20px; font-weight: 800; color: #FFFFFF; display: flex; align-items: center; gap: 10px; }}
    .badge {{ background: #1E3A8A; color: #93C5FD; font-size: 11px; padding: 3px 8px; border-radius: 12px; font-weight: 600; }}
    .btn {{ background: #10B981; color: white; padding: 8px 14px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 6px; cursor: pointer; border: none; }}
    .btn:hover {{ background: #059669; }}
    .btn-secondary {{ background: #1E293B; color: #CBD5E1; }}
    .btn-secondary:hover {{ background: #334155; }}
    .card {{ background: #0F172A; border: 1px solid #1E293B; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.3); }}
    table {{ width: 100%; border-collapse: collapse; text-align: left; }}
    th {{ background: #0F2C59; color: #FFFFFF; font-weight: 700; padding: 10px 14px; text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.5px; border-bottom: 1px solid #1E293B; }}
    td {{ padding: 9px 14px; border-bottom: 1px solid #1E293B; font-size: 12px; color: #CBD5E1; }}
    tr:hover td {{ background: #1E293B80; }}
    .status-pill {{ padding: 2px 8px; border-radius: 12px; font-size: 10.5px; font-weight: 700; display: inline-block; }}
    .status-Applied {{ background: #0284C720; color: #38BDF8; border: 1px solid #0284C740; }}
    .status-Under-Review {{ background: #D9770620; color: #FBBF24; border: 1px solid #D9770640; }}
    .status-Interviewing {{ background: #D9770620; color: #FBBF24; border: 1px solid #D9770640; }}
    .status-Offer-Received {{ background: #10B98120; color: #34D399; border: 1px solid #10B98140; }}
    .status-Wishlist {{ background: #64748B20; color: #94A3B8; border: 1px solid #64748B40; }}
    .status-Rejected {{ background: #E11D4820; color: #FB7185; border: 1px solid #E11D4840; }}
    .mono {{ font-family: 'JetBrains Mono', monospace; font-size: 11.5px; }}
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1><span>📊 JobPulse Tracker</span> <span class="badge">Desktop Excel Sync</span></h1>
      <p style="color: #64748B; font-size: 12px; margin-top: 4px;">{PRIMARY_EXCEL_PATH}</p>
    </div>
    <div style="display: flex; gap: 8px;">
      <button onclick="launchExcel(this)" class="btn">📊 Open in Desktop Excel</button>
      <button onclick="location.reload()" class="btn btn-secondary">🔄 Refresh</button>
    </div>
  </div>

  <div id="launchToast" style="display: none; padding: 10px 14px; background: #064E3B; border: 1px solid #059669; color: #A7F3D0; border-radius: 8px; margin-bottom: 16px; font-weight: 600; font-size: 12px;"></div>

  <div class="card">
    <table>
      <thead>
        <tr>"""
    if rows:
        headers = rows[0]
        for h in headers[:9]:
            html += f"<th>{h}</th>"
        html += "<th>Salary</th><th>Notes</th></tr></thead><tbody>"

        for row in rows[1:]:
            html += "<tr>"
            # App ID
            html += f"<td class='mono' style='color: #60A5FA; font-weight: bold;'>{row[0] if len(row) > 0 else ''}</td>"
            # Date
            html += f"<td class='mono' style='color: #94A3B8;'>{row[1] if len(row) > 1 else ''}</td>"
            # Company
            html += f"<td style='color: #FFFFFF; font-weight: 600;'>{row[2] if len(row) > 2 else ''}</td>"
            # Job Title
            html += f"<td style='color: #E2E8F0;'>{row[3] if len(row) > 3 else ''}</td>"
            # Job Type
            html += f"<td style='color: #94A3B8;'>{row[4] if len(row) > 4 else ''}</td>"
            # Location
            html += f"<td style='color: #94A3B8;'>{row[5] if len(row) > 5 else ''}</td>"
            # Platform
            html += f"<td style='color: #CBD5E1;'>{row[6] if len(row) > 6 else ''}</td>"
            # Status
            status_text = row[7] if len(row) > 7 else 'Applied'
            status_class = f"status-{status_text.replace(' ', '-').replace('/', '-')}"
            html += f"<td><span class='status-pill {status_class}'>{status_text}</span></td>"
            # URL
            url_text = row[8] if len(row) > 8 else ''
            if url_text.startswith("http"):
                html += f"<td><a href='{url_text}' target='_blank' style='color: #38BDF8;'>Link ↗</a></td>"
            else:
                html += f"<td style='color: #64748B;'>-</td>"
            # Salary
            sal = row[10] if len(row) > 10 else ''
            html += f"<td class='mono' style='color: #34D399; font-weight: 600;'>{sal}</td>"
            # Notes
            notes = row[12] if len(row) > 12 else ''
            html += f"<td style='color: #94A3B8; max-width: 320px; font-size: 11px;'>{notes}</td>"
            html += "</tr>"
        html += """</tbody></table></div>
  <script>
    async function launchExcel(btn) {
      const orig = btn.innerHTML;
      btn.innerHTML = "⏳ Opening...";
      btn.disabled = true;
      const toast = document.getElementById("launchToast");
      try {
        const res = await fetch("/open-excel");
        const data = await res.json();
        toast.style.display = "block";
        toast.innerHTML = "✅ Excel command sent to your desktop OS!";
        setTimeout(() => { toast.style.display = "none"; }, 6000);
      } catch (err) {
        toast.style.display = "block";
        toast.style.background = "#7F1D1D";
        toast.innerText = "Error contacting bridge: " + err;
      } finally {
        btn.innerHTML = orig;
        btn.disabled = false;
      }
    }
  </script>
</body></html>"""
    return html

def append_job_to_excel(job):
    init_excel_if_missing()

    wb = openpyxl.load_workbook(PRIMARY_EXCEL_PATH)
    ws = wb.active

    max_id_num = 0
    for row in range(2, ws.max_row + 1):
        cell_val = str(ws.cell(row=row, column=1).value or "")
        if cell_val.startswith("APP-"):
            try:
                num = int(cell_val.replace("APP-", ""))
                if num > max_id_num:
                    max_id_num = num
            except ValueError:
                pass
    next_id = f"APP-{max_id_num + 1:03d}"

    date_applied = job.get("date_applied") or datetime.now().strftime("%Y-%m-%d")
    company = job.get("company", "").strip() or "N/A"
    job_title = job.get("job_title", "").strip() or "N/A"
    job_type = job.get("job_type", "").strip() or "Full-time"
    location = job.get("location", "").strip() or "Remote / Hybrid"
    platform_src = job.get("platform", "").strip() or "Direct"
    status = job.get("status", "").strip() or "Applied"
    url = job.get("url", "").strip()
    email = job.get("email", "").strip()
    salary = job.get("salary", "").strip()
    follow_up = job.get("follow_up", "").strip() or datetime.now().strftime("%Y-%m-%d")
    notes = job.get("notes", "").strip()

    new_row = [
        next_id,
        date_applied,
        company,
        job_title,
        job_type,
        location,
        platform_src,
        status,
        url,
        email,
        salary,
        follow_up,
        notes
    ]

    ws.append(new_row)
    new_row_idx = ws.max_row
    ws.row_dimensions[new_row_idx].height = 22

    status_bg = STATUS_COLORS.get(status, "FFFFFF")

    for col_idx in range(1, len(new_row) + 1):
        cell = ws.cell(row=new_row_idx, column=col_idx)
        cell.font = Font(name="Calibri", size=9.5)
        cell.border = THIN_BORDER
        
        if col_idx in [1, 2, 5, 8, 12]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in [11]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        if col_idx == 8:
            cell.fill = PatternFill(start_color=status_bg, end_color=status_bg, fill_type="solid")
            cell.font = Font(name="Calibri", size=9.5, bold=True)

    try:
        wb.save(PRIMARY_EXCEL_PATH)
        wb.close()
    except PermissionError:
        raise PermissionError("Your Excel file is currently open in another app. Please close it and retry.")

    return next_id, new_row_idx - 1

class TrackerHandler(BaseHTTPRequestHandler):
    def _send_cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(200)
        self._send_cors_headers()
        self.end_headers()

    def do_GET(self):
        if self.path == "/status" or self.path == "/":
            total_records = 0
            if os.path.exists(PRIMARY_EXCEL_PATH):
                try:
                    wb = openpyxl.load_workbook(PRIMARY_EXCEL_PATH, read_only=True)
                    ws = wb.active
                    total_records = max(0, ws.max_row - 1)
                    wb.close()
                except Exception:
                    pass

            response_data = {
                "status": "connected",
                "os": OS_SYSTEM,
                "excel_path": PRIMARY_EXCEL_PATH,
                "total_records": total_records,
                "message": "JobPulse Sync Bridge is Active"
            }
            self.send_response(200)
            self._send_cors_headers()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(response_data).encode("utf-8"))

        elif self.path == "/view":
            html_content = get_tracker_html()
            self.send_response(200)
            self._send_cors_headers()
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(html_content.encode("utf-8"))

        elif self.path == "/open-excel":
            success, msg = launch_excel()
            self.send_response(200 if success else 500)
            self._send_cors_headers()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": msg}).encode("utf-8"))

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/add-job":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                job_data = json.loads(post_data.decode("utf-8"))
                app_id, total = append_job_to_excel(job_data)
                
                response_data = {
                    "success": True,
                    "app_id": app_id,
                    "total_records": total,
                    "message": f"Successfully logged {app_id} to {PRIMARY_EXCEL_PATH}!"
                }
                self.send_response(200)
                self._send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(response_data).encode("utf-8"))
            except PermissionError as pe:
                self.send_response(423)
                self._send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(pe)}).encode("utf-8"))
            except Exception as e:
                self.send_response(500)
                self._send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode("utf-8"))

        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        return

def run_server():
    init_excel_if_missing()
    server_address = ("127.0.0.1", PORT)
    httpd = HTTPServer(server_address, TrackerHandler)
    print("=" * 60)
    print(f"   JobPulse Excel Sync Bridge ({OS_SYSTEM})")
    print(f"   Target Excel: {PRIMARY_EXCEL_PATH}")
    print(f"   Dashboard:    http://127.0.0.1:{PORT}/view")
    print("=" * 60)
    print("Bridge is running. Keep this terminal window open.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping bridge...")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
