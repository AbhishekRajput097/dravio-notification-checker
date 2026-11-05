"""
================================================================
Dravio — NEET UG Counselling Notification Monitor
================================================================
Version     : 3.0.0
Description : Monitors 6 official NEET UG counselling sites for
              new notifications. Tracks individual notification
              links — sends email alerts with exact titles & PDF
              URLs, and logs each notification as a structured
              row in Google Sheets for the frontend team.
Sites       : Gujarat (ACPUGMEC), Maharashtra (CET Cell),
              MP (DME), Rajasthan (rajugneet2026), MCC (AIQ),
              NTA NEET UG
================================================================
"""

import os
import re
import sys
import json
import time
import hashlib
import smtplib
import logging
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from urllib.parse import urljoin

import requests
import schedule
import gspread
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------
# Load .env file
# ----------------------------------------------------------------
load_dotenv()

# ----------------------------------------------------------------
# Logging Setup
# ----------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%d-%m-%Y %I:%M:%S %p",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("dravio_monitor.log", encoding="utf-8")
    ]
)
log = logging.getLogger("DravioMonitor")

# ----------------------------------------------------------------
# Configuration from .env
# ----------------------------------------------------------------
GMAIL_SENDER           = os.getenv("GMAIL_SENDER", "")
GMAIL_APP_PASSWORD     = os.getenv("GMAIL_APP_PASSWORD", "")
ALERT_RECIPIENTS       = [e.strip() for e in os.getenv("ALERT_RECIPIENTS", "").split(",") if e.strip()]
GOOGLE_SHEET_ID        = os.getenv("GOOGLE_SHEET_ID", "")
CREDENTIALS_FILE       = os.getenv("CREDENTIALS_FILE", "credentials.json")
WORKSHEET_NAME         = os.getenv("WORKSHEET_NAME", "Notification Log")
CHECK_INTERVAL_MINUTES = int(os.getenv("CHECK_INTERVAL_MINUTES", "60"))
MAX_RETRIES            = int(os.getenv("MAX_RETRIES", "3"))
RESTART_WAIT_MINUTES   = int(os.getenv("RESTART_WAIT_MINUTES", "5"))

STATE_FILE          = "last_state.json"
STATUS_SHEET_NAME   = "Monitor Status"
HEARTBEAT_TIME      = os.getenv("HEARTBEAT_TIME", "09:00")   # daily summary email time
SCRIPT_START_TIME   = datetime.now()                          # set once at startup

# ----------------------------------------------------------------
# Sites Configuration
# ----------------------------------------------------------------
SITES = [
    {
        "name"       : "Gujarat (ACPUGMEC)",
        "state"      : "Gujarat",
        "url"        : "https://medadmgujarat.org/ug/home.aspx",
        "filter_kw"  : None,
        "strip_noise": [],
    },
    {
        "name"       : "Maharashtra (CET Cell)",
        "state"      : "Maharashtra",
        "url"        : "https://cetcell.mahacet.org/",
        # Only track links whose text contains medical/health keywords
        "filter_kw"  : ["medical", "health science", "mbbs", "neet", "nursing",
                         "bds", "ayush", "dmer", "health", "dental"],
        "strip_noise": [],
    },
    {
        "name"       : "Madhya Pradesh (DME)",
        "state"      : "Madhya Pradesh",
        "url"        : "https://dme.mponline.gov.in/Portal/Services/DMEMP/DMEUG/Profile/Instructions.aspx",
        "filter_kw"  : None,
        # Strip rotating load-balancer server names that change on every request
        "strip_noise": [r"^ARYABHATA$", r"^VINDHYACHAL$", r"^ARYABHAT$",
                        r"Powered by\s*:", r"^[A-Z]{5,15}$"],
    },
    {
        "name"       : "Rajasthan (RUHS)",
        "state"      : "Rajasthan",
        "url"        : "https://www.ruhsraj.org/",
        # RUHS = Rajasthan University of Health Sciences — official body for
        # Rajasthan NEET UG / MBBS / BDS admissions and counselling.
        # Previous URL (rajugneet2026.com) was a private blog, not official.
        "filter_kw"  : None,
        "strip_noise": [],
    },
    {
        "name"       : "MCC (AIQ / Central Counselling)",
        "state"      : "All India",
        "url"        : "https://mcc.nic.in/current-events-ug/",
        "filter_kw"  : None,
        "strip_noise": [],
    },
    {
        "name"       : "NTA NEET UG (Public Notices)",
        "state"      : "National (NTA)",
        "url"        : "https://neet.nta.nic.in/public-notices/",
        "filter_kw"  : None,
        "strip_noise": [],
    },
]

# ----------------------------------------------------------------
# Google Sheets Setup
# Columns designed as a handoff document for the frontend team:
# each row = one specific notification with its direct link.
# ----------------------------------------------------------------
SHEET_HEADERS = [
    "Sr No",
    "Detected Date",
    "Detected Time (IST)",
    "State / Body",
    "Site Name",
    "Notification Title",
    "Direct Link / PDF URL",
    "Source Page URL",
    "Status",
    "Remarks",
]


def get_worksheet():
    """Connect to Google Sheets and return the target worksheet."""
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sheet  = client.open_by_key(GOOGLE_SHEET_ID)

        try:
            ws = sheet.worksheet(WORKSHEET_NAME)
            # Ensure headers are up-to-date (in case of old worksheet)
            existing_headers = ws.row_values(1)
            if existing_headers != SHEET_HEADERS:
                ws.update("A1", [SHEET_HEADERS])
                log.info("✅ Updated Google Sheet headers.")
        except gspread.WorksheetNotFound:
            ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=12)
            ws.append_row(SHEET_HEADERS)
            ws.format("A1:J1", {
                "backgroundColor" : {"red": 0.1, "green": 0.27, "blue": 0.43},
                "textFormat"      : {
                    "bold"            : True,
                    "foregroundColor" : {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "fontSize"        : 11
                },
                "horizontalAlignment": "CENTER"
            })
            log.info("✅ Created new Google Sheet worksheet.")
        return ws

    except Exception as e:
        log.error(f"❌ Google Sheets connection failed: {e}")
        return None


STATUS_SHEET_HEADERS = [
    "Site Name", "State", "Last Checked (IST)", "HTTP Status",
    "Notifications Indexed", "Last Result", "Last Error",
]


def get_status_worksheet():
    """Get or create the Monitor Status worksheet (separate tab in same sheet)."""
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sheet  = client.open_by_key(GOOGLE_SHEET_ID)

        try:
            ws = sheet.worksheet(STATUS_SHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sheet.add_worksheet(title=STATUS_SHEET_NAME, rows=20, cols=10)
            ws.append_row(STATUS_SHEET_HEADERS)
            ws.format("A1:G1", {
                "backgroundColor" : {"red": 0.13, "green": 0.37, "blue": 0.23},
                "textFormat"      : {
                    "bold"            : True,
                    "foregroundColor" : {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "fontSize"        : 11
                },
                "horizontalAlignment": "CENTER"
            })
            # Pre-fill one row per site so they exist to be updated in-place
            for i, site in enumerate(SITES, start=2):
                ws.update(f"A{i}:G{i}", [[
                    site["name"], site["state"], "Not checked yet",
                    "—", "—", "—", "—"
                ]])
            log.info("✅ Created Monitor Status worksheet.")
        return ws

    except Exception as e:
        log.error(f"❌ Could not get status worksheet: {e}")
        return None


def update_status_sheet(ws_status, site_results):
    """
    Overwrite the per-site rows in the Monitor Status sheet.
    site_results = list of dicts: {name, state, checked_at, http_ok,
                                    links_count, result, error}
    Rows 2..N are fixed — one per site in SITES order.
    """
    if not ws_status:
        return
    try:
        now_str = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        for i, r in enumerate(site_results, start=2):
            row_data = [
                r["name"],
                r["state"],
                r.get("checked_at", now_str),
                "✅ OK" if r.get("http_ok") else "❌ Failed",
                r.get("links_count", "—"),
                r.get("result", "—"),
                r.get("error", "")[:200],
            ]
            ws_status.update(f"A{i}:G{i}", [row_data])

            # Green row = OK, red row = failed
            color = ({"red": 0.85, "green": 0.95, "blue": 0.85}
                     if r.get("http_ok")
                     else {"red": 0.98, "green": 0.85, "blue": 0.85})
            ws_status.format(f"A{i}:G{i}", {"backgroundColor": color})

        log.info("✅ Monitor Status sheet updated.")
    except Exception as e:
        log.error(f"❌ Could not update status sheet: {e}")


def build_heartbeat_email(site_results, total_notifications_today):
    """
    Daily summary email — confirms the script is alive and shows
    the status of every monitored site.
    """
    now     = datetime.now().strftime("%d %B %Y, %I:%M %p")
    subject = f"✅ Dravio Monitor — Daily Health Check ({datetime.now().strftime('%d %b %Y')})"

    uptime_hours = round((datetime.now() - SCRIPT_START_TIME).total_seconds() / 3600, 1)

    rows = ""
    for r in site_results:
        status_cell = ("✅ OK" if r.get("http_ok")
                       else "<span style='color:#c0392b;font-weight:bold;'>❌ Failed</span>")
        result_cell = r.get("result", "—")
        rows += f"""
        <tr>
          <td style="padding:9px 10px;border:1px solid #dee2e6;font-size:13px;">
            {r['name']}</td>
          <td style="padding:9px 10px;border:1px solid #dee2e6;font-size:13px;
                     text-align:center;">{status_cell}</td>
          <td style="padding:9px 10px;border:1px solid #dee2e6;font-size:13px;">
            {r.get('checked_at', '—')}</td>
          <td style="padding:9px 10px;border:1px solid #dee2e6;font-size:13px;
                     text-align:center;">{r.get('links_count', '—')}</td>
          <td style="padding:9px 10px;border:1px solid #dee2e6;font-size:13px;
                     color:#555;">{result_cell}</td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f4f4;">

      <div style="background:#1a6e3c;padding:18px 24px;">
        <h1 style="color:#fff;margin:0;font-size:20px;">
          ✅ Dravio Monitor — Daily Health Report
        </h1>
        <p style="color:#a8d8be;margin:4px 0 0;font-size:12px;">
          Script is running normally · {now} (IST)
        </p>
      </div>

      <div style="max-width:700px;margin:20px auto;background:#fff;
                  border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <!-- Summary Strip -->
        <div style="display:flex;padding:0;">
          <div style="flex:1;padding:16px;text-align:center;
                      border-right:1px solid #dee2e6;background:#f8f9fa;">
            <div style="font-size:28px;font-weight:bold;color:#1a6e3c;">
              {sum(1 for r in site_results if r.get('http_ok'))}/{len(site_results)}
            </div>
            <div style="font-size:11px;color:#666;margin-top:4px;">Sites Reachable</div>
          </div>
          <div style="flex:1;padding:16px;text-align:center;
                      border-right:1px solid #dee2e6;background:#f8f9fa;">
            <div style="font-size:28px;font-weight:bold;color:#1a3c6e;">
              {total_notifications_today}
            </div>
            <div style="font-size:11px;color:#666;margin-top:4px;">
              New Notifications (today)</div>
          </div>
          <div style="flex:1;padding:16px;text-align:center;background:#f8f9fa;">
            <div style="font-size:28px;font-weight:bold;color:#555;">
              {uptime_hours}h
            </div>
            <div style="font-size:11px;color:#666;margin-top:4px;">Uptime This Session</div>
          </div>
        </div>

        <!-- Site Status Table -->
        <div style="padding:20px 24px;">
          <p style="font-weight:bold;color:#1a3c6e;margin:0 0 10px;font-size:14px;">
            📊 Site Status — Last Check Cycle
          </p>
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <tr style="background:#1a3c6e;">
              <th style="padding:8px 10px;color:#fff;text-align:left;">Site</th>
              <th style="padding:8px;color:#fff;width:70px;">Status</th>
              <th style="padding:8px 10px;color:#fff;text-align:left;">Last Checked</th>
              <th style="padding:8px;color:#fff;width:80px;">Links</th>
              <th style="padding:8px 10px;color:#fff;text-align:left;">Result</th>
            </tr>
            {rows}
          </table>
        </div>

        <div style="padding:0 24px 20px;">
          <p style="color:#888;font-size:11px;margin:0;text-align:center;">
            Check the
            <strong>Monitor Status</strong> tab in Google Sheets for live status at any time.<br>
            Next daily report: tomorrow at {HEARTBEAT_TIME} IST
          </p>
        </div>

        <div style="background:#f8f9fa;padding:12px 20px;
                    border-top:1px solid #dee2e6;text-align:center;">
          <p style="margin:0;color:#6c757d;font-size:11px;">
            Automated daily report from <strong>Dravio Monitor v3.0</strong>
          </p>
        </div>
      </div>

    </body>
    </html>
    """
    return subject, html


def log_to_sheet(ws, site, notification):
    """
    Append one row per new notification.
    notification = {title, link}  (from extract_notification_links)
    or a plain string for error rows.
    """
    if not ws:
        log.warning("⚠️  Skipping sheet log — worksheet unavailable.")
        return

    try:
        existing = ws.get_all_values()
        sr_no    = len(existing)
        now      = datetime.now()

        if isinstance(notification, dict):
            title  = notification.get("title", "")[:500]
            link   = notification.get("link", "")
            status = "New — Pending Review"
        else:
            title  = str(notification)[:500]
            link   = ""
            status = "Error"

        row = [
            sr_no,
            now.strftime("%d-%m-%Y"),
            now.strftime("%I:%M %p"),
            site["state"],
            site["name"],
            title,
            link,
            site["url"],
            status,
            "Auto-detected by Dravio Monitor",
        ]
        ws.append_row(row)

        row_num = len(existing) + 1
        ws.format(f"A{row_num}:J{row_num}", {
            "backgroundColor": {"red": 1.0, "green": 0.98, "blue": 0.80}
        })
        log.info(f"   📊 Logged to Sheet row {row_num}: {title[:60]}")

    except Exception as e:
        log.error(f"❌ Failed to log to Google Sheet: {e}")


# ----------------------------------------------------------------
# Email Alerts
# ----------------------------------------------------------------
def send_email(subject, html_body, recipients=None):
    """Send an HTML email alert."""
    if recipients is None:
        recipients = ALERT_RECIPIENTS
    try:
        msg            = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"Dravio Monitor <{GMAIL_SENDER}>"
        msg["To"]      = ", ".join(recipients)
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_SENDER, recipients, msg.as_string())

        log.info(f"   📧 Email sent to: {', '.join(recipients)}")

    except Exception as e:
        log.error(f"❌ Email sending failed: {e}")


def build_notification_email(site, new_notifications):
    """
    Build a rich HTML email listing each new notification with its
    title and a direct clickable link to the PDF / page.
    new_notifications = list of {title, link}
    """
    now     = datetime.now().strftime("%d %B %Y, %I:%M %p")
    count   = len(new_notifications)
    subject = (f"🔔 Dravio Alert: {count} New Notification"
               f"{'s' if count > 1 else ''} — {site['name']}")

    # Build one card per notification
    notif_cards = ""
    for i, n in enumerate(new_notifications, 1):
        is_pdf  = n["link"].lower().endswith(".pdf")
        link_label = "📄 Download PDF" if is_pdf else "🔗 Open Link"
        bg_color   = "#f8f9fa" if i % 2 == 0 else "#ffffff"
        notif_cards += f"""
        <tr style="background:{bg_color};">
          <td style="padding:10px 8px;border:1px solid #dee2e6;
                     font-size:12px;color:#555;text-align:center;width:28px;">
            {i}
          </td>
          <td style="padding:10px;border:1px solid #dee2e6;font-size:13px;color:#212529;">
            {n['title']}
          </td>
          <td style="padding:10px;border:1px solid #dee2e6;text-align:center;white-space:nowrap;">
            <a href="{n['link']}"
               style="background:#1a3c6e;color:#fff;padding:6px 14px;
                      border-radius:4px;text-decoration:none;font-size:12px;
                      font-weight:bold;display:inline-block;">
              {link_label}
            </a>
          </td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f4f4;">

      <div style="background:#1a3c6e;padding:18px 24px;">
        <h1 style="color:#fff;margin:0;font-size:20px;">
          🔔 Dravio — New Counselling Notification
        </h1>
        <p style="color:#a8c4e0;margin:4px 0 0;font-size:12px;">
          NEET UG MBBS Counselling Monitor · Automated Alert
        </p>
      </div>

      <div style="max-width:680px;margin:20px auto;background:#fff;
                  border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <div style="background:#fff3cd;border-left:5px solid #ffc107;padding:14px 20px;">
          <p style="margin:0;font-size:14px;font-weight:bold;color:#856404;">
            ⚠️ {count} new notification{'s' if count > 1 else ''} detected on {site['name']}
          </p>
        </div>

        <div style="padding:20px 24px 8px;">
          <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">
            <tr>
              <td style="padding:8px 10px;background:#f8f9fa;border:1px solid #dee2e6;
                         width:30%;font-weight:bold;color:#495057;">🏛️ State / Body</td>
              <td style="padding:8px 10px;border:1px solid #dee2e6;
                         font-weight:bold;color:#1a3c6e;">{site['state']}</td>
            </tr>
            <tr>
              <td style="padding:8px 10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">📋 Authority</td>
              <td style="padding:8px 10px;border:1px solid #dee2e6;color:#212529;">
                {site['name']}</td>
            </tr>
            <tr>
              <td style="padding:8px 10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">🕐 Detected At</td>
              <td style="padding:8px 10px;border:1px solid #dee2e6;color:#212529;">
                {now} (IST)</td>
            </tr>
          </table>

          <p style="font-weight:bold;color:#1a3c6e;margin:16px 0 8px;font-size:14px;">
            📄 New Notifications:
          </p>
          <table style="width:100%;border-collapse:collapse;">
            <tr style="background:#1a3c6e;">
              <th style="padding:8px;color:#fff;font-size:12px;width:28px;">#</th>
              <th style="padding:8px;color:#fff;font-size:12px;text-align:left;">
                Notification Title</th>
              <th style="padding:8px;color:#fff;font-size:12px;width:120px;">Link</th>
            </tr>
            {notif_cards}
          </table>
        </div>

        <div style="padding:16px 24px;">
          <div style="text-align:center;">
            <a href="{site['url']}"
               style="background:#1a3c6e;color:#fff;padding:11px 28px;
                      border-radius:5px;text-decoration:none;
                      font-size:13px;font-weight:bold;display:inline-block;">
              🌐 Open Official Site
            </a>
          </div>
          <p style="color:#888;font-size:11px;text-align:center;margin-top:12px;">
            These notifications have also been logged in the Dravio Google Sheet.
          </p>
        </div>

        <div style="background:#f8f9fa;padding:12px 20px;
                    border-top:1px solid #dee2e6;text-align:center;">
          <p style="margin:0;color:#6c757d;font-size:11px;">
            Automated alert from <strong>Dravio Monitor v3.0</strong><br>
            Gujarat · Maharashtra · Madhya Pradesh · Rajasthan · MCC (AIQ) · NTA
          </p>
        </div>
      </div>

    </body>
    </html>
    """
    return subject, html


def build_failure_email(error_details):
    """Build HTML email for monitor failure/crash alert."""
    now     = datetime.now().strftime("%d %B %Y, %I:%M %p")
    subject = "🚨 URGENT: Dravio Monitor Has Stopped Working"

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;">
      <div style="background:#c0392b;padding:18px 24px;">
        <h1 style="color:#fff;margin:0;font-size:20px;">
          🚨 Dravio Monitor — Critical Failure Alert
        </h1>
      </div>
      <div style="max-width:620px;margin:20px auto;background:#fff;
                  border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.1);">
        <div style="background:#fdecea;border-left:5px solid #c0392b;padding:14px 20px;">
          <p style="margin:0;font-size:14px;font-weight:bold;color:#c0392b;">
            ⚠️ The Dravio monitoring script has stopped running!
          </p>
        </div>
        <div style="padding:24px;">
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         width:32%;font-weight:bold;color:#495057;">🕐 Failed At</td>
              <td style="padding:10px;border:1px solid #dee2e6;">{now} (IST)</td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">❌ Error</td>
              <td style="padding:10px;border:1px solid #dee2e6;
                         font-family:monospace;font-size:11px;color:#c0392b;">
                {error_details[:600].replace('<','&lt;').replace('>','&gt;')}
              </td>
            </tr>
            <tr>
              <td style="padding:10px;background:#f8f9fa;border:1px solid #dee2e6;
                         font-weight:bold;color:#495057;">⏰ Auto Restart</td>
              <td style="padding:10px;border:1px solid #dee2e6;">
                Will attempt restart in {RESTART_WAIT_MINUTES} minutes</td>
            </tr>
          </table>
        </div>
        <div style="background:#f8f9fa;padding:12px 20px;
                    border-top:1px solid #dee2e6;text-align:center;">
          <p style="margin:0;color:#888;font-size:11px;">
            Automated failure notification from <strong>Dravio Monitor</strong>
          </p>
        </div>
      </div>
    </body>
    </html>
    """
    return subject, html


# ----------------------------------------------------------------
# Page Fetching & Notification Extraction
# ----------------------------------------------------------------
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Link texts that are navigation / chrome — never real notifications
_NAV_SKIP = {
    "home", "about us", "contact us", "useful links", "login", "register",
    "previous", "next", "back", "sitemap", "screen reader",
    "skip to main content", "rss", "print", "share on facebook",
    "share on x (formerly twitter)", "share on linkedin", "share on x",
    "view", "download", "click here", "read more", "more", "here",
    "work with us", "rti", "user manual", "statistics", "downloads",
    "candidate", "institute", "cets", "cap", "contact", "accessible version",
    "all rights reserved.", "foreign candidate registration a.y 2026-27",
}


def fetch_page(site):
    """
    Fetch site HTML. Returns raw html string.
    SSL verification disabled — many Indian govt sites use expired certs.
    """
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    resp = requests.get(site["url"], headers=HTTP_HEADERS, timeout=30, verify=False)
    resp.raise_for_status()
    return resp.text


def extract_notification_links(html, site):
    """
    Parse page HTML and return a list of notification dicts:
      {id, title, link}

    - Skips navigation links, very short texts, anchors, JS links
    - Applies filter_kw to link text when configured (e.g. Maharashtra)
    - ID = MD5 of the href so the same PDF URL is never double-counted
    """
    base_url  = site["url"]
    filter_kw = site.get("filter_kw")

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "header", "meta", "noscript"]):
        tag.decompose()

    results  = []
    seen_ids = set()

    for a in soup.find_all("a", href=True):
        text = a.get_text(separator=" ", strip=True)
        href = a["href"].strip()

        # Skip blank, very short, or pure-navigation links
        if not text or len(text) < 15:
            continue
        if text.lower().rstrip(".…!") in _NAV_SKIP:
            continue
        if href.startswith("#") or href.lower().startswith("javascript"):
            continue

        # Resolve relative URLs
        if not href.startswith("http"):
            href = urljoin(base_url, href)

        # For Maharashtra, only keep links related to medical/health courses
        if filter_kw:
            if not any(kw.lower() in text.lower() for kw in filter_kw):
                continue

        nid = hashlib.md5(href.encode("utf-8")).hexdigest()
        if nid in seen_ids:
            continue
        seen_ids.add(nid)

        results.append({
            "id"   : nid,
            "title": text[:300],
            "link" : href,
        })

    return results


def extract_text_for_fallback(html, site):
    """
    Extract plain text from page (used as a fallback change detector
    for sites like MP that currently have no notification links).
    Applies strip_noise and filter_kw.
    """
    strip_noise = site.get("strip_noise", [])
    filter_kw   = site.get("filter_kw")

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header",
                     "meta", "link", "noscript"]):
        tag.decompose()

    raw = soup.get_text(separator="\n", strip=True)

    if strip_noise:
        lines = raw.splitlines()
        raw   = "\n".join(
            l for l in lines
            if not any(re.search(p, l.strip(), re.IGNORECASE) for p in strip_noise)
        )

    if filter_kw:
        lines    = raw.splitlines()
        filtered = [l for l in lines if any(kw.lower() in l.lower() for kw in filter_kw)]
        return "\n".join(filtered) if filtered else raw

    return raw


def get_hash(text):
    return hashlib.md5(text.encode("utf-8")).hexdigest()


# ----------------------------------------------------------------
# State Management
# State format (v3):
#   {
#     "site_name": {
#       "seen_ids":    ["md5_of_href", ...],   # tracked notification links
#       "text_hash":   "md5",                  # fallback for link-less pages
#       "last_checked": "ISO datetime"
#     }
#   }
# ----------------------------------------------------------------
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def migrate_site_state(old_entry, notifications, text_hash):
    """
    Convert old-format state (had 'hash'/'snapshot' keys) to v3 format.
    Treats all currently-visible notifications as already-seen so we
    don't flood alerts on the first run after upgrade.
    """
    return {
        "seen_ids"    : [n["id"] for n in notifications],
        "text_hash"   : text_hash,
        "last_checked": datetime.now().isoformat(),
    }


# ----------------------------------------------------------------
# Notification counter (resets at midnight via date comparison)
# ----------------------------------------------------------------
_today_count      = {"date": datetime.now().date(), "count": 0}

def _increment_today(n=1):
    if _today_count["date"] != datetime.now().date():
        _today_count["date"]  = datetime.now().date()
        _today_count["count"] = 0
    _today_count["count"] += n


# ----------------------------------------------------------------
# Main Check Function
# ----------------------------------------------------------------
def check_all_sites():
    log.info("=" * 60)
    log.info(f"🔍 Starting check — {datetime.now().strftime('%d-%m-%Y %I:%M %p IST')}")
    log.info("=" * 60)

    state        = load_state()
    ws           = get_worksheet()
    ws_status    = get_status_worksheet()
    any_change   = False
    site_results = []   # collected for status sheet + heartbeat email

    for site in SITES:
        name = site["name"]
        log.info(f"\n📡 Checking: {name}")
        log.info(f"   URL: {site['url']}")

        retry_count = 0
        success     = False
        site_record = {
            "name"       : name,
            "state"      : site["state"],
            "checked_at" : datetime.now().strftime("%d-%m-%Y %I:%M %p"),
            "http_ok"    : False,
            "links_count": 0,
            "result"     : "—",
            "error"      : "",
        }

        while retry_count < MAX_RETRIES and not success:
            try:
                html          = fetch_page(site)
                notifications = extract_notification_links(html, site)
                fallback_text = extract_text_for_fallback(html, site)
                text_hash     = get_hash(fallback_text)
                current_ids   = {n["id"] for n in notifications}

                site_record["http_ok"]     = True
                site_record["links_count"] = len(notifications)
                site_record["checked_at"]  = datetime.now().strftime("%d-%m-%Y %I:%M %p")

                site_state = state.get(name, {})

                # ── Migrate old state format ──────────────────────
                if "hash" in site_state and "seen_ids" not in site_state:
                    log.info(f"   🔄 Migrating state to v3 format")
                    state[name] = migrate_site_state(site_state, notifications, text_hash)
                    site_record["result"] = f"Migrated baseline ({len(notifications)} links)"
                    log.info(f"   ✅ Baseline saved ({len(notifications)} existing links indexed)")
                    success = True
                    break

                # ── First run for this site ───────────────────────
                if not site_state:
                    state[name] = {
                        "seen_ids"    : list(current_ids),
                        "text_hash"   : text_hash,
                        "last_checked": datetime.now().isoformat(),
                    }
                    site_record["result"] = f"First run — {len(notifications)} links baselined"
                    log.info(f"   ✅ First run — baseline saved "
                             f"({len(notifications)} notifications indexed)")

                # ── Subsequent runs ───────────────────────────────
                else:
                    seen_ids       = set(site_state.get("seen_ids", []))
                    new_notifs     = [n for n in notifications if n["id"] not in seen_ids]
                    prev_text_hash = site_state.get("text_hash", "")

                    if new_notifs:
                        any_change = True
                        _increment_today(len(new_notifs))
                        log.info(f"   🔔 {len(new_notifs)} NEW notification(s) found!")
                        for n in new_notifs:
                            log.info(f"      + {n['title'][:80]}")
                            log.info(f"        {n['link']}")

                        subject, html_body = build_notification_email(site, new_notifs)
                        send_email(subject, html_body)
                        for n in new_notifs:
                            log_to_sheet(ws, site, n)

                        site_record["result"] = f"🔔 {len(new_notifs)} new notification(s) sent"

                    elif text_hash != prev_text_hash and not notifications:
                        # Fallback for link-less pages (e.g. MP)
                        old_lines = {l.strip() for l in
                                     site_state.get("snapshot_text", "").splitlines()
                                     if len(l.strip()) > 20}
                        new_lines = [l.strip() for l in fallback_text.splitlines()
                                     if len(l.strip()) > 20]
                        added     = [l for l in new_lines if l not in old_lines]

                        if added:
                            any_change = True
                            _increment_today(1)
                            log.info(f"   🔔 New text content detected on {name}!")
                            text_notif = {
                                "title": " | ".join(added[:3])[:300],
                                "link" : site["url"],
                            }
                            subject, html_body = build_notification_email(site, [text_notif])
                            send_email(subject, html_body)
                            log_to_sheet(ws, site, text_notif)
                            site_record["result"] = "🔔 New text notification sent"
                        else:
                            site_record["result"] = "Page reformatted — no new content"
                            log.info(f"   ℹ️  Page changed (no new content added)")
                    else:
                        site_record["result"] = "✅ No new notifications"
                        log.info(f"   ✅ No new notifications")

                    # Always update state
                    state[name] = {
                        "seen_ids"     : list(current_ids),
                        "text_hash"    : text_hash,
                        "snapshot_text": fallback_text[:2000],
                        "last_checked" : datetime.now().isoformat(),
                    }

                success = True

            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code if e.response is not None else 0
                err_msg = f"HTTP {status_code}: {e}"
                site_record["error"]  = err_msg
                site_record["result"] = f"❌ HTTP {status_code}"
                if status_code in (403, 401, 404, 410):
                    # No point retrying — server is actively refusing or site is gone
                    log.warning(f"   ⚠️  HTTP {status_code} — skipping retries for {name}")
                    success = True   # mark done so we don't loop
                else:
                    retry_count += 1
                    log.warning(f"   ⚠️  Fetch failed (attempt {retry_count}/{MAX_RETRIES}): {e}")
                    if retry_count < MAX_RETRIES:
                        time.sleep(30)

            except requests.exceptions.RequestException as e:
                retry_count += 1
                err_msg = str(e)
                log.warning(f"   ⚠️  Fetch failed (attempt {retry_count}/{MAX_RETRIES}): {e}")
                site_record["error"] = err_msg
                if retry_count < MAX_RETRIES:
                    time.sleep(30)

            except Exception as e:
                err_msg = str(e)
                log.error(f"   ❌ Unexpected error for {name}: {e}")
                log_to_sheet(ws, site, f"Error: {err_msg}")
                site_record["error"]  = err_msg
                site_record["result"] = "❌ Error"
                break

        if not success:
            log.error(f"   ❌ All {MAX_RETRIES} retries failed for {name}")
            site_record["result"] = f"❌ Failed after {MAX_RETRIES} retries"
            # Only log to Sheet on first failure (avoid flooding with repeated errors)
            last_error = state.get(name, {}).get("last_error", "")
            current_error = site_record.get("error", "fetch failed")
            if last_error != current_error:
                log_to_sheet(ws, site, f"Fetch failed: {current_error}")
                state.setdefault(name, {})["last_error"] = current_error

        site_results.append(site_record)

    save_state(state)

    # Update the Monitor Status sheet with this cycle's results
    update_status_sheet(ws_status, site_results)

    if not any_change:
        log.info(f"\n✅ All sites checked — No new notifications found.")

    log.info(f"\n⏰ Next check in {CHECK_INTERVAL_MINUTES} minutes...")


def send_daily_heartbeat():
    """Send the morning health-check summary email."""
    log.info("📬 Sending daily heartbeat email...")
    state        = load_state()
    site_results = []
    for site in SITES:
        s = state.get(site["name"], {})
        raw_ts = s.get("last_checked", "")
        try:
            checked_at = (datetime.fromisoformat(raw_ts)
                          .strftime("%d-%m-%Y %I:%M %p") if raw_ts else "Never")
        except ValueError:
            checked_at = raw_ts
        site_results.append({
            "name"       : site["name"],
            "state"      : site["state"],
            "checked_at" : checked_at,
            "http_ok"    : bool(raw_ts),
            "links_count": len(s.get("seen_ids", [])),
            "result"     : "Refer to Notification Log sheet for details",
            "error"      : "",
        })

    subject, html_body = build_heartbeat_email(site_results, _today_count["count"])
    send_email(subject, html_body)
    log.info("✅ Daily heartbeat email sent.")


# ----------------------------------------------------------------
# Startup Validation
# ----------------------------------------------------------------
def validate_config():
    errors = []
    if not GMAIL_SENDER:
        errors.append("GMAIL_SENDER is not set in .env")
    if not GMAIL_APP_PASSWORD or GMAIL_APP_PASSWORD == "xxxx xxxx xxxx xxxx":
        errors.append("GMAIL_APP_PASSWORD is not set in .env")
    if not ALERT_RECIPIENTS:
        errors.append("ALERT_RECIPIENTS is not set in .env")
    if not GOOGLE_SHEET_ID or GOOGLE_SHEET_ID == "your_sheet_id_here":
        errors.append("GOOGLE_SHEET_ID is not set in .env")
    if not os.path.exists(CREDENTIALS_FILE):
        errors.append(f"credentials.json not found at: {CREDENTIALS_FILE}")

    if errors:
        log.error("❌ Configuration errors found:")
        for err in errors:
            log.error(f"   • {err}")
        sys.exit(1)

    log.info("✅ Configuration validated successfully.")


# ----------------------------------------------------------------
# Entry Point with Auto-Restart on Crash
# ----------------------------------------------------------------
def run_test_mode():
    """
    --test mode: fetch all sites right now, treat the first 2 notifications
    from each as 'new', send a real test email and write test rows to the
    Sheet, then exit. Your real state/baseline is never touched.
    """
    log.info("=" * 60)
    log.info("   🧪 Dravio Monitor — PIPELINE TEST MODE")
    log.info("=" * 60)
    log.info("   Fetching all sites and simulating new notifications...")
    log.info("   (Real state is NOT modified. Sheet rows are marked [TEST].)")
    log.info("=" * 60)

    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    ws = get_worksheet()
    overall_ok = True

    for site in SITES:
        name = site["name"]
        log.info(f"\n📡 Testing: {name}")

        try:
            html          = fetch_page(site)
            notifications = extract_notification_links(html, site)
            fallback_text = extract_text_for_fallback(html, site)

            log.info(f"   ✅ Site reachable")
            log.info(f"   📎 Notifications extracted: {len(notifications)}")

            if not notifications:
                log.info(f"   ⚠️  No notification links found on this site.")
                log.info(f"      (Fallback text length: {len(fallback_text)} chars)")
                # Show a snippet of what text was extracted
                snippet = " | ".join(
                    l.strip() for l in fallback_text.splitlines()
                    if len(l.strip()) > 20
                )[:300]
                log.info(f"      Text preview: {snippet}")
                continue

            # Show what was found
            for i, n in enumerate(notifications[:5], 1):
                log.info(f"      {i}. {n['title'][:80]}")
                log.info(f"         {n['link']}")
            if len(notifications) > 5:
                log.info(f"      ... and {len(notifications) - 5} more")

            # Use first 2 as simulated "new" notifications
            test_notifs = [
                {**n, "title": f"[TEST] {n['title']}"}
                for n in notifications[:2]
            ]

            # Send test email
            subject, html_body = build_notification_email(site, test_notifs)
            subject = f"[TEST] {subject}"
            send_email(subject, html_body)
            log.info(f"   📧 Test email sent")

            # Write test rows to Sheet (Status = [TEST])
            if ws:
                for n in test_notifs:
                    existing = ws.get_all_values()
                    sr_no    = len(existing)
                    now      = datetime.now()
                    row = [
                        sr_no,
                        now.strftime("%d-%m-%Y"),
                        now.strftime("%I:%M %p"),
                        site["state"],
                        site["name"],
                        n["title"],
                        n["link"],
                        site["url"],
                        "[TEST] — Delete this row",
                        "Pipeline test run",
                    ]
                    ws.append_row(row)
                    row_num = len(existing) + 1
                    # Light blue to distinguish test rows
                    ws.format(f"A{row_num}:J{row_num}", {
                        "backgroundColor": {"red": 0.80, "green": 0.90, "blue": 1.0}
                    })
                log.info(f"   📊 {len(test_notifs)} test row(s) written to Sheet (blue rows)")

        except Exception as e:
            overall_ok = False
            log.error(f"   ❌ FAILED: {e}")

    log.info("\n" + "=" * 60)
    if overall_ok:
        log.info("   ✅ TEST COMPLETE — All sites reachable, email sent, Sheet updated.")
        log.info("   Check your inbox and the Google Sheet to verify the output.")
        log.info("   Delete the blue [TEST] rows from the Sheet when done.")
    else:
        log.info("   ⚠️  TEST FINISHED WITH ERRORS — check the output above.")
    log.info("=" * 60)


if __name__ == "__main__":

    validate_config()

    if "--test" in sys.argv:
        run_test_mode()
        sys.exit(0)

    log.info("=" * 60)
    log.info("   🏥 Dravio Counselling Monitor v3.0 — Starting Up")
    log.info("=" * 60)
    log.info(f"   Monitoring  : {len(SITES)} sites")
    for s in SITES:
        log.info(f"                 • {s['name']}")
    log.info(f"   Interval    : Every {CHECK_INTERVAL_MINUTES} minutes")
    log.info(f"   Alerts to   : {', '.join(ALERT_RECIPIENTS)}")
    log.info(f"   Started At  : {datetime.now().strftime('%d-%m-%Y %I:%M %p IST')}")
    log.info("=" * 60)

    consecutive_failures = 0

    while True:
        try:
            check_all_sites()
            consecutive_failures = 0

            schedule.every(CHECK_INTERVAL_MINUTES).minutes.do(check_all_sites)
            schedule.every().day.at(HEARTBEAT_TIME).do(send_daily_heartbeat)
            log.info(f"   📬 Daily heartbeat scheduled at {HEARTBEAT_TIME} IST")

            while True:
                schedule.run_pending()
                time.sleep(30)

        except KeyboardInterrupt:
            log.info("\n⛔ Monitor stopped by user.")
            sys.exit(0)

        except Exception as e:
            consecutive_failures += 1
            error_details = traceback.format_exc()
            log.error(f"💥 Critical crash (failure #{consecutive_failures}):\n{error_details}")

            try:
                subject, html = build_failure_email(error_details)
                send_email(subject, html)
                log.info("📧 Failure alert email sent.")
            except Exception as mail_err:
                log.error(f"❌ Could not send failure email: {mail_err}")

            schedule.clear()
            log.info(f"⏳ Waiting {RESTART_WAIT_MINUTES} minutes before restarting...")
            time.sleep(RESTART_WAIT_MINUTES * 60)
            log.info("🔄 Restarting monitor...")
