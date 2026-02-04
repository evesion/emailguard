#!/usr/bin/env python3
"""
===============================================================================
                    EMAILGUARD INBOX PLACEMENT TESTER
===============================================================================

A tool to test email inbox placement rates across Google and Microsoft
using the EmailGuard.io API.

SETUP:
    1. Create a .env file with your API key:
       EMAILGUARD_API_KEY=your_api_key_here

    2. Place your email accounts CSV file in the same folder
       The CSV must have columns: from_name, from_email, user_name, password, smtp_host

USAGE:
    python3 emailguard.py

GitHub: https://github.com/evesion/emailguard
===============================================================================
"""

__version__ = "1.1.0"
__repo__ = "evesion/emailguard"

import csv
import json
import logging
import os
import shutil
import smtplib
import ssl
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from email.mime.text import MIMEText
from email.utils import formatdate

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

CSV_INPUT_FILE = 'smartlead_output.csv'
SMTP_PORT = 465
EMAIL_SUBJECT = "Team Meeting Code"
EMAIL_BODY = """Hello Team, Please find here todays meeting code:"""

MAX_DOMAINS_PER_BATCH = 50
MAX_PARALLEL_WORKERS = 5
EMAIL_DELAY_SECONDS = 3
POLL_INTERVAL_SECONDS = 30

# ═══════════════════════════════════════════════════════════════════════════════
# INTERNAL SETTINGS
# ═══════════════════════════════════════════════════════════════════════════════

API_BASE_URL = "https://app.emailguard.io/api/v1"
DATA_DIR = '.emailguard_data'
CONFIG_FILE = '.env'
BATCH_STATE_FILE = os.path.join(DATA_DIR, 'batch_state.json')
OUTPUT_CSV_FILE = os.path.join(DATA_DIR, 'test_queue.csv')
RESULTS_CSV_FILE = 'inbox_placement_results.csv'
PDF_REPORT_FILE = 'inbox_placement_report.pdf'

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

API_KEY = None


# ═══════════════════════════════════════════════════════════════════════════════
# CORE FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def load_env_file(filepath):
    env_vars = {}
    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    env_vars[key.strip()] = value.strip().strip('"\'')
    return env_vars


def load_api_key():
    global API_KEY
    API_KEY = os.environ.get('EMAILGUARD_API_KEY')
    if not API_KEY:
        env_vars = load_env_file(CONFIG_FILE)
        API_KEY = env_vars.get('EMAILGUARD_API_KEY')
    return API_KEY is not None


def save_api_key(key):
    global API_KEY
    API_KEY = key
    with open(CONFIG_FILE, 'w') as f:
        f.write(f"# EmailGuard Configuration\n")
        f.write(f"EMAILGUARD_API_KEY={key}\n")


def ensure_data_dir():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)


def create_api_session():
    session = requests.Session()
    retry_strategy = Retry(total=3, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=MAX_PARALLEL_WORKERS, pool_maxsize=MAX_PARALLEL_WORKERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"})
    return session


def load_batch_state():
    ensure_data_dir()
    if os.path.exists(BATCH_STATE_FILE):
        try:
            with open(BATCH_STATE_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {'processed_domains': [], 'batch_number': 0}


def save_batch_state(state):
    ensure_data_dir()
    with open(BATCH_STATE_FILE, 'w') as f:
        json.dump(state, f, indent=2)


def get_unique_domains(csv_file):
    domains = []
    domain_to_row = {}
    try:
        with open(csv_file, mode='r', encoding='utf-8-sig') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                domain = row['from_email'].split('@')[-1]
                if domain not in domain_to_row:
                    domains.append(domain)
                    domain_to_row[domain] = row
    except FileNotFoundError:
        return None, None, f"CSV file '{csv_file}' not found"
    except KeyError as e:
        return None, None, f"CSV missing required column: {e}"
    return domains, domain_to_row, None


def send_email(from_name, from_email, user_name, password, smtp_host, recipients, filter_phrase):
    server = None
    try:
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        server = smtplib.SMTP_SSL(host=smtp_host, port=SMTP_PORT, context=context, timeout=30)
        server.login(user_name, password)
        recipients_list = recipients.split(';')
        msg = MIMEText(f"{EMAIL_BODY}\n\n{filter_phrase}", 'plain')
        msg['From'] = f"{from_name} <{from_email}>"
        msg['To'] = ', '.join(recipients_list)
        msg['Subject'] = EMAIL_SUBJECT
        msg['Date'] = formatdate(localtime=True)
        server.sendmail(from_email, recipients_list, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        if server:
            try:
                server.quit()
            except:
                pass


def create_test(session, test_name):
    try:
        response = session.post(f"{API_BASE_URL}/inbox-placement-tests", json={"name": test_name})
        response.raise_for_status()
        return response.json(), None
    except Exception as e:
        return None, str(e)


def get_test_results(session, test_uuid):
    try:
        response = session.get(f"{API_BASE_URL}/inbox-placement-tests/{test_uuid}")
        response.raise_for_status()
        return response.json(), None
    except Exception as e:
        return None, str(e)


def calculate_stats(test_emails):
    if not test_emails:
        return {'total': 0, 'inbox': 0, 'spam': 0, 'inbox_rate': 0, 'spam_rate': 0,
                'google_total': 0, 'google_inbox': 0, 'google_inbox_rate': 0,
                'microsoft_total': 0, 'microsoft_inbox': 0, 'microsoft_inbox_rate': 0, 'waiting': 0}

    total = len(test_emails)
    inbox = spam = waiting = 0
    google_total = google_inbox = microsoft_total = microsoft_inbox = 0

    for email in test_emails:
        folder = (email.get('folder') or '').lower()
        status = (email.get('status') or '').lower()
        provider = (email.get('provider') or '').lower()
        is_google = 'google' in provider
        is_microsoft = 'microsoft' in provider

        if status == 'waiting_for_email':
            waiting += 1
        elif folder == 'inbox':
            inbox += 1
            if is_google: google_inbox += 1
            if is_microsoft: microsoft_inbox += 1
        elif folder in ['spam', 'junk']:
            spam += 1

        if status != 'waiting_for_email':
            if is_google: google_total += 1
            if is_microsoft: microsoft_total += 1

    return {
        'total': total, 'inbox': inbox, 'spam': spam, 'waiting': waiting,
        'inbox_rate': (inbox / total) * 100 if total > 0 else 0,
        'spam_rate': (spam / total) * 100 if total > 0 else 0,
        'google_total': google_total, 'google_inbox': google_inbox,
        'google_inbox_rate': (google_inbox / google_total) * 100 if google_total > 0 else 0,
        'microsoft_total': microsoft_total, 'microsoft_inbox': microsoft_inbox,
        'microsoft_inbox_rate': (microsoft_inbox / microsoft_total) * 100 if microsoft_total > 0 else 0
    }


def check_for_updates(silent=False):
    try:
        response = requests.get(f"https://api.github.com/repos/{__repo__}/releases/latest", timeout=5)
        if response.status_code == 200:
            data = response.json()
            latest_version = data.get('tag_name', '').lstrip('v')
            if latest_version and compare_versions(latest_version, __version__) > 0:
                return {'available': True, 'current': __version__, 'latest': latest_version,
                        'url': data.get('html_url', '')}
        return {'available': False, 'current': __version__}
    except:
        return {'available': False, 'current': __version__}


def compare_versions(v1, v2):
    def parse(v):
        return [int(x) for x in v.split('.')]
    try:
        p1, p2 = parse(v1), parse(v2)
        for i in range(max(len(p1), len(p2))):
            a = p1[i] if i < len(p1) else 0
            b = p2[i] if i < len(p2) else 0
            if a > b: return 1
            if a < b: return -1
        return 0
    except:
        return 0


def download_update():
    try:
        update_info = check_for_updates()
        if not update_info.get('available'):
            return False, "Already up to date"

        raw_url = f"https://raw.githubusercontent.com/{__repo__}/v{update_info['latest']}/emailguard.py"
        response = requests.get(raw_url, timeout=30)
        if response.status_code != 200:
            return False, f"Download failed (HTTP {response.status_code})"

        new_content = response.text
        try:
            compile(new_content, '<string>', 'exec')
        except SyntaxError as e:
            return False, f"Invalid update file: {e}"

        current_file = os.path.abspath(__file__)
        backup_file = current_file + '.backup'
        shutil.copy2(current_file, backup_file)

        with open(current_file, 'w', encoding='utf-8') as f:
            f.write(new_content)

        return True, f"Updated to v{update_info['latest']}! Please restart."
    except Exception as e:
        return False, str(e)


def fetch_single_result(session, test_info):
    from_email, test_uuid, test_url = test_info
    result, error = get_test_results(session, test_uuid)
    if error or not result or 'data' not in result:
        return {'from_email': from_email, 'test_uuid': test_uuid, 'status': 'FAILED', 'test_url': test_url}
    data = result['data']
    stats = calculate_stats(data.get('inbox_placement_test_emails', []))
    return {
        'from_email': from_email, 'test_uuid': test_uuid, 'test_url': test_url,
        'test_name': data.get('name', ''), 'status': data.get('status', ''),
        'overall_score': data.get('overall_score', ''), 'stats': stats
    }


def generate_pdf(all_results):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.enums import TA_CENTER

    doc = SimpleDocTemplate(PDF_REPORT_FILE, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=24, spaceAfter=30,
                                  alignment=TA_CENTER, textColor=colors.Color(0.2, 0.3, 0.5))
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'], fontSize=12,
                                     alignment=TA_CENTER, textColor=colors.gray, spaceAfter=20)
    section_style = ParagraphStyle('Section', parent=styles['Heading2'], fontSize=14,
                                    spaceBefore=20, spaceAfter=10, textColor=colors.Color(0.2, 0.3, 0.5))

    story = []
    story.append(Paragraph("Email Inbox Placement Report", title_style))
    story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %H:%M')}", subtitle_style))
    story.append(Spacer(1, 20))

    completed = [r for r in all_results if r.get('status') in ['completed', 'complete']]
    pending = [r for r in all_results if r.get('status') not in ['completed', 'complete', 'FAILED']]
    failed = [r for r in all_results if r.get('status') == 'FAILED']

    story.append(Paragraph("Executive Summary", section_style))
    summary_data = [['Total Tests', 'Completed', 'Pending', 'Failed'],
                    [str(len(all_results)), str(len(completed)), str(len(pending)), str(len(failed))]]
    summary_table = Table(summary_data, colWidths=[120, 120, 120, 120])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.3, 0.5)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 14),
        ('BACKGROUND', (0, 1), (-1, 1), colors.Color(0.95, 0.95, 0.95)),
        ('BOX', (0, 0), (-1, -1), 2, colors.Color(0.2, 0.3, 0.5)),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 20))

    if completed:
        avg_inbox = sum(r['stats']['inbox_rate'] for r in completed) / len(completed)
        avg_spam = sum(r['stats']['spam_rate'] for r in completed) / len(completed)
        google_tests = [r for r in completed if r['stats']['google_total'] > 0]
        microsoft_tests = [r for r in completed if r['stats']['microsoft_total'] > 0]
        avg_google = sum(r['stats']['google_inbox_rate'] for r in google_tests) / len(google_tests) if google_tests else 0
        avg_microsoft = sum(r['stats']['microsoft_inbox_rate'] for r in microsoft_tests) / len(microsoft_tests) if microsoft_tests else 0

        def get_color(val, is_spam=False):
            if is_spam:
                return colors.Color(0.1, 0.6, 0.2) if val <= 10 else (colors.Color(0.8, 0.6, 0) if val <= 25 else colors.Color(0.8, 0.2, 0.2))
            return colors.Color(0.1, 0.6, 0.2) if val >= 70 else (colors.Color(0.8, 0.6, 0) if val >= 50 else colors.Color(0.8, 0.2, 0.2))

        story.append(Paragraph("Overall Performance", section_style))
        metrics_data = [['Total Inbox Rate', 'Spam Rate'], [f'{avg_inbox:.1f}%', f'{avg_spam:.1f}%']]
        metrics_table = Table(metrics_data, colWidths=[240, 240])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.3, 0.4, 0.6)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 20),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 1), (0, 1), get_color(avg_inbox)),
            ('TEXTCOLOR', (1, 1), (1, 1), get_color(avg_spam, is_spam=True)),
            ('BACKGROUND', (0, 1), (-1, 1), colors.Color(0.98, 0.98, 0.98)),
            ('BOX', (0, 0), (-1, -1), 2, colors.Color(0.3, 0.4, 0.6)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
            ('TOPPADDING', (0, 0), (-1, -1), 15),
        ]))
        story.append(metrics_table)
        story.append(Spacer(1, 20))

        story.append(Paragraph("Inbox Rate by Provider", section_style))
        provider_data = [['Google', 'Microsoft'], [f'{avg_google:.1f}%', f'{avg_microsoft:.1f}%']]
        provider_table = Table(provider_data, colWidths=[240, 240])
        provider_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.3, 0.4, 0.6)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 24),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 1), (0, 1), get_color(avg_google)),
            ('TEXTCOLOR', (1, 1), (1, 1), get_color(avg_microsoft)),
            ('BACKGROUND', (0, 1), (-1, 1), colors.Color(0.98, 0.98, 0.98)),
            ('BOX', (0, 0), (-1, -1), 2, colors.Color(0.3, 0.4, 0.6)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 18),
            ('TOPPADDING', (0, 0), (-1, -1), 18),
        ]))
        story.append(provider_table)
        story.append(Spacer(1, 30))

    story.append(Paragraph("Detailed Results", section_style))
    table_data = [['Domain', 'Inbox %', 'Google %', 'Microsoft %', 'Spam %', 'Status']]
    for r in all_results:
        domain = r.get('from_email', '').split('@')[-1]
        if len(domain) > 18: domain = domain[:15] + '...'
        stats = r.get('stats', {})
        table_data.append([
            domain,
            f"{stats.get('inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('google_inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('microsoft_inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('spam_rate', 0):.0f}%" if stats else '-',
            r.get('status', 'Unknown')
        ])

    detail_table = Table(table_data, colWidths=[110, 65, 70, 80, 65, 80])
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.3, 0.5)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.8, 0.8, 0.8)),
        ('BOX', (0, 0), (-1, -1), 1, colors.Color(0.2, 0.3, 0.5)),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    for i in range(1, len(table_data)):
        bg = colors.Color(0.95, 0.95, 0.97) if i % 2 == 0 else colors.white
        style_cmds.append(('BACKGROUND', (0, i), (-1, i), bg))
    detail_table.setStyle(TableStyle(style_cmds))
    story.append(detail_table)
    doc.build(story)


# ═══════════════════════════════════════════════════════════════════════════════
# GUI APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False


class EmailGuardApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title(f"EmailGuard Inbox Placement Tester v{__version__}")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # Set theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.session = None
        self.running = False
        self.poll_thread = None

        self.setup_ui()
        self.load_status()
        self.check_api_key()

        # Check for updates
        self.root.after(1000, self.check_updates_async)

    def setup_ui(self):
        # Main container
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        title_label = ctk.CTkLabel(header_frame, text="📧 EmailGuard", font=ctk.CTkFont(size=28, weight="bold"))
        title_label.pack(side="left")

        version_label = ctk.CTkLabel(header_frame, text=f"v{__version__}", font=ctk.CTkFont(size=14), text_color="gray")
        version_label.pack(side="left", padx=(10, 0), pady=(10, 0))

        self.update_btn = ctk.CTkButton(header_frame, text="🔄 Update Available", width=140,
                                         fg_color="#2d5a27", hover_color="#3d7a37", command=self.do_update)
        self.update_btn.pack(side="right")
        self.update_btn.pack_forget()  # Hidden initially

        # Status cards
        cards_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        cards_frame.pack(fill="x", pady=(0, 20))

        # Domains card
        self.domains_card = self.create_stat_card(cards_frame, "Domains", "0 / 0", "📋")
        self.domains_card.pack(side="left", expand=True, fill="x", padx=(0, 10))

        # Tests card
        self.tests_card = self.create_stat_card(cards_frame, "Tests in Queue", "0", "🧪")
        self.tests_card.pack(side="left", expand=True, fill="x", padx=(0, 10))

        # Status card
        self.status_card = self.create_stat_card(cards_frame, "Status", "Ready", "✅")
        self.status_card.pack(side="left", expand=True, fill="x")

        # Action buttons
        buttons_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=(0, 20))

        self.run_btn = ctk.CTkButton(buttons_frame, text="▶️  Run New Tests", height=50,
                                      font=ctk.CTkFont(size=16, weight="bold"), command=self.run_tests)
        self.run_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.results_btn = ctk.CTkButton(buttons_frame, text="📊  Get Results", height=50,
                                          font=ctk.CTkFont(size=16, weight="bold"), command=self.get_results)
        self.results_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.poll_btn = ctk.CTkButton(buttons_frame, text="🔄  Auto-Poll", height=50,
                                       font=ctk.CTkFont(size=16, weight="bold"), command=self.toggle_polling)
        self.poll_btn.pack(side="left", expand=True, fill="x")

        # Progress bar
        self.progress = ctk.CTkProgressBar(self.main_frame)
        self.progress.pack(fill="x", pady=(0, 10))
        self.progress.set(0)

        # Log output
        log_label = ctk.CTkLabel(self.main_frame, text="Activity Log", font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        log_label.pack(fill="x")

        self.log_text = ctk.CTkTextbox(self.main_frame, height=250, font=ctk.CTkFont(family="Courier", size=12))
        self.log_text.pack(fill="both", expand=True, pady=(5, 20))

        # Bottom buttons
        bottom_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bottom_frame.pack(fill="x")

        self.settings_btn = ctk.CTkButton(bottom_frame, text="⚙️ Settings", width=120, command=self.show_settings)
        self.settings_btn.pack(side="left")

        self.reset_btn = ctk.CTkButton(bottom_frame, text="🗑️ Reset All", width=120,
                                        fg_color="#8B0000", hover_color="#A52A2A", command=self.reset_all)
        self.reset_btn.pack(side="left", padx=(10, 0))

        self.csv_btn = ctk.CTkButton(bottom_frame, text="📁 Select CSV", width=120, command=self.select_csv)
        self.csv_btn.pack(side="right")

        self.open_report_btn = ctk.CTkButton(bottom_frame, text="📄 Open Report", width=120, command=self.open_report)
        self.open_report_btn.pack(side="right", padx=(0, 10))

    def create_stat_card(self, parent, title, value, icon):
        card = ctk.CTkFrame(parent, corner_radius=10)

        icon_label = ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=24))
        icon_label.pack(pady=(15, 5))

        value_label = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=20, weight="bold"))
        value_label.pack()
        card.value_label = value_label

        title_label = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12), text_color="gray")
        title_label.pack(pady=(0, 15))

        return card

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

    def load_status(self):
        domains, _, error = get_unique_domains(CSV_INPUT_FILE)
        state = load_batch_state()

        if domains:
            processed = len(state.get('processed_domains', []))
            self.domains_card.value_label.configure(text=f"{processed} / {len(domains)}")
        else:
            self.domains_card.value_label.configure(text="No CSV")

        if os.path.exists(OUTPUT_CSV_FILE):
            with open(OUTPUT_CSV_FILE, 'r') as f:
                tests = len(list(csv.reader(f)))
            self.tests_card.value_label.configure(text=str(tests))
        else:
            self.tests_card.value_label.configure(text="0")

    def check_api_key(self):
        if not load_api_key():
            self.show_settings(first_run=True)

    def check_updates_async(self):
        def check():
            update_info = check_for_updates(silent=True)
            if update_info.get('available'):
                self.root.after(0, lambda: self.show_update_available(update_info))
        threading.Thread(target=check, daemon=True).start()

    def show_update_available(self, update_info):
        self.update_btn.configure(text=f"🔄 Update to v{update_info['latest']}")
        self.update_btn.pack(side="right")
        self.log(f"Update available: v{update_info['latest']}")

    def do_update(self):
        self.log("Downloading update...")
        success, message = download_update()
        if success:
            messagebox.showinfo("Update Complete", message)
            self.root.quit()
        else:
            messagebox.showerror("Update Failed", message)

    def run_tests(self):
        if self.running:
            return

        if not API_KEY:
            messagebox.showerror("Error", "Please configure your API key first")
            self.show_settings()
            return

        self.running = True
        self.run_btn.configure(state="disabled")
        self.status_card.value_label.configure(text="Running...")

        def run():
            try:
                state = load_batch_state()
                processed = set(state['processed_domains'])
                batch_num = state['batch_number'] + 1

                domains, domain_map, error = get_unique_domains(CSV_INPUT_FILE)
                if error:
                    self.root.after(0, lambda: self.log(f"Error: {error}"))
                    return

                remaining = [d for d in domains if d not in processed]
                if not remaining:
                    self.root.after(0, lambda: self.log("All domains processed!"))
                    return

                batch = remaining[:MAX_DOMAINS_PER_BATCH]
                self.root.after(0, lambda: self.log(f"Starting batch #{batch_num} with {len(batch)} domains"))

                session = create_api_session()
                ensure_data_dir()
                results = []

                for i, domain in enumerate(batch):
                    progress = (i + 1) / len(batch)
                    self.root.after(0, lambda p=progress: self.progress.set(p))
                    self.root.after(0, lambda d=domain: self.log(f"Processing: {d}"))

                    row = domain_map[domain]
                    test_name = f"Inbox Test - {domain} - {time.strftime('%Y-%m-%d %H:%M:%S')}"
                    test_result, error = create_test(session, test_name)

                    if not test_result or 'data' not in test_result:
                        self.root.after(0, lambda d=domain: self.log(f"  ❌ Failed to create test"))
                        continue

                    test_data = test_result['data']
                    test_uuid = test_data['uuid']
                    filter_phrase = test_data['filter_phrase']
                    test_emails = test_data['comma_separated_test_email_addresses']

                    success, err = send_email(
                        row.get('from_name', ''), row['from_email'], row['user_name'],
                        row['password'], row['smtp_host'], test_emails.replace(',', ';'), filter_phrase
                    )

                    if success:
                        self.root.after(0, lambda: self.log(f"  ✅ Email sent"))
                        results.append({
                            'domain': domain, 'from_email': row['from_email'],
                            'test_uuid': test_uuid, 'filter_phrase': filter_phrase,
                            'test_url': f"https://app.emailguard.io/inbox-placement-tests/{test_uuid}"
                        })
                    else:
                        self.root.after(0, lambda e=err: self.log(f"  ❌ Email failed: {e[:50]}"))

                    if i < len(batch) - 1:
                        time.sleep(EMAIL_DELAY_SECONDS)

                if results:
                    with open(OUTPUT_CSV_FILE, mode='a', newline='') as f:
                        writer = csv.writer(f)
                        for r in results:
                            writer.writerow([r['from_email'], r['test_uuid'], r['filter_phrase'], r['test_url']])

                state['processed_domains'] = list(processed) + [r['domain'] for r in results]
                state['batch_number'] = batch_num
                save_batch_state(state)
                session.close()

                self.root.after(0, lambda: self.log(f"Batch complete: {len(results)} successful"))
                self.root.after(0, self.load_status)

            finally:
                self.running = False
                self.root.after(0, lambda: self.run_btn.configure(state="normal"))
                self.root.after(0, lambda: self.status_card.value_label.configure(text="Ready"))
                self.root.after(0, lambda: self.progress.set(0))

        threading.Thread(target=run, daemon=True).start()

    def get_results(self):
        if self.running:
            return

        if not os.path.exists(OUTPUT_CSV_FILE):
            messagebox.showinfo("Info", "No tests found. Run tests first.")
            return

        self.running = True
        self.results_btn.configure(state="disabled")
        self.status_card.value_label.configure(text="Fetching...")

        def fetch():
            try:
                with open(OUTPUT_CSV_FILE, mode='r', encoding='utf-8-sig') as f:
                    tests = list(csv.reader(f))

                test_infos = [(row[0], row[1], row[3] if len(row) > 3 else '') for row in tests if len(row) >= 2]
                self.root.after(0, lambda: self.log(f"Fetching {len(test_infos)} results..."))

                session = create_api_session()
                all_results = []

                with ThreadPoolExecutor(max_workers=MAX_PARALLEL_WORKERS) as executor:
                    futures = {executor.submit(fetch_single_result, session, info): info for info in test_infos}
                    for i, future in enumerate(as_completed(futures)):
                        result = future.result()
                        all_results.append(result)
                        progress = (i + 1) / len(test_infos)
                        self.root.after(0, lambda p=progress: self.progress.set(p))

                session.close()

                completed = [r for r in all_results if r.get('status') in ['completed', 'complete']]
                pending = [r for r in all_results if r.get('status') not in ['completed', 'complete', 'FAILED']]

                self.root.after(0, lambda: self.log(f"Results: {len(completed)} completed, {len(pending)} pending"))

                # Generate outputs
                with open(RESULTS_CSV_FILE, mode='w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['from_email', 'status', 'inbox_rate_%', 'google_%', 'microsoft_%', 'spam_%'])
                    for r in all_results:
                        stats = r.get('stats', {})
                        writer.writerow([
                            r.get('from_email', ''), r.get('status', ''),
                            f"{stats.get('inbox_rate', 0):.1f}", f"{stats.get('google_inbox_rate', 0):.1f}",
                            f"{stats.get('microsoft_inbox_rate', 0):.1f}", f"{stats.get('spam_rate', 0):.1f}"
                        ])

                try:
                    generate_pdf(all_results)
                    self.root.after(0, lambda: self.log("PDF report generated!"))
                except Exception as e:
                    self.root.after(0, lambda: self.log(f"PDF error: {e}"))

                if completed:
                    avg_inbox = sum(r['stats']['inbox_rate'] for r in completed) / len(completed)
                    self.root.after(0, lambda: self.log(f"Average inbox rate: {avg_inbox:.1f}%"))

            finally:
                self.running = False
                self.root.after(0, lambda: self.results_btn.configure(state="normal"))
                self.root.after(0, lambda: self.status_card.value_label.configure(text="Ready"))
                self.root.after(0, lambda: self.progress.set(0))

        threading.Thread(target=fetch, daemon=True).start()

    def toggle_polling(self):
        if self.running:
            self.running = False
            self.poll_btn.configure(text="🔄  Auto-Poll")
            self.status_card.value_label.configure(text="Stopped")
            self.log("Polling stopped")
        else:
            self.running = True
            self.poll_btn.configure(text="⏹️  Stop Polling")
            self.status_card.value_label.configure(text="Polling...")
            self.log("Starting auto-poll...")

            def poll():
                while self.running:
                    if not os.path.exists(OUTPUT_CSV_FILE):
                        self.root.after(0, lambda: self.log("No tests to poll"))
                        break

                    with open(OUTPUT_CSV_FILE, mode='r', encoding='utf-8-sig') as f:
                        tests = list(csv.reader(f))

                    test_infos = [(row[0], row[1], row[3] if len(row) > 3 else '') for row in tests if len(row) >= 2]
                    session = create_api_session()
                    all_results = []

                    with ThreadPoolExecutor(max_workers=MAX_PARALLEL_WORKERS) as executor:
                        futures = {executor.submit(fetch_single_result, session, info): info for info in test_infos}
                        for future in as_completed(futures):
                            all_results.append(future.result())

                    session.close()

                    completed = len([r for r in all_results if r.get('status') in ['completed', 'complete']])
                    pending = len([r for r in all_results if r.get('status') not in ['completed', 'complete', 'FAILED']])

                    self.root.after(0, lambda c=completed, p=pending: self.log(f"Poll: {c} completed, {p} pending"))

                    if pending == 0:
                        self.root.after(0, lambda: self.log("All tests complete!"))
                        # Generate final report
                        with open(RESULTS_CSV_FILE, mode='w', newline='', encoding='utf-8') as f:
                            writer = csv.writer(f)
                            writer.writerow(['from_email', 'status', 'inbox_%', 'google_%', 'microsoft_%', 'spam_%'])
                            for r in all_results:
                                stats = r.get('stats', {})
                                writer.writerow([
                                    r.get('from_email', ''), r.get('status', ''),
                                    f"{stats.get('inbox_rate', 0):.1f}", f"{stats.get('google_inbox_rate', 0):.1f}",
                                    f"{stats.get('microsoft_inbox_rate', 0):.1f}", f"{stats.get('spam_rate', 0):.1f}"
                                ])
                        try:
                            generate_pdf(all_results)
                            self.root.after(0, lambda: self.log("PDF report generated!"))
                        except:
                            pass
                        break

                    for _ in range(POLL_INTERVAL_SECONDS):
                        if not self.running:
                            break
                        time.sleep(1)

                self.running = False
                self.root.after(0, lambda: self.poll_btn.configure(text="🔄  Auto-Poll"))
                self.root.after(0, lambda: self.status_card.value_label.configure(text="Ready"))

            threading.Thread(target=poll, daemon=True).start()

    def show_settings(self, first_run=False):
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Settings")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ctk.CTkFrame(dialog)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        if first_run:
            ctk.CTkLabel(frame, text="Welcome! Please configure your API key.",
                        font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(0, 20))

        ctk.CTkLabel(frame, text="EmailGuard API Key:", anchor="w").pack(fill="x")
        api_entry = ctk.CTkEntry(frame, width=400, placeholder_text="Enter your API key")
        api_entry.pack(fill="x", pady=(5, 10))
        if API_KEY:
            api_entry.insert(0, API_KEY)

        ctk.CTkLabel(frame, text="Get your API key from:", anchor="w").pack(fill="x")
        link = ctk.CTkLabel(frame, text="https://app.emailguard.io/settings/api",
                           text_color="#4a9eff", cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda e: os.system("open https://app.emailguard.io/settings/api"))

        def save():
            key = api_entry.get().strip()
            if key:
                save_api_key(key)
                self.log("API key saved")
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Please enter an API key")

        ctk.CTkButton(frame, text="Save", command=save).pack(pady=20)

    def reset_all(self):
        if messagebox.askyesno("Confirm Reset", "This will delete all test data. Continue?"):
            for f in [BATCH_STATE_FILE, OUTPUT_CSV_FILE]:
                if os.path.exists(f):
                    os.remove(f)
            if os.path.exists(DATA_DIR):
                try:
                    os.rmdir(DATA_DIR)
                except:
                    pass
            self.log("All data reset")
            self.load_status()

    def select_csv(self):
        global CSV_INPUT_FILE
        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file:
            CSV_INPUT_FILE = file
            self.log(f"CSV file: {os.path.basename(file)}")
            self.load_status()

    def open_report(self):
        if os.path.exists(PDF_REPORT_FILE):
            os.system(f"open '{PDF_REPORT_FILE}'")
        else:
            messagebox.showinfo("Info", "No report generated yet")

    def run(self):
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    if GUI_AVAILABLE:
        app = EmailGuardApp()
        app.run()
    else:
        print("=" * 60)
        print("  GUI requires CustomTkinter. Install with:")
        print("  pip3 install customtkinter")
        print("=" * 60)
        print("\nFalling back to terminal mode...")
        print("Run the terminal version or install the GUI package.")
        sys.exit(1)


if __name__ == "__main__":
    main()
