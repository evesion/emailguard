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

__version__ = "1.4.1"
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

SMTP_PORT = 465
EMAIL_SUBJECT = "Team Meeting Code"
EMAIL_BODY = """Hello Team, Please find here todays meeting code:"""

DEFAULT_BATCH_SIZE = 50
MAX_PARALLEL_WORKERS = 5
EMAIL_DELAY_SECONDS = 3
POLL_INTERVAL_SECONDS = 30

# Required CSV columns
REQUIRED_CSV_COLUMNS = ['from_email', 'user_name', 'password', 'smtp_host']

# ═══════════════════════════════════════════════════════════════════════════════
# INTERNAL SETTINGS
# ═══════════════════════════════════════════════════════════════════════════════

API_BASE_URL = "https://app.emailguard.io/api/v1"
DATA_DIR = '.emailguard_data'
CONFIG_FILE = '.env'
CUSTOMERS_FILE = os.path.join(DATA_DIR, 'customers.json')
SETTINGS_FILE = os.path.join(DATA_DIR, 'settings.json')

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

API_KEY = None


# ═══════════════════════════════════════════════════════════════════════════════
# SETTINGS MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def load_settings():
    ensure_data_dir()
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {'batch_size': DEFAULT_BATCH_SIZE}


def save_settings(settings):
    ensure_data_dir()
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# CSV VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

def validate_csv(csv_file):
    """
    Validate CSV file format and return detailed error messages.
    Returns: (is_valid, error_message, row_count, domain_count)
    """
    if not csv_file:
        return False, "No CSV file selected", 0, 0

    if not os.path.exists(csv_file):
        return False, f"File not found: {csv_file}", 0, 0

    try:
        with open(csv_file, mode='r', encoding='utf-8-sig') as file:
            # Check if file is empty
            content = file.read()
            if not content.strip():
                return False, "CSV file is empty", 0, 0

            file.seek(0)
            csv_reader = csv.DictReader(file)

            # Check headers
            if not csv_reader.fieldnames:
                return False, "CSV file has no headers", 0, 0

            headers = [h.strip().lower() for h in csv_reader.fieldnames]
            missing_columns = []
            for col in REQUIRED_CSV_COLUMNS:
                if col.lower() not in headers:
                    missing_columns.append(col)

            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}", 0, 0

            # Validate rows
            row_count = 0
            domains = set()
            errors = []

            for i, row in enumerate(csv_reader, start=2):  # Start at 2 (1 is header)
                row_count += 1

                # Check for empty required fields
                for col in REQUIRED_CSV_COLUMNS:
                    # Find the actual column name (case-insensitive)
                    actual_col = None
                    for h in csv_reader.fieldnames:
                        if h.strip().lower() == col.lower():
                            actual_col = h
                            break

                    if actual_col and not row.get(actual_col, '').strip():
                        if len(errors) < 5:  # Limit error messages
                            errors.append(f"Row {i}: Empty '{col}'")

                # Extract domain
                from_email_col = None
                for h in csv_reader.fieldnames:
                    if h.strip().lower() == 'from_email':
                        from_email_col = h
                        break

                if from_email_col and row.get(from_email_col):
                    email = row[from_email_col].strip()
                    if '@' in email:
                        domain = email.split('@')[-1]
                        domains.add(domain)
                    elif len(errors) < 5:
                        errors.append(f"Row {i}: Invalid email format")

            if row_count == 0:
                return False, "CSV file has no data rows", 0, 0

            if errors:
                error_msg = "CSV validation errors:\n• " + "\n• ".join(errors)
                if len(errors) >= 5:
                    error_msg += "\n• (more errors...)"
                return False, error_msg, row_count, len(domains)

            return True, None, row_count, len(domains)

    except UnicodeDecodeError:
        return False, "CSV file encoding error. Please save as UTF-8.", 0, 0
    except Exception as e:
        return False, f"Error reading CSV: {str(e)}", 0, 0


# ═══════════════════════════════════════════════════════════════════════════════
# CUSTOMER & BATCH MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def ensure_data_dir():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)


def load_customers():
    ensure_data_dir()
    if os.path.exists(CUSTOMERS_FILE):
        try:
            with open(CUSTOMERS_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {'customers': [], 'active_customer': None, 'active_batch': None}


def save_customers(data):
    ensure_data_dir()
    with open(CUSTOMERS_FILE, 'w') as f:
        json.dump(data, f, indent=2)


def get_customer_dir(customer_name):
    safe_name = "".join(c for c in customer_name if c.isalnum() or c in (' ', '-', '_')).strip()
    return os.path.join(DATA_DIR, safe_name)


def get_batch_dir(customer_name, batch_name):
    safe_batch = "".join(c for c in batch_name if c.isalnum() or c in (' ', '-', '_')).strip()
    return os.path.join(get_customer_dir(customer_name), safe_batch)


def create_customer(name):
    data = load_customers()
    if name not in data['customers']:
        data['customers'].append(name)
        customer_dir = get_customer_dir(name)
        if not os.path.exists(customer_dir):
            os.makedirs(customer_dir)
        # Initialize customer state
        customer_state = {'batches': [], 'created': datetime.now().isoformat()}
        with open(os.path.join(customer_dir, 'customer_state.json'), 'w') as f:
            json.dump(customer_state, f, indent=2)
    data['active_customer'] = name
    data['active_batch'] = None
    save_customers(data)
    return True


def delete_customer(customer_name):
    """Delete a customer and all their data"""
    data = load_customers()

    if customer_name not in data['customers']:
        return False, "Customer not found"

    # Remove customer directory
    customer_dir = get_customer_dir(customer_name)
    if os.path.exists(customer_dir):
        shutil.rmtree(customer_dir)

    # Update customers list
    data['customers'].remove(customer_name)

    # Clear active customer if it was deleted
    if data['active_customer'] == customer_name:
        data['active_customer'] = data['customers'][0] if data['customers'] else None
        data['active_batch'] = None

    save_customers(data)
    return True, "Customer deleted"


def delete_batch(customer_name, batch_name):
    """Delete a batch and all its data"""
    customer_dir = get_customer_dir(customer_name)
    batch_dir = get_batch_dir(customer_name, batch_name)

    # Remove batch directory
    if os.path.exists(batch_dir):
        shutil.rmtree(batch_dir)

    # Update customer state
    state_file = os.path.join(customer_dir, 'customer_state.json')
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            customer_state = json.load(f)

        if batch_name in customer_state.get('batches', []):
            customer_state['batches'].remove(batch_name)
            with open(state_file, 'w') as f:
                json.dump(customer_state, f, indent=2)

    # Update active batch if it was deleted
    data = load_customers()
    if data['active_batch'] == batch_name and data['active_customer'] == customer_name:
        batches = get_customer_batches(customer_name)
        data['active_batch'] = batches[0] if batches else None
        save_customers(data)

    return True, "Batch deleted"


def create_batch(customer_name, batch_name):
    customer_dir = get_customer_dir(customer_name)
    batch_dir = get_batch_dir(customer_name, batch_name)

    # Load customer state
    state_file = os.path.join(customer_dir, 'customer_state.json')
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            customer_state = json.load(f)
    else:
        customer_state = {'batches': [], 'created': datetime.now().isoformat()}

    # Add batch if not exists
    if batch_name not in customer_state['batches']:
        customer_state['batches'].append(batch_name)
        with open(state_file, 'w') as f:
            json.dump(customer_state, f, indent=2)

    # Create batch directory
    if not os.path.exists(batch_dir):
        os.makedirs(batch_dir)

    # Initialize batch state
    batch_state_file = os.path.join(batch_dir, 'batch_state.json')
    if not os.path.exists(batch_state_file):
        batch_state = {
            'name': batch_name,
            'processed_domains': [],
            'created': datetime.now().isoformat(),
            'csv_file': None
        }
        with open(batch_state_file, 'w') as f:
            json.dump(batch_state, f, indent=2)

    # Update active batch
    data = load_customers()
    data['active_batch'] = batch_name
    save_customers(data)

    return True


def get_customer_batches(customer_name):
    customer_dir = get_customer_dir(customer_name)
    state_file = os.path.join(customer_dir, 'customer_state.json')
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            return json.load(f).get('batches', [])
    return []


def get_batch_state(customer_name, batch_name):
    batch_dir = get_batch_dir(customer_name, batch_name)
    state_file = os.path.join(batch_dir, 'batch_state.json')
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            return json.load(f)
    return {'name': batch_name, 'processed_domains': [], 'csv_file': None}


def save_batch_state(customer_name, batch_name, state):
    batch_dir = get_batch_dir(customer_name, batch_name)
    if not os.path.exists(batch_dir):
        os.makedirs(batch_dir)
    state_file = os.path.join(batch_dir, 'batch_state.json')
    with open(state_file, 'w') as f:
        json.dump(state, f, indent=2)


def get_batch_files(customer_name, batch_name):
    batch_dir = get_batch_dir(customer_name, batch_name)
    return {
        'dir': batch_dir,
        'state': os.path.join(batch_dir, 'batch_state.json'),
        'queue': os.path.join(batch_dir, 'test_queue.csv'),
        'results': os.path.join(batch_dir, 'results.csv'),
        'report': os.path.join(batch_dir, 'report.pdf')
    }


def get_customer_combined_report(customer_name):
    return os.path.join(get_customer_dir(customer_name), 'combined_report.pdf')


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


def create_api_session():
    session = requests.Session()
    retry_strategy = Retry(total=3, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=MAX_PARALLEL_WORKERS, pool_maxsize=MAX_PARALLEL_WORKERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"})
    return session


def get_unique_domains(csv_file):
    domains = []
    domain_to_row = {}
    try:
        with open(csv_file, mode='r', encoding='utf-8-sig') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                # Case-insensitive column lookup
                from_email = None
                for key in row.keys():
                    if key.strip().lower() == 'from_email':
                        from_email = row[key]
                        break

                if from_email and '@' in from_email:
                    domain = from_email.split('@')[-1]
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


def generate_pdf(all_results, output_file, title="Email Inbox Placement Report", subtitle=None):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.enums import TA_CENTER

    doc = SimpleDocTemplate(output_file, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=24, spaceAfter=10,
                                  alignment=TA_CENTER, textColor=colors.Color(0.2, 0.3, 0.5))
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'], fontSize=14,
                                     alignment=TA_CENTER, textColor=colors.Color(0.3, 0.4, 0.5), spaceAfter=5)
    date_style = ParagraphStyle('Date', parent=styles['Normal'], fontSize=12,
                                 alignment=TA_CENTER, textColor=colors.gray, spaceAfter=20)
    section_style = ParagraphStyle('Section', parent=styles['Heading2'], fontSize=14,
                                    spaceBefore=20, spaceAfter=10, textColor=colors.Color(0.2, 0.3, 0.5))

    story = []
    story.append(Paragraph(title, title_style))
    if subtitle:
        story.append(Paragraph(subtitle, subtitle_style))
    story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %H:%M')}", date_style))
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


def generate_combined_pdf(customer_name, batches_data, output_file):
    """Generate a combined report for all batches of a customer"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.enums import TA_CENTER
    from reportlab.graphics.shapes import Drawing, Line, String, Rect
    from reportlab.graphics.charts.lineplots import LinePlot
    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics.widgets.markers import makeMarker

    doc = SimpleDocTemplate(output_file, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=24, spaceAfter=10,
                                  alignment=TA_CENTER, textColor=colors.Color(0.2, 0.3, 0.5))
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'], fontSize=16,
                                     alignment=TA_CENTER, textColor=colors.Color(0.3, 0.4, 0.5), spaceAfter=5)
    date_style = ParagraphStyle('Date', parent=styles['Normal'], fontSize=12,
                                 alignment=TA_CENTER, textColor=colors.gray, spaceAfter=20)
    section_style = ParagraphStyle('Section', parent=styles['Heading2'], fontSize=14,
                                    spaceBefore=20, spaceAfter=10, textColor=colors.Color(0.2, 0.3, 0.5))
    batch_title_style = ParagraphStyle('BatchTitle', parent=styles['Heading2'], fontSize=16,
                                        spaceBefore=10, spaceAfter=15, textColor=colors.Color(0.2, 0.4, 0.6))

    story = []

    # Title page
    story.append(Paragraph("Combined Inbox Placement Report", title_style))
    story.append(Paragraph(f"Customer: {customer_name}", subtitle_style))
    story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %H:%M')}", date_style))
    story.append(Spacer(1, 30))

    # Summary across all batches
    all_results = []
    for batch_name, results in batches_data.items():
        all_results.extend(results)

    completed = [r for r in all_results if r.get('status') in ['completed', 'complete']]

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

        story.append(Paragraph("Overall Summary (All Batches)", section_style))

        # Batch summary table - also collect data for chart
        batch_summary = [['Batch', 'Tests', 'Inbox %', 'Google %', 'Microsoft %']]
        chart_data = {'names': [], 'inbox': [], 'google': [], 'microsoft': []}

        for batch_name, results in batches_data.items():
            batch_completed = [r for r in results if r.get('status') in ['completed', 'complete']]
            if batch_completed:
                b_inbox = sum(r['stats']['inbox_rate'] for r in batch_completed) / len(batch_completed)
                b_google_tests = [r for r in batch_completed if r['stats']['google_total'] > 0]
                b_microsoft_tests = [r for r in batch_completed if r['stats']['microsoft_total'] > 0]
                b_google = sum(r['stats']['google_inbox_rate'] for r in b_google_tests) / len(b_google_tests) if b_google_tests else 0
                b_microsoft = sum(r['stats']['microsoft_inbox_rate'] for r in b_microsoft_tests) / len(b_microsoft_tests) if b_microsoft_tests else 0
                batch_summary.append([batch_name[:20], str(len(results)), f"{b_inbox:.0f}%", f"{b_google:.0f}%", f"{b_microsoft:.0f}%"])

                # Store for chart
                chart_data['names'].append(batch_name[:15] if len(batch_name) > 15 else batch_name)
                chart_data['inbox'].append(b_inbox)
                chart_data['google'].append(b_google)
                chart_data['microsoft'].append(b_microsoft)

        if len(batch_summary) > 1:
            batch_table = Table(batch_summary, colWidths=[140, 60, 80, 80, 90])
            batch_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.3, 0.5)),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.8, 0.8, 0.8)),
                ('BOX', (0, 0), (-1, -1), 1, colors.Color(0.2, 0.3, 0.5)),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            story.append(batch_table)
            story.append(Spacer(1, 20))

        # Trend Chart - only show if we have 2+ batches
        if len(chart_data['names']) >= 2:
            story.append(Paragraph("Inbox Rate Trend", section_style))

            # Create the chart
            drawing = Drawing(500, 250)

            # Add background
            drawing.add(Rect(40, 30, 420, 180, fillColor=colors.Color(0.98, 0.98, 0.98),
                            strokeColor=colors.Color(0.8, 0.8, 0.8), strokeWidth=1))

            # Create line plot
            lp = LinePlot()
            lp.x = 60
            lp.y = 50
            lp.width = 380
            lp.height = 150

            # Prepare data points - (x, y) tuples for each line
            num_batches = len(chart_data['names'])
            inbox_data = [(i, chart_data['inbox'][i]) for i in range(num_batches)]
            google_data = [(i, chart_data['google'][i]) for i in range(num_batches)]
            microsoft_data = [(i, chart_data['microsoft'][i]) for i in range(num_batches)]

            lp.data = [inbox_data, google_data, microsoft_data]

            # Style the lines
            lp.lines[0].strokeColor = colors.Color(0.2, 0.4, 0.8)  # Blue - Overall
            lp.lines[0].strokeWidth = 3
            lp.lines[0].symbol = makeMarker('FilledCircle')
            lp.lines[0].symbol.fillColor = colors.Color(0.2, 0.4, 0.8)
            lp.lines[0].symbol.size = 8

            lp.lines[1].strokeColor = colors.Color(0.85, 0.2, 0.2)  # Red - Google
            lp.lines[1].strokeWidth = 2
            lp.lines[1].symbol = makeMarker('FilledSquare')
            lp.lines[1].symbol.fillColor = colors.Color(0.85, 0.2, 0.2)
            lp.lines[1].symbol.size = 6

            lp.lines[2].strokeColor = colors.Color(0.1, 0.6, 0.2)  # Green - Microsoft
            lp.lines[2].strokeWidth = 2
            lp.lines[2].symbol = makeMarker('FilledDiamond')
            lp.lines[2].symbol.fillColor = colors.Color(0.1, 0.6, 0.2)
            lp.lines[2].symbol.size = 6

            # Configure axes
            lp.xValueAxis.valueMin = -0.2
            lp.xValueAxis.valueMax = num_batches - 0.8
            lp.xValueAxis.valueSteps = list(range(num_batches))
            lp.xValueAxis.labelTextFormat = lambda x: chart_data['names'][int(x)] if 0 <= int(x) < num_batches else ''
            lp.xValueAxis.labels.fontName = 'Helvetica'
            lp.xValueAxis.labels.fontSize = 8
            lp.xValueAxis.labels.angle = 0

            lp.yValueAxis.valueMin = 0
            lp.yValueAxis.valueMax = 100
            lp.yValueAxis.valueSteps = [0, 25, 50, 75, 100]
            lp.yValueAxis.labelTextFormat = '%d%%'
            lp.yValueAxis.labels.fontName = 'Helvetica'
            lp.yValueAxis.labels.fontSize = 9

            drawing.add(lp)

            # Add legend
            legend = Legend()
            legend.x = 150
            legend.y = 15
            legend.dx = 8
            legend.dy = 8
            legend.fontName = 'Helvetica'
            legend.fontSize = 9
            legend.boxAnchor = 'c'
            legend.columnMaximum = 1
            legend.strokeWidth = 0
            legend.strokeColor = None
            legend.deltax = 100
            legend.deltay = 0
            legend.autoXPadding = 10
            legend.alignment = 'right'
            legend.dividerLines = 0
            legend.colorNamePairs = [
                (colors.Color(0.2, 0.4, 0.8), 'Overall Inbox'),
                (colors.Color(0.85, 0.2, 0.2), 'Google'),
                (colors.Color(0.1, 0.6, 0.2), 'Microsoft'),
            ]
            drawing.add(legend)

            # Add title
            drawing.add(String(250, 235, 'Inbox Placement Rate by Batch',
                              fontSize=12, fontName='Helvetica-Bold',
                              fillColor=colors.Color(0.2, 0.3, 0.5), textAnchor='middle'))

            story.append(drawing)
            story.append(Spacer(1, 20))

            # Add trend analysis text
            first_inbox = chart_data['inbox'][0]
            last_inbox = chart_data['inbox'][-1]
            change = last_inbox - first_inbox

            if change > 5:
                trend_text = f"📈 <b>Improving:</b> Inbox rate increased by {change:.1f}% from first to last batch"
                trend_color = colors.Color(0.1, 0.5, 0.1)
            elif change < -5:
                trend_text = f"📉 <b>Declining:</b> Inbox rate decreased by {abs(change):.1f}% from first to last batch"
                trend_color = colors.Color(0.7, 0.1, 0.1)
            else:
                trend_text = f"➡️ <b>Stable:</b> Inbox rate remained relatively stable (±{abs(change):.1f}%)"
                trend_color = colors.Color(0.4, 0.4, 0.4)

            trend_style = ParagraphStyle('Trend', parent=styles['Normal'], fontSize=11,
                                         textColor=trend_color, spaceAfter=20)
            story.append(Paragraph(trend_text, trend_style))

        # Overall metrics
        story.append(Paragraph("Combined Performance", section_style))
        metrics_data = [['Total Inbox Rate', 'Google', 'Microsoft', 'Spam Rate'],
                        [f'{avg_inbox:.1f}%', f'{avg_google:.1f}%', f'{avg_microsoft:.1f}%', f'{avg_spam:.1f}%']]
        metrics_table = Table(metrics_data, colWidths=[120, 120, 120, 120])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.3, 0.4, 0.6)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 18),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 1), (0, 1), get_color(avg_inbox)),
            ('TEXTCOLOR', (1, 1), (1, 1), get_color(avg_google)),
            ('TEXTCOLOR', (2, 1), (2, 1), get_color(avg_microsoft)),
            ('TEXTCOLOR', (3, 1), (3, 1), get_color(avg_spam, is_spam=True)),
            ('BACKGROUND', (0, 1), (-1, 1), colors.Color(0.98, 0.98, 0.98)),
            ('BOX', (0, 0), (-1, -1), 2, colors.Color(0.3, 0.4, 0.6)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(metrics_table)

    # Individual batch details
    for batch_name, results in batches_data.items():
        story.append(PageBreak())
        story.append(Paragraph(f"Batch: {batch_name}", batch_title_style))

        batch_completed = [r for r in results if r.get('status') in ['completed', 'complete']]
        if batch_completed:
            b_inbox = sum(r['stats']['inbox_rate'] for r in batch_completed) / len(batch_completed)
            story.append(Paragraph(f"Tests: {len(results)} | Average Inbox Rate: {b_inbox:.1f}%", date_style))

        # Detailed results table
        table_data = [['Domain', 'Inbox %', 'Google %', 'Microsoft %', 'Spam %', 'Status']]
        for r in results:
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
    from tkinter import filedialog, messagebox, simpledialog
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False


class EmailGuardApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title(f"EmailGuard Inbox Placement Tester v{__version__}")
        self.root.geometry("950x800")
        self.root.minsize(900, 750)

        # Set theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.session = None
        self.running = False
        self.poll_thread = None

        # Current selection
        self.current_customer = None
        self.current_batch = None
        self.current_csv = None

        # Load settings
        self.settings = load_settings()
        self.batch_size = self.settings.get('batch_size', DEFAULT_BATCH_SIZE)

        self.setup_ui()
        self.load_saved_state()
        self.check_api_key()

        # Check for updates
        self.root.after(1000, self.check_updates_async)

    def setup_ui(self):
        # Main container
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 15))

        title_label = ctk.CTkLabel(header_frame, text="📧 EmailGuard", font=ctk.CTkFont(size=28, weight="bold"))
        title_label.pack(side="left")

        version_label = ctk.CTkLabel(header_frame, text=f"v{__version__}", font=ctk.CTkFont(size=14), text_color="gray")
        version_label.pack(side="left", padx=(10, 0), pady=(10, 0))

        self.update_btn = ctk.CTkButton(header_frame, text="🔄 Update Available", width=140,
                                         fg_color="#2d5a27", hover_color="#3d7a37", command=self.do_update)
        self.update_btn.pack(side="right")
        self.update_btn.pack_forget()  # Hidden initially

        # Customer & Batch Selection Frame
        selection_frame = ctk.CTkFrame(self.main_frame)
        selection_frame.pack(fill="x", pady=(0, 15))

        # Customer selection
        customer_frame = ctk.CTkFrame(selection_frame, fg_color="transparent")
        customer_frame.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)

        ctk.CTkLabel(customer_frame, text="👤 Customer:", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        self.customer_dropdown = ctk.CTkOptionMenu(customer_frame, width=150, values=["No customers"],
                                                    command=self.on_customer_change)
        self.customer_dropdown.pack(side="left", padx=(10, 5))
        ctk.CTkButton(customer_frame, text="+", width=30, command=self.new_customer).pack(side="left")
        ctk.CTkButton(customer_frame, text="🗑", width=30, fg_color="#8B0000", hover_color="#A52A2A",
                      command=self.delete_customer_dialog).pack(side="left", padx=(5, 0))

        # Batch selection
        batch_frame = ctk.CTkFrame(selection_frame, fg_color="transparent")
        batch_frame.pack(side="left", fill="x", expand=True, padx=(5, 10), pady=10)

        ctk.CTkLabel(batch_frame, text="📦 Batch:", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        self.batch_dropdown = ctk.CTkOptionMenu(batch_frame, width=150, values=["No batches"],
                                                 command=self.on_batch_change)
        self.batch_dropdown.pack(side="left", padx=(10, 5))
        ctk.CTkButton(batch_frame, text="+", width=30, command=self.new_batch).pack(side="left")
        ctk.CTkButton(batch_frame, text="🗑", width=30, fg_color="#8B0000", hover_color="#A52A2A",
                      command=self.delete_batch_dialog).pack(side="left", padx=(5, 0))

        # Batch size slider frame
        slider_frame = ctk.CTkFrame(self.main_frame)
        slider_frame.pack(fill="x", pady=(0, 15))

        slider_inner = ctk.CTkFrame(slider_frame, fg_color="transparent")
        slider_inner.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(slider_inner, text="📊 Batch Size:", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")

        self.batch_size_label = ctk.CTkLabel(slider_inner, text=f"{self.batch_size} domains",
                                              font=ctk.CTkFont(size=13), width=100)
        self.batch_size_label.pack(side="right")

        self.batch_size_slider = ctk.CTkSlider(slider_inner, from_=10, to=100, number_of_steps=9,
                                                command=self.on_batch_size_change, width=300)
        self.batch_size_slider.set(self.batch_size)
        self.batch_size_slider.pack(side="right", padx=(20, 10))

        ctk.CTkLabel(slider_inner, text="10", font=ctk.CTkFont(size=11), text_color="gray").pack(side="right")

        # Status cards
        cards_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        cards_frame.pack(fill="x", pady=(0, 15))

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
        buttons_frame.pack(fill="x", pady=(0, 15))

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

        self.log_text = ctk.CTkTextbox(self.main_frame, height=180, font=ctk.CTkFont(family="Courier", size=12))
        self.log_text.pack(fill="both", expand=True, pady=(5, 15))

        # Bottom buttons - Row 1
        bottom_frame1 = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bottom_frame1.pack(fill="x", pady=(0, 10))

        self.csv_btn = ctk.CTkButton(bottom_frame1, text="📁 Select CSV", width=130, command=self.select_csv)
        self.csv_btn.pack(side="left")

        self.batch_report_btn = ctk.CTkButton(bottom_frame1, text="📄 Batch Report", width=130, command=self.open_batch_report)
        self.batch_report_btn.pack(side="left", padx=(10, 0))

        self.combined_report_btn = ctk.CTkButton(bottom_frame1, text="📑 Combined Report", width=140,
                                                  fg_color="#1a5f1a", hover_color="#2d7a2d", command=self.generate_combined_report)
        self.combined_report_btn.pack(side="left", padx=(10, 0))

        self.settings_btn = ctk.CTkButton(bottom_frame1, text="⚙️ Settings", width=100, command=self.show_settings)
        self.settings_btn.pack(side="right")

        self.reset_btn = ctk.CTkButton(bottom_frame1, text="🔄 Reset Batch", width=120,
                                        fg_color="#555555", hover_color="#666666", command=self.reset_batch)
        self.reset_btn.pack(side="right", padx=(0, 10))

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

    def on_batch_size_change(self, value):
        self.batch_size = int(value)
        self.batch_size_label.configure(text=f"{self.batch_size} domains")
        # Save to settings
        self.settings['batch_size'] = self.batch_size
        save_settings(self.settings)

    def load_saved_state(self):
        data = load_customers()
        customers = data.get('customers', [])

        if customers:
            self.customer_dropdown.configure(values=customers)
            active = data.get('active_customer')
            if active and active in customers:
                self.customer_dropdown.set(active)
                self.current_customer = active
                self.load_batches_for_customer(active)

                # Load active batch
                active_batch = data.get('active_batch')
                batches = get_customer_batches(active)
                if active_batch and active_batch in batches:
                    self.batch_dropdown.set(active_batch)
                    self.current_batch = active_batch
                    self.load_batch_status()
            else:
                self.customer_dropdown.set(customers[0])
                self.current_customer = customers[0]
                self.load_batches_for_customer(customers[0])
        else:
            self.customer_dropdown.configure(values=["No customers"])
            self.customer_dropdown.set("No customers")
            self.batch_dropdown.configure(values=["No batches"])
            self.batch_dropdown.set("No batches")

    def load_batches_for_customer(self, customer_name):
        batches = get_customer_batches(customer_name)
        if batches:
            self.batch_dropdown.configure(values=batches)
            self.batch_dropdown.set(batches[0])
            self.current_batch = batches[0]
            self.load_batch_status()
        else:
            self.batch_dropdown.configure(values=["No batches"])
            self.batch_dropdown.set("No batches")
            self.current_batch = None
            self.domains_card.value_label.configure(text="No batch")
            self.tests_card.value_label.configure(text="0")

    def load_batch_status(self):
        if not self.current_customer or not self.current_batch:
            return

        files = get_batch_files(self.current_customer, self.current_batch)
        state = get_batch_state(self.current_customer, self.current_batch)

        csv_file = state.get('csv_file')
        if csv_file and os.path.exists(csv_file):
            self.current_csv = csv_file
            domains, _, error = get_unique_domains(csv_file)
            if domains:
                processed = len(state.get('processed_domains', []))
                self.domains_card.value_label.configure(text=f"{processed} / {len(domains)}")
            else:
                self.domains_card.value_label.configure(text="CSV Error")
        else:
            self.current_csv = None
            self.domains_card.value_label.configure(text="No CSV")

        if os.path.exists(files['queue']):
            with open(files['queue'], 'r') as f:
                tests = len(list(csv.reader(f)))
            self.tests_card.value_label.configure(text=str(tests))
        else:
            self.tests_card.value_label.configure(text="0")

    def on_customer_change(self, customer_name):
        if customer_name == "No customers":
            return
        self.current_customer = customer_name
        data = load_customers()
        data['active_customer'] = customer_name
        save_customers(data)
        self.load_batches_for_customer(customer_name)
        self.log(f"Switched to customer: {customer_name}")

    def on_batch_change(self, batch_name):
        if batch_name == "No batches":
            return
        self.current_batch = batch_name
        data = load_customers()
        data['active_batch'] = batch_name
        save_customers(data)
        self.load_batch_status()
        self.log(f"Switched to batch: {batch_name}")

    def new_customer(self):
        dialog = ctk.CTkInputDialog(text="Enter customer name:", title="New Customer")
        name = dialog.get_input()
        if name and name.strip():
            name = name.strip()
            create_customer(name)
            data = load_customers()
            self.customer_dropdown.configure(values=data['customers'])
            self.customer_dropdown.set(name)
            self.current_customer = name
            self.load_batches_for_customer(name)
            self.log(f"Created customer: {name}")

    def delete_customer_dialog(self):
        if not self.current_customer or self.current_customer == "No customers":
            messagebox.showinfo("Info", "No customer selected")
            return

        if messagebox.askyesno("Delete Customer",
                               f"Are you sure you want to delete '{self.current_customer}' and ALL their batches?\n\nThis cannot be undone!"):
            success, msg = delete_customer(self.current_customer)
            if success:
                self.log(f"Deleted customer: {self.current_customer}")
                self.load_saved_state()
            else:
                messagebox.showerror("Error", msg)

    def new_batch(self):
        if not self.current_customer or self.current_customer == "No customers":
            messagebox.showerror("Error", "Please create a customer first")
            return

        dialog = ctk.CTkInputDialog(text="Enter batch name:", title="New Batch")
        name = dialog.get_input()
        if name and name.strip():
            name = name.strip()
            create_batch(self.current_customer, name)
            batches = get_customer_batches(self.current_customer)
            self.batch_dropdown.configure(values=batches)
            self.batch_dropdown.set(name)
            self.current_batch = name
            self.load_batch_status()
            self.log(f"Created batch: {name}")

    def delete_batch_dialog(self):
        if not self.current_batch or self.current_batch == "No batches":
            messagebox.showinfo("Info", "No batch selected")
            return

        if messagebox.askyesno("Delete Batch",
                               f"Are you sure you want to delete batch '{self.current_batch}'?\n\nThis will delete all test data and cannot be undone!"):
            success, msg = delete_batch(self.current_customer, self.current_batch)
            if success:
                self.log(f"Deleted batch: {self.current_batch}")
                self.load_batches_for_customer(self.current_customer)
            else:
                messagebox.showerror("Error", msg)

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

        if not self.current_customer or self.current_customer == "No customers":
            messagebox.showerror("Error", "Please create a customer first")
            return

        if not self.current_batch or self.current_batch == "No batches":
            messagebox.showerror("Error", "Please create a batch first")
            return

        if not self.current_csv:
            messagebox.showerror("Error", "Please select a CSV file first")
            return

        if not API_KEY:
            messagebox.showerror("Error", "Please configure your API key first")
            self.show_settings()
            return

        # Validate CSV before running
        is_valid, error_msg, row_count, domain_count = validate_csv(self.current_csv)
        if not is_valid:
            messagebox.showerror("CSV Validation Error", error_msg)
            return

        self.log(f"CSV validated: {row_count} rows, {domain_count} unique domains")

        self.running = True
        self.run_btn.configure(state="disabled")
        self.status_card.value_label.configure(text="Running...")

        def run():
            try:
                files = get_batch_files(self.current_customer, self.current_batch)
                state = get_batch_state(self.current_customer, self.current_batch)
                processed = set(state.get('processed_domains', []))

                domains, domain_map, error = get_unique_domains(self.current_csv)
                if error:
                    self.root.after(0, lambda: self.log(f"Error: {error}"))
                    return

                remaining = [d for d in domains if d not in processed]
                if not remaining:
                    self.root.after(0, lambda: self.log("All domains processed!"))
                    return

                batch = remaining[:self.batch_size]
                self.root.after(0, lambda: self.log(f"Starting with {len(batch)} domains (batch size: {self.batch_size})"))

                session = create_api_session()
                results = []

                for i, domain in enumerate(batch):
                    progress = (i + 1) / len(batch)
                    self.root.after(0, lambda p=progress: self.progress.set(p))
                    self.root.after(0, lambda d=domain: self.log(f"Processing: {d}"))

                    row = domain_map[domain]

                    # Case-insensitive column access
                    def get_col(row, col_name):
                        for key in row.keys():
                            if key.strip().lower() == col_name.lower():
                                return row[key]
                        return ''

                    test_name = f"{self.current_customer} - {self.current_batch} - {domain}"
                    test_result, error = create_test(session, test_name)

                    if not test_result or 'data' not in test_result:
                        self.root.after(0, lambda d=domain: self.log(f"  ❌ Failed to create test"))
                        continue

                    test_data = test_result['data']
                    test_uuid = test_data['uuid']
                    filter_phrase = test_data['filter_phrase']
                    test_emails = test_data['comma_separated_test_email_addresses']

                    success, err = send_email(
                        get_col(row, 'from_name'), get_col(row, 'from_email'),
                        get_col(row, 'user_name'), get_col(row, 'password'),
                        get_col(row, 'smtp_host'), test_emails.replace(',', ';'), filter_phrase
                    )

                    if success:
                        self.root.after(0, lambda: self.log(f"  ✅ Email sent"))
                        results.append({
                            'domain': domain, 'from_email': get_col(row, 'from_email'),
                            'test_uuid': test_uuid, 'filter_phrase': filter_phrase,
                            'test_url': f"https://app.emailguard.io/inbox-placement-tests/{test_uuid}"
                        })
                    else:
                        self.root.after(0, lambda e=err: self.log(f"  ❌ Email failed: {e[:50]}"))

                    if i < len(batch) - 1:
                        time.sleep(EMAIL_DELAY_SECONDS)

                if results:
                    with open(files['queue'], mode='a', newline='') as f:
                        writer = csv.writer(f)
                        for r in results:
                            writer.writerow([r['from_email'], r['test_uuid'], r['filter_phrase'], r['test_url']])

                state['processed_domains'] = list(processed) + [r['domain'] for r in results]
                save_batch_state(self.current_customer, self.current_batch, state)
                session.close()

                self.root.after(0, lambda: self.log(f"Complete: {len(results)} successful"))
                self.root.after(0, self.load_batch_status)

            finally:
                self.running = False
                self.root.after(0, lambda: self.run_btn.configure(state="normal"))
                self.root.after(0, lambda: self.status_card.value_label.configure(text="Ready"))
                self.root.after(0, lambda: self.progress.set(0))

        threading.Thread(target=run, daemon=True).start()

    def get_results(self):
        if self.running:
            return

        if not self.current_customer or not self.current_batch:
            messagebox.showinfo("Info", "Please select a customer and batch first")
            return

        files = get_batch_files(self.current_customer, self.current_batch)
        if not os.path.exists(files['queue']):
            messagebox.showinfo("Info", "No tests found. Run tests first.")
            return

        self.running = True
        self.results_btn.configure(state="disabled")
        self.status_card.value_label.configure(text="Fetching...")

        def fetch():
            try:
                with open(files['queue'], mode='r', encoding='utf-8-sig') as f:
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
                with open(files['results'], mode='w', newline='', encoding='utf-8') as f:
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
                    generate_pdf(all_results, files['report'],
                                title=f"Inbox Placement Report",
                                subtitle=f"{self.current_customer} - {self.current_batch}")
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
            if not self.current_customer or not self.current_batch:
                messagebox.showinfo("Info", "Please select a customer and batch first")
                return

            files = get_batch_files(self.current_customer, self.current_batch)
            if not os.path.exists(files['queue']):
                messagebox.showinfo("Info", "No tests to poll. Run tests first.")
                return

            self.running = True
            self.poll_btn.configure(text="⏹️  Stop Polling")
            self.status_card.value_label.configure(text="Polling...")
            self.log("Starting auto-poll...")

            def poll():
                while self.running:
                    files = get_batch_files(self.current_customer, self.current_batch)
                    if not os.path.exists(files['queue']):
                        self.root.after(0, lambda: self.log("No tests to poll"))
                        break

                    with open(files['queue'], mode='r', encoding='utf-8-sig') as f:
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
                        with open(files['results'], mode='w', newline='', encoding='utf-8') as f:
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
                            generate_pdf(all_results, files['report'],
                                        title=f"Inbox Placement Report",
                                        subtitle=f"{self.current_customer} - {self.current_batch}")
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
        dialog.geometry("500x380")
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

        ctk.CTkButton(frame, text="Save", command=save).pack(pady=(20, 10))

        # Separator
        separator = ctk.CTkFrame(frame, height=2, fg_color="gray")
        separator.pack(fill="x", pady=15)

        # Update section
        update_frame = ctk.CTkFrame(frame, fg_color="transparent")
        update_frame.pack(fill="x")

        ctk.CTkLabel(update_frame, text="Software Updates", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(update_frame, text=f"Current version: v{__version__}", text_color="gray").pack(anchor="w", pady=(2, 8))

        update_status_label = ctk.CTkLabel(update_frame, text="", text_color="gray")
        update_status_label.pack(anchor="w")

        def check_updates_clicked():
            update_status_label.configure(text="Checking for updates...", text_color="gray")
            dialog.update()

            update_info = check_for_updates()
            if update_info.get('available'):
                update_status_label.configure(
                    text=f"✅ Update available: v{update_info['latest']}",
                    text_color="#2d7a2d"
                )
                # Show update button
                update_btn = ctk.CTkButton(update_frame, text=f"Download v{update_info['latest']}",
                                           fg_color="#2d5a27", hover_color="#3d7a37",
                                           command=lambda: self.do_update_from_settings(dialog))
                update_btn.pack(anchor="w", pady=(5, 0))
            else:
                update_status_label.configure(
                    text="✓ You're running the latest version",
                    text_color="gray"
                )

        ctk.CTkButton(update_frame, text="🔄 Check for Updates", width=160,
                      command=check_updates_clicked).pack(anchor="w")

    def do_update_from_settings(self, settings_dialog):
        settings_dialog.destroy()
        self.log("Downloading update...")
        success, message = download_update()
        if success:
            messagebox.showinfo("Update Complete", message)
            self.root.quit()
        else:
            messagebox.showerror("Update Failed", message)

    def reset_batch(self):
        if not self.current_customer or not self.current_batch:
            return

        if messagebox.askyesno("Confirm Reset", f"This will reset progress for batch '{self.current_batch}' (keeps test data). Continue?"):
            files = get_batch_files(self.current_customer, self.current_batch)
            for f in [files['queue'], files['results'], files['report']]:
                if os.path.exists(f):
                    os.remove(f)

            # Reset batch state but keep the batch
            state = get_batch_state(self.current_customer, self.current_batch)
            state['processed_domains'] = []
            save_batch_state(self.current_customer, self.current_batch, state)

            self.log(f"Batch '{self.current_batch}' reset")
            self.load_batch_status()

    def select_csv(self):
        if not self.current_customer or not self.current_batch:
            messagebox.showerror("Error", "Please create a customer and batch first")
            return

        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file:
            # Validate CSV before accepting
            is_valid, error_msg, row_count, domain_count = validate_csv(file)

            if not is_valid:
                messagebox.showerror("CSV Validation Error", error_msg)
                return

            self.current_csv = file
            state = get_batch_state(self.current_customer, self.current_batch)
            state['csv_file'] = file
            save_batch_state(self.current_customer, self.current_batch, state)
            self.log(f"CSV loaded: {os.path.basename(file)} ({domain_count} domains)")
            self.load_batch_status()

    def open_batch_report(self):
        if not self.current_customer or not self.current_batch:
            messagebox.showinfo("Info", "Please select a customer and batch first")
            return

        files = get_batch_files(self.current_customer, self.current_batch)
        if os.path.exists(files['report']):
            os.system(f"open '{files['report']}'")
        else:
            messagebox.showinfo("Info", "No report generated yet. Click 'Get Results' first.")

    def generate_combined_report(self):
        if not self.current_customer or self.current_customer == "No customers":
            messagebox.showinfo("Info", "Please select a customer first")
            return

        batches = get_customer_batches(self.current_customer)
        if not batches:
            messagebox.showinfo("Info", "No batches found for this customer")
            return

        self.log(f"Generating combined report for {self.current_customer}...")

        def generate():
            try:
                session = create_api_session()
                batches_data = {}

                for batch_name in batches:
                    files = get_batch_files(self.current_customer, batch_name)
                    if not os.path.exists(files['queue']):
                        continue

                    with open(files['queue'], mode='r', encoding='utf-8-sig') as f:
                        tests = list(csv.reader(f))

                    test_infos = [(row[0], row[1], row[3] if len(row) > 3 else '') for row in tests if len(row) >= 2]

                    if not test_infos:
                        continue

                    all_results = []
                    with ThreadPoolExecutor(max_workers=MAX_PARALLEL_WORKERS) as executor:
                        futures = {executor.submit(fetch_single_result, session, info): info for info in test_infos}
                        for future in as_completed(futures):
                            all_results.append(future.result())

                    if all_results:
                        batches_data[batch_name] = all_results
                        self.root.after(0, lambda b=batch_name: self.log(f"  Fetched: {b}"))

                session.close()

                if batches_data:
                    output_file = get_customer_combined_report(self.current_customer)
                    generate_combined_pdf(self.current_customer, batches_data, output_file)
                    self.root.after(0, lambda: self.log("Combined report generated!"))
                    self.root.after(0, lambda: os.system(f"open '{output_file}'"))
                else:
                    self.root.after(0, lambda: self.log("No data found to generate report"))

            except Exception as e:
                self.root.after(0, lambda: self.log(f"Error: {e}"))

        threading.Thread(target=generate, daemon=True).start()

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
