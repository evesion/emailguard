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

GitHub: https://github.com/USERNAME/emailguard
===============================================================================
"""

__version__ = "1.0.0"
__repo__ = "USERNAME/emailguard"  # Update this after creating your repo

import csv
import json
import logging
import os
import shutil
import smtplib
import ssl
import sys
import tempfile
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

CSV_INPUT_FILE = 'smartlead_output.csv'  # Your email accounts CSV file

# Email settings
SMTP_PORT = 465
EMAIL_SUBJECT = "Team Meeting Code"
EMAIL_BODY = """Hello Team, Please find here todays meeting code:"""

# Performance settings
MAX_DOMAINS_PER_BATCH = 50
MAX_PARALLEL_WORKERS = 5
EMAIL_DELAY_SECONDS = 3
POLL_INTERVAL_SECONDS = 30

# ═══════════════════════════════════════════════════════════════════════════════
# INTERNAL SETTINGS - DO NOT MODIFY
# ═══════════════════════════════════════════════════════════════════════════════

API_BASE_URL = "https://app.emailguard.io/api/v1"
DATA_DIR = '.emailguard_data'
CONFIG_FILE = '.env'
BATCH_STATE_FILE = os.path.join(DATA_DIR, 'batch_state.json')
OUTPUT_CSV_FILE = os.path.join(DATA_DIR, 'test_queue.csv')
RESULTS_CSV_FILE = 'inbox_placement_results.csv'
PDF_REPORT_FILE = 'inbox_placement_report.pdf'

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

# Global API key (loaded from config)
API_KEY = None


# ═══════════════════════════════════════════════════════════════════════════════
# AUTO-UPDATER
# ═══════════════════════════════════════════════════════════════════════════════

def check_for_updates(silent=False):
    """Check GitHub for newer version"""
    try:
        response = requests.get(
            f"https://api.github.com/repos/{__repo__}/releases/latest",
            timeout=5
        )
        if response.status_code == 200:
            data = response.json()
            latest_version = data.get('tag_name', '').lstrip('v')

            if latest_version and compare_versions(latest_version, __version__) > 0:
                return {
                    'available': True,
                    'current': __version__,
                    'latest': latest_version,
                    'url': data.get('html_url', ''),
                    'notes': data.get('body', '')[:200]
                }
        return {'available': False, 'current': __version__}
    except Exception as e:
        if not silent:
            logger.debug(f"Update check failed: {e}")
        return {'available': False, 'current': __version__, 'error': str(e)}


def compare_versions(v1, v2):
    """Compare two version strings. Returns 1 if v1 > v2, -1 if v1 < v2, 0 if equal"""
    def parse(v):
        return [int(x) for x in v.split('.')]

    try:
        p1, p2 = parse(v1), parse(v2)
        for i in range(max(len(p1), len(p2))):
            a = p1[i] if i < len(p1) else 0
            b = p2[i] if i < len(p2) else 0
            if a > b:
                return 1
            if a < b:
                return -1
        return 0
    except:
        return 0


def download_update():
    """Download and install the latest version"""
    try:
        print("\n  Checking for updates...")
        update_info = check_for_updates()

        if not update_info.get('available'):
            print("  You already have the latest version!")
            return False

        print(f"  New version available: v{update_info['latest']}")
        print(f"  Current version: v{update_info['current']}")

        confirm = input("\n  Download and install update? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("  Update cancelled.")
            return False

        print("\n  Downloading update...")

        # Get the raw file from GitHub
        raw_url = f"https://raw.githubusercontent.com/{__repo__}/v{update_info['latest']}/emailguard.py"
        response = requests.get(raw_url, timeout=30)

        if response.status_code != 200:
            print(f"  Failed to download update (HTTP {response.status_code})")
            return False

        new_content = response.text

        # Verify it's valid Python
        try:
            compile(new_content, '<string>', 'exec')
        except SyntaxError as e:
            print(f"  Downloaded file has errors: {e}")
            return False

        # Backup current file
        current_file = os.path.abspath(__file__)
        backup_file = current_file + '.backup'

        print("  Creating backup...")
        shutil.copy2(current_file, backup_file)

        # Write new version
        print("  Installing update...")
        with open(current_file, 'w', encoding='utf-8') as f:
            f.write(new_content)

        print(f"\n  ✓ Updated to v{update_info['latest']}!")
        print("  Please restart the script to use the new version.")
        print(f"  Backup saved as: {os.path.basename(backup_file)}")

        return True

    except Exception as e:
        print(f"\n  Update failed: {e}")
        return False


def show_update_notification(update_info):
    """Show update notification banner"""
    if update_info.get('available'):
        print("\n  ┌─────────────────────────────────────────────────────┐")
        print(f"  │  UPDATE AVAILABLE: v{update_info['latest']} (current: v{update_info['current']})".ljust(54) + "│")
        print("  │  Select option [6] to update                        │")
        print("  └─────────────────────────────────────────────────────┘")


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def load_env_file(filepath):
    """Load environment variables from .env file"""
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
    """Load API key from environment or .env file"""
    global API_KEY

    # Try environment variable first
    API_KEY = os.environ.get('EMAILGUARD_API_KEY')

    # Try .env file
    if not API_KEY:
        env_vars = load_env_file(CONFIG_FILE)
        API_KEY = env_vars.get('EMAILGUARD_API_KEY')

    return API_KEY is not None


def setup_api_key():
    """Interactive setup for API key"""
    clear_screen()
    print_header()
    print("\n  API KEY SETUP")
    print("  " + "─" * 40)
    print("\n  No API key found. You need to configure your EmailGuard API key.")
    print("\n  You can get your API key from:")
    print("  https://app.emailguard.io/settings/api")

    api_key = input("\n  Enter your API key: ").strip()

    if not api_key:
        print("\n  No API key entered. Exiting...")
        sys.exit(1)

    # Save to .env file
    with open(CONFIG_FILE, 'w') as f:
        f.write(f"# EmailGuard Configuration\n")
        f.write(f"EMAILGUARD_API_KEY={api_key}\n")

    global API_KEY
    API_KEY = api_key

    print(f"\n  ✓ API key saved to {CONFIG_FILE}")
    print("  You can edit this file later to change your key.")
    input("\n  Press Enter to continue...")


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def clear_screen():
    """Clear the terminal screen"""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """Print the application header"""
    print("\n" + "═" * 60)
    print("           EMAILGUARD INBOX PLACEMENT TESTER")
    print(f"                      v{__version__}")
    print("═" * 60)


def print_menu():
    """Print the main menu"""
    print("\n┌────────────────────────────────────────────────────────┐")
    print("│                      MAIN MENU                         │")
    print("├────────────────────────────────────────────────────────┤")
    print("│  [1]  Run New Tests      - Send test emails            │")
    print("│  [2]  Get Results        - Fetch results & generate PDF│")
    print("│  [3]  Auto-Poll Results  - Wait until all complete     │")
    print("│  [4]  Reset All          - Clear data and start fresh  │")
    print("│  [5]  Help               - Show instructions           │")
    print("│  [6]  Check for Updates  - Update to latest version    │")
    print("│  [0]  Exit                                             │")
    print("└────────────────────────────────────────────────────────┘")


def print_help():
    """Print detailed help instructions"""
    clear_screen()
    print_header()
    print("""
HOW TO USE THIS TOOL
────────────────────────────────────────────────────────────

STEP 1: CONFIGURE API KEY
   Create a .env file with your EmailGuard API key:

   EMAILGUARD_API_KEY=your_api_key_here

   Get your API key from: https://app.emailguard.io/settings/api

STEP 2: PREPARE YOUR CSV FILE
   Create a CSV file with your email accounts. Required columns:

   from_name    | from_email           | user_name            | password     | smtp_host
   -------------|----------------------|----------------------|--------------|------------------
   John Doe     | john@company.com     | john@company.com     | app_password | smtp.company.com

STEP 3: RUN NEW TESTS (Option 1)
   - Processes up to 50 unique domains per run
   - Creates a test for each domain via EmailGuard API
   - Sends test emails from your accounts

STEP 4: GET RESULTS (Option 2 or 3)
   - Option 2: Fetch current results once
   - Option 3: Keep polling until all tests complete (recommended)
   - Generates a PDF report with inbox placement rates

OUTPUT FILES
────────────────────────────────────────────────────────────
   inbox_placement_results.csv  - Detailed results data
   inbox_placement_report.pdf   - Visual PDF report

COLOR CODING
────────────────────────────────────────────────────────────
   Green  = Good (≥70%)
   Yellow = Warning (50-69%)
   Red    = Poor (<50%)

Press Enter to return to menu...""")
    input()


def ensure_data_dir():
    """Ensure the data directory exists"""
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)


def create_api_session():
    """Create a requests session with retry logic"""
    session = requests.Session()
    retry_strategy = Retry(
        total=3,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
    )
    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=MAX_PARALLEL_WORKERS, pool_maxsize=MAX_PARALLEL_WORKERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"})
    return session


def load_batch_state():
    """Load the batch state"""
    ensure_data_dir()
    if os.path.exists(BATCH_STATE_FILE):
        try:
            with open(BATCH_STATE_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {'processed_domains': [], 'batch_number': 0}


def save_batch_state(state):
    """Save the batch state"""
    ensure_data_dir()
    with open(BATCH_STATE_FILE, 'w') as f:
        json.dump(state, f, indent=2)


def get_unique_domains(csv_file):
    """Get unique domains from CSV"""
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
        print(f"\n  ERROR: CSV file '{csv_file}' not found!")
        print(f"  Please place your email accounts CSV in this folder.")
        return None, None
    except KeyError as e:
        print(f"\n  ERROR: CSV is missing required column: {e}")
        print(f"  Required columns: from_name, from_email, user_name, password, smtp_host")
        return None, None

    return domains, domain_to_row


# ═══════════════════════════════════════════════════════════════════════════════
# EMAIL & API FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def send_email(from_name, from_email, user_name, password, smtp_host, recipients, filter_phrase):
    """Send test email"""
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
    """Create inbox placement test"""
    try:
        response = session.post(f"{API_BASE_URL}/inbox-placement-tests", json={"name": test_name})
        response.raise_for_status()
        return response.json(), None
    except Exception as e:
        return None, str(e)


def get_test_results(session, test_uuid):
    """Get test results"""
    try:
        response = session.get(f"{API_BASE_URL}/inbox-placement-tests/{test_uuid}")
        response.raise_for_status()
        return response.json(), None
    except Exception as e:
        return None, str(e)


def calculate_stats(test_emails):
    """Calculate placement statistics"""
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


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN MENU FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def run_new_tests():
    """Run new inbox placement tests"""
    clear_screen()
    print_header()
    print("\n  RUNNING NEW TESTS")
    print("  " + "─" * 40)

    state = load_batch_state()
    processed = set(state['processed_domains'])
    batch_num = state['batch_number'] + 1

    all_domains, domain_map = get_unique_domains(CSV_INPUT_FILE)
    if all_domains is None:
        input("\n  Press Enter to return to menu...")
        return

    remaining = [d for d in all_domains if d not in processed]

    if not remaining:
        print(f"\n  All {len(all_domains)} domains have been processed!")
        print("  Use 'Reset All' to start fresh.")
        input("\n  Press Enter to return to menu...")
        return

    batch = remaining[:MAX_DOMAINS_PER_BATCH]

    print(f"\n  Batch #{batch_num}")
    print(f"  ├─ Total domains in CSV:    {len(all_domains)}")
    print(f"  ├─ Already processed:       {len(processed)}")
    print(f"  ├─ Processing this batch:   {len(batch)}")
    print(f"  └─ Remaining after batch:   {len(remaining) - len(batch)}")

    input("\n  Press Enter to start...")

    session = create_api_session()
    ensure_data_dir()
    results = []
    failed = []

    print()
    for i, domain in enumerate(batch, 1):
        row = domain_map[domain]
        print(f"  [{i}/{len(batch)}] {domain}...", end=" ", flush=True)

        test_name = f"Inbox Test - {domain} - {time.strftime('%Y-%m-%d %H:%M:%S')}"
        test_result, error = create_test(session, test_name)

        if not test_result or 'data' not in test_result:
            print("FAILED (API)")
            failed.append((domain, error))
            continue

        test_data = test_result['data']
        test_uuid = test_data['uuid']
        filter_phrase = test_data['filter_phrase']
        test_emails = test_data['comma_separated_test_email_addresses']

        success, error = send_email(
            row.get('from_name', ''), row['from_email'], row['user_name'],
            row['password'], row['smtp_host'], test_emails.replace(',', ';'), filter_phrase
        )

        if success:
            print("OK")
            results.append({
                'domain': domain,
                'from_email': row['from_email'],
                'test_uuid': test_uuid,
                'filter_phrase': filter_phrase,
                'test_url': f"https://app.emailguard.io/inbox-placement-tests/{test_uuid}"
            })
        else:
            err_msg = str(error)[:30] + "..." if len(str(error)) > 30 else str(error)
            print(f"FAILED ({err_msg})")
            failed.append((domain, error))

        if i < len(batch):
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

    print("\n  " + "─" * 40)
    print(f"  BATCH #{batch_num} COMPLETE")
    print(f"  ├─ Successful: {len(results)}")
    print(f"  ├─ Failed:     {len(failed)}")
    print(f"  └─ Total processed: {len(state['processed_domains'])}/{len(all_domains)}")

    if failed:
        print("\n  Failed domains:")
        for domain, error in failed[:5]:
            err_msg = str(error)[:40] + "..." if len(str(error)) > 40 else str(error)
            print(f"    - {domain}: {err_msg}")
        if len(failed) > 5:
            print(f"    ... and {len(failed) - 5} more")

    remaining_count = len(all_domains) - len(state['processed_domains'])
    if remaining_count > 0:
        print(f"\n  Run again to process {min(remaining_count, MAX_DOMAINS_PER_BATCH)} more domains.")

    input("\n  Press Enter to return to menu...")


def fetch_single_result(session, test_info):
    """Fetch a single test result"""
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


def get_results(auto_poll=False):
    """Fetch and display results"""
    clear_screen()
    print_header()
    print(f"\n  {'AUTO-POLLING' if auto_poll else 'FETCHING'} RESULTS")
    print("  " + "─" * 40)

    if not os.path.exists(OUTPUT_CSV_FILE):
        print("\n  No tests found. Run 'New Tests' first.")
        input("\n  Press Enter to return to menu...")
        return

    with open(OUTPUT_CSV_FILE, mode='r', encoding='utf-8-sig') as f:
        tests = list(csv.reader(f))

    if not tests:
        print("\n  No tests found. Run 'New Tests' first.")
        input("\n  Press Enter to return to menu...")
        return

    test_infos = [(row[0], row[1], row[3] if len(row) > 3 else '') for row in tests if len(row) >= 2]

    session = create_api_session()

    while True:
        print(f"\n  Fetching {len(test_infos)} test results...")

        all_results = []
        with ThreadPoolExecutor(max_workers=MAX_PARALLEL_WORKERS) as executor:
            futures = {executor.submit(fetch_single_result, session, info): info for info in test_infos}
            for future in as_completed(futures):
                all_results.append(future.result())

        completed = [r for r in all_results if r.get('status') in ['completed', 'complete']]
        pending = [r for r in all_results if r.get('status') not in ['completed', 'complete', 'FAILED']]
        failed = [r for r in all_results if r.get('status') == 'FAILED']

        print(f"\n  Status: {len(completed)} completed, {len(pending)} pending, {len(failed)} failed")

        if not auto_poll or not pending:
            break

        print(f"\n  Waiting {POLL_INTERVAL_SECONDS} seconds before next check...")
        print("  (Press Ctrl+C to stop and generate report with current results)")

        try:
            time.sleep(POLL_INTERVAL_SECONDS)
            clear_screen()
            print_header()
            print("\n  AUTO-POLLING RESULTS")
            print("  " + "─" * 40)
        except KeyboardInterrupt:
            print("\n\n  Stopping auto-poll...")
            break

    session.close()

    generate_outputs(all_results)

    completed = [r for r in all_results if r.get('status') in ['completed', 'complete']]

    if completed:
        avg_inbox = sum(r['stats']['inbox_rate'] for r in completed) / len(completed)
        avg_google = sum(r['stats']['google_inbox_rate'] for r in completed if r['stats']['google_total'] > 0)
        avg_microsoft = sum(r['stats']['microsoft_inbox_rate'] for r in completed if r['stats']['microsoft_total'] > 0)

        google_count = len([r for r in completed if r['stats']['google_total'] > 0])
        microsoft_count = len([r for r in completed if r['stats']['microsoft_total'] > 0])

        if google_count > 0:
            avg_google /= google_count
        if microsoft_count > 0:
            avg_microsoft /= microsoft_count

        print("\n  " + "─" * 40)
        print("  SUMMARY")
        print(f"  ├─ Total Inbox Rate:     {avg_inbox:.1f}%")
        print(f"  ├─ Google Inbox Rate:    {avg_google:.1f}%")
        print(f"  └─ Microsoft Inbox Rate: {avg_microsoft:.1f}%")

    print("\n  Files generated:")
    print(f"  ├─ {RESULTS_CSV_FILE}")
    print(f"  └─ {PDF_REPORT_FILE}")

    input("\n  Press Enter to return to menu...")


def generate_outputs(all_results):
    """Generate CSV and PDF outputs"""
    with open(RESULTS_CSV_FILE, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['from_email', 'test_uuid', 'status', 'overall_score', 'inbox_rate_%',
                        'spam_rate_%', 'google_inbox_rate_%', 'microsoft_inbox_rate_%', 'test_url'])

        for r in all_results:
            stats = r.get('stats', {})
            writer.writerow([
                r.get('from_email', ''), r.get('test_uuid', ''), r.get('status', ''),
                r.get('overall_score', ''), f"{stats.get('inbox_rate', 0):.1f}",
                f"{stats.get('spam_rate', 0):.1f}", f"{stats.get('google_inbox_rate', 0):.1f}",
                f"{stats.get('microsoft_inbox_rate', 0):.1f}", r.get('test_url', '')
            ])

    try:
        generate_pdf(all_results)
    except ImportError:
        print("\n  Note: Install 'reportlab' for PDF generation: pip install reportlab")
    except Exception as e:
        print(f"\n  Warning: Could not generate PDF: {e}")


def generate_pdf(all_results):
    """Generate PDF report"""
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
        if len(domain) > 18:
            domain = domain[:15] + '...'
        stats = r.get('stats', {})
        status = r.get('status', 'Unknown')

        table_data.append([
            domain,
            f"{stats.get('inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('google_inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('microsoft_inbox_rate', 0):.0f}%" if stats else '-',
            f"{stats.get('spam_rate', 0):.0f}%" if stats else '-',
            status
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


def reset_all():
    """Reset all data"""
    clear_screen()
    print_header()
    print("\n  RESET ALL DATA")
    print("  " + "─" * 40)
    print("\n  This will delete:")
    print(f"    - Batch state (processed domains)")
    print(f"    - Test queue")
    print("\n  Output files (CSV, PDF) will NOT be deleted.")
    print("  Your .env config file will NOT be deleted.")

    confirm = input("\n  Type 'yes' to confirm: ").strip().lower()

    if confirm == 'yes':
        files = [BATCH_STATE_FILE, OUTPUT_CSV_FILE]
        for f in files:
            if os.path.exists(f):
                os.remove(f)

        if os.path.exists(DATA_DIR):
            try:
                os.rmdir(DATA_DIR)
            except:
                pass

        print("\n  ✓ All data has been reset!")
    else:
        print("\n  Reset cancelled.")

    input("\n  Press Enter to return to menu...")


def check_updates_menu():
    """Check for updates menu option"""
    clear_screen()
    print_header()
    print("\n  CHECK FOR UPDATES")
    print("  " + "─" * 40)

    download_update()

    input("\n  Press Enter to return to menu...")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    """Main entry point"""
    # Load API key
    if not load_api_key():
        setup_api_key()

    # Check for updates on startup (silent)
    update_info = check_for_updates(silent=True)

    while True:
        clear_screen()
        print_header()

        # Show update notification if available
        show_update_notification(update_info)

        # Show current status
        state = load_batch_state()
        domains, _ = get_unique_domains(CSV_INPUT_FILE)

        if domains:
            processed = len(state.get('processed_domains', []))
            print(f"\n  Status: {processed}/{len(domains)} domains processed")

            if os.path.exists(OUTPUT_CSV_FILE):
                with open(OUTPUT_CSV_FILE, 'r') as f:
                    tests = len(list(csv.reader(f)))
                print(f"  Tests in queue: {tests}")

        print_menu()

        choice = input("\n  Enter choice [0-6]: ").strip()

        if choice == '1':
            run_new_tests()
        elif choice == '2':
            get_results(auto_poll=False)
        elif choice == '3':
            get_results(auto_poll=True)
        elif choice == '4':
            reset_all()
        elif choice == '5':
            print_help()
        elif choice == '6':
            check_updates_menu()
            update_info = check_for_updates(silent=True)  # Refresh after update attempt
        elif choice == '0':
            clear_screen()
            print("\n  Goodbye!\n")
            sys.exit(0)
        else:
            print("\n  Invalid choice. Press Enter to try again...")
            input()


if __name__ == "__main__":
    main()
