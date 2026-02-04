# EmailGuard Inbox Placement Tester

A modern desktop application to test email inbox placement rates across Google and Microsoft using the [EmailGuard.io](https://emailguard.io) API.

## Features

- **Multi-Customer Support** - Manage multiple customers with separate data
- **Batch Testing** - Create named batches for each test run (e.g., "January Warmup", "Post-Migration Check")
- **Bulk Processing** - Process up to 50 domains per batch automatically
- **Real-time Polling** - Auto-poll for results until all tests complete
- **PDF Reports**:
  - Individual batch reports
  - Combined customer reports (all batches in one PDF)
  - Google & Microsoft inbox rates breakdown
- **Progress Tracking** - Remembers progress between sessions
- **Auto-Updates** - Notifies when new versions are available

## Quick Start

### 1. Install Dependencies

```bash
pip3 install requests reportlab customtkinter
```

### 2. Clone & Run

```bash
git clone https://github.com/evesion/emailguard.git
cd emailguard
python3 emailguard.py
```

### 3. First Run Setup

On first launch, you'll be prompted to:
1. Enter your EmailGuard API key (get it from https://app.emailguard.io/settings/api)
2. Create a customer
3. Create a batch
4. Select a CSV file with your email accounts

## CSV Format

Your CSV file must have these columns:

| from_name | from_email | user_name | password | smtp_host |
|-----------|------------|-----------|----------|-----------|
| John Doe | john@company.com | john@company.com | app_password | smtp.company.com |
| Jane Smith | jane@company.com | jane@company.com | app_password | smtp.company.com |

## Usage

### Workflow

1. **Create Customer** - Click "+ New" next to customer dropdown
2. **Create Batch** - Click "+ New" next to batch dropdown and give it a name
3. **Select CSV** - Click "Select CSV" to choose your email accounts file
4. **Run Tests** - Click "Run New Tests" to start inbox placement tests
5. **Get Results** - Click "Get Results" or "Auto-Poll" to fetch test results
6. **View Reports**:
   - "Batch Report" - View report for current batch
   - "Combined Report" - Generate report for all batches (current customer)

### GUI Features

- 📊 Status cards showing domains processed and tests in queue
- ▶️ Run New Tests - Process up to 50 domains per batch
- 📊 Get Results - Fetch results and generate PDF report
- 🔄 Auto-Poll - Keep checking until all tests complete
- 📄 Batch Report - Open the current batch's PDF report
- 📑 Combined Report - Generate a combined PDF for all batches
- ⚙️ Settings - Configure API key
- 🗑️ Reset Batch - Clear current batch data and start over

## Data Structure

```
.emailguard_data/
├── customers.json
├── Customer A/
│   ├── customer_state.json
│   ├── Batch - January Test/
│   │   ├── batch_state.json
│   │   ├── test_queue.csv
│   │   ├── results.csv
│   │   └── report.pdf
│   ├── Batch - February Test/
│   │   └── ...
│   └── combined_report.pdf
├── Customer B/
│   └── ...
```

## PDF Reports

### Batch Report
- Executive Summary - Total tests, completed, pending, failed
- Overall Performance - Total inbox rate and spam rate
- Inbox Rate by Provider - Separate rates for Google and Microsoft
- Detailed Results - Per-domain breakdown

### Combined Report
- Summary across all batches
- Batch comparison table
- Combined performance metrics
- Individual batch details with full results

### Color Coding

- 🟢 Green: Good (≥70%)
- 🟡 Yellow: Warning (50-69%)
- 🔴 Red: Poor (<50%)

## Configuration

You can modify these settings at the top of `emailguard.py`:

```python
MAX_DOMAINS_PER_BATCH = 50               # Domains per run
EMAIL_DELAY_SECONDS = 3                  # Delay between emails
POLL_INTERVAL_SECONDS = 30               # Auto-poll check interval
```

## Auto-Updates

The app automatically checks for updates on startup. If a new version is available, you'll see an "Update Available" button in the header.

## Requirements

- Python 3.8+
- `requests` >= 2.28.0
- `reportlab` >= 4.0.0
- `customtkinter` >= 5.2.0

## License

MIT License

## Support

For issues or questions, please open an issue on [GitHub](https://github.com/evesion/emailguard).
