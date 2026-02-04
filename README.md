# EmailGuard Inbox Placement Tester

A tool to test email inbox placement rates across Google and Microsoft using the [EmailGuard.io](https://emailguard.io) API.

## Features

- ✅ Test inbox placement for multiple email domains
- ✅ Batch processing (50 domains per run)
- ✅ Separate inbox rates for Google and Microsoft
- ✅ Visual PDF reports with color-coded results
- ✅ Auto-polling until all tests complete
- ✅ Automatic updates from GitHub

## Quick Start

### 1. Install Dependencies

```bash
pip3 install requests reportlab customtkinter
```

### 2. Configure API Key

Create a `.env` file in the same folder as the script:

```
EMAILGUARD_API_KEY=your_api_key_here
```

Get your API key from: https://app.emailguard.io/settings/api

### 3. Prepare Your CSV File

Create a CSV file named `smartlead_output.csv` with your email accounts:

| from_name | from_email | user_name | password | smtp_host |
|-----------|------------|-----------|----------|-----------|
| John Doe | john@company.com | john@company.com | app_password | smtp.company.com |
| Jane Smith | jane@company.com | jane@company.com | app_password | smtp.company.com |

### 4. Run the Script

```bash
python3 emailguard.py
```

## Usage

The script launches a modern desktop GUI:

![EmailGuard GUI](https://via.placeholder.com/800x500?text=EmailGuard+GUI)

**Features:**
- 📊 Status cards showing domains processed and tests in queue
- ▶️ Run New Tests - Process up to 50 domains per batch
- 📊 Get Results - Fetch results and generate PDF report
- 🔄 Auto-Poll - Keep checking until all tests complete
- ⚙️ Settings - Configure API key
- 📄 Open Report - View generated PDF

### Typical Workflow

1. **Run New Tests** (Option 1) - Processes up to 50 domains at a time
2. **Auto-Poll Results** (Option 3) - Waits until all tests complete, then generates report
3. View your `inbox_placement_report.pdf`

## Output Files

| File | Description |
|------|-------------|
| `inbox_placement_results.csv` | Detailed results data |
| `inbox_placement_report.pdf` | Visual PDF report |

## PDF Report

The PDF report includes:

- **Executive Summary** - Total tests, completed, pending, failed
- **Overall Performance** - Total inbox rate and spam rate
- **Inbox Rate by Provider** - Separate rates for Google and Microsoft
- **Detailed Results** - Per-domain breakdown with color coding

### Color Coding

- 🟢 Green: Good (≥70%)
- 🟡 Yellow: Warning (50-69%)
- 🔴 Red: Poor (<50%)

## Configuration

You can modify these settings at the top of `emailguard.py`:

```python
CSV_INPUT_FILE = 'smartlead_output.csv'  # Your CSV filename
MAX_DOMAINS_PER_BATCH = 50               # Domains per run
EMAIL_DELAY_SECONDS = 3                  # Delay between emails
POLL_INTERVAL_SECONDS = 30               # Auto-poll check interval
```

## Auto-Updates

The script automatically checks for updates on startup. If a new version is available, you'll see a notification:

```
┌─────────────────────────────────────────────────────┐
│  UPDATE AVAILABLE: v1.1.0 (current: v1.0.0)        │
│  Select option [6] to update                        │
└─────────────────────────────────────────────────────┘
```

Select option [6] to download and install the update.

## Requirements

- Python 3.7+
- `requests` - HTTP library
- `reportlab` - PDF generation
- `customtkinter` - Modern GUI framework

## License

MIT License

## Support

For issues or questions, please open an issue on [GitHub](https://github.com/evesion/emailguard).
