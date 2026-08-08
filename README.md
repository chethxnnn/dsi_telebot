# Telegram Placement Scraper & Google Sheets Web App 🎓

A modern, automated placement & job posting scraper built with **Telethon**, **FastAPI**, and **Google Sheets**. It extracts structured details (Company Name, Roles, Salary, Location, Eligibility, Registration Links, and Deadlines) from Telegram placement posts and syncs them incrementally to a live Google Sheet or Excel workbook.

---

## 🌟 Features

- **Incremental Syncing**: Remembers the last processed message (`last_id`). Every run only fetches new posts without creating duplicates.
- **Hierarchical Parser**: High-accuracy regex pipeline to handle campus drive posts, probation/post-probation salaries, and multiple job roles per post.
- **One-Click Cloud Dashboard**: Modern web interface deployed on Vercel to sync data on demand with a single click.
- **Google Sheets Integration**: Automatically appends parsed records to a styled Google Sheet accessible anywhere.
- **Excel Export**: Local script support (`telegram_to_excel.py`) with custom column formatting, header styling, and soft-red flags for unparsed fields.

---

## 🛠️ Project Structure

```
.
├── api/
│   └── index.py            # FastAPI Serverless app for Vercel deployment
├── find_group.py           # Helper script to list all Telegram groups & IDs
├── telegram_to_excel.py    # Standalone CLI script to scrape & generate placements.xlsx
├── generate_session.py     # Script to generate portable StringSession for cloud auth
├── google_script.js        # Google Apps Script for Google Sheet Webhook integration
├── vercel.json             # Vercel deployment configuration
├── requirements.txt        # Python dependencies
└── README.md               # Project documentation
```

---

## 🚀 Quick Setup & Local Usage

### 1. Installation
Clone the repository and install Python dependencies:
```bash
git clone https://github.com/chethxnnn/dsi_telebot.git
cd dsi_telebot
pip install -r requirements.txt
```

### 2. Telegram API Credentials
1. Log in to [my.telegram.org](https://my.telegram.org) and create an app.
2. Update `API_ID` and `API_HASH` in `find_group.py` and `telegram_to_excel.py`.

### 3. Find Placement Group ID
Run the helper script to list your Telegram groups:
```bash
python find_group.py
```
Copy your target group ID (e.g., `-1002020152383`) into `GROUP` in `telegram_to_excel.py`.

### 4. Run Local Scraper
Generate your Excel spreadsheet (`placements.xlsx`):
```bash
python telegram_to_excel.py
```

---

## ☁️ Vercel & Google Sheets Deployment

### Step 1: Create Google Sheet Webhook
1. Open a blank spreadsheet on [Google Sheets](https://sheets.google.com).
2. Go to **Extensions** → **Apps Script**.
3. Paste the contents of [`google_script.js`](./google_script.js) and click **Save**.
4. Click **Deploy** → **New deployment** → Choose type **Web app**.
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
5. Copy the generated Web App URL.

### Step 2: Deploy to Vercel
1. Import this repository into [Vercel](https://vercel.com).
2. (Optional) Set Environment Variables in Vercel settings:
   - `TELEGRAM_API_ID`
   - `TELEGRAM_API_HASH`
   - `TELEGRAM_SESSION_STRING` (Generated via `python generate_session.py`)
   - `TELEGRAM_GROUP_ID`
   - `GOOGLE_WEBHOOK_URL`
3. Click **Deploy**.

---

## 📊 Extracted Fields

| Field | Description |
|---|---|
| **Date** | Message posting date (YYYY-MM-DD) |
| **Company** | Inferred or explicit company name |
| **Role** | Job title / roles concatenated with `\|` |
| **Salary (Probation)** | Initial stipend or probation rate |
| **Salary (Post-Probation)** | Confirmed salary or CTC in LPA / monthly |
| **Salary (Raw)** | Verbatim compensation block |
| **Location** | Extracted Indian cities / remote status |
| **Eligibility** | Batch, department, degree requirements |
| **Registration Link** | Google Forms / portal application links |
| **Deadline** | Explicit application cutoff time/date |
| **Needs Review** | Flagged as `Yes` if Company or Role could not be inferred |
| **Message Link** | Direct link to Telegram post |

---

## 📄 License

MIT License. Created for DSATM / College placement automated tracking.
