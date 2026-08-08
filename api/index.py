# -*- coding: utf-8 -*-
"""
Vercel Serverless Web Application for Telegram Placement Scraper
Serves a modern single-page dashboard and handles incremental sync to Google Sheets.
"""

import os
import re
import json
import asyncio
from typing import Optional
import urllib.request
import urllib.error

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, JSONResponse
from telethon import TelegramClient
from telethon.sessions import StringSession

app = FastAPI(title="Placement Scraper Dashboard")

def parse_int_env(key: str, default: int = 0) -> int:
    val = os.environ.get(key, "").strip()
    return int(val) if val.lstrip('-').isdigit() else default

# Safe Environment Variable Parsing
DEFAULT_API_ID = parse_int_env("TELEGRAM_API_ID", 39806525)
DEFAULT_API_HASH = os.environ.get("TELEGRAM_API_HASH", "20561160a2f41ad9cbb3f9e45e9bdf67")
DEFAULT_SESSION_STRING = os.environ.get("TELEGRAM_SESSION_STRING", "1BVtsOHkBu7kXAtq0qkp9Iw8iknD56onfN5y5MdVv42D_sgoFkb-9bwt0c5DFVSBjw5T3BbDW5apscHjXR7sI2soMAnEo4BmzVUpR6KrIsF_PWjeg4zOlqB5BK2Z-w02D2-jCY00QhbO4ybhD8oQ4L4dQRhd1scPKB4Qy4oteLYdO3hyyE-IPd3wjGtK47KPiRL3pjQL3ckkqj-KQVePNNszaOW9FOnqb4E9n5uU_C95oS5ZaUPmkMkjNwHXLA9ILA-qGAuJnltuqjHUMM3eNF1Ei6Wx3eAOFY9xLirQbzIPThP1v-QSR5iKwi_8LfAZ31ecOswoklygwPIxwNHSUT34BTWX5Eio=")
DEFAULT_GROUP_ID = os.environ.get("TELEGRAM_GROUP_ID", "-1002020152383")
DEFAULT_WEBHOOK_URL = os.environ.get("GOOGLE_WEBHOOK_URL", "")

# --- Extractor Utilities ---
def is_placement_post(text: str) -> bool:
    if not text or len(text) < 80:
        return False
    keywords = [
        'company', 'role', 'salary', 'compensation', 'ctc', 'lpa', 
        'campus drive', 'job title', 'eligibility', 'register', 
        'recruitment', 'openings', 'hiring', 'walk-in', 'package', 
        'stipend', 'job description', 'vacancy', 'fresher'
    ]
    text_lower = text.lower()
    hits = sum(1 for kw in keywords if kw in text_lower)
    return hits >= 2

def extract_company(text: str) -> Optional[str]:
    invalid_companies = {
        'about company', 'details', 'key details', 'hrms', 'workshop highlights',
        'alert', 'opportunity with', 'remainder', 'registration', 'who can',
        'ask how', 'updates', 'policy', 'note', 'important note', 'final reminder',
        'gentle reminder', 'call for', 'incorrect', 'who can participate',
        'registration deadline', 'test configuration', 'public keys', 'the company',
        'for women coders', 'challenge alert'
    }

    def clean_comp_name(name):
        if not name:
            return None
        name = re.sub(r'^(?:Greetings?\s+from|Hurry\s+Up\s*[\-\:]?|Alert\s*[\-\:]?|Attention\s+to\s+All\s+(?:Registered\s+)?Students\s+(?:for\s+)?|Important\s+Note\s+(?:for\s+[\w\s]+\s+on\s+)?)\s*', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'\s*-\s*Pune\'s\s+Top.*$', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'\s*[\-\:]\s*All\s+Branches.*$', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'[^\w\s\-\&\.\,\']', '', name).strip('- ').strip()
        if name.lower() in invalid_companies or any(inv in name.lower() for inv in ['registration', 'updates', 'reminder', 'highlight', 'webinar']):
            return None
        if len(name.split()) > 7 or len(name) > 60:
            return None
        return name

    m1 = re.search(r'Company\s*:\s*(.+)', text, re.IGNORECASE)
    if m1:
        val = clean_comp_name(m1.group(1).strip().split('\n')[0].strip())
        if val: return val

    m2 = re.search(r'Greetings from\s+(.+?)[\.!\n]', text, re.IGNORECASE)
    if m2:
        val = clean_comp_name(m2.group(1).strip())
        if val: return val

    m3 = re.search(r'^(.+?)\s+is\s+(?:currently\s+)?(?:hiring|recruiting)', text, re.IGNORECASE | re.MULTILINE)
    if m3:
        val = clean_comp_name(m3.group(1).strip())
        if val: return val

    m4 = re.search(r'^(.+?)\s+wants to hire', text, re.IGNORECASE | re.MULTILINE)
    if m4:
        val = clean_comp_name(m4.group(1).strip())
        if val: return val

    first_line = text.split('\n')[0].strip()
    m5 = re.match(r'^[^\w]*([\w][\w\s\-\&\.]+?)\s*(?:Campus\s+(?:drive|recruitment)|Recruitment|Off[- ]?Campus)', first_line, re.IGNORECASE)
    if m5:
        val = clean_comp_name(m5.group(1))
        if val: return val

    clean_line = re.sub(r'[^\w\s\-,\&]', '', first_line).strip()
    clean_line = re.sub(r'\s*-\s*\d{4}', '', clean_line).strip()
    suffixes = [r'Campus\s+drive', r'Hiring', r'Recruitment', r'Placements?', r'Drive']
    for suf in suffixes:
        clean_line = re.sub(rf'\b{suf}\b', '', clean_line, flags=re.IGNORECASE).strip()
    clean_line = clean_line.strip('- ').strip()

    greetings = ['dear', 'hello', 'hi ', 'attention', 'note', 'important', 'reminder',
                 'final reminder', 'students', 'all ', 'this ', 'today', 'we ', 'the ',
                 'registration', 'who can', 'ask how', 'only for', 'gentle reminder',
                 'call for', 'incorrect', 'job description']
    if clean_line and len(clean_line.split()) <= 5:
        if not any(clean_line.lower().startswith(g) for g in greetings):
            val = clean_comp_name(clean_line)
            if val: return val

    return None

def extract_roles(text: str):
    roles = []
    def is_valid_role(r):
        if not r: return False
        r_str = r.strip()
        if re.search(r'https?://|www\.|drive\.google\.com|forms\.gle|\.php|\.pdf', r_str, re.IGNORECASE): return False
        if len(r_str.split()) > 8 or len(r_str) > 70: return False
        if any(w in r_str.lower() for w in ['strictly', 'unplaced', 'eligible', 'selection', 'written test', 'interview', 'round', 'responsibility', 'responsibilities', 'criteria', 'note:']): return False
        return True

    m1 = re.search(r'(?:Job Title|Position|Profile|Designation|Role Title)\s*:\s*(.+)', text, re.IGNORECASE)
    if m1:
        val = m1.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val): return [val]

    m2 = re.search(r'Job Description\s*[–\-:]\s*(.+)', text, re.IGNORECASE)
    if m2:
        val = m2.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val): return [val]

    for m in re.finditer(r'(?:^|\n)\s*Role\s*:\s*(.+)', text, re.IGNORECASE):
        val = m.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val) and 'overview' not in val.lower():
            roles.append(val)
    if roles: return roles[:5]

    for line in text.split('\n'):
        m = re.match(r'^\s*\d+[\.\)]\s*([A-Za-z0-9\s\-\/\(\)\&]+?)(?:\s*:|\s*–|\s*-|$)', line)
        if m:
            val = m.group(1).strip()
            if is_valid_role(val) and any(kw in val.lower() for kw in ['engineer', 'developer', 'associate', 'executive', 'analyst', 'trainee', 'intern', 'manager', 'lead', 'consultant', 'bda', 'boe', 'sde', 'get']):
                roles.append(val)
    if roles: return roles[:5]

    role_section = False
    for line in text.split('\n'):
        line_lower = line.lower().strip()
        if re.search(r'(?:hiring|openings?|positions?|vacancies|following roles|open roles)\s*:?\s*$', line_lower):
            role_section = True
            continue
        if role_section:
            if re.search(r'(?:eligibility|criteria|responsibilit|qualification|selection|process|compensation|salary|about|key resp|requirements|skills)', line_lower):
                role_section = False
                continue
            m = re.match(r'^\s*[•\-\*]\s+(.+)', line, re.UNICODE)
            if m:
                candidate = m.group(1).strip(': ')
                if is_valid_role(candidate): roles.append(candidate)
            elif line.strip() and not re.match(r'^\s*[•\-\*]', line, re.UNICODE):
                role_section = False
    if roles: return roles[:5]

    known_roles = [
        'Software Development Engineer', 'Software Engineer', 'Full Stack Developer',
        'Frontend Developer', 'Backend Developer', 'DevOps Engineer', 'Data Analyst',
        'Data Scientist', 'Business Analyst', 'Business Development Associate',
        'Business Development Executive', 'Marketing Associate', 'Marketing Executive',
        'System Engineer', 'Graduate Engineer Trainee', 'Programmer Analyst Trainee',
        'Programmer Analyst', 'Associate Software Engineer', 'QA Engineer',
        'Quality Assurance Engineer', 'Testing Engineer', 'Cyber Security Intern',
        'Cloud Engineer', 'Sales Executive', 'Customer Success Associate'
    ]
    for kr in known_roles:
        if re.search(rf'\b{re.escape(kr)}\b', text, re.IGNORECASE):
            roles.append(kr)
            break
    return roles

def extract_salary(text: str) -> dict:
    result = {'probation': None, 'post_probation': None, 'raw': None}
    raw_match = re.search(
        r'(?:Compensation(?: Structure)?|Salary|CTC)\s*:\s*(.+?)(?=\n\s*(?:[A-Z][a-z]+|Eligibility|Location|Selection|Key Resp)\s*:|$)',
        text, re.IGNORECASE | re.DOTALL
    )
    if raw_match:
        result['raw'] = raw_match.group(1).strip()[:300]
    else:
        m_sal = re.search(r'(?:Salary|CTC)\s*:\s*(.+)', text, re.IGNORECASE)
        if m_sal: result['raw'] = m_sal.group(1).strip()

    m_prob = re.search(r'During Probation[^:]*:\s*₹?\s*([\d,]+)\s*(?:per month|p\.?m\.?|/month)?', text, re.IGNORECASE | re.UNICODE)
    m_post = re.search(r'Post Probation[^:]*:\s*₹?\s*([\d,]+)\s*(?:per month|p\.?m\.?|/month)?', text, re.IGNORECASE | re.UNICODE)
    if m_prob and m_post:
        result['probation'] = f'₹{m_prob.group(1).strip()}/month'
        result['post_probation'] = f'₹{m_post.group(1).strip()}/month'
        m_post_lpa = re.search(r'Post Probation[^\n]*(\d+(?:\.\d+)?)\s*LPA', text, re.IGNORECASE)
        if m_post_lpa: result['post_probation'] = f'{m_post_lpa.group(1)} LPA'
        return result

    flat_match = re.search(r'(?:Salary|CTC)\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if flat_match:
        val = flat_match.group(1).strip()
        result['probation'] = val
        result['post_probation'] = val
        return result

    if not result['probation'] and not result['post_probation']:
        lpas = re.findall(r'(\d+(?:\.\d+)?)\s*LPA', text, re.IGNORECASE)
        monthly_rs = re.findall(r'₹\s*([\d,]+)\s*(?:per month|p\.?m\.?|/month)', text, re.IGNORECASE | re.UNICODE)
        if lpas:
            lpa_vals = sorted(set(float(l) for l in lpas))
            if len(lpa_vals) >= 2:
                result['probation'] = f'{lpa_vals[0]} LPA'
                result['post_probation'] = f'{lpa_vals[-1]} LPA'
            else:
                result['probation'] = f'{lpa_vals[0]} LPA'
                result['post_probation'] = f'{lpa_vals[0]} LPA'
        elif monthly_rs:
            rs_vals = sorted(set(int(r.replace(',', '')) for r in monthly_rs))
            if len(rs_vals) >= 2:
                result['probation'] = f'₹{rs_vals[0]:,}/month'
                result['post_probation'] = f'₹{rs_vals[-1]:,}/month'
            else:
                result['probation'] = f'₹{rs_vals[0]:,}/month'
                result['post_probation'] = f'₹{rs_vals[0]:,}/month'
    return result

def extract_locations(text: str) -> Optional[str]:
    m = re.search(r'(?:Job\s+)?Locations?\s*:\s*(.+)', text, re.IGNORECASE)
    locations = []
    if m:
        raw_loc = m.group(1)
        locations = [l.strip() for l in re.split(r'[,|]', raw_loc) if l.strip()]
    else:
        cities = ['Bangalore', 'Bengaluru', 'Hyderabad', 'Chennai', 'Mumbai', 'Delhi', 'NCR', 'Pune', 'Noida', 'Gurgaon', 'Gurugram', 'Kolkata', 'Ahmedabad', 'Jaipur', 'Kochi', 'Coimbatore', 'Chandigarh', 'Lucknow', 'Indore', 'Remote', 'Work from Home', 'WFH', 'Pan India']
        for city in cities:
            if re.search(rf'\b{city}\b', text, re.IGNORECASE):
                locations.append(city)
    seen = set()
    dedup = []
    for loc in locations:
        if loc.lower() not in seen:
            seen.add(loc.lower())
            dedup.append(loc)
    return ", ".join(dedup) if dedup else None

def extract_eligibility(text: str) -> Optional[str]:
    m = re.search(r'Eligibility\s*(?:Criteria)?\s*:?\s*\n?(.*?)(?=\n\s*(?:[A-Z][a-zA-Z\s]+:|Required Skills|Compensation|Selection Process|About|Key Resp)|\Z)', text, re.IGNORECASE | re.DOTALL)
    if m:
        content = m.group(1).strip()
        if not content: return None
        content = re.sub(r'^\s*[•\-\*]\s*', '', content, flags=re.MULTILINE | re.UNICODE)
        content = re.sub(r'\n+', ' | ', content).strip(' |')
        return content if len(content) > 5 else None
    return None

def extract_links(text: str) -> Optional[str]:
    urls = re.findall(r'https?://[^\s<>"\'\)]+', text, re.IGNORECASE)
    return " | ".join(urls) if urls else None

def extract_deadline(text: str) -> Optional[str]:
    patterns = [
        r'(?:on or before|last date[:\s]*)\s*(.+?)(?:\.|\n|$)',
        r'(?:register|apply)\s+(?:by|before)\s+(.+?)(?:\.|\n|using|$)',
        r'(?:before|by)\s+(\d{1,2}[:\d]*\s*(?:AM|PM|am|pm)\s+(?:today|tomorrow|\d{1,2}(?:st|nd|rd|th)?\s+\w+(?:\s+\d{4})?))',
        r'(?:deadline|last date)\s*[:\-]\s*(.+?)(?:\.|\n|$)',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if len(val) > 5: return val
    return None

def parse_message(msg):
    text = msg.text or ""
    company = extract_company(text)
    roles = extract_roles(text)
    salary = extract_salary(text)
    location = extract_locations(text)
    eligibility = extract_eligibility(text)
    links = extract_links(text)
    deadline = extract_deadline(text)
    role_str = ' | '.join(roles) if roles else None
    chat_id = str(msg.chat_id).replace('-100', '') if msg.chat_id else ''
    return {
        'date': msg.date.strftime('%Y-%m-%d') if msg.date else '',
        'company': company,
        'role': role_str,
        'salary_probation': salary['probation'],
        'salary_post_probation': salary['post_probation'],
        'salary_raw': salary['raw'],
        'location': location,
        'eligibility': eligibility,
        'registration_link': links,
        'deadline': deadline,
        'needs_review': 'Yes' if not company or not role_str else 'No',
        'message_link': f'https://t.me/c/{chat_id}/{msg.id}' if chat_id else '',
        'raw_message': text[:500]
    }

# --- Web Application Routes ---
@app.get("/", response_class=HTMLResponse)
async def dashboard():
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Placement Telegram Scraper Dashboard</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {
                --bg: #0F172A;
                --card-bg: rgba(30, 41, 59, 0.7);
                --card-border: rgba(255, 255, 255, 0.1);
                --primary: #3B82F6;
                --primary-hover: #2563EB;
                --text-main: #F8FAFC;
                --text-muted: #94A3B8;
                --success: #10B981;
                --warning: #F59E0B;
            }
            body {
                font-family: 'Inter', sans-serif;
                background-color: var(--bg);
                color: var(--text-main);
                margin: 0;
                padding: 40px 20px;
                display: flex;
                justify-content: center;
            }
            .container {
                max-width: 800px;
                width: 100%;
            }
            .header {
                text-align: center;
                margin-bottom: 30px;
            }
            .header h1 {
                font-size: 28px;
                font-weight: 700;
                margin: 0 0 8px 0;
                background: linear-gradient(135deg, #60A5FA, #A78BFA);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }
            .header p {
                color: var(--text-muted);
                margin: 0;
                font-size: 14px;
            }
            .card {
                background: var(--card-bg);
                border: 1px solid var(--card-border);
                backdrop-filter: blur(12px);
                border-radius: 16px;
                padding: 24px;
                margin-bottom: 24px;
                box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
            }
            .form-group {
                margin-bottom: 16px;
            }
            label {
                display: block;
                font-size: 13px;
                font-weight: 500;
                color: var(--text-muted);
                margin-bottom: 6px;
            }
            input {
                width: 100%;
                padding: 12px;
                border-radius: 8px;
                border: 1px solid var(--card-border);
                background: rgba(15, 23, 42, 0.6);
                color: #FFF;
                font-size: 14px;
                box-sizing: border-box;
            }
            input:focus {
                outline: none;
                border-color: var(--primary);
            }
            .btn {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 100%;
                padding: 14px;
                background: var(--primary);
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 15px;
                font-weight: 600;
                cursor: pointer;
                transition: background 0.2s, transform 0.1s;
            }
            .btn:hover {
                background: var(--primary-hover);
            }
            .btn:active {
                transform: scale(0.98);
            }
            .btn-secondary {
                background: rgba(255, 255, 255, 0.1);
                color: var(--text-main);
                margin-top: 10px;
            }
            .btn-secondary:hover {
                background: rgba(255, 255, 255, 0.18);
            }
            #status-box {
                margin-top: 20px;
                padding: 16px;
                border-radius: 10px;
                background: rgba(15, 23, 42, 0.8);
                font-family: monospace;
                font-size: 13px;
                color: #A7F3D0;
                display: none;
                white-space: pre-wrap;
            }
            .badge {
                display: inline-block;
                padding: 4px 10px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 600;
                background: rgba(16, 185, 129, 0.15);
                color: var(--success);
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>🎓 Telegram Placement Scraper</h1>
                <p>One-click Cloud Sync to Google Sheets</p>
            </div>

            <div class="card">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                    <div>
                        <h2 style="font-size: 18px; margin: 0;">Telegram Status</h2>
                        <span style="font-size: 13px; color: var(--text-muted);">Target: Engineering 2026 batch (-1002020152383)</span>
                    </div>
                    <span class="badge">● Authorized</span>
                </div>

                <form id="sync-form">
                    <div class="form-group">
                        <label for="webhook_url">Google Sheet Webhook URL</label>
                        <input type="url" id="webhook_url" name="webhook_url" placeholder="https://script.google.com/macros/s/.../exec" required>
                    </div>

                    <button type="submit" class="btn" id="sync-btn">⚡ Run Incremental Sync Now</button>
                </form>

                <div id="status-box"></div>
            </div>

            <div class="card" style="text-align: center;">
                <h3 style="font-size: 16px; margin-top: 0;">View Your Placement Data</h3>
                <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 16px;">Open your Google Sheet to view, filter, and export the real-time placement records.</p>
                <a id="sheet-link" href="#" target="_blank" class="btn btn-secondary" style="text-decoration: none;">📊 Open Google Sheet</a>
            </div>
        </div>

        <script>
            // Restore saved Webhook URL from localStorage
            const savedWebhook = localStorage.getItem('google_webhook_url');
            if (savedWebhook) {
                document.getElementById('webhook_url').value = savedWebhook;
                document.getElementById('sheet-link').href = savedWebhook.replace('/exec', '/edit');
            }

            document.getElementById('sync-form').addEventListener('submit', async (e) => {
                e.preventDefault();
                const webhookUrl = document.getElementById('webhook_url').value.trim();
                localStorage.setItem('google_webhook_url', webhookUrl);
                document.getElementById('sheet-link').href = webhookUrl.replace('/exec', '/edit');

                const btn = document.getElementById('sync-btn');
                const box = document.getElementById('status-box');

                btn.disabled = true;
                btn.innerText = '⌛ Syncing with Telegram...';
                box.style.display = 'block';
                box.innerText = 'Connecting to Telegram...\nFetching new messages since last sync...';

                try {
                    const res = await fetch('/api/sync', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: new URLSearchParams({ 'webhook_url': webhookUrl })
                    });
                    const data = await res.json();
                    
                    if (data.status === 'success') {
                        box.innerText = `✅ Sync Complete!\n\nNew Messages Processed: ${data.messages_processed}\nNew Placement Rows Added: ${data.rows_added}\nSkipped: ${data.skipped}\nLast Message ID: ${data.last_id}`;
                    } else {
                        box.innerText = `❌ Error: ${data.message}`;
                    }
                } catch (err) {
                    box.innerText = `❌ Request failed: ${err.message}`;
                } finally {
                    btn.disabled = false;
                    btn.innerText = '⚡ Run Incremental Sync Now';
                }
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.post("/api/sync")
async def sync_placements(webhook_url: str = Form(...)):
    if not webhook_url:
        return JSONResponse({"status": "error", "message": "Webhook URL is required."}, status_code=400)

    try:
        # 1. Query Google Sheet via Webhook to get current last_id
        req_get = urllib.request.Request(webhook_url, method="GET")
        with urllib.request.urlopen(req_get, timeout=10) as resp:
            sheet_meta = json.loads(resp.read().decode('utf-8'))
            last_id = int(sheet_meta.get("last_id", 0))
    except Exception as e:
        last_id = 0

    # 2. Connect to Telethon statelessly using StringSession
    try:
        session = StringSession(DEFAULT_SESSION_STRING)
        client = TelegramClient(session, DEFAULT_API_ID, DEFAULT_API_HASH)
        await client.connect()

        if not await client.is_user_authorized():
            return JSONResponse({"status": "error", "message": "Telegram Session invalid or expired."}, status_code=401)

        group_entity = int(DEFAULT_GROUP_ID) if str(DEFAULT_GROUP_ID).lstrip('-').isdigit() else DEFAULT_GROUP_ID
        messages = []
        async for msg in client.iter_messages(group_entity, min_id=last_id, reverse=True):
            if msg.text and not msg.action:
                messages.append(msg)

        await client.disconnect()
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Telegram API error: {str(e)}"}, status_code=500)

    if not messages:
        return JSONResponse({
            "status": "success",
            "messages_processed": 0,
            "rows_added": 0,
            "skipped": 0,
            "last_id": last_id
        })

    # 3. Parse placement messages
    new_rows = []
    skipped = 0
    for msg in messages:
        if not is_placement_post(msg.text):
            skipped += 1
            continue
        new_rows.append(parse_message(msg))

    max_id = max(m.id for m in messages)

    # 4. Post parsed rows + max_id to Google Sheet Webhook
    try:
        payload = json.dumps({
            "rows": new_rows,
            "last_id": max_id
        }).encode('utf-8')
        
        req_post = urllib.request.Request(
            webhook_url,
            data=payload,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_post, timeout=30) as resp_post:
            res_data = json.loads(resp_post.read().decode('utf-8'))

        return JSONResponse({
            "status": "success",
            "messages_processed": len(messages),
            "rows_added": len(new_rows),
            "skipped": skipped,
            "last_id": max_id
        })
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Failed to post to Google Sheets: {str(e)}"}, status_code=500)
