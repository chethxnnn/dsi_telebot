# -*- coding: utf-8 -*-
"""
Telegram to Excel Scraper for Placement Posts

This script scrapes job postings from a Telegram group and outputs them into a structured Excel workbook.
"""

import re
import os
import asyncio
from datetime import datetime
from telethon import TelegramClient
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# Configuration Block
# ==========================================
API_ID   = int(os.environ.get("TELEGRAM_API_ID", 0))
API_HASH = os.environ.get("TELEGRAM_API_HASH", "YOUR_API_HASH")
GROUP    = os.environ.get("TELEGRAM_GROUP_ID", "-1002020152383")
EXCEL    = 'placements.xlsx'
LAST_ID  = 'last_id.txt'

# ==========================================
# 1. Last ID Tracking
# ==========================================
def read_last_id():
    if os.path.exists(LAST_ID):
        try:
            with open(LAST_ID, 'r', encoding='utf-8') as f:
                return int(f.read().strip())
        except ValueError:
            pass
    return 0

def write_last_id(msg_id):
    with open(LAST_ID, 'w', encoding='utf-8') as f:
        f.write(str(msg_id))

# ==========================================
# 2. Message Fetcher
# ==========================================
async def fetch_new_messages(client, last_id):
    messages = []
    group_entity = int(GROUP) if str(GROUP).lstrip('-').isdigit() else GROUP
    async for msg in client.iter_messages(group_entity, min_id=last_id, reverse=True):
        if msg.text and not msg.action:
            messages.append(msg)
    return messages

# ==========================================
# 3. Placement Post Filter
# ==========================================
def is_placement_post(msg):
    text = msg.text or ""
    if len(text) < 80:
        return False
    if msg.reply_to is not None:
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

# ==========================================
# 4. Parsers
# ==========================================
def extract_company(text):
    """Extract company name using labeled patterns, then careful heuristics."""
    
    # Blacklist of invalid company names / noisy section titles
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
        # Remove common noise phrases from start
        name = re.sub(r'^(?:Greetings?\s+from|Hurry\s+Up\s*[\-\:]?|Alert\s*[\-\:]?|Attention\s+to\s+All\s+(?:Registered\s+)?Students\s+(?:for\s+)?|Important\s+Note\s+(?:for\s+[\w\s]+\s+on\s+)?)\s*', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'\s*-\s*Pune\'s\s+Top.*$', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'\s*[\-\:]\s*All\s+Branches.*$', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'[^\w\s\-\&\.\,\']', '', name).strip('- ').strip()
        
        if name.lower() in invalid_companies or any(inv in name.lower() for inv in ['registration', 'updates', 'reminder', 'highlight', 'webinar']):
            return None
        if len(name.split()) > 7 or len(name) > 60:
            return None
        return name

    # Priority 1: Explicit "Company: X" label
    m1 = re.search(r'Company\s*:\s*(.+)', text, re.IGNORECASE)
    if m1:
        val = clean_comp_name(m1.group(1).strip().split('\n')[0].strip())
        if val:
            return val
    
    # Priority 2: "Greetings from X"
    m2 = re.search(r'Greetings from\s+(.+?)[\.!\n]', text, re.IGNORECASE)
    if m2:
        val = clean_comp_name(m2.group(1).strip())
        if val:
            return val
    
    # Priority 3: "X is hiring" / "X is currently hiring"
    m3 = re.search(r'^(.+?)\s+is\s+(?:currently\s+)?(?:hiring|recruiting)', text, re.IGNORECASE | re.MULTILINE)
    if m3:
        val = clean_comp_name(m3.group(1).strip())
        if val:
            return val
    
    # Priority 4: "X wants to hire"
    m4 = re.search(r'^(.+?)\s+wants to hire', text, re.IGNORECASE | re.MULTILINE)
    if m4:
        val = clean_comp_name(m4.group(1).strip())
        if val:
            return val
    
    # Priority 5: "X Campus drive" or "X Campus Recruitment" in first line
    first_line = text.split('\n')[0].strip()
    m5 = re.match(r'^[^\w]*([\w][\w\s\-\&\.]+?)\s*(?:Campus\s+(?:drive|recruitment)|Recruitment|Off[- ]?Campus)', first_line, re.IGNORECASE)
    if m5:
        val = clean_comp_name(m5.group(1))
        if val:
            return val
    
    # Priority 6: First line ONLY if it looks like a clean company name
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
            if val:
                return val
    
    return None

def extract_roles(text):
    """Extract job role/title names — NOT responsibilities, instructions, or URLs."""
    roles = []

    def is_valid_role(r):
        if not r:
            return False
        r_str = r.strip()
        # Reject URLs, links, or file paths
        if re.search(r'https?://|www\.|drive\.google\.com|forms\.gle|\.php|\.pdf', r_str, re.IGNORECASE):
            return False
        # Reject sentences / long instructions / responsibilities
        if len(r_str.split()) > 8 or len(r_str) > 70:
            return False
        if any(w in r_str.lower() for w in ['strictly', 'unplaced', 'eligible', 'selection', 'written test', 'interview', 'round', 'responsibility', 'responsibilities', 'criteria', 'note:']):
            return False
        return True

    # Priority 1: Explicit "Job Title:", "Role:", "Position:", "Designation:", "Profile:"
    m1 = re.search(r'(?:Job Title|Position|Profile|Designation|Role Title)\s*:\s*(.+)', text, re.IGNORECASE)
    if m1:
        val = m1.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val):
            roles.append(val)
            return roles
        
    # Priority 2: "Job Description – X" (role name after dash)
    m2 = re.search(r'Job Description\s*[–\-:]\s*(.+)', text, re.IGNORECASE)
    if m2:
        val = m2.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val):
            roles.append(val)
            return roles

    # Priority 3: "Role: X" (explicit label, not "Role Overview" or "Role & Responsibilities")
    for m in re.finditer(r'(?:^|\n)\s*Role\s*:\s*(.+)', text, re.IGNORECASE):
        val = m.group(1).strip().split('\n')[0].strip()
        if is_valid_role(val) and 'overview' not in val.lower():
            roles.append(val)
    if roles:
        return roles[:5]

    # Priority 4: Numbered items like "1. Software Engineer:" or "1) Software Engineer"
    for line in text.split('\n'):
        m = re.match(r'^\s*\d+[\.\)]\s*([A-Za-z0-9\s\-\/\(\)\&]+?)(?:\s*:|\s*–|\s*-|$)', line)
        if m:
            val = m.group(1).strip()
            if is_valid_role(val) and any(kw in val.lower() for kw in ['engineer', 'developer', 'associate', 'executive', 'analyst', 'trainee', 'intern', 'manager', 'lead', 'consultant', 'bda', 'boe', 'sde', 'get']):
                roles.append(val)
    if roles:
        return roles[:5]

    # Priority 5: Bullet points ONLY in a "hiring for" / "openings" / "positions" section
    lines = text.split('\n')
    role_section = False
    for line in lines:
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
                if is_valid_role(candidate):
                    roles.append(candidate)
            elif line.strip() and not re.match(r'^\s*[•\-\*]', line, re.UNICODE):
                role_section = False
    if roles:
        return roles[:5]

    # Priority 6: Fallback scan for standard job titles in placement posts
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

def extract_salary(text):
    result = {'probation': None, 'post_probation': None, 'raw': None}
    
    # Extract raw salary block
    raw_match = re.search(
        r'(?:Compensation(?: Structure)?|Salary|CTC)\s*:\s*(.+?)(?=\n\s*(?:[A-Z][a-z]+|Eligibility|Location|Selection|Key Resp)\s*:|$)',
        text, re.IGNORECASE | re.DOTALL
    )
    if raw_match:
        result['raw'] = raw_match.group(1).strip()[:300]  # cap raw at 300 chars
    else:
        m_sal = re.search(r'(?:Salary|CTC)\s*:\s*(.+)', text, re.IGNORECASE)
        if m_sal:
            result['raw'] = m_sal.group(1).strip()
            
    # Priority 1: Probation vs Post Probation (check FIRST — more specific)
    m_prob = re.search(r'During Probation[^:]*:\s*₹?\s*([\d,]+)\s*(?:per month|p\.?m\.?|/month)?', text, re.IGNORECASE | re.UNICODE)
    m_post = re.search(r'Post Probation[^:]*:\s*₹?\s*([\d,]+)\s*(?:per month|p\.?m\.?|/month)?', text, re.IGNORECASE | re.UNICODE)
    if m_prob and m_post:
        result['probation'] = f'₹{m_prob.group(1).strip()}/month'
        result['post_probation'] = f'₹{m_post.group(1).strip()}/month'
        # Also check for LPA in post-probation line
        m_post_lpa = re.search(r'Post Probation[^\n]*(\d+(?:\.\d+)?)\s*LPA', text, re.IGNORECASE)
        if m_post_lpa:
            result['post_probation'] = f'{m_post_lpa.group(1)} LPA'
        return result
    
    # Priority 2: Flat salary in Salary/CTC line
    flat_match = re.search(r'(?:Salary|CTC)\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if flat_match:
        val = flat_match.group(1).strip()
        result['probation'] = val
        result['post_probation'] = val
        return result
    
    # Fallback: LPA values (annual) — keep separate from monthly ₹ amounts
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

def extract_locations(text):
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
                
    # Deduplicate (case-insensitive)
    seen = set()
    dedup = []
    for loc in locations:
        if loc.lower() not in seen:
            seen.add(loc.lower())
            dedup.append(loc)
            
    return ", ".join(dedup) if dedup else None

def extract_eligibility(text):
    # Match 'Eligibility Criteria:' or 'Eligibility:' or just 'Eligibility' on its own line
    m = re.search(r'Eligibility\s*(?:Criteria)?\s*:?\s*\n?(.*?)(?=\n\s*(?:[A-Z][a-zA-Z\s]+:|Required Skills|Compensation|Selection Process|About|Key Resp)|\Z)', text, re.IGNORECASE | re.DOTALL)
    if m:
        content = m.group(1).strip()
        if not content:
            return None
        # Clean up bullet points
        content = re.sub(r'^\s*[•\-\*]\s*', '', content, flags=re.MULTILINE | re.UNICODE)
        content = re.sub(r'\n+', ' | ', content)
        content = content.strip(' |')
        return content if len(content) > 5 else None
    return None

def extract_links(text):
    urls = re.findall(r'https?://[^\s<>"\'\)]+', text, re.IGNORECASE)
    return " | ".join(urls) if urls else None

def extract_deadline(text):
    # Try specific patterns first — 'by' alone is too common
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
            if len(val) > 5:  # avoid noise like single words
                return val
    return None

# ==========================================
# 5. Main Parser
# ==========================================
def parse_message(msg):
    """Parse a single message into ONE row dict."""
    text = msg.text or ""
    company = extract_company(text)
    roles = extract_roles(text)
    salary = extract_salary(text)
    location = extract_locations(text)
    eligibility = extract_eligibility(text)
    links = extract_links(text)
    deadline = extract_deadline(text)
    
    # Join multiple roles into one cell with ' | ' separator
    role_str = ' | '.join(roles) if roles else None
    
    chat_id = str(msg.chat_id).replace('-100', '') if msg.chat_id else ''
    
    row = {
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
    return row

# ==========================================
# 6. Excel Handler
# ==========================================
def ensure_excel(filepath):
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = Workbook()
    ws = wb.active
    ws.title = 'Placements'
    headers = [
        'Date', 'Company', 'Role', 'Salary (Probation)', 'Salary (Post-Probation)', 
        'Salary (Raw)', 'Location', 'Eligibility', 'Registration Link', 
        'Deadline', 'Needs Review', 'Message Link', 'Raw Message'
    ]
    ws.append(headers)
    
    # Header styling
    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        bottom=Side(style='thin', color='2F5496')
    )
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30
    
    return wb, ws

def append_rows(ws, rows):
    count = 0
    keys = [
        'date', 'company', 'role', 'salary_probation', 'salary_post_probation',
        'salary_raw', 'location', 'eligibility', 'registration_link', 'deadline',
        'needs_review', 'message_link', 'raw_message'
    ]
    # Alternating row colors
    light_fill = PatternFill(start_color='F2F7FB', end_color='F2F7FB', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    for r in rows:
        ws.append([r[k] for k in keys])
        row_num = ws.max_row
        fill = light_fill if row_num % 2 == 0 else white_fill
        for cell in ws[row_num]:
            cell.fill = fill
            cell.alignment = Alignment(vertical='top', wrap_text=False)
        count += 1
    return count

def format_columns(ws):
    """Apply column-specific widths and alignment to fit content cleanly."""
    # Column widths tuned per field
    col_config = {
        'A': {'width': 12, 'align': 'center'},     # Date
        'B': {'width': 22, 'align': 'left'},        # Company
        'C': {'width': 32, 'align': 'left'},        # Role
        'D': {'width': 20, 'align': 'center'},      # Salary (Probation)
        'E': {'width': 22, 'align': 'center'},      # Salary (Post-Probation)
        'F': {'width': 30, 'align': 'left'},        # Salary (Raw)
        'G': {'width': 28, 'align': 'left'},        # Location
        'H': {'width': 35, 'align': 'left'},        # Eligibility
        'I': {'width': 40, 'align': 'left'},        # Registration Link
        'J': {'width': 22, 'align': 'center'},      # Deadline
        'K': {'width': 14, 'align': 'center'},      # Needs Review
        'L': {'width': 35, 'align': 'left'},        # Message Link
        'M': {'width': 50, 'align': 'left'},        # Raw Message
    }
    
    for col_letter, config in col_config.items():
        ws.column_dimensions[col_letter].width = config['width']
        align = Alignment(horizontal=config['align'], vertical='top', wrap_text=(col_letter in ('F', 'H', 'M')))
        for row in range(2, ws.max_row + 1):
            cell = ws[f'{col_letter}{row}']
            cell.alignment = align
    
    # Highlight "Yes" in Needs Review column with a soft red
    review_fill = PatternFill(start_color='FCE4E4', end_color='FCE4E4', fill_type='solid')
    review_font = Font(color='C00000', bold=True)
    for row in range(2, ws.max_row + 1):
        cell = ws[f'K{row}']
        if cell.value == 'Yes':
            cell.fill = review_fill
            cell.font = review_font

# ==========================================
# 7. Main Flow
# ==========================================
client = TelegramClient('session', API_ID, API_HASH)

async def main():
    await client.start()
    
    last_id = read_last_id()
    print('Fetching messages...')
    messages = await fetch_new_messages(client, last_id)
    
    if not messages:
        print('No new messages.')
        return
        
    print(f'Processing {len(messages)} messages...')
    
    all_rows = []
    skipped = 0
    for msg in messages:
        if not is_placement_post(msg):
            skipped += 1
            continue
        row = parse_message(msg)
        all_rows.append(row)
        
    max_id = max(m.id for m in messages)
    
    if all_rows:
        wb, ws = ensure_excel(EXCEL)
        count = append_rows(ws, all_rows)
        format_columns(ws)
        try:
            wb.save(EXCEL)
            print(f'Added {count} rows to {EXCEL}')
            write_last_id(max_id)  # Only update after successful save
            print(f'Processed {len(messages)} messages ({skipped} skipped). Last ID: {max_id}')
        except PermissionError:
            print(f"Error: {EXCEL} is locked. Please close it in Excel and try again.")
            print("Last ID NOT updated — next run will re-process these messages.")
    else:
        print('No placement posts found in new messages.')
        write_last_id(max_id)  # No rows but still advance the pointer
        print(f'Processed {len(messages)} messages ({skipped} skipped). Last ID: {max_id}')
    
    print('Done.')

if __name__ == "__main__":
    with client:
        client.loop.run_until_complete(main())
