"""
Resume Hub Cloud — Mobile job queue + AI resume builder + Google Drive save
Deploy to Render.com.
Required env vars: ANTHROPIC_API_KEY, GOOGLE_CREDENTIALS_JSON (base64), GOOGLE_DRIVE_FOLDER_ID
Optional env vars: GMAIL_USER, GMAIL_PASS (to also email the resume)
"""
import os, sqlite3, threading, re, smtplib, time, subprocess, tempfile, base64, json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, jsonify, g

BASE       = os.path.dirname(os.path.abspath(__file__))
DB_PATH    = '/tmp/resume_queue.db'
TEMPLATE   = os.path.join(BASE, 'master_template.docx')
TMP_DIR    = tempfile.gettempdir()
DEFAULT_TO = 'ksunilmech29@gmail.com'


# ── GitHub DB Persistence (survives Render redeploys) ─────────────────────
GITHUB_REPO     = 'ksunilmech29-ship-it/resume-hub'
BACKUP_FILE     = 'docs/cloud_jobs_backup.json'
GITHUB_RAW_URL  = f'https://raw.githubusercontent.com/{GITHUB_REPO}/main/{BACKUP_FILE}'
GITHUB_API_URL  = f'https://api.github.com/repos/{GITHUB_REPO}/contents/{BACKUP_FILE}'
_gh_lock        = threading.Lock()

def _gh_token():
    token = (os.environ.get('RESUME_HUB_TOKEN')
             or os.environ.get('Resume_Hub_Token')
             or os.environ.get('resume_hub_token', ''))
    if not token:
        try:
            conn = sqlite3.connect(DB_PATH)
            row = conn.execute("SELECT value FROM settings WHERE key='resume_hub_token'").fetchone()
            if row: token = row[0]
            conn.close()
        except Exception:
            pass
    return token

def restore_jobs_from_github():
    """Restore jobs from GitHub backup if local DB is empty."""
    import urllib.request as _ur
    try:
        conn = sqlite3.connect(DB_PATH)
        count = conn.execute('SELECT COUNT(*) FROM jobs').fetchone()[0]
        conn.close()
        if count > 0:
            return  # DB already has data
    except Exception:
        return

    try:
        with _ur.urlopen(GITHUB_RAW_URL, timeout=10) as resp:
            data = json.loads(resp.read())
        jobs = data.get('jobs', [])
        if not jobs:
            return
        conn = sqlite3.connect(DB_PATH)
        for job in jobs:
            conn.execute(
                'INSERT OR IGNORE INTO jobs (id,title,company,url,jd_text,status,build_log,drive_link,created_at,built_at) VALUES (?,?,?,?,?,?,?,?,?,?)',
                (job.get('id'), job.get('title',''), job.get('company',''), job.get('url',''),
                 job.get('jd_text',''), job.get('status','pending'), job.get('build_log',''),
                 job.get('drive_link',''), job.get('created_at',''), job.get('built_at',''))
            )
        conn.commit()
        conn.close()
        print(f'[persistence] Restored {len(jobs)} jobs from GitHub backup')
    except Exception as e:
        print(f'[persistence] Restore failed: {e}')

def backup_jobs_to_github():
    """Push current jobs to GitHub so they survive redeploys."""
    token = _gh_token()
    if not token:
        return
    import urllib.request as _ur
    with _gh_lock:
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            jobs = [dict(r) for r in conn.execute('SELECT * FROM jobs ORDER BY id').fetchall()]
            conn.close()

            payload = json.dumps({'jobs': jobs, 'backed_up_at': datetime.now().isoformat()}, ensure_ascii=False)
            content_b64 = base64.b64encode(payload.encode('utf-8')).decode()

            # Get current file SHA (needed for update)
            sha = ''
            try:
                req = _ur.Request(GITHUB_API_URL,
                    headers={'Authorization': f'token {token}', 'Accept': 'application/vnd.github.v3+json'})
                with _ur.urlopen(req, timeout=8) as resp:
                    sha = json.loads(resp.read()).get('sha', '')
            except Exception:
                pass

            body = {'message': f'backup: {len(jobs)} jobs', 'content': content_b64, 'branch': 'main'}
            if sha:
                body['sha'] = sha

            data = json.dumps(body).encode()
            req = _ur.Request(GITHUB_API_URL, data=data, method='PUT',
                headers={'Authorization': f'token {token}',
                         'Content-Type': 'application/json',
                         'Accept': 'application/vnd.github.v3+json'})
            _ur.urlopen(req, timeout=15)
        except Exception as e:
            print(f'[persistence] Backup failed: {e}')

def _backup_async():
    threading.Thread(target=backup_jobs_to_github, daemon=True).start()

def _backup_heartbeat():
    """Backup every 5 minutes unconditionally — belt-and-suspenders."""
    import time as _t
    _t.sleep(30)
    while True:
        try:
            backup_jobs_to_github()
        except Exception:
            pass
        _t.sleep(300)

app = Flask(__name__)

# ── Database ──────────────────────────────────────────────────────────────────
# Init runs at import time so Gunicorn workers have the schema ready

SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT NOT NULL,
    company     TEXT NOT NULL,
    url         TEXT DEFAULT '',
    jd_text     TEXT DEFAULT '',
    status      TEXT DEFAULT 'pending',
    build_log   TEXT DEFAULT '',
    drive_link  TEXT DEFAULT '',
    created_at  TEXT DEFAULT (datetime('now','localtime')),
    built_at    TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS board_jobs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT NOT NULL,
    company     TEXT DEFAULT '',
    location    TEXT DEFAULT '',
    job_url     TEXT DEFAULT '',
    date_posted TEXT DEFAULT '',
    source      TEXT DEFAULT 'LinkedIn Alert',
    description TEXT DEFAULT '',
    dedup_key   TEXT UNIQUE,
    imported_at TEXT DEFAULT (datetime('now','localtime'))
);
"""

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(_):
    db = g.pop('db', None)
    if db:
        db.close()

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()

init_db()  # runs at import time
restore_jobs_from_github()  # restore from GitHub if DB was wiped
threading.Thread(target=_backup_heartbeat, daemon=True).start()  # backup every 5 min

def get_setting(key, default=''):
    row = get_db().execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
    return row['value'] if row else default

def set_setting(key, value):
    get_db().execute('INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)', (key, value))
    get_db().commit()

def _get_setting_direct(db_path, key, default=''):
    conn = sqlite3.connect(db_path)
    row  = conn.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
    conn.close()
    return row[0] if row else default

# ── Candidate profile ─────────────────────────────────────────────────────────

PROFILE = """
==========================================================================
GOLD STANDARD REFERENCE: REFERENCE_FORMAT_RESUME.docx (same folder)
Match its format EXACTLY — especially these 3 elements:
  1. Career Highlights: 5-column borderless table, pipes aligned, 32pt stats
  2. Education: always on page 1 after Experience
  3. Page 2: filled to within 2 lines of bottom, no empty space
==========================================================================

==========================================================================
STEP 1 — READ THE JD FIRST. BEFORE WRITING ANYTHING:
Extract from the JD: the core role theme, key responsibilities, required
skills, and priority keywords. Every section of this resume — header
keywords, summary, competency headings, experience descriptions, and KA
section headings — must reflect those JD themes. The resume reader must
immediately see that Sunil matches what the JD asks for.
==========================================================================

NAME: SUNIL KUMAR K  (three words, K is mandatory — never drop it)
Contact: +91 9741114967 | ksunilmech29@gmail.com | linkedin.com/in/sunil-kumar-k-8b1b72a
Experience: 19+ years, Fintech and Financial Services domain

==========================================================================
MASTER TEMPLATE — STRUCTURE ONLY, NEVER COPY ITS TEXT
==========================================================================
The template is a Delivery Leader resume. Use it ONLY for XML structure,
fonts, colors, spacing, table layout, and background colors.
Replace EVERY paragraph with JD-driven content. Do not carry over any
phrase, heading, or bullet from the template text.

==========================================================================
TARGET ROLE vs PAST EMPLOYERS
==========================================================================
Title and Company in this request = the job Sunil is APPLYING FOR.
Never add the target company to the Experience section.

==========================================================================
CAREER HISTORY — exactly these 4, most recent first
==========================================================================
1. Open Financial Technologies | Vice President | Apr 2022 - Present
   Fintech platform company. Adapt title and 2-sentence description to JD.

2. Attra | Program Manager | May 2018 - Mar 2022
   NEVER "Synechron", "Syntel", or "Program Director".
   US-based clients, SaaS, fintech delivery. Adapt 2-sentence description to JD.

3. Deutsche Bank | Assistant Vice President | Nov 2009 - Apr 2018
   Multi-region: India, EU, US. IB Division. Adapt 2-sentence description to JD.

4. Accenture | Software Engineer | Nov 2006 - Oct 2009
   VERBATIM: "Started as a QA Engineer and progressively contributed to product
   and delivery, built user stories, supported product definition and process
   standardization across Banking, Insurance and Healthcare domain."

==========================================================================
EDUCATION AND CERTIFICATIONS — exactly these 4, verbatim, always Page 1
==========================================================================
1. Executive Global Management Programme  -  IIM Calcutta  |  2014 - 2016
2. Bachelor of Engineering  -  Visvesvaraya Technological University  |  2002 - 2006
3. SAFe 5 Agilist  -  Scaled Agile Inc.  |  Dec 2022
4. Certified Scrum Master (CSM)  -  Scrum Alliance
Never add PMP, Prince2, TOGAF, CFA, or any other certification.

Domain: always Fintech and Financial Services.

==========================================================================
PAGE 1 FORMAT
==========================================================================
HEADER — 3 lines:
  Line 1 [BG F0F7FB, ALL CAPS bold Navy]: SUNIL KUMAR K
  Line 2 [BG F0F7FB, teal]:
    [Safe generic senior title — NOT the verbatim job posting title. Derive from JD theme.
     Use "Head of [Theme]" or "Vice President | [Theme]" form. Examples:
       JD = "Account Director" → "Head of Account Management"
       JD = "Director Customer Success" → "Head of Customer Success"
       JD = "Associate Director Growth" → "Head of Growth"
     Never copy the exact job posting title. Never expose level mismatches.
     Do NOT alter any Experience designation (VP, Program Manager, AVP, etc.)]
    |  [3-5 keywords drawn directly from the JD]
    Keywords must reflect what the JD asks for — not generic phrases.
  Line 3 [no BG, gray]:
    Bangalore, India  |  +91 9741114967  |  ksunilmech29@gmail.com  |  linkedin.com/in/sunil-kumar-k-8b1b72a

PROFESSIONAL SUMMARY — 2 paragraphs, min 6 lines:
  Every sentence must address what the JD asks for.
  Para 1: positioning + domain + years + leadership angle matching JD.
  Para 2: quantified outcomes relevant to the JD requirements.
  No em dashes, no semicolons.

CAREER HIGHLIGHTS — BUILD AS A 5-COLUMN BORDERLESS TABLE (not paragraphs):
  Structure: [item] [sep] [item] [sep] [item]  — 2 table rows, BG E3F2F8 throughout
  Column widths (DXA): 3300 | 378 | 3300 | 378 | 3300  (total = 10656)
  All borders: none. All cells: center-aligned. Table shading: E3F2F8.

  Item cells — each cell contains ONE paragraph, center-aligned, with 2 runs:
    Run 1: STAT   — bold, Navy 1A3557, font size 14pt (sz=28 half-points)
    Run 2: LABEL  — italic, Teal 0D6E8A, font size 10pt (sz=20 half-points), space-prefixed
    LABEL must be MAX 2 WORDS — never 3. Stat+label combined must stay under 20 chars.

  Separator cells — ONE paragraph, text "|":
    Color: Light gray 888888, font size 11pt (sz=22), center-aligned

  Row spacing: space_before=80, space_after=20 for row 1
               space_before=20, space_after=80 for row 2

  Content: 2 rows x 3 items = 6 stats chosen to match JD priorities.
  NEVER use paragraphs. NEVER use ◆. NEVER left-align. ALWAYS use this 5-column table.

CORE COMPETENCIES — 3-column borderless table:
  3 column headings derived from the JD's key requirement areas.
  Each column: 1 bold teal heading + up to 5 dash-prefixed bullets (3-5 words).

PROFESSIONAL EXPERIENCE — 4 roles, company line [BG F5FAFD]:
  Each 2-sentence description must emphasise the angle the JD cares about.
  Domain label in description must match the JD context.
  Max 2 sentences per role. Accenture always verbatim.

PAGE 1 MUST CONTAIN — in this exact order, all fitting within page 1:
  1. Header (Name + Role + Contact)
  2. Professional Summary (2 paragraphs)
  3. Career Highlights (2 rows)
  4. Core Competencies (3-col table)
  5. Professional Experience (4 roles, 2 sentences each)
  6. Education and Certifications (all 4 entries)

EDUCATION MUST ALWAYS BE ON PAGE 1 — THIS IS A HARD RULE.
If page 1 is overflowing:
  - Cut Summary to 2 very tight sentences per paragraph (4 lines total)
  - Cut each Experience description to 1-2 lines maximum
  - NEVER remove Education. NEVER push Education to page 2.
  - Reduce font size of body text to 9pt if needed, but keep Education on page 1.

==========================================================================
PAGE 2 FORMAT
==========================================================================
Page break before Key Achievements. Page 2 = Key Achievements ONLY.

KEY ACHIEVEMENTS — section headings [BG EBF4FF, Navy ALL CAPS]:
  Headings must come from the JD's key themes and requirements.
  Order: most JD-relevant sections first.

Each bullet:
  â—  Bold Sub-heading:  Sentence 1 with outcome. Sentence 2 with metric.
  200-250 chars, varied opening verbs (Led/Drove/Designed/Established/
  Owned/Championed/Developed/Governed/Delivered/Managed/Scaled/Directed).
  Never repeat same verb more than twice. Include metrics in every bullet.

PAGE 2 FILL RULE — MANDATORY, NO EXCEPTIONS:
Page 2 MUST be filled to within 2 lines of the bottom edge.
A page 2 with empty space at the bottom = AUTOMATIC FAIL.

COUNTING RULE: A standard page 2 holds approximately 55-60 lines of body text.
After writing your bullets, count the lines used. If below 53 lines, ADD MORE CONTENT.

ACTION SEQUENCE — execute in order until page 2 is full:
  STEP 1: Extend the last 2-3 bullets — add one extra sentence to each (~80 chars).
  STEP 2: If still not full, add a new bullet to the last section.
  STEP 3: If still not full, add a new section "STAKEHOLDER ENGAGEMENT AND
           CROSS-FUNCTIONAL LEADERSHIP" with 2 fully written bullets.
  STEP 4: If still not full, add another section "DELIVERY EXCELLENCE AND
           OPERATIONAL GOVERNANCE" with 2 bullets.

Keep executing steps until page 2 is full. Do not stop at STEP 1 if it is not enough.
Page 2 must NEVER overflow to a third page.

==========================================================================
HARD RULES — THESE ARE NON-NEGOTIABLE. VIOLATION = AUTOMATIC FAIL.
==========================================================================
NAME (HARDEST RULE):
  The name paragraph MUST read exactly: SUNIL KUMAR K
  All three words. The K is mandatory. Never "Sunil Kumar". Never "SUNIL KUMAR".
  In your XML: <w:t>SUNIL KUMAR K</w:t>   NOT <w:t>SUNIL KUMAR</w:t>
  Check this before saving. If it says SUNIL KUMAR without K, FIX IT before proceeding.

ATTRA TITLE (HARD RULE):
  Attra job title MUST be exactly: Program Manager
  NEVER write "Program Director" for Attra. Not in the experience header. Not anywhere.
  The experience line must read: Attra   Program Manager   May 2018 - Mar 2022

ACCENTURE DESCRIPTION (VERBATIM — COPY EXACTLY):
  "Started as a QA Engineer and progressively contributed to product and delivery, built user stories, supported product definition and process standardization across Banking, Insurance and Healthcare domain."
  Copy character-for-character. No paraphrasing.

CAREER HIGHLIGHTS (HARD RULE — MUST BE 5-COLUMN TABLE):
  NEVER write Career Highlights as paragraphs with symbols like ◆ or •.
  ALWAYS build a 5-column borderless w:tbl with col widths 3300|378|3300|378|3300 DXA.
  2 rows, 3 items per row = 6 stat+label cells.
  Item cells: stat bold Navy 14pt + label italic Teal 10pt. Separator cells: "|" gray 11pt.
  If you find yourself writing "19+ Years ◆ 30% Faster" as a paragraph, STOP. Build the table.

PAGE 2 FILL (HARD RULE):
  Page 2 MUST be filled to within 2 lines of the bottom. Count your lines.
  Standard page 2 = 55-60 lines. If below 53 lines, ADD MORE CONTENT — do not stop early.
  Add sections "STAKEHOLDER ENGAGEMENT AND CROSS-FUNCTIONAL LEADERSHIP" and
  "DELIVERY EXCELLENCE AND OPERATIONAL GOVERNANCE" with 2 bullets each if needed.
  Keep adding until page 2 is full.

OTHER HARD RULES:
- 2 pages, both fully filled
- No em dashes, no semicolons anywhere in the document
- No tables except Core Competencies 3-column table and Career Highlights 5-column table
- No certifications beyond the canonical 4
- Colors: Navy 1A3557, Teal 0D6E8A, BG F0F7FB/E3F2F8/F5FAFD/EBF4FF
- Margins: 576 DXA top/bottom, 792 DXA left/right
- $20M only if JD explicitly asks for P&L or revenue ownership

FINAL VALIDATION CHECKLIST — run through this before saving the file:
  [1] Name paragraph reads exactly: SUNIL KUMAR K (three words, K present)
  [2] Attra title reads: Program Manager (NOT Program Director)
  [3] Accenture text starts with: "Started as a QA Engineer..."
  [4] Career Highlights is a 5-column borderless w:tbl (NOT paragraphs with ◆)
  [5] Page 2 is full, within 2 lines of the bottom
  [6] No em dashes, no semicolons in any paragraph
  [7] All 4 canonical certifications present on page 1
==========================================================================
- Name line 1: SUNIL KUMAR K (all 3 words, ALL CAPS) — missing K = fail
- 2 pages, both fully filled
- No em dashes, no semicolons
- No tables except Core Competencies
- Attra: never Synechron/Syntel/Program Director
- No certifications beyond canonical 4
- Colors: Navy 1A3557, Teal 0D6E8A, BG F0F7FB/E3F2F8/F5FAFD/EBF4FF
- Margins: 576 DXA top/bottom, 792 DXA left/right
- $20M only if JD explicitly asks for P&L/revenue ownership
- Gap check internally — do not pause to ask. Proceed to build.
- No spurious commas: never write "X, with" or trailing commas before conjunctions/prepositions. Proofread ALL punctuation before saving.
- Validate all rules. Must reach 10/10 before saving.
==========================================================================
"""

# ── Google Drive ──────────────────────────────────────────────────────────────

def _get_drive_service():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    creds = Credentials(
        token=None,
        refresh_token=os.environ['GOOGLE_REFRESH_TOKEN'],
        client_id=os.environ['GOOGLE_CLIENT_ID'],
        client_secret=os.environ['GOOGLE_CLIENT_SECRET'],
        token_uri='https://oauth2.googleapis.com/token',
    )
    return build('drive', 'v3', credentials=creds)

def _get_or_create_folder(service, name, parent_id):
    """Get existing subfolder by name or create it under parent_id."""
    query = (f"name='{name}' and mimeType='application/vnd.google-apps.folder' "
             f"and '{parent_id}' in parents and trashed=false")
    results = service.files().list(q=query, fields='files(id)').execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    folder = service.files().create(
        body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]},
        fields='id'
    ).execute()
    return folder['id']

def upload_to_drive(file_path, filename, job=None):
    from googleapiclient.http import MediaFileUpload
    root_folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID', '')
    service = _get_drive_service()

    # Build subfolder name: YYYY-MM-DD_Title_Company
    date_str = datetime.now().strftime('%Y-%m-%d')
    if job:
        def slug(s): return re.sub(r'[^A-Za-z0-9]', '_', s).strip('_')
        title_slug   = '_'.join(slug(job.get('title', '')).split('_')[:4])
        company_slug = '_'.join(slug(job.get('company', '')).split('_')[:2])
        subfolder_name = f"{date_str}_{title_slug}_{company_slug}"
    else:
        subfolder_name = date_str

    # Create or reuse subfolder inside root Drive folder
    if root_folder_id:
        target_folder_id = _get_or_create_folder(service, subfolder_name, root_folder_id)
    else:
        target_folder_id = None

    meta = {
        'name': filename,
        'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }
    if target_folder_id:
        meta['parents'] = [target_folder_id]

    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    f = service.files().create(body=meta, media_body=media, fields='id,webViewLink').execute()
    service.permissions().create(
        fileId=f['id'],
        body={'role': 'reader', 'type': 'anyone'}
    ).execute()
    download_url = f'https://drive.google.com/uc?export=download&id={f["id"]}'

    # Upload job_info.txt with full job details including URL
    if job and target_folder_id:
        import io
        info = (
            f"Job Title : {job.get('title', '')}\n"
            f"Company   : {job.get('company', '')}\n"
            f"Job URL   : {job.get('url', '(not provided)')}\n"
            f"Built On  : {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"Resume    : {filename}\n"
        )
        from googleapiclient.http import MediaIoBaseUpload
        info_media = MediaIoBaseUpload(
            io.BytesIO(info.encode('utf-8')),
            mimetype='text/plain'
        )
        info_file = service.files().create(
            body={'name': 'job_info.txt', 'parents': [target_folder_id]},
            media_body=info_media,
            fields='id'
        ).execute()
        service.permissions().create(
            fileId=info_file['id'],
            body={'role': 'reader', 'type': 'anyone'}
        ).execute()

    return f.get('webViewLink', ''), download_url

# ── Resume builder ────────────────────────────────────────────────────────────

def safe_filename(title, company):
    def slug(s):
        return re.sub(r'[^A-Za-z0-9]', '_', s).strip('_')
    date  = datetime.now().strftime('%Y-%m-%d')
    title_slug   = '_'.join(slug(title).split('_')[:4])
    company_slug = '_'.join(slug(company).split('_')[:2])
    return f'Sunil_Kumar_{title_slug}_{company_slug}_{date}.docx'

def build_prompt(job, output_path, extra=''):
    return f"""You are building a tailored ATS resume for Sunil Kumar. Follow ALL rules exactly.

{PROFILE}

ROLE TO BUILD FOR:
Title: {job['title']}
Company: {job['company']}

JOB DESCRIPTION:
{job['jd_text'] or '(No JD provided - tailor using the role title and Fintech domain context)'}

{('ADDITIONAL INSTRUCTIONS:\n' + extra) if extra else ''}

TECHNICAL BUILD INSTRUCTIONS:
- Master template: {TEMPLATE}
- Output file: {output_path}
- Unpack the master template using zipfile (it is a .docx ZIP), modify word/document.xml, repack
- Never rebuild from a blank python-docx Document() - always start from the template
- Use python-docx or direct XML manipulation via zipfile + lxml
- Only import: os, re, copy, zipfile, shutil, lxml, docx (python-docx)

Write a complete self-contained Python script that builds this resume.
Return ONLY the Python script inside a ```python ... ``` code block. No explanation.
"""

def _extract_script(text):
    # Try standard ```python ... ``` block
    m = re.search(r'```python\s*(.*?)```', text, re.DOTALL)
    if m:
        return m.group(1).strip()
    # Try ``` ... ``` without language tag
    m = re.search(r'```\s*(import .*?)```', text, re.DOTALL)
    if m:
        return m.group(1).strip()
    # If response itself starts with import/# it's already raw Python
    stripped = text.strip()
    if stripped.startswith('import ') or stripped.startswith('#') or stripped.startswith('from '):
        return stripped
    # Last resort: find first 'import' and take everything from there
    idx = text.find('import os')
    if idx == -1:
        idx = text.find('import zipfile')
    if idx != -1:
        return text[idx:].strip()
    return None

def _do_build(job_id, db_path, extra=''):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    job = dict(conn.execute('SELECT * FROM jobs WHERE id=?', (job_id,)).fetchone())

    fname       = safe_filename(job['title'], job['company'])
    output_path = os.path.join(TMP_DIR, fname)

    def log(msg):
        conn.execute("UPDATE jobs SET build_log=? WHERE id=?", (msg, job_id))
        conn.commit()

    try:
        api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
        if not api_key:
            raise ValueError('ANTHROPIC_API_KEY not set in Render environment variables.')

        import anthropic
        import ast
        client = anthropic.Anthropic(api_key=api_key)

        script = None
        last_error = ''
        for attempt in range(1, 4):
            log(f'Calling Claude API to generate resume (attempt {attempt}/3)...')
            messages = [{'role': 'user', 'content': build_prompt(job, output_path, extra)}]
            if last_error:
                messages.append({'role': 'assistant', 'content': '```python\n' + (script or '') + '\n```'})
                messages.append({'role': 'user', 'content': f'That script had an error:\n{last_error}\n\nFix it and return the complete corrected Python script inside ```python ... ```.'})

            response = client.messages.create(
                model='claude-sonnet-4-6',
                max_tokens=8192,
                messages=messages
            )
            candidate = _extract_script(response.content[0].text)
            if not candidate:
                last_error = 'No Python script found in response.'
                continue
            try:
                ast.parse(candidate)
                script = candidate
                break
            except SyntaxError as e:
                last_error = f'SyntaxError: {e}'
                script = candidate

        if not script:
            raise ValueError(f'Claude failed to produce a valid script after 3 attempts. Last error: {last_error}')

        try:
            ast.parse(script)
        except SyntaxError as e:
            raise ValueError(f'Generated script has a syntax error after retries: {e}\n\nScript excerpt:\n{script[:500]}')

        log('Building resume...')
        with tempfile.NamedTemporaryFile(suffix='.py', mode='w', delete=False, encoding='utf-8') as f:
            f.write(script)
            script_path = f.name

        result = subprocess.run(
            ['python', script_path],
            capture_output=True, text=True, timeout=180, cwd=TMP_DIR
        )
        try:
            os.unlink(script_path)
        except Exception:
            pass

        if not os.path.exists(output_path):
            err = (result.stdout + '\n' + result.stderr).strip()
            raise ValueError('Build script ran but no .docx produced.\n\n' + err[-2000:])

        # Upload to Google Drive
        log('Uploading to Google Drive...')
        drive_link = ''
        download_url = ''
        drive_error = ''
        try:
            drive_link, download_url = upload_to_drive(output_path, fname, job=job)
            log(f'Drive upload OK: {drive_link}')
        except Exception as e:
            drive_error = str(e)
            log(f'Drive upload FAILED: {drive_error}')

        # Convert to PDF and upload to same Drive subfolder
        try:
            import shutil
            soffice = shutil.which('soffice') or '/usr/bin/soffice'
            if os.path.exists(soffice):
                pdf_out = os.path.splitext(output_path)[0] + '.pdf'
                subprocess.run(
                    [soffice, '--headless', '--convert-to', 'pdf',
                     '--outdir', os.path.dirname(output_path), output_path],
                    capture_output=True, timeout=60
                )
                if os.path.exists(pdf_out):
                    upload_to_drive(pdf_out, os.path.basename(pdf_out), job=job)
                    log('PDF uploaded to Drive.')
                    try: os.unlink(pdf_out)
                    except Exception: pass
        except Exception:
            pass

        # Also email if configured
        gmail_user = _get_setting_direct(db_path, 'gmail_user', os.environ.get('GMAIL_USER', ''))
        gmail_pass = _get_setting_direct(db_path, 'gmail_pass', os.environ.get('GMAIL_PASS', ''))
        to_addr    = _get_setting_direct(db_path, 'email_to', DEFAULT_TO)
        email_status = ''
        if gmail_user and gmail_pass:
            try:
                _send_email(job, output_path, to_addr, gmail_user, gmail_pass, drive_link)
                email_status = f'Email sent to {to_addr}.'
            except Exception as e:
                email_status = f'Email failed: {e}'

        try:
            os.unlink(output_path)
        except Exception:
            pass

        status  = 'done'
        _backup_async()  # persist to GitHub
        parts = []
        if drive_link:
            parts.append(f'Saved to Google Drive: {drive_link}')
        if download_url:
            parts.append(f'Download: {download_url}')
        elif drive_error:
            parts.append(f'Drive upload failed: {drive_error}')
        else:
            parts.append('Drive upload skipped (no credentials).')
        if email_status:
            parts.append(email_status)
        log_msg = '\n'.join(parts)

        conn.execute(
            "UPDATE jobs SET status=?, built_at=?, build_log=?, drive_link=? WHERE id=?",
            (status, datetime.now().strftime('%Y-%m-%d %H:%M'), log_msg, drive_link, job_id)
        )

    except subprocess.TimeoutExpired:
        conn.execute(
            "UPDATE jobs SET status='error', build_log='Build timed out after 3 minutes.' WHERE id=?",
            (job_id,)
        )
    except Exception as e:
        _backup_async()  # persist error state
        conn.execute(
            "UPDATE jobs SET status='error', build_log=? WHERE id=?",
            (str(e)[:2000], job_id)
        )
    finally:
        conn.commit()
        conn.close()

def _trigger_build(job_id, extra=''):
    t = threading.Thread(target=_do_build, args=(job_id, DB_PATH, extra), daemon=True)
    t.start()

# ── Email ─────────────────────────────────────────────────────────────────────

def _send_email(job, resume_path, to_addr, gmail_user, gmail_pass, drive_link=''):
    msg = MIMEMultipart()
    msg['From']    = gmail_user
    msg['To']      = to_addr
    msg['Subject'] = f"Resume: {job['title']} at {job['company']}"
    body = (f"Hi Sunil,\n\nYour resume for {job['title']} at {job['company']} is ready.\n"
            + (f"\nGoogle Drive: {drive_link}\n" if drive_link else '')
            + (f"\nJob link: {job['url']}\n" if job.get('url') else '')
            + "\nReady to apply!\n")
    msg.attach(MIMEText(body, 'plain'))
    fname = os.path.basename(resume_path)
    with open(resume_path, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{fname}"')
    msg.attach(part)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=20) as s:
        s.login(gmail_user, gmail_pass)
        s.send_message(msg)

# ── API routes ────────────────────────────────────────────────────────────────

@app.route('/api/jobs', methods=['GET'])
def api_list():
    jobs = get_db().execute('SELECT * FROM jobs ORDER BY id DESC').fetchall()
    return jsonify([dict(j) for j in jobs])

@app.route('/api/jobs', methods=['POST'])
def api_add():
    d       = request.json or {}
    title   = (d.get('title') or '').strip()
    company = (d.get('company') or '').strip()
    if not title or not company:
        return jsonify({'error': 'title and company required'}), 400
    db  = get_db()
    cur = db.execute('INSERT INTO jobs(title,company,url,jd_text) VALUES(?,?,?,?)',
                     (title, company, d.get('url', ''), d.get('jd', '')))
    db.commit()
    _backup_async()
    job_id = cur.lastrowid
    # Auto-trigger build immediately — no separate "Build" tap needed
    _trigger_build(job_id)
    row = db.execute('SELECT * FROM jobs WHERE id=?', (job_id,)).fetchone()
    return jsonify(dict(row)), 201

@app.route('/api/jobs/<int:jid>', methods=['PATCH'])
def api_edit(jid):
    d   = request.json or {}
    db  = get_db()
    row = db.execute('SELECT * FROM jobs WHERE id=?', (jid,)).fetchone()
    if not row:
        return jsonify({'error': 'not found'}), 404
    title   = d['title'].strip()   if 'title'   in d else row['title']
    company = d['company'].strip() if 'company' in d else row['company']
    jd      = d['jd'].strip()      if 'jd'      in d else row['jd_text']
    db.execute(
        "UPDATE jobs SET title=?, company=?, jd_text=?, status='pending', build_log='', built_at='', drive_link='' WHERE id=?",
        (title, company, jd, jid)
    )
    db.commit()
    return jsonify({'ok': True})

@app.route('/api/jobs/<int:jid>', methods=['DELETE'])
def api_delete(jid):
    get_db().execute('DELETE FROM jobs WHERE id=?', (jid,))
    get_db().commit()
    return jsonify({'ok': True})

@app.route('/api/jobs/<int:jid>/status')
def api_status(jid):
    row = get_db().execute('SELECT * FROM jobs WHERE id=?', (jid,)).fetchone()
    return jsonify(dict(row)) if row else ('', 404)

@app.route('/api/jobs/<int:jid>/build', methods=['POST'])
def api_build(jid):
    db  = get_db()
    row = db.execute('SELECT * FROM jobs WHERE id=?', (jid,)).fetchone()
    if not row:
        return jsonify({'error': 'not found'}), 404
    if row['status'] == 'building':
        return jsonify({'error': 'already building'})
    db.execute("UPDATE jobs SET status='building', build_log='Starting...', drive_link='' WHERE id=?", (jid,))
    db.commit()
    _trigger_build(jid)
    return jsonify({'ok': True})

@app.route('/api/jobs/<int:jid>/reply', methods=['POST'])
def api_reply(jid):
    db    = get_db()
    row   = db.execute('SELECT * FROM jobs WHERE id=?', (jid,)).fetchone()
    if not row:
        return jsonify({'error': 'not found'}), 404
    reply = (request.json or {}).get('reply', '').strip()
    if not reply:
        return jsonify({'error': 'reply text required'}), 400
    db.execute(
        "UPDATE jobs SET status='building', build_log='Retrying with your instructions...' WHERE id=?",
        (jid,)
    )
    db.commit()
    _trigger_build(jid, extra=reply)
    return jsonify({'ok': True})

@app.route('/api/build-all', methods=['POST'])
def api_build_all():
    db   = get_db()
    jobs = db.execute("SELECT * FROM jobs WHERE status='pending'").fetchall()
    for job in jobs:
        db.execute("UPDATE jobs SET status='building', build_log='Queued...' WHERE id=?", (job['id'],))
        db.commit()
        _trigger_build(job['id'])
        time.sleep(1)
    return jsonify({'started': len(jobs)})

@app.route('/api/jobs/clear-done', methods=['POST'])
def api_clear_done():
    cur = get_db().execute("DELETE FROM jobs WHERE status IN ('done','emailed','error')")
    get_db().commit()
    return jsonify({'ok': True, 'deleted': cur.rowcount})

@app.route('/api/settings', methods=['GET'])
def api_settings_get():
    return jsonify({
        'email_to':   get_setting('email_to', DEFAULT_TO),
        'gmail_user': get_setting('gmail_user', os.environ.get('GMAIL_USER', '')),
        'gmail_pass': ('•' * 8) if (get_setting('gmail_pass') or os.environ.get('GMAIL_PASS')) else '',
    })

@app.route('/api/settings', methods=['POST'])
def api_settings_save():
    d = request.json or {}
    for key in ('email_to', 'gmail_user'):
        if d.get(key):
            set_setting(key, d[key].strip())
    if d.get('gmail_pass') and not d['gmail_pass'].startswith('•'):
        set_setting('gmail_pass', d['gmail_pass'].strip())
    return jsonify({'ok': True})

def _dedup_key(title, company, url=''):
    if url:
        m = re.search(r'/jobs/view/(\d+)', url)
        if m:
            return f'li_{m.group(1)}'
    slug = re.sub(r'[^a-z0-9]', '', (title + company).lower())
    return f'tc_{slug[:60]}'

def _get_setting_cloud(key, default=''):
    try:
        row = get_db().execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
        return row[0] if row else os.environ.get(key.upper(), default)
    except Exception:
        return default

def parse_linkedin_alert_emails(gmail_user, gmail_pass, days=7):
    import imaplib, email as _email
    from datetime import timedelta

    jobs = []
    seen_keys = set()

    conn = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    conn.login(gmail_user, gmail_pass)
    conn.select('INBOX')

    since = (datetime.now() - timedelta(days=days)).strftime('%d-%b-%Y')
    _, msg_ids = conn.search(None,
        f'(FROM "jobalerts-noreply@linkedin.com" SINCE {since})')
    ids = msg_ids[0].split()

    for mid in ids:
        _, data = conn.fetch(mid, '(RFC822)')
        msg = _email.message_from_bytes(data[0][1])

        html_body = ''
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    html_body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    break
        elif msg.get_content_type() == 'text/html':
            html_body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        if not html_body:
            continue

        job_blocks = re.findall(
            r'href="(https://www\.linkedin\.com/comm/jobs/view/\d+[^"]*)"[^>]*>(.*?)</a>'
            r'.*?<span[^>]*>(.*?)</span>'
            r'(?:.*?<span[^>]*>(.*?)</span>)?',
            html_body, re.DOTALL | re.IGNORECASE
        )

        if not job_blocks:
            job_blocks_simple = re.findall(
                r'href="(https://www\.linkedin\.com/(?:comm/)?jobs/view/\d+[^"]*)"[^>]*>(.*?)</a>',
                html_body, re.DOTALL | re.IGNORECASE
            )
            for url, title_html in job_blocks_simple:
                title = re.sub(r'<[^>]+>', '', title_html).strip()
                if not title or len(title) < 4:
                    continue
                url_clean = url.split('?')[0]
                key = _dedup_key(title, '', url_clean)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                jobs.append({'title': title, 'company': '', 'location': '',
                             'job_url': url_clean, 'date_posted': '', 'dedup_key': key})
        else:
            for url, title_html, co_html, loc_html in job_blocks:
                title    = re.sub(r'<[^>]+>', '', title_html).strip()
                company  = re.sub(r'<[^>]+>', '', co_html).strip()
                location = re.sub(r'<[^>]+>', '', (loc_html or '')).strip()
                if not title or len(title) < 4:
                    continue
                url_clean = url.split('?')[0]
                key = _dedup_key(title, company, url_clean)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                jobs.append({'title': title, 'company': company, 'location': location,
                             'job_url': url_clean, 'date_posted': '', 'dedup_key': key})

    conn.logout()
    return jobs

@app.route('/api/import-linkedin-emails', methods=['POST'])
def api_import_linkedin_emails():
    gmail_user = _get_setting_cloud('gmail_user', DEFAULT_TO)
    gmail_pass = _get_setting_cloud('gmail_pass', os.environ.get('GMAIL_PASS', ''))
    if not gmail_pass:
        return jsonify({'error': 'Gmail app password not set in Settings'}), 400

    days = int(request.json.get('days', 7)) if request.json else 7
    try:
        parsed = parse_linkedin_alert_emails(gmail_user, gmail_pass, days=days)
    except Exception as e:
        return jsonify({'error': f'Email fetch failed: {e}'}), 500

    db = get_db()
    existing_keys = set(
        r[0] for r in db.execute('SELECT dedup_key FROM board_jobs').fetchall()
    )
    jobs_file = os.path.join(BASE, 'docs', 'jobs_data.json')
    if os.path.exists(jobs_file):
        try:
            with open(jobs_file, encoding='utf-8') as f:
                daily = json.load(f).get('jobs', [])
            for j in daily:
                existing_keys.add(_dedup_key(j.get('title',''), j.get('company',''), j.get('job_url','')))
        except Exception:
            pass

    inserted = 0
    for j in parsed:
        key = j.get('dedup_key') or _dedup_key(j['title'], j.get('company',''), j.get('job_url',''))
        if key in existing_keys:
            continue
        try:
            db.execute(
                'INSERT INTO board_jobs (title, company, location, job_url, date_posted, source, dedup_key) '
                'VALUES (?,?,?,?,?,?,?)',
                (j['title'], j.get('company',''), j.get('location',''),
                 j.get('job_url',''), j.get('date_posted',''), 'LinkedIn Alert', key)
            )
            existing_keys.add(key)
            inserted += 1
        except sqlite3.IntegrityError:
            pass
    db.commit()

    return jsonify({'ok': True, 'found': len(parsed), 'inserted': inserted,
                    'skipped': len(parsed) - inserted})

HISTORY_CUTOFF = '2026-08-01'

def _valid_date(d):
    if not d or str(d).lower() in ('nan', 'none', 'null', ''):
        return ''
    return d[:10] if d[:10] >= HISTORY_CUTOFF else None

@app.route('/api/jobboard')
def api_jobboard():
    jobs = []
    sources = []
    generated_at = ''

    # Cloud: reads from jobs_data.json (daily scrape + desktop sync)
    jobs_file = os.path.join(BASE, 'docs', 'jobs_data.json')
    if os.path.exists(jobs_file):
        try:
            with open(jobs_file, encoding='utf-8') as f:
                data = json.load(f)
            generated_at = data.get('generated_at', '')
            scrape_date = generated_at[:10] if generated_at else ''
            added = 0
            for j in data.get('jobs', []):
                d = _valid_date(j.get('date_posted', ''))
                if d is None:
                    continue
                if not d and scrape_date >= HISTORY_CUTOFF:
                    d = scrape_date
                elif not d:
                    continue
                jobs.append({
                    'title':       j.get('title', ''),
                    'company':     j.get('company', ''),
                    'location':    j.get('location', ''),
                    'source':      j.get('source', 'Daily Scrape'),
                    'date_posted': d,
                    'job_url':     j.get('job_url', ''),
                    'description': j.get('description', ''),
                    'score':       '',
                    'key_skills':  '',
                    'is_new':      1,
                    '_from':       'daily',
                })
                added += 1
            if added:
                sources.append(f'Jobs ({added})')
        except Exception:
            pass

    # Source 2: board_jobs table — LinkedIn alert email imports
    try:
        db = get_db()
        db.execute('''CREATE TABLE IF NOT EXISTS board_jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
            company TEXT DEFAULT '', location TEXT DEFAULT '',
            job_url TEXT DEFAULT '', date_posted TEXT DEFAULT '',
            source TEXT DEFAULT 'LinkedIn Alert', description TEXT DEFAULT '',
            dedup_key TEXT UNIQUE,
            imported_at TEXT DEFAULT (datetime('now','localtime')))''')
        alert_jobs = db.execute(
            'SELECT title, company, location, job_url, date_posted, source, imported_at '
            'FROM board_jobs ORDER BY imported_at DESC'
        ).fetchall()
        loaded_keys = set(
            _dedup_key(j['title'], j.get('company',''), j.get('job_url','')) for j in jobs
        )
        added = 0
        for r in alert_jobs:
            key = _dedup_key(r[0], r[1], r[3])
            if key in loaded_keys:
                continue
            loaded_keys.add(key)
            jobs.append({
                'title': r[0], 'company': r[1], 'location': r[2],
                'job_url': r[3], 'date_posted': r[4],
                'source': r[5] or 'LinkedIn Alert',
                'description': '', 'score': '', 'key_skills': '', 'is_new': 1,
                '_from': 'email_alert',
            })
            added += 1
        if added:
            sources.append(f'LinkedIn Alerts ({added} jobs)')
    except Exception:
        pass

    return jsonify({'jobs': jobs, 'total': len(jobs), 'sources': sources, 'generated_at': generated_at})

@app.route('/health')
def health():
    token = _gh_token()
    env_keys = [k for k in os.environ if 'TOKEN' in k.upper() or 'HUB' in k.upper() or 'GOOGLE' in k.upper()]
    drive_ok = bool(
        os.environ.get('GOOGLE_REFRESH_TOKEN') and
        os.environ.get('GOOGLE_CLIENT_ID') and
        os.environ.get('GOOGLE_CLIENT_SECRET')
    )
    missing = [v for v in ['GOOGLE_REFRESH_TOKEN','GOOGLE_CLIENT_ID','GOOGLE_CLIENT_SECRET']
               if not os.environ.get(v)]
    return jsonify({
        'ok':            True,
        'template':      os.path.exists(TEMPLATE),
        'api_key':       bool(os.environ.get('ANTHROPIC_API_KEY')),
        'drive':         drive_ok,
        'drive_missing': missing,
        'gh_token':      bool(token),
        'gh_token_hint': (token[:6] + '…') if token else 'NOT SET',
        'env_google_keys': [k for k in os.environ if 'GOOGLE' in k.upper()],
    })

@app.route('/api/backup-now', methods=['POST'])
def api_backup_now():
    """Trigger a synchronous GitHub backup and return result."""
    token = _gh_token()
    if not token:
        return jsonify({'ok': False, 'error': 'RESUME_HUB_TOKEN not set — add it as a Render env var'}), 400
    import urllib.request as _ur
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        jobs = [dict(r) for r in conn.execute('SELECT * FROM jobs ORDER BY id').fetchall()]
        conn.close()

        payload = json.dumps({'jobs': jobs, 'backed_up_at': datetime.now().isoformat()}, ensure_ascii=False)
        content_b64 = base64.b64encode(payload.encode('utf-8')).decode()

        sha = ''
        try:
            req = _ur.Request(GITHUB_API_URL,
                headers={'Authorization': f'token {token}', 'Accept': 'application/vnd.github.v3+json'})
            with _ur.urlopen(req, timeout=8) as resp:
                sha = json.loads(resp.read()).get('sha', '')
        except Exception:
            pass

        body = {'message': f'manual backup: {len(jobs)} jobs', 'content': content_b64, 'branch': 'main'}
        if sha:
            body['sha'] = sha

        data = json.dumps(body).encode()
        req = _ur.Request(GITHUB_API_URL, data=data, method='PUT',
            headers={'Authorization': f'token {token}',
                     'Content-Type': 'application/json',
                     'Accept': 'application/vnd.github.v3+json'})
        _ur.urlopen(req, timeout=15)
        return jsonify({'ok': True, 'jobs_backed_up': len(jobs)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# ── UI ────────────────────────────────────────────────────────────────────────

@app.route('/')
def ui():
    return HTML

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1"/>
<meta name="theme-color" content="#ffffff"/>
<title>Resume Hub</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;color:#111;font-size:15px;min-height:100vh}
.topbar{background:#fff;border-bottom:1px solid #e5e5e5;padding:14px 16px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:10}
.topbar h1{font-size:17px;font-weight:600;flex:1}
.badge-count{background:#e8f0fe;color:#1a56db;font-size:12px;padding:3px 9px;border-radius:20px;font-weight:500}
.tabs{display:flex;background:#fff;border-bottom:1px solid #e5e5e5;position:sticky;top:53px;z-index:9}
.tab{flex:1;padding:12px 0;text-align:center;font-size:13px;color:#666;cursor:pointer;border-bottom:2.5px solid transparent;user-select:none}
.tab.active{color:#111;border-bottom:2.5px solid #111;font-weight:500}
.pane{display:none;padding:14px;padding-bottom:40px}
.pane.active{display:block}
.card{background:#fff;border:1px solid #e5e5e5;border-radius:12px;padding:16px;margin-bottom:12px}
label{display:block;font-size:12px;color:#666;margin:14px 0 5px;font-weight:500}
label:first-child{margin-top:0}
input,textarea{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;font-family:inherit;background:#fafafa;color:#111;outline:none;-webkit-appearance:none}
input:focus,textarea:focus{border-color:#888;background:#fff}
textarea{height:120px;resize:vertical}
.btn{display:block;width:100%;padding:12px;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;border:none;margin-top:10px;text-align:center}
.btn-primary{background:#111;color:#fff}
.bulk-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:12px}
.bulk-btn{padding:10px 6px;background:#fff;border:1px solid #ddd;border-radius:8px;font-size:13px;cursor:pointer;text-align:center;font-weight:500}
.job-card{background:#fff;border:1px solid #e5e5e5;border-radius:12px;padding:14px;margin-bottom:10px}
.job-head{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
.job-title{font-size:14px;font-weight:600;line-height:1.3}
.job-co{font-size:12px;color:#666;margin-top:2px}
.badge{font-size:11px;font-weight:500;padding:3px 9px;border-radius:20px;white-space:nowrap;flex-shrink:0}
.b-pending{background:#fff7e6;color:#b45309}
.b-building{background:#eff6ff;color:#1d4ed8}
.b-done{background:#f0fdf4;color:#16a34a}
.b-error{background:#fef2f2;color:#b91c1c}
.job-meta{font-size:11px;color:#999;margin-top:6px}
.drive-link{display:block;margin-top:8px;padding:9px 12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;font-size:13px;color:#16a34a;font-weight:500;text-decoration:none;text-align:center}
.job-actions{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:7px;margin-top:10px}
.ja-btn{padding:8px 6px;font-size:12px;font-weight:500;border:1px solid #ddd;border-radius:7px;background:#fff;color:#111;cursor:pointer;text-align:center}
.ja-btn:disabled{opacity:.35;cursor:default}
.ja-del{border-color:#fecaca;color:#dc2626}
.log-box{background:#111;color:#9effa0;font-size:11px;font-family:monospace;padding:10px;border-radius:7px;max-height:120px;overflow-y:auto;margin-top:8px;white-space:pre-wrap;word-break:break-all}
.empty{text-align:center;color:#999;padding:40px 0;font-size:14px}
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#111;color:#fff;padding:10px 20px;border-radius:20px;font-size:13px;z-index:100;opacity:0;transition:opacity .2s;pointer-events:none}
.toast.show{opacity:1}
</style>
</head>
<body>
<div class="topbar">
  <span>&#128221;</span>
  <h1>Resume Hub</h1>
  <span class="badge-count" id="pending-badge">0 pending</span>
</div>
<div class="tabs">
  <div class="tab active" onclick="showTab('add')">Add Job</div>
  <div class="tab" onclick="showTab('board')">Job Board</div>
  <div class="tab" onclick="showTab('queue')">Queue</div>
  <div class="tab" onclick="showTab('settings')">Settings</div>
</div>

<div class="pane active" id="pane-add">
  <div class="card">
    <label>Job Title *</label>
    <input type="text" id="add-title" placeholder="VP of Delivery"/>
    <label>Company *</label>
    <input type="text" id="add-company" placeholder="Acme Fintech"/>
    <label>Job URL (optional)</label>
    <input type="url" id="add-url" placeholder="https://linkedin.com/jobs/view/..."/>
    <label>Job Description (paste full JD for best results)</label>
    <textarea id="add-jd" placeholder="Paste the full job description here..."></textarea>
    <button class="btn btn-primary" onclick="addAndBuild()">&#9654; Submit &amp; Build Resume</button>
  </div>
</div>

<div class="pane" id="pane-board">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
    <span style="font-size:12px;color:#999" id="board-meta">Loading…</span>
    <div style="display:flex;gap:6px">
      <button onclick="importLinkedInAlerts()" id="btn-import-li" style="padding:6px 12px;border:1px solid #0a66c2;border-radius:8px;background:#fff;color:#0a66c2;font-size:12px;font-weight:500;cursor:pointer">&#128231; LinkedIn Alerts</button>
      <button onclick="loadBoard()" style="padding:6px 12px;border:1px solid #ddd;border-radius:8px;background:#fff;font-size:12px;cursor:pointer">&#8635; Refresh</button>
    </div>
  </div>
  <div id="import-li-status" style="display:none;font-size:12px;color:#555;margin-bottom:10px;padding:8px 12px;background:#f0f7ff;border-radius:8px;border:1px solid #cce0ff"></div>
  <div id="board-filter" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap"></div>
  <div id="board-list"><div class="empty">Loading jobs…</div></div>
</div>

<div class="pane" id="pane-queue">
  <div class="bulk-row">
    <button class="bulk-btn" onclick="buildAll()">&#9654; Build all pending</button>
    <button class="bulk-btn" onclick="loadJobs()">&#8635; Refresh</button>
    <button class="bulk-btn" onclick="clearDone()" style="color:#dc2626;border-color:#fecaca">&#128465; Clear done</button>
  </div>
  <div id="job-list"><div class="empty">No jobs yet. Add one on the Add Job tab.</div></div>
</div>

<div class="pane" id="pane-settings">
  <div class="card">
    <p style="font-size:12px;color:#666;margin-bottom:14px">Resumes are saved to your Google Drive automatically. Gmail is optional — enables email delivery too.</p>
    <label>Gmail address (optional sender)</label>
    <input type="email" id="s-gmail-user" placeholder="you@gmail.com"/>
    <label>Gmail App Password</label>
    <input type="password" id="s-gmail-pass" placeholder="xxxx xxxx xxxx xxxx"/>
    <label>Send resumes to</label>
    <input type="email" id="s-email-to" placeholder="ksunilmech29@gmail.com"/>
    <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
  </div>
</div>

<div class="toast" id="toast"></div>

<div id="edit-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:50;padding:20px;align-items:flex-end">
  <div style="background:#fff;border-radius:16px 16px 0 0;padding:20px;width:100%;max-height:90vh;overflow-y:auto;margin-top:auto">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
      <span style="font-weight:600;font-size:15px">Edit job</span>
      <button onclick="closeEdit()" style="background:none;border:none;font-size:22px;cursor:pointer;color:#666">&times;</button>
    </div>
    <input type="hidden" id="edit-id"/>
    <label style="display:block;font-size:12px;color:#666;margin-bottom:4px;font-weight:500">Job Title</label>
    <input type="text" id="edit-title" style="width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:12px"/>
    <label style="display:block;font-size:12px;color:#666;margin-bottom:4px;font-weight:500">Company</label>
    <input type="text" id="edit-company" style="width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:12px"/>
    <label style="display:block;font-size:12px;color:#666;margin-bottom:4px;font-weight:500">Job Description</label>
    <textarea id="edit-jd" style="width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;height:160px;resize:vertical"></textarea>
    <button onclick="saveEdit()" style="display:block;width:100%;padding:12px;background:#111;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;margin-top:12px">Save &amp; reset to pending</button>
  </div>
</div>

<script>
let jobs = [];

function showTab(name) {
  document.querySelectorAll('.pane').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('pane-' + name).classList.add('active');
  const tabs = ['add','board','queue','settings'];
  document.querySelectorAll('.tab')[tabs.indexOf(name)].classList.add('active');
  if (name === 'queue') loadJobs();
  if (name === 'settings') loadSettings();
  if (name === 'board') loadBoard();
}

// ── Job Board ──────────────────────────────────────────────────────────────
let boardJobs = [];
let boardFilter = '';
let boardSearch = '';

// ── Board preferences stored in localStorage ───────────────────
const DISMISSED_KEY = 'board_dismissed_v1';
const FIT_KEY       = 'board_fit_v1';
const STOP_WORDS = new Set(['and','the','of','for','in','at','a','an','to','with','head','vice',
  'president','senior','junior','global','chief','officer','lead','principal','associate']);

function getDismissed() {
  try { return new Set(JSON.parse(localStorage.getItem(DISMISSED_KEY)||'[]')); } catch { return new Set(); }
}
function saveDismissed(s) { localStorage.setItem(DISMISSED_KEY, JSON.stringify([...s])); }
function getFitSignals() {
  try { return JSON.parse(localStorage.getItem(FIT_KEY)||'[]'); } catch { return []; }
}
function saveFitSignals(arr) { localStorage.setItem(FIT_KEY, JSON.stringify(arr.slice(-60))); }

function extractWords(title) {
  return (title||'').toLowerCase().split(/\W+/).filter(w => w.length > 3 && !STOP_WORDS.has(w));
}

function fitScore(job, signals) {
  if (!signals.length) return 0;
  const words = new Set(extractWords(job.title));
  let score = 0;
  for (const sig of signals) {
    for (const w of (sig.words||[])) { if (words.has(w)) score += 2; }
    if (sig.source && job.source && job.source.toLowerCase().includes(sig.source.toLowerCase())) score += 0.5;
  }
  return score;
}

function dismissJob(idx) {
  const j = boardJobs[idx];
  if (!j) return;
  const key = j.job_url || (j.title + '|' + j.company);
  const d = getDismissed();
  d.add(key);
  saveDismissed(d);
  boardJobs.splice(idx, 1);
  renderBoard();
  toast('Removed — won\'t show again');
}

async function loadBoard() {
  document.getElementById('board-list').innerHTML = '<div class="empty">Loading…</div>';
  const d = await api('/api/jobboard');
  if (d.error && !d.jobs) { document.getElementById('board-list').innerHTML = `<div class="empty">${esc(d.error)}</div>`; return; }
  // Filter already-dismissed jobs
  const dismissed = getDismissed();
  boardJobs = (d.jobs || []).filter(j => {
    const key = j.job_url || (j.title + '|' + j.company);
    return !dismissed.has(key);
  });
  const srcLabel = (d.sources||[]).join(' + ') || `${boardJobs.length} jobs`;
  document.getElementById('board-meta').textContent = `${srcLabel}${d.generated_at ? ' · ' + d.generated_at : ''}`;
  renderBoardFilter();
  renderBoard();
}

function renderBoardFilter() {
  const el = document.getElementById('board-filter');
  const chip = (val, label, active) =>
    `<button onclick="setBoardFilter('${val}')" style="padding:5px 12px;border-radius:20px;border:1px solid ${active?'#1a3557':'#ddd'};background:${active?'#1a3557':'#fff'};color:${active?'#fff':'#444'};font-size:12px;cursor:pointer">${label}</button>`;
  const hasLinkedIn = boardJobs.some(j => (j.source||'').toLowerCase().includes('linkedin'));
  const hasIndeed   = boardJobs.some(j => (j.source||'').toLowerCase().includes('indeed'));
  el.innerHTML =
    `<input type="text" placeholder="Search title / company…" oninput="setBoardSearch(this.value)"
      style="width:100%;padding:8px 12px;border:1px solid #ddd;border-radius:8px;font-size:13px;margin-bottom:8px;font-family:inherit"/>` +
    chip('', 'All', boardFilter==='') +
    chip('linkedin', '🔵 LinkedIn', boardFilter==='linkedin') +
    (hasIndeed ? chip('indeed', 'Indeed', boardFilter==='indeed') : '');
}

function setBoardFilter(src) {
  boardFilter = src;
  renderBoardFilter();
  renderBoard();
}

function setBoardSearch(val) {
  boardSearch = val.toLowerCase();
  renderBoard();
}

function renderBoard() {
  const el = document.getElementById('board-list');
  let filtered = boardFilter
    ? boardJobs.filter(j => (j.source||'').toLowerCase().includes(boardFilter))
    : boardJobs;
  if (boardSearch) filtered = filtered.filter(j =>
    (j.title||'').toLowerCase().includes(boardSearch) ||
    (j.company||'').toLowerCase().includes(boardSearch)
  );
  if (!filtered.length) { el.innerHTML = '<div class="empty">No jobs found.</div>'; return; }

  // Score and sort — highest fit first, then latest date, then original order
  const signals = getFitSignals();
  const scored = filtered.map((j, i) => ({ j, i, fs: fitScore(j, signals) }));
  const dateKey = j => (j.date_posted && j.date_posted !== 'nan' && j.date_posted !== '') ? j.date_posted : '0000-00-00';
  scored.sort((a, b) => b.fs - a.fs || dateKey(b.j).localeCompare(dateKey(a.j)) || a.i - b.i);

  el.innerHTML = scored.slice(0, 150).map(({ j, fs }) => {
    const idx = boardJobs.indexOf(j);
    const posted = (j.date_posted && j.date_posted !== 'nan') ? j.date_posted : '';
    const scoreHtml = j.score ? `<span style="background:#f0fdf4;color:#16a34a;font-size:11px;padding:2px 8px;border-radius:20px;margin-left:6px">Score ${j.score}</span>` : '';
    const newBadge  = j.is_new  ? `<span style="background:#fef9c3;color:#854d0e;font-size:11px;padding:2px 8px;border-radius:20px;margin-left:4px">New</span>` : '';
    const fitBadge  = fs > 0    ? `<span style="background:#ede9fe;color:#5b21b6;font-size:11px;padding:2px 8px;border-radius:20px;margin-left:4px">&#127919; Recommended</span>` : '';
    const skills = j.key_skills ? `<div style="font-size:11px;color:#666;margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(j.key_skills.slice(0,120))}</div>` : '';
    return `<div class="job-card" id="bcard-${idx}">
  <div class="job-head">
    <div style="flex:1;min-width:0">
      <div class="job-title">${esc(j.title)}${newBadge}${scoreHtml}${fitBadge}</div>
      <div class="job-co">${esc(j.company)}${j.location ? ' · ' + esc(j.location) : ''}</div>
    </div>
    <span class="badge" style="background:#f0f4ff;color:#1a3557;flex-shrink:0">${esc(j.source||'')}</span>
  </div>
  <div class="job-meta">${posted}</div>
  ${skills}
  <div style="display:flex;gap:8px;margin-top:10px">
    ${j.job_url ? `<a href="${esc(j.job_url)}" target="_blank" style="flex:1;padding:8px 0;text-align:center;border:1px solid #ddd;border-radius:7px;font-size:12px;color:#1a56db;text-decoration:none">&#128279; View</a>` : '<span style="flex:1"></span>'}
    <button onclick="addBoardJobToQueue(${idx})" id="badd-${idx}" style="flex:2;padding:8px 0;background:#111;color:#fff;border:none;border-radius:7px;font-size:12px;font-weight:500;cursor:pointer">+ Add to Queue</button>
    <button onclick="dismissJob(${idx})" title="Not for me" style="padding:8px 10px;background:#fff;color:#999;border:1px solid #eee;border-radius:7px;font-size:14px;cursor:pointer">&#10005;</button>
  </div>
</div>`;
  }).join('') + (scored.length > 150 ? `<div class="empty" style="padding:12px">Showing 150 of ${scored.length} — use search to narrow down</div>` : '');
}

async function addBoardJobToQueue(idx) {
  const j = boardJobs[idx];
  if (!j) return;
  const btn = document.getElementById('badd-' + idx);
  if (btn) { btn.disabled = true; btn.textContent = 'Adding…'; }
  const d = await api('/api/jobs', 'POST', {
    title:   j.title,
    company: j.company,
    url:     j.job_url || '',
    jd:      j.description || '',
  });
  if (d.id) {
    if (btn) { btn.textContent = '✓ Added'; btn.style.background = '#16a34a'; }
    toast(`"${j.title}" added to queue`);
    // Learn from this job — save fit signals
    const words = extractWords(j.title);
    if (words.length) {
      const signals = getFitSignals();
      signals.push({ words, source: j.source || '' });
      saveFitSignals(signals);
    }
  } else {
    if (btn) { btn.disabled = false; btn.textContent = '+ Add to Queue'; }
    toast('Error: ' + (d.error || 'unknown'), 3000);
  }
}

async function importLinkedInAlerts() {
  const btn = document.getElementById('btn-import-li');
  const status = document.getElementById('import-li-status');
  btn.disabled = true;
  btn.textContent = 'Importing…';
  status.style.display = 'none';
  try {
    const r = await api('/api/import-linkedin-emails', 'POST', { days: 7 });
    if (r.error) {
      status.textContent = 'Error: ' + r.error;
      status.style.display = 'block';
    } else {
      status.textContent = `LinkedIn Alerts: ${r.found} found, ${r.inserted} new jobs added, ${r.skipped} duplicates skipped.`;
      status.style.display = 'block';
      if (r.inserted > 0) loadBoard();
    }
  } catch(e) {
    status.textContent = 'Import failed: ' + e;
    status.style.display = 'block';
  }
  btn.disabled = false;
  btn.textContent = '✉ LinkedIn Alerts';
}

function toast(msg, ms=2500) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), ms);
}

async function api(path, method='GET', body=null) {
  const opts = { method, headers: {'Content-Type':'application/json'} };
  if (body) opts.body = JSON.stringify(body);
  try {
    const r = await fetch(path, opts);
    if (!r.ok && r.status !== 201) {
      const err = await r.json().catch(() => ({}));
      return { error: err.error || 'Server error ' + r.status };
    }
    return await r.json();
  } catch(e) {
    toast('Cannot reach server', 4000);
    return { error: e.message };
  }
}

function getFormData() {
  return {
    title:   document.getElementById('add-title').value.trim(),
    company: document.getElementById('add-company').value.trim(),
    url:     document.getElementById('add-url').value.trim(),
    jd:      document.getElementById('add-jd').value.trim(),
  };
}

function clearForm() {
  ['add-title','add-company','add-url','add-jd'].forEach(id =>
    document.getElementById(id).value = '');
}

async function addAndBuild() {
  const { title, company, url, jd } = getFormData();
  if (!title || !company) { toast('Title and company are required'); return; }
  if (!jd.trim()) {
    toast('Tip: paste the job description for a tailored resume', 3000);
  }
  const d = await api('/api/jobs', 'POST', { title, company, url, jd });
  if (!d.id) { toast('Error: ' + (d.error||''), 4000); return; }
  clearForm();
  // Build starts automatically on the server — no separate trigger needed
  showTab('queue');
  toast('Submitted! Building resume — check Queue in 2-3 mins');
}

async function addOnly() {
  const { title, company, url, jd } = getFormData();
  if (!title || !company) { toast('Title and company are required'); return; }
  const d = await api('/api/jobs', 'POST', { title, company, url, jd });
  if (d.id) { clearForm(); showTab('queue'); toast('Added to queue!'); }
  else toast('Error: ' + (d.error||''), 4000);
}

async function loadJobs() {
  const result = await api('/api/jobs');
  jobs = Array.isArray(result) ? result : [];
  renderJobs();
  updateBadge();
}

function renderJobs() {
  const el = document.getElementById('job-list');
  if (!jobs.length) { el.innerHTML = '<div class="empty">No jobs yet. Add one on the Add Job tab.</div>'; return; }
  el.innerHTML = jobs.map(j => {
    const bClass = {pending:'b-pending',building:'b-building',done:'b-done',error:'b-error'}[j.status] || 'b-pending';
    const bLabel = {pending:'Pending',building:'Building...',done:'Done',error:'Error'}[j.status] || j.status;
    const canBuild    = j.status === 'pending' || j.status === 'error';
    const canRecreate = j.status === 'done';
    const showLog     = j.build_log && (j.status === 'building' || j.status === 'error' || (j.status === 'done' && !j.drive_link));
    const logHtml     = showLog ? `<div class="log-box">${esc(j.build_log.slice(-600))}</div>` : '';
    const driveHtml   = j.drive_link ? `<a class="drive-link" href="${esc(j.drive_link)}" target="_blank">&#128194; Open in Google Drive</a>` : '';
    const replyHtml   = j.status === 'error' ? `
      <div style="margin-top:8px">
        <textarea id="reply-${j.id}" placeholder="Instructions e.g. reframe domain as enterprise SaaS..."
          style="width:100%;padding:9px;border:1px solid #ddd;border-radius:8px;font-size:13px;height:70px;resize:none;font-family:inherit"></textarea>
        <button onclick="replyBuild(${j.id})" style="width:100%;margin-top:6px;padding:9px;background:#111;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:500;cursor:pointer">
          &#9654; Retry with instructions
        </button>
      </div>` : '';
    const recreateHtml = canRecreate ? `
      <div id="recreate-box-${j.id}" style="display:none;margin-top:10px">
        <textarea id="recreate-note-${j.id}" placeholder="Reason for recreating, e.g: better match the strategy focus, sharpen the CS angle, fix page 2 gaps…"
          style="width:100%;padding:9px;border:1px solid #ddd;border-radius:8px;font-size:13px;height:70px;resize:none;font-family:inherit"></textarea>
        <button onclick="submitRecreate(${j.id})" style="width:100%;margin-top:6px;padding:9px;background:#1a3557;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:500;cursor:pointer">
          &#9654; Recreate with these instructions
        </button>
        <button onclick="document.getElementById('recreate-box-${j.id}').style.display='none'" style="width:100%;margin-top:5px;padding:8px;background:#f0f0f0;color:#444;border:none;border-radius:8px;font-size:12px;cursor:pointer">Cancel</button>
      </div>` : '';
    return `<div class="job-card" id="card-${j.id}">
  <div class="job-head">
    <div><div class="job-title">${esc(j.title)}</div><div class="job-co">${esc(j.company)}</div></div>
    <span class="badge ${bClass}">${bLabel}</span>
  </div>
  <div class="job-meta">${j.created_at.slice(0,16).replace('T',' ')}${j.built_at ? ' &middot; built ' + j.built_at : ''}</div>
  ${driveHtml}${logHtml}${replyHtml}${recreateHtml}
  <div class="job-actions">
    <button class="ja-btn" onclick="buildOne(${j.id})" ${canBuild?'':'disabled'}>&#9654; Build</button>
    <button class="ja-btn" onclick="toggleRecreate(${j.id})" ${canRecreate?'':'disabled'} style="${canRecreate?'border-color:#1a3557;color:#1a3557':''}">&#128260; Recreate</button>
    <button class="ja-btn" onclick="editJob(${j.id})">&#9998; Edit</button>
    <button class="ja-btn ja-del" onclick="deleteJob(${j.id})">&#128465; Delete</button>
  </div>
</div>`;
  }).join('');

  if (jobs.some(j => j.status === 'building')) setTimeout(loadJobs, 4000);
}

function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

async function buildOne(id) {
  const d = await api(`/api/jobs/${id}/build`, 'POST');
  if (d.ok) { toast('Build started — 2-3 mins'); setTimeout(loadJobs, 1500); }
  else toast('Error: ' + (d.error||''), 4000);
}

function toggleRecreate(id) {
  const box = document.getElementById('recreate-box-' + id);
  if (!box) return;
  box.style.display = box.style.display === 'none' ? 'block' : 'none';
}

async function submitRecreate(id) {
  const note = (document.getElementById('recreate-note-' + id) || {}).value || '';
  const instructions = note.trim() ? `Recreate this resume. Reason: ${note}` : 'Recreate this resume with improved JD alignment.';
  const d = await api(`/api/jobs/${id}/reply`, 'POST', {reply: instructions});
  if (d.ok) {
    document.getElementById('recreate-box-' + id).style.display = 'none';
    toast('Recreating resume…');
    setTimeout(loadJobs, 1500);
  } else {
    toast('Error: ' + (d.error || 'unknown'), 4000);
  }
}

async function replyBuild(id) {
  const reply = (document.getElementById('reply-' + id)||{}).value || '';
  const d = await api(`/api/jobs/${id}/reply`, 'POST', { reply });
  if (d.ok) { toast('Retrying...'); setTimeout(loadJobs, 1500); }
  else toast('Error: ' + (d.error||''), 4000);
}

async function deleteJob(id) {
  if (!confirm('Delete this job?')) return;
  await api(`/api/jobs/${id}`, 'DELETE');
  jobs = jobs.filter(j => j.id !== id);
  renderJobs(); updateBadge();
}

async function clearDone() {
  const n = jobs.filter(j => ['done','error'].includes(j.status)).length;
  if (!n) { toast('No completed jobs to clear'); return; }
  if (!confirm(`Delete ${n} completed job(s)?`)) return;
  const d = await api('/api/jobs/clear-done', 'POST');
  if (d.ok) { toast(`Cleared ${d.deleted} job(s)`); await loadJobs(); }
}

async function buildAll() {
  const pending = jobs.filter(j => j.status === 'pending').length;
  if (!pending) { toast('No pending jobs'); return; }
  const d = await api('/api/build-all', 'POST');
  toast(`Building ${d.started} resume(s)...`);
  setTimeout(loadJobs, 2000);
}

function updateBadge() {
  const n = jobs.filter(j => j.status === 'pending').length;
  document.getElementById('pending-badge').textContent = n + ' pending';
}

async function loadSettings() {
  const d = await api('/api/settings');
  document.getElementById('s-gmail-user').value = d.gmail_user || '';
  document.getElementById('s-gmail-pass').value = d.gmail_pass || '';
  document.getElementById('s-email-to').value   = d.email_to  || '';
}

async function saveSettings() {
  const d = await api('/api/settings', 'POST', {
    gmail_user: document.getElementById('s-gmail-user').value,
    gmail_pass: document.getElementById('s-gmail-pass').value,
    email_to:   document.getElementById('s-email-to').value,
  });
  toast(d.ok ? 'Settings saved!' : 'Error saving');
}

function editJob(id) {
  const j = jobs.find(x => x.id === id);
  if (!j) return;
  document.getElementById('edit-id').value      = id;
  document.getElementById('edit-title').value   = j.title;
  document.getElementById('edit-company').value = j.company;
  document.getElementById('edit-jd').value      = j.jd_text || '';
  document.getElementById('edit-modal').style.display = 'flex';
}

function closeEdit() {
  document.getElementById('edit-modal').style.display = 'none';
}

async function saveEdit() {
  const id      = document.getElementById('edit-id').value;
  const title   = document.getElementById('edit-title').value.trim();
  const company = document.getElementById('edit-company').value.trim();
  const jd      = document.getElementById('edit-jd').value.trim();
  if (!title || !company) { toast('Title and company required'); return; }
  const d = await api(`/api/jobs/${id}`, 'PATCH', { title, company, jd });
  if (d.ok) { closeEdit(); await loadJobs(); toast('Saved'); }
  else toast('Error: ' + (d.error||''));
}

setInterval(() => {
  if (document.getElementById('pane-queue').classList.contains('active')) {
    if (jobs.some(j => j.status === 'building')) loadJobs();
  }
}, 5000);

updateBadge();
</script>
</body>
</html>
"""

# ── Start ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5002))
    print('=' * 50)
    print('  Resume Hub Cloud')
    print(f'  Port:     {port}')
    print(f'  Template: {"FOUND" if os.path.exists(TEMPLATE) else "MISSING"}')
    print(f'  API Key:  {"SET" if os.environ.get("ANTHROPIC_API_KEY") else "NOT SET"}')
    print(f'  Drive:    {"SET" if os.environ.get("GOOGLE_CREDENTIALS_JSON") else "NOT SET"}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
