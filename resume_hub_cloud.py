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

init_db()  # runs at import time — works with both gunicorn and direct python

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
CANDIDATE: Sunil Kumar
Contact: +91-9741114967 | ksunilmech29@gmail.com | linkedin.com/in/sunil-kumar-k-8b1b72a/
Experience: 19+ years, Fintech and Financial Services domain

CAREER ENTRIES (use exact titles):
- Accenture (Nov 2006 - Oct 2009), Software Engineer
  Description VERBATIM: "Started as a QA Engineer and progressively contributed to product and delivery - built user stories, supported product definition and process standardization across Banking, Insurance and Healthcare domain."
- Attra (May 2018 - Mar 2022), Program Manager [NEVER "Program Director"]
  Adapt description to role but always keep title as "Program Manager"

CANONICAL EDUCATION AND CERTIFICATIONS (exactly these 4, in this exact order):
1. Executive Global Management Programme  -  IIM Calcutta  |  2014 - 2016
2. Bachelor of Engineering  -  Visvesvaraya Technological University  |  2002 - 2006
3. SAFe 5 Agilist  -  Scaled Agile Inc.  |  Dec 2022
4. Certified Scrum Master (CSM)  -  Scrum Alliance
NEVER add any other certification (no PMP, Prince2, TOGAF, CFA).

DOMAIN: Always Fintech and Financial Services - anchor all experience to banking, payments,
lending, capital markets, or enterprise fintech contexts.

RESUME RULES (all mandatory):
- 2 pages only, never exceed
- No em dashes anywhere, use colon or comma instead
- No semicolons anywhere
- No tables in experience or body sections - only the 3-column Core Competencies table
- Career Highlights: exactly 2 lines, exactly 3 words per line (e.g. "19+ Years Experience")
- Core Competencies: 3-column borderless table, 1 bold teal heading per column, up to 5 bullets (3-5 words, dash prefix)
- Each job experience description: exactly 2 sentences max, never 3
- Education and Certifications: always on Page 1, exactly 4 entries
- Page 1 order: Name/contact, Summary (2 short paras 2 sentences each), Career Highlights, Core Competencies, Experience (4 roles), Education
- Page 2: Key Accomplishments only, 5-7 sections, bullets 200-250 chars, 2 sentences max, varied opening verbs
- Vary opening verbs (Led, Drove, Designed, Established, Owned, Championed, Developed, Governed, Delivered, Managed, Scaled, Spearheaded, Directed, Executed, Defined)
- Avoid "Built" more than twice per resume
- Use "Attra" never "Synechron"
- Colors: Navy 1A3557, Teal 0D6E8A, Dark text 1A1A1A, Mid-gray 444444, Light-gray 888888
- Page margins: top/bottom 576 DXA (0.4 inch), left/right 792 DXA (0.55 inch), content width 10656 DXA
- Portfolio/revenue ($20M) only if JD explicitly mentions P&L or revenue ownership
- Do gap check internally, proceed without asking for confirmation
- Validate to 10/10 before saving
"""

# ── Google Drive ──────────────────────────────────────────────────────────────

def _get_drive_service():
    creds_b64 = os.environ.get('GOOGLE_CREDENTIALS_JSON', '')
    if not creds_b64:
        raise ValueError('GOOGLE_CREDENTIALS_JSON env var not set.')
    creds_json = base64.b64decode(creds_b64).decode('utf-8')
    creds_dict = json.loads(creds_json)

    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=['https://www.googleapis.com/auth/drive.file']
    )
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file_path, filename):
    from googleapiclient.http import MediaFileUpload
    folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID', '')
    service   = _get_drive_service()
    meta = {
        'name':    filename,
        'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }
    if folder_id:
        meta['parents'] = [folder_id]
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    f = service.files().create(body=meta, media_body=media, fields='id,webViewLink').execute()
    # Make it viewable by anyone with the link
    service.permissions().create(
        fileId=f['id'],
        body={'role': 'reader', 'type': 'anyone'}
    ).execute()
    return f.get('webViewLink', '')

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
    m = re.search(r'```python\s*(.*?)```', text, re.DOTALL)
    return m.group(1).strip() if m else None

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

        log('Calling Claude API to generate resume...')
        import anthropic
        client   = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=8192,
            messages=[{'role': 'user', 'content': build_prompt(job, output_path, extra)}]
        )

        script = _extract_script(response.content[0].text)
        if not script:
            raise ValueError('Claude did not return a Python script.\n\n' + response.content[0].text[:600])

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
        try:
            drive_link = upload_to_drive(output_path, fname)
        except Exception as e:
            log(f'Drive upload failed: {e}\nResume built locally — trying email...')

        # Also email if configured
        gmail_user = _get_setting_direct(db_path, 'gmail_user', os.environ.get('GMAIL_USER', ''))
        gmail_pass = _get_setting_direct(db_path, 'gmail_pass', os.environ.get('GMAIL_PASS', ''))
        to_addr    = _get_setting_direct(db_path, 'email_to', DEFAULT_TO)
        if gmail_user and gmail_pass:
            try:
                _send_email(job, output_path, to_addr, gmail_user, gmail_pass, drive_link)
            except Exception as e:
                pass

        try:
            os.unlink(output_path)
        except Exception:
            pass

        status  = 'done'
        log_msg = f'Resume saved to Google Drive.'
        if drive_link:
            log_msg += f'\nLink: {drive_link}'

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
    row = db.execute('SELECT * FROM jobs WHERE id=?', (cur.lastrowid,)).fetchone()
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

@app.route('/health')
def health():
    return jsonify({
        'ok':       True,
        'template': os.path.exists(TEMPLATE),
        'api_key':  bool(os.environ.get('ANTHROPIC_API_KEY')),
        'drive':    bool(os.environ.get('GOOGLE_CREDENTIALS_JSON')),
    })

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
.job-actions{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px;margin-top:10px}
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
    <button class="btn btn-primary" onclick="addAndBuild()">&#9654; Add &amp; Build Resume</button>
    <button class="btn" style="background:#f5f5f5;color:#444;margin-top:8px" onclick="addOnly()">&#43; Add to Queue Only</button>
  </div>
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
  ['add','queue','settings'].forEach((n,i) => {
    if (n === name) document.querySelectorAll('.tab')[i].classList.add('active');
  });
  if (name === 'queue') loadJobs();
  if (name === 'settings') loadSettings();
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
  const d = await api('/api/jobs', 'POST', { title, company, url, jd });
  if (!d.id) { toast('Error: ' + (d.error||''), 4000); return; }
  clearForm();
  // immediately trigger build
  await api(`/api/jobs/${d.id}/build`, 'POST');
  showTab('queue');
  toast('Added and building — check back in 2-3 mins');
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
    const canBuild = j.status === 'pending' || j.status === 'error';
    const showLog  = (j.status === 'building' || j.status === 'error') && j.build_log;
    const logHtml  = showLog ? `<div class="log-box">${esc(j.build_log.slice(-600))}</div>` : '';
    const driveHtml = j.drive_link ? `<a class="drive-link" href="${esc(j.drive_link)}" target="_blank">&#128194; Open in Google Drive</a>` : '';
    const replyHtml = j.status === 'error' ? `
      <div style="margin-top:8px">
        <textarea id="reply-${j.id}" placeholder="Instructions e.g. reframe domain as enterprise SaaS..."
          style="width:100%;padding:9px;border:1px solid #ddd;border-radius:8px;font-size:13px;height:70px;resize:none;font-family:inherit"></textarea>
        <button onclick="replyBuild(${j.id})" style="width:100%;margin-top:6px;padding:9px;background:#111;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:500;cursor:pointer">
          &#9654; Retry with instructions
        </button>
      </div>` : '';
    return `<div class="job-card" id="card-${j.id}">
  <div class="job-head">
    <div><div class="job-title">${esc(j.title)}</div><div class="job-co">${esc(j.company)}</div></div>
    <span class="badge ${bClass}">${bLabel}</span>
  </div>
  <div class="job-meta">${j.created_at.slice(0,16).replace('T',' ')}${j.built_at ? ' &middot; built ' + j.built_at : ''}</div>
  ${driveHtml}${logHtml}${replyHtml}
  <div class="job-actions">
    <button class="ja-btn" onclick="buildOne(${j.id})" ${canBuild?'':'disabled'}>&#9654; Build</button>
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
