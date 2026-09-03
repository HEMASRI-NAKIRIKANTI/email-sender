"""
Resume Mailer — paste a job requirement, AI drafts a tailored email,
and it gets sent to the recruiter with your resume attached.

Supports multiple profiles in one session (e.g. different people sharing a
browser, or one person with multiple resume variants). Nothing is written to
disk or a database — everything lives only in this browser tab's
session_state and is cleared when the tab closes.
"""

import io
import json
import re
import smtplib
import ssl
import time
import uuid
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import streamlit as st

st.set_page_config(page_title="Resume Mailer", page_icon="📧", layout="centered")

WORK_AUTH_OPTIONS = [
    "STEM OPT", "OPT (Post-Completion)", "H1B", "H1B Transfer", "Green Card",
    "US Citizen", "CPT", "TN Visa", "Other / custom",
]
AVAILABILITY_OPTIONS = [
    "Immediate Joiner", "2 Weeks Notice", "1 Month Notice", "Flexible", "Other / custom",
]
EMAIL_REGEX = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+[a-zA-Z]"
PHONE_REGEX = r"\+?\d[\d\-.\s()]{7,}\d"

# --------------------------------------------------------------------------
# Text extraction / formatting helpers
# --------------------------------------------------------------------------

def extract_resume_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    try:
        if name.endswith(".pdf"):
            import pdfplumber
            text = []
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                for page in pdf.pages:
                    text.append(page.extract_text() or "")
            return "\n".join(text).strip()
        if name.endswith(".docx"):
            import docx
            doc = docx.Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs).strip()
        return data.decode("utf-8", errors="ignore").strip()
    except Exception as e:
        st.warning(f"Couldn't extract text from resume for AI context ({e}). "
                    "The file will still be attached to the email as-is.")
        return ""


def markdown_bold_to_html(text: str) -> str:
    """Convert **bold** markdown + newlines into a simple HTML email body."""
    escaped = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    bolded = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped)
    paragraphs = bolded.split("\n\n")
    html_paragraphs = "".join(f"<p>{p.replace(chr(10), '<br>')}</p>" for p in paragraphs if p.strip())
    return f"<html><body style='font-family:Calibri,\"Segoe UI\",Arial,sans-serif;font-size:15px;line-height:1.5;'>{html_paragraphs}</body></html>"


def strip_markdown_bold(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"\1", text)


_BOLD_MAP = {}
for i in range(26):
    _BOLD_MAP[chr(ord('A') + i)] = chr(0x1D400 + i)  # 𝗔-𝗭
    _BOLD_MAP[chr(ord('a') + i)] = chr(0x1D41A + i)  # 𝗮-𝘇
for i in range(10):
    _BOLD_MAP[chr(ord('0') + i)] = chr(0x1D7CE + i)  # 𝟬-𝟵


def to_unicode_bold(text: str) -> str:
    """Render plain text using Unicode Mathematical Bold codepoints, so it
    shows up bold in an email SUBJECT line (which can't render HTML/markdown)."""
    return "".join(_BOLD_MAP.get(ch, ch) for ch in text)


def build_facts_block(profile: dict) -> str:
    lines = []
    if profile.get("current_location"):
        lines.append(f"**Current Location:** {profile['current_location']}")
    if profile.get("work_authorization"):
        lines.append(f"**Work Authorization:** {profile['work_authorization']}")
    if profile.get("years_experience"):
        lines.append(f"**Professional Experience:** {profile['years_experience']}")
    if profile.get("availability"):
        lines.append(f"**Availability:** {profile['availability']}")
    return "\n".join(lines)


def build_signature(profile: dict) -> str:
    """Contact block only (name is placed separately, right under 'Best Regards,').
    Labels are bold, values are plain — matches the 'Phone: ..., Gmail: ...' style."""
    lines = []
    if profile.get("phone"):
        lines.append(f"**Phone:** {profile['phone']}")
    if profile.get("contact_email"):
        lines.append(f"**Gmail:** {profile['contact_email']}")
    if profile.get("linkedin"):
        lines.append(f"**LinkedIn:** {profile['linkedin']}")
    if profile.get("github"):
        lines.append(f"**GitHub:** {profile['github']}")
    return "\n".join(lines)


def assemble_email(profile: dict, narrative: str, ai_subject: str, greeting_name: str) -> tuple[str, str]:
    """Builds the final body (greeting + AI pitch + facts + closing + signature)
    and a Unicode-bold subject. Shared by both the single and bulk send flows."""
    facts_block = build_facts_block(profile)
    signature = build_signature(profile)
    full_name = profile.get("full_name", "")

    greeting = f"Dear {greeting_name or 'Hiring Team'},\n\nI hope this message finds you well. "
    full_body = greeting + narrative
    if facts_block:
        full_body += f"\n\nBelow are my details\n\n{facts_block}"
    full_body += (
        "\n\nI have attached my updated resume for your review. I look forward to "
        "the opportunity to discuss how I can contribute to your team."
        "\n\nThank you for your consideration."
    )
    full_body += f"\n\nBest Regards,\n{full_name}"
    if signature:
        full_body += f"\n{signature}"

    subject_text = (ai_subject or "").strip() or "Application"
    if full_name:
        subject_text = f"{subject_text} - {full_name}"
    bold_subject = to_unicode_bold(subject_text)

    return full_body, bold_subject


# --------------------------------------------------------------------------
# AI generation — static (cacheable) context first, dynamic job req last
# --------------------------------------------------------------------------

def build_static_context(profile: dict) -> str:
    """Everything that stays IDENTICAL across emails in this session.
    Putting this first (and unchanged) is what makes prompt caching kick in:
    OpenAI caches matching prefixes >1024 tokens automatically; for Claude we
    mark this block with an explicit cache_control breakpoint below."""
    return f"""CANDIDATE PROFILE
Name: {profile.get('full_name', '')}
Years of experience: {profile.get('years_experience', '')}
Current location: {profile.get('current_location', '')}
Work authorization: {profile.get('work_authorization', '')}
Availability: {profile.get('availability', '')}

RESUME CONTENT
{profile.get('resume_text') or '(no resume text extracted; write generically but confidently)'}
"""


def build_dynamic_instructions(job_description: str, tone: str) -> str:
    return f"""Write a short, {tone.lower()} email pitch (2-3 short paragraphs, under 200 words)
applying for the role described below, using the candidate profile and resume above.
Reference only real resume content — never invent skills, achievements, or numbers not in the resume.
Wrap 3-5 genuinely essential keywords/skills in **double asterisks** for emphasis (e.g. **AWS**, **Python**).

The email already opens with "Hello/Hi <Recruiter>,\\n\\nI hope this message finds you well. " before
your text, so your FIRST sentence must continue naturally straight after "well." — start it with
something like "I am writing to express my interest in the <role> position at <company>..." if a
role/company name is identifiable in the job requirement below, otherwise phrase the opening
naturally without inventing a company name.

Do NOT include any greeting/salutation yourself (that's already added), do NOT include a facts
list (location/visa/experience/availability — added separately), and do NOT include a closing
line, sign-off, or signature (added separately) — write ONLY the pitch paragraphs.

Return your answer as exactly two sections, nothing else:
SUBJECT: <subject line, no company boilerplate like "Application for" needed, just the role/value prop>
BODY:
<the pitch paragraphs>

--- JOB REQUIREMENT / DESCRIPTION ---
{job_description}
"""


def generate_email(provider: str, api_key: str, profile: dict, job_description: str, tone: str) -> tuple[str, str]:
    static_context = build_static_context(profile)
    dynamic_instructions = build_dynamic_instructions(job_description, tone)

    if provider == "OpenAI":
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        # Static context first, dynamic last -> stable prefix so OpenAI's
        # automatic prompt caching can match it across repeated generations.
        prompt = static_context + "\n" + dynamic_instructions
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
        )
        raw = resp.choices[0].message.content

    else:  # Claude
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=600,
            system=[
                {
                    "type": "text",
                    "text": static_context,
                    "cache_control": {"type": "ephemeral"},  # explicit cache breakpoint
                }
            ],
            messages=[{"role": "user", "content": dynamic_instructions}],
        )
        raw = resp.content[0].text

    subject, body = "Application", raw
    if "SUBJECT:" in raw and "BODY:" in raw:
        subject_part, body_part = raw.split("BODY:", 1)
        subject = subject_part.replace("SUBJECT:", "").strip()
        body = body_part.strip()
    return subject, body


def extract_phone_candidates(text: str) -> list[str]:
    """Best-effort phone number detection — keeps only matches with a
    plausible digit count so we don't pick up random numbers/years."""
    candidates = []
    for m in re.findall(PHONE_REGEX, text or ""):
        digits = re.sub(r"\D", "", m)
        if 7 <= len(digits) <= 15:
            candidates.append(m.strip())
    # de-dup while preserving order
    seen = set()
    out = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def extract_recruiter_fields_ai(provider: str, api_key: str, job_description: str) -> dict:
    """One lightweight AI call to pull company / recruiter name / role title
    out of free-text requirement. Regex handles email/phone separately since
    those are more reliably pattern-matched than inferred."""
    prompt = f"""Extract these fields from the recruiter message / job posting below.
Return STRICT JSON only — no markdown fences, no commentary, no extra keys:
{{"company": "", "recruiter_name": "", "role_title": ""}}
Use an empty string "" for anything you cannot confidently determine. Never invent a value.

--- MESSAGE ---
{job_description}
"""
    if provider == "OpenAI":
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
        )
        raw = resp.choices[0].message.content
    else:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text

    cleaned = re.sub(r"^```(json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    try:
        data = json.loads(cleaned)
    except Exception:
        data = {}
    return {
        "company": (data.get("company") or "").strip(),
        "recruiter_name": (data.get("recruiter_name") or "").strip(),
        "role_title": (data.get("role_title") or "").strip(),
    }


# --------------------------------------------------------------------------
# Email sending
# --------------------------------------------------------------------------

def send_email(smtp_server, smtp_port, sender_email, sender_password, use_tls,
                recruiter_email, cc_self, subject, body_markdown,
                attachment_bytes, attachment_filename):
    from email.header import Header

    msg = MIMEMultipart("mixed")
    msg["From"] = sender_email
    msg["To"] = recruiter_email
    msg["Subject"] = Header(subject, "utf-8")  # subject may contain Unicode bold glyphs
    if cc_self:
        msg["Cc"] = sender_email

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(strip_markdown_bold(body_markdown), "plain"))
    alt.attach(MIMEText(markdown_bold_to_html(body_markdown), "html"))
    msg.attach(alt)

    part = MIMEApplication(attachment_bytes, Name=attachment_filename)
    part["Content-Disposition"] = f'attachment; filename="{attachment_filename}"'
    msg.attach(part)

    recipients = [recruiter_email] + ([sender_email] if cc_self else [])
    context = ssl.create_default_context()
    if use_tls:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipients, msg.as_string())
    else:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipients, msg.as_string())


# --------------------------------------------------------------------------
# SQL Server (shared across all profiles) — profile persistence + send log
# --------------------------------------------------------------------------
# Uses pymssql (bundles FreeTDS in its Python wheel — no system ODBC driver
# needs to be installed on the host, which matters on Streamlit Community
# Cloud where you can't add Microsoft's driver repo) with SQL Server
# authentication, so anyone with the host/port/db/username/password you give
# them can connect remotely.

DB_PROFILE_COLUMNS = [
    "profile_id", "label", "full_name", "phone", "contact_email", "linkedin", "github",
    "current_location", "work_authorization", "years_experience", "availability",
    "provider", "api_key", "smtp_server", "smtp_port", "sender_email", "sender_password",
    "use_tls", "cc_self", "resume_filename", "resume_bytes",
]

LOG_COLUMNS = [
    "sent_at", "profile_label", "applicant_name", "recruiter_company", "recruiter_name",
    "recruiter_email", "recruiter_phone", "role_title", "subject", "job_description",
    "status", "error_message",
]


def get_db_connection(db: dict):
    import pymssql
    return pymssql.connect(
        server=db["server"],
        port=str(int(db["port"])),
        database=db["database"],
        user=db["username"],
        password=db["password"],
        login_timeout=10,
        timeout=10,
    )


def ensure_tables(conn):
    cur = conn.cursor()
    cur.execute("""
    IF OBJECT_ID('dbo.profiles', 'U') IS NULL
    CREATE TABLE dbo.profiles (
        profile_id NVARCHAR(64) PRIMARY KEY,
        label NVARCHAR(200), full_name NVARCHAR(200), phone NVARCHAR(50),
        contact_email NVARCHAR(200), linkedin NVARCHAR(300), github NVARCHAR(300),
        current_location NVARCHAR(200), work_authorization NVARCHAR(100),
        years_experience NVARCHAR(50), availability NVARCHAR(100),
        provider NVARCHAR(50), api_key NVARCHAR(500), smtp_server NVARCHAR(200),
        smtp_port INT, sender_email NVARCHAR(200), sender_password NVARCHAR(500),
        use_tls BIT, cc_self BIT, resume_filename NVARCHAR(300), resume_bytes VARBINARY(MAX),
        created_at DATETIME DEFAULT GETDATE(), updated_at DATETIME DEFAULT GETDATE()
    )
    """)
    cur.execute("""
    IF OBJECT_ID('dbo.email_log', 'U') IS NULL
    CREATE TABLE dbo.email_log (
        log_id INT IDENTITY(1,1) PRIMARY KEY,
        sent_at DATETIME DEFAULT GETDATE(),
        profile_label NVARCHAR(200), applicant_name NVARCHAR(200),
        recruiter_company NVARCHAR(200), recruiter_name NVARCHAR(200),
        recruiter_email NVARCHAR(200), recruiter_phone NVARCHAR(50),
        role_title NVARCHAR(300), subject NVARCHAR(500),
        job_description NVARCHAR(MAX), status NVARCHAR(20), error_message NVARCHAR(1000)
    )
    """)
    conn.commit()


def save_profile_to_db(conn, pid: str, profile: dict):
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM dbo.profiles WHERE profile_id = %s", (pid,))
    exists = cur.fetchone()[0] > 0
    vals = (
        profile.get("label"), profile.get("full_name"), profile.get("phone"),
        profile.get("contact_email"), profile.get("linkedin"), profile.get("github"),
        profile.get("current_location"), profile.get("work_authorization"),
        profile.get("years_experience"), profile.get("availability"),
        profile.get("provider"), profile.get("api_key"), profile.get("smtp_server"),
        int(profile.get("smtp_port") or 587), profile.get("sender_email"), profile.get("sender_password"),
        bool(profile.get("use_tls")), bool(profile.get("cc_self")),
        profile.get("resume_filename"), profile.get("resume_bytes"),
    )
    if exists:
        cur.execute("""
            UPDATE dbo.profiles SET
                label=%s, full_name=%s, phone=%s, contact_email=%s, linkedin=%s, github=%s,
                current_location=%s, work_authorization=%s, years_experience=%s, availability=%s,
                provider=%s, api_key=%s, smtp_server=%s, smtp_port=%s, sender_email=%s, sender_password=%s,
                use_tls=%s, cc_self=%s, resume_filename=%s, resume_bytes=%s, updated_at=GETDATE()
            WHERE profile_id=%s
        """, vals + (pid,))
    else:
        cur.execute("""
            INSERT INTO dbo.profiles
                (profile_id, label, full_name, phone, contact_email, linkedin, github,
                 current_location, work_authorization, years_experience, availability,
                 provider, api_key, smtp_server, smtp_port, sender_email, sender_password,
                 use_tls, cc_self, resume_filename, resume_bytes)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (pid,) + vals)
    conn.commit()


def load_profiles_from_db(conn) -> dict:
    cur = conn.cursor()
    cur.execute(f"SELECT {', '.join(DB_PROFILE_COLUMNS)} FROM dbo.profiles")
    rows = cur.fetchall()
    loaded = {}
    for row in rows:
        d = dict(zip(DB_PROFILE_COLUMNS, row))
        pid = d.pop("profile_id")
        resume_bytes = d.get("resume_bytes")
        d["resume_bytes"] = bytes(resume_bytes) if resume_bytes else None
        d["smtp_port"] = int(d.get("smtp_port") or 587)
        d["use_tls"] = bool(d.get("use_tls"))
        d["cc_self"] = bool(d.get("cc_self"))
        loaded[pid] = d
    return loaded


def log_email_attempt(db: dict, profile: dict, recruiter_email: str, recruiter_company: str,
                       recruiter_name: str, recruiter_phone: str, role_title: str,
                       subject: str, job_description: str, status: str, error_message: str = None):
    conn = get_db_connection(db)
    try:
        ensure_tables(conn)
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO dbo.email_log
                (profile_label, applicant_name, recruiter_company, recruiter_name,
                 recruiter_email, recruiter_phone, role_title, subject, job_description,
                 status, error_message)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (profile.get("label"), profile.get("full_name"), recruiter_company, recruiter_name,
              recruiter_email, recruiter_phone, role_title, subject, job_description,
              status, error_message))
        conn.commit()
    finally:
        conn.close()


def fetch_log(conn, today_only: bool, limit: int = 300):
    cur = conn.cursor()
    where = "WHERE CAST(sent_at AS DATE) = CAST(GETDATE() AS DATE)" if today_only else ""
    cur.execute(f"""
        SELECT TOP {int(limit)} {', '.join(LOG_COLUMNS)}
        FROM dbo.email_log {where}
        ORDER BY sent_at DESC
    """)
    rows = cur.fetchall()
    return [dict(zip(LOG_COLUMNS, row)) for row in rows]


# --------------------------------------------------------------------------
# Session state
# --------------------------------------------------------------------------
st.session_state.setdefault("profiles", {})       # id -> profile dict
st.session_state.setdefault("active_profile_id", None)
st.session_state.setdefault("bulk_requirements", [])   # list of requirement text strings
st.session_state.setdefault("bulk_queue", [])           # list of generated bulk-send items
# Per-profile generated drafts are stored dynamically as gen_subject_<id> / gen_body_<id>


def profile_form(existing: dict | None, form_key: str):
    """Renders the add/edit form. Returns the saved profile dict, or None."""
    d = existing or {}
    with st.form(form_key, clear_on_submit=False):
        st.markdown("**Profile label** (just for you to tell profiles apart)")
        label = st.text_input("Label", value=d.get("label", ""), placeholder="e.g. Jane – Backend roles")

        st.markdown("**Contact & links**")
        c1, c2 = st.columns(2)
        full_name = c1.text_input("Full name", value=d.get("full_name", ""))
        phone = c2.text_input("Phone number", value=d.get("phone", ""))
        contact_email = c1.text_input("Contact email (shown in signature)", value=d.get("contact_email", ""))
        linkedin = c2.text_input("LinkedIn URL", value=d.get("linkedin", ""))
        github = c1.text_input("GitHub URL", value=d.get("github", ""))

        st.markdown("**Candidate details**")
        current_location = c2.text_input("Current location", value=d.get("current_location", ""), placeholder="Liberty Hill, TX")
        wa_default = d.get("work_authorization", WORK_AUTH_OPTIONS[0])
        wa_index = WORK_AUTH_OPTIONS.index(wa_default) if wa_default in WORK_AUTH_OPTIONS else len(WORK_AUTH_OPTIONS) - 1
        work_auth_choice = c1.selectbox("Work authorization", WORK_AUTH_OPTIONS, index=wa_index)
        work_authorization = c1.text_input("Custom work authorization", value=d.get("work_authorization", "")) \
            if work_auth_choice == "Other / custom" else work_auth_choice
        years_experience = c2.text_input("Years of experience", value=d.get("years_experience", ""), placeholder="5+ Years")
        av_default = d.get("availability", AVAILABILITY_OPTIONS[0])
        av_index = AVAILABILITY_OPTIONS.index(av_default) if av_default in AVAILABILITY_OPTIONS else len(AVAILABILITY_OPTIONS) - 1
        avail_choice = c1.selectbox("Availability", AVAILABILITY_OPTIONS, index=av_index)
        availability = c1.text_input("Custom availability", value=d.get("availability", "")) \
            if avail_choice == "Other / custom" else avail_choice

        st.markdown("**AI provider**")
        provider = st.radio("Provider", ["OpenAI", "Claude"], horizontal=True,
                             index=0 if d.get("provider", "OpenAI") == "OpenAI" else 1)
        api_key = st.text_input(f"{provider} API key", type="password", value=d.get("api_key", ""))

        st.markdown("**Resume** (upload once, reused for every email from this profile)")
        uploaded_resume = st.file_uploader("Resume file", type=["pdf", "docx", "txt"], key=f"resume_{form_key}")

        st.markdown("**SMTP (your email account)**")
        s1, s2 = st.columns(2)
        sender_email = s1.text_input("SMTP login email", value=d.get("sender_email", ""))
        sender_password = s2.text_input("App password / SMTP password", type="password", value=d.get("sender_password", ""))
        smtp_server = s1.text_input("SMTP server", value=d.get("smtp_server", "smtp.gmail.com"))
        smtp_port = s2.number_input("Port", value=int(d.get("smtp_port", 587)), step=1)
        use_tls = st.checkbox("Use STARTTLS (587)", value=d.get("use_tls", True))
        cc_self = st.checkbox("CC myself on the email", value=d.get("cc_self", True))

        submitted = st.form_submit_button("💾 Save profile", type="primary", use_container_width=True)

        if submitted:
            if not label.strip():
                st.error("Give this profile a label.")
                return None
            resume_bytes = d.get("resume_bytes")
            resume_filename = d.get("resume_filename")
            resume_text = d.get("resume_text", "")
            if uploaded_resume is not None:
                resume_bytes = uploaded_resume.getvalue()
                resume_filename = uploaded_resume.name
                resume_text = extract_resume_text(uploaded_resume)
            if not resume_bytes:
                st.error("Upload a resume for this profile.")
                return None
            return {
                "label": label, "full_name": full_name, "phone": phone,
                "contact_email": contact_email, "linkedin": linkedin, "github": github,
                "current_location": current_location, "work_authorization": work_authorization,
                "years_experience": years_experience, "availability": availability,
                "provider": provider, "api_key": api_key,
                "resume_bytes": resume_bytes, "resume_filename": resume_filename, "resume_text": resume_text,
                "sender_email": sender_email, "sender_password": sender_password,
                "smtp_server": smtp_server, "smtp_port": smtp_port, "use_tls": use_tls, "cc_self": cc_self,
            }
    return None


# --------------------------------------------------------------------------
# Sidebar — profile manager
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("👤 Profiles")
    st.caption("Add as many people/resume variants as you like. Nothing here is saved to a server — "
               "it only lives in this browser tab.")

    profile_ids = list(st.session_state.profiles.keys())
    labels = ["➕ Add new profile"] + [st.session_state.profiles[pid]["label"] for pid in profile_ids]
    default_index = 0
    if st.session_state.active_profile_id in profile_ids:
        default_index = profile_ids.index(st.session_state.active_profile_id) + 1
    choice = st.selectbox("Active profile", labels, index=default_index)

    if choice == "➕ Add new profile":
        new_profile = profile_form(None, form_key="new_profile")
        if new_profile:
            pid = str(uuid.uuid4())
            st.session_state.profiles[pid] = new_profile
            st.session_state.active_profile_id = pid
            st.success(f"Profile '{new_profile['label']}' saved.")
            st.rerun()
    else:
        pid = profile_ids[labels.index(choice) - 1]
        st.session_state.active_profile_id = pid
        with st.expander("✏️ Edit this profile"):
            updated = profile_form(st.session_state.profiles[pid], form_key=f"edit_{pid}")
            if updated:
                st.session_state.profiles[pid] = updated
                st.success("Profile updated.")
                st.rerun()
        if st.button("🗑️ Delete this profile", use_container_width=True):
            del st.session_state.profiles[pid]
            st.session_state.active_profile_id = None
            st.rerun()

    st.subheader("Tone")
    tone = st.select_slider("Email tone", options=["Formal", "Professional", "Friendly"], value="Professional")

    st.divider()
    st.header("🗄️ SQL Server (shared)")
    st.caption("One shared connection for everyone using this app — saves profiles for next time "
               "and logs every sent email (recruiter, company, JD, timestamp) for tracking.")

    db = st.session_state.setdefault("db_config", {
        "server": "", "port": 1433, "database": "", "username": "", "password": "",
    })
    db["server"] = st.text_input("Server / host", value=db["server"], placeholder="myserver.database.windows.net")
    dcol1, dcol2 = st.columns(2)
    db["port"] = dcol1.number_input("Port", value=int(db["port"]), step=1)
    db["database"] = dcol2.text_input("Database", value=db["database"])
    db["username"] = st.text_input("SQL username", value=db["username"])
    db["password"] = st.text_input("SQL password", type="password", value=db["password"])

    tcol1, tcol2 = st.columns(2)
    with tcol1:
        if st.button("🔌 Test", use_container_width=True):
            try:
                conn = get_db_connection(db)
                conn.close()
                st.success("Connected ✅")
            except Exception as e:
                st.error(f"Failed: {e}")
    with tcol2:
        if st.button("🛠️ Setup tables", use_container_width=True):
            try:
                conn = get_db_connection(db)
                ensure_tables(conn)
                conn.close()
                st.success("Tables ready ✅")
            except Exception as e:
                st.error(f"Failed: {e}")

    ucol1, ucol2 = st.columns(2)
    with ucol1:
        if st.button("⬆️ Save profiles", use_container_width=True):
            try:
                conn = get_db_connection(db)
                ensure_tables(conn)
                for pid, profile in st.session_state.profiles.items():
                    save_profile_to_db(conn, pid, profile)
                conn.close()
                st.success(f"Saved {len(st.session_state.profiles)} profile(s) ✅")
            except Exception as e:
                st.error(f"Failed: {e}")
    with ucol2:
        if st.button("⬇️ Load profiles", use_container_width=True):
            try:
                conn = get_db_connection(db)
                ensure_tables(conn)
                loaded = load_profiles_from_db(conn)
                conn.close()
                st.session_state.profiles.update(loaded)
                st.success(f"Loaded {len(loaded)} profile(s) ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Failed: {e}")

    st.session_state["db_log_enabled"] = st.checkbox(
        "📝 Log every sent email to SQL Server", value=st.session_state.get("db_log_enabled", False),
    )
    st.caption("⚠️ API keys and SMTP passwords are stored in the `profiles` table as plain values on "
               "'Save profiles' — anyone with DB access can read them. Restrict DB access accordingly.")

# --------------------------------------------------------------------------
# Main panel
# --------------------------------------------------------------------------
st.title("📧 Resume Mailer")

if not st.session_state.profiles:
    st.info("👈 Add a profile in the sidebar to get started (contact info, resume, work auth, SMTP login).")
    st.stop()

all_ids = list(st.session_state.profiles.keys())
all_labels = [st.session_state.profiles[pid]["label"] for pid in all_ids]

tab_compose, tab_bulk, tab_log = st.tabs(["✉️ Compose & Send", "📦 Bulk Send", "📊 Tracking"])

with tab_compose:
    st.write("Paste a recruiter's job requirement below and pick which profiles it applies to. "
             "Each selected profile gets its **own** personalized email — generated from *their* "
             "resume and details, sent from *their* SMTP account — not a copy of one profile's email.")

    job_description = st.text_area(
        "Paste the job requirement / recruiter message here", height=200,
        placeholder="e.g. 'Looking for a Senior Backend Engineer with 5+ years in Python, AWS, and distributed "
                    "systems. Reach out to jane.recruiter@company.com if interested.'",
    )

    # --- Recruiter email (regex-detected) ---
    detected_emails = sorted(set(re.findall(EMAIL_REGEX, job_description or "")))
    st.session_state.setdefault("recruiter_email", "")

    detected_email_default = ""
    if len(detected_emails) == 1:
        detected_email_default = detected_emails[0]
    elif len(detected_emails) > 1:
        detected_email_default = st.selectbox(
            "Multiple email addresses found in the requirement — which one is the recruiter's?",
            detected_emails,
        )
    if detected_email_default and not st.session_state.recruiter_email:
        st.session_state.recruiter_email = detected_email_default

    email_col, refresh_col = st.columns([4, 1])
    with email_col:
        recruiter_email = st.text_input("Recruiter's email address", key="recruiter_email")
    with refresh_col:
        st.write("")
        if st.button("🔄 Use detected", use_container_width=True, disabled=not detected_email_default,
                     help="Overwrite the field above with the email found in the requirement text."):
            st.session_state.recruiter_email = detected_email_default
            st.rerun()

    if detected_emails:
        st.caption(f"📩 Found {len(detected_emails)} email address(es) in the requirement text.")
    elif job_description.strip():
        st.caption("No email address found in the requirement text — enter it manually above.")

    # --- Recruiter name / company / role / phone (regex + one AI call) ---
    st.session_state.setdefault("recruiter_name", "")
    st.session_state.setdefault("recruiter_company", "")
    st.session_state.setdefault("role_title", "")
    st.session_state.setdefault("recruiter_phone", "")

    selected_labels = st.multiselect(
        "Send for these profiles", all_labels, default=all_labels,
        help="Each one is generated and sent independently — its own resume, its own facts/signature, its own SMTP login.",
    )
    selected_ids = [all_ids[all_labels.index(lbl)] for lbl in selected_labels]

    extract_clicked = st.button(
        "🔍 Extract recruiter, company & role details from the requirement",
        use_container_width=True, disabled=not (job_description.strip() and selected_ids),
    )
    if extract_clicked:
        helper_profile = st.session_state.profiles[selected_ids[0]]
        if not helper_profile.get("api_key"):
            st.error(f"'{helper_profile['label']}' has no {helper_profile['provider']} API key saved — "
                     "needed to run this extraction. Edit that profile, or select a different profile first.")
        else:
            with st.spinner("Reading the requirement..."):
                try:
                    fields = extract_recruiter_fields_ai(
                        helper_profile["provider"], helper_profile["api_key"], job_description,
                    )
                    if fields["company"] and not st.session_state.recruiter_company:
                        st.session_state.recruiter_company = fields["company"]
                    if fields["recruiter_name"] and not st.session_state.recruiter_name:
                        st.session_state.recruiter_name = fields["recruiter_name"]
                    if fields["role_title"] and not st.session_state.role_title:
                        st.session_state.role_title = fields["role_title"]
                    phone_candidates = extract_phone_candidates(job_description)
                    if phone_candidates and not st.session_state.recruiter_phone:
                        st.session_state.recruiter_phone = phone_candidates[0]
                    st.success("Extracted — review the fields below before sending.")
                except Exception as e:
                    st.error(f"Extraction failed: {e}")

    rc1, rc2 = st.columns(2)
    with rc1:
        recruiter_name = st.text_input(
            "Recruiter's name (for the greeting)", key="recruiter_name", placeholder="e.g. Srinija",
            help="Used for 'Dear <name>,' at the top of the email. Leave blank to use 'Hiring Team'.",
        )
        recruiter_company = st.text_input("Recruiter's company", key="recruiter_company")
    with rc2:
        recruiter_phone = st.text_input("Recruiter's phone number", key="recruiter_phone")
        role_title = st.text_input("Role / job title", key="role_title")

    if recruiter_email or recruiter_phone or role_title:
        st.caption("📋 Copy:")
        st.code(f"Role: {role_title}\nEmail: {recruiter_email}\nPhone: {recruiter_phone}", language=None)

    gen_col, _ = st.columns([1, 3])
    with gen_col:
        n = len(selected_ids)
        generate_clicked = st.button(
            f"✨ Generate {'email' if n == 1 else f'{n} personalized emails'}",
            type="primary", use_container_width=True, disabled=(n == 0),
        )

    if generate_clicked:
        if not job_description.strip():
            st.error("Paste the job requirement first.")
        else:
            progress = st.progress(0.0)
            for i, pid in enumerate(selected_ids):
                profile = st.session_state.profiles[pid]
                progress.progress(i / len(selected_ids), text=f"Drafting for {profile['label']}...")
                if not profile.get("api_key"):
                    st.error(f"'{profile['label']}': no {profile['provider']} API key saved — skipped.")
                    continue
                try:
                    subject, narrative = generate_email(
                        profile["provider"], profile["api_key"], profile, job_description, tone,
                    )
                    full_body, bold_subject = assemble_email(
                        profile, narrative, subject, recruiter_name.strip() or "Hiring Team",
                    )
                    st.session_state[f"gen_subject_{pid}"] = bold_subject
                    st.session_state[f"gen_body_{pid}"] = full_body
                except Exception as e:
                    st.error(f"'{profile['label']}': generation failed ({e}).")
            progress.progress(1.0, text="Done.")

    ready_ids = [pid for pid in selected_ids if st.session_state.get(f"gen_body_{pid}")]

    if ready_ids:
        st.divider()
        st.subheader(f"Review & send ({len(ready_ids)} ready)")

        for pid in ready_ids:
            profile = st.session_state.profiles[pid]
            subj_key, body_key = f"gen_subject_{pid}", f"gen_body_{pid}"
            with st.expander(f"✉️ {profile['label']}", expanded=(len(ready_ids) == 1)):
                st.session_state[subj_key] = st.text_input("Subject", value=st.session_state[subj_key], key=f"subj_input_{pid}")
                st.session_state[body_key] = st.text_area(
                    "Body (use **word** to bold it in the sent email)",
                    value=st.session_state[body_key], height=280, key=f"body_input_{pid}",
                )
                st.markdown("**Preview:**")
                st.markdown(st.session_state[body_key].replace("\n", "  \n"))
                st.caption(f"📎 {profile['resume_filename']}  |  From: {profile.get('sender_email', '(not set)')}")

                if st.button("📤 Send this one", key=f"send_{pid}", use_container_width=True):
                    if not recruiter_email:
                        st.error("Enter the recruiter's email address above.")
                    elif not profile.get("sender_email") or not profile.get("sender_password"):
                        st.error(f"'{profile['label']}' is missing SMTP email/password — edit it in the sidebar.")
                    else:
                        try:
                            send_email(
                                profile["smtp_server"], int(profile["smtp_port"]),
                                profile["sender_email"], profile["sender_password"],
                                profile["use_tls"], recruiter_email, profile["cc_self"],
                                st.session_state[subj_key], st.session_state[body_key],
                                profile["resume_bytes"], profile["resume_filename"],
                            )
                            st.success(f"Sent for {profile['label']} ✅")
                            if st.session_state.get("db_log_enabled"):
                                try:
                                    log_email_attempt(
                                        st.session_state["db_config"], profile, recruiter_email,
                                        recruiter_company, recruiter_name, recruiter_phone, role_title,
                                        st.session_state[subj_key], job_description, "sent",
                                    )
                                except Exception as e:
                                    st.warning(f"Sent, but logging to SQL Server failed: {e}")
                        except Exception as e:
                            st.error(f"Failed to send for {profile['label']}: {e}")
                            if st.session_state.get("db_log_enabled"):
                                try:
                                    log_email_attempt(
                                        st.session_state["db_config"], profile, recruiter_email,
                                        recruiter_company, recruiter_name, recruiter_phone, role_title,
                                        st.session_state[subj_key], job_description, "failed", str(e),
                                    )
                                except Exception:
                                    pass

        st.divider()
        send_all_col, _ = st.columns([1, 3])
        with send_all_col:
            send_all_clicked = st.button(f"📤 Send all {len(ready_ids)}", type="primary", use_container_width=True)

        if send_all_clicked:
            if not recruiter_email:
                st.error("Enter the recruiter's email address above.")
            else:
                for pid in ready_ids:
                    profile = st.session_state.profiles[pid]
                    subj_key, body_key = f"gen_subject_{pid}", f"gen_body_{pid}"
                    if not profile.get("sender_email") or not profile.get("sender_password"):
                        st.error(f"Skipped '{profile['label']}': missing SMTP email/password.")
                        continue
                    try:
                        send_email(
                            profile["smtp_server"], int(profile["smtp_port"]),
                            profile["sender_email"], profile["sender_password"],
                            profile["use_tls"], recruiter_email, profile["cc_self"],
                            st.session_state[subj_key], st.session_state[body_key],
                            profile["resume_bytes"], profile["resume_filename"],
                        )
                        st.success(f"Sent for {profile['label']} ✅")
                        if st.session_state.get("db_log_enabled"):
                            try:
                                log_email_attempt(
                                    st.session_state["db_config"], profile, recruiter_email,
                                    recruiter_company, recruiter_name, recruiter_phone, role_title,
                                    st.session_state[subj_key], job_description, "sent",
                                )
                            except Exception as e:
                                st.warning(f"Sent for {profile['label']}, but logging failed: {e}")
                    except Exception as e:
                        st.error(f"Failed for {profile['label']}: {e}")
                        if st.session_state.get("db_log_enabled"):
                            try:
                                log_email_attempt(
                                    st.session_state["db_config"], profile, recruiter_email,
                                    recruiter_company, recruiter_name, recruiter_phone, role_title,
                                    st.session_state[subj_key], job_description, "failed", str(e),
                                )
                            except Exception:
                                pass

    st.divider()
    st.caption("🔒 Note: with 'Store everything in DB' in effect, API keys and SMTP passwords are "
               "saved in plain columns when you click 'Save profiles' — restrict who can reach that "
               "database. Resumes, once saved, are stored as VARBINARY(MAX).\n\n"
               "⚡ Prompt caching: resume/profile context is sent as a stable prefix so OpenAI's automatic "
               "caching applies; for Claude it's marked with an explicit cache_control breakpoint.")

with tab_bulk:
    st.write("Paste **multiple** job requirements at once, separated by a line containing just `---`. "
             "Each requirement gets its own recruiter/company/role extraction, and a personalized email "
             "per selected profile. Sends go out one at a time with a delay between them, so a recruiter's "
             "inbox (or your SMTP account) doesn't get hit with a burst all at once.")

    bulk_text = st.text_area(
        "Paste requirements here, separated by lines containing only ---", height=260,
        placeholder="Looking for a Backend Engineer... jane@company.com\n\n---\n\n"
                    "Looking for a Data Scientist... john@otherco.com\n\n---\n\n"
                    "Looking for a DevOps Engineer... hr@thirdco.com",
    )

    parse_col, _ = st.columns([1, 3])
    with parse_col:
        parse_clicked = st.button("📋 Parse requirements", use_container_width=True, disabled=not bulk_text.strip())

    if parse_clicked:
        parts = re.split(r"\n\s*-{3,}\s*\n", bulk_text.strip())
        parts = [p.strip() for p in parts if p.strip()]
        st.session_state.bulk_requirements = parts
        st.session_state.bulk_queue = []  # parsing invalidates any previous queue
        st.success(f"Found {len(parts)} requirement(s).")

    bulk_reqs = st.session_state.get("bulk_requirements", [])
    if bulk_reqs:
        st.caption(f"📄 {len(bulk_reqs)} requirement(s) parsed:")
        for i, r in enumerate(bulk_reqs):
            st.caption(f"{i + 1}. {r[:100]}{'...' if len(r) > 100 else ''}")

        bulk_selected_labels = st.multiselect(
            "Send for these profiles (applied to every requirement above)",
            all_labels, default=all_labels, key="bulk_profile_select",
        )
        bulk_selected_ids = [all_ids[all_labels.index(lbl)] for lbl in bulk_selected_labels]

        gen_bulk_col, _ = st.columns([1, 3])
        with gen_bulk_col:
            n_emails = len(bulk_reqs) * len(bulk_selected_ids)
            generate_bulk_clicked = st.button(
                f"✨ Generate {n_emails} email(s)", type="primary",
                use_container_width=True, disabled=(n_emails == 0),
            )

        if generate_bulk_clicked:
            helper_profile = st.session_state.profiles[bulk_selected_ids[0]]
            queue = []
            progress = st.progress(0.0)
            for ri, req_text in enumerate(bulk_reqs):
                progress.progress(ri / len(bulk_reqs), text=f"Processing requirement {ri + 1}/{len(bulk_reqs)}...")

                found_emails = sorted(set(re.findall(EMAIL_REGEX, req_text)))
                req_recruiter_email = found_emails[0] if found_emails else ""
                found_phones = extract_phone_candidates(req_text)
                req_recruiter_phone = found_phones[0] if found_phones else ""

                fields = {"company": "", "recruiter_name": "", "role_title": ""}
                if helper_profile.get("api_key"):
                    try:
                        fields = extract_recruiter_fields_ai(
                            helper_profile["provider"], helper_profile["api_key"], req_text,
                        )
                    except Exception as e:
                        st.warning(f"Requirement {ri + 1}: extraction failed ({e}) — company/name/role left blank.")

                for pid in bulk_selected_ids:
                    profile = st.session_state.profiles[pid]
                    if not profile.get("api_key"):
                        queue.append({
                            "req_index": ri, "req_preview": req_text[:100], "job_description": req_text,
                            "profile_id": pid, "profile_label": profile["label"],
                            "recruiter_email": req_recruiter_email, "recruiter_company": fields["company"],
                            "recruiter_name": fields["recruiter_name"], "recruiter_phone": req_recruiter_phone,
                            "role_title": fields["role_title"], "subject": "", "body": "",
                            "status": "failed", "error": f"No {profile['provider']} API key saved.",
                        })
                        continue
                    try:
                        subject, narrative = generate_email(
                            profile["provider"], profile["api_key"], profile, req_text, tone,
                        )
                        full_body, bold_subject = assemble_email(
                            profile, narrative, subject, fields["recruiter_name"] or "Hiring Team",
                        )
                        queue.append({
                            "req_index": ri, "req_preview": req_text[:100], "job_description": req_text,
                            "profile_id": pid, "profile_label": profile["label"],
                            "recruiter_email": req_recruiter_email, "recruiter_company": fields["company"],
                            "recruiter_name": fields["recruiter_name"], "recruiter_phone": req_recruiter_phone,
                            "role_title": fields["role_title"], "subject": bold_subject, "body": full_body,
                            "status": "ready", "error": None,
                        })
                    except Exception as e:
                        queue.append({
                            "req_index": ri, "req_preview": req_text[:100], "job_description": req_text,
                            "profile_id": pid, "profile_label": profile["label"],
                            "recruiter_email": req_recruiter_email, "recruiter_company": fields["company"],
                            "recruiter_name": fields["recruiter_name"], "recruiter_phone": req_recruiter_phone,
                            "role_title": fields["role_title"], "subject": "", "body": "",
                            "status": "failed", "error": str(e),
                        })
            progress.progress(1.0, text="Done.")
            st.session_state.bulk_queue = queue

    bulk_queue = st.session_state.get("bulk_queue", [])
    if bulk_queue:
        st.divider()
        ready_count = sum(1 for it in bulk_queue if it["status"] == "ready")
        sent_count = sum(1 for it in bulk_queue if it["status"] == "sent")
        failed_count = sum(1 for it in bulk_queue if it["status"] == "failed")
        st.subheader(f"Queue — {ready_count} ready, {sent_count} sent, {failed_count} failed")

        req_indices = sorted(set(it["req_index"] for it in bulk_queue))
        for ri in req_indices:
            items = [it for it in bulk_queue if it["req_index"] == ri]
            with st.expander(f"📄 Requirement {ri + 1}: {items[0]['req_preview']}...", expanded=False):
                for item in items:
                    key_base = f"bulk_{ri}_{item['profile_id']}"
                    icon = {"ready": "⬜", "sent": "✅", "failed": "❌"}.get(item["status"], "⬜")
                    st.markdown(f"**{icon} {item['profile_label']}**")
                    if item["status"] == "failed":
                        st.caption(f"Error: {item.get('error', '')}")
                    else:
                        rcol1, rcol2 = st.columns(2)
                        with rcol1:
                            item["recruiter_email"] = st.text_input(
                                "Recruiter email", value=item["recruiter_email"], key=f"{key_base}_email",
                            )
                            item["recruiter_company"] = st.text_input(
                                "Recruiter's company", value=item.get("recruiter_company", ""), key=f"{key_base}_company",
                            )
                        with rcol2:
                            item["recruiter_phone"] = st.text_input(
                                "Recruiter phone", value=item.get("recruiter_phone", ""), key=f"{key_base}_phone",
                            )
                            item["role_title"] = st.text_input(
                                "Role / job title", value=item.get("role_title", ""), key=f"{key_base}_role",
                            )
                        st.caption("📋 Copy:")
                        st.code(
                            f"Role: {item['role_title']}\nEmail: {item['recruiter_email']}\nPhone: {item['recruiter_phone']}",
                            language=None,
                        )
                        item["subject"] = st.text_input("Subject", value=item["subject"], key=f"{key_base}_subj")
                        item["body"] = st.text_area("Body", value=item["body"], height=180, key=f"{key_base}_body")
                    st.divider()

        st.subheader("Send the queue")
        delay_choice = st.radio(
            "Delay between each send", ["Immediate", "1 minute", "3 minutes"], horizontal=True,
            help="Applied between every individual email sent, so a burst of sends doesn't hit "
                 "recruiters' inboxes (or trip spam filters on your SMTP account) all at once.",
        )
        delay_seconds = {"Immediate": 0, "1 minute": 60, "3 minutes": 180}[delay_choice]
        pending = [it for it in bulk_queue if it["status"] == "ready"]
        if pending and delay_seconds:
            est_minutes = (len(pending) - 1) * delay_seconds / 60
            st.caption(f"⏱️ Estimated total time: ~{est_minutes:.0f} min for {len(pending)} pending email(s). "
                       "Keep this tab open while it runs.")

        send_bulk_col, clear_col = st.columns([1, 1])
        with send_bulk_col:
            start_sending_clicked = st.button(
                f"🚀 Send {len(pending)} pending", type="primary",
                use_container_width=True, disabled=(len(pending) == 0),
            )
        with clear_col:
            if st.button("🗑️ Clear queue", use_container_width=True):
                st.session_state.bulk_queue = []
                st.session_state.bulk_requirements = []
                st.rerun()

        if start_sending_clicked:
            overall_progress = st.progress(0.0)
            status_placeholder = st.empty()
            total = len(pending)
            for i, item in enumerate(pending):
                if not item["recruiter_email"]:
                    item["status"] = "failed"
                    item["error"] = "No recruiter email set."
                    st.error(f"❌ {i + 1}/{total}: {item['profile_label']} — no recruiter email set, skipped.")
                    overall_progress.progress((i + 1) / total)
                    continue

                profile = st.session_state.profiles[item["profile_id"]]
                status_placeholder.info(f"📤 Sending {i + 1}/{total}: {item['profile_label']} → {item['recruiter_email']}...")
                try:
                    send_email(
                        profile["smtp_server"], int(profile["smtp_port"]),
                        profile["sender_email"], profile["sender_password"],
                        profile["use_tls"], item["recruiter_email"], profile["cc_self"],
                        item["subject"], item["body"],
                        profile["resume_bytes"], profile["resume_filename"],
                    )
                    item["status"] = "sent"
                    st.success(f"✅ Sent {i + 1}/{total}: {item['profile_label']} → {item['recruiter_email']}")
                    if st.session_state.get("db_log_enabled"):
                        try:
                            log_email_attempt(
                                st.session_state["db_config"], profile, item["recruiter_email"],
                                item["recruiter_company"], item["recruiter_name"], item["recruiter_phone"],
                                item["role_title"], item["subject"], item["job_description"], "sent",
                            )
                        except Exception as e:
                            st.warning(f"Sent, but logging failed: {e}")
                except Exception as e:
                    item["status"] = "failed"
                    item["error"] = str(e)
                    st.error(f"❌ Failed {i + 1}/{total}: {item['profile_label']} → {item['recruiter_email']}: {e}")
                    if st.session_state.get("db_log_enabled"):
                        try:
                            log_email_attempt(
                                st.session_state["db_config"], profile, item["recruiter_email"],
                                item["recruiter_company"], item["recruiter_name"], item["recruiter_phone"],
                                item["role_title"], item["subject"], item["job_description"], "failed", str(e),
                            )
                        except Exception:
                            pass

                overall_progress.progress((i + 1) / total)

                if delay_seconds and i < total - 1:
                    remaining = delay_seconds
                    while remaining > 0:
                        mins, secs = divmod(remaining, 60)
                        status_placeholder.info(f"⏳ Waiting {mins}m {secs:02d}s before next send ({i + 2}/{total})...")
                        time.sleep(1)
                        remaining -= 1

            status_placeholder.success("🎉 Bulk send complete.")

with tab_log:
    st.subheader("🗄️ Database browser")
    db = st.session_state.get("db_config", {})
    if not db.get("server"):
        st.info("Configure SQL Server in the sidebar to enable this.")
    else:
        view_choice = st.radio("View", ["📊 Sent Log", "👤 Profiles in DB"], horizontal=True)

        if view_choice == "📊 Sent Log":
            log_view = st.radio("Show", ["Today", "All (last 300)"], horizontal=True)
            if st.button("🔄 Refresh", key="refresh_log"):
                st.rerun()
            try:
                conn = get_db_connection(db)
                ensure_tables(conn)
                rows = fetch_log(conn, today_only=(log_view == "Today"))
                conn.close()
                if not rows:
                    st.write("No sent emails logged yet.")
                else:
                    import pandas as pd
                    df = pd.DataFrame(rows)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download CSV", df.to_csv(index=False).encode("utf-8"),
                        "sent_emails_log.csv", "text/csv",
                    )
            except Exception as e:
                st.error(f"Couldn't load the log: {e}")

        else:
            if st.button("🔄 Refresh", key="refresh_profiles"):
                st.rerun()
            show_secrets = st.checkbox("Show API keys & SMTP passwords (instead of masked)", value=False)
            try:
                conn = get_db_connection(db)
                ensure_tables(conn)
                loaded = load_profiles_from_db(conn)
                conn.close()
                if not loaded:
                    st.write("No profiles saved to the DB yet — use '⬆️ Save profiles' in the sidebar.")
                else:
                    import pandas as pd
                    display_cols = [c for c in DB_PROFILE_COLUMNS if c != "resume_bytes"]
                    rows = []
                    for pid, p in loaded.items():
                        row = {"profile_id": pid, **{c: p.get(c) for c in display_cols if c != "profile_id"}}
                        if not show_secrets:
                            row["api_key"] = "••••••" if row.get("api_key") else ""
                            row["sender_password"] = "••••••" if row.get("sender_password") else ""
                        rows.append(row)
                    df = pd.DataFrame(rows)
                    st.dataframe(df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Couldn't load profiles: {e}")