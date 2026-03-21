"""
Bookkeeping File Reviewer
Streamlit + ChatGPT — Senior CPA review tool for Alberta bookkeeping files.
"""
import streamlit as st
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Bookkeeping Reviewer",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Light theme with dark readable text ── */
.stApp { background: #f5f7fa !important; }
[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0;
}
[data-testid="stSidebar"] * { color: #1e293b !important; }

/* Welcome banner */
.welcome-banner {
    background: linear-gradient(135deg, #3b37cc 0%, #7c3aed 100%);
    border-radius: 14px; padding: 1.1rem 1.8rem;
    margin-bottom: 1.2rem; display: flex; align-items: center;
    box-shadow: 0 4px 18px rgba(59,55,204,.22);
}
.welcome-text {
    font-size: 22px; font-weight: 800; color: #ffffff;
    letter-spacing: .01em;
}
.welcome-sub {
    font-size: 12px; color: rgba(255,255,255,.72);
    margin-top: 2px; font-weight: 400;
}

/* Check cards — strong dark text */
.check-card {
    border-radius: 8px; padding: 10px 14px; margin: 5px 0;
    border-left: 4px solid; font-size: 13px;
    box-shadow: 0 1px 3px rgba(0,0,0,.07);
}
.check-critical { background: #fff0f0; border-color: #dc2626; color: #1e293b; }
.check-warning  { background: #fffbeb; border-color: #d97706; color: #1e293b; }
.check-info     { background: #eff6ff; border-color: #4f46e5; color: #1e293b; }
.check-pass     { background: #f0fdf4; border-color: #16a34a; color: #1e293b; }
.check-card strong { color: #0f172a !important; font-size: 13px; }

.check-badge {
    display: inline-block; padding: 2px 9px; border-radius: 4px;
    font-size: 10px; font-weight: 800; margin-right: 8px;
    letter-spacing: .05em;
}
.badge-critical { background: #dc2626; color: #ffffff; }
.badge-warning  { background: #d97706; color: #ffffff; }
.badge-info     { background: #4f46e5; color: #ffffff; }
.badge-pass     { background: #16a34a; color: #ffffff; }

/* Stat boxes */
.stat-box {
    background: #ffffff; border: 1px solid #e2e8f0;
    border-radius: 12px; padding: 16px; text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.stat-num { font-size: 28px; font-weight: 800; margin-bottom: 4px; }
.stat-lbl { font-size: 12px; color: #475569; font-weight: 600; }

/* Client header */
.client-header {
    background: linear-gradient(135deg, #3b37cc, #7c3aed);
    border-radius: 14px; padding: 1.1rem 1.5rem; margin-bottom: 1.2rem;
    box-shadow: 0 4px 14px rgba(59,55,204,.25);
}
.client-name { font-size: 20px; font-weight: 800; color: #ffffff; }
.client-sub  { font-size: 13px; color: rgba(255,255,255,.78); margin-top: 4px; }

/* AI result — dark text on white */
.ai-result {
    background: #ffffff; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 1.25rem; margin-top: 1rem;
    white-space: pre-wrap; font-size: 13px; line-height: 1.8;
    color: #0f172a; max-height: 600px; overflow-y: auto;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}

/* Category header */
.cat-header {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: .08em; color: #4f46e5;
    margin: 14px 0 5px 0; padding: 5px 8px;
    background: #e0e7ff; border-radius: 4px;
}

/* Detail text under checks — dark readable */
.check-detail {
    font-size: 12px; color: #374151; margin-top: 4px;
    padding-left: 4px;
}

/* Buttons */
.stButton > button {
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 12px !important;
}

/* Tab text */
[data-baseweb="tab"] { color: #1e293b !important; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Bookkeeping Reviewer")
    st.markdown("*Senior CPA review tool*")
    st.divider()

    uploaded_file = st.file_uploader(
        "Upload bookkeeping file (.xlsx or .xlsm)",
        type=["xlsx", "xlsm"],
        help="Upload the completed bookkeeping Excel file prepared by your staff."
    )

    st.divider()
    import os
    api_key = ""
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    if not api_key:
        api_key = os.environ.get("OPENAI_API_KEY", "")
    if api_key:
        st.success("✅ API key loaded automatically")
    else:
        api_key = st.text_input(
            "OpenAI API key",
            type="password",
            placeholder="sk-proj-...",
            help="Or store as OPENAI_API_KEY in Streamlit Secrets for auto-load"
        )
        if api_key:
            st.success("API key set ✓")

    st.divider()
    st.markdown("**How to use:**")
    st.markdown("""
1. Upload the completed Excel file
2. View automatic checks instantly
3. Enter your OpenAI API key
4. Click any AI review button
5. Download PDF or Word report
""")
    st.divider()
    st.markdown("*All files processed in memory — not stored.*", unsafe_allow_html=True)


# ── Firm name (edit this to your firm name) ────────────────────────────────────
FIRM_NAME    = "Tanupriya Prasad Professional Corporation"
FIRM_TAGLINE = "Bookkeeping & Tax Review Portal  ·  Powered by AI"

# ── Main area ──────────────────────────────────────────────────────────────────
if not uploaded_file:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;background:#f5f7fa">
        <div style="font-size:60px;margin-bottom:1rem">📋</div>
        <div style="font-size:24px;font-weight:700;color:#1e293b;margin-bottom:8px">Bookkeeping File Reviewer</div>
        <div style="font-size:15px;color:#64748b;margin-bottom:2rem">
            Upload a completed bookkeeping Excel file to begin your CPA review
        </div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;max-width:700px;margin:0 auto">
            <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)">
                <div style="font-size:24px;margin-bottom:8px">⚡</div>
                <div style="font-weight:600;color:#1e293b;margin-bottom:4px">25 Auto Checks</div>
                <div style="font-size:12px;color:#64748b">Math, balancing & consistency checks in seconds</div>
            </div>
            <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)">
                <div style="font-size:24px;margin-bottom:8px">🤖</div>
                <div style="font-weight:600;color:#1e293b;margin-bottom:4px">8 AI Prompts</div>
                <div style="font-size:12px;color:#64748b">Tax planning, missing expenses, staff queries & more</div>
            </div>
            <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)">
                <div style="font-size:24px;margin-bottom:8px">📄</div>
                <div style="font-weight:600;color:#1e293b;margin-bottom:4px">PDF & Word</div>
                <div style="font-size:12px;color:#64748b">Professional formatted reports ready for your file</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ── Load file ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Reading Excel file...")
def load_data(file_bytes, filename):
    from excel_reader import read_workbook, extract_data
    wb = read_workbook(BytesIO(file_bytes))
    return extract_data(wb)


@st.cache_data(show_spinner="Running 25 automatic checks...")
def load_checks(file_bytes, filename):
    from excel_reader import read_workbook, extract_data
    from auto_checks import run_checks
    wb = read_workbook(BytesIO(file_bytes))
    data = extract_data(wb)
    return run_checks(data)


file_bytes = uploaded_file.read()
data   = load_data(file_bytes, uploaded_file.name)
checks = load_checks(file_bytes, uploaded_file.name)

from auto_checks import summarize_checks
counts = summarize_checks(checks)

# ── Welcome banner ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="welcome-banner">
    <div>
        <div class="welcome-text">🏢 &nbsp;Welcome to {FIRM_NAME}</div>
        <div class="welcome-sub">{FIRM_TAGLINE}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Client header ──────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="client-header">
    <div class="client-name">📋 {data.get('client_name','')}</div>
    <div class="client-sub">
        Year ended {str(data.get('year_end',''))[:10]} &nbsp;·&nbsp;
        Prepared by {data.get('prepared_by','')} &nbsp;·&nbsp;
        Signer: {data.get('signer','')} &nbsp;·&nbsp;
        {data.get('version','')}
    </div>
</div>
""", unsafe_allow_html=True)

# ── Summary stats ──────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
def fmt(v): return f"${v:,.0f}"

with c1:
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#ef4444">{counts["critical"]}</div><div class="stat-lbl">Critical</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#f59e0b">{counts["warning"]}</div><div class="stat-lbl">Warnings</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#6366f1">{counts["info"]}</div><div class="stat-lbl">Info</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#10b981">{counts["pass"]}</div><div class="stat-lbl">Passed</div></div>', unsafe_allow_html=True)
with c5:
    rev = data.get("total_revenue_cy", 0)
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#fff;font-size:20px">{fmt(rev)}</div><div class="stat-lbl">Revenue CY</div></div>', unsafe_allow_html=True)
with c6:
    ni = data.get("net_income_cy", 0)
    color = "#10b981" if ni >= 0 else "#ef4444"
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:{color};font-size:20px">{fmt(ni)}</div><div class="stat-lbl">Net Income CY</div></div>', unsafe_allow_html=True)

st.divider()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["⚡ Automatic Checks", "🤖 AI Review", "📄 Download Report", "📧 Email Report"])


# ── Tab 1: Automatic Checks ────────────────────────────────────────────────────
with tab1:
    st.markdown("#### 25 Automatic Checks — instant, no AI required")

    # Filter
    col_a, col_b = st.columns([3, 1])
    with col_b:
        show_pass = st.checkbox("Show passed checks", value=False)

    # Group by category
    from itertools import groupby
    cats = {}
    for c in checks:
        cat = c.get("category", "General")
        cats.setdefault(cat, []).append(c)

    level_order = {"critical": 0, "warning": 1, "info": 2, "pass": 3}
    badge_map   = {"critical": "badge-critical", "warning": "badge-warning",
                   "info": "badge-info", "pass": "badge-pass"}
    card_map    = {"critical": "check-critical", "warning": "check-warning",
                   "info": "check-info", "pass": "check-pass"}
    label_map   = {"critical": "CRITICAL", "warning": "WARNING", "info": "INFO", "pass": "PASS"}

    for cat, cat_checks in cats.items():
        visible = [c for c in cat_checks if show_pass or c["level"] != "pass"]
        if not visible: continue

        st.markdown(f'<div class="cat-header">▸ {cat}</div>', unsafe_allow_html=True)
        for c in sorted(visible, key=lambda x: level_order.get(x["level"], 9)):
            detail_html = f'<div class="check-detail">{c.get("detail","")}</div>' if c.get("detail") and c["level"] != "pass" else ""
            st.markdown(f"""
            <div class="check-card {card_map.get(c['level'], '')}">
                <span class="check-badge {badge_map.get(c['level'], '')}">{label_map.get(c['level'], c['level'].upper())}</span>
                <strong style="font-size:13px">{c['title']}</strong>
                {detail_html}
            </div>
            """, unsafe_allow_html=True)


# ── Tab 2: AI Review ───────────────────────────────────────────────────────────
with tab2:
    if not api_key:
        st.warning("⚠️ Enter your OpenAI API key in the sidebar to use AI review features.")
        st.info("💡 Don't have one? Get it at platform.openai.com → API Keys. Cost: ~$0.02–0.05 per review.")
        st.markdown("---")
        st.markdown("**Once your API key is entered, 8 CPA review prompts will appear here:**")
        for label in ["🔍 Full File Review", "💡 Tax Planning", "🔎 Missing Expenses", "📋 Staff Queries",
                      "🔍 Full File Review", "🏦 Bank Statement Review", "🔎 Missing Expenses", "🚩 Flag Unusual Items", "📊 Management Report", "✉️ Client Summary", "💡 Tax Planning", "📁 Engagement Notes", "📋 Staff Queries"]:
            st.markdown(f"&nbsp;&nbsp;&nbsp;▸ {label}", unsafe_allow_html=True)
    else:
        from ai_review import run_prompt, get_prompt_labels

        st.markdown("#### 9 Pre-built CPA Review Prompts")
        st.markdown("*Click any button to run that analysis. Results appear below and are included in your report.*")

        # Prompt buttons — 4 per row
        prompt_labels = get_prompt_labels()
        if "ai_results" not in st.session_state:
            st.session_state.ai_results = {}
        if "active_prompt" not in st.session_state:
            st.session_state.active_prompt = None

        # Dynamic grid: 3 columns, as many rows as needed
        n_cols = 3
        n_prompts = len(prompt_labels)
        n_rows = (n_prompts + n_cols - 1) // n_cols
        all_cols = []
        for _r in range(n_rows):
            all_cols += st.columns(n_cols)

        for i, (key, label) in enumerate(prompt_labels):
            with all_cols[i]:
                already_run = key in st.session_state.ai_results
                btn_label = f"✓ {label}" if already_run else label
                if st.button(btn_label, key=f"btn_{key}", use_container_width=True):
                    st.session_state.active_prompt = key
                    with st.spinner(f"Running: {label}..."):
                        try:
                            result = run_prompt(key, data, checks, api_key)
                            st.session_state.ai_results[key] = result
                        except Exception as e:
                            st.error(f"API error: {e}")

        st.divider()

        # Run all button
        col_run, col_clear = st.columns([2, 1])
        with col_run:
            if st.button("⚡ Engage AI Power — Run All 9 Prompts", type="primary", use_container_width=True):
                progress = st.progress(0)
                status   = st.empty()
                for i, (key, label) in enumerate(prompt_labels):
                    status.info(f"Running {i+1}/9: {label}...")
                    try:
                        result = run_prompt(key, data, checks, api_key)
                        st.session_state.ai_results[key] = result
                    except Exception as e:
                        st.session_state.ai_results[key] = f"Error: {e}"
                    progress.progress((i + 1) / len(prompt_labels))
                status.success("✅ All 9 prompts complete! Go to Download & Email tab.")
        with col_clear:
            if st.button("Clear results", use_container_width=True):
                st.session_state.ai_results = {}
                st.session_state.active_prompt = None
                st.rerun()

        # Display active result
        active = st.session_state.active_prompt
        if active and active in st.session_state.ai_results:
            label = dict(prompt_labels).get(active, active)
            st.markdown(f"#### {label}")
            with st.container():
                st.markdown(
                    f'<div class="ai-result">' +
                    st.session_state.ai_results[active].replace("\n", "<br>") +
                    '</div>',
                    unsafe_allow_html=True
                )
            col_copy, _ = st.columns([1, 3])
            with col_copy:
                st.download_button(
                    "⬇ Save this analysis as .txt",
                    data=st.session_state.ai_results[active],
                    file_name=f"{active}_analysis.txt",
                    mime="text/plain",
                    use_container_width=True,
                    key=f"dl_{active}"
                )
        elif st.session_state.ai_results:
            # Show most recently run
            last_key = list(st.session_state.ai_results.keys())[-1]
            label = dict(prompt_labels).get(last_key, last_key)
            st.markdown(f"#### {label}")
            with st.container():
                st.markdown(
                    f'<div class="ai-result">' +
                    st.session_state.ai_results[last_key].replace("\n", "<br>") +
                    '</div>',
                    unsafe_allow_html=True
                )

        # Financial data expander
        with st.expander("🔍 View extracted financial data sent to AI"):
            from ai_review import build_context
            st.code(build_context(data, checks), language="text")


# ── Tab 3: Download ────────────────────────────────────────────────────────────
with tab3:
    st.markdown("#### Download Review Report")

    ai_results = st.session_state.get("ai_results", {})
    has_ai = len(ai_results) > 0

    if not has_ai:
        st.info("Run at least one AI prompt in the AI Review tab to include AI analysis in your report. "
                "You can also download a report with just the automatic checks.")

    st.markdown(f"**Report will include:**")
    st.markdown(f"- Client details and summary stats")
    st.markdown(f"- All 25 automatic checks ({counts['critical']} critical, {counts['warning']} warnings, {counts['info']} info, {counts['pass']} passed)")
    if has_ai:
        st.markdown(f"- {len(ai_results)} AI analysis section(s): {', '.join(ai_results.keys())}")
    else:
        st.markdown("- *(No AI sections yet — run prompts in AI Review tab)*")

    st.divider()

    col_pdf, col_word = st.columns(2)

    with col_pdf:
        st.markdown("##### 📕 PDF Report")
        st.markdown("Formatted with colour-coded checks, firm branding header and page numbers.")
        if st.button("Generate PDF", type="primary", use_container_width=True):
            with st.spinner("Generating PDF..."):
                try:
                    from report_gen import generate_pdf
                    pdf_bytes = generate_pdf(data, checks, ai_results)
                    client_safe = data.get("client_name","client").replace(" ","_").replace(".","")
                    yr = str(data.get("year_end",""))[:7].replace("-","")
                    filename = f"Review_{client_safe}_{yr}.pdf"
                    st.download_button(
                        "⬇ Download PDF",
                        data=pdf_bytes,
                        file_name=filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
                    st.success("PDF ready!")
                except Exception as e:
                    st.error(f"PDF error: {e}")

    with col_word:
        st.markdown("##### 📘 Word Document")
        st.markdown("Editable .docx — add your letterhead, notes, and customize before printing.")
        if st.button("Generate Word", type="primary", use_container_width=True):
            with st.spinner("Generating Word document..."):
                try:
                    from report_gen import generate_word
                    word_bytes = generate_word(data, checks, ai_results)
                    client_safe = data.get("client_name","client").replace(" ","_").replace(".","")
                    yr = str(data.get("year_end",""))[:7].replace("-","")
                    filename = f"Review_{client_safe}_{yr}.docx"
                    st.download_button(
                        "⬇ Download Word",
                        data=word_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success("Word document ready!")
                except Exception as e:
                    st.error(f"Word error: {e}")

    st.divider()
    st.markdown("*Files are generated in memory and not stored anywhere.*")


# ── Tab 4: Email Report ────────────────────────────────────────────────────────
with tab4:
    st.markdown("#### 📧 Email Review Report")
    st.markdown("Generate a report and send it directly to any email address.")
    st.divider()

    ai_results_em = st.session_state.get("ai_results", {})

    # ── Format selector ───────────────────────────────────────────────────
    st.markdown("**Step 1 — Choose report format**")
    col_fmt1, col_fmt2, col_fmt3 = st.columns(3)
    with col_fmt1:
        send_pdf  = st.checkbox("📕 Attach PDF report",  value=True)
    with col_fmt2:
        send_word = st.checkbox("📘 Attach Word document", value=False)
    with col_fmt3:
        send_inline = st.checkbox("📝 Include summary in email body", value=True)

    st.divider()

    # ── Auto-load SMTP from Streamlit secrets ─────────────────────────────
    def _secret(key, default=""):
        try:
            return st.secrets[key]
        except Exception:
            return default

    _smtp_host_def = _secret("SMTP_HOST", "smtp.gmail.com")
    _smtp_port_def = int(_secret("SMTP_PORT", "587"))
    _smtp_user_def = _secret("SMTP_USER", "")
    _smtp_pass_def = _secret("SMTP_PASSWORD", "")
    _smtp_from_def = _secret("SMTP_FROM", "")

    _secrets_loaded = bool(_smtp_user_def and _smtp_pass_def)
    if _secrets_loaded:
        st.success("✅ Email credentials loaded from Streamlit Secrets — no setup needed")

    # ── Email details ─────────────────────────────────────────────────────
    st.markdown("**Step 2 — Enter email details**")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        to_email = st.text_input(
            "To (recipient email)",
            placeholder="partner@yourfirm.com",
            help="Who should receive this report?"
        )
    with col_e2:
        from_email = st.text_input(
            "From (your email)",
            value=_smtp_from_def,
            placeholder="you@yourfirm.com",
            help="Auto-loaded from Streamlit Secrets if set"
        )

    col_e3, col_e4 = st.columns(2)
    with col_e3:
        client_nm = data.get("client_name", "Client")
        yr_e = str(data.get("year_end", ""))[:10]
        subject = st.text_input(
            "Subject",
            value=f"Bookkeeping Review Report — {client_nm} — Year ended {yr_e}"
        )
    with col_e4:
        cc_email = st.text_input(
            "CC (optional)",
            placeholder="manager@yourfirm.com"
        )

    st.divider()

    # ── SMTP settings ─────────────────────────────────────────────────────
    expander_label = (
        "✅ SMTP settings loaded from Streamlit Secrets — click to override"
        if _secrets_loaded else
        "⚙️ Configure email server (required)"
    )
    with st.expander(expander_label, expanded=not _secrets_loaded):
        if _secrets_loaded:
            st.info(
                f"Using: **{_smtp_user_def}** via **{_smtp_host_def}:{_smtp_port_def}** "
                f"— loaded from Streamlit Secrets automatically. "
                f"Expand this panel only if you want to override."
            )
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            smtp_host = st.text_input("SMTP Host",
                value=_smtp_host_def,
                help="Gmail: smtp.gmail.com  |  Outlook: smtp.office365.com")
            smtp_user = st.text_input("SMTP Username / Email",
                value=_smtp_user_def,
                placeholder="you@gmail.com")
        with col_s2:
            smtp_port = st.number_input("SMTP Port", value=_smtp_port_def, step=1)
            smtp_pass = st.text_input("SMTP Password / App Password",
                value=_smtp_pass_def,
                type="password",
                help="Auto-loaded from Streamlit Secrets if set")

        if not _secrets_loaded:
            st.info("""
**Save permanently:** Go to share.streamlit.io → your app → ⋯ → Settings → Secrets and add:

```
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = "587"
SMTP_USER     = "you@gmail.com"
SMTP_PASSWORD = "your-16-char-app-password"
SMTP_FROM     = "you@gmail.com"
```

**Gmail App Password:** myaccount.google.com → Security → App Passwords → Mail → Generate
**Outlook:** smtp.office365.com, port 587, your regular password
        """)

    # ── Preview & send ────────────────────────────────────────────────────
    st.divider()
    st.markdown("**Step 4 — Preview & send**")

    # Build inline summary
    from auto_checks import summarize_checks
    counts_em = summarize_checks(checks)

    def build_email_body():
        lines = []
        lines.append(f"BOOKKEEPING REVIEW REPORT")
        lines.append(f"{'='*50}")
        lines.append(f"Client:      {data.get('client_name','')}")
        lines.append(f"Year ended:  {yr_e}")
        lines.append(f"Prepared by: {data.get('prepared_by','')}")
        lines.append(f"Reviewer:    {data.get('signer','')}")
        lines.append(f"Generated:   {datetime.now().strftime('%B %d, %Y %H:%M')}")
        lines.append("")
        lines.append("REVIEW SUMMARY")
        lines.append("-" * 40)
        lines.append(f"  Critical issues : {counts_em['critical']}")
        lines.append(f"  Warnings        : {counts_em['warning']}")
        lines.append(f"  Info items      : {counts_em['info']}")
        lines.append(f"  Passed checks   : {counts_em['pass']}")
        lines.append("")
        lines.append("KEY FINANCIALS")
        lines.append("-" * 40)
        lines.append(f"  Revenue CY   : ${data.get('total_revenue_cy',0):,.0f}  |  PY: ${data.get('total_revenue_py',0):,.0f}")
        lines.append(f"  Expenses CY  : ${data.get('total_expenses_cy',0):,.0f}  |  PY: ${data.get('total_expenses_py',0):,.0f}")
        lines.append(f"  Net Income CY: ${data.get('net_income_cy',0):,.0f}  |  PY: ${data.get('net_income_py',0):,.0f}")
        lines.append(f"  Total Assets : ${data.get('total_assets_cy',0):,.0f}")
        if ai_results_em:
            lines.append("")
            lines.append("AI REVIEW SECTIONS INCLUDED")
            lines.append("-" * 40)
            PTITLES = {
                "full_review":"Full File Review","bank_statement_review":"Bank Statement Audit",
                "missing_expenses":"Missing Expenses","unusual_items":"Flag Unusual Items",
                "management_summary":"Management Report","client_summary":"Client Summary",
                "tax_planning":"Tax Planning","engagement_notes":"Engagement Notes",
                "staff_queries":"Staff Queries"
            }
            for k in ai_results_em:
                lines.append(f"  • {PTITLES.get(k, k)}")
        lines.append("")
        lines.append("This report was generated by the Bookkeeping Reviewer system.")
        lines.append("Please review the attached document for full details.")
        return "\n".join(lines)

    email_body_preview = build_email_body()

    if send_inline:
        with st.expander("📋 Preview email body", expanded=False):
            st.code(email_body_preview, language="text")

    col_send, col_test = st.columns([2, 1])

    with col_send:
        send_clicked = st.button(
            "📧 Send Report Email",
            type="primary",
            use_container_width=True,
            disabled=not (to_email and from_email)
        )

    with col_test:
        if st.button("Send test to myself", use_container_width=True,
                     disabled=not from_email):
            to_email = from_email  # override to = from for test

    if send_clicked or (not (to_email and from_email) and st.session_state.get("_send_attempted")):
        st.session_state["_send_attempted"] = True
        if not to_email:
            st.error("Please enter a recipient email address.")
        elif not from_email:
            st.error("Please enter your from email address.")
        elif not locals().get("smtp_host") or not locals().get("smtp_user") or not locals().get("smtp_pass"):
            st.warning("⚠️ Configure your SMTP settings above before sending.")
        else:
            with st.spinner("Generating report and sending email..."):
                try:
                    import smtplib
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.text import MIMEText
                    from email.mime.base import MIMEBase
                    from email import encoders

                    msg = MIMEMultipart()
                    msg["From"]    = from_email
                    msg["To"]      = to_email
                    msg["Subject"] = subject
                    if cc_email:
                        msg["Cc"] = cc_email

                    # Body
                    body_text = email_body_preview if send_inline else (
                        f"Please find the bookkeeping review report for {client_nm} attached.\n\n"
                        f"Year ended: {yr_e}\n"
                        f"Generated: {datetime.now().strftime('%B %d, %Y')}"
                    )
                    msg.attach(MIMEText(body_text, "plain"))

                    client_safe = data.get("client_name","client").replace(" ","_").replace(".","")
                    yr_fn = str(data.get("year_end",""))[:7].replace("-","")

                    # PDF attachment
                    if send_pdf:
                        from report_gen import generate_pdf
                        pdf_bytes = generate_pdf(data, checks, ai_results_em)
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(pdf_bytes)
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition",
                            f"attachment; filename=Review_{client_safe}_{yr_fn}.pdf")
                        msg.attach(part)

                    # Word attachment
                    if send_word:
                        from report_gen import generate_word
                        word_bytes = generate_word(data, checks, ai_results_em)
                        part2 = MIMEBase("application", "octet-stream")
                        part2.set_payload(word_bytes)
                        encoders.encode_base64(part2)
                        part2.add_header("Content-Disposition",
                            f"attachment; filename=Review_{client_safe}_{yr_fn}.docx")
                        msg.attach(part2)

                    # Send
                    recipients = [to_email]
                    if cc_email:
                        recipients.append(cc_email)

                    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
                        server.ehlo()
                        server.starttls()
                        server.login(smtp_user, smtp_pass)
                        server.sendmail(from_email, recipients, msg.as_string())

                    attachments = []
                    if send_pdf:  attachments.append("PDF")
                    if send_word: attachments.append("Word")
                    attach_str = " + ".join(attachments) if attachments else "body only"
                    st.success(f"✅ Email sent successfully to **{to_email}** ({attach_str})")

                except smtplib.SMTPAuthenticationError:
                    st.error("❌ Authentication failed. Check your email/password. Gmail users: use an App Password.")
                except smtplib.SMTPException as e:
                    st.error(f"❌ SMTP error: {e}")
                except Exception as e:
                    st.error(f"❌ Error: {e}")

    st.divider()
    st.markdown("""
**Quick setup guide:**
- **Gmail:** Enable 2-step verification → Google Account → Security → App Passwords → Mail → copy 16-char password
- **Outlook:** Use smtp.office365.com, port 587, your regular email and password
- **Other:** Ask your IT team for SMTP relay credentials

*Email credentials are never stored — entered only for the current session.*
""")
