"""
Bookkeeping File Reviewer
Streamlit + ChatGPT — Senior CPA review tool for Alberta bookkeeping files.
"""
import streamlit as st
from io import BytesIO

st.set_page_config(
    page_title="Bookkeeping Reviewer",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.main { background: #0f0e1a; }
.stApp { background: #0f0e1a; }

/* Check cards */
.check-card {
    border-radius: 10px; padding: 10px 14px; margin: 6px 0;
    border-left: 4px solid; font-size: 13px;
}
.check-critical { background: rgba(239,68,68,.1); border-color: #ef4444; }
.check-warning  { background: rgba(245,158,11,.1); border-color: #f59e0b; }
.check-info     { background: rgba(99,102,241,.1); border-color: #6366f1; }
.check-pass     { background: rgba(16,185,129,.08); border-color: #10b981; }

.check-badge {
    display: inline-block; padding: 1px 8px; border-radius: 4px;
    font-size: 10px; font-weight: 700; margin-right: 8px;
}
.badge-critical { background: #ef4444; color: white; }
.badge-warning  { background: #f59e0b; color: white; }
.badge-info     { background: #6366f1; color: white; }
.badge-pass     { background: #10b981; color: white; }

/* Stat boxes */
.stat-box {
    background: #16152a; border: 1px solid rgba(255,255,255,.1);
    border-radius: 12px; padding: 16px; text-align: center;
}
.stat-num { font-size: 28px; font-weight: 800; margin-bottom: 4px; }
.stat-lbl { font-size: 12px; color: rgba(226,232,240,.5); }

/* Client header */
.client-header {
    background: linear-gradient(135deg, rgba(99,102,241,.2), rgba(139,92,246,.15));
    border: 1px solid rgba(99,102,241,.3); border-radius: 16px;
    padding: 1.2rem 1.5rem; margin-bottom: 1.2rem;
}
.client-name { font-size: 20px; font-weight: 800; color: #fff; }
.client-sub  { font-size: 13px; color: rgba(226,232,240,.6); margin-top: 4px; }

/* AI result */
.ai-result {
    background: #16152a; border: 1px solid rgba(99,102,241,.25);
    border-radius: 12px; padding: 1.25rem; margin-top: 1rem;
    white-space: pre-wrap; font-size: 13px; line-height: 1.7;
    color: #e2e8f0; max-height: 600px; overflow-y: auto;
}

/* Category header */
.cat-header {
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: .07em; color: rgba(165,180,252,.8);
    margin: 12px 0 4px 0; padding: 4px 0;
    border-bottom: 1px solid rgba(99,102,241,.2);
}
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
    api_key = st.text_input(
        "OpenAI API key",
        type="password",
        placeholder="sk-...",
        help="Your ChatGPT API key. Get one at platform.openai.com"
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


# ── Main area ──────────────────────────────────────────────────────────────────
if not uploaded_file:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem">
        <div style="font-size:60px;margin-bottom:1rem">📋</div>
        <div style="font-size:24px;font-weight:700;color:#fff;margin-bottom:8px">Bookkeeping File Reviewer</div>
        <div style="font-size:15px;color:rgba(226,232,240,.5);margin-bottom:2rem">
            Upload a completed bookkeeping Excel file to begin your CPA review
        </div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;max-width:700px;margin:0 auto">
            <div style="background:#16152a;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px">
                <div style="font-size:24px;margin-bottom:8px">⚡</div>
                <div style="font-weight:600;color:#fff;margin-bottom:4px">25 Auto Checks</div>
                <div style="font-size:12px;color:rgba(226,232,240,.4)">Math, balancing & consistency checks in seconds</div>
            </div>
            <div style="background:#16152a;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px">
                <div style="font-size:24px;margin-bottom:8px">🤖</div>
                <div style="font-weight:600;color:#fff;margin-bottom:4px">8 AI Prompts</div>
                <div style="font-size:12px;color:rgba(226,232,240,.4)">Tax planning, missing expenses, staff queries & more</div>
            </div>
            <div style="background:#16152a;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px">
                <div style="font-size:24px;margin-bottom:8px">📄</div>
                <div style="font-weight:600;color:#fff;margin-bottom:4px">PDF & Word</div>
                <div style="font-size:12px;color:rgba(226,232,240,.4)">Professional formatted reports ready for your file</div>
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
tab1, tab2, tab3 = st.tabs(["⚡ Automatic Checks", "🤖 AI Review", "📄 Download Report"])


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
            detail_html = f'<div style="font-size:12px;color:rgba(226,232,240,.55);margin-top:4px">{c.get("detail","")}</div>' if c.get("detail") and c["level"] != "pass" else ""
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
        st.warning("Enter your OpenAI API key in the sidebar to use AI review features.")
        st.info("Don't have one? Get it at platform.openai.com → API keys. Cost: ~$0.02–0.05 per review.")
        st.stop()

    from ai_review import run_prompt, get_prompt_labels

    st.markdown("#### 8 Pre-built CPA Review Prompts")
    st.markdown("*Click any button to run that analysis. Results appear below and are included in your download.*")

    # Prompt buttons — 4 per row
    prompt_labels = get_prompt_labels()
    if "ai_results" not in st.session_state:
        st.session_state.ai_results = {}
    if "active_prompt" not in st.session_state:
        st.session_state.active_prompt = None

    row1 = st.columns(4)
    row2 = st.columns(4)
    all_cols = row1 + row2

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
        if st.button("▶ Run ALL 8 prompts", type="primary", use_container_width=True):
            progress = st.progress(0)
            status   = st.empty()
            for i, (key, label) in enumerate(prompt_labels):
                status.info(f"Running {i+1}/8: {label}...")
                try:
                    result = run_prompt(key, data, checks, api_key)
                    st.session_state.ai_results[key] = result
                except Exception as e:
                    st.session_state.ai_results[key] = f"Error: {e}"
                progress.progress((i + 1) / len(prompt_labels))
            status.success("All 8 prompts complete! Go to Download tab.")
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
        st.markdown(f'<div class="ai-result">{st.session_state.ai_results[active]}</div>', unsafe_allow_html=True)
    elif st.session_state.ai_results:
        # Show most recently run
        last_key = list(st.session_state.ai_results.keys())[-1]
        label = dict(prompt_labels).get(last_key, last_key)
        st.markdown(f"#### {label}")
        st.markdown(f'<div class="ai-result">{st.session_state.ai_results[last_key]}</div>', unsafe_allow_html=True)

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
