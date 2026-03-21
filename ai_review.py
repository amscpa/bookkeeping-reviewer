"""
ai_review.py — 8 pre-built ChatGPT prompts for bookkeeping file review.
"""
from openai import OpenAI


def build_context(data, checks):
    """Build a structured financial context string for ChatGPT."""
    exp = data.get("expense_items", {})
    inc = data.get("income_items", {})

    def fmt(v): return f"${v:,.0f}"
    def chg(cy, py): return f"{((cy-py)/abs(py)*100):+.1f}%" if py else "N/A"

    income_lines = "\n".join(
        f"  {k}: CY {fmt(v['cy'])} | PY {fmt(v['py'])} | Change {chg(v['cy'],v['py'])}"
        for k, v in inc.items() if v['cy'] or v['py']
    )
    expense_lines = "\n".join(
        f"  {k}: CY {fmt(v['cy'])} | PY {fmt(v['py'])} | Change {chg(v['cy'],v['py'])}"
        for k, v in exp.items() if v['cy'] or v['py']
    )
    check_lines = "\n".join(
        f"  [{c['level'].upper()}] {c['title']}: {c['detail']}"
        for c in checks if c['level'] != 'pass'
    )
    sh_lines = "\n".join(
        f"  {str(t.get('date',''))[:10]} | {t.get('description','')} | {t.get('amount','?')}"
        for t in data.get("sh_transactions", [])[:20]
    )
    existing_q = "\n".join(data.get("existing_queries", []))

    return f"""
CLIENT: {data.get('client_name')}
YEAR-END: {str(data.get('year_end',''))[:10]}
CURRENT YEAR (CY): {int(data.get('cy',0))} | PRIOR YEAR (PY): {int(data.get('py',0))}
PREPARED BY: {data.get('prepared_by')} | SIGNER: {data.get('signer')}
FIRST YEAR: {data.get('first_year')} | NTR: {data.get('ntr')}

INCOME STATEMENT:
Revenue Items:
{income_lines}
  TOTAL REVENUE: CY {fmt(data.get('total_revenue_cy',0))} | PY {fmt(data.get('total_revenue_py',0))}

Expense Items:
{expense_lines}
  TOTAL EXPENSES: CY {fmt(data.get('total_expenses_cy',0))} | PY {fmt(data.get('total_expenses_py',0))}

  Income before tax: CY {fmt(data.get('income_before_tax_cy',0))} | PY {fmt(data.get('income_before_tax_py',0))}
  Tax provision: CY {fmt(data.get('tax_provision_cy',0))} | PY {fmt(data.get('tax_provision_py',0))}
  NET INCOME: CY {fmt(data.get('net_income_cy',0))} | PY {fmt(data.get('net_income_py',0))}

BALANCE SHEET:
  Bank: CY {fmt(data.get('bank_cy',0))} | PY {fmt(data.get('bank_py',0))}
  Total Assets: CY {fmt(data.get('total_assets_cy',0))} | PY {fmt(data.get('total_assets_py',0))}
  GST Payable (BS): CY {fmt(data.get('gst_payable_bs_cy',0))} | PY {fmt(data.get('gst_payable_bs_py',0))}
  Taxes Payable: CY {fmt(data.get('taxes_payable_cy',0))} | PY {fmt(data.get('taxes_payable_py',0))}
  Shareholder Loan: CY {fmt(data.get('sh_loan_cy',0))} | PY {fmt(data.get('sh_loan_py',0))}
  Retained Earnings: CY {fmt(data.get('retained_earnings_cy',0))} | PY {fmt(data.get('retained_earnings_py',0))}

RETAINED EARNINGS:
  Opening: {fmt(data.get('re_opening',0))}
  Net Income: {fmt(data.get('re_net_income',0))}
  Dividends: {fmt(data.get('re_dividends',0))}
  Closing: {fmt(data.get('re_closing',0))}

GST:
  Sales: {fmt(data.get('gst_sales',0))}
  Output GST: {fmt(data.get('gst_output',0))}
  Input GST (ITC): {fmt(data.get('gst_input',0))}
  GST Payable (calc): {fmt(data.get('gst_payable_calc',0))}

AMORTIZATION:
  IS expense: {fmt(data.get('expense_items',{}).get('Amortization',{}).get('cy',0))}
  Schedule total: {fmt(data.get('amort_expense_schedule',0))}
  Closing WDV: {fmt(data.get('amort_closing_wdv',0))}

SHAREHOLDER TRANSACTIONS (first 20):
{sh_lines}

AUTOMATIC CHECKS — ISSUES FOUND:
{check_lines if check_lines else '  No issues found by automatic checks.'}

EXISTING QUERIES IN FILE:
{existing_q if existing_q else '  None'}
""".strip()


SYSTEM_PROMPT = """You are a senior Canadian CPA (CGA/CPA designation) reviewing a completed 
bookkeeping and financial statement preparation file for an Alberta small business corporation. 
You have 20+ years of experience in Canadian tax and accounting. 
Be specific, practical, and professional. Reference CRA rules where relevant. 
All dollar amounts are in CAD. Format your response in clear sections with headers.
Do not pad your response — be concise and actionable."""


def run_prompt(prompt_type, data, checks, api_key):
    """Run a specific pre-built prompt. Returns response text."""
    client = OpenAI(api_key=api_key)
    context = build_context(data, checks)

    prompts = {

        "full_review": {
            "label": "🔍 Full File Review",
            "user": f"""Review this completed bookkeeping file thoroughly. 
Based on the financial data below, provide:
1. EXECUTIVE SUMMARY — 3-4 sentences on the overall quality of the file and any urgent concerns
2. CRITICAL ISSUES — anything that must be fixed before this file can go to the client
3. REVIEW QUERIES — specific questions for the staff member who prepared this file
4. ITEMS TO VERIFY WITH CLIENT — what needs client confirmation before finalizing
5. OVERALL ASSESSMENT — is this file ready to review with the client? Rate: Ready / Minor corrections needed / Major corrections needed

Financial data:
{context}"""
        },

        "tax_planning": {
            "label": "💡 Tax Planning Opportunities",
            "user": f"""Based on this Alberta corporation's financial statements, identify all tax planning opportunities. 
Consider:
1. SALARY VS DIVIDEND OPTIMIZATION — given net income of ${data.get('net_income_cy',0):,.0f}, 
   what is the optimal mix? Consider both corporate and personal tax.
2. RRSP PLANNING — based on salary paid, what is the RRSP room being created?
3. SMALL BUSINESS DEDUCTION — is the company eligible? Any risk of association?
4. INCOME SPLITTING — any opportunities for family members?
5. CAPITAL DIVIDEND ACCOUNT — any capital gains or life insurance proceeds?
6. CORPORATE CLASS — any assets that should be held personally vs in corporation?
7. TAX INSTALMENTS — based on current year tax, what should next year's instalments be?
8. OTHER OPPORTUNITIES specific to this file.

For each opportunity, state the estimated tax saving in dollars if possible.

Financial data:
{context}"""
        },

        "missing_expenses": {
            "label": "🔎 Missing Expenses Analysis",
            "user": f"""Analyze this file for potentially missing or understated expenses.
I am the reviewing CPA. The file was prepared by staff from client-provided bank statements and receipts.

1. MISSING EXPENSE ANALYSIS — Compare CY vs PY. For each expense that dropped significantly or is zero 
   in CY but had a balance in PY, assess: Is this likely missing (client forgot to provide) or genuinely not incurred?
   
2. UNDERCLAIMED EXPENSES — Based on the nature of the business and revenue level, what expenses would 
   you expect to see that are not in this file? (e.g., home office, vehicle, professional development)
   
3. POTENTIAL MISCODINGS — Based on the description in transactions, flag any amounts that look 
   like they may have been coded to the wrong account by staff.
   
4. RECOMMENDED CLIENT QUESTIONS — List the specific questions I should ask the client to recover 
   any missing deductions.

Financial data:
{context}"""
        },

        "staff_queries": {
            "label": "📋 Staff Queries & Corrections",
            "user": f"""I am the reviewing CPA. This file was prepared by {data.get('prepared_by','my staff')}.
Generate a formal list of review queries and required corrections for the staff member.

Format as a numbered query list. For each query include:
- QUERY NUMBER
- SHEET/AREA affected
- SPECIFIC ISSUE
- WHAT ACTION IS REQUIRED
- PRIORITY (Urgent / Normal / Low)

Focus on:
1. Technical accounting errors that must be corrected
2. Missing entries or calculations that need to be completed  
3. Supporting documentation that needs to be obtained or verified
4. Consistency issues between sheets
5. Any unusual items that need staff explanation

Be specific — reference actual numbers from the file.

Financial data:
{context}"""
        },

        "management_summary": {
            "label": "📊 Management Report (Internal)",
            "user": f"""Prepare a brief internal management report for this client file for my accounting practice records.
This is NOT for the client — it is for my internal file documentation.

Include:
1. FILE SUMMARY — client, year-end, engagement type, staff, key metrics
2. FINANCIAL HIGHLIGHTS — revenue, net income, key balance sheet items vs PY
3. NOTABLE ITEMS — anything unusual or noteworthy about this year
4. TAX POSITION — estimated tax payable, effective rate, installment recommendation
5. GST COMPLIANCE — status and any issues
6. FILE STATUS — outstanding items before the file can be finalized
7. BILLING NOTE — any scope changes or extra work performed vs standard engagement

Keep this concise — max 1 page equivalent.

Financial data:
{context}"""
        },

        "client_summary": {
            "label": "✉️ Client Executive Summary",
            "user": f"""Draft a professional executive summary letter to present to the client at the year-end meeting.
This should be written in plain English — the client is a business owner, not an accountant.

Include:
1. HIGHLIGHTS of the year (revenue, profit, key changes vs last year)
2. TAX PAYABLE for the year and how it compares to last year
3. CASH POSITION and what changed
4. 2-3 OBSERVATIONS about the business performance
5. 1-2 RECOMMENDATIONS for next year
6. ANY OUTSTANDING ITEMS the client needs to address

Tone: Professional but accessible. No accounting jargon without explanation.
Length: 1 page.
Format: Letter format, addressed to the client.
Do NOT include our firm's address or letterhead — that will be added separately.

Financial data:
{context}"""
        },

        "engagement_notes": {
            "label": "📁 Engagement File Notes",
            "user": f"""Prepare formal engagement file notes for this bookkeeping/financial statement file.
These are professional CPA working paper notes.

Include:
1. ENGAGEMENT OVERVIEW — scope, basis of accounting, standards applied
2. SIGNIFICANT ACCOUNTING POLICIES applied in this file
3. AREAS OF JUDGMENT — where professional judgment was applied
4. SUBSEQUENT EVENTS — any items to consider (based on year-end date)
5. GOING CONCERN — any indicators noted
6. RELATED PARTY TRANSACTIONS — shareholder transactions summary
7. CONTINGENCIES — any items noted
8. FILE COMPLETION CHECKLIST — items confirmed complete vs outstanding

Use professional CPA language. Reference ASPE where applicable.
These notes should be ready to include in the working paper file.

Financial data:
{context}"""
        },

        "unusual_items": {
            "label": "🚩 Flag Unusual Items",
            "user": f"""Review this file and flag all unusual, suspicious, or high-risk items that a senior CPA should investigate before finalizing.

I want you to think like a CRA auditor — what would raise red flags?

Categories to assess:
1. REVENUE FLAGS — unusual revenue patterns, timing, recognition issues
2. EXPENSE FLAGS — unusual amounts, categories, or timing
3. RELATED PARTY FLAGS — shareholder transactions, related company transactions
4. TAX FLAGS — aggressive positions, incomplete provisions, installment issues
5. GST FLAGS — under-reporting risks, ITC issues, filing compliance
6. DOCUMENTATION FLAGS — items that likely need supporting documentation
7. POTENTIAL CRA AUDIT TRIGGERS — items that historically attract CRA attention

For each flag, state:
- The specific item and dollar amount
- Why it is concerning
- What action should be taken
- Risk level (High / Medium / Low)

Be thorough — this is a professional liability review.

Financial data:
{context}"""
        },

        "bank_statement_review": {
            "label": "🏦 Bank Statement Review",
            "user": f"""You are a senior bookkeeping auditor, tax reviewer, and financial statement reviewer.

Your task is to audit the bookkeeping data extracted from this workbook and determine whether it is accurate, complete, and logically consistent. Focus especially on the Bank Statement data.

Workbook structure and rules:
- Rows represent transactions; columns represent accounts.
- Opening balances are at the start of the period; closing balances at the end.
- Row 603 (if present) indicates whether each account is an Income Statement or Balance Sheet account.
- Positive values represent debits; negative values represent credits.
- Range F605:L623 (if present) contains GST calculation, tax provision entries, and final tax liability.

Your objectives:
1. Review the full Bank Statement data and determine whether the bookkeeping appears accurate.
2. Detect bookkeeping errors including:
   - Entries posted to the wrong account
   - Unbalanced or illogical journal patterns
   - Incorrect debit/credit signs
   - Missing postings or duplicate transactions
   - Inconsistent descriptions versus account coding
   - Opening and closing balance issues
   - Accounts classified incorrectly between Income Statement and Balance Sheet
   - GST calculation issues
   - Tax provision or tax liability errors
3. Check whether transaction behaviour is consistent with normal bookkeeping logic.
4. Check whether totals, balances, and account movement appear reasonable.
5. Pay special attention to:
   - Whether opening balances reconcile logically to transaction activity
   - Whether closing balances are consistent with postings
   - Whether Income Statement accounts behave as periodic accounts
   - Whether Balance Sheet accounts behave as cumulative accounts
   - Whether GST entries are correctly applied
   - Whether tax provision and final tax liability are reasonable

Classify each finding as one of:
- Confirmed error
- Likely error
- Possible issue
- Review note

Provide your answer in these sections:

A. OVERALL CONCLUSION
State whether the bookkeeping appears accurate, mostly accurate with exceptions, or materially incorrect.

B. KEY ERRORS FOUND
For each issue: classification, row/account involved, why it appears incorrect, suggested correction.

C. BALANCE AND LOGIC CHECKS
Explain whether opening balances, closing balances, and account classifications appear consistent.
Note unusual patterns or accounts requiring manual review.

D. GST AND TAX REVIEW
Review GST calculation, tax provision entries, and final tax liability.
State any inconsistencies, calculation concerns, or presentation issues.

E. TOP 10 HIGHEST-RISK ITEMS FOR MANUAL REVIEW
Rank the top 10 rows, accounts, or tax items that most need human review.

F. ITEMS REQUIRING HUMAN CONFIRMATION
List any items that may be valid but need client or bookkeeper confirmation.

Be skeptical and analytical. Do not assume the workbook is correct.
Use exact row and column references wherever possible.
If evidence is insufficient to prove an error, label it as "Possible issue".

Financial and transaction data extracted from this file:
{context}"""
        },
    }

    if prompt_type not in prompts:
        return "Unknown prompt type."

    p = prompts[prompt_type]
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": p["user"]}
        ],
        max_tokens=2000,
        temperature=0.3
    )
    return response.choices[0].message.content


def get_prompt_labels():
    """Return list of (key, label) for all prompts."""
    return [
        ("full_review",           "🔍 Full File Review"),
        ("bank_statement_review", "🏦 Bank Statement Review"),
        ("missing_expenses",      "🔎 Missing Expenses"),
        ("unusual_items",         "🚩 Flag Unusual Items"),
        ("management_summary",    "📊 Management Report"),
        ("client_summary",        "✉️ Client Summary"),
        ("tax_planning",          "💡 Tax Planning"),
        ("engagement_notes",      "📁 Engagement Notes"),
        ("staff_queries",         "📋 Staff Queries"),
    ]
