"""
auto_checks.py — 25 automatic checks on the bookkeeping file.
Pure Python / math — no AI required. Fast and reliable.
"""


def pct_change(cy, py):
    """Return % change from PY to CY."""
    if not py: return None
    return (cy - py) / abs(py) * 100


def fmt_dollar(v):
    return f"${v:,.0f}"


def fmt_pct(v):
    return f"{v:+.1f}%"


def run_checks(data):
    """
    Run all automatic checks. Returns list of check dicts:
    {level: 'critical'|'warning'|'info'|'pass', title, detail, category}
    """
    checks = []

    def add(level, title, detail, category="General"):
        checks.append({"level": level, "title": title, "detail": detail, "category": category})

    # ── 1. Engagement setup ──────────────────────────────────────────────────
    if not data.get("client_name") or data["client_name"] == "Unknown":
        add("critical", "Client name missing", "Input sheet B2 is blank.", "Setup")
    else:
        add("pass", "Client name present", data["client_name"], "Setup")

    if not data.get("year_end"):
        add("critical", "Year-end date missing", "Input sheet B1 is blank.", "Setup")
    else:
        add("pass", "Year-end date present", str(data["year_end"])[:10], "Setup")

    if not data.get("prepared_by"):
        add("warning", "Prepared By is blank", "Input sheet B8 should have the staff name.", "Setup")
    else:
        add("pass", "Prepared by filled", data["prepared_by"], "Setup")

    if not data.get("signer"):
        add("warning", "Signer name is blank", "Input sheet B7 — who is signing the NTR?", "Setup")

    # ── 2. Balance Sheet balances ────────────────────────────────────────────
    total_assets = data.get("total_assets_cy", 0)
    total_liab   = data.get("total_liabilities_cy", 0)
    total_equity = data.get("total_equity_cy", 0)
    l_plus_e     = total_liab + total_equity
    diff         = abs(total_assets - l_plus_e)

    if diff > 5:
        add("critical",
            f"Balance sheet out of balance by {fmt_dollar(diff)}",
            f"Total Assets: {fmt_dollar(total_assets)} | "
            f"Liabilities + Equity: {fmt_dollar(l_plus_e)} | "
            f"Difference: {fmt_dollar(diff)}",
            "Balance Sheet")
    else:
        add("pass", "Balance sheet balances",
            f"Assets = Liabilities + Equity = {fmt_dollar(total_assets)}", "Balance Sheet")

    # ── 3. Retained earnings reconciliation ─────────────────────────────────
    re_open  = data.get("re_opening", 0)
    re_ni    = data.get("re_net_income", 0)
    re_div   = data.get("re_dividends", 0)
    re_close = data.get("re_closing", 0)
    re_calc  = re_open + re_ni + re_div  # dividends stored as negative
    re_diff  = abs(re_calc - re_close)

    if re_diff > 5:
        add("critical",
            f"Retained earnings don't reconcile — difference {fmt_dollar(re_diff)}",
            f"Opening {fmt_dollar(re_open)} + NI {fmt_dollar(re_ni)} + Dividends {fmt_dollar(re_div)} "
            f"= {fmt_dollar(re_calc)} but RE closing = {fmt_dollar(re_close)}",
            "Retained Earnings")
    else:
        add("pass", "Retained earnings reconcile",
            f"Opening {fmt_dollar(re_open)} → Closing {fmt_dollar(re_close)}", "Retained Earnings")

    # ── 4. Net income ties between IS and RE ────────────────────────────────
    ni_is = data.get("net_income_cy", 0)
    ni_re = data.get("re_net_income", 0)
    if abs(ni_is - ni_re) > 5:
        add("critical",
            f"Net income mismatch between IS and RE",
            f"IS net income: {fmt_dollar(ni_is)} | RE net income: {fmt_dollar(ni_re)} | "
            f"Difference: {fmt_dollar(abs(ni_is-ni_re))}",
            "Net Income")
    elif ni_is:
        add("pass", "Net income ties between IS and RE",
            f"Both show {fmt_dollar(ni_is)}", "Net Income")

    # ── 5. RE closing = BS retained earnings ────────────────────────────────
    re_bs = data.get("retained_earnings_cy", 0)
    if abs(re_close - re_bs) > 5:
        add("critical",
            "RE closing balance doesn't match BS Retained Earnings",
            f"RE statement closing: {fmt_dollar(re_close)} | BS Retained Earnings: {fmt_dollar(re_bs)}",
            "Retained Earnings")
    elif re_close:
        add("pass", "RE closing matches BS", f"{fmt_dollar(re_close)}", "Retained Earnings")

    # ── 6. Amortization expense ties to schedule ────────────────────────────
    amort_is   = data.get("expense_items", {}).get("Amortization", {}).get("cy", 0)
    amort_sched = data.get("amort_expense_schedule", 0)
    if amort_is and amort_sched:
        if abs(amort_is - amort_sched) > 5:
            add("warning",
                f"Amortization mismatch — IS vs schedule",
                f"IS shows {fmt_dollar(amort_is)} | Schedule shows {fmt_dollar(amort_sched)} | "
                f"Difference: {fmt_dollar(abs(amort_is - amort_sched))}",
                "Amortization")
        else:
            add("pass", "Amortization ties to schedule",
                f"Both show ~{fmt_dollar(amort_is)}", "Amortization")

    # ── 7. GST payable — calc vs BS ─────────────────────────────────────────
    gst_calc = data.get("gst_payable_calc", 0)
    gst_bs   = data.get("gst_payable_bs_cy", 0)
    if gst_calc and gst_bs:
        if abs(gst_calc - gst_bs) > 10:
            add("warning",
                f"GST payable mismatch — GST sheet vs BS",
                f"GST sheet calculated: {fmt_dollar(gst_calc)} | BS GST payable: {fmt_dollar(gst_bs)} | "
                f"Difference: {fmt_dollar(abs(gst_calc-gst_bs))}",
                "GST")
        else:
            add("pass", "GST payable ties (GST sheet vs BS)",
                f"~{fmt_dollar(gst_bs)}", "GST")

    # ── 8. GST rate reasonableness (Alberta = 5%) ───────────────────────────
    gst_output = data.get("gst_output", 0)
    gst_sales  = data.get("gst_sales", 0)
    if gst_sales and gst_output:
        implied_rate = gst_output / gst_sales * 100
        if implied_rate < 3 or implied_rate > 8:
            add("warning",
                f"GST output rate looks unusual: {implied_rate:.1f}% of sales",
                f"Expected ~5% for Alberta. Output GST {fmt_dollar(gst_output)} on sales {fmt_dollar(gst_sales)}. "
                f"Check if all taxable supplies are captured.",
                "GST")
        else:
            add("pass", "GST output rate reasonable",
                f"{implied_rate:.1f}% of sales — within expected 5% Alberta range", "GST")

    # ── 9. Revenue change YoY ────────────────────────────────────────────────
    rev_cy = data.get("total_revenue_cy", 0)
    rev_py = data.get("total_revenue_py", 0)
    if rev_py:
        rev_chg = pct_change(rev_cy, rev_py)
        if rev_chg is not None:
            if abs(rev_chg) > 30:
                level = "critical" if abs(rev_chg) > 50 else "warning"
                add(level,
                    f"Revenue changed significantly: {fmt_pct(rev_chg)} YoY",
                    f"CY revenue: {fmt_dollar(rev_cy)} | PY revenue: {fmt_dollar(rev_py)}. "
                    f"Requires explanation — client situation change, lost contract, or coding issue?",
                    "Revenue")
            else:
                add("pass", f"Revenue change normal: {fmt_pct(rev_chg)} YoY",
                    f"CY {fmt_dollar(rev_cy)} vs PY {fmt_dollar(rev_py)}", "Revenue")

    # ── 10. Missing expenses — PY had balance, CY is zero ───────────────────
    missing_expenses = []
    exp_items = data.get("expense_items", {})
    for name, vals in exp_items.items():
        py_val = vals.get("py", 0)
        cy_val = vals.get("cy", 0)
        # PY had >$500, CY is 0
        if py_val > 500 and cy_val == 0:
            missing_expenses.append(f"{name}: PY had {fmt_dollar(py_val)}, CY = $0")

    if missing_expenses:
        add("warning",
            f"{len(missing_expenses)} expense(s) present in PY but zero in CY",
            "Missing: " + " | ".join(missing_expenses[:6]) +
            (" + more..." if len(missing_expenses) > 6 else ""),
            "Expenses")
    else:
        add("pass", "No significant missing expenses vs PY",
            "All PY expense categories accounted for in CY", "Expenses")

    # ── 11. Large expense swings (>100% increase) ───────────────────────────
    large_swings = []
    for name, vals in exp_items.items():
        py_val = vals.get("py", 0)
        cy_val = vals.get("cy", 0)
        if py_val > 200 and cy_val > 0:
            chg = pct_change(cy_val, py_val)
            if chg is not None and chg > 100:
                large_swings.append(f"{name}: {fmt_pct(chg)} (PY {fmt_dollar(py_val)} → CY {fmt_dollar(cy_val)})")

    if large_swings:
        add("warning",
            f"{len(large_swings)} expense(s) more than doubled vs PY",
            " | ".join(large_swings[:5]),
            "Expenses")

    # ── 12. Large expense drops (>50% decrease from PY) ─────────────────────
    large_drops = []
    for name, vals in exp_items.items():
        py_val = vals.get("py", 0)
        cy_val = vals.get("cy", 0)
        if py_val > 500 and cy_val > 0:
            chg = pct_change(cy_val, py_val)
            if chg is not None and chg < -50:
                large_drops.append(f"{name}: {fmt_pct(chg)} (PY {fmt_dollar(py_val)} → CY {fmt_dollar(cy_val)})")

    if large_drops:
        add("info",
            f"{len(large_drops)} expense(s) dropped >50% vs PY",
            " | ".join(large_drops[:5]) + " — confirm intentional.",
            "Expenses")

    # ── 13. Meals & entertainment vs revenue ─────────────────────────────────
    meals_cy = exp_items.get("Meals and entertainment", {}).get("cy", 0)
    if rev_cy and meals_cy:
        meals_pct = meals_cy / rev_cy * 100
        if meals_pct > 5:
            add("warning",
                f"Meals & entertainment is {meals_pct:.1f}% of revenue — CRA may scrutinize",
                f"{fmt_dollar(meals_cy)} on revenue of {fmt_dollar(rev_cy)}. "
                f"CRA expects 50% deductibility and adequate documentation.",
                "Expenses")
        else:
            add("pass", f"Meals & entertainment reasonable: {meals_pct:.1f}% of revenue",
                fmt_dollar(meals_cy), "Expenses")

    # ── 14. Travel vs revenue ────────────────────────────────────────────────
    travel_cy = exp_items.get("Travel", {}).get("cy", 0)
    if rev_cy and travel_cy:
        travel_pct = travel_cy / rev_cy * 100
        if travel_pct > 10:
            add("warning",
                f"Travel expense is {travel_pct:.1f}% of revenue — needs documentation",
                f"{fmt_dollar(travel_cy)} — confirm client has receipts and business purpose.",
                "Expenses")

    # ── 15. Salaries — payroll tax check ─────────────────────────────────────
    salaries_cy = exp_items.get("Salaries", {}).get("cy", 0)
    benefits_cy = exp_items.get("Employee benefits", {}).get("cy", 0)
    payroll_tax = data.get("bs_items", {}).get("Wages and deductions payable", {}).get("cy", 0)
    if salaries_cy > 5000 and not payroll_tax and not benefits_cy:
        add("info",
            "Salaries recorded but no payroll deductions payable on BS",
            f"Salaries = {fmt_dollar(salaries_cy)}. Verify CPP/EI/Income Tax deductions are accounted for. "
            f"Cross-check with T4 summary.",
            "Payroll")
    elif salaries_cy > 5000:
        add("pass", "Salaries present with supporting accounts",
            fmt_dollar(salaries_cy), "Payroll")

    # ── 16. Dividends consistency ────────────────────────────────────────────
    div_re   = abs(data.get("re_dividends", 0))
    div_sh   = data.get("sh_total_outflows", 0)
    div_is   = exp_items.get("Dividend", {}).get("cy", 0)
    if div_re and abs(div_sh) and abs(abs(div_re) - abs(div_sh)) > 100:
        add("warning",
            "Dividends in RE don't match shareholder outflows",
            f"RE dividends: {fmt_dollar(div_re)} | Shareholder outflows: {fmt_dollar(abs(div_sh))}. "
            f"Verify dividend declaration date and amount.",
            "Dividends")
    elif div_re:
        add("pass", f"Dividends recorded: {fmt_dollar(div_re)}", "", "Dividends")

    # ── 17. Shareholder loan — blank amounts ─────────────────────────────────
    blank_sh = data.get("sh_blank_amounts", 0)
    if blank_sh:
        add("warning",
            f"{blank_sh} shareholder transaction(s) have no amount",
            "Open ShareholderTrans sheet — some rows have descriptions but no $ amount. "
            "These may be uncoded transactions.",
            "Shareholder")
    else:
        add("pass", "All shareholder transactions have amounts", "", "Shareholder")

    # ── 18. Shareholder loan balance ─────────────────────────────────────────
    sh_bal_cy = data.get("sh_loan_cy", 0)
    sh_bal_py = data.get("sh_loan_py", 0)
    if sh_bal_cy and abs(sh_bal_cy) > 10000:
        add("info",
            f"Shareholder loan balance: {fmt_dollar(sh_bal_cy)}",
            f"PY was {fmt_dollar(sh_bal_py)}. If this is a loan TO the corporation from the shareholder, "
            f"confirm interest is being charged at CRA prescribed rate. "
            f"If loan FROM corporation, verify section 15(2) rules.",
            "Shareholder")

    # ── 19. Tax provision reasonableness ─────────────────────────────────────
    ibt = data.get("income_before_tax_cy", 0)
    tax = data.get("tax_provision_cy", 0)
    if ibt and tax:
        eff_rate = tax / ibt * 100
        if eff_rate > 30 or eff_rate < 5:
            add("warning",
                f"Effective tax rate looks unusual: {eff_rate:.1f}%",
                f"Tax provision {fmt_dollar(tax)} on income {fmt_dollar(ibt)}. "
                f"Alberta small business rate ~11%, general rate ~23%. "
                f"Verify tax provision calculation.",
                "Tax")
        else:
            add("pass", f"Effective tax rate reasonable: {eff_rate:.1f}%",
                f"Tax {fmt_dollar(tax)} on income {fmt_dollar(ibt)}", "Tax")
    elif ibt > 0 and not tax:
        add("warning",
            "Income before tax is positive but no tax provision recorded",
            f"Income before tax: {fmt_dollar(ibt)}. Is the income tax provision entry missing?",
            "Tax")

    # ── 20. Taxes payable YoY change ─────────────────────────────────────────
    tax_cy = data.get("taxes_payable_cy", 0)
    tax_py = data.get("taxes_payable_py", 0)
    if tax_py and tax_cy:
        chg = pct_change(tax_cy, tax_py)
        if chg is not None and abs(chg) > 50:
            add("info",
                f"Taxes payable changed {fmt_pct(chg)} YoY",
                f"CY {fmt_dollar(tax_cy)} vs PY {fmt_dollar(tax_py)}. "
                f"Consistent with income change? Review installment schedule.",
                "Tax")

    # ── 21. Bank balance reasonableness ─────────────────────────────────────
    bank_cy = data.get("bank_cy", 0)
    bank_py = data.get("bank_py", 0)
    if bank_cy < 0:
        add("critical", "Bank balance is negative",
            f"Bank account shows {fmt_dollar(bank_cy)}. "
            f"Possible missing deposit, timing issue, or overdraft not disclosed.",
            "Bank")
    elif bank_cy:
        chg = pct_change(bank_cy, bank_py)
        if chg is not None:
            add("pass" if abs(chg) < 100 else "info",
                f"Bank balance: {fmt_dollar(bank_cy)} ({fmt_pct(chg) if chg else ''} vs PY)",
                f"PY was {fmt_dollar(bank_py)}", "Bank")

    # ── 22. Professional fees — recurring ────────────────────────────────────
    prof_cy = exp_items.get("Professional fees", {}).get("cy", 0)
    if not prof_cy:
        add("info",
            "No professional fees recorded",
            "Typical for this client? If accounting fees are being accrued, "
            "confirm the AP/accrual entry was made.",
            "Expenses")

    # ── 23. Existing queries from file ───────────────────────────────────────
    existing = data.get("existing_queries", [])
    if existing:
        add("warning",
            f"{len(existing)} open quer{'y' if len(existing)==1 else 'ies'} in the Queries sheet",
            " | ".join(existing[:5]),
            "Queries")

    # ── 24. First year check ─────────────────────────────────────────────────
    if data.get("first_year") == "Yes":
        add("info",
            "First year of filing — additional review required",
            "Verify: opening balances are nil, share capital entry present, "
            "incorporation date matches, and Articles of Incorporation on file.",
            "Setup")

    # ── 25. NTR report ───────────────────────────────────────────────────────
    if data.get("ntr") == "Yes":
        add("info",
            "Notice to Reader engaged — verify standard NTR wording",
            "Confirm engagement letter on file, NTR dated after year-end, "
            "and no scope limitations noted.",
            "Setup")

    return checks


def summarize_checks(checks):
    """Return counts by level."""
    return {
        "critical": sum(1 for c in checks if c["level"] == "critical"),
        "warning":  sum(1 for c in checks if c["level"] == "warning"),
        "info":     sum(1 for c in checks if c["level"] == "info"),
        "pass":     sum(1 for c in checks if c["level"] == "pass"),
    }
