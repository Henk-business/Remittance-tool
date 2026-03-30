import streamlit as st
import pandas as pd
from io import BytesIO
from reconcile import run_reconciliation, build_report
import datetime

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Remittance Reconciliation",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .main { background: #f8fafc; }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1100px; }

    /* Header */
    .recon-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
        border-radius: 16px;
        padding: 36px 40px;
        color: white;
        margin-bottom: 28px;
    }
    .recon-header h1 { font-size: 28px; font-weight: 700; margin: 0 0 6px; }
    .recon-header p { font-size: 15px; opacity: .8; margin: 0; font-weight: 300; }

    /* Metric cards */
    .metric-row { display: flex; gap: 14px; margin: 20px 0; }
    .metric-card {
        flex: 1;
        border-radius: 12px;
        padding: 18px 20px;
        border: 1px solid;
    }
    .metric-card.green  { background: #f0fdf4; border-color: #bbf7d0; }
    .metric-card.red    { background: #fef2f2; border-color: #fecaca; }
    .metric-card.amber  { background: #fffbeb; border-color: #fde68a; }
    .metric-card.purple { background: #f5f3ff; border-color: #ddd6fe; }
    .metric-card.blue   { background: #eff6ff; border-color: #bfdbfe; }
    .metric-card .label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; opacity: .6; margin-bottom: 6px; }
    .metric-card .value { font-size: 26px; font-weight: 700; }
    .metric-card .sub   { font-size: 12px; opacity: .65; margin-top: 2px; }
    .green  .value { color: #16a34a; }
    .red    .value { color: #dc2626; }
    .amber  .value { color: #d97706; }
    .purple .value { color: #7c3aed; }
    .blue   .value { color: #2563eb; }

    /* Upload boxes */
    .upload-box {
        background: white;
        border: 2px dashed #cbd5e1;
        border-radius: 14px;
        padding: 28px;
        text-align: center;
        transition: all .2s;
        margin-bottom: 16px;
    }
    .upload-box:hover { border-color: #2563eb; background: #eff6ff; }
    .upload-box .icon { font-size: 32px; margin-bottom: 8px; }
    .upload-box h3 { font-size: 15px; font-weight: 600; margin: 0 0 4px; color: #1e293b; }
    .upload-box p  { font-size: 13px; color: #64748b; margin: 0; }

    /* Step badges */
    .step-badge {
        display: inline-flex; align-items: center; justify-content: center;
        width: 28px; height: 28px; border-radius: 50%;
        background: #2563eb; color: white;
        font-size: 13px; font-weight: 700; margin-right: 10px;
    }
    .step-title { font-size: 16px; font-weight: 600; color: #1e293b; display: flex; align-items: center; margin-bottom: 12px; }

    /* Result rows */
    .result-ok   { background: #f0fdf4; border-left: 4px solid #22c55e; }
    .result-warn { background: #fef2f2; border-left: 4px solid #ef4444; }
    .result-info { background: #fffbeb; border-left: 4px solid #f59e0b; }

    /* Download button */
    .stDownloadButton button {
        background: linear-gradient(135deg, #1e3a5f, #2563eb) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        padding: 14px 32px !important;
        width: 100% !important;
        margin-top: 8px;
    }
    .stDownloadButton button:hover { opacity: .9 !important; }

    div[data-testid="stExpander"] { border: 1px solid #e2e8f0; border-radius: 10px; }

    /* Hide streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="recon-header">
  <h1>🔍 Remittance Reconciliation</h1>
  <p>Upload your SAP export and the client's remittance — the tool automatically flags matches, doubles, missing items, and discrepancies.</p>
</div>
""", unsafe_allow_html=True)


# ── STEP 1: FILE UPLOADS ─────────────────────────────────────────────────────
st.markdown('<div class="step-title"><span class="step-badge">1</span> Upload your files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="upload-box">
      <div class="icon">🗂️</div>
      <h3>SAP Export</h3>
      <p>FBL5N / ALV customer open items (.xlsx)</p>
    </div>
    """, unsafe_allow_html=True)
    sap_file = st.file_uploader("SAP Export", type=['xlsx', 'xls'], label_visibility='collapsed', key='sap')

with col2:
    st.markdown("""
    <div class="upload-box">
      <div class="icon">📄</div>
      <h3>Client Remittance</h3>
      <p>Payment advice Excel from the client (.xlsx)</p>
    </div>
    """, unsafe_allow_html=True)
    rem_file = st.file_uploader("Remittance", type=['xlsx', 'xls'], label_visibility='collapsed', key='rem')


# ── STEP 2: DETAILS ──────────────────────────────────────────────────────────
st.markdown('<div class="step-title" style="margin-top:24px"><span class="step-badge">2</span> Enter payment details <span style="font-size:13px;font-weight:400;color:#94a3b8;margin-left:8px">(optional but improves the report)</span></div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    customer_name = st.text_input("Customer name", placeholder="e.g. AB-InBev Belgium NV")
with c2:
    payment_amount = st.number_input("Payment amount (€)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
with c3:
    payment_date = st.date_input("Payment date", value=None)


# ── STEP 3: RUN ──────────────────────────────────────────────────────────────
st.markdown('<div class="step-title" style="margin-top:24px"><span class="step-badge">3</span> Run the reconciliation</div>', unsafe_allow_html=True)

run_col, _ = st.columns([1, 2])
with run_col:
    run = st.button("▶  Run Reconciliation", use_container_width=True, type="primary")

if run:
    if not sap_file:
        st.error("⚠️ Please upload the SAP export file.")
    elif not rem_file:
        st.error("⚠️ Please upload the client remittance file.")
    else:
        with st.spinner("Running reconciliation — this usually takes a few seconds…"):
            try:
                pmt = float(payment_amount) if payment_amount and payment_amount > 0 else None
                pmt_date = str(payment_date) if payment_date else None
                cname = customer_name.strip() or "Customer"

                results = run_reconciliation(
                    BytesIO(sap_file.read()),
                    BytesIO(rem_file.read()),
                    payment_amount=pmt,
                    payment_date=pmt_date,
                )
                report_bytes = build_report(results, cname)

                # ── RESULTS ──────────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### ✅ Reconciliation Complete")

                n_matched  = len(results['matched'])
                n_cleared  = len(results['already_cleared'])
                n_missing  = len(results['not_found'])
                n_diff     = len(results['amount_diff'])
                n_sap_only = len(results['missing_from_remittance'])

                # Metric cards
                st.markdown(f"""
                <div class="metric-row">
                  <div class="metric-card green">
                    <div class="label">✓ Matched</div>
                    <div class="value">{n_matched}</div>
                    <div class="sub">Open in SAP &amp; on remittance</div>
                  </div>
                  <div class="metric-card red">
                    <div class="label">⚠ Already Cleared</div>
                    <div class="value">{n_cleared}</div>
                    <div class="sub">Potential double payments</div>
                  </div>
                  <div class="metric-card purple">
                    <div class="label">✗ Not Found</div>
                    <div class="value">{n_missing}</div>
                    <div class="sub">On remittance, not in SAP</div>
                  </div>
                  <div class="metric-card amber">
                    <div class="label">△ Amount Diff</div>
                    <div class="value">{n_diff}</div>
                    <div class="sub">Amounts don't match</div>
                  </div>
                  <div class="metric-card blue">
                    <div class="label">📋 SAP Only</div>
                    <div class="value">{n_sap_only}</div>
                    <div class="sub">Open in SAP, not on remittance</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                # ── FINDINGS PREVIEW ─────────────────────────────────────────
                if n_cleared > 0:
                    with st.expander(f"⚠️  Already Cleared — {n_cleared} potential double payments", expanded=True):
                        st.warning("These items appear on the remittance but already have a clearing document in SAP. Verify before processing — they may have already been paid.")
                        rows = []
                        for item in results['already_cleared']:
                            rows.append({
                                "Reference": item['ref'],
                                "Invoice # (remittance)": item.get('invoice_num', ''),
                                "Amount (€)": item.get('amount'),
                                "Cleared in SAP on": str(item.get('cleared_date', ''))[:10] if item.get('cleared_date') else '',
                                "SAP Clearing Doc": str(item.get('cleared_by', '')),
                            })
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if n_missing > 0:
                    with st.expander(f"✗  Not Found in SAP — {n_missing} items"):
                        st.info("These references are on the remittance but cannot be found anywhere in the SAP export.")
                        rows = [{"Reference": i['ref'], "Invoice #": i.get('invoice_num',''), "Amount (€)": i.get('amount')} for i in results['not_found']]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if n_diff > 0:
                    with st.expander(f"△  Amount Differences — {n_diff} items"):
                        st.info("Found in SAP but the amounts differ from the remittance by more than €1.")
                        rows = []
                        for item in results['amount_diff']:
                            rows.append({
                                "Reference": item['ref'],
                                "Remittance (€)": item.get('amount'),
                                "SAP Amount (€)": item.get('sap_amount'),
                                "Difference (€)": item.get('difference'),
                            })
                        df_diff = pd.DataFrame(rows)
                        st.dataframe(df_diff, use_container_width=True, hide_index=True)

                if n_matched > 0:
                    with st.expander(f"✓  Matched — {n_matched} items (all good)"):
                        rows = [{"Reference": i['ref'], "Invoice #": i.get('invoice_num',''), "SAP Amount (€)": i.get('sap_amount')} for i in results['matched']]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                # ── DOWNLOAD ─────────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### 📥 Download Full Report")
                st.caption("The Excel report contains all findings with full detail across 6 colour-coded tabs.")

                safe = cname.replace(' ', '_').replace('/', '-')[:30]
                fname = f"Reconciliation_{safe}_{datetime.date.today()}.xlsx"

                st.download_button(
                    label="⬇  Download Excel Report",
                    data=report_bytes.getvalue(),
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                import traceback
                st.error(f"Something went wrong: {e}")
                with st.expander("Technical detail"):
                    st.code(traceback.format_exc())


# ── HOW IT WORKS (always visible at bottom) ──────────────────────────────────
st.markdown("---")
with st.expander("ℹ️  How this tool works"):
    c1, c2, c3, c4 = st.columns(4)
    steps = [
        ("1", "Parse inputs", "Reads the SAP export and scans the remittance for 9954xxxxxx reference numbers"),
        ("2", "Cross-reference", "Matches every remittance reference against open items and cleared items in SAP"),
        ("3", "Flag issues", "Identifies double payments, missing items, amount differences, and open invoices not on the remittance"),
        ("4", "Report", "Downloads a formatted Excel with one tab per finding category"),
    ]
    for col, (num, title, desc) in zip([c1, c2, c3, c4], steps):
        with col:
            st.markdown(f"**{num}. {title}**")
            st.caption(desc)

with st.expander("📋  What SAP export format does this expect?"):
    st.markdown("""
The tool is designed around the standard **FBL5N customer line item** export, saved as **.xlsx**.

It expects these columns (exact names matter):

| Column | What it is |
|---|---|
| `Assignment` | The AB-InBev reference number (9954xxxxxx) |
| `Document Number` | SAP document number |
| `Document Type` | RV, AB, RU, DZ, ZP etc |
| `Document Date` | Invoice date |
| `Net due date` | Payment due date |
| `Amount in local currency` | Invoice amount |
| `Clearing Document` | Blank if open, filled if cleared |
| `Clearing date` | Date it was cleared |
| `Text` | Line item description |
| `Document Header Text` | Order/PO reference |

**Tip:** Go to FBL5N in SAP → run for your customer → click the spreadsheet icon to export → save as .xlsx. That's it.
    """)
