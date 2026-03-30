import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import traceback
from reconcile import run_reconciliation, build_report

st.set_page_config(
    page_title="Remittance Reconciliation",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main { background: #f8fafc; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1080px; }

.banner {
    background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
    border-radius: 16px; padding: 36px 40px; color: white; margin-bottom: 28px;
}
.banner h1 { font-size: 28px; font-weight: 700; margin: 0 0 8px; }
.banner p  { font-size: 15px; opacity: .8; margin: 0; font-weight: 300; }

.upload-card {
    background: white; border: 2px dashed #cbd5e1; border-radius: 14px;
    padding: 28px; text-align: center; margin-bottom: 4px;
}
.upload-card .icon { font-size: 30px; margin-bottom: 8px; }
.upload-card h3    { font-size: 14px; font-weight: 600; margin: 0 0 4px; color: #1e293b; }
.upload-card p     { font-size: 12px; color: #64748b; margin: 0; }

.metric-row { display: flex; gap: 12px; margin: 16px 0; }
.metric     { flex: 1; border-radius: 12px; padding: 16px 18px; border: 1px solid; }
.metric .lbl { font-size: 11px; font-weight: 600; text-transform: uppercase;
               letter-spacing: .06em; opacity: .6; margin-bottom: 4px; }
.metric .val { font-size: 26px; font-weight: 700; }
.metric .sub { font-size: 11px; opacity: .6; margin-top: 2px; }
.green  { background:#f0fdf4; border-color:#bbf7d0; } .green  .val { color:#16a34a; }
.red    { background:#fef2f2; border-color:#fecaca; } .red    .val { color:#dc2626; }
.amber  { background:#fffbeb; border-color:#fde68a; } .amber  .val { color:#d97706; }
.purple { background:#f5f3ff; border-color:#ddd6fe; } .purple .val { color:#7c3aed; }
.blue   { background:#eff6ff; border-color:#bfdbfe; } .blue   .val { color:#2563eb; }

.stDownloadButton button {
    background: linear-gradient(135deg, #1e3a5f, #2563eb) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
    font-size: 15px !important; padding: 14px 28px !important;
    width: 100% !important;
}
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── BANNER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="banner">
  <h1>🔍 Remittance Reconciliation</h1>
  <p>Upload a SAP customer open items export and any client remittance file.
     The tool matches references, classifies invoices vs credit notes using SAP as the
     source of truth, and flags any discrepancies.</p>
</div>
""", unsafe_allow_html=True)

# ── STEP 1: FILES ─────────────────────────────────────────────────────────────
st.markdown("### 1 · Upload files")
c1, c2 = st.columns(2)

with c1:
    st.markdown("""<div class="upload-card">
      <div class="icon">🗂️</div><h3>SAP Export</h3>
      <p>FBL5N or any ALV customer line-item export saved as .xlsx</p>
    </div>""", unsafe_allow_html=True)
    sap_file = st.file_uploader("SAP", type=['xlsx','xls'],
                                label_visibility='collapsed', key='sap')

with c2:
    st.markdown("""<div class="upload-card">
      <div class="icon">📄</div><h3>Client Remittance</h3>
      <p>Payment advice or remittance Excel from the customer (.xlsx)</p>
    </div>""", unsafe_allow_html=True)
    rem_file = st.file_uploader("Remittance", type=['xlsx','xls'],
                                label_visibility='collapsed', key='rem')

# ── STEP 2: DETAILS ───────────────────────────────────────────────────────────
st.markdown("### 2 · Payment details <span style='font-size:13px;font-weight:400;color:#94a3b8;'>(optional)</span>", unsafe_allow_html=True)
d1, d2, d3 = st.columns(3)
with d1:
    customer_name = st.text_input("Customer name", placeholder="e.g. Acme Corp")
with d2:
    payment_amount = st.number_input("Payment amount (€)", min_value=0.0, value=0.0,
                                     step=0.01, format="%.2f")
with d3:
    payment_date = st.date_input("Payment date", value=None)

# ── STEP 3: RUN ───────────────────────────────────────────────────────────────
st.markdown("### 3 · Run")
btn_col, _ = st.columns([1, 2])
with btn_col:
    run = st.button("▶  Run Reconciliation", use_container_width=True, type="primary")

if run:
    if not sap_file:
        st.error("Please upload the SAP export file.")
    elif not rem_file:
        st.error("Please upload the client remittance file.")
    else:
        with st.spinner("Matching references and checking SAP…"):
            try:
                pmt    = float(payment_amount) if payment_amount and payment_amount > 0 else None
                cname  = customer_name.strip() or "Customer"

                sap_bytes = BytesIO(sap_file.read())
                rem_bytes = BytesIO(rem_file.read())

                results = run_reconciliation(sap_bytes, rem_bytes, pmt, cname)

                mi = results['matched_invoices']
                mc = results['matched_credits']
                ac = results['already_cleared']
                nf = results['not_found']
                mfr = results['missing_from_rem']

                # ── METRICS ──────────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### Results")
                st.markdown(f"""
                <div class="metric-row">
                  <div class="metric green">
                    <div class="lbl">✓ Invoices Matched</div>
                    <div class="val">{len(mi)}</div>
                    <div class="sub">€{results['t_inv']:,.2f}</div>
                  </div>
                  <div class="metric green">
                    <div class="lbl">✓ Credits Matched</div>
                    <div class="val">{len(mc)}</div>
                    <div class="sub">€{results['t_cred']:,.2f}</div>
                  </div>
                  <div class="metric red">
                    <div class="lbl">⚠ Already Cleared</div>
                    <div class="val">{len(ac)}</div>
                    <div class="sub">Potential doubles</div>
                  </div>
                  <div class="metric purple">
                    <div class="lbl">✗ Not Found</div>
                    <div class="val">{len(nf)}</div>
                    <div class="sub">Not in SAP</div>
                  </div>
                  <div class="metric blue">
                    <div class="lbl">📋 SAP Only</div>
                    <div class="val">{len(mfr)}</div>
                    <div class="sub">Not on remittance</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                # ── PREVIEW PANELS ────────────────────────────────────────────
                if ac:
                    with st.expander(f"⚠️  Already Cleared — {len(ac)} items (check for double payments)", expanded=True):
                        st.warning("These references are on the remittance but are already cleared in SAP. Investigate before processing.")
                        rows = [{
                            "SAP Reference":       i['sap_ref'],
                            "SAP Classification":  i.get('sap_class',''),
                            "Cleared Date (SAP)":  str(i.get('cleared_date',''))[:10],
                            "Clearing Doc":        str(i.get('cleared_by','')),
                        } for i in ac]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if nf:
                    with st.expander(f"✗  Not Found in SAP — {len(nf)} items"):
                        st.info("These values were on the remittance but cannot be matched to any RV or RU document in SAP.")
                        rows = [{"Value from Remittance": i['sap_ref'], "Context": i.get('context','')} for i in nf]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if mi:
                    with st.expander(f"✓  Matched Invoices — {len(mi)} items"):
                        rows = [{
                            "SAP Reference":  i['sap_ref'],
                            "SAP Amount (€)": i.get('sap_amount'),
                            "Due Date":       str(i.get('sap_due_date',''))[:10],
                        } for i in mi]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if mc:
                    with st.expander(f"✓  Matched Credit Notes — {len(mc)} items"):
                        rows = [{"SAP Reference": i['sap_ref'], "SAP Amount (€)": i.get('sap_amount')} for i in mc]
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                # ── DOWNLOAD ──────────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### Download full report")
                st.caption("Excel with one tab per category — matched invoices, credits, already cleared, not found, and SAP open items not on remittance.")

                report = build_report(results)
                safe   = cname.replace(' ', '_').replace('/', '-')[:30]
                fname  = f"Reconciliation_{safe}_{datetime.date.today()}.xlsx"

                st.download_button(
                    "⬇  Download Excel Report",
                    data=report.getvalue(),
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                with st.expander("Technical detail"):
                    st.code(traceback.format_exc())

# ── INFO ──────────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("ℹ️  How matching works"):
    st.markdown("""
**The tool does not rely on the customer's formatting.**
It builds a lookup from every unique value in the SAP **Assignment** and **Document Number** columns,
then scans every cell of the remittance to find anything that matches — even if the reference is embedded inside a longer string.

**SAP is the source of truth for classification:**

| SAP condition | Classification |
|---|---|
| Doc type RV, amount > 0 | Invoice |
| Doc type RV, amount < 0 | Credit note |
| Doc type RU | Credit note (goods return) |
| Clearing document present | Already cleared |

The customer's signs, BILL/RBILL labels, and positive/negative conventions are all ignored.
""")

with st.expander("📋  What SAP export format does this expect?"):
    st.markdown("""
Any FBL5N or ALV customer line-item export saved as **.xlsx**. The tool recognises these column names
(and their Dutch/German equivalents):

`Assignment` · `Document Number` · `Document Type` · `Document Date` · `Net due date` ·
`Amount in local currency` · `Clearing Document` · `Clearing date` · `Text` · `Document Header Text`

**Tip:** Run FBL5N for your customer in SAP → click the export/spreadsheet icon → Save as .xlsx.
    """)
