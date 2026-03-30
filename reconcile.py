import pandas as pd
import re
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# ─── COLOURS ────────────────────────────────────────────────────────────────
BG = {
    'dk_blue':  '1F3864', 'md_blue':  '2E75B6', 'lt_blue':  'D6E4F0',
    'dk_green': '1E4620', 'md_green': '375623', 'lt_green': 'E2EFDA',
    'dk_red':   '7B0000', 'md_red':   'C00000', 'lt_red':   'FFE2E2', 'pink': 'FFD7D7',
    'yellow':   'FFF2CC', 'orange':   'FCE4D6', 'amber':    'FFC000',
    'grey':     'F2F2F2', 'mid_grey': 'BFBFBF', 'white':    'FFFFFF',
    'purple':   '4A235A', 'lt_purple':'F5EEF8',
}
FG = {
    'white':'FFFFFF','black':'000000','md_red':'C00000','md_green':'375623',
    'dk_red':'7B0000','md_blue':'2E75B6','grey':'595959','purple':'4A235A',
    'lt_purple':'4A235A','amber':'B8860B',
}

def _c(ws,row,col,val=None,bold=False,bg='white',fg='black',sz=10,ha='left',wrap=False,fmt=None,italic=False):
    cell=ws.cell(row=row,column=col,value=val)
    cell.font=Font(name='Arial',bold=bold,color=FG.get(fg,fg),size=sz,italic=italic)
    cell.fill=PatternFill('solid',fgColor=BG.get(bg,bg))
    cell.alignment=Alignment(horizontal=ha,vertical='center',wrap_text=wrap)
    if fmt: cell.number_format=fmt
    return cell

def _mr(ws,row,c1,c2,val=None,**kw):
    ws.merge_cells(f'{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}')
    _c(ws,row,c1,val,**kw)
    bg=kw.get('bg','white')
    for c in range(c1+1,c2+1):
        ws.cell(row=row,column=c).fill=PatternFill('solid',fgColor=BG.get(bg,'FFFFFF'))

def _hdr(ws,row,cols,bg='md_blue'):
    for col,lbl in cols:
        cell=ws.cell(row=row,column=col,value=lbl)
        cell.font=Font(name='Arial',bold=True,color='FFFFFF',size=9)
        cell.fill=PatternFill('solid',fgColor=BG[bg])
        cell.alignment=Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[row].height=15

def fd(ts):
    try: return pd.Timestamp(ts).strftime('%d/%m/%Y') if pd.notna(ts) else ''
    except: return ''

def _extract_refs(text):
    """Extract all 9954xxxxxx references from any string."""
    return re.findall(r'9954\d{6}', str(text))


# ─── SAP EXPORT PARSER ──────────────────────────────────────────────────────
def parse_sap_export(filepath_or_buffer):
    """
    Reads a SAP FBL5N / ALV customer line-item export.
    Returns a normalised DataFrame.
    SAP is the source of truth for everything — doc type, amount sign,
    and whether an item is open (no clearing doc) or already cleared.
    """
    raw = pd.read_excel(filepath_or_buffer, sheet_name=0, header=0)
    raw.columns = [str(c).strip() for c in raw.columns]

    col_map = {
        'Assignment':               'assignment',
        'Document Number':          'doc_number',
        'Document Type':            'doc_type',
        'Document Date':            'doc_date',
        'Net due date':             'due_date',
        'Amount in local currency': 'amount',
        'Clearing Document':        'clearing_doc',
        'Clearing date':            'clearing_date',
        'Text':                     'text',
        'Document Header Text':     'header_text',
    }
    rename={k:v for k,v in col_map.items() if k in raw.columns}
    df=raw.rename(columns=rename)

    for col in ['doc_date','due_date','clearing_date']:
        if col in df.columns:
            df[col]=pd.to_datetime(df[col],errors='coerce')

    if 'amount' in df.columns:
        df['amount']=pd.to_numeric(df['amount'],errors='coerce').fillna(0)

    df['ref']=df.get('assignment',pd.Series(dtype=str)).astype(str).str.strip()

    # ── SAP-AUTHORITATIVE CLASSIFICATION ──────────────────────────────────
    # An item is an INVOICE if: doc_type=RV and amount > 0
    # An item is a CREDIT NOTE if: doc_type=RV and amount < 0
    #                           OR: doc_type=RU (goods return)
    # AB items are clearing residuals, not true invoices/credits
    # DZ/ZP are payments
    def classify(row):
        dt = str(row.get('doc_type','')).strip()
        amt = row.get('amount', 0)
        if dt == 'RV':
            return 'CREDIT_NOTE' if amt < 0 else 'INVOICE'
        if dt == 'RU':
            return 'CREDIT_NOTE'
        if dt == 'AB':
            return 'CLEARING_RESIDUAL'
        if dt in ('DZ','ZP'):
            return 'PAYMENT'
        return 'OTHER'

    df['sap_classification'] = df.apply(classify, axis=1)
    df['is_open'] = df['clearing_doc'].isna()
    return df


# ─── REMITTANCE PARSER ───────────────────────────────────────────────────────
def parse_remittance(filepath_or_buffer):
    """
    Reads any client remittance/payment advice Excel.

    We deliberately DO NOT trust the client's signs or credit/invoice labels.
    We only extract: the reference number(s) and optionally the invoice number.
    SAP will tell us what each reference actually is.

    Returns list of dicts: {ref, invoice_num, raw_ref, row_index}
    """
    raw = pd.read_excel(filepath_or_buffer, sheet_name=0, header=None)
    items = []
    seen_refs = set()

    # First pass: look for a header row to identify ref/invoice columns
    ref_col = None
    inv_col = None
    data_start = 0

    for hdr_idx in range(min(10, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[hdr_idx]]
        for ci, v in enumerate(row_vals):
            if any(k in v for k in ['referentie','reference','ref']):
                ref_col = ci
            if any(k in v for k in ['factuur','invoice','nummer','number']):
                if ci != ref_col:
                    inv_col = ci
        if ref_col is not None:
            data_start = hdr_idx + 1
            break

    # Extract refs from the identified columns
    for row_idx in range(data_start, len(raw)):
        row = raw.iloc[row_idx]

        # Gather all cells to scan
        cells_to_scan = list(row)

        # Prioritise the ref column if found
        primary_refs = []
        if ref_col is not None and ref_col < len(row):
            primary_refs = _extract_refs(row.iloc[ref_col])

        # Also scan every cell for 9954xxxxxx patterns (catches refs embedded in text)
        all_refs = []
        for cell in cells_to_scan:
            all_refs.extend(_extract_refs(cell))

        # Combine: primary first, then any others found
        combined = primary_refs + [r for r in all_refs if r not in primary_refs]

        # Get invoice number if available (used only for display)
        inv_num = ''
        if inv_col is not None and inv_col < len(row):
            inv_num = str(row.iloc[inv_col]).strip()
            if inv_num.lower() in ('nan','none',''): inv_num = ''

        # If no inv_col, scan for BILL/RBILL patterns
        if not inv_num:
            for cell in cells_to_scan:
                m = re.search(r'[RB]BILL/\d{4}/\d{2}/\d{4}', str(cell))
                if m:
                    inv_num = m.group(0)
                    break

        raw_ref = str(row.iloc[ref_col]).strip() if ref_col is not None and ref_col < len(row) else ''

        for ref in combined:
            if ref not in seen_refs:
                seen_refs.add(ref)
                items.append({
                    'ref': ref,
                    'invoice_num': inv_num,
                    'raw_ref': raw_ref,
                    'row_index': row_idx,
                })

    return items


# ─── MAIN RECONCILIATION ────────────────────────────────────────────────────
def run_reconciliation(sap_file, remittance_file, payment_amount=None, payment_date=None):
    """
    SAP is the source of truth.
    For each reference on the remittance:
      1. Find it in SAP
      2. Use SAP's doc type and amount to determine if it's an invoice or credit note
      3. Check if it's open or already cleared
      4. Flag any issues
    """
    sap = parse_sap_export(sap_file)
    remittance = parse_remittance(remittance_file)

    # ── BUILD SAP LOOKUPS ──────────────────────────────────────────────────
    # Only RV and RU items are real invoices/credits
    # We look at open items first, then cleared
    rv_ru = sap[sap['doc_type'].isin(['RV','RU'])].copy()
    rv_ru_open    = rv_ru[rv_ru['is_open']].copy()
    rv_ru_cleared = rv_ru[~rv_ru['is_open']].copy()

    def make_lookup(df_in):
        lkp = {}
        for _, row in df_in.iterrows():
            ref = row['ref']
            if ref and ref != 'nan':
                lkp.setdefault(ref, []).append(row)
        return lkp

    open_lkp    = make_lookup(rv_ru_open)
    cleared_lkp = make_lookup(rv_ru_cleared)
    # Also keep a lookup of ALL open items (RV+RU+AB) for complete picture
    all_open_lkp = make_lookup(sap[sap['is_open']])

    # ── MATCH EACH REMITTANCE ITEM ─────────────────────────────────────────
    matched_invoices   = []   # open RV invoice — good match
    matched_credits    = []   # open RV/RU credit note — client is including it (legit)
    already_cleared    = []   # has a clearing doc — potential double
    not_found          = []   # nowhere in SAP RV/RU
    rem_refs_seen = set()

    for item in remittance:
        ref = item['ref']
        rem_refs_seen.add(ref)

        if ref in open_lkp:
            sap_rows = open_lkp[ref]
            sap_net  = sum(r['amount'] for r in sap_rows)
            sap_class = sap_rows[0]['sap_classification']  # SAP authoritative
            entry = {
                **item,
                'sap_amount':  sap_net,
                'sap_class':   sap_class,
                'sap_doc_type': sap_rows[0]['doc_type'],
                'sap_due_date': sap_rows[0]['due_date'],
                'sap_doc_date': sap_rows[0]['doc_date'],
                'sap_header':   str(sap_rows[0].get('header_text','')) if pd.notna(sap_rows[0].get('header_text')) else '',
            }
            if sap_class == 'INVOICE':
                matched_invoices.append(entry)
            else:
                matched_credits.append(entry)

        elif ref in cleared_lkp:
            sap_rows = cleared_lkp[ref]
            already_cleared.append({
                **item,
                'sap_amount':   sum(r['amount'] for r in sap_rows),
                'sap_class':    sap_rows[0]['sap_classification'],
                'cleared_by':   str(sap_rows[0].get('clearing_doc','')),
                'cleared_date': sap_rows[0].get('clearing_date'),
            })
        else:
            not_found.append(item)

    # ── SAP OPEN INVOICES NOT ON REMITTANCE ───────────────────────────────
    open_invoices_sap = rv_ru_open[rv_ru_open['sap_classification']=='INVOICE'].copy()
    missing_from_rem  = open_invoices_sap[
        ~open_invoices_sap['ref'].isin(rem_refs_seen)
    ].copy()

    open_credits_sap = rv_ru_open[rv_ru_open['sap_classification']=='CREDIT_NOTE'].copy()

    # ── TOTALS ─────────────────────────────────────────────────────────────
    return {
        'matched_invoices':     matched_invoices,
        'matched_credits':      matched_credits,
        'already_cleared':      already_cleared,
        'not_found':            not_found,
        'missing_from_rem':     missing_from_rem,
        'open_invoices_sap':    open_invoices_sap,
        'open_credits_sap':     open_credits_sap,
        'payment_amount':       payment_amount,
        'payment_date':         payment_date,
        'remittance':           remittance,
        'sap':                  sap,
        # Totals
        't_matched_inv':   sum(i['sap_amount'] for i in matched_invoices),
        't_matched_cred':  sum(i['sap_amount'] for i in matched_credits),
        't_cleared':       len(already_cleared),
        't_not_found':     len(not_found),
        't_missing':       missing_from_rem['amount'].sum() if len(missing_from_rem) else 0,
        't_open_inv':      open_invoices_sap['amount'].sum(),
        't_open_cred':     open_credits_sap['amount'].sum(),
    }


# ─── EXCEL REPORT ───────────────────────────────────────────────────────────
def build_report(results, customer_name='Customer'):
    wb = openpyxl.Workbook()

    mi   = results['matched_invoices']
    mc   = results['matched_credits']
    ac   = results['already_cleared']
    nf   = results['not_found']
    mfr  = results['missing_from_rem']
    pmt  = results['payment_amount']

    # ── SHEET 1: SUMMARY ──────────────────────────────────────────────────
    ws = wb.active; ws.title = 'Summary'
    for ci_,w in enumerate([4,44,4,18,4,18,4,4],1):
        ws.column_dimensions[get_column_letter(ci_)].width=w

    r=1
    title = f'REMITTANCE RECONCILIATION — {customer_name}'
    if pmt: title += f'  ·  Payment €{pmt:,.2f}'
    _mr(ws,r,1,8,title,bold=True,bg='dk_blue',fg='white',sz=14,ha='center')
    ws.row_dimensions[r].height=36; r+=1
    _mr(ws,r,1,8,'SAP is the source of truth for all classifications below. Client signs and labels are not used.',
        bg='md_blue',fg='white',sz=9,ha='center',italic=True)
    ws.row_dimensions[r].height=18; r+=2

    _mr(ws,r,1,8,'RECONCILIATION SUMMARY',bold=True,bg='dk_blue',fg='white',sz=11,ha='center')
    ws.row_dimensions[r].height=24; r+=1

    rows=[
        (f'✓  Invoices matched — open in SAP, on remittance ({len(mi)} items)',    results['t_matched_inv'],  'lt_green','md_green',False),
        (f'✓  Credit notes matched — open in SAP, on remittance ({len(mc)} items)',results['t_matched_cred'], 'lt_green','md_green',False),
        ('',None,'white','black',False),
        (f'⚠  Already cleared in SAP — check for double payment ({len(ac)} items)',None,'lt_red','md_red',False),
        (f'✗  Not found anywhere in SAP ({len(nf)} items)',                         None,'pink','md_red',False),
        ('',None,'white','black',False),
        (f'Open SAP invoices not on remittance ({len(mfr)} items)',                results['t_missing'],  'lt_blue','black',False),
        (f'Open SAP credit notes available to apply',                               results["t_open_cred"],'lt_green','md_green',False),
    ]
    for desc,amt,bg,fg,bold in rows:
        if desc=='': _mr(ws,r,1,8,None,bg='white'); ws.row_dimensions[r].height=6; r+=1; continue
        _mr(ws,r,1,5,desc,bold=bold,bg=bg,fg=fg,sz=10)
        if amt is not None:
            _c(ws,r,6,amt,bold=bold,bg=bg,fg=fg,fmt='#,##0.00',ha='right',sz=10)
        else:
            _c(ws,r,6,None,bg=bg)
        _c(ws,r,7,None,bg=bg); _c(ws,r,8,None,bg=bg)
        ws.row_dimensions[r].height=20; r+=1

    # ── SHEET 2: MATCHED INVOICES ──────────────────────────────────────────
    def make_sheet(wb, tab_name, title_text, subtitle, items, bg_hdr, cols_def, row_fn):
        ws2=wb.create_sheet(tab_name)
        nc=len(cols_def)
        for ci,(lbl,w) in enumerate(cols_def,1):
            ws2.column_dimensions[get_column_letter(ci)].width=w
        r2=1
        _mr(ws2,r2,1,nc,title_text,bold=True,bg=bg_hdr,fg='white',sz=11)
        ws2.row_dimensions[r2].height=24; r2+=1
        _mr(ws2,r2,1,nc,subtitle,bg='lt_'+bg_hdr if 'lt_'+bg_hdr in BG else 'grey',fg='black',sz=9,italic=True,wrap=True)
        ws2.row_dimensions[r2].height=18; r2+=1
        _hdr(ws2,r2,[(i+1,lbl) for i,(lbl,_) in enumerate(cols_def)])
        r2+=1
        total=0
        for idx,item in enumerate(items,1):
            bg='lt_green' if idx%2==0 else 'white'
            row_fn(ws2,r2,idx,item,bg)
            amt=item.get('sap_amount',0) or 0
            total+=amt
            ws2.row_dimensions[r2].height=13; r2+=1
        _mr(ws2,r2,1,nc-1,'TOTAL',bold=True,bg=bg_hdr,fg='white',sz=10)
        _c(ws2,r2,nc,total,bold=True,bg=bg_hdr,fg='white',fmt='#,##0.00',ha='right',sz=10)
        ws2.row_dimensions[r2].height=16
        return ws2

    inv_cols=[('#',4),('AB-InBev Reference (SAP Assignment)',26),('Invoice # (remittance)',22),
              ('SAP Doc Type',12),('Invoice Date',13),('Due Date',13),('SAP Amount (€)',15)]
    def inv_row(ws,r,idx,item,bg):
        _c(ws,r,1,idx,bg=bg,sz=8,ha='center')
        _c(ws,r,2,item['ref'],bg=bg,sz=9,bold=True)
        _c(ws,r,3,item.get('invoice_num',''),bg=bg,sz=9)
        _c(ws,r,4,item.get('sap_doc_type',''),bg=bg,sz=8,ha='center')
        _c(ws,r,5,fd(item.get('sap_doc_date')),bg=bg,sz=9,ha='center')
        _c(ws,r,6,fd(item.get('sap_due_date')),bg=bg,sz=9,ha='center')
        _c(ws,r,7,item.get('sap_amount'),bg=bg,fmt='#,##0.00',ha='right',sz=9)

    make_sheet(wb,'✓ Matched — Invoices',
        f'✓  MATCHED INVOICES — Open in SAP, on remittance  ({len(mi)} items  ·  €{results["t_matched_inv"]:,.2f})',
        'These are confirmed open invoices in SAP that the customer is paying. SAP classification is used — client sign/label is ignored.',
        mi,'md_green',inv_cols,inv_row)

    make_sheet(wb,'✓ Matched — Credits',
        f'✓  MATCHED CREDIT NOTES — Open in SAP, on remittance  ({len(mc)} items  ·  €{results["t_matched_cred"]:,.2f})',
        'SAP classifies these as credit notes (RV negative or RU). The client has included them on the remittance — they will offset against the payment.',
        mc,'md_green',inv_cols,inv_row)

    # ── SHEET: ALREADY CLEARED ──────────────────────────────────────────────
    ws4=wb.create_sheet('⚠ Already Cleared')
    for ci,w in enumerate([4,26,22,15,18,22,4],1):
        ws4.column_dimensions[get_column_letter(ci)].width=w
    r4=1
    _mr(ws4,r4,1,7,f'⚠  ALREADY CLEARED IN SAP — Potential Double Payments  ({len(ac)} items)',
        bold=True,bg='md_red',fg='white',sz=11)
    ws4.row_dimensions[r4].height=24; r4+=1
    _mr(ws4,r4,1,7,'These references appear on the remittance but already have a SAP clearing document. '
        'This could mean the client is paying them twice. Verify before processing.',
        bg='lt_red',fg='dk_red',sz=9,italic=True,wrap=True)
    ws4.row_dimensions[r4].height=20; r4+=1
    _hdr(ws4,r4,[(1,'#'),(2,'Reference'),(3,'Invoice # (remittance)'),(4,'SAP Classification'),(5,'Cleared Date'),(6,'SAP Clearing Doc'),(7,'')])
    r4+=1
    for idx,item in enumerate(ac,1):
        _c(ws4,r4,1,idx,bg='lt_red',sz=8,ha='center')
        _c(ws4,r4,2,item['ref'],bg='lt_red',sz=9,bold=True)
        _c(ws4,r4,3,item.get('invoice_num',''),bg='lt_red',sz=9)
        _c(ws4,r4,4,item.get('sap_class',''),bg='lt_red',sz=9,ha='center')
        _c(ws4,r4,5,fd(item.get('cleared_date')),bg='lt_red',sz=9,ha='center')
        _c(ws4,r4,6,str(item.get('cleared_by','')),bg='lt_red',sz=9)
        _c(ws4,r4,7,None,bg='lt_red')
        ws4.row_dimensions[r4].height=14; r4+=1

    # ── SHEET: NOT FOUND ────────────────────────────────────────────────────
    ws5=wb.create_sheet('✗ Not Found in SAP')
    for ci_,w in enumerate([4,26,30,4,4],1):
        ws5.column_dimensions[get_column_letter(ci_)].width=w
    r5=1
    _mr(ws5,r5,1,5,f'✗  NOT FOUND IN SAP  ({len(nf)} items)',bold=True,bg='purple',fg='white',sz=11)
    ws5.row_dimensions[r5].height=24; r5+=1
    _mr(ws5,r5,1,5,'These references appear on the remittance but cannot be found in SAP as RV or RU documents '
        '(open or cleared). They may not exist, use a different reference, or may not yet be booked.',
        bg='lt_purple',fg='black',sz=9,italic=True,wrap=True)
    ws5.row_dimensions[r5].height=20; r5+=1
    _hdr(ws5,r5,[(1,'#'),(2,'Reference from Remittance'),(3,'Invoice # (remittance)'),(4,''),(5,'')])
    r5+=1
    for idx,item in enumerate(nf,1):
        bg='lt_purple' if idx%2==0 else 'white'
        _c(ws5,r5,1,idx,bg=bg,sz=8,ha='center')
        _c(ws5,r5,2,item['ref'],bg=bg,sz=9)
        _c(ws5,r5,3,item.get('invoice_num',''),bg=bg,sz=9)
        _c(ws5,r5,4,None,bg=bg); _c(ws5,r5,5,None,bg=bg)
        ws5.row_dimensions[r5].height=13; r5+=1

    # ── SHEET: SAP OPEN NOT ON REMITTANCE ───────────────────────────────────
    if len(mfr)>0:
        ws6=wb.create_sheet('SAP Open — Not on Remittance')
        for ci_,w in enumerate([4,26,30,13,13,15],1):
            ws6.column_dimensions[get_column_letter(ci_)].width=w
        r6=1
        _mr(ws6,r6,1,6,f'OPEN IN SAP — NOT ON REMITTANCE  ({len(mfr)} items  ·  €{mfr["amount"].sum():,.2f})',
            bold=True,bg='md_blue',fg='white',sz=11)
        ws6.row_dimensions[r6].height=24; r6+=1
        _mr(ws6,r6,1,6,'These open SAP invoices were not mentioned on the remittance. '
            'The client may have forgotten them or plans a separate payment.',
            bg='lt_blue',fg='md_blue',sz=9,italic=True,wrap=True)
        ws6.row_dimensions[r6].height=18; r6+=1
        _hdr(ws6,r6,[(1,'#'),(2,'SAP Reference'),(3,'Order / Description'),(4,'Invoice Date'),(5,'Due Date'),(6,'Amount (€)')])
        r6+=1
        for idx,(_,row) in enumerate(mfr.sort_values('due_date').iterrows(),1):
            bg='lt_blue' if idx%2==0 else 'white'
            _c(ws6,r6,1,idx,bg=bg,sz=8,ha='center')
            _c(ws6,r6,2,row['ref'],bg=bg,sz=9)
            _c(ws6,r6,3,str(row.get('header_text','')) if pd.notna(row.get('header_text')) else '',bg=bg,sz=9)
            _c(ws6,r6,4,fd(row.get('doc_date')),bg=bg,sz=9,ha='center')
            _c(ws6,r6,5,fd(row.get('due_date')),bg=bg,sz=9,ha='center')
            _c(ws6,r6,6,row['amount'],bg=bg,fmt='#,##0.00',ha='right',sz=9)
            ws6.row_dimensions[r6].height=13; r6+=1
        _mr(ws6,r6,1,5,'TOTAL',bold=True,bg='md_blue',fg='white',sz=10)
        _c(ws6,r6,6,mfr['amount'].sum(),bold=True,bg='md_blue',fg='white',fmt='#,##0.00',ha='right',sz=10)
        ws6.row_dimensions[r6].height=16

    out=BytesIO()
    wb.save(out)
    out.seek(0)
    return out
