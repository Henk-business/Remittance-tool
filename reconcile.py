"""
Remittance Reconciliation Engine
---------------------------------
General purpose — works for any customer, any SAP export format.

SAP export is the source of truth for:
  - Whether an item is an invoice or credit note (doc type + sign)
  - Whether an item is open or already cleared (clearing doc present or not)

The remittance is scanned for any reference numbers that can be matched
back to the SAP export Assignment field or Document Number field.
We do NOT trust the customer's signs, labels, or formatting.
"""

import pandas as pd
import re
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')


# ── COLOURS ─────────────────────────────────────────────────────────────────
BG = {
    'dk_blue': '1F3864', 'md_blue': '2E75B6', 'lt_blue': 'D6E4F0',
    'md_green': '375623', 'lt_green': 'E2EFDA',
    'md_red': 'C00000',  'lt_red': 'FFE2E2', 'pink': 'FFD7D7',
    'yellow': 'FFF2CC',  'orange': 'FCE4D6',
    'grey': 'F2F2F2',    'mid_grey': 'BFBFBF', 'white': 'FFFFFF',
    'purple': '4A235A',  'lt_purple': 'F5EEF8',
    'amber': 'FFC000',
}
FG = {
    'white': 'FFFFFF', 'black': '000000', 'md_red': 'C00000',
    'md_green': '375623', 'dk_red': '7B0000', 'md_blue': '2E75B6',
    'grey': '595959', 'purple': '4A235A',
}

def _c(ws, row, col, val=None, bold=False, bg='white', fg='black',
       sz=10, ha='left', wrap=False, fmt=None, italic=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name='Arial', bold=bold, color=FG.get(fg, fg), size=sz, italic=italic)
    c.fill = PatternFill('solid', fgColor=BG.get(bg, bg))
    c.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    return c

def _mr(ws, row, c1, c2, val=None, **kw):
    ws.merge_cells(f'{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}')
    _c(ws, row, c1, val, **kw)
    bg = kw.get('bg', 'white')
    for c in range(c1 + 1, c2 + 1):
        ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=BG.get(bg, 'FFFFFF'))

def _hdr(ws, row, cols, bg='md_blue'):
    for col, lbl in cols:
        c = ws.cell(row=row, column=col, value=lbl)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=9)
        c.fill = PatternFill('solid', fgColor=BG[bg])
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 15

def _fd(ts):
    try:
        return pd.Timestamp(ts).strftime('%d/%m/%Y') if pd.notna(ts) else ''
    except Exception:
        return ''

def _col_w(ws, widths):
    """Set column widths by list (1-indexed)."""
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── SAP EXPORT PARSER ────────────────────────────────────────────────────────
# Maps every known SAP column name variant to a standard internal name.
SAP_COL_MAP = {
    # Assignment / reference
    'Assignment':                    'assignment',
    'Zuordnung':                     'assignment',
    'Reference Key 1':               'ref_key1',
    'Referentie 1':                  'ref_key1',
    # Document number
    'Document Number':               'doc_number',
    'Belegnummer':                   'doc_number',
    'Factuurnummer':                 'doc_number',
    # Document type
    'Document Type':                 'doc_type',
    'Boekingssoort':                 'doc_type',
    'Belegtyp':                      'doc_type',
    # Dates
    'Document Date':                 'doc_date',
    'Boekingsdatum':                 'doc_date',
    'Belegdatum':                    'doc_date',
    'Net due date':                  'due_date',
    'Netto-vervaldatum':             'due_date',
    'Nettofälligkeitsdatum':         'due_date',
    # Amount
    'Amount in local currency':      'amount',
    'Bedrag in lokale valuta':       'amount',
    'Betrag in Hauswährung':         'amount',
    'Amount in document currency':   'amount',
    # Clearing
    'Clearing Document':             'clearing_doc',
    'Verrekeningsdocument':          'clearing_doc',
    'Ausgleichsbeleg':               'clearing_doc',
    'Clearing date':                 'clearing_date',
    'Verrekeningsdatum':             'clearing_date',
    # Text / description
    'Text':                          'text',
    'Tekst':                         'text',
    'Document Header Text':          'header_text',
    'Documentkoptekst':              'header_text',
    # Other useful fields
    'Customer':                      'customer',
    'Account':                       'customer',
    'Debtor':                        'customer',
}

def parse_sap_export(file_obj):
    """
    Parse a SAP customer open items export (FBL5N or similar).
    Returns a normalised DataFrame. SAP is authoritative on all classifications.
    """
    df = pd.read_excel(file_obj, sheet_name=0, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # Rename columns
    rename = {k: v for k, v in SAP_COL_MAP.items() if k in df.columns}
    df = df.rename(columns=rename)

    # Ensure essential columns exist
    for col in ['assignment', 'doc_number', 'doc_type', 'amount',
                'clearing_doc', 'clearing_date', 'doc_date', 'due_date',
                'header_text', 'text']:
        if col not in df.columns:
            df[col] = ''

    # Type conversions
    df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0.0)
    for col in ['doc_date', 'due_date', 'clearing_date']:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    # Normalised reference: assignment field, stripped
    df['ref'] = df['assignment'].astype(str).str.strip()

    # Also keep a clean doc_number string for matching
    df['doc_number_str'] = df['doc_number'].astype(str).str.strip().str.split('.').str[0]

    # ── SAP CLASSIFICATION (source of truth) ──────────────────────────────
    # Invoice:      RV doc type with positive amount
    # Credit note:  RV doc type with negative amount, OR RU doc type
    # Payment:      DZ or ZP
    # Clearing:     AB
    # Other:        anything else
    def classify(row):
        dt = str(row['doc_type']).strip().upper()
        amt = row['amount']
        if dt == 'RV':
            return 'CREDIT_NOTE' if amt < 0 else 'INVOICE'
        if dt == 'RU':
            return 'CREDIT_NOTE'
        if dt in ('DZ', 'ZP'):
            return 'PAYMENT'
        if dt == 'AB':
            return 'CLEARING_RESIDUAL'
        return 'OTHER'

    df['sap_class'] = df.apply(classify, axis=1)

    # Open = no clearing document
    df['is_open'] = df['clearing_doc'].isna() | (df['clearing_doc'].astype(str).str.strip() == '')

    return df


# ── REMITTANCE PARSER ────────────────────────────────────────────────────────
def parse_remittance(file_obj, sap_df):
    """
    Parse any client remittance/payment advice Excel.

    Strategy:
    1. Build a lookup of every value in the SAP 'ref' and 'doc_number_str' columns
    2. Scan every cell of the remittance for anything that matches
    3. Return matched pairs: {remittance_value, sap_ref, context}

    We deliberately ignore the customer's signs, labels and formatting.
    SAP will tell us what each matched item actually is.
    """
    raw = pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)
    raw = raw.fillna('')

    # Build SAP lookup sets for fast matching
    sap_refs       = set(sap_df['ref'].unique())
    sap_doc_nums   = set(sap_df['doc_number_str'].unique())
    # Remove unhelpful values
    for junk in ('', 'nan', 'None', '0', '0.0'):
        sap_refs.discard(junk)
        sap_doc_nums.discard(junk)

    found = {}  # ref -> {remittance_value, row, col, context_row}

    for row_idx, row in raw.iterrows():
        for col_idx, cell_val in row.items():
            cell_str = str(cell_val).strip()
            if not cell_str or cell_str.lower() in ('nan', 'none', ''):
                continue

            # Direct match against SAP refs
            if cell_str in sap_refs and cell_str not in found:
                found[cell_str] = {
                    'matched_value':    cell_str,
                    'sap_ref':          cell_str,
                    'match_type':       'assignment',
                    'row':              row_idx,
                    'col':              col_idx,
                    'context':          ' | '.join(str(v) for v in row.values if str(v).strip() and str(v).strip().lower() not in ('nan','none')),
                }
                continue

            # Direct match against SAP doc numbers
            if cell_str in sap_doc_nums and cell_str not in found:
                found[cell_str] = {
                    'matched_value':    cell_str,
                    'sap_ref':          cell_str,
                    'match_type':       'doc_number',
                    'row':              row_idx,
                    'col':              col_idx,
                    'context':          ' | '.join(str(v) for v in row.values if str(v).strip() and str(v).strip().lower() not in ('nan','none')),
                }
                continue

            # Substring search: sometimes a cell contains "ref 9954884107" or similar
            for sap_ref in sap_refs:
                sap_ref_str = str(sap_ref)
                if len(sap_ref_str) >= 6 and sap_ref_str in cell_str and sap_ref not in found:
                    found[sap_ref] = {
                        'matched_value':    cell_str,
                        'sap_ref':          sap_ref,
                        'match_type':       'substring',
                        'row':              row_idx,
                        'col':              col_idx,
                        'context':          ' | '.join(str(v) for v in row.values if str(v).strip() and str(v).strip().lower() not in ('nan','none')),
                    }

    return list(found.values())


# ── MAIN RECONCILIATION ──────────────────────────────────────────────────────
def run_reconciliation(sap_file, remittance_file, payment_amount=None, customer_name=''):
    sap = parse_sap_export(sap_file)
    matches = parse_remittance(remittance_file, sap)

    # SAP lookups
    rv_ru = sap[sap['doc_type'].str.upper().isin(['RV', 'RU'])]
    rv_ru_open    = rv_ru[rv_ru['is_open']]
    rv_ru_cleared = rv_ru[~rv_ru['is_open']]

    def make_lkp(df_in, key='ref'):
        lkp = {}
        for _, row in df_in.iterrows():
            k = row[key]
            if k and k != 'nan':
                lkp.setdefault(k, []).append(row)
        return lkp

    open_ref_lkp    = make_lkp(rv_ru_open, 'ref')
    open_docnum_lkp = make_lkp(rv_ru_open, 'doc_number_str')
    cleared_ref_lkp = make_lkp(rv_ru_cleared, 'ref')
    cleared_doc_lkp = make_lkp(rv_ru_cleared, 'doc_number_str')

    matched_invoices = []
    matched_credits  = []
    already_cleared  = []
    not_found        = []
    matched_refs     = set()

    for item in matches:
        ref   = item['sap_ref']
        mtype = item['match_type']

        # Look up in open items first
        sap_rows = open_ref_lkp.get(ref) or open_docnum_lkp.get(ref)

        if sap_rows:
            matched_refs.add(ref)
            net = sum(r['amount'] for r in sap_rows)
            cls = 'INVOICE' if net > 0 else 'CREDIT_NOTE'
            entry = {
                **item,
                'sap_amount':   net,
                'sap_class':    cls,
                'sap_doc_type': sap_rows[0]['doc_type'],
                'sap_due_date': sap_rows[0]['due_date'],
                'sap_doc_date': sap_rows[0]['doc_date'],
                'sap_header':   str(sap_rows[0]['header_text']) if pd.notna(sap_rows[0]['header_text']) else '',
            }
            if cls == 'INVOICE':
                matched_invoices.append(entry)
            else:
                matched_credits.append(entry)
            continue

        # Check cleared items
        cleared_rows = cleared_ref_lkp.get(ref) or cleared_doc_lkp.get(ref)
        if cleared_rows:
            matched_refs.add(ref)
            already_cleared.append({
                **item,
                'sap_amount':   sum(r['amount'] for r in cleared_rows),
                'sap_class':    cleared_rows[0]['sap_class'],
                'cleared_by':   str(cleared_rows[0]['clearing_doc']),
                'cleared_date': cleared_rows[0]['clearing_date'],
            })
            continue

        not_found.append(item)

    # Open SAP invoices not on remittance
    open_invoices_all = rv_ru_open[rv_ru_open['sap_class'] == 'INVOICE'].copy()
    open_credits_all  = rv_ru_open[rv_ru_open['sap_class'] == 'CREDIT_NOTE'].copy()

    missing_from_rem = open_invoices_all[
        ~open_invoices_all['ref'].isin(matched_refs) &
        ~open_invoices_all['doc_number_str'].isin(matched_refs)
    ].copy()

    return {
        'matched_invoices':  matched_invoices,
        'matched_credits':   matched_credits,
        'already_cleared':   already_cleared,
        'not_found':         not_found,
        'missing_from_rem':  missing_from_rem,
        'open_credits_all':  open_credits_all,
        'payment_amount':    payment_amount,
        'customer_name':     customer_name,
        # totals
        't_inv':     sum(i['sap_amount'] for i in matched_invoices),
        't_cred':    sum(i['sap_amount'] for i in matched_credits),
        't_missing': missing_from_rem['amount'].sum() if len(missing_from_rem) else 0,
        't_open_cr': open_credits_all['amount'].sum() if len(open_credits_all) else 0,
    }


# ── EXCEL REPORT ─────────────────────────────────────────────────────────────
def build_report(results):
    wb = openpyxl.Workbook()

    mi  = results['matched_invoices']
    mc  = results['matched_credits']
    ac  = results['already_cleared']
    nf  = results['not_found']
    mfr = results['missing_from_rem']
    pmt = results['payment_amount']
    cname = results['customer_name'] or 'Customer'

    # ── SUMMARY ──────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = 'Summary'
    _col_w(ws, [4, 46, 4, 18, 4, 16, 4, 4])

    r = 1
    title = f'REMITTANCE RECONCILIATION — {cname}'
    if pmt:
        title += f'  ·  Payment: €{pmt:,.2f}'
    _mr(ws, r, 1, 8, title, bold=True, bg='dk_blue', fg='white', sz=14, ha='center')
    ws.row_dimensions[r].height = 36; r += 1
    _mr(ws, r, 1, 8,
        'SAP is the source of truth. Invoice/credit classification uses SAP doc type and amount — '
        'the customer\'s signs and labels are ignored.',
        bg='md_blue', fg='white', sz=9, ha='center', italic=True)
    ws.row_dimensions[r].height = 18; r += 2

    _mr(ws, r, 1, 8, 'RESULTS', bold=True, bg='dk_blue', fg='white', sz=11, ha='center')
    ws.row_dimensions[r].height = 24; r += 1

    summary_rows = [
        (f'✓  Invoices matched — open in SAP  ({len(mi)} items)',
         results['t_inv'], 'lt_green', 'md_green'),
        (f'✓  Credit notes matched — open in SAP  ({len(mc)} items)',
         results['t_cred'], 'lt_green', 'md_green'),
        ('', None, 'white', 'black'),
        (f'⚠  Already cleared in SAP — potential double payment  ({len(ac)} items)',
         None, 'lt_red', 'md_red'),
        (f'✗  Reference not found anywhere in SAP  ({len(nf)} items)',
         None, 'pink', 'md_red'),
        ('', None, 'white', 'black'),
        (f'Open SAP invoices not on remittance  ({len(mfr)} items)',
         results['t_missing'], 'lt_blue', 'black'),
        (f'Open SAP credit notes available to offset',
         results['t_open_cr'], 'lt_green', 'md_green'),
    ]
    for desc, amt, bg, fg in summary_rows:
        if not desc:
            _mr(ws, r, 1, 8, None, bg='white')
            ws.row_dimensions[r].height = 6; r += 1; continue
        _mr(ws, r, 1, 5, desc, bg=bg, fg=fg, sz=10)
        _c(ws, r, 6, amt, bg=bg, fg=fg, fmt='#,##0.00', ha='right', sz=10)
        _c(ws, r, 7, None, bg=bg)
        _c(ws, r, 8, None, bg=bg)
        ws.row_dimensions[r].height = 20; r += 1

    # ── HELPER: STANDARD INVOICE/CREDIT SHEET ────────────────────────────────
    def make_item_sheet(title, subtitle, items, tab_name, hdr_bg):
        ws2 = wb.create_sheet(tab_name)
        _col_w(ws2, [4, 24, 34, 12, 12, 16, 4])
        r2 = 1
        _mr(ws2, r2, 1, 7, title, bold=True, bg=hdr_bg, fg='white', sz=11)
        ws2.row_dimensions[r2].height = 24; r2 += 1
        _mr(ws2, r2, 1, 7, subtitle, bg='lt_' + hdr_bg if 'lt_' + hdr_bg in BG else 'grey',
            fg='black', sz=9, italic=True, wrap=True)
        ws2.row_dimensions[r2].height = 18; r2 += 1
        _hdr(ws2, r2, [(1, '#'), (2, 'SAP Reference'), (3, 'Remittance Context'),
                       (4, 'Invoice Date'), (5, 'Due Date'), (6, 'SAP Amount (€)'), (7, '')])
        r2 += 1
        total = 0.0
        for idx, item in enumerate(items, 1):
            bg = 'lt_green' if idx % 2 == 0 else 'white'
            _c(ws2, r2, 1, idx, bg=bg, sz=8, ha='center')
            _c(ws2, r2, 2, item['sap_ref'], bg=bg, sz=9, bold=True)
            _c(ws2, r2, 3, item.get('context', ''), bg=bg, sz=8)
            _c(ws2, r2, 4, _fd(item.get('sap_doc_date')), bg=bg, sz=9, ha='center')
            _c(ws2, r2, 5, _fd(item.get('sap_due_date')), bg=bg, sz=9, ha='center')
            _c(ws2, r2, 6, item.get('sap_amount', 0), bg=bg, fmt='#,##0.00', ha='right', sz=9)
            _c(ws2, r2, 7, None, bg=bg)
            total += item.get('sap_amount', 0) or 0
            ws2.row_dimensions[r2].height = 13; r2 += 1
        _mr(ws2, r2, 1, 5, 'TOTAL', bold=True, bg=hdr_bg, fg='white', sz=10)
        _c(ws2, r2, 6, total, bold=True, bg=hdr_bg, fg='white', fmt='#,##0.00', ha='right', sz=10)
        _c(ws2, r2, 7, None, bg=hdr_bg)
        ws2.row_dimensions[r2].height = 16

    make_item_sheet(
        f'✓  INVOICES MATCHED — Open in SAP  ({len(mi)} items  ·  €{results["t_inv"]:,.2f})',
        'These are open RV invoices in SAP that were found on the remittance. '
        'Classification is from SAP only — customer signs are ignored.',
        mi, '✓ Matched Invoices', 'md_green')

    make_item_sheet(
        f'✓  CREDIT NOTES MATCHED — Open in SAP  ({len(mc)} items  ·  €{results["t_cred"]:,.2f})',
        'SAP classifies these as credit notes (negative RV or RU). '
        'Found on the remittance — they will offset against the payment.',
        mc, '✓ Matched Credits', 'md_green')

    # ── ALREADY CLEARED ───────────────────────────────────────────────────────
    ws4 = wb.create_sheet('⚠ Already Cleared')
    _col_w(ws4, [4, 24, 16, 14, 20, 4])
    r4 = 1
    _mr(ws4, r4, 1, 6,
        f'⚠  ALREADY CLEARED IN SAP — Check for Double Payment  ({len(ac)} items)',
        bold=True, bg='md_red', fg='white', sz=11)
    ws4.row_dimensions[r4].height = 24; r4 += 1
    _mr(ws4, r4, 1, 6,
        'These references appear on the remittance but already have a SAP clearing document. '
        'They may have been paid in a previous payment run. Verify before processing.',
        bg='lt_red', fg='black', sz=9, italic=True, wrap=True)
    ws4.row_dimensions[r4].height = 20; r4 += 1
    _hdr(ws4, r4, [(1, '#'), (2, 'SAP Reference'), (3, 'SAP Classification'),
                   (4, 'Cleared Date in SAP'), (5, 'SAP Clearing Doc'), (6, '')])
    r4 += 1
    for idx, item in enumerate(ac, 1):
        _c(ws4, r4, 1, idx, bg='lt_red', sz=8, ha='center')
        _c(ws4, r4, 2, item['sap_ref'], bg='lt_red', sz=9, bold=True)
        _c(ws4, r4, 3, item.get('sap_class', ''), bg='lt_red', sz=9, ha='center')
        _c(ws4, r4, 4, _fd(item.get('cleared_date')), bg='lt_red', sz=9, ha='center')
        _c(ws4, r4, 5, str(item.get('cleared_by', '')), bg='lt_red', sz=9)
        _c(ws4, r4, 6, None, bg='lt_red')
        ws4.row_dimensions[r4].height = 14; r4 += 1

    # ── NOT FOUND ─────────────────────────────────────────────────────────────
    ws5 = wb.create_sheet('✗ Not Found in SAP')
    _col_w(ws5, [4, 24, 40, 4])
    r5 = 1
    _mr(ws5, r5, 1, 4,
        f'✗  NOT FOUND IN SAP  ({len(nf)} items)',
        bold=True, bg='purple', fg='white', sz=11)
    ws5.row_dimensions[r5].height = 24; r5 += 1
    _mr(ws5, r5, 1, 4,
        'These values appeared on the remittance and matched no open or cleared RV/RU '
        'document in SAP. They may be booked under a different reference, not yet posted, '
        'or may not exist on this account.',
        bg='lt_purple', fg='black', sz=9, italic=True, wrap=True)
    ws5.row_dimensions[r5].height = 20; r5 += 1
    _hdr(ws5, r5, [(1, '#'), (2, 'Value from Remittance'), (3, 'Full Row Context'), (4, '')])
    r5 += 1
    for idx, item in enumerate(nf, 1):
        bg = 'lt_purple' if idx % 2 == 0 else 'white'
        _c(ws5, r5, 1, idx, bg=bg, sz=8, ha='center')
        _c(ws5, r5, 2, item['sap_ref'], bg=bg, sz=9, bold=True)
        _c(ws5, r5, 3, item.get('context', ''), bg=bg, sz=8)
        _c(ws5, r5, 4, None, bg=bg)
        ws5.row_dimensions[r5].height = 13; r5 += 1

    # ── SAP OPEN NOT ON REMITTANCE ────────────────────────────────────────────
    if len(mfr) > 0:
        ws6 = wb.create_sheet('SAP Open — Not on Remittance')
        _col_w(ws6, [4, 24, 32, 13, 13, 15])
        r6 = 1
        _mr(ws6, r6, 1, 6,
            f'OPEN IN SAP — NOT ON REMITTANCE  ({len(mfr)} items  ·  €{mfr["amount"].sum():,.2f})',
            bold=True, bg='md_blue', fg='white', sz=11)
        ws6.row_dimensions[r6].height = 24; r6 += 1
        _mr(ws6, r6, 1, 6,
            'These invoices are open in SAP but were not found on the remittance. '
            'The customer may have missed them or plans a separate payment.',
            bg='lt_blue', fg='black', sz=9, italic=True, wrap=True)
        ws6.row_dimensions[r6].height = 18; r6 += 1
        _hdr(ws6, r6, [(1, '#'), (2, 'SAP Reference'), (3, 'Description'),
                       (4, 'Invoice Date'), (5, 'Due Date'), (6, 'Amount (€)')])
        r6 += 1
        for idx, (_, row) in enumerate(mfr.sort_values('due_date').iterrows(), 1):
            bg = 'lt_blue' if idx % 2 == 0 else 'white'
            _c(ws6, r6, 1, idx, bg=bg, sz=8, ha='center')
            _c(ws6, r6, 2, row['ref'], bg=bg, sz=9)
            hdr_txt = str(row['header_text']) if pd.notna(row['header_text']) else ''
            _c(ws6, r6, 3, hdr_txt, bg=bg, sz=9)
            _c(ws6, r6, 4, _fd(row['doc_date']), bg=bg, sz=9, ha='center')
            _c(ws6, r6, 5, _fd(row['due_date']), bg=bg, sz=9, ha='center')
            _c(ws6, r6, 6, row['amount'], bg=bg, fmt='#,##0.00', ha='right', sz=9)
            ws6.row_dimensions[r6].height = 13; r6 += 1
        _mr(ws6, r6, 1, 5, 'TOTAL', bold=True, bg='md_blue', fg='white', sz=10)
        _c(ws6, r6, 6, mfr['amount'].sum(), bold=True, bg='md_blue',
           fg='white', fmt='#,##0.00', ha='right', sz=10)
        ws6.row_dimensions[r6].height = 16

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
