from flask import Flask, request, send_file, render_template, jsonify
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# ── Network colours for UI feedback ─────────────────────────────────────────
NETWORK_COLORS = {
    'zong':    '#d62b2b',
    'ufone':   '#00a651',
    'jazz':    '#f7941d',
    'telenor': '#0073c2'
}

NETWORK_NAMES = {
    'zong':    'Zong',
    'ufone':   'Ufone',
    'jazz':    'Jazz',
    'telenor': 'Telenor'
}

# ── Service numbers to hide from pivot per network ───────────────────────────
SERVICE_NUMBERS = {
    'zong': {
        "230","6009","310","1700","15","1122","6008","25","102","211",
        "700","777","828","829","2161","2300","3238","3239","3458","3557",
        "3737","6911","7028","7078","7258","7861","8227","8558","9080",
        "44342","47650","2545"
    },
    'ufone': {
        "(blank)","414b5548","4554484e4943","476f50","497466617120486f6d657",
        "5054434c","536869666120496e742e","55464f4e45","55666f6e65",
        "180","1166","3404","6525","55506169"
    },
    'jazz': {
        "(blank)","3111","5188","8696","33313131","80000016","302771212",
        "2353494D4C4147","4A415A5A","4A415A5A3447","4A617A7A","4A617A7A203447",
        "4A617A7A436173","2211","2299","3737","8388","32323131","4A617A7A74756E",
        "111","5716","8300","123","668","3444","3445","6009","6060","6064",
        "6080","6633","7770","1","5","8","47","70","188","558","773","940",
        "3977","6381","51876","347534","8885988","MO","ikTok","hatsapp",
        "BAUTH","azzCash","azz 4G","azz","4A415A5A203447","5.34555E+57"
    },
    'telenor': {
        "(blank)","230","5797","6557","7770","7788","8632","727251","727287",
        "150MB43Mins33","Bari Bachat","BudgetOffer","CALLING0Rs6","Din Bhar",
        "Freefire","FULLDAYCALL","internet","MUFT 3GB","PakvAusx","Telenorx",
        "WhtsApp0Rs5","Win Balance"
    }
}

# ── Carrier prefixes to strip per network ────────────────────────────────────
CARRIER_PREFIXES = {
    'zong':    ["110", "92", "38"],
    'ufone':   ["9292", "92"],
    'jazz':    ["640", "92", "38", "40", "2"],
    'telenor': ["9292", "92"]
}

# ────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ────────────────────────────────────────────────────────────────────────────

def thin_border():
    t = Side(style='thin')
    return Border(left=t, right=t, top=t, bottom=t)

def autofit_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 45)

def clean(val):
    """Convert a raw cell value to a clean string, empty string if null."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).strip()
    return '' if s.lower() == 'nan' else s

def strip_prefix(value, network):
    """Strip carrier prefix from a B-Party number."""
    s = clean(value)
    if not s or s == '---':
        return s
    for prefix in CARRIER_PREFIXES.get(network, []):
        if s.startswith(prefix):
            return s[len(prefix):]
    return s

def read_file(file_bytes, filename):
    """Read xlsx / xls / csv into a raw DataFrame (all as strings)."""
    ext = filename.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        for enc in ['utf-8', 'latin-1', 'cp1252']:
            try:
                return pd.read_csv(
                    io.BytesIO(file_bytes), header=None,
                    dtype=str, encoding=enc
                )
            except Exception:
                continue
        raise ValueError("Could not decode CSV — try saving as UTF-8")
    elif ext in ['xlsx', 'xls']:
        engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'
        return pd.read_excel(
            io.BytesIO(file_bytes), header=None,
            dtype=str, engine=engine
        )
    raise ValueError(f"Unsupported file type: .{ext}")

# ────────────────────────────────────────────────────────────────────────────
# NETWORK DETECTION
# ────────────────────────────────────────────────────────────────────────────

def detect_network(df):
    """
    Detect network from raw CDR structure.

    Detection rules (in priority order):
      Jazz    → any header contains 'aparty', 'bparty', or 'calltype' (no space)
      Zong    → any header contains 'call_type' (underscore)
      Telenor → 13+ columns in the raw file
      Ufone   → fallback
    """
    row0 = [clean(v).lower() for v in df.iloc[0]] if len(df) > 0 else []

    if any(h in ('aparty', 'bparty', 'calltype') for h in row0):
        return 'jazz'
    if any('call_type' in h for h in row0):
        return 'zong'
    if df.shape[1] >= 13:
        return 'telenor'
    return 'ufone'

# ────────────────────────────────────────────────────────────────────────────
# ZONG PROCESSOR
# ────────────────────────────────────────────────────────────────────────────

def process_zong(df):
    """
    Raw layout: rows 0-2 = info, row 3 = blank, row 4 = headers, row 5+ = data.
    VBA: move col D → C, delete row 3, set headers in row 5.
    """
    # Info rows (Number / Name / CNIC)
    info_rows = df.iloc[:3].copy()

    # Move column D (index 3) before column C (index 2)
    cols = list(range(df.shape[1]))
    if df.shape[1] >= 4:
        cols = [0, 1, 3, 2] + cols[4:]
    df = df.iloc[:, cols].copy()
    df.columns = range(df.shape[1])

    # Delete blank row 3 (index 3)
    df = df.drop(index=3).reset_index(drop=True)

    # Row index 3 is now the header row
    raw_hdrs = [clean(v) for v in df.iloc[3]]
    headers = raw_hdrs if raw_hdrs else [f'Col{i}' for i in range(df.shape[1])]

    # Force known column names
    def set_hdr(lst, idx, name):
        if idx < len(lst):
            lst[idx] = name
    set_hdr(headers, 1, 'A Party')
    set_hdr(headers, 2, 'B Party')
    set_hdr(headers, 3, 'Date & Time')
    if len(headers) > 9:
        set_hdr(headers, 9, 'Location')

    # Data from index 4 onwards
    data = df.iloc[4:].copy()
    data.columns = headers[:data.shape[1]]
    data = data.reset_index(drop=True)

    # Strip carrier prefixes
    if 'B Party' in data.columns:
        data['B Party'] = data['B Party'].apply(lambda x: strip_prefix(x, 'zong'))

    return info_rows, headers[:data.shape[1]], data

# ────────────────────────────────────────────────────────────────────────────
# UFONE PROCESSOR
# ────────────────────────────────────────────────────────────────────────────

def process_ufone(df):
    """
    Raw layout: C2 holds subscriber number.
    VBA: delete cols B,F,G,L → move A→G pos, K→C pos → combine G+F for Call Type.
    We detect columns by their header names for robustness.
    """
    # Subscriber number from C2
    sub_num = clean(df.iloc[1, 2]) if df.shape[0] > 1 and df.shape[1] > 2 else ''

    # Delete columns at indices 1(B), 5(F), 6(G), 11(L) — go right to left
    drop_idx = sorted([i for i in [1, 5, 6, 11] if i < df.shape[1]], reverse=True)
    df = df.drop(columns=df.columns[drop_idx]).copy()
    df.columns = range(df.shape[1])

    # After deletion, remaining columns (original indices → new):
    # A(0)→0, C(2)→1, D(3)→2, E(4)→3, H(7)→4, I(8)→5, J(9)→6, K(10)→7
    # Move original-A (now idx 0) to last, original-K (now idx 7) to idx 1(C position)
    # Simplified: rearrange so the key columns are in logical order
    n = df.shape[1]
    if n >= 8:
        # Put col 7 (was K) at position 2, col 0 (was A) at position 6
        order = [1, 2, 7, 3, 4, 5, 6, 0] + list(range(8, n))
        df = df.iloc[:, order[:n]].copy()
        df.columns = range(df.shape[1])

    # Combine what are now cols 5 & 4 (date + time) into a new Call Type col
    # (mirrors: Range("C:C").Formula = "=G2 & " " & F2")
    if df.shape[1] >= 6:
        df.insert(2, 'combined', df.iloc[:, 5].astype(str) + ' ' + df.iloc[:, 4].astype(str))
        df.columns = range(df.shape[1])

    # Drop the now-redundant original time columns
    if df.shape[1] >= 8:
        df = df.drop(columns=[df.columns[6], df.columns[7]]).copy()
        df.columns = range(df.shape[1])

    # Headers
    n = df.shape[1]
    headers = [f'Col{i}' for i in range(n)]
    names = ['A Party', 'B Party', 'Call Type', 'Duration', 'Date & Time']
    for i, name in enumerate(names):
        if i < n:
            headers[i] = name

    # Skip header row 0 and use data rows
    data = df.iloc[1:].copy()
    data.columns = headers[:data.shape[1]]
    data = data.reset_index(drop=True)

    # Fill blanks with "---"
    data = data.fillna('---').replace('nan', '---').replace('', '---')

    # Strip prefixes from B Party
    if 'B Party' in data.columns:
        data['B Party'] = data['B Party'].apply(
            lambda x: strip_prefix(x, 'ufone') if x != '---' else x
        )

    # Sort by Date & Time
    if 'Date & Time' in data.columns:
        try:
            data = data.sort_values('Date & Time').reset_index(drop=True)
        except Exception:
            pass

    info_rows = pd.DataFrame([
        ['Number', sub_num, ''],
        ['Name', '', ''],
        ['CNIC', '', '']
    ])

    return info_rows, headers[:data.shape[1]], data

# ────────────────────────────────────────────────────────────────────────────
# JAZZ PROCESSOR
# ────────────────────────────────────────────────────────────────────────────

def process_jazz(df):
    """
    Raw layout: Row 0 = headers (AParty/BParty/CallType — no spaces).
    B2 = subscriber number. Delete cols J:O. Prefixes: 640,92,38,2,40.
    """
    # Subscriber number from B2
    sub_num = clean(df.iloc[1, 1]) if df.shape[0] > 1 and df.shape[1] > 1 else ''

    # Raw headers from row 0
    raw_hdrs = [clean(v) for v in df.iloc[0]]

    # Data rows (row 1 onwards)
    data = df.iloc[1:].copy()
    data.columns = raw_hdrs[:data.shape[1]]
    data = data.reset_index(drop=True)

    # Drop columns J onwards (index 9+) — equivalent to deleting J:O
    if data.shape[1] > 9:
        data = data.iloc[:, :9].copy()

    # Normalise header names (remove spaces / case differences)
    rename = {}
    for col in data.columns:
        cl = col.lower().replace(' ', '').replace('_', '').replace('-', '')
        if cl == 'aparty':         rename[col] = 'A Party'
        elif cl == 'bparty':       rename[col] = 'B Party'
        elif cl == 'calltype':     rename[col] = 'Call Type'
        elif cl == 'datetime':     rename[col] = 'Date & Time'
        elif cl == 'duration':     rename[col] = 'Duration'
        elif cl in ('cellid','cell'): rename[col] = 'Cell ID'
        elif cl == 'imsi':         rename[col] = 'IMSI'
        elif cl == 'imei':         rename[col] = 'IMEI'
        elif cl == 'location':     rename[col] = 'Location'
    data = data.rename(columns=rename)

    # Force positional names for any still-unnamed key columns
    cols = list(data.columns)
    pos_map = {0: 'Call Type', 1: 'A Party', 2: 'B Party', 3: 'Date & Time', 4: 'Duration'}
    for pos, name in pos_map.items():
        if pos < len(cols) and name not in cols:
            cols[pos] = name
    data.columns = cols

    # Strip prefixes from B Party
    if 'B Party' in data.columns:
        data['B Party'] = data['B Party'].apply(lambda x: strip_prefix(x, 'jazz'))

    headers = list(data.columns)
    info_rows = pd.DataFrame([
        ['Number', sub_num, ''],
        ['Name', '', ''],
        ['CNIC', '', '']
    ])

    return info_rows, headers, data

# ────────────────────────────────────────────────────────────────────────────
# TELENOR PROCESSOR
# ────────────────────────────────────────────────────────────────────────────

def process_telenor(df):
    """
    Raw layout: 13+ columns. A2 = subscriber number.
    VBA: sort by col F, replace own number in B:C, copy C→B(skipblanks),
         delete C, combine G+N for Call Type, build Location from N:S.
    """
    # Subscriber number from A2
    sub_num = clean(df.iloc[1, 0]) if df.shape[0] > 1 else ''

    # Sort by column F (index 5)
    if df.shape[1] > 5:
        try:
            df = df.sort_values(by=df.columns[5]).reset_index(drop=True)
        except Exception:
            pass

    # Replace subscriber number in cols B(1) and C(2) with empty string
    for ci in [1, 2]:
        if ci < df.shape[1] and sub_num:
            df.iloc[:, ci] = df.iloc[:, ci].replace(sub_num, '')

    # Copy col C(2) to col B(1) where B is blank, then delete C
    if df.shape[1] > 2:
        b_blank = df.iloc[:, 1].apply(lambda v: clean(v) == '')
        df.iloc[b_blank, 1] = df.iloc[b_blank, 2]
        df = df.drop(columns=df.columns[2]).copy()
        df.columns = range(df.shape[1])

    # Combine col G(now ~5) & col N(now ~11) → Call Type (mirrors VBA: =G2 & " " & N2)
    # After C deletion: A=0,B=1,D=2,E=3,F=4,G=5,H=6,I=7,J=8,K=9,L=10,M=11,N=12,...
    g_idx, n_idx = 5, 11
    if df.shape[1] > n_idx:
        call_type = (df.iloc[:, g_idx].astype(str) + ' ' + df.iloc[:, n_idx].astype(str)).str.strip()
        df.insert(g_idx, 'CallTypeCombined', call_type)
        df.columns = range(df.shape[1])
        # Delete original G (now g_idx+1) and original N (now n_idx+2 after insert)
        drop_g = g_idx + 1
        drop_n = n_idx + 2
        drop_cols = sorted([i for i in [drop_g, drop_n] if i < df.shape[1]], reverse=True)
        df = df.drop(columns=[df.columns[i] for i in drop_cols]).copy()
        df.columns = range(df.shape[1])

    # Build Location by combining remaining high-index columns (original N:S)
    # After all operations, location parts start around index 11
    loc_start = 11
    if df.shape[1] > loc_start:
        loc_parts = [df.iloc[:, i].fillna('').astype(str).replace('nan','')
                     for i in range(loc_start, min(df.shape[1], loc_start + 7))]
        location_col = loc_parts[0]
        for part in loc_parts[1:]:
            location_col = location_col + ' ' + part
        location_col = location_col.str.strip()
        # Keep only columns up to loc_start, then append location
        df = df.iloc[:, :loc_start].copy()
        df.columns = range(df.shape[1])
        df['Location'] = location_col.values

    # Assign headers
    n = df.shape[1]
    headers = [f'Col{i}' for i in range(n)]
    pos_map = {
        0: 'A Party',
        1: 'B Party',
        4: 'Date & Time',
        5: 'Call Type',
        6: 'Duration',
    }
    if 'Location' in df.columns:
        pos_map[n - 1] = 'Location'
    for pos, name in pos_map.items():
        if pos < n:
            headers[pos] = name

    # Insert 4 info rows at top, use row at index 0 as header row
    # (skip raw header row 0 — VBA renames them all manually)
    data = df.iloc[1:].copy()
    data.columns = headers[:data.shape[1]]
    data = data.reset_index(drop=True)

    # Fill blanks with "---"
    data = data.fillna('---').replace('nan', '---').replace('', '---')

    # Strip prefixes from B Party
    if 'B Party' in data.columns:
        data['B Party'] = data['B Party'].apply(
            lambda x: strip_prefix(x, 'telenor') if x != '---' else x
        )

    info_rows = pd.DataFrame([
        ['Number', sub_num, ''],
        ['Name', '', ''],
        ['CNIC', '', '']
    ])

    return info_rows, headers[:data.shape[1]], data

# ────────────────────────────────────────────────────────────────────────────
# SHEET BUILDERS
# ────────────────────────────────────────────────────────────────────────────

def build_cdr_sheet(ws, info_rows, headers, data_df, network):
    bold14  = Font(size=14, bold=True)
    bold    = Font(bold=True)
    normal  = Font()
    hdr_fill = PatternFill(fill_type='solid', fgColor='D9D9D9')
    net_color = NETWORK_COLORS.get(network, 'D62B2B').replace('#', '')
    info_fill = PatternFill(fill_type='solid', fgColor='FFF2CC')

    # ── Info header block (rows 1-3) ──────────────────────────────────────
    labels = ['Number', 'Name', 'CNIC']
    for ri, label in enumerate(labels, start=1):
        c = ws.cell(row=ri, column=1, value=label)
        c.font = bold14
        c.border = thin_border()
        c.fill = info_fill

        ws.merge_cells(f'B{ri}:C{ri}')
        val = ''
        if ri - 1 < len(info_rows):
            raw = info_rows.iloc[ri - 1, 1] if info_rows.shape[1] > 1 else ''
            val = clean(raw)
        vc = ws.cell(row=ri, column=2, value=val)
        vc.font = bold14
        vc.border = thin_border()

    # ── Column headers (row 5) ────────────────────────────────────────────
    for ci, hdr in enumerate(headers, start=1):
        c = ws.cell(row=5, column=ci, value=hdr)
        c.font = bold
        c.border = thin_border()
        c.fill = hdr_fill

    # ── Data rows (row 6 onwards) ─────────────────────────────────────────
    for ri, row_data in data_df.iterrows():
        excel_row = ri + 6
        for ci, val in enumerate(row_data, start=1):
            c = ws.cell(row=excel_row, column=ci, value=clean(val) or val)
            c.font = normal
            c.border = thin_border()

    autofit_columns(ws)

    # ── Page setup ────────────────────────────────────────────────────────
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize   = 5      # Legal
    ws.page_setup.scale       = 70
    ws.print_title_rows       = '5:5'
    ws.page_margins.top       = 0.41
    ws.page_margins.right     = 0.25
    ws.page_margins.bottom    = 0.38
    ws.page_margins.left      = 1.14


def build_summary_sheet(ws, data_df, network):
    bold      = Font(bold=True)
    hdr_fill  = PatternFill(fill_type='solid', fgColor='D9D9D9')
    white_fill = PatternFill(fill_type='solid', fgColor='FFFFFF')

    # Find Call Type column (name varies per network)
    ct_col = None
    for col in data_df.columns:
        if col.replace(' ', '').replace('_', '').lower() == 'calltype':
            ct_col = col
            break

    a_col = next((c for c in data_df.columns if 'a party' in c.lower()), None)
    b_col = next((c for c in data_df.columns if 'b party' in c.lower()), None)

    if not all([ct_col, a_col, b_col]):
        ws.cell(row=3, column=1,
                value='Summary unavailable — A Party / B Party / Call Type columns not found')
        return

    # Filter out service numbers
    svc = SERVICE_NUMBERS.get(network, set())
    df  = data_df[~data_df[b_col].astype(str).isin(svc)].copy()
    df  = df[df[b_col].astype(str).str.strip().ne('')].copy()
    df  = df[df[b_col].astype(str).str.strip().ne('---')].copy()

    # Pivot: count of each call type grouped by A Party + B Party
    call_types = sorted(df[ct_col].dropna().unique())

    try:
        pivot = (
            df.groupby([a_col, b_col])[ct_col]
            .value_counts()
            .unstack(fill_value=0)
        )
        for ct in call_types:
            if ct not in pivot.columns:
                pivot[ct] = 0
        pivot = pivot[call_types]
        pivot['Grand Total'] = pivot.sum(axis=1)
        pivot = pivot.sort_values('Grand Total', ascending=False).reset_index()
    except Exception as e:
        ws.cell(row=3, column=1, value=f'Summary error: {e}')
        return

    # Write headers at row 3
    col_headers = [a_col, b_col] + call_types + ['Grand Total']
    for ci, hdr in enumerate(col_headers, start=1):
        c = ws.cell(row=3, column=ci, value=hdr)
        c.font = bold
        c.border = thin_border()
        c.fill = hdr_fill

    # Write data from row 4
    for ri, row_data in pivot.iterrows():
        er = ri + 4
        for ci, val in enumerate(row_data, start=1):
            c = ws.cell(row=er, column=ci, value=val)
            c.border = thin_border()
            c.fill = white_fill

    # A-party numbers joined (equivalent to the "combine text" section in VBA)
    if a_col in pivot.columns:
        a_parties = ', '.join(
            str(v) for v in pivot[a_col].unique()
            if str(v).strip() not in ('', 'Grand Total', 'nan')
        )
        ws.cell(row=2, column=len(col_headers), value=a_parties)

    autofit_columns(ws)
    ws.page_setup.scale       = 70
    ws.page_margins.top       = 0.41
    ws.page_margins.right     = 0.25
    ws.page_margins.bottom    = 0.38
    ws.page_margins.left      = 1.14


def build_i2_sheet(ws, headers, data_df, network):
    """
    I2 sheet: same as CDR data but strips first 2 chars from A Party (Telenor/Ufone)
    or B Party (Zong/Jazz).
    """
    bold     = Font(bold=True)
    hdr_fill = PatternFill(fill_type='solid', fgColor='D9D9D9')

    strip_col = 'B Party' if network in ('zong', 'jazz') else 'A Party'

    # Headers at row 1
    for ci, hdr in enumerate(headers, start=1):
        c = ws.cell(row=1, column=ci, value=hdr)
        c.font = bold
        c.border = thin_border()
        c.fill = hdr_fill

    # Data from row 2
    for ri, row_data in data_df.iterrows():
        er = ri + 2
        for ci, (col_name, val) in enumerate(row_data.items(), start=1):
            s = clean(val) or str(val)
            if col_name == strip_col and len(s) > 2:
                s = s[2:]
            c = ws.cell(row=er, column=ci, value=s)
            c.border = thin_border()

    autofit_columns(ws)

# ────────────────────────────────────────────────────────────────────────────
# MAIN ASSEMBLY
# ────────────────────────────────────────────────────────────────────────────

def create_output_excel(info_rows, headers, data_df, network):
    wb = Workbook()
    ws_cdr     = wb.active
    ws_cdr.title = 'CDR'
    ws_summary = wb.create_sheet('Summary')
    ws_i2      = wb.create_sheet('I2')

    build_cdr_sheet(ws_cdr, info_rows, headers, data_df, network)
    build_summary_sheet(ws_summary, data_df, network)
    build_i2_sheet(ws_i2, headers, data_df, network)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ────────────────────────────────────────────────────────────────────────────
# FLASK ROUTES
# ────────────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/detect', methods=['POST'])
def detect():
    """Quick detection endpoint — returns network name before full processing."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    try:
        df  = read_file(file.read(), file.filename)
        net = detect_network(df)
        return jsonify({
            'network': net,
            'name':    NETWORK_NAMES[net],
            'color':   NETWORK_COLORS[net]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400

    ext = file.filename.rsplit('.', 1)[-1].lower()
    if ext not in {'xlsx', 'xls', 'csv'}:
        return jsonify({'error': f'.{ext} is not supported'}), 400

    try:
        raw       = file.read()
        df_raw    = read_file(raw, file.filename)
        network   = detect_network(df_raw)

        processors = {
            'zong':    process_zong,
            'ufone':   process_ufone,
            'jazz':    process_jazz,
            'telenor': process_telenor,
        }
        info_rows, headers, data_df = processors[network](df_raw)

        output     = create_output_excel(info_rows, headers, data_df, network)
        out_name   = file.filename.rsplit('.', 1)[0] + f'_{NETWORK_NAMES[network]}_CDR_Processed.xlsx'

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=out_name
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
