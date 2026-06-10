import os
import re
import urllib.request
import pandas as pd
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
from jinja2 import ChoiceLoader, FileSystemLoader

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Support both local 'templates' folder and Render root directory template setups
app.jinja_loader = ChoiceLoader([
    FileSystemLoader(os.path.join(app.root_path, 'templates')),
    FileSystemLoader(app.root_path)
])

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

dashboard_state = {'data': None}

SHEET_CONFIG_PATH = os.path.join('uploads', 'google_sheet_config.txt')
SHEET_LOCAL_PATH = os.path.join('uploads', 'google_sheet_latest.xlsx')
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1soS5--lwUvYTlGZHCtovmUa_H7zl_RGoZFZ1b4ZeD7s/edit?gid=0#gid=0"

def extract_spreadsheet_id(url):
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else None

def download_google_sheet(spreadsheet_id, dest_path):
    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    req = urllib.request.Request(
        export_url,
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    )
    with urllib.request.urlopen(req, timeout=15) as response:
        with open(dest_path, 'wb') as f:
            f.write(response.read())

ORDERED_QTY_ALIASES = ['Ordered Qty', 'OrderQuantity', 'Units Ordered', 'Po Qty', 'Order Quantity']
PAYMENT_ALIASES     = ['Payment Received', 'Payment Received ']

PLATFORM_DISPLAY = {
    'Master Data Instamart': 'Instamart',
    'Master Data Zepto':     'Zepto',
    'Master Data Blinkit':   'Blinkit',
    'Master Data Dmart':     'Dmart',
    'Master Data BB':        'Big Basket',
}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def find_col(df, aliases):
    stripped_map = {c.strip(): c for c in df.columns}
    for a in aliases:
        if a in df.columns:
            return a
        if a.strip() in stripped_map:
            return stripped_map[a.strip()]
    return None


def to_num(series):
    return pd.to_numeric(
        series.astype(str).str.strip().replace(['-', '', 'nan', 'NaN', 'None'], '0'),
        errors='coerce'
    ).fillna(0)


def process_sheet(df, platform_name, month_filter=None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all')

    # Apply month filter from Invoice Month column
    if month_filter and 'Invoice Month' in df.columns:
        df['_month_parsed'] = pd.to_datetime(df['Invoice Month'], errors='coerce')
        try:
            target = datetime.strptime(month_filter, '%b %Y')
            df = df[
                (df['_month_parsed'].dt.year  == target.year) &
                (df['_month_parsed'].dt.month == target.month)
            ]
        except ValueError:
            pass

    ordered_col = find_col(df, ORDERED_QTY_ALIASES)
    payment_col = find_col(df, PAYMENT_ALIASES)

    ordered   = to_num(df[ordered_col])    if ordered_col   else pd.Series(0, index=df.index)
    dispatch  = to_num(df['Dispatch Qty']) if 'Dispatch Qty' in df.columns else pd.Series(0, index=df.index)
    grn_qty   = to_num(df['GRN Qty'])      if 'GRN Qty'     in df.columns else pd.Series(0, index=df.index)
    damaged   = to_num(df['Damaged Qty'])  if 'Damaged Qty' in df.columns else pd.Series(0, index=df.index)
    short_qty = to_num(df['Short Qty'])    if 'Short Qty'   in df.columns else pd.Series(0, index=df.index)
    excess    = to_num(df['Excess Qty'])   if 'Excess Qty'  in df.columns else pd.Series(0, index=df.index)
    rejected  = to_num(df['Total Rejected Qty']) if 'Total Rejected Qty' in df.columns else (damaged + short_qty + excess)
    dn_val    = to_num(df['DN Value'])       if 'DN Value'       in df.columns else pd.Series(0, index=df.index)
    grn_val   = to_num(df['GRN Value'])      if 'GRN Value'      in df.columns else pd.Series(0, index=df.index)
    incl_gst  = to_num(df['Including GST'])  if 'Including GST'  in df.columns else pd.Series(0, index=df.index)
    excl_gst  = to_num(df['Excluding GST'])  if 'Excluding GST'  in df.columns else pd.Series(0, index=df.index)

    # Deduplicate by Invoice No â€” payment is invoice-level but rows repeat per SKU
    if payment_col and 'Invoice No' in df.columns:
        df['_payment_raw'] = to_num(df[payment_col])
        tot_payment = float(
            df.drop_duplicates(subset=['Invoice No'])['_payment_raw'].sum()
        )
    elif payment_col:
        tot_payment = float(to_num(df[payment_col]).sum())
    else:
        tot_payment = 0.0

    tot_ordered  = float(ordered.sum())
    tot_dispatch = float(dispatch.sum())
    tot_grn      = float(grn_qty.sum())
    tot_rejected = float(rejected.sum())
    tot_damaged  = float(damaged.sum())
    tot_short    = float(short_qty.sum())
    tot_excess   = float(excess.sum())
    tot_dn       = float(dn_val.sum())
    tot_grn_val  = float(grn_val.sum())
    tot_incl_gst = float(incl_gst.sum())
    tot_excl_gst = float(excl_gst.sum())

    fill_rate       = round(tot_dispatch / tot_ordered  * 100, 1) if tot_ordered  > 0 else 0
    acceptance_rate = round(tot_grn      / tot_dispatch * 100, 1) if tot_dispatch > 0 else 0
    rejection_rate  = round(tot_rejected / tot_dispatch * 100, 1) if tot_dispatch > 0 else 0
    payment_pct     = round(tot_payment  / tot_grn_val  * 100, 1) if tot_grn_val  > 0 else 0

    # GRN Status
    grn_status = {}
    if 'GRN Status' in df.columns:
        vc = df['GRN Status'].astype(str).str.strip()
        vc = vc[vc.notna() & ~vc.isin(['nan', '0', ''])]
        grn_status = {k: int(v) for k, v in vc.value_counts().items()}

    # Delivery outcome
    log_remark = {}
    if 'Log. Remark' in df.columns:
        vc = df['Log. Remark'].astype(str).str.strip()
        vc = vc[vc.notna() & ~vc.isin(['nan', ''])]
        log_remark = {k: int(v) for k, v in vc.value_counts().items()}

    # Monthly trend (only meaningful when month_filter is None)
    monthly = []
    if not month_filter and 'Invoice Month' in df.columns:
        df['_month'] = pd.to_datetime(df['Invoice Month'], errors='coerce')
        df['_dispatch'] = dispatch
        df['_grn']      = grn_qty
        df['_dn_val']   = dn_val
        df['_grn_val']  = grn_val
        grp = df.dropna(subset=['_month']).groupby('_month').agg(
            dispatch=('_dispatch', 'sum'),
            grn=('_grn', 'sum'),
            dn_val=('_dn_val', 'sum'),
            grn_val=('_grn_val', 'sum'),
        ).reset_index().sort_values('_month')
        for _, row in grp.iterrows():
            monthly.append({
                'month':    row['_month'].strftime('%b %Y'),
                'dispatch': round(float(row['dispatch']), 0),
                'grn':      round(float(row['grn']), 0),
                'dn_val':   round(float(row['dn_val']), 0),
                'grn_val':  round(float(row['grn_val']), 0),
            })

    # City breakdown
    city_data = []
    if 'City' in df.columns:
        df['_dispatch'] = dispatch
        df['_grn']      = grn_qty
        df['_rej']      = rejected
        grp = df[df['City'].astype(str).str.strip() != ''].groupby('City').agg(
            dispatch=('_dispatch', 'sum'),
            grn=('_grn', 'sum'),
            rejected=('_rej', 'sum'),
        ).reset_index().sort_values('dispatch', ascending=False).head(10)
        for _, row in grp.iterrows():
            city_data.append({
                'city':     str(row['City']).strip(),
                'dispatch': round(float(row['dispatch']), 0),
                'grn':      round(float(row['grn']), 0),
                'rejected': round(float(row['rejected']), 0),
            })

    # Top SKUs by dispatch
    sku_col  = find_col(df, ['Zeel Code', 'ZEEL CODE', 'Zeel code'])
    top_skus = []
    if sku_col:
        df['_sku_dispatch'] = dispatch
        df['_sku_rej'] = rejected
        grp = df[df[sku_col].astype(str).str.strip().notna()].groupby(sku_col).agg(
            dispatch=('_sku_dispatch', 'sum'),
            rejected=('_sku_rej', 'sum'),
        ).reset_index().sort_values('dispatch', ascending=False).head(8)
        for _, row in grp.iterrows():
            top_skus.append({
                'sku':      str(row[sku_col])[:45],
                'dispatch': round(float(row['dispatch']), 0),
                'rejected': round(float(row['rejected']), 0),
            })

    return {
        'platform':  platform_name,
        'row_count': len(df),
        'kpis': {
            'ordered':          round(tot_ordered, 0),
            'dispatch':         round(tot_dispatch, 0),
            'grn':              round(tot_grn, 0),
            'rejected':         round(tot_rejected, 0),
            'damaged':          round(tot_damaged, 0),
            'short':            round(tot_short, 0),
            'excess':           round(tot_excess, 0),
            'fill_rate':        fill_rate,
            'acceptance_rate':  acceptance_rate,
            'rejection_rate':   rejection_rate,
            'dn_value':         round(tot_dn, 0),
            'grn_value':        round(tot_grn_val, 0),
            'payment':          round(tot_payment, 0),
            'payment_pct':      payment_pct,
            'dispatch_incl_gst': round(tot_incl_gst, 0),
            'dispatch_excl_gst': round(tot_excl_gst, 0),
        },
        'grn_status':      grn_status,
        'log_remark':      log_remark,
        'monthly':         monthly,
        'city':            city_data,
        'rejection_split': {
            'Damaged': round(tot_damaged, 0),
            'Short':   round(tot_short, 0),
            'Excess':  round(tot_excess, 0),
        },
        'top_skus': top_skus,
    }


def build_summary(platforms):
    k = [p['kpis'] for p in platforms]

    tot_ordered  = sum(x['ordered']  for x in k)
    tot_dispatch = sum(x['dispatch'] for x in k)
    tot_grn      = sum(x['grn']      for x in k)
    tot_rejected = sum(x['rejected'] for x in k)
    tot_dn       = sum(x['dn_value'] for x in k)
    tot_grn_val  = sum(x['grn_value'] for x in k)
    tot_payment  = sum(x['payment']  for x in k)

    totals = {
        'ordered':         round(tot_ordered,  0),
        'dispatch':        round(tot_dispatch, 0),
        'grn':             round(tot_grn,      0),
        'rejected':        round(tot_rejected, 0),
        'damaged':         round(sum(x['damaged'] for x in k), 0),
        'short':           round(sum(x['short']   for x in k), 0),
        'excess':          round(sum(x['excess']  for x in k), 0),
        'fill_rate':       round(tot_dispatch / tot_ordered  * 100, 1) if tot_ordered  > 0 else 0,
        'acceptance_rate': round(tot_grn      / tot_dispatch * 100, 1) if tot_dispatch > 0 else 0,
        'rejection_rate':  round(tot_rejected / tot_dispatch * 100, 1) if tot_dispatch > 0 else 0,
        'dn_value':          round(tot_dn,       0),
        'grn_value':         round(tot_grn_val,  0),
        'payment':           round(tot_payment,  0),
        'payment_pct':       round(tot_payment / tot_grn_val * 100, 1) if tot_grn_val > 0 else 0,
        'dispatch_incl_gst': round(sum(x['dispatch_incl_gst'] for x in k), 0),
        'dispatch_excl_gst': round(sum(x['dispatch_excl_gst'] for x in k), 0),
    }

    labels = [p['platform'] for p in platforms]
    platform_comparison = {
        'labels':          labels,
        'dispatch':        [p['kpis']['dispatch']        for p in platforms],
        'grn':             [p['kpis']['grn']             for p in platforms],
        'rejected':        [p['kpis']['rejected']        for p in platforms],
        'fill_rate':       [p['kpis']['fill_rate']       for p in platforms],
        'acceptance_rate': [p['kpis']['acceptance_rate'] for p in platforms],
        'dn_value':        [p['kpis']['dn_value']        for p in platforms],
        'grn_value':       [p['kpis']['grn_value']       for p in platforms],
        'payment':         [p['kpis']['payment']         for p in platforms],
    }

    # Merge monthly trends from all platforms
    month_map = {}
    for p in platforms:
        for m in p['monthly']:
            key = m['month']
            if key not in month_map:
                month_map[key] = {'month': key, 'dispatch': 0, 'grn': 0, 'dn_val': 0, 'grn_val': 0}
            month_map[key]['dispatch'] += m['dispatch']
            month_map[key]['grn']      += m['grn']
            month_map[key]['dn_val']   += m['dn_val']
            month_map[key]['grn_val']  += m['grn_val']

    def sort_key(e):
        try:    return datetime.strptime(e['month'], '%b %Y')
        except: return datetime.min

    return {
        'kpis':               totals,
        'platform_comparison': platform_comparison,
        'monthly':            sorted(month_map.values(), key=sort_key),
        'rejection_split': {
            'Damaged': round(sum(x['damaged'] for x in k), 0),
            'Short':   round(sum(x['short']   for x in k), 0),
            'Excess':  round(sum(x['excess']  for x in k), 0),
        },
    }


@app.route('/')
def index():
    return render_template('index.html')


def parse_excel_file(filepath, filename):
    xl = pd.ExcelFile(filepath)

    # Cache all sheet DataFrames
    sheet_dfs = {}
    for sheet in xl.sheet_names:
        display = PLATFORM_DISPLAY.get(sheet)
        if display:
            df = pd.read_excel(filepath, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]
            sheet_dfs[display] = df

    if not sheet_dfs:
        raise ValueError('No recognised platform sheets found.')

    # Collect all unique months across all sheets
    all_months_set = set()
    for display, df in sheet_dfs.items():
        if 'Invoice Month' in df.columns:
            months = pd.to_datetime(df['Invoice Month'], errors='coerce').dropna()
            for m in months:
                all_months_set.add(m.strftime('%b %Y'))

    def sort_month(m):
        try:    return datetime.strptime(m, '%b %Y')
        except: return datetime.min

    all_months = sorted(all_months_set, key=sort_month)

    # Full data (all months combined)
    platforms_all = [process_sheet(df, display) for display, df in sheet_dfs.items()]
    summary_all   = build_summary(platforms_all)

    # Per-month slices â€” pre-compute so frontend filters instantly
    monthly_views = {}
    for month in all_months:
        p_month = [process_sheet(df, display, month_filter=month) for display, df in sheet_dfs.items()]
        monthly_views[month] = {
            'summary':   build_summary(p_month),
            'platforms': p_month,
        }

    return {
        'filename':      filename,
        'all_months':    all_months,
        'summary':       summary_all,
        'platforms':     platforms_all,
        'monthly_views': monthly_views,
    }


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Please upload a .xlsx or .xls file'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        result = parse_excel_file(filepath, filename)
        dashboard_state['data'] = result
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'Failed to parse file: {str(e)}'}), 500


@app.route('/connect_sheet', methods=['POST'])
def connect_sheet():
    data = request.get_json() or {}
    url = data.get('url', '').strip()
    if not url:
        return jsonify({'error': 'Please provide a valid Google Sheets URL'}), 400
    
    sheet_id = extract_spreadsheet_id(url)
    if not sheet_id:
        return jsonify({'error': 'Could not extract spreadsheet ID. Please ensure the link is a valid Google Sheets URL.'}), 400
    
    try:
        os.makedirs('uploads', exist_ok=True)
        download_google_sheet(sheet_id, SHEET_LOCAL_PATH)
        
        # Save config
        with open(SHEET_CONFIG_PATH, 'w', encoding='utf-8') as f:
            f.write(url)
            
        result = parse_excel_file(SHEET_LOCAL_PATH, 'Google Sheets Data')
        dashboard_state['data'] = result
        
        # Include sheet URL in the returned state so frontend can show it
        result['sheet_url'] = url
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'Failed to connect and download sheet: {str(e)}. Please check if "Anyone with the link can view" sharing is enabled.'}), 500


@app.route('/sync_sheet', methods=['POST'])
def sync_sheet():
    sheet_url = None
    if os.path.exists(SHEET_CONFIG_PATH):
        try:
            with open(SHEET_CONFIG_PATH, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content != 'disconnected':
                    sheet_url = content
        except Exception:
            pass
    else:
        sheet_url = DEFAULT_GOOGLE_SHEET_URL

    if not sheet_url or sheet_url == 'disconnected':
        return jsonify({'error': 'No Google Sheet is currently connected.'}), 400
        
    try:
        url = sheet_url
            
        sheet_id = extract_spreadsheet_id(url)
        if not sheet_id:
            return jsonify({'error': 'Invalid saved Google Sheets URL.'}), 400
            
        download_google_sheet(sheet_id, SHEET_LOCAL_PATH)
        result = parse_excel_file(SHEET_LOCAL_PATH, 'Google Sheets Data')
        dashboard_state['data'] = result
        result['sheet_url'] = url
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'Failed to sync: {str(e)}. Please verify the internet connection and sheet sharing settings.'}), 500


@app.route('/disconnect_sheet', methods=['POST'])
def disconnect_sheet():
    try:
        os.makedirs('uploads', exist_ok=True)
        with open(SHEET_CONFIG_PATH, 'w', encoding='utf-8') as f:
            f.write('disconnected')
    except Exception:
        pass
    if os.path.exists(SHEET_LOCAL_PATH):
        try:
            os.remove(SHEET_LOCAL_PATH)
        except Exception:
            pass
    dashboard_state['data'] = None
    return jsonify({'success': True})


@app.route('/data')
def get_data():
    sheet_url = None
    if os.path.exists(SHEET_CONFIG_PATH):
        try:
            with open(SHEET_CONFIG_PATH, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content != 'disconnected':
                    sheet_url = content
        except Exception:
            pass
    else:
        # Fallback to default Google Sheet URL if no config file exists yet
        sheet_url = DEFAULT_GOOGLE_SHEET_URL

    if dashboard_state['data'] is None:
        # Try loading Google Sheet first if config exists
        if sheet_url:
            sheet_id = extract_spreadsheet_id(sheet_url)
            if sheet_id:
                try:
                    download_google_sheet(sheet_id, SHEET_LOCAL_PATH)
                    result = parse_excel_file(SHEET_LOCAL_PATH, 'Google Sheets Data')
                    dashboard_state['data'] = result
                except Exception:
                    # Fallback to local copy if offline
                    if os.path.exists(SHEET_LOCAL_PATH):
                        try:
                            result = parse_excel_file(SHEET_LOCAL_PATH, 'Google Sheets Data')
                            dashboard_state['data'] = result
                        except Exception:
                            pass
        
        # Fallback to standard uploaded Excel files
        if dashboard_state['data'] is None:
            upload_dir = app.config['UPLOAD_FOLDER']
            if os.path.exists(upload_dir):
                files = [os.path.join(upload_dir, f) for f in os.listdir(upload_dir) if allowed_file(f) and 'google_sheet' not in f]
                if files:
                    latest_filepath = max(files, key=os.path.getmtime)
                    latest_filename = os.path.basename(latest_filepath)
                    try:
                        result = parse_excel_file(latest_filepath, latest_filename)
                        dashboard_state['data'] = result
                    except Exception:
                        pass

    if dashboard_state['data'] is None:
        return jsonify({'empty': True})
        
    # Make a copy and inject sheet_url if connected
    response_data = dict(dashboard_state['data'])
    if sheet_url:
        response_data['sheet_url'] = sheet_url
        
    return jsonify(response_data)


if __name__ == '__main__':
    app.run(debug=True, port=5050)
