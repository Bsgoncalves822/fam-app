import os, re, uuid, threading, zipfile, io, time, requests, logging, json, sys, hashlib
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# SELF-UPDATE
# ---------------------------------------------------------------------------
UPDATE_URL = "https://raw.githubusercontent.com/Bsgoncalves822/fam-app/main/app.py"
HASH_FILE  = os.path.join(BASE_DIR, 'data', 'update_hashes.json')
BRANCH     = "main"

def _md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()

def _load_hashes():
    try:
        with open(HASH_FILE) as f: return json.load(f)
    except Exception: return {}

def _save_hashes(d):
    os.makedirs(os.path.dirname(HASH_FILE), exist_ok=True)
    with open(HASH_FILE, 'w') as f: json.dump(d, f)

def try_self_update():
    try:
        import subprocess
        result = subprocess.run(['git', 'fetch', '--quiet', 'origin', BRANCH],
                                cwd=BASE_DIR, capture_output=True, timeout=15)
        if result.returncode == 0:
            diff = subprocess.run(['git', 'diff', '--name-only', f'HEAD..origin/{BRANCH}'],
                                  cwd=BASE_DIR, capture_output=True, text=True, timeout=10,
                                  encoding='utf-8', errors='replace')
            changed = [l.strip() for l in diff.stdout.splitlines() if l.strip()]
            if changed:
                print(f"[UPDATE] Git: {len(changed)} arquivo(s) atualizado(s): {', '.join(changed)}")
                subprocess.run(['git', 'reset', '--hard', f'origin/{BRANCH}'],
                               cwd=BASE_DIR, capture_output=True, timeout=30)
                import shutil
                pycache = os.path.join(BASE_DIR, '__pycache__')
                if os.path.isdir(pycache): shutil.rmtree(pycache, ignore_errors=True)
                print("[UPDATE] Reiniciando app...")
                os.execv(sys.executable, [sys.executable] + sys.argv)
            else:
                print("[UPDATE] Git: ja esta na versao mais recente.")
            return
    except Exception as e:
        print(f"[UPDATE] Git indisponivel ({e}), tentando URL fallback...")
    try:
        hashes = _load_hashes()
        resp = requests.get(UPDATE_URL + f"?cb={int(time.time())}", timeout=10)
        if resp.status_code != 200: return
        remote_content = resp.content
        remote_hash = hashlib.md5(remote_content).hexdigest()
        local_hash  = _md5(__file__) if os.path.exists(__file__) else ''
        if remote_hash != local_hash and remote_hash != hashes.get('app.py', ''):
            with open(__file__, 'wb') as f: f.write(remote_content)
            hashes['app.py'] = remote_hash
            _save_hashes(hashes)
            import shutil
            pycache = os.path.join(BASE_DIR, '__pycache__')
            if os.path.isdir(pycache): shutil.rmtree(pycache, ignore_errors=True)
            print("[UPDATE] Reiniciando app...")
            os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception as e:
        print(f"[UPDATE] Falha no update ({e}), continuando com versao atual.")

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
LOG_DIR = os.path.join(BASE_DIR, 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, datetime.now().strftime('%Y-%m-%d') + '.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.FileHandler(log_file, encoding='utf-8'), logging.StreamHandler()]
)
logger = logging.getLogger('fam')

# ---------------------------------------------------------------------------
# FLASK
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
for d in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'],
          os.path.join(BASE_DIR, 'data')]:
    os.makedirs(d, exist_ok=True)

tasks = {}

# ---------------------------------------------------------------------------
# FILENAME PARSER  (ported from parse_fam_all_months.py)
# ---------------------------------------------------------------------------
def fix_year(y):
    y = y[:4]
    if y.startswith('200') and int(y) > 2030: y = '20' + y[2:]
    return y

def parse_amount(raw):
    if raw is None: return None
    raw = str(raw).strip()
    is_negative = raw.startswith('-')
    raw = raw.lstrip('-').strip()
    raw = re.sub(r'[^\d,.]', '', raw)
    raw = re.sub(r'L$', '', raw)
    raw = re.sub(r'\.(?=\d{3},)', '', raw)
    raw = re.sub(r'(\d)\s+(\d)', r'\1\2', raw)
    if ',' in raw:
        parts = raw.rsplit(',', 1)
        integer_part = parts[0].replace('.', '').replace(',', '')
        decimal_part = parts[1]
        try:
            val = float(f"{integer_part}.{decimal_part}")
            return -val if is_negative else val
        except:
            return None
    return None

def preprocess(name):
    name = re.sub(r'^[a-zA-Z]+(?=\d{1,2}[-.\s]\d{2}[-.\s])', '', name)
    name = re.sub(r'\s*\(VENC[^)]*\)\s*$', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*\(\d+\)\s*$', '', name)
    name = re.sub(
        r'^(\d{1,2}[-.]?\d{2}[-.]?)(\d{1})([A-Z])',
        lambda m: m.group(1) + '2025 ' + m.group(3), name
    )
    return name

def parse_filename(filename):
    name = re.sub(r'\.pdf$', '', filename, flags=re.IGNORECASE).strip()
    if re.match(r'^sicredi_[A-Za-z0-9]', name, re.IGNORECASE):
        return None, None, None, None, filename, "padrão hash — sem data"
    name = preprocess(name)
    date_m = re.match(r'^(\d{1,2})[\s.\-]+(\d{2})[\s.\-+]+(\d{4,5})[\s.\-]*\s*', name)
    if not date_m:
        return None, None, None, None, filename, "data não encontrada"
    day   = date_m.group(1).zfill(2)
    month = date_m.group(2)
    year  = fix_year(date_m.group(3))
    date_str  = f"{day}/{month}/{year}"
    remainder = name[date_m.end():].strip()
    remainder = re.sub(r'^[-–\s]+', '', remainder).strip()

    amount_m = re.search(r'[-–]\s*(-?[\d.,]+,\d{2,})\s*$', remainder)
    if amount_m:
        amount = parse_amount(amount_m.group(1))
        body   = remainder[:amount_m.start()].strip().rstrip('-–').strip()
    else:
        fallback_m = re.search(r'([\d.]+,\d{2,})\s*$', remainder)
        if fallback_m:
            amount = parse_amount(fallback_m.group(1))
            body   = remainder[:fallback_m.start()].strip().rstrip('.-– ').strip()
        else:
            amount = None
            body   = remainder.strip()

    paren_m = re.search(r'\(\s*([^)]+?)\s*\)\s*$', body)
    if paren_m:
        payment = paren_m.group(1).strip()
        supplier = body[:paren_m.start()].strip().rstrip('-–').strip()
    else:
        pay_kw = re.search(r'\b(PIX\s+\w+|BOLETO|TED|DOC|CHEQUE|DEBITO|CREDITO|SICREDI|BANCO\s+\w+)\b',
                           body, re.IGNORECASE)
        if pay_kw:
            payment  = pay_kw.group(1).strip()
            supplier = (body[:pay_kw.start()] + body[pay_kw.end():]).strip().rstrip('-–').strip()
        else:
            payment  = None
            supplier = body.strip().lstrip('-–').strip()

    flag = None
    if amount is None: flag = "valor não encontrado"
    return date_str, supplier, payment, amount, filename, flag

# ---------------------------------------------------------------------------
# EXCEL BUILDER
# ---------------------------------------------------------------------------
ACCENT  = "00D4FF"
GREEN   = "00E676"
SURFACE = "141720"
BG      = "0D0F12"
BORDER  = "252A38"

def hex_fill(hex_color): return PatternFill("solid", fgColor=hex_color)

def build_excel(rows_by_month, out_path):
    wb = Workbook()
    wb.remove(wb.active)

    header_cols = ["Data", "Fornecedor / Descrição", "Forma de Pagamento", "Valor (R$)", "Arquivo"]
    col_widths   = [14, 52, 22, 16, 60]

    all_rows = []
    for month_label in sorted(rows_by_month.keys()):
        rows = rows_by_month[month_label]
        ws = wb.create_sheet(title=month_label)
        ws.sheet_view.showGridLines = False

        # Title row
        ws.merge_cells(f"A1:{get_column_letter(len(header_cols))}1")
        title_cell = ws["A1"]
        title_cell.value = f"FAM — Comprovantes · {month_label}"
        title_cell.font = Font(name="Arial", bold=True, size=13, color=ACCENT)
        title_cell.fill = hex_fill(BG)
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 32

        # Header row
        for ci, (hdr, w) in enumerate(zip(header_cols, col_widths), 1):
            cell = ws.cell(row=2, column=ci, value=hdr)
            cell.font      = Font(name="Arial", bold=True, size=9, color=ACCENT)
            cell.fill      = hex_fill(SURFACE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[2].height = 20

        thin = Side(style="thin", color=BORDER)
        border = Border(bottom=Side(style="thin", color="1C2030"))

        total = 0.0
        for ri, row in enumerate(rows, 3):
            date, supplier, payment, amount, fname, flag = row
            row_fill = hex_fill("0A0C10") if ri % 2 == 0 else hex_fill(BG)
            values = [date or "", supplier or fname, payment or "", amount or "", fname]
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.fill      = row_fill
                cell.alignment = Alignment(vertical="center")
                cell.border    = border
                if ci == 4:
                    cell.number_format = '#.##0,00'
                    cell.font = Font(name="Arial", size=9,
                                     color="FF5252" if flag else "E8ECF4")
                else:
                    cell.font = Font(name="Arial", size=9,
                                     color="6B7599" if flag else "E8ECF4")
            ws.row_dimensions[ri].height = 16
            if amount: total += amount
            all_rows.append(row)

        # Total row
        tr = len(rows) + 3
        ws.merge_cells(f"A{tr}:C{tr}")
        tot_label = ws[f"A{tr}"]
        tot_label.value     = "TOTAL"
        tot_label.font      = Font(name="Arial", bold=True, size=9, color=ACCENT)
        tot_label.fill      = hex_fill(SURFACE)
        tot_label.alignment = Alignment(horizontal="right", vertical="center")
        tot_val = ws.cell(row=tr, column=4, value=total)
        tot_val.font         = Font(name="Arial", bold=True, size=9, color=GREEN)
        tot_val.fill         = hex_fill(SURFACE)
        tot_val.number_format = '#.##0,00'
        tot_val.alignment    = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[tr].height = 20

    # Summary sheet
    ws = wb.create_sheet(title="Resumo", index=0)
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f"A1:{get_column_letter(len(header_cols))}1")
    c = ws["A1"]
    c.value = "FAM — Resumo Consolidado"
    c.font  = Font(name="Arial", bold=True, size=13, color=ACCENT)
    c.fill  = hex_fill(BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    for ci, (hdr, w) in enumerate(zip(header_cols, col_widths), 1):
        cell = ws.cell(row=2, column=ci, value=hdr)
        cell.font      = Font(name="Arial", bold=True, size=9, color=ACCENT)
        cell.fill      = hex_fill(SURFACE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[2].height = 20

    border = Border(bottom=Side(style="thin", color="1C2030"))
    grand_total = 0.0
    for ri, row in enumerate(all_rows, 3):
        date, supplier, payment, amount, fname, flag = row
        row_fill = hex_fill("0A0C10") if ri % 2 == 0 else hex_fill(BG)
        values = [date or "", supplier or fname, payment or "", amount or "", fname]
        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill      = row_fill
            cell.alignment = Alignment(vertical="center")
            cell.border    = border
            if ci == 4:
                cell.number_format = '#.##0,00'
                cell.font = Font(name="Arial", size=9,
                                 color="FF5252" if flag else "E8ECF4")
            else:
                cell.font = Font(name="Arial", size=9,
                                 color="6B7599" if flag else "E8ECF4")
        ws.row_dimensions[ri].height = 16
        if amount: grand_total += amount

    tr = len(all_rows) + 3
    ws.merge_cells(f"A{tr}:C{tr}")
    tl = ws[f"A{tr}"]
    tl.value = "TOTAL GERAL"
    tl.font  = Font(name="Arial", bold=True, size=9, color=ACCENT)
    tl.fill  = hex_fill(SURFACE)
    tl.alignment = Alignment(horizontal="right", vertical="center")
    tv = ws.cell(row=tr, column=4, value=grand_total)
    tv.font         = Font(name="Arial", bold=True, size=9, color=GREEN)
    tv.fill         = hex_fill(SURFACE)
    tv.number_format = '#.##0,00'
    tv.alignment    = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[tr].height = 20

    wb.save(out_path)

# ---------------------------------------------------------------------------
# TASK RUNNER
# ---------------------------------------------------------------------------
def run_parse(task_id, zip_path):
    def log(msg):
        tasks[task_id]['log'].append(msg)
        logger.info(msg)

    try:
        log("Lendo arquivos do ZIP...")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            names = [n for n in zf.namelist() if n.lower().endswith('.pdf')]

        log(f"{len(names)} comprovantes encontrados.")

        rows_by_month = {}
        flagged = 0
        for fname in names:
            basename = os.path.basename(fname)
            date, supplier, payment, amount, original, flag = parse_filename(basename)
            if flag: flagged += 1

            # Derive month label from parsed date or filename
            if date:
                parts = date.split('/')
                month_label = f"{parts[1]}-{parts[2]}"
            else:
                # Try to grab month from filename directly
                m = re.search(r'\d{1,2}[-./]\d{2}[-./](\d{4})', basename)
                month_label = "Sem Data"

            rows_by_month.setdefault(month_label, []).append(
                (date, supplier, payment, amount, original, flag)
            )

        log(f"Parseados: {len(names) - flagged} OK · {flagged} com flag")
        log("Gerando Excel...")

        out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{task_id}.xlsx")
        build_excel(rows_by_month, out_path)

        tasks[task_id]['status'] = 'DONE'
        tasks[task_id]['file']   = out_path
        log(f"Concluído! {len(names)} comprovantes → Excel pronto.")

    except Exception as e:
        import traceback
        tasks[task_id]['status'] = 'ERROR'
        tasks[task_id]['log'].append(f"ERRO: {str(e)}")
        tasks[task_id]['log'].append(traceback.format_exc())
        logger.error(f"Run failed: {e}")

# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload/zip', methods=['POST'])
def upload_zip():
    if 'file' not in request.files: return jsonify({"error": "No file"}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f"comp_{uuid.uuid4().hex}.zip")
    f.save(path)
    return jsonify({"path": path})

@app.route('/processar', methods=['POST'])
def processar():
    data     = request.json
    zip_path = data.get('zip_path')
    if not zip_path: return jsonify({"error": "Missing zip"}), 400
    task_id = uuid.uuid4().hex
    tasks[task_id] = {"status": "PENDING", "log": [], "file": None}
    threading.Thread(target=run_parse, args=(task_id, zip_path)).start()
    return jsonify({"task_id": task_id})

@app.route('/status/<task_id>')
def status(task_id):
    if task_id not in tasks: return jsonify({"error": "Invalid task"}), 404
    return jsonify(tasks[task_id])

@app.route('/download/<task_id>')
def download(task_id):
    if task_id not in tasks or tasks[task_id]['status'] != 'DONE':
        return jsonify({"error": "Not ready"}), 400
    return send_file(tasks[task_id]['file'], as_attachment=True,
                     download_name='FAM_Comprovantes.xlsx')

# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    try_self_update()
    logger.info("FAM App (ZIP Parser) starting on port 5002")
    app.run(host='0.0.0.0', port=5002, debug=False)
