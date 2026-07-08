import re, io, os
from flask import Flask, request, jsonify, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

UPLOAD = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD, exist_ok=True)

def norm(s):
    return re.sub(r'\s+', '', str(s)).lower()

def detect_amount_cols(ws):
    for row in ws.iter_rows(min_row=1, max_row=8, values_only=True):
        cols = [i for i, c in enumerate(row)
                if c and '금액' in re.sub(r'\s+', '', str(c))]
        if len(cols) >= 4:
            return cols[0], cols[1], cols[2], cols[3]
    return 5, 7, 9, 11

def extract_items(ws):
    ci_mat, ci_lab, ci_exp, ci_tot = detect_amount_cols(ws)
    items = []; grand = None
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        raw = row[0]
        if not raw: continue
        name = str(raw).strip()
        if ('합' in name and '계' in name and ('[' in name or '합     계' in name)):
            if isinstance(row[ci_tot], (int, float)):
                grand = {'mat': row[ci_mat] or 0, 'lab': row[ci_lab] or 0,
                         'exp': row[ci_exp] or 0, 'tot': row[ci_tot] or 0}
            continue
        if not isinstance(row[ci_tot], (int, float)): continue
        m = re.match(r'^(\d{2,})\s+', name)
        if m:
            code = m.group(1); level = len(code) // 2; clean = name[m.end():].strip()
        else:
            code_cell = row[13] if len(row) > 13 else None
            code = str(code_cell).strip() if code_cell else None
            level = 2; clean = name
        items.append({'name': clean, 'code': code, 'mat': row[ci_mat] or 0,
                      'lab': row[ci_lab] or 0, 'exp': row[ci_exp] or 0,
                      'tot': row[ci_tot] or 0, 'level': level})
    return items, grand

def do_compare(ws1, ws2):
    items1, grand1 = extract_items(ws1)
    items2, grand2 = extract_items(ws2)
    codes1 = {i['code'] for i in items1 if i['code']}
    codes2 = {i['code'] for i in items2 if i['code']}
    use_code = bool(codes1 and codes2)
    if use_code:
        map1 = {i['code']: i for i in items1}
        key_fn = lambda it: it['code']
    else:
        map1 = {norm(i['name']): i for i in items1}
        key_fn = lambda it: norm(it['name'])
    result = []; used1 = set()
    for it2 in items2:
        key = key_fn(it2)
        if key and key in map1:
            it1 = map1[key]; used1.add(key)
            it1['level'] = it2['level']
            result.append(('match', it1, it2))
        else:
            result.append(('add', None, it2))
    for it1 in items1:
        key = key_fn(it1)
        if not key or key not in used1:
            result.append(('del', it1, None))
    return result, grand1, grand2

def build_excel(rows, grand1, grand2, label1, label2, sheet_name):
    out = Workbook(); ws = out.active
    ws.title = f'{sheet_name}_비교'[:31]
    thin = Side(style='thin', color='CCCCCC')
    med  = Side(style='medium', color='888888')
    def bd(c, heavy=False):
        s = med if heavy else thin
        c.border = Border(left=s, right=s, top=s, bottom=s)

    HDR='1F3864'; WHT='FFFFFF'
    widths = [3, 36, 14, 14, 12, 14, 14, 14, 12, 14, 16, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells('A1:L1')
    c = ws.cell(1, 1, f'{sheet_name} 1대1 비교   ① {label1}  vs  ② {label2}   【원 단위】')
    c.font = Font(name='맑은 고딕', bold=True, size=12, color=WHT)
    c.fill = PatternFill('solid', start_color=HDR)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:B2'); ws.merge_cells('C2:F2'); ws.merge_cells('G2:J2')
    for ci, col_name, bg in [(3, f'① {label1}', '1F3864'),
                              (7, f'② {label2}', 'C55A11'),
                              (11, '증감액', '1A5C38'),
                              (12, '비고', '374151')]:
        c = ws.cell(2, ci, col_name)
        c.font = Font(name='맑은 고딕', bold=True, size=10, color=WHT)
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = Alignment(horizontal='center', vertical='center')
        bd(c, True)
    ws.cell(2, 1).fill = PatternFill('solid', start_color='374151')
    ws.row_dimensions[2].height = 22

    sub = ['L', '', '재료비', '노무비', '경비', '합계', '재료비', '노무비', '경비', '합계', '(합계)', '']
    sub_bgs = ['374151','374151','2C4F84','2C4F84','2C4F84','2C4F84','A04A0E','A04A0E','A04A0E','A04A0E','1F6B3E','374151']
    for ci, (h, bg) in enumerate(zip(sub, sub_bgs), 1):
        c = ws.cell(3, ci, h)
        c.font = Font(name='맑은 고딕', bold=True, size=8, color=WHT)
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = Alignment(horizontal='center', vertical='center')
        bd(c, True)
    ws.row_dimensions[3].height = 22
    ws.freeze_panes = 'A4'

    def fnum(c, v, bg, bold=False):
        c.value = v if v != 0 else None
        c.number_format = '#,##0;(#,##0);-'
        c.font = Font(name='맑은 고딕', size=9, bold=bold,
                      color='C00000' if isinstance(v, (int, float)) and v < 0 else '000000')
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = Alignment(horizontal='right', vertical='center')
        bd(c)
    def ftxt(c, v, bg, bold=False, center=False):
        c.value = v
        c.font = Font(name='맑은 고딕', size=9, bold=bold)
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = Alignment(horizontal='center' if center else 'left',
                                vertical='center', wrap_text=True)
        bd(c)

    r = 4
    for rtype, it1, it2 in rows:
        lv = (it2['level'] if it2 else (it1 or {}).get('level', 2)) or 2
        indent = '  ' * (min(lv, 3) - 1); bold = lv <= 2
        if rtype == 'match':
            diff = it2['tot'] - it1['tot']
            bg = 'E2EFDA' if diff > 0 else ('FCE4D6' if diff < 0 else 'F5F5F5')
            ftxt(ws.cell(r,1), str(lv), bg, center=True)
            ftxt(ws.cell(r,2), indent+it1['name'], bg, bold=bold)
            fnum(ws.cell(r,3), it1['mat'], bg); fnum(ws.cell(r,4), it1['lab'], bg)
            fnum(ws.cell(r,5), it1['exp'], bg); fnum(ws.cell(r,6), it1['tot'], bg, bold=bold)
            fnum(ws.cell(r,7), it2['mat'], bg); fnum(ws.cell(r,8), it2['lab'], bg)
            fnum(ws.cell(r,9), it2['exp'], bg); fnum(ws.cell(r,10), it2['tot'], bg, bold=bold)
            fnum(ws.cell(r,11), diff, bg, bold=bold)
            pct = diff / it1['tot'] if it1['tot'] else 0
            c2 = ws.cell(r, 12, f'{pct:+.1%}' if it1['tot'] else '-')
            c2.font = Font(name='맑은 고딕', size=9)
            c2.fill = PatternFill('solid', start_color=bg)
            c2.alignment = Alignment(horizontal='center', vertical='center'); bd(c2)
        elif rtype == 'add':
            bg = 'DDEBF7'
            ftxt(ws.cell(r,1), str(lv), bg, center=True)
            ftxt(ws.cell(r,2), indent+it2['name'], bg, bold=bold)
            for ci in [3,4,5,6]: ftxt(ws.cell(r,ci), '-', bg, center=True)
            fnum(ws.cell(r,7), it2['mat'], bg); fnum(ws.cell(r,8), it2['lab'], bg)
            fnum(ws.cell(r,9), it2['exp'], bg); fnum(ws.cell(r,10), it2['tot'], bg, bold=True)
            fnum(ws.cell(r,11), it2['tot'], bg, bold=True)
            c2 = ws.cell(r, 12, '내역추가')
            c2.font = Font(name='맑은 고딕', size=9, bold=True, color=WHT)
            c2.fill = PatternFill('solid', start_color='2E75B6')
            c2.alignment = Alignment(horizontal='center', vertical='center'); bd(c2)
        else:
            bg = 'FFF2CC'
            ftxt(ws.cell(r,1), str(lv), bg, center=True)
            ftxt(ws.cell(r,2), indent+it1['name'], bg, bold=bold)
            fnum(ws.cell(r,3), it1['mat'], bg); fnum(ws.cell(r,4), it1['lab'], bg)
            fnum(ws.cell(r,5), it1['exp'], bg); fnum(ws.cell(r,6), it1['tot'], bg, bold=True)
            for ci in [7,8,9,10]: ftxt(ws.cell(r,ci), '-', bg, center=True)
            fnum(ws.cell(r,11), -it1['tot'], bg, bold=True)
            c2 = ws.cell(r, 12, '내역삭제')
            c2.font = Font(name='맑은 고딕', size=9, bold=True, color=WHT)
            c2.fill = PatternFill('solid', start_color='C55A11')
            c2.alignment = Alignment(horizontal='center', vertical='center'); bd(c2)
        ws.row_dimensions[r].height = 16; r += 1

    r += 1
    ws.merge_cells(f'A{r}:B{r}')
    c = ws.cell(r, 1, '■ 공식 합계 ([합 계] 행 기준)')
    c.font = Font(name='맑은 고딕', bold=True, size=10, color=WHT)
    c.fill = PatternFill('solid', start_color=HDR)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(r, 2).fill = PatternFill('solid', start_color=HDR)
    bg = 'D6E4FF'
    g1 = grand1 or {'mat':0,'lab':0,'exp':0,'tot':0}
    g2 = grand2 or {'mat':0,'lab':0,'exp':0,'tot':0}
    fnum(ws.cell(r,3), g1['mat'], bg, True); fnum(ws.cell(r,4), g1['lab'], bg, True)
    fnum(ws.cell(r,5), g1['exp'], bg, True); fnum(ws.cell(r,6), g1['tot'], bg, True)
    fnum(ws.cell(r,7), g2['mat'], bg, True); fnum(ws.cell(r,8), g2['lab'], bg, True)
    fnum(ws.cell(r,9), g2['exp'], bg, True); fnum(ws.cell(r,10), g2['tot'], bg, True)
    diff_tot = g2['tot'] - g1['tot']
    fnum(ws.cell(r,11), diff_tot, bg, True)
    pct = diff_tot / g1['tot'] if g1['tot'] else 0
    c2 = ws.cell(r, 12, f'{pct:+.1%}')
    c2.font = Font(name='맑은 고딕', size=10, bold=True)
    c2.fill = PatternFill('solid', start_color=bg)
    c2.alignment = Alignment(horizontal='center', vertical='center'); bd(c2)
    ws.row_dimensions[r].height = 22

    buf = io.BytesIO(); out.save(buf); buf.seek(0)
    return buf


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/sheets', methods=['POST'])
def get_sheets():
    files = {}
    for key in ['file1', 'file2']:
        f = request.files.get(key)
        if not f: return jsonify(error=f'{key} 없음'), 400
        path = os.path.join(UPLOAD, f'{key}_{f.filename}')
        f.save(path)
        wb = load_workbook(path, data_only=True, read_only=True)
        files[key] = {'name': f.filename, 'sheets': list(wb.sheetnames), 'path': path}
        wb.close()
    return jsonify(files)

@app.route('/api/compare', methods=['POST'])
def compare():
    data = request.json
    wb1 = load_workbook(data['path1'], data_only=True)
    wb2 = load_workbook(data['path2'], data_only=True)
    rows, grand1, grand2 = do_compare(wb1[data['sheet1']], wb2[data['sheet2']])

    table = []
    for rtype, it1, it2 in rows:
        lv = (it2['level'] if it2 else (it1 or {}).get('level', 2)) or 2
        name = it2['name'] if it2 else it1['name']
        diff = (it2['tot'] - it1['tot']) if rtype == 'match' else \
               (it2['tot'] if rtype == 'add' else -it1['tot'])
        pct = diff / it1['tot'] if rtype == 'match' and it1['tot'] else None
        table.append({
            'type': rtype, 'level': lv, 'name': name,
            'mat1': it1['mat'] if it1 else None, 'lab1': it1['lab'] if it1 else None,
            'exp1': it1['exp'] if it1 else None, 'tot1': it1['tot'] if it1 else None,
            'mat2': it2['mat'] if it2 else None, 'lab2': it2['lab'] if it2 else None,
            'exp2': it2['exp'] if it2 else None, 'tot2': it2['tot'] if it2 else None,
            'diff': diff, 'pct': pct,
        })

    return jsonify({
        'table': table, 'grand1': grand1, 'grand2': grand2,
        'summary': {
            'match': sum(1 for r in rows if r[0] == 'match'),
            'add':   sum(1 for r in rows if r[0] == 'add'),
            'del':   sum(1 for r in rows if r[0] == 'del'),
        },
        'path1': data['path1'], 'path2': data['path2'],
        'sheet1': data['sheet1'], 'sheet2': data['sheet2'],
        'label1': data['label1'], 'label2': data['label2'],
    })

@app.route('/api/download', methods=['POST'])
def download():
    data = request.json
    wb1 = load_workbook(data['path1'], data_only=True)
    wb2 = load_workbook(data['path2'], data_only=True)
    rows, grand1, grand2 = do_compare(wb1[data['sheet1']], wb2[data['sheet2']])
    buf = build_excel(rows, grand1, grand2, data['label1'], data['label2'], data['sheet1'])
    return send_file(buf, as_attachment=True,
                     download_name=f"{data['sheet1']}_1대1비교.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
