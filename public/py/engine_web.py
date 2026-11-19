# -*- coding: utf-8 -*-
"""ブラウザ（Pyodide）から呼ぶための入口。引数も戻り値も JSON 文字列"""
import base64, io, json
import openpyxl
import shift_engine as E


def _cfg(b64):
    return E.Config(io.BytesIO(base64.b64decode(b64)))


def _vjson(V):
    return [dict(kind=v['kind'], date=v['date'].isoformat() if v['date'] else None,
                 staff=v['staff'], msg=v['msg'], target=v.get('target')) for v in V]


def _specials(lst):
    out = []
    for x in lst or []:
        d = E.to_date(x.get('date'))
        if d and E.norm(x.get('target')):
            out.append(dict(date=d, target=E.norm(x['target']), min=E.num(x.get('min')),
                            max=E.num(x.get('max')), note=x.get('note') or ''))
    return out


def _targets(req_by_day):
    """その月に必要人数が決まっている対象（日別の人数を表示する行）"""
    return list(dict.fromkeys(t for lst in req_by_day.values() for (t, *_) in lst))


def _out(fn):
    """例外（SystemExit も含む）を {error: ...} にして返す"""
    def wrap(*a):
        try:
            return json.dumps(fn(*a), ensure_ascii=False)
        except BaseException as e:
            return json.dumps({'error': str(e) or e.__class__.__name__}, ensure_ascii=False)
    return wrap


@_out
def info(settings_b64):
    cfg = _cfg(settings_b64)
    return dict(
        staff=[dict(no=s['no'], name=s['name'], key=s['key'], groups=sorted(s['groups']),
                    holidays=E.num(s['holidays'])) for s in cfg.staff],
        codes=[dict(code=c, name=str(v['name']), kind=v['kind'], color=v['color'], auto=v['auto'],
                    tags=sorted(v['tags']), group=v['group']) for c, v in cfg.codes.items()],
        groups=cfg.groups,
        tags=sorted(cfg.all_tags),
        holidays={d.isoformat(): str(n) for d, n in cfg.holidays.items()},
        basic={k: (v if isinstance(v, (int, float, str)) or v is None else str(v)) for k, v in cfg.basic.items()},
        errors=list(dict.fromkeys(cfg.errors)),
    )


@_out
def solve(req_json):
    req = json.loads(req_json)
    cfg = _cfg(req['settings'])
    month = E.to_date(req['month'])
    nd = len(E.month_dates(month))
    keys = {s['key'] for s in cfg.staff}
    cells = {}
    for c in req.get('cells', []):
        d = int(c['day'])
        if c['key'] in keys and 1 <= d <= nd and str(c.get('value') or '').strip():
            cells[(c['key'], d - 1)] = str(c['value']).strip()
    specials = _specials(req.get('specials'))
    buf = io.BytesIO()
    r = E.run_solve(cfg, month, cells, E.num(req.get('holidays_n')), specials,
                    E.num(req.get('seconds')), buf, log=lambda *a: None)
    return dict(
        status=r['status'], notes=r['notes'], violations=_vjson(r['violations']), holidays_n=r['hol_n'],
        targets=_targets(E.day_requirements(cfg, r['dates'], specials)),
        table={s['key']: [r['table'].get((s['key'], i)) for i in range(nd)] for s in r['staff']},
        xlsx=base64.b64encode(buf.getvalue()).decode())


@_out
def check(settings_b64, xlsx_b64, sheet, month, filename, specials_json='[]'):
    cfg = _cfg(settings_b64)
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(xlsx_b64)), data_only=True)
    sheet = sheet or None
    if sheet and sheet not in wb.sheetnames:
        raise ValueError(f'シート「{sheet}」がありません（{"、".join(wb.sheetnames[-6:])} など）')
    # 手作りの勤務表には特別日シートが無いので、画面で登録したその月の特別日を足す
    extra = json.loads(specials_json or '[]')
    if extra and '特別日' not in wb.sheetnames:
        sp = wb.create_sheet('特別日')
        sp.append(['日付', '対象', '最小', '最大', '備考'])
        for x in extra:
            sp.append([E.to_date(x.get('date')), x.get('target'), x.get('min'), x.get('max'), x.get('note')])
    buf = io.BytesIO()
    r = E.run_check(cfg, wb, sheet, month or None, buf, filename or '')
    return dict(month=r['month'].isoformat(), violations=_vjson(r['violations']), notes=r['notes'],
                sheets=wb.sheetnames, xlsx=base64.b64encode(buf.getvalue()).decode())


def month_of(xlsx_b64, sheet, month):
    """チェック前に、その表が何年何月か調べる（特別日を探すため）"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(xlsx_b64)), data_only=True)
        ws = wb[sheet] if sheet else (wb['勤務表'] if '勤務表' in wb.sheetnames else wb.worksheets[0])
        return json.dumps(E.month_of(ws, month or None).isoformat()[:7])
    except BaseException:
        return json.dumps('')


# ------------------------------------------------------------ 設定をブラウザで編集する
EDIT_SHEETS = ['基本設定', '勤務種類', '職員', '必要人数', 'ルール', '祝日']


def _cell_out(v):
    """Excel の値 → 画面に出す文字（日付は YYYY-MM-DD、時刻は H:MM）"""
    import datetime
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        if v.hour == v.minute == 0 and v.year > 1900:
            return v.date().isoformat()
        return f'{v.hour}:{v.minute:02d}'
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, datetime.time):
        return f'{v.hour}:{v.minute:02d}'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _cell_in(v, header):
    """画面の文字 → Excel に書く値（数字は数値、祝日の日付は日付）"""
    import datetime, re
    s = str(v if v is not None else '').strip()
    if s == '':
        return None
    if header == '日付':
        d = E.to_date(s)
        return datetime.datetime(d.year, d.month, d.day) if d else s
    if header in ('開始', '終了', 'コード', 'No', '使える勤務', 'タグ', '対象職員', 'グループ'):
        return int(s) if header == 'No' and re.fullmatch(r'\d+', s) else s
    if re.fullmatch(r'-?\d+', s):
        return int(s)
    if re.fullmatch(r'-?\d+\.\d+', s):
        return float(s)
    return s


@_out
def sheets(settings_b64):
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(settings_b64)))
    out = {}
    for name in EDIT_SHEETS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        rows = [[_cell_out(v) for v in r] for r in ws.iter_rows(values_only=True)]
        if name == '基本設定':
            head, body = ['項目', '値'], [r[:2] + [''] * (2 - len(r[:2])) for r in rows if any(r[:2])]
        else:
            if not rows:
                continue
            head = []
            for h in rows[0]:          # 見出しは最初の空欄まで（右側のメモ書きは見出しにしない）
                if not h:
                    break
                head.append(h)
            body = [r[:len(head)] + [''] * (len(head) - len(r[:len(head)])) for r in rows[1:] if any(r[:len(head)])]
            if name == '勤務種類' and '勤務グループ' not in head:
                head.append('勤務グループ')
                body = [r + [''] for r in body]
        out[name] = dict(headers=head, rows=body)
    return out


@_out
def save_sheets(settings_b64, data_json):
    from openpyxl.styles import PatternFill
    data = json.loads(data_json)
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(settings_b64)))
    for name, t in data.items():
        if name not in EDIT_SHEETS:
            continue
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        head = t['headers']
        start = 1 if name == '基本設定' else 2
        if ws.max_row >= start:
            ws.delete_rows(start, ws.max_row - start + 1)
        if name != '基本設定':
            from copy import copy
            for j, h in enumerate(head):
                c = ws.cell(1, j + 1)
                if c.value != h:
                    c.value = h
                    src = ws.cell(1, 1)
                    if src.has_style:
                        c.font, c.fill, c.alignment, c.border = copy(src.font), copy(src.fill), copy(src.alignment), copy(src.border)
                        ws.column_dimensions[c.column_letter].width = 12
        for k, row in enumerate(t['rows']):
            r = start + k
            for j, h in enumerate(head):
                c = ws.cell(r, j + 1, _cell_in(row[j] if j < len(row) else '', h if name != '基本設定' else ''))
                if h == '日付' and c.value is not None and not isinstance(c.value, str):
                    c.number_format = 'yyyy/mm/dd'
                if h == '背景色' and c.value:
                    col = str(c.value).lstrip('#').upper()
                    c.value = col
                    if len(col) == 6:
                        c.fill = PatternFill('solid', fgColor=col)
    buf = io.BytesIO()
    wb.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()
    return dict(b64=b64, info=json.loads(info(b64)))


@_out
def recheck(req_json):
    """画面で手直しした勤務表をチェックし直し、Excel も作り直す"""
    req = json.loads(req_json)
    cfg = _cfg(req['settings'])
    month = E.to_date(req['month'])
    dates = E.month_dates(month)
    nd = len(dates)
    tbl = req.get('table') or {}
    staff = [s for s in cfg.staff if s['key'] in tbl]
    table = {}
    for s in staff:
        row = tbl[s['key']]
        for i in range(nd):
            v = str(row[i]).strip() if i < len(row) and row[i] not in (None, '') else ''
            if v:
                table[(s['key'], i)] = v
    hol_n = E.num(req.get('holidays_n'))
    if hol_n is None:
        hol_n = E.num(cfg.basic.get(E.norm('既定公休数')))
    specials = _specials(req.get('specials'))
    rq = E.day_requirements(cfg, dates, specials)
    V = E.evaluate(cfg, dates, table, rq, hol_n, staff)
    buf = io.BytesIO()
    E.write_result(cfg, buf, month, dates, table, staff, rq, hol_n, V,
                   f'{month:%Y年%m月} 勤務表（自動作成・手直しあり）', req.get('status') or '', list(dict.fromkeys(cfg.errors)), specials)
    return dict(violations=_vjson(V), targets=_targets(rq), holidays_n=hol_n,
                xlsx=base64.b64encode(buf.getvalue()).decode())
