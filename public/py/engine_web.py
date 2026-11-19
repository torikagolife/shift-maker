# -*- coding: utf-8 -*-
"""ブラウザ（Pyodide）から呼ぶための入口。引数も戻り値も JSON 文字列"""
import base64, io, json
import openpyxl
import shift_engine as E


def _cfg(b64):
    return E.Config(io.BytesIO(base64.b64decode(b64)))


def _vjson(V):
    return [dict(kind=v['kind'], date=v['date'].isoformat() if v['date'] else None,
                 staff=v['staff'], msg=v['msg']) for v in V]


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
                    tags=sorted(v['tags'])) for c, v in cfg.codes.items()],
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
    specials = []
    for x in req.get('specials', []):
        d = E.to_date(x.get('date'))
        if d and E.norm(x.get('target')):
            specials.append(dict(date=d, target=E.norm(x['target']), min=E.num(x.get('min')),
                                 max=E.num(x.get('max')), note=x.get('note') or ''))
    buf = io.BytesIO()
    r = E.run_solve(cfg, month, cells, E.num(req.get('holidays_n')), specials,
                    E.num(req.get('seconds')), buf, log=lambda *a: None)
    return dict(
        status=r['status'], notes=r['notes'], violations=_vjson(r['violations']), holidays_n=r['hol_n'],
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
