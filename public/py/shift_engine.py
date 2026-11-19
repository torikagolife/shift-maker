# -*- coding: utf-8 -*-
"""
シフト自動作成エンジン（汎用版）

使い方:
  python shift_engine.py new   設定.xlsx 2026-12                 … 月の予定表（希望休などを書き込む用）を作る
  python shift_engine.py solve 設定.xlsx 予定_2026-12.xlsx         … 予定表をもとに勤務表を自動作成
  python shift_engine.py check 設定.xlsx 勤務表.xlsx [シート名] [--month 2026-11]
                                                                 … できあがった勤務表のルール違反をチェック

必要なもの: pip install ortools openpyxl（OR-Tools が無ければ scipy で解きます: pip install scipy openpyxl）
"""
import sys, os, re, datetime, calendar, collections, unicodedata
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WEEK = '月火水木金土日'
DAY_TYPES = ['毎日', '平日', '土曜', '日曜', '祝日', '月', '火', '水', '木', '金']
BIG = 1000          # ルール違反1件あたりのペナルティ（「できるだけ守る」ではなく「必ず守りたい」扱い）
SOFT_WISH = 20      # ?付き希望の重み（基本設定で変更可）


# ------------------------------------------------------------ 共通ヘルパー
def norm(s):
    """全角半角・空白のゆれを吸収して比較用の文字列にする"""
    if s is None:
        return ''
    if isinstance(s, float) and s.is_integer():
        s = int(s)
    s = unicodedata.normalize('NFKC', str(s)).strip()
    return re.sub(r'\s+', '', s)


def split_list(s):
    return [t for t in (norm(x) for x in re.split(r'[,、，/／\s]+', str(s or ''))) if t]


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = norm(v).replace('/', '-').replace('.', '-')
    m = re.match(r'(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$', s)
    if not m:
        return None
    return datetime.date(int(m[1]), int(m[2]), int(m[3] or 1))


def read_table(ws):
    """1行目を見出しとして、各行を dict で返す（空行は飛ばす）"""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [norm(h) for h in rows[0]]
    out = []
    for r in rows[1:]:
        if all(v is None or norm(v) == '' for v in r):
            continue
        out.append({h: v for h, v in zip(head, r) if h})
    return out


def flag(v, default=True):
    s = norm(v)
    if s == '':
        return default
    return s in ('○', '〇', 'o', 'O', '1', 'はい', 'TRUE', 'True', 'true', 'yes', '有効', 'する')


# ------------------------------------------------------------ 設定の読み込み
class Config:
    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        self.errors = []

        # 基本設定
        self.basic = {}
        if '基本設定' in wb.sheetnames:
            for r in wb['基本設定'].iter_rows(values_only=True):
                if r and r[0]:
                    self.basic[norm(r[0])] = r[1]

        # 勤務種類
        self.codes = collections.OrderedDict()
        for r in read_table(wb['勤務種類']):
            c = norm(r.get('コード'))
            if not c:
                continue
            kind = norm(r.get('種別')) or '勤務'
            tags = set(split_list(r.get('タグ')))
            tags.add(kind)
            # 勤務グループ（診察室・受付など）。グループ名はタグにもなるので必要人数やルールで使える
            grp = norm(r.get('勤務グループ'))
            if grp:
                tags.add(grp)
            else:
                grp = '休み' if kind in ('休み', '休暇') else '未分類'
            self.codes[c] = dict(
                code=c, name=r.get('名称') or c, kind=kind, tags=tags, group=grp,
                start=r.get('開始'), end=r.get('終了'),
                auto=flag(r.get('自動配置'), True),
                color=norm(r.get('背景色')).lstrip('#') or None)
        self.all_tags = set().union(*[v['tags'] for v in self.codes.values()]) if self.codes else set()
        # グループの並び順は勤務種類シートに出てくる順（休み・未分類は最後）
        g = list(dict.fromkeys(v['group'] for v in self.codes.values()))
        self.groups = [x for x in g if x not in ('未分類', '休み')] + [x for x in ('未分類', '休み') if x in g]

        # 職員
        self.staff = []
        for r in read_table(wb['職員']):
            name = r.get('氏名')
            if not norm(name):
                continue
            s = dict(no=norm(r.get('No')), name=str(name).strip(), key=norm(name),
                     groups=set(split_list(r.get('グループ'))),
                     holidays=r.get('公休数'))
            allowed = set()
            for t in split_list(r.get('使える勤務')) or ['勤務']:
                allowed |= self.resolve(t, f'職員「{s["name"]}」の使える勤務')
            s['allowed'] = allowed
            self.staff.append(s)
        self.staff_by_key = {s['key']: s for s in self.staff}
        self.staff_by_no = {s['no']: s for s in self.staff if s['no']}

        # 必要人数
        self.reqs = []
        for r in read_table(wb['必要人数']):
            dt = norm(r.get('日の種類')) or '毎日'
            target = norm(r.get('対象'))
            if not target:
                continue
            if dt not in DAY_TYPES:
                self.errors.append(f'必要人数: 日の種類「{dt}」は使えません（{"・".join(DAY_TYPES)}）')
                continue
            self.reqs.append(dict(daytype=dt, target=target, codes=self.resolve(target, '必要人数'),
                                  min=num(r.get('最小')), max=num(r.get('最大')), note=r.get('備考') or ''))

        # ルール
        self.rules = []
        for r in read_table(wb['ルール']):
            if not flag(r.get('有効'), True):
                continue
            kind = norm(r.get('ルール'))
            if not kind:
                continue
            rule = dict(kind=kind, who=r.get('対象職員'), daytype=norm(r.get('日の種類')) or '毎日',
                        t1=norm(r.get('タグ1')) or '出勤', t2=norm(r.get('タグ2')),
                        n=num(r.get('数値')), w=num(r.get('重み')) or 10, desc=r.get('説明') or '')
            rule['c1'] = self.resolve(rule['t1'], f'ルール「{kind}」')
            rule['c2'] = self.resolve(rule['t2'], f'ルール「{kind}」') if rule['t2'] else set()
            rule['staff'] = self.pick_staff(rule['who'])
            self.rules.append(rule)

        # 祝日・休診日
        self.holidays = {}
        if '祝日' in wb.sheetnames:
            for r in read_table(wb['祝日']):
                d = to_date(r.get('日付'))
                if d:
                    self.holidays[d] = r.get('名称') or '祝日'

    def resolve(self, token, where=''):
        """コード名またはタグ名 → 該当する勤務コードの集合"""
        t = norm(token)
        if t in self.codes:
            return {t}
        hit = {c for c, v in self.codes.items() if t in v['tags']}
        if not hit:
            self.errors.append(f'{where}: 「{token}」は勤務コードにもタグにもありません')
        return hit

    def pick_staff(self, who):
        """対象職員: 空欄/全員 → 全員、グループ名、No のカンマ区切り、氏名"""
        toks = split_list(who)
        if not toks or toks == ['全員']:
            return list(self.staff)
        out = []
        for s in self.staff:
            if any(t in s['groups'] or t == s['no'] or t == s['key'] for t in toks):
                out.append(s)
        if not out:
            self.errors.append(f'対象職員「{who}」に当てはまる職員がいません')
        return out

    def day_type_match(self, d, dt):
        hol = d in self.holidays
        wd = d.weekday()
        if dt == '毎日':
            return True
        if dt == '祝日':
            return hol
        if dt == '日曜':
            return wd == 6 and not hol
        if dt == '土曜':
            return wd == 5 and not hol
        if dt == '平日':
            return wd < 5 and not hol
        if dt in '月火水木金':
            return WEEK[wd] == dt and not hol
        return False


def num(v):
    if v is None or norm(v) == '':
        return None
    try:
        return int(float(norm(v)))
    except ValueError:
        return None


# ------------------------------------------------------------ 勤務表（予定表）の読み込み
def find_layout(ws):
    """日付(1,2,3…)が並んだ行と、氏名の列を探す。既存の勤務表もこの形ならそのまま読める"""
    day_row = day_col = None
    for r in range(1, 15):
        for c in range(1, 12):
            if num(ws.cell(r, c).value) == 1 and num(ws.cell(r, c + 1).value) == 2 and num(ws.cell(r, c + 2).value) == 3:
                day_row, day_col = r, c
                break
        if day_row:
            break
    if not day_row:
        raise SystemExit(f'シート「{ws.title}」で日付（1,2,3…）の行が見つかりません')
    ndays = 0
    while num(ws.cell(day_row, day_col + ndays).value) == ndays + 1:
        ndays += 1
    name_col, head_row = day_col - 1, day_row
    for r in range(1, day_row + 5):
        for c in range(1, day_col):
            if norm(ws.cell(r, c).value) in ('氏名', 'スタッフ名', '名前', '職員名'):
                name_col, head_row = c, max(head_row, r)
    return day_row, day_col, ndays, name_col, head_row


def read_schedule(ws, cfg):
    day_row, day_col, ndays, name_col, head_row = find_layout(ws)
    cells, unknown_staff, order = {}, [], []
    for r in range(head_row + 1, ws.max_row + 1):
        key = norm(ws.cell(r, name_col).value)
        if not key:
            continue
        if key.startswith('▼'):      # 自動作成で出力した表の「▼日別の人数」から下は集計なので読まない
            break
        s = cfg.staff_by_key.get(key)
        if not s:
            unknown_staff.append(str(ws.cell(r, name_col).value))
            continue
        order.append(s)
        for i in range(ndays):
            v = ws.cell(r, day_col + i).value
            if v is not None and norm(v) != '':
                cells[(s['key'], i)] = str(v).strip()
    return cells, ndays, unknown_staff, order


def month_of(ws, given=None):
    if given:
        return to_date(given)
    for coord in ('B1', 'C1', 'A1'):
        d = to_date(ws[coord].value)
        if d:
            return d.replace(day=1)
    m = re.search(r'(\d{4})\D?(\d{1,2})', ws.title)
    if m:
        return datetime.date(int(m[1]), int(m[2]), 1)
    raise SystemExit('対象年月がわかりません。--month 2026-11 のように指定してください')


def read_month_extras(wb, ws):
    """予定ファイルの公休数（E1）と、特別日シート"""
    hol_n = num(ws['E1'].value)
    specials = []
    if '特別日' in wb.sheetnames:
        for r in read_table(wb['特別日']):
            d = to_date(r.get('日付'))
            if d and norm(r.get('対象')):
                specials.append(dict(date=d, target=norm(r.get('対象')),
                                     min=num(r.get('最小')), max=num(r.get('最大')), note=r.get('備考') or ''))
    return hol_n, specials


# ------------------------------------------------------------ その日の必要人数を決める
def day_requirements(cfg, dates, specials):
    """{日付index: [(対象名, コード集合, 最小, 最大, 備考)]}。特別日の行は同じ対象の通常行より優先"""
    out = {}
    for i, d in enumerate(dates):
        sp = [x for x in specials if x['date'] == d]
        sp_targets = {x['target'] for x in sp}
        lst = []
        for q in cfg.reqs:
            if q['target'] in sp_targets:
                continue
            if cfg.day_type_match(d, q['daytype']):
                lst.append((q['target'], q['codes'], q['min'], q['max'], q['note']))
        for x in sp:
            lst.append((x['target'], cfg.resolve(x['target'], '特別日'), x['min'], x['max'], x['note'] or '特別日'))
        out[i] = lst
    return out


def able(s, codes):
    """その職員の「使える勤務」に codes のどれかが入っているか"""
    return bool(s['allowed'] & codes)


def holidays_for(cfg, s, default_n):
    n = num(s['holidays'])
    return n if n is not None else default_n


# ------------------------------------------------------------ チェック（できあがった表を評価）
def evaluate(cfg, dates, table, req_by_day, hol_n, staff_list):
    """table: {(職員key, 日index): コード}。違反の一覧を返す"""
    V = []
    def add(kind, i, s, msg, target=None):
        V.append(dict(kind=kind, date=dates[i] if i is not None else None, staff=s['name'] if s else '', msg=msg, target=target))
    nd = len(dates)
    has = lambda s, i, cs: table.get((s['key'], i)) in cs

    for s in staff_list:
        for i in range(nd):
            c = table.get((s['key'], i))
            if c is None:
                add('空欄', i, s, '勤務が入っていません')
            elif norm(c) not in cfg.codes:
                add('未登録コード', i, s, f'「{c}」は勤務種類にありません（入力ミス？）')
            elif cfg.codes[norm(c)]['kind'] == '勤務' and norm(c) not in s['allowed']:
                add('担当外', i, s, f'「{c}」はこの職員の使える勤務に入っていません')
        n = holidays_for(cfg, s, hol_n)
        if n is not None:
            got = sum(1 for i in range(nd) if cfg.codes.get(norm(table.get((s['key'], i)) or ''), {}).get('tags', set()) & {'公休'})
            if got != n:
                add('公休数', None, s, f'公休 {got}日（予定 {n}日）')

    for i in range(nd):
        for (tname, cs, mn, mx, note) in req_by_day[i]:
            cnt = sum(1 for s in staff_list if norm(table.get((s['key'], i)) or '') in cs)
            if mn is not None and cnt < mn:
                add('人数不足', i, None, f'{tname} {cnt}人（最小 {mn}人）{note}', tname)
            if mx is not None and cnt > mx:
                add('人数超過', i, None, f'{tname} {cnt}人（最大 {mx}人）{note}', tname)

    T = {k: norm(v) for k, v in table.items()}
    for r in cfg.rules:
        k, c1, c2, n = r['kind'], r['c1'], r['c2'], r['n']
        days = [i for i, d in enumerate(dates) if cfg.day_type_match(d, r['daytype'])]
        if k == '連勤上限' and n:
            for s in r['staff']:
                run = 0
                for i in range(nd):
                    run = run + 1 if T.get((s['key'], i)) in c1 else 0
                    if run == n + 1:
                        add('連勤', i, s, f'{r["t1"]}が{n + 1}日以上連続（上限 {n}日）')
        elif k == '翌日禁止':
            for s in r['staff']:
                for i in range(nd - 1):
                    if T.get((s['key'], i)) in c1 and T.get((s['key'], i + 1)) in c2:
                        add('並び', i + 1, s, f'{r["t1"]}の翌日に{r["t2"]}')
        elif k in ('月回数上限', '月回数下限') and n is not None:
            for s in r['staff']:
                cnt = sum(1 for i in days if T.get((s['key'], i)) in c1)
                if k == '月回数上限' and cnt > n:
                    add('回数', None, s, f'{r["t1"]}（{r["daytype"]}）{cnt}回（上限 {n}回）')
                if k == '月回数下限' and cnt < n:
                    add('回数', None, s, f'{r["t1"]}（{r["daytype"]}）{cnt}回（下限 {n}回）')
        elif k == '同日上限' and n is not None:
            for i in days:
                who = [s['name'] for s in r['staff'] if T.get((s['key'], i)) in c1]
                if len(who) > n:
                    add('同日', i, None, f'{r["t1"]}: {"・".join(who)}（同じ日は{n}人まで）')
        elif k == '均等化':
            cnts = [sum(1 for i in days if T.get((s['key'], i)) in c1) for s in r['staff'] if able(s, c1)]
            if cnts and max(cnts) - min(cnts) > (n or 1):
                add('かたより', None, None, f'{r["t1"]}（{r["daytype"]}）の回数が {min(cnts)}〜{max(cnts)}回でばらついています')
    return V


# ------------------------------------------------------------ 自動作成（OR-Tools CP-SAT）
def solve(cfg, dates, cells, req_by_day, hol_n, staff_list, seconds):
    try:
        from ortools.sat.python import cp_model
    except ImportError:
        # OR-Tools が無い環境（ブラウザの Pyodide など）は scipy の HiGHS で解く
        import solver_highs
        return solver_highs.solve(cfg, dates, cells, req_by_day, hol_n, staff_list, seconds)
    m = cp_model.CpModel()
    nd = len(dates)
    rest = [c for c, v in cfg.codes.items() if v['kind'] == '休み' and v['auto']]
    if not rest:
        raise SystemExit('勤務種類に「種別=休み」で自動配置できるコード（例: 休）が必要です')
    wish_w = num(cfg.basic.get(norm('希望の重み'))) or SOFT_WISH

    X = {}       # (key, i, code) -> BoolVar
    obj = []
    notes = []
    for s in staff_list:
        base = (s['allowed'] & {c for c, v in cfg.codes.items() if v['auto']}) | set(rest)
        for i in range(nd):
            raw = cells.get((s['key'], i))
            cand, prefer = set(base), None
            if raw:
                t = norm(raw)
                if t[0] in '×xX✕' and len(t) > 1:
                    cand -= cfg.resolve(t[1:], f'{s["name"]} {i + 1}日')
                elif t[0] in '?？' and len(t) > 1:
                    prefer = cfg.resolve(t[1:], f'{s["name"]} {i + 1}日')
                else:
                    fixed = set()
                    for part in re.split(r'[/／]', t):
                        fixed |= cfg.resolve(part, f'{s["name"]} {i + 1}日')
                    cand = fixed
            if not cand:
                notes.append(f'{s["name"]} {i + 1}日: 入れられる勤務がありません')
                cand = set(rest)
            vs = []
            for c in cand:
                X[(s['key'], i, c)] = v = m.NewBoolVar('')
                vs.append(v)
            m.AddExactlyOne(vs)
            if prefer:
                hit = [X[(s['key'], i, c)] for c in prefer if (s['key'], i, c) in X]
                if hit:
                    miss = m.NewBoolVar('')
                    m.Add(sum(hit) == 0).OnlyEnforceIf(miss)
                    m.Add(sum(hit) >= 1).OnlyEnforceIf(miss.Not())
                    obj.append(wish_w * miss)

    def cnt(s, i, cs):
        return sum(X[(s['key'], i, c)] for c in cs if (s['key'], i, c) in X)

    def soft_le(expr, ub, w=BIG):
        sl = m.NewIntVar(0, 1000, '')
        m.Add(expr <= ub + sl)
        obj.append(w * sl)

    def soft_ge(expr, lb, w=BIG):
        sl = m.NewIntVar(0, 1000, '')
        m.Add(expr + sl >= lb)
        obj.append(w * sl)

    # 公休数
    kokyu = {c for c, v in cfg.codes.items() if '公休' in v['tags']}
    for s in staff_list:
        n = holidays_for(cfg, s, hol_n)
        if n is not None:
            e = sum(cnt(s, i, kokyu) for i in range(nd))
            soft_le(e, n); soft_ge(e, n)

    # 必要人数
    for i in range(nd):
        for (tname, cs, mn, mx, note) in req_by_day[i]:
            e = sum(cnt(s, i, cs) for s in staff_list)
            if mn is not None:
                soft_ge(e, mn)
            if mx is not None:
                soft_le(e, mx)

    # ルール
    for r in cfg.rules:
        k, c1, c2, n = r['kind'], r['c1'], r['c2'], r['n']
        days = [i for i, d in enumerate(dates) if cfg.day_type_match(d, r['daytype'])]
        ss = [s for s in r['staff'] if s in staff_list]
        if k == '連勤上限' and n:
            for s in ss:
                for i in range(nd - n):
                    soft_le(sum(cnt(s, j, c1) for j in range(i, i + n + 1)), n)
        elif k == '翌日禁止':
            for s in ss:
                for i in range(nd - 1):
                    soft_le(cnt(s, i, c1) + cnt(s, i + 1, c2), 1)
        elif k == '月回数上限' and n is not None:
            for s in ss:
                soft_le(sum(cnt(s, i, c1) for i in days), n)
        elif k == '月回数下限' and n is not None:
            for s in ss:
                soft_ge(sum(cnt(s, i, c1) for i in days), n)
        elif k == '同日上限' and n is not None:
            for i in days:
                soft_le(sum(cnt(s, i, c1) for s in ss), n)
        elif k == '均等化' and [s for s in ss if able(s, c1)]:
            ss = [s for s in ss if able(s, c1)]
            hi, lo = m.NewIntVar(0, nd, ''), m.NewIntVar(0, nd, '')
            for s in ss:
                e = sum(cnt(s, i, c1) for i in days)
                m.Add(e <= hi); m.Add(e >= lo)
            obj.append(r['w'] * (hi - lo))
        elif k not in ('連勤上限', '翌日禁止', '月回数上限', '月回数下限', '同日上限', '均等化'):
            notes.append(f'ルール「{k}」は未対応のため無視しました')

    m.Minimize(sum(obj))
    sv = cp_model.CpSolver()
    sv.parameters.max_time_in_seconds = seconds
    sv.parameters.num_workers = int(os.environ.get('SHIFT_WORKERS', '8'))
    st = sv.Solve(m)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise SystemExit('解が見つかりませんでした。予定表で固定した勤務どうしが矛盾していないか確認してください')
    table = {}
    for (k, i, c), v in X.items():
        if sv.Value(v):
            table[(k, i)] = c
    status = '最適解' if st == cp_model.OPTIMAL else f'時間内の最良解（{seconds}秒で打ち切り）'
    return table, status, notes


# ------------------------------------------------------------ Excel出力
THIN = Side(style='thin', color='BBBBBB')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED = PatternFill('solid', fgColor='FF9999')
GRAY = PatternFill('solid', fgColor='E0E0E0')


def write_result(cfg, path, month, dates, table, staff_list, req_by_day, hol_n, V, title, status='', notes=(), specials=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '勤務表'
    nd = len(dates)
    ws['A1'] = '対象年月'; ws['B1'] = month; ws['B1'].number_format = 'yyyy"年"m"月"'
    ws['D1'] = '公休数'; ws['E1'] = hol_n
    ws['G1'] = title
    for i, d in enumerate(dates):
        col = 3 + i
        ws.cell(3, col, i + 1); ws.cell(4, col, WEEK[d.weekday()])
        if d in cfg.holidays:
            ws.cell(5, col, cfg.holidays[d])
        for r in (3, 4, 5):
            ws.cell(r, col).alignment = Alignment(horizontal='center', shrink_to_fit=True)
            if d.weekday() >= 5 or d in cfg.holidays:
                ws.cell(r, col).fill = GRAY
        ws.column_dimensions[get_column_letter(col)].width = 5.2
    ws.cell(5, 1, 'No'); ws.cell(5, 2, '氏名')
    ws.column_dimensions['B'].width = 14

    bad_cell = {(v['staff'], v['date']) for v in V if v['staff'] and v['date']}
    for r, s in enumerate(staff_list, start=6):
        ws.cell(r, 1, s['no']); ws.cell(r, 2, s['name'])
        for i, d in enumerate(dates):
            c = table.get((s['key'], i))
            cell = ws.cell(r, 3 + i, c)
            cell.alignment = Alignment(horizontal='center', shrink_to_fit=True)
            cell.border = BOX
            info = cfg.codes.get(norm(c or ''))
            if (s['name'], d) in bad_cell:
                cell.fill = RED
            elif info and info['color']:
                cell.fill = PatternFill('solid', fgColor=info['color'])

    # 日別の人数（必要人数の対象ごと）
    r0 = 6 + len(staff_list) + 1
    ws.cell(r0, 2, '▼日別の人数').font = Font(bold=True)
    targets = list(dict.fromkeys(t for i in range(nd) for (t, *_) in req_by_day[i]))
    for k, t in enumerate(targets):
        r = r0 + 1 + k
        ws.cell(r, 2, t)
        for i in range(nd):
            hit = [q for q in req_by_day[i] if q[0] == t]
            cs = hit[0][1] if hit else cfg.resolve(t)
            n = sum(1 for s in staff_list if norm(table.get((s['key'], i)) or '') in cs)
            cell = ws.cell(r, 3 + i, n)
            cell.alignment = Alignment(horizontal='center'); cell.border = BOX
            for (_, _, mn, mx, _) in hit:
                if (mn is not None and n < mn) or (mx is not None and n > mx):
                    cell.fill = RED

    # 個人別の回数（公休＋ルールで使うタグ）
    cols = ['公休'] + list(dict.fromkeys(r['t1'] for r in cfg.rules if r['kind'] in ('均等化', '月回数上限', '月回数下限')))
    c0 = 3 + nd + 1
    for j, t in enumerate(cols):
        ws.cell(5, c0 + j, t).alignment = Alignment(horizontal='center', shrink_to_fit=True)
        ws.column_dimensions[get_column_letter(c0 + j)].width = 6
        cs = cfg.resolve(t)
        for r, s in enumerate(staff_list, start=6):
            ws.cell(r, c0 + j, sum(1 for i in range(nd) if norm(table.get((s['key'], i)) or '') in cs)).alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'C6'

    write_group_sheet(cfg, wb, month, dates, table, staff_list, title)

    # チェック結果
    ck = wb.create_sheet('チェック')
    ck['A1'] = title; ck['A1'].font = Font(bold=True, size=13)
    ck['A2'] = status
    row = 3
    for n_ in notes:
        ck.cell(row, 1, '注意'); ck.cell(row, 4, n_); row += 1
    row += 1
    for j, h in enumerate(['種類', '日付', '職員', '内容']):
        ck.cell(row, 1 + j, h).font = Font(bold=True)
    if not V:
        ck.cell(row + 1, 1, '違反なし')
    for k, v in enumerate(sorted(V, key=lambda v: (v['date'] or datetime.date.min, v['kind'])), start=row + 1):
        ck.cell(k, 1, v['kind'])
        if v['date']:
            ck.cell(k, 2, f'{v["date"].month}/{v["date"].day}({WEEK[v["date"].weekday()]})')
        ck.cell(k, 3, v['staff']); ck.cell(k, 4, v['msg'])
    for col, w in zip('ABCD', (12, 10, 16, 70)):
        ck.column_dimensions[col].width = w

    # 特別日も一緒に保存しておく（手直し後に check し直すとき用）
    sp = wb.create_sheet('特別日')
    for j, h in enumerate(['日付', '対象', '最小', '最大', '備考']):
        sp.cell(1, 1 + j, h).font = Font(bold=True)
    for x in specials:
        sp.append([x['date'], x['target'], x['min'], x['max'], x['note']])
        sp.cell(sp.max_row, 1).number_format = 'yyyy/mm/dd'
    sp.column_dimensions['A'].width = 12
    wb.save(path)


def write_group_sheet(cfg, wb, month, dates, table, staff_list, title):
    """グループ別の勤務表: 日付ごとに、各グループに誰が入っているか"""
    ws = wb.create_sheet('グループ別', 1)
    nd = len(dates)
    ws['A1'] = title.replace('勤務表', '勤務表（グループ別）', 1) if '勤務表' in title else title
    ws['A1'].font = Font(bold=True, size=13)
    ws.column_dimensions['A'].width = 12
    for i, d in enumerate(dates):
        col = 2 + i
        ws.cell(3, col, i + 1); ws.cell(4, col, WEEK[d.weekday()])
        for r in (3, 4):
            c = ws.cell(r, col); c.alignment = Alignment(horizontal='center'); c.border = BOX
            if d.weekday() >= 5 or d in cfg.holidays:
                c.fill = GRAY
        ws.column_dimensions[get_column_letter(col)].width = 11
    ws.cell(3, 1, '日'); ws.cell(4, 1, '曜日')
    # 日ごと・グループごとの [(職員, コード)]
    cell_of = {}
    for s in staff_list:
        for i in range(nd):
            c = norm(table.get((s['key'], i)) or '')
            info = cfg.codes.get(c)
            g = info['group'] if info else ('未登録' if c else None)
            if g:
                cell_of.setdefault((g, i), []).append((s, c))
    groups = cfg.groups + (['未登録'] if any(k[0] == '未登録' for k in cell_of) else [])
    HEAD = PatternFill('solid', fgColor='2F5D8A')
    r = 5
    for g in groups:
        n = max([len(cell_of.get((g, i), [])) for i in range(nd)] or [0])
        if n == 0:
            continue
        ws.cell(r, 1, g).font = Font(bold=True, color='FFFFFF')
        ws.cell(r, 1).fill = HEAD
        for i in range(nd):
            c = ws.cell(r, 2 + i, len(cell_of.get((g, i), [])))
            c.font = Font(bold=True, color='FFFFFF'); c.fill = HEAD; c.alignment = Alignment(horizontal='center')
        for k in range(n):
            rr = r + 1 + k
            ws.cell(rr, 1).border = BOX
            for i in range(nd):
                lst = cell_of.get((g, i), [])
                c = ws.cell(rr, 2 + i)
                c.border = BOX
                c.alignment = Alignment(horizontal='left', shrink_to_fit=True)
                if k < len(lst):
                    s, code = lst[k]
                    c.value = f'{s["name"].replace(chr(12288), "").replace(" ", "")} {code}'
                    info = cfg.codes.get(code)
                    if info and info['color']:
                        c.fill = PatternFill('solid', fgColor=info['color'])
        r += n + 2
    ws.freeze_panes = 'B5'
    ws.page_setup.orientation = 'landscape'


# ------------------------------------------------------------ コマンド
def month_dates(month):
    n = calendar.monthrange(month.year, month.month)[1]
    return [datetime.date(month.year, month.month, d) for d in range(1, n + 1)]


def cmd_new(cfg, ym):
    month = to_date(ym)
    dates = month_dates(month)
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = '予定'
    ws['A1'] = '対象年月'; ws['B1'] = month; ws['B1'].number_format = 'yyyy"年"m"月"'
    ws['D1'] = '公休数'; ws['E1'] = num(cfg.basic.get(norm('既定公休数')))
    ws['G1'] = '書き方: コード=固定（例 休・有・B）／ A/AS=どちらか／ ×B=Bにしない／ ?休=できれば休み（希望）／ 空欄=おまかせ'
    for i, d in enumerate(dates):
        col = 3 + i
        ws.cell(3, col, i + 1); ws.cell(4, col, WEEK[d.weekday()])
        if d in cfg.holidays:
            ws.cell(5, col, cfg.holidays[d])
        ws.column_dimensions[get_column_letter(col)].width = 5.2
        for r in (3, 4, 5):
            ws.cell(r, col).alignment = Alignment(horizontal='center', shrink_to_fit=True)
            if d.weekday() >= 5 or d in cfg.holidays:
                ws.cell(r, col).fill = GRAY
    ws.cell(5, 1, 'No'); ws.cell(5, 2, '氏名'); ws.column_dimensions['B'].width = 14
    for r, s in enumerate(cfg.staff, start=6):
        ws.cell(r, 1, s['no']); ws.cell(r, 2, s['name'])
        for i in range(len(dates)):
            ws.cell(r, 3 + i).border = BOX
            ws.cell(r, 3 + i).alignment = Alignment(horizontal='center', shrink_to_fit=True)
    ws.freeze_panes = 'C6'
    sp = wb.create_sheet('特別日')
    for j, h in enumerate(['日付', '対象', '最小', '最大', '備考']):
        sp.cell(1, 1 + j, h).font = Font(bold=True)
    sp['F1'] = '← その日だけ必要人数を変えたいとき（同じ対象の通常の必要人数より優先）'
    sp.column_dimensions['A'].width = 12
    out = f'予定_{month:%Y-%m}.xlsx'
    wb.save(out)
    print(f'作成しました: {out}')


def run_solve(cfg, month, cells, hol_n, specials, seconds, out, staff_list=None, unknown=(), log=print):
    """自動作成の本体。cells: {(職員key, 日index): 予定表の書き込み}。out はパスか BytesIO"""
    dates = month_dates(month)
    if hol_n is None:
        hol_n = num(cfg.basic.get(norm('既定公休数')))
    staff_list = staff_list or cfg.staff
    req = day_requirements(cfg, dates, specials)
    sec = seconds or num(cfg.basic.get(norm('計算時間（秒）'))) or 60
    log(f'{month:%Y年%m月}  職員{len(staff_list)}人  計算中（最大{sec}秒）…')
    table, status, notes = solve(cfg, dates, cells, req, hol_n, staff_list, sec)
    notes = list(dict.fromkeys(list(cfg.errors) + notes + [f'職員シートにいない名前: {u}' for u in unknown]))
    V = evaluate(cfg, dates, table, req, hol_n, staff_list)
    write_result(cfg, out, month, dates, table, staff_list, req, hol_n, V,
                 f'{month:%Y年%m月} 勤務表（自動作成）', status, notes, specials)
    return dict(status=status, notes=notes, violations=V, table=table, staff=staff_list, dates=dates, hol_n=hol_n)


def cmd_solve(cfg, path, seconds=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['予定'] if '予定' in wb.sheetnames else wb.active
    month = month_of(ws)
    cells, nd, unknown, order = read_schedule(ws, cfg)
    if nd != len(month_dates(month)):
        raise SystemExit(f'日付の列が {nd} 日分しかありません（{month:%Y年%m月}は {len(month_dates(month))} 日）')
    hol_n, specials = read_month_extras(wb, ws)
    if cfg.errors:
        print('設定の確認:'); [print('  -', e) for e in dict.fromkeys(cfg.errors)]
    out = os.path.join(os.path.dirname(os.path.abspath(path)), f'勤務表_{month:%Y-%m}_自動作成.xlsx')
    r = run_solve(cfg, month, cells, hol_n, specials, seconds, out, order, unknown)
    print(f'{r["status"]}  残った違反 {len(r["violations"])}件')
    print(f'作成しました: {out}')


def run_check(cfg, wb, sheet=None, ym=None, out=None, filename=''):
    """できあがった勤務表（openpyxl Workbook）をチェック"""
    ws = wb[sheet] if sheet else (wb['勤務表'] if '勤務表' in wb.sheetnames else wb.active)
    month = month_of(ws, ym)
    cells, nd, unknown, order = read_schedule(ws, cfg)
    dates = month_dates(month)[:nd]
    hol_n = num(ws['E1'].value) if norm(ws['D1'].value) == '公休数' else None
    specials = read_month_extras(wb, ws)[1] if '特別日' in wb.sheetnames else []
    req = day_requirements(cfg, dates, specials)
    V = evaluate(cfg, dates, cells, req, hol_n, order)
    notes = list(dict.fromkeys(list(cfg.errors) + [f'職員シートにいない名前（チェック対象外）: {u}' for u in unknown]))
    if out is not None:
        write_result(cfg, out, month, dates, cells, order, req, hol_n, V,
                     f'{month:%Y年%m月} 勤務表チェック（{filename} / {ws.title}）', '', notes, specials)
    return dict(month=month, violations=V, notes=notes, staff=order, dates=dates, table=cells)


def cmd_check(cfg, path, sheet=None, ym=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    probe = run_check(cfg, wb, sheet, ym)
    out = os.path.join(os.path.dirname(os.path.abspath(path)), f'チェック結果_{probe["month"]:%Y-%m}.xlsx')
    r = run_check(cfg, wb, sheet, ym, out, os.path.basename(path))
    kinds = collections.Counter(v['kind'] for v in r['violations'])
    print(f'違反 {len(r["violations"])}件: ' + '、'.join(f'{k}{n}件' for k, n in kinds.most_common()))
    print(f'作成しました: {out}')


def main(argv):
    if len(argv) < 3:
        print(__doc__); return
    cmd, cfg_path = argv[0], argv[1]
    cfg = Config(cfg_path)
    if cmd == 'new':
        cmd_new(cfg, argv[2])
    elif cmd == 'solve':
        sec = None
        if '--sec' in argv:
            sec = int(argv[argv.index('--sec') + 1])
        cmd_solve(cfg, argv[2], sec)
    elif cmd == 'check':
        rest = argv[3:]
        ym = None
        if '--month' in rest:
            k = rest.index('--month'); ym = rest[k + 1]; rest = rest[:k] + rest[k + 2:]
        cmd_check(cfg, argv[2], rest[0] if rest else None, ym)
    else:
        print(__doc__)


if __name__ == '__main__':
    main(sys.argv[1:])
