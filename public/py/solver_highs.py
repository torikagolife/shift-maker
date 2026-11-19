# -*- coding: utf-8 -*-
"""OR-Tools の代わりに scipy（HiGHS）で解く版。ブラウザ（Pyodide）でも動く"""
import re
import numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds
from scipy.sparse import coo_matrix
from shift_engine import norm, num, able, holidays_for, BIG, SOFT_WISH


class LP:
    def __init__(self):
        self.lb, self.ub, self.integ, self.c = [], [], [], []
        self.rows, self.rlo, self.rhi = [], [], []

    def var(self, lb=0, ub=1, cost=0):
        self.lb.append(lb); self.ub.append(ub); self.integ.append(1); self.c.append(cost)
        return len(self.lb) - 1

    def add(self, terms, lo=-np.inf, hi=np.inf):
        """terms: [(var, coef), ...]"""
        self.rows.append(terms); self.rlo.append(lo); self.rhi.append(hi)

    def le(self, terms, ub, w=BIG):     # できるだけ ≦ ub（超えた分にペナルティ）
        s = self.var(0, 1000, w); self.add(terms + [(s, -1)], hi=ub)

    def ge(self, terms, lb, w=BIG):
        s = self.var(0, 1000, w); self.add(terms + [(s, 1)], lo=lb)

    def solve(self, seconds):
        r, cidx, val = [], [], []
        for i, terms in enumerate(self.rows):
            for v, k in terms:
                r.append(i); cidx.append(v); val.append(k)
        A = coo_matrix((val, (r, cidx)), shape=(len(self.rows), len(self.lb))).tocsr()
        res = milp(c=np.array(self.c, float), integrality=np.array(self.integ),
                   bounds=Bounds(np.array(self.lb, float), np.array(self.ub, float)),
                   constraints=LinearConstraint(A, np.array(self.rlo, float), np.array(self.rhi, float)),
                   options={'time_limit': seconds, 'mip_rel_gap': 0.0})
        return res


def solve(cfg, dates, cells, req_by_day, hol_n, staff_list, seconds):
    lp = LP()
    nd = len(dates)
    rest = [c for c, v in cfg.codes.items() if v['kind'] == '休み' and v['auto']]
    if not rest:
        raise SystemExit('勤務種類に「種別=休み」で自動配置できるコード（例: 休）が必要です')
    wish_w = num(cfg.basic.get(norm('希望の重み'))) or SOFT_WISH
    X, notes = {}, []
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
            for c in sorted(cand):
                X[(s['key'], i, c)] = lp.var()
            lp.add([(X[(s['key'], i, c)], 1) for c in cand], 1, 1)       # 1日1つ
            if prefer:
                hit = [X[(s['key'], i, c)] for c in prefer if (s['key'], i, c) in X]
                if hit:
                    miss = lp.var(0, 1, wish_w)
                    lp.add([(miss, 1)] + [(h, 1) for h in hit], lo=1)

    def cnt(s, i, cs):
        return [(X[(s['key'], i, c)], 1) for c in cs if (s['key'], i, c) in X]

    kokyu = {c for c, v in cfg.codes.items() if '公休' in v['tags']}
    for s in staff_list:
        n = holidays_for(cfg, s, hol_n)
        if n is not None:
            e = [t for i in range(nd) for t in cnt(s, i, kokyu)]
            lp.le(e, n); lp.ge(e, n)

    for i in range(nd):
        for (tname, cs, mn, mx, note) in req_by_day[i]:
            e = [t for s in staff_list for t in cnt(s, i, cs)]
            if mn is not None: lp.ge(e, mn)
            if mx is not None: lp.le(e, mx)

    for r in cfg.rules:
        k, c1, c2, n = r['kind'], r['c1'], r['c2'], r['n']
        days = [i for i, d in enumerate(dates) if cfg.day_type_match(d, r['daytype'])]
        ss = [s for s in r['staff'] if s in staff_list]
        if k == '連勤上限' and n:
            for s in ss:
                for i in range(nd - n):
                    lp.le([t for j in range(i, i + n + 1) for t in cnt(s, j, c1)], n)
        elif k == '翌日禁止':
            for s in ss:
                for i in range(nd - 1):
                    lp.le(cnt(s, i, c1) + cnt(s, i + 1, c2), 1)
        elif k == '月回数上限' and n is not None:
            for s in ss:
                lp.le([t for i in days for t in cnt(s, i, c1)], n)
        elif k == '月回数下限' and n is not None:
            for s in ss:
                lp.ge([t for i in days for t in cnt(s, i, c1)], n)
        elif k == '同日上限' and n is not None:
            for i in days:
                lp.le([t for s in ss for t in cnt(s, i, c1)], n)
        elif k == '均等化':
            ss = [s for s in ss if able(s, c1)]
            if ss:
                hi, lo = lp.var(0, nd, r['w']), lp.var(0, nd, -r['w'])
                for s in ss:
                    e = [t for i in days for t in cnt(s, i, c1)]
                    lp.add(e + [(hi, -1)], hi=0)
                    lp.add(e + [(lo, -1)], lo=0)
        else:
            notes.append(f'ルール「{k}」は未対応のため無視しました')

    res = lp.solve(seconds)
    if res.x is None:
        raise SystemExit('解が見つかりませんでした（' + str(res.message) + '）')
    table = {(k, i): c for (k, i, c), v in X.items() if res.x[v] > 0.5}
    status = '最適解' if res.status == 0 else f'時間内の最良解（{seconds}秒で打ち切り）'
    return table, status, notes
