import json, io
import numpy as np
import openpyxl
book = None
level = 'basic'

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def initialize(raw, lv='basic'):
    global book, codes, names, L, M, V, sheets, valid, level
    level = lv
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    base = book['국산투입계수(Ad)']
    codes = [str(c) for c in base[4][2:] if c is not None and str(c).isdigit() and int(str(c)) < 9000]
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('부문 코드를 읽지 못했습니다. 시트 구조를 확인하세요.')
    if lv == 'basic' and len(codes) != 380:
        raise ValueError('기본부문은 380개여야 하는데 ' + str(len(codes)) + '개를 읽었습니다.')
    nm = {str(r[0]): str(r[1]) for r in base[6:] if r}
    names = [nm[c] for c in codes]

    def matrix(s):
        rows = book[s]
        ri = {str(r[0]): i for i, r in enumerate(rows) if r}
        ci = {str(c): i for i, c in enumerate(rows[4])}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a

    L = matrix('생산유발계수')
    M = matrix('수입유발계수')
    V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)')
    am = matrix('수입투입계수(Am)')

    arows = book['총투입계수(A)']
    aci = {str(c): i for i, c in enumerate(arows[4])}
    vr = next(r for r in arows if r and (str(r[0]) == '9690' or str(r[1]).strip() == '부가가치계'))
    v = np.array([vr[aci[c]] for c in codes], dtype=float)

    n = len(codes)
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    return emit({'sectors': [{'code': c, 'name': n_} for c, n_ in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals})

def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if str(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]),
                 'unit': str(rows[3][0] or '係数').replace('係数', '계수 / 비율: 행·열 명칭 확인'), 'cells': cells})

def simulate(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload)
    f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0:
            raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0:
        raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f
    m = M @ f
    v = V @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})

def linkage():
    mean = L.sum() / len(codes)
    back = L.sum(axis=0)
    fwd = L.sum(axis=1)
    rows = [{'code': c, 'name': n_, 'influence': float(b / mean), 'sensitivity': float(f / mean),
             'key': bool(b / mean > 1 and f / mean > 1)}
            for c, n_, b, f in zip(codes, names, back, fwd)]
    return emit({'level': level, 'rows': rows})
