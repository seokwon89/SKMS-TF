"""사업보고서 텍스트에서 '직원 등 현황' 표의 성별 행과 합계 행을 파싱하고 산술 검증한다."""
import re, sys, glob, os, json
def n(t): return 0 if t in ('-', '') else int(t.replace(',', ''))
out = {}
for f in sorted(glob.glob(sys.argv[1] + '/[[]*.txt')):
    L = open(f, encoding='utf-8').read().split('\n')
    s = [i for i, l in enumerate(L) if re.search(r'직원 등 현황\s*$', l)][-1]
    rows, total, base = [], None, None
    for i in range(s, s + 120):
        l = L[i]
        if '기준일' in l and base is None:
            base = re.sub(r'\s+', '', l)
        m = re.match(r'^\s*(.*?)\s+(남|여)\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s', l)
        if m:
            rows.append(dict(seg=m.group(1).strip(), sex=m.group(2), reg=n(m.group(3)), reg_pt=n(m.group(4)), tmp=n(m.group(5)), tmp_pt=n(m.group(6)), sum=n(m.group(7)), line=i + 1))
            continue
        m = re.match(r'^\s*합계\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s+([\d,\-]+)\s', l)
        if m:
            total = dict(reg=n(m.group(1)), reg_pt=n(m.group(2)), tmp=n(m.group(3)), tmp_pt=n(m.group(4)), sum=n(m.group(5)), line=i + 1)
            break
    co = os.path.basename(f).split(']')[0].strip('[')
    chk = []
    for r in rows:
        chk.append(r['reg'] + r['tmp'] == r['sum'])
    s_reg, s_tmp, s_sum = (sum(r[k] for r in rows) for k in ('reg', 'tmp', 'sum'))
    ok = all(chk) and total and (s_reg, s_tmp, s_sum) == (total['reg'], total['tmp'], total['sum']) and total['reg'] + total['tmp'] == total['sum']
    out[co] = dict(base=base, rows=rows, total=total, arithmetic_ok=bool(ok))
    print(f"{co:12s} {base} | 정규 {total['reg']:>6,} 기간제 {total['tmp']:>5,} 합계 {total['sum']:>6,} | 행합 {s_sum:,} | 산술검증 {'OK' if ok else 'FAIL'} | 행 {len(rows)}")
json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
