"""사업보고서 텍스트에서 기준일(2025.12.31) 임원 현황 행만 파싱한다.
기준일 이후 변동·선임후보·겸직 표는 제외한다."""
import re, sys, glob, json, os
pat = re.compile(r'^\s*([가-힣]{2,5}[a-z]?|[A-Za-z][A-Za-z .\-]+?)\*?\s+(남|여)(?:성)?\s+(\d{4})[.\-년]\s?(\d{1,2})월?\s+(?:(\S+(?:\s\S+)?)\s+)??(사내이사|사외이사|기타비상무이사|미등기|감사위원|상근감사|감사)\s+(상근|비상근)')
OFF = re.compile(r'이후|선임 후보자|해임 대상자|변동|퇴임한|신규로 선임|겸직')
ON = re.compile(r'임원\s?(의\s)?현황')
out = {}
for f in sorted(glob.glob(sys.argv[1] + '/[[]*.txt')):
    L = open(f, encoding='utf-8').read().split('\n')
    s = next(i for i, l in enumerate(L) if re.match(r'^1\. 임원 및 직원 등의 현황\s*$', l))
    e = next(i for i in range(s + 1, len(L)) if re.search(r'^\S*\.? ?직원 등(의)? 현황', L[i]))
    rows, on, seen = [], True, set()
    for i in range(s, e):
        l = L[i]
        if OFF.search(l) and not pat.match(l):
            on = False
        elif ON.search(l) and not pat.match(l):
            on = True
        m = pat.match(l)
        if m and on:
            k = (m.group(1), m.group(3), m.group(4))
            if k in seen:
                continue
            seen.add(k)
            rows.append(dict(name=m.group(1).strip(), sex=m.group(2), birth=f'{m.group(3)}.{int(m.group(4)):02d}',
                             reg=m.group(6), full=m.group(7), line_no=i + 1))
    co = os.path.basename(f).split(']')[0].strip('[')
    out[co] = rows
    reg = [r for r in rows if r['reg'] != '미등기']
    print(f"{co}: 등기 {len(reg)} / 미등기 {len(rows) - len(reg)} (상근 {sum(r['reg']=='미등기' and r['full']=='상근' for r in rows)})")
json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
