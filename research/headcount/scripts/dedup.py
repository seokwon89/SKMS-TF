"""멤버사 간 겸직 미등기임원이 여러 회사 직원수에 중복 포함될 수 있는 최대 인원을 계산한다."""
import json, sys, collections
ex = json.load(open(sys.argv[1]))
# 직원수에 미등기임원 포함 여부 (사업보고서 주석 기준)
INCL = {'SK에너지': 'Y', '에스케이지오센트릭': 'Y', '에스케이온': 'Y', 'SK이노베이션': 'Y', 'SK가스': 'Y',
        'SK하이닉스': 'Y(회장·부회장 제외)', 'SK텔레콤': '?', '에스케이에코플랜트': '?', 'SK디스커버리': '?'}
EXCL_PERSON = {('SK하이닉스', '최태원')}  # 하이닉스 직원수는 회장 제외 명시
m = collections.defaultdict(list)
for co, rows in ex.items():
    for r in rows:
        if r['reg'] == '미등기' and (co, r['name']) not in EXCL_PERSON:
            m[(r['name'].rstrip('ab'), r['birth'])].append(co)
strict = maxd = 0
lines = []
for k, cos in sorted(m.items()):
    cos = sorted(set(cos))
    if len(cos) < 2:
        continue
    y = [c for c in cos if INCL[c].startswith('Y')]
    a = [c for c in cos if INCL[c] != 'N']
    strict += max(len(y) - 1, 0)
    maxd += max(len(a) - 1, 0)
    lines.append(f"{k[0]}({k[1]}): {', '.join(cos)}  → 확정포함사 중복 {max(len(y)-1,0)}, 최대 {max(len(a)-1,0)}")
print('\n'.join(lines))
print(f'중복 가능 인원(확정 포함사 기준) {strict} / 최대(미기재사 포함) {maxd}')
