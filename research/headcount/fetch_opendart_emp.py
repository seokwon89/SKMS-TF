"""OpenDART에서 SUPEX 멤버사 21개사의 사업보고서 '직원 현황'과 '미등기임원 수'를 받아 CSV로 저장한다.

사용법:
    export DART_API_KEY=발급받은키   # https://opendart.fss.or.kr 에서 무료 발급
    python3 fetch_opendart_emp.py --year 2025 --out 02_직원현황_2025사업보고서_raw.csv

- reprt_code 11011 = 사업보고서, 11012 = 반기보고서
- 비상장사는 corpCode 목록에서 회사명으로 찾는다. 후보가 여러 개면 모두 출력하니 확인 후 CORP_CODE_OVERRIDE에 고정할 것.
- 모회사 안의 사내 조직(SK이노베이션 E&S, SK엔무브, SK㈜ AX 등)은 사업부문(fo_bbm)별 행을 *_by_segment.csv로 따로 저장한다.
- 실제 API 응답으로 아직 실행해 보지 못한 스크립트다. 첫 실행 결과를 DART 원문 한두 건과 대조할 것.
"""
import argparse
import csv
import io
import json
import os
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

API = "https://opendart.fss.or.kr/api"

# 첨부 「'26년 멤버사 및 협의회 위원 건제」(2026.1.) 건제 순.
# (멤버사, 상장 종목코드 또는 None, 비상장일 때 corpCode 검색용 이름 후보, 집계 법인)
MEMBERS = [
    ("SK㈜", "034730", [], "SK㈜"),
    ("SK하이닉스", "000660", [], "SK하이닉스"),
    ("SK이노베이션", "096770", [], "SK이노베이션"),
    ("SK텔레콤", "017670", [], "SK텔레콤"),
    ("SK디스커버리", "006120", [], "SK디스커버리"),
    ("SK온", None, ["SK온", "에스케이온"], "SK온"),
    ("SK에너지", None, ["SK에너지", "에스케이에너지"], "SK에너지"),
    ("SK에코플랜트", None, ["SK에코플랜트", "에스케이에코플랜트"], "SK에코플랜트"),
    ("SK이노베이션 E&S", None, [], "SK이노베이션"),  # 2024.11 합병, CIC
    ("SK지오센트릭", None, ["SK지오센트릭", "에스케이지오센트릭"], "SK지오센트릭"),
    ("SK브로드밴드", None, ["SK브로드밴드", "에스케이브로드밴드"], "SK브로드밴드"),
    ("SK가스", "018670", [], "SK가스"),
    ("SK실트론", None, ["SK실트론", "에스케이실트론"], "SK실트론"),
    ("SK스퀘어", "402340", [], "SK스퀘어"),
    ("SK네트웍스", "001740", [], "SK네트웍스"),
    ("SK엔무브", None, ["SK엔무브", "에스케이엔무브"], "SK온"),  # 2025.11 SK온 흡수합병(완료 여부 확인 필요)
    ("SKC", "011790", [], "SKC"),
    ("SK케미칼", "285130", [], "SK케미칼"),
    ("SK㈜ AX", None, [], "SK㈜"),  # 구 SK㈜ C&C, 사내독립기업으로 추정
    ("SK에코플랜트 머티리얼즈", None, ["SK에코플랜트머티리얼즈", "SK에코플랜트 머티리얼즈", "에스케이에코플랜트머티리얼즈"], None),
    ("SK바이오팜", "326030", [], "SK바이오팜"),
]

CORP_CODE_OVERRIDE = {}  # 예: {"SK온": "01234567"}


def get(path, **params):
    url = f"{API}/{path}?{urllib.parse.urlencode(params)}"
    with urllib.request.urlopen(url, timeout=60) as r:
        return r.read()


def load_corp_codes(key):
    raw = get("corpCode.xml", crtfc_key=key)
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        root = ET.fromstring(z.read(z.namelist()[0]))
    rows = []
    for el in root.iter("list"):
        rows.append({
            "corp_code": el.findtext("corp_code"),
            "corp_name": (el.findtext("corp_name") or "").strip(),
            "stock_code": (el.findtext("stock_code") or "").strip(),
        })
    return rows


def resolve(corps, stock_code, names, label):
    if label in CORP_CODE_OVERRIDE:
        return CORP_CODE_OVERRIDE[label], ""
    if stock_code:
        hits = [c for c in corps if c["stock_code"] == stock_code]
    else:
        wanted = {n.replace(" ", "") for n in names}
        hits = [c for c in corps if c["corp_name"].replace(" ", "") in wanted]
    if len(hits) == 1:
        return hits[0]["corp_code"], ""
    note = "후보 없음" if not hits else "후보 여러 개: " + ", ".join(f'{h["corp_name"]}({h["corp_code"]})' for h in hits)
    return None, note


def num(s):
    s = (s or "").replace(",", "").strip()
    return int(s) if s.lstrip("-").isdigit() else 0


def fetch_list(key, path, corp_code, year, reprt_code):
    data = json.loads(get(path, crtfc_key=key, corp_code=corp_code, bsns_year=year, reprt_code=reprt_code))
    if data.get("status") != "000":
        return None, f'{data.get("status")} {data.get("message")}'
    return data.get("list", []), ""


def is_total(row):
    return any(t in (row.get("fo_bbm") or "") + (row.get("sexdstn") or "") for t in ("합계", "총계", "전체"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="2025")
    ap.add_argument("--reprt-code", default="11011")
    ap.add_argument("--out", default="02_직원현황_2025사업보고서_raw.csv")
    a = ap.parse_args()
    key = os.environ.get("DART_API_KEY")
    if not key:
        sys.exit("DART_API_KEY 환경변수가 필요합니다.")

    corps = load_corp_codes(key)
    cache, out, seg_rows = {}, [], []
    for label, stock, names, entity in MEMBERS:
        row = {"멤버사": label, "집계법인": entity or "", "corp_code": "", "정규직": "", "기간제": "",
               "직원합계": "", "미등기임원": "", "비고": ""}
        if entity and entity != label:
            row["비고"] = f"{entity} 내부 조직 — {entity} 사업부문별 행에서 인원 확인 후 기입"
            out.append(row)
            continue
        code, note = resolve(corps, stock, names, label)
        if not code:
            row["비고"] = note + " — 사업보고서 미제출이면 기업집단포털 종업원수로 보완"
            out.append(row)
            continue
        row["corp_code"] = code
        emp, err = fetch_list(key, "empSttus.json", code, a.year, a.reprt_code)
        if emp is None:
            row["비고"] = f"직원현황 없음({err}) — 기업집단포털로 보완"
            out.append(row)
            continue
        totals = [r for r in emp if is_total(r)]
        base = totals if totals else emp
        row["정규직"] = sum(num(r.get("rgllbr_co")) for r in base)
        row["기간제"] = sum(num(r.get("cnttk_co")) for r in base)
        row["직원합계"] = sum(num(r.get("sm")) for r in base)
        if totals:
            row["비고"] = "합계 행 사용 — 원문 대조 필요"
        for r in emp:
            seg_rows.append({"멤버사": label, "사업부문": r.get("fo_bbm"), "성별": r.get("sexdstn"),
                             "정규직": num(r.get("rgllbr_co")), "기간제": num(r.get("cnttk_co")),
                             "합계": num(r.get("sm")), "rcept_no": r.get("rcept_no")})
        ex, _ = fetch_list(key, "unrstExctvMendngSttus.json", code, a.year, a.reprt_code)
        if ex:
            row["미등기임원"] = sum(num(r.get("nmpr")) for r in ex)
        out.append(row)

    with open(a.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
        w.writeheader()
        w.writerows(out)
    seg_path = a.out.replace(".csv", "_by_segment.csv")
    if seg_rows:
        with open(seg_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(seg_rows[0].keys()))
            w.writeheader()
            w.writerows(seg_rows)
    print(f"저장: {a.out}, {seg_path}")


if __name__ == "__main__":
    main()
