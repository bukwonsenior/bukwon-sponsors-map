# -*- coding: utf-8 -*-
"""
후원처 엑셀(후원처_지도관리.xlsx)  ->  sponsors.json

담당자는 상호명·업종·주소·전화만 입력한다. 좌표(위도·경도)는 비워둬도 된다.
이 스크립트가 카카오 지도로 주소(→ 안 되면 '원주 상호명')를 좌표로 바꿔 자동으로 채운다.
못 찾은 가게는 지도에서만 빠지고 나머지는 정상 갱신된다(_geocode_report.json 에 기록).
(원래 구글시트 Apps Script 'Untitled1.js' 가 하던 일을 파이썬으로 옮긴 것)

카카오 REST 키는 코드에 두지 않는다. GitHub Actions Secret(KAKAO_REST_KEY)에서 꺼내 쓴다.
한 번 찾은 좌표는 _coords_cache.json 에 저장해, 다음부터는 다시 검색하지 않는다.
"""
import json, os, re, sys, time, urllib.parse, urllib.request
from pathlib import Path

import openpyxl

XLSX  = Path("북원_후원업체_지도관리.xlsx")
OUT   = Path("sponsors.json")
CACHE = Path("_coords_cache.json")
REPORT = Path("_geocode_report.json")   # 편집기가 읽어 '좌표 못 찾은 가게'를 알려줌
KEY   = os.environ.get("KAKAO_REST_KEY", "").strip()

errors = []
def err(row, msg): errors.append(f"{row}행: {msg}")

def s(v):
    return "" if v is None else str(v).strip()

# 주민등록번호만 차단한다. (전화는 후원 가게 번호라 010 이어도 정상)
JUMIN = re.compile(r"\b\d{6}[-]\d{7}\b")

def load_cache():
    if CACHE.exists():
        try: return json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception: return {}
    return {}

# 원주시 대략 범위 (이 밖으로 찍히면 잘못 찾은 것으로 본다)
WONJU = (37.15, 37.60, 127.70, 128.25)
def in_wonju(lat, lng):
    try: lat, lng = float(lat), float(lng)
    except (TypeError, ValueError): return False
    return WONJU[0] <= lat <= WONJU[1] and WONJU[2] <= lng <= WONJU[3]

def _kakao(kind, query):
    url = (f"https://dapi.kakao.com/v2/local/search/{kind}.json?query=" + urllib.parse.quote(query))
    req = urllib.request.Request(url, headers={"Authorization": "KakaoAK " + KEY})
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read().decode("utf-8")).get("documents", [])

def kakao_geocode(name, addr):
    """(상호명, 주소) → (lat, lng, 방법). ① 주소검색 ② 주소 키워드 ③ '원주 상호명' 키워드.
    ①②는 주소가 가리키는 곳을 그대로 믿는다(서울·춘천 후원처도 있음).
    ③은 이름만으로 찾는 것이라 원주 범위 안 결과만 쓴다. 못 찾으면 None."""
    tries = [("address", addr, False), ("keyword", addr, False)]
    if name: tries.append(("keyword", f"원주 {name}", True))
    for kind, q, wonju_only in tries:
        try:
            docs = _kakao(kind, q)
        except Exception as e:
            raise RuntimeError(f"카카오 API 호출 실패: {e}")
        for d in docs[:5]:
            lat, lng = float(d["y"]), float(d["x"])
            if not wonju_only or in_wonju(lat, lng):
                return lat, lng, f"{kind}:{q}"
        time.sleep(0.1)
    return None

def main():
    if not XLSX.exists():
        print(f"::error::{XLSX} 파일이 없습니다.")
        sys.exit(1)

    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "후원처" not in wb.sheetnames:
        print("::error::'후원처' 시트가 없습니다. 시트 이름을 바꾸지 마세요.")
        sys.exit(1)
    ws = wb["후원처"]

    head = [s(c.value) for c in ws[4]]
    need = ["게시","상호명","업종","주소","전화","위도","경도","비고"]
    optional = ["후원내용","소개"]          # 없어도 됨 (예전 엑셀 호환)
    for h in need:
        if h not in head:
            errors.append(f"'후원처' 시트 4행에 '{h}' 열이 없습니다. 열 이름을 바꾸지 마세요.")
    if errors:
        for e in errors: print("  ✗ " + e)
        print(f"::error::엑셀 열 이름 문제 {len(errors)}건")
        sys.exit(1)
    ix = {h: head.index(h) for h in need + optional if h in head}
    def col(row, h):
        i = ix.get(h)
        return s(row[i]) if (i is not None and i < len(row)) else ""

    cache = load_cache()
    staged, to_geocode, warns = [], [], []
    report = {"좌표못찾음": [], "새로찾음": []}

    for r in range(5, ws.max_row + 1):
        row = [c.value for c in ws[r]]
        if not any(s(v) for v in row): continue
        pub = col(row, "게시").upper()
        if pub in ("X", "×"): continue
        if pub != "O":
            err(r, f"게시 칸은 O 또는 X 만 (지금: '{col(row, '게시')}')"); continue
        name, addr = col(row, "상호명"), col(row, "주소")
        if not name: err(r, "상호명이 비어 있습니다"); continue
        if not addr: err(r, f"'{name}' 의 주소가 비어 있습니다"); continue
        if JUMIN.search(" ".join(s(v) for v in row)):
            err(r, "주민등록번호로 보이는 값이 있습니다. 공개 저장소라 올릴 수 없습니다")
        rec = {"name": name, "cat": col(row, "업종"), "addr": addr, "tel": col(row, "전화"),
               "memo": col(row, "비고"), "support": col(row, "후원내용"), "intro": col(row, "소개"),
               "lat": col(row, "위도"), "lng": col(row, "경도"), "row": r}
        if rec["lat"] and rec["lng"]:
            pass                       # 엑셀에 좌표를 직접 넣은 경우 그대로
        elif addr in cache:
            rec["lat"], rec["lng"] = cache[addr][0], cache[addr][1]
            rec["fill"] = True
        else:
            to_geocode.append(rec)
        staged.append(rec)

    if errors:
        print("\n엑셀에서 고쳐야 할 곳이 있습니다. 지도는 그대로 둡니다.\n")
        for e in errors: print("  ✗ " + e)
        print(f"::error::엑셀 오류 {len(errors)}건")
        sys.exit(1)

    # 새 주소만 좌표 검색 — 못 찾은 곳은 지도에서만 빼고 나머지는 정상 갱신
    if to_geocode and not KEY:
        print("::warning::KAKAO_REST_KEY 가 없어 새 주소 좌표를 찾지 못했습니다 (GitHub Secrets 확인).")
    for rec in to_geocode:
        coord = None
        if KEY:
            try:
                coord = kakao_geocode(rec["name"], rec["addr"])
            except RuntimeError as e:
                print("::warning::" + str(e))
            time.sleep(0.12)
        if coord:
            cache[rec["addr"]] = [coord[0], coord[1]]
            rec["lat"], rec["lng"] = coord[0], coord[1]
            rec["fill"] = True
            report["새로찾음"].append({"행": rec["row"], "상호명": rec["name"], "주소": rec["addr"], "방법": coord[2]})
        else:
            warns.append(f"{rec['row']}행 {rec['name']}: 주소로 위치를 못 찾아 지도에서 뺐습니다 → 주소를 정확히 고치거나 위도·경도를 직접 넣어 주세요 ({rec['addr']})")
            report["좌표못찾음"].append({"행": rec["row"], "상호명": rec["name"], "주소": rec["addr"]})

    # 찾은 좌표를 엑셀 위도·경도 칸에도 채워 넣음 → 다음에 편집기·엑셀을 열면 좌표가 보임
    filled = [x for x in staged if x.get("fill")]
    if filled:
        for x in filled:
            ws.cell(x["row"], ix["위도"] + 1, float(x["lat"]))
            ws.cell(x["row"], ix["경도"] + 1, float(x["lng"]))
        wb.save(XLSX)
        print(f"  · 엑셀에 좌표 {len(filled)}곳 채움")

    shown = [x for x in staged if x["lat"] and x["lng"]]
    header = ["연번","상호명","업종","주소","전화","위도","경도","비고","후원내용","소개"]
    arr = [header]
    for i, rec in enumerate(shown, start=1):
        arr.append([i, rec["name"], rec["cat"], rec["addr"], rec["tel"],
                    rec["lat"], rec["lng"], rec["memo"], rec["support"], rec["intro"]])
    OUT.write_text(json.dumps(arr, ensure_ascii=False, indent=1), encoding="utf-8")
    CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    for w in warns: print("::warning::" + w)
    print(f"✓ {OUT} 생성 — 지도 {len(shown)}곳 (새로 좌표 찾음 {len(report['새로찾음'])}곳, 못 찾음 {len(report['좌표못찾음'])}곳)")

if __name__ == "__main__":
    main()
