# 북원노인종합복지관 후원처 지도

후원 가게 위치를 카카오 지도에 표시합니다.

## 담당자가 하는 일

`data/북원_후원업체_지도관리.xlsx` 를 고쳐서 GitHub `data/` 폴더에 업로드하면 끝입니다.

- 상호명·업종·주소·전화만 입력하세요. **위도·경도는 비워두면 자동으로 채워집니다.**
- 후원처를 내리려면 A열 게시를 `X` 로 (줄은 지우지 마세요).
- 자세한 건 엑셀의 '사용법' 시트를 보세요.

## 자동으로 되는 일

엑셀을 올리면 GitHub Actions 가:
1. 주소를 카카오 지도로 검색해 좌표(위도·경도)를 채우고
2. `data/sponsors.json` 을 만들고
3. 지도(`index.html`)가 그 파일을 읽어 표시합니다.

한 번 찾은 좌표는 `data/_coords_cache.json` 에 저장돼, 새 주소만 다시 검색합니다.

## 설정 (최초 1회)

- 저장소 **Settings → Secrets and variables → Actions → New repository secret**
  - Name: `KAKAO_REST_KEY` / Value: 카카오 REST 키
- 카카오 개발자콘솔에서 **JS 앱키의 허용 도메인**을 이 저장소의 Pages 주소로 제한
  - 예: `https://bukwonsenior.github.io`

## 파일

| 경로 | 역할 | 누가 |
|---|---|---|
| `data/북원_후원업체_지도관리.xlsx` | 원본 데이터 | 담당자 |
| `data/sponsors.json` | 자동 생성 | 자동 |
| `data/_coords_cache.json` | 좌표 캐시 (자동) | 자동 |
| `scripts/build_sponsors.py` | 변환·지오코딩 | 개발 시 |
| `.github/workflows/build-sponsors.yml` | 자동 실행 | 개발 시 |
| `index.html` | 지도 화면 | 개발 시 |

**개인정보 금지** — 후원'자'(개인) 정보는 넣지 마세요. 공개 저장소이며, 후원 '가게' 정보만 올립니다.
