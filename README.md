# 신한카드 검색광고 게재보고 PPT 자동 생성

네이버 카드검색 (`card-search.naver.com`) 게재지면을 자동 캡처해
샘플과 동일한 레이아웃의 PPT 게재보고를 만들어 줍니다.

## 사용법

### 🌟 방법 1. 웹 UI (추천)
1. **`게재보고_웹UI_시작.bat`** 더블클릭 — 로컬 서버가 뜨고 브라우저가 자동으로 열립니다 (`http://127.0.0.1:5180`)
2. PC URL / MO URL 두 칸에 붙여넣기 (카드명은 자동 추출됨)
3. **[게재보고 만들기]** 클릭 → 진행률 표시되면서 PC 캡처 → MO 캡처 → PPT 조립 (총 ~30초)
4. 완료되면 **[PowerPoint 로 열기]** 또는 **[폴더 열기]** / **[다운로드]** 선택
5. 우측 패널에 최근 생성 이력이 누적됨 (재열기 가능)

### 방법 2. 콘솔 (간단)
- **`게재보고_생성.bat`** 더블클릭 → 안내에 따라 URL/카드명 입력

### 방법 3. CLI
```powershell
python C:\Users\FIN\Desktop\게재보고_자동화\generate_report.py `
    --pc "https://card-search.naver.com/item?cardAdId=..." `
    --mo "https://m-card-search.naver.com/item?cardAdId=..." `
    --card "신한카드 Deep Once" `
    --out "C:\Users\FIN\Desktop\게재보고_DeepOnce.pptx"
```

출력은 항상 바탕화면(`%USERPROFILE%\Desktop`)에 `(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_<카드명>.pptx` 형태로 저장됩니다 (동명 파일 있으면 자동으로 `_1`, `_2` 부여).

## 슬라이드 구성

| 슬라이드 | 내용 | 입력 |
|----------|------|------|
| PC 섹션 맨 앞 | 네이버 검색결과 OneBox (2장, 텍스트 안 잘리게 분할) | `--search-pc` / 웹 UI |
| PC 본문 | 카드 상세 페이지 (혜택 전부 펼친 풀페이지) | `--pc` (필수) |
| PC 섹션 맨 끝 | 신한카드 자체 안내 페이지 (PC 뷰포트) | `--landing` / 웹 UI |
| MO 섹션 맨 앞 | 네이버 검색결과 OneBox (2장 분할) | `--search-mo` / 웹 UI |
| MO 본문 | 카드 상세 페이지 (5장 그리드) | `--mo` (필수) |
| MO 섹션 맨 끝 | 신한카드 자체 안내 페이지 (모바일 뷰포트, 5장 그리드) | `--landing` / 웹 UI |

`--landing` 은 한 번만 입력하면 PC/MO 두 디바이스로 자동 캡처됩니다.

## 동작 방식

### PC 캡처
1. 1920×1080 뷰포트로 카드 상세 페이지 로딩
2. 모든 혜택 `<summary>` (렌탈/관리비/문화 등) 를 자동 클릭해 펼침
   — `.open=true` 만으론 lazy-load 가 안 되므로 **반드시 click()**
3. 페이지 내 `position:fixed/sticky` 요소를 `absolute` 로 변환
   (Playwright 풀페이지 캡처에서 sticky 헤더가 매 viewport 마다 중복 캡처되는 문제 해결)
4. 뷰포트를 페이지 전체 높이로 키워 한 번에 캡처 → stitching 없이 깔끔한 풀페이지 이미지
5. 989px 단위로 슬라이스 (샘플 PPT 와 동일 비율, 60px 겹침으로 누락 방지)
6. 별도 narrow viewport (801×1311) 로 타이틀 슬라이드용 vertical 캡처

### MO 캡처
- 모바일 UA (Galaxy S22) + 404×874 / 391×1258 / 검색결과 페이지 캡처
- 동일한 펼침 + sticky 회피 + 슬라이스 (874px, 50px 겹침)

### PPT 조립
- 16:9 widescreen, 샘플과 동일한 헤더 스타일 (좌: 회사명, 우: PC/MO 표시, 청색 구분선)
- PC: 타이틀 1장 + 풀폭 캡처 N장
- MO: 타이틀(리스트+상세) 1장 + 풀폭 캡처 5장씩 그리드 N장
- 누락 없이 페이지 처음부터 끝까지 모두 포함

## 파일 구조

| 파일 | 역할 |
|------|------|
| `generate_report.py` | 메인 캡처/PPT 조립 모듈 (CLI 도 됨) |
| `server.py` | Flask 로컬 웹서버 (job queue + progress polling) |
| `static/index.html` | 웹 UI 메인 페이지 |
| `static/styles.css` | UI 스타일 (Pretendard, 카드/패널 레이아웃) |
| `static/app.js` | UI 로직 (붙여넣기, 자동 추출, 진행률 표시) |
| `게재보고_웹UI_시작.bat` | 웹 UI launcher (서버 + 브라우저 자동 열기) |
| `게재보고_생성.bat` / `.ps1` | 콘솔 모드 launcher |
| `_history.json` | 최근 생성 이력 (자동 관리, 30 건) |

## 의존성

이미 설치되어 있어야 함 (현 시스템엔 모두 설치 확인됨):
- Python 3.x
- `playwright` (브라우저 포함: `python -m playwright install chromium`)
- `python-pptx`, `Pillow`, `flask`

## API (외부에서 호출하고 싶을 때)

서버 가동 중일 때 `http://127.0.0.1:5180/api/...`:

| 엔드포인트 | 설명 |
|-----------|------|
| `POST /api/generate` | `{pc_url, mo_url, card_name}` 받아 job 시작, `{job_id}` 반환 |
| `GET  /api/status/<id>` | `{status, step, progress, file_name, …}` |
| `GET  /api/download/<id>` | PPT 파일 다운로드 |
| `POST /api/open/<id>` | 로컬 PowerPoint 로 파일 열기 |
| `POST /api/reveal/<id>` | 탐색기에서 파일 위치 보기 |
| `GET  /api/history` | 최근 생성 이력 |
| `POST /api/extract-card` | `{url}` 의 query 파라미터에서 카드명 추출 |

## 트러블슈팅

- **혜택 본문이 비어 있음**: 네이버 페이지가 `summary` 클릭 시 lazy-load.
  `_open_all_details()` 가 click 처리. 새 혜택 카테고리가 추가됐다면 `details.details > summary` 셀렉터가 여전히 매칭되는지 확인.
- **PPT 가 너무 길거나 짧음**: 카드 페이지 자체 길이에 따라 슬라이드 수가 자동 결정. 강제로 줄이려면 `PC_SCREENSHOT_H`(989) 를 키우거나 `PC_OVERLAP`(60) 를 키움.
- **서버 포트 충돌**: 5180 이 다른 프로그램에 잡혀 있으면 `server.py` 의 `port = 5180` 을 다른 번호로 변경.
- **카드명에 한글 깨짐**: 브라우저 fetch() 는 UTF-8 기본이라 정상. PowerShell `ConvertTo-Json` 으로 직접 API 호출 시 한글이 깨지는 건 PS 5.1 의 이슈 — 브라우저 사용 권장.
