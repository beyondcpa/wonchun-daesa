# 원천세 대사표 자동생성 웹앱

급여대장·이행상황신고서·급여명세서·간이지급조서가 든 폴더(zip)를 올리면 원천세 대사표(엑셀)를 자동 생성하는 웹앱입니다. 여러 컴퓨터에서 웹 주소로 접속해 사용합니다.

## 구성 파일
- `app.py` — 웹 화면(업로드·미리보기·다운로드)
- `engine.py` — 원천파일 파싱(값 자동추출)
- `builder.py` — 대사표 엑셀 조립
- `requirements.txt` — 필요한 파이썬 패키지

## 업로드 자료 구조
연도 폴더를 통째로 zip으로 압축해서 올립니다. 폴더 안에 월별 하위폴더가 있어야 합니다.

```
2026.zip
└─ 2026/
   ├─ 01월/  (급여대장.xls, 원천세신고서.pdf, 급여명세서*.pdf, 사업간이지급조서.pdf …)
   ├─ 02월/
   └─ …
```

파일명 인식 규칙: `급여대장`(.xls), `신고`/`원천세`(.pdf), `급여명세`(.pdf, zip 안도 인식), `근로간이`(.pdf), `사업간이`(.pdf).

---

## 배포 방법 (GitHub → Streamlit Community Cloud, 무료)

### 1) GitHub 저장소 만들기
1. https://github.com 로그인 → 우측 상단 **+** → **New repository**
2. 이름 예: `wonchun-daesa` → **Create repository**

### 2) 파일 업로드
1. 만든 저장소에서 **Add file → Upload files**
2. `app.py`, `engine.py`, `builder.py`, `requirements.txt` 4개를 끌어다 놓기
3. 아래 **Commit changes** 클릭

### 3) Streamlit 클라우드에 배포
1. https://share.streamlit.io 접속 → **Continue with GitHub** 로 로그인(저장소 접근 허용)
2. **Create app → Deploy a public app from GitHub**
3. Repository: `본인아이디/wonchun-daesa`, Branch: `main`, Main file path: `app.py`
4. **Deploy** 클릭 → 1~2분 후 배포 완료

### 4) 주소 공유
- 배포되면 `https://<앱이름>.streamlit.app` 주소가 생깁니다.
- 이 주소를 가족·회사·직원에게 보내면 각자 컴퓨터 브라우저에서 바로 사용합니다.

### 업데이트
- 로직을 고치려면 GitHub 저장소의 파일을 수정(Upload files로 덮어쓰기)하면 Streamlit이 자동으로 다시 배포합니다.

---

## 로컬에서 먼저 실행해 보기 (선택)
파이썬이 깔린 PC에서:
```
pip install -r requirements.txt
streamlit run app.py
```
브라우저가 열리면 zip을 올려 테스트합니다.

## 참고 / 한계
- 현재 파일명·서식은 이노파인더스 자료 기준으로 인식합니다. 다른 회사도 파일명 규칙(급여대장/신고/급여명세/근로간이/사업간이)만 맞으면 동작하며, 서식이 크게 다르면 `engine.py`의 정규식을 조정하면 됩니다.
- 6월 급여대장이 이미지(PNG)이거나 형식이 표준과 다르면 해당 값은 대사표에서 주황(자료 미확보)으로 표시됩니다 → 그 셀만 수기 보완.
- 미제출비과세·상여·중도퇴사/정산 등은 해당 자료가 급여명세/별도자료로 들어오면 자동 반영되도록 확장할 수 있습니다.
