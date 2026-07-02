# 구글드라이브 자동읽기 설정 가이드

앱에서 **☁️ 구글드라이브에서 읽기**를 쓰려면, 앱(서비스계정)이 드라이브 폴더를 읽을 수 있게 1회 설정이 필요합니다. (약 10분)

zip 업로드 방식만 쓸 거면 이 설정은 필요 없습니다.

---

## 1) 구글 클라우드에서 서비스계정 만들기

1. https://console.cloud.google.com 접속 (구글 로그인)
2. 상단 프로젝트 선택 → **새 프로젝트** 생성 (이름 예: `wonchun-daesa`)
3. 좌측 메뉴 **API 및 서비스 → 라이브러리** → "Google Drive API" 검색 → **사용 설정(Enable)**
4. **API 및 서비스 → 사용자 인증 정보(Credentials)** → 상단 **+ 사용자 인증 정보 만들기 → 서비스 계정**
   - 이름 아무거나(예: `daesa-reader`) → 만들기 → 역할은 건너뛰어도 됨 → 완료
5. 만들어진 서비스계정 클릭 → **키(KEYS)** 탭 → **키 추가 → 새 키 만들기 → JSON** → 다운로드
   - JSON 파일이 받아집니다(중요). 그 안의 `client_email` 값(…@….iam.gserviceaccount.com)을 복사해 두세요.

## 2) 드라이브 폴더를 서비스계정에 공유

1. 구글드라이브에서 원천자료가 든 **연도 폴더(예: 2026)** 를 우클릭 → **공유**
2. 위에서 복사한 서비스계정 이메일(`...iam.gserviceaccount.com`)을 **뷰어(보기 권한)** 로 추가
   - 공유드라이브라면, 그 공유드라이브 멤버로 서비스계정을 추가

## 3) Streamlit 앱에 인증정보 등록 (Secrets)

1. https://share.streamlit.io 에서 배포한 앱 → 우측 **⋮ → Settings → Secrets**
2. 아래 형식으로 붙여넣기 (다운로드한 JSON 내용을 그대로 옮김):

```toml
[gcp_service_account]
type = "service_account"
project_id = "여기에-project_id"
private_key_id = "여기에-private_key_id"
private_key = "-----BEGIN PRIVATE KEY-----\n...여러줄...\n-----END PRIVATE KEY-----\n"
client_email = "daesa-reader@....iam.gserviceaccount.com"
client_id = "여기에-client_id"
token_uri = "https://oauth2.googleapis.com/token"
```

> JSON 파일의 각 값을 그대로 옮기면 됩니다. `private_key`는 줄바꿈이 `\n`으로 들어간 긴 문자열 그대로 큰따옴표 안에 넣으세요.

3. **Save** → 앱이 자동 재시작됩니다.

## 4) 사용

1. 앱 접속 → 비번 `beyond`
2. **☁️ 구글드라이브에서 읽기** 선택
3. 연도 폴더의 **드라이브 링크**(주소창의 `.../folders/XXXX`) 붙여넣기 → **드라이브에서 불러오기**
4. 자동으로 파일 받아 대사표 생성 → 다운로드

---

## 참고
- 서비스계정에 **공유 안 된 폴더는 못 읽습니다.** 새 회사·새 연도 폴더가 생기면 그 폴더(또는 상위 공유드라이브)에 서비스계정을 공유해 두세요.
- 파일명·폴더구조는 자유입니다(월·종류 자동 인식). 다만 자료 종류를 알 수 있는 단어(급여대장/신고서/급여명세/근로간이/사업간이)는 파일명에 있어야 합니다.
- 속도: 드라이브에서 파일 받는 시간 + 파싱(3~5초). zip 업로드와 비슷하며, zip 만드는 수고가 없어집니다.
