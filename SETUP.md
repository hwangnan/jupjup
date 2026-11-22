# 줍줍 설치 가이드

구조: **화면(web/) → GitHub Pages**, **데이터 API(gas/) → Google Apps Script**, **DB → Google Sheets**

```
web/      앱 화면 (PWA). GitHub Pages에 올라가요
gas/      Apps Script 코드. 스프레드시트에 붙여넣어요
.github/  web/ 폴더를 GitHub Pages에 자동 배포
```

연결 전에는 앱이 **예시 데이터**로 동작해요. 아래 1~2단계를 마치면 실제 기록으로 바뀌어요.

---

## 1. Google Sheets + Apps Script (데이터 API)

1. [sheets.new](https://sheets.new)에서 새 스프레드시트를 만들고 이름을 `줍줍 DB`로 바꿔요.
2. 메뉴 **확장 프로그램 → Apps Script**를 열어요.
3. 기본 `Code.gs` 내용을 지우고 `gas/Code.gs` 내용을 통째로 붙여넣어요.
4. 왼쪽 **⚙ 프로젝트 설정 → "편집기에서 appsscript.json 매니페스트 파일 표시"**를 켜고, 편집기에 생긴 `appsscript.json`에 `gas/appsscript.json` 내용을 붙여넣어요.
5. 상단 함수 목록에서 `setup`을 고르고 **실행**해요.
   - 처음엔 권한 승인 창이 떠요. 본인 계정으로 허용해요. ("확인되지 않은 앱" 경고가 나오면 **고급 → 이동**)
   - 실행 로그에 `API 토큰: xxxx…`이 나오면 복사해 둬요. 시트에 `Items / Comps / Settings / Log` 탭이 생겨요.
6. 오른쪽 위 **배포 → 새 배포 → 유형: 웹 앱**
   - 실행 사용자: **나**
   - 액세스 권한: **모든 사용자**
   - 배포 후 나오는 `https://script.google.com/macros/s/…/exec` 주소를 복사해요.

> 토큰은 비밀번호처럼 다뤄요. 잊어버리면 `showToken`, 바꾸려면 `rotateToken` 함수를 실행해요.
> 코드를 고친 뒤에는 **배포 → 배포 관리 → ✏ 수정 → 버전: 새 버전**으로 올려야 반영돼요. (주소는 그대로 유지)

## 2. 앱과 연결

앱 **설정 탭 → 연결 → Google Sheets**에서 웹 앱 주소와 API 토큰을 넣고 **연결하기**를 눌러요.
연결 정보는 그 기기(브라우저)에만 저장돼요. 폰과 PC에서 각각 한 번씩 연결해요.

## 3. GitHub Pages (앱 화면 배포)

1. GitHub에 새 저장소를 만들어요. (예: `jupjup`, Public)
2. 이 폴더를 올려요.
   ```powershell
   git init -b main
   git add .
   git commit -m "줍줍 첫 버전"
   git remote add origin https://github.com/<내아이디>/jupjup.git
   git push -u origin main
   ```
3. 저장소 **Settings → Pages → Source: GitHub Actions**로 바꿔요.
4. Actions 탭에서 배포가 끝나면 `https://<내아이디>.github.io/jupjup/` 주소로 열려요.
5. 폰에서 열고 **홈 화면에 추가**하면 앱처럼 설치돼요.

> 저장소가 공개여도 괜찮아요. 코드에는 토큰이나 시트 주소가 들어 있지 않아요.

## 로컬에서 미리보기

```powershell
python -m http.server 5173 --directory web
```
브라우저에서 `http://localhost:5173` 을 열어요.

---

## 데이터 구조 (Google Sheets)

| 시트 | 내용 |
|---|---|
| Items | 매물 1건 = 1행. `status`: new(새 후보) · watch(관심) · pass · bought(매입) · selling(판매중) · sold(완료) |
| Comps | 판매완료 시세. 매물(`itemId`)마다 여러 건 |
| Settings | 매입 기준, 건당 비용, 알림 키워드 |
| Log | 상태 변경 기록 (개인화 분석용) |

- `askPrice` = 처음 올라온 가격, `buyPrice` = 실제 매입가 → 네고 효과를 나중에 분석할 수 있어요.
- 시세(Comps)는 **실제로 팔린 가격만** 넣어요. 정가(`retail`)는 참고용이고 계산에 쓰지 않아요.
- 매물 사진은 내 Google Drive의 `줍줍 매물 사진` 폴더에 "링크가 있는 사용자 보기" 권한으로 저장돼요.

## 4. MacroDroid 알림 수집 (안드로이드)

당근 키워드 알림 → MacroDroid → 웹 앱 `ingest` → 가격·동네 읽기 → 키워드 최대 가격 필터 → 48시간 중복 제거 → `Items`에 `new`로 저장.
모든 알림 원문은 `Inbox` 시트에 결과(`created` / `duplicate` / `over-max` / `noprice`)와 함께 남아요.

> Apps Script 코드를 바꿨다면 먼저 **배포 → 배포 관리 → ✏ → 새 버전**으로 다시 배포하세요.

1. 당근 앱 **나의 당근 → 키워드 알림**에 키워드를 등록하고, 줍줍 **설정 → 알림 키워드**에도 같은 이름과 최대 가격을 넣어요.
2. MacroDroid에서 **매크로 추가**
   - **트리거**: 알림 → 알림 수신 → 앱 선택: **당근**
   - **액션**: 연결 → **HTTP 요청**
     - 방식 **POST**, URL: 줍줍 **설정 → MacroDroid 알림 수집**에서 복사한 주소
     - **쿼리 매개변수**: `action`=`ingest`, `token`=(토큰), `title`=알림 제목, `text`=알림 텍스트, `app`=앱 이름
       (알림 제목·텍스트·앱 이름은 **…** 매직 텍스트 버튼에서 골라 넣어요)
   - **액션 2**: 알림 → **알림 지우기** → 앱 **당근** (키워드 알림이 알림창에 쌓이지 않게)
     채팅 알림까지 지워지지 않게, 가능하면 "제목 포함" 조건에 ` - ` 를 넣어요. 키워드 알림 제목은 `키워드 - 구 동` 형식이에요.
3. 휴대폰 설정에서 MacroDroid의 **알림 접근 권한**을 허용하고, **배터리 최적화 제외**로 둬요. (안 하면 백그라운드에서 멈춰요)
4. 알림이 오면 줍줍 발굴 탭에 `키워드 알림`으로 떠요. 안 뜨면 `Inbox` 시트의 `result`를 확인해요.

알림 문구를 어떻게 읽는지 미리 보려면 Apps Script 편집기에서 `testParse`를 실행해요 (시트에 쓰지 않아요).

## 다음 단계

- AI 분석: 사진·제목으로 상품 식별, 상태 체크리스트 (Claude API 키 필요, 사용량 과금)
