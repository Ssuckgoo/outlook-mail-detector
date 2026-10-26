# 메일 텍스트 감지기 (Outlook 추가기능)

메일을 **보내는 순간** 본문에서 `#요청:값` 형태의 텍스트를 찾아 외부 API로 전송합니다.
클래식 Outlook, 새 Outlook, 웹 Outlook에서 모두 동작합니다.

## 동작 흐름

```
[보내기] 클릭
  → Outlook이 등록된 manifest를 보고 launchevent.js를 받아와 실행
  → 본문에서 #요청:값 검색
      ├─ 없음 → 바로 발송
      └─ 있음 → API로 POST
            ├─ 성공 → 발송
            └─ 실패 → "그래도 보내기 / 보내지 않기" 선택창
```

## 폴더 구조

| 파일 | 역할 |
|---|---|
| `manifest.xml` | **실사용용** 설정. 코드를 GitHub Pages(https://ssuckgoo.github.io/outlook-mail-detector)에서 받아옴 |
| `manifest.dev.xml` | **개발용** 설정. 코드를 내 PC의 `server.js`(https://localhost:3000)에서 받아옴 |
| `docs/launchevent.js` | 실제 로직 (본문 검사 + API 호출). **패턴과 API 주소는 맨 위에서 수정** |
| `docs/commands.html` | 새 Outlook / 웹 Outlook이 launchevent.js를 불러올 때 쓰는 빈 페이지 |
| `docs/taskpane.html` | 추가기능 소개 페이지 |
| `docs/index.html` | GitHub Pages 기본 주소(`/`)로 들어오면 소개 페이지로 이동 |
| `docs/assets/` | 아이콘 |
| `server.js` | 개발용 HTTPS 서버 + 가짜 API (`/api/hook`, 받은 내용을 콘솔에 출력) |

`docs/` 폴더가 GitHub Pages로 공개되는 폴더입니다. 여기를 수정해서 푸시하면 실사용 추가기능에 반영됩니다.

## 실사용 (GitHub Pages)

### 1. GitHub Pages 켜기 (최초 1회)
저장소 → **Settings → Pages** → Source: **Deploy from a branch** → Branch: **main**, 폴더: **/docs** → Save

1~2분 뒤 https://ssuckgoo.github.io/outlook-mail-detector/ 에서 소개 페이지가 보이면 성공입니다.

### 2. Outlook에 등록 (최초 1회)
1. 브라우저에서 https://aka.ms/olksideload 접속
2. **내 추가 기능** → **사용자 지정 추가 기능** → **파일에서 추가**
3. `manifest.xml` 선택 → 설치

메일 계정에 등록되므로 클래식, 새 Outlook에도 자동 적용됩니다.
(클래식 Outlook은 재시작해야 반영될 수 있습니다)

### 3. 코드 수정 반영
`docs/` 안의 파일을 고쳐서 푸시하면 끝입니다. 다시 등록할 필요는 없습니다.
(Outlook이 이전 파일을 캐시하고 있으면 재시작해야 반영될 수 있습니다)

`manifest.xml` 자체를 고쳤을 때만 2번을 다시 합니다.

## 개발 (내 PC에서 테스트)

```
npm install   # 최초 1회
npm start     # https://localhost:3000 서버 실행
```
- 처음 실행하면 **개발용 인증서 설치 확인창**이 뜹니다. "예"를 누릅니다.
- Outlook에는 `manifest.dev.xml`을 등록합니다 (실사용용과 이름이 달라서 동시에 설치 가능).
- 메일 본문에 `#요청:A001`을 넣고 보내면 서버 콘솔에 감지된 내용이 출력됩니다.

## 실제 API로 바꾸기

`docs/launchevent.js` 맨 위 두 줄을 수정합니다.
```js
var PATTERN = /#요청:(\S+)/g;                       // 찾을 텍스트 규칙
var TARGET_API = "https://localhost:3000/api/hook"; // 보낼 API 주소
```
- 기본값은 개발용 가짜 API라서, `npm start`가 꺼져 있으면 패턴이 있는 메일을 보낼 때 실패 선택창이 뜹니다.
- 실제 API는 **https**여야 하고, **CORS를 허용**해야 합니다 (`Access-Control-Allow-Origin`).

## 제약 사항

- **Microsoft 메일 계정**(Microsoft 365, outlook.com, hotmail)에서만 동작합니다. Outlook에 Gmail이나 네이버를 IMAP으로 연결한 계정은 추가기능을 지원하지 않습니다.
- 클래식 Outlook은 **Microsoft 365 구독 버전**이어야 합니다 (Outlook 2016/2019 단품 버전은 보내기 이벤트 미지원).
- 개인 계정(outlook.com)에서 보내기 이벤트가 동작하는지는 실제 테스트로 확인이 필요합니다.

## 설정 파일 검사
```
npm run validate
```
