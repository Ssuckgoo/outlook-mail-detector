// 메일 보내기 버튼을 누르면 Outlook 이 실행하는 코드
// 클래식 Outlook 은 이 파일을 브라우저가 아닌 JS 전용 엔진으로 직접 실행하므로
// 최신 문법(async/await 등) 대신 Promise 와 일반 함수로 작성한다.

// 본문에서 찾을 텍스트 규칙: "#요청:A001" -> "A001"
var PATTERN = /#요청:(\S+)/g;

// 찾은 내용을 보낼 API 주소
//  - 외부 테스트용: webhook.site (받은 요청을 웹에서 확인, 2026-11-07 만료)
//  - 로컬 테스트용: "https://localhost:3000/api/hook" (server.js 의 가짜 API, manifest.dev.xml 과 함께 사용)
var TARGET_API = "https://webhook.site/e4022ef2-a425-4e7b-88ba-fb3ce5d4f0d0";

// API 가 이 시간 안에 응답하지 않으면 실패로 처리 (Outlook 이 무한히 기다리지 않도록)
var TIMEOUT_MS = 10000;

function withTimeout(promise, ms) {
  return new Promise(function (resolve, reject) {
    var timer = setTimeout(function () {
      reject(new Error(ms / 1000 + "초 동안 API 응답이 없습니다 (" + TARGET_API + ")"));
    }, ms);
    promise.then(
      function (value) { clearTimeout(timer); resolve(value); },
      function (err) { clearTimeout(timer); reject(err); }
    );
  });
}

// 디버그: 실행 단계를 서버 터미널에 남긴다 (빈 문자열이면 끔)
// 이미지 요청은 CORS 영향을 받지 않아 fetch 가 막혀도 서버까지 도달한다. 클래식 Outlook(Image 없음)에서는 콘솔에만 남긴다.
// GitHub Pages 에서 실행될 때는 브라우저가 localhost 요청을 막으므로 꺼둔다. 로컬 디버깅 시 "https://localhost:3000/trace"
var TRACE_URL = "";

function trace(step) {
  console.log("[감지기] " + step);
  try {
    if (TRACE_URL && typeof Image !== "undefined") {
      new Image().src = TRACE_URL + "?step=" + encodeURIComponent(step) + "&t=" + Date.now();
    }
  } catch (e) {}
}

trace("0. 스크립트 로드됨");

// Office 의 콜백 방식 getAsync 를 Promise 로 바꿔주는 도우미
function getAsync(target, options) {
  return new Promise(function (resolve, reject) {
    var callback = function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    };
    if (options === undefined) {
      target.getAsync(callback);
    } else {
      target.getAsync(options, callback);
    }
  });
}

function findMatches(text) {
  var matches = [];
  var m;
  PATTERN.lastIndex = 0;
  while ((m = PATTERN.exec(text)) !== null) {
    matches.push(m[1]);
  }
  return matches;
}

function onMessageSendHandler(event) {
  trace("1. 보내기 이벤트 시작");
  var item = Office.context.mailbox.item;

  getAsync(item.body, Office.CoercionType.Text)
    .then(function (body) {
      var matches = findMatches(body);
      trace("2. 본문 읽음, 찾은 값: " + matches.join(","));
      if (matches.length === 0) {
        // 찾는 텍스트가 없으면 평소처럼 바로 보낸다
        event.completed({ allowEvent: true });
        return;
      }

      return Promise.all([getAsync(item.subject), getAsync(item.to)]).then(function (values) {
        var payload = {
          subject: values[0],
          to: values[1].map(function (r) { return r.emailAddress; }),
          matches: matches,
          sentAt: new Date().toISOString()
        };

        trace("3. API 호출: " + TARGET_API);
        return withTimeout(fetch(TARGET_API, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        }), TIMEOUT_MS).then(function (res) {
          trace("4. API 응답: " + res.status);
          if (!res.ok) {
            throw new Error("HTTP " + res.status);
          }
          event.completed({ allowEvent: true });
        });
      });
    })
    .catch(function (err) {
      trace("실패: " + (err && err.message ? err.message : err));
      // 실패하면 발송을 멈추고 "그래도 보내기" 선택창을 띄운다
      event.completed({
        allowEvent: false,
        errorMessage: "API 전송에 실패했습니다: " + (err && err.message ? err.message : err)
      });
    });
}

// manifest.xml 의 FunctionName 과 실제 함수를 연결
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
trace("0-1. 함수 연결 완료");

// 웹/새 Outlook 은 Office 초기화(onReady)가 끝나야 보내기 이벤트 함수를 호출한다
Office.onReady(function (info) {
  trace("0-2. Office 준비됨: " + (info && info.host));
});
