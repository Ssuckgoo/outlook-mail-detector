// 메일 보내기 버튼을 누르면 Outlook 이 실행하는 코드
// 클래식 Outlook 은 이 파일을 브라우저가 아닌 JS 전용 엔진으로 직접 실행하므로
// 최신 문법(async/await 등) 대신 Promise 와 일반 함수로 작성한다.

// 본문에서 찾을 텍스트 규칙: "#요청:A001" -> "A001"
var PATTERN = /#요청:(\S+)/g;

// 찾은 내용을 보낼 API 주소 (테스트용: server.js 의 가짜 API)
var TARGET_API = "https://localhost:3000/api/hook";

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
  var item = Office.context.mailbox.item;

  getAsync(item.body, Office.CoercionType.Text)
    .then(function (body) {
      var matches = findMatches(body);
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

        return fetch(TARGET_API, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        }).then(function (res) {
          if (!res.ok) {
            throw new Error("HTTP " + res.status);
          }
          event.completed({ allowEvent: true });
        });
      });
    })
    .catch(function (err) {
      // 실패하면 발송을 멈추고 "그래도 보내기" 선택창을 띄운다
      event.completed({
        allowEvent: false,
        errorMessage: "API 전송에 실패했습니다: " + (err && err.message ? err.message : err)
      });
    });
}

// manifest.xml 의 FunctionName 과 실제 함수를 연결
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
