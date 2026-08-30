// 【注意】GitHub公開時は実際のキーを書かないこと！
const LINE_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN';
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';

function doPost(e) {
  // 🔒 LockServiceによる排他制御（同時アクセスの衝突防止）
  const lock = LockService.getScriptLock();
  let replyText = "";
  
  const event = JSON.parse(e.postData.contents).events[0];
  if (event.type !== 'message' || event.message.type !== 'text') return;
  
  const userMessage = event.message.text.trim();
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  // 🤖 AIスタッフ（Gemini 2.5 Flash 搭載）
  function analyzeMessageWithGemini(message, todayStr, eventScheduleStr) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;
    
    const prompt = `あなたはサークルの出欠管理AIです。今日の基準日は「${todayStr}」です。
    
    【サークルのイベントスケジュール】
    ${eventScheduleStr}
    
    ユーザーのメッセージから「日付」「状況」「理由」を抽出し、必ず【JSONの配列（Array）形式】のみで返答してください。
    ユーザーが複数日の連絡をまとめて行っている場合は、1日ごとにオブジェクトに分けて、すべて配列に格納してください。
    
    フォーマット: 
    [
      {"date": "〇月〇日", "status": "欠席" または "遅刻" または "取消", "reason": "理由の文字列"}
      ...
    ]
    ・上記のイベントスケジュールに載っているイベント名がメッセージに含まれていたら、対応する日付（〇月〇日）に自動変換してください。
    ・「明日の集会」などの場合は、基準日から計算して「〇月〇日」にしてください。
    ・具体的な日付（〇月〇日）が特定できない、またはスケジュールにない不明なイベント名の場合は、dateの値を必ず "特定不可" としてください。
    ・取消の場合は reason は空文字 "" にしてください。
    ・出欠連絡と関係ないメッセージの場合は [{"error": true}] と返してください。`;

    const payload = {
      "contents": [{"parts": [{"text": prompt + "\n\nユーザーのメッセージ: " + message}]}],
      "generationConfig": { "temperature": 0.0, "responseMimeType": "application/json" }
    };

    try {
      const response = UrlFetchApp.fetch(url, {
        'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true
      });
      const json = JSON.parse(response.getContentText());
      if (json.error) return [{ "error": true, "debug": "APIエラー: " + json.error.message }];
      
      const text = json.candidates[0].content.parts[0].text;
      const cleanText = text.replace(/```json/g, '').replace(/```/g, '').trim();
      return JSON.parse(cleanText);
    } catch (error) {
      return [{ "error": true, "debug": "処理エラー: " + error.message }];
    }
  }

  try {
    // 最大30秒間、他の人のアクセスを待機させる
    lock.waitLock(30000); 
    
    const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ss = SpreadsheetApp.openById(sheetId);
    const sheetMembers = ss.getSheetByName('名簿');
    const sheetPassword = ss.getSheetByName('パスワード');
    const sheetEvents = ss.getSheetByName('イベント予定'); 

    const currentPassword = sheetPassword.getRange('B1').getDisplayValue();
    const currentDate = sheetPassword.getRange('B2').getDisplayValue() || '未分類の日付';

    // イベント予定の読み込み
    let eventScheduleStr = "";
    if (sheetEvents) {
      const eventData = sheetEvents.getDataRange().getDisplayValues();
      for (let i = 1; i < eventData.length; i++) {
        if (eventData[i][0] && eventData[i][1]) {
          eventScheduleStr += `✨ ${eventData[i][0]} ： ${eventData[i][1]}\n`;
        }
      }
    }

    // 名簿との照合
    const dataMembers = sheetMembers.getDataRange().getValues();
    let isRegistered = false;
    let userName = "", userGrade = "", userKana = "";

    for (let i = 1; i < dataMembers.length; i++) {
      if (dataMembers[i][0] === userId) {
        isRegistered = true;
        userGrade = dataMembers[i][1];
        userName = dataMembers[i][2];
        userKana = dataMembers[i][3] || "";
        break;
      }
    }

    // 出欠記録用の関数（userIdも記録して本人特定を確実にする）
    function recordAttendance(targetSheetName, status, reason) {
      let targetSheet = ss.getSheetByName(targetSheetName);
      if (!targetSheet) {
        targetSheet = ss.insertSheet(targetSheetName);
        targetSheet.appendRow(['氏名', 'ふりがな', '学年', '状況', '理由', '打刻日時', 'userId']);
      }
      const now = new Date();
      targetSheet.appendRow([userName, userKana, userGrade, status, reason, now, userId]);
    }

    // 取消用の関数（userIdで本人を特定。userId未記録の古いデータは氏名でフォールバック）
    function cancelAttendance(targetSheetName) {
      let targetSheet = ss.getSheetByName(targetSheetName);
      if (!targetSheet) return false;

      let deleted = false;
      const data = targetSheet.getDataRange().getValues();
      for (let i = data.length - 1; i > 0; i--) {
        const rowUserId = data[i][6]; // G列：userId
        const isMatch = rowUserId ? (rowUserId === userId) : (data[i][0] === userName);
        if (isMatch) {
          targetSheet.deleteRow(i + 1);
          deleted = true;
        }
      }
      return deleted;
    }

    // 🚦 条件分岐ルーター
    if (!isRegistered && !userMessage.startsWith('登録')) {
      replyText = "初回登録をお願いします！\n「登録 1年 山田太郎」のように送信してください。";
    } 
    else if (userMessage === '出席') {
      replyText = "今日のパスワードを入力してください🔑";
    } 
    else if (userMessage.startsWith('登録')) {
      const parts = userMessage.replace('登録', '').trim().split(/[  ]+/);
      if (parts.length < 2 || !parts[0] || !parts[1]) {
         replyText = "登録エラー😢\n学年と名前が正しく読み取れませんでした。";
      } else {
         const grade = parts[0];
         const name = parts[1];
         sheetMembers.appendRow([userId, grade, name]);
         replyText = `名簿に登録しました！\n（学年: ${grade} / 名前: ${name}）`;
      }
    } 
    else if (userMessage === currentPassword) {
      recordAttendance(currentDate, '出席', '');
      replyText = `${userName}さん、${currentDate}の出席を受け付けました！🎸\n※間違えた場合は「${currentDate} 取消」と送ってください。`;
    } 
    else if (userMessage === '予定表') {
      if (eventScheduleStr === "") {
        replyText = "✨ 予定表 📋 ✨\n\n現在、登録されているサークルの予定はありません。";
      } else {
        replyText = `✨ 予定表 📋 ✨\n\n現在のイベント予定スケジュールです！\n━━━━━━━━━━━━━━\n${eventScheduleStr}━━━━━━━━━━━━━━\n\n※欠席・遅刻連絡は、上記の日付やイベント名を入れてメッセージを送ってね！🎸`;
      }
    } 
    else if (userMessage === 'しまう') {
      replyText = "予定表を閉じました。出欠連絡を続ける場合はメニューをタップしてください✨";
    } 
    else if (userMessage === '欠席' || userMessage === '遅刻' || userMessage === '取消') {
      if (userMessage === '取消') {
         replyText = "取消ですね。\n「3月15日 取消」のように、対象の日付とスペースで区切って送信してください。";
      } else {
         replyText = `${userMessage}ですね！対象の日付と理由を記録するため、\n「3月15日 ${userMessage} 体調不良」\nのように、日付と理由をスペースで区切って送信してください。`;
      }
    } 
    // ⚡ ルートA: マニュアル入力（AI完全回避・高速処理ルート）
    else if (userMessage.match(/^([0-9０-９]{1,2}月[0-9０-９]{1,2}日)[  ]+(欠席|遅刻|取消)(?:[  ]+(.*))?$/)) {
      const match = userMessage.match(/^([0-9０-９]{1,2}月[0-9０-９]{1,2}日)[  ]+(欠席|遅刻|取消)(?:[  ]+(.*))?$/);
      const dateStr = match[1];
      const statusStr = match[2];
      const reasonStr = match[3] || "";
      
      if (statusStr === '取消') {
         const isCanceled = cancelAttendance(dateStr);
         if (isCanceled) {
             replyText = `【${dateStr}】の連絡を取り消しました！`;
         } else {
             replyText = `該当する日付の連絡が見つからなかったか、すでに取り消されています。`;
         }
      } else {
         recordAttendance(dateStr, statusStr, reasonStr);
         replyText = `【${dateStr}(${statusStr})】の連絡を受け付けました！✨`;
      }
    }
    // 🤖 ルートB: 自由入力メッセージ（Gemini API解析ルート）
    else {
      // ログに残す工夫
      console.log("🤖 ルートB(Gemini)へ突撃: " + userMessage); 

      const today = new Date();
      const todayStr = (today.getMonth() + 1) + '月' + today.getDate() + '日';
      
      const aiResults = analyzeMessageWithGemini(userMessage, todayStr, eventScheduleStr);

      if (!Array.isArray(aiResults) || (aiResults[0] && aiResults[0].error)) {
         if (aiResults[0] && aiResults[0].debug) {
             replyText = "【システム調査中】AI連携でエラーが発生しました。\n原因: " + aiResults[0].debug;
         } else {
             replyText = "メッセージを確認できませんでした。\n・出席する場合はパスワードを入力\n・欠席/遅刻は理由を添えて送信してください。";
         }
      } else {
         let successDetails = [];
         let cancelDetails = [];
         let hasInvalidDate = false; 

         for (let i = 0; i < aiResults.length; i++) {
           const res = aiResults[i];
           const dateStr = res.date;
           
           const dateRegex = /^[0-9０-９]{1,2}月[0-9０-９]{1,2}日$/;
           if (!dateStr || dateStr === "特定不可" || !dateRegex.test(dateStr)) {
              hasInvalidDate = true;
              break; 
           }
         }

         if (hasInvalidDate) {
            replyText = "日程を特定できませんでした😢\n「〇月〇日」とはっきりと日付を書くか、登録されているイベント名（例：夏定期1日目）を使って、もう一度送信してください！";
         } else {
            for (let i = 0; i < aiResults.length; i++) {
              const res = aiResults[i];
              const dateStr = res.date;
              const statusStr = res.status;
              const reasonStr = res.reason;

              if (statusStr === '取消') {
                 const isCanceled = cancelAttendance(dateStr);
                 if (isCanceled) cancelDetails.push(dateStr);
              } else {
                 recordAttendance(dateStr, statusStr, reasonStr);
                 successDetails.push(`${dateStr}(${statusStr})`);
              }
            }

            let replyParts = [];
            if (successDetails.length > 0) {
               replyParts.push(`【${successDetails.join(', ')}】の連絡を受け付けました！`);
            }
            if (cancelDetails.length > 0) {
               replyParts.push(`【${cancelDetails.join(', ')}】の連絡を取り消しました！`);
            }
            
            if (replyParts.length > 0) {
               replyText = `AIが読み取りました🤖✨\n` + replyParts.join('\n');
            } else {
               replyText = `該当する日付の連絡が見つからなかったか、すでに取り消されています。`;
            }
         }
      }
    }
  } catch (e) {
    replyText = "アクセスが集中しています💦 スプレッドシートが混雑しているので、数秒待ってからもう一度送信してください！";
  } finally {
    // 処理が終わったら必ずロックを解除
    lock.releaseLock();
  }

  // LINEへ返信
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{'type': 'text', 'text': replyText}],
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}
