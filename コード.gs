function onFormSubmit(e) {
  // エディタから直接実行された場合のクラッシュを防ぐ
  if (!e) {
    showMessage("エラー: この関数はフォーム送信時に自動実行されます。エディタから直接実行することはできません。");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ====================================================
  // フォームの回答1シートから最新行のデータを取得
  // ====================================================
  const latestRowData = e.values;

  // 動的に列の位置を取得するために、ヘッダーを取得
  const formSheet = ss.getSheetByName("フォームの回答 1");
  if (!formSheet) {
    showMessage("エラー: 「フォームの回答 1」シートが見つかりません。");
    return;
  }
  const header = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

  // 各項目の列インデックスをすべて動的に取得
  const nameIndex = header.indexOf("氏名");
  const deptIndex = header.indexOf("部署");
  const destIndex = header.indexOf("出張先");
  const startIndex = header.indexOf("出張開始日");
  const endIndex = header.indexOf("出張終了日");
  const purposeIndex = header.indexOf("出張目的");
  const transitIndex = header.indexOf("交通費（円）");
  const lodgingIndex = header.indexOf("宿泊費（円）");
  const otherIndex = header.indexOf("その他経費（円）");
  const remarkIndex = header.indexOf("備考");

  if (transitIndex === -1 || lodgingIndex === -1 || otherIndex === -1 || nameIndex === -1 || remarkIndex === -1) {
    showMessage("エラー: フォーム回答シートに必要な列が見つかりません。ヘッダー名を確認してください。");
    return;
  }

  // ====================================================
  // 自動採番ロジック（年月-3桁連番）
  // ====================================================
  const historySheet = ss.getSheetByName("申請履歴");
  if (!historySheet) {
    showMessage("エラー: 「申請履歴」シートが見つかりません。");
    return;
  }

  // 現在の年月を取得 (例: "202606")
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0"); // 1月が0から始まるため+1し、2桁にゼロ埋め
  const currentYearMonth = `${year}${month}`;

  // 申請履歴シートのA列（申請番号）から最大連番を探す
  const lastRowHistory = historySheet.getLastRow();
  let maxSerialNumber = 0;

  if (lastRowHistory > 1) {
    // A列の2行目から最終行までの「申請番号」をまとめて取得
    const appNumbers = historySheet.getRange(2, 1, lastRowHistory - 1, 1).getValues();
    
    for (let i = 0; i < appNumbers.length; i++) {
      const fullNumber = appNumbers[i][0].toString(); // 例: "202606-002"
      
      // 現在の年月から始まっている場合のみ連番をチェック
      if (fullNumber.startsWith(currentYearMonth)) {
        // ハイフンの後ろの文字列（連番部分）を切り出して数値化
        const serialPart = Number(fullNumber.split("-")[1]);
        if (serialPart > maxSerialNumber) {
          maxSerialNumber = serialPart; // 最大値を更新
        }
      }
    }
  }

  // 最大連番+1で新しい番号を生成し、3桁にゼロ埋め
  const newSerialNumber = String(maxSerialNumber + 1).padStart(3, "0");
  const applicationNumber = `${currentYearMonth}-${newSerialNumber}`;

  // ====================================================
  // 合計金額を計算
  // ====================================================
  const transitCost = Number(latestRowData[transitIndex]);
  const lodgingCost = Number(latestRowData[lodgingIndex]);
  const otherCost = Number(latestRowData[otherIndex]);

  const totalAmount = transitCost + lodgingCost + otherCost;

  // ====================================================
  // Slackメッセージ用変数の定義
  // ====================================================
  const applicantName = latestRowData[nameIndex].toString().trim();
  const department = deptIndex !== -1 ? latestRowData[deptIndex] : "未設定";
  const destination = destIndex !== -1 ? latestRowData[destIndex] : "未設定";
  const startDate = startIndex !== -1 ? latestRowData[startIndex] : "-";
  const endDate = endIndex !== -1 ? latestRowData[endIndex] : "-";
  const purpose = purposeIndex !== -1 ? latestRowData[purposeIndex] : "未設定";
  const remarkData = latestRowData[remarkIndex];
  const remark = remarkData !== "" ? remarkData : "特になし";
  const fTotal = totalAmount.toLocaleString();

  // ====================================================
  // 「申請履歴」シートに1行追記（申請番号を先頭、合計経費を備考の前に挿入）
  // ====================================================
  // 備考より前のデータ、備考のデータを切り分けて組み替え
  const beforeRemarkData = latestRowData.slice(0, remarkIndex);
  
  // 先頭に [applicationNumber] を、備考の前に「totalAmount」を挿入する
  const outputRowData = [applicationNumber, ...beforeRemarkData, totalAmount, remarkData];

  // 申請履歴シートに追記
  historySheet.appendRow(outputRowData);

  // ====================================================
  // Slackに通知を送信
  // ====================================================
  const slackMessage = `【出張経費申請】申請番号：${applicationNumber}
  申請者：${applicantName}（${department}）
  出張先：${destination}
  出張期間：${startDate} ～ ${endDate}
  出張目的：${purpose}
  合計金額：${fTotal}円
  備考：${remark}`;

  sendToSlack(slackMessage);
}

// =====================
// Slack通知ヘルパー関数
// =====================
function sendToSlack(message) {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");
  const channelId = "C0BB2B0EV6Z";
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelId,
      text: message
    })
  };
  
  const response = UrlFetchApp.fetch(
    "https://slack.com/api/chat.postMessage", options);
  console.log(response.getContentText());
}
