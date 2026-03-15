// スプレッドシートのIDとシート名（後で実際の値に置き換えます）
const SPREADSHEET_ID = '1996BJT0IJoHYebMcoQaB0V6JNerrgnlyOCJOtUACT94';
const SHEET_NAME_MASTER = '備品マスタ';
const SHEET_NAME_HISTORY = '履歴データ';

function doGet(e) {
  const params = e ? e.parameter : {};
  const itemId = params.id || '';
  
  const template = HtmlService.createTemplateFromFile('index');
  template.itemId = itemId;
  
  return template.evaluate()
                 .setTitle('備品カルテシステム')
                 .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 備品データを取得する関数
function getItemData(itemId) {
  if (!itemId) return null;
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  const masterData = masterSheet.getDataRange().getValues();
  
  let item = null;
  // ヘッダーを飛ばして検索
  const searchId = String(itemId).trim();
  for (let i = 1; i < masterData.length; i++) {
    const rowId = String(masterData[i][0]).trim();
    if (rowId == searchId) {
      item = {
        id: rowId,
        name: masterData[i][1],
        status: masterData[i][2],
        department: masterData[i][3],
        purchaseDate: masterData[i][4] instanceof Date ? Utilities.formatDate(masterData[i][4], "JST", "yyyy/MM/dd") : masterData[i][4],
        lastInspection: masterData[i][5] instanceof Date ? Utilities.formatDate(masterData[i][5], "JST", "yyyy/MM/dd") : masterData[i][5],
        nextInspection: masterData[i][6] instanceof Date ? Utilities.formatDate(masterData[i][6], "JST", "yyyy/MM/dd") : masterData[i][6],
        manualUrl: masterData[i][7],
        history: []
      };
      break;
    }
  }
  
  if (!item) return null;

  // 履歴データを取得
  const historySheet = ss.getSheetByName(SHEET_NAME_HISTORY);
  const historyData = historySheet.getDataRange().getValues();
  for (let j = 1; j < historyData.length; j++) {
    const histId = String(historyData[j][1]).trim();
    if (histId == searchId) {
      item.history.push({
        date: historyData[j][0] instanceof Date ? Utilities.formatDate(historyData[j][0], "JST", "yyyy/MM/dd") : historyData[j][0],
        content: historyData[j][2],
        user: historyData[j][3]
      });
    }
  }
  
  // 履歴を日付の新しい順にソート
  item.history.reverse();
  
  return item;
}

// 点検・修理報告を登録する関数
function registerReport(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. 履歴データへの追記
  const historySheet = ss.getSheetByName(SHEET_NAME_HISTORY);
  const today = new Date();
  historySheet.appendRow([
    today,
    data.itemId,
    data.content,
    data.user
  ]);
  
  // 2. 備品マスタの更新（ステータスと最終点検日）
  const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] == data.itemId) {
      masterSheet.getRange(i + 1, 3).setValue(data.status); // ステータス
      masterSheet.getRange(i + 1, 6).setValue(today);      // 最終点検日
      break;
    }
  }
  
  return { success: true };
}
