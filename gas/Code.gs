/**
 * システム設計仕様書に基づいた備品管理バックエンド
 * 2026/03/16 刷新
 */

const SPREADSHEET_ID = '1996BJT0IJoHYebMcoQaB0V6JNerrgnlyOCJOtUACT94';

// シート定数
const SHEETS = {
  ASSETS: 'T_Assets',
  HISTORY: 'T_History',
  DEPTS: 'M_Depts',
  CATEGORIES: 'M_Categories'
};

// ステータス・タイプ定数
const STATUS = {
  ACTIVE: '稼働中',
  FAULTY: '故障中',
  SPARE: '予備',
  WAIT_DISPOSAL: '廃棄待',
  DISPOSED: '廃棄済'
};

const LOG_TYPES = {
  REPAIR: '修理',
  INSPECTION: '点検',
  MOVE: '移動',
  INVENTORY: '棚卸',
  DISPOSAL: '廃棄'
};

/**
 * Webアプリのエントリポイント
 * 全てSPAとして index.html を返す
 * @param {Object} e - URLパラメータ (tokenまたはidを受け取る)
 */
function doGet(e) {
  const token = e?.parameter?.token || '';
  const id = e?.parameter?.id || '';
  
  try {
    const template = HtmlService.createTemplateFromFile('index');
    template.token = token;
    template.id = id;
    template.appUrl = ScriptApp.getService().getUrl();
    
    return template.evaluate()
                   .setTitle('病院備品管理システム')
                   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
                   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(error) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif; padding:20px;">システムの読み込みに失敗しました。</div>');
  }
}

/**
 * トークンまたはIDに基づいて備品情報と履歴を取得
 * @param {string} token - ユニークなアクセス権限トークン (QRコード由来)
 * @param {string} id - システム内部ID
 */
function getAssetData(token, id) {
  console.log('getAssetData call with token:', token, 'id:', id);
  if (!token && !id) return null;
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. 備品データの検索 (T_Assets)
    const assetsSheet = ss.getSheetByName(SHEETS.ASSETS);
    if (!assetsSheet) throw new Error('Sheet "' + SHEETS.ASSETS + '" not found');
    const assets = getSheetDataAsObjects(assetsSheet);
    
    // キーの正規化（日本語ヘッダー対応）
    const normalizedAssets = assets.map(a => normalizeKeys(a, 'ASSETS'));
    
    let asset = null;
    if (token) {
      asset = normalizedAssets.find(a => a.qr_token === token);
    } else if (id) {
      asset = normalizedAssets.find(a => String(a.id) === String(id));
    }
    
    console.log('Asset found:', asset ? asset.name : 'not found');
    if (!asset) return null;

    // 2. 履歴データの取得 (T_History)
    const historySheet = ss.getSheetByName(SHEETS.HISTORY);
    if (!historySheet) throw new Error('Sheet "' + SHEETS.HISTORY + '" not found');
    const allHistory = getSheetDataAsObjects(historySheet);
    
    // 該当備品の履歴を抽出して日付降順に
    const history = allHistory
      .map(h => normalizeKeys(h, 'HISTORY'))
      .filter(h => h.asset_id === asset.id)
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
      .map(h => ({
        ...h,
        timestamp: formatDate(h.timestamp)
      }));

    asset.history = history;
    asset.purchase_date = formatDate(asset.purchase_date);

    return asset;
  } catch (e) {
    console.error('getAssetData Error:', e.toString());
    throw e; // クライアント側にエラーを投げ返す
  }
}

/**
 * 日本語ヘッダーをプログラム用キーに変換
 */
function normalizeKeys(obj, type) {
  const mapping = {
    ASSETS: {
      'ID': 'id', 'システム識別ID': 'id', 'id': 'id',
      '型番': 'model_number', 'model_number': 'model_number',
      '備品管理番号': 'asset_tag', '管理ID': 'asset_tag', 'asset_tag': 'asset_tag',
      'カテゴリ名': 'category_id', 'カテゴリID': 'category_id', 'category_id': 'category_id',
      '備品名称': 'name', '名称': 'name', 'name': 'name',
      '現在の状態': 'status', 'ステータス': 'status', 'status': 'status',
      '管理部署名': 'dept_id', '部署ID': 'dept_id', 'dept_id': 'dept_id',
      '設置階': 'floor', 'floor': 'floor',
      '設置場所': 'location', '場所': 'location', 'location': 'location',
      'QRアクセスキー': 'qr_token', 'トークン': 'qr_token', 'qr_token': 'qr_token',
      '説明書リンク': 'manual_url', 'マニュアルURL': 'manual_url', 'manual_url': 'manual_url',
      '購入年月日': 'purchase_date', '購入日': 'purchase_date', 'purchase_date': 'purchase_date',
      '購入業者': 'vendor', 'vendor': 'vendor',
      '購入金額': 'price', '価格': 'price', 'price': 'price',
      '耐用年数': 'useful_life', 'useful_life': 'useful_life',
      '修理依頼先': 'repair_vendor', 'repair_vendor': 'repair_vendor',
      '備考': 'note', 'note': 'note'
    },
    HISTORY: {
      '報告日時': 'timestamp', '日時': 'timestamp', 'timestamp': 'timestamp',
      '対象備品ID': 'asset_id', '備品ID': 'asset_id', 'asset_id': 'asset_id',
      '報告種別': 'type', '種別': 'type', 'type': 'type',
      '担当者名': 'operator', '担当者': 'operator', 'operator': 'operator',
      '備考・メモ': 'note', '備考': 'note', 'note': 'note'
    }
  };
  
  const rules = mapping[type];
  const newObj = {};
  for (let key in obj) {
    if (!key) continue;
    const cleanKey = String(key).trim();
    const normalizedKey = rules[cleanKey] || cleanKey;
    newObj[normalizedKey] = obj[key];
  }
  return newObj;
}

/**
 * 全備品リストを取得 (一覧画面用)
 */
function getAssetList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assetsSheet = ss.getSheetByName(SHEETS.ASSETS);
    if (!assetsSheet) throw new Error('Sheet not found');
    const assets = getSheetDataAsObjects(assetsSheet);
    
    // 不要な情報の除外・フロント用の整形
    return assets.map(a => normalizeKeys(a, 'ASSETS'))
                 .map(a => ({
                   id: a.id,
                   asset_tag: a.asset_tag,
                   category_id: a.category_id,
                   name: a.name,
                   model_number: a.model_number,
                   status: a.status,
                   dept_id: a.dept_id,
                   floor: a.floor,
                   location: a.location,
                   qr_token: a.qr_token
                 }));
  } catch (e) {
    console.error('getAssetList Error:', e.toString());
    throw e;
  }
}

/**
 * 備品管理番号の自動生成
 * カテゴリ名 & 購入年月日の西暦下2桁+月2桁
 */
function generateAssetTag(categoryId, purchaseDateString) {
  const cat = categoryId || '未分類';
  let datePart = '----';
  
  if (purchaseDateString) {
    const d = new Date(purchaseDateString);
    if (!isNaN(d.getTime())) {
      const yy = String(d.getFullYear()).slice(-2);
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      datePart = `${yy}${mm}`;
    }
  }
  return `${cat}${datePart}`;
}

/**
 * 新規備品の登録処理

 */
function registerNewAsset(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assetsSheet = ss.getSheetByName(SHEETS.ASSETS);
    
    // ヘッダー取得してインデックスに合わせて配列を作成
    const headers = assetsSheet.getDataRange().getValues()[0];
    const newRow = new Array(headers.length).fill('');
    
    // IDは行番号-1とする
    const newId = assetsSheet.getLastRow(); 
    // 16文字のランダムな文字列をトークンとして生成
    const newToken = Utilities.getUuid().replace(/-/g, '').substring(0, 16); 
    
    // 逆引き用マッピング (プログラム用キー -> 予想される日本語ヘッダーの配列)
    const reverseMapping = {
      'id': ['ID', 'システム識別ID', 'id'],
      'model_number': ['型番', 'model_number'],
      'asset_tag': ['備品管理番号', '管理ID', 'asset_tag'],
      'category_id': ['カテゴリ名', 'カテゴリID', 'category_id'],
      'name': ['備品名称', '名称', 'name'],
      'status': ['現在の状態', 'ステータス', 'status'],
      'dept_id': ['管理部署名', '部署ID', 'dept_id'],
      'floor': ['設置階', 'floor'],
      'location': ['設置場所', '場所', 'location'],
      'qr_token': ['QRアクセスキー', 'トークン', 'qr_token'],
      'manual_url': ['説明書リンク', 'マニュアルURL', 'manual_url'],
      'purchase_date': ['購入年月日', '購入日', 'purchase_date'],
      'vendor': ['購入業者', 'vendor'],
      'price': ['購入金額', '価格', 'price'],
      'useful_life': ['耐用年数', 'useful_life'],
      'repair_vendor': ['修理依頼先', 'repair_vendor'],
      'note': ['備考', 'note']
    };
    
    // 備品管理番号を自動計算
    data.asset_tag = generateAssetTag(data.category_id, data.purchase_date);
    
    const insertData = { ...data, id: newId, qr_token: newToken, status: STATUS.ACTIVE };

    headers.forEach((h, index) => {
      const cleanHeader = String(h).trim();
      for (const [key, possibleHeaders] of Object.entries(reverseMapping)) {
        if (possibleHeaders.includes(cleanHeader)) {
          newRow[index] = insertData[key] || '';
          break;
        }
      }
    });
    
    assetsSheet.appendRow(newRow);
    
    // 履歴シートにも初期登録のログを残す
    const historySheet = ss.getSheetByName(SHEETS.HISTORY);
    historySheet.appendRow([
      new Date(),
      newId,
      '点検', // 登録という種別がないため初期点検相当とする
      insertData.operator || 'システム',
      '新規登録'
    ]);
    
    return { success: true, token: newToken };
  } catch (e) {
    console.error('registerNewAsset Error:', e);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 既存備品データの直接更新 (単票フォームでの修正保存用)
 */
function updateAsset(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assetsSheet = ss.getSheetByName(SHEETS.ASSETS);
    
    const assetsData = assetsSheet.getDataRange().getValues();
    const headers = assetsData[0];
    
    // idで検索
    let rowIndex = -1;
    for (let i = 1; i < assetsData.length; i++) {
      if (assetsData[i][0] === data.id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: '指定された備品が見つかりません。' };
    }

    // 逆引き用マッピング (プログラム用キー -> 予想される日本語ヘッダーの配列)
    const reverseMapping = {
      'model_number': ['型番', 'model_number'],
      'asset_tag': ['備品管理番号', '管理ID', 'asset_tag'],
      'category_id': ['カテゴリ名', 'カテゴリID', 'category_id'],
      'name': ['備品名称', '名称', 'name'],
      'status': ['現在の状態', 'ステータス', 'status'],
      'dept_id': ['管理部署名', '部署ID', 'dept_id'],
      'floor': ['設置階', 'floor'],
      'location': ['設置場所', '場所', 'location'],
      'manual_url': ['説明書リンク', 'マニュアルURL', 'manual_url'],
      'purchase_date': ['購入年月日', '購入日', 'purchase_date'],
      'vendor': ['購入業者', 'vendor'],
      'price': ['購入金額', '価格', 'price'],
      'useful_life': ['耐用年数', 'useful_life'],
      'repair_vendor': ['修理依頼先', 'repair_vendor'],
      'note': ['備考', 'note']
    };

    // 備品管理番号を自動計算で上書き
    data.asset_tag = generateAssetTag(data.category_id, data.purchase_date);

    // 編集された項目のみを更新する
    headers.forEach((h, index) => {
      const cleanHeader = String(h).trim();
      for (const [key, possibleHeaders] of Object.entries(reverseMapping)) {
        if (possibleHeaders.includes(cleanHeader) && data[key] !== undefined) {
          assetsSheet.getRange(rowIndex, index + 1).setValue(data[key]);
          break;
        }
      }
    });
    
    // 更新履歴を残す
    const historySheet = ss.getSheetByName(SHEETS.HISTORY);
    historySheet.appendRow([
      new Date(),
      data.id,
      '点検', // マスタ修正のため点検扱い
      data.operator || 'システム',
      '台帳情報修正'
    ]);
    
    return { success: true };
  } catch (e) {
    console.error('updateAsset Error:', e);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


/**
 * 報告の登録処理
 * @param {Object} data - 報告内容 (asset_id, type, operator, note, status, new_floor, new_location 等)
 */
function registerReport(data) {
  const lock = LockService.getScriptLock();
  try {
    // 30秒間ロックを試みる
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();

    // 1. T_Historyへの追記
    const historySheet = ss.getSheetByName(SHEETS.HISTORY);
    historySheet.appendRow([
      now,
      data.asset_id,
      data.type,
      data.operator,
      data.note || ''
    ]);

    // 2. T_Assetsの更新 (ステータス、場所など)
    const assetsSheet = ss.getSheetByName(SHEETS.ASSETS);
    const assetsData = assetsSheet.getDataRange().getValues();
    const headers = assetsData[0];
    
    // asset_id(A列)で検索
    let rowIndex = -1;
    for (let i = 1; i < assetsData.length; i++) {
      if (assetsData[i][0] === data.asset_id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      // ステータス列を更新
      const statusIdx = headers.indexOf('status');
      if (statusIdx !== -1) assetsSheet.getRange(rowIndex, statusIdx + 1).setValue(data.status);
      
      // 移動・棚卸の場合は場所も更新
      if (data.type === LOG_TYPES.MOVE || data.type === LOG_TYPES.INVENTORY) {
        const floorIdx = headers.indexOf('floor');
        const locIdx = headers.indexOf('location');
        if (floorIdx !== -1 && data.new_floor) assetsSheet.getRange(rowIndex, floorIdx + 1).setValue(data.new_floor);
        if (locIdx !== -1 && data.new_location) assetsSheet.getRange(rowIndex, locIdx + 1).setValue(data.new_location);
      }
    }

    return { success: true };

  } catch (e) {
    console.error(e);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 補助関数: シートデータをObjectの配列に変換
 */
function getSheetDataAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    console.warn(`Sheet ${sheet.getName()} is empty or has no data.`);
    return [];
  }
  
  const headers = data[0];
  console.log(`Sheet ${sheet.getName()} headers:`, headers);
  
  const rows = data.slice(1);
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      // header が文字列でない、または空の列などの場合はスキップするか、そのまま文字列表現に
      const key = (header !== undefined && header !== null) ? String(header).trim() : `Col${index}`;
      obj[key] = row[index];
    });
    return obj;
  });
}

/**
 * 日付フォーマット
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return date;
  return Utilities.formatDate(date, "JST", "yyyy/MM/dd HH:mm");
}
