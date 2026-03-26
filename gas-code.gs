/**
 * ME機器管理システム - Google Apps Script バックエンド
 *
 * このスクリプトをGoogle Spreadsheetの「拡張機能 > Apps Script」に貼り付けてください。
 * デプロイ後のURLをアプリのGAS_URL設定に入力してください。
 *
 * スプレッドシートID: 1hORasvyl7H0KcVy33AhUfISv3-1HXatxduBXHuijspM
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1hORasvyl7H0KcVy33AhUfISv3-1HXatxduBXHuijspM';

// シート名の定義
const SHEET_NAMES = {
  equipment: '機器マスター',    // メインの機器マスターシート
  lending: '貸出履歴',          // 貸出履歴シート
  inspection: '点検記録',       // 点検記録シート
  repair: '修理記録',           // 修理記録シート
  staff: 'スタッフ',            // スタッフマスターシート
  departments: '部署',          // 部署マスターシート
  areas: '貸出エリア',           // 貸出エリアシート
  items: '物品マスター'          // 物品マスターシート
};

// ===== ヘルパー関数 =====

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * シートの全データをJSON配列として取得
 */
function getSheetData(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };

  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      const val = data[i][j];
      row[headers[j]] = (val === null || val === undefined) ? '' : String(val);
      if (val !== '' && val !== null && val !== undefined) hasData = true;
    }
    if (hasData) rows.push(row);
  }
  return { headers, rows };
}

/**
 * ヘッダー行のカラムマッピングを取得
 */
function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h] = i + 1; });
  return { headers, map };
}

/**
 * 機器マスターのヘッダーとフィールドのマッピング
 * スプレッドシートの列名 → アプリ内フィールド名
 */
function getEquipmentFieldMap() {
  return {
    // スプレッドシート列名: アプリフィールド名
    '機器ID': 'id',
    '機器区分': 'category',
    '機器名': 'name',
    '型式': 'model',
    '製造番号': 'serial',
    '管理部署': 'dept',
    '製造元': 'manufacturer',
    '代理店': 'dealer',
    'ステータス': 'status',
    '貸出状態': 'lending',
    '現在地': 'location',
    '保管場所': 'storage',
    '廃棄': 'discard',
    'CH': 'channel',
    '資産区分': 'assetType',
    '納入日': 'deliveryDate',
    '購入価格': 'purchasePrice',
    '代替レンタルデモ開始日': 'rentalStartDate',
    '代替レンタルデモ終了日': 'rentalEndDate',
    '保守契約': 'maintenanceContract',
    '廃棄日': 'discardDate',
    '廃棄理由': 'discardReason',
    '添付文書URL': 'documentUrl',
    '備考': 'notes',
    // 以下、スプレッドシートに追加の列がある場合も自動対応
  };
}

/**
 * アプリフィールド名 → スプレッドシート列名の逆マッピング
 */
function getReverseFieldMap() {
  const map = getEquipmentFieldMap();
  const reverse = {};
  for (const [sheetCol, appField] of Object.entries(map)) {
    reverse[appField] = sheetCol;
  }
  return reverse;
}

// ===== GET リクエスト =====
function doGet(e) {
  const action = e.parameter.action || 'getEquipment';
  let result;

  try {
    switch (action) {
      case 'getEquipment':
        result = handleGetEquipment();
        break;
      case 'getAll':
        result = handleGetAll();
        break;
      case 'getHeaders':
        result = handleGetHeaders();
        break;
      case 'getLending':
        result = handleGetLending();
        break;
      case 'getInspections':
        result = handleGetInspections();
        break;
      case 'getRepairs':
        result = handleGetRepairs();
        break;
      case 'getStaff':
        result = handleGetStaff();
        break;
      case 'getDepartments':
        result = handleGetDepartments();
        break;
      case 'getAreas':
        result = handleGetAreas();
        break;
      case 'getItems':
        result = handleGetItems();
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== POST リクエスト =====
function doPost(e) {
  let result;

  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case 'saveEquipment':
        result = handleSaveEquipment(body.data);
        break;
      case 'saveEquipmentBatch':
        result = handleSaveEquipmentBatch(body.data);
        break;
      case 'deleteEquipment':
        result = handleDeleteEquipment(body.id);
        break;
      case 'addLending':
        result = handleAddLending(body.data);
        break;
      case 'addInspection':
        result = handleAddInspection(body.data);
        break;
      case 'addRepair':
        result = handleAddRepair(body.data);
        break;
      case 'saveStaff':
        result = handleSaveStaff(body.data);
        break;
      case 'deleteStaff':
        result = handleDeleteStaff(body.id);
        break;
      case 'saveDepartment':
        result = handleSaveDepartment(body.data);
        break;
      case 'deleteDepartment':
        result = handleDeleteDepartment(body.id);
        break;
      case 'saveArea':
        result = handleSaveArea(body.data);
        break;
      case 'deleteArea':
        result = handleDeleteArea(body.id);
        break;
      case 'saveItem':
        result = handleSaveItem(body.data);
        break;
      case 'deleteItem':
        result = handleDeleteItem(body.id);
        break;
      case 'syncAll':
        result = handleSyncAll(body.data);
        break;
      case 'updateEquipmentField':
        result = handleUpdateEquipmentField(body.id, body.field, body.value);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 機器マスター =====

function handleGetEquipment() {
  const sheetName = SHEET_NAMES.equipment;
  const sheet = getOrCreateSheet(sheetName);

  if (sheet.getLastRow() < 1) {
    // ヘッダーがない場合、デフォルトヘッダーを作成
    const defaultHeaders = ['機器ID','機器区分','機器名','型式','製造番号','管理部署','製造元','代理店','ステータス','貸出状態','現在地','保管場所','CH','資産区分','納入日','購入価格','代替レンタルデモ開始日','代替レンタルデモ終了日','保守契約','廃棄','廃棄日','廃棄理由','添付文書URL','備考'];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    return { success: true, data: [], headers: defaultHeaders };
  }

  const result = getSheetData(sheetName);
  const fieldMap = getEquipmentFieldMap();

  // スプレッドシートの列名をアプリのフィールド名に変換
  const converted = result.rows.map(row => {
    const obj = {};
    for (const [sheetCol, value] of Object.entries(row)) {
      const appField = fieldMap[sheetCol] || sheetCol;
      obj[appField] = value;
    }
    return obj;
  });

  return {
    success: true,
    data: converted,
    headers: result.headers,
    fieldMap: fieldMap
  };
}

function handleGetHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  if (sheet.getLastRow() < 1) return { success: true, headers: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { success: true, headers: headers };
}

function handleSaveEquipment(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  const { headers, map } = getHeaderMap(sheet);
  const reverseMap = getReverseFieldMap();

  // 機器ID列を探す
  const idCol = map['機器ID'];
  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };

  // 既存行を検索
  const lastRow = sheet.getLastRow();
  let targetRow = -1;

  if (lastRow > 1) {
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    // 新規追加
    targetRow = lastRow + 1;
  }

  // データを書き込む
  for (const [appField, value] of Object.entries(data)) {
    const sheetCol = reverseMap[appField] || appField;
    const col = map[sheetCol];
    if (col) {
      sheet.getRange(targetRow, col).setValue(value || '');
    } else {
      // スプレッドシートにない列は、カスタムフィールドとして末尾に追加
      // ただし、マッピングにない内部フィールドはスキップ
    }
  }

  return { success: true, message: '保存しました: ' + data.id };
}

function handleSaveEquipmentBatch(dataArray) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  const { headers, map } = getHeaderMap(sheet);
  const reverseMap = getReverseFieldMap();
  const idCol = map['機器ID'];

  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };

  // 既存IDマップを作成
  const lastRow = sheet.getLastRow();
  const existingIds = {};
  if (lastRow > 1) {
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    ids.forEach((row, i) => { existingIds[String(row[0])] = i + 2; });
  }

  let nextRow = lastRow + 1;
  const updates = [];
  const newRows = [];

  for (const data of dataArray) {
    const rowData = new Array(headers.length).fill('');

    for (const [appField, value] of Object.entries(data)) {
      const sheetCol = reverseMap[appField] || appField;
      const colIdx = map[sheetCol];
      if (colIdx) {
        rowData[colIdx - 1] = value || '';
      }
    }

    if (existingIds[data.id]) {
      updates.push({ row: existingIds[data.id], data: rowData });
    } else {
      newRows.push(rowData);
    }
  }

  // 既存行更新
  for (const update of updates) {
    sheet.getRange(update.row, 1, 1, headers.length).setValues([update.data]);
  }

  // 新規行追加
  if (newRows.length > 0) {
    sheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { success: true, message: `更新:${updates.length}件, 新規:${newRows.length}件` };
}

function handleDeleteEquipment(id) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  const { map } = getHeaderMap(sheet);
  const idCol = map['機器ID'];
  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'データがありません' };

  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { success: true, message: id + ' を削除しました' };
    }
  }

  return { success: false, error: id + ' が見つかりません' };
}

function handleUpdateEquipmentField(id, field, value) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  const { headers, map } = getHeaderMap(sheet);
  const reverseMap = getReverseFieldMap();
  const idCol = map['機器ID'];

  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };

  const sheetCol = reverseMap[field] || field;
  const targetCol = map[sheetCol];
  if (!targetCol) return { success: false, error: sheetCol + '列が見つかりません' };

  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.getRange(i + 2, targetCol).setValue(value || '');
      return { success: true };
    }
  }

  return { success: false, error: id + ' が見つかりません' };
}

// ===== 貸出履歴 =====

function handleGetLending() {
  const sheet = getOrCreateSheet(SHEET_NAMES.lending);
  if (sheet.getLastRow() < 1) {
    const headers = ['日時','機器ID','機器名','操作','貸出先','返却元','操作者'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.lending);
  return { success: true, data: result.rows };
}

function handleAddLending(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.lending);
  if (sheet.getLastRow() < 1) {
    const headers = ['日時','機器ID','機器名','操作','貸出先','返却元','操作者'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const row = [
    data.datetime || new Date().toISOString(),
    data.equipmentId || '',
    data.equipmentName || '',
    data.action || '',
    data.destination || '',
    data.returnFrom || '',
    data.operator || ''
  ];

  sheet.appendRow(row);
  return { success: true };
}

// ===== 点検記録 =====

function handleGetInspections() {
  const sheet = getOrCreateSheet(SHEET_NAMES.inspection);
  if (sheet.getLastRow() < 1) {
    const headers = ['点検ID','機器ID','機器名','点検日','点検者','点検種別','結果','備考'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.inspection);
  return { success: true, data: result.rows };
}

function handleAddInspection(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.inspection);
  if (sheet.getLastRow() < 1) {
    const headers = ['点検ID','機器ID','機器名','点検日','点検者','点検種別','結果','備考'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const row = [
    data.inspectionId || ('INS-' + Date.now()),
    data.equipmentId || '',
    data.equipmentName || '',
    data.date || new Date().toISOString().slice(0,10),
    data.inspector || '',
    data.type || '',
    data.result || '',
    data.notes || ''
  ];

  sheet.appendRow(row);
  return { success: true };
}

// ===== 修理記録 =====

const REPAIR_HEADERS = [
  '修理ID','機器ID','機器名','受付日','依頼部署','発生日','発生状況',
  '詳細内容','作業区分','状態','修理内容','見積もり依頼','修理点検依頼',
  '金額','交換部品1','部品単価1','交換部品2','部品単価2','交換部品3','部品単価3',
  '交換部品4','部品単価4','交換部品5','部品単価5','部品合計金額',
  '修理完了','修理完了日','修理完了時担当','ステータス'
];

function handleGetRepairs() {
  const sheet = getOrCreateSheet(SHEET_NAMES.repair);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, REPAIR_HEADERS.length).setValues([REPAIR_HEADERS]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.repair);
  return { success: true, data: result.rows };
}

function handleAddRepair(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.repair);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, REPAIR_HEADERS.length).setValues([REPAIR_HEADERS]);
  }

  const row = [
    data.repairId || ('REP-' + Date.now()),
    data.equipmentId || '',
    data.equipmentName || '',
    data.receptionDate || '',
    data.requestDept || '',
    data.occurrenceDate || '',
    data.occurrence || '',
    data.description || '',
    data.workCategory || '',
    data.condition || '',
    data.action || '',
    data.estimateRequest ? 'TRUE' : 'FALSE',
    data.inspectionRequest ? 'TRUE' : 'FALSE',
    data.cost || '',
    data.part1 || '', data.partPrice1 || '',
    data.part2 || '', data.partPrice2 || '',
    data.part3 || '', data.partPrice3 || '',
    data.part4 || '', data.partPrice4 || '',
    data.part5 || '', data.partPrice5 || '',
    data.partsTotal || '',
    data.repairCompleted ? 'TRUE' : 'FALSE',
    data.completedDate || '',
    data.completedStaff || '',
    data.repairStatus || '対応中'
  ];

  sheet.appendRow(row);
  return { success: true };
}

// ===== スタッフ =====

function handleGetStaff() {
  const sheet = getOrCreateSheet(SHEET_NAMES.staff);
  if (sheet.getLastRow() < 1) {
    const headers = ['社員ID','氏名','部署','メール','役割'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.staff);
  return { success: true, data: result.rows.map(r => ({
    staffId: r['社員ID'] || '',
    name: r['氏名'] || '',
    dept: r['部署'] || '',
    email: r['メール'] || '',
    role: r['役割'] || ''
  }))};
}

function handleSaveStaff(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.staff);
  if (sheet.getLastRow() < 1) {
    const headers = ['社員ID','氏名','部署','メール','役割'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // 既存行検索
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.staffId)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const row = [data.staffId, data.name, data.dept, data.email || '', data.role || ''];

  if (targetRow === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  }

  return { success: true };
}

function handleDeleteStaff(id) {
  return deleteRowById(SHEET_NAMES.staff, id);
}

// ===== 部署 =====

function handleGetDepartments() {
  const sheet = getOrCreateSheet(SHEET_NAMES.departments);
  if (sheet.getLastRow() < 1) {
    const headers = ['部署ID','部署名','説明'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.departments);
  return { success: true, data: result.rows.map(r => ({
    deptId: r['部署ID'] || '',
    name: r['部署名'] || '',
    description: r['説明'] || ''
  }))};
}

function handleSaveDepartment(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.departments);
  if (sheet.getLastRow() < 1) {
    const headers = ['部署ID','部署名','説明'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.deptId)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const row = [data.deptId, data.name, data.description || ''];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

  return { success: true };
}

function handleDeleteDepartment(id) {
  return deleteRowById(SHEET_NAMES.departments, id);
}

// ===== 貸出エリア =====

function handleGetAreas() {
  const sheet = getOrCreateSheet(SHEET_NAMES.areas);
  if (sheet.getLastRow() < 1) {
    const headers = ['エリアID','エリア名','アイコン','並び順'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.areas);
  return { success: true, data: result.rows.map(r => ({
    areaId: r['エリアID'] || '',
    name: r['エリア名'] || '',
    icon: r['アイコン'] || '',
    sortOrder: parseInt(r['並び順']) || 0
  }))};
}

function handleSaveArea(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.areas);
  if (sheet.getLastRow() < 1) {
    const headers = ['エリアID','エリア名','アイコン','並び順'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.areaId)) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const row = [data.areaId, data.name, data.icon || '', data.sortOrder || 0];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

  return { success: true };
}

function handleDeleteArea(id) {
  return deleteRowById(SHEET_NAMES.areas, id);
}

// ===== 物品マスター =====

const ITEM_HEADERS = [
  '物品ID','スキャン用','JANコード','区分','商品名','規格','製造会社','納入業者',
  '納入価格','単価','入数','定数','発注単位','発注点','使用方法及び説明','発注条件','交換パーツ','添付文書'
];

const ITEM_FIELD_MAP = {
  '物品ID': 'itemId', 'スキャン用': 'scanCode', 'JANコード': 'janCode',
  '区分': 'category', '商品名': 'productName', '規格': 'spec',
  '製造会社': 'manufacturer', '納入業者': 'supplier', '納入価格': 'deliveryPrice',
  '単価': 'unitPrice', '入数': 'packageQty', '定数': 'parLevel',
  '発注単位': 'orderUnit', '発注点': 'reorderPoint', '使用方法及び説明': 'description',
  '発注条件': 'orderCondition', '交換パーツ': 'isReplacementPart', '添付文書': 'packageInsert'
};

function handleGetItems() {
  const sheet = getOrCreateSheet(SHEET_NAMES.items);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.items);
  const converted = result.rows.map(row => {
    const obj = {};
    for (const [sheetCol, value] of Object.entries(row)) {
      const appField = ITEM_FIELD_MAP[sheetCol] || sheetCol;
      obj[appField] = value;
    }
    return obj;
  });
  return { success: true, data: converted };
}

function handleSaveItem(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.items);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  }
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.itemId)) { targetRow = i + 2; break; }
    }
  }
  const row = [
    data.itemId || '', data.scanCode || '', data.janCode || '',
    data.category || '', data.productName || '', data.spec || '',
    data.manufacturer || '', data.supplier || '',
    data.deliveryPrice || '', data.unitPrice || '',
    data.packageQty || '', data.parLevel || '',
    data.orderUnit || '', data.reorderPoint || '',
    data.description || '', data.orderCondition || '',
    data.isReplacementPart || '', data.packageInsert || ''
  ];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function handleDeleteItem(id) {
  return deleteRowById(SHEET_NAMES.items, id);
}

// ===== 全データ取得 =====

function handleGetAll() {
  return {
    success: true,
    equipment: handleGetEquipment(),
    staff: handleGetStaff(),
    departments: handleGetDepartments(),
    areas: handleGetAreas(),
    lending: handleGetLending(),
    items: handleGetItems()
  };
}

// ===== 全データ同期 =====

function handleSyncAll(data) {
  const results = {};

  if (data.equipment && data.equipment.length > 0) {
    results.equipment = handleSaveEquipmentBatch(data.equipment);
  }

  if (data.staff && data.staff.length > 0) {
    for (const s of data.staff) {
      handleSaveStaff(s);
    }
    results.staff = { success: true, count: data.staff.length };
  }

  if (data.departments && data.departments.length > 0) {
    for (const d of data.departments) {
      handleSaveDepartment(d);
    }
    results.departments = { success: true, count: data.departments.length };
  }

  if (data.areas && data.areas.length > 0) {
    for (const a of data.areas) {
      handleSaveArea(a);
    }
    results.areas = { success: true, count: data.areas.length };
  }

  if (data.items && data.items.length > 0) {
    for (const item of data.items) {
      handleSaveItem(item);
    }
    results.items = { success: true, count: data.items.length };
  }

  return { success: true, results };
}

// ===== 共通ヘルパー =====

function deleteRowById(sheetName, id) {
  const sheet = getOrCreateSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'データがありません' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { success: true, message: id + ' を削除しました' };
    }
  }

  return { success: false, error: id + ' が見つかりません' };
}

// ===== テスト用関数 =====

function testGetEquipment() {
  const result = handleGetEquipment();
  Logger.log(JSON.stringify(result).substring(0, 1000));
}

function testGetAll() {
  const result = handleGetAll();
  Logger.log(JSON.stringify(result).substring(0, 1000));
}
