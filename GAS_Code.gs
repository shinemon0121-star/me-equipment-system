/**
 * ME機器管理システム - Google Apps Script バックエンド
 * このスクリプトをGoogle Spreadsheetの「拡張機能 > Apps Script」に貼り付けてください。
 * デプロイ後のURLをアプリのGAS_URL設定に入力してください。
 *
 * スプレッドシートID: 1hORasvyl7H0KcVy33AhUfISv3-1HXatxduBXHuijspM
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1hORasvyl7H0KcVy33AhUfISv3-1HXatxduBXHuijspM';

const SHEET_NAMES = {
  equipment: '機器マスター',
  extendedSpecs: 'extended_specs',
  lending: '貸出履歴',
  inspection: '点検記録',
  repair: '修理記録',
  staff: 'スタッフ',
  departments: '部署',
  areas: '貸出エリア',
  items: '物品マスター',
  orders: '発注データ',
  inventoryLog: '入出庫履歴'
};

// 各種点検シート名の一覧（タイプ別点検データ取得用）
const TYPED_INSPECTION_SHEETS = [
  '輸液ポンプ点検', 'シリンジポンプ点検', '人工呼吸器点検', '除細動器点検',
  '医用テレメーター点検', 'ベッドサイドモニター点検', '麻酔器点検',
  '低圧持続吸引機点検', '低圧持続吸引器点検', 'IABP点検', 'ECMO_PCPS点検', '保育器点検', '電気メス点検', 'AED点検',
  'パルスオキシメータ点検', '経腸栄養ポンプ点検'
];

// ===== ヘルパー関数 =====
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

// シート名一覧を返す（デバッグ用）
function handleListSheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets().map(s => ({ name: s.getName(), gid: s.getSheetId(), rows: s.getLastRow() }));
  return { success: true, sheets };
}

// gid からシートを取得するヘルパー
function getSheetByGid(gid) {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getSheetId() === gid) return s;
  }
  return null;
}

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
      if (val === null || val === undefined) { row[headers[j]] = ''; continue; }
      // Date オブジェクトを yyyy/M/d 形式に変換
      if (val instanceof Date && !isNaN(val.getTime())) {
        row[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy/M/d');
      } else {
        const strVal = String(val || '').trim();
        // Date.toString() 形式（例: "Sat May 07 2022..."）を検出して yyyy/M/d に変換
        if (strVal.match(/^[A-Z][a-z]{2}\s+[A-Z][a-z]{2}\s+\d{2}\s+\d{4}/)) {
          try {
            const parsed = new Date(strVal);
            if (!isNaN(parsed.getTime())) {
              row[headers[j]] = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy/M/d');
            } else {
              row[headers[j]] = strVal;
            }
          } catch (e) {
            row[headers[j]] = strVal;
          }
        } else {
          row[headers[j]] = strVal;
        }
      }
      if (val !== '' && val !== null && val !== undefined) hasData = true;
    }
    if (hasData) rows.push(row);
  }
  return { headers, rows };
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h] = i + 1; });
  return { headers, map };
}

function getEquipmentFieldMap() {
  return {
    '機器ID': 'id', '機器区分': 'category', '機器名': 'name', '型式': 'model',
    '製造番号': 'serial', '管理部署': 'dept', '製造元': 'manufacturer', '代理店': 'dealer',
    'ステータス': 'status', '貸出状態': 'lending', '現在地': 'location',
    '保管場所': 'storage', '廃棄': 'discard',
    'CH': 'channel', '資産区分': 'assetType', '納入日': 'deliveryDate', '購入価格': 'purchasePrice',
    '代替レンタルデモ開始日': 'rentalStartDate', '代替レンタルデモ終了日': 'rentalEndDate',
    '保守契約': 'maintenanceContract', '廃棄日': 'discardDate', '廃棄理由': 'discardReason',
    '添付文書URL': 'documentUrl', '備考': 'notes',
    '管理部門': 'managedByDept', '保守契約形態': 'maintenanceType',
    '年間保守費': 'annualMaintenanceCost', '保守契約開始日': 'maintenanceStartDate', '保守契約終了日': 'maintenanceEndDate',
    '初回点検月': 'firstInspectionMonth', '耐久年数': 'durabilityYears',
    'スポット点検予定日': 'nextSpotInspDate', 'スポット点検備考': 'nextSpotInspNote',
    '取扱説明書URL': 'manualUrl',
  };
}

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
      case 'getEquipment': result = handleGetEquipment(); break;
      case 'getAll': result = handleGetAll(); break;
      case 'getHeaders': result = handleGetHeaders(); break;
      case 'getLending': result = handleGetLending(); break;
      case 'getInspections': result = handleGetInspections(); break;
      case 'getRepairs': result = handleGetRepairs(); break;
      case 'getStaff': result = handleGetStaff(); break;
      case 'getDepartments': result = handleGetDepartments(); break;
      case 'getAreas': result = handleGetAreas(); break;
      case 'getItems': result = handleGetItems(); break;
      case 'getOrders': result = handleGetOrders(); break;
      case 'getInventoryLog': result = handleGetInventoryLog(); break;
      case 'getExtendedSpecs': result = handleGetExtendedSpecs(); break;
      case 'getAllTypedInspections': result = handleGetAllTypedInspections(); break;
      case 'listSheets': result = handleListSheets(); break;
      default: result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== POST リクエスト =====
function doPost(e) {
  Logger.log('🔵 doPost called!');
  let result;
  try {
    Logger.log('POST data received, size: ' + (e.postData.contents ? e.postData.contents.length : 0));
    const body = JSON.parse(e.postData.contents);
    Logger.log('✅ JSON parsed successfully');
    Logger.log('Action: ' + body.action);
    if (body.equipment) Logger.log('Equipment array length: ' + body.equipment.length);
    const action = body.action;
    switch (action) {
      case 'saveEquipment': result = handleSaveEquipment(body.data); break;
      case 'saveEquipmentBatch': result = handleSaveEquipmentBatch(body.data); break;
      case 'deleteEquipment': result = handleDeleteEquipment(body.id); break;
      case 'addLending': result = handleAddLending(body.data); break;
      case 'addInspection': result = handleAddInspection(body.data); break;
      case 'addTypedInspection': result = handleAddTypedInspection(body.sheetName, body.data); break;
      case 'addRepair': result = handleAddRepair(body.data); break;
      case 'deleteRepair': result = handleDeleteRepair(body.repairId); break;
      case 'saveStaff': result = handleSaveStaff(body.data); break;
      case 'deleteStaff': result = handleDeleteStaff(body.id); break;
      case 'saveDepartment': result = handleSaveDepartment(body.data); break;
      case 'deleteDepartment': result = handleDeleteDepartment(body.id); break;
      case 'saveArea': result = handleSaveArea(body.data); break;
      case 'deleteArea': result = handleDeleteArea(body.id); break;
      case 'saveItem': result = handleSaveItem(body.data); break;
      case 'deleteItem': result = handleDeleteItem(body.id); break;
      case 'saveOrder': result = handleSaveOrder(body.data); break;
      case 'deleteOrder': result = handleDeleteOrder(body.id); break;
      case 'addInventoryLog': result = handleAddInventoryLog(body.data); break;
      case 'saveExtendedSpec': result = handleSaveExtendedSpec(body.data); break;
      case 'deleteExtendedSpec': result = handleDeleteExtendedSpec(body.id); break;
      case 'syncAll': result = handleSyncAll(body); break;
      case 'updateEquipmentField': result = handleUpdateEquipmentField(body.id, body.field, body.value); break;
      default: result = { success: false, error: 'Unknown action: ' + action };
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
    const defaultHeaders = ['機器ID','機器区分','機器名','型式','製造番号','管理部署','製造元','代理店','ステータス','貸出状態','現在地','保管場所','CH','資産区分','納入日','購入価格','代替レンタルデモ開始日','代替レンタルデモ終了日','保守契約','廃棄','廃棄日','廃棄理由','添付文書URL','備考','管理部門','保守契約形態','年間保守費','保守契約開始日','保守契約終了日','初回点検月','耐久年数','スポット点検予定日','スポット点検備考','取扱説明書URL'];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    return { success: true, data: [], headers: defaultHeaders };
  }
  const result = getSheetData(sheetName);
  const fieldMap = getEquipmentFieldMap();
  const converted = result.rows.map(row => {
    const obj = {};
    for (const [sheetCol, value] of Object.entries(row)) {
      const appField = fieldMap[sheetCol] || sheetCol;
      obj[appField] = value;
    }
    return obj;
  });
  return { success: true, data: converted, headers: result.headers, fieldMap: fieldMap };
}

function handleGetHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  if (sheet.getLastRow() < 1) return { success: true, headers: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { success: true, headers: headers };
}

function handleSaveEquipment(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);
  const reverseMap = getReverseFieldMap();

  // 必要な列を自動作成
  ensureEquipmentColumns(sheet, data);

  const { headers, map } = getHeaderMap(sheet);
  const idCol = map['機器ID'];
  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) { targetRow = i + 2; break; }
    }
  }
  if (targetRow === -1) targetRow = lastRow + 1;
  for (const [appField, value] of Object.entries(data)) {
    const sheetCol = reverseMap[appField] || appField;
    const col = map[sheetCol];
    if (col) sheet.getRange(targetRow, col).setValue(value || '');
  }
  return { success: true, message: '保存しました: ' + data.id };
}

// 機器シートに必要な列があなければ自動作成
function ensureEquipmentColumns(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const reverseMap = getReverseFieldMap();
  const requiredCols = new Set(Object.values(reverseMap));

  let newColIndex = sheet.getLastColumn() + 1;
  for (const [appField, sheetCol] of Object.entries(reverseMap)) {
    if (!headers.includes(sheetCol)) {
      // 新しい列を追加
      sheet.getRange(1, newColIndex).setValue(sheetCol);
      newColIndex++;
      Logger.log('✅ 列を追加: ' + sheetCol);
    }
  }
}

function handleSaveEquipmentBatch(dataArray) {
  const sheet = getOrCreateSheet(SHEET_NAMES.equipment);

  // 必要な列を自動作成
  if (dataArray.length > 0) {
    ensureEquipmentColumns(sheet, dataArray[0]);
  }

  const { headers, map } = getHeaderMap(sheet);
  const reverseMap = getReverseFieldMap();
  const idCol = map['機器ID'];
  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };
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
      if (colIdx) rowData[colIdx - 1] = value || '';
    }
    if (existingIds[data.id]) updates.push({ row: existingIds[data.id], data: rowData });
    else newRows.push(rowData);
  }
  for (const update of updates) {
    sheet.getRange(update.row, 1, 1, headers.length).setValues([update.data]);
  }
  if (newRows.length > 0) {
    sheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
  }
  return { success: true, message: '更新:' + updates.length + '件, 新規:' + newRows.length + '件' };
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
    data.equipmentId || '', data.equipmentName || '',
    data.action || '', data.destination || '',
    data.returnFrom || '', data.operator || ''
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
  return { success: true, data: getSheetData(SHEET_NAMES.inspection).rows };
}

function handleAddInspection(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.inspection);
  if (sheet.getLastRow() < 1) {
    const headers = ['点検ID','機器ID','機器名','点検日','点検者','点検種別','結果','備考'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.appendRow([
    data.inspectionId || ('INS-' + Date.now()),
    data.equipmentId || '', data.equipmentName || '',
    data.date || new Date().toISOString().slice(0,10),
    data.inspector || '', data.type || '',
    data.result || '', data.notes || ''
  ]);
  return { success: true };
}

function handleAddTypedInspection(sheetName, data) {
  if (!sheetName || !data) return { success: false, error: 'Missing sheetName or data' };
  const ss = getSpreadsheet();
  // まず名前で検索
  let sheet = ss.getSheetByName(sheetName);
  // 名前で見つからなければ gid で検索
  if (!sheet) {
    const gidEntry = TYPED_INSPECTION_GIDS.find(e => e.name === sheetName);
    if (gidEntry) {
      sheet = getSheetByGid(gidEntry.gid);
    }
  }
  // それでもなければ新規作成
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const keys = Object.keys(data);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Add missing headers
  keys.forEach(k => {
    if (!headers.includes(k)) {
      headers.push(k);
      sheet.getRange(1, headers.length).setValue(k);
    }
  });
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return { success: true, sheetUsed: sheet.getName() };
}

// ===== 修理記録 =====
const REPAIR_HEADERS = [
  '修理ID','機器ID','機器名','受付日','依頼部署','発生日','発生状況',
  '詳細内容','作業区分','状態','修理内容','見積もり依頼','修理点検依頼',
  '金額','交換部品1','部品単価1','交換部品2','部品単価2','交換部品3','部品単価3',
  '交換部品4','部品単価4','交換部品5','部品単価5','部品合計金額',
  '修理完了','修理完了日','修理完了時担当','ステータス',
  '概算費用','ダウンタイム時間','故障区分','臨床的重要度'
];

function handleGetRepairs() {
  const ss = getSpreadsheet();
  // まず名前で検索
  let sheet = ss.getSheetByName(SHEET_NAMES.repair);
  // 名前で見つからなければ gid=895124206 で検索
  if (!sheet) {
    sheet = getSheetByGid(895124206);
    // 見つかったら SHEET_NAMES を実際のタブ名に更新（以降のアクセス用）
    if (sheet) SHEET_NAMES.repair = sheet.getName();
  }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.repair);
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, REPAIR_HEADERS.length).setValues([REPAIR_HEADERS]);
    return { success: true, data: [] };
  }

  // 必要なカラムを確認・追加
  ensureRepairColumns(sheet, {});
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      const val = data[i][j];
      if (val === null || val === undefined) { row[headers[j]] = ''; continue; }
      if (val instanceof Date && !isNaN(val.getTime())) {
        row[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy/M/d');
      } else {
        row[headers[j]] = String(val);
      }
      if (val !== '' && val !== null && val !== undefined) hasData = true;
    }
    if (hasData) {
      // 修理IDがない場合は自動生成
      if (!row['修理ID'] || row['修理ID'].trim() === '') {
        const equipmentId = row['機器ID'] || '';
        const dateStr = row['受付日'] || '';
        row['修理ID'] = 'REP-' + equipmentId + '-' + dateStr.replace(/\//g, '') + '-' + i;
        Logger.log('修理ID自動生成: ' + row['修理ID'] + ' (機器ID: ' + equipmentId + ')');
      } else {
        Logger.log('修理ID既存: ' + row['修理ID']);
      }
      rows.push(row);
    }
  }
  return { success: true, data: rows, sheetName: sheet.getName() };
}

function handleAddRepair(data) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.repair);
  if (!sheet) { sheet = getSheetByGid(895124206); if (sheet) SHEET_NAMES.repair = sheet.getName(); }
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.repair);

  // ヘッダー行を確認・修正
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, REPAIR_HEADERS.length).setValues([REPAIR_HEADERS]);
  } else {
    // 既存シートの場合、必要なカラムを確認・追加
    ensureRepairColumns(sheet, data);
  }

  const row = [
    data.repairId || ('REP-' + Date.now()),
    data.equipmentId || '', data.equipmentName || '',
    data.receptionDate || '', data.requestDept || '',
    data.occurrenceDate || '', data.occurrence || '',
    data.description || '', data.workCategory || '',
    data.condition || '', data.action || '',
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
    data.completedDate || '', data.completedStaff || '',
    data.repairStatus || '対応中',
    data.estimatedCost || '', data.downtimeHours || '',
    data.failureType || '', data.clinicalImportance || ''
  ];
  // 既存レコードがあれば更新、なければ追加
  const lastRow = sheet.getLastRow();
  if (lastRow > 1 && data.repairId) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.repairId)) {
        sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return { success: true };
      }
    }
  }
  sheet.appendRow(row);
  return { success: true };
}

// 修理シートに必要なカラムがない場合、自動的に追加
function ensureRepairColumns(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const requiredHeaders = REPAIR_HEADERS;
  let columnsAdded = false;

  for (let i = 0; i < requiredHeaders.length; i++) {
    if (!headers.includes(requiredHeaders[i])) {
      // カラムが足りない場合は末尾に追加
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(requiredHeaders[i]);
      columnsAdded = true;
      Logger.log('修理シートにカラムを追加: ' + requiredHeaders[i]);
    }
  }

  if (columnsAdded) {
    Logger.log('修理シートにカラムを追加しました');
  }
}

function handleDeleteRepair(repairId) {
  Logger.log('🔴 handleDeleteRepair called with repairId: ' + repairId);

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.repair);
  if (!sheet) {
    sheet = getSheetByGid(895124206);
    if (sheet) SHEET_NAMES.repair = sheet.getName();
  }
  if (!sheet) return { success: false, error: '修理シートが見つかりません' };

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'データがありません' };

  // 全データを取得して検索
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  Logger.log('シートの総行数: ' + lastRow + ', ヘッダ: ' + JSON.stringify(headers.slice(0, 3)));

  // 複数のパターンで検索
  for (let i = data.length - 1; i >= 1; i--) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = String(data[i][j] || '');
    }

    // パターン1: 修理IDが完全一致
    if (row['修理ID'] === String(repairId)) {
      Logger.log('✅ パターン1で一致。行' + i + 'を削除');
      sheet.deleteRow(i + 1);
      return { success: true, message: repairId + ' を削除しました（パターン1）' };
    }

    // パターン2: 機器ID + 受付日 で一致（自動生成IDの場合）
    const genId = `REP-${row['機器ID'] || ''}-${(row['受付日'] || '').replace(/\//g, '')}`;
    if (genId === String(repairId)) {
      Logger.log('✅ パターン2で一致。生成ID: ' + genId + ', 行' + i + 'を削除');
      sheet.deleteRow(i + 1);
      return { success: true, message: repairId + ' を削除しました（パターン2）' };
    }

    // デバッグ: 最初の3行だけログに出力
    if (i <= 3) {
      Logger.log('行' + i + ': 修理ID=' + row['修理ID'] + ', 生成ID=' + genId);
    }
  }

  Logger.log('❌ ' + repairId + ' が見つかりません');
  return { success: false, error: repairId + ' が見つかりません' };
}

// ===== スタッフ =====
function handleGetStaff() {
  const sheet = getOrCreateSheet(SHEET_NAMES.staff);
  if (sheet.getLastRow() < 1) {
    const headers = ['社員ID','氏名','部署','メール','役割'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { success: true, data: [] };
  }
  return { success: true, data: getSheetData(SHEET_NAMES.staff).rows.map(r => ({
    staffId: r['社員ID']||'', name: r['氏名']||'',
    dept: r['部署']||'', email: r['メール']||'', role: r['役割']||''
  }))};
}

function handleSaveStaff(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.staff);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 5).setValues([['社員ID','氏名','部署','メール','役割']]);
  }
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.staffId)) { targetRow = i + 2; break; }
    }
  }
  const row = [data.staffId, data.name, data.dept, data.email||'', data.role||''];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function handleDeleteStaff(id) { return deleteRowById(SHEET_NAMES.staff, id); }

// ===== 部署 =====
function handleGetDepartments() {
  const sheet = getOrCreateSheet(SHEET_NAMES.departments);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['部署ID','部署名','説明']]);
    return { success: true, data: [] };
  }
  return { success: true, data: getSheetData(SHEET_NAMES.departments).rows.map(r => ({
    deptId: r['部署ID']||'', name: r['部署名']||'', description: r['説明']||''
  }))};
}

function handleSaveDepartment(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.departments);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['部署ID','部署名','説明']]);
  }
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.deptId)) { targetRow = i + 2; break; }
    }
  }
  const row = [data.deptId, data.name, data.description||''];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function handleDeleteDepartment(id) { return deleteRowById(SHEET_NAMES.departments, id); }

// ===== 貸出エリア =====
function handleGetAreas() {
  const sheet = getOrCreateSheet(SHEET_NAMES.areas);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 4).setValues([['エリアID','エリア名','アイコン','並び順']]);
    return { success: true, data: [] };
  }
  return { success: true, data: getSheetData(SHEET_NAMES.areas).rows.map(r => ({
    areaId: r['エリアID']||'', name: r['エリア名']||'',
    icon: r['アイコン']||'', sortOrder: parseInt(r['並び順'])||0
  }))};
}

function handleSaveArea(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.areas);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 4).setValues([['エリアID','エリア名','アイコン','並び順']]);
  }
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.areaId)) { targetRow = i + 2; break; }
    }
  }
  const row = [data.areaId, data.name, data.icon||'', data.sortOrder||0];
  if (targetRow === -1) sheet.appendRow(row);
  else sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function handleDeleteArea(id) { return deleteRowById(SHEET_NAMES.areas, id); }

// ===== 物品マスター =====
const ITEM_HEADERS = [
  '物品ID','スキャン用','JANコード','区分','商品名','規格','製造会社','納入業者',
  '納入価格','単価','入数','定数','発注単位','発注点','使用方法及び説明','発注条件','交換パーツ','添付文書'
];
const ITEM_FIELD_MAP = {
  '物品ID':'itemId','スキャン用':'scanCode','JANコード':'janCode','区分':'category',
  '商品名':'productName','規格':'spec','製造会社':'manufacturer','納入業者':'supplier',
  '納入価格':'deliveryPrice','単価':'unitPrice','入数':'packageQty','定数':'parLevel',
  '発注単位':'orderUnit','発注点':'reorderPoint','使用方法及び説明':'description',
  '発注条件':'orderCondition','交換パーツ':'isReplacementPart','添付文書':'packageInsert'
};
function handleGetItems() {
  const sheet = getOrCreateSheet(SHEET_NAMES.items);
  if (sheet.getLastRow() < 1) { sheet.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]); return {success:true,data:[]}; }
  const result = getSheetData(SHEET_NAMES.items);
  return {success:true, data: result.rows.map(row => {
    const obj = {};
    for (const [k,v] of Object.entries(row)) { obj[ITEM_FIELD_MAP[k]||k] = v; }
    return obj;
  })};
}
function handleSaveItem(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.items);
  if (sheet.getLastRow() < 1) sheet.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2,1,lastRow-1,1).getValues();
    for (let i=0;i<ids.length;i++) { if (String(ids[i][0])===String(data.itemId)) { targetRow=i+2; break; } }
  }
  const row = [data.itemId||'',data.scanCode||'',data.janCode||'',data.category||'',data.productName||'',
    data.spec||'',data.manufacturer||'',data.supplier||'',data.deliveryPrice||'',data.unitPrice||'',
    data.packageQty||'',data.parLevel||'',data.orderUnit||'',data.reorderPoint||'',
    data.description||'',data.orderCondition||'',data.isReplacementPart||'',data.packageInsert||''];
  if (targetRow===-1) sheet.appendRow(row);
  else sheet.getRange(targetRow,1,1,row.length).setValues([row]);
  return {success:true};
}
function handleDeleteItem(id) { return deleteRowById(SHEET_NAMES.items, id); }

// ===== 発注データ =====
const ORDER_HEADERS = ['発注ID','発注日','スキャン用','JANコード','発注数量','ステータス','備考'];
const ORDER_FIELD_MAP = {'発注ID':'orderId','発注日':'date','スキャン用':'scanCode','JANコード':'janCode','発注数量':'quantity','ステータス':'status','備考':'remarks'};
function handleGetOrders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.orders);
  if (sheet.getLastRow() < 1) { sheet.getRange(1,1,1,ORDER_HEADERS.length).setValues([ORDER_HEADERS]); return {success:true,data:[]}; }
  const result = getSheetData(SHEET_NAMES.orders);
  return {success:true, data: result.rows.map(row => {
    const obj = {};
    for (const [k,v] of Object.entries(row)) { obj[ORDER_FIELD_MAP[k]||k] = v; }
    return obj;
  })};
}
function handleSaveOrder(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.orders);
  if (sheet.getLastRow() < 1) sheet.getRange(1,1,1,ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2,1,lastRow-1,1).getValues();
    for (let i=0;i<ids.length;i++) { if (String(ids[i][0])===String(data.orderId)) { targetRow=i+2; break; } }
  }
  const row = [data.orderId||'',data.date||'',data.scanCode||'',data.janCode||'',data.quantity||'',data.status||'',data.remarks||''];
  if (targetRow===-1) sheet.appendRow(row);
  else sheet.getRange(targetRow,1,1,row.length).setValues([row]);
  return {success:true};
}
function handleDeleteOrder(id) { return deleteRowById(SHEET_NAMES.orders, id); }

// ===== 入出庫履歴 =====
const INVLOG_HEADERS = ['記録ID','日時','物品ID','物品名','JANコード','カテゴリ','種別','数量','担当者','備考'];
const INVLOG_FIELD_MAP = {'記録ID':'logId','日時':'timestamp','物品ID':'itemId','物品名':'itemName','JANコード':'janCode','カテゴリ':'category','種別':'type','数量':'quantity','担当者':'staff','備考':'note'};
const INVLOG_REVERSE_MAP = {};
for (const [k,v] of Object.entries(INVLOG_FIELD_MAP)) INVLOG_REVERSE_MAP[v] = k;

function handleGetInventoryLog() {
  const sheet = getOrCreateSheet(SHEET_NAMES.inventoryLog);
  if (sheet.getLastRow() < 1) { sheet.getRange(1,1,1,INVLOG_HEADERS.length).setValues([INVLOG_HEADERS]); return {success:true,data:[]}; }
  const result = getSheetData(SHEET_NAMES.inventoryLog);
  return {success:true, data: result.rows.map(row => {
    const obj = {};
    for (const [k,v] of Object.entries(row)) { obj[INVLOG_FIELD_MAP[k]||k] = v; }
    return obj;
  })};
}
function handleAddInventoryLog(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.inventoryLog);
  if (sheet.getLastRow() < 1) sheet.getRange(1,1,1,INVLOG_HEADERS.length).setValues([INVLOG_HEADERS]);
  sheet.appendRow(INVLOG_HEADERS.map(h => data[INVLOG_FIELD_MAP[h]||h] || data[h] || ''));
  return {success:true};
}

// ===== 各種点検データ一括取得 =====
// 各種点検シートの gid 一覧（名前で見つからない場合のフォールバック）
const TYPED_INSPECTION_GIDS = [
  { name: '輸液ポンプ点検', gid: 305816429 },
  { name: 'シリンジポンプ点検', gid: 1915372484 },
  { name: '人工呼吸器点検', gid: 200938452 },
  { name: '除細動器点検', gid: 1434266603 },
  { name: '医用テレメーター点検', gid: 683083212 },
  { name: 'ベッドサイドモニター点検', gid: 929301363 },
  { name: '麻酔器点検', gid: 1657040241 },
  { name: '低圧持続吸引機点検', gid: 1769894837 },
  { name: 'IABP点検', gid: 1723261589 },
  { name: 'ECMO_PCPS点検', gid: 1141402612 }
];

function handleGetAllTypedInspections() {
  const allData = [];
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();

  // 名前で検索 → gid で検索の順でシートを探す
  const sheetsToRead = [];
  // まず名前ベースで探す
  for (const sheetName of TYPED_INSPECTION_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheetsToRead.push({ sheet, label: sheetName });
    }
  }
  // 名前で見つからなかったものを gid で補完
  for (const entry of TYPED_INSPECTION_GIDS) {
    const alreadyFound = sheetsToRead.some(s => s.label === entry.name);
    if (!alreadyFound) {
      const sheet = allSheets.find(s => s.getSheetId() === entry.gid);
      if (sheet && sheet.getLastRow() > 1) {
        sheetsToRead.push({ sheet, label: entry.name + '(' + sheet.getName() + ')' });
      }
    }
  }

  for (const { sheet, label } of sheetsToRead) {
    try {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      for (let i = 1; i < data.length; i++) {
        const row = {};
        let hasData = false;
        for (let j = 0; j < headers.length; j++) {
          const val = data[i][j];
          if (val === null || val === undefined) { row[headers[j]] = ''; continue; }
          if (val instanceof Date && !isNaN(val.getTime())) {
            row[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy/M/d');
          } else {
            row[headers[j]] = String(val);
          }
          if (val !== '' && val !== null && val !== undefined) hasData = true;
        }
        if (!hasData) continue;
        row._sheetName = label.split('(')[0]; // 正規化された名前
        row._actualSheetName = sheet.getName();
        allData.push(row);
      }
    } catch(e) { /* シートが存在しない場合はスキップ */ }
  }
  return { success: true, data: allData, sheetCount: sheetsToRead.length };
}

// ===== 全データ取得・同期 =====
function handleGetAll() {
  return {
    success: true,
    equipment: handleGetEquipment(), staff: handleGetStaff(),
    departments: handleGetDepartments(), areas: handleGetAreas(),
    lending: handleGetLending(), items: handleGetItems(),
    orders: handleGetOrders(),
    inspections: handleGetInspections(),
    repairs: handleGetRepairs(),
    inventoryLog: handleGetInventoryLog(),
    typedInspections: handleGetAllTypedInspections(),
    extendedSpecs: handleGetExtendedSpecs()
  };
}

function handleSyncAll(data) {
  try {
    if (!data) return { success: false, error: 'データが送信されていません (data is undefined)' };
    const results = {};
    if (data.equipment && Array.isArray(data.equipment) && data.equipment.length > 0) {
      results.equipment = handleSaveEquipmentBatch(data.equipment);
    }
    if (data.staff && Array.isArray(data.staff) && data.staff.length > 0) {
      for (const s of data.staff) handleSaveStaff(s);
      results.staff = { success: true, count: data.staff.length };
    }
    if (data.departments && Array.isArray(data.departments) && data.departments.length > 0) {
      for (const d of data.departments) handleSaveDepartment(d);
      results.departments = { success: true, count: data.departments.length };
    }
    if (data.areas && Array.isArray(data.areas) && data.areas.length > 0) {
      for (const a of data.areas) handleSaveArea(a);
      results.areas = { success: true, count: data.areas.length };
    }
    if (data.items && Array.isArray(data.items) && data.items.length > 0) {
      for (const it of data.items) handleSaveItem(it);
      results.items = { success: true, count: data.items.length };
    }
    if (data.orders && Array.isArray(data.orders) && data.orders.length > 0) {
      for (const o of data.orders) handleSaveOrder(o);
      results.orders = { success: true, count: data.orders.length };
    }
    return { success: true, results };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
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

// ===== extended_specs =====
const EXT_SPEC_HEADERS = ['機器ID','管球型番','管球交換日','管球累積照射回数','管球交換費用','線量管理対象','高額保守契約詳細','保守対応時間','年間稼働時間目標','薬事分類','特記事項'];
const EXT_SPEC_FIELD_MAP = {
  '機器ID':'equipmentId','管球型番':'tubeModel','管球交換日':'tubeReplacedDate',
  '管球累積照射回数':'tubeShotCount','管球交換費用':'tubeReplaceCost','線量管理対象':'isDoseManaged',
  '高額保守契約詳細':'maintenanceDetail','保守対応時間':'slaResponseHours',
  '年間稼働時間目標':'annualTargetHours','薬事分類':'regulatoryClass','特記事項':'specialNotes'
};
const EXT_SPEC_REVERSE_MAP = {};
for (const [k,v] of Object.entries(EXT_SPEC_FIELD_MAP)) EXT_SPEC_REVERSE_MAP[v] = k;

function handleGetExtendedSpecs() {
  const sheet = getOrCreateSheet(SHEET_NAMES.extendedSpecs);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1,1,1,EXT_SPEC_HEADERS.length).setValues([EXT_SPEC_HEADERS]);
    return { success: true, data: [] };
  }
  const result = getSheetData(SHEET_NAMES.extendedSpecs);
  return { success: true, data: result.rows.map(row => {
    const obj = {};
    for (const [k,v] of Object.entries(row)) { obj[EXT_SPEC_FIELD_MAP[k]||k] = v; }
    return obj;
  })};
}

function handleSaveExtendedSpec(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.extendedSpecs);
  if (sheet.getLastRow() < 1) sheet.getRange(1,1,1,EXT_SPEC_HEADERS.length).setValues([EXT_SPEC_HEADERS]);
  const { headers, map } = getHeaderMap(sheet);
  const idCol = map['機器ID'];
  if (!idCol) return { success: false, error: '機器ID列が見つかりません' };
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.equipmentId)) { targetRow = i + 2; break; }
    }
  }
  if (targetRow === -1) targetRow = lastRow + 1;
  for (const [appField, value] of Object.entries(data)) {
    const sheetCol = EXT_SPEC_REVERSE_MAP[appField] || appField;
    const col = map[sheetCol];
    if (col) sheet.getRange(targetRow, col).setValue(value || '');
  }
  return { success: true, message: '拡張仕様を保存しました: ' + data.equipmentId };
}

function handleDeleteExtendedSpec(id) {
  return deleteRowById(SHEET_NAMES.extendedSpecs, id);
}