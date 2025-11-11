// -----------------------------------------------
// グローバル変数 & 設定
// -----------------------------------------------

// (v21: スクリプトプロパティからIDを取得)
// (v23: データ用とマスタ用を分離)
var DATA_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
var MASTER_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');


// シート名 (v20: 日報機能を削除 / v21: マスタ追加)
var SHEET_NAMES = {
  SETTINGS: '設定',
  EMPLOYEES: '社員マスタ',
  EXPENSE_ITEMS: '経費項目マスタ',
  REPORTS: '経費ヘッダー', 
  EXPENSES: '経費明細',
  
  // (v21: 新規追加マスタ)
  TRAVEL_RULES: '出張旅費規定マスタ',
  AREA_MASTER: '宿泊費エリアマスタ',
  POSITION_MASTER: '役職マスタ',
  DEPARTMENT_MASTER: '部門マスタ',
  JOB_TYPE_MASTER: '職種区分マスタ'
};

// -----------------------------------------------
// メイン & 初期化
// -----------------------------------------------

/**
 * WebアプリのGETリクエストハンドラ (v20: 修正 / v21: マスタ読み込み拡張 / v23: マスタ分離対応)
 */
function doGet(e) {
  var ui = HtmlService.createTemplateFromFile('index.html');
  // (v23: ssの即時オープンを削除)
  
  // 1. ユーザー情報 (v21: 拡張)
  var userInfo = getUserInfo_();
  
  // 2. 経費項目マスタ (v23: ss引数を削除)
  var expenseItems = getSheetData_(SHEET_NAMES.EXPENSE_ITEMS);
  
  // (v21: 新規マスタ読み込み / v23: ss引数を削除)
  var travelRules = getSheetData_(SHEET_NAMES.TRAVEL_RULES);
  var areaMaster = getSheetData_(SHEET_NAMES.AREA_MASTER);
  var positionMaster = getSheetData_(SHEET_NAMES.POSITION_MASTER);
  var departmentMaster = getSheetData_(SHEET_NAMES.DEPARTMENT_MASTER);
  var jobTypeMaster = getSheetData_(SHEET_NAMES.JOB_TYPE_MASTER);
  
  // 3. テンプレートにデータを渡す
  ui.employeeInfo = JSON.stringify(userInfo.employeeInfo); // ユーザー情報 (v21: 拡張された情報を含む)
  ui.isManager = userInfo.isManager; // 管理者フラグ
  ui.expenseItems = JSON.stringify(expenseItems); // 経費項目
  
  // (v21: 新規マスタを渡す)
  ui.travelRules = JSON.stringify(travelRules);
  ui.areaMaster = JSON.stringify(areaMaster);
  ui.positionMaster = JSON.stringify(positionMaster);
  ui.departmentMaster = JSON.stringify(departmentMaster);
  ui.jobTypeMaster = JSON.stringify(jobTypeMaster);
  
  // 4. HTMLを生成
  return ui.evaluate()
    .setTitle('経費精算アプリ') // (v20: タイトル変更)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ログインユーザーの情報を取得 (v2 / v21: 取得項目拡張 / v23: マスタ分離対応)
 * @returns {object} { employeeInfo: { id, name, positionCode, deptCode, jobTypeCode ... }, isManager: boolean }
 */
function getUserInfo_() {
  var email = Session.getActiveUser().getEmail();
  
  // (v23: ssの即時オープンを削除)
  var data = getSheetData_(SHEET_NAMES.EMPLOYEES); // (v23: ss引数を削除)
  
  var userInfo = {
    employeeInfo: {
      id: null,
      name: 'ゲスト',
      email: email,
      positionCode: null, // (v21: 役職コード)
      deptCode: null, // (v21: 所属部門コード)
      jobTypeCode: null // (v21: 職種区分コード)
    },
    isManager: false
  };

  for (var i = 0; i < data.length; i++) {
    // (v21: キーを要件定義の「アドレス」に変更)
    if (data[i]['アドレス'] === email) {
      userInfo.employeeInfo.id = data[i]['社員番号'];
      userInfo.employeeInfo.name = data[i]['社員名'];
      
      // (v21: 拡張項目を取得)
      userInfo.employeeInfo.positionCode = data[i]['役職 (コード)'];
      userInfo.employeeInfo.deptCode = data[i]['所属部門 (コード)'];
      userInfo.employeeInfo.jobTypeCode = data[i]['職種区分 (コード)'];

      // (v21: 役職名での判定を継続)
      var positionName = data[i]['役職名']; 
      if (positionName && (positionName.indexOf('部長') > -1 || positionName.indexOf('課長') > -1)) {
        userInfo.isManager = true;
      }
      break;
    }
  }
  return userInfo;
}

/**
 * 設定シートをK-Vオブジェクトで取得 (キャッシュ対応) (v9 修正 / v23: データID使用)
 * @returns {object} 設定の K-V
 */
function getSettings_() {
  var cache = CacheService.getScriptCache();
  var CACHE_KEY = 'SETTINGS_CACHE';
  
  var cached = cache.get(CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }

  var ss = SpreadsheetApp.openById(DATA_SPREADSHEET_ID); // (v23: データIDを使用)
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }
  if (sheet.getLastColumn() < 2) {
    return {}; 
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  
  var settings = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) { 
      settings[data[i][0]] = data[i][1];
    }
  }
  
  cache.put(CACHE_KEY, JSON.stringify(settings), 21600);
  return settings;
}


// -----------------------------------------------
// (A) 経費精算入力 (v20: 大幅修正 / v21: 拡張)
// -----------------------------------------------

/**
 * 経費精算の保存または申請 (v20: 日報機能を削除 / v21: areaCode追加, 日当重複チェック / v23: データID使用)
 * @param {object} data { header, expenses }
 * @param {string} status "一時保存" / "申請中"
 * @returns {object} { status: "success", message: "..." }
 */
function saveOrSubmitExpenseReport(data, status) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); 
  
  var ss = SpreadsheetApp.openById(DATA_SPREADSHEET_ID); // (v23: データIDを使用)
  var userInfo = getUserInfo_();
  
  try {
    var header = data.header;
    var expenses = data.expenses;
    var reportId = header.reportId;

    if (!reportId) {
      reportId = 'KEIHI-' + Utilities.getUuid(); 
    }
    
    deleteReportData_(reportId, ss); 
    
    var totalExpenseAmount = 0;
    var allowanceCount = 0; // (v21: 日当(101)のカウント用)

    for (var k = 0; k < expenses.length; k++) {
      totalExpenseAmount += parseFloat(expenses[k].amount || 0);
      
      // (v21: 要件2-4 日当(101)の重複チェック)
      if (String(expenses[k].itemCode) === '101') {
        allowanceCount++;
      }
    }

    // (v21: 要件2-4 日当(101)が2件以上あればエラー)
    if (allowanceCount > 1) {
      throw new Error('日当(101)が複数登録されています。明細を確認してください。');
    }

    var headerRow = [
      reportId, // 経費ID (A)
      header.date, // 日付 (B) (yyyy-MM-dd)
      userInfo.employeeInfo.name, // 申請者 (C)
      header.title, // (v20 新規) 件名 (D)
      header.remarks, // 備考 (E)
      status, // 承認ステータス (F)
      userInfo.employeeInfo.id, // 申請者社員番号 (G)
      totalExpenseAmount, // 合計経費 (H)
      (status === '申請中') ? new Date() : null, // 申請日 (I)
      null // 差戻し理由 (J)
    ];

    var expensesRows = expenses.map(function(exp) {
      var expenseId = exp.expenseId;
      if (!expenseId || expenseId.startsWith('temp_')) {
          expenseId = 'EXP-' + Utilities.getUuid();
      }
      return [
        reportId, // 経費ID (A)
        expenseId, // 明細ID (B)
        exp.itemCode, // 経費項目コード (C)
        exp.amount, // 金額 (D)
        exp.receiptUrl, // 領収書URL (E)
        exp.remarks, // 備考 (F)
        exp.useDate, // 利用日 (G) (yyyy-MM-dd)
        exp.areaCode || null // (v21: 要件2-3 エリアコード (H))
      ];
    });

    if (headerRow.length > 0) {
      ss.getSheetByName(SHEET_NAMES.REPORTS).appendRow(headerRow);
    }
    if (expensesRows.length > 0) {
      var expSheet = ss.getSheetByName(SHEET_NAMES.EXPENSES);
      expSheet.getRange(expSheet.getLastRow() + 1, 1, expensesRows.length, expensesRows[0].length)
        .setValues(expensesRows);
    }
    
    if (status === '申請中') {
      var settings = getSettings_();
      var approverEmail = settings.APPROVER_EMAIL_1;
      if (approverEmail) {
        MailApp.sendEmail({
          to: approverEmail,
          subject: '[経費精算] 承認依頼: ' + userInfo.employeeInfo.name + ' (' + header.date + ')',
          body: userInfo.employeeInfo.name + ' さんから経費精算の承認依頼が届きました。\n\n' +
                '日付: ' + header.date + '\n' +
                '件名: ' + header.title + '\n' +
                '合計経費: ' + totalExpenseAmount + ' 円\n' +
                '備考: ' + header.remarks + '\n\n' +
                'システムにログインして内容を確認・承認してください。\n' +
                ScriptApp.getService().getUrl()
        });
      }
    }

    return { status: "success", message: status + "が完了しました。", reportId: reportId };
  } catch (e) {
    Logger.log('saveOrSubmitExpenseReport エラー: ' + e);
    // (v21: エラーメッセージを具体的に返す)
    return { status: "error", message: "保存/申請中にエラーが発生しました: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * (v22: 削除) AI-OCRによる画像処理 (processImageOCR)
 */

/**
 * (v22: 新規) Base64画像をDriveにアップロードし、URLを返す
 * @param {string} base64Image
 * @returns {object} { status, url, message }
*/
function uploadBase64Image(base64Image) {
  try {
    var settings = getSettings_();
    var folderId = settings.RECEIPT_FOLDER_ID;

    if (!folderId) {
      throw new Error('設定シート (RECEIPT_FOLDER_ID) を確認してください。');
    }

    // 1. 画像デコードとDrive保存
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Image.split(',')[1]), 
      'image/png', // PNG形式で固定
      'receipt-' + new Date().getTime() + '.png'
    );
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);
    
    // ★ 修正箇所: 組織外への共有 (ANYONE_WITH_LINK) がポリシーで拒否される場合、
    // 組織内 (DOMAIN_WITH_LINK) での共有に変更します。
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = file.getUrl();
    
    return { 
      status: "success", 
      url: fileUrl 
    };

  } catch (e) {
    Logger.log('uploadBase64Image エラー: ' + e);
    return { status: "error", message: "画像アップロードエラー: " + e.message };
  }
}


// -----------------------------------------------
// (B) (C) 上長・閲覧用 (v20: 修正あり / v21: 拡張)
// -----------------------------------------------

/**
 * (B) 承認待ちの申請一覧を取得 (v20: 修正 / v23: マスタ分離対応)
 */
function getPendingReports() {
  // (v23: ssの即時オープンを削除)
  var data = getSheetData_(SHEET_NAMES.REPORTS); // (v23: ss引数を削除)
  
  var pendingReports = data.filter(function(row) {
    return row['承認ステータス'] === '申請中';
  });
  return pendingReports;
}

/**
 * (C) 特定の申請の詳細データを取得 (v20: 修正 / v21: areaCode取得 / v23: マスタ分離対応)
 */
function getReportDetails(reportId) {
  try {
    if (!reportId) {
      throw new Error('経費IDが指定されていません。');
    }
    
    // (v23: ssの即時オープンを削除)
    
    var headerData = getSheetData_(SHEET_NAMES.REPORTS); // (v23: ss引数を削除)
    var header = headerData.find(function(row) {
      return String(row['経費ID']) === String(reportId); 
    });
    
    var expensesData = getSheetData_(SHEET_NAMES.EXPENSES); // (v23: ss引数を削除)
    var expenses = expensesData.filter(function(row) {
      return String(row['経費ID']) === String(reportId); 
    });

    var expenseItemsMaster = getSheetData_(SHEET_NAMES.EXPENSE_ITEMS); // (v23: ss引数を削除)
    var expenseItemsMap = {};
    expenseItemsMaster.forEach(function(item) {
      expenseItemsMap[item['経費項目コード']] = item['経費項目'];
    });

    // (v21: 要件3-5 エリアマスタを取得してマップを作成)
    var areaMaster = getSheetData_(SHEET_NAMES.AREA_MASTER); // (v23: ss引数を削除)
    var areaMap = {};
    areaMaster.forEach(function(item) {
      areaMap[item['地域区分 (コード)']] = item['地域区分名'];
    });
    
    expenses = expenses.map(function(exp) {
      // (v21: 要件2-3 エリアコードを取得)
      var areaCode = exp['エリアコード']; 
      
      return {
        expenseId: exp['明細ID'],
        itemCode: exp['経費項目コード'],
        amount: exp['金額'],
        receiptUrl: exp['領収書URL'],
        remarks: exp['備考'],
        useDate: exp['利用日'],
        itemName: expenseItemsMap[exp['経費項目コード']] || exp['経費項目コード'],
        areaCode: areaCode, // (v21: 取得したエリアコード)
        areaName: areaMap[areaCode] || '' // (v21: 要件3-5 エリア名に変換)
      };
    });

    return {
      header: header || {},
      expenses: expenses
    };
  } catch (e) {
    Logger.log('getReportDetails エラー: ' + e.message + '\nStack: ' + e.stack);
    return { 
      header: {}, 
      expenses: [] 
    };
  }
}

/**
 * (B) 申請ステータスを更新 (内部関数) (v20: 修正 / v23: データID使用)
 */
function updateReportStatus_(reportId, newStatus, rejectionReason) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    var ss = SpreadsheetApp.openById(DATA_SPREADSHEET_ID); // (v23: データIDを使用)
    var sheet = ss.getSheetByName(SHEET_NAMES.REPORTS); 
    
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('経費ヘッダーシートが見つからないか、空です。');
    }
    
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error('経費ヘッダーシートにカラムがありません。');
    }
    
    var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    
    var headers = data[0];
    var reportIdCol = headers.indexOf('経費ID');
    var statusCol = headers.indexOf('承認ステータス');
    var applyDateCol = headers.indexOf('申請日');
    var rejectionReasonCol = headers.indexOf('差戻し理由');
    
    if (reportIdCol === -1 || statusCol === -1) {
      throw new Error('シートの列(経費ID or 承認ステータス)が見つかりません。');
    }
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][reportIdCol]) === String(reportId)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        
        if (newStatus === '差戻し') {
          if (applyDateCol > -1) {
            sheet.getRange(i + 1, applyDateCol + 1).setValue(null);
          }
          if (rejectionReasonCol > -1 && rejectionReason) {
            sheet.getRange(i + 1, rejectionReasonCol + 1).setValue(rejectionReason);
          }
        }
        
        if (newStatus === '承認済') {
          if (rejectionReasonCol > -1) {
            sheet.getRange(i + 1, rejectionReasonCol + 1).setValue(null);
          }
        }
        
        return { status: "success", message: "ステータスを「" + newStatus + "」に更新しました。" };
      }
    }
    
    throw new Error('対象の申請データが見つかりません。');
  } catch (e) {
    Logger.log('updateReportStatus_ エラー: ' + e);
    return { status: "error", message: "ステータス更新エラー: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * (B) 承認
 */
function approveReport(reportId) {
  return updateReportStatus_(reportId, '承認済', null);
}

/**
 * (B) 差戻し
 */
function rejectReport(reportId, rejectionReason) {
  return updateReportStatus_(reportId, '差戻し', rejectionReason);
}


// -----------------------------------------------
// (D) 経費精算用 (v20: 修正あり / v21: 拡張)
// -----------------------------------------------

/**
 * (D) 承認済みの経費データを取得 (v20: 修正 / v21: areaCode取得 / v23: マスタ分離対応)
 */
function getApprovedExpenses() {
  var userInfo = getUserInfo_();
  var userId = userInfo.employeeInfo.id;
  
  // (v23: ssの即時オープンを削除)
  
  var allHeaders = getSheetData_(SHEET_NAMES.REPORTS); // (v23: ss引数を削除)
  var approvedReports = allHeaders.filter(function(row) {
    return String(row['申請者社員番号']) === String(userId) && row['承認ステータス'] === '承認済';
  });
  
  if (approvedReports.length === 0) {
    return { reports: [], expenses: [] };
  }
  
  var approvedReportIds = approvedReports.map(function(r) { return r['経費ID']; });
  
  var allExpenses = getSheetData_(SHEET_NAMES.EXPENSES); // (v23: ss引数を削除)
  var approvedExpenses = allExpenses.filter(function(exp) {
    return approvedReportIds.some(function(approvedId) {
      return String(exp['経費ID']) === String(approvedId);
    });
  });

  var expenseItemsMaster = getSheetData_(SHEET_NAMES.EXPENSE_ITEMS); // (v23: ss引数を削除)
  var expenseItemsMap = {};
  expenseItemsMaster.forEach(function(item) {
    expenseItemsMap[item['経費項目コード']] = item['経費項目'];
  });

  // (v21: エリアマスタ)
  var areaMaster = getSheetData_(SHEET_NAMES.AREA_MASTER); // (v23: ss引数を削除)
  var areaMap = {};
  areaMaster.forEach(function(item) {
    areaMap[item['地域区分 (コード)']] = item['地域区分名'];
  });
  
  approvedExpenses = approvedExpenses.map(function(exp) {
    exp['itemName'] = expenseItemsMap[exp['経費項目コード']] || exp['経費項目コード'];
    // (v21: 取得)
    var areaCode = exp['エリアコード'];
    exp['areaCode'] = areaCode;
    exp['areaName'] = areaMap[areaCode] || '';
    return exp;
  });

  return {
    reports: approvedReports,
    expenses: approvedExpenses
  };
}

/**
 * (D) 最終的な経費精算申請 (v20: 修正 / v23: データID使用)
 */
function submitFinalExpenses(reportIds) {
  if (!reportIds || reportIds.length === 0) {
    return { status: "error", message: "申請対象が選択されていません。" };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    var ss = SpreadsheetApp.openById(DATA_SPREADSHEET_ID); // (v23: データIDを使用)
    var sheet = ss.getSheetByName(SHEET_NAMES.REPORTS);
    
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('経費ヘッダーシートが見つからないか、空です。');
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error('経費ヘッダーシートにカラムがありません。');
    }
    var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    
    var headers = data[0];
    var reportIdCol = headers.indexOf('経費ID');
    var statusCol = headers.indexOf('承認ステータス');
    var applicantCol = headers.indexOf('申請者');
    var dateCol = headers.indexOf('日付');
    var totalCol = headers.indexOf('合計経費');
    
    if (reportIdCol === -1 || statusCol === -1) {
      throw new Error('シートの列(経費ID or 承認ステータス)が見つかりません。');
    }
    
    var updatedCount = 0;
    var applicantName = '';
    var totalAmount = 0;
    var targetDetails = [];

    var reportIdsStr = reportIds.map(function(id) { return String(id); });

    for (var i = 1; i < data.length; i++) {
      if (reportIdsStr.indexOf(String(data[i][reportIdCol])) > -1) {
        sheet.getRange(i + 1, statusCol + 1).setValue('精算申請済');
        updatedCount++;
        
        if (!applicantName) {
          applicantName = data[i][applicantCol];
        }
        totalAmount += parseFloat(data[i][totalCol] || 0);
        
        var dateStr = (data[i][dateCol] instanceof Date) 
          ? Utilities.formatDate(data[i][dateCol], Session.getScriptTimeZone(), 'yyyy/MM/dd')
          : data[i][dateCol];
        
        targetDetails.push('   - ' + dateStr + ' (¥' + data[i][totalCol] + ')');
      }
    }
    
    if (updatedCount === 0) {
      throw new Error('対象のデータが見つかりませんでした。');
    }
    
    var settings = getSettings_();
    var accountingEmail = settings.ACCOUNTING_EMAIL;
    if (accountingEmail) {
      MailApp.sendEmail({
        to: accountingEmail,
        subject: '[経費精算] 最終申請: ' + applicantName,
        body: applicantName + ' さんから経費の最終申請が届きました。\n\n' +
              '合計件数: ' + updatedCount + ' 件\n' +
              '合計金額: ' + totalAmount + ' 円\n\n' +
              '対象申請:\n' +
              targetDetails.join('\n') + '\n\n' +
              'システムで振込処理を行ってください。'
      });
    }

    return { status: "success", message: updatedCount + "件の経費精算を申請しました。" };
  } catch (e) {
    Logger.log('submitFinalExpenses エラー: ' + e);
    return { status: "error", message: "経費精算申請エラー: " + e.message };
  } finally {
    lock.releaseLock();
  }
}


// -----------------------------------------------
// (E) 申請一覧用 (v20: 修正あり / v21: 拡張)
// -----------------------------------------------

/**
 * (E) 自分の申請一覧を取得 (v20: 修正 / v23: マスタ分離対応)
 */
function getMyReports() {
  try {
    var userInfo = getUserInfo_();
    var targetEmployeeId = userInfo.employeeInfo.id;
    
    if (!targetEmployeeId) {
      return []; 
    }

    // (v23: ssの即時オープンを削除)
    var allData = getSheetData_(SHEET_NAMES.REPORTS); // (v23: ss引数を削除)

    var idxApplicant = '申請者社員番号';
    
    var myReports = [];
    
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][idxApplicant]) === String(targetEmployeeId)) {
        
        myReports.push({
          '経費ID': allData[i]['経費ID'],
          '日付': allData[i]['日付'],
          '件名': allData[i]['件名'],
          '承認ステータス': allData[i]['承認ステータス'],
          '合計経費': allData[i]['合計経費'],
          '備考': allData[i]['備考']
        });
      }
    }
    
    myReports.sort(function(a, b) {
      return b['日付'].localeCompare(a['日付']);
    });
    
    return myReports;

  } catch (e) {
    Logger.log('getMyReports エラー: ' + e.message + '\nStack: ' + e.stack);
    throw new Error('申請一覧の取得に失敗しました: ' + e.message);
  }
}

/**
 * (E) 編集用に申請データを取得 (v20: 修正 / v21: 拡張)
 * (v21: getReportDetailsが拡張されているため、この関数は修正不要)
*/
function getReportDataForEdit(reportId) {
  try {
    var userInfo = getUserInfo_();
    var userId = userInfo.employeeInfo.id;
    
    // (v21: getReportDetailsは既にareaCode/areaNameを含むよう修正済み)
    var details = getReportDetails(reportId); 
    
    if (!details.header || !details.header['経費ID']) {
      Logger.log('getReportDataForEdit: getReportDetails が空のヘッダーを返しました。ReportID: ' + reportId);
      return { status: "error", message: "対象の申請データが見つかりません。" };
    }
    
    if (String(details.header['申請者社員番号']) !== String(userId)) {
      Logger.log('getReportDataForEdit: 権限エラー。 UserID: ' + userId + ', ReportApplicantID: ' + details.header['申請者社員番号']);
      return { status: "error", message: "この申請を編集する権限がありません。" };
    }
    
    var status = details.header['承認ステータス'];
    if (status !== '一時保存' && status !== '差戻し') {
      Logger.log('getReportDataForEdit: ステータスエラー。 Status: ' + status);
      return { status: "error", message: "「" + status + "」の申請は編集できません。" };
    }
    
    return { status: "success", data: details };
  } catch (e) {
    Logger.log('getReportDataForEdit エラー: ' + e.message + '\nStack: ' + e.stack);
    return { status: "error", message: "編集データの読み込みエラー: " + e.message };
  }
}


// -----------------------------------------------
// 汎用ヘルパー関数
// -----------------------------------------------

/**
 * 汎用: シートデータをオブジェクト配列に変換 (v17 修正 / v23: ID自動選択)
 */
function getSheetData_(sheetName) { // (v23: ss引数を削除)
  var idToUse;
  
  // (v23: 'マスタ' が含まれるシート名は MASTER_SPREADSHEET_ID を使用)
  if (sheetName.indexOf('マスタ') > -1) {
    idToUse = MASTER_SPREADSHEET_ID;
    if (!idToUse) {
      Logger.log('エラー: スクリプトプロパティに MASTER_SPREADSHEET_ID が設定されていません。');
      throw new Error('マスターSpreadsheet IDが設定されていません。');
    }
  } else {
    idToUse = DATA_SPREADSHEET_ID;
    if (!idToUse) {
      Logger.log('エラー: スクリプトプロパティに SPREADSHEET_ID が設定されていません。');
      throw new Error('データSpreadsheet IDが設定されていません。');
    }
  }
  
  var ss = SpreadsheetApp.openById(idToUse);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + sheetName + ' (ID: ' + idToUse + ')');
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return [];
  }
  
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data.shift(); 
  
  var result = data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      if (row[index] instanceof Date) {
        try {
          var date = row[index];
          if (date.getFullYear() === 1899 && date.getMonth() === 11 && date.getDate() === 30) {
             obj[header] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm');
          }
          else if (date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0) {
             obj[header] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
          } 
          else {
             obj[header] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
          }
        } catch(e) {
           obj[header] = row[index].toString();
        }
      } else {
        obj[header] = row[index];
      }
    });
    return obj;
  });
  
  return result;
}

/**
 * 汎用: 特定の経費IDに関連するデータを削除 (v20: 修正 / v23: データID使用)
 */
function deleteReportData_(reportId, ss) {
  if (!reportId) {
    return;
  }
  
  if (!ss) {
    ss = SpreadsheetApp.openById(DATA_SPREADSHEET_ID); // (v23: データIDを使用)
  }
  
  var sheetsToDelete = [
    SHEET_NAMES.REPORTS, 
    SHEET_NAMES.EXPENSES
  ];

  sheetsToDelete.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }
    
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      return;
    }
    
    var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    var headers = data[0];
    
    var reportIdCol = headers.indexOf('経費ID'); 
    if (reportIdCol === -1) {
      reportIdCol = headers.indexOf('日報ID'); 
      if(reportIdCol === -1) {
         return;
      }
    }
    
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][reportIdCol]) === String(reportId)) {
        rowsToDelete.push(i + 1);
      }
    }
    
    rowsToDelete.sort(function(a, b) { return b - a; });
    
    rowsToDelete.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
    });
  });
}

/**
 * (v22: デバッグ用) 
 * 設定シートのキャッシュを強制的にクリアします。
 */
function clearSettingsCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('SETTINGS_CACHE');
  Logger.log('設定キャッシュ (SETTINGS_CACHE) をクリアしました。');
  Browser.msgBox('設定キャッシュをクリアしました。');
}
