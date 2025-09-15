/**
 * 在試算表菜單中添加一個自定義菜單項。
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('管理帳號與群組')
    .addItem('依[新建更新]表批次處理', 'processUsersAndGroups_V2')
    .addSeparator()
    .addItem('1.匯出[全部@tea清單]"', 'exportAllUsers')
    .addItem('2.依據匯出sheet 只更新使用者姓、名、機構單位、職稱', 'updateUsersFromSheet')
    .addSeparator()
    .addItem('1.匯出群組成員 (互動式)', 'showGroupManagementSidebar')
    .addItem('2.依據匯出的sheet更新群組成員', 'updateGroupMembersFromSheet') // 【新增這個功能】
    .addSeparator()
    .addItem('匯出所有機構單位 (含人數)', 'exportOUsAndUserCounts')
    .addToUi();
}

/**
 * [優化版] 處理試算表中的使用者資料，新增/更新帳號並支援加入多個指定群組。
 */
function processUsersAndGroups_V2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var headers = values[0];
  var data = values.slice(1);

  // 查找各欄位的索引
  var firstNameCol = headers.indexOf('名');
  var lastNameCol = headers.indexOf('姓');
  var emailCol = headers.indexOf('Email Address [Required]');
  var passwordCol = headers.indexOf('空白代表不改密碼');
  var orgUnitPathCol = headers.indexOf('機構路徑');
  var employeeTitleCol = headers.indexOf('Employee Title');
  var groupEmailCol = headers.indexOf('加入群組');

  if ([firstNameCol, lastNameCol, emailCol, passwordCol, orgUnitPathCol, groupEmailCol].includes(-1)) {
    SpreadsheetApp.getUi().alert('錯誤', '試算表標題欄位不正確，請確保包含: 名, 姓, Email Address [Required], 空白代表不改密碼, 機構路徑, 加入群組。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var successCount = 0;
  var failCount = 0;
  var noActionCount = 0; // [優化] 新增計數器，用於記錄“無需操作”的情況
  var logMessages = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = String(row[emailCol] || '').trim();
    if (!email) {
      continue; // 如果 Email 為空，直接跳過此行
    }

    var logPrefix = '第 ' + (i + 2) + ' 行 (' + email + '): ';

    try {
      var firstName = String(row[firstNameCol] || '').trim();
      var lastName = String(row[lastNameCol] || '').trim();
      var password = String(row[passwordCol] || '').trim();
      var orgUnitPath = String(row[orgUnitPathCol] || '').trim();
      var employeeTitle = String(row[employeeTitleCol] || '').trim();
      var groupEmails = String(row[groupEmailCol] || '').trim();

      if (!firstName || !lastName || !orgUnitPath) {
        logMessages.push(logPrefix + '錯誤 - 必要的欄位 (名, 姓, 機構路徑) 不完整，跳過。');
        failCount++;
        continue;
      }

      var user;
      try {
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail" }); // 優化：只獲取必要的欄位，API 調用更輕量
      } catch (e) {
        user = null;
      }

      var userObj = {
        name: { givenName: firstName, familyName: lastName },
        orgUnitPath: orgUnitPath,
        // 如果 employeeTitle 為空字串，API 可能會報錯，所以只有在有值時才加入
        ...(employeeTitle && { title: employeeTitle })
      };

      if (user) { // 使用者已存在，執行更新
        logMessages.push(logPrefix + '帳號已存在，密碼不修改。');
        AdminDirectory.Users.update(userObj, email);
        logMessages.push(logPrefix + '使用者帳號其他資訊已更新。');
      } else { // 使用者不存在，執行新增
        if (!password) {
          logMessages.push(logPrefix + '錯誤 - 創建新使用者時「空白代表不改密碼」欄位不能為空。');
          failCount++;
          continue;
        }
        userObj.primaryEmail = email;
        userObj.password = password;
        userObj.changePasswordAtNextLogin = true;
        AdminDirectory.Users.insert(userObj);
        logMessages.push(logPrefix + '使用者帳號已成功創建。');
      }

      // [優化] 處理多個群組
      if (groupEmails) {
        var groups = groupEmails.split(',').map(function (g) { return g.trim(); });
        for (var j = 0; j < groups.length; j++) {
          var groupEmail = groups[j];
          if (groupEmail) {
            try {
              AdminDirectory.Members.insert({ email: email, role: "MEMBER" }, groupEmail);
              logMessages.push(logPrefix + '已成功加入群組 ' + groupEmail + '。');
            } catch (groupError) {
              // 檢查是否為“成員已存在”的錯誤
              if (groupError.message.includes("Member already exists") || groupError.message.includes("duplicate")) {
                logMessages.push(logPrefix + '已是群組 ' + groupEmail + ' 的成員，無需操作。');
                noActionCount++; // 歸入“無須操作”計數
              } else {
                // 其他所有群組相關錯誤（包括權限問題）都視為失敗
                logMessages.push(logPrefix + '加入群組 ' + groupEmail + ' 時失敗: ' + groupError.message);
                failCount++;
              }
            }
          }
        }
      }

      successCount++;

    } catch (e) {
      logMessages.push(logPrefix + '處理帳號時發生嚴重錯誤: ' + e.message);
      failCount++;
    }

    // Utilities.sleep(300); // 如果處理大量資料(>100筆)，建議取消此行註解以避免 API 速率限制
  }

  var resultMsg = '帳號與群組處理完成！\n\n' +
    '成功處理行數: ' + successCount + '\n' +
    '失敗/錯誤數: ' + failCount + '\n' +
    '無需操作數 (例如成員已存在): ' + noActionCount + '\n\n' + // [優化] 新增報告項
    '詳細日誌請查看 Apps Script 編輯器中的「執行作業」。\n\n' +
    '--- 部分日誌預覽 ---\n' + logMessages.slice(0, 15).join('\n') + (logMessages.length > 15 ? '\n...(更多日誌省略)' : '');

  ui.alert('處理結果', resultMsg, ui.ButtonSet.OK);
  Logger.log('--- 完整執行日誌 ---\n' + logMessages.join('\n'));
}

/**
 * [升級版][危險操作] 清除指定 Google 群組中的所有成員。
 * 此函數可以獨立執行（彈出輸入框），也可以接收從側邊欄傳來的 groupEmail。
 * @param {string} [groupEmail] (可選) 從側邊欄傳遞過來的群組 Email。
 * @returns {object} 回傳一個包含操作結果的物件給側邊欄。
 */
function clearGroupMembers(groupEmail) {
  var ui = SpreadsheetApp.getUi();

  // 如果函數不是由側邊欄觸發（沒有傳入 groupEmail），則彈出輸入框
  if (!groupEmail) {
    var response = ui.prompt(
      '危險操作確認',
      '您即將清除一個群組中的所有成員。\n此操作不可逆！\n請輸入完整的群組 Email 以確認執行:',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() != ui.Button.OK) {
      return { message: '操作已取消。', success: true };
    }
    groupEmail = response.getResponseText().trim();
  }

  if (!groupEmail || groupEmail.indexOf('@') === -1) {
    ui.alert('輸入無效', '您沒有提供有效的群組 Email，操作已取消。', ui.ButtonSet.OK);
    return { message: '輸入無效，操作已取消。', success: false };
  }

  // 第二重安全確認：再次彈窗確認目標
  var finalConfirmation = ui.alert(
    '最終確認',
    '您【確定】要將群組【' + groupEmail + '】中的所有成員都移除嗎？\n\n這個動作無法復原！',
    ui.ButtonSet.YES_NO
  );

  if (finalConfirmation != ui.Button.YES) {
    ui.alert('操作已取消。'); // 彈窗提示使用者
    return { message: '操作已取消。', success: true }; // 回傳結果給側邊欄
  }

  var removedCount = 0;
  var errorCount = 0;
  var allMembers = [];

  try {
    var pageToken;
    do {
      var page = AdminDirectory.Members.list(groupEmail, { maxResults: 500, pageToken: pageToken });
      if (page.members) {
        allMembers = allMembers.concat(page.members);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allMembers.length === 0) {
      var noMemberMsg = '群組 ' + groupEmail + ' 本身就是空的，無需操作。';
      ui.alert('提示', noMemberMsg, ui.ButtonSet.OK);
      return { message: noMemberMsg, success: true };
    }

    for (var i = 0; i < allMembers.length; i++) {
      try {
        AdminDirectory.Members.remove(groupEmail, allMembers[i].email);
        removedCount++;
      } catch (e) {
        errorCount++;
      }
    }

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
    ui.alert('錯誤', '無法處理群組 ' + groupEmail + '。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
    return { message: errorMsg, success: false };
  }

  var resultMsg = '群組清除作業完成！\n\n' +
    '目標群組: ' + groupEmail + '\n' +
    '成功移除成員數: ' + removedCount + '\n' +
    '失敗數: ' + errorCount;

  ui.alert('作業報告', resultMsg, ui.ButtonSet.OK);
  return { message: resultMsg.replace(/\n/g, '<br>'), success: true };
}
/**
 * 匯出指定 Google 群組的所有成員到一個新的工作表中。
 */
function exportGroupMembersToSheet() {
  var ui = SpreadsheetApp.getUi();

  // 彈出輸入框，要求使用者輸入群組 Email
  var response = ui.prompt(
    '匯出群組成員',
    '請輸入您想要匯出成員列表的群組 Email (例如: staffmembers@tea.nknush.kh.edu.tw):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('操作已取消。');
    return;
  }

  var groupEmail = response.getResponseText().trim();
  if (!groupEmail || groupEmail.indexOf('@') === -1) {
    ui.alert('輸入無效', '您沒有輸入有效的群組 Email，操作已取消。', ui.ButtonSet.OK);
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在讀取成員列表，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始讀取群組: ' + groupEmail];
  var allMembers = [];

  try {
    // 處理分頁，循環獲取所有成員
    var pageToken;
    do {
      var page = AdminDirectory.Members.list(groupEmail, {
        maxResults: 500,
        pageToken: pageToken
      });
      if (page.members) {
        allMembers = allMembers.concat(page.members);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allMembers.length === 0) {
      ui.alert('結果', '群組 ' + groupEmail + ' 中沒有任何成員。', ui.ButtonSet.OK);
      return;
    }

    logMessages.push('共找到 ' + allMembers.length + ' 位成員，開始寫入新工作表...');

    // 準備要寫入工作表的資料 (2D 陣列)
    var outputData = [['成員 Email', '角色 (Role)', '類型 (Type)', '狀態 (Status)']]; // 標題行
    for (var i = 0; i < allMembers.length; i++) {
      var member = allMembers[i];
      outputData.push([member.email, member.role, member.type, member.status]); // 資料行
    }

    // 建立新的工作表
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "[群組成員] " + groupEmail.split('@')[0];
    var newSheet = spreadsheet.insertSheet(sheetName);

    // 將資料一次性寫入新工作表
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 設定固定欄位寬度和自動裁剪
    newSheet.setColumnWidth(1, 200); // 成員 Email
    newSheet.setColumnWidth(2, 80);  // 角色 (Role)
    newSheet.setColumnWidth(3, 80);  // 類型 (Type)
    newSheet.setColumnWidth(4, 80);  // 狀態 (Status)

    // 設定資料範圍的自動裁剪
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 4);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
    }

    // 切換到新建立的工作表，讓使用者可以直接看到結果
    newSheet.activate();

    ui.alert('匯出成功！', allMembers.length + ' 位成員的資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法讀取群組 ' + groupEmail + ' 的成員。\n\n請檢查群組是否存在，或您是否有權限查看此群組的成員。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
  }
}

// =================================================================================
// ============ 互動式側邊欄 - 讀取群組與成員功能 (開始) =====================
// =================================================================================

/**
 * 顯示一個包含所有群組列表的側邊欄，讓使用者可以選擇並查詢成員。
 */
function showGroupManagementSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('群組成員查詢工具')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * [升級版] 獲取網域中的所有群組列表，包含每個群組的成員數量。
 * @returns {Array} 一個包含群組物件 {name, email, memberCount} 的陣列。
 */
function listAllGroups() {
  var allGroups = [];
  var pageToken;
  try {
    do {
      var page = AdminDirectory.Groups.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        // 【主要變更 1】明確指定我們需要的欄位，包含 directMembersCount
        fields: 'nextPageToken,groups(name,email,directMembersCount)'
      });
      if (page.groups) {
        // 【主要變更 2】提取需要的資訊，並將成員數量也加入
        var groups = page.groups.map(function (group) {
          return {
            name: group.name,
            email: group.email,
            memberCount: group.directMembersCount || 0 // 如果沒有這個欄位，預設為 0
          };
        });
        allGroups = allGroups.concat(groups);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    allGroups.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    return allGroups;
  } catch (e) {
    Logger.log('無法獲取群組列表: ' + e.toString());
    return [{ error: '無法獲取群組列表: ' + e.message }];
  }
}


/**
 * [最終版] 根據給定的群組 Email，獲取其所有成員（包含姓名、最後登入時間、機構單位路徑和所屬群組），並直接匯出到一個新的工作表。
 * 這個函數會被 HTML 側邊欄呼叫。
 * @param {string} groupEmail 要查詢並匯出的群組 Email。
 * @returns {object} 一個包含操作結果的物件。
 */
function getGroupMembersForSidebar(groupEmail) {
  if (!groupEmail) {
    return { success: false, message: '未提供有效的群組 Email。' };
  }

  var allMembers = [];
  var pageToken;

  try {
    // 步驟 1: 獲取所有成員列表
    do {
      var page = AdminDirectory.Members.list(groupEmail, {
        maxResults: 500,
        pageToken: pageToken
      });
      if (page.members) {
        allMembers = allMembers.concat(page.members);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allMembers.length === 0) {
      return { success: true, message: '群組 ' + groupEmail + ' 中沒有任何成員。', noMembers: true };
    }

    // 步驟 2: 準備要寫入工作表的資料
    var outputData = [[
      '成員 Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '最後登入時間 (Last Login)',
      '角色 (Role)',
      '類型 (Type)',
      '狀態 (Status)',
      '機構單位路徑 (OU Path)',
      '所屬群組 (Groups)',
      '是否需要更新'
    ]];

    // 步驟 3: 遍歷每一位成員，以獲取他們的詳細資訊
    for (var i = 0; i < allMembers.length; i++) {
      var member = allMembers[i];
      var firstName = '';
      var lastName = '';
      var lastLogin = 'N/A';
      var orgUnitPath = 'N/A';
      var userGroups = 'N/A';

      if (member.type === 'USER') {
        try {
          // 獲取使用者基本資訊（包含機構單位路徑）
          var user = AdminDirectory.Users.get(member.email, {
            fields: 'name,lastLoginTime,orgUnitPath'
          });
          firstName = user.name.givenName || '';
          lastName = user.name.familyName || '';
          orgUnitPath = user.orgUnitPath || '/';

          // 處理並格式化最後登入時間
          if (user.lastLoginTime) {
            var loginDate = new Date(user.lastLoginTime);
            if (loginDate.getFullYear() > 1970) {
              lastLogin = loginDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
            } else {
              lastLogin = '從未登入';
            }
          } else {
            lastLogin = '從未登入';
          }

          // 獲取使用者所屬的所有群組
          try {
            var memberGroups = [];
            var groupPageToken;
            do {
              var groupPage = AdminDirectory.Groups.list({
                userKey: member.email,
                maxResults: 200,
                pageToken: groupPageToken,
                fields: 'nextPageToken,groups(name)'
              });
              if (groupPage.groups) {
                for (var g = 0; g < groupPage.groups.length; g++) {
                  memberGroups.push(groupPage.groups[g].name);
                }
              }
              groupPageToken = groupPage.nextPageToken;
            } while (groupPageToken);

            userGroups = memberGroups.length > 0 ? memberGroups.join(', ') : '無群組';
          } catch (groupError) {
            userGroups = '無法獲取';
            Logger.log('無法獲取使用者 ' + member.email + ' 的群組資訊: ' + groupError.message);
          }

        } catch (userError) {
          firstName = 'N/A';
          lastName = 'N/A';
          lastLogin = '無法獲取';
          orgUnitPath = '無法獲取';
          userGroups = '無法獲取';
          Logger.log('無法獲取使用者 ' + member.email + ' 的詳細資訊: ' + userError.message);
        }
      } else {
        firstName = '(巢狀群組)';
        lastName = '(Nested Group)';
        lastLogin = '不適用';
        orgUnitPath = '不適用';
        userGroups = '不適用';
      }

      // 將包含新欄位的完整資料加入到輸出陣列中
      outputData.push([
        member.email,
        lastName,
        firstName,
        lastLogin,
        member.role,
        member.type,
        member.status,
        orgUnitPath,
        userGroups,
        '無需更新' // 預設為無需更新
      ]);
    }

    // 步驟 4: 建立新的工作表
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var safeSheetName = "[群組成員] "+groupEmail.split('@')[0].replace(/[^\w\s]/gi, '_');

    var existingSheet = spreadsheet.getSheetByName(safeSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(safeSheetName, 0);

    // 步驟 5: 一次性寫入所有資料（移除說明行）
    newSheet.getRange(1, 1, outputData.length, 10).setValues(outputData);

    // 步驟 6: 在工作表底部建立原始值參考區域（儲存 B、C、I 欄的原始值）
    var referenceStartRow = outputData.length + 3; // 留空間避免衝突
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '']]; // 修正：改為 3 個元素

    // 複製 B、C、I 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) { // 從第2行開始（跳過標題）
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][8]  // I欄：所屬群組 (Groups)
      ]);
    }

    // 寫入參考區域
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 3).setValues(referenceData); // 修正：改為 3 欄

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 7: 設定檢測公式（只檢測 B、C、I 欄的變化）
    // 資料行從第2行開始（第1行是標題）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var dataIndex = rowIndex - 1; // 對應到 outputData 中的索引（第2行對應 outputData[1]）
      var refRowIndex = referenceStartRow + dataIndex; // 對應的參考行

      // 只有在是資料行時才設定檢測公式（跳過標題行）
      if (dataIndex >= 1 && dataIndex < outputData.length) {
        var detectionFormula =
          '=IF(OR(' +
          'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
          'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
          'I' + rowIndex + '<>$C$' + refRowIndex +        // I欄：所屬群組 ✅ 修正！
          '),"需要更新","無需更新")';

        newSheet.getRange(rowIndex, 10).setFormula(detectionFormula); // J欄（第10欄）
      }
    }

    // 步驟 8: 設定範圍保護 + 視覺提示
    if (outputData.length > 1) {
      var dataRowCount = outputData.length - 1;
      
      // 對每個不可編輯的範圍設定個別保護
      var protectedRanges = [
        newSheet.getRange(2, 1, dataRowCount, 1),  // A欄：Email
        newSheet.getRange(2, 4, dataRowCount, 1),  // D欄：最後登入
        newSheet.getRange(2, 5, dataRowCount, 1),  // E欄：角色
        newSheet.getRange(2, 6, dataRowCount, 1),  // F欄：類型
        newSheet.getRange(2, 7, dataRowCount, 1),  // G欄：狀態
        newSheet.getRange(2, 8, dataRowCount, 1),  // H欄：機構單位
        newSheet.getRange(2, 10, dataRowCount, 1)  // J欄：檢測狀態
      ];

      // 為每個不可編輯範圍設定保護
      for (var p = 0; p < protectedRanges.length; p++) {
        var protection = protectedRanges[p].protect()
          .setDescription('系統產生的唯讀資料 - 請勿修改');
        
        // 移除所有編輯者（包括擁有者），但這對擁有者無效
        protection.removeEditors(protection.getEditors());
        
        // 設為警告模式，至少會彈出提醒
        protection.setWarningOnly(true);
      }

      // 用強烈的視覺區別
      // 可編輯欄位：綠色背景
      newSheet.getRange(2, 2, dataRowCount, 1).setBackground('#C8E6C9'); // B欄：綠色
      newSheet.getRange(2, 3, dataRowCount, 1).setBackground('#C8E6C9'); // C欄：綠色  
      newSheet.getRange(2, 9, dataRowCount, 1).setBackground('#C8E6C9'); // I欄：綠色

      // 不可編輯欄位：灰色背景 + 斜體
      var readOnlyRanges = [
        newSheet.getRange(2, 1, dataRowCount, 1),  // A欄
        newSheet.getRange(2, 4, dataRowCount, 1),  // D欄
        newSheet.getRange(2, 5, dataRowCount, 1),  // E欄
        newSheet.getRange(2, 6, dataRowCount, 1),  // F欄
        newSheet.getRange(2, 7, dataRowCount, 1),  // G欄
        newSheet.getRange(2, 8, dataRowCount, 1),  // H欄
        newSheet.getRange(2, 10, dataRowCount, 1)  // J欄
      ];

      for (var r = 0; r < readOnlyRanges.length; r++) {
        readOnlyRanges[r]
          .setBackground('#EEEEEE')       // 灰色背景
          .setFontStyle('italic')         // 斜體字
          .setFontColor('#666666');       // 灰色文字
      }

      // 在標題行加上視覺提示
      var headerRange = newSheet.getRange(1, 1, 1, 10);
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');

      // 可編輯欄位的標題加上 ✅ 符號
      newSheet.getRange(1, 2).setValue('✅ 姓 (Family Name)');
      newSheet.getRange(1, 3).setValue('✅ 名 (Given Name)');
      newSheet.getRange(1, 9).setValue('✅ 所屬群組 (Groups)');

      // 不可編輯欄位的標題加上 🔒 符號
      newSheet.getRange(1, 1).setValue('🔒 成員 Email');
      newSheet.getRange(1, 4).setValue('🔒 最後登入時間 (Last Login)');
      newSheet.getRange(1, 5).setValue('🔒 角色 (Role)');
      newSheet.getRange(1, 6).setValue('🔒 類型 (Type)');
      newSheet.getRange(1, 7).setValue('🔒 狀態 (Status)');
      newSheet.getRange(1, 8).setValue('🔒 機構單位路徑 (OU Path)');
      newSheet.getRange(1, 10).setValue('🔒 是否需要更新');
    }

    // 步驟 9: 設定欄位寬度（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 60,  // A欄：成員 Email
      2: 60,  // B欄：姓 (Family Name)
      3: 60,  // C欄：名 (Given Name)
      4: 60,  // D欄：最後登入時間
      5: 50,   // E欄：角色 (Role)
      6: 50,   // F欄：類型 (Type)
      7: 50,   // G欄：狀態 (Status)
      8: 300,  // H欄：機構單位路徑
      9: 200,  // I欄：所屬群組 (Groups)
      10: 60  // J欄：是否需要更新
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 10; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定所有資料範圍的自動裁剪（文字換行）
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 10);
      dataRange.setWrap(true); // 啟用自動換行以適應固定寬度
      dataRange.setVerticalAlignment('top'); // 垂直對齊頂部
    }

    newSheet.setFrozenRows(1); // 凍結標題行

    // 步驟 10: 設定「是否需要更新」欄位的條件格式
    if (outputData.length > 1) {
      var detectionRange = newSheet.getRange(2, 10, outputData.length - 1, 1); // J欄（第10欄）- 修正！

      var needUpdateRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("需要更新")
        .setBackground("#FFA500")
        .setFontColor("#FFFFFF")
        .setRanges([detectionRange])
        .build();

      var noUpdateRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("無需更新")
        .setBackground("#90EE90")
        .setFontColor("#000000")
        .setRanges([detectionRange])
        .build();

      var alreadyUpdatedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("已更新")
        .setBackground("#87CEEB")
        .setFontColor("#000000")
        .setRanges([detectionRange])
        .build();

      var rules = newSheet.getConditionalFormatRules();
      rules.push(needUpdateRule);
      rules.push(noUpdateRule);
      rules.push(alreadyUpdatedRule);
      newSheet.setConditionalFormatRules(rules);
    }

    // 步驟 11: 回傳成功的結果給側邊欄
    return {
      success: true,
      sheetName: safeSheetName,
      memberCount: allMembers.length
    };

  } catch (e) {
    Logger.log('從側邊匯出群組 ' + groupEmail + ' 成員時失敗: ' + e.toString());
    return { success: false, message: '無法獲取成員: ' + e.message };
  }
}

// =================================================================================
// ============ 互動式側邊欄 - 讀取群組與成員功能 (結束) =====================
// =================================================================================

// =================================================================================
// ============ 匯出機構單位與人數功能 (開始) ========================
// =================================================================================

/**
 * 掃描整個網域，獲取所有機構單位 (OU) 及其內部的使用者數量，並匯出到一個新的工作表。
 */
function exportOUsAndUserCounts() {
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在掃描您的組織結構與使用者，這可能需要一些時間，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始掃描機構單位與使用者...'];

  try {
    // --- 步驟 1: 獲取所有使用者，並在記憶體中計算每個 OU 的人數 ---
    var ouUserCounts = {};
    var pageToken;
    do {
      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        fields: 'nextPageToken,users(orgUnitPath)' // 只獲取我們需要的 orgUnitPath 欄位，極大提升效率
      });
      if (page.users) {
        page.users.forEach(function (user) {
          var ouPath = user.orgUnitPath;
          if (ouUserCounts[ouPath]) {
            ouUserCounts[ouPath]++;
          } else {
            ouUserCounts[ouPath] = 1;
          }
        });
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    logMessages.push('使用者人數統計完成。');

    // --- 步驟 2: 獲取所有機構單位 ---
    var allOUs = [];
    pageToken = null; // 重置 pageToken
    do {
      var ouPage = AdminDirectory.Orgunits.list({
        customerId: 'C01mdd9w2',
        pageToken: pageToken
      });
      if (ouPage.organizationUnits) {
        allOUs = allOUs.concat(ouPage.organizationUnits);
      }
      pageToken = ouPage.nextPageToken;
    } while (pageToken);

    logMessages.push('機構單位列表獲取完成，共找到 ' + allOUs.length + ' 個子單位。');

    // --- 步驟 3: 合併數據並準備匯出 ---
    var outputData = [['機構單位路徑 (OU Path)', '機構單位名稱 (OU Name)', '使用者人數']];

    for (var i = 0; i < allOUs.length; i++) {
      var ou = allOUs[i];
      var count = ouUserCounts[ou.orgUnitPath] || 0; // 如果某個 OU 是空的，人數為 0
      outputData.push([ou.orgUnitPath, ou.name, count]);
    }

    // 手動加入根機構單位 ("/")，因為 API 不會將其作為一個單位返回
    var rootCount = ouUserCounts['/'] || 0;
    outputData.push(['/', '根機構單位 (Root)', rootCount]);

    // 依照路徑排序，方便閱讀
    // 我們將根單位暫時移出，排序後再放回第一位
    var rootRow = outputData.pop();
    outputData.sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });
    outputData.unshift(rootRow); // 將根單位放回最前面

    // --- 步驟 4: 建立新工作表並寫入資料 ---
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "機構單位人數統計";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // 設定固定欄位寬度
    newSheet.setColumnWidth(1, 250); // 機構單位路徑
    newSheet.setColumnWidth(2, 200); // 機構單位名稱
    newSheet.setColumnWidth(3, 100); // 使用者人數

    // 設定資料範圍的自動裁剪
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 3);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
    }

    newSheet.activate();

    ui.alert('匯出成功！', '包含 ' + (outputData.length - 1) + ' 個機構單位的統計資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法完成機構單位掃描。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉側邊欄的 "處理中" 提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}

/**
 * 匯出整個 tea 網域中的所有使用者資料到一個新的工作表。
 * 包含使用者的基本資訊、機構單位、最後登入時間等詳細資訊。
 */
function exportAllUsers() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '匯出所有使用者',
    '您即將匯出整個 tea 網域的所有使用者清單。\n\n此操作可能需要較長時間，確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在讀取所有使用者資料，這可能需要幾分鐘時間，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始讀取所有使用者...'];
  var allUsers = [];
  var processedCount = 0;

  try {
    // 步驟 1: 獲取所有使用者
    var pageToken;
    do {
      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,organizations,suspended,creationTime,lastLoginTime)'
      });

      if (page.users) {
        allUsers = allUsers.concat(page.users);
        processedCount += page.users.length;
        logMessages.push('已讀取 ' + processedCount + ' 位使用者...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allUsers.length === 0) {
      ui.alert('結果', '未找到任何使用者。', ui.ButtonSet.OK);
      return;
    }

    logMessages.push('使用者資料讀取完成，共 ' + allUsers.length + ' 位使用者，開始整理資料...');

    // 步驟 2: 準備要寫入工作表的資料
    var outputData = [[
      '使用者 Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      'Employee Title',
      '帳號狀態',
      '建立時間',
      '最後登入時間',
      '是否需要更新'
    ]];

    // 步驟 3: 處理每位使用者的資料
    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];

      var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
      var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
      var orgUnitPath = user.orgUnitPath || '/';

      var employeeTitle = 'N/A';
      if (user.organizations && user.organizations.length > 0) {
        for (var j = 0; j < user.organizations.length; j++) {
          var org = user.organizations[j];
          if (org.title) {
            employeeTitle = org.title;
            break;
          }
        }
      }

      var status = user.suspended ? '已停用' : '啟用中';

      var creationTime = 'N/A';
      if (user.creationTime) {
        var createdDate = new Date(user.creationTime);
        creationTime = createdDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
      }

      var lastLoginTime = 'N/A';
      if (user.lastLoginTime) {
        var loginDate = new Date(user.lastLoginTime);
        if (loginDate.getFullYear() > 1970) {
          lastLoginTime = loginDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
        } else {
          lastLoginTime = '從未登入';
        }
      } else {
        lastLoginTime = '從未登入';
      }

      outputData.push([
        user.primaryEmail,
        familyName,
        givenName,
        orgUnitPath,
        employeeTitle,
        status,
        creationTime,
        lastLoginTime,
        '無需更新'
      ]);
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '_');
    var sheetName = "[全部@tea清單]" + timestamp;

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 先寫入資料（不包含公式）
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 步驟 5: 在工作表底部建立原始值參考區域
    var referenceStartRow = outputData.length + 3;
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '']]; // 修正：改為 4 個元素

    // 複製 B、C、D、E 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) { // 從第2行開始（跳過標題）
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4]  // E欄：Employee Title
      ]);
    }

    // 寫入參考區域
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 4).setValues(referenceData); // 修正：改為 4 欄

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E 欄的變化）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var refRowIndex = referenceStartRow + (rowIndex - 1); // 對應的參考行

      var detectionFormula =
        '=IF(OR(' +
        'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
        'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
        'D' + rowIndex + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
        'E' + rowIndex + '<>$D$' + refRowIndex +        // E欄：Employee Title
        '),"需要更新","無需更新")';

      newSheet.getRange(rowIndex, 9).setFormula(detectionFormula); // I欄（第9欄）
    }

    // 步驟 7: 設定格式（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 200,  // A欄：使用者 Email
      2: 100,  // B欄：姓 (Family Name)
      3: 100,  // C欄：名 (Given Name)
      4: 180,  // D欄：機構單位路徑
      5: 120,  // E欄：Employee Title
      6: 80,   // F欄：帳號狀態
      7: 150,  // G欄：建立時間
      8: 150,  // H欄：最後登入時間
      9: 120   // I欄：是否需要更新
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 9; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定所有資料範圍的自動裁剪（文字換行）
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 9);
      dataRange.setWrap(true); // 啟用自動換行以適應固定寬度
      dataRange.setVerticalAlignment('top'); // 垂直對齊頂部
    }

    newSheet.setFrozenRows(1); // 凍結標題行

    // 步驟 8: 設定「是否需要更新」欄位的條件格式
    if (outputData.length > 1) {
      var detectionRange = newSheet.getRange(2, 9, outputData.length - 1, 1); // I欄（第9欄）

      var needUpdateRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("需要更新")
        .setBackground("#FFA500")
        .setFontColor("#FFFFFF")
        .setRanges([detectionRange])
        .build();

      var noUpdateRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("無需更新")
        .setBackground("#90EE90")
        .setFontColor("#000000")
        .setRanges([detectionRange])
        .build();

      var alreadyUpdatedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("已更新")
        .setBackground("#87CEEB")
        .setFontColor("#000000")
        .setRanges([detectionRange])
        .build();

      var rules = newSheet.getConditionalFormatRules();
      rules.push(needUpdateRule);
      rules.push(noUpdateRule);
      rules.push(alreadyUpdatedRule);
      newSheet.setConditionalFormatRules(rules);
    }

    newSheet.activate();

    ui.alert('匯出成功！', allUsers.length + ' 位使用者的資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法完成使用者匯出。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉側邊欄的 "處理中" 提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}

/**
 * 根據試算表中的資料更新使用者的機構單位路徑和職稱。
 * 讀取目前工作表中的資料，並更新對應使用者的 orgUnitPath 和 Employee Title。
 * 只處理 I 欄標記為「需要更新」的行。
 */
function updateUsersFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新使用者資訊',
    '此功能將讀取目前工作表的資料，並更新使用者的姓名、機構單位路徑和職稱。\n\n' +
    '★ 智能更新：只會處理 I 欄標記為「需要更新」的使用者。\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的使用者資料\n' +
    '2. 資料格式正確（包含 Email、姓、名、機構單位路徑、Employee Title 欄位）\n' +
    '3. 您已經手動修改了需要更新的資料\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  if (values.length < 2) {
    ui.alert('錯誤', '工作表中沒有足夠的資料。請確保至少有標題行和一行資料。', ui.ButtonSet.OK);
    return;
  }

  var headers = values[0];
  var data = values.slice(1);

  // 查找各欄位的索引
  var emailCol = headers.indexOf('使用者 Email');
  var familyNameCol = headers.indexOf('姓 (Family Name)');
  var givenNameCol = headers.indexOf('名 (Given Name)');
  var orgUnitPathCol = headers.indexOf('機構單位路徑');
  var employeeTitleCol = headers.indexOf('Employee Title');
  var updateStatusCol = headers.indexOf('是否需要更新'); // 新增：檢測欄位的索引

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「使用者 Email」欄位。請確保工作表包含正確的標題。', ui.ButtonSet.OK);
    return;
  }

  if (familyNameCol === -1 && givenNameCol === -1 && orgUnitPathCol === -1 && employeeTitleCol === -1) {
    ui.alert('錯誤', '找不到任何可更新的欄位（姓、名、機構單位路徑、Employee Title）。請確保工作表包含至少其中一個欄位。', ui.ButtonSet.OK);
    return;
  }

  // 篩選出需要更新的行
  var rowsToUpdate = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = String(row[emailCol] || '').trim();
    var updateStatus = updateStatusCol !== -1 ? String(row[updateStatusCol] || '').trim() : '';

    // 如果有檢測欄位，只處理標記為「需要更新」的行；如果沒有檢測欄位，處理所有行
    if (email && (updateStatusCol === -1 || updateStatus === '需要更新')) {
      rowsToUpdate.push({
        index: i,
        rowNumber: i + 2, // 實際行號（包含標題行）
        data: row
      });
    }
  }

  if (rowsToUpdate.length === 0) {
    ui.alert('提示', '沒有找到需要更新的使用者。\n\n' +
      (updateStatusCol !== -1 ?
        '所有使用者的 I 欄都顯示「無需更新」，或沒有有效的 Email。' :
        '沒有找到有效的 Email。'),
      ui.ButtonSet.OK);
    return;
  }

  // 確認要處理的行數
  var confirmationFields = [];
  if (familyNameCol !== -1) confirmationFields.push('• 更新姓氏');
  if (givenNameCol !== -1) confirmationFields.push('• 更新名字');
  if (orgUnitPathCol !== -1) confirmationFields.push('• 更新機構單位路徑');
  if (employeeTitleCol !== -1) confirmationFields.push('• 更新職稱資訊');

  var finalConfirmation = ui.alert(
    '最終確認',
    '即將處理 ' + rowsToUpdate.length + ' 位使用者的資料' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的使用者）' : '') + '。\n\n' +
    '此操作將會：\n' +
    confirmationFields.join('\n') +
    '\n\n確定要執行嗎？',
    ui.ButtonSet.YES_NO
  );

  if (finalConfirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在更新使用者資料，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始更新使用者資料...'];
  var successCount = 0;
  var failCount = 0;
  var skipCount = 0;

  for (var i = 0; i < rowsToUpdate.length; i++) {
    var rowInfo = rowsToUpdate[i];
    var row = rowInfo.data;
    var email = String(row[emailCol] || '').trim();

    var logPrefix = '第 ' + rowInfo.rowNumber + ' 行 (' + email + '): ';

    try {
      // 檢查使用者是否存在
      var user;
      try {
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail,name,orgUnitPath,organizations" });
      } catch (e) {
        logMessages.push(logPrefix + '使用者不存在，跳過。');
        skipCount++;
        continue;
      }

      var needsUpdate = false;
      var userObj = {};

      // 處理姓名更新
      var nameObj = {};
      var nameUpdated = false;

      if (familyNameCol !== -1) {
        var newFamilyName = String(row[familyNameCol] || '').trim();
        var currentFamilyName = (user.name && user.name.familyName) ? user.name.familyName : '';

        if (newFamilyName && newFamilyName !== currentFamilyName) {
          nameObj.familyName = newFamilyName;
          nameUpdated = true;
          logMessages.push(logPrefix + '姓氏將從 "' + currentFamilyName + '" 更新為 "' + newFamilyName + '"');
        }
      }

      if (givenNameCol !== -1) {
        var newGivenName = String(row[givenNameCol] || '').trim();
        var currentGivenName = (user.name && user.name.givenName) ? user.name.givenName : '';

        if (newGivenName && newGivenName !== currentGivenName) {
          nameObj.givenName = newGivenName;
          nameUpdated = true;
          logMessages.push(logPrefix + '名字將從 "' + currentGivenName + '" 更新為 "' + newGivenName + '"');
        }
      }

      if (nameUpdated) {
        // 保留現有的姓名資料，只更新有變化的部分
        if (user.name) {
          if (!nameObj.familyName && user.name.familyName) {
            nameObj.familyName = user.name.familyName;
          }
          if (!nameObj.givenName && user.name.givenName) {
            nameObj.givenName = user.name.givenName;
          }
        }
        userObj.name = nameObj;
        needsUpdate = true;
      }

      // 處理機構單位路徑更新
      if (orgUnitPathCol !== -1) {
        var newOrgUnitPath = String(row[orgUnitPathCol] || '').trim();
        if (newOrgUnitPath && newOrgUnitPath !== user.orgUnitPath) {
          userObj.orgUnitPath = newOrgUnitPath;
          needsUpdate = true;
          logMessages.push(logPrefix + '機構單位路徑將從 "' + user.orgUnitPath + '" 更新為 "' + newOrgUnitPath + '"');
        }
      }

      // 處理職稱更新
      if (employeeTitleCol !== -1) {
        var newEmployeeTitle = String(row[employeeTitleCol] || '').trim();

        // 取得目前的職稱
        var currentTitle = '';
        if (user.organizations && user.organizations.length > 0) {
          for (var j = 0; j < user.organizations.length; j++) {
            if (user.organizations[j].title) {
              currentTitle = user.organizations[j].title;
              break;
            }
          }
        }

        // 比較職稱是否需要更新
        if (newEmployeeTitle !== currentTitle) {
          // 準備 organizations 資料結構
          if (newEmployeeTitle) {
            userObj.organizations = [{
              title: newEmployeeTitle,
              primary: true,
              type: 'work'
            }];
          } else {
            // 如果新職稱為空，清除職稱
            userObj.organizations = [];
          }
          needsUpdate = true;
          logMessages.push(logPrefix + '職稱將從 "' + currentTitle + '" 更新為 "' + newEmployeeTitle + '"');
        }
      }

      // 執行更新
      if (needsUpdate) {
        AdminDirectory.Users.update(userObj, email);
        logMessages.push(logPrefix + '使用者資料已成功更新。');
        successCount++;

        // 更新工作表中的檢測欄位狀態為「已更新」
        if (updateStatusCol !== -1) {
          sheet.getRange(rowInfo.rowNumber, updateStatusCol + 1).setValue('已更新');
        }
      } else {
        logMessages.push(logPrefix + '實際檢查後無需更新，資料相同。');
        skipCount++;
      }

      // 避免 API 速率限制
      if (i % 10 === 9) {
        Utilities.sleep(100);
      }

    } catch (e) {
      logMessages.push(logPrefix + '更新時發生錯誤: ' + e.message);
      failCount++;
    }
  }

  var resultMsg = '使用者資料更新完成！\n\n' +
    '處理了 ' + rowsToUpdate.length + ' 位使用者' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的使用者）' : '') + '：\n' +
    '成功更新: ' + successCount + ' 位使用者\n' +
    '跳過/無需更新: ' + skipCount + ' 位使用者\n' +
    '失敗/錯誤: ' + failCount + ' 位使用者\n\n' +
    '詳細日誌請查看 Apps Script 編輯器中的「執行作業」。\n\n' +
    '--- 部分日誌預覽 ---\n' +
    logMessages.slice(0, 15).join('\n') +
    (logMessages.length > 15 ? '\n...(更多日誌省略)' : '');

  ui.alert('更新結果', resultMsg, ui.ButtonSet.OK);
  Logger.log('--- 完整更新日誌 ---\n' + logMessages.join('\n'));

  // 關閉處理中提示
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>更新完成！</b>').setTitle('進度'));
}

/**
 * 根據工作表中的資料更新使用者所屬的群組。
 * 讀取目前工作表中的「所屬群組 (Groups)」欄位，並更新使用者實際所屬的群組。
 * 只處理 I 欄標記為「需要更新」的行。
 * 自動跳過巢狀群組（Nested Group）。
 */
function updateGroupMembersFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新群組成員歸屬',
    '此功能將讀取目前工作表的「所屬群組 (Groups)」欄位資料，並更新使用者實際所屬的群組。\n\n' +
    '★ 智能更新：只會處理 I 欄標記為「需要更新」的使用者。\n' +
    '★ 自動跳過：巢狀群組（Nested Group）不會被處理。\n\n' +
    '請確認：\n' +
    '1. 目前工作表是群組成員匯出的工作表\n' +
    '2. 您已經手動修改了「所屬群組 (Groups)」欄位\n' +
    '3. 群組名稱格式正確（用逗號分隔多個群組）\n\n' +
    '⚠️ 注意：此操作會完全替換使用者的群組歸屬！\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  if (values.length < 2) {
    ui.alert('錯誤', '工作表中沒有足夠的資料。請確保至少有標題行和一行資料。', ui.ButtonSet.OK);
    return;
  }

  var headers = values[0];
  var data = values.slice(1);

  // 查找各欄位的索引
  var emailCol = headers.indexOf('成員 Email');
  if (emailCol === -1) {
    // 如果找不到純文字版本，嘗試尋找帶有emoji的版本
    emailCol = headers.indexOf('🔒 成員 Email');
  }
  
  var typeCol = headers.indexOf('類型 (Type)');
  if (typeCol === -1) {
    typeCol = headers.indexOf('🔒 類型 (Type)');
  }
  
  var groupsCol = headers.indexOf('所屬群組 (Groups)');
  if (groupsCol === -1) {
    groupsCol = headers.indexOf('✅ 所屬群組 (Groups)');
  }
  
  var updateStatusCol = headers.indexOf('是否需要更新');
  if (updateStatusCol === -1) {
    updateStatusCol = headers.indexOf('🔒 是否需要更新');
  }

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「成員 Email」或「🔒 成員 Email」欄位。請確保工作表是從群組成員匯出功能產生的。', ui.ButtonSet.OK);
    return;
  }

  if (groupsCol === -1) {
    ui.alert('錯誤', '找不到「所屬群組 (Groups)」或「✅ 所屬群組 (Groups)」欄位。請確保工作表包含群組資訊。', ui.ButtonSet.OK);
    return;
  }

  // 篩選出需要更新的行（排除巢狀群組）
  var rowsToUpdate = [];
  var nestedGroupCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = String(row[emailCol] || '').trim();
    var type = typeCol !== -1 ? String(row[typeCol] || '').trim() : '';
    var updateStatus = updateStatusCol !== -1 ? String(row[updateStatusCol] || '').trim() : '';

    // 檢查是否為巢狀群組
    if (type === 'GROUP') {
      nestedGroupCount++;
      continue; // 跳過巢狀群組
    }

    // 如果有檢測欄位，只處理標記為「需要更新」的行；如果沒有檢測欄位，處理所有行
    if (email && (updateStatusCol === -1 || updateStatus === '需要更新')) {
      rowsToUpdate.push({
        index: i,
        rowNumber: i + 2, // 實際行號（包含標題行）
        data: row
      });
    }
  }

  if (rowsToUpdate.length === 0) {
    var noUpdateMsg = '沒有找到需要更新的使用者。\n\n';
    if (nestedGroupCount > 0) {
      noUpdateMsg += '• 已自動跳過 ' + nestedGroupCount + ' 個巢狀群組\n';
    }
    noUpdateMsg += (updateStatusCol !== -1 ?
      '• 所有使用者的 I 欄都顯示「無需更新」，或沒有有效的 Email。' :
      '沒有找到有效的 Email。');
    
    ui.alert('提示', noUpdateMsg, ui.ButtonSet.OK);
    return;
  }

  // 最終確認
  var confirmationMsg = '即將處理 ' + rowsToUpdate.length + ' 位使用者的群組歸屬' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的使用者）' : '') + '。\n\n';
  
  if (nestedGroupCount > 0) {
    confirmationMsg += '✓ 已自動跳過 ' + nestedGroupCount + ' 個巢狀群組。\n\n';
  }
  
  confirmationMsg += '⚠️ 重要提醒：\n' +
    '• 此操作會移除使用者原有的所有群組\n' +
    '• 然後將使用者加入到新指定的群組中\n' +
    '• 空白的群組欄位將使使用者不屬於任何群組\n\n' +
    '確定要執行嗎？';

  var finalConfirmation = ui.alert('最終確認', confirmationMsg, ui.ButtonSet.YES_NO);

  if (finalConfirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在更新群組成員歸屬，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始更新群組成員歸屬...'];
  if (nestedGroupCount > 0) {
    logMessages.push('已自動跳過 ' + nestedGroupCount + ' 個巢狀群組（類型為 GROUP）。');
  }
  
  var successCount = 0;
  var failCount = 0;
  var skipCount = 0;
  
  // 建立群組名稱到群組Email的對應表
  var groupNameToEmailMap = {};
  try {
    var allGroups = listAllGroups();
    for (var g = 0; g < allGroups.length; g++) {
      if (!allGroups[g].error) {
        groupNameToEmailMap[allGroups[g].name] = allGroups[g].email;
      }
    }
    logMessages.push('已建立群組名稱對應表，共 ' + Object.keys(groupNameToEmailMap).length + ' 個群組。');
  } catch (e) {
    logMessages.push('建立群組對應表時發生錯誤: ' + e.message);
  }

  for (var i = 0; i < rowsToUpdate.length; i++) {
    var rowInfo = rowsToUpdate[i];
    var row = rowInfo.data;
    var email = String(row[emailCol] || '').trim();
    var newGroupsText = String(row[groupsCol] || '').trim();

    var logPrefix = '第 ' + rowInfo.rowNumber + ' 行 (' + email + '): ';

    try {
      // 檢查使用者是否存在
      var user;
      try {
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail" });
      } catch (e) {
        logMessages.push(logPrefix + '使用者不存在，跳過。');
        skipCount++;
        continue;
      }

      // 解析新的群組列表
      var newGroups = [];
      if (newGroupsText && newGroupsText !== '無群組' && newGroupsText !== 'N/A' && newGroupsText !== '無法獲取' && newGroupsText !== '不適用') {
        var groupNames = newGroupsText.split(',').map(function(name) { return name.trim(); });
        
        for (var j = 0; j < groupNames.length; j++) {
          var groupName = groupNames[j];
          if (groupName && groupNameToEmailMap[groupName]) {
            newGroups.push({
              name: groupName,
              email: groupNameToEmailMap[groupName]
            });
          } else if (groupName) {
            logMessages.push(logPrefix + '警告：找不到群組 "' + groupName + '" 的 Email，將跳過此群組。');
          }
        }
      }

      // 步驟 1: 獲取使用者目前所屬的所有群組
      var currentGroups = [];
      try {
        var groupPageToken;
        do {
          var groupPage = AdminDirectory.Groups.list({
            userKey: email,
            maxResults: 200,
            pageToken: groupPageToken,
            fields: 'nextPageToken,groups(name,email)'
          });
          if (groupPage.groups) {
            currentGroups = currentGroups.concat(groupPage.groups);
          }
          groupPageToken = groupPage.nextPageToken;
               } while (groupPageToken);
      } catch (e) {
        logMessages.push(logPrefix + '無法獲取目前群組歸屬: ' + e.message);
      }

      logMessages.push(logPrefix + '目前屬於 ' + currentGroups.length + ' 個群組，將更新為 ' + newGroups.length + ' 個群組。');

      // 步驟 2: 從所有目前群組中移除該使用者
      var removeCount = 0;
      var removeErrors = 0;
      for (var k = 0; k < currentGroups.length; k++) {
        try {
          AdminDirectory.Members.remove(currentGroups[k].email, email);
          removeCount++;
        } catch (removeError) {
          removeErrors++;
          logMessages.push(logPrefix + '從群組 "' + currentGroups[k].name + '" 移除時失敗: ' + removeError.message);
        }
      }

      if (removeCount > 0) {
        logMessages.push(logPrefix + '成功從 ' + removeCount + ' 個群組中移除' + (removeErrors > 0 ? '（失敗 ' + removeErrors + ' 個）' : '') + '。');
      }

      // 步驟 3: 將使用者加入到新的群組中
      var addCount = 0;
      var addErrors = 0;
      for (var k = 0; k < newGroups.length; k++) {
        try {
          AdminDirectory.Members.insert({
            email: email,
            role: "MEMBER"
          }, newGroups[k].email);
          addCount++;
        } catch (addError) {
          if (addError.message.includes("Member already exists") || addError.message.includes("duplicate")) {
            logMessages.push(logPrefix + '已是群組 "' + newGroups[k].name + '" 的成員。');
            addCount++; // 視為成功
          } else {
            addErrors++;
            logMessages.push(logPrefix + '加入群組 "' + newGroups[k].name + '" 時失敗: ' + addError.message);
          }
        }
      }

      if (newGroups.length > 0) {
        logMessages.push(logPrefix + '成功加入 ' + addCount + ' 個群組' + (addErrors > 0 ? '（失敗 ' + addErrors + ' 個）' : '') + '。');
      } else {
        logMessages.push(logPrefix + '群組欄位為空，使用者現在不屬於任何群組。');
      }

      successCount++;

      // 更新工作表中的檢測欄位狀態為「已更新」
      if (updateStatusCol !== -1) {
        sheet.getRange(rowInfo.rowNumber, updateStatusCol + 1).setValue('已更新');
     
      }

      // 避免 API 速率限制
      if (i % 5 === 4) {
        Utilities.sleep(200);
      }

    } catch (e) {
      logMessages.push(logPrefix + '處理時發生嚴重錯誤: ' + e.message);
      failCount++;
    }
  }

  var resultMsg = '群組成員歸屬更新完成！\n\n' +
    '處理了 ' + rowsToUpdate.length + ' 位使用者' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的使用者）' : '') + '：\n' +
    '成功更新: ' + successCount + ' 位使用者\n' +
    '跳過/不存在: ' + skipCount + ' 位使用者\n' +
    '失敗/錯誤: ' + failCount + ' 位使用者\n' +
    (nestedGroupCount > 0 ? '自動跳過巢狀群組: ' + nestedGroupCount + ' 個\n' : '') +
    '\n詳細日誌請查看 Apps Script 編輯器中的「執行作業」。\n\n' +
    '--- 部分日誌預覽 ---\n' +
    logMessages.slice(0, 15).join('\n') +
    (logMessages.length > 15 ? '\n...(更多日誌省略)' : '');

  ui.alert('更新結果', resultMsg, ui.ButtonSet.OK);
  Logger.log('--- 完整群組更新日誌 ---\n' + logMessages.join('\n'));

  // 關閉處理中提示
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>群組更新完成！</b>').setTitle('進度'));
}

