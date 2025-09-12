/**
 * 在試算表菜單中添加一個自定義菜單項。
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('管理帳號與群組')
      .addItem('依試算表資料批次處理', 'processUsersAndGroups_V2')
      .addSeparator()
      .addItem('查詢/匯出群組成員 (互動式)', 'showGroupManagementSidebar')
      .addItem('匯出所有機構單位 (含人數)', 'exportOUsAndUserCounts')
      .addSeparator()
      .addItem('1.匯出全部@tea 清單', 'exportAllUsers')
      .addItem('2.依據匯出sheet 只更新使用者機構單位與職稱', 'updateUsersFromSheet') // 【新增這一行】
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
        user = AdminDirectory.Users.get(email, {fields: "primaryEmail"}); // 優化：只獲取必要的欄位，API 調用更輕量
      } catch (e) {
        user = null;
      }

      var userObj = {
        name: { givenName: firstName, familyName: lastName },
        orgUnitPath: orgUnitPath,
        // 如果 employeeTitle 為空字串，API 可能會報錯，所以只有在有值時才加入
        ...(employeeTitle && {title: employeeTitle}) 
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
        var groups = groupEmails.split(',').map(function(g) { return g.trim(); });
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
    var sheetName = groupEmail.split('@')[0] + "_成員列表";
    var newSheet = spreadsheet.insertSheet(sheetName);

    // 將資料一次性寫入新工作表
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // 自動調整欄寬以利閱讀
    newSheet.autoResizeColumns(1, 4);

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
        var groups = page.groups.map(function(group) {
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
    
    allGroups.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    return allGroups;
  } catch (e) {
    Logger.log('無法獲取群組列表: ' + e.toString());
    return [{ error: '無法獲取群組列表: ' + e.message }];
  }
}


/**
 * [最終版] 根據給定的群組 Email，獲取其所有成員（包含姓名和最後登入時間），並直接匯出到一個新的工作表。
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

    // 步驟 2: 準備要寫入工作表的資料，並更新標題行以包含新欄位
    var outputData = [['成員 Email', '姓 (Family Name)', '名 (Given Name)', '最後登入時間 (Last Login)', '角色 (Role)', '類型 (Type)', '狀態 (Status)']]; // 【新增】更新標題行

    // 步驟 3: 遍歷每一位成員，以獲取他們的詳細資訊
    for (var i = 0; i < allMembers.length; i++) {
      var member = allMembers[i];
      var firstName = '';
      var lastName = '';
      var lastLogin = 'N/A'; // 預設值
      
      if (member.type === 'USER') {
        try {
          // 【更新】擴大 API 請求範圍，加入 lastLoginTime
          var user = AdminDirectory.Users.get(member.email, {
            fields: 'name,lastLoginTime' 
          });
          firstName = user.name.givenName || '';
          lastName = user.name.familyName || '';

          // 【新增】處理並格式化時間
          if (user.lastLoginTime) {
            // Google 回傳的是 Epoch time (since 1970-01-01 in milliseconds) or ISO string
            // 我們將它轉換為本地時區的可讀格式
            var loginDate = new Date(user.lastLoginTime);
            // 檢查日期是否有效 (有些帳號可能從未登入過，回傳的日期是 1970-01-01)
            if (loginDate.getFullYear() > 1970) {
              lastLogin = loginDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
            } else {
              lastLogin = '從未登入';
            }
          } else {
            lastLogin = '從未登入';
          }
        } catch (userError) {
          firstName = 'N/A';
          lastName = 'N/A';
          lastLogin = '無法獲取';
          Logger.log('無法獲取使用者 ' + member.email + ' 的詳細資訊: ' + userError.message);
        }
      } else {
        firstName = '(巢狀群組)';
        lastName = '(Nested Group)';
        lastLogin = '不適用';
      }

      // 將包含新欄位的完整資料加入到輸出陣列中
      outputData.push([member.email, lastName, firstName, lastLogin, member.role, member.type, member.status]);
    }

    // 步驟 4: 建立新的工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var safeSheetName = groupEmail.split('@')[0].replace(/[^\w\s]/gi, '_') + "_成員列表"; 
    
    var existingSheet = spreadsheet.getSheetByName(safeSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    var newSheet = spreadsheet.insertSheet(safeSheetName, 0);

    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    newSheet.autoResizeColumns(1, 7); // 【更新】欄數從 6 改為 7

    newSheet.activate();

    // 步驟 5: 回傳成功的結果給側邊欄
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

  var logMessages = ['開始掃描機構單位與使用者人數...'];

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
        page.users.forEach(function(user) {
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
    outputData.sort(function(a, b) {
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
    newSheet.autoResizeColumns(1, 3);
    newSheet.activate();

    ui.alert('匯出成功！', '包含 ' + (outputData.length -1) + ' 個機構單位的統計資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
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
    // 步驟 1: 獲取所有使用者 - 【修改】使用 organizations 欄位來取得 title
    var pageToken;
    do {
      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        // 【修改】加入 organizations 欄位來取得職稱資訊
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

    // 步驟 2: 準備要寫入工作表的資料 - 在機構單位路徑後加入 Employee Title 欄位
    var outputData = [[
      '使用者 Email',
      '姓 (Family Name)', 
      '名 (Given Name)',
      '機構單位路徑',
      'Employee Title',
      '帳號狀態',
      '建立時間',
      '最後登入時間'
    ]];

    // 步驟 3: 處理每位使用者的資料
    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];
      
      var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
      var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
      var orgUnitPath = user.orgUnitPath || '/';
      
      // 【新增】從 organizations 陣列中取得職稱資訊
      var employeeTitle = 'N/A';
      if (user.organizations && user.organizations.length > 0) {
        // 取得第一個組織的職稱，如果有多個組織，取主要的那個
        for (var j = 0; j < user.organizations.length; j++) {
          var org = user.organizations[j];
          if (org.title) {
            employeeTitle = org.title;
            break; // 找到第一個有職稱的組織就停止
          }
        }
      }
      
      var status = user.suspended ? '已停用' : '啟用中';
      
      // 格式化建立時間
      var creationTime = 'N/A';
      if (user.creationTime) {
        var createdDate = new Date(user.creationTime);
        creationTime = createdDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
      }
      
      // 格式化最後登入時間
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

      // 將 Employee Title 欄位加入到輸出資料中（在機構單位路徑後面）
      outputData.push([
        user.primaryEmail,
        familyName,
        givenName,
        orgUnitPath,
        employeeTitle,
        status,
        creationTime,
        lastLoginTime
      ]);
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '_');
    var sheetName = "所有使用者清單_" + timestamp;
    
    // 檢查是否有同名工作表並刪除
    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    var newSheet = spreadsheet.insertSheet(sheetName, 0);
    
    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // 自動調整欄寬
    newSheet.autoResizeColumns(1, outputData[0].length);
    
    // 凍結標題行
    newSheet.setFrozenRows(1);
    
    // 設定標題行格式
    var headerRange = newSheet.getRange(1, 1, 1, outputData[0].length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 切換到新工作表
    newSheet.activate();

    ui.alert('匯出成功！', 
             allUsers.length + ' 位使用者的基本資料已成功匯出至新的工作表 "' + sheetName + '"。\n\n' +
             '工作表包含使用者的基本資訊：Email、姓名、機構單位、職稱、狀態及登入時間。', 
             ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', 
             '無法完成使用者清單匯出。\n\n' +
             '可能的原因：\n' +
             '- API 權限不足\n' +
             '- 網路連線問題\n' +
             '- 資料量過大導致超時\n\n' +
             '錯誤詳情: ' + e.message, 
             ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉處理中提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}

/**
 * 根據試算表中的資料更新使用者的機構單位路徑和職稱。
 * 讀取目前工作表中的資料，並更新對應使用者的 orgUnitPath 和 Employee Title。
 */
function updateUsersFromSheet() {
  var ui = SpreadsheetApp.getUi();
  
  // 第一層確認
  var confirmation = ui.alert(
    '更新使用者資訊',
    '此功能將讀取目前工作表的資料，並更新使用者的機構單位路徑和職稱。\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的使用者資料\n' +
    '2. 資料格式正確（包含 Email、機構單位路徑、Employee Title 欄位）\n' +
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
  var orgUnitPathCol = headers.indexOf('機構單位路徑');
  var employeeTitleCol = headers.indexOf('Employee Title');

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「使用者 Email」欄位。請確保工作表包含正確的標題。', ui.ButtonSet.OK);
    return;
  }
  
  if (orgUnitPathCol === -1 && employeeTitleCol === -1) {
    ui.alert('錯誤', '找不到「機構單位路徑」或「Employee Title」欄位。請確保工作表包含至少其中一個欄位。', ui.ButtonSet.OK);
    return;
  }

  // 最後確認
  var finalConfirmation = ui.alert(
    '最終確認',
    '即將處理 ' + data.length + ' 位使用者的資料。\n\n' +
    '此操作將會：\n' +
    (orgUnitPathCol !== -1 ? '• 更新機構單位路徑\n' : '') +
    (employeeTitleCol !== -1 ? '• 更新職稱資訊\n' : '') +
    '\n確定要執行嗎？',
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

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = String(row[emailCol] || '').trim();
    
    if (!email) {
      skipCount++;
      continue; // 如果 Email 為空，直接跳過此行
    }

    var logPrefix = '第 ' + (i + 2) + ' 行 (' + email + '): ';
    
    try {
      // 檢查使用者是否存在
      var user;
      try {
        user = AdminDirectory.Users.get(email, {fields: "primaryEmail,orgUnitPath,organizations"});
      } catch (e) {
        logMessages.push(logPrefix + '使用者不存在，跳過。');
        skipCount++;
        continue;
      }

      var needsUpdate = false;
      var userObj = {};

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
      } else {
        logMessages.push(logPrefix + '無需更新，資料相同。');
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

