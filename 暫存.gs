/**
 * 在試算表菜單中添加一個自定義菜單項。
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('管理帳號與群組')
    .addItem('1. 匯出[新建範本tea]sheet 範本', 'exportNewUserTemplate')
    .addItem('2. 依[新建範本tea]批次新增(不更動現有資料)', 'processUsersAndGroups_V2')
    .addSeparator()
    .addItem('1.匯出群組成員 (互動式)', 'showGroupManagementSidebar')
    .addItem('2.依據匯出的sheet更新群組成員', 'updateGroupMembersFromSheet')
    .addSeparator()
    .addItem('匯出所有機構單位 (含人數)', 'exportOUsAndUserCounts')
    .addToUi();
}

/**
 * [純新增版] 處理試算表中的使用者資料，僅新增不存在的帳號並支援加入多個指定群組。
 * 已存在的使用者將完全跳過，不做任何更動。
 */
function processUsersAndGroups_V2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var headers = values[0];
  var data = values.slice(1);

  // 查找各欄位的索引 - 更新為新的欄位名稱
  var firstNameCol = headers.indexOf('名');
  var lastNameCol = headers.indexOf('姓');
  var emailCol = headers.indexOf('Email'); // 更新欄位名稱
  var passwordCol = headers.indexOf('密碼'); // 更新欄位名稱
  var orgUnitPathCol = headers.indexOf('機構路徑');
  var employeeTitleCol = headers.indexOf('Employee Title(部別領域)'); // 更新欄位名稱
  var groupEmailCol = headers.indexOf('所屬群組');

  // 支援舊版欄位名稱的向後相容性
  if (emailCol === -1) {
    emailCol = headers.indexOf('Email Address [Required]'); // 舊版欄位名稱
  }
  if (passwordCol === -1) {
    passwordCol = headers.indexOf('空白代表不改密碼'); // 舊版欄位名稱
  }
  if (employeeTitleCol === -1) {
    employeeTitleCol = headers.indexOf('Employee Title'); // 舊版欄位名稱
  }
  if (groupEmailCol === -1) {
    groupEmailCol = headers.indexOf('加入群組'); // 支援舊版欄位名稱
  }

  // 檢查必要欄位是否存在
  if ([firstNameCol, lastNameCol, emailCol, orgUnitPathCol].includes(-1)) {
    var missingFields = [];
    if (firstNameCol === -1) missingFields.push('名');
    if (lastNameCol === -1) missingFields.push('姓');
    if (emailCol === -1) missingFields.push('Email (或 Email Address [Required])');
    if (orgUnitPathCol === -1) missingFields.push('機構路徑');
    
    SpreadsheetApp.getUi().alert('錯誤', 
      '試算表標題欄位不正確，缺少以下必要欄位：\n• ' + missingFields.join('\n• ') + 
      '\n\n請確保工作表包含這些欄位。', 
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ui = SpreadsheetApp.getUi();
  
  // 確認對話框明確說明只會新增
  var confirmation = ui.alert(
    '批次新增帳號確認',
    '此功能將【僅新增】不存在的使用者帳號。\n\n' +
    '★ 重要說明：\n' +
    '• 已存在的使用者將完全跳過，不做任何更動\n' +
    '• 只會新增不存在的帳號並設定群組\n' +
    '• 新增的帳號會要求首次登入時更改密碼\n' +
    '• 處理結果會顯示在最後一欄\n\n' +
    '確定要繼續執行純新增操作嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  // 在最後一欄加上處理結果標題（如果還沒有的話）
  var resultColIndex = headers.length;
  if (headers[resultColIndex - 1] !== '處理結果') {
    sheet.getRange(1, resultColIndex + 1).setValue('處理結果');
    sheet.getRange(1, resultColIndex + 1).setBackground('#4285F4');
    sheet.getRange(1, resultColIndex + 1).setFontColor('white');
    sheet.getRange(1, resultColIndex + 1).setFontWeight('bold');
    sheet.getRange(1, resultColIndex + 1).setHorizontalAlignment('center');
  } else {
    resultColIndex = resultColIndex - 1; // 如果已存在，使用現有的欄位
  }

  var successCount = 0;
  var failCount = 0;
  var skipCount = 0; // 跳過的已存在使用者數量
  var noActionCount = 0; // 群組操作中無需操作的數量
  var logMessages = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = String(row[emailCol] || '').trim();
    var resultMessage = '';
    var resultColor = '#FFFFFF'; // 預設白色背景
    
    if (!email) {
      sheet.getRange(i + 2, resultColIndex + 1).setValue('空白Email，跳過');
      sheet.getRange(i + 2, resultColIndex + 1).setBackground('#FFE0B2'); // 淺橘色
      continue; // 如果 Email 為空，直接跳過此行
    }

    var logPrefix = '第 ' + (i + 2) + ' 行 (' + email + '): ';

    try {
      // 【重要修改】先檢查使用者是否已存在，再檢查必填欄位
      var userExists = false;
      try {
        AdminDirectory.Users.get(email, { fields: "primaryEmail" });
        userExists = true;
      } catch (e) {
        userExists = false;
      }

      if (userExists) {
        // 使用者已存在，完全跳過
        resultMessage = '使用者已存在，跳過';
        resultColor = '#FFF3E0'; // 淺橘色
        logMessages.push(logPrefix + '使用者已存在，跳過處理（不做任何更動）。');
        skipCount++;
        sheet.getRange(i + 2, resultColIndex + 1).setValue(resultMessage);
        sheet.getRange(i + 2, resultColIndex + 1).setBackground(resultColor);
        continue;
      }

      // 使用者不存在，才檢查必填欄位
      var firstName = String(row[firstNameCol] || '').trim();
      var lastName = String(row[lastNameCol] || '').trim();
      var password = String(row[passwordCol] || '').trim();
      var orgUnitPath = String(row[orgUnitPathCol] || '').trim();
      var employeeTitle = (employeeTitleCol !== -1) ? String(row[employeeTitleCol] || '').trim() : '';
      var groupEmails = (groupEmailCol !== -1) ? String(row[groupEmailCol] || '').trim() : '';

      if (!firstName || !lastName || !orgUnitPath || !password) {
        resultMessage = '必填欄位不完整';
        resultColor = '#FFCDD2'; // 淺紅色
        logMessages.push(logPrefix + '錯誤 - 必要的欄位 (名, 姓, 機構路徑, 密碼) 不完整，跳過。');
        failCount++;
        sheet.getRange(i + 2, resultColIndex + 1).setValue(resultMessage);
        sheet.getRange(i + 2, resultColIndex + 1).setBackground(resultColor);
        continue;
      }

      // 執行新增使用者
      var userObj = {
        name: { givenName: firstName, familyName: lastName },
        orgUnitPath: orgUnitPath,
        primaryEmail: email,
        password: password,
        changePasswordAtNextLogin: true
      };

      // 如果 employeeTitle 有值才加入
      if (employeeTitle) {
        userObj.organizations = [{
          title: employeeTitle,
          primary: true,
          type: 'work'
        }];
      }

      AdminDirectory.Users.insert(userObj);
      logMessages.push(logPrefix + '使用者帳號已成功創建。');
      
      var groupResults = [];
      var hasGroupError = false;

      // 處理群組加入（僅對新建立的使用者）
      if (groupEmails) {
        var groups = groupEmails.split(',').map(function (g) { return g.trim(); });
        for (var j = 0; j < groups.length; j++) {
          var groupEmail = groups[j];
          if (groupEmail) {
            try {
              AdminDirectory.Members.insert({ email: email, role: "MEMBER" }, groupEmail);
              logMessages.push(logPrefix + '已成功加入群組 ' + groupEmail + '。');
              groupResults.push('✓' + groupEmail);
            } catch (groupError) {
              // 檢查是否為"成員已存在"的錯誤
              if (groupError.message.includes("Member already exists") || groupError.message.includes("duplicate")) {
                logMessages.push(logPrefix + '已是群組 ' + groupEmail + ' 的成員，無需操作。');
                groupResults.push('○' + groupEmail);
                noActionCount++;
              } else {
                // 其他群組相關錯誤
                logMessages.push(logPrefix + '加入群組 ' + groupEmail + ' 時失敗: ' + groupError.message);
                groupResults.push('✗' + groupEmail);
                hasGroupError = true;
              }
            }
          }
        }
      }

      // 設定處理結果訊息
      if (hasGroupError) {
        resultMessage = '帳號已新增，部分群組失敗';
        resultColor = '#FFECB3'; // 淺黃色
      } else {
        resultMessage = '帳號已新增成功';
        if (groupResults.length > 0) {
          resultMessage += ' (群組: ' + groupResults.length + ')';
        }
        resultColor = '#C8E6C9'; // 淺綠色
      }

      successCount++;

    } catch (e) {
      resultMessage = '處理失敗: ' + e.message;
      resultColor = '#FFCDD2'; // 淺紅色
      logMessages.push(logPrefix + '處理帳號時發生嚴重錯誤: ' + e.message);
      failCount++;
    }

    // 寫入處理結果到工作表
    sheet.getRange(i + 2, resultColIndex + 1).setValue(resultMessage);
    sheet.getRange(i + 2, resultColIndex + 1).setBackground(resultColor);

    // 避免 API 速率限制
    if (i % 10 === 9) {
      Utilities.sleep(200);
    }
  }

  var resultMsg = '批次新增帳號處理完成！\n\n' +
    '成功新增帳號數: ' + successCount + '\n' +
    '跳過已存在帳號數: ' + skipCount + '\n' +
    '失敗/錯誤數: ' + failCount + '\n' +
    '群組無需操作數: ' + noActionCount + '\n\n' +
    '處理結果已顯示在工作表最後一欄。\n' +
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
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('批次新增帳號、群組@tea_Sidebar')
        .setWidth(400)
        .setTitle('群組管理工具');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('錯誤', 
      '無法載入側邊欄：' + e.message + '\n\n' +
      '請確保 HTML 檔案 "批次新增帳號、群組@tea_Sidebar.html" 存在。', 
      ui.ButtonSet.OK);
    Logger.log('側邊欄載入錯誤: ' + e.toString());
  }
}

/**
 * 為側邊欄提供群組成員匯出功能
 * @param {string} groupEmail 群組Email
 * @returns {object} 包含匯出結果的物件
 */
function getGroupMembersForSidebar(groupEmail) {
  var logMessages = ['側邊欄匯出群組成員: ' + groupEmail];
  
  try {
    // 獲取群組資訊
    var groupInfo;
    try {
      groupInfo = AdminDirectory.Groups.get(groupEmail);
    } catch (e) {
      return { 
        success: false, 
        message: '無法找到群組 ' + groupEmail + '。請檢查群組Email是否正確。' 
      };
    }

    var allMembers = [];
    var pageToken;
    
    // 分頁獲取所有成員
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
      return { 
        success: true, 
        noMembers: true,
        message: '群組 ' + groupEmail + ' 中沒有任何成員。' 
      };
    }

    logMessages.push('找到 ' + allMembers.length + ' 位成員，開始建立工作表...');

    // 準備輸出資料
    var outputData = [['成員 Email', '角色 (Role)', '類型 (Type)', '狀態 (Status)']];
    
    for (var i = 0; i < allMembers.length; i++) {
      var member = allMembers[i];
      outputData.push([
        member.email || '',
        member.role || '',
        member.type || '',
        member.status || ''
      ]);
    }

    // 建立新工作表
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "[群組成員] " + (groupInfo.name || groupEmail.split('@')[0]);
    
    // 如果同名工作表存在，加上時間戳記
    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      var timestamp = new Date().toISOString().slice(11, 19).replace(/:/g, '');
      sheetName = sheetName + "_" + timestamp;
    }

    var newSheet = spreadsheet.insertSheet(sheetName);

    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 設定格式
    newSheet.setColumnWidth(1, 250); // 成員 Email
    newSheet.setColumnWidth(2, 100); // 角色
    newSheet.setColumnWidth(3, 100); // 類型
    newSheet.setColumnWidth(4, 100); // 狀態

    // 標題行格式
    var headerRange = newSheet.getRange(1, 1, 1, 4);
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // 資料格式
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 4);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
    }

    // 凍結標題行
    newSheet.setFrozenRows(1);

    // 切換到新工作表
    newSheet.activate();

    logMessages.push('匯出完成，共 ' + allMembers.length + ' 位成員');
    Logger.log('側邊欄群組成員匯出日誌:\n' + logMessages.join('\n'));

    return { 
      success: true, 
      memberCount: allMembers.length,
      sheetName: sheetName,
      groupName: groupInfo.name || groupEmail
    };

  } catch (e) {
    var errorMsg = '匯出過程中發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    Logger.log('側邊欄匯出錯誤:\n' + logMessages.join('\n'));
    
    return { 
      success: false, 
      message: errorMsg
    };
  }
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
        maxResults: 200,
        pageToken: pageToken,
        fields: 'nextPageToken,groups(name,email,directMembersCount)'
      });
      
      if (page.groups) {
        var groups = page.groups.map(function (group) {
          return {
            name: group.name || '未命名群組',
            email: group.email || '',
            memberCount: group.directMembersCount || 0
          };
        });
        allGroups = allGroups.concat(groups);
      }
      
      pageToken = page.nextPageToken;
      
      // 在分頁請求間稍作暫停
      if (pageToken) {
        Utilities.sleep(100);
      }
      
    } while (pageToken);

    // 按群組名稱排序
    allGroups.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    Logger.log('成功獲取 ' + allGroups.length + ' 個群組');
    return allGroups;
    
  } catch (e) {
    Logger.log('獲取群組列表時發生錯誤: ' + e.toString());
    return [{ error: '無法獲取群組列表: ' + e.message }];
  }
}

/**
 * 根據工作表中的資料批次更新群組成員
 * 工作表應包含 '群組Email' 和 '成員Email' 欄位
 */
function updateGroupMembersFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // 確認對話框
  var confirmation = ui.alert(
    '批次更新群組成員',
    '此功能將根據目前工作表的資料批次更新群組成員。\n\n' +
    '工作表應包含以下欄位：\n' +
    '• 群組Email 或 Group Email\n' +
    '• 成員Email 或 Member Email\n\n' +
    '操作類型：\n' +
    '• 如果成員不在群組中，將會被加入\n' +
    '• 如果成員已在群組中，將保持不變\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  if (values.length < 2) {
    ui.alert('錯誤', '工作表中沒有足夠的資料。請確保至少有標題行和一行資料。', ui.ButtonSet.OK);
    return;
  }

  var headers = values[0];
  var data = values.slice(1);

  // 尋找必要的欄位
  var groupEmailCol = headers.indexOf('群組Email');
  if (groupEmailCol === -1) {
    groupEmailCol = headers.indexOf('Group Email');
  }
  
  var memberEmailCol = headers.indexOf('成員Email');
  if (memberEmailCol === -1) {
    memberEmailCol = headers.indexOf('Member Email');
  }

  if (groupEmailCol === -1 || memberEmailCol === -1) {
    ui.alert('錯誤', 
      '找不到必要的欄位。請確保工作表包含以下欄位之一：\n' +
      '• 群組Email 或 Group Email\n' +
      '• 成員Email 或 Member Email', 
      ui.ButtonSet.OK);
    return;
  }

  // 在最後一欄加上處理結果標題
  var resultColIndex = headers.length;
  if (headers[resultColIndex - 1] !== '處理結果') {
    sheet.getRange(1, resultColIndex + 1).setValue('處理結果');
    sheet.getRange(1, resultColIndex + 1).setBackground('#4285F4');
    sheet.getRange(1, resultColIndex + 1).setFontColor('white');
    sheet.getRange(1, resultColIndex + 1).setFontWeight('bold');
    sheet.getRange(1, resultColIndex + 1).setHorizontalAlignment('center');
  } else {
    resultColIndex = resultColIndex - 1;
  }

  var successCount = 0;
  var failCount = 0;
  var skipCount = 0;
  var logMessages = [];

  // 處理每一行資料
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var groupEmail = String(row[groupEmailCol] || '').trim();
    var memberEmail = String(row[memberEmailCol] || '').trim();
    var resultMessage = '';
    var resultColor = '#FFFFFF';

    if (!groupEmail || !memberEmail) {
      resultMessage = '群組或成員Email為空';
      resultColor = '#FFE0B2'; // 淺橘色
      sheet.getRange(i + 2, resultColIndex + 1).setValue(resultMessage);
      sheet.getRange(i + 2, resultColIndex + 1).setBackground(resultColor);
      skipCount++;
      continue;
    }

    var logPrefix = '第 ' + (i + 2) + ' 行 (' + memberEmail + ' → ' + groupEmail + '): ';

    try {
      // 檢查成員是否已在群組中
      var memberExists = false;
      try {
        AdminDirectory.Members.get(groupEmail, memberEmail);
        memberExists = true;
      } catch (e) {
        memberExists = false;
      }

      if (memberExists) {
        resultMessage = '成員已存在群組中';
        resultColor = '#FFF3E0'; // 淺橘色
        logMessages.push(logPrefix + '成員已在群組中，無需操作。');
        skipCount++;
      } else {
        // 加入成員到群組
        AdminDirectory.Members.insert({
          email: memberEmail,
          role: 'MEMBER'
        }, groupEmail);
        
        resultMessage = '成功加入群組';
        resultColor = '#C8E6C9'; // 淺綠色
        logMessages.push(logPrefix + '成功加入群組。');
        successCount++;
      }

    } catch (e) {
      resultMessage = '處理失敗: ' + e.message;
      resultColor = '#FFCDD2'; // 淺紅色
      logMessages.push(logPrefix + '處理失敗: ' + e.message);
      failCount++;
    }

    // 寫入處理結果
    sheet.getRange(i + 2, resultColIndex + 1).setValue(resultMessage);
    sheet.getRange(i + 2, resultColIndex + 1).setBackground(resultColor);

    // 每處理10行暫停一下
    if (i % 10 === 9) {
      Utilities.sleep(200);
    }
  }

  // 顯示完成訊息
  var resultMsg = '批次更新群組成員處理完成！\n\n' +
    '成功加入群組數: ' + successCount + '\n' +
    '跳過已存在成員數: ' + skipCount + '\n' +
    '失敗/錯誤數: ' + failCount + '\n\n' +
    '處理結果已顯示在工作表最後一欄。\n' +
    '詳細日誌請查看 Apps Script 編輯器中的「執行作業」。\n\n' +
    '--- 部分日誌預覽 ---\n' + logMessages.slice(0, 15).join('\n') + 
    (logMessages.length > 15 ? '\n...(更多日誌省略)' : '');

  ui.alert('處理結果', resultMsg, ui.ButtonSet.OK);
  Logger.log('--- 完整執行日誌 ---\n' + logMessages.join('\n'));
}

/**
 * 匯出所有機構單位及其人數統計
 */
function exportOUsAndUserCounts() {
  var ui = SpreadsheetApp.getUi();
  
  var confirmation = ui.alert(
    '匯出機構單位',
    '此功能將匯出所有機構單位及其使用者人數統計。\n\n' +
    '這可能需要較長時間處理，確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  try {
    ui.alert('處理中', '正在獲取機構單位資料，請稍候...', ui.ButtonSet.OK);

    // 獲取所有機構單位
    var allOUs = [];
    var pageToken;
    
    do {
      var page = AdminDirectory.Orgunits.list('my_customer', {
        maxResults: 500,
        pageToken: pageToken
      });
      
      if (page.organizationUnits) {
        allOUs = allOUs.concat(page.organizationUnits);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    // 準備輸出資料
    var outputData = [['機構路徑', '機構名稱', '描述', '使用者人數']];
    
    for (var i = 0; i < allOUs.length; i++) {
      var ou = allOUs[i];
      
      // 計算該機構單位的使用者人數
      var userCount = 0;
      try {
        var userPage = AdminDirectory.Users.list({
          customer: 'my_customer',
          query: 'orgUnitPath=' + ou.orgUnitPath,
          maxResults: 1
        });
        userCount = userPage.totalCount || 0;
      } catch (e) {
        userCount = 0;
      }
      
      outputData.push([
        ou.orgUnitPath || '',
        ou.name || '',
        ou.description || '',
        userCount
      ]);
    }

    // 建立新工作表
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    var sheetName = '[機構單位統計]_' + timestamp;
    
    var newSheet = spreadsheet.insertSheet(sheetName);
    
    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // 設定格式
    newSheet.setColumnWidth(1, 200); // 機構路徑
    newSheet.setColumnWidth(2, 150); // 機構名稱
    newSheet.setColumnWidth(3, 200); // 描述
    newSheet.setColumnWidth(4, 80);  // 使用者人數
    
    // 標題行格式
    var headerRange = newSheet.getRange(1, 1, 1, 4);
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 凍結標題行
    newSheet.setFrozenRows(1);
    
    newSheet.activate();
    
    ui.alert('匯出成功！', 
      '共匯出 ' + allOUs.length + ' 個機構單位的資料至工作表 "' + sheetName + '"。', 
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', '匯出過程中發生錯誤: ' + e.message, ui.ButtonSet.OK);
  }
}
