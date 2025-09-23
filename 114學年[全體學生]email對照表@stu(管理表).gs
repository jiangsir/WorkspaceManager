/**
 * 這個函數會在試算表檔案被開啟時自動執行，
 * 並在工具列上建立一個名為「管理工具」的自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('管理工具')
    .addItem('1.匯出[全部@stu清單]', 'exportAllStudents')
    .addItem('2.依據[全部@stu清單] 更新 B,C,D,E,F,G 欄位內容', 'updateStudentsFromSheet')
    .addSeparator()
    .addItem('匯出[預約停權]範本', 'exportSuspensionTemplate')
    .addItem('--1.依據"停權時間"啟動停權程序', 'scheduleCompleteSuspensionProcess')
    .addItem('--2.列出本工作表內所有觸發器', 'listAllTriggers')
    .addItem('--3.清理本工作表內所有觸發器', 'cleanAllSuspensionTriggers')
    .addToUi();
}

/**
 * 匯出整個 stu 網域中的所有使用者資料到一個新的工作表。
 * 針對學生版本優化，移除群組資訊以加速處理，並採用分批處理避免逾時。
 */
function exportAllStudents() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '匯出所有學生',
    '您即將匯出整個 stu 網域的所有學生清單。\n\n此操作可能需要較長時間，如果資料量很大會分批處理，確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  // 清除之前的進度記錄
  PropertiesService.getScriptProperties().deleteProperty('EXPORT_PROGRESS');
  
  // 開始新的匯出 - 直接處理，不保存大量資料
  performDirectExport();
}

/**
 * 執行直接匯出處理（不保存大量資料到 Properties）
 */
function performDirectExport() {
  var ui = SpreadsheetApp.getUi();
  var startTime = new Date().getTime();
  var maxExecutionTime = 4.5 * 60 * 1000; // 4.5分鐘限制，留更多緩衝
  
  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在處理學生資料，請稍候...</b>').setTitle('處理中'));
  
  var logMessages = [];
  var allUsers = [];
  var totalProcessed = 0;
  
  try {
    // 檢查是否有之前的進度（只保存基本資訊）
    var savedProgress = PropertiesService.getScriptProperties().getProperty('EXPORT_PROGRESS');
    var startPageToken = null;
    var skipCount = 0;
    
    if (savedProgress) {
      var progress = JSON.parse(savedProgress);
      startPageToken = progress.pageToken || null;
      skipCount = progress.processedCount || 0;
      logMessages.push('恢復進度，從第 ' + (skipCount + 1) + ' 位開始');
    } else {
      logMessages.push('開始新的匯出作業');
    }

    // 準備輸出資料的標題
    var outputData = [[
      'Email', '姓 (Family Name)', '名 (Given Name)', '機構單位路徑',
      'Employee ID(真實姓名)', 'Employee Title(部別領域)', 'Department(註解)',
      '帳號狀態', '建立時間', '最後登入時間', '是否需要更新', '在學狀態'
    ]];

    // 獲取和處理使用者資料（邊獲取邊處理）
    var pageToken = startPageToken;
    var currentSkipped = 0;
    
    do {
      // 檢查執行時間
      var currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        logMessages.push('時間接近限制，保存輕量進度');
        
        // 只保存基本進度資訊（不保存使用者資料）
        var lightProgress = {
          pageToken: pageToken,
          processedCount: totalProcessed,
          phase: 'FETCHING'
        };
        
        PropertiesService.getScriptProperties().setProperty('EXPORT_PROGRESS', JSON.stringify(lightProgress));
        ui.alert('處理中', '已處理 ' + totalProcessed + ' 位學生。\n\n請點選「2.繼續未完成的匯出」來繼續處理。', ui.ButtonSet.OK);
        return;
      }

      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 100, // 減少批次大小
        pageToken: pageToken,
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,organizations,externalIds,suspended,creationTime,lastLoginTime)'
      });

      if (page.users) {
        for (var i = 0; i < page.users.length; i++) {
          // 如果需要跳過已處理的資料
          if (currentSkipped < skipCount) {
            currentSkipped++;
            continue;
          }
          
          var user = page.users[i];
          var userData = processUserData(user);
          outputData.push(userData);
          totalProcessed++;
          
          if (totalProcessed % 100 === 0) {
            logMessages.push('已處理 ' + totalProcessed + ' 位學生的資料');
            
            // 定期檢查時間
            var checkTime = new Date().getTime();
            if (checkTime - startTime > maxExecutionTime) {
              logMessages.push('時間接近限制，保存進度並暫停');
              
              var lightProgress = {
                pageToken: pageToken,
                processedCount: totalProcessed,
                phase: 'PROCESSING'
              };
              
              PropertiesService.getScriptProperties().setProperty('EXPORT_PROGRESS', JSON.stringify(lightProgress));
              ui.alert('處理中', '已處理 ' + totalProcessed + ' 位學生。\n\n請點選「2.繼續未完成的匯出」來繼續處理。', ui.ButtonSet.OK);
              return;
            }
          }
        }
        
        logMessages.push('已獲取並處理 ' + totalProcessed + ' 位學生');
      }
      
      pageToken = page.nextPageToken;
    } while (pageToken);

    logMessages.push('資料獲取完成，共處理 ' + totalProcessed + ' 位學生，開始建立工作表');

    // 建立工作表
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "[全部@stu清單]";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 分批寫入資料
    var writeBatchSize = 1000;
    for (var startRow = 0; startRow < outputData.length; startRow += writeBatchSize) {
      // 檢查時間
      var currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        logMessages.push('寫入過程中時間接近限制');
        ui.alert('處理中', '資料處理完成，正在寫入工作表。\n\n如果沒有完成，請重新執行匯出。', ui.ButtonSet.OK);
        break;
      }

      var endRow = Math.min(startRow + writeBatchSize, outputData.length);
      var batchData = outputData.slice(startRow, endRow);
      
      newSheet.getRange(startRow + 1, 1, batchData.length, batchData[0].length).setValues(batchData);
      logMessages.push('已寫入第 ' + (startRow + 1) + ' 到第 ' + endRow + ' 行');
    }

    // 簡化的格式設定
    setupSimpleFormatting(newSheet, outputData.length);
    
    newSheet.activate();
    
    // 清除進度記錄
    PropertiesService.getScriptProperties().deleteProperty('EXPORT_PROGRESS');
    
    ui.alert('匯出成功！', totalProcessed + ' 位學生的資料已成功匯出至工作表 "' + sheetName + '"。', ui.ButtonSet.OK);
    logMessages.push('匯出完成');

  } catch (e) {
    logMessages.push('發生錯誤: ' + e.message);
    ui.alert('錯誤', '匯出過程中發生錯誤：\n\n' + e.message + '\n\n可嘗試重新執行匯出。', ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>處理完成或已暫停</b>').setTitle('狀態'));
  }
}

/**
 * 根據試算表中的資料更新學生的姓名、機構單位路徑和職稱。
 * 讀取目前工作表中的資料，並更新對應學生的 姓名、機構單位路徑、Employee ID、Employee Title、Department。
 * 只處理 K 欄標記為「需要更新」的行。
 */
function updateStudentsFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新學生資訊',
    '此功能將讀取目前工作表的資料，並更新學生的姓名、機構單位路徑、員工編號、職稱和部門。\n\n' +
    '★ 智能更新：只會處理 K 欄標記為「需要更新」的學生。\n' +
    '★ 可更新欄位：B(姓)、C(名)、D(機構單位)、E(員工編號)、F(職稱)、G(部門)\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的學生資料\n' +
    '2. 資料格式正確\n' +
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
  var emailCol = headers.indexOf('Email');                        // A欄
  var familyNameCol = headers.indexOf('姓 (Family Name)');        // B欄
  var givenNameCol = headers.indexOf('名 (Given Name)');          // C欄
  var orgUnitPathCol = headers.indexOf('機構單位路徑');            // D欄
  var employeeIdCol = headers.indexOf('Employee ID(真實姓名)');   // E欄
  var employeeTitleCol = headers.indexOf('Employee Title(部別領域)'); // F欄
  var departmentCol = headers.indexOf('Department(註解)');        // G欄
  var updateStatusCol = headers.indexOf('是否需要更新');           // K欄

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「Email」欄位。請確保工作表包含正確的標題。', ui.ButtonSet.OK);
    return;
  }

  if (familyNameCol === -1 && givenNameCol === -1 && orgUnitPathCol === -1 && employeeIdCol === -1 && employeeTitleCol === -1 && departmentCol === -1) {
    ui.alert('錯誤', '找不到任何可更新的欄位。請確保工作表包含至少其中一個欄位。', ui.ButtonSet.OK);
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
    ui.alert('提示', '沒有找到需要更新的學生。\n\n' +
      (updateStatusCol !== -1 ?
        '所有學生的 K 欄都顯示「無需更新」，或沒有有效的 Email。' :
        '沒有找到有效的 Email。'),
      ui.ButtonSet.OK);
    return;
  }

  // 確認要處理的行數
  var confirmationFields = [];
  if (familyNameCol !== -1) confirmationFields.push('• 更新姓氏 (B欄)');
  if (givenNameCol !== -1) confirmationFields.push('• 更新名字 (C欄)');
  if (orgUnitPathCol !== -1) confirmationFields.push('• 更新機構單位路徑 (D欄)');
  if (employeeIdCol !== -1) confirmationFields.push('• 更新員工編號 (E欄)');
  if (employeeTitleCol !== -1) confirmationFields.push('• 更新職稱 (F欄)');
  if (departmentCol !== -1) confirmationFields.push('• 更新部門 (G欄)');

  var finalConfirmation = ui.alert(
    '最終確認',
    '即將處理 ' + rowsToUpdate.length + ' 位學生的資料' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的學生）' : '') + '。\n\n' +
    '此操作將會：\n' +
    confirmationFields.join('\n') +
    '\n\n確定要執行嗎？',
    ui.ButtonSet.YES_NO
  );

  if (finalConfirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在更新學生資料，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始更新學生資料...'];
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
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail,name,orgUnitPath,organizations,externalIds" });
      } catch (e) {
        logMessages.push(logPrefix + '使用者不存在，跳過。');
        skipCount++;
        continue;
      }

      var needsUserUpdate = false;
      var userObj = {};

      // 處理姓名更新
      var nameObj = {};
      var nameUpdated = false;

      if (familyNameCol !== -1) {
        var newFamilyName = String(row[familyNameCol] || '').trim();
        var currentFamilyName = (user.name && user.name.familyName) ? user.name.familyName : '';

        if (newFamilyName && newFamilyName !== 'N/A' && newFamilyName !== currentFamilyName) {
          nameObj.familyName = newFamilyName;
          nameUpdated = true;
          logMessages.push(logPrefix + '姓氏將從 "' + currentFamilyName + '" 更新為 "' + newFamilyName + '"');
        }
      }

      if (givenNameCol !== -1) {
        var newGivenName = String(row[givenNameCol] || '').trim();
        var currentGivenName = (user.name && user.name.givenName) ? user.name.givenName : '';

        if (newGivenName && newGivenName !== 'N/A' && newGivenName !== currentGivenName) {
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
        needsUserUpdate = true;
      }

      // 處理機構單位路徑更新
      if (orgUnitPathCol !== -1) {
        var newOrgUnitPath = String(row[orgUnitPathCol] || '').trim();
        if (newOrgUnitPath && newOrgUnitPath !== user.orgUnitPath) {
          userObj.orgUnitPath = newOrgUnitPath;
          needsUserUpdate = true;
          logMessages.push(logPrefix + '機構單位路徑將從 "' + user.orgUnitPath + '" 更新為 "' + newOrgUnitPath + '"');
        }
      }

      // 處理 Employee ID 更新
      if (employeeIdCol !== -1) {
        var newEmployeeId = String(row[employeeIdCol] || '').trim();
        if (newEmployeeId === 'N/A') newEmployeeId = '';

        // 取得目前的 Employee ID
        var currentEmployeeId = '';
        if (user.externalIds && user.externalIds.length > 0) {
          for (var j = 0; j < user.externalIds.length; j++) {
            var externalId = user.externalIds[j];
            if (externalId.type === 'organization' || externalId.type === 'work') {
              currentEmployeeId = externalId.value;
              break;
            }
          }
        }

        // 比較 Employee ID 是否需要更新
        if (newEmployeeId !== currentEmployeeId) {
          if (newEmployeeId) {
            userObj.externalIds = [{
              value: newEmployeeId,
              type: 'organization'
            }];
          } else {
            // 如果新 Employee ID 為空，清除 Employee ID
            userObj.externalIds = [];
          }
          needsUserUpdate = true;
          logMessages.push(logPrefix + 'Employee ID 將從 "' + currentEmployeeId + '" 更新為 "' + newEmployeeId + '"');
        }
      }

      // 處理 Employee Title 和 Department 更新
      var needsOrgUpdate = false;
      var newEmployeeTitle = '';
      var newDepartment = '';
      var currentTitle = '';
      var currentDepartment = '';

      // 取得目前的 Employee Title 和 Department
      if (user.organizations && user.organizations.length > 0) {
        for (var j = 0; j < user.organizations.length; j++) {
          var org = user.organizations[j];
          if (org.title) {
            currentTitle = org.title;
          }
          if (org.department) {
            currentDepartment = org.department;
          }
        }
      }

      // 檢查 Employee Title 更新
      if (employeeTitleCol !== -1) {
        newEmployeeTitle = String(row[employeeTitleCol] || '').trim();
        if (newEmployeeTitle === 'N/A') newEmployeeTitle = '';
        if (newEmployeeTitle !== currentTitle) {
          needsOrgUpdate = true;
          logMessages.push(logPrefix + 'Employee Title 將從 "' + currentTitle + '" 更新為 "' + newEmployeeTitle + '"');
        } else {
          newEmployeeTitle = currentTitle; // 保持原值
        }
      } else {
        newEmployeeTitle = currentTitle; // 保持原值
      }

      // 檢查 Department 更新
      if (departmentCol !== -1) {
        newDepartment = String(row[departmentCol] || '').trim();
        if (newDepartment === 'N/A') newDepartment = '';
        if (newDepartment !== currentDepartment) {
          needsOrgUpdate = true;
          logMessages.push(logPrefix + 'Department 將從 "' + currentDepartment + '" 更新為 "' + newDepartment + '"');
        } else {
          newDepartment = currentDepartment; // 保持原值
        }
      } else {
        newDepartment = currentDepartment; // 保持原值
      }

      // 如果需要更新 organizations
      if (needsOrgUpdate) {
        if (newEmployeeTitle || newDepartment) {
          var orgObj = {
            primary: true
          };
          if (newEmployeeTitle) {
            orgObj.title = newEmployeeTitle;
          }
          if (newDepartment) {
            orgObj.department = newDepartment;
          }
          userObj.organizations = [orgObj];
        } else {
          // 如果都為空，清除 organizations
          userObj.organizations = [];
        }
        needsUserUpdate = true;
      }

      // 執行使用者資料更新
      if (needsUserUpdate) {
        AdminDirectory.Users.update(userObj, email);
        logMessages.push(logPrefix + '學生資料已成功更新。');
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
      if (i % 5 === 4) {
        Utilities.sleep(200);
      }

    } catch (e) {
      logMessages.push(logPrefix + '更新時發生錯誤: ' + e.message);
      failCount++;
    }
  }

  var resultMsg = '學生資料更新完成！\n\n' +
    '處理了 ' + rowsToUpdate.length + ' 位學生' +
    (updateStatusCol !== -1 ? '（僅處理標記為「需要更新」的學生）' : '') + '：\n' +
    '成功更新: ' + successCount + ' 位學生\n' +
    '跳過/無需更新: ' + skipCount + ' 位學生\n' +
    '失敗/錯誤: ' + failCount + ' 位學生\n\n' +
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
 * 清理本工作表內所有觸發器
 * 刪除與當前試算表相關聯的所有觸發器
 */
function cleanAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  var deletedCount = 0;

  // 刪除所有觸發器
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if (trigger.getHandlerFunction() === 'exportAllStudents' || 
        trigger.getHandlerFunction() === 'updateStudentsFromSheet' || 
        trigger.getHandlerFunction() === 'exportSuspensionTemplate' ||
        trigger.getHandlerFunction() === 'suspendUsersAtTime' ||
        trigger.getHandlerFunction() === 'sendNotificationEmails') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }

  ui.alert('清理完成', '已成功刪除 ' + deletedCount + ' 個觸發器。', ui.ButtonSet.OK);
}

/**
 * 列出本工作表內所有觸發器
 * 顯示當前試算表所有觸發器的詳細資訊
 */
function listAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var currentSheet = SpreadsheetApp.getActiveSheet().getName();
  
  if (triggers.length === 0) {
    ui.alert('觸發器狀態', '目前整個專案中沒有任何觸發器。', ui.ButtonSet.OK);
    return;
  }

  // 分類觸發器
  var suspendTriggers = [];
  var notificationTriggers = [];
  var otherTriggers = [];

  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    var handlerFunction = trigger.getHandlerFunction();
    var uniqueId = trigger.getUniqueId();
    
    // 基本資訊
    var triggerInfo = {
      id: uniqueId,
      handler: handlerFunction,
      eventType: trigger.getEventType().toString(),
      source: trigger.getTriggerSource().toString(),
      sourceId: trigger.getTriggerSourceId()
    };

    // 獲取詳細資訊
    if (handlerFunction === 'suspendUsersAtTime') {
      var propKey = `trigger_${uniqueId}`;
      var storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          var triggerData = JSON.parse(storedData);
          triggerInfo.targetTime = triggerData.targetTime;
          triggerInfo.sheetName = triggerData.sheetName;
          triggerInfo.accountCount = triggerData.accountCount;
          triggerInfo.isCurrentSheet = (triggerData.sheetName === currentSheet);
        } catch (e) {
          triggerInfo.error = '資料格式錯誤';
        }
      } else {
        triggerInfo.error = '找不到觸發器資料';
      }
      suspendTriggers.push(triggerInfo);
      
    } else if (handlerFunction === 'sendNotificationEmails') {
      var propKey = `notification_trigger_${uniqueId}`;
      var storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          var triggerData = JSON.parse(storedData);
          triggerInfo.notificationTime = triggerData.notificationTime;
          triggerInfo.weeksBeforeSuspend = triggerData.weeksBeforeSuspend;
          triggerInfo.hoursBeforeSuspend = triggerData.hoursBeforeSuspend;
          triggerInfo.isHourNotification = triggerData.isHourNotification;
          triggerInfo.sheetName = triggerData.sheetName;
          triggerInfo.accountCount = triggerData.accountCount;
          triggerInfo.isCurrentSheet = (triggerData.sheetName === currentSheet);
        } catch (e) {
          triggerInfo.error = '資料格式錯誤';
        }
      } else {
        triggerInfo.error = '找不到觸發器資料';
      }
      notificationTriggers.push(triggerInfo);
      
    } else {
      otherTriggers.push(triggerInfo);
    }
  }

  // 建立 HTML 內容
  var htmlContent = `
    <style>
      body { font-family: 'Microsoft JhengHei', Arial, sans-serif; margin: 10px; }
      h3 { color: #1a73e8; margin-bottom: 15px; }
      h4 { color: #d73027; margin-top: 20px; margin-bottom: 10px; }
      .section { margin-bottom: 25px; border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; }
      .trigger-item { 
        background: #f8f9fa; 
        border-left: 4px solid #1a73e8; 
        margin: 10px 0; 
        padding: 12px; 
        border-radius: 4px;
      }
      .current-sheet { border-left-color: #34a853 !important; background: #e8f5e8; }
      .error { border-left-color: #ea4335 !important; background: #fce8e6; }
      .info-row { margin: 5px 0; }
      .label { font-weight: bold; color: #5f6368; }
      .value { color: #202124; }
      .time { color: #1967d2; font-weight: 500; }
      .count { color: #137333; font-weight: 500; }
      .error-text { color: #d93025; font-weight: 500; }
      .summary { background: #e3f2fd; padding: 12px; border-radius: 6px; margin-bottom: 20px; }
      .no-data { color: #5f6368; font-style: italic; text-align: center; padding: 20px; }
    </style>
  `;

  htmlContent += `<h3>📋 觸發器詳細列表</h3>`;
  
  // 摘要資訊
  var currentSheetSuspendCount = suspendTriggers.filter(t => t.isCurrentSheet).length;
  var currentSheetNotificationCount = notificationTriggers.filter(t => t.isCurrentSheet).length;
  var totalCurrentSheet = currentSheetSuspendCount + currentSheetNotificationCount;
  
  htmlContent += `
    <div class="summary">
      <strong>📊 摘要統計</strong><br>
      • 總觸發器數量：<span class="count">${triggers.length}</span> 個<br>
      • 目前工作表「${currentSheet}」相關：<span class="count">${totalCurrentSheet}</span> 個<br>
      • 停權觸發器：<span class="count">${suspendTriggers.length}</span> 個（其中 ${currentSheetSuspendCount} 個屬於目前工作表）<br>
      • 通知觸發器：<span class="count">${notificationTriggers.length}</span> 個（其中 ${currentSheetNotificationCount} 個屬於目前工作表）<br>
      • 其他觸發器：<span class="count">${otherTriggers.length}</span> 個
    </div>
  `;

  // 停權觸發器詳情
  htmlContent += `<div class="section">`;
  htmlContent += `<h4>🚫 停權觸發器 (${suspendTriggers.length} 個)</h4>`;
  
  if (suspendTriggers.length === 0) {
    htmlContent += `<div class="no-data">目前沒有停權觸發器</div>`;
  } else {
    suspendTriggers.forEach(function(trigger, index) {
      var itemClass = 'trigger-item';
      if (trigger.isCurrentSheet) itemClass += ' current-sheet';
      if (trigger.error) itemClass += ' error';
      
      htmlContent += `<div class="${itemClass}">`;
      htmlContent += `<div class="info-row"><span class="label">📌 觸發器 #${index + 1}</span></div>`;
      
      if (trigger.error) {
        htmlContent += `<div class="info-row"><span class="label">❌ 錯誤：</span><span class="error-text">${trigger.error}</span></div>`;
      } else {
        var targetDate = new Date(trigger.targetTime);
        htmlContent += `<div class="info-row"><span class="label">⏰ 停權時間：</span><span class="time">${targetDate.toLocaleString('zh-TW')}</span></div>`;
        htmlContent += `<div class="info-row"><span class="label">📄 工作表：</span><span class="value">${trigger.sheetName}</span> ${trigger.isCurrentSheet ? '(目前工作表)' : ''}</div>`;
        htmlContent += `<div class="info-row"><span class="label">👥 影響帳號：</span><span class="count">${trigger.accountCount}</span> 個</div>`;
      }
      
      htmlContent += `<div class="info-row"><span class="label">🔧 函數：</span><span class="value">${trigger.handler}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">🆔 ID：</span><span class="value">${trigger.id}</span></div>`;
      htmlContent += `</div>`;
    });
  }
  htmlContent += `</div>`;

  // 通知觸發器詳情
  htmlContent += `<div class="section">`;
  htmlContent += `<h4>📧 通知觸發器 (${notificationTriggers.length} 個)</h4>`;
  
  if (notificationTriggers.length === 0) {
    htmlContent += `<div class="no-data">目前沒有通知觸發器</div>`;
  } else {
    notificationTriggers.forEach(function(trigger, index) {
      var itemClass = 'trigger-item';
      if (trigger.isCurrentSheet) itemClass += ' current-sheet';
      if (trigger.error) itemClass += ' error';
      
      htmlContent += `<div class="${itemClass}">`;
      htmlContent += `<div class="info-row"><span class="label">📌 觸發器 #${index + 1}</span></div>`;
      
      if (trigger.error) {
        htmlContent += `<div class="info-row"><span class="label">❌ 錯誤：</span><span class="error-text">${trigger.error}</span></div>`;
      } else {
        var notificationDate = new Date(trigger.notificationTime);
        var timeDesc = trigger.isHourNotification ? 
          `停權前 ${trigger.hoursBeforeSuspend} 小時` : 
          `停權前 ${trigger.weeksBeforeSuspend} 週`;
        
        htmlContent += `<div class="info-row"><span class="label">📨 通知時間：</span><span class="time">${notificationDate.toLocaleString('zh-TW')}</span></div>`;
        htmlContent += `<div class="info-row"><span class="label">⏱️ 通知類型：</span><span class="value">${timeDesc}</span></div>`;
        htmlContent += `<div class="info-row"><span class="label">📄 工作表：</span><span class="value">${trigger.sheetName}</span> ${trigger.isCurrentSheet ? '(目前工作表)' : ''}</div>`;
        htmlContent += `<div class="info-row"><span class="label">👥 影響帳號：</span><span class="count">${trigger.accountCount}</span> 個</div>`;
      }
      
      htmlContent += `<div class="info-row"><span class="label">🔧 函數：</span><span class="value">${trigger.handler}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">🆔 ID：</span><span class="value">${trigger.id}</span></div>`;
      htmlContent += `</div>`;
    });
  }
  htmlContent += `</div>`;

  // 其他觸發器詳情
  if (otherTriggers.length > 0) {
    htmlContent += `<div class="section">`;
    htmlContent += `<h4>🔧 其他觸發器 (${otherTriggers.length} 個)</h4>`;
    
    otherTriggers.forEach(function(trigger, index) {
      htmlContent += `<div class="trigger-item">`;
      htmlContent += `<div class="info-row"><span class="label">📌 觸發器 #${index + 1}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">🔧 函數：</span><span class="value">${trigger.handler}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">📋 事件類型：</span><span class="value">${trigger.eventType}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">📂 觸發來源：</span><span class="value">${trigger.source}</span></div>`;
      htmlContent += `<div class="info-row"><span class="label">🆔 ID：</span><span class="value">${trigger.id}</span></div>`;
      htmlContent += `</div>`;
    });
    
    htmlContent += `</div>`;
  }

  // 說明文字
  htmlContent += `
    <div class="section">
      <h4>📋 說明</h4>
      <div style="font-size: 14px; line-height: 1.6;">
        <p><strong>🟢 綠色背景</strong>：屬於目前工作表「${currentSheet}」的觸發器</p>
        <p><strong>🔵 藍色背景</strong>：其他工作表的觸發器</p>
        <p><strong>🔴 紅色背景</strong>：有錯誤或資料缺失的觸發器</p>
        <br>
        <p><strong>停權觸發器</strong>：在指定時間自動停權使用者帳號</p>
        <p><strong>通知觸發器</strong>：在停權前的指定時間發送通知信</p>
        <p><strong>其他觸發器</strong>：非停權相關的觸發器（如定時匯出等）</p>
      </div>
    </div>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(800)
    .setHeight(600);

  ui.showModalDialog(htmlOutput, `📋 觸發器詳細列表 (共 ${triggers.length} 個)`);
}

/**
 * 匯出機構單位路徑為 "/預約停權" 的使用者到新工作表
 */
function exportSuspensionTemplate() {
  var ui = SpreadsheetApp.getUi();

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在匯出離校學生清單，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始匯出離校學生清單...'];

  try {
    // 步驟 1: 先獲取所有使用者，然後篩選出機構單位路徑為 "/離校學生" 的使用者
    var retiredStudents = [];
    var processedCount = 0;
    var totalCount = 0;

    logMessages.push('正在讀取所有學生資料並篩選離校學生...');

    var pageToken;
    do {
      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,organizations,suspended,creationTime,lastLoginTime)'
      });

      if (page.users) {
        totalCount += page.users.length;
        
        // 篩選出機構單位路徑為 "/預約停權" 的使用者
        for (var i = 0; i < page.users.length; i++) {
          var user = page.users[i];
          if (user.orgUnitPath === '/預約停權') {
            retiredStudents.push(user);
            processedCount++;
          }
        }
        
        logMessages.push('已掃描 ' + totalCount + ' 位學生，找到 ' + processedCount + ' 位預約停權學生...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (retiredStudents.length === 0) {
      ui.alert(
        '結果', 
        '未找到任何機構單位路徑為 "/預約停權" 的使用者。\n\n' +
        '已掃描總學生數：' + totalCount + '\n' +
        '找到離校學生數：0\n\n' +
        '請確認：\n' +
        '1. 機構單位 "/預約停權" 是否存在\n' +
        '2. 是否有學生被分配到此機構單位', 
        ui.ButtonSet.OK
      );
      return;
    }

    logMessages.push('學生掃描完成，總共掃描 ' + totalCount + ' 位學生，找到 ' + retiredStudents.length + ' 位離校學生，開始整理資料...');

    // 步驟 2: 準備要寫入工作表的資料（在 H 欄之後新增四個欄位）
    var outputData = [[
      'Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      'Department(註解)',
      '帳號狀態',
      '建立時間',
      '最後登入時間',
      '停權日期',           // I欄：新增
      '目前進度',           // J欄：新增
      '錯誤訊息',           // K欄：新增
      '郵件通知進度'        // L欄：新增
    ]];

    // 步驟 3: 處理每位離校學生的資料
    for (var i = 0; i < retiredStudents.length; i++) {
      var user = retiredStudents[i];

      var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
      var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
      var orgUnitPath = user.orgUnitPath || '/';

      // 取得 Department
      var department = 'N/A';
      if (user.organizations && user.organizations.length > 0) {
        for (var j = 0; j < user.organizations.length; j++) {
          var org = user.organizations[j];
          if (org.department) {
            department = org.department;
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
        user.primaryEmail,    // A欄: Email
        familyName,           // B欄: 姓 (Family Name)
        givenName,            // C欄: 名 (Given Name)
        orgUnitPath,          // D欄: 機構單位路徑
        department,           // E欄: Department(註解)
        status,               // F欄: 帳號狀態
        creationTime,         // G欄: 建立時間
        lastLoginTime,        // H欄: 最後登入時間
        '',                   // I欄: 停權日期（留空）
        '待處理',             // J欄: 目前進度
        '',                   // K欄: 錯誤訊息（留空）
        '未通知'              // L欄: 郵件通知進度
      ]);

      // 顯示進度（每處理 10 位學生顯示一次）
      if ((i + 1) % 10 === 0 || i === retiredStudents.length - 1) {
        logMessages.push('已處理 ' + (i + 1) + '/' + retiredStudents.length + ' 位離校學生的資料...');
      }
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "[預約停權]";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 步驟 5: 設定格式（保持您原有的欄位寬度，並為新欄位設定寬度）
    var columnWidths = {
      1: 60,   // A欄：Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 100,  // D欄：機構單位路徑
      5: 80,   // E欄：Department(註解)
      6: 60,   // F欄：帳號狀態
      7: 80,   // G欄：建立時間
      8: 80,   // H欄：最後登入時間
      9: 80,   // I欄：停權日期
      10: 80,  // J欄：目前進度
      11: 100, // K欄：錯誤訊息
      12: 100  // L欄：郵件通知進度
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 12; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定標題行格式
    var headerRange = newSheet.getRange(1, 1, 1, 12);
    headerRange.setBackground('#FF6B6B')
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');

    // 設定所有資料範圍的格式
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 12);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
    }

    // 凍結標題行
    newSheet.setFrozenRows(1);

    // 設定資料驗證 - 停權日期欄位（I欄）
    if (outputData.length > 1) {
      var dateRange = newSheet.getRange(2, 9, outputData.length - 1, 1);
      
      // 修改資料驗證，允許日期時間格式
      var dateValidation = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .setHelpText('請輸入日期時間，格式範例：\n• 2024/12/25 14:30\n• 2024-12-25 14:30:00\n• 或直接輸入 =NOW() 取得現在時間')
        .build();
      dateRange.setDataValidation(dateValidation);
      
      // 設定 I 欄的數字格式為日期時間格式
      dateRange.setNumberFormat('yyyy/mm/dd hh:mm:ss');
    }

    // 設定帳號狀態欄位的條件格式（F欄）
    if (outputData.length > 1) {
      var statusRange = newSheet.getRange(2, 6, outputData.length - 1, 1); // F欄

      var suspendedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("已停用")
        .setBackground("#FFE6E6")
        .setFontColor("#CC0000")
        .setRanges([statusRange])
        .build();

      var activeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("啟用中")
        .setBackground("#E6F7E6")
        .setFontColor("#008000")
        .setRanges([statusRange])
        .build();

      var rules = newSheet.getConditionalFormatRules();
      rules.push(suspendedRule);
      rules.push(activeRule);
      newSheet.setConditionalFormatRules(rules);
    }

    // 設定條件格式 - 目前進度欄位（J欄）
    if (outputData.length > 1) {
      var progressRange = newSheet.getRange(2, 10, outputData.length - 1, 1);

      var waitingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("待處理")
        .setBackground("#FFF2CC")
        .setFontColor("#BF9000")
        .setRanges([progressRange])
        .build();

      var processingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("處理中")
        .setBackground("#FCE5CD")
        .setFontColor("#B45F06")
        .setRanges([progressRange])
        .build();

      var completedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("已停權")
        .setBackground("#D9EAD3")
        .setFontColor("#274E13")
        .setRanges([progressRange])
        .build();

      var errorRule = SpreadsheetApp.newConditionalFormatRule()  
        .whenTextEqualTo("錯誤")
        .setBackground("#F4CCCC")
        .setFontColor("#CC0000")
        .setRanges([progressRange])
        .build();

      var currentRules = newSheet.getConditionalFormatRules();
      currentRules.push(waitingRule);
      currentRules.push(processingRule);
      currentRules.push(completedRule);
      currentRules.push(errorRule);
      newSheet.setConditionalFormatRules(currentRules);
    }

    // 設定條件格式 - 郵件通知進度欄位（L欄）
    if (outputData.length > 1) {
      var notificationRange = newSheet.getRange(2, 12, outputData.length - 1, 1);

      var notNotifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("未通知")
        .setBackground("#FFF2CC")
        .setFontColor("#BF9000")
        .setRanges([notificationRange])
        .build();

      var notifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("已通知")
        .setBackground("#D9EAD3")
        .setFontColor("#274E13")
        .setRanges([notificationRange])
        .build();

      var notificationErrorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("通知失敗")
        .setBackground("#F4CCCC")
        .setFontColor("#CC0000")
        .setRanges([notificationRange])
        .build();

      var allRules = newSheet.getConditionalFormatRules();
      allRules.push(notNotifiedRule);
      allRules.push(notifiedRule);
      allRules.push(notificationErrorRule);
      newSheet.setConditionalFormatRules(allRules);
    }

    // 步驟 7: 在工作表底部添加統計資訊
    var statsStartRow = outputData.length + 3;
    var activeCount = 0;
    var suspendedCount = 0;

    for (var i = 1; i < outputData.length; i++) {
      if (outputData[i][5] === '啟用中') {  // F欄：帳號狀態
        activeCount++;
      } else if (outputData[i][5] === '已停用') {
        suspendedCount++;
      }
    }

    var statsData = [
      ['=== 離校學生統計資訊 ==='],
      [''],
      ['掃描範圍：全部學生 (' + totalCount + ' 位)'],
      ['總離校學生數：' + (outputData.length - 1)],
      ['啟用中帳號：' + activeCount],
      ['已停用帳號：' + suspendedCount],
      [''],
      ['匯出時間：' + new Date().toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() })],
      ['篩選條件：機構單位路徑 = "/離校學生"']
    ];

    newSheet.getRange(statsStartRow, 1, statsData.length, 1).setValues(statsData);

    // 設定統計資訊格式
    var statsRange = newSheet.getRange(statsStartRow, 1, statsData.length, 1);
    statsRange.setFontSize(10)
             .setFontColor('#666666');

    newSheet.getRange(statsStartRow, 1, 1, 1)
           .setFontWeight('bold')
           .setFontColor('#FF6B6B');

    newSheet.activate();

    logMessages.push('離校學生清單匯出完成！共包含 ' + (outputData.length - 1) + ' 位離校學生。');

    ui.alert(
      '匯出成功！', 
      '離校學生清單已成功匯出，共包含 ' + (outputData.length - 1) + ' 位離校學生。\n\n' +
      '掃描統計：\n' +
      '• 總掃描學生：' + totalCount + ' 位\n' +
      '• 找到離校學生：' + (outputData.length - 1) + ' 位\n' +
      '• 啟用中帳號：' + activeCount + ' 位\n' +
      '• 已停用帳號：' + suspendedCount + ' 位\n\n' +
      '功能特點：\n' +
      '• 已設定自動篩選功能\n' +
      '• 包含條件格式和資料驗證\n' +
      '• 包含統計資訊\n' +
      '• 新增停權管理相關欄位\n\n' +
      '工作表名稱：「' + sheetName + '」', 
      ui.ButtonSet.OK
    );

  } catch (e) {
    var errorMsg = '匯出離校學生清單時發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法匯出離校學生清單。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉處理中提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}

/**
 * 啟動完整的停權程序（包含通知信和停權觸發器）
 */
function scheduleCompleteSuspensionProcess() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  
  // 確認對話框
  const confirmation = ui.alert(
    '啟動完整停權程序',
    '此功能將依據工作表中的「停權時間」啟動完整的停權程序：\n\n' +
    '📧 通知信排程：\n' +
    '• 停權前 4 週通知\n' +
    '• 停權前 3 週通知\n' +
    '• 停權前 2 週通知\n' +
    '• 停權前 1 週通知\n' +
    '• 停權前 6 小時最後通知\n\n' +
    '⏰ 停權觸發器：\n' +
    '• 在指定時間自動停權帳號\n\n' +
    '⚠️ 注意：此操作會清除現有的相關觸發器並重新建立。\n\n' +
    '確定要啟動完整停權程序嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在啟動完整停權程序，請稍候...</b>').setTitle('處理中'));

  try {
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    let processedAccounts = 0;
    let validAccounts = 0;

    // 欄位索引
    const emailColumnIndex = 0;    // A欄：Email
    const timeColumnIndex = 8;     // I欄：停權日期
    const statusColumnIndex = 9;   // J欄：目前進度
    const errorColumnIndex = 10;   // K欄：錯誤訊息
    const mailStatusColumnIndex = 11; // L欄：郵件通知進度

    // 第一步：驗證資料並統計
    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const timeStr = data[row][timeColumnIndex];
      
      if (!email || !timeStr) continue;
      processedAccounts++;

      const suspendDate = new Date(timeStr);
      if (isNaN(suspendDate.getTime())) {
        sheet.getRange(row + 1, errorColumnIndex + 1).setValue('時間格式錯誤');
        continue;
      }

      if (suspendDate <= now) {
        sheet.getRange(row + 1, errorColumnIndex + 1).setValue('時間已過期');
        continue;
      }

      validAccounts++;
      // 清除錯誤訊息
      sheet.getRange(row + 1, errorColumnIndex + 1).setValue('');
    }

    if (validAccounts === 0) {
      ui.alert(
        '無有效資料',
        `在工作表「${sheetName}」中找到 ${processedAccounts} 筆資料，但沒有有效的未來停權時間。\n\n` +
        '請檢查：\n' +
        '• I欄停權日期格式是否正確\n' +
        '• 停權時間是否為未來時間\n' +
        '• A欄是否有有效的 Email',
        ui.ButtonSet.OK
      );
      return;
    }

    // 第二步：建立通知信觸發器
    console.log('開始建立通知信觸發器...');
    const notificationResult = createNotificationTriggers(sheet, sheetName, data, now);

    // 第三步：建立停權觸發器
    console.log('開始建立停權觸發器...');
    const suspensionResult = createSuspensionTriggers(sheet, sheetName, data, now);

    // 第四步：更新工作表狀態
    updateSheetStatus(sheet, data, notificationResult.notificationTimes, suspensionResult.futureTimes, now);

    // 顯示結果
    const resultMessage = 
      `完整停權程序啟動成功！\n\n` +
      `工作表：「${sheetName}」\n` +
      `處理帳號：${validAccounts} 個有效帳號\n\n` +
      `📧 通知信觸發器：${notificationResult.createdCount} 個\n` +
      `${notificationResult.summary}\n\n` +
      `⏰ 停權觸發器：${suspensionResult.createdCount} 個\n` +
      `${suspensionResult.summary}\n\n` +
      `✅ 停權程序已完全啟動，系統將自動：\n` +
      `• 在預定時間發送通知信\n` +
      `• 在停權時間執行帳號停權`;

    ui.alert('停權程序啟動成功', resultMessage, ui.ButtonSet.OK);

  } catch (error) {
    console.error('啟動停權程序時發生錯誤:', error);
    ui.alert('錯誤', `啟動停權程序時發生錯誤：\n\n${error.message}`, ui.ButtonSet.OK);
  } finally {
    ui.showSidebar(HtmlService.createHtmlOutput('<b>停權程序啟動完成！</b>').setTitle('完成'));
  }
}

/**
 * 建立通知信觸發器（內部函數）
 */
function createNotificationTriggers(sheet, sheetName, data, now) {
  const notificationTimes = new Set();
  const emailColumnIndex = 0;
  const timeColumnIndex = 8;
  const errorColumnIndex = 10;
  const mailStatusColumnIndex = 11;

  // 收集所有通知時間點
  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex];
    const timeStr = data[row][timeColumnIndex];
    if (!email || !timeStr) continue;

    const suspendDate = new Date(timeStr);
    if (isNaN(suspendDate.getTime()) || suspendDate <= now) continue;

    // 計算通知時間點（4、3、2、1週前 + 6小時前）
    for (let weeks = 4; weeks >= 1; weeks--) {
      const notificationDate = new Date(suspendDate.getTime() - (weeks * 7 * 24 * 60 * 60 * 1000));
      if (notificationDate > now) {
        notificationTimes.add(`${notificationDate.toISOString()}_${weeks}week`);
      }
    }

    const sixHoursBeforeDate = new Date(suspendDate.getTime() - (6 * 60 * 60 * 1000));
    if (sixHoursBeforeDate > now) {
      notificationTimes.add(`${sixHoursBeforeDate.toISOString()}_6hour`);
    }
  }

  // 刪除現有通知觸發器
  const allTriggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  for (let trig of allTriggers) {
    if (trig.getHandlerFunction() === 'sendNotificationEmails') {
      const propKey = `notification_trigger_${trig.getUniqueId()}`;
      const storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          const triggerData = JSON.parse(storedData);
          if (triggerData.sheetName === sheetName) {
            ScriptApp.deleteTrigger(trig);
            PropertiesService.getScriptProperties().deleteProperty(propKey);
            deletedCount++;
          }
        } catch (e) {
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          deletedCount++;
        }
      }
    }
  }

  // 建立新的通知觸發器
  let createdCount = 0;
  const triggerSummary = [];

  for (const timeTypeStr of notificationTimes) {
    const [timeStr, typeStr] = timeTypeStr.split('_');
    const triggerTime = new Date(timeStr);

    let weeksBeforeSuspend = null;
    let hoursBeforeSuspend = null;
    let isHourNotification = false;

    if (typeStr.endsWith('week')) {
      weeksBeforeSuspend = parseInt(typeStr);
    } else if (typeStr.endsWith('hour')) {
      hoursBeforeSuspend = parseInt(typeStr);
      isHourNotification = true;
    }

    // 統計帳號數量
    let accountCount = 0;
    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const rowTimeStr = data[row][timeColumnIndex];
      if (!email || !rowTimeStr) continue;

      const suspendDate = new Date(rowTimeStr);
      if (isNaN(suspendDate.getTime())) continue;

      let expectedNotificationDate;
      if (isHourNotification) {
        expectedNotificationDate = new Date(suspendDate.getTime() - (hoursBeforeSuspend * 60 * 60 * 1000));
      } else {
        expectedNotificationDate = new Date(suspendDate.getTime() - (weeksBeforeSuspend * 7 * 24 * 60 * 60 * 1000));
      }

      if (Math.abs(expectedNotificationDate.getTime() - triggerTime.getTime()) < 60 * 1000) {
        accountCount++;
      }
    }

    // 建立觸發器
    const trigger = ScriptApp.newTrigger('sendNotificationEmails')
      .timeBased()
      .at(triggerTime)
      .create();

    // 儲存觸發器資訊
    const triggerData = {
      notificationTime: timeStr,
      weeksBeforeSuspend: weeksBeforeSuspend,
      hoursBeforeSuspend: hoursBeforeSuspend,
      isHourNotification: isHourNotification,
      sheetName: sheetName,
      accountCount: accountCount
    };

    PropertiesService.getScriptProperties().setProperty(
      `notification_trigger_${trigger.getUniqueId()}`,
      JSON.stringify(triggerData)
    );

    const displayText = isHourNotification ?
      `停權前 ${hoursBeforeSuspend} 小時` :
      `停權前 ${weeksBeforeSuspend} 週`;

    triggerSummary.push(`• ${displayText}: ${accountCount} 個帳號`);
    createdCount++;
  }

  return {
    notificationTimes,
    createdCount,
    deletedCount,
    summary: triggerSummary.join('\n')
  };
}

/**
 * 建立停權觸發器（內部函數）
 */
function createSuspensionTriggers(sheet, sheetName, data, now) {
  const futureTimes = new Set();
  const emailColumnIndex = 0;
  const timeColumnIndex = 8;
  const errorColumnIndex = 10;

  // 收集所有未來停權時間
  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex];
    const timeStr = data[row][timeColumnIndex];
    if (!email || !timeStr) continue;

    const date = new Date(timeStr);
    if (isNaN(date.getTime()) || date <= now) continue;

    futureTimes.add(date.toISOString());
  }

  // 刪除現有停權觸發器
  const allTriggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  for (let trig of allTriggers) {
    if (trig.getHandlerFunction() === 'suspendUsersAtTime') {
      const propKey = `trigger_${trig.getUniqueId()}`;
      const storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          const triggerData = JSON.parse(storedData);
          if (triggerData.sheetName === sheetName) {
            ScriptApp.deleteTrigger(trig);
            PropertiesService.getScriptProperties().deleteProperty(propKey);
            deletedCount++;
          }
        } catch (e) {
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          deletedCount++;
        }
      }
    }
  }

  // 建立新的停權觸發器
  let createdCount = 0;
  const triggerSummary = [];

  for (const timeStr of futureTimes) {
    const triggerTime = new Date(timeStr);

    // 統計帳號數量
    let accountCount = 0;
    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const rowTimeStr = data[row][timeColumnIndex];
      if (!email || !rowTimeStr) continue;

      const rowDate = new Date(rowTimeStr);
      if (isNaN(rowDate.getTime())) continue;

      if (Math.abs(rowDate.getTime() - triggerTime.getTime()) < 60 * 1000) {
        accountCount++;
      }
    }

    // 建立觸發器
    const trigger = ScriptApp.newTrigger('suspendUsersAtTime')
      .timeBased()
      .at(triggerTime)
      .create();

    // 儲存觸發器資訊
    const triggerData = {
      targetTime: timeStr,
      sheetName: sheetName,
      accountCount: accountCount
    };

    PropertiesService.getScriptProperties().setProperty(
      `trigger_${trigger.getUniqueId()}`,
      JSON.stringify(triggerData)
    );

    triggerSummary.push(`• ${triggerTime.toLocaleString('zh-TW')}: ${accountCount} 個帳號`);
    createdCount++;
  }

  return {
    futureTimes,
    createdCount,
    deletedCount,
    summary: triggerSummary.join('\n')
  };
}

/**
 * 更新工作表狀態（內部函數）
 */
function updateSheetStatus(sheet, data, notificationTimes, futureTimes, now) {
  const emailColumnIndex = 0;
  const timeColumnIndex = 8;
  const statusColumnIndex = 9;
  const mailStatusColumnIndex = 11;

  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex];
    const timeStr = data[row][timeColumnIndex];
    if (!email || !timeStr) continue;

    const suspendDate = new Date(timeStr);
    if (isNaN(suspendDate.getTime()) || suspendDate <= now) continue;

    // 檢查是否有停權觸發器
    const suspendKey = suspendDate.toISOString();
    if (futureTimes.has(suspendKey)) {
      sheet.getRange(row + 1, statusColumnIndex + 1).setValue('已預約停權');
    }

    // 檢查是否有通知觸發器
    let hasNotifications = false;
    
    // 檢查週通知
    for (let weeks = 4; weeks >= 1; weeks--) {
      const notificationDate = new Date(suspendDate.getTime() - (weeks * 7 * 24 * 60 * 60 * 1000));
      if (notificationDate > now) {
        const key = `${notificationDate.toISOString()}_${weeks}week`;
        if (notificationTimes.has(key)) {
          hasNotifications = true;
          break;
        }
      }
    }

    // 檢查小時通知
    if (!hasNotifications) {
      const sixHoursBeforeDate = new Date(suspendDate.getTime() - (6 * 60 * 60 * 1000));
      if (sixHoursBeforeDate > now) {
        const key = `${sixHoursBeforeDate.toISOString()}_6hour`;
        if (notificationTimes.has(key)) {
          hasNotifications = true;
        }
      }
    }

    if (hasNotifications) {
      sheet.getRange(row + 1, mailStatusColumnIndex + 1).setValue('已預約連續通知信');
    }
  }
}

/**
 * 清理預約停權相關的所有觸發器
 */
function cleanAllSuspensionTriggers() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const allTriggers = ScriptApp.getProjectTriggers();

  let deletedSuspendTriggers = 0;
  let deletedNotificationTriggers = 0;

  // 清理停權觸發器
  for (let trig of allTriggers) {
    if (trig.getHandlerFunction() === 'suspendUsersAtTime') {
      const propKey = `trigger_${trig.getUniqueId()}`;
      const storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          const triggerData = JSON.parse(storedData);
          if (triggerData.sheetName === sheetName) {
            ScriptApp.deleteTrigger(trig);
            PropertiesService.getScriptProperties().deleteProperty(propKey);
            console.log(`刪除工作表 ${sheetName} 的停權觸發器（UID=${trig.getUniqueId()}）`);
            deletedSuspendTriggers++;
          }
        } catch (e) {
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          deletedSuspendTriggers++;
        }
      }
    }
  }

  // 清理通知觸發器
  for (let trig of allTriggers) {
    if (trig.getHandlerFunction() === 'sendNotificationEmails') {
      const propKey = `notification_trigger_${trig.getUniqueId()}`;
      const storedData = PropertiesService.getScriptProperties().getProperty(propKey);
      if (storedData) {
        try {
          const triggerData = JSON.parse(storedData);
          if (triggerData.sheetName === sheetName) {
            ScriptApp.deleteTrigger(trig);
            PropertiesService.getScriptProperties().deleteProperty(propKey);
            console.log(`刪除工作表 ${sheetName} 的通知觸發器（UID=${trig.getUniqueId()}）`);
            deletedNotificationTriggers++;
          }
        } catch (e) {



          // JSON 解析失敗但仍屬於該工作表的觸發器，也刪除
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          console.log(`刪除工作表 ${sheetName} 的損壞通知觸發器（UID=${trig.getUniqueId()}）`);
          deletedNotificationTriggers++;
        }
      }
    }
  }

  // 清空相關狀態欄位
  const data = sheet.getDataRange().getValues();
  let clearedCells = 0;

  // 欄位索引（根據 [預約停權] 工作表的結構）
  const emailColumnIndex = 0;    // A欄：Email
  const statusColumnIndex = 9;   // J欄：目前進度
  const mailStatusColumnIndex = 11; // L欄：郵件通知進度

  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex];
    if (!email) continue; // 跳過沒有 email 的列

    // 清空 J 欄（狀態欄）- 只清理觸發器設定的狀態
    const statusCell = sheet.getRange(row + 1, statusColumnIndex + 1);
    const currentStatus = statusCell.getValue();
    if (currentStatus === '已預約停權' || currentStatus === '已預約') {
      statusCell.setValue('待處理');
      clearedCells++;
    }

    // 清空 L 欄（郵件狀態欄）- 只清理觸發器設定的狀態
    const mailStatusCell = sheet.getRange(row + 1, mailStatusColumnIndex + 1);
    const currentMailStatus = mailStatusCell.getValue();
    if (currentMailStatus && (
      currentMailStatus.includes('已預約連續通知信') ||
      currentMailStatus.includes('已發送') ||
      currentMailStatus.includes('前通知')
    )) {
      mailStatusCell.setValue('未通知');
      clearedCells++;
    }
  }

  const totalDeleted = deletedSuspendTriggers + deletedNotificationTriggers;

  if (totalDeleted > 0 || clearedCells > 0) {
    console.log(`工作表「${sheetName}」清理完成：`);
    console.log(`- 停權觸發器：${deletedSuspendTriggers} 個`);
    console.log(`- 通知觸發器：${deletedNotificationTriggers} 個`);
    console.log(`- 清空相關狀態：${clearedCells} 個儲存格`);

    ui.alert(
      '清理完成',
      `工作表「${sheetName}」清理完成：\n\n` +
      `• 停權觸發器：${deletedSuspendTriggers} 個\n` +
      `• 通知觸發器：${deletedNotificationTriggers} 個\n` +
      `• 清空相關狀態：${clearedCells} 個儲存格\n\n` +
      `已將狀態重置為初始值：\n` +
      `• J欄：重置為「待處理」\n` +
      `• L欄：重置為「未通知」`,
      ui.ButtonSet.OK
    );
  } else {
    console.log(`工作表「${sheetName}」目前沒有任何觸發器或相關狀態需要清理`);
    ui.alert(
      '無需清理',
      `工作表「${sheetName}」目前沒有任何觸發器或相關狀態需要清理。`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 停權指定時間的所有帳號（由觸發器自動執行）
 */
function suspendUsersAtTime(e) {
  try {
    console.log('觸發器開始執行');

    const thisTriggerId = e?.triggerUid;
    console.log('觸發器 ID:', thisTriggerId);

    let targetTime = null;
    let sheetName = null;

    if (thisTriggerId) {
      const storedData = PropertiesService.getScriptProperties().getProperty(`trigger_${thisTriggerId}`);
      if (storedData) {
        const triggerData = JSON.parse(storedData);
        targetTime = triggerData.targetTime;
        sheetName = triggerData.sheetName;
        console.log('從 Properties 獲取的目標時間:', targetTime);
        console.log('從 Properties 獲取的工作表名稱:', sheetName);
      }
    }

    // 使用指定的工作表，如果沒有則使用活躍工作表
    let sheet;
    if (sheetName) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        console.log(`❌ 找不到工作表：${sheetName}`);
        return;
      }
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }

    const data = sheet.getDataRange().getValues();
    const now = new Date();

    // 欄位索引（根據 [預約停權] 工作表的結構）
    const emailColumnIndex = 0;    // A欄：Email
    const timeColumnIndex = 8;     // I欄：停權日期
    const statusColumnIndex = 9;   // J欄：目前進度
    const errorColumnIndex = 10;   // K欄：錯誤訊息

    console.log(`處理工作表：${sheet.getName()}`);
    console.log('處理的資料筆數:', data.length);
    console.log('當前時間:', now.toISOString());

    let processedCount = 0;

    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const timeStr = data[row][timeColumnIndex];
      if (!email || !timeStr) continue;

      const date = new Date(timeStr);
      if (isNaN(date.getTime())) continue;

      console.log(`檢查第 ${row + 1} 列 - 帳號: ${email}, 預定時間: ${timeStr}`);
      console.log(`  轉換後的時間: ${date.toISOString()}`);

      let shouldSuspend = false;

      if (targetTime) {
        // 有指定目標時間，比對是否一致
        const targetDate = new Date(targetTime);
        const timeDiff = Math.abs(date.getTime() - targetDate.getTime());
        console.log(`  目標時間: ${targetDate.toISOString()}`);
        console.log(`  時間差異: ${timeDiff / 1000} 秒`);

        // 改為使用 1 分鐘誤差，與建立觸發器時一致
        if (timeDiff < 60 * 1000) {
          shouldSuspend = true;
          console.log(`  ✅ 時間匹配 (目標時間比對)`);
        } else {
          console.log(`  ❌ 時間不匹配`);
        }
      } else {
        // 沒有指定目標時間，檢查是否已到預定時間
        // 同樣改為 1 分鐘誤差
        if (date <= now && (now.getTime() - date.getTime()) < 60 * 1000) {
          shouldSuspend = true;
          console.log(`  ✅ 時間匹配 (當前時間比對)`);
        }
      }

      if (shouldSuspend) {
        console.log(`準備停權: ${email}`);
        try {
          AdminDirectory.Users.update({ suspended: true }, email);
          sheet.getRange(row + 1, statusColumnIndex + 1).setValue('已停權');
          console.log(`✅ 停權成功：${email}`);
          processedCount++;
        } catch (err) {
          console.log(`❌ 停權失敗 (${email}): ${err.message}`);
          sheet.getRange(row + 1, errorColumnIndex + 1).setValue(`錯誤: ${err.message}`);
        }
      }
    }

    console.log(`觸發器執行完成，共處理 ${processedCount} 個帳號`);

    // 清理 Properties
    if (thisTriggerId) {
      PropertiesService.getScriptProperties().deleteProperty(`trigger_${thisTriggerId}`);
    }

  } catch (error) {
    console.log('觸發器執行發生錯誤:', error.message);
    console.log('錯誤詳細:', error.toString());
  }
}

/**
 * 發送通知信（由觸發器自動執行）
 */
function sendNotificationEmails(e) {
  try {
    console.log('通知信觸發器開始執行');

    const thisTriggerId = e?.triggerUid;
    console.log('觸發器 ID:', thisTriggerId);

    let notificationTime = null;
    let weeksBeforeSuspend = null;
    let hoursBeforeSuspend = null;
    let isHourNotification = false;
    let sheetName = null;

    if (thisTriggerId) {
      const storedData = PropertiesService.getScriptProperties().getProperty(`notification_trigger_${thisTriggerId}`);
      if (storedData) {
        const triggerData = JSON.parse(storedData);
        notificationTime = triggerData.notificationTime;
        weeksBeforeSuspend = triggerData.weeksBeforeSuspend;
        hoursBeforeSuspend = triggerData.hoursBeforeSuspend;
        isHourNotification = triggerData.isHourNotification;
        sheetName = triggerData.sheetName;
        console.log('通知時間:', notificationTime);
        console.log('停權前週數:', weeksBeforeSuspend);
        console.log('停權前小時數:', hoursBeforeSuspend);
        console.log('是否為小時通知:', isHourNotification);
        console.log('工作表名稱:', sheetName);
      }
    }

    // 使用指定的工作表
    let sheet;
    if (sheetName) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        console.log(`❌ 找不到工作表：${sheetName}`);
        return;
      }
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }

    const data = sheet.getDataRange().getValues();
    const now = new Date();

    // 欄位索引（根據 [預約停權] 工作表的結構）
    const emailColumnIndex = 0;    // A欄：Email
    const timeColumnIndex = 8;     // I欄：停權日期
    const errorColumnIndex = 10;   // K欄：錯誤訊息
    const mailStatusColumnIndex = 11; // L欄：郵件通知進度

    console.log(`處理工作表：${sheet.getName()}`);
    console.log('當前時間:', now.toISOString());

    let sentCount = 0;

    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const timeStr = data[row][timeColumnIndex];
      if (!email || !timeStr) continue;

      const suspendDate = new Date(timeStr);
      if (isNaN(suspendDate.getTime())) continue;

      // 計算預期的通知時間
      let expectedNotificationTime;
      if (isHourNotification) {
        expectedNotificationTime = new Date(suspendDate.getTime() - (hoursBeforeSuspend * 60 * 60 * 1000));
      } else {
        expectedNotificationTime = new Date(suspendDate.getTime() - (weeksBeforeSuspend * 7 * 24 * 60 * 60 * 1000));
      }

      const timeDiff = Math.abs(expectedNotificationTime.getTime() - now.getTime());

      console.log(`檢查第 ${row + 1} 列 - 帳號: ${email}, 停權時間: ${timeStr}`);
      console.log(`  預期通知時間: ${expectedNotificationTime.toISOString()}`);
      console.log(`  時間差異: ${timeDiff / 1000} 秒`);

      // 如果時間匹配（允許1分鐘誤差）
      if (timeDiff < 60 * 1000) {
        console.log(`準備發送通知信給: ${email}`);
        try {
          // 發送通知信
          let subject, body;
          const timeInfo = isHourNotification ?
            `${hoursBeforeSuspend} 小時` :
            `${weeksBeforeSuspend} 週`;

          subject = `[信箱停用通知] 因應國教署資安政策，離校/畢業帳號停權通知 - 本帳號預計將於 ${suspendDate.toLocaleString('zh-TW')} 實施停權`;

          if (isHourNotification) {
            body = `
親愛的使用者，

為因應國教署資安政策，本[離校/畢業]帳號 ${email} 將於 ${suspendDate.toLocaleString('zh-TW')} 停權。

⚠️ 這是停權前 ${hoursBeforeSuspend} 小時的最後提醒通知，請立即處理資料轉移事宜！

此信件為系統自動發送，請勿直接回覆。
            `;
          } else {
            body = `
親愛的使用者，

為因應國教署資安政策，本[離校/畢業]帳號 ${email} 將於 ${suspendDate.toLocaleString('zh-TW')} 停權。

這是停權前 ${weeksBeforeSuspend} 週的提醒通知，請儘速處理資料轉移事宜。

此信件為系統自動發送，請勿直接回覆。
            `;
          }

          GmailApp.sendEmail(email, subject, body);
          console.log(`✅ 通知信發送成功：${email} (停權前 ${timeInfo})`);
          sentCount++;

          // 在工作表中記錄發送狀態
          const currentStatus = sheet.getRange(row + 1, mailStatusColumnIndex + 1).getValue() || '';
          const newStatus = currentStatus ?
            `${currentStatus}; 已發送${timeInfo}前通知` :
            `已發送${timeInfo}前通知`;
          sheet.getRange(row + 1, mailStatusColumnIndex + 1).setValue(newStatus);
        } catch (err) {
          sheet.getRange(row + 1, errorColumnIndex + 1).setValue(err.message);
          console.log(`❌ 通知信發送失敗 (${email}): ${err.message}`);
        }
      }
    }

    console.log(`通知信觸發器執行完成，共發送 ${sentCount} 封信`);

    // 清理 Properties
    if (thisTriggerId) {
      PropertiesService.getScriptProperties().deleteProperty(`notification_trigger_${thisTriggerId}`);
    }

  } catch (error) {
    console.log('通知信觸發器執行發生錯誤:', error.message);
    console.log('錯誤詳細:', error.toString());
  }
}

/**
 * 處理單個使用者的資料
 * @param {Object} user - 使用者物件
 * @return {Array} 處理後的使用者資料陣列
 */
function processUserData(user) {
  var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
  var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
  var orgUnitPath = user.orgUnitPath || '/';

  // 取得 Employee ID
  var employeeId = 'N/A';
  if (user.externalIds && user.externalIds.length > 0) {
    for (var j = 0; j < user.externalIds.length; j++) {
      var externalId = user.externalIds[j];
      if (externalId.type === 'organization' || externalId.type === 'work') {
        employeeId = externalId.value;
        break;
      }
    }
  }

  // 取得 Employee Title 和 Department
  var employeeTitle = 'N/A';
  var department = 'N/A';
  if (user.organizations && user.organizations.length > 0) {
    for (var j = 0; j < user.organizations.length; j++) {
      var org = user.organizations[j];
      if (org.title) employeeTitle = org.title;
      if (org.department) department = org.department;
      if (employeeTitle !== 'N/A' && department !== 'N/A') break;
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

  return [
    user.primaryEmail, familyName, givenName, orgUnitPath,
    employeeId, employeeTitle, department, status,
    creationTime, lastLoginTime, '無需更新', ''
  ];
}

/**
 * 設定簡化的格式（避免逾時）
 * @param {Sheet} sheet - 工作表物件
 * @param {number} dataLength - 資料行數
 */
function setupSimpleFormatting(sheet, dataLength) {
  Logger.log('開始設定基本格式');
  
  // 設定欄位寬度
  var columnWidths = [80, 60, 60, 250, 60, 60, 60, 60, 60, 80, 80, 80];
  for (var col = 1; col <= columnWidths.length; col++) {
    sheet.setColumnWidth(col, columnWidths[col - 1]);
  }
  
  // 凍結標題行
  sheet.setFrozenRows(1);
  
  // 建立原始值參考區域（用於偵測變更）
  if (dataLength > 1) {
    var referenceStartRow = dataLength + 3;
    
    // 建立參考區域標題
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '', '']]; // 6欄對應B~G
    
    // 複製 B、C、D、E、F、G 欄的原始值到參考區域
    var originalData = sheet.getRange(2, 2, dataLength - 1, 6).getValues(); // 從第2行開始，取B~G欄（6欄）
    for (var i = 0; i < originalData.length; i++) {
      referenceData.push([
        originalData[i][0], // B欄：姓 (Family Name)
        originalData[i][1], // C欄：名 (Given Name)
        originalData[i][2], // D欄：機構單位路徑
        originalData[i][3], // E欄：Employee ID
        originalData[i][4], // F欄：Employee Title
        originalData[i][5]  // G欄：Department
      ]);
    }
    
    // 寫入參考區域
    sheet.getRange(referenceStartRow, 1, referenceData.length, 6).setValues(referenceData);
    
    // 隱藏參考區域
    if (referenceData.length > 1) {
      sheet.hideRows(referenceStartRow, referenceData.length);
    }
    
    Logger.log('參考區域建立完成，開始設定檢測公式');
    
    // 設定 K 欄的變更檢測公式（批次處理）`
    var batchSize = 500;
    for (var startRow = 2; startRow <= dataLength; startRow += batchSize) {
      var endRow = Math.min(startRow + batchSize - 1, dataLength);
      var detectionFormulas = [];
      
      for (var row = startRow; row <= endRow; row++) {
        var refRowIndex = referenceStartRow + (row - 1); // 對應的參考區域行號
        
        var detectionFormula = 
          '=IF(OR(' +
          'B' + row + '<>$A$' + refRowIndex + ',' +  // B欄：姓
          'C' + row + '<>$B$' + refRowIndex + ',' +  // C欄：名
          'D' + row + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
          'E' + row + '<>$D$' + refRowIndex + ',' +  // E欄：Employee ID
          'F' + row + '<>$E$' + refRowIndex + ',' +  // F欄：Employee Title
          'G' + row + '<>$F$' + refRowIndex +        // G欄：Department
          '),"需要更新","無需更新")';
        
        detectionFormulas.push([detectionFormula]);
      }
      
      if (detectionFormulas.length > 0) {
        sheet.getRange(startRow, 11, detectionFormulas.length, 1).setFormulas(detectionFormulas); // K欄（第11欄）
      }
    }
    
    Logger.log('檢測公式設定完成');
    
    // 設定 K 欄（是否需要更新）的條件格式
    var detectionRange = sheet.getRange(2, 11, dataLength - 1, 1); // K欄
    
    var needUpdateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("需要更新")
      .setBackground("#FFA500")  // 橘色背景
      .setFontColor("#FFFFFF")   // 白色文字
      .setRanges([detectionRange])
      .build();

    var noUpdateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("無需更新")
      .setBackground("#90EE90")  // 淺綠色背景
      .setFontColor("#000000")   // 黑色文字
      .setRanges([detectionRange])
      .build();

    var alreadyUpdatedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("已更新")
      .setBackground("#87CEEB")  // 淺藍色背景
      .setFontColor("#000000")   // 黑色文字
      .setRanges([detectionRange])
      .build();

    var rules = sheet.getConditionalFormatRules();
    rules.push(needUpdateRule);
    rules.push(noUpdateRule);
    rules.push(alreadyUpdatedRule);
    sheet.setConditionalFormatRules(rules);
    
    Logger.log('條件格式設定完成');
  }
  
  // 設定 L 欄在學狀態公式（批次處理，同時比對高中部、國中部、國小部的E欄和現職教師的C欄）
  if (dataLength > 1) {
    var batchSize = 500;
    for (var startRow = 2; startRow <= dataLength; startRow += batchSize) {
      var endRow = Math.min(startRow + batchSize - 1, dataLength);
      var formulas = [];
      
      for (var row = startRow; row <= endRow; row++) {
        // 修改後的公式：前三個工作表比對Email(A欄對E欄)，現職教師比對姓名(C欄對C欄)
        var formula = '=IF(' +
          'NOT(ISNA(VLOOKUP(A' + row + ',\'高中部\'!E:E,1,FALSE))),' + // 如果在高中部找到(Email比對)
          '"高中部在學",' +
          'IF(' +
            'NOT(ISNA(VLOOKUP(A' + row + ',\'國中部\'!E:E,1,FALSE))),' + // 如果在國中部找到(Email比對)
            '"國中部在學",' +
            'IF(' +
              'NOT(ISNA(VLOOKUP(A' + row + ',\'國小部\'!E:E,1,FALSE))),' + // 如果在國小部找到(Email比對)
              '"國小部在學",' +
              'IF(' +
                'NOT(ISNA(VLOOKUP(C' + row + ',\'114學年全校教職員工對照表\'!C:C,1,FALSE))),' + // 如果在現職教師找到(姓名比對)
                '"114現職",' +
                '""' + // 四個都沒找到就顯示空白
              ')' +
            ')' +
          ')' +
        ')';
        
        formulas.push([formula]);
      }
      
      if (formulas.length > 0) {
        sheet.getRange(startRow, 12, formulas.length, 1).setFormulas(formulas);
      }
    }
  }
  
  // 設定標題行格式
  var headerRange = sheet.getRange(1, 1, 1, 12);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  Logger.log('基本格式設定完成');
}