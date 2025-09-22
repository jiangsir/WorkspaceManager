function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('自訂工具') // 您可以自訂選單名稱
      .addItem('取代姓名中間字', 'replaceMiddleName') // '取代姓名中間字' 是選單項目名稱，'replaceMiddleName' 是您要執行的函數名稱
      .addSeparator()
      .addItem('1.匯出[全部@stu清單]', 'exportAllStudentUsers')
      .addItem('2.依據[全部@stu清單] 更新 B,C,D,E,F,G,H 欄位內容', 'updateStudentsFromSheet')
      .addToUi();
}

function replaceMiddleName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // 設定要處理的範圍，從 D2 開始到 D 欄的最後一列
  var startRow = 2;
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(startRow, 4, lastRow - startRow + 1, 1);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var name = values[i][0];
    // 確保處理的是字串且長度大於等於 2
    if (name && typeof name == 'string' && name.length >= 2) {
      var len = name.length;
      
      if (len === 2) {
        // 【處理兩字的情況】
        // 取出第一個字，後面加上 "O"
        values[i][0] = name.substring(0, 1) + 'O';
      } else { // len > 2
        // 【處理三字以上的情況】
        var firstChar = name.substring(0, 1); // 取出第一個字
        var lastChar = name.substring(len - 1); // 取出最後一個字
        var middleCircles = 'O'.repeat(len - 2); // 根據中間字數，產生對應數量的 "O"
        
        values[i][0] = firstChar + middleCircles + lastChar;
      }
    }
  }
  range.setValues(values);
}

/**
 * 匯出整個 stu 網域中的所有學生使用者資料到一個新的工作表。
 * 包含學生的基本資訊、機構單位、最後登入時間等詳細資訊。
 */
function exportAllStudentUsers() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '匯出所有學生使用者',
    '您即將匯出整個 stu 網域的所有學生使用者清單。\n\n此操作可能需要較長時間，確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在讀取所有學生使用者資料，這可能需要幾分鐘時間，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始讀取所有學生使用者...'];
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
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,suspended,creationTime,lastLoginTime,externalIds,organizations,addresses)'
      });

      if (page.users) {
        // 篩選出學生使用者（Email 包含 @stu）
        var studentUsers = page.users.filter(function(user) {
          return user.primaryEmail && user.primaryEmail.includes('@stu');
        });
        
        allUsers = allUsers.concat(studentUsers);
        processedCount += page.users.length;
        logMessages.push('已讀取 ' + processedCount + ' 位使用者，找到 ' + allUsers.length + ' 位學生...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allUsers.length === 0) {
      ui.alert('結果', '未找到任何學生使用者。', ui.ButtonSet.OK);
      return;
    }

    logMessages.push('學生使用者資料讀取完成，共 ' + allUsers.length + ' 位學生，開始整理資料...');

    // 步驟 2: 準備要寫入工作表的資料（調整欄位順序）
    var outputData = [[
      'Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      '所屬群組',
      'Employee ID(真實姓名)',
      'Employee Title(部別領域)',
      'Department(註解)',
      '帳號狀態',
      '建立時間',
      '最後登入時間',
      '是否需要更新'
    ]];

    // 步驟 3: 處理每位學生的資料
    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];

      var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
      var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
      
      // 從 externalIds 中提取 Employee ID
      var employeeId = 'N/A';
      if (user.externalIds && user.externalIds.length > 0) {
        for (var j = 0; j < user.externalIds.length; j++) {
          if (user.externalIds[j].type === 'organization' || user.externalIds[j].type === 'custom') {
            employeeId = user.externalIds[j].value || 'N/A';
            break;
          }
        }
      }
      
      // 從 organizations 中提取職稱和部門
      var employeeType = 'N/A';
      var department = 'N/A';
      if (user.organizations && user.organizations.length > 0) {
        var org = user.organizations[0]; // 取第一個組織資訊
        employeeType = org.title || 'N/A';
        department = org.department || 'N/A';
      }
      
      var orgUnitPath = user.orgUnitPath || '/';

      // 取得學生所屬的所有群組 Email
      var userGroups = '';
      try {
        var memberGroupEmails = [];
        var groupPageToken;
        do {
          var groupPage = AdminDirectory.Groups.list({
            userKey: user.primaryEmail,
            maxResults: 200,
            pageToken: groupPageToken,
            fields: 'nextPageToken,groups(email)'
          });
          if (groupPage.groups) {
            for (var g = 0; g < groupPage.groups.length; g++) {
              memberGroupEmails.push(groupPage.groups[g].email);
            }
          }
          groupPageToken = groupPage.nextPageToken;
        } while (groupPageToken);

        userGroups = memberGroupEmails.length > 0 ? memberGroupEmails.join(', ') : '';
      } catch (groupError) {
        userGroups = '無法獲取';
        Logger.log('無法獲取學生 ' + user.primaryEmail + ' 的群組資訊: ' + groupError.message);
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
        user.primaryEmail,        // A欄：Email
        familyName,              // B欄：姓 (Family Name)
        givenName,               // C欄：名 (Given Name)
        orgUnitPath,             // D欄：機構單位路徑
        userGroups,              // E欄：所屬群組
        employeeId,              // F欄：Employee ID(真實姓名)
        employeeType,            // G欄：Employee Title(部別領域)
        department,              // H欄：Department(註解)
        status,                  // I欄：帳號狀態
        creationTime,            // J欄：建立時間
        lastLoginTime,           // K欄：最後登入時間
        '無需更新'               // L欄：是否需要更新
      ]);

      // 顯示進度（每處理 50 位學生顯示一次）
      if ((i + 1) % 50 === 0) {
        logMessages.push('已處理 ' + (i + 1) + '/' + allUsers.length + ' 位學生的群組資訊...');
      }
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "[全部@stu清單]";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 步驟 5: 在工作表底部建立原始值參考區域
    var referenceStartRow = outputData.length + 3;
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '', '', '', '']];

    // 複製 B、C、D、E、F、G、H 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) {
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4], // E欄：所屬群組
        outputData[i][5], // F欄：Employee ID(真實姓名)
        outputData[i][6], // G欄：Employee Title(部別領域)
        outputData[i][7], // H欄：Department(註解)
        ''
      ]);
    }

    // 寫入參考區域
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 8).setValues(referenceData);

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E、F、G、H 欄的變化）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var refRowIndex = referenceStartRow + (rowIndex - 1);

      var detectionFormula =
        '=IF(OR(' +
        'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
        'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
        'D' + rowIndex + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
        'E' + rowIndex + '<>$D$' + refRowIndex + ',' +  // E欄：所屬群組
        'F' + rowIndex + '<>$E$' + refRowIndex + ',' +  // F欄：Employee ID
        'G' + rowIndex + '<>$F$' + refRowIndex + ',' +  // G欄：Employee Title
        'H' + rowIndex + '<>$G$' + refRowIndex +        // H欄：Department
        '),"需要更新","無需更新")';

      newSheet.getRange(rowIndex, 12).setFormula(detectionFormula); // L欄（第12欄）
    }

    // 步驟 7: 設定格式（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 80,  // A欄：學生 Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 200,  // D欄：機構單位路徑
      5: 150,  // E欄：所屬群組
      6: 60,   // F欄：Employee ID(真實姊名)
      7: 60,  // G欄：Employee Title(部別領域)
      8: 80,   // H欄：Department(註解)
      9: 50,   // I欄：帳號狀態
      10: 60,  // J欄：建立時間
      11: 80,  // K欄：最後登入時間
      12: 60   // L欄：是否需要更新
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 12; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定所有資料範圍的自動裁剪（文字換行）
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 12);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
    }

    newSheet.setFrozenRows(1); // 凍結標題行

    // 步驟 8: 設定「是否需要更新」欄位的條件格式
    if (outputData.length > 1) {
      var detectionRange = newSheet.getRange(2, 12, outputData.length - 1, 1); // L欄（第12欄）

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

    ui.alert('匯出成功！', allUsers.length + ' 位學生的資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法完成學生使用者匯出。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉側邊欄的 "處理中" 提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}

/**
 * 根據試算表中的資料更新學生的機構單位路徑和職稱。
 * 讀取目前工作表中的資料，並更新對應學生的 orgUnitPath、Employee ID、Employee Title、Department 和群組歸屬。
 * 只處理 L 欄標記為「需要更新」的行。
 */
function updateStudentsFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新學生資訊',
    '此功能將讀取目前工作表的資料，並更新學生的姓名、機構單位路徑、員工編號、職稱、部門和群組歸屬。\n\n' +
    '★ 智能更新：只會處理 L 欄標記為「需要更新」的學生。\n' +
    '★ 可更新欄位：B(姓)、C(名)、D(機構單位)、E(所屬群組)、F(員工編號)、G(職稱)、H(部門)\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的學生資料\n' +
    '2. 資料格式正確\n' +
    '3. 您已經手動修改了需要更新的資料\n\n' +
    '⚠️ 注意：群組更新會完全替換學生的群組歸屬！\n\n' +
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

  // 查找各欄位的索引（對應學生管理的欄位順序）
  var emailCol = headers.indexOf('Email');                        // A欄
  var familyNameCol = headers.indexOf('姓 (Family Name)');        // B欄
  var givenNameCol = headers.indexOf('名 (Given Name)');          // C欄
  var orgUnitPathCol = headers.indexOf('機構單位路徑');            // D欄
  var groupsCol = headers.indexOf('所屬群組');                    // E欄
  var employeeIdCol = headers.indexOf('Employee ID(真實姓名)');   // F欄
  var employeeTitleCol = headers.indexOf('Employee Title(部別領域)'); // G欄
  var departmentCol = headers.indexOf('Department(註解)');        // H欄
  var updateStatusCol = headers.indexOf('是否需要更新');           // L欄

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「Email」欄位。請確保工作表包含正確的標題。', ui.ButtonSet.OK);
    return;
  }

  if (familyNameCol === -1 && givenNameCol === -1 && orgUnitPathCol === -1 && groupsCol === -1 && employeeIdCol === -1 && employeeTitleCol === -1 && departmentCol === -1) {
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
        '所有學生的 L 欄都顯示「無需更新」，或沒有有效的 Email。' :
        '沒有找到有效的 Email。'),
      ui.ButtonSet.OK);
    return;
  }

  // 建立群組名稱到群組Email的對應表（保留以支援群組名稱格式）
  var groupNameToEmailMap = {};
  try {
    var allGroups = [];
    var pageToken;
    do {
      var page = AdminDirectory.Groups.list({
        customer: 'my_customer',
        maxResults: 200,
        pageToken: pageToken,
        fields: 'nextPageToken,groups(name,email)'
      });
      if (page.groups) {
        allGroups = allGroups.concat(page.groups);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    for (var g = 0; g < allGroups.length; g++) {
      groupNameToEmailMap[allGroups[g].name] = allGroups[g].email;
    }
  } catch (e) {
    Logger.log('建立群組對應表時發生錯誤: ' + e.message);
  }

  // 確認要處理的行數
  var confirmationFields = [];
  if (familyNameCol !== -1) confirmationFields.push('• 更新姓氏 (B欄)');
  if (givenNameCol !== -1) confirmationFields.push('• 更新名字 (C欄)');
  if (orgUnitPathCol !== -1) confirmationFields.push('• 更新機構單位路徑 (D欄)');
  if (groupsCol !== -1) confirmationFields.push('• 更新群組歸屬 (E欄)');
  if (employeeIdCol !== -1) confirmationFields.push('• 更新員工編號 (F欄)');
  if (employeeTitleCol !== -1) confirmationFields.push('• 更新職稱 (G欄)');
  if (departmentCol !== -1) confirmationFields.push('• 更新部門 (H欄)');

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
      // 檢查學生是否存在
      var user;
      try {
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail,name,orgUnitPath,organizations,externalIds" });
      } catch (e) {
        logMessages.push(logPrefix + '學生不存在，跳過。');
        skipCount++;
        continue;
      }

      var needsUserUpdate = false;
      var userObj = {};
      var needsGroupUpdate = false;

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
            if (externalId.type === 'organization' || externalId.type === 'custom') {
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

      // 處理群組更新
      if (groupsCol !== -1) {
        var newGroupsText = String(row[groupsCol] || '').trim();
        needsGroupUpdate = true;

        // 解析新的群組列表
        var newGroups = [];
        if (newGroupsText && newGroupsText !== '無群組' && newGroupsText !== 'N/A' && newGroupsText !== '無法獲取' && newGroupsText !== '不適用') {
          var groupIdentifiers = newGroupsText.split(',').map(function (identifier) { return identifier.trim(); });

          for (var j = 0; j < groupIdentifiers.length; j++) {
            var groupIdentifier = groupIdentifiers[j];
            if (groupIdentifier) {
              // 檢查是否為群組 Email 格式（包含 @ 符號）
              if (groupIdentifier.indexOf('@') !== -1) {
                // 直接使用群組 Email
                newGroups.push({
                  identifier: groupIdentifier,
                  email: groupIdentifier
                });
              } else if (groupNameToEmailMap[groupIdentifier]) {
                // 使用群組名稱查找對應的 Email
                newGroups.push({
                  identifier: groupIdentifier,
                  email: groupNameToEmailMap[groupIdentifier]
                });
              } else {
                logMessages.push(logPrefix + '警告：無法識別群組 "' + groupIdentifier + '"，將跳過此群組。');
              }
            }
          }
        }

        // 步驟 1: 獲取學生目前所屬的所有群組
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

        // 步驟 2: 從所有目前群組中移除該學生
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

        // 步驟 3: 將學生加入到新的群組中
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
              logMessages.push(logPrefix + '已是群組 "' + newGroups[k].identifier + '" 的成員。');
              addCount++; // 視為成功
            } else {
              addErrors++;
              logMessages.push(logPrefix + '加入群組 "' + newGroups[k].identifier + '" 時失敗: ' + addError.message);
            }
          }
        }

        if (newGroups.length > 0) {
          logMessages.push(logPrefix + '成功加入 ' + addCount + ' 個群組' + (addErrors > 0 ? '（失敗 ' + addErrors + ' 個）' : '') + '。');
        } else {
          logMessages.push(logPrefix + '群組欄位為空，學生現在不屬於任何群組。');
        }
      }

      // 執行學生資料更新
      if (needsUserUpdate) {
        AdminDirectory.Users.update(userObj, email);
        logMessages.push(logPrefix + '學生基本資料已成功更新。');
      }

      if (needsUserUpdate || needsGroupUpdate) {
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