/**
 * 這個函數會在試算表檔案被開啟時自動執行，
 * 並在工具列上建立一個名為「管理工具」的自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('管理工具')
      .addItem('處理新進員工帳號', 'processNewUsers')
      .addSeparator()
      .addItem('1.匯出[全部@tea清單]', 'exportAllUsers')
      .addItem('2.依據[全部@tea清單] 更新 B,C,D,E,F,G 欄位內容', 'updateUsersFromSheet')
      .addToUi();
}

/**
 * 這是主要的核心函數。
 * 它會讀取試算表中的資料，處理所有狀態為 "建立帳號" 的使用者。
 */
function processNewUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 取得第一行標題列，並動態找到各欄位的索引
  // 這樣做的好處是，就算您調整了試算表欄位的順序，程式碼也不需要修改
  const headers = values[0]; 
  const firstNameIndex = headers.indexOf("FirstName");
  const lastNameIndex = headers.indexOf("LastName");
  const emailIndex = headers.indexOf("Email");
  const passwordIndex = headers.indexOf("Password");
  const groupIndex = headers.indexOf("GroupEmail");
  const statusIndex = headers.indexOf("Status");

  // 從第二行 (索引為 1) 開始遍歷所有資料
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[statusIndex];
    const userEmail = row[emailIndex];

    // 如果狀態是 "建立帳號"，並且 Email 欄位不是空的，就開始處理
    if (status === '建立帳號' && userEmail) {
      
      // 使用 try...catch 結構來處理錯誤
      // 這樣可以確保即使某一個帳號建立失敗，程式也會繼續處理下一個，不會中斷
      try {
        // 1. 準備要建立的使用者物件
        const newUser = {
          primaryEmail: userEmail,
          name: {
            givenName: row[firstNameIndex],
            familyName: row[lastNameIndex]
          },
          password: row[passwordIndex],
          changePasswordAtNextLogin: true // 強制使用者下次登入時變更密碼 (安全性考量)
        };

        // 呼叫 Admin SDK API 來新增使用者
        AdminDirectory.Users.insert(newUser);
        Logger.log(`成功建立使用者: ${userEmail}`);

        // 2. 準備將使用者加入群組
        const groupEmail = row[groupIndex];
        // 檢查 GroupEmail 欄位是否有填寫
        if (groupEmail) {
          const member = {
            email: userEmail,
            role: 'MEMBER'
          };
          // 呼叫 Admin SDK API 將成員加入群組
          AdminDirectory.Members.insert(member, groupEmail);
          Logger.log(`成功將 ${userEmail} 加入群組 ${groupEmail}`);
        }

        // 3. 所有操作成功後，在試算表中回寫狀態為 "已完成"
        sheet.getRange(i + 1, statusIndex + 1).setValue('已完成');
        
      } catch (err) {
        // 如果在 try 區塊中發生任何錯誤 (例如帳號已存在)
        Logger.log(`處理 ${userEmail} 時發生錯誤: ${err.toString()}`);
        // 將錯誤訊息回寫到狀態欄，方便管理者查看
        sheet.getRange(i + 1, statusIndex + 1).setValue(`錯誤: ${err.message}`);
      }
    }
  }
  // 所有處理完成後，跳出一個提示視窗
  SpreadsheetApp.getUi().alert('批次處理完成！請檢查 Status 欄位的結果。');
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
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,organizations,externalIds,suspended,creationTime,lastLoginTime)'
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
      'Employee ID',
      'Employee Title',     // ← 調換順序：Employee Title 在前
      'Department',         // ← 調換順序：Department 在後
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
          if (org.title) {
            employeeTitle = org.title;
          }
          if (org.department) {
            department = org.department;
          }
          // 如果都找到了就跳出循環
          if (employeeTitle !== 'N/A' && department !== 'N/A') {
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
        employeeId,
        employeeTitle,   // ← F欄：Employee Title 在前
        department,      // ← G欄：Department 在後
        status,
        creationTime,
        lastLoginTime,
        '無需更新'
      ]);
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '_');
    var sheetName = "[全部@tea清單]";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 先寫入資料（不包含公式）
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 步驟 5: 在工作表底部建立原始值參考區域
    var referenceStartRow = outputData.length + 3;
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '', '']];

    // 複製 B、C、D、E、F、G 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) { // 從第2行開始（跳過標題）
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4], // E欄：Employee ID
        outputData[i][5], // F欄：Employee Title（調換後）
        outputData[i][6]  // G欄：Department（調換後）
      ]);
    }

    // 寫入參考區域
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 6).setValues(referenceData);

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E、F、G 欄的變化）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var refRowIndex = referenceStartRow + (rowIndex - 1); // 對應的參考行

      var detectionFormula =
        '=IF(OR(' +
        'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
        'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
        'D' + rowIndex + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
        'E' + rowIndex + '<>$D$' + refRowIndex + ',' +  // E欄：Employee ID
        'F' + rowIndex + '<>$E$' + refRowIndex + ',' +  // F欄：Employee Title
        'G' + rowIndex + '<>$F$' + refRowIndex +        // G欄：Department
        '),"需要更新","無需更新")';

      newSheet.getRange(rowIndex, 11).setFormula(detectionFormula); // K欄（第11欄）
    }

    // 步驟 7: 設定格式（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 60,  // A欄：使用者 Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 300,  // D欄：機構單位路徑
      5: 60,  // E欄：Employee ID
      6: 60,  // F欄：Employee Title
      7: 60,  // G欄：Department
      8: 50,   // H欄：帳號狀態
      9: 60,  // I欄：建立時間
      10: 80, // J欄：最後登入時間
      11: 60  // K欄：是否需要更新
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 11; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定所有資料範圍的自動裁剪（文字換行）
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 11);
      dataRange.setWrap(true); // 啟用自動換行以適應固定寬度
      dataRange.setVerticalAlignment('top'); // 垂直對齊頂部
    }

    newSheet.setFrozenRows(1); // 凍結標題行

    // 步驟 8: 設定「是否需要更新」欄位的條件格式
    if (outputData.length > 1) {
      var detectionRange = newSheet.getRange(2, 11, outputData.length - 1, 1); // K欄（第11欄）

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
 * 讀取目前工作表中的資料，並更新對應使用者的 orgUnitPath、Employee ID、Employee Title 和 Department。
 * 只處理 K 欄標記為「需要更新」的行。
 */
function updateUsersFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新使用者資訊',
    '此功能將讀取目前工作表的資料，並更新使用者的姓名、機構單位路徑、員工編號、職稱和部門。\n\n' +
    '★ 智能更新：只會處理 K 欄標記為「需要更新」的使用者。\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的使用者資料\n' +
    '2. 資料格式正確（包含 Email、姓、名、機構單位路徑、Employee ID、Employee Title、Department 欄位）\n' +
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

  // 查找各欄位的索引（調整順序後）
  var emailCol = headers.indexOf('使用者 Email');
  var familyNameCol = headers.indexOf('姓 (Family Name)');
  var givenNameCol = headers.indexOf('名 (Given Name)');
  var orgUnitPathCol = headers.indexOf('機構單位路徑');
  var employeeIdCol = headers.indexOf('Employee ID');
  var employeeTitleCol = headers.indexOf('Employee Title');  // ← 現在是 F 欄
  var departmentCol = headers.indexOf('Department');         // ← 現在是 G 欄
  var updateStatusCol = headers.indexOf('是否需要更新'); // 檢測欄位的索引

  // 檢查必要欄位是否存在
  if (emailCol === -1) {
    ui.alert('錯誤', '找不到「使用者 Email」欄位。請確保工作表包含正確的標題。', ui.ButtonSet.OK);
    return;
  }

  if (familyNameCol === -1 && givenNameCol === -1 && orgUnitPathCol === -1 && employeeIdCol === -1 && employeeTitleCol === -1 && departmentCol === -1) {
    ui.alert('錯誤', '找不到任何可更新的欄位（姓、名、機構單位路徑、Employee ID、Employee Title、Department）。請確保工作表包含至少其中一個欄位。', ui.ButtonSet.OK);
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
        '所有使用者的 K 欄都顯示「無需更新」，或沒有有效的 Email。' :
        '沒有找到有效的 Email。'),
      ui.ButtonSet.OK);
    return;
  }

  // 確認要處理的行數
  var confirmationFields = [];
  if (familyNameCol !== -1) confirmationFields.push('• 更新姓氏');
  if (givenNameCol !== -1) confirmationFields.push('• 更新名字');
  if (orgUnitPathCol !== -1) confirmationFields.push('• 更新機構單位路徑');
  if (employeeIdCol !== -1) confirmationFields.push('• 更新員工編號');
  if (employeeTitleCol !== -1) confirmationFields.push('• 更新職稱資訊');
  if (departmentCol !== -1) confirmationFields.push('• 更新部門');

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
        user = AdminDirectory.Users.get(email, { fields: "primaryEmail,name,orgUnitPath,organizations,externalIds" });
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

      // 處理 Employee ID 更新
      if (employeeIdCol !== -1) {
        var newEmployeeId = String(row[employeeIdCol] || '').trim();
        
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
          needsUpdate = true;
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
        needsUpdate = true;
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