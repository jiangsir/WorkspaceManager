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
    .addItem('2.依據[全部@tea清單] 更新 B,C,D,E,F,G,H 欄位內容', 'updateUsersFromSheet')
    .addSeparator()
    .addItem('匯出[預約停權]範本', 'exportSuspensionTemplate')
    .addItem('--1.依據"停權時間"啟動停權程序', 'scheduleCompleteSuspensionProcess')
    .addItem('--2.列出本工作表內所有觸發器', 'listAllTriggers')
    .addItem('--3.清理本工作表內所有觸發器', 'cleanAllSuspensionTriggers')
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

    // 步驟 2: 準備要寫入工作表的資料（新增「所屬群組」欄位和「現職狀態」欄位）
    var outputData = [[
      'Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      '所屬群組',              // ← 新增：在 D、E 欄之間插入
      'Employee ID(真實姓名)',
      'Employee Title(部別領域)',
      'Department(註解)',
      '帳號狀態',
      '建立時間',
      '最後登入時間',
      '是否需要更新',
      '現職狀態'              // ← 新增：M欄現職狀態
    ]];

    // 步驟 3: 處理每位使用者的資料
    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];

      var familyName = (user.name && user.name.familyName) ? user.name.familyName : 'N/A';
      var givenName = (user.name && user.name.givenName) ? user.name.givenName : 'N/A';
      var orgUnitPath = user.orgUnitPath || '/';

      // 新增：取得使用者所屬的所有群組 Email
      var userGroups = '';  // 修改：改為空字串，無群組時顯示空白
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

        userGroups = memberGroupEmails.length > 0 ? memberGroupEmails.join(', ') : '';  // 修改：無群組時留空白
      } catch (groupError) {
        userGroups = '無法獲取';
        Logger.log('無法獲取使用者 ' + user.primaryEmail + ' 的群組資訊: ' + groupError.message);
      }

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
        userGroups,      // ← 修改：E欄顯示所屬群組（群組 Email），無群組時留空白
        employeeId,      // ← 原 E欄變成 F欄
        employeeTitle,   // ← 原 F欄變成 G欄
        department,      // ← 原 G欄變成 H欄
        status,          // ← 原 H欄變成 I欄
        creationTime,    // ← 原 I欄變成 J欄
        lastLoginTime,   // ← 原 J欄變成 K欄
        '無需更新',       // ← 原 K欄變成 L欄
        ''               // ← 新增：M欄現職狀態，先填空值，稍後設定公式
      ]);

      // 顯示進度（每處理 50 位使用者顯示一次）
      if ((i + 1) % 50 === 0) {
        logMessages.push('已處理 ' + (i + 1) + '/' + allUsers.length + ' 位使用者的群組資訊...');
      }
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

    // 設定 M 欄的現職狀態公式
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var statusFormula = '=IF(ISNA(VLOOKUP(A' + rowIndex + ',\'114學年全校教職員工對照表\'!F:F,1,FALSE)),"","現職")';
      newSheet.getRange(rowIndex, 13).setFormula(statusFormula); // M欄（第13欄）
    }

    // 步驟 5: 在工作表底部建立原始值參考區域（修正欄數）
    var referenceStartRow = outputData.length + 3;
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '', '', '', '']]; // 修正：8欄標題

    // 複製 B、C、D、E、F、G、H 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) { // 從第2行開始（跳過標題）
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4], // E欄：所屬群組（新增）
        outputData[i][5], // F欄：Employee ID
        outputData[i][6], // G欄：Employee Title
        outputData[i][7], // H欄：Department
        ''               // 第8欄：留空以配合標題行的8欄
      ]);
    }

    // 寫入參考區域（修正：改為8欄）
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 8).setValues(referenceData);

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E、F、G、H 欄的變化）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var refRowIndex = referenceStartRow + (rowIndex - 1); // 對應的參考行

      var detectionFormula =
        '=IF(OR(' +
        'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
        'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
        'D' + rowIndex + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
        'E' + rowIndex + '<>$D$' + refRowIndex + ',' +  // E欄：所屬群組（新增）
        'F' + rowIndex + '<>$E$' + refRowIndex + ',' +  // F欄：Employee ID
        'G' + rowIndex + '<>$F$' + refRowIndex + ',' +  // G欄：Employee Title
        'H' + rowIndex + '<>$G$' + refRowIndex +        // H欄：Department
        '),"需要更新","無需更新")';

      newSheet.getRange(rowIndex, 12).setFormula(detectionFormula); // L欄（第12欄）
    }

    // 步驟 7: 設定格式（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 60,   // A欄：使用者 Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 350,  // D欄：機構單位路徑
      5: 150,  // E欄：所屬群組（新增，較寬以容納多個群組 Email）
      6: 60,   // F欄：Employee ID
      7: 60,   // G欄：Employee Title
      8: 60,   // H欄：Department
      9: 50,   // I欄：帳號狀態
      10: 60,  // J欄：建立時間
      11: 80,  // K欄：最後登入時間
      12: 60,  // L欄：是否需要更新
      13: 60   // M欄：現職狀態
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 13; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定所有資料範圍的自動裁剪（文字換行）
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 13);
      dataRange.setWrap(true); // 啟用自動換行以適應固定寬度
      dataRange.setVerticalAlignment('top'); // 垂直對齊頂部
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
 * 讀取目前工作表中的資料，並更新對應使用者的 orgUnitPath、Employee ID、Employee Title、Department 和群組歸屬。
 * 只處理 L 欄標記為「需要更新」的行。
 */
function updateUsersFromSheet() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '更新使用者資訊',
    '此功能將讀取目前工作表的資料，並更新使用者的姓名、機構單位路徑、員工編號、職稱、部門和群組歸屬。\n\n' +
    '★ 智能更新：只會處理 L 欄標記為「需要更新」的使用者。\n' +
    '★ 可更新欄位：B(姓)、C(名)、D(機構單位)、E(所屬群組)、F(員工編號)、G(職稱)、H(部門)\n\n' +
    '請確認：\n' +
    '1. 目前工作表包含正確的使用者資料\n' +
    '2. 資料格式正確\n' +
    '3. 您已經手動修改了需要更新的資料\n\n' +
    '⚠️ 注意：群組更新會完全替換使用者的群組歸屬！\n\n' +
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

  // 查找各欄位的索引（對應新的欄位順序）
  var emailCol = headers.indexOf('Email');                        // A欄
  var familyNameCol = headers.indexOf('姓 (Family Name)');        // B欄
  var givenNameCol = headers.indexOf('名 (Given Name)');          // C欄
  var orgUnitPathCol = headers.indexOf('機構單位路徑');            // D欄
  var groupsCol = headers.indexOf('所屬群組');                    // E欄 (新增)
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
    ui.alert('錯誤', '找不到任何可更新的欄位。請確保工作表包含至少其中一个欄位。', ui.ButtonSet.OK);
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
        '所有使用者的 L 欄都顯示「無需更新」，或沒有有效的 Email。' :
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
          logMessages.push(logPrefix + '群組欄位為空，使用者現在不屬於任何群組。');
        }
      }

      // 執行使用者資料更新
      if (needsUserUpdate) {
        AdminDirectory.Users.update(userObj, email);
        logMessages.push(logPrefix + '使用者基本資料已成功更新。');
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
    if (trigger.getHandlerFunction() === 'processNewUsers' || trigger.getHandlerFunction() === 'exportAllUsers' || trigger.getHandlerFunction() === 'updateUsersFromSheet' || trigger.getHandlerFunction() === 'exportSuspensionTemplate') {
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
  var triggerList = [];

  // 收集所有觸發器的資訊
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    triggerList.push([
      trigger.getHandlerFunction(),
      trigger.getEventType(),
      trigger.getTriggerSource(),
      trigger.getTriggerSourceId()
    ]);
  }

  // 如果沒有觸發器，顯示提示訊息
  if (triggerList.length === 0) {
    ui.alert('目前沒有任何觸發器。');
    return;
  }

  // 顯示觸發器資訊
  var htmlOutput = HtmlService.createHtmlOutput('<h3>當前試算表觸發器列表</h3>');
  htmlOutput.append('<table style="width:100%; border-collapse:collapse;">');
  htmlOutput.append('<tr style="background-color:#f2f2f2; font-weight:bold;">');
  htmlOutput.append('<th style="border:1px solid #ddd; padding:8px;">處理函數</th>');
  htmlOutput.append('<th style="border:1px solid #ddd; padding:8px;">事件類型</th>');
  htmlOutput.append('<th style="border:1px solid #ddd; padding:8px;">觸發來源</th>');
  htmlOutput.append('<th style="border:1px solid #ddd; padding:8px;">來源 ID</th>');
  htmlOutput.append('</tr>');

  for (var i = 0; i < triggerList.length; i++) {
    var trigger = triggerList[i];
    htmlOutput.append('<tr>');
    htmlOutput.append('<td style="border:1px solid #ddd; padding:8px;">' + trigger[0] + '</td>');
    htmlOutput.append('<td style="border:1px solid #ddd; padding:8px;">' + trigger[1] + '</td>');
    htmlOutput.append('<td style="border:1px solid #ddd; padding:8px;">' + trigger[2] + '</td>');
    htmlOutput.append('<td style="border:1px solid #ddd; padding:8px;">' + trigger[3] + '</td>');
    htmlOutput.append('</tr>');
  }

  htmlOutput.append('</table>');
  htmlOutput.append('<p>總共 ' + triggerList.length + ' 個觸發器。</p>');

  ui.showModalDialog(htmlOutput, '觸發器列表');
}

/**
 * 匯出機構單位路徑為 "/離職人員" 的使用者到新工作表
 */
function exportSuspensionTemplate() {
  var ui = SpreadsheetApp.getUi();

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在匯出離職人員清單，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始匯出離職人員清單...'];

  try {
    // 步驟 1: 先獲取所有使用者，然後篩選出機構單位路徑為 "/離職人員" 的使用者
    var retiredUsers = [];
    var processedCount = 0;
    var totalCount = 0;

    logMessages.push('正在讀取所有使用者資料並篩選離職人員...');

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
        
        // 篩選出機構單位路徑為 "/離職人員" 的使用者
        for (var i = 0; i < page.users.length; i++) {
          var user = page.users[i];
          if (user.orgUnitPath === '/離職人員') {
            retiredUsers.push(user);
            processedCount++;
          }
        }
        
        logMessages.push('已掃描 ' + totalCount + ' 位使用者，找到 ' + processedCount + ' 位離職人員...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (retiredUsers.length === 0) {
      ui.alert(
        '結果', 
        '未找到任何機構單位路徑為 "/離職人員" 的使用者。\n\n' +
        '已掃描總使用者數：' + totalCount + '\n' +
        '找到離職人員數：0\n\n' +
        '請確認：\n' +
        '1. 機構單位 "/離職人員" 是否存在\n' +
        '2. 是否有使用者被分配到此機構單位', 
        ui.ButtonSet.OK
      );
      return;
    }

    logMessages.push('使用者掃描完成，總共掃描 ' + totalCount + ' 位使用者，找到 ' + retiredUsers.length + ' 位離職人員，開始整理資料...');

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

    // 步驟 3: 處理每位離職人員的資料
    for (var i = 0; i < retiredUsers.length; i++) {
      var user = retiredUsers[i];

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

      // 顯示進度（每處理 10 位使用者顯示一次）
      if ((i + 1) % 10 === 0 || i === retiredUsers.length - 1) {
        logMessages.push('已處理 ' + (i + 1) + '/' + retiredUsers.length + ' 位離職人員的資料...');
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
      ['=== 離職人員統計資訊 ==='],
      [''],
      ['掃描範圍：全部使用者 (' + totalCount + ' 位)'],
      ['總離職人員數：' + (outputData.length - 1)],
      ['啟用中帳號：' + activeCount],
      ['已停用帳號：' + suspendedCount],
      [''],
      ['匯出時間：' + new Date().toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() })],
      ['篩選條件：機構單位路徑 = "/離職人員"']
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

    logMessages.push('離職人員清單匯出完成！共包含 ' + (outputData.length - 1) + ' 位離職人員。');

    ui.alert(
      '匯出成功！', 
      '離職人員清單已成功匯出，共包含 ' + (outputData.length - 1) + ' 位離職人員。\n\n' +
      '掃描統計：\n' +
      '• 總掃描使用者：' + totalCount + ' 位\n' +
      '• 找到離職人員：' + (outputData.length - 1) + ' 位\n' +
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
    var errorMsg = '匯出離職人員清單時發生錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法匯出離職人員清單。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
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
          // JSON 解析失敗但仍屬於該工作表的觸發器，也刪除
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          console.log(`刪除工作表 ${sheetName} 的損壞停權觸發器（UID=${trig.getUniqueId()}）`);
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

          subject = `[信箱停用通知] 因應國教署資安政策，離職/畢業帳號停權通知 - 本帳號預計將於 ${suspendDate.toLocaleString('zh-TW')} 實施停權`;

          if (isHourNotification) {
            body = `
親愛的使用者，

為因應國教署資安政策，本[離職/畢業]帳號 ${email} 將於 ${suspendDate.toLocaleString('zh-TW')} 停權。

⚠️ 這是停權前 ${hoursBeforeSuspend} 小時的最後提醒通知，請立即處理資料轉移事宜！

此信件為系統自動發送，請勿直接回覆。
            `;
          } else {
            body = `
親愛的使用者，

為因應國教署資安政策，本[離職/畢業]帳號 ${email} 將於 ${suspendDate.toLocaleString('zh-TW')} 停權。

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
