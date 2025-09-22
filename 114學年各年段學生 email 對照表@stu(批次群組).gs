/**
 * 這個函數會在試算表檔案被開啟時自動執行，
 * 並在工具列上建立一個名為「管理工具」的自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('管理工具')
    .addItem('1.匯出[全部@stu清單]', 'exportAllStudents')
    .addToUi();
}

/**
 * 匯出整個 stu 網域中的所有使用者資料到一個新的工作表。
 * 包含使用者的基本資訊、機構單位、最後登入時間等詳細資訊。
 * 針對學生版本優化，移除群組資訊以加速處理。
 */
function exportAllStudents() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '匯出所有學生',
    '您即將匯出整個 stu 網域的所有學生清單。\n\n此操作可能需要較長時間，確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在讀取所有學生資料，這可能需要幾分鐘時間，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始讀取所有學生...'];
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
        logMessages.push('已讀取 ' + processedCount + ' 位學生...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (allUsers.length === 0) {
      ui.alert('結果', '未找到任何使用者。', ui.ButtonSet.OK);
      return;
    }

    logMessages.push('使用者資料讀取完成，共 ' + allUsers.length + ' 位學生，開始整理資料...');

    // 步驟 2: 準備要寫入工作表的資料（移除「所屬群組」欄位以加速處理）
    var outputData = [[
      'Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      'Employee ID(真實姓名)',
      'Employee Title(部別領域)',
      'Department(註解)',
      '帳號狀態',
      '建立時間',
      '最後登入時間',
      '是否需要更新',
      '現職狀態'              // ← L欄現職狀態
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
        employeeId,      // E欄：Employee ID
        employeeTitle,   // F欄：Employee Title
        department,      // G欄：Department
        status,          // H欄：帳號狀態
        creationTime,    // I欄：建立時間
        lastLoginTime,   // J欄：最後登入時間
        '無需更新',       // K欄：是否需要更新
        ''               // L欄：現職狀態，先填空值，稍後設定公式
      ]);

      // 顯示進度（每處理 100 位學生顯示一次，因為學生人數較多）
      if ((i + 1) % 100 === 0) {
        logMessages.push('已處理 ' + (i + 1) + '/' + allUsers.length + ' 位學生的資料...');
      }
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '_');
    var sheetName = "[全部@stu清單]";

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 先寫入資料（不包含公式）
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 設定 L 欄的現職狀態公式（假設參照的工作表名稱為學生版本）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var statusFormula = '=IF(ISNA(VLOOKUP(A' + rowIndex + ',\'114學年各年段學生對照表\'!F:F,1,FALSE)),"","在學")';
      newSheet.getRange(rowIndex, 12).setFormula(statusFormula); // L欄（第12欄）
    }

    // 步驟 5: 在工作表底部建立原始值參考區域（調整為7欄，移除群組欄位）
    var referenceStartRow = outputData.length + 3;
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '', '', '']]; // 7欄標題

    // 複製 B、C、D、E、F、G 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) { // 從第2行開始（跳過標題）
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4], // E欄：Employee ID
        outputData[i][5], // F欄：Employee Title
        outputData[i][6], // G欄：Department
        ''               // 第7欄：留空以配合標題行的7欄
      ]);
    }

    // 寫入參考區域（修正：改為7欄）
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 7).setValues(referenceData);

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E、F、G 欄的變化，移除群組欄位）
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
      1: 60,   // A欄：使用者 Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 350,  // D欄：機構單位路徑
      5: 60,   // E欄：Employee ID
      6: 60,   // F欄：Employee Title
      7: 60,   // G欄：Department
      8: 50,   // H欄：帳號狀態
      9: 60,   // I欄：建立時間
      10: 80,  // J欄：最後登入時間
      11: 60,  // K欄：是否需要更新
      12: 60   // L欄：現職狀態
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

    ui.alert('匯出成功！', allUsers.length + ' 位學生的資料已成功匯出至新的工作表 "' + sheetName + '"。', ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法完成學生匯出。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉側邊欄的 "處理中" 提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}