function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('自訂工具') // 您可以自訂選單名稱
      .addItem('取代姓名中間字', 'replaceMiddleName') // '取代姓名中間字' 是選單項目名稱，'replaceMiddleName' 是您要執行的函數名稱
      .addSeparator()
      .addItem('1.匯出[全部@stu清單]', 'exportAllStudentUsers')
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
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,suspended,creationTime,lastLoginTime)'
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

    // 步驟 2: 準備要寫入工作表的資料（簡化版本）
    var outputData = [[
      'Email',
      '姓 (Family Name)',
      '名 (Given Name)',
      '機構單位路徑',
      '所屬群組',
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
        user.primaryEmail,
        familyName,
        givenName,
        orgUnitPath,
        userGroups,
        status,
        creationTime,
        lastLoginTime,
        '無需更新'
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
    var referenceData = [['=== 原始值參考區域（系統用，請勿修改）===', '', '', '', '']];

    // 複製 B、C、D、E 欄的原始值到參考區域
    for (var i = 1; i < outputData.length; i++) {
      referenceData.push([
        outputData[i][1], // B欄：姓 (Family Name)
        outputData[i][2], // C欄：名 (Given Name)  
        outputData[i][3], // D欄：機構單位路徑
        outputData[i][4], // E欄：所屬群組
        ''
      ]);
    }

    // 寫入參考區域
    newSheet.getRange(referenceStartRow, 1, referenceData.length, 5).setValues(referenceData);

    // 隱藏參考區域
    if (referenceData.length > 1) {
      newSheet.hideRows(referenceStartRow, referenceData.length);
    }

    // 步驟 6: 設定檢測公式（檢測 B、C、D、E 欄的變化）
    for (var rowIndex = 2; rowIndex <= outputData.length; rowIndex++) {
      var refRowIndex = referenceStartRow + (rowIndex - 1);

      var detectionFormula =
        '=IF(OR(' +
        'B' + rowIndex + '<>$A$' + refRowIndex + ',' +  // B欄：姓
        'C' + rowIndex + '<>$B$' + refRowIndex + ',' +  // C欄：名
        'D' + rowIndex + '<>$C$' + refRowIndex + ',' +  // D欄：機構單位路徑
        'E' + rowIndex + '<>$D$' + refRowIndex +        // E欄：所屬群組
        '),"需要更新","無需更新")';

      newSheet.getRange(rowIndex, 9).setFormula(detectionFormula); // I欄（第9欄）
    }

    // 步驟 7: 設定格式（固定寬度 + 自動裁剪內容）
    var columnWidths = {
      1: 60,   // A欄：學生 Email
      2: 60,   // B欄：姓 (Family Name)
      3: 60,   // C欄：名 (Given Name)
      4: 200,  // D欄：機構單位路徑
      5: 150,  // E欄：所屬群組
      6: 50,   // F欄：帳號狀態
      7: 60,   // G欄：建立時間
      8: 80,   // H欄：最後登入時間
      9: 60    // I欄：是否需要更新
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
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
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