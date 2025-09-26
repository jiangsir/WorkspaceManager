/**
 * 在試算表菜單中添加一個自定義菜單項。
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('管理帳號與群組')
    .addItem('1. 匯出[新建範本stu]sheet 範本', 'exportNewUserTemplate')
    .addItem('2. 依[新建範本tea]批次新增(不更動現有資料)', 'processUsersAndGroups_V2')
    .addSeparator()
    .addItem('1.匯出[群組成員] (互動式)', 'showGroupManagementSidebar')
    .addItem('2.依[群組成員]更新群組', 'updateGroupMembersFromSheet')
    .addSeparator()
    .addItem('匯出所有機構單位 (含人數)', 'exportOUsAndUserCounts')
    .addToUi();
}

/**
 * 匯出近一年新增的使用者作為新建範本。
 * 包含所需的欄位格式，方便批次處理新使用者資料。
 */
function exportNewUserTemplate() {
  var ui = SpreadsheetApp.getUi();

  // 第一層確認
  var confirmation = ui.alert(
    '匯出新建範本',
    '此功能將匯出近一年新增的使用者作為範本，並包含批次處理所需的欄位格式。\n\n' +
    '匯出欄位包含：\n' +
    '• Email、姓、名\n' +
    '• 密碼、機構路徑\n' +
    '• Employee ID(真實姓名)、Employee Title(部別領域)、Department(註解)\n' +
    '• 建立日期、所屬群組 (群組 Email)\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (confirmation != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.showSidebar(HtmlService.createHtmlOutput('<b>正在讀取近一年新增的使用者資料，請稍候...</b>').setTitle('處理中'));

  var logMessages = ['開始讀取近一年新增的使用者...'];
  var recentUsers = [];
  var processedCount = 0;

  // 計算一年前的日期
  var oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  var oneYearAgoISO = oneYearAgo.toISOString();

  try {
    // 步驟 1: 獲取所有使用者並篩選近一年新增的
    var pageToken;
    do {
      var page = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        fields: 'nextPageToken,users(primaryEmail,name,orgUnitPath,organizations,creationTime,externalIds)'
      });

      if (page.users) {
        // 篩選近一年新增的使用者
        for (var i = 0; i < page.users.length; i++) {
          var user = page.users[i];
          if (user.creationTime && user.creationTime >= oneYearAgoISO) {
            recentUsers.push(user);
          }
        }
        processedCount += page.users.length;
        logMessages.push('已處理 ' + processedCount + ' 位使用者，找到 ' + recentUsers.length + ' 位近一年新增的使用者...');
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    if (recentUsers.length === 0) {
      ui.alert('結果', '近一年內未找到任何新增的使用者。', ui.ButtonSet.OK);
      return;
    }

    logMessages.push('篩選完成，共找到 ' + recentUsers.length + ' 位近一年新增的使用者，開始整理資料...');

    // 步驟 2: 準備要寫入工作表的資料
    var outputData = [[
      'Email',
      '姓',
      '名',
      '密碼',
      '機構路徑',
      'Employee ID(真實姓名)',
      'Employee Title(部別領域)',
      'Department(註解)',
      '建立日期',
      '所屬群組'
    ]];

    // 步驟 3: 處理每位使用者的資料並獲取群組資訊
    for (var i = 0; i < recentUsers.length; i++) {
      var user = recentUsers[i];

      var givenName = (user.name && user.name.givenName) ? user.name.givenName : '';
      var familyName = (user.name && user.name.familyName) ? user.name.familyName : '';
      var email = user.primaryEmail || '';
      var orgUnitPath = user.orgUnitPath || '/';

      // 獲取 Employee ID（從 externalIds 中提取）
      var employeeId = '';
      if (user.externalIds && user.externalIds.length > 0) {
        for (var j = 0; j < user.externalIds.length; j++) {
          if (user.externalIds[j].type === 'organization' || user.externalIds[j].type === 'custom') {
            employeeId = user.externalIds[j].value;
            break;
          }
        }
      }

      // 獲取 Employee Title 和 Department
      var employeeTitle = '';
      var department = '';
      if (user.organizations && user.organizations.length > 0) {
        for (var j = 0; j < user.organizations.length; j++) {
          var org = user.organizations[j];
          if (org.title) {
            employeeTitle = org.title;
          }
          if (org.department) {
            department = org.department;
          }
          if (employeeTitle && department) break;
        }
      }

      // 處理建立日期
      var creationTime = 'N/A';
      if (user.creationTime) {
        var createdDate = new Date(user.creationTime);
        creationTime = createdDate.toLocaleString('zh-TW', { timeZone: Session.getScriptTimeZone() });
      }

      // 獲取使用者所屬的群組 Email
      var userGroups = '';
      try {
        var memberGroupEmails = [];
        var groupPageToken;
        do {
          var groupPage = AdminDirectory.Groups.list({
            userKey: email,
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
        userGroups = '';
        logMessages.push('無法獲取使用者 ' + email + ' 的群組資訊: ' + groupError.message);
      }

      // 將資料加入到輸出陣列中
      outputData.push([
        email,                        // Email
        familyName,                   // 姓
        givenName,                    // 名
        '',                           // 密碼（預設空白）
        orgUnitPath,                  // 機構路徑
        employeeId,                   // Employee ID(真實姓名)
        employeeTitle,                // Employee Title(部別領域)
        department,                   // Department(註解)
        creationTime,                 // 建立日期
        userGroups                    // 所屬群組
      ]);

      // 每處理10位使用者就稍作暫停，避免API速率限制
      if (i % 10 === 9) {
        Utilities.sleep(100);
      }
    }

    // 步驟 4: 建立新工作表並寫入資料
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    var sheetName = "[新建範本]_" + timestamp;

    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(sheetName, 0);

    // 寫入資料
    newSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 步驟 5: 設定格式
    var columnWidths = {
      1: 60,  // Email
      2: 60,   // 姓
      3: 60,   // 名
      4: 60,   // 密碼
      5: 200,  // 機構路徑
      6: 60,  // Employee ID(真實姓名)
      7: 60,  // Employee Title(部別領域)
      8: 60,  // Department(註解)
      9: 60,  // 建立日期
      10: 100  // 所屬群組
    };

    // 設定固定欄位寬度
    for (var col = 1; col <= 10; col++) {
      if (columnWidths[col]) {
        newSheet.setColumnWidth(col, columnWidths[col]);
      }
    }

    // 設定標題行格式
    var headerRange = newSheet.getRange(1, 1, 1, 10);
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // 設定資料範圍的自動裁剪
    if (outputData.length > 1) {
      var dataRange = newSheet.getRange(2, 1, outputData.length - 1, 10);
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment('top');
      
      // 為密碼欄位設定淺色背景提示
      newSheet.getRange(2, 4, outputData.length - 1, 1).setBackground('#FFF9C4');
    }

    // 凍結標題行
    newSheet.setFrozenRows(1);

    newSheet.activate();

    ui.alert('匯出成功！', 
      '近一年新增的 ' + recentUsers.length + ' 位使用者資料已成功匯出至新的工作表 "' + sheetName + '"。\n\n' +
      '工作表包含所有批次處理所需的欄位格式，您可以：\n' +
      '1. 編輯「密碼」欄位來設定新密碼\n' +
      '2. 修改其他欄位資料\n' +
      '3. 查看「建立日期」了解帳號建立時間\n' +
      '4. 「所屬群組」欄位顯示群組 Email（便於批次處理）\n' +
      '5. 使用「依[新建範本tea]批次新增」功能進行批次處理', 
      ui.ButtonSet.OK);

  } catch (e) {
    var errorMsg = '處理過程中發生嚴重錯誤: ' + e.message;
    logMessages.push(errorMsg);
    ui.alert('錯誤', '無法完成新建範本匯出。\n\n錯誤詳情: ' + e.message, ui.ButtonSet.OK);
  } finally {
    Logger.log(logMessages.join('\n'));
    // 關閉側邊欄的 "處理中" 提示
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('<b>完成！</b>').setTitle('進度'));
  }
}
