const emailColumnIndex = 2; // 假設 email 在第 3 欄（C 欄）
const timeColumnIndex = 5; // 假設時間在第 6 欄（F 欄）
const statusColumnIndex = 6; // 假設狀態在第 7 欄（G 欄）
const errorColumnIndex = 7; // 假設錯誤訊息在第 8 欄（H 欄）
const MailStatusColumnIndex = 8; // 假設郵件狀態在第 9 欄（I 欄）


/**
 * onOpen 時加入自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('使用者管理工具')
    .addItem('匯出範本', 'xxxxxx')
    .addSeparator()
    .addItem('寄發本工作表內的連續通知信', 'scheduleNotificationEmails')
    .addItem('建立/更新本工作表內的停權觸發器', 'scheduleSuspendUsersByTime')
    .addSeparator()
    .addItem('立即停權本工作表所有使用者', 'SuspendAllUser')
    .addItem('清理本工作表所有觸發器', 'cleanAllTriggers')
    .addSeparator()
    .addItem('列出所有觸發器', 'xxxxxx')
    .addToUi();
}

/**
 * 掃描所有帳號與停權時間，建立觸發器
 */
function scheduleSuspendUsersByTime() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const futureTimes = new Set(); // 改用 Set 來收集所有不同的未來時間

  // 1️⃣ 掃描所有列，收集不同的未來時間點
  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
    const timeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
    if (!email || !timeStr) continue;

    const date = new Date(timeStr);
    if (isNaN(date.getTime())) {
      sheet.getRange(row + 1, errorColumnIndex + 1).setValue('時間格式錯誤');
      continue;
    }

    if (date <= now) {
      sheet.getRange(row + 1, errorColumnIndex + 1).setValue('時間已過');
      continue;
    }

    // 將未來時間加入 Set（自動去重）
    futureTimes.add(date.toISOString());
  }

  console.log(`工作表「${sheetName}」發現 ${futureTimes.size} 個不同的未來時間點`);

  // 2️⃣ 刪除此工作表的現有觸發器
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
            console.log(`刪除工作表 ${sheetName} 的舊觸發器（UID=${trig.getUniqueId()}）`);
            deletedCount++;
          }
        } catch (e) {
          // JSON 解析失敗，也刪除這個觸發器
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          console.log(`刪除損壞的觸發器（UID=${trig.getUniqueId()}）`);
          deletedCount++;
        }
      }
    }
  }

  if (deletedCount > 0) {
    console.log(`已刪除 ${deletedCount} 個舊觸發器`);
  }

  // 3️⃣ 為每個不同的未來時間點建立觸發器
  if (futureTimes.size > 0) {
    let createdCount = 0;
    const triggerInfos = [];

    for (const timeStr of futureTimes) {
      const triggerTime = new Date(timeStr);

      // 統計這個時間點有多少個帳號
      let accountCount = 0;
      for (let row = 1; row < data.length; row++) {
        const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
        const rowTimeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
        if (!email || !rowTimeStr) continue;

        const rowDate = new Date(rowTimeStr);
        if (isNaN(rowDate.getTime())) continue;

        // 使用時間差比對，允許 1 分鐘誤差
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

      console.log(`✅ 為工作表 ${sheetName} 建立觸發器：${timeStr} (${accountCount} 個帳號) (UID=${trigger.getUniqueId()})`);

      triggerInfos.push({
        time: triggerTime.toLocaleString('zh-TW'),
        count: accountCount
      });

      createdCount++;
    }

    // 4️⃣ 標記已預約的帳號
    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const timeStr = data[row][timeColumnIndex];
      if (!email || !timeStr) continue;

      const date = new Date(timeStr);
      if (isNaN(date.getTime())) continue;

      const key = date.toISOString();
      if (futureTimes.has(key)) {
        sheet.getRange(row + 1, statusColumnIndex + 1).setValue('已預約');
      }
    }

    // 5️⃣ 顯示建立結果
    let message = `已為工作表「${sheetName}」建立 ${createdCount} 個觸發器：\n\n`;
    for (const info of triggerInfos) {
      message += `• ${info.time} - ${info.count} 個帳號\n`;
    }
    SpreadsheetApp.getUi().alert(message);

  } else {
    SpreadsheetApp.getUi().alert(`工作表「${sheetName}」目前沒有任何「未來時間」，不需要建立觸發器。`);
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

    console.log(`處理工作表：${sheet.getName()}`);
    console.log('處理的資料筆數:', data.length);
    console.log('當前時間:', now.toISOString());

    let processedCount = 0;

    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
      const timeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
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

        // 🔧 修正：改為使用 1 分鐘誤差，與建立觸發器時一致
        if (timeDiff < 60 * 1000) {
          shouldSuspend = true;
          console.log(`  ✅ 時間匹配 (目標時間比對)`);
        } else {
          console.log(`  ❌ 時間不匹配`);
        }
      } else {
        // 沒有指定目標時間，檢查是否已到預定時間
        // 🔧 修正：同樣改為 1 分鐘誤差
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
 * 建立連續通知信的觸發器
 */
function scheduleNotificationEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const notificationTimes = new Set(); // 收集所有通知時間點

  // 1️⃣ 掃描所有列，計算通知時間點
  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
    const timeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
    if (!email || !timeStr) continue;

    const suspendDate = new Date(timeStr);
    if (isNaN(suspendDate.getTime()) || suspendDate <= now) continue;

    // 計算四個通知時間點（停權前 4、3、2、1 週）
    for (let weeks = 4; weeks >= 1; weeks--) {
      const notificationDate = new Date(suspendDate.getTime() - (weeks * 7 * 24 * 60 * 60 * 1000));
      if (notificationDate > now) {
        notificationTimes.add(`${notificationDate.toISOString()}_${weeks}week`);
      }
    }

    // 🆕 新增：停權前 6 小時的通知
    const sixHoursBeforeDate = new Date(suspendDate.getTime() - (6 * 60 * 60 * 1000));
    if (sixHoursBeforeDate > now) {
      notificationTimes.add(`${sixHoursBeforeDate.toISOString()}_6hour`);
    }
  }

  console.log(`工作表「${sheetName}」發現 ${notificationTimes.size} 個通知時間點`);

  // 2️⃣ 刪除此工作表的現有通知觸發器
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
            console.log(`刪除工作表 ${sheetName} 的舊通知觸發器（UID=${trig.getUniqueId()}）`);
            deletedCount++;
          }
        } catch (e) {
          ScriptApp.deleteTrigger(trig);
          PropertiesService.getScriptProperties().deleteProperty(propKey);
          console.log(`刪除損壞的通知觸發器（UID=${trig.getUniqueId()}）`);
          deletedCount++;
        }
      }
    }
  }

  if (deletedCount > 0) {
    console.log(`已刪除 ${deletedCount} 個舊通知觸發器`);
  }

  // 3️⃣ 為每個通知時間點建立觸發器
  if (notificationTimes.size > 0) {
    let createdCount = 0;
    const triggerInfos = [];

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

      // 統計這個時間點需要通知的帳號數量
      let accountCount = 0;
      for (let row = 1; row < data.length; row++) {
        const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
        const rowTimeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
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

      console.log(`✅ 為工作表 ${sheetName} 建立通知觸發器：${displayText} (${triggerTime.toLocaleString('zh-TW')}) - ${accountCount} 個帳號`);

      triggerInfos.push({
        time: triggerTime.toLocaleString('zh-TW'),
        type: displayText,
        count: accountCount
      });

      createdCount++;
    }

    // 4️⃣ 標記已預約連續通知信的帳號
    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex];
      const timeStr = data[row][timeColumnIndex];
      if (!email || !timeStr) continue;

      const suspendDate = new Date(timeStr);
      if (isNaN(suspendDate.getTime()) || suspendDate <= now) continue;

      // 檢查這個帳號是否有任何通知時間點
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
        sheet.getRange(row + 1, MailStatusColumnIndex + 1).setValue('已預約連續通知信');
      }
    }

    // 5️⃣ 顯示建立結果
    let message = `已為工作表「${sheetName}」建立 ${createdCount} 個通知觸發器：\n\n`;
    for (const info of triggerInfos) {
      message += `• ${info.type} (${info.time}) - ${info.count} 個帳號\n`;
    }
    SpreadsheetApp.getUi().alert(message);

  } else {
    SpreadsheetApp.getUi().alert(`工作表「${sheetName}」目前沒有需要設定通知的帳號。`);
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

    console.log(`處理工作表：${sheet.getName()}`);
    console.log('當前時間:', now.toISOString());

    let sentCount = 0;

    for (let row = 1; row < data.length; row++) {
      const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
      const timeStr = data[row][timeColumnIndex]; // 假設時間在第 6 欄（F 欄）
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
          const currentStatus = sheet.getRange(row + 1, MailStatusColumnIndex + 1).getValue() || '';
          const newStatus = currentStatus ?
            `${currentStatus}; 已發送${timeInfo}前通知` :
            `已發送${timeInfo}前通知`;
          sheet.getRange(row + 1, MailStatusColumnIndex + 1).setValue(newStatus);
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
 * Suspend All User
 */
function SuspendAllUser() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  console.log(`開始處理 ${data.length - 1} 筆使用者資料`);

  // 從第2列開始讀取（跳過標題列）
  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex]; // 假設 email 在第 3 列（C 欄）

    // 跳過空白的 email
    if (!email || email.trim() === '') {
      console.log(`第 ${row + 1} 列：email 為空，跳過`);
      continue;
    }

    console.log(`第 ${row + 1} 列：準備停權 ${email}`);

    try {
      AdminDirectory.Users.update({ suspended: true }, email);
      sheet.getRange(row + 1, statusColumnIndex + 1).setValue('已停權'); // 在 G 欄標記狀態
      console.log(`✅ 停權成功：${email}`);
    } catch (err) {
      console.log(`❌ 停權失敗 (${email}): ${err.message}`);
      sheet.getRange(row + 1, errorColumnIndex + 1).setValue(`錯誤: ${err.message}`); // 在 H 欄標記錯誤
    }
  }

  console.log('批次停權作業完成');
  SpreadsheetApp.getUi().alert('批次停權作業完成，請查看 G 欄的執行結果。');
}

/**
 * 清理目前分頁的所有觸發器（手動執行用）
 */
function cleanAllTriggers() {
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

  // 🆕 清空 G 欄（狀態）和 I 欄（郵件狀態）- 只清理觸發器相關的狀態
  const data = sheet.getDataRange().getValues();
  let clearedCells = 0;

  for (let row = 1; row < data.length; row++) {
    const email = data[row][emailColumnIndex]; // 假設 email 在第 3 欄（C 欄）
    if (!email) continue; // 跳過沒有 email 的列

    // 清空 G 欄（狀態欄）- 只清理觸發器設定的狀態
    const statusCell = sheet.getRange(row + 1, statusColumnIndex + 1);
    const currentStatus = statusCell.getValue();
    if (currentStatus === '已預約') {
      statusCell.setValue('');
      clearedCells++;
    }

    // 清空 I 欄（郵件狀態欄）- 只清理觸發器設定的狀態
    const mailStatusCell = sheet.getRange(row + 1, MailStatusColumnIndex + 1);
    const currentMailStatus = mailStatusCell.getValue();
    if (currentMailStatus && (
      currentMailStatus.includes('已預約連續通知信') ||
      currentMailStatus.includes('已發送') ||
      currentMailStatus.includes('前通知')
    )) {
      mailStatusCell.setValue('');
      clearedCells++;
    }
  }

  const totalDeleted = deletedSuspendTriggers + deletedNotificationTriggers;

  if (totalDeleted > 0 || clearedCells > 0) {
    console.log(`工作表「${sheetName}」清理完成：`);
    console.log(`- 停權觸發器：${deletedSuspendTriggers} 個`);
    console.log(`- 通知觸發器：${deletedNotificationTriggers} 個`);
    console.log(`- 清空相關狀態：${clearedCells} 個儲存格`);

    SpreadsheetApp.getUi().alert(`工作表「${sheetName}」清理完成：\n\n• 停權觸發器：${deletedSuspendTriggers} 個\n• 通知觸發器：${deletedNotificationTriggers} 個\n• 清空相關狀態：${clearedCells} 個儲存格`);
  } else {
    console.log(`工作表「${sheetName}」目前沒有任何觸發器或相關狀態需要清理`);
    SpreadsheetApp.getUi().alert(`工作表「${sheetName}」目前沒有任何觸發器或相關狀態需要清理。`);
  }
}