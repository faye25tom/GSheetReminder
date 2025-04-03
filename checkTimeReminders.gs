/**
 * @OnlyCurrentDoc // Limits the script to only affect the current spreadsheet.
 */

// --- 全局配置区 ---
// 将配置移到全局，以便所有函数都能访问

// 工作表名称 (将 "Sheet1" 替换为你的工作表名称)
const SHEET_NAME = "Sheet1";
// 包含提醒时间的列字母 (例如: "A")
const TIME_COLUMN_LETTER = "A";
// 用于标记提醒状态的列字母 (例如: "B")
const STATUS_COLUMN_LETTER = "B";
// 包含提醒消息内容的列字母 (例如: "C")
const MESSAGE_COLUMN_LETTER = "C";
// 状态列中标记已发送提醒的文本
const SENT_STATUS_TEXT = "提醒已发送";

// --- 邮件提醒逻辑 (由时间触发器运行) ---

/**
 * 主函数，用于检查时间并通过邮件发送提醒。
 * 你需要为此函数设置一个时间驱动的触发器。
 */
function checkTimeReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const recipientEmail = Session.getActiveUser().getEmail();

  // 检查工作表是否存在
  if (!sheet) {
    const errorMsg = `错误：找不到名为 "${SHEET_NAME}" 的工作表。请检查脚本中的 SHEET_NAME 配置。邮件提醒脚本无法继续运行。`;
    Logger.log(errorMsg);
    if (recipientEmail) {
      try {
        MailApp.sendEmail(recipientEmail, `Google Sheet 提醒脚本错误`, errorMsg);
      } catch (e) {
        Logger.log(`发送配置错误邮件失败: ${e}`);
      }
    }
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("工作表数据为空或只有标题行，邮件提醒无需检查。");
    return;
  }

  const now = new Date();
  const timeCol = columnLetterToIndex(TIME_COLUMN_LETTER);
  const statusCol = columnLetterToIndex(STATUS_COLUMN_LETTER);
  const messageCol = columnLetterToIndex(MESSAGE_COLUMN_LETTER);

  // 检查列字母是否有效
   if (!timeCol || !statusCol || !messageCol) {
     const errorMsg = `错误：脚本配置中的列字母 (${TIME_COLUMN_LETTER}, ${STATUS_COLUMN_LETTER}, ${MESSAGE_COLUMN_LETTER}) 无效。邮件提醒脚本无法继续运行。`;
     Logger.log(errorMsg);
     if (recipientEmail) {
       try {
         MailApp.sendEmail(recipientEmail, `Google Sheet 提醒脚本配置错误`, errorMsg);
       } catch (e) {
         Logger.log(`发送配置错误邮件失败: ${e}`);
       }
     }
     return;
  }

  const timeRange = sheet.getRange(2, timeCol, lastRow - 1, 1);
  const statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1);
  const messageRange = sheet.getRange(2, messageCol, lastRow - 1, 1);
  const timeValues = timeRange.getValues();
  const statusValues = statusRange.getValues();
  const messageValues = messageRange.getValues();

  for (let i = 0; i < timeValues.length; i++) {
    const reminderTimeCell = timeValues[i][0];
    const statusCell = statusValues[i][0];
    const messageCell = messageValues[i][0];
    const currentRow = i + 2;

    if (statusCell === SENT_STATUS_TEXT) {
      continue;
    }

    if (reminderTimeCell instanceof Date && !isNaN(reminderTimeCell)) {
      const reminderTime = new Date(reminderTimeCell);
      const reminderTimeStripped = new Date(reminderTime.getFullYear(), reminderTime.getMonth(), reminderTime.getDate(), reminderTime.getHours(), reminderTime.getMinutes());
      const nowStripped = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes());

      if (reminderTimeStripped <= nowStripped) {
        Logger.log(`找到匹配时间：行 ${currentRow}, 时间: ${reminderTime}, 消息: ${messageCell}`);
        const messageBody = messageCell ? messageCell.toString() : `时间到了！(来自工作表 "${SHEET_NAME}" 的单元格 ${TIME_COLUMN_LETTER}${currentRow})`;
        const subject = `Google Sheet 提醒: ${messageCell ? messageCell.toString().substring(0, 50) : '时间到了！'}`;

        if (!recipientEmail) {
           Logger.log(`错误：无法获取当前用户的邮箱地址，无法发送邮件提醒 (行 ${currentRow})。`);
           continue;
        }

        try {
          MailApp.sendEmail(recipientEmail, subject, `来自工作表 "${SHEET_NAME}" (行 ${currentRow}) 的提醒：\n\n${messageBody}`);
          Logger.log(`邮件提醒已发送至 ${recipientEmail} (行 ${currentRow})`);
          sheet.getRange(currentRow, statusCol).setValue(SENT_STATUS_TEXT);
          Logger.log(`行 ${currentRow} 的状态已更新为 "${SENT_STATUS_TEXT}"`);
        } catch (e) {
          Logger.log(`发送邮件提醒失败 (行 ${currentRow}): ${e}`);
        }
        // Utilities.sleep(1000); // 可选延迟
      }
    } else if (reminderTimeCell !== "") {
      Logger.log(`警告：行 ${currentRow} 的时间单元格 (${TIME_COLUMN_LETTER}${currentRow}) 内容 "${reminderTimeCell}" 不是有效的日期/时间格式。`);
    }
  }
  Logger.log("邮件提醒时间检查完成。");
}

// --- 状态自动重置逻辑 (由 onEdit 触发器自动运行) ---

/**
 * 当用户编辑电子表格时自动运行。
 * 如果时间列 ('A') 被修改为一个晚于当前时间的值，则清除状态列 ('B')。
 * @param {Object} e 事件对象，包含有关编辑的信息。
 */
function onEdit(e) {
  // 检查事件对象是否存在
  if (!e || !e.range) {
    // Logger.log("onEdit 事件对象无效或缺少范围信息。");
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const newValue = e.value; // 编辑后的新值

  // 检查是否是正确的表单，并且编辑发生在数据行 (非标题行)
  if (sheet.getName() === SHEET_NAME && editedRow > 1) {
    const timeColIndex = columnLetterToIndex(TIME_COLUMN_LETTER);
    const statusColIndex = columnLetterToIndex(STATUS_COLUMN_LETTER);

    // 检查是否编辑了时间列，并且列索引有效
    if (editedCol === timeColIndex && timeColIndex && statusColIndex) {
      Logger.log(`监测到编辑：工作表 "${SHEET_NAME}", 单元格 ${TIME_COLUMN_LETTER}${editedRow}, 新值: ${newValue}`);

      // 获取当前时间
      const now = new Date();

      // 检查新值是否是有效的日期对象并且晚于当前时间
      if (newValue instanceof Date && !isNaN(newValue) && newValue > now) {
         Logger.log(`新时间 (${newValue}) 晚于当前时间 (${now})。准备清除行 ${editedRow} 的状态。`);
         // 清除同一行的状态单元格内容
         sheet.getRange(editedRow, statusColIndex).clearContent();
         Logger.log(`行 ${editedRow} 的状态 (${STATUS_COLUMN_LETTER}${editedRow}) 已被清除。`);
      } else {
         // 如果新值不是日期，或是过去的日期，或者被清空了，则不执行清除操作
         Logger.log(`新值不是未来的有效时间，行 ${editedRow} 的状态未被清除。`);
      }
    }
  }
}


// --- 辅助函数 ---

/**
 * 辅助函数：将列字母转换为列索引 (A=1, B=2, ...)
 * @param {string} letter 列字母 (例如 "A", "B", "AA"). 大小写不敏感。
 * @return {number|null} 列索引，如果字母无效则返回 null.
 */
function columnLetterToIndex(letter) {
  if (typeof letter !== 'string' || letter.length === 0) {
    return null;
  }
  letter = letter.toUpperCase();
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    const charCode = letter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) { // ASCII A=65, Z=90
       return null; // 非法字符
    }
    column += (charCode - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

// --- 可选菜单 (用于手动测试) ---

/**
 * 在电子表格打开时添加一个自定义菜单。
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('提醒工具')
      .addItem('手动检查提醒 (发送邮件)', 'checkTimeReminders')
      .addToUi();
}
