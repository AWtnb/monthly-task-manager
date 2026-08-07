/**
 * プロパティサービスから値を取得する
 * @param key - キー名
 * @returns 取得した値
 */
const getProperty = (key: string): string => {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error(`Property not found: ${key}`);
  return value;
};

const CALENDAR_ID = getProperty("CALENDAR_ID");
const SHEET_ID = getProperty("SHEET_ID");
const WEBHOOK_URL = getProperty("WEBHOOK_URL");
const SHEET = SpreadsheetApp.openById(SHEET_ID);

/**
 * シート名からシートオブジェクトを取得する
 * @param sheetName - シート名
 * @returns シートオブジェクト
 */
const getSheetByName = (
  sheetName: string,
): GoogleAppsScript.Spreadsheet.Sheet | null => {
  return SHEET.getSheetByName(sheetName) ?? null;
};

/**
 * カレンダーIDからカレンダーオブジェクトを取得する
 * @param calendarId - カレンダーID
 * @returns カレンダーオブジェクト
 */
const getCalendarById = (
  calendarId: string,
): GoogleAppsScript.Calendar.Calendar | null => {
  return CalendarApp.getCalendarById(calendarId) ?? null;
};

/**
 * SlackのIncoming WebhookでメッセージをPOSTする
 * @param webhookUrl - Webhook URL
 * @param blocks: object[] - 投稿するメッセージのSlack Bot Kit形式データ
 */
const postToSlack = (webhookUrl: string, blocks: object[]): void => {
  const payload = JSON.stringify({ blocks });

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload,
  });
};
