/**
 * プロパティサービスから値を取得する
 * @param {string} key - キー名
 * @returns {string} 取得した値
 */
const getProperty = (key) => {
  return PropertiesService.getScriptProperties().getProperty(key);
};

const BOT_USER_TOKEN = getProperty("BOT_USER_TOKEN");
const CHANNEL_ID = getProperty("CHANNEL_ID");
const USER_ID = getProperty("USER_ID");
const SHEET_ID = getProperty("SHEET_ID");
const SHEET = SpreadsheetApp.openById(SHEET_ID);

/**
 * シート名からシートオブジェクトを取得する
 * @param {string} sheetName - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} シートオブジェクト
 */
const getSheetByName = (sheetName) => {
  return SHEET.getSheetByName(sheetName) ?? null;
};

/**
 * カレンダーIDからカレンダーオブジェクトを取得する
 * @param {string} calendarId - カレンダーID
 * @returns {GoogleAppsScript.Calendar.Calendar | null} カレンダーオブジェクト
 */
const getCalendarById = (calendarId) => {
  return CalendarApp.getCalendarById(calendarId) ?? null;
};

/**
 * カレンダーから1週間後の終日イベント一覧を取得する
 * @param {GoogleAppsScript.Calendar.Calendar} calendar - カレンダーオブジェクト
 * @returns {GoogleAppsScript.Calendar.CalendarEvent[]} 終日イベントの配列
 */
const getAllDayEventsNextWeek = (calendar) => {
  const now = new Date();
  const nextWeek = new Date(now);
  nextWeek.setDate(now.getDate() + 7);

  return calendar
    .getEventsForDay(nextWeek)
    .filter((event) => event.isAllDayEvent());
};

/**
 * カレンダーから過去31日以内（当日除く）の終日イベントを取得する
 * @param {GoogleAppsScript.Calendar.Calendar} calendar - カレンダーオブジェクト
 * @returns {GoogleAppsScript.Calendar.CalendarEvent[]} 終日イベントの配列
 */
const getAllDayEventsWithinMonth = (calendar) => {
  const now = new Date();
  const start = new Date(now);
  start.setDate(now.getDate() - 31);

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  return calendar
    .getEvents(start, today)
    .filter((event) => event.isAllDayEvent());
};

/**
 * シート名からシートオブジェクトを取得する
 * @param {string} sheetName - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} シートオブジェクト
 */
const getSheetByName = (sheetName) => {
  return SHEET.getSheetByName(sheetName) ?? null;
};

/**
 * MEMBERSシートから人名とSlack IDのMapを生成する
 * @returns {Map<string, string>} 人名をキー、Slack IDを値とするMap
 */
const getMembersMap = () => {
  const sheet = getSheetByName("MEMBERS");
  const rows = sheet.getDataRange().getValues();

  return rows.slice(1).reduce((map, [name, slackId]) => {
    map.set(name, slackId);
    return map;
  }, new Map());
};

/**
 * SlackのIncoming WebhookでメッセージをPOSTする
 * @param {string} hookUrl - Webhook URL
 * @param {string} msg - 投稿するメッセージ
 */
const postToSlack = (hookUrl, msg) => {
  const payload = {
    text: msg,
  };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };
  UrlFetchApp.fetch(hookUrl, options);
};

const checkNextTask = () => {
  getAllDayEventsNextWeek().forEach((event) => {
    event;
  });
};
