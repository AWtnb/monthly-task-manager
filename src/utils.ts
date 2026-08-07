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
const CALENDAR = CalendarApp.getCalendarById(CALENDAR_ID);

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
 * `HOLIDAYS` シートから祝日の一覧を取得する
 * @returns 日付型の配列
 */
const getHolidays = (): Date[] => {
  const sheet = SHEET.getSheetByName("HOLIDAYS");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  return values
    .map((row) => row[0])
    .filter((cell) => 0 < String(cell).trim().length)
    .map((cell) => new Date(cell));
};

/**
 * 祝日かどうかを判定する
 * @param targetDate - 判定対象の日付
 * @returns boolean - 祝日ならtrue、そうでなければfalse
 */
const isHoliday = (targetDate: Date): boolean => {
  const targetYmd = Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyyMMdd");

  return getHolidays().some((holiday) => {
    const holidayYmd = Utilities.formatDate(holiday, "Asia/Tokyo", "yyyyMMdd");
    return holidayYmd === targetYmd;
  });
};

/**
 * SlackのIncoming WebhookでメッセージをPOSTする
 * @param blocks: object[] - 投稿するメッセージのSlack Bot Kit形式データ
 */
const postToSlack = (blocks: object[]): void => {
  const payload = JSON.stringify({ blocks });
  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload,
  });
};
