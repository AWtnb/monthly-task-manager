"use strict";
/**
 * プロパティサービスから値を取得する
 * @param key - キー名
 * @returns 取得した値
 */
const getProperty = (key) => {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    if (!value)
        throw new Error(`Property not found: ${key}`);
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
const getSheetByName = (sheetName) => {
    var _a;
    return (_a = SHEET.getSheetByName(sheetName)) !== null && _a !== void 0 ? _a : null;
};
/**
 * SlackのIncoming WebhookでメッセージをPOSTする
 * @param blocks: object[] - 投稿するメッセージのSlack Bot Kit形式データ
 */
const postToSlack = (blocks) => {
    const payload = JSON.stringify({ blocks });
    UrlFetchApp.fetch(WEBHOOK_URL, {
        method: "post",
        contentType: "application/json",
        payload,
    });
};
