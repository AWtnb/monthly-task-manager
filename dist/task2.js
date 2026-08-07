"use strict";
/**
 * カレンダーから過去31日以内（当日除く）の終日イベントを取得する
 * @param calendar - カレンダーオブジェクト
 * @returns 終日イベントの配列
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
 * シートから担当者ごと・期限日ごとに未了工程を集約する
 * @param sheet - 対象シート
 * @param dataStartRow - データ開始行（1始まり）
 * @returns 担当者 → 期限日（昇順） → タスク一覧
 */
const collectPendingTasks = (sheet, dataStartRow) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < dataStartRow)
        return new Map();
    const data = sheet
        .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7)
        .getValues();
    const result = new Map();
    for (const row of data) {
        const [task, done, person, , , , deadline] = row;
        if (!person || done !== false)
            continue;
        const deadlineKey = Utilities.formatDate(new Date(deadline), Session.getScriptTimeZone(), "yyyy/MM/dd");
        if (!result.has(person)) {
            result.set(person, new Map());
        }
        const byDeadline = result.get(person);
        if (!byDeadline.has(deadlineKey)) {
            byDeadline.set(deadlineKey, []);
        }
        byDeadline.get(deadlineKey).push(task);
    }
    // 各担当者のdeadlineをキーの文字列昇順（= 日付昇順）でソート
    for (const [person, byDeadline] of result) {
        const sorted = new Map([...byDeadline.entries()].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0));
        result.set(person, sorted);
    }
    return result;
};
const SLACK_ID_MAPPING = (() => {
    const map = new Map();
    const sheet = getSheetByName("MEMBER");
    if (!sheet)
        return map;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2)
        return map;
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const [name, slackId] of data) {
        if (!name)
            continue;
        map.set(name, slackId);
    }
    return map;
})();
/**
 * collectPendingTasksの結果をSlack Block Kit形式に変換する
 * @param pendingTasks - 担当者 → 期限日（昇順） → タスク一覧
 * @returns Slack Block Kit の blocks 配列
 */
const buildPendingTasksBlocks = (pendingTasks) => {
    const blocks = [
        {
            type: "header",
            text: { type: "plain_text", text: "📋 未了タスク一覧", emoji: true },
        },
        { type: "divider" },
    ];
    for (const [person, byDeadline] of pendingTasks) {
        // 担当者ヘッダー
        blocks.push({
            type: "section",
            text: { type: "mrkdwn", text: `*👤 ${person}*` },
        });
        // 期限日ごとのタスクリスト
        for (const [deadlineKey, tasks] of byDeadline) {
            const taskLines = tasks.map((t) => `• ${t}`).join("\n");
            blocks.push({
                type: "section",
                text: {
                    type: "mrkdwn",
                    text: `*${deadlineKey}*\n${taskLines}`,
                },
            });
        }
        blocks.push({ type: "divider" });
    }
    return blocks;
};
/**
 * 【定期実行】
 * 現時点のタスクをSlackに通知する
 */
const checkCurrentTask = () => {
    const calendar = getCalendarById(CALENDAR_ID);
    if (!calendar) {
        return;
    }
    getAllDayEventsWithinMonth(calendar).forEach((event) => {
        const title = event.getTitle();
        const sheet = getSheetByName(title);
        if (!sheet)
            return;
        const tasks = collectPendingTasks(sheet, 7);
        if (tasks.size === 0)
            return;
        postToSlack(WEBHOOK_URL, buildPendingTasksBlocks(tasks));
    });
};
