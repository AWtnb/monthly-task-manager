"use strict";
/**
 * カレンダーから過去31日以内（当日含む）の終日イベントを取得する
 * @returns 終日イベントの配列
 */
const getAllDayEventsWithinMonth = () => {
    const now = new Date();
    const start = new Date(now);
    start.setDate(now.getDate() - 31);
    const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    return CALENDAR.getEvents(start, end).filter((event) => event.isAllDayEvent());
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
        const sorted = new Map([...byDeadline.entries()].sort(([a], [b]) => {
            if (a < b)
                return -1;
            if (b < a)
                return 1;
            return 0;
        }));
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
 * @param sheetUrl - シートのURL
 * @param pendingTasks - 担当者 → 期限日（昇順） → タスク一覧
 * @returns Slack Block Kit の blocks 配列
 */
const buildPendingTasksBlocks = (title, sheetUrl, pendingTasks) => {
    const blocks = [
        {
            type: "header",
            text: {
                type: "plain_text",
                text: `${title} 未了タスク一覧`,
                emoji: true,
            },
            level: 1,
        },
        {
            type: "rich_text",
            elements: [
                {
                    type: "rich_text_section",
                    elements: [
                        {
                            type: "link",
                            url: sheetUrl,
                            text: "シート",
                            style: {
                                bold: true,
                            },
                        },
                        {
                            type: "text",
                            text: "を確認しましょう！\n",
                        },
                    ],
                },
            ],
        },
    ];
    for (const [person, byDeadline] of pendingTasks) {
        blocks.push({ type: "divider" });
        const slackId = SLACK_ID_MAPPING.get(person);
        const mention = slackId ? `<@${slackId}>` : person;
        const personHeader = {
            type: "section",
            text: {
                type: "mrkdwn",
                text: `:white_check_mark: *Task of ${mention}*`,
            },
        };
        blocks.push(personHeader);
        for (const [deadlineKey, tasks] of byDeadline) {
            const deadlineHeader = {
                type: "section",
                text: {
                    type: "mrkdwn",
                    text: `*${deadlineKey}* 〆`,
                },
            };
            blocks.push(deadlineHeader);
            blocks.push({
                type: "context",
                elements: tasks.map((task) => ({
                    type: "mrkdwn",
                    text: `• ${task}`,
                })),
            });
        }
    }
    return blocks;
};
/**
 * 【定期実行】
 * 現時点のタスクをSlackに通知する
 */
const checkCurrentTask = () => {
    getAllDayEventsWithinMonth().forEach((event) => {
        const title = event.getTitle();
        const sheet = getSheetByName(title);
        if (!sheet)
            return;
        const tasks = collectPendingTasks(sheet, 7);
        if (tasks.size === 0)
            return;
        const sheetId = sheet.getSheetId();
        const sheetUrl = `${SHEET.getUrl()}?gid=${sheetId}`;
        postToSlack(buildPendingTasksBlocks(title, sheetUrl, tasks));
    });
};
