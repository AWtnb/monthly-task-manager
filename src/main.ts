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
 * カレンダーから1週間後の終日イベント一覧を取得する
 * @param calendar - カレンダーオブジェクト
 * @returns 終日イベントの配列
 */
const getAllDayEventsNextWeek = (
  calendar: GoogleAppsScript.Calendar.Calendar,
): GoogleAppsScript.Calendar.CalendarEvent[] => {
  const now = new Date();
  const nextWeek = new Date(now);
  nextWeek.setDate(now.getDate() + 7);

  return calendar
    .getEventsForDay(nextWeek)
    .filter((event) => event.isAllDayEvent());
};

/**
 * カレンダーから過去31日以内（当日除く）の終日イベントを取得する
 * @param calendar - カレンダーオブジェクト
 * @returns 終日イベントの配列
 */
const getAllDayEventsWithinMonth = (
  calendar: GoogleAppsScript.Calendar.Calendar,
): GoogleAppsScript.Calendar.CalendarEvent[] => {
  const now = new Date();
  const start = new Date(now);
  start.setDate(now.getDate() - 31);

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  return calendar
    .getEvents(start, today)
    .filter((event) => event.isAllDayEvent());
};

/**
 * TEMPLATEシートをコピーして新しいシートを作成する
 * @param event - カレンダーイベント
 * @returns 作成したシートのURL
 */
const createSheetFromTemplate = (
  event: GoogleAppsScript.Calendar.CalendarEvent,
): string => {
  const title = event.getTitle();
  const start = event.getStartTime();
  const template = SHEET.getSheetByName("TEMPLATE");
  if (!template) throw new Error("TEMPLATE sheet not found");

  const newSheet = template.copyTo(SHEET);
  newSheet.setName(title);
  newSheet.getRange("A2").setValue(title);
  newSheet.getRange("A4").setValue(start);
  const sheetId = newSheet.getSheetId();
  return `${SHEET.getUrl()}?gid=${sheetId}`;
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

/**
 * 【定期実行】
 * 翌週の終日イベントをSlackに通知する
 */
const checkNextTask = () => {
  const calendar = getCalendarById(CALENDAR_ID);
  if (!calendar) {
    return;
  }
  const events = getAllDayEventsNextWeek(calendar);
  if (events.length < 1) {
    return;
  }

  const result = events
    .filter((event) => {
      const title = event.getTitle();
      return !getSheetByName(title);
    })
    .map((event) => {
      const title = event.getTitle();
      const shtUrl = createSheetFromTemplate(event);
      return { title: title, url: shtUrl };
    });

  if (result.length < 1) {
    return;
  }

  const start = events[0].getStartTime();
  const dateStr = Utilities.formatDate(
    start,
    Session.getScriptTimeZone(),
    "M月d日",
  );

  const blocks: object[] = [];
  blocks.push({
    type: "header",
    text: {
      type: "plain_text",
      text: `${dateStr}校了予定`,
      emoji: true,
    },
    level: 1,
  });

  const linkListElems = result.map((r) => {
    return {
      type: "rich_text_section",
      elements: [
        {
          type: "link",
          url: r.url,
          text: r.title,
          style: {
            bold: true,
          },
        },
      ],
    };
  });
  blocks.push({
    type: "rich_text",
    elements: [
      {
        type: "rich_text_section",
        elements: [
          {
            type: "broadcast",
            range: "channel",
          },
          {
            type: "text",
            text: "以下の校了予定があります。各シートを確認しておきましょう。\n",
          },
        ],
      },
      {
        type: "rich_text_list",
        style: "bullet",
        indent: 0,
        elements: linkListElems,
      },
    ],
  });

  postToSlack(WEBHOOK_URL, blocks);
};

/**
 * シートから担当者ごと・期限日ごとに未了工程を集約する
 * @param sheet - 対象シート
 * @param dataStartRow - データ開始行（1始まり）
 * @returns 担当者 → 期限日（昇順） → タスク一覧
 */
const collectPendingTasks = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dataStartRow: number,
): Map<string, Map<string, string[]>> => {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return new Map();

  const data = sheet
    .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7)
    .getValues() as [
    string,
    boolean,
    string,
    unknown,
    unknown,
    unknown,
    Date,
  ][];

  const result = new Map<string, Map<string, string[]>>();

  for (const row of data) {
    const [task, done, person, , , , deadline] = row;
    if (!person || done !== false) continue;

    const deadlineKey = Utilities.formatDate(
      new Date(deadline),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd",
    );

    if (!result.has(person)) {
      result.set(person, new Map<string, string[]>());
    }

    const byDeadline = result.get(person)!;
    if (!byDeadline.has(deadlineKey)) {
      byDeadline.set(deadlineKey, []);
    }
    byDeadline.get(deadlineKey)!.push(task);
  }

  // 各担当者のdeadlineをキーの文字列昇順（= 日付昇順）でソート
  for (const [person, byDeadline] of result) {
    const sorted = new Map(
      [...byDeadline.entries()].sort(([a], [b]) =>
        a < b ? -1 : a > b ? 1 : 0,
      ),
    );
    result.set(person, sorted);
  }

  return result;
};

const SLACK_ID_MAPPING = (() => {
  const map = new Map<string, string>();
  const sheet = getSheetByName("MEMBER");
  if (!sheet) return map;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues() as [
    string,
    string,
  ][];
  for (const [name, slackId] of data) {
    if (!name) continue;
    map.set(name, slackId);
  }

  return map;
})();

/**
 * collectPendingTasksの結果をSlack Block Kit形式に変換する
 * @param pendingTasks - 担当者 → 期限日（昇順） → タスク一覧
 * @returns Slack Block Kit の blocks 配列
 */
const buildPendingTasksBlocks = (
  pendingTasks: Map<string, Map<string, string[]>>,
): object[] => {
  const blocks: object[] = [
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
    if (!sheet) return;
    const tasks = collectPendingTasks(sheet, 7);
    if (tasks.size === 0) return;
    postToSlack(WEBHOOK_URL, buildPendingTasksBlocks(tasks));
  });
};
