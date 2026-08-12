/**
 * カレンダーから過去31日以内（当日含む）の終日イベントを取得する
 * @returns 終日イベントの配列
 */
const getAllDayEventsWithinMonth =
  (): GoogleAppsScript.Calendar.CalendarEvent[] => {
    const now = new Date();
    const start = new Date(now);
    start.setDate(now.getDate() - 31);

    const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

    return CALENDAR.getEvents(start, end).filter((event) =>
      event.isAllDayEvent(),
    );
  };

type PendingTasksByPerson = Map<string, Map<string, string[]>>;

const DEADLINE_UNDEFINED = "(日付未指定)";

/**
 * 締切文字列を "yyyy/MM/dd" 形式のキーに変換する
 * 不正な日付文字列の場合は `(日付未指定)` を返す
 * @param deadline 締切日を表す文字列
 * @returns フォーマット済みの日付キー、または `(日付未指定)`
 */
const parseDeadlineKey = (deadline: string): string => {
  const date = new Date(deadline);
  if (Number.isNaN(date.getTime())) {
    return DEADLINE_UNDEFINED;
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
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
): PendingTasksByPerson => {
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
    string,
  ][];

  const result: PendingTasksByPerson = new Map();

  for (const row of data) {
    const [task, done, person, , , , deadline] = row;
    if (!person || done !== false) continue;

    const deadlineKey = parseDeadlineKey(deadline);

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
  // （ただし(日付未指定)は必ず最後に配置する）
  for (const [person, byDeadline] of result) {
    const sorted = new Map(
      [...byDeadline.entries()].sort(([a], [b]) => {
        if (a === DEADLINE_UNDEFINED) return 1;
        if (b === DEADLINE_UNDEFINED) return -1;
        if (a < b) return -1;
        if (b < a) return 1;
        return 0;
      }),
    );
    result.set(person, sorted);
  }

  return result;
};

const getSlackIdMapping = (): Map<string, string> => {
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
};

/**
 * collectPendingTasksの結果をSlack Block Kit形式に変換する
 * @param sheetUrl - シートのURL
 * @param pendingTasks - 担当者 → 期限日（昇順） → タスク一覧
 * @returns Slack Block Kit の blocks 配列
 */
const buildPendingTasksBlocks = (
  title: string,
  sheetUrl: string,
  pendingTasks: PendingTasksByPerson,
): object[] => {
  const blocks: object[] = [
    {
      type: "header",
      text: {
        type: "plain_text",
        text: `:spiral_calendar_pad: ${title} 未了タスク一覧`,
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

  const mapping = getSlackIdMapping();
  for (const [person, byDeadline] of pendingTasks) {
    blocks.push({ type: "divider" });
    const slackId = mapping.get(person);
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
      const deadlineHeader =
        deadlineKey === DEADLINE_UNDEFINED
          ? DEADLINE_UNDEFINED
          : `${deadlineKey} 〆`;
      const deadlineHeaderBlock = {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*${deadlineHeader}*`,
        },
      };
      blocks.push(deadlineHeaderBlock);
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
  const now = new Date();
  if (now.getDay() === 0 || now.getDay() === 6 || isHoliday(now)) return;

  getAllDayEventsWithinMonth().forEach((event) => {
    const title = event.getTitle();
    const sheet = getSheetByName(title);
    if (!sheet) return;
    const tasks = collectPendingTasks(sheet, 7);
    if (tasks.size === 0) return;

    const sheetId = sheet.getSheetId();
    const sheetUrl = `${SHEET.getUrl()}?gid=${sheetId}`;
    postToSlack(buildPendingTasksBlocks(title, sheetUrl, tasks));
  });
};
