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
