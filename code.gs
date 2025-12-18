/**
 * タスクリマインダー GAS（メール + カレンダー）
 *
 * 機能:
 * - 期限の 7日前 / 3日前 / 1日前 にメールでリマインド
 * - カレンダーにタスクのイベントを自動作成
 *
 * 使い方:
 * 1. CONFIG.SPREADSHEET_ID をスプレッドシートのIDに変更
 * 2. createDailyTrigger を実行 → 毎日9時に自動実行
 */

// =============================================
// 設定
// =============================================
const CONFIG = {
  // スプレッドシートのID（URLの /d/XXXXX/edit の XXXXX 部分）
  SPREADSHEET_ID: "YOUR_SPREADSHEET_ID_HERE",

  // シート名
  SHEET_NAME: "Sheet1",

  // 列の位置（A列=1, B列=2, ...）
  COLUMNS: {
    TASK: 1, // A列: タスク名
    STATUS: 2, // B列: ステータス
    OWNER: 3, // C列: 担当者名
    DEADLINE: 4, // D列: 期限（日付または日時）
    EMAIL: 5, // E列: メールアドレス
  },

  // 「完了」とみなすステータス
  COMPLETED_STATUS: "完了",

  // リマインドを送る日（期限の何日前か）
  REMINDER_DAYS: [7, 3, 1],

  // メール設定
  EMAIL: {
    SENDER_NAME: "タスクリマインダー",
    SUBJECT_PREFIX: "【リマインド】",
  },

  // カレンダー設定
  CALENDAR: {
    ENABLED: true, // カレンダー連携を有効にするか
    REMINDERS: [60, 1440], // 通知（1時間前、1日前）
    DEFAULT_HOUR: 17, // イベントの開始時間（0-23）
    DEFAULT_MINUTE: 0, // イベントの開始分（0-59）
  },
};

// =============================================
// メイン処理（トリガーで毎日実行）
// =============================================

/**
 * 毎日実行されるメイン関数
 * - リマインド対象のタスクにメールを送信
 * - 未登録のタスクをカレンダーに追加
 */
function checkAndSendReminders() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(
    CONFIG.SHEET_NAME
  );
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const calendar = CalendarApp.getDefaultCalendar();
  let emailCount = 0;
  let calendarCount = 0;

  // 2行目以降（ヘッダーをスキップ）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const task = row[CONFIG.COLUMNS.TASK - 1];
    const status = row[CONFIG.COLUMNS.STATUS - 1];
    const owner = row[CONFIG.COLUMNS.OWNER - 1];
    const deadlineRaw = new Date(row[CONFIG.COLUMNS.DEADLINE - 1]);
    const email = row[CONFIG.COLUMNS.EMAIL - 1] || null;

    // 完了済みはスキップ
    if (status === CONFIG.COMPLETED_STATUS) continue;

    // --- メール送信 ---
    const deadlineForCalc = new Date(deadlineRaw);
    deadlineForCalc.setHours(0, 0, 0, 0);
    const daysUntil = Math.ceil(
      (deadlineForCalc - today) / (1000 * 60 * 60 * 24)
    );

    if (CONFIG.REMINDER_DAYS.includes(daysUntil)) {
      sendReminderEmail_({
        task,
        owner,
        email,
        deadline: deadlineForCalc,
        daysUntil,
        status,
      });
      emailCount++;
    }

    // --- カレンダー作成 ---
    if (CONFIG.CALENDAR.ENABLED) {
      const eventTitle = `📋 ${task} - ${owner}`;
      const existingEvents = calendar.getEventsForDay(deadlineRaw, {
        search: task,
      });

      if (existingEvents.length === 0) {
        createCalendarEvent_(calendar, eventTitle, deadlineRaw, {
          description: `タスク: ${task}\n担当者: ${owner}\nステータス: ${status}`,
          email: email,
        });
        calendarCount++;
      }
    }
  }

  console.log(`完了: メール ${emailCount}件, カレンダー ${calendarCount}件`);
}

// =============================================
// 公開関数（メニューに表示）
// =============================================

/**
 * 全タスクのカレンダーイベントを一括作成
 */
function createCalendarEventsForAllTasks() {
  console.log("カレンダーイベントを一括作成中...");
  checkAndSendReminders();
}

/**
 * 毎日9時に自動実行するトリガーを設定
 */
function createDailyTrigger() {
  // 既存のトリガーを削除
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));

  // 新しいトリガーを作成
  ScriptApp.newTrigger("checkAndSendReminders")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  console.log("✅ トリガー設定完了: 毎日9時に実行");
}

/**
 * テスト: スプレッドシートの最初のタスクでカレンダーイベントを作成
 */
function testCreateCalendarEvent() {
  const myEmail = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(
    CONFIG.SHEET_NAME
  );
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    console.error("データがありません");
    return;
  }

  const row = data[1];
  const task = row[CONFIG.COLUMNS.TASK - 1];
  const owner = row[CONFIG.COLUMNS.OWNER - 1];
  const status = row[CONFIG.COLUMNS.STATUS - 1];
  const deadline = new Date(row[CONFIG.COLUMNS.DEADLINE - 1]);
  const eventTitle = `📋 ${task} - ${owner}`;

  console.log(`カレンダー作成テスト:`);
  console.log(`  タスク: ${task}`);
  console.log(
    `  期限: ${formatDate_(deadline)} ${CONFIG.CALENDAR.DEFAULT_HOUR}:${String(
      CONFIG.CALENDAR.DEFAULT_MINUTE
    ).padStart(2, "0")}`
  );

  // 重複チェック
  const calendar = CalendarApp.getDefaultCalendar();
  const existingEvents = calendar.getEventsForDay(deadline, { search: task });

  if (existingEvents.length > 0) {
    console.log(`⚠️ 既にカレンダーに登録済みです（スキップ）`);
    return;
  }

  createCalendarEvent_(calendar, eventTitle, deadline, {
    description: `タスク: ${task}\n担当者: ${owner}\nステータス: ${status}\n\n※ テスト作成`,
    email: myEmail,
  });

  console.log("✅ カレンダーを確認してください");
}

/**
 * テスト: 自分にテストメールを送信
 */
function sendTestEmail() {
  const myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) {
    console.error("メールアドレス取得不可");
    return;
  }

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(
    CONFIG.SHEET_NAME
  );
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    console.error("データがありません");
    return;
  }

  const row = data[1];
  sendReminderEmail_({
    task: row[CONFIG.COLUMNS.TASK - 1],
    owner: row[CONFIG.COLUMNS.OWNER - 1],
    status: row[CONFIG.COLUMNS.STATUS - 1],
    deadline: new Date(row[CONFIG.COLUMNS.DEADLINE - 1]),
    email: myEmail,
    daysUntil: 3,
  });

  console.log("✅ テストメール送信完了");
}

// =============================================
// 内部関数（メニューに非表示）
// =============================================

/**
 * リマインドメールを送信
 */
function sendReminderEmail_(r) {
  if (!r.email) return;

  const daysText = r.daysUntil === 1 ? "明日" : `${r.daysUntil}日後`;
  const subject = `${CONFIG.EMAIL.SUBJECT_PREFIX}「${r.task}」の期限が${daysText}です`;

  const body = [
    `${r.owner} さん`,
    ``,
    `以下のタスクの期限が ${daysText}（${formatDate_(
      r.deadline
    )}）に迫っています。`,
    ``,
    `━━━━━━━━━━━━━━━━━━━━`,
    `タスク: ${r.task}`,
    `期限: ${formatDate_(r.deadline)}`,
    `ステータス: ${r.status}`,
    `━━━━━━━━━━━━━━━━━━━━`,
    ``,
    `期限までにタスクを完了してください。`,
    ``,
    `---`,
    `このメールは自動送信されています。`,
  ].join("\n");

  try {
    MailApp.sendEmail({
      to: r.email,
      subject: subject,
      body: body,
      name: CONFIG.EMAIL.SENDER_NAME,
    });
    console.log(`[メール] ${r.owner}: ${r.task}`);
  } catch (e) {
    console.error(`[メールエラー] ${e}`);
  }
}

/**
 * カレンダーイベントを作成
 */
function createCalendarEvent_(calendar, title, startTime, opts) {
  const options = { description: opts.description };

  // デフォルト時間を適用
  const eventStart = new Date(startTime);
  eventStart.setHours(
    CONFIG.CALENDAR.DEFAULT_HOUR,
    CONFIG.CALENDAR.DEFAULT_MINUTE,
    0,
    0
  );

  const eventEnd = new Date(eventStart);
  eventEnd.setHours(eventStart.getHours() + 1);

  const event = calendar.createEvent(title, eventStart, eventEnd, options);

  if (opts.email) event.addGuest(opts.email);

  event.removeAllReminders();
  CONFIG.CALENDAR.REMINDERS.forEach((min) => event.addPopupReminder(min));

  const timeStr = `${CONFIG.CALENDAR.DEFAULT_HOUR}:${String(
    CONFIG.CALENDAR.DEFAULT_MINUTE
  ).padStart(2, "0")}`;
  console.log(`[カレンダー] ${title} (${timeStr})`);
}

/**
 * 日付を「YYYY年M月D日」形式でフォーマット
 */
function formatDate_(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}
