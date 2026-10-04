/**
 * @OnlyCurrentDoc
 *
 * 日報・日次決算テンプレート 提出スクリプト（Google Apps Script）
 *
 * 入れ方:
 *   1. スプレッドシートの「ファイル」→「設定」で、タイムゾーンを「(GMT+09:00) Tokyo」にする
 *      （日付・提出日時の基準。Googleドライブで変換したファイルは米国時間になっていることがある）
 *   2. メニュー「拡張機能」→「Apps Script」を開く
 *   3. 最初からあるコードをすべて消し、このファイルの中身を貼り付けて保存する
 *   4. Apps Script の「プロジェクトの設定」（歯車のアイコン）で、タイムゾーンが東京になっているか確認する
 *      （リマインドの時刻の基準）
 *   5. スプレッドシートを再読み込みすると、メニューに「日報」が表示される
 *   6. 初回の実行時だけ承認画面が出るので、自分のアカウントを選んで許可する
 *
 * メニュー「日報」:
 *   本日の日報を提出        … 「月間日報」の「提出する日」（空欄なら今日）の列を「日次ログ」に1行保存する
 *                              （同じ日付があれば上書き確認）。設定シートに提出先メールアドレスがあればメールも送る
 *   月を締めて翌月へ進む    … 「月間日報」を「日報 2026年10月」のようなシートとして残し、入力欄を空にして翌月にする
 *   毎日のリマインドを設定    … 設定シートの「リマインドの時刻」に、その日の日報が未提出ならメールで知らせる
 *   リマインドを解除
 */

const SHEET_MAIN = '月間日報';
const SHEET_SUBMIT = '提出データ';
const SHEET_LOG = '日次ログ';
const SHEET_SETTINGS = '設定';
// 月間日報の場所（名前付き範囲）
const RANGE_SUBMIT_DATE = 'REPORT_DATE'; // 提出する日（空欄なら今日）
const RANGE_MONTH = 'REPORT_MONTH'; // 対象の月
const RANGE_DATES = 'DATE_ROW'; // 1日〜31日の日付の行
const RANGE_SUBMITTED = 'SUBMIT_ROW'; // 提出済み（○）の行
const RANGE_INPUTS = 'INPUT_BLOCK'; // 入力欄（薄い黄色のセル）の範囲
const INPUT_BACKGROUND = '#fff2cc'; // 入力セルの背景色
const LOG_FIRST_ROW = 3; // 日次ログのデータ開始行（1〜2行目は見出し、3行目は記入例）
const ARCHIVE_PREFIX = '日報 '; // 月を締めたときに残すシートの名前（例：日報 2026年10月）
const REMINDER_HANDLER = 'remindIfNotSubmitted';
const DEFAULT_REMINDER_HOUR = 21;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報')
    .addItem('本日の日報を提出', 'submitDailyReport')
    .addItem('月を締めて翌月へ進む', 'closeMonth')
    .addSeparator()
    .addItem('毎日のリマインドを設定', 'setupReminder')
    .addItem('リマインドを解除', 'removeReminder')
    .addToUi();
}

function submitDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const submit = ss.getSheetByName(SHEET_SUBMIT);
  const log = ss.getSheetByName(SHEET_LOG);
  if (!submit || !log || !ss.getRangeByName(RANGE_DATES)) {
    ui.alert('「' + SHEET_MAIN + '」「' + SHEET_SUBMIT + '」「' + SHEET_LOG + '」のいずれかが見つかりません。');
    return;
  }
  SpreadsheetApp.flush();

  const width = submit.getLastColumn();
  const headers = submit.getRange(1, 1, 1, width).getDisplayValues()[0];
  const row = submit.getRange(2, 1, 1, width);
  const values = row.getValues()[0];
  const display = row.getDisplayValues()[0];
  const formats = row.getNumberFormats()[0];

  const date = values[0];
  if (!(date instanceof Date) || date.getFullYear() < 2000) {
    ui.alert('「' + SHEET_MAIN + '」の「提出する日」を確認してください。');
    return;
  }
  const tz = ss.getSpreadsheetTimeZone();
  const dateText = Utilities.formatDate(date, tz, 'yyyy/MM/dd');

  const col = dayColumn_(ss, date, tz);
  if (col < 0) {
    ui.alert(dateText + ' は「' + SHEET_MAIN + '」の対象の月にありません。「提出する日」か「対象の月」を確認してください。');
    return;
  }
  if (!hasInput_(ss, col)) {
    ui.alert(dateText + ' の列がまだ空です。「' + SHEET_MAIN + '」に入力してから提出してください。');
    return;
  }

  const existing = findLogRow_(log, date, tz);
  if (existing) {
    const answer = ui.alert(
      '提出の確認',
      dateText + ' の日報はすでに日次ログの ' + existing + ' 行目にあります。上書きしますか？',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }

  const stampCol = headers.indexOf('提出日時');
  if (stampCol >= 0) values[stampCol] = new Date();

  const target = existing || Math.max(log.getLastRow() + 1, LOG_FIRST_ROW);
  const dest = log.getRange(target, 1, 1, width);
  dest.setValues([values]);
  dest.setNumberFormats([formats]);

  let mailNote = '';
  const to = String(getSetting_(ss, '提出先メールアドレス')).trim();
  if (to) {
    try {
      MailApp.sendEmail({
        to: to,
        subject: buildSubject_(ss, dateText),
        body: buildBody_(ss, headers, display),
      });
      mailNote = '\n' + to + ' にメールで送りました。';
    } catch (e) {
      mailNote = '\nメールは送れませんでした：' + e.message;
    }
  }

  // 「提出する日」を空欄に戻す（次は今日の日付で提出される）
  const dateCell = ss.getRangeByName(RANGE_SUBMIT_DATE);
  if (dateCell && dateCell.getValue() !== '') dateCell.clearContent();

  ui.alert('提出しました', dateText + ' の日報を日次ログの ' + target + ' 行目に保存しました。' + mailNote);
}

function closeMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(SHEET_MAIN);
  const monthCell = ss.getRangeByName(RANGE_MONTH);
  if (!sheet || !monthCell || !ss.getRangeByName(RANGE_INPUTS)) {
    ui.alert('「' + SHEET_MAIN + '」シートが見つかりません。');
    return;
  }
  const month = monthCell.getValue();
  if (!(month instanceof Date)) {
    ui.alert('「' + SHEET_MAIN + '」の「対象の月」を入力してください。');
    return;
  }
  const tz = ss.getSpreadsheetTimeZone();
  const label = Utilities.formatDate(month, tz, 'yyyy年M月');
  const name = uniqueSheetName_(ss, ARCHIVE_PREFIX + label);
  const pending = unsubmittedDays_(ss, tz);
  const answer = ui.alert(
    '月を締める',
    '「' + SHEET_MAIN + '」を「' + name + '」シートとして残し、入力欄を空にして翌月に進みます。' +
      (pending.length ? '\n\nまだ提出していない日があります：' + pending.join('、') +
        '\n（提出していない日は、日次ログ・月次集計・10年計画に入りません）' : '') +
      '\n\nよろしいですか？',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const copy = sheet.copyTo(ss).setName(name);
  ss.setActiveSheet(copy);
  ss.moveActiveSheet(ss.getNumSheets());
  clearInputs_(ss);
  const next = nextMonth_(month, tz);
  monthCell.setValue(next.text);
  ss.setActiveSheet(sheet);
  ui.alert('翌月に進みました', label + 'の日報を「' + name + '」シートに残しました。「' + SHEET_MAIN + '」は ' +
    next.label + ' になりました。');
}

function setupReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setting = getSetting_(ss, 'リマインドの時刻');
  const hour = setting !== '' && Number(setting) >= 0 && Number(setting) <= 23 ?
    Math.floor(Number(setting)) : DEFAULT_REMINDER_HOUR;
  removeReminderTriggers_();
  ScriptApp.newTrigger(REMINDER_HANDLER).timeBased().everyDays(1).atHour(hour).create();
  SpreadsheetApp.getUi().alert(
    '毎日 ' + hour + ' 時台（' + Session.getScriptTimeZone() + '）に、その日の日報が未提出ならメールで知らせます。\n' +
      '時刻がずれる場合は、Apps Scriptの「プロジェクトの設定」でタイムゾーンを日本時間（Asia/Tokyo）にしてください。');
}

function removeReminder() {
  removeReminderTriggers_();
  SpreadsheetApp.getUi().alert('リマインドを解除しました。');
}

function remindIfNotSubmitted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const today = new Date();
  if (!log || findLogRow_(log, today, tz)) return;
  const to = Session.getEffectiveUser().getEmail();
  if (!to) return;
  MailApp.sendEmail({
    to: to,
    subject: '【日報】' + Utilities.formatDate(today, tz, 'yyyy/MM/dd') + ' の日報がまだ提出されていません',
    body: '今日の日報がまだ提出されていません。\n\n' + ss.getUrl() +
      '\n\n「' + SHEET_MAIN + '」の今日の列に入力し、メニュー「日報」→「本日の日報を提出」で提出してください。',
  });
}

// ---------------------------------------------------------------- 内部の関数

/** 日次ログで、指定した日付の行番号を返す（なければ 0）。 */
function findLogRow_(log, date, tz) {
  const last = log.getLastRow();
  if (last < LOG_FIRST_ROW) return 0;
  const key = Utilities.formatDate(date, tz, 'yyyyMMdd');
  const dates = log.getRange(LOG_FIRST_ROW, 1, last - LOG_FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    if (d instanceof Date && Utilities.formatDate(d, tz, 'yyyyMMdd') === key) {
      return LOG_FIRST_ROW + i;
    }
  }
  return 0;
}

/** 月間日報で、指定した日付が何列目か（0始まり）を返す（対象の月になければ -1）。 */
function dayColumn_(ss, date, tz) {
  const dates = ss.getRangeByName(RANGE_DATES).getValues()[0];
  const key = Utilities.formatDate(date, tz, 'yyyyMMdd');
  for (let c = 0; c < dates.length; c++) {
    if (dates[c] instanceof Date && Utilities.formatDate(dates[c], tz, 'yyyyMMdd') === key) return c;
  }
  return -1;
}

/** 月間日報の入力欄で、指定した列に1つでも入力があるか。 */
function hasInput_(ss, col) {
  return ss.getRangeByName(RANGE_INPUTS).getValues().some(function (row) { return row[col] !== ''; });
}

/** 入力があるのに日次ログに提出していない日（例：10/5）の一覧。 */
function unsubmittedDays_(ss, tz) {
  const dates = ss.getRangeByName(RANGE_DATES).getValues()[0];
  const done = ss.getRangeByName(RANGE_SUBMITTED).getValues()[0];
  const inputs = ss.getRangeByName(RANGE_INPUTS).getValues();
  const days = [];
  for (let c = 0; c < dates.length; c++) {
    if (!(dates[c] instanceof Date) || done[c] !== '') continue;
    if (inputs.some(function (row) { return row[c] !== ''; })) days.push(Utilities.formatDate(dates[c], tz, 'M/d'));
  }
  return days;
}

/** 月間日報の入力欄（数式のない薄い黄色のセル）と「提出する日」を空にする。 */
function clearInputs_(ss) {
  const block = ss.getRangeByName(RANGE_INPUTS);
  const sheet = block.getSheet();
  const backgrounds = block.getBackgrounds();
  const formulas = block.getFormulas();
  const targets = [];
  for (let r = 0; r < backgrounds.length; r++) {
    let start = -1;
    for (let c = 0; c <= backgrounds[r].length; c++) {
      const input = c < backgrounds[r].length &&
        String(backgrounds[r][c]).toLowerCase() === INPUT_BACKGROUND && !formulas[r][c];
      if (input && start < 0) start = c;
      if (!input && start >= 0) {
        const rowNo = block.getRow() + r;
        targets.push(columnLetter_(block.getColumn() + start) + rowNo + ':' +
          columnLetter_(block.getColumn() + c - 1) + rowNo);
        start = -1;
      }
    }
  }
  if (targets.length) sheet.getRangeList(targets).clearContent();
  const dateCell = ss.getRangeByName(RANGE_SUBMIT_DATE);
  if (dateCell) dateCell.clearContent();
}

/** 翌月の1日（セルに入れる 'yyyy-MM-dd' と、表示用の 'yyyy年M月'）。 */
function nextMonth_(month, tz) {
  const p = Utilities.formatDate(month, tz, 'yyyy-MM').split('-').map(Number);
  const y = p[1] === 12 ? p[0] + 1 : p[0];
  const m = p[1] === 12 ? 1 : p[1] + 1;
  return { text: y + '-' + (m < 10 ? '0' : '') + m + '-01', label: y + '年' + m + '月' };
}

/** 同じ名前のシートがあれば「 (2)」などを付けた名前を返す。 */
function uniqueSheetName_(ss, base) {
  if (!ss.getSheetByName(base)) return base;
  let i = 2;
  while (ss.getSheetByName(base + ' (' + i + ')')) i++;
  return base + ' (' + i + ')';
}

function columnLetter_(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

/** 設定シートのA列から項目名を探し、B列の値を返す（なければ空文字）。 */
function getSetting_(ss, label) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 1) return '';
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === label) return values[i][1] === null ? '' : values[i][1];
  }
  return '';
}

function buildSubject_(ss, dateText) {
  const name = String(getSetting_(ss, '記入者')).trim();
  return '【日報】' + dateText + (name ? ' ' + name : '');
}

function buildBody_(ss, headers, display) {
  const lines = [];
  for (let i = 0; i < headers.length; i++) {
    const label = String(headers[i]).trim();
    const value = String(display[i] === undefined ? '' : display[i]).trim();
    if (!label || !value || value === '-') continue;
    lines.push(label + '：' + value);
  }
  lines.push('', 'スプレッドシート：' + ss.getUrl());
  return lines.join('\n');
}

function removeReminderTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === REMINDER_HANDLER; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}
