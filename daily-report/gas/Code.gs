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
 *   本日の日報を提出        … 「提出データ」2行目を「日次ログ」に1行保存（同じ日付があれば上書き確認）。
 *                              設定シートに提出先メールアドレスがあればメールも送る
 *   入力欄を空にして翌日へ進む … 「本日の日報」の入力欄（薄い黄色のセル）を空にし、日付を翌日にする
 *   毎日のリマインドを設定    … 設定シートの「リマインドの時刻」に、その日の日報が未提出ならメールで知らせる
 *   リマインドを解除
 */

const SHEET_FORM = '本日の日報';
const SHEET_SUBMIT = '提出データ';
const SHEET_LOG = '日次ログ';
const SHEET_SETTINGS = '設定';
const DATE_RANGE_NAME = 'REPORT_DATE'; // 本日の日報の日付セル（名前付き範囲）
const INPUT_BACKGROUND = '#fff2cc'; // 入力セルの背景色
const LOG_FIRST_ROW = 3; // 日次ログのデータ開始行（1〜2行目は見出し、3行目は記入例）
const REMINDER_HANDLER = 'remindIfNotSubmitted';
const DEFAULT_REMINDER_HOUR = 21;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報')
    .addItem('本日の日報を提出', 'submitDailyReport')
    .addItem('入力欄を空にして翌日へ進む', 'clearFormForNextDay')
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
  if (!submit || !log) {
    ui.alert('「' + SHEET_SUBMIT + '」または「' + SHEET_LOG + '」シートが見つかりません。');
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
    ui.alert('「' + SHEET_FORM + '」の日付を入力してから提出してください。');
    return;
  }
  const tz = ss.getSpreadsheetTimeZone();
  const dateText = Utilities.formatDate(date, tz, 'yyyy/MM/dd');

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

  const next = ui.alert(
    '提出しました',
    dateText + ' の日報を日次ログの ' + target + ' 行目に保存しました。' + mailNote +
      '\n\n入力欄を空にして、日付を翌日に進めますか？',
    ui.ButtonSet.YES_NO);
  if (next === ui.Button.YES) clearForm_(ss, nextDayText_(date, tz));
}

function clearFormForNextDay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '入力欄を空にする',
    '「' + SHEET_FORM + '」の入力欄（薄い黄色のセル）を空にして、日付を翌日に進めます。' +
      '提出していない内容は消えます。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const tz = ss.getSpreadsheetTimeZone();
  const dateCell = ss.getRangeByName(DATE_RANGE_NAME);
  const current = dateCell ? dateCell.getValue() : null;
  const base = current instanceof Date ? current : new Date();
  clearForm_(ss, nextDayText_(base, tz));
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
      '\n\n「' + SHEET_FORM + '」に入力し、メニュー「日報」→「本日の日報を提出」で提出してください。',
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

/** 本日の日報の入力欄（数式のない薄い黄色のセル）を空にし、日付セルに nextDate を入れる。 */
function clearForm_(ss, nextDate) {
  const form = ss.getSheetByName(SHEET_FORM);
  const dateCell = ss.getRangeByName(DATE_RANGE_NAME);
  const dateA1 = dateCell ? dateCell.getA1Notation() : '';
  const range = form.getDataRange();
  const backgrounds = range.getBackgrounds();
  const formulas = range.getFormulas();
  const targets = [];
  for (let r = 0; r < backgrounds.length; r++) {
    for (let c = 0; c < backgrounds[r].length; c++) {
      if (String(backgrounds[r][c]).toLowerCase() !== INPUT_BACKGROUND || formulas[r][c]) continue;
      const a1 = columnLetter_(c + 1) + (r + 1);
      if (a1 !== dateA1) targets.push(a1);
    }
  }
  if (targets.length) form.getRangeList(targets).clearContent();
  if (dateCell && nextDate) dateCell.setValue(nextDate);
}

/** 日付の翌日を 'yyyy-MM-dd' の文字列で返す（セルに入れると日付として扱われる）。 */
function nextDayText_(date, tz) {
  const parts = Utilities.formatDate(date, tz, 'yyyy-MM-dd').split('-').map(Number);
  const next = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + 1));
  return Utilities.formatDate(next, 'UTC', 'yyyy-MM-dd');
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
