import { DateTime } from 'luxon';

type UpdateData = {
  mailDate: DateTime
  targetDate: DateTime
  quantity?: number,
  price: {
    al?: number, am?: number, as?: number
  }
};

const gas: any = global;

gas._main = () => {
  const SPINACH_SPREAD_SHEET_ID = getProperty('SPINACH_SPREAD_SHEET_ID');

  // メールを検索する条件
  const SEARCH_KEYWORD = 'label:市況-ほうれん草';
  // 設定シートのメール検索日のセル
  const SETTINGS_SHEET_SEARCH_MAIL_DATE = 'B2';
  // Webhook Url
  const WEBHOOK_URLS = getProperty('WEBHOOK_URLS').split('|');

  const spreadSheet = SpreadsheetApp.openById(SPINACH_SPREAD_SHEET_ID);
  const settingsSheet = spreadSheet.getSheetByName('SETTINGS');
  if (!settingsSheet) throw new Error('SETTINGSシートが存在しません');

  const searchMailDateRange = settingsSheet.getRange(SETTINGS_SHEET_SEARCH_MAIL_DATE);
  const searchMailDateValue: string = searchMailDateRange.getValue();

  let searchMailDate: DateTime | undefined = undefined;
  let latestMailDate: DateTime | undefined = undefined;

  // メールの検索キーワードを組み立て
  let searchKeyword = SEARCH_KEYWORD;
  if (searchMailDateValue) {
    searchMailDate = DateTime.fromISO(searchMailDateValue);
    // 最終検索日以降
    searchKeyword += ' after:' + searchMailDate.toFormat('yyyy/MM/dd');
  }

  let messages: GoogleAppsScript.Gmail.GmailMessage[] = [];
  for (const thread of GmailApp.search(searchKeyword)) {
    for (const message of thread.getMessages()) {
      messages.push(message);
    }
  }

  messages = messages.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  Logger.log('searchKeyword: %s, messageCount: %s', searchKeyword, messages.length);

  let updateDatas: UpdateData[] = [];

  // メールから市況データを集計
  for (const message of messages) {
    const plainBody = message.getPlainBody();
    const nextLineGenerator = (function* () {
      for (let line of plainBody.split('\r\n')) {
        line = normalize(line.trim());
        if (line) yield line;
      }
    })();
    const readBody = () => {
      const value = nextLineGenerator.next().value;
      return value ? value : '';
    };

    // Mail
    const mailDate = DateTime.fromMillis(message.getDate().getTime());
    if (searchMailDate && mailDate <= searchMailDate) continue;
    latestMailDate = mailDate;

    for (const url of WEBHOOK_URLS) {
      try {
        UrlFetchApp.fetch(url, {
          method: 'post',
          payload: {
            username: message.getSubject(),
            content: normalize(plainBody)
          }
        });
      } catch (error) {
        console.error(error);
      }
    }

    const mailMonth = mailDate.month;
    // mm月dd日出荷
    const linePd = readBody();
    const pdMatcher = linePd.match(/(.+)月\s*(.+)日出荷/);
    if (!pdMatcher) continue;
    // AL, AM, AS
    const lineAl = readBody();
    const lineAm = readBody();
    const lineAs = readBody();
    // label 出荷数量
    readBody();
    // n箱
    const lineQty = readBody();
    // 本文に年がないので、メールの時刻から取得する
    let year = mailDate.year;
    const month = Number(pdMatcher[1]);
    const day = Number(pdMatcher[2]);
    // 前年の市況が年初に送られてきた場合
    if (month === 12 && mailMonth === 1) year--;

    const updateData: UpdateData = {
      mailDate: mailDate,
      targetDate: DateTime.local(year, month, day),
      quantity: formatNumber(lineQty),
      price: {
        al: formatNumber(lineAl),
        am: formatNumber(lineAm),
        as: formatNumber(lineAs),
      }
    };

    updateDatas.push(updateData);
  };

  updateDatas = updateDatas.sort((a, b) => a.targetDate.diff(b.targetDate).milliseconds);

  // シートに書き出し
  for (const updateData of updateDatas) {
    const sheetName = String(updateData.targetDate.year);
    let sheet = spreadSheet.getSheetByName(sheetName);

    // シートが存在しない場合、雛形からコピーして作成する
    if (!sheet) {
      const templateSheet = spreadSheet.getSheetByName('TEMPLATE');
      if (!templateSheet) throw new Error('SETTINGSシートが存在しません');

      sheet = templateSheet.copyTo(spreadSheet);
      spreadSheet.setActiveSheet(sheet);
      spreadSheet.moveActiveSheet(1);
      sheet.setName(sheetName).showSheet();
    }

    const row = sheet.getLastRow() + 1;
    let column = 1;
    sheet.getRange(row, column++).setValue(updateData.targetDate.toFormat('yyyy/MM/dd'));
    sheet.getRange(row, column++).setValue(updateData.price.al);
    sheet.getRange(row, column++).setValue(updateData.price.am);
    sheet.getRange(row, column++).setValue(updateData.price.as);
    sheet.getRange(row, column++).setValue(updateData.quantity);
    sheet.getRange(row, column++).setValue(updateData.mailDate.toFormat('yyyy/MM/dd HH:mm:ss'));
  };

  // 全てが正常終了したら、設定シートを更新する
  if (latestMailDate) {
    searchMailDateRange.setValue(latestMailDate.toISO());
  }
};

function getProperty(key: string, defaultValue?: any): string {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) return value;
  if (defaultValue) return defaultValue;
  throw new Error(`Undefined property: ${key}`);
}

function normalize(s: string) {
  // F*ck Zenkaku
  return s.replace(/[Ａ-Ｚａ-ｚ０-９]/g,
    s => String.fromCharCode(s.charCodeAt(0) - 65248)).replace(/　/g, ' ');
}

function formatNumber(s: string): number | undefined {
  const ematcher = s.match(/(\d+)$/);
  if (ematcher) {
    return Number(ematcher[1]);
  }
  const smatcher = s.match(/^(\d+)/);
  if (smatcher) {
    return Number(smatcher[1]);
  }

  return undefined;
}
