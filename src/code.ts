const mail = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const googlePlay = getGooglePlayBillingDetails();

  // シートへの書き込み用の変数
  const varArray: Array<Array<string | GoogleAppsScript.Base.Date>> = [];
  const values = varArray.concat(googlePlay);
  if (values.length > 0) {
    sh.getRange(2, 1, values.length, values[0].length).setValues(values);
  }
};

const getGooglePlayBillingDetails = () => {
  const searchQuery =
    'from:(googleplay-noreply@google.com) subject:(Google Play のご注文明細) ';
  const messages = getBillingMessages(searchQuery);
  const values = [];

  for (const message of messages) {
    const date = message.getDate();
    const body = message.getPlainBody();
    const from = message.getFrom();
    const subject = message.getSubject();
    const perma = message.getThread().getPermalink();

    const contents = body.split(/(\r\n|\n|\r)/);
    const receipt = contents.filter((content) => content.includes('￥'));
    const detail = receipt[0];
    const billingItem = detail.match(/^.* \(.*\)\s*\uFFE5.*\d+$/);
    if (billingItem === null) {
      continue;
    }

    const itemCharEnd = detail.indexOf('(');
    const item = detail.substring(0, itemCharEnd);
    const gameTitleStart = itemCharEnd + 1;
    const gameTitleEnd = detail.indexOf(')');
    const gameTitle = detail.substring(gameTitleStart, gameTitleEnd);
    const priceStart = detail.indexOf('￥');
    const price = detail.substring(priceStart);

    values.push([date, item, gameTitle, price, from, subject, perma]);
  }
  return values;
};

const getBillingMessages = (searchQuery: string) => {
  const threads = GmailApp.search(searchQuery, 0, 500);
  const thread = GmailApp.getMessagesForThreads(threads);
  const values = [];
  for (const messages of thread) {
    for (const message of messages) {
      values.push(message);
    }
  }
  return values;
};
