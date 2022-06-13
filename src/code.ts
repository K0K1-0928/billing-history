const mail = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const googlePlay = getGooglePlayBillingDetails();
  const apple = getAppleBillingDetails();
  const details = googlePlay.concat(apple);
  details.sort((a, b) => {
    return a.date < b.date ? 1 : -1;
  });

  // シート書き込み用の変数
  const values = [];
  for (const detail of details) {
    const value = [
      detail.date,
      detail.title,
      detail.item,
      detail.price,
      detail.from,
      detail.subject,
      detail.permalink,
    ];
    values.push(value);
  }
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
    for (const detail of receipt) {
      const billingItem = detail.match(/^.* \(.*\)\s*\uFFE5.*\d+$/);
      if (billingItem === null) {
        continue;
      }
      const itemCharEnd = detail.indexOf('(');
      const item = detail.substring(0, itemCharEnd).trim();
      const appTitleStart = itemCharEnd + 1;
      const appTitleEnd = detail.indexOf(')');
      const appTitle = detail.substring(appTitleStart, appTitleEnd).trim();
      const priceStart = detail.indexOf('￥') + 1;
      const price = detail.substring(priceStart).trim();

      values.push({
        date: date,
        title: appTitle,
        item: item,
        price: price,
        from: from,
        subject: subject,
        permalink: perma,
      });
    }
  }
  return values;
};

const getAppleBillingDetails = () => {
  const searchQuery =
    'from:(no_reply@email.apple.com) subject:(Apple からの領収書です。) ';
  const messages = getBillingMessages(searchQuery);
  const values = [];

  for (const message of messages) {
    const date = message.getDate();
    const body = message.getPlainBody();
    const from = message.getFrom();
    const subject = message.getSubject();
    const perma = message.getThread().getPermalink();

    const contents = body.split(/(\r\n|\n|\r)/);
    const receipt = contents.filter((content) => content.includes('¥'));
    for (const detail of receipt) {
      const billingItem = detail.match(/^.*,.*\s*App 内課金\s*¥.*\d+$/);
      if (billingItem === null) {
        continue;
      }
      const appTitleEnd = detail.indexOf(', ');
      const appTitle = detail.substring(0, appTitleEnd).trim();
      const itemStart = appTitleEnd + 1;
      const itemEnd = detail.indexOf('App 内課金') - 1;
      const item = detail.substring(itemStart, itemEnd).trim();
      const priceStart = detail.indexOf('¥') + 1;
      const price = detail.substring(priceStart).trim();

      values.push({
        date: date,
        title: appTitle,
        item: item,
        price: price,
        from: from,
        subject: subject,
        permalink: perma,
      });
    }
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
