"use strict";
const mail = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getActiveSheet();
    const googlePlay = getGooglePlayBillingDetails();
    const apple = getAppleBillingDetails();
    // シートへの書き込み用の変数
    const values = googlePlay.concat(apple);
    values.sort((a, b) => {
        return a[0] < b[0] ? 1 : -1;
    });
    if (values.length > 0) {
        sh.getRange(2, 1, values.length, values[0].length).setValues(values);
    }
};
const getGooglePlayBillingDetails = () => {
    const searchQuery = 'from:(googleplay-noreply@google.com) subject:(Google Play のご注文明細) ';
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
        const item = detail.substring(0, itemCharEnd).trim();
        const appTitleStart = itemCharEnd + 1;
        const appTitleEnd = detail.indexOf(')');
        const appTitle = detail.substring(appTitleStart, appTitleEnd).trim();
        const priceStart = detail.indexOf('￥') + 1;
        const price = detail.substring(priceStart).trim();
        values.push([date, appTitle, item, price, from, subject, perma]);
    }
    return values;
};
const getAppleBillingDetails = () => {
    const searchQuery = 'from:(no_reply@email.apple.com) subject:(Apple からの領収書です。) ';
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
            values.push([date, appTitle, item, price, from, subject, perma]);
        }
    }
    return values;
};
const getBillingMessages = (searchQuery) => {
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
