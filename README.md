# Billing History

G Mail から Apple/Google の課金明細メールを取得し、あなたの課金履歴をリスト化するプログラムです。  
コンテナバインド型の Google Apps Script ですので、  
以下のリンクからスプレッドシートをコピーして使用するのが簡単です。  
[Billing History](https://docs.google.com/spreadsheets/d/1iZx57lF7GugtRb6TjqZXRP2PJpJLe0vSR0fm4YdN_VY/edit?usp=sharing)

## How To Use

スプレッドシートをコピーして使用する事を前提とします。

### 初回

1. [Billing History](https://docs.google.com/spreadsheets/d/1iZx57lF7GugtRb6TjqZXRP2PJpJLe0vSR0fm4YdN_VY/edit?usp=sharing)を開き、 ファイル > コピーを作成 でスプレッドシートをコピーします。
2. ツール > スクリプトエディタを開きます。
3. スクリプトエディタから、1 度 `getBillingHistory()` を実行します。この際、スプレッドシートと G Mail への編集・閲覧権限を要求されるので、許可します。
4. 実行完了後、メールから取得したデータがスプレッドシートに書き込まれます。

### 2 回目以降

1. スプレッドシートのメニューボタンから、「課金履歴取得」ボタンをクリックし、実行します。
2. 実行完了後、メールから取得したデータがスプレッドシートに書き込まれます。
