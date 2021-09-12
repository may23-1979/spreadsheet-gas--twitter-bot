/**
 * スプレッドシートからTweetに投稿するBOTプログラム
 * 下記サイトを参考にしています。
 * https://cravelweb.com/webdesign/google-apps-script-twitter-tweet-bot
 * 
 * TODO: secret.gs で、各種キーを設定してください。
 * CONSUMER_KEY, TOKEN, CONSUMER_SECRET, TOKEN_SECRET
 */

// Twitter APIの認証とレスポンス取得
function run() {
  var service = getService();
  Logger.log(service.getCallbackUrl());
  var tweet = pickUpTweet(); // ツイートの内容を取得
  if (tweet == '') {
    Logger.log('Tweetが選択できませんでした');
    return false; // 終了
  }
  Logger.log('Tweet Selected : '+tweet);

  if (service.hasAccess()) {
    var url = 'https://api.twitter.com/1.1/statuses/update.json';
    var payload = {
      status: tweet
    };
    var response = service.fetch(url, {
      method: 'post',
      payload: payload
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.authorize();
    Logger.log('URLを確認してください: %s',
        authorizationUrl);
  }
} 

function doGet() {
  return HtmlService.createHtmlOutput(ScriptApp.getService().getUrl());
}

// 認証リセット関数。デバッグ時の初期化用。
function reset() {
  var service = getService();
  service.reset();
}

// サービス設定
function getService() {
  return OAuth1.createService('Twitter')
      .setConsumerKey(CONSUMER_KEY) // コンシューマーキー＆シークレット
      .setConsumerSecret(CONSUMER_SECRET) // コンシューマーシークレット
      .setAccessToken(TOKEN, TOKEN_SECRET) // アクセストークンキー＆シークレット

      // oAuthエンドポイントURL
      .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
      .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
      .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')

      .setCallbackFunction('authCallback') // コールバック関数名 
}

// OAuthコールバック
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('成功しました');
  } else {
    return HtmlService.createHtmlOutput('失敗しました');
  }
}

// Googleスプレッドシートからツイートする内容を取得する
// カスタマイズ
function pickUpTweet() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1"); // シート名
  if (targetSheet.getLastRow() == 1) { return "" } // シートにデータが無い
  var targetCells = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 4);
  Logger.log(JSON.stringify(targetCells, null, 2));

  var cells = targetCells.getValues(); // ツイートのリストを格納

  var tweetText = "";
  var tweetCount = 0;
  var previousCount = cells[0][2]; 
  var lastCount = cells[cells.length-1][2];
Logger.log("previousCount:" + previousCount + ", lastCount:" + lastCount);
  if ( isNaN(previousCount) || previousCount == 0 || previousCount == lastCount) {
    
    // 投稿回数を更新
    targetCells.getCell(1,3).setValue(cells[0][2]+1);
    // 投稿日時
    targetCells.getCell(1,4).setValue(Utilities.formatDate((new Date()), 'Asia/Tokyo', 'yyyy/MM/dd hh:mm:ss'))
    return cells[0][0];
  }

  for (var i = 0, il = cells.length; i < il; i++ ) {
    tweetText = cells[i][0];
    tweetCount = cells[i][2];
    if (isNaN(tweetCount)) { tweetCount = 0; }
    Logger.log("previousCount:" + previousCount);
    if ( tweetCount == 0 || ( previousCount != tweetCount) ) {
      Logger.log("tweetCount:" + tweetCount);
      // 投稿回数を更新
      targetCells.getCell(i+1,3).setValue(tweetCount+1)
      // 投稿日時
      targetCells.getCell(i+1,4).setValue(Utilities.formatDate((new Date()), 'Asia/Tokyo', 'yyyy/MM/dd hh:mm:ss'))
      break;
    }
    previousCount = tweetCount;
  }
  return tweetText;
}