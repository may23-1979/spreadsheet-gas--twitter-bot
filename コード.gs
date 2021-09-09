// TODO: secret.gs で、各種キーを設定してください。
// CONSUMER_KEY, TOKEN, CONSUMER_SECRET, TOKEN_SECRET
//
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
function pickUpTweet() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1"); // シート名
  if (targetSheet.getLastRow() == 1) { return "" } // シートにデータが無い
  var cells = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).getValues(); // ツイートのリストを格納
  var grossWeight = 0;
  for (var i = 0; i < cells.length; i++ ) { // ウェイトの合計値を算出
    grossWeight += cells[i][1];
  }
  if (grossWeight == 0) { return ""; } // ウェイト総数がゼロ
  var targetWeight = grossWeight * Math.random(); // ウェイト合計値を最大値としてランダムでターゲットの数値を生成

  var tweetText = "";
  for (var i = 0, il = cells.length; i < il; i++ ) {
    targetWeight -= cells[i][1]; // セルに記入されたウェイトの値を減算
    if (targetWeight < 0) { // ターゲットの数値がマイナスになったらそのセルの内容を返す
      tweetText = cells[i][0];
      break;
    }
  }
  return tweetText;
}