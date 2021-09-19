/**
 * スプレッドシートからTweetに投稿するBOTプログラム
 * 下記サイトを参考にしています。
 * https://cravelweb.com/webdesign/google-apps-script-twitter-tweet-bot
 * 
 * TODO: secret.gs で、各種キーを設定してください。
 * CONSUMER_KEY, TOKEN, CONSUMER_SECRET, TOKEN_SECRET
 */

const SHEET_NAME_TITLE = "基本情報";
const SHEET_NAME_TWEET = "シート-テスト用";

// Twitter APIの認証とレスポンス取得
function run() {
  var service = getService();
  Logger.log(service.getCallbackUrl());
  var targetCulumn = pickUpTweet(); // ツイートの内容を取得
  var tweetText = targetCulumn.getCell(1,1).getValue();
  var tweetCount = targetCulumn.getCell(1,4);

  if (tweetText == '') {
    Logger.log('Tweetが選択できませんでした');
    return false; // 終了
  }
  // タイトル(ハッシュタグ)を取得
  tweetText += "\n" + getTweetTitle();
  Logger.log('tweetText:' + tweetText);

  if (service.hasAccess()) {
    var url = 'https://api.twitter.com/1.1/statuses/update.json';
    var payload = {
      status: tweetText
    };

    try {
      var response = service.fetch(url, {
        method: 'post',
        payload: payload
      });
      var result = JSON.parse(response.getContentText());
      Logger.log(JSON.stringify(result, null, 2));
      Logger.log("Tweet成功");

      // 投稿回数を更新
      tweetCount.setValue(tweetCount.getValue()+1);
      // 投稿日時
      targetCulumn.getCell(1,5).setValue(Utilities.formatDate((new Date()), 'Asia/Tokyo', 'yyyy/MM/dd hh:mm:ss'));
    } catch(error) {
      Logger.log("失敗");
      console.error(getErrorInfo(error));
      
      // http://westplain.sakuraweb.com/translate/twitter/API-Overview/Error-Codes-and-Responses.cgi
      targetCulumn.getCell(1,5).setValue("投稿失敗\n" + error.message);
    }

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
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TWEET); // シート名
  if (targetSheet.getLastRow() == 1) { return [] } // シートにデータが無い
  var targetCells = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 5);
  Logger.log(JSON.stringify(targetCells, null, 2));

  var cells = targetCells.getValues(); // ツイートのリストを格納

  var tweetCount = 0;
  var previousCount = cells[0][3];
  var lastCount = cells[cells.length-1][3];
  Logger.log("previousCount:" + previousCount + ", lastCount:" + lastCount);
  if ( isNaN(previousCount) || previousCount == 0 || previousCount == lastCount) {   
    Logger.log("pickUpTweet() > 1行目をリターン");
    return targetSheet.getRange(2,1, 2,5);
  }

  for (var i = 0, il = cells.length; i < il; i++ ) {
    tweetCount = cells[i][3];
    if (isNaN(tweetCount)) { tweetCount = 0; }
    Logger.log("previousCount:" + previousCount);
    if ( tweetCount == 0 || ( previousCount != tweetCount) ) {
      Logger.log("pickUpTweet() > i+2行目をリターン, i:" + i);
      return targetSheet.getRange(i+2,1, i+2,5);
    }
    previousCount = tweetCount;
  }
  return targetSheet.getRange(i+2,1, i+2,5);
}

// Googleスプレッドシートからツイートする内容を取得する
function getTweetTitle() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TITLE); // シート名
  if (targetSheet.getLastRow() == 1) { return "" } // シートにデータが無い
  var targetCells = targetSheet.getRange(2, 1);
  return targetCells.getValue();
}

function getErrorInfo(error){
  var errorInfo = "[名前] " + error.name + "\n" +
         "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n" +
         "[メッセージ]" + error.message + "\n" +      
         "[StackTrace]\n" + error.stack;
  return errorInfo;
}