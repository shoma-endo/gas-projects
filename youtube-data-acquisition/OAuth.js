var token_url = 'https://oauth2.googleapis.com/token';
var client_id = PropertiesService.getScriptProperties().getProperty('CLIENT_ID');
var client_secret = PropertiesService.getScriptProperties().getProperty('CLIENT_SECRET');
var code = "4/0ARtbsJpFFL1tV-Oqk32aUsunFcEkq2WfHACEdjIJWpby-2ZhTh28uANnQdvC_6O5n-u8tg"

/************************************
認可コードを利用してトークン情報を取得して返す（初回のみ使用する）
次回からはリフレッシュトークンを使ってトークン情報を更新できる
************************************/
function getAccessToken(obj, e) {
  var payload = {
    'grant_type': 'authorization_code',
    'client_id': obj['client_id'],
    'client_secret': obj['client_secret'],
    'code': code,
    'redirect_uri': "https://www.google.com/"
  }
  
  var response = UrlFetchApp.fetch(token_url, getOptions(payload));
  return response;
}

/************************************
refresh_tokenを使って更新したトークン情報を返す
************************************/
function runRefresh(obj, refresh_token) {  
  var payload = {
    'grant_type': 'refresh_token',
    'client_id': obj['client_id'],
    'client_secret': obj['client_secret'],
    'refresh_token': refresh_token
  }
  var response = UrlFetchApp.fetch(token_url, getOptions(payload));
  return response;
}

/************************************
optionsを作って返す
************************************/
function getOptions(payload) {
  var options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload,
    'muteHttpExceptions' : true
  }
  return options;
}

/************************************
渡されたmethodで実行する
************************************/
function runMethod(method, url, access_token, payload) {
  var options = {
    'method': method,
    'contentType': 'application/json',
    'headers': { 'Authorization': 'Bearer ' + access_token },
    'payload': payload,
    'muteHttpExceptions': true
  }
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/************************************
アクセストークン、リフレッシュトークン取得時に使うオブジェクト
************************************/
var appInfo = {
  'client_id': client_id,
  'client_secret': client_secret,
  'redirect_uri': "https://www.google.com/"
}

/************************************
Web認証用URLを開いたときに動く処理
************************************/
function doGet(e) {
  var response = getAccessToken(appInfo, e);
  setScriptProperties(JSON.parse(response));
  //getCompanies();
  //getWalletables();
  return ContentService.createTextOutput(response);// ブラウザに表示する
}

/************************************
リフレッシュトークンを使ってトークン情報を更新してスクリプトプロパティを上書きする
************************************/
function refreshTokens() {
  var response = runRefresh(appInfo, getScriptProperties('refresh_token'))
  setScriptProperties(JSON.parse(response));
}

/************************************
PropertiesService
************************************/
function setScriptProperties(jobj) {// スクリプトのプロパティに値を保存する
  PropertiesService.getScriptProperties().setProperties(jobj);
}

function getScriptProperties(key) {// スクリプトのプロパティから値を取得する
  return PropertiesService.getScriptProperties().getProperty(key);
}