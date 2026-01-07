var now = new Date();
var today = new Date(now.getFullYear(),now.getMonth(),now.getDate())
var one_year_ago = new Date(now.getFullYear(),now.getMonth(),now.getDate())
one_year_ago.setDate(one_year_ago.getDate()-365)

var options = {
    'method' : 'get',
    'muteHttpExceptions': true,
  };

var ss = SpreadsheetApp.getActiveSpreadsheet(); 
var id = ss.getId();
var sheet_token = ss.getSheetByName("トークンとID")
var business_id = sheet_token.getRange(2,2).getValue()
var token = sheet_token.getRange(1,2).getValue()


function onOpen() {
  var subMenus = [];
  subMenus.push({
    name: "トリガー設定 ※初回1回のみクリック",
    functionName: "first_time" 
  });
  ss.addMenu("GAS起動", subMenus);
}


function first_time(){ //初回起動時のみ設定
  ScriptApp.newTrigger('getCompete').timeBased().everyHours(2).create();
  //ScriptApp.newTrigger('countCheck').forSpreadsheet(ss).onEdit().create();
}

function getComment(arr){

  let comment_text = ""
  let comment = ""

  //コメント検索
  for(let i=0;i<arr.length;i++){
    let url = "https://graph.facebook.com/v12.0/" + arr[i][1]+ "/comments?access_token=" + token
    let encodedURI = encodeURI(url);
    let response = UrlFetchApp.fetch(encodedURI,options).getContentText();
    let json = JSON.parse(response)

    comment_text = ""
    comment = ""
    for (var num in json["data"]) {
      comment = json["data"][num]["text"]
      if(comment.match(/#/)){
        comment = comment.replace("#","")
        comment_text += "," + comment
      }
    }
    arr[i].push(comment_text)
  }
  
}
