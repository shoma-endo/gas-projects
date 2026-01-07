var now = new Date();
var today = new Date(now.getFullYear(),now.getMonth(),now.getDate())
var one_year_ago = new Date(now.getFullYear(),now.getMonth(),now.getDate())
one_year_ago.setDate(one_year_ago.getDate()-730)
var days_2_later = new Date(now.getFullYear(),now.getMonth(),now.getDate())
days_2_later.setDate(days_2_later.getDate()+2)

var options = {
    'method' : 'get',
    'muteHttpExceptions': true,
  };

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet_token = ss.getSheetByName("トークンとID")
var business_id = sheet_token.getRange(2,2).getValue()
var token = sheet_token.getRange(1,2).getValue()


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

function all_Triggers(){
  delTrigger()
  delTrigger2()
  delTrigger3()
  delTrigger_Reel()
  delTrigger2_Reel()
  delTrigger3_Reel()
  delTrigger4_Reel()
}
