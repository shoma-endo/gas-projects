var key = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
var ss = SpreadsheetApp.getActiveSpreadsheet()
var now = new Date();
var today = new Date(now.getFullYear(),now.getMonth(),now.getDate());

function getAccount() {

  const sheets = ss.getSheets()
  let dates = ""
  let url = "https://www.googleapis.com/youtube/v3/channels?key=" + key + "&part=snippet,statistics&id="

  let response = ""
  let json = ""
  let arr = []

  let id = ""

  for(let i=0;i<sheets.length;i++){
    arr=[]
    if(sheets[i].getSheetName().match(/アカウント情報/)){
      id = sheets[i].getRange(1,2).getValue()
      dates = sheets[i].getRange(6,1,sheets[i].getLastRow()-5,1).getValues()

      //for(let i=0;i<datas.length;i++){
        //if(datas[i][0] !== ""){
          response = UrlFetchApp.fetch(url + id).getContentText();
          json = JSON.parse(response)
          sheets[i].getRange(2,2).setValue(json["items"]["0"]["snippet"]["title"])
          sheets[i].getRange(3,2).setValue(json["items"]["0"]["snippet"]["publishedAt"].split("T")[0])
          arr.push([json["items"]["0"]["statistics"]["subscriberCount"],json["items"]["0"]["statistics"]["videoCount"],json["items"]["0"]["statistics"]["viewCount"]])
        //}
      //}

      if(arr.length>0){
        for(let j=0;j<dates.length;j++){
          if(dates[j][0].getTime() == today.getTime()){
            sheets[i].getRange(j+6,2,arr.length,arr[0].length).setValues(arr)
          }
        }
      }
    }

  }

}
