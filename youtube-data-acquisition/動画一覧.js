function getList() {

  const sheets = ss.getSheets()
  
  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){
      let url = "https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=" + sheets[i-1].getRange(1,2).getValue() + "&maxResults=50&order=date&type=video&key=" + key

      let response = UrlFetchApp.fetch(url).getContentText();
      let json = JSON.parse(response)
      let arr = []
      let ts = ""
      let date = ""
      let date_1hours = ""
      let date_3hours = ""
      let date_7days = ""

      for(var j=0;j<json["items"].length;j++){
        ts = json["items"][j]["snippet"]["publishedAt"]
        date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
        date_1hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+1,date.getMinutes())
        date_3hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+3,date.getMinutes())
        date_7days = new Date(date.getFullYear(),date.getMonth(),date.getDate()+7,date.getHours()+3,date.getMinutes())
        arr.push([json["items"][j]["id"]["videoId"],json["items"][j]["snippet"]["title"],"https://www.youtube.com/watch?v="+json["items"][j]["id"]["videoId"],"","",date,date_1hours,date_3hours,date_7days])
      }

      arr.sort(sorting_asc)

      if(sheets[i].getLastRow() > 2){
        let datas = sheets[i].getRange(3,1,sheets[i].getLastRow()-2,1).getValues();

        let flg = false
        let arr2 = []
        for(let i=0;i<arr.length;i++){
          flg = false
          for(let j=0;j<datas.length;j++){
            if(arr[i][0] == datas[j][0]){
              flg = true
              break;
            }
          }
          if(flg == false){
            arr2.push(arr[i])
          }
        }

        if(arr2.length>0){
          sheets[i].getRange(sheets[i].getLastRow()+1,1,arr2.length,arr2[0].length).setValues(arr2)
        }
      }else{
        if(arr.length>0){
          sheets[i].getRange(sheets[i].getLastRow()+1,1,arr.length,arr[0].length).setValues(arr)
        }
      }
    }
  }
  getVideos()
  all_Triggers()
}

function sorting_asc(a, b){
  if(a[5] < b[5]){
    return -1;
  }else if(a[5] > b[5] ){
    return 1;
  }else{
   return 0;
  }
}


function getVideos() {

  const sheets = ss.getSheets()
  let datas = ""
  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){
      datas = sheets[i].getRange(3,1,sheets[i].getLastRow()-2,1).getValues()
      let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id="

      let response = ""
      let json = ""
      let arr = []
      let tags = ""
      let arr2 = []
      let duration = ""
      for(let i=0;i<datas.length;i++){
        if(datas[i][0] !== ""){
          response = UrlFetchApp.fetch(url + datas[i][0]).getContentText();
          json = JSON.parse(response)
          arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])
          if(String(json["items"]["0"]["snippet"]["tags"]) !== "undefined"){
            for(let j=0;j<json["items"]["0"]["snippet"]["tags"].length;j++){
              tags += json["items"]["0"]["snippet"]["tags"][j] + ","
            }
          }
          duration = json["items"]["0"]["contentDetails"]["duration"]
          duration = duration.replace("PT","").replace("H","時間").replace("M","分").replace("S","秒")
          arr2.push([duration,tags])
        }
      }

      if(arr.length>0){
        sheets[i].getRange(3,4,arr2.length,arr2[0].length).setValues(arr2)
        sheets[i].getRange(3,10,arr.length,arr[0].length).setValues(arr)
      }
    }
  }

}

function getVideos_1hours(){

  const sheets = ss.getSheets()
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){

      let time = ""
      let search_id = ""
      let row_num = ""

      datas = sheets[i].getRange(3,1,sheets[i].getLastRow()-2,sheets[i].getLastColumn()).getValues()
      for(let i=0;i<datas.length;i++){
        time = datas[i][6]
        if(time !== ""){
          if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
            search_id = datas[i][0]
            row_num = i+3
            break;
          }
        }
      }

      let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

      let response = ""
      let json = ""
      let arr = []

      response = UrlFetchApp.fetch(url).getContentText();
      json = JSON.parse(response)
      arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


      if(arr.length>0){
        sheets[i].getRange(row_num,13,arr.length,arr[0].length).setValues(arr)
      }
    }
  }

}

function getVideos_3hours(){

  const sheets = ss.getSheets()
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){

      let time = ""
      let search_id = ""
      let row_num = ""
      datas = sheets[i].getRange(3,1,sheets[i].getLastRow()-2,sheets[i].getLastColumn()).getValues()
      for(let i=0;i<datas.length;i++){
        time = datas[i][7]
        if(time !== ""){
          if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
            search_id = datas[i][0]
            row_num = i+3
            break;
          }
        }
      }

      let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

      let response = ""
      let json = ""
      let arr = []

      response = UrlFetchApp.fetch(url).getContentText();
      json = JSON.parse(response)
      arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


      if(arr.length>0){
        sheets[i].getRange(row_num,16,arr.length,arr[0].length).setValues(arr)
      }
    }
  }

}


function getVideos_7days(){

  const sheets = ss.getSheets()
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){

      let time = ""
      let search_id = ""
      let row_num = ""
      datas = sheets[i].getRange(3,1,sheets[i].getLastRow()-2,sheets[i].getLastColumn()).getValues()
      for(let i=0;i<datas.length;i++){
        time = datas[i][8]
        if(time !== ""){
          if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
            search_id = datas[i][0]
            row_num = i+3
            break;
          }
        }
      }

      let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

      let response = ""
      let json = ""
      let arr = []

      response = UrlFetchApp.fetch(url).getContentText();
      json = JSON.parse(response)
      arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


      if(arr.length>0){
        sheets[i].getRange(row_num,19,arr.length,arr[0].length).setValues(arr)
      }
    }
  }

}



function delTrigger() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_1hours"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheets = ss.getSheets()
  let datas = ""
  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){
      datas = sheets[i].getRange(3,7,sheets[i].getLastRow()-2,1).getValues()

      for(let i=0;i<datas.length;i++){
        if(datas[i][0].getTime() >= now.getTime()){
          ScriptApp.newTrigger('getVideos_1hours').timeBased().at(datas[i][0]).create();
        }
      }
    }
  }
}

function delTrigger2() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_3hours"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheets = ss.getSheets()
  let datas = ""
  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){
      datas = sheets[i].getRange(3,8,sheets[i].getLastRow()-2,1).getValues()

      for(let i=0;i<datas.length;i++){
        if(datas[i][0].getTime() >= now.getTime()){
          ScriptApp.newTrigger('getVideos_3hours').timeBased().at(datas[i][0]).create();
        }
      }
    }
  }
}

function delTrigger3() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_7days"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheets = ss.getSheets()
  let datas = ""
  for(let i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName().match(/動画一覧/)){
      datas = sheets[i].getRange(3,9,sheets[i].getLastRow()-2,1).getValues()

      for(let i=0;i<datas.length;i++){
        if(datas[i][0].getTime() >= now.getTime()){
          ScriptApp.newTrigger('getVideos_7days').timeBased().at(datas[i][0]).create();
        }
      }
    }
  }
}


function all_Triggers(){
  delTrigger()
  delTrigger2()
  delTrigger3()
}



function getList_mine() {

  const sheet = ss.getSheetByName("自アカウント動画")
  let url = "https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=UCz7ehWImLdo__rnenoFTwJQ&maxResults=50&order=date&type=video&key=" + key

  let response = UrlFetchApp.fetch(url).getContentText();
  let json = JSON.parse(response)
  let arr = []
  let ts = ""
  let date = ""
  let date_7days = ""
  let date_1months = ""
  let date_3months = ""

  for(var j=0;j<json["items"].length;j++){
    ts = json["items"][j]["snippet"]["publishedAt"]
    date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
    date_7days = new Date(date.getFullYear(),date.getMonth(),date.getDate()+7,date.getHours(),date.getMinutes())
    date_1months = new Date(date.getFullYear(),date.getMonth()+1,date.getDate(),date.getHours(),date.getMinutes())
    date_3months = new Date(date.getFullYear(),date.getMonth()+3,date.getDate(),date.getHours(),date.getMinutes())
    arr.push([json["items"][j]["id"]["videoId"],"",json["items"][j]["snippet"]["title"],"https://www.youtube.com/watch?v="+json["items"][j]["id"]["videoId"],"","",date,date_7days,date_1months,date_3months])
  }

  arr.sort(sorting_asc)

  if(sheet.getLastRow() > 2){
    let datas = sheet.getRange(3,1,sheet.getLastRow()-2,1).getValues();

    let flg = false
    let arr2 = []
    for(let i=0;i<arr.length;i++){
      flg = false
      for(let j=0;j<datas.length;j++){
        if(arr[i][0] == datas[j][0]){
          flg = true
          break;
        }
      }
      if(flg == false){
        arr2.push(arr[i])
      }
    }

    if(arr2.length>0){
      sheet.getRange(sheet.getLastRow()+1,1,arr2.length,arr2[0].length).setValues(arr2)
    }
  }else{
    if(arr.length>0){
      sheet.getRange(sheet.getLastRow()+1,1,arr.length,arr[0].length).setValues(arr)
    }
  }

  const requests = {updateDimensionProperties: {
    properties: {pixelSize: 150},
    range: {sheetId: sheet.getSheetId(), startIndex:2, endIndex: sheet.getLastRow(), dimension: "ROWS"},
    fields: "pixelSize"
  }};
  Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());

  //getVideos_mine()
  //all_Triggers_mine()
}


function getVideos_mine() {

  refreshTokens()

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = sheet.getRange(3,1,sheet.getLastRow()-2,7).getValues()
  let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id="

  let sdate = ""
  let edate = Utilities.formatDate(today,"JST","yyyy-MM-dd")
  let url2 = ""

  let response = ""
  let json = ""
  let arr = []
  let tags = ""
  let arr2 = []
  let duration = ""

  let response2 = ""
  let json2 = ""
  let sec = ""
  let arr3 = []
  let thumbnails = ""
  let arr4 = []
  for(let i=0;i<datas.length;i++){
    tags = ""
    if(datas[i][0] !== ""){

      response = UrlFetchApp.fetch(url + datas[i][0]).getContentText();
      json = JSON.parse(response)
      arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])
      if(String(json["items"]["0"]["snippet"]["tags"]) !== "undefined"){
        for(let j=0;j<json["items"]["0"]["snippet"]["tags"].length;j++){
          tags += json["items"]["0"]["snippet"]["tags"][j] + ","
        }
      }

      if(String(json["items"]["0"]["snippet"]["thumbnails"]["high"]["url"]) !== "undefined"){
        thumbnails = json["items"]["0"]["snippet"]["thumbnails"]["high"]["url"]
      }else{
        thumbnails = ""//json["items"]["0"]["snippet"]["thumbnails"]["default"]["url"]
      }

      if(thumbnails !== ""){
        arr4.push(['=IMAGE("' + thumbnails + '")'])
      }else{
        arr4.push([""])
      }
      duration = json["items"]["0"]["contentDetails"]["duration"]
      duration = duration.replace("PT","").replace("H","時間").replace("M","分").replace("S","秒")
      arr2.push([duration,tags])

      sec = ""
      sdate = Utilities.formatDate(datas[i][6],"JST","yyyy-MM-dd")
      //sdate = "2022-01-01"
      url2 = "https://youtubeanalytics.googleapis.com/v2/reports?dimensions=video&filters=video==" + datas[i][0] + "&ids=channel==MINE&metrics=likes,averageViewDuration&startDate=" + sdate + "&endDate=" + edate

      var accessToken = getScriptProperties('access_token');
      var params = {
        "method" : "get",
        "headers" : {"Authorization":"Bearer " + accessToken}
      };
      response2 = UrlFetchApp.fetch(url2,params).getContentText();
      json2 = JSON.parse(response2)
      Logger.log(json2)
      if(json2["rows"].length>0){
        sec = json2["rows"]["0"]["1"]
      }

      let min = Math.floor(sec / 60);
      let rem =  Math.floor(sec % 60);
      arr3.push([min + "分" + rem + "秒"])

    }
  }

  if(arr.length>0){
    sheet.getRange(3,2,arr4.length,arr4[0].length).setValues(arr4)
    sheet.getRange(3,5,arr2.length,arr2[0].length).setValues(arr2)
    sheet.getRange(3,11,arr.length,arr[0].length).setValues(arr)
    sheet.getRange(3,14,arr3.length,arr3[0].length).setValues(arr3)
  }

  SpreadsheetApp.flush()
  getRate()
}

function getRate(){
  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = sheet.getRange(3,1,sheet.getLastRow()-2,13).getValues()
  let time1 = ""
  let time1_sec = ""
  let time2 = ""
  let time2_sec = ""
  let arr = []
  for(let i=0;i<datas.length;i++){
    time1 = datas[i][3]
    time2 = datas[i][12]
    if(time1 !== "" && time2 !== ""){
      if(time1.match("分")){
        time1_sec = (Number(time1.split("分")[0]) * 60) + Number(time1.split("分")[1].split("秒")[0])
      }else{
        time1_sec = time1.split("秒")[0]
      }

      if(time2.match("分")){
        time2_sec = (Number(time2.split("分")[0]) * 60) + Number(time2.split("分")[1].split("秒")[0])
      }else{
        time2_sec = time2.split("秒")[0]
      }
      arr.push([time2_sec / time1_sec])
    }
  }
  sheet.getRange(3,14,arr.length,arr[0].length).setValues(arr)
}


function getVideos_1hours_mine(){

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  let time = ""
  let search_id = ""
  let row_num = ""

  datas = sheet.getRange(3,1,sheet.getLastRow()-2,sheet.getLastColumn()).getValues()
  for(let i=0;i<datas.length;i++){
    time = datas[i][6]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][0]
        row_num = i+3
        break;
      }
    }
  }

  let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

  let response = ""
  let json = ""
  let arr = []

  response = UrlFetchApp.fetch(url).getContentText();
  json = JSON.parse(response)
  arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


  if(arr.length>0){
    sheet.getRange(row_num,15,arr.length,arr[0].length).setValues(arr)
  }

}

function getVideos_3hours_mine(){

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  let time = ""
  let search_id = ""
  let row_num = ""

  datas = sheet.getRange(3,1,sheet.getLastRow()-2,sheet.getLastColumn()).getValues()
  for(let i=0;i<datas.length;i++){
    time = datas[i][7]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][0]
        row_num = i+3
        break;
      }
    }
  }

  let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

  let response = ""
  let json = ""
  let arr = []

  response = UrlFetchApp.fetch(url).getContentText();
  json = JSON.parse(response)
  arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


  if(arr.length>0){
    sheet.getRange(row_num,18,arr.length,arr[0].length).setValues(arr)
  }

}


function getVideos_7days_mine(){

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = ""
  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  let time = ""
  let search_id = ""
  let row_num = ""

  datas = sheet.getRange(3,1,sheet.getLastRow()-2,sheet.getLastColumn()).getValues()
  for(let i=0;i<datas.length;i++){
    time = datas[i][8]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][0]
        row_num = i+3
        break;
      }
    }
  }

  let url = "https://www.googleapis.com/youtube/v3/videos?key=" + key + "&part=statistics,snippet,contentDetails&id=" + search_id

  let response = ""
  let json = ""
  let arr = []

  response = UrlFetchApp.fetch(url).getContentText();
  json = JSON.parse(response)
  arr.push([json["items"]["0"]["statistics"]["viewCount"],json["items"]["0"]["statistics"]["likeCount"],json["items"]["0"]["statistics"]["commentCount"]])


  if(arr.length>0){
    sheet.getRange(row_num,21,arr.length,arr[0].length).setValues(arr)
  }

}



function delTrigger_mine() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_1hours_mine"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = sheet.getRange(3,7,sheet.getLastRow()-2,1).getValues()

  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getVideos_1hours_mine').timeBased().at(datas[i][0]).create();
    }
  }

}

function delTrigger2_mine() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_3hours_mine"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = sheet.getRange(3,8,sheet.getLastRow()-2,1).getValues()

  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getVideos_3hours_mine').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger3_mine() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getVideos_7days_mine"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("自アカウント動画")
  let datas = sheet.getRange(3,9,sheet.getLastRow()-2,1).getValues()

  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getVideos_7days_mine').timeBased().at(datas[i][0]).create();
    }
  }
}


function all_Triggers_mine(){
  delTrigger_mine()
  delTrigger2_mine()
  delTrigger3_mine()
}

