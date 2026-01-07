function getNewReel(){　//新規リール

  const sheet = ss.getSheetByName("リール");
  let datas = ""
  let lastRow = ""
  if(sheet.getRange(4,3).getValue() == ""){
    lastRow = 3
  }else{
    lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  }
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(3,1,lastRow-1,sheet.getLastColumn()-1).getValues()
  }

  //リール検索
  const fields_media = "media{media_product_type,caption,id,like_count,comments_count,media_url,thumbnail_url,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = []; //取得用
  let date = ""
  let date_1hours = ""
  let date_3hours = ""
  let date_7days = ""
  let date_1month = ""
  //let like_count = ""
  let ts = ""
  let id = ""
  let caption = ""
  let url = ""
  let type = ""
  for (var num in json_media["media"]["data"]) {
    type = json_media["media"]["data"][num]["media_product_type"]
    if(type == "REELS"){
      id = json_media["media"]["data"][num]["id"]
      caption = json_media["media"]["data"][num]["caption"]//.substr(0,40)
      like_count = json_media["media"]["data"][num]["like_count"]
      ts = json_media["media"]["data"][num]["timestamp"]
      url = json_media["media"]["data"][num]["permalink"]
      date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
      date_1hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+1,date.getMinutes())
      date_3hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+3,date.getMinutes())
      date_7days = new Date(date.getFullYear(),date.getMonth(),date.getDate()+7,date.getHours()+3,date.getMinutes())
      date_1month = new Date(date.getFullYear(),date.getMonth()+1,date.getDate(),date.getHours()+3,date.getMinutes())
      //if(date.getTime() > one_year_ago.getTime()){
      arr.push([today,id,date,caption,url,date_1hours,date_3hours,date_7days,date_1month])
      //}else{
        //break;
      //}
    }
  }

  arr.sort(sorting_asc)
  //getComment(arr)

  let id2 = ""
  let flg = false;
  let arr2 = []
  for(let i=0;i<arr.length;i++){
    id2 = arr[i][1];
    flg = false;
    for(let j=0;j<datas.length;j++){
      if(id2 == datas[j][1]){
        flg = true;
        break;
      }
    }
    if(flg == false){
      arr2.push(arr[i])
      if(arr[i][5].getTime() >= today.getTime()){
        ScriptApp.newTrigger('getNewMedia_follow_1hours_Reel').timeBased().at(arr[i][5]).create();
      }
      if(arr[i][6].getTime() >= today.getTime()){
        ScriptApp.newTrigger('getNewMedia_follow_3hours_Reel').timeBased().at(arr[i][6]).create();
      }
      if(today.getTime() <= arr[i][7].getTime() && arr[i][7].getTime() <= days_2_later.getTime() ){ //多くなりすぎないように絞る
        ScriptApp.newTrigger('getNewMedia_follow_7days_Reel').timeBased().at(arr[i][7]).create();
      }
      if(today.getTime() <= arr[i][8].getTime() && arr[i][8].getTime() <= days_2_later.getTime() ){ //多くなりすぎないように絞る
        ScriptApp.newTrigger('getNewMedia_follow_1month_Reel').timeBased().at(arr[i][8]).create();
      }
    }
  }

  if(arr2.length > 0){
    sheet.getRange(lastRow+1,1,arr2.length,arr2[0].length).setValues(arr2)
    SpreadsheetApp.flush()
  }

  const requests = {updateDimensionProperties: {
    properties: {pixelSize: 40},
    range: {sheetId: sheet.getSheetId(), startIndex:1, endIndex: sheet.getLastRow(), dimension: "ROWS"},
    fields: "pixelSize"
  }};
  Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());

}

function getNewMedia_follow_1hours_Reel(){

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()

  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  //ten_before.setDate(ten_before.getDate()-1)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)
  //ten_after.setDate(ten_after.getDate()+1)

  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  for(let i=0;i<datas.length;i++){
    time = datas[i][5]
    date = datas[i][0]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][1]
        row_num = i+4
        break;
      }
    }
  }

  //リール検索
  const fields_media = "media{media_product_type,caption,id,like_count,comments_count,media_url,thumbnail_url,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let like_count = ""
  let comments_count = ""
  let id = ""
  for (var num in json_media["media"]["data"]) {
    id = json_media["media"]["data"][num]["id"]
    like_count = json_media["media"]["data"][num]["like_count"]
    comments_count = json_media["media"]["data"][num]["comments_count"]
    if(id == search_id){
      arr.push([like_count,comments_count])
    }
  }

  if(arr.length > 0){
    //リールインサイト検索
    const metric_media = "total_interactions,plays,reach,saved,shares"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let total_interactions = ""
    let plays = ""
    let reach = ""
    let saved = ""
    let share = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "total_interactions":
          total_interactions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "plays":
          plays = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "shares":
            share = json_mediainsight["data"][num]["values"][0]["value"];
            break
      }
    }
    arr[0].push(saved,share,total_interactions,"",reach,plays)
    sheet.getRange(row_num,18,1,arr[0].length).setValues(arr)
  }
  //delTrigger3_Reel()

}


function getNewMedia_follow_3hours_Reel(){

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()

  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-10)
  //ten_before.setDate(ten_before.getDate()-1)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+10)
  //ten_after.setDate(ten_after.getDate()+1)

  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  for(let i=0;i<datas.length;i++){
    time = datas[i][6]
    date = datas[i][0]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][1]
        row_num = i+4
        break;
      }
    }
  }

  //リール検索
  const fields_media = "media{caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let like_count = ""
  let comments_count = ""
  let id = ""
  for (var num in json_media["media"]["data"]) {
    id = json_media["media"]["data"][num]["id"]
    like_count = json_media["media"]["data"][num]["like_count"]
    comments_count = json_media["media"]["data"][num]["comments_count"]
    if(id == search_id){
      arr.push([like_count,comments_count])
    }
  }

  if(arr.length > 0){
    //リールインサイト検索
    const metric_media = "total_interactions,plays,reach,saved,shares"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let total_interactions = ""
    let plays = ""
    let reach = ""
    let saved = ""
    let share = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "total_interactions":
          total_interactions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "plays":
          plays = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "shares":
            share = json_mediainsight["data"][num]["values"][0]["value"];
            break
      }
    }
    arr[0].push(saved,share,total_interactions,"",reach,plays)
    sheet.getRange(row_num,26,1,arr[0].length).setValues(arr)
  }
  //delTrigger_Reel()

}

function getNewMedia_follow_7days_Reel(){

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()

  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  for(let i=0;i<datas.length;i++){
    time = datas[i][7]
    date = datas[i][0]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][1]
        row_num = i+4
        break;
      }
    }
  }

  //リール検索
  const fields_media = "media{caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let like_count = ""
  let comments_count = ""
  let id = ""
  for (var num in json_media["media"]["data"]) {
    id = json_media["media"]["data"][num]["id"]
    like_count = json_media["media"]["data"][num]["like_count"]
    comments_count = json_media["media"]["data"][num]["comments_count"]
    if(id == search_id){
      arr.push([like_count,comments_count])
    }
  }

  if(arr.length > 0){
    //リールインサイト検索
    const metric_media = "total_interactions,plays,reach,saved,shares"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let total_interactions = ""
    let plays = ""
    let reach = ""
    let saved = ""
    let share = ""
    let rate = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "total_interactions":
          total_interactions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "plays":
          plays = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "shares":
            share = json_mediainsight["data"][num]["values"][0]["value"];
            break
      }
    }
    rate = diffusion_followers(reach)
    arr[0].push(saved,share,total_interactions,"",reach,plays,rate)
    sheet.getRange(row_num,34,1,arr[0].length).setValues(arr)
  }
  //delTrigger2_Reel()
}

function getNewMedia_follow_1month_Reel(){

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()

  let ten_before = new Date();
  ten_before.setMinutes(ten_before.getMinutes()-5)
  let ten_after = new Date();
  ten_after.setMinutes(ten_after.getMinutes()+5)

  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  for(let i=0;i<datas.length;i++){
    time = datas[i][8]
    date = datas[i][0]
    if(time !== ""){
      if(ten_before.getTime() <= time.getTime() && time.getTime() <=ten_after.getTime()){
        search_id = datas[i][1]
        row_num = i+4
        break;
      }
    }
  }

  //リール検索
  const fields_media = "media{caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let like_count = ""
  let comments_count = ""
  let id = ""
  for (var num in json_media["media"]["data"]) {
    id = json_media["media"]["data"][num]["id"]
    like_count = json_media["media"]["data"][num]["like_count"]
    comments_count = json_media["media"]["data"][num]["comments_count"]
    if(id == search_id){
      arr.push([like_count,comments_count])
    }
  }

  if(arr.length > 0){
    //リールインサイト検索
    const metric_media = "total_interactions,plays,reach,saved,shares"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let total_interactions = ""
    let plays = ""
    let reach = ""
    let saved = ""
    let share = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "total_interactions":
          total_interactions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "plays":
          plays = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "shares":
            share = json_mediainsight["data"][num]["values"][0]["value"];
            break
      }
    }
    arr[0].push(saved,share,total_interactions,"",reach,plays)
    sheet.getRange(row_num,43,1,arr[0].length).setValues(arr)
  }
  //delTrigger4_Reel()
}


function delTrigger_Reel() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_3hours_Reel"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,7,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getNewMedia_follow_3hours_Reel').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger2_Reel() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_7days_Reel"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,8,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(now.getTime() <= datas[i][0].getTime() && datas[i][0].getTime() <= days_2_later.getTime() ){ //多くなりすぎないように絞る
      ScriptApp.newTrigger('getNewMedia_follow_7days_Reel').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger3_Reel() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_1hours_Reel"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,6,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getNewMedia_follow_1hours_Reel').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger4_Reel() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_1month_Reel"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,9,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(now.getTime() <= datas[i][0].getTime() && datas[i][0].getTime() <= days_2_later.getTime() ){ //多くなりすぎないように絞る
      ScriptApp.newTrigger('getNewMedia_follow_1month_Reel').timeBased().at(datas[i][0]).create();
    }
  }
}


function getMediaCount_All_Reel(){

  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()
  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  
//いいねとコメントを検索
  const fields_media = "media{caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v22.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let like_count = ""
  let comments_count = ""
  let id = ""
  let next_url = ""

  for (var num in json_media["media"]["data"]) {
    next_url = ""
    id = String(json_media["media"]["data"][num]["id"])
    caption = json_media["media"]["data"][num]["caption"]//.substr(0,40)
    like_count = json_media["media"]["data"][num]["like_count"]
    comments_count = ""
    if(String(json_media["media"]["data"][num]["comments_count"]) !== "undefined"){
      comments_count = json_media["media"]["data"][num]["comments_count"]
    }
    ts = json_media["media"]["data"][num]["timestamp"]
    //url = json_media["media"]["data"][num]["permalink"]
    date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
    if(date.getTime() > one_year_ago.getTime()){
      arr.push([id,like_count,comments_count])
      if(num == 24){   
        next_url = json_media["media"]["paging"]["next"]
        response_media = UrlFetchApp.fetch(next_url,options).getContentText();
        json_media = JSON.parse(response_media)
      }
    }else{
      break;
    }
  }

  while(next_url !== ""){
    for (var num in json_media["data"]) {
      next_url = ""
      id = String(json_media["data"][num]["id"])
      caption = json_media["data"][num]["caption"]//.substr(0,40)
      like_count = json_media["data"][num]["like_count"]
      comments_count = ""
      if(String(json_media["data"][num]["comments_count"]) !== "undefined"){
        comments_count = json_media["data"][num]["comments_count"]
      }
      ts = json_media["data"][num]["timestamp"]
      //url = json_media["data"][num]["permalink"]
      date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
      if(date.getTime() > one_year_ago.getTime()){
        arr.push([id,like_count,comments_count])
        if(num == 24){
          next_url = json_media["paging"]["next"]
          response_media = UrlFetchApp.fetch(next_url,options).getContentText();
          json_media = JSON.parse(response_media)
        }
      }else{
        break;
      }
    }
  }

  //Logger.log(arr)

//保存、再生数、リーチ、インプレッションを検索
  if(arr.length > 0){
    for(let i=0;i<arr.length;i++){
      search_id = arr[i][0];
      //リールインサイト検索
      //const metric_media = "total_interactions,plays,reach,saved,shares"
      const metric_media = "total_interactions,reach,saved,shares"
      let url_mediainsight = ""
      let encodedURI_mediainsight = ""
      let response_mediainsight = ""
      let json_mediainsight = ""

      let total_interactions = ""
      let plays = ""
      let reach = ""
      let saved = ""
      let share = ""

      url_mediainsight = "https://graph.facebook.com/v22.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
      encodedURI_mediainsight = encodeURI(url_mediainsight);
      response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
      json_mediainsight = JSON.parse(response_mediainsight)
      
      for (var num in json_mediainsight["data"]) {
        name = json_mediainsight["data"][num]["name"];
        switch (name) {
          case "total_interactions":
            total_interactions = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "plays":
            plays = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "reach":
            reach = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "saved":
            saved = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "shares":
            share = json_mediainsight["data"][num]["values"][0]["value"];
            break
        }
      }
      arr[i].push(saved,share,total_interactions,reach,plays)
    }
  }

  //Logger.log(arr)
  let new_arr = [];
  let exist_flg = false;

  let content1 = ""
  let content2 = ""
  let content3 = ""
  let content4 = ""
  let content5 = ""
  let content6 = ""
  let content7 = ""

  for(let i=0;i<datas.length;i++){
    exist_flg = false;
    for(let j=0;j<arr.length;j++){
      if(datas[i][1] == arr[j][0]){
        exist_flg = true;
        content1 = arr[j][1]
        content2 = arr[j][2]
        content3 = arr[j][3]
        content4 = arr[j][4]
        content5 = arr[j][5]
        content6 = arr[j][6]
        content7 = arr[j][7]
        break;
        //new_arr.push([arr[j][1],arr[j][2],arr[j][3],arr[j][4],"-",arr[j][5],arr[j][6]])
      }
    }
    if(exist_flg == true){
      new_arr.push([content1,content2,content3,content4,content5,"",content6,content7])
    }else{
      new_arr.push(["","","","","","","",""])
    }
  }
  //Logger.log(new_arr)

  sheet.getRange(4,10,new_arr.length,new_arr[0].length).setValues(new_arr)
  new_engagement_rate_Reel()
}

function new_engagement_rate_Reel(){

  const sheet = ss.getSheetByName("リール");
  let datas = ""
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  if(sheet.getRange(4,3).getValue() !== ""){
    datas = sheet.getRange(4,1,lastRow-2,sheet.getLastColumn()).getValues()
  }

  const account_sheet = ss.getSheetByName("アカウント情報")
  const datas2 = account_sheet.getRange(2,1,account_sheet.getLastRow()-1,5).getValues();

  let rate1 = ""
  let rate2 = ""
  let rate3 = ""
  let rate4 = ""
  let engagement = ""
  let follower = "" 
  let reach = ""
  for(let i=0;i<datas.length;i++){
    if(datas[i][13] !== ""){
      engagement = datas[i][13]
      reach = datas[i][15]
      rate1 = engagement / reach
      sheet.getRange(i+4,15).setValue(rate1)
    }
    if(datas[i][22] == ""){
      if(datas[i][21] !== ""){
        //Logger.log(datas[i])
        engagement = datas[i][21]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate1 = engagement / follower
              sheet.getRange(i+4,23).setValue(rate1)
              break;
            }
          }
        }
      }
    }
    if(datas[i][30] == ""){
      if(datas[i][29] !== ""){
        engagement = datas[i][29]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate2 = engagement / follower
              sheet.getRange(i+4,31).setValue(rate2)
              break;
            }
          }
        }
      }
    }
    if(datas[i][38] == ""){
      if(datas[i][37] !== ""){
        engagement = datas[i][37]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate3 = engagement / follower
              sheet.getRange(i+4,39).setValue(rate3)
              break;
            }
          }
        }
      }
    }
    if(datas[i][47] == ""){
      if(datas[i][46] !== ""){
        engagement = datas[i][46]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate4 = engagement / follower
              sheet.getRange(i+4,48).setValue(rate4)
              break;
            }
          }
        }
      }
    }
  }
}

function diffusion_followers(reach){

  const account_sheet = ss.getSheetByName("アカウント情報")
  const lastRow = account_sheet.getRange(1,5).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const datas2 = account_sheet.getRange(2,5,lastRow-1,1).getValues();

  let follower = datas2[datas2.length-1][0]
  return reach / follower
}