function getNewMedia(){　//新規投稿

  const sheet = ss.getSheetByName("投稿");
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

  //投稿検索
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
  //let like_count = ""
  let ts = ""
  let id = ""
  let caption = ""
  let url = ""
  for (var num in json_media["media"]["data"]) {
    //Logger.log(json_media["media"]["data"][num]["media_url"])
    //Logger.log(json_media["media"]["data"][num]["thumbnail_url"])
    //Logger.log(json_media["media"]["data"][num]["permalink"])
    type = json_media["media"]["data"][num]["media_product_type"]
    if(type == "FEED"){
      id = json_media["media"]["data"][num]["id"]
      caption = json_media["media"]["data"][num]["caption"]//.substr(0,40)
      like_count = json_media["media"]["data"][num]["like_count"]
      ts = json_media["media"]["data"][num]["timestamp"]
      url = json_media["media"]["data"][num]["permalink"]
      date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])
      date_1hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+1,date.getMinutes())
      date_3hours = new Date(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours()+3,date.getMinutes())
      date_7days = new Date(date.getFullYear(),date.getMonth(),date.getDate()+7,date.getHours()+3,date.getMinutes())
      //if(date.getTime() > one_year_ago.getTime()){
      arr.push([today,id,date,caption,url,date_1hours,date_3hours,date_7days])
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
        ScriptApp.newTrigger('getNewMedia_follow_1hours').timeBased().at(arr[i][5]).create();
      }
      if(arr[i][6].getTime() >= today.getTime()){
        ScriptApp.newTrigger('getNewMedia_follow_3hours').timeBased().at(arr[i][6]).create();
      }
      if(today.getTime() <= arr[i][7].getTime() && arr[i][7].getTime() <= days_2_later.getTime() ){
      //if(arr[i][7].getTime() >= today.getTime()){
        ScriptApp.newTrigger('getNewMedia_follow_7days').timeBased().at(arr[i][7]).create();
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

  getNewReel()

  getTopten_new()

}

function sorting_asc(a, b){
  if(a[2] < b[2]){
    return -1;
  }else if(a[2] > b[2] ){
    return 1;
  }else{
   return 0;
  }
}

function getNewMedia_follow_1hours(){

  const sheet = ss.getSheetByName("投稿");
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

  //投稿検索
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
    //投稿インサイト検索
    const metric_media = "engagement,impressions,reach,saved"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let engagement = ""
    let impressions = ""
    let reach = ""
    let saved = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "engagement":
          engagement = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "impressions":
          impressions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
      }
    }
    let rate = ""
    //rate = engagement_rate(date,engagement)

    arr[0].push(saved,engagement,rate,reach,impressions)
    sheet.getRange(row_num,17,1,arr[0].length).setValues(arr)
  }
  //delTrigger3()

}


function getNewMedia_follow_3hours(){

  const sheet = ss.getSheetByName("投稿");
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

  //投稿検索
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
    //投稿インサイト検索
    const metric_media = "engagement,impressions,reach,saved"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let engagement = ""
    let impressions = ""
    let reach = ""
    let saved = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "engagement":
          engagement = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "impressions":
          impressions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
      }
    }
    let rate = ""
    //rate = engagement_rate(date,engagement)

    arr[0].push(saved,engagement,rate,reach,impressions)
    sheet.getRange(row_num,24,1,arr[0].length).setValues(arr)
  }
  //delTrigger()

}

function getNewMedia_follow_7days(){

  const sheet = ss.getSheetByName("投稿");
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

  //投稿検索
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
    //投稿インサイト検索
    const metric_media = "engagement,impressions,reach,saved"
    let url_mediainsight = ""
    let encodedURI_mediainsight = ""
    let response_mediainsight = ""
    let json_mediainsight = ""

    let engagement = ""
    let impressions = ""
    let reach = ""
    let saved = ""

    url_mediainsight = "https://graph.facebook.com/v12.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
    encodedURI_mediainsight = encodeURI(url_mediainsight);
    response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
    json_mediainsight = JSON.parse(response_mediainsight)
    
    for (var num in json_mediainsight["data"]) {
      name = json_mediainsight["data"][num]["name"];
      switch (name) {
        case "engagement":
          engagement = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "impressions":
          impressions = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "reach":
          reach = json_mediainsight["data"][num]["values"][0]["value"];
          break
        case "saved":
          saved = json_mediainsight["data"][num]["values"][0]["value"];
          break
      }
    }

    let rate = ""
    //rate = engagement_rate(date,engagement)
    rate = diffusion_followers(reach)
    arr[0].push(saved,engagement,"",reach,impressions,rate)
    sheet.getRange(row_num,31,1,arr[0].length).setValues(arr)
  }
  //delTrigger2()
}


function delTrigger() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_3hours"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("投稿");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,7,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() >= now.getTime()){
      ScriptApp.newTrigger('getNewMedia_follow_3hours').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger2() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_7days"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("投稿");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,8,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    //if(datas[i][0].getTime() >= now.getTime()){
    if(now.getTime() <= datas[i][0].getTime() && datas[i][0].getTime() <= days_2_later.getTime() ){ //多くなりすぎないように絞る
      ScriptApp.newTrigger('getNewMedia_follow_7days').timeBased().at(datas[i][0]).create();
    }
  }
}

function delTrigger3() {

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getNewMedia_follow_1hours"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const sheet = ss.getSheetByName("投稿");
  let lastRow = sheet.getRange(2,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = ""
  if(sheet.getRange(4,1).getValue() !== ""){
    datas = sheet.getRange(4,6,lastRow-3,1).getValues()
  }
  for(let i=0;i<datas.length;i++){
    if(datas[i][0] !== ""){
      if(datas[i][0].getTime() >= now.getTime()){
        ScriptApp.newTrigger('getNewMedia_follow_1hours').timeBased().at(datas[i][0]).create();
      }
    }
  }
}

function getTagList(){ //投稿一覧からタグ一覧を作成する

  const sheet1 = ss.getSheetByName("投稿");
  const datas = sheet1.getRange(4,3,sheet1.getLastRow()-2,7).getDisplayValues();
  
  const sheet2 = ss.getSheetByName("投稿タグ一覧(投稿数)");
  let datas2 = ""
  if(sheet2.getRange(2,1).getValue() !== ""){
    datas2 = sheet2.getRange(2,1,sheet2.getLastRow()-1,1).getValues();
  }
  
  const sheet3 = ss.getSheetByName("投稿タグ一覧(いいね数)");
  /*if(sheet3.getRange(2,1).getValue() == ""){
    sheet3.getRange(2,1).setValue("-")
  }
  sheet3.getRange(2,1,sheet3.getLastRow()-1,1).clear();*/

  const sheet4 = ss.getSheetByName("投稿タグ一覧(リーチ)");
  /*if(sheet4.getRange(2,1).getValue() == ""){
    sheet4.getRange(2,1).setValue("-")
  }
  sheet4.getRange(2,1,sheet4.getLastRow()-1,1).clear();*/

  const sheet5 = ss.getSheetByName("投稿タグ一覧(インプレッション)");
  /*if(sheet5.getRange(2,1).getValue() == ""){
    sheet5.getRange(2,1).setValue("-")
  }
  sheet5.getRange(2,1,sheet5.getLastRow()-1,1).clear();*/
  
  

  let content = ""
  let content_arr = ""
  let tag = ""
  let arr = []
  let flg = false
  let tag_arr = "";
  for(let i=0;i<datas.length;i++){　//キャプション内タグ
    content = datas[i][1]
    content_arr = ""
    tag_arr = ""
    if(content.match(/#/)){
      content_arr = content.split("\n")
      tag = ""
      for(let j=0;j<content_arr.length;j++){
        tag = content_arr[j]
        if(tag.match(/#/)){
          tag_arr = tag.split("#")
          for(let l=0;l<tag_arr.length;l++){
            tag = tag_arr[l]
            if(tag !== ""){
              tag = tag.split(" ")[0]
              tag = tag.replace(/ /g,"").replace(/　/g,"")
              if(arr.length > 0){
                flg = false
                for(let k=0;k<arr.length;k++){
                  if(tag == arr[k][0]){
                    flg = true
                    break;
                  }
                }
                if(flg == false){
                  arr.push([tag])
                }
              }else{
                arr.push([tag])
              }
            }
          }
        }
      }
    }
  }
  for(let i=0;i<datas.length;i++){　//コメント内タグ
    content = datas[i][6]
    content_arr = ""
    if(content.match(/#/)){
      content_arr = content.split("\n")
      tag = ""
      for(let j=0;j<content_arr.length;j++){
        tag = content_arr[j]
        if(tag.match(/#/)){
          tag_arr = tag.split("#")
          for(let l=0;l<tag_arr.length;l++){
            tag = tag_arr[l]
            if(tag !== ""){
              tag = tag.replace(/ /g,"").replace(/　/g,"")
              if(arr.length > 0){
                flg = false
                for(let k=0;k<arr.length;k++){
                  if(tag == arr[k][0]){
                    flg = true
                    break;
                  }
                }
                if(flg == false){
                  arr.push([tag])
                }
              }else{
                arr.push([tag])
              }
            }
          }
        }
      }
    }
  }

  if(arr.length > 0){
    if(sheet2.getRange(2,1).getValue() == ""){
      sheet2.getRange(2,1,arr.length,1).setValues(arr)
      sheet3.getRange(2,1,arr.length,1).setValues(arr)
      sheet4.getRange(2,1,arr.length,1).setValues(arr)
      sheet5.getRange(2,1,arr.length,1).setValues(arr)
    }else{
      const datas2 = sheet2.getRange(2,1,sheet2.getLastRow()-1,1).getValues();
      let exist_flg = false
      let new_arr = []
      for(let i=0;i<arr.length;i++){
        exist_flg = false
        for(let j=0;j<datas2.length;j++){
          if(arr[i][0] == datas2[j][0]){
            exist_flg = true
            break;
          }
        }
        if(exist_flg == false){
          new_arr.push([arr[i][0]])
        }
      }
      if(new_arr.length > 0){
        sheet2.getRange(sheet2.getLastRow()+1,1,new_arr.length,1).setValues(new_arr) 
        sheet3.getRange(sheet3.getLastRow()+1,1,new_arr.length,1).setValues(new_arr) 
        sheet4.getRange(sheet4.getLastRow()+1,1,new_arr.length,1).setValues(new_arr) 
        sheet5.getRange(sheet5.getLastRow()+1,1,new_arr.length,1).setValues(new_arr) 
      }
    }
  

  //postNewCount1()
  //postNewCount2()
  postNewCount3()
  //postNewCount4()
  //get_posting_Tags()
  }
  
}

/*
function postNewCount1(){
  
  const sheet1 = ss.getSheetByName("投稿");
  const sheet2 = ss.getSheetByName("投稿タグ一覧(投稿数)");
  
  const datas = sheet1.getRange(4,3,sheet1.getLastRow()-1,15).getValues();
  const datas2 = sheet2.getRange(1,1,sheet2.getLastRow(),1).getValues();
  if(sheet2.getRange(2,3).getValue()==""){
    sheet2.getRange(2,3).setValue("-")
  }
  sheet2.getRange(2,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).clear()
  const datas2_2 = sheet2.getRange(1,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).getValues();

  let month = ""
  let next_month = ""
  let tag = ""
  let caption = ""
  let comment = ""
  let ts = ""

  for(let i=1;i<datas2.length;i++){
    if(datas2[i][0] !== ""){
      tag = "#" + datas2[i][0]
      for(let j=0;j<datas.length;j++){
        caption = datas[j][1]
        comment = datas[j][6]
        //Logger.log(caption)
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          ts = datas[j][0]
          if(ts !== ""){
            for(let k=0;k<datas2_2[0].length;k++){ 
              month = datas2_2[0][k]
              next_month = datas2_2[0][k+1]
              if(typeof(month)=="object"){
                if(ts.getTime() >= month.getTime()){
                  if(typeof(next_month)=="object"){
                    if(ts.getTime() < next_month.getTime()){
                      if(datas2_2[i][k] == ""){
                        datas2_2[i][k] = 1
                        break;
                      }else{
                        datas2_2[i][k] += 1
                        break;
                      }
                    }
                  }else{
                    if(datas2_2[i][k] == ""){
                      datas2_2[i][k] = 1
                      break;
                    }else{
                      datas2_2[i][k] += 1
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  datas2_2.shift();
  sheet2.getRange(2,3,datas2_2.length,datas2_2[0].length).setValues(datas2_2)
}
*/

/*
function postNewCount2(){
  
  const sheet1 = ss.getSheetByName("投稿");
  const sheet2 = ss.getSheetByName("投稿タグ一覧(いいね数)");
  
  const datas = sheet1.getRange(4,3,sheet1.getLastRow()-1,22).getValues();
  const datas2 = sheet2.getRange(1,1,sheet2.getLastRow(),1).getValues();
  if(sheet2.getRange(2,3).getValue()==""){
    sheet2.getRange(2,3).setValue("-")
  }
  sheet2.getRange(2,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).clear()
  const datas2_2 = sheet2.getRange(1,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).getValues();

  let month = ""
  let next_month = ""
  let tag = ""
  let caption = ""
  let comment = ""
  let ts = ""
  let like = ""
  for(let i=1;i<datas2.length;i++){
    if(datas2[i][0] !== ""){
      tag = "#" + datas2[i][0]
      for(let j=0;j<datas.length;j++){
        caption = datas[j][1]
        comment = datas[j][6]
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          like = datas[j][7]
          ts = datas[j][0]
          if(ts !== ""){
            for(let k=0;k<datas2_2[0].length;k++){ 
              month = datas2_2[0][k]
              next_month = datas2_2[0][k+1]
              if(typeof(month)=="object"){
                if(ts.getTime() >= month.getTime()){
                  if(typeof(next_month)=="object"){
                    if(ts.getTime() < next_month.getTime()){
                      if(datas2_2[i][k] == ""){
                        datas2_2[i][k] = Number(like)
                        break;
                      }else{
                        datas2_2[i][k] += Number(like)
                        break;
                      }
                    }
                  }else{
                    if(datas2_2[i][k] == ""){
                        datas2_2[i][k] = Number(like)
                      break;
                    }else{
                        datas2_2[i][k] += Number(like)
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  datas2_2.shift();
  sheet2.getRange(2,3,datas2_2.length,datas2_2[0].length).setValues(datas2_2)
}
*/

function postNewCount3(){
  
  const sheet1 = ss.getSheetByName("投稿");
  const sheet2 = ss.getSheetByName("投稿タグ一覧(リーチ)");
  
  const datas = sheet1.getRange(4,3,sheet1.getLastRow()-1,27).getValues();
  const datas2 = sheet2.getRange(1,1,sheet2.getLastRow(),1).getValues();
  if(sheet2.getRange(2,3).getValue()==""){
    sheet2.getRange(2,3).setValue("-")
  }
  sheet2.getRange(2,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).clear()
  const datas2_2 = sheet2.getRange(1,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).getValues();

  let month = ""
  let next_month = ""
  let tag = ""
  let caption = ""
  let comment = ""
  let ts = ""
  let reach = ""

  for(let i=1;i<datas2.length;i++){
    if(datas2[i][0] !== ""){
      tag = "#" + datas2[i][0]
      for(let j=0;j<datas.length;j++){
        caption = datas[j][1]
        comment = datas[j][6]
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          reach = datas[j][12]
          ts = datas[j][0]
          if(ts !== ""){
            for(let k=0;k<datas2_2[0].length;k++){ 
              month = datas2_2[0][k]
              next_month = datas2_2[0][k+1]
              if(typeof(month)=="object"){
                if(ts.getTime() >= month.getTime()){
                  if(typeof(next_month)=="object"){
                    if(ts.getTime() < next_month.getTime()){
                      if(datas2_2[i][k] == ""){
                        datas2_2[i][k] = Number(reach)
                        break;
                      }else{
                        datas2_2[i][k] += Number(reach)
                        break;
                      }
                    }
                  }else{
                    if(datas2_2[i][k] == ""){
                      datas2_2[i][k] = Number(reach)
                      break;
                    }else{
                      datas2_2[i][k] += Number(reach)
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  datas2_2.shift();
  sheet2.getRange(2,3,datas2_2.length,datas2_2[0].length).setValues(datas2_2)
}

/*
function postNewCount4(){
  
  const sheet1 = ss.getSheetByName("投稿");
  const sheet2 = ss.getSheetByName("投稿タグ一覧(インプレッション)");
  
  const datas = sheet1.getRange(4,3,sheet1.getLastRow()-1,28).getValues();
  const datas2 = sheet2.getRange(1,1,sheet2.getLastRow(),1).getValues();
  if(sheet2.getRange(2,3).getValue()==""){
    sheet2.getRange(2,3).setValue("-")
  }
  sheet2.getRange(2,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).clear()
  const datas2_2 = sheet2.getRange(1,3,sheet2.getLastRow(),sheet2.getLastColumn()-2).getValues();

  let month = ""
  let next_month = ""
  let tag = ""
  let caption = ""
  let comment = ""
  let ts = ""
  let impressions = ""
  for(let i=1;i<datas2.length;i++){
    if(datas2[i][0] !== ""){
      tag = "#" + datas2[i][0]
      for(let j=0;j<datas.length;j++){
        caption = datas[j][1]
        comment = datas[j][6]
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          impressions = datas[j][13]
          ts = datas[j][0]
          if(ts !== ""){
            for(let k=0;k<datas2_2[0].length;k++){ 
              month = datas2_2[0][k]
              next_month = datas2_2[0][k+1]
              if(typeof(month)=="object"){
                if(ts.getTime() >= month.getTime()){
                  if(typeof(next_month)=="object"){
                    if(ts.getTime() < next_month.getTime()){
                      if(datas2_2[i][k] == ""){
                        datas2_2[i][k] = Number(impressions)
                        break;
                      }else{
                        datas2_2[i][k] += Number(impressions)
                        break;
                      }
                    }
                  }else{
                    if(datas2_2[i][k] == ""){
                      datas2_2[i][k] = Number(impressions)
                      break;
                    }else{
                      datas2_2[i][k] += Number(impressions)
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  datas2_2.shift();
  sheet2.getRange(2,3,datas2_2.length,datas2_2[0].length).setValues(datas2_2)
}

function engagement_rate(date,engagement){

  const account_sheet = ss.getSheetByName("アカウント情報")
  const datas = account_sheet.getRange(2,1,account_sheet.getLastRow()-1,5).getValues();

  let rate = ""
  for(let i=0;i<datas.length;i++){
    if(datas[i][0].getTime() == date.getTime()){
      rate = engagement / datas[i][4]
      break;
    }
  }

  return rate
}
*/

function new_engagement_rate(){

  const sheet = ss.getSheetByName("投稿");
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
    if(datas[i][12] !== ""){
      engagement = datas[i][12]
      reach = datas[i][14]
      rate1 = engagement / reach
      sheet.getRange(i+4,14).setValue(rate1)
    }
    if(datas[i][20] == ""){
      if(datas[i][19] !== ""){
        engagement = datas[i][19]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate2 = engagement / follower
              sheet.getRange(i+4,21).setValue(rate2)
              break;
            }
          }
        }
      }
    }
    if(datas[i][27] == ""){
      if(datas[i][26] !== ""){
        engagement = datas[i][26]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate3 = engagement / follower
              sheet.getRange(i+4,28).setValue(rate3)
              break;
            }
          }
        }
      }
    }
    if(datas[i][34] == ""){
      if(datas[i][33] !== ""){
        engagement = datas[i][33]
        for(let j=0;j<datas2.length;j++){
          if(datas2[j][0].getTime() == datas[i][0].getTime()){
            if(datas2[j][4] !== ""){
              follower = datas2[j][4]
              rate4 = engagement / follower
              sheet.getRange(i+4,35).setValue(rate4)
              break;
            }
          }
        }
      }
    }
  }
  new_engagement_rate_Reel()
}



function getMediaCount_All(){

  const sheet = ss.getSheetByName("投稿");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,1,lastRow-3,sheet.getLastColumn()).getValues()
  let time = ""
  let search_id = ""
  let row_num = ""
  let date = ""
  
//いいねとコメントを検索
  const fields_media = "media{caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v21.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
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
    url = json_media["media"]["data"][num]["permalink"]
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
      url = json_media["data"][num]["permalink"]
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

//保存、エンゲージメント、リーチ、インプレッションを検索
  if(arr.length > 0){
    for(let i=0;i<arr.length;i++){
      search_id = arr[i][0];
      //投稿インサイト検索
      //const metric_media = "engagement,impressions,reach,saved"
      const metric_media = "total_interactions,reach,saved"
      let url_mediainsight = ""
      let encodedURI_mediainsight = ""
      let response_mediainsight = ""
      let json_mediainsight = ""

      let engagement = ""
      let impressions = ""
      let reach = ""
      let saved = ""

      url_mediainsight = "https://graph.facebook.com/v22.0/" + search_id + "/insights?metric=" + metric_media + "&access_token=" + token
      encodedURI_mediainsight = encodeURI(url_mediainsight);
      response_mediainsight = UrlFetchApp.fetch(encodedURI_mediainsight,options).getContentText();
      json_mediainsight = JSON.parse(response_mediainsight)
      
      for (var num in json_mediainsight["data"]) {
        name = json_mediainsight["data"][num]["name"];
        switch (name) {
          case "total_interactions":
            engagement = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "impressions":
            impressions = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "reach":
            reach = json_mediainsight["data"][num]["values"][0]["value"];
            break
          case "saved":
            saved = json_mediainsight["data"][num]["values"][0]["value"];
            break
        }
      }
      arr[i].push(saved,engagement,reach,impressions)
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
        break;
        //new_arr.push([arr[j][1],arr[j][2],arr[j][3],arr[j][4],"-",arr[j][5],arr[j][6]])
      }
    }
    if(exist_flg == true){
      new_arr.push([content1,content2,content3,content4,"",content5,content6])
    }else{
      new_arr.push(["","","","","","",""])
    }
  }
  //Logger.log(new_arr)

  sheet.getRange(4,10,new_arr.length,new_arr[0].length).setValues(new_arr)
  new_engagement_rate()

  //getMediaCount_All_Reel()
}

function getDate_first(){ //初回だけ上書きする処理
  const sheet = ss.getSheetByName("投稿");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,3,lastRow-3,sheet.getLastColumn()).getValues()

  let arr = [];
  let date = ""
  for(let i=0;i<datas.length;i++){
    date = ""
    if(datas[i][0] !== ""){
      date = Utilities.formatDate(datas[i][0],"JST","yyyy/MM/dd")
      arr.push([date])
    }
  }
  if(arr.length > 0){
    sheet.getRange(4,1,arr.length,1).setValues(arr)
  }
  getDate_first_reel()
  
}

function getDate_first_reel(){ //初回だけ上書きする処理
  const sheet = ss.getSheetByName("リール");
  let lastRow = sheet.getRange(4,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let datas = sheet.getRange(4,3,lastRow-3,sheet.getLastColumn()).getValues()

  let arr = [];
  let date = ""
  for(let i=0;i<datas.length;i++){
    date = ""
    if(datas[i][0] !== ""){
      date = Utilities.formatDate(datas[i][0],"JST","yyyy/MM/dd")
      arr.push([date])
    }
  }
  sheet.getRange(4,1,arr.length,1).setValues(arr)

}

