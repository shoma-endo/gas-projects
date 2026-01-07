function onOpen() {
  var subMenus = [];
  subMenus.push({
    name: "認証",
    functionName: "test" 
  });
  subMenus.push({
    name: "初期設定 ※1回しか押さない",
    functionName: "setTrigger" 
  });
  subMenus.push({
    name: "トリガー削除 ※メンテナンス時のみ選択",
    functionName: "clearTrigger" 
  });
  ss.addMenu("GAS起動", subMenus);
}

function test(){

    
}

function setTrigger(){

  ScriptApp.newTrigger('getAccount').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getFollowers').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getNewMedia').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('getTagList').timeBased().atHour(1).everyDays(1).create();
  //ScriptApp.newTrigger('posting_Tags').timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger('new_engagement_rate').timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger('getMediaCount_All').timeBased().atHour(3).everyDays(1).create();
  ScriptApp.newTrigger('getMediaCount_All_Reel').timeBased().atHour(4).everyDays(1).create();
  getpastMedia()
  getNewMedia()
  getTagList()
  new_engagement_rate()
  getDate_first()

  const time = new Date();
  time.setMinutes(time.getMinutes()+3)
  ScriptApp.newTrigger('getMediaCount_All').timeBased().at(time).create();

  const time2 = new Date();
  time2.setMinutes(time2.getMinutes()+3)
  ScriptApp.newTrigger('getMediaCount_All_Reel').timeBased().at(time2).create();

  Browser.msgBox("初期設定が完了しました!!")
}

function clearTrigger(){

  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "getAccount"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getFollowers"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getNewMedia"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getTagList"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "posting_Tags"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "new_engagement_rate"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getCompete"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getHashTagID"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getMediaCount_All"){
      ScriptApp.deleteTrigger(trigger);
    }else if(trigger.getHandlerFunction() == "getMediaCount_All_Reel"){
      ScriptApp.deleteTrigger(trigger);
    }

  }

}

function resetTrigger(){

  ScriptApp.newTrigger('getAccount').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getFollowers').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getNewMedia').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('getTagList').timeBased().atHour(1).everyDays(1).create();
  //ScriptApp.newTrigger('posting_Tags').timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger('new_engagement_rate').timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger('getMediaCount_All').timeBased().atHour(3).everyDays(1).create();
  ScriptApp.newTrigger('getMediaCount_All_Reel').timeBased().atHour(4).everyDays(1).create();

  
}

function getpastMedia(){

  const sheet = ss.getSheetByName("過去投稿");

  //投稿検索
  const fields_media = "media{media_product_type,caption,id,like_count,comments_count,permalink,timestamp}"
  const url_media = "https://graph.facebook.com/v12.0/" + business_id + "?fields=" + fields_media + "&access_token=" + token
  let encodedURI_media = encodeURI(url_media);
  let response_media = UrlFetchApp.fetch(encodedURI_media,options).getContentText();
  let json_media = JSON.parse(response_media)

  let arr = [];
  let date = ""
  let like_count = ""
  let comments_count = ""
  let ts = ""
  let id = ""
  let caption = ""
  let next_url = "-"
  let url = ""

  for (var num in json_media["media"]["data"]) {
    if(json_media["media"]["data"][num]["media_product_type"] == "REELS"){
      type = "リール"
    }else if(json_media["media"]["data"][num]["media_product_type"] == "FEED"){
      type = "投稿"
    }
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
      arr.push([today,id,type,caption,date,url,like_count,comments_count])
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
      if(json_media["data"][num]["media_product_type"] == "REELS"){
        type = "リール"
      }else if(json_media["data"][num]["media_product_type"] == "FEED"){
        type = "投稿"
      }
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
        arr.push([today,id,type,caption,date,url,like_count,comments_count])
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

  arr.sort(sorting_asc_past)

  //投稿インサイト検索
  let media_id = ""
  const f_metric_media = "engagement,impressions,reach,saved"
  const r_metric_media = "total_interactions,plays,reach,saved,shares"
  let url_mediainsight = ""
  let encodedURI_mediainsight = ""
  let response_mediainsight = ""
  let json_mediainsight = ""

  let engagement = ""
  let impressions = ""
  let reach = ""
  let saved = ""

  for(let i=0;i<arr.length;i++){
    media_id = arr[i][1]
    if(arr[i][2] == "投稿"){
      url_mediainsight = "https://graph.facebook.com/v12.0/" + media_id + "/insights?metric=" + f_metric_media + "&access_token=" + token
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
      arr[i].push(saved,"-",engagement,reach,impressions,"-")
    }else if(arr[i][2] == "リール"){
      url_mediainsight = "https://graph.facebook.com/v12.0/" + media_id + "/insights?metric=" + r_metric_media + "&access_token=" + token
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
      arr[i].push(saved,share,total_interactions,reach,"-",plays)
    }
  }
  
  if(arr.length > 0){
    //getComment(arr)

    sheet.getRange(2,1,arr.length,arr[0].length).setValues(arr)

    //sheet.getRange(200,1,arr.length,arr[0].length).setValues(arr)
    SpreadsheetApp.flush()

    const requests = {updateDimensionProperties: {
      properties: {pixelSize: 40},
      range: {sheetId: sheet.getSheetId(), startIndex:1, endIndex: sheet.getLastRow(), dimension: "ROWS"},
      fields: "pixelSize"
    }};
    Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());

    //getpastTagList()
    //postCount()
    getTopten()
  }

/*
  ScriptApp.newTrigger('getAccount').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getFollowers').timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('getNewMedia').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('getTagList').timeBased().atHour(1).everyDays(1).create();
  //ScriptApp.newTrigger('getCompete').timeBased().atHour(0).everyDays(1).create();
  //ScriptApp.newTrigger('getHashTagID').timeBased().atHour(0).everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).create();
  ScriptApp.newTrigger('posting_Tags').timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger('new_engagement_rate').timeBased().atHour(2).everyDays(1).create();
  //ScriptApp.newTrigger('get_posting_Tags').timeBased().atHour(4).everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SATURDAY).create();
  
*/

}

/*
function getpastTagList(){ //過去投稿一覧からタグ一覧を作成する

  const sheet1 = ss.getSheetByName("過去投稿");
  const sheet2 = ss.getSheetByName("過去投稿タグ");
  
  const datas = sheet1.getRange(2,3,sheet1.getLastRow()-1,10).getValues();
  //Logger.log(datas[5])

  let content = ""
  let content_arr = ""
  let tag = ""
  let arr = []
  let flg = false
  let tag_arr = "";
  for(let i=0;i<datas.length;i++){　//キャプション内タグ
    content = datas[i][0]
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
              tag = tag.replace(" ","")
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
    content = datas[i][9]
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
              tag = tag.replace(" ","")
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
  sheet2.getRange(3,1,arr.length,1).setValues(arr)
  sheet2.getRange(3,16,arr.length,1).setValues(arr)
  sheet2.getRange(3,31,arr.length,1).setValues(arr)
  sheet2.getRange(3,46,arr.length,1).setValues(arr)
  sheet2.getRange(3,61,arr.length,1).setValues(arr)

}
*/
/*
function postCount(){
  
  const sheet1 = ss.getSheetByName("過去投稿");
  const sheet2 = ss.getSheetByName("過去投稿タグ");
  
  const datas = sheet1.getRange(2,3,sheet1.getLastRow()-1,10).getValues();
  const datas2 = sheet2.getRange(2,1,sheet2.getLastRow()-1,sheet2.getLastColumn()).getValues();

  let month = ""
  let next_month = ""
  let tag = ""
  let caption = ""
  let comment = ""
  let ts = ""
  let like = ""
  let reach = ""
  let impressions = ""
  //let cnt = ""
  for(let i=1;i<datas2.length;i++){
    if(datas2[i][0] !== ""){
      tag = "#" + datas2[i][0]
      for(let j=0;j<datas.length;j++){
        caption = datas[j][0]
        comment = datas[j][9]
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          like = datas[j][3]
          reach = datas[j][7]
          impressions = datas[j][8]
          ts = datas[j][1]
          for(let k=1;k<datas2[0].length;k++){ 
            month = datas2[0][k]
            next_month = datas2[0][k+1]
            if(typeof(month)=="object"){
              if(ts.getTime() >= month.getTime()){
                if(typeof(next_month)=="object"){
                  if(ts.getTime() < next_month.getTime()){
                    if(datas2[i][k] == ""){
                      datas2[i][k] = 1
                      datas2[i][k+15] = Number(like)
                      datas2[i][k+30] = Number(reach)
                      datas2[i][k+45] = Number(impressions)
                      break;
                    }else{
                      datas2[i][k] += 1
                      datas2[i][k+15] += Number(like)
                      datas2[i][k+30] += Number(reach)
                      datas2[i][k+45] += Number(impressions)
                      break;
                    }
                  }
                }else{
                  if(datas2[i][k] == ""){
                      datas2[i][k] = 1
                      datas2[i][k+15] = Number(like)
                      datas2[i][k+30] = Number(reach)
                      datas2[i][k+45] = Number(impressions)
                    break;
                  }else{
                      datas2[i][k] += 1
                      datas2[i][k+15] += Number(like)
                      datas2[i][k+30] += Number(reach)
                      datas2[i][k+45] += Number(impressions)
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

  datas2.shift();
  for(let i=0;i<datas2.length;i++){
    for(let j=1;j<14;j++){
      if(datas2[i][14] == ""){
        datas2[i][14] = 0
      }
      if(datas2[i][29] == ""){
        datas2[i][29] = 0
      }
      if(datas2[i][44] == ""){
        datas2[i][44] = 0
      }
      if(datas2[i][59] == ""){
        datas2[i][59] = 0
      }
      datas2[i][14] += Number(datas2[i][j])
      datas2[i][29] += Number(datas2[i][j+15])
      datas2[i][44] += Number(datas2[i][j+30])
      datas2[i][59] += Number(datas2[i][j+45])
    }
  }
  datas2.sort(sorting_desc_past)
  //Logger.log(datas2)
  sheet2.getRange(3,1,datas2.length,datas2[0].length).setValues(datas2)
}
*/

function sorting_desc_past(a, b){
  if(a[14] > b[14]){
    return -1;
  }else if(a[14] < b[14] ){
    return 1;
  }else{
   return 0;
  }
}

function getTopten(){

  const sheet = ss.getSheetByName("過去投稿");
  const sheet_topten = ss.getSheetByName("過去投稿トップテン")

  let cnt = 0

  //トップテン（投稿）
  let datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_like)

  if(datas_like_comment.length < 10){
    cnt = datas_like_comment.length
  }else{
    cnt = 10
  }

  let arr_like = [];
  let new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "投稿"){
      if(new_cnt < cnt){
        arr_like.push([datas_like_comment[i][4],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_like.length > 0){
    sheet_topten.getRange(3,2,new_cnt,3).setValues(arr_like)
  }

  datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_save)
  let arr_save = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "投稿"){
      if(new_cnt < cnt){
        arr_save.push([datas_like_comment[i][6],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_save.length > 0){
    sheet_topten.getRange(3,5,new_cnt,3).setValues(arr_save)
  }

  datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_reach)
  let arr_reach = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "投稿"){
      if(new_cnt < cnt){
        arr_reach.push([datas_like_comment[i][9],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_reach.length > 0){
    sheet_topten.getRange(3,8,new_cnt,3).setValues(arr_reach)
  }
  

/*
  const requests = {updateDimensionProperties: {
      properties: {pixelSize: 60},
      range: {sheetId: sheet_topten.getSheetId(), startIndex:2, endIndex: sheet_topten.getLastRow(), dimension: "ROWS"},
      fields: "pixelSize"
    }};
    Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());
*/
  //トップテンに色をつける
  sheet_topten.getRange(3,2,10,3).setBackground(null)
  sheet_topten.getRange(3,5,10,3).setBackground(null)
  sheet_topten.getRange(3,8,10,3).setBackground(null)
  for(let i=0;i<arr_like.length;i++){
    if(arr_like[i][1] !== ""){
      for(let j=0;j<arr_save.length;j++){
        if(arr_save[j][1] !== ""){
          if(arr_like[i][1].getTime() == arr_save[j][1].getTime()){
            for(let k=0;k<arr_reach.length;k++){
              if(arr_reach[k][1] !== ""){
                if(arr_save[j][1].getTime() == arr_reach[k][1].getTime()){
                  sheet_topten.getRange(3+i,2,1,3).setBackground("#e2ce90")
                  sheet_topten.getRange(3+j,5,1,3).setBackground("#e2ce90")
                  sheet_topten.getRange(3+k,8,1,3).setBackground("#e2ce90")
                }
              }
            }
          }
        }
      }
    }
  }

  getTopten_Reel()
}

function getTopten_Reel(){

  const sheet = ss.getSheetByName("過去投稿");
  const sheet_topten = ss.getSheetByName("過去投稿トップテン")

  let cnt = 0

  //トップテン（リール）
  let datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_like)

  if(datas_like_comment.length < 10){
    cnt = datas_like_comment.length
  }else{
    cnt = 10
  }

  let arr_like = [];
  let new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "リール"){
      if(new_cnt < cnt){
        arr_like.push([datas_like_comment[i][4],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_like.length > 0){
    sheet_topten.getRange(18,2,new_cnt,3).setValues(arr_like)
  }

  datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_save)
  let arr_save = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "リール"){
      if(new_cnt < cnt){
        arr_save.push([datas_like_comment[i][6],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_save.length > 0){
    sheet_topten.getRange(18,5,new_cnt,3).setValues(arr_save)
  }

  datas_like_comment = sheet.getRange(2,3,sheet.getLastRow()-1,10).getValues()
  datas_like_comment.sort(sorting_desc_reach)
  let arr_reach = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(datas_like_comment[i][0] == "リール"){
      if(new_cnt < cnt){
        arr_reach.push([datas_like_comment[i][9],datas_like_comment[i][2],datas_like_comment[i][3]])
        new_cnt++
      }else{
        break;
      }
    }
  }
  if(arr_reach.length > 0){
    sheet_topten.getRange(18,8,new_cnt,3).setValues(arr_reach)
  }
  

/*
  const requests = {updateDimensionProperties: {
      properties: {pixelSize: 60},
      range: {sheetId: sheet_topten.getSheetId(), startIndex:2, endIndex: sheet_topten.getLastRow(), dimension: "ROWS"},
      fields: "pixelSize"
    }};
    Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());
*/
  //トップテンに色をつける
  sheet_topten.getRange(18,2,10,3).setBackground(null)
  sheet_topten.getRange(18,5,10,3).setBackground(null)
  sheet_topten.getRange(18,8,10,3).setBackground(null)
  for(let i=0;i<arr_like.length;i++){
    if(arr_like[i][1] !== ""){
      for(let j=0;j<arr_save.length;j++){
        if(arr_save[j][1] !== ""){
          if(arr_like[i][1].getTime() == arr_save[j][1].getTime()){
            for(let k=0;k<arr_reach.length;k++){
              if(arr_reach[k][1] !== ""){
                if(arr_save[j][1].getTime() == arr_reach[k][1].getTime()){
                  sheet_topten.getRange(18+i,2,1,3).setBackground("#e2ce90")
                  sheet_topten.getRange(18+j,5,1,3).setBackground("#e2ce90")
                  sheet_topten.getRange(18+k,8,1,3).setBackground("#e2ce90")
                }
              }
            }
          }
        }
      }
    }
  }


}

function sorting_desc_like(a, b){
  if(a[4] > b[4]){
    return -1;
  }else if(a[4] < b[4] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_save(a, b){
  if(a[6] > b[6]){
    return -1;
  }else if(a[6] < b[6] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_reach(a, b){
  if(a[9] > b[9]){
    return -1;
  }else if(a[9] < b[9] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_asc_past(a, b){
  if(a[4] < b[4]){
    return -1;
  }else if(a[4] > b[4] ){
    return 1;
  }else{
   return 0;
  }
}