function getHashTagID(){

  //ハッシュタグ検索(自分の投稿)
  const mytag_sheet = ss.getSheetByName("投稿タグ一覧(投稿数)");
  const tag_lastRow = mytag_sheet.getRange(1,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const tagid_lastRow = mytag_sheet.getRange(1,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  if(tag_lastRow-tagid_lastRow > 0){
    let mytag_data = mytag_sheet.getRange(tagid_lastRow+1,1,tag_lastRow-tagid_lastRow,1).getValues();
    let tag_name = ""
    let tag_id = ""
    for(let i=0;i<mytag_data.length;i++){
      tag_name = mytag_data[i][0]
      const url = "https://graph.facebook.com/v12.0/ig_hashtag_search?user_id=" + business_id + "&q=" + tag_name + "&access_token=" + token
      let encodedURI = encodeURI(url);
      let response = UrlFetchApp.fetch(encodedURI,options).getContentText();
      let json = JSON.parse(response)
      try{ 
        tag_id = String(json["data"][0]["id"])
        mytag_data[i].push(tag_id)
      }catch{
        Logger.log("エラー:"+tag_name)
        mytag_data[i].push("-")
      }
    }
    mytag_sheet.getRange(tagid_lastRow+1,1,mytag_data.length,2).setValues(mytag_data) //ID一覧作成
  }

}

function searchHashTag_Top(){

  const sheet = ss.getSheetByName("タグ検索(トップ)");
  let hashtag_id = sheet.getRange(2,3).getValue();

  if(sheet.getRange(5,1).getValue()==""){
    sheet.getRange(5,1).setValue("-")
  }
  sheet.getRange(5,1,sheet.getLastRow(),2).clear();

  const fields = "id,media_type,comments_count,like_count,media_url,caption"
  const url = "https://graph.facebook.com/v12.0/" + hashtag_id + "/top_media?user_id=" + business_id + "&fields=" + fields + "&access_token=" + token
  let encodedURI = encodeURI(url);
  let response = UrlFetchApp.fetch(encodedURI,options).getContentText();
  let json = JSON.parse(response)

  let like = ""
  let caption = ""

  let arr = []
  //Logger.log(json)
  for (var num in json["data"]) {
    //like = ""
    like = String(json["data"][num]["like_count"])
    caption = json["data"][num]["caption"]
    if(like !== "undefined"){
      arr.push([caption,like])
    }
  }
  arr.sort(sorting_desc)
  sheet.getRange(5,1,arr.length,arr[0].length).setValues(arr)
  SpreadsheetApp.flush()

  const requests = {updateDimensionProperties: {
    properties: {pixelSize: 80},
    range: {sheetId: sheet.getSheetId(), startIndex:1, endIndex: sheet.getLastRow(), dimension: "ROWS"},
    fields: "pixelSize"
  }};
  Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());

}

function searchHashTag_Recent(){

  const sheet = ss.getSheetByName("タグ検索(最近)");
  let hashtag_id = sheet.getRange(2,3).getValue();

  if(sheet.getRange(5,1).getValue()==""){
    sheet.getRange(5,1).setValue("-")
  }
  sheet.getRange(5,1,sheet.getLastRow(),2).clear();

  const fields = "id,media_type,comments_count,like_count,media_url,caption"
  const url = "https://graph.facebook.com/v12.0/" + hashtag_id + "/recent_media?user_id=" + business_id + "&fields=" + fields + "&access_token=" + token
  let encodedURI = encodeURI(url);
  let response = UrlFetchApp.fetch(encodedURI,options).getContentText();
  let json = JSON.parse(response)

  let like = ""
  let caption = ""

  let arr = []
  for (var num in json["data"]) {
    like = String(json["data"][num]["like_count"])
    caption = json["data"][num]["caption"]
    if(like !== "undefined"){
      arr.push([caption,like])
    }
  }
  arr.sort(sorting_desc)
  sheet.getRange(5,1,arr.length,arr[0].length).setValues(arr)
  SpreadsheetApp.flush()

  const requests = {updateDimensionProperties: {
    properties: {pixelSize: 80},
    range: {sheetId: sheet.getSheetId(), startIndex:1, endIndex: sheet.getLastRow(), dimension: "ROWS"},
    fields: "pixelSize"
  }};
  Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());

}


function posting_Tags(){ //投稿数を月次で転記する

  const sheet1 = ss.getSheetByName("投稿タグ一覧(投稿数)")
  const datas1 = sheet1.getRange(2,1,sheet1.getLastRow()-1,2).getValues()
  const sheet2 = ss.getSheetByName("投稿タグ推移")
  const datas2 = sheet2.getRange(1,1,sheet2.getLastRow(),sheet2.getLastColumn()).getDisplayValues()

  const month = now.getFullYear() + "年" + (now.getMonth()+1) + "月"
  let num = ""
  let flg = false
  let arr = []
  let col_num = ""

  for(let i=0;i<datas1.length;i++){
    flg = false
    for(let j=0;j<datas2.length;j++){
      if(datas1[i][0] == datas2[j][0]){
        flg = true
        num = j
        break;
      }
    }
    if(flg == true){ //過去にあるタグだったら
      for(let k=0;k<datas2[0].length;k++){
        if(datas2[0][k] == month){
          datas2[num][k] = datas1[i][1]
          col_num = k
          //num = k
          break;
        }
      }
    }else{ //なかったら
      arr = []
      for(let k=0;k<datas2[0].length;k++){
        if(datas2[0][k] == month){
          num = k;
          col_num = k
          break;
        }
      }
      arr.push(datas1[i][0])
      for(let l=0;l<num-1;l++){
        arr.push("")
      }
      arr.push(datas1[i][1])
      datas2.push(arr)
      for(let l=datas2[datas2.length-1].length;l<datas2[0].length;l++){
        datas2[datas2.length-1].push("")
      }
    }   
  }
  datas2.shift();
  datas2.sort((a, b) => {
    if(a[col_num] > b[col_num]){
      return -1;
    }else if(a[col_num] < b[col_num] ){
      return 1;
    }else{
    return 0;
    }
  });

  sheet2.getRange(2,1,datas2.length,datas2[0].length).setValues(datas2)

}

function sorting_desc_posting(a, b){
  if(a[1] > b[1]){
    return -1;
  }else if(a[1] < b[1] ){
    return 1;
  }else{
   return 0;
  }
}