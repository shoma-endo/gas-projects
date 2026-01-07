function getCompete() {
  //競合検索

  const sheet = ss.getSheetByName("競合");
  if(sheet.getLastRow() > 2){
      const datas = sheet.getRange(3,1,sheet.getLastRow()-2,1).getValues()
      const sheet2 = ss.getSheetByName("競合（原本）")
      const sheet3 = ss.getSheetByName("タグ一覧（原本）")

      //シートがあるか確認
      let sheets_cnt = ss.getNumSheets();
      let sheets = ss.getSheets()
      let arr_sheetname = [];
      let s_name = ""
      let sheet_num = ""
      for(let i=0;i<sheets_cnt;i++){
        s_name = sheets[i].getSheetName()
        arr_sheetname.push(s_name)
        if(s_name == "競合"){
          sheet_num = i+2
        }
      }

      let arr = [];
      let id = ""
      let like_count = ""
      let url = ""
      let comments_count = ""
      let display_name = ""
      let user_id = ""
      let userid_arr = [];
      let ts = ""
      let date = ""
      let tag = ""
      let json_compete = ""
      let sheet_exist_flg = false;
      let sheet_new = ""
      let sheet_new2 = ""
      let follower_count = ""
      let media_count = ""

      for(let i=0;i<datas.length;i++){
        display_name = datas[i][0]
        if(display_name !== ""){
          if(display_name !== "taimu.t"){
            sheet_exist_flg = false
            arr = [];

            for(let j=0;j<arr_sheetname.length;j++){
              if(display_name == arr_sheetname[j]){
                sheet_exist_flg = true
                break;
              }
            }
            if(sheet_exist_flg == false){
              sheet3.copyTo(ss).setName("タグ:"+display_name)
              sheet_new2 = ss.getSheetByName("タグ:"+display_name);
              sheet_new2.setTabColor("#155aae")
              sheet_new2.activate();
              ss.moveActiveSheet(sheet_num);

              sheet2.copyTo(ss).setName(display_name)
              sheet_new = ss.getSheetByName(display_name);
              sheet_new.setTabColor("#38761d")
              sheet_new.activate();
              ss.moveActiveSheet(sheet_num);
              sheet_new.getRange(10,2).setValue(one_year_ago)
              sheet_new.getRange(9,17).setValue(today)
            }else{
              sheet_new = ss.getSheetByName(display_name);
              sheet_new2 = ss.getSheetByName("タグ:"+display_name);
            }
            SpreadsheetApp.flush()
            let new_datas = sheet_new.getRange(10,2,sheet_new.getLastRow()-9,11).getValues()
            let last_date = sheet_new.getRange(9,24).getValue();
            if(last_date == ""){
              last_date = one_year_ago //便宜上当てはめ
            }

            if(last_date.getTime() !== today.getTime()){

              const fields_competition = "{followers_count,media_count,media{caption,comments_count,like_count,permalink,timestamp},id}"
              const url_compete = "https://graph.facebook.com/v12.0/" + business_id + "?fields=business_discovery.username(" + display_name + ")" + fields_competition + "&access_token=" + token
              let encodedURI_compete = encodeURI(url_compete);
              let response_compete = UrlFetchApp.fetch(encodedURI_compete,options).getContentText();
              json_compete = JSON.parse(response_compete)

              follower_count = ""
              media_count = ""

              try{

                follower_count = String(json_compete["business_discovery"]["followers_count"])
                media_count = String(json_compete["business_discovery"]["media_count"])

                user_id = json_compete["business_discovery"]["id"]
                datas[i].push(user_id)
                
                for (var num in json_compete["business_discovery"]["media"]["data"]) {
                  id = json_compete["business_discovery"]["media"]["data"][num]["id"];
                  like_count = json_compete["business_discovery"]["media"]["data"][num]["like_count"];
                  comments_count = json_compete["business_discovery"]["media"]["data"][num]["comments_count"];
                  url = json_compete["business_discovery"]["media"]["data"][num]["permalink"];
                  ts = json_compete["business_discovery"]["media"]["data"][num]["timestamp"];
                  caption = json_compete["business_discovery"]["media"]["data"][num]["caption"]
                  date = new Date(ts.split("T")[0].split("-")[0],ts.split("T")[0].split("-")[1]-1,ts.split("T")[0].split("-")[2],Number(ts.split("T")[1].split(":")[0])+9,ts.split("T")[1].split(":")[1])

                  if(String(like_count) == "undefined"){
                    like_count = 0
                  }
                  if(String(comments_count) == "undefined"){
                    comments_count = 0
                  }
                  
                  arr.push([id,like_count,comments_count,date,url,caption])
                }

                //Logger.log(arr)

                getCommentCompete(arr)

                let day_flg = false
                for(let k=0;k<new_datas.length;k++){
                  day_flg = false
                  if(new_datas[k][0] !== ""){
                    for(let l=0;l<arr.length;l++){
                      let new_date = new Date(arr[l][3].getFullYear(),arr[l][3].getMonth(),arr[l][3].getDate())
                      if(new_datas[k][0].getTime() == new_date.getTime()){
                        id = arr[l][0];
                        like_count = arr[l][1];
                        comments_count = arr[l][2];
                        url = arr[l][4];
                        ts = arr[l][3];
                        caption = arr[l][5];
                        tag = arr[l][6];
                          
                        if(new_datas[k][3] == "" || new_datas[k][3] == id){ //はじめて、もしくは同じIDのとき
                          day_flg = true;
                          break;
                        }
                      }
                    }
                    if(day_flg == true){
                      new_datas[k][3] = id
                      new_datas[k][7] = like_count
                      new_datas[k][8] = comments_count
                      new_datas[k][9] = like_count + comments_count
                      new_datas[k][1] = ts
                      new_datas[k][2] = url
                      new_datas[k][4] = caption
                      new_datas[k][5] = tag
                    }

                    if(new_datas[k][0].getTime() == today.getTime()){
                      new_datas[k-1][10] = follower_count
                      new_datas[k-1][6] = media_count
                    }
                  }
                }
                sheet_new.getRange(10,2,new_datas.length,11).setValues(new_datas)
                sheet_new.getRange(10,8,new_datas.length,5).setNumberFormat("#,##0") //★
                sheet_new.getRange(9,24).setValue(today)
                SpreadsheetApp.flush()

                const requests = {updateDimensionProperties: {
                  properties: {pixelSize: 40},
                  range: {sheetId: sheet_new.getSheetId(), startIndex:9, endIndex: sheet_new.getLastRow(), dimension: "ROWS"},
                  fields: "pixelSize"
                }};
                Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());
                
                //トップテン
                let datas_like_comment = sheet_new.getRange(10,3,sheet_new.getLastRow()-9,7).getValues()
                datas_like_comment.sort(sorting_desc_like)
                let arr_like1 = [];
                let arr_like2 = [];
                for(let i=0;i<10;i++){
                  if(i<5){
                  arr_like1.push([datas_like_comment[i][6],datas_like_comment[i][0],datas_like_comment[i][1]])
                  }else{
                    arr_like2.push([datas_like_comment[i][6],datas_like_comment[i][0],datas_like_comment[i][1]])
                  }
                }
                sheet_new.getRange(3,3,5,3).setValues(arr_like1)
                sheet_new.getRange(3,7,5,3).setValues(arr_like2)
                sheet_new.getRange(3,3,5,1).setNumberFormat("#,##0") //★
                sheet_new.getRange(3,7,5,1).setNumberFormat("#,##0") //★

                datas_like_comment = sheet_new.getRange(10,3,sheet_new.getLastRow()-9,8).getValues()
                datas_like_comment.sort(sorting_desc_comment)
                let arr_comment1 = [];
                let arr_comment2 = [];
                for(let i=0;i<10;i++){
                  if(i<5){
                    arr_comment1.push([datas_like_comment[i][7],datas_like_comment[i][0],datas_like_comment[i][1]])
                  }else{
                    arr_comment2.push([datas_like_comment[i][7],datas_like_comment[i][0],datas_like_comment[i][1]])
                  }
                }
                sheet_new.getRange(3,11,5,3).setValues(arr_comment1)
                sheet_new.getRange(3,15,5,3).setValues(arr_comment2)
                sheet_new.getRange(3,11,5,1).setNumberFormat("#,##0") //★
                sheet_new.getRange(3,15,5,1).setNumberFormat("#,##0") //★

                //sheet_new.getRange(6,4,10,3).setBackground(null)
                //sheet_new.getRange(6,7,10,3).setBackground(null)
                //トップテンに色をつける
                /*
                for(let i=0;i<arr_like.length;i++){
                  if(arr_like[i][1] !== ""){
                    for(let j=0;j<arr_comment.length;j++){
                      if(arr_comment[j][1] !== ""){
                        if(arr_like[i][1].getTime() == arr_comment[j][1].getTime()){
                          sheet_new.getRange(6+i,4).setBackground("#e2ce90")
                          sheet_new.getRange(6+j,7).setBackground("#e2ce90")
                        }
                      }
                    }
                  }
                }
                */
                getTagList_compete(sheet_new,sheet_new2)
                competeCount(sheet_new,sheet_new2)

              }catch(e){
                Logger.log(e.message)
              }
            }
          }
        }
      }
      for(let i=0;i<datas.length;i++){
        if(datas[i].length == 1){
          datas[i].push("")
        }
      }
      sheet.getRange(3,1,datas.length,2).setValues(datas)
    //}else{
      //Browser.msgBox("このシートはデータ取得ができません")
    //}

    competeAccount()
  }

function sorting_desc_like(a, b){
  if(a[6] > b[6]){
    return -1;
  }else if(a[6] < b[6] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_comment(a, b){
  if(a[7] > b[7]){
    return -1;
  }else if(a[7] < b[7] ){
    return 1;
  }else{
   return 0;
  }
}

function getCommentCompete(arr){

  let comment_text = ""
  let comment = ""

  //コメント検索
  for(let i=0;i<arr.length;i++){
    let url = "https://graph.facebook.com/v12.0/" + arr[i][2]+ "/comments?access_token=" + token
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


function getTagList_compete(sheet_new,sheet_new2){ //投稿一覧からタグ一覧を作成する
  
  const datas = sheet_new.getRange(10,6,sheet_new.getLastRow()-9,2).getValues();

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
    content = datas[i][1]
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
  if(arr.length > 0){
    sheet_new2.getRange(3,1,arr.length,1).setValues(arr)
  }
}

function competeCount(sheet_new,sheet_new2){
  
  const datas = sheet_new.getRange(10,1,sheet_new.getLastRow()-9,9).getValues();
  const datas2 = sheet_new2.getRange(2,1,sheet_new2.getLastRow()-1,sheet_new2.getLastColumn()).getValues();

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
        caption = datas[j][5]
        comment = datas[j][6]
        if(caption.indexOf(tag) !== -1 || comment.indexOf(tag) !== -1){ //タグがキャプションに存在する  
          like = datas[j][8]
          ts = datas[j][2]
          for(let k=1;k<datas2[0].length;k++){ 
            month = datas2[0][k]
            next_month = datas2[0][k+1]
            if(typeof(month)=="object"){
              if(ts.getTime() >= month.getTime()){
                if(typeof(next_month)=="object"){
                  if(ts.getTime() < next_month.getTime()){
                    if(datas2[i][k] == ""){
                      datas2[i][k] = 1
                      datas2[i][k+14] = Number(like)
                      break;
                    }else{
                      datas2[i][k] += 1
                      datas2[i][k+14] += Number(like)
                      break;
                    }
                  }
                }else{
                  if(datas2[i][k] == ""){
                      datas2[i][k] = 1
                      datas2[i][k+14] = Number(like)
                    break;
                  }else{
                      datas2[i][k] += 1
                      datas2[i][k+14] += Number(like)
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
  if(datas2.length > 0){
    for(let i=0;i<datas2.length;i++){
      for(let j=1;j<14;j++){
        if(datas2[i][14] == ""){
          datas2[i][14] = 0
        }
        if(datas2[i][28] == ""){
          datas2[i][28] = 0
        }
        datas2[i][14] += Number(datas2[i][j])
        datas2[i][28] += Number(datas2[i][j+15])
      }
    }
    datas2.sort(sorting_desc_compete)
    sheet_new2.getRange(3,1,datas2.length,datas2[0].length).setValues(datas2)
  }
}

function sorting_desc_compete(a, b){
  if(a[14] > b[14]){
    return -1;
  }else if(a[14] < b[14] ){
    return 1;
  }else{
   return 0;
  }
}