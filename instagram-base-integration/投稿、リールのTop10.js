function getTopten_new(){

  const sheet = ss.getSheetByName("投稿");
  const sheet_topten = ss.getSheetByName("トップテン")
  sheet_topten.getRange(3,2,10,9).clearContent()

  let cnt = 0

  //トップテン（投稿）
  let datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,13).getValues()
  datas_like_comment.sort(sorting_desc_like2)

  if(datas_like_comment.length < 10){
    cnt = datas_like_comment.length
  }else{
    cnt = 10
  }

  let arr_like = [];
  let new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_like.push([datas_like_comment[i][7],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_like.length > 0){
    sheet_topten.getRange(3,2,new_cnt,3).setValues(arr_like)
  }

  datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,13).getValues()
  datas_like_comment.sort(sorting_desc_save2)
  let arr_save = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_save.push([datas_like_comment[i][9],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_save.length > 0){
    sheet_topten.getRange(3,5,new_cnt,3).setValues(arr_save)
  }

  datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,13).getValues()
  datas_like_comment.sort(sorting_desc_reach2)
  let arr_reach = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_reach.push([datas_like_comment[i][12],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_reach.length > 0){
    sheet_topten.getRange(3,8,new_cnt,3).setValues(arr_reach)
  }

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

  getTopten_Reel_new()
}

function getTopten_Reel_new(){

  const sheet = ss.getSheetByName("リール");
  const sheet_topten = ss.getSheetByName("トップテン")
  sheet_topten.getRange(18,2,10,9).clearContent()

  let cnt = 0

  //トップテン（リール）
  let datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,14).getValues()
  datas_like_comment.sort(sorting_desc_like2)

  if(datas_like_comment.length < 10){
    cnt = datas_like_comment.length
  }else{
    cnt = 10
  }

  let arr_like = [];
  let new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_like.push([datas_like_comment[i][7],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_like.length > 0){
    sheet_topten.getRange(18,2,new_cnt,3).setValues(arr_like)
  }

  datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,14).getValues()
  datas_like_comment.sort(sorting_desc_save2)
  let arr_save = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_save.push([datas_like_comment[i][9],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_save.length > 0){
    sheet_topten.getRange(18,5,new_cnt,3).setValues(arr_save)
  }

  datas_like_comment = sheet.getRange(4,3,sheet.getLastRow()-1,14).getValues()
  datas_like_comment.sort(sorting_desc_reach2_reel)
  let arr_reach = [];
  new_cnt = 0
  for(let i=0;i<datas_like_comment.length;i++){
    if(new_cnt < cnt){
      arr_reach.push([datas_like_comment[i][13],datas_like_comment[i][0],datas_like_comment[i][2]])
      new_cnt++
    }else{
      break;
    }
  }
  if(arr_reach.length > 0){
    sheet_topten.getRange(18,8,new_cnt,3).setValues(arr_reach)
  }
  
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


function sorting_desc_like2(a, b){
  if(a[7] > b[7]){
    return -1;
  }else if(a[7] < b[7] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_save2(a, b){
  if(a[9] > b[9]){
    return -1;
  }else if(a[6] < b[6] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_reach2(a, b){
  if(a[12] > b[12]){
    return -1;
  }else if(a[12] < b[12] ){
    return 1;
  }else{
   return 0;
  }
}

function sorting_desc_reach2_reel(a, b){
  if(a[13] > b[13]){
    return -1;
  }else if(a[13] < b[13] ){
    return 1;
  }else{
   return 0;
  }
}