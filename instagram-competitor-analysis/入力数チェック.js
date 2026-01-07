function countCheck(e) {
  
  const sheet = ss.getSheetByName("競合");
  const datas = sheet.getRange(3,1,sheet.getLastRow()-2,1).getValues();

  let cnt = 0
  for(let i=0;i<datas.length;i++){
    if(datas[i][0] !== ""){
      cnt += 1
    }
  }
  if(cnt > 5){
    Browser.msgBox("最大入力数を超えています")
    e.range.clear()
  }
  
}
