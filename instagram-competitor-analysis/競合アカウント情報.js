function competeAccount() {
  const sheets = ss.getSheets();
  
  const sheet1 = ss.getSheetByName("競合");
  const datas1 = sheet1.getRange(3,1,sheet1.getLastRow()-2,12).getValues()

  let name = ""
  let datas2 = ""
  let count_arr = ""
  
  for(let i=0;i<datas1.length;i++){
    name = datas1[i][0]
    if(name !== ""){
      for(let j=0;j<sheets.length;j++){
        if(sheets[j].getSheetName() == name){
          datas2 = sheets[j].getRange(10,2,sheets[j].getLastRow()-9,11).getValues()
          for(let k=0;k<datas2.length;k++){
            if(today.getTime() == datas2[k][0].getTime()){
              datas1[i][2] = datas2[k-1][10]
              datas1[i][3] = datas2[k-1][6]
            }
          }
          count_arr = sheets[j].getRange(3,21,1,8).getValues()
          datas1[i][4] = count_arr[0][0]
          datas1[i][5] = count_arr[0][1]
          datas1[i][6] = count_arr[0][2]
          datas1[i][7] = count_arr[0][3]
          datas1[i][8] = count_arr[0][4]
          datas1[i][9] = count_arr[0][5]
          datas1[i][10] = count_arr[0][6]
          datas1[i][11] = count_arr[0][7]
          break;
        }
      }
    }
  }
  sheet1.getRange(3,1,datas1.length,12).setValues(datas1)
}
