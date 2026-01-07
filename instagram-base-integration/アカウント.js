function getAccount() { //アカウントインサイト検索（day:インプレッション、プロフィールビュー、新規フォロワー数）

  const sheet = ss.getSheetByName("アカウント情報");
  const datas = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues()

  let name = ""
  let impressions = ""
  let reach = ""
  let profile_views = ""
  let website_clicks = ""

  const url_user = "https://graph.facebook.com/v22.0/" + business_id + "?fields=followers_count&access_token=" + token
  let encodedURI_user = encodeURI(url_user);
  let response_user = UrlFetchApp.fetch(encodedURI_user,options).getContentText();
  let json_user = JSON.parse(response_user)
  let follower_count = json_user["followers_count"]


  
  //const metric_user_day = "impressions,profile_views,reach,follower_count,website_clicks&period=day"
  const metric_user_day = "profile_views,reach,website_clicks&period=day&metric_type=total_value"
  //const metric_user_day = "reach&period=day"
  //const metric_user_day = "profile_activity&breakdown=action_type"
  const url_user_day_insight = "https://graph.facebook.com/v22.0/" + business_id + "/insights?metric=" + metric_user_day + "&access_token=" + token
  let encodedURI_user_day_insight = encodeURI(url_user_day_insight);
  let response_user_day_insight = UrlFetchApp.fetch(encodedURI_user_day_insight,options).getContentText();
  let json_user_day_insight = JSON.parse(response_user_day_insight)

  let arr = [];
  let arr2 = [];
  //let date = ""
  for (var num in json_user_day_insight["data"]) {
    name = json_user_day_insight["data"][num]["name"];
    //date = json_user_day_insight["data"][num]["values"][1]["end_time"];
    //date = new Date(date.split("T")[0].split("-")[0],date.split("T")[0].split("-")[1]-1,date.split("T")[0].split("-")[2])
    switch (name) {
      case "impressions":
        impressions = json_user_day_insight["data"][num]["total_value"]["value"];
        break
    　case "profile_views":
        profile_views = json_user_day_insight["data"][num]["total_value"]["value"];
        break
    　case "reach":
        reach = json_user_day_insight["data"][num]["total_value"]["value"];
        break
      case "website_clicks":
        website_clicks = json_user_day_insight["data"][num]["total_value"]["value"];
        break
      //case "follower_count":
        //new_follower_count = json_user_day_insight["data"][num]["values"][1]["value"];
        //break
    }
  }
  arr.push([reach,impressions,profile_views,follower_count])
  arr2.push([website_clicks])

  let sheet_date = ""
  for(let i=0;i<datas.length;i++){
    sheet_date = datas[i][0]
    if(sheet_date.getTime() == today.getTime()){
      sheet.getRange(i+2,2,1,arr[0].length).setValues(arr)
      sheet.getRange(i+2,7,1,1).setValues(arr2)
      break;
    }
  }

  all_Triggers() //各種トリガーを同じタイミングでリフレッシュする
}

function getFollowers() { //フォロワー検索

  const sheet = ss.getSheetByName("フォロワー情報");
  sheet.getRange(4,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  const sheet_country = ss.getSheetByName("国名コード");
  const country_datas = sheet_country.getRange(2,1,sheet_country.getLastRow(),2).getValues()
  const sheet2 = ss.getSheetByName("フォロワー属性変化");
  const datas = sheet2.getRange(4,1,sheet2.getLastRow(),sheet2.getLastColumn()).getValues()

  const now = new Date();
  const now_dt = Math.round(now.getTime() / 1000) 

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate()-1)
  const yesterday_dt = Math.round(yesterday.getTime() / 1000)
  const yesterday_st = Utilities.formatDate(yesterday,"JST","yyyy/MM/dd")

  const yesterday2 = new Date(yesterday.getFullYear(),yesterday.getMonth(),yesterday.getDate());

  let name = ""
  let audience_city = ""
  let audience_country = ""
  let audience_gender_age = ""

  //const metric_user_lifetime = "follower_demographics&period=lifetime&breakdown=city,country,gender,age&metric_type=total_value"//"audience_city,audience_country,audience_gender_age&period=lifetime"



  const metric_user_lifetime_city = "follower_demographics&period=lifetime&breakdown=city&metric_type=total_value"
  const url_userinsight_city = "https://graph.facebook.com/v22.0/" + business_id + "/insights?metric=" + metric_user_lifetime_city + "&access_token=" + token
  let encodedURI_userinsight_city = encodeURI(url_userinsight_city);
  let response_userinsight_city = UrlFetchApp.fetch(encodedURI_userinsight_city,options).getContentText();
  let json_userinsigh_city = JSON.parse(response_userinsight_city)

  let city = ""
  let city_cnt = ""
  let arr_city = []
  for (var num in json_userinsigh_city["data"][0]["total_value"]["breakdowns"][0]["results"]) {
    city = json_userinsigh_city["data"][0]["total_value"]["breakdowns"][0]["results"][num]["dimension_values"][0];
    city_cnt = json_userinsigh_city["data"][0]["total_value"]["breakdowns"][0]["results"][num]["value"]
    arr_city.push([city,city_cnt])
  }
  arr_city.sort(sorting_desc)
  sheet.getRange(4,1,arr_city.length,2).setValues(arr_city)
  


  const metric_user_lifetime_country = "follower_demographics&period=lifetime&breakdown=country&metric_type=total_value"
  const url_userinsight_country = "https://graph.facebook.com/v22.0/" + business_id + "/insights?metric=" + metric_user_lifetime_country + "&access_token=" + token
  let encodedURI_userinsight_country = encodeURI(url_userinsight_country);
  let response_userinsight_country = UrlFetchApp.fetch(encodedURI_userinsight_country,options).getContentText();
  let json_userinsight_country = JSON.parse(response_userinsight_country)

  let country = ""
  let country_cnt = ""
  let arr_country = []
  for (var num in json_userinsight_country["data"][0]["total_value"]["breakdowns"][0]["results"]) {
    country = json_userinsight_country["data"][0]["total_value"]["breakdowns"][0]["results"][num]["dimension_values"][0];
    country_cnt = json_userinsight_country["data"][0]["total_value"]["breakdowns"][0]["results"][num]["value"]
    for(let i=0;i<country_datas.length;i++){
      if(country == country_datas[i][0]){
        arr_country.push([country_datas[i][1],country_cnt])
        break;
      }
    }
    //arr_country.push([country,country_cnt])
  }
  arr_country.sort(sorting_desc)
  sheet.getRange(4,4,arr_country.length,2).setValues(arr_country)
  

  const metric_user_lifetime_age = "follower_demographics&period=lifetime&breakdown=age&metric_type=total_value"
  const url_userinsight_age = "https://graph.facebook.com/v22.0/" + business_id + "/insights?metric=" + metric_user_lifetime_age + "&access_token=" + token
  let encodedURI_userinsight_age = encodeURI(url_userinsight_age);
  let response_userinsight_age = UrlFetchApp.fetch(encodedURI_userinsight_age,options).getContentText();
  let json_userinsight_age = JSON.parse(response_userinsight_age)

  let age = ""
  let age_cnt = ""
  let arr_age = []
  for (var num in json_userinsight_age["data"][0]["total_value"]["breakdowns"][0]["results"]) {
    age = json_userinsight_age["data"][0]["total_value"]["breakdowns"][0]["results"][num]["dimension_values"][0];
    age_cnt = json_userinsight_age["data"][0]["total_value"]["breakdowns"][0]["results"][num]["value"]
    arr_age.push([age,age_cnt])
  }
  //arr_age.sort(sorting_desc)
  sheet.getRange(4,7,arr_age.length,2).setValues(arr_age)



  const metric_user_lifetime_gender = "follower_demographics&period=lifetime&breakdown=gender&metric_type=total_value"
  const url_userinsight_gender = "https://graph.facebook.com/v22.0/" + business_id + "/insights?metric=" + metric_user_lifetime_gender + "&access_token=" + token
  let encodedURI_userinsight_gender = encodeURI(url_userinsight_gender);
  let response_userinsight_gender = UrlFetchApp.fetch(encodedURI_userinsight_gender,options).getContentText();
  let json_userinsight_gender = JSON.parse(response_userinsight_gender)

  let gender = ""
  let gender_cnt = ""
  let arr_gender = []
  for (var num in json_userinsight_gender["data"][0]["total_value"]["breakdowns"][0]["results"]) {
    gender = json_userinsight_gender["data"][0]["total_value"]["breakdowns"][0]["results"][num]["dimension_values"][0];
    gender_cnt = json_userinsight_gender["data"][0]["total_value"]["breakdowns"][0]["results"][num]["value"]
    switch(gender){
    case "F":
      gender = "女性"
      break;
    case "M":
      gender = "男性"
      break;
    case "U":
      gender = "不明"
      break;
    }
    arr_gender.push([gender,gender_cnt])
  }
  //arr_age.sort(sorting_desc)
  sheet.getRange(4,10,arr_gender.length,2).setValues(arr_gender)

  sheet.getRange(1,1).setValue(Utilities.formatDate(now,"JST","yyyy/MM/dd HH:mm") + "時点")

/*
  let arr_city = [];
  let arr_country = [];
  let arr_gender = [];
  let arr = [];
  let arr2 = [];
  let arr_age_format = ["13-17","18-24","25-34","35-44","45-54","55-64","65+","13-17","18-24","25-34","35-44","45-54","55-64","65+","13-17","18-24","25-34","35-44","45-54","55-64","65+"]
  //let date = ""
  let date2 = yesterday_st + " 17:00時点"
  let gender = ""
  for (var num in json_userinsight["data"]) {
    name = json_userinsight["data"][num]["name"];
    //date = json_userinsight["data"][num]["values"][0]["end_time"];
    //date = new Date(date.split("T")[0].split("-")[0],date.split("T")[0].split("-")[1]-1,date.split("T")[0].split("-")[2])
    //date2 = Utilities.formatDate(date,"JST","yyyy/MM/dd") + " 17:00時点"
    switch (name) {
      case "audience_city":
        audience_city = json_userinsight["data"][num]["values"][0]["value"];
        for (var key in audience_city) {
          arr_city.push([key,audience_city[key]])
        }
        arr_city.sort(sorting_desc)
        sheet.getRange(4,1,arr_city.length,2).setValues(arr_city)
        break
    　case "audience_country":
        audience_country = json_userinsight["data"][num]["values"][0]["value"];
        for (var key in audience_country) {
          for(let i=0;i<country_datas.length;i++){
            if(key == country_datas[i][0]){
              arr_country.push([country_datas[i][1],audience_country[key]])
              break;
            }
          }
        }
        arr_country.sort(sorting_desc)
        sheet.getRange(4,4,arr_country.length,2).setValues(arr_country)
        break
    　case "audience_gender_age":
        audience_gender_age = json_userinsight["data"][num]["values"][0]["value"];

        //2022/3/28 フォロワー属性修正
        for (var key in audience_gender_age) {
          arr2.push(key.split(".")[1])
        }

        for (var key in audience_gender_age) {
          switch(key.split(".")[0]){
          case "F":
            gender = "女性"
            break;
          case "M":
            gender = "男性"
            break;
          case "U":
            gender = "不明"
            break;
          }
          arr_gender.push([gender,key.split(".")[1],audience_gender_age[key]])
          arr.push(audience_gender_age[key])
        }
        sheet.getRange(4,7,arr_gender.length,3).setValues(arr_gender)
        break
    }
  }
  sheet.getRange(1,1).setValue(date2)

  let arr_age = [];
  let age_num = 0
  for(let i=0;i<arr2.length;i++){
    for(let j=age_num;j<arr_age_format.length;j++){
      if(arr2[i] == arr_age_format[j]){
        arr_age.push(arr[i])
        age_num = j+1
        break;
      }else{
        arr_age.push("")
      }
    }
  }

  let sheet_date = ""
  for(let i=0;i<datas.length;i++){
    sheet_date = datas[i][0]
    //if(sheet_date.getTime() == date.getTime()){
    if(sheet_date.getTime() == yesterday2.getTime()){
      sheet2.getRange(i+4,23,1,arr_age.length).setValues([arr_age])
      break;
    }
  }
  */
}

function sorting_desc(a, b){
  if(a[1] > b[1]){
    return -1;
  }else if(a[1] < b[1] ){
    return 1;
  }else{
   return 0;
  }
}