function findRow(sheet,val,col){

  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] == val){
      return i+1;
    }
  }
  return 0;
}

function findMultiRow(sheet,val,col){

  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  var targetRows = []
  var data = []
  for(var i=0;i<dat.length;i++){
    if(dat[i][col-1] == val){
      targetRows.push(i+1)
    }
  }
  targetRows = Array.from(new Set(targetRows))
  for (let i = 0; i < targetRows.length; i++) {
    // 検索にヒットしたレコードの取得
    let tmpdata = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn()).getValues();
    data.push(tmpdata[0]);
  }
  return data;
}

function crossfindRow(sheet,key1,col1,key2,col2){

  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i][col1-1] == key1 && dat[i][col2-1] == key2){
      return i+1;
    }
  }
  return 0;
}

function graduation(sheet,key){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  var result = []
  
  for(var i=1;i<dat.length;i++){
    if(Number(dat[i][0]) < (key+1)*10000 && dat[i][5] == "部員"){
      result.push(i+1);
    }
  }
  return result;
}

function findNearDataRow(sheet){

  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  //Dateオブジェクトからインスタンスを生成
  const today = new Date();
  for(var i=0;i<dat.length;i++){
    var dt = new Date(Number(dat[i][0]),Number(Number(dat[i][1])-1),Number(Number(dat[i][2])+1));
    if(dt<today){
      sheet.deleteRow(i+1)
    }else{
      return
    }
  }
}

function search(sheet,searchW){
  let data = []; // 検索にヒットしたデータの格納先配列
  dat = sheet.getDataRange().getValues();
  if (searchW == null || searchW == '') {
    return dat // 検索ワードがnullの場合は全件取得
 
  } else {
    let ranges = sheet.createTextFinder(searchW).findAll(); // キーワードによる検索を実施
    let targetRows = []; // 検索にヒットしたレコード行の格納先
    data.push(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0])
   
    // 検索にヒットしたRangeとレコード行を格納
    for (let i = 0; i < ranges.length; i++) {
      targetRows.push(ranges[i].getRow());
    }
    targetRows = Array.from(new Set(targetRows))
    for (let i = 0; i < targetRows.length; i++) {
      // 検索にヒットしたレコードの取得
      let tmpdata = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn()).getValues();
      data.push(tmpdata[0]);
    }
    return data
  }
}

function twoInt(number){
  if(number.length < 2){
    return "0" + number
  }else{
    return number
  }
}
function doGet(e) {
  // URLのexec/(またはdev/)以降を取得
  var page = e.pathInfo ? e.pathInfo : "index"


  // 該当するテンプレートを取得する
  var template = (() => {
    try {
      return HtmlService.createTemplateFromFile("index");
      //return HtmlService.createTemplateFromFile("templete");
    } catch(e) {
      return HtmlService.createTemplateFromFile("error");
    }
  })();

  var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');

  var LOGIN_USER = Session.getActiveUser().getEmail();
  try{
    var user_permission = member_db.getRange(findRow(member_db,LOGIN_USER,5),7).getValue()
    var user_name = member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue()
  }catch{
    var user_permission = "外部"
    var user_name = "匿名"
  }

  // htmlを返す
  template.user_name = user_name
  template.user_permission = user_permission
  template.page = page
  template.url = ScriptApp.getService().getUrl();   // テンプレートにアプリのURLを渡す
  return template.evaluate()                     // テンプレートを評価してhtmlを返す
    .setTitle("Polaris-U")                           // タイトルをセット
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');  // viewportを設定
}

function getData() {
  var LOGIN_USER = Session.getActiveUser().getEmail();
  var schedule_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部活日程');
  var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');
  var item_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機材情報');
  var absence_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('欠席連絡');
  var form_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('フォーム情報');
  var setting_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  var event_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('行事情報');

  switch(arguments[0]){
    case "index":
      findNearDataRow(schedule_db)
      schedule_db.getDataRange().sort({column: 10, ascending: true})
      return schedule_db.getRange(2,1,1,schedule_db.getLastColumn()).getValues()
    
    case "schedule_list":
      findNearDataRow(schedule_db)
      schedule_db.getDataRange().sort({column: 10, ascending: true})
      return ["",schedule_db.getRange(1,1,schedule_db.getLastRow(),schedule_db.getLastColumn()).getValues()]

    case "schedule_detail":
      var schedule = schedule_db.getRange(arguments[1],1,1,schedule_db.getLastColumn()).getValues()
      var key_date = twoInt(schedule[0][0]) + "-" + twoInt(schedule[0][1]) + "-" + twoInt(schedule[0][2])
      var key_activity = schedule[0][4]
      var absence_list = absence_db.getDataRange().getValues()
      var absence_result = []
      for(i=1;i<absence_list.length;i++){
        var check = absence_list[i][4].indexOf(key_activity);
        if(String(absence_list[i][3])== key_date && check != -1 ){
          absence_result.push(absence_list[i])
        }
      }

      if(member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue()==String(schedule[0][8])){
        var permittion = "allow"
      }else{
        var permittion = "reject"
      }
      // ["欠席者","活動情報","権限"]
      return [absence_result,schedule,permittion,arguments[1]]

    case "item_list_default":
      return item_db.getDataRange().getValues()
    
    case "absence_list":
      return search(absence_db,arguments[1])
    
    case "form_list":
      return form_db.getDataRange().getValues()

    case "form_public":
      return findMultiRow(form_db,"受付中",8)
    
    case "item_list_search":
      return search(item_db,arguments[1])

    case "item_inquery":
      judge = findRow(item_db,arguments[1],1)
      if(judge == 0){
        return ["bad"]
      }else{
        return ["ok",item_db.getRange(judge,2).getValue(),item_db.getRange(judge,4).getValue(),item_db.getRange(judge,5).getValue(),item_db.getRange(judge,6).getValue()]
      }
    
    case "absence_inquery":
      result = search(schedule_db,arguments[1])
      if(result.length == 1){
        return ["bad"]
      }else{
        return ["ok",result]
      }
    
    case "absence_edit_inquery":
      result = absence_db.getRange(arguments[1],1,1,6).getValues()
      if(result[0][0]== ""){
        return ["bad"]
      }else if(member_db.getRange(findRow(member_db,LOGIN_USER,5),1).getValue() != result[0][0]){
        return ["notallow"]
      }else{
        return ["ok",result]
      }
    
    case "form_inquery":
      result = form_db.getRange(arguments[1],1,1,9).getValues()
      if(result[0][0]== ""){
        return ["bad"]
      }else{
        return ["ok",result]
      }
    
    case "member_new_inquery":
      // コード認証
      right_code = setting_db.getRange(1,2).getValue()
      judge = findMultiRow(member_db,arguments[2],1)
      if(right_code != arguments[1]){
        return ["bad"]
      }else if(judge.length != 0){
        return ["already"]
      }else{
        return ["ok"]
      }
    
    case "member_list_search":
      member_db.getRange(2,1,member_db.getLastRow()-1,member_db.getLastColumn()).sort({column: 2, ascending: true})
      return search(member_db,arguments[1]);

    case "member_graduetion":
      target = graduation(member_db,arguments[1])
      for(var i=0;i<target.length;i++){
        member_db.getRange(target[i],6).setValue("引退")
      }
      return
    
    case "member_edit_inquery":
      result_row = findRow(member_db,arguments[1],1)
      if(result_row == 0){
        return ["bad"]
      }else{
        result = member_db.getRange(result_row,1,1,member_db.getLastColumn()).getValues()
        return ["ok",result[0]]
      }
    
    case "event_list":
      result = []
      getsheetdata = event_db.getDataRange().getValues()
      for(var i=0;i<getsheetdata.length;i++){
        result.push(getsheetdata[i][0])
      }
      return result
    
    case "event_detail":
      return event_db.getRange(Number(arguments[1])+1,1,1,9).getValues()
  }
}

function sendData() {
  var LOGIN_USER = Session.getActiveUser().getEmail();
  var schedule_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部活日程');
  var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');
  var item_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機材情報');
  var absence_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('欠席連絡');
  var form_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('フォーム情報');
  findNearDataRow(schedule_db)

  switch(arguments[0]){
    case "schedule_new":
      var WeekChars = [ "日", "月", "火", "水", "木", "金", "土" ];
      var date = String(arguments[1])
      var date_array = date.split('-');
      var hold_day = new Date(Number(date_array[0]),Number(date_array[1])-1,date_array[2].split('T')[0])
      var day_youbi = WeekChars[hold_day.getDay()]
      schedule_db.appendRow([date_array[0],date_array[1],date_array[2].split('T')[0],day_youbi,String(date_array[2].split('T')[1]),arguments[2],arguments[3],arguments[4],member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[1],schedule_db.getLastRow()+1]);
      schedule_db.getRange(schedule_db.getLastRow(),10).setNumberFormat('yyyy"-"mm"-"dd hh":"mm');
      schedule_db.getRange(schedule_db.getLastRow(),1).setNumberFormat('@');
      schedule_db.getRange(schedule_db.getLastRow(),2).setNumberFormat('@');
      schedule_db.getRange(schedule_db.getLastRow(),3).setNumberFormat('@');
      schedule_db.getRange(schedule_db.getLastRow(),5).setNumberFormat('@');
      schedule_db.getRange(schedule_db.getLastRow(),10).setNumberFormat('@');
      return [arguments[2],String(arguments[1])]

    case "schedule_update":
      var row = arguments[1]
      var WeekChars = [ "日", "月", "火", "水", "木", "金", "土" ];
      var date = String(arguments[2])
      var date_array = date.split('-');
      var hold_day = new Date(Number(date_array[0]),Number(date_array[1])+1,date_array[2].split('T')[0])
      var day_youbi = WeekChars[hold_day.getDay()]
      schedule_db.getRange(row,1,1,10).setValues([[date_array[0],date_array[1],date_array[2].split('T')[0],day_youbi,String(date_array[2].split('T')[1]),arguments[3],arguments[4],arguments[5],member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[2]]]);
      schedule_db.getRange(row,10).setNumberFormat('yyyy"-"mm"-"dd hh":"mm');
      schedule_db.getRange(row,1).setNumberFormat('@');
      schedule_db.getRange(row,2).setNumberFormat('@');
      schedule_db.getRange(row,3).setNumberFormat('@');
      schedule_db.getRange(row,5).setNumberFormat('@');
      schedule_db.getRange(row,10).setNumberFormat('@');
      return [arguments[2],String(arguments[1])]
    
    case "schedule_delete":
      var row = arguments[1]
      schedule_db.deleteRow(row)
      return

    case "item_new":
      item_db.appendRow([arguments[1],arguments[2],arguments[3],"健康",arguments[4],arguments[5]]);
      item_db.getRange(item_db.getLastRow(),3).setNumberFormat('@')
      return [arguments[1],arguments[2]]
    
    case "item_update":
      rownumber = findRow(item_db,arguments[1],1)
      if(rownumber == 0){
        return ["失敗しました","該当の機材IDが見つかりませんでした"]
      }else{
        item_db.getRange(rownumber,4,1,3).setValues([[String(arguments[3]),String(arguments[4]),String(arguments[5])]])
      }
      return [arguments[1],arguments[2]]

    case "item_update":
      rownumber = findRow(item_db,arguments[1],1)
      if(rownumber == 0){
        return ["失敗しました","該当の機材IDが見つかりませんでした"]
      }else{
        item_db.getRange(rownumber,4,1,3).setValues([[String(arguments[3]),String(arguments[4]),String(arguments[5])]])
      }
      return [arguments[1],arguments[2]]

    case "absence_new":
      absence_db.appendRow([member_db.getRange(findRow(member_db,LOGIN_USER,5),1).getValue(),member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[3],arguments[1],arguments[2],arguments[4],absence_db.getLastRow()+1]);
      absence_db.getRange(absence_db.getLastRow(),4).setNumberFormat('@')
      absence_db.getRange(absence_db.getLastRow(),1).setNumberFormat('@')
      absence_db.getRange(absence_db.getLastRow(),7).setNumberFormat('@')
      return [member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[1],arguments[2]]
    
    case "absence_new_slist":
      absence_db.appendRow([member_db.getRange(findRow(member_db,LOGIN_USER,5),1).getValue(),member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[2],schedule_db.getRange(arguments[1],10).getValue().split(" ")[0],(schedule_db.getRange(arguments[1],10).getValue().split(" ")[1] + " " + schedule_db.getRange(arguments[1],6).getValue()),arguments[3],absence_db.getLastRow()+1]);
      absence_db.getRange(absence_db.getLastRow(),4).setNumberFormat('@')
      absence_db.getRange(absence_db.getLastRow(),1).setNumberFormat('@')
      absence_db.getRange(absence_db.getLastRow(),7).setNumberFormat('@')
      return [member_db.getRange(findRow(member_db,LOGIN_USER,5),4).getValue(),arguments[1],arguments[2]]
    
    case "absence_delete":
      absence_db.deleteRow(arguments[1])
      return []

    case "form_new":
      form_db.appendRow([arguments[1].split("-")[1],arguments[1].split("-")[2].split("T")[0],arguments[1].split("T")[1],arguments[2],arguments[3],arguments[4],arguments[5],"受付中",form_db.getLastRow()+1]);
      form_db.getRange(form_db.getLastRow(),1).setNumberFormat('@')
      form_db.getRange(form_db.getLastRow(),2).setNumberFormat('@')
      form_db.getRange(form_db.getLastRow(),3).setNumberFormat('@')
      return [arguments[2],arguments[3]]
    
    case "form_update":
      form_db.getRange(arguments[1],1,1,8).setValues([[arguments[2],arguments[3],arguments[4],arguments[5],arguments[6],arguments[7],arguments[8],arguments[9]]])
      form_db.getRange(absence_db.getLastRow(),1).setNumberFormat('@')
      form_db.getRange(absence_db.getLastRow(),2).setNumberFormat('@')
      form_db.getRange(absence_db.getLastRow(),3).setNumberFormat('@')
      return []
    
    case "form_delete":
      form_db.deleteRow(arguments[1])
      return []
    
    case "member_new":
      member_db.appendRow([arguments[1],arguments[2],arguments[3],arguments[4],LOGIN_USER,"部員"]);
      member_db.getRange(form_db.getLastRow(),1).setNumberFormat('@')
      member_db.getRange(form_db.getLastRow(),2).setNumberFormat('@')
      return [arguments[1],arguments[4]]
    
    case "member_update":
      member_db.getRange(findRow(member_db,arguments[1],1),1,1,6).setValues([[arguments[1],arguments[2],arguments[3],arguments[4],arguments[5],arguments[6]]])
      member_db.getRange(member_db.getLastRow(),1).setNumberFormat('@')
      member_db.getRange(member_db.getLastRow(),2).setNumberFormat('@')
      return [arguments[1],arguments[4]]
    
    case "member_upgrade":
      member_db.getRange(findRow(member_db,arguments[1],1),2).setValue(arguments[2])
      member_db.getRange(member_db.getLastRow(),2).setNumberFormat('@')
      return []
    
    case "member_delete":
      member_db.deleteRow(arguments[1])
      return []
  }
}