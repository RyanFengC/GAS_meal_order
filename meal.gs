
 // 和Google試算表連動
var SpreadSheet = SpreadsheetApp.getActive();
var Sheet_parameter = SpreadSheet.getSheetByName("參數");
var Sheet_origin = SpreadSheet.getSheetByName("原始資料");
var Sheet_order_meal = SpreadSheet.getSheetByName("資料整理");
var Sheet_result = SpreadSheet.getSheetByName("結果");
var date = new Date();
var CHANNEL_ACCESS_TOKEN=Sheet_parameter.getRange(5,3).getValue()

function doPost(e) {
 
  var msg = JSON.parse(e.postData.contents);
  
    
  try {
      
    // 取出 replayToken 和 發送的訊息文字
    var replyToken = msg.events[0].replyToken;
    var userMessage = msg.events[0].message.text;                 
    var userID = msg.events[0].source.userId;   
    if (typeof replyToken === 'undefined') {
      return;
    }
    

    var Sheet_origin_length = Sheet_origin.getLastRow()
    // 特殊指令的處理
    switch(userMessage) {

      // 清空
      case 'clear':
      case '/clear':
      case '清空':
      case '清除':
        clear_all_report();
        //send_msg(CHANNEL_ACCESS_TOKEN, replyToken, '已清空');
        Sheet_parameter.getRange(2,1).setValue('off');
        return;

      // 列出當前回報情況
      case '/list':
      case 'list':
      case '結單':
      case '結果':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken ,Sheet_result.getRange(1,1).getValue());
        return;

      //開始統計(將參數設為on)
      case '開始':
        clear_all_report();
        Sheet_parameter.getRange(2,1).setValue('on');
        return;

      //寄送結果至電子信箱
      case 'Email':
      case '電子郵件':
      case 'email':
      case 'mail':
        send_email()//寄送電子信箱
        return  

      //啟動加1模式
      case '加1模式':
      case '加ㄧ模式':
      case '+1模式':
      case '+一模式':
        clear_all_report();
        Sheet_parameter.getRange(2,2).setValue('on');
        return;

      default:
        console.log('not special command')
    }
    
  //replace用法:使用反斜線(/)表示需替換之符號,如/ 表示需替換空白，
  //           /g表示替換全部,如/ /g表示替換全部字串中的空白符號，如多個條件則使用or(|)
  //console.log(userMessage_edit_split.length)
    // 訊息內容資料分割
    userMessage_edit = userMessage.replace(/\r?\n|\r/g,',')//取代換行
    userMessage_edit = userMessage_edit.replace(/(×)|(x)|(X)/g,'*')//取代乘(x)符號
    userMessage_edit = userMessage_edit.replace(/ /g,'')//取代空白
    userMessage_edit_split = userMessage_edit.split(',')//從,做分割
      
    // 當參數表中顯示為on才開始統計，以免統計錯誤
    if(Sheet_parameter.getRange(2, 1).getValue()=='on'){
      //輸入原始資料分頁
      var username = JSON.parse(getusername(CHANNEL_ACCESS_TOKEN,userID)).displayName//取得輸入人員名稱
      Sheet_origin.getRange(Sheet_origin_length + 1, 1).setValue(Sheet_origin_length);//填入項次欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 3).setValue(userMessage);//填入內容欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 2).setValue(username);//填入訊息發送者欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 4).setValue(date);//填入時間欄位
      //輸入統計資料分頁
      data_arange(userMessage,username)
      Sheet_result.getRange(1,1).setValue(print_report_list(Sheet_order_meal,Sheet_parameter))

    }else if(Sheet_parameter.getRange(2, 2).getValue()=='on'){
      //----------------------------加1模式-----------------------------------
      //輸入原始資料分頁
      var username = JSON.parse(getusername(CHANNEL_ACCESS_TOKEN,userID)).displayName//取得輸入人員名稱
      Sheet_origin.getRange(Sheet_origin_length + 1, 1).setValue(Sheet_origin_length);//填入項次欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 3).setValue(userMessage);//填入內容欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 2).setValue(username);//填入訊息發送者欄位
      Sheet_origin.getRange(Sheet_origin_length + 1, 4).setValue(date);//填入時間欄位
      //輸入統計資料分頁
      add_one_mode(userMessage,username)
      Sheet_result.getRange(1,1).setValue(print_report_list(Sheet_order_meal,Sheet_parameter))

    }/*else{
       send_msg(CHANNEL_ACCESS_TOKEN, replyToken,'統計未開始，請先輸入 開始 或 加1模式 ');
    }*/
    
  }
  catch(err) {
    console.log(err);
  }
}
//加1模式訊息處理
function add_one_mode(message,name){
  message = message.replace(/(×)|(x)|(X)|(\+)/g,'*')
  mesg_split = message.replace(/ /g,'')
  meal_number = Sheet_order_meal.getRange(3,4).getValue()
  meal = mesg_split.split('*')
  
  if(isNaN(parseInt(meal_number))==true){
    total_number=parseInt(meal[1])//計算總數量
  }else{
    total_number=parseInt(meal_number)+parseInt(meal[1])//計算總數量
  }
  
  Sheet_order_meal.getRange(3,4).setValue(total_number)
  Sheet_order_meal.getRange(3,5).setValue(Sheet_order_meal.getRange(3,5).getValue()
  + '\r\n' + name + '*' + meal[1])
  //計算目前總金額
  total_money = parseInt(total_number)*parseInt(Sheet_order_meal.getRange(3,3).getValue())
  Sheet_order_meal.getRange(3,7).setValue(total_money)    

}
//一般模式訊息處理
//輸入之格式需限定如多個品項用換行
//輸入之格式須為  餐點 價錢 x 數量
function data_arange(message,name){
  message = message.replace(/\r?\n|\r/g,',')
  message = message.replace(/(×)|(x)|(X)|(\+)/g,'*')
  message = message.replace(/ /g,'')
  mesg_split = message.split(',')
  //replace用法:使用反斜線(/)表示需替換之符號,如/ 表示需替換空白，
  //           /g表示替換全部,如/ /g表示替換全部字串中的空白符號，如多個條件則使用or(|)
  console.log(mesg_split.length)
  console.log(mesg_split)


  for(i=0; i < mesg_split.length; i++){
    meal = mesg_split[i].split('*')
    var Sheet_order_meal_length= Sheet_order_meal.getLastRow()
    var check_index = NaN //用於判斷資料庫是否已存在該品名
    //console.log(meal)
    for(ii=0; ii < Sheet_order_meal_length; ii++){

      meal_name = Sheet_order_meal.getRange(ii+1,2).getValue()
      meal_number = Sheet_order_meal.getRange(ii+1,4).getValue()

      //判斷資料庫中是否已存在此餐點
      if(meal_name == determine_money(meal[0])[0]){
        check_index=ii
        break
      }
    }
//餐點未存在資料庫
    if(isNaN(check_index)){
            //輸入項次欄位
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,1).setValue(Sheet_order_meal_length - 1)  
            //輸入餐點欄位
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,2).setValue(determine_money(meal[0])[0])
            //輸入價錢欄位
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,3).setValue(determine_money(meal[0])[1])
            //輸入數量欄位
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,4).setValue(meal[1])    
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,5).setValue(name+ '*' + meal[1])
            //計算目前總金額
            total_money = parseInt(determine_money(meal[0])[1])*parseInt(meal[1])
            Sheet_order_meal.getRange(Sheet_order_meal_length + 1,7).setValue(total_money)    
    }else{//餐點已存在資料庫
        total_number=parseInt(meal_number)+parseInt(meal[1])//計算總數量
        Sheet_order_meal.getRange(check_index+1,4).setValue(total_number)
        Sheet_order_meal.getRange(check_index+1,5).setValue(Sheet_order_meal.getRange(check_index+1,5).getValue()
        + ',' + name + '*' + meal[1])
        //計算目前總金額
        total_money = parseInt(total_number)*parseInt(Sheet_order_meal.getRange(check_index+1,3).getValue())
        Sheet_order_meal.getRange(check_index + 1,7).setValue(total_money)    

    }

  }
}

//取得訊息中之價錢
function determine_money(meal){
  if(isNaN(Number(meal.substr(-3)))==false){
    meal_name = meal.split(meal.substr(-3))[0]
    meal_money = Number(meal.substr(-3))

  }else if(isNaN(Number(meal.substr(-2)))==false){
    meal_name = meal.split(meal.substr(-2))[0]
    meal_money = Number(meal.substr(-2))
  }
  return [meal_name,meal_money]
}

//取得輸入人之名稱
function getusername(CHANNEL_ACCESS_TOKEN,userID){

  var url = "https://api.line.me/v2/bot/profile/" + userID;

  return UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'get',
  });
}
  




function print_report_list(Sheet_order_meal,Sheet_parameter) {
  report = '';
  for(i = 2; i <Sheet_order_meal.getLastRow(); i++){
      report = report +  Sheet_order_meal.getRange(i+1, 2).getValue() + '*' 
              + Sheet_order_meal.getRange(i+1, 4).getValue()+ '\r\n';
  }
  list_url = '結果顯示頁面\r\n'+Sheet_parameter.getRange(2, 3).getValue();

  report = report +'-------\r\n'+'總金額'+ Sheet_order_meal.getRange(2, 9).getValue()+ '\r\n'
            +'總數量'+ Sheet_order_meal.getRange(2, 10).getValue()+ '\r\n' + list_url;
  return report;

}



function clear_all_report() {
  for(var i= 0; i<100000000;i++ ){
    if (Sheet_origin.getLastRow()==1){
      break
    }
    else{
      Sheet_origin.deleteRow(Sheet_origin.getLastRow())
    }
  }
  for(var i= 0; i<1000000;i++ ){
    if (Sheet_order_meal.getLastRow()==2){
      break
    }
    else{
      Sheet_order_meal.deleteRow(Sheet_order_meal.getLastRow())
    }
  }
  Sheet_result.getRange(1,1).setValue('無資料')
  Sheet_parameter.getRange(2,2).setValue('off');
  Sheet_parameter.getRange(2,1).setValue('off');
}


function send_msg(token, replyToken, text) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + token,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
        'text': text,
      }],
    }),
  });
  
}

 // 連結HTML檔案            
function doGet(){

  var x=HtmlService.createTemplateFromFile("Index");
  var y= x.evaluate();
  //  避免無法顯示在網際網路上
  var z = y.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return z;
}
//取得日期
function getDate(){
  dataValues = [
    date.getFullYear(),
    date.getMonth() + 1,
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
];
  get_date = dataValues[0]+'/'+dataValues[1]+'/'+dataValues[2]+'訂餐統計'
  console.log(get_date)
return get_date
}

// 抓試算表裡的資料
function getData(){
  var Sheet_origin_data=Sheet_origin.getDataRange();
  var Sheet_order_meal_data = Sheet_order_meal.getDataRange();
  return [Sheet_order_meal_data.getValues(),Sheet_origin_data.getValues()];

}

function send_email(){
  let emailAddress = Sheet_parameter.getRange(9, 3).getValue();
  let subject = getDate();
  let message = Sheet_result.getRange(1, 1).getValue();
  MailApp.sendEmail(emailAddress, subject, message);

}




//用於google試算表中之資料清除功能
function datacleanall(){
  for(var i= 0; i<100000000;i++ ){
    if (Sheet_origin.getLastRow()==1){
      break
    }
    else{
      Sheet_origin.deleteRow(Sheet_origin.getLastRow())
    }
  }
  for(var i= 0; i<100000000;i++ ){
    if (Sheet_order_meal.getLastRow()==2){
      break
    }
    else{
      Sheet_order_meal.deleteRow(Sheet_order_meal.getLastRow())
    }
  }
  Sheet_result.getRange(1,1).setValue('無資料')
  Sheet_parameter.getRange(2,1).setValue('off')
  Sheet_parameter.getRange(2,2).setValue('off')
    
}
//建立google sheets的功能
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("資料庫功能")
      .addItem("資料清除","datacleanall")
      .addToUi();
}