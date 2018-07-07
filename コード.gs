
function myFunction() {
  
  var key = 'AIzaSyDp4itNxFAkTHMCpWLNdvRgcYnAb8_NGQM';
  var start = new Date();
  
  //シートを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //行、列を定義
  var rowStartData = 2;
  var rowEndData = sheet.getDataRange().getLastRow();
  var colEndData = sheet.getDataRange().getLastColumn();
  var colTitle = 1;
  var colChannelUrl = 2;
  var colTodaysData = colEndData + 1;
  
  //日付を挿入
  var today = new Date()
  var formatDate = Utilities.formatDate(new Date(),"JST","yy/MM/dd HH:mm");
  sheet.getRange(1,colTodaysData).setValue(formatDate);
  
  
  //for文で繰り返し処理
  for (i = rowStartData; i <= rowEndData; i++) {
    
    
    //登録者数を取得
    var channelUrl = sheet.getRange(i,colChannelUrl).getValue();
    var channelID = channelUrl.slice(32);
    var getSubscUrl = 'https://www.googleapis.com/youtube/v3/channels?part=statistics&id=' + channelID +'&key=' + key;
    var responseSubsc = UrlFetchApp.fetch(getSubscUrl)
    var jsonSubsc = JSON.parse(responseSubsc.getContentText()).items[0].statistics.subscriberCount;

    
    //チャンネルタイトルを取得
    var getTitleUrl = 'https://www.googleapis.com/youtube/v3/channels?part=snippet&id=' + channelID +'&key=' + key ;
    var responseTitle = UrlFetchApp.fetch(getTitleUrl)
    var jsonTitle = JSON.parse(responseTitle.getContentText()).items[0].snippet.title;
    
    
    //登録者数をシートに挿入
    sheet.getRange(i,colTodaysData).setValue(jsonSubsc);
    
    //チャンネルタイトルをシートに挿入
    sheet.getRange(i,colTitle).setValue(jsonTitle);　
  }
  
    var end = new Date();
  var time_past = (end - start)/1000;
  Logger.log(time_past);
  
}