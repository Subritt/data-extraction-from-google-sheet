/*
############################################################################################################################################################################################################
############################################################################### Creator : Subritt Burlakoti                   ##############################################################################
############################################################################### Email : subritt.burlakoti@es.cloudfactory.com ##############################################################################
############################################################################### Delivery Solution Project Support             ##############################################################################
############################################################################################################################################################################################################
*/

/*
function to get session
*/
function getSession(){
  return Session.getEffectiveUser();
}

/*
switch toggle
*/
function switch_toggle(sheet){
  return sheet.getRange('J1').getValue();
}


/*
 ____________________________________
| Function to extract throughput data |
 -------------------------------------
*/
function extractThroughput(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VC Implemented');
  
  /*
  Trigger
  */
  const start_time = new Date().getTime();
  Logger.log(start_time);
  var all_trigger = ScriptApp.getProjectTriggers();
  Logger.log(all_trigger.length);

  //  if(all_trigger.length > 1){
  //    Logger.log(all_trigger[1].getUniqueId());
  //    ScriptApp.deleteTrigger(all_trigger[1]);
  //    Logger.log("Trigger Deleted");
  //  }
  
  /*
  Setting active user
  */
  sheet.getRange('J2').setValue(getSession());
  
  /*
  calling switch toggle
  */
  if(switch_toggle(sheet) == false){
    Logger.log('Switch is turned off');
    ss.toast('Switch is off. Turn on to extract Throughput Data');
    sheet.getRange('J2').clearContent();
    return;
  }
  
  /*
  Fetching data
  */
  var driveID = sheet.getRange('B1').getValues();
  var data = sheet.getRange(4, 1, sheet.getLastRow()-3, sheet.getLastColumn()).getValues();
  Logger.log(data.length);
  
  /*
  Looping arround workstream
  */
  for(var i = 0 ; i < data.length ; i++){
    
    /*
    check validation
    */
    if(data[i][14] == 'No'){
      Logger.log('Validation : ' + data[i][14]);
      ss.toast('Validation : ' + data[i][14]);
      continue;
    }
    
    var error_count = 0;
    
    /*
    Script duration
    */
    var current_time = new Date().getTime();
    var elapsed_time = current_time - start_time;
    
    Logger.log(elapsed_time/60000 + "minutes");
    
    /*
    Create Trigger
    28 minutes = 1680000 miliseconds
    20 minutes = 1200000 miliseconds
    */
    if(elapsed_time > 1200000){
      //      ScriptApp.newTrigger('extractThroughput')
      //      .timeBased()
      //      .everyMinutes(30)
      //      .create();
      Logger.log("Trigger Created!");
      sheet.getRange('J2').clearContent();
      return;
    }
    
    var internalUrl = data[i][1];
    var throughputVCurl = data[i][3];
    Logger.log(internalUrl + " | " + throughputVCurl);
    
    var sourceSS = SpreadsheetApp.openByUrl(internalUrl);
    var sourceSheet = sourceSS.getSheetByName("Throughput");
    var wsSetting = sourceSS.getSheetByName("WorkStream Setting");
    var targetSS = SpreadsheetApp.openByUrl(throughputVCurl);
    var targetSheet = targetSS.getSheetByName("Throughput");
    
    /*
    Looping through data
    */
    if(data[i][0] == "" || (new Date().getDate()) == (new Date(data[i][8]).getDate())){
      Logger.log("Loop Continued");
      ss.toast('SKIPPED!!!' + 'Workstream : ' + data[i][0] + 'Updated Date : ' + Utilities.formatDate(new Date(data[i][8]), "GMT+5:45", "MM/dd/yyyy"));
      continue;
    }
    Logger.log("after checking date == null");
    ss.toast("Working On : " + data[i][0]);
    /*
    clear status
    */
    sheet.getRange(i+4, 10).clearContent();
    
    //    if(data[i][8] != ""){
    //      if(new Date(new Date().setDate(new Date().getDate()-1)) != new Date(data[i][8]) && data[i][8] != ""){
    //        if(new Date(new Date().setDate(new Date().getDate()-1)) > new Date(targetSheet.getRange(targetSheet.getLastRow(),1).getValue())){
    //          date = Utilities.formatDate(new Date(targetSheet.getRange(targetSheet.getLastRow(),1).getValue()), "GMT+5:45", "MM/dd/yyyy");
    //        }else{
    //          continue;
    //        }
    //      }
    //    }
    var throughputEndData = targetSheet.getRange(targetSheet.getLastRow(),1).getValue();
    
    if(throughputEndData == "Date (MM/DD/YYYY)" || throughputEndData == ""){
      var today = new Date();
      var newdate = Utilities.formatDate(new Date((today.getMonth())+"/"+1+"/"+today.getYear()), "GMT+5:45", "MM/dd/yyyy");
      var date = new Date(new Date().setDate(new Date(newdate).getDate()-1));
    }else{
      date = Utilities.formatDate(new Date(throughputEndData), "GMT+5:45", "MM/dd/yyyy");
    }
    
    /*
    Fetch Data
    */
    var teamCaptain = wsSetting.getRange("E3:E"+(wsSetting.getLastRow())).getValues();
    var sourceData = sourceSheet.getRange("A2:G"+(sourceSheet.getLastRow())).getValues();
    Logger.log(sourceData[sourceData.length-1]);
    
    var arr = [];
    
    for(var j = 0 ; j < sourceData.length ; j++){
      //      if(sourceData[j][0] == "" || sourceData[j][0] == null){
      //        Logger.log("inside if : " + sourceData[j][0]);
      //        continue;
      //      }
      
      /*
      Condition to check date and user email in Internal Dashboard
      */
      if(sourceData[j][0] == "" && sourceData[j][1] == ""){
        continue;
      }
      
      Logger.log(sourceSS.getName() + " : " + sourceData[j][0]);
      try{
        var sourceDate = Utilities.formatDate(sourceData[j][0], "GMT+5:45", "MM/dd/yyyy");
      }catch(e){
        Logger.log("Date Exception : " + e);
        sheet.getRange(i+4, 10).setValue(data[i][0] + " : "+ "Row Number : "+(j+2)+" ||| "+" Error Message : "+e);
        e="";
        error_count = 1;
      }
      if(error_count == 1){
        break;
      }
      var email = sourceData[j][1];
      
      //      Logger.log(sourceDate);
      var flag = 0;
      for(var k = 0 ; k < teamCaptain.length ; k++){
        if(email == teamCaptain[k][0] || email.split("@")[1] != "es.cloudfactory.com"){
          flag = 1;
          break;
        }
      }
      if((new Date(sourceDate)) > (new Date(date)) && (new Date(sourceDate)) < (new Date(Utilities.formatDate(new Date(), "GMT+5:45", "MM/dd/yyyy")))  && flag == 0){
        Logger.log("inside if conition " + sourceDate);
        arr.push(sourceData[j]);
      }
    }
    
    Logger.log(arr);
    if(arr != "" && error_count == 0){
      targetSheet.getRange(targetSheet.getLastRow()+1, 1, arr.length, arr[0].length).setValues(arr);
      setFormula(targetSheet);
    }
    sheet.getRange("I4:I"+(i+4)).setValue(Utilities.formatDate(new Date(), "GMT+5;45", "MM/dd/yyyy"));
  }
  
  //  if(i == data.length-1){
  //    sheet.getRange('J2').clearContent();
  //  }
  sheet.getRange('J2').clearContent();
  //  return;
  
}