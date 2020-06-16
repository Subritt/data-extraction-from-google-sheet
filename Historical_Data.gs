/*
############################################################################################################################################################################################################
############################################################################### Creator : Subritt Burlakoti        #########################################################################################
############################################################################### Email : subrittburlakoti@gmail.com #########################################################################################
############################################################################################################################################################################################################
*/

function extractHistoricalData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VC Implemented');
  
  /*
  Extract WS
  */
  var ws = sheet.getRange(4, 1, sheet.getLastRow()-3, sheet.getLastColumn()).getValues();
  //  Logger.log(ws);
  
  for(var k = 0 ; k < ws.length ; k++){
    
    ss.toast("Working on : " + ws[k][0]);
    
    if(ws[k][14] == "No"){
      Logger.log("Validation : " + ws[k][14]);
      ss.toast("Skipped!!!" + "Validation : " + ws[k][14]);
      continue;
    }
    if(ws[k][11] == "Extracted"){
      Logger.log("Already Extracted");
      ss.toast("Already Extracted");
      continue;
    }
    
    var warehouse = SpreadsheetApp.openByUrl(ws[k][7]);
    var throughput = warehouse.getSheetByName('Throughput');
    var throughputVCss = SpreadsheetApp.openByUrl(ws[k][3]);
    var throughputSheet = throughputVCss.getSheetByName('Throughput');
    var workstreamSetting = throughputVCss.getSheetByName('WorkStream Setting');
    
    throughputSheet.getRange(2, 1, throughputSheet.getMaxRows(), throughputSheet.getMaxColumns()).clearContent();
    
    /*
    Hours to RampUp
    */
    var rampDuration = workstreamSetting.getRange("K3:K"+(workstreamSetting.getLastRow())).getValues();
    for(var i = 0; i < rampDuration.length ; i++){
      var rampFlag = 0;
      if(rampDuration[i] != "" && parseInt(rampDuration[i]) > 40){
        break;
      }
      //      Logger.log("Hours to RampUp <= 40");
      //      extractThroughput();
      rampFlag = 1;
      break;
    }
    if(rampFlag == 1){
      sheet.getRange(k+4, 12).setValue("Extracted");
      Logger.log("Extracted");
      sheet.getRange(k+4, 9).clearContent();
      continue;
    }
    
    var historicalData = throughput.getRange(2, 1, throughput.getLastRow()-1, throughput.getLastColumn()).getValues();
    //  Logger.log(historicalData.length);
    //  return;
    
    var email = [];
    var usecase = [];
    var tasktype = [];
    
    var flag;
    var count = 0;
    email.push(historicalData[8][1]);
    usecase.push(historicalData[8][2]);
    tasktype.push(historicalData[8][3]);
    for(var i = 0 ; i < historicalData.length ; i++){
      if(historicalData[i][0] == ""){
        continue;
      }
      
      var emailData = historicalData[i][1];
      var usecaseData = historicalData[i][2];
      var tasktypeData = historicalData[i][3];
      
      if(count == 0){
        var finalArray = [[emailData,usecaseData,tasktypeData]];
        count = 1;
      }
      
      //    Logger.log(emailData);
      
      for(var j = 0 ; j < finalArray.length ; j++){
        //      Logger.log(emailData + " | " + email[j]);
        if(emailData == finalArray[j][0] && usecaseData == finalArray[j][1] && tasktypeData == finalArray[j][2]){
          //        Logger.log("Inside IF");
          flag = 1;
          break;
        }
        
        flag = 0; 
      }
      
      if(flag == 0){
        email.push(emailData);
        usecase.push(usecaseData);
        tasktype.push(tasktypeData);
        finalArray.push([emailData,usecaseData,tasktypeData]);
      }
      
    }
    //  Logger.log(finalArray);
    var finalData = totalSum(finalArray, throughput, historicalData);
    
    throughputSheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
    //    extractThroughput();
    sheet.getRange(k+4, 12).setValue("Extracted");
    Logger.log((k+4) + " : " + "Extracted");
    sheet.getRange(k+4, 9).clearContent();
  }
  
}

/*
Function to calculate sum of task duration
*/
function totalSum(arr, throughputSheet, data){
  Logger.log("Before sum");
  Logger.log(arr.length);
  Logger.log(arr);
  var finalData = [];
  for(var i = 0 ; i < arr.length ; i++){
    var task = 0;
    var duration = 0;
    for(var j = 0 ; j < data.length ; j++){
      if(data[j][0] == ""){
        continue;
      }
      if(data[j][1] == arr[i][0] && data[j][2] == arr[i][1] && data[j][3] == arr[i][2]){
        Logger.log(data[j][1] + " ||| " + data[j][2] + " ||| " + data[j][3] + " ||| " + Number(data[j][4]));
        task += Number(data[j][4]);
        duration += Number(data[j][5]);
      }
    }
    arr[i].push(task);
    arr[i].push(duration);
  }
  
  Logger.log("After SUM")
  Logger.log(arr.length);
  Logger.log(arr);
  
  /*
  Appending Data from arr to finalData
  */
  for(var row = 0 ; row < arr.length ; row++){
    finalData.push([""]);
    for(var col = 0 ; col < arr[0].length ; col++){
      finalData[row].push(arr[row][col]);
    }
  }
  Logger.log("Final DATA");
  Logger.log(finalData.length);
  Logger.log(finalData);
  
  return finalData;
}