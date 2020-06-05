/*
############################################################################################################################################################################################################
############################################################################### Creator : Subritt Burlakoti        #########################################################################################
############################################################################### Email : subrittburlakoti@gmail.com #########################################################################################
############################################################################################################################################################################################################
*/

// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

// Test if value is a date and if so format
// otherwise, reflect input variable back as-is. 
function isDate(sDate) {
  if (isValidDate(sDate)) {
    sDate = Utilities.formatDate(new Date(sDate), "GMT+5:45", "MM/dd/yyyy");
  }
  return sDate;
}

/*
Parent FUnction
*/
function main(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VC Implemented');
  
  /*
  Fetching data
  */
  var driveID = sheet.getRange('B1').getValues();
  var data = sheet.getRange(4, 1, sheet.getLastRow()-3, sheet.getLastColumn()).getValues();
  Logger.log(data[0][14]);
  
  //  for(var j = 0 ; j < data.length ; j++){
  //    var wsName = data[j][0];
  //    var wsUrl = data[j][7];
  //    var throughputUrl = data[j][3];
  //    var accuracyUrl = data[j][5];
  //    //    Logger.log(wsName + " | " + wsUrl + " | " + throughputUrl + " | " + accuracyUrl);
  //    Logger.log(j + " : " + data[j][10]);
  //    if(wsName == "" || data[j][10] == "Done"){
  //      Logger.log("Condition met");
  //      continue;
  //    }
  //    Logger.log("Condition did not meet!");
  //  }
  //  return;
  
  var drive = DriveApp.getFolderById(driveID);
  var name;
  var folderId;
  var flag = [];
  for(var j = 0 ; j < data.length ; j++){
    var wsName = data[j][0];
    var wsUrl = data[j][7];
    var throughputUrl = data[j][3];
    var accuracyUrl = data[j][5];
    Logger.log(wsName + " | " + wsUrl + " | " + throughputUrl + " | " + accuracyUrl);
    Logger.log(j + " : " + data[j][10]);
    
    ss.toast("Working on : " + data[j][0]);
    
    if(data[j][14] == 'No'){
      Logger.log("Validation : " + data[j][14]);
      ss.toast("Skipped!!!" + "Validation : " + data[j][14]);
      continue;
    }
    
    if(wsName == "" || data[j][10] == "Done"){
      Logger.log("Condition met!");
      ss.toast("Skipped!!! : " + "Status ---> Done")
      continue;
    }
    
    if(wsUrl != ""){
      appendHistoricalData(wsName, wsUrl, throughputUrl, accuracyUrl);
    }else{
      var vcFolder = "VC of "+data[j][0];
      var folders = drive.getFolders();
      while(folders.hasNext()){
        var folder = folders.next();
        if(folder.getName() == vcFolder){
          Logger.log(folder.getName()+" || "+folder.getId());
          name = data[j][0]+" VC Data";
          folderId = folder.getId();
          var fileUrl = createSheet(name, folderId);
          sheet.getRange(j+4, 8).setValue(fileUrl);
          appendHistoricalData(wsName, fileUrl, throughputUrl, accuracyUrl);
          //          return;
        }
      }
    }
    sheet.getRange(j+4, 11).setValue("Done");
  }
  
}

/*
Function to create a spreadsheet, 
insert "Throughput", "Accuracy", "Overall Throughput", "Overall Accuracy"
and add headers to these sheets
*/
function createSheet(name, folderId){
  var throughputHeader = [['Date (MM/DD/YYYY)','Cloud Worker','Use Case','Task Type','Tasks Completed','Task Duration',
                           'Notes','Average Task/Hr','Ramp Percentage','Target (Min) (Actual)','Target (Max) (Actual)',
                           'Target (Min) (After Ramp)','Target (Max) (After Ramp)','VC%']];
  var accuracyHeader = [['Date','#Task Reviewed','Task_Url (Optional)','Task Completed By','Task Type',
                         '# Review Duration (hours)','Task Reviewed By','Accuracy','Min','Max','VC']];
  var overallThroughputHeader = [['Date','Email Address','Average VC %']];
  var overallAccuracyHeader = [['Date','Email Address','Average VC %']];
  var resource = {
    title : name,
    mimeType : MimeType.GOOGLE_SHEETS,
    parents : [{ id : folderId}]
  };
  var fileJson = Drive.Files.insert(resource);
  var file = SpreadsheetApp.openById(fileJson.id);
  file.getActiveSheet().setName("Throughput");
  file.getActiveSheet().getRange(1, 1, 1, throughputHeader[0].length).setValues(throughputHeader).setFontWeight('bold');
  //  file.getActiveSheet().getRange(1, file.getActiveSheet().getLastColumn()+1, 1, file.getActiveSheet().getMaxColumns())
  file.insertSheet().setName("Accuracy");
  file.getActiveSheet().getRange(1, 1, 1, accuracyHeader[0].length).setValues(accuracyHeader).setFontWeight('bold');
  file.insertSheet().setName("Overall Throughput");
  file.getActiveSheet().getRange(1, 1, 1, overallThroughputHeader[0].length).setValues(overallThroughputHeader).setFontWeight('bold');
  file.insertSheet().setName("Overall Accuracy");
  file.getActiveSheet().getRange(1, 1, 1, overallAccuracyHeader[0].length).setValues(overallAccuracyHeader).setFontWeight('bold');
  return file.getUrl();
}

/*
Function to append historical data for
"Throughput", "Accuracy", "Overall Throughput", "Overall Accuracy"
*/
function appendHistoricalData(wsName, wsUrl, throughputUrl, accuracyUrl){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VC Implemented');
  
  Logger.log("Throughput : " + throughputUrl + " | " + "Accuracy : " + accuracyUrl);
  
  /*
  Extracting Data
  */
  var throughputSS = SpreadsheetApp.openByUrl(throughputUrl);
  var throughput = throughputSS.getSheetByName('Throughput');
  var overallThroughput = throughputSS.getSheetByName('Overall');
  var accuracySS = SpreadsheetApp.openByUrl(accuracyUrl);
  var accuracy = accuracySS.getSheetByName('Accuracy');
  var overallAccuracy = accuracySS.getSheetByName('Overall Accuracy');
  
  //  var file = DriveApp.getFileById(throughputSS.getId());
  //  Logger.log(file);
  //  var driveID = file.getParents().next().getId();
  //  Logger.log(driveID);
  
  var throughputData = throughput.getRange(2, 1, throughput.getLastRow()-1, throughput.getLastColumn()).getValues();
  var formattedThroughput = dateFormate(throughputData);
  
  var overallThroughputData = overallThroughput.getRange(2, 4, overallThroughput.getLastRow()-1, overallThroughput.getLastColumn()).getValues();
  var formattedOverallThroughput = dateFormate(overallThroughputData);
  
  var accuracyData = accuracy.getRange(2, 1, accuracy.getLastRow()-1, accuracy.getLastColumn()).getValues();
  var formattedAccuracy = dateFormate(accuracyData);
  
  var overallAccuracyData = overallAccuracy.getRange(2, 4, overallAccuracy.getLastRow()-1, overallAccuracy.getLastColumn()).getValues();
  var formattedOverallAccuracy = dateFormate(overallAccuracyData);
  
  /*
  pushing Historical Values
  */    
  var vcDataSS = SpreadsheetApp.openByUrl(wsUrl);
  var vcDataThroughput = vcDataSS.getSheetByName('Throughput');
  var vcDataAccuracy = vcDataSS.getSheetByName('Accuracy');
  var vcOverallThroughput = vcDataSS.getSheetByName('Overall Throughput');
  var vcOverallAccuracy = vcDataSS.getSheetByName('Overall Accuracy');
  
  try{
    vcDataThroughput.getRange(vcDataThroughput.getLastRow()+1, 1, formattedThroughput.length, formattedThroughput[0].length).setValues(formattedThroughput);
    vcOverallThroughput.getRange(vcOverallThroughput.getLastRow()+1, 1, formattedOverallThroughput.length, formattedOverallThroughput[0].length).setValues(formattedOverallThroughput);
    vcDataAccuracy.getRange(vcDataAccuracy.getLastRow()+1, 1, formattedAccuracy.length, formattedAccuracy[0].length).setValues(formattedAccuracy);
    vcOverallAccuracy.getRange(vcOverallAccuracy.getLastRow()+1, 1, formattedOverallAccuracy.length, formattedOverallAccuracy[0].length).setValues(formattedOverallAccuracy);
  }catch(e){
    Logger.log("Append Function : " + " | " + e);
  }
  
  return;
}

/*
managing date format
*/
function dateFormate(value){
  var arr = [];
  for(var i = 0 ; i < value.length ; i++){
    var date = "";
    if(value[i][0] != ""){
      try{
        date = Utilities.formatDate(value[i][0], "GMT+5:45", "MM/dd/yyyy");
      }catch(e){
        Logger.log("dateFormat function : " + " | " + e);
        continue;
      }
      arr.push([date]);
    }else{
      arr.push([date]);
    }
    for(var j = 1 ; j < value[0].length ; j++){
      try{
        arr[i].push(value[i][j]);
      }catch(e){
        Logger.log("dateFormate() error : " + e);
      }
    }
  }
  return arr;
}