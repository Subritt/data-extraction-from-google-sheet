/*
############################################################################################################################################################################################################
############################################################################### Creator : Subritt Burlakoti        #########################################################################################
############################################################################### Email : subrittburlakoti@gmail.com #########################################################################################
############################################################################################################################################################################################################
*/

function setFormula(throughput){
  /*
  sheet declaration
  */
  //  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //  var throughput = ss.getSheetByName('Throughput');
  
  /*
  getting last row for appending formula
  */
  var data = throughput.getRange(2, 1, throughput.getLastRow()-1, 8).getValues();
  Logger.log(data.length);
  var flag = 0;
  for(var i = 0 ; i < data.length; i++){
    var value = data[i][7];
    var date = data[i][0];
    if(date != "" && value == ""){
      Logger.log(value);
      var flag = i+1;
      break;
    }
    flag = i+2;
  }
  Logger.log(i + " | " + flag);
  
  /*
  Looping Formulas
  */
  var range = throughput.getLastRow()-flag;
  Logger.log(range);
  if(range == 0){
    return;
  }
  
  var arrFormula = [];
  for(var j = 1 ; j <= range ; j++){
    var taskPerHour = "=IFERROR(E"+(flag+j)+"/F"+(flag+j)+',"")';
    var rampPercentage = "=iferror(filter('VC Calculation'!AA:AA,'VC Calculation'!Y:Y=B"+(flag+j)+",'VC Calculation'!Z:Z=C"+(flag+j)+"),"+'""'+")";
    var minTarget = "=iferror(filter('Daily Targets Records'!C:D,'Daily Targets Records'!B:B=D"+(flag+j)+",'Daily Targets Records'!A:A=A"+(flag+j)+"),"+'""'+")";
    var minTargetAfterRamp = "=I"+(flag+j)+"%*J"+(flag+j);
    var maxTargetAfterRamp = "=I"+(flag+j)+"%*K"+(flag+j);
    //    arrFormula = [[taskPerHour],[rampPercentage],[minTarget],[minTargetAfterRamp],[maxTargetAfterRamp]];
    arrFormula.push([taskPerHour,rampPercentage,minTarget,,minTargetAfterRamp,maxTargetAfterRamp]);
  }
  Logger.log(arrFormula);
  
  throughput.getRange(flag+1, 8, arrFormula.length, arrFormula[0].length).setValues(arrFormula);
  throughput.getRange(flag+1, 11, arrFormula.length).clearContent();
  return;
  
  Logger.log(throughput.getRange(flag+1, 8, arrFormula.length, arrFormula[0].length).getValues());
  Utilities.sleep(5000);
  throughput.getRange(flag+1, 8, arrFormula.length, arrFormula[0].length).activate();
  throughput.getRange(flag+1, 8, arrFormula.length, arrFormula[0].length).copyTo(throughput.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
//  Utilities.sleep(5000);
  
  //  DataChange();
  
  return;
}