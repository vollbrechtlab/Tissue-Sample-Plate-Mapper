/**
 * File name   : ProjectUtilities.gs
 * Created on  : 6/7/2017
 * Author      : Takao Shibamoto
 * Description : Functions that extend API
 */

// fix the user data if it is broken
function fixUserData(userData)
{
  /* Fix project name */
  if(typeof(userData["Project Name"]) != "string" || userData["Project Name"] === ""){
    userData["Project Name"] = "UnknownProj";
  }

  /* fix station/site */
  if(typeof(userData["Station/Site"]) != "string" || userData["Station/Site"] === ""){
    userData["Station/Site"] = "UnknownSite";
  }

  /* fix sample date */
  if(typeof(userData["Sample Date"]) != "string" || userData["Sample Date"] === ""){
    userData["Sample Date"] = "UnknownDate";
  } else {
    userData["Sample Date"] = isoDateToUsDate(userData["Sample Date"]);
  }

  /* fix crop */
  if(typeof(userData["Crop"]) != "string" || userData["Crop"] === ""){
    userData["Crop"] = "UnknownCrop";
  }

  /* fix rack format */
  if(typeof(userData["Rack Format"]) != "string" || userData["Rack Format"] === ""){
    userData["Rack Format"] = "96-all (d)";
  }

  /* fix shading option */
  if(typeof(userData["Shading Option"]) != "string"){
    userData["Shading Option"] = "No";
  }

  /* fix First plate number */
  // check if it is string
  if(typeof(userData["First plate number"]) != "string"){
    userData["First plate number"] = 1;
  } else {
    // change page number to integer 
    userData["First plate number"] = parseInt(userData["First plate number"], 10);
    // if it is not a number or it is not a valid number
    if(isNaN(userData["First plate number"]) || userData["First plate number"] <= 0){
      userData["First plate number"] = 1;
    }
  }

  return userData;
}

// calculate the total number of samples and racks
function countSamplesAndRacks(orgDataArr)
{
  var orgDataInfo = {numSamples: 0, numRacks: 0};

  // calculate the number of total samples
  for(var i = 0; i < orgDataArr.length; i++){
    orgDataInfo.numSamples += orgDataArr[i]["Plant #"];
  }
  // calculate the number of total racks
  orgDataInfo.numRacks = Math.ceil(orgDataInfo.numSamples/96);

  orgDataInfo.numRows = orgDataInfo.numRacks * 96;;

  return orgDataInfo;
}


// read original data sheet
function readOrgData(ss, userData)
{
  var sheet = ss.getSheetByName("Original Data");
  var sheetData = sheet.getDataRange().getValues();
  
  // check if there are columns called "Plot Name/#" and "Plant #"
  var plotNameCol, plantNumCol, noteCol;
  for(var i = 0; i < sheetData[0].length; i++){
    if(sheetData[0][i] == "Plot Name/#"){
      plotNameCol = i;
    }
    if(sheetData[0][i] == "Plant #"){
      plantNumCol = i;
    }
    if(sheetData[0][i] == "Note (optional)"){
      noteCol = i;
    }
  }
  if(plotNameCol == undefined || plantNumCol == undefined){
    // there is no column called Plant#/Shade
    return null;
  }

  // start reading the data
  var orgDataArr = [];
  for(var i = 0; i < sheetData.length-1; i++){
    orgDataArr[i] = {};
    orgDataArr[i]["Plot Name/#"] = sheetData[i+1][plotNameCol];
    orgDataArr[i]["Plant #"] = sheetData[i+1][plantNumCol];

    if(sheetData[i+1][noteCol] == undefined){
      orgDataArr[i]["Note"] = " ";
    } else {
      orgDataArr[i]["Note"] = sheetData[i+1][noteCol];
    }
  }

  // no data
  if(orgDataArr.length == 0){
    return null;
  }

  return orgDataArr;
}


// check if there is any output sheets
function checkOutputSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rackLabelsSheet = ss.getSheetByName("Rack Labels");
  var rackLabelsForPrintingSheet = ss.getSheetByName("Rack Labels for printing");
  var sampleListSheet = ss.getSheetByName("Sample List");
  var summarySheet = ss.getSheetByName("Summary");
  if(rackLabelsSheet != null || rackLabelsForPrintingSheet != null || sampleListSheet != null || summarySheet != null){
    var ui = SpreadsheetApp.getUi(); 
    var result = ui.alert(
      'Output sheets already exist',
      'Overwrite them (Yes) \nStart a new spreadsheet manually (No)',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
      return true;
    } 

    return false;
  }

  return true;
}

// check if the original data sheet exists
// if it does not, create one
function checkOrgDataSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orgDataSheet = ss.getSheetByName("Original Data");
  var ui = SpreadsheetApp.getUi(); 
 
  // if the original data sheet does not exist
  if (orgDataSheet == null) {
    var result = ui.alert(
      'Original data sheet does not exist',
      'A sheet named Original Data does not exist. \nDo you want to create one? (Yes) \nOr please create one manually. (No)',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
      // Add original data sheet with the nessesary column names
      newSheet = ss.insertSheet("Original Data");
      data = []
      data.push(['Plot Name/#', 'Plant #', 'Note (optional)']);
      newSheet.getRange(1, 1, 1, 3).setValues(data);
      newSheet.getRange(1, 1, 1, 1).setBackground('Yellow');
      newSheet.getRange(1, 1, 1, 1).setNote('Experiment/Entry, Rng/Row, Numerical Entry, etc.');
      newSheet.getRange(1, 2, 1, 1).setBackground('Lime');
      newSheet.getRange(1, 2, 1, 1).setNote('Number of plants in plot (list to be expanded), or number of specific plant within the plot (list not to be expanded).  Column Must Contain a Number.');
      newSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      ss.setActiveSheet(newSheet);
    } else {
      // do nothing
    }
    return false;
  } 

  // it does exist 
  var values = orgDataSheet.getDataRange().getValues();
  var found1 = false;
  var found2 = false;
  // check if the column names exist
  checking_loop:
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] == 'Plot Name/#'){
        found1 = true;
      } else if(values[i][j] == 'Plant #'){
        found2 = true;
      }
      if(found1 && found2){
        break checking_loop;
      }
    }
  }
  if(!found1 || !found2){
    var result = ui.alert(
      'Column names are not set',
      '\"Plot Name/#\"" or \"Plant #\"" columns do not exist. \nDo you want to create them? (Yes) \nOr please add them manually. (No)',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
      // Add original data sheet with the nessesary column names
      data = []
      data.push(['Plot Name/#', 'Plant #']);
      orgDataSheet.getRange(1, 1, 1, 2).setValues(data);
      orgDataSheet.getRange(1, 1, 1, 1).setBackground('Yellow');
      orgDataSheet.getRange(1, 1, 1, 1).setNote('Experiment/Entry, Rng/Row, Numerical Entry, etc.');
      orgDataSheet.getRange(1, 2, 1, 1).setBackground('Lime');
      orgDataSheet.getRange(1, 2, 1, 1).setNote('Number of plants in plot (list to be expanded), or number of specific plant within the plot (list not to be expanded).  Column Must Contain a Number.');
      ss.setActiveSheet(orgDataSheet);
    } else {
      // do nothing
    }

    return false;
  }
  return true;
}

function finalInstructionPopup(){
  var ui = SpreadsheetApp.getUi(); 
  var result = ui.alert(
    'Done!',
    'Go to \"Plate Labels for printing\" sheet.\nPrint the sheet with the following settings.\nPrint: Current sheet, \nPage orientation: Landscape, \nScale: 100%, \nMargins: Normal, \nFormatting: No gridlines, \nHorizontal Alignment: Center, \nVertical Alignment: Center\n(Some browsers save PDF instead of printing)',
    ui.ButtonSet.OK);

  // Process the user's response.
  if (result == ui.Button.OK) {
    return true;
  } 

  return false;
}

