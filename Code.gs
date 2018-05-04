/**
 * File name   : Code.gs
 * Created on  : 5/19/2017
 * Author      : Takao Shibamoto
 * Description : The main code
 */

/*
 * Structures of data used many times
 * 
 * userData {Dictionary}
 * userData["Project Name"] {String}
 * userData["Station/Site"] {String}
 * userData["Sample Date"] {String}
 * userData["Crop"] {String} 
 * userData["Rack Format"] {String}
 * userData["Shading Option"] {String}
 * userData["First plate number"] {String}
 *
 * orgDataArr {Array}
 * orgDataArr[i] {Dictionary}
 * orgDataArr[i]["Plot Name/#"] {String}
 * orgDataArr[i]["Plant #"] {Number}
 * orgDataArr[i]["Note"] {String}
 */

function onOpen(e) {
  SpreadsheetApp.getUi() 
    .createAddonMenu()
    .addItem('Generate labels', 'showSidebar')
    .addToUi();  
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Label Generation')
    .setWidth(300);
  SpreadsheetApp.getUi() 
    .showSidebar(html);

  if(!checkOrgDataSheet()){
    // if the original data is not set properly,
    // dont do anything
    return;
  }
}

/**
 * start generating labels
 * @param {Dictionary} user data from the form
 */
function startGeneration(userData) 
{ 
  if(!checkOrgDataSheet()){
    // if the original data is not set properly,
    // dont do anything
    return;
  }
  if(!checkOutputSheets()){
    // check if user wants to override the existing ouputs or not
    return;
  }

  // get the current spreasheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  userData = fixUserData(userData);

  // read the original data and store it in a dictionary
  var orgDataArr = readOrgData(ss, userData);
  if(orgDataArr == null){
    ui = SpreadsheetApp.getUi();
    ui.alert(
      'Data is empty',
      'Please add some data to the original data sheet',
      ui.ButtonSet.OK);
    return;
  }

  // calculate the number of samples and racks
  var orgDataInfo = countSamplesAndRacks(orgDataArr);
  
  // create a Plate Labels sheet
  createRackLabelsSheet(ss, userData, orgDataArr, orgDataInfo);

  // create a Plate Labels for printing sheet
  createRackLabelsForPrintingSheet(ss, userData, orgDataInfo);

  // create a sample list sheet
  createSampleListSheet(ss, userData, orgDataArr, orgDataInfo);

  // create a summary sheet
  createSummarySheet(ss, userData, orgDataArr, orgDataInfo);

  // change the currently active sheet to the summary sheet
  // ss.setActiveSheet(ss.getSheetByName("Summary"));

  //finalInstructionPopup();

}

/**
 * creates a summary sheet based on the user data and original sheet
 * @param {SpreadSheet} spreadsheet
 * @param {Dictionary} user data from the form
 * @param {Array} original data from the sheet as an array
 * @param {Array} extra info from the original data from the sheet
 */
function createSummarySheet(ss, userData, orgDataArr, orgDataInfo) 
{
  /* create the summary sheet */
  var sheet = createNewSheet(ss, "Summary");

  /* array of basic summary data that will be copied on the sheet */
  var summaryData = [];
  summaryData.push(["Project Name", userData["Project Name"]]);
  summaryData.push(["Process Data", getCurrentDate()]);
  summaryData.push(["Sample Date", userData["Sample Date"]]);
  summaryData.push(["Station/Site", userData["Station/Site"]]);
  summaryData.push(["Crop", userData["Crop"]]);
  summaryData.push(["Plate Format", userData["Rack Format"]]);
  summaryData.push(["Total Plates", orgDataInfo.numRacks]);
  summaryData.push(["Total Samples", orgDataInfo.numSamples]);
  summaryData.push(["Comments", " "]);
  summaryData.push([" ", " "]); // empty row
  summaryData.push(["Plate ID", "Plots"]);  // rack IDs to the sheet  

  /* create the summary of each plate */
  var currPlateIdx = 1;
  var remainingCellSpace = 96;
  for(var i = 0; i < orgDataArr.length; i++)
  {
    // check if the there is enough space for the plot
    if(remainingCellSpace / orgDataArr[i]["Plant #"] >= 1){
      summaryData.push([userData["Project Name"]+"_"+currPlateIdx ,orgDataArr[i]["Plot Name/#"] + " from 1" + " to " + orgDataArr[i]["Plant #"]]);
      remainingCellSpace -= orgDataArr[i]["Plant #"];
    } 
    else { // if there is not enough space, go to the next plate and update remaining cell space
      summaryData.push([userData["Project Name"]+"_"+currPlateIdx, orgDataArr[i]["Plot Name/#"] + " from 1" + " to " + remainingCellSpace]);
      currPlateIdx++;
      summaryData.push([userData["Project Name"]+"_"+currPlateIdx, orgDataArr[i]["Plot Name/#"] + " from " + (remainingCellSpace+1) + " to " + orgDataArr[i]["Plant #"]]);
      remainingCellSpace = remainingCellSpace - orgDataArr[i]["Plant #"] + 96;
    }
  }

  /* copy the data array on the sheet */
  sheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);

  /* format the sheet */
  sheet.getRange('A1:A9').setFontWeight('bold');
  sheet.getRange('A11:B11').setFontWeight('bold');
  sheet.getRange(1, 1, summaryData.length, 2).setHorizontalAlignment("center");
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
}

/**
 * creates a Plate Labels sheet based on the user data and original sheet
 * @param {SpreadSheet} spreadsheet
 * @param {Dictionary} user data from the form
 * @param {Array} original data from the sheet as an array
 * @param {Array} extra info from the original data from the sheet
 */
function createRackLabelsSheet(ss, userData, orgDataArr, orgDataInfo)
{
  /* create a Plate Labels sheet */
  var sheet = createNewSheet(ss, "Plate Labels");

  /* create a list of samples in the "Plot Name/# \n Plant #" format
     it also store if the cell should be shaded or not */
  var samples = [];
  var shadeToggle = true; // used for shading every other full
  for(var i = 0; i < orgDataArr.length; i++){
    for(var j = 1; j <= orgDataArr[i]["Plant #"]; j++)
    {   
      // Shading option: Shade First Plant of Every Entry is selected
      if(userData["Shading Option"] == "First Plant") {
        if(j == 1){
          samples.push({name:orgDataArr[i]["Plot Name/#"] + "\n" + j, isShaded:true});
        } else {
          samples.push({name:orgDataArr[i]["Plot Name/#"] + "\n" + j, isShaded:false});
        }
      } 
      // Shading option: Shade Every Other Full Entry is selected
      else if(userData["Shading Option"] == "Every Other Full") {
        if(j == 1){
          shadeToggle = !shadeToggle;
        }
        samples.push({name:orgDataArr[i]["Plot Name/#"] + "\n" + j, isShaded:shadeToggle});
      } 
      // No shading
      else {
        samples.push({name:orgDataArr[i]["Plot Name/#"] + "\n" + j, isShaded:false});
      }
    }
  }


  /* create and manipulate the 2d array that will be written to the sheet*/

  var currSampleIdx = 0; // index of the current sample
  var shadedPositions = []; // array of positions that will be shaded
  for (var i = 0; i < orgDataInfo.numRacks; i++)
  {
    // create a 2d array containing data in each label
    // this array will be written on the sheet
    var rackData = create2dArray(8, 12);
    
    if(userData["Rack Format"] == "96-All (d)") {
      // put the samples vertically to the data array
      for (var col = 0; col < 12; col++) {
        for (var row = 0; row < 8; row++) {
          if(currSampleIdx < samples.length){
            if(samples[currSampleIdx].isShaded == true){
              shadedPositions.push({row:row+2+10*i, col:col+3});
            }
            rackData[row][col] = samples[currSampleIdx].name;
            currSampleIdx++;
          } else {
            rackData[row][col] = " ";
          }
        }
      }
    } else if (userData["Rack Format"] == "96-All (a)") {
      // put the samples horizontally to the data array
      for (var row = 0; row < 8; row++) {
        for (var col = 0; col < 12; col++) {
          if(currSampleIdx < samples.length){
            if(samples[currSampleIdx].isShaded == true){
              shadedPositions.push({row:row+2+10*i, col:col+3});
            }
            rackData[row][col] = samples[currSampleIdx].name;
            currSampleIdx++;
          } else {
            rackData[row][col] = " ";
          }
        }
      }
    }

    // add column names and ranges to the data array
    rackData.splice(0, 0, ["1","2","3","4","5","6","7","8","9","10","11","12"]);
    rackData.splice(9, 0, ["1/8","9/16","17/24","25/32","33/40","41/48","49/56","57/64","65/72","73/80","81/88","89/96"]);
    
    // add row names and descriptions to the data array
    var rowNames = ["X", "A", "B", "C", "D", "E", "F", "G", "H", "X"];
    var descriptions = [" ", userData["Project Name"]+"-"+(i+1), 
                        "[ "+(userData["First plate number"]+i)+" ]",  
                        userData["Sample Date"], " ", " ", " ", " ", " ", " "];
    for (var row = 0; row < 10; row++) {
      // add the row names to the rack data
      rackData[row].splice(0, 0, rowNames[row]);
      // add the descriptions to the rack data
      rackData[row].splice(0, 0, descriptions[row]);
      rackData[row].splice(14, 0, descriptions[row]);
      rackData[row].splice(14, 0, " ");
    }

    /* write the rack data to the sheet */
    sheet.getRange(10*i+1, 1, 10, 16).setValues(rackData);


    /* format each rack label */
    // format column names
    sheet.getRange(10*i+1, 2, 1, 13).setFontWeight('bold');
    sheet.getRange(10*i+1, 1, 1, 16).setBackground('#c0c0c0');
    // format the sample ranges
    sheet.getRange(10*i+10, 1, 1, 16).setBackground('#c0c0c0');
    // set row heights
    sheet.setRowHeight(10*i+1, 15);
    sheet.setRowHeight(10*i+10, 15);
    for(var row = 10*i+2; row <= 10*i+9; row++){
      sheet.setRowHeight(row, 34);
    }
    // change font size
    sheet.getRange(10*i+2, 3, 9, 12).setFontSize(7);
    // set borders
    sheet.getRange(10*i+2, 3, 8, 12)
      .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(10*i+2, 3, 8, 12)
      .setBorder(null, null, null, null, true, true, "black", SpreadsheetApp.BorderStyle.DOTTED);
    sheet.getRange(10*i+2, 2, 8, 1)
      .setBorder(true, true, true, null, null, true, "black", SpreadsheetApp.BorderStyle.DOTTED);
    sheet.getRange(10*i+2, 15, 8, 1)
      .setBorder(true, null, true, true, null, null, "black", SpreadsheetApp.BorderStyle.DOTTED);
    sheet.getRange(10*i+1, 3, 1, 12)
      .setBorder(true, true, null, true, true, null, "black", SpreadsheetApp.BorderStyle.DOTTED);
    sheet.getRange(10*i+10, 3, 1, 12)
      .setBorder(null, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.DOTTED);

  }

  // shade the cells
  for(var i = 0; i < shadedPositions.length; i++){
    sheet.getRange(shadedPositions[i].row, shadedPositions[i].col, 1, 1).setBackground('#c0c0c0');
  }


  /* format the sheet */

  // set column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(16, 100);
  for(var col = 2; col <= 15; col++){
    sheet.setColumnWidth(col, 33);
  }
  // left and right edges
  sheet.getRange(1, 2, 10*orgDataInfo.numRacks, 13).setHorizontalAlignment("center");
  sheet.getRange(1, 2, 10*orgDataInfo.numRacks, 13).setVerticalAlignment("middle");
  // left edge
  sheet.getRange(1, 1, 10*orgDataInfo.numRacks, 1).setHorizontalAlignment("left");
  sheet.getRange(1, 1, 10*orgDataInfo.numRacks, 1).setFontWeight('bold');
  // right edge
  sheet.getRange(1, 16, 10*orgDataInfo.numRacks, 1).setHorizontalAlignment("right");
  sheet.getRange(1, 16, 10*orgDataInfo.numRacks, 1).setFontWeight('bold');
  // row names
  sheet.getRange(1, 2, 10*orgDataInfo.numRacks, 1).setBackground('#c0c0c0');
  sheet.getRange(1, 2, 10*orgDataInfo.numRacks, 1).setFontWeight('bold');
  // right blank
  sheet.getRange(1, 15, 10*orgDataInfo.numRacks, 1).setBackground('#c0c0c0');

  // set wrap
  //cells = sheet.getRange(1, 1, 10*orgDataInfo.numRacks, 1);
  //cells.setWrap(true);
  //cells = sheet.getRange(1, 16, 10*orgDataInfo.numRacks, 1);
  //cells.setWrap(true);

}

/**
 * creates a Plate Labels sheet for the printing purpose
 * use this method only after createRackLabelsSheet()
 * @param {SpreadSheet} spreadsheet
 * @param {Dictionary} user data from the form
 * @param {Array} original data from the sheet as an array
 */
function createRackLabelsForPrintingSheet(ss, userData, orgDataInfo)
{
  // create the new sheet
  var sheet = copySheet(ss, ss.getSheetByName("Plate Labels"), "Plate Labels for printing");

  // insert rows, clear the format, and change the row height 
  for(var i = 0; i < orgDataInfo.numRacks; i++){
    sheet.insertRowBefore(12*i+1); // insert the top row
    sheet.getRange(12*i+1, 1, 1, 16).clearFormat();
    sheet.setRowHeight(12*i+1, 160); 
    sheet.insertRowAfter(12*i+11); // insert the bottom row
    sheet.getRange(12*i+12, 1, 1, 16).clearFormat();
    sheet.setRowHeight(12*i+12, 160);
    //var currPlateNum = sheet.getRange(12*i+4, 1, 1, 1).getValue();
    var currPlateNum = userData["First plate number"] + i;
    sheet.getRange(12*i+12, 1, 1, 1).setValue(currPlateNum); // write the page number
    sheet.getRange(12*i+12, 1, 1, 16).merge(); // merge columns
    sheet.getRange(12*i+12, 1, 1, 16).setVerticalAlignment('middle');
    sheet.getRange(12*i+12, 1, 1, 16).setHorizontalAlignment('center');
    sheet.getRange(12*i+12, 1, 1, 16).setFontSize(30);
    sheet.getRange(12*i+1, 1, 1, 16).merge(); // merge columns

  }
  /*
  // create a space to add a note
  sheet.getRange(1, 17, 12, 8).merge();
  // add printing instruction as a note
  sheet.getRange(1, 17, 12, 8)
    .setNote( "Print this sheet with the following settings.\n" +
              "Print: Current sheet, \n" +
              "Page orientation: Landscape, \n" +
              "Scale: 100%, \n" +
              "Margins: Normal, \n" +
              "Formatting: No gridlines, \n" +
              "Horizontal Alignment: Center, \n" +
              "Vertical Alignment: Center");*/

  // insert column
  //sheet.insertColumnBefore(1);
  // clear the format
  //sheet.getRange(1, 1, orgDataInfo.numRows+2*orgDataInfo.numRacks, 1).clearFormat();
  // adjast the column width
  //sheet.setColumnWidth(1, 80);
}

/**
 * creates a sample list sheet based on the user data and original sheet
 * @param {SpreadSheet} spreadsheet
 * @param {Dictionary} user data from the form
 * @param {Array} original data from the sheet as an array
 * @param {Array} extra info from the original data from the sheet
 */
function createSampleListSheet(ss, userData, orgDataArr, orgDataInfo)
{
  /* create a sample list sheet */
  var sheet = createNewSheet(ss, "Sample List");

  /* create a sample list filled with empty cells */
  var sampleData = [];
  for(var i = 0; i < orgDataInfo.numRows; i++){
    sampleData.push([" ", " ", " ", " ", " ", " ", " ", " ", " ", " "]);
  }

  /* add the sample index, rack name, plate number, and rack sample index to the sample list */
  var currPlateSamp = 1;
  var currPlateNum = userData["First plate number"];
  var currProjPlate = userData["Project Name"] + "-1";
  var currProjPlateNum = 1;
  for(var i = 0; i < orgDataInfo.numRows; i++)
  {
    if(currPlateSamp == 97){
        currPlateSamp = 1;
        currPlateNum++;
        currProjPlateNum++;
        currProjPlate = userData["Project Name"] + "-" + currProjPlateNum;
    }

    sampleData[i][0] = i+1; // add sample index
    sampleData[i][1] = currProjPlate; // add plate name
    sampleData[i][2] = currPlateNum; // add plate number
    sampleData[i][3] = currPlateSamp; // add plate sample index

    currPlateSamp++;
  }

  /* add "Cell Column" to the sample list */
  var rowNames = ["A", "B", "C", "D", "E", "F", "G", "H"];
  var colNames = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
  var currIdx = 0;
  // Rack Format: 96-All (d) is selected -> Sort down from top to bottom
  if(userData["Rack Format"] == "96-All (d)"){
    for(var i = 0; i < orgDataInfo.numRacks; i++){
      for(var col = 0; col < 12; col++){
        for(var row = 0; row < 8; row++)
        {
          sampleData[currIdx][4] = rowNames[row]+colNames[col];
          currIdx++;
        }
      }
    }
  } 
  // Rack Format: 96-All (a) is selected -> Sort accross from left to right
  else if(userData["Rack Format"] == "96-All (a)"){
    for(var i = 0; i < orgDataInfo.numRacks; i++){
      for(var row = 0; row < 8; row++){
        for(var col = 0; col < 12; col++)
        {
          sampleData[currIdx][4] = rowNames[row]+colNames[col];
          currIdx++;
        }
      }
    }
  }

  /* read original data array and
     add the samples and notes to the sample list */
  currIdx = 0;
  var shadedRows = [];
  for(var i = 0; i < orgDataArr.length; i++){
    for(var j = 1; j <= orgDataArr[i]["Plant #"]; j++)
    {
      // add "Plot Name/#" to the sample list
      sampleData[currIdx][6] = orgDataArr[i]["Plot Name/#"];

      // add "Plant #" to the sample list
      sampleData[currIdx][7] = j;

      if(j == 1) {
        // record the cells that have to be shaded
        shadedRows.push(currIdx+2);
      }
      
      // add note to the list
      sampleData[currIdx][8] = orgDataArr[i]["Note"];

      currIdx++;
    }
  }

  /* add column names to the sample data */
  sampleData.splice(0, 0, ["Sample", "Proj.Plate", "Plate #", "PlateSamp", "Cell", " ", "Plot Name/#", "Plant #", "Note", "Result"]);
  
  /* write the sample list array to the sheet */
  sheet.getRange(1, 1, orgDataInfo.numRows+1, 10).setValues(sampleData);

  /* format the sheet */
  sheet.getRange('A1:J1').setFontWeight('bold');
  sheet.getRange(1, 1, orgDataInfo.numRows+1, 10).setHorizontalAlignment('center');
  for(var i = 1; i <= 10; i++){
    sheet.autoResizeColumn(i);
  }
  sheet.setColumnWidth(6, 50);
  
  /* shade the plot name */
  if(userData["Shading Option"] == "First Plant") {
    for(var i = 0; i < shadedRows.length; i++){
      sheet.getRange(shadedRows[i], 7, 1, 1).setBackground('#c0c0c0');
    }
  } 
  else if(userData["Shading Option"] == "Every Other Full") {
    var currRow = 2;
    for(var i = 0; i < orgDataArr.length; i++){
      if(i%2 == 1){
        sheet.getRange(currRow, 7, orgDataArr[i]["Plant #"], 1).setBackground('#c0c0c0');
      }
      currRow += orgDataArr[i]["Plant #"];
    }
  } 
}
