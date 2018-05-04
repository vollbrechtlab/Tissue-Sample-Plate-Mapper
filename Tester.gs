function testD(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (d)";
  userData["Shading Option"] = "First Plant";
  userData["First page number"] = 4;
  startGeneration(userData);
}

function testD2(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (d)";
  userData["Shading Option"] = "Every Other";
  userData["First page number"] = "144";
  startGeneration(userData);
}

function testA(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (a)";
  userData["Shading Option"] = "First Plant";
  userData["First page number"] = 4;
  startGeneration(userData);
}

function testA2(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (a)";
  userData["Shading Option"] = "Every Other";
  userData["First page number"] = 4;
  startGeneration(userData);
}

function test0(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (d)";
  userData["Shading Option"] = "First Plant";
  userData["First page number"] = 4;
  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  userData = fixUserData(userData);
  var orgDataArr = readOrgData(ss, userData);
  //createSummarySheet(ss, userData, orgDataArr);
  createRackLabelsSheet(ss, userData, orgDataArr);
  //createSampleListSheet(ss, userData, orgDataArr);
  //ss.setActiveSheet(ss.getSheetByName("Summary"));
}

function testPrintingSheet(){
  var userData = {};
  userData["Project Name"] = "TestProject";
  userData["Station/Site"] = "somewhere";
  userData["Sample Date"] = "7/1/2017";
  userData["Crop"] = "Corn (Field)";
  userData["Rack Format"] = "96-All (d)";
  userData["Shading Option"] = "First Plant";
  userData["First page number"] = 4;


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  userData = fixUserData(userData);

  // read the original data and store it in a dictionary
  var orgDataArr = readOrgData(ss, userData);

  // calculate the number of samples and racks
  var orgDataInfo = countSamplesAndRacks(orgDataArr);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createRackLabelsForPrintingSheet(ss, userData, orgDataInfo);
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}