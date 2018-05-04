/*
  File name   : BasicUtilities.gs
  Created on  : 6/7/2017
  Author      : Takao Shibamoto
  Description : Basic javascript functions
*/

// convert the IOS date format to US date
function isoDateToUsDate(date)
{
  if(date.search('-') >= 0 && date.search('/') < 0){
    date = date.replace(/-/g, '/');
    date += '/';
    date = date.slice(4, 11) + date.slice(0, 4);
    date = date.substring(1);
  }

  return date;
}


// Add a new sheet in the given spreadsheet
function createNewSheet(ss, newSheetName)
{
  var newSheet = ss.getSheetByName(newSheetName);
  
  // if the new sheet name already exists, update it
  if (newSheet != null) {
    ss.deleteSheet(newSheet);
  }
  
  newSheet = ss.insertSheet();
  newSheet.setName(newSheetName);

  return newSheet;
}


// Add a new sheet in the given spreadsheet
function copySheet(ss, orgSheet, newSheetName)
{
  var newSheet = ss.getSheetByName(newSheetName);
  
  // if the new sheet name already exists, update it
  if (newSheet != null) {
    ss.deleteSheet(newSheet);
  }
  
  newSheet = ss.insertSheet(newSheetName, {template: orgSheet});

  return newSheet;
}


// print matrix in the Logger
function printMatrix(matrix)
{
  for (var row = 0; row < matrix.length; row++) {
    var line = "";
    for (var col = 0; col < matrix[0].length; col++) {
      line += matrix[row][col] + " ";
    }
    Logger.log(line);
  }
}


// create a 2d array filled with undefined
function create2dArray(row, col)
{
  var arr = new Array(row);
  for (var i = 0; i < row; i++) {
    arr[i] = new Array(col);
  }
  return arr;
}


// get the current data in a clean format
function getCurrentDate(){
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();

  if(dd<10) {
    dd='0'+dd
  } 

  if(mm<10) {
    mm='0'+mm
  } 

  today = mm+'/'+dd+'/'+yyyy;
  
  return today;
}
