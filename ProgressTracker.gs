var segment = 1200000; //20 minutes
// This function creates a custom menu in Google Sheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Time Tracking')
    .addItem('Timers', 'showSidebar')
    .addToUi();
}

// This function shows the sidebar with the floating button
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Timers');
  SpreadsheetApp.getUi().showSidebar(html);
}

// This function is triggered by the floating button
function logTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = getLastRowInColumn(16);  // Get the last row with data in column 16
  var timestamp = new Date();

  var initialDate = new Date(1899, 11, 30);
  var storedValue = sheet.getRange(lastRow+1, 21).getValue();
  var storedValueDuration;
  
  if (storedValue === ""){
    storedValueDuration = 0;
  }
  else
  {
    storedValueDuration = storedValue - initialDate;
  }

  // If this is not the first row, calculate the duration from the previous timestamp
  if (lastRow > 0) {
    var prevTimestamp = new Date(sheet.getRange(lastRow, 16).getValue());
    var durationMs = timestamp - prevTimestamp + storedValueDuration;
    var duration = formatDuration(durationMs);
    sheet.getRange(lastRow + 1, 6).setValue(duration);
  }

  var level = sheet.getRange(4, 17).getValue();
  var handicap = sheet.getRange(5, 17).getValue();

  sheet.getRange(lastRow+1, 18).setValue(level);
  sheet.getRange(lastRow+1, 19).setValue(handicap);

  // Log the current timestamp
  sheet.getRange(lastRow + 1, 16).setValue(timestamp);
}

function startTime(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = getLastRowInColumn(16)
  var timestamp = new Date();

  if (lastRow == 3){
    sheet.getRange(lastRow + 1, 16).setValue(timestamp);
  }
  else
  {
    sheet.getRange(lastRow, 16).setValue(timestamp);
  }
  return 
}

function storeTimeSegment(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = getLastRowInColumn(16);  // Get the last row with data in column 16
  
  var initialDate = new Date(1899, 11, 30);
  var storedValue = sheet.getRange(lastRow+1, 21).getValue();
  var storedValueDuration;
  
  if (storedValue === ""){
    storedValueDuration = 0;
  }
  else
  {
    storedValueDuration = storedValue - initialDate;
  }
  
  var prevTimestamp = new Date(sheet.getRange(lastRow, 16).getValue());
  var currentTimestamp = new Date();
  
  var durationMs = currentTimestamp - prevTimestamp + storedValueDuration;
  var duration = formatDuration(durationMs);

  sheet.getRange(lastRow + 1, 21).setValue(duration);
}

function getLastRowInColumn(column) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange(1, column, sheet.getMaxRows(), 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== '') {
      return i + 1;
    }
  }
  return 0;
}

// Helper function to format duration in HH:MM:SS, handles negative durations
function formatDuration(durationMs) {
  var isNegative = durationMs < 0;
  var totalSeconds = Math.abs(Math.floor(durationMs / 1000));
  var hours = Math.floor(totalSeconds / 3600);
  var minutes = Math.floor((totalSeconds % 3600) / 60);
  var seconds = totalSeconds % 60;

  // Format the duration
  var formatted = pad(hours) + ':' + pad(minutes) + ':' + pad(seconds);

  // Add minus sign if the duration is negative
  return isNegative ? '-' + formatted : formatted;
}

// Helper function to pad single digit numbers with a leading zero
function pad(number) {
  return number < 10 ? '0' + number : number;
}

// Function to get the value from a cell
function getCellValue(cell, key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getRange(cell).getDisplayValue();
  return value.toString();
}

// Function to fetch values from multiple cells
function fetchCellValues() {
  var cellValues = {};
  cellValues.P2 = getCellValue('P2');
  cellValues.L2 = getCellValue('L2');
  cellValues.Q2 = getCellValue('Q2');
  cellValues.N2 = getCellValue('N2');
  return cellValues;
}

function pad2(number, digits = 2) {
  return number.toString().padStart(digits, '0');
}

// Helper function to format duration in HH:MM:SS.mmm, handles negative durations
function formatDuration2(durationMs) {
  var isNegative = durationMs < 0;
  var totalMilliseconds = Math.abs(durationMs);
  var totalSeconds = Math.floor(totalMilliseconds / 1000);
  var milliseconds = Math.floor(totalMilliseconds % 1000);
  var hours = Math.floor(totalSeconds / 3600);
  var minutes = Math.floor((totalSeconds % 3600) / 60);
  var seconds = totalSeconds % 60;

  // Format the duration
  var formatted = pad2(hours) + ':' + pad2(minutes) + ':' + pad2(seconds) + '.' + pad2(milliseconds, 3);

  // Add minus sign if the duration is negative
  return isNegative ? '-' + formatted : formatted;
}


function calculateMultiplier() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var initialDate = new Date(1899, 11, 30);
  var updateCell = false;

  var lastRow = getLastRowInColumn(8)
  var difference = sheet.getRange(lastRow, 8).getValue();
  
  var differenceMs = difference - initialDate; //Format is in terms of milliseconds

  if(differenceMs < 0){
    if (sheet.getRange(4, 17).getValue() != 0)
    {
      sheet.getRange(5, 17).setValue(sheet.getRange(5, 17).getValue() + 1);
    }
    updateCell = true;
  }
  else if (differenceMs >= 0){
    sheet.getRange(4, 17).setValue(sheet.getRange(4, 17).getValue() + 1);
    updateCell = true;
  }

  if (updateCell)
  {
    sheet.getRange(6, 17).setValue("0:00:00");
    sheet.getRange(8, 17).setValue(lastRow);

    var level = sheet.getRange(4, 17).getValue();
    var handicap = sheet.getRange(5, 17).getValue();
    var oldRate =  sheet.getRange(lastRow, 11).getValue() / sheet.getRange(4, 3).getValue();
    var percentage = 0.8;
    var oldMultiplier = 1 / oldRate;

    var newMultiplier = (oldMultiplier*2) + percentage*(oldMultiplier-oldMultiplier*2);

    sheet.getRange(4, 3).setValue(newMultiplier);
    sheet.getRange(lastRow, 18).setValue(level);
    sheet.getRange(lastRow, 19).setValue(handicap);
    SpreadsheetApp.flush();
    Utilities.sleep(250); 
    sheet.getRange(6, 17).setValue(formatDuration2(sheet.getRange(lastRow, 8).getValue() - initialDate));

  }
  //Needs to be here or potential error. 
  // Need to return variables level and handicap
  return [sheet.getRange(4, 17).getValue(),sheet.getRange(5, 17).getValue()];
}

function fetchStats(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  return [sheet.getRange(4, 17).getValue(),sheet.getRange(5, 17).getValue()]
}

// This function calculates the current duration of the segment
function calculateCurrentDuration() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = getLastRowInColumn(16);  // Get the last row with data in column 16
  var initialDate = new Date(1899, 11, 30);
  
  var timeLeft = new Date(sheet.getRange(lastRow + 1, 3).getValue());
  var timeLeftFormatted = formatDuration(timeLeft- initialDate);

  var totalTimeLeft = new Date(sheet.getRange(lastRow, 8).getValue());
  var totalTimeLeftFormatted = formatDuration(totalTimeLeft- initialDate);

  var storedValue = sheet.getRange(lastRow+1, 21).getValue();
  var storedValueDuration;

  //Finding % progress
  var currentProgress = Math.floor((sheet.getRange(lastRow, 8).getValue() - initialDate) / (sheet.getRange(7, 17).getValue() - initialDate)*100);
  var maxProgress = Math.floor((sheet.getRange(9, 17).getValue() - initialDate) / (sheet.getRange(7, 17).getValue() - initialDate)*100);

  if (storedValue === ""){
    storedValueDuration = 0;
  }
  else
  {
    storedValueDuration = storedValue - initialDate;
  }

  if (lastRow > 0) {
    var prevTimestamp = new Date(sheet.getRange(lastRow, 16).getValue());
    var currentTimestamp = new Date();
    var durationMs = currentTimestamp - prevTimestamp+storedValueDuration;
    var duration = formatDuration(durationMs);

    var timeLeftDuration = formatDuration(timeLeft - initialDate - durationMs);
    var totalTimeDuration = formatDuration(totalTimeLeft - initialDate + (timeLeft - initialDate - durationMs)+segment);

    return [duration, timeLeftDuration,totalTimeDuration, currentProgress, maxProgress];
  } else {
    return ['00:00:00', timeLeftFormatted, totalTimeLeftFormatted, currentProgress, maxProgress]; // If there is no previous timestamp, return 0 duration
  }
}