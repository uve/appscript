function run() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  sheet.setFrozenRows(1);
  
  
  var label = 'run() time';  // Labels the timing log entry.
  console.time(label);              // Starts the timer.
    
  
  var lastColumn = sheet.getLastColumn();  
  
  var top = sheet.getRange(1, 1, 1, lastColumn);
  top.setFontWeight("bold");
  top.setHorizontalAlignment("center");
  
  
  var lastRow = sheet.getLastRow();
  
  var columnDate = sheet.getRange(2, 2, lastRow, 2);
  columnDate.setHorizontalAlignment("center");
    

  var dataRange = sheet.getRange(2, 1, lastRow, 2);
  var data = dataRange.getValues();
    
    
  var numbersRange = sheet.getRange(2, 4, lastRow, 6);
  numbersRange.setNumberFormat("0.00");
  
  
  setRealDiff();
  setDateColor();
  setLatestQuarterColor();
  
 
   
  console.timeEnd(label);
 
      
  /*
  var entries = [{name : "Check Duplicates",functionName : "checkDuplicates"}];
  sheet.addMenu("Scripts", entries);
  */
};


function getLastQuarter() {
  var date = new Date() ;
  var month = date.getMonth();
  
  if (month < 3) {
    return "Q4";
  } else if (month < 6) {
    return "Q1";
  } else if (month < 9) {
    return "Q2";
  } else {
    return "Q3";
  }
}

function setLatestQuarterColor() { 
   // Set color data to highlight the latest day
   var sheet = SpreadsheetApp.getActiveSheet(); 
   var lastRow = sheet.getLastRow() - 1;
   
   var lastQuarter = getLastQuarter();
   var columnIndex = 3;
   var rangeGreen = [];

   var dataRange = sheet.getRange(1, columnIndex, lastRow, 1);
   dataRange.setBackground("white");  
    
 
   var data = dataRange.getValues();
  
   for (var i = 2; i < lastRow; i++) {
     var cell = dataRange.getCell(i, 1);
     var value = cell.getValue();
     
     if (value == lastQuarter) {
        rangeGreen.push("C" + i);
     }
 }
  
 var rangeList = sheet.getRangeList(rangeGreen);
 rangeList.setBackground("lightgreen");   
 
 
}

function setDateColor() { 
   // Set color data to highlight the latest day
   var sheet = SpreadsheetApp.getActiveSheet(); 
   var lastRow = sheet.getLastRow() - 1;
   
   var today = new Date() ;
   
   var columnIndex = 2;

   var rangeGreen = [];
   var rangeYellow = [];


   var dataRange = sheet.getRange(1, columnIndex, lastRow, 1);
   dataRange.setBackground("white");
    
 
   var data = dataRange.getValues();
  
   for (var i = 2; i < lastRow; i++) {
     var cell = dataRange.getCell(i, 1);
     var value = cell.getValue();
     
     var diff = daysDiff(value, today);
  
     if (diff < 7) {
        rangeGreen.push("B" + i);
     } else if ((diff >= 7) && (diff < 14)) {
       rangeYellow.push("B" + i);
     } 
 }
  
 var rangeList = sheet.getRangeList(rangeGreen);
 rangeList.setBackground("lightgreen");   
 
 rangeList = sheet.getRangeList(rangeYellow);
 rangeList.setBackground("yellow");   
 
}


function setRealDiff() { 
   // Set google finance price
   var sheet = SpreadsheetApp.getActiveSheet(); 
   var lastRow = sheet.getLastRow() - 1;
   
   var columnRealDiffIndex = 12;

   var top = sheet.getRange(1, columnRealDiffIndex);
   top.setValue("Real Diff");
   
   var columnGoogleDiff = sheet.getRange(2, columnRealDiffIndex, lastRow, 1);
   columnGoogleDiff.setFormula("=K2-E2");       
 
   var dataRange = sheet.getRange(1, columnRealDiffIndex, lastRow, 1);
   dataRange.setBackground("white");
  
   var rangeGreen = [];
   var data = dataRange.getValues();
  
   for (var i = 2; i < lastRow; i++) {
     var cell = dataRange.getCell(i, 1);
     var value = cell.getValue();
  
     if (value > 0) {   
       rangeGreen.push("L" + i);
     }
 }
  
 var rangeList = sheet.getRangeList(rangeGreen);
 rangeList.setBackground("lightgreen");   
}



function daysDiff(a, b) {
    var oneDay = 1000 * 60 * 60 * 24;
    return Math.floor(b.getTime() / oneDay) - Math.floor(a.getTime() / oneDay);
}




