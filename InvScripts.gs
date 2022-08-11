// Inventory Processing
// v.1.1
// June 28, 2022
// Run in Google Apps Script

// Last update: August 11, 2022
// adjusted clean() to account for errors when deleting rows

// Adds menu functions for processing
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Add Prefix", functionName: "prefix"});
  menuEntries.push({name: "Clean Report", functionName: "clean"});
  menuEntries.push({name: "Generate Stats", functionName: "reportStats"});
  menuEntries.push({name: "Generate Scan List", functionName: "genScanList"});
  
  ss.addMenu("Inventory",menuEntries);
}

// Adds prefix for processing
function prefix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scans = ss.getSheetByName("Scans");
  var scanLen = scans.getLastRow();
  for (var row = 1; row <= scanLen; row = row+1) {
    var bcode = scans.getRange(row,1).getValue();
    scans.getRange(row,1).setValue("n:"+bcode);
  }
}

// Cleans raw report data for processing
function clean() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawData = ss.getSheetByName("RawData");
  var listLen = rawData.getLastRow();
  for (line = listLen; line >= 57; line = line - 1) {
    var value = rawData.getRange(line,1).getValue();
    if (value == '') {
      rawData.deleteRows(line-2,3);
      if (listLen >= line-2) {
        line = line-3;
      }
    }
  }
}

// Generates Stats summary
function reportStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawData = ss.getSheetByName("RawData");

  // Checks for stats sheet and creates if one does not exist
  if(ss.getSheetByName("InvStats")== null) {
    ss.insertSheet('InvStats');
  }

  // Creates array & counter for stats export
  var statsExport = new Array();
  var statsCount = 0;

  //Sets headings for stats
  var statSheet = ss.getSheetByName("InvStats");
  var statHead = statSheet.getRange(1,1);
  statHead.setValue("Report Date:").setFontWeight("bold").setHorizontalAlignment("right");
  var reportDate = rawData.getRange(1,1).getValue().match("[\\D]{3}\\s[\\D]{3}\\s[\\d]{1,2}\\s[\\d]{4}");
  statHead.offset(0,1).setValue(reportDate);

  statsExport[statsCount] = reportDate;
  statsCount=statsCount+1;
  
  statHead.offset(0,2).setValue("Report Location:").setFontWeight("bold").setHorizontalAlignment("right");
  var reportLoc = rawData.getRange(1,1).getValue().match("(?<=LOCATION).*");
  statHead.offset(0,3).setValue(reportLoc);

  statsExport[statsCount]=reportLoc;
  statsCount=statsCount+1;
  
  statHead.offset(2,0,2).setValues([["Beginning Call#:"],["Ending Call#:"]]).setFontWeight("bold").setHorizontalAlignment("right");
  statHead.offset(2,2,2).setValues([["Beginning Barcode:"],["Ending Barcode:"]]).setFontWeight("bold").setHorizontalAlignment("right");
  var statTotals = statSheet.getRange(6,1,4);
  statTotals.setValues([["Total Items in Range:"],["Currently Checked Out:"],["Off-Shelf Status:"],["Expected on Shelf:"]]).setFontWeight("bold").    setHorizontalAlignment("right");
  var statOverview = statSheet.getRange(11,1,4);
  statOverview.setValues([["Total Barcodes in Input File:"],["Total Items Misshelved (wrong unit):"],["Total Items Misshelved (correct unit):"],["Total Items Missing:"]]).setFontWeight("bold").setHorizontalAlignment("right");
  
  // Extracts & writes summary numbers from report
  for (var line = 3; line <= 4; line = line+1){
    var callString = rawData.getRange(line,1).getValue();
    var printBCode = callString.match("\\d{14}");
    statSheet.getRange(line,4).setValue(printBCode);
    var printCall = callString.match("(?<=\\d{14}).*");
    statSheet.getRange(line,2).setValue(printCall);

    statsExport[statsCount]=printCall;
    statsCount = statsCount+1;
  }
  
  for (var line = 6; line <= 9; line = line+1){
    var callString = rawData.getRange(line,1).getValue();
    var printSummary = callString.match("\\d{1,}");
    statSheet.getRange(line,2).setValue(printSummary);

    statsExport[statsCount]=printSummary;
    statsCount = statsCount+1;
  }
  
  var count = 0;
  for (var line = 11; line <= 17; line = line+2){
    var callString = rawData.getRange(line,1).getValue();
    var printTotals = callString.match("\\d{1,10}");
    var printLine = 11+count;
     statSheet.getRange(printLine,2).setValue(printTotals);
    count = count+1;

    statsExport[statsCount]=printTotals;
    statsCount = statsCount+1;
  }

  // Exports stats to dashboard
  var exportSS = SpreadsheetApp.openById("1vgi6bw0QclXI6k2sJPDcUHI0Q5FTz3TnGCobb3H4KrI");
  var exportSheet = exportSS.getSheetByName("RawData");
  var row = exportSheet.getLastRow();
  for (i=0;i<=11;i++) {
    var newRow = exportSheet.getRange(row+1,1);
    newRow.offset(0,i).setValue(statsExport[i]);
  } 
}

// Generates Scanned Item Report
function genScanList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawData = ss.getSheetByName("RawData");

  // Checks for ScanList sheet and creates if one does not exist
  if(ss.getSheetByName("ScanList")== null) {
    ss.insertSheet('ScanList');
  }
  var scanList = ss.getSheetByName("ScanList");
  scanList.getRange(1,1,1,3).setValues([["Shelf Placement","Status", "Call # & Title"]]).setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);

  // Extracts item information
  var listLength = (rawData.getLastRow())-23;
  

  for (var line = 1; line <= listLength+1; line = line+1) {
    var item = rawData.getRange(22+line,1).getValue();
    var itemStart = item.match("^\\d{1,4}\.");
    if (itemStart == null) {

      // Extracts & writes scan order
      var itemPlace = null;
      scanList.getRange(line+1,1).setValue(itemPlace);

      // Extracts & writes status
      var itemStatus = item.match("^.*(?=\\s[\\d]{6})");
      scanList.getRange(line+1,2).setValue(itemStatus);

      // Extracts & writes Call number & Title
      var itemCall = item.match("(?<=\\d{6})\\s.*");
      scanList.getRange(line+1,3).setValue(itemCall);

      // Highlights errors on Scan List
      var status = scanList.getRange(line+1,2).getValue();
      if (status.match("ERR")) {
        scanList.getRange(line+1,1,1,3).setBackground('yellow');
      }

    } else {
      // Extracts & writes scan order
      var itemPlace = item.match("^\\d{1,4}");
      scanList.getRange(line+1,1).setValue(itemPlace);

      // Extracts & writes status
      var itemStatus = item.match("[A-Za-z\\s=\\d]*(?=\\s[\\d]{6})");
      scanList.getRange(line+1,2).setValue(itemStatus);

      // Extracts & writes Call number & Title
      var itemCall = item.match("(?<=\\d{6})\\s.*");
      scanList.getRange(line+1,3).setValue(itemCall);

      // Highlights errors on Scan List
      var status = scanList.getRange(line+1,2).getValue();
      if (status.match("ERR")) {
        scanList.getRange(line+1,1,1,3).setBackground('yellow');
       }
    }
  }
}
