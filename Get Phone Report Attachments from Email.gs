/*
 * This script runs daily to process phone system use data. It:
 *
 *   ** Makes sure there's a Phone Use Statistics file for the current year, and creates one if there isn't.
 *   ** Looks for emails that have been labelled as "Phone Reports" (automatically labelled by filter) that 
 *      have been received in the last day.
 *   ** Extracts the CSV report attachment and, if it's the right type, turns the data into an array.
 *   ** Pulls phone extension mapping data out of the Phone Use Statistics sheet.
 *   ** Loops through the report data to get all incoming and outgoing traffic recorded for each extension
 *      and pushing it into a new array. (This means creating a duplicate entry for in-house calls to represent 
 *      each side of the call.)
 *   ** Loops through the new array to compare to the mapping data, and pushes entries for any matching 
 *      extensions into the final array, adding extension name, branch and service desk information.
 *   ** Dumps the contents of the final array into the end of the Data tab.
 *
 * The main tab of the Phone Use Statistics uses dsum functions and conditional formatting to create a heat map
 * of phone traffic that can be filtered by branch, service desk, extension, and date range.
 *
 */


function getPhoneStats() {
  // look for a Drive file that has the right name; if it's not there, create it
  var now = new Date();
  var year = now.getFullYear();
  var files = DriveApp.searchFiles('title contains "' + year + ' Phone Use Statistics"');
  var id, spreadsheet;
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    var regExp = new RegExp(year);
    if (name.match(regExp) && name.match(/^[^(Copy)]/i)) {
      id = file.getId();
    }
  }
  if (id) {
    spreadsheet = SpreadsheetApp.openById(id);
  } else {
    spreadsheet = SpreadsheetApp.create(year + ' Phone Use Statistics');
    spreadsheet.insertSheet('Phone Traffic Patterns', 0);
    var dataSheet = spreadsheet.insertSheet('Data', 1);
    dataSheet.getRange(1, 1, 1, 10).setValues([['Date', 'Day', 'Time', 'Branch', 'Service Desk', 'Ext Name', 'Ext', 'Type', 'Where', 'Duration']]);
    dataSheet.deleteColumns(11, dataSheet.getLastColumn - 10);
    spreadsheet.insertSheet('Data Mapping', 2);
    var emptySheet = spreadsheet.getSheets()[3];
    spreadsheet.deleteSheet(emptySheet);
  }
  // look in email for yesterday's data reports
  var detailFile;
  //var now = new Date(now.getTime() - 0 * 24 * 60 * 60 * 1000);
  var yesterday = new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000);  
  var dateRegExp = new RegExp(yesterday);
  var label = GmailApp.getUserLabelByName("Phone Reports");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    var message = threads[i].getMessages()[0];
    var messageDate = message.getDate();
    if (messageDate < now && messageDate > yesterday) {
      var subject = message.getSubject();
      if (subject.match(/Detail/)) {
        detailFile = message.getAttachments()[0].copyBlob();
        break;
      }
    }
  }
  // get extension data ready
  var dataSheet = spreadsheet.getSheetByName('Data Mapping');
  var extData = dataSheet.getDataRange().getValues();
  // parse through report data and prepare for import
  if (detailFile) {
    var detailData;
    var newData = [];
    detailData = Utilities.parseCsv(detailFile.getDataAsString());
    for (var d = 0; d < detailData.length; d++) {
      var detail = detailData[d];
      var date = detail[0];
      var day = new Date(date).getDay() + 1;
      var time = detail[1];
      var fromExt = detail[2];
      var toExt = detail[3];
      var callType = detail[7];
      var dur = detail[9];
      if (callType == 'INC') {
        newData.push([date, day, time, '', '', '', fromExt, 'received', 'outside', dur]);
      } else if (callType == 'IHC') {
        newData.push([date, day, time, '', '', '', fromExt, 'placed', toExt, dur]);
        newData.push([date, day, time, '', '', '', toExt, 'received', fromExt, dur]);
      } else {
        newData.push([date, day, time, '', '', '', fromExt, 'placed', 'outside', dur]);
      }
    }
    var data = [];
    var branch, serviceDesk, extName;
    for (var x = 0; x < newData.length; x++) {
      var ext = newData[x][6];
      for (var e = 0; e < extData.length; e++) {
        var testExt = extData[e][0];
        if (ext == testExt) {
          branch = extData[e][2];
          serviceDesk = extData[e][3];
          extName = extData[e][1];
          break;
        }
      }
      if (extName) {
        data.push([newData[x][0], newData[x][1], newData[x][2], branch, serviceDesk, extName, newData[x][6], newData[x][7], newData[x][8], newData[x][9]]);
        branch = null;
        serviceDesk = null;
        extName = null;
      }
    }
    if (data.length > 0) {
      var sheet = spreadsheet.getSheetByName('Data');
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, data.length, 10).setValues(data);
    }
  } else {
    GmailApp.sendEmail('lhoffman@sclibnj.org', 'Couldn\'t get phone reports', '');
  }
}
