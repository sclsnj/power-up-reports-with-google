/*
 * This script runs daily to process people counter data. It:
 *
 *   ** Looks for emails that have been labelled as "xTraffic Reports" (automatically labelled by filter) that 
 *      have been received in the last day.
 *   ** Extracts the CSV report attachment and turns the data into an array.
 *   ** Loops through the report data to get all hourly traffic data by branch.
 *   ** Dumps the contents of the final array into the end of the Data tab (1mgWnI4pYLN-OMiks8Bv9_ySjoKLEGwiFzQCyUCTr5pY).
 *
 * The main tab of People Count Trends uses dsum functions and conditional formatting to create a heat map
 * of patron traffic that can be filtered by branch and date range.
 *
 * This script is bound to this Google Sheet: https://docs.google.com/spreadsheets/d/1mgWnI4pYLN-OMiks8Bv9_ySjoKLEGwiFzQCyUCTr5pY/edit
 */


function getPeopleCountData() {
  // look in email for yesterday's data reports
  var detailFile;
  var now = new Date();
  var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);  
  var label = GmailApp.getUserLabelByName("xTraffic Reports");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    var message = threads[i].getMessages()[0];
    var messageDate = message.getDate();
    if (messageDate < now && messageDate > yesterday) {
      var subject = message.getSubject();
      if (subject.match(/Daily traffic/i)) {
        var attachments = message.getAttachments();
        for (var a in attachments) {
          var ext = attachments[a].getName();
          if (ext.match(/\.csv$/i)) {
            detailFile = attachments[a].getAs('csv');
            break;
          }
        }
      }
    }
  }
  // parse through report data and prepare for import
  var detailData;
  var newData = [];
  var date;
  detailData = Utilities.parseCsv(detailFile.getDataAsString());
  for (var d = 0; d < detailData.length; d++) {
    var detail = detailData[d];
    if (detail[0].match(/show/i)) {
      date = detail[0].match(/\d{1,2}\/\d{1,2}\/\d{4}/);
      var day = new Date(date).getDay() + 1;
    } else if (detail[0].match(/\d{2} [a-z]{4}/i)) {
      var branch = getBranch(detail[0]);
      newData.push([date, day, '9', branch, detail[3]]);
      newData.push([date, day, '10', branch, detail[4]]);
      newData.push([date, day, '11', branch, detail[5]]);
      newData.push([date, day, '12', branch, detail[6]]);
      newData.push([date, day, '13', branch, detail[7]]);
      newData.push([date, day, '14', branch, detail[8]]);
      newData.push([date, day, '15', branch, detail[9]]);
      newData.push([date, day, '16', branch, detail[10]]);
      newData.push([date, day, '17', branch, detail[11]]);
      newData.push([date, day, '18', branch, detail[12]]);
      newData.push([date, day, '19', branch, detail[13]]);
      newData.push([date, day, '20', branch, detail[14]]);
      newData.push([date, day, '21', branch, detail[15]]);
    }
  }
  if (newData.length > 0) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Data');
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newData.length, 5).setValues(newData);
  }
}

function getBranch(name) {
  var branch;
  if (name.match(/01/)) {
    branch = 'BBROOK';
  } else if (name.match(/02/)) {
    branch = 'BRIDGE';
  } else if (name.match(/03/)) {
    branch = 'HILLSB';
  } else if (name.match(/04/)) {
    branch = 'MANVLE';
  } else if (name.match(/05/)) {
    branch = 'NPLAIN';
  } else if (name.match(/06/)) {
    branch = 'PEAGLA';
  } else if (name.match(/07/)) {
    branch = 'SOMERV';
  } else if (name.match(/08/)) {
    branch = 'WARREN';
  } else if (name.match(/09/)) {
    branch = 'MJACOB';
  } else if (name.match(/10/)) {
    branch = 'WTCHNG';
  }
  return branch;
}
