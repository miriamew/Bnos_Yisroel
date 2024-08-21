function convertArmy(time) {
  if (time[5] === "p") {
    if (time[0] === "1" && time[1] === "2") {
      time = time;
    }
    else {
      time = String(parseInt(time[0]) + 1) + String(parseInt(time[1]) + 2) + time.slice(2);
    }
  }
  return time
}
function subtractTimes(inTime, shouldstart) {

  results = ((parseInt(inTime[0] + inTime[1]) - parseInt(shouldstart[0] + shouldstart[1])) * 60 +
    (parseInt(inTime[3] + inTime[4]) - parseInt(shouldstart[3] + shouldstart[4])));
  return results
}

function payrollDate() {
  today = new Date()
  console.log(today.getDate())
  console.log(((new Date(today.getFullYear(), today.getMonth(), 15))))
  newdate = (new Date(today.getFullYear(), today.getMonth() + 1, 0)).getDate();
  console.log(newdate);
  //if (parseInt(today.getDate()) == parseInt(newdate) - 5){
  if (parseInt(today.getDate()) > 15) {
    console.log((new Date(today.getFullYear(), today.getMonth() + 1, 0)))
    return ((new Date(today.getFullYear(), today.getMonth() + 1, 0)))

  }
  else {
    console.log("test")
    return ((new Date(today.getFullYear(), today.getMonth(), 15)))
  }
}
function timeclocks() {
  var spreadsheet = SpreadsheetApp.openById('--');
  var spreadsheet2 = SpreadsheetApp.openById('')
  var timeclockSheet = spreadsheet.getSheetByName('Newest Icon');
  var scheduleSheet = spreadsheet.getSheetByName('Employees');
  var current = spreadsheet.getSheetByName('Current');
  var scheduleSheet2 = spreadsheet2.getSheetByName('Employees');
  var current2 = spreadsheet2.getSheetByName('Current');
  var scheduledata = scheduleSheet.getDataRange().getValues();
  var scheduledata2 = scheduleSheet2.getDataRange().getValues();
  var timeclockdata = timeclockSheet.getDataRange().getValues();
  console.log(payrollDate())
  scheduleSheet.getRange(1, scheduledata[0].length + 1, 1, 2).setValues([[payrollDate(), "Hours"]])
  current.getRange(1, 3, 1, 2).setValues([[payrollDate(), "Hours"]])
  scheduleSheet2.getRange(1, scheduledata2[0].length + 1, 1, 2).setValues([[payrollDate(), "Hours"]])
  current2.getRange(1, 3, 1, 2).setValues([[payrollDate(), "Hours"]])

  //scheduleSheet.getRange(1, scheduledata[0].length+1,1,1).setValue(" ")
  for (let i = 1; i < 22; i++) {
    found = false
    Logger.log(i);
    date = ""
    line = (scheduledata[i]);
    flag = false;
    let d = new Date();
    var sheetName = String(line[0]);
    rate = line[3]
    var personsheet = spreadsheet.getSheetByName(sheetName);
    var personsheet2 = spreadsheet2.getSheetByName(sheetName);
    personsheet2.getRange(2, 1, personsheet2.getLastRow(), 9).clearContent();
    //console.log(personsheet)
    var persondata = personsheet.getDataRange().getValues();
    amountofrows = persondata.length;
    personsheet.appendRow(["h"]);
    var upto = 0;
    for (let k = 1; k < timeclockdata.length; k++) {
      row = (timeclockdata[k]);
      if (line[2] === row[2] && (row[2] != "")) {
        found = true
        if (String(line[2]) != "60184" || String(row[5])[1] != "2") {
          realdate = new Date(String(row[3]).substring(4, 15));
          realdate = Utilities.formatDate(realdate, 'America/New_York', 'MM/dd/yyyy');
          date = (String(row[3]).substring(4, 15));
          console.log(date);
          inDay = row[4];
          outDay = row[6];
          inTime = row[5];
          outTime = row[7];
          savedintime = inTime;
          savedouttime = outTime;
          inTime = convertArmy(inTime);
          outTime = convertArmy(outTime);
          minutes = subtractTimes(outTime, inTime);
          hours = minutes / 60

          hours = Math.round(hours * 100) / 100
          if (String(inDay) !== String(outDay)) {
            //newmonth(realdate, personsheet);                
            personsheet.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], "", "", "Didn't clock OUT"]);
            personsheet2.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], "", "", "Didn't clock OUT"]);
            flag = true

          }
          else if (inTime == outTime) {
            //newmonth(realdate, personsheet);
            personsheet.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], "", "", "Didn't clock IN"]);
            personsheet2.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], "", "", "Didn't clock IN"]);
            flag = true
          }
          else {
            //newmonth( realdate, personsheet);
            personsheet.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], minutes, hours]);
            personsheet2.appendRow([row[0], row[2], realdate, row[4], row[5], row[6], row[7], minutes, hours]);
          }
        }
      }
    }
    personsheet.getRange(amountofrows + 1, 1, 1, 1).clearContent();
    formula = "=SUM(I" + String(amountofrows + 1) + ":I" + String(personsheet.getLastRow() + ")")
    formula2 = "=ROUND(B" + String(personsheet.getLastRow() + 1) + "*" + String(rate) + ")"
    formula3 = "=SUM(I2" + ":I" + String(personsheet2.getLastRow() + ")")
    formula4 = "=ROUND(B" + String(personsheet2.getLastRow() + 1) + "*" + String(rate) + ")"
    personsheet.appendRow(["Total hours:", formula, "Payment", formula2])
    personsheet2.appendRow(["Total hours:", formula3, "Payment", formula4])
    personsheet.getRange(personsheet.getLastRow(), 1, 1, 4).setFontWeight("bold").setFontFamily("Calibri").setBackground("lime");
    personsheet.getRange(personsheet.getLastRow(), 4, 1, 1).setNumberFormat("[$$]#");
    personsheet2.getRange(personsheet2.getLastRow(), 1, 1, 4).setFontWeight("bold").setFontFamily("Calibri").setBackground("lime");;
    personsheet2.getRange(personsheet2.getLastRow(), 4, 1, 1).setNumberFormat("[$$]#");
    totalpayment = personsheet.getRange(personsheet.getLastRow(), 4, 1, 1).getValue();
    totalhours = personsheet.getRange(personsheet.getLastRow(), 2, 1, 1).getValue()
    if (found == false) {
      totalhours = 0
      totalpayment = 0
    }
    //scheduleSheet.getRange(i+1, scheduleSheet.getLastColumn()-1,1,2).setValues([["$" + String(totalpayment),String(Math.round(totalhours * 100) / 100) + " hrs" ]]).setFontWeight("bold");
    if (flag) {
      scheduleSheet.getRange(i + 1, scheduleSheet.getLastColumn() - 1, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("red");
      current.getRange(i + 1, 3, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("red");
      scheduleSheet2.getRange(i + 1, scheduleSheet2.getLastColumn() - 1, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("red");
      current2.getRange(i + 1, 3, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("red");

    }
    if (!flag) {
      scheduleSheet.getRange(i + 1, scheduleSheet.getLastColumn() - 1, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold");
      current.getRange(i + 1, 3, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("white");
      scheduleSheet2.getRange(i + 1, scheduleSheet2.getLastColumn() - 1, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold");
      current2.getRange(i + 1, 3, 1, 2).setValues([["$" + String(totalpayment), String(Math.round(totalhours * 100) / 100) + " hrs"]]).setFontWeight("bold").setBackground("white");

    }
    //scheduleSheet.getRange(i+1, scheduleSheet.getLastColumn(),1,1).setValue()
  }
  formatSheets()
}

function formatSheets() {
  var spreadsheet = SpreadsheetApp.openById('--');
  var spreadsheet2 = SpreadsheetApp.openById('')
  var scheduleSheet = spreadsheet.getSheetByName('Employees');
  var scheduledata = scheduleSheet.getDataRange().getValues();
  header = [["Name", "ID", "Date", "Day In", "Time In", "Day Out", "Time Out", "Total Minutes", "Total Hours"]]
  for (let i = 1; i < 22; i++) {
    line = (scheduledata[i]);
    var sheetName = String(line[0]);
    console.log(sheetName)
    var personsheet1 = spreadsheet.getSheetByName(sheetName);
    var personsheet2 = spreadsheet2.getSheetByName(sheetName);
    personsheet1.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
    personsheet2.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
    var personsheet2 = spreadsheet2.getSheetByName(sheetName);
    personsheet1.getRange(1, 1, 1, 9).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan");;//;
    personsheet1.getRange(1, 1, 1, 9).setValues(header)
    personsheet2.getRange(1, 1, 1, 9).setValues(header)
    //personsheet1.getRange(2, 1, 200, 9).clear()
    //personsheet2.getRange(2, 1, 200, 9).clear()
    personsheet2.getRange(1, 1, 1, 9).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan");;//
    if (String(scheduledata[i][4]) == "") {
      var spreadsheet3 = SpreadsheetApp.create(String(scheduledata[i][2]));
      scheduleSheet.getRange(i + 1, 5).setValue(spreadsheet3.getId());
      var file = DriveApp.getFileById(spreadsheet3.getId());
      var folder = DriveApp.getFolderById("1sDb7q9synz2kd5OXAKD6PX0Tv2XC18eU");
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);

    }
    if (String(scheduledata[i][4]) != "") {
      var targetSheet = SpreadsheetApp.openById(scheduledata[i][4]);
      var targetSheet = targetSheet.getSheetByName('Sheet1');
      sourceData = personsheet1.getDataRange().getValues();
      console.log(sourceData.length)
      var sBG = personsheet1.getDataRange().getBackgrounds();
      var sFC = personsheet1.getDataRange().getFontColors();
      var sFF = personsheet1.getDataRange().getFontFamilies();
      targetSheet.clearContents()
      targetSheet.getRange(1, 1, 30, 20).clearFormat().setBackground(null).setBorder(false, false, false, false, false, false);; // Clear existing data in the target sheet//
      targetSheet.getRange(1, 1, sourceData.length, 9).setValues(sourceData).setFontFamily("Calibri").setHorizontalAlignment("left").setBackgrounds(sBG).setFontColors(sFC).setFontFamilies(sFF);//
      targetSheet.getRange(1, 1, sourceData.length, 9).setBorder(true, true, true, true, true, true);
      targetSheet.getRange(1, 1, 1, 9).setValues(header)
      targetSheet.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
      targetSheet.getRange(1, 1, 1, 9).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan")
      targetSheet.autoResizeColumn(7);
      targetSheet.autoResizeColumn(8);




    }

  }
}
