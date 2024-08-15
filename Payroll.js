function test() {

}

function formatTime(time) {
  hours = (time.getHours())
  console.log(hours)
  minutes = (time.getMinutes())
  s = " am"
  if (hours > 12) {
    console.log("afternoon")
    hours = hours - 12;
    s = " pm"
  }
  if (hours == 12) {
    console.log("noon");
    s = " pm";
  }
  if (minutes < 10) {
    minutes = "0" + String(minutes);
  }
  if (hours == 0) {
    hours = 12
  }
  console.log((String(hours) + ":" + String(minutes) + s))
  return (String(hours) + ":" + String(minutes) + s)
}

function emails() {
  rightnow = new Date();
  var spreadsheet = SpreadsheetApp.openById('1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU');
  var emailsheet = spreadsheet.getSheetByName('Emails');
  emaildata = emailsheet.getDataRange().getValues();
  time = (String(rightnow.getHours()));
  day = rightnow.getDay();
  if (day > 0 && day < 6 && time === "6") { //if it is a work day, reset all columns to false
    for (i = 1; i < emaildata.length + 1; i++) {
      emailsheet.getRange(i + 1, 4, 1, 1).setValue(false);
      emailsheet.getRange(i + 1, 6, 1, 1).setValue(false);
      emailsheet.getRange(i + 1, 8, 1, 1).setValue(false);
      emailsheet.getRange(i + 1, 10, 1, 1).setValue(false);
      emailsheet.getRange(i + 1, 12, 1, 1).setValue(false);
    }
  }
  if (day > 0 && day < 6 && time === "17") { //in the 5:00 hour, send out reminder emails to whoever didn't submit their form yet
    for (i = 1; i < emaildata.length + 1; i++) {
      email = emailsheet.getRange(i + 1, 2, 1, 1).getValues();
      teaches = emailsheet.getRange(i + 1, ((day) * 2) + 1, 1, 1).getValues();
      taught = emailsheet.getRange(i + 1, (day) * 2 + 2, 1, 1).getValues();
      console.log(email[0][0])
      if (teaches[0][0] == true && taught[0][0] == false) {
        Utilities.sleep(2000)
        console.log("sending")
        GmailApp.sendEmail(email[0][0], "Bnos Yisroel Payroll", "", {
          name: "Bnos Yisroel",
          htmlBody: `<p>Our records indicate that you have job(s) that require you to report your hours/periods worked today.
If you are receiving this email, we do not have a record of you submitting the form for today. Please submit the <a href=https://docs.google.com/forms/d/e/1FAIpQLSdnmYmN39J4BGt7SxlZGY0Kjtcu9Bt_dodAMezxYzGnIivTWw/viewform>PAYROLL FORM.</a> Thank you!<br /></p>`
        });
      }
    }
  }
}

function onFormSubmit(e) {
  var spreadsheet = SpreadsheetApp.openById('1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU');
  var form = FormApp.openById('1nQfaAIQ75yV-qjbxdGHjlYYLl5lFOtKJpcRANIl_FPk');
  var responses = e.values;
  var sheet = spreadsheet.getSheetByName('Form Responses 1');
  var rateSheet = spreadsheet.getSheetByName('Rates');
  var emailsheet = spreadsheet.getSheetByName('Emails');
  emaildata = emailsheet.getDataRange().getValues();
  responses = responses.filter(function (element) {
    return element.trim() !== "";
  });
  var nameResponse = responses[2];
  var jobResponse = responses[4];
  rate = 0
  type = ""
  payment = 0
  minlate = 0
  comments = "None"
  comments2 = ""
  error = ""
  errorMsg = ""
  ratedata = rateSheet.getDataRange().getValues()
  for (i = 0; i < ratedata.length; i++) {
    line = ratedata[i];
    if (line[0] == nameResponse) {
      if (line[1] == jobResponse) {
        console.log(nameResponse + " " + jobResponse)
        rate = line[3];
        type = line[2];
        var spreadsheet2 = SpreadsheetApp.openById(line[5]);
        var sheet2 = spreadsheet2.getSheetByName("Sheet1");
        var personal = spreadsheet.getSheetByName(nameResponse);
        dateworked = new Date(responses[3])
        if (type == "Period") {
          var periodResponse = responses[5];
          if (String(jobResponse) == "Meeting with Adina Bulman") { periodResponse = 1 }
          var formResponses = form.getResponses();
          for (var m = formResponses.length - 1; m < formResponses.length; m++) {
            var formResponse = formResponses[m];
            var itemResponses = formResponse.getItemResponses();
            for (var b = 0; b < itemResponses.length; b++) {
              var itemResponse = itemResponses[b];
              if (itemResponse.getItem().getTitle() == "Additional Comments and Notes") {
                comments = String(itemResponse.getResponse());
              }
              if (itemResponse.getItem().getTitle() == "If you arrived late or left early, how many minutes late/left early were you? " && parseFloat(itemResponse.getResponse()) != NaN) {
                minlate = (itemResponse.getResponse());
                comments2 = "<br />Minutes arrived late: " + minlate
                periodResponse = periodResponse - (minlate / 42)
                periodResponse = Math.round((periodResponse) * 100) / 100;
              }
            }
          }
          personinfo = [responses[0], responses[1], responses[2], responses[3], responses[4], "", "", periodResponse, comments]
          ratemessage = "$" + String(rate) + "/Period"
          payment = Math.round(parseFloat(periodResponse) * parseFloat(rate) * 100) / 100
          moreinfo = `Total Period(s): ${periodResponse} `
          if (String(jobResponse) == "Meeting with Adina Bulman") {
            persondata = personal.getDataRange().getValues()
            for (entry = persondata.length - 1; entry > 0; entry--) {
              olddateworked = new Date(persondata[entry][3])
              if (((dateworked.getTime() - olddateworked.getTime()) / (24 * 3600 * 1000)) > 7) {
                break
              }
              if (dateworked.getDay() >= olddateworked.getDay() && String(persondata[entry][4]) == "Meeting with Adina Bulman") {
                personinfo = []
                error += "Error - Can only submit one meeting per week\n";
                errorMsg += "Can only submit one meeting per week\n";
                break
              }
            }
          }
        }
        if (type == "Hour") {
          console.log("in hours")
          comments = String(responses[7])
          ratemessage = "$" + String(rate) + "/hr"
          var time1 = new Date("January 1, 2000 " + responses[5]);
          var time2 = new Date("January 1, 2000 " + responses[6]);
          var hours = Math.round(((time2 - time1) / 3600000) * 100) / 100;
          payment = Math.round((hours * parseFloat(rate)) * 100) / 100
          personinfo = [responses[0], responses[1], responses[2], responses[3], responses[4], responses[5], responses[6], hours, responses[7],]
          time1st = formatTime(time1)
          time2st = formatTime(time2)
          day = new Date(time1.getFullYear(), time1.getMonth(), time1.getDate())
          time1 = time1.getTime() - day.getTime();
          day = new Date(time2.getFullYear(), time2.getMonth(), time2.getDate())
          time2 = time2.getTime() - day.getTime();
          moreinfo = `Start Time: ${time1st}  <br /> End Time: ${time2st}<br /> Total Hour(s): ${hours}`
          if (time1 > time2) {
            error += "Error - End time not later than start time\n";
            errorMsg += `<li>The end time (${time2st}) is not later than the start time (${time1st}).</li><br />`;
          }
          if (time1 < 28800000 || time1 > 64800000) {
            error += "Error - Start time not between 8:00 AM and 6:00 PM\n";
            errorMsg += `<li>The start time (${time1st}) is not between 8:00 AM and 6:00 PM.</li><br />`;
          }
          if (time2 < 28800000 || time2 > 64800000) {
            error += "Error - End time not between 8:00 AM and 6:00 PM\n";
            errorMsg += `<li>The end time (${time2st}) is not between 8:00 AM and 6:00 PM.</li><br />`;
          }
        }
        if (type == "Vacation") {
          var formResponses = form.getResponses();
          for (var m = formResponses.length - 1; m < formResponses.length; m++) {
            var formResponse = formResponses[m];
            var itemResponses = formResponse.getItemResponses();
            for (var b = 0; b < itemResponses.length; b++) {
              var itemResponse = itemResponses[b];
              if (itemResponse.getItem().getTitle() == "Additional Comments and Notes") {
                comments = String(itemResponse.getResponse());
              }
            }
          }
          ratemessage = "$" + String(rate) + "/hr"

          hours = 0
          for (n = 5; n < 10; n++) {
            if (responses[n]) {
              timesworked = responses[n].split(":")
              console.log(timesworked)
              if (timesworked.length == 3) {
                console.log(parseFloat(timesworked[0]))
                console.log(parseFloat(timesworked[1] / 60))
                hours = hours + parseFloat(timesworked[0]) + parseFloat(timesworked[1] / 60)
                console.log(hours)
              }
            }
          }
          hours = Math.round((hours) * 100) / 100
          payment = Math.round((hours * parseFloat(rate)) * 100) / 100
          moreinfo = `Total Hour(s): ${hours}`
          personinfo = [responses[0], responses[1], responses[2], responses[3], responses[4], , , hours, responses[7],]
        }
        if (String(comments) == "undefined") {
          comments = "None"
        }
        body = "<p>Thank you for submitting your payroll form. As a reminder, forms submitted by the 10th  of the month will be added to the payroll on the 15th, forms submitted by the 25th of the month will be added to the payroll on the last day of the month. <br /> Thank you for all your hard work!</p><br />";
        if (error != "") {
          body = `<p>There was an error with your form. Please resubmit the <a href=https://docs.google.com/forms/d/e/1FAIpQLSdnmYmN39J4BGt7SxlZGY0Kjtcu9Bt_dodAMezxYzGnIivTWw/viewform>PAYROLL FORM.</a> Thank you!<br /><br /> Error(s): <ol>${errorMsg}</ol> <br /></p>`;
          payment = 0;
        }
        else {
          sheet2.appendRow(personinfo)
          personal.appendRow(personinfo)
          sheet2.getRange(sheet2.getLastRow(), 1, 1, 9).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
          personal.getRange(personal.getLastRow(), 1, 1, 9).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
          sheet.getRange(e.range.rowStart, sheet.getLastColumn() - 1).setValue(payment);
        }
        for (m = 0; m < emaildata.length; m++) {
          if (emaildata[m][0] == nameResponse) {
            emailsheet.getRange(m + 1, dateworked.getDay() * 2 + 2, 1, 1).setValue(true);
          }
        }
        timestamp = new Date(responses[0])
        timestamp = String(Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MM/dd/yyyy")) + " " + formatTime(timestamp)
        body += `Timestamp: ${timestamp}  <br /> Name: ${responses[2]} <br /> Email: ${responses[1]} <br /> Date Worked: ${responses[3]} <br /> Job: ${responses[4]} <br />  ${moreinfo}<br /> Rate: ${ratemessage}  <br />Total Payment pre-tax: $${payment}<br /> Additional Comments and Notes: ${comments} `;
        email = responses[1];
        GmailApp.sendEmail(email, "Bnos Yisroel Payroll", "", {
          name: "Bnos Yisroel",
          htmlBody: body,
          //bcc: "aheyman@bnosyisroel.org,cdresdner@bnosyisroel.org"
        });
        // Append the result to the current row in a specific column (adjust column index)

      }
    }
  }

}
function triggerPayroll() {
  today = new Date()
  console.log(today.getDate())
  console.log(((new Date(today.getFullYear(), today.getMonth(), 15))))
  newdate = (new Date(today.getFullYear(), today.getMonth() + 1, 0)).getDate();
  console.log(newdate);
  //if (parseInt(today.getDate()) == parseInt(newdate) - 5){
  if (parseInt(today.getDate()) == 25) {
    console.log((new Date(today.getFullYear(), today.getMonth() + 1, 0)))
    payroll((new Date(today.getFullYear(), today.getMonth() + 1, 0)))
    console.log("test")
  }
  else if (parseInt(today.getDate()) == 10) {
    console.log("test")
    payroll((new Date(today.getFullYear(), today.getMonth(), 15)))
  }
}
function payroll(date) {
  today = Utilities.formatDate(date, 'America/New_York', 'MM/dd/yyyy');
  var spreadsheet = SpreadsheetApp.openById('1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU');
  var sheet = spreadsheet.getSheetByName('Form Responses 1');
  var rateSheet = spreadsheet.getSheetByName('Rates');
  sheetdata = sheet.getDataRange().getValues()
  ratedata = rateSheet.getDataRange().getValues()
  spreadsheet.insertSheet().setName(today);
  newsheet = spreadsheet.getSheetByName(today);
  found = false;
  todelete = -1
  newsheet.appendRow(["NAME", "DATE", "JOB", "DEPT", "HOURS", "RATE", "TOTAL", "NOTES"]);
  newdata = newsheet.getDataRange().getValues();
  for (i = sheetdata.length - 1; i > -1; i--) {
    if (sheetdata[i][sheet.getLastColumn() - 1] == "Payroll" && !found) {
      lastrow = i;
      console.log(String(lastrow) + " is last row")
      found = true;
    }
  }
  for (k = 1; k < ratedata.length; k++) {
    worked = false
    sum = 0;
    time = 0;
    line = ratedata[k];
    rate = line[3];
    type = line[2];
    job = line[1];
    dept = line[4];
    for (i = lastrow + 1; i < sheetdata.length; i++) {
      console.log(String(i) + " is sheet line")
      if ((line[0] == sheetdata[i][2]) && line[0] != "" && sheetdata[i][0] != "") {
        console.log("found")
        answers = sheetdata[i];
        answers = answers.filter(function (element) {
          return String(element).trim() !== "";
        });
        payment = 0
        if (line[1] == answers[4]) {
          worked = true
          console.log("found job")

          sum = sum + parseFloat(sheet.getRange(i + 1, sheet.getLastColumn() - 1).getValue())
          if (type == "Period") {
            var periodResponse = answers[5];
            time = time + answers[5]
            payment = parseInt(periodResponse) * parseInt(rate)
            newsheet.appendRow([(line[0]), answers[3], job, dept, answers[5], "$" + String(parseFloat(rate)), "$" + String(parseFloat(answers[answers.length - 1]))])
            if (todelete > -1) {
              newsheet.getRange(todelete, 1, 1, 1).clearContent()
              todelete = -1
            }
          }
          if (type == "Hour" || type == "Vacation") {
            hours = parseFloat(sheet.getRange(i + 1, sheet.getLastColumn() - 1).getValue()) / rate
            hours = Math.round((hours) * 100) / 100
            time = time + hours
            //numWithZeroes = num.toFixed(Math.max(((num+'').split(".")[1]||"").length, 2));
            newsheet.appendRow([(line[0]), answers[3], job, dept, String(parseFloat(hours)), "$" + String(parseFloat(rate)), "$" + String(parseFloat(answers[answers.length - 1]))])
            if (todelete > -1) {
              newsheet.getRange(todelete, 1, 1, 1).clearContent()
              todelete = -1
            }
          }
        }
      }
    }
    if (worked) {
      console.log(job)
      if (job == "HSLC") {
        console.log("HSLC")
        var individualSheet = SpreadsheetApp.openById(line[5]).getSheetByName('Sheet1');
        var individualTab = spreadsheet.getSheetByName(line[0]);
        individualSheet.appendRow([, "", (line[0]), today, "Paid Prep", Math.floor(time / 5)])
        individualTab.appendRow([, "", (line[0]), today, "Paid Prep", Math.floor(time / 5)])
        newsheet.appendRow([line[0], today, "Paid Prep", dept, Math.floor(time / 5), "$" + String(parseFloat(rate)), "$" + String(parseFloat(Math.floor(time / 5) * rate))])
        individualSheet.getRange(individualSheet.getLastRow(), 1, 1, 9).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
        individualTab.getRange(individualTab.getLastRow(), 1, 1, 9).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
        sum = sum + Math.floor(time / 5) * rate;
        time = time + Math.floor(time / 5)
      }
      newsheet.appendRow(["", "", "", "", time, "", "$" + String(parseFloat(sum))])
      newsheet.getRange(newsheet.getLastRow(), 1, 1, 7).setFontWeight("bold")
      newsheet.appendRow(["h"]);
      todelete = newsheet.getLastRow();
    }


  }
  if (todelete > -1) {
    newsheet.getRange(todelete, 1, 1, 1).clearContent()
    todelete = -1
  }
  newsheet.getRange(1, 1, 1, 9).setFontWeight("bold")

  newsheet.getDataRange().setFontSize(10).setFontFamily("Calibri").setHorizontalAlignment("left")
  sheet.getRange(sheet.getLastRow(), sheet.getLastColumn(), 1, 1).setValue("Payroll");
}

function formatSheets() {
  var spreadsheetId = '1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  // Get the submitted responses
  var rateSheet = spreadsheet.getSheetByName('Rates');
  rateSheetData = rateSheet.getDataRange().getValues()
  for (i = 1; i < rateSheetData.length; i++) {
    if (rateSheetData[i][2] == "Hour") {
      header = [["Timestamp", "Email Address", "Name", "Date worked", "Job", "Start Time", "End Time", "Hours Worked", "Notes"]]
    }
    else {
      header = [["Timestamp", "Email Address", "Name", "Date worked", "Job", "Start Time", "End Time", "Periods Worked", "Notes"]]
    }
    console.log(String(rateSheetData[i][5]))
    console.log(String(rateSheetData[i][0]))
    if ((String(rateSheetData[i][5]) == "" || (String(rateSheetData[i][5]) == "undefined")) && String(rateSheetData[i][0]) != String(rateSheetData[i - 1][0])) {
      var spreadsheet2 = SpreadsheetApp.create(String(rateSheetData[i][0]));
      rateSheet.getRange(i + 1, 6).setValue(spreadsheet2.getId());
      var file = DriveApp.getFileById(spreadsheet2.getId());
      var folder = DriveApp.getFolderById("1vNWqocE_xNf3Ke1fczQer-_SQBhNhfWk");
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
    if (String(rateSheetData[i][5]) == "" && String(rateSheetData[i][0]) == String(rateSheetData[i - 1][0])) {
      rateSheet.getRange(i + 1, 6).setValue(rateSheet.getRange(i, 6).getValue());
    }
    var sourceTab = spreadsheet.getSheetByName(String(rateSheetData[i][0]));
    if (!sourceTab) {
      sheetadd = spreadsheet.insertSheet(String(rateSheetData[i][0]));
    }
    if (sourceTab) {
      sourceTab.clear()
      sourceTab.getRange(1, 1, 1, 9).setValues(header)
      if (rateSheetData[i][2] == "Period") {
        sourceTab.hideColumns(6);
        sourceTab.hideColumns(7);
      }
      sourceTab.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
      sourceTab.getRange(1, 1, 1, 9).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan")

    }
    if (String(rateSheetData[i][5]) != "" && String(rateSheetData[i][5]) != "undefined") {
      var targetSheet = SpreadsheetApp.openById(rateSheetData[i][5]);

      var targetSheet = targetSheet.getSheetByName('Sheet1');
      targetSheet.clear()
      targetSheet.getRange(1, 1, 1, 9).setValues(header)
      if (rateSheetData[i][2] == "Period") {
        targetSheet.hideColumns(6);
        targetSheet.hideColumns(7);
      }
      targetSheet.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)
      targetSheet.getRange(1, 1, 1, 9).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan")


    }
  }
}
