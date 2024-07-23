function formtest(){
  // Open a form by ID and log the responses to each question.
var form = FormApp.openById('1nQfaAIQ75yV-qjbxdGHjlYYLl5lFOtKJpcRANIl_FPk');
var formResponses = form.getResponses();
for (var i = formResponses.length-1; i < formResponses.length; i++) {
  var formResponse = formResponses[i];
  var itemResponses = formResponse.getItemResponses();
  for (var j = 0; j < itemResponses.length; j++) {
    var itemResponse = itemResponses[j];
        if (itemResponse.getItem().getTitle() == "Additional Comments and Notes"){
          console.log(itemResponse.getResponse());
        }
        if (itemResponse.getItem().getTitle() == "If you arrived late, how many minutes late were you? "){
          console.log("late" +itemResponse.getResponse());
        }

  }
}
}
function play (){

  const person = {}
  taughtM = true
  array = [taughtM]

  person.taughtM = true;
  array.push(person.taughtM)
  person.taughtM = false;
  console.log(person.array[0])
}
function formatTime(time){
  //time = new Date(time)
  hours = (time.getHours()) 
  console.log(hours)
  minutes = (time.getMinutes())
  s = " am"
  if (hours > 12){
    console.log("afternoon")
    hours = hours - 12;
    s = " pm"
  } 
  if (hours == 12){
    console.log("noon");
    s=" pm";}
  if (minutes < 10){
    minutes = "0" + String(minutes);
  }
  if (hours == 0){
    hours = 12
  }
  console.log((String(hours) + ":" + String(minutes) + s))
  return (String(hours) + ":" + String(minutes) + s)
 


  }
function emails(){
  rightnow = new Date();
  var spreadsheetId = '1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var emailsheet = spreadsheet.getSheetByName('Emails');
  emaildata = emailsheet.getDataRange().getValues();
  time = (String(rightnow.getHours()));
  day = rightnow.getDay();
  console.log(day)
  
  if (day >0 && day <6 && time === "6"){
    for (i = 1; i < emaildata.length+1;i++){
      emailsheet.getRange(i+1, 4,1,1).setValue(false);
      emailsheet.getRange(i+1, 6,1,1).setValue(false);
      emailsheet.getRange(i+1, 8,1,1).setValue(false);
      emailsheet.getRange(i+1, 10,1,1).setValue(false);
      emailsheet.getRange(i+1, 12,1,1).setValue(false);
    }

  }
  if (day >0 && day < 6 && time === "17"){
    for (i = 1; i < emaildata.length+1;i++){
      email = emailsheet.getRange(i+1, 2,1,1).getValues();
      teaches = emailsheet.getRange(i+1, ((day)*2)+1,1,1).getValues();
      taught = emailsheet.getRange(i+1, (day)*2+2,1,1).getValues();
      console.log(email[0][0])
      if (teaches[0][0] == true && taught[0][0] == false){
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
   
  console.log(responses)
  var sheet = spreadsheet.getSheetByName('Form Responses 1');
  var rateSheet = spreadsheet.getSheetByName('Rates');
  var emailsheet = spreadsheet.getSheetByName('Emails');
  emaildata = emailsheet.getDataRange().getValues();
  responses = responses.filter(function(element) {
  // Remove elements that are empty or contain only whitespace
  return element.trim() !== "";
});

console.log(responses);
  
  var nameResponse =responses[2];
  var jobResponse = responses[4];
  
  rate = 0
  type = ""
  payment = 0
  minlate = 0
  comments = "None"
  comments2 = ""
  ratedata = rateSheet.getDataRange().getValues()
  console.log(ratedata)
  for (i = 0; i < ratedata.length; i++){
    line = ratedata[i];
    console.log(line[0])
    spreadsheetid = line[5];
    console.log(line[1])
    console.log(nameResponse)
    console.log(jobResponse)
    if (line[0] == nameResponse){
      if(line[1] == jobResponse){
        console.log("found job")
        rate = line[3];
        type = line[2];
        var spreadsheet2 = SpreadsheetApp.openById(spreadsheetid);
        var sheet2 = spreadsheet2.getSheetByName("Sheet1");
        var personal = spreadsheet.getSheetByName(nameResponse);

        personinfo = []
        for (pi = 0; pi<responses.length;pi++){
          personinfo.push(responses[pi])
        }
        console.log(personinfo)
        sheet2.appendRow(personinfo)
        personal.appendRow(personinfo)
        sheet2.getRange(sheet2.getLastRow(), 1, 1, 8).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)    
        sheet2.autoResizeColumn(7);
        sheet2.autoResizeColumn(8);
        personal.getRange(personal.getLastRow(), 1, 1, 8).setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)    
        personal.autoResizeColumn(7);
        personal.autoResizeColumn(8);
      }
    }
  }
  console.log(rate);
  
  console.log(type);
  if (type == "Period"){
    console.log("in periods")
    var periodResponse =responses[5];
      var formResponses = form.getResponses();
      for (var i = formResponses.length-1; i < formResponses.length; i++) {
        var formResponse = formResponses[i];
        var itemResponses = formResponse.getItemResponses();
        for (var j = 0; j < itemResponses.length; j++) {
          var itemResponse = itemResponses[j];
              if (itemResponse.getItem().getTitle() == "Additional Comments and Notes"){
                comments = String(itemResponse.getResponse());
              }
              if (itemResponse.getItem().getTitle() == "If you arrived late, how many minutes late were you? " && parseFloat(itemResponse.getResponse()) != NaN){
                console.log("in here")
                minlate = (itemResponse.getResponse());
                console.log(minlate)
                comments2 = "<br />Minutes arrived late: "+minlate
                periodResponse = periodResponse -(minlate/42)
                periodResponse = Math.round((periodResponse)*100)/100;
              }

        }
      }
   
    console.log(responses[6])
    ratemessage = "$" + String(rate)+"/Period"
    
    foremail = String(responses[5]) + " Period(s)"
    payment = parseFloat(periodResponse) * parseFloat(rate)
    error = ""
    moreinfo = `Total Period(s): ${periodResponse} `
  }
  if (type == "Hour"){
    console.log("in hours")
    comments = String(responses[7])
    ratemessage = "$" + String(rate)+"/hr"
    var inTimeResponse =responses[5];
    var outTimeResponse =responses[6];
    var time1 = new Date("January 1, 2000 " + inTimeResponse);
    var time2 = new Date("January 1, 2000 " + outTimeResponse);
    var timeDifference = time2 - time1;
    var hours = Math.round((timeDifference / 3600000)*100)/100;
    payment = (hours * parseFloat(rate));
    payment = Math.round(payment * 100) / 100
    foremail = String(hours) + " Hour(s)"
    error = ""
    errorMsg = ""
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
  
  console.log(comments)
  if (String(comments) == "undefined"){
    comments = "None"
  }
  body = `<p>Regarding form submitted by: ${nameResponse}</p>`
  body = "<p>Thank you for submitting your payroll form. As a reminder, forms submitted by the 10th  of the month will be added to the payroll on the 15th, forms submitted by the 25th of the month will be added to the payroll on the last day of the month. <br /> Thank you for all your hard work!</p><br />";
  
  if (error != "") {
    body = `<p>There was an error with your form. Please resubmit the <a href=https://docs.google.com/forms/d/e/1FAIpQLSdnmYmN39J4BGt7SxlZGY0Kjtcu9Bt_dodAMezxYzGnIivTWw/viewform>PAYROLL FORM.</a> Thank you!<br /><br /> Error(s): <ol>${errorMsg}</ol> <br /></p>`;
    payment = 0;
    
  }
  timestamp = new Date(responses[0])
  for(m = 0; m<emaildata.length;m++){
    if (emaildata[m][0] == nameResponse){
      emailsheet.getRange(m+1, timestamp.getDay() *2+2,1,1).setValue(true);
    }
  }
  timestamp = String(Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MM/dd/yyyy")) +" "+ formatTime(timestamp)
  body += `Timestamp: ${timestamp}  <br /> Name: ${responses[2]} <br /> Email: ${responses[1]} <br /> Date Worked: ${responses[3]} <br /> Job: ${responses[4]} <br />  ${moreinfo}<br /> Rate: ${ratemessage}  <br />Total Payment pre-tax: $${payment}<br /> Additional Comments and Notes: ${comments} `;

  console.log(payment);
  email = responses[1];
  //GmailApp.sendEmail(email, "Bnos Yisroel Payroll", "", {
  //name: "Bnos Yisroel",
  //htmlBody: body,
  //bcc: "aheyman@bnosyisroel.org,cdresdner@bnosyisroel.org"
  //});

  // Append the result to the current row in a specific column (adjust column index)
  sheet.getRange(e.range.rowStart, sheet.getLastColumn()- 1).setValue(payment);
  
}
function triggerPayroll(){
  today = new Date()
  console.log(today.getDate())
  console.log(((new Date(today.getFullYear(), today.getMonth(), 15))))
  newdate = (new Date(today.getFullYear(), today.getMonth()+1, 0)).getDate();
  console.log(newdate);
  //if (parseInt(today.getDate()) == parseInt(newdate) - 5){
  if (parseInt(today.getDate()) == 25){  
    console.log((new Date(today.getFullYear(), today.getMonth() +1, 0)))
    payroll((new Date(today.getFullYear(), today.getMonth()+1, 0)))
    console.log("test")
  }
  else if (parseInt(today.getDate()) == 10){
    console.log("test")
    payroll((new Date(today.getFullYear(), today.getMonth(), 15)))
  }
}
function payroll(date) {
  today = date
  today = Utilities.formatDate(today, 'America/New_York', 'MM/dd/yyyy');
  var spreadsheetId = '1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  // Get the submitted responses
  var sheet = spreadsheet.getSheetByName('Form Responses 1');
  var rateSheet = spreadsheet.getSheetByName('Rates');
  sheetdata = sheet.getDataRange().getValues()
  console.log(sheetdata)
  ratedata = rateSheet.getDataRange().getValues()
  spreadsheet.insertSheet().setName(today);
  newsheet = spreadsheet.getSheetByName(today);
  found = false;
  todelete = -1
  newsheet.appendRow(["NAME",	"DATE",	"JOB",	"DEPT",	"HOURS",	"RATE",	"TOTAL", 	"NOTES"]);
  newdata = newsheet.getDataRange().getValues();
  for (i = sheetdata.length - 1; i > -1; i--){
        if (sheetdata[i][sheet.getLastColumn()-1] ==  "Payroll" && !found){
          lastrow = i;
          console.log(String(lastrow) + " is last row")
          found = true;
        }
      }
    for (k = 1; k < ratedata.length; k++){
      sum = 0;
      time = 0;
      
      line = ratedata[k];
      
      for (i = lastrow+1 ;i < sheetdata.length;  i++){
          console.log(String(i) + " is sheet line")
            if((line[0] == sheetdata[i][2]) && line[0] !=""&& sheetdata[i][0] !=""){
              console.log("found")
              answers = sheetdata[i];
              answers = answers.filter(function(element) {
                return String(element).trim() !== "";
                });
              rate = 0
              type = ""
              payment = 0  
                  if(line[1] == answers[4]){
                    console.log("found job")
                    rate = line[3];
                    type = line[2];
                    job = line[1];
                    dept = line[4];
                    sum = sum+parseFloat(sheet.getRange(i+1, sheet.getLastColumn()-1).getValue())
                  
              if (type == "Period"){
                var periodResponse =answers[5];
                time = time + answers[5]
                payment = parseInt(periodResponse) * parseInt(rate)
                newsheet.appendRow([(line[0]) ,answers[3], job, dept,answers[5], "$" + String(parseFloat(rate)), "$" + String(parseFloat(answers[answers.length-1])) ])
                if (todelete > -1 ){ 
                  newsheet.getRange(todelete,1,1,1).clearContent()
                  todelete = -1
                  }
                if (job == "HS Learning Center"){
                  extraperiods = Math.floor(periodResponse /5);
                  sum = sum + extraperiods*rate;
                }
              }
              if (type == "Hour"){
                hours = parseFloat(sheet.getRange(i+1, sheet.getLastColumn()-1).getValue())/rate
                time = time+hours
                //numWithZeroes = num.toFixed(Math.max(((num+'').split(".")[1]||"").length, 2));
                newsheet.appendRow([(line[0]),answers[3], job, dept,String(parseFloat(hours)), "$" + String(parseFloat(rate)), "$" + String(parseFloat(answers[answers.length-1])) ])
                if (todelete > -1 ){ 
                  newsheet.getRange(todelete,1,1,1).clearContent()
                  todelete = -1
                }
              }
              }
              }
            }
        newdata = newsheet.getDataRange().getValues();
        totalled = false;
        for (r = 0; r< newdata.length; r++){
          console.log("hi")
          
          if (line[0] == newdata[r][0] && !totalled){
            newsheet.appendRow(["","","","",time,"","$" + String(parseFloat(sum))])
            newsheet.getRange(newsheet.getLastRow(),1,1,7).setFontWeight("bold")
            totalled = true;
            newsheet.appendRow(["h"]);
            todelete = newsheet.getLastRow();
        }
      }
    }    
    if (todelete > -1 ){ 
                  newsheet.getRange(todelete,1,1,1).clearContent()
                  todelete = -1
                  }
  newsheet.getRange(1,1,1,8).setFontWeight("bold")
  
  newsheet.getDataRange().setFontSize(10).setFontFamily("Calibri").setHorizontalAlignment("left")
  sheet.getRange(sheet.getLastRow(),sheet.getLastColumn(),1,1).setValue("Payroll");
}

function formatSheets(){
  var spreadsheetId = '1rhBuITb0CwJ_0dPj1SdhQu90JsmOb4V_2wsKvXJPhiU';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  // Get the submitted responses
  var rateSheet = spreadsheet.getSheetByName('Emails');
  rateSheetData = rateSheet.getDataRange().getValues()
  for(i = 1; i<rateSheetData.length;i++){
    if(String(rateSheetData[i][5]) != ""){
      console.log(rateSheetData[i][12])
        var targetSheet = SpreadsheetApp.openById(rateSheetData[i][12]); 
        var targetSheet = targetSheet.getSheetByName('Sheet1'); 
        targetSheet.getDataRange().setFontFamily("Calibri").setHorizontalAlignment("left").setBorder(true, true, true, true, true, true)     
        targetSheet.getRange(1, 1, 1, 8).setHorizontalAlignment("center").setFontWeight("bold").setBackground("cyan")
        targetSheet.autoResizeColumn(7);
        targetSheet.autoResizeColumn(8);


    }
  }
}
