
function importXlsxFromEmail2() {
  const emailLabelName = "bloomerang"; // Change this to your label name
  const targetFolderId = "1wm2YYavLugH6MlOpB9tAwlEvWdDNDv-Y"; // Replace with your target folder ID
  const emailLabel = GmailApp.getUserLabelByName(emailLabelName);
  if (!emailLabel) {
    Logger.log(`Label "${emailLabelName}" not found.`);
    return;
  }
  const threads = GmailApp.search(`label:${emailLabelName}`);
  if (threads.length === 0) {
    Logger.log("No emails found with the specified label.");
    return;
  }
  const thread = threads[0];
  const messages = thread.getMessages();
  for (let j = 0; j < messages.length; j++) {
    const message = messages[j];
    const attachments = message.getAttachments();
    const emailSubject = message.getSubject();
    for (let i = 0; i < attachments.length; i++) {
        const blob = attachments[i].copyBlob(); // Use copyBlob() to get the blob without conversion
        const file = DriveApp.createFile(blob).setName(emailSubject);
        Logger.log(`Temporary file created: ${file.getId()}`);
        const folder = DriveApp.getFolderById(targetFolderId);
        const sheetFile = DriveApp.getFileById(file.getId());
      folder.addFile(sheetFile);
        DriveApp.getRootFolder().removeFile(sheetFile);
        Logger.log(`File added to folder with ID: ${targetFolderId}`);
        message.moveToTrash();
        return
       
    }
  
  }
}
function updateBloomerangInfo(){
  var spreadsheet = SpreadsheetApp.openById("1uNr5GlXqjI4SCKgVveBpS45z1L4P9HYhFeHX32DvN4M");
  var reportSheet = spreadsheet.getSheetByName("Trend Data");
  var reportSheetData = reportSheet.getDataRange().getValues();
  var sourcespreadsheet = SpreadsheetApp.openById("1avWCeq4oPCREfIWuSklJMoyYI68eC-tSDc5Si8EKiwI");
  var sourceSheet = sourcespreadsheet.getSheetByName("Data");
  var sourceSheetData = sourceSheet.getDataRange().getValues();
  for (let k = 1; k < reportSheetData.length; k++) {
    line = reportSheetData[k]
    sourceName = String(sourcespreadsheet.getName()).substring(0, sourcespreadsheet.getName().length-5)
    //console.log(sourceName)
    if (sourceName == line[0]){
      if (line[0] == "Income from Endowment Draws"){
        console.log("Income from Endowment Draws")
        console.log(line[34])
      }
      if (line[0] == "($1K+) Supporters"){
        console.log("($1K+) Supporters")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($0-$999.99) Friends"){
        console.log("($0-$999.99) Friends")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($2,500+) Benefactors"){
        console.log("($2,500+) Benefactors")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($10K+) Sustainers"){
        console.log("($10K+) Sustainers")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
       console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($25K+) Visionaries"){
        console.log("($25K+) Visionaries")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($5K+) Partners"){
        console.log("($5K+) Partners")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "($50K+) Founders"){
        console.log("($50K+) Founders")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Capital Gifts"){
        console.log("Capital Gifts")
        total = sourceSheetData[1][3]
        donors = sourceSheet.getLastRow() - 4
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Total Annual + Scholarship Support"){
        console.log("Total Annual + Scholarship Support")
        donations = 0
        donors = 0
        for (row = 1; row<sourceSheetData.length;row++){
          if (String(sourceSheetData[row][1]) != "" ){
          donors ++;
          donations = donations + sourceSheetData[row][3]
          }
          
        }
        
        reportSheet.getRange(k+2, 38, 1,1).setValues([[donations]])
      }
      if (line[0] == "Pledges"){
        console.log("Pledges")
        donations = sourceSheetData[1][2]
        donors = 0
        for (row = 1; row<sourceSheetData.length;row++){
          if (String(sourceSheetData[row][1]) != "" ){
            console.log(String(sourceSheetData[row][1]))
          donors ++;
          }
          
        }
        
        reportSheet.getRange(k+1, 38, 1,1).setValues([[donations]])
        reportSheet.getRange(k+3, 38, 1,1).setValues([[donors]])
      }
      if (line[0] == "Eruv Community"){
        console.log("Eruv Community")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Former Board Members"){
        console.log("Former Board Members")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Board"){
        console.log("Board")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "GP Alumni"){
        console.log("GP Alumni")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Transactional Income"){
        console.log("Transactional Income")
        total = sourceSheetData[1][2]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Gifts/Pledges to Endowment"){
        console.log("Gifts/Pledges to Endowment")
        total = sourceSheetData[1][3]
        donors = sourceSheet.getLastRow() - 6
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Income from Endowment Draws"){
        console.log("Income from Endowment Draws")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Jewish Community"){
        console.log("Jewish Community")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
       console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Alumnae"){
        console.log("Alumnae")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Students"){
        console.log("Students")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "GPs"){
        console.log("GPs")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
       console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "New Capital Commitments"){
        console.log("New Capital Commitments")
        total = sourceSheetData[1][3]
        donors = sourceSheet.getLastRow() - 4
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "New Endowment Commitments"){
        console.log("New Endowment Commitments")
        total = sourceSheetData[1][3]
        donors = sourceSheet.getLastRow() - 4
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Other Donors"){
        console.log("Other Donors")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Parent Alumni"){
        console.log("Parent Alumni")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
       console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Parents"){
        console.log("Parents")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Restricted Giving"){
        console.log("Restricted Giving")
        total = sourceSheetData[1][3]
        donors = 0
        for (row = 1; row<sourceSheetData.length;row++){
          if (sourceSheetData[row][1] != "")
          donors ++;
        }
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
      if (line[0] == "Staff"){
        console.log("Staff")
        total = sourceSheetData[1][1]
        donors = sourceSheet.getLastRow() - 2
        console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
        reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
        reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])

      }
    }
  if (sourceName == "Annual Campaign by Appeal"){
        donors = 0
        total = 0
        //console.log("Annual Campaign by Appeal")
        for (row = 1; row < sourceSheetData.length;row++){
          if(sourceSheetData[row][0] == "" && sourceSheetData[row][1] != ""){
            sourceSheetData[row][0] = "None"
          }
          if((sourceSheetData[row][0] == line[0]) && (line[0]!="") && (line[0]!= "Charidy Day")){
          donors = donors+1
          total = total + sourceSheetData[row][3]
          //console.log(donors)
          //console.log(total)
          console.log(reportSheet.getRange(k+1, 38, 1,1).getValues())
          reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
          reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])
          }
          if((String(sourceSheetData[row][0]).substring(0,13) == "Charidy Day (" ) && (line[0]== "Charidy Day")){
            console.log(String(sourceSheetData[row][0]).substring(13,String(sourceSheetData[row][0]).length-1))
            donors = String(sourceSheetData[row][0]).substring(13,String(sourceSheetData[row][0]).length-1)
            console.log(sourceSheetData[row+1][3])
            total = sourceSheetData[row+1][3]
            reportSheet.getRange(k+1, 38, 1,1).setValues([[total]])
            reportSheet.getRange(k+2, 38, 1,1).setValue([[donors]])
          }
        }
        }

  }

}
function createTrigger() {
  var allTriggers = ScriptApp.getProjectTriggers();
  //console.log(allTriggers.length) 
    ScriptApp.newTrigger('importXlsxFromEmail2')
      .timeBased()
      .everyMinutes(10)
      .create();
    ScriptApp.newTrigger('updateBloomerangInfo')
      .timeBased()
      .everyMinutes(10)
      .create();
  
}

function deleteAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if((allTriggers[i].getHandlerFunction())=="importXlsxFromEmail2"){
    ScriptApp.deleteTrigger(allTriggers[i]);

      }
      if((allTriggers[i].getHandlerFunction())=="updateBloomerangInfo"){
    ScriptApp.deleteTrigger(allTriggers[i]);

      }
  }
}


