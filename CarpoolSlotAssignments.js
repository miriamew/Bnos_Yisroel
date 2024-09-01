function createFormResultsTabs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formResultsSheet = spreadsheet.getSheetByName('Form Results'); // The sheet with form results
  const divisions = ['PS', 'ES', 'MS', 'HS'];
  
  // Define column indices
  const dismissalColumn = 1; // Adjust if necessary
  const typeColumn = 2; // Adjust if necessary
  const walkerColumn = 11; // Adjust if necessary
  const studentColumns = [12, 13, 14, 15, 16, 17]; // Adjust if necessary
  
  const data = formResultsSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers from data

  divisions.forEach(division => {
    // Create or clear the division sheet
    let divisionSheet = spreadsheet.getSheetByName(division + ' Form Results');
    if (!divisionSheet) {
      divisionSheet = spreadsheet.insertSheet(division + ' Form Results');
    } else {
      divisionSheet.clear(); // Clear existing data if the sheet already exists
    }
    divisionSheet.appendRow(headers); // Add headers to the new sheet
    
    // Filter and add unique rows to the division sheet
    const addedRows = new Set();
    data.forEach(row => {
      const dismissalDate = row[dismissalColumn - 1];
      const dismissalTime = formatTime(dismissalDate); // Extract time part
      const type = row[typeColumn - 1];
      const walker = row[walkerColumn - 1];
      const students = studentColumns.map(col => row[col - 1]).filter(cell => cell);
      
      // Determine division based on formatted time
      if ((dismissalTime === '05:15' && division === 'PS') ||
          (dismissalTime === '06:45' && division === 'ES') ||
          (dismissalTime === '07:15' && division === 'MS') ||
          (dismissalTime === '07:55' && division === 'HS')) {
        
        // Create a unique identifier for the form result
        const rowId = row.join('|'); // Unique identifier based on row content
        if (!addedRows.has(rowId)) {
          divisionSheet.appendRow(row);
          addedRows.add(rowId);
        }
      }
    });
  });
}

function formatTime(date) {
  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

function updateCarpoolSlots() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const divisions = ['PS', 'ES', 'MS', 'HS'];
  const slots = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z','1','2','3','4','5','6','7','8','9','10','11','12','13', '14', '15', '16','17','18']; // Define possible slots

  // Define column indices for the form results and existing sheets
  const walkerColumn = 11; // Adjust if necessary
  const studentColumns = [12, 13, 14, 15, 16, 17]; // Adjust if necessary
  const largeVanColumn = 3; // Adjust if necessary for the 'Large Van' column
  const commentsColumn = 18; // Adjust if necessary for the comments column

  // Function to extract last name, first name, and grade from student string
  function parseStudentInfo(studentString) {
    const parts = studentString.split(',');
    const lastName = parts[0].trim();
    const rest = parts[1] ? parts[1].trim() : '';
    var grade = rest.split(' ').pop(); // Extract the last word as the grade
    const firstName = rest.substring(0, rest.lastIndexOf(' ')).trim(); // The remaining part is the first name
    if (grade == "SN"){grade = "Senior Nursery"}
    if (grade == "JN"){grade = "Junior Nursery"}
    return { lastName, firstName, grade };
  }

  // Function to generate carpool slot based on existing slots and large van requirement
  function getCarpoolSlot(existingSlots, largeVan) {
    console.log(existingSlots)
    const regularSlots = slots.slice(0, 14); // Slots before 'P'
    const largeVanSlots = slots.slice(14); // Slots from 'P' onward
    console.log(regularSlots)
    console.log(largeVanSlots)
    let availableSlots = largeVan ? [] : regularSlots; // Initialize available slots
    let alternateAvailableSlots = []; // Initialize alternate available slots
    
    if (largeVan) {
      // Alternate large van slots
      alternateAvailableSlots = largeVanSlots.filter((slot, index) => index % 2 === 0);
      availableSlots = alternateAvailableSlots.concat(
        largeVanSlots.filter((slot, index) => index % 2 !== 0)
      );
    } else {
      availableSlots = regularSlots.concat(largeVanSlots);
    }

    for (let i = 0; i < availableSlots.length; i++) {
      const slot = availableSlots[i];
      if (!existingSlots.has(slot)) {
        return slot;
      }
    }
    return ''; // Return empty if no slot is available
  }

  // Iterate over each division
  divisions.forEach(division => {
    const formResultsSheetName = division + ' Form Results';
    const existingSheetName = division; // Existing tab with student info

    // Get the form results sheet for the division
    const formResultsSheet = spreadsheet.getSheetByName(formResultsSheetName);
    const existingSheet = spreadsheet.getSheetByName(existingSheetName);

    if (formResultsSheet && existingSheet) {
      const formResultsData = formResultsSheet.getDataRange().getValues();
      const existingDataRange = existingSheet.getDataRange();
      const existingData = existingDataRange.getValues();
      
      let carpoolSlotIndex = existingData[0].indexOf('Carpool Slot'); // Find the column index for Carpool Slot
      let commentsIndex = existingData[0].indexOf('Comments'); // Find the column index for Comments

      // Ensure there's a Carpool Slot column in the existing sheet
      if (carpoolSlotIndex === -1) {
        // Add headers if not present
        existingSheet.getRange(1, existingData[0].length + 1).setValue('Carpool Slot');
        carpoolSlotIndex = existingData[0].length; // Update index after adding the column
      }
      
      // Ensure there's a Comments column in the existing sheet
      if (commentsIndex === -1) {
        // Add Comments header if not present
        existingSheet.getRange(1, existingData[0].length + 1).setValue('Comments');
        commentsIndex = existingData[0].length; // Update index after adding the column
      }

      // Create a map to track carpool slot assignments and comments
      const studentUpdates = new Map();
      const existingSlots = new Set(); // To keep track of assigned slots

      formResultsData.forEach(row => {
        const type = row[1]; // Carpool type
        const walker = row[walkerColumn - 1]; // Walking authorization
        const largeVan = row[largeVanColumn - 1] === 'Yes'; // Large Van
        const comments = row[commentsColumn - 1]; // Comments
        const students = studentColumns.map(col => row[col - 1]).filter(cell => cell);
        
    
        if (type === 'Carpool') {
          const carpoolSlot = getCarpoolSlot(existingSlots, largeVan);
          console.log(carpoolSlot)
          if (carpoolSlot) {
            students.forEach(student => {
              console.log(student)
              const { lastName, firstName, grade } = parseStudentInfo(student);
              const studentId = `${lastName}|${firstName}|${grade}`;
              studentUpdates.set(studentId, { slot: carpoolSlot, comment: comments });
            });
            existingSlots.add(carpoolSlot); // Add the assigned slot to the set
          }
        } else {
          students.forEach(student => {
            const { lastName, firstName, grade } = parseStudentInfo(student);
            const studentId = `${lastName}|${firstName}|${grade}`;
            if (!studentUpdates.has(studentId)) {
              studentUpdates.set(studentId, { slot: type, comment: comments }); // Assign type to Carpool Slot and add comments
            } else {
              // Update comments if the student already has a slot assignment
              const existingEntry = studentUpdates.get(studentId);
              studentUpdates.set(studentId, { slot: existingEntry.slot, comment: `${existingEntry.comment} | ${comments}` });
            }
          });
        }
      });

      // Update the Carpool Slot and Comments columns in the existing sheet
      for (let i = 1; i < existingData.length; i++) {
        const existingRow = existingData[i];
        const lastName = existingRow[0];
        const firstName = existingRow[1];
        const grade = existingRow[3];
        const studentId = `${lastName}|${firstName}|${grade}`;
        
        if (studentUpdates.has(studentId)) {
          const { slot, comment } = studentUpdates.get(studentId);
          existingSheet.getRange(i + 1, carpoolSlotIndex + 1).setValue(slot);
          existingSheet.getRange(i + 1, commentsIndex + 1).setValue(comment);
        }
      }
    }
  });
}
function updateCarpoolInfoForDifferentDivisions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const divisions = ['PS', 'ES', 'MS', 'HS'];
  const timeToDivision = {
    '05:15': 'PS',
    '06:45': 'ES',
    '07:15': 'MS',
    '07:55': 'HS'
  };

  // Define column indices for student data in Form Results
  const studentColumns = [12, 13, 14, 15, 16, 17]; // Adjust based on the actual columns

  // Function to extract last name, first name, and grade from student string
  function parseStudentInfo(studentString) {
    const parts = studentString.split(',');
    const lastName = parts[0].trim();
    const rest = parts[1] ? parts[1].trim() : '';
    var grade = rest.split(' ').pop(); // Extract the last word as the grade
    const firstName = rest.substring(0, rest.lastIndexOf(' ')).trim(); // The remaining part is the first name
    if (grade == "SN"){grade = "Senior Nursery"}
    if (grade == "JN"){grade = "Junior Nursery"}
    return { lastName, firstName, grade };
  }

  // Create a map to track students and their carpool times
  const studentCarpoolMap = new Map();

  // Get the Form Results sheet and populate the studentCarpoolMap
  const formResultsSheet = spreadsheet.getSheetByName('Form Results');
  if (formResultsSheet) {
    const formResultsData = formResultsSheet.getDataRange().getValues();
    const timeColumnIndex = 0; // Column A (0-based index is 0)

    // Fill the map with student info from Form Results
    formResultsData.forEach(row => {
      console.log(row[timeColumnIndex])
      const time = formatTime(row[timeColumnIndex]);
      const type = row[1]
      console.log(type)
      const students = studentColumns.map(col => row[col - 1]).filter(cell => cell);

      students.forEach(student => {
        const { lastName, firstName, grade } = parseStudentInfo(student);
        const studentId = `${lastName}|${firstName}|${grade}`;
        if (!studentCarpoolMap.has(studentId) && timeToDivision[time]) {
          studentCarpoolMap.set(studentId, timeToDivision[time] + "-" +type);
        }
      });
    });
  }

  // Log the studentCarpoolMap for debugging
  console.log('Student Carpool Map:', Array.from(studentCarpoolMap.entries()));

  // Iterate over each division
  divisions.forEach(division => {
    const divisionSheet = spreadsheet.getSheetByName(division);

    if (divisionSheet) {
      const divisionData = divisionSheet.getDataRange().getValues();

      // Get the indices for Carpool Slot (Column E)
      const carpoolSlotIndex = 4; // Column E (0-based index is 4)

      // Update the division sheet
      for (let i = 1; i < divisionData.length; i++) { // Skip header row
        const existingRow = divisionData[i];
        const lastName = existingRow[0];
        const firstName = existingRow[1];
        const grade = existingRow[3];
        const studentId = `${lastName}|${firstName}|${grade}`;
        
        if (!existingRow[carpoolSlotIndex]) { // Check if Carpool Slot is blank
          if (studentCarpoolMap.has(studentId)) {
            const divisionWithCarpool = studentCarpoolMap.get(studentId);
            divisionSheet.getRange(i + 1, carpoolSlotIndex + 1).setValue(divisionWithCarpool);
          } 
        }
      }
    }
  });
}
