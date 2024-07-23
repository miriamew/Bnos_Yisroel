function populateStudentClassesForTab(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = ss.getSheetByName('Categories');
  const studentClassesSheet = ss.getSheetByName('Student Classes');
  const termSheet = ss.getSheetByName(sheetName);
  
  
  // Set the headers for the Student Classes sheet
  const studentClassesHeader = [
    'Student', 'Category', 'Class', 'Year', 'Term', 
    'Fitness/Sports Taken', 'Fitness/Sports Remaining', 
    'Arts Taken', 'Arts Remaining', 
    'Skills Taken', 'Skills Remaining', 
    'Computers Taken', 'Computers Remaining', 
    'STEM Taken', 'STEM Remaining', 
    'Literacy Taken', 'Literacy Remaining', 
    'Music Taken', 'Music Remaining', 
    'Additional Taken', 'Additional Remaining'
  ];
  studentClassesSheet.appendRow(studentClassesHeader);
  
  // Read categories and classes mapping from Categories sheet
  const categoriesRange = categoriesSheet.getDataRange();
  const categoriesData = categoriesRange.getValues();
  
  const categoriesMap = new Map();
  const numRows = categoriesData.length;
  const numCols = categoriesData[0].length;
  
  for (let col = 0; col < numCols; col++) {
    const category = categoriesData[0][col];
    if (category) {
      categoriesMap.set(category, []);
      for (let row = 1; row < numRows; row++) {
        const className = categoriesData[row][col];
        if (className) {
          categoriesMap.get(category).push(className);
        }
      }
    }
  }
  
  // Extract year and term from sheetName
  const [year, term] = sheetName.split(' ');
  
  // Read data from the specified term sheet
  const range = termSheet.getDataRange();
  const data = range.getValues();
  
  const headers = data[0];
  const studentClassesMap = new Map();
  
  for (let col = 0; col < headers.length; col++) {
    const className = headers[col];
    if (className) {
      for (let row = 1; row < data.length; row++) {
        const student = data[row][col];
        if (student) {
          if (!studentClassesMap.has(student)) {
            studentClassesMap.set(student, {
              classes: [],
              counts: {
                'Fitness/Sports': 0, 'Arts': 0, 'Skills': 0, 
                'Computers': 0, 'STEM': 0, 'Literacy': 0, 'Music': 0, 
                'Additional': 0
              }
            });
          }
          
          let category = 'Uncategorized';
          for (const [cat, classes] of categoriesMap.entries()) {
            if (classes.includes(className)) {
              category = cat;
              break;
            }
          }
          
          studentClassesMap.get(student).classes.push({ className, category });
          if (studentClassesMap.get(student).counts[category] !== undefined) {
            studentClassesMap.get(student).counts[category]++;
          } else {
            studentClassesMap.get(student).counts['Additional']++;
          }
        }
      }
    }
  }
  
  // Define the requirements for each category
  const requirements = {
    'Fitness/Sports': 3, 'Arts': 3, 'Skills': 2, 
    'Computers': 2, 'STEM': 2, 'Literacy': 2, 'Music': 1, 
    'Additional': 3
  };
  
  // Adjust counts and allocate extra classes to Additional category
  for (const [student, info] of studentClassesMap.entries()) {
    const counts = info.counts;
    let additionalCount = counts['Additional'];
    
    for (const category in requirements) {
      if (counts[category] > requirements[category]) {
        additionalCount += counts[category] - requirements[category];
        counts[category] = requirements[category];
      }
    }
    
    counts['Additional'] = additionalCount;
  }
  
  // Write student classes data to the Student Classes sheet
  for (const [student, info] of studentClassesMap.entries()) {
    const classes = info.classes;
    classes.sort((a, b) => a.category.localeCompare(b.category));
    
    const counts = info.counts;
    const remaining = {};
    for (const [category, count] of Object.entries(counts)) {
      remaining[category] = Math.max(0, requirements[category] - count);
    }
    
    for (const classInfo of classes) {
      studentClassesSheet.appendRow([
        student, classInfo.category, classInfo.className, year, term,
        counts['Fitness/Sports'], remaining['Fitness/Sports'],
        counts['Arts'], remaining['Arts'],
        counts['Skills'], remaining['Skills'],
        counts['Computers'], remaining['Computers'],
        counts['STEM'], remaining['STEM'],
        counts['Literacy'], remaining['Literacy'],
        counts['Music'], remaining['Music'],
        counts['Additional'], remaining['Additional']
      ]);
    }
  }
}

// Trigger functions for each tab
function populateStudentClassesForTab1() {
  populateStudentClassesForTab('23-24 Term1');
}

function populateStudentClassesForTab2() {
  populateStudentClassesForTab('23-24 Term2');
}

function populateStudentClassesForTab3() {
  populateStudentClassesForTab('23-24 Term3');
}

function populateStudentClassesForTab4() {
  populateStudentClassesForTab('22-23 Term1');
}

function populateStudentClassesForTab5() {
  populateStudentClassesForTab('22-23 Term2');
}
function consolidateStudentRequirements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentClassesSheet = ss.getSheetByName('Student Classes');
  const consolidatedSheet = ss.getSheetByName('Consolidated Requirements') || ss.insertSheet('Consolidated Requirements');
  
  // Clear existing data in the Consolidated Requirements sheet
  consolidatedSheet.clearContents();
  
  // Set the headers for the Consolidated Requirements sheet
  const consolidatedHeader = [
    'Student', 
    'Fitness/Sports Taken', 'Fitness/Sports Remaining',
    'Arts Taken', 'Arts Remaining', 
    'Skills Taken', 'Skills Remaining', 
    'Computers Taken', 'Computers Remaining', 
    'STEM Taken', 'STEM Remaining', 
    'Literacy Taken', 'Literacy Remaining', 
    'Music Taken', 'Music Remaining', 
    'Additional Taken', 'Additional Remaining'
  ];
  consolidatedSheet.appendRow(consolidatedHeader);
  
  // Define the requirements for each category
  const requirements = {
    'Fitness/Sports': 3, 'Arts': 3, 'Skills': 2, 
    'Computers': 2, 'STEM': 2, 'Literacy': 2, 'Music': 1, 
    'Additional': 3
  };

  // Read data from the Student Classes sheet
  const range = studentClassesSheet.getDataRange();
  const data = range.getValues();
  
  const studentRequirementsMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const student = row[0];
    const category = row[1];

    if (!studentRequirementsMap.has(student)) {
      studentRequirementsMap.set(student, {
        'Fitness/Sports': 0,
        'Arts': 0,
        'Skills': 0,
        'Computers': 0,
        'STEM': 0,
        'Literacy': 0,
        'Music': 0,
        'Additional': 0
      });
    }

    const studentRequirements = studentRequirementsMap.get(student);

    if (category in studentRequirements) {
      studentRequirements[category]++;
    } else {
      studentRequirements['Additional']++;
    }
  }

  // Adjust counts and allocate extra classes to Additional category
  for (const [student, counts] of studentRequirementsMap.entries()) {
    let additionalCount = counts['Additional'];
    
    for (const category in requirements) {
      if (counts[category] > requirements[category]) {
        additionalCount += counts[category] - requirements[category];
        counts[category] = requirements[category];
      }
    }

    counts['Additional'] = additionalCount;
  }

  // Write consolidated data to the Consolidated Requirements sheet
  for (const [student, counts] of studentRequirementsMap.entries()) {
    const remaining = {};
    for (const category in requirements) {
      remaining[category] = Math.max(0, requirements[category] - counts[category]);
    }
    
    consolidatedSheet.appendRow([
      student,
      counts['Fitness/Sports'], remaining['Fitness/Sports'],
      counts['Arts'], remaining['Arts'],
      counts['Skills'], remaining['Skills'],
      counts['Computers'], remaining['Computers'],
      counts['STEM'], remaining['STEM'],
      counts['Literacy'], remaining['Literacy'],
      counts['Music'], remaining['Music'],
      counts['Additional'], remaining['Additional']
    ]);
  }
}
function updateStudentIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName("Student Classes");
  var idsSheet = ss.getSheetByName("Student IDs");

  var classesData = classesSheet.getDataRange().getValues();
  var idsData = idsSheet.getDataRange().getValues();

  // Assuming headers are in the first row, and data starts from the second row
  var classesHeaders = classesData[0];
  var idsHeaders = idsData[0];

  var nameIndex = classesHeaders.indexOf("Student");
  var idIndex = idsHeaders.indexOf("ID");

  // Iterate through each row in Student Classes sheet (starting from the second row)
  for (var i = 1; i < classesData.length; i++) {
    var studentName = classesData[i][nameIndex];
    if (studentName) {
      // Find matching student in Student IDs sheet
      for (var j = 1; j < idsData.length; j++) {
        if (idsData[j][2] === studentName) {
          // Update Student Classes sheet with ID from Student IDs sheet
          classesSheet.getRange(i + 1, nameIndex + 6).setValue(idsData[j][3]);
          break; // Exit inner loop once match is found
        }
      }
    }
  }
}
function organizeStudentClasses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName("Student Classes");

  var classesData = classesSheet.getDataRange().getValues();

  // Headers of the Student Classes sheet
  var headers = classesData[0];
  
  // Indices of relevant columns
  var nameIndex = headers.indexOf("Student Name");
  var categoryIndex = headers.indexOf("Category");
  var classIndex = headers.indexOf("Class");
  var yearIndex = headers.indexOf("Year");
  var termIndex = headers.indexOf("Term");
  var idIndex = headers.indexOf("Student ID");

  // Create a new sheet for organized data
  var organizedSheet = ss.insertSheet("Organized Classes");
  
  // Define categories and their required counts
  var categories = [
    { name: "Fitness/Sports", required: 3 },
    { name: "Arts", required: 3 },
    { name: "Skills", required: 2 },
    { name: "Computers", required: 2 },
    { name: "STEM", required: 2 },
    { name: "Literacy", required: 2 },
    { name: "Music", required: 1 }
  ];
  
  // Set headers for organized sheet
  var organizedHeaders = ["Student ID", "Student Name"];
  var categoryHeaders = [];
  
  // Add headers for each category
  categories.forEach(function(category) {
    organizedHeaders.push(category.name, "Classes Count");
    categoryHeaders.push(category.name);
  });
  
  // Add header for additional column
  organizedHeaders.push("Additional Classes");
  organizedSheet.getRange(1, 1, 1, organizedHeaders.length).setValues([organizedHeaders]);
  
  // Initialize organized data structure
  var organizedData = {};

  // Iterate through data in Student Classes sheet
  for (var i = 1; i < classesData.length; i++) {
    var studentID = classesData[i][idIndex];
    var studentName = classesData[i][nameIndex];
    var category = classesData[i][categoryIndex];
    var classInfo = classesData[i][classIndex] + " (" + classesData[i][termIndex] + " " + classesData[i][yearIndex] + ")";
    
    // Initialize student in organized data if not already
    if (!organizedData[studentID]) {
      organizedData[studentID] = {
        "name": studentName,
        "additionalClasses": 0
      };
      
      // Initialize categories for the student
      categories.forEach(function(category) {
        organizedData[studentID][category.name] = {
          "classes": [],
          "count": 0
        };
      });
    }
    
    // Check if the category exists in the organized data
    if (organizedData[studentID][category]) {
      // Add class info to the appropriate category for the student
      organizedData[studentID][category].classes.push(classInfo);
      organizedData[studentID][category].count++;
      
      // Check if they exceed the required amount
      if (organizedData[studentID][category].count > categories.find(cat => cat.name === category).required) {
        organizedData[studentID].additionalClasses++;
      }
    }
  }

  // Prepare data for insertion into the organized sheet
  var outputData = [];
  for (var id in organizedData) {
    var rowData = [id, organizedData[id].name];
    categories.forEach(function(category) {
      var classes = organizedData[id][category.name].classes || [];
      var count = organizedData[id][category.name].count || 0;
      var required = categories.find(cat => cat.name === category.name).required;
      
      rowData.push(classes.join(", "));
      rowData.push(count + "/" + required + " " + category.name + " classes taken");
    });
    
    // Add additional column for extra classes
    rowData.push(organizedData[id].additionalClasses);
    
    outputData.push(rowData);
  }

  // Insert data into the organized sheet
  organizedSheet.getRange(2, 1, outputData.length, organizedHeaders.length).setValues(outputData);
}
