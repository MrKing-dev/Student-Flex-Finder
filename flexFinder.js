
//Example first lines
//Cycle Day: Day 3
//Instructor: Brown

function onOpen() {
  var ui = DocumentApp.getUi(); // Get the UI of the document
  ui.createMenu('Custom Menu')  // Add a custom menu
    .addItem('Fetch Students', 'fetchStudentsFromDocInputs')  // Add a menu item that calls the script
    .addToUi();
}

//1GoBtzm-RBtl7KsAwQWxpPqlnI3bB5mQ8F5hfSLgOtlw
//Student Schedules

function fetchStudentsFromDocInputs() {
  var sheetId = "1GoBtzm-RBtl7KsAwQWxpPqlnI3bB5mQ8F5hfSLgOtlw"; // Replace with your actual sheet ID
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Student Schedules"); // Change if needed
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody().getText(); // Get all text from the document

  // Extract values from the document text
  var dayMatch = body.match(/Cycle Day:\s*(Day \d+)/);
  var instructorMatch = body.match(/Instructor:\s*(.*)/);  // Match everything after "Instructor:"

  if (!dayMatch || !instructorMatch) {
    Logger.log("Could not find cycle day or instructor in the document.");
    return;
  }

  var day = dayMatch[1];
  var instructorString = instructorMatch[1].trim();  // Get the instructor(s) string

  // Split the instructor string into an array of possible names (separated by commas or semicolons)
  var instructors = instructorString.split(/,|;/).map(name => name.trim().toLowerCase());  // Convert to lowercase

  // Define which columns correspond to each cycle day
  var dayColumns = {
    "Day 1": ["F", "G"],
    "Day 2": ["H", "I"],
    "Day 3": ["J", "K"],
    "Day 4": ["L", "M"],
    "Day 5": ["N", "O"],
    "Day 6": ["P", "Q"]
  };

  if (!(day in dayColumns)) {
    Logger.log("Invalid day provided.");
    return;
  }

  // Get column indexes
  var columns = dayColumns[day].map(letter => columnLetterToIndex(letter));
  var data = sheet.getDataRange().getValues(); // Get all sheet data

  var studentNames = [];
  for (var i = 1; i < data.length; i++) { // Skip header row
    var studentName = data[i][1]; // Column B (index 1)
    var instructorInFirstColumn = data[i][columns[0]]; // Instructor name from the first column (e.g., F for Day 1)
    var instructorInSecondColumn = data[i][columns[1]]; // Instructor name from the second column (e.g., G for Day 1)

    // Check if any of the instructor names match (case-insensitive and trimmed)
    var matchFound = columns.some(colIndex => {
      return instructors.some(instructor => {
        // Use regular expression to find instructor anywhere in the cell (case-insensitive)
        var regex = new RegExp(instructor, 'i');  // 'i' for case-insensitive match
        return regex.test(data[i][colIndex].toString().trim());  // Match anywhere in the cell
      });
    });

    if (matchFound && studentName) {
      // Push both student name and instructor name in the second column to the array
      studentNames.push(studentName + " | " + instructorInSecondColumn);
    }
  }

  // Insert results into the Google Doc
  body += "\n\nStudents assigned to " + instructors.join(", ") + " on " + day + ":\n" + studentNames.join("\n");
  doc.getBody().setText(body);

  Logger.log("Script completed successfully.");
}

// Helper function to convert column letter to index (0-based)
function columnLetterToIndex(letter) {
  return letter.charCodeAt(0) - "A".charCodeAt(0);
}
