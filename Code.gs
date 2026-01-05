// ============================================
// TEACHER ALLOCATION SYSTEM - APPS SCRIPT
// ============================================
// Installation: Tools > Script Editor > Paste this code > Save
// ============================================

// ============================================
// 1. COURSE FILTERING FUNCTIONS
// ============================================

/**
 * Filters courses based on category and search query
 * Triggered when user types in search box or changes category
 */
function filterCourses() {
  // This function is replaced by updateCourseDropdown()
  // Kept for backward compatibility
  updateCourseDropdown();
}

/**
 * Sets the selected course when user clicks on filtered result
 */
function selectCourse(courseName) {
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  dashboard.getRange('B8').setValue(courseName);
  dashboard.getRange('B5').setValue(courseName); // Update search box
}

/**
 * Clears the selected course
 */
function clearSelectedCourse() {
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // 1. Reset Category Filter to "All"
  dashboard.getRange('B3').setValue('All');
  
  // 2. Clear other fields
  dashboard.getRange('B9').clearContent(); // Clear selected course
  dashboard.getRange('B5').clearContent(); // Clear search box
  dashboard.getRange('B7').clearContent(); // Clear dropdown
  dashboard.getRange('D7').clearContent(); // Clear count
  dashboard.getRange('B11').clearContent();// Clear date
  dashboard.getRange('B13').clearContent();// Clear time
  
  // 3. Clear the search results area
  dashboard.getRange('18:101').clear();

  // 4. Refresh the dropdown options
  updateCourseDropdown();
}

// ============================================
// 2. TEACHER SEARCH FUNCTIONS
// ============================================

/**
 * Main function - Searches for available teachers
 * Triggered by "Search Available Teachers" button
 */
function searchAvailableTeachers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName('Dashboard');
    
    // Get search criteria
    var selectedCourse = dashboard.getRange('B9').getValue();
    var selectedDate = dashboard.getRange('B11').getValue();
    var selectedTime = dashboard.getRange('B13').getValue();
    
    Logger.log('Starting search with: Course=' + selectedCourse + ', Date=' + selectedDate + ', Time=' + selectedTime);
    
    // Validation
    if (!selectedCourse || !selectedDate || !selectedTime) {
      SpreadsheetApp.getUi().alert('Please select Course, Date, and Time before searching!');
      return;
    }
    
    // Clear previous results
    dashboard.getRange('A18:E100').clearContent();
    
    // Show loading message
    dashboard.getRange('A18').setValue('Searching...');
    SpreadsheetApp.flush();
    
    // Find qualified teachers
    var qualifiedTeachers = getQualifiedTeachers(selectedCourse);
    Logger.log('Found ' + qualifiedTeachers.length + ' qualified teachers');
    
    if (qualifiedTeachers.length === 0) {
      dashboard.getRange('A18').setValue('No teachers found who can teach this course.');
      return;
    }
    
    // Check availability for each teacher
    var results = [];
    for (var i = 0; i < qualifiedTeachers.length; i++) {
      var teacher = qualifiedTeachers[i];
      var availability = checkTeacherAvailability(teacher, selectedDate, selectedTime);
      results.push(availability);
    }
    
    Logger.log('Checked availability for ' + results.length + ' teachers');
    
    // Sort results: Available first, then by workload
    results.sort(function(a, b) {
      if (a.isAvailable && !b.isAvailable) return -1;
      if (!a.isAvailable && b.isAvailable) return 1;
      return a.workloadPercent - b.workloadPercent;
    });
    
    // Display results
    displaySearchResults(results);
    Logger.log('Results displayed successfully');
    
    // Log search
    logSearch(selectedCourse, selectedDate, selectedTime, results.length);
    
  } catch (error) {
    Logger.log('ERROR in searchAvailableTeachers: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard')
      .getRange('A18').setValue('Error: ' + error.message);
    throw error;
  }
}

/**
 * Gets list of teachers qualified to teach the selected course
 */
function getQualifiedTeachers(courseName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skillsSheet = ss.getSheetByName('Teacher_Skills_Master');
  var skillsData = skillsSheet.getDataRange().getValues();
  
  var qualifiedTeachers = [];
  
  // If using vertical format (Teacher | Course | Skill Level)
  for (var i = 1; i < skillsData.length; i++) {
    var teacher = skillsData[i][0];
    var course = skillsData[i][1];
    
    if (course === courseName && qualifiedTeachers.indexOf(teacher) === -1) {
      qualifiedTeachers.push(teacher);
    }
  }
  
  // If no teachers found, check horizontal format
  if (qualifiedTeachers.length === 0) {
    qualifiedTeachers = getQualifiedTeachersHorizontal(courseName);
  }
  
  return qualifiedTeachers;
}

/**
 * Alternative: Gets qualified teachers from horizontal format
 * (Teacher Name in Column A, Courses in subsequent columns)
 */
function getQualifiedTeachersHorizontal(courseName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skillsSheet = ss.getSheetByName('Teacher_Skills_Master');
  var skillsData = skillsSheet.getDataRange().getValues();
  
  var qualifiedTeachers = [];
  var headers = skillsData[0]; // First row contains course names
  
  for (var i = 1; i < skillsData.length; i++) {
    var teacher = skillsData[i][0];
    
    for (var j = 1; j < skillsData[i].length; j++) {
      if (headers[j] === courseName && skillsData[i][j] === 'Yes') {
        qualifiedTeachers.push(teacher);
        break;
      }
    }
  }
  
  return qualifiedTeachers;
}

/**
 * Checks if a teacher is available at specific date/time
 */
function checkTeacherAvailability(teacherName, date, time) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Try to find individual teacher sheet
  var teacherSheet = ss.getSheetByName(teacherName + '_Availability');
  
  if (!teacherSheet) {
    // Try consolidated availability sheet
    return checkConsolidatedAvailability(teacherName, date, time);
  }
  
  var data = teacherSheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      teacher: teacherName,
      isAvailable: false,
      reason: 'No availability data found',
      availableSlots: 0,
      workloadPercent: 100
    };
  }
  
  var headers = data[0]; // First row contains time slots
  
  // Format the time consistently (remove leading zeros for comparison)
  var searchTime = time.toString().trim();
  
  // Find time column
  var timeColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    var headerTime = headers[i].toString().trim();
    if (headerTime === searchTime) {
      timeColIndex = i;
      break;
    }
  }
  
  if (timeColIndex === -1) {
    Logger.log('Time slot not found: ' + searchTime + ' in teacher: ' + teacherName);
    Logger.log('Available headers: ' + headers.join(', '));
    return {
      teacher: teacherName,
      isAvailable: false,
      reason: 'Time slot ' + searchTime + ' not found in schedule',
      availableSlots: 0,
      workloadPercent: 100
    };
  }
  
  // Format date for comparison
  var searchDateStr = formatDate(date);
  
  // Find date row and check availability
  var isAvailable = false;
  var availableSlots = 0;
  var totalSlots = 0;
  var foundDate = false;
  
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowDateStr = formatDate(rowDate);
    
    if (rowDateStr === searchDateStr) {
      foundDate = true;
      // Check if available at this time
      var cellValue = data[i][timeColIndex];
      isAvailable = (cellValue !== 'Not Available');
      
      // Count available slots for this date
      for (var j = 1; j < data[i].length; j++) {
        if (data[i][j] !== '') { // Only count actual time slots
          totalSlots++;
          if (data[i][j] !== 'Not Available') {
            availableSlots++;
          }
        }
      }
      break;
    }
  }
  
  if (!foundDate) {
    Logger.log('Date not found: ' + searchDateStr + ' for teacher: ' + teacherName);
    return {
      teacher: teacherName,
      isAvailable: false,
      reason: 'Date not found in schedule',
      availableSlots: 0,
      workloadPercent: 100
    };
  }
  
  var workloadPercent = totalSlots > 0 ? Math.round((1 - availableSlots/totalSlots) * 100) : 0;
  
  return {
    teacher: teacherName,
    isAvailable: isAvailable,
    reason: isAvailable ? 'Available' : 'Not available at this time',
    availableSlots: availableSlots,
    workloadPercent: workloadPercent
  };
}

/**
 * Checks availability from consolidated Teacher_Availability_Master sheet
 */
function checkConsolidatedAvailability(teacherName, date, time) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availSheet = ss.getSheetByName('Teacher_Availability_Master');
  
  if (!availSheet) {
    return {
      teacher: teacherName,
      isAvailable: false,
      reason: 'No availability data found',
      availableSlots: 0,
      workloadPercent: 100
    };
  }
  
  var data = availSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find time column
  var timeColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == time) {
      timeColIndex = i;
      break;
    }
  }
  
  // Find matching row
  var isAvailable = false;
  var availableSlots = 0;
  var totalSlots = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === teacherName && formatDate(data[i][1]) === formatDate(date)) {
      var cellValue = data[i][timeColIndex];
      isAvailable = (cellValue !== 'Not Available');
      
      // Count available slots
      for (var j = 2; j < data[i].length; j++) {
        totalSlots++;
        if (data[i][j] !== 'Not Available') {
          availableSlots++;
        }
      }
      break;
    }
  }
  
  var workloadPercent = totalSlots > 0 ? Math.round((1 - availableSlots/totalSlots) * 100) : 0;
  
  return {
    teacher: teacherName,
    isAvailable: isAvailable,
    reason: isAvailable ? 'Available' : 'Not available at this time',
    availableSlots: availableSlots,
    workloadPercent: workloadPercent
  };
}

/**
 * Displays search results in the dashboard
 */

function displaySearchResults(results) {
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Clear previous results (6 columns wide)
  dashboard.getRange('A18:F100').clearContent();
  
  // Get the selected course to look up skills
  var selectedCourse = dashboard.getRange('B9').getValue();
  
  // Set headers (Skill Level is now at the end)
  var headerRange = dashboard.getRange('A17:F17');
  headerRange.setValues([
    ['Teacher Name', 'Status', 'Available Slots Today', 'Workload', 'Action', 'Skill Level']
  ]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9ead3');
  
  if (results.length === 0) {
    dashboard.getRange('A18').setValue('No qualified teachers found for this course.');
    return;
  }
  
  // Prepare results data
  var displayData = [];
  for (var i = 0; i < results.length; i++) {
    var result = results[i];
    
    // Get the skill level
    var skillLevel = getTeacherSkillLevel(result.teacher, selectedCourse);
    
    var status = result.isAvailable ? '✅ Available' : '⚠️ ' + result.reason;
    var workload = result.workloadPercent + '% busy';
    var action = result.isAvailable ? 'Assign' : 'View Schedule';
    
    // Order: Name, Status, Slots, Workload, Action, Skill
    displayData.push([
      result.teacher,
      status,
      result.availableSlots + ' slots',
      workload,
      action,
      skillLevel // Skill is now last
    ]);
  }
  
  // Write results starting at row 18
  if (displayData.length > 0) {
    var resultsRange = dashboard.getRange(18, 1, displayData.length, 6);
    resultsRange.setValues(displayData);
    
    // Format results
    formatSearchResults(18, displayData.length);
  }
}

/**
 * Formats the search results with colors
 */
function formatSearchResults(startRow, numRows) {
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  if (numRows === 0) return;
  
  try {
    for (var i = 0; i < numRows; i++) {
      var row = startRow + i;
      
      // 1. Status Formatting (Column B / 2)
      var statusCell = dashboard.getRange(row, 2); 
      var status = statusCell.getValue();
      
      if (status && status.toString().indexOf('✅') > -1) {
        statusCell.setBackground('#d4edda'); // Green
        statusCell.setFontColor('#155724');
      } else if (status && status.toString().indexOf('⚠️') > -1) {
        statusCell.setBackground('#fff3cd'); // Yellow
        statusCell.setFontColor('#856404');
      }
      
      // 2. Skill Level Formatting (Column F / 6)
      var skillCell = dashboard.getRange(row, 6);
      var skill = skillCell.getValue();
      if (skill === 'Expert') skillCell.setFontColor('#0b5394'); // Blue
      if (skill === 'Beginner') skillCell.setFontColor('#e69138'); // Orange
    }
    
    // Add borders to results table (6 columns wide)
    var tableRange = dashboard.getRange(startRow, 1, numRows, 6);
    tableRange.setBorder(true, true, true, true, true, true);
    
  } catch (e) {
    Logger.log('Error formatting results: ' + e);
  }
}
// ============================================
// 3. ASSIGNMENT FUNCTIONS
// ============================================

/**
 * Assigns a teacher to the selected course/time
 */
function assignTeacher(teacherName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  var assignmentsLog = ss.getSheetByName('Assignments_Log');
  
  // Get assignment details
  var course = dashboard.getRange('B8').getValue();
  var date = dashboard.getRange('B10').getValue();
  var time = dashboard.getRange('B12').getValue();
  var assignedBy = Session.getActiveUser().getEmail();
  
  // Generate assignment ID
  var lastRow = assignmentsLog.getLastRow();
  var assignmentId = 'ASG' + String(lastRow).padStart(4, '0');
  
  // Log assignment
  assignmentsLog.appendRow([
    assignmentId,
    course,
    teacherName,
    date,
    time,
    assignedBy,
    new Date()
  ]);
  
  // Update teacher availability
  updateTeacherAvailability(teacherName, date, time, 'Not Available');
  
  // Show confirmation
  SpreadsheetApp.getUi().alert(
    'Success!',
    'Assigned ' + teacherName + ' to ' + course + '\nDate: ' + date + '\nTime: ' + time,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Refresh search results
  searchAvailableTeachers();
}

/**
 * Updates teacher availability after assignment
 */
function updateTeacherAvailability(teacherName, date, time, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teacherSheet = ss.getSheetByName(teacherName + '_Availability');
  
  if (!teacherSheet) {
    // Update consolidated sheet if exists
    updateConsolidatedAvailability(teacherName, date, time, status);
    return;
  }
  
  var data = teacherSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find time column
  var timeColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == time) {
      timeColIndex = i;
      break;
    }
  }
  
  // Find date row and update
  for (var i = 1; i < data.length; i++) {
    if (formatDate(data[i][0]) === formatDate(date)) {
      teacherSheet.getRange(i + 1, timeColIndex + 1).setValue(status);
      
      // Apply formatting
      if (status === 'Not Available') {
        teacherSheet.getRange(i + 1, timeColIndex + 1).setBackground('#f8d7da');
      }
      break;
    }
  }
}

/**
 * Updates consolidated availability sheet
 */
function updateConsolidatedAvailability(teacherName, date, time, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availSheet = ss.getSheetByName('Teacher_Availability_Master');
  
  if (!availSheet) return;
  
  var data = availSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find time column
  var timeColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == time) {
      timeColIndex = i;
      break;
    }
  }
  
  // Find and update row
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === teacherName && formatDate(data[i][1]) === formatDate(date)) {
      availSheet.getRange(i + 1, timeColIndex + 1).setValue(status);
      break;
    }
  }
}

// ============================================
// 4. UTILITY FUNCTIONS
// ============================================

/**
 * Formats date for consistent comparison
 */
function formatDate(date) {
  if (!date) return '';
  
  // If it's already a string with day name, extract just the date part
  if (typeof date === 'string') {
    // Handle format like "02/01/2026 (Friday)"
    var match = date.match(/(\d{2}\/\d{2}\/\d{4})/);
    if (match) {
      return match[1];
    }
    return date;
  }
  
  // If it's a Date object, format it
  try {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  } catch (e) {
    Logger.log('Error formatting date: ' + date + ', Error: ' + e);
    return date.toString();
  }
}

/**
 * Logs search activity
 */
function logSearch(course, date, time, resultsCount) {
  // Optional: Create a Search_Log sheet to track searches
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Search_Log');
  
  if (logSheet) {
    logSheet.appendRow([
      new Date(),
      Session.getActiveUser().getEmail(),
      course,
      date,
      time,
      resultsCount
    ]);
  }
}

// ============================================
// 5. DEBUG & TESTING FUNCTIONS
// ============================================

/**
 * Debug function - Run this to test the search manually
 */
function debugSearch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  
  var course = dashboard.getRange('B9').getValue();
  var date = dashboard.getRange('B11').getValue();
  var time = dashboard.getRange('B13').getValue();
  
  Logger.log('=== DEBUG SEARCH ===');
  Logger.log('Course: ' + course);
  Logger.log('Date: ' + date);
  Logger.log('Time: ' + time);
  
  // Get qualified teachers
  var teachers = getQualifiedTeachers(course);
  Logger.log('Qualified teachers: ' + teachers.join(', '));
  
  // Check each teacher
  for (var i = 0; i < teachers.length; i++) {
    Logger.log('--- Checking: ' + teachers[i] + ' ---');
    var result = checkTeacherAvailability(teachers[i], date, time);
    Logger.log('Available: ' + result.isAvailable);
    Logger.log('Reason: ' + result.reason);
    Logger.log('Available Slots: ' + result.availableSlots);
    Logger.log('Workload: ' + result.workloadPercent + '%');
  }
  
  Logger.log('=== END DEBUG ===');
}

/**
 * Test function - Check if teacher sheets exist
 */
function checkTeacherSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  Logger.log('=== CHECKING TEACHER SHEETS ===');
  
  var teacherSheets = [];
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf('_Availability') > -1) {
      teacherSheets.push(sheetName);
      Logger.log('Found: ' + sheetName);
    }
  }
  
  Logger.log('Total teacher availability sheets: ' + teacherSheets.length);
  Logger.log('=== END CHECK ===');
  
  return teacherSheets;
}

/**
 * Test function - Check teacher skills data
 */
function checkTeacherSkills() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skillsSheet = ss.getSheetByName('Teacher_Skills_Master');
  
  if (!skillsSheet) {
    Logger.log('ERROR: Teacher_Skills_Master sheet not found!');
    return;
  }
  
  var data = skillsSheet.getDataRange().getValues();
  
  Logger.log('=== TEACHER SKILLS DATA ===');
  Logger.log('Total rows: ' + data.length);
  Logger.log('Headers: ' + data[0].join(', '));
  Logger.log('Sample data (first 5 rows):');
  
  for (var i = 1; i < Math.min(6, data.length); i++) {
    Logger.log('Row ' + i + ': ' + data[i].join(' | '));
  }
  
  Logger.log('=== END CHECK ===');
}

// ============================================
// 5. MENU & TRIGGER SETUP
// ============================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Teacher Allocation')
    .addItem('Search Available Teachers', 'searchAvailableTeachers')
    .addItem('Refresh Courses', 'filterCourses')
    .addItem('Clear Selection', 'clearSelectedCourse')
    .addSeparator()
    .addItem('Setup Dashboard', 'setupDashboard')
    .addToUi();
}

/**
 * Sets up the dashboard with buttons and formatting
 */
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  
  // Add buttons (using drawings or images with assigned scripts)
  SpreadsheetApp.getUi().alert(
    'Setup Instructions',
    'To add buttons:\n\n' +
    '1. Insert > Drawing\n' +
    '2. Create a button shape with text\n' +
    '3. Click the three dots > Assign script\n' +
    '4. Enter function name: searchAvailableTeachers\n\n' +
    'Create buttons for:\n' +
    '- Search Available Teachers (B14)\n' +
    '- Filter Courses (C5)\n' +
    '- Clear Selection (C8)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Trigger for course filtering on cell edit
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  if (sheet.getName() !== 'Dashboard') return;
  
  var cellAddress = range.getA1Notation();
  
  // Auto-filter when search box or category changes
  if (cellAddress === 'B5' || cellAddress === 'B3') {
    updateCourseDropdown();
  }
  
  // When a course is selected from B7 dropdown, move it to B9 (Selected Course)
  if (cellAddress === 'B7') {
    var selectedCourse = range.getValue();
    if (selectedCourse) {
      sheet.getRange('B9').setValue(selectedCourse);
      range.clearContent(); // Clear the dropdown after selection
      sheet.getRange('B5').setValue(selectedCourse); // Update search box
    }
  }
}

/**
 * Updates the course dropdown in B7 based on category and search filters
 */
//Old version can be deleted
// function updateCourseDropdown() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var dashboard = ss.getSheetByName('Dashboard');
//   var coursesSheet = ss.getSheetByName('Courses_Master');
  
//   var category = dashboard.getRange('B3').getValue();
//   var searchQuery = dashboard.getRange('B5').getValue().toString().toLowerCase();
  
//   // Get all courses
//   var coursesData = coursesSheet.getDataRange().getValues();
//   var filteredCourses = [];
  
//   // Skip header row
//   for (var i = 1; i < coursesData.length; i++) {
//     var courseName = coursesData[i][0];
//     var courseCategory = coursesData[i][1];
//     var isActive = coursesData[i][2];
    
//     if (!courseName || isActive !== 'Yes') continue;
    
//     // Category filter
//     if (category && category !== 'All' && courseCategory !== category) {
//       continue;
//     }
    
//     // Search filter
//     if (searchQuery && !courseName.toLowerCase().includes(searchQuery)) {
//       continue;
//     }
    
//     filteredCourses.push(courseName);
//   }
  
//   // Update dropdown in B7 with filtered courses
//   var cell = dashboard.getRange('B7');
  
//   if (filteredCourses.length > 0) {
//     var rule = SpreadsheetApp.newDataValidation()
//       .requireValueInList(filteredCourses, true)
//       .setAllowInvalid(false)
//       .build();
//     cell.setDataValidation(rule);
//     cell.setValue(''); // Clear previous selection
    
//     // Show count
//     dashboard.getRange('D7').setValue(filteredCourses.length + ' courses found');
//   } else {
//     cell.clearDataValidations();
//     cell.setValue('');
//     dashboard.getRange('D7').setValue('No courses found');
//   }
// }
function updateCourseDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  var coursesSheet = ss.getSheetByName('Courses_Master');
  
  var category = dashboard.getRange('B3').getValue();
  var searchQuery = dashboard.getRange('B5').getValue().toString().toLowerCase();
  
  // Get all courses
  var coursesData = coursesSheet.getDataRange().getValues();
  var filteredCourses = [];
  
  // Skip header row
  for (var i = 1; i < coursesData.length; i++) {
    var courseName = coursesData[i][0];
    var courseCategory = coursesData[i][1];
    var isActive = coursesData[i][2];
    
    if (!courseName || isActive !== 'Yes') continue;
    
    // Category filter
    if (category && category !== 'All' && courseCategory !== category) {
      continue;
    }
    
    // Search filter
    if (searchQuery && !courseName.toLowerCase().includes(searchQuery)) {
      continue;
    }
    
    // --- NEW: PREVENT DUPLICATES ---
    // Only add if it's not already in the list
    if (filteredCourses.indexOf(courseName) === -1) {
      filteredCourses.push(courseName);
    }
  }
  
  // Update dropdown in B7 with filtered courses
  var cell = dashboard.getRange('B7');
  
  if (filteredCourses.length > 0) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(filteredCourses, true)
      .setAllowInvalid(false)
      .build();
    cell.setDataValidation(rule);
    // cell.setValue(''); // Optional: Only clear if current value is invalid
    
    // Show count
    dashboard.getRange('D7').setValue(filteredCourses.length + ' courses found');
  } else {
    cell.clearDataValidations();
    cell.setValue('');
    dashboard.getRange('D7').setValue('No courses found');
  }
}

/**
 * NEW: Helper to get skill level for a specific teacher and course
 */
function getTeacherSkillLevel(teacherName, courseName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Teacher_Skills_Master');
  var data = sheet.getDataRange().getValues();
  
  // Iterate through skills master to find the match
  for (var i = 1; i < data.length; i++) {
    // Col A is Teacher, Col B is Course, Col C is Skill Level
    if (data[i][0] === teacherName && data[i][1] === courseName) {
      return data[i][2]; // Return Skill Level (e.g., "Expert")
    }
  }
  return 'N/A';
}
