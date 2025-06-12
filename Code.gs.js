// Shared helper functions
function getOrCreateFolder(path) {
  const root = DriveApp.getRootFolder();
  const folders = path.split('/');
  let currentFolder = root;
  
  for (const folderName of folders) {
    const existingFolders = currentFolder.getFoldersByName(folderName);
    currentFolder = existingFolders.hasNext() ? existingFolders.next() : 
                    currentFolder.createFolder(folderName);
  }
  
  return currentFolder;
}

function getFileExtension(filename) {
  const match = filename.match(/\.[0-9a-z]+$/i);
  return match ? match[0] : '';
}

function parseDurationToSeconds(duration) {
  if (!duration) return 0;
  
  // Handle Excel time values (numbers)
  if (typeof duration === 'number') {
    return duration * 86400; // Convert day fraction to seconds
  }
  
  // Handle string formats
  if (typeof duration === 'string') {
    duration = duration.trim();
    
    // Try "Xh Ym Zs" format (e.g., "00h 00m 24s")
    const hmsFormat = duration.match(/(\d+)h\s*(\d+)m\s*(\d+)s/i);
    if (hmsFormat) {
      const hours = parseInt(hmsFormat[1]) || 0;
      const minutes = parseInt(hmsFormat[2]) || 0;
      const seconds = parseInt(hmsFormat[3]) || 0;
      return (hours * 3600) + (minutes * 60) + seconds;
    }
    
    // Try "HH:MM:SS" format (e.g., "00:00:24")
    const timeParts = duration.split(':');
    if (timeParts.length === 3) {
      const hours = parseInt(timeParts[0]) || 0;
      const minutes = parseInt(timeParts[1]) || 0;
      const seconds = parseInt(timeParts[2]) || 0;
      return (hours * 3600) + (minutes * 60) + seconds;
    }
  }
  
  return 0; // Unknown format
}

// Main functions
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Student Engagement Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserNames() {
  const cache = CacheService.getScriptCache();
  const cachedNames = cache.get('user_names');
  
  if (cachedNames) {
    return JSON.parse(cachedNames);
  }
  
  const ss = SpreadsheetApp.openById('1zlZCmKxS1gyzFRRrEaMUHgclJaHpMqjpZiC-rTr_l2A');
  const sheet = ss.getSheetByName('JUNE CPL(Report)');
  const lastRow = sheet.getLastRow();
  
  const names = sheet.getRange(9, 1, lastRow - 3, 1).getValues()
    .flat()
    .filter(name => name !== '');
  
  // Cache for 1 hour
  cache.put('user_names', JSON.stringify(names), 3600);
  
  return names;
}

function processCPLForm(formData, files) {
  const startTime = new Date();
  const lock = LockService.getScriptLock();
  
  try {
    // Acquire lock to prevent concurrent execution
    if (!lock.tryLock(10000)) {
      throw new Error('Another submission is already in progress. Please wait.');
    }
    
    const ss = SpreadsheetApp.openById('1zlZCmKxS1gyzFRRrEaMUHgclJaHpMqjpZiC-rTr_l2A');
    const timeZone = Session.getScriptTimeZone();
    const today = new Date();
    const dateString = Utilities.formatDate(today, timeZone, 'd MMMM');
    
    // Get sheets in parallel
    const [cplSheet, actualCallingSheet] = [
      ss.getSheetByName('JUNE CPL(Report)'),
      ss.getSheetByName('Actual Callings')
    ];
    
    // Get headers and user names in parallel
    const [cplHeader, userNames] = [
      cplSheet.getRange(7, 1, 1, cplSheet.getLastColumn()).getValues()[0],
      cplSheet.getRange('A8:A').getValues().flat()
    ];
    
    // Find date column
    let dateCol = -1;
    for (let i = 0; i < cplHeader.length; i++) {
      const cell = cplHeader[i];
      if (cell instanceof Date && Utilities.formatDate(cell, timeZone, 'd MMMM') === dateString) {
        dateCol = i + 1;
        break;
      }
    }
    if (dateCol === -1) throw new Error("Today's date column not found in CPL sheet");
    
    // Find user row
    const userRow = userNames.indexOf(formData.userName) + 8;
    if (userRow < 8) throw new Error("User not found in CPL sheet");
    
    // Check for duplicate entry
    if (cplSheet.getRange(userRow, dateCol).getValue() !== '') {
      throw new Error('You have already submitted CPL data today');
    }
    
    // Prepare data
    const counts = [
      formData.thapar,
      formData.thaparArts,
      formData.lmThapar,
      formData.niit,
      formData.velTech,
      formData.alpha,
      formData.amity,
      formData.dsu,
      formData.kl,
      formData.sikkimManipal
    ];
    const total = counts.reduce((sum, count) => sum + count, 0);
    
    // Process files and update spreadsheet in parallel
    const excelAnalysis = analyzeExcelFile(files.excel);
    
    // Batch update spreadsheet
    cplSheet.getRange(userRow, dateCol, 1, 14).setValues([[
      formData.totalCalls,
      formData.studentEngaged,
      ...counts,
      0, // CPS
      total
    ]]);
    
    // Handle file uploads (don't wait for completion)
    handleFileUploads(formData.userName, files.screenshots, files.excel);
    
    // Update Actual Callings with validated count
    updateActualCallings(formData.userName, formData.studentEngaged, excelAnalysis);
    
    return { 
      success: true,
      excelAnalysis: excelAnalysis,
      processingTime: (new Date() - startTime) / 1000 + ' seconds'
    };
  } catch (error) {
    console.error('Error in processCPLForm:', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function analyzeExcelFile(excelFile) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(excelFile.data), excelFile.type);
    const tempFile = DriveApp.createFile(blob);
    const spreadsheet = SpreadsheetApp.open(tempFile);
    const sheet = spreadsheet.getSheets()[0];
    
    // Get all data in one read operation
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { countHHMMSS: 0, countHMS: 0, maxCount: 0, isValid: false };
    
    const range = sheet.getRange(`A1:H${lastRow}`);
    const data = range.getValues();
    
    let countHHMMSS = 0;
    let countHMS = 0;
    
    // Skip header row (index 0) and process from row 2 (index 1)
    for (let i = 1; i < data.length; i++) {
      const status = data[i][1]; // Column B (index 1)
      const duration = data[i][7]; // Column H (index 7)
      
      if (status === "Unknown" || status === "UNKNOWN") {
        const seconds = parseDurationToSeconds(duration);
        if (seconds > 24) {
          countHHMMSS++;
          countHMS++;
        }
      }
    }
    
    // Delete temporary file immediately
    tempFile.setTrashed(true);
    
    return {
      countHHMMSS: countHHMMSS,
      countHMS: countHMS,
      maxCount: Math.max(countHHMMSS, countHMS),
      isValid: countHHMMSS > 0 || countHMS > 0
    };
  } catch (error) {
    console.error('Error analyzing Excel file:', error);
    return {
      countHHMMSS: 0,
      countHMS: 0,
      maxCount: 0,
      isValid: false,
      error: error.message
    };
  }
}


function processCPSForm(formData, files) {
  const startTime = new Date();
  console.time('processCPSForm');
  
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      throw new Error('Another submission is already in progress. Please wait.');
    }
    
    const ss = SpreadsheetApp.openById('1zlZCmKxS1gyzFRRrEaMUHgclJaHpMqjpZiC-rTr_l2A');
    const sheet = ss.getSheetByName('JUNE CPL(Report)');
    const actualCallingSheet = ss.getSheetByName('Actual Callings');
    const timeZone = Session.getScriptTimeZone();
    const today = new Date();
    const dateString = Utilities.formatDate(today, timeZone, 'd MMMM');
    const cpsDateString = Utilities.formatDate(today, timeZone, 'yyyy-MM-dd');
    const timeString = Utilities.formatDate(today, timeZone, 'HH:mm:ss');

    // Find user row
    const userNames = sheet.getRange('A8:A').getValues().flat();
    const userRow = userNames.indexOf(formData.userName) + 8;
    if (userRow < 8) throw new Error("User not found");

    // Find date column
    const headerRow = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues()[0];
    let dateCol = -1;
    let alreadySubmitted = false;

    for (let i = 0; i < headerRow.length; i++) {
      const cell = headerRow[i];
      if (cell instanceof Date && Utilities.formatDate(cell, timeZone, 'd MMMM') === dateString) {
        dateCol = i + 1;
        if (sheet.getRange(userRow, dateCol).getValue() !== '') {
          alreadySubmitted = true;
        }
        break;
      }
    }

    if (alreadySubmitted) {
      throw new Error('You have already submitted data today');
    }
    if (dateCol === -1) throw new Error("Today's date column not found");

    // Write to CPL Report
    const dataToWrite = [
      formData.totalCalls,
      formData.totalCallsEngaged,
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, // university counts
      formData.students.length,     // CPS
      formData.students.length      // Total
    ];
    sheet.getRange(userRow, dateCol, 1, dataToWrite.length).setValues([dataToWrite]);

    // Only write to CPS Details Sheet if studentName has length > 0
    const cpsSpreadsheetId = '1mu-DRQm-VkYmJ-_Q6XbkYgTqS_2WA62TUP8umKRErh0';
    const cpsSheet = SpreadsheetApp.openById(cpsSpreadsheetId).getSheetByName('Sheet1');

    formData.students.forEach(student => {
      // Only process if studentName has content
      if (student.studentName && student.studentName.trim().length > 0) {
        cpsSheet.appendRow([
          cpsDateString,
          timeString,
          formData.userName,
          student.studentName.trim(),
          student.studentMobile || '',
          student.universityName || '',
          student.visitingDate || '',
          student.registration || 'Completed',
          student.application || 'Completed',
          student.admission || 'Completed'
        ]);
      }
    });

    // Handle file uploads
    handleFileUploads(formData.userName, files.screenshots, files.excel);

    // Analyze Excel file for call validation
    const excelAnalysis = analyzeExcelFile(files.excel);
    
    // Update Actual Callings sheet with validated count
    updateActualCallings(formData.userName, formData.totalCallsEngaged, excelAnalysis);

    console.timeEnd('processCPSForm');
    return { 
      success: true, 
      count: formData.students.filter(s => s.studentName && s.studentName.trim().length > 0).length,
      processingTime: (new Date() - startTime) / 1000 + ' seconds'
    };
  } catch (error) {
    console.error('Error in processCPSForm:', error);
    throw error;
  } finally {
    LockService.getScriptLock().releaseLock();
  }
}


function updateActualCallings(userName, formEngagedCount, excelAnalysis = null) {
  try {
    const ss = SpreadsheetApp.openById('1zlZCmKxS1gyzFRRrEaMUHgclJaHpMqjpZiC-rTr_l2A');
    const sheet = ss.getSheetByName('Actual Callings');
    const timeZone = Session.getScriptTimeZone();
    const today = new Date();
    const dateString = Utilities.formatDate(today, timeZone, 'd MMMM');

    // Get all needed data in one operation
    const [callingNames, callingHeader] = [
      sheet.getRange('B3:B').getValues().flat(),
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    ];
    
    const callingRow = callingNames.indexOf(userName) + 3;
    if (callingRow < 3) throw new Error("User not found in Actual Callings");

    let callingDateCol = -1;
    for (let i = 0; i < callingHeader.length; i++) {
      const cell = callingHeader[i];
      if (cell instanceof Date && Utilities.formatDate(cell, timeZone, 'd MMMM') === dateString) {
        callingDateCol = i + 1;
        break;
      }
    }
    if (callingDateCol === -1) throw new Error("Today's date column not found in Actual Callings");

    const finalCount = excelAnalysis?.isValid 
  ? Math.max(Number(formEngagedCount), excelAnalysis.maxCount)
  : Number(formEngagedCount);

// Write finalCount to today's column
sheet.getRange(callingRow, callingDateCol).setValue(finalCount);



// If excelAnalysis is valid, also write its count in next column
if (excelAnalysis?.isValid) {
  sheet.getRange(callingRow, callingDateCol + 1).setValue(excelAnalysis.maxCount);
}

    return true;
  } catch (error) {
    console.error('Error in updateActualCallings:', error);
    throw error;
  }
}

function handleFileUploads(userName, screenshotFiles, excelFile) {
  try {
    const timeZone = Session.getScriptTimeZone();
    const today = new Date();
    const dateString = Utilities.formatDate(today, timeZone, 'dMMM');
    
    // Process Screenshots
    if (screenshotFiles && screenshotFiles.length > 0) {
      const screenshotFolder = getOrCreateFolder('Report_Screenshot/' + userName);
      for (let i = 0; i < screenshotFiles.length; i++) {
        const file = screenshotFiles[i];
        const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.type, file.name);
        const fileName = `${dateString}(${i+1})${getFileExtension(file.name)}`;
        screenshotFolder.createFile(blob).setName(fileName);
      }
    }
    
    // Process Excel
    if (excelFile) {
      const excelFolder = getOrCreateFolder('Report_Excle_Sheet/' + userName);
      const excelBlob = Utilities.newBlob(Utilities.base64Decode(excelFile.data), excelFile.type, excelFile.name);
      const excelFileName = `${dateString}_excle_Report${getFileExtension(excelFile.name)}`;
      excelFolder.createFile(excelBlob).setName(excelFileName);
    }
    
    return true;
  } catch (error) {
    console.error('Error in handleFileUploads:', error);
    throw error;
  }
}