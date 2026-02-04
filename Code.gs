/**
 * Creates the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Merge Tools')
    .addItem('Run Merger', 'openMergeDialog')
    .addToUi();
}

/**
 * Opens the modal dialog for user to select a worksheet.
 */
function openMergeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Worksheet to Merge');
}

/**
 * Reads the 'Setup' sheet to get available worksheet names.
 * Used by the client-side JavaScript to populate the dropdown.
 */
function getSetupWorksheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName('Setup');
  
  if (!setupSheet) {
    throw new Error('The "Setup" sheet was not found.');
  }

  // Assumes headers in row 1, data starts in row 2
  // Col A: Worksheet Name, Col B: Master Doc ID, Col C: Mapping, Col D: Folder ID
  const lastRow = setupSheet.getLastRow();
  if (lastRow < 2) return [];

  // Get all worksheet names from Column A (A2:A)
  const names = setupSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  
  // Filter out empty rows
  return names.filter(name => name !== '');
}

/**
 * Main logic to execute the merge for a specific worksheet.
 * @param {string} worksheetName - The name of the tab selected by the user.
 */
function runMergeProcess(worksheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName('Setup');
  
  // --- 1. GET CONFIGURATION FROM SETUP ---
  const setupData = setupSheet.getDataRange().getValues();
  // Find the row corresponding to the selected worksheet
  const configRow = setupData.find(row => row[0] === worksheetName);
  
  if (!configRow) {
    throw new Error(`Configuration for sheet "${worksheetName}" not found in Setup.`);
  }

  const masterDocId = configRow[1];
  const rawMapping = configRow[2];
  const folderId = configRow[3];

  if (!masterDocId || !rawMapping || !folderId) {
    throw new Error('Missing configuration data (Doc ID, Mapping, or Folder ID) in Setup.');
  }

  // --- IMPROVED JSON PARSING ---
  // Tries to parse strictly first, then cleans up smart quotes and single quotes if needed.
  let mapping;
  let jsonString = String(rawMapping).trim();

  // Replace smart/curly quotes with straight quotes just in case
  jsonString = jsonString.replace(/[\u201C\u201D]/g, '"').replace(/[\u2018\u2019]/g, "'");

  try {
    // Attempt 1: Parse as provided (handles cases like {"1": "Student's Name"} correctly)
    mapping = JSON.parse(jsonString);
  } catch (e1) {
    try {
      // Attempt 2: If failed, try converting single quotes to double quotes (handles {'1':'val'})
      // This is the fallback for users accustomed to writing loose JSON.
      const looseJson = jsonString.replace(/'/g, '"');
      mapping = JSON.parse(looseJson);
    } catch (e2) {
       // If both fail, throw a clear error with the string for debugging
       throw new Error(`Invalid JSON in Setup Mapping column.\nCheck for missing quotes or commas.\nValue: "${rawMapping}"`);
    }
  }

  // --- 2. PREPARE DESTINATION ---
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error('Invalid Destination Folder ID.');
  }

  // --- 3. PROCESS THE TARGET WORKSHEET ---
  const dataSheet = ss.getSheetByName(worksheetName);
  if (!dataSheet) {
    throw new Error(`Sheet "${worksheetName}" does not exist.`);
  }

  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  
  // If no data rows (only header), exit
  if (lastRow < 2) return { success: true, count: 0 };

  const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataRange = dataSheet.getRange(2, 1, lastRow - 1, lastCol);
  const dataValues = dataRange.getValues();

  // Find Indices of Status Columns
  const colMap = {
    docId: headers.findIndex(h => h.toString().toLowerCase().includes('merged doc id')),
    docUrl: headers.findIndex(h => h.toString().toLowerCase().includes('merged doc url')),
    link: headers.findIndex(h => h.toString().toLowerCase().includes('link to merged doc')),
    status: headers.findIndex(h => h.toString().toLowerCase().includes('document merge status')),
  };

  // Find CODIGO and AÑO column indices for filename generation
  const codigoIndex = headers.findIndex(h => h.toString().trim().toUpperCase() === 'CODIGO');
  const anoIndex = headers.findIndex(h => h.toString().trim().toUpperCase() === 'AÑO');

  // Validate columns exist
  if (Object.values(colMap).some(idx => idx === -1)) {
    throw new Error('One or more required status columns (Merged Doc ID, URL, Link, Status) are missing in the worksheet.');
  }

  let processedCount = 0;
  const userEmail = Session.getActiveUser().getEmail();

  // Loop through rows
  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const currentRowNum = i + 2; // +2 because data starts at row 2 and array is 0-indexed

    // Check if "Merged Doc ID" is empty. If it has a value, skip.
    if (row[colMap.docId] !== '') {
      continue;
    }

    try {
      // A. Create Copy of Master Doc
      const templateFile = DriveApp.getFileById(masterDocId);
      // Create a temporary name for the file (e.g. Merged Doc - Row X)
      // We will rename it properly when converting to PDF
      const tempDocFile = templateFile.makeCopy(`Temp_${worksheetName}_Row${currentRowNum}`);
      const tempDocId = tempDocFile.getId();
      const doc = DocumentApp.openById(tempDocId);
      const body = doc.getBody();

      // B. Replace Placeholders
      // Mapping format: {'1': '<<Placeholder>>'} where key is Column Number
      for (const [colNumStr, placeholder] of Object.entries(mapping)) {
        const colIndex = parseInt(colNumStr) - 1; // Convert "1" to index 0
        if (colIndex >= 0 && colIndex < row.length) {
          let replaceValue = row[colIndex];
          
          // Format dates specifically if the value is a Date object
          if (replaceValue instanceof Date) {
             replaceValue = Utilities.formatDate(replaceValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          
          // Force string conversion
          replaceValue = replaceValue === null || replaceValue === undefined ? '' : String(replaceValue);
          
          body.replaceText(placeholder, replaceValue);
        }
      }

      doc.saveAndClose();

      // C. Determine Filename: sheetName - <<CODIGO>> - <<AÑO>>
      let codigoVal = '';
      if (codigoIndex !== -1 && row[codigoIndex] !== undefined) {
          codigoVal = row[codigoIndex];
          if (codigoVal instanceof Date) {
             codigoVal = Utilities.formatDate(codigoVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
      }
      
      let anoVal = '';
      if (anoIndex !== -1 && row[anoIndex] !== undefined) {
          anoVal = row[anoIndex];
          if (anoVal instanceof Date) {
             anoVal = Utilities.formatDate(anoVal, Session.getScriptTimeZone(), "yyyy");
          }
      }
      
      // Force string and trim
      codigoVal = String(codigoVal);
      anoVal = String(anoVal);
      
      const fileName = `${worksheetName} - ${codigoVal} - ${anoVal}`;

      // D. Convert to PDF
      const pdfBlob = tempDocFile.getAs(MimeType.PDF);
      pdfBlob.setName(fileName);
      const pdfFile = folder.createFile(pdfBlob);
      
      // E. Set Permission (Anyone with link can view)
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // F. Clean up Temp Doc
      tempDocFile.setTrashed(true);

      // G. Generate Output Data
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd yyyy h:mm a");
      const statusMsg = `Document successfully created; Document successfully merged; Manually run by ${userEmail}; Timestamp: ${timestamp}`;
      const pdfUrl = pdfFile.getUrl();
      const pdfId = pdfFile.getId();
      
      // Use semicolon for formula separator as requested
      const hyperlinkFormula = `=HYPERLINK("${pdfUrl}"; "${fileName}")`;

      // H. Write to Sheet (Updating cell by cell to ensure safety if script times out later)
      dataSheet.getRange(currentRowNum, colMap.docId + 1).setValue(pdfId);
      dataSheet.getRange(currentRowNum, colMap.docUrl + 1).setValue(pdfUrl);
      dataSheet.getRange(currentRowNum, colMap.link + 1).setValue(hyperlinkFormula);
      dataSheet.getRange(currentRowNum, colMap.status + 1).setValue(statusMsg);

      processedCount++;

    } catch (err) {
      console.error(`Row ${currentRowNum} failed: ${err.message}`);
      dataSheet.getRange(currentRowNum, colMap.status + 1).setValue(`Error: ${err.message}`);
    }
  }

  return { success: true, count: processedCount };
}