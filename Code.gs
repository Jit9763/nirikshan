/**
 * Census Inspection Web Form Backend (Code.gs)
 * Saves reports to Google Sheets and serves the premium web interface.
 */

// Name of the Sheet where data will be stored
const SHEET_NAME = 'FieldReports';
const SYNC_SHEET_NAME = 'HLBSyncData';

/**
 * Serves the HTML Web Form.
 */
function doGet(e) {
  initSyncSheet();
  const htmlOutput = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('जनगणना निरीक्षण फील्ड रिपोर्ट - Census Inspection')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

/**
 * Handles CORS POST requests from external pages (like GitHub Pages).
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data && data.action === 'getReports') {
      const dashboardData = getAdminDashboardData();
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        reports: dashboardData.reports,
        syncData: dashboardData.syncData
      }))
      .setMimeType(ContentService.MimeType.JSON);
    }
    const result = saveReport(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'सर्वर एरर (डेटा पार्सिंग): ' + err.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Saves the submitted form data into Google Sheet.
 * @param {Object} data - Form fields submitted by frontend
 * @return {Object} Result of save operation
 */
function saveReport(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // If the sheet doesn't exist, create it and add headers
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        'Timestamp',
        'निरीक्षक का पद (Role)',
        'निरीक्षक का नाम (Inspector Name)',
        'निरीक्षण की तिथि (Inspection Date)',
        'राज्य/जिला (State/District)',
        'तहसील/ब्लॉक/वार्ड (Tehsil/Block/Ward)',
        'HLB नंबर (HLB No.)',
        'प्रगणक का नाम (Enumerator Name)',
        'पर्यवेक्षक का नाम (Supervisor Name)',
        'पर्यवेक्षक सर्कल नंबर (Supervisor Circle No.)',
        'जनगणना मकान अपेक्षित (Expected Houses)',
        'अब तक पूरे किए गए घर (Completed)',
        'बचे हुए घर (Pending)',
        'क्या नक़्शा सही है? (Map Correct)',
        'क्या कोई घर छूटा हुआ मिला? (Missed Houses)',
        'बिंदु 1: घर-घर जाकर डेटा भर रहा है (Door-to-Door)',
        'बिंदु 1: टिप्पणी (Door-to-Door Remarks)',
        'बिंदु 2: मकान सूचीकरण सही है (Houselisting)',
        'बिंदु 2: टिप्पणी (Houselisting Remarks)',
        'बिंदु 3: मुखिया का नाम सही दर्ज है (Head of Family)',
        'बिंदु 3: टिप्पणी (Head of Family Remarks)',
        'बिंदु 4: सदस्यों की संख्या/लिंग सही है (Members/Gender)',
        'बिंदु 4: टिप्पणी (Members/Gender Remarks)',
        'बिंदु 5: मोबाइल/आईडी स्वेच्छा से लिया (ID Taken)',
        'बिंदु 5: टिप्पणी (ID Taken Remarks)',
        'बिंदु 6: लोग सहयोग कर रहे हैं (Cooperation)',
        'बिंदु 6: टिप्पणी (Cooperation Remarks)',
        'बिंदु 7: डिजिटल ऐप/किट सही काम कर रही है (App Working)',
        'बिंदु 7: टिप्पणी (App Working Remarks)',
        'शिकायत 1 (Complaint 1)',
        'शिकायत 2 (Complaint 2)',
        'शिकायत 3 (Complaint 3)',
        'सुझाव 1 (Suggestion 1)',
        'सुझाव 2 (Suggestion 2)',
        'सुझाव 3 (Suggestion 3)',
        'समग्र फीडबैक (Overall Rating)',
        'हस्ताक्षर (Signature)'
      ];
      sheet.appendRow(headers);
      
      // Make header row bold and frozen
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
      sheet.setFrozenRows(1);
    } else {
      // Ensure the headers are upgraded
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
        const headerRange = sheet.getRange(1, 1, 1, lastColumn);
        let existingHeaders = headerRange.getValues()[0];
        
        // 1. Delete EB Number column if it exists
        const ebHeaderIndex = existingHeaders.indexOf('गणना ब्लॉक नंबर (EB Number)');
        if (ebHeaderIndex !== -1) {
          const colToDel = ebHeaderIndex + 1; // 1-indexed
          sheet.deleteColumn(colToDel);
          
          // Refetch headers after deletion
          const updatedLastCol = sheet.getLastColumn();
          existingHeaders = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
        }
        
        // 2. Ensure HLB number is at column 7 (index 6, which is column G)
        const hlbHeaderIndex = existingHeaders.indexOf('HLB नंबर (HLB No.)');
        if (hlbHeaderIndex === -1) {
          sheet.insertColumnBefore(7);
          sheet.getRange(1, 7).setValue('HLB नंबर (HLB No.)').setFontWeight('bold').setBackground('#f1f3f4');
        }
        
        // 3. Ensure Signature header exists
        const updatedLastCol = sheet.getLastColumn();
        const updatedHeaders = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
        const sigHeaderIndex = updatedHeaders.indexOf('हस्ताक्षर (Signature)');
        if (sigHeaderIndex === -1) {
          sheet.getRange(1, updatedLastCol + 1).setValue('हस्ताक्षर (Signature)').setFontWeight('bold').setBackground('#f1f3f4');
        }
      }
    }
    
    // Process Signature if present
    let signatureCellVal = '';
    if (data.signature_base64 && data.signature_base64.indexOf('data:image/png;base64,') === 0) {
      try {
        const base64Data = data.signature_base64.split(',')[1];
        const decoded = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decoded, 'image/png', 'Signature_' + (data.inspector_name || 'Inspector').replace(/[^a-zA-Z0-9]/g, '_') + '_' + new Date().getTime() + '.png');
        
        // Try to save inside the folder of the active Spreadsheet
        let folder = null;
        try {
          const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
          const file = DriveApp.getFileById(fileId);
          const folders = file.getParents();
          if (folders.hasNext()) {
            folder = folders.next();
          }
        } catch(e) {
          // Folder retrieval failed, will save to Drive root
        }
        
        const sigFile = folder ? folder.createFile(blob) : DriveApp.createFile(blob);
        sigFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileId = sigFile.getId();
        
        // Generate =IMAGE() formula to make it visible inline in Google Sheets
        signatureCellVal = `=IMAGE("https://drive.google.com/uc?export=view&id=${fileId}")`;
      } catch(e) {
        // Fallback to storing raw base64 if DriveApp is not authorized or fails
        signatureCellVal = data.signature_base64;
      }
    }
    
    // Process input data structure
    const rowData = [
      new Date(),
      data.inspector_role || '',
      data.inspector_name || '',
      data.inspection_date || '',
      data.district || '',
      data.block || '',
      data.hlb_no || '',
      data.enumerator_name || '',
      data.supervisor_name || '',
      data.supervisor_circle_no || '',
      Number(data.assigned_houses) || 0,
      Number(data.completed_houses) || 0,
      Number(data.pending_houses) || 0,
      data.map_correct || '',
      data.missed_houses || '',
      
      // Checklist Items (Status & Remarks)
      data.chk_1_status || '',
      data.chk_1_remarks || '',
      data.chk_2_status || '',
      data.chk_2_remarks || '',
      data.chk_3_status || '',
      data.chk_3_remarks || '',
      data.chk_4_status || '',
      data.chk_4_remarks || '',
      data.chk_5_status || '',
      data.chk_5_remarks || '',
      data.chk_6_status || '',
      data.chk_6_remarks || '',
      data.chk_7_status || '',
      data.chk_7_remarks || '',
      
      // Complaints
      data.complaint_1 || '',
      data.complaint_2 || '',
      data.complaint_3 || '',
      
      // Suggestions
      data.suggestion_1 || '',
      data.suggestion_2 || '',
      data.suggestion_3 || '',
      
      data.overall_rating || '',
      signatureCellVal
    ];
    
    // Write row
    sheet.appendRow(rowData);
    
    return {
      status: 'success',
      message: 'डेटा और हस्ताक्षर सफलतापूर्वक सेव हो गया है।',
      timestamp: new Date().toLocaleString()
    };
  } catch(e) {
    return {
      status: 'error',
      message: 'सेव करने में त्रुटि: ' + e.toString()
    };
  }
}

/**
 * Fetches all inspection reports from the Google Sheet.
 * @return {Array} List of reports as objects
 */
function getReportsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return [];
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Map headers to normalized keys
    const headerKeys = {
      'Timestamp': 'timestamp',
      'निरीक्षक का पद (Role)': 'inspector_role',
      'निरीक्षक का नाम (Inspector Name)': 'inspector_name',
      'निरीक्षण की तिथि (Inspection Date)': 'inspection_date',
      'राज्य/जिला (State/District)': 'district',
      'तहसील/ब्लॉक/वार्ड (Tehsil/Block/Ward)': 'block',
      'HLB नंबर (HLB No.)': 'hlb_no',
      'प्रगणक का नाम (Enumerator Name)': 'enumerator_name',
      'पर्यवेक्षक का नाम (Supervisor Name)': 'supervisor_name',
      'पर्यवेक्षक सर्कल नंबर (Supervisor Circle No.)': 'supervisor_circle_no',
      'जनगणना मकान अपेक्षित (Expected Houses)': 'assigned_houses',
      'अब तक पूरे किए गए घर (Completed)': 'completed_houses',
      'बचे हुए घर (Pending)': 'pending_houses',
      'क्या नक़्शा सही है? (Map Correct)': 'map_correct',
      'क्या कोई घर छूटा हुआ मिला? (Missed Houses)': 'missed_houses',
      'बिंदु 1: घर-घर जाकर डेटा भर रहा है (Door-to-Door)': 'chk_1_status',
      'बिंदु 1: टिप्पणी (Door-to-Door Remarks)': 'chk_1_remarks',
      'बिंदु 2: मकान सूचीकरण सही है (Houselisting)': 'chk_2_status',
      'बिंदु 2: टिप्पणी (Houselisting Remarks)': 'chk_2_remarks',
      'बिंदु 3: मुखिया का नाम सही दर्ज है (Head of Family)': 'chk_3_status',
      'बिंदु 3: टिप्पणी (Head of Family Remarks)': 'chk_3_remarks',
      'बिंदु 4: सदस्यों की संख्या/लिंग सही है (Members/Gender)': 'chk_4_status',
      'बिंदु 4: टिप्पणी (Members/Gender Remarks)': 'chk_4_remarks',
      'बिंदु 5: मोबाइल/आईडी स्वेच्छा से लिया (ID Taken)': 'chk_5_status',
      'बिंदु 5: टिप्पणी (ID Taken Remarks)': 'chk_5_remarks',
      'बिंदु 6: लोग सहयोग कर रहे हैं (Cooperation)': 'chk_6_status',
      'बिंदु 6: टिप्पणी (Cooperation Remarks)': 'chk_6_remarks',
      'बिंदु 7: डिजिटल ऐप/किट सही काम कर रही है (App Working)': 'chk_7_status',
      'बिंदु 7: टिप्पणी (App Working Remarks)': 'chk_7_remarks',
      'शिकायत 1 (Complaint 1)': 'complaint_1',
      'शिकायत 2 (Complaint 2)': 'complaint_2',
      'शिकायत 3 (Complaint 3)': 'complaint_3',
      'सुझाव 1 (Suggestion 1)': 'suggestion_1',
      'सुझाव 2 (Suggestion 2)': 'suggestion_2',
      'सुझाव 3 (Suggestion 3)': 'suggestion_3',
      'समग्र फीडबैक (Overall Rating)': 'overall_rating',
      'हस्ताक्षर (Signature)': 'signature'
    };
    
    const reports = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        const key = headerKeys[headers[j]] || headers[j];
        let val = row[j];
        if (val instanceof Date) {
          if (headers[j] === 'निरीक्षण की तिथि (Inspection Date)') {
            try {
              val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
            } catch(e) {
              val = val.toISOString().split('T')[0];
            }
          } else {
            try {
              val = val.toISOString();
            } catch(e) {
              val = val.toLocaleString();
            }
          }
        }
        obj[key] = val;
      }
      reports.push(obj);
    }
    return reports;
  } catch(e) {
    Logger.log("Error in getReportsData: " + e.toString());
    return [];
  }
}

/**
 * Combined function to fetch all reports and the latest HLB sync data mapping.
 */
function getAdminDashboardData() {
  try {
    return {
      reports: getReportsData(),
      syncData: getSyncData()
    };
  } catch(e) {
    Logger.log("Error in getAdminDashboardData: " + e.toString());
    return { reports: [], syncData: {} };
  }
}

/**
 * Reads synchronization data from the HLBSyncData sheet.
 */
function getSyncData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SYNC_SHEET_NAME);
    if (!sheet) {
      initSyncSheet();
      sheet = ss.getSheetByName(SYNC_SHEET_NAME);
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {};
    }
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const syncMap = {};
    for (let i = 0; i < data.length; i++) {
      const hlbNum = String(data[i][0]).trim();
      const syncVal = data[i][1];
      if (hlbNum) {
        syncMap[hlbNum] = syncVal;
      }
    }
    return syncMap;
  } catch(e) {
    Logger.log("Error in getSyncData: " + e.toString());
    return {};
  }
}

/**
 * Initializes the HLB Sync sheet with headers and pre-populates it with HLBs 1 to 242.
 */
function initSyncSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SYNC_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SYNC_SHEET_NAME);
      const headers = ['HLB Number', 'Sync Data'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9e1f2');
      sheet.setFrozenRows(1);
      
      // Populate HLB 1 to 242 rows
      const rows = [];
      for (let i = 1; i <= 242; i++) {
        rows.push([i, '']);
      }
      sheet.getRange(2, 1, 242, 2).setValues(rows);
    }
  } catch(e) {
    Logger.log("Error in initSyncSheet: " + e.toString());
  }
}
