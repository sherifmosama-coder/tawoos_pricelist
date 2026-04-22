// --- Code.gs ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Price List Generator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- PASSCODE SECURITY ENGINE ---
function verifyPasscode(pin) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date().getTime();
  
  // 1. Check if currently locked out
  let lockoutUntil = parseInt(props.getProperty('lockoutUntil') || '0');
  if (now < lockoutUntil) {
    let remaining = Math.ceil((lockoutUntil - now) / 1000);
    return { success: false, message: `System locked.`, locked: true, remaining: remaining };
  }
  
  // 2. Success Case
  if (pin === "3334") {
    props.setProperty('failedAttempts', '0'); // Reset counters
    return { success: true };
  } 
  
  // 3. Failure Case
  let attempts = parseInt(props.getProperty('failedAttempts') || '0');
  let firstFailure = parseInt(props.getProperty('firstFailureTime') || '0');
  
  // Reset the rolling window if it's been more than 1 hour since the first failure
  if (now - firstFailure > 60 * 60 * 1000) {
    attempts = 0;
    firstFailure = now;
    props.setProperty('firstFailureTime', firstFailure.toString());
  }
  
  attempts++;
  props.setProperty('failedAttempts', attempts.toString());
  
  // 4. Send Email Alert if 10 attempts reached
  if (attempts === 10) {
    try {
      MailApp.sendEmail({
        to: "sherif.m.osama@gmail.com",
        subject: "SECURITY ALERT: Multiple Failed Login Attempts",
        body: "SECURITY ALERT:\n\nThere have been 10 failed login attempts to your Tawoos Price list Engine within the last hour.\n\nPlease monitor your application."
      });
    } catch(e) {
      console.log("MailApp not authorized or failed.");
    }
  }
  
  // 5. Trigger 30-second Lockout every 3 attempts
  if (attempts % 3 === 0) {
    let cooldownSeconds = 30;
    props.setProperty('lockoutUntil', (now + cooldownSeconds * 1000).toString());
    return { success: false, message: `Too many attempts.`, locked: true, remaining: cooldownSeconds };
  }
  
  return { success: false, message: `Incorrect passcode. Attempt ${attempts % 3}/3`, locked: false };
}

function getInitialData(pin) {
  // --- GATEWAY CHECK ---
  let auth = verifyPasscode(pin);
  if (!auth.success) {
    return { authError: true, message: auth.message, locked: auth.locked, remaining: auth.remaining };
  }

  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  
  // Fetch Products
  const productsSheet = ss.getSheetByName('Products');
  const productsData = productsSheet.getDataRange().getValues();
  const products = [];
  
  for (let i = 1; i < productsData.length; i++) {
    
    // 1. Process Imported Products (Column A)
    let impName = String(productsData[i][0] || '');
    if (impName) { 
      products.push({
        index: 'imp_' + i, // Unique ID for imported
        arabicName: impName,      
        englishName: String(productsData[i][1] || ''),     
        unit: String(productsData[i][2] || ''),            
        barcode: String(productsData[i][9] || ''),         
        unitCapacity: parseFloat(productsData[i][4]) || 1, 
        taxRate: parseFloat(productsData[i][7]) || 0,      
        packDescAr: String(productsData[i][12] || ''),             
        packDescEn: String(productsData[i][13] || ''),     
        netCapacity: String(productsData[i][15] || '')     
      });
    }

    // 2. Process Manual Products (Column V)
    // First, ensure the row actually extends to column V to avoid errors
    if (productsData[i].length > 21) {
      let manName = String(productsData[i][21] || '');
      if (manName) {
        products.push({
          index: 'man_' + i, // Unique ID for manual
          arabicName: manName,      
          englishName: String(productsData[i][22] || ''),     
          unit: String(productsData[i][23] || ''),            
          barcode: String(productsData[i][30] || ''),         
          unitCapacity: parseFloat(productsData[i][25]) || 1, 
          taxRate: parseFloat(productsData[i][28]) || 0,      
          packDescAr: String(productsData[i][33] || ''),             
          packDescEn: String(productsData[i][34] || ''),     
          netCapacity: String(productsData[i][36] || '')     
        });
      }
    }
  }

  // Fetch Clients
  const clientsSheet = ss.getSheetByName('Clients');
  const clientsData = clientsSheet.getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < clientsData.length; i++) {
    let importedClient = String(clientsData[i][0] || ''); // Col A
    let manualClient = String(clientsData[i][2] || '');   // Col C
    
    if (importedClient) clients.push(importedClient);
    if (manualClient) clients.push(manualClient);
  }

  // Fetch Signatories 
  let sigSheet = ss.getSheetByName('Signatories');
  if (!sigSheet) {
    sigSheet = ss.insertSheet('Signatories');
    sigSheet.appendRow(['Name', 'Title', 'Phone']);
  }
  const sigData = sigSheet.getDataRange().getValues();
  const signatories = [];
  for (let i = 1; i < sigData.length; i++) {
    if (sigData[i][0]) {
      signatories.push({ 
        name: String(sigData[i][0]), 
        title: String(sigData[i][1]), 
        phone: String(sigData[i][2]) 
      });
    }
  }

  let logSheet = ss.getSheetByName('Log');
  let nextSeq = logSheet ? Math.max(1, logSheet.getLastRow()) : 1; 

  // Fetch Existing Archives for Dashboard (STRINGIFYING DATES)
  let rawLogs = logSheet ? logSheet.getDataRange().getValues() : [];
  let docLogs = [];
  if (rawLogs.length > 1) {
    for (let i = 1; i < rawLogs.length; i++) {
      let r = rawLogs[i];
      if (!r[0]) continue;
      
      docLogs.push({
        rowIdx: i + 1,
        refNum: String(r[0]),
        docDate: r[1] instanceof Date ? r[1].toISOString() : String(r[1] || ''),
        isPromo: String(r[2] || ''),
        promoQty: String(r[3] || ''),
        promoStart: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''),
        promoEnd: r[5] instanceof Date ? r[5].toISOString() : String(r[5] || ''),
        client: String(r[6] || ''),
        vatRate: String(r[7] || ''),
        productAr: String(r[8] || ''),
        offPrice: String(r[9] || ''),
        promoPrice: String(r[10] || ''),
        currency: String(r[11] || ''),
        rsp: String(r[12] || ''),
        sendTo: String(r[13] || ''),
        notes: [String(r[14] || ''), String(r[15] || ''), String(r[16] || ''), String(r[17] || '')],
        urlEn: String(r[18] || ''),
        urlAr: String(r[19] || ''),
        timestamp: r[20] instanceof Date ? r[20].toISOString() : String(r[20] || ''),
        signatory: String(r[21] || '') // Column V
      });
    }
  }

  return { products, clients, signatories, nextSeq, docLogs };
}

function addNewClient(clientName, pin) {
  const ss = SpreadsheetApp.openById('1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw');
  const sheet = ss.getSheetByName('Clients');
  
  // Find first empty row safely in Column C (Index 3)
  const colC = sheet.getRange("C:C").getValues();
  let nextRow = 2; // Always skip row 1 (headers) so read loop doesn't skip it
  while (nextRow <= colC.length && colC[nextRow-1][0] !== "") {
    nextRow++;
  }
  
  sheet.getRange(nextRow, 3).setValue(clientName);
  SpreadsheetApp.flush(); // Force immediate database commit
  
  return getInitialData(pin); 
}

function addNewProduct(product, pin) {
  const ss = SpreadsheetApp.openById('1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw');
  const sheet = ss.getSheetByName('Products');
  
  // Find first empty row safely in Column V (Index 22)
  const colV = sheet.getRange("V:V").getValues();
  let nextRow = 2; // Always skip row 1 (headers)
  while (nextRow <= colV.length && colV[nextRow-1][0] !== "") {
    nextRow++;
  }
  
  // Map exact data structure starting from Column V through AK
  let rowData = new Array(16).fill("");
  rowData[0] = product.arabicName;        // Col V
  rowData[1] = product.englishName;       // Col W
  rowData[2] = product.unitType;          // Col X
  rowData[4] = product.unitCapacity;      // Col Z
  rowData[7] = product.taxRate;           // Col AC
  rowData[9] = product.barcode;           // Col AE
  rowData[12] = product.smallUnitAr;      // Col AH
  rowData[13] = product.pieceCap;         // Col AI
  rowData[15] = product.pieceCap;         // Col AK (Net capacity)
  
  sheet.getRange(nextRow, 22, 1, 16).setValues([rowData]);
  SpreadsheetApp.flush(); // Force immediate database commit
  
  return getInitialData(pin);
}

function saveToLogAndSignatories(payload) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  const logSheet = ss.getSheetByName('Log');
  
  // 1. Create PDF in Google Drive
  let fileUrl = "";
  try {
    const folderId = "1MwSDzthUTaXuEYQlSrCWvcXKDn-2fABK"; // Dedicated Archive Folder
    const folder = DriveApp.getFolderById(folderId);
    
    // Decode the Base64 PDF string sent from html2pdf in the browser
    const decodedPdf = Utilities.base64Decode(payload.pdfData);
    const pdfBlob = Utilities.newBlob(decodedPdf, 'application/pdf', payload.refNum + ".pdf");
    const file = folder.createFile(pdfBlob);
    fileUrl = file.getUrl();
  } catch (e) {
    Logger.log("Drive Error: " + e.toString());
    fileUrl = "Error Saving to Drive";
  }

  // 2. Map URL to specific columns (S = index 18, T = index 19)
  const urlColIndex = (payload.lang === 'EN') ? 18 : 19;

  // 3. Save New Signatory if provided
  if (payload.newSignatory) {
    let sigSheet = ss.getSheetByName('Signatories');
    sigSheet.appendRow([payload.newSignatory.name, payload.newSignatory.title, payload.newSignatory.phone]);
  }

  // 4. Save to Log Sheet
  if (logSheet && payload.logData && payload.logData.length > 0) {
    payload.logData.forEach(row => {
      // row[18] is Col S, row[19] is Col T
      row[urlColIndex] = fileUrl; 
      logSheet.appendRow(row);
    });
  }
  return { success: true, fileUrl: fileUrl };
}

function forceDriveWriteScope() {
  DriveApp.createFile("Temp Auth File.txt", "You can delete this.", "text/plain");
}

function updatePriceList(payload) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  const logSheet = ss.getSheetByName('Log');
  
  // 1. Create PDF in Google Drive
  let fileUrl = "";
  try {
    const folderId = "1MwSDzthUTaXuEYQlSrCWvcXKDn-2fABK"; // Dedicated Archive Folder
    const folder = DriveApp.getFolderById(folderId);
    const decodedPdf = Utilities.base64Decode(payload.pdfData);
    const pdfBlob = Utilities.newBlob(decodedPdf, 'application/pdf', payload.refNum + ".pdf");
    const file = folder.createFile(pdfBlob);
    fileUrl = file.getUrl();
  } catch (e) {
    Logger.log("Drive Error: " + e.toString());
  }
  const urlColIndex = (payload.lang === 'EN') ? 18 : 19;
  
  // 2. Find specific rows to overwrite
  const data = logSheet.getDataRange().getValues();
  let rowIndices = [];
  for(let i=0; i<data.length; i++){
    if(String(data[i][0]) === String(payload.refNum)) rowIndices.push(i+1);
  }
  
  let timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd HH:mm");
  let userEmail = Session.getActiveUser().getEmail() || 'User';

  // 3. Update Existing Rows & Add Native Notes
  for(let p=0; p < payload.logData.length; p++) {
    let newRow = payload.logData[p];
    if(fileUrl) newRow[urlColIndex] = fileUrl;

    if (p < rowIndices.length) {
      let rowIndex = rowIndices[p];
      let oldRow = data[rowIndex - 1];
      
      for(let col=0; col<newRow.length; col++) {
        if (col === 20 || col === 18 || col === 19) continue; // Skip Timestamp and URLs for notes
        
        if (String(oldRow[col]) !== String(newRow[col])) {
          let cell = logSheet.getRange(rowIndex, col + 1);
          let oldNote = cell.getNote() || "";
          let changeLog = `[${timestampStr}] ${userEmail}: Changed from "${oldRow[col]}" to "${newRow[col]}"`;
          cell.setNote(oldNote ? oldNote + "\n" + changeLog : changeLog);
          cell.setValue(newRow[col]);
        }
      }
      logSheet.getRange(rowIndex, 21).setValue(newRow[20]); // Force timestamp update
      if(fileUrl) logSheet.getRange(rowIndex, urlColIndex + 1).setValue(fileUrl);
      
    } else {
      // Append if new products were added during edit
      logSheet.appendRow(newRow);
    }
  }
  
  // 4. Clear leftover rows if products were removed during edit
  for (let p = payload.logData.length; p < rowIndices.length; p++) {
    logSheet.getRange(rowIndices[p], 1, 1, logSheet.getLastColumn()).clearContent().clearNote();
  }

  return { success: true, fileUrl: fileUrl, ref: payload.refNum };
}

function getEditHistory(refNum) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  const logSheet = ss.getSheetByName('Log');
  const data = logSheet.getDataRange().getValues();
  let history = [];
  
  for(let i=0; i<data.length; i++) {
    if(String(data[i][0]) === String(refNum)) {
       let notes = logSheet.getRange(i+1, 1, 1, data[i].length).getNotes()[0];
       for(let n=0; n<notes.length; n++) {
         if(notes[n]) {
           history.push({
             product: data[i][8] || 'Global Setting',
             column: data[0][n], // Header name
             note: notes[n]
           });
         }
       }
    }
  }
  return history;
}

function authorizeEmailScope() { MailApp.sendEmail(Session.getActiveUser().getEmail(), "Auth", "Auth"); }
