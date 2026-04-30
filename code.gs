// --- Code.gs ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Price List Generator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- PASSCODE SECURITY ENGINE ---
function verifyPasscode(pin) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  
  // Multi-User Engine Initialization
  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['User Name', 'Passcode']);
    usersSheet.appendRow(['Admin', '3334']); // Default fallback
  }

  const props = PropertiesService.getScriptProperties();
  const now = new Date().getTime();
  
  // 1. Check if currently locked out
  let lockoutUntil = parseInt(props.getProperty('lockoutUntil') || '0');
  if (now < lockoutUntil) {
    let remaining = Math.ceil((lockoutUntil - now) / 1000);
    return { success: false, message: `System locked.`, locked: true, remaining: remaining };
  }
  
  // 2. Dynamic Success Case
  const usersData = usersSheet.getDataRange().getValues();
  let matchedUser = null;
  for (let i = 1; i < usersData.length; i++) {
    if (String(usersData[i][1]) === String(pin)) {
      matchedUser = String(usersData[i][0]);
      break;
    }
  }

  if (matchedUser) {
    props.setProperty('failedAttempts', '0');
    return { success: true, userName: matchedUser };
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
  let currentUser = auth.userName;

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
        isPrivateLabel: productsData[i][10] === true || String(productsData[i][10]).toLowerCase() === 'true', // Column K
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
          barcode: String(productsData[i][9] || ''),
          isPrivateLabel: productsData[i][10] === true || String(productsData[i][10]).toLowerCase() === 'true', // Column K
          unitCapacity: parseFloat(productsData[i][4]) || 1, 
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
    sigSheet.appendRow(['Name EN', 'Name AR', 'Title EN', 'Title AR', 'Phone', 'Email']);
  }
  const sigData = sigSheet.getDataRange().getValues();
  const signatories = [];
  for (let i = 1; i < sigData.length; i++) {
    let ownerRaw = String(sigData[i][6] || '').trim(); // Column G
    
    // Split the cell by commas and trim spaces (e.g. "Sherif, Admin" becomes ["Sherif", "Admin"])
    let ownersList = ownerRaw.split(',').map(name => name.trim());
    
    // Load if the cell is empty (legacy) OR the currentUser's name is found anywhere in the list
    if ((sigData[i][0] || sigData[i][1]) && (ownerRaw === '' || ownersList.includes(currentUser))) {
      signatories.push({ 
        nameEn: String(sigData[i][0] || ''), 
        nameAr: String(sigData[i][1] || ''), 
        titleEn: String(sigData[i][2] || ''),
        titleAr: String(sigData[i][3] || ''),
        phone: String(sigData[i][4] || ''),
        email: String(sigData[i][5] || '')
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
        // Force native formatting to prevent timezone shifts from altering the day
        docDate: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[1] || ''),
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
        notes: String(r[14] || ''), // Single Note Column
        urlEn: String(r[15] || ''), // Shifted from 18 to 15
        urlAr: String(r[16] || ''), // Shifted from 19 to 16
        timestamp: r[17] instanceof Date ? Utilities.formatDate(r[17], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : String(r[17] || ''), // Shifted from 20 to 17
        signatory: String(r[18] || ''), // Shifted from 21 to 18 (Column S)
        createdBy: String(r[19] || '')  // Column T (Creator Name)
      });
    }
  }

  // ==========================================
  // COMMERCIAL TERMS & CONTRACTS EXTRACTION
  // ==========================================
  let invoiceDiscounts = [];
  let clientContracts = [];
  let clientVolumes = {};

  try {
    // 1. MIGRATION FIX: Read directly from the Active Spreadsheet
    const commSs = SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. Extract Invoice Discounts (خصومات فواتير ثابتة)
    let invSheet = commSs.getSheetByName("خصومات فواتير ثابتة");
    if (invSheet) {
      let invData = invSheet.getDataRange().getValues();
      let invDisplayData = invSheet.getDataRange().getDisplayValues(); // Grabs exact formatted string (e.g., "5%")
      let productHeaders = invData[1]; // Row 2
      
      for (let i = 2; i < invData.length; i++) { 
        let row = invData[i];
        let displayRow = invDisplayData[i];
        if (!row[0]) continue; 
        
        let clientDiscount = {
          client: String(row[0]).trim(),
          start: row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[5],
          end: row[7] instanceof Date ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[7],
          products: {}
        };
        
        for (let j = 9; j < row.length; j++) {
          let discountVal = String(displayRow[j]).trim(); 
          if (discountVal && discountVal !== "0" && discountVal !== "0%" && discountVal !== "0.00%" && discountVal !== "") {
            let pName = String(productHeaders[j]).trim();
            // Automatically append % if the sheet didn't include it
            clientDiscount.products[pName] = discountVal.includes('%') ? discountVal : discountVal + '%';
          }
        }
        if (Object.keys(clientDiscount.products).length > 0) invoiceDiscounts.push(clientDiscount);
      }
    }

    // 3. Extract Client Contracts (عقود عملاء)
    let contractSheet = commSs.getSheetByName("عقود عملاء");
    if (contractSheet) {
      let cData = contractSheet.getDataRange().getValues();
      let cDisplayData = contractSheet.getDataRange().getDisplayValues(); 
      for (let i = 3; i < cData.length; i++) { 
        let row = cData[i];
        let displayRow = cDisplayData[i];
        if (!row[7]) continue; 
        
        clientContracts.push({
          fileLink: String(row[2] || '').trim(), // Column C
          start: row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[3],
          end: row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[4],
          client: String(row[7]).trim(),
          isPrivateLabel: row[8] === true || String(row[8]).toLowerCase() === 'true',
          payTerms: `${row[10] || 0} Days - ${row[11] || ''}`,
          fees: {
            monthlyFixed: parseFloat(row[14]) || 0,
            monthlyPct: String(displayRow[16] || "0"), // Q
            centralPct: String(displayRow[18] || "0"), // S
            q1Fixed: parseFloat(row[20]) || 0,
            q2Fixed: parseFloat(row[22]) || 0,
            q3Fixed: parseFloat(row[24]) || 0,
            q4Fixed: parseFloat(row[26]) || 0,
            quarterlyPct: String(displayRow[28] || "0"),
            annualFixed: parseFloat(row[30]) || 0,
            annualPct: String(displayRow[32] || "0")
          }
        });
      }
    }

    // 4. Extract JSON Sales Volumes from Google Drive (Using NetSales for precision)
    let productScopeMap = {};
    products.forEach(p => { productScopeMap[p.arabicName.trim()] = p.isPrivateLabel; });

    let files = DriveApp.getFilesByName("Tawoos_Cache_NetSales.json");
    if (files.hasNext()) {
      let file = files.next();
      let jsonData = JSON.parse(file.getBlob().getDataAsString());
      if (jsonData && jsonData.data) {
        jsonData.data.forEach(record => {
          let cName = String(record.client).trim();
          let pName = String(record.product).trim();
          let qty = parseFloat(record.qty) || 0; 
          
          let isPL = productScopeMap[pName] === true;
          let scopeKey = isPL ? "privateLabel" : "ownBrand";
          
          let year = "Unknown";
          if (record.dateMs) {
             let d = new Date(record.dateMs);
             if (!isNaN(d.getFullYear())) year = d.getFullYear().toString();
          } else if (record.date) {
             let d = new Date(record.date);
             if (!isNaN(d.getFullYear())) year = d.getFullYear().toString();
          }

          // Safety Net: Keep track of specific scope AND total mixed volume to prevent missing data
          if (!clientVolumes[cName]) clientVolumes[cName] = { privateLabel: { years: {} }, ownBrand: { years: {} }, total: { years: {} } };
          
          if (!clientVolumes[cName][scopeKey].years[year]) clientVolumes[cName][scopeKey].years[year] = 0;
          clientVolumes[cName][scopeKey].years[year] += qty;

          if (!clientVolumes[cName].total.years[year]) clientVolumes[cName].total.years[year] = 0;
          clientVolumes[cName].total.years[year] += qty;
        });
      }
    }

    // 5. Extract Costing Data from Cache
    var costingData = {}; // FIX: Changed 'let' to 'var' so it survives outside the try/catch block!
    let costFiles = DriveApp.getFilesByName("Tawoos_Cache_Costing.json");
    if (costFiles.hasNext()) {
        let costFile = costFiles.next();
        let costJson = JSON.parse(costFile.getBlob().getDataAsString());
        if (costJson && costJson.data) {
            costJson.data.forEach(record => {
                let pName = String(record.product).trim();
                // Pick worst-case scenario (highest) between Actual and Market cost
                let cost = Math.max(parseFloat(record.unitActual) || 0, parseFloat(record.unitMarket) || 0);
                let dateStr = record.date;
                
                if (!costingData[pName]) costingData[pName] = [];
                costingData[pName].push({ date: dateStr, cost: cost });
            });
            // Sort history by date descending (newest first)
            for (let p in costingData) {
                costingData[p].sort((a, b) => new Date(b.date) - new Date(a.date));
            }
        }
    }

    // 6. Extract Raw Cadence Data for Scenario Builder Item-Level Insights
    var cadenceData = [];
    // Prioritize the richer NetSales file for actual revenue/margin history
    let salesFiles = DriveApp.getFilesByName("Tawoos_Cache_NetSales.json");
    
    if (salesFiles.hasNext()) {
        let sFile = salesFiles.next();
        try {
            let sJson = JSON.parse(sFile.getBlob().getDataAsString());
            cadenceData = sJson.data ? sJson.data : (Array.isArray(sJson) ? sJson : []);
        } catch(e) {
            console.error("Failed to parse NetSales JSON: " + e.message);
        }
    }

  } catch (e) {
    console.error("Error fetching commercial terms: " + e.message);
  }

  // Extract the absolute latest (current) cost for every product for the Deal Simulator
  const currentCosts = {};
  if (typeof costingData !== 'undefined') {
    Object.keys(costingData).forEach(productName => {
      let history = costingData[productName];
      if (history && history.length > 0) {
        // The map is already sorted by date descending, so index [0] is the latest
        currentCosts[productName] = parseFloat(history[0].cost || history[0].totalCost || history[0].unitCost || 0);
      }
    });
  }

  return { 
    products, clients, signatories, nextSeq, docLogs, userName: currentUser, 
    commercialData: { invoiceDiscounts, clientContracts, clientVolumes },
    costingData: costingData,
    currentCosts: currentCosts, // The Live Cost Bridge for the Simulator
    cadenceData: cadenceData
  };
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

  // 2. Map URL to specific columns (Shifted: P = index 15, Q = index 16)
  const urlColIndex = (payload.lang === 'EN') ? 15 : 16;

  // Authenticate user to link the new signatory to them
  let auth = verifyPasscode(payload.authPin);
  let currentUser = auth.success ? auth.userName : 'Unknown User';

  // 3. Save New Signatory if provided
  if (payload.newSignatory) {
    let sigSheet = ss.getSheetByName('Signatories');
    sigSheet.appendRow([
      payload.newSignatory.nameEn, 
      payload.newSignatory.nameAr, 
      payload.newSignatory.titleEn, 
      payload.newSignatory.titleAr, 
      payload.newSignatory.phone, 
      payload.newSignatory.email,
      currentUser // Column G (Owner)
    ]);
  }

  // 4. Save to Log Sheet
  if (logSheet && payload.logData && payload.logData.length > 0) {
    payload.logData.forEach(row => {
      row[urlColIndex] = fileUrl; 
      row[19] = currentUser; // Add Created By to Column T (Index 19)
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
  const urlColIndex = (payload.lang === 'EN') ? 15 : 16;
  
  // 2. Find specific rows to overwrite
  const data = logSheet.getDataRange().getValues();
  let rowIndices = [];
  for(let i=0; i<data.length; i++){
    if(String(data[i][0]) === String(payload.refNum)) rowIndices.push(i+1);
  }
  
  let timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd HH:mm");
  
  // Authenticate user to get their real name for the audit log
  let auth = verifyPasscode(payload.authPin);
  let currentUser = auth.success ? auth.userName : 'Unknown User';

  // 3. Update Existing Rows & Add Native Notes
  for(let p=0; p < payload.logData.length; p++) {
    let newRow = payload.logData[p];
    if(fileUrl) newRow[urlColIndex] = fileUrl;

    if (p < rowIndices.length) {
      let rowIndex = rowIndices[p];
      let oldRow = data[rowIndex - 1];
      
      for(let col=0; col<newRow.length; col++) {
        if (col === 17 || col === 15 || col === 16) continue; // Skip Timestamp and URLs for notes (Shifted left by 3)
        
        if (String(oldRow[col]) !== String(newRow[col])) {
          let cell = logSheet.getRange(rowIndex, col + 1);
          let oldNote = cell.getNote() || "";
          let changeLog = `[${timestampStr}] ${currentUser}: Changed from "${oldRow[col]}" to "${newRow[col]}"`;
          cell.setNote(oldNote ? oldNote + "\n" + changeLog : changeLog);
          cell.setValue(newRow[col]);
        }
      }
      logSheet.getRange(rowIndex, 18).setValue(newRow[17]); // Force timestamp update (Shifted left by 3)
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

// ==========================================
// COMMERCIAL CONTRACT MANAGER
// ==========================================
function saveCommercialContract(payload) {
  try {
    const folderId = '11vM83H5B-Z6oFVoC76gyCSCAU78JjVnM';
    let fileUrl = payload.existingLink || "";

    // 1. Handle File Upload to Google Drive
    if (payload.fileBase64) {
      let folder = DriveApp.getFolderById(folderId);
      let blob = Utilities.newBlob(Utilities.base64Decode(payload.fileBase64), payload.mimeType, payload.fileName);
      let newFile = folder.createFile(blob);
      fileUrl = newFile.getUrl();
    }

    // 2. Locate Existing Row or Create New
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("عقود عملاء");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 3; i < data.length; i++) {
      let rowClient = String(data[i][7]).trim();
      let rowPL = data[i][8] === true || String(data[i][8]).toLowerCase() === 'true';
      if (rowClient === payload.client && rowPL === payload.isPrivateLabel) {
        rowIndex = i + 1; // +1 for 1-based getRange
        break;
      }
    }

    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow() + 1; // Append new
    }

    // 3. Update Specific Columns (1-based index)
    // C(3): Link, D(4): Start, E(5): End, H(8): Client, I(9): PL
    // K(11): PayDays, L(12): PayCond, O(15): M.Fixed, Q(17): M.Pct, S(19): C.Pct
    // U(21): Q1, W(23): Q2, Y(25): Q3, AA(27): Q4, AC(29): Q.Pct
    // AE(31): A.Fixed, AG(33): A.Pct
    const updates = [
      {col: 3, val: fileUrl}, {col: 4, val: payload.start}, {col: 5, val: payload.end},
      {col: 8, val: payload.client}, {col: 9, val: payload.isPrivateLabel},
      {col: 11, val: payload.payDays}, {col: 12, val: payload.payCond},
      {col: 15, val: payload.mFixed || ""}, {col: 17, val: payload.mPct || ""}, {col: 19, val: payload.cPct || ""},
      {col: 21, val: payload.q1 || ""}, {col: 23, val: payload.q2 || ""}, {col: 25, val: payload.q3 || ""}, 
      {col: 27, val: payload.q4 || ""}, {col: 29, val: payload.qPct || ""},
      {col: 31, val: payload.aFixed || ""}, {col: 33, val: payload.aPct || ""}
    ];

    updates.forEach(u => sheet.getRange(rowIndex, u.col).setValue(u.val));

    return { success: true, message: "Contract Saved Successfully!", fileUrl: fileUrl };
  } catch (e) {
    return { success: false, message: "Server Error: " + e.message };
  }
}

// ==========================================
// AI CONTRACT ANALYZER ENGINE (100% FREE TIER ROUTER)
// ==========================================
function analyzeContractWithAI(payload) {
  try {
    // 1. YOUR NEW FREE API KEYS (Get these from aistudio.google.com, console.groq.com, openrouter.ai)
    const API_KEYS = {
      GEMINI: "******",
      GROQ: "******",
      OPENROUTER: "******"
    };

    let documentString = "";
    payload.pages.forEach(p => { documentString += `\n\n--- PAGE ${p.page} ---\n${p.text}`; });

    // AGGRESSIVE COMPRESSOR FOR OPENROUTER: Free tier limits bandwidth drastically. 
    // We restrict fallback text to the first 8,000 characters (where 90% of financial terms live).
    let compressedDocumentString = documentString.length > 8000 
        ? documentString.substring(0, 8000) + "\n\n[TEXT TRUNCATED FOR BANDWIDTH LIMITS]" 
        : documentString;

    let langInstruction = payload.language === 'ar' ? 
        "Output all keys, categories, and values strictly in professional Arabic." : 
        "Output all keys, categories, and values in English.";

    // THE MASTERCLASS PROMPT: FIGURES ONLY
    let systemPrompt = `You are an elite AI trained specifically for FMCG legal contract analysis in Egypt.
    ${langInstruction}

    ### EXTRACTION TARGETS:
    1. **Identity (CRITICAL)**:
       - Client Name: The specific retailer/distributor name.
       - Contract Scope: You MUST classify this as exactly one of these: "Standard" (Supplier Brand), "Private Label" (Client Brand/Tashnee' le-hisab al-ghayr), or "Both".
       - Validity Dates: The duration of the agreement.

    2. **Financials (ONLY WITH FIGURES)**:
       - Invoice Discount (%), Rebates (%), Fixed Fees (EGP), Listing Fees (EGP).
       
    3. **Risks (ONLY WITH FIGURES)**:
       - Short Supply Penalty (%), Payment Terms (Days), quality penalties.

    ### RULES:
    - For Contract Scope, search for keywords like "العلامة الخاصة" or "Private Label" to distinguish from standard trading.
    - Provide the EXACT page number.
    - Respond ONLY with a raw JSON array. 
    Schema:
    [
      { "category": "Identity", "key": "Client Name", "value": "Hyperone", "pageFound": 1 },
      { "category": "Financial", "key": "Quarterly Rebate", "value": "13%", "pageFound": 2 },
      { "category": "Risk", "key": "Short Supply Penalty", "value": "50% of PO", "pageFound": 3 }
    ]`;

    const extractJson = (text) => {
       let clean = text.replace(/```json/gi, '').replace(/```/g, '').trim();
       let start = clean.indexOf('[');
       let end = clean.lastIndexOf(']');
       if(start !== -1 && end !== -1) clean = clean.substring(start, end + 1);
       return JSON.parse(clean);
    };

    // --- AGENT 1: GEMINI 2.5 FLASH ---
    // Note: If you receive a quota "limit: 0" error, your Google Cloud Project requires billing to be enabled.
    const runGemini = () => {
       if (!API_KEYS.GEMINI || API_KEYS.GEMINI.includes("PUT_")) throw new Error("Gemini Key Missing");
       let url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEYS.GEMINI}`;
       let reqPayload = {
          "contents": [{"parts": [{"text": systemPrompt}, {"text": "DOCUMENT TEXT:\n" + documentString}]}],
          "generationConfig": { "temperature": 0.0 } 
       };
       let res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(reqPayload), muteHttpExceptions: true });
       let json = JSON.parse(res.getContentText());
       if (json.error) throw new Error("Gemini Error: " + json.error.message);
       return extractJson(json.candidates[0].content.parts[0].text);
    };

    // --- UNIVERSAL AGENT: Handles Groq and OpenRouter ---
    const runOpenAICompatible = (baseUrl, apiKey, modelId) => {
       if (!apiKey || apiKey.includes("PUT_")) throw new Error("API Key Missing for " + modelId);
       let reqPayload = {
          model: modelId,
          messages: [
             { role: "system", content: systemPrompt },
             { role: "user", content: "DOCUMENT TEXT:\n" + compressedDocumentString }
          ],
          temperature: 0.0
       };
       let res = UrlFetchApp.fetch(baseUrl, {
          method: "post",
          contentType: "application/json",
          headers: { 
             "Authorization": "Bearer " + apiKey,
             "HTTP-Referer": "https://tawoos-erp.com", 
             "X-Title": "Tawoos ERP" 
          },
          payload: JSON.stringify(reqPayload),
          muteHttpExceptions: true
       });
       let json = JSON.parse(res.getContentText());
       if (json.error) throw new Error(modelId + " Error: " + json.error.message);
       
       let rawContent = json.choices[0].message.content;
       try { return extractJson(rawContent); } catch(e) { throw new Error("Could not parse JSON output from " + modelId); }
    };

    // 3. THE ROUTER
    let finalData = null;
    let agentUsed = "";

    if (payload.agent === 'auto') {
        try {
            finalData = runGemini();
            agentUsed = "Gemini 2.5 Flash";
        } catch (geminiError) {
            // SILENT FALLBACK TO NVIDIA NEMOTRON
            finalData = runOpenAICompatible("https://openrouter.ai/api/v1/chat/completions", API_KEYS.OPENROUTER, "nvidia/nemotron-3-super:free");
            agentUsed = "OpenRouter Nemotron 3 (Fallback)";
        }
    } else if (payload.agent === 'groq_llama3') {
        finalData = runOpenAICompatible("https://api.groq.com/openai/v1/chat/completions", API_KEYS.GROQ, "llama-3.3-70b-versatile");
        agentUsed = "Groq Llama 3.3 70B";
    } else if (payload.agent === 'openrouter_free') {
        // NVIDIA Nemotron 3 Super: Stable, highly intelligent, April 2026 endpoint
        finalData = runOpenAICompatible("https://openrouter.ai/api/v1/chat/completions", API_KEYS.OPENROUTER, "nvidia/nemotron-3-super:free");
        agentUsed = "OpenRouter (NVIDIA Nemotron 3)";
    } else {
        finalData = runGemini();
        agentUsed = "Gemini 2.5 Flash";
    }

    return { success: true, data: finalData, usedAgent: agentUsed };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// PHASE 3: UPLOAD TO DRIVE & SAVE TO ARCHIVE
// ==========================================
function saveContractToSheet(payload) {
  try {
    // 1. Upload PDF to Google Drive
    let folderName = "Contract Archives";
    let folders = DriveApp.getFoldersByName(folderName);
    let targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    let blob = Utilities.newBlob(Utilities.base64Decode(payload.fileData), 'application/pdf', payload.docName);
    let file = targetFolder.createFile(blob);
    let fileUrl = file.getUrl();

    // 2. Intelligent Data Mapping (Mining the JSON for core Identity Terms)
    let clientName = "غير محدد";
    let scope = "غير محدد";
    let validity = "غير محدد";
    let risksList = [];

    payload.approvedJson.forEach(item => {
        if (!item.approved) return; // Skip unapproved items in summaries
        
        let k = (item.key || "").toLowerCase();
        let c = (item.category || "").toLowerCase();
        
        // Match Identity Terms (English or Arabic)
        if (k.includes('client') || k.includes('عميل') || k.includes('name') || k.includes('اسم')) clientName = item.value;
        if (k.includes('scope') || k.includes('نطاق') || k.includes('label')) scope = item.value;
        if (k.includes('valid') || k.includes('date') || k.includes('تاريخ') || k.includes('فترة')) validity = item.value;
        
        // Compile Risks
        if (c.includes('risk') || c.includes('مخاطر')) risksList.push(`• ${item.key}: ${item.value}`);
    });

    let financialSummary = `الخصومات: ${payload.totalRebates}\nالرسوم الثابتة: ${payload.totalFees}`;
    let riskSummary = risksList.length > 0 ? risksList.join("\n") : "لا توجد مخاطر واضحة";

    // 3. Save to Google Sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = 'أرشيف تحليل العقود';
    let sheet = ss.getSheetByName(sheetName);
    
    // Create sheet with 11 exact columns if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "تاريخ التحليل", "اسم المستند", "رابط العقد (Drive)", "اسم العميل", "نطاق التعاقد", 
        "فترة التعاقد", "حالة المراجعة", "الملخص المالي", "المخاطر التشغيلية", 
        "بيانات الذكاء الاصطناعي الأصلية (Raw AI JSON)", "البيانات المعتمدة (Approved JSON)"
      ]);
      sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#f3f4f6");
    }

    let dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    
    let rowData = [
       dateStamp,                            // 1: Date
       payload.docName,                      // 2: Doc Name
       fileUrl,                              // 3: Live Drive Link
       clientName,                           // 4: Client
       scope,                                // 5: Scope
       validity,                             // 6: Validity Dates
       "معتمد (Approved)",                   // 7: Status
       financialSummary,                     // 8: Finances
       riskSummary,                          // 9: Risks
       JSON.stringify(payload.rawAiJson),    // 10: Raw
       JSON.stringify(payload.approvedJson)  // 11: Final
    ];

    sheet.appendRow(rowData);
    
    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, 9);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getContractArchive() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('أرشيف تحليل العقود');
    if (!sheet) return [];

    // CRITICAL FIX: getDisplayValues() converts everything (including Dates) to safe strings
    let values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return [];

    // Remove header and map to objects
    return values.slice(1).reverse().map(row => ({
      date: row[0] || "",
      docName: row[1] || "",
      link: row[2] || "",
      client: row[3] || "غير محدد",
      scope: row[4] || "غير محدد",
      financials: row[7] || "0.00%",
      // Safety fallback to empty array in case of older 8-column test rows
      approvedJson: row[10] || "[]" 
    }));
  } catch (e) {
    return [];
  }
}

// ==========================================
// PHASE 3: FETCH PDF FROM DRIVE FOR ARCHIVE VIEWER
// ==========================================
function getDriveFileBase64(url) {
  try {
    // Extract the File ID from the Drive URL
    let match = url.match(/\/d\/(.*?)\//);
    if (!match || match.length < 2) {
        throw new Error("Invalid Google Drive URL. Ensure the contract was saved properly.");
    }
    
    let id = match[1];
    let file = DriveApp.getFileById(id);
    let blob = file.getBlob();
    
    // Encode as Base64 to safely transmit to the frontend
    return { 
        success: true, 
        base64: Utilities.base64Encode(blob.getBytes()) 
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
