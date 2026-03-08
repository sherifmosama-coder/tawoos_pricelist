// --- Code.gs ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Price List Generator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  
  // Fetch Products
  const productsSheet = ss.getSheetByName('Products');
  const productsData = productsSheet.getDataRange().getValues();
  const products = [];
  
  for (let i = 1; i < productsData.length; i++) {
    if (productsData[i][0]) { 
      products.push({
        index: i,
        arabicName: String(productsData[i][0] || ''),      
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
  }

  // Fetch Clients
  const clientsSheet = ss.getSheetByName('Clients');
  const clientsData = clientsSheet.getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < clientsData.length; i++) {
    if (clientsData[i][0]) clients.push(String(clientsData[i][0]));
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

function addNewClient(clientName) {
  const ss = SpreadsheetApp.openById('1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw');
  const sheet = ss.getSheetByName('Clients');
  const data = sheet.getDataRange().getValues();
  
  // Check for duplicates
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientName) throw new Error('Client already exists');
  }
  
  sheet.appendRow([clientName]);
  return getInitialData(); // Refresh UI data
}

function addNewProduct(p) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('Products');
  
  // Find first empty row based on Column A
  const colA = sheet.getRange("A:A").getValues();
  let nextRow = 1;
  while (colA[nextRow - 1] && colA[nextRow - 1][0] !== "") {
    nextRow++;
  }

  // Prepare 16-column row (A to P)
  // Mapped Columns: A, B, C, E, F, H, J, P
  // Formula Columns (Set to empty): D, G, M, N
let rowData = new Array(16).fill(""); 
  rowData[0] = p.arabicName;     // A
  rowData[1] = p.englishName;    // B
  rowData[2] = p.unitType;       // C
  rowData[3] = "";               // D (EMPTY - Formula)
  rowData[4] = p.unitCapacity;   // E
  rowData[5] = p.smallUnitAr;    // F
  rowData[6] = "";               // G (EMPTY - Formula)
  rowData[7] = p.taxRate;        // H
  rowData[9] = p.barcode;        // J
  rowData[15] = p.pieceCap;      // P

  // Write range A-L (indices 0-11)
  sheet.getRange(nextRow, 1, 1, 12).setValues([rowData.slice(0, 12)]);
  // Write range O-P (indices 14-15) - skips M and N
  sheet.getRange(nextRow, 15, 1, 2).setValues([rowData.slice(14, 16)]);
  
  return getInitialData();
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
