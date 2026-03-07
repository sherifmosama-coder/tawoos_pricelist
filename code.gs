// --- Code.gs ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Modern Price List Generator')
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
        arabicName: productsData[i][0] || '',      // Col A
        englishName: productsData[i][1] || '',     // Col B
        unit: productsData[i][2] || '',            // Col C
        barcode: productsData[i][9] || '',                 // Col J (Barcode)
        unitCapacity: parseFloat(productsData[i][4]) || 1, // Col E (Unit Capacity)
        taxRate: parseFloat(productsData[i][7]) || 0,      // Col H (Tax Category)
        packDescAr: productsData[i][12] || '',             // Col M (Pack Desc Ar)
        packDescEn: productsData[i][13] || '',     // Col N (Pack Desc En)
        netCapacity: productsData[i][15] || ''     // Col P (Net Piece Capacity)
      });
    }
  }

  // Fetch Clients
  const clientsSheet = ss.getSheetByName('Clients');
  const clientsData = clientsSheet.getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < clientsData.length; i++) {
    if (clientsData[i][0]) clients.push(clientsData[i][0]);
  }

  // Fetch Signatories (Creates sheet if it doesn't exist)
  let sigSheet = ss.getSheetByName('Signatories');
  if (!sigSheet) {
    sigSheet = ss.insertSheet('Signatories');
    sigSheet.appendRow(['Name', 'Title', 'Phone']);
  }
  const sigData = sigSheet.getDataRange().getValues();
  const signatories = [];
  for (let i = 1; i < sigData.length; i++) {
    if (sigData[i][0]) {
      signatories.push({ name: sigData[i][0], title: sigData[i][1], phone: sigData[i][2] });
    }
  }

  // Get Sequence Number (Based on Log sheet rows)
  let logSheet = ss.getSheetByName('Log');
  let nextSeq = logSheet ? Math.max(1, logSheet.getLastRow()) : 1; // Assumes header is row 1

  return { products, clients, signatories, nextSeq };
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
  // but here we write the full row excluding M/N indices to be safe.
  sheet.getRange(nextRow, 1, 1, 12).setValues([rowData.slice(0, 12)]); // Write A-L
  sheet.getRange(nextRow, 15, 1, 2).setValues([rowData.slice(14, 16)]); // Write O-P (skips M, N)
  
  return getInitialData();
}

function saveToLogAndSignatories(payload) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  const logSheet = ss.getSheetByName('Log');
  
  // 1. Generate PDF in Drive and get URL
  let fileUrl = "No File Generated";
  try {
    const folderId = "YOUR_FOLDER_ID_HERE"; // Replace with your target Folder ID
    const folder = DriveApp.getFolderById(folderId);
    
    // Create a PDF blob from the HTML content sent from the frontend
    const htmlBlob = Utilities.newBlob(payload.htmlContent, 'text/html', payload.refNum + ".html");
    const pdfBlob = htmlBlob.getAs('application/pdf').setName(payload.refNum + ".pdf");
    
    const file = folder.createFile(pdfBlob);
    fileUrl = file.getUrl();
  } catch (e) {
    Logger.log("Drive Save Error: " + e.toString());
  }

  // 2. Map Signatory if New
  if (payload.newSignatory) {
    const sigSheet = ss.getSheetByName('Signatories');
    sigSheet.appendRow([payload.newSignatory.name, payload.newSignatory.title, payload.newSignatory.phone]);
  }

  // 3. Append Logs with the new File URL
  payload.logData.forEach(row => {
    // We add the fileUrl to the last column of the log
    row[18] = fileUrl; 
    logSheet.appendRow(row);
  });

  return { success: true, ref: payload.refNum };
}
