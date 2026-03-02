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
        barcode: productsData[i][3] || '',                 // Col D
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

function saveToLogAndSignatories(payload) {
  const ssId = '1ojtygvWrUn1Zmb0CowQqbCOMtFL4z47u2ov3Elf8YQw';
  const ss = SpreadsheetApp.openById(ssId);
  
  // 1. Save New Signatory if provided
  if (payload.newSignatory) {
    let sigSheet = ss.getSheetByName('Signatories');
    sigSheet.appendRow([payload.newSignatory.name, payload.newSignatory.title, payload.newSignatory.phone]);
  }

  // 2. Save to Log Sheet
  let logSheet = ss.getSheetByName('Log');
  if (logSheet && payload.logData && payload.logData.length > 0) {
    payload.logData.forEach(row => {
      logSheet.appendRow(row);
    });
  }
  return true;
}
