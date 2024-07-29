function Menu()
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Finances')
    .addItem('Insert Sale', 'showHtmlForm')
    .addItem('Performance Report', 'businessPerformance')
    .addItem('Generate Invoice', 'exportSelectedRowToPDF')
    .addItem('Generate Receipt', 'exportSelectedRowToPDF_2')
    .addToUi();
}

function exportSelectedRowToPDF() {
  const companyInfo = {
    name: "L & L's Bakery Sales",
    address: "Lot 1A, Jalan Teknologi 5, Taman Teknologi Malaysia, 57000",
    website: "https://kidsavage2310.wixsite.com/luvenandlucksbakery"
  };
  const achRemittanceInfo = {
    bankName: "Maybank",
    accountName: "L & L's Bakery Sales",
    accountNumber: "123-456-789",
    routingNumber: "321-654-987",
    additionalInfo: "Please include the invoice number with your payment."
  };
  const zelleInfo = {
    notice: "Payments via Zelle are accepted.",
    emailAddress: "qking4352@google.com"
  };
  const checkRemittanceInfo = {
    payableTo: "L & L's Bakery Sales",
    address: "Lot 1A, Jalan Teknologi 5, Taman Teknologi Malaysia, 57000",
    additionalInfo: "Please include the invoice number on your check."
  };

  const imageUrl = 'https://drive.google.com/uc?id=1Phhi8cTeJoLZAGAVJhQy7bbimC3n8u1p';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedRow = sheet.getActiveRange().getRow();
  if (selectedRow <= 1) {
    SpreadsheetApp.getUi().alert('Please select a row other than the header row.');
    return;
  }
  let [orderID, orderDate, name, contactNumber, email, address, 
      productID, productName, quantity, unitPrice, totalPrice, paymentStatus] = 
    sheet.getRange(selectedRow, 1, 1, 12).getValues()[0];
  const dueDate = new Date(new Date(orderDate).getTime() + (30 * 24 * 60 * 60 * 1000));
  const doc = DocumentApp.create(`Invoice-${orderID}`);
  const body = doc.getBody();
  body.setMarginTop(72); // 1 inch
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // Insert Image
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  const image = body.appendImage(imageBlob);
  image.setWidth(100);
  image.setHeight(100);
  
  // Adjust position of the image
  body.appendParagraph("").setSpacingBefore(20);

  // Document Header
  body.appendParagraph(companyInfo.name)
      .setFontSize(16)
      .setBold(true)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(companyInfo.address)
      .setFontSize(10)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`${companyInfo.website}`)
      .setFontSize(10)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("");
  
  // Invoice Details
  body.appendParagraph(`Invoice #: ${orderID}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`Invoice Date: ${new Date(orderDate).toLocaleDateString()}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`Due Date: ${dueDate.toLocaleDateString()}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("");
  
  // Bill To Section
  body.appendParagraph("BILL TO:").setFontSize(10).setBold(true);
  body.appendParagraph(name).setFontSize(10);
  body.appendParagraph(address).setFontSize(10);
  body.appendParagraph(contactNumber).setFontSize(10);
  body.appendParagraph(email).setFontSize(10);
  body.appendParagraph("");
  
  // Services Table
  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  headerRow.appendTableCell('PRODUCT').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('UNIT PRICE (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('QUANTITY').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('TOTAL PRICE (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  
  const productRow = table.appendTableRow();
  productRow.appendTableCell(productName).setFontSize(10);
  productRow.appendTableCell(`RM ${parseFloat(unitPrice).toFixed(2)}`).setFontSize(10);
  productRow.appendTableCell(`${parseInt(quantity, 10)}`).setFontSize(10);
  productRow.appendTableCell(`RM ${parseFloat(totalPrice).toFixed(2)}`).setFontSize(10);
  
  // Financial Summary
  body.appendParagraph(`Total Price: RM ${parseFloat(totalPrice).toFixed(2)}`).setFontSize(10).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`Payment Status: ${paymentStatus}`).setFontSize(10).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("");
  
  // ACH Remittance Info
  body.appendParagraph("ACH Remittance to:").setFontSize(10).setBold(true);
  body.appendParagraph(`Bank Name: ${achRemittanceInfo.bankName}`).setFontSize(10);
  body.appendParagraph(`Account Name: ${achRemittanceInfo.accountName}`).setFontSize(10);
  body.appendParagraph(`Account Number: ${achRemittanceInfo.accountNumber}`).setFontSize(10);
  body.appendParagraph(`Routing Number: ${achRemittanceInfo.routingNumber}`).setFontSize(10);
  body.appendParagraph(achRemittanceInfo.additionalInfo).setFontSize(10);
  body.appendParagraph("");
  
  // Zelle Payment Information
  body.appendParagraph(zelleInfo.notice).setFontSize(10).setBold(true);
  body.appendParagraph(`Email: ${zelleInfo.emailAddress}`).setFontSize(10);
  body.appendParagraph("");
  
  // Physical Check Remittance Information
  body.appendParagraph("To remit by physical check, please send to:").setBold(true).setFontSize(10);
  body.appendParagraph(checkRemittanceInfo.payableTo).setFontSize(10);
  body.appendParagraph(checkRemittanceInfo.address).setFontSize(10);
  body.appendParagraph(checkRemittanceInfo.additionalInfo).setFontSize(10);

  // PDF Generation and Sharing
  doc.saveAndClose();
  const pdfBlob = doc.getAs('application/pdf');
  const folders = DriveApp.getFoldersByName("Invoices");
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Invoices");
  let version = 1;
  let pdfFileName = `Invoice-${orderID}_V${String(version).padStart(2, '0')}.pdf`;
  while (folder.getFilesByName(pdfFileName).hasNext()) {
    version++;
    pdfFileName = `Invoice-${orderID}_V${String(version).padStart(2, '0')}.pdf`;
  }
  const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = pdfFile.getUrl();
  const htmlOutput = HtmlService.createHtmlOutput(`<html><body><p>Invoice PDF generated successfully. Version: ${version}. <a href="${pdfUrl}" target="_blank" rel="noopener noreferrer">Click here to view and download your Invoice PDF</a>.</p></body></html>`)
                                .setWidth(300)
                                .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Invoice PDF Download');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}

function exportSelectedRowToPDF_2() {
  const companyInfo = {
    name: "L & L's Bakery Sales",
    address: "Lot 1A, Jalan Teknologi 5, Taman Teknologi Malaysia, 57000",
    website: "https://kidsavage2310.wixsite.com/luvenandlucksbakery"
  };
  
  const imageUrl = 'https://drive.google.com/uc?id=1Phhi8cTeJoLZAGAVJhQy7bbimC3n8u1p';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedRow = sheet.getActiveRange().getRow();
  
  const achRemittanceInfo = {
    bankName: "Maybank",
    accountNumber: "xxx-xxx-xxx",
    routingNumber: "xxx-xxx-xxx",
  };
  
  if (selectedRow <= 1) {
    SpreadsheetApp.getUi().alert('Please select a row other than the header row.');
    return;
  }
  
  let [orderID, orderDate, name, contactNumber, email, address, 
      productID, productName, quantity, unitPrice, totalPrice, paymentStatus] = 
    sheet.getRange(selectedRow, 1, 1, 12).getValues()[0];

  if (paymentStatus.toLowerCase() === 'pending') {
    SpreadsheetApp.getUi().alert('Cannot generate a receipt for a pending payment.');
    return;
  }
  
  const today = new Date();
  
  const doc = DocumentApp.create(`Receipt-${orderID}`);
  const body = doc.getBody();
  body.setMarginTop(72); // 1 inch
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // Insert Image
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  const image = body.appendImage(imageBlob);
  image.setWidth(100);
  image.setHeight(100);

  // Adjust position of the image
  body.appendParagraph("").setSpacingBefore(20);

  // Document Header
  body.appendParagraph(companyInfo.name)
      .setFontSize(16)
      .setBold(true)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(companyInfo.address)
      .setFontSize(10)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`${companyInfo.website}`)
      .setFontSize(10)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("");
  
  // Receipt Details
  body.appendParagraph(`Receipt #: ${orderID}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`Receipt Date: ${today.toLocaleDateString()}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("");
  
  // Bill To Section
  body.appendParagraph("BILL TO:").setFontSize(10).setBold(true);
  body.appendParagraph(name).setFontSize(10);
  body.appendParagraph(address).setFontSize(10);
  body.appendParagraph(contactNumber).setFontSize(10);
  body.appendParagraph(email).setFontSize(10);
  body.appendParagraph("");
  
  // Services Table
  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  headerRow.appendTableCell('DESCRIPTION').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('UNIT PRICE (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('QUANTITY').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  headerRow.appendTableCell('AMOUNT').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
  
  const productRow = table.appendTableRow();
  productRow.appendTableCell(productName).setFontSize(10);
  productRow.appendTableCell(`RM ${parseFloat(unitPrice).toFixed(2)}`).setFontSize(10);
  productRow.appendTableCell(`${parseInt(quantity, 10)}`).setFontSize(10);
  productRow.appendTableCell(`RM ${parseFloat(totalPrice).toFixed(2)}`).setFontSize(10);
  
  // Financial Summary
  body.appendParagraph(`Receipt Total: RM ${parseFloat(totalPrice).toFixed(2)}`).setFontSize(10).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // ACH Remittance Info
  body.appendParagraph("Customer Payment Details:").setFontSize(10).setBold(true);
  body.appendParagraph(`Bank Name: ${achRemittanceInfo.bankName}`).setFontSize(10);
  body.appendParagraph(`Account Number: ${achRemittanceInfo.accountNumber}`).setFontSize(10);
  body.appendParagraph(`Routing Number: ${achRemittanceInfo.routingNumber}`).setFontSize(10);
  body.appendParagraph("");
  
  // PDF Generation and Sharing
  doc.saveAndClose();
  const pdfBlob = doc.getAs('application/pdf');
  const folders = DriveApp.getFoldersByName("Receipts");
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Receipts");
  let version = 1;
  let pdfFileName = `Receipt-${orderID}_V${String(version).padStart(2, '0')}.pdf`;
  while (folder.getFilesByName(pdfFileName).hasNext()) {
    version++;
    pdfFileName = `Receipt-${orderID}_V${String(version).padStart(2, '0')}.pdf`;
  }
  const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = pdfFile.getUrl();
  const htmlOutput = HtmlService.createHtmlOutput(`<html><body><p>Receipt PDF generated successfully. Version: ${version}. <a href="${pdfUrl}" target="_blank" rel="noopener noreferrer">Click here to view and download your Receipt PDF</a>.</p></body></html>`)
                                .setWidth(300)
                                .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Receipt PDF Download');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}


function showHtmlForm() 
{
  var htmlOutput = HtmlService.createHtmlOutputFromFile('main');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'New Sale Form');
}

function storeData(order) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var orders = sheet.getSheetByName("Sales Sheet");

  var inventory = sheet.getSheetByName("Inventory Sheet");
  var inventoryData = inventory.getDataRange().getValues();
  var referenceList = inventoryData.map(row=>row[1]);

  var index = referenceList.indexOf(order.product);

  // Data Retrieval & Calculation
  let product_id = inventoryData[index][0];
  let price = inventoryData[index][5];
  let total = price*order.quantity;

  // OrderId Generator
  var lastRow = orders.getLastRow();
  var lastOrderId = orders.getRange(lastRow, 1).getValue();

  var newOrderId;
  var orderIdNumber = parseInt(lastOrderId.slice(1));
  newOrderId = 'S' + ('000' + (orderIdNumber + 1)).slice(-3);

  // Append Data
  orders.appendRow([newOrderId, order.orderDate, order.name, order.contact, order.email, order.address, product_id, order.product, order.quantity, price, total, "Pending"]);
  var newRow = orders.getRange(orders.getLastRow(), 1, 1, 12); 
  newRow.setBorder(true, true, true, true, true, true, '#F6E5AE', SpreadsheetApp.BorderStyle.SOLID);
}

function businessPerformance() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName('Sales Sheet');
  var inventorySheet = ss.getSheetByName('Inventory Sheet');
  
  var salesDataRange = salesSheet.getDataRange();
  var salesData = salesDataRange.getValues();
  
  var inventoryDataRange = inventorySheet.getDataRange();
  var inventoryData = inventoryDataRange.getValues();
  
  // Variable Intializing
  let totalSales = 0;
  let totalCost = 0;
  let totalProfit = 0;
  let totalQuantitySold = 0;
  let orderCount = 0;
  let paidOrderCount = 0;
  let itemSalesData = [];
  let totalInventoryQuantity = 0;
  let totalInventoryInStockQuantity = 0;
  
  // Inventory Costs Map
  let costMap = {};
  for (let i = 3; i < inventoryData.length; i++) 
  {
    let productId = inventoryData[i][0];

    let costPriceStr = inventoryData[i][4].toString(); 
    let costPrice = parseFloat(costPriceStr.replace('RM ', '').replace(',', ''));

    let totalQuantity = inventoryData[i][6];
    let inStockQuantity = inventoryData[i][8];
    
    costMap[productId] = 
    {
      costPrice: costPrice,
      totalQuantity: totalQuantity,
      inStockQuantity: inStockQuantity
    };
    
    totalInventoryQuantity += totalQuantity;
    totalInventoryInStockQuantity += inStockQuantity;
    totalCost += costPrice * inStockQuantity;
  }
  
  // Calculations for Sales
  for (let i = 3; i < salesData.length; i++) 
  {
    let quantity = salesData[i][8];

    let priceStr = salesData[i][9].toString(); 
    let price = parseFloat(priceStr.replace('RM ', '').replace(',', ''));

    let sales = quantity * price;
    totalSales += sales;
    itemSalesData.push([salesData[i][7], sales]);

    totalQuantitySold += quantity;
    totalProfit = totalSales - totalCost;
    
    if (salesData[i][11] === 'Paid') 
    {
      paidOrderCount++;
    }
    orderCount++;
  }
  
  let salesPerItem = totalSales / (itemSalesData.length || 1);
  let bestSeller = itemSalesData.reduce((max, item) => item[1] > max[1] ? item : max, ['', 0])[0];
  
  let stockPercentage = ((totalInventoryQuantity - totalQuantitySold) / totalInventoryQuantity) * 100;
  let paidOrderPercentage = (paidOrderCount / orderCount) * 100;
  
  // Creating Performance Report
  let perfSheet = ss.getSheetByName('Performance Sheet');
  if (!perfSheet) {
    perfSheet = ss.insertSheet('Performance Sheet');
  } else {
    perfSheet.clear();
  }
  
  // Generating Performance Report
  const imageUrl = 'https://drive.google.com/uc?id=16fOEigakwe4kfA6PJnzaCDhX_8l1jWJC'; 
  perfSheet.insertImage(imageUrl, 1,1); 
  
  perfSheet.appendRow(['', 'Performance Report']);
  perfSheet.appendRow(['', '‎']);
  perfSheet.appendRow(['Total Sales', `RM ${totalSales.toFixed(2)}`]);
  perfSheet.appendRow(['Total Cost', `RM ${totalCost.toFixed(2)}`]);
  perfSheet.appendRow(['Gross Income', `RM ${totalProfit.toFixed(2)}`]);
  perfSheet.appendRow(['Average Sale (per item)   ', `RM ${salesPerItem.toFixed(2)}`]);
  perfSheet.appendRow(['Best Selling Item', bestSeller]);
  perfSheet.appendRow(['Quantity Sold', totalQuantitySold.toFixed(2)]);
  perfSheet.appendRow(['Orders Created', orderCount.toFixed(2)]);
  perfSheet.appendRow(['Stock Percentage', `${stockPercentage.toFixed(2)}%`]);
  perfSheet.appendRow(['Paid Percentage', `${paidOrderPercentage.toFixed(2)}%`]);
  
  // Formatting Performance Report
  let range = perfSheet.getRange('A1:B' + perfSheet.getLastRow());
  range.setFontFamily('Arial');
  for (let i = 3; i <= perfSheet.getLastRow(); i++) 
  {
    let rowRange = perfSheet.getRange('A' + i + ':B' + i);
    if (i % 2 === 0) {
      rowRange.setBackground('#F8D5D3');
    } else {
      rowRange.setBackground('#F2ABA8');
    }
  }

  let profitCell = perfSheet.getRange('B5');
  profitCell.setFontColor(totalProfit >= 0 ? 'green' : 'red').setFontWeight('bold');

  let stockPercentageCell = perfSheet.getRange('B10');
  stockPercentageCell.setFontColor(stockPercentage >= 50 ? 'green' : 'red').setFontWeight('bold');

  let paidPercentageCell = perfSheet.getRange('B11');
  paidPercentageCell.setFontColor(paidOrderPercentage >= 50 ? 'green' : 'red').setFontWeight('bold');

  perfSheet.getRange('A1:B1').setFontSize(40).setFontWeight('bold').setBackground('white').setFontColor('black').setFontFamily("Schoolbell").setHorizontalAlignment('right');
  perfSheet.getRange('A2:B2').setBackground('white');
  perfSheet.getRange('A3:B' + perfSheet.getLastRow()).setFontSize(12).setHorizontalAlignment('left').setBorder(true, true, true, true, true, true, '#F6E5AE', SpreadsheetApp.BorderStyle.SOLID);
  perfSheet.getRange('A3:A' + perfSheet.getLastRow()).setFontWeight('bold').setBackground("#E9BF35");
  perfSheet.getRange('B3:B' + perfSheet.getLastRow()).setHorizontalAlignment('right');

  perfSheet.setRowHeight(1, 85).setColumnWidth(2, 420).autoResizeColumn(1);
}