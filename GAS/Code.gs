function myFunction() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let menuSheet = spreadsheet.getSheetByName("メニュー");
  let customerSheet = spreadsheet.getSheetByName("顧客データ");

  let customerName = menuSheet.getRange("B3").getValue();

  let newSheetName = `請求書(${customerName})`;
  let invoiceSheet = spreadsheet.getSheetByName(newSheetName);
  
  if (invoiceSheet != null) {
    spreadsheet.deleteSheet(invoiceSheet);
  }

  spreadsheet.setSpreadsheetTimeZone("Asia/Tokyo");
  let date = menuSheet.getRange("B6").getValues();

  let invoice = spreadsheet.getSheetByName("請求書テンプレート").copyTo(spreadsheet);
  invoice.activate();
  
  invoice.setName(newSheetName);

  let customerData = customerSheet.getDataRange().getValues();

  let customerAddress = "";
  // ifの中で直接定義してもエラー(Not Defined)

  for (let j = 1; j < customerData.length; ++j) {
    if (customerData[j][1] == customerName) {
      customerAddress = customerData[j][4];
      // Logger.log(customerData[j][1]);
      // Logger.log(customerData[j][4]);
      break;
    }
  }


  invoice.getRange("B1").setValue(date);
  invoice.getRange("A4").setValue(customerName);

  let data = spreadsheet.getSheetByName("注文データ").getDataRange().getValues();

  for (let i = 1, m = 11; i < data.length; ++i) {
    // 注文データのi行目D列の値が、設定した企業名と同じだったら
    if (data[i][3] == customerName) {
      invoice.getRange(3, 1).setValue(customerAddress);

      //日付(B->A)、商品名(G->B)、単価(H->E)、個数(I->F)、金額(K->G)
      invoice.getRange(m, 1).setValue(data[i][1]);
      invoice.getRange(m, 2).setValue(data[i][6]);
      invoice.getRange(m, 5).setValue(data[i][7]);
      invoice.getRange(m, 6).setValue(data[i][8]);
      invoice.getRange(m, 7).setValue(data[i][9]);
      // getRangeで行、列を指定するときは1スタート
      ++m;
      }
  }


}

function deleteFunction() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let menuSheet = spreadsheet.getSheetByName("メニュー");
  let customerName = menuSheet.getRange("B3").getValue();

  let invoice = spreadsheet.getSheetByName(`請求書(${customerName})`);
  
  if (invoice != null) {
    spreadsheet.deleteSheet(invoice);
  }
  // else {
  //   let invoice = spreadsheet.getSheetByName("請求書テンプレート").copyTo(spreadsheet);
  //   invoice.activate();
  //   invice.setName(`請求書(${customerName})`);
  // }
}
