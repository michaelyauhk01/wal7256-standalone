// Envs? Maybe? This stuff is to map the values on the spreadsheet
const config = {
  spreadsheet: {
    //For resetting the spreadsheet with the original copy (testing purpose)
    originalOrderMgmtReport: {
      url: "https://docs.google.com/spreadsheets/d/1ladSAv2Z-2f3Ri-VCGEuckFboEARi8J35A1nEXBxyrM/edit",
      sheet: {
        tab1: "工作表1",
      },
    },
    orderMgmtReport: {
      url: "https://docs.google.com/spreadsheets/d/1GjJmrkOiuyxAyF6EM7ZgJF_RPI0uB80B8JI3GnYATzM/edit",
      sheet: {
        tab1: "工作表1",
        mapMerchant: "Map Merchant",
        mapCOG: "Map COG"
      },
      column: {
        commission: "佣金",
        cog: "單件COG",
        settlePrice: "結算價",
        orderStatus: "訂單狀態",
        deliveryFee: "運費",
        childOrderId: "子訂單ID",
        parentOrderId: "主訂單ID",
        shopId: "店鋪ID"
      }
    },
    settleReport: {
      url: "https://docs.google.com/spreadsheets/d/1X7vtKzmoay2_uU2vQXlPsPOyA-BYx1HmZJmEPha_e0A/edit",
      sheet: {
        mapMerchant: "Map Merchant",
        consignmentMerchant: "Consignment Merchant"
      },
      column: {
        commissionRate: "Commission Rate",
        shopId: "01mall Shop #"
      }
    }
  }
}

// Helper to provide wrapper of some app script APIs
const helper = {
  // I can't find any native methods to get the position of a column
  // currently for loop and search by name/text
  getColumnPos: (columnName = '', sheet) => {
    const data = sheet.getDataRange().getValues()
    const columnList = data[0]

    let index = null;
    for (let i = 0; i < columnList.length; i++) {
      if (columnList[i] === columnName) {
        index = i;
        break;
      }
    }

    if (index === null) {
      throw new Error("Column not found")
    }

    return index;
  },
  // the native insertSheet() does not check duplicate/ does not return the Sheet instance after insert
  // return Sheet instance is useful
  createSheetBySpreadsheet: (spreadSheet, sheetName) => {
    if (typeof sheetName !== 'string') {
      sheetName = `${sheetName}`
    }

    const sheet = spreadSheet.getSheetByName(sheetName)

    if (sheet) {
      return sheet
    }

    spreadSheet.insertSheet(sheetName)
    return spreadSheet.getSheetByName(sheetName)
  },
  // the native copyTo() does not support cross-spreadsheet copy
  // this one store the entire sheet in memory and set the value into another
  copyTo: (from, to) => {
    const source = from.getDataRange().getValues()
    const target = to.getRange(1, 1, source.length, source[1].length)
    target.setValues(source)
  },
  // delete a row by a column value
  deleteRowByColumnValue: (sheet, column, value) => {
    if (!sheet) {
      throw new Error("Unknown sheet")
    }

    const data = sheet.getDataRange().getValues()
    const firstRow = data[0]
    let colIdx = null;

    for (let i = 0; i < firstRow.length; i++) {
      if (firstRow[i] === column) {
        colIdx = i
        break
      }
    }

    if (colIdx === null) {
      throw new Error(`Column ${value} not found`)
    }

    let rowsDeleted = 0;

    for (let i = 0; i < data.length; i++) {
      const colValue = data[i][colIdx]

      if (colValue === value) {

        sheet.deleteRow(i + 1 - rowsDeleted)
        rowsDeleted++;
      }
    }
  }
}

/**
 * Append three columns to 01網購_訂單管理報表Template
 */
function createColumns() {
  const orderReportSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url)
  const orderSheet = orderReportSpreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1)

  const lastColIdx = orderSheet.getLastColumn();

  const newColumns = [
    config.spreadsheet.orderMgmtReport.column.commission,
    config.spreadsheet.orderMgmtReport.column.cog,
    config.spreadsheet.orderMgmtReport.column.settlePrice,
  ]

  for (let i = 0; i < newColumns.length; i++) {
    const range = orderSheet.getRange(1, lastColIdx + i + 1)
    range.setValue(newColumns[i])
  }
}


/**
 * Create sheet named "Map Merchant" and "Map COG"
 */
function createSheet() {
  const spreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);
  helper.createSheetBySpreadsheet(spreadSheet, config.spreadsheet.orderMgmtReport.sheet.mapMerchant)
  helper.createSheetBySpreadsheet(spreadSheet, config.spreadsheet.orderMgmtReport.sheet.mapCOG)
}

/**
 * Copy Map Merchant sheet from 01mall 結算 | Checking log sheet to 01網購_訂單管理報表Template
 */
function copyMapMerchant() {
  const orderSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);
  const settleSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.settleReport.url);

  const mapMerchantFromSettle = settleSpreadSheet.getSheetByName(config.spreadsheet.settleReport.sheet.mapMerchant)
  const mapMerchantFromOrder = orderSpreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.mapMerchant)

  helper.copyTo(mapMerchantFromSettle, mapMerchantFromOrder)
}

/**
 * Copy Consignment Merchant sheet from 01mall 結算 | Checking log sheet to 01網購_訂單管理報表Template
 */
function copyMapCog() {
  const orderSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);
  const settleSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.settleReport.url);

  const mapCogtFromSettle = settleSpreadSheet.getSheetByName(config.spreadsheet.settleReport.sheet.consignmentMerchant)
  const mapCogFromOrder = orderSpreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.mapCOG)

  helper.copyTo(mapCogtFromSettle, mapCogFromOrder)
}

/**
 * Delete the order(row) if its status is either "已取消(主動)" or "已取消(自動)"
 */
function deleteOrderByStatus() {
  const OrderStatus = {
    MANUAL_CANCELLED: "已取消(主動)",
    AUTO_CANCELLED: "已取消(自動)",
  }

  const spreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);  // get Spreadsheet and Sheet instance
  const sheet = spreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1);

  helper.deleteRowByColumnValue(sheet, config.spreadsheet.orderMgmtReport.column.orderStatus, OrderStatus.AUTO_CANCELLED)
  helper.deleteRowByColumnValue(sheet, config.spreadsheet.orderMgmtReport.column.orderStatus, OrderStatus.MANUAL_CANCELLED)
}

/**
 * For duplicated order IDs, set the rest of its delivery to 0 if there is one >0
 * If all delivery fees are 0, leave them to be 0
 */
function removeDuplicatedDeliveryFee() {
  const spreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);  // get Spreadsheet and Sheet instance
  const sheet = spreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1);

  const data = sheet.getDataRange().getValues()
  const delieryFeeColIdx = helper.getColumnPos(config.spreadsheet.orderMgmtReport.column.deliveryFee, sheet)

  const orderIdColIdx = helper.getColumnPos(config.spreadsheet.orderMgmtReport.column.parentOrderId, sheet)

  let scannedOrder = [];

  for (let i = 0; i < data.length; i++) {
    const orderId = data[i][orderIdColIdx]
    const deliveryFee = data[i][delieryFeeColIdx]

    if (orderId !== scannedOrder[scannedOrder.length - 1]?.id) {
      let hasFee = false;

      for (let j = 0; j < scannedOrder.length; j++) {
        const range = sheet.getRange(scannedOrder[j].row + 1, delieryFeeColIdx + 1)

        const fee = scannedOrder[j].fee
        if (fee !== '$0.00' && hasFee === false) {
          hasFee = true;
          continue;
        }

        if (hasFee) {
          range.setValue(0)
        }
      }

      scannedOrder = [{ row: i, id: orderId, fee: deliveryFee }]
      continue;
    }

    scannedOrder.push({
      row: i,
      id: orderId,
      fee: deliveryFee
    })
  }
}

/**
 * Copy Commision Rate from one spreadsheet to another
 */
function copyCommissionRate() {
  //Retrieve all sheet instances
  const settleReportSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.settleReport.url)
  const mapMerchantSheet = settleReportSpreadSheet.getSheetByName(config.spreadsheet.settleReport.sheet.mapMerchant)
  const orderReportSpreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url)
  const orderSheet = orderReportSpreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1)

  //Use helper to get specific column index/position
  const settleReportCols = {
    comissionRate: helper.getColumnPos(config.spreadsheet.settleReport.column.commissionRate, mapMerchantSheet),
    shopId: helper.getColumnPos(config.spreadsheet.settleReport.column.shopId, mapMerchantSheet)
  }

  const orderReportCols = {
    commission: helper.getColumnPos(config.spreadsheet.orderMgmtReport.column.commission, orderSheet),
    shopId: helper.getColumnPos(config.spreadsheet.orderMgmtReport.column.shopId, orderSheet)
  }

  // Get sheets data in 2D array
  const settleSheetData = mapMerchantSheet.getDataRange().getValues();
  const orderSheetData = orderSheet.getDataRange().getValues();


  // Match shop ID and copy the commission rate
  for (let i = 0; i < settleSheetData.length; i++) {
    const settleCommRate = settleSheetData[i][settleReportCols.comissionRate]
    const settleShopId = settleSheetData[i][settleReportCols.shopId]
    for (let j = 0; j < orderSheetData.length; j++) {
      const orderShopId = orderSheetData[j][orderReportCols.shopId];

      if (settleShopId === orderShopId) {
        const range = orderSheet.getRange(j + 1, orderReportCols.commission + 1)
        range.setValue(settleCommRate)
        break;
      }
    }
  }
}

/**
 * By shop Id (店鋪ID) create a sheet under the same spreadsheet with the ID as the sheet name
 */
function createSpreadsheetByShopId() {
  const spreadSheet = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url);  // get Spreadsheet and Sheet instance
  const sheet = spreadSheet.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1);

  const shopIdColPos = helper.getColumnPos(config.spreadsheet.orderMgmtReport.column.shopId, sheet); // Search where '店鋪id' is on the first row (column position)

  const data = sheet.getDataRange().getValues()  // Get sheet data in 2D Array

  const shopIdSet = new Set();// to remove duplicated id in array

  for (let i = 0; i < data.length; i++) {
    const shopId = data[i][shopIdColPos]

    if (shopId === config.spreadsheet.orderMgmtReport.column.shopId) { //skip first column
      continue;
    }

    shopIdSet.add(shopId)
  }

  const uniqueShopIds = [...shopIdSet]

  for (let i = 0; i < uniqueShopIds.length; i++) {
    helper.createSheetBySpreadsheet(spreadSheet, `${uniqueShopIds[i]}`) // itereate create spreadsheet by shop id
  }
}

/**
 * Restore spreadsheet back to original
 */
function restoreMockData() {
  const defaultMock = SpreadsheetApp.openByUrl(config.spreadsheet.originalOrderMgmtReport.url)
  const from = defaultMock.getSheetByName(config.spreadsheet.originalOrderMgmtReport.sheet.tab1)
  const target = SpreadsheetApp.openByUrl(config.spreadsheet.orderMgmtReport.url)
  const to = target.getSheetByName(config.spreadsheet.orderMgmtReport.sheet.tab1)
  helper.copyTo(from, to)
}