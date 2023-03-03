const MOCK_RESTORE = "https://docs.google.com/spreadsheets/d/1ladSAv2Z-2f3Ri-VCGEuckFboEARi8J35A1nEXBxyrM/edit"
const HK01MALL_ORDER_MGMT_REPORT_URL = "https://docs.google.com/spreadsheets/d/1GjJmrkOiuyxAyF6EM7ZgJF_RPI0uB80B8JI3GnYATzM/edit"
const HK01MALL_ORDER_MGMT_REPORT_SHEET_NAME = "工作表1"
const HK01MALL_SETTLE_REPORT_URL = "https://docs.google.com/spreadsheets/d/1X7vtKzmoay2_uU2vQXlPsPOyA-BYx1HmZJmEPha_e0A/edit"

const COL_COMMISSION = "佣金"
const COL_COG = "單件COG"
const COL_SETTLEMENT_PRICE = "結算價"
const COL_ORDER_STATUS = "訂單狀態"
const COL_DELIVERY_FEE = "運費"
const COL_CHILD_ORDER_ID = "子訂單ID"
const COL_ORDER_ID = "主訂單ID"

const OrderStatus = {
  MANUAL_CANCELLED: "已取消(主動)",
  AUTO_CANCELLED: "已取消(自動)",
  FINISHED: "已完成",
  PENDING_FOR_TAKEN: "待取貨",
  PENDING_FOR_SENT: "待發貨"
}

const Helper = {
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

    if (!index) {
      throw new Error("Column not found")
    }

    return index;
  }
}

function appendColumns(sheet, columnNames = []) {
  if (!sheet) {
    throw new Error("Unknown sheet")
  }

  if (columnNames.length <= 0) {
    return;
  }

  const lastCol = sheet.getLastColumn();

  columnNames.forEach((name, i) => {
    const range = sheet.getRange(1, lastCol + i + 1)
    range.setValue(name)
  })
}

function deleteOrderByStatus(sheet, status) {
  if (!sheet) {
    throw new Error("Unknown sheet")
  }

  const data = sheet.getDataRange().getValues()
  const firstRow = data[0]
  let invoiceStatusCellIdx = null;

  for (let i = 0; i < firstRow.length; i++) {
    if (firstRow[i] === COL_ORDER_STATUS) {
      invoiceStatusCellIdx = i
      break
    }
  }

  if (invoiceStatusCellIdx === null) {
    const msg = `Column ${COL_ORDER_STATUS} not found`
    throw new Error(msg)
  }

  let rowsDeleted = 0;

  for (let i = 0; i < data.length; i++) {
    const orderStatus = data[i][invoiceStatusCellIdx]

    if (orderStatus === status) {

      sheet.deleteRow(i + 1 - rowsDeleted)
      rowsDeleted++;
    }
  }
}

function removeDuplicatedDeliveryFee(sheet) {
  if (!sheet) {
    throw new Error("unknown sheet")
  }

  const data = sheet.getDataRange().getValues()
  const delieryFeeColIdx = Helper.getColumnPos(COL_DELIVERY_FEE, sheet)
  const orderIdColIdx = Helper.getColumnPos(COL_CHILD_ORDER_ID, sheet)

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

function createSheetBySpreadsheet(spreadSheet, sheetName) {
  const sheet = spreadSheet.getSheetByName(sheetName)

  if (sheet) {
    return sheet
  }

  spreadSheet.insertSheet(sheetName)
  return spreadSheet.getSheetByName(sheetName)
}

// copyTo doesn't work across different spreadsheet
// need to make our own copy function
function copyTo(from, to) {
  const source = from.getDataRange().getValues()
  const range = to.getRange(1, 1, source.length, source[1].length)
  range.setValues(source)
}

function restoreMockData() {
  //for reset the example spreadsheet to test
  const defaultMock = SpreadsheetApp.openByUrl(MOCK_RESTORE)
  const from = defaultMock.getSheetByName(HK01MALL_ORDER_MGMT_REPORT_SHEET_NAME)
  const target = SpreadsheetApp.openByUrl(HK01MALL_ORDER_MGMT_REPORT_URL)
  const to = target.getSheetByName(HK01MALL_ORDER_MGMT_REPORT_SHEET_NAME)

  copyTo(from, to)
}

function main() {
  const hk01MallOrderMgmtSS = SpreadsheetApp.openByUrl(HK01MALL_ORDER_MGMT_REPORT_URL)
  const orderSheet = hk01MallOrderMgmtSS.getSheetByName(HK01MALL_ORDER_MGMT_REPORT_SHEET_NAME)

  // Add text [佣金] [單件COG] [結算價] at AQ2; AR2; AS2
  // appendColumns(orderSheet, [COL_COMMISSION, COL_COG, COL_SETTLEMENT_PRICE])

  // Base on column O data, remove roll with 已取消(自動) 已取消(主動) in column O 
  // deleteOrderByStatus(orderSheet, OrderStatus.AUTO_CANCELLED)
  // deleteOrderByStatus(orderSheet, OrderStatus.MANUAL_CANCELLED)

  // Create two new tab with name [MAP Merchant] [MAP COG]
  const mapMerchantSheetDest = createSheetBySpreadsheet(hk01MallOrderMgmtSS, "MAP Merchant")
  const mapCogSheetDest = createSheetBySpreadsheet(hk01MallOrderMgmtSS, "MAP COG")

  // Copy from 01mall 結算 | Checking log sheet into 01網購_訂單管理報表Template
  const hk01MallSettleSS = SpreadsheetApp.openByUrl(HK01MALL_SETTLE_REPORT_URL)
  console.log(hk01MallSettleSS.getName())
  const mapMerchantSheet = hk01MallSettleSS.getSheetByName("MAP Merchant")
  const mapCogSheet = hk01MallSettleSS.getSheetByName("Consignment Merchant")

  // copyTo(mapMerchantSheet, mapMerchantSheetDest)
  // copyTo(mapCogSheet, mapCogSheetDest)

  // Remove duplicated order by delivery fee 
  // If any of orders has delivery fee > 0, remove orders except that one
  // If none of orders have delivery fee > 0, use any one of them
  removeDuplicatedDeliveryFee(orderSheet)

  // restoreMockData()
}