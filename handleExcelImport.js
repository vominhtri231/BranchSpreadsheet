const START_ROW = 3;
const BRANCH_POSITION = 1;
const SALES_POSITION = 3;
const STOCK_POSITION = 6;

function onSubmit(branchMap) {
  const thisMonthItems = getThisMonthItems(branchMap);
  saveImportResult(thisMonthItems);
  sendSalesReport();
}

function getThisMonthItems(branchMap) {
  const lastMonthItems = getLastMonthItems();
  const thisMonthBranchMap = getThisMonthBranchMap(branchMap, lastMonthItems);
  const thisMonthItems = getItems(thisMonthBranchMap);
  // sort by sales in descending order
  thisMonthItems.sort((item1, item2) => item2.sales - item1.sales);
  return thisMonthItems;
}

function getLastMonthItems() {
  const sheet = SpreadsheetApp.getActiveSheet();

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) {
    // no old rows
    return [];
  }
  const lastRowValues = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, STOCK_POSITION).getValues();

  // array count from 0 instead of 1 in sheet
  return lastRowValues.map(values => ({
    branch: values[BRANCH_POSITION - 1],
    sales: values[SALES_POSITION - 1],
    stock: values[STOCK_POSITION - 1]
  }));
}

function getThisMonthBranchMap(branchMap, lastMonthItems) {
  return lastMonthItems
    .reduce((acc, lastMonthItem) => {
      const branch = lastMonthItem.branch;
      acc[branch] = acc[branch] || {};
      acc[branch] = {
        ...acc[branch],
        ...{
          lastMonthSales: lastMonthItem.sales,
          lastMonthStocks: lastMonthItem.stock
        }
      };
      return acc;
    }, branchMap);
}

function getItems(branchMap) {
  const items = [];
  for (const branch in branchMap) {
    if (!branch || branch === 'undefined') {
      continue;
    }
    items.push({ ...branchMap[branch], branch: branch });
  }
  return items;
}

function saveImportResult(items) {
  const sheet = SpreadsheetApp.getActiveSheet();
  deleteOldRows(sheet);
  const savingValues = getSavingValues(items);
  saveNewRows(sheet, savingValues);
  SpreadsheetApp.flush();
}

function deleteOldRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) {
    // no old rows
    return;
  }
  sheet.deleteRows(START_ROW, lastRow - START_ROW + 1);
}

function getSavingValues(items) {
  return items.map(item => {
    const row =
      [item.branch, , item.sales, item.cost, ,
      item.stock, item.lastMonthSales, item.lastMonthStocks, '', ''];

    const margin = calculatePercentage(item.sales - item.cost, item.cost);
    if (margin) {
      row[4] = margin;
    }

    
    const salesTrend = calculatePercentage(item.sales, item.lastMonthSales);
    Logger.log(item.sale, item.lastMonthSales, salesTrend);
    if (salesTrend) {
      row[8] = salesTrend;
    }

    
    const stockTrend = calculatePercentage(item.stock, item.lastMonthStocks);
    Logger.log(item.stock, item.lastMonthStocks, stockTrend);
    if (stockTrend) {
      row[9] = stockTrend;
    }

    return row;
  });
}

function calculatePercentage(numerator, denominator) {
  if (!numerator || !denominator || denominator === 0) {
    return;
  }
  return numerator / denominator;
  return Math.round(((numerator / denominator) + Number.EPSILON) * 100) / 100;
}

function saveNewRows(sheet, savingValues) {
  const numberOfRow = savingValues.length;
  setFormat(sheet, numberOfRow);
  const savingRange = sheet.getRange(START_ROW, 1, numberOfRow, 10);
  savingRange.setValues(savingValues);
}

function setFormat(sheet, numberOfRow) {
  const branchRange = sheet.getRange(START_ROW, 1, numberOfRow, 1);
  branchRange.setNumberFormat("@");
}
