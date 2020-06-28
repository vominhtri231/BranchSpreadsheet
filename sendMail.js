const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1B50ODPwoY_T8L4E4_0A-hNtUGylKyyKmjcWquqjnccg/edit";
const SALES_SHEET = "売上・棚卸";
const RECIPIENT_SHEET = "アドレス";
const SUBJECT = "ブランド毎売上レポート";

function sendSalesReport() {
  const emailBody = getEmailBody();
  getRecipients().forEach(recipient => {
    MailApp.sendEmail({
      to: recipient,
      subject: SUBJECT,
      htmlBody: emailBody,
    });
  });
}

function getEmailBody() {
  const emailTemplate = HtmlService.createTemplateFromFile('salesEmailTemplate');
  emailTemplate.items = getSalesItems();
  emailTemplate.spreadsheetUrl = SPREADSHEET_URL;
  return emailTemplate.evaluate().getContent();
}

function getSalesItems() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = spreadsheet.getSheetByName(SALES_SHEET);
  return sheet.getDataRange().getValues().slice(1)
    .map(values => ({
      branch: values[0],
      sales: getCurrencyValue(values[2]),
      cost: getRoundedCurrencyValue(values[3]),
      margin: getPercentageValue(values[4]),
      stock: getRoundedCurrencyValue(values[5]),
      lastMonthSales: getCurrencyValue(values[6]),
      lastMonthStock: getRoundedCurrencyValue(values[7]),
      salesTrend: getPercentageValue(values[8]),
      stockTrend: getPercentageValue(values[9]),
    }));
}

function getRoundedCurrencyValue(value) {
  return getCurrencyValue(Math.ceil(value));;
}

function getCurrencyValue(value) {
  return String(value).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
}

function getPercentageValue(value) {
  return Math.ceil(value * 100) + "%";
}

function getRecipients() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = spreadsheet.getSheetByName(RECIPIENT_SHEET);
  return sheet.getDataRange().getValues().map(values => values[1]);
}