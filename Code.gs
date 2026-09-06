/* =========================================================
   PRODUCTS HUB
   Code.gs
   Backend / Google Sheets connection
========================================================= */


/* =========================================================
   CONFIGURATION
========================================================= */

const SHEET_NAMES = [
  'Products',
  'Applications',
  'Industries',
  'Competitors',
  'Cable Knowledge',
  'Learning',
  'Career Growth',
  'Documents'
];


/* =========================================================
   WEB APP
========================================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Products Hub')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* =========================================================
   INCLUDE HTML FILES
========================================================= */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =========================================================
   GET SPREADSHEET
========================================================= */

function getSpreadsheet() {

  return SpreadsheetApp.getActiveSpreadsheet();

}


/* =========================================================
   CHECK REQUIRED SHEETS
========================================================= */

function checkSheets() {

  const ss = getSpreadsheet();

  const result = {};

  SHEET_NAMES.forEach(function(sheetName) {

    const sheet = ss.getSheetByName(sheetName);

    result[sheetName] = !!sheet;

  });

  return result;

}


/* =========================================================
   GET SHEET DATA
========================================================= */

function getSheetData(sheetName) {

  const ss = getSpreadsheet();

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' + sheetName
    );
  }


  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();


  if (lastRow < 1 || lastColumn < 1) {
    return [];
  }


  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();


  if (values.length < 2) {
    return [];
  }


  const headers = values[0];


  return values
    .slice(1)
    .filter(function(row) {

      return row.some(function(value) {

        return value !== '' &&
               value !== null &&
               value !== undefined;

      });

    })
    .map(function(row) {

      const item = {};

      headers.forEach(function(header, index) {

        if (header !== '') {

          item[header] = row[index];

        }

      });

      return item;

    });

}


/* =========================================================
   DASHBOARD COUNTS
========================================================= */

function getDashboardData() {

  const ss = getSpreadsheet();

  const result = {};

  SHEET_NAMES.forEach(function(sheetName) {

    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {

      result[sheetName] = 0;

      return;

    }


    const lastRow = sheet.getLastRow();

    result[sheetName] =
      Math.max(0, lastRow - 1);

  });


  return result;

}


/* =========================================================
   PRODUCTS
========================================================= */

function getProducts() {

  return getSheetData('Products');

}


/* =========================================================
   APPLICATIONS
========================================================= */

function getApplications() {

  return getSheetData('Applications');

}


/* =========================================================
   INDUSTRIES
========================================================= */

function getIndustries() {

  return getSheetData('Industries');

}


/* =========================================================
   COMPETITORS
========================================================= */

function getCompetitors() {

  return getSheetData('Competitors');

}


/* =========================================================
   CABLE KNOWLEDGE
========================================================= */

function getCableKnowledge() {

  return getSheetData('Cable Knowledge');

}


/* =========================================================
   LEARNING
========================================================= */

function getLearning() {

  return getSheetData('Learning');

}


/* =========================================================
   CAREER GROWTH
========================================================= */

function getCareerGrowth() {

  return getSheetData('Career Growth');

}


/* =========================================================
   DOCUMENTS
========================================================= */

function getDocuments() {

  return getSheetData('Documents');

}


/* =========================================================
   TEST CONNECTION
========================================================= */

function testConnection() {

  const ss = getSpreadsheet();

  return {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    sheets: checkSheets()
  };

}