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
   SAFE VERSION
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


  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }


  /*
   * IMPORTANT:
   * getDisplayValues() returns everything as strings.
   * This prevents Date / number / serialization
   * problems when sending data to the web app.
   */

  const values = sheet
    .getRange(
      1,
      1,
      lastRow,
      lastColumn
    )
    .getDisplayValues();


  const headers = values[0];


  return values
    .slice(1)

    .filter(function(row) {

      return row.some(function(value) {

        return String(value).trim() !== '';

      });

    })

    .map(function(row) {

      const item = {};


      headers.forEach(function(header, index) {

        const cleanHeader =
          String(header).trim();


        if (!cleanHeader) {
          return;
        }


        item[cleanHeader] =
          String(row[index] || '');

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

/* =========================================================
   ADD PRODUCT
========================================================= */

function addProduct(product) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    throw new Error('Products sheet not found.');
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error('Products sheet has no headers.');
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];


  /* -----------------------------------------
     GENERATE PRODUCT ID
  ----------------------------------------- */

  const productId = generateProductId(sheet);


  /* -----------------------------------------
     CREATE ROW
  ----------------------------------------- */

  const row = headers.map(function(header) {

    switch (header) {

      case 'Product ID':
        return productId;

      case 'Product Name':
        return product.productName || '';

      case 'Category':
        return product.category || '';

      case 'Sub Category':
        return product.subCategory || '';

      case 'Application':
        return product.application || '';

      case 'Industry':
        return product.industry || '';

      case 'Voltage':
        return product.voltage || '';

      case 'Size Range':
        return product.sizeRange || '';

      case 'Standard':
        return product.standard || '';

      case 'Key Features':
        return product.keyFeatures || '';

      case 'Sales Points':
        return product.salesPoints || '';

      case 'Customer Type':
        return product.customerType || '';

      case 'Competitors':
        return product.competitors || '';

      case 'Product Advantage':
        return product.productAdvantage || '';

      case 'Technical Notes':
        return product.technicalNotes || '';

      case 'Datasheet':
        return product.datasheet || '';

      case 'Catalogue':
        return product.catalogue || '';

      case 'Image':
        return product.image || '';

      case 'Status':
        return product.status || 'Active';

      case 'Last Updated':
        return new Date();

      default:
        return '';

    }

  });


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);


  return {
    success: true,
    productId: productId
  };

}


/* =========================================================
   GENERATE PRODUCT ID
========================================================= */

function generateProductId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'PH-0001';
  }


  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues();


  let maxNumber = 0;


  values.forEach(function(row) {

    const value = String(row[0] || '');

    const match =
      value.match(/^PH-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });


  return 'PH-' +
    String(maxNumber + 1)
      .padStart(4, '0');

}

/* =========================================================
   UPDATE PRODUCT
========================================================= */

function updateProduct(productId, product) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    throw new Error('Products sheet not found.');
  }


  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();


  if (lastRow < 2) {
    throw new Error('No products found.');
  }


  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];


  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();


  let targetRow = -1;


  ids.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(productId).trim()
    ) {

      targetRow = index + 2;

    }

  });


  if (targetRow === -1) {

    throw new Error(
      'Product not found: ' + productId
    );

  }


  const existingRow =
    sheet
      .getRange(
        targetRow,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  const row = headers.map(function(header, index) {

    switch (header) {

      case 'Product ID':
        return productId;

      case 'Product Name':
        return product.productName || '';

      case 'Category':
        return product.category || '';

      case 'Sub Category':
        return product.subCategory || '';

      case 'Application':
        return product.application || '';

      case 'Industry':
        return product.industry || '';

      case 'Voltage':
        return product.voltage || '';

      case 'Size Range':
        return product.sizeRange || '';

      case 'Standard':
        return product.standard || '';

      case 'Key Features':
        return product.keyFeatures || '';

      case 'Sales Points':
        return product.salesPoints || '';

      case 'Customer Type':
        return product.customerType || '';

      case 'Competitors':
        return product.competitors || '';

      case 'Product Advantage':
        return product.productAdvantage || '';

      case 'Technical Notes':
        return product.technicalNotes || '';

      case 'Datasheet':
        return product.datasheet || '';

      case 'Catalogue':
        return product.catalogue || '';

      case 'Image':
        return product.image || '';

      case 'Status':
        return product.status || 'Active';

      case 'Last Updated':
        return new Date();

      default:
        return existingRow[index];

    }

  });


  sheet
    .getRange(
      targetRow,
      1,
      1,
      row.length
    )
    .setValues([row]);


  return {
    success: true,
    productId: productId
  };

}


/* =========================================================
   DELETE PRODUCT
========================================================= */

function deleteProduct(productId) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    throw new Error('Products sheet not found.');
  }


  const lastRow = sheet.getLastRow();


  if (lastRow < 2) {
    throw new Error('No products found.');
  }


  const ids = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getDisplayValues();


  let targetRow = -1;


  ids.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(productId).trim()
    ) {

      targetRow = index + 2;

    }

  });


  if (targetRow === -1) {

    throw new Error(
      'Product not found: ' + productId
    );

  }


  sheet.deleteRow(targetRow);


  return {
    success: true,
    productId: productId
  };

}

/* =========================================================
   ADD APPLICATION
========================================================= */

function addApplication(application) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Applications');

  if (!sheet) {
    throw new Error('Applications sheet not found.');
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error('Applications sheet has no headers.');
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  const row = headers.map(function(header) {

    switch (String(header).trim()) {

      case 'Application':
        return application.application || '';

      case 'Industry':
        return application.industry || '';

      case 'Process / Area':
        return application.processArea || '';

      case 'Cable Required':
        return application.cableRequired || '';

      case 'Why Cable Is Required':
        return application.whyCableRequired || '';

      case 'Customer Type':
        return application.customerType || '';

      case 'Typical Buyer':
        return application.typicalBuyer || '';

      case 'Customer Requirement':
        return application.customerRequirement || '';

      case 'Opportunity Identification':
        return application.opportunityIdentification || '';

      case 'Questions to Ask':
        return application.questionsToAsk || '';

      case 'Key Selling Points':
        return application.keySellingPoints || '';

      case 'Technical Considerations':
        return application.technicalConsiderations || '';

      case 'Common Objections':
        return application.commonObjections || '';

      case 'Objection Answer':
        return application.objectionAnswer || '';

      case 'Related Products':
        return application.relatedProducts || '';

      case 'Status':
        return application.status || 'Active';

      case 'Last Updated':
        return new Date();

      default:
        return '';
    }

  });

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);

  return {
    success: true,
    application: application.application || ''
  };
}


/* =========================================================
   UPDATE APPLICATION
========================================================= */

function updateApplication(applicationName, application) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Applications');

  if (!sheet) {
    throw new Error('Applications sheet not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error('No applications found.');
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const applicationColumn =
    headers.findIndex(function(header) {
      return String(header).trim() === 'Application';
    });

  if (applicationColumn === -1) {
    throw new Error(
      'Application column not found.'
    );
  }

  const values = sheet
    .getRange(
      2,
      applicationColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  let targetRow = -1;

  values.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(applicationName).trim()
    ) {

      targetRow = index + 2;

    }

  });

  if (targetRow === -1) {
    throw new Error(
      'Application not found: ' +
      applicationName
    );
  }

  const existingRow =
    sheet
      .getRange(
        targetRow,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  const row = headers.map(function(header, index) {

    switch (String(header).trim()) {

      case 'Application':
        return application.application || '';

      case 'Industry':
        return application.industry || '';

      case 'Process / Area':
        return application.processArea || '';

      case 'Cable Required':
        return application.cableRequired || '';

      case 'Why Cable Is Required':
        return application.whyCableRequired || '';

      case 'Customer Type':
        return application.customerType || '';

      case 'Typical Buyer':
        return application.typicalBuyer || '';

      case 'Customer Requirement':
        return application.customerRequirement || '';

      case 'Opportunity Identification':
        return application.opportunityIdentification || '';

      case 'Questions to Ask':
        return application.questionsToAsk || '';

      case 'Key Selling Points':
        return application.keySellingPoints || '';

      case 'Technical Considerations':
        return application.technicalConsiderations || '';

      case 'Common Objections':
        return application.commonObjections || '';

      case 'Objection Answer':
        return application.objectionAnswer || '';

      case 'Related Products':
        return application.relatedProducts || '';

      case 'Status':
        return application.status || 'Active';

      case 'Last Updated':
        return new Date();

      default:
        return existingRow[index];
    }

  });

  sheet
    .getRange(
      targetRow,
      1,
      1,
      row.length
    )
    .setValues([row]);

  return {
    success: true,
    application:
      application.application || ''
  };
}


/* =========================================================
   DELETE APPLICATION
========================================================= */

function deleteApplication(applicationName) {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Applications');

  if (!sheet) {
    throw new Error('Applications sheet not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error('No applications found.');
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const applicationColumn =
    headers.findIndex(function(header) {
      return String(header).trim() === 'Application';
    });

  if (applicationColumn === -1) {
    throw new Error(
      'Application column not found.'
    );
  }

  const values = sheet
    .getRange(
      2,
      applicationColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  let targetRow = -1;

  values.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(applicationName).trim()
    ) {

      targetRow = index + 2;

    }

  });

  if (targetRow === -1) {
    throw new Error(
      'Application not found: ' +
      applicationName
    );
  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    application: applicationName
  };
}