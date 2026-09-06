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
   APPLICATIONS
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE APPLICATION ID
========================================================= */

function generateApplicationId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'AP-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });

  if (idColumn === -1) {

    throw new Error(
      'Application ID column not found.'
    );

  }

  const ids =
    sheet
      .getRange(
        2,
        idColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^AP-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'AP-' +
    String(maxNumber + 1)
      .padStart(4, '0');

}


/* =========================================================
   ADD APPLICATION
========================================================= */

function addApplication(application) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Applications');


  if (!sheet) {

    throw new Error(
      'Applications sheet not found.'
    );

  }


  const lastColumn =
    sheet.getLastColumn();


  if (lastColumn === 0) {

    throw new Error(
      'Applications sheet has no headers.'
    );

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  /* -----------------------------------------
     GENERATE APPLICATION ID
  ----------------------------------------- */

  const applicationId =
    generateApplicationId(sheet);


  /* -----------------------------------------
     CREATE ROW
  ----------------------------------------- */

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Application ID':
          return applicationId;

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

    applicationId:
      applicationId,

    application:
      application.application || ''

  };

}


/* =========================================================
   FIND APPLICATION ROW
   Accepts Application ID OR Application Name
========================================================= */

function findApplicationRow(
  sheet,
  applicationKey
) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  if (lastRow < 2) {
    return -1;
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });


  const applicationColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application';

    });


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    /* Check Application ID */

    if (
      idColumn !== -1 &&
      String(
        values[i][idColumn] || ''
      ).trim() ===
      String(applicationKey).trim()
    ) {

      return i + 2;

    }


    /* Check Application Name */

    if (
      applicationColumn !== -1 &&
      String(
        values[i][applicationColumn] || ''
      ).trim() ===
      String(applicationKey).trim()
    ) {

      return i + 2;

    }

  }


  return -1;

}


/* =========================================================
   UPDATE APPLICATION
   Accepts Application ID OR Application Name
========================================================= */

function updateApplication(
  applicationKey,
  application
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Applications');


  if (!sheet) {

    throw new Error(
      'Applications sheet not found.'
    );

  }


  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  if (lastRow < 2) {

    throw new Error(
      'No applications found.'
    );

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const targetRow =
    findApplicationRow(
      sheet,
      applicationKey
    );


  if (targetRow === -1) {

    throw new Error(
      'Application not found: ' +
      applicationKey
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


  const row =
    headers.map(function(
      header,
      index
    ) {

      switch (
        String(header).trim()
      ) {

        /* KEEP EXISTING ID */

        case 'Application ID':
          return existingRow[index] || '';

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


  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });


  const finalApplicationId =
    idColumn !== -1
      ? sheet
          .getRange(
            targetRow,
            idColumn + 1
          )
          .getDisplayValue()
      : '';


  return {

    success: true,

    applicationId:
      finalApplicationId,

    application:
      application.application || ''

  };

}


/* =========================================================
   DELETE APPLICATION
   Accepts Application ID OR Application Name
========================================================= */

function deleteApplication(
  applicationKey
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Applications');


  if (!sheet) {

    throw new Error(
      'Applications sheet not found.'
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      'No applications found.'
    );

  }


  const targetRow =
    findApplicationRow(
      sheet,
      applicationKey
    );


  if (targetRow === -1) {

    throw new Error(
      'Application not found: ' +
      applicationKey
    );

  }


  const lastColumn =
    sheet.getLastColumn();


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });


  let applicationId = '';


  if (idColumn !== -1) {

    applicationId =
      sheet
        .getRange(
          targetRow,
          idColumn + 1
        )
        .getDisplayValue();

  }


  const applicationColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application';

    });


  let applicationName = '';


  if (applicationColumn !== -1) {

    applicationName =
      sheet
        .getRange(
          targetRow,
          applicationColumn + 1
        )
        .getDisplayValue();

  }


  sheet.deleteRow(targetRow);


  return {

    success: true,

    applicationId:
      applicationId,

    application:
      applicationName

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

/* =========================================================
   INDUSTRIES
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE INDUSTRY ID
========================================================= */

function generateIndustryId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'IN-0001';
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  let maxNumber = 0;


  values.forEach(function(row) {

    const value =
      String(row[0] || '').trim();


    const match =
      value.match(/^IN-(\d+)$/i);


    if (match) {

      const number =
        parseInt(match[1], 10);


      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });


  return 'IN-' +
    String(maxNumber + 1)
      .padStart(4, '0');

}


/* =========================================================
   ADD INDUSTRY
========================================================= */

function addIndustry(industry) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Industries');


  if (!sheet) {
    throw new Error(
      'Industries sheet not found.'
    );
  }


  const lastColumn =
    sheet.getLastColumn();


  if (lastColumn === 0) {
    throw new Error(
      'Industries sheet has no headers.'
    );
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  const industryId =
    generateIndustryId(sheet);


  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Industry ID':
          return industryId;

        case 'Industry':
          return industry.industry || '';

        case 'Major Segments':
          return industry.majorSegments || '';

        case 'Typical Projects':
          return industry.typicalProjects || '';

        case 'Cables Used':
          return industry.cablesUsed || '';

        case 'Key Applications':
          return industry.keyApplications || '';

        case 'Customer Types':
          return industry.customerTypes || '';

        case 'Typical Buyers':
          return industry.typicalBuyers || '';

        case 'Decision Makers':
          return industry.decisionMakers || '';

        case 'Customer Requirements':
          return industry.customerRequirements || '';

        case 'Opportunity Identification':
          return industry.opportunityIdentification || '';

        case 'Questions to Ask':
          return industry.questionsToAsk || '';

        case 'Key Selling Points':
          return industry.keySellingPoints || '';

        case 'Technical Considerations':
          return industry.technicalConsiderations || '';

        case 'Common Objections':
          return industry.commonObjections || '';

        case 'Objection Answers':
          return industry.objectionAnswers || '';

        case 'Related Products':
          return industry.relatedProducts || '';

        case 'Status':
          return industry.status || 'Active';

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

    industryId:
      industryId,

    industry:
      industry.industry || ''

  };

}


/* =========================================================
   UPDATE INDUSTRY
========================================================= */

function updateIndustry(
  industryId,
  industry
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Industries');


  if (!sheet) {
    throw new Error(
      'Industries sheet not found.'
    );
  }


  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  if (lastRow < 2) {
    throw new Error(
      'No industries found.'
    );
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Industry ID';

    });


  if (idColumn === -1) {

    throw new Error(
      'Industry ID column not found.'
    );

  }


  const ids =
    sheet
      .getRange(
        2,
        idColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  let targetRow = -1;


  ids.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(industryId).trim()
    ) {

      targetRow =
        index + 2;

    }

  });


  if (targetRow === -1) {

    throw new Error(
      'Industry not found: ' +
      industryId
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


  const row =
    headers.map(function(
      header,
      index
    ) {

      switch (
        String(header).trim()
      ) {

        case 'Industry ID':
          return industryId;

        case 'Industry':
          return industry.industry || '';

        case 'Major Segments':
          return industry.majorSegments || '';

        case 'Typical Projects':
          return industry.typicalProjects || '';

        case 'Cables Used':
          return industry.cablesUsed || '';

        case 'Key Applications':
          return industry.keyApplications || '';

        case 'Customer Types':
          return industry.customerTypes || '';

        case 'Typical Buyers':
          return industry.typicalBuyers || '';

        case 'Decision Makers':
          return industry.decisionMakers || '';

        case 'Customer Requirements':
          return industry.customerRequirements || '';

        case 'Opportunity Identification':
          return industry.opportunityIdentification || '';

        case 'Questions to Ask':
          return industry.questionsToAsk || '';

        case 'Key Selling Points':
          return industry.keySellingPoints || '';

        case 'Technical Considerations':
          return industry.technicalConsiderations || '';

        case 'Common Objections':
          return industry.commonObjections || '';

        case 'Objection Answers':
          return industry.objectionAnswers || '';

        case 'Related Products':
          return industry.relatedProducts || '';

        case 'Status':
          return industry.status || 'Active';

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

    industryId:
      industryId,

    industry:
      industry.industry || ''

  };

}


/* =========================================================
   DELETE INDUSTRY
========================================================= */

function deleteIndustry(
  industryId
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Industries');


  if (!sheet) {
    throw new Error(
      'Industries sheet not found.'
    );
  }


  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  if (lastRow < 2) {
    throw new Error(
      'No industries found.'
    );
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Industry ID';

    });


  if (idColumn === -1) {

    throw new Error(
      'Industry ID column not found.'
    );

  }


  const ids =
    sheet
      .getRange(
        2,
        idColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  let targetRow = -1;


  ids.forEach(function(row, index) {

    if (
      String(row[0]).trim() ===
      String(industryId).trim()
    ) {

      targetRow =
        index + 2;

    }

  });


  if (targetRow === -1) {

    throw new Error(
      'Industry not found: ' +
      industryId
    );

  }


  sheet.deleteRow(targetRow);


  return {

    success: true,

    industryId:
      industryId

  };

}

/* =========================================================
   ONE-TIME: FILL MISSING INDUSTRY IDS
========================================================= */

function fillMissingIndustryIds() {

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Industries');

  if (!sheet) {
    throw new Error('Industries sheet not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    return {
      success: true,
      updated: 0
    };
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Industry ID';
  });

  const nameColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Industry';
  });

  if (idColumn === -1) {
    throw new Error('Industry ID column not found.');
  }

  if (nameColumn === -1) {
    throw new Error('Industry column not found.');
  }

  const ids = sheet
    .getRange(
      2,
      idColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  const names = sheet
    .getRange(
      2,
      nameColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^IN-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }
    }

  });

  let updated = 0;

  ids.forEach(function(row, index) {

    const currentId =
      String(row[0] || '').trim();

    const industryName =
      String(names[index][0] || '').trim();

    if (!currentId && industryName) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'IN-' +
          String(maxNumber).padStart(4, '0')
        );

      updated++;

    }

  });

  return {
    success: true,
    updated: updated
  };

}

/* =========================================================
   APPLICATION ID
========================================================= */

function generateApplicationId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'AP-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Application ID';
  });

  if (idColumn === -1) {
    throw new Error('Application ID column not found.');
  }

  const ids = sheet
    .getRange(
      2,
      idColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^AP-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'AP-' +
    String(maxNumber + 1).padStart(4, '0');
}


/* =========================================================
   ONE-TIME: ADD APPLICATION ID COLUMN
========================================================= */

function addApplicationIdColumn() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Applications');

  if (!sheet) {
    throw new Error(
      'Applications sheet not found.'
    );
  }

  const lastColumn =
    sheet.getLastColumn();

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const existing =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });

  if (existing !== -1) {

    return {
      success: true,
      message: 'Application ID column already exists.'
    };

  }

  sheet.insertColumnBefore(1);

  sheet
    .getRange(1, 1)
    .setValue('Application ID');

  return {
    success: true,
    message: 'Application ID column added.'
  };

}


/* =========================================================
   ONE-TIME: FILL EXISTING APPLICATION IDS
========================================================= */

function fillMissingApplicationIds() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Applications');

  if (!sheet) {
    throw new Error(
      'Applications sheet not found.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {

    return {
      success: true,
      updated: 0
    };

  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application ID';

    });

  const applicationColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Application';

    });

  if (idColumn === -1) {

    throw new Error(
      'Application ID column not found.'
    );

  }

  if (applicationColumn === -1) {

    throw new Error(
      'Application column not found.'
    );

  }

  const ids =
    sheet
      .getRange(
        2,
        idColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  const applications =
    sheet
      .getRange(
        2,
        applicationColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^AP-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  let updated = 0;

  ids.forEach(function(row, index) {

    const currentId =
      String(row[0] || '').trim();

    const applicationName =
      String(
        applications[index][0] || ''
      ).trim();

    if (
      !currentId &&
      applicationName
    ) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'AP-' +
          String(maxNumber).padStart(4, '0')
        );

      updated++;

    }

  });

  return {
    success: true,
    updated: updated
  };

}

/* =========================================================
   COMPETITORS
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE COMPETITOR ID
========================================================= */

function generateCompetitorId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'CO-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Competitor ID';
  });

  if (idColumn === -1) {
    throw new Error('Competitor ID column not found.');
  }

  const ids = sheet
    .getRange(
      2,
      idColumn + 1,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^CO-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'CO-' +
    String(maxNumber + 1).padStart(4, '0');

}


/* =========================================================
   ADD COMPETITOR
========================================================= */

function addCompetitor(competitor) {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Competitors');

  if (!sheet) {
    throw new Error(
      'Competitors sheet not found.'
    );
  }

  const lastColumn =
    sheet.getLastColumn();

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  const competitorId =
    generateCompetitorId(sheet);

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Competitor ID':
          return competitorId;

        case 'Competitor':
          return competitor.competitor || '';

        case 'Company Type':
          return competitor.companyType || '';

        case 'Main Products':
          return competitor.mainProducts || '';

        case 'Industries Served':
          return competitor.industriesServed || '';

        case 'Applications':
          return competitor.applications || '';

        case 'Strengths':
          return competitor.strengths || '';

        case 'Weaknesses':
          return competitor.weaknesses || '';

        case 'Price Position':
          return competitor.pricePosition || '';

        case 'Market Position':
          return competitor.marketPosition || '';

        case 'Customer Type':
          return competitor.customerType || '';

        case 'Typical Buyers':
          return competitor.typicalBuyers || '';

        case 'Key Differentiators':
          return competitor.keyDifferentiators || '';

        case 'Polycab Advantage':
          return competitor.polycabAdvantage || '';

        case 'Common Objections':
          return competitor.commonObjections || '';

        case 'How to Respond':
          return competitor.howToRespond || '';

        case 'Opportunity Strategy':
          return competitor.opportunityStrategy || '';

        case 'Related Products':
          return competitor.relatedProducts || '';

        case 'Status':
          return competitor.status || 'Active';

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
    competitorId: competitorId,
    competitor: competitor.competitor || ''
  };

}


/* =========================================================
   FIND COMPETITOR ROW
   Accepts ID OR NAME
========================================================= */

function findCompetitorRow(
  sheet,
  competitorKey
) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {
    return -1;
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const idColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Competitor ID';

    });

  const nameColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Competitor';

    });

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      idColumn !== -1 &&
      String(
        values[i][idColumn] || ''
      ).trim() ===
      String(competitorKey).trim()
    ) {

      return i + 2;

    }

    if (
      nameColumn !== -1 &&
      String(
        values[i][nameColumn] || ''
      ).trim() ===
      String(competitorKey).trim()
    ) {

      return i + 2;

    }

  }

  return -1;

}


/* =========================================================
   UPDATE COMPETITOR
========================================================= */

function updateCompetitor(
  competitorId,
  competitor
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Competitors');

  if (!sheet) {
    throw new Error(
      'Competitors sheet not found.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error(
      'No competitors found.'
    );
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const targetRow =
    findCompetitorRow(
      sheet,
      competitorId
    );

  if (targetRow === -1) {
    throw new Error(
      'Competitor not found: ' +
      competitorId
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

  const row =
    headers.map(function(
      header,
      index
    ) {

      switch (
        String(header).trim()
      ) {

        case 'Competitor ID':
          return existingRow[index];

        case 'Competitor':
          return competitor.competitor || '';

        case 'Company Type':
          return competitor.companyType || '';

        case 'Main Products':
          return competitor.mainProducts || '';

        case 'Industries Served':
          return competitor.industriesServed || '';

        case 'Applications':
          return competitor.applications || '';

        case 'Strengths':
          return competitor.strengths || '';

        case 'Weaknesses':
          return competitor.weaknesses || '';

        case 'Price Position':
          return competitor.pricePosition || '';

        case 'Market Position':
          return competitor.marketPosition || '';

        case 'Customer Type':
          return competitor.customerType || '';

        case 'Typical Buyers':
          return competitor.typicalBuyers || '';

        case 'Key Differentiators':
          return competitor.keyDifferentiators || '';

        case 'Polycab Advantage':
          return competitor.polycabAdvantage || '';

        case 'Common Objections':
          return competitor.commonObjections || '';

        case 'How to Respond':
          return competitor.howToRespond || '';

        case 'Opportunity Strategy':
          return competitor.opportunityStrategy || '';

        case 'Related Products':
          return competitor.relatedProducts || '';

        case 'Status':
          return competitor.status || 'Active';

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
    competitorId: competitorId,
    competitor: competitor.competitor || ''
  };

}


/* =========================================================
   DELETE COMPETITOR
========================================================= */

function deleteCompetitor(
  competitorId
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Competitors');

  if (!sheet) {
    throw new Error(
      'Competitors sheet not found.'
    );
  }

  const targetRow =
    findCompetitorRow(
      sheet,
      competitorId
    );

  if (targetRow === -1) {
    throw new Error(
      'Competitor not found: ' +
      competitorId
    );
  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    competitorId: competitorId
  };

}