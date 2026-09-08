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
   PRELOAD SECTION DATA
   Returns all navigation data in one web-app request so the
   client can render sections without a separate round trip.
========================================================= */

function getAllSectionData() {

  const result = {};

  SHEET_NAMES.forEach(function(sheetName) {
    result[sheetName] = getSheetData(sheetName);
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

/* =========================================================
   CABLE KNOWLEDGE
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE KNOWLEDGE ID
========================================================= */

function generateKnowledgeId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'CK-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Knowledge ID';
  });

  if (idColumn === -1) {
    throw new Error('Knowledge ID column not found.');
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
      value.match(/^CK-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'CK-' +
    String(maxNumber + 1).padStart(4, '0');
}


/* =========================================================
   ONE-TIME: ADD KNOWLEDGE ID COLUMN
========================================================= */

function addKnowledgeIdColumn() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Cable Knowledge');

  if (!sheet) {
    throw new Error(
      'Cable Knowledge sheet not found.'
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
        'Knowledge ID';

    });

  if (existing !== -1) {

    return {
      success: true,
      message: 'Knowledge ID column already exists.'
    };

  }

  sheet.insertColumnBefore(1);

  sheet
    .getRange(1, 1)
    .setValue('Knowledge ID');

  return {
    success: true,
    message: 'Knowledge ID column added.'
  };

}


/* =========================================================
   ONE-TIME: FILL EXISTING KNOWLEDGE IDS
========================================================= */

function fillMissingKnowledgeIds() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Cable Knowledge');

  if (!sheet) {
    throw new Error(
      'Cable Knowledge sheet not found.'
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
        'Knowledge ID';

    });

  if (idColumn === -1) {

    throw new Error(
      'Knowledge ID column not found.'
    );

  }

  /* Find Topic or Question column */

  let referenceColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Topic';

    });

  if (referenceColumn === -1) {

    referenceColumn =
      headers.findIndex(function(header) {

        return String(header).trim() ===
          'Question';

      });

  }

  if (referenceColumn === -1) {

    throw new Error(
      'Topic or Question column not found.'
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

  const references =
    sheet
      .getRange(
        2,
        referenceColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^CK-(\d+)$/i);

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

    const reference =
      String(
        references[index][0] || ''
      ).trim();

    if (
      !currentId &&
      reference
    ) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'CK-' +
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
   ADD KNOWLEDGE
========================================================= */

function addKnowledge(knowledge) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Cable Knowledge');

  if (!sheet) {
    throw new Error(
      'Cable Knowledge sheet not found.'
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

  const knowledgeId =
    generateKnowledgeId(sheet);

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Knowledge ID':
          return knowledgeId;

        case 'Topic':
          return knowledge.topic || '';

        case 'Category':
          return knowledge.category || '';

        case 'Sub Category':
          return knowledge.subCategory || '';

        case 'Question':
          return knowledge.question || '';

        case 'Answer':
          return knowledge.answer || '';

        case 'Key Points':
          return knowledge.keyPoints || '';

        case 'Sales Relevance':
          return knowledge.salesRelevance || '';

        case 'Learning Level':
          return knowledge.learningLevel || '';

        case 'Related Products':
          return knowledge.relatedProducts || '';

        case 'Related Applications':
          return knowledge.relatedApplications || '';

        case 'Related Industries':
          return knowledge.relatedIndustries || '';

        case 'Document Link':
          return knowledge.documentLink || '';

        case 'Source':
          return knowledge.source || '';

        case 'Status':
          return knowledge.status || 'Active';

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
    knowledgeId: knowledgeId
  };

}


/* =========================================================
   FIND KNOWLEDGE ROW
========================================================= */

function findKnowledgeRow(
  sheet,
  knowledgeId
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
        'Knowledge ID';

    });

  if (idColumn === -1) {
    return -1;
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

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0] || '').trim() ===
      String(knowledgeId).trim()
    ) {

      return i + 2;

    }

  }

  return -1;

}


/* =========================================================
   UPDATE KNOWLEDGE
========================================================= */

function updateKnowledge(
  knowledgeId,
  knowledge
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Cable Knowledge');

  if (!sheet) {
    throw new Error(
      'Cable Knowledge sheet not found.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  const targetRow =
    findKnowledgeRow(
      sheet,
      knowledgeId
    );

  if (targetRow === -1) {

    throw new Error(
      'Knowledge not found: ' +
      knowledgeId
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

        case 'Knowledge ID':
          return existingRow[index];

        case 'Topic':
          return knowledge.topic || '';

        case 'Category':
          return knowledge.category || '';

        case 'Sub Category':
          return knowledge.subCategory || '';

        case 'Question':
          return knowledge.question || '';

        case 'Answer':
          return knowledge.answer || '';

        case 'Key Points':
          return knowledge.keyPoints || '';

        case 'Sales Relevance':
          return knowledge.salesRelevance || '';

        case 'Learning Level':
          return knowledge.learningLevel || '';

        case 'Related Products':
          return knowledge.relatedProducts || '';

        case 'Related Applications':
          return knowledge.relatedApplications || '';

        case 'Related Industries':
          return knowledge.relatedIndustries || '';

        case 'Document Link':
          return knowledge.documentLink || '';

        case 'Source':
          return knowledge.source || '';

        case 'Status':
          return knowledge.status || 'Active';

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
    knowledgeId: knowledgeId
  };

}


/* =========================================================
   DELETE KNOWLEDGE
========================================================= */

function deleteKnowledge(
  knowledgeId
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Cable Knowledge');

  if (!sheet) {
    throw new Error(
      'Cable Knowledge sheet not found.'
    );
  }

  const targetRow =
    findKnowledgeRow(
      sheet,
      knowledgeId
    );

  if (targetRow === -1) {

    throw new Error(
      'Knowledge not found: ' +
      knowledgeId
    );

  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    knowledgeId: knowledgeId
  };

}

/* =========================================================
   LEARNING
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE LEARNING ID
========================================================= */

function generateLearningId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'LR-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {

    return String(header).trim() === 'Learning ID';

  });

  if (idColumn === -1) {

    throw new Error(
      'Learning ID column not found.'
    );

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
      value.match(/^LR-(\d+)$/i);

    if (match) {

      const number =
        parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'LR-' +
    String(maxNumber + 1).padStart(4, '0');

}


/* =========================================================
   ADD LEARNING ID COLUMN
   RUN ONCE ONLY IF NEEDED
========================================================= */

function addLearningIdColumn() {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Learning');

  if (!sheet) {

    throw new Error(
      'Learning sheet not found.'
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

  const exists =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Learning ID';

    });

  if (exists !== -1) {

    return {
      success: true,
      message:
        'Learning ID column already exists.'
    };

  }

  sheet.insertColumnBefore(1);

  sheet
    .getRange(1, 1)
    .setValue('Learning ID');

  return {
    success: true,
    message:
      'Learning ID column added.'
  };

}


/* =========================================================
   FILL EXISTING LEARNING IDS
========================================================= */

function fillMissingLearningIds() {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Learning');

  if (!sheet) {

    throw new Error(
      'Learning sheet not found.'
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
        'Learning ID';

    });

  if (idColumn === -1) {

    throw new Error(
      'Learning ID column not found.'
    );

  }

  const topicColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Topic';

    });

  if (topicColumn === -1) {

    throw new Error(
      'Topic column not found.'
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

  const topics =
    sheet
      .getRange(
        2,
        topicColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^LR-(\d+)$/i);

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

    const topic =
      String(
        topics[index][0] || ''
      ).trim();

    if (
      !currentId &&
      topic
    ) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'LR-' +
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
   ADD LEARNING
========================================================= */

function addLearning(learning) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Learning');

  if (!sheet) {

    throw new Error(
      'Learning sheet not found.'
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

  const learningId =
    generateLearningId(sheet);

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Learning ID':
          return learningId;

        case 'Topic':
          return learning.topic || '';

        case 'Category':
          return learning.category || '';

        case 'Sub Category':
          return learning.subCategory || '';

        case 'What to Learn':
          return learning.whatToLearn || '';

        case 'Why It Matters':
          return learning.whyItMatters || '';

        case 'Key Concepts':
          return learning.keyConcepts || '';

        case 'Practical Example':
          return learning.practicalExample || '';

        case 'Sales Application':
          return learning.salesApplication || '';

        case 'Questions to Practice':
          return learning.questionsToPractice || '';

        case 'Related Products':
          return learning.relatedProducts || '';

        case 'Related Industries':
          return learning.relatedIndustries || '';

        case 'Learning Level':
          return learning.learningLevel || '';

        case 'Status':
          return learning.status || 'Active';

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
    learningId: learningId
  };

}


/* =========================================================
   FIND LEARNING ROW
========================================================= */

function findLearningRow(
  sheet,
  learningId
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
        'Learning ID';

    });

  if (idColumn === -1) {
    return -1;
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

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0] || '').trim() ===
      String(learningId).trim()
    ) {

      return i + 2;

    }

  }

  return -1;

}


/* =========================================================
   UPDATE LEARNING
========================================================= */

function updateLearning(
  learningId,
  learning
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Learning');

  if (!sheet) {

    throw new Error(
      'Learning sheet not found.'
    );

  }

  const lastColumn =
    sheet.getLastColumn();

  const targetRow =
    findLearningRow(
      sheet,
      learningId
    );

  if (targetRow === -1) {

    throw new Error(
      'Learning record not found: ' +
      learningId
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

        case 'Learning ID':
          return existingRow[index];

        case 'Topic':
          return learning.topic || '';

        case 'Category':
          return learning.category || '';

        case 'Sub Category':
          return learning.subCategory || '';

        case 'What to Learn':
          return learning.whatToLearn || '';

        case 'Why It Matters':
          return learning.whyItMatters || '';

        case 'Key Concepts':
          return learning.keyConcepts || '';

        case 'Practical Example':
          return learning.practicalExample || '';

        case 'Sales Application':
          return learning.salesApplication || '';

        case 'Questions to Practice':
          return learning.questionsToPractice || '';

        case 'Related Products':
          return learning.relatedProducts || '';

        case 'Related Industries':
          return learning.relatedIndustries || '';

        case 'Learning Level':
          return learning.learningLevel || '';

        case 'Status':
          return learning.status || 'Active';

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
    learningId: learningId
  };

}


/* =========================================================
   DELETE LEARNING
========================================================= */

function deleteLearning(
  learningId
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Learning');

  if (!sheet) {

    throw new Error(
      'Learning sheet not found.'
    );

  }

  const targetRow =
    findLearningRow(
      sheet,
      learningId
    );

  if (targetRow === -1) {

    throw new Error(
      'Learning record not found: ' +
      learningId
    );

  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    learningId: learningId
  };

}

/* =========================================================
   CAREER GROWTH
   ADD / UPDATE / DELETE
========================================================= */


/* =========================================================
   GENERATE CAREER ID
========================================================= */

function generateCareerId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'CG-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Career ID';
  });

  if (idColumn === -1) {
    throw new Error('Career ID column not found.');
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

    const value = String(row[0] || '').trim();

    const match = value.match(/^CG-(\d+)$/i);

    if (match) {

      const number = parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'CG-' +
    String(maxNumber + 1).padStart(4, '0');
}


/* =========================================================
   ADD CAREER ID COLUMN
   RUN ONLY IF COLUMN DOES NOT EXIST
========================================================= */

function addCareerIdColumn() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Career Growth');

  if (!sheet) {
    throw new Error(
      'Career Growth sheet not found.'
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

  const exists =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Career ID';

    });

  if (exists !== -1) {

    return {
      success: true,
      message:
        'Career ID column already exists.'
    };

  }

  sheet.insertColumnBefore(1);

  sheet
    .getRange(1, 1)
    .setValue('Career ID');

  return {
    success: true,
    message:
      'Career ID column added.'
  };

}


/* =========================================================
   FILL EXISTING CAREER IDS
========================================================= */

function fillMissingCareerIds() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Career Growth');

  if (!sheet) {
    throw new Error(
      'Career Growth sheet not found.'
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
        'Career ID';

    });

  const topicColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Topic';

    });

  if (idColumn === -1) {
    throw new Error(
      'Career ID column not found.'
    );
  }

  if (topicColumn === -1) {
    throw new Error(
      'Topic column not found.'
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

  const topics =
    sheet
      .getRange(
        2,
        topicColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let maxNumber = 0;

  ids.forEach(function(row) {

    const value =
      String(row[0] || '').trim();

    const match =
      value.match(/^CG-(\d+)$/i);

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

    const topic =
      String(
        topics[index][0] || ''
      ).trim();

    if (!currentId && topic) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'CG-' +
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
   ADD CAREER
========================================================= */

function addCareerGrowth(career) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Career Growth');

  if (!sheet) {
    throw new Error(
      'Career Growth sheet not found.'
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

  const careerId =
    generateCareerId(sheet);

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Career ID':
          return careerId;

        case 'Topic':
          return career.topic || '';

        case 'Category':
          return career.category || '';

        case 'Sub Category':
          return career.subCategory || '';

        case 'Goal':
          return career.goal || '';

        case 'Why It Matters':
          return career.whyItMatters || '';

        case 'Skills to Develop':
          return career.skillsToDevelop || '';

        case 'Actions':
          return career.actions || '';

        case 'Measure of Progress':
          return career.measureOfProgress || '';

        case 'Sales Application':
          return career.salesApplication || '';

        case 'Resources':
          return career.resources || '';

        case 'Learning Level':
          return career.learningLevel || '';

        case 'Status':
          return career.status || 'Active';

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
    careerId: careerId
  };

}


/* =========================================================
   FIND CAREER ROW
========================================================= */

function findCareerRow(
  sheet,
  careerId
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
        'Career ID';

    });

  if (idColumn === -1) {
    return -1;
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

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0] || '').trim() ===
      String(careerId).trim()
    ) {

      return i + 2;

    }

  }

  return -1;

}


/* =========================================================
   UPDATE CAREER
========================================================= */

function updateCareerGrowth(
  careerId,
  career
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Career Growth');

  if (!sheet) {
    throw new Error(
      'Career Growth sheet not found.'
    );
  }

  const lastColumn =
    sheet.getLastColumn();

  const targetRow =
    findCareerRow(
      sheet,
      careerId
    );

  if (targetRow === -1) {

    throw new Error(
      'Career record not found: ' +
      careerId
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

        case 'Career ID':
          return existingRow[index];

        case 'Topic':
          return career.topic || '';

        case 'Category':
          return career.category || '';

        case 'Sub Category':
          return career.subCategory || '';

        case 'Goal':
          return career.goal || '';

        case 'Why It Matters':
          return career.whyItMatters || '';

        case 'Skills to Develop':
          return career.skillsToDevelop || '';

        case 'Actions':
          return career.actions || '';

        case 'Measure of Progress':
          return career.measureOfProgress || '';

        case 'Sales Application':
          return career.salesApplication || '';

        case 'Resources':
          return career.resources || '';

        case 'Learning Level':
          return career.learningLevel || '';

        case 'Status':
          return career.status || 'Active';

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
    careerId: careerId
  };

}


/* =========================================================
   DELETE CAREER
========================================================= */

function deleteCareerGrowth(
  careerId
) {

  const ss =
    getSpreadsheet();

  const sheet =
    ss.getSheetByName('Career Growth');

  if (!sheet) {
    throw new Error(
      'Career Growth sheet not found.'
    );
  }

  const targetRow =
    findCareerRow(
      sheet,
      careerId
    );

  if (targetRow === -1) {

    throw new Error(
      'Career record not found: ' +
      careerId
    );

  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    careerId: careerId
  };

}

/* =========================================================
   DOCUMENTS
   ADD / UPDATE / DELETE
========================================================= */

function generateDocumentId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'DOC-0001';
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const idColumn = headers.findIndex(function(header) {
    return String(header).trim() === 'Document ID';
  });

  if (idColumn === -1) {
    throw new Error('Document ID column not found.');
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

    const value = String(row[0] || '').trim();

    const match = value.match(/^DOC-(\d+)$/i);

    if (match) {

      const number = parseInt(match[1], 10);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });

  return 'DOC-' +
    String(maxNumber + 1).padStart(4, '0');

}


/* =========================================================
   ADD DOCUMENT ID COLUMN
   RUN ONLY IF IT DOES NOT EXIST
========================================================= */

function addDocumentIdColumn() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
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

  const exists =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Document ID';

    });

  if (exists !== -1) {

    return {
      success: true,
      message: 'Document ID column already exists.'
    };

  }

  sheet.insertColumnBefore(1);

  sheet
    .getRange(1, 1)
    .setValue('Document ID');

  return {
    success: true,
    message: 'Document ID column added.'
  };

}


/* =========================================================
   FILL MISSING DOCUMENT IDS
========================================================= */

function fillMissingDocumentIds() {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
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
        'Document ID';

    });

  const nameColumn =
    headers.findIndex(function(header) {

      return String(header).trim() ===
        'Document Name';

    });

  if (idColumn === -1) {
    throw new Error('Document ID column not found.');
  }

  if (nameColumn === -1) {
    throw new Error('Document Name column not found.');
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

  const names =
    sheet
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
      value.match(/^DOC-(\d+)$/i);

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

    const name =
      String(
        names[index][0] || ''
      ).trim();

    if (!currentId && name) {

      maxNumber++;

      sheet
        .getRange(
          index + 2,
          idColumn + 1
        )
        .setValue(
          'DOC-' +
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
   ADD DOCUMENT
========================================================= */

function addDocument(documentData) {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
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

  const documentId =
    generateDocumentId(sheet);

  const row =
    headers.map(function(header) {

      switch (
        String(header).trim()
      ) {

        case 'Document ID':
          return documentId;

        case 'Document Name':
          return documentData.documentName || '';

        case 'Document Type':
          return documentData.documentType || '';

        case 'Category':
          return documentData.category || '';

        case 'Related Products':
          return documentData.relatedProducts || '';

        case 'Related Industries':
          return documentData.relatedIndustries || '';

        case 'Description':
          return documentData.description || '';

        case 'Document Link':
          return documentData.documentLink || '';

        case 'Source':
          return documentData.source || '';

        case 'Learning Level':
          return documentData.learningLevel || '';

        case 'Status':
          return documentData.status || 'Active';

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
    documentId: documentId
  };

}


/* =========================================================
   FIND DOCUMENT ROW
========================================================= */

function findDocumentRow(
  sheet,
  documentId
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
        'Document ID';

    });

  if (idColumn === -1) {
    return -1;
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

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0] || '').trim() ===
      String(documentId).trim()
    ) {

      return i + 2;

    }

  }

  return -1;

}


/* =========================================================
   UPDATE DOCUMENT
========================================================= */

function updateDocument(
  documentId,
  documentData
) {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
  }

  const lastColumn =
    sheet.getLastColumn();

  const targetRow =
    findDocumentRow(
      sheet,
      documentId
    );

  if (targetRow === -1) {

    throw new Error(
      'Document not found: ' +
      documentId
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

        case 'Document ID':
          return existingRow[index];

        case 'Document Name':
          return documentData.documentName || '';

        case 'Document Type':
          return documentData.documentType || '';

        case 'Category':
          return documentData.category || '';

        case 'Related Products':
          return documentData.relatedProducts || '';

        case 'Related Industries':
          return documentData.relatedIndustries || '';

        case 'Description':
          return documentData.description || '';

        case 'Document Link':
          return documentData.documentLink || '';

        case 'Source':
          return documentData.source || '';

        case 'Learning Level':
          return documentData.learningLevel || '';

        case 'Status':
          return documentData.status || 'Active';

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
    documentId: documentId
  };

}


/* =========================================================
   DELETE DOCUMENT
========================================================= */

function deleteDocument(documentId) {

  const ss = getSpreadsheet();

  const sheet =
    ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
  }

  const targetRow =
    findDocumentRow(
      sheet,
      documentId
    );

  if (targetRow === -1) {

    throw new Error(
      'Document not found: ' +
      documentId
    );

  }

  sheet.deleteRow(targetRow);

  return {
    success: true,
    documentId: documentId
  };

}
