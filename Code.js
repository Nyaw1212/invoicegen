/**
 * Elliott invoice generator.
 * Bound to the Elliott Invoice Generator spreadsheet using clasp.
 */
const INVOICE_CONFIG = Object.freeze({
  TEMPLATE_ID: '1QM-GM6cv0bZ7yhHDPBaE2wEOtlStZmb2uu4yXTqsgPU',
  SHEET_NAME: 'Invoice',
  FIRST_ITEM_ROW: 9,
  DUE_DAYS: 8,
  GENERATED_LINK_CELL: 'K2'
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Elliott Invoice')
    .addItem('Generate Invoice', 'generateInvoice')
    .addToUi();
}

function generateInvoice() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = createInvoice_();
    ui.alert('Invoice created!\n' + result.url);
  } catch (err) {
    ui.alert('Could not create invoice:\n' + err.message);
    throw err;
  }
}

function createInvoice_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVOICE_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Invoice tab is missing.');

  const id = sheet.getRange('B1').getDisplayValue().trim();
  const dateValue = sheet.getRange('B2').getValue();
  const job = sheet.getRange('B3').getDisplayValue().trim();
  const shipTo = sheet.getRange('B4').getDisplayValue().trim();
  if (!id || !job || !shipTo || !dateValue) {
    throw new Error('Complete B1:B4 before generating.');
  }

  const orderDate = invoiceDate_(dateValue);
  const dueDate = new Date(orderDate);
  dueDate.setDate(dueDate.getDate() + INVOICE_CONFIG.DUE_DAYS);
  const spreadsheetTimeZone = ss.getSpreadsheetTimeZone();
  const tz = (typeof spreadsheetTimeZone === 'string' && spreadsheetTimeZone.trim())
    ? spreadsheetTimeZone
    : Session.getScriptTimeZone() || 'America/Chicago';
  const dateText = value => Utilities.formatDate(value, tz, 'MM/dd/yyyy');

  const count = sheet.getLastRow() - INVOICE_CONFIG.FIRST_ITEM_ROW + 1;
  if (count <= 0) throw new Error('Add an item in row 9.');

  const sourceRows = sheet.getRange(INVOICE_CONFIG.FIRST_ITEM_ROW, 1, count, 8).getValues();
  const items = sourceRows.map((r, offset) => ({row: r, sheetRow: offset + INVOICE_CONFIG.FIRST_ITEM_ROW}))
    .filter(({row}) => row.slice(0, 8).some(v => v !== '' && v !== null))
    .map(({row, sheetRow}) => {
      const numeric = [0, 1, 2, 6].map(i => {
        if (row[i] === '' || row[i] === null) throw new Error('Missing quantity or price in row ' + sheetRow);
        const n = Number(row[i]);
        if (!Number.isFinite(n) || n < 0) throw new Error('Invalid quantity or price in row ' + sheetRow);
        return n;
      });
      return {
        ordered: numeric[0],
        shipped: numeric[1],
        backorder: numeric[2],
        catalog: String(row[3] ?? ''),
        vendor: String(row[4] ?? ''),
        description: String(row[5] ?? ''),
        price: numeric[3],
        unit: String(row[7] ?? ''),
        extended: Math.round(numeric[1] * numeric[3] * 100) / 100
      };
    });
  if (!items.length) throw new Error('Add at least one item.');
  const total = Math.round(items.reduce((s, i) => s + i.extended, 0) * 100) / 100;

  // Create and edit a copy; the template is never modified.
  const name = 'Elliott Invoice - ' + id + ' - ' + job;
  const copy = DriveApp.getFileById(INVOICE_CONFIG.TEMPLATE_ID).makeCopy(name);
  try {
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    const material = findMaterialRow_(body);
    if (!material) throw new Error('Template material row ({{CATALOG_NUMBER}}) was not found.');

    // Clone the row BEFORE filling placeholders in the original row.
    const clones = items.slice(1).map(() => material.row.copy());
    fillMaterialRow_(material.row, items[0]);
    clones.forEach((clone, i) => {
      const inserted = material.table.insertTableRow(material.rowIndex + i + 1, clone);
      fillMaterialRow_(inserted, items[i + 1]);
    });

    const fields = {
      INVOICE_ID: id,
      INVOICE_DATE: dateText(orderDate),
      DATE_DUE: dateText(dueDate),
      CUSTOMER_JOB: job,
      SHIP_TO: shipTo,
      SUBTOTAL: money_(total),
      TOTAL: money_(total)
    };
    Object.entries(fields).forEach(([key, value]) => {
      body.replaceText('\\{\\{' + key + '\\}\\}', String(value));
    });
    doc.saveAndClose();
  } catch (err) {
    copy.setTrashed(true);
    throw err;
  }

  const url = copy.getUrl();
  sheet.getRange('K1').setValue('Generated Invoice');
  sheet.getRange(INVOICE_CONFIG.GENERATED_LINK_CELL).setFormula(
    '=HYPERLINK("' + url.replace(/"/g, '""') + '","Open invoice")'
  );
  return {id, url, itemCount: items.length, total};
}

function findMaterialRow_(body) {
  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getText().includes('{{CATALOG_NUMBER}}')) {
        return {table, row, rowIndex: r};
      }
    }
  }
  return null;
}

function fillMaterialRow_(row, item) {
  const fields = {
    ORDER_QTY: item.ordered,
    SHIP_QTY: item.shipped,
    BACKORDER_QTY: item.backorder,
    CATALOG_NUMBER: item.catalog,
    VENDOR_CODE: item.vendor,
    ITEM_DESCRIPTION: item.description,
    UNIT_PRICE: '
    UNIT_CODE: item.unit,
    EXTENDED_PRICE: '
  };
  Object.entries(fields).forEach(([key, value]) => {
    const pattern = '\\{\\{' + key + '\\}\\}';
    for (let i = 0; i < row.getNumCells(); i++) {
      row.getCell(i).replaceText(pattern, String(value));
    }
  });
}

function elliottPrice_(n) {
  // Match the source Elliott invoice's dollar-sign spacing.
  return '$       ' + money_(n);
}

function money_(n) {
  return Number(n).toFixed(2);
}

function invoiceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value);
  const m = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error('B2 must contain a date in MM/DD/YYYY format.');
  const month = Number(m[1]), day = Number(m[2]), year = Number(m[3]);
  const date = new Date(year, month - 1, day, 12);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('B2 contains an invalid date.');
  }
  return date;
}
 + money_(item.price),
    UNIT_CODE: item.unit,
    EXTENDED_PRICE: money_(item.extended)
  };
  Object.entries(fields).forEach(([key, value]) => {
    const pattern = '\\{\\{' + key + '\\}\\}';
    for (let i = 0; i < row.getNumCells(); i++) {
      row.getCell(i).replaceText(pattern, String(value));
    }
  });
}

function money_(n) {
  return Number(n).toFixed(2);
}

function invoiceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value);
  const m = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error('B2 must contain a date in MM/DD/YYYY format.');
  const month = Number(m[1]), day = Number(m[2]), year = Number(m[3]);
  const date = new Date(year, month - 1, day, 12);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('B2 contains an invalid date.');
  }
  return date;
}
 + money_(item.extended)
  };
  Object.entries(fields).forEach(([key, value]) => {
    const pattern = '\\{\\{' + key + '\\}\\}';
    for (let i = 0; i < row.getNumCells(); i++) {
      row.getCell(i).replaceText(pattern, String(value));
    }
  });
}

function money_(n) {
  return Number(n).toFixed(2);
}

function invoiceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value);
  const m = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error('B2 must contain a date in MM/DD/YYYY format.');
  const month = Number(m[1]), day = Number(m[2]), year = Number(m[3]);
  const date = new Date(year, month - 1, day, 12);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('B2 contains an invalid date.');
  }
  return date;
}
 + money_(item.price),
    UNIT_CODE: item.unit,
    EXTENDED_PRICE: money_(item.extended)
  };
  Object.entries(fields).forEach(([key, value]) => {
    const pattern = '\\{\\{' + key + '\\}\\}';
    for (let i = 0; i < row.getNumCells(); i++) {
      row.getCell(i).replaceText(pattern, String(value));
    }
  });
}

function money_(n) {
  return Number(n).toFixed(2);
}

function invoiceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value);
  const m = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error('B2 must contain a date in MM/DD/YYYY format.');
  const month = Number(m[1]), day = Number(m[2]), year = Number(m[3]);
  const date = new Date(year, month - 1, day, 12);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error('B2 contains an invalid date.');
  }
  return date;
}
