var SHEETS = {
  ITEMS: 'Items List',
  ORDERS: 'Order List',
  ORDER_ITEMS: 'Order Items'
};

var ORDER_STATUSES = ['PENDING', 'APPROVED', 'REJECTED', 'CANCELLED'];
var PAYMENT_STATUSES = ['PENDING', 'PAID', 'UNPAID', 'REFUNDED'];

function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_');
  });

  return values.slice(1)
    .filter(function (row) {
      return row.some(function (v) { return v !== '' && v !== null && v !== undefined; });
    })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        obj[h] = row[i];
      });
      return obj;
    });
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'list').toLowerCase();

  if (action === 'health') {
    return jsonResponse({
      success: true,
      service: 'inventory-api',
      sheets: SHEETS
    });
  }

  if (action !== 'list' && action !== 'get') {
    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  }

  const ss = getSpreadsheet();
  if (!ss) {
    return jsonResponse({ success: false, error: 'Spreadsheet unavailable' });
  }

  const key = String(p.sheet || p.type || 'items')
    .toLowerCase()
    .replace(/[^a-z]/g, '');

  const sheetMap = {
    item: SHEETS.ITEMS,
    items: SHEETS.ITEMS,
    order: SHEETS.ORDERS,
    orders: SHEETS.ORDERS,
    orderitem: SHEETS.ORDER_ITEMS,
    orderitems: SHEETS.ORDER_ITEMS
  };

  const name = sheetMap[key] || SHEETS.ITEMS;
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    return jsonResponse({ success: false, error: 'Sheet not found: ' + name });
  }

  let data = sheetData(sheet);

  if (action === 'get' && p.id) {
    data = data.filter(function (row) {
      return Number(row.id) === Number(p.id);
    });
  }

  return jsonResponse({
    success: true,
    action: action,
    sheet: name,
    count: data.length,
    data: data
  });
}

function doPost(e) {
  const p = Object.assign({}, (e && e.parameter) || {});
  let body = {};

  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonResponse({ success: false, error: 'Invalid JSON body' });
    }
  }

  const input = Object.assign({}, p, body);
  const action = String(input.action || '').toLowerCase();

  try {
    switch (action) {
      case 'add-item':
        return jsonResponse({
          success: true,
          id: addItemRecord(input.name, input.price, input.stock, input.image)
        });

      case 'add-order':
        return jsonResponse({ success: true, id: createOrderRecord() });

      case 'add-order-item':
        return jsonResponse({
          success: true,
          id: addOrderItemRecord(input.orderId, input.itemId, input.quantity, input.deductStock)
        });

      case 'update-order-status':
        setOrderStatusRecord(input.id, 3, input.status, ORDER_STATUSES);
        return jsonResponse({ success: true });

      case 'update-payment-status':
        setOrderStatusRecord(input.id, 4, input.status, PAYMENT_STATUSES);
        return jsonResponse({ success: true });

      case 'update-order-item-status': {
        const row = findRow(getSheet(SHEETS.ORDER_ITEMS), 1, input.id);
        if (!row) throw new Error('Order item not found');
        applyOrderItemStatus(row, input.status);
        return jsonResponse({ success: true });
      }

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory System')
    .addItem('Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('Add New Item', 'addItem')
    .addItem('Create Order', 'createOrder')
    .addItem('Add Item to Order', 'addOrderItem')
    .addSeparator()
    .addItem('Update Order Status', 'updateOrderStatus')
    .addItem('Update Payment Status', 'updatePaymentStatus')
    .addToUi();
}

// CREATE SHEETS AND HEADERS
function setupSheets() {
  const ss = getSpreadsheet();

  const headers = {
    [SHEETS.ITEMS]: [
      'id', 'name', 'price', 'stock', 'image url'
    ],
    [SHEETS.ORDERS]: [
      'id', 'total', 'status', 'payment status'
    ],
    [SHEETS.ORDER_ITEMS]: [
      'id', 'order id', 'item id', 'quantity',
      'price', 'total', 'status', 'stock deducted'
    ]
  };

  Object.keys(headers).forEach(name => {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers[name]);
    }
  });

  // Status dropdowns
  const orderItems = ss.getSheetByName(SHEETS.ORDER_ITEMS);
  const orders = ss.getSheetByName(SHEETS.ORDERS);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ORDER_STATUSES, true)
    .setAllowInvalid(false)
    .build();

  const paymentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PAYMENT_STATUSES, true)
    .setAllowInvalid(false)
    .build();

  orderItems.getRange('G2:G1000').setDataValidation(statusRule);
  orders.getRange('C2:C1000').setDataValidation(statusRule);
  orders.getRange('D2:D1000').setDataValidation(paymentRule);

  SpreadsheetApp.getUi().alert('Inventory system setup completed!');
}

// GENERATE NEXT ID
function nextId(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return 1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(Number)
    .filter(n => !isNaN(n));

  return ids.length ? Math.max(...ids) + 1 : 1;
}

// ADD NEW ITEM
function addItem() {
  const ui = SpreadsheetApp.getUi();

  const name = ui.prompt('Add Item', 'Enter item name:', ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() !== ui.Button.OK) return;

  const price = ui.prompt('Add Item', 'Enter item price:', ui.ButtonSet.OK_CANCEL);
  if (price.getSelectedButton() !== ui.Button.OK) return;

  const stock = ui.prompt('Add Item', 'Enter stock quantity:', ui.ButtonSet.OK_CANCEL);
  if (stock.getSelectedButton() !== ui.Button.OK) return;

  const image = ui.prompt('Add Item', 'Enter image URL (optional):', ui.ButtonSet.OK_CANCEL);
  if (image.getSelectedButton() !== ui.Button.OK) return;

  try {
    const id = addItemRecord(
      name.getResponseText(),
      price.getResponseText(),
      stock.getResponseText(),
      image.getResponseText()
    );
    ui.alert('Item added successfully! Item ID: ' + id);
  } catch (err) {
    ui.alert(err.message);
  }
}

function addItemRecord(name, price, stock, image) {
  const itemName = String(name || '').trim();
  const itemPrice = Number(price);
  const itemStock = Number(stock);

  if (!itemName || !Number.isFinite(itemPrice) ||
      itemPrice < 0 || !Number.isInteger(itemStock) || itemStock < 0) {
    throw new Error('Please enter a valid name, price and stock quantity.');
  }

  const sheet = getSheet(SHEETS.ITEMS);
  const id = nextId(sheet);

  sheet.appendRow([
    id,
    itemName,
    itemPrice,
    itemStock,
    String(image || '').trim()
  ]);

  return id;
}

// CREATE NEW ORDER
function createOrder() {
  const id = createOrderRecord();

  SpreadsheetApp.getUi().alert(
    'Order created successfully!\nOrder ID: ' + id
  );
}

function createOrderRecord() {
  const sheet = getSheet(SHEETS.ORDERS);
  const id = nextId(sheet);

  sheet.appendRow([
    id,
    0,
    'PENDING',
    'PENDING'
  ]);

  return id;
}

// ADD ITEM TO AN EXISTING ORDER
function addOrderItem() {
  const ui = SpreadsheetApp.getUi();

  const orderPrompt = ui.prompt(
    'Add Order Item',
    'Enter Order ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (orderPrompt.getSelectedButton() !== ui.Button.OK) return;

  const itemPrompt = ui.prompt(
    'Add Order Item',
    'Enter Item ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (itemPrompt.getSelectedButton() !== ui.Button.OK) return;

  const qtyPrompt = ui.prompt(
    'Add Order Item',
    'Enter quantity:',
    ui.ButtonSet.OK_CANCEL
  );
  if (qtyPrompt.getSelectedButton() !== ui.Button.OK) return;

  try {
    const id = addOrderItemRecord(
      orderPrompt.getResponseText(),
      itemPrompt.getResponseText(),
      qtyPrompt.getResponseText(),
      false
    );
    ui.alert('Item added to order successfully! Order item ID: ' + id);
  } catch (err) {
    ui.alert(err.message);
  }
}

function addOrderItemRecord(orderIdRaw, itemIdRaw, quantityRaw, deductStock) {
  const orderId = Number(orderIdRaw);
  const itemId = Number(itemIdRaw);
  const quantity = Number(quantityRaw);

  const ss = getSpreadsheet();
  const orders = ss.getSheetByName(SHEETS.ORDERS);
  const items = ss.getSheetByName(SHEETS.ITEMS);
  const orderItems = ss.getSheetByName(SHEETS.ORDER_ITEMS);

  const orderRow = findRow(orders, 1, orderId);
  if (!orderRow) throw new Error('Order not found!');

  const itemRow = findRow(items, 1, itemId);
  if (!itemRow) throw new Error('Item not found!');

  if (!Number.isInteger(quantity) || quantity <= 0) {
    throw new Error('Enter a valid quantity.');
  }

  const orderStatus = String(orders.getRange(orderRow, 3).getValue()).toUpperCase();

  if (orderStatus === 'APPROVED') {
    throw new Error('This order is already approved. Create a new order to add more items.');
  }

  const itemData = items.getRange(itemRow, 1, 1, 5).getValues()[0];
  const price = Number(itemData[2]);
  const total = price * quantity;
  const id = nextId(orderItems);

  orderItems.appendRow([
    id,
    orderId,
    itemId,
    quantity,
    price,
    total,
    'PENDING',
    false
  ]);

  recalculateOrderTotal(orderId);

  if (deductStock === true || String(deductStock).toLowerCase() === 'true') {
    applyOrderItemStatus(findRow(orderItems, 1, id), 'APPROVED');
  }

  return id;
}

// FIND ROW USING ID
function findRow(sheet, column, value) {
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(
    2, column, sheet.getLastRow() - 1, 1
  ).getValues();

  const index = values.findIndex(
    row => Number(row[0]) === Number(value)
  );

  return index === -1 ? null : index + 2;
}

// RECALCULATE ORDER TOTAL
function recalculateOrderTotal(orderId) {
  const ss = getSpreadsheet();
  const orders = ss.getSheetByName(SHEETS.ORDERS);
  const orderItems = ss.getSheetByName(SHEETS.ORDER_ITEMS);

  const orderRow = findRow(orders, 1, orderId);
  if (!orderRow) return;

  let total = 0;

  if (orderItems.getLastRow() >= 2) {
    const data = orderItems.getRange(
      2, 1, orderItems.getLastRow() - 1, 8
    ).getValues();

    data.forEach(row => {
      if (Number(row[1]) === Number(orderId)) {
        total += Number(row[5]) || 0;
      }
    });
  }

  orders.getRange(orderRow, 2).setValue(total);
}

function applyOrderItemStatus(row, statusRaw) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ORDER_ITEMS);
  const status = String(statusRaw || '').trim().toUpperCase();

  if (!ORDER_STATUSES.includes(status)) {
    throw new Error('Invalid status!');
  }

  const orderId = sheet.getRange(row, 2).getValue();
  const itemId = sheet.getRange(row, 3).getValue();
  const quantity = Number(sheet.getRange(row, 4).getValue());
  const deductedCell = sheet.getRange(row, 8);
  const alreadyDeducted = deductedCell.getValue() === true;

  const items = ss.getSheetByName(SHEETS.ITEMS);
  const itemRow = findRow(items, 1, itemId);

  if (!itemRow || !Number.isInteger(quantity) || quantity <= 0) {
    throw new Error('Invalid item or quantity.');
  }

  const stockCell = items.getRange(itemRow, 4);
  const currentStock = Number(stockCell.getValue());

  // APPROVE: REDUCE STOCK ONCE
  if (status === 'APPROVED' && !alreadyDeducted) {
    if (currentStock < quantity) {
      throw new Error('Insufficient stock! Available: ' + currentStock);
    }

    stockCell.setValue(currentStock - quantity);
    deductedCell.setValue(true);
  }

  // CANCEL / REJECT: RESTORE PREVIOUSLY DEDUCTED STOCK
  if (status !== 'APPROVED' && alreadyDeducted) {
    stockCell.setValue(currentStock + quantity);
    deductedCell.setValue(false);
  }

  sheet.getRange(row, 7).setValue(status);
  recalculateOrderTotal(orderId);
}

// AUTOMATIC STOCK MANAGEMENT
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== SHEETS.ORDER_ITEMS) return;
  if (e.range.getRow() < 2 || e.range.getColumn() !== 7) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(10000)) return;

  try {
    const row = e.range.getRow();

    try {
      applyOrderItemStatus(row, e.range.getValue());
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Order item status updated!',
        'Inventory System'
      );
    } catch (err) {
      e.range.setValue(e.oldValue || 'PENDING');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        err.message,
        'Inventory Error'
      );
    }
  } finally {
    lock.releaseLock();
  }
}

// UPDATE ORDER STATUS
function updateOrderStatus() {
  updateStatus(SHEETS.ORDERS, 3, ORDER_STATUSES, 'Order');
}

// UPDATE PAYMENT STATUS
function updatePaymentStatus() {
  updateStatus(SHEETS.ORDERS, 4, PAYMENT_STATUSES, 'Payment');
}

function updateStatus(sheetName, column, options, label) {
  const ui = SpreadsheetApp.getUi();

  const idPrompt = ui.prompt(
    'Update ' + label + ' Status',
    'Enter Order ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (idPrompt.getSelectedButton() !== ui.Button.OK) return;

  const statusPrompt = ui.prompt(
    'Update ' + label + ' Status',
    'Enter: ' + options.join(', '),
    ui.ButtonSet.OK_CANCEL
  );

  if (statusPrompt.getSelectedButton() !== ui.Button.OK) return;

  try {
    setOrderStatusRecord(
      idPrompt.getResponseText(),
      column,
      statusPrompt.getResponseText(),
      options
    );
    ui.alert(label + ' status updated successfully!');
  } catch (err) {
    ui.alert(err.message);
  }
}

function setOrderStatusRecord(idRaw, column, statusRaw, options) {
  const id = Number(idRaw);
  const status = String(statusRaw || '').trim().toUpperCase();

  if (!Number.isFinite(id) || !options.includes(status)) {
    throw new Error('Invalid status!');
  }

  const sheet = getSheet(SHEETS.ORDERS);
  const row = findRow(sheet, 1, id);

  if (!row) {
    throw new Error('Order not found!');
  }

  sheet.getRange(row, column).setValue(status);
}
