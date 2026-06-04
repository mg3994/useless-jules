/**
 * Antinna Backend - Google Apps Script
 * Core logic for handling e-commerce and service marketplace data.
 *
 * Clean Architecture:
 * - Handlers (doPost)
 * - Controllers (Action Router)
 * - Services (Business Logic)
 * - Repository (Google Sheets Abstraction)
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // Placeholder

/**
 * Main entry point for POST requests
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    let response;
    switch(action) {
      case 'SYNC_DEVICE':
        response = DeviceService.sync(payload);
        break;
      case 'LOGOUT_DEVICE':
        response = DeviceService.logout(payload);
        break;
      case 'CHECKOUT_ORDER':
        response = OrderService.checkout(payload);
        break;
      case 'VERIFY_PAYMENT':
        response = OrderService.verify(payload);
        break;
      case 'BOOK_REPAIR':
        response = ServiceService.book(payload);
        break;
      case 'GET_INVENTORY':
        response = InventoryService.getAll();
        break;
      default:
        throw new Error('Invalid action: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: response
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Device Service - Manages FCM tokens and client sessions
 */
const DeviceService = {
  sync: function(data) {
    const sheet = Repository.getSheet('Devices');
    const rows = Repository.getAll(sheet);

    const rowIndex = rows.findIndex(row => row[0] === data.clientId);
    const timestamp = new Date();

    if (rowIndex > -1) {
      Repository.updateRow(sheet, rowIndex + 2, [
        data.clientId,
        data.uid || 'guest',
        data.deviceToken,
        data.clientName,
        timestamp
      ]);
    } else {
      Repository.appendRow(sheet, [
        data.clientId,
        data.uid || 'guest',
        data.deviceToken,
        data.clientName,
        timestamp
      ]);
    }
    return { synced: true };
  },

  logout: function(data) {
    // Optional: Implement logic to clear/invalidate tokens on logout
    return { loggedOut: true };
  }
};

/**
 * Order Service - Manages e-commerce transactions
 */
const OrderService = {
  checkout: function(data) {
    const sheet = Repository.getSheet('Orders');
    const orderId = 'ORD-' + Date.now();
    Repository.appendRow(sheet, [
      orderId,
      new Date(),
      data.uid,
      data.name,
      data.phone,
      JSON.stringify(data.items),
      data.total,
      '', // UPI_TxnRef
      'Pending'
    ]);
    return { orderId: orderId };
  },

  verify: function(data) {
    const sheet = Repository.getSheet('Orders');
    const rows = Repository.getAll(sheet);
    const rowIndex = rows.findIndex(row => row[0] === data.orderId);

    if (rowIndex > -1) {
      Repository.updateCell(sheet, rowIndex + 2, 8, data.txnRef);
      Repository.updateCell(sheet, rowIndex + 2, 9, 'Success');
      return { verified: true };
    }
    throw new Error('Order not found');
  }
};

/**
 * Service Service - Manages repair bookings
 */
const ServiceService = {
  book: function(data) {
    const sheet = Repository.getSheet('ServiceBookings');
    const bookingId = 'BK-' + Date.now();
    Repository.appendRow(sheet, [
      bookingId,
      new Date(),
      data.uid,
      data.serviceType,
      data.details,
      data.scheduledDate,
      'Pending'
    ]);
    return { bookingId: bookingId };
  }
};

/**
 * Inventory Service - Stock management
 */
const InventoryService = {
  getAll: function() {
    const sheet = Repository.getSheet('Inventory');
    return Repository.getAll(sheet);
  }
};

/**
 * Repository - Low-level Google Sheets operations
 */
const Repository = {
  getSheet: function(name) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getSheetByName(name) || ss.insertSheet(name);
  },

  getAll: function(sheet) {
    const range = sheet.getDataRange();
    return range.getValues().slice(1); // Exclude header
  },

  appendRow: function(sheet, data) {
    sheet.appendRow(data);
  },

  updateRow: function(sheet, rowIndex, data) {
    sheet.getRange(rowIndex, 1, 1, data.length).setValues([data]);
  },

  updateCell: function(sheet, rowIndex, colIndex, value) {
    sheet.getRange(rowIndex, colIndex).setValue(value);
  }
};
