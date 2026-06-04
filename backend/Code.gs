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
 * Main entry point for GET requests (CORS-friendly for data retrieval)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    let response;

    if (action === 'GET_INVENTORY') {
      response = InventoryService.getFiltered({
        page: parseInt(e.parameter.page) || 1,
        pageSize: parseInt(e.parameter.pageSize) || 12,
        category: e.parameter.category || '',
        sortBy: e.parameter.sortBy || 'name'
      });
    } else {
      throw new Error('Invalid GET action');
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
 * Inventory Service - Stock management with Pagination, Filtering, and Sorting
 */
const InventoryService = {
  getFiltered: function(params) {
    const {
      page = 1,
      pageSize = 12,
      category = '',
      search = '',
      sortBy = 'name', // 'name', 'price_asc', 'price_desc'
      status = 'Active'
    } = params;

    const sheet = Repository.getSheet('Inventory');
    let data = Repository.getAll(sheet); // Array of arrays [[ID, Category, Name, Stock, Price, Status], ...]

    // 1. Filtering
    if (category) {
      data = data.filter(row => row[1] === category);
    }
    if (search) {
      const query = search.toLowerCase();
      data = data.filter(row => row[2].toLowerCase().includes(query));
    }
    if (status) {
      data = data.filter(row => row[5] === status);
    }

    // 2. Sorting
    data.sort((a, b) => {
      switch(sortBy) {
        case 'price_asc': return a[4] - b[4];
        case 'price_desc': return b[4] - a[4];
        case 'name': return a[2].localeCompare(b[2]);
        default: return 0;
      }
    });

    // 3. Pagination
    const totalCount = data.length;
    const startIndex = (page - 1) * pageSize;
    const paginatedData = data.slice(startIndex, startIndex + pageSize);

    return {
      products: paginatedData.map(row => ({
        id: row[0],
        category: row[1],
        name: row[2],
        stock: row[3],
        price: row[4],
        status: row[5]
      })),
      pagination: {
        currentPage: page,
        pageSize: pageSize,
        totalCount: totalCount,
        totalPages: Math.ceil(totalCount / pageSize)
      }
    };
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
