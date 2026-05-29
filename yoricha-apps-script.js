// ============================================================
//  YORICHA — Google Apps Script Backend v3
//  Sheets: Members, History, Menu, Orders
// ============================================================

var SPREADSHEET_ID = '1her5toztbL38dtNy9Y0S9W7qzB3jUERKvbxSv4_jzpg';
var SHEET_MEMBERS  = 'Members';
var SHEET_HISTORY  = 'History';
var SHEET_MENU     = 'Menu';
var SHEET_ORDERS   = 'Orders';
var STAFF_PASSWORD = '1234';
var MAPS_KEY       = 'AIzaSyDKG3-6nsURa-1MT8gtxgLWZR64UuywjMQ';
var SHOP_ADDR      = '24/10 Đường số 6, Linh Chiểu, Thủ Đức, TP.HCM';

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet(e)  { return route(e); }
function doPost(e) { return route(e); }

function route(e) {
  try {
    var params = e.parameter || {};
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var p = {};
    for (var k in params) p[k] = params[k];
    for (var k in body)   p[k] = body[k];

    if (p.action === 'register')       return makeResponse(registerMember(p));
    if (p.action === 'lookup')         return makeResponse(lookupMember(p));
    if (p.action === 'purchase')       return makeResponse(addPurchase(p));
    if (p.action === 'getMenu')        return makeResponse(getMenu());
    if (p.action === 'saveMenu')       return makeResponse(saveMenu(p));
    if (p.action === 'placeOrder')     return makeResponse(placeOrder(p));
    if (p.action === 'getOrders')      return makeResponse(getOrders(p));
    if (p.action === 'updateOrder')    return makeResponse(updateOrder(p));
    if (p.action === 'getOrderStatus') return makeResponse(getOrderStatus(p));
    if (p.action === 'getMembers')     return makeResponse(getMembers(p));
    if (p.action === 'calcDistance')   return makeResponse(calcDistance(p));

    return makeResponse({ ok: false, error: 'Unknown action: ' + p.action });
  } catch (err) {
    return makeResponse({ ok: false, error: err.message });
  }
}

// ── HELPERS ──────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Không tìm thấy sheet "' + name + '"');
  return sheet;
}
function normalizePhone(phone) {
  var p = String(phone).replace(/\s+/g, '');
  if (p.charAt(0) === '0') p = '84' + p.slice(1);
  return p;
}
function nowVN() {
  return new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
}
function makeOrderID() {
  return 'YC' + new Date().getTime().toString().slice(-6);
}
function authStaff(password) {
  return password === STAFF_PASSWORD;
}

// ── MEMBERS ───────────────────────────────────────────────────
function registerMember(p) {
  var name = p.name; var phone = p.phone;
  if (!name || !phone) return { ok: false, error: 'Thiếu tên hoặc số điện thoại' };
  var ph = normalizePhone(phone);
  var sheet = getSheet(SHEET_MEMBERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === ph) return { ok: false, error: 'Số điện thoại đã được đăng ký' };
  }
  sheet.appendRow([ph, name.trim(), 0, nowVN()]);
  return { ok: true, member: { phone: ph, name: name.trim(), stamps: 0, history: [] } };
}

function lookupMember(p) {
  var phone = p.phone;
  if (!phone) return { ok: false, error: 'Thiếu số điện thoại' };
  var ph = normalizePhone(phone);
  var mSheet = getSheet(SHEET_MEMBERS);
  var mData = mSheet.getDataRange().getValues();
  var member = null;
  for (var i = 1; i < mData.length; i++) {
    if (String(mData[i][0]) === ph) {
      member = { phone: String(mData[i][0]), name: String(mData[i][1]), stamps: Number(mData[i][2]) };
      break;
    }
  }
  if (!member) return { ok: false, error: 'Không tìm thấy thành viên' };
  var hSheet = getSheet(SHEET_HISTORY);
  var hData = hSheet.getDataRange().getValues();
  var history = [];
  for (var j = 1; j < hData.length; j++) {
    if (String(hData[j][0]) === ph) {
      history.push({ product: String(hData[j][1]), stamp: Number(hData[j][2]), date: String(hData[j][3]) });
    }
  }
  member.history = history;
  return { ok: true, member: member };
}

function addPurchase(p) {
  if (!authStaff(p.password)) return { ok: false, error: 'Sai mật khẩu' };
  var phone = p.phone; var product = p.product;
  if (!phone || !product) return { ok: false, error: 'Thiếu thông tin' };
  var ph = normalizePhone(phone);
  var mSheet = getSheet(SHEET_MEMBERS);
  var mData = mSheet.getDataRange().getValues();
  var memberRow = -1; var stamps = 0;
  for (var i = 1; i < mData.length; i++) {
    if (String(mData[i][0]) === ph) { memberRow = i + 1; stamps = Number(mData[i][2]); break; }
  }
  if (memberRow === -1) return { ok: false, error: 'Không tìm thấy thành viên' };
  var newStamps = (stamps >= 10) ? 1 : stamps + 1;
  mSheet.getRange(memberRow, 3).setValue(newStamps);
  getSheet(SHEET_HISTORY).appendRow([ph, product, newStamps, nowVN()]);
  return lookupMember({ phone: ph });
}

function getMembers(p) {
  if (!authStaff(p.password)) return { ok: false, error: 'Sai mật khẩu' };
  var mSheet = getSheet(SHEET_MEMBERS);
  var mData = mSheet.getDataRange().getValues();
  var members = [];
  for (var i = 1; i < mData.length; i++) {
    if (!mData[i][0]) continue;
    members.push({ phone: String(mData[i][0]), name: String(mData[i][1]), stamps: Number(mData[i][2]), registeredAt: String(mData[i][3]) });
  }
  return { ok: true, members: members };
}

// ── MENU ──────────────────────────────────────────────────────
function getMenu() {
  var sheet = getSheet(SHEET_MENU);
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var imageRaw = data[i][7] ? String(data[i][7]) : '';
    // Convert Drive share link to thumbnail if needed
    var imageUrl = '';
    if (imageRaw) {
      var match = imageRaw.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (match) {
        imageUrl = 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400';
      } else {
        imageUrl = imageRaw;
      }
    }
    items.push({
      id: String(data[i][0]), name: String(data[i][1]), category: String(data[i][2]),
      price_s: Number(data[i][3]), price_l: Number(data[i][4]),
      description: String(data[i][5]),
      available: String(data[i][6]).toUpperCase() === 'TRUE' || data[i][6] === true,
      image: imageUrl
    });
  }
  return { ok: true, menu: items };
}

function saveMenu(p) {
  if (!authStaff(p.password)) return { ok: false, error: 'Sai mật khẩu' };
  var items = typeof p.items === 'string' ? JSON.parse(p.items) : p.items;
  var sheet = getSheet(SHEET_MENU);
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 7).clearContent();
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    sheet.appendRow([it.id, it.name, it.category, it.price_s, it.price_l, it.description, it.available, it.image || '']);
  }
  return { ok: true };
}

// ── ORDERS ────────────────────────────────────────────────────
function placeOrder(p) {
  var phone = p.phone || '';
  var name  = p.name  || 'Khách';
  var note  = p.note  || '';
  var items = typeof p.items === 'string' ? JSON.parse(p.items) : p.items;
  if (!items || !items.length) return { ok: false, error: 'Giỏ hàng trống' };

  var total = p.total ? Number(p.total) : 0;
  if (!total) {
    for (var i = 0; i < items.length; i++) total += items[i].price * items[i].qty;
  }

  var orderID = makeOrderID();
  var itemsStr = items.map(function(it) {
    var parts = [it.name, '('+it.size+')'];
    if (it.bot) parts.push(it.bot);
    if (it.sua) parts.push(it.sua);
    if (it.sweet) parts.push('Ngọt '+it.sweet);
    if (it.note) parts.push('📝'+it.note);
    parts.push('x'+it.qty);
    return parts.join(' ');
  }).join(' | ');

  getSheet(SHEET_ORDERS).appendRow([orderID, phone ? normalizePhone(phone) : '', name, itemsStr, total, 'Mới', note, nowVN(), '']);
  return { ok: true, orderID: orderID, total: total };
}

function getOrders(p) {
  if (!authStaff(p.password)) return { ok: false, error: 'Sai mật khẩu' };
  var sheet = getSheet(SHEET_ORDERS);
  var data = sheet.getDataRange().getValues();
  var orders = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    orders.push({
      orderID: String(data[i][0]), phone: String(data[i][1]), name: String(data[i][2]),
      items: String(data[i][3]), total: Number(data[i][4]), status: String(data[i][5]),
      note: String(data[i][6]), createdAt: String(data[i][7]), updatedAt: String(data[i][8]), row: i + 1
    });
  }
  orders.reverse();
  return { ok: true, orders: orders };
}

function updateOrder(p) {
  if (!authStaff(p.password)) return { ok: false, error: 'Sai mật khẩu' };
  var orderID = p.orderID; var status = p.status;
  if (!orderID || !status) return { ok: false, error: 'Thiếu thông tin' };
  var sheet = getSheet(SHEET_ORDERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === orderID) {
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 9).setValue(nowVN());
      // Auto cộng điểm khi Xong
      if (status === 'Xong' && data[i][1]) {
        var phone = String(data[i][1]);
        var itemsStr = String(data[i][3]);
        var itemList = itemsStr.split(' | ');
        var mSheet = getSheet(SHEET_MEMBERS);
        var mData = mSheet.getDataRange().getValues();
        for (var r = 1; r < mData.length; r++) {
          if (String(mData[r][0]) === phone) {
            var stamps = Number(mData[r][2]);
            for (var s = 0; s < itemList.length; s++) {
              stamps = (stamps >= 10) ? 1 : stamps + 1;
              getSheet(SHEET_HISTORY).appendRow([phone, itemList[s].split(' (')[0], stamps, nowVN()]);
            }
            mSheet.getRange(r + 1, 3).setValue(stamps);
            break;
          }
        }
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Không tìm thấy đơn' };
}

function getOrderStatus(p) {
  var orderID = p.orderID;
  if (!orderID) return { ok: false, error: 'Thiếu orderID' };
  var sheet = getSheet(SHEET_ORDERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === orderID) return { ok: true, status: String(data[i][5]) };
  }
  return { ok: false, error: 'Không tìm thấy đơn' };
}

// ── DISTANCE (via Google Maps) ────────────────────────────────
function calcDistance(p) {
  var destination = p.destination;
  if (!destination) return { ok: false, error: 'Thiếu địa chỉ' };
  try {
    var url = 'https://maps.googleapis.com/maps/api/distancematrix/json'
      + '?origins=' + encodeURIComponent(SHOP_ADDR)
      + '&destinations=' + encodeURIComponent(destination)
      + '&key=' + MAPS_KEY
      + '&language=vi';
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());
    if (result.status !== 'OK') return { ok: false, error: 'Không tìm được địa chỉ' };
    var element = result.rows[0].elements[0];
    if (element.status !== 'OK') return { ok: false, error: 'Địa chỉ không hợp lệ' };
    var meters = element.distance.value;
    var km = meters / 1000;
    var text = element.distance.text;
    return { ok: true, km: km, text: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// NOTE: getMenu() đã được cập nhật để đọc thêm cột Image (cột H = index 7)
// Hàm getMenu cũ sẽ được override bởi hàm dưới đây - XOÁ hàm getMenu cũ và dùng cái này
