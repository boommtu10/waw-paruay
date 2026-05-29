// ========================================
// วาว พารวย - ระบบรับหวย
// Code.gs - Google Apps Script Backend
// ========================================

const SHEET_AGENTS = 'เจ้าที่ส่ง';
const SHEET_BILLS = 'บิลลูกค้า';
const SHEET_ITEMS = 'รายการเลข';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('วาว พารวย - ระบบรับหวย')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Set headers
    if (name === SHEET_AGENTS) {
      sheet.appendRow([
        'id','ชื่อ','งวด',
        'จ่าย3ตัวตรง','จ่าย3ตัวโต๊ด','จ่าย2ตัวบน','จ่าย2ตัวล่าง',
        'จ่าย4ตัวโต๊ดบน','จ่าย5ตัวโต๊ดบน','จ่ายวิ่งบน','จ่ายวิ่งล่าง',
        'เปอร์เซ็นต์','เลขอั้น','วันที่บันทึก'
      ]);
    } else if (name === SHEET_BILLS) {
      sheet.appendRow([
        'billId','ชื่อลูกค้า','งวด','วันที่บันทึก','ยอดรวมทั้งหมด'
      ]);
    } else if (name === SHEET_ITEMS) {
      sheet.appendRow([
        'itemId','billId','ชื่อลูกค้า','งวด','ชื่อเจ้า','ประเภท',
        'เลข','ราคาตรง','ราคาโต๊ด','ตัวเลือกX_T','ราคารวม',
        'หมายเหตุ','วันที่บันทึก'
      ]);
    }
  }
  return sheet;
}

// ===== AGENT FUNCTIONS =====
function saveAgent(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_AGENTS);
    const id = 'A' + Date.now();
    const banned = "'" + JSON.stringify(data.bannedNumbers || []);
    const periodSafe = "'" + normalizePeriod(data.period);
    sheet.appendRow([
      id, data.name, periodSafe,
      data.pay3straight, data.pay3tod, data.pay2top, data.pay2bottom,
      data.pay4tod, data.pay5tod, data.payRunTop, data.payRunBottom,
      data.percent, banned, new Date().toLocaleString('th-TH'),
      data.percentRun || 0  // col 15 (index 14): เปอร์เซ็นต์วิ่ง
    ]);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ===== อัปเดตเลขอั้นของเจ้าที่ส่ง (ไม่กระทบบิลที่บันทึกไปแล้ว) =====
function updateAgentBanned(id, bannedNumbers) {
  try {
    const sheet = getOrCreateSheet(SHEET_AGENTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const banned = "'" + JSON.stringify(bannedNumbers || []);
        sheet.getRange(i + 1, 13).setValue(banned); // col 13 = เลขอั้น (index 12)
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบเจ้า id: ' + id };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ===== DEBUG: เรียกจาก Console เพื่อตรวจสอบว่า period ใน Sheet เก็บอะไร =====
function debugBills() {
  const sheet = getOrCreateSheet(SHEET_BILLS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < Math.min(data.length, 6); i++) {
    const raw = data[i][2];
    result.push({
      row: i+1,
      rawType: typeof raw,
      isDate: raw instanceof Date,
      rawValue: String(raw),
      normalized: normalizePeriod(raw)
    });
  }
  Logger.log(JSON.stringify(result));
  return result;
}

function getAgents() {
  try {
    const sheet = getOrCreateSheet(SHEET_AGENTS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, agents: [] };
    const agents = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      let banned = [];
      try {
        // ลบ apostrophe ที่เติมไว้ป้องกัน Sheets ออกก่อน parse
        const rawBanned = String(row[12] || '[]').replace(/^'/, '');
        banned = JSON.parse(rawBanned);
        if (!Array.isArray(banned)) banned = [];
      } catch(e) { banned = []; }
      agents.push({
        id: String(row[0]), name: String(row[1]), period: normalizePeriod(row[2]),
        pay3straight: row[3], pay3tod: row[4], pay2top: row[5], pay2bottom: row[6],
        pay4tod: row[7], pay5tod: row[8], payRunTop: row[9], payRunBottom: row[10],
        percent: row[11], bannedNumbers: banned,
        percentRun: row[14] !== undefined && row[14] !== '' ? Number(row[14]) : Number(row[11]) || 0
      });
    }
    return { success: true, agents: agents };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function deleteAgent(id) {
  try {
    const sheet = getOrCreateSheet(SHEET_AGENTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบข้อมูล' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// นับจำนวนบิลทั้งหมดในระบบ (สำหรับ sync billCounter)
function getBillCount() {
  try {
    const sheet = getOrCreateSheet(SHEET_BILLS);
    const data = sheet.getDataRange().getValues();
    // ลบ header row ออก
    const count = Math.max(0, data.length - 1);
    return { success: true, count: count };
  } catch(e) {
    return { success: false, count: 0, error: e.toString() };
  }
}


function saveBill(billData) {
  try {
    const billSheet = getOrCreateSheet(SHEET_BILLS);
    const itemSheet = getOrCreateSheet(SHEET_ITEMS);
    const billId = 'B' + Date.now();
    const now = new Date().toLocaleString('th-TH');
    // เติม ' นำหน้าเพื่อป้องกัน Sheets แปลงงวดเป็น Date อัตโนมัติ
    const periodSafe = "'" + normalizePeriod(billData.period);

    billSheet.appendRow([
      billId, billData.customerName, periodSafe, now, billData.grandTotal
    ]);

    // Save each item
    billData.items.forEach((item, idx) => {
      const itemId = billId + '_' + (idx + 1);
      itemSheet.appendRow([
        itemId, billId, billData.customerName, periodSafe,
        item.agentName, item.type,
        item.number, item.priceMain, item.priceTod || 0,
        item.xOrT || '', item.totalPrice,
        item.note || '', now
      ]);
    });

    return { success: true, billId: billId };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// แปลงค่างวดจาก Sheets ให้เป็น string dd/mm/yy พ.ศ. เสมอ
function normalizePeriod(val) {
  if (!val) return '';
  if (val instanceof Date) {
    // ใช้ timezone ไทย UTC+7
    const offset = 7 * 60;
    const local = new Date(val.getTime() + offset * 60000);
    const d = String(local.getUTCDate()).padStart(2,'0');
    const m = String(local.getUTCMonth() + 1).padStart(2,'0');
    const beYear = local.getUTCFullYear() + 543;
    const yy = String(beYear).slice(-2);
    return d + '/' + m + '/' + yy;
  }
  // ลบ apostrophe นำหน้า (ที่เราเติมไว้ป้องกัน Sheets) ออก
  return String(val).trim().replace(/^'/, '');
}

function getBills(filter) {
  try {
    const billSheet = getOrCreateSheet(SHEET_BILLS);
    const itemSheet = getOrCreateSheet(SHEET_ITEMS);
    const billData = billSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();

    if (billData.length <= 1) return { success: true, bills: [] };

    const filterPeriod = filter && filter.period ? normalizePeriod(filter.period) : '';

    // index items by billId for fast lookup
    const itemsByBill = {};
    for (let j = 1; j < itemData.length; j++) {
      const irow = itemData[j];
      if (!irow[0]) continue;
      const bid = String(irow[1]).trim();
      if (!itemsByBill[bid]) itemsByBill[bid] = [];
      itemsByBill[bid].push({
        itemId: String(irow[0]), agentName: String(irow[4]), type: String(irow[5]),
        number: String(irow[6]), priceMain: Number(irow[7]) || 0, priceTod: Number(irow[8]) || 0,
        xOrT: String(irow[9] || ''), totalPrice: Number(irow[10]) || 0, note: String(irow[11] || '')
      });
    }

    const bills = [];
    for (let i = 1; i < billData.length; i++) {
      const row = billData[i];
      if (!row[0]) continue;
      const periodVal = normalizePeriod(row[2]);
      const billId = String(row[0]).trim();

      // filter
      if (filterPeriod && periodVal !== filterPeriod) continue;
      if (filter && filter.customerName && String(row[1]) !== String(filter.customerName)) continue;

      const bill = {
        billId: billId,
        customerName: String(row[1]),
        period: periodVal,
        date: row[3],
        grandTotal: Number(row[4]) || 0,
        items: itemsByBill[billId] || []
      };
      bills.push(bill);
    }
    return { success: true, bills: bills };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// เรียกครั้งเดียวได้ทั้ง customer, agent, all เพื่อลดจำนวน round trips
function getAllSummary(filter) {
  try {
    const billSheet = getOrCreateSheet(SHEET_BILLS);
    const itemSheet = getOrCreateSheet(SHEET_ITEMS);
    const agentSheet = getOrCreateSheet(SHEET_AGENTS);

    const billData = billSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();
    const agentData = agentSheet.getDataRange().getValues();

    const filterPeriod = filter && filter.period ? normalizePeriod(filter.period) : '';

    // สร้าง agent percent map
    const agentPercentMap = {};
    const agentPercentRunMap = {};
    for (var ai = 1; ai < agentData.length; ai++) {
      const ar = agentData[ai];
      if (ar[0]) {
        agentPercentMap[String(ar[1])] = Number(ar[11]) || 0;
        // col 15 (index 14) = percentRun; ถ้าไม่มีให้ใช้ percent ปกติ
        agentPercentRunMap[String(ar[1])] = (ar[14] !== undefined && ar[14] !== '') ? Number(ar[14]) : Number(ar[11]) || 0;
      }
    }

    // index bill grandTotal
    const billGrandTotal = {};
    const billCustomer = {};
    const billPeriod = {};
    const billDate = {};
    const billNetTransfer = {};      // col 7 (index 6)
    const billAdjustedBet = {};      // col 9 (index 8)
    for (var bi = 1; bi < billData.length; bi++) {
      const br = billData[bi];
      if (!br[0]) continue;
      const bid = String(br[0]).trim();
      const per = normalizePeriod(br[2]);
      if (filterPeriod && per !== filterPeriod) continue;
      billGrandTotal[bid] = Number(br[4]) || 0;
      billCustomer[bid] = String(br[1]);
      billPeriod[bid] = per;
      billDate[bid] = br[3] ? String(br[3]) : '';
      // prize columns (บันทึกจาก savePrizeResult)
      billNetTransfer[bid] = (br[6] !== '' && br[6] !== undefined && br[6] !== null) ? Number(br[6]) : null;
      billAdjustedBet[bid] = (br[8] !== '' && br[8] !== undefined && br[8] !== null) ? Number(br[8]) : null;
    }

    // คำนวณจาก items
    const customerTotal = {};   // { customerName: total }
    const agentTotal = {};      // { agentName: total }
    const agentRunTotal = {};   // { agentName: total เฉพาะ runtop+runbot }
    const agentBills = {};      // { agentName: { billId: { customerName, total } } }
    const billItems = {};       // { billId: [ {agentName,type,number,priceMain,priceTod,xOrT,totalPrice,note} ] }

    for (var ii = 1; ii < itemData.length; ii++) {
      const ir = itemData[ii];
      if (!ir[0]) continue;
      const bid = String(ir[1]).trim();
      if (!billCustomer[bid]) continue; // ไม่อยู่ใน filter
      const agentName = String(ir[4]);
      const price = Number(ir[10]) || 0;
      const itemType = String(ir[5]);

      if (!agentTotal[agentName]) agentTotal[agentName] = 0;
      agentTotal[agentName] += price;

      // แยก run total สำหรับ percentRun
      if (itemType === 'runtop' || itemType === 'runbot') {
        if (!agentRunTotal[agentName]) agentRunTotal[agentName] = 0;
        agentRunTotal[agentName] += price;
      }

      if (!agentBills[agentName]) agentBills[agentName] = {};
      if (!agentBills[agentName][bid]) agentBills[agentName][bid] = { customerName: billCustomer[bid], total: 0 };
      agentBills[agentName][bid].total += price;

      // เก็บ items ต่อบิล สำหรับแสดงรายละเอียด
      if (!billItems[bid]) billItems[bid] = [];
      billItems[bid].push({
        agentName: agentName,
        type: itemType,
        number: String(ir[6]),
        priceMain: Number(ir[7]) || 0,
        priceTod: Number(ir[8]) || 0,
        xOrT: String(ir[9] || ''),
        totalPrice: price,
        note: String(ir[11] || '')
      });
    }

    // customer total มาจาก grandTotal ของ bill (ไม่ใช่ items เพื่อให้ตรงกับที่ลูกค้าจ่าย)
    const customerBills = {}; // { customerName: [ {billId, period, date, total, items} ] }
    Object.keys(billCustomer).forEach(function(bid) {
      const name = billCustomer[bid];
      if (!customerTotal[name]) customerTotal[name] = 0;
      customerTotal[name] += billGrandTotal[bid];
      if (!customerBills[name]) customerBills[name] = [];
      customerBills[name].push({
        billId: bid,
        period: billPeriod[bid],
        date: billDate[bid],
        total: billGrandTotal[bid],
        items: billItems[bid] || [],
        netTransfer:      billNetTransfer[bid],
        adjustedBetTotal: billAdjustedBet[bid]
      });
    });

    // สร้าง agentMap สำหรับ tab agent
    const agentMap = {};
    Object.keys(agentBills).forEach(function(name) {
      const pct = agentPercentMap[name] || 0;
      const pctRun = agentPercentRunMap[name] || pct;
      const tot = agentTotal[name] || 0;
      const runTot = agentRunTotal[name] || 0;
      const nonRunTot = tot - runTot;
      const profit = Math.round(nonRunTot * pct / 100) + Math.round(runTot * pctRun / 100);
      agentMap[name] = {
        total: tot, percent: pct, percentRun: pctRun,
        runTotal: runTot,
        profit: profit,
        bills: Object.keys(agentBills[name]).map(function(bid) {
          return { billId: bid, customerName: agentBills[name][bid].customerName, total: agentBills[name][bid].total };
        })
      };
    });

    // สร้าง allAgentProfits สำหรับ tab all/profit
    var totalProfit = 0;
    const agentProfits = {};
    Object.keys(agentTotal).forEach(function(name) {
      const pct = agentPercentMap[name] || 0;
      const pctRun = agentPercentRunMap[name] || pct;
      const tot = agentTotal[name] || 0;
      const runTot = agentRunTotal[name] || 0;
      const nonRunTot = tot - runTot;
      const profit = Math.round(nonRunTot * pct / 100) + Math.round(runTot * pctRun / 100);
      agentProfits[name] = { total: tot, percent: pct, percentRun: pctRun, runTotal: runTot, profit: profit };
      totalProfit += profit;
    });

    return {
      success: true,
      customer: customerTotal,
      customerBills: customerBills,
      agent: agentMap,
      all: { customers: customerTotal, agents: agentProfits, totalProfit: totalProfit }
    };
  } catch(e) {
    Logger.log('getAllSummary error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}


// ===== ลบบิล (billId) พร้อม items ทั้งหมดที่เกี่ยวข้อง =====
function deleteBill(billId) {
  try {
    const billSheet = getOrCreateSheet(SHEET_BILLS);
    const itemSheet = getOrCreateSheet(SHEET_ITEMS);

    // ลบ rows ใน SHEET_ITEMS ที่มี billId ตรงกัน (ลบจากด้านล่างขึ้นบนเพื่อ index ไม่เลื่อน)
    const itemData = itemSheet.getDataRange().getValues();
    for (let i = itemData.length - 1; i >= 1; i--) {
      if (String(itemData[i][1]).trim() === String(billId).trim()) {
        itemSheet.deleteRow(i + 1);
      }
    }

    // ลบ row ใน SHEET_BILLS
    const billData = billSheet.getDataRange().getValues();
    for (let i = 1; i < billData.length; i++) {
      if (String(billData[i][0]).trim() === String(billId).trim()) {
        billSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบบิล: ' + billId };
  } catch(e) {
    Logger.log('deleteBill error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ===== PAGE 4: บันทึกผลรางวัล — อัปเดต adjustedBetTotal ในชีต Bills =====
function savePrizeResult(data) {
  try {
    var billSheet = getOrCreateSheet(SHEET_BILLS);
    var billData  = billSheet.getDataRange().getValues();

    // หา row ของ billId นี้ แล้วอัปเดต col F (index 5) = adjustedBetTotal
    // และ col G (index 6) = winTotal, col H (index 7) = netTransfer
    // ถ้า column ยังไม่มีให้เพิ่ม header ก่อน
    var headers = billData[0];
    // ตรวจว่ามี header ครบไหม ถ้าไม่มีให้เพิ่ม
    if (headers.length < 6) billSheet.getRange(1, 6).setValue('ยอดถูกรวม');
    if (headers.length < 7) billSheet.getRange(1, 7).setValue('ยอดสุทธิ');
    if (headers.length < 8) billSheet.getRange(1, 8).setValue('โอนให้ลูกค้า');
    if (headers.length < 9) billSheet.getRange(1, 9).setValue('ยอดแทงคงเหลือ');

    for (var i = 1; i < billData.length; i++) {
      if (String(billData[i][0]).trim() === String(data.billId).trim()) {
        var row = i + 1;
        billSheet.getRange(row, 6).setValue(data.winTotal      || 0); // ยอดถูกรวม
        billSheet.getRange(row, 7).setValue(data.netTransfer   || 0); // ยอดสุทธิ (บวก=โอน, ลบ=ค้าง)
        billSheet.getRange(row, 8).setValue(data.transferAmt   || 0); // โอนให้ลูกค้า
        billSheet.getRange(row, 9).setValue(data.adjustedBetTotal || 0); // ยอดแทงคงเหลือ
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบ billId: ' + data.billId };
  } catch(e) {
    Logger.log('savePrizeResult error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ===== ล้างข้อมูลทั้งหมดทุกชีต (เริ่มงวดใหม่) =====
function clearAllData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsToWipe = [SHEET_BILLS, SHEET_ITEMS, SHEET_AGENTS];
    sheetsToWipe.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) return;
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    });
    return { success: true };
  } catch(e) {
    Logger.log('clearAllData error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
// ===== ดึงบิลตามงวด =====
function getBillsByPeriod(periodStr) {
  try {
    var billSheet = getOrCreateSheet(SHEET_BILLS);
    var itemSheet = getOrCreateSheet(SHEET_ITEMS);
    var billData  = billSheet.getDataRange().getValues();
    var itemData  = itemSheet.getDataRange().getValues();

    var filterPeriod = periodStr ? String(periodStr).trim() : '';

    // index items by billId
    var itemsByBill = {};
    for (var j = 1; j < itemData.length; j++) {
      var ir = itemData[j];
      if (!ir[0]) continue;
      var bid = String(ir[1]).trim();
      if (!itemsByBill[bid]) itemsByBill[bid] = [];
      itemsByBill[bid].push({
        agentName:  String(ir[4]),
        type:       String(ir[5]),
        number:     String(ir[6]),
        priceMain:  Number(ir[7])  || 0,
        priceTod:   Number(ir[8])  || 0,
        xOrT:       String(ir[9]  || ''),
        totalPrice: Number(ir[10]) || 0,
        note:       String(ir[11] || '')
      });
    }

    var bills = [];
    for (var i = 1; i < billData.length; i++) {
      var row = billData[i];
      if (!row[0]) continue;
      var per   = normalizePeriod(row[2]);
      if (filterPeriod && per !== filterPeriod) continue;
      var billId = String(row[0]).trim();
      bills.push({
        billId:       billId,
        customerName: String(row[1]),
        period:       per,
        date:         row[3] ? String(row[3]) : '',
        grandTotal:   Number(row[4]) || 0,
        items:        itemsByBill[billId] || []
      });
    }
  
      return { success: true, bills: bills }; 
    
  } catch(e) {
    Logger.log('getBillsByPeriod error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ================================================================
//  วาว พารวย - Code.gs เพิ่มเติม สำหรับหน้าที่ 5: สถิติ
//  วิธีใช้: คัดลอกทั้งหมดด้านล่างต่อท้าย Code.gs เดิม
// ================================================================

const SHEET_HISTORY = 'ประวัติงวด';   // ชีตเก็บ snapshot แต่ละงวด

// ----------------------------------------------------------------
// getOrCreateHistorySheet — สร้างชีตประวัติงวด (ถ้ายังไม่มี)
// ----------------------------------------------------------------
function getOrCreateHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HISTORY);
    sheet.appendRow([
      'period',       // A: งวด dd/mm/yy
      'prize3',       // B: เลข 3 ตัวบน
      'prize2top',    // C: เลข 2 ตัวบน
      'prize2bot',    // D: เลข 2 ตัวล่าง
      'customersJSON',// E: JSON สรุปลูกค้า { name: { totalBet, totalWin } }
      'agentsJSON',   // F: JSON สรุปเจ้าที่ส่ง { name: { totalSent, totalPaid, topCustomer, topCustomerAmt } }
      'billsJSON',    // G: JSON บิลที่ถูกรางวัล [ { customerName, grandTotal, winTotal, netTransfer } ]
      'savedAt'       // H: วันที่บันทึก
    ]);
  }
  return sheet;
}

// ----------------------------------------------------------------
// archivePeriodData — บันทึก snapshot งวดปัจจุบันลงชีต ประวัติงวด
//   data = { period, prize3, prize2top, prize2bot }
// ----------------------------------------------------------------
function archivePeriodData(data) {
  try {
    var period   = String(data.period   || '').trim();
    var prize3   = String(data.prize3   || '');
    var prize2top= String(data.prize2top|| '');
    var prize2bot= String(data.prize2bot|| '');

    if (!period) return { success: false, error: 'ไม่ระบุงวด' };

    // ดึงข้อมูลบิลและ items งวดนี้
    var billSheet  = getOrCreateSheet(SHEET_BILLS);
    var itemSheet  = getOrCreateSheet(SHEET_ITEMS);
    var agentSheet = getOrCreateSheet(SHEET_AGENTS);
    var billData   = billSheet.getDataRange().getValues();
    var itemData   = itemSheet.getDataRange().getValues();
    var agentData  = agentSheet.getDataRange().getValues();

    // สร้าง map percent ของเจ้าที่ส่ง { agentName: percent }
    var agentPercentMap = {};
    var agentPercentRunMap = {};
    for (var ap = 1; ap < agentData.length; ap++) {
      var ar = agentData[ap];
      if (ar[0]) {
        agentPercentMap[String(ar[1])] = Number(ar[11]) || 0;
        // col 15 (index 14) = percentRun; ถ้าไม่มีให้ใช้ percent ปกติ
        agentPercentRunMap[String(ar[1])] = (ar[14] !== undefined && ar[14] !== '') ? Number(ar[14]) : Number(ar[11]) || 0;
      }
    }

    // index items by billId
    var itemsByBill = {};
    for (var j = 1; j < itemData.length; j++) {
      var ir = itemData[j];
      if (!ir[0]) continue;
      var bid = String(ir[1]).trim();
      if (!itemsByBill[bid]) itemsByBill[bid] = [];
      itemsByBill[bid].push({
        agentName:  String(ir[4]),
        type:       String(ir[5]),
        totalPrice: Number(ir[10]) || 0
      });
    }

    // สรุปลูกค้า + เจ้า
    var custMap  = {};  // { name: { totalBet, totalWin } }
    var agentMap = {};  // { agentName: { totalSent, totalPaid, custAmts: { custName: amt } } }
    var billSnaps = []; // [ { customerName, grandTotal, winTotal, netTransfer } ]

    for (var i = 1; i < billData.length; i++) {
      var row = billData[i];
      if (!row[0]) continue;
      var per = normalizePeriod(row[2]);
      if (per !== period) continue;

      var billId   = String(row[0]).trim();
      var custName = String(row[1]);
      var total    = Number(row[4]) || 0;
      var winTotal = Number(row[5]) || 0;   // col F (index 5) ยอดถูกรวม (จาก savePrizeResult)
      var netTrans = Number(row[6]) || 0;   // col G (index 6) ยอดสุทธิ

      // customer
      if (!custMap[custName]) custMap[custName] = { totalBet: 0, totalWin: 0 };
      custMap[custName].totalBet += total;
      custMap[custName].totalWin += winTotal;

      // bills snapshot (เฉพาะที่ถูก)
      if (winTotal > 0) {
        billSnaps.push({ customerName: custName, grandTotal: total, winTotal: winTotal, netTransfer: netTrans });
      }

      // agent — คำนวณจาก items
      var items = itemsByBill[billId] || [];
      // รวมยอดส่งต่อเจ้าในบิลนี้ก่อน แล้วค่อยกระจาย winTotal ครั้งเดียวต่อเจ้า
      // (ป้องกัน rounding error จากการ Math.round ทีละ item แล้วผลรวมคลาดเคลื่อน)
      var agentSentInBill = {}; // { agentName: ยอดรวมของเจ้านี้ในบิลนี้ }
      items.forEach(function(it) {
        var ag = it.agentName;
        if (!agentMap[ag]) agentMap[ag] = { totalSent: 0, runSent: 0, totalPaid: 0, custAmts: {} };
        agentMap[ag].totalSent += it.totalPrice;
        // แยกยอดเลขวิ่ง
        if (it.type === 'runtop' || it.type === 'runbot') {
          agentMap[ag].runSent += it.totalPrice;
        }
        if (!agentMap[ag].custAmts[custName]) agentMap[ag].custAmts[custName] = 0;
        agentMap[ag].custAmts[custName] += it.totalPrice;
        if (!agentSentInBill[ag]) agentSentInBill[ag] = 0;
        agentSentInBill[ag] += it.totalPrice;
      });
      // กระจาย winTotal ต่อเจ้า 1 ครั้ง — เจ้าสุดท้ายรับยอดคงเหลือเพื่อให้รวมได้ winTotal พอดี
      if (total > 0 && winTotal > 0) {
        var agNamesInBill = Object.keys(agentSentInBill);
        var assigned = 0;
        for (var ai = 0; ai < agNamesInBill.length - 1; ai++) {
          var agName = agNamesInBill[ai];
          var share = Math.round(winTotal * agentSentInBill[agName] / total);
          agentMap[agName].totalPaid += share;
          assigned += share;
        }
        // เจ้าสุดท้ายรับส่วนที่เหลือ ทำให้ผลรวม = winTotal เสมอ
        var lastAgName = agNamesInBill[agNamesInBill.length - 1];
        if (lastAgName) agentMap[lastAgName].totalPaid += (winTotal - assigned);
      }
    }

    // หาลูกค้าส่งเยอะสุดต่อเจ้า
    var agentsOut = {};
    Object.keys(agentMap).forEach(function(ag) {
      var entry = agentMap[ag];
      var topCust = '', topAmt = 0;
      Object.keys(entry.custAmts).forEach(function(cn) {
        if (entry.custAmts[cn] > topAmt) { topAmt = entry.custAmts[cn]; topCust = cn; }
      });
      var pct = agentPercentMap[ag] || 0;
      var pctRun = agentPercentRunMap[ag] || pct;
      var runTot = entry.runSent || 0;
      var nonRunTot = entry.totalSent - runTot;
      var profit = Math.round(nonRunTot * pct / 100) + Math.round(runTot * pctRun / 100);
      agentsOut[ag] = {
        totalSent:       entry.totalSent,
        totalPaid:       entry.totalPaid,
        percent:         pct,
        percentRun:      pctRun,
        profit:          profit,
        topCustomer:     topCust,
        topCustomerAmt:  topAmt
      };
    });

    // ตรวจว่ามีงวดนี้อยู่แล้วหรือไม่ (update แทน insert)
    var histSheet = getOrCreateHistorySheet();
    var histData  = histSheet.getDataRange().getValues();
    var targetRow = -1;
    for (var h = 1; h < histData.length; h++) {
      if (String(histData[h][0]).trim() === period) { targetRow = h + 1; break; }
    }

    var rowValues = [
      "'" + period,
      prize3,
      prize2top,
      prize2bot,
      JSON.stringify(custMap),
      JSON.stringify(agentsOut),
      JSON.stringify(billSnaps),
      new Date().toLocaleString('th-TH')
    ];

    if (targetRow > 0) {
      histSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      histSheet.appendRow(rowValues);
    }

    return { success: true };
  } catch(e) {
    Logger.log('archivePeriodData error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// getStatHistory — ดึงข้อมูลทุกงวดจากชีต ประวัติงวด
// ----------------------------------------------------------------
function getStatHistory() {
  try {
    var histSheet = getOrCreateHistorySheet();
    var data = histSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, history: [] };

    var history = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var customers = {};
      var agents    = {};
      var bills     = [];
      try { customers = JSON.parse(String(row[4] || '{}')); } catch(e) {}
      try { agents    = JSON.parse(String(row[5] || '{}')); } catch(e) {}
      try { bills     = JSON.parse(String(row[6] || '[]')); } catch(e) {}

      history.push({
        period:   String(row[0]).trim().replace(/^'/,''),
        prize3:   String(row[1] || ''),
        prize2top:String(row[2] || ''),
        prize2bot:String(row[3] || ''),
        customers: customers,
        agents:    agents,
        bills:     bills,
        savedAt:  String(row[7] || '')
      });
    }

    // เรียงตาม period (dd/mm/yy → ใช้ index เวลา)
    history.sort(function(a, b) {
      return parsePeriodDate(a.period) - parsePeriodDate(b.period);
    });

    return { success: true, history: history };
  } catch(e) {
    Logger.log('getStatHistory error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// parsePeriodDate — แปลง "dd/mm/yy" (พ.ศ.) → timestamp สำหรับ sort
// ----------------------------------------------------------------
function parsePeriodDate(str) {
  try {
    var parts = String(str).split('/');
    if (parts.length < 3) return 0;
    var d  = parseInt(parts[0]) || 1;
    var m  = parseInt(parts[1]) || 1;
    var yy = parseInt(parts[2]) || 0;
    // yy เป็น พ.ศ. 2 หลัก (เช่น 68 = 2568)
    var fullYear = (yy >= 0 && yy < 100) ? 2500 + yy - 543 : yy - 543;
    return new Date(fullYear, m - 1, d).getTime();
  } catch(e) {
    return 0;
  }
}
