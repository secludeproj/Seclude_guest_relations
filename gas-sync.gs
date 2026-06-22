/**
 * Seclude — Djubo Gmail Auto-Sync
 * Google Apps Script — paste into script.google.com and run setupTrigger() once.
 *
 * What it does:
 *   • Runs every 2 hours automatically
 *   • Searches Gmail for Reservation Confirmation / Cancellation emails
 *   • Parses each email using the same logic as the browser app
 *   • Writes new/updated bookings to Firestore seclude_gr_pending
 *   • Browser app reads & consumes that queue on next page load
 */

// ─── CONFIG ─────────────────────────────────────────────────────────────────
var FIRESTORE_PROJECT = 'seclude-ops';
var FIRESTORE_API_KEY = 'AIzaSyCzbDofACiuSslCH2pVMZzNDfLLQa2_dYI';
var FIRESTORE_BASE    = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT + '/databases/(default)/documents';
var PENDING_COLLECTION = 'seclude_gr_pending';
var PROCESSED_KEY      = 'djubo_processed_ids';   // PropertiesService key
var LAST_SYNC_KEY      = 'djubo_last_sync_date';
var MAX_STORED_IDS     = 3000;
var BATCH_SIZE         = 10;

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────
function syncDjuboEmails() {
  var props   = PropertiesService.getScriptProperties();
  var processed = loadProcessedIds(props);
  var lastSync  = props.getProperty(LAST_SYNC_KEY) || '2026/06/01';
  var afterDate = gmailAfterDate(lastSync);

  var added = 0, updated = 0, cancelled = 0;

  // ── 1. Confirmation emails ──────────────────────────────────────────────
  var threads = GmailApp.search('subject:"Reservation Confirmation" after:' + afterDate);
  Logger.log('[Sync] found ' + threads.length + ' confirmation threads after ' + afterDate);

  var toProcess = [];
  var consecutiveKnown = 0;
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var msgId = msgs[j].getId();
      if (processed.has(msgId)) {
        if (++consecutiveKnown >= 5) { Logger.log('[Sync] 5 consecutive known — stopping early'); break; }
        continue;
      }
      consecutiveKnown = 0;
      toProcess.push(msgs[j]);
    }
    if (consecutiveKnown >= 5) break;
  }

  Logger.log('[Sync] ' + toProcess.length + ' new confirmation messages to process');

  for (var k = 0; k < toProcess.length; k++) {
    var msg  = toProcess[k];
    var msgId = msg.getId();
    try {
      var subject = msg.getSubject();
      var body    = getMessageBody(msg);
      var emailDate = Utilities.formatDate(msg.getDate(), 'Asia/Kolkata', 'yyyy-MM-dd');

      var parsed = parseDjuboBooking(body, msgId, subject);
      if (parsed) {
        if (!parsed.addedOn) parsed.addedOn = emailDate;
        parsed.emailReceivedDate = emailDate;
        writeToFirestore(parsed);
        added++;
        Logger.log('[Sync] ✅ ' + parsed.reservationId + ' — ' + parsed.guestName + ' — ' + parsed.property);
      } else {
        Logger.log('[Sync] ⚠ could not parse: ' + subject);
      }
      processed.add(msgId);
    } catch(e) {
      Logger.log('[Sync] ✗ error on ' + msgId + ': ' + e.message);
    }
  }

  // ── 2. Cancellation emails ──────────────────────────────────────────────
  var cancelThreads = GmailApp.search('subject:"Reservation Cancellation" after:' + afterDate);
  for (var ci = 0; ci < cancelThreads.length; ci++) {
    var cmsgs = cancelThreads[ci].getMessages();
    for (var cj = 0; cj < cmsgs.length; cj++) {
      var cmsg   = cmsgs[cj];
      var cmsgId = cmsg.getId();
      if (processed.has(cmsgId)) continue;
      try {
        var csubject = cmsg.getSubject();
        var cbody    = getMessageBody(cmsg);
        var resId    = parseDjuboCancellation(cbody, csubject);
        if (resId) {
          writeCancellationToFirestore(resId);
          cancelled++;
          Logger.log('[Sync] ❌ cancellation: ' + resId);
        }
        processed.add(cmsgId);
      } catch(e) {
        Logger.log('[Sync] ✗ cancel error: ' + e.message);
      }
    }
  }

  // ── Save state ──────────────────────────────────────────────────────────
  saveProcessedIds(props, processed);
  props.setProperty(LAST_SYNC_KEY, Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy/MM/dd'));
  Logger.log('[Sync] Done — added:' + added + ' cancelled:' + cancelled);
}

// ─── TRIGGER SETUP (run once manually) ───────────────────────────────────────
function setupTrigger() {
  // Delete existing triggers to avoid duplicates
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'syncDjuboEmails') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger('syncDjuboEmails')
    .timeBased()
    .everyHours(2)
    .create();
  Logger.log('✅ Trigger created — syncDjuboEmails will run every 2 hours');
}

// ─── PARSING HELPERS ─────────────────────────────────────────────────────────
var MONTH_IDX = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
var KNOWN_MEAL_PLANS = ['Room Only','CP - Breakfast','CP – Breakfast','MAP - Breakfast & Dinner','MAP – Breakfast & Dinner','AP - All Inclusive','AP','EP','CP'];
var PROP_FROM_PREFIX = {
  SECP:'Pangot Perch',
  SECW:'Willows',SECT:'Taraview',SECD:'Taradale',SECE:'Cliffs Edge',
  SECM:'Mussoorie',SECA:"Annie's Song",SECS:'Two Stones',SECB:'Bantony Cottage',
  SECOL:'One Love',SECTE:'Tea Estate',SECRC:'Red Cedar Cottage',
  SECFC:'Falling Cashews',SECH:'Hampi 1800',
  SECK:'Karthika Nivas',SECMS:'Marari Sand',SECBH:'Beach House by the Lake',
  SECBV:'Blue Villa',SECSV:'Sunrise Villa',SECGV:'Green Villa'
};
var DJUBO_NAME_MAP = {
  'pangot, perch':'Pangot Perch','pangot perch':'Pangot Perch',
  'willows':'Willows','taraview':'Taraview','taradale':'Taradale','cliffs edge':'Cliffs Edge',
  'mussoorie':'Mussoorie',"annie's song":"Annie's Song",'two stones':'Two Stones','bantony cottage':'Bantony Cottage',
  'one love':'One Love','tea estate':'Tea Estate','red cedar cottage':'Red Cedar Cottage',
  'falling cashews':'Falling Cashews','hampi 1800':'Hampi 1800','hampi':'Hampi 1800',
  'karthika nivas':'Karthika Nivas','marari sand':'Marari Sand',
  'beach house by the lake':'Beach House by the Lake','beach house':'Beach House by the Lake',
  'blue villa':'Blue Villa','sunrise villa':'Sunrise Villa','green villa':'Green Villa'
};

function djuboDateToISO(str) {
  if (!str) return '';
  var m = str.match(/(\d{1,2})\s+(\w{3})[,\s]+(\d{4})/);
  if (!m) return '';
  var mo = MONTH_IDX[m[2]] || 1;
  return m[3] + '-' + (mo < 10 ? '0' : '') + mo + '-' + (m[1].length < 2 ? '0' : '') + m[1];
}
function djuboTime(str) {
  var m = str ? str.match(/(\d{2}:\d{2})/) : null;
  return m ? m[1] : '14:00';
}
function todayISO() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

function getMessageBody(msg) {
  // Try plain text first, then HTML-stripped
  var plain = msg.getPlainBody();
  if (plain && plain.length > 50) return plain;
  var html = msg.getBody();
  return htmlToText(html || '');
}

function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n').replace(/<\/td>/gi, ' ').replace(/<\/th>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}

function parseDjuboBooking(body, msgId, subject) {
  if (!/Reservation\s+Confirmation/i.test(subject)) return null;

  var subjParts = subject.split(/\s*[-–]\s*/);
  var resId = null;
  for (var i = subjParts.length - 1; i >= 0; i--) {
    var m = subjParts[i].match(/([A-Z]{2,}\d{4,})/);
    if (m) { resId = m[1].trim(); break; }
  }
  if (!resId) {
    var bm = body.match(/([A-Z]{2,}\d{4,})/);
    if (bm) resId = bm[1].trim();
  }
  if (!resId) return null;

  // Property from ResID prefix
  var px = resId.match(/^([A-Z]+)/);
  var property = (px && PROP_FROM_PREFIX[px[1]]) || '';
  if (!property) {
    var midIdx = -1;
    for (var j = 0; j < subjParts.length; j++) {
      if (/Reservation\s+Confirmation/i.test(subjParts[j])) { midIdx = j; break; }
    }
    if (midIdx >= 0 && midIdx + 1 < subjParts.length) {
      var raw = subjParts[midIdx + 1].replace(/^Seclude\s*/i, '').trim();
      property = DJUBO_NAME_MAP[raw.toLowerCase()] || raw;
    }
  }

  function grab(re) {
    var m = body.match(re);
    return m ? m[1].replace(/[\r\n\t]+/g, ' ').trim() : '';
  }
  var D = '\\d{1,2}\\s+\\w{3}[\\s,]+\\d{4}[^\\n\\r]*';
  var ciStr  = grab(new RegExp('Check.{0,5}in[\\s\\S]{0,80}?(' + D + ')', 'i'));
  var coStr  = grab(new RegExp('Check.{0,5}out[\\s\\S]{0,80}?(' + D + ')', 'i'));
  var guestName   = grab(/Guest.{0,5}Name.{0,15}:\s*([^\n\r:]{2,60})/i);
  var guestEmail  = grab(/Guest.{0,5}Email.{0,15}:\s*([^\s\n\r]{3,80})/i);
  var guestPhone  = grab(/Guest.{0,10}Contact.{0,20}:\s*([^\n\r]{3,30})/i);
  var createdBy   = grab(/Created.{0,4}By.{0,10}:\s*([^\n\r]{2,60})/i);
  var createdOnStr= grab(/Created.{0,4}On.{0,10}:\s*([^\n\r]{2,40})/i);
  var resType     = grab(/Type.{0,5}Reservation.{0,10}:\s*([^\n\r]{2,40})/i) || 'Direct';

  // Rooms table
  var rooms = [];
  var lines = body.split('\n');
  var tblIdx = -1;
  for (var li = 0; li < lines.length; li++) {
    if (lines[li].indexOf('S.No.') >= 0 && lines[li].indexOf('Room') >= 0) { tblIdx = li; break; }
  }
  if (tblIdx >= 0) {
    for (var ri = tblIdx + 1; ri < lines.length; ri++) {
      var line = lines[ri].trim();
      if (!line || /^Room Rent|^Total\b|^Discount|^Sub Total|^Total Taxes|^Advance|^Total outstanding/.test(line)) break;
      if (!/^\d+\s/.test(line)) continue;
      var rest = line.replace(/^\d+\s+/, '');
      var mealPlan = 'Room Only';
      for (var mi = 0; mi < KNOWN_MEAL_PLANS.length; mi++) {
        if (rest.indexOf(KNOWN_MEAL_PLANS[mi]) >= 0) { mealPlan = KNOWN_MEAL_PLANS[mi]; break; }
      }
      var mpIdx  = rest.indexOf(mealPlan);
      var before = rest.slice(0, mpIdx).trim();
      var after  = rest.slice(mpIdx + mealPlan.length).trim();
      var bp     = before.split(/\s+/);
      var priceM = after.match(/INR\s*([\d,]+\.?\d*)/);
      var nums   = bp.slice(-4);
      rooms.push({
        roomName: bp.slice(0, -4).join(' ') || rest.split(/\s+/)[0],
        nights:   parseInt(nums[0]) || 1,
        adults:   parseInt(nums[1]) || 2,
        children: parseInt(nums[2]) || 0,
        infants:  parseInt(nums[3]) || 0,
        mealPlan: mealPlan,
        netPrice: priceM ? parseFloat(priceM[1].replace(/,/g, '')) : 0
      });
    }
  }
  if (!rooms.length) rooms.push({roomName:'',nights:1,adults:2,children:0,infants:0,mealPlan:'Room Only',netPrice:0});

  function fin(re) {
    var m = body.match(re);
    return m ? parseFloat(m[1].replace(/,/g, '')) : 0;
  }
  var financials = {
    roomRent:    fin(/Room\s*Rent[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i),
    discount:    String(fin(/Discount[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i) || ''),
    subTotal:    fin(/Sub\s*Total[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i),
    taxes:       String(fin(/Total\s*Tax[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i) || ''),
    advance:     String(fin(/Advance[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i) || ''),
    outstanding: fin(/outstanding[\s\S]{0,20}?INR\s*([\d,]+\.?\d*)/i)
  };

  return {
    id: 'djubo_' + resId,
    reservationId: resId,
    property: property,
    guestName: guestName,
    guestEmail: guestEmail,
    guestPhone: guestPhone,
    checkIn:        djuboDateToISO(ciStr),
    checkInTime:    djuboTime(ciStr),
    checkOut:       djuboDateToISO(coStr),
    checkOutTime:   djuboTime(coStr) || '11:00',
    reservationType: resType,
    status: 'Confirmed',
    upsellStatus: 'Pending',
    rooms: rooms,
    financials: financials,
    callLog: [],
    pipelineStages: {},
    createdBy: createdBy,
    addedById: null,
    addedOn: djuboDateToISO(createdOnStr) || todayISO(),
    source: 'gas_auto',
    gmailMsgId: msgId
  };
}

function parseDjuboCancellation(body, subject) {
  if (!/Reservation\s+Cancellation/i.test(subject)) return null;
  var subjParts = subject.split(/\s*[-–]\s*/);
  for (var i = subjParts.length - 1; i >= 0; i--) {
    var m = subjParts[i].match(/([A-Z]{2,}\d{4,})/);
    if (m) return m[1];
  }
  var bm = body.match(/Reservation ID[\s\S]{0,15}\n\s*([A-Z]{2,}\d{4,})/);
  return bm ? bm[1] : (body.match(/\b([A-Z]{2,}\d{4,})\b/) || [null,null])[1];
}

// ─── FIRESTORE REST ──────────────────────────────────────────────────────────
function jsToFirestore(val) {
  if (val === null || val === undefined) return {nullValue: null};
  if (typeof val === 'boolean') return {booleanValue: val};
  if (typeof val === 'number')  return Number.isInteger(val) ? {integerValue: String(val)} : {doubleValue: val};
  if (typeof val === 'string')  return {stringValue: val};
  if (Array.isArray(val))       return {arrayValue: {values: val.map(jsToFirestore)}};
  if (typeof val === 'object') {
    var fields = {};
    for (var k in val) { if (val.hasOwnProperty(k)) fields[k] = jsToFirestore(val[k]); }
    return {mapValue: {fields: fields}};
  }
  return {stringValue: String(val)};
}
function bookingToFirestoreFields(booking) {
  var fields = {};
  for (var k in booking) {
    if (booking.hasOwnProperty(k)) fields[k] = jsToFirestore(booking[k]);
  }
  return fields;
}

function writeToFirestore(booking) {
  var docId  = booking.reservationId;
  var url    = FIRESTORE_BASE + '/' + PENDING_COLLECTION + '/' + docId + '?key=' + FIRESTORE_API_KEY;
  var body   = JSON.stringify({fields: bookingToFirestoreFields(booking)});
  var resp   = UrlFetchApp.fetch(url, {method:'PATCH', contentType:'application/json', payload:body, muteHttpExceptions:true});
  if (resp.getResponseCode() >= 300) {
    Logger.log('[Firestore] write error ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0,200));
  }
}

function writeCancellationToFirestore(resId) {
  var docId = 'cancel_' + resId;
  var url   = FIRESTORE_BASE + '/' + PENDING_COLLECTION + '/' + docId + '?key=' + FIRESTORE_API_KEY;
  var body  = JSON.stringify({fields:{
    _type:    {stringValue:'cancellation'},
    reservationId: {stringValue: resId},
    cancelledAt:   {stringValue: todayISO()}
  }});
  UrlFetchApp.fetch(url, {method:'PATCH', contentType:'application/json', payload:body, muteHttpExceptions:true});
}

// ─── PROCESSED IDs ────────────────────────────────────────────────────────────
function loadProcessedIds(props) {
  try {
    var raw = props.getProperty(PROCESSED_KEY);
    var arr = raw ? JSON.parse(raw) : [];
    var set = {};
    for (var i = 0; i < arr.length; i++) set[arr[i]] = true;
    return {
      has: function(id) { return !!set[id]; },
      add: function(id) { set[id] = true; },
      toArray: function() { return Object.keys(set); }
    };
  } catch(e) { return {has:function(){return false;},add:function(){},toArray:function(){return [];}}; }
}
function saveProcessedIds(props, processed) {
  try {
    var arr = processed.toArray();
    if (arr.length > MAX_STORED_IDS) arr = arr.slice(arr.length - MAX_STORED_IDS);
    props.setProperty(PROCESSED_KEY, JSON.stringify(arr));
  } catch(e) { Logger.log('[Sync] could not save processed IDs: ' + e.message); }
}

// ─── DATE HELPERS ─────────────────────────────────────────────────────────────
function gmailAfterDate(lastSyncStr) {
  // lastSyncStr format: yyyy/MM/dd — go back 1 day for safety overlap
  try {
    var parts = lastSyncStr.split('/');
    var d = new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
    d.setDate(d.getDate() - 1);
    return d.getFullYear() + '/' +
      (d.getMonth()+1 < 10 ? '0' : '') + (d.getMonth()+1) + '/' +
      (d.getDate() < 10 ? '0' : '') + d.getDate();
  } catch(e) { return '2026/06/01'; }
}

// ─── MANUAL TEST ──────────────────────────────────────────────────────────────
// Run this in the GAS editor to test without waiting for the trigger
function testSync() {
  syncDjuboEmails();
}
