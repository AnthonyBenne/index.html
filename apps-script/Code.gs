/**
 * Ben Anthony website scheduler.
 * GET  -> open 30-minute consultation slots (Eastern time)
 * POST -> books a slot on the default calendar and emails the invitation
 */
var TZ = 'America/New_York';
// Booking windows by weekday (0 = Sunday): [startHour, endHour)
var WINDOWS = {0:[12,18], 1:[9,18], 2:[9,18], 3:[9,18], 4:[9,18], 5:[9,18], 6:[15,18]};
var SLOT_MIN = 30;
var NOTICE_HOURS = 4;
var DAYS_AHEAD = 21;
// CRM sheet "Ben Anthony CRM - Leads" (Sheet1, columns A-J)
var CRM_ID = '1l__KNOxcnC7nzACHRVuEJ_GALETw1GelOP-zdg6OO8c';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    return json_({ok: true, zone: 'Eastern Time', days: getDays_()});
  } catch (err) {
    return json_({ok: false, error: 'unavailable'});
  }
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }

function getDays_() {
  var cal = CalendarApp.getDefaultCalendar();
  var now = new Date();
  var earliest = new Date(now.getTime() + NOTICE_HOURS * 3600000);
  var days = [];
  for (var i = 0; i <= DAYS_AHEAD; i++) {
    var base = new Date(now.getTime() + i * 86400000);
    var ymd = Utilities.formatDate(base, TZ, 'yyyy-MM-dd');
    var noon = new Date(ymd + 'T12:00:00Z');
    var offset = Utilities.formatDate(noon, TZ, 'XXX');
    var dow = parseInt(Utilities.formatDate(noon, TZ, 'u'), 10) % 7;
    var win = WINDOWS[dow];
    if (!win) continue;
    var dayStart = new Date(ymd + 'T' + pad_(win[0]) + ':00:00' + offset);
    var dayEnd = new Date(ymd + 'T' + pad_(win[1]) + ':00:00' + offset);
    var busy = cal.getEvents(dayStart, dayEnd).filter(function (ev) {
      return !ev.isAllDayEvent();
    }).map(function (ev) {
      return [ev.getStartTime().getTime(), ev.getEndTime().getTime()];
    });
    var slots = [];
    for (var t = dayStart.getTime(); t + SLOT_MIN * 60000 <= dayEnd.getTime(); t += SLOT_MIN * 60000) {
      if (t < earliest.getTime()) continue;
      var end = t + SLOT_MIN * 60000;
      var clash = busy.some(function (b) { return t < b[1] && end > b[0]; });
      if (clash) continue;
      var d = new Date(t);
      slots.push({
        iso: Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
        label: Utilities.formatDate(d, TZ, 'h:mm a'),
        hour: parseInt(Utilities.formatDate(d, TZ, 'H'), 10)
      });
    }
    if (slots.length) {
      days.push({
        date: ymd,
        weekday: Utilities.formatDate(noon, TZ, 'EEE'),
        day: Utilities.formatDate(noon, TZ, 'd'),
        month: Utilities.formatDate(noon, TZ, 'MMM'),
        long: Utilities.formatDate(noon, TZ, 'EEEE, MMMM d'),
        slots: slots
      });
    }
  }
  return days;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f<>]/g, ' ').trim().slice(0, max);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    if (clean_(data.website, 50)) return json_({ok: true}); // hidden field: bots only
    var first = clean_(data.first, 60), last = clean_(data.last, 60);
    var phone = clean_(data.phone, 30), email = clean_(data.email, 120);
    var market = clean_(data.market, 20), iso = clean_(data.start, 40);
    if (!first || !last) return json_({ok: false, error: 'name'});
    if (phone.replace(/\D/g, '').length < 10) return json_({ok: false, error: 'phone'});
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ok: false, error: 'email'});
    if (market !== 'Connecticut' && market !== 'New Jersey') market = '';
    var method = clean_(data.method, 10) === 'video' ? 'video' : 'phone';

    lock.waitLock(20000);
    var match = null;
    getDays_().forEach(function (day) {
      day.slots.forEach(function (s) { if (s.iso === iso) match = {day: day, slot: s}; });
    });
    if (!match) return json_({ok: false, error: 'taken'});

    var start = new Date(iso);
    var end = new Date(start.getTime() + SLOT_MIN * 60000);
    var title = (method === 'video' ? 'Video call with ' : 'Call with ') + first + ' ' + last + (market ? ' - ' + market : '');
    var lines = [
      '30-minute consultation booked on benanthony.com',
      'Format: ' + (method === 'video' ? 'Video call (Google Meet)' : 'Phone call'),
      'Phone: ' + phone,
      'Email: ' + email
    ];
    if (market) lines.push('Market: ' + market);
    lines.push('');
    lines.push(method === 'video'
      ? 'Join with the Google Meet link in this invitation.'
      : 'This is a phone call. Ben will call you at the number above.');
    var description = lines.join('\n');
    var fmt = "yyyy-MM-dd'T'HH:mm:ssXXX";
    var done = false;
    try {
      // Calendar API: lets a phone booking carry no video link and a video booking carry one
      var resource = {
        summary: title,
        description: description,
        start: {dateTime: Utilities.formatDate(start, TZ, fmt), timeZone: TZ},
        end: {dateTime: Utilities.formatDate(end, TZ, fmt), timeZone: TZ},
        attendees: [{email: email}]
      };
      if (method === 'video') {
        resource.conferenceData = {createRequest: {requestId: Utilities.getUuid(), conferenceSolutionKey: {type: 'hangoutsMeet'}}};
      }
      var created = Calendar.Events.insert(resource, 'primary', {conferenceDataVersion: 1, sendUpdates: 'all'});
      done = true;
      if (method === 'phone' && created.conferenceData) {
        try {
          Calendar.Events.patch({conferenceData: null}, 'primary', created.id, {conferenceDataVersion: 1, sendUpdates: 'none'});
        } catch (x) {}
      }
    } catch (apiErr) {
      if (done) throw apiErr;
      CalendarApp.getDefaultCalendar().createEvent(title, start, end, {
        description: description,
        guests: email,
        sendInvites: true
      });
    }
    try {
      // Timestamp | Name | Phone | Email | Market | Reason | Source | Appointment | Stage | Notes
      SpreadsheetApp.openById(CRM_ID).getSheets()[0].appendRow([
        new Date(), first + ' ' + last, phone, email, market, '', 'Website',
        match.day.long + ' at ' + match.slot.label + ' ET', 'New',
        method === 'video' ? 'Video call (Google Meet)' : 'Phone call'
      ]);
    } catch (crmErr) { console.error('CRM write failed: ' + crmErr); }
    return json_({ok: true, when: match.day.long + ' at ' + match.slot.label, method: method});
  } catch (err) {
    return json_({ok: false, error: 'server'});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
