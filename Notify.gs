// Notify.gs - phone notifications through OneSignal. A separate file in the main Apps Script project.
// It reads the same data as the app and sends a short message to each person's phone.
//
// Script Properties (Project Settings) needed:
//   ONESIGNAL_APP_ID    from OneSignal > Settings > Keys & IDs
//   ONESIGNAL_REST_KEY  the REST API Key from the same page (keep it private)
//
// Optional rows in the Settings tab of your Sheet:
//   NotifyDetail        "full" shows the details in the message. Anything else keeps it generic (default).
//   NotifyPeople        names to notify, separated by commas (default: everyone on the People tab)
//   NotifyMorningHour   hour for the morning message, 0-23 (default 7)
//   NotifyEveningHour   hour for the evening message, 0-23 (default 18)
//   AppUrl              the address the notification opens

// Falls back to this household's own web app URL (set per-Sheet via the
// Settings tab's AppUrl key, or the multi-tenant web app URL itself).
function notifyAppUrl_() { return getSetting_('AppUrl') || getWebAppUrl(); }

function notifyProp_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

// Sends one push message to the given people (their OneSignal external IDs are their names)
function oneSignalSend_(names, title, body) {
  var appId = notifyProp_('ONESIGNAL_APP_ID'), key = notifyProp_('ONESIGNAL_REST_KEY');
  if (!appId || !key) throw new Error('Add ONESIGNAL_APP_ID and ONESIGNAL_REST_KEY in Project Settings first.');
  var res = UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Key ' + key },
    payload: JSON.stringify({
      app_id: appId, target_channel: 'push', include_aliases: { external_id: names },
      headings: { en: title }, contents: { en: body }, url: notifyAppUrl_()
    })
  });
  var code = res.getResponseCode(), data = {};
  try { data = JSON.parse(res.getContentText()); } catch (e) {}
  if (code < 200 || code >= 300) throw new Error('OneSignal error ' + code + ': ' + String(res.getContentText()).slice(0, 200));
  return data; // data.id is empty when nobody has turned notifications on yet
}

function notifyRecipients_() {
  var all = rows_('People').map(function (r) { return String(r[0]).trim(); }).filter(Boolean);
  var setting = getSetting_('NotifyPeople');
  if (!setting) return all;
  var wanted = setting.split(',').map(function (x) { return x.trim().toLowerCase(); });
  return all.filter(function (n) { return wanted.indexOf(n.toLowerCase()) >= 0; });
}

function dueWords_(daysLeft) {
  return daysLeft < 0 ? 'overdue' : daysLeft === 0 ? 'due today' : daysLeft === 1 ? 'due tomorrow' : 'due in ' + daysLeft + ' days';
}

// What matters today, for one person
function todayItems_(home, docs, person) {
  var items = [];
  (home.events || []).forEach(function (e) { items.push((e.time === 'All day' ? 'All day' : e.time) + ' ' + e.title); });
  if (home.nextClinic && home.nextClinic.daysLeft === 0) items.push('Baby clinic: ' + home.nextClinic.what);
  (home.bills || []).forEach(function (b) { if (!b.paid && b.daysLeft <= 2) items.push(b.name + ' bill ' + dueWords_(b.daysLeft)); });
  (home.reminders || []).forEach(function (r) { if (r.daysLeft <= 0) items.push(r.name + ' ' + dueWords_(r.daysLeft)); });
  ((docs && docs.soon) || []).forEach(function (d) { if (d.daysLeft <= 0) items.push(d.title + (d.daysLeft < 0 ? ' has expired' : ' expires today')); });
  var me = String(person).toLowerCase();
  (home.chores || []).forEach(function (c) {
    if (!c.done && c.status === 'due' && String(c.who || '').trim().toLowerCase() === me) items.push('Chore: ' + c.name);
  });
  return items;
}

function digestMessage_(prefix, items) {
  var n = items.length;
  var title = prefix + ': ' + (n === 1 ? '1 thing' : n + ' things');
  if (getSetting_('NotifyDetail').toLowerCase() !== 'full') return { title: title, body: 'Open the app to see what needs you.' };
  var lines = items.slice(0, 5);
  if (items.length > 5) lines.push('and ' + (items.length - 5) + ' more');
  return { title: title, body: lines.join('\n') };
}

// Runs every morning. Sends nothing to people who have nothing today.
function sendMorningDigest() {
  var home = getHomeData(), docs = getDocsSummary();
  notifyRecipients_().forEach(function (person) {
    var items = todayItems_(home, docs, person);
    if (!items.length) return;
    var m = digestMessage_('Today', items);
    oneSignalSend_([person], m.title, m.body);
  });
}

// Runs every evening: what is coming tomorrow
function sendEveningPreview() {
  var tomorrow = dayStr_(new Date(Date.now() + 86400000));
  var items = [];
  try {
    getCalendarData().events.forEach(function (e) {
      if (dayStr_(new Date(e.start)) === tomorrow) items.push((e.allDay ? 'All day' : Utilities.formatDate(new Date(e.start), tz_(), 'HH:mm')) + ' ' + e.title);
    });
  } catch (err) { /* calendar not available */ }
  getRemindersData().items.forEach(function (r) { if (!r.done && r.daysLeft === 1) items.push(r.name + ' due tomorrow'); });
  var home = getHomeData();
  (home.bills || []).forEach(function (b) { if (!b.paid && b.daysLeft === 1) items.push(b.name + ' bill due tomorrow'); });
  if (home.nextClinic && home.nextClinic.daysLeft === 1) items.push('Baby clinic: ' + home.nextClinic.what);
  if (!items.length) return;
  var m = digestMessage_('Tomorrow', items);
  notifyRecipients_().forEach(function (person) { oneSignalSend_([person], m.title, m.body); });
}

// Run this by hand to check that everything is connected
function sendTestNotification() {
  var names = notifyRecipients_();
  var data = oneSignalSend_(names, 'Household test', 'If you can read this, notifications are working.');
  Logger.log(data.id ? 'Sent to: ' + names.join(', ') : 'Nobody has turned notifications on yet. Open the app and tap Turn on.');
  return data;
}

// Run this by hand once. It sets up the daily morning and evening messages.
function installNotificationTriggers() {
  removeNotificationTriggers();
  var morning = Math.min(23, Math.max(0, parseInt(getSetting_('NotifyMorningHour'), 10) || 7));
  var evening = Math.min(23, Math.max(0, parseInt(getSetting_('NotifyEveningHour'), 10) || 18));
  ScriptApp.newTrigger('sendMorningDigest').timeBased().everyDays(1).atHour(morning).inTimezone(tz_()).create();
  ScriptApp.newTrigger('sendEveningPreview').timeBased().everyDays(1).atHour(evening).inTimezone(tz_()).create();
  Logger.log('Morning message at ' + morning + ':00 and evening message at ' + evening + ':00 (' + tz_() + ').');
}

function removeNotificationTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'sendMorningDigest' || fn === 'sendEveningPreview') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- Weekly digest email ---------- */
// A Sunday-morning email (not a push notification, so it works even if
// nobody has turned push notifications on) summarising the week ahead:
// bills due, meals planned, chores waiting, reminders coming up, documents
// expiring, and calendar events. Sent to every email on the People tab.

function weekRangeLabel_() {
  var start = new Date(), end = new Date(start.getTime() + 6 * 86400000);
  return Utilities.formatDate(start, tz_(), 'd MMM') + '–' + Utilities.formatDate(end, tz_(), 'd MMM');
}

function buildWeeklyDigest_() {
  var home = getHomeData();
  var docs = {}; try { docs = getDocsSummary(); } catch (err) { /* Documents tab not set up */ }
  var meals = getMealsData();
  var cal = { events: [] }; try { cal = getCalendarData(); } catch (err) { /* calendar not found */ }
  var weekEnd = dayStr_(new Date(Date.now() + 7 * 86400000));
  var sections = [];

  var billLines = (home.bills || []).filter(function (b) { return !b.paid && b.daysLeft >= 0 && b.daysLeft <= 7; })
    .sort(function (a, b) { return a.daysLeft - b.daysLeft; })
    .map(function (b) { return b.name + ' — ' + dueWords_(b.daysLeft); });
  if (billLines.length) sections.push({ title: 'Bills due this week', lines: billLines });

  var mealLines = (meals.items || []).filter(function (m) { return m.date <= weekEnd; })
    .map(function (m) { return Utilities.formatDate(Utilities.parseDate(m.date, tz_(), 'yyyy-MM-dd'), tz_(), 'EEE') + ': ' + m.meal; });
  if (mealLines.length) sections.push({ title: 'Meals planned', lines: mealLines });

  var choreLines = (home.chores || []).filter(function (c) { return c.status === 'due'; })
    .map(function (c) { return c.name + (c.who ? ' (' + c.who + ')' : ''); });
  if (choreLines.length) sections.push({ title: 'Chores waiting', lines: choreLines });

  var reminderLines = (home.reminders || []).filter(function (r) { return r.daysLeft >= 0 && r.daysLeft <= 7; })
    .sort(function (a, b) { return a.daysLeft - b.daysLeft; })
    .map(function (r) { return r.name + ' — ' + dueWords_(r.daysLeft); });
  if (reminderLines.length) sections.push({ title: 'Reminders coming up', lines: reminderLines });

  var docLines = ((docs && docs.soon) || []).filter(function (d) { return d.daysLeft >= 0 && d.daysLeft <= 7; })
    .map(function (d) { return d.title + (d.daysLeft === 0 ? ' expires today' : ' expires in ' + d.daysLeft + ' days'); });
  if (docLines.length) sections.push({ title: 'Documents expiring', lines: docLines });

  var eventLines = (cal.events || []).filter(function (e) { return dayStr_(new Date(e.start)) <= weekEnd; })
    .map(function (e) {
      return Utilities.formatDate(new Date(e.start), tz_(), 'EEE d MMM') +
        (e.allDay ? ' (all day)' : ' ' + Utilities.formatDate(new Date(e.start), tz_(), 'HH:mm')) + ': ' + e.title;
    });
  if (eventLines.length) sections.push({ title: 'On the calendar', lines: eventLines });

  return sections;
}

function weeklyDigestHtml_(sections) {
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  var body = sections.map(function (s) {
    var items = s.lines.map(function (l) { return '<li style="margin:0 0 6px;">' + esc(l) + '</li>'; }).join('');
    return '<h2 style="font-family:Georgia,serif;font-size:17px;color:#2B3528;margin:22px 0 8px;">' + esc(s.title) + '</h2>' +
      '<ul style="margin:0;padding-left:20px;color:#2B3528;font-size:14px;line-height:1.5;">' + items + '</ul>';
  }).join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:480px;margin:0 auto;background:#FFFDF8;padding:24px;border-radius:12px;border:1px solid #E1DAC7;">' +
    '<p style="font-family:Georgia,serif;font-weight:bold;font-size:20px;color:#2B3528;margin:0 0 4px;">PorterPilot</p>' +
    '<p style="color:#74806E;font-size:13px;margin:0 0 4px;">Your week ahead, ' + weekRangeLabel_() + '</p>' +
    body +
    '<p style="margin-top:24px;"><a href="' + notifyAppUrl_() + '" style="background:#55684B;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none;font-size:14px;">Open PorterPilot</a></p>' +
    '</div>';
}

function weeklyDigestText_(sections) {
  return sections.map(function (s) { return s.title + ':\n' + s.lines.map(function (l) { return '- ' + l; }).join('\n'); }).join('\n\n');
}

// Runs once a week. Sends nothing when there is nothing due, and nothing to
// a member with no email on the People tab.
function sendWeeklyDigestEmail() {
  var sections = buildWeeklyDigest_();
  if (!sections.length) return;
  var subject = 'PorterPilot: your week ahead (' + weekRangeLabel_() + ')';
  var html = weeklyDigestHtml_(sections);
  var text = weeklyDigestText_(sections);
  var emails = rows_('People').map(function (r) { return String(r[1] || '').trim(); }).filter(function (e) { return /@/.test(e); });
  var uniq = emails.filter(function (e, i) { return emails.indexOf(e) === i; });
  uniq.forEach(function (email) {
    try { MailApp.sendEmail({ to: email, subject: subject, body: text, htmlBody: html }); } catch (err) { /* keep going for the rest */ }
  });
}

// Run this by hand once (e.g. from the Apps Script editor). Add a
// NotifyWeeklyHour row (0-23) in the Settings tab to change the send time.
function installWeeklyDigestTrigger() {
  removeWeeklyDigestTrigger();
  var hour = Math.min(23, Math.max(0, parseInt(getSetting_('NotifyWeeklyHour'), 10) || 8));
  ScriptApp.newTrigger('sendWeeklyDigestEmail').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(hour).inTimezone(tz_()).create();
  Logger.log('Weekly digest email at ' + hour + ':00 on Sundays (' + tz_() + ').');
}

function removeWeeklyDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyDigestEmail') ScriptApp.deleteTrigger(t);
  });
}

// Run this by hand to check the email looks right, without waiting for Sunday.
function sendTestWeeklyDigest() {
  var sections = buildWeeklyDigest_();
  if (!sections.length) { Logger.log('Nothing due in the next 7 days, so no email would be sent.'); return; }
  var me = Session.getActiveUser().getEmail();
  MailApp.sendEmail({ to: me, subject: 'PorterPilot: your week ahead (' + weekRangeLabel_() + ') [TEST]', body: weeklyDigestText_(sections), htmlBody: weeklyDigestHtml_(sections) });
  Logger.log('Test email sent to ' + me);
}
