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
