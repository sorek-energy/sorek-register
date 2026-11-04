/* Sorek Register - Send reminder (Smart Alerts).
   Runs when the user clicks Send. Only NEW emails (not replies/forwards) to a project-party domain
   without a SOR number get a reminder. The user can always choose "Send Anyway". */
/* global Office */

var SOREK_DEFAULT_NAMES = { GOV: "State of Israel", FIC: "Fichtner", SNR: "SENER", SIE: "SIEMENS", ISK: "Isa Khoury", GAL: "Galil", MLK: "MELKAM" };

var SOREK_DEFAULT_DOMAINS = {
  GOV: ["gov.il"],
  FIC: ["fichtner.de"],
  SNR: ["corporate.sener", "energy.sener"],
  SIE: ["siemens-energy.com"],
  ISK: ["ikgenergy.com", "tomicic.com", "isakhoury.com", "trisak-en.com"],
  GAL: ["galileng.com"],
  MLK: ["melkam.co.il"]
};

function sorekAllow(event) { event.completed({ allowEvent: true }); }

function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox.item;
    item.subject.getAsync(function (sr) {
      var subject = (sr && sr.value) || "";
      if (/SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/i.test(subject)) { sorekAllow(event); return; }
      if (/^\s*(re|fw|fwd|aw|wg)\s*:/i.test(subject) || /^\s*(תשובה|השב|הועבר)\s*:/.test(subject)) { sorekAllow(event); return; }

      var rs = null;
      try { rs = Office.context.roamingSettings; } catch (e) { rs = null; }
      if (rs && rs.get("sorReminder") === false) { sorekAllow(event); return; }

      var map = null;
      try { map = rs ? JSON.parse(rs.get("sorDomains") || "null") : null; } catch (e) { map = null; }
      if (!map) map = SOREK_DEFAULT_DOMAINS;

      item.to.getAsync(function (tr) {
        item.cc.getAsync(function (cr) {
          var all = [].concat((tr && tr.value) || [], (cr && cr.value) || []);
          var hit = null;
          for (var i = 0; i < all.length && !hit; i++) {
            var addr = (all[i].emailAddress || "").toLowerCase();
            var d = addr.split("@")[1];
            if (!d) continue;
            for (var code in map) {
              var doms = map[code] || [];
              for (var j = 0; j < doms.length; j++) {
                var x = String(doms[j]).toLowerCase();
                if (x && (d === x || d.slice(-(x.length + 1)) === "." + x)) { hit = code; break; }
              }
              if (hit) break;
            }
          }
          if (!hit) { sorekAllow(event); return; }

          var names = null;
          try { names = rs ? JSON.parse(rs.get("sorNames") || "null") : null; } catch (e) { names = null; }
          var who = (names && names[hit]) || SOREK_DEFAULT_NAMES[hit] || hit;
          var msg = "This email is going to " + who + " (" + hit + ").\n" +
                    "New project emails get a SOR number for the register.\n\n" +
                    "Get number: add one now.\nSend Anyway: send it without a number.";
          var opts = { allowEvent: false, errorMessage: msg };
          try {
            if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
              opts.errorMessageMarkdown = "**This email is going to " + who + " (" + hit + ").**\n\n" +
                "New project emails get a SOR number for the register.\n\n" +
                "- **Get number** – add one now\n- **Send Anyway** – send it without a number";
            }
          } catch (e) { /* plain text only */ }
          try {
            if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
              opts.cancelLabel = "Get number";
              opts.commandId = "msgComposeOpenPaneButton";
              opts.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
            }
          } catch (e) { /* older Outlook: default Send Anyway prompt */ }
          event.completed(opts);
        });
      });
    });
  } catch (e) {
    // Never block sending because of the helper
    sorekAllow(event);
  }
}

/* Project-email bar: shown while writing a NEW email to a project party, with a link that opens the pane. */
function sorekProjectHit(item, rs, done) {
  var map = null;
  try { map = rs ? JSON.parse(rs.get("sorDomains") || "null") : null; } catch (e) { map = null; }
  if (!map) map = SOREK_DEFAULT_DOMAINS;
  item.to.getAsync(function (tr) {
    item.cc.getAsync(function (cr) {
      var all = [].concat((tr && tr.value) || [], (cr && cr.value) || []);
      for (var i = 0; i < all.length; i++) {
        var d = String(all[i].emailAddress || "").toLowerCase().split("@")[1];
        if (!d) continue;
        for (var code in map) {
          var doms = map[code] || [];
          for (var j = 0; j < doms.length; j++) {
            var x = String(doms[j]).toLowerCase();
            if (x && (d === x || d.slice(-(x.length + 1)) === "." + x)) { done(code); return; }
          }
        }
      }
      done(null);
    });
  });
}

function onRecipientsChangedHandler(event) {
  var finish = function () { try { event.completed(); } catch (e) { /* ignore */ } };
  try {
    var item = Office.context.mailbox.item;
    var rs = null;
    try { rs = Office.context.roamingSettings; } catch (e) { rs = null; }
    var clear = function () { item.notificationMessages.removeAsync("sorekHint", function () { finish(); }); };
    if (rs && rs.get("sorBar") === false) { clear(); return; }
    item.subject.getAsync(function (sr) {
      var subject = (sr && sr.value) || "";
      if (/SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/i.test(subject) || /^\s*(re|fw|fwd|aw|wg)\s*:/i.test(subject) ||
          /^\s*(\u05EA\u05E9\u05D5\u05D1\u05D4|\u05D4\u05E9\u05D1|\u05D4\u05D5\u05E2\u05D1\u05E8)\s*:/.test(subject)) { clear(); return; }
      sorekProjectHit(item, rs, function (hit) {
        if (!hit) { clear(); return; }
        var names = null;
        try { names = rs ? JSON.parse(rs.get("sorNames") || "null") : null; } catch (e) { names = null; }
        var who = (names && names[hit]) || SOREK_DEFAULT_NAMES[hit] || hit;
        item.notificationMessages.replaceAsync("sorekHint", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
          message: ("Project email to " + who + " (" + hit + "). Give it a SOR number for the register.").slice(0, 150),
          icon: "Icon.16",
          actions: [{
            actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
            actionText: "Get SOR number",
            commandId: "msgComposeOpenPaneButton",
            contextData: "{}"
          }]
        }, function () { finish(); });
      });
    });
  } catch (e) { finish(); }
}

/* Reply / Reply All / Forward to an email with a SOR number: add References + Subject at the top automatically.
   No new number is taken. Turn off in the pane: Settings > Add the header to replies and forwards automatically. */
var SOREK_REPLY_RE = /^\s*(re|fw|fwd|aw|wg|תשובה|השב|הועבר)\s*:/i;
function sorekEsc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function sorekTopic(subject) {
  var s = String(subject || "");
  for (var i = 0; i < 5; i++) s = s.replace(SOREK_REPLY_RE, "").trim();
  s = s.replace(/SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/gi, " ");
  for (var j = 0; j < 5; j++) s = s.replace(SOREK_REPLY_RE, "").trim();
  return s.replace(/\s{2,}/g, " ").replace(/^[\s\-_–:]+|[\s\-_–:]+$/g, "").trim();
}
function sorekHeaderHtml(refs, topic) {
  var base = "mso-line-height-rule:exactly;line-height:13pt;font-family:Calibri,Arial,sans-serif;font-size:10.5pt;vertical-align:top;";
  var td1 = 'style="' + base + 'padding:0 12px 0 0;color:#16324F;font-weight:bold;white-space:nowrap"';
  var td2 = 'style="' + base + 'padding:0"';
  return '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:0">' +
    '<tr><td ' + td1 + '>References:</td><td ' + td2 + '>' + refs.map(function (r) { return sorekEsc(r) + (topic ? " – " + sorekEsc(topic) : ""); }).join("<br>") + '</td></tr>' +
    '<tr><td ' + td1 + '>Subject:</td><td ' + td2 + '>' + sorekEsc(topic) + '</td></tr></table><div style="line-height:10pt">&nbsp;</div>';
}
function onNewComposeHandler(event) {
  var finish = function () { try { event.completed(); } catch (e) { /* ignore */ } };
  try {
    var item = Office.context.mailbox.item;
    var rs = null;
    try { rs = Office.context.roamingSettings; } catch (e) { rs = null; }
    if (rs && rs.get("sorAutoHdr") === false) { finish(); return; }
    item.subject.getAsync(function (sr) {
      var subject = (sr && sr.value) || "";
      if (!SOREK_REPLY_RE.test(subject)) { finish(); return; }
      var refs = [], m, re = /SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/gi;
      while ((m = re.exec(subject))) { var r = m[0].toUpperCase(); if (refs.indexOf(r) < 0) refs.push(r); }
      if (!refs.length) { finish(); return; }
      var topic = sorekTopic(subject);
      item.body.getTypeAsync(function (tr) {
        var html = !(tr && tr.status === Office.AsyncResultStatus.Succeeded && tr.value === Office.CoercionType.Text);
        var content = html ? sorekHeaderHtml(refs, topic)
          : "References: " + refs.map(function (r) { return r + (topic ? " – " + topic : ""); }).join("; ") + "\nSubject: " + topic + "\n\n";
        item.body.prependAsync(content, { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text }, function () { finish(); });
      });
    });
  } catch (e) { finish(); }
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
  Office.actions.associate("onNewComposeHandler", onNewComposeHandler);
}
