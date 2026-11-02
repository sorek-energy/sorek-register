/* Sorek Register - Send reminder (Smart Alerts).
   Runs when the user clicks Send. Only NEW emails (not replies/forwards) to a project-party domain
   without a SOR number get a reminder. The user can always choose "Send Anyway". */
/* global Office */

var SOREK_DEFAULT_NAMES = { GOV: "the State of Israel", FIC: "Fichtner", SNR: "Sener", SIE: "Siemens Energy" };

var SOREK_DEFAULT_DOMAINS = {
  GOV: ["gov.il"],
  FIC: ["fichtner.de"],
  SNR: ["corporate.sener", "energy.sener"],
  SIE: ["siemens-energy.com"]
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

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
}
