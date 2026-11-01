/* Sorek Register - Send reminder (Smart Alerts).
   Runs when the user clicks Send. Only NEW emails (not replies/forwards) to a project-party domain
   without a SOR number get a reminder. The user can always choose "Send Anyway". */
/* global Office */

var SOREK_DEFAULT_DOMAINS = {
  GOV: ["gov.il"],
  FIC: ["fichtner.de"],
  SNR: ["sener.es", "group.sener"],
  SIE: ["siemens.com", "siemens-energy.com"]
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

          var msg = "New email to a project party (" + hit + "). Give it a SOR number? " +
                    "Open the Sorek Register pane to number it, or choose Send Anyway to send without a number.";
          var opts = { allowEvent: false, errorMessage: msg };
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

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
