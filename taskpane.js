/* Sorek Register - Outlook task pane.
   Compose: give a new email a SOR number (central counter in SharePoint list "Companies"),
            write it to the subject, add the general mailbox to CC, log it in "EmailRegister".
   Read:    find the email in the register (or log it) and set Status / Variation / Impact / Claims notes. */
/* global Office, msal */
(function () {
  "use strict";
  var C = window.SOREK_CONFIG;
  var SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All"];
  var TASK_SCOPES = ["https://graph.microsoft.com/Tasks.ReadWrite"];
  var NUM_RE = /SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/i;
  var NUM_RE_G = /SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/gi;
  var REPLY_RE = /^\s*(re|fw|fwd|aw|wg|\u05EA\u05E9\u05D5\u05D1\u05D4|\u05D4\u05E9\u05D1|\u05D4\u05D5\u05E2\u05D1\u05E8)\s*:/i;
  // keyword Category -> Impact option; categories that also suggest Variation
  var IMPACT_OF = { "cost": "Cost", "time & delay": "Time", "time": "Time", "scope": "Scope", "ehs": "EHS", "quality": "Quality" };
  var VARIATION_CATS = ["cost", "time & delay", "time", "scope", "change / variation", "contractual notice"];
  var MSAL_URLS = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@4/lib/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js"
  ];
  var STATUS_OUT = ["Issued", "Follow-up", "Claim", "Cancelled"];
  var STATUS_IN = ["Logged", "Follow-up", "Claim"];

  var regCols = null, pca = null, siteId = null, listIds = {}, companies = [], keywords = [], item = null;
  var $ = function (s) { return document.querySelector(s); };
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function pad4(n) { return String(n).padStart(4, "0"); }
  var msgTimer;
  function toast(t, isErr) { var m = $("#msg"); m.textContent = t; m.className = "msg" + (isErr ? " e" : ""); m.hidden = false; clearTimeout(msgTimer); msgTimer = setTimeout(function () { m.hidden = true; }, isErr ? 8000 : 4000); }
  function show(html) { $("#app").innerHTML = html; }
  function showError(title, detail) {
    show('<div class="err"><b>' + esc(title) + '</b><br>' + esc(detail || "") + '</div><button class="btns" id="retry">Try again</button>');
    $("#retry").onclick = start;
  }

  /* ---------- Office helpers ---------- */
  function oget(prop, opt) {
    return new Promise(function (res, rej) {
      var cb = function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error); };
      opt ? prop.getAsync(opt, cb) : prop.getAsync(cb);
    });
  }
  function oset(prop, val) {
    return new Promise(function (res, rej) { prop.setAsync(val, function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error); }); });
  }
  function bodyText() {
    return new Promise(function (res) { item.body.getAsync(Office.CoercionType.Text, function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : ""); }); });
  }
  function roam() { try { return Office.context.roamingSettings; } catch (e) { return null; } }
  function saveRoam() { var rs = roam(); if (rs) rs.saveAsync(function () {}); }

  /* ---------- Auth (MSAL, nested app auth when available) ---------- */
  function loadScript(u) { return new Promise(function (res, rej) { var s = document.createElement("script"); s.src = u; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); }
  async function loadMsal() {
    for (var i = 0; i < MSAL_URLS.length; i++) {
      try { await loadScript(MSAL_URLS[i]); if (window.msal) return; } catch (e) { /* next */ }
    }
    throw new Error("The sign-in library could not be loaded. Check the internet connection.");
  }
  async function initAuth() {
    await loadMsal();
    var cfg = {
      auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/" + C.tenantId, redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" }
    };
    var naa = false;
    try { naa = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1") && !!msal.createNestablePublicClientApplication; } catch (e) { naa = false; }
    if (naa) { pca = await msal.createNestablePublicClientApplication(cfg); }
    else { pca = new msal.PublicClientApplication(cfg); await pca.initialize(); }
  }
  async function token(interactive) {
    var acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
    var req = acc ? { scopes: SCOPES, account: acc } : { scopes: SCOPES };
    try {
      var r = await pca.acquireTokenSilent(req);
      return r.accessToken;
    } catch (e) {
      if (!interactive) throw Object.assign(new Error("signin"), { needSignIn: true });
      var p = await pca.acquireTokenPopup({ scopes: SCOPES });
      if (p.account) pca.setActiveAccount(p.account);
      return p.accessToken;
    }
  }

  // Separate token for To Do tasks: never prompts, so the rest works before IT approves Tasks.ReadWrite
  async function taskToken() {
    var acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
    var r = await pca.acquireTokenSilent(acc ? { scopes: TASK_SCOPES, account: acc } : { scopes: TASK_SCOPES });
    return r.accessToken;
  }
  async function gTask(method, path, body) {
    var t = await taskToken();
    var r = await fetch("https://graph.microsoft.com/v1.0" + path, { method: method, headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    if (r.status === 204) return null;
    var j = null; try { j = await r.json(); } catch (e) { j = null; }
    if (!r.ok) { var err = new Error((j && j.error && j.error.message) || ("HTTP " + r.status)); err.status = r.status; throw err; }
    return j;
  }

  /* ---------- Graph / SharePoint ---------- */
  async function g(method, path, body, headers) {
    var t = await token(false);
    var h = Object.assign({ Authorization: "Bearer " + t, "Content-Type": "application/json", Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }, headers || {});
    var r = await fetch(path.indexOf("https://") === 0 ? path : "https://graph.microsoft.com/v1.0" + path, { method: method, headers: h, body: body ? JSON.stringify(body) : undefined });
    if (r.status === 204) return null;
    var txt = await r.text(), j = null;
    try { j = txt ? JSON.parse(txt) : null; } catch (e) { j = null; }
    if (!r.ok) { var err = new Error((j && j.error && j.error.message) || ("HTTP " + r.status)); err.status = r.status; throw err; }
    return j;
  }
  function listPath(key) { return "/sites/" + siteId + "/lists/" + listIds[key]; }
  async function items(key, select, filter) {
    var url = listPath(key) + "/items?$expand=fields($select=" + select + ")&$top=500" + (filter ? "&$filter=" + encodeURIComponent(filter) : "");
    var out = [];
    while (url) { var r = await g("GET", url); out = out.concat(r.value || []); url = r["@odata.nextLink"] || null; }
    return out;
  }
  function q(s) { return String(s).replace(/'/g, "''"); }

  async function loadData() {
    var s = await g("GET", "/sites/" + C.siteHost + ":" + C.sitePath + "?$select=id");
    siteId = s.id;
    var ls = await g("GET", "/sites/" + siteId + "/lists?$select=id,displayName,name&$top=200");
    Object.keys(C.lists).forEach(function (k) {
      var n = C.lists[k].toLowerCase();
      var f = (ls.value || []).find(function (l) { return (l.displayName || "").toLowerCase() === n || (l.name || "").toLowerCase() === n; });
      if (!f) throw new Error('The SharePoint list "' + C.lists[k] + '" was not found on the site.');
      listIds[k] = f.id;
    });
    try {
      var cols = await g("GET", listPath("register") + "/columns?$select=name,choice&$top=200");
      regCols = {};
      (cols.value || []).forEach(function (c) { regCols[c.name] = c; });
    } catch (e) { regCols = null; }
    var co;
    try { co = await items("companies", "Title,CompanyName,Domains,LastNumber,RefPrefix"); }
    catch (e) { co = await items("companies", "Title,CompanyName,Domains,LastNumber"); }
    companies = co.map(function (x) {
      var f = x.fields || {};
      return { id: x.id, code: String(f.Title || "").trim().toUpperCase(), name: f.CompanyName || "", last: Number(f.LastNumber || 0),
               domains: String(f.Domains || "").split(/[;,\s]+/).map(function (d) { return d.trim().toLowerCase().replace(/^@/, ""); }).filter(Boolean),
               prefixes: String(f.RefPrefix || "").split(/[;,\s]+/).map(function (d) { return d.trim().toUpperCase(); }).filter(Boolean) };
    }).filter(function (c) { return c.code && c.code !== C.issuer; }).sort(function (a, b) { return a.code.localeCompare(b.code); });
    var kw = await items("keywords", "Title,Category");
    keywords = kw.map(function (x) { var f = x.fields || {}; return { term: String(f.Title || "").trim(), cat: String(f.Category || "").trim().toLowerCase() }; })
                 .filter(function (k) { return k.term; });
    // cache domains for the Send reminder
    var rs = roam();
    if (rs) {
      var map = {}; companies.forEach(function (c) { if (c.domains.length) map[c.code] = c.domains; });
      var names = {}; companies.forEach(function (c) { if (c.name) names[c.code] = c.name; });
      rs.set("sorDomains", JSON.stringify(map)); rs.set("sorNames", JSON.stringify(names)); saveRoam();
    }
  }

  function scan(text) {
    var t = (text || "").toLowerCase(), hits = [];
    keywords.forEach(function (k) { if (t.indexOf(k.term.toLowerCase()) >= 0 && hits.indexOf(k.term) < 0) hits.push(k.term); });
    return hits;
  }
  // What the keywords found suggest: Impact options to pre-tick, and whether to pre-tick Variation
  function suggest(hits) {
    var opts = impactOptions(), imp = [], variation = false;
    hits.forEach(function (term) {
      var k = keywords.find(function (x) { return x.term === term; }); if (!k) return;
      var i = IMPACT_OF[k.cat];
      if (i && opts.indexOf(i) >= 0 && imp.indexOf(i) < 0) imp.push(i);
      if (VARIATION_CATS.indexOf(k.cat) >= 0) variation = true;
    });
    return { impact: imp, variation: variation };
  }
  function impactOptions() {
    var col = regCols && regCols.Impact;
    return (col && col.choice && col.choice.choices && col.choice.choices.length) ? col.choice.choices.slice() : ["Cost", "Time", "Scope"];
  }
  function impactChecks(selected) {
    return '<div><div class="lbl">Impact</div><div class="checks">' + impactOptions().map(function (k) {
      return '<label><input type="checkbox" data-imp="' + esc(k) + '"' + (selected.indexOf(k) >= 0 ? " checked" : "") + '> ' + esc(k) + '</label>';
    }).join("") + '</div></div>';
  }
  function readImpact() {
    return Array.prototype.slice.call(document.querySelectorAll("[data-imp]")).filter(function (c) { return c.checked; }).map(function (c) { return c.dataset.imp; });
  }
  /* ---------- Due date ("Reply expected by") ---------- */
  function isoDay(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function addDays(n) { var d = new Date(); d.setDate(d.getDate() + n); return isoDay(d); }
  function showDay(iso) { var p = String(iso || "").slice(0, 10).split("-"); return p.length === 3 ? p[2] + "/" + p[1] + "/" + p[0] : ""; }
  function dueBlock(prefill) {
    var v = prefill ? String(prefill).slice(0, 10) : "";
    return '<div><div class="lbl">Reply expected by</div><div class="pills" id="duePills">' +
      [["0", "None"], ["3", "3 days"], ["7", "7 days"], ["14", "14 days"]].map(function (x) {
        return '<button class="pill" data-due="' + x[0] + '" aria-pressed="' + (x[0] === "0" && !v) + '">' + x[1] + '</button>';
      }).join("") + '</div><input type="date" id="dueDate" style="margin-top:8px" value="' + esc(v) + '">' +
      '<p class="muted" style="margin-top:4px">Optional. With a date, you get a reminder in your To Do on that morning.</p></div>';
  }
  function wireDue() {
    var p = $("#duePills"), inp = $("#dueDate"); if (!p) return;
    p.onclick = function (e) {
      var b = e.target.closest("[data-due]"); if (!b) return;
      document.querySelectorAll("#duePills .pill").forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); });
      inp.value = b.dataset.due === "0" ? "" : addDays(Number(b.dataset.due));
    };
    inp.onchange = function () { document.querySelectorAll("#duePills .pill").forEach(function (x) { x.setAttribute("aria-pressed", String(!inp.value && x.dataset.due === "0")); }); };
  }
  function readDue() { var i = $("#dueDate"); return i && i.value ? i.value : ""; }
  // Reminder task in the user's own Microsoft To Do (needs Tasks.ReadWrite approved by IT)
  async function createTask(num, topic, due, recips) {
    try {
      var lists = await gTask("GET", "/me/todo/lists");
      var list = (lists.value || []).find(function (l) { return l.wellknownListName === "defaultList"; }) || (lists.value || [])[0];
      if (!list) return { ok: false, msg: "No To Do list found" };
      var tz = "Israel Standard Time";
      var t = await gTask("POST", "/me/todo/lists/" + list.id + "/tasks", {
        title: "Follow-up: " + num + " – " + topic,
        body: { contentType: "text", content: "Reply expected by " + showDay(due) + ".\nRecipients: " + (recips || []).join("; ") },
        dueDateTime: { dateTime: due + "T00:00:00", timeZone: tz },
        isReminderOn: true,
        reminderDateTime: { dateTime: due + "T09:00:00", timeZone: tz }
      });
      var rs = roam();
      if (rs) {
        var map = {}; try { map = JSON.parse(rs.get("sorTasks") || "{}"); } catch (e) { map = {}; }
        map[num] = { l: list.id, t: t.id };
        var keys = Object.keys(map); while (keys.length > 40) { delete map[keys.shift()]; }
        rs.set("sorTasks", JSON.stringify(map)); saveRoam();
      }
      return { ok: true };
    } catch (e) {
      return { ok: false, msg: "Reminder not created yet – waiting for IT approval" };
    }
  }
  async function deleteTask(num) {
    var rs = roam(); if (!rs) return;
    var map = {}; try { map = JSON.parse(rs.get("sorTasks") || "{}"); } catch (e) { return; }
    var x = map[num]; if (!x) return;
    try { await gTask("DELETE", "/me/todo/lists/" + x.l + "/tasks/" + x.t); } catch (e) { /* already gone */ }
    delete map[num]; rs.set("sorTasks", JSON.stringify(map)); saveRoam();
  }
  /* ---------- Header block at the top of the email body ---------- */
  function headerHtml(num, refs, topic) {
    var td1 = 'style="padding:1px 14px 1px 0;color:#16324F;font-weight:bold;white-space:nowrap;vertical-align:top"';
    var td2 = 'style="padding:1px 0;vertical-align:top"';
    var rows = '<tr><td ' + td1 + '>Correspondence ID:</td><td ' + td2 + '>' + esc(num) + '</td></tr>';
    if (refs.length) rows += '<tr><td ' + td1 + '>References:</td><td ' + td2 + '>' + refs.map(function (r) { return esc(r.ref) + (r.subject ? " – " + esc(r.subject) : ""); }).join("<br>") + '</td></tr>';
    rows += '<tr><td ' + td1 + '>Subject:</td><td ' + td2 + '>' + esc(topic) + '</td></tr>';
    return '<table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10.5pt;margin:0 0 14px 0">' + rows + '</table><p>&nbsp;</p>';
  }
  function headerText(num, refs, topic) {
    return "Correspondence ID: " + num + "\n" +
      (refs.length ? "References: " + refs.map(function (r) { return r.ref + (r.subject ? " – " + r.subject : ""); }).join("; ") + "\n" : "") +
      "Subject: " + topic + "\n\n";
  }
  function bodyType() { return new Promise(function (res) { item.body.getTypeAsync(function (r) { res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "html"); }); }); }
  async function insertHeader(num, refs, topic) {
    var html = (await bodyType()) === Office.CoercionType.Html;
    await new Promise(function (res, rej) {
      item.body.prependAsync(html ? headerHtml(num, refs, topic) : headerText(num, refs, topic), { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text },
        function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error); });
    });
  }
  async function removeHeader(num) {
    try {
      var html = (await bodyType()) === Office.CoercionType.Html;
      var cur = await new Promise(function (res, rej) { item.body.getAsync(html ? Office.CoercionType.Html : Office.CoercionType.Text, function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res(r.value || "") : rej(r.error); }); });
      var at = cur.indexOf("Correspondence ID");
      if (at < 0 || cur.indexOf(num, at) < 0 || cur.indexOf(num, at) - at > 800) return false;
      var out;
      if (html) {
        var a = cur.lastIndexOf("<table", at), b = cur.indexOf("</table>", at);
        if (a < 0 || b < 0) return false;
        out = cur.slice(0, a) + cur.slice(b + 8);
      } else {
        var e = cur.indexOf("Subject:", at); var nl = e >= 0 ? cur.indexOf("\n", e) : -1;
        if (nl < 0) return false;
        out = cur.slice(0, at) + cur.slice(nl + 1).replace(/^\n/, "");
      }
      await new Promise(function (res, rej) { item.body.setAsync(out, { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text }, function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error); }); });
      return true;
    } catch (e) { return false; }
  }
  // Topic without RE:/FW: prefixes and without correspondence numbers
  function cleanTopic(subject) {
    var s = String(subject || "");
    for (var i = 0; i < 5; i++) s = s.replace(REPLY_RE, "").trim();
    s = s.replace(NUM_RE_G, " ");
    externalRefs(s).forEach(function (r) { s = s.split(r.ref).join(" ").split(r.ref.toLowerCase()).join(" "); });
    for (var j = 0; j < 5; j++) s = s.replace(REPLY_RE, "").trim();
    return s.replace(/\s{2,}/g, " ").replace(/^[\s\-_–:]+|[\s\-_–:]+$/g, "").trim();
  }
  function refsInSubject(subject) {
    var out = [], seen = {};
    (String(subject || "").match(NUM_RE_G) || []).forEach(function (n) { n = n.toUpperCase(); if (!seen[n]) { seen[n] = true; out.push(n); } });
    externalRefs(String(subject || "")).forEach(function (r) { if (!seen[r.ref]) { seen[r.ref] = true; out.push(r.ref); } });
    return out;
  }
  async function subjectOf(ref) {
    try {
      var r = await items("register", "Title,MailSubject", "fields/Title eq '" + q(ref) + "'");
      return r.length ? cleanTopic((r[0].fields || {}).MailSubject || "") : "";
    } catch (e) { return ""; }
  }
  async function recentFor(code) {
    try {
      var r = await items("register", "Title,MailSubject,MailDate,Status,Company", "fields/Company eq '" + q(code) + "'");
      return r.map(function (x) { return x.fields || {}; })
        .filter(function (f) { return f.Status !== "Cancelled" && f.Title && f.Title !== "(no number)"; })
        .sort(function (a, b) { return String(b.MailDate || "").localeCompare(String(a.MailDate || "")); })
        .slice(0, 5)
        .map(function (f) { return { ref: f.Title, subject: cleanTopic(f.MailSubject || ""), date: f.MailDate }; });
    } catch (e) { return []; }
  }
  function domainOf(addr) { return String(addr || "").toLowerCase().split("@")[1] || ""; }
  function codeForAddresses(addrs) {
    for (var i = 0; i < addrs.length; i++) {
      var d = domainOf(addrs[i]); if (!d) continue;
      for (var j = 0; j < companies.length; j++) {
        var c = companies[j];
        if (c.domains.some(function (x) { return d === x || d.endsWith("." + x); })) return c.code;
      }
    }
    return null;
  }
  function codeFromNumber(num) {
    var p = String(num || "").toUpperCase().split("-");
    if (p.length < 4) return null;
    return p[1] === C.issuer ? p[2] : p[1];
  }
  function coNameOf(code) { var c = companies.find(function (x) { return x.code === String(code || "").toUpperCase(); }); return c && c.name && c.name !== c.code ? c.name : ""; }
  function companyChoice(code) {
    if (!code) return "Other";
    var col = regCols && regCols.Company;
    if (col && col.choice) {
      var ok = (col.choice.choices || []).indexOf(code) >= 0 || col.choice.allowTextEntry;
      return ok ? code : "Other";
    }
    return code;
  }
  function isOwnDomain(d) { return (C.ownDomains || []).some(function (x) { x = String(x).toLowerCase(); return d === x || d.endsWith("." + x); }); }
  // Other parties' own correspondence numbers, e.g. Siemens "IL5002-SE-SPE-00187" (prefix set in Companies.RefPrefix).
  // Returns [{ref, code}] in the order they appear: the first is usually the email's own number, the rest are references.
  function externalRefs(text) {
    var out = [], seen = {};
    companies.forEach(function (c) {
      c.prefixes.forEach(function (p) {
        var re = new RegExp(p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "[A-Z0-9-]*\\d", "gi"), m;
        while ((m = re.exec(text)) !== null) {
          var r = m[0].toUpperCase();
          if (!seen[r]) { seen[r] = true; out.push({ ref: r, code: c.code, at: m.index }); }
        }
      });
    });
    return out.sort(function (a, b) { return a.at - b.at; });
  }
  // Create a register row; tolerate a Company column without the new code and a missing ExternalRef column
  async function createRegister(fields) {
    if (regCols) {
      fields = Object.assign({}, fields);
      Object.keys(fields).forEach(function (k) { if (k !== "Title" && k.indexOf("@") < 0 && !regCols[k]) delete fields[k]; });
    }
    try { return await g("POST", listPath("register") + "/items", { fields: fields }); }
    catch (e) {
      if (e.status !== 400) throw e;
      var f2 = Object.assign({}, fields);
      if (/ExternalRef/i.test(e.message)) delete f2.ExternalRef;
      else if (f2.Company && f2.Company !== "Other") f2.Company = "Other";
      else if ("ExternalRef" in f2) delete f2.ExternalRef;
      else throw e;
      return await createRegister(f2);
    }
  }
  function regSelect() {
    var want = ["Title", "Direction", "Status", "Variation", "Impact", "ClaimsNotes", "MailSubject", "Company", "ExternalRef", "CancelReason", "DueDate", "Recipients"];
    return (regCols ? want.filter(function (k) { return k === "Title" || regCols[k]; }) : want.slice(0, 8)).join(",");
  }
  async function patchRegister(id, fields) {
    var f = Object.assign({}, fields);
    if (regCols) Object.keys(f).forEach(function (k) { if (k.indexOf("@") < 0 && !regCols[k]) { delete f[k]; delete f[k + "@odata.type"]; } });
    return g("PATCH", listPath("register") + "/items/" + id + "/fields", f);
  }
  function reasonForm(prefill) {
    return '<div class="lbl">Reason for cancelling</div><input type="text" id="rsText" value="' + esc(prefill || "Email not sent") + '">';
  }
  function wireReasons() { /* plain text field, nothing to wire */ }
  function fmt(code, n) { return C.project + "-" + C.issuer + "-" + code + "-" + pad4(n); }

  /* ---------- Number reservation (safe for several users at once) ---------- */
  async function reserve(code, buildFields) {
    var co = companies.find(function (c) { return c.code === code; });
    if (!co) throw new Error("Company " + code + " is not in the Companies list.");
    for (var attempt = 0; attempt < 8; attempt++) {
      var cur = await g("GET", listPath("companies") + "/items/" + co.id + "?$expand=fields($select=Title,LastNumber)");
      var n = Number((cur.fields || {}).LastNumber || 0) + 1;
      var num = fmt(code, n);
      try {
        await g("PATCH", listPath("companies") + "/items/" + co.id + "/fields", { LastNumber: n }, cur["@odata.etag"] ? { "If-Match": cur["@odata.etag"] } : {});
      } catch (e) {
        if (e.status === 412 || e.status === 409) { await sleep(300 + Math.random() * 700); continue; }
        if (e.status === 400) { await g("PATCH", listPath("companies") + "/items/" + co.id + "/fields", { LastNumber: n }); }
        else throw e;
      }
      var created = await createRegister(Object.assign(buildFields(num), { Title: num }));
      var same = await items("register", "Title", "fields/Title eq '" + q(num) + "'");
      if (same.length > 1) {
        var ids = same.map(function (x) { return Number(x.id); }).sort(function (a, b) { return a - b; });
        if (Number(created.id) !== ids[0]) { await g("DELETE", listPath("register") + "/items/" + created.id); continue; }
      }
      co.last = n;
      return { num: num, id: created.id };
    }
    throw new Error("Could not get a free number. Please try again.");
  }

  async function findRecord(opts) {
    if (opts.messageId) {
      var a = await items("register", regSelect(), "fields/MessageId eq '" + q(opts.messageId) + "'");
      if (a.length) return a[0];
    }
    if (opts.number) {
      var b = await items("register", regSelect(), "fields/Title eq '" + q(opts.number) + "'");
      var out = b.filter(function (x) { return (x.fields || {}).Direction === "Out"; });
      if (out.length) return out[0];
    }
    return null;
  }

  /* ---------- Record panel (shared by read and compose) ---------- */
  function recordPanel(rec, hits, opts) {
    opts = opts || {};
    var f = rec.fields || {};
    var dir = f.Direction || "In";
    var ownNumber = dir === "Out" && String(f.Title || "").toUpperCase().indexOf(C.project + "-" + C.issuer + "-") === 0 && f.Status !== "Logged";
    var statuses = ownNumber ? STATUS_OUT : STATUS_IN;
    var impact = Array.isArray(f.Impact) ? f.Impact : [];
    var st = { status: f.Status || statuses[0] };
    var refs = String(f.ExternalRef || "").split(/;\s*/).filter(Boolean);
    var refLine = "";
    if (refs.length) {
      if (ownNumber) refLine = 'References: <span class="mono">' + esc(refs.join(", ")) + '</span>';
      else {
        var other = refs.filter(function (r) { return r.toUpperCase() !== String(f.Title || "").toUpperCase(); });
        refLine = other.length && other.length < refs.length ? 'Refers to: <span class="mono">' + esc(other.join(", ")) + '</span>' : 'Their ref: <span class="mono">' + esc(refs.join(", ")) + '</span>';
      }
    }
    show(
      '<div class="ok"><b>In the register</b><div class="num" style="margin-top:4px">' + esc(f.Title || "") + '</div>' +
      (refLine ? '<p class="muted" style="margin-top:4px">' + refLine + '</p>' : "") +
      '<p class="muted">' + (dir === "Out" ? "Outgoing" : "Incoming") + (f.Company ? " · " + esc(f.Company) + (coNameOf(f.Company) ? " – " + esc(coNameOf(f.Company)) : "") : "") + '</p></div>' +
      (hits && hits.length ? '<div><div class="lbl">Keywords found</div>' + hits.map(function (k) { return '<span class="kw">' + esc(k) + '</span>'; }).join("") + '</div>' : "") +
      '<div><div class="lbl">Status</div><div class="pills" id="stPills">' + statuses.map(function (s) { return '<button class="pill" data-st="' + esc(s) + '" aria-pressed="' + (s === st.status) + '">' + esc(s) + '</button>'; }).join("") + '</div></div>' +
      '<label class="varbox"><input type="checkbox" id="recVar"' + (f.Variation ? " checked" : "") + '><span><b>Variation</b><br><span class="muted">Tracked by the Claims Manager</span></span></label>' +
      impactChecks(impact) +
      (regCols && !regCols.DueDate ? "" : dueBlock(f.DueDate)) +
      '<div><div class="lbl">Claims notes</div><textarea id="recNotes">' + esc(f.ClaimsNotes || "") + '</textarea></div>' +
      '<div id="cancelBox"' + (st.status === "Cancelled" ? "" : " hidden") + '>' + reasonForm(f.CancelReason) + '</div>' +
      '<button class="btnp" id="recSave">Save</button>' +
      (opts.compose && ownNumber && f.Status !== "Cancelled" ? '<div class="warn" id="relBox" style="margin-top:4px">Not sending this email? <button class="btns" id="relBtn" style="margin-top:8px;width:100%">Release number</button></div>' : "")
    );
    wireDue();
    if ($("#relBtn")) $("#relBtn").onclick = function () { askRelease(rec.id, f.Title, false); };
    $("#stPills").onclick = function (e) {
      var b = e.target.closest("[data-st]"); if (!b) return;
      st.status = b.dataset.st;
      document.querySelectorAll("#stPills .pill").forEach(function (p) { p.setAttribute("aria-pressed", String(p === b)); });
      if (st.status === "Claim") $("#recVar").checked = true;
      $("#cancelBox").hidden = st.status !== "Cancelled";
    };
    var dueBefore = f.DueDate ? String(f.DueDate).slice(0, 10) : "";
    $("#recSave").onclick = async function () {
      var btn = this;
      var reason = $("#rsText") ? $("#rsText").value.trim() : "";
      if (st.status === "Cancelled" && !reason) { toast("Write why the number is cancelled", true); $("#rsText").focus(); return; }
      btn.disabled = true; btn.textContent = "Saving…";
      var due = readDue();
      try {
        var upd = { Status: st.status, Variation: $("#recVar").checked, ClaimsNotes: $("#recNotes").value,
          "Impact@odata.type": "Collection(Edm.String)", Impact: readImpact() };
        if (st.status === "Cancelled") upd.CancelReason = reason;
        if (due !== dueBefore) upd.DueDate = due ? due + "T12:00:00Z" : null;
        await patchRegister(rec.id, upd);
        var note = "Saved in the register";
        if (due && due !== dueBefore) {
          var t = await createTask(f.Title, cleanTopic(f.MailSubject || ""), due, String(f.Recipients || "").split(/;\s*/).filter(Boolean));
          note += t.ok ? " · reminder added to your To Do" : " · " + t.msg;
          dueBefore = due;
        }
        toast(note);
      } catch (e) { toast("Not saved: " + e.message, true); }
      btn.disabled = false; btn.textContent = "Save";
    };
  }

  /* ---------- Compose ---------- */
  async function composeView() {
    var subject = await oget(item.subject);
    var to = await oget(item.to), cc = await oget(item.cc);
    var addrs = to.concat(cc).map(function (r) { return r.emailAddress; });
    var isReply = REPLY_RE.test(subject || "");
    // Already numbered by us (a new email after Insert, or a saved draft)
    var own = !isReply && (subject || "").match(new RegExp("^\\s*" + C.project + "-" + C.issuer + "-[A-Z]{3}-\\d{4}", "i"));
    if (own) {
      var num0 = own[0].trim().toUpperCase();
      var rec = await findRecord({ number: num0 });
      if (rec) { recordPanel(rec, [], { compose: true }); return; }
      show('<div class="warn">This email already has the number <b class="mono">' + esc(num0) + '</b> in the subject, but it is not in the register.</div>');
      return;
    }
    var body = await bodyText();
    var hits = scan(subject + " " + body);
    var sug = suggest(hits);
    var subjRefs = refsInSubject(subject);
    var detected = codeForAddresses(addrs);
    if (!detected && subjRefs.length) { var c0 = codeFromNumber(subjRefs[0]) || (externalRefs(subjRefs[0])[0] || {}).code; if (c0 && companies.some(function (c) { return c.code === c0; })) detected = c0; }
    var st = { code: detected || (companies[0] && companies[0].code), picked: false };
    var autoRefs = [];
    for (var i = 0; i < subjRefs.length && i < 5; i++) autoRefs.push({ ref: subjRefs[i], subject: await subjectOf(subjRefs[i]) });
    var recentCache = {};
    var form = { topic: cleanTopic(subject) || "", variation: sug.variation, impact: sug.impact, refs: autoRefs.map(function (r) { return r.ref; }), extra: "", due: "" };

    try {
      item.addHandlerAsync(Office.EventType.RecipientsChanged, async function () {
        try {
          var t2 = await oget(item.to), c2 = await oget(item.cc);
          addrs = t2.concat(c2).map(function (r) { return r.emailAddress; });
          var d2 = codeForAddresses(addrs);
          if (d2) detected = d2;
          if (!st.picked && d2 && d2 !== st.code) { st.code = d2; keep(); await render(); }
          else if ($("#insert") && !$("#insert").disabled) { keep(); await render(); }
        } catch (e) { /* ignore */ }
      });
    } catch (e) { /* older Outlook */ }

    function keep() {
      if (!$("#topic")) return;
      form.topic = $("#topic").value; form.variation = $("#var").checked; form.impact = readImpact();
      form.refs = Array.prototype.slice.call(document.querySelectorAll("[data-ref]")).filter(function (c) { return c.checked; }).map(function (c) { return c.dataset.ref; });
      form.extra = $("#refExtra").value; form.due = readDue();
    }

    async function render() {
      var co = companies.find(function (c) { return c.code === st.code; });
      var next = co ? fmt(co.code, co.last + 1) : "—";
      if (co && !recentCache[co.code]) recentCache[co.code] = await recentFor(co.code);
      var recent = co ? recentCache[co.code].filter(function (r) { return !autoRefs.some(function (a) { return a.ref === r.ref; }); }) : [];
      var refRow = function (r, auto) {
        return '<label style="display:flex;gap:8px;align-items:flex-start;font-size:13px;padding:3px 0"><input type="checkbox" data-ref="' + esc(r.ref) + '" data-sub="' + esc(r.subject || "") + '"' + (form.refs.indexOf(r.ref) >= 0 ? " checked" : "") + ' style="margin-top:3px">' +
          '<span><b class="mono">' + esc(r.ref) + '</b>' + (r.subject ? ' · ' + esc(r.subject) : "") + (r.date ? ' <span class="muted">· ' + esc(showDay(r.date)) + '</span>' : "") + (auto ? ' <span class="muted">(from the subject)</span>' : "") + '</span></label>';
      };
      show(
        (isReply ? '<div class="note" style="background:var(--blue-soft);border-radius:6px;padding:10px 12px;font-size:13px">Reply – a number is optional. References are filled in from the original email.</div>' : "") +
        (addrs.length ? "" : '<div class="warn">No recipients yet. Add them first, so the company is detected and the recipients are recorded.</div>') +
        '<div><div class="lbl">Recipient' + (detected ? " · detected" : "") + '</div><div class="pills" id="coPills">' +
        companies.map(function (c) { return '<button class="pill" data-co="' + esc(c.code) + '" aria-pressed="' + (c.code === st.code) + '" title="' + esc(c.name) + '">' + esc(c.code) + '</button>'; }).join("") + '</div>' +
        (co ? '<div class="coname"><b>' + esc(co.code) + '</b> – ' + esc(co.name || co.code) + '</div>' : "") + '</div>' +
        '<div><div class="lbl">Next number</div><div class="num">' + esc(next) + '</div><p class="muted">Taken only when you click Insert. If someone else takes it first, you get the next one.</p></div>' +
        '<div><div class="lbl">Topic</div><input type="text" id="topic" value="' + esc(form.topic) + '" placeholder="Short topic of the email"></div>' +
        '<div><div class="lbl">References</div>' +
          autoRefs.map(function (r) { return refRow(r, true); }).join("") +
          (recent.length ? '<p class="muted" style="margin:6px 0 2px">Recent with ' + esc(st.code) + ':</p>' + recent.map(function (r) { return refRow(r, false); }).join("") : "") +
          '<input type="text" id="refExtra" style="margin-top:6px" placeholder="Other reference (optional)" value="' + esc(form.extra) + '"></div>' +
        '<label class="varbox"><input type="checkbox" id="var"' + (form.variation ? " checked" : "") + '><span><b>Track as variation</b><br><span class="muted">' +
        (hits.length ? "Keywords found: " + esc(hits.join(", ")) : "Tick if this email may support a variation request or claim.") + '</span></span></label>' +
        impactChecks(form.impact) +
        (regCols && !regCols.DueDate ? "" : dueBlock(form.due)) +
        '<button class="btnp" id="insert"' + (co ? "" : " disabled") + '>Insert number</button>'
      );
      wireDue();
      $("#coPills").onclick = async function (e) { var b = e.target.closest("[data-co]"); if (!b) return; keep(); st.code = b.dataset.co; st.picked = true; await render(); };
      $("#insert").onclick = doInsert;
    }

    async function doInsert() {
      keep();
      var btn = $("#insert"); btn.disabled = true; btn.textContent = "Getting number…";
      var topic = form.topic.trim() || "Topic";
      var refs = Array.prototype.slice.call(document.querySelectorAll("[data-ref]")).filter(function (c) { return c.checked; })
        .map(function (c) { return { ref: c.dataset.ref, subject: c.dataset.sub }; });
      form.extra.split(/[;,]/).map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (x) { refs.push({ ref: x, subject: "" }); });
      var due = form.due;
      try {
        addrs = (await oget(item.to)).concat(await oget(item.cc)).map(function (r) { return r.emailAddress; });
        var res = await reserve(st.code, function (num) {
          var fl = {
            Direction: "Out", Company: companyChoice(st.code), MailSubject: num + "_" + topic, MailDate: new Date().toISOString(),
            Status: "Issued", Variation: form.variation, Keywords: hits.join(", "),
            Sender: (Office.context.mailbox.userProfile || {}).emailAddress || "", Recipients: addrs.join("; "),
            ExternalRef: refs.map(function (r) { return r.ref; }).join("; "),
            "Impact@odata.type": "Collection(Edm.String)", Impact: form.impact
          };
          if (due) fl.DueDate = due + "T12:00:00Z";
          return fl;
        });
        try { await oset(item.subject, res.num + "_" + topic); }
        catch (err) {
          // the number exists in the register but did not reach the email: cancel it so nothing is left hanging
          try { await patchRegister(res.id, { Status: "Cancelled", CancelReason: "Not inserted in the email (connection error)" }); } catch (e2) { /* best effort */ }
          throw err;
        }
        var all = (await oget(item.to)).concat(await oget(item.cc)).map(function (r) { return String(r.emailAddress || "").toLowerCase(); });
        var ccAdded = false;
        if (all.indexOf(C.generalMailbox.toLowerCase()) < 0) {
          await new Promise(function (r) { item.cc.addAsync([C.generalMailbox], function () { r(); }); });
          ccAdded = true;
        }
        var headerOk = true;
        try { await insertHeader(res.num, refs, topic); } catch (e) { headerOk = false; }
        var task = null;
        if (due) task = await createTask(res.num, topic, due, addrs);
        show(
          '<div class="ok"><b>Number inserted</b><div class="num" style="margin-top:4px">' + esc(res.num) + '</div>' + (coNameOf(st.code) ? '<p class="muted">' + esc(coNameOf(st.code)) + '</p>' : "") + '</div>' +
          '<ul class="done"><li>Number added to the subject</li>' +
          '<li>' + (ccAdded ? "sorek@sorek.energy added to CC" : "sorek@sorek.energy already in CC") + '</li>' +
          '<li>' + (headerOk ? "Correspondence ID" + (refs.length ? ", References" : "") + " and Subject added at the top of the email" : "The header could not be added – add it by hand if needed") + '</li>' +
          '<li>Logged in the register as <b>Issued</b></li>' + (form.variation ? '<li>Marked <b>Variation</b></li>' : "") +
          (form.impact.length ? '<li>Impact: ' + esc(form.impact.join(", ")) + '</li>' : "") +
          (due ? '<li>Reply expected by ' + esc(showDay(due)) + (task && task.ok ? ' – reminder in your To Do' : ' – ' + esc(task ? task.msg : "")) + '</li>' : "") + '</ul>' +
          '<p class="muted">Finish the email and click Send as usual.</p>' +
          '<div class="warn" id="relBox">Not sending this email after all? <button class="btns" id="relBtn" style="margin-top:8px;width:100%">Release number</button><span class="muted" style="display:block;margin-top:6px">The number is marked Cancelled in the register and removed from the email. Cancelled numbers are not reused.</span></div>'
        );
        $("#relBtn").onclick = function () { askRelease(res.id, res.num, ccAdded); };
        try { item.notificationMessages.removeAsync("sorekHint", function () {}); } catch (e) { /* none */ }
      } catch (e) {
        toast("The number was not inserted: " + e.message, true);
        btn.disabled = false; btn.textContent = "Insert number";
      }
    }
    await render();
  }

  /* ---------- Release a number that will not be sent ---------- */
  function askRelease(recId, num, removeCc) {
    var box = $("#relBox"), before = box.innerHTML;
    box.innerHTML = '<b>Release ' + esc(num) + '</b><p class="muted" style="margin:4px 0 10px">Marked Cancelled in the register. Change the reason if needed.</p>' +
      reasonForm("") +
      '<div style="display:flex;gap:8px;margin-top:10px"><button class="btnp" id="relGo" style="flex:1">Release</button><button class="btns" id="relBack">Back</button></div>';
    $("#relBack").onclick = function () { box.innerHTML = before; $("#relBtn").onclick = function () { askRelease(recId, num, removeCc); }; };
    $("#relGo").onclick = function () {
      var reason = $("#rsText").value.trim();
      if (!reason) { toast("Write why the number is cancelled", true); $("#rsText").focus(); return; }
      this.disabled = true; this.textContent = "Releasing…";
      releaseNumber(recId, num, removeCc, reason);
    };
  }
  async function releaseNumber(recId, num, removeCc, reason) {
    try {
      await patchRegister(recId, { Status: "Cancelled", CancelReason: reason || "" });
      var subj = await oget(item.subject);
      var clean = String(subj || "").replace(new RegExp(String(num).replace(/[-]/g, "\\-") + "_?", "i"), "").trim();
      await oset(item.subject, clean);
      if (removeCc) {
        var cc = await oget(item.cc);
        var keepCc = cc.filter(function (r) { return String(r.emailAddress || "").toLowerCase() !== C.generalMailbox.toLowerCase(); });
        await new Promise(function (r) { item.cc.setAsync(keepCc, function () { r(); }); });
      }
      var hdr = await removeHeader(num);
      await deleteTask(num);
      show('<div class="ok"><b>Number released</b><div class="num" style="margin-top:4px">' + esc(num) + '</div><p class="muted">Marked Cancelled in the register (' + esc(reason || "") + ') and removed from the subject' + (hdr ? " and the email" : "") + '.</p></div>' +
           (hdr ? "" : '<div class="warn">Delete the Correspondence ID lines at the top of the email by hand.</div>') +
           '<button class="btns" id="again">Get a new number for this email</button>');
      $("#again").onclick = function () { composeView().catch(function (e) { showError("Something went wrong", e.message); }); };
      toast(num + " released");
    } catch (e) {
      toast("Not released: " + e.message, true);
      var go = $("#relGo"); if (go) { go.disabled = false; go.textContent = "Release"; }
    }
  }

  /* ---------- Read ---------- */
  async function readView() {
    var subject = item.subject || "";
    var from = (item.from && item.from.emailAddress) || "";
    var fromUs = isOwnDomain(domainOf(from));
    var m = subject.match(NUM_RE);
    var number = m ? m[0].toUpperCase() : null;
    var body = await bodyText();
    var hits = scan(subject + " " + body);
    var sug = suggest(hits);
    var refs = externalRefs(subject);
    var refText = refs.map(function (r) { return r.ref; }).join("; ");
    var rec = await findRecord({ messageId: item.internetMessageId, number: fromUs && number && number.indexOf(C.project + "-" + C.issuer + "-") === 0 ? number : null });
    if (rec) {
      var rf = rec.fields || {};
      if (refText && !rf.ExternalRef && rf.Direction !== "Out") {
        // fill in their reference on rows logged before this feature existed
        var patch = { ExternalRef: refText };
        if (!rf.Title || rf.Title === "(no number)") patch.Title = refs[0].ref;
        try { await patchRegister(rec.id, patch); Object.assign(rf, patch); } catch (e) { /* column may not exist yet */ }
      }
      recordPanel(rec, hits); return;
    }

    var recips = (item.to || []).concat(item.cc || []).map(function (r) { return r.emailAddress; });
    var dir = fromUs ? "Out" : "In";
    var code = (number && codeFromNumber(number)) || (fromUs ? codeForAddresses(recips) : codeForAddresses([from])) || (refs[0] && refs[0].code) || null;
    show(
      '<div class="warn"><b>Not in the register yet.</b><br><span class="muted">' + (dir === "In" ? "Incoming" : "Outgoing") + (code ? " · " + esc(code) : "") + (number ? " · ref " + esc(number) : "") + '</span>' +
      (refs.length ? '<div style="margin-top:6px;font-size:13px">Their ref: <b class="mono">' + esc(refs[0].ref) + '</b>' + (refs.length > 1 ? '<br><span class="muted">Refers to: ' + esc(refs.slice(1).map(function (r) { return r.ref; }).join(", ")) + '</span>' : "") + '</div>' : "") +
      '</div>' +
      (hits.length ? '<div><div class="lbl">Keywords found</div>' + hits.map(function (k) { return '<span class="kw">' + esc(k) + '</span>'; }).join("") + '</div>' : "") +
      '<label class="varbox"><input type="checkbox" id="var"' + (sug.variation ? " checked" : "") + '><span><b>Variation</b><br><span class="muted">Track for the Claims Manager</span></span></label>' +
      impactChecks(sug.impact) +
      '<button class="btnp" id="logBtn">Log to register</button>'
    );
    $("#logBtn").onclick = async function () {
      var btn = this; btn.disabled = true; btn.textContent = "Logging…";
      try {
        var created = await createRegister({
          Title: number || (refs[0] && refs[0].ref) || "(no number)", ExternalRef: refText, Direction: dir, Company: code ? companyChoice(code) : "Other",
          MailSubject: subject, MailDate: (item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date()).toISOString(),
          Status: "Logged", Variation: $("#var").checked, Keywords: hits.join(", "),
          "Impact@odata.type": "Collection(Edm.String)", Impact: readImpact(),
          Sender: from, Recipients: recips.join("; "), MessageId: item.internetMessageId || ""
        });
        var fresh = await g("GET", listPath("register") + "/items/" + created.id + "?$expand=fields");
        toast("Logged in the register");
        recordPanel(fresh, hits);
      } catch (e) { toast("Not logged: " + e.message, true); btn.disabled = false; btn.textContent = "Log to register"; }
    };
  }

  /* ---------- Footer (settings) ---------- */
  function footer(compose) {
    var rs = roam();
    var on = !(rs && rs.get("sorReminder") === false);
    var barOn = !(rs && rs.get("sorBar") === false);
    $("#foot").innerHTML = (compose ? '<label><input type="checkbox" id="remind"' + (on ? " checked" : "") + '> Remind me on Send for new emails to project parties</label>' +
      '<label><input type="checkbox" id="barTgl"' + (barOn ? " checked" : "") + '> Show the project-email bar while writing</label>' : "") +
      '<span>Signed in as ' + esc((Office.context.mailbox.userProfile || {}).emailAddress || "") + '</span>';
    var r = $("#remind");
    if (r) r.onchange = function () { if (rs) { rs.set("sorReminder", r.checked); saveRoam(); toast(r.checked ? "Send reminder on" : "Send reminder off"); } };
    var bt = $("#barTgl");
    if (bt) bt.onchange = function () {
      if (rs) { rs.set("sorBar", bt.checked); saveRoam(); toast(bt.checked ? "Project-email bar on" : "Project-email bar off"); }
      if (!bt.checked) { try { item.notificationMessages.removeAsync("sorekHint", function () {}); } catch (e) { /* none */ } }
    };
  }

  /* ---------- Start ---------- */
  async function start() {
    show('<p class="muted">Connecting to the register…</p>');
    try { await initAuth(); }
    catch (e) { showError("Sign-in could not start", e.message); return; }
    try { await token(false); }
    catch (e) {
      show('<div class="warn">Sign in once to connect the add-in to the SharePoint register.</div><button class="btnp" id="signin">Sign in</button>');
      $("#signin").onclick = async function () {
        try { await token(true); run(); }
        catch (err) { showError("Sign-in failed", err.message + (/consent|AADSTS65001|AADSTS90094/i.test(err.message) ? " — your organisation may require an administrator to approve the app." : "")); }
      };
      return;
    }
    run();
  }
  async function run() {
    show('<p class="muted">Loading the register…</p>');
    try { await loadData(); }
    catch (e) {
      var hint = e.status === 403 ? "You have no access to the SharePoint site, or the app permission was not approved." :
                 e.status === 404 ? "The SharePoint site was not found. Check sitePath in config.js." : "";
      showError("The register could not be loaded", e.message + (hint ? " " + hint : "")); return;
    }
    var compose = !!(item.subject && item.subject.getAsync);
    footer(compose);
    try { compose ? await composeView() : await readView(); }
    catch (e) { showError("Something went wrong", e.message); }
  }

  Office.onReady(function () {
    item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { show('<p class="muted">Open an email to use Sorek Register.</p>'); return; }
    start();
  });
})();
