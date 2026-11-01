/* Sorek Register - Outlook task pane.
   Compose: give a new email a SOR number (central counter in SharePoint list "Companies"),
            write it to the subject, add the general mailbox to CC, log it in "EmailRegister".
   Read:    find the email in the register (or log it) and set Status / Variation / Impact / Claims notes. */
/* global Office, msal */
(function () {
  "use strict";
  var C = window.SOREK_CONFIG;
  var SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All"];
  var NUM_RE = /SOR-[A-Z]{3}-[A-Z]{3}-\d{4}/i;
  var REPLY_RE = /^\s*(re|fw|fwd|aw|wg)\s*:/i;
  var MSAL_URLS = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@4/lib/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js"
  ];
  var STATUS_OUT = ["Issued", "Follow-up", "Claim", "Cancelled"];
  var STATUS_IN = ["Logged", "Follow-up", "Claim"];

  var pca = null, siteId = null, listIds = {}, companies = [], keywords = [], item = null;
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
    var co = await items("companies", "Title,CompanyName,Domains,LastNumber");
    companies = co.map(function (x) {
      var f = x.fields || {};
      return { id: x.id, code: String(f.Title || "").trim().toUpperCase(), name: f.CompanyName || "", last: Number(f.LastNumber || 0),
               domains: String(f.Domains || "").split(/[;,\s]+/).map(function (d) { return d.trim().toLowerCase().replace(/^@/, ""); }).filter(Boolean) };
    }).filter(function (c) { return c.code && c.code !== C.issuer; }).sort(function (a, b) { return a.code.localeCompare(b.code); });
    var kw = await items("keywords", "Title,Category");
    keywords = kw.map(function (x) { return String((x.fields || {}).Title || "").trim(); }).filter(Boolean);
    // cache domains for the Send reminder
    var rs = roam();
    if (rs) {
      var map = {}; companies.forEach(function (c) { if (c.domains.length) map[c.code] = c.domains; });
      rs.set("sorDomains", JSON.stringify(map)); saveRoam();
    }
  }

  function scan(text) {
    var t = (text || "").toLowerCase(), hits = [];
    keywords.forEach(function (k) { if (t.indexOf(k.toLowerCase()) >= 0 && hits.indexOf(k) < 0) hits.push(k); });
    return hits;
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
  function companyChoice(code) { return C.companyChoices.indexOf(code) >= 0 ? code : "Other"; }
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
      var created = await g("POST", listPath("register") + "/items", { fields: Object.assign(buildFields(num), { Title: num }) });
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
      var a = await items("register", "Title,Direction,Status,Variation,Impact,ClaimsNotes,MailSubject,Company", "fields/MessageId eq '" + q(opts.messageId) + "'");
      if (a.length) return a[0];
    }
    if (opts.number) {
      var b = await items("register", "Title,Direction,Status,Variation,Impact,ClaimsNotes,MailSubject,Company", "fields/Title eq '" + q(opts.number) + "'");
      var out = b.filter(function (x) { return (x.fields || {}).Direction === "Out"; });
      if (out.length) return out[0];
    }
    return null;
  }

  /* ---------- Record panel (shared by read and compose) ---------- */
  function recordPanel(rec, hits) {
    var f = rec.fields || {};
    var dir = f.Direction || "In";
    var statuses = dir === "Out" ? STATUS_OUT : STATUS_IN;
    var impact = Array.isArray(f.Impact) ? f.Impact : [];
    var st = { status: f.Status || statuses[0], variation: !!f.Variation, impact: impact.slice() };
    show(
      '<div class="ok"><b>In the register</b><div class="num" style="margin-top:4px">' + esc(f.Title || "") + '</div>' +
      '<p class="muted">' + (dir === "Out" ? "Outgoing" : "Incoming") + (f.Company ? " · " + esc(f.Company) : "") + '</p></div>' +
      (hits && hits.length ? '<div><div class="lbl">Keywords found</div>' + hits.map(function (k) { return '<span class="kw">' + esc(k) + '</span>'; }).join("") + '</div>' : "") +
      '<div><div class="lbl">Status</div><div class="pills" id="stPills">' + statuses.map(function (s) { return '<button class="pill" data-st="' + esc(s) + '" aria-pressed="' + (s === st.status) + '">' + esc(s) + '</button>'; }).join("") + '</div></div>' +
      '<label class="varbox"><input type="checkbox" id="recVar"' + (st.variation ? " checked" : "") + '><span><b>Variation</b><br><span class="muted">Tracked by the Claims Manager</span></span></label>' +
      '<div><div class="lbl">Impact</div><div class="checks">' + ["Cost", "Time", "Scope"].map(function (k) { return '<label><input type="checkbox" data-imp="' + k + '"' + (impact.indexOf(k) >= 0 ? " checked" : "") + '> ' + k + '</label>'; }).join("") + '</div></div>' +
      '<div><div class="lbl">Claims notes</div><textarea id="recNotes">' + esc(f.ClaimsNotes || "") + '</textarea></div>' +
      '<button class="btnp" id="recSave">Save</button>'
    );
    $("#stPills").onclick = function (e) {
      var b = e.target.closest("[data-st]"); if (!b) return;
      st.status = b.dataset.st;
      document.querySelectorAll("#stPills .pill").forEach(function (p) { p.setAttribute("aria-pressed", String(p === b)); });
      if (st.status === "Claim") $("#recVar").checked = true;
    };
    $("#recSave").onclick = async function () {
      var btn = this; btn.disabled = true; btn.textContent = "Saving…";
      var imp = Array.prototype.slice.call(document.querySelectorAll("[data-imp]")).filter(function (c) { return c.checked; }).map(function (c) { return c.dataset.imp; });
      try {
        await g("PATCH", listPath("register") + "/items/" + rec.id + "/fields", {
          Status: st.status, Variation: $("#recVar").checked, ClaimsNotes: $("#recNotes").value,
          "Impact@odata.type": "Collection(Edm.String)", Impact: imp
        });
        toast("Saved in the register");
      } catch (e) { toast("Not saved: " + e.message, true); }
      btn.disabled = false; btn.textContent = "Save";
    };
  }

  /* ---------- Compose ---------- */
  async function composeView() {
    var subject = await oget(item.subject);
    var to = await oget(item.to), cc = await oget(item.cc);
    var addrs = to.concat(cc).map(function (r) { return r.emailAddress; });
    var existing = (subject || "").match(NUM_RE);
    if (existing) {
      var rec = await findRecord({ number: existing[0].toUpperCase() });
      if (rec) { recordPanel(rec, []); return; }
      show('<div class="warn">This email already has the number <b class="mono">' + esc(existing[0]) + '</b> in the subject, but it is not in the register.</div>');
      return;
    }
    var body = await bodyText();
    var hits = scan(subject + " " + body);
    var detected = codeForAddresses(addrs);
    var st = { code: detected || (companies[0] && companies[0].code), variation: hits.length > 0 };
    var isReply = REPLY_RE.test(subject || "");

    function render() {
      var co = companies.find(function (c) { return c.code === st.code; });
      var next = co ? fmt(co.code, co.last + 1) : "—";
      show(
        (isReply ? '<div class="warn">This looks like a reply or forward. Replies are not numbered; keep the original subject.</div>' : "") +
        '<div><div class="lbl">Recipient' + (detected ? " · detected from the address" : "") + '</div><div class="pills" id="coPills">' +
        companies.map(function (c) { return '<button class="pill" data-co="' + esc(c.code) + '" aria-pressed="' + (c.code === st.code) + '" title="' + esc(c.name) + '">' + esc(c.code) + '</button>'; }).join("") + '</div></div>' +
        '<div><div class="lbl">Next number</div><div class="num">' + esc(next) + '</div><p class="muted">Taken only when you click Insert. If someone else takes it first, you get the next one.</p></div>' +
        '<div><div class="lbl">Topic</div><input type="text" id="topic" value="' + esc(subject || "") + '" placeholder="Short topic of the email"></div>' +
        '<label class="varbox"><input type="checkbox" id="var"' + (st.variation ? " checked" : "") + '><span><b>Track as variation</b><br><span class="muted">' +
        (hits.length ? "Keywords found: " + esc(hits.join(", ")) : "Tick if this email may support a variation request or claim.") + '</span></span></label>' +
        '<button class="btnp" id="insert"' + (co ? "" : " disabled") + '>Insert number</button>'
      );
      $("#coPills").onclick = function (e) { var b = e.target.closest("[data-co]"); if (!b) return; st.code = b.dataset.co; st.variation = $("#var").checked; var t = $("#topic").value; render(); $("#topic").value = t; };
      $("#insert").onclick = doInsert;
    }

    async function doInsert() {
      var btn = $("#insert"); btn.disabled = true; btn.textContent = "Getting number…";
      var topic = $("#topic").value.trim() || "Topic";
      var variation = $("#var").checked;
      try {
        var res = await reserve(st.code, function (num) {
          return {
            Direction: "Out", Company: companyChoice(st.code), MailSubject: num + "_" + topic, MailDate: new Date().toISOString(),
            Status: "Issued", Variation: variation, Keywords: hits.join(", "),
            Sender: (Office.context.mailbox.userProfile || {}).emailAddress || "", Recipients: addrs.join("; ")
          };
        });
        await oset(item.subject, res.num + "_" + topic);
        var all = (await oget(item.to)).concat(await oget(item.cc)).map(function (r) { return String(r.emailAddress || "").toLowerCase(); });
        var ccAdded = false;
        if (all.indexOf(C.generalMailbox.toLowerCase()) < 0) {
          await new Promise(function (r) { item.cc.addAsync([C.generalMailbox], function () { r(); }); });
          ccAdded = true;
        }
        show(
          '<div class="ok"><b>Number inserted</b><div class="num" style="margin-top:4px">' + esc(res.num) + '</div></div>' +
          '<ul class="done"><li>Number added to the subject</li>' +
          '<li>' + (ccAdded ? "sorek@sorek.energy added to CC" : "sorek@sorek.energy already in CC") + '</li>' +
          '<li>Logged in the register as <b>Issued</b></li>' + (variation ? '<li>Marked <b>Variation</b></li>' : "") + '</ul>' +
          '<p class="muted">Finish the email and click Send as usual.</p>'
        );
      } catch (e) {
        toast("The number was not inserted: " + e.message, true);
        btn.disabled = false; btn.textContent = "Insert number";
      }
    }
    render();
  }

  /* ---------- Read ---------- */
  async function readView() {
    var subject = item.subject || "";
    var from = (item.from && item.from.emailAddress) || "";
    var fromUs = domainOf(from) === C.ownDomain || domainOf(from).endsWith("." + C.ownDomain);
    var m = subject.match(NUM_RE);
    var number = m ? m[0].toUpperCase() : null;
    var body = await bodyText();
    var hits = scan(subject + " " + body);
    var rec = await findRecord({ messageId: item.internetMessageId, number: fromUs && number && number.indexOf(C.project + "-" + C.issuer + "-") === 0 ? number : null });
    if (rec) { recordPanel(rec, hits); return; }

    var recips = (item.to || []).concat(item.cc || []).map(function (r) { return r.emailAddress; });
    var dir = fromUs ? "Out" : "In";
    var code = fromUs ? codeForAddresses(recips) : codeForAddresses([from]);
    show(
      '<div class="warn"><b>Not in the register yet.</b><br><span class="muted">' + (dir === "In" ? "Incoming" : "Outgoing") + (code ? " · " + esc(code) : "") + (number ? " · ref " + esc(number) : "") + '</span></div>' +
      (hits.length ? '<div><div class="lbl">Keywords found</div>' + hits.map(function (k) { return '<span class="kw">' + esc(k) + '</span>'; }).join("") + '</div>' : "") +
      '<label class="varbox"><input type="checkbox" id="var"' + (hits.length ? " checked" : "") + '><span><b>Variation</b><br><span class="muted">Track for the Claims Manager</span></span></label>' +
      '<button class="btnp" id="logBtn">Log to register</button>'
    );
    $("#logBtn").onclick = async function () {
      var btn = this; btn.disabled = true; btn.textContent = "Logging…";
      try {
        var created = await g("POST", listPath("register") + "/items", { fields: {
          Title: number || "(no number)", Direction: dir, Company: code ? companyChoice(code) : "Other",
          MailSubject: subject, MailDate: (item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date()).toISOString(),
          Status: dir === "Out" ? "Issued" : "Logged", Variation: $("#var").checked, Keywords: hits.join(", "),
          Sender: from, Recipients: recips.join("; "), MessageId: item.internetMessageId || ""
        } });
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
    $("#foot").innerHTML = (compose ? '<label><input type="checkbox" id="remind"' + (on ? " checked" : "") + '> Remind me on Send for new emails to project parties</label>' : "") +
      '<span>Signed in as ' + esc((Office.context.mailbox.userProfile || {}).emailAddress || "") + '</span>';
    var r = $("#remind");
    if (r) r.onchange = function () { if (rs) { rs.set("sorReminder", r.checked); saveRoam(); toast(r.checked ? "Send reminder on" : "Send reminder off"); } };
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
