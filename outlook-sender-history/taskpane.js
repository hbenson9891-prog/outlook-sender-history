/* =====================================================================
   Sender History - taskpane.js

   >>> ONE THING TO EDIT: paste your Azure "Application (client) ID"
       into CONFIG.clientId below. <<<
   ===================================================================== */

const CONFIG = {
  // From the Azure app registration (Overview > Application (client) ID)
  clientId: "3b0c33ce-4b87-4734-9c05-7e23425d583b",

  // "common" = work/school + personal accounts.
  // To lock this to your firm only, replace "common" with your Tenant ID.
  authority: "https://login.microsoftonline.com/common",

  // Least-privilege scope: read-only mailbox access.
  graphScopes: ["Mail.Read"],

  maxResults: 25,   // how many message cards to show
  fetchSize: 100    // how many to pull from Graph for counts/first-contact
};

let msalInstance = null;
const cache = {};            // key "mode|email" -> { items, truncated }
let currentSender = null;    // lower-cased sender email of the open message
let currentMode = "from";    // "from" or "all"
let myEmail = "";            // the signed-in user's own address

/* ---------- Office bootstrap ---------- */
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  try {
    myEmail = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  } catch (e) { /* leave blank */ }

  document.getElementById("refresh-btn").addEventListener("click", () => {
    clearCacheForSender(currentSender);
    loadForCurrentItem(false);
  });

  document.getElementById("mode-from").addEventListener("click", () => setMode("from"));
  document.getElementById("mode-all").addEventListener("click", () => setMode("all"));

  // If the user pins the pane, refresh automatically when they select a
  // different message. (Requires Mailbox 1.5+, which our manifest targets.)
  try {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => loadForCurrentItem(false)
    );
  } catch (e) {
    // Older host without ItemChanged support: pane still works per message.
  }

  loadForCurrentItem(false);
});

function setMode(mode) {
  if (mode === currentMode) return;
  currentMode = mode;
  document.getElementById("mode-from").classList.toggle("active", mode === "from");
  document.getElementById("mode-all").classList.toggle("active", mode === "all");
  loadForCurrentItem(false);
}

function clearCacheForSender(email) {
  if (!email) return;
  delete cache["from|" + email];
  delete cache["all|" + email];
}

/* ---------- MSAL (Nested App Authentication) ---------- */
async function initMsal() {
  if (msalInstance) return;
  msalInstance = await msal.createNestablePublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: CONFIG.authority
    },
    cache: { cacheLocation: "localStorage" }
  });
}

/**
 * Try to get a Graph token silently. If interaction is required and the
 * caller allows it (i.e. user clicked a button), show the sign-in popup.
 * Returns the access token, or null if sign-in is needed but not yet allowed.
 */
async function getToken(allowInteractive) {
  await initMsal();
  const request = { scopes: CONFIG.graphScopes };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    if (allowInteractive) {
      const result = await msalInstance.acquireTokenPopup(request);
      return result.accessToken;
    }
    return null; // signal: caller should show a "Sign in" prompt
  }
}

/* ---------- Main flow ---------- */
async function loadForCurrentItem(allowInteractive) {
  const item = Office.context.mailbox.item;

  if (!item) {
    setSenderBar(null);
    showState("envelope", "No message selected", "Select an email in your inbox to see history for that sender.");
    return;
  }

  const from = item.from; // { emailAddress, displayName } in read mode
  if (!from || !from.emailAddress) {
    setSenderBar(null);
    showState("envelope", "No sender to look up", "This item (a draft, meeting, or report) has no sender address.");
    return;
  }

  const senderEmail = from.emailAddress.toLowerCase();
  currentSender = senderEmail;
  setSenderBar(from.displayName || from.emailAddress, from.emailAddress);

  const key = currentMode + "|" + senderEmail;
  if (cache[key]) {
    renderResult(cache[key]);
    return;
  }

  showSpinner(currentMode === "all"
    ? "Loading all correspondence\u2026"
    : "Loading prior emails\u2026");

  let token;
  try {
    token = await getToken(allowInteractive);
  } catch (err) {
    showSignInPrompt();
    return;
  }
  if (!token) {
    showSignInPrompt();
    return;
  }

  const modeAtRequest = currentMode;
  try {
    const result = await fetchMessages(token, senderEmail, modeAtRequest);
    cache[modeAtRequest + "|" + senderEmail] = result;
    // Guard against the user switching message/mode while we fetched.
    if (currentSender === senderEmail && currentMode === modeAtRequest) {
      renderResult(result);
    }
  } catch (err) {
    showState("warn", "Couldn't load emails", "", {
      errorText: (err && err.message) ? err.message : String(err),
      label: "Try again", onClick: () => loadForCurrentItem(true)
    });
  }
}

function showSignInPrompt() {
  showState("lock", "Sign in to continue",
    "Connect your Microsoft 365 account to look up this sender's history.",
    { label: "Sign in", onClick: () => loadForCurrentItem(true) });
}

/* ---------- Microsoft Graph ---------- */
/**
 * Uses $search (KQL), because Graph cannot $filter on recipient fields.
 *   mode "from" -> messages where this address is the sender
 *   mode "all"  -> any message where this address is a participant
 *                  (sender OR to/cc/bcc), i.e. full two-way correspondence
 */
async function fetchMessages(token, email, mode) {
  const safe = email.replace(/"/g, ""); // keep the KQL string valid
  const term = (mode === "all" ? "participants:" : "from:") + safe;
  const search = encodeURIComponent('"' + term + '"');
  const select = "subject,from,receivedDateTime,bodyPreview,webLink,hasAttachments,isRead";
  const url = "https://graph.microsoft.com/v1.0/me/messages"
            + "?$search=" + search
            + "&$select=" + select
            + "&$top=" + CONFIG.fetchSize;

  const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });

  if (!res.ok) {
    let detail = "HTTP " + res.status;
    try {
      const body = await res.json();
      if (body && body.error && body.error.message) detail = body.error.message;
    } catch (_) { /* ignore */ }
    throw new Error(detail);
  }

  const data = await res.json();
  const items = (data.value || []).sort(
    (a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime)
  );
  const truncated = !!data["@odata.nextLink"]; // more exist than we fetched
  return { items, truncated, mode };
}

/* ---------- Rendering ---------- */
function setSenderBar(name, email) {
  const bar = document.getElementById("sender-bar");
  const refresh = document.getElementById("refresh-btn");
  if (!name) {
    bar.hidden = true;
    refresh.hidden = true;
    return;
  }
  document.getElementById("sender-name").textContent = name;
  document.getElementById("sender-email").textContent = email;
  bar.hidden = false;
  refresh.hidden = false;
}

function directionOf(m) {
  const f = (m.from && m.from.emailAddress && m.from.emailAddress.address || "").toLowerCase();
  return (myEmail && f === myEmail) ? "sent" : "received";
}

function renderResult(result) {
  const content = document.getElementById("content");
  const items = result.items;

  if (!items.length) {
    const msg = result.mode === "all"
      ? "No emails to or from this address are in your mailbox yet."
      : "This appears to be the first message from this sender in your mailbox.";
    showState("envelope", "No prior emails", msg);
    return;
  }

  const frag = document.createDocumentFragment();
  frag.appendChild(buildContextLine(result));

  const display = items.slice(0, CONFIG.maxResults);
  const wrap = document.createElement("div");
  wrap.className = "msg-list";

  display.forEach((m, i) => {
    const card = document.createElement("div");
    const dir = directionOf(m);
    card.className = "msg-card" + (m.isRead === false ? " unread" : "");
    card.style.animationDelay = (i * 16) + "ms";
    if (m.webLink) {
      card.addEventListener("click", () => window.open(m.webLink, "_blank"));
    }

    const top = document.createElement("div");
    top.className = "msg-top";
    const subj = document.createElement("div");
    subj.className = "msg-subject";
    subj.textContent = m.subject || "(no subject)";
    const date = document.createElement("div");
    date.className = "msg-date";
    date.textContent = formatDate(m.receivedDateTime);
    top.appendChild(subj);
    top.appendChild(date);

    const preview = document.createElement("div");
    preview.className = "msg-preview";
    preview.textContent = m.bodyPreview || "";

    card.appendChild(top);
    card.appendChild(preview);

    const meta = document.createElement("div");
    meta.className = "msg-meta";
    // Show direction only in "all" mode (in "from" mode everything is received).
    if (result.mode === "all") {
      const d = document.createElement("span");
      d.className = "badge dir " + dir;
      d.textContent = dir === "sent" ? "\u2192 Sent" : "\u2190 Received";
      meta.appendChild(d);
    }
    if (m.hasAttachments) {
      const a = document.createElement("span");
      a.className = "badge";
      a.textContent = "\uD83D\uDCCE Attachment";
      meta.appendChild(a);
    }
    if (meta.childNodes.length) card.appendChild(meta);

    wrap.appendChild(card);
  });

  frag.appendChild(wrap);
  content.replaceChildren(frag);
}

/* Matter / conflict context summary */
function buildContextLine(result) {
  const items = result.items;
  const total = items.length;
  const totalLabel = total + (result.truncated ? "+" : "");

  let sent = 0, received = 0;
  let earliest = null;
  items.forEach((m) => {
    if (directionOf(m) === "sent") sent++; else received++;
    const d = new Date(m.receivedDateTime);
    if (!earliest || d < earliest) earliest = d;
  });

  const wrap = document.createElement("div");
  wrap.className = "context";

  const line1 = document.createElement("div");
  line1.className = "context-main";
  if (result.mode === "all") {
    line1.textContent = totalLabel + " message" + (total === 1 ? "" : "s")
      + " with this contact";
  } else {
    line1.textContent = totalLabel + " message" + (total === 1 ? "" : "s")
      + " from this sender";
  }
  wrap.appendChild(line1);

  const bits = [];
  if (result.mode === "all") {
    bits.push(received + " received");
    bits.push(sent + " sent");
  }
  if (earliest) {
    bits.push((result.truncated ? "earliest shown " : "first contact ")
      + formatDate(earliest.toISOString(), true));
  }
  if (bits.length) {
    const line2 = document.createElement("div");
    line2.className = "context-sub";
    line2.textContent = bits.join("  \u00b7  ");
    wrap.appendChild(line2);
  }

  if (total > CONFIG.maxResults || (result.truncated && total >= CONFIG.maxResults)) {
    const line3 = document.createElement("div");
    line3.className = "context-note";
    line3.textContent = "Showing the " + Math.min(CONFIG.maxResults, total) + " most recent";
    wrap.appendChild(line3);
  }

  return wrap;
}

/* ---------- State helpers ---------- */
function showSpinner(text) {
  const content = document.getElementById("content");
  const s = stateBlock("", "", text);
  const sp = document.createElement("div");
  sp.className = "spinner";
  s.insertBefore(sp, s.firstChild);
  content.replaceChildren(s);
}

function showState(icon, title, body, opts) {
  opts = opts || {};
  const glyphs = { envelope: "\u2709", lock: "\uD83D\uDD12", warn: "\u26A0", info: "\u2139" };
  const content = document.getElementById("content");
  const s = stateBlock(glyphs[icon] || "", title, body);

  if (opts.errorText) {
    const e = document.createElement("div");
    e.className = "error-text";
    e.textContent = opts.errorText;
    s.appendChild(e);
  }
  if (opts.label && opts.onClick) {
    const btn = document.createElement("button");
    btn.className = "primary-btn";
    btn.textContent = opts.label;
    btn.addEventListener("click", opts.onClick);
    s.appendChild(btn);
  }
  content.replaceChildren(s);
}

function stateBlock(glyph, title, body) {
  const s = document.createElement("div");
  s.className = "state";
  if (glyph) {
    const g = document.createElement("div");
    g.className = "big";
    g.textContent = glyph;
    s.appendChild(g);
  }
  if (title) {
    const h = document.createElement("h3");
    h.textContent = title;
    s.appendChild(h);
  }
  if (body) {
    const p = document.createElement("p");
    p.textContent = body;
    s.appendChild(p);
  }
  return s;
}

function formatDate(iso, withYear) {
  if (!iso) return "";
  const d = new Date(iso);
  const now = new Date();
  const sameYear = d.getFullYear() === now.getFullYear();
  const opts = (sameYear && !withYear)
    ? { month: "short", day: "numeric" }
    : { month: "short", day: "numeric", year: "numeric" };
  return d.toLocaleDateString(undefined, opts);
}
