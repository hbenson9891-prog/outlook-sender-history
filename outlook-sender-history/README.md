# Sender History (Outlook add-in)

A task pane that shows your email history with whoever sent the message you have
open. Switch between mail *from* that sender and *all correspondence* (both
directions), see a quick count and first-contact date for conflict/intake
context, and click any result to open it. Pin it open and it refreshes
automatically as you click through your inbox.

It reads your mailbox through Microsoft Graph using your own Microsoft 365
sign-in. No data is sent to any third party, and there is no separate server: the
code runs inside Outlook and talks only to Microsoft.

---

## What's in this folder

```
manifest.xml      The add-in definition (what you sideload into Outlook)
taskpane.html     The panel shell
taskpane.css      Styling
taskpane.js       The logic  <-- paste your Azure client ID here
assets/           Icons
README.md         This file
```

There are exactly **two placeholders** to fill in:

1. `YOUR DOMAIN` (where you host the files): find-and-replace `REPLACE-WITH-YOUR-DOMAIN` in `manifest.xml`.
2. `YOUR AZURE CLIENT ID`: paste into `CONFIG.clientId` at the top of `taskpane.js`.

---

## Setup (about 20 to 30 minutes, one time)

### Step 1: Host the files on an HTTPS URL

Outlook will only load an add-in from `https://`. The free, no-cost path is
**GitHub Pages**:

1. Create a free GitHub account if you don't have one.
2. Create a new public repository, e.g. `outlook-sender-history`.
3. Upload all the files in this folder (keep the `assets` subfolder intact).
4. In the repo, go to **Settings > Pages**, set **Source** to `main` branch, root, and Save.
5. After a minute, GitHub gives you a URL like
   `https://YOURNAME.github.io/outlook-sender-history/`.

Your domain for the next steps is the part without `https://` and without a
trailing slash, e.g. `YOURNAME.github.io/outlook-sender-history`.

> Firm note: if IT prefers files not sit on public GitHub, host them on **Azure
> Static Web Apps** (also free tier) or any firm web server instead. The only
> requirement is HTTPS.

Now open `manifest.xml` and replace every `REPLACE-WITH-YOUR-DOMAIN` with your
domain. There are several occurrences. Re-upload the edited `manifest.xml`.

### Step 2: Register the app in Azure (this is what authorizes the Graph read)

1. Go to the [Azure portal](https://portal.azure.com) and open **App registrations**.
2. Click **New registration**.
   - **Name**: `Sender History`
   - **Supported account types**: *Accounts in any organizational directory and personal Microsoft accounts*.
   - Leave Redirect URI blank for now. Click **Register**.
3. Copy the **Application (client) ID** from the Overview page.
4. Go to **Manage > Authentication > Add a platform > Single-page application**, and add **two** redirect URIs:
   - `brk-multihub://YOURNAME.github.io` (just the origin, no subpath)
   - `https://YOURNAME.github.io/outlook-sender-history/taskpane.html`

   The first one (the `brk-multihub://` entry) is what tells Microsoft that
   Outlook is allowed to broker sign-in for your add-in. It is required.
5. Save.

You do **not** need to pre-add API permissions. The add-in requests `Mail.Read`
at runtime and you consent the first time you use it.

> Firm note: to restrict the add-in to firm accounts only, set Supported account
> types to single-tenant in step 2, and change `authority` in `taskpane.js` from
> `common` to your Tenant ID.

### Step 3: Drop in your client ID

Open `taskpane.js`, and at the top set:

```js
clientId: "the-application-client-id-you-copied",
```

Re-upload `taskpane.js`.

### Step 4: Sideload the add-in

**Outlook on the web** (easiest to test):
1. Open Outlook in a browser, click the **gear / Settings**.
2. Go to **General > Manage add-ins** (or **Get Add-ins**), choose
   **My add-ins > Add a custom add-in > Add from file**, and upload `manifest.xml`.

**New Outlook for Windows / classic desktop**: same idea, find **Get Add-ins >
My add-ins > Custom Addins > Add from file** and select `manifest.xml`.

### Step 5: Use it

1. Open any received email.
2. On the ribbon (or the `...` menu on an open message), click **Sender History**.
3. The first time, click **Sign in** and approve the `Mail.Read` consent.
4. Click the **pin** icon at the top of the task pane so it stays open. Now it
   updates on its own every time you select a different email.

---

## Confidentiality (worth reading, given the practice)

- Sign-in uses your own Microsoft 365 identity. The add-in never sees or stores a password.
- The only outbound calls are to `login.microsoftonline.com` (sign-in) and
  `graph.microsoft.com` (your mailbox), both Microsoft endpoints inside your tenant.
- Nothing is sent to Anthropic, to me, or to any outside analytics or server.
- The hosting location (GitHub Pages, etc.) only serves the static HTML/JS. Your
  email content does not pass through it; the code executes in Outlook and reads
  Graph directly.
- Scope is read-only (`Mail.Read`). The add-in cannot send, delete, or move mail.
- A short-lived cache of results lives in memory for the session only and clears when the pane closes.

If you later deploy firm-wide, lock the app registration to your tenant and have
IT push it through **Microsoft 365 admin center > Integrated apps** rather than
per-user sideloading.

---

## Things you may not have considered (and what I did about them)

Built in already:
- **Two directions.** A toggle at the top switches between *From them* (mail the
  sender sent you) and *All correspondence* (everything to or from that address,
  including what you sent them). In the All view each result is tagged Received
  or Sent.
- **Matter / conflict context.** A summary line shows how many messages exist
  with that contact, the received/sent split, and the date of first contact
  (or "earliest shown" for very long histories).
- **Searches the whole mailbox, not just the Inbox.** Results are pulled from
  every folder, so archived threads and your Sent items both show up.
- **Newest first, capped at 25 cards** so a chatty contact doesn't flood the pane
  (the context count still reflects the larger total).
- **Click a result to open that email.**
- **Attachment indicator** flags messages that carried files (handy when hunting
  for an earlier draft or exhibit).
- **Auto-refresh on selection** when pinned, plus a manual refresh button.
- **Edge cases handled:** nothing selected, items with no sender (drafts,
  meeting requests), first-ever contact, sign-in needed, and Graph errors.

Worth knowing:
- **How the two-way lookup works.** Graph cannot filter on recipient fields, so
  the add-in uses Graph keyword search (`from:` and `participants:`). It is
  reliable but matches by address keyword rather than an exact-equality filter,
  so an unusual address format could occasionally include or miss an edge case.
- **Old Outlook.** Outlook 2019 and earlier perpetual versions use a legacy
  web engine that does not support this modern sign-in. New Outlook, Microsoft
  365 Outlook, and Outlook on the web all work. If you need 2019 support, there's
  a documented fallback flow we can bolt on.
- **Mobile.** This targets desktop and web. Outlook mobile add-in support for
  this pattern is limited.

---

## Troubleshooting

- **Button doesn't appear:** confirm the manifest URLs all use your real domain
  and that the files actually load at those `https://` URLs in a browser.
- **"Sign in" loops or errors:** the most common cause is a missing or mistyped
  `brk-multihub://YOURDOMAIN` redirect URI in Azure, or the wrong client ID in
  `taskpane.js`.
- **Graph error about a query being too complex or recipients not filterable:**
  the code uses keyword search (`$search`) precisely to avoid this, so if you see
  it, you likely edited the query.
- **Nothing returns for a known contact:** in *From them* view, confirm you
  opened a *received* message (the lookup keys off the From address). In *All
  correspondence* view it keys off the address as a participant. Also confirm
  consent was granted.
