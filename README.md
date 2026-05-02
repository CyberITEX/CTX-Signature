# CTX-Signature — CyberITEX Outlook Email Signature Add-in

## Quick Reference URLs

| Resource | URL |
|---|---|
| **Manifest** | `https://cyberitex.github.io/CTX-Signature/manifest.xml` |
| **commands.js** | `https://cyberitex.github.io/CTX-Signature/commands.js` |
| **commands.html** | `https://cyberitex.github.io/CTX-Signature/commands.html` |
| **taskpane.html** | `https://cyberitex.github.io/CTX-Signature/taskpane.html` |
| **Logo (full)** | `https://cyberitex.github.io/CTX-Signature/assets/logo.png` |
| **Logo (32px)** | `https://cyberitex.github.io/CTX-Signature/assets/logo-32.png` |

> Use the **Manifest URL** when installing via Microsoft 365 Admin Center → Integrated Apps → Upload custom app.

---

An Microsoft Outlook Add-in that automatically inserts a static CyberITEX email signature whenever a new email or appointment is composed. Works on Outlook Web (OWA), New Outlook (Windows), and Classic Outlook (desktop).

---

## How It Works

```
User clicks New Mail / New Appointment
            ↓
Outlook fires OnNewMessageCompose event
            ↓
commands.js runs onNewMessageCompose()
            ↓
setSignatureAsync() injects the HTML signature
            ↓
Signature appears in the compose window automatically
```

The add-in uses **event-based activation** (Office.js Mailbox API 1.10). No button click is needed — the signature is inserted the moment the compose window opens.

---

## Project Files

### [`manifest.xml`](manifest.xml)
The configuration file that tells Outlook what this add-in is, where its files are hosted, and when to run it.

**Key sections:**

| Section | Purpose |
|---|---|
| `<Id>` | Unique GUID identifying this add-in in the Microsoft ecosystem |
| `<DisplayName>` | Name shown in Outlook's add-in list |
| `<IconUrl>` | 32×32 icon shown in the Outlook UI |
| `<AppDomains>` | Whitelist of domains the add-in is allowed to communicate with |
| `<FormSettings>` | Legacy fallback — loads `taskpane.html` in compose form for old Outlook clients |
| `<Permissions>` | `ReadWriteItem` — required to write the signature into the email body |
| `<Rule>` | Legacy rule — activates on message compose form (fallback for very old clients) |
| `VersionOverridesV1_0` | Adds a **Signature** button to the compose ribbon (fallback for Mailbox 1.3+ clients) |
| `VersionOverridesV1_1` | Event-based activation for modern Outlook (Mailbox 1.10+) — fires automatically |

**VersionOverrides layering:**
```
OfficeApp (base — Mailbox 1.1)
└── VersionOverridesV1_0 (ribbon button — Mailbox 1.3+)
    └── VersionOverridesV1_1 (auto-insert on compose — Mailbox 1.10+)
```
Outlook uses the deepest layer it supports. Modern clients use V1_1 (auto). Older clients fall back to the ribbon button. Very old clients use the legacy `FormSettings`.

**Runtimes (V1_1):**

| Runtime | File | Purpose |
|---|---|---|
| `WebViewRuntime.Url` | `taskpane.html` | Renders the taskpane UI (signature preview) |
| `JSRuntime.Url` | `commands.js` | Runs the event handler in a lightweight JS engine |

---

### [`commands.js`](commands.js)
The event handler script. Runs in a **JS-only runtime** — no browser DOM, no UI. Office.js is pre-loaded by Outlook automatically.

**What it does:**
1. Defines the `SIGNATURE_HTML` constant — the full HTML table for the email signature
2. Defines `onNewMessageCompose(event)` — called by Outlook on every new message/appointment compose
3. Calls `Office.context.mailbox.item.body.setSignatureAsync()` to inject the signature as HTML
4. Calls `event.completed()` to tell Outlook the handler finished successfully
5. Registers the function name via `Office.actions.associate()` so Outlook can find it by the name declared in the manifest

**Signature content (edit here to update):**
```
Lines 4–26  →  SIGNATURE_HTML  →  the HTML table with logo, name, title, phones, links
```

**To update the signature:** edit `SIGNATURE_HTML` between the backticks on lines 4–26, then upload `commands.js` to GitHub. Changes go live within ~1 minute — no redeployment needed.

---

### [`commands.html`](commands.html)
A minimal HTML shell page. It has no visible content — its only job is to load `Office.js` and `commands.js` in the correct order for the `FunctionFile` reference in the manifest.

**Why it exists:**  
The manifest's `FunctionFile` element must point to an HTML page (not a raw `.js` file). Outlook loads this page in a hidden browser context to initialise the add-in's JavaScript environment before any events fire.

**Load order:**
```
Outlook loads commands.html
    → loads office.js  (provides Office.* APIs)
    → loads commands.js  (registers onNewMessageCompose)
```

---

### [`taskpane.html`](taskpane.html)
The visible task pane panel shown when the user clicks the **Signature** button in the compose ribbon (V1_0 fallback). Displays a status message and a live preview of what the signature looks like.

**Contains:**
- A status banner confirming the signature is auto-inserted
- A rendered preview of the signature table (same design as `commands.js`)
- References `assets/logo.png` via a relative path (works because GitHub Pages serves from the repo root)

---

### [`assets/logo.png`](assets/logo.png)
Full-size CyberITEX logo used inside the email signature body. Displayed at 100px wide in the signature table. Served from GitHub Pages.

### [`assets/logo-32.png`](assets/logo-32.png)
32×32 pixel square icon used by Outlook's add-in UI (ribbon button, add-in list). Must be exactly 32×32.

### [`.nojekyll`](.nojekyll)
Empty file that disables Jekyll processing on GitHub Pages. Without it, GitHub Pages would ignore files and folders starting with `_` and may misprocess the project. Required for correct static file serving.

---

## Hosting

All files are hosted as a static site on **GitHub Pages**:

| File | Live URL |
|---|---|
| `manifest.xml` | `https://cyberitex.github.io/CTX-Signature/manifest.xml` |
| `commands.js` | `https://cyberitex.github.io/CTX-Signature/commands.js` |
| `commands.html` | `https://cyberitex.github.io/CTX-Signature/commands.html` |
| `taskpane.html` | `https://cyberitex.github.io/CTX-Signature/taskpane.html` |
| `assets/logo.png` | `https://cyberitex.github.io/CTX-Signature/assets/logo.png` |
| `assets/logo-32.png` | `https://cyberitex.github.io/CTX-Signature/assets/logo-32.png` |

No server, no backend, no build step. Push to `main` → GitHub Pages serves it within ~1 minute.

---

## Deployment (M365 Admin)

1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings → Integrated apps → Upload custom apps**
3. Select **Office Add-in** → upload `manifest.xml`
4. Assign to users → **Deploy**

| Client | Time to activate after deploy |
|---|---|
| Outlook on the Web | ~5 minutes |
| New Outlook (Windows) | ~5–30 minutes |
| Classic Outlook (desktop) | up to 24 hours |

---

## Updating the Signature

Only `commands.js` and `taskpane.html` need editing:

| File | What to edit |
|---|---|
| [`commands.js` lines 4–26](commands.js) | `SIGNATURE_HTML` — the actual signature injected into emails |
| [`taskpane.html` lines 45–65](taskpane.html) | Preview table — keep in sync with `commands.js` |

After editing, upload both files to GitHub. No redeployment in M365 Admin Center required.

---

## Validation

Run the official Microsoft manifest validator locally:
```bash
npx office-addin-manifest validate manifest.xml
```
Expected output: `The manifest is valid.`

---

## Compatibility

| Outlook client | Supported | Mechanism |
|---|---|---|
| Outlook on the Web (OWA) | Yes | Event-based (V1_1) |
| New Outlook for Windows | Yes | Event-based (V1_1) |
| Classic Outlook (Microsoft 365) | Yes | Event-based (V1_1) |
| Classic Outlook (older, Mailbox 1.3+) | Yes | Ribbon button (V1_0 fallback) |
| Outlook Mobile | No | Event-based activation not supported on mobile |
