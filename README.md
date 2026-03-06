<p align="center">
  <img src="extension/icons/icon128.png" width="80" alt="Outlook Time Matrix" />
</p>

<h1 align="center">Outlook Time Matrix</h1>

<p align="center">
  Stop reacting. Start prioritizing.
</p>

<p align="center">
  <a href="https://chromewebstore.google.com"><img src="https://img.shields.io/badge/Chrome%20Web%20Store-available-brightgreen?logo=googlechrome" alt="Chrome Web Store" /></a>
  <img src="https://img.shields.io/badge/manifest-v3-orange" alt="Manifest V3" />
</p>

---

Most people live in their calendar — but a calendar alone doesn't tell you what *actually* matters. This extension overlays an **Eisenhower Matrix** directly onto your Outlook events, so you can see at a glance what to do now, what to schedule, and what to let go.

---

## Screenshots

<p align="center">
  <img src="tmp/ex/resized/ex1.png" width="48%" />
  <img src="tmp/ex/resized/ex2.png" width="48%" />
</p>
<p align="center">
  <img src="tmp/ex/resized/ex4.png" width="48%" />
  <img src="tmp/ex/resized/ex5.png" width="48%" />
</p>

---

## Features

| | |
|---|---|
| **Drag & drop** | Move events between Q1–Q4 and the sidebar |
| **Zoom** | Single-click a quadrant to expand it |
| **Finished tracking** | Checkbox marks events as done, synced to Outlook category color |
| **Per-account memory** | Placements and settings saved per Microsoft account and per calendar |
| **Open in Outlook** | ↗ button navigates an existing Outlook tab to the event |
| **Grid / List view** | Toggle between 2×2 matrix and 4×1 list |
| **2-week window** | Loads events from Monday of the current week + 14 days |

### Quadrants

|  | Urgent | Not Urgent |
|---|---|---|
| **Important** | Q1 · Do Now | Q2 · Schedule |
| **Not Important** | Q3 · Delegate | Q4 · Eliminate |

---

## Installation

> **Chrome Web Store** — search "Outlook Time Matrix" or use the link above.

To run locally:

1. Clone this repo
2. Copy `extension/src/config.example.js` → `extension/src/config.js` and fill in your Azure credentials
3. Go to `chrome://extensions` → Enable **Developer mode** → **Load unpacked** → select `extension/`

---

## Setup (for self-hosting)

You need a free **Azure App Registration** to authenticate with Microsoft Graph.

1. [Azure Portal](https://portal.azure.com) → App registrations → New registration
2. Authentication → Add platform → Web → paste the Redirect URI from the extension options page
3. API permissions → Microsoft Graph delegated: `User.Read`, `Calendars.ReadWrite`, `offline_access`
4. Copy **Client ID** and **Tenant ID** into `extension/src/config.js`

> If you install from the Chrome Web Store, no setup is needed — just sign in with your Microsoft account.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Platform | Chrome Extension Manifest V3 |
| UI | Vanilla JS + CSS (no framework, no build step) |
| Auth | OAuth 2.0 PKCE via `chrome.identity` |
| API | Microsoft Graph API (`/me/calendars`, `/me/calendars/{id}/calendarView`) |
| Storage | `chrome.storage.local` + `chrome.storage.sync` |

---

## Privacy Policy

Outlook Time Matrix does **not** collect, store, or transmit any personal data to external servers.

- OAuth tokens are stored locally in `chrome.storage.local` on your device only
- Calendar placements and settings are stored locally per account
- No analytics, no tracking, no third-party data sharing

---

