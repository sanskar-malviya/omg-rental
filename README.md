# OMG Rental — Smart Rental Management

Rental-management web app for **OMG Rental, Indore**. It starts with Garba chaniya choli rentals and is built to grow into wedding, bridal and party collections. It is linked with OMG Salon.

Every physical piece (G = ghagra, C = choli, D = dupatta, J = jewellery) is tracked on its own. Bookings, deposits, returns, damage, late fees, refunds, sales, vouchers and reports all work from that item-level data.

## What's in this repo

| Path | What it is |
|---|---|
| `omg-rental-app/netlify-site/` | The web app. Deploy this folder to Netlify as-is. |
| `omg-rental-app/netlify-site/config.js` | The Google Apps Script Web App URL the app talks to. If left empty, the app runs in demo mode. |
| `omg-rental-app/google-apps-script/Code.gs` | The API. Paste it into the Google Sheet (Extensions → Apps Script). |
| `omg-rental-app/SETUP_GUIDE.md` | Step-by-step go-live guide. |
| `OMG_Rental_Prototype.html` | Standalone demo with sample data and a guided tour. Open it directly in Chrome; nothing is saved. |

## Architecture

```
Tablet / phone / laptop ──▶ Netlify (static app, index.html)
                               │  HTTPS, JSON
                               ▼
                         Google Apps Script web app (Code.gs)
                         · PIN login + 6 h sessions
                         · admin-only tables (Users, Settings, Packages, Collections)
                         · script lock + double-booking check on every new booking
                               │
                               ▼
                         Google Sheet "OMG Rental Database" (single source of truth)
                         16 tabs: Items, Customers, Bookings, Booking_Items, Payments,
                         Booking_Events, Returns, Return_Items, Sales, Vouchers, Packages,
                         Collections, Users, Settings, Item_History, Audit_Log
```

- **No deletes.** Bookings, payments and sales change status (cancelled / void). Payments, Item_History and Audit_Log are append-only.
- **Sync.** The app diffs its in-memory state against the last saved snapshot and saves only the changed rows. It retries when offline and refreshes from the sheet every 60 s when idle.
- **Double booking.** A piece is blocked from pickup until the return deadline plus a turnaround buffer. The server re-checks this inside `LockService` before saving any new booking.
- **Demo mode.** If no API URL is set, the app runs on built-in sample data, with a demo clock and a 19-step guided tour.

## Run locally

Open `omg-rental-app/netlify-site/index.html` through any static server, for example:

```bash
npx serve omg-rental-app/netlify-site
```

## Code map (`index.html`)

1. CSS design tokens and components
2. Utilities (dates, money, icons, garment illustrations)
3. Seed / sample data
4. Domain rules: `availability`, `quote`, `checkVoucher`, `lateFee`, `money`, `completeCore`
5. Roles, permissions and audit log
6. Screens (one `scr*` function per screen)
7. Live mode: `toTables` / `fromTables` (sheet mapping), sync engine, login
8. Router and boot

The in-app **Dev Handoff** screen documents the data model, state machines, the availability algorithm, money rules and edge cases for a future move to a full backend (for example PostgreSQL).
