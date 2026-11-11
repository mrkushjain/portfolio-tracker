# 📊 Portfolio Tracker

A **zero-backend, zero-subscription** personal stock portfolio bookkeeping app — hosted free on GitHub Pages, with Google Sheets + Apps Script as the backend API.

🚀 **Live App:** [mrkushjain.github.io/portfolio-tracker](https://mrkushjain.github.io/portfolio-tracker/)

---

## ✨ Features

### 🏗️ Fully Dynamic — Zero Hardcoding
- Create **Account Profiles** at runtime (e.g. Me, Wife, Parent)
- Add **Brokers / Platforms** dynamically and link them to accounts (Robinhood, Zerodha, Schwab, Fidelity, E\*TRADE, INDmoney, etc.)
- **Multi-currency** support: USD ($) and INR (₹) per broker
- Seed initial cash balances and existing stock positions from the UI — no spreadsheet editing needed

### 🐷 Virtual Piggy Banks / Goal Allocation
- Create custom goal buckets: *Shivya's Wedding*, *Goa Trip*, *Retirement*, *Buying House*, etc.
- **Earmark specific share quantities** from your holdings to individual goals (e.g. 5 of 10 MSFT shares → Wedding, 3 → House)
- Visual progress bars showing current market value vs. goal amount
- When selling a stock that has goal allocations, a prompt lets you specify which piggy bank(s) to deduct from

### 📋 Order Bookkeeping Log
- Log BUY / SELL orders with: symbol, quantity, price, fees, execution type (Market / Limit), tax term (Short/Long)
- **Atomic fulfillment state machine**: marking an order Fulfilled automatically:
  - BUY → deducts cash, adds holding, recalculates weighted average cost, auto-allocates to tagged piggy bank
  - SELL → deducts holding, adds cash, calculates realized P&L, deducts piggy bank allocations
- Cancel or delete Pending orders
- **📖 Book Stock** quick action: one click to pre-fill a new BUY order from any fulfilled order

### 📊 Analytics Dashboard
- Portfolio value, total cash, invested amount, and realized P&L summary cards
- Per-broker breakdown table
- Holdings overview with unrealized P&L (when live prices are fetched)
- Realized P&L chart by stock (Chart.js)
- Portfolio allocation donut chart by broker

### 👁️ Watchlist & Live Prices
- Add any stock ticker to a watchlist (supports US equities, ETFs, Indian stocks via `NSE:TCS`, etc.)
- Fetches live price, % change, open, high, low via Google Sheets' `GOOGLEFINANCE()` function
- Shows your existing holding quantity next to each watchlist item

---

## 🗂️ Tech Stack

| Layer | Technology |
|---|---|
| Frontend | Vanilla JS · HTML5 · [Tailwind CSS CDN](https://tailwindcss.com) · [Chart.js](https://www.chartjs.org/) |
| Backend / API | [Google Apps Script](https://script.google.com) Web App (`doGet` returning JSON) |
| Database | Google Sheets (8 sheets: Accounts, Brokers, Cash\_Balances, Holdings, Piggy\_Banks, Goal\_Allocations, Orders, Watchlist) |
| Hosting | GitHub Pages (static, free) |

---

## 🚀 Deployment Guide

### Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **new blank spreadsheet**.
2. Name it **Portfolio Tracker** (or anything you like).

### Step 2 — Add the Apps Script

1. In your Google Sheet, click **Extensions → Apps Script**.
2. Delete the placeholder `function myFunction() {}`.
3. Paste the entire contents of [`Code.gs`](./Code.gs) into the editor.
4. Click **Save** (Ctrl+S). Name the project *Portfolio Tracker*.

### Step 3 — Initialize Sheets

1. In the Apps Script editor, select **`initSheets`** from the function dropdown.
2. Click **▶ Run**. Accept the permissions prompt when asked.
3. Switch back to your spreadsheet — you should see 8 new sheets with styled headers.

### Step 4 — Deploy as a Web App

1. In the Apps Script editor: **Deploy → New deployment**.
2. Click the ⚙️ gear → **Web app**.
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy** and copy the **Web App URL** (`https://script.google.com/macros/s/…/exec`).

> **After any code change:** Deploy → Manage deployments → Edit → New version → Deploy. The URL stays the same.

### Step 5 — Configure the App

1. Open the live app (GitHub Pages URL below, or open `index.html` locally).
2. Click the **⚙ button** in the top-right corner (it will be red if unconfigured).
3. Paste your Web App URL and click **Save & Connect**.

> The URL is saved in your browser's `localStorage` — it never leaves your device and survives page refreshes.

### Step 6 — Host on GitHub Pages (optional fork)

If you fork this repo:

1. Go to your fork → **Settings → Pages**.
2. Source: **Branch: main / (root)** → **Save**.
3. Your app will be live at `https://<your-username>.github.io/portfolio-tracker/`.

---

## 📐 Data Model

```
Accounts ──< Brokers ──< Cash_Balances
                    └──< Holdings ──< Goal_Allocations >── Piggy_Banks
                    └──< Orders
Watchlist  (independent)
```

---

## 🔧 CORS Note

All write operations are sent as `GET ?payload=<JSON>` requests instead of `POST`. This avoids the GAS `POST /exec → 302 redirect → cross-origin /echo` chain that browsers block in CORS mode. The `doGet` handler in `Code.gs` detects the `payload` parameter and routes it to the same write handlers as `doPost`.

---

## 📄 License

MIT
