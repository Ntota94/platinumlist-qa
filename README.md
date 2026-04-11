# Platinumlist QA & D-SAT System

Internal quality assurance and D-SAT validation tool for the Platinumlist GCC Customer Support team.

## Features
- **QA Entry** — Score agent interactions across 15 parameters with live KO detection
- **D-SAT Import** — Upload Intercom CSV and auto-filter rating 1–3 cases
- **Validation Queue** — Validate each D-SAT case with responsible party, issue type, and coaching steps
- **Agent View** — Monthly performance breakdown with parameter chart
- **Manager Dashboard** — Team ranking, C-SAT targets, and manual C-SAT input
- **Coaching Email** — Auto-generated coaching emails per agent
- **Settings** — Coordinator info, max transactions, done-by names

## Stack
- **Frontend**: Vanilla HTML/CSS/JS + Bootstrap 5 + Chart.js (hosted on GitHub Pages)
- **Backend**: Google Apps Script (connected via `doPost`)
- **Data**: Google Sheets (one spreadsheet, multiple tabs)

## Setup

### 1. Apps Script (backend)
1. Open your Google Sheet → **Extensions → Apps Script**
2. Paste `Code.gs` content and save
3. Run `setupSheet()` once to create the required tabs
4. **Deploy → New Deployment → Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the deployment URL

### 2. GitHub Pages (frontend)
1. Open `index.html`
2. Find the line `var SCRIPT_URL = "";` near the top of the `<script>` block
3. Paste your Apps Script deployment URL:
   ```js
   var SCRIPT_URL = "https://script.google.com/macros/s/YOUR_ID/exec";
   ```
4. Commit and push — GitHub Pages serves the file automatically

## Targets
| Metric | Target |
|--------|--------|
| QA Score | ≥ 90% |
| C-SAT after validation | ≥ 85% |

## Version
**v2.1.0** — April 2026 · Platinumlist GCC Team
