# Auto Form Filler

Bulk website form submitter using Playwright. Automatically fills and submits contact forms on multiple websites.

## ✨ Features

- 🚀 Bulk processing — submit forms on 100+ websites
- 📂 Two input modes: 5-file mode or CSV/Excel mode
- 🔄 Auto-resume — skips already-processed websites
- 🎯 Smart form detection — tries homepage → contact page → newsletter
- 👤 Splits First/Last name if form needs it
- 📝 Handles Subject dropdown + textarea + text input
- 🍪 Auto closes popups and accepts cookies
- 💾 Saves results to JSON for tracking

## 📋 Requirements

- Node.js v18 or higher
- Windows / Mac / Linux
- Chrome browser

## 🛠️ Installation

```bash
# 1. Clone this repo
git clone https://github.com/huzikhan66/auto-form-filler.git
cd auto-form-filler

# 2. Install dependencies
npm install

# 3. Install Playwright browser
npx playwright install chromium