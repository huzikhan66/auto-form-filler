// website-monitor.mjs — Combined Universal Version
// Supports: 5-file mode (urls.txt, first-name.txt, last-name.txt, subject.txt, message.txt)
//        OR CSV/Excel mode (urls.csv with url,email,name,subject,message columns)

import readline from "readline";
import axios from "axios";
import * as cheerio from "cheerio";
import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from 'url';
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ═══════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════
const CONFIG = {
  slowMo: 300,
  timeout: 30000,
  maxTabs: 1, // 1 tab at a time — safe, sequential
  resultsFile: path.join(__dirname, 'monitoring-results.json'),
  logsFile: path.join(__dirname, 'monitoring-logs.txt')
};

// ─── Default contact info (fallback if files don't have data) ───
const DEFAULT_CONTACT = {
  firstName: "Albert",
  lastName: "Khan",
  name: "Albert Khan",
  email: "albert.psestimation@gmail.com",
  phone: "",
  subject: "Custom Packaging Inquiry",
  message: `Hi,

We provide custom packaging products like rigid boxes, magnetic closure boxes, mailer boxes, drawer boxes, display boxes, custom inserts, window boxes, gift boxes, branded labels, stickers, hang tags, printed tissue paper, custom tape, eco paper bags, stand-up pouches, cardboard tubes, food packaging, subscription boxes, candle boxes, cosmetic boxes, and apparel packaging.

We offer affordable prices, short delivery times, and low MOQ.

Do you need any custom packaging material?`
};

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════
function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question(question, (answer) => { rl.close(); resolve(answer.trim()); });
  });
}

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }
function getTimestamp() { return new Date().toISOString(); }

async function logMessage(message) {
  const logEntry = `[${getTimestamp()}] ${message}\n`;
  console.log(message);
  try { await fs.appendFile(CONFIG.logsFile, logEntry); } catch {}
}

// ─── CSV line parser (respects quoted commas) ───
function splitCsvLine(line) {
  const parts = [];
  let current = "";
  let inQuotes = false;
  for (const ch of line) {
    if (ch === '"') inQuotes = !inQuotes;
    else if (ch === "," && !inQuotes) { parts.push(current.trim()); current = ""; }
    else current += ch;
  }
  parts.push(current.trim());
  return parts.map(p => p.replace(/^"|"$/g, "").trim());
}

// ─── File finder (tries multiple names) ───
async function findFileByPatterns(patterns) {
  for (const name of patterns) {
    const p = path.join(__dirname, name);
    try { await fs.access(p); return p; } catch {}
  }
  return null;
}

async function loadLines(filePath) {
  if (!filePath) return [];
  try {
    const raw = await fs.readFile(filePath, 'utf8');
    return raw.split(/\r?\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0 && !l.startsWith('#'));
  } catch { return []; }
}

function pickRotating(list, index, fallback) {
  if (!list || list.length === 0) return fallback;
  return list[index % list.length];
}

// ═══════════════════════════════════════════════════════════
// INPUT LOADING
// ═══════════════════════════════════════════════════════════

// MODE 1: 5 separate files
async function loadFiveFileMode() {
  const files = {
    firstName: await findFileByPatterns(['first-name.txt', 'firstname.txt', 'fname.txt', 'first_name.txt']),
    lastName:  await findFileByPatterns(['last-name.txt', 'lastname.txt', 'lname.txt', 'last_name.txt']),
    subject:   await findFileByPatterns(['subject.txt', 'subjects.txt']),
    message:   await findFileByPatterns(['message.txt', 'messages.txt', 'msg.txt']),
    urls:      await findFileByPatterns(['urls.txt', 'url.txt', 'urls.csv', 'websites.txt', 'links.txt'])
  };

  if (!files.urls) return null;

  const urlLines = await loadLines(files.urls);
  const firstNames = await loadLines(files.firstName);
  const lastNames = await loadLines(files.lastName);
  const subjects = await loadLines(files.subject);
  const messages = await loadLines(files.message);

  const tasks = [];
  for (let i = 0; i < urlLines.length; i++) {
    const cols = splitCsvLine(urlLines[i]);
    const url = cols[0];
    if (!url) continue;
    const email = (cols[1] || DEFAULT_CONTACT.email).trim();

    const firstName = pickRotating(firstNames, i, DEFAULT_CONTACT.firstName);
    const lastName = pickRotating(lastNames, i, DEFAULT_CONTACT.lastName);
    const subject = pickRotating(subjects, i, DEFAULT_CONTACT.subject);
    const message = messages.length > 0 ? pickRotating(messages, i, DEFAULT_CONTACT.message) : DEFAULT_CONTACT.message;

    tasks.push({
      url: url.trim(),
      email,
      firstName: firstName.trim(),
      lastName: lastName.trim(),
      name: `${firstName} ${lastName}`.trim(),
      phone: DEFAULT_CONTACT.phone,
      subject: subject.trim(),
      message
    });
  }

  console.log(`📂 5-File Mode:`);
  console.log(`   • URLs:       ${path.basename(files.urls)} (${urlLines.length} rows)`);
  console.log(`   • First Name: ${files.firstName ? path.basename(files.firstName) : '(default)'}`);
  console.log(`   • Last Name:  ${files.lastName ? path.basename(files.lastName) : '(default)'}`);
  console.log(`   • Subject:    ${files.subject ? path.basename(files.subject) : '(default)'}`);
  console.log(`   • Message:    ${files.message ? path.basename(files.message) : '(default)'}\n`);

  return tasks;
}

// MODE 2: single CSV/Excel file
async function findInputFile() {
  try {
    const files = await fs.readdir(__dirname);
    const candidates = files.filter(f => {
      const lower = f.toLowerCase();
      return (lower.endsWith('.csv') || lower.endsWith('.txt') || lower.endsWith('.xlsx') || lower.endsWith('.xls'))
        && !lower.startsWith('monitoring-logs')
        && !lower.startsWith('monitoring-results')
        && !lower.startsWith('~$')
        && !lower.match(/^(first|last|subject|message|urls?|fname|lname)/i); // avoid 5-file mode files
    });
    if (candidates.length === 0) return null;
    const preferred = ['input.xlsx', 'input.csv', 'input.txt', 'leads.xlsx', 'leads.csv'];
    for (const name of preferred) {
      const match = candidates.find(f => f.toLowerCase() === name);
      if (match) return path.join(__dirname, match);
    }
    const xlsx = candidates.find(f => /\.xlsx?$/i.test(f));
    return path.join(__dirname, xlsx || candidates[0]);
  } catch { return null; }
}

async function loadSingleFileMode(filePath) {
  let raw;
  if (/\.xlsx?$/i.test(filePath)) {
    const wb = XLSX.readFile(filePath);
    raw = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]], { blankrows: false });
  } else {
    raw = await fs.readFile(filePath, "utf8");
  }

  const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) return [];

  let startIndex = 0;
  let colIndex = { url: 0, email: 1, name: 2, phone: 3, subject: 4, message: 5 };
  const firstLineLower = lines[0].toLowerCase();
  if (firstLineLower.includes("url")) {
    const headerCols = splitCsvLine(lines[0]).map(c => c.toLowerCase());
    colIndex = {
      url: headerCols.indexOf("url"),
      email: headerCols.indexOf("email"),
      name: headerCols.indexOf("name"),
      phone: headerCols.indexOf("phone"),
      subject: headerCols.indexOf("subject"),
      message: headerCols.indexOf("message")
    };
    startIndex = 1;
  }

  const tasks = [];
  for (let i = startIndex; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i]);
    const url = colIndex.url >= 0 ? cols[colIndex.url] : cols[0];
    if (!url) continue;

    const email = (colIndex.email >= 0 ? cols[colIndex.email] : cols[1]) || DEFAULT_CONTACT.email;
    const name = (colIndex.name >= 0 ? cols[colIndex.name] : cols[2]) || DEFAULT_CONTACT.name;
    const phone = (colIndex.phone >= 0 ? cols[colIndex.phone] : cols[3]) || DEFAULT_CONTACT.phone;
    const subject = (colIndex.subject >= 0 ? cols[colIndex.subject] : cols[4]) || DEFAULT_CONTACT.subject;
    const message = (colIndex.message >= 0 ? cols[colIndex.message] : cols[5]) || DEFAULT_CONTACT.message;

    const nameParts = name.trim().split(/\s+/);
    tasks.push({
      url: url.trim(),
      email: email.trim(),
      name: name.trim(),
      firstName: nameParts[0] || DEFAULT_CONTACT.firstName,
      lastName: nameParts.slice(1).join(" ") || DEFAULT_CONTACT.lastName,
      phone: (phone || "").trim(),
      subject: (subject || "").trim(),
      message
    });
  }
  return tasks;
}

// ─── Resume support ───
async function loadAlreadyProcessedKeys() {
  const keys = new Set();
  try {
    const content = await fs.readFile(CONFIG.resultsFile, 'utf8');
    const data = JSON.parse(content);
    for (const r of data) {
      if (r.url) keys.add(`${r.url}::${r.email || ''}`);
    }
  } catch {}
  return keys;
}

// ═══════════════════════════════════════════════════════════
// WEBSITE TYPE DETECTION
// ═══════════════════════════════════════════════════════════
function detectWebsiteType(html) {
  const $ = cheerio.load(html);
  const pageText = $("html").html()?.toLowerCase() || "";

  if (pageText.includes("wp-content") || pageText.includes("wordpress") || pageText.includes("/wp-")) {
    return { type: "WORDPRESS", confidence: "HIGH" };
  }
  if (pageText.includes("cdn.shopify.com") || pageText.includes("shopify")) {
    return { type: "SHOPIFY", confidence: "HIGH" };
  }
  if ($('form').length > 0) {
    const hasInput = $('input[type="text"], input[type="email"], textarea').length > 0;
    const hasSubmit = $('button[type="submit"], input[type="submit"]').length > 0;
    if (hasInput && hasSubmit) return { type: "GENERIC_HTML_FORM", confidence: "HIGH" };
  }
  return { type: "UNKNOWN", confidence: "LOW" };
}

// ═══════════════════════════════════════════════════════════
// POPUP / COOKIE HANDLING
// ═══════════════════════════════════════════════════════════
async function handlePopupsAndCookies(page) {
  let popupClosed = false;
  let cookiesAccepted = false;

  await wait(1500);

  const popupSelectors = [
    'button[aria-label="Close"]', 'button[aria-label="Close popup"]',
    'button[aria-label="Close dialog"]', 'button[aria-label="Dismiss"]',
    '.close-button', '.popup-close', '.modal-close', '.popup-close-btn',
    'button.popup-close', 'button.modal-close', 'button.close',
    '.close-popup', '.dismiss-button', '.newsletter-close',
    '.popup .close', '.modal .close',
    '[data-dismiss="modal"]', '[data-close="true"]',
    '.fancybox-close', '.mfp-close', '.remodal-close',
    '.klaviyo-close-form', '.needsclick[aria-label="Close"]',
    '.privy-close-icon', '.om-close-icon',
    '.sumome-signup-form-close', '.justuno-close', '.mc-closeModal',
    'button:has-text("×")', 'button:has-text("✕")'
  ];

  for (const selector of popupSelectors) {
    try {
      const elements = page.locator(selector);
      const count = await elements.count();
      for (let i = 0; i < count; i++) {
        const el = elements.nth(i);
        if (await el.isVisible()) {
          await el.click({ timeout: 2000 });
          popupClosed = true;
          await wait(500);
          break;
        }
      }
      if (popupClosed) break;
    } catch {}
  }

  if (!popupClosed) {
    try { await page.keyboard.press('Escape'); await wait(300); } catch {}
  }

  const cookieSelectors = [
    'button:has-text("Accept")', 'button:has-text("Accept All")',
    'button:has-text("Accept Cookies")', 'button:has-text("Allow Cookies")',
    'button:has-text("OK")', 'button:has-text("Got it")',
    'button:has-text("I Agree")', 'button:has-text("Agree")',
    'button:has-text("Consent")', 'button:has-text("Yes")',
    '.cookie-accept', '.cookie-consent-accept', '.cookie-accept-button',
    '.accept-cookies', '#cookie-accept', '.cc-accept', '.cc-allow',
    '.gdpr-accept', '.cky-btn-accept', '.cky-btn-allow',
    '.onetrust-accept-btn-handler', '.optanon-allow-all',
    '[aria-label="Accept cookies"]', '[aria-label="Accept all cookies"]'
  ];

  for (const selector of cookieSelectors) {
    try {
      const elements = page.locator(selector);
      const count = await elements.count();
      for (let i = 0; i < count; i++) {
        const el = elements.nth(i);
        if (await el.isVisible()) {
          await el.click({ timeout: 2000 });
          cookiesAccepted = true;
          await wait(500);
          break;
        }
      }
      if (cookiesAccepted) break;
    } catch {}
  }

  return { popupClosed, cookiesAccepted };
}

// ═══════════════════════════════════════════════════════════
// FORM FILLING
// ═══════════════════════════════════════════════════════════

async function findVisibleLocator(page, selectors) {
  for (const selector of selectors) {
    try {
      const candidates = page.locator(selector);
      const count = await candidates.count();
      for (let i = 0; i < count; i++) {
        const c = candidates.nth(i);
        if (await c.isVisible().catch(() => false)) return c;
      }
    } catch {}
  }
  return null;
}

async function fillNameFields(page, tabIndex, contact) {
  const firstSelectors = [
    'input[name*="first" i][name*="name" i]', 'input[name*="fname" i]',
    'input[id*="first" i][id*="name" i]', 'input[placeholder*="first name" i]',
    'input[name="first_name"]', 'input[name="firstName"]', 'input[name="your-first-name"]'
  ];
  const lastSelectors = [
    'input[name*="last" i][name*="name" i]', 'input[name*="lname" i]',
    'input[id*="last" i][id*="name" i]', 'input[placeholder*="last name" i]',
    'input[name="last_name"]', 'input[name="lastName"]', 'input[name="your-last-name"]'
  ];

  const firstInput = await findVisibleLocator(page, firstSelectors);
  const lastInput = await findVisibleLocator(page, lastSelectors);

  if (!firstInput && !lastInput) return { filled: false };

  if (firstInput) {
    await firstInput.scrollIntoViewIfNeeded();
    await wait(300);
    await firstInput.fill(contact.firstName, { timeout: 5000 });
    await logMessage(`✅ Tab ${tabIndex}: Filled First Name`);
  }
  if (lastInput) {
    await lastInput.scrollIntoViewIfNeeded();
    await wait(300);
    await lastInput.fill(contact.lastName, { timeout: 5000 });
    await logMessage(`✅ Tab ${tabIndex}: Filled Last Name`);
  }
  return { filled: true };
}

async function selectFirstVisibleOption(page, tabIndex, selectors) {
  for (const selector of selectors) {
    try {
      const dropdowns = page.locator(selector);
      const count = await dropdowns.count();
      for (let i = 0; i < count; i++) {
        const dd = dropdowns.nth(i);
        if (!(await dd.isVisible().catch(() => false))) continue;
        const options = await dd.locator('option').allTextContents();
        const values = await dd.locator('option').evaluateAll(els => els.map(e => e.value));
        const preferredWords = ['general', 'other', 'sales', 'inquiry', 'product', 'quote'];
        let idx = -1;
        for (const word of preferredWords) {
          const found = options.findIndex(o => o && o.toLowerCase().includes(word));
          if (found >= 0) { idx = found; break; }
        }
        if (idx === -1) idx = options.findIndex((o, k) => o && o.trim() && values[k]);
        if (idx === -1 || !values[idx]) continue;
        await dd.selectOption(values[idx]);
        await logMessage(`✅ Tab ${tabIndex}: Selected dropdown "${options[idx].trim()}"`);
        return true;
      }
    } catch {}
  }
  return false;
}

async function findAndFillForm(page, tabIndex, contact) {
  const results = {
    nameFilled: false, emailFilled: false, phoneFilled: false,
    subjectFilled: false, messageFilled: false, submitted: false, formFound: false
  };

  await wait(1500);

  const formCount = await page.locator('form').count();
  if (formCount === 0) {
    await logMessage(`⚠️ Tab ${tabIndex}: No form found`);
    return results;
  }

  results.formFound = true;
  await logMessage(`📝 Tab ${tabIndex}: Found ${formCount} form(s)`);

  async function fillFirstVisible(selectors, value, label) {
    for (const selector of selectors) {
      const candidates = page.locator(selector);
      const count = await candidates.count();
      for (let i = 0; i < count; i++) {
        const c = candidates.nth(i);
        try {
          if (await c.isVisible()) {
            await c.scrollIntoViewIfNeeded();
            await wait(400);
            await c.focus();
            await c.fill(value, { timeout: 5000 });
            await wait(400);
            await logMessage(`✅ Tab ${tabIndex}: Filled ${label}`);
            return true;
          }
        } catch {}
      }
    }
    await logMessage(`⚠️ Tab ${tabIndex}: No ${label} field found`);
    return false;
  }

  // Name
  const splitName = await fillNameFields(page, tabIndex, contact);
  if (splitName.filled) results.nameFilled = true;
  else {
    results.nameFilled = await fillFirstVisible([
      'input[name*="name" i]:not([name*="user" i])',
      'input[id*="name" i]', 'input[placeholder*="name" i]',
      'input[name*="fullname" i]', 'input[name*="your-name" i]',
      'input[name="name"]', 'input[id="name"]'
    ], contact.name, "Name");
  }

  // Email
  results.emailFilled = await fillFirstVisible([
    'input[type="email"]', 'input[name*="email" i]', 'input[id*="email" i]',
    'input[placeholder*="email" i]', 'input[name*="your-email" i]',
    'input[name="email"]', 'input[id="email"]'
  ], contact.email, "Email");

  // Phone
  if (contact.phone) {
    results.phoneFilled = await fillFirstVisible([
      'input[type="tel"]', 'input[name*="phone" i]', 'input[id*="phone" i]',
      'input[placeholder*="phone" i]', 'input[name*="mobile" i]'
    ], contact.phone, "Phone");
  }

  // Subject
  if (contact.subject) {
    results.subjectFilled = await fillFirstVisible([
      'input[name*="subject" i]', 'input[id*="subject" i]',
      'input[placeholder*="subject" i]', 'input[name="your-subject"]',
      'input[name*="topic" i]', 'textarea[name*="subject" i]'
    ], contact.subject, "Subject");

    if (!results.subjectFilled) {
      results.subjectFilled = await selectFirstVisibleOption(page, tabIndex, [
        'select[name*="subject" i]', 'select[id*="subject" i]',
        'select[name*="topic" i]', 'select[name*="reason" i]',
        'select[name*="inquiry" i]'
      ]);
    }
  }

  // Message
  results.messageFilled = await fillFirstVisible([
    'textarea[name*="message" i]', 'textarea[id*="message" i]',
    'textarea[name*="comment" i]', 'textarea',
    'textarea[name="your-message"]', 'textarea[name="message"]'
  ], contact.message, "Message");

  // Submit
  try {
    const submitSelectors = [
      'button[type="submit"]', 'input[type="submit"]',
      'button:has-text("Submit")', 'button:has-text("Send")',
      'button:has-text("Send Message")', 'button:has-text("Contact")',
      'button:has-text("Get in Touch")', 'input[value*="Submit" i]',
      'input[value*="Send" i]', 'button[name="submit"]', 'input[name="submit"]'
    ];
    for (const selector of submitSelectors) {
      const btn = page.locator(selector).first();
      if (await btn.count() > 0 && await btn.isVisible()) {
        await btn.scrollIntoViewIfNeeded();
        await wait(800);
        await btn.click({ timeout: 5000 });
        await wait(1500);
        await logMessage(`✅ Tab ${tabIndex}: Clicked submit`);
        results.submitted = true;
        break;
      }
    }
  } catch (err) {
    await logMessage(`❌ Tab ${tabIndex}: Submit error: ${err.message}`);
  }

  return results;
}

// ═══════════════════════════════════════════════════════════
// CONFIRM SUBMISSION
// ═══════════════════════════════════════════════════════════
async function confirmSubmission(page, originalUrl, tabIndex) {
  await wait(2000);
  const currentUrl = page.url();
  if (currentUrl !== originalUrl) {
    await logMessage(`✅ Tab ${tabIndex}: Page redirected - confirmed!`);
    return { confirmed: true, method: "redirect" };
  }
  try {
    const bodyText = await page.textContent('body');
    const keywords = ['thank', 'success', 'sent', 'received', 'submitted', 'message sent', 'thanks for contacting'];
    for (const word of keywords) {
      if (bodyText.toLowerCase().includes(word)) {
        await logMessage(`✅ Tab ${tabIndex}: Success message found: "${word}"`);
        return { confirmed: true, method: "success_message" };
      }
    }
  } catch {}
  await logMessage(`⚠️ Tab ${tabIndex}: No confirmation found`);
  return { confirmed: false, method: "no_confirmation" };
}

// ═══════════════════════════════════════════════════════════
// CONTACT PAGE FALLBACK
// ═══════════════════════════════════════════════════════════
const CONTACT_PATHS = [
  "/contact", "/contact-us", "/contact-us/", "/contactus",
  "/pages/contact", "/pages/contact-us", "/get-in-touch",
  "/contact.html", "/contact.php", "/about/contact", "/support/contact"
];

async function findContactLinkOnPage(page, baseUrl) {
  const selectors = [
    'a:has-text("Contact Us")', 'a:has-text("Contact us")',
    'a:has-text("Contact")', 'a:has-text("Get in Touch")',
    'a[href*="contact" i]'
  ];
  for (const sel of selectors) {
    try {
      const link = page.locator(sel).first();
      if (await link.count() > 0 && await link.isVisible().catch(() => false)) {
        const href = await link.getAttribute('href');
        if (href) return new URL(href, baseUrl).toString();
      }
    } catch {}
  }
  return null;
}

async function tryContactPageFallback(page, tabIndex, baseUrl) {
  let candidateUrl = null;
  try { candidateUrl = await findContactLinkOnPage(page, baseUrl); } catch {}

  const candidates = candidateUrl
    ? [candidateUrl, ...CONTACT_PATHS.map(p => new URL(p, baseUrl).toString())]
    : CONTACT_PATHS.map(p => new URL(p, baseUrl).toString());

  const tried = new Set();
  for (const candidate of candidates.slice(0, 6)) {
    if (tried.has(candidate) || candidate === page.url()) continue;
    tried.add(candidate);
    try {
      await logMessage(`↪️ Tab ${tabIndex}: Trying contact page: ${candidate}`);
      const response = await page.goto(candidate, { waitUntil: "load", timeout: 15000 });
      if (!response || !response.ok()) continue;
      await wait(1000);
      const formCount = await page.locator('form').count();
      if (formCount > 0) {
        await logMessage(`✅ Tab ${tabIndex}: Found contact page with form`);
        return true;
      }
    } catch {}
  }
  return false;
}

// ═══════════════════════════════════════════════════════════
// NEWSLETTER FALLBACK
// ═══════════════════════════════════════════════════════════
async function tryNewsletterSignup(page, tabIndex, email) {
  const result = { attempted: false, submitted: false };
  try {
    const forms = page.locator('form');
    const formCount = await forms.count();
    for (let f = 0; f < formCount; f++) {
      const form = forms.nth(f);
      try {
        const hasTextarea = await form.locator('textarea').count();
        const hasName = await form.locator('input[name*="name" i], input[placeholder*="name" i]').count();
        if (hasTextarea > 0 || hasName > 0) continue;

        const emailInput = form.locator('input[type="email"], input[name*="email" i], input[placeholder*="email" i]').first();
        if (await emailInput.count() === 0) continue;
        if (!(await emailInput.isVisible().catch(() => false))) continue;

        result.attempted = true;
        await emailInput.scrollIntoViewIfNeeded();
        await wait(300);
        await emailInput.fill(email, { timeout: 5000 });
        await wait(500);

        const btn = form.locator(
          'button[type="submit"], input[type="submit"], button:has-text("Subscribe"), button:has-text("Sign Up"), button:has-text("Join")'
        ).first();

        if (await btn.count() > 0 && await btn.isVisible().catch(() => false)) {
          await btn.scrollIntoViewIfNeeded();
          await wait(500);
          await btn.click({ timeout: 5000 });
          await wait(1000);
          result.submitted = true;
          await logMessage(`✅ Tab ${tabIndex}: Newsletter box submitted`);
          return result;
        }
      } catch {}
    }
  } catch {}
  return result;
}

// ═══════════════════════════════════════════════════════════
// MAIN PROCESSING
// ═══════════════════════════════════════════════════════════
async function processWebsiteInTab(page, task, tabIndex, total) {
  const { url } = task;
  const result = {
    url, email: task.email, timestamp: getTimestamp(),
    status: "PENDING", websiteType: null, httpStatus: null,
    formSubmission: null, error: null, submissionConfirmed: false,
    popupClosed: false, cookiesAccepted: false, tabIndex
  };

  try {
    await logMessage(`\n📑 Tab ${tabIndex}/${total}: Opening ${url} (email: ${task.email})`);
    await page.goto(url, { waitUntil: "load", timeout: 30000 });
    await logMessage(`✅ Tab ${tabIndex}: Page loaded`);
    await wait(2000);

    try {
      const response = await axios.get(url, {
        timeout: 10000, validateStatus: false,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36' }
      });
      result.httpStatus = response.status;
    } catch {}

    try {
      const html = await page.content();
      const typeInfo = detectWebsiteType(html);
      result.websiteType = typeInfo;
      await logMessage(`📋 Tab ${tabIndex}: Type: ${typeInfo.type}`);
    } catch {}

    const popup = await handlePopupsAndCookies(page);
    result.popupClosed = popup.popupClosed;
    result.cookiesAccepted = popup.cookiesAccepted;
    if (popup.popupClosed) await logMessage(`✅ Tab ${tabIndex}: Popup closed`);
    if (popup.cookiesAccepted) await logMessage(`✅ Tab ${tabIndex}: Cookies accepted`);

    // Scroll
    await page.evaluate(() => window.scrollTo({ top: document.body.scrollHeight / 2, behavior: "smooth" }));
    await wait(800);
    await page.evaluate(() => window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" }));
    await wait(800);
    await page.evaluate(() => window.scrollTo({ top: 0, behavior: "smooth" }));
    await wait(800);

    let formResult = await findAndFillForm(page, tabIndex, task);
    let usedContactPage = false;

    if (!formResult.formFound || !formResult.submitted) {
      const found = await tryContactPageFallback(page, tabIndex, url);
      if (found) {
        usedContactPage = true;
        const popup2 = await handlePopupsAndCookies(page);
        result.popupClosed = result.popupClosed || popup2.popupClosed;
        result.cookiesAccepted = result.cookiesAccepted || popup2.cookiesAccepted;

        await page.evaluate(() => window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" }));
        await wait(800);
        await page.evaluate(() => window.scrollTo({ top: 0, behavior: "smooth" }));
        await wait(800);

        const contactResult = await findAndFillForm(page, tabIndex, task);
        if (contactResult.formFound) formResult = contactResult;
      }
    }

    result.formSubmission = formResult;
    result.contactPageUsed = usedContactPage;

    let newsletterResult = { attempted: false, submitted: false };
    if (!formResult.submitted) {
      newsletterResult = await tryNewsletterSignup(page, tabIndex, task.email);
      result.newsletterFallback = newsletterResult;
    }

    if (formResult.submitted) {
      const conf = await confirmSubmission(page, page.url(), tabIndex);
      result.submissionConfirmed = conf.confirmed;
      if (conf.confirmed) {
        result.status = "SUCCESS";
        await logMessage(`🎉✅ Tab ${tabIndex}: SUBMISSION CONFIRMED!`);
      } else {
        result.status = "SUBMITTED_BUT_UNCONFIRMED";
        await logMessage(`⚠️ Tab ${tabIndex}: Submitted but not confirmed`);
      }
    } else if (newsletterResult.submitted) {
      result.status = "NEWSLETTER_SUBMITTED";
      await logMessage(`📧 Tab ${tabIndex}: Newsletter submitted`);
    } else {
      result.status = "NO_SUBMIT";
      await logMessage(`⚠️ Tab ${tabIndex}: No form/submit`);
    }
  } catch (error) {
    result.status = "ERROR";
    result.error = error.message;
    await logMessage(`❌ Tab ${tabIndex}: Error: ${error.message}`);
  }

  return result;
}

// ═══════════════════════════════════════════════════════════
// SAVE RESULTS
// ═══════════════════════════════════════════════════════════
async function saveResults(newResults) {
  try {
    let existingData = [];
    try {
      const content = await fs.readFile(CONFIG.resultsFile, 'utf8');
      existingData = JSON.parse(content);
    } catch {}
    existingData.push(...newResults);
    if (existingData.length > 1000) existingData = existingData.slice(-1000);
    await fs.writeFile(CONFIG.resultsFile, JSON.stringify(existingData, null, 2));
    await logMessage(`💾 Results saved (${newResults.length} new)`);
  } catch (error) {
    await logMessage(`❌ Save error: ${error.message}`);
  }
}

// ═══════════════════════════════════════════════════════════
// ROLLING POOL
// ═══════════════════════════════════════════════════════════
async function runPool(tasks, browser, maxConcurrent) {
  const results = [];
  let nextIndex = 0;
  let active = 0;

  return new Promise((resolve) => {
    const launchNext = async () => {
      if (nextIndex >= tasks.length) {
        if (active === 0) resolve(results);
        return;
      }

      const task = tasks[nextIndex];
      const tabIndex = nextIndex + 1;
      nextIndex++;
      active++;

      await logMessage(`\n🆕 Tab ${tabIndex}/${tasks.length}: Launching (active: ${active})`);

      let page;
      try {
        page = await browser.newPage();
        const result = await processWebsiteInTab(page, task, tabIndex, tasks.length);
        await page.close();
        await logMessage(`🔒 Tab ${tabIndex}: Closed`);
        results.push(result);
        await saveResults([result]);
      } catch (error) {
        try { if (page) await page.close(); } catch {}
        const errResult = {
          url: task.url, email: task.email, timestamp: getTimestamp(),
          status: "ERROR", error: error.message, tabIndex
        };
        results.push(errResult);
        await saveResults([errResult]);
      }

      const ok = results.filter(r => r.status === "SUCCESS").length;
      const err = results.filter(r => r.status === "ERROR").length;
      await logMessage(`📊 Progress: ${results.length}/${tasks.length} | ${ok} ✅ | ${err} ❌`);

      active--;
      launchNext();
    };

    for (let i = 0; i < maxConcurrent; i++) launchNext();
  });
}

// ═══════════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════════
async function main() {
  console.log("═══════════════════════════════════════════════════════════");
  console.log("     🔍 WEBSITE MONITOR — UNIVERSAL COMBINED VERSION");
  console.log("     📦 Supports 5-File Mode + CSV Mode + Resume");
  console.log("═══════════════════════════════════════════════════════════\n");

  console.log("🛡️ Features:");
  console.log("   ✅ Auto-detects 5 files (urls.txt, first-name.txt, last-name.txt, subject.txt, message.txt)");
  console.log("   ✅ OR single CSV/Excel file (url,email,name,subject,message)");
  console.log("   ✅ Resume — already-done websites skip");
  console.log("   ✅ Tries home page → Contact page → Newsletter box");
  console.log("   ✅ Splits First/Last name if form needs it");
  console.log("   ✅ Handles Subject dropdown + textarea + text input");
  console.log("   ✅ Auto closes popups and accepts cookies\n");

  // MODE 1: 5-file mode
  let tasks = await loadFiveFileMode();

  // MODE 2: CSV/Excel mode
  if (!tasks || tasks.length === 0) {
    console.log("ℹ️  5-file mode not found, trying CSV/Excel mode...\n");
    let filePath = await findInputFile();
    if (filePath) {
      console.log(`📂 Auto-detected: ${path.basename(filePath)}\n`);
    } else {
      filePath = await ask("Enter path to input file (CSV/Excel/TXT): ");
      if (!filePath) { console.log("❌ Cancelled"); return; }
    }
    try {
      tasks = await loadSingleFileMode(filePath.trim());
    } catch (err) {
      console.log(`❌ Could not read file: ${err.message}`);
      return;
    }
  }

  if (!tasks || tasks.length === 0) {
    console.log("❌ No valid rows found!");
    console.log("   Put urls.txt (one URL per line) OR input.csv next to this script.");
    return;
  }

  // Resume
  const processed = await loadAlreadyProcessedKeys();
  const total = tasks.length;
  tasks = tasks.filter(t => !processed.has(`${t.url}::${t.email}`));
  const skipped = total - tasks.length;

  if (skipped > 0) console.log(`⏭️  Skipping ${skipped} already-processed website(s).\n`);

  if (tasks.length === 0) {
    console.log("✅ All websites already processed!");
    console.log(`   Check ${path.basename(CONFIG.resultsFile)}`);
    console.log("   Delete it to reprocess from scratch.");
    return;
  }

  console.log(`\n📋 ${tasks.length} website(s) to process:`);
  tasks.slice(0, 5).forEach((t, i) => console.log(`   ${i + 1}: ${t.url}  →  ${t.email}`));
  if (tasks.length > 5) console.log(`   ... and ${tasks.length - 5} more`);

  console.log(`\n⏰ Processing one website at a time (single Chrome tab)`);
  console.log("✅ Press Ctrl+C to stop — progress saves after every website\n");

  const confirm = await ask("Start? (yes/no): ");
  if (confirm.toLowerCase() !== 'yes') { console.log("❌ Cancelled"); return; }

  const startTime = Date.now();
  let browser = null;
  let results = [];

  try {
    await logMessage("🌐 Opening browser...");
    try {
      browser = await chromium.launch({ headless: false, slowMo: CONFIG.slowMo });
    } catch {
      browser = await chromium.launch({ headless: false, slowMo: CONFIG.slowMo, channel: "chrome" });
    }
    await logMessage("✅ Browser opened");

    results = await runPool(tasks, browser, CONFIG.maxTabs);
  } catch (error) {
    await logMessage(`❌ Browser error: ${error.message}`);
  } finally {
    if (browser) {
      await browser.close();
      await logMessage("🔒 Browser closed");
    }
  }

  const totalTime = ((Date.now() - startTime) / 1000).toFixed(2);

  console.log(`\n${"=".repeat(60)}`);
  console.log(`📊 FINAL SUMMARY`);
  console.log(`${"=".repeat(60)}`);
  const successCount = results.filter(r => r.status === "SUCCESS").length;
  const submittedCount = results.filter(r => r.status === "SUBMITTED_BUT_UNCONFIRMED").length;
  const errorCount = results.filter(r => r.status === "ERROR").length;
  const noSubmitCount = results.filter(r => r.status === "NO_SUBMIT").length;
  const newsletterCount = results.filter(r => r.status === "NEWSLETTER_SUBMITTED").length;

  console.log(`   ✅ Success (Confirmed):     ${successCount}`);
  console.log(`   ⚠️ Submitted (Unconfirmed): ${submittedCount}`);
  console.log(`   📧 Newsletter Submitted:    ${newsletterCount}`);
  console.log(`   ❌ Failed (Error):          ${errorCount}`);
  console.log(`   📝 No Form/Submit:          ${noSubmitCount}`);
  console.log(`   📊 Total:                   ${results.length}`);
  console.log(`   ⏱️ Total Time:              ${totalTime}s`);
  console.log(`${"=".repeat(60)}`);

  if (successCount > 0) {
    console.log(`\n✅ Successful:`);
    results.filter(r => r.status === "SUCCESS").forEach((r, i) => console.log(`   ${i + 1}. ${r.url}`));
  }
  if (errorCount > 0) {
    console.log(`\n❌ Failed:`);
    results.filter(r => r.status === "ERROR").forEach((r, i) => console.log(`   ${i + 1}. ${r.url} (${r.error})`));
  }

  console.log(`\n✅ Done! Check ${path.basename(CONFIG.resultsFile)} for details.`);
}

process.on('SIGINT', async () => {
  console.log('\n\n🛑 Stopped by user');
  process.exit(0);
});

main().catch(console.error);