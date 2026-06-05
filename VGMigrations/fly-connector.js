/**
 * fly-connector.js
 *
 * Playwright automation of the AvePoint Fly "Create Connection" flow.
 * Hardened: anchored on inputs (not panel wrappers), aggressive recovery
 * between tasks, DOM dump on failure for diagnostics.
 *
 * Version: 1.2.0
 * Last Updated: 2026-05-17
 *
 * Changelog:
 * v1.2.0 (2026-05-17) - Optimized timeouts for faster execution
 * v1.1.0 (2026-05-17) - Fixed dropdown selector timeouts
 * v1.0.0 (2026-05-16) - Initial release
 */

const { chromium } = require('playwright');
const fs   = require('fs');
const path = require('path');
const readline = require('readline');

const FLY_URL          = 'https://fly.avepointonlineservices.com/#/dashboard';
const CONNECTIONS_URL  = 'https://fly.avepointonlineservices.com/#/settings/connection';
const TENANT_MGMT_URL  = 'https://www.avepointonlineservices.com/#/management/tenant';
const APP_MGMT_URL     = 'https://www.avepointonlineservices.com/#/management/app';
const STORAGE_STATE    = path.join(__dirname, 'auth', 'storageState.json');
const LOG_DIR          = path.join(__dirname, 'logs');

fs.mkdirSync(path.dirname(STORAGE_STATE), { recursive: true });
fs.mkdirSync(LOG_DIR, { recursive: true });

const RUN_MODE   = (process.argv.find(a => a.startsWith('--mode=')) || '').split('=')[1] || 'unknown';
const DISPLAY_NM = (process.argv.find(a => a.startsWith('--display-name=')) || '').split('=')[1] || '';
const HEADLESS   = process.argv.includes('--headless') || process.env.FLY_HEADLESS === '1';

// Per-run folder under logs/ so artifacts don't pile up across 200+ tenants.
// Folder name: <SanitizedDisplayName>_<yyyy-MM-dd_HHmmss>
// For login mode (no tenant context): login_<timestamp>
function sanitiseForPath(s) {
  return (s || '').replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
}
function localStamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}
const RUN_FOLDER_NAME = (function () {
  const stamp = localStamp();
  if (RUN_MODE === 'login') return `login_${stamp}`;
  const name = sanitiseForPath(DISPLAY_NM) || RUN_MODE;
  return `${name}_${stamp}`;
})();
const RUN_DIR  = path.join(LOG_DIR, RUN_FOLDER_NAME);
const LOG_FILE = path.join(RUN_DIR, 'run.log');

fs.mkdirSync(RUN_DIR, { recursive: true });

let logStream = null;
try { logStream = fs.createWriteStream(LOG_FILE, { flags: 'a' }); }
catch (e) { process.stderr.write(`WARN: cannot open log: ${e.message}\n`); }

function ts() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ` +
         `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.${String(d.getMilliseconds()).padStart(3,'0')}`;
}
function logLine(level, msg) {
  const line = `${ts()} [${level.padEnd(5)}] ${msg}`;
  if (logStream) { try { logStream.write(line + '\n'); } catch {} }
}
const logInfo  = m => logLine('INFO',  m);
const logWarn  = m => logLine('WARN',  m);
const logError = m => logLine('ERROR', m);
const logStep  = m => logLine('STEP',  m);
const logTask  = m => logLine('TASK',  m);

function emit(obj) {
  process.stdout.write(JSON.stringify(obj) + '\n');
  if (obj.event) {
    const level = ({ 'fatal':'ERROR','error':'ERROR','warn':'WARN' }[obj.event]) || 'INFO';
    logLine(level, `[${obj.event}] ${obj.message || ''}`);
  } else if (obj.id && obj.status) {
    logTask(`${obj.id} ${obj.status.padEnd(8)} ${obj.message || ''}`);
  } else {
    logInfo(JSON.stringify(obj));
  }
}

process.on('uncaughtException',  err => {
  logError(`UNCAUGHT: ${err && err.stack || err}`);
  emit({ event: 'fatal', message: `Uncaught: ${err && err.message || err}` });
  setTimeout(() => process.exit(98), 200);
});
process.on('unhandledRejection', err => {
  logError(`UNHANDLED REJECTION: ${err && err.stack || err}`);
  emit({ event: 'fatal', message: `Unhandled rejection: ${err && err.message || err}` });
  setTimeout(() => process.exit(97), 200);
});

async function login() {
  logInfo('==== LOGIN MODE START ====');
  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext();
  const page    = await context.newPage();

  emit({ event: 'info', message: 'Opening Fly portal - please complete Microsoft SSO sign-in.' });
  await page.goto(FLY_URL);

  try {
    await page.waitForURL(/fly\.avepointonlineservices\.com\/#\//, { timeout: 5 * 60 * 1000 });
    await page.waitForLoadState('load', { timeout: 15_000 }).catch(() => {});
  } catch (err) {
    logError(`Login timeout: ${err.message}`);
    emit({ event: 'error', message: 'Timed out waiting for sign-in.' });
    await browser.close();
    process.exit(2);
  }

  // Also visit the AOS management portal so its session cookies are included
  // in the saved state (needed by the AOS Tenant & App Setup flow).
  emit({ event: 'info', message: 'Capturing AOS management session...' });
  logStep('visiting AOS management portal to capture session');
  try {
    await page.goto('https://www.avepointonlineservices.com/#/dashboard');
    await page.waitForLoadState('load', { timeout: 20_000 }).catch(() => {});
    await page.waitForTimeout(2000); // let any SSO redirect settle
  } catch (err) {
    logWarn(`AOS portal visit failed: ${err.message}`);
  }

  await context.storageState({ path: STORAGE_STATE });
  emit({ event: 'login-ok', message: `Session persisted to ${STORAGE_STATE}` });
  await browser.close();
  logInfo('==== LOGIN MODE END ====');
}

async function create() {
  logInfo('==== CREATE MODE START ====');
  logInfo(`Log file: ${LOG_FILE}`);

  if (!fs.existsSync(STORAGE_STATE)) {
    emit({ event: 'fatal', message: 'No saved session. Run --mode=login first.' });
    process.exit(3);
  }

  const tasks = await readTasks();
  logInfo(`Received ${tasks.length} tasks`);
  if (tasks.length === 0) {
    emit({ event: 'fatal', message: 'No tasks received on stdin.' });
    process.exit(4);
  }

  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext({ storageState: STORAGE_STATE });
  const page    = await context.newPage();

  page.on('pageerror',     err => logWarn(`page error: ${err.message}`));
  page.on('console',       msg => { if (msg.type() === 'error') logWarn(`console: ${msg.text()}`); });
  page.on('requestfailed', req => logWarn(`req failed: ${req.method()} ${req.url()}`));

  logStep(`Navigating to ${CONNECTIONS_URL}`);
  try {
    await navigateToConnectionsPage(page);
  } catch (err) {
    logError(`Failed to reach Connections page: ${err.message}`);
    emit({ event: 'fatal', message: `Could not reach Connections page. Session may be stale - re-run --mode=login. (${err.message})` });
    await browser.close();
    process.exit(5);
  }

  const existing = await readExistingConnectionNames(page);
  logInfo(`Existing connections cached: ${existing.size}`);
  emit({ event: 'info', message: `Found ${existing.size} existing connections.` });

  for (const task of tasks) {
    await ensureCleanListPage(page, task.id);

    logTask(`${task.id} STARTED  workload="${task.workloadLabel}" name="${task.connectionName}"`);
    emit({ id: task.id, status: 'WORKING', message: '' });

    if (existing.has(task.connectionName)) {
      emit({ id: task.id, status: 'SKIPPED', message: 'Already exists' });
      continue;
    }

    try {
      await createOneConnection(page, task);
      existing.add(task.connectionName);
      emit({ id: task.id, status: 'CREATED', message: 'OK' });
    } catch (err) {
      await dumpFailure(page, task.id, err);
      const msg = err && err.message || String(err);
      emit({
        id: task.id,
        status: 'FAILED',
        message: `${msg.split('\n')[0]}  (see ${RUN_DIR})`
      });
    }
  }

  await browser.close();
  logInfo('==== CREATE MODE END ====');
  emit({ event: 'done', message: `Log: ${LOG_FILE}` });
}

// ---------------------------------------------------------------------------
// Navigate to the Connections page by clicking through the menu.
// Direct hash navigation to #/settings/connection does not work reliably -
// the AOS SPA needs the dashboard mounted first and the Settings submenu
// expanded via click before its router responds to the connection route.
// ---------------------------------------------------------------------------
async function navigateToConnectionsPage(page) {
  logStep('navigateToConnectionsPage: start');

  // 1. Start at the dashboard so the SPA fully mounts
  const onConnections = page.url().includes('/#/settings/connection');
  if (!onConnections) {
    await page.goto(FLY_URL);
    await page.waitForLoadState('load', { timeout: 15_000 }).catch(() => {});
  }

  // 2. Wait for the left nav to render. Dashboard button is the most reliable anchor.
  logStep('waiting for left nav to render');
  await page.locator('button[aria-label="Dashboard"]').first()
            .waitFor({ state: 'visible', timeout: 15_000 });

  // 3. If the Create connection button is already visible, we're done
  const createBtnFast = page.getByRole('button', { name: /^create connection$/i }).first();
  if (await createBtnFast.isVisible({ timeout: 1500 }).catch(() => false)) {
    logStep('navigateToConnectionsPage: already on Connections page');
    return;
  }

  // 4. Click the Settings menu item to expand its sub-items.
  //    Fluent UI v9 nav uses <button aria-label="Settings">.
  logStep('clicking Settings menu (button[aria-label="Settings"])');
  const settingsBtn = page.locator('button[aria-label="Settings"]').first();
  await settingsBtn.waitFor({ state: 'visible', timeout: 10_000 });
  await settingsBtn.click();
  await page.waitForTimeout(300); // submenu animation

  // 5. Click the Connection sub-item. After expanding Settings the sub-item
  //    typically renders as <button aria-label="Connection"> or as an <a>
  //    with href="#/settings/connection". Try both.
  logStep('clicking Connection sub-item');
  const connectionTargets = [
    'button[aria-label="Connection"]',
    'button[aria-label="Connections"]',
    'a[href$="#/settings/connection"]',
    'a[href="#/settings/connection"]'
  ];
  let clicked = false;
  for (const sel of connectionTargets) {
    const loc = page.locator(sel).first();
    if (await loc.isVisible({ timeout: 1500 }).catch(() => false)) {
      logStep(`Connection sub-item found via "${sel}"`);
      await loc.click();
      clicked = true;
      break;
    }
  }
  if (!clicked) {
    // Last resort: any element whose accessible name is exactly "Connection"
    logStep('Connection sub-item: fallback to role-based match');
    await page.getByRole('button', { name: /^Connection$/, exact: true }).first()
              .click({ timeout: 10_000 });
  }

  // 6. Wait for the Create connection button as the success signal
  logStep('waiting for Create connection button on list page');
  await page.getByRole('button', { name: /^create connection$/i }).first()
            .waitFor({ state: 'visible', timeout: 15_000 });
  await page.waitForLoadState('load', { timeout: 8_000 }).catch(() => {});
  logStep('navigateToConnectionsPage: done');
}

// ---------------------------------------------------------------------------
async function ensureCleanListPage(page, taskId) {
  logStep(`${taskId} ensureCleanListPage: start`);

  for (let i = 0; i < 3; i++) {
    await page.keyboard.press('Escape').catch(() => {});
    await page.waitForTimeout(150);
  }

  const closeBtn = page.locator(
    '.ant-drawer-close, .ms-Panel-closeButton, button[aria-label="Close"], button[title="Close"]'
  ).first();
  if (await closeBtn.isVisible({ timeout: 500 }).catch(() => false)) {
    logStep(`${taskId} clicking drawer close button`);
    await closeBtn.click().catch(() => {});
    await page.waitForTimeout(300);
  }

  // Verify the Create button is visible AND stable. Visibility alone isn't
  // enough - the button may flicker as a previous drawer's close transition
  // completes. Require it to remain visible across a short delay.
  const createBtn = page.getByRole('button', { name: /^create connection$/i }).first();
  const firstCheck = await createBtn.isVisible({ timeout: 2_000 }).catch(() => false);
  if (firstCheck) {
    await page.waitForTimeout(400);
    const secondCheck = await createBtn.isVisible({ timeout: 1_000 }).catch(() => false);
    if (secondCheck) {
      logStep(`${taskId} ensureCleanListPage: done`);
      return;
    }
  }

  logStep(`${taskId} Create button not stable; renavigating via menu`);
  try {
    await navigateToConnectionsPage(page);
  } catch (err) {
    logWarn(`${taskId} renavigation failed: ${err.message}`);
  }
  logStep(`${taskId} ensureCleanListPage: done`);
}

// ---------------------------------------------------------------------------
// Dropdown helpers - work against Fluent UI v9 (the framework AOS uses).
//
// Field rows in this form have a label (e.g. "App profile") followed by a
// clickable dropdown control. The control is typically a button or div with
// role="combobox"/role="button" and may include trailing icon buttons (refresh).
// We pick the first "interactive" element after the label.
// ---------------------------------------------------------------------------
async function openDropdownAfterLabel(page, labelText, opts) {
  opts = opts || {};
  // Build the label predicate. By default uses starts-with for prefix-tolerance
  // (handles trailing asterisk in required-field labels). Pass exact:true when
  // the label is a prefix of ANOTHER label on the same form (e.g. "Service
  // account" vs "Service account authentication" - exact:true on the former).
  let labelPred;
  if (opts.exact) {
    // Match exact text, plus tolerant of trailing asterisk-with-space variants
    const variants = [
      labelText,
      labelText + '*',
      labelText + ' *'
    ].map(s => `normalize-space(.)="${s}"`);
    labelPred = '(' + variants.join(' or ') + ')';
  } else {
    labelPred = `starts-with(normalize-space(.), "${labelText}")`;
  }
  const xp = `//label[${labelPred}]/following::*[@role="combobox"][1]`;
  const dd = page.locator(`xpath=${xp}`).first();
  await dd.waitFor({ state: 'visible', timeout: 10_000 });
  await dd.click();
  await page.waitForTimeout(300);
}

// After opening a dropdown that contains a search box, type into the first
// visible search input. AOS uses several placeholder texts ("Search tenant",
// "Search app profile", and the generic "Search" for service accounts).
// If no search input is visible (the dropdown has no filter), this is a no-op.
// Click the "Destination" radio button if the workload form includes the
// "Configure source or destination connection" radio group. Currently this
// applies to Teams Chat (and possibly other workloads in the future).
// No-op when the radio isn't present.
async function pickDestinationIfPresent(page, taskId) {
  // Look for the heading text first - if it isn't on the form, there's
  // nothing to do. Bail out fast (short timeout) so we don't waste time
  // on every workload that doesn't have this option.
  const heading = page.locator('text=Configure source or destination connection').first();
  if (!(await heading.isVisible({ timeout: 1500 }).catch(() => false))) {
    return;
  }
  logStep(`${taskId} Source/Destination radio present - picking Destination`);

  // Fluent UI v9 radios are <input type="radio">. The visible label sits next
  // to it. Try the most stable selectors in turn.
  const strategies = [
    page.getByRole('radio', { name: /^\s*Destination\s*$/i }).first(),
    page.locator('label:has-text("Destination") input[type="radio"]').first(),
    page.locator('input[type="radio"][value*="estination" i]').first()
  ];
  for (const loc of strategies) {
    if (await loc.isVisible({ timeout: 1000 }).catch(() => false)) {
      // .check() is the right Playwright primitive for radios - it tolerates
      // being already-checked and clicks the label rather than the input.
      try { await loc.check({ timeout: 5_000 }); return; }
      catch { /* try next strategy */ }
    }
  }
  // Last resort - click any element whose accessible name is "Destination"
  await page.getByText('Destination', { exact: true }).first().click({ timeout: 5_000 });
}

async function fillFirstVisibleSearch(page, text) {
  // Named-placeholder search inputs ONLY. Do NOT fall back to a generic
  // input[placeholder="Search"] because AOS has a page-wide search at the
  // top of every page with that placeholder, and matching it would (a) put
  // text in the wrong place and (b) steal focus and close the open dropdown.
  // If no named search box matches, return false and let the option-matcher
  // pick from the unfiltered list.
  const candidates = [
    'input[placeholder="Search tenant"]',
    'input[placeholder="Search app profile"]',
    'input[placeholder="Search delegated app profile"]',
    'input[placeholder="Search delegated app profiles"]',
    'input[placeholder="Search service account"]',
    'input[placeholder="Search service accounts"]'
  ];
  for (const sel of candidates) {
    const loc = page.locator(sel).first();
    if (await loc.isVisible({ timeout: 500 }).catch(() => false)) {
      await loc.fill('');
      await loc.fill(text);
      await page.waitForTimeout(400);
      return true;
    }
  }
  return false;
}

async function pickDropdownOptionContaining(page, needle) {  // Match option by title attribute (exact), then by visible text (substring).
  // Fluent UI v9 uses .fui-Option with role="option" and a title attribute.
  const escaped = needle.replace(/"/g, '\\"');
  const needleRe = new RegExp(escapeRegex(needle), 'i');

  const strategies = [
    // 1. Exact title attribute match
    page.locator(`[role="option"][title*="${escaped}"]`).first(),
    page.locator(`.fui-Option[title*="${escaped}"]`).first(),
    // 2. Substring text match across various option containers
    page.locator('[role="option"]').filter({ hasText: needleRe }).first(),
    page.locator('.fui-Option').filter({ hasText: needleRe }).first(),
    page.locator('.ms-Dropdown-item').filter({ hasText: needleRe }).first(),
    page.locator('.ant-select-item-option').filter({ hasText: needleRe }).first(),
    page.locator('[class*="OptionItem"]').filter({ hasText: needleRe }).first(),
    // 3. Plain <li> elements (used by some AOS dropdowns)
    page.locator('ul li, [role="listbox"] li, [class*="dropdown"] li')
        .filter({ hasText: needleRe }).first(),
    page.locator('li').filter({ hasText: needleRe }).first(),
  ];

  for (const loc of strategies) {
    if (await loc.isVisible({ timeout: 1500 }).catch(() => false)) {
      await loc.click();
      return;
    }
  }
  // Last-ditch: wait for any of the above with a longer timeout
  await page.locator('[role="option"], li').filter({ hasText: needleRe }).first()
            .waitFor({ state: 'visible', timeout: 5_000 });
  await page.locator('[role="option"], li').filter({ hasText: needleRe }).first().click();
}

async function createOneConnection(page, task) {
  logStep(`${task.id} click Create connection button`);
  await page.getByRole('button', { name: /create connection/i }).first().click({ timeout: 30_000 });

  logStep(`${task.id} wait for Connection name input`);
  const nameField = page.getByLabel('Connection name', { exact: false }).first();
  await nameField.waitFor({ state: 'visible', timeout: 15_000 });

  logStep(`${task.id} fill name="${task.connectionName}"`);
  await nameField.click();
  await nameField.fill('');
  await nameField.fill(task.connectionName);

  logStep(`${task.id} open Connection type dropdown`);
  // Library-agnostic: find clickable dropdown control following the label.
  // Matches Ant Design (.ant-select), Fluent UI (.ms-Dropdown), or generic role/combobox.
  const typeSelect = page.locator(
    'xpath=//*[normalize-space(text())="Connection type"]/ancestor-or-self::*[1]/following::*[' +
      'contains(@class,"ant-select") or contains(@class,"ms-Dropdown") or ' +
      '@role="combobox" or @role="listbox" or @role="button"' +
    '][1]'
  ).first();
  await typeSelect.waitFor({ state: 'visible', timeout: 10_000 });
  await typeSelect.click();
  await page.waitForTimeout(200);

  logStep(`${task.id} pick workload="${task.workloadLabel}"`);
  // Fluent UI v9 options have role="option" AND a title attribute matching the
  // visible label. Title attribute match is bulletproof vs whitespace/checkmark
  // icons that defeat hasText regexes. Multiple selector strategies for safety.
  const escapedLabel = task.workloadLabel.replace(/"/g, '\\"');
  const typeOption = page.locator([
    `[role="option"][title="${escapedLabel}"]`,
    `.fui-Option[title="${escapedLabel}"]`,
    `[role="option"]:has-text("${escapedLabel}")`,
    `.fui-Option:has-text("${escapedLabel}")`
  ].join(', ')).first();
  await typeOption.waitFor({ state: 'visible', timeout: 10_000 });
  await typeOption.click();
  await page.waitForTimeout(200);

  // --- Source/Destination radio (Teams Chat and possibly others) ---
  // For workloads that show "Configure source or destination connection",
  // we want Destination. The radio only appears for some workloads; this is
  // a no-op when absent.
  await pickDestinationIfPresent(page, task.id);

  logStep(`${task.id} open Tenant dropdown`);
  // Primary: Fluent v9 combobox uses aria-label="Select tenant" while empty.
  // Fallback: label-anchored helper if AOS ever changes the placeholder.
  let tenantSelect = page.locator('button[role="combobox"][aria-label="Select tenant"]').first();
  if (!(await tenantSelect.isVisible({ timeout: 2_000 }).catch(() => false))) {
    logStep(`${task.id} Tenant aria-label not found, falling back to label helper`);
    await openDropdownAfterLabel(page, 'Tenant');
  } else {
    await tenantSelect.click();
    await page.waitForTimeout(300);
  }

  logStep(`${task.id} type tenant search="${task.tenantSearch}"`);
  const searchBox = page.getByPlaceholder(/search tenant/i).first();
  await searchBox.waitFor({ state: 'visible', timeout: 10_000 });
  await searchBox.fill(task.tenantSearch);
  await page.waitForTimeout(400);

  logStep(`${task.id} pick tenant row`);
  const tenantOption = page.locator(
    `.ant-select-item-option, .ms-Dropdown-item, [role="option"]`
  ).filter({ hasText: new RegExp(escapeRegex(task.tenantSearch), 'i') }).first();
  await tenantOption.waitFor({ state: 'visible', timeout: 10_000 });
  await tenantOption.click();

  // --- Wait for the Credentials section to render (App profile dropdown appears) ---
  logStep(`${task.id} wait for App profile field to appear`);
  const appProfileLabel = page.locator(
    'xpath=//*[normalize-space(text())="App profile"]'
  ).first();
  await appProfileLabel.waitFor({ state: 'visible', timeout: 15_000 });
  await page.waitForTimeout(600); // let the section finish hydrating

  // --- App profile: open dropdown, type the credentials name, pick the match ---
  logStep(`${task.id} open App profile dropdown`);
  await openDropdownAfterLabel(page, 'App profile');
  logStep(`${task.id} search App profile for "${task.credentialsName}"`);
  // Search input in the opened listbox - placeholder may be specific or just "Search"
  await fillFirstVisibleSearch(page, task.credentialsName);
  await pickDropdownOptionContaining(page, task.credentialsName);

  // --- Service account authentication: pick Modern authentication ---
  // This field is absent for Teams Chat Destination (and any other flow that
  // uses delegated app auth only). Skip gracefully when not present.
  const authLabel = page.locator(
    'xpath=//label[starts-with(normalize-space(.), "Service account authentication")]'
  ).first();
  if (await authLabel.isVisible({ timeout: 2_000 }).catch(() => false)) {
    logStep(`${task.id} open Service account authentication dropdown`);
    await openDropdownAfterLabel(page, 'Service account authentication');
    logStep(`${task.id} pick "Modern authentication"`);
    await pickDropdownOptionContaining(page, 'Modern authentication');
    await page.waitForTimeout(600); // wait for auth menu to close + downstream list to refresh
  } else {
    logStep(`${task.id} Service account authentication field not present - skipping`);
  }

  // --- Microsoft delegated app profile: open, type credentials name, pick the match ---
  // Selecting Modern authentication (above) or being on a Destination form (e.g.
  // Teams Chat) both surface this required field. Same credentials substring
  // matches in either case.
  logStep(`${task.id} open Microsoft delegated app profile dropdown`);
  await openDropdownAfterLabel(page, 'Microsoft delegated app profile');
  logStep(`${task.id} search delegated app profile for "${task.credentialsName}"`);
  await fillFirstVisibleSearch(page, task.credentialsName);
  await pickDropdownOptionContaining(page, task.credentialsName);

  // --- Placeholder account (Teams Chat Destination): derive email and fill ---
  // Format: <credentialsName>@<tenantSearch>.onmicrosoft.com
  // The field is a plain text input, not a dropdown. Only present on certain
  // destination flows; no-op otherwise.
  const placeholderLabel = page.locator(
    'xpath=//label[starts-with(normalize-space(.), "Placeholder account")]'
  ).first();
  if (await placeholderLabel.isVisible({ timeout: 1_500 }).catch(() => false)) {
    let credName = task.credentialsName || '';
    const placeholderEmail = credName.includes('@')
      ? credName
      : `${credName}@${task.tenantSearch}.onmicrosoft.com`;
    logStep(`${task.id} fill Placeholder account="${placeholderEmail}"`);
    // The textbox is the next focusable input after the label.
    const placeholderInput = page.locator(
      'xpath=//label[starts-with(normalize-space(.), "Placeholder account")]' +
      '/following::input[@type="text" or @type="email" or not(@type)][1]'
    ).first();
    await placeholderInput.waitFor({ state: 'visible', timeout: 5_000 });
    await placeholderInput.fill('');
    await placeholderInput.fill(placeholderEmail);
    await page.waitForTimeout(200);
  }

  // SharePoint admin center URL is auto-filled by AOS when the tenant is picked.
  // We leave it alone unless a future requirement says otherwise.

  logStep(`${task.id} click Save`);
  const saveBtn = page.getByRole('button', { name: /^\s*save\s*$/i }).first();
  await saveBtn.click({ timeout: 10_000 });

  // Robust "form closed" detection:
  //   1. Drawer/dialog disappears (the form container, not just one input)
  //   2. Connection-list "Create connection" button is visible and stable
  //   3. Network idle (the POST to create the connection has completed)
  // Trusting "input is hidden" alone was racing transient DOM detaches mid-
  // render and proceeding before AOS had actually saved.
  logStep(`${task.id} wait for drawer to close`);
  await page.locator('.fui-Drawer, [role="dialog"]').first()
            .waitFor({ state: 'hidden', timeout: 60_000 })
            .catch(() => { /* drawer may already be gone */ });

  logStep(`${task.id} wait for list page to settle`);
  const listCreateBtn = page.getByRole('button', { name: /^create connection$/i }).first();
  await listCreateBtn.waitFor({ state: 'visible', timeout: 30_000 });
  // Stability check: button must remain visible for 800ms continuously, no
  // mid-transition state.
  await page.waitForTimeout(800);
  if (!(await listCreateBtn.isVisible({ timeout: 1_000 }).catch(() => false))) {
    throw new Error('List page did not settle after Save - Create connection button disappeared');
  }
  await page.waitForLoadState('networkidle', { timeout: 15_000 }).catch(() => {});
  logStep(`${task.id} create flow complete`);
}

async function dumpFailure(page, taskId, err) {
  const stamp    = Date.now();
  const pngPath  = path.join(RUN_DIR, `fail-${taskId}-${stamp}.png`);
  const htmlPath = path.join(RUN_DIR, `fail-${taskId}-${stamp}.html`);

  await page.screenshot({ path: pngPath, fullPage: true }).catch(e => logWarn(`screenshot failed: ${e.message}`));
  try {
    const html = await page.content();
    fs.writeFileSync(htmlPath, html, 'utf8');
    logError(`${taskId} DOM dumped to ${htmlPath}`);
  } catch (e) {
    logWarn(`html dump failed: ${e.message}`);
  }

  const msg = err && err.message || String(err);
  logError(`${taskId} FAILED   ${msg}`);
  logError(`${taskId} stack: ${err && err.stack || '(no stack)'}`);
  logError(`${taskId} screenshot: ${pngPath}`);
}

async function readExistingConnectionNames(page) {
  const names = new Set();
  try {
    await page.waitForSelector('text=Connection name', { timeout: 15_000 });
    let prev = -1;
    for (let i = 0; i < 50; i++) {
      const cells = await page.locator('[role="row"] [role="cell"]:first-child, td:first-child').allInnerTexts();
      for (const c of cells) {
        const t = c.trim();
        if (t && t !== 'Connection name') names.add(t);
      }
      if (names.size === prev) break;
      prev = names.size;
      await page.keyboard.press('PageDown').catch(() => {});
      await page.waitForTimeout(300);
    }
  } catch (err) {
    logWarn(`could not scrape existing names: ${err.message}`);
  }
  return names;
}

function readTasks() {
  return new Promise((resolve, reject) => {
    const rl = readline.createInterface({ input: process.stdin });
    const tasks = [];
    rl.on('line', line => {
      line = line.trim();
      if (!line) return;
      try { tasks.push(JSON.parse(line)); } catch (e) { logWarn(`bad task: ${line}`); }
    });
    rl.on('close', () => resolve(tasks));
    rl.on('error', reject);
  });
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// ---------------------------------------------------------------------------
// SETUP MODE  –  Tenant Management + App Management + Admin Consent
// ---------------------------------------------------------------------------
// Each stdin task:
//   { id, tenantDisplayName, tenantSearch, appProfileName, clientId, clientSecret }
// ---------------------------------------------------------------------------
async function setup() {
  logInfo('==== SETUP MODE START ====');
  if (!fs.existsSync(STORAGE_STATE)) {
    emit({ event: 'fatal', message: 'No saved session. Run Sign In to AOS first.' });
    process.exit(3);
  }

  const tasks = await readTasks();
  if (tasks.length === 0) {
    emit({ event: 'fatal', message: 'No tasks received on stdin.' });
    process.exit(4);
  }

  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext({ storageState: STORAGE_STATE });
  const page    = await context.newPage();

  page.on('pageerror',     err => logWarn(`page error: ${err.message}`));
  page.on('console',       msg => { if (msg.type() === 'error') logWarn(`console: ${msg.text()}`); });
  page.on('requestfailed', req => logWarn(`req failed: ${req.method()} ${req.url()}`));

  // Verify session is alive
  logStep('navigating to dashboard to verify session');
  await page.goto(FLY_URL);
  await page.waitForLoadState('load', { timeout: 15_000 }).catch(() => {});
  try {
    await page.locator('button[aria-label="Dashboard"]').first()
              .waitFor({ state: 'visible', timeout: 15_000 });
  } catch {
    emit({ event: 'fatal', message: 'Session appears stale — re-run Sign In to AOS.' });
    await browser.close();
    process.exit(5);
  }

  for (const task of tasks) {
    logTask(`${task.id} STARTED tenant="${task.tenantDisplayName}"`);

    // Step 1: Tenant Management
    emit({ id: task.id, status: 'WORKING', message: 'Adding tenant to Tenant Management...' });
    try {
      const tenantResult = await addOneTenant(page, task);
      logTask(`${task.id} tenant: ${tenantResult}`);
      emit({ id: task.id, status: 'WORKING', message: `Tenant: ${tenantResult}. Adding app profile...` });
    } catch (err) {
      await dumpFailure(page, `${task.id}-tenant`, err);
      emit({ id: task.id, status: 'FAILED',
             message: `Tenant step failed: ${err.message.split('\n')[0]}  (see ${RUN_DIR})` });
      continue;
    }

    // Step 2: App Management
    try {
      const appResult = await addOneAppProfile(page, task);
      logTask(`${task.id} app: ${appResult}`);
      emit({ id: task.id, status: 'DONE',
             message: `Tenant registered. App profile: ${appResult}` });
    } catch (err) {
      await dumpFailure(page, `${task.id}-app`, err);
      emit({ id: task.id, status: 'FAILED',
             message: `App profile step failed: ${err.message.split('\n')[0]}  (see ${RUN_DIR})` });
    }
  }

  await browser.close();
  logInfo('==== SETUP MODE END ====');
  emit({ event: 'done', message: `Log: ${LOG_FILE}` });
}

// Navigate directly to the AOS management portal page for 'tenant' or 'app'.
// TENANT_MGMT_URL / APP_MGMT_URL live on www.avepointonlineservices.com — a
// different subdomain from fly.avepointonlineservices.com. Going via the Fly
// Settings nav does not reach them, so we use the direct URL instead.
async function navigateToAosManagement(page, section) {
  const url = section === 'tenant' ? TENANT_MGMT_URL : APP_MGMT_URL;
  logStep(`navigateToAosManagement: ${url}`);

  await page.goto(url);
  await page.waitForLoadState('load', { timeout: 20_000 }).catch(() => {});
  await page.waitForTimeout(500); // let SPA router finish

  // Confirm we are on the right page, not a login redirect
  const finalUrl = page.url();
  if (!finalUrl.includes('avepointonlineservices.com') ||
      finalUrl.toLowerCase().includes('login') ||
      finalUrl.toLowerCase().includes('signin')) {
    throw new Error(
      `Navigation landed on unexpected URL: ${finalUrl}\n` +
      'Session may be stale — re-run "Sign in to AOS" to refresh it.'
    );
  }

  // Use the breadcrumb heading as the readiness signal — it is stable and
  // doesn't depend on button label wording (which differs between sections).
  const headingRe = section === 'tenant' ? /tenant management/i : /app management/i;
  await page.getByText(headingRe).first()
            .waitFor({ state: 'visible', timeout: 20_000 });

  await page.waitForTimeout(300);
  logStep(`navigateToAosManagement: done (${section})`);
}

// ---------------------------------------------------------------------------
// Tenant Management — add one tenant
// Returns: 'added' | 'already exists'
// ---------------------------------------------------------------------------
async function addOneTenant(page, task) {
  logStep(`${task.id} addOneTenant: "${task.tenantSearch}"`);

  await navigateToAosManagement(page, 'tenant');

  // Wait for the table data to load (heading visible ≠ rows rendered)
  await page.waitForTimeout(1500);

  // Substring match scoped to data cells only (avoids false positives from nav
  // elements — the cell scope is what prevents those, not strict matching).
  const tenantRe = new RegExp(escapeRegex(task.tenantSearch), 'i');
  const nameCell = page.locator('td, [role="gridcell"], [role="cell"]')
                       .filter({ hasText: tenantRe }).first();
  if (await nameCell.isVisible({ timeout: 1_000 }).catch(() => false)) {
    logStep(`${task.id} tenant "${task.tenantSearch}" already exists`);
    return 'already exists';
  }

  // Click the "Connect tenant" / "Add" button (label varies by AOS version)
  logStep(`${task.id} clicking Connect tenant`);
  const addBtns = [
    page.getByRole('button', { name: /connect tenant/i }).first(),
    page.locator('button').filter({ hasText: /connect tenant/i }).first(),
    page.getByRole('button', { name: /^add tenant$/i }).first(),
    page.getByRole('button', { name: /^add$/i }).first(),
    page.locator('button').filter({ hasText: /^add$/i }).first(),
  ];
  let clicked = false;
  for (const b of addBtns) {
    if (await b.isVisible({ timeout: 1500 }).catch(() => false)) { await b.click(); clicked = true; break; }
  }
  if (!clicked) throw new Error('Could not find Connect tenant / Add button on Tenant Management page');
  await page.waitForTimeout(500);

  // The "Connect tenant" dialog first shows a platform picker (Microsoft /
  // Google / Salesforce / Amazon). The dialog may be present-but-hidden during
  // its CSS open animation, so wait for the Microsoft card text to be visible
  // rather than waiting on the dialog container itself.
  logStep(`${task.id} selecting Microsoft platform`);

  // Wait up to 10s for any of the Microsoft card strategies to become visible
  const msStrategies = [
    page.locator('li, [role="listitem"], [role="option"]').filter({ hasText: /^microsoft$/i }).first(),
    page.locator('li, [role="listitem"]').filter({ hasText: /microsoft/i }).first(),
    page.getByRole('radio', { name: /^microsoft$/i }).first(),
    page.locator('label').filter({ hasText: /^microsoft$/i }).first(),
    page.locator('[class*="option"], [class*="card"], [class*="item"], [class*="platform"], [class*="provider"]')
        .filter({ hasText: /microsoft/i }).first(),
    page.getByText('Microsoft', { exact: true }).first(),
  ];
  let msPicked = false;
  const msDeadline = Date.now() + 10_000;
  while (!msPicked && Date.now() < msDeadline) {
    for (const loc of msStrategies) {
      if (await loc.isVisible({ timeout: 500 }).catch(() => false)) {
        await loc.click(); msPicked = true; break;
      }
    }
    if (!msPicked) await page.waitForTimeout(300);
  }
  if (!msPicked) throw new Error('Microsoft platform option not found — dialog may not have opened');
  await page.waitForTimeout(300);

  logStep(`${task.id} clicking Connect (platform selection)`);
  const platformDialog = page.locator('[role="dialog"], .ant-modal-content, .fui-Drawer').first();
  const platformConnectBtn = [
    platformDialog.getByRole('button', { name: /^connect$/i }).first(),
    page.getByRole('button', { name: /^connect$/i }).first(),
  ];
  let platformConnectClicked = false;
  for (const b of platformConnectBtn) {
    if (await b.isVisible({ timeout: 2_000 }).catch(() => false)) {
      // Listen for a Playwright-managed popup (window.open OAuth flow)
      const popupPromise = page.waitForEvent('popup', { timeout: 5_000 }).catch(() => null);
      await b.click();
      platformConnectClicked = true;

      const popup = await popupPromise;
      if (popup) {
        // Playwright caught the OAuth window — wait for user to finish logging in
        logStep(`${task.id} Microsoft login popup detected — waiting for authentication (up to 5 min)...`);
        await popup.waitForEvent('close', { timeout: 300_000 });
        logStep(`${task.id} login popup closed — continuing`);
        await page.waitForTimeout(2_000);
      }
      // If popup is null the OAuth may be opening in a separate OS window that
      // Playwright cannot track. Fall through and wait for the search box below.
      break;
    }
  }
  if (!platformConnectClicked) throw new Error('Connect button not found in platform selection dialog');

  // After OAuth completes, AvePoint either:
  //   (a) auto-connects the tenant and closes the drawer immediately, OR
  //   (b) shows a tenant-search/confirmation form in the drawer
  // Check for a search box quickly; if none appears, assume auto-connect (a).
  logStep(`${task.id} checking for search form after OAuth...`);
  await page.waitForTimeout(1500);

  const drawerScope = page.locator('[role="dialog"], .ant-modal-content, .fui-Drawer').first();
  const searchCandidates = [
    drawerScope.getByPlaceholder(/search tenant/i).first(),
    drawerScope.getByPlaceholder(/search/i).first(),
    drawerScope.locator('input[type="text"], input[type="search"]').first(),
  ];
  let searchBox = null;
  for (const s of searchCandidates) {
    if (await s.isVisible({ timeout: 4_000 }).catch(() => false)) { searchBox = s; break; }
  }

  if (searchBox) {
    // Flow (b): manual search/confirm step
    logStep(`${task.id} search form present — filling tenant name`);
    await searchBox.fill('');
    await searchBox.fill(task.tenantSearch);
    await page.waitForTimeout(600);
    logStep(`${task.id} picking tenant option`);
    await pickDropdownOptionContaining(page, task.tenantSearch);
    await page.waitForTimeout(300);
    logStep(`${task.id} saving tenant`);
    const saveBtn = page.getByRole('button', { name: /^\s*(ok|save|confirm|add)\s*$/i }).first();
    await saveBtn.click({ timeout: 10_000 });
    await page.locator('[role="dialog"], .ant-modal-content, .fui-Drawer').first()
              .waitFor({ state: 'hidden', timeout: 30_000 }).catch(() => {});
    await page.waitForTimeout(400);
  } else {
    // Flow (a): drawer already closed — OAuth handled the connection automatically
    logStep(`${task.id} drawer closed — tenant connected automatically via OAuth`);
    await page.waitForTimeout(1000);
  }

  // Verify the tenant now appears in the list (substring match — AOS may use
  // the full domain name rather than the short tenantSearch value)
  logStep(`${task.id} verifying tenant in list...`);
  await navigateToAosManagement(page, 'tenant');
  await page.waitForTimeout(1500);
  const verifyRe = new RegExp(escapeRegex(task.tenantSearch), 'i');
  const verified = page.locator('td, [role="gridcell"], [role="cell"]')
                       .filter({ hasText: verifyRe }).first();
  if (!await verified.isVisible({ timeout: 5_000 }).catch(() => false)) {
    throw new Error(
      `Tenant "${task.tenantSearch}" not found in list after connection. ` +
      'Check that the Microsoft account used belongs to that tenant.'
    );
  }

  logStep(`${task.id} addOneTenant: done`);
  return 'added';
}

// ---------------------------------------------------------------------------
// App Management — add one app profile via the 3-step wizard then grant consent
// Wizard steps observed in AOS:
//   1. Select services  (tenant dropdown + service radio e.g. "Fly")
//   2. Choose setup method  (manual credentials entry)
//   3. Consent to apps  (grant admin consent button)
// Returns: 'created + consent granted' | 'created (consent skipped)' | 'already exists'
// ---------------------------------------------------------------------------
async function addOneAppProfile(page, task) {
  logStep(`${task.id} addOneAppProfile: "${task.appProfileName}"`);

  await navigateToAosManagement(page, 'app');

  // Wait for the table data to load before checking
  await page.waitForTimeout(1500);

  // Check whether it already exists — cell-scoped substring match
  for (const needle of [task.appProfileName, task.tenantDisplayName, task.tenantSearch]) {
    if (!needle) continue;
    const existing = page.locator('td, [role="gridcell"], [role="cell"]')
                         .filter({ hasText: new RegExp(escapeRegex(needle), 'i') }).first();
    if (await existing.isVisible({ timeout: 1_000 }).catch(() => false)) {
      logStep(`${task.id} app profile already exists (matched "${needle}")`);
      const row = existing.locator('xpath=ancestor::tr | ancestor::*[@role="row"]').first();
      const granted = await grantAdminConsent(page, task, row);
      return granted ? 'already exists (consent granted)' : 'already exists';
    }
  }

  // Click Add — opens the multi-step wizard as a full page
  logStep(`${task.id} clicking Add (app profile)`);
  const addBtns = [
    page.getByRole('button', { name: /^add app(lication)?( profile)?$/i }).first(),
    page.locator('button').filter({ hasText: /^add app(lication)?( profile)?$/i }).first(),
    page.getByRole('button', { name: /^add$/i }).first(),
    page.locator('button').filter({ hasText: /^add$/i }).first(),
    page.getByRole('button', { name: /^create/i }).first(),
  ];
  let appAddClicked = false;
  for (const b of addBtns) {
    if (await b.isVisible({ timeout: 1500 }).catch(() => false)) { await b.click(); appAddClicked = true; break; }
  }
  if (!appAddClicked) throw new Error('Could not find Add button on App Management page');
  await page.waitForTimeout(600);

  // ── STEP 1: Select services ──────────────────────────────────────────────
  logStep(`${task.id} wizard step 1: Select services`);
  await page.getByText(/select services/i).first()
            .waitFor({ state: 'visible', timeout: 10_000 });

  // Select tenant from the Tenant dropdown
  logStep(`${task.id} step 1: selecting tenant "${task.tenantSearch}"`);
  await openDropdownAfterLabel(page, 'Tenant');
  const tenantSearchFilled = await fillFirstVisibleSearch(page, task.tenantSearch);
  if (!tenantSearchFilled) await page.waitForTimeout(200);
  try {
    await pickDropdownOptionContaining(page, task.tenantSearch);
  } catch (err) {
    // Tenant not in the dropdown — the wizard only shows tenants without an app
    // profile yet, so if it's missing the profile already exists.
    logStep(`${task.id} tenant not found in wizard dropdown — app profile likely already set up`);
    await page.keyboard.press('Escape').catch(() => {});
    await page.waitForTimeout(300);
    await page.getByRole('button', { name: /^cancel$/i }).first().click({ timeout: 3_000 }).catch(() => {});
    await navigateToAosManagement(page, 'app');
    return 'already exists (tenant not in wizard dropdown)';
  }
  await page.waitForTimeout(400);

  // Select the "Fly" service radio / card
  logStep(`${task.id} step 1: selecting Fly service`);
  const flyStrategies = [
    page.getByRole('radio',  { name: /^fly$/i }).first(),
    page.locator('label').filter({ hasText: /^fly$/i }).first(),
    page.locator('[class*="card"], [class*="item"], [class*="option"]')
        .filter({ hasText: /^fly$/i }).first(),
    page.getByText('Fly', { exact: true }).first(),
  ];
  let flySelected = false;
  for (const loc of flyStrategies) {
    if (await loc.isVisible({ timeout: 1500 }).catch(() => false)) {
      await loc.click(); flySelected = true; break;
    }
  }
  if (!flySelected) logWarn(`${task.id} Fly service option not found — proceeding`);
  await page.waitForTimeout(300);

  logStep(`${task.id} step 1: clicking Next`);
  await page.getByRole('button', { name: /^next$/i }).first().click({ timeout: 10_000 });
  await page.waitForTimeout(800);

  // ── STEP 2: Choose setup method ──────────────────────────────────────────
  logStep(`${task.id} wizard step 2: Choose setup method`);
  await page.getByText(/choose setup method/i).first()
            .waitFor({ state: 'visible', timeout: 10_000 });

  // Always select "Modern" mode (uses Microsoft OAuth — no manual credentials needed)
  logStep(`${task.id} step 2: selecting Modern mode`);
  const modernStrategies = [
    page.getByRole('radio', { name: /^modern$/i }).first(),
    page.locator('label').filter({ hasText: /^modern$/i }).first(),
    page.locator('[class*="card"], [class*="option"], [class*="item"]').filter({ hasText: /^modern$/i }).first(),
    page.getByText('Modern', { exact: true }).first(),
  ];
  let modernSelected = false;
  for (const loc of modernStrategies) {
    if (await loc.isVisible({ timeout: 1500 }).catch(() => false)) {
      await loc.click(); modernSelected = true; break;
    }
  }
  if (!modernSelected) logWarn(`${task.id} step 2: Modern option not found — using default`);
  await page.waitForTimeout(300);

  // Profile name field (present in some versions of the wizard on this step)
  const nameStrategies = [
    page.getByLabel(/^(profile\s*)?name\s*\*?$/i).first(),
    page.locator('xpath=//label[contains(normalize-space(.), "Name")]/following::input[@type!="hidden"][1]').first(),
  ];
  for (const f of nameStrategies) {
    if (await f.isVisible({ timeout: 1000 }).catch(() => false)) {
      logStep(`${task.id} step 2: filling profile name`);
      await f.fill(task.appProfileName); break;
    }
  }

  logStep(`${task.id} step 2: clicking Next`);
  await page.getByRole('button', { name: /^next$/i }).first().click({ timeout: 10_000 });
  await page.waitForTimeout(800);

  // ── STEP 3: Consent to apps ──────────────────────────────────────────────
  logStep(`${task.id} wizard step 3: Consent to apps`);
  await page.getByText(/consent to apps/i).first()
            .waitFor({ state: 'visible', timeout: 10_000 });

  const granted = await grantAdminConsent(page, task, page.locator('body'));

  // Click the final Finish / Save / Submit button
  logStep(`${task.id} step 3: clicking Finish`);
  const finishBtns = [
    page.getByRole('button', { name: /^(finish|save|done|submit|complete)$/i }).first(),
    page.getByRole('button', { name: /^next$/i }).first(),
  ];
  for (const b of finishBtns) {
    if (await b.isVisible({ timeout: 2_000 }).catch(() => false)) {
      await b.click({ timeout: 10_000 }); break;
    }
  }
  await page.waitForTimeout(1000);

  // Return to App Management list to confirm row exists
  logStep(`${task.id} returning to App Management list`);
  await navigateToAosManagement(page, 'app');

  return granted ? 'created + consent granted' : 'created (consent skipped — grant manually)';
}

// ---------------------------------------------------------------------------
// Grant Admin Consent for all consent links/buttons on the current page.
// The "Consent to apps" wizard step shows one <a> "Consent" link per app
// (Fly, Fly for Power Platform, Fly delegated app, etc.) — all must be clicked.
// Also handles the App Management list row case (button or link in a row).
// Returns true if at least one consent was processed.
// ---------------------------------------------------------------------------
async function grantAdminConsent(page, task, rowLocator) {
  logStep(`${task.id} grantAdminConsent: looking for consent links`);

  // Collect all visible "Consent" links/buttons — prefer links since the
  // wizard uses <a> elements, fall back to buttons for the list-row case.
  const candidateSels = [
    page.getByRole('link',   { name: /^consent$/i }),
    page.getByRole('button', { name: /^consent$/i }),
    page.locator('a').filter({ hasText: /^consent$/i }),
    page.locator('button').filter({ hasText: /^consent$/i }),
    page.getByRole('button', { name: /grant.*consent/i }),
  ];

  let consentLinks = [];
  for (const sel of candidateSels) {
    const n = await sel.count().catch(() => 0);
    for (let i = 0; i < n; i++) {
      if (await sel.nth(i).isVisible({ timeout: 300 }).catch(() => false)) {
        consentLinks.push(sel.nth(i));
      }
    }
    if (consentLinks.length > 0) break;
  }

  // Fallback: try row-level context menu for the App Management list case
  if (consentLinks.length === 0) {
    logStep(`${task.id} no consent links visible — trying row right-click`);
    await rowLocator.click({ button: 'right', timeout: 3_000 }).catch(() => {});
    await page.waitForTimeout(300);
    const ctxItem = page.locator('[role="menuitem"]').filter({ hasText: /consent/i }).first();
    if (await ctxItem.isVisible({ timeout: 1000 }).catch(() => false)) {
      consentLinks.push(ctxItem);
    } else {
      await page.keyboard.press('Escape').catch(() => {});
      logWarn(`${task.id} consent links not found — manual consent required`);
      return false;
    }
  }

  logStep(`${task.id} found ${consentLinks.length} consent link(s) — processing each`);
  let anyGranted = false;

  for (let i = 0; i < consentLinks.length; i++) {
    // Re-locate by index each iteration in case the DOM updated after a prior consent
    let link = consentLinks[i];
    if (!await link.isVisible({ timeout: 500 }).catch(() => false)) {
      logStep(`${task.id} consent ${i + 1}/${consentLinks.length} no longer visible — skipping`);
      continue;
    }

    logStep(`${task.id} clicking consent ${i + 1}/${consentLinks.length}`);
    const popupPromise = page.waitForEvent('popup', { timeout: 5_000 }).catch(() => null);
    await link.click({ timeout: 5_000 });

    const popup = await popupPromise;
    if (popup) {
      logStep(`${task.id} OAuth popup opened: ${popup.url()}`);
      emit({ event: 'info', message: `Admin consent ${i + 1}/${consentLinks.length}: approve in the Microsoft popup.` });
      // Wait for the popup to close on its own — AvePoint's OAuth callback page
      // calls window.close() once it has processed the consent response.
      // Do NOT close it early (e.g. on redirect) or the consent won't register.
      await popup.waitForEvent('close', { timeout: 300_000 });
      logStep(`${task.id} consent ${i + 1} popup closed`);
      await page.waitForTimeout(800);
    } else {
      // Popup not caught by Playwright — may have opened in an untracked window.
      // Wait a moment for the page to settle then continue.
      logWarn(`${task.id} consent ${i + 1} popup not tracked — waiting 3s`);
      await page.waitForTimeout(3_000);
    }

    await page.waitForTimeout(600);
    anyGranted = true;
  }

  logStep(`${task.id} grantAdminConsent: done`);
  return anyGranted;
}

(async () => {
  try {
    if      (RUN_MODE === 'login')  await login();
    else if (RUN_MODE === 'create') await create();
    else if (RUN_MODE === 'setup')  await setup();
    else { console.error('Usage: node fly-connector.js --mode=login|create|setup'); process.exit(1); }
  } catch (err) {
    logError(`FATAL: ${err && err.stack || err}`);
    emit({ event: 'fatal', message: err && err.stack || String(err) });
    process.exit(99);
  } finally {
    if (logStream) { logStream.end(); }
  }
})();