#!/usr/bin/env node
// Build the one-page executive brief (HTML + PDF) from a FLY report export folder.
// Uses the same analysis/rendering code as FLY_Executive_Brief.html, so output matches the template.
//
//   node fly-brief.js --in "<export folder>" --client "AEP Italia" [--out "<folder>"] [--workload OneDrive|SharePoint|Exchange|Teams|M365Group|TeamChat] [--internal "a.com,b.com"] [--no-pdf]
const fs = require('fs'), path = require('path'), { execFileSync } = require('child_process');

function args() {
  const a = {}, v = process.argv.slice(2);
  for (let i = 0; i < v.length; i++) {
    if (v[i] === '--no-pdf') a.noPdf = true;
    else if (v[i].startsWith('--')) a[v[i].slice(2)] = v[++i];
  }
  return a;
}
const TEMPLATE = path.join(__dirname, 'FLY_Executive_Brief.html');

function loadCore() {
  const html = fs.readFileSync(TEMPLATE, 'utf8');
  const m = /<script id="core">([\s\S]*?)<\/script>/.exec(html);
  if (!m) throw new Error('Core script block not found in template');
  const mod = { exports: {} };
  new Function('module', m[1])(mod);
  return { core: mod.exports, html };
}

function findEdge() {
  const c = [process.env.EDGE_PATH, 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'].filter(Boolean);
  return c.find(p => fs.existsSync(p));
}

async function buildBrief({ input, client, workload, out, internal, noPdf, log = console.log }) {
  const { core, html } = loadCore();
  const files = [];
  const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]);
  for (const p of walk(input)) {
    const name = path.basename(p);
    if (!/\.csv$/i.test(name)) continue;
    if (/Site objects/i.test(name)) continue; // not needed for the brief
    const blob = await fs.openAsBlob(p);
    files.push({ name, size: blob.size, stream: () => blob.stream() });
  }
  const summary = await core.analyse(files);
  const internalList = internal ? internal.split(',').map(s => s.trim()).filter(Boolean) : summary.domains;
  fs.mkdirSync(out, { recursive: true });
  const base = [client || 'Migration', workload].filter(Boolean).join(' ').replace(/[^\w.-]+/g, '_') + '_Executive_Brief';
  fs.writeFileSync(path.join(out, base + '.summary.json'), JSON.stringify(summary, null, 1));
  const inject = '<script>window.__BRIEF__=' + JSON.stringify({ summary, client, workload: workload || '', internal: internalList }).replace(/</g, '\\u003c') + ';</script>\n';
  const page = html.replace('<script id="core">', inject + '<script id="core">');
  fs.writeFileSync(path.join(out, base + '.html'), page);
  // Edge cannot open 8.3 short paths (C:\Users\ABC~1\...), so always hand it the long path.
  out = fs.realpathSync.native(out);
  const htmlOut = path.join(out, base + '.html');
  log('HTML: ' + htmlOut);
  if (!noPdf) {
    const edge = findEdge();
    if (!edge) { log('Edge/Chrome not found; skipped PDF (set EDGE_PATH).'); return { htmlOut, summary }; }
    const pdfOut = path.join(out, base + '.pdf');
    // A fresh profile per run: with a shared --user-data-dir, a second run while the previous headless
    // Edge is still alive hands its job to that instance and exits, so its PDF is never written.
    const profile = fs.mkdtempSync(path.join(require('os').tmpdir(), 'fly-brief-edge-'));
    try { fs.unlinkSync(pdfOut); } catch { /* no previous PDF */ }
    execFileSync(edge, ['--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--no-first-run', '--user-data-dir=' + profile,
      '--virtual-time-budget=5000', '--print-to-pdf=' + pdfOut, require('url').pathToFileURL(htmlOut).href], { stdio: 'ignore' });
    // msedge.exe can return before the PDF is flushed: wait until it exists and its size stops changing.
    const deadline = Date.now() + 60000;
    let last = -1;
    while (Date.now() < deadline) {
      const size = fs.existsSync(pdfOut) ? fs.statSync(pdfOut).size : -1;
      if (size > 0 && size === last) break;
      last = size;
      await new Promise(r => setTimeout(r, 500));
    }
    try { fs.rmSync(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 }); } catch { /* best effort */ }
    if (!fs.existsSync(pdfOut)) { log('PDF was not produced by ' + path.basename(edge) + ' within 60s.'); return { htmlOut, summary }; }
    log('PDF:  ' + pdfOut);
    return { htmlOut, pdfOut, summary };
  }
  return { htmlOut, summary };
}

module.exports = { buildBrief };

if (require.main === module) {
  const a = args();
  if (!a.in) { console.error('Usage: node fly-brief.js --in "<export folder>" --client "<name>" [--out "<folder>"] [--workload OneDrive|SharePoint|Exchange|Teams|M365Group|TeamChat] [--internal "a.com,b.com"] [--no-pdf]'); process.exit(1); }
  buildBrief({ input: a.in, client: a.client || '', workload: a.workload || '', out: a.out || a.in, internal: a.internal, noPdf: a.noPdf })
    .catch(e => { console.error(e.message); process.exit(1); });
}
