#!/usr/bin/env node
// Build executive briefs (HTML + PDF) from AvePoint FLY migration reports.
// Uses the same analysis/rendering code as FLY_Executive_Brief.html, so output matches the template.
//
// One brief per area:
//   node fly-brief.js --in "<export folder>"        --client "AEP Italia" [--out "<folder>"] [--workload OneDrive] [--internal "a.com,b.com"] [--no-pdf]
//   node fly-brief.js --report "<...Mappings.csv>"  --client "AEP Italia" [--out "<folder>"]   (only that report's own files)
//   node fly-brief.js --summary "<summary.json>"    --client "AEP Italia" --out "<folder>"     (pre-built summary, e.g. live FLY data)
// All areas in one document (summary page + one page per area, static HTML for email + PDF):
//   node fly-brief.js --pack "<a.summary.json>;<b.summary.json>;..." --client "AEP Italia" --out "<folder>" [--internal ...] [--no-pdf]
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
const AREA_ORDER = ['Exchange', 'OneDrive', 'SharePoint', 'Teams', 'M365Group', 'TeamChat'];

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

const readJson = f => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
const safe = s => s.replace(/[^\w.-]+/g, '_');

// Print an HTML file to PDF with headless Edge/Chrome. Returns the PDF path, or null.
async function printPdf(htmlOut, pdfOut, log) {
  const edge = findEdge();
  if (!edge) { log('Edge/Chrome not found; skipped PDF (set EDGE_PATH).'); return null; }
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
  if (!fs.existsSync(pdfOut)) { log('PDF was not produced by ' + path.basename(edge) + ' within 60s.'); return null; }
  return pdfOut;
}

// The CSVs that belong to one FLY report: same folder, same workload prefix, same export second.
// (A folder can hold several areas' reports, or copies; reading them all together mixes their data.)
function reportFiles(mappingsPath) {
  const dir = path.dirname(mappingsPath), name = path.basename(mappingsPath);
  const m = /^(\w+?)(?:Summary|Migration)Report_(\d{14})/i.exec(name);
  if (!m) return [mappingsPath];
  const re = new RegExp('^' + m[1] + '(Summary|Migration)Report_' + m[2] + '\\d*_.*\\.csv$', 'i');
  return fs.readdirSync(dir).filter(f => re.test(f)).map(f => path.join(dir, f));
}

async function analyseFiles(core, paths) {
  const files = [];
  for (const p of paths) {
    const name = path.basename(p);
    if (!/\.csv$/i.test(name) || /Site objects|Mailbox objects|Chat objects|Team objects/i.test(name)) continue; // item-level detail is not needed
    const blob = await fs.openAsBlob(p);
    files.push({ name, size: blob.size, stream: () => blob.stream() });
  }
  return core.analyse(files);
}

async function buildBrief({ input, report, summaryFile, client, workload, out, internal, noPdf, log = console.log }) {
  const { core, html } = loadCore();
  let summary;
  if (summaryFile) summary = readJson(summaryFile);     // pre-built (e.g. live FLY data)
  else if (report) summary = await analyseFiles(core, reportFiles(report));
  else {
    const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]);
    summary = await analyseFiles(core, walk(input));
  }
  workload = workload || summary.workload || '';
  const internalList = internal ? internal.split(',').map(s => s.trim()).filter(Boolean) : summary.domains;
  fs.mkdirSync(out, { recursive: true });
  const base = safe([client || 'Migration', workload].filter(Boolean).join(' ')) + '_Executive_Brief';
  fs.writeFileSync(path.join(out, base + '.summary.json'), JSON.stringify(summary, null, 1));
  const inject = '<script>window.__BRIEF__=' + JSON.stringify({ summary, client, workload, internal: internalList }).replace(/</g, '\\u003c') + ';</script>\n';
  fs.writeFileSync(path.join(out, base + '.html'), html.replace('<script id="core">', inject + '<script id="core">'));
  // Edge cannot open 8.3 short paths (C:\Users\ABC~1\...), so always hand it the long path.
  out = fs.realpathSync.native(out);
  const htmlOut = path.join(out, base + '.html');
  log('HTML: ' + htmlOut);
  if (noPdf) return { htmlOut, summary };
  const pdfOut = await printPdf(htmlOut, path.join(out, base + '.pdf'), log);
  if (pdfOut) log('PDF:  ' + pdfOut);
  return { htmlOut, pdfOut, summary };
}

// All areas in one static document: summary page, then one page per area. No scripts, so it renders
// the same when opened from an email attachment as it does in the PDF.
async function buildPack({ summaryFiles, client, out, internal, noPdf, log = console.log }) {
  const { core, html } = loadCore();
  const items = summaryFiles.map(f => { const s = readJson(f); return { summary: s, workload: s.workload }; })
    .sort((a, b) => AREA_ORDER.indexOf(a.workload) - AREA_ORDER.indexOf(b.workload));
  if (!items.length) throw new Error('No area summaries to combine.');
  const internalList = internal ? internal.split(',').map(s => s.trim()).filter(Boolean) : [...new Set(items.flatMap(i => i.summary.domains || []))];
  const opts = { client, internal: internalList };
  const head = /<head>([\s\S]*?)<\/head>/.exec(html)[1].replace(/<title>[\s\S]*?<\/title>/, '<title>' + (client || 'Migration').replace(/[<&]/g, '') + ' Microsoft 365 Migration - Executive Brief</title>');
  // Summary page lists every area; areas with no mappings in FLY get no page of their own
  const sheets = [core.renderOverview(items, opts)].concat(items.filter(i => (i.summary.m || {}).total > 0)
    .map(i => core.render(i.summary, Object.assign({ workload: i.workload }, opts))));
  const page = '<!doctype html>\n<html lang="en">\n<head>' + head + '</head>\n<body class="pack">\n' +
    sheets.map(s => '<div class="sheet">\n' + s + '\n</div>').join('\n') + '\n</body>\n</html>\n';
  fs.mkdirSync(out, { recursive: true });
  out = fs.realpathSync.native(out);
  const base = safe(client || 'Migration') + '_Migration_Executive_Brief';
  const htmlOut = path.join(out, base + '.html');
  fs.writeFileSync(htmlOut, page);
  log('PACK HTML: ' + htmlOut);
  if (noPdf) return { htmlOut };
  const pdfOut = await printPdf(htmlOut, path.join(out, base + '.pdf'), log);
  if (pdfOut) log('PACK PDF:  ' + pdfOut);
  return { htmlOut, pdfOut, areas: items.map(i => i.workload) };
}

module.exports = { buildBrief, buildPack, reportFiles };

if (require.main === module) {
  const a = args();
  const fail = e => { console.error(e.message); process.exit(1); };
  if (a.pack) {
    if (!a.out) { console.error('--out is required with --pack'); process.exit(1); }
    buildPack({ summaryFiles: a.pack.split(';').map(s => s.trim()).filter(Boolean), client: a.client || '', out: a.out, internal: a.internal, noPdf: a.noPdf }).catch(fail);
  } else {
    if (!a.in && !a.summary && !a.report) { console.error('Usage: see the header of fly-brief.js'); process.exit(1); }
    const out = a.out || a.in || (a.report && path.dirname(a.report));
    if (!out) { console.error('--out is required with --summary'); process.exit(1); }
    buildBrief({ input: a.in, report: a.report, summaryFile: a.summary, client: a.client || '', workload: a.workload || '', out, internal: a.internal, noPdf: a.noPdf }).catch(fail);
  }
}
