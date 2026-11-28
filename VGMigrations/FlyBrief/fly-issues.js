#!/usr/bin/env node
// FLY TechTeam Report: an interactive, read-only drill-down of migration problems for the tech team.
// Area -> issue type (grouped by FLY error code / message) or user/site -> the individual failed items.
//
//   node fly-issues.js --manifest "<manifest.json>" --out "<folder>"
//
// manifest: { client, areas: [ { workload, report: "<...Mappings.csv>" }            // downloaded FLY report
//                            | { workload, live: { project, mappings: [...] } } ] }  // live FLY mapping status
// Successful and skipped rows are left out. Very large issue types are capped (the count of the rest is shown).
const fs = require('fs'), path = require('path');
const { reportFiles } = require('./fly-brief.js');

const CAP_PER_GROUP = 1500, CAP_PER_AREA = 40000;
const AREA_ORDER = ['Exchange', 'OneDrive', 'SharePoint', 'Teams', 'M365Group', 'TeamChat'];
const AREA_TITLE = { Exchange: 'Exchange', OneDrive: 'OneDrive', SharePoint: 'SharePoint', Teams: 'Teams', M365Group: 'Microsoft 365 Groups', TeamChat: 'Teams Chat' };

function args() {
  const a = {}, v = process.argv.slice(2);
  for (let i = 0; i < v.length; i++) if (v[i].startsWith('--')) a[v[i].slice(2)] = v[++i];
  return a;
}

// Streaming CSV reader (quoted fields, embedded newlines). onRow(object) per data row.
function readCsv(file, onRow) {
  return new Promise((resolve, reject) => {
    let header = null, row = [], field = '', inQ = false, prevQ = false, first = true;
    const emit = () => {
      if (!header) { header = row.map(h => h.replace(/^﻿/, '').trim()); return; }
      if (row.length === 1 && row[0] === '') return;
      const o = {}; for (let i = 0; i < header.length; i++) o[header[i]] = row[i] || ''; onRow(o);
    };
    const s = fs.createReadStream(file, { encoding: 'utf8', highWaterMark: 1 << 20 });
    s.on('data', chunk => {
      if (first) { chunk = chunk.replace(/^﻿/, ''); first = false; }
      for (let i = 0; i < chunk.length; i++) {
        const c = chunk[i];
        if (inQ) { if (c === '"') { inQ = false; prevQ = true; } else field += c; continue; }
        if (c === '"') { if (prevQ) field += '"'; inQ = true; prevQ = false; continue; }
        prevQ = false;
        if (c === ',') { row.push(field); field = ''; }
        else if (c === '\n') { row.push(field.endsWith('\r') ? field.slice(0, -1) : field); field = ''; emit(); row = []; }
        else field += c;
      }
    });
    s.on('end', () => { if (field || row.length) { row.push(field); emit(); } resolve(); });
    s.on('error', reject);
  });
}

const lastSeg = u => { const p = String(u || '').split(/[\\/]/).filter(Boolean); return p.length ? p[p.length - 1] : ''; };
const trim = (s, n) => { s = String(s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
// Similar messages differ only in names, URLs, GUIDs and numbers: blank those out to group them
const norm = s => String(s || '').replace(/"[^"]*"/g, '"…"').replace(/https?:\/\/\S+/g, '<url>').replace(/[\w.+-]+@[\w.-]+/g, '<email>')
  .replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '<id>').replace(/\d+/g, '#').replace(/\s+/g, ' ').trim().slice(0, 200);
const severity = st => /error|fail|not found/i.test(st) ? 'Error' : /not supported/i.test(st) ? 'Info' : 'Warning';

// Owner (user / site / mailbox / team) of a row, as a stable key + label
function ownerOf(r) {
  const url = r['Source URL'];
  if (url) {
    let m = /\/personal\/([^/]+)/i.exec(url); if (m) return { key: 'p:' + m[1].toLowerCase(), label: m[1] };
    m = /^(https?:\/\/[^/]+\/(?:sites|teams)\/[^/]+)/i.exec(url); if (m) return { key: 's:' + m[1].toLowerCase(), label: m[1] };
    m = /^(https?:\/\/[^/]+)/i.exec(url); if (m) return { key: 's:' + m[1].toLowerCase(), label: m[1] };
  }
  const src = r['Source'] || r['Source email address'] || r['Source user'] || '';
  const head = String(src).split('\\')[0].trim();
  if (/@/.test(head)) return { key: 'e:' + head.toLowerCase(), label: head };
  if (head) return { key: 'n:' + head.toLowerCase(), label: head };
  return { key: '?', label: '(unknown)' };
}
function mappingKeys(src) {
  src = String(src || '').trim();
  if (/^https?:/i.test(src)) return [ownerOf({ 'Source URL': src }).key];
  if (/@/.test(src)) return ['e:' + src.toLowerCase(), 'p:' + src.toLowerCase().replace(/[@.\-]/g, '_')];
  return src ? ['n:' + src.toLowerCase()] : [];
}

function newArea(workload, source) {
  return { workload, title: AREA_TITLE[workload] || workload, source, mappings: [], owners: [], groups: [], items: [], msgs: [],
    _owner: {}, _group: {}, _msg: {}, counts: { Error: 0, Warning: 0, Info: 0, Security: 0 }, capped: 0 };
}
function ownerIdx(A, o) {
  if (A._owner[o.key] === undefined) { A._owner[o.key] = A.owners.length; A.owners.push({ label: o.label, mapping: null, n: 0 }); }
  return A._owner[o.key];
}
function msgIdx(A, s) { s = trim(s, 400); if (A._msg[s] === undefined) { A._msg[s] = A.msgs.length; A.msgs.push(s); } return A._msg[s]; }
function addIssue(A, kind, r, sev, code, comment, title, detail, pathStr) {
  const gk = kind + '|' + (code || norm(comment));
  let g = A._group[gk];
  // key: stable across regenerations, so the report's Fixed / Ignored marks survive a re-run
  if (g === undefined) { g = A._group[gk] = A.groups.length; A.groups.push({ key: gk, kind, sev, code: code || '', msg: trim(comment, 400), n: 0, kept: 0, owners: {} }); }
  const G = A.groups[g], o = ownerIdx(A, ownerOf(r));
  G.n++; G.owners[o] = (G.owners[o] || 0) + 1; A.owners[o].n++; A.counts[sev]++;
  if (G.kept >= CAP_PER_GROUP || A.items.length >= CAP_PER_AREA) { A.capped++; return; }
  G.kept++;
  // item: [group, owner, title, detail, path, status, message]
  A.items.push([g, o, trim(title, 200), trim(detail, 200), trim(pathStr, 300), r.Status || '', msgIdx(A, comment)]);
}

async function fromReport(workload, mappingsCsv, log) {
  const A = newArea(workload, 'FLY report ' + (/_(\d{4})(\d{2})(\d{2})/.exec(path.basename(mappingsCsv)) || []).slice(1).join('-'));
  for (const f of reportFiles(mappingsCsv)) {
    const name = path.basename(f), t0 = Date.now();
    let n = 0;
    if (/Mappings\.csv$/i.test(name)) {
      await readCsv(f, r => {
        n++;
        const srcCol = Object.keys(r).find(k => /^Source/i.test(k)), dstCol = Object.keys(r).find(k => /^Destination/i.test(k) && !/type/i.test(k));
        A.mappings.push({ src: r[srcCol] || '', dst: r[dstCol] || '', status: r['Last job status'] || '', comment: trim(r['Last job comment'] || r['Comment'], 300),
          gb: parseFloat(r['Migrated size (GB)']) || 0, ok: +r['Successful objects'] || 0, exc: +r['Objects with exceptions'] || 0, err: +r['Error objects'] || 0,
          skip: +r['Skipped objects'] || 0, start: (r['Last job start time'] || '').slice(0, 16), end: (r['Last job end time'] || '').slice(0, 16) });
      });
    } else if (/Site security\.csv$|_security\.csv$/i.test(name)) {
      await readCsv(f, r => {
        n++; const st = r.Status || ''; if (!st || /^(Successful|Skipped)$/i.test(st)) return;
        const who = r['Source display name'] || r['Source login name'];
        addIssue(A, 'Permission', r, severity(st), r['Error code'], r.Comment, who + (r.Permission ? ' – ' + r.Permission : ''),
          [r['User/Group'], r['Source login name'], r['Member of'] ? 'member of ' + r['Member of'] : ''].filter(Boolean).join(' · '), r['Source URL']);
      });
    } else if (/SharedLink/i.test(name)) {
      await readCsv(f, r => {
        n++; const st = r.Status || '', type = r['Shared link type'] || '', perm = r['Shared link permission'] || '';
        const title = lastSeg(r['Source URL']), detail = [type, perm, r['Shared with source user/group email'] ? 'shared with ' + r['Shared with source user/group email'] : ''].filter(Boolean).join(' · ');
        if (/anyone/i.test(type) && /edit/i.test(perm)) addIssue(A, 'Sharing link', r, 'Security', 'ANYONE-EDIT', 'Anonymous edit link: anyone who has the link can edit the file, with no sign-in.', title, detail, r['Source URL']);
        if (!st || /^(Successful|Skipped)$/i.test(st)) return;
        addIssue(A, 'Sharing link', r, severity(st), r['Error code'], r.Comment, title, detail, r['Source URL']);
      });
    } else if (/objects(\(\d+\))?\.csv$|configurations\.csv$/i.test(name)) {
      const kind = /configurations/i.test(name) ? 'Configuration' : 'Item';
      await readCsv(f, r => {
        n++; const st = r.Status || ''; if (!st || /^(Successful|Skipped)$/i.test(st)) return;
        const title = r.Title || (r.Object ? r.Object + ': ' : '') + lastSeg(r['Source URL']);
        addIssue(A, kind, r, severity(st), r['Error code'], r.Comment, title, [r.Type || r.Object, r.Size].filter(Boolean).join(' · '), r['Source URL'] || r.Source);
      });
    } else continue;
    log(`  ${name}: ${n.toLocaleString()} rows (${((Date.now() - t0) / 1000).toFixed(1)}s)`);
  }
  return A;
}

// Live FLY mapping status (no report downloaded): every mapping that did not finish cleanly is an issue
function fromLive(workload, live) {
  const A = newArea(workload, 'FLY live data (' + live.project + ')');
  const ST = { 0: 'Not started', 1: 'In progress', 2: 'Stopped', 3: 'Waiting', 4: 'Finished', 5: 'Finished with exceptions', 6: 'Failed', 7: 'Scheduled' };
  for (const m of live.mappings || []) {
    const status = ST[m.status] || ('Status ' + m.status);
    A.mappings.push({ src: m.src, dst: m.dst, status, comment: '', gb: m.gb || 0, ok: null, exc: null, err: m.errorItems || 0, skip: null, start: m.start || '', end: '' });
    if (m.status === 4) continue;
    const r = /^https?:/i.test(m.src) ? { 'Source URL': m.src, Status: status } : { Source: m.src, Status: status };
    const sev = m.status === 6 ? 'Error' : m.status === 5 ? 'Warning' : 'Info';
    const msg = m.status === 6 ? 'The last migration job failed for this mapping.' : m.status === 5 ? 'The last migration job finished, but some items failed or have warnings.'
      : m.status === 0 ? 'This mapping has not been migrated yet.' : 'The mapping is ' + status.toLowerCase() + '.';
    addIssue(A, 'Mapping', r, sev, 'MAPPING-' + status.toUpperCase().replace(/\W+/g, '-'), msg, m.src,
      (m.errorItems ? m.errorItems.toLocaleString() + ' items with errors (FLY live counter)' : '') + (m.gb ? ' · ' + m.gb.toFixed(1) + ' GB migrated' : ''), m.dst ? '→ ' + m.dst : '');
  }
  return A;
}

function finish(A) {
  // link owners to their mapping row
  const mk = {}; A.mappings.forEach((m, i) => mappingKeys(m.src).forEach(k => { mk[k] = i; }));
  Object.keys(A._owner).forEach(k => { const i = mk[k]; if (i !== undefined) { const o = A.owners[A._owner[k]]; o.mapping = i; o.label = A.mappings[i].src || o.label; } });
  A.groups.forEach(G => { G.users = Object.keys(G.owners).length; G.top = Object.entries(G.owners).sort((a, b) => b[1] - a[1]).slice(0, 200).map(([o, n]) => [+o, n]); delete G.owners; });
  delete A._owner; delete A._group; delete A._msg;
  return A;
}

async function main() {
  const a = args();
  if (!a.manifest || !a.out) { console.error('Usage: node fly-issues.js --manifest "<manifest.json>" --out "<folder>"'); process.exit(1); }
  const man = JSON.parse(fs.readFileSync(a.manifest, 'utf8').replace(/^﻿/, ''));
  const areas = [], used = {};
  for (const ar of (man.areas || []).sort((x, y) => AREA_ORDER.indexOf(x.workload) - AREA_ORDER.indexOf(y.workload))) {
    console.log(`[${ar.workload}] ${ar.report ? 'reading ' + path.basename(ar.report) + ' and its detail files' : 'live FLY mapping status'}...`);
    const A = ar.report ? await fromReport(ar.workload, ar.report, console.log) : fromLive(ar.workload, ar.live);
    // unique id per area (a customer can have two projects for one workload, e.g. "OneDrive - USA" / "OneDrive - CAN")
    used[ar.workload] = (used[ar.workload] || 0) + 1;
    A.id = used[ar.workload] > 1 ? ar.workload + used[ar.workload] : ar.workload;
    if (ar.live && ar.live.project) A.title += ' (' + ar.live.project + ')';
    areas.push(finish(A));
    console.log(`[${ar.workload}] ${A.counts.Error.toLocaleString()} errors, ${A.counts.Warning.toLocaleString()} warnings, ${A.counts.Security.toLocaleString()} security, ${A.counts.Info.toLocaleString()} info in ${A.groups.length} issue types` + (A.capped ? ` (${A.capped.toLocaleString()} items beyond the per-type cap not listed)` : ''));
  }
  const data = { client: man.client || 'Migration', generated: new Date().toISOString(), areas };
  fs.mkdirSync(a.out, { recursive: true });
  const out = path.join(fs.realpathSync.native(a.out), (man.client || 'Migration').replace(/[^\w.-]+/g, '_') + '_TechTeam_Report.html');
  const tpl = fs.readFileSync(path.join(__dirname, 'FLY_TechTeam_Report.html'), 'utf8');
  fs.writeFileSync(out, tpl.replace('/*__DATA__*/null', JSON.stringify(data).replace(/</g, '\\u003c')));
  console.log('TECH HTML: ' + out);
}
main().catch(e => { console.error(e.stack || e.message); process.exit(1); });
