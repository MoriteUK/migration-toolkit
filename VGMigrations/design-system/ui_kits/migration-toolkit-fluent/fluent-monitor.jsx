/* global React, F, Icon, Button, Card, Pill, Field, Select, Toggle, PageHeader, CommandBar, CommandDivider, Page */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · Monitor list + App Registration
// ============================================================================
const { useState: _uSm, useEffect: _uEm, useRef: _uRm } = React;

const MONITOR_ROWS = [
  { project: 'Contoso — SharePoint', icon: 'cloud', total: 412, notStarted: 0, inProgress: 38, complete: 374, failed: 0, warnings: 0 },
  { project: 'Contoso — Exchange', icon: 'mail', total: 206, notStarted: 4, inProgress: 12, complete: 187, failed: 0, warnings: 3 },
  { project: 'Contoso — OneDrive', icon: 'folder', total: 318, notStarted: 0, inProgress: 5, complete: 313, failed: 0, warnings: 0 },
  { project: 'Contoso — Teams', icon: 'teams', total: 58, notStarted: 10, inProgress: 2, complete: 44, failed: 2, warnings: 0 },
  { project: 'Contoso — Teams Chat', icon: 'teams', err: 'Pre-scan not supported for this workload' },
  { project: 'Contoso — Groups', icon: 'users', total: 73, notStarted: 1, inProgress: 0, complete: 72, failed: 0, warnings: 0 },
];

function statusFor(r) {
  if (r.err) return { tone: 'grey', label: 'Not available' };
  if (r.failed > 0) return { tone: 'red', label: r.failed + ' failed' };
  if (r.warnings > 0) return { tone: 'amber', label: r.warnings + ' warnings' };
  if (r.inProgress > 0) return { tone: 'blue', label: 'In progress' };
  return { tone: 'green', label: 'Complete' };
}

function Monitor({ project, setProject }) {
  const [sel, setSel] = _uSm(0);
  const [auto, setAuto] = _uSm(true);
  const cols = ['Project', 'Status', 'Total', 'Not started', 'In progress', 'Complete', 'Failed', 'Warnings'];
  const align = ['left', 'left', 'right', 'right', 'right', 'right', 'right', 'right'];
  return (
    <Page>
      <PageHeader title="Project Monitor" subtitle="Live migration progress · last refreshed 14:42:14"
        trail={['Home', 'AvePoint Fly', 'Project Monitor']} />
      <CommandBar
        right={<>
          <Toggle checked={auto} onChange={setAuto} label="Auto refresh" />
          <Select value="5 min" options={['1 min', '2 min', '5 min', '10 min', '15 min', '30 min']} onChange={() => {}} width={92} />
        </>}>
        <Button variant="subtle" icon="sync">Refresh now</Button>
        <Button variant="subtle" icon="download">Export status</Button>
        <CommandDivider />
        <Select value={project} options={['Contoso', 'Fabrikam', 'Northwind']} onChange={setProject} width={150} />
        <Button variant="subtle" icon="filter">Filter</Button>
      </CommandBar>

      <Card pad={0} style={{ overflow: 'hidden' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 13.5, fontFamily: F.font }}>
          <thead>
            <tr>
              {cols.map((c, i) => (
                <th key={c} style={{
                  textAlign: align[i], padding: '11px 16px', fontSize: 12, fontWeight: 600, color: F.muted,
                  textTransform: 'uppercase', letterSpacing: '.03em', borderBottom: `1px solid ${F.border}`,
                  background: '#fbfbfd', whiteSpace: 'nowrap',
                }}>{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {MONITOR_ROWS.map((r, ri) => {
              const st = statusFor(r);
              const isSel = sel === ri;
              return (
                <tr key={ri} onClick={() => setSel(ri)} style={{
                  cursor: 'pointer', background: isSel ? F.accentTint : 'transparent',
                  borderBottom: `1px solid ${F.divider}`,
                }}
                  onMouseEnter={(e) => { if (!isSel) e.currentTarget.style.background = '#f7f8fa'; }}
                  onMouseLeave={(e) => { if (!isSel) e.currentTarget.style.background = 'transparent'; }}>
                  <td style={{ padding: '12px 16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 11 }}>
                      <span style={{ width: 30, height: 30, borderRadius: F.radiusSm, background: '#f3f4f7', display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
                        <Icon name={r.icon} size={17} color={F.inkSoft} />
                      </span>
                      <span style={{ fontWeight: 600, color: F.ink, whiteSpace: 'nowrap' }}>{r.project}</span>
                    </div>
                  </td>
                  <td style={{ padding: '12px 16px' }}><Pill tone={st.tone}>{st.label}</Pill></td>
                  {r.err ? (
                    <td colSpan={6} style={{ padding: '12px 16px', color: F.muted, fontStyle: 'italic' }}>{r.err}</td>
                  ) : (
                    <Nums r={r} />
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      <p style={{ fontSize: 12.5, color: F.faint, marginTop: 12 }}>
        Showing {MONITOR_ROWS.length} workloads · select a row to view per-item detail.
      </p>
    </Page>
  );
}

function Num({ children, strong, tone }) {
  return <td style={{
    padding: '12px 16px', textAlign: 'right', fontVariantNumeric: 'tabular-nums',
    fontWeight: strong ? 600 : 400, color: tone || F.inkSoft,
  }}>{children}</td>;
}
function Nums({ r }) {
  return (<>
    <Num strong>{r.total}</Num>
    <Num tone={F.muted}>{r.notStarted}</Num>
    <Num tone={r.inProgress ? F.blue : F.muted}>{r.inProgress}</Num>
    <Num tone={F.green} strong>{r.complete}</Num>
    <Num tone={r.failed ? F.red : F.muted} strong={!!r.failed}>{r.failed}</Num>
    <Num tone={r.warnings ? F.amber : F.muted} strong={!!r.warnings}>{r.warnings}</Num>
  </>);
}

// ── App Registration workflow form ──────────────────────────────────────────
const LVL = { INFO: '#7aa2e8', OK: '#4cc97a', WARN: '#e3b341', ERROR: '#e86a6a' };
function AppReg({ onBack }) {
  const [tenant, setTenant] = _uSm('contoso.onmicrosoft.com');
  const [appName, setAppName] = _uSm('AvePoint Fly Migration');
  const [lines, setLines] = _uSm([{ ts: '14:30:02', level: 'INFO', msg: 'Ready. Sign in as Global Administrator to begin.' }]);
  const ref = _uRm(null);
  _uEm(() => { if (ref.current) ref.current.scrollTop = ref.current.scrollHeight; }, [lines]);
  const run = () => {
    const seq = [
      { level: 'INFO', msg: 'Connecting to Microsoft Graph…' },
      { level: 'OK', msg: 'Signed in as admin@' + tenant },
      { level: 'INFO', msg: 'Creating app registration "' + appName + '"…' },
      { level: 'INFO', msg: 'Granting Sites.FullControl.All, Mail.ReadWrite, Group.ReadWrite.All…' },
      { level: 'OK', msg: 'App registered. Client ID 8f3c2a91-… secret expires in 24 months.' },
    ];
    seq.forEach((s, i) => setTimeout(() => {
      const ts = new Date(Date.now() + i * 1000).toLocaleTimeString('en-GB', { hour12: false });
      setLines((p) => [...p, { ts, ...s }]);
    }, i * 360));
  };
  return (
    <Page>
      <PageHeader title="Create App Registration" subtitle="Step 1 of 5 · register the Entra ID app and grant API permissions"
        trail={['Home', 'AvePoint Fly', 'App Registration']}
        actions={<Button variant="default" onClick={onBack}>Back to workflow</Button>} />
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.1fr', gap: 20, alignItems: 'start' }}>
        <Card>
          <h2 style={{ margin: '0 0 16px', fontSize: 16, fontWeight: 600, color: F.ink }}>Entra ID application</h2>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <Field label="Tenant domain" value={tenant} onChange={setTenant} />
            <Field label="Application name" value={appName} onChange={setAppName} hint="Shown in the Entra ID app registrations list." />
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
            <Button variant="primary" icon="key" onClick={run}>Create registration</Button>
            <Button variant="default" onClick={() => setLines([{ ts: '14:30:02', level: 'INFO', msg: 'Ready.' }])}>Reset</Button>
          </div>
        </Card>
        <div>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: F.muted, textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 8 }}>Activity log</div>
          <div ref={ref} style={{
            background: F.logBg, border: `1px solid ${F.logBorder}`, borderRadius: F.radius,
            padding: '14px 16px', height: 232, overflowY: 'auto', fontFamily: F.mono, fontSize: 12.5, lineHeight: 1.6,
          }}>
            {lines.map((l, i) => (
              <div key={i}>
                <span style={{ color: '#5a6178' }}>{l.ts}</span>{' '}
                <span style={{ color: LVL[l.level], fontWeight: 600 }}>[{l.level}]</span>{' '}
                <span style={{ color: '#cdd2e2' }}>{l.msg}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </Page>
  );
}

Object.assign(window, { Monitor, AppReg });
