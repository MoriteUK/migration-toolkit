/* global React, FLY, Window, Header, Footer, FooterButton, Tile, Button, Card, Field, Select, Checkbox, StatusDot, LogConsole, MonitorGrid */
// ============================================================================
// Fly Migration Toolkit — UI Kit · screens
// Recreates: Launcher, sub-menus, AvePoint Fly menu, App Registration form,
// Project Monitor and the Settings dialog. Static/cosmetic recreations.
// ============================================================================
const { useState: uS, useEffect: uE } = React;

// ── Launcher (main-menu.ps1 · Show-Launcher) ───────────────────────────────
function Launcher({ go, gear, close, updateAvailable }) {
  return (
    <Window title="Migration Tools" width={480} onGear={gear} onClose={close}>
      {updateAvailable && (
        <div style={{
          height: 36, background: FLY.bannerWarn, display: 'flex', alignItems: 'center',
          padding: '0 8px', gap: 8,
        }}>
          <span style={{ fontWeight: 600, fontSize: 13, color: '#212121', flex: 1 }}>
            &nbsp;Update available: v2.0.3 → v2.1.0
          </span>
          <Button onClick={() => {}} style={{ height: 26, fontSize: 12 }}>Install Update</Button>
        </div>
      )}
      <div style={{ padding: '26px 40px 6px' }}>
        <Tile label="Discovery" subtitle="M365 tenant assessment — mailboxes, sites, OneDrive, groups" onClick={() => go('discovery')} />
        <Tile label="AvePoint Fly" subtitle="Migration toolkit — connections, mappings, monitoring" onClick={() => go('fly')} />
        <Tile label="Misc Scripts" subtitle="Utility and helper scripts" onClick={() => go('misc')} />
        <Tile label="Domain Removal" subtitle="Scripts for removing and cleaning up domains" onClick={() => go('domain')} />
      </div>
    </Window>
  );
}

// ── Generic sub-menu shell ─────────────────────────────────────────────────
function SubMenu({ title, gear, back, tiles }) {
  return (
    <Window title={title} width={480} onGear={gear} onClose={back}>
      <div style={{ padding: '26px 40px 6px' }}>
        {tiles.map((t, i) => <Tile key={i} {...t} />)}
      </div>
    </Window>
  );
}

// ── AvePoint Fly menu (menu.ps1) ───────────────────────────────────────────
function FlyMenu({ go, gear, back }) {
  const items = [
    { label: '1. Create App Registration', subtitle: 'Register the Entra ID app and grant required API permissions', onClick: () => go('appreg') },
    { label: '2. Setup AOS Tenant & App', subtitle: 'Configure the AvePoint Online Services tenant and application' },
    { label: '3. Connections & Migration Mappings', subtitle: 'Manage connections, source/destination accounts and job mappings' },
    { label: '4. View Migration Reports', subtitle: 'Review per-user migration results and export status reports' },
    { label: '5. Monitor Projects', subtitle: 'Live project monitoring and migration progress tracking', onClick: () => go('monitor') },
  ];
  return <SubMenu title="Migration Toolkit" gear={gear} back={back} tiles={items} />;
}

// ── App Registration workflow form (card + fields + log console) ───────────
function AppRegForm({ gear, back }) {
  const [tenant, setTenant] = uS('contoso.onmicrosoft.com');
  const [appName, setAppName] = uS('AvePoint Fly Migration');
  const [lines, setLines] = uS([
    { ts: '14:30:02', level: 'INFO', msg: 'Ready. Sign in as Global Administrator to begin.' },
  ]);
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
    }, i * 380));
  };
  return (
    <Window title="Create App Registration" width={560} onGear={gear} onClose={back}>
      <div style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: 14 }}>
        <Card title="Entra ID App">
          <Field label="Tenant domain" value={tenant} onChange={setTenant} />
          <Field label="Application name" value={appName} onChange={setAppName} />
          <div style={{ display: 'flex', gap: 10, marginTop: 4 }}>
            <Button onClick={run}>Create Registration</Button>
            <Button variant="secondary" onClick={() => setLines([{ ts: '14:30:02', level: 'INFO', msg: 'Ready.' }])}>Reset</Button>
          </div>
        </Card>
        <div>
          <div style={{ fontSize: 10, fontWeight: 600, letterSpacing: '.06em', textTransform: 'uppercase', color: FLY.muted, marginBottom: 6 }}>Activity log</div>
          <LogConsole lines={lines} height={170} />
        </div>
      </div>
    </Window>
  );
}

// ── Project Monitor (monitor.ps1) ──────────────────────────────────────────
const MON_ROWS = [
  { project: 'Contoso - SharePoint', total: 412, notStarted: 0, inProgress: 38, complete: 374, failed: 0, warnings: 0, refresh: '14:42:10' },
  { project: 'Contoso - Exchange', total: 206, notStarted: 4, inProgress: 12, complete: 187, failed: 0, warnings: 3, refresh: '14:42:11' },
  { project: 'Contoso - OneDrive', total: 318, notStarted: 0, inProgress: 5, complete: 313, failed: 0, warnings: 0, refresh: '14:42:12' },
  { project: 'Contoso - Teams', total: 58, notStarted: 10, inProgress: 2, complete: 44, failed: 2, warnings: 0, refresh: '14:42:13' },
  { project: 'Contoso - Teams Chat', err: 'PreScan not supported for this workload', refresh: '14:42:13' },
  { project: 'Contoso - Groups', total: 73, notStarted: 1, inProgress: 0, complete: 72, failed: 0, warnings: 0, refresh: '14:42:14' },
];

function ProjectMonitor({ gear, back }) {
  const [sel, setSel] = uS(0);
  const [auto, setAuto] = uS(true);
  const [project, setProject] = uS('Contoso');
  const connExtra = (
    <>
      <StatusDot color={FLY.green} size={12} />
      <span style={{ color: '#78e678', fontSize: 12.5 }}>Connected: contoso.avepointonline.com</span>
    </>
  );
  return (
    <Window title="Project Monitor" width={920} onGear={gear} onClose={back}
      footerColor={FLY.footerAlt} footerExtra={connExtra}>
      <div style={{ display: 'flex', flexDirection: 'column', height: 520 }}>
        {/* selector bar */}
        <div style={{
          height: 50, background: FLY.selTint, display: 'flex', alignItems: 'center',
          gap: 12, padding: '0 14px', flex: 'none',
        }}>
          <span style={{ fontWeight: 600, fontSize: 13 }}>Project:</span>
          <Select value={project} onChange={setProject} width={200}
            options={['Contoso', 'Fabrikam', 'Northwind']} />
          <Button onClick={() => {}} style={{ height: 28 }}>Refresh Now</Button>
          <Checkbox label="Auto refresh" checked={auto} onChange={setAuto} />
          <Select value="5 min" onChange={() => {}} width={84}
            options={['1 min', '2 min', '5 min', '10 min', '15 min', '30 min']} />
          <span style={{ fontSize: 12.5, color: FLY.muted, marginLeft: 4 }}>Last refresh: 14:42:14</span>
        </div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <MonitorGrid rows={MON_ROWS} selected={sel} onSelect={setSel} />
        </div>
      </div>
    </Window>
  );
}

// ── Settings dialog (modal, tabbed) ────────────────────────────────────────
function SettingsDialog({ onClose }) {
  const [tab, setTab] = uS('Config');
  const tabs = ['Config', 'Customer', 'Workloads', 'Discovery'];
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(15,20,35,.45)', display: 'flex',
      alignItems: 'center', justifyContent: 'center', zIndex: 50,
    }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: 520, background: FLY.bg, borderRadius: 4, overflow: 'hidden',
        boxShadow: '0 24px 60px rgba(10,15,35,.4)', fontFamily: FLY.font,
        display: 'flex', flexDirection: 'column',
      }}>
        <Header title="Settings" height={48} onGear={() => {}} />
        <div style={{ display: 'flex', gap: 2, background: '#dfe3ee', padding: '0 8px' }}>
          {tabs.map((t) => (
            <button key={t} onClick={() => setTab(t)} style={{
              border: 0, background: tab === t ? FLY.bg : 'transparent', cursor: 'pointer',
              padding: '10px 16px', fontSize: 13, fontWeight: tab === t ? 600 : 400,
              color: tab === t ? FLY.accent : FLY.muted, fontFamily: FLY.font,
              borderBottom: tab === t ? `2px solid ${FLY.accent}` : '2px solid transparent',
            }}>{t}</button>
          ))}
        </div>
        <div style={{ padding: 20, minHeight: 220 }}>
          {tab === 'Config' && (
            <Card title="Fly API">
              <Field label="Fly API URL" value="https://contoso.avepointonline.com" onChange={() => {}} />
              <Field label="Client ID" value="8f3c2a91-7d4e-…-a91b" onChange={() => {}} />
              <Field label="Client Secret" value="••••••••••••••••" password onChange={() => {}} />
            </Card>
          )}
          {tab === 'Customer' && (
            <Card title="Customer">
              <Field label="Tenant prefix" value="CONTOSO" onChange={() => {}} />
              <Field label="SharePoint admin URL" value="https://contoso-admin.sharepoint.com" onChange={() => {}} />
              <Field label="Secret expiry" value="2028-05-31" onChange={() => {}} />
            </Card>
          )}
          {tab === 'Workloads' && (
            <Card title="Project suffixes">
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px 16px' }}>
                {['SharePoint', 'Exchange', 'OneDrive', 'Teams', 'Teams Chat', 'Groups'].map((w) => (
                  <Field key={w} label={w} value={w} onChange={() => {}} />
                ))}
              </div>
            </Card>
          )}
          {tab === 'Discovery' && (
            <Card title="Discovery output">
              <Field label="Discovery output path" value="C:\\Migration\\Discovery" onChange={() => {}} />
            </Card>
          )}
        </div>
        <Footer onClose={onClose} extra={<Button variant="primary" onClick={onClose}>Save</Button>} />
      </div>
    </div>
  );
}

Object.assign(window, { Launcher, SubMenu, FlyMenu, AppRegForm, ProjectMonitor, SettingsDialog });
