/* global React, ReactDOM, F, TopBar, NavRail, SettingsPanel, Home, FlyMenu, Monitor, AppReg, DomainRemoval, PageHeader, Page, Card, Button, Icon */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · app shell
// ============================================================================
const { useState: _uSa } = React;

function Stub({ title, trail, subtitle }) {
  return (
    <Page>
      <PageHeader title={title} subtitle={subtitle} trail={trail} />
      <Card style={{ display: 'flex', alignItems: 'center', gap: 14, color: F.muted }}>
        <Icon name="grid" size={22} color={F.faint} />
        <span style={{ fontSize: 14 }}>This surface exists in the toolkit but isn't part of the redesigned sample. The shell, nav and patterns above apply here too.</span>
      </Card>
    </Page>
  );
}

function FluentApp() {
  const [nav, setNav] = _uSa('home');
  const [sub, setSub] = _uSa(null); // appreg
  const [settings, setSettings] = _uSa(false);
  const [project, setProject] = _uSa('Contoso');

  const onNav = (id) => { setSub(null); setNav(id); };

  let body;
  if (sub === 'appreg') body = <AppReg onBack={() => setSub(null)} />;
  else if (nav === 'home') body = <Home onNav={onNav} onGear={() => setSettings(true)} />;
  else if (nav === 'fly') body = <FlyMenu onNav={onNav} onOpen={(id) => setSub(id)} />;
  else if (nav === 'monitor') body = <Monitor project={project} setProject={setProject} />;
  else if (nav === 'discovery') body = <Stub title="Discovery" subtitle="M365 tenant assessment — mailboxes, sites, OneDrive, groups and devices." trail={['Home', 'Discovery']} />;
  else if (nav === 'reports') body = <Stub title="Migration Reports" subtitle="Review per-user results and export status reports." trail={['Home', 'Reports']} />;
  else if (nav === 'domain') body = <DomainRemoval />;
  else if (nav === 'misc') body = <Stub title="Misc Scripts" subtitle="Utility and helper scripts." trail={['Home', 'Misc Scripts']} />;

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', background: F.shell, fontFamily: F.font, color: F.ink }}>
      <TopBar onGear={() => setSettings(true)} />
      <div style={{ flex: 1, display: 'flex', minHeight: 0 }}>
        <NavRail active={sub ? 'fly' : nav} onNav={onNav} />
        <main style={{ flex: 1, overflowY: 'auto', background: F.content }}>{body}</main>
      </div>
      {settings && <SettingsPanel onClose={() => setSettings(false)} />}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<FluentApp />);
