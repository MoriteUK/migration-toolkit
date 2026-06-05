/* global React, ReactDOM, FLY, Launcher, SubMenu, FlyMenu, AppRegForm, ProjectMonitor, SettingsDialog */
// ============================================================================
// Fly Migration Toolkit — UI Kit · interactive app shell
// Swaps "windows" the way the WinForms launcher opens child dialogs.
// ============================================================================
const { useState: uSt } = React;

function App() {
  const [screen, setScreen] = uSt('launcher'); // launcher | discovery | fly | misc | domain | appreg | monitor
  const [settings, setSettings] = uSt(false);
  const go = (s) => setScreen(s);
  const gear = () => setSettings(true);
  const home = () => setScreen('launcher');
  const backFly = () => setScreen('fly');

  let view;
  switch (screen) {
    case 'discovery':
      view = <SubMenu title="Discovery Tools" gear={gear} back={home} tiles={[
        { label: 'M365 Discovery', subtitle: 'M365 tenant assessment — mailboxes, sites, OneDrive, groups, devices' },
      ]} />;
      break;
    case 'misc':
      view = <SubMenu title="Misc Scripts" gear={gear} back={home} tiles={[
        { label: 'Provision OneDrives', subtitle: 'Pre-provision OneDrive for Business sites from a mapping file' },
        { label: 'Set Teams Owners', subtitle: 'Add a user as owner to Teams and M365 Groups from a CSV' },
      ]} />;
      break;
    case 'domain':
      view = <SubMenu title="Domain Removal" gear={gear} back={home} tiles={[
        { label: 'Domain Removal Workflow', danger: true, subtitle: '3-step workflow: Update on-prem UPN → AD Sync → Remove domain' },
        { label: 'Update On-Premise UPNs', subtitle: 'Update UPN, email, and aliases in on-premise Active Directory' },
        { label: 'Run AD Sync', subtitle: 'Trigger Azure AD Connect delta sync on VOL-ane-aad1 server' },
        { label: 'Remove Domain', subtitle: 'Remove a verified domain and all associated M365 objects' },
        { label: 'Hide from Address Book', subtitle: 'Bulk hide Exchange Online recipients from the Global Address List' },
      ]} />;
      break;
    case 'fly':
      view = <FlyMenu go={go} gear={gear} back={home} />;
      break;
    case 'appreg':
      view = <AppRegForm gear={gear} back={backFly} />;
      break;
    case 'monitor':
      view = <ProjectMonitor gear={gear} back={backFly} />;
      break;
    default:
      view = <Launcher go={go} gear={gear} close={home} updateAvailable />;
  }

  return (
    <div style={{
      minHeight: '100vh', width: '100%', background:
        'radial-gradient(1200px 700px at 50% -10%, #e8edf6 0%, #f0f2f7 45%, #e6e9f1 100%)',
      display: 'flex', alignItems: 'flex-start', justifyContent: 'center',
      padding: '52px 24px 64px', boxSizing: 'border-box', fontFamily: FLY.font,
    }}>
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 18 }}>
        <Crumb screen={screen} go={go} home={home} />
        {view}
      </div>
      {settings && <SettingsDialog onClose={() => setSettings(false)} />}
    </div>
  );
}

function Crumb({ screen, go, home }) {
  const labels = {
    launcher: 'Launcher', discovery: 'Discovery', fly: 'AvePoint Fly',
    misc: 'Misc Scripts', domain: 'Domain Removal', appreg: 'AvePoint Fly · App Registration',
    monitor: 'AvePoint Fly · Project Monitor',
  };
  return (
    <div style={{
      fontSize: 12, color: FLY.muted, letterSpacing: '.02em', display: 'flex',
      alignItems: 'center', gap: 8, fontFamily: FLY.font,
    }}>
      <button onClick={home} style={{
        border: 0, background: 'none', color: screen === 'launcher' ? FLY.accent : FLY.muted,
        cursor: 'pointer', fontSize: 12, fontFamily: FLY.font, padding: 0,
        fontWeight: screen === 'launcher' ? 600 : 400,
      }}>Migration Tools</button>
      {screen !== 'launcher' && <span>›</span>}
      {screen !== 'launcher' && <span style={{ color: FLY.text, fontWeight: 600 }}>{labels[screen]}</span>}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
