/* global React, F, Icon, Button, Card, Field, Select, Toggle, SearchBox */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · chrome
// TopBar (suite bar), NavRail (M365 admin nav), CommandBar, PageHeader,
// Breadcrumb, SettingsPanel (right slide-over with pivot tabs).
// ============================================================================
const { useState: _uSc } = React;

// ── Suite / top bar ──────────────────────────────────────────────────────────
function TopBar({ onGear }) {
  return (
    <div style={{
      height: 48, background: '#fff', borderBottom: `1px solid ${F.border}`,
      display: 'flex', alignItems: 'center', gap: 12, padding: '0 14px', flex: 'none', zIndex: 5,
    }}>
      <span style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', width: 32, height: 32, borderRadius: F.radiusSm, color: F.muted, cursor: 'pointer' }}>
        <Icon name="waffle" size={20} />
      </span>
      <img src="../../assets/fly-logo-64.png" alt="" width={24} height={24} style={{ borderRadius: 5, display: 'block' }} />
      <span style={{ fontSize: 15, fontWeight: 600, color: F.ink }}>Fly Migration Toolkit</span>
      <div style={{ flex: 1, display: 'flex', justifyContent: 'center' }}>
        <SearchBox placeholder="Search projects, users, mappings" width={440} />
      </div>
      <button onClick={onGear} title="Settings" style={{
        width: 32, height: 32, border: 0, background: 'transparent', borderRadius: F.radiusSm,
        cursor: 'pointer', color: F.muted, display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}><Icon name="settings" size={20} /></button>
      <span style={{
        width: 30, height: 30, borderRadius: '50%', background: F.accent, color: '#fff',
        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12.5, fontWeight: 600,
      }}>MA</span>
    </div>
  );
}

// ── Left navigation rail ─────────────────────────────────────────────────────
const NAV = [
  { type: 'item', id: 'home', label: 'Home', icon: 'home' },
  { type: 'group', label: 'Migrate' },
  { type: 'item', id: 'discovery', label: 'Discovery', icon: 'search' },
  { type: 'item', id: 'fly', label: 'AvePoint Fly', icon: 'send' },
  { type: 'item', id: 'monitor', label: 'Project Monitor', icon: 'chart' },
  { type: 'item', id: 'reports', label: 'Reports', icon: 'doc' },
  { type: 'group', label: 'Clean up' },
  { type: 'item', id: 'domain', label: 'Domain Removal', icon: 'trash' },
  { type: 'item', id: 'misc', label: 'Misc Scripts', icon: 'grid' },
];

function NavRail({ active, onNav }) {
  return (
    <nav style={{
      width: 240, background: F.nav, borderRight: `1px solid ${F.border}`, flex: 'none',
      padding: '10px 8px', display: 'flex', flexDirection: 'column', gap: 1, overflowY: 'auto',
    }}>
      {NAV.map((n, i) => n.type === 'group' ? (
        <div key={i} style={{
          fontSize: 11, fontWeight: 600, letterSpacing: '.04em', textTransform: 'uppercase',
          color: F.faint, padding: '14px 12px 6px',
        }}>{n.label}</div>
      ) : (
        <NavItem key={n.id} item={n} active={active === n.id} onClick={() => onNav(n.id)} />
      ))}
    </nav>
  );
}

function NavItem({ item, active, onClick }) {
  const [h, setH] = _uSc(false);
  return (
    <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        position: 'relative', display: 'flex', alignItems: 'center', gap: 11, width: '100%',
        height: 38, padding: '0 12px', border: 0, borderRadius: F.radiusSm, cursor: 'pointer',
        background: active ? F.accentTint : h ? '#f3f4f7' : 'transparent',
        color: active ? F.accent : F.inkSoft, fontFamily: F.font, fontSize: 14,
        fontWeight: active ? 600 : 500, textAlign: 'left', transition: 'background .1s',
      }}>
      {active && <span style={{ position: 'absolute', left: 0, top: 8, bottom: 8, width: 3, borderRadius: 2, background: F.accent }} />}
      <Icon name={item.icon} size={19} color={active ? F.accent : F.muted} />
      {item.label}
    </button>
  );
}

// ── Breadcrumb + page header ─────────────────────────────────────────────────
function Breadcrumb({ trail }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, color: F.muted, marginBottom: 6 }}>
      {trail.map((t, i) => (
        <React.Fragment key={i}>
          {i > 0 && <Icon name="chevronRight" size={13} color={F.faint} />}
          <span style={{ color: i === trail.length - 1 ? F.inkSoft : F.muted, fontWeight: i === trail.length - 1 ? 600 : 400 }}>{t}</span>
        </React.Fragment>
      ))}
    </div>
  );
}

function PageHeader({ title, subtitle, trail, actions }) {
  return (
    <div style={{ marginBottom: 18 }}>
      {trail && <Breadcrumb trail={trail} />}
      <div style={{ display: 'flex', alignItems: 'flex-end', gap: 16 }}>
        <div style={{ flex: 1 }}>
          <h1 style={{ margin: 0, fontSize: 28, fontWeight: 600, color: F.ink, letterSpacing: '-.01em', lineHeight: 1.15 }}>{title}</h1>
          {subtitle && <p style={{ margin: '6px 0 0', fontSize: 14, color: F.muted, lineHeight: 1.5 }}>{subtitle}</p>}
        </div>
        {actions && <div style={{ display: 'flex', gap: 8, flex: 'none' }}>{actions}</div>}
      </div>
    </div>
  );
}

// ── Command bar (toolbar row above content) ──────────────────────────────────
function CommandBar({ children, right }) {
  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 6, padding: '8px 0', marginBottom: 14,
      borderBottom: `1px solid ${F.divider}`,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, flex: 1, flexWrap: 'wrap' }}>{children}</div>
      {right && <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>{right}</div>}
    </div>
  );
}
function CommandDivider() {
  return <span style={{ width: 1, height: 20, background: F.divider, margin: '0 4px' }} />;
}

// ── Settings panel (right slide-over with pivot tabs) ────────────────────────
function SettingsPanel({ onClose }) {
  const [tab, setTab] = _uSc('Config');
  const tabs = ['Config', 'Customer', 'Workloads', 'Discovery'];
  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 60 }}>
      <div onClick={onClose} style={{ position: 'absolute', inset: 0, background: 'rgba(20,26,42,.32)' }} />
      <div style={{
        position: 'absolute', top: 0, right: 0, bottom: 0, width: 460, background: '#fff',
        boxShadow: '-8px 0 32px rgba(16,24,40,.18)', display: 'flex', flexDirection: 'column',
        fontFamily: F.font, animation: 'flySlideIn .18s ease-out',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', padding: '18px 20px 14px', borderBottom: `1px solid ${F.divider}` }}>
          <h2 style={{ margin: 0, flex: 1, fontSize: 20, fontWeight: 600, color: F.ink }}>Settings</h2>
          <button onClick={onClose} style={{ border: 0, background: 'transparent', cursor: 'pointer', color: F.muted, padding: 6, borderRadius: F.radiusSm, display: 'flex' }}>
            <Icon name="x" size={20} />
          </button>
        </div>
        <div style={{ display: 'flex', gap: 4, padding: '0 20px', borderBottom: `1px solid ${F.divider}` }}>
          {tabs.map((t) => (
            <button key={t} onClick={() => setTab(t)} style={{
              border: 0, background: 'transparent', cursor: 'pointer', padding: '12px 6px', marginRight: 12,
              fontSize: 14, fontWeight: 600, fontFamily: F.font, color: tab === t ? F.accent : F.muted,
              borderBottom: tab === t ? `2px solid ${F.accent}` : '2px solid transparent', marginBottom: -1,
            }}>{t}</button>
          ))}
        </div>
        <div style={{ flex: 1, overflowY: 'auto', padding: 20, display: 'flex', flexDirection: 'column', gap: 16 }}>
          {tab === 'Config' && (<>
            <Field label="Fly API URL" value="https://contoso.avepointonline.com" onChange={() => {}} hint="Your AvePoint Online Services region endpoint." />
            <Field label="Client ID" value="8f3c2a91-7d4e-4b2a-9f10-a91b" onChange={() => {}} />
            <Field label="Client Secret" value="••••••••••••••••" password onChange={() => {}} hint="Expires 2028-05-31." />
          </>)}
          {tab === 'Customer' && (<>
            <Field label="Tenant prefix" value="CONTOSO" onChange={() => {}} />
            <Field label="SharePoint admin URL" value="https://contoso-admin.sharepoint.com" onChange={() => {}} />
          </>)}
          {tab === 'Workloads' && (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              {['SharePoint', 'Exchange', 'OneDrive', 'Teams', 'Teams Chat', 'Groups'].map((w) => (
                <Field key={w} label={w} value={w} onChange={() => {}} />
              ))}
            </div>
          )}
          {tab === 'Discovery' && (
            <Field label="Discovery output path" value="C:\\Migration\\Discovery" onChange={() => {}} />
          )}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, padding: '14px 20px', borderTop: `1px solid ${F.divider}` }}>
          <Button variant="default" onClick={onClose}>Cancel</Button>
          <Button variant="primary" onClick={onClose}>Save changes</Button>
        </div>
      </div>
    </div>
  );
}

Object.assign(window, { TopBar, NavRail, NavItem, Breadcrumb, PageHeader, CommandBar, CommandDivider, SettingsPanel });
