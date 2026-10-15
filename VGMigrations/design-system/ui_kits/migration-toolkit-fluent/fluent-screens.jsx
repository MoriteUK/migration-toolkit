/* global React, F, Icon, Button, Card, Pill, Field, Select, Toggle, PageHeader, CommandBar, CommandDivider */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · screens
// Home dashboard · AvePoint Fly menu · Project Monitor (light list) · App Reg.
// ============================================================================
const { useState: _uSs, useEffect: _uEs, useRef: _uRs } = React;

// ── HOME DASHBOARD ───────────────────────────────────────────────────────────
const WORKLOADS = [
  { name: 'SharePoint', icon: 'cloud', total: 412, done: 374, tone: 'green', status: 'On track' },
  { name: 'Exchange', icon: 'mail', total: 206, done: 187, tone: 'amber', status: '3 warnings' },
  { name: 'OneDrive', icon: 'folder', total: 318, done: 313, tone: 'green', status: 'On track' },
  { name: 'Teams', icon: 'teams', total: 58, done: 44, tone: 'red', status: '2 failed' },
];
const AREAS = [
  { id: 'discovery', icon: 'search', title: 'Discovery', desc: 'Assess mailboxes, sites, OneDrive and groups before you migrate.' },
  { id: 'fly', icon: 'send', title: 'AvePoint Fly', desc: 'Connections, mappings and the numbered migration workflow.' },
  { id: 'monitor', icon: 'chart', title: 'Project Monitor', desc: 'Live progress across every workload with auto-refresh.' },
  { id: 'domain', icon: 'trash', title: 'Domain Removal', desc: 'Update UPNs, sync AD and remove decommissioned domains.' },
];

function Home({ onNav, onGear }) {
  const totalItems = WORKLOADS.reduce((s, w) => s + w.total, 0);
  const totalDone = WORKLOADS.reduce((s, w) => s + w.done, 0);
  const pct = Math.round((totalDone / totalItems) * 100);
  return (
    <Page>
      <PageHeader
        title="Migration overview"
        subtitle="Contoso → Fabrikam · tenant migration in progress"
        trail={['Home']}
        actions={<>
          <Button variant="default" icon="sync">Refresh</Button>
          <Button variant="primary" icon="send" onClick={() => onNav('fly')}>Open AvePoint Fly</Button>
        </>}
      />

      {/* summary strip */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 16 }}>
        <Stat label="Items in scope" value={totalItems.toLocaleString()} sub="across 6 workloads" />
        <Stat label="Completed" value={totalDone.toLocaleString()} sub={pct + '% of scope'} tone="green" />
        <Stat label="In progress" value="57" sub="actively migrating" tone="blue" />
        <Stat label="Needs attention" value="5" sub="2 failed · 3 warnings" tone="red" />
      </div>

      {/* progress by workload */}
      <Card pad={0} style={{ marginBottom: 24 }}>
        <div style={{ padding: '16px 20px', borderBottom: `1px solid ${F.divider}`, display: 'flex', alignItems: 'center' }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: F.ink, flex: 1 }}>Progress by workload</h2>
          <Button variant="subtle" icon="chart" onClick={() => onNav('monitor')}>Open monitor</Button>
        </div>
        <div>
          {WORKLOADS.map((w, i) => {
            const p = Math.round((w.done / w.total) * 100);
            return (
              <div key={w.name} style={{
                display: 'flex', alignItems: 'center', gap: 16, padding: '14px 20px',
                borderTop: i ? `1px solid ${F.divider}` : 'none',
              }}>
                <span style={{ width: 36, height: 36, borderRadius: F.radiusSm, background: F.accentTint, display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
                  <Icon name={w.icon} size={20} color={F.accent} />
                </span>
                <div style={{ width: 130, flex: 'none', fontSize: 14, fontWeight: 600, color: F.ink }}>{w.name}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ height: 8, borderRadius: 4, background: '#eceef3', overflow: 'hidden' }}>
                    <div style={{ width: p + '%', height: '100%', borderRadius: 4, background: w.tone === 'red' ? F.red : w.tone === 'amber' ? F.amber : F.green }} />
                  </div>
                </div>
                <div style={{ width: 96, flex: 'none', textAlign: 'right', fontSize: 13, color: F.muted, fontVariantNumeric: 'tabular-nums' }}>{w.done}/{w.total}</div>
                <div style={{ width: 110, flex: 'none', display: 'flex', justifyContent: 'flex-end' }}>
                  <Pill tone={w.tone}>{w.status}</Pill>
                </div>
              </div>
            );
          })}
        </div>
      </Card>

      {/* quick actions */}
      <h2 style={{ margin: '0 0 14px', fontSize: 16, fontWeight: 600, color: F.ink }}>Tools</h2>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 16 }}>
        {AREAS.map((a) => (
          <Card key={a.id} hover onClick={() => onNav(a.id)} style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
            <span style={{ width: 44, height: 44, borderRadius: F.radiusSm, background: F.accentTint, display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
              <Icon name={a.icon} size={24} color={F.accent} />
            </span>
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <h3 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: F.ink, flex: 1 }}>{a.title}</h3>
                <Icon name="chevronRight" size={18} color={F.faint} />
              </div>
              <p style={{ margin: '5px 0 0', fontSize: 13.5, color: F.muted, lineHeight: 1.5 }}>{a.desc}</p>
            </div>
          </Card>
        ))}
      </div>
    </Page>
  );
}

function Stat({ label, value, sub, tone }) {
  const accent = { green: F.green, blue: F.blue, red: F.red }[tone] || F.ink;
  return (
    <Card pad={16}>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: F.muted, textTransform: 'uppercase', letterSpacing: '.03em' }}>{label}</div>
      <div style={{ fontSize: 30, fontWeight: 600, color: accent, lineHeight: 1.1, margin: '8px 0 2px', fontVariantNumeric: 'tabular-nums' }}>{value}</div>
      <div style={{ fontSize: 12.5, color: F.faint }}>{sub}</div>
    </Card>
  );
}

// ── AVEPOINT FLY MENU (numbered workflow as a stepped list) ──────────────────
const FLY_STEPS = [
  { n: 1, id: 'appreg', icon: 'key', title: 'Create App Registration', desc: 'Register the Entra ID app and grant required API permissions.', done: true },
  { n: 2, id: 'aos', icon: 'server', title: 'Setup AOS Tenant & App', desc: 'Configure the AvePoint Online Services tenant and application.', done: true },
  { n: 3, id: 'conn', icon: 'plug', title: 'Connections & Migration Mappings', desc: 'Manage connections, source/destination accounts and job mappings.', done: false, active: true },
  { n: 4, id: 'reports', icon: 'doc', title: 'View Migration Reports', desc: 'Review per-user results and export status reports.', done: false },
  { n: 5, id: 'monitor', icon: 'chart', title: 'Monitor Projects', desc: 'Live project monitoring and migration progress tracking.', done: false },
];

function FlyMenu({ onNav, onOpen }) {
  return (
    <Page>
      <PageHeader title="AvePoint Fly" subtitle="Run the migration workflow in order. Steps 1–2 are complete."
        trail={['Home', 'AvePoint Fly']}
        actions={<Button variant="primary" icon="chart" onClick={() => onNav('monitor')}>Open monitor</Button>} />
      <Card pad={0}>
        {FLY_STEPS.map((s, i) => {
          const clickable = s.id === 'appreg' || s.id === 'monitor';
          return (
            <div key={s.n} onClick={() => clickable && (s.id === 'monitor' ? onNav('monitor') : onOpen(s.id))}
              style={{
                display: 'flex', alignItems: 'center', gap: 16, padding: '16px 20px',
                borderTop: i ? `1px solid ${F.divider}` : 'none', cursor: clickable ? 'pointer' : 'default',
                background: s.active ? F.accentTint : 'transparent',
              }}>
              <StepBadge n={s.n} done={s.done} active={s.active} />
              <span style={{ width: 36, height: 36, borderRadius: F.radiusSm, background: '#f3f4f7', display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
                <Icon name={s.icon} size={20} color={F.inkSoft} />
              </span>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 15, fontWeight: 600, color: F.ink }}>{s.title}</div>
                <div style={{ fontSize: 13, color: F.muted, marginTop: 2 }}>{s.desc}</div>
              </div>
              {s.done && <Pill tone="green">Complete</Pill>}
              {s.active && <Pill tone="blue">In progress</Pill>}
              {clickable && <Icon name="chevronRight" size={18} color={F.faint} />}
            </div>
          );
        })}
      </Card>
    </Page>
  );
}

function StepBadge({ n, done, active }) {
  const bg = done ? F.green : active ? F.accent : '#fff';
  const fg = done || active ? '#fff' : F.muted;
  const bd = done ? F.green : active ? F.accent : F.border;
  return (
    <span style={{
      width: 28, height: 28, borderRadius: '50%', background: bg, color: fg, border: `1px solid ${bd}`,
      display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 13, fontWeight: 700, flex: 'none',
    }}>{done ? <Icon name="check" size={16} color="#fff" /> : n}</span>
  );
}

// shared page wrapper
function Page({ children }) {
  return <div style={{ maxWidth: 1080, margin: '0 auto', padding: '28px 32px 48px' }}>{children}</div>;
}

Object.assign(window, { Home, FlyMenu, Page, Stat });
