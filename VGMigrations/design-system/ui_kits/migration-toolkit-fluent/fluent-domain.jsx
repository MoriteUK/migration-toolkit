/* global React, F, Icon, Button, Card, Pill, Field, Select, Toggle, PageHeader, Page */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · Domain Removal
// The most complex / destructive surface: a guided 3-step workflow
// (Update on-prem UPNs → Run AD sync → Remove domain), standalone tools,
// a live activity log, and a type-to-confirm dialog for the destructive step.
// ============================================================================
const { useState: _uSd, useEffect: _uEd, useRef: _uRd } = React;

const DLVL = { INFO: '#7aa2e8', OK: '#4cc97a', WARN: '#e3b341', ERROR: '#e86a6a' };

// ── Fluent MessageBar (full-width tinted notice, not a left-border card) ─────
function MessageBar({ tone = 'warn', icon = 'alert', children, style }) {
  const map = {
    warn: [F.amberTint, F.amberBorder, F.amber],
    info: [F.blueTint, F.blueBorder, F.blue],
    error: [F.redTint, F.redBorder, F.red],
    success: [F.greenTint, F.greenBorder, F.green],
  };
  const [bg, bd, fg] = map[tone] || map.warn;
  return (
    <div style={{
      display: 'flex', alignItems: 'flex-start', gap: 11, padding: '12px 16px',
      background: bg, border: `1px solid ${bd}`, borderRadius: F.radiusSm,
      fontSize: 13.5, color: F.inkSoft, lineHeight: 1.5, ...style,
    }}>
      <Icon name={icon} size={18} color={fg} style={{ marginTop: 1 }} />
      <div style={{ flex: 1 }}>{children}</div>
    </div>
  );
}

// ── Type-to-confirm destructive dialog ───────────────────────────────────────
function ConfirmRemoveDialog({ domain, onClose, onConfirm }) {
  const [val, setVal] = _uSd('');
  const match = val.trim().toLowerCase() === domain.toLowerCase();
  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 70 }}>
      <div onClick={onClose} style={{ position: 'absolute', inset: 0, background: 'rgba(20,26,42,.4)' }} />
      <div style={{
        position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%,-50%)',
        width: 460, background: '#fff', borderRadius: F.radius, boxShadow: '0 24px 60px rgba(10,15,35,.4)',
        fontFamily: F.font, overflow: 'hidden',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '18px 22px 14px' }}>
          <span style={{ width: 36, height: 36, borderRadius: '50%', background: F.redTint, display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
            <Icon name="alert" size={20} color={F.red} />
          </span>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 600, color: F.ink }}>Remove domain?</h2>
        </div>
        <div style={{ padding: '0 22px 18px' }}>
          <p style={{ margin: '0 0 14px', fontSize: 14, color: F.inkSoft, lineHeight: 1.55 }}>
            This permanently removes <b>{domain}</b> as a verified domain and detaches all associated
            Microsoft&nbsp;365 objects. <b style={{ color: F.red }}>This cannot be undone.</b>
          </p>
          <Field label={`Type "${domain}" to confirm`} value={val} onChange={setVal} placeholder={domain} />
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, padding: '14px 22px', borderTop: `1px solid ${F.divider}`, background: '#fbfbfd' }}>
          <Button variant="default" onClick={onClose}>Cancel</Button>
          <Button variant="danger" disabled={!match} onClick={() => match && onConfirm()}>Remove domain</Button>
        </div>
      </div>
    </div>
  );
}

// ── Guided step row ───────────────────────────────────────────────────────────
function StepRow({ step, index, state, onRun, isLast }) {
  const status = state.status; // pending | active | running | done
  const locked = status === 'pending';
  const badgeBg = status === 'done' ? F.green : status === 'pending' ? '#fff' : (step.danger ? F.red : F.accent);
  const badgeFg = status === 'pending' ? F.muted : '#fff';
  const badgeBd = status === 'done' ? F.green : status === 'pending' ? F.border : (step.danger ? F.red : F.accent);
  return (
    <div style={{ display: 'flex', gap: 16, padding: '18px 4px' }}>
      {/* rail + badge */}
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', flex: 'none' }}>
        <span style={{
          width: 30, height: 30, borderRadius: '50%', background: badgeBg, color: badgeFg,
          border: `1px solid ${badgeBd}`, display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 13.5, fontWeight: 700,
        }}>{status === 'done' ? <Icon name="check" size={16} color="#fff" /> : index + 1}</span>
        {!isLast && <span style={{ width: 2, flex: 1, marginTop: 6, background: status === 'done' ? F.green : F.divider, minHeight: 28 }} />}
      </div>
      {/* content */}
      <div style={{ flex: 1, paddingBottom: isLast ? 0 : 4 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: locked ? F.muted : F.ink, flex: 1 }}>{step.title}</h3>
          {status === 'done' && <Pill tone="green">Done</Pill>}
          {status === 'running' && <Pill tone="blue">Running…</Pill>}
          {status === 'active' && <Pill tone="grey" dot={false}>Ready</Pill>}
          {status === 'pending' && <Pill tone="grey" dot={false}>Waiting</Pill>}
        </div>
        <p style={{ margin: '5px 0 0', fontSize: 13.5, color: F.muted, lineHeight: 1.5 }}>{step.desc}</p>
        {step.meta && <p style={{ margin: '6px 0 0', fontSize: 12.5, color: F.faint, fontFamily: F.mono }}>{step.meta}</p>}
        <div style={{ marginTop: 12 }}>
          <Button
            variant={step.danger ? 'danger' : status === 'done' ? 'default' : 'primary'}
            icon={step.icon}
            disabled={locked || status === 'running'}
            onClick={onRun}>
            {status === 'done' ? step.redo : status === 'running' ? 'Running…' : step.cta}
          </Button>
        </div>
      </div>
    </div>
  );
}

// ── Domain Removal screen ─────────────────────────────────────────────────────
const STEPS = [
  { key: 'upn', icon: 'users', title: '1. Update on-premise UPNs', cta: 'Update UPNs', redo: 'Re-run',
    desc: 'Change UPN, primary email and aliases in on-premise Active Directory away from the domain being removed.',
    meta: 'Target: VOL-ane-dc1 · scope: 412 users' },
  { key: 'sync', icon: 'sync', title: '2. Run Azure AD Connect sync', cta: 'Run delta sync', redo: 'Re-run sync',
    desc: 'Trigger a delta sync cycle so the updated on-premise attributes flow to Microsoft 365.',
    meta: 'Server: VOL-ane-aad1 · cycle: Delta' },
  { key: 'remove', icon: 'trash', title: '3. Remove domain', cta: 'Remove domain…', redo: 'Removed', danger: true,
    desc: 'Remove the verified domain and detach all associated Microsoft 365 objects. Destructive — runs only after steps 1–2.' },
];

function DomainRemoval() {
  const [domain, setDomain] = _uSd('legacy-contoso.com');
  const [states, setStates] = _uSd({ upn: 'active', sync: 'pending', remove: 'pending' });
  const [lines, setLines] = _uSd([]);
  const [confirm, setConfirm] = _uSd(false);
  const ref = _uRd(null);
  _uEd(() => { if (ref.current) ref.current.scrollTop = ref.current.scrollHeight; }, [lines]);

  const log = (level, msg, delay) => setTimeout(() => {
    const ts = new Date().toLocaleTimeString('en-GB', { hour12: false });
    setLines((p) => [...p, { ts, level, msg }]);
  }, delay);

  const complete = (key, nextKey) => {
    setStates((s) => ({ ...s, [key]: 'done', ...(nextKey ? { [nextKey]: 'active' } : {}) }));
  };

  const runStep = (key) => {
    if (key === 'remove') { setConfirm(true); return; }
    setStates((s) => ({ ...s, [key]: 'running' }));
    if (key === 'upn') {
      log('INFO', `Connecting to on-premise AD (VOL-ane-dc1)…`, 0);
      log('INFO', `Rewriting UPN suffix for 412 users from @${domain}…`, 500);
      log('OK', 'UPNs, primary SMTP and aliases updated for 412 users.', 1400);
      setTimeout(() => complete('upn', 'sync'), 1500);
    } else if (key === 'sync') {
      log('INFO', 'Triggering Azure AD Connect delta sync on VOL-ane-aad1…', 0);
      log('INFO', 'Sync cycle started. Waiting for export to Microsoft 365…', 600);
      log('OK', 'Delta sync complete. 412 objects exported.', 1600);
      setTimeout(() => complete('sync', 'remove'), 1700);
    }
  };

  const doRemove = () => {
    setConfirm(false);
    setStates((s) => ({ ...s, remove: 'running' }));
    log('WARN', `Removing verified domain ${domain}…`, 0);
    log('INFO', 'Detaching 412 associated Microsoft 365 objects…', 600);
    log('OK', `Domain ${domain} removed successfully.`, 1700);
    setTimeout(() => complete('remove', null), 1800);
  };

  const allDone = states.remove === 'done';

  return (
    <Page>
      <PageHeader title="Domain Removal"
        subtitle="Decommission a verified domain and clean up associated Microsoft 365 objects."
        trail={['Home', 'Domain Removal']}
        actions={<Select label="" value={domain} width={210}
          options={['legacy-contoso.com', 'old.fabrikam.com', 'northwind-legacy.com']} onChange={setDomain} />} />

      <MessageBar tone={allDone ? 'success' : 'warn'} icon={allDone ? 'check' : 'alert'} style={{ marginBottom: 20 }}>
        {allDone
          ? <>Domain <b>{domain}</b> has been removed. You can verify in the Microsoft 365 admin center.</>
          : <>Run these steps <b>in order</b>. The domain cannot be removed until on-premise UPNs are updated and synced, or Microsoft 365 will block the removal.</>}
      </MessageBar>

      <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 1fr', gap: 20, alignItems: 'start' }}>
        {/* guided workflow */}
        <Card pad={0}>
          <div style={{ padding: '16px 20px', borderBottom: `1px solid ${F.divider}` }}>
            <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: F.ink }}>Guided removal</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: F.muted }}>Three steps for <b style={{ color: F.inkSoft }}>{domain}</b></p>
          </div>
          <div style={{ padding: '4px 20px 16px' }}>
            {STEPS.map((s, i) => (
              <StepRow key={s.key} step={s} index={i} state={{ status: states[s.key] }}
                onRun={() => runStep(s.key)} isLast={i === STEPS.length - 1} />
            ))}
          </div>
        </Card>

        {/* activity log */}
        <div>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: F.muted, textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 8 }}>Activity log</div>
          <div ref={ref} style={{
            background: F.logBg, border: `1px solid ${F.logBorder}`, borderRadius: F.radius,
            padding: '14px 16px', height: 268, overflowY: 'auto', fontFamily: F.mono, fontSize: 12.5, lineHeight: 1.65,
          }}>
            {lines.length === 0
              ? <span style={{ color: '#5a6178' }}>Output appears here as you run each step.</span>
              : lines.map((l, i) => (
                <div key={i}>
                  <span style={{ color: '#5a6178' }}>{l.ts}</span>{' '}
                  <span style={{ color: DLVL[l.level], fontWeight: 600 }}>[{l.level}]</span>{' '}
                  <span style={{ color: '#cdd2e2' }}>{l.msg}</span>
                </div>
              ))}
          </div>
        </div>
      </div>

      {/* standalone tools */}
      <h2 style={{ margin: '28px 0 14px', fontSize: 16, fontWeight: 600, color: F.ink }}>Standalone tools</h2>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 16 }}>
        <ToolCard icon="users" title="Update On-Premise UPNs" desc="Update UPN, email and aliases in on-premise AD from a mapping file." cta="Open" />
        <ToolCard icon="sync" title="Run AD Sync" desc="Trigger an Azure AD Connect delta sync on VOL-ane-aad1." cta="Run sync" />
        <ToolCard icon="user" title="Hide from Global Address List" desc="Bulk-hide Exchange Online recipients from the GAL." cta="Open" />
        <ToolCard icon="trash" title="Remove Domain" desc="Remove a verified domain and its associated objects." cta="Remove…" danger onClick={() => setConfirm(true)} />
      </div>

      {confirm && <ConfirmRemoveDialog domain={domain} onClose={() => setConfirm(false)} onConfirm={doRemove} />}
    </Page>
  );
}

function ToolCard({ icon, title, desc, cta, danger, onClick }) {
  return (
    <Card style={{ display: 'flex', gap: 14, alignItems: 'flex-start' }}>
      <span style={{ width: 40, height: 40, borderRadius: F.radiusSm, background: danger ? F.redTint : F.accentTint, display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 'none' }}>
        <Icon name={icon} size={22} color={danger ? F.red : F.accent} />
      </span>
      <div style={{ flex: 1 }}>
        <h3 style={{ margin: 0, fontSize: 14.5, fontWeight: 600, color: F.ink }}>{title}</h3>
        <p style={{ margin: '5px 0 12px', fontSize: 13, color: F.muted, lineHeight: 1.5 }}>{desc}</p>
        <Button variant={danger ? 'danger' : 'default'} onClick={onClick}>{cta}</Button>
      </div>
    </Card>
  );
}

Object.assign(window, { DomainRemoval, MessageBar, ConfirmRemoveDialog });
