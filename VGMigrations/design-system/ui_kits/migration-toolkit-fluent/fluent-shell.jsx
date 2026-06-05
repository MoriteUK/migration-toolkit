/* global React */
// ============================================================================
// Fly Migration Toolkit — Fluent refresh · tokens, icons & primitives
// Modern Microsoft-365-admin-center styling, kept to what WinForms can render
// (flat panels, hairline borders, Segoe UI, Segoe Fluent Icons glyphs).
// On the web we add a faint shadow + 4px radius for polish — see README caveats.
// ============================================================================
const { useState: _uS } = React;

const F = {
  // brand (unchanged from the product)
  accent: '#0064b4', accentHover: '#004e98', accentPressed: '#003a72',
  accentTint: '#eaf1fb', accentTintBorder: '#cfe0f5',
  // fluent neutrals
  shell: '#f3f4f7', nav: '#ffffff', content: '#f7f8fa', card: '#ffffff',
  ink: '#1b1b1f', inkSoft: '#42434a', muted: '#6a6c78', faint: '#8a8c98',
  border: '#e5e7ee', borderSoft: '#eef0f5', divider: '#e9ebf1',
  // status
  green: '#0f7a36', greenTint: '#e6f4ea', greenBorder: '#bfe3cb',
  amber: '#9a6700', amberTint: '#fbf3e0', amberBorder: '#efdcab',
  red: '#b42318', redTint: '#fdecec', redBorder: '#f3c5c0',
  blue: '#0064b4', blueTint: '#eaf1fb', blueBorder: '#cfe0f5',
  grey: '#6a6c78', greyTint: '#eef0f3', greyBorder: '#dadce4',
  // console (kept dark — modern editor convention)
  logBg: '#1e1f2b', logBorder: '#2c2e3e',
  font: '"Segoe UI Variable","Segoe UI",-apple-system,system-ui,sans-serif',
  mono: '"Cascadia Code","Consolas","SF Mono",Menlo,monospace',
  // web-only polish (drop for 1:1 WinForms)
  radius: 6, radiusSm: 4,
  shadowCard: '0 1px 2px rgba(16,24,40,.06), 0 1px 3px rgba(16,24,40,.05)',
  shadowRaise: '0 4px 12px rgba(16,24,40,.10)',
};

// ── Icons ───────────────────────────────────────────────────────────────────
// Line glyphs in Segoe-Fluent spirit. In the real WinForms build these map to
// Segoe Fluent Icons codepoints; rendered here as SVG so they show everywhere.
const ICON_PATHS = {
  home: 'M3 11.5 12 4l9 7.5M5 10v10h5v-6h4v6h5V10',
  send: 'M4 12 20 4l-5 16-3.5-6.5L4 12z',
  search: 'M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zm6.5-1.5L21 21',
  settings: 'M12 15.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7zM12 2.5v2.5M12 19v2.5M4.6 6.1l1.8 1.8M17.6 16.1l1.8 1.8M2.5 12H5M19 12h2.5M4.6 17.9l1.8-1.8M17.6 7.9l1.8-1.8',
  grid: 'M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z',
  server: 'M4 5h16v5H4zM4 14h16v5H4zM7.5 7.5h.01M7.5 16.5h.01',
  mail: 'M3 6h18v12H3zM3 7l9 6 9-6',
  cloud: 'M7 18a4 4 0 0 1 0-8 5 5 0 0 1 9.6-1.3A3.5 3.5 0 0 1 17.5 18H7z',
  users: 'M16 19v-1.5a3.5 3.5 0 0 0-3.5-3.5h-5A3.5 3.5 0 0 0 4 17.5V19M10 11a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7zM20 19v-1.5a3.5 3.5 0 0 0-2.6-3.4M15 4.2a3.5 3.5 0 0 1 0 6.6',
  teams: 'M4 8h9v9a4.5 4.5 0 0 1-9 0V8zM13.5 9h6.5v6a3 3 0 0 1-6 0M16.5 7a2 2 0 1 0 0-4 2 2 0 0 0 0 4z',
  folder: 'M3 6h6l2 2.5h10V19H3V6z',
  trash: 'M5 7h14M9 7V4.5h6V7M6.5 7l.8 12.5h9.4L17.5 7M10 10.5v6M14 10.5v6',
  sync: 'M20 11a8 8 0 0 0-14.3-4.2M4 5v3.5h3.5M4 13a8 8 0 0 0 14.3 4.2M20 19v-3.5h-3.5',
  download: 'M12 4v11M7.5 10.5 12 15l4.5-4.5M5 19h14',
  chevronRight: 'M9 5l7 7-7 7',
  chevronDown: 'M5 9l7 7 7-7',
  check: 'M5 12.5 10 17.5 19.5 7',
  alert: 'M12 4 2.5 20h19L12 4zM12 10v4.5M12 17.5h.01',
  x: 'M6 6l12 12M18 6 6 18',
  shield: 'M12 3 5 6v6c0 4.5 3 7.5 7 9 4-1.5 7-4.5 7-9V6l-7-3z',
  plug: 'M9 3v5M15 3v5M7 8h10v3a5 5 0 0 1-10 0V8zM12 16v5',
  chart: 'M4 20V10M10 20V4M16 20v-7M22 20H2',
  user: 'M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM5 20a7 7 0 0 1 14 0',
  waffle: 'M5 5h3v3H5zM10.5 5h3v3h-3zM16 5h3v3h-3zM5 10.5h3v3H5zM10.5 10.5h3v3h-3zM16 10.5h3v3h-3zM5 16h3v3H5zM10.5 16h3v3h-3zM16 16h3v3h-3z',
  key: 'M14 9a4 4 0 1 0-3.5 3.97L11 14h2v2h2v2h3v-3l-3.5-3.5A4 4 0 0 0 14 9z',
  clock: 'M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18zM12 7.5V12l3 2',
  filter: 'M4 5h16l-6 7v6l-4 2v-8L4 5z',
  more: 'M6 12h.01M12 12h.01M18 12h.01',
  doc: 'M6 3h8l4 4v14H6zM14 3v4h4',
};
function Icon({ name, size = 20, color = 'currentColor', filled = false, style }) {
  const d = ICON_PATHS[name] || ICON_PATHS.grid;
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
      stroke={color} strokeWidth={1.7} strokeLinecap="round" strokeLinejoin="round"
      style={{ flex: 'none', display: 'block', ...style }} aria-hidden="true">
      <path d={d} />
    </svg>
  );
}

// ── Button ───────────────────────────────────────────────────────────────────
function Button({ children, variant = 'default', icon, onClick, disabled, style }) {
  const [h, setH] = _uS(false);
  const [p, setP] = _uS(false);
  const map = {
    primary: { bg: F.accent, hov: F.accentHover, pre: F.accentPressed, fg: '#fff', bd: 'transparent' },
    default: { bg: '#fff', hov: '#f5f6f9', pre: '#eceef3', fg: F.ink, bd: F.border },
    subtle: { bg: 'transparent', hov: '#eef0f5', pre: '#e4e7ee', fg: F.inkSoft, bd: 'transparent' },
    danger: { bg: '#fff', hov: F.redTint, pre: '#f8dcd8', fg: F.red, bd: F.redBorder },
  };
  const c = map[variant] || map.default;
  return (
    <button onClick={onClick} disabled={disabled}
      onMouseEnter={() => setH(true)} onMouseLeave={() => { setH(false); setP(false); }}
      onMouseDown={() => setP(true)} onMouseUp={() => setP(false)}
      style={{
        height: 32, padding: icon && !children ? 0 : '0 14px', width: icon && !children ? 32 : 'auto',
        minWidth: icon && !children ? 32 : undefined,
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 8,
        border: `1px solid ${c.bd}`, borderRadius: F.radiusSm, cursor: disabled ? 'default' : 'pointer',
        background: disabled ? '#f3f4f7' : (p ? c.pre : h ? c.hov : c.bg),
        color: disabled ? F.faint : c.fg, fontWeight: variant === 'primary' ? 600 : 500,
        fontSize: 14, fontFamily: F.font, transition: 'background .1s, border-color .1s',
        whiteSpace: 'nowrap', ...style,
      }}>
      {icon && <Icon name={icon} size={18} />}
      {children}
    </button>
  );
}

// ── Card ─────────────────────────────────────────────────────────────────────
function Card({ children, pad = 20, style, hover, onClick }) {
  const [h, setH] = _uS(false);
  return (
    <div onClick={onClick}
      onMouseEnter={() => hover && setH(true)} onMouseLeave={() => hover && setH(false)}
      style={{
        background: F.card, border: `1px solid ${F.border}`, borderRadius: F.radius,
        boxShadow: h ? F.shadowRaise : F.shadowCard, padding: pad,
        cursor: onClick ? 'pointer' : 'default', transition: 'box-shadow .12s, border-color .12s',
        borderColor: h ? F.accentTintBorder : F.border, ...style,
      }}>{children}</div>
  );
}

// ── Status pill ────────────────────────────────────────────────────────────
function Pill({ tone = 'grey', children, dot = true }) {
  const map = {
    green: [F.green, F.greenTint, F.greenBorder], amber: [F.amber, F.amberTint, F.amberBorder],
    red: [F.red, F.redTint, F.redBorder], blue: [F.blue, F.blueTint, F.blueBorder],
    grey: [F.grey, F.greyTint, F.greyBorder],
  };
  const [fg, bg, bd] = map[tone] || map.grey;
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 6, height: 22, padding: '0 9px',
      borderRadius: 11, background: bg, border: `1px solid ${bd}`, color: fg,
      fontSize: 12, fontWeight: 600, fontFamily: F.font, whiteSpace: 'nowrap',
    }}>
      {dot && <span style={{ width: 6, height: 6, borderRadius: '50%', background: fg }} />}
      {children}
    </span>
  );
}

// ── Field / Select / Toggle / SearchBox ─────────────────────────────────────
function Field({ label, value, onChange, password, placeholder, hint, style }) {
  const [f, setF] = _uS(false);
  return (
    <label style={{ display: 'block', ...style }}>
      {label && <span style={{ display: 'block', fontSize: 13, fontWeight: 600, color: F.inkSoft, marginBottom: 5 }}>{label}</span>}
      <input type={password ? 'password' : 'text'} value={value} placeholder={placeholder}
        onChange={(e) => onChange && onChange(e.target.value)}
        onFocus={() => setF(true)} onBlur={() => setF(false)}
        style={{
          width: '100%', height: 34, boxSizing: 'border-box', padding: '0 11px',
          border: `1px solid ${f ? F.accent : F.border}`, borderRadius: F.radiusSm,
          boxShadow: f ? `0 0 0 2px ${F.accentTint}` : 'none',
          fontFamily: F.font, fontSize: 14, color: F.ink, background: '#fff', outline: 'none',
          transition: 'border-color .1s, box-shadow .1s',
        }} />
      {hint && <span style={{ display: 'block', fontSize: 12, color: F.muted, marginTop: 4 }}>{hint}</span>}
    </label>
  );
}

function Select({ label, value, options, onChange, width, style }) {
  return (
    <label style={{ display: 'inline-flex', flexDirection: 'column', gap: 5, ...style }}>
      {label && <span style={{ fontSize: 13, fontWeight: 600, color: F.inkSoft }}>{label}</span>}
      <div style={{ position: 'relative', width: width || 'auto' }}>
        <select value={value} onChange={(e) => onChange && onChange(e.target.value)}
          style={{
            height: 34, width: width || 'auto', padding: '0 30px 0 11px', border: `1px solid ${F.border}`,
            borderRadius: F.radiusSm, fontFamily: F.font, fontSize: 14, color: F.ink, background: '#fff',
            appearance: 'none', cursor: 'pointer', outline: 'none',
          }}>
          {options.map((o) => <option key={o} value={o}>{o}</option>)}
        </select>
        <span style={{ position: 'absolute', right: 9, top: 8, pointerEvents: 'none' }}>
          <Icon name="chevronDown" size={16} color={F.muted} />
        </span>
      </div>
    </label>
  );
}

function Toggle({ checked, onChange, label }) {
  return (
    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 10, cursor: 'pointer', fontSize: 14, color: F.ink, fontFamily: F.font }}>
      <span onClick={() => onChange && onChange(!checked)} style={{
        width: 38, height: 20, borderRadius: 10, background: checked ? F.accent : '#c4c7d0',
        position: 'relative', transition: 'background .15s', flex: 'none',
      }}>
        <span style={{
          position: 'absolute', top: 2, left: checked ? 20 : 2, width: 16, height: 16, borderRadius: '50%',
          background: '#fff', transition: 'left .15s', boxShadow: '0 1px 2px rgba(0,0,0,.3)',
        }} />
      </span>
      {label}
    </label>
  );
}

function SearchBox({ placeholder = 'Search', width = 320 }) {
  const [f, setF] = _uS(false);
  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 8, height: 32, width, padding: '0 11px',
      border: `1px solid ${f ? F.accent : F.border}`, borderRadius: F.radiusSm, background: '#fff',
      boxShadow: f ? `0 0 0 2px ${F.accentTint}` : 'none',
    }}>
      <Icon name="search" size={17} color={F.muted} />
      <input placeholder={placeholder} onFocus={() => setF(true)} onBlur={() => setF(false)}
        style={{ border: 0, outline: 'none', flex: 1, fontFamily: F.font, fontSize: 14, color: F.ink, background: 'transparent' }} />
    </div>
  );
}

Object.assign(window, { F, Icon, Button, Card, Pill, Field, Select, Toggle, SearchBox });
