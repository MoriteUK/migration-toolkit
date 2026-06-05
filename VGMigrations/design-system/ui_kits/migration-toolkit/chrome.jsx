/* global React */
// ============================================================================
// Fly Migration Toolkit — UI Kit · chrome & primitives
// Faithful HTML/React recreation of the WinForms controls in lib.ps1.
// ============================================================================
const { useState } = React;

const FLY = {
  accent: '#0064b4', accentHover: '#004e98', accentTint: '#dce6f8',
  bg: '#f0f2f7', panel: '#ffffff', text: '#1c1c20', muted: '#646c78',
  border: '#d2d7e4', grey: '#afb6c3', logBg: '#1a1b26', footer: '#141826',
  footerAlt: '#1c2030', green: '#129b3c', amber: '#c38700', red: '#c31e1e',
  closeRed: '#c83737', bannerWarn: '#fff3cd', selTint: '#dce6f8',
  font: '"Segoe UI","Segoe UI Web (West European)",-apple-system,system-ui,sans-serif',
  mono: '"Cascadia Code","Consolas","SF Mono",Menlo,monospace',
};

// ── Window shell: blue banner + content + navy footer ──────────────────────
function Window({ title, width = 480, onGear, onClose, footerExtra, children, footerColor }) {
  return (
    <div style={{
      width, background: FLY.bg, borderRadius: 4, overflow: 'hidden',
      boxShadow: '0 18px 50px rgba(20,28,55,.28)', border: '1px solid #c2c9da',
      display: 'flex', flexDirection: 'column', fontFamily: FLY.font,
    }}>
      <Header title={title} onGear={onGear} />
      <div style={{ flex: 1, minHeight: 0 }}>{children}</div>
      <Footer onClose={onClose} extra={footerExtra} color={footerColor} />
    </div>
  );
}

function Header({ title, onGear, height = 56 }) {
  const logoH = height >= 56 ? 38 : 30;
  return (
    <div style={{
      height, background: FLY.accent, display: 'flex', alignItems: 'center',
      gap: 12, padding: '0 14px', flex: 'none',
    }}>
      <img src="../../assets/fly-logo-64.png" alt="" width={logoH} height={logoH}
        style={{ borderRadius: logoH * 0.18, display: 'block' }} />
      <span style={{
        color: '#fff', fontWeight: 600, fontSize: height >= 56 ? 19 : 17,
        flex: 1, letterSpacing: '.005em',
      }}>{title}</span>
      <GearButton onClick={onGear} />
    </div>
  );
}

function GearButton({ onClick }) {
  const [h, setH] = useState(false);
  return (
    <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      title="Settings" style={{
        width: 38, height: 38, border: 0, background: h ? FLY.accentHover : FLY.accent,
        cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
        padding: 0, borderRadius: 3,
      }}>
      <span style={{
        width: 28, height: 28, borderRadius: '50%', background: '#fff', color: FLY.accent,
        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 17,
        lineHeight: 1,
      }}>⚙</span>
    </button>
  );
}

function Footer({ onClose, extra, color }) {
  return (
    <div style={{
      height: 46, background: color || FLY.footer, display: 'flex', alignItems: 'center',
      padding: '0 10px', flex: 'none', gap: 10,
    }}>
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: 10 }}>{extra}</div>
      <FooterButton label="Close" onClick={onClose} />
    </div>
  );
}

function FooterButton({ label, onClick, color = FLY.closeRed }) {
  const [h, setH] = useState(false);
  return (
    <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        height: 30, padding: '0 18px', border: 0, borderRadius: 2, cursor: 'pointer',
        background: h ? '#b02f2f' : color, color: '#fff', fontWeight: 600, fontSize: 13,
        fontFamily: FLY.font,
      }}>{label}</button>
  );
}

// ── Navigation tile (oversized accent button + muted subtitle) ─────────────
function Tile({ label, subtitle, danger, onClick }) {
  const [h, setH] = useState(false);
  const base = danger ? FLY.red : FLY.accent;
  const hover = danger ? '#a81a1a' : FLY.accentHover;
  return (
    <div style={{ marginBottom: 26 }}>
      <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
        style={{
          display: 'block', width: '100%', height: 64, border: 0, borderRadius: 2,
          background: h ? hover : base, color: '#fff', fontWeight: 600, fontSize: 19,
          textAlign: 'left', padding: '0 20px', cursor: 'pointer', fontFamily: FLY.font,
          transition: 'background .12s',
        }}>{label}</button>
      {subtitle && <div style={{ fontSize: 12, color: FLY.muted, margin: '8px 0 0 4px' }}>{subtitle}</div>}
    </div>
  );
}

// ── Standard buttons ────────────────────────────────────────────────────────
function Button({ children, variant = 'primary', onClick, style }) {
  const [h, setH] = useState(false);
  const map = {
    primary: { bg: FLY.accent, hov: FLY.accentHover, fg: '#fff' },
    secondary: { bg: '#e1e4ee', hov: '#d4d8e6', fg: FLY.text },
    danger: { bg: FLY.red, hov: '#a81a1a', fg: '#fff' },
  };
  const c = map[variant];
  return (
    <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        height: 30, padding: '0 18px', border: 0, borderRadius: 2, cursor: 'pointer',
        background: h ? c.hov : c.bg, color: c.fg, fontWeight: 600, fontSize: 13,
        fontFamily: FLY.font, transition: 'background .12s', ...style,
      }}>{children}</button>
  );
}

// ── Accent-stripe card ──────────────────────────────────────────────────────
function Card({ title, children, style }) {
  return (
    <div style={{ background: FLY.border, borderRadius: 3, padding: 1, ...style }}>
      <div style={{ background: FLY.panel, borderRadius: 2, display: 'flex', overflow: 'hidden' }}>
        <div style={{ width: 4, background: FLY.accent, flex: 'none' }} />
        <div style={{ flex: 1, padding: '14px 18px' }}>
          {title && <div style={{
            fontSize: 10, fontWeight: 600, letterSpacing: '.06em', textTransform: 'uppercase',
            color: FLY.muted, marginBottom: 10,
          }}>{title}</div>}
          {children}
        </div>
      </div>
    </div>
  );
}

// ── Form controls ───────────────────────────────────────────────────────────
function Field({ label, value, onChange, password, placeholder }) {
  const [f, setF] = useState(false);
  return (
    <label style={{ display: 'block', marginBottom: 12 }}>
      <span style={{ display: 'block', fontSize: 13, color: FLY.text, marginBottom: 5 }}>{label}</span>
      <input type={password ? 'password' : 'text'} value={value} placeholder={placeholder}
        onChange={(e) => onChange && onChange(e.target.value)}
        onFocus={() => setF(true)} onBlur={() => setF(false)}
        style={{
          width: '100%', height: 28, boxSizing: 'border-box', padding: '0 9px',
          border: `1px solid ${f ? FLY.accent : FLY.border}`, borderRadius: 2,
          fontFamily: FLY.font, fontSize: 13, color: FLY.text, background: '#fff', outline: 'none',
        }} />
    </label>
  );
}

function Select({ label, value, options, onChange, width }) {
  return (
    <label style={{ display: 'inline-flex', flexDirection: 'column', gap: 5 }}>
      {label && <span style={{ fontSize: 13, color: FLY.text }}>{label}</span>}
      <select value={value} onChange={(e) => onChange && onChange(e.target.value)}
        style={{
          height: 28, width: width || 'auto', padding: '0 8px', border: `1px solid ${FLY.border}`,
          borderRadius: 2, fontFamily: FLY.font, fontSize: 13, color: FLY.text, background: '#fff',
        }}>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </label>
  );
}

function Checkbox({ label, checked, onChange }) {
  return (
    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: 13, color: FLY.text, cursor: 'pointer' }}>
      <input type="checkbox" checked={checked} onChange={(e) => onChange && onChange(e.target.checked)}
        style={{ accentColor: FLY.accent, width: 15, height: 15 }} />
      {label}
    </label>
  );
}

function StatusDot({ color = FLY.grey, size = 13 }) {
  return <span style={{ width: size, height: size, borderRadius: '50%', background: color, display: 'inline-block', flex: 'none' }} />;
}

Object.assign(window, {
  FLY, Window, Header, GearButton, Footer, FooterButton, Tile, Button, Card,
  Field, Select, Checkbox, StatusDot,
});
