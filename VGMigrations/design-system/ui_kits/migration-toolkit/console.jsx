/* global React, FLY */
// ============================================================================
// Fly Migration Toolkit — UI Kit · console & data grid
// Recreates the Write-Log RichTextBox and the Project Monitor DataGridView.
// ============================================================================

// ── Coloured log console (Write-Log in lib.ps1) ────────────────────────────
const LOG_LEVEL = {
  INFO: '#789bdc', OK: '#41c36e', WARN: '#dca52d', ERROR: '#e15050',
};
function LogConsole({ lines, height = 200 }) {
  const ref = React.useRef(null);
  React.useEffect(() => { if (ref.current) ref.current.scrollTop = ref.current.scrollHeight; }, [lines]);
  return (
    <div ref={ref} style={{
      background: FLY.logBg, borderRadius: 3, padding: '12px 14px', height, overflowY: 'auto',
      fontFamily: FLY.mono, fontSize: 12.5, lineHeight: 1.55,
    }}>
      {lines.map((l, i) => (
        <div key={i}>
          <span style={{ color: '#505f78' }}>{l.ts}</span>{' '}
          <span style={{ color: LOG_LEVEL[l.level] || LOG_LEVEL.INFO, fontWeight: 600 }}>[{l.level}]</span>{' '}
          <span style={{ color: '#cdd4e6' }}>{l.msg}</span>
        </div>
      ))}
    </div>
  );
}

// ── Project Monitor data grid (dark themed DataGridView) ───────────────────
const MON_COLS = ['Project', 'Total', 'Not Started', 'In Progress', 'Complete', 'Failed', 'Warnings', 'Last Refresh'];

function MonitorGrid({ rows, selected, onSelect }) {
  return (
    <div style={{ background: FLY.logBg, height: '100%', overflow: 'auto' }}>
      <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12.5, fontFamily: FLY.font }}>
        <thead>
          <tr>
            {MON_COLS.map((c, i) => (
              <th key={c} style={{
                background: FLY.accent, color: '#fff', fontWeight: 700, textAlign: i === 0 ? 'left' : 'center',
                padding: '8px 10px', borderRight: '1px solid #2d374b', whiteSpace: 'nowrap', position: 'sticky', top: 0,
              }}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => {
            const isSel = selected === ri;
            let bg = 'transparent', fg = '#bed2ff';
            if (r.failed > 0 || r.err) { bg = '#461919'; fg = '#f07d7d'; }
            else if (r.warnings > 0) { bg = '#41300c'; fg = '#ebc350'; }
            if (isSel) { bg = FLY.accentHover; fg = '#fff'; }
            const cells = r.err
              ? [r.project, 'ERR', r.err, '', '', '', '', r.refresh]
              : [r.project, r.total, r.notStarted, r.inProgress, r.complete, r.failed, r.warnings, r.refresh];
            return (
              <tr key={ri} onClick={() => onSelect && onSelect(ri)} style={{ cursor: 'pointer' }}>
                {cells.map((val, ci) => (
                  <td key={ci} style={{
                    background: bg, color: fg, padding: '6px 10px', borderTop: '1px solid #2d374b',
                    borderRight: '1px solid #2d374b', textAlign: ci === 0 || (r.err && ci === 2) ? 'left' : 'center',
                    fontFamily: ci === 0 ? FLY.font : FLY.mono, whiteSpace: 'nowrap',
                  }}>{val}</td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

Object.assign(window, { LogConsole, MonitorGrid });
