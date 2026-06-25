// Migration Toolkit - Renderer Process (UI Logic)

// ---------------------------------------------------------------------------
// File browser helper — opens a native CSV file-open dialog and writes the
// chosen path into the given input element.  Works for any Browse button.
// ---------------------------------------------------------------------------
async function openFileBrowser(inputEl, opts) {
  try {
    const result = await window.electronAPI.showOpenDialog(Object.assign({
      properties: ['openFile'],
      filters: [
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    }, opts || {}));
    if (!result.canceled && result.filePaths.length > 0) {
      inputEl.value = result.filePaths[0];
      inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
  } catch (err) {
    console.error('File dialog error:', err);
  }
}

// Event delegation — handles every Browse button that sits inside an
// .input-with-button container (all views, current and future).
// Buttons may carry data-extensions="xlsx" (or other comma-separated exts)
// to override the default CSV filter.
document.addEventListener('click', (e) => {
  const btn = e.target.closest('.input-with-button button');
  if (!btn) return;
  const input = btn.closest('.input-with-button')?.querySelector('input');
  if (!input) return;
  const exts = btn.dataset.extensions;
  if (btn.dataset.folder) {
    window.electronAPI.showOpenDialog({ properties: ['openDirectory'] }).then(result => {
      if (!result.canceled && result.filePaths.length > 0) {
        input.value = result.filePaths[0];
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }).catch(err => console.error('Folder dialog error:', err));
  } else if (exts) {
    const extList = exts.split(',').map(s => s.trim());
    openFileBrowser(input, {
      filters: [
        { name: extList.map(x => x.toUpperCase()).join('/') + ' Files', extensions: extList },
        { name: 'All Files', extensions: ['*'] }
      ]
    });
  } else {
    openFileBrowser(input);
  }
});

console.log('=== RENDERER.JS LOADING ===');
console.log('Document ready state:', document.readyState);

// Wait for DOM to load
document.addEventListener('DOMContentLoaded', async () => {
  console.log('=== DOMCONTENTLOADED EVENT FIRED ===');
  console.log('Migration Toolkit loaded');
  console.log('electronAPI available:', !!window.electronAPI);


  if (!window.electronAPI) {
    console.error('FATAL: electronAPI not available! Preload script may not be working.');
    document.body.innerHTML = '<div style="padding:40px;text-align:center;color:red;"><h1>Error</h1><p>electronAPI not loaded. Check console for details.</p></div>';
    return;
  }

  // Get version and check for updates on startup
  await loadVersion();
  await checkForUpdates();

  // Tile click handlers
  const tiles = document.querySelectorAll('.tile');
  console.log('Attaching click handlers to', tiles.length, 'tiles');

  tiles.forEach((tile, index) => {
    const scriptName = tile.getAttribute('data-script');
    console.log(`  Tile ${index}: ${scriptName}`);

    tile.addEventListener('click', async (event) => {
      console.log('!!! TILE CLICKED !!!', scriptName);
      if (scriptName) {
        await launchScript(scriptName, tile);
      }
    });

    // Verify listener was attached
    console.log(`  ✓ Click handler attached to tile ${index}`);
  });

  console.log('All click handlers attached');

  // Sidebar navigation - regular items
  const sidebarItems = document.querySelectorAll('.sidebar-item:not(.sidebar-item-expandable)');
  sidebarItems.forEach((item) => {
    item.addEventListener('click', (event) => {
      const view = item.getAttribute('data-view');

      // Update active state
      sidebarItems.forEach(i => i.classList.remove('active'));
      document.querySelectorAll('.sidebar-subitem').forEach(i => i.classList.remove('active'));
      item.classList.add('active');

      // Switch views
      if (view) {
        switchView(view);
      }
    });
  });

  // Sidebar navigation - expandable items
  const expandableItems = document.querySelectorAll('.sidebar-item-expandable');
  expandableItems.forEach((item) => {
    item.addEventListener('click', (event) => {
      const menuId = item.getAttribute('data-menu');
      const submenu = document.getElementById(`submenu-${menuId}`);

      // Toggle this menu
      const isExpanded = item.classList.contains('expanded');

      // Close all other menus
      expandableItems.forEach(i => {
        if (i !== item) {
          i.classList.remove('expanded');
        }
      });
      document.querySelectorAll('.sidebar-submenu').forEach(s => {
        if (s !== submenu) {
          s.classList.remove('expanded');
        }
      });

      // Toggle current menu
      if (isExpanded) {
        item.classList.remove('expanded');
        submenu.classList.remove('expanded');
      } else {
        item.classList.add('expanded');
        submenu.classList.add('expanded');
      }
    });
  });

  // Sidebar navigation - sub-items
  const sidebarSubitems = document.querySelectorAll('.sidebar-subitem');
  sidebarSubitems.forEach((subitem) => {
    subitem.addEventListener('click', (event) => {
      event.stopPropagation();
      const view = subitem.getAttribute('data-view');

      // Update active state
      sidebarItems.forEach(i => i.classList.remove('active'));
      sidebarSubitems.forEach(i => i.classList.remove('active'));
      subitem.classList.add('active');

      // Switch views
      if (view) {
        switchView(view);
      }
    });
  });

  // Launch buttons in views - any button with data-script attribute
  const launchButtons = document.querySelectorAll('button[data-script]');
  launchButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      const scriptName = button.getAttribute('data-script');
      if (scriptName) {
        await launchScript(scriptName, button);
      }
    });
  });

  // Discovery form - domain mode toggle
  const domainModeRadios = document.querySelectorAll('input[name="domainMode"]');
  domainModeRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
      const singleSection = document.getElementById('singleDomainSection');
      const multipleSection = document.getElementById('multipleDomainSection');

      if (e.target.value === 'single') {
        singleSection.classList.remove('hidden');
        multipleSection.classList.add('hidden');
      } else {
        singleSection.classList.add('hidden');
        multipleSection.classList.remove('hidden');
      }
    });
  });

  // Discovery domain combobox
  (function initDomainCombobox() {
    const input  = document.getElementById('discoveryDomain');
    const btn    = document.getElementById('domainDropdownBtn');
    const panel  = document.getElementById('domainDropdownPanel');
    const list   = document.getElementById('domainDropdownList');
    if (!input || !btn || !panel || !list) return;

    let highlightedIndex = -1;

    function selectDomain(domain) {
      input.value = domain;
      const vbuInput = document.getElementById('discoveryVbuId');
      if (vbuInput) vbuInput.value = _vbuMap[domain.toLowerCase()] || '';
      closePanel();
    }

    function getVisibleItems() {
      return Array.from(list.querySelectorAll('.domain-dropdown-item'));
    }

    function setHighlight(index) {
      const items = getVisibleItems();
      items.forEach(el => el.classList.remove('highlighted'));
      if (index >= 0 && index < items.length) {
        items[index].classList.add('highlighted');
        items[index].scrollIntoView({ block: 'nearest' });
      }
      highlightedIndex = index;
    }

    function renderItems(filter) {
      highlightedIndex = -1;
      const q = (filter || '').toLowerCase();
      const rows = _vbuRows.filter(r => !q || r.domain.includes(q) || (r.vbuName && r.vbuName.toLowerCase().includes(q)));
      if (rows.length === 0) {
        list.innerHTML = `<div class="domain-dropdown-empty">No matches</div>`;
      } else {
        list.innerHTML = rows.map(r =>
          `<div class="domain-dropdown-item" data-domain="${r.domain}">
            <span class="di-domain">${r.domain}</span>
            ${r.vbuName ? `<span class="di-vbu">${r.vbuName}</span>` : ''}
          </div>`
        ).join('');
        list.querySelectorAll('.domain-dropdown-item').forEach(el => {
          el.addEventListener('mousedown', e => { e.preventDefault(); selectDomain(el.dataset.domain); });
        });
      }
    }

    function openPanel() {
      // Don't show panel if no data is loaded — just let user type freely
      if (_vbuRows.length === 0) { closePanel(); return; }
      renderItems(input.value);
      panel.classList.remove('hidden');
    }

    function closePanel() {
      panel.classList.add('hidden');
      highlightedIndex = -1;
    }

    btn.addEventListener('click', () => {
      if (_vbuRows.length === 0) { input.focus(); return; }
      if (panel.classList.contains('hidden')) { input.focus(); openPanel(); } else { closePanel(); }
    });

    input.addEventListener('input', () => { openPanel(); });
    input.addEventListener('focus', () => { openPanel(); });
    input.addEventListener('blur', () => {
      setTimeout(() => {
        closePanel();
        // auto-fill VBU if exact match
        const domain = input.value.trim().toLowerCase();
        const vbuInput = document.getElementById('discoveryVbuId');
        if (vbuInput && _vbuMap[domain] !== undefined) vbuInput.value = _vbuMap[domain];
      }, 150);
    });

    input.addEventListener('keydown', e => {
      if (panel.classList.contains('hidden')) return;
      const items = getVisibleItems();
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setHighlight(Math.min(highlightedIndex + 1, items.length - 1));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setHighlight(Math.max(highlightedIndex - 1, 0));
      } else if (e.key === 'Enter' && highlightedIndex >= 0) {
        e.preventDefault();
        selectDomain(items[highlightedIndex].dataset.domain);
      } else if (e.key === 'Escape') {
        closePanel();
      }
    });

    document.addEventListener('click', e => {
      if (!input.contains(e.target) && !btn.contains(e.target) && !panel.contains(e.target)) closePanel();
    });
  })();

  // Discovery - Start button
  const startDiscoveryBtn = document.getElementById('startDiscoveryBtn');
  if (startDiscoveryBtn) {
    startDiscoveryBtn.addEventListener('click', async () => {
      const domainMode = document.querySelector('input[name="domainMode"]:checked')?.value || 'single';
      const vbuId = document.getElementById('discoveryVbuId').value.trim();
      const skipPP = document.getElementById('skipPowerPlatform').checked;
      const hybrid = document.getElementById('hybrid').checked;
      const members = document.getElementById('includeMembers').checked;
      const continueOnError = document.getElementById('continueOnError').checked;
      const outputFolder = document.getElementById('outputFolder').value.trim() || 'E:\\Work\\Jolera\\Volaris\\Discovery';

      let domainsToRun = [];
      if (domainMode === 'single') {
        const domain = document.getElementById('discoveryDomain').value.trim();
        if (!domain) { alert('Please enter a domain name'); return; }
        domainsToRun = [domain];
      } else {
        domainsToRun = (document.getElementById('discoveryDomains').value || '')
          .split('\n').map(d => d.trim().toLowerCase()).filter(d => d && !d.startsWith('#') && d.includes('.'));
        if (domainsToRun.length === 0) { alert('Please enter at least one domain name'); return; }
      }

      // Show log section
      const logSection = document.getElementById('discoveryLog');
      const logOutput = document.getElementById('discoveryLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = `Starting discovery for: ${domainsToRun.join(', ')}\n`;
      if (vbuId) logOutput.textContent += `VBU ID: ${vbuId}\n`;
      logOutput.textContent += `Options: SkipPP=${skipPP}  Hybrid=${hybrid}  Members=${members}\n`;
      logOutput.textContent += `Output: ${outputFolder}\n\n`;

      startDiscoveryBtn.disabled = true;
      startDiscoveryBtn.textContent = 'Running...';
      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });

      let lastResult;
      try {
        for (const domain of domainsToRun) {
          if (domainsToRun.length > 1) logOutput.textContent += `\n=== ${domain} ===\n`;
          const args = ['-Domain', domain, '-OutputPath', outputFolder];
          if (vbuId) args.push('-BusinessUnitId', vbuId);
          if (skipPP) args.push('-SkipPowerPlatform');
          if (hybrid) args.push('-Hybrid');
          if (members) args.push('-IncludeMembers');
          lastResult = await window.electronAPI.streamPowerShell('search-domain.ps1', args);
          if (!lastResult.success && !continueOnError) break;
        }
        logOutput.textContent += lastResult?.success ? '\n✓ Discovery complete\n' : `\n✗ Failed (exit ${lastResult?.code})\n`;
      } catch (err) {
        logOutput.textContent += `\n✗ Error: ${err.message}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        startDiscoveryBtn.disabled = false;
        startDiscoveryBtn.textContent = 'Start Discovery';
      }
    });
  }

  // Option launch button click handlers
  const launchBtns = document.querySelectorAll('.option-launch-btn');
  launchBtns.forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const optionItem = btn.closest('.option-item');
      const scriptName = optionItem.getAttribute('data-script');
      if (scriptName) {
        await launchScript(scriptName, btn);
      }
    });
  });

  // Logs button
  document.getElementById('logsBtn').addEventListener('click', async () => {
    try {
      const result = await window.electronAPI.openLogs();
      if (!result.success) {
        alert('Failed to open logs folder: ' + result.error);
      }
    } catch (error) {
      console.error('Error opening logs:', error);
      alert('Error opening logs folder: ' + error.message);
    }
  });

  // Settings button
  document.getElementById('settingsBtn').addEventListener('click', async () => {
    await openSettings();
  });

  // Settings tab switching
  const settingsTabs = document.querySelectorAll('.settings-tab');
  settingsTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const tabName = tab.getAttribute('data-tab');

      // Update active tab
      settingsTabs.forEach(t => t.classList.remove('settings-tab-active'));
      tab.classList.add('settings-tab-active');

      // Update active content
      document.querySelectorAll('.settings-tab-content').forEach(content => {
        content.classList.remove('settings-tab-content-active');
        content.classList.add('hidden');
      });

      const activeContent = document.getElementById(tabName + 'Tab');
      if (activeContent) {
        activeContent.classList.add('settings-tab-content-active');
        activeContent.classList.remove('hidden');
      }
    });
  });

  // Helper function to add empty customer row
  function addEmptyCustomerRow() {
    const tbody = document.getElementById('customerTableBody');
    const newRow = document.createElement('tr');

    const td1 = document.createElement('td');
    const td2 = document.createElement('td');
    const td3 = document.createElement('td');
    const td4 = document.createElement('td');

    const input1 = document.createElement('input');
    input1.type = 'text';
    input1.className = 'form-input form-input-compact customer-prefix';
    input1.placeholder = 'Project prefix';

    const input2 = document.createElement('input');
    input2.type = 'text';
    input2.className = 'form-input form-input-compact customer-account';
    input2.placeholder = 'account@domain.onmicrosoft.com';

    const input3 = document.createElement('input');
    input3.type = 'text';
    input3.className = 'form-input form-input-compact customer-domain';
    input3.placeholder = 'e.g. mbufara';

    const input4 = document.createElement('input');
    input4.type = 'text';
    input4.className = 'form-input form-input-compact customer-spo';
    input4.placeholder = 'https://tenant-admin.sharepoint.com';

    td1.appendChild(input1);
    td2.appendChild(input2);
    td3.appendChild(input3);
    td4.appendChild(input4);

    newRow.appendChild(td1);
    newRow.appendChild(td2);
    newRow.appendChild(td3);
    newRow.appendChild(td4);

    tbody.appendChild(newRow);
  }

  // Auto-populate domain and SPO from account name (e.g. admin@contoso.onmicrosoft.com → contoso)
  function autoFillCustomerRow(accountInput) {
    const row = accountInput.closest('tr');
    if (!row) return;
    const domainInput = row.querySelector('.customer-domain');
    const spoInput    = row.querySelector('.customer-spo');
    if (!domainInput || !spoInput) return;

    const m = accountInput.value.trim().match(/@([^@]+)\.onmicrosoft\.com$/i);
    if (!m) return;
    const tenant = m[1];

    if (!domainInput.value.trim()) domainInput.value = tenant;
    if (!spoInput.value.trim())    spoInput.value    = `https://${tenant}-admin.sharepoint.com`;
  }

  document.getElementById('customerTableBody')?.addEventListener('input', (e) => {
    if (e.target.classList.contains('customer-account')) autoFillCustomerRow(e.target);
  });

  // Add/Remove customer rows
  document.getElementById('addCustomerBtn')?.addEventListener('click', () => {
    addEmptyCustomerRow();
  });

  document.getElementById('removeCustomerBtn')?.addEventListener('click', () => {
    const tbody = document.getElementById('customerTableBody');
    if (tbody.rows.length > 1) {
      tbody.deleteRow(tbody.rows.length - 1);
    }
  });

  // Close button
  document.getElementById('closeBtn').addEventListener('click', () => {
    window.close();
  });

  // Settings dialog buttons
  document.getElementById('cancelSettingsBtn').addEventListener('click', closeSettings);
  document.getElementById('saveSettingsBtn').addEventListener('click', saveSettings);
  document.getElementById('testConnectionBtn').addEventListener('click', testConnection);
  document.getElementById('checkUpdatesBtn').addEventListener('click', manualCheckUpdates);
  document.getElementById('settingsAosSignInBtn').addEventListener('click', settingsAosSignIn);

  // Install update button
  document.getElementById('installUpdateBtn').addEventListener('click', installUpdate);

  // Close dialog on overlay click
  document.getElementById('settingsDialog').addEventListener('click', (e) => {
    if (e.target.classList.contains('dialog-overlay')) {
      closeSettings();
    }
  });

  // Monitor functionality
  let monitorTimer = null;
  let lastMonitorData = null;
  const monitorRefreshBtn = document.getElementById('monitorRefreshBtn');
  const monitorAutoRefresh = document.getElementById('monitorAutoRefresh');
  const monitorInterval = document.getElementById('monitorInterval');
  const monitorProject = document.getElementById('monitorProject');
  const monitorStatus = document.getElementById('monitorStatus');
  const monitorTableBody = document.getElementById('monitorTableBody');
  const monitorConnection = document.getElementById('monitorConnection');
  let _cachedPortalUrl = null;

  if (monitorRefreshBtn) {
    monitorRefreshBtn.addEventListener('click', async () => {
      await refreshMonitor();
    });
  }

  if (monitorAutoRefresh) {
    monitorAutoRefresh.addEventListener('change', () => {
      if (monitorAutoRefresh.checked) {
        startMonitorTimer();
      } else {
        stopMonitorTimer();
      }
    });
  }

  async function refreshMonitor() {
    const project = monitorProject.value.trim();
    if (!project) {
      monitorStatus.textContent = 'Please enter a project prefix';
      return;
    }

    monitorStatus.textContent = 'Refreshing...';
    monitorTableBody.innerHTML = '<tr class="monitor-empty-row"><td colspan="8">Loading project data...</td></tr>';
    updateConnectionStatus('checking');

    try {
      // Get real migration data from API
      const result = await window.electronAPI.getMigrationData(project);
      const now = new Date().toLocaleTimeString();

      if (result.success && result.data) {
        const data = result.data;
        lastMonitorData = data;
        monitorTableBody.innerHTML = '';

        // Display each workload
        const workloads = data.Workloads || {};
        if (Object.keys(workloads).length === 0) {
          monitorTableBody.innerHTML = '<tr class="monitor-empty-row"><td colspan="8">No workload data found for this project</td></tr>';
          monitorStatus.textContent = 'No data available';
          updateConnectionStatus('failed');
          return;
        }

        Object.keys(workloads).forEach(workloadName => {
          const wl = workloads[workloadName];
          const total = wl.Total || 0;
          const complete = wl.Completed || 0;
          const failed = wl.Failed || 0;
          const inProgress = wl.InProgress || 0;
          const notStarted = wl.NotStarted || 0;
          const warnings = wl.Warnings || 0;

          const row = document.createElement('tr');
          const projectFound = wl.ProjectFound !== false;
          row.style.cursor = 'pointer';
          row.title = projectFound ? 'Click to open in Fly portal' : 'Project not yet created';

          // Apply row styling based on status
          if (!projectFound) {
            row.style.opacity = '0.45';
          } else if (failed > 0) {
            row.classList.add('monitor-row-failed');
          } else if (warnings > 0) {
            row.classList.add('monitor-row-warning');
          } else if (complete === total && total > 0) {
            row.classList.add('monitor-row-success');
          }

          row.innerHTML = `
            <td>${project} - ${workloadName}</td>
            <td>${total}</td>
            <td>${notStarted}</td>
            <td>${inProgress}</td>
            <td>${complete}</td>
            <td>${warnings}</td>
            <td>${failed}</td>
            <td>${now}</td>
          `;

          row.addEventListener('click', async () => {
            if (!projectFound) return;
            if (_cachedPortalUrl === null) {
              try {
                const cfg = await window.electronAPI.getConfig();
                _cachedPortalUrl = (cfg.success && cfg.config?.PortalUrl) ? cfg.config.PortalUrl : '';
              } catch (_) { _cachedPortalUrl = ''; }
            }
            if (!_cachedPortalUrl) return;
            const wlPaths = {
              Exchange: 'exchange', SharePoint: 'sharepoint', OneDrive: 'onedrive',
              Teams: 'teams', TeamChat: 'teamchat', Groups: 'm365group'
            };
            const origin = _cachedPortalUrl.replace(/#.*$/, '').replace(/\/$/, '');
            const projectId = wl.ProjectId;
            const flyUrl = projectId
              ? `${origin}/#/project/${projectId}/mappings`
              : origin;
            await window.electronAPI.openExternal(flyUrl);
          });
          monitorTableBody.appendChild(row);
        });

        monitorStatus.textContent = `Last refresh: ${now} — click a row to open in Fly`;
        updateConnectionStatus('connected');
      } else {
        monitorTableBody.innerHTML = '<tr class="monitor-empty-row"><td colspan="8">Failed to load data - check console for details</td></tr>';
        monitorStatus.textContent = `Error: ${result.error || 'Unknown error'}`;
        updateConnectionStatus('failed');
        console.error('Monitor refresh error:', result);
      }
    } catch (error) {
      monitorTableBody.innerHTML = '<tr class="monitor-empty-row"><td colspan="8">Error loading data</td></tr>';
      monitorStatus.textContent = `Error: ${error.message}`;
      updateConnectionStatus('failed');
      console.error('Monitor refresh exception:', error);
    }
  }

  // ── Monitor detail modal ──────────────────────────────────────────────────
  const monitorDetailOverlay = document.getElementById('monitorDetailOverlay');
  const monitorDetailTitle   = document.getElementById('monitorDetailTitle');
  const monitorDetailBody    = document.getElementById('monitorDetailBody');
  const monitorDetailClose   = document.getElementById('monitorDetailClose');

  if (monitorDetailClose) {
    monitorDetailClose.addEventListener('click', () => monitorDetailOverlay.classList.add('hidden'));
  }
  if (monitorDetailOverlay) {
    monitorDetailOverlay.addEventListener('click', (e) => {
      if (e.target === monitorDetailOverlay) monitorDetailOverlay.classList.add('hidden');
    });
  }

  function showMonitorDetail(workloadName, data) {
    if (!data || !monitorDetailOverlay) return;

    monitorDetailTitle.textContent = `${monitorProject.value.trim()} — ${workloadName}`;

    const failed    = (data.FailedItems    || []).filter(i => i.Workload === workloadName);
    const warnings  = (data.WarningItems   || []).filter(i => i.Workload === workloadName);
    const inProg    = (data.InProgressItems|| []).filter(i => i.Workload === workloadName);
    const completed = (data.CompletedItems || []).filter(i => i.Workload === workloadName);

    // Store items by index so expand click handlers can look them up
    window._monitorItems = [...failed, ...warnings, ...inProg, ...completed];

    // Columns whose values we skip when hunting for error text (source/dest/status noise)
    const _noiseCol = /^(source|target|destination|.*principal.*name|upn|stage.?status|job.?progress|workload|project|status|state)$/i;

    // Finds the most meaningful error text for an item.
    // Strategy: captured Exception field → error-named AllFields key → longest AllFields value (≥60 chars) → Error/Warning fallback
    function findErrorText(item) {
      if (item.Exception && item.Exception.trim()) return item.Exception.trim();

      const af = item.AllFields || {};
      let bestNamed = '';   // longest value from a key matching error/exception/fail/reason/message/detail
      let longestVal = '';  // longest value overall (real error messages are much longer than status codes/emails)

      for (const [k, v] of Object.entries(af)) {
        if (_noiseCol.test(k)) continue;
        const s = String(v).trim();
        if (!s || s === '0' || s === '-') continue;
        if (/error|exception|fail|reason|message|detail/i.test(k) && s.length > bestNamed.length) bestNamed = s;
        if (s.length >= 60 && s.length > longestVal.length) longestVal = s;
      }

      return bestNamed || longestVal || item.Error || item.Warning || '';
    }

    // Returns summary text for the collapsed row (first line, truncated)
    function bestErrorSummary(item) {
      const full = findErrorText(item);
      if (!full) return item.Status || '';
      const firstLine = full.split(/\r?\n/)[0].trim();
      return firstLine.length > 160 ? firstLine.substring(0, 157) + '…' : firstLine;
    }

    // Builds the inline expanded detail — status info + Fly link
    function buildInlineDetail(item, flyUrl) {
      const isFailed = item.Status === 'Failed';
      const color = isFailed ? '#dc2626' : '#d97706';
      const errCount = item.ErrorCount > 0 ? ` (${item.ErrorCount} item${item.ErrorCount !== 1 ? 's' : ''} affected)` : '';
      const msg = isFailed
        ? `Migration failed${errCount} — open in Fly portal for full error details`
        : `Completed with exceptions${errCount} — open in Fly portal for details`;

      const msgHtml = flyUrl
        ? `<span style="font-size:12px;color:${color};flex:1;">${msg}</span>`
        : `<span style="font-size:12px;color:#6b7280;flex:1;">Error details only available in the Fly portal.</span>`;

      const flyLinkHtml = flyUrl
        ? `<a class="detail-fly-link" href="#" style="font-size:12px;font-weight:600;color:#2563eb;text-decoration:underline;white-space:nowrap;flex-shrink:0;padding-top:1px;">↗ Open in Fly</a>`
        : '';

      return `<div style="padding:7px 14px 9px 32px;background:#f9fafb;border-bottom:2px solid #e5e7eb;display:flex;gap:14px;align-items:flex-start;">
        ${msgHtml}
        ${flyLinkHtml}
      </div>`;
    }

    function section(title, color, items, renderRow) {
      if (!items.length) return '';
      const rows = items.map(renderRow).join('');
      return `
        <h3 style="margin:16px 0 8px; font-size:13px; font-weight:600; color:${color};">${title} (${items.length})</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px;">
          <tbody>${rows}</tbody>
        </table>`;
    }

    const tdStyle = 'padding:5px 8px; border-bottom:1px solid #f0f0f0; word-break:break-all;';

    const html = [
      section('Failed', '#dc2626', failed, i => {
        const idx = window._monitorItems.indexOf(i);
        const summary = bestErrorSummary(i);
        return `<tr class="monitor-item-row" data-item-idx="${idx}" style="cursor:pointer;" title="Open in Fly portal">
          <td style="${tdStyle} width:38%">${i.Name}</td>
          <td style="${tdStyle} color:#dc2626;">${summary}</td></tr>`;
      }),
      section('Warnings', '#d97706', warnings, i => {
        const idx = window._monitorItems.indexOf(i);
        const summary = bestErrorSummary(i);
        return `<tr class="monitor-item-row" data-item-idx="${idx}" style="cursor:pointer;" title="Open in Fly portal">
          <td style="${tdStyle} width:38%">${i.Name}</td>
          <td style="${tdStyle} color:#d97706;">${summary || i.Warning || i.Status}</td></tr>`;
      }),
      section('In Progress', '#2563eb', inProg, i =>
        `<tr style="cursor:default;"><td style="${tdStyle} width:38%">${i.Name}</td><td style="${tdStyle} color:#2563eb">${i.Status}</td><td style="${tdStyle} color:#9ca3af;font-size:11px;">currently running</td></tr>`),
      section('Recently Completed', '#16a34a', completed, i =>
        `<tr><td style="${tdStyle} width:38%">${i.Name}</td><td style="${tdStyle} color:#16a34a">${i.Status}</td></tr>`)
    ].join('');

    monitorDetailBody.innerHTML = html ||
      '<p style="color:#6b7280; font-size:13px;">No detail items available for this workload.</p>';

    const workloadPaths = {
      Exchange: 'exchange', SharePoint: 'sharepoint', OneDrive: 'onedrive',
      Teams: 'teams', TeamChat: 'teamchat', Groups: 'm365group'
    };

    monitorDetailBody.querySelectorAll('.monitor-item-row').forEach(row => {
      row.addEventListener('mouseenter', () => { row.style.backgroundColor = '#f9fafb'; });
      row.addEventListener('mouseleave', () => { row.style.backgroundColor = ''; });
      row.addEventListener('click', async () => {
        const idx  = parseInt(row.getAttribute('data-item-idx'), 10);
        const item = window._monitorItems[idx];

        if (_cachedPortalUrl === null) {
          try {
            const cfg = await window.electronAPI.getConfig();
            _cachedPortalUrl = (cfg.success && cfg.config?.PortalUrl) ? cfg.config.PortalUrl : '';
          } catch (_) { _cachedPortalUrl = ''; }
        }
        if (!_cachedPortalUrl) return;

        const origin = _cachedPortalUrl.replace(/#.*$/, '').replace(/\/$/, '');
        const flyUrl = item.ProjectId
          ? `${origin}/#/project/${item.ProjectId}/mappings`
          : origin;

        await window.electronAPI.openExternal(flyUrl);
      });
    });

    monitorDetailOverlay.classList.remove('hidden');
  }

  function updateConnectionStatus(status) {
    const dot = monitorConnection.querySelector('.connection-dot');
    const text = monitorConnection.querySelectorAll('span')[1]; // Second span is the text

    dot.className = 'connection-dot';
    if (status === 'connected') {
      dot.classList.add('connection-connected');
      text.textContent = 'Connected to AvePoint Fly';
    } else if (status === 'failed') {
      dot.classList.add('connection-failed');
      text.textContent = 'Connection failed';
    } else {
      dot.classList.add('connection-checking');
      text.textContent = 'Checking connection...';
    }
  }

  function startMonitorTimer() {
    stopMonitorTimer();
    const intervalText = monitorInterval.value;
    const minutes = parseInt(intervalText);
    const ms = minutes * 60 * 1000;

    monitorTimer = setInterval(() => {
      refreshMonitor();
    }, ms);
  }

  // ── Monitor table sorting ──────────────────────────────────────────────────
  let _monitorSortCol = -1;
  let _monitorSortAsc = true;

  function getMonitorCellValue(row, col) {
    const cells = row.querySelectorAll('td');
    if (!cells[col]) return '';
    const text = cells[col].textContent.trim();
    // Numeric columns (1-7 except 7 which is time string)
    if (col >= 1 && col <= 6) return parseInt(text, 10) || 0;
    return text.toLowerCase();
  }

  function applyMonitorSort() {
    if (_monitorSortCol < 0) return;
    const tbody = monitorTableBody;
    const rows = Array.from(tbody.querySelectorAll('tr:not(.monitor-empty-row)'));
    if (!rows.length) return;
    rows.sort((a, b) => {
      const va = getMonitorCellValue(a, _monitorSortCol);
      const vb = getMonitorCellValue(b, _monitorSortCol);
      if (va < vb) return _monitorSortAsc ? -1 : 1;
      if (va > vb) return _monitorSortAsc ? 1 : -1;
      return 0;
    });
    rows.forEach(r => tbody.appendChild(r));
    // Update arrow indicators
    document.querySelectorAll('#monitorTableHead th').forEach(th => {
      const arrow = th.querySelector('.sort-arrow');
      if (!arrow) return;
      const col = parseInt(th.getAttribute('data-col'), 10);
      arrow.textContent = col === _monitorSortCol ? (_monitorSortAsc ? ' ▲' : ' ▼') : '';
    });
  }

  document.getElementById('monitorTableHead')?.addEventListener('click', (e) => {
    const th = e.target.closest('th[data-col]');
    if (!th) return;
    const col = parseInt(th.getAttribute('data-col'), 10);
    if (_monitorSortCol === col) {
      _monitorSortAsc = !_monitorSortAsc;
    } else {
      _monitorSortCol = col;
      _monitorSortAsc = col === 0; // text column → A-Z first; numeric → desc first
    }
    applyMonitorSort();
  });

  // Re-apply sort after every refresh
  const _origRefreshMonitor = refreshMonitor;
  refreshMonitor = async function() {
    await _origRefreshMonitor();
    applyMonitorSort();
  };

  function stopMonitorTimer() {
    if (monitorTimer) {
      clearInterval(monitorTimer);
      monitorTimer = null;
    }
  }

  // Stop timer when switching views
  const originalSwitchView = switchView;
  switchView = function(viewName) {
    stopMonitorTimer();
    originalSwitchView(viewName);
  };

  // Dashboard functionality
  const dashboardDomainSelect = document.getElementById('dashboardDomainSelect');
  const dashboardRefreshBtn = document.getElementById('dashboardRefreshBtn');
  const dashboardSubtitle = document.getElementById('dashboardSubtitle');
  const workloadBars = document.getElementById('workloadBars');
  const openMonitorBtn = document.getElementById('openMonitorBtn');
  const openAvepointBtn = document.getElementById('openAvepointBtn');

  // Dashboard auto-refresh timer
  let dashboardTimer = null;

  function startDashboardAutoRefresh() {
    stopDashboardAutoRefresh();
    // Auto-refresh every 60 seconds
    dashboardTimer = setInterval(() => {
      const domain = dashboardDomainSelect?.value;
      if (domain && document.getElementById('dashboardView')?.classList.contains('hidden') === false) {
        console.log('Auto-refreshing dashboard...');
        refreshDashboard();
      }
    }, 60000); // 60 seconds
  }

  function stopDashboardAutoRefresh() {
    if (dashboardTimer) {
      clearInterval(dashboardTimer);
      dashboardTimer = null;
    }
  }

  if (dashboardRefreshBtn) {
    dashboardRefreshBtn.addEventListener('click', refreshDashboard);
  }

  if (dashboardDomainSelect) {
    dashboardDomainSelect.addEventListener('change', () => {
      refreshDashboard();
      // Restart auto-refresh timer when domain changes
      startDashboardAutoRefresh();
    });
  }

  if (openMonitorBtn) {
    openMonitorBtn.addEventListener('click', () => {
      switchView('avepoint-monitor');
    });
  }

  if (openAvepointBtn) {
    openAvepointBtn.addEventListener('click', async () => {
      // Open AvePoint Fly portal
      const config = await window.electronAPI.getConfig();
      if (config.success && config.config && config.config.PortalUrl) {
        await window.electronAPI.openExternal(config.config.PortalUrl);
      } else {
        alert('Portal URL not configured. Please set it in Settings > Config tab.');
      }
    });
  }

  // Populate domain dropdown from customer settings
  async function loadCustomerDomains() {
    try {
      const config = await window.electronAPI.getConfig();
      if (config.success && config.config && config.config.Customers) {
        const customers = [...config.config.Customers].sort((a, b) => (a.Prefix || '').localeCompare(b.Prefix || ''));
        dashboardDomainSelect.innerHTML = '<option value="">Select migration...</option>';

        customers.forEach(customer => {
          if (customer.Prefix) {
            const option = document.createElement('option');
            option.value = customer.Prefix;
            option.textContent = customer.Prefix;
            dashboardDomainSelect.appendChild(option);
          }
        });

        // Select first domain if available
        if (customers.length > 0 && customers[0].Prefix) {
          dashboardDomainSelect.value = customers[0].Prefix;
          refreshDashboard();
          // Start auto-refresh timer
          startDashboardAutoRefresh();
        }
      }
    } catch (error) {
      console.error('Error loading customer domains:', error);
    }
  }

  async function refreshDashboard() {
    const domain = dashboardDomainSelect.value;
    if (!domain) return;

    const domainText = dashboardDomainSelect.options[dashboardDomainSelect.selectedIndex].text;

    // Show loading state
    document.getElementById('statItemsInScope').textContent = '...';
    document.getElementById('statCompleted').textContent = '...';
    document.getElementById('statInProgress').textContent = '...';
    document.getElementById('statNeedsAttention').textContent = '...';

    try {
      // Get real migration data from PowerShell
      console.log('Fetching migration data for:', domain);
      const result = await window.electronAPI.getMigrationData(domain);
      console.log('Migration data result:', result);

      if (result.success && result.data) {
        console.log('Using real data:', result.data);
        const data = result.data;

        const totalItems = data.TotalItems || 0;
        const completed = data.Completed || 0;
        const inProgress = data.InProgress || 0;
        const failed = data.Failed || 0;
        const warnings = data.Warnings || 0;
        const completedPercent = totalItems > 0 ? Math.round((completed / totalItems) * 100) : 0;

        // Store data for detail views
        currentStatData = {
          totalItems,
          completed,
          inProgress,
          failed,
          warnings,
          completedPercent,
          notStarted: data.NotStarted || 0,
          workloads: data.Workloads || {},
          failedItems: data.FailedItems || [],
          warningItems: data.WarningItems || [],
          inProgressItems: data.InProgressItems || [],
          completedItems: data.CompletedItems || []
        };

        // Update UI with real data
        document.getElementById('statItemsInScope').textContent = totalItems.toLocaleString();
        const activeWorkloads = Object.keys(data.Workloads || {}).length;
        document.getElementById('statItemsDetail').textContent = `across ${activeWorkloads} workloads`;
        document.getElementById('statCompleted').textContent = completed.toLocaleString();
        document.getElementById('statCompletedPercent').textContent = `${completedPercent}% of scope`;
        document.getElementById('statInProgress').textContent = inProgress.toLocaleString();
        document.getElementById('statNeedsAttention').textContent = failed + warnings;
        document.getElementById('statAttentionDetail').textContent = `${failed} failed · ${warnings} warnings`;
      } else {
        // Fallback to sample data if API fails
        console.warn('Failed to get real data, using sample data. Result:', result);

        // Only show detailed error if user manually clicked refresh
        // Don't show on auto-load to avoid annoying popup on startup
        const errorMsg = result.error || 'Failed to parse migration data';
        console.error('Migration data error:', errorMsg);

        // Show subtle error in UI instead of popup
        document.getElementById('statItemsInScope').textContent = '—';
        document.getElementById('statItemsDetail').textContent = 'No data available';
        document.getElementById('statCompleted').textContent = '—';
        document.getElementById('statCompletedPercent').textContent = 'Configure API in Settings';
        document.getElementById('statInProgress').textContent = '—';
        document.getElementById('statNeedsAttention').textContent = '—';
        document.getElementById('statAttentionDetail').textContent = errorMsg;
      }
    } catch (error) {
      console.error('Error fetching migration data:', error);

      // Show error in UI
      document.getElementById('statItemsInScope').textContent = '—';
      document.getElementById('statItemsDetail').textContent = 'Error loading data';
      document.getElementById('statCompleted').textContent = '—';
      document.getElementById('statCompletedPercent').textContent = error.message;
      document.getElementById('statInProgress').textContent = '—';
      document.getElementById('statNeedsAttention').textContent = '—';
      document.getElementById('statAttentionDetail').textContent = 'Check console for details';
    }

    // Update workload bars
    updateWorkloadBars();
  }

  function useSampleData() {
    const totalItems = Math.floor(Math.random() * 500) + 500;
    const completed = Math.floor(totalItems * (Math.random() * 0.3 + 0.6));
    const inProgress = Math.floor((totalItems - completed) * (Math.random() * 0.4 + 0.3));
    const failed = Math.floor(Math.random() * 5) + 2;
    const warnings = Math.floor(Math.random() * 10) + 3;
    const completedPercent = Math.round((completed / totalItems) * 100);

    currentStatData = {
      totalItems,
      completed,
      inProgress,
      failed,
      warnings,
      completedPercent,
      workloads: {},
      failedItems: [],
      warningItems: [],
      inProgressItems: [],
      completedItems: []
    };

    document.getElementById('statItemsInScope').textContent = totalItems.toLocaleString();
    document.getElementById('statItemsDetail').textContent = 'across 4 workloads';
    document.getElementById('statCompleted').textContent = completed.toLocaleString();
    document.getElementById('statCompletedPercent').textContent = `${completedPercent}% of scope`;
    document.getElementById('statInProgress').textContent = inProgress.toLocaleString();
    document.getElementById('statNeedsAttention').textContent = failed + warnings;
    document.getElementById('statAttentionDetail').textContent = `${failed} failed · ${warnings} warnings`;

    updateWorkloadBars();
  }

  function buildBarGradient(completed, inProgress, failed) {
    const filled = completed + inProgress + failed;
    if (filled === 0) return null;

    // Segment boundaries as % of the filled portion
    const gEnd = (completed   / filled) * 100;
    const oEnd = ((completed + inProgress) / filled) * 100;
    const blend = 7; // blend overlap in percentage points

    // Single-colour fast paths
    if (inProgress === 0 && failed === 0) return 'linear-gradient(90deg, #16a34a, #22c55e)';
    if (completed  === 0 && failed === 0) return 'linear-gradient(90deg, #ea580c, #f97316)';
    if (completed  === 0 && inProgress === 0) return 'linear-gradient(90deg, #dc2626, #ef4444)';

    const stops = [];
    if (completed > 0) {
      stops.push('#22c55e 0%');
      stops.push(`#22c55e ${Math.max(gEnd - blend, 0).toFixed(1)}%`);
    }
    if (inProgress > 0) {
      stops.push(`#f97316 ${(completed > 0 ? Math.min(gEnd + blend, oEnd) : 0).toFixed(1)}%`);
      if (failed > 0) {
        stops.push(`#f97316 ${Math.max(oEnd - blend, gEnd).toFixed(1)}%`);
      } else {
        stops.push('#f97316 100%');
      }
    }
    if (failed > 0) {
      stops.push(`#ef4444 ${((completed > 0 || inProgress > 0) ? Math.min(oEnd + blend, 100) : 0).toFixed(1)}%`);
      stops.push('#ef4444 100%');
    }
    return `linear-gradient(90deg, ${stops.join(', ')})`;
  }

  function updateWorkloadBars() {
    workloadBars.innerHTML = '';

    const workloadsData = currentStatData.workloads || {};

    if (Object.keys(workloadsData).length === 0) {
      workloadBars.innerHTML = '<p style="color: #6c757d; padding: 20px; text-align: center;">No workload data available</p>';
      return;
    }

    Object.keys(workloadsData).forEach(workloadName => {
      const wlData     = workloadsData[workloadName];
      const total      = wlData.Total      || 0;
      const completed  = wlData.Completed  || 0;
      const failed     = wlData.Failed     || 0;
      const inProgress = wlData.InProgress || 0;

      const filledPct  = total > 0 ? ((completed + inProgress + failed) / total) * 100 : 0;
      const gradient   = buildBarGradient(completed, inProgress, failed);
      const fillStyle  = gradient
        ? `width:${filledPct.toFixed(1)}%; background:${gradient};`
        : `width:0%;`;

      // Status badge — show all active counts
      const parts = [];
      if (completed > 0)  parts.push(`<span class="status-badge badge-ontrack">${completed} done</span>`);
      if (inProgress > 0) parts.push(`<span class="status-badge badge-warning">${inProgress} in progress</span>`);
      if (failed > 0)     parts.push(`<span class="status-badge badge-failed">${failed} failed</span>`);
      if (parts.length === 0 && total > 0) parts.push(`<span class="status-badge badge-neutral">Not started</span>`);

      const bar = document.createElement('div');
      bar.className = 'workload-bar-item';
      bar.setAttribute('data-workload', workloadName.toLowerCase());

      bar.innerHTML = `
        <div class="workload-bar-name">${workloadName}</div>
        <div class="workload-bar-progress">
          <div class="workload-bar-fill" style="${fillStyle}"></div>
        </div>
        <div class="workload-bar-count">${completed}/${total}</div>
        <div class="workload-bar-status">${parts.join('')}</div>
      `;

      workloadBars.appendChild(bar);
    });
  }

  // Initialize dashboard on load
  if (workloadBars) {
    loadCustomerDomains();
  }

  // Dashboard stat box click handlers
  let currentStatData = {};

  function showStatDetail(statType) {
    const dialog = document.getElementById('statDetailDialog');
    const title = document.getElementById('statDetailTitle');
    const content = document.getElementById('statDetailContent');

    let titleText = '';
    let contentHTML = '';

    switch(statType) {
      case 'scope':
        titleText = 'Items in Scope - Details';
        contentHTML = `
          <h3 style="margin-top: 0;">Breakdown by Workload</h3>
          <table class="detail-table">
            <thead>
              <tr>
                <th>Workload</th>
                <th>Total Items</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><span class="workload-icon-inline">☁️</span> SharePoint</td>
                <td>${currentStatData.sharepoint || 0}</td>
                <td><span class="status-badge badge-ontrack">Active</span></td>
              </tr>
              <tr>
                <td><span class="workload-icon-inline">📧</span> Exchange</td>
                <td>${currentStatData.exchange || 0}</td>
                <td><span class="status-badge badge-ontrack">Active</span></td>
              </tr>
              <tr>
                <td><span class="workload-icon-inline">📁</span> OneDrive</td>
                <td>${currentStatData.onedrive || 0}</td>
                <td><span class="status-badge badge-ontrack">Active</span></td>
              </tr>
              <tr>
                <td><span class="workload-icon-inline">👥</span> Teams</td>
                <td>${currentStatData.teams || 0}</td>
                <td><span class="status-badge badge-ontrack">Active</span></td>
              </tr>
            </tbody>
          </table>
        `;
        break;

      case 'completed':
        titleText = 'Completed Items - Details';
        contentHTML = `
          <h3 style="margin-top: 0;">Recently Completed</h3>
          <div class="detail-list">
            ${generateCompletedItems()}
          </div>
          <div style="margin-top: 24px; padding: 16px; background: #e8f5e9; border-radius: 8px; border-left: 4px solid #28a745;">
            <strong>Success Rate:</strong> ${currentStatData.completedPercent}% of total scope completed successfully
          </div>
        `;
        break;

      case 'inprogress':
        titleText = 'In Progress - Details';
        contentHTML = `
          <h3 style="margin-top: 0;">Currently Migrating</h3>
          <table class="detail-table">
            <thead>
              <tr>
                <th>Item</th>
                <th>Type</th>
                <th>Progress</th>
                <th>ETA</th>
              </tr>
            </thead>
            <tbody>
              ${generateInProgressItems()}
            </tbody>
          </table>
        `;
        break;

      case 'attention':
        titleText = 'Needs Attention - Details';
        contentHTML = `
          <h3 style="margin-top: 0; color: #dc3545;">Items Requiring Attention</h3>

          <div class="attention-section">
            <h4 style="color: #c82333;">❌ Failed Items (${currentStatData.failed || 0})</h4>
            <table class="detail-table">
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Workload</th>
                  <th>Error</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                ${generateFailedItems()}
              </tbody>
            </table>
          </div>

          <div class="attention-section" style="margin-top: 24px;">
            <h4 style="color: #e65100;">⚠️ Warnings (${currentStatData.warnings || 0})</h4>
            <table class="detail-table">
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Workload</th>
                  <th>Warning</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                ${generateWarningItems()}
              </tbody>
            </table>
          </div>
        `;
        break;
    }

    title.textContent = titleText;
    content.innerHTML = contentHTML;
    dialog.classList.remove('hidden');
  }

  function generateCompletedItems() {
    const items = currentStatData.completedItems && currentStatData.completedItems.length > 0
      ? currentStatData.completedItems
      : [
          { Name: 'No completed items yet', Workload: '-', Status: '-' }
        ];

    return items.map(item => `
      <div class="detail-item">
        <div class="detail-item-header">
          <strong>${item.Name || item.name}</strong>
          <span class="detail-item-badge">${item.Workload || item.type}</span>
        </div>
        <div class="detail-item-meta">
          <span>✅ ${item.Status || 'Completed'}</span>
        </div>
      </div>
    `).join('');
  }

  function generateInProgressItems() {
    const items = currentStatData.inProgressItems && currentStatData.inProgressItems.length > 0
      ? currentStatData.inProgressItems.slice(0, 10)
      : [];

    if (items.length === 0) {
      return '<tr><td colspan="4" style="text-align: center; color: var(--color-ink-muted);">No items currently in progress</td></tr>';
    }

    return items.map((item, idx) => `
      <tr>
        <td><strong>${item.Name || item.name}</strong></td>
        <td>${item.Workload || item.type}</td>
        <td>
          <div class="inline-progress">
            <div class="inline-progress-bar">
              <div class="inline-progress-fill" style="width: ${(idx + 1) * 10 + 30}%"></div>
            </div>
            <span>${(idx + 1) * 10 + 30}%</span>
          </div>
        </td>
        <td>Calculating...</td>
      </tr>
    `).join('');
  }

  function generateFailedItems() {
    const items = currentStatData.failedItems && currentStatData.failedItems.length > 0
      ? currentStatData.failedItems
      : [];

    if (items.length === 0) {
      return '<tr><td colspan="4" style="text-align: center; color: #28a745;">✅ No failed items - all migrations successful!</td></tr>';
    }

    return items.map(item => `
      <tr>
        <td><strong>${item.Name || item.name}</strong></td>
        <td>${item.Workload || item.workload}</td>
        <td><span style="color: #dc3545;">${item.Error || item.error}</span></td>
        <td><button class="btn btn-secondary btn-sm">Retry</button></td>
      </tr>
    `).join('');
  }

  function generateWarningItems() {
    const items = currentStatData.warningItems && currentStatData.warningItems.length > 0
      ? currentStatData.warningItems
      : [];

    if (items.length === 0) {
      return '<tr><td colspan="4" style="text-align: center; color: #28a745;">✅ No warnings - all migrations clean!</td></tr>';
    }

    return items.map(item => `
      <tr>
        <td><strong>${item.Name || item.name}</strong></td>
        <td>${item.Workload || item.workload}</td>
        <td><span style="color: #e65100;">${item.Warning || item.warning}</span></td>
        <td>${item.Status || item.status}</td>
      </tr>
    `).join('');
  }

  // Attach click handlers to stat boxes
  document.addEventListener('click', (e) => {
    const statBox = e.target.closest('.dashboard-stat-box');
    if (statBox) {
      const boxes = Array.from(document.querySelectorAll('.dashboard-stat-box'));
      const index = boxes.indexOf(statBox);

      const statTypes = ['scope', 'completed', 'inprogress', 'attention'];
      if (index >= 0 && index < statTypes.length) {
        showStatDetail(statTypes[index]);
      }
    }
  });

  // Close stat detail dialog
  document.getElementById('closeStatDetailBtn')?.addEventListener('click', () => {
    document.getElementById('statDetailDialog').classList.add('hidden');
  });

  // Close on overlay click
  document.getElementById('statDetailDialog')?.addEventListener('click', (e) => {
    if (e.target.classList.contains('dialog-overlay')) {
      document.getElementById('statDetailDialog').classList.add('hidden');
    }
  });
});

// Populate OneDrive tenant URL dropdown from customer settings
async function loadOneDriveTenants() {
  try {
    console.log('Loading OneDrive tenants...');
    const onedriveTenantUrl = document.getElementById('onedriveTenantUrl');
    if (!onedriveTenantUrl) {
      console.log('onedriveTenantUrl element not found');
      return;
    }

    const config = await window.electronAPI.getConfig();
    console.log('Config loaded:', config);

    if (config.success && config.config && config.config.Customers) {
      const customers = [...config.config.Customers].sort((a, b) => (a.Prefix || '').localeCompare(b.Prefix || ''));
      console.log('Customers found:', customers.length);
      onedriveTenantUrl.innerHTML = '<option value="">Select tenant...</option>';

      customers.forEach(customer => {
        console.log('Processing customer:', customer.Prefix, customer.SharePointAdminUrl);
        if (customer.SharePointAdminUrl) {
          const option = document.createElement('option');
          option.value = customer.SharePointAdminUrl;
          option.textContent = `${customer.Prefix} - ${customer.SharePointAdminUrl}`;
          onedriveTenantUrl.appendChild(option);
        }
      });
      console.log('Dropdown populated with', onedriveTenantUrl.options.length, 'options');
    }
  } catch (error) {
    console.error('Error loading OneDrive tenants:', error);
  }
}

// Populate Monitor project dropdown from customer settings
async function loadMonitorProjects() {
  try {
    const monitorProject = document.getElementById('monitorProject');
    if (!monitorProject) return;

    const config = await window.electronAPI.getConfig();
    if (config.success && config.config && config.config.Customers) {
      const customers = [...config.config.Customers].sort((a, b) => (a.Prefix || '').localeCompare(b.Prefix || ''));
      monitorProject.innerHTML = '<option value="">Select project...</option>';

      customers.forEach(customer => {
        if (customer.Prefix) {
          const option = document.createElement('option');
          option.value = customer.Prefix;
          option.textContent = customer.Prefix;
          monitorProject.appendChild(option);
        }
      });
    }
  } catch (error) {
    console.error('Error loading monitor projects:', error);
  }
}

// Load saved AOS tenant details into the AOS Setup view
async function loadAosConfig() {
  try {
    const result = await window.electronAPI.getSharedConfig();
    if (!result.success || !result.config) return;
    const cfg = result.config;
    const dn = document.getElementById('aosDisplayName');
    const sc = document.getElementById('aosSearchCode');
    const pn = document.getElementById('aosProfileName');
    if (dn && cfg.TenantName)   dn.value = cfg.TenantName;
    if (sc && cfg.TenantSearch) sc.value = cfg.TenantSearch;
    if (pn && cfg.AppProfileName) pn.value = cfg.AppProfileName;
    else if (pn && cfg.TenantName && !pn.value) pn.value = cfg.TenantName + ' App';
  } catch (error) {
    console.error('Error loading AOS config:', error);
  }
}

// Populate customer dropdown in Connections view
async function loadConnectionsCustomers() {
  try {
    const psCustomerPrefix = document.getElementById('psCustomerPrefix');
    if (!psCustomerPrefix) {
      console.log('psCustomerPrefix element not found');
      return;
    }

    const config = await window.electronAPI.getConfig();

    if (config.success && config.config && config.config.Customers) {
      const customers = [...config.config.Customers].sort((a, b) => (a.Prefix || '').localeCompare(b.Prefix || ''));
      console.log('Loading customers into dropdown:', customers.length);

      psCustomerPrefix.innerHTML = '<option value="">Select customer...</option>';

      customers.forEach(customer => {
        if (customer.Prefix) {
          const option = document.createElement('option');
          option.value = customer.Prefix;
          option.textContent = customer.Prefix;
          psCustomerPrefix.appendChild(option);
        }
      });

      console.log('Customer dropdown populated with', psCustomerPrefix.options.length, 'options');
    }
  } catch (error) {
    console.error('Error loading connections customers:', error);
  }
}

function fitConnectionsLog() {
  const logPre = document.getElementById('connMappingsLogPre');
  const view   = document.getElementById('avepointConnectionsView');
  if (!logPre || !view || view.classList.contains('hidden')) return;
  requestAnimationFrame(() => {
    const top     = logPre.getBoundingClientRect().top;
    const content = document.querySelector('.content');
    const padBot  = content ? (parseInt(getComputedStyle(content).paddingBottom) || 40) : 40;
    const height  = window.innerHeight - top - padBot;
    logPre.style.height = Math.max(80, height) + 'px';
  });
}

window.addEventListener('resize', fitConnectionsLog);

// Show full error/exception detail for a single migration item
async function showItemDetail(item) {
  if (!item) return;
  const overlay = document.getElementById('itemDetailOverlay');
  const title   = document.getElementById('itemDetailTitle');
  const body    = document.getElementById('itemDetailBody');
  if (!overlay) return;

  title.textContent = item.Name || 'Item Details';
  body.innerHTML = '<p style="color:#6b7280;font-size:13px;">Loading...</p>';
  overlay.classList.remove('hidden');

  // Get portal URL for the Fly link
  let portalUrl = null;
  try {
    const cfg = await window.electronAPI.getConfig();
    if (cfg.success && cfg.config) portalUrl = cfg.config.PortalUrl || null;
  } catch (_) {}

  // Workload → Fly portal URL path segment
  const workloadPaths = {
    Exchange: 'exchange', SharePoint: 'sharepoint', OneDrive: 'onedrive',
    Teams: 'teams', TeamChat: 'teamchat', Groups: 'm365group'
  };
  const flyUrl = portalUrl
    ? `${portalUrl.replace(/\/$/, '')}` +
      (workloadPaths[item.Workload] ? `/migration/${workloadPaths[item.Workload]}/projects` : '')
    : null;

  function field(label, value, color) {
    if (!value) return '';
    return `<div style="margin-bottom:14px;">
      <div style="font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;">${label}</div>
      <div style="font-size:13px;color:${color || '#1c1c20'};word-break:break-word;">${value}</div>
    </div>`;
  }

  const errorValue = item.Error || item.Warning || '';
  const errorColor = item.Error ? '#dc2626' : item.Warning ? '#d97706' : '#1c1c20';

  // Build "All Fly Fields" section from the raw CSV row — this shows every column
  // AvePoint Fly exported, including error codes and exception detail regardless of column naming
  const skipKeys = new Set(['Name','Destination','Workload','Project','Status','Error','Warning',
    'Exception','TotalItems','MigratedItems','FailedItemCount','LastRunTime','AllFields']);
  let allFieldsHtml = '';
  const allFields = item.AllFields || {};
  const flyFieldEntries = Object.entries(allFields).filter(([k]) => !skipKeys.has(k) && k !== 'SourceUserPrincipalName');
  if (flyFieldEntries.length) {
    const rows = flyFieldEntries.map(([k, v]) => {
      const isError = /error|exception|fail/i.test(k);
      const color = isError ? (item.Error ? '#dc2626' : '#d97706') : '#374151';
      return `<tr>
        <td style="padding:4px 10px 4px 0;font-size:11px;font-weight:600;color:#6b7280;white-space:nowrap;vertical-align:top;">${k}</td>
        <td style="padding:4px 0;font-size:12px;color:${color};word-break:break-word;">${v}</td>
      </tr>`;
    }).join('');
    allFieldsHtml = `<div style="margin-bottom:14px;">
      <div style="font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:6px;">Fly Migration Data</div>
      <table style="width:100%;border-collapse:collapse;">${rows}</table>
    </div>`;
  }

  const flyLinkHtml = flyUrl ? `
    <div style="margin-bottom:16px;">
      <button id="openInFlyBtn" class="btn btn-secondary btn-compact" style="font-size:12px;">
        ↗ Open in AvePoint Fly
      </button>
    </div>` : '';

  body.innerHTML = `
    <div>
      ${flyLinkHtml}
      ${field('Source', item.Name)}
      ${field('Destination', item.Destination)}
      ${item.Project ? field('Project', item.Project) : ''}
      ${field('Workload', item.Workload)}
      ${field('Status / Error', errorValue, errorColor)}
      ${item.Exception ? `<div style="margin-bottom:14px;">
        <div style="font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;">Exception</div>
        <pre style="font-size:11px;background:#1a1b26;color:#cdd4e6;padding:14px;border-radius:6px;overflow-x:auto;white-space:pre-wrap;word-break:break-word;margin:0;line-height:1.5;">${item.Exception}</pre>
      </div>` : ''}
      ${allFieldsHtml}
      ${field('Last Run', item.LastRunTime)}
    </div>`;

  if (flyUrl) {
    document.getElementById('openInFlyBtn')?.addEventListener('click', async () => {
      await window.electronAPI.openExternal(flyUrl);
    });
  }
}

// Switch between views
function switchView(viewName) {
  console.log('Switching to view:', viewName);

  // Hide all views
  const allViews = document.querySelectorAll('.view-container');
  allViews.forEach(view => view.classList.add('hidden'));

  // Show selected view
  const viewMap = {
    'dashboard': 'dashboardView',
    'discovery': 'discoveryView',
    // AvePoint Fly sub-views
    'avepoint-appreg': 'avepointAppRegView',
    'avepoint-aos': 'avepointAosView',
    'avepoint-connections': 'avepointConnectionsView',
    'avepoint-reports': 'avepointReportsView',
    'avepoint-monitor': 'avepointMonitorView',
    // Misc Scripts sub-views
    'misc-onedrive': 'miscOneDriveView',
    'misc-teams': 'miscTeamsView',
    'misc-deduplicate': 'miscDeduplicateView',
    'misc-purge-spo':     'miscPurgeSpoView',
    'misc-domain-devices': 'miscDomainDevicesView',
    // Domain Removal sub-views
    'domain-workflow': 'domainWorkflowView',
    'domain-remove': 'domainRemoveView',
    'domain-onprem': 'domainOnPremView',
    'domain-cloud': 'domainCloudView',
    'domain-hide': 'domainHideView',
    'domain-alias': 'domainAliasView',
    'domain-sip': 'domainSIPView',
    'domain-entra-remove': 'domainEntraRemoveView'
  };

  const targetViewId = viewMap[viewName];
  if (targetViewId) {
    const targetView = document.getElementById(targetViewId);
    if (targetView) {
      targetView.classList.remove('hidden');

      // Load dropdowns when specific views are shown
      if (viewName === 'discovery') {
        loadDiscoveryDomains();
      } else if (viewName === 'misc-onedrive') {
        loadOneDriveTenants();
      } else if (viewName === 'avepoint-monitor') {
        loadMonitorProjects();
      } else if (viewName === 'avepoint-aos') {
        loadAosConfig();
      } else if (viewName === 'avepoint-connections') {
        loadConnectionsCustomers();
        fitConnectionsLog();
      } else if (viewName === 'dashboard') {
        // Restart dashboard auto-refresh when returning to dashboard
        const domain = dashboardDomainSelect?.value;
        if (domain) {
          startDashboardAutoRefresh();
        }
      }
    }
  }

  // Stop dashboard auto-refresh when navigating away from dashboard
  if (viewName !== 'dashboard') {
    stopDashboardAutoRefresh();
  }
}

// ── Discovery: domain/VBU dropdown ───────────────────────────────────────────
let _vbuMap = {};
let _vbuRows = [];

async function loadDiscoveryDomains() {
  try {
    const cfgResult = await window.electronAPI.getConfig();
    const csvPath = cfgResult?.config?.VbuCsvPath;
    if (!csvPath) return;

    const result = await window.electronAPI.readVbuCsv(csvPath);
    if (!result.success || !result.rows.length) return;

    _vbuMap = {};
    _vbuRows = result.rows;
    result.rows.forEach(r => { _vbuMap[r.domain] = r.vbuId; });
  } catch (err) {
    console.error('loadDiscoveryDomains failed:', err);
  }
}

// Launch PowerShell script
async function launchScript(scriptName, buttonElement) {
  try {
    console.log(`Launching script: ${scriptName}`);

    // Check if electronAPI is available
    if (!window.electronAPI) {
      console.error('electronAPI not found! Check preload.js');
      alert('Error: electronAPI not available. Make sure preload.js is loaded.');
      return;
    }

    // Show visual feedback - button flash
    if (buttonElement) {
      buttonElement.style.transform = 'scale(0.98)';
      setTimeout(() => {
        buttonElement.style.transform = '';
      }, 150);
    }

    const result = await window.electronAPI.launchScript(scriptName);
    console.log('Launch result:', result);

    if (result.success) {
      console.log('✓ Script launched successfully');

      // Show success flash
      if (buttonElement) {
        const originalFilter = buttonElement.style.filter;
        buttonElement.style.filter = 'brightness(1.2)';
        setTimeout(() => {
          buttonElement.style.filter = originalFilter;
        }, 300);
      }
    } else {
      console.error('Script launch failed:', result.error);
      alert(`Failed to launch script:\n${result.error}`);
    }
  } catch (error) {
    console.error('Error launching script:', error);
    alert(`Error launching script:\n${error.message}`);
  }
}

// Load version
async function loadVersion() {
  try {
    const result = await window.electronAPI.getVersion();
    if (result.success) {
      document.getElementById('versionDisplay').textContent = `Version ${result.version}`;
    }
  } catch (error) {
    console.error('Error loading version:', error);
  }
}

// Check for updates
async function checkForUpdates() {
  try {
    // This would normally check GitHub, but for now we'll skip auto-check
    // The PowerShell script handles this
    console.log('Update check skipped (handled by PowerShell)');
  } catch (error) {
    console.error('Error checking updates:', error);
  }
}

// AOS sign-in from Settings Config tab
async function settingsAosSignIn() {
  const btn = document.getElementById('settingsAosSignInBtn');
  const statusSpan = document.getElementById('settingsAosSessionStatus');
  btn.disabled = true;
  btn.textContent = 'Opening browser...';
  statusSpan.style.color = 'var(--color-text-muted)';
  statusSpan.textContent = 'Waiting for sign-in...';
  try {
    await window.electronAPI.streamPowerShell('Login-AOS.ps1', []);
    statusSpan.style.color = 'var(--color-success, #4caf50)';
    statusSpan.textContent = 'Session saved';
  } catch (err) {
    statusSpan.style.color = 'var(--color-error, #f44336)';
    statusSpan.textContent = `Failed: ${err.message || err}`;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Sign in to AOS';
  }
}

// Manual check for updates (from settings) — streams output so user can see progress
async function manualCheckUpdates() {
  const statusSpan = document.getElementById('updateStatus');
  const logPre = document.getElementById('updateLog');
  const btn = document.getElementById('checkUpdatesBtn');

  btn.disabled = true;
  statusSpan.textContent = 'Checking...';
  statusSpan.style.color = '#6c757d';
  logPre.textContent = '';
  logPre.style.display = 'block';

  let accumulated = '';

  const onOutput = (text) => {
    accumulated += text;
    logPre.textContent += text;
    logPre.scrollTop = logPre.scrollHeight;
  };

  window.electronAPI.onPsOutput(onOutput);

  try {
    const result = await window.electronAPI.streamPowerShell('Check-Updates.ps1', ['-Force']);

    if (result.success || result.code === 0) {
      if (accumulated.includes('UPDATE_AVAILABLE')) {
        statusSpan.textContent = '✓ Update available!';
        statusSpan.style.color = '#28a745';
      } else {
        statusSpan.textContent = '✓ Already up to date';
        statusSpan.style.color = '#28a745';
      }
    } else {
      statusSpan.textContent = `❌ Check failed (exit ${result.code ?? '?'})`;
      statusSpan.style.color = '#dc3545';
    }
  } catch (error) {
    statusSpan.textContent = `❌ Error: ${error.message}`;
    statusSpan.style.color = '#dc3545';
  } finally {
    window.electronAPI.offPsOutput(onOutput);
    btn.disabled = false;
  }
}

// Install update
async function installUpdate() {
  const btn = document.getElementById('installUpdateBtn');
  const originalText = btn.textContent;

  try {
    btn.textContent = 'Installing...';
    btn.disabled = true;

    await window.electronAPI.launchScript('Check-Updates.ps1');

    btn.textContent = '✓ Complete';
    setTimeout(() => {
      document.getElementById('updateBanner').classList.add('hidden');
    }, 2000);
  } catch (error) {
    btn.textContent = originalText;
    btn.disabled = false;
    alert(`Update failed:\n${error.message}`);
  }
}

// Open settings dialog
async function openSettings() {
  try {
    // Load current config
    const result = await window.electronAPI.getConfig();

    if (result.success && result.config) {
      // Config tab
      document.getElementById('apiUrl').value = result.config.Url || '';
      document.getElementById('clientId').value = result.config.ClientId || '';

      // Show placeholder dots if secret exists, otherwise leave empty
      const clientSecretInput = document.getElementById('clientSecret');
      if (result.config.EncSecret) {
        clientSecretInput.placeholder = '••••••••••••••••';
      } else {
        clientSecretInput.placeholder = 'Enter Client Secret';
      }
      clientSecretInput.value = ''; // Never show actual secret

      document.getElementById('portalUrl').value = result.config.PortalUrl || '';
      document.getElementById('sharePointAdminUrl').value = result.config.SharePointAdminUrl || '';
      document.getElementById('secretExpiry').value = result.config.SecretExpiry || '';
      document.getElementById('discoveryOutputPath').value = result.config.DiscoveryOutputPath || '';
      document.getElementById('vbuCsvPath').value = result.config.VbuCsvPath || '';

      // Load customers into table
      const customerTableBody = document.getElementById('customerTableBody');
      customerTableBody.innerHTML = '';

      if (result.config.Customers && result.config.Customers.length > 0) {
        result.config.Customers.forEach(customer => {
          const row = document.createElement('tr');

          // Create table cells
          const td1 = document.createElement('td');
          const td2 = document.createElement('td');
          const td3 = document.createElement('td');
          const td4 = document.createElement('td');

          const input1 = document.createElement('input');
          input1.type = 'text';
          input1.className = 'form-input form-input-compact customer-prefix';
          input1.value = customer.Prefix || '';
          input1.placeholder = 'Project prefix';

          const input2 = document.createElement('input');
          input2.type = 'text';
          input2.className = 'form-input form-input-compact customer-account';
          input2.value = customer.AccountName || '';
          input2.placeholder = 'account@domain.onmicrosoft.com';

          const input3 = document.createElement('input');
          input3.type = 'text';
          input3.className = 'form-input form-input-compact customer-domain';
          input3.value = customer.Domain || '';
          input3.placeholder = 'e.g. mbufara';

          const input4 = document.createElement('input');
          input4.type = 'text';
          input4.className = 'form-input form-input-compact customer-spo';
          input4.value = customer.SharePointAdminUrl || '';
          input4.placeholder = 'https://tenant-admin.sharepoint.com';

          td1.appendChild(input1);
          td2.appendChild(input2);
          td3.appendChild(input3);
          td4.appendChild(input4);

          row.appendChild(td1);
          row.appendChild(td2);
          row.appendChild(td3);
          row.appendChild(td4);

          customerTableBody.appendChild(row);
        });
      } else {
        // Add one empty row
        addEmptyCustomerRow();
      }
    }

    document.getElementById('settingsDialog').classList.remove('hidden');
  } catch (error) {
    console.error('Error opening settings:', error);
    document.getElementById('settingsDialog').classList.remove('hidden');
  }
}

// Close settings dialog
function closeSettings() {
  document.getElementById('settingsDialog').classList.add('hidden');
}

// Save settings
async function saveSettings() {
  try {
    // Get existing config first
    const existingResult = await window.electronAPI.getConfig();
    const existingConfig = (existingResult.success && existingResult.config) ? existingResult.config : {};

    const apiUrl = document.getElementById('apiUrl').value.trim();
    const clientId = document.getElementById('clientId').value.trim();
    const clientSecret = document.getElementById('clientSecret').value.trim();
    const portalUrl = document.getElementById('portalUrl').value.trim();
    const sharePointAdminUrl = document.getElementById('sharePointAdminUrl').value.trim();
    const secretExpiry = document.getElementById('secretExpiry').value.trim();
    const discoveryOutputPath = document.getElementById('discoveryOutputPath').value.trim();
    const vbuCsvPath = document.getElementById('vbuCsvPath').value.trim();

    // Collect customer data from table
    const customerRows = document.querySelectorAll('#customerTableBody tr');
    const customers = [];

    customerRows.forEach(row => {
      const prefix = row.querySelector('.customer-prefix').value.trim();
      const accountName = row.querySelector('.customer-account').value.trim();
      const domain = row.querySelector('.customer-domain')?.value.trim() || '';
      const spoUrl = row.querySelector('.customer-spo').value.trim();

      if (prefix) {
        customers.push({
          Prefix: prefix,
          AccountName: accountName,
          Domain: domain,
          SharePointAdminUrl: spoUrl
        });
      }
    });

    // Merge with existing config to preserve all fields
    const config = {
      ...existingConfig,  // Start with existing config
      Url: apiUrl,
      ClientId: clientId,
      PortalUrl: portalUrl,
      SharePointAdminUrl: sharePointAdminUrl,
      SecretExpiry: secretExpiry,
      DiscoveryOutputPath: discoveryOutputPath,
      VbuCsvPath: vbuCsvPath,
      Customers: customers
    };

    // Only update secret if provided (will be encrypted by PowerShell)
    if (clientSecret) {
      config.ClientSecret = clientSecret;
    }

    console.log('Saving config:', config);

    const result = await window.electronAPI.saveConfig(config);

    if (result.success) {
      alert('Settings saved successfully');
      closeSettings();

      // Reload customer data in all views that cache it
      if (typeof loadCustomerDomains === 'function') loadCustomerDomains();
      if (typeof loadMonitorProjects === 'function') loadMonitorProjects();
      if (typeof loadConnectionsCustomers === 'function') loadConnectionsCustomers();
    } else {
      alert(`Failed to save settings:\n${result.error}`);
    }
  } catch (error) {
    console.error('Error saving settings:', error);
    alert(`Error:\n${error.message}`);
  }
}

  // Test connection to Fly API
  async function testConnection() {
    const statusSpan = document.getElementById('connectionStatus');
    const testBtn = document.getElementById('testConnectionBtn');

    try {
      testBtn.disabled = true;
      statusSpan.textContent = 'Testing...';
      statusSpan.style.color = '#6c757d';

      const apiUrl = document.getElementById('apiUrl').value.trim();
      const clientId = document.getElementById('clientId').value.trim();
      const clientSecret = document.getElementById('clientSecret').value.trim();

      if (!apiUrl || !clientId) {
        statusSpan.textContent = '❌ Please fill in API URL and Client ID';
        statusSpan.style.color = '#dc3545';
        testBtn.disabled = false;
        return;
      }

      // Check if we have a saved secret or a new one
      const config = await window.electronAPI.getConfig();
      const hasExistingSecret = config.success && config.config && config.config.EncSecret;

      if (!clientSecret && !hasExistingSecret) {
        statusSpan.textContent = '❌ Please enter Client Secret';
        statusSpan.style.color = '#dc3545';
        testBtn.disabled = false;
        return;
      }

      // If a new secret is provided, save credentials first
      if (clientSecret) {
        const tempConfig = {
          Url: apiUrl,
          ClientId: clientId,
          ClientSecret: clientSecret
        };
        const saveResult = await window.electronAPI.saveConfig(tempConfig);
        if (!saveResult.success) {
          statusSpan.textContent = `❌ Failed to save credentials: ${saveResult.error}`;
          statusSpan.style.color = '#dc3545';
          testBtn.disabled = false;
          return;
        }
      } else {
        // Using existing secret, but need to save URL and ClientId in case they changed
        const tempConfig = {
          Url: apiUrl,
          ClientId: clientId
        };
        await window.electronAPI.saveConfig(tempConfig);
      }

      // Test the connection
      const result = await window.electronAPI.testConnection();

      if (result.success) {
        statusSpan.textContent = '✓ Connection successful';
        statusSpan.style.color = '#28a745';
      } else {
        statusSpan.textContent = `❌ ${result.error || 'Connection failed'}`;
        statusSpan.style.color = '#dc3545';
      }
    } catch (error) {
      statusSpan.textContent = `❌ Error: ${error.message}`;
      statusSpan.style.color = '#dc3545';
    } finally {
      testBtn.disabled = false;
    }
  }

// Wire up item detail modal close
document.addEventListener('DOMContentLoaded', () => {
  const itemOverlay = document.getElementById('itemDetailOverlay');
  document.getElementById('itemDetailClose')?.addEventListener('click', () => {
    itemOverlay?.classList.add('hidden');
  });
  itemOverlay?.addEventListener('click', (e) => {
    if (e.target === itemOverlay) itemOverlay.classList.add('hidden');
  });
});

// PowerShell Automation - needs to be called during DOMContentLoaded
// Initialize after DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  const psCustomerPrefix    = document.getElementById('psCustomerPrefix');
  const psCreateProjectBtn  = document.getElementById('psCreateProjectBtn');
  const psImportMappingsBtn = document.getElementById('psImportMappingsBtn');
  const psCreateSitesBtn    = document.getElementById('psCreateSitesBtn');
  const psVerifyBtn         = document.getElementById('psVerifyBtn');
  const psPreScanBtn        = document.getElementById('psPreScanBtn');
  const psFullMigrBtn       = document.getElementById('psFullMigrBtn');
  const psIncrMigrBtn       = document.getElementById('psIncrMigrBtn');
  const psStartWorkflowBtn  = document.getElementById('psStartWorkflowBtn');
  const psStopJobsBtn       = document.getElementById('psStopJobsBtn');
  const psClearMappingsBtn  = document.getElementById('psClearMappingsBtn');
  const psViewDocsBtn       = document.getElementById('psViewDocsBtn');
  const connMappingsLog     = document.getElementById('connMappingsLog');
  const connMappingsLogPre  = document.getElementById('connMappingsLogPre');

  function appendConnLog(text) {
    if (!connMappingsLogPre) return;
    connMappingsLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    connMappingsLogPre.scrollTop = connMappingsLogPre.scrollHeight;
  }

  function showConnLog(clearFirst) {
    if (connMappingsLog) connMappingsLog.style.display = '';
    if (clearFirst && connMappingsLogPre) connMappingsLogPre.textContent = '';
    fitConnectionsLog();
  }

  const connWorkloadDefs = [
    { id: 'exchangeMapping',   workload: 'Exchange'   },
    { id: 'onedriveMapping',   workload: 'OneDrive'   },
    { id: 'sharepointMapping', workload: 'SharePoint' },
    { id: 'teamsMapping',      workload: 'Teams'      },
    { id: 'teamschatsMapping', workload: 'TeamChat'   },
    { id: 'groupsMapping',     workload: 'Groups'     },
  ];

  function getFilledWorkloads() {
    return connWorkloadDefs
      .map(w => ({ ...w, file: document.getElementById(w.id)?.value?.trim() }))
      .filter(w => w.file);
  }

  function getSelectedWorkloads() {
    return connWorkloadDefs
      .filter(w => document.getElementById(`wlChk-${w.workload}`)?.checked)
      .map(w => w.workload);
  }

  if (psCreateProjectBtn) {
    psCreateProjectBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix?.value?.trim();
      if (!prefix) { alert('Please select a customer first.'); return; }

      const cfgResult = await window.electronAPI.getConfig();
      const cfg = cfgResult.success ? cfgResult.config : {};

      const sourceConnMap = {
        Exchange:   'OurVolaris - EXO',
        OneDrive:   'OurVolaris - OneDrive',
        SharePoint: 'OurVolaris - SPO',
        Teams:      'OurVolaris - MS Teams',
        TeamChat:   'OurVolaris - Teams Chats',
        Groups:     'OurVolaris - M365 Groups',
      };

      const customer = (cfg.Customers || []).find(c => c.Prefix === prefix);
      const customerDomain = customer?.Domain?.trim() || '';
      if (!customerDomain) {
        alert(`No domain found for customer "${prefix}".\n\nAdd the customer's domain in Settings → Customer.`);
        return;
      }

      psCreateProjectBtn.disabled = true;
      psCreateProjectBtn.textContent = 'Creating...';
      showConnLog(true);
      appendConnLog(`=== Creating Projects for ${prefix} ===\n\n`);

      window.electronAPI.onPsOutput(appendConnLog);
      try {
        for (const item of connWorkloadDefs) {
          const projectName = `${prefix} - ${item.workload}`;
          appendConnLog(`\n--- ${projectName} ---\n`);
          const result = await window.electronAPI.streamPowerShell('New-FlyProject.ps1', [
            '-ProjectName',      projectName,
            '-Workload',         item.workload,
            '-SourceConnection', sourceConnMap[item.workload],
            '-CustomerDomain',   customerDomain
          ]);
          appendConnLog(result.success ? `\n✓ Done\n` : `\n✗ Failed (exit ${result.code})\n`);
        }
        appendConnLog('\n=== Finished ===\n');
      } catch (err) {
        appendConnLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        psCreateProjectBtn.disabled = false;
        psCreateProjectBtn.textContent = '📁 Create Projects';
      }
    });
  }

  if (psImportMappingsBtn) {
    psImportMappingsBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix?.value?.trim();
      if (!prefix) { alert('Please select a customer first.'); return; }

      const toImport = getFilledWorkloads();
      if (toImport.length === 0) { alert('Please browse for at least one mapping file.'); return; }

      psImportMappingsBtn.disabled = true;
      psImportMappingsBtn.textContent = 'Importing...';
      showConnLog(true);
      appendConnLog(`=== Importing Mappings for ${prefix} ===\n\n`);

      window.electronAPI.onPsOutput(appendConnLog);
      try {
        for (const item of toImport) {
          const projectName = `${prefix} - ${item.workload}`;
          appendConnLog(`\n--- ${item.workload} ---\n`);
          psImportMappingsBtn.textContent = `Importing ${item.workload}…`;
          const result = await window.electronAPI.streamPowerShell('Import-FlyMappings.ps1', [
            '-ProjectName', projectName,
            '-Workload',    item.workload,
            '-MappingFile', item.file
          ]);
          appendConnLog(result.success ? `\n✓ Done\n` : `\n✗ Failed (exit ${result.code})\n`);
        }
        appendConnLog('\n=== Finished ===\n');
      } catch (err) {
        appendConnLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        psImportMappingsBtn.disabled = false;
        psImportMappingsBtn.textContent = '📥 Import Mappings';
      }
    });
  }

  if (psCreateSitesBtn) {
    psCreateSitesBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix?.value?.trim();
      if (!prefix) { alert('Please select a customer first.'); return; }

      const spFile = document.getElementById('sharepointMapping')?.value?.trim();
      if (!spFile) {
        alert('Please browse for the SharePoint mapping CSV file first.');
        return;
      }

      const cfgResult = await window.electronAPI.getConfig();
      const cfg = cfgResult.success ? cfgResult.config : {};
      const customer = (cfg.Customers || []).find(c => c.Prefix === prefix);
      const ownerEmail = customer?.AccountName?.trim();
      if (!ownerEmail) {
        alert(`No AccountName found for customer "${prefix}".\n\nAdd the destination admin email in Settings → Customer.`);
        return;
      }

      psCreateSitesBtn.disabled = true;
      psCreateSitesBtn.textContent = 'Creating Sites...';
      showConnLog(true);
      appendConnLog(`=== Create SharePoint Sites — ${prefix} ===\n`);
      appendConnLog(`Site owner: ${ownerEmail}\n\n`);

      window.electronAPI.onPsOutput(appendConnLog);
      try {
        const result = await window.electronAPI.streamPowerShell('New-SharePointSites.ps1', [
          '-MappingFile', spFile,
          '-SiteOwner',   ownerEmail
        ]);
        appendConnLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
        appendConnLog('\n=== Finished ===\n');
      } catch (err) {
        appendConnLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        psCreateSitesBtn.disabled = false;
        psCreateSitesBtn.textContent = '🌐 Create Sites';
      }
    });
  }

  async function runMigrationStage(stage, label, btn, originalLabel) {
    const prefix = psCustomerPrefix?.value?.trim();
    if (!prefix) { alert('Please select a customer first.'); return; }

    btn.disabled = true;
    btn.textContent = `${label}…`;
    showConnLog(true);
    appendConnLog(`=== ${label} — ${prefix} ===\n\n`);

    window.electronAPI.onPsOutput(appendConnLog);
    try {
      const args = ['-CustomerPrefix', prefix, '-Stage', stage];
      // Only pass -Workloads when a non-empty subset is explicitly ticked;
      // if nothing (or everything) is ticked, let the PS script run all workloads
      // and skip any projects that don't exist automatically.
      const selected = getSelectedWorkloads();
      if (selected.length > 0 && selected.length < connWorkloadDefs.length) {
        args.push('-Workloads', selected.join(','));
      }
      const result = await window.electronAPI.streamPowerShell('Start-FlyMigrationStage.ps1', args);
      appendConnLog(result.success ? `\n✓ Done\n` : `\n✗ Failed (exit ${result.code})\n`);
      appendConnLog('\n=== Finished ===\n');
    } catch (err) {
      appendConnLog(`\nError: ${err.message || err}\n`);
    } finally {
      window.electronAPI.offPsOutput();
      btn.disabled = false;
      btn.textContent = originalLabel;
    }
  }

  if (psVerifyBtn) {
    psVerifyBtn.addEventListener('click', () =>
      runMigrationStage('Verify', 'Verify', psVerifyBtn, '✓ Verify'));
  }

  if (psPreScanBtn) {
    psPreScanBtn.addEventListener('click', () =>
      runMigrationStage('PreScan', 'Pre-Scan', psPreScanBtn, '🔍 Pre-Scan'));
  }

  if (psFullMigrBtn) {
    psFullMigrBtn.addEventListener('click', () =>
      runMigrationStage('FullMigration', 'Full Migration', psFullMigrBtn, '▶ Full'));
  }

  if (psIncrMigrBtn) {
    psIncrMigrBtn.addEventListener('click', () =>
      runMigrationStage('IncrementalMigration', 'Incremental Migration', psIncrMigrBtn, '↺ Incremental'));
  }

  if (psStopJobsBtn) {
    psStopJobsBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix?.value?.trim();
      if (!prefix) { alert('Please select a customer first.'); return; }
      if (!confirm(`Stop all in-progress jobs for "${prefix}"?`)) return;

      psStopJobsBtn.disabled = true;
      psStopJobsBtn.textContent = 'Stopping…';
      showConnLog(true);
      appendConnLog(`=== Stop Jobs — ${prefix} ===\n\n`);

      window.electronAPI.onPsOutput(appendConnLog);
      try {
        const args = ['-CustomerPrefix', prefix];
        const selected = getSelectedWorkloads();
        if (selected.length > 0 && selected.length < connWorkloadDefs.length) {
          args.push('-Workloads', selected.join(','));
        }
        const result = await window.electronAPI.streamPowerShell('Stop-FlyMigrationStage.ps1', args);
        appendConnLog(result.success ? `\n✓ Done\n` : `\n✗ Failed (exit ${result.code})\n`);
        appendConnLog('\n=== Finished ===\n');
      } catch (err) {
        appendConnLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        psStopJobsBtn.disabled = false;
        psStopJobsBtn.textContent = '⏹ Stop Jobs';
      }
    });
  }

  if (psStartWorkflowBtn) {
    psStartWorkflowBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix?.value?.trim();
      if (!prefix) { alert('Please select a customer first.'); return; }

      const toRun = getFilledWorkloads();
      if (toRun.length === 0) { alert('Please browse for at least one mapping file.'); return; }

      const cfgResult2 = await window.electronAPI.getConfig();
      const cfg2 = cfgResult2.success ? cfgResult2.config : {};
      const wfCustomer = (cfg2.Customers || []).find(c => c.Prefix === prefix);
      const wfDomain = wfCustomer?.Domain?.trim() || '';

      psStartWorkflowBtn.disabled = true;
      psStartWorkflowBtn.textContent = 'Running...';
      showConnLog(true);
      appendConnLog(`=== Full Migration Workflow for ${prefix} ===\n\n`);

      window.electronAPI.onPsOutput(appendConnLog);
      try {
        for (const item of toRun) {
          appendConnLog(`\n--- ${item.workload} ---\n`);
          psStartWorkflowBtn.textContent = `Running ${item.workload}…`;
          const wfArgs = [
            '-CustomerPrefix', prefix,
            '-Workload',       item.workload,
            '-MappingFile',    item.file
          ];
          if (wfDomain) { wfArgs.push('-CustomerDomain', wfDomain); }
          const result = await window.electronAPI.streamPowerShell('Start-FlyMigrationWorkflow.ps1', wfArgs);
          appendConnLog(result.success ? `\n✓ Done\n` : `\n✗ Failed (exit ${result.code})\n`);
        }
        appendConnLog('\n=== Finished ===\n');
      } catch (err) {
        appendConnLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        psStartWorkflowBtn.disabled = false;
        psStartWorkflowBtn.textContent = '🚀 Full Workflow';
      }
    });
  }

  if (psClearMappingsBtn) {
    psClearMappingsBtn.addEventListener('click', () => {
      connWorkloadDefs.forEach(w => {
        const el = document.getElementById(w.id);
        if (el) el.value = '';
      });
      if (connMappingsLogPre) connMappingsLogPre.textContent = '';
      if (connMappingsLog)    connMappingsLog.style.display = 'none';
    });
  }

  if (psViewDocsBtn) {
    psViewDocsBtn.addEventListener('click', async () => {
      try {
        await window.electronAPI.openExternal('https://docs.avepoint.com/fly/');
      } catch (error) {
        alert(`Error opening documentation: ${error.message}`);
      }
    });
  }

  // App Registration handlers
  const appRegTenantId = document.getElementById('appRegTenantId');
  const appRegName = document.getElementById('appRegName');
  const createAppRegBtn = document.getElementById('createAppRegBtn');
  const viewAppRegDocsBtn = document.getElementById('viewAppRegDocsBtn');
  const appRegStatus = document.getElementById('appRegStatus');

  const appRegLog    = document.getElementById('appRegLog');
  const appRegLogPre = document.getElementById('appRegLogPre');

  function appendAppRegLog(text) {
    if (!appRegLogPre) return;
    // Strip ANSI colour codes that PowerShell emits
    appRegLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    appRegLogPre.scrollTop = appRegLogPre.scrollHeight;
  }

  if (createAppRegBtn) {
    createAppRegBtn.addEventListener('click', async () => {
      const tenantId = appRegTenantId.value.trim();
      const appName = appRegName.value.trim();

      if (!tenantId) {
        alert('Please enter the Destination Tenant ID');
        return;
      }

      if (!appName) {
        alert('Please enter an Application Name');
        return;
      }

      const confirmed = confirm(
        `Create Azure AD App Registration?\n\n` +
        `Tenant ID: ${tenantId}\n` +
        `App Name: ${appName}\n\n` +
        `This will:\n` +
        `1. Show a device code in PowerShell window\n` +
        `2. You visit https://microsoft.com/devicelogin\n` +
        `3. Enter the code to authenticate\n` +
        `4. Create app in Azure AD\n` +
        `5. Configure Graph & SharePoint permissions\n` +
        `6. Generate Client ID & Secret\n\n` +
        `You must be Global Admin or Application Admin.\n\n` +
        `Continue?`
      );
      if (!confirmed) return;

      createAppRegBtn.disabled = true;
      createAppRegBtn.textContent = 'Creating...';

      if (appRegLog)    { appRegLog.style.display = 'block'; }
      if (appRegLogPre) { appRegLogPre.textContent = ''; }
      appRegStatus.style.display = 'block';
      appRegStatus.style.color = '#6c757d';
      appRegStatus.textContent = '⏳ Running — device code will appear in the log below. Copy it and visit https://microsoft.com/devicelogin';

      appendAppRegLog('Starting app registration...\n\n');

      // Stream output line-by-line so the device code appears immediately
      window.electronAPI.onPsOutput((data) => appendAppRegLog(data));

      try {
        const result = await window.electronAPI.streamPowerShell('New-AzureAppRegistration.ps1', [
          '-TenantId', tenantId,
          '-AppName', appName,
          '-SkipSavePrompt'
        ]);

        if (result.success) {
          appendAppRegLog('\n✓ Completed successfully.\n');
          appRegStatus.style.color = '#28a745';
          appRegStatus.textContent = '✓ App registration created — see log above for credentials.';
        } else {
          appendAppRegLog('\n✗ Failed (exit code ' + result.code + ').\n');
          appRegStatus.style.color = '#dc3545';
          appRegStatus.textContent = '✗ Failed — see log above for details.';
        }
      } catch (error) {
        appendAppRegLog('\n✗ Error: ' + (error.message || error) + '\n');
        appRegStatus.style.color = '#dc3545';
        appRegStatus.textContent = `✗ Error: ${error.message || error}`;
      } finally {
        window.electronAPI.offPsOutput();
        createAppRegBtn.disabled = false;
        createAppRegBtn.textContent = '🚀 Create App Registration';
      }
    });
  }

  if (viewAppRegDocsBtn) {
    viewAppRegDocsBtn.addEventListener('click', async () => {
      try {
        await window.electronAPI.openExternal('https://docs.avepoint.com/fly/user_guide/migration/get_started_with_fly_migration/create_an_azure_ad_application.htm');
      } catch (error) {
        alert(`Error opening documentation: ${error.message}`);
      }
    });
  }

  // AOS Setup handlers
  const aosDisplayNameInput = document.getElementById('aosDisplayName');
  if (aosDisplayNameInput) {
    aosDisplayNameInput.addEventListener('input', () => {
      const pn = document.getElementById('aosProfileName');
      if (pn && (!pn.value || pn.value.endsWith(' App'))) {
        pn.value = aosDisplayNameInput.value.trim() ? aosDisplayNameInput.value.trim() + ' App' : '';
      }
    });
  }

  const aosLog    = document.getElementById('aosLog');
  const aosLogPre = document.getElementById('aosLogPre');

  function appendAosLog(text) {
    if (!aosLogPre) return;
    aosLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    aosLogPre.scrollTop = aosLogPre.scrollHeight;
  }

  function showAosStatus(msg, type) {
    const el = document.getElementById('aosStatus');
    if (!el) return;
    el.style.display = 'block';
    el.style.color = type === 'error' ? '#dc3545' : type === 'warning' ? '#856404' : '#0064b4';
    el.textContent = msg;
  }

  function aosGetFields() {
    return {
      displayName: (document.getElementById('aosDisplayName') || {}).value?.trim() || '',
      searchCode:  (document.getElementById('aosSearchCode')  || {}).value?.trim() || '',
      profileName: (document.getElementById('aosProfileName') || {}).value?.trim() || ''
    };
  }

  const aosSignInBtn = document.getElementById('aosSignInBtn');
  if (aosSignInBtn) {
    aosSignInBtn.addEventListener('click', async () => {
      const { displayName, searchCode, profileName } = aosGetFields();
      if (!displayName || !searchCode) {
        showAosStatus('Enter Display Name and Search Code first.', 'warning');
        return;
      }
      aosSignInBtn.disabled = true;
      aosSignInBtn.textContent = 'Signing in...';
      if (aosLog)    { aosLog.style.display = 'block'; }
      if (aosLogPre) { aosLogPre.textContent = ''; }
      showAosStatus('Opening browser for AOS sign-in...', 'info');
      appendAosLog('Starting sign-in...\n\n');
      window.electronAPI.onPsOutput((data) => appendAosLog(data));
      try {
        await window.electronAPI.saveSharedConfig({ TenantName: displayName, TenantSearch: searchCode, AppProfileName: profileName });
        const result = await window.electronAPI.streamPowerShell('Aos-SignIn.ps1');
        if (result.success) {
          appendAosLog('\n✓ Sign-in complete.\n');
          showAosStatus('✓ Signed in. Session saved.', 'info');
        } else {
          appendAosLog('\n✗ Sign-in failed (exit code ' + result.code + ').\n');
          showAosStatus('✗ Sign-in failed — see log above.', 'error');
        }
      } catch (error) {
        appendAosLog('\n✗ Error: ' + (error.message || error) + '\n');
        showAosStatus('Error: ' + (error.message || error), 'error');
      } finally {
        window.electronAPI.offPsOutput();
        aosSignInBtn.disabled = false;
        aosSignInBtn.textContent = '🔐 Sign in to AOS';
      }
    });
  }

  const aosRunSetupBtn = document.getElementById('aosRunSetupBtn');
  if (aosRunSetupBtn) {
    aosRunSetupBtn.addEventListener('click', async () => {
      const { displayName, searchCode, profileName } = aosGetFields();
      if (!displayName || !searchCode || !profileName) {
        showAosStatus('Please fill in all three fields before running setup.', 'warning');
        return;
      }
      aosRunSetupBtn.disabled = true;
      aosRunSetupBtn.textContent = 'Running...';
      if (aosLog)    { aosLog.style.display = 'block'; }
      if (aosLogPre) { aosLogPre.textContent = ''; }
      showAosStatus('Browser automation running — approve any consent prompts that appear.', 'info');
      appendAosLog('Starting app profile setup...\n\n');
      window.electronAPI.onPsOutput((data) => appendAosLog(data));
      try {
        await window.electronAPI.saveSharedConfig({ TenantName: displayName, TenantSearch: searchCode, AppProfileName: profileName });
        const result = await window.electronAPI.streamPowerShell('Aos-Setup.ps1');
        if (result.success) {
          appendAosLog('\n✓ Setup complete.\n');
          showAosStatus('✓ App profile created and consent granted.', 'info');
        } else {
          appendAosLog('\n✗ Setup failed (exit code ' + result.code + ').\n');
          showAosStatus('✗ Setup failed — see log above.', 'error');
        }
      } catch (error) {
        appendAosLog('\n✗ Error: ' + (error.message || error) + '\n');
        showAosStatus('Error: ' + (error.message || error), 'error');
      } finally {
        window.electronAPI.offPsOutput();
        aosRunSetupBtn.disabled = false;
        aosRunSetupBtn.textContent = '🚀 Create App Profile & Grant Consent';
      }
    });
  }

  // --- dead code kept for reference, remove this block if aosTestConnectionBtn no longer exists ---
  const aosTestConnectionBtn = document.getElementById('aosTestConnectionBtn');
  const aosConnectionStatus  = document.getElementById('aosConnectionStatus');
  if (aosTestConnectionBtn) {
    aosTestConnectionBtn.addEventListener('click', async () => {
      aosTestConnectionBtn.disabled = true;
      aosTestConnectionBtn.textContent = 'Testing...';
      aosConnectionStatus.style.display = 'block';
      aosConnectionStatus.style.color = '#6c757d';
      aosConnectionStatus.textContent = '⏳ Testing connection...';

      try {
        const result = await window.electronAPI.testConnection();

        if (result.success) {
          aosConnectionStatus.style.color = '#28a745';
          aosConnectionStatus.textContent = '✓ Connection successful';
        } else {
          aosConnectionStatus.style.color = '#dc3545';
          aosConnectionStatus.textContent = `❌ ${result.error || 'Connection failed'}`;
        }
      } catch (error) {
        aosConnectionStatus.style.color = '#dc3545';
        aosConnectionStatus.textContent = `❌ Error: ${error.message}`;
      } finally {
        aosTestConnectionBtn.disabled = false;
        aosTestConnectionBtn.textContent = '🔌 Test Connection';
      }
    });
  }

});

// ── Misc Scripts handlers ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {

  // ── Provision OneDrives ───────────────────────────────────────────────────
  const provisionOneDriveBtn      = document.getElementById('provisionOneDriveBtn');
  const checkOneDriveStatusBtn    = document.getElementById('checkOneDriveStatusBtn');
  const onedriveTenantUrl         = document.getElementById('onedriveTenantUrl');
  const onedriveAdHocUrl          = document.getElementById('onedriveAdHocUrl');
  const onedriveMappingFile       = document.getElementById('onedriveMappingFile');
  const onedriveColumnOverride    = document.getElementById('onedriveColumnOverride');
  const onedriveWhatIf            = document.getElementById('onedriveWhatIf');
  const onedriveExportClean       = document.getElementById('onedriveExportClean');
  const onedriveExportCleanBrowse = document.getElementById('onedriveExportCleanBrowse');
  const provisionOneDriveLog      = document.getElementById('provisionOneDriveLog');
  const provisionOneDriveLogPre   = document.getElementById('provisionOneDriveLogPre');

  function appendProvLog(text) {
    if (!provisionOneDriveLogPre) return;
    provisionOneDriveLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    provisionOneDriveLogPre.scrollTop = provisionOneDriveLogPre.scrollHeight;
  }

  if (provisionOneDriveBtn) {
    provisionOneDriveBtn.addEventListener('click', async () => {
      const adminUrl = onedriveAdHocUrl?.value?.trim() || onedriveTenantUrl?.value?.trim();
      const mapFile  = onedriveMappingFile?.value?.trim();
      if (!adminUrl) { alert('Please select an SPO Admin URL or enter a one-off URL.'); return; }
      if (!mapFile)  { alert('Please browse for a mapping file.'); return; }

      provisionOneDriveBtn.disabled = true;
      provisionOneDriveBtn.textContent = 'Running…';
      if (provisionOneDriveLog) provisionOneDriveLog.style.display = '';
      if (provisionOneDriveLogPre) provisionOneDriveLogPre.textContent = '';

      const args = ['-MappingFile', mapFile, '-AdminUrl', adminUrl];
      const col  = onedriveColumnOverride?.value?.trim();
      if (col) args.push('-Column', col);
      if (onedriveWhatIf?.checked) args.push('-WhatIf');

      window.electronAPI.onPsOutput(appendProvLog);
      try {
        const result = await window.electronAPI.streamPowerShell('Provision-OneDrives.ps1', args);
        appendProvLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
      } catch (err) {
        appendProvLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        provisionOneDriveBtn.disabled = false;
        provisionOneDriveBtn.textContent = '▶ Start Provisioning';
      }
    });
  }

  if (onedriveExportCleanBrowse) {
    onedriveExportCleanBrowse.addEventListener('click', async () => {
      const mapFile = onedriveMappingFile?.value?.trim() || '';
      const defaultName = mapFile
        ? mapFile.replace(/(\.[^.]+)$/, '_clean$1')
        : 'mapping_clean.csv';
      const result = await window.electronAPI.showSaveDialog({
        defaultPath: defaultName,
        filters: [{ name: 'CSV Files', extensions: ['csv'] }, { name: 'All Files', extensions: ['*'] }]
      });
      if (!result.canceled && result.filePath) {
        onedriveExportClean.value = result.filePath;
      }
    });
  }

  if (checkOneDriveStatusBtn) {
    checkOneDriveStatusBtn.addEventListener('click', async () => {
      const adminUrl = onedriveAdHocUrl?.value?.trim() || onedriveTenantUrl?.value?.trim();
      const mapFile  = onedriveMappingFile?.value?.trim();
      if (!adminUrl) { alert('Please select an SPO Admin URL or enter a one-off URL.'); return; }
      if (!mapFile)  { alert('Please browse for a mapping file.'); return; }

      checkOneDriveStatusBtn.disabled = true;
      provisionOneDriveBtn.disabled   = true;
      checkOneDriveStatusBtn.textContent = 'Checking…';
      if (provisionOneDriveLog) provisionOneDriveLog.style.display = '';
      if (provisionOneDriveLogPre) provisionOneDriveLogPre.textContent = '';

      const args = ['-MappingFile', mapFile, '-AdminUrl', adminUrl];
      const col  = onedriveColumnOverride?.value?.trim();
      if (col) args.push('-Column', col);
      const exportPath = onedriveExportClean?.value?.trim();
      if (exportPath) args.push('-ExportCleanCsv', exportPath);

      window.electronAPI.onPsOutput(appendProvLog);
      try {
        const result = await window.electronAPI.streamPowerShell('Check-OneDriveStatus.ps1', args);
        appendProvLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
      } catch (err) {
        appendProvLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        checkOneDriveStatusBtn.disabled = false;
        provisionOneDriveBtn.disabled   = false;
        checkOneDriveStatusBtn.textContent = 'Check Status';
      }
    });
  }

  // ── Deduplicate Inventory ─────────────────────────────────────────────────
  const deduplicateRunBtn   = document.getElementById('deduplicateRunBtn');
  const deduplicateWorkbook = document.getElementById('deduplicateWorkbook');
  const deduplicateLogPre   = document.getElementById('deduplicateLogPre');

  function appendDeduplicateLog(text) {
    if (!deduplicateLogPre) return;
    deduplicateLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    deduplicateLogPre.scrollTop = deduplicateLogPre.scrollHeight;
  }

  if (deduplicateRunBtn) {
    deduplicateRunBtn.addEventListener('click', async () => {
      const workbook = deduplicateWorkbook?.value?.trim();
      if (!workbook) { alert('Please browse for a discovery workbook.'); return; }

      deduplicateRunBtn.disabled = true;
      deduplicateRunBtn.textContent = 'Running…';
      if (deduplicateLogPre) deduplicateLogPre.textContent = '';

      const args = ['-SourceWorkbook', workbook];

      window.electronAPI.onPsOutput(appendDeduplicateLog);
      try {
        const result = await window.electronAPI.streamPowerShell('Deduplicate-Inventory.ps1', args);
        appendDeduplicateLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
      } catch (err) {
        appendDeduplicateLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        deduplicateRunBtn.disabled = false;
        deduplicateRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Purge Deleted SPO Sites ───────────────────────────────────────────────
  const purgeSpoRunBtn     = document.getElementById('purgeSpoRunBtn');
  const purgeSpoMappingFile = document.getElementById('purgeSpoMappingFile');
  const purgeSpoLogPre     = document.getElementById('purgeSpoLogPre');

  if (purgeSpoRunBtn) {
    purgeSpoRunBtn.addEventListener('click', async () => {
      const mappingFile = purgeSpoMappingFile?.value?.trim();
      if (!mappingFile) { alert('Please browse for the SharePoint mapping CSV file.'); return; }

      purgeSpoRunBtn.disabled = true;
      purgeSpoRunBtn.textContent = 'Purging…';
      if (purgeSpoLogPre) purgeSpoLogPre.textContent = '';

      const appendLog = text => {
        if (!purgeSpoLogPre) return;
        purgeSpoLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
        purgeSpoLogPre.scrollTop = purgeSpoLogPre.scrollHeight;
      };

      window.electronAPI.onPsOutput(appendLog);
      try {
        const result = await window.electronAPI.streamPowerShell('Remove-DeletedSharePointSites.ps1', [
          '-MappingFile', mappingFile
        ]);
        appendLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
      } catch (err) {
        appendLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        purgeSpoRunBtn.disabled = false;
        purgeSpoRunBtn.textContent = '🗑 Purge Deleted Sites';
      }
    });
  }

  // ── Set Teams Owners ──────────────────────────────────────────────────────
  const setTeamsOwnersBtn  = document.getElementById('setTeamsOwnersBtn');
  const teamsOwnerUpn      = document.getElementById('teamsOwnerUpn');
  const teamsCsvFile       = document.getElementById('teamsCsvFile');
  const teamsWhatIf        = document.getElementById('teamsWhatIf');
  const teamsOwnersLog     = document.getElementById('teamsOwnersLog');
  const teamsOwnersLogPre  = document.getElementById('teamsOwnersLogPre');

  function appendTeamsLog(text) {
    if (!teamsOwnersLogPre) return;
    teamsOwnersLogPre.textContent += text.replace(/\x1b\[[0-9;]*m/g, '');
    teamsOwnersLogPre.scrollTop = teamsOwnersLogPre.scrollHeight;
  }

  if (setTeamsOwnersBtn) {
    setTeamsOwnersBtn.addEventListener('click', async () => {
      const ownerUpn = teamsOwnerUpn?.value?.trim();
      const csvFile  = teamsCsvFile?.value?.trim();
      if (!ownerUpn) { alert('Please enter the owner UPN.'); return; }
      if (!csvFile)  { alert('Please browse for a CSV file.'); return; }

      setTeamsOwnersBtn.disabled = true;
      setTeamsOwnersBtn.textContent = 'Running…';
      if (teamsOwnersLog) teamsOwnersLog.style.display = '';
      if (teamsOwnersLogPre) teamsOwnersLogPre.textContent = '';

      const args = ['-CsvFile', csvFile, '-OwnerUpn', ownerUpn];
      if (teamsWhatIf?.checked) args.push('-WhatIf');

      window.electronAPI.onPsOutput(appendTeamsLog);
      try {
        const result = await window.electronAPI.streamPowerShell('Set-TeamsOwners-Run.ps1', args);
        appendTeamsLog(result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`);
      } catch (err) {
        appendTeamsLog(`\nError: ${err.message || err}\n`);
      } finally {
        window.electronAPI.offPsOutput();
        setTeamsOwnersBtn.disabled = false;
        setTeamsOwnersBtn.textContent = '▶ Add Owner';
      }
    });
  }

  // ── Get Domain Devices ────────────────────────────────────────────────────
  const getDomainDevicesBtn = document.getElementById('getDomainDevicesBtn');
  if (getDomainDevicesBtn) {
    getDomainDevicesBtn.addEventListener('click', async () => {
      await launchScript('Get-DomainDevices.ps1', getDomainDevicesBtn);
    });
  }

  // ── Reports → redirect to Monitor ────────────────────────────────────────
  const reportsGoToMonitorBtn = document.getElementById('reportsGoToMonitorBtn');
  if (reportsGoToMonitorBtn) {
    reportsGoToMonitorBtn.addEventListener('click', () => {
      switchView('avepoint-monitor');
    });
  }

  // ── Remove Domain — browse & run ──────────────────────────────────────────
  const removeBrowseBtn = document.getElementById('removeBrowseBtn');
  if (removeBrowseBtn) {
    removeBrowseBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.showOpenDialog({ properties: ['openDirectory'] });
      if (result && !result.canceled && result.filePaths.length > 0) {
        document.getElementById('removeDiscoveryFolder').value = result.filePaths[0];
      }
    });
  }

  const removeRunBtn = document.getElementById('removeRunBtn');
  if (removeRunBtn) {
    removeRunBtn.addEventListener('click', async () => {
      const folder  = document.getElementById('removeDiscoveryFolder').value.trim();
      const whatIf  = document.getElementById('removeWhatIf').checked;
      const checked = [...document.querySelectorAll('.remove-section:checked')].map(cb => cb.value);

      if (!folder) { alert('Please select a Discovery folder.'); return; }
      if (checked.length === 0) { alert('Please select at least one object type.'); return; }

      const logSection = document.getElementById('removeLog');
      const logOutput  = document.getElementById('removeLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      removeRunBtn.disabled = true;
      removeRunBtn.textContent = 'Running…';

      const args = ['-DiscoveryFolder', folder, '-Sections', checked.join(',')];
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('remove-domain.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        removeRunBtn.disabled = false;
        removeRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Update On-Prem UPNs ───────────────────────────────────────────────────
  const onpremCsvBrowseBtn = document.getElementById('onpremCsvBrowseBtn');
  if (onpremCsvBrowseBtn) {
    onpremCsvBrowseBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.showOpenDialog({ properties: ['openDirectory'] });
      if (!result.canceled && result.filePaths.length > 0) {
        document.getElementById('onpremCsvFolder').value = result.filePaths[0];
      }
    });
  }

  const onpremRunBtn = document.getElementById('onpremRunBtn');
  if (onpremRunBtn) {
    onpremRunBtn.addEventListener('click', async () => {
      const folder = document.getElementById('onpremCsvFolder').value.trim();
      const src    = document.getElementById('onpremSourceDomain').value.trim();
      const tgt    = document.getElementById('onpremTargetDomain').value.trim();
      const whatIf = document.getElementById('onpremWhatIf').checked;

      if (!folder) { alert('Please select a CSV folder.'); return; }
      if (!src || !tgt) { alert('Please enter both source and target domains.'); return; }

      const logSection = document.getElementById('onpremLog');
      const logOutput  = document.getElementById('onpremLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      onpremRunBtn.disabled = true;
      onpremRunBtn.textContent = 'Running…';

      const args = ['-CSVFolder', folder, '-SourceDomain', src, '-TargetDomain', tgt];
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Update-OnPremUPN.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        onpremRunBtn.disabled = false;
        onpremRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Update Cloud UPNs ─────────────────────────────────────────────────────
  const cloudRunBtn = document.getElementById('cloudRunBtn');
  if (cloudRunBtn) {
    cloudRunBtn.addEventListener('click', async () => {
      const oldDomain = document.getElementById('cloudOldDomain').value.trim();
      const newDomain = document.getElementById('cloudNewDomain').value.trim();
      const whatIf    = document.getElementById('cloudWhatIf').checked;

      if (!oldDomain || !newDomain) { alert('Please enter both old and new domains.'); return; }

      const logSection = document.getElementById('cloudLog');
      const logOutput  = document.getElementById('cloudLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      cloudRunBtn.disabled = true;
      cloudRunBtn.textContent = 'Running…';

      const args = ['-OldDomain', oldDomain, '-NewDomain', newDomain];
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Update-UPN.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        cloudRunBtn.disabled = false;
        cloudRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Hide from Address Book ────────────────────────────────────────────────
  const hideBrowseBtn = document.getElementById('hideBrowseBtn');
  if (hideBrowseBtn) {
    hideBrowseBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.showOpenDialog({ properties: ['openDirectory'] });
      if (!result.canceled && result.filePaths.length > 0) {
        document.getElementById('hideDiscoveryFolder').value = result.filePaths[0];
      }
    });
  }

  const hideRunBtn = document.getElementById('hideRunBtn');
  if (hideRunBtn) {
    hideRunBtn.addEventListener('click', async () => {
      const folder = document.getElementById('hideDiscoveryFolder').value.trim();
      const whatIf = document.getElementById('hideWhatIf').checked;

      if (!folder) { alert('Please select a discovery folder.'); return; }

      const logSection = document.getElementById('hideLog');
      const logOutput  = document.getElementById('hideLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      hideRunBtn.disabled = true;
      hideRunBtn.textContent = 'Running…';

      const args = ['-DiscoveryFolder', folder];
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Hide-AddressBook.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        hideRunBtn.disabled = false;
        hideRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Remove Alias Addresses ────────────────────────────────────────────────
  const aliasBrowseBtn = document.getElementById('aliasBrowseBtn');
  if (aliasBrowseBtn) {
    aliasBrowseBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.showOpenDialog({ properties: ['openDirectory'] });
      if (!result.canceled && result.filePaths.length > 0) {
        document.getElementById('aliasDiscoveryFolder').value = result.filePaths[0];
      }
    });
  }

  const aliasRunBtn = document.getElementById('aliasRunBtn');
  if (aliasRunBtn) {
    aliasRunBtn.addEventListener('click', async () => {
      const folder         = document.getElementById('aliasDiscoveryFolder').value.trim();
      const domain         = document.getElementById('aliasDomain').value.trim();
      const removeAliases  = document.getElementById('aliasRemoveAliases').checked;
      const removeSIP      = document.getElementById('aliasRemoveSIP').checked;
      const whatIf         = document.getElementById('aliasWhatIf').checked;

      if (!folder) { alert('Please select a discovery folder.'); return; }
      if (!removeAliases && !removeSIP) { alert('Select at least one address type to remove.'); return; }

      const logSection = document.getElementById('aliasLog');
      const logOutput  = document.getElementById('aliasLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      aliasRunBtn.disabled = true;
      aliasRunBtn.textContent = 'Running…';

      const args = ['-DiscoveryFolder', folder];
      if (!removeAliases) args.push('-SkipAliases');
      if (!removeSIP)     args.push('-SkipSIP');
      if (domain) args.push('-Domain', domain);
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Remove-AliasAddresses.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        aliasRunBtn.disabled = false;
        aliasRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── SIP / IM Addresses — mode toggle ──────────────────────────────────────
  const sipModeEl = document.getElementById('sipMode');
  if (sipModeEl) {
    const sipToggle = () => {
      const isRemove = sipModeEl.value === 'Remove';
      document.getElementById('sipRemoveFields').style.display   = isRemove ? '' : 'none';
      document.getElementById('sipNewDomainGroup').style.display = isRemove ? 'none' : '';
    };
    sipModeEl.addEventListener('change', sipToggle);
    sipToggle();
  }

  // ── SIP / IM Addresses — run ───────────────────────────────────────────────
  const sipRunBtn = document.getElementById('sipRunBtn');
  if (sipRunBtn) {
    sipRunBtn.addEventListener('click', async () => {
      const mode      = document.getElementById('sipMode').value;
      const oldDomain = document.getElementById('sipOldDomain').value.trim();
      const newDomain = document.getElementById('sipNewDomain').value.trim();
      const skuId     = document.getElementById('sipLicenceSkuId').value.trim();
      const waitMins  = parseInt(document.getElementById('sipWaitMinutes').value, 10) || 5;
      const whatIf    = document.getElementById('sipWhatIf').checked;

      if (!oldDomain) { alert('Please enter the old domain.'); return; }
      if (mode === 'Replace' && !newDomain) { alert('Please enter the new domain.'); return; }

      const logSection = document.getElementById('sipLog');
      const logOutput  = document.getElementById('sipLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      sipRunBtn.disabled = true;
      sipRunBtn.textContent = 'Running…';

      const args = ['-OldDomain', oldDomain, '-Mode', mode];
      if (mode === 'Replace') {
        args.push('-NewDomain', newDomain);
      } else {
        if (skuId) args.push('-LicenseSkuId', skuId);
        args.push('-WaitMinutes', String(waitMins));
      }
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Update-SIPDomain.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        sipRunBtn.disabled = false;
        sipRunBtn.textContent = '▶ Run';
      }
    });
  }

  // ── Remove Entra Users (last resort) ─────────────────────────────────────
  const entraRemoveBrowseBtn = document.getElementById('entraRemoveBrowseBtn');
  if (entraRemoveBrowseBtn) {
    entraRemoveBrowseBtn.addEventListener('click', async () => {
      const result = await window.electronAPI.showOpenDialog({ properties: ['openDirectory'] });
      if (result && !result.canceled && result.filePaths.length > 0) {
        document.getElementById('entraRemoveOutputFolder').value = result.filePaths[0];
      }
    });
  }

  const entraRemoveRunBtn = document.getElementById('entraRemoveRunBtn');
  if (entraRemoveRunBtn) {
    entraRemoveRunBtn.addEventListener('click', async () => {
      const domain       = document.getElementById('entraRemoveDomain').value.trim();
      const outputFolder = document.getElementById('entraRemoveOutputFolder').value.trim();
      const whatIf       = document.getElementById('entraRemoveWhatIf').checked;

      if (!domain) { alert('Please enter the domain to target.'); return; }

      const logSection = document.getElementById('entraRemoveLog');
      const logOutput  = document.getElementById('entraRemoveLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = '';

      entraRemoveRunBtn.disabled = true;
      entraRemoveRunBtn.textContent = 'Running…';

      const args = ['-Domain', domain];
      if (outputFolder) args.push('-OutputFolder', outputFolder);
      if (whatIf) args.push('-WhatIf');

      window.electronAPI.onPsOutput((text) => {
        logOutput.textContent += text;
        logOutput.scrollTop = logOutput.scrollHeight;
      });
      try {
        const result = await window.electronAPI.streamPowerShell('Remove-EntraUsers.ps1', args);
        logOutput.textContent += result.success ? '\n✓ Done\n' : `\n✗ Failed (exit ${result.code})\n`;
      } catch (err) {
        logOutput.textContent += `\nError: ${err.message || err}\n`;
      } finally {
        window.electronAPI.offPsOutput();
        entraRemoveRunBtn.disabled = false;
        entraRemoveRunBtn.textContent = '▶ Export & Remove';
      }
    });
  }

  // ── Workflow step navigation buttons ──────────────────────────────────────
  document.querySelectorAll('.workflow-nav-btn[data-view]').forEach((btn) => {
    btn.addEventListener('click', () => switchView(btn.getAttribute('data-view')));
  });

});
