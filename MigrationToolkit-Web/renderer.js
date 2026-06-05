// Migration Toolkit - Renderer Process (UI Logic)

console.log('=== RENDERER.JS LOADING ===');
console.log('Document ready state:', document.readyState);

// Wait for DOM to load
document.addEventListener('DOMContentLoaded', async () => {
  console.log('=== DOMCONTENTLOADED EVENT FIRED ===');
  console.log('Migration Toolkit loaded');
  console.log('electronAPI available:', !!window.electronAPI);

  // Test if clicks work at all
  document.body.addEventListener('click', (e) => {
    console.log('!!! BODY CLICK DETECTED !!!', e.target.tagName, e.target.className);
  });

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

  // Discovery - Start button
  const startDiscoveryBtn = document.getElementById('startDiscoveryBtn');
  if (startDiscoveryBtn) {
    startDiscoveryBtn.addEventListener('click', async () => {
      const domain = document.getElementById('discoveryDomain').value.trim();
      const vbuId = document.getElementById('discoveryVbuId').value.trim();
      const skipPP = document.getElementById('skipPowerPlatform').checked;
      const hybrid = document.getElementById('hybrid').checked;
      const members = document.getElementById('includeMembers').checked;
      const continueOnError = document.getElementById('continueOnError').checked;

      if (!domain) {
        alert('Please enter a domain name');
        return;
      }

      // Show log section
      const logSection = document.getElementById('discoveryLog');
      const logOutput = document.getElementById('discoveryLogOutput');
      logSection.classList.remove('hidden');
      logOutput.textContent = `Starting discovery for ${domain}...\n`;
      logOutput.textContent += `VBU ID: ${vbuId || 'Not specified'}\n`;
      logOutput.textContent += `Skip Power Platform: ${skipPP}\n`;
      logOutput.textContent += `Hybrid Mode: ${hybrid}\n`;
      logOutput.textContent += `Include Members: ${members}\n`;
      logOutput.textContent += `Continue on Error: ${continueOnError}\n\n`;
      logOutput.textContent += `Launching PowerShell discovery script...\n`;

      // Launch the actual PowerShell script
      await launchScript('discovery-menu.ps1', startDiscoveryBtn);
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
    input3.className = 'form-input form-input-compact customer-spo';
    input3.placeholder = 'https://tenant-admin.sharepoint.com';

    td1.appendChild(input1);
    td2.appendChild(input2);
    td3.appendChild(input3);

    newRow.appendChild(td1);
    newRow.appendChild(td2);
    newRow.appendChild(td3);

    tbody.appendChild(newRow);
  }

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
  const monitorRefreshBtn = document.getElementById('monitorRefreshBtn');
  const monitorAutoRefresh = document.getElementById('monitorAutoRefresh');
  const monitorInterval = document.getElementById('monitorInterval');
  const monitorProject = document.getElementById('monitorProject');
  const monitorStatus = document.getElementById('monitorStatus');
  const monitorTableBody = document.getElementById('monitorTableBody');
  const monitorConnection = document.getElementById('monitorConnection');

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

          // Apply row styling based on status
          if (failed > 0) {
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
            <td>${failed}</td>
            <td>${warnings}</td>
            <td>${now}</td>
          `;
          monitorTableBody.appendChild(row);
        });

        monitorStatus.textContent = `Last refresh: ${now}`;
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
        const customers = config.config.Customers;
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

  function updateWorkloadBars() {
    workloadBars.innerHTML = '';

    // Use real workload data from currentStatData
    const workloadsData = currentStatData.workloads || {};

    // If no workload data, show nothing
    if (Object.keys(workloadsData).length === 0) {
      workloadBars.innerHTML = '<p style="color: #6c757d; padding: 20px; text-align: center;">No workload data available</p>';
      return;
    }

    // Process each workload from real data
    Object.keys(workloadsData).forEach(workloadName => {
      const wlData = workloadsData[workloadName];
      const total = wlData.Total || 0;
      const completed = wlData.Completed || 0;
      const failed = wlData.Failed || 0;
      const warnings = wlData.Warnings || 0;
      const inProgress = wlData.InProgress || 0;
      const progress = total > 0 ? Math.round((completed / total) * 100) : 0;

      let statusClass = 'status-ontrack';
      let statusBadge = 'On track';
      let badgeClass = 'badge-ontrack';

      if (failed > 0) {
        statusClass = 'status-failed';
        statusBadge = `${failed} failed`;
        badgeClass = 'badge-failed';
      } else if (warnings > 0) {
        statusClass = 'status-warning';
        statusBadge = `${warnings} warnings`;
        badgeClass = 'badge-warning';
      } else if (inProgress > 0) {
        statusBadge = `${inProgress} in progress`;
        badgeClass = 'badge-ontrack';
      }

      const bar = document.createElement('div');
      bar.className = 'workload-bar-item';
      bar.setAttribute('data-workload', workloadName.toLowerCase());

      bar.innerHTML = `
        <div class="workload-bar-name">${workloadName}</div>
        <div class="workload-bar-progress">
          <div class="workload-bar-fill ${statusClass}" style="width: ${progress}%"></div>
        </div>
        <div class="workload-bar-count">${completed}/${total}</div>
        <div class="workload-bar-status">
          <span class="status-badge ${badgeClass}">${statusBadge}</span>
        </div>
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
      const customers = config.config.Customers;
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
      const customers = config.config.Customers;
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
    // Domain Removal sub-views
    'domain-workflow': 'domainWorkflowView',
    'domain-remove': 'domainRemoveView',
    'domain-onprem': 'domainOnPremView',
    'domain-cloud': 'domainCloudView',
    'domain-hide': 'domainHideView'
  };

  const targetViewId = viewMap[viewName];
  if (targetViewId) {
    const targetView = document.getElementById(targetViewId);
    if (targetView) {
      targetView.classList.remove('hidden');

      // Load dropdowns when specific views are shown
      if (viewName === 'misc-onedrive') {
        loadOneDriveTenants();
      } else if (viewName === 'avepoint-monitor') {
        loadMonitorProjects();
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

// Manual check for updates (from settings)
async function manualCheckUpdates() {
  const statusSpan = document.getElementById('updateStatus');
  const btn = document.getElementById('checkUpdatesBtn');

  try {
    btn.disabled = true;
    statusSpan.textContent = 'Checking...';
    statusSpan.style.color = '#6c757d';

    const result = await window.electronAPI.checkUpdates();

    if (result.success) {
      if (result.output && result.output.includes('UPDATE_AVAILABLE')) {
        statusSpan.textContent = '✓ Update available! Restart to install.';
        statusSpan.style.color = '#28a745';
      } else {
        statusSpan.textContent = '✓ You have the latest version';
        statusSpan.style.color = '#28a745';
      }
    } else {
      statusSpan.textContent = `❌ ${result.error || 'Check failed'}`;
      statusSpan.style.color = '#dc3545';
    }
  } catch (error) {
    statusSpan.textContent = `❌ Error: ${error.message}`;
    statusSpan.style.color = '#dc3545';
  } finally {
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

          // Create inputs properly to avoid HTML escaping issues
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
          input3.className = 'form-input form-input-compact customer-spo';
          input3.value = customer.SharePointAdminUrl || '';
          input3.placeholder = 'https://tenant-admin.sharepoint.com';

          td1.appendChild(input1);
          td2.appendChild(input2);
          td3.appendChild(input3);

          row.appendChild(td1);
          row.appendChild(td2);
          row.appendChild(td3);

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

    // Collect customer data from table
    const customerRows = document.querySelectorAll('#customerTableBody tr');
    const customers = [];

    customerRows.forEach(row => {
      const prefix = row.querySelector('.customer-prefix').value.trim();
      const accountName = row.querySelector('.customer-account').value.trim();
      const spoUrl = row.querySelector('.customer-spo').value.trim();

      if (prefix) {  // Only add if prefix is not empty
        customers.push({
          Prefix: prefix,
          AccountName: accountName,
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

      // Reload customer domains in dashboard
      if (typeof loadCustomerDomains === 'function') {
        loadCustomerDomains();
      }
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

// PowerShell Automation - needs to be called during DOMContentLoaded
// Initialize after DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  const psCustomerPrefix = document.getElementById('psCustomerPrefix');
  const psWorkload = document.getElementById('psWorkload');
  const psMappingFile = document.getElementById('psMappingFile');
  const psBrowseMappingBtn = document.getElementById('psBrowseMappingBtn');
  const psCreateProjectBtn = document.getElementById('psCreateProjectBtn');
  const psImportMappingsBtn = document.getElementById('psImportMappingsBtn');
  const psStartWorkflowBtn = document.getElementById('psStartWorkflowBtn');
  const psViewDocsBtn = document.getElementById('psViewDocsBtn');

  if (psBrowseMappingBtn) {
    psBrowseMappingBtn.addEventListener('click', async () => {
      // TODO: Implement file browser dialog
      alert('File browser not yet implemented. Please type the path manually for now.');
    });
  }

  if (psCreateProjectBtn) {
    psCreateProjectBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix.value.trim();
      const workload = psWorkload.value;

      if (!prefix || !workload) {
        alert('Please enter Customer Prefix and select Workload');
        return;
      }

      const projectName = `${prefix} - ${workload}`;
      const confirmed = confirm(`Create project: ${projectName}?`);
      if (!confirmed) return;

      psCreateProjectBtn.disabled = true;
      psCreateProjectBtn.textContent = 'Creating...';

      try {
        const result = await window.electronAPI.executePowerShell('New-FlyProject.ps1', [
          '-ProjectName', projectName,
          '-Workload', workload
        ]);

        if (result.success) {
          alert(`✓ Project created successfully!\n\n${result.output}`);
        } else {
          alert(`✗ Failed to create project:\n\n${result.error}`);
        }
      } catch (error) {
        alert(`Error: ${error.message}`);
      } finally {
        psCreateProjectBtn.disabled = false;
        psCreateProjectBtn.textContent = '📁 Create Project';
      }
    });
  }

  if (psImportMappingsBtn) {
    psImportMappingsBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix.value.trim();
      const workload = psWorkload.value;
      const mappingFile = psMappingFile.value.trim();

      if (!prefix || !workload || !mappingFile) {
        alert('Please enter Customer Prefix, Workload, and Mapping File');
        return;
      }

      const projectName = `${prefix} - ${workload}`;
      const confirmed = confirm(`Import mappings to: ${projectName}?\n\nFile: ${mappingFile}`);
      if (!confirmed) return;

      psImportMappingsBtn.disabled = true;
      psImportMappingsBtn.textContent = 'Importing...';

      try {
        const result = await window.electronAPI.executePowerShell('Import-FlyMappings.ps1', [
          '-ProjectName', projectName,
          '-Workload', workload,
          '-MappingFile', mappingFile
        ]);

        if (result.success) {
          alert(`✓ Mappings imported successfully!\n\n${result.output}`);
        } else {
          alert(`✗ Failed to import mappings:\n\n${result.error}`);
        }
      } catch (error) {
        alert(`Error: ${error.message}`);
      } finally {
        psImportMappingsBtn.disabled = false;
        psImportMappingsBtn.textContent = '📥 Import Mappings';
      }
    });
  }

  if (psStartWorkflowBtn) {
    psStartWorkflowBtn.addEventListener('click', async () => {
      const prefix = psCustomerPrefix.value.trim();
      const workload = psWorkload.value;
      const mappingFile = psMappingFile.value.trim();

      if (!prefix || !workload || !mappingFile) {
        alert('Please enter Customer Prefix, Workload, and Mapping File');
        return;
      }

      const confirmed = confirm(
        `Run FULL migration workflow?\n\n` +
        `Customer: ${prefix}\n` +
        `Workload: ${workload}\n` +
        `Mapping: ${mappingFile}\n\n` +
        `This will:\n` +
        `1. Create project\n` +
        `2. Import mappings\n` +
        `3. Run pre-scan\n` +
        `4. Run verification\n` +
        `5. Start migration\n\n` +
        `Continue?`
      );
      if (!confirmed) return;

      psStartWorkflowBtn.disabled = true;
      psStartWorkflowBtn.textContent = 'Running Workflow...';

      try {
        const result = await window.electronAPI.executePowerShell('Start-FlyMigrationWorkflow.ps1', [
          '-CustomerPrefix', prefix,
          '-Workload', workload,
          '-MappingFile', mappingFile
        ]);

        if (result.success) {
          alert(`✓ Workflow completed successfully!\n\n${result.output}`);
        } else {
          alert(`✗ Workflow failed:\n\n${result.error}`);
        }
      } catch (error) {
        alert(`Error: ${error.message}`);
      } finally {
        psStartWorkflowBtn.disabled = false;
        psStartWorkflowBtn.textContent = '🚀 Run Full Workflow';
      }
    });
  }

  if (psViewDocsBtn) {
    psViewDocsBtn.addEventListener('click', async () => {
      try {
        await window.electronAPI.openExternal('file:///C:/Temp/Scripts/VGMigrations/POWERSHELL-AUTOMATION.md');
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
        `1. Open a browser for authentication\n` +
        `2. Create app in Azure AD\n` +
        `3. Configure Graph & SharePoint permissions\n` +
        `4. Generate Client ID & Secret\n\n` +
        `You must be Global Admin or Application Admin.\n\n` +
        `Continue?`
      );
      if (!confirmed) return;

      createAppRegBtn.disabled = true;
      createAppRegBtn.textContent = 'Creating...';
      appRegStatus.style.display = 'block';
      appRegStatus.style.color = '#6c757d';
      appRegStatus.textContent = '⏳ Launching PowerShell script... Check the PowerShell window for authentication prompts.';

      try {
        const result = await window.electronAPI.executePowerShell('New-AzureAppRegistration.ps1', [
          '-TenantId', tenantId,
          '-AppName', appName,
          '-Interactive'
        ]);

        if (result.success) {
          appRegStatus.style.color = '#28a745';
          appRegStatus.textContent = '✓ App registration created! Check the PowerShell window for credentials.';
          alert(
            `✓ App Registration Created!\n\n` +
            `IMPORTANT: Copy the credentials from the PowerShell window.\n\n` +
            `Next Steps:\n` +
            `1. Copy the Client ID and Secret from PowerShell\n` +
            `2. Grant admin consent in Azure Portal\n` +
            `3. Enter credentials in Settings > Config tab\n` +
            `4. Test connection\n\n` +
            `See PowerShell window for full details.`
          );
        } else {
          appRegStatus.style.color = '#dc3545';
          appRegStatus.textContent = `✗ Failed: ${result.error}`;
          alert(`✗ Failed to create app registration:\n\n${result.error}`);
        }
      } catch (error) {
        appRegStatus.style.color = '#dc3545';
        appRegStatus.textContent = `✗ Error: ${error.message}`;
        alert(`Error: ${error.message}`);
      } finally {
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

});
