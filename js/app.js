/**
 * DMC Hospital Statistics Dashboard (2026) - Application Controller (app.js)
 * Clean tab navigation, dynamic monthly tables, and interactive drilldowns.
 */

document.addEventListener('DOMContentLoaded', () => {
  HospitalApp.init();
});

const HospitalApp = {
  currentFilter: {
    period: 'all',
    monthIndex: null,
    departmentId: 'all',
    activeTableTab: 'opd',
    currentTab: 'overview'
  },

  init() {
    this.initTheme();
    this.initClock();
    this.bindEvents();
    ExcelDataParser.init();
    HospitalCharts.initAll();
    this.updateDashboardMetrics();
    this.populateDepartmentFilter();
    this.renderAllTables();
  },

  /**
   * Theme Manager
   */
  initTheme() {
    const savedTheme = localStorage.getItem('dmc_theme') || 'light';
    this.setTheme(savedTheme);

    const themeToggleBtn = document.getElementById('themeToggleBtn');
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
        const nextTheme = currentTheme === 'light' ? 'dark' : 'light';
        this.setTheme(nextTheme);
      });
    }
  },

  setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('dmc_theme', theme);

    const icon = document.querySelector('#themeToggleBtn i');
    if (icon) {
      icon.className = theme === 'dark' ? 'fas fa-sun' : 'fas fa-moon';
    }

    if (HospitalCharts.instances.ipdDonut) {
      HospitalCharts.updateAll(this.currentFilter.monthIndex);
    }
  },

  /**
   * Header Live Clock
   */
  initClock() {
    const clockEl = document.getElementById('liveClockText');
    const update = () => {
      if (!clockEl) return;
      const now = new Date();
      clockEl.textContent = now.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
    };
    update();
    setInterval(update, 1000);
  },

  /**
   * Bind event listeners
   */
  bindEvents() {
    // Sidebar Tab Navigation
    const navLinks = document.querySelectorAll('.sidebar-menu li a[data-tab]');
    navLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        const tabId = link.getAttribute('data-tab');
        this.switchTab(tabId);
      });
    });

    // Mobile Sidebar Toggle
    const sidebarToggleBtn = document.getElementById('sidebarToggleBtn');
    const sidebar = document.querySelector('.app-sidebar');
    if (sidebarToggleBtn && sidebar) {
      sidebarToggleBtn.addEventListener('click', () => {
        sidebar.classList.toggle('sidebar-open');
      });
    }

    // Period Preset Buttons (All, Q1, Q2, Jul, Aug)
    const presetBtns = document.querySelectorAll('.preset-pill[data-period]');
    presetBtns.forEach(btn => {
      btn.addEventListener('click', (e) => {
        presetBtns.forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        const period = e.target.dataset.period;
        this.handlePeriodChange(period);
      });
    });

    // Single Month Selector Dropdown
    const monthFilterSelect = document.getElementById('monthFilterSelect');
    if (monthFilterSelect) {
      monthFilterSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        const presetBtns = document.querySelectorAll('.preset-pill[data-period]');
        presetBtns.forEach(b => b.classList.remove('active'));

        if (val === 'all') {
          const allBtn = document.querySelector('.preset-pill[data-period="all"]');
          if (allBtn) allBtn.classList.add('active');
          this.handlePeriodChange('all');
        } else if (val === '6') {
          const julBtn = document.querySelector('.preset-pill[data-period="jul"]');
          if (julBtn) julBtn.classList.add('active');
          this.handlePeriodChange(6);
        } else if (val === '7') {
          const augBtn = document.querySelector('.preset-pill[data-period="aug"]');
          if (augBtn) augBtn.classList.add('active');
          this.handlePeriodChange(7);
        } else {
          this.handlePeriodChange(parseInt(val, 10));
        }
      });
    }

    // Department Filter Dropdown (Deep Dive / Drilldown)
    const deptFilterSelect = document.getElementById('deptFilterSelect');
    if (deptFilterSelect) {
      deptFilterSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        this.currentFilter.departmentId = val;
        if (val && val !== 'all') {
          const isOPD = val.startsWith('opd_');
          const deptId = val.replace(/^(opd_|ipd_)/, '');
          const dept = isOPD
            ? hospitalData.outpatient.find(d => d.id === deptId)
            : hospitalData.inpatient.find(w => w.id === deptId);
          if (dept) {
            this.showDepartmentDrilldown(isOPD ? 'outpatient' : 'inpatient', dept.name);
          }
        }
      });
    }

    // Print / PDF Button
    const printReportBtn = document.getElementById('printReportBtn');
    if (printReportBtn) {
      printReportBtn.addEventListener('click', () => {
        window.print();
      });
    }

    // Share / Snapshot Button
    const shareBtn = document.getElementById('shareDashboardBtn');
    if (shareBtn) {
      shareBtn.addEventListener('click', () => {
        this.shareSnapshot();
      });
    }

    // Excel Upload Modal
    const openUploadBtn = document.getElementById('openUploadModalBtn');
    const uploadModal = document.getElementById('uploadModal');
    const closeUploadBtn = document.getElementById('closeUploadModalBtn');
    if (openUploadBtn && uploadModal) {
      openUploadBtn.addEventListener('click', () => {
        uploadModal.classList.add('show');
      });
    }
    if (closeUploadBtn && uploadModal) {
      closeUploadBtn.addEventListener('click', () => {
        uploadModal.classList.remove('show');
      });
    }

    // Drilldown modal close
    const drilldownModal = document.getElementById('drilldownModal');
    const closeDrilldownBtn = document.getElementById('closeDrilldownModalBtn');
    if (closeDrilldownBtn && drilldownModal) {
      closeDrilldownBtn.addEventListener('click', () => {
        drilldownModal.classList.remove('show');
      });
    }

    [uploadModal, drilldownModal].forEach(m => {
      if (m) {
        m.addEventListener('click', (e) => {
          if (e.target === m) m.classList.remove('show');
        });
      }
    });

    // Tab Table Switcher in Master Table
    const tabTableOPD = document.getElementById('tabTableOPD');
    const tabTableIPD = document.getElementById('tabTableIPD');
    if (tabTableOPD && tabTableIPD) {
      tabTableOPD.addEventListener('click', () => {
        tabTableOPD.classList.add('active');
        tabTableIPD.classList.remove('active');
        this.currentFilter.activeTableTab = 'opd';
        this.renderMasterTable();
      });
      tabTableIPD.addEventListener('click', () => {
        tabTableIPD.classList.add('active');
        tabTableOPD.classList.remove('active');
        this.currentFilter.activeTableTab = 'ipd';
        this.renderMasterTable();
      });
    }

    // Global callback on Excel file ingestion
    window.onHospitalDataUpdated = () => {
      this.populateDepartmentFilter();
      this.updateDashboardMetrics();
      HospitalCharts.updateAll(this.currentFilter.periodFilter);
      this.renderAllTables();
      if (uploadModal) uploadModal.classList.remove('show');
    };
  },

  /**
   * Switch Active Tab Panel
   */
  switchTab(tabId) {
    this.currentFilter.currentTab = tabId;

    // Update Sidebar Active state
    document.querySelectorAll('.sidebar-menu li').forEach(li => li.classList.remove('active'));
    const activeLink = document.querySelector(`.sidebar-menu li a[data-tab="${tabId}"]`);
    if (activeLink && activeLink.parentElement) {
      activeLink.parentElement.classList.add('active');
    }

    // Show selected panel, hide others
    document.querySelectorAll('.dashboard-tab-panel').forEach(panel => {
      panel.classList.add('tab-hidden');
    });

    const targetPanel = document.getElementById(`tab-${tabId}`);
    if (targetPanel) {
      targetPanel.classList.remove('tab-hidden');
    }

    // Scroll to top of content
    window.scrollTo({ top: 0, behavior: 'smooth' });

    // Resize Chart.js instances so they redraw correctly inside newly visible tabs
    setTimeout(() => {
      Object.values(HospitalCharts.instances).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
          chart.resize();
        }
      });
    }, 50);

    // Close mobile drawer if open
    const sidebar = document.querySelector('.app-sidebar');
    if (sidebar) sidebar.classList.remove('sidebar-open');
  },

  /**
   * Handle Period Filter (All, Q1, Q2, Jul, Aug, or individual month index)
   */
  handlePeriodChange(period) {
    this.currentFilter.period = String(period);
    const monthSelect = document.getElementById('monthFilterSelect');

    if (period === 'all') {
      this.currentFilter.periodFilter = null;
      this.currentFilter.monthIndex = null;
      if (monthSelect) monthSelect.value = 'all';
    } else if (period === 'q1') {
      this.currentFilter.periodFilter = [0, 1, 2]; // Jan, Feb, Mar
      this.currentFilter.monthIndex = null;
      if (monthSelect) monthSelect.value = 'all';
    } else if (period === 'q2') {
      this.currentFilter.periodFilter = [3, 4, 5]; // Apr, May, Jun
      this.currentFilter.monthIndex = null;
      if (monthSelect) monthSelect.value = 'all';
    } else if (period === 'jul') {
      this.currentFilter.periodFilter = 6; // July
      this.currentFilter.monthIndex = 6;
      if (monthSelect) monthSelect.value = '6';
    } else if (period === 'aug') {
      this.currentFilter.periodFilter = 7; // August
      this.currentFilter.monthIndex = 7;
      if (monthSelect) monthSelect.value = '7';
    } else {
      const idx = parseInt(period, 10);
      if (!isNaN(idx) && idx >= 0 && idx < 8) {
        this.currentFilter.periodFilter = idx;
        this.currentFilter.monthIndex = idx;
        if (monthSelect) monthSelect.value = idx.toString();
      } else {
        this.currentFilter.periodFilter = null;
        this.currentFilter.monthIndex = null;
        if (monthSelect) monthSelect.value = 'all';
      }
    }

    this.updateDashboardMetrics();
    HospitalCharts.updateAll(this.currentFilter.periodFilter);
    this.renderAllTables();
  },

  /**
   * Populate Department Dropdown Filter
   */
  populateDepartmentFilter() {
    const select = document.getElementById('deptFilterSelect');
    if (!select) return;

    select.innerHTML = '<option value="all">All Departments (13 OPD / 9 IPD)</option>';

    const opdGroup = document.createElement('optgroup');
    opdGroup.label = "Outpatient (OPD)";
    hospitalData.outpatient.forEach(d => {
      const opt = document.createElement('option');
      opt.value = `opd_${d.id}`;
      opt.textContent = `${d.name} (OPD)`;
      select.appendChild(opt);
    });
    select.appendChild(opdGroup);

    const ipdGroup = document.createElement('optgroup');
    ipdGroup.label = "Inpatient (IPD)";
    hospitalData.inpatient.forEach(w => {
      const opt = document.createElement('option');
      opt.value = `ipd_${w.id}`;
      opt.textContent = `${w.name} (IPD)`;
      ipdGroup.appendChild(opt);
    });
    select.appendChild(ipdGroup);
  },

  /**
   * Update Top Hero KPI Cards
   */
  updateDashboardMetrics() {
    const pf = this.currentFilter.periodFilter;

    const opdTotal = HospitalAnalytics.getOPDTotal(pf);
    this.animateValue('kpiOPDCount', opdTotal);

    const ipdTotal = HospitalAnalytics.getIPDTotal(pf);
    this.animateValue('kpiIPDCount', ipdTotal);

    const grandTotal = opdTotal + ipdTotal;
    const grandTotalEl = document.getElementById('kpiGrandTotal');
    if (grandTotalEl) grandTotalEl.innerHTML = `<i class="fas fa-users-medical"></i> ${grandTotal.toLocaleString()} Total Patients`;

    const deliveriesTotal = HospitalAnalytics.getDeliveriesTotal(pf);
    const liveBirthsTotal = HospitalAnalytics.getLiveBirthsTotal(pf);
    this.animateValue('kpiDeliveriesCount', deliveriesTotal);
    const liveBirthsEl = document.getElementById('kpiLiveBirthsSub');
    if (liveBirthsEl) {
      liveBirthsEl.textContent = `${liveBirthsTotal.toLocaleString()} Live Births`;
    }

    const csStats = HospitalAnalytics.getCSStats(pf);
    const csRateEl = document.getElementById('kpiCSRate');
    if (csRateEl) {
      csRateEl.textContent = `${csStats.rate}%`;
    }
    const csRatioEl = document.getElementById('kpiCSRatioSub');
    if (csRatioEl) {
      csRatioEl.textContent = `${csStats.cs.toLocaleString()} CS · ${csStats.svd.toLocaleString()} SVD`;
    }

    const mortBadge = document.getElementById('mortalityRatePillBadge');
    if (mortBadge) {
      const mortDeaths = HospitalAnalytics.getMortalityTotal(pf);
      const mortRate = ipdTotal > 0 ? ((mortDeaths / ipdTotal) * 100).toFixed(2) : '0.00';
      mortBadge.textContent = `${mortRate}% Mortality`;
    }
  },

  /**
   * Animate numeric value changes
   */
  animateValue(elementId, endVal, duration = 400) {
    const el = document.getElementById(elementId);
    if (!el) return;

    const startVal = parseInt(el.textContent.replace(/[^0-9]/g, ''), 10) || 0;
    if (startVal === endVal) {
      el.textContent = endVal.toLocaleString();
      return;
    }

    const startTime = performance.now();
    const step = (currentTime) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      const easeProgress = 1 - Math.pow(1 - progress, 3);
      const current = Math.floor(startVal + (endVal - startVal) * easeProgress);
      el.textContent = current.toLocaleString();

      if (progress < 1) {
        requestAnimationFrame(step);
      } else {
        el.textContent = endVal.toLocaleString();
      }
    };
    requestAnimationFrame(step);
  },

  /**
   * Render All Tables Across Tabs
   */
  renderAllTables() {
    this.renderInpatientTable();
    this.renderOutpatientTable();
    this.renderMaternityTable();
    this.renderMaternityHistoricalTable();
    this.renderMortalityTable();
    this.renderMasterTable();
  },

  /**
   * Render Inpatient Table (Tab 2)
   */
  renderInpatientTable() {
    const table = document.getElementById('inpatientFullTable');
    if (!table) return;

    const months = hospitalData.metadata.months;
    const data = hospitalData.inpatient;
    const monthlySums = months.map(() => 0);
    let totalSum = 0;

    let html = `
      <thead>
        <tr>
          <th style="min-width: 180px;">INPATIENT DEPARTMENT</th>
          ${months.map(m => `<th class="num-mono" style="text-align: right;">${m.toUpperCase()}</th>`).join('')}
          <th class="num-mono" style="text-align: right; font-weight: 700;">TOTAL</th>
          <th style="text-align: center;">ACTION</th>
        </tr>
      </thead>
      <tbody>
    `;

    data.forEach(dept => {
      totalSum += dept.total;
      months.forEach((_, i) => monthlySums[i] += (dept.monthly[i] || 0));

      html += `
        <tr>
          <td><strong style="color: var(--text-primary); font-weight: 600;">${dept.name}</strong></td>
          ${dept.monthly.map(v => `<td class="num-mono" style="text-align: right;">${v ? v.toLocaleString() : '0'}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--primary);">${dept.total.toLocaleString()}</td>
          <td style="text-align: center;">
            <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('inpatient', '${dept.name}')">
              <i class="fas fa-chart-line"></i> Inspect
            </button>
          </td>
        </tr>
      `;
    });

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td>GRAND TOTAL</td>
          ${monthlySums.map(s => `<td class="num-mono" style="text-align: right;">${s.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--primary);">${totalSum.toLocaleString()}</td>
          <td>—</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Outpatient Table (Tab 3)
   */
  renderOutpatientTable() {
    const table = document.getElementById('outpatientFullTable');
    if (!table) return;

    const months = hospitalData.metadata.months;
    const data = hospitalData.outpatient;
    const monthlySums = months.map(() => 0);
    let totalSum = 0;

    let html = `
      <thead>
        <tr>
          <th style="min-width: 180px;">OUTPATIENT DEPARTMENT</th>
          ${months.map(m => `<th class="num-mono" style="text-align: right;">${m.toUpperCase()}</th>`).join('')}
          <th class="num-mono" style="text-align: right; font-weight: 700;">TOTAL</th>
          <th style="text-align: center;">ACTION</th>
        </tr>
      </thead>
      <tbody>
    `;

    data.forEach(dept => {
      totalSum += dept.total;
      months.forEach((_, i) => monthlySums[i] += (dept.monthly[i] || 0));

      html += `
        <tr>
          <td><strong style="color: var(--text-primary); font-weight: 600;">${dept.name}</strong></td>
          ${dept.monthly.map(v => `<td class="num-mono" style="text-align: right;">${v ? v.toLocaleString() : '0'}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--secondary);">${dept.total.toLocaleString()}</td>
          <td style="text-align: center;">
            <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('outpatient', '${dept.name}')">
              <i class="fas fa-chart-line"></i> Inspect
            </button>
          </td>
        </tr>
      `;
    });

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td>GRAND TOTAL</td>
          ${monthlySums.map(s => `<td class="num-mono" style="text-align: right;">${s.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--secondary);">${totalSum.toLocaleString()}</td>
          <td>—</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Maternity Monthly Table (Tab 4)
   */
  renderMaternityTable() {
    const table = document.getElementById('maternityFullTable');
    if (!table) return;

    const data = hospitalData.maternity.monthly;
    let totDel = 0, totLive = 0, totCS = 0, totSVD = 0, totDeath = 0;

    let html = `
      <thead>
        <tr>
          <th>Period (Month)</th>
          <th class="num-mono" style="text-align: right;">Deliveries</th>
          <th class="num-mono" style="text-align: right;">Live Births</th>
          <th class="num-mono" style="text-align: right;">C-Section (CS)</th>
          <th class="num-mono" style="text-align: right;">SVD</th>
          <th class="num-mono" style="text-align: right;">CS Rate</th>
          <th class="num-mono" style="text-align: right;">Stillbirths</th>
          <th>Circumstance</th>
        </tr>
      </thead>
      <tbody>
    `;

    data.forEach(m => {
      totDel += m.deliveries;
      totLive += m.liveBirths;
      totCS += m.cs;
      totSVD += m.svd;
      totDeath += m.deaths;
      const csRate = m.deliveries > 0 ? ((m.cs / m.deliveries) * 100).toFixed(1) : "0.0";

      html += `
        <tr>
          <td><strong>${m.month} 2026</strong></td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${m.deliveries.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right;">${m.liveBirths.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 600;">${m.cs.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary);">${m.svd.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${csRate}%</td>
          <td class="num-mono" style="text-align: right; color: ${m.deaths > 0 ? 'var(--danger)' : 'inherit'};">${m.deaths}</td>
          <td style="font-size: 0.8rem; color: var(--text-muted);">${m.deathReason}</td>
        </tr>
      `;
    });

    const overallCSRate = totDel > 0 ? ((totCS / totDel) * 100).toFixed(1) : "0.0";

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td>TOTAL (JAN - AUG 2026)</td>
          <td class="num-mono" style="text-align: right;">${totDel.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right;">${totLive.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 800;">${totCS.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary);">${totSVD.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">${overallCSRate}%</td>
          <td class="num-mono" style="text-align: right; color: var(--danger); font-weight: 800;">${totDeath}</td>
          <td>Still births macerated</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Multi-Year Maternity Historical Record (2019 - 2026)
   * Source: Sheet 3: BABIES TOTAL (Deliveries/Babies from 2019 up to August 2026)
   */
  renderMaternityHistoricalTable() {
    const table = document.getElementById('maternityHistoricalTable');
    if (!table) return;

    const data = hospitalData.maternity.historical;
    let totDel = 0, totLive = 0, totCS = 0, totSVD = 0;

    let html = `
      <thead>
        <tr>
          <th>PERIOD (YEAR)</th>
          <th class="num-mono" style="text-align: right;">DELIVERIES AT HF</th>
          <th class="num-mono" style="text-align: right;">LIVE BIRTHS</th>
          <th class="num-mono" style="text-align: right;">C-SECTION (CS)</th>
          <th class="num-mono" style="text-align: right;">SVD</th>
          <th class="num-mono" style="text-align: right;">CS RATE</th>
        </tr>
      </thead>
      <tbody>
    `;

    data.forEach(item => {
      totDel += item.deliveries;
      totLive += item.liveBirths;
      totCS += item.cs;
      totSVD += item.svd;
      const csRate = item.deliveries > 0 ? ((item.cs / item.deliveries) * 100).toFixed(1) : "0.0";

      html += `
        <tr>
          <td><strong style="color: var(--text-primary);">${item.year}</strong></td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${item.deliveries.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right;">${item.liveBirths.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 600;">${item.cs.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary);">${item.svd.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${csRate}%</td>
        </tr>
      `;
    });

    const cumCSRate = totDel > 0 ? ((totCS / totDel) * 100).toFixed(1) : "0.0";

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td>CUMULATIVE (2019 – 2026)</td>
          <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--primary); font-weight: 800;">${totDel.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">${totLive.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 800;">${totCS.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary); font-weight: 800;">${totSVD.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">${cumCSRate}%</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Documented Clinical Mortality Register (Tab 5)
   * Source: Sheet 4: Death
   */
  renderMortalityTable() {
    const table = document.getElementById('mortalityLogTable');
    if (!table) return;

    const data = hospitalData.mortality;
    let totalDeaths = 0;

    let html = `
      <thead>
        <tr>
          <th style="min-width: 120px;">PERIOD (DATE)</th>
          <th>MONTH</th>
          <th class="num-mono" style="text-align: right;">DEATHS (HF)</th>
          <th>AGE</th>
          <th>DEPARTMENT</th>
          <th>CIRCUMSTANCE OF DEATH</th>
        </tr>
      </thead>
      <tbody>
    `;

    data.forEach(item => {
      totalDeaths += item.count;
      html += `
        <tr>
          <td class="num-mono" style="color: var(--text-primary); font-weight: 500;">${item.date}</td>
          <td><strong>${item.month}</strong></td>
          <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--danger); font-size: 1rem;">${item.count}</td>
          <td>${item.age === '0' ? '<span class="kpi-badge badge-warning" style="font-size: 0.75rem; padding: 2px 8px;">0 (Neonate)</span>' : '<span style="color: var(--text-muted);">' + (item.age || '—') + '</span>'}</td>
          <td><strong style="color: var(--text-primary);">${item.department}</strong></td>
          <td><span style="color: var(--text-muted); font-size: 0.85rem;">${item.circumstance}</span></td>
        </tr>
      `;
    });

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td colspan="2">TOTAL IN-FACILITY DEATHS (JAN – AUG 2026)</td>
          <td class="num-mono" style="text-align: right; font-weight: 800; color: var(--danger); font-size: 1.05rem;">${totalDeaths}</td>
          <td colspan="3" style="color: var(--text-muted); font-size: 0.82rem;">${((totalDeaths / HospitalAnalytics.getIPDTotal()) * 100).toFixed(2)}% mortality rate across ${HospitalAnalytics.getIPDTotal().toLocaleString()} inpatient admissions (Sheet 4: Death)</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Master Data Table (Tab 6)
   */
  renderMasterTable() {
    const thead = document.getElementById('dataTableHead');
    const tbody = document.getElementById('dataTableBody');
    const tfoot = document.getElementById('dataTableFoot');
    if (!tbody || !thead) return;

    const isOPD = this.currentFilter.activeTableTab === 'opd';
    const dataset = isOPD ? hospitalData.outpatient : hospitalData.inpatient;
    const months = hospitalData.metadata.months;

    thead.innerHTML = `
      <tr>
        <th scope="col" style="min-width: 180px;">Department (${isOPD ? 'OUTPATIENT' : 'INPATIENT'})</th>
        ${months.map(m => `<th scope="col" class="num-mono" style="text-align: right;">${m.toUpperCase()}</th>`).join('')}
        <th scope="col" class="num-mono" style="text-align: right; font-weight: 700;">Total</th>
        <th scope="col" style="text-align: center;">Action</th>
      </tr>
    `;

    tbody.innerHTML = '';
    const monthlyTotals = months.map(() => 0);
    let grandSum = 0;

    dataset.forEach(dept => {
      grandSum += dept.total;
      months.forEach((_, idx) => {
        monthlyTotals[idx] += (dept.monthly[idx] || 0);
      });

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong style="color: var(--text-primary); font-weight: 600;">${dept.name}</strong></td>
        ${dept.monthly.map(val => `<td class="num-mono" style="text-align: right;">${val ? val.toLocaleString() : '0'}</td>`).join('')}
        <td class="num-mono" style="text-align: right; font-weight: 700; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'};">${dept.total.toLocaleString()}</td>
        <td style="text-align: center;">
          <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('${isOPD ? 'outpatient' : 'inpatient'}', '${dept.name}')" title="Inspect monthly progression">
            <i class="fas fa-chart-line"></i> Inspect
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    if (tfoot) {
      tfoot.innerHTML = `
        <tr>
          <td>GRAND TOTAL</td>
          ${monthlyTotals.map(sum => `<td class="num-mono" style="text-align: right;">${sum.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 800; font-size: 1rem; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'};">${grandSum.toLocaleString()}</td>
          <td>—</td>
        </tr>
      `;
    }
  },

  /**
   * Department Performance Drilldown Modal
   */
  showDepartmentDrilldown(type, deptName) {
    const modal = document.getElementById('drilldownModal');
    const modalTitle = document.getElementById('drilldownModalTitle');
    const modalBody = document.getElementById('drilldownModalBody');
    if (!modal || !modalBody) return;

    // 1. Mortality Incident Drilldown
    if (type === 'mortality') {
      modalTitle.textContent = `${deptName} — Clinical Mortalities Detail`;
      const matching = hospitalData.mortality.filter(m => {
        const full = `${m.department} (${m.circumstance})`;
        return full.toLowerCase().includes(deptName.toLowerCase()) || deptName.toLowerCase().includes(m.department.toLowerCase());
      });
      const deathsCount = matching.reduce((s, m) => s + m.count, 0);
      const totalFacilityDeaths = hospitalData.mortality.reduce((s, m) => s + m.count, 0);
      const pctOfDeaths = totalFacilityDeaths > 0 ? ((deathsCount / totalFacilityDeaths) * 100).toFixed(1) : 0;

      modalBody.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px;">
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Recorded Mortalities</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--danger); margin-top: 4px;">
              ${deathsCount} ${deathsCount === 1 ? 'Death' : 'Deaths'}
            </div>
          </div>
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Share of Facility Deaths</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--primary); margin-top: 4px;">
              ${pctOfDeaths}%
            </div>
            <div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 2px;">of all ${totalFacilityDeaths} recorded facility deaths</div>
          </div>
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Clinical Register</div>
            <div style="font-size: 1.2rem; font-weight: 700; color: var(--text-primary); margin-top: 6px;">
              Sheet 4: Death
            </div>
            <div style="font-size: 0.72rem; color: var(--text-muted);">Verified clinical register</div>
          </div>
        </div>

        <h4 style="margin-bottom: 12px; font-size: 1rem; color: var(--text-primary);">Recorded Incidents</h4>
        <div class="table-responsive-wrapper" style="margin-bottom: 20px;">
          <table class="hospital-table" style="font-size: 0.85rem;">
            <thead>
              <tr>
                <th>Date</th>
                <th>Month</th>
                <th class="num-mono" style="text-align: right;">Count</th>
                <th>Age</th>
                <th>Department</th>
                <th>Circumstance</th>
              </tr>
            </thead>
            <tbody>
              ${matching.map(m => `
                <tr>
                  <td class="num-mono">${m.date}</td>
                  <td><strong>${m.month}</strong></td>
                  <td class="num-mono" style="text-align: right; color: var(--danger); font-weight: 700;">${m.count}</td>
                  <td>${m.age === '0' ? '<span class="kpi-badge badge-warning">0 (Neonate)</span>' : (m.age || '—')}</td>
                  <td><strong>${m.department}</strong></td>
                  <td>${m.circumstance}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>

        <div style="display: flex; justify-content: flex-end; gap: 10px;">
          <button class="btn btn-primary" onclick="HospitalApp.switchTab('mortality'); document.getElementById('closeDrilldownModalBtn').click();">
            <i class="fas fa-arrow-right"></i> Open Full Mortality Register Tab
          </button>
        </div>
      `;

      modal.classList.add('show');
      return;
    }

    // 2. Delivery Mode Drilldown
    if (type === 'delivery') {
      const isCS = deptName.toLowerCase().includes('caesarean') || deptName.toLowerCase().includes('cs');
      modalTitle.textContent = isCS ? 'Caesarean Section (CS) — Delivery Breakdown' : 'Spontaneous Vaginal Delivery (SVD) — Breakdown';
      const list = hospitalData.maternity.monthly;
      const months = list.map(m => m.month);
      const dataVals = list.map(m => isCS ? m.cs : m.svd);
      const totalDeliveries = list.reduce((s, m) => s + m.deliveries, 0);
      const totalVal = dataVals.reduce((s, v) => s + v, 0);
      const rate = totalDeliveries > 0 ? ((totalVal / totalDeliveries) * 100).toFixed(1) : 0;
      const peakIdx = dataVals.indexOf(Math.max(...dataVals));

      modalBody.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px;">
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Total ${isCS ? 'C-Sections' : 'SVD Deliveries'}</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: ${isCS ? '#E84A2D' : '#10B981'}; margin-top: 4px;">
              ${totalVal.toLocaleString()}
            </div>
          </div>
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Delivery Mode Share</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--primary); margin-top: 4px;">
              ${rate}%
            </div>
            <div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 2px;">of all ${totalDeliveries.toLocaleString()} facility deliveries</div>
          </div>
          <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Peak Delivery Month</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--text-primary); margin-top: 4px;">
              ${months[peakIdx]}
            </div>
            <div style="font-size: 0.72rem; color: var(--text-muted);">${Math.max(...dataVals).toLocaleString()} ${isCS ? 'C-Sections' : 'SVDs'}</div>
          </div>
        </div>

        <h4 style="margin-bottom: 12px; font-size: 1rem; color: var(--text-primary);">Monthly Progression (Jan - Aug 2026)</h4>
        <div style="height: 220px; width: 100%; position: relative;">
          <canvas id="drilldownChartCanvas"></canvas>
        </div>

        <div style="margin-top: 16px; display: flex; justify-content: space-between; align-items: center;">
          <span style="font-size: 0.85rem; color: var(--text-muted);">Source: Sheet 3 (BABIES TOTAL)</span>
          <button class="btn btn-primary" onclick="HospitalApp.switchTab('maternity'); document.getElementById('closeDrilldownModalBtn').click();">
            <i class="fas fa-arrow-right"></i> Open Full Maternity & Deliveries Tab
          </button>
        </div>
      `;

      modal.classList.add('show');

      setTimeout(() => {
        const canvas = document.getElementById('drilldownChartCanvas');
        if (canvas) {
          new Chart(canvas, {
            type: 'line',
            data: {
              labels: months,
              datasets: [{
                label: isCS ? 'Caesarean Section (CS)' : 'Spontaneous Vaginal Delivery (SVD)',
                data: dataVals,
                borderColor: isCS ? '#E84A2D' : '#10B981',
                backgroundColor: isCS ? 'rgba(232, 74, 45, 0.12)' : 'rgba(16, 185, 129, 0.12)',
                fill: true,
                tension: 0.35,
                borderWidth: 3,
                pointBackgroundColor: isCS ? '#E84A2D' : '#10B981',
                pointRadius: 5
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              scales: {
                y: { ticks: { callback: v => v.toLocaleString() } }
              },
              plugins: { legend: { display: false } }
            }
          });
        }
      }, 100);
      return;
    }

    // 3. Clinical Inpatient / Outpatient Department Drilldown
    modalTitle.textContent = `${deptName} — Monthly Breakdown`;

    const isOPD = type === 'outpatient';
    const dept = isOPD
      ? hospitalData.outpatient.find(d => d.name.toUpperCase() === deptName.toUpperCase() || deptName.toUpperCase().includes(d.name.toUpperCase()))
      : hospitalData.inpatient.find(w => w.name.toUpperCase() === deptName.toUpperCase() || deptName.toUpperCase().includes(w.name.toUpperCase()));

    if (!dept) return;

    const months = hospitalData.metadata.months;
    const totalHospital = isOPD ? HospitalAnalytics.getOPDTotal() : HospitalAnalytics.getIPDTotal();
    const sharePercent = totalHospital > 0 ? ((dept.total / totalHospital) * 100).toFixed(1) : 0;

    modalBody.innerHTML = `
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px;">
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Total Patients (Jan-Aug)</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: ${dept.color || 'var(--primary)'}; margin-top: 4px;">
            ${dept.total.toLocaleString()}
          </div>
        </div>
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">${isOPD ? 'Outpatient Share' : 'Inpatient Share'}</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: ${dept.color || 'var(--accent)'}; margin-top: 4px;">
            ${sharePercent}%
          </div>
          <div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 2px;">of all ${isOPD ? '64,105 OPD visits' : '3,120 IPD admissions'}</div>
        </div>
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Monthly Average</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--text-primary); margin-top: 4px;">
            ${Math.round(dept.total / 8).toLocaleString()}
          </div>
        </div>
      </div>

      <h4 style="margin-bottom: 12px; font-size: 1rem; color: var(--text-primary);">Monthly Patient Progression (Jan - Aug 2026)</h4>
      <div style="height: 220px; width: 100%; position: relative;">
        <canvas id="drilldownChartCanvas"></canvas>
      </div>

      <div style="margin-top: 16px; display: flex; justify-content: space-between; font-size: 0.82rem; color: var(--text-muted); background: var(--bg-surface-alt); padding: 10px 14px; border-radius: 6px;">
        <span>Service: <strong>${isOPD ? 'Outpatient Consultation (OPD)' : 'Inpatient Admission (IPD)'}</strong></span>
        <span>Peak Month: <strong>${months[dept.monthly.indexOf(Math.max(...dept.monthly))]} (${Math.max(...dept.monthly).toLocaleString()} pts)</strong></span>
      </div>
    `;

    modal.classList.add('show');

    setTimeout(() => {
      const canvas = document.getElementById('drilldownChartCanvas');
      if (canvas) {
        new Chart(canvas, {
          type: 'line',
          data: {
            labels: months,
            datasets: [{
              label: `${dept.name}`,
              data: dept.monthly,
              borderColor: dept.color || (isOPD ? '#2563EB' : '#EC4899'),
              backgroundColor: (dept.color || '#2563EB') + '1E',
              fill: true,
              tension: 0.35,
              borderWidth: 3,
              pointBackgroundColor: dept.color || (isOPD ? '#2563EB' : '#EC4899'),
              pointRadius: 5
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              y: {
                ticks: {
                  callback: v => v.toLocaleString()
                }
              }
            },
            plugins: {
              legend: { display: false }
            }
          }
        });
      }
    }, 100);
  },

  /**
   * Share Snapshot
   */
  shareSnapshot() {
    const opd = HospitalAnalytics.getOPDTotal();
    const ipd = HospitalAnalytics.getIPDTotal();
    const del = HospitalAnalytics.getDeliveriesTotal();
    const live = HospitalAnalytics.getLiveBirthsTotal();
    const cs = HospitalAnalytics.getCSStats();
    const deaths = HospitalAnalytics.getDeathsTotal();
    const mortRate = HospitalAnalytics.getMortalityRate();

    const text = `DMC Hospital Statistics Summary (Jan-Aug 2026):
• Outpatient Consultations (OPD): ${opd.toLocaleString()}
• Inpatient Admissions (IPD): ${ipd.toLocaleString()}
• Total Deliveries: ${del.toLocaleString()} (${live.toLocaleString()} Live Births)
• Caesarean Section Rate: ${cs.rate}% (${cs.cs.toLocaleString()} CS vs ${cs.svd.toLocaleString()} SVD)
• In-Facility Mortalities: ${deaths} Deaths (${mortRate}% mortality rate)
Report Source: ${hospitalData.metadata.sourceFile}`;

    navigator.clipboard.writeText(text).then(() => {
      ExcelDataParser.showToast("Summary copied to clipboard!", "success");
    }).catch(() => {
      ExcelDataParser.showToast("Summary generated.", "info");
    });
  }
};

window.showDepartmentDrilldown = (type, deptName) => HospitalApp.showDepartmentDrilldown(type, deptName);
