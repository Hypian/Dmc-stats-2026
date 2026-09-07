/**
 * DMC Hospital Statistics Dashboard - Application Controller (app.js)
 * Clean tab navigation, dynamic monthly tables, interactive drilldowns,
 * and multi-year dataset switching (2019-2024, 2025, and 2026).
 */

document.addEventListener('DOMContentLoaded', () => {
  HospitalApp.init();
});

const HospitalApp = {
  currentYear: '2026',
  currentFilter: {
    period: 'all',
    periodFilter: null,
    monthIndex: null,
    departmentId: 'all',
    activeTableTab: 'opd',
    currentTab: 'overview'
  },

  init() {
    this.initTheme();
    this.initClock();
    this.initYearSelector();
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
      HospitalCharts.updateAll(this.currentFilter.periodFilter);
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
   * Year Selector Manager
   */
  initYearSelector() {
    const yearPills = document.querySelectorAll('#yearSelectorPills .preset-pill');
    yearPills.forEach(btn => {
      btn.addEventListener('click', (e) => {
        const targetBtn = e.target.closest('.preset-pill');
        if (!targetBtn) return;
        const year = targetBtn.getAttribute('data-year');
        this.handleYearChange(year);
      });
    });
  },

  /**
   * Handle Year Selection (2026, 2025, or multi)
   */
  handleYearChange(yearKey) {
    this.currentYear = yearKey;

    // Update Year selector pills active state
    const yearPills = document.querySelectorAll('#yearSelectorPills .preset-pill');
    yearPills.forEach(p => {
      p.classList.toggle('active', p.getAttribute('data-year') === yearKey);
    });

    if (yearKey === 'multi') {
      this.switchTab('multiyear');
      const dateRangeEl = document.getElementById('headerDateRangeText');
      if (dateRangeEl) dateRangeEl.textContent = '2019 – 2026 Longitudinal Period';
      const sourceFileEl = document.getElementById('headerSourceFileText');
      if (sourceFileEl) sourceFileEl.textContent = 'Combined 2019-2024, 2025 & 2026 Workbooks';
      const grandTotalEl = document.getElementById('kpiGrandTotal');
      if (grandTotalEl) grandTotalEl.innerHTML = `<i class="fas fa-users-medical"></i> 528,639 Total Patients Served`;
      return;
    }

    if (this.currentFilter.currentTab === 'multiyear') {
      this.switchTab('overview');
    }

    HospitalAnalytics.setActiveYear(yearKey);

    // Update header meta
    const dateRangeEl = document.getElementById('headerDateRangeText');
    if (dateRangeEl) {
      dateRangeEl.textContent = hospitalData.metadata.periodBadge || hospitalData.metadata.periodCovered;
    }
    const sourceFileEl = document.getElementById('headerSourceFileText');
    if (sourceFileEl) {
      sourceFileEl.textContent = hospitalData.metadata.sourceFile;
    }
    const sidebarPeriodEl = document.getElementById('sidebarPeriodText');
    if (sidebarPeriodEl) {
      sidebarPeriodEl.textContent = `${hospitalData.metadata.year} Records`;
    }

    // Update Period Preset Buttons HTML
    const periodPillsContainer = document.getElementById('periodPresetPills');
    if (periodPillsContainer) {
      if (yearKey === '2025') {
        periodPillsContainer.innerHTML = `
          <button class="preset-pill active" data-period="all">All Months (Jan–Dec)</button>
          <button class="preset-pill" data-period="q1">Q1 (Jan–Mar)</button>
          <button class="preset-pill" data-period="q2">Q2 (Apr–Jun)</button>
          <button class="preset-pill" data-period="q3">Q3 (Jul–Sep)</button>
          <button class="preset-pill" data-period="q4">Q4 (Oct–Dec)</button>
        `;
      } else if (hospitalData.metadata.isAnnualOnly) {
        periodPillsContainer.innerHTML = `
          <button class="preset-pill active" data-period="all">Full Year ${yearKey} (Annual Register)</button>
        `;
      } else {
        periodPillsContainer.innerHTML = `
          <button class="preset-pill active" data-period="all">All Months (Jan–Aug)</button>
          <button class="preset-pill" data-period="q1">Q1 (Jan–Mar)</button>
          <button class="preset-pill" data-period="q2">Q2 (Apr–Jun)</button>
          <button class="preset-pill" data-period="jul">Jul</button>
          <button class="preset-pill" data-period="aug">Aug</button>
        `;
      }
      this.bindPeriodButtons();
    }

    // Update Single Month Filter Dropdown
    const monthSelect = document.getElementById('monthFilterSelect');
    if (monthSelect) {
      if (hospitalData.metadata.isAnnualOnly) {
        monthSelect.innerHTML = `<option value="all">Full Year ${yearKey} (Annual Total)</option>`;
      } else {
        let optHtml = `<option value="all">All Months (${hospitalData.metadata.months.length === 12 ? 'Jan–Dec' : 'Jan–Aug'})</option>`;
        hospitalData.metadata.months.forEach((m, idx) => {
          optHtml += `<option value="${idx}">${m} ${hospitalData.metadata.year}</option>`;
        });
        monthSelect.innerHTML = optHtml;
      }
    }

    // Update Alerts Banner
    const del = HospitalAnalytics.getDeliveriesTotal();
    const csStats = HospitalAnalytics.getCSStats();
    const live = HospitalAnalytics.getLiveBirthsTotal();
    const opdTot = HospitalAnalytics.getOPDTotal();
    const ipdTot = HospitalAnalytics.getIPDTotal();

    const alertMaternity = document.getElementById('alertTextMaternity');
    if (alertMaternity) {
      alertMaternity.innerHTML = `<strong>Obstetric Delivery Record (${hospitalData.metadata.year}):</strong> Total <strong>${del.toLocaleString()} deliveries</strong> (${csStats.cs.toLocaleString()} Caesarean Section [<strong>${csStats.rate}%</strong>] vs ${csStats.svd.toLocaleString()} Spontaneous Vaginal Delivery [<strong>${(100 - csStats.rate).toFixed(1)}%</strong>], with ${live.toLocaleString()} live births).`;
    }
    const alertSafety = document.getElementById('alertTextSafety');
    if (alertSafety) {
      if (yearKey === '2026') {
        alertSafety.innerHTML = `<strong>Facility Inpatient Safety:</strong> <strong>6 in-facility deaths</strong> recorded across 3,120 inpatient admissions (0.19% mortality rate): 4 stillbirths macerated in Maternity, 1 internal medicine, 1 emergency.`;
      } else {
        alertSafety.innerHTML = `<strong>Annual Clinical Inflow (${hospitalData.metadata.year}):</strong> <strong>${opdTot.toLocaleString()} Outpatient consultations</strong> and <strong>${ipdTot.toLocaleString()} Inpatient admissions</strong> recorded across all departments (${hospitalData.metadata.sourceFile}).`;
      }
    }

    // Update Section Subtitles
    const ipdDonutSub = document.getElementById('inpatientDonutSubtext');
    if (ipdDonutSub) ipdDonutSub.textContent = `Distribution of ${HospitalAnalytics.getIPDTotal().toLocaleString()} admissions across ${hospitalData.inpatient.length} wards (${hospitalData.metadata.sourceFile})`;
    const opdDonutSub = document.getElementById('outpatientDonutSubtext');
    if (opdDonutSub) opdDonutSub.textContent = `Volume distribution of ${HospitalAnalytics.getOPDTotal().toLocaleString()} patient visits (${hospitalData.metadata.sourceFile})`;
    const trajTitle = document.getElementById('trajectoryChartTitle');
    if (trajTitle) trajTitle.textContent = `Monthly Patient Inflow Trajectory (${hospitalData.metadata.periodCovered})`;
    const ipdRankTitle = document.getElementById('ipdRankingTitle');
    if (ipdRankTitle) ipdRankTitle.textContent = `Inpatient Department Volumes & Rankings (${hospitalData.metadata.periodCovered})`;
    const ipdRankSub = document.getElementById('ipdRankingSubtext');
    if (ipdRankSub) ipdRankSub.textContent = `Total admissions per specialized ward extracted directly from Sheet: IP (${hospitalData.metadata.sourceFile})`;
    const ipdBadge = document.getElementById('ipdBadgeRankingTotal');
    if (ipdBadge) ipdBadge.textContent = `Total: ${HospitalAnalytics.getIPDTotal().toLocaleString()} Admissions`;
    const ipdTableSub = document.getElementById('ipdTableSubtext');
    if (ipdTableSub) ipdTableSub.textContent = `Monthly admissions and totals for all ${hospitalData.inpatient.length} wards`;
    const opdRankTitle = document.getElementById('opdRankingTitle');
    if (opdRankTitle) opdRankTitle.textContent = `Outpatient Department Volume Rankings (${hospitalData.metadata.year})`;
    const opdRankSub = document.getElementById('opdRankingSubtext');
    if (opdRankSub) opdRankSub.textContent = `Total consultations across all 13 outpatient services (${HospitalAnalytics.getOPDTotal().toLocaleString()} visits)`;
    const opdBadge = document.getElementById('opdBadgeRankingTotal');
    if (opdBadge) opdBadge.textContent = `Total: ${HospitalAnalytics.getOPDTotal().toLocaleString()} Visits`;
    const matRatioTitle = document.getElementById('maternityDeliveryRatioTitle');
    if (matRatioTitle) matRatioTitle.textContent = `Maternity Delivery Mode Ratio (${hospitalData.metadata.year})`;
    const matBadge = document.getElementById('maternityDonutBadge');
    if (matBadge) matBadge.textContent = `${HospitalAnalytics.getDeliveriesTotal().toLocaleString()} Deliveries`;

    // Reset filters
    this.currentFilter.period = 'all';
    this.currentFilter.periodFilter = null;
    this.currentFilter.monthIndex = null;
    this.currentFilter.departmentId = 'all';

    // If currently on multiyear, switch back to overview
    if (this.currentFilter.currentTab === 'multiyear') {
      this.switchTab('overview');
    }

    this.populateDepartmentFilter();
    this.updateDashboardMetrics();
    HospitalCharts.updateAll(null);
    this.renderAllTables();
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
        if (tabId === 'multiyear') {
          this.handleYearChange('multi');
        } else {
          this.switchTab(tabId);
        }
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

    // Bind Period Preset Buttons
    this.bindPeriodButtons();

    // Single Month Selector Dropdown
    const monthFilterSelect = document.getElementById('monthFilterSelect');
    if (monthFilterSelect) {
      monthFilterSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        const presetBtns = document.querySelectorAll('#periodPresetPills .preset-pill');
        presetBtns.forEach(b => b.classList.remove('active'));

        if (val === 'all') {
          const allBtn = document.querySelector('#periodPresetPills .preset-pill[data-period="all"]');
          if (allBtn) allBtn.classList.add('active');
          this.handlePeriodChange('all');
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
    window.onHospitalDataUpdated = (yearToActivate) => {
      if (yearToActivate) {
        this.handleYearChange(yearToActivate);
      } else {
        this.populateDepartmentFilter();
        this.updateDashboardMetrics();
        HospitalCharts.updateAll(this.currentFilter.periodFilter);
        this.renderAllTables();
      }
      if (uploadModal) uploadModal.classList.remove('show');
    };
  },

  /**
   * Bind Period Preset Buttons
   */
  bindPeriodButtons() {
    const presetBtns = document.querySelectorAll('#periodPresetPills .preset-pill');
    presetBtns.forEach(btn => {
      btn.addEventListener('click', (e) => {
        presetBtns.forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        const period = e.target.dataset.period;
        this.handlePeriodChange(period);
      });
    });
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

    // Update Year selector pills if switching to/from multiyear
    if (tabId === 'multiyear') {
      const yearPills = document.querySelectorAll('#yearSelectorPills .preset-pill');
      yearPills.forEach(p => p.classList.toggle('active', p.getAttribute('data-year') === 'multi'));
    } else if (this.currentYear === 'multi') {
      this.handleYearChange('2026');
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
    }, 60);

    // Close mobile drawer if open
    const sidebar = document.querySelector('.app-sidebar');
    if (sidebar) sidebar.classList.remove('sidebar-open');
  },

  /**
   * Handle Period Filter (All, Q1, Q2, Q3, Q4, Jul, Aug, or individual month index)
   */
  handlePeriodChange(period) {
    this.currentFilter.period = String(period);
    const monthSelect = document.getElementById('monthFilterSelect');
    const numMonths = hospitalData.metadata.months.length;

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
    } else if (period === 'q3') {
      this.currentFilter.periodFilter = [6, 7, 8]; // Jul, Aug, Sep
      this.currentFilter.monthIndex = null;
      if (monthSelect) monthSelect.value = 'all';
    } else if (period === 'q4') {
      this.currentFilter.periodFilter = [9, 10, 11]; // Oct, Nov, Dec
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
      if (!isNaN(idx) && idx >= 0 && idx < numMonths) {
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

    select.innerHTML = `<option value="all">All Departments (${hospitalData.outpatient.length} OPD / ${hospitalData.inpatient.length} IPD)</option>`;

    const opdGroup = document.createElement('optgroup');
    opdGroup.label = 'Outpatient (OPD)';
    hospitalData.outpatient.forEach(d => {
      const opt = document.createElement('option');
      opt.value = `opd_${d.id}`;
      opt.textContent = `${d.name} (OPD)`;
      opdGroup.appendChild(opt);
    });
    select.appendChild(opdGroup);

    const ipdGroup = document.createElement('optgroup');
    ipdGroup.label = 'Inpatient (IPD)';
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
    this.renderMultiYearOPDTable();
    this.renderMultiYearIPDTable();
  },

  /**
   * Render Inpatient Table (Tab 2)
   */
  renderInpatientTable() {
    const table = document.getElementById('inpatientFullTable');
    if (!table) return;

    const isAnnual = hospitalData.metadata && hospitalData.metadata.isAnnualOnly;
    const data = hospitalData.inpatient;
    let totalSum = data.reduce((s, d) => s + d.total, 0);

    if (isAnnual) {
      let html = `
        <thead>
          <tr>
            <th style="min-width: 200px;">INPATIENT WARD</th>
            <th class="num-mono" style="text-align: right; font-weight: 700;">ANNUAL ADMISSIONS (${hospitalData.metadata.year})</th>
            <th class="num-mono" style="text-align: right; font-weight: 700;">HOSPITAL SHARE %</th>
            <th style="text-align: center;">ACTION</th>
          </tr>
        </thead>
        <tbody>
      `;
      data.forEach(dept => {
        const share = totalSum > 0 ? ((dept.total / totalSum) * 100).toFixed(1) : '0.0';
        html += `
          <tr>
            <td>
              <div style="display: flex; align-items: center; gap: 8px;">
                <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${dept.color};"></span>
                <strong style="color: var(--text-primary); font-weight: 600;">${dept.name}</strong>
              </div>
            </td>
            <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--primary); font-size: 0.95rem;">${dept.total.toLocaleString()}</td>
            <td class="num-mono" style="text-align: right; font-weight: 600;">${share}%</td>
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
            <td>GRAND TOTAL (${hospitalData.metadata.year})</td>
            <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--primary); font-weight: 800;">${totalSum.toLocaleString()}</td>
            <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
            <td>—</td>
          </tr>
        </tfoot>
      `;
      table.innerHTML = html;
      return;
    }

    const months = hospitalData.metadata.months;
    const monthlySums = months.map(() => 0);

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

    const isAnnual = hospitalData.metadata && hospitalData.metadata.isAnnualOnly;
    const data = hospitalData.outpatient;
    let totalSum = data.reduce((s, d) => s + d.total, 0);

    if (isAnnual) {
      let html = `
        <thead>
          <tr>
            <th style="min-width: 200px;">OUTPATIENT DEPARTMENT</th>
            <th class="num-mono" style="text-align: right; font-weight: 700;">ANNUAL CONSULTATIONS (${hospitalData.metadata.year})</th>
            <th class="num-mono" style="text-align: right; font-weight: 700;">HOSPITAL SHARE %</th>
            <th style="text-align: center;">ACTION</th>
          </tr>
        </thead>
        <tbody>
      `;
      data.forEach(dept => {
        const share = totalSum > 0 ? ((dept.total / totalSum) * 100).toFixed(1) : '0.0';
        html += `
          <tr>
            <td>
              <div style="display: flex; align-items: center; gap: 8px;">
                <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${dept.color};"></span>
                <strong style="color: var(--text-primary); font-weight: 600;">${dept.name}</strong>
              </div>
            </td>
            <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--secondary); font-size: 0.95rem;">${dept.total.toLocaleString()}</td>
            <td class="num-mono" style="text-align: right; font-weight: 600;">${share}%</td>
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
            <td>GRAND TOTAL (${hospitalData.metadata.year})</td>
            <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--secondary); font-weight: 800;">${totalSum.toLocaleString()}</td>
            <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
            <td>—</td>
          </tr>
        </tfoot>
      `;
      table.innerHTML = html;
      return;
    }

    const months = hospitalData.metadata.months;
    const monthlySums = months.map(() => 0);

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

    const data = hospitalData.maternity.monthly || [];
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

    if (data.length > 0) {
      data.forEach(m => {
        totDel += m.deliveries;
        totLive += m.liveBirths;
        totCS += m.cs;
        totSVD += m.svd;
        totDeath += m.deaths;
        const csRate = m.deliveries > 0 ? ((m.cs / m.deliveries) * 100).toFixed(1) : '0.0';

        html += `
          <tr>
            <td><strong>${m.month} ${hospitalData.metadata.year}</strong></td>
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
    } else if (hospitalData.maternity.annual) {
      const ann = hospitalData.maternity.annual;
      totDel = ann.deliveries;
      totLive = ann.liveBirths;
      totCS = ann.cs;
      totSVD = ann.svd;

      html += `
        <tr>
          <td><strong>Full Year ${hospitalData.metadata.year} (Annual Total)</strong></td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${ann.deliveries.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right;">${ann.liveBirths.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 600;">${ann.cs.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary);">${ann.svd.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${ann.csRate}%</td>
          <td class="num-mono" style="text-align: right;">0</td>
          <td style="font-size: 0.8rem; color: var(--text-muted);">Verified historical register (Sheet: BABIES TOTAL)</td>
        </tr>
      `;
    }

    const overallCSRate = totDel > 0 ? ((totCS / totDel) * 100).toFixed(1) : '0.0';

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td>TOTAL (${hospitalData.metadata.year})</td>
          <td class="num-mono" style="text-align: right; font-size: 1rem; color: var(--primary); font-weight: 800;">${totDel.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">${totLive.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 800;">${totCS.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary); font-weight: 800;">${totSVD.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">${overallCSRate}%</td>
          <td class="num-mono" style="text-align: right; font-weight: 800; color: ${totDeath > 0 ? 'var(--danger)' : 'inherit'};">${totDeath}</td>
          <td>—</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Maternity Historical Register (Tab 4)
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
      const csRate = item.deliveries > 0 ? ((item.cs / item.deliveries) * 100).toFixed(1) : '0.0';
      const isActive = item.year.startsWith(String(hospitalData.metadata.year));

      html += `
        <tr style="${isActive ? 'background-color: var(--primary-tint); font-weight: 600;' : ''}">
          <td>
            <strong style="color: ${isActive ? 'var(--primary)' : 'var(--text-primary)'};">
              ${item.year} ${isActive ? '<span style="background: var(--primary); color: #fff; font-size: 0.65rem; padding: 2px 6px; border-radius: 4px; margin-left: 6px;">Active</span>' : ''}
            </strong>
          </td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${item.deliveries.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right;">${item.liveBirths.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--accent); font-weight: 600;">${item.cs.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; color: var(--secondary);">${item.svd.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 700;">${csRate}%</td>
        </tr>
      `;
    });

    const cumCSRate = totDel > 0 ? ((totCS / totDel) * 100).toFixed(1) : '0.0';

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
   * Render Clinical Mortality Register (Tab 5)
   */
  renderMortalityTable() {
    const table = document.getElementById('mortalityLogTable');
    if (!table) return;

    const data = hospitalData.mortality || [];
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

    if (data.length > 0) {
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
    } else {
      html += `
        <tr>
          <td colspan="6" style="text-align: center; color: var(--text-muted); padding: 28px;">
            <i class="fas fa-shield-check" style="color: #10B981; font-size: 1.4rem; margin-bottom: 8px; display: block;"></i>
            No clinical in-facility deaths were recorded in the ${hospitalData.metadata.year} workbook register.
          </td>
        </tr>
      `;
    }

    const ipdTotal = HospitalAnalytics.getIPDTotal();
    const mortRate = ipdTotal > 0 ? ((totalDeaths / ipdTotal) * 100).toFixed(2) : '0.00';

    html += `
      </tbody>
      <tfoot>
        <tr>
          <td colspan="2">TOTAL IN-FACILITY DEATHS (${hospitalData.metadata.year})</td>
          <td class="num-mono" style="text-align: right; font-weight: 800; color: ${totalDeaths > 0 ? 'var(--danger)' : 'var(--text-muted)'}; font-size: 1.05rem;">${totalDeaths}</td>
          <td colspan="3" style="color: var(--text-muted); font-size: 0.82rem;">${mortRate}% mortality rate across ${ipdTotal.toLocaleString()} inpatient admissions (${hospitalData.metadata.sourceFile})</td>
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
    const isAnnual = hospitalData.metadata && hospitalData.metadata.isAnnualOnly;
    const grandTotal = isOPD ? HospitalAnalytics.getOPDTotal() : HospitalAnalytics.getIPDTotal();

    if (isAnnual) {
      thead.innerHTML = `
        <tr>
          <th scope="col" style="min-width: 200px;">Department (${isOPD ? 'OUTPATIENT' : 'INPATIENT'})</th>
          <th scope="col" class="num-mono" style="text-align: right; font-weight: 700;">Annual Total (${hospitalData.metadata.year})</th>
          <th scope="col" class="num-mono" style="text-align: right;">Share %</th>
          <th scope="col" style="text-align: center;">Action</th>
        </tr>
      `;
      let rowsHtml = '';
      dataset.forEach(item => {
        const share = grandTotal > 0 ? ((item.total / grandTotal) * 100).toFixed(1) : 0;
        rowsHtml += `
          <tr>
            <td>
              <div style="display: flex; align-items: center; gap: 8px;">
                <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${item.color || '#E84A2D'};"></span>
                <strong>${item.name}</strong>
              </div>
            </td>
            <td class="num-mono" style="text-align: right; font-weight: 700; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'}; font-size: 0.95rem;">
              ${item.total.toLocaleString()}
            </td>
            <td class="num-mono" style="text-align: right; font-weight: 600;">
              ${share}%
            </td>
            <td style="text-align: center;">
              <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('${isOPD ? 'outpatient' : 'inpatient'}', '${item.name}')">
                <i class="fas fa-chart-line"></i> Inspect
              </button>
            </td>
          </tr>
        `;
      });
      tbody.innerHTML = rowsHtml;
      if (tfoot) {
        tfoot.innerHTML = `
          <tr>
            <td style="font-weight: 700;">GRAND TOTAL (${hospitalData.metadata.year})</td>
            <td class="num-mono" style="text-align: right; font-size: 1.05rem; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'}; font-weight: 800;">
              ${grandTotal.toLocaleString()}
            </td>
            <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
            <td>—</td>
          </tr>
        `;
      }
      return;
    }

    const months = hospitalData.metadata.months;

    thead.innerHTML = `
      <tr>
        <th scope="col" style="min-width: 180px;">Department (${isOPD ? 'OUTPATIENT' : 'INPATIENT'})</th>
        ${months.map(m => `<th scope="col" class="num-mono" style="text-align: right;">${m.toUpperCase()}</th>`).join('')}
        <th scope="col" class="num-mono" style="text-align: right; font-weight: 700;">Total</th>
        <th scope="col" class="num-mono" style="text-align: right;">Share %</th>
        <th scope="col" style="text-align: center;">Action</th>
      </tr>
    `;

    const monthlySums = months.map(() => 0);
    let rowsHtml = '';
    dataset.forEach(item => {
      const share = grandTotal > 0 ? ((item.total / grandTotal) * 100).toFixed(1) : 0;
      months.forEach((_, i) => monthlySums[i] += (item.monthly[i] || 0));

      rowsHtml += `
        <tr>
          <td>
            <div style="display: flex; align-items: center; gap: 8px;">
              <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${item.color || '#E84A2D'};"></span>
              <strong>${item.name}</strong>
            </div>
          </td>
          ${item.monthly.map(v => `<td class="num-mono" style="text-align: right;">${v ? v.toLocaleString() : '0'}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 700; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'};">
            ${item.total.toLocaleString()}
          </td>
          <td class="num-mono" style="text-align: right; font-weight: 600;">
            ${share}%
          </td>
          <td style="text-align: center;">
            <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('${isOPD ? 'outpatient' : 'inpatient'}', '${item.name}')">
              <i class="fas fa-chart-line"></i> Inspect
            </button>
          </td>
        </tr>
      `;
    });

    tbody.innerHTML = rowsHtml;

    if (tfoot) {
      tfoot.innerHTML = `
        <tr>
          <td style="font-weight: 700;">GRAND TOTAL</td>
          ${monthlySums.map(s => `<td class="num-mono" style="text-align: right; font-weight: 700;">${s.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-size: 1.05rem; color: ${isOPD ? 'var(--secondary)' : 'var(--primary)'}; font-weight: 800;">
            ${grandTotal.toLocaleString()}
          </td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
          <td>—</td>
        </tr>
      `;
    }
  },

  /**
   * Render Multi-Year Outpatient Table (Tab 7)
   */
  renderMultiYearOPDTable() {
    const table = document.getElementById('multiYearOPDTable');
    if (!table) return;

    const hist = HISTORICAL_MULTI_YEAR;
    const years = hist.years;
    const depts = hist.opdDepartments;
    const grandTotal = hist.cumulativeSummary.totalOPD;

    let html = `
      <thead>
        <tr>
          <th style="min-width: 180px;">OUTPATIENT SPECIALTY</th>
          ${years.map(y => `<th class="num-mono" style="text-align: right;">${y}</th>`).join('')}
          <th class="num-mono" style="text-align: right; font-weight: 700; color: var(--primary);">8-YR TOTAL</th>
          <th class="num-mono" style="text-align: right; font-weight: 700;">SHARE %</th>
          <th style="text-align: center;">ACTION</th>
        </tr>
      </thead>
      <tbody>
    `;

    depts.forEach(dept => {
      const pct = grandTotal > 0 ? ((dept.total / grandTotal) * 100).toFixed(1) : '0.0';
      html += `
        <tr>
          <td>
            <div style="display: flex; align-items: center; gap: 8px;">
              <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${dept.color};"></span>
              <strong>${dept.name}</strong>
            </div>
          </td>
          ${dept.years.map(v => `<td class="num-mono" style="text-align: right;">${v.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--primary); font-size: 0.95rem;">${dept.total.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 600;">${pct}%</td>
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
          <td>ANNUAL OPD GRAND TOTAL</td>
          ${hist.annualTotals.opd.map(t => `<td class="num-mono" style="text-align: right; font-weight: 700;">${t.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-size: 1.05rem; color: var(--primary); font-weight: 800;">${grandTotal.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
          <td>—</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Render Multi-Year Inpatient Table (Tab 7)
   */
  renderMultiYearIPDTable() {
    const table = document.getElementById('multiYearIPDTable');
    if (!table) return;

    const hist = HISTORICAL_MULTI_YEAR;
    const years = hist.years;
    const depts = hist.ipdDepartments;
    const grandTotal = hist.cumulativeSummary.totalIPD;

    let html = `
      <thead>
        <tr>
          <th style="min-width: 180px;">INPATIENT WARD</th>
          ${years.map(y => `<th class="num-mono" style="text-align: right;">${y}</th>`).join('')}
          <th class="num-mono" style="text-align: right; font-weight: 700; color: var(--secondary);">8-YR TOTAL</th>
          <th class="num-mono" style="text-align: right; font-weight: 700;">SHARE %</th>
          <th style="text-align: center;">ACTION</th>
        </tr>
      </thead>
      <tbody>
    `;

    depts.forEach(ward => {
      const pct = grandTotal > 0 ? ((ward.total / grandTotal) * 100).toFixed(1) : '0.0';
      html += `
        <tr>
          <td>
            <div style="display: flex; align-items: center; gap: 8px;">
              <span style="width: 10px; height: 10px; border-radius: 50%; background-color: ${ward.color};"></span>
              <strong>${ward.name}</strong>
            </div>
          </td>
          ${ward.years.map(v => `<td class="num-mono" style="text-align: right;">${v.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-weight: 700; color: var(--secondary); font-size: 0.95rem;">${ward.total.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 600;">${pct}%</td>
          <td style="text-align: center;">
            <button class="btn-chart-action" onclick="HospitalApp.showDepartmentDrilldown('inpatient', '${ward.name}')">
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
          <td>ANNUAL IPD GRAND TOTAL</td>
          ${hist.annualTotals.ipd.map(t => `<td class="num-mono" style="text-align: right; font-weight: 700;">${t.toLocaleString()}</td>`).join('')}
          <td class="num-mono" style="text-align: right; font-size: 1.05rem; color: var(--secondary); font-weight: 800;">${grandTotal.toLocaleString()}</td>
          <td class="num-mono" style="text-align: right; font-weight: 800;">100.0%</td>
          <td>—</td>
        </tr>
      </tfoot>
    `;

    table.innerHTML = html;
  },

  /**
   * Department Performance Drilldown Modal
   */
  showDepartmentDrilldown(type, deptName) {
    const modal = document.getElementById('drilldownModal');
    const modalTitle = document.getElementById('drilldownModalTitle');
    const modalBody = document.getElementById('drilldownModalBody');
    if (!modal || !modalTitle || !modalBody) return;

    // 1. Mortality Incident Drilldown
    if (type === 'mortality') {
      modalTitle.textContent = `${deptName} — Clinical Mortalities Detail`;
      const mortList = hospitalData.mortality || [];
      const matching = mortList.filter(m => {
        const full = `${m.department} (${m.circumstance})`;
        return full.toLowerCase().includes(deptName.toLowerCase()) || deptName.toLowerCase().includes(m.department.toLowerCase());
      });
      const deathsCount = matching.reduce((s, m) => s + m.count, 0);
      const totalFacilityDeaths = mortList.reduce((s, m) => s + m.count, 0);
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
      const list = hospitalData.maternity.monthly || [];
      const hasMonthly = list.length > 0;
      const months = hasMonthly ? list.map(m => m.month) : HISTORICAL_MULTI_YEAR.years;
      const dataVals = hasMonthly ? list.map(m => isCS ? m.cs : m.svd) : (isCS ? HISTORICAL_MULTI_YEAR.annualTotals.cs : HISTORICAL_MULTI_YEAR.annualTotals.svd);
      const totalDeliveries = hasMonthly ? list.reduce((s, m) => s + m.deliveries, 0) : HISTORICAL_MULTI_YEAR.cumulativeSummary.totalDeliveries;
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
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Peak Period</div>
            <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--text-primary); margin-top: 4px;">
              ${months[peakIdx]}
            </div>
            <div style="font-size: 0.72rem; color: var(--text-muted);">${Math.max(...dataVals).toLocaleString()} ${isCS ? 'C-Sections' : 'SVDs'}</div>
          </div>
        </div>

        <h4 style="margin-bottom: 12px; font-size: 1rem; color: var(--text-primary);">Progression (${hasMonthly ? hospitalData.metadata.periodCovered : '2019 – 2026'})</h4>
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
    const avgMonthly = Math.round(dept.total / months.length);

    modalBody.innerHTML = `
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px;">
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Total Patients (${hospitalData.metadata.periodBadge || hospitalData.metadata.year})</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: ${dept.color || 'var(--primary)'}; margin-top: 4px;">
            ${dept.total.toLocaleString()}
          </div>
        </div>
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">${isOPD ? 'Outpatient Share' : 'Inpatient Share'}</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: ${dept.color || 'var(--accent)'}; margin-top: 4px;">
            ${sharePercent}%
          </div>
          <div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 2px;">of all ${totalHospital.toLocaleString()} ${isOPD ? 'OPD visits' : 'IPD admissions'}</div>
        </div>
        <div style="padding: 14px; background: var(--bg-surface-alt); border-radius: 8px; border: 1px solid var(--border-light);">
          <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">Monthly Average</div>
          <div style="font-size: 1.8rem; font-weight: 700; font-family: var(--font-mono); color: var(--text-primary); margin-top: 4px;">
            ${avgMonthly.toLocaleString()}
          </div>
        </div>
      </div>

      <h4 style="margin-bottom: 12px; font-size: 1rem; color: var(--text-primary);">Monthly Patient Progression (${hospitalData.metadata.periodCovered || hospitalData.metadata.year})</h4>
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
    const data = {
      title: 'DMC Hospital Statistics Report',
      text: `Hospital Statistics (${hospitalData.metadata.periodCovered}): ${HospitalAnalytics.getOPDTotal().toLocaleString()} OPD Visits, ${HospitalAnalytics.getIPDTotal().toLocaleString()} IPD Admissions, ${HospitalAnalytics.getDeliveriesTotal().toLocaleString()} Deliveries.`,
      url: window.location.href
    };

    if (navigator.share) {
      navigator.share(data).catch(() => {});
    } else {
      navigator.clipboard.writeText(`${data.title}\n${data.text}\n${data.url}`).then(() => {
        if (typeof ExcelDataParser !== 'undefined' && ExcelDataParser.showToast) {
          ExcelDataParser.showToast('Report summary copied to clipboard!', 'success');
        } else {
          alert('Report summary copied to clipboard!');
        }
      });
    }
  }
};

window.HospitalApp = HospitalApp;
window.HospitalAnalytics = HospitalAnalytics;
window.showDepartmentDrilldown = (type, name) => HospitalApp.showDepartmentDrilldown(type, name);
