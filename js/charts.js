/**
 * DMC Hospital Statistics - Chart.js Visualization Engine (charts.js)
 * Clean, readable data visualizations directly mapping the 4 Excel sheets:
 * 1. Inpatient Admissions by Department (Donut)
 * 2. Outpatient Consultations by Department (Donut)
 * 3. Maternity Delivery Mode: CS vs SVD (Donut)
 * 4. In-Facility Deaths by Department & Cause (Donut)
 * 5. Monthly Trend Line: Outpatient vs Inpatient (Jan - Aug 2026)
 * 6. Historical Deliveries Evolution (2019 - 2026)
 * 7. Department Comparison Horizontal Bars (IPD and OPD)
 */

const HospitalCharts = {
  instances: {},

  // DMC Dream Red & White Healthcare Color Palette
  colors: {
    primary: '#E84A2D',
    secondary: '#3D3532',
    accent: '#FA8974',
    danger: '#DC2626',
    info: '#71645B'
  },

  /**
   * Register custom plugin for Centered Inner Text on Donut Charts
   */
  registerPlugins() {
    if (typeof Chart === 'undefined') return;

    const centerTextPlugin = {
      id: 'donutCenterText',
      afterDraw(chart) {
        if (chart.config.type !== 'doughnut' || !chart.config.options.plugins.donutCenterText) return;

        const centerConfig = chart.config.options.plugins.donutCenterText;
        if (!centerConfig.display) return;

        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        if (!meta || !meta.data || !meta.data.length) return;

        const x = meta.data[0].x;
        const y = meta.data[0].y;

        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

        // Main Metric Number
        ctx.font = `bold ${centerConfig.valueFontSize || '26px'} 'JetBrains Mono', monospace`;
        ctx.fillStyle = centerConfig.valueColor || (isDark ? '#FF5E42' : '#E84A2D');
        ctx.fillText(centerConfig.value || '', x, y - (centerConfig.subtext ? 10 : 0));

        // Subtext Label
        if (centerConfig.subtext) {
          ctx.font = `700 ${centerConfig.labelFontSize || '11px'} 'Poppins', sans-serif`;
          ctx.fillStyle = centerConfig.labelColor || (isDark ? '#B8ABA5' : '#8C7D77');
          ctx.fillText(centerConfig.subtext.toUpperCase(), x, y + 16);
        }

        ctx.restore();
      }
    };

    Chart.register(centerTextPlugin);
  },

  /**
   * Initialize all visualizations
   */
  initAll() {
    this.registerPlugins();
    this.renderInpatientDonut();
    this.renderOutpatientDonut();
    this.renderDeliveryModeDonut();
    this.renderMortalityDonut();
    this.renderAdmissionsTrendChart();
    this.renderHistoricalDeliveriesChart();
    this.renderDepartmentComparisonBar('ipd');
    this.renderOPDRankingBar();
    this.renderSparklines();
    this.renderMultiYearTrajectoryChart();
    this.renderMultiYearDeliveriesChart();
    this.renderMultiYearDeptRankingChart();
  },

  /**
   * 1. Inpatient Admissions by Department Donut
   */
  renderInpatientDonut(periodFilter) {
    const canvas = document.getElementById('inpatientDonutChart');
    if (!canvas) return;

    if (this.instances.ipdDonut) {
      this.instances.ipdDonut.destroy();
    }

    const labels = hospitalData.inpatient.map(w => w.name);
    const dataVals = hospitalData.inpatient.map(w => {
      if (Array.isArray(periodFilter)) {
        return periodFilter.reduce((s, idx) => s + (w.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return w.monthly[periodFilter] || 0;
      }
      return w.total;
    });

    const totalIPD = dataVals.reduce((a, b) => a + b, 0);
    const colors = hospitalData.inpatient.map(w => w.color);
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.ipdDonut = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataVals,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: isDark ? '#1E1A18' : '#FFFFFF',
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: '66%',
        animation: { duration: 800 },
        onHover: (evt, el) => {
          if (evt.native && evt.native.target) {
            evt.native.target.style.cursor = el.length ? 'pointer' : 'default';
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                const val = context.raw;
                const pct = totalIPD > 0 ? ((val / totalIPD) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${val.toLocaleString()} admissions (${pct}%)`;
              }
            }
          },
          donutCenterText: {
            display: true,
            value: totalIPD.toLocaleString(),
            subtext: 'Inpatients',
            valueColor: isDark ? '#FF5E42' : '#E84A2D'
          }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('inpatient', labels[index]);
            }
          }
        }
      }
    });

    this.renderCustomLegend('inpatientLegend', labels, dataVals, colors, totalIPD, 'inpatient');
  },

  /**
   * 2. Outpatient Consultations by Department Donut
   */
  renderOutpatientDonut(periodFilter) {
    const canvas = document.getElementById('outpatientDonutChart');
    if (!canvas) return;

    if (this.instances.opdDonut) {
      this.instances.opdDonut.destroy();
    }

    const getVal = (d) => {
      if (Array.isArray(periodFilter)) {
        return periodFilter.reduce((s, idx) => s + (d.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return d.monthly[periodFilter] || 0;
      }
      return d.total;
    };

    const sorted = [...hospitalData.outpatient].sort((a, b) => getVal(b) - getVal(a));

    const labels = sorted.map(d => d.name);
    const dataVals = sorted.map(d => getVal(d));
    const colors = sorted.map(d => d.color);

    const totalOPD = dataVals.reduce((a, b) => a + b, 0);
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.opdDonut = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataVals,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: isDark ? '#1E1A18' : '#FFFFFF',
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: '66%',
        animation: { duration: 800 },
        onHover: (evt, el) => {
          if (evt.native && evt.native.target) {
            evt.native.target.style.cursor = el.length ? 'pointer' : 'default';
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                const val = context.raw;
                const pct = totalOPD > 0 ? ((val / totalOPD) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${val.toLocaleString()} visits (${pct}%)`;
              }
            }
          },
          donutCenterText: {
            display: true,
            value: totalOPD >= 1000 ? `${(totalOPD / 1000).toFixed(1)}k` : totalOPD.toLocaleString(),
            subtext: 'Outpatients',
            valueColor: isDark ? '#FF5E42' : '#E84A2D'
          }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('outpatient', labels[index]);
            }
          }
        }
      }
    });

    this.renderCustomLegend('outpatientLegend', labels, dataVals, colors, totalOPD, 'outpatient');
  },

  /**
   * 3. Maternity Delivery Mode Donut
   */
  renderDeliveryModeDonut(periodFilter) {
    const canvas = document.getElementById('deliveryModeDonutChart');
    if (!canvas) return;

    if (this.instances.deliveryDonut) {
      this.instances.deliveryDonut.destroy();
    }

    const stats = HospitalAnalytics.getCSStats(periodFilter);
    const labels = ['Caesarean Section (CS)', 'Spontaneous Vaginal Delivery (SVD)'];
    const dataVals = [stats.cs, stats.svd];
    const colors = ['#E84A2D', '#10B981'];
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.deliveryDonut = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataVals,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: isDark ? '#1E1A18' : '#FFFFFF',
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: '66%',
        animation: { duration: 800 },
        onHover: (evt, el) => {
          if (evt.native && evt.native.target) {
            evt.native.target.style.cursor = el.length ? 'pointer' : 'default';
          }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('delivery', labels[index]);
            }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                const pct = stats.total > 0 ? ((context.raw / stats.total) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${context.raw.toLocaleString()} (${pct}%)`;
              }
            }
          },
          donutCenterText: {
            display: true,
            value: `${stats.rate}%`,
            subtext: 'C-Section Rate',
            valueColor: isDark ? '#FF5E42' : '#E84A2D'
          }
        }
      }
    });

    this.renderCustomLegend('deliveryModeLegend', labels, dataVals, colors, stats.total, 'delivery');
  },

  /**
   * 4. In-Facility Deaths Donut (Dynamically mapped from Sheet 4: Death)
   */
  renderMortalityDonut(periodFilter) {
    const canvas = document.getElementById('mortalityDonutChart');
    if (!canvas) return;

    if (this.instances.mortalityDonut) {
      this.instances.mortalityDonut.destroy();
    }

    const months = hospitalData.metadata.months;
    let filteredDeaths = hospitalData.mortality;

    if (Array.isArray(periodFilter)) {
      const targetMonths = periodFilter.map(i => months[i]);
      filteredDeaths = hospitalData.mortality.filter(d => targetMonths.includes(d.month));
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      const targetMonth = months[periodFilter];
      filteredDeaths = hospitalData.mortality.filter(d => d.month === targetMonth);
    }

    const groupMap = {};
    filteredDeaths.forEach(d => {
      const key = `${d.department} (${d.circumstance})`;
      groupMap[key] = (groupMap[key] || 0) + d.count;
    });

    const hasData = Object.keys(groupMap).length > 0;
    const labels = hasData ? Object.keys(groupMap) : ['No In-Facility Mortalities'];
    const dataVals = hasData ? Object.values(groupMap) : [1];
    const totalDeaths = filteredDeaths.reduce((sum, d) => sum + d.count, 0);

    const colorMap = {
      'Maternity (Still births macerated)': '#E84A2D',
      'Internal medecine (Natural disease ( cancer))': '#3D3532',
      'Emergency (Natural disease)': '#71645B'
    };
    const colors = hasData ? labels.map(l => colorMap[l] || '#FA8974') : ['#E84A2D'];
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.mortalityDonut = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataVals,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: isDark ? '#1E1A18' : '#FFFFFF',
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: '66%',
        animation: { duration: 800 },
        onHover: (evt, el) => {
          if (evt.native && evt.native.target) {
            evt.native.target.style.cursor = (el.length && hasData) ? 'pointer' : 'default';
          }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0 && hasData) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('mortality', labels[index]);
            }
          }
        },
        plugins: {
          legend: { display: false },
          donutCenterText: {
            display: true,
            value: totalDeaths.toString(),
            subtext: totalDeaths === 1 ? 'Death' : 'Deaths',
            valueColor: totalDeaths > 0 ? (isDark ? '#FF5E42' : '#E84A2D') : '#16A34A'
          }
        }
      }
    });

    this.renderCustomLegend('mortalityLegend', labels, hasData ? dataVals : [0], colors, totalDeaths, 'mortality');
  },

  /**
   * 5. Monthly Admissions Trend Chart
   */
  renderAdmissionsTrendChart() {
    const canvas = document.getElementById('patientTrendLineChart');
    if (!canvas) return;

    if (this.instances.trendLine) {
      this.instances.trendLine.destroy();
    }

    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    const isAnnual = hospitalData.metadata && hospitalData.metadata.isAnnualOnly;

    let chartLabels, opdData, ipdData;
    let pointRadiusOPD = 5, pointRadiusIPD = 5;
    let hist = null, activeIdx = -1;

    if (isAnnual) {
      hist = HISTORICAL_MULTI_YEAR;
      chartLabels = hist.years;
      opdData = hist.annualTotals.opd;
      ipdData = hist.annualTotals.ipd;
      activeIdx = hist.years.findIndex(y => y.startsWith(String(hospitalData.metadata.year)));
      pointRadiusOPD = hist.years.map((_, i) => i === activeIdx ? 9 : 4);
      pointRadiusIPD = hist.years.map((_, i) => i === activeIdx ? 9 : 4);
    } else {
      chartLabels = hospitalData.metadata.months;
      opdData = chartLabels.map((_, i) => HospitalAnalytics.getOPDTotal(i));
      ipdData = chartLabels.map((_, i) => HospitalAnalytics.getIPDTotal(i));
    }

    const trajTitle = document.getElementById('trajectoryChartTitle');
    if (trajTitle) {
      trajTitle.textContent = isAnnual 
        ? `Longitudinal Patient Trajectory (2019–2026 · Highlighted: ${hospitalData.metadata.year})`
        : `Monthly Patient Inflow Trajectory (${hospitalData.metadata.periodCovered})`;
    }

    this.instances.trendLine = new Chart(canvas, {
      type: isAnnual ? 'bar' : 'line',
      data: {
        labels: chartLabels,
        datasets: isAnnual ? [
          {
            type: 'bar',
            label: 'Outpatients (OPD)',
            data: opdData,
            backgroundColor: hist.years.map((_, i) => i === activeIdx ? '#E84A2D' : (isDark ? 'rgba(232, 74, 45, 0.35)' : 'rgba(232, 74, 45, 0.4)')),
            borderColor: hist.years.map((_, i) => i === activeIdx ? '#C0392B' : 'transparent'),
            borderWidth: hist.years.map((_, i) => i === activeIdx ? 2 : 0),
            borderRadius: 4,
            barPercentage: 0.9,
            categoryPercentage: 0.8,
            yAxisID: 'yOPD'
          },
          {
            type: 'bar',
            label: 'Inpatients (IPD)',
            data: ipdData,
            backgroundColor: hist.years.map((_, i) => i === activeIdx ? '#2563EB' : (isDark ? 'rgba(37, 99, 235, 0.35)' : 'rgba(37, 99, 235, 0.4)')),
            borderColor: hist.years.map((_, i) => i === activeIdx ? '#1D4ED8' : 'transparent'),
            borderWidth: hist.years.map((_, i) => i === activeIdx ? 2 : 0),
            borderRadius: 4,
            barPercentage: 0.9,
            categoryPercentage: 0.8,
            yAxisID: 'yIPD'
          }
        ] : [
          {
            label: 'Outpatients (OPD)',
            data: opdData,
            borderColor: '#E84A2D',
            backgroundColor: 'rgba(232, 74, 45, 0.08)',
            borderWidth: 3,
            fill: true,
            tension: 0.35,
            pointBackgroundColor: '#E84A2D',
            pointBorderColor: '#E84A2D',
            pointBorderWidth: 1,
            pointRadius: pointRadiusOPD,
            yAxisID: 'yOPD'
          },
          {
            label: 'Inpatients (IPD)',
            data: ipdData,
            borderColor: '#2563EB',
            backgroundColor: 'rgba(37, 99, 235, 0.08)',
            borderWidth: 3,
            fill: false,
            tension: 0.35,
            pointBackgroundColor: '#2563EB',
            pointBorderColor: '#2563EB',
            pointBorderWidth: 1,
            pointRadius: pointRadiusIPD,
            yAxisID: 'yIPD'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { family: 'Poppins', weight: '600' } }
          },
          yOPD: {
            type: 'linear',
            position: 'left',
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.05)' },
            ticks: {
              color: isDark ? '#FF5E42' : '#E84A2D',
              font: { family: 'JetBrains Mono', weight: '600' },
              callback: val => val.toLocaleString()
            },
            title: { display: true, text: 'OPD Consultations', color: isDark ? '#FF5E42' : '#E84A2D', font: { weight: 'bold' } }
          },
          yIPD: {
            type: 'linear',
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: {
              color: isDark ? '#93C5FD' : '#2563EB',
              font: { family: 'JetBrains Mono', weight: '600' }
            },
            title: { display: true, text: 'IPD Admissions', color: isDark ? '#93C5FD' : '#2563EB', font: { weight: 'bold' } }
          }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              usePointStyle: !isAnnual,
              boxWidth: isAnnual ? 14 : 10,
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          }
        }
      }
    });
  },

  /**
   * 6. Multi-Year Historical Deliveries Chart
   */
  renderHistoricalDeliveriesChart() {
    const canvas = document.getElementById('historicalDeliveriesChart');
    if (!canvas) return;

    if (this.instances.historyChart) {
      this.instances.historyChart.destroy();
    }

    const hist = hospitalData.maternity.historical;
    const labels = hist.map(h => h.year);
    const csVals = hist.map(h => h.cs);
    const svdVals = hist.map(h => h.svd);
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.historyChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'C-Section (CS)',
            data: csVals,
            backgroundColor: '#E84A2D',
            borderRadius: 4
          },
          {
            label: 'Spontaneous Vaginal (SVD)',
            data: svdVals,
            backgroundColor: '#10B981',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            stacked: true,
            grid: { display: false },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { weight: '600' } }
          },
          y: {
            stacked: true,
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.05)' },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { family: 'JetBrains Mono' } }
          }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              usePointStyle: true,
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          }
        }
      }
    });
  },

  /**
   * 7. Department Comparison Horizontal Bar Chart
   */
  renderDepartmentComparisonBar(metric = 'ipd', periodFilter) {
    const canvas = document.getElementById('deptComparisonBarChart');
    if (!canvas) return;

    if (this.instances.deptBar) {
      this.instances.deptBar.destroy();
    }

    const getVal = (d) => {
      if (Array.isArray(periodFilter)) {
        return periodFilter.reduce((s, idx) => s + (d.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return d.monthly[periodFilter] || 0;
      }
      return d.total;
    };

    const items = [...hospitalData.inpatient].sort((a, b) => getVal(b) - getVal(a));
    const labels = items.map(d => d.name);
    const dataVals = items.map(d => getVal(d));
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.deptBar = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: periodFilter ? 'Inpatients (Filtered Period)' : `Total Inpatients (${hospitalData.metadata.periodCovered || hospitalData.metadata.year})`,
          data: dataVals,
          backgroundColor: items.map(d => d.color || '#EC4899'),
          hoverBackgroundColor: items.map(d => d.color || '#DB2777'),
          borderRadius: 6,
          maxBarThickness: 24
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)' },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { family: 'JetBrains Mono' } }
          },
          y: {
            grid: { display: false },
            ticks: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          }
        },
        plugins: {
          legend: { display: false }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('inpatient', labels[index]);
            }
          }
        }
      }
    });
  },

  /**
   * 8. Outpatient Ranking Horizontal Bar
   */
  renderOPDRankingBar(periodFilter) {
    const canvas = document.getElementById('opdRankingBarChart');
    if (!canvas) return;

    if (this.instances.opdBar) {
      this.instances.opdBar.destroy();
    }

    const getVal = (d) => {
      if (Array.isArray(periodFilter)) {
        return periodFilter.reduce((s, idx) => s + (d.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return d.monthly[periodFilter] || 0;
      }
      return d.total;
    };

    const items = [...hospitalData.outpatient].sort((a, b) => getVal(b) - getVal(a));
    const labels = items.map(d => d.name);
    const dataVals = items.map(d => getVal(d));
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.opdBar = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: periodFilter ? 'Outpatients (Filtered Period)' : `Total Outpatients (${hospitalData.metadata.periodCovered || hospitalData.metadata.year})`,
          data: dataVals,
          backgroundColor: items.map(d => d.color || '#2563EB'),
          hoverBackgroundColor: items.map(d => d.color || '#1D4ED8'),
          borderRadius: 6,
          maxBarThickness: 22
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)' },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { family: 'JetBrains Mono' } }
          },
          y: {
            grid: { display: false },
            ticks: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          }
        },
        plugins: {
          legend: { display: false }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            if (typeof window.showDepartmentDrilldown === 'function') {
              window.showDepartmentDrilldown('outpatient', labels[index]);
            }
          }
        }
      }
    });
  },

  /**
   * 9. Mini Sparklines in KPI Cards
   */
  renderSparklines() {
    this.createSparkline('opdSparkline', hospitalData.metadata.months.map((_, i) => HospitalAnalytics.getOPDTotal(i)), '#E84A2D');
    this.createSparkline('ipdSparkline', hospitalData.metadata.months.map((_, i) => HospitalAnalytics.getIPDTotal(i)), '#2563EB');
    const delData = (hospitalData.maternity.monthly && hospitalData.maternity.monthly.length > 0)
      ? hospitalData.maternity.monthly.map(m => m.deliveries)
      : [109, 669, 898, 1141, 1252, 1441, 1437, 968];
    const csData = (hospitalData.maternity.monthly && hospitalData.maternity.monthly.length > 0)
      ? hospitalData.maternity.monthly.map(m => m.cs)
      : [99, 457, 547, 738, 779, 854, 824, 543];
    this.createSparkline('delSparkline', delData, '#10B981');
    this.createSparkline('csSparkline', csData, '#F97316');
  },

  createSparkline(canvasId, dataPoints, strokeColor) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (this.instances[canvasId]) {
      this.instances[canvasId].destroy();
    }

    this.instances[canvasId] = new Chart(canvas, {
      type: 'line',
      data: {
        labels: dataPoints.map((_, i) => i),
        datasets: [{
          data: dataPoints,
          borderColor: strokeColor,
          borderWidth: 2.5,
          pointRadius: 0,
          fill: false,
          tension: 0.4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        events: [],
        scales: {
          x: { display: false },
          y: { display: false }
        },
        plugins: {
          legend: { display: false },
          tooltip: { enabled: false }
        }
      }
    });
  },

  /**
   * Interactive Legend Component: Perfectly aligned columns for Department, Count, and Percentage
   */
  renderCustomLegend(containerId, labels, values, colors, total, type) {
    const container = document.getElementById(containerId);
    if (!container) return;

    container.innerHTML = '';
    labels.forEach((label, idx) => {
      const val = values[idx] || 0;
      const pct = total > 0 ? ((val / total) * 100).toFixed(1) : "0.0";
      const color = colors[idx] || '#999';

      const item = document.createElement('div');
      item.className = 'legend-row';
      item.title = `Click to inspect ${label}: ${val.toLocaleString()} (${pct}%)`;
      item.innerHTML = `
        <div class="legend-dept">
          <span class="legend-color-dot" style="background-color: ${color};"></span>
          <span class="legend-label">${label}</span>
        </div>
        <div class="legend-metrics">
          <span class="legend-count num-mono">${val.toLocaleString()}</span>
          <span class="legend-pct-badge num-mono">${pct}%</span>
        </div>
      `;

      item.addEventListener('click', () => {
        if (typeof window.showDepartmentDrilldown === 'function') {
          window.showDepartmentDrilldown(type, label);
        }
      });

      container.appendChild(item);
    });
  },

  /**
   * Refresh all visualizations
   */
  updateAll(periodFilter) {
    this.renderInpatientDonut(periodFilter);
    this.renderOutpatientDonut(periodFilter);
    this.renderDeliveryModeDonut(periodFilter);
    this.renderMortalityDonut(periodFilter);
    this.renderAdmissionsTrendChart();
    this.renderHistoricalDeliveriesChart();
    this.renderDepartmentComparisonBar('ipd', periodFilter);
    this.renderOPDRankingBar(periodFilter);
    this.renderSparklines();
    this.renderMultiYearTrajectoryChart();
    this.renderMultiYearDeliveriesChart();
    this.renderMultiYearDeptRankingChart();
  },

  /**
   * 10. Multi-Year Patient Volume Trajectory Chart (2019 - 2026)
   */
  renderMultiYearTrajectoryChart() {
    const canvas = document.getElementById('multiYearTrajectoryChart');
    if (!canvas) return;

    if (this.instances.multiYearTrajectory) {
      this.instances.multiYearTrajectory.destroy();
    }

    const hist = HISTORICAL_MULTI_YEAR;
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.multiYearTrajectory = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: hist.years,
        datasets: [
          {
            type: 'bar',
            label: 'Outpatients (OPD)',
            data: hist.annualTotals.opd,
            backgroundColor: '#E84A2D',
            hoverBackgroundColor: '#C0392B',
            borderRadius: 4,
            barPercentage: 0.9,
            categoryPercentage: 0.8,
            yAxisID: 'yOPD'
          },
          {
            type: 'bar',
            label: 'Inpatients (IPD)',
            data: hist.annualTotals.ipd,
            backgroundColor: '#2563EB',
            hoverBackgroundColor: '#1D4ED8',
            borderRadius: 4,
            barPercentage: 0.9,
            categoryPercentage: 0.8,
            yAxisID: 'yIPD'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { weight: '600', family: 'Inter' } }
          },
          yOPD: {
            type: 'linear',
            position: 'left',
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)' },
            ticks: {
              color: isDark ? '#FFA392' : '#E84A2D',
              font: { family: 'JetBrains Mono', weight: '600' },
              callback: (val) => val >= 1000 ? `${(val / 1000).toFixed(0)}k` : val
            },
            title: { display: true, text: 'OPD Visits', color: isDark ? '#FFA392' : '#E84A2D', font: { weight: 'bold' } }
          },
          yIPD: {
            type: 'linear',
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: {
              color: isDark ? '#93C5FD' : '#2563EB',
              font: { family: 'JetBrains Mono', weight: '600' },
              callback: (val) => val >= 1000 ? `${(val / 1000).toFixed(1)}k` : val
            },
            title: { display: true, text: 'IPD Admissions', color: isDark ? '#93C5FD' : '#2563EB', font: { weight: 'bold' } }
          }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              usePointStyle: false,
              boxWidth: 14,
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          },
          tooltip: {
            callbacks: {
              label(context) {
                return ` ${context.dataset.label}: ${context.raw.toLocaleString()}`;
              }
            }
          }
        }
      }
    });
  },

  /**
   * 11. Multi-Year Maternity Deliveries (CS vs SVD) Chart (2019 - 2026)
   */
  renderMultiYearDeliveriesChart() {
    const canvas = document.getElementById('multiYearDeliveriesChart');
    if (!canvas) return;

    if (this.instances.multiYearDeliveries) {
      this.instances.multiYearDeliveries.destroy();
    }

    const hist = HISTORICAL_MULTI_YEAR;
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.multiYearDeliveries = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: hist.years,
        datasets: [
          {
            label: 'C-Section (CS)',
            data: hist.annualTotals.cs,
            backgroundColor: '#E84A2D',
            hoverBackgroundColor: '#C0392B',
            borderRadius: 4
          },
          {
            label: 'Spontaneous Vaginal (SVD)',
            data: hist.annualTotals.svd,
            backgroundColor: '#10B981',
            hoverBackgroundColor: '#059669',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            stacked: true,
            grid: { display: false },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { weight: '600' } }
          },
          y: {
            stacked: true,
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)' },
            ticks: { color: isDark ? '#B8ABA5' : '#5C504B', font: { family: 'JetBrains Mono' } },
            title: { display: true, text: 'Total Deliveries', font: { weight: '600' } }
          }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              usePointStyle: true,
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          },
          tooltip: {
            callbacks: {
              label(context) {
                const yearIdx = context.dataIndex;
                const total = hist.annualTotals.deliveries[yearIdx];
                const pct = total > 0 ? ((context.raw / total) * 100).toFixed(1) : '0.0';
                return ` ${context.dataset.label}: ${context.raw.toLocaleString()} (${pct}%)`;
              },
              footer(items) {
                if (!items.length) return '';
                const yearIdx = items[0].dataIndex;
                const total = hist.annualTotals.deliveries[yearIdx];
                const live = hist.annualTotals.liveBirths[yearIdx];
                return `Total Deliveries: ${total.toLocaleString()} (${live.toLocaleString()} Live Births)`;
              }
            }
          }
        }
      }
    });
  },

  /**
   * 12. Top Department Volume Drivers (8-Year Cumulative OPD)
   */
  renderMultiYearDeptRankingChart() {
    const canvas = document.getElementById('multiYearDeptRankingChart');
    if (!canvas) return;

    if (this.instances.multiYearDeptRanking) {
      this.instances.multiYearDeptRanking.destroy();
    }

    const hist = HISTORICAL_MULTI_YEAR;
    const sortedDepts = [...hist.opdDepartments].sort((a, b) => b.total - a.total);
    const labels = sortedDepts.map(d => d.name);
    const dataVals = sortedDepts.map(d => d.total);
    const colors = sortedDepts.map(d => d.color);
    const grandTotal = hist.cumulativeSummary.totalOPD;
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    this.instances.multiYearDeptRanking = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: 'Cumulative OPD Consultations (2019-2026)',
          data: dataVals,
          backgroundColor: colors,
          borderRadius: 6,
          maxBarThickness: 22
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            grid: { color: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)' },
            ticks: {
              color: isDark ? '#B8ABA5' : '#5C504B',
              font: { family: 'JetBrains Mono' },
              callback: (val) => val >= 1000 ? `${(val / 1000).toFixed(0)}k` : val
            }
          },
          y: {
            grid: { display: false },
            ticks: {
              color: isDark ? '#FAF6F5' : '#241E1C',
              font: { family: 'Inter', size: 12, weight: '600' }
            }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                const val = context.raw;
                const pct = grandTotal > 0 ? ((val / grandTotal) * 100).toFixed(1) : '0.0';
                return ` ${val.toLocaleString()} visits (${pct}% of all 506,735 OPD visits)`;
              }
            }
          }
        }
      }
    });
  }
};

window.HospitalCharts = HospitalCharts;
