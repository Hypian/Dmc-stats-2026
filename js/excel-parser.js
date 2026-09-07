/**
 * DMC Hospital Statistics - Dynamic Multi-Workbook Excel Parser (SheetJS integration)
 * Intelligently recognizes and ingests:
 *  1. PATIENTS STATISTICS - JAN TO AUG 2026.xlsx (Monthly OPD, IPD, Maternity, Mortality)
 *  2. PATIENTS NUMBERS 2025.xlsx (Full 12-Month OPD and IPD)
 *  3. PATIENTS NUMBERS 2019-2024.xlsx (6-Year Longitudinal Matrix)
 * 
 * Automatically synchronizes HOSPITAL_DATASETS and HISTORICAL_MULTI_YEAR,
 * recalculates all cumulative metrics with zero discrepancies, and triggers real-time UI updates.
 */

const ExcelDataParser = {
  isParsing: false,

  /**
   * Standardize department names across varying workbook spelling conventions
   */
  normalizeDeptName(rawName) {
    const clean = String(rawName || '').toUpperCase().trim();
    if (clean.includes('ORTHO')) return 'ORTHOPEDIC SURGERY';
    if (clean.includes('GN SURG') || clean.includes('GENERAL SURG')) return 'GENERAL SURGERY';
    if (clean.includes('GENERAL MED') || clean.includes('GN MED')) return 'GENERAL MED';
    if (clean.includes('GYNE') || clean.includes('OBS')) return 'GYNE & OBS';
    if (clean.includes('INTERNAL')) return 'INTERNAL MED';
    if (clean.includes('PEDIATRIC')) return 'PEDIATRICS';
    if (clean.includes('PHYSIO')) return 'PHYSIOTHERAPY';
    if (clean.includes('DENTAL')) return 'DENTAL';
    if (clean.includes('VACCIN')) return 'VACCINATION';
    if (clean.includes('ENT')) return 'ENT SERVICE';
    if (clean.includes('OPHT')) return 'OPHTALMOLOGY';
    if (clean.includes('UROL')) return 'UROLOGY';
    if (clean.includes('EMERG')) return 'EMERGENCY';
    if (clean.includes('NEONAT')) return 'NEONATOLOGY';
    return clean;
  },

  /**
   * Department color mapping helper
   */
  getColorForDept(normName) {
    const colorMap = {
      'PEDIATRICS': '#2563EB',
      'GENERAL MED': '#10B981',
      'GYNE & OBS': '#EC4899',
      'INTERNAL MED': '#8B5CF6',
      'PHYSIOTHERAPY': '#06B6D4',
      'DENTAL': '#14B8A6',
      'VACCINATION': '#F59E0B',
      'ORTHOPEDIC SURGERY': '#F97316',
      'ENT SERVICE': '#6366F1',
      'OPHTALMOLOGY': '#0EA5E9',
      'UROLOGY': '#64748B',
      'EMERGENCY': '#EF4444',
      'GENERAL SURGERY': '#E84A2D',
      'NEONATOLOGY': '#14B8A6'
    };
    return colorMap[normName] || '#16A085';
  },

  /**
   * Initialize file upload listeners (dropzone and file input)
   */
  init() {
    const fileInput = document.getElementById('excelFileInput');
    const dropZone = document.getElementById('excelDropZone');

    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) this.parseFile(file);
      });
    }

    if (dropZone) {
      ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.add('drag-over');
        });
      });

      ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.remove('drag-over');
        });
      });

      dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const file = dt.files[0];
        if (file) this.parseFile(file);
      });
    }
  },

  /**
   * Parse an Excel file using SheetJS (XLSX)
   */
  parseFile(file) {
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
      this.showToast("Invalid file format. Please upload an Excel (.xlsx or .xls) file.", "danger");
      return;
    }

    this.showToast(`Loading and analyzing "${file.name}"...`, "info");
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        const parseResult = this.processWorkbook(workbook, file.name);
        if (parseResult.success) {
          this.showToast(parseResult.message, "success");
          
          // Re-render entire dashboard for the detected dataset
          if (typeof window.onHospitalDataUpdated === 'function') {
            window.onHospitalDataUpdated(parseResult.targetYear);
          }
        } else {
          this.showToast(`Notice: ${parseResult.message}`, "warning");
        }
      } catch (err) {
        console.error("Excel parse error:", err);
        this.showToast(`Failed to parse file: ${err.message}`, "danger");
      }
    };

    reader.onerror = () => {
      this.showToast("Failed to read the file. Please check file permissions.", "danger");
    };

    reader.readAsArrayBuffer(file);
  },

  /**
   * Intelligent Workbook Type Detector & Processor
   */
  processWorkbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    const lowerFileName = fileName.toLowerCase();

    // Check for 2019-2024 Historical Workbook
    const isHistoricalFile = lowerFileName.includes('2019-2024') || lowerFileName.includes('2019_2024');
    let hasHistoricalHeaders = false;

    // Scan sheets for annual "JAN-DEC 2019" or 2025 month patterns
    for (const name of sheetNames) {
      const ws = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r] || [];
        if (row.some(cell => typeof cell === 'string' && (cell.includes('2019') || cell.includes('JAN-DEC')))) {
          hasHistoricalHeaders = true;
          break;
        }
      }
      if (hasHistoricalHeaders) break;
    }

    // 1. Process 2019-2024 Longitudinal Workbook
    if (isHistoricalFile || hasHistoricalHeaders) {
      return this.processHistoricalWorkbook(workbook, fileName);
    }

    // 2. Process 2025 Full-Year Workbook (or check if sheets have 12 month columns)
    const is2025File = lowerFileName.includes('2025');
    let has12MonthColumns = false;
    for (const name of sheetNames) {
      const lower = name.toLowerCase();
      if (lower.includes('opd') || lower.includes('out patient') || lower.includes('ip')) {
        const ws = workbook.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        const monthCols = this.detectMonthColumns(rows);
        if (monthCols.length === 12) {
          has12MonthColumns = true;
          break;
        }
      }
    }

    if (is2025File || has12MonthColumns) {
      return this.process2025Workbook(workbook, fileName);
    }

    // 3. Process 2026 Monthly / Standard Report (or Default)
    return this.process2026Workbook(workbook, fileName);
  },

  /**
   * Helper: Detect month column indices from sheet header rows
   */
  detectMonthColumns(rows) {
    const monthAliases = [
      ['JAN', 'JANUARY'], ['FEB', 'FEBRUARY'], ['MAR', 'MARCH'],
      ['APR', 'APRIL'], ['MAY'], ['JUN', 'JUNE'],
      ['JUL', 'JULY'], ['AUG', 'AUGUST'], ['SEP', 'SEPTEMBER'],
      ['OCT', 'OCTOBER'], ['NOV', 'NOVEMBER'], ['DEC', 'DECEMBER']
    ];

    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      const cols = [];

      row.forEach((cell, colIdx) => {
        if (!cell) return;
        const str = String(cell).trim().toUpperCase();
        if (str.includes('TOTAL') || str.includes('DEPART') || str.includes('WARD')) return;

        monthAliases.forEach((aliases, mIdx) => {
          if (aliases.some(alias => str === alias || str.startsWith(alias + ' '))) {
            cols.push({ colIndex: colIdx, monthIndex: mIdx, label: monthAliases[mIdx][0] });
          }
        });
      });

      if (cols.length >= 6) {
        // Sort by column index
        cols.sort((a, b) => a.colIndex - b.colIndex);
        return cols;
      }
    }

    // Default fallback: 8 months (cols C through J, index 2-9)
    return [2, 3, 4, 5, 6, 7, 8, 9].map((colIdx, i) => ({
      colIndex: colIdx,
      monthIndex: i,
      label: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug'][i]
    }));
  },

  /**
   * 1. PROCESS 2019-2024 HISTORICAL WORKBOOK
   */
  processHistoricalWorkbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    let opdParsed = false;
    let ipdParsed = false;

    sheetNames.forEach(sheetName => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const lower = sheetName.toLowerCase().trim();

      // Find header row with year columns
      let headerRowIndex = -1;
      let yearIndices = []; // { year: '2019', colIndex: 1 }

      for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r] || [];
        const foundYears = [];
        row.forEach((cell, idx) => {
          if (cell === null || cell === undefined) return;
          const str = String(cell).toUpperCase();
          for (let y = 2019; y <= 2024; y++) {
            if (str.includes(String(y))) {
              foundYears.push({ year: String(y), colIndex: idx });
            }
          }
        });
        if (foundYears.length >= 4) {
          headerRowIndex = r;
          yearIndices = foundYears;
          break;
        }
      }

      if (headerRowIndex === -1) return;

      // Parse OPD Sheet
      if (lower === 'opd' || lower.includes('out patient')) {
        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[0] ? String(row[0]).trim() : '';
          if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('DEPART')) continue;

          const norm = this.normalizeDeptName(rawName);
          const deptEntry = HISTORICAL_MULTI_YEAR.opdDepartments.find(d => this.normalizeDeptName(d.name) === norm);

          if (deptEntry) {
            yearIndices.forEach(item => {
              const yIdx = parseInt(item.year, 10) - 2019; // 2019 -> 0, 2024 -> 5
              const val = Number(row[item.colIndex]) || 0;
              if (yIdx >= 0 && yIdx < 6) {
                deptEntry.years[yIdx] = val;
              }
            });
            // Recalculate row total
            deptEntry.total = deptEntry.years.reduce((a, b) => a + b, 0);
          }
        }
        opdParsed = true;
      }

      // Parse IP Sheet
      if (lower === 'ip' || lower.includes('inpatient')) {
        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[0] ? String(row[0]).trim() : '';
          if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('WARD')) continue;

          const norm = this.normalizeDeptName(rawName);
          const wardEntry = HISTORICAL_MULTI_YEAR.ipdDepartments.find(d => this.normalizeDeptName(d.name) === norm);

          if (wardEntry) {
            yearIndices.forEach(item => {
              const yIdx = parseInt(item.year, 10) - 2019;
              const val = Number(row[item.colIndex]) || 0;
              if (yIdx >= 0 && yIdx < 6) {
                wardEntry.years[yIdx] = val;
              }
            });
            wardEntry.total = wardEntry.years.reduce((a, b) => a + b, 0);
          }
        }
        ipdParsed = true;
      }
    });

    // Recompute Annual Totals and Cumulative Summary
    for (let i = 0; i < 6; i++) {
      HISTORICAL_MULTI_YEAR.annualTotals.opd[i] = HISTORICAL_MULTI_YEAR.opdDepartments.reduce((acc, d) => acc + (d.years[i] || 0), 0);
      HISTORICAL_MULTI_YEAR.annualTotals.ipd[i] = HISTORICAL_MULTI_YEAR.ipdDepartments.reduce((acc, d) => acc + (d.years[i] || 0), 0);
    }

    const cum = HISTORICAL_MULTI_YEAR.cumulativeSummary;
    cum.totalOPD = HISTORICAL_MULTI_YEAR.annualTotals.opd.reduce((a, b) => a + b, 0);
    cum.totalIPD = HISTORICAL_MULTI_YEAR.annualTotals.ipd.reduce((a, b) => a + b, 0);
    cum.totalPatientsServed = cum.totalOPD + cum.totalIPD;

    return {
      success: true,
      targetYear: 'multi',
      message: `Successfully ingested "${fileName}"! Updated 6-year longitudinal baseline (2019–2024). Cumulative OPD: ${cum.totalOPD.toLocaleString()}, IPD: ${cum.totalIPD.toLocaleString()}.`
    };
  },

  /**
   * 2. PROCESS 2025 FULL-YEAR WORKBOOK
   */
  process2025Workbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    const newOPD = [];
    const newIPD = [];

    sheetNames.forEach(sheetName => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const lower = sheetName.toLowerCase().trim();

      // Find header row and month columns
      let headerRowIndex = -1;
      let monthCols = [];

      for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r] || [];
        const detected = [];
        row.forEach((cell, cIdx) => {
          if (!cell) return;
          const str = String(cell).toUpperCase();
          if (str.includes('JAN')) detected.push({ col: cIdx, m: 0 });
          else if (str.includes('FEB')) detected.push({ col: cIdx, m: 1 });
          else if (str.includes('MAR')) detected.push({ col: cIdx, m: 2 });
          else if (str.includes('APR')) detected.push({ col: cIdx, m: 3 });
          else if (str.includes('MAY')) detected.push({ col: cIdx, m: 4 });
          else if (str.includes('JUN')) detected.push({ col: cIdx, m: 5 });
          else if (str.includes('JUL')) detected.push({ col: cIdx, m: 6 });
          else if (str.includes('AUG')) detected.push({ col: cIdx, m: 7 });
          else if (str.includes('SEP')) detected.push({ col: cIdx, m: 8 });
          else if (str.includes('OCT')) detected.push({ col: cIdx, m: 9 });
          else if (str.includes('NOV')) detected.push({ col: cIdx, m: 10 });
          else if (str.includes('DEC')) detected.push({ col: cIdx, m: 11 });
        });
        if (detected.length >= 6) {
          headerRowIndex = r;
          monthCols = detected;
          break;
        }
      }

      if (headerRowIndex === -1) {
        // Fallback: columns C through N (index 2-13)
        headerRowIndex = 1;
        monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13].map((c, m) => ({ col: c, m }));
      }

      // Outpatient 2025
      if (lower.includes('opd') || lower.includes('out patient')) {
        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[0] ? String(row[0]).trim() : '';
          if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('DEPART')) continue;

          const monthly = new Array(12).fill(0);
          monthCols.forEach(item => {
            const val = Number(row[item.col]) || 0;
            if (item.m >= 0 && item.m < 12) monthly[item.m] = val;
          });

          const total = monthly.reduce((a, b) => a + b, 0);
          if (total > 0) {
            const norm = this.normalizeDeptName(rawName);
            newOPD.push({
              id: rawName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
              name: rawName.toUpperCase(),
              monthly: monthly,
              total: total,
              color: this.getColorForDept(norm)
            });

            // Sync to Multi-Year 2025 column (index 6)
            const myEntry = HISTORICAL_MULTI_YEAR.opdDepartments.find(d => this.normalizeDeptName(d.name) === norm);
            if (myEntry) {
              myEntry.years[6] = total;
              myEntry.total = myEntry.years.reduce((a, b) => a + b, 0);
            }
          }
        }
      }

      // Inpatient 2025
      if (lower === 'ip' || lower.includes('inpatient')) {
        for (let r = headerRowIndex + 1; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[0] ? String(row[0]).trim() : '';
          if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('WARD')) continue;

          const monthly = new Array(12).fill(0);
          monthCols.forEach(item => {
            const val = Number(row[item.col]) || 0;
            if (item.m >= 0 && item.m < 12) monthly[item.m] = val;
          });

          const total = monthly.reduce((a, b) => a + b, 0);
          if (total > 0) {
            const norm = this.normalizeDeptName(rawName);
            newIPD.push({
              id: rawName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
              name: rawName.toUpperCase(),
              monthly: monthly,
              total: total,
              color: this.getColorForDept(norm)
            });

            // Sync to Multi-Year 2025 column (index 6)
            const myEntry = HISTORICAL_MULTI_YEAR.ipdDepartments.find(d => this.normalizeDeptName(d.name) === norm);
            if (myEntry) {
              myEntry.years[6] = total;
              myEntry.total = myEntry.years.reduce((a, b) => a + b, 0);
            }
          }
        }
      }
    });

    if (newOPD.length > 0) HOSPITAL_DATASETS['2025'].outpatient = newOPD;
    if (newIPD.length > 0) HOSPITAL_DATASETS['2025'].inpatient = newIPD;

    // Recalculate Multi-Year 2025 Total
    const opd2025Total = newOPD.reduce((acc, d) => acc + d.total, 0);
    const ipd2025Total = newIPD.reduce((acc, d) => acc + d.total, 0);
    if (opd2025Total > 0) HISTORICAL_MULTI_YEAR.annualTotals.opd[6] = opd2025Total;
    if (ipd2025Total > 0) HISTORICAL_MULTI_YEAR.annualTotals.ipd[6] = ipd2025Total;

    const cum = HISTORICAL_MULTI_YEAR.cumulativeSummary;
    cum.totalOPD = HISTORICAL_MULTI_YEAR.annualTotals.opd.reduce((a, b) => a + b, 0);
    cum.totalIPD = HISTORICAL_MULTI_YEAR.annualTotals.ipd.reduce((a, b) => a + b, 0);
    cum.totalPatientsServed = cum.totalOPD + cum.totalIPD;

    // If currently on 2025, refresh working copy
    if (HospitalAnalytics.getActiveYear() === '2025') {
      HospitalAnalytics.setActiveYear('2025');
    }

    return {
      success: true,
      targetYear: '2025',
      message: `Successfully processed "${fileName}"! Updated full-year 2025: ${opd2025Total.toLocaleString()} OPD visits and ${ipd2025Total.toLocaleString()} IPD admissions across 12 months.`
    };
  },

  /**
   * 3. PROCESS 2026 WORKBOOK (Monthly report with OPD, IPD, Babies, Death)
   */
  process2026Workbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    let foundSheets = { opd: false, ipd: false, babies: false, death: false };
    let stats = { opdCount: 0, ipdCount: 0, deliveries: 0 };

    sheetNames.forEach(name => {
      const lower = name.toLowerCase().trim();
      const ws = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      // Outpatient Sheet
      if (lower.includes('out patient') || lower.includes('outpatient') || lower === 'opd') {
        foundSheets.opd = true;
        stats.opdCount = this.parse2026Outpatient(rows);
      }
      // Inpatient Sheet
      else if (lower.includes('inpatient') || lower.includes('in patient') || lower === 'ip') {
        foundSheets.ipd = true;
        stats.ipdCount = this.parse2026Inpatient(rows);
      }
      // Babies Sheet
      else if (lower.includes('babies') || lower.includes('deliveries') || lower.includes('maternity')) {
        foundSheets.babies = true;
        stats.deliveries = this.parse2026Babies(rows);
      }
      // Death Sheet
      else if (lower.includes('death') || lower.includes('mortality')) {
        foundSheets.death = true;
        this.parse2026Death(rows);
      }
    });

    if (!foundSheets.opd && !foundSheets.ipd && !foundSheets.babies) {
      return {
        success: false,
        message: "No recognizable hospital sheets found in this file. Please verify format."
      };
    }

    // Recalculate 2026 multi-year baseline column (index 7)
    const opd2026Total = HOSPITAL_DATASETS['2026'].outpatient.reduce((acc, d) => acc + d.total, 0);
    const ipd2026Total = HOSPITAL_DATASETS['2026'].inpatient.reduce((acc, d) => acc + d.total, 0);
    HISTORICAL_MULTI_YEAR.annualTotals.opd[7] = opd2026Total;
    HISTORICAL_MULTI_YEAR.annualTotals.ipd[7] = ipd2026Total;

    const cum = HISTORICAL_MULTI_YEAR.cumulativeSummary;
    cum.totalOPD = HISTORICAL_MULTI_YEAR.annualTotals.opd.reduce((a, b) => a + b, 0);
    cum.totalIPD = HISTORICAL_MULTI_YEAR.annualTotals.ipd.reduce((a, b) => a + b, 0);
    cum.totalPatientsServed = cum.totalOPD + cum.totalIPD;

    if (HospitalAnalytics.getActiveYear() === '2026') {
      HospitalAnalytics.setActiveYear('2026');
    }

    return {
      success: true,
      targetYear: '2026',
      message: `Successfully processed "${fileName}"! Updated ${stats.opdCount} OPD services, ${stats.ipdCount} IPD wards, and ${stats.deliveries} deliveries.`
    };
  },

  parse2026Outpatient(rows) {
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      if (row.some(c => typeof c === 'string' && (c.toUpperCase().includes('DEPART') || c.toUpperCase().includes('OUT PATIENT')))) {
        headerRowIndex = r;
        break;
      }
    }
    if (headerRowIndex === -1) headerRowIndex = 2;

    const newOPD = [];
    const monthCols = [2, 3, 4, 5, 6, 7, 8, 9]; // columns C through J (Jan-Aug)

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const rawName = row[0] ? String(row[0]).trim() : '';
      if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('DEPART')) continue;

      const monthly = monthCols.map(c => Number(row[c]) || 0);
      const total = monthly.reduce((a, b) => a + b, 0);

      if (total > 0) {
        const norm = this.normalizeDeptName(rawName);
        newOPD.push({
          id: rawName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
          name: rawName.toUpperCase(),
          monthly: monthly,
          total: total,
          color: this.getColorForDept(norm)
        });

        // Sync to multi-year 2026 (index 7)
        const myEntry = HISTORICAL_MULTI_YEAR.opdDepartments.find(d => this.normalizeDeptName(d.name) === norm);
        if (myEntry) {
          myEntry.years[7] = total;
          myEntry.total = myEntry.years.reduce((a, b) => a + b, 0);
        }
      }
    }

    if (newOPD.length > 0) {
      HOSPITAL_DATASETS['2026'].outpatient = newOPD;
    }
    return newOPD.length;
  },

  parse2026Inpatient(rows) {
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      if (row.some(c => typeof c === 'string' && c.toUpperCase().includes('INPATIENT'))) {
        headerRowIndex = r;
        break;
      }
    }
    if (headerRowIndex === -1) headerRowIndex = 1;

    const newIPD = [];
    const monthCols = [2, 3, 4, 5, 6, 7, 8, 9]; // Columns C through J (Jan-Aug)

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const rawName = row[0] ? String(row[0]).trim() : '';
      if (!rawName || rawName.toUpperCase().includes('TOTAL') || rawName.toUpperCase().includes('INPATIENT')) continue;

      // Ensure August (col 9) is included without the Excel row formula bug!
      const monthly = monthCols.map(c => Number(row[c]) || 0);
      const total = monthly.reduce((a, b) => a + b, 0);

      if (total > 0) {
        const norm = this.normalizeDeptName(rawName);
        newIPD.push({
          id: rawName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
          name: rawName.toUpperCase(),
          monthly: monthly,
          total: total,
          color: this.getColorForDept(norm)
        });

        // Sync to multi-year 2026 (index 7)
        const myEntry = HISTORICAL_MULTI_YEAR.ipdDepartments.find(d => this.normalizeDeptName(d.name) === norm);
        if (myEntry) {
          myEntry.years[7] = total;
          myEntry.total = myEntry.years.reduce((a, b) => a + b, 0);
        }
      }
    }

    if (newIPD.length > 0) {
      HOSPITAL_DATASETS['2026'].inpatient = newIPD;
    }
    return newIPD.length;
  },

  parse2026Babies(rows) {
    const monthlyData = [];
    let totalDeliveries = 0;

    for (let r = 1; r < Math.min(rows.length, 14); r++) {
      const row = rows[r] || [];
      if (!row[0] || String(row[0]).toUpperCase().includes('TOTAL')) break;

      const del = Number(row[1]) || 0;
      const live = Number(row[2]) || 0;
      const death = Number(row[3]) || 0;
      const cs = Number(row[5]) || 0;
      const svd = Number(row[6]) || 0;

      if (del > 0 || live > 0) {
        let monthName = "M" + r;
        if (row[0] instanceof Date) {
          monthName = row[0].toLocaleString('default', { month: 'short' });
        } else if (typeof row[0] === 'string' && row[0].includes('-')) {
          const parts = row[0].split('-');
          const mIdx = parseInt(parts[1], 10) - 1;
          monthName = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug'][mIdx] || ("M" + r);
        }

        monthlyData.push({
          month: monthName,
          period: String(row[0]).substring(0, 10),
          deliveries: del,
          liveBirths: live,
          deaths: death,
          deathReason: row[4] ? String(row[4]) : 'None',
          cs: cs,
          svd: svd
        });
        totalDeliveries += del;
      }
    }

    if (monthlyData.length > 0) {
      HOSPITAL_DATASETS['2026'].maternity.monthly = monthlyData;
      // Sync 2026 maternity totals into historical
      HISTORICAL_MULTI_YEAR.annualTotals.deliveries[7] = totalDeliveries;
      HISTORICAL_MULTI_YEAR.annualTotals.liveBirths[7] = monthlyData.reduce((a, b) => a + b.liveBirths, 0);
      HISTORICAL_MULTI_YEAR.annualTotals.cs[7] = monthlyData.reduce((a, b) => a + b.cs, 0);
      HISTORICAL_MULTI_YEAR.annualTotals.svd[7] = monthlyData.reduce((a, b) => a + b.svd, 0);
    }
    return totalDeliveries;
  },

  parse2026Death(rows) {
    const deaths = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (!row[0] || String(row[0]).toUpperCase().includes('TOTAL')) continue;
      const count = Number(row[1]) || 0;
      if (count > 0) {
        let mName = "Month";
        if (row[0] instanceof Date) {
          mName = row[0].toLocaleString('default', { month: 'short' });
        }
        deaths.push({
          date: String(row[0]).substring(0, 10),
          month: mName,
          count: count,
          age: row[2] !== undefined && row[2] !== null ? String(row[2]) : "Unspecified",
          department: row[3] ? String(row[3]).trim() : "General",
          circumstance: row[4] ? String(row[4]).trim() : "Unspecified",
          category: "Clinical",
          severity: "high"
        });
      }
    }
    if (deaths.length > 0) {
      HOSPITAL_DATASETS['2026'].mortality = deaths;
    }
  },

  /**
   * Display toast notification
   */
  showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `dashboard-toast toast-${type}`;
    
    let iconClass = 'fa-info-circle';
    if (type === 'success') iconClass = 'fa-check-circle';
    if (type === 'warning') iconClass = 'fa-exclamation-triangle';
    if (type === 'danger') iconClass = 'fa-times-circle';

    toast.innerHTML = `
      <i class="fas ${iconClass}"></i>
      <span class="toast-message">${message}</span>
      <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;

    container.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('fade-out');
      setTimeout(() => toast.remove(), 400);
    }, 4500);
  }
};
