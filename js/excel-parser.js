/**
 * DMC Hospital Statistics - Dynamic Excel Parser (SheetJS integration)
 * Enables hospital administrators to upload new or updated Excel workbooks
 * and immediately re-calculate all dashboard KPIs, charts, and tables in real-time.
 */

const ExcelDataParser = {
  isParsing: false,

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

    this.showToast(`Loading and analyzing ${file.name}...`, "info");
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        const parseResult = this.processWorkbook(workbook, file.name);
        if (parseResult.success) {
          this.showToast(`Successfully processed "${file.name}"! Updated ${parseResult.stats.opdCount} OPD departments, ${parseResult.stats.ipdCount} IPD wards, and ${parseResult.stats.deliveries} deliveries.`, "success");
          
          // Re-render entire dashboard
          if (typeof window.onHospitalDataUpdated === 'function') {
            window.onHospitalDataUpdated(hospitalData);
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
   * Process workbook sheets and update `hospitalData`
   */
  processWorkbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    let foundSheets = { opd: false, ipd: false, babies: false, death: false };
    let stats = { opdCount: 0, ipdCount: 0, deliveries: 0 };

    sheetNames.forEach(name => {
      const lower = name.toLowerCase();
      const ws = workbook.Sheets[name];
      const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      // Outpatient Sheet
      if (lower.includes('out patient') || lower.includes('outpatient') || lower.includes('opd')) {
        foundSheets.opd = true;
        stats.opdCount = this.parseOutpatientSheet(jsonRows);
      }
      // Inpatient Sheet
      else if (lower.includes('inpatient') || lower.includes('in patient') || lower.includes('ipd')) {
        foundSheets.ipd = true;
        stats.ipdCount = this.parseInpatientSheet(jsonRows);
      }
      // Babies & Deliveries Sheet
      else if (lower.includes('babies') || lower.includes('deliveries') || lower.includes('maternity')) {
        foundSheets.babies = true;
        stats.deliveries = this.parseBabiesSheet(jsonRows);
      }
      // Mortality / Death Sheet
      else if (lower.includes('death') || lower.includes('mortality')) {
        foundSheets.death = true;
        this.parseDeathSheet(jsonRows);
      }
    });

    // Update metadata
    hospitalData.metadata.sourceFile = fileName;
    hospitalData.metadata.lastUpdated = new Date().toISOString();

    if (!foundSheets.opd && !foundSheets.ipd && !foundSheets.babies) {
      return {
        success: false,
        message: "No recognized sheets found. Expected sheet names like 'OUT PATIENT', 'INPATIENT', 'BABIES', or 'Death'."
      };
    }

    return { success: true, stats };
  },

  /**
   * Extract Outpatient statistics
   */
  parseOutpatientSheet(rows) {
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      if (row.some(cell => typeof cell === 'string' && (cell.toUpperCase().includes('DEPART') || cell.toUpperCase().includes('OUT PATIENT')))) {
        headerRowIndex = r;
        break;
      }
    }

    if (headerRowIndex === -1) headerRowIndex = 2; // fallback row 3

    const newOPD = [];
    const monthCols = [2, 3, 4, 5, 6, 7, 8, 9]; // columns C through J (Jan-Aug)

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const deptName = row[0] ? String(row[0]).trim() : '';
      if (!deptName || deptName.toUpperCase().includes('TOTAL') || deptName.toUpperCase().includes('DEPART')) continue;

      const monthlyVals = monthCols.map(colIdx => {
        const val = Number(row[colIdx]);
        return isNaN(val) ? 0 : val;
      });

      const total = monthlyVals.reduce((a, b) => a + b, 0);
      if (total > 0) {
        newOPD.push({
          id: deptName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
          name: deptName,
          monthly: monthlyVals,
          total: total,
          category: "specialty"
        });
      }
    }

    if (newOPD.length > 0) {
      hospitalData.outpatient = newOPD;
    }
    return newOPD.length;
  },

  /**
   * Extract Inpatient statistics
   */
  parseInpatientSheet(rows) {
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      if (row.some(cell => typeof cell === 'string' && cell.toUpperCase().includes('INPATIENT'))) {
        headerRowIndex = r;
        break;
      }
    }
    if (headerRowIndex === -1) headerRowIndex = 1;

    const newIPD = [];
    const monthCols = [2, 3, 4, 5, 6, 7, 8, 9];

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const wardName = row[0] ? String(row[0]).trim() : '';
      if (!wardName || wardName.toUpperCase().includes('TOTAL') || wardName.toUpperCase().includes('INPATIENT')) continue;

      const monthlyVals = monthCols.map(colIdx => {
        const val = Number(row[colIdx]);
        return isNaN(val) ? 0 : val;
      });

      const total = monthlyVals.reduce((a, b) => a + b, 0);
      if (total > 0) {
        // match bed capacity if existing
        const existing = hospitalData.inpatient.find(i => i.name.toLowerCase().includes(wardName.toLowerCase()) || wardName.toLowerCase().includes(i.name.toLowerCase()));
        newIPD.push({
          id: wardName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
          name: wardName,
          monthly: monthlyVals,
          total: total,
          color: existing ? existing.color : "#16A085"
        });
      }
    }

    if (newIPD.length > 0) {
      hospitalData.inpatient = newIPD;
    }
    return newIPD.length;
  },

  /**
   * Extract Deliveries & Babies statistics
   */
  parseBabiesSheet(rows) {
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
          monthName = hospitalData.metadata.allMonths[mIdx] || ("M" + r);
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
      hospitalData.maternity.monthly = monthlyData;
    }
    return totalDeliveries;
  },

  /**
   * Extract Mortality entries
   */
  parseDeathSheet(rows) {
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
      hospitalData.mortality = deaths;
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
