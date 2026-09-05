/**
 * DMC Hospital Statistics Dashboard (2026) - Real Excel Data Store
 * Source: PATIENTS_STATISTICS JAN TO JULY 2026.xlsx
 * Strictly includes data present in the Excel workbook:
 * - Sheet 1: OUT PATIENT PER DEPARTEMENT
 * - Sheet 2: INPATIENT PER DEPARTEMENT
 * - Sheet 3: BABIES TOTAL (Monthly & Historical 2019-2026)
 * - Sheet 4: Death (In-facility deaths, age, department, circumstance)
 */

const HOSPITAL_RAW_DATA = {
  metadata: {
    hospitalName: "Dream Medical Center Hospital",
    reportTitle: "Hospital Statistics & Clinical Performance Report",
    year: 2026,
    periodCovered: "January - August 2026",
    months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"],
    allMonths: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    lastUpdated: "2026-09-05",
    sourceFile: "PATIENTS STATISTICS - JAN TO AUG 2026.xlsx"
  },

  // Sheet 1: OUT PATIENT PER DEPARTEMENT (Jan - Aug 2026)
  outpatient: [
    { id: "pediatrics", name: "PEDIATRICS", monthly: [1508, 1629, 1740, 1321, 1579, 1487, 1191, 1099], total: 11554, color: "#2563EB" },
    { id: "general_med", name: "GENERAL MED", monthly: [1355, 1281, 1540, 1361, 1332, 1560, 1342, 1515], total: 11286, color: "#10B981" },
    { id: "physiotherapy", name: "PHYSIOTHERAPY", monthly: [1069, 1013, 1260, 1263, 1399, 1516, 1404, 1337], total: 10261, color: "#06B6D4" },
    { id: "gyne_obs", name: "GYNE & OBS", monthly: [1218, 1095, 1095, 1103, 1181, 1033, 1136, 1091], total: 8952, color: "#EC4899" },
    { id: "internal_med", name: "INTERNAL MED", monthly: [1080, 965, 1016, 1031, 1123, 1079, 1105, 1073], total: 8472, color: "#8B5CF6" },
    { id: "vaccination", name: "VACCINATION", monthly: [506, 415, 489, 412, 500, 454, 458, 408], total: 3642, color: "#F59E0B" },
    { id: "ortho_surg", name: "ORTHO SURG", monthly: [424, 462, 433, 390, 378, 374, 479, 519], total: 3459, color: "#F97316" },
    { id: "dental", name: "DENTAL", monthly: [324, 314, 345, 327, 268, 305, 333, 333], total: 2549, color: "#14B8A6" },
    { id: "ent_service", name: "ENT SERVICE", monthly: [191, 209, 220, 241, 186, 178, 216, 204], total: 1645, color: "#6366F1" },
    { id: "ophtalmology", name: "OPHTALMOLOGY", monthly: [131, 139, 161, 171, 148, 143, 150, 143], total: 1186, color: "#0EA5E9" },
    { id: "general_surgery", name: "GENERAL SURGERY", monthly: [59, 58, 54, 66, 35, 77, 182, 195], total: 726, color: "#E84A2D" },
    { id: "urology", name: "UROLOGY", monthly: [50, 47, 45, 61, 48, 36, 37, 29], total: 353, color: "#64748B" },
    { id: "emergency", name: "EMERGENCY", monthly: [2, 4, 3, 1, 4, 3, 1, 2], total: 20, color: "#EF4444" }
  ],

  // Sheet 2: INPATIENT PER DEPARTEMENT (Jan - Aug 2026)
  inpatient: [
    { id: "gyne_obs_in", name: "GYNE & OBS", monthly: [165, 154, 169, 177, 169, 150, 161, 143], total: 1288, color: "#EC4899" },
    { id: "pediatrics_in", name: "PEDIATRICS", monthly: [64, 92, 111, 74, 95, 92, 63, 63], total: 654, color: "#2563EB" },
    { id: "orthopedic_in", name: "ORTHOPEDIC", monthly: [53, 63, 47, 34, 40, 52, 59, 67], total: 415, color: "#F97316" },
    { id: "ent_service_in", name: "ENT SERVICE", monthly: [27, 34, 27, 84, 42, 44, 55, 96], total: 409, color: "#6366F1" },
    { id: "internal_med_in", name: "INTERNAL MED", monthly: [23, 26, 18, 35, 25, 28, 27, 31], total: 213, color: "#8B5CF6" },
    { id: "gn_surgery_in", name: "GN SURGERY", monthly: [7, 8, 7, 15, 6, 9, 15, 16], total: 83, color: "#E84A2D" },
    { id: "neonatology_in", name: "NEONATOLOGY", monthly: [4, 4, 2, 3, 8, 5, 5, 3], total: 34, color: "#14B8A6" },
    { id: "urology_surgery_in", name: "UROLOGY SURGERY", monthly: [5, 3, 6, 2, 0, 1, 3, 2], total: 22, color: "#64748B" },
    { id: "emergency_in", name: "EMERGENCY", monthly: [1, 0, 0, 0, 0, 0, 0, 1], total: 2, color: "#EF4444" }
  ],

  // Sheet 3: BABIES TOTAL
  maternity: {
    // Monthly 2026 Deliveries (Jan - Aug 2026)
    monthly: [
      { month: "Jan", period: "2026-01-31", deliveries: 116, liveBirths: 115, deaths: 1, deathReason: "Still births macerated", cs: 61, svd: 55 },
      { month: "Feb", period: "2026-02-28", deliveries: 114, liveBirths: 115, deaths: 0, deathReason: "None", cs: 59, svd: 55 },
      { month: "Mar", period: "2026-03-31", deliveries: 135, liveBirths: 135, deaths: 0, deathReason: "None", cs: 82, svd: 53 },
      { month: "Apr", period: "2026-04-30", deliveries: 132, liveBirths: 135, deaths: 0, deathReason: "None", cs: 72, svd: 61 },
      { month: "May", period: "2026-05-31", deliveries: 133, liveBirths: 137, deaths: 0, deathReason: "None", cs: 62, svd: 71 },
      { month: "Jun", period: "2026-06-30", deliveries: 112, liveBirths: 115, deaths: 0, deathReason: "None", cs: 72, svd: 40 },
      { month: "Jul", period: "2026-07-31", deliveries: 126, liveBirths: 125, deaths: 2, deathReason: "Still births macerated", cs: 75, svd: 51 },
      { month: "Aug", period: "2026-08-31", deliveries: 100, liveBirths: 100, deaths: 1, deathReason: "Still births macerated", cs: 60, svd: 40 }
    ],
    // Historical Multi-Year Record (2019 to 2026) from Sheet 3
    historical: [
      { year: "2019", deliveries: 109, liveBirths: 149, cs: 99, svd: 10 },
      { year: "2020", deliveries: 669, liveBirths: 664, cs: 457, svd: 212 },
      { year: "2021", deliveries: 898, liveBirths: 900, cs: 547, svd: 351 },
      { year: "2022", deliveries: 1141, liveBirths: 1137, cs: 738, svd: 403 },
      { year: "2023", deliveries: 1252, liveBirths: 1248, cs: 779, svd: 473 },
      { year: "2024", deliveries: 1441, liveBirths: 1444, cs: 854, svd: 587 },
      { year: "2025", deliveries: 1437, liveBirths: 1439, cs: 824, svd: 613 },
      { year: "2026 (8M)", deliveries: 968, liveBirths: 977, cs: 543, svd: 426 }
    ]
  },

  // Sheet 4: Death (In-facility mortalities)
  mortality: [
    { date: "2026-01-31", month: "Jan", count: 1, age: "0", department: "Maternity", circumstance: "Still births macerated" },
    { date: "2026-05-31", month: "May", count: 1, age: "Unspecified", department: "Internal medecine", circumstance: "Natural disease ( cancer)" },
    { date: "2026-06-30", month: "Jun", count: 1, age: "Unspecified", department: "Emergency", circumstance: "Natural disease" },
    { date: "2026-07-31", month: "Jul", count: 2, age: "0", department: "Maternity", circumstance: "Still births macerated" },
    { date: "2026-08-31", month: "Aug", count: 1, age: "0", department: "Maternity", circumstance: "Still births macerated" }
  ]
};

// Global reactive dataset instance
let hospitalData = JSON.parse(JSON.stringify(HOSPITAL_RAW_DATA));

/**
 * Exact Data Calculator & Accessor
 * Supports:
 * - null / 'all': full year (Jan - Aug)
 * - Array of month indices (e.g. [0, 1, 2] for Q1, [3, 4, 5] for Q2)
 * - Single month index (e.g. 7 for August)
 */
const HospitalAnalytics = {
  // Outpatient total
  getOPDTotal(periodFilter) {
    return hospitalData.outpatient.reduce((acc, dept) => {
      if (Array.isArray(periodFilter)) {
        return acc + periodFilter.reduce((s, idx) => s + (dept.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return acc + (dept.monthly[periodFilter] || 0);
      }
      return acc + dept.total;
    }, 0);
  },

  // Inpatient total
  getIPDTotal(periodFilter) {
    return hospitalData.inpatient.reduce((acc, dept) => {
      if (Array.isArray(periodFilter)) {
        return acc + periodFilter.reduce((s, idx) => s + (dept.monthly[idx] || 0), 0);
      } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
        return acc + (dept.monthly[periodFilter] || 0);
      }
      return acc + dept.total;
    }, 0);
  },

  // Total Deliveries (Recorded Jan - Jul 2026)
  getDeliveriesTotal(periodFilter) {
    const list = hospitalData.maternity.monthly;
    if (Array.isArray(periodFilter)) {
      return periodFilter.reduce((acc, idx) => acc + (list[idx] ? list[idx].deliveries : 0), 0);
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      return list[periodFilter] ? list[periodFilter].deliveries : 0;
    }
    return list.reduce((acc, m) => acc + m.deliveries, 0);
  },

  // Total Live Births (Recorded Jan - Jul 2026)
  getLiveBirthsTotal(periodFilter) {
    const list = hospitalData.maternity.monthly;
    if (Array.isArray(periodFilter)) {
      return periodFilter.reduce((acc, idx) => acc + (list[idx] ? list[idx].liveBirths : 0), 0);
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      return list[periodFilter] ? list[periodFilter].liveBirths : 0;
    }
    return list.reduce((acc, m) => acc + m.liveBirths, 0);
  },

  // C-Section vs SVD stats
  getCSStats(periodFilter) {
    const list = hospitalData.maternity.monthly;
    let cs = 0, svd = 0, total = 0;

    if (Array.isArray(periodFilter)) {
      periodFilter.forEach(idx => {
        if (list[idx]) {
          cs += list[idx].cs;
          svd += list[idx].svd;
          total += list[idx].deliveries;
        }
      });
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      if (list[periodFilter]) {
        cs = list[periodFilter].cs;
        svd = list[periodFilter].svd;
        total = list[periodFilter].deliveries;
      }
    } else {
      cs = list.reduce((acc, m) => acc + m.cs, 0);
      svd = list.reduce((acc, m) => acc + m.svd, 0);
      total = list.reduce((acc, m) => acc + m.deliveries, 0);
    }

    const rate = total > 0 ? ((cs / total) * 100).toFixed(1) : "0.0";
    return { cs, svd, total, rate: parseFloat(rate) };
  },

  // In-Facility Deaths Total
  getDeathsTotal(periodFilter) {
    const months = hospitalData.metadata.months;
    if (Array.isArray(periodFilter)) {
      const targetMonths = periodFilter.map(i => months[i]);
      return hospitalData.mortality.filter(d => targetMonths.includes(d.month)).reduce((acc, d) => acc + d.count, 0);
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      const targetMonth = months[periodFilter];
      return hospitalData.mortality.filter(d => d.month === targetMonth).reduce((acc, d) => acc + d.count, 0);
    }
    return hospitalData.mortality.reduce((acc, d) => acc + d.count, 0);
  },

  // Alias for getDeathsTotal
  getMortalityTotal(periodFilter) {
    return this.getDeathsTotal(periodFilter);
  },

  // Inpatient Mortality Rate (% of inpatients)
  getMortalityRate(periodFilter) {
    const totalIPD = this.getIPDTotal(periodFilter);
    const totalDeaths = this.getDeathsTotal(periodFilter);
    return totalIPD > 0 ? ((totalDeaths / totalIPD) * 100).toFixed(2) : "0.00";
  }
};
