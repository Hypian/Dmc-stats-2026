/**
 * DMC Hospital Statistics Dashboard - Multi-Year Real Excel Data Store
 * Sources:
 * 1. DMC_Patient_Statistics 2026 YTD.xlsx (Jan - Sep 2026 outpatient, inpatient, maternity, and mortality)
 * 2. PATIENTS NUMBERS 2025.xlsx (Jan - Dec 2025)
 * 3. PATIENTS NUMBERS 2019-2024.xlsx (Annual 2019 - 2024)
 * 
 * Verified 100% accurate against original workbook formulas and records.
 */

// Master dataset by Year
const HOSPITAL_DATASETS = {
  '2026': {
    metadata: {
      hospitalName: 'Dream Medical Center Hospital',
      reportTitle: 'Hospital Statistics & Clinical Performance Report',
      year: 2026,
      periodCovered: 'January - September 2026',
      periodBadge: 'Jan 01 – Sep 30, 2026',
      months: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'],
      allMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      lastUpdated: '2026-10-06',
      sourceFile: 'DMC_Patient_Statistics 2026 YTD.xlsx',
      hasMortality: true,
      hasMonthlyMaternity: true
    },
    outpatient: [
      { id: 'pediatrics', name: 'PEDIATRICS', monthly: [1508, 1629, 1740, 1321, 1579, 1487, 1191, 1099, 1324], total: 12878, color: '#2563EB' },
      { id: 'general_med', name: 'GENERAL MED', monthly: [1355, 1281, 1540, 1361, 1332, 1560, 1342, 1515, 1490], total: 12776, color: '#10B981' },
      { id: 'physiotherapy', name: 'PHYSIOTHERAPY', monthly: [1069, 1013, 1260, 1263, 1399, 1516, 1404, 1337, 1383], total: 11644, color: '#06B6D4' },
      { id: 'gyne_obs', name: 'GYNE & OBS', monthly: [1218, 1095, 1095, 1103, 1181, 1033, 1136, 1091, 1174], total: 10126, color: '#EC4899' },
      { id: 'internal_med', name: 'INTERNAL MED', monthly: [1080, 965, 1016, 1031, 1123, 1079, 1105, 1073, 1093], total: 9565, color: '#8B5CF6' },
      { id: 'vaccination', name: 'VACCINATION', monthly: [506, 415, 489, 412, 500, 454, 458, 408, 479], total: 4121, color: '#F59E0B' },
      { id: 'ortho_surg', name: 'ORTHO SURG', monthly: [424, 462, 433, 390, 378, 374, 479, 519, 395], total: 3854, color: '#F97316' },
      { id: 'dental', name: 'DENTAL', monthly: [324, 314, 345, 327, 268, 305, 333, 333, 416], total: 2965, color: '#14B8A6' },
      { id: 'ent_service', name: 'ENT SERVICE', monthly: [191, 209, 220, 241, 186, 178, 216, 204, 172], total: 1817, color: '#6366F1' },
      { id: 'ophtalmology', name: 'OPHTALMOLOGY', monthly: [131, 139, 161, 171, 148, 143, 150, 143, 139], total: 1325, color: '#0EA5E9' },
      { id: 'general_surgery', name: 'GENERAL SURGERY', monthly: [59, 58, 54, 66, 35, 77, 182, 195, 122], total: 848, color: '#E84A2D' },
      { id: 'urology', name: 'UROLOGY', monthly: [50, 47, 45, 61, 48, 36, 37, 29, 39], total: 392, color: '#64748B' },
      { id: 'emergency', name: 'EMERGENCY', monthly: [2, 4, 3, 1, 4, 3, 1, 2, 3], total: 23, color: '#EF4444' }
    ],
    inpatient: [
      { id: 'gyne_obs_in', name: 'GYNE & OBS', monthly: [165, 154, 169, 177, 169, 150, 161, 143, 166], total: 1454, color: '#EC4899' },
      { id: 'pediatrics_in', name: 'PEDIATRICS', monthly: [64, 92, 111, 74, 95, 92, 63, 63, 82], total: 736, color: '#2563EB' },
      { id: 'orthopedic_in', name: 'ORTHOPEDIC', monthly: [53, 63, 47, 34, 40, 52, 59, 67, 43], total: 458, color: '#F97316' },
      { id: 'ent_service_in', name: 'ENT SERVICE', monthly: [27, 34, 27, 84, 42, 44, 55, 96, 41], total: 450, color: '#6366F1' },
      { id: 'internal_med_in', name: 'INTERNAL MED', monthly: [23, 26, 18, 35, 25, 28, 27, 31, 23], total: 236, color: '#8B5CF6' },
      { id: 'gn_surgery_in', name: 'GN SURGERY', monthly: [7, 8, 7, 15, 6, 9, 15, 16, 12], total: 95, color: '#E84A2D' },
      { id: 'neonatology_in', name: 'NEONATOLOGY', monthly: [4, 4, 2, 3, 8, 5, 5, 3, 3], total: 37, color: '#14B8A6' },
      { id: 'urology_surgery_in', name: 'UROLOGY SURGERY', monthly: [5, 3, 6, 2, 0, 1, 3, 2, 5], total: 27, color: '#64748B' },
      { id: 'emergency_in', name: 'EMERGENCY', monthly: [1, 0, 0, 0, 0, 0, 0, 1, 0], total: 2, color: '#EF4444' }
    ],
    maternity: {
      monthly: [
        { month: 'Jan', period: '2026-01-31', deliveries: 116, liveBirths: 115, deaths: 1, deathReason: 'Still births macerated', cs: 61, svd: 55 },
        { month: 'Feb', period: '2026-02-28', deliveries: 114, liveBirths: 115, deaths: 0, deathReason: 'None', cs: 59, svd: 55 },
        { month: 'Mar', period: '2026-03-31', deliveries: 135, liveBirths: 135, deaths: 0, deathReason: 'None', cs: 82, svd: 53 },
        { month: 'Apr', period: '2026-04-30', deliveries: 132, liveBirths: 135, deaths: 0, deathReason: 'None', cs: 72, svd: 60 },
        { month: 'May', period: '2026-05-31', deliveries: 133, liveBirths: 137, deaths: 0, deathReason: 'None', cs: 62, svd: 71 },
        { month: 'Jun', period: '2026-06-30', deliveries: 112, liveBirths: 115, deaths: 0, deathReason: 'None', cs: 72, svd: 40 },
        { month: 'Jul', period: '2026-07-31', deliveries: 126, liveBirths: 125, deaths: 2, deathReason: 'Still births macerated', cs: 75, svd: 51 },
        { month: 'Aug', period: '2026-08-31', deliveries: 100, liveBirths: 100, deaths: 1, deathReason: 'Still births macerated', cs: 60, svd: 40 },
        { month: 'Sep', period: '2026-09-30', deliveries: 147, liveBirths: 149, deaths: 1, deathReason: 'Still births macerated', cs: 80, svd: 67 }
      ],
      historical: [
        { year: '2019', deliveries: 109, liveBirths: 149, cs: 99, svd: 10 },
        { year: '2020', deliveries: 669, liveBirths: 664, cs: 457, svd: 212 },
        { year: '2021', deliveries: 898, liveBirths: 900, cs: 547, svd: 351 },
        { year: '2022', deliveries: 1141, liveBirths: 1137, cs: 738, svd: 403 },
        { year: '2023', deliveries: 1252, liveBirths: 1248, cs: 779, svd: 473 },
        { year: '2024', deliveries: 1441, liveBirths: 1444, cs: 854, svd: 587 },
        { year: '2025', deliveries: 1437, liveBirths: 1439, cs: 824, svd: 613 },
        { year: '2026 (9M)', deliveries: 1115, liveBirths: 1126, cs: 623, svd: 492 }
      ]
    },
    mortality: [
      { date: '2026-01-31', month: 'Jan', count: 1, age: '0', department: 'Maternity', circumstance: 'Still births macerated' },
      { date: '2026-05-31', month: 'May', count: 1, age: 'Unspecified', department: 'Internal Medicine', circumstance: 'Natural disease (cancer)' },
      { date: '2026-06-30', month: 'Jun', count: 1, age: 'Unspecified', department: 'Emergency', circumstance: 'Natural disease' },
      { date: '2026-07-31', month: 'Jul', count: 2, age: '0', department: 'Maternity', circumstance: 'Still births macerated' },
      { date: '2026-08-31', month: 'Aug', count: 1, age: '0', department: 'Maternity', circumstance: 'Still births macerated' },
      { date: '2026-09-30', month: 'Sep', count: 1, age: '0', department: 'Maternity', circumstance: 'Still births macerated' }
    ]
  },

  '2025': {
    metadata: {
      hospitalName: 'Dream Medical Center Hospital',
      reportTitle: 'Hospital Statistics & Clinical Performance Report 2025',
      year: 2025,
      periodCovered: 'January - December 2025 (Full Year)',
      periodBadge: 'Jan 01 – Dec 31, 2025',
      months: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      allMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      lastUpdated: '2026-09-07',
      sourceFile: 'PATIENTS NUMBERS 2025.xlsx',
      hasMortality: false,
      hasMonthlyMaternity: false
    },
    outpatient: [
      { id: 'pediatrics', name: 'PEDIATRICS', monthly: [1580, 1844, 1720, 1402, 1595, 1942, 1271, 1099, 1169, 1754, 2241, 1838], total: 19455, color: '#2563EB' },
      { id: 'physiotherapy', name: 'PHYSIOTHERAPY', monthly: [1355, 1406, 1572, 1699, 1818, 1721, 1767, 1458, 1259, 1406, 1405, 1198], total: 18064, color: '#06B6D4' },
      { id: 'general_med', name: 'GENERAL MED', monthly: [1436, 1231, 1459, 1376, 1494, 1632, 1294, 1328, 1129, 1423, 1866, 1844], total: 17512, color: '#10B981' },
      { id: 'gyne_obs', name: 'GYNE & OBS', monthly: [1229, 1092, 1162, 1123, 1238, 1103, 1117, 1100, 1050, 1221, 1128, 1087], total: 13650, color: '#EC4899' },
      { id: 'internal_med', name: 'INTERNAL MED', monthly: [1202, 1012, 1060, 1075, 1185, 1208, 1116, 1085, 961, 1170, 1216, 1225], total: 13515, color: '#8B5CF6' },
      { id: 'vaccination', name: 'VACCINATION', monthly: [450, 375, 435, 457, 451, 437, 530, 459, 436, 472, 431, 468], total: 5401, color: '#F59E0B' },
      { id: 'ortho_surg', name: 'ORTHO SURG', monthly: [436, 380, 356, 448, 473, 402, 609, 441, 375, 482, 473, 427], total: 5302, color: '#F97316' },
      { id: 'dental', name: 'DENTAL', monthly: [440, 413, 473, 446, 401, 339, 342, 358, 325, 355, 310, 327], total: 4529, color: '#14B8A6' },
      { id: 'ent_service', name: 'ENT SERVICE', monthly: [164, 120, 111, 123, 136, 141, 127, 126, 102, 127, 227, 186], total: 1690, color: '#6366F1' },
      { id: 'ophtalmology', name: 'OPHTALMOLOGY', monthly: [137, 100, 122, 125, 123, 88, 88, 112, 115, 128, 122, 112], total: 1372, color: '#0EA5E9' },
      { id: 'urology', name: 'UROLOGY', monthly: [69, 91, 78, 88, 108, 83, 94, 98, 58, 65, 59, 40], total: 931, color: '#64748B' },
      { id: 'general_surgery', name: 'GENERAL SURGERY', monthly: [4, 4, 2, 1, 1, 2, 0, 0, 0, 2, 59, 54], total: 129, color: '#E84A2D' },
      { id: 'emergency', name: 'EMERGENCY', monthly: [2, 3, 0, 1, 0, 0, 6, 6, 3, 1, 0, 4], total: 26, color: '#EF4444' }
    ],
    inpatient: [
      { id: 'gyne_obs_in', name: 'GYNE & OBS', monthly: [169, 141, 155, 173, 182, 183, 168, 188, 132, 166, 157, 188], total: 2002, color: '#EC4899' },
      { id: 'pediatrics_in', name: 'PEDIATRICS', monthly: [79, 80, 67, 59, 82, 85, 43, 59, 72, 66, 116, 100], total: 908, color: '#2563EB' },
      { id: 'orthopedic_in', name: 'ORTHOPEDIC', monthly: [55, 46, 50, 56, 74, 48, 78, 67, 51, 63, 60, 53], total: 701, color: '#F97316' },
      { id: 'ent_service_in', name: 'ENT SERVICE', monthly: [31, 26, 27, 48, 24, 27, 26, 35, 23, 11, 9, 23], total: 310, color: '#6366F1' },
      { id: 'internal_med_in', name: 'INTERNAL MED', monthly: [17, 15, 25, 24, 28, 32, 25, 18, 19, 26, 27, 25], total: 281, color: '#8B5CF6' },
      { id: 'urology_surgery_in', name: 'UROLOGY SURGERY', monthly: [1, 6, 1, 7, 3, 7, 7, 3, 2, 1, 3, 5], total: 46, color: '#64748B' },
      { id: 'neonatology_in', name: 'NEONATOLOGY', monthly: [6, 0, 1, 2, 0, 5, 2, 3, 4, 2, 3, 7], total: 35, color: '#14B8A6' },
      { id: 'gn_surgery_in', name: 'GN SURGERY', monthly: [1, 0, 1, 1, 2, 1, 0, 0, 0, 2, 4, 7], total: 19, color: '#E84A2D' }
    ],
    maternity: {
      annual: { deliveries: 1437, liveBirths: 1439, cs: 824, svd: 613, csRate: 57.3 },
      monthly: [],
      historical: [
        { year: '2019', deliveries: 109, liveBirths: 149, cs: 99, svd: 10 },
        { year: '2020', deliveries: 669, liveBirths: 664, cs: 457, svd: 212 },
        { year: '2021', deliveries: 898, liveBirths: 900, cs: 547, svd: 351 },
        { year: '2022', deliveries: 1141, liveBirths: 1137, cs: 738, svd: 403 },
        { year: '2023', deliveries: 1252, liveBirths: 1248, cs: 779, svd: 473 },
        { year: '2024', deliveries: 1441, liveBirths: 1444, cs: 854, svd: 587 },
        { year: '2025', deliveries: 1437, liveBirths: 1439, cs: 824, svd: 613 },
        { year: '2026 (9M)', deliveries: 1115, liveBirths: 1126, cs: 623, svd: 492 }
      ]
    },
    mortality: []
  }
};

// Master 8-Year Multi-Year Historical Matrix (2019 - 2026)
const HISTORICAL_MULTI_YEAR = {
  years: ['2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026 (9M)'],
  patientYears: ['2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026 (9M)'],
  annualTotals: {
    opd: [12540, 34068, 66165, 67518, 75117, 85646, 101576, 72334], // sum = 514,964
    ipd: [569, 1662, 2265, 3045, 3255, 3686, 4302, 3495],            // sum = 22,279
    deliveries: [109, 669, 898, 1141, 1252, 1441, 1437, 1115],       // sum = 8,062
    liveBirths: [149, 664, 900, 1137, 1248, 1444, 1439, 1126],       // sum = 8,107
    cs: [99, 457, 547, 738, 779, 854, 824, 623],                    // sum = 4,921
    svd: [10, 212, 351, 403, 473, 587, 613, 492]                    // sum = 3,141
  },
  cumulativeSummary: {
    totalOPD: 514964,
    totalIPD: 22279,
    totalDeliveries: 8062,
    totalLiveBirths: 8107,
    totalCS: 4921,
    totalSVD: 3141,
    totalPatientsServed: 537243, // 514,964 OPD + 22,279 IPD
    allTimeCSRate: 61.0
  },
  opdDepartments: [
    { name: 'PEDIATRICS', years: [3446, 9210, 15192, 19386, 18868, 17527, 19455, 12878], total: 115962, color: '#2563EB' },
    { name: 'GENERAL MED', years: [2762, 4777, 6280, 10314, 12190, 14824, 17512, 12776], total: 81435, color: '#10B981' },
    { name: 'GYNE & OBS', years: [2092, 6845, 11039, 10588, 12375, 13808, 13650, 10126], total: 80523, color: '#EC4899' },
    { name: 'INTERNAL MED', years: [2138, 4888, 17427, 6201, 8339, 10672, 13515, 9565], total: 72745, color: '#8B5CF6' },
    { name: 'PHYSIOTHERAPY', years: [57, 2087, 6411, 8307, 10732, 14132, 18064, 11644], total: 71434, color: '#06B6D4' },
    { name: 'DENTAL', years: [1308, 2734, 3318, 3454, 3970, 4044, 4529, 2965], total: 26322, color: '#14B8A6' },
    { name: 'VACCINATION', years: [0, 1063, 2190, 2915, 3890, 4819, 5401, 4121], total: 24399, color: '#F59E0B' },
    { name: 'ORTHOPEDIC SURGERY', years: [0, 1132, 2485, 3139, 3292, 3573, 5302, 3854], total: 22777, color: '#F97316' },
    { name: 'ENT SERVICE', years: [191, 412, 295, 401, 299, 779, 1690, 1817], total: 5884, color: '#6366F1' },
    { name: 'OPHTALMOLOGY', years: [312, 175, 171, 216, 204, 638, 1372, 1325], total: 4413, color: '#0EA5E9' },
    { name: 'UROLOGY', years: [98, 557, 493, 280, 352, 313, 931, 392], total: 3416, color: '#64748B' },
    { name: 'EMERGENCY', years: [0, 0, 662, 1996, 82, 75, 26, 23], total: 2864, color: '#EF4444' },
    { name: 'GENERAL SURGERY', years: [136, 188, 202, 321, 524, 442, 129, 848], total: 2790, color: '#E84A2D' }
  ],
  ipdDepartments: [
    { name: 'GYNE & OBS', years: [282, 923, 1145, 1491, 1637, 1841, 2002, 1454], total: 10775, color: '#EC4899' },
    { name: 'PEDIATRICS', years: [181, 425, 546, 847, 800, 909, 908, 736], total: 5352, color: '#2563EB' },
    { name: 'ORTHOPEDIC SURGERY', years: [0, 60, 295, 428, 459, 465, 701, 458], total: 2866, color: '#F97316' },
    { name: 'INTERNAL MED', years: [86, 148, 124, 194, 223, 215, 281, 236], total: 1507, color: '#8B5CF6' },
    { name: 'ENT SERVICE', years: [0, 0, 30, 6, 49, 184, 310, 450], total: 1029, color: '#6366F1' },
    { name: 'UROLOGY', years: [0, 57, 97, 37, 49, 23, 46, 27], total: 336, color: '#64748B' },
    { name: 'GENERAL SURGERY', years: [20, 48, 24, 39, 34, 35, 19, 95], total: 314, color: '#E84A2D' },
    { name: 'NEONATOLOGY', years: [0, 0, 0, 0, 0, 0, 35, 37], total: 72, color: '#14B8A6' },
    { name: 'DENTAL', years: [0, 1, 4, 3, 4, 14, 0, 0], total: 26, color: '#10B981' },
    { name: 'EMERGENCY', years: [0, 0, 0, 0, 0, 0, 0, 2], total: 2, color: '#EF4444' }
  ]
};

// Dynamically build individual annual datasets for 2019 through 2024 from validated matrix
['2019', '2020', '2021', '2022', '2023', '2024'].forEach((yr, idx) => {
  const opdList = HISTORICAL_MULTI_YEAR.opdDepartments.map(dept => ({
    id: dept.name.toLowerCase().replace(/[^a-z0-9]/g, '_'),
    name: dept.name,
    monthly: [dept.years[idx]],
    total: dept.years[idx],
    color: dept.color
  })).filter(d => d.total > 0);

  const ipdList = HISTORICAL_MULTI_YEAR.ipdDepartments.map(ward => ({
    id: ward.name.toLowerCase().replace(/[^a-z0-9]/g, '_'),
    name: ward.name,
    monthly: [ward.years[idx]],
    total: ward.years[idx],
    color: ward.color
  })).filter(w => w.total > 0);

  const matTotals = HISTORICAL_MULTI_YEAR.annualTotals;
  const del = matTotals.deliveries[idx];
  const live = matTotals.liveBirths[idx];
  const cs = matTotals.cs[idx];
  const svd = matTotals.svd[idx];
  const csRate = del > 0 ? parseFloat(((cs / del) * 100).toFixed(1)) : 0;

  HOSPITAL_DATASETS[yr] = {
    metadata: {
      hospitalName: 'Dream Medical Center Hospital',
      reportTitle: `Hospital Statistics & Clinical Performance Report ${yr}`,
      year: parseInt(yr, 10),
      periodCovered: `January - December ${yr} (Full Year)`,
      periodBadge: `Jan 01 – Dec 31, ${yr}`,
      months: [`Full Year ${yr}`],
      allMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      lastUpdated: '2026-09-07',
      sourceFile: 'PATIENTS NUMBERS 2019-2024.xlsx',
      hasMortality: false,
      hasMonthlyMaternity: false,
      isAnnualOnly: true
    },
    outpatient: opdList,
    inpatient: ipdList,
    maternity: {
      annual: { deliveries: del, liveBirths: live, cs: cs, svd: svd, csRate: csRate },
      monthly: [],
      historical: [
        { year: '2019', deliveries: 109, liveBirths: 149, cs: 99, svd: 10 },
        { year: '2020', deliveries: 669, liveBirths: 664, cs: 457, svd: 212 },
        { year: '2021', deliveries: 898, liveBirths: 900, cs: 547, svd: 351 },
        { year: '2022', deliveries: 1141, liveBirths: 1137, cs: 738, svd: 403 },
        { year: '2023', deliveries: 1252, liveBirths: 1248, cs: 779, svd: 473 },
        { year: '2024', deliveries: 1441, liveBirths: 1444, cs: 854, svd: 587 },
        { year: '2025', deliveries: 1437, liveBirths: 1439, cs: 824, svd: 613 },
        { year: '2026 (9M)', deliveries: 1115, liveBirths: 1126, cs: 623, svd: 492 }
      ]
    },
    mortality: []
  };
});

// Active working year dataset (defaults to 2026)
let currentActiveYear = '2026';
let hospitalData = JSON.parse(JSON.stringify(HOSPITAL_DATASETS['2026']));

/**
 * Hospital Analytics Engine
 */
const HospitalAnalytics = {
  // Set the active year dataset
  setActiveYear(yearKey) {
    if (HOSPITAL_DATASETS[yearKey]) {
      currentActiveYear = yearKey;
      hospitalData = JSON.parse(JSON.stringify(HOSPITAL_DATASETS[yearKey]));
      return true;
    }
    return false;
  },

  getActiveYear() {
    return currentActiveYear;
  },

  getDataset(yearKey) {
    return HOSPITAL_DATASETS[yearKey] || hospitalData;
  },

  getHistoricalData() {
    return HISTORICAL_MULTI_YEAR;
  },

  // Outpatient total for current active year
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

  // Inpatient total for current active year
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

  // Total Deliveries
  getDeliveriesTotal(periodFilter) {
    if (hospitalData.maternity && hospitalData.maternity.annual) {
      return hospitalData.maternity.annual.deliveries;
    }
    const list = hospitalData.maternity.monthly || [];
    if (Array.isArray(periodFilter)) {
      return periodFilter.reduce((acc, idx) => acc + (list[idx] ? list[idx].deliveries : 0), 0);
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      return list[periodFilter] ? list[periodFilter].deliveries : 0;
    }
    return list.reduce((acc, m) => acc + m.deliveries, 0);
  },

  // Total Live Births
  getLiveBirthsTotal(periodFilter) {
    if (hospitalData.maternity && hospitalData.maternity.annual) {
      return hospitalData.maternity.annual.liveBirths;
    }
    const list = hospitalData.maternity.monthly || [];
    if (Array.isArray(periodFilter)) {
      return periodFilter.reduce((acc, idx) => acc + (list[idx] ? list[idx].liveBirths : 0), 0);
    } else if (typeof periodFilter === 'number' && periodFilter >= 0) {
      return list[periodFilter] ? list[periodFilter].liveBirths : 0;
    }
    return list.reduce((acc, m) => acc + m.liveBirths, 0);
  },

  // C-Section vs SVD stats
  getCSStats(periodFilter) {
    if (hospitalData.maternity && hospitalData.maternity.annual) {
      const ann = hospitalData.maternity.annual;
      return { cs: ann.cs, svd: ann.svd, total: ann.deliveries, rate: ann.csRate };
    }
    const list = hospitalData.maternity.monthly || [];
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

    const rate = total > 0 ? ((cs / total) * 100).toFixed(1) : '0.0';
    return { cs, svd, total, rate: parseFloat(rate) };
  },

  // In-Facility Deaths Total
  getDeathsTotal(periodFilter) {
    if (!hospitalData.mortality || hospitalData.mortality.length === 0) return 0;
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

  getMortalityTotal(periodFilter) {
    return this.getDeathsTotal(periodFilter);
  },

  getMortalityRate(periodFilter) {
    const totalIPD = this.getIPDTotal(periodFilter);
    const totalDeaths = this.getDeathsTotal(periodFilter);
    return totalIPD > 0 ? ((totalDeaths / totalIPD) * 100).toFixed(2) : '0.00';
  }
};

window.hospitalData = hospitalData;
window.HOSPITAL_DATASETS = HOSPITAL_DATASETS;
window.HISTORICAL_MULTI_YEAR = HISTORICAL_MULTI_YEAR;
window.HospitalAnalytics = HospitalAnalytics;
