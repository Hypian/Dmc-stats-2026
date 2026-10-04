# Agent Instructions — Dream Medical Center (DMC) Hospital Statistics Dashboard

## 1. Project Mission & Identity
This project is the official executive clinical performance dashboard for **Dream Medical Center (DMC) Hospital**.
- **Institution**: Dream Medical Center Hospital
- **Theme**: Dream Red & Warm White / Dark Charcoal medical dashboard (`#E84A2D` primary, `#3D3532` dark surface, `#FAF6F5` light surface)
- **Primary Data Sources**: Microsoft Excel registers covering 2019 through 2026.
  - Multi-sheet 2026 Monthly Clinical Workbook (`PATIENTS_STATISTICS  2026.xlsx`)
  - 2025 Annual Department Workbook (`PATIENTS NUMBERS 2025.xlsx`)
  - 2019–2024 Multi-Year Longitudinal Workbook (`PATIENTS NUMBERS 2019-2024.xlsx`)
  - Longitudinal Payer Distribution Workbook (`insurance distribution from november 2018 to september 17.xlsx`)

---

## 2. Technology Stack & Hard Constraints
1. **Core Runtime**: Pure Vanilla HTML5 and ES6+ JavaScript. No front-end build steps (no React, no Vue, no Angular, no Webpack).
2. **Styling**: Pure Vanilla CSS organized modularly in `css/`:
   - [main.css](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/css/main.css): Global variables, theme tokens (light/dark), reset, typography, and page shell.
   - [components.css](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/css/components.css): KPI metric cards, donut containers, tables, buttons, modals, badges, filters.
   - [responsive.css](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/css/responsive.css): Media queries for mobile/tablet responsive layouts.
   - [print.css](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/css/print.css): Executive printer & PDF snapshot layout.
   - **DO NOT** introduce TailwindCSS or CSS preprocessors.
3. **Libraries Loaded via CDN**:
   - `Chart.js 4.4.1` (UMD version)
   - `SheetJS / xlsx 0.18.5` (UMD version)
   - `FontAwesome 6.4.2`
   - Google Fonts: `Inter`, `JetBrains Mono`, `Poppins`

---

## 3. Architecture & File Responsibilities
- [index.html](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/index.html): Semantic layout containing sidebar navigation, executive header, 6 tab panels (`overview`, `inpatient`, `outpatient`, `maternity`, `mortality`, `insurers`, `historical`), upload modal, and department drilldown modal.
- [js/data.js](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/js/data.js): Default verified snapshot and historical multi-year records. Serves as reliable fallback when no file is uploaded.
- [js/insurer-data.js](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/js/insurer-data.js): Asynchronous parser for historical insurer distribution workbook, merging RHIC/non-RHIC duplicate entities.
- [js/excel-parser.js](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/js/excel-parser.js): In-browser SheetJS workbook reader that detects workbook schemas, maps sheets, extracts clinical metrics, and updates `hospitalData`.
- [js/charts.js](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/js/charts.js): Visualization engine creating Chart.js instances, custom donut center text plugin, custom legends, and theme-adaptive coloring.
- [js/app.js](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/js/app.js): Application controller managing tabs, period filters, year selector, table rendering, KPI cards, theme toggling, live clock, and drilldown modals.

---

## 4. Donut Chart Design Standards
All donut charts in this dashboard must adhere to the following design system tokens:
1. **Container Wrapper**: Must be enclosed in `<div class="donut-container"><canvas id="..."></canvas></div>`.
2. **Ring Proportions**:
   - `type: 'doughnut'`
   - `cutout: '66%'`
   - `maintainAspectRatio: true`
   - `responsive: true`
   - `borderWidth: 2`
   - `borderColor: isDark ? '#1E1A18' : '#FFFFFF'`
   - `hoverOffset: 8`
   - `animation: { duration: 800 }`
3. **Cursor & Hover Interaction**:
   - Must implement `onHover: (evt, el) => { if (evt.native && evt.native.target) evt.native.target.style.cursor = el.length ? 'pointer' : 'default'; }`
4. **Donut Center Text Plugin (`donutCenterText`)**:
   - Must specify:
     ```javascript
     donutCenterText: {
       display: true,
       value: formattedValue, // bold JetBrains Mono 26px
       subtext: 'UPPERCASE SUBTEXT', // Poppins 700 11px
       valueColor: isDark ? '#FF5E42' : '#E84A2D'
     }
     ```
5. **Tooltips**:
   - `plugins.legend.display: false` (Chart.js native legend must be hidden).
   - Tooltips must format count with thousand separators and show relative percentage `(xx.x%)`.
6. **Custom Legend (`renderCustomLegend`)**:
   - Render into adjacent container `<div id="...Legend" class="donut-legend-custom"></div>`.
   - Displays color dot, title, monospace count, and percentage badge.
   - Rows trigger `window.showDepartmentDrilldown(type, label)` on click (except insurer items).

---

## 5. Agent Workflow Guidelines
- **Modifying Data**: When updating clinical numbers or sheet schemas, verify against [docs/DATA_UPDATE_GUIDE.md](file:///c:/Users/kagin/OneDrive/Desktop/My%20Projects/Dmc%20stats%202026/docs/DATA_UPDATE_GUIDE.md).
- **Adding New Visualizations**: Register the Chart instance in `HospitalCharts.instances[key]`, ensure `destroy()` is called on re-render, and listen to theme changes.
- **Testing**: When making DOM changes, ensure all element IDs referenced in `app.js` and `charts.js` exist in `index.html`.
