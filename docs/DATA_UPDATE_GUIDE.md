# DMC Hospital Statistics Dashboard — Data Update Guide

Welcome to the **DMC Hospital Executive Statistics & KPI Dashboard (2026)**. This guide outlines how to update the dashboard data, either by uploading an updated Excel file in the browser or modifying the JavaScript dataset directly.

---

## 1. Quick In-Browser Update (Recommended)

The dashboard includes a built-in **SheetJS** Excel parser. You can update all statistics without writing any code:

1. Open the dashboard in your web browser.
2. Click the **"Import Excel"** button in the top-right header toolbar.
3. Drag and drop your updated Excel file (e.g., `PATIENTS_STATISTICS JAN TO JULY 2026.xlsx` or any future monthly update) into the upload area.
4. The dashboard will instantly:
   - Read all sheets (`OUT PATIENT PER DEPARTEMENT`, `INPATIENT PER DEPARTEMENT`, `BABIES TOTAL`, `Death`).
   - Recalculate totals, bed occupancy rates, and C-Section ratios.
   - Dynamically re-render all primary donut charts, trend lines, sparklines, and data tables.
   - Show a toast notification confirming the number of parsed departments and deliveries.

---

## 2. Expected Excel Sheet Structure

When preparing your Excel workbook, ensure the following sheet naming and structure conventions are preserved:

### Sheet 1: `OUT PATIENT PER DEPARTEMENT`
- **Row 3**: Header containing `DEPARTMENTS` followed by monthly date headers (e.g. `2026-01-01`, `2026-02-01`, up to `2026-12-01`).
- **Rows 4–16**: Department names in Column A (e.g. `INTERNAL MED`, `PEDIATRICS`, `GENERAL MED`, `GYNE & OBS`, `PHYSIOTHERAPY`, etc.) with monthly counts across Columns C through J.
- **Row 17**: Total row.

### Sheet 2: `INPATIENT PER DEPARTEMENT`
- **Row 2**: Header containing `INPATIENT BY DEPARTMENTS` followed by monthly date headers.
- **Rows 3–11**: Ward names in Column A (e.g. `GYNE & OBS`, `PEDIATRICS`, `ORTHOPEDIC`, `INTERNAL MED`, `ENT SERVICE`, etc.) with monthly admissions.
- **Row 12**: Total row.

### Sheet 3: `BABIES TOTAL`
- **Row 1**: Headers: `period`, `Deliveries at health facility`, `births live`, `Death`, `Circumstances of death`, `Deliveries by Caesarean Section`, `SVD`.
- **Rows 2–8**: Monthly records for the current year.
- **Rows 17–26**: Multi-year historical records (2019 to 2026) for longitudinal delivery tracking.

### Sheet 4: `Death`
- **Row 1**: Headers: `period`, `Death at the HF`, `Age`, `Departement`, `Circonstance`.
- **Rows 2–8**: In-facility mortality records with date, count, age, department, and cause.

---

## 3. Direct Code Update (for Pre-bundled Offline Data)

If you wish to update the default dataset bundled with the dashboard permanently:

1. Open [`js/data.js`](../js/data.js) in your code editor.
2. Locate the `HOSPITAL_RAW_DATA` object:
   - **`outpatient`**: Array of departments with their monthly counts and totals.
   - **`inpatient`**: Array of wards with monthly admissions, assigned beds, and average length of stay (`alos`).
   - **`maternity.monthly`**: Array of monthly delivery metrics (`deliveries`, `liveBirths`, `cs`, `svd`, `deaths`).
   - **`mortality`**: Array of individual clinical death records.
   - **`bedCapacity.wards`**: Dedicated bed counts per ward.
3. Save the file. When you reload `index.html`, the new values will be loaded automatically.

---

## 4. Key Metric Calculations

| Metric | Formula |
|---|---|
| **Bed Occupancy Rate** | $\frac{\sum (\text{Inpatient Admissions} \times \text{ALOS})}{\text{Total Beds} \times \text{Days in Period}} \times 100$ |
| **C-Section Rate** | $\frac{\text{Caesarean Deliveries}}{\text{Total Facility Deliveries}} \times 100$ |
| **Inpatient Mortality Rate** | $\frac{\text{Total Facility Deaths}}{\text{Total Inpatient Admissions}} \times 100$ |
| **Average Length of Stay (ALOS)** | Weighted mean of inpatient stay durations across all 9 wards |

---

## 5. Exporting & Printing Executive Reports

- Click **"Export PDF"** to open your browser's print dialog.
- The `@media print` stylesheet automatically hides navigation bars, filters, and buttons, producing a crisp, multi-page executive report.
- Select **"Save as PDF"** in the destination dropdown to generate the report file.
