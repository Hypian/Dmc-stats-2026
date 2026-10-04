/** Insurer distribution source: November 2018 through September 17, 2026. */
const INSURER_DISTRIBUTION = { metadata: { periodCovered: 'November 2018 – September 17, 2026', sourceFile: 'insurance distribution from november 2018 to september 17.xlsx' }, inpatient: [], outpatient: [] };
(async () => {
  try {
    const response = await fetch(encodeURI(INSURER_DISTRIBUTION.metadata.sourceFile));
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: null });
    const records = (start, end) => rows.slice(start - 1, end).map(row => [String(row[0] || '').trim(), Number(row[1]) || 0]).filter(([name, value]) => name && value > 0);
    INSURER_DISTRIBUTION.inpatient = records(3, 130);
    INSURER_DISTRIBUTION.outpatient = records(134, 287);
    window.dispatchEvent(new CustomEvent('insurerDataLoaded'));
  } catch (error) { console.warn('Unable to load insurer distribution workbook:', error); }
})();
