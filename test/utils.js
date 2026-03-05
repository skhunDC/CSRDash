function getWeekStartISO(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d.toISOString().slice(0, 10);
}

function fmtCurrency(n) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0
  }).format(Number(n || 0));
}

function fmtPercent(n) {
  return `${(Number(n || 0) * 100).toFixed(1)}%`;
}

function buildSalesComparison(rows, weekStart) {
  const grouped = {};
  rows
    .filter((r) => r.weekStart === weekStart)
    .forEach((row) => {
      if (!grouped[row.store]) grouped[row.store] = { store: row.store, sales2025: 0, sales2026: 0 };
      if (Number(row.year) === 2025) grouped[row.store].sales2025 += Number(row.sales || 0);
      if (Number(row.year) === 2026) grouped[row.store].sales2026 += Number(row.sales || 0);
    });

  return Object.values(grouped).map((item) => {
    const delta = item.sales2026 - item.sales2025;
    return {
      ...item,
      delta,
      deltaPct: item.sales2025 ? delta / item.sales2025 : 0
    };
  });
}

module.exports = { getWeekStartISO, fmtCurrency, fmtPercent, buildSalesComparison };
