const { getWeekStartISO, fmtCurrency, fmtPercent, buildSalesComparison } = require('./utils');

describe('date helpers', () => {
  test('returns monday for mid-week date', () => {
    expect(getWeekStartISO('2026-02-18T10:00:00Z')).toBe('2026-02-16');
  });

  test('returns previous monday for sunday date', () => {
    expect(getWeekStartISO('2026-02-22T10:00:00Z')).toBe('2026-02-16');
  });
});

describe('format helpers', () => {
  test('formats currency', () => {
    expect(fmtCurrency(12345)).toBe('$12,345');
  });

  test('formats percent', () => {
    expect(fmtPercent(0.256)).toBe('25.6%');
  });
});

describe('payload transformations', () => {
  test('builds sales comparison payload shape', () => {
    const rows = [
      { weekStart: '2026-02-16', store: 'Dublin', year: 2025, sales: 100 },
      { weekStart: '2026-02-16', store: 'Dublin', year: 2026, sales: 160 },
      { weekStart: '2026-02-16', store: 'Pleasanton', year: 2025, sales: 300 },
      { weekStart: '2026-02-16', store: 'Pleasanton', year: 2026, sales: 330 }
    ];
    const out = buildSalesComparison(rows, '2026-02-16');
    expect(out).toHaveLength(2);
    expect(out[0]).toEqual(expect.objectContaining({
      store: expect.any(String),
      sales2025: expect.any(Number),
      sales2026: expect.any(Number),
      delta: expect.any(Number),
      deltaPct: expect.any(Number)
    }));
  });
});
