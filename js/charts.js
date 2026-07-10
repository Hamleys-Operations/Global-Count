/* ==========================================================================
   Hamleys Global Count Dashboard — charts.js
   Thin wrapper around Chart.js: creates/updates/destroys chart instances,
   applies Hamleys theme colors, handles dark-mode refresh & PNG export.
   ========================================================================== */

'use strict';

const ChartsLib = (() => {
  const registry = {}; // canvasId -> Chart instance

  const PALETTE = ['#d71920', '#d4af37', '#2f7fd1', '#17a673', '#8e44ad', '#e67e22', '#16a2b8', '#c0392b', '#7f8c8d', '#2c3e50'];

  function isDark() { return document.documentElement.getAttribute('data-theme') === 'dark'; }
  function textColor() { return isDark() ? '#eef0f5' : '#1f2430'; }
  function gridColor() { return isDark() ? 'rgba(255,255,255,0.06)' : 'rgba(20,20,40,0.06)'; }

  function baseOptions(extra) {
    return Object.assign({
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 700, easing: 'easeOutQuart' },
      plugins: {
        legend: { labels: { color: textColor(), boxWidth: 12, font: { size: 11 } } },
        tooltip: {
          backgroundColor: isDark() ? '#20242f' : '#1f2430',
          titleColor: '#fff', bodyColor: '#fff',
          padding: 10, cornerRadius: 8, displayColors: true,
        }
      },
      scales: {
        x: { ticks: { color: textColor(), font: { size: 10 } }, grid: { color: gridColor() } },
        y: { ticks: { color: textColor(), font: { size: 10 } }, grid: { color: gridColor() }, beginAtZero: true }
      }
    }, extra || {});
  }

  function destroy(id) {
    if (registry[id]) { registry[id].destroy(); delete registry[id]; }
  }

  function colorize(datasets, opts) {
    opts = opts || {};
    return datasets.map((ds, i) => Object.assign({
      borderColor: PALETTE[i % PALETTE.length],
      backgroundColor: opts.fill ? hexAlpha(PALETTE[i % PALETTE.length], opts.alpha || 0.18) : (opts.solid ? datasets.length > 1 ? PALETTE.map((c, j) => hexAlpha(c, 0.85)) : hexAlpha(PALETTE[i % PALETTE.length], 0.85) : PALETTE[i % PALETTE.length]),
      borderWidth: opts.borderWidth ?? 2,
      tension: 0.35,
      fill: !!opts.fill,
      pointRadius: opts.points === false ? 0 : 3,
      pointHoverRadius: 6,
      borderRadius: opts.radius ?? 0,
    }, ds));
  }

  function hexAlpha(hex, alpha) {
    const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
    return `rgba(${r},${g},${b},${alpha})`;
  }

  function lineChart(id, labels, datasets) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    registry[id] = new Chart(el, {
      type: 'line',
      data: { labels, datasets: colorize(datasets, { points: true }) },
      options: baseOptions()
    });
  }

  function areaChart(id, labels, datasets) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    registry[id] = new Chart(el, {
      type: 'line',
      data: { labels, datasets: colorize(datasets, { fill: true, alpha: 0.25 }) },
      options: baseOptions()
    });
  }

  function barChart(id, labels, datasets, diverging) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    registry[id] = new Chart(el, {
      type: 'bar',
      data: { labels, datasets: colorize(datasets, { solid: true, radius: 6 }) },
      options: baseOptions({
        scales: diverging ? {
          x: { ticks: { color: textColor(), font: { size: 10 } }, grid: { color: gridColor() } },
          y: { ticks: { color: textColor() }, grid: { color: gridColor() } }
        } : undefined
      })
    });
  }

  function horizontalBar(id, labels, datasets, gold) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    const ds = datasets.map(d => Object.assign({}, d, {
      backgroundColor: gold ? 'rgba(212,175,55,0.85)' : 'rgba(215,25,32,0.85)',
      borderRadius: 6,
    }));
    registry[id] = new Chart(el, {
      type: 'bar',
      data: { labels, datasets: ds },
      options: baseOptions({ indexAxis: 'y' })
    });
  }

  function pieChart(id, labels, data) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    registry[id] = new Chart(el, {
      type: 'pie',
      data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length]), borderColor: isDark() ? '#181c27' : '#fff', borderWidth: 2 }] },
      options: baseOptions({ scales: undefined })
    });
  }

  function doughnutChart(id, labels, data) {
    const el = document.getElementById(id); if (!el) return;
    destroy(id);
    registry[id] = new Chart(el, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length]), borderColor: isDark() ? '#181c27' : '#fff', borderWidth: 2 }] },
      options: baseOptions({ scales: undefined, cutout: '62%' })
    });
  }

  function downloadChart(id) {
    const chart = registry[id];
    if (!chart) return;
    const a = document.createElement('a');
    a.href = chart.toBase64Image('image/png', 1);
    a.download = `Hamleys_${id}_${Date.now()}.png`;
    a.click();
  }

  function refreshAllThemes() {
    Object.keys(registry).forEach(id => {
      const chart = registry[id];
      if (!chart) return;
      chart.options.plugins.legend.labels.color = textColor();
      if (chart.options.plugins.tooltip) {
        chart.options.plugins.tooltip.backgroundColor = isDark() ? '#20242f' : '#1f2430';
      }
      if (chart.options.scales) {
        Object.values(chart.options.scales).forEach(scale => {
          if (scale.ticks) scale.ticks.color = textColor();
          if (scale.grid) scale.grid.color = gridColor();
        });
      }
      chart.update();
    });
  }

  return { lineChart, areaChart, barChart, horizontalBar, pieChart, doughnutChart, downloadChart, refreshAllThemes };
})();
