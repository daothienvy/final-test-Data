/**
 * POWER BI EXECUTIVE DASHBOARD - APPLICATION CONTROLLER
 * Handles tab transitions, interactive Chart.js charts, live slicers, and action center.
 */

// Global Chart References
let chartP1Trend = null;
let chartP1Donut = null;
let chartP1TopCat = null;

let chartP2SegDonut = null;
let chartP2SegSpend = null;
let chartP2Geo = null;
let chartP2Recency = null;

let chartP3LeadTier = null;
let chartP3CatShipRatio = null;
let chartP3StateLead = null;
let chartP3Weight = null;

let chartP4Forecast = null;

// Application State
let currentTab = 'page1';
let currentCustomerModel = 'ENHANCED';
let filteredActions = [];

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  renderPaymentTable();
  initPage1Charts();
  initPage2Charts();
  initPage3Charts();
  initPage4Charts();
  renderActionTable();
});

// Tab Navigation
function switchTab(tabId) {
  currentTab = tabId;
  document.querySelectorAll('.page-view').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

  const targetPage = document.getElementById(tabId);
  if (targetPage) targetPage.classList.add('active');

  const activeBtn = document.querySelector(`button[onclick="switchTab('${tabId}')"]`);
  if (activeBtn) activeBtn.classList.add('active');

  // Trigger resize on visible charts to ensure crisp rendering
  window.dispatchEvent(new Event('resize'));
}

// Format Currency
function formatCurrency(val) {
  if (val >= 1000000) return '$' + (val / 1000000).toFixed(2) + 'M';
  if (val >= 1000) return '$' + (val / 1000).toFixed(1) + 'K';
  return '$' + val.toFixed(2);
}

// -------------------------------------------------------------
// TRANG 1: SALES & REVENUE CHARTS
// -------------------------------------------------------------
function initPage1Charts() {
  const months = DASHBOARD_DATA.monthly.map(m => m.order_month);
  const revenues = DASHBOARD_DATA.monthly.map(m => m.revenue);
  const aovs = DASHBOARD_DATA.monthly.map(m => m.aov);

  // 1. Combo Trend Chart
  const ctxTrend = document.getElementById('chart-p1-trend').getContext('2d');
  chartP1Trend = new Chart(ctxTrend, {
    type: 'bar',
    data: {
      labels: months,
      datasets: [
        {
          label: 'Doanh thu Hàng tháng ($)',
          data: revenues,
          backgroundColor: 'rgba(37, 99, 235, 0.75)',
          borderColor: '#2563EB',
          borderWidth: 1,
          borderRadius: 4,
          yAxisID: 'yRev'
        },
        {
          label: 'Giá trị Đơn hàng TB AOV ($)',
          data: aovs,
          type: 'line',
          borderColor: '#F59E0B',
          backgroundColor: '#F59E0B',
          borderWidth: 2.5,
          pointRadius: 3,
          tension: 0.3,
          yAxisID: 'yAov'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { family: 'Inter', size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => ctx.dataset.label.includes('AOV') 
              ? `AOV: $${ctx.raw.toFixed(2)}`
              : `Doanh thu: $${ctx.raw.toLocaleString()}`
          }
        }
      },
      scales: {
        yRev: {
          type: 'linear',
          position: 'left',
          grid: { color: '#F1F5F9' },
          ticks: { callback: (val) => formatCurrency(val) }
        },
        yAov: {
          type: 'linear',
          position: 'right',
          grid: { drawOnChartArea: false },
          ticks: { callback: (val) => `$${val}` }
        },
        x: { grid: { display: false }, ticks: { maxRotation: 45, minRotation: 45 } }
      }
    }
  });

  // 2. Category Donut Chart
  const catLabels = DASHBOARD_DATA.categories.slice(0, 6).map(c => c.product_category_name);
  const catRevs = DASHBOARD_DATA.categories.slice(0, 6).map(c => c.revenue);
  const ctxDonut = document.getElementById('chart-p1-donut').getContext('2d');
  chartP1Donut = new Chart(ctxDonut, {
    type: 'doughnut',
    data: {
      labels: catLabels,
      datasets: [{
        data: catRevs,
        backgroundColor: [
          '#1E3A8A', '#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6'
        ],
        borderWidth: 2,
        borderColor: '#FFFFFF'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.label}: $${ctx.raw.toLocaleString()} (${((ctx.raw / DASHBOARD_DATA.kpi.revenue) * 100).toFixed(1)}%)`
          }
        }
      },
      cutout: '65%'
    }
  });

  // 3. Top 10 Categories Bar Chart
  const top10 = DASHBOARD_DATA.categories.slice(0, 10);
  const ctxTopCat = document.getElementById('chart-p1-topcat').getContext('2d');
  chartP1TopCat = new Chart(ctxTopCat, {
    type: 'bar',
    data: {
      labels: top10.map(c => c.product_category_name),
      datasets: [{
        label: 'Doanh thu ($)',
        data: top10.map(c => c.revenue),
        backgroundColor: '#1E3A8A',
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => `Doanh thu: $${ctx.raw.toLocaleString()}` } }
      },
      scales: {
        x: { ticks: { callback: (val) => formatCurrency(val) }, grid: { color: '#F1F5F9' } },
        y: { grid: { display: false } }
      }
    }
  });
}

function renderPaymentTable() {
  const tbody = document.getElementById('tbody-p1-payment');
  tbody.innerHTML = '';
  DASHBOARD_DATA.payments.forEach(p => {
    const share = ((p.revenue / DASHBOARD_DATA.kpi.revenue) * 100).toFixed(1);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${p.payment_type.toUpperCase()}</strong></td>
      <td>${p.orders.toLocaleString()}</td>
      <td><strong>$${p.revenue.toLocaleString()}</strong></td>
      <td><span class="badge-tag badge-positive">${share}%</span></td>
      <td>$${p.aov.toFixed(2)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// -------------------------------------------------------------
// TRANG 2: CUSTOMER SEGMENTATION CHARTS
// -------------------------------------------------------------
function initPage2Charts() {
  // 1. Segmentation Donut
  const ctxSeg = document.getElementById('chart-p2-seg-donut').getContext('2d');
  const segData = getSegmentData();

  chartP2SegDonut = new Chart(ctxSeg, {
    type: 'doughnut',
    data: {
      labels: segData.labels,
      datasets: [{
        data: segData.values,
        backgroundColor: ['#10B981', '#2563EB', '#F59E0B', '#EF4444'],
        borderWidth: 2,
        borderColor: '#FFFFFF'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.label}: ${ctx.raw.toLocaleString()} khách (${((ctx.raw / DASHBOARD_DATA.kpi.customers) * 100).toFixed(1)}%)`
          }
        }
      },
      cutout: '60%'
    }
  });

  // 2. Average Spend by Segment
  const ctxSpend = document.getElementById('chart-p2-seg-spend').getContext('2d');
  chartP2SegSpend = new Chart(ctxSpend, {
    type: 'bar',
    data: {
      labels: ['1. VIP High-Value', '2. Loyal / High-Tier', '3. Potential', '4. Lost / At-Risk'],
      datasets: [{
        label: 'Chi tiêu Trung bình ($ / Khách)',
        data: [1538.45, 519.57, 187.32, 114.60],
        backgroundColor: ['#10B981', '#2563EB', '#F59E0B', '#EF4444'],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => `Mức chi tiêu TB: $${ctx.raw.toFixed(2)}` } }
      },
      scales: {
        y: { ticks: { callback: (val) => `$${val}` }, grid: { color: '#F1F5F9' } },
        x: { grid: { display: false } }
      }
    }
  });

  // 3. Top States Bar
  const topStates = DASHBOARD_DATA.states.slice(0, 8);
  const ctxGeo = document.getElementById('chart-p2-geo').getContext('2d');
  chartP2Geo = new Chart(ctxGeo, {
    type: 'bar',
    data: {
      labels: topStates.map(s => s.customer_state),
      datasets: [{
        label: 'Doanh thu ($)',
        data: topStates.map(s => s.revenue),
        backgroundColor: '#1E3A8A',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => `Doanh thu: $${ctx.raw.toLocaleString()}` } }
      },
      scales: {
        y: { ticks: { callback: (val) => formatCurrency(val) }, grid: { color: '#F1F5F9' } },
        x: { grid: { display: false } }
      }
    }
  });

  // 4. Recency Distribution
  const ctxRecency = document.getElementById('chart-p2-recency').getContext('2d');
  chartP2Recency = new Chart(ctxRecency, {
    type: 'pie',
    data: {
      labels: ['Active (≤30d)', 'Warm (31-90d)', 'Inactive (>90d)'],
      datasets: [{
        data: [5695, 10304, 73317],
        backgroundColor: ['#10B981', '#F59E0B', '#EF4444'],
        borderWidth: 2,
        borderColor: '#FFFFFF'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.label}: ${ctx.raw.toLocaleString()} (${((ctx.raw / 89316) * 100).toFixed(1)}%)`
          }
        }
      }
    }
  });
}

function getSegmentData() {
  if (currentCustomerModel === 'STANDARD') {
    return {
      labels: ['1. VIP Customer', '2. Loyal Customer', '3. Potential Customer', '4. Lost / At-Risk'],
      values: [0, 0, 15999, 73317]
    };
  } else {
    return {
      labels: ['1. VIP High-Value', '2. Loyal / High-Tier', '3. Potential', '4. Lost / At-Risk'],
      values: [5281, 14341, 11337, 58357]
    };
  }
}

function toggleCustomerModel() {
  const select = document.getElementById('slicer-model-view');
  currentCustomerModel = select.value;

  const segData = getSegmentData();
  chartP2SegDonut.data.labels = segData.labels;
  chartP2SegDonut.data.datasets[0].data = segData.values;
  chartP2SegDonut.update();

  const desc = document.getElementById('desc-p2-seg');
  const vipVal = document.getElementById('kpi-p2-vip-val');
  const vipSub = document.getElementById('kpi-p2-vip-sub');
  const lostVal = document.getElementById('kpi-p2-lost-val');

  if (currentCustomerModel === 'STANDARD') {
    desc.textContent = 'Mô hình Barem Chuẩn Đề bài (Standard DAX Section 6.1)';
    vipVal.textContent = '0';
    vipSub.innerHTML = '<span class="badge-tag badge-alert">0 khách (Single Order token)</span>';
    lostVal.textContent = '73,317';
  } else {
    desc.textContent = 'Mô hình Phân loại Giá trị Tích lũy (Enhanced Monetary Model)';
    vipVal.textContent = '5,281';
    vipSub.innerHTML = '<span class="badge-tag badge-positive">Đóng góp 26.7% doanh thu</span> (Chi tiêu > $1k)';
    lostVal.textContent = '58,357';
  }
}

// -------------------------------------------------------------
// TRANG 3: SUPPLY CHAIN & LOGISTICS CHARTS
// -------------------------------------------------------------
function initPage3Charts() {
  // 1. Delivery Tier Clustered Bar
  const ctxTier = document.getElementById('chart-p3-lead-tier').getContext('2d');
  chartP3LeadTier = new Chart(ctxTier, {
    type: 'bar',
    data: {
      labels: ['1. Express (≤0d)', '2. Standard (1-3d)', '3. Delayed (>3d)', 'Not Shipped/Cancelled'],
      datasets: [{
        label: 'Số lượng Đơn hàng',
        data: [0, 4378, 83049, 1889],
        backgroundColor: ['#10B981', '#2563EB', '#EF4444', '#94A3B8'],
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `Số đơn: ${ctx.raw.toLocaleString()} (${((ctx.raw / 89316) * 100).toFixed(1)}%)`
          }
        }
      },
      scales: {
        y: { ticks: { callback: (val) => val.toLocaleString() }, grid: { color: '#F1F5F9' } },
        x: { grid: { display: false } }
      }
    }
  });

  // 2. Shipping Ratio % by Category
  const topCats = DASHBOARD_DATA.categories.slice(0, 10);
  const ctxShipRatio = document.getElementById('chart-p3-cat-ship-ratio').getContext('2d');
  chartP3CatShipRatio = new Chart(ctxShipRatio, {
    type: 'bar',
    data: {
      labels: topCats.map(c => c.product_category_name),
      datasets: [{
        label: 'Tỷ lệ Chi phí Shipping (%)',
        data: topCats.map(c => c.shipping_ratio),
        backgroundColor: topCats.map(c => c.shipping_ratio > 15 ? '#EF4444' : (c.shipping_ratio < 8 ? '#10B981' : '#F59E0B')),
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `Shipping Ratio: ${ctx.raw}% (${ctx.raw > 15 ? 'BÁO ĐỘNG ĐỎ' : 'Bình thường'})`
          }
        }
      },
      scales: {
        y: {
          ticks: { callback: (val) => `${val}%` },
          grid: { color: '#F1F5F9' },
          suggestedMax: 20
        },
        x: { ticks: { maxRotation: 45, minRotation: 45 } }
      }
    }
  });

  // 3. Lead Time by State
  const topStates = DASHBOARD_DATA.states.slice(0, 10);
  const ctxStateLead = document.getElementById('chart-p3-state-lead').getContext('2d');
  chartP3StateLead = new Chart(ctxStateLead, {
    type: 'bar',
    data: {
      labels: topStates.map(s => s.customer_state),
      datasets: [{
        label: 'Thời gian Giao TB (Ngày)',
        data: topStates.map(s => s.avg_lead),
        backgroundColor: topStates.map(s => s.avg_lead > 15 ? '#EF4444' : '#1E3A8A'),
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => `Lead Time: ${ctx.raw} ngày` } }
      },
      scales: {
        x: { ticks: { callback: (val) => `${val}d` }, grid: { color: '#F1F5F9' } },
        y: { grid: { display: false } }
      }
    }
  });

  // 4. Weight Class Impact
  const ctxWeight = document.getElementById('chart-p3-weight').getContext('2d');
  chartP3Weight = new Chart(ctxWeight, {
    type: 'bar',
    data: {
      labels: ['Hàng nhẹ (<1kg)', 'Trung bình (1-5kg)', 'Hàng nặng (>5kg)'],
      datasets: [
        {
          label: 'Phí Vận chuyển TB ($)',
          data: [43.47, 44.96, 46.61],
          backgroundColor: '#2563EB',
          borderRadius: 4,
          yAxisID: 'yFee'
        },
        {
          label: 'Lead Time TB (Ngày)',
          data: [11.87, 12.85, 14.19],
          backgroundColor: '#F59E0B',
          borderRadius: 4,
          yAxisID: 'yDays'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { family: 'Inter', size: 11 } } }
      },
      scales: {
        yFee: {
          position: 'left',
          ticks: { callback: (val) => `$${val}` },
          grid: { color: '#F1F5F9' },
          suggestedMax: 60
        },
        yDays: {
          position: 'right',
          ticks: { callback: (val) => `${val}d` },
          grid: { drawOnChartArea: false },
          suggestedMax: 20
        },
        x: { grid: { display: false } }
      }
    }
  });
}

// -------------------------------------------------------------
// TRANG 4: DEMAND FORECASTING & ACTION CENTER
// -------------------------------------------------------------
function initPage4Charts() {
  // 1. Forecast Line Chart
  const histMonths = DASHBOARD_DATA.monthly.slice(-8).map(m => m.order_month);
  const histRevs = DASHBOARD_DATA.monthly.slice(-8).map(m => m.revenue);

  const forecastMonths = ['2018-10', '2018-11 (BF)', '2018-12'];
  const forecastRevs = [1620000, 2450000, 2100000];

  const labels = [...histMonths, ...forecastMonths];
  const histData = [...histRevs, null, null, null];
  const fcData = [...Array(histMonths.length - 1).fill(null), histRevs[histRevs.length - 1], ...forecastRevs];

  const ctxForecast = document.getElementById('chart-p4-forecast').getContext('2d');
  chartP4Forecast = new Chart(ctxForecast, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Doanh thu Thực tế ($)',
          data: histData,
          borderColor: '#1E3A8A',
          backgroundColor: 'rgba(30, 58, 138, 0.1)',
          borderWidth: 2.5,
          tension: 0.3,
          pointRadius: 4
        },
        {
          label: 'Dự báo AI (Confidence 95%)',
          data: fcData,
          borderColor: '#10B981',
          backgroundColor: 'rgba(16, 185, 129, 0.15)',
          borderWidth: 2.5,
          borderDash: [6, 6],
          tension: 0.3,
          pointRadius: 5,
          fill: true
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { family: 'Inter', size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: $${ctx.raw ? ctx.raw.toLocaleString() : 'N/A'}`
          }
        }
      },
      scales: {
        y: { ticks: { callback: (val) => formatCurrency(val) }, grid: { color: '#F1F5F9' } },
        x: { ticks: { maxRotation: 45, minRotation: 45 } }
      }
    }
  });
}

function renderActionTable(cases = DASHBOARD_DATA.actions) {
  filteredActions = cases;
  const tbody = document.getElementById('tbody-action');
  tbody.innerHTML = '';

  cases.slice(0, 25).forEach(c => {
    const tr = document.createElement('tr');
    const badgeUrgency = c.urgency === 'HIGH' 
      ? '<span class="badge-urgency-high">CỰC KỲ KHẨN CẤP</span>' 
      : '<span class="badge-urgency-med">TRUNG BÌNH</span>';

    const statusBadge = c.status === 'canceled' 
      ? '<span class="badge-tag badge-alert">CANCELED</span>'
      : `<span class="badge-tag badge-warning">${c.status.toUpperCase()}</span>`;

    tr.innerHTML = `
      <td><code>${c.order_id}</code></td>
      <td><code>${c.customer_id}</code></td>
      <td><strong>${c.state}</strong></td>
      <td>${c.category}</td>
      <td>$${c.sales.toFixed(2)}</td>
      <td>$${c.shipping_charge.toFixed(2)}</td>
      <td><strong>${c.shipping_ratio}%</strong></td>
      <td>${c.lead_time}</td>
      <td>${statusBadge}</td>
      <td>${badgeUrgency}</td>
      <td><span style="color: #1E3A8A; font-weight: 600;">${c.action}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function filterActionTable() {
  const query = document.getElementById('action-search').value.toLowerCase();
  const urgency = document.getElementById('action-urgency-filter').value;

  const result = DASHBOARD_DATA.actions.filter(item => {
    const matchQuery = 
      item.order_id.toLowerCase().includes(query) ||
      item.customer_id.toLowerCase().includes(query) ||
      item.state.toLowerCase().includes(query) ||
      item.category.toLowerCase().includes(query);
    
    const matchUrgency = urgency === 'ALL' || item.urgency === urgency;
    return matchQuery && matchUrgency;
  });

  renderActionTable(result);
}

function exportActionTableToCSV() {
  let csv = 'Order ID,Customer ID,State,Category,Sales,Shipping Charge,Shipping Ratio %,Lead Time,Status,Urgency,Action\n';
  filteredActions.forEach(c => {
    csv += `"${c.order_id}","${c.customer_id}","${c.state}","${c.category}",${c.sales},${c.shipping_charge},${c.shipping_ratio},"${c.lead_time}","${c.status}","${c.urgency}","${c.action}"\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.setAttribute('download', 'Executive_Action_Tasks.csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function applyFilters() {
  const state = document.getElementById('slicer-state').value;
  const category = document.getElementById('slicer-category').value;

  // Visual notification of active filter
  console.log(`Filters applied: State=${state}, Category=${category}`);
}
