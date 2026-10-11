const overviewError = document.querySelector('#overview-error');
const recentBody = document.querySelector('#overview-recent');
const vendorList = document.querySelector('#overview-vendors-list');

function formatAmount(value) {
  return `₹${new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value) || 0)}`;
}

function formatDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || '').slice(0, 10));
  if (!match) return '—';
  return new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
    .format(new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function paidAmount(bill) {
  return Number(bill.paidAmount ?? (bill.paymentStatus === 'Paid' ? bill.totalAmount : 0)) || 0;
}

function paymentState(bill) {
  const total = Number(bill.totalAmount) || 0;
  const paid = paidAmount(bill);
  if (paid >= total) return 'Paid';
  return paid > 0 ? 'Partial' : 'Unpaid';
}

function statusBadge(status) {
  const badge = document.createElement('span');
  badge.className = `status-badge ${status.toLocaleLowerCase()}`;
  badge.textContent = status === 'Partial' ? 'Partially paid' : status;
  return badge;
}

function cell(value, className = '') {
  const tableCell = document.createElement('td');
  tableCell.textContent = value;
  if (className) tableCell.className = className;
  return tableCell;
}

function renderMonthlyChart(bills) {
  const months = new Map();
  for (const bill of bills) {
    const match = /^(\d{4})-(\d{2})/.exec(String(bill.invoiceDate || ''));
    if (!match) continue;
    const key = `${match[1]}-${match[2]}`;
    months.set(key, (months.get(key) || 0) + (Number(bill.totalAmount) || 0));
  }
  const entries = [...months].sort(([left], [right]) => left.localeCompare(right)).slice(-8);
  const highest = Math.max(...entries.map(([, amount]) => amount), 1);
  const chart = document.querySelector('#monthly-chart');
  chart.replaceChildren();
  if (!entries.length) {
    chart.textContent = 'Monthly totals will appear here after bills are recorded.';
    chart.classList.add('chart-empty');
    return;
  }
  chart.classList.remove('chart-empty');
  for (const [month, amount] of entries) {
    const item = document.createElement('div');
    item.className = 'month-bar-item';
    const label = document.createElement('div');
    label.className = 'month-bar-total';
    label.textContent = formatAmount(amount);
    const track = document.createElement('div');
    track.className = 'month-bar-track';
    const bar = document.createElement('span');
    bar.className = 'month-bar';
    bar.style.height = `${Math.max(5, amount / highest * 100)}%`;
    bar.title = `${month}: ${formatAmount(amount)}`;
    track.append(bar);
    const monthLabel = document.createElement('span');
    monthLabel.className = 'month-bar-label';
    monthLabel.textContent = new Intl.DateTimeFormat('en', { month: 'short', year: '2-digit' })
      .format(new Date(Number(month.slice(0, 4)), Number(month.slice(5, 7)) - 1, 1));
    item.append(label, track, monthLabel);
    chart.append(item);
  }
  document.querySelector('#overview-period').textContent = `${entries.length} recent months`;
}

function renderVendors(vendors) {
  vendorList.replaceChildren();
  const withBalances = vendors.filter((vendor) => vendor.unpaidAmount > 0)
    .sort((left, right) => right.unpaidAmount - left.unpaidAmount)
    .slice(0, 5);
  if (!withBalances.length) {
    const empty = document.createElement('p');
    empty.className = 'overview-list-empty';
    empty.textContent = 'No customer balances are outstanding.';
    vendorList.append(empty);
    return;
  }
  for (const vendor of withBalances) {
    const row = document.createElement('a');
    row.className = 'overview-vendor-row';
    row.href = '/explore.html';
    const details = document.createElement('span');
    details.className = 'overview-vendor-copy';
    const name = document.createElement('strong');
    name.textContent = vendor.vendorName;
    const count = document.createElement('small');
    count.textContent = `${vendor.billCount} ${vendor.billCount === 1 ? 'bill' : 'bills'}`;
    details.append(name, count);
    const amount = document.createElement('strong');
    amount.className = 'overview-vendor-amount';
    amount.textContent = formatAmount(vendor.unpaidAmount);
    row.append(details, amount);
    vendorList.append(row);
  }
}

function renderRecent(bills) {
  recentBody.replaceChildren();
  const recent = bills.slice().sort((left, right) => (
    String(right.invoiceDate || '').localeCompare(String(left.invoiceDate || ''))
    || String(right.billId).localeCompare(String(left.billId))
  )).slice(0, 7);
  document.querySelector('#overview-empty').hidden = recent.length !== 0;
  for (const bill of recent) {
    const row = document.createElement('tr');
    row.append(
      cell(bill.billId, 'bill-number-cell'),
      cell(bill.shipToName || '—'),
      cell(formatDate(bill.invoiceDate)),
      cell(formatAmount(bill.totalAmount), 'amount-column'),
      cell(formatAmount(paidAmount(bill)), 'amount-column'),
      cell(formatAmount((Number(bill.totalAmount) || 0) - paidAmount(bill)), 'amount-column')
    );
    const statusCell = document.createElement('td');
    statusCell.append(statusBadge(paymentState(bill)));
    row.append(statusCell);
    recentBody.append(row);
  }
}

async function loadOverview() {
  try {
    const [billsResponse, exploreResponse] = await Promise.all([
      fetch('/api/bills'),
      fetch('/api/explore')
    ]);
    const [bills, explore] = await Promise.all([billsResponse.json(), exploreResponse.json()]);
    if (!billsResponse.ok) throw new Error(bills.error || 'Unable to load bill history.');
    if (!exploreResponse.ok) throw new Error(explore.error || 'Unable to load account balances.');
    if (!Array.isArray(bills) || !explore.summary || !Array.isArray(explore.vendors)) {
      throw new Error('The overview response is incomplete. Restart the bill book server and try again.');
    }
    const summary = explore.summary;
    document.querySelector('#overview-bills').textContent = new Intl.NumberFormat('en-IN').format(bills.length);
    document.querySelector('#overview-vendors').textContent = `${summary.vendorCount} customer accounts`;
    document.querySelector('#overview-total').textContent = formatAmount(summary.totalAmount);
    document.querySelector('#overview-paid').textContent = formatAmount(summary.paidAmount);
    document.querySelector('#overview-unpaid').textContent = formatAmount(summary.unpaidAmount);
    const profitTotal = bills.reduce((sum, bill) => sum + (Number(bill.profitAmount) || 0), 0);
    document.querySelector('#overview-profit').textContent = formatAmount(profitTotal);
    const purchaseTotal = bills.reduce((sum, bill) => sum + (Number(bill.purchasePrice) || 0), 0);
    const profitRate = purchaseTotal > 0 ? profitTotal / purchaseTotal * 100 : 0;
    document.querySelector('#overview-profit-rate').textContent = `${profitRate.toFixed(2)}% on purchase price`;
    const collectionRate = summary.totalAmount > 0 ? summary.paidAmount / summary.totalAmount * 100 : 0;
    document.querySelector('#overview-collection-rate').textContent = `${collectionRate.toFixed(1)}% collected`;
    renderMonthlyChart(bills);
    renderVendors(explore.vendors);
    renderRecent(bills);
    document.querySelector('#tracker-note').hidden = !bills.some((bill) => bill.source === 'payment-tracker');
  } catch (error) {
    overviewError.textContent = error.message || 'Unable to load the business overview.';
    overviewError.hidden = false;
  }
}

loadOverview();
