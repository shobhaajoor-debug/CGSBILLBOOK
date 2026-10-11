const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const ExcelJS = require('exceljs');
const { buildBillsFromWorkbook } = require('../scripts/import-payment-tracker');
const { createPaymentTrackerSync } = require('../lib/payment-tracker-sync');

function addTrackerSheet(workbook, name, billRows) {
  const sheet = workbook.addWorksheet(name);
  sheet.addRow(['CGS GLOBAL ENTERPRISES']);
  sheet.addRow([
    'Sl.NO', 'Date', 'Company', 'Invoice NO.', 'Recived price', 'Billed price',
    '% age profit', 'Amount of profit', 'Amount Recived', 'Amount Balanced', 'Due date', 'Ship to Address'
  ]);
  sheet.getCell('A2').value = { richText: [{ text: 'Sl.' }, { text: 'NO' }] };
  for (const values of billRows) sheet.addRow(values);
  return sheet;
}

test('imports actual tracker values and ignores formatted headers and non-bill sheets', () => {
  const workbook = new ExcelJS.Workbook();
  addTrackerSheet(workbook, 'Cashflow-Dec', [
    [1, new Date('2025-12-02T00:00:00.000Z'), 'Customer One', 'CGS-01', 100, 118, 0.18, 18, { formula: 'F3', result: 118 }, 0, new Date('2025-12-12T00:00:00.000Z'), 'Address not retained'],
    [2, '03-12-2025', 'Customer Two', 2, 200, 236, 0.18, 36, 100, 136, '25 days', 'Address not retained']
  ]);
  addTrackerSheet(workbook, 'E-COMER', [
    [1, new Date('2025-12-02T00:00:00.000Z'), 'Non Bill Row', 'CGS-99', 10, 20, 0.5, 10, 20, 0, '2025-12-12']
  ]);

  const bills = buildBillsFromWorkbook(workbook, '2026-10-06T00:00:00.000Z');

  assert.equal(bills.length, 2);
  assert.deepEqual(bills.map((bill) => bill.billId), ['CGS-01', '2']);
  assert.equal(bills[1].sourceInvoiceNo, '2');
  assert.equal(bills[0].invoiceDate, '2025-12-02');
  assert.equal(bills[0].dueDate, '2025-12-12');
  assert.equal(bills[0].totalAmount, 118);
  assert.equal(bills[0].paidAmount, 118);
  assert.equal(bills[0].paymentStatus, 'Paid');
  assert.equal(bills[0].profitRate, 0.18);
  assert.equal(bills[0].sourceFields['Amount Balanced'], 0);
  assert.equal(bills[0].sourceFields['Sl.NO'], 1);
  assert.equal(Object.hasOwn(bills[0].sourceFields, '[object Object]'), false);
  assert.equal('sellerAddress' in bills[0], false);
  assert.equal('shipToAddress' in bills[0], false);
  assert.equal(Object.keys(bills[0].sourceFields).some((label) => /address/i.test(label)), false);
  assert.equal(bills[1].invoiceDate, '2025-12-03');
  assert.equal(bills[1].paidAmount, 100);
  assert.equal(bills[1].paymentStatus, 'Unpaid');
  assert.equal(bills[1].dueDays, 25);
  assert.deepEqual(bills[1].items, []);
});

test('rejects duplicate invoice numbers instead of silently dropping a bill', () => {
  const workbook = new ExcelJS.Workbook();
  addTrackerSheet(workbook, 'Cashflow-Dec', [
    [1, new Date('2025-12-02T00:00:00.000Z'), 'Customer One', 'CGS-01', 100, 118, 0.18, 18, 0, 118, '', ''],
    [2, new Date('2025-12-03T00:00:00.000Z'), 'Customer Two', 'CGS-01', 100, 118, 0.18, 18, 0, 118, '', '']
  ]);

  assert.throws(() => buildBillsFromWorkbook(workbook), /Duplicate invoice number CGS-01/);
});

test('re-syncs edited tracker data while retaining bills created in the app', async () => {
  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'cgs-payment-tracker-sync-'));
  try {
    const workbookPath = path.join(temporaryDirectory, 'tracker.xlsx');
    const ledgerPath = path.join(temporaryDirectory, 'bills.json');
    const workbook = new ExcelJS.Workbook();
    const sheet = addTrackerSheet(workbook, 'Cashflow-Dec', [
      [1, new Date('2025-12-02T00:00:00.000Z'), 'Customer One', 'CGS-01', 100, 118, 0.18, 18, 0, 118, '', '']
    ]);
    await workbook.xlsx.writeFile(workbookPath);

    const sync = createPaymentTrackerSync({ workbookPath, ledgerPath });
    assert.equal(await sync(), true);
    const firstRecords = JSON.parse(await fs.readFile(ledgerPath, 'utf8'));
    firstRecords.push({ billId: 'CGS-0002', source: 'app', shipToName: 'Manual Customer' });
    await fs.writeFile(ledgerPath, `${JSON.stringify(firstRecords, null, 2)}\n`);

    sheet.getCell('F3').value = 125;
    sheet.addRow([2, new Date('2025-12-03T00:00:00.000Z'), 'Customer Two', 'CGS-02', 100, 130, 0.3, 30, 0, 130, '', '']);
    await workbook.xlsx.writeFile(workbookPath);
    assert.equal(await sync(), true);

    const refreshedRecords = JSON.parse(await fs.readFile(ledgerPath, 'utf8'));
    assert.deepEqual(refreshedRecords.map((bill) => bill.billId), ['CGS-01', 'CGS-02', 'CGS-0002']);
    assert.equal(refreshedRecords[0].totalAmount, 125);
    assert.equal(refreshedRecords[2].shipToName, 'Manual Customer');
    assert.equal(await sync(), false);

    refreshedRecords[0].paidAmount = 12;
    await fs.writeFile(ledgerPath, `${JSON.stringify(refreshedRecords, null, 2)}\n`);
    const syncAfterRestart = createPaymentTrackerSync({ workbookPath, ledgerPath });
    assert.equal(await syncAfterRestart(), false);
    const recordsAfterRestart = JSON.parse(await fs.readFile(ledgerPath, 'utf8'));
    assert.equal(recordsAfterRestart[0].paidAmount, 12);
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
});
