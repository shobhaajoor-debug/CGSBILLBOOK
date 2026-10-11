const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const projectRoot = path.resolve(__dirname, '..');
const defaultWorkbookPath = path.join(projectRoot, 'CGS-Payment Tracker-Latest.xlsx');
const dataDirectory = process.env.DATA_DIRECTORY || path.join(projectRoot, 'data');

function resolvedCellValue(cell) {
  const value = cell.value;
  if (value && typeof value === 'object') {
    if (Object.hasOwn(value, 'formula') || Object.hasOwn(value, 'sharedFormula')) {
      return resolvedValue(value.result);
    }
    if (Object.hasOwn(value, 'error')) return null;
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || '').join('');
    if (typeof value.text === 'string') return value.text;
  }
  return value ?? null;
}

function resolvedValue(value) {
  if (value && typeof value === 'object') {
    if (Object.hasOwn(value, 'error')) return null;
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || '').join('');
    if (typeof value.text === 'string') return value.text;
  }
  return value ?? null;
}

function sourceFieldValue(cell) {
  const rawValue = cell.value;
  if (rawValue && typeof rawValue === 'object' && (Object.hasOwn(rawValue, 'formula') || Object.hasOwn(rawValue, 'sharedFormula'))) {
    const result = rawValue.result;
    if (result && typeof result === 'object' && Object.hasOwn(result, 'error')) return result.error;
    return result === null || result === undefined ? '' : sourceFieldValue({ value: result });
  }
  if (rawValue instanceof Date) return dateText(rawValue);
  if (rawValue && typeof rawValue === 'object' && Object.hasOwn(rawValue, 'error')) return rawValue.error;
  if (rawValue && typeof rawValue === 'object' && Array.isArray(rawValue.richText)) {
    return rawValue.richText.map((part) => part.text || '').join('');
  }
  if (rawValue && typeof rawValue === 'object' && typeof rawValue.text === 'string') return rawValue.text;
  return rawValue ?? '';
}

function normalizedHeader(value) {
  return String(value ?? '').toLocaleLowerCase().replace(/[^a-z0-9]/g, '');
}

function findHeaderRow(sheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const columns = new Map();
    const labels = new Map();
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const label = String(resolvedCellValue(cell) ?? '').trim();
      columns.set(normalizedHeader(label), columnNumber);
      labels.set(columnNumber, label);
    });
    const company = columns.get('company');
    const invoice = [...columns].find(([name]) => name.startsWith('invoiceno'))?.[1];
    const total = columns.get('billedprice') || columns.get('billamount');
    if (company && invoice && total) return { rowNumber, columns, labels };
  }
  return null;
}

function column(columns, ...names) {
  for (const name of names) {
    if (columns.has(name)) return columns.get(name);
  }
  return null;
}

function amountAt(row, columnNumber, label, required = false) {
  if (!columnNumber) return 0;
  const raw = resolvedCellValue(row.getCell(columnNumber));
  if (raw === null || raw === '') {
    if (required) throw new Error(`${label} is missing at worksheet row ${row.number}.`);
    return 0;
  }
  const amount = Number(raw);
  if (!Number.isFinite(amount) || amount < 0) {
    if (required) throw new Error(`${label} is invalid at worksheet row ${row.number}.`);
    return 0;
  }
  return Math.round((amount + Number.EPSILON) * 100) / 100;
}

function dateText(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
  }
  const text = String(value ?? '').trim();
  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  const localMatch = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(text);
  const year = Number(isoMatch ? isoMatch[1] : localMatch?.[3]);
  const month = Number(isoMatch ? isoMatch[2] : localMatch?.[2]);
  const day = Number(isoMatch ? isoMatch[3] : localMatch?.[1]);
  if (!year || !month || !day) return '';
  const date = new Date(Date.UTC(year, month - 1, day));
  return Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
    ? ''
    : date.toISOString().slice(0, 10);
}

function sourceDate(row, columnNumber, label, required = false) {
  if (!columnNumber) {
    if (required) throw new Error(`${label} column is missing.`);
    return '';
  }
  const date = dateText(resolvedCellValue(row.getCell(columnNumber)));
  if (!date && required) throw new Error(`${label} is invalid at worksheet row ${row.number}.`);
  return date;
}

function sourceBillId(value) {
  const text = String(value ?? '').trim();
  if (/^\d+$/.test(text)) return text;
  if (/^CGS-\d+$/i.test(text)) return `CGS-${text.slice(4)}`;
  return '';
}

function sourceDueDays(value) {
  const match = /(\d+(?:\.\d+)?)\s*(?:days?)?/i.exec(String(value ?? '').trim());
  if (!match) return null;
  const days = Number(match[1]);
  return Number.isFinite(days) && days >= 0 ? Math.round(days) : null;
}

function buildBillsFromWorkbook(workbook, importedAt = new Date().toISOString()) {
  const bills = [];
  const seenBillIds = new Set();

  for (const sheet of workbook.worksheets) {
    if (/e[- ]?comer/i.test(sheet.name)) continue;
    const header = findHeaderRow(sheet);
    if (!header) continue;
    const { rowNumber: headerRowNumber, columns, labels } = header;
    const companyColumn = column(columns, 'company');
    const invoiceColumn = [...columns].find(([name]) => name.startsWith('invoiceno'))[1];
    const billColumn = column(columns, 'billedprice', 'billamount');
    const dateColumn = column(columns, 'date') || 2;
    const receivedColumn = column(columns, 'amountrecived', 'amountreceived');
    const purchaseColumn = column(columns, 'recivedprice', 'receivedprice');
    const dueDateColumn = column(columns, 'duedate', 'dueon');
    const dueDaysColumn = column(columns, 'duedays', 'noofduedays');
    const statusColumn = column(columns, 'status');
    const delayColumn = column(columns, 'noofdaysdealy', 'noofduedaysdealy', 'noofduedelaydays', 'noofdaysdelay');

    for (let rowNumber = headerRowNumber + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const shipToName = String(resolvedCellValue(row.getCell(companyColumn)) ?? '').replace(/\s+/g, ' ').trim();
      const rawInvoiceNumber = resolvedCellValue(row.getCell(invoiceColumn));
      const billId = sourceBillId(rawInvoiceNumber);
      if (!shipToName || !billId) continue;

      if (seenBillIds.has(billId)) {
        throw new Error(`Duplicate invoice number ${billId} at ${sheet.name}, row ${rowNumber}.`);
      }
      const totalAmount = amountAt(row, billColumn, 'Billed amount', true);
      const paidAmount = amountAt(row, receivedColumn, 'Amount received');
      const invoiceDate = sourceDate(row, dateColumn, 'Invoice date', true);
      const dueDate = sourceDate(row, dueDateColumn, 'Due date');
      const rawDueValue = dueDateColumn ? resolvedCellValue(row.getCell(dueDateColumn)) : null;
      const dueTerms = rawDueValue !== null && !dueDate ? String(rawDueValue).trim() : '';
      const dueDays = dueDate
        ? Math.round((Date.parse(`${dueDate}T00:00:00Z`) - Date.parse(`${invoiceDate}T00:00:00Z`)) / 86400000)
        : sourceDueDays(dueDaysColumn ? resolvedCellValue(row.getCell(dueDaysColumn)) : rawDueValue);
      const sourceFields = {};
      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        const label = labels.get(columnNumber);
        if (!label || /address/i.test(label)) return;
        sourceFields[label] = sourceFieldValue(cell);
      });

      const purchasePrice = amountAt(row, purchaseColumn, 'Purchase price', true);
      const profitAmount = Math.round((totalAmount - purchasePrice + Number.EPSILON) * 100) / 100;
      const profitRate = purchasePrice > 0 ? profitAmount / purchasePrice : 0;
      seenBillIds.add(billId);
      bills.push({
        billId,
        source: 'payment-tracker',
        sourceSheet: sheet.name,
        sourceInvoiceNo: String(rawInvoiceNumber).trim(),
        createdAt: importedAt,
        invoiceDate,
        totalAmount,
        ...(dueDate ? { dueDate } : {}),
        ...(dueDays !== null && dueDays >= 0 ? { dueDays } : {}),
        ...(dueTerms ? { sourceDueTerms: dueTerms } : {}),
        sourceFields,
        shipToName,
        vehicleNo: '',
        paidAmount,
        amountReceived: paidAmount,
        amountBalanced: Math.round((totalAmount - paidAmount + Number.EPSILON) * 100) / 100,
        purchasePrice,
        receivedPrice: purchasePrice,
        billedPrice: totalAmount,
        profitRate,
        profitAmount,
        sourceStatus: statusColumn ? String(resolvedCellValue(row.getCell(statusColumn)) ?? '').trim() : '',
        sourceDelayDays: delayColumn ? sourceFieldValue(row.getCell(delayColumn)) : '',
        paymentStatus: paidAmount >= totalAmount ? 'Paid' : 'Unpaid',
        paidOn: '',
        items: []
      });
    }
  }

  if (bills.length === 0) throw new Error('No bill records were found in the payment tracker workbook.');
  return bills;
}

async function writeAtomically(filePath, content) {
  const temporaryPath = `${filePath}.${process.pid}.tmp`;
  try {
    await fs.writeFile(temporaryPath, content, 'utf8');
    await fs.rename(temporaryPath, filePath);
  } catch (error) {
    await fs.unlink(temporaryPath).catch(() => {});
    throw error;
  }
}

async function importPaymentTracker(workbookPath = defaultWorkbookPath, { replace = false } = {}) {
  const billsPath = path.join(dataDirectory, 'bills.json');
  const paymentsPath = path.join(dataDirectory, 'payments.json');
  if (!replace) {
    for (const target of [billsPath, paymentsPath]) {
      try {
        await fs.access(target);
        throw new Error(`Refusing to replace existing ${path.basename(target)} without --replace.`);
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
    }
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const bills = buildBillsFromWorkbook(workbook);
  await fs.mkdir(dataDirectory, { recursive: true });
  await writeAtomically(billsPath, `${JSON.stringify(bills, null, 2)}\n`);
  await writeAtomically(paymentsPath, '[]\n');
  return bills;
}

if (require.main === module) {
  const replace = process.argv.includes('--replace');
  const workbookPath = process.argv.slice(2).find((argument) => argument !== '--replace') || defaultWorkbookPath;
  importPaymentTracker(workbookPath, { replace })
    .then((bills) => {
      const accounts = new Map();
      for (const bill of bills) {
        const key = bill.shipToName.toLocaleLowerCase();
        const account = accounts.get(key) || { total: 0, received: 0 };
        account.total += bill.totalAmount;
        account.received += bill.paidAmount;
        accounts.set(key, account);
      }
      const totalAmount = [...accounts.values()].reduce((sum, account) => sum + account.total, 0);
      const paidAmount = [...accounts.values()].reduce((sum, account) => sum + Math.min(account.total, account.received), 0);
      console.log(`Imported ${bills.length} bills from ${path.basename(workbookPath)}.`);
      console.log(`Billed: ₹${totalAmount.toFixed(2)} | Received: ₹${paidAmount.toFixed(2)} | Outstanding: ₹${(totalAmount - paidAmount).toFixed(2)}`);
    })
    .catch((error) => {
      console.error(`Payment tracker import failed: ${error.message}`);
      process.exitCode = 1;
    });
}

module.exports = { buildBillsFromWorkbook, importPaymentTracker };
