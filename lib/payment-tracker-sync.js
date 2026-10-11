const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');
const { buildBillsFromWorkbook } = require('../scripts/import-payment-tracker');

function createPaymentTrackerSync({ workbookPath, ledgerPath }) {
  let lastLoadedSignature = null;
  const signaturePath = `${ledgerPath}.tracker-sync.json`;

  return async function syncPaymentTracker() {
    if (!workbookPath) return false;

    const stats = await fs.stat(workbookPath);
    const signature = `${stats.size}:${stats.mtimeMs}`;
    if (signature === lastLoadedSignature) return false;

    let previousBills = [];
    let ledgerExists = true;
    try {
      previousBills = JSON.parse(await fs.readFile(ledgerPath, 'utf8'));
      if (!Array.isArray(previousBills)) throw new Error(`${path.basename(ledgerPath)} must contain a JSON array.`);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
      ledgerExists = false;
    }

    if (ledgerExists) {
      try {
        const previousSignature = JSON.parse(await fs.readFile(signaturePath, 'utf8'));
        if (previousSignature === signature) {
          lastLoadedSignature = signature;
          return false;
        }
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const freshBills = buildBillsFromWorkbook(workbook);
    const previousTrackerBills = new Map(
      previousBills.filter((bill) => bill.source === 'payment-tracker').map((bill) => [bill.billId, bill])
    );
    const retainedBills = previousBills.filter((bill) => bill.source !== 'payment-tracker');
    const retainedBillIds = new Set(retainedBills.map((bill) => bill.billId));
    const refreshedBills = freshBills.map((bill) => {
      if (retainedBillIds.has(bill.billId)) {
        throw new Error(`Payment tracker invoice ${bill.billId} conflicts with an app-created bill.`);
      }
      const previousBill = previousTrackerBills.get(bill.billId);
      return previousBill ? { ...bill, createdAt: previousBill.createdAt || bill.createdAt } : bill;
    });
    const records = [...refreshedBills, ...retainedBills];

    if (JSON.stringify(records) !== JSON.stringify(previousBills)) {
      const temporaryPath = `${ledgerPath}.${process.pid}.tmp`;
      try {
        await fs.mkdir(path.dirname(ledgerPath), { recursive: true });
        await fs.writeFile(temporaryPath, `${JSON.stringify(records, null, 2)}\n`, 'utf8');
        await fs.rename(temporaryPath, ledgerPath);
      } catch (error) {
        await fs.unlink(temporaryPath).catch(() => {});
        throw error;
      }
    }

    const signatureTemporaryPath = `${signaturePath}.${process.pid}.tmp`;
    try {
      await fs.mkdir(path.dirname(signaturePath), { recursive: true });
      await fs.writeFile(signatureTemporaryPath, JSON.stringify(signature), 'utf8');
      await fs.rename(signatureTemporaryPath, signaturePath);
    } catch (error) {
      await fs.unlink(signatureTemporaryPath).catch(() => {});
      throw error;
    }

    lastLoadedSignature = signature;
    return true;
  };
}

module.exports = { createPaymentTrackerSync };
