const { summarizeTaxByHsn } = require('./tax-summary');

function ewayDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
  return match ? `${match[3]}/${match[2]}/${match[1]}` : null;
}

function buildEwayBillDraft(bill) {
  const items = Array.isArray(bill.items) ? bill.items : [];
  const taxSummary = summarizeTaxByHsn(items);
  const taxableValue = taxSummary.reduce((total, row) => total + row.taxableValue, 0);
  const totalTax = taxSummary.reduce((total, row) => total + row.taxAmount, 0);
  const missingFields = [
    'fromGstin',
    'fromPlace',
    'fromPincode',
    'fromStateCode',
    'toGstin',
    'toPlace',
    'toPincode',
    'toStateCode',
    'transactionType',
    'transMode',
    'transDistance',
    'vehicleType',
    'itemList[].taxRates',
    'itemList[].qtyUnitMapping',
    'otherValue (discount and round-off mapping)'
  ];

  return {
    readyToSubmit: false,
    missingFields,
    calculatedTaxValue: Number(totalTax.toFixed(2)),
    payload: {
      supplyType: 'O',
      subSupplyType: '1',
      subSupplyDesc: '',
      docType: 'INV',
      docNo: bill.sourceInvoiceNo || bill.billId,
      docDate: ewayDate(bill.invoiceDate),
      fromGstin: null,
      fromTrdName: bill.sellerName || null,
      fromAddr1: bill.sellerAddress || null,
      fromAddr2: null,
      fromPlace: null,
      fromPincode: null,
      actFromStateCode: null,
      fromStateCode: null,
      toGstin: null,
      toTrdName: bill.shipToName || null,
      toAddr1: bill.shipToAddress || null,
      toAddr2: null,
      toPlace: null,
      toPincode: null,
      actToStateCode: null,
      toStateCode: null,
      transactionType: null,
      dispatchFromGSTIN: null,
      dispatchFromTradeName: null,
      shipToGSTIN: null,
      shipToTradeName: null,
      totalValue: Number(taxableValue.toFixed(2)),
      cgstValue: null,
      sgstValue: null,
      igstValue: null,
      cessValue: null,
      cessNonAdvolValue: null,
      otherValue: null,
      totInvValue: Number((Number(bill.totalAmount) || 0).toFixed(2)),
      transporterId: null,
      transporterName: null,
      transDocNo: null,
      transDocDate: null,
      transMode: null,
      transDistance: null,
      vehicleNo: bill.vehicleNo || null,
      vehicleType: null,
      itemList: items.filter((item) => !item.isDiscount).map((item) => ({
        productName: item.description || null,
        productDesc: item.description || null,
        hsnCode: item.hsn || null,
        quantity: Number(item.quantity),
        qtyUnit: item.unit || null,
        cgstRate: null,
        sgstRate: null,
        igstRate: null,
        cessRate: null,
        cessNonAdvol: null,
        taxableAmount: Number((Number(item.amount) || 0).toFixed(2))
      }))
    }
  };
}

module.exports = { buildEwayBillDraft };
