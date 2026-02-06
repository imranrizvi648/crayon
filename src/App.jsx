import React, { useState, useMemo } from 'react';
import { Calculator, Plus, Trash2, Save, Send, FileText, DollarSign, Users, ChevronDown, ChevronUp, CheckCircle, XCircle, Clock, Edit3, Download, FilePlus, Clipboard, FileSpreadsheet } from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const PRODUCT_CATEGORIES = {
  ENTERPRISE_ONLINE: { label: 'Enterprise Online Products', short: 'Enterprise Online' },
  ADDITIONAL: { label: 'Additional Products', short: 'Additional' },
  ADDITIONAL_ON_PREMISE: { label: 'Additional Products - On Premise', short: 'On Premise' }
};

// Helper to parse Excel pasted data
const parseNumber = (str) => {
  if (!str || str === '-' || str.trim() === '') return 0;
  const cleaned = str.toString().replace(/[$,\s%]/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
};

const parsePercent = (str) => {
  const num = parseNumber(str);
  if (str && str.toString().includes('%')) {
    return num / 100;
  }
  if (num > 1) return num / 100;
  return num;
};

const cleanItemName = (str) => {
  if (!str) return '';
  return str.replace(/^["']|["']$/g, '').replace(/\s+/g, ' ').trim();
};

const detectCategory = (line) => {
  const lower = line.toLowerCase();
  if (lower.includes('on premise') || lower.includes('on-premise')) return 'ADDITIONAL_ON_PREMISE';
  if (lower.includes('additional products')) return 'ADDITIONAL';
  if (lower.includes('enterprise online')) return 'ENTERPRISE_ONLINE';
  return null;
};

const looksLikePartNumber = (str) => {
  if (!str) return false;
  const trimmed = str.trim();
  return /^[A-Z0-9]{2,}-[A-Z0-9]+$/i.test(trimmed);
};

const joinMultiLineQuotedFields = (text) => {
  const lines = text.split(/\r?\n/);
  const result = [];
  let current = '';
  
  const looksLikeCategoryRow = (str) => {
    const lower = str.toLowerCase();
    return lower.includes('additional products') || 
           lower.includes('enterprise online') || 
           lower.includes('on premise');
  };
  
  for (const line of lines) {
    if (!line.trim()) continue;
    
    const cols = line.split('\t');
    const firstCol = cols[0]?.trim() || '';
    
    const isNewRow = looksLikePartNumber(firstCol) || 
                     (firstCol === '' && looksLikeCategoryRow(line));
    
    if (isNewRow) {
      if (current) {
        result.push(current);
      }
      current = line;
    } else if (current) {
      current = current + ' ' + line;
    } else {
      current = line;
    }
  }
  
  if (current) {
    result.push(current);
  }
  
  return result;
};

const parseExcelRow = (row, currentCategory) => {
  const cols = row.split('\t');
  
  const fullRowText = cols.join(' ');
  const detectedCat = detectCategory(fullRowText);
  if (detectedCat && (!cols[0] || cols[0].trim() === '' || detectCategory(cols[0]) || detectCategory(cols[1]))) {
    return { isCategory: true, category: detectedCat };
  }
  
  const partNum = cols[0]?.trim() || '';
  if (!partNum || !looksLikePartNumber(partNum)) return null;
  
  return {
    isCategory: false,
    data: {
      partNumber: partNum,
      itemName: cleanItemName(cols[1] || ''),
      unitNetUsd: parseNumber(cols[2]),
      unitErpUsd: parseNumber(cols[3]),
      msDiscountPct: parsePercent(cols[5]),
      crayonMarkupPct: parsePercent(cols[6]),
      unitType: parseInt(parseNumber(cols[7])) || 12,
      quantity: parseInt(parseNumber(cols[12])) || 0,
      rebatePct: parsePercent(cols[16]),
      swoGpPct: cols[19] !== undefined && cols[19] !== '' ? parsePercent(cols[19]) : 0.5, // Default 50% for Africa GP split
      category: currentCategory
    }
  };
};

const getEmptyHeader = () => ({
  customerName: 'Etihad Water and Electricity (EtihadWE)', opportunityId: '', region: 'ME', currencyCode: 'AED', exchangeRate: 3.6725, accountManager: 'Mohammed Areff', agreementType: 'Enterprise Enrollment', vatRate: 0.05,
  salesLocation: 'UAE', newOrRenewal: 'Renewal', agreementLevelSystem: 'D', agreementLevelServer: 'D', agreementLevelApplication: 'D',
  bidBondPct: 0, bankChargesPct: 0.03, performanceBondPct: 0, performanceBankChargesPct: 0.01,
  tenderCost: 0,
  otherLspRebateY1: 0, otherLspRebateY2: 0, otherLspRebateY3: 0,
  cifM365E5: 3321665.88, cifM365E3: 0, cifAzure: 0, cifDynamics365: 0,
  partnerName: '',
  dealType: 'Normal',
  erpCustomerId: '',
  customerSegment: '',
  businessArea: '',
  producerName: '',
  accountManagerId: ''
});

const getEmptyLineItems = () => [
  { id: 1, partNumber: '', itemName: '', category: 'ENTERPRISE_ONLINE', unitNetUsd: 0, unitErpUsd: 0, msDiscountPct: 0, crayonMarkupPct: 0, unitType: 12, quantity: 0, rebatePct: 0, swoGpPct: 0.5 },
];

const getEmptyDiscounts = () => ({ year1: 18362.50, year2: 14690.00, year3: 11017.50 });

const getSampleHeader = () => ({
  customerName: 'Etihad Water and Electricity (EtihadWE)', opportunityId: 'OPP-2025-ME-001', region: 'ME', currencyCode: 'AED', exchangeRate: 3.6725, accountManager: 'Mohammed Areff', agreementType: 'Enterprise Enrollment', vatRate: 0.05,
  salesLocation: 'UAE', newOrRenewal: 'Renewal', agreementLevelSystem: 'D', agreementLevelServer: 'D', agreementLevelApplication: 'D',
  bidBondPct: 0, bankChargesPct: 0.03, performanceBondPct: 0, performanceBankChargesPct: 0.01,
  tenderCost: 0,
  otherLspRebateY1: 0, otherLspRebateY2: 0, otherLspRebateY3: 0,
  cifM365E5: 3321665.88, cifM365E3: 0, cifAzure: 0, cifDynamics365: 0,
  partnerName: '',
  dealType: 'Normal',
  erpCustomerId: '',
  customerSegment: '',
  businessArea: '',
  producerName: '',
  accountManagerId: ''
});

const getSampleLineItems = () => [
  { id: 1, partNumber: 'AAA-28605', itemName: 'M365 E5 Original Existing Customer Sub Per User', category: 'ENTERPRISE_ONLINE', unitNetUsd: 50.54, unitErpUsd: 52.2, msDiscountPct: 0.20, crayonMarkupPct: 0.015, unitType: 12, quantity: 33, rebatePct: 0.0325, swoGpPct: 0.5 },
  { id: 2, partNumber: 'AAD-33177', itemName: 'M365 E5 Unified FSA Renewal Sub Per User', category: 'ENTERPRISE_ONLINE', unitNetUsd: 45.86, unitErpUsd: 47.3, msDiscountPct: 0.20, crayonMarkupPct: 0.015, unitType: 12, quantity: 2018, rebatePct: 0.0325, swoGpPct: 0.5 },
  { id: 3, partNumber: '', itemName: '', category: 'ENTERPRISE_ONLINE', unitNetUsd: 0, unitErpUsd: 0, msDiscountPct: 0, crayonMarkupPct: 0, unitType: 12, quantity: 0, rebatePct: 0, swoGpPct: 0.5 },
];

const getSampleDiscounts = () => ({ year1: 18362.50, year2: 14690.00, year3: 11017.50 });

// FIXED: Match Excel calculation methodology exactly
// Excel uses ROUNDED values at each intermediate step (cell-by-cell calculation)
// Each cell formula rounds its result, and subsequent cells use that rounded value
// 
// IMPORTANT: Normal and Ramped deals use DIFFERENT EUP formulas!
// Normal: EUP = ROUND(MSDiscNet * (1 + CrayonMarkup%), 2) - MULTIPLICATION
// Ramped: EUP = ROUND(MSDiscNet / (1 - CrayonMarkup%), 2) - DIVISION
//
// CRITICAL: Excel rounds EACH intermediate value, not just the final result
const calculateLineItem = (item, exchangeRate, dealType = 'Normal') => {
  // Step 1: MS Disc Net = ROUND(UnitNet * (1-MSDisc%) * ExRate, 2)
  // This is a cell value - Excel rounds and stores this
  const discountedNet = Math.round(item.unitNetUsd * (1 - item.msDiscountPct) * exchangeRate * 100) / 100;
  
  // Step 2: MS Disc ERP = ROUND(UnitERP * (1-MSDisc%) * ExRate, 2)
  // This is a cell value - Excel rounds and stores this
  const discountedErp = Math.round(item.unitErpUsd * (1 - item.msDiscountPct) * exchangeRate * 100) / 100;
  
  // Step 3: Total Net = ROUND(MSDiscNet * UnitType * Qty, 2)
  // Uses the ROUNDED discountedNet value
  const totalNet = Math.round(discountedNet * item.unitType * item.quantity * 100) / 100;
  
  // Step 4: Total ERP = ROUND(MSDiscERP * UnitType * Qty, 2)
  // Uses the ROUNDED discountedErp value
  const totalErp = Math.round(discountedErp * item.unitType * item.quantity * 100) / 100;
  
  // Step 5: Default Markup calculation
  const defaultMarkup = item.unitErpUsd > 0 ? (item.unitErpUsd - item.unitNetUsd) / item.unitErpUsd : 0;
  
  // Step 6: EUP - CRITICAL BUSINESS RULE:
  // If Default Markup is ~0% (Unit Net ≈ Unit ERP), use MS Disc ERP (ignore Crayon Markup)
  // Otherwise, apply the appropriate formula based on deal type:
  // - Normal: EUP = MSDiscNet * (1 + CrayonMarkup%) - MULTIPLICATION
  // - Ramped: EUP = MSDiscNet / (1 - CrayonMarkup%) - DIVISION
  let eupUnit;
  const hasDefaultMarkup = Math.abs(item.unitErpUsd - item.unitNetUsd) > 0.001;
  
  if (hasDefaultMarkup && item.crayonMarkupPct > 0) {
    if (dealType === 'Ramped') {
      // Ramped: EUP = MSDiscNet / (1 - CrayonMarkup%)
      eupUnit = Math.round(discountedNet / (1 - item.crayonMarkupPct) * 100) / 100;
    } else {
      // Normal: EUP = MSDiscNet * (1 + CrayonMarkup%)
      eupUnit = Math.round(discountedNet * (1 + item.crayonMarkupPct) * 100) / 100;
    }
  } else {
    // Products with no default markup: use ROUNDED MS Disc ERP
    eupUnit = discountedErp;
  }
  
  // Step 7: Total EUP = ROUND(EUP * Qty * UnitType, 2)
  // Uses the ROUNDED eupUnit value
  const totalEup = Math.round(eupUnit * item.quantity * item.unitType * 100) / 100;
  
  // Step 8: Rebate = ROUND(Total Net * Rebate %, 2)
  const rebateAmount = Math.round(totalNet * item.rebatePct * 100) / 100;
  
  // Step 9: GP Split calculations (for Africa region)
  const gp = Math.round((totalEup - totalNet) * 100) / 100;
  const swoGpPct = item.swoGpPct !== undefined && item.swoGpPct !== null ? item.swoGpPct : 0.5;
  const swoGp = Math.round(gp * swoGpPct * 100) / 100;
  const partnerGp = Math.round((gp - swoGp) * 100) / 100;
  
  // Calculated markup for display
  const calculatedMarkup = discountedNet > 0 ? (eupUnit - discountedNet) / discountedNet : 0;
  
  return { 
    ...item, 
    discountedNet,
    discountedErp,
    totalNet,
    totalErp,
    eupUnit,
    totalEup,
    rebateAmount,
    defaultMarkup, 
    calculatedMarkup,
    gp,
    swoGp,
    partnerGp
  };
};

const fmt = (v, c = 'AED') => {
  if (v == null || isNaN(v)) return '—';
  try {
    // Only use currency formatting if we have a valid 3-letter code
    if (c && c.length === 3) {
      return new Intl.NumberFormat('en-US', { style: 'currency', currency: c, minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
    }
  } catch (e) {
    // Fall through to default formatting
  }
  // Fallback: format as number with currency code prefix
  return `${c || ''} ${new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v)}`.trim();
};
const fmtNum = (v, d = 2) => v == null || isNaN(v) ? '—' : v.toFixed(d);
const fmtPct = (v, d = 2) => `${(v * 100).toFixed(d)}%`;

const generateSheetId = () => `CS-2025-${String(Math.floor(Math.random() * 999) + 1).padStart(3, '0')}`;

const StatusBadge = ({ status }) => {
  const cfg = { draft: { c: 'bg-gray-100 text-gray-700', i: Edit3, l: 'Draft' }, submitted: { c: 'bg-blue-100 text-blue-700', i: Send, l: 'Submitted' }, approved: { c: 'bg-green-100 text-green-700', i: CheckCircle, l: 'Approved' }, rejected: { c: 'bg-red-100 text-red-700', i: XCircle, l: 'Rejected' } }[status] || { c: 'bg-gray-100', i: Clock, l: status };
  const I = cfg.i;
  return <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold ${cfg.c}`}><I size={12} /> {cfg.l}</span>;
};

const Section = ({ title, children, open: defaultOpen = true, badge }) => {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="border border-gray-200 rounded-xl bg-white shadow-sm overflow-hidden">
      <button onClick={() => setOpen(!open)} className="w-full flex items-center justify-between p-4 bg-gray-50 border-b hover:bg-gray-100">
        <div className="flex items-center gap-3"><span className="font-semibold text-gray-800">{title}</span>{badge}</div>
        {open ? <ChevronUp size={20} className="text-gray-500" /> : <ChevronDown size={20} className="text-gray-500" />}
      </button>
      {open && <div className="p-4">{children}</div>}
    </div>
  );
};

const MetricCard = ({ label, value, sub, color = 'gray', big }) => (
  <div className={`p-4 rounded-xl border ${color === 'green' ? 'border-green-200 bg-green-50' : color === 'blue' ? 'border-blue-200 bg-blue-50' : 'border-gray-200'}`}>
    <p className="text-xs font-medium text-gray-500 uppercase">{label}</p>
    <p className={`font-bold ${big ? 'text-2xl' : 'text-xl'} ${color === 'green' ? 'text-green-700' : color === 'blue' ? 'text-blue-700' : 'text-gray-900'}`}>{value}</p>
    {sub && <p className="text-xs text-gray-500 mt-1">{sub}</p>}
  </div>
);

const CatBadge = ({ cat }) => {
  const c = { ENTERPRISE_ONLINE: 'bg-blue-100 text-blue-700', ADDITIONAL: 'bg-purple-100 text-purple-700', ADDITIONAL_ON_PREMISE: 'bg-orange-100 text-orange-700' }[cat] || 'bg-gray-100';
  return <span className={`inline-block px-2 py-0.5 rounded text-xs font-medium ${c}`}>{PRODUCT_CATEGORIES[cat]?.short}</span>;
};

const ConfirmModal = ({ isOpen, onConfirm, onCancel, title, message }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-xl p-6 max-w-md mx-4">
        <h3 className="text-lg font-semibold text-gray-900 mb-2">{title}</h3>
        <p className="text-gray-600 mb-6">{message}</p>
        <div className="flex justify-end gap-3">
          <button onClick={onCancel} className="px-4 py-2 border border-gray-300 rounded-lg hover:bg-gray-50">Cancel</button>
          <button onClick={onConfirm} className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Create New</button>
        </div>
      </div>
    </div>
  );
};

const Toast = ({ message, onClose }) => {
  React.useEffect(() => {
    const timer = setTimeout(onClose, 3000);
    return () => clearTimeout(timer);
  }, [onClose]);
  
  return (
    <div className="fixed top-20 right-4 bg-green-600 text-white px-4 py-3 rounded-lg shadow-lg z-50 flex items-center gap-2">
      <Clipboard size={18} />
      {message}
    </div>
  );
};

export default function CrayonCostingApp() {
  const [view, setView] = useState('form');
  const [status, setStatus] = useState('draft');
  const [sheetId, setSheetId] = useState('CS-2025-001');
  const [showConfirm, setShowConfirm] = useState(false);
  const [toast, setToast] = useState(null);
  
  const [header, setHeader] = useState(getEmptyHeader);
  const [lineItems, setLineItems] = useState(getEmptyLineItems);
  const [lineItemsY2, setLineItemsY2] = useState(getEmptyLineItems);
  const [lineItemsY3, setLineItemsY3] = useState(getEmptyLineItems);
  const [discounts, setDiscounts] = useState(getEmptyDiscounts);
  const [activeYear, setActiveYear] = useState(1);
  
  const handlePaste = (e, itemId) => {
    const pastedText = e.clipboardData.getData('text');
    
    if (!pastedText.includes('\t')) {
      return;
    }
    
    e.preventDefault();
    
    const rows = joinMultiLineQuotedFields(pastedText);
    
    if (rows.length === 0) return;
    
    // Get current year's line items
    const currentLineItems = getCurrentLineItems();
    const setCurrentLineItems = header.dealType === 'Ramped' 
      ? (activeYear === 1 ? setLineItems : activeYear === 2 ? setLineItemsY2 : setLineItemsY3)
      : setLineItems;
    
    const currentIndex = currentLineItems.findIndex(i => i.id === itemId);
    const currentItem = currentLineItems[currentIndex];
    let currentCategory = currentItem?.category || 'ENTERPRISE_ONLINE';
    
    const newItems = [];
    let maxId = Math.max(...currentLineItems.map(i => i.id), 0);
    
    for (const row of rows) {
      const parsed = parseExcelRow(row, currentCategory);
      
      if (!parsed) continue;
      
      if (parsed.isCategory) {
        currentCategory = parsed.category;
        continue;
      }
      
      maxId++;
      newItems.push({
        id: maxId,
        ...parsed.data,
        category: currentCategory
      });
    }
    
    if (newItems.length === 0) return;
    
    const updatedItems = [...currentLineItems];
    
    if (newItems.length === 1) {
      const singleItem = newItems[0];
      updatedItems[currentIndex] = {
        ...currentItem,
        ...singleItem,
        id: currentItem.id
      };
    } else {
      updatedItems.splice(currentIndex, 1, ...newItems.map((item, idx) => ({
        ...item,
        id: idx === 0 ? currentItem.id : item.id
      })));
    }
    
    setCurrentLineItems(updatedItems);
    setToast(`Pasted ${newItems.length} row${newItems.length > 1 ? 's' : ''} from Excel${header.dealType === 'Ramped' ? ` (Year ${activeYear})` : ''}`);
  };
  
  const handleSave = () => {
    // In a real app, this would call an API to save the data
    // For now, show a confirmation toast
    setToast(`Costing sheet ${sheetId} saved successfully!`);
  };
  
  const handleNewSheetClick = () => setShowConfirm(true);
  
  const confirmNewSheet = () => {
    setSheetId(generateSheetId());
    setHeader(getEmptyHeader());
    setLineItems(getEmptyLineItems());
    setLineItemsY2(getEmptyLineItems());
    setLineItemsY3(getEmptyLineItems());
    setDiscounts(getEmptyDiscounts());
    setActiveYear(1);
    setStatus('draft');
    setView('form');
    setShowConfirm(false);
  };
  
  const cancelNewSheet = () => setShowConfirm(false);
  
  const calculated = useMemo(() => lineItems.map(i => calculateLineItem(i, header.exchangeRate, header.dealType)), [lineItems, header.exchangeRate, header.dealType]);
  const calculatedY2 = useMemo(() => lineItemsY2.map(i => calculateLineItem(i, header.exchangeRate, header.dealType)), [lineItemsY2, header.exchangeRate, header.dealType]);
  const calculatedY3 = useMemo(() => lineItemsY3.map(i => calculateLineItem(i, header.exchangeRate, header.dealType)), [lineItemsY3, header.exchangeRate, header.dealType]);
  
  const sorted = useMemo(() => [...calculated].sort((a, b) => {
    const order = ['ENTERPRISE_ONLINE', 'ADDITIONAL', 'ADDITIONAL_ON_PREMISE'];
    return order.indexOf(a.category) - order.indexOf(b.category);
  }), [calculated]);
  
  const sortedY2 = useMemo(() => [...calculatedY2].sort((a, b) => {
    const order = ['ENTERPRISE_ONLINE', 'ADDITIONAL', 'ADDITIONAL_ON_PREMISE'];
    return order.indexOf(a.category) - order.indexOf(b.category);
  }), [calculatedY2]);
  
  const sortedY3 = useMemo(() => [...calculatedY3].sort((a, b) => {
    const order = ['ENTERPRISE_ONLINE', 'ADDITIONAL', 'ADDITIONAL_ON_PREMISE'];
    return order.indexOf(a.category) - order.indexOf(b.category);
  }), [calculatedY3]);
  
  // Summary uses already-rounded values from each line item (matching Excel)
  const summary = useMemo(() => {
    const isRamped = header.dealType === 'Ramped';
    
    // Year 1 calculations (always from lineItems / calculated)
    const netY1 = calculated.reduce((s, i) => s + i.totalNet, 0);
    const erpY1 = calculated.reduce((s, i) => s + i.totalErp, 0);
    const eupY1 = calculated.reduce((s, i) => s + i.totalEup, 0);
    const rebY1 = calculated.reduce((s, i) => s + i.rebateAmount, 0);
    const gpY1Item = calculated.reduce((s, i) => s + i.gp, 0);
    const swoGpY1 = calculated.reduce((s, i) => s + i.swoGp, 0);
    const partnerGpY1 = calculated.reduce((s, i) => s + i.partnerGp, 0);
    
    // Year 2 & 3 calculations (different for Normal vs Ramped)
    let netY2, netY3, erpY2, erpY3, eupY2, eupY3, rebY2, rebY3;
    let gpY2Item, gpY3Item, swoGpY2, swoGpY3, partnerGpY2, partnerGpY3;
    
    if (isRamped) {
      // Ramped: Use actual Year 2 and Year 3 line items
      netY2 = calculatedY2.reduce((s, i) => s + i.totalNet, 0);
      netY3 = calculatedY3.reduce((s, i) => s + i.totalNet, 0);
      erpY2 = calculatedY2.reduce((s, i) => s + i.totalErp, 0);
      erpY3 = calculatedY3.reduce((s, i) => s + i.totalErp, 0);
      eupY2 = calculatedY2.reduce((s, i) => s + i.totalEup, 0);
      eupY3 = calculatedY3.reduce((s, i) => s + i.totalEup, 0);
      rebY2 = calculatedY2.reduce((s, i) => s + i.rebateAmount, 0);
      rebY3 = calculatedY3.reduce((s, i) => s + i.rebateAmount, 0);
      gpY2Item = calculatedY2.reduce((s, i) => s + i.gp, 0);
      gpY3Item = calculatedY3.reduce((s, i) => s + i.gp, 0);
      swoGpY2 = calculatedY2.reduce((s, i) => s + i.swoGp, 0);
      swoGpY3 = calculatedY3.reduce((s, i) => s + i.swoGp, 0);
      partnerGpY2 = calculatedY2.reduce((s, i) => s + i.partnerGp, 0);
      partnerGpY3 = calculatedY3.reduce((s, i) => s + i.partnerGp, 0);
    } else {
      // Normal: Year 2 & 3 same as Year 1
      netY2 = netY1; netY3 = netY1;
      erpY2 = erpY1; erpY3 = erpY1;
      eupY2 = eupY1; eupY3 = eupY1;
      rebY2 = 0; rebY3 = 0; // Rebate only Year 1 for Normal
      gpY2Item = gpY1Item; gpY3Item = gpY1Item;
      swoGpY2 = swoGpY1; swoGpY3 = swoGpY1;
      partnerGpY2 = partnerGpY1; partnerGpY3 = partnerGpY1;
    }
    
    const dY1 = discounts.year1;
    const dY2 = discounts.year2;
    const dY3 = discounts.year3;
    
    const edY1 = eupY1 - dY1, edY2 = eupY2 - dY2, edY3 = eupY3 - dY3;
    const ed3 = edY1 + edY2 + edY3;
    
    const gpY1Profit = edY1 - netY1, gpY2Profit = edY2 - netY2, gpY3Profit = edY3 - netY3;
    const gp3 = gpY1Profit + gpY2Profit + gpY3Profit;
    
    const grY1 = gpY1Profit + rebY1;
    const grY2 = gpY2Profit + rebY2;
    const grY3 = gpY3Profit + rebY3;
    const gr3 = grY1 + grY2 + grY3;
    
    const netTotal = netY1 + netY2 + netY3;
    const erpTotal = erpY1 + erpY2 + erpY3;
    const eupTotal = eupY1 + eupY2 + eupY3;
    const swoGpTotal = swoGpY1 + swoGpY2 + swoGpY3;
    const partnerGpTotal = partnerGpY1 + partnerGpY2 + partnerGpY3;
    
    return { 
      net: { y1: Math.round(netY1 * 100) / 100, y2: Math.round(netY2 * 100) / 100, y3: Math.round(netY3 * 100) / 100, t: Math.round(netTotal * 100) / 100 }, 
      erp: { y1: Math.round(erpY1 * 100) / 100, y2: Math.round(erpY2 * 100) / 100, y3: Math.round(erpY3 * 100) / 100, t: Math.round(erpTotal * 100) / 100 },
      eup: { y1: Math.round(eupY1 * 100) / 100, y2: Math.round(eupY2 * 100) / 100, y3: Math.round(eupY3 * 100) / 100, t: Math.round(eupTotal * 100) / 100 }, 
      disc: { y1: Math.round(dY1 * 100) / 100, y2: Math.round(dY2 * 100) / 100, y3: Math.round(dY3 * 100) / 100, t: Math.round((dY1 + dY2 + dY3) * 100) / 100 }, 
      ed: { y1: Math.round(edY1 * 100) / 100, y2: Math.round(edY2 * 100) / 100, y3: Math.round(edY3 * 100) / 100, t: Math.round(ed3 * 100) / 100 }, 
      edVat: Math.round(ed3 * (1 + header.vatRate) * 100) / 100, 
      gp: { y1: Math.round(gpY1Item * 100) / 100, y2: Math.round(gpY2Item * 100) / 100, y3: Math.round(gpY3Item * 100) / 100, t: Math.round((gpY1Item + gpY2Item + gpY3Item) * 100) / 100 },
      swoGp: { y1: Math.round(swoGpY1 * 100) / 100, y2: Math.round(swoGpY2 * 100) / 100, y3: Math.round(swoGpY3 * 100) / 100, t: Math.round(swoGpTotal * 100) / 100 },
      partnerGp: { y1: Math.round(partnerGpY1 * 100) / 100, y2: Math.round(partnerGpY2 * 100) / 100, y3: Math.round(partnerGpY3 * 100) / 100, t: Math.round(partnerGpTotal * 100) / 100 },
      gr: { y1: Math.round(grY1 * 100) / 100, y2: Math.round(grY2 * 100) / 100, y3: Math.round(grY3 * 100) / 100, t: Math.round(gr3 * 100) / 100, m: netTotal > 0 ? gr3 / netTotal : 0 }, 
      reb: { y1: Math.round(rebY1 * 100) / 100, y2: Math.round(rebY2 * 100) / 100, y3: Math.round(rebY3 * 100) / 100 } 
    };
  }, [calculated, calculatedY2, calculatedY3, discounts, header.exchangeRate, header.vatRate, header.dealType]);
  
  const add = () => {
    const newItem = { id: Date.now(), partNumber: '', itemName: '', category: 'ENTERPRISE_ONLINE', unitNetUsd: 0, unitErpUsd: 0, msDiscountPct: 0, crayonMarkupPct: 0, unitType: 12, quantity: 0, rebatePct: 0, swoGpPct: 0.5 };
    if (header.dealType === 'Ramped') {
      if (activeYear === 1) setLineItems([...lineItems, { ...newItem, id: Math.max(...lineItems.map(i => i.id), 0) + 1 }]);
      else if (activeYear === 2) setLineItemsY2([...lineItemsY2, { ...newItem, id: Math.max(...lineItemsY2.map(i => i.id), 0) + 1 }]);
      else setLineItemsY3([...lineItemsY3, { ...newItem, id: Math.max(...lineItemsY3.map(i => i.id), 0) + 1 }]);
    } else {
      setLineItems([...lineItems, { ...newItem, id: Math.max(...lineItems.map(i => i.id), 0) + 1 }]);
    }
  };
  
  const upd = (id, f, v) => {
    if (header.dealType === 'Ramped') {
      if (activeYear === 1) setLineItems(lineItems.map(i => i.id === id ? { ...i, [f]: v } : i));
      else if (activeYear === 2) setLineItemsY2(lineItemsY2.map(i => i.id === id ? { ...i, [f]: v } : i));
      else setLineItemsY3(lineItemsY3.map(i => i.id === id ? { ...i, [f]: v } : i));
    } else {
      setLineItems(lineItems.map(i => i.id === id ? { ...i, [f]: v } : i));
    }
  };
  
  const del = id => {
    if (header.dealType === 'Ramped') {
      if (activeYear === 1) setLineItems(lineItems.filter(i => i.id !== id));
      else if (activeYear === 2) setLineItemsY2(lineItemsY2.filter(i => i.id !== id));
      else setLineItemsY3(lineItemsY3.filter(i => i.id !== id));
    } else {
      setLineItems(lineItems.filter(i => i.id !== id));
    }
  };
  
  // Copy Year 1 to Year 2
  const copyY1toY2 = () => {
    const copied = lineItems.map((item, idx) => ({ ...item, id: idx + 1 }));
    setLineItemsY2(copied);
    setToast('Year 1 data copied to Year 2');
  };
  
  // Copy Year 1 to Year 3
  const copyY1toY3 = () => {
    const copied = lineItems.map((item, idx) => ({ ...item, id: idx + 1 }));
    setLineItemsY3(copied);
    setToast('Year 1 data copied to Year 3');
  };
  
  // Copy Year 1 to both Year 2 and Year 3
  const copyY1toAll = () => {
    const copied = lineItems.map((item, idx) => ({ ...item, id: idx + 1 }));
    setLineItemsY2(copied);
    setLineItemsY3(copied);
    setToast('Year 1 data copied to Year 2 & Year 3');
  };
  
  // Get current year's line items and calculated values
  const getCurrentLineItems = () => {
    if (header.dealType !== 'Ramped') return lineItems;
    if (activeYear === 1) return lineItems;
    if (activeYear === 2) return lineItemsY2;
    return lineItemsY3;
  };
  
  const getCurrentCalculated = () => {
    if (header.dealType !== 'Ramped') return calculated;
    if (activeYear === 1) return calculated;
    if (activeYear === 2) return calculatedY2;
    return calculatedY3;
  };
  
  const getCurrentSorted = () => {
    if (header.dealType !== 'Ramped') return sorted;
    if (activeYear === 1) return sorted;
    if (activeYear === 2) return sortedY2;
    return sortedY3;
  };
  
  const ed = status === 'draft';
  
  // Excel Export Function
  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    
    // === COSTING SHEET ===
    const costingData = [];
    
    // Header info
    costingData.push(['Customer Name', header.customerName]);
    costingData.push(['Sales Location', header.salesLocation]);
    costingData.push(['Account Manager', header.accountManager]);
    if (header.region === 'AF') costingData.push(['Partner Name', header.partnerName]);
    costingData.push(['Agreement Type', header.agreementType]);
    costingData.push(['New/Renewal', header.newOrRenewal]);
    costingData.push(['Currency', header.currencyCode]);
    costingData.push(['Exchange Rate', header.exchangeRate]);
    costingData.push(['VAT Rate', header.vatRate * 100 + '%']);
    costingData.push(['Region', header.region === 'ME' ? 'Middle East' : 'Africa']);
    costingData.push(['Deal Type', header.dealType]);
    costingData.push([]);
    
    // Line items header
    const lineItemHeaders = ['Category', 'Part Number', 'Item Name', 'Unit Net USD', 'Unit ERP USD', 'Default Markup %', 'MS Discount %', 'Crayon Markup %', 'Unit Type', 'MS Disc Net', 'MS Disc ERP', 'Total Net', 'Total ERP', 'Qty', 'EUP Unit', 'Total EUP/Yr', 'Markup %', 'Rebate %', 'Rebate'];
    if (header.region === 'AF') {
      lineItemHeaders.push('GP', 'SWO GP %', 'SWO GP', 'Partner GP');
    }
    
    // Helper function to add line items
    const addLineItems = (items, yearLabel) => {
      if (yearLabel) {
        costingData.push([]);
        costingData.push([`=== ${yearLabel} ===`]);
      }
      costingData.push(lineItemHeaders);
      
      items.forEach(item => {
        if (item.partNumber || item.itemName) {
          const row = [
            PRODUCT_CATEGORIES[item.category]?.short || item.category,
            item.partNumber,
            item.itemName,
            item.unitNetUsd,
            item.unitErpUsd,
            ((item.unitErpUsd - item.unitNetUsd) / item.unitNetUsd * 100).toFixed(2) + '%',
            (item.msDiscountPct * 100).toFixed(2) + '%',
            (item.crayonMarkupPct * 100).toFixed(2) + '%',
            item.unitType,
            item.discountedNet,
            item.discountedErp,
            item.totalNet,
            item.totalErp,
            item.quantity,
            item.eupUnit,
            item.totalEup,
            (item.calculatedMarkup * 100).toFixed(2) + '%',
            (item.rebatePct * 100).toFixed(2) + '%',
            item.rebateAmount
          ];
          if (header.region === 'AF') {
            row.push(item.gp, (item.swoGpPct * 100).toFixed(2) + '%', item.swoGp, item.partnerGp);
          }
          costingData.push(row);
        }
      });
    };
    
    if (header.dealType === 'Ramped') {
      // Ramped: Add 3 separate year tables
      addLineItems(sorted, 'YEAR 1');
      const netY1 = sorted.reduce((s, i) => s + i.totalNet, 0);
      const erpY1 = sorted.reduce((s, i) => s + i.totalErp, 0);
      const eupY1 = sorted.reduce((s, i) => s + i.totalEup, 0);
      costingData.push(['YEAR 1 TOTALS', '', '', '', '', '', '', '', '', '', '', netY1, erpY1, '', '', eupY1]);
      
      addLineItems(sortedY2, 'YEAR 2');
      const netY2 = sortedY2.reduce((s, i) => s + i.totalNet, 0);
      const erpY2 = sortedY2.reduce((s, i) => s + i.totalErp, 0);
      const eupY2 = sortedY2.reduce((s, i) => s + i.totalEup, 0);
      costingData.push(['YEAR 2 TOTALS', '', '', '', '', '', '', '', '', '', '', netY2, erpY2, '', '', eupY2]);
      
      addLineItems(sortedY3, 'YEAR 3');
      const netY3 = sortedY3.reduce((s, i) => s + i.totalNet, 0);
      const erpY3 = sortedY3.reduce((s, i) => s + i.totalErp, 0);
      const eupY3 = sortedY3.reduce((s, i) => s + i.totalEup, 0);
      costingData.push(['YEAR 3 TOTALS', '', '', '', '', '', '', '', '', '', '', netY3, erpY3, '', '', eupY3]);
      
      costingData.push([]);
      costingData.push(['GRAND TOTAL (3 Years)', '', '', '', '', '', '', '', '', '', '', netY1 + netY2 + netY3, erpY1 + erpY2 + erpY3, '', '', eupY1 + eupY2 + eupY3]);
    } else {
      // Normal: Single table
      addLineItems(sorted, null);
      costingData.push([]);
      costingData.push(['TOTALS (Yearly)', '', '', '', '', '', '', '', '', '', '', summary.net.y1, summary.erp.y1, '', '', summary.eup.y1]);
      costingData.push(['TOTALS (3 Years)', '', '', '', '', '', '', '', '', '', '', summary.net.t, summary.erp.t, '', '', summary.eup.t]);
    }
    
    const wsCosting = XLSX.utils.aoa_to_sheet(costingData);
    XLSX.utils.book_append_sheet(wb, wsCosting, 'Costing');
    
    // === MERGED SHEET ===
    const mergedData = [];
    
    // Header section
    mergedData.push(['Customer Name', header.customerName, '', 'Exchange Rate', header.exchangeRate]);
    mergedData.push(['Sales Location', header.salesLocation]);
    mergedData.push(['Account Manager', header.accountManager]);
    if (header.region === 'AF') mergedData.push(['Partner Name', header.partnerName]);
    mergedData.push(['Agreement', header.agreementType]);
    mergedData.push(['New/Renewal', header.newOrRenewal]);
    mergedData.push(['Agreement Level - System', header.agreementLevelSystem]);
    mergedData.push(['Agreement Level - Server', header.agreementLevelServer]);
    mergedData.push(['Agreement Level - Application', header.agreementLevelApplication]);
    mergedData.push(['Currency', header.currencyCode]);
    mergedData.push([]);
    
    // Cost Price section
    mergedData.push(['Cost Price / CPS Price', 'Values']);
    mergedData.push(['Total Net Year 1', summary.net.y1]);
    mergedData.push(['Total Net Year 2', summary.net.y1]);
    mergedData.push(['Total Net Year 3', summary.net.y1]);
    mergedData.push(['Grand Total Net Over 3 Years', summary.net.t]);
    mergedData.push([]);
    
    // Estimated Retail Price
    const defaultMarkup = summary.net.y1 > 0 ? (summary.erp.y1 - summary.net.y1) / summary.net.y1 : 0;
    mergedData.push(['Estimated Retail Price', 'Values', 'Default Markup %', 'Default GP']);
    mergedData.push(['Total ERP Year 1', summary.erp.y1, (defaultMarkup * 100).toFixed(2) + '%', summary.erp.y1 - summary.net.y1]);
    mergedData.push(['Total ERP Year 2', summary.erp.y1, (defaultMarkup * 100).toFixed(2) + '%', summary.erp.y1 - summary.net.y1]);
    mergedData.push(['Total ERP Year 3', summary.erp.y1, (defaultMarkup * 100).toFixed(2) + '%', summary.erp.y1 - summary.net.y1]);
    mergedData.push(['Grand Total ERP Over 3 Years', summary.erp.t, '', summary.erp.t - summary.net.t]);
    mergedData.push([]);
    
    // EUP without Discount
    mergedData.push(['End User Price without Crayon Discount', 'Values']);
    mergedData.push(['Total EUP Year 1', summary.eup.y1]);
    mergedData.push(['Total EUP Year 2', summary.eup.y1]);
    mergedData.push(['Total EUP Year 3', summary.eup.y1]);
    mergedData.push(['Grand Total EUP (3 Years) w/o Discount', summary.eup.t]);
    mergedData.push([]);
    
    // Crayon Discount
    mergedData.push(['Crayon Discount/Funding', 'Values']);
    mergedData.push(['Discount Value Year 1', summary.disc.y1]);
    mergedData.push(['Discount Value Year 2', summary.disc.y2]);
    mergedData.push(['Discount Value Year 3', summary.disc.y3]);
    mergedData.push(['Total Discount', summary.disc.t]);
    mergedData.push([]);
    
    // EUP with Discount
    mergedData.push(['End User Price with Crayon Discount', 'Values']);
    mergedData.push(['Total EUP Year 1 with Discount', summary.ed.y1]);
    mergedData.push(['Total EUP Year 2 with Discount', summary.ed.y2]);
    mergedData.push(['Total EUP Year 3 with Discount', summary.ed.y3]);
    mergedData.push(['Grand Total EUP (3 Years) w/ Discount', summary.ed.t]);
    mergedData.push(['Grand Total EUP (3 Years) w/ Discount + VAT', summary.edVat]);
    mergedData.push([]);
    
    // Crayon Rebate
    mergedData.push(['Crayon Rebate', 'Values']);
    mergedData.push(['Rebate Year 1', summary.reb.y1]);
    mergedData.push(['Rebate Year 2', '—']);
    mergedData.push(['Rebate Year 3', '—']);
    mergedData.push(['Total Rebate Over 3 Years', summary.reb.y1]);
    mergedData.push([]);
    
    // GP without Rebates
    if (header.region === 'AF') {
      mergedData.push(['GP without Rebates', 'Crayon GP', 'Partner GP']);
      mergedData.push(['GP Year 1', summary.swoGp.y1, summary.partnerGp.y1]);
      mergedData.push(['GP Year 2', summary.swoGp.y1, summary.partnerGp.y1]);
      mergedData.push(['GP Year 3', summary.swoGp.y1, summary.partnerGp.y1]);
      mergedData.push(['GP Over 3 Years', summary.swoGp.t, summary.partnerGp.t]);
      mergedData.push(['Markup', (summary.swoGp.t / summary.net.t * 100).toFixed(2) + '%', (summary.partnerGp.t / summary.net.t * 100).toFixed(2) + '%']);
    } else {
      mergedData.push(['GP without Rebates', 'Values']);
      mergedData.push(['GP Year 1', summary.ed.y1 - summary.net.y1]);
      mergedData.push(['GP Year 2', summary.ed.y2 - summary.net.y1]);
      mergedData.push(['GP Year 3', summary.ed.y3 - summary.net.y1]);
      mergedData.push(['GP Over 3 Years', summary.ed.t - summary.net.t]);
      mergedData.push(['Markup %', ((summary.ed.t - summary.net.t) / summary.net.t * 100).toFixed(2) + '%']);
    }
    mergedData.push([]);
    
    // Gross Profit with Rebates
    if (header.region === 'AF') {
      mergedData.push(['Gross Profit with Rebates', 'Values']);
      mergedData.push(['GP + Rebate Year 1', summary.swoGp.y1 + summary.reb.y1]);
      mergedData.push(['GP + Rebate Year 2', summary.swoGp.y2 + summary.reb.y2]);
      mergedData.push(['GP + Rebate Year 3', summary.swoGp.y3 + summary.reb.y3]);
      mergedData.push(['Total GP + Rebate Over 3 Years', summary.swoGp.t + summary.reb.y1 + summary.reb.y2 + summary.reb.y3]);
      mergedData.push(['Overall Markup', ((summary.swoGp.t + summary.reb.y1 + summary.reb.y2 + summary.reb.y3) / summary.net.t * 100).toFixed(2) + '%']);
    } else {
      mergedData.push(['Gross Profit with Rebates', 'Values']);
      mergedData.push(['GP + Rebate Year 1', summary.gr.y1]);
      mergedData.push(['GP + Rebate Year 2', summary.gr.y2]);
      mergedData.push(['GP + Rebate Year 3', summary.gr.y3]);
      mergedData.push(['Total GP + Rebate Over 3 Years', summary.gr.t]);
      mergedData.push(['Overall Markup', (summary.gr.m * 100).toFixed(2) + '%']);
    }
    mergedData.push([]);
    
    // Gross Profit with Rebates + Crayon Cost (after bid bond/performance bond/tender costs)
    // Bid Bond Cost = only Year 1
    const bidBondCost = summary.edVat * header.bidBondPct * header.bankChargesPct;
    // Performance Bond Cost = EACH year
    const perfBondCostPerYear = summary.edVat * header.performanceBondPct * header.performanceBankChargesPct;
    const totalPBCost = perfBondCostPerYear * 3;
    // Tender Cost = only Year 1
    const tenderCost = header.tenderCost || 0;
    const totalCrayonCost = bidBondCost + totalPBCost + tenderCost;
    
    const gpRebY1 = header.region === 'AF' ? summary.swoGp.y1 + summary.reb.y1 : summary.gr.y1;
    const gpRebY2 = header.region === 'AF' ? summary.swoGp.y2 + summary.reb.y2 : summary.gr.y2;
    const gpRebY3 = header.region === 'AF' ? summary.swoGp.y3 + summary.reb.y3 : summary.gr.y3;
    
    // Year 1: subtract BB cost + PB cost + Tender cost
    // Year 2 & 3: subtract only PB cost
    const gpRebCrayonY1 = gpRebY1 - bidBondCost - perfBondCostPerYear - tenderCost;
    const gpRebCrayonY2 = gpRebY2 - perfBondCostPerYear;
    const gpRebCrayonY3 = gpRebY3 - perfBondCostPerYear;
    const gpRebCrayonTotal = gpRebCrayonY1 + gpRebCrayonY2 + gpRebCrayonY3;
    const overallCrayonMarkup = summary.net.t > 0 ? gpRebCrayonTotal / summary.net.t : 0;
    
    mergedData.push(['Gross Profit with Rebates + Crayon Cost', 'Values']);
    mergedData.push(['GP + Rebate + Crayon Cost Year 1', gpRebCrayonY1]);
    mergedData.push(['GP + Rebate + Crayon Cost Year 2', gpRebCrayonY2]);
    mergedData.push(['GP + Rebate + Crayon Cost Year 3', gpRebCrayonY3]);
    mergedData.push(['Total GP + Rebate + Crayon Cost Over 3 Years', gpRebCrayonTotal]);
    mergedData.push(['Overall Markup %', (overallCrayonMarkup * 100).toFixed(2) + '%']);
    mergedData.push([]);
    
    // Bid Bond & Bank Charges
    mergedData.push(['Bid Bond & Bank Charges', '', '']);
    mergedData.push(['', '%', 'Value']);
    mergedData.push(['Bid Bond', (header.bidBondPct * 100).toFixed(2) + '%', summary.edVat * header.bidBondPct]);
    mergedData.push(['Bank Charges', (header.bankChargesPct * 100).toFixed(2) + '%', '—']);
    mergedData.push(['Total BB Cost', '', bidBondCost]);
    mergedData.push([]);
    mergedData.push(['Performance Bond', (header.performanceBondPct * 100).toFixed(2) + '%', summary.edVat * header.performanceBondPct]);
    mergedData.push(['Bank Charges', (header.performanceBankChargesPct * 100).toFixed(2) + '%', '—']);
    mergedData.push(['Cost Year 1', '', perfBondCostPerYear]);
    mergedData.push(['Cost Year 2', '', perfBondCostPerYear]);
    mergedData.push(['Cost Year 3', '', perfBondCostPerYear]);
    mergedData.push(['Total PB Cost over 3 years', '', totalPBCost]);
    mergedData.push([]);
    mergedData.push(['Tender Cost', '', tenderCost || '—']);
    mergedData.push([]);
    mergedData.push(['Total Crayon Cost', '', totalCrayonCost]);
    mergedData.push([]);
    
    // CIF Products
    mergedData.push(['CIF Products', 'Yearly Value']);
    mergedData.push(['M365E5', header.cifM365E5 || '—']);
    mergedData.push(['M365E3', header.cifM365E3 || '—']);
    mergedData.push(['Azure', header.cifAzure || '—']);
    mergedData.push(['Dynamics365', header.cifDynamics365 || '—']);
    
    const wsMerged = XLSX.utils.aoa_to_sheet(mergedData);
    XLSX.utils.book_append_sheet(wb, wsMerged, 'Merged');
    
    // === FINAL PRICE TABLE ===
    const priceTableData = [];
    
    // Header
    priceTableData.push(['PRICE QUOTATION']);
    priceTableData.push(['Customer:', header.customerName]);
    priceTableData.push(['Currency:', header.currencyCode, 'Exchange Rate:', header.exchangeRate]);
    priceTableData.push(['VAT:', (header.vatRate * 100) + '%']);
    if (header.dealType === 'Ramped') priceTableData.push(['Deal Type:', 'Ramped Deal']);
    priceTableData.push([]);
    
    // Helper function to add year table for Final Price Table
    const addPriceYearTable = (items, yearLabel, yearDiscount, yearNum) => {
      priceTableData.push([]);
      priceTableData.push([`=== ${yearLabel} ===`]);
      priceTableData.push(['Part Number', 'Item Name', 'Qty', 'Unit Type', 'EUP', `${yearLabel} Total`]);
      
      // Enterprise Online Products
      const enterpriseItems = items.filter(i => i.category === 'ENTERPRISE_ONLINE' && (i.partNumber || i.itemName));
      if (enterpriseItems.length > 0) {
        priceTableData.push(['Enterprise Online Products']);
        enterpriseItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup]);
        });
      }
      
      // Additional Products
      const additionalItems = items.filter(i => i.category === 'ADDITIONAL' && (i.partNumber || i.itemName));
      if (additionalItems.length > 0) {
        priceTableData.push(['Additional Products']);
        additionalItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup]);
        });
      }
      
      // Additional Products - On Premise
      const onPremiseItems = items.filter(i => i.category === 'ADDITIONAL_ON_PREMISE' && (i.partNumber || i.itemName));
      if (onPremiseItems.length > 0) {
        priceTableData.push(['Additional Products - On Premise']);
        onPremiseItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup]);
        });
      }
      
      // Year totals
      const yearEup = items.reduce((s, i) => s + i.totalEup, 0);
      const yearEd = yearEup - yearDiscount;
      const yearVat = yearEd * header.vatRate;
      const yearTotal = yearEd + yearVat;
      
      priceTableData.push([]);
      priceTableData.push(['', '', '', '', 'Total ' + header.currencyCode, yearEup]);
      priceTableData.push(['', '', '', '', 'Further Discount from Crayon', yearDiscount]);
      priceTableData.push(['', '', '', '', 'Total after discount', yearEd]);
      priceTableData.push(['', '', '', '', 'VAT ' + (header.vatRate * 100) + '%', yearVat]);
      priceTableData.push(['', '', '', '', 'Grand Total with VAT ' + header.currencyCode, yearTotal]);
      
      return { eup: yearEup, disc: yearDiscount, ed: yearEd, vat: yearVat, total: yearTotal };
    };
    
    if (header.dealType === 'Ramped') {
      // Ramped: Add 3 separate year tables + Grand Total Summary
      const y1Totals = addPriceYearTable(sorted, 'Year 1', discounts.year1, 1);
      const y2Totals = addPriceYearTable(sortedY2, 'Year 2', discounts.year2, 2);
      const y3Totals = addPriceYearTable(sortedY3, 'Year 3', discounts.year3, 3);
      
      // Grand Total Summary
      priceTableData.push([]);
      priceTableData.push([]);
      priceTableData.push(['=== GRAND TOTAL SUMMARY (3 Years) ===']);
      priceTableData.push(['', '', '', '', 'Year 1', 'Year 2', 'Year 3', 'Total 3 Years']);
      priceTableData.push(['', '', '', 'Total EUP', y1Totals.eup, y2Totals.eup, y3Totals.eup, y1Totals.eup + y2Totals.eup + y3Totals.eup]);
      priceTableData.push(['', '', '', 'Discount', y1Totals.disc, y2Totals.disc, y3Totals.disc, y1Totals.disc + y2Totals.disc + y3Totals.disc]);
      priceTableData.push(['', '', '', 'After Discount', y1Totals.ed, y2Totals.ed, y3Totals.ed, y1Totals.ed + y2Totals.ed + y3Totals.ed]);
      priceTableData.push(['', '', '', 'VAT', y1Totals.vat, y2Totals.vat, y3Totals.vat, y1Totals.vat + y2Totals.vat + y3Totals.vat]);
      priceTableData.push(['', '', '', 'Grand Total with VAT', y1Totals.total, y2Totals.total, y3Totals.total, y1Totals.total + y2Totals.total + y3Totals.total]);
    } else {
      // Normal: Single table with Yr.1, Yr.2, Yr.3 columns
      priceTableData.push(['Part Number', 'Item Name', 'Qty', 'Unit Type', 'EUP', 'Yr.1 Total', 'Yr.2 Total', 'Yr.3 Total', 'Total Over 3 Years']);
      
      // Enterprise Online Products
      const enterpriseItems = sorted.filter(i => i.category === 'ENTERPRISE_ONLINE' && (i.partNumber || i.itemName));
      if (enterpriseItems.length > 0) {
        priceTableData.push(['Enterprise Online Products']);
        enterpriseItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup, item.totalEup, item.totalEup, item.totalEup * 3]);
        });
      }
      
      // Additional Products
      const additionalItems = sorted.filter(i => i.category === 'ADDITIONAL' && (i.partNumber || i.itemName));
      if (additionalItems.length > 0) {
        priceTableData.push(['Additional Products']);
        additionalItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup, item.totalEup, item.totalEup, item.totalEup * 3]);
        });
      }
      
      // Additional Products - On Premise
      const onPremiseItems = sorted.filter(i => i.category === 'ADDITIONAL_ON_PREMISE' && (i.partNumber || i.itemName));
      if (onPremiseItems.length > 0) {
        priceTableData.push(['Additional Products - On Premise']);
        onPremiseItems.forEach(item => {
          priceTableData.push([item.partNumber, item.itemName, item.quantity, item.unitType, item.eupUnit, item.totalEup, item.totalEup, item.totalEup, item.totalEup * 3]);
        });
      }
      
      priceTableData.push([]);
      
      // Totals
      priceTableData.push(['', '', '', '', 'Total ' + header.currencyCode, summary.eup.y1, summary.eup.y1, summary.eup.y1, summary.eup.t]);
      priceTableData.push(['', '', '', '', 'Further Discount from Crayon', summary.disc.y1, summary.disc.y2, summary.disc.y3, summary.disc.t]);
      priceTableData.push(['', '', '', '', 'Total after discount', summary.ed.y1, summary.ed.y2, summary.ed.y3, summary.ed.t]);
      priceTableData.push(['', '', '', '', 'VAT ' + (header.vatRate * 100) + '%', summary.ed.y1 * header.vatRate, summary.ed.y2 * header.vatRate, summary.ed.y3 * header.vatRate, summary.ed.t * header.vatRate]);
      priceTableData.push(['', '', '', '', 'Grand Total with VAT ' + header.currencyCode, summary.ed.y1 * (1 + header.vatRate), summary.ed.y2 * (1 + header.vatRate), summary.ed.y3 * (1 + header.vatRate), summary.edVat]);
    }
    
    const wsPriceTable = XLSX.utils.aoa_to_sheet(priceTableData);
    
    // Set column widths
    wsPriceTable['!cols'] = [
      { wch: 15 }, { wch: 45 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 18 }
    ];
    
    XLSX.utils.book_append_sheet(wb, wsPriceTable, 'FinalPriceTable');
    
    // Generate file and download
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const fileName = `${header.customerName || 'CostingSheet'}_${sheetId}_${new Date().toISOString().split('T')[0]}.xlsx`;
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), fileName);
    
    setToast('Excel file exported successfully!');
  };
  
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {toast && <Toast message={toast} onClose={() => setToast(null)} />}
      
      <ConfirmModal 
        isOpen={showConfirm}
        onConfirm={confirmNewSheet}
        onCancel={cancelNewSheet}
        title="Create New Costing Sheet"
        message="This will clear all current data and create a blank costing sheet. Are you sure you want to continue?"
      />
      
      <header className="bg-gradient-to-r from-blue-600 via-blue-700 to-indigo-800 text-white shadow-lg">
        <div className="max-w-full mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-lg"><Calculator className="w-6 h-6" /></div>
            <div><h1 className="text-xl font-bold">Crayon Costing Application</h1><p className="text-blue-200 text-sm">Microsoft License Sales</p></div>
          </div>
          <div className="flex items-center gap-4">
            <button onClick={handleNewSheetClick} className="flex items-center gap-2 px-4 py-2 bg-white/20 hover:bg-white/30 rounded-lg transition text-sm font-medium">
              <FilePlus size={18} /> New Costing Sheet
            </button>
            <div className="text-sm bg-white/10 px-3 py-1 rounded font-mono">{sheetId}</div>
            <StatusBadge status={status} />
            <div className="flex items-center gap-2 bg-white/10 px-3 py-2 rounded-lg"><Users size={16} /><span className="text-sm">John Sales</span></div>
          
           <button 
  onClick={() => {
   
    localStorage.removeItem("demo_auth");
    sessionStorage.removeItem("demo_auth");
    
    // 2. Redirect with replace (taaki back button loop na bane)
    // Production mein "/login.html" ya sirf "/login" aapke setup par depend karta hai
    window.location.replace("/login.html");
  }}
  style={{
    padding: "8px 16px",
    color: "Red",
    borderRadius: "6px",
    cursor: "pointer",
    border: "none",
    fontWeight: "bold"
  }}
>
  Logout
</button>
          </div>
        </div>
      </header>
      
      <div className="bg-white border-b sticky top-0 z-10 shadow-sm">
        <div className="max-w-full mx-auto px-4 flex gap-1">
          {[{ id: 'form', l: 'Costing Form', i: Edit3 }, { id: 'summary', l: 'Merged', i: DollarSign }, { id: 'preview', l: 'Final Price Table', i: FileText }].map(t => (
            <button key={t.id} onClick={() => setView(t.id)} className={`flex items-center gap-2 px-5 py-3 font-medium border-b-2 transition ${view === t.id ? 'border-blue-600 text-blue-600 bg-blue-50/50' : 'border-transparent text-gray-500 hover:bg-gray-50'}`}><t.i size={18} /> {t.l}</button>
          ))}
        </div>
      </div>
      
      <main className="max-w-full mx-auto px-4 py-6 pb-24">
        {view === 'form' && (
          <div className="space-y-6">
            <Section title="Customer & Agreement Details">
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
                <div className="col-span-2"><label className="block text-sm font-medium text-gray-700 mb-1">Customer Name *</label><input type="text" value={header.customerName} onChange={e => setHeader({...header, customerName: e.target.value})} placeholder="Enter customer name" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">ERP Customer ID</label><input type="text" value={header.erpCustomerId} onChange={e => setHeader({...header, erpCustomerId: e.target.value})} placeholder="ERP ID" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Opportunity ID</label><input type="text" value={header.opportunityId} onChange={e => setHeader({...header, opportunityId: e.target.value})} placeholder="OPP-XXXX" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Customer Segment</label><input type="text" value={header.customerSegment} onChange={e => setHeader({...header, customerSegment: e.target.value})} placeholder="Segment" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Business Area</label><input type="text" value={header.businessArea} onChange={e => setHeader({...header, businessArea: e.target.value})} placeholder="Business Area" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Region</label><select value={header.region} onChange={e => setHeader({...header, region: e.target.value})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed}><option value="ME">Middle East</option><option value="AF">Africa</option></select></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Deal Type</label><select value={header.dealType} onChange={e => { setHeader({...header, dealType: e.target.value}); setActiveYear(1); }} className={`w-full px-3 py-2 border rounded-lg ${header.dealType === 'Ramped' ? 'bg-purple-50 border-purple-300' : ''}`} disabled={!ed}><option value="Normal">Normal</option><option value="Ramped">Ramped</option></select></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Sales Location</label><input type="text" value={header.salesLocation} onChange={e => setHeader({...header, salesLocation: e.target.value})} placeholder="UAE" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Producer Name</label><input type="text" value={header.producerName} onChange={e => setHeader({...header, producerName: e.target.value})} placeholder="Producer" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Account Manager</label><input type="text" value={header.accountManager} onChange={e => setHeader({...header, accountManager: e.target.value})} placeholder="Name" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Account Manager ID</label><input type="text" value={header.accountManagerId} onChange={e => setHeader({...header, accountManagerId: e.target.value})} placeholder="AM ID" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                {header.region === 'AF' && <div><label className="block text-sm font-medium text-gray-700 mb-1">Partner Name</label><input type="text" value={header.partnerName} onChange={e => setHeader({...header, partnerName: e.target.value})} placeholder="Enter partner name" className="w-full px-3 py-2 border rounded-lg bg-orange-50 border-orange-200" disabled={!ed} /></div>}
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Agreement Type</label><input type="text" value={header.agreementType} onChange={e => setHeader({...header, agreementType: e.target.value})} placeholder="Enterprise Enrollment" className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">New/Renewal</label><select value={header.newOrRenewal} onChange={e => setHeader({...header, newOrRenewal: e.target.value})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed}><option>New</option><option>Renewal</option></select></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Currency</label><input type="text" value={header.currencyCode} onChange={e => setHeader({...header, currencyCode: e.target.value.toUpperCase()})} placeholder="USD" className="w-full px-3 py-2 border rounded-lg" maxLength={3} disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Exchange Rate</label><input type="number" step="0.0001" value={header.exchangeRate} onChange={e => setHeader({...header, exchangeRate: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">VAT %</label><input type="number" step="0.1" value={header.vatRate * 100} onChange={e => setHeader({...header, vatRate: (parseFloat(e.target.value) || 0) / 100})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Agreement Level - System</label><select value={header.agreementLevelSystem} onChange={e => setHeader({...header, agreementLevelSystem: e.target.value})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed}><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Agreement Level - Server</label><select value={header.agreementLevelServer} onChange={e => setHeader({...header, agreementLevelServer: e.target.value})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed}><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Agreement Level - Application</label><select value={header.agreementLevelApplication} onChange={e => setHeader({...header, agreementLevelApplication: e.target.value})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed}><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select></div>
              </div>
            </Section>
            
            <Section title="Line Items" badge={<span className="bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full text-xs font-medium">{getCurrentCalculated().length} products{header.dealType === 'Ramped' ? ` (Year ${activeYear})` : ''}</span>}>
              {header.dealType === 'Ramped' && (
                <div className="mb-4">
                  <div className="flex flex-wrap items-center gap-2 mb-3">
                    <div className="flex rounded-lg overflow-hidden border border-purple-300">
                      {[1, 2, 3].map(year => (
                        <button key={year} onClick={() => setActiveYear(year)} className={`px-4 py-2 text-sm font-medium transition-colors ${activeYear === year ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 hover:bg-purple-50'}`}>
                          Year {year}
                        </button>
                      ))}
                    </div>
                    {ed && activeYear === 1 && (
                      <div className="flex gap-2 ml-4">
                        <button onClick={copyY1toY2} className="px-3 py-1.5 text-xs bg-purple-100 text-purple-700 rounded-lg hover:bg-purple-200 transition-colors">Copy to Year 2</button>
                        <button onClick={copyY1toY3} className="px-3 py-1.5 text-xs bg-purple-100 text-purple-700 rounded-lg hover:bg-purple-200 transition-colors">Copy to Year 3</button>
                        <button onClick={copyY1toAll} className="px-3 py-1.5 text-xs bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors">Copy to All Years</button>
                      </div>
                    )}
                  </div>
                  <div className="p-2 bg-purple-50 border border-purple-200 rounded-lg text-sm text-purple-700">
                    <strong>Ramped Deal:</strong> Each year can have different MS Discount %, Crayon Markup %, and Quantity values. Edit Year 1 first, then copy to other years as a starting point.
                  </div>
                </div>
              )}
              <div className="mb-3 p-3 bg-blue-50 border border-blue-200 rounded-lg text-sm text-blue-700 flex items-center gap-2">
                <Clipboard size={16} />
                <span><strong>Tip:</strong> Copy rows from Excel and paste into the Part Number field to auto-populate. Supports single or multiple rows!</span>
              </div>
              
              <div className="overflow-x-auto -mx-4">
                <table className="w-full text-xs border-collapse min-w-max">
                  <thead>
                    <tr className="bg-gray-100 text-gray-600 text-left">
                      <th className="px-2 py-2 font-semibold border-b sticky left-0 bg-gray-100 z-10 min-w-28">Category</th>
                      <th className="px-2 py-2 font-semibold border-b bg-green-50 sticky left-28 z-10">Part Number</th>
                      <th className="px-2 py-2 font-semibold border-b min-w-52 sticky left-56 bg-gray-100 z-10">Item Name</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Unit Net USD</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Unit ERP USD</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-gray-200">Default Markup %</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">MS Discount %</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Crayon Markup %</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Unit Type</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-blue-50">MS Disc Net</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-blue-50">MS Disc ERP</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-blue-100">Total Net</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-blue-100">Total ERP</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Qty</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-green-50">EUP</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-green-100">Total EUP/Yr</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-gray-200">Markup %</th>
                      <th className="px-2 py-2 font-semibold border-b text-right">Rebate %</th>
                      <th className="px-2 py-2 font-semibold border-b text-right bg-yellow-50">Rebate</th>
                      {header.region === 'AF' && (
                        <>
                          <th className="px-2 py-2 font-semibold border-b text-right bg-purple-50">GP</th>
                          <th className="px-2 py-2 font-semibold border-b text-right bg-purple-100">SWO GP %</th>
                          <th className="px-2 py-2 font-semibold border-b text-right bg-purple-50">SWO GP</th>
                          <th className="px-2 py-2 font-semibold border-b text-right bg-orange-50">Partner GP</th>
                        </>
                      )}
                      <th className="px-2 py-2 border-b w-8"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {getCurrentSorted().map((item) => (
                      <tr key={item.id} className="hover:bg-blue-50/30 border-b border-gray-100">
                        <td className="px-2 py-1 sticky left-0 bg-white z-10">
                          <select value={item.category} onChange={e => upd(item.id, 'category', e.target.value)} className="w-full px-1 py-1 text-xs border rounded" disabled={!ed}>
                            {Object.entries(PRODUCT_CATEGORIES).map(([k, v]) => <option key={k} value={k}>{v.short}</option>)}
                          </select>
                        </td>
                        <td className="px-2 py-1 bg-green-50/30 sticky left-28 z-10">
                          <input type="text" value={item.partNumber} onChange={e => upd(item.id, 'partNumber', e.target.value)} onPaste={e => handlePaste(e, item.id)} placeholder="Paste Excel here" className="w-28 px-1 py-1 border border-green-300 rounded text-xs bg-white" disabled={!ed} />
                        </td>
                        <td className="px-2 py-1 sticky left-56 bg-white z-10"><input type="text" value={item.itemName} onChange={e => upd(item.id, 'itemName', e.target.value)} className="w-full min-w-52 px-1 py-1 border rounded text-xs" disabled={!ed} /></td>
                        <td className="px-2 py-1"><input type="number" step="0.01" value={item.unitNetUsd} onChange={e => upd(item.id, 'unitNetUsd', parseFloat(e.target.value) || 0)} className="w-20 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1"><input type="number" step="0.01" value={item.unitErpUsd} onChange={e => upd(item.id, 'unitErpUsd', parseFloat(e.target.value) || 0)} className="w-20 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1 text-right bg-gray-50 font-mono text-gray-600">{fmtPct(item.defaultMarkup, 2)}</td>
                        <td className="px-2 py-1"><input type="number" step="any" value={Math.round(item.msDiscountPct * 10000) / 100} onChange={e => upd(item.id, 'msDiscountPct', (parseFloat(e.target.value) || 0) / 100)} className="w-16 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1"><input type="number" step="any" value={Math.round(item.crayonMarkupPct * 10000) / 100} onChange={e => upd(item.id, 'crayonMarkupPct', (parseFloat(e.target.value) || 0) / 100)} className="w-16 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1"><input type="number" value={item.unitType} onChange={e => upd(item.id, 'unitType', parseInt(e.target.value) || 1)} className="w-12 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1 text-right bg-blue-50/50 font-mono text-blue-800">{fmtNum(item.discountedNet, 2)}</td>
                        <td className="px-2 py-1 text-right bg-blue-50/50 font-mono text-blue-800">{fmtNum(item.discountedErp, 2)}</td>
                        <td className="px-2 py-1 text-right bg-blue-100/50 font-mono font-medium">{fmt(item.totalNet, header.currencyCode)}</td>
                        <td className="px-2 py-1 text-right bg-blue-100/50 font-mono">{fmt(item.totalErp, header.currencyCode)}</td>
                        <td className="px-2 py-1"><input type="number" value={item.quantity} onChange={e => upd(item.id, 'quantity', parseInt(e.target.value) || 0)} className="w-14 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1 text-right bg-green-50 font-mono text-green-700">{fmtNum(item.eupUnit, 2)}</td>
                        <td className="px-2 py-1 text-right bg-green-100 font-mono font-semibold text-green-800">{fmt(item.totalEup, header.currencyCode)}</td>
                        <td className="px-2 py-1 text-right bg-gray-50 font-mono text-gray-600">{fmtPct(item.calculatedMarkup)}</td>
                        <td className="px-2 py-1"><input type="number" step="any" value={Math.round(item.rebatePct * 10000) / 100} onChange={e => upd(item.id, 'rebatePct', (parseFloat(e.target.value) || 0) / 100)} className="w-14 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                        <td className="px-2 py-1 text-right bg-yellow-50 font-mono text-yellow-700">{fmt(item.rebateAmount, header.currencyCode)}</td>
                        {header.region === 'AF' && (
                          <>
                            <td className="px-2 py-1 text-right bg-purple-50 font-mono text-purple-700">{fmt(item.gp, header.currencyCode)}</td>
                            <td className="px-2 py-1 bg-purple-100"><input type="number" step="any" value={Math.round(item.swoGpPct * 100)} onChange={e => upd(item.id, 'swoGpPct', (parseFloat(e.target.value) || 0) / 100)} className="w-14 px-1 py-1 border rounded text-xs text-right" disabled={!ed} /></td>
                            <td className="px-2 py-1 text-right bg-purple-50 font-mono text-purple-700">{fmt(item.swoGp, header.currencyCode)}</td>
                            <td className="px-2 py-1 text-right bg-orange-50 font-mono text-orange-700">{fmt(item.partnerGp, header.currencyCode)}</td>
                          </>
                        )}
                        <td className="px-2 py-1">{ed && <button onClick={() => del(item.id)} className="p-1 text-red-400 hover:text-red-600 rounded"><Trash2 size={14} /></button>}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-gray-100 font-semibold">
                      <td colSpan={11} className="px-2 py-2 text-right">{header.dealType === 'Ramped' ? `Year ${activeYear} Total:` : 'Yearly Installment:'}</td>
                      <td className="px-2 py-2 text-right bg-blue-200">{fmt(header.dealType === 'Ramped' ? summary.net[`y${activeYear}`] : summary.net.y1, header.currencyCode)}</td>
                      <td className="px-2 py-2 text-right bg-blue-200">{fmt(header.dealType === 'Ramped' ? summary.erp[`y${activeYear}`] : summary.erp.y1, header.currencyCode)}</td>
                      <td className="px-2 py-2"></td>
                      <td className="px-2 py-2"></td>
                      <td className="px-2 py-2 text-right bg-green-200">{fmt(header.dealType === 'Ramped' ? summary.eup[`y${activeYear}`] : summary.eup.y1, header.currencyCode)}</td>
                      <td colSpan={2}></td>
                      <td className="px-2 py-2 text-right bg-yellow-100">{fmt(header.dealType === 'Ramped' ? summary.reb[`y${activeYear}`] : summary.reb.y1, header.currencyCode)}</td>
                      {header.region === 'AF' && (
                        <>
                          <td className="px-2 py-2 text-right bg-purple-200">{fmt(header.dealType === 'Ramped' ? summary.gp[`y${activeYear}`] : summary.gp.y1, header.currencyCode)}</td>
                          <td className="px-2 py-2"></td>
                          <td className="px-2 py-2 text-right bg-purple-200">{fmt(header.dealType === 'Ramped' ? summary.swoGp[`y${activeYear}`] : summary.swoGp.y1, header.currencyCode)}</td>
                          <td className="px-2 py-2 text-right bg-orange-200">{fmt(header.dealType === 'Ramped' ? summary.partnerGp[`y${activeYear}`] : summary.partnerGp.y1, header.currencyCode)}</td>
                        </>
                      )}
                      <td></td>
                    </tr>
                    <tr className="bg-blue-100 font-semibold">
                      <td colSpan={11} className="px-2 py-2 text-right">Total over 3 Years:</td>
                      <td className="px-2 py-2 text-right bg-blue-300">{fmt(summary.net.t, header.currencyCode)}</td>
                      <td className="px-2 py-2 text-right bg-blue-300">{fmt(summary.erp.t, header.currencyCode)}</td>
                      <td className="px-2 py-2"></td>
                      <td className="px-2 py-2"></td>
                      <td className="px-2 py-2 text-right bg-green-300">{fmt(summary.eup.t, header.currencyCode)}</td>
                      <td colSpan={2}></td>
                      <td className="px-2 py-2"></td>
                      {header.region === 'AF' && (
                        <>
                          <td className="px-2 py-2 text-right bg-purple-300">{fmt(summary.gp.t, header.currencyCode)}</td>
                          <td className="px-2 py-2"></td>
                          <td className="px-2 py-2 text-right bg-purple-300">{fmt(summary.swoGp.t, header.currencyCode)}</td>
                          <td className="px-2 py-2 text-right bg-orange-300">{fmt(summary.partnerGp.t, header.currencyCode)}</td>
                        </>
                      )}
                      <td></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
              
              {/* Profit Summary Box */}
              <div className="mt-4 flex justify-end">
                <div className="border border-gray-300 rounded-lg overflow-hidden">
                  <table className="text-sm">
                    <tbody>
                      {header.dealType === 'Ramped' ? (
                        <>
                          <tr className="bg-gray-50 border-b">
                            <td className="px-4 py-2 font-medium text-gray-700">Profit /Year 1</td>
                            <td className="px-4 py-2 text-right font-semibold text-green-700">{fmt(summary.eup.y1 - summary.net.y1, header.currencyCode)}</td>
                          </tr>
                          <tr className="bg-gray-50 border-b">
                            <td className="px-4 py-2 font-medium text-gray-700">Profit /Year 2</td>
                            <td className="px-4 py-2 text-right font-semibold text-green-700">{fmt(summary.eup.y2 - summary.net.y2, header.currencyCode)}</td>
                          </tr>
                          <tr className="bg-gray-50 border-b">
                            <td className="px-4 py-2 font-medium text-gray-700">Profit /Year 3</td>
                            <td className="px-4 py-2 text-right font-semibold text-green-700">{fmt(summary.eup.y3 - summary.net.y3, header.currencyCode)}</td>
                          </tr>
                        </>
                      ) : (
                        <tr className="bg-gray-50 border-b">
                          <td className="px-4 py-2 font-medium text-gray-700">Profit /Year</td>
                          <td className="px-4 py-2 text-right font-semibold text-green-700">{fmt(summary.eup.y1 - summary.net.y1, header.currencyCode)}</td>
                        </tr>
                      )}
                      <tr className="bg-gray-100 border-b">
                        <td className="px-4 py-2 font-medium text-gray-700">Profit /3Years</td>
                        <td className="px-4 py-2 text-right font-bold text-green-800">{fmt(summary.eup.t - summary.net.t, header.currencyCode)}</td>
                      </tr>
                      <tr className="bg-green-100 border-b">
                        <td className="px-4 py-2 font-medium text-gray-700">Overall Margin</td>
                        <td className="px-4 py-2 text-right font-bold text-green-800">{fmtPct(summary.eup.t > 0 ? (summary.eup.t - summary.net.t) / summary.eup.t : 0)}</td>
                      </tr>
                      {header.region === 'AF' && (
                        <>
                          {header.dealType === 'Ramped' ? (
                            <>
                              <tr className="bg-purple-50 border-b">
                                <td className="px-4 py-2 font-medium text-purple-700">SWO GP /Year 1</td>
                                <td className="px-4 py-2 text-right font-semibold text-purple-700">{fmt(summary.swoGp.y1, header.currencyCode)}</td>
                              </tr>
                              <tr className="bg-purple-50 border-b">
                                <td className="px-4 py-2 font-medium text-purple-700">SWO GP /Year 2</td>
                                <td className="px-4 py-2 text-right font-semibold text-purple-700">{fmt(summary.swoGp.y2, header.currencyCode)}</td>
                              </tr>
                              <tr className="bg-purple-50 border-b">
                                <td className="px-4 py-2 font-medium text-purple-700">SWO GP /Year 3</td>
                                <td className="px-4 py-2 text-right font-semibold text-purple-700">{fmt(summary.swoGp.y3, header.currencyCode)}</td>
                              </tr>
                            </>
                          ) : (
                            <tr className="bg-purple-50 border-b">
                              <td className="px-4 py-2 font-medium text-purple-700">SWO GP /Year</td>
                              <td className="px-4 py-2 text-right font-semibold text-purple-700">{fmt(summary.swoGp.y1, header.currencyCode)}</td>
                            </tr>
                          )}
                          <tr className="bg-purple-100 border-b">
                            <td className="px-4 py-2 font-medium text-purple-700">SWO GP /3Years</td>
                            <td className="px-4 py-2 text-right font-bold text-purple-800">{fmt(summary.swoGp.t, header.currencyCode)}</td>
                          </tr>
                          {header.dealType === 'Ramped' ? (
                            <>
                              <tr className="bg-orange-50 border-b">
                                <td className="px-4 py-2 font-medium text-orange-700">Partner GP /Year 1</td>
                                <td className="px-4 py-2 text-right font-semibold text-orange-700">{fmt(summary.partnerGp.y1, header.currencyCode)}</td>
                              </tr>
                              <tr className="bg-orange-50 border-b">
                                <td className="px-4 py-2 font-medium text-orange-700">Partner GP /Year 2</td>
                                <td className="px-4 py-2 text-right font-semibold text-orange-700">{fmt(summary.partnerGp.y2, header.currencyCode)}</td>
                              </tr>
                              <tr className="bg-orange-50 border-b">
                                <td className="px-4 py-2 font-medium text-orange-700">Partner GP /Year 3</td>
                                <td className="px-4 py-2 text-right font-semibold text-orange-700">{fmt(summary.partnerGp.y3, header.currencyCode)}</td>
                              </tr>
                            </>
                          ) : (
                            <tr className="bg-orange-50 border-b">
                              <td className="px-4 py-2 font-medium text-orange-700">Partner GP /Year</td>
                              <td className="px-4 py-2 text-right font-semibold text-orange-700">{fmt(summary.partnerGp.y1, header.currencyCode)}</td>
                            </tr>
                          )}
                          <tr className="bg-orange-100">
                            <td className="px-4 py-2 font-medium text-orange-700">Partner GP /3Years</td>
                            <td className="px-4 py-2 text-right font-bold text-orange-800">{fmt(summary.partnerGp.t, header.currencyCode)}</td>
                          </tr>
                        </>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
              
              {ed && (
                <div className="mt-4">
                  <button onClick={add} className="flex items-center gap-1 px-3 py-1.5 text-xs bg-blue-50 text-blue-700 hover:bg-blue-100 rounded-lg border border-blue-200"><Plus size={14} /> Add Row</button>
                </div>
              )}
            </Section>
            
            <Section title={`Crayon Discount/Funding (${header.currencyCode || 'Local Currency'})`} open={false}>
              <div className="grid grid-cols-3 gap-4">
                {['year1', 'year2', 'year3'].map((k, i) => (
                  <div key={k}><label className="block text-sm font-medium text-gray-700 mb-1">Year {i + 1}</label><input type="number" value={discounts[k]} onChange={e => setDiscounts({...discounts, [k]: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                ))}
              </div>
            </Section>
            
            <Section title="Bid Bond & Bank Charges" open={false}>
              <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Bid Bond %</label><input type="number" step="0.01" value={header.bidBondPct * 100} onChange={e => setHeader({...header, bidBondPct: (parseFloat(e.target.value) || 0) / 100})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Bank Charges %</label><input type="number" step="0.01" value={header.bankChargesPct * 100} onChange={e => setHeader({...header, bankChargesPct: (parseFloat(e.target.value) || 0) / 100})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Performance Bond %</label><input type="number" step="0.01" value={header.performanceBondPct * 100} onChange={e => setHeader({...header, performanceBondPct: (parseFloat(e.target.value) || 0) / 100})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Performance Bank Charges %</label><input type="number" step="0.01" value={header.performanceBankChargesPct * 100} onChange={e => setHeader({...header, performanceBankChargesPct: (parseFloat(e.target.value) || 0) / 100})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Tender Cost ({header.currencyCode})</label><input type="number" step="0.01" value={header.tenderCost} onChange={e => setHeader({...header, tenderCost: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg bg-yellow-50" disabled={!ed} /></div>
              </div>
            </Section>
            
            <Section title={`Other LSP Rebate (${header.currencyCode || 'Local Currency'})`} open={false}>
              <div className="grid grid-cols-3 gap-4">
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Rebate Year 1</label><input type="number" value={header.otherLspRebateY1} onChange={e => setHeader({...header, otherLspRebateY1: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Rebate Year 2</label><input type="number" value={header.otherLspRebateY2} onChange={e => setHeader({...header, otherLspRebateY2: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Rebate Year 3</label><input type="number" value={header.otherLspRebateY3} onChange={e => setHeader({...header, otherLspRebateY3: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
              </div>
            </Section>
            
            <Section title={`CIF Products (${header.currencyCode || 'Local Currency'}) - Yearly Value`} open={false}>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <div><label className="block text-sm font-medium text-gray-700 mb-1">M365E5</label><input type="number" value={header.cifM365E5} onChange={e => setHeader({...header, cifM365E5: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">M365E3</label><input type="number" value={header.cifM365E3} onChange={e => setHeader({...header, cifM365E3: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Azure</label><input type="number" value={header.cifAzure} onChange={e => setHeader({...header, cifAzure: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Dynamics365</label><input type="number" value={header.cifDynamics365} onChange={e => setHeader({...header, cifDynamics365: parseFloat(e.target.value) || 0})} className="w-full px-3 py-2 border rounded-lg" disabled={!ed} /></div>
              </div>
            </Section>
          </div>
        )}
        
        {view === 'summary' && (
          <div className="space-y-6">
            {/* Header Info Section */}
            <div className="bg-white rounded-xl shadow border overflow-hidden">
              <div className="bg-gradient-to-r from-teal-600 to-teal-800 text-white p-4">
                <h2 className="text-lg font-bold">Customer & Agreement Summary</h2>
              </div>
              <div className="p-4">
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                  <div><span className="text-gray-500">Customer Name:</span><p className="font-semibold">{header.customerName || '—'}</p></div>
                  <div><span className="text-gray-500">Sales Location:</span><p className="font-semibold">{header.salesLocation || '—'}</p></div>
                  <div><span className="text-gray-500">Account Manager:</span><p className="font-semibold">{header.accountManager || '—'}</p></div>
                  {header.region === 'AF' && <div><span className="text-gray-500">Partner Name:</span><p className="font-semibold text-orange-600">{header.partnerName || '—'}</p></div>}
                  <div><span className="text-gray-500">Agreement:</span><p className="font-semibold">{header.agreementType}</p></div>
                  <div><span className="text-gray-500">New/Renewal:</span><p className="font-semibold">{header.newOrRenewal}</p></div>
                  <div><span className="text-gray-500">Deal Type:</span><p className={`font-semibold ${header.dealType === 'Ramped' ? 'text-purple-600' : ''}`}>{header.dealType}{header.dealType === 'Ramped' && <span className="ml-1 text-xs bg-purple-100 text-purple-700 px-1 py-0.5 rounded">3 Year Pricing</span>}</p></div>
                  <div><span className="text-gray-500">Currency:</span><p className="font-semibold">{header.currencyCode}</p></div>
                  <div><span className="text-gray-500">Exchange Rate:</span><p className="font-semibold">{header.exchangeRate}</p></div>
                  <div><span className="text-gray-500">Agreement Levels:</span><p className="font-semibold">Sys: {header.agreementLevelSystem} | Srv: {header.agreementLevelServer} | App: {header.agreementLevelApplication}</p></div>
                </div>
              </div>
            </div>
            
            {/* Key Metrics */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <MetricCard label="Total Net (3 Years)" value={fmt(summary.net.t, header.currencyCode)} sub="Cost" />
              <MetricCard label="Total ERP (3 Years)" value={fmt(summary.erp.t, header.currencyCode)} sub="Retail" color="blue" />
              <MetricCard label="Total EUP (3 Years)" value={fmt(summary.eup.t, header.currencyCode)} sub="Without Discount" color="blue" />
              <MetricCard label="GP without Rebates" value={fmt(header.region === 'AF' ? (summary.swoGp.t + summary.partnerGp.t) : (summary.ed.t - summary.net.t), header.currencyCode)} sub={header.region === 'AF' ? `Crayon: ${fmt(summary.swoGp.t, header.currencyCode)} | Partner: ${fmt(summary.partnerGp.t, header.currencyCode)}` : `Markup: ${fmtPct(summary.net.t > 0 ? (summary.ed.t - summary.net.t) / summary.net.t : 0)}`} color="green" big />
            </div>
            
            {/* Cost Price / CPS Price */}
            <Section title="Cost Price / CPS Price">
              <table className="w-full text-sm">
                <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                <tbody className="divide-y">
                  <tr><td className="px-4 py-2">Total Net Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.net.y1, header.currencyCode)}</td></tr>
                  <tr><td className="px-4 py-2">Total Net Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.net.y2, header.currencyCode)}</td></tr>
                  <tr><td className="px-4 py-2">Total Net Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.net.y3, header.currencyCode)}</td></tr>
                  <tr className="bg-blue-50 font-semibold"><td className="px-4 py-2">Grand Total Net Over 3 Years</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.net.t, header.currencyCode)}</td></tr>
                </tbody>
              </table>
            </Section>
            
            {/* Estimated Retail Price */}
            <Section title="Estimated Retail Price">
              <table className="w-full text-sm">
                <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th><th className="px-4 py-2 text-right font-semibold">Default Markup %</th><th className="px-4 py-2 text-right font-semibold">Default GP</th></tr></thead>
                <tbody className="divide-y">
                  <tr><td className="px-4 py-2">Total ERP Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right">{fmtPct(summary.net.y1 > 0 ? (summary.erp.y1 - summary.net.y1) / summary.net.y1 : 0)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y1 - summary.net.y1, header.currencyCode)}</td></tr>
                  <tr><td className="px-4 py-2">Total ERP Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y2, header.currencyCode)}</td><td className="px-4 py-2 text-right">{fmtPct(summary.net.y2 > 0 ? (summary.erp.y2 - summary.net.y2) / summary.net.y2 : 0)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y2 - summary.net.y2, header.currencyCode)}</td></tr>
                  <tr><td className="px-4 py-2">Total ERP Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y3, header.currencyCode)}</td><td className="px-4 py-2 text-right">{fmtPct(summary.net.y3 > 0 ? (summary.erp.y3 - summary.net.y3) / summary.net.y3 : 0)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.y3 - summary.net.y3, header.currencyCode)}</td></tr>
                  <tr className="bg-blue-50 font-semibold"><td className="px-4 py-2">Grand Total ERP Over 3 Years</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.t, header.currencyCode)}</td><td className="px-4 py-2 text-right">{fmtPct(summary.net.t > 0 ? (summary.erp.t - summary.net.t) / summary.net.t : 0)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.erp.t - summary.net.t, header.currencyCode)}</td></tr>
                </tbody>
              </table>
            </Section>
            
            {/* End User Price Section */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* EUP without Discount */}
              <Section title="End User Price without Crayon Discount">
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    <tr><td className="px-4 py-2">Total EUP Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.eup.y1, header.currencyCode)}</td></tr>
                    <tr><td className="px-4 py-2">Total EUP Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.eup.y2, header.currencyCode)}</td></tr>
                    <tr><td className="px-4 py-2">Total EUP Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.eup.y3, header.currencyCode)}</td></tr>
                    <tr className="bg-blue-50 font-semibold"><td className="px-4 py-2">Grand Total EUP (3 Years) w/o Discount</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.eup.t, header.currencyCode)}</td></tr>
                  </tbody>
                </table>
              </Section>
              
              {/* Crayon Discount */}
              <Section title={`Crayon Discount/Funding from Group (${header.currencyCode})`}>
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    <tr className="bg-orange-50"><td className="px-4 py-2 text-orange-700">Discount Value Year 1</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.disc.y1, header.currencyCode)}</td></tr>
                    <tr className="bg-orange-50"><td className="px-4 py-2 text-orange-700">Discount Value Year 2</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.disc.y2, header.currencyCode)}</td></tr>
                    <tr className="bg-orange-50"><td className="px-4 py-2 text-orange-700">Discount Value Year 3</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.disc.y3, header.currencyCode)}</td></tr>
                    <tr className="bg-orange-100 font-semibold"><td className="px-4 py-2 text-orange-800">Total Discount</td><td className="px-4 py-2 text-right font-mono text-orange-800">{fmt(summary.disc.t, header.currencyCode)}</td></tr>
                  </tbody>
                </table>
              </Section>
            </div>
            
            {/* EUP with Discount & Rebates */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* EUP with Discount */}
              <Section title="End User Price with Crayon Discount">
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    <tr><td className="px-4 py-2">Total EUP Year 1 with Discount</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.ed.y1, header.currencyCode)}</td></tr>
                    <tr><td className="px-4 py-2">Total EUP Year 2 with Discount</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.ed.y2, header.currencyCode)}</td></tr>
                    <tr><td className="px-4 py-2">Total EUP Year 3 with Discount</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.ed.y3, header.currencyCode)}</td></tr>
                    <tr className="bg-blue-50 font-semibold"><td className="px-4 py-2">Grand Total EUP (3 Years) w/ Discount</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.ed.t, header.currencyCode)}</td></tr>
                    <tr className="bg-blue-100 font-bold"><td className="px-4 py-2">Grand Total EUP (3 Years) w/ Discount + VAT</td><td className="px-4 py-2 text-right font-mono text-blue-700">{fmt(summary.edVat, header.currencyCode)}</td></tr>
                  </tbody>
                </table>
              </Section>
              
              {/* Crayon Rebate */}
              <Section title="Crayon Rebate">
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    <tr><td className="px-4 py-2">Rebate Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.reb.y1, header.currencyCode)}</td></tr>
                    <tr><td className="px-4 py-2">Rebate Year 2</td><td className="px-4 py-2 text-right font-mono">—</td></tr>
                    <tr><td className="px-4 py-2">Rebate Year 3</td><td className="px-4 py-2 text-right font-mono">—</td></tr>
                    <tr className="bg-yellow-50 font-semibold"><td className="px-4 py-2 text-yellow-700">Total Rebate Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-yellow-700">{fmt(summary.reb.y1, header.currencyCode)}</td></tr>
                  </tbody>
                </table>
              </Section>
              
              {/* Other LSP Rebate */}
              <Section title="Other LSP Rebate">
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    <tr><td className="px-4 py-2">Rebate Year 1</td><td className="px-4 py-2 text-right font-mono">{header.otherLspRebateY1 > 0 ? fmt(header.otherLspRebateY1, header.currencyCode) : '—'}</td></tr>
                    <tr><td className="px-4 py-2">Rebate Year 2</td><td className="px-4 py-2 text-right font-mono">{header.otherLspRebateY2 > 0 ? fmt(header.otherLspRebateY2, header.currencyCode) : '—'}</td></tr>
                    <tr><td className="px-4 py-2">Rebate Year 3</td><td className="px-4 py-2 text-right font-mono">{header.otherLspRebateY3 > 0 ? fmt(header.otherLspRebateY3, header.currencyCode) : '—'}</td></tr>
                    <tr className="bg-gray-100 font-semibold"><td className="px-4 py-2">Total Rebate Over 3 Years</td><td className="px-4 py-2 text-right font-mono">{(header.otherLspRebateY1 + header.otherLspRebateY2 + header.otherLspRebateY3) > 0 ? fmt(header.otherLspRebateY1 + header.otherLspRebateY2 + header.otherLspRebateY3, header.currencyCode) : '—'}</td></tr>
                  </tbody>
                </table>
              </Section>
            </div>
            
            {/* CIF Products */}
            <Section title="CIF Products">
              <table className="w-full text-sm">
                <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold">CIF Products</th><th className="px-4 py-2 text-right font-semibold">Yearly Value</th></tr></thead>
                <tbody className="divide-y">
                  <tr className={header.cifM365E5 > 0 ? 'bg-yellow-50' : ''}><td className="px-4 py-2">M365E5</td><td className="px-4 py-2 text-right font-mono">{header.cifM365E5 > 0 ? fmt(header.cifM365E5, header.currencyCode) : '—'}</td></tr>
                  <tr className={header.cifM365E3 > 0 ? 'bg-yellow-50' : ''}><td className="px-4 py-2">M365E3</td><td className="px-4 py-2 text-right font-mono">{header.cifM365E3 > 0 ? fmt(header.cifM365E3, header.currencyCode) : '—'}</td></tr>
                  <tr className={header.cifAzure > 0 ? 'bg-yellow-50' : ''}><td className="px-4 py-2">Azure</td><td className="px-4 py-2 text-right font-mono">{header.cifAzure > 0 ? fmt(header.cifAzure, header.currencyCode) : '—'}</td></tr>
                  <tr className={header.cifDynamics365 > 0 ? 'bg-yellow-50' : ''}><td className="px-4 py-2">Dynamics365</td><td className="px-4 py-2 text-right font-mono">{header.cifDynamics365 > 0 ? fmt(header.cifDynamics365, header.currencyCode) : '—'}</td></tr>
                </tbody>
              </table>
            </Section>
            
            {/* GP Section */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* GP without Rebates */}
              <Section title="GP without Rebates">
                {header.region === 'AF' ? (
                  <table className="w-full text-sm">
                    <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold text-purple-700">Crayon GP</th><th className="px-4 py-2 text-right font-semibold text-orange-700">Partner GP</th></tr></thead>
                    <tbody className="divide-y">
                      <tr className="bg-gray-50"><td className="px-4 py-2">GP Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.partnerGp.y1, header.currencyCode)}</td></tr>
                      <tr className="bg-gray-50"><td className="px-4 py-2">GP Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y2, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.partnerGp.y2, header.currencyCode)}</td></tr>
                      <tr className="bg-gray-50"><td className="px-4 py-2">GP Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y3, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.partnerGp.y3, header.currencyCode)}</td></tr>
                      <tr className="bg-gray-100 font-semibold"><td className="px-4 py-2">GP Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-purple-700">{fmt(summary.swoGp.t, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.partnerGp.t, header.currencyCode)}</td></tr>
                      <tr><td className="px-4 py-2 font-medium">Markup</td><td className="px-4 py-2 text-right font-bold text-red-600">{fmtPct(summary.net.t > 0 ? summary.swoGp.t / summary.net.t : 0)}</td><td className="px-4 py-2 text-right font-bold text-red-600">{fmtPct(summary.net.t > 0 ? summary.partnerGp.t / summary.net.t : 0)}</td></tr>
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full text-sm">
                    <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                    <tbody className="divide-y">
                      <tr className="bg-green-50"><td className="px-4 py-2 text-green-700">GP Year 1</td><td className="px-4 py-2 text-right font-mono text-green-700">{fmt(summary.ed.y1 - summary.net.y1, header.currencyCode)}</td></tr>
                      <tr className="bg-green-50"><td className="px-4 py-2 text-green-700">GP Year 2</td><td className="px-4 py-2 text-right font-mono text-green-700">{fmt(summary.ed.y2 - summary.net.y2, header.currencyCode)}</td></tr>
                      <tr className="bg-green-50"><td className="px-4 py-2 text-green-700">GP Year 3</td><td className="px-4 py-2 text-right font-mono text-green-700">{fmt(summary.ed.y3 - summary.net.y3, header.currencyCode)}</td></tr>
                      <tr className="bg-green-100 font-semibold"><td className="px-4 py-2 text-green-800">GP Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-green-800">{fmt(summary.ed.t - summary.net.t, header.currencyCode)}</td></tr>
                      <tr><td className="px-4 py-2 font-medium">Markup %</td><td className="px-4 py-2 text-right font-bold text-green-700">{fmtPct(summary.net.t > 0 ? (summary.ed.t - summary.net.t) / summary.net.t : 0)}</td></tr>
                    </tbody>
                  </table>
                )}
              </Section>
              
              {/* Gross Profit with Rebates (matching Excel Merged sheet) */}
              <Section title="Gross Profit with Rebates">
                <table className="w-full text-sm">
                  <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                  <tbody className="divide-y">
                    {header.region === 'AF' ? (
                      <>
                        <tr><td className="px-4 py-2">GP + Rebate Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y1 + summary.reb.y1, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y2 + summary.reb.y2, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.swoGp.y3 + summary.reb.y3, header.currencyCode)}</td></tr>
                        <tr className="bg-green-100 font-semibold"><td className="px-4 py-2 text-green-800">Total GP + Rebate Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-green-800">{fmt(summary.swoGp.t + summary.reb.y1 + summary.reb.y2 + summary.reb.y3, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2 font-medium">Overall Markup</td><td className="px-4 py-2 text-right font-bold text-green-700">{fmtPct(summary.net.t > 0 ? (summary.swoGp.t + summary.reb.y1 + summary.reb.y2 + summary.reb.y3) / summary.net.t : 0)}</td></tr>
                      </>
                    ) : (
                      <>
                        <tr><td className="px-4 py-2">GP + Rebate Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gr.y1, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gr.y2, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gr.y3, header.currencyCode)}</td></tr>
                        <tr className="bg-green-100 font-semibold"><td className="px-4 py-2 text-green-800">Total GP + Rebate Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-green-800">{fmt(summary.gr.t, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2 font-medium">Overall Markup</td><td className="px-4 py-2 text-right font-bold text-green-700">{fmtPct(summary.gr.m)}</td></tr>
                      </>
                    )}
                  </tbody>
                </table>
              </Section>
              
              {/* Gross Profit with Rebates + Crayon Cost */}
              <Section title="Gross Profit with Rebates + Crayon Cost">
                {(() => {
                  // Calculate Crayon Costs
                  // Bid Bond Cost = only Year 1
                  const bidBondCost = summary.edVat * header.bidBondPct * header.bankChargesPct;
                  // Performance Bond Cost = each year
                  const perfBondCostPerYear = summary.edVat * header.performanceBondPct * header.performanceBankChargesPct;
                  // Tender Cost = only Year 1
                  const tenderCost = header.tenderCost || 0;
                  
                  // GP + Rebate values
                  const gpRebY1 = header.region === 'AF' ? summary.swoGp.y1 + summary.reb.y1 : summary.gr.y1;
                  const gpRebY2 = header.region === 'AF' ? summary.swoGp.y2 + summary.reb.y2 : summary.gr.y2;
                  const gpRebY3 = header.region === 'AF' ? summary.swoGp.y3 + summary.reb.y3 : summary.gr.y3;
                  
                  // GP + Rebate + Crayon Cost
                  // Year 1: subtract BB cost + PB cost + Tender cost (Excel formula: =B47-(E42+E47+E52))
                  // Year 2 & 3: subtract only PB cost
                  const gpRebCrayonY1 = gpRebY1 - bidBondCost - perfBondCostPerYear - tenderCost;
                  const gpRebCrayonY2 = gpRebY2 - perfBondCostPerYear;
                  const gpRebCrayonY3 = gpRebY3 - perfBondCostPerYear;
                  const gpRebCrayonTotal = gpRebCrayonY1 + gpRebCrayonY2 + gpRebCrayonY3;
                  const overallMarkup = summary.net.t > 0 ? gpRebCrayonTotal / summary.net.t : 0;
                  
                  return (
                    <table className="w-full text-sm">
                      <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Values</th></tr></thead>
                      <tbody className="divide-y">
                        <tr><td className="px-4 py-2">GP + Rebate + Crayon Cost Year 1</td><td className="px-4 py-2 text-right font-mono">{fmt(gpRebCrayonY1, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate + Crayon Cost Year 2</td><td className="px-4 py-2 text-right font-mono">{fmt(gpRebCrayonY2, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">GP + Rebate + Crayon Cost Year 3</td><td className="px-4 py-2 text-right font-mono">{fmt(gpRebCrayonY3, header.currencyCode)}</td></tr>
                        <tr className="bg-green-200 font-bold"><td className="px-4 py-2 text-green-900">Total GP + Rebate + Crayon Cost Over 3 Years</td><td className="px-4 py-2 text-right font-mono text-green-900 text-lg">{fmt(gpRebCrayonTotal, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2 font-medium">Overall Markup %</td><td className="px-4 py-2 text-right font-bold text-green-700 text-lg">{fmtPct(overallMarkup)}</td></tr>
                      </tbody>
                    </table>
                  );
                })()}
              </Section>
            </div>
            
            {/* Bid Bond & Bank Charges */}
            <Section title="Bid Bond & Bank Charges">
              {(() => {
                const bidBondValue = summary.edVat * header.bidBondPct;
                const totalBBCost = bidBondValue * header.bankChargesPct;
                const perfBondValue = summary.edVat * header.performanceBondPct;
                const perfBondCostPerYear = perfBondValue * header.performanceBankChargesPct;
                const totalPBCost = perfBondCostPerYear * 3;
                const tenderCost = header.tenderCost || 0;
                const totalCrayonCost = totalBBCost + totalPBCost + tenderCost;
                
                return (
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <table className="w-full text-sm">
                      <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold">Bid Bond</th><th className="px-4 py-2 text-right font-semibold">%</th><th className="px-4 py-2 text-right font-semibold">Value</th></tr></thead>
                      <tbody className="divide-y">
                        <tr><td className="px-4 py-2">Bid Bond</td><td className="px-4 py-2 text-right">{fmtPct(header.bidBondPct)}</td><td className="px-4 py-2 text-right font-mono">{fmt(bidBondValue, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">Bank Charges</td><td className="px-4 py-2 text-right">{fmtPct(header.bankChargesPct)}</td><td className="px-4 py-2 text-right font-mono">—</td></tr>
                        <tr className="font-semibold bg-yellow-100"><td className="px-4 py-2">Total BB Cost</td><td className="px-4 py-2 text-right"></td><td className="px-4 py-2 text-right font-mono">{fmt(totalBBCost, header.currencyCode)}</td></tr>
                      </tbody>
                    </table>
                    <table className="w-full text-sm">
                      <thead><tr className="bg-gray-50"><th className="px-4 py-2 text-left font-semibold">Performance Bond</th><th className="px-4 py-2 text-right font-semibold">%</th><th className="px-4 py-2 text-right font-semibold">Value</th></tr></thead>
                      <tbody className="divide-y">
                        <tr><td className="px-4 py-2">Performance Bond</td><td className="px-4 py-2 text-right">{fmtPct(header.performanceBondPct)}</td><td className="px-4 py-2 text-right font-mono">{fmt(perfBondValue, header.currencyCode)}</td></tr>
                        <tr><td className="px-4 py-2">Bank Charges</td><td className="px-4 py-2 text-right">{fmtPct(header.performanceBankChargesPct)}</td><td className="px-4 py-2 text-right font-mono">—</td></tr>
                        <tr className="bg-blue-50"><td className="px-4 py-2">Cost Year 1</td><td className="px-4 py-2 text-right"></td><td className="px-4 py-2 text-right font-mono">{fmt(perfBondCostPerYear, header.currencyCode)}</td></tr>
                        <tr className="bg-blue-50"><td className="px-4 py-2">Cost Year 2</td><td className="px-4 py-2 text-right"></td><td className="px-4 py-2 text-right font-mono">{fmt(perfBondCostPerYear, header.currencyCode)}</td></tr>
                        <tr className="bg-blue-50"><td className="px-4 py-2">Cost Year 3</td><td className="px-4 py-2 text-right"></td><td className="px-4 py-2 text-right font-mono">{fmt(perfBondCostPerYear, header.currencyCode)}</td></tr>
                        <tr className="font-semibold bg-yellow-100"><td className="px-4 py-2">Total PB Cost over 3 years</td><td className="px-4 py-2 text-right"></td><td className="px-4 py-2 text-right font-mono">{fmt(totalPBCost, header.currencyCode)}</td></tr>
                      </tbody>
                    </table>
                  </div>
                );
              })()}
              {/* Tender Cost and Total Crayon Cost */}
              <div className="mt-4 flex justify-end gap-4">
                <table className="text-sm">
                  <tbody>
                    <tr className="bg-yellow-50">
                      <td className="px-6 py-3 font-semibold">Tender Cost</td>
                      <td className="px-6 py-3 text-right font-mono">{header.tenderCost ? fmt(header.tenderCost, header.currencyCode) : '—'}</td>
                    </tr>
                    <tr className="bg-red-100 font-bold">
                      <td className="px-6 py-3 text-red-800">Total Crayon Cost</td>
                      <td className="px-6 py-3 text-right font-mono text-red-800 text-lg">{fmt(summary.edVat * header.bidBondPct * header.bankChargesPct + summary.edVat * header.performanceBondPct * header.performanceBankChargesPct * 3 + (header.tenderCost || 0), header.currencyCode)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </Section>
            
            {/* Africa GP Split - Only show for Africa region */}
            {header.region === 'AF' && (
              <Section title="Africa GP Split (SWO vs Partner)">
                <table className="w-full text-sm">
                  <thead><tr className="bg-purple-50"><th className="px-4 py-2 text-left font-semibold"></th><th className="px-4 py-2 text-right font-semibold">Year 1</th><th className="px-4 py-2 text-right font-semibold">Year 2</th><th className="px-4 py-2 text-right font-semibold">Year 3</th><th className="px-4 py-2 text-right font-semibold bg-purple-100">3-Year Total</th></tr></thead>
                  <tbody className="divide-y">
                    <tr><td className="px-4 py-2 font-medium">Total GP</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono">{fmt(summary.gp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono font-semibold bg-purple-50">{fmt(summary.gp.t, header.currencyCode)}</td></tr>
                    <tr className="bg-purple-50"><td className="px-4 py-2 font-medium text-purple-700">SWO GP</td><td className="px-4 py-2 text-right font-mono text-purple-700">{fmt(summary.swoGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono text-purple-700">{fmt(summary.swoGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono text-purple-700">{fmt(summary.swoGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono font-semibold text-purple-700 bg-purple-100">{fmt(summary.swoGp.t, header.currencyCode)}</td></tr>
                    <tr className="bg-orange-50"><td className="px-4 py-2 font-medium text-orange-700">Partner GP</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.partnerGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.partnerGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono text-orange-700">{fmt(summary.partnerGp.y1, header.currencyCode)}</td><td className="px-4 py-2 text-right font-mono font-semibold text-orange-700 bg-orange-100">{fmt(summary.partnerGp.t, header.currencyCode)}</td></tr>
                  </tbody>
                </table>
              </Section>
            )}
          </div>
        )}
        
        {view === 'preview' && (
          <div className="bg-white rounded-xl shadow-lg border overflow-hidden">
            <div className="bg-gradient-to-r from-blue-600 to-blue-800 text-white p-6">
              <div className="flex justify-between items-start">
                <div>
                  <h2 className="text-2xl font-bold">Price Quotation</h2>
                  <p className="text-blue-100 mt-1">{header.customerName || '(Customer Name)'}</p>
                  <p className="text-blue-200 text-sm">{header.agreementType} | {header.currencyCode} | VAT {fmtPct(header.vatRate)}{header.dealType === 'Ramped' && <span className="ml-2 bg-purple-500 text-white px-2 py-0.5 rounded text-xs">Ramped Deal</span>}</p>
                </div>
                <div className="text-right text-sm">
                  <p className="text-blue-200">Sheet ID: <span className="font-mono text-white">{sheetId}</span></p>
                  <p className="text-blue-200">Exchange Rate: <span className="text-white">{header.exchangeRate}</span></p>
                  {header.accountManager && <p className="text-blue-200">Account Manager: <span className="text-white">{header.accountManager}</span></p>}
                  {header.region === 'AF' && header.partnerName && <p className="text-blue-200">Partner: <span className="text-orange-300">{header.partnerName}</span></p>}
                </div>
              </div>
            </div>
            
            <div className="p-6 overflow-x-auto">
              {/* Normal Deal - Single table with Year columns */}
              {header.dealType !== 'Ramped' && (
                <table className="w-full text-sm border-collapse">
                  <colgroup>
                    <col style={{width: '12%'}} />
                    <col style={{width: '28%'}} />
                    <col style={{width: '6%'}} />
                    <col style={{width: '6%'}} />
                    <col style={{width: '8%'}} />
                    <col style={{width: '10%'}} />
                    <col style={{width: '10%'}} />
                    <col style={{width: '10%'}} />
                    <col style={{width: '10%'}} />
                  </colgroup>
                  
                  {/* Enterprise Online Products */}
                  {sorted.filter(i => i.category === 'ENTERPRISE_ONLINE' && (i.partNumber || i.itemName)).length > 0 && (
                    <>
                      <thead>
                        <tr><th colSpan={9} className="text-left text-sm font-bold text-blue-700 bg-blue-50 px-3 py-2 border border-blue-200">Enterprise Online Products</th></tr>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.1 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.2 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.3 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-blue-100">Total Over 3 Years</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.filter(i => i.category === 'ENTERPRISE_ONLINE' && (i.partNumber || i.itemName)).map(i => (
                          <tr key={i.id} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-blue-50">{fmt(i.totalEup * 3, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </>
                  )}
                  
                  {/* Additional Products */}
                  {sorted.filter(i => i.category === 'ADDITIONAL' && (i.partNumber || i.itemName)).length > 0 && (
                    <>
                      <thead>
                        <tr><th colSpan={9} className="text-left text-sm font-bold text-purple-700 bg-purple-50 px-3 py-2 border border-purple-200 pt-4">Additional Products</th></tr>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.1 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.2 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.3 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-blue-100">Total Over 3 Years</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.filter(i => i.category === 'ADDITIONAL' && (i.partNumber || i.itemName)).map(i => (
                          <tr key={i.id} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-blue-50">{fmt(i.totalEup * 3, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </>
                  )}
                  
                  {/* Additional Products - On Premise */}
                  {sorted.filter(i => i.category === 'ADDITIONAL_ON_PREMISE' && (i.partNumber || i.itemName)).length > 0 && (
                    <>
                      <thead>
                        <tr><th colSpan={9} className="text-left text-sm font-bold text-orange-700 bg-orange-50 px-3 py-2 border border-orange-200 pt-4">Additional Products - On Premise</th></tr>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.1 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.2 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-50">Yr.3 Total</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-blue-100">Total Over 3 Years</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.filter(i => i.category === 'ADDITIONAL_ON_PREMISE' && (i.partNumber || i.itemName)).map(i => (
                          <tr key={i.id} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono border bg-green-50/50">{fmt(i.totalEup, header.currencyCode)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-blue-50">{fmt(i.totalEup * 3, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </>
                  )}
                  
                  {/* Grand Totals - Normal */}
                  <tfoot>
                    <tr className="font-semibold bg-gray-100">
                      <td colSpan={5} className="px-3 py-3 text-right border">Total {header.currencyCode}</td>
                      <td className="px-3 py-3 text-right border bg-green-100">{fmt(summary.eup.y1, header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border bg-green-100">{fmt(summary.eup.y1, header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border bg-green-100">{fmt(summary.eup.y1, header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border bg-blue-200 font-bold">{fmt(summary.eup.t, header.currencyCode)}</td>
                    </tr>
                    <tr className="text-orange-700 bg-orange-50">
                      <td colSpan={5} className="px-3 py-2 text-right border">Further Discount from Crayon</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.disc.y1, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.disc.y2, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.disc.y3, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border font-semibold">{fmt(summary.disc.t, header.currencyCode)}</td>
                    </tr>
                    <tr className="font-semibold bg-gray-50">
                      <td colSpan={5} className="px-3 py-2 text-right border">Total after discount</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y1, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y2, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y3, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border bg-blue-100">{fmt(summary.ed.t, header.currencyCode)}</td>
                    </tr>
                    <tr className="bg-gray-50">
                      <td colSpan={5} className="px-3 py-2 text-right border">VAT {fmtPct(header.vatRate)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y1 * header.vatRate, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y2 * header.vatRate, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border">{fmt(summary.ed.y3 * header.vatRate, header.currencyCode)}</td>
                      <td className="px-3 py-2 text-right border font-semibold">{fmt(summary.ed.t * header.vatRate, header.currencyCode)}</td>
                    </tr>
                    <tr className="bg-teal-600 text-white text-base font-bold">
                      <td colSpan={5} className="px-3 py-3 text-right border border-teal-500">Grand Total with VAT {header.currencyCode}</td>
                      <td className="px-3 py-3 text-right border border-teal-500">{fmt(summary.ed.y1 * (1 + header.vatRate), header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border border-teal-500">{fmt(summary.ed.y2 * (1 + header.vatRate), header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border border-teal-500">{fmt(summary.ed.y3 * (1 + header.vatRate), header.currencyCode)}</td>
                      <td className="px-3 py-3 text-right border border-teal-500 text-lg">{fmt(summary.edVat, header.currencyCode)}</td>
                    </tr>
                  </tfoot>
                </table>
              )}
              
              {/* Ramped Deal - 3 Separate Year Tables */}
              {header.dealType === 'Ramped' && (
                <div className="space-y-8">
                  {/* Year 1 Table */}
                  <div>
                    <h3 className="text-lg font-bold text-purple-700 mb-3 flex items-center gap-2">
                      <span className="bg-purple-600 text-white px-3 py-1 rounded">Year 1</span>
                    </h3>
                    <table className="w-full text-sm border-collapse">
                      <colgroup>
                        <col style={{width: '15%'}} />
                        <col style={{width: '40%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '15%'}} />
                      </colgroup>
                      <thead>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-100">Yr.1 Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.filter(i => i.partNumber || i.itemName).map(i => (
                          <tr key={`y1-${i.id}`} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-green-50">{fmt(i.totalEup, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr className="font-semibold bg-gray-100">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border bg-green-200">{fmt(summary.eup.y1, header.currencyCode)}</td>
                        </tr>
                        <tr className="text-orange-700 bg-orange-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Further Discount from Crayon</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.disc.y1, header.currencyCode)}</td>
                        </tr>
                        <tr className="font-semibold bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total after discount</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y1, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">VAT {fmtPct(header.vatRate)}</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y1 * header.vatRate, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-purple-600 text-white font-bold">
                          <td colSpan={5} className="px-3 py-2 text-right border border-purple-500">Grand Total with VAT {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border border-purple-500 text-lg">{fmt(summary.ed.y1 * (1 + header.vatRate), header.currencyCode)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  
                  {/* Year 2 Table */}
                  <div>
                    <h3 className="text-lg font-bold text-purple-700 mb-3 flex items-center gap-2">
                      <span className="bg-purple-600 text-white px-3 py-1 rounded">Year 2</span>
                    </h3>
                    <table className="w-full text-sm border-collapse">
                      <colgroup>
                        <col style={{width: '15%'}} />
                        <col style={{width: '40%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '15%'}} />
                      </colgroup>
                      <thead>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-100">Yr.2 Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortedY2.filter(i => i.partNumber || i.itemName).map(i => (
                          <tr key={`y2-${i.id}`} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-green-50">{fmt(i.totalEup, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr className="font-semibold bg-gray-100">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border bg-green-200">{fmt(summary.eup.y2, header.currencyCode)}</td>
                        </tr>
                        <tr className="text-orange-700 bg-orange-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Further Discount from Crayon</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.disc.y2, header.currencyCode)}</td>
                        </tr>
                        <tr className="font-semibold bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total after discount</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y2, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">VAT {fmtPct(header.vatRate)}</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y2 * header.vatRate, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-purple-600 text-white font-bold">
                          <td colSpan={5} className="px-3 py-2 text-right border border-purple-500">Grand Total with VAT {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border border-purple-500 text-lg">{fmt(summary.ed.y2 * (1 + header.vatRate), header.currencyCode)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  
                  {/* Year 3 Table */}
                  <div>
                    <h3 className="text-lg font-bold text-purple-700 mb-3 flex items-center gap-2">
                      <span className="bg-purple-600 text-white px-3 py-1 rounded">Year 3</span>
                    </h3>
                    <table className="w-full text-sm border-collapse">
                      <colgroup>
                        <col style={{width: '15%'}} />
                        <col style={{width: '40%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '10%'}} />
                        <col style={{width: '15%'}} />
                      </colgroup>
                      <thead>
                        <tr className="bg-gray-50">
                          <th className="px-3 py-2 text-left font-semibold border">Part Number</th>
                          <th className="px-3 py-2 text-left font-semibold border">Item Name</th>
                          <th className="px-3 py-2 text-right font-semibold border">Qty</th>
                          <th className="px-3 py-2 text-right font-semibold border">Unit Type</th>
                          <th className="px-3 py-2 text-right font-semibold border">EUP</th>
                          <th className="px-3 py-2 text-right font-semibold border bg-green-100">Yr.3 Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortedY3.filter(i => i.partNumber || i.itemName).map(i => (
                          <tr key={`y3-${i.id}`} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-mono text-xs border">{i.partNumber}</td>
                            <td className="px-3 py-2 border">{i.itemName}</td>
                            <td className="px-3 py-2 text-right border">{i.quantity}</td>
                            <td className="px-3 py-2 text-right border">{i.unitType}</td>
                            <td className="px-3 py-2 text-right font-mono border">{fmtNum(i.eupUnit, 2)}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold border bg-green-50">{fmt(i.totalEup, header.currencyCode)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr className="font-semibold bg-gray-100">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border bg-green-200">{fmt(summary.eup.y3, header.currencyCode)}</td>
                        </tr>
                        <tr className="text-orange-700 bg-orange-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Further Discount from Crayon</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.disc.y3, header.currencyCode)}</td>
                        </tr>
                        <tr className="font-semibold bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">Total after discount</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y3, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-gray-50">
                          <td colSpan={5} className="px-3 py-2 text-right border">VAT {fmtPct(header.vatRate)}</td>
                          <td className="px-3 py-2 text-right border">{fmt(summary.ed.y3 * header.vatRate, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-purple-600 text-white font-bold">
                          <td colSpan={5} className="px-3 py-2 text-right border border-purple-500">Grand Total with VAT {header.currencyCode}</td>
                          <td className="px-3 py-2 text-right border border-purple-500 text-lg">{fmt(summary.ed.y3 * (1 + header.vatRate), header.currencyCode)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  
                  {/* Grand Total Summary - All 3 Years */}
                  <div className="mt-8">
                    <h3 className="text-lg font-bold text-teal-700 mb-3 flex items-center gap-2">
                      <span className="bg-teal-600 text-white px-3 py-1 rounded">Grand Total (Year 1 + Year 2 + Year 3)</span>
                    </h3>
                    <table className="w-full text-sm border-collapse">
                      <colgroup>
                        <col style={{width: '70%'}} />
                        <col style={{width: '30%'}} />
                      </colgroup>
                      <tbody>
                        <tr className="font-semibold bg-gray-100">
                          <td className="px-4 py-3 text-right border">Grand Total {header.currencyCode} (Yr1+Yr2+Yr3)</td>
                          <td className="px-4 py-3 text-right border bg-blue-100 font-bold text-lg">{fmt(summary.eup.t, header.currencyCode)}</td>
                        </tr>
                        <tr className="text-orange-700 bg-orange-50">
                          <td className="px-4 py-3 text-right border">Further Discount from Crayon</td>
                          <td className="px-4 py-3 text-right border font-semibold">{fmt(summary.disc.t, header.currencyCode)}</td>
                        </tr>
                        <tr className="font-semibold bg-gray-50">
                          <td className="px-4 py-3 text-right border">Total after discount</td>
                          <td className="px-4 py-3 text-right border bg-blue-50">{fmt(summary.ed.t, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-gray-50">
                          <td className="px-4 py-3 text-right border">VAT {fmtPct(header.vatRate)}</td>
                          <td className="px-4 py-3 text-right border">{fmt(summary.ed.t * header.vatRate, header.currencyCode)}</td>
                        </tr>
                        <tr className="bg-teal-600 text-white font-bold text-lg">
                          <td className="px-4 py-4 text-right border border-teal-500">Grand Total with VAT {header.currencyCode} (Yr1+Yr2+Yr3)</td>
                          <td className="px-4 py-4 text-right border border-teal-500 text-xl">{fmt(summary.edVat, header.currencyCode)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              
              {/* Footer Notes */}
              <div className="mt-6 text-xs text-gray-500 border-t pt-4">
                <p>• All prices are in {header.currencyCode} and subject to {fmtPct(header.vatRate)} VAT</p>
                <p>• This quotation is valid for 30 days from the date of issue</p>
                <p>• Payment terms: As per agreement</p>
                {header.dealType === 'Ramped' && <p>• This is a Ramped Deal with different pricing per year</p>}
              </div>
            </div>
          </div>
        )}
      </main>
      
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t shadow-lg z-20">
        <div className="max-w-full mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3 text-sm">
            <span><span className="font-semibold">{calculated.length}</span> products</span>
            <span className="text-gray-300">|</span>
            <span className="text-blue-600 font-semibold">{fmt(summary.eup.t, header.currencyCode)} EUP</span>
            <span className="text-gray-300">|</span>
            <span className="text-green-600 font-semibold">{fmt(summary.gr.t, header.currencyCode)} GP ({fmtPct(summary.gr.m)})</span>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={exportToExcel} className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 shadow"><FileSpreadsheet size={18} /> Export Excel</button>
            {status === 'draft' && (<><button onClick={handleSave} className="flex items-center gap-2 px-4 py-2 border rounded-lg hover:bg-gray-50"><Save size={18} /> Save</button><button onClick={() => setStatus('submitted')} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 shadow"><Send size={18} /> Submit</button></>)}
            {status === 'submitted' && (<><button onClick={() => setStatus('rejected')} className="flex items-center gap-2 px-4 py-2 border border-red-300 text-red-600 rounded-lg hover:bg-red-50"><XCircle size={18} /> Reject</button><button onClick={() => setStatus('approved')} className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 shadow"><CheckCircle size={18} /> Approve</button></>)}
            {status === 'approved' && <button className="flex items-center gap-2 px-4 py-2 bg-gray-800 text-white rounded-lg hover:bg-gray-900 shadow"><Download size={18} /> Export PDF</button>}
            {status === 'rejected' && <button onClick={() => setStatus('draft')} className="flex items-center gap-1 px-4 py-2 bg-blue-600 text-white rounded-lg shadow"><Edit3 size={18} /> Revise</button>}
          </div>
        </div>
      </div>
    </div>
  );
}
