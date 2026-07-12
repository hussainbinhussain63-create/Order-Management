'use client';

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip as ChartTooltip,
  Legend
} from 'recharts';
import {
  LayoutDashboard,
  UploadCloud,
  FileSpreadsheet,
  Clock,
  AlertTriangle,
  Package,
  XCircle,
  Truck,
  TrendingUp,
  Search,
  Download,
  Trash2,
  Calendar,
  AlertCircle,
  ClipboardCopy,
  Timer,
  Info,
  ChevronRight,
  ClipboardList,
  CheckCircle,
  HelpCircle,
  PackageCheck,
  RefreshCcw,
  BarChart3,
  CalendarDays
} from 'lucide-react';

const viewIcon = (view: string, className = "w-4 h-4") => {
  switch (view) {
    case 'dash': return <LayoutDashboard className={className} />;
    case 'import': return <UploadCloud className={className} />;
    case 'orders': return <FileSpreadsheet className={className} />;
    case 'pending': return <Clock className={className} />;
    case 'wh': return <AlertTriangle className={className} />;
    case 'instock': return <PackageCheck className={className} />;
    case 'stockout': return <XCircle className={className} />;
    case 'courier-ewe': return <Truck className={className} />;
    case 'courier-tfm': return <Truck className={className} />;
    case 'reports': return <TrendingUp className={className} />;
    case 'orderstatus': return <Search className={className} />;
    case 'importlogs': return <ClipboardList className={className} />;
    default: return <Info className={className} />;
  }
};

// Types matching Mongoose Order schema
interface SkuDetail {
  supplier?: string;
  instock?: string;
  supplierStatus?: string;
  supplierDispatchDate?: string;
  receivedInWH?: string;
}

interface OrderRecord {
  _id?: string;
  orderDate: string;
  orderNo: string;
  store: string;
  customer: string;
  mobile: string;
  phone2?: string;
  city: string;
  orderValue: string;
  noItems: string;
  sku: string;
  supplier: string;
  instock: string;
  orderStatus: string;
  supplierStatus: string;
  supplierDispatchDate: string;
  receivedInWH: string;
  dispatchFromOffice: string;
  dispatchDateEWE: string;
  dispatchDateTFM: string;
  courier: string;
  tfmCourier: string;
  courierStatusEWE: string;
  courierStatusTFM: string;
  codStatus: string;
  codAmountReceived: string;
  codReceived?: string;
  returnReceived: string;
  replaced: string;
  trackingEWE: string;
  trackingTFM: string;
  skuDetails?: Record<string, SkuDetail>;
  comment?: string;
  lastComment?: string;
}

// Sidebar View Mappings
const TITLES: Record<string, string> = {
  dash: 'Dashboard',
  import: 'Import data',
  orders: 'All orders',
  pending: 'Pending supplier',
  wh: 'Not dispatched 5+ days',
  instock: 'Instock queue',
  stockout: 'Stock out',
  'courier-ewe': 'Courier — EWE',
  'courier-tfm': 'Courier — TFM',
  reports: 'Reports',
  orderstatus: 'Order status',
  importlogs: 'Import Skipped Logs'
};

// Colors mapping for Hue badges
const hashColor = (str: string): number => {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  return Math.abs(hash) % 360;
};

// Date helpers
const parseOrderDate = (dateStr: string): Date | null => {
  if (!dateStr) return null;
  const clean = dateStr.trim();
  
  // Split by slashes, dashes, or dots
  const parts = clean.split(/[\/\-\.]/);
  if (parts.length === 3) {
    // If year is first (YYYY-MM-DD)
    if (parts[0].length === 4) {
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      const d = parseInt(parts[2], 10);
      const res = new Date(y, m, d);
      if (!isNaN(res.getTime())) return res;
    }
    // If year is last (DD/MM/YYYY or MM/DD/YYYY)
    if (parts[2].length === 4 || parts[2].length === 2) {
      const y = parseInt(parts[2], 10) + (parts[2].length === 2 ? 2000 : 0);
      const p0 = parseInt(parts[0], 10);
      const p1 = parseInt(parts[1], 10);
      
      // If first part is > 12, it must be day (DD/MM/YYYY)
      if (p0 > 12) {
        const res = new Date(y, p1 - 1, p0);
        if (!isNaN(res.getTime())) return res;
      }
      // If second part is > 12, it must be day (MM/DD/YYYY)
      else if (p1 > 12) {
        const res = new Date(y, p0 - 1, p1);
        if (!isNaN(res.getTime())) return res;
      }
      // Default to DD/MM/YYYY
      else {
        const res = new Date(y, p1 - 1, p0);
        if (!isNaN(res.getTime())) return res;
      }
    }
  }
  
  const native = new Date(clean);
  if (!isNaN(native.getTime())) return native;
  
  return null;
};

const orderMonthKey = (dateStr: string): string => {
  const d = parseOrderDate(dateStr);
  if (!d) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
};

const monthLabel = (m: string): string => {
  const parts = m.split('-');
  const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthIdx = parseInt(parts[1], 10) - 1;
  return (names[monthIdx] || '') + ' ' + parts[0];
};

// Text helpers
const isCancelled = (o: OrderRecord): boolean => {
  return (o.orderStatus || '').toLowerCase().includes('cancel');
};

const isDispatchedFromOffice = (o: OrderRecord): boolean => {
  if ((o.orderStatus || '').toLowerCase().includes('dispatch')) return true;
  if (o.courier || o.tfmCourier) return true;
  if (o.courierStatusEWE || o.courierStatusTFM) return true;
  if (o.dispatchDateEWE || o.dispatchDateTFM) return true;
  if ((o.dispatchFromOffice || '').toLowerCase() === 'yes') return true;
  return false;
};

const isDelivered = (o: OrderRecord): boolean => {
  return /deliver/i.test(o.courierStatusEWE || '') || /deliver/i.test(o.courierStatusTFM || '');
};

const isDeliveredEWE = (o: OrderRecord): boolean => /deliver/i.test(o.courierStatusEWE || '');
const isDeliveredTFM = (o: OrderRecord): boolean => /deliver/i.test(o.courierStatusTFM || '');

const isReturned = (o: OrderRecord): boolean => {
  if (o.returnReceived) return true;
  return o.replaced === 'Yes' || /return|rto/i.test(o.courierStatusEWE || '') || /return|rto/i.test(o.courierStatusTFM || '');
};

const skuCode = (s: string): string => {
  return (s || '').split('|')[0].trim().toUpperCase();
};

const COLUMNS = [
  { key: 'orderDate', label: 'Order date' },
  { key: 'orderNo', label: 'Order no.' },
  { key: 'store', label: 'Store' },
  { key: 'customer', label: 'Customer' },
  { key: 'mobile', label: 'Mobile' },
  { key: 'orderValue', label: 'Order value' },
  { key: 'noItems', label: 'No. items' },
  { key: 'sku', label: 'SKU' },
  { key: 'supplier', label: 'Supplier' },
  { key: 'instock', label: 'Instock' },
  { key: 'orderStatus', label: 'Order status' },
  { key: 'supplierStatus', label: 'Supplier status' },
  { key: 'supplierDispatchDate', label: 'Supplier dispatch date' },
  { key: 'dispatchFromOffice', label: 'Dispatch from office' },
  { key: 'dispatchDateEWE', label: 'Dispatch date EWE' },
  { key: 'dispatchDateTFM', label: 'Dispatch date TFM' },
  { key: 'courier', label: 'Courier' },
  { key: 'courierStatusEWE', label: 'Courier status EWE' },
  { key: 'courierStatusTFM', label: 'Courier status TFM' },
  { key: 'codStatus', label: 'COD status' },
  { key: 'codAmountReceived', label: 'COD amt received' },
  { key: 'returnReceived', label: 'Return received' },
  { key: 'trackingEWE', label: 'Tracking EWE' },
  { key: 'trackingTFM', label: 'Tracking TFM' }
];

export default function OrderFlowApp() {
  const [db, setDb] = useState<OrderRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeView, setActiveView] = useState('dash');
  const [monthFilter, setMonthFilter] = useState('');
  const [selectedOrder, setSelectedOrder] = useState<OrderRecord | null>(null);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    orderDate: 90,
    orderNo: 85,
    store: 95,
    customer: 120,
    mobile: 95,
    orderValue: 85,
    noItems: 70,
    sku: 200,
    supplier: 90,
    instock: 80,
    orderStatus: 100,
    supplierStatus: 100,
    supplierDispatchDate: 125,
    dispatchFromOffice: 125,
    dispatchDateEWE: 125,
    dispatchDateTFM: 125,
    courier: 90,
    courierStatusEWE: 125,
    courierStatusTFM: 125,
    codStatus: 90,
    codAmountReceived: 110,
    returnReceived: 100,
    trackingEWE: 120,
    trackingTFM: 120
  });

  const [editingComment, setEditingComment] = useState('');

  const startResize = (colKey: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.pageX;
    const startWidth = colWidths[colKey] || 100;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.pageX - startX;
      const newWidth = Math.max(40, startWidth + deltaX);
      setColWidths(prev => ({
        ...prev,
        [colKey]: newWidth
      }));
    };

    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  };

  useEffect(() => {
    if (selectedOrder) {
      setEditingComment(selectedOrder.comment || '');
    } else {
      setEditingComment('');
    }
  }, [selectedOrder]);

  // Filter States
  const [ordersStore, setOrdersStore] = useState('');
  const [ordersSup, setOrdersSup] = useState('');
  const [ordersSupStatus, setOrdersSupStatus] = useState('');
  const [ordersQuery, setOrdersQuery] = useState('');

  const [pendingStore, setPendingStore] = useState('');
  const [pendingSup, setPendingSup] = useState('');
  const [pendingQuery, setPendingQuery] = useState('');

  const [whStore, setWhStore] = useState('');
  const [whSup, setWhSup] = useState('');
  const [whQuery, setWhQuery] = useState('');

  const [instockStore, setInstockStore] = useState('');
  const [instockQuery, setInstockQuery] = useState('');

  const [stockoutStore, setStockoutStore] = useState('');
  const [stockoutQuery, setStockoutQuery] = useState('');

  const [eweStore, setEweStore] = useState('');
  const [eweStatus, setEweStatus] = useState('');
  const [eweQuery, setEweQuery] = useState('');

  const [tfmStore, setTfmStore] = useState('');
  const [tfmStatus, setTfmStatus] = useState('');
  const [tfmQuery, setTfmQuery] = useState('');

  const [reportsStore, setReportsStore] = useState('');
  const [reportsQuery, setReportsQuery] = useState('');

  const [statusQuery, setStatusQuery] = useState('');

  // Dashboard Filter States
  const [dashStore, setDashStore] = useState('');
  const [dashSup, setDashSup] = useState('');
  const [dashMonth, setDashMonth] = useState('');
  const [dashQuery, setDashQuery] = useState('');

  // Recharts Analytical states
  const [chartData, setChartData] = useState<{
    monthlyTrend: any[];
    storeDistribution: any[];
    supplierDistribution: any[];
    courierStats: any[];
    codSummary: { totalReceived: number; totalPending: number };
  } | null>(null);
  const [chartLoading, setChartLoading] = useState(false);

  const loadChartData = async () => {
    try {
      setChartLoading(true);
      const res = await fetch('/api/reports/dashboard');
      const data = await res.json();
      if (data.success) {
        setChartData(data);
      }
    } catch (err) {
      console.error('Error loading chart aggregates:', err);
    } finally {
      setChartLoading(false);
    }
  };

  // Import Page State
  const [importTab, setImportTab] = useState('at');
  const [importText, setImportText] = useState('');
  const [importFileName, setImportFileName] = useState('');
  const [importResult, setImportResult] = useState<{ ok: boolean; message: string } | null>(null);
  const [previewRows, setPreviewRows] = useState<any[]>([]);
  const [previewHeaders, setPreviewHeaders] = useState<string[]>([]);
  const [forceImport, setForceImport] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Toast State
  const [toast, setToast] = useState<{ message: string; ok: boolean } | null>(null);

  // Custom Delete DB Confirm Modal State
  const [showDeleteModal, setShowDeleteModal] = useState(false);

  // Active Import job tracker
  const [activeJob, setActiveJob] = useState<{
    jobId: string;
    progress: number;
    processedRows: number;
    totalRows: number;
    status: string;
    added: number;
    updated: number;
    error?: string;
  } | null>(null);

  // Import Diagnostic Logs
  const [importLogs, setImportLogs] = useState<any[]>([]);
  const [logsLoading, setLogsLoading] = useState(false);

  const loadImportLogs = async () => {
    try {
      setLogsLoading(true);
      const res = await fetch('/api/import/logs');
      const data = await res.json();
      if (data.success) {
        setImportLogs(data.logs || []);
      }
    } catch (err) {
      console.error('Error loading import logs:', err);
    } finally {
      setLogsLoading(false);
    }
  };

  // Load database on mount
  const loadData = async () => {
    try {
      setLoading(true);
      const res = await fetch('/api/orders');
      const data = await res.json();
      if (data.success) {
        setDb(data.orders || []);
        await loadChartData();
        await loadImportLogs();
      } else {
        showToast(data.error || 'Failed to load shared database', false);
      }
    } catch (err) {
      console.error(err);
      showToast('Could not load database records', false);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  const showToast = (message: string, ok = true) => {
    setToast({ message, ok });
    setTimeout(() => setToast(null), 3000);
  };

  // CSV/TSV Parser Logic from HTML
  const splitDelim = (line: string, sep: string): string[] => {
    const out: string[] = [];
    let cur = '';
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"') {
          if (line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQ = false;
          }
        } else {
          cur += ch;
        }
      } else {
        if (ch === '"') {
          inQ = true;
        } else if (ch === sep) {
          out.push(cur.trim());
          cur = '';
        } else {
          cur += ch;
        }
      }
    }
    out.push(cur.trim());
    return out;
  };

  const norm = (h: string): string => {
    const s = h.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (s.includes('note')) return '';
    if (s.includes('orderstatus')) return 'orderStatus';
    if (s.includes('replaced') || s.includes('replacement')) return 'replaced';
    if (s.includes('lastcomment')) return 'tfmLastComment';
    if (s === 'date' || s.includes('orderdate') || s === 'orderdat') return 'orderDate';
    if ((s.includes('orderno') || s.includes('ordernumber')) && !s.includes('note')) return 'orderNo';
    if (s.includes('customername') || s === 'customer' || s.includes('consignee')) return 'customer';
    if (s.includes('mobilenumber') || s === 'mobile' || s === 'contact') return 'mobile';
    if (s.includes('telephone') || s.includes('contactno') || s.includes('contactnumber')) return 'mobile2';
    if (s.includes('selectstore') || s === 'store') return 'store';
    if (s === 'city' || s.includes('deliverycity') || s.includes('customercity') || s.includes('consigneecity') || s.includes('shippingcity') || s.includes('emirate')) return 'city';
    if (s.includes('totalordervalue') || s.includes('ordervalue') || s === 'totalprice') return 'orderValue';
    if (s.includes('noofitem') || s.includes('numberofitem') || s === 'noofitems' || s.includes('quantity') || s === 'qty' || s === 'qua') return 'noItems';
    if (s === 'sku') return 'sku';
    if (s === 'supplier') return 'supplier';
    if (s === 'instock' || s === 'instockfield') return 'instock';
    if (s.includes('supplierstatus')) return 'supplierStatus';
    if (s.includes('supplierdateofdispatch') || s.includes('dateofdispatch') || s.includes('billno') || s === 'bill_no' || s === 'billno') return 'supplierDispatchDate';
    if (s.includes('receivedinwh') || s.includes('received_in_wh')) return 'receivedInWH';
    if (s.includes('soldout')) return 'soldOut';
    if (s.includes('dispatchfromoffice') || s.includes('dispatchfronoffice')) return 'dispatchFromOffice';
    if (s.includes('dispatchdatefronoffice') || s.includes('dispatchdatefromoffice')) return 'dispatchDateEWE';
    if (s.includes('courierstatus') && s.includes('ewe')) return 'courierStatusEWE';
    if (s.includes('courierstatus') && s.includes('tfm')) return 'courierStatusTFM';
    if (s.includes('courierstatus')) return 'courierStatusEWE';
    if (s.includes('trackingid') && s.includes('ewe')) return 'trackingEWE';
    if (s.includes('trackingid') && s.includes('tfm')) return 'trackingTFM';
    if (s.includes('codstatus')) return 'codStatus';
    if (s.includes('courier') && !s.includes('status') && !s.includes('track')) return 'courier';
    if (s === 'tfm') return 'tfmCourier';
    return s;
  };

  const parseTSV = (text: string) => {
    const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim());
    if (lines.length < 2) return [];
    
    const tabCount = (lines[0].match(/\t/g) || []).length;
    const commaCount = (lines[0].match(/,/g) || []).length;
    const sep = tabCount > commaCount ? '\t' : ',';
    
    const headers = splitDelim(lines[0], sep).map(norm);
    return lines.slice(1).map(l => {
      const vals = splitDelim(l, sep);
      const o: any = {};
      headers.forEach((m, i) => {
        if (m && vals[i] !== undefined && vals[i] !== '') {
          if (m === 'mobile2') {
            if (!o['mobile']) o['mobile'] = vals[i];
            return;
          }
          if (!o[m]) o[m] = vals[i];
        }
      });
      // Store raw fields
      Object.defineProperty(o, '_raw', { value: vals, enumerable: false });
      return o;
    });
  };

  const findValPreview = (r: any, subs: string[]): string => {
    for (const s of subs) { if (r[s] !== undefined && r[s] !== null) return String(r[s]); }
    for (const k of Object.keys(r)) {
      const kl = k.toLowerCase();
      if (subs.some(s => kl.includes(s))) return String(r[k]);
    }
    return '';
  };

  const validatePreviewRow = (row: any, type: string) => {
    if (type === 'orders' || type === 'at') {
      if (!row.orderNo && !row.sku) return 'Row is empty: missing both Order Number and SKU code.';
      if (!row.orderNo) return 'Missing Order Number (orderNo).';
      if (!row.sku) return 'Missing SKU code.';
    } else {
      let ref = '';
      if (type === 'ewe') {
        ref = findValPreview(row, ['referencenumber', 'reference']);
      } else if (type === 'tfm') {
        ref = findValPreview(row, ['shipperref', 'shipper']);
      } else if (type === 'cod' || type === 'returns' || type === 'ret') {
        ref = findValPreview(row, ['orderno', 'order_no', 'orderNo']);
      } else {
        ref = row.orderNo;
      }
      ref = (ref || '').trim();
      if (!ref) {
        return 'Missing Order Reference Number.';
      }

      const exists = db.some(o => o.orderNo === ref);
      if (!exists) {
        if (forceImport) {
          return 'AUTO-CREATE';
        } else {
          return `Order number "${ref}" not found in database (import Airtable orders first)`;
        }
      }
    }
    return null;
  };

  useEffect(() => {
    if (!importText.trim()) {
      setPreviewRows([]);
      setPreviewHeaders([]);
      return;
    }
    try {
      const parsed = parseTSV(importText);
      if (parsed && parsed.length > 0) {
        setPreviewHeaders(Object.keys(parsed[0]));
        setPreviewRows(parsed);
      } else {
        setPreviewRows([]);
        setPreviewHeaders([]);
      }
    } catch (e) {
      setPreviewRows([]);
      setPreviewHeaders([]);
    }
  }, [importText]);

  // Drag and Drop File Handlers
  const handleFile = (file: File) => {
    setImportFileName(file.name);
    const reader = new FileReader();
    const isExcel = /\.xlsx?$/i.test(file.name);
    if (isExcel) {
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const csv = XLSX.utils.sheet_to_csv(worksheet);
          setImportText(csv);
        } catch (err) {
          console.error(err);
          showToast('Could not parse Excel spreadsheet file', false);
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        setImportText(e.target?.result as string || '');
      };
      reader.readAsText(file);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  // Import Submission (Sequential Chunked uploads to prevent Vercel 10s timeouts)
  const handleImportSubmit = async (type: string) => {
    const trimmed = importText.trim();
    if (!trimmed) {
      showToast('Please upload or paste import data first', false);
      return;
    }

    const rows = parseTSV(trimmed);
    if (rows.length === 0) {
      showToast('Could not parse spreadsheet data (no headers/rows)', false);
      return;
    }

    const jobId = 'job_' + Date.now();
    const totalRows = rows.length;
    const chunkSize = 100;

    setActiveJob({
      jobId,
      progress: 0,
      processedRows: 0,
      totalRows,
      status: 'processing',
      added: 0,
      updated: 0
    });

    // Clear inputs and result tab immediately to let the user navigate away
    setImportText('');
    setImportFileName('');
    setImportResult(null);

    // Toast feedback
    showToast('Import started in background');

    // Run chunk loop in async IIFE context (non-blocking)
    (async () => {
      let totalAdded = 0;
      let totalUpdated = 0;
      
      try {
        for (let index = 0; index < totalRows; index += chunkSize) {
          const chunk = rows.slice(index, index + chunkSize);
          
          const res = await fetch('/api/orders/chunk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type, chunk, jobId, startIndex: index, forceImport })
          });
          const data = await res.json();
          
          if (!data.success) {
            throw new Error(data.error || 'Batch import failed');
          }
          
          totalAdded += data.added || 0;
          totalUpdated += data.updated || 0;
          
          const processed = Math.min(totalRows, index + chunk.length);
          const progress = Math.min(100, Math.round((processed / totalRows) * 100));
          
          setActiveJob(prev => prev ? {
            ...prev,
            progress,
            processedRows: processed,
            added: totalAdded,
            updated: totalUpdated
          } : null);
        }
        
        // Completed successfully!
        setActiveJob(prev => prev ? { ...prev, status: 'completed', progress: 100 } : null);
        let msg = '';
        if (type === 'orders' || type === 'at') {
          msg = `Imported ${totalAdded} new orders, updated ${totalUpdated} orders.`;
        } else {
          msg = `Processed ${totalRows} rows successfully (Added: ${totalAdded}, Updated: ${totalUpdated}).`;
        }
        setImportResult({ ok: true, message: msg });
        showToast(msg);
        
        // Load data on finish
        await loadData();
        
        // Auto clear active job indicator after 4 seconds
        setTimeout(() => {
          setActiveJob(null);
        }, 4000);

      } catch (err: any) {
        console.error('Batch import crashed:', err);
        setActiveJob(prev => prev ? { ...prev, status: 'failed', error: err.message } : null);
        setImportResult({ ok: false, message: err.message || 'Import job failed' });
        showToast(err.message || 'Import job failed', false);
      }
    })();
  };

  // Clear Database Execution
  const executeClearDatabase = async () => {
    try {
      setLoading(true);
      const res = await fetch('/api/orders', { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        setDb([]);
        setImportLogs([]);
        showToast('All order database records successfully cleared');
      } else {
        showToast(data.error || 'Failed to clear records', false);
      }
    } catch (err) {
      console.error(err);
      showToast('Network error during delete', false);
    } finally {
      setLoading(false);
      setShowDeleteModal(false);
    }
  };

  // Comment updates
  const handleUpdateComment = async (orderNo: string, sku: string, comment: string) => {
    try {
      const res = await fetch('/api/orders/comment', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orderNo, sku, comment })
      });
      const data = await res.json();
      if (data.success) {
        // Update state in memory
        setDb(prev => prev.map(o => {
          if (o.orderNo === orderNo && o.sku === sku) {
            return { ...o, comment };
          }
          return o;
        }));
        showToast('Comment updated');
      } else {
        showToast(data.error || 'Failed to save comment', false);
      }
    } catch (err) {
      console.error(err);
      showToast('Could not save comment update', false);
    }
  };

  // Months available
  const availableMonths = useMemo(() => {
    const months = db.map(o => orderMonthKey(o.orderDate)).filter(Boolean);
    return [...new Set(months)].sort().reverse();
  }, [db]);

  // Active records inside month filter and not cancelled
  const activeRecords = useMemo(() => {
    return db.filter(o => {
      if (isCancelled(o)) return false;
      if (!monthFilter) return true;
      return orderMonthKey(o.orderDate) === monthFilter;
    });
  }, [db, monthFilter]);

  // Dynamic Dashboard Filtered Records
  const dashFilteredRecords = useMemo(() => {
    return db.filter(o => {
      if (isCancelled(o)) return false;
      if (dashStore && o.store !== dashStore) return false;
      if (dashSup && o.supplier !== dashSup) return false;
      if (dashMonth && orderMonthKey(o.orderDate) !== dashMonth) return false;
      if (dashQuery) {
        const q = dashQuery.toLowerCase();
        return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
      }
      return true;
    });
  }, [db, dashStore, dashSup, dashMonth, dashQuery]);

  // Dashboard Chart Aggregate Computations
  const dashboardCharts = useMemo(() => {
    // 1. Monthly Trend
    const trendMap: Record<string, { month: string; orders: number; value: number }> = {};
    dashFilteredRecords.forEach(o => {
      const dateStr = o.orderDate;
      if (!dateStr) return;
      
      const monthKey = orderMonthKey(dateStr);
      if (!monthKey) return;
      
      const val = parseFloat(o.orderValue) || 0;
      if (!trendMap[monthKey]) {
        trendMap[monthKey] = { month: monthKey, orders: 0, value: 0 };
      }
      trendMap[monthKey].orders++;
      trendMap[monthKey].value += val;
    });
    const monthlyTrend = Object.values(trendMap).sort((a, b) => a.month.localeCompare(b.month));

    // 2. Store Distribution
    const storeMap: Record<string, number> = {};
    dashFilteredRecords.forEach(o => {
      const s = o.store || 'Unknown';
      storeMap[s] = (storeMap[s] || 0) + 1;
    });
    const storeDistribution = Object.entries(storeMap)
      .map(([store, count]) => ({ name: store, value: count }))
      .sort((a, b) => b.value - a.value);

    // 3. Supplier Load
    const supplierMap: Record<string, number> = {};
    dashFilteredRecords.forEach(o => {
      const s = o.supplier || 'Unassigned';
      supplierMap[s] = (supplierMap[s] || 0) + 1;
    });
    const supplierDistribution = Object.entries(supplierMap)
      .map(([supplier, count]) => ({ name: supplier, count }))
      .sort((a, b) => b.count - a.count);

    // 4. Courier Performance
    let eweDelivered = 0, eweReturned = 0, eweHold = 0, ewePending = 0;
    let tfmDelivered = 0, tfmReturned = 0, tfmHold = 0, tfmPending = 0;
    
    dashFilteredRecords.forEach(o => {
      if (o.courier === 'EWE' || o.courierStatusEWE) {
        const s = o.courierStatusEWE || '';
        if (/deliver/i.test(s)) eweDelivered++;
        else if (/return|rto/i.test(s)) eweReturned++;
        else if (/hold/i.test(s)) eweHold++;
        else ewePending++;
      }
      if (o.tfmCourier === 'TFM' || o.courierStatusTFM) {
        const s = o.courierStatusTFM || '';
        if (/deliver/i.test(s)) tfmDelivered++;
        else if (/return|rto/i.test(s)) tfmReturned++;
        else if (/hold/i.test(s)) tfmHold++;
        else tfmPending++;
      }
    });
    const courierStats = [
      { name: 'Delivered', EWE: eweDelivered, TFM: tfmDelivered },
      { name: 'Returned', EWE: eweReturned, TFM: tfmReturned },
      { name: 'On Hold', EWE: eweHold, TFM: tfmHold },
      { name: 'Pending', EWE: ewePending, TFM: tfmPending }
    ];

    // 5. COD Cashflow Summary
    let totalReceived = 0;
    let totalPending = 0;
    dashFilteredRecords.forEach(o => {
      const amt = parseFloat(o.orderValue) || 0;
      if (o.codStatus === 'Received' || o.codReceived === 'Received') {
        totalReceived += amt;
      } else {
        totalPending += amt;
      }
    });
    const codSummary = { totalReceived, totalPending };

    return {
      monthlyTrend,
      storeDistribution,
      supplierDistribution,
      courierStats,
      codSummary
    };
  }, [dashFilteredRecords]);

  // Alert Badge Counts under Dashboard Filters
  const dashAlertCounts = useMemo(() => {
    const today = new Date();
    const daysSince = (d: string) => {
      if (!d) return null;
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return null;
      return Math.max(0, Math.floor((today.getTime() - dt.getTime()) / 86400000));
    };

    const billCreated = dashFilteredRecords.filter(o => {
      const tfmCreated = (o.courierStatusTFM || '').trim().toLowerCase() === 'created';
      const eweCreated = (o.courierStatusEWE || '').trim().toLowerCase() === 'pending';
      return tfmCreated || eweCreated;
    }).length;

    const supNotDispatched2 = dashFilteredRecords.filter(o => {
      if (o.supplierStatus !== 'Not Dispatched') return false;
      if (isDelivered(o)) return false;
      const d = daysSince(o.orderDate);
      return d !== null && d > 2;
    }).length;

    const order5dNotDisp = dashFilteredRecords.filter(o => {
      if (isDispatchedFromOffice(o)) return false;
      const d = daysSince(o.orderDate);
      return d !== null && d > 5;
    }).length;

    const onHoldEWE = dashFilteredRecords.filter(o => {
      if (!(o.courier === 'EWE' || o.courierStatusEWE)) return false;
      if (isDeliveredEWE(o) || isReturned(o)) return false;
      if ((o.courierStatusEWE || '').trim().toLowerCase() === 'pending') return false;
      const replReturnKeys = dashFilteredRecords.filter(x => x.replaced === 'Yes' && (x.returnReceived || '').trim().toLowerCase() !== 'received').map(x => (x.orderNo || '') + '||' + (x.sku || ''));
      return !replReturnKeys.includes((o.orderNo || '') + '||' + (o.sku || ''));
    });

    const onHoldTFM = dashFilteredRecords.filter(o => {
      if (!(o.tfmCourier === 'TFM' || o.courierStatusTFM)) return false;
      if (isDeliveredTFM(o) || isReturned(o)) return false;
      if ((o.courierStatusTFM || '').trim().toLowerCase() === 'pending') return false;
      const replReturnKeys = dashFilteredRecords.filter(x => x.replaced === 'Yes' && (x.returnReceived || '').trim().toLowerCase() !== 'received').map(x => (x.orderNo || '') + '||' + (x.sku || ''));
      return !replReturnKeys.includes((o.orderNo || '') + '||' + (o.sku || ''));
    });

    return {
      billCreated,
      supNotDispatched2,
      order5dNotDisp,
      onHoldEWE: onHoldEWE.length,
      onHoldTFM: onHoldTFM.length,
      onHoldEWERecords: onHoldEWE,
      onHoldTFMRecords: onHoldTFM
    };
  }, [dashFilteredRecords]);

  // Alert Badge Counts
  const alertCounts = useMemo(() => {
    const today = new Date();
    const daysSince = (d: string) => {
      if (!d) return null;
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return null;
      return Math.max(0, Math.floor((today.getTime() - dt.getTime()) / 86400000));
    };

    const billCreated = activeRecords.filter(o => {
      const tfmCreated = (o.courierStatusTFM || '').trim().toLowerCase() === 'created';
      const eweCreated = (o.courierStatusEWE || '').trim().toLowerCase() === 'pending';
      return tfmCreated || eweCreated;
    }).length;

    const supNotDispatched2 = activeRecords.filter(o => {
      if (o.supplierStatus !== 'Not Dispatched') return false;
      if (isDelivered(o)) return false;
      const d = daysSince(o.orderDate);
      return d !== null && d > 2;
    }).length;

    const order5dNotDisp = activeRecords.filter(o => {
      if (isDispatchedFromOffice(o)) return false;
      const d = daysSince(o.orderDate);
      return d !== null && d > 5;
    }).length;

    const onHoldEWE = activeRecords.filter(o => {
      if (!(o.courier === 'EWE' || o.courierStatusEWE)) return false;
      if (isDeliveredEWE(o) || isReturned(o)) return false;
      if ((o.courierStatusEWE || '').trim().toLowerCase() === 'pending') return false;
      const replReturnKeys = activeRecords.filter(x => x.replaced === 'Yes' && (x.returnReceived || '').trim().toLowerCase() !== 'received').map(x => (x.orderNo || '') + '||' + (x.sku || ''));
      return !replReturnKeys.includes((o.orderNo || '') + '||' + (o.sku || ''));
    }).length;

    const onHoldTFM = activeRecords.filter(o => {
      if (!(o.tfmCourier === 'TFM' || o.courierStatusTFM)) return false;
      if (isDeliveredTFM(o) || isReturned(o)) return false;
      if ((o.courierStatusTFM || '').trim().toLowerCase() === 'created') return false;
      const replReturnKeys = activeRecords.filter(x => x.replaced === 'Yes' && (x.returnReceived || '').trim().toLowerCase() !== 'received').map(x => (x.orderNo || '') + '||' + (x.sku || ''));
      return !replReturnKeys.includes((o.orderNo || '') + '||' + (o.sku || ''));
    }).length;

    return { billCreated, supNotDispatched2, order5dNotDisp, onHoldEWE, onHoldTFM };
  }, [activeRecords]);

  // Jump handlers for Sidebar alert badges
  const handleAlertJump = (view: string, elementId: string) => {
    setActiveView(view);
    setTimeout(() => {
      const el = document.getElementById(elementId);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }, 100);
  };

  // CSV Exporter
  const handleExportCSV = () => {
    const headers = [
      'Order Date', 'Order Number', 'Store', 'Customer Name', 'Mobile Number', 'City', 'Order Value', 'No of Items', 'SKU', 'Supplier', 'INSTOCK', 'Order Status', 'Supplier Status', 'Supplier Date of Dispatch', 'Received in WH', 'Dispatch From Office', 'Dispatch Date From Office EWE', 'Dispatch Date From Office TFM', 'Courier', 'TFM', 'Courier Status EWE', 'Courier Status TFM', 'COD Status', 'COD Amount Received', 'Return Received', 'Replaced', 'Tracking ID EWE', 'Tracking ID TFM'
    ];

    const recordsToExport = db.filter(o => {
      if (!monthFilter) return true;
      return orderMonthKey(o.orderDate) === monthFilter;
    });

    const rows = recordsToExport.map(o => [
      o.orderDate, o.orderNo, o.store, o.customer, o.mobile, o.city, o.orderValue, o.noItems, o.sku, o.supplier, o.instock, o.orderStatus, o.supplierStatus, o.supplierDispatchDate, o.receivedInWH, isDispatchedFromOffice(o) ? 'Yes' : 'No', o.dispatchDateEWE, o.dispatchDateTFM, o.courier, o.tfmCourier, o.courierStatusEWE, o.courierStatusTFM, o.codStatus, o.codAmountReceived, o.returnReceived, o.replaced, o.trackingEWE, o.trackingTFM
    ].map(val => `"${(val || '').toString().replace(/"/g, '""')}"`).join(','));

    const csvContent = 'data:text/csv;charset=utf-8,' + encodeURIComponent([headers.join(','), ...rows].join('\n'));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute('href', csvContent);
    downloadAnchor.setAttribute('download', `orders-export${monthFilter ? '-' + monthFilter : ''}.csv`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    document.body.removeChild(downloadAnchor);
  };

  // Shared Helper Badge Render
  const renderBadge = (text: string) => {
    if (!text || text === '-' || text === '#N/A') return <span className="px-2 py-0.5 text-[10px] font-medium bg-zinc-100 text-zinc-500 border border-zinc-200 rounded">R</span>;
    const hue = hashColor(text);
    return (
      <span
        className="px-2 py-0.5 text-[10px] font-medium border rounded whitespace-nowrap inline-block"
        style={{
          backgroundColor: `hsl(${hue}, 65%, 95%)`,
          color: `hsl(${hue}, 60%, 25%)`,
          borderColor: `hsl(${hue}, 50%, 80%)`
        }}
      >
        {text}
      </span>
    );
  };

  const renderStatusBadge = (s: string) => {
    if (!s || s === '-' || s === '#N/A') return <span className="px-2 py-0.5 text-[10px] font-medium bg-zinc-50 text-zinc-400 border border-zinc-200 rounded">—</span>;
    if (s === 'Delivered' || s === 'Received') return <span className="px-2 py-0.5 text-[10px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200 rounded">{s}</span>;
    if (['STOCK OUT', 'Recalled', 'Cancelled', 'Not Dispatched', 'Not Received'].includes(s)) return <span className="px-2 py-0.5 text-[10px] font-medium bg-rose-50 text-rose-700 border border-rose-200 rounded">{s}</span>;
    if (['Dispatched', 'INSTOCK', 'Received in WH'].includes(s)) return <span className="px-2 py-0.5 text-[10px] font-medium bg-blue-50 text-blue-700 border border-blue-200 rounded">{s}</span>;
    if (['In Warehouse', 'Order in the way', 'OFD', 'Pending'].includes(s) || s.toLowerCase().includes('ofd')) return <span className="px-2 py-0.5 text-[10px] font-medium bg-amber-50 text-amber-700 border border-amber-200 rounded">{s}</span>;
    return <span className="px-2 py-0.5 text-[10px] font-medium bg-zinc-100 text-zinc-600 border border-zinc-200 rounded">{s}</span>;
  };

  // Get Order Breakdown (siblings + SKU lists)
  const itemsBreakdown = (o: OrderRecord) => {
    const siblings = db.filter(x => x.orderNo === o.orderNo);
    if (!siblings.length) return null;
    const rows: React.ReactNode[] = [];
    siblings.forEach(item => {
      const parts = (item.sku || '').split(',').map(s => s.trim()).filter(Boolean);
      const list = parts.length ? parts : [item.sku || ''];
      list.forEach((part, idx) => {
        const code = skuCode(part);
        const skuDetailsObj = item.skuDetails ? (item.skuDetails instanceof Map ? Object.fromEntries(item.skuDetails.entries()) : item.skuDetails) : {};
        const hasDetail = !!(skuDetailsObj && Object.prototype.hasOwnProperty.call(skuDetailsObj, code));
        const detail = hasDetail ? skuDetailsObj[code] : {};

        const supplier = hasDetail ? (detail.supplier || item.supplier || '-') : (item.supplier || '-');
        const supplierStatus = hasDetail ? (detail.supplierStatus || '') : (item.supplierStatus || '');
        const date = (hasDetail ? (detail.supplierDispatchDate || '') : (item.supplierDispatchDate || '')).trim();
        const receivedInWH = hasDetail ? (detail.receivedInWH || '') : (item.receivedInWH || '');

        rows.push(
          <tr key={`${item._id || item.orderNo}-${part}-${idx}`} className="hover:bg-zinc-50 border-b border-zinc-100 text-[11px]">
            <td className="p-2 truncate max-w-[200px]" title={part}>{part || '-'}</td>
            <td className="p-2">{renderBadge(supplier)}</td>
            <td className="p-2">
              {date ? (
                <span className="px-2 py-0.5 text-[10px] font-medium bg-blue-50 text-blue-700 border border-blue-200 rounded">Dispatched — {date}</span>
              ) : (
                renderStatusBadge(supplierStatus || 'Not Dispatched')
              )}
            </td>
            <td className="p-2">
              {receivedInWH === 'Yes' ? (
                <span className="px-2 py-0.5 text-[10px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200 rounded">Yes</span>
              ) : (
                <span className="px-2 py-0.5 text-[10px] font-medium bg-zinc-50 text-zinc-400 border border-zinc-200 rounded">No</span>
              )}
            </td>
          </tr>
        );
      });
    });

    return (
      <div className="mt-4 border border-zinc-200 rounded-lg overflow-hidden bg-white shadow-sm">
        <div className="bg-zinc-50 p-2 text-xs font-semibold text-zinc-700 border-b border-zinc-200">
          All items on this order ({rows.length} SKU{rows.length > 1 ? 's' : ''}):
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                <th className="p-2 font-normal">SKU</th>
                <th className="p-2 font-normal">Supplier</th>
                <th className="p-2 font-normal">Supplier status</th>
                <th className="p-2 font-normal">Received in WH</th>
              </tr>
            </thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
      </div>
    );
  };

  const handleRowClick = (o: OrderRecord) => {
    setSelectedOrder(o);
  };

  return (
    <div className="flex h-screen overflow-hidden bg-zinc-100 text-zinc-800 font-sans">
      {/* Toast Notification (Shifted to bottom-left to prevent overlap) */}
      {toast && (
        <div
          className={`fixed bottom-6 left-6 px-4 py-3 border rounded-lg shadow-lg text-sm transition-all duration-300 z-50 flex items-center gap-2 ${
            toast.ok
              ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
              : 'bg-amber-50 text-amber-800 border-amber-200'
          }`}
        >
          <span className={`w-2 h-2 rounded-full ${toast.ok ? 'bg-emerald-500' : 'bg-amber-500'}`}></span>
          {toast.message}
        </div>
      )}

      {/* Floating Background Import Progress Notification Card (Non-blocking) */}
      {activeJob && (
        <div className="fixed bottom-6 right-6 z-40 max-w-md w-80 md:w-96 animate-slide-up">
          <div className="bg-white rounded-2xl border border-indigo-150 p-5 shadow-2xl flex flex-col gap-3.5 hover:shadow-indigo-500/10 transition-all">
            <div className="flex justify-between items-center">
              <h3 className="font-extrabold text-zinc-900 text-xs flex items-center gap-1.5">
                <UploadCloud className="w-4.5 h-4.5 text-indigo-500 animate-bounce" />
                Importing data (background)...
              </h3>
              <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-lg">
                {activeJob.progress}%
              </span>
            </div>

            <div className="w-full bg-zinc-150 rounded-full h-1.5 overflow-hidden border border-zinc-200/30">
              <div 
                className="bg-indigo-600 h-1.5 rounded-full transition-all duration-300 ease-out" 
                style={{ width: `${activeJob.progress}%` }}
              ></div>
            </div>

            <div className="grid grid-cols-2 gap-3 text-[10px] mt-1 border-t border-zinc-100 pt-2.5">
              <div>
                <span className="text-zinc-400 block uppercase font-bold text-[8px] tracking-wider">Processed</span>
                <span className="font-semibold text-zinc-800">{activeJob.processedRows} / {activeJob.totalRows} rows</span>
              </div>
              <div>
                <span className="text-zinc-400 block uppercase font-bold text-[8px] tracking-wider">Status</span>
                <span className="font-semibold text-zinc-800 capitalize flex items-center gap-1">
                  <span className={`w-1.5 h-1.5 rounded-full ${
                    activeJob.status === 'completed' 
                      ? 'bg-emerald-500' 
                      : activeJob.status === 'failed' 
                        ? 'bg-rose-500' 
                        : 'bg-indigo-500 animate-ping'
                  }`}></span>
                  {activeJob.status}
                </span>
              </div>
            </div>
            
            <div className="grid grid-cols-2 gap-3 text-[10px] border-t border-zinc-100 pt-2.5">
              <div>
                <span className="text-zinc-400 block uppercase font-bold text-[8px] tracking-wider">Added (New)</span>
                <span className="font-semibold text-emerald-600">{activeJob.added} records</span>
              </div>
              <div>
                <span className="text-zinc-400 block uppercase font-bold text-[8px] tracking-wider">Updated / Matched</span>
                <span className="font-semibold text-indigo-600">{activeJob.updated} records</span>
              </div>
            </div>

            {activeJob.error && (
              <div className="mt-1.5 text-[9px] text-rose-700 bg-rose-50 border border-rose-100 p-2 rounded-lg font-medium">
                Exception: {activeJob.error}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Custom Delete Database Confirmation Modal */}
      {showDeleteModal && (
        <div className="fixed inset-0 bg-zinc-950/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-zinc-200 p-6 shadow-2xl w-full max-w-md flex flex-col gap-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-2 bg-rose-50 rounded-full border border-rose-100">
                <Trash2 className="w-6 h-6 text-rose-500" />
              </div>
              <h3 className="font-extrabold text-zinc-900 text-sm">Clear All Database Records?</h3>
            </div>

            <div className="text-zinc-600 text-xs leading-relaxed space-y-2">
              <p>
                You are about to clear the database. This action is <strong className="text-rose-600">permanent and irreversible</strong>.
              </p>
              <p className="bg-rose-50 border border-rose-100 rounded-lg p-3 text-[11px] text-rose-700 font-medium flex items-start gap-1.5">
                <AlertTriangle className="w-4.5 h-4.5 text-rose-500 flex-shrink-0 mt-0.5" />
                This will delete all orders, supplier status updates, courier tracking details (EWE/TFM), COD values, returns log, comment logs, and diagnostic records from the database.
              </p>
            </div>

            <div className="grid grid-cols-2 gap-3 mt-4">
              <button
                onClick={() => setShowDeleteModal(false)}
                className="px-4 py-2 border border-zinc-200 hover:bg-zinc-50 text-zinc-700 rounded-lg text-xs font-semibold shadow-sm transition"
              >
                No, Cancel
              </button>
              <button
                onClick={executeClearDatabase}
                disabled={loading}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-semibold shadow-sm shadow-rose-600/10 flex items-center justify-center gap-1.5 transition disabled:opacity-50"
              >
                {loading ? (
                  <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                ) : (
                  <Trash2 className="w-3.5 h-3.5" />
                )}
                Yes, Clear Everything
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Sidebar Navigation */}
      <aside className="w-60 flex-shrink-0 bg-white border-r border-zinc-200 flex flex-col z-20 shadow-xl">
        <div className="p-4 border-b border-zinc-100 flex items-center gap-3 text-zinc-900 font-bold text-lg tracking-tight">
          <Package className="w-6 h-6 text-indigo-600" />
          <span>OrderFlow</span>
        </div>

        <nav className="flex-1 overflow-y-auto py-3 space-y-0.5 px-2">
          {Object.keys(TITLES).map(viewName => {
            const isActive = activeView === viewName;
            return (
              <button
                key={viewName}
                onClick={() => {
                  setActiveView(viewName);
                  setImportResult(null);
                }}
                className={`w-full text-left flex items-center justify-between px-3 py-2 rounded-lg text-xs font-medium transition-all ${
                  isActive
                    ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/20'
                    : 'text-zinc-600 hover:bg-zinc-100 hover:text-zinc-950'
                }`}
              >
                <div className="flex items-center gap-2">
                  {viewIcon(viewName, "w-4 h-4 flex-shrink-0")}
                  <span>{TITLES[viewName]}</span>
                </div>
                {/* Visual arrow indicators for specific menus */}
                <ChevronRight className={`w-3.5 h-3.5 transition-transform ${isActive ? 'rotate-90 text-white' : 'text-zinc-400'}`} />
              </button>
            );
          })}

          {/* Sidebar Alerts panel */}
          <div className="pt-4 mt-4 border-t border-zinc-200 px-2">
            <h4 className="text-[10px] font-semibold text-zinc-400 uppercase tracking-wider mb-2 px-1">Alerts</h4>
            <div className="space-y-1">
              <button
                onClick={() => handleAlertJump('dash', 'sec-dash-bc')}
                className="w-full flex items-center justify-between text-left py-1.5 px-2 rounded hover:bg-zinc-100 text-[11px] text-zinc-600 hover:text-zinc-950"
              >
                <div className="flex items-center gap-1.5">
                  <ClipboardList className="w-3.5 h-3.5 text-amber-500" />
                  <span>Bill created, not dispatched</span>
                </div>
                <span className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full text-[10px] font-bold">{alertCounts.billCreated}</span>
              </button>
              <button
                onClick={() => handleAlertJump('dash', 'sec-dash-s2')}
                className="w-full flex items-center justify-between text-left py-1.5 px-2 rounded hover:bg-zinc-100 text-[11px] text-zinc-600 hover:text-zinc-950"
              >
                <div className="flex items-center gap-1.5">
                  <Timer className="w-3.5 h-3.5 text-red-500" />
                  <span>Suppliers not dispatched 2+ days</span>
                </div>
                <span className="bg-red-100 text-red-800 px-2 py-0.5 rounded-full text-[10px] font-bold">{alertCounts.supNotDispatched2}</span>
              </button>
              <button
                onClick={() => handleAlertJump('dash', 'sec-dash-7d')}
                className="w-full flex items-center justify-between text-left py-1.5 px-2 rounded hover:bg-zinc-100 text-[11px] text-zinc-600 hover:text-zinc-950"
              >
                <div className="flex items-center gap-1.5">
                  <AlertTriangle className="w-3.5 h-3.5 text-red-500" />
                  <span>Order 5+ days, not dispatched</span>
                </div>
                <span className="bg-red-100 text-red-800 px-2 py-0.5 rounded-full text-[10px] font-bold">{alertCounts.order5dNotDisp}</span>
              </button>
              <button
                onClick={() => handleAlertJump('reports', 'sec-rep-holdewe')}
                className="w-full flex items-center justify-between text-left py-1.5 px-2 rounded hover:bg-zinc-800 text-[11px] text-zinc-400 hover:text-white"
              >
                <div className="flex items-center gap-1.5">
                  <Clock className="w-3.5 h-3.5 text-amber-500" />
                  <span>On hold — Easyway</span>
                </div>
                <span className="bg-amber-500/20 text-amber-400 px-2 py-0.5 rounded-full text-[10px] font-bold">{alertCounts.onHoldEWE}</span>
              </button>
              <button
                onClick={() => handleAlertJump('reports', 'sec-rep-holdtfm')}
                className="w-full flex items-center justify-between text-left py-1.5 px-2 rounded hover:bg-zinc-800 text-[11px] text-zinc-400 hover:text-white"
              >
                <div className="flex items-center gap-1.5">
                  <Clock className="w-3.5 h-3.5 text-amber-500" />
                  <span>On hold — TFM</span>
                </div>
                <span className="bg-amber-500/20 text-amber-400 px-2 py-0.5 rounded-full text-[10px] font-bold">{alertCounts.onHoldTFM}</span>
              </button>
            </div>
          </div>
        </nav>
      </aside>

      {/* Main Container */}
      <main className="flex-1 flex flex-col overflow-hidden min-w-0 bg-zinc-50 relative">
        {/* Top Header */}
        <header className="h-14 bg-white border-b border-zinc-200 px-6 flex items-center justify-between z-10 flex-shrink-0">
          <h1 className="font-semibold text-zinc-800 text-base flex items-center gap-2">
            {TITLES[activeView]}
            {loading && (
              <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-zinc-100 text-zinc-500 animate-pulse">
                Syncing MongoDB...
              </span>
            )}
          </h1>

          <div className="flex items-center gap-2.5">
            {/* Global Month Selection Dropdown */}
            <div className="relative flex items-center">
              <CalendarDays className="absolute left-2.5 w-3.5 h-3.5 text-zinc-400 pointer-events-none" />
              <select
                value={monthFilter}
                onChange={e => setMonthFilter(e.target.value)}
                className="pl-8 pr-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs text-zinc-700 shadow-sm focus:outline-none focus:border-indigo-500"
              >
                <option value="">All months</option>
                {availableMonths.map(m => (
                  <option key={m} value={m}>{monthLabel(m)}</option>
                ))}
              </select>
            </div>

            {/* Export CSV */}
            <button
              onClick={handleExportCSV}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 border border-zinc-200 text-zinc-700 rounded-lg text-xs bg-white hover:bg-zinc-50 transition shadow-sm font-medium"
            >
              <Download className="w-3.5 h-3.5 text-zinc-500" />
              Export
            </button>

            {/* Clear Database (Dangerous Action) */}
            <button
              onClick={() => setShowDeleteModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 border border-rose-200 text-rose-700 rounded-lg text-xs bg-white hover:bg-rose-50 transition shadow-sm font-medium"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-500" />
              Clear
            </button>
          </div>
        </header>

        {/* Content Wrapper */}
        <div className="flex-1 overflow-y-auto p-6">
          {loading && db.length === 0 ? (
            <div className="h-full w-full flex flex-col items-center justify-center gap-2 text-zinc-500 text-xs">
              <div className="w-6 h-6 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin"></div>
              <span>Connecting to database & loading orders...</span>
            </div>
          ) : !loading && db.length === 0 && activeView !== 'import' ? (
            <div className="h-[calc(100vh-160px)] w-full flex flex-col items-center justify-center gap-4 text-center p-6 bg-white rounded-2xl border border-zinc-200 shadow-sm max-w-2xl mx-auto my-8">
              <div className="p-4 bg-indigo-50 text-indigo-600 rounded-full">
                <FileSpreadsheet className="w-10 h-10" />
              </div>
              <div className="max-w-sm space-y-1">
                <h3 className="font-bold text-zinc-800 text-base">No Data Added Yet</h3>
                <p className="text-zinc-500 text-xs leading-relaxed">
                  Start by uploading your Airtable orders tracking sheet in the **Import Data** page to populate the dashboard metrics, charts, and tables.
                </p>
              </div>
              <button
                onClick={() => setActiveView('import')}
                className="mt-2 inline-flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs shadow-md shadow-indigo-600/10 font-semibold transition-all"
              >
                <UploadCloud className="w-4 h-4" />
                Go to Import Data
              </button>
            </div>
          ) : (
            <>
              {/* VIEWS SWITCHBOARD */}

              {/* 1. DASHBOARD */}
              {activeView === 'dash' && (
                <div className="space-y-6">
                  {/* Dashboard Dynamic Filter Bar */}
                  <div className="bg-white p-4 rounded-xl border border-zinc-200 shadow-sm flex flex-wrap gap-3 items-center justify-between">
                    <div className="flex flex-wrap gap-3 items-center flex-1">
                      <select
                        value={dashStore}
                        onChange={e => setDashStore(e.target.value)}
                        className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none focus:border-indigo-500"
                      >
                        <option value="">All stores</option>
                        {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                      </select>

                      <select
                        value={dashSup}
                        onChange={e => setDashSup(e.target.value)}
                        className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none focus:border-indigo-500"
                      >
                        <option value="">All suppliers</option>
                        {[...new Set(db.map(o => o.supplier))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                      </select>

                      <select
                        value={dashMonth}
                        onChange={e => setDashMonth(e.target.value)}
                        className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none focus:border-indigo-500"
                      >
                        <option value="">All months</option>
                        {availableMonths.map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}
                      </select>

                      <input
                        type="text"
                        value={dashQuery}
                        onChange={e => setDashQuery(e.target.value)}
                        placeholder="Search order / customer / SKU..."
                        className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none focus:border-indigo-500 flex-1 max-w-xs"
                      />
                    </div>

                    {(dashStore || dashSup || dashMonth || dashQuery) && (
                      <button
                        onClick={() => {
                          setDashStore('');
                          setDashSup('');
                          setDashMonth('');
                          setDashQuery('');
                        }}
                        className="px-3 py-1.5 text-xs text-indigo-600 hover:text-indigo-800 font-semibold"
                      >
                        Clear Filters
                      </button>
                    )}
                  </div>

                  {/* Dashboard Metrics Grid */}
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                    {[
                      { label: 'Total orders', val: [...new Set(dashFilteredRecords.map(o => o.orderNo))].length, style: 'text-indigo-600 bg-indigo-50/50', icon: FileSpreadsheet },
                      { label: 'Total dispatched', val: dashFilteredRecords.filter(isDispatchedFromOffice).length, style: 'text-emerald-600 bg-emerald-50/50', icon: Truck },
                      { label: 'Instock — not dispatched', val: dashFilteredRecords.filter(o => o.supplierStatus === 'INSTOCK' && !isDispatchedFromOffice(o)).length, style: 'text-amber-600 bg-amber-50/50', icon: Package },
                      { label: 'Delivered', val: dashFilteredRecords.filter(isDelivered).length, style: 'text-emerald-600 bg-emerald-50/50', icon: CheckCircle },
                      { label: 'Returned', val: dashFilteredRecords.filter(isReturned).length, style: 'text-rose-600 bg-rose-50/50', icon: RefreshCcw },
                      { label: 'On hold with Easyway', val: dashFilteredRecords.filter(o => (o.courier === 'EWE' || o.courierStatusEWE) && !isDelivered(o) && !isReturned(o)).length, style: 'text-amber-600 bg-amber-50/50', icon: HelpCircle },
                      { label: 'On hold with TFM', val: dashFilteredRecords.filter(o => (o.tfmCourier === 'TFM' || o.courierStatusTFM) && !isDelivered(o) && !isReturned(o)).length, style: 'text-amber-600 bg-amber-50/50', icon: HelpCircle },
                      { label: 'Order taken 5+ days, no dispatch', val: dashAlertCounts.order5dNotDisp, style: 'text-rose-600 bg-rose-50/50 border-rose-200', icon: AlertTriangle },
                      { label: 'Bill created, not dispatched', val: dashAlertCounts.billCreated, style: 'text-amber-600 bg-amber-50/50 border-amber-200', icon: ClipboardList },
                      { label: 'Suppliers not dispatched > 2 days', val: dashAlertCounts.supNotDispatched2, style: 'text-rose-600 bg-rose-50/50 border-rose-200', icon: Timer }
                    ].map((stat, idx) => {
                      const IconComponent = stat.icon;
                      return (
                        <div key={idx} className="bg-white p-4 rounded-xl border border-zinc-200 shadow-sm flex flex-col justify-between relative overflow-hidden group hover:border-zinc-300 transition-all duration-200">
                          <div className="flex justify-between items-start mb-2">
                            <span className="text-[10px] uppercase font-bold tracking-wider text-zinc-400 leading-tight">{stat.label}</span>
                            <IconComponent className="w-4 h-4 text-zinc-400 group-hover:text-indigo-500 transition-colors flex-shrink-0" />
                          </div>
                          <span className={`text-2xl font-bold ${stat.style} px-2 py-0.5 rounded-lg w-fit`}>{stat.val}</span>
                        </div>
                      );
                    })}
                  </div>

                  {/* Dashboard Analytical Graphs */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Chart 1: Volume & Value Trend */}
                    <div className="bg-white p-5 rounded-xl border border-zinc-200 shadow-sm flex flex-col">
                      <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider mb-4 flex items-center gap-1.5">
                        <TrendingUp className="w-4 h-4 text-zinc-500" />
                        Monthly Volume & Value Trend
                      </h3>
                      <div className="h-64 w-full">
                        {dashboardCharts.monthlyTrend.length === 0 ? (
                          <div className="h-full flex items-center justify-center text-zinc-400 text-xs">No trend data available</div>
                        ) : (
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart data={dashboardCharts.monthlyTrend} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                              <defs>
                                <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="#4f46e5" stopOpacity={0.15}/>
                                  <stop offset="95%" stopColor="#4f46e5" stopOpacity={0}/>
                                </linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                              <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                              <YAxis tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                              <ChartTooltip contentStyle={{ fontSize: '11px', borderRadius: '8px' }} />
                              <Legend wrapperStyle={{ fontSize: '10px', marginTop: '10px' }} />
                              <Area name="Order Value (AED)" type="monotone" dataKey="value" stroke="#4f46e5" strokeWidth={2} fillOpacity={1} fill="url(#colorValue)" />
                              <Area name="No. Orders" type="monotone" dataKey="orders" stroke="#06b6d4" strokeWidth={2} fillOpacity={0} />
                            </AreaChart>
                          </ResponsiveContainer>
                        )}
                      </div>
                    </div>

                    {/* Chart 2: Store Wise Distribution */}
                    <div className="bg-white p-5 rounded-xl border border-zinc-200 shadow-sm flex flex-col">
                      <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider mb-4 flex items-center gap-1.5">
                        <PieChart className="w-4 h-4 text-zinc-500" />
                        Orders by Store
                      </h3>
                      <div className="h-64 w-full flex items-center justify-center">
                        {dashboardCharts.storeDistribution.length === 0 ? (
                          <div className="text-zinc-400 text-xs">No store data available</div>
                        ) : (
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie
                                data={dashboardCharts.storeDistribution}
                                cx="50%"
                                cy="45%"
                                innerRadius={55}
                                outerRadius={80}
                                paddingAngle={3}
                                dataKey="value"
                              >
                                {dashboardCharts.storeDistribution.map((entry, index) => {
                                  const colors = ['#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'];
                                  return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />;
                                })}
                              </Pie>
                              <ChartTooltip contentStyle={{ fontSize: '11px', borderRadius: '8px' }} />
                              <Legend layout="horizontal" verticalAlign="bottom" align="center" wrapperStyle={{ fontSize: '10px' }} />
                            </PieChart>
                          </ResponsiveContainer>
                        )}
                      </div>
                    </div>

                    {/* Chart 3: Courier Delivery Performance */}
                    <div className="bg-white p-5 rounded-xl border border-zinc-200 shadow-sm flex flex-col">
                      <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider mb-4 flex items-center gap-1.5">
                        <Truck className="w-4 h-4 text-zinc-500" />
                        Courier Performance Comparison
                      </h3>
                      <div className="h-64 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardCharts.courierStats} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                            <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                            <YAxis tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                            <ChartTooltip contentStyle={{ fontSize: '11px', borderRadius: '8px' }} />
                            <Legend wrapperStyle={{ fontSize: '10px', marginTop: '10px' }} />
                            <Bar name="Easyway (EWE)" dataKey="EWE" fill="#6366f1" radius={[4, 4, 0, 0]} />
                            <Bar name="TFM Courier" dataKey="TFM" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* Chart 4: Supplier Loads */}
                    <div className="bg-white p-5 rounded-xl border border-zinc-200 shadow-sm flex flex-col">
                      <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider mb-4 flex items-center gap-1.5">
                        <ClipboardList className="w-4 h-4 text-zinc-500" />
                        Supplier Load (Top 8)
                      </h3>
                      <div className="h-64 w-full">
                        {dashboardCharts.supplierDistribution.length === 0 ? (
                          <div className="text-zinc-400 text-xs h-full flex items-center justify-center">No supplier data available</div>
                        ) : (
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardCharts.supplierDistribution.slice(0, 8)} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                              <YAxis tickLine={false} axisLine={false} tick={{ fill: '#71717a', fontSize: 10 }} />
                              <ChartTooltip contentStyle={{ fontSize: '11px', borderRadius: '8px' }} />
                              <Bar name="No. Items" dataKey="count" fill="#10b981" radius={[4, 4, 0, 0]} />
                            </BarChart>
                          </ResponsiveContainer>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Table Sections */}

                  {/* Section A: Instock - Not Dispatched */}
                  <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                      <h3 className="font-semibold text-zinc-800 text-xs flex items-center gap-1.5">
                        <PackageCheck className="w-4 h-4 text-zinc-500" />
                        Instock — not dispatched to customer
                      </h3>
                      <span className="bg-amber-100 text-amber-800 font-bold px-2.5 py-0.5 rounded-full text-[10px]">{dashFilteredRecords.filter(o => o.supplierStatus === 'INSTOCK' && !isDispatchedFromOffice(o)).length}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                            <th className="p-3 font-medium">Order date</th>
                            <th className="p-3 font-medium">Order no.</th>
                            <th className="p-3 font-medium">Store</th>
                            <th className="p-3 font-medium">Customer</th>
                            <th className="p-3 font-medium">Mobile</th>
                            <th className="p-3 font-medium">SKU</th>
                            <th className="p-3 font-medium">Order value</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashFilteredRecords.filter(o => o.supplierStatus === 'INSTOCK' && !isDispatchedFromOffice(o)).length === 0 ? (
                            <tr><td colSpan={7} className="p-6 text-center text-zinc-400">No instock items pending</td></tr>
                          ) : (
                            dashFilteredRecords.filter(o => o.supplierStatus === 'INSTOCK' && !isDispatchedFromOffice(o)).map(o => (
                              <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                <td className="p-3">{o.orderDate || '-'}</td>
                                <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                <td className="p-3">{renderBadge(o.store)}</td>
                                <td className="p-3">{o.customer || '-'}</td>
                                <td className="p-3">{o.mobile || '-'}</td>
                                <td className="p-3 truncate max-w-[200px]" title={o.sku}>{o.sku}</td>
                                <td className="p-3 font-medium text-zinc-900">{o.orderValue || '-'}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Section B: Pending from supplier */}
                  <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                      <h3 className="font-semibold text-zinc-800 text-xs flex items-center gap-1.5">
                        <ClipboardList className="w-4 h-4 text-zinc-500" />
                        Pending from supplier
                      </h3>
                      <span className="bg-rose-100 text-rose-800 font-bold px-2.5 py-0.5 rounded-full text-[10px]">{dashFilteredRecords.filter(o => o.supplierStatus === 'Not Dispatched' && !isDelivered(o)).length}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                            <th className="p-3 font-medium">Order date</th>
                            <th className="p-3 font-medium">Order no.</th>
                            <th className="p-3 font-medium">Store</th>
                            <th className="p-3 font-medium">Customer</th>
                            <th className="p-3 font-medium">SKU</th>
                            <th className="p-3 font-medium">Supplier</th>
                            <th className="p-3 font-medium">Supplier status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashFilteredRecords.filter(o => o.supplierStatus === 'Not Dispatched' && !isDelivered(o)).length === 0 ? (
                            <tr><td colSpan={7} className="p-6 text-center text-zinc-400">No pending orders from suppliers</td></tr>
                          ) : (
                            dashFilteredRecords.filter(o => o.supplierStatus === 'Not Dispatched' && !isDelivered(o)).map(o => (
                              <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                <td className="p-3">{o.orderDate || '-'}</td>
                                <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                <td className="p-3">{renderBadge(o.store)}</td>
                                <td className="p-3">{o.customer || '-'}</td>
                                <td className="p-3 truncate max-w-[200px]">{o.sku}</td>
                                <td className="p-3">{renderBadge(o.supplier)}</td>
                                <td className="p-3">{renderStatusBadge(o.supplierStatus)}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Section C: Taken 5+ days, not dispatched */}
                  <div id="sec-dash-7d" className="bg-white border border-rose-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="px-4 py-3 bg-rose-50 border-b border-rose-100 flex justify-between items-center">
                      <h3 className="font-semibold text-rose-950 text-xs flex items-center gap-1.5">
                        <AlertTriangle className="w-4 h-4 text-rose-500" />
                        Order taken 5+ days ago — not yet dispatched from office
                      </h3>
                      <span className="bg-rose-100 text-rose-800 font-bold px-2.5 py-0.5 rounded-full text-[10px]">{dashAlertCounts.order5dNotDisp}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                            <th className="p-3 font-medium">Order date</th>
                            <th className="p-3 font-medium">Order no.</th>
                            <th className="p-3 font-medium">Store</th>
                            <th className="p-3 font-medium">Customer</th>
                            <th className="p-3 font-medium">Mobile</th>
                            <th className="p-3 font-medium">SKU</th>
                            <th className="p-3 font-medium">Supplier</th>
                            <th className="p-3 font-medium">Days pending</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashAlertCounts.order5dNotDisp === 0 ? (
                            <tr><td colSpan={8} className="p-6 text-center text-zinc-400">Nothing pending 5+ days</td></tr>
                          ) : (
                            dashFilteredRecords.filter(o => {
                              if (isDispatchedFromOffice(o)) return false;
                              const dt = new Date(o.orderDate);
                              if (isNaN(dt.getTime())) return false;
                              const days = Math.max(0, Math.floor((new Date().getTime() - dt.getTime()) / 86400000));
                              return days > 5;
                            }).map(o => {
                              const dt = new Date(o.orderDate);
                              const days = Math.floor((new Date().getTime() - dt.getTime()) / 86400000);
                              return (
                                <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                  <td className="p-3">{o.orderDate || '-'}</td>
                                  <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                  <td className="p-3">{renderBadge(o.store)}</td>
                                  <td className="p-3">{o.customer || '-'}</td>
                                  <td className="p-3">{o.mobile || '-'}</td>
                                  <td className="p-3 truncate max-w-[180px]">{o.sku}</td>
                                  <td className="p-3">{renderBadge(o.supplier)}</td>
                                  <td className="p-3"><span className="px-2 py-0.5 bg-rose-100 text-rose-800 rounded font-semibold text-[10px]">{days} days</span></td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Section D: Bill created on courier but not dispatched */}
                  <div id="sec-dash-bc" className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                      <h3 className="font-semibold text-zinc-800 text-xs flex items-center gap-1.5">
                        <ClipboardCopy className="w-4 h-4 text-zinc-500" />
                        Bill created on courier — not yet dispatched
                      </h3>
                      <span className="bg-amber-100 text-amber-800 font-bold px-2.5 py-0.5 rounded-full text-[10px]">{dashAlertCounts.billCreated}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                            <th className="p-3 font-medium">Order no.</th>
                            <th className="p-3 font-medium">Store</th>
                            <th className="p-3 font-medium">Customer</th>
                            <th className="p-3 font-medium">Mobile</th>
                            <th className="p-3 font-medium">SKU</th>
                            <th className="p-3 font-medium">Courier</th>
                            <th className="p-3 font-medium">Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashAlertCounts.billCreated === 0 ? (
                            <tr><td colSpan={7} className="p-6 text-center text-zinc-400">No bill created orders pending dispatch</td></tr>
                          ) : (
                            dashFilteredRecords.filter(o => {
                              const tfmCreated = (o.courierStatusTFM || '').trim().toLowerCase() === 'created';
                              const eweCreated = (o.courierStatusEWE || '').trim().toLowerCase() === 'pending';
                              return tfmCreated || eweCreated;
                            }).map(o => {
                              const isTFM = (o.courierStatusTFM || '').trim().toLowerCase() === 'created';
                              return (
                                <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                  <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                  <td className="p-3">{renderBadge(o.store)}</td>
                                  <td className="p-3">{o.customer || '-'}</td>
                                  <td className="p-3">{o.mobile || '-'}</td>
                                  <td className="p-3 truncate max-w-[180px]">{o.sku}</td>
                                  <td className="p-3">{renderBadge(isTFM ? 'TFM' : 'EWE')}</td>
                                  <td className="p-3">{renderStatusBadge(isTFM ? o.courierStatusTFM : o.courierStatusEWE)}</td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Section E: Suppliers not dispatched > 2 days */}
                  <div id="sec-dash-s2" className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                    <div className="px-4 py-3 bg-rose-50 border-b border-rose-100 flex justify-between items-center">
                      <h3 className="font-semibold text-rose-950 text-xs flex items-center gap-1.5">
                        <Timer className="w-4 h-4 text-rose-500" />
                        Suppliers not dispatched — more than 2 days
                      </h3>
                      <span className="bg-rose-100 text-rose-800 font-bold px-2.5 py-0.5 rounded-full text-[10px]">{dashAlertCounts.supNotDispatched2}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                            <th className="p-3 font-medium">Order date</th>
                            <th className="p-3 font-medium">Order no.</th>
                            <th className="p-3 font-medium">Store</th>
                            <th className="p-3 font-medium">Customer</th>
                            <th className="p-3 font-medium">SKU</th>
                            <th className="p-3 font-medium">Supplier</th>
                            <th className="p-3 font-medium">Days waiting</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashAlertCounts.supNotDispatched2 === 0 ? (
                            <tr><td colSpan={7} className="p-6 text-center text-zinc-400">None</td></tr>
                          ) : (
                            dashFilteredRecords.filter(o => {
                              if (o.supplierStatus !== 'Not Dispatched') return false;
                              if (isDelivered(o)) return false;
                              const dt = new Date(o.orderDate);
                              if (isNaN(dt.getTime())) return false;
                              const days = Math.max(0, Math.floor((new Date().getTime() - dt.getTime()) / 86400000));
                              return days > 2;
                            }).map(o => {
                              const dt = new Date(o.orderDate);
                              const days = Math.floor((new Date().getTime() - dt.getTime()) / 86400000);
                              return (
                                <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                  <td className="p-3">{o.orderDate || '-'}</td>
                                  <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                  <td className="p-3">{renderBadge(o.store)}</td>
                                  <td className="p-3">{o.customer || '-'}</td>
                                  <td className="p-3 truncate max-w-[180px]">{o.sku}</td>
                                  <td className="p-3">{renderBadge(o.supplier)}</td>
                                  <td className="p-3"><span className="px-2 py-0.5 bg-rose-100 text-rose-800 rounded font-semibold text-[10px]">{days} days</span></td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* 2. IMPORT PAGE */}
              {activeView === 'import' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm w-full">
                  {/* Tab Selector */}
                  <div className="flex border-b border-zinc-200 text-xs">
                    {[
                      { id: 'at', label: 'Orders (Airtable)' },
                      { id: 'sup', label: 'Suppliers' },
                      { id: 'ewe', label: 'EWE courier' },
                      { id: 'tfm', label: 'TFM courier' },
                      { id: 'cod', label: 'COD status' },
                      { id: 'ret', label: 'Returns' }
                    ].map(t => (
                      <button
                        key={t.id}
                        onClick={() => {
                          setImportTab(t.id);
                          setImportText('');
                          setImportFileName('');
                          setImportResult(null);
                        }}
                        className={`flex-1 py-3 text-center font-medium border-b-2 transition ${
                          importTab === t.id
                            ? 'border-indigo-600 text-indigo-600 bg-indigo-50/10'
                            : 'border-transparent text-zinc-500 hover:text-zinc-700'
                        }`}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>

                  <div className="p-6 space-y-6">
                    <div className="text-xs text-zinc-600 bg-zinc-50 p-4 rounded-lg border border-zinc-200">
                      {importTab === 'at' && (
                        <p>Upload a CSV/XLS file or paste your Airtable orders export. Columns auto-detected (order date, order no., customer, mobile, SKU, order value, order status, etc). Re-importing updates existing rows by order no. + SKU and preserves any supplier or courier data already matched. If the <strong>order status</strong> column says "Cancelled", the order is marked Cancelled everywhere. Dispatch from office is derived automatically — it's true when order status says dispatched, or when any courier is assigned.</p>
                      )}
                      {importTab === 'sup' && (
                        <p>Upload a CSV file or paste your supplier export. Matched to existing orders by <strong>order no</strong>. If a <strong>SKU</strong> column is present, each row's status is applied only to that SKU (so a multi-SKU order shows a separate status per item); otherwise the status is applied to the whole order. Updates: supplier, instock, supplier status, supplier dispatch date, received in WH, dispatch from office. Import your orders first so there's something to match against.</p>
                      )}
                      {importTab === 'ewe' && (
                        <p>Upload file or paste export from EWE portal. Matched to orders by <strong>Reference Number</strong> = order number. Updates: Courier Status EWE, Tracking ID EWE, Dispatch Date from Office EWE. If a <strong>Replaced</strong> column is present and checked for a row, that order is automatically logged as a return on the same order no.</p>
                      )}
                      {importTab === 'tfm' && (
                        <p>Upload file or paste export from TFM portal. Matched by <strong>SHIPPER REF #</strong> = order number. Updates: Courier Status TFM, Tracking ID TFM, Dispatch Date from Office TFM. If a <strong>Replaced</strong> column is present and checked for a row, that order is automatically logged as a return on the same order no.</p>
                      )}
                      {importTab === 'cod' && (
                        <p>Upload a CSV/XLS file or paste data with columns: <strong>order no</strong>, <strong>amount</strong>, <strong>received</strong>. Matched to existing orders by order no. Sets COD status to "Received" and records the COD amount received when the received column is marked (yes / y / true / 1 / x).</p>
                      )}
                      {importTab === 'ret' && (
                        <p>Upload a CSV/XLS file or paste data with columns: <strong>order no</strong>, <strong>received</strong> (whether the return has been received back or not). Matched to existing orders by order no.</p>
                      )}
                    </div>

                    {/* File Drag and Drop */}
                    <div
                      onDragOver={e => e.preventDefault()}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                      className="border-2 border-dashed border-zinc-300 hover:border-indigo-400 rounded-xl p-8 text-center bg-zinc-50 hover:bg-zinc-100/50 cursor-pointer transition flex flex-col items-center justify-center gap-2 group"
                    >
                      <svg className="w-8 h-8 text-zinc-400 group-hover:text-indigo-500 transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                      </svg>
                      <span className="text-xs text-zinc-500 font-medium">
                        {importFileName ? (
                          <span className="text-indigo-600 font-semibold">{importFileName}</span>
                        ) : (
                          'Click to upload or drag & drop CSV, TSV, TXT, XLS, or XLSX'
                        )}
                      </span>
                      <input
                        type="file"
                        ref={fileInputRef}
                        onChange={e => {
                          const file = e.target.files?.[0];
                          if (file) handleFile(file);
                        }}
                        accept=".csv,.tsv,.txt,.xlsx,.xls"
                        className="hidden"
                      />
                    </div>

                    {/* Raw Text Input */}
                    <div className="space-y-1">
                      <label className="text-[10px] uppercase font-bold text-zinc-400 tracking-wider">Raw Text Paste Area</label>
                      <textarea
                        value={importText}
                        onChange={e => setImportText(e.target.value)}
                        placeholder="— Paste tab or comma separated data directly here —"
                        className="w-full h-32 border border-zinc-200 rounded-xl p-3 text-xs bg-zinc-50 font-mono focus:outline-none focus:border-indigo-500 resize-y"
                      />
                    </div>
                    {/* Submit and Result UI */}
                    <div className="flex flex-col gap-4">
                      {previewRows.length > 0 && (
                        <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-4 text-xs flex flex-col gap-1.5 text-indigo-800">
                          <span className="font-bold flex items-center gap-1.5 text-[13px]">
                            <Info className="w-4 h-4 text-indigo-500" />
                            Ready to Process {previewRows.length} Rows
                          </span>
                          <p className="opacity-90">
                            Clicking the button below will start the background import job. <strong>{previewRows.length}</strong> records will be processed and parsed.
                          </p>
                        </div>
                      )}

                      {previewRows.length > 0 && importTab !== 'at' && importTab !== 'orders' && (
                        <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-900 flex items-center gap-2.5">
                          <input 
                            type="checkbox" 
                            id="chk-force-import"
                            checked={forceImport} 
                            onChange={e => setForceImport(e.target.checked)}
                            className="rounded text-amber-650 focus:ring-amber-500 border-amber-300 w-4 h-4 cursor-pointer"
                          />
                          <label htmlFor="chk-force-import" className="font-semibold cursor-pointer select-none">
                            Force Import: Auto-create missing placeholder orders in the database instead of skipping them.
                          </label>
                        </div>
                      )}

                      {(() => {
                        if (previewRows.length === 0) return null;
                        const invalidRowsCount = previewRows.filter(row => {
                          const err = validatePreviewRow(row, importTab);
                          return err !== null && err !== 'AUTO-CREATE';
                        }).length;
                        if (invalidRowsCount === 0) return null;
                        return (
                          <div className="bg-rose-50 border border-rose-100 rounded-xl p-4 text-xs flex flex-col gap-1.5 text-rose-800 animate-pulse">
                            <span className="font-bold flex items-center gap-1.5 text-[13px] text-rose-700">
                              <AlertTriangle className="w-4 h-4 text-rose-500" />
                              Integrity Alert: {invalidRowsCount} Invalid / Skipped Rows Detected
                            </span>
                            <p className="opacity-90">
                              There are <strong>{invalidRowsCount}</strong> rows in your spreadsheet copy that are missing required keys (such as Order Number or SKU code) and will be bypassed during database ingestion. Flags are highlighted below.
                            </p>
                          </div>
                        );
                      })()}

                      <div className="flex items-center gap-4">
                        <button
                          onClick={() => handleImportSubmit(importTab)}
                          className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-xs font-semibold hover:bg-indigo-700 shadow-md hover:shadow-indigo-600/10 transition"
                        >
                          Import {importTab.toUpperCase()} Data ({previewRows.length} rows)
                        </button>
                        <span className="text-[11px] text-zinc-500">
                          {db.length} total SKUs stored in DB
                        </span>
                      </div>
                    </div>

                    {previewRows.length > 0 && (
                      <div className="border border-zinc-200 rounded-xl overflow-hidden shadow-sm bg-white mt-4">
                        <div className="px-4 py-2.5 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                          <h4 className="font-semibold text-zinc-800 text-[11px]">
                            Parsed Preview ({previewRows.length} rows parsed)
                          </h4>
                          <span className="text-[10px] text-zinc-400 font-medium">Showing first 10 rows for validation</span>
                        </div>
                        <div className="overflow-x-auto max-h-96">
                          <table className="w-full text-left border-collapse text-[11px]">
                            <thead>
                              <tr className="bg-zinc-100 text-[10px] text-zinc-500 uppercase border-b border-zinc-200 font-bold">
                                <th className="p-2 border-r border-zinc-200 text-rose-600 font-bold whitespace-nowrap bg-zinc-150">Pre-Import Validation Alert</th>
                                {previewHeaders.map(h => (
                                  <th key={h} className="p-2 border-r border-zinc-200 whitespace-nowrap">{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {previewRows.slice(0, 10).map((row, idx) => {
                                const validationError = validatePreviewRow(row, importTab);
                                const isAutoCreate = validationError === 'AUTO-CREATE';
                                const isInvalid = !!validationError && !isAutoCreate;

                                return (
                                  <tr 
                                    key={idx} 
                                    className={`border-b border-zinc-100 ${
                                      isInvalid 
                                        ? 'bg-rose-50/70 text-rose-950 font-medium hover:bg-rose-100/50' 
                                        : isAutoCreate 
                                          ? 'bg-amber-50/60 text-amber-950 font-medium hover:bg-amber-100/40 border-amber-100'
                                          : 'hover:bg-zinc-50'
                                    }`}
                                  >
                                    <td className={`p-2 border-r font-bold ${
                                      isInvalid 
                                        ? 'text-rose-600 border-rose-100' 
                                        : isAutoCreate 
                                          ? 'text-amber-600 border-amber-100'
                                          : 'text-zinc-400 border-zinc-100'
                                    }`}>
                                      {isAutoCreate ? '✓ Will Auto-Create Order' : validationError || '✓ Valid'}
                                    </td>
                                    {previewHeaders.map(h => (
                                      <td key={h} className={`p-2 border-r max-w-[200px] truncate ${isInvalid ? 'border-rose-100/60' : isAutoCreate ? 'border-amber-100/60' : 'border-zinc-100'}`} title={row[h]}>
                                        {typeof row[h] === 'object' ? JSON.stringify(row[h]) : String(row[h] ?? '-')}
                                      </td>
                                    ))}
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {importResult && (
                      <div className={`p-4 rounded-xl text-xs border ${
                        importResult.ok
                          ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                          : 'bg-rose-50 text-rose-800 border-rose-200'
                      }`}>
                        {importResult.message}
                      </div>
                    )}

                    {/* Auto Status Logic Helper footer */}
                    {importTab === 'at' && (
                      <div className="pt-6 border-t border-zinc-200 flex flex-wrap gap-2 text-[10px] text-zinc-500">
                        <strong className="text-zinc-700 w-full mb-1">Auto-status logic:</strong>
                        <span className="px-2 py-0.5 bg-indigo-50 border border-indigo-200 rounded text-indigo-800 font-medium">INSTOCK field = Yes / no supplier → Instock, not dispatched</span>
                        <span className="px-2 py-0.5 bg-blue-50 border border-blue-200 rounded text-blue-800 font-medium">Supplier + bill no (dispatch date) → Received in WH</span>
                        <span className="px-2 py-0.5 bg-rose-50 border border-rose-200 rounded text-rose-800 font-medium">Supplier assigned, no bill no → Pending supplier</span>
                        <span className="px-2 py-0.5 bg-rose-50 border border-rose-200 rounded text-rose-800 font-medium">Sold Out → Stock out</span>
                        <span className="px-2 py-0.5 bg-rose-50 border border-rose-200 rounded text-rose-800 font-medium">order status = Cancelled → Cancelled (overrides everything)</span>
                        <span className="px-2 py-0.5 bg-emerald-50 border border-emerald-200 rounded text-emerald-800 font-medium">order status = dispatched, or any courier assigned → Dispatch from office</span>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* 3. ALL ORDERS */}
              {activeView === 'orders' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={ordersStore}
                      onChange={e => setOrdersStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={ordersSup}
                      onChange={e => setOrdersSup(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All suppliers</option>
                      {[...new Set(db.map(o => o.supplier))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={ordersSupStatus}
                      onChange={e => setOrdersSupStatus(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All supplier statuses</option>
                      <option value="Not Dispatched">Not Dispatched</option>
                      <option value="Dispatched">Dispatched</option>
                      <option value="INSTOCK">INSTOCK</option>
                      <option value="STOCK OUT">STOCK OUT</option>
                    </select>

                    <input
                      type="text"
                      value={ordersQuery}
                      onChange={e => setOrdersQuery(e.target.value)}
                      placeholder="Search order no / customer / SKU..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  {/* Main Grid View */}
                  <div className="overflow-x-auto overflow-y-auto max-h-[calc(100vh-220px)] border border-zinc-200 rounded-lg shadow-inner">
                    <table className="text-left border-collapse text-xs" style={{ tableLayout: 'fixed', width: 'max-content' }}>
                      <thead className="sticky top-0 bg-zinc-50 z-20 border-b border-zinc-200 shadow-[0_1px_0_rgba(0,0,0,0.05)]">
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200 font-semibold">
                          {COLUMNS.map(col => (
                            <th
                              key={col.key}
                              style={{ width: colWidths[col.key] || 100, minWidth: colWidths[col.key] || 100, maxWidth: colWidths[col.key] || 100 }}
                              className="relative p-2.5 border-r border-zinc-200 whitespace-nowrap font-semibold select-none group text-left"
                            >
                              <div className="truncate pr-3" title={col.label}>{col.label}</div>
                              {/* Drag handle for resizing */}
                              <div
                                onMouseDown={(e) => startResize(col.key, e)}
                                className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-indigo-500/50 active:bg-indigo-600 z-10 transition-colors"
                              />
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {db.filter(o => {
                          if (!monthFilter) return true;
                          return orderMonthKey(o.orderDate) === monthFilter;
                        }).filter(o => {
                          if (ordersStore && o.store !== ordersStore) return false;
                          if (ordersSup && o.supplier !== ordersSup) return false;
                          if (ordersSupStatus && o.supplierStatus !== ordersSupStatus) return false;
                          if (ordersQuery) {
                            const q = ordersQuery.toLowerCase();
                            return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                          }
                          return true;
                        }).length === 0 ? (
                          <tr><td colSpan={COLUMNS.length} className="p-6 text-center text-zinc-400">No results found</td></tr>
                        ) : (
                          db.filter(o => {
                            if (!monthFilter) return true;
                            return orderMonthKey(o.orderDate) === monthFilter;
                          }).filter(o => {
                            if (ordersStore && o.store !== ordersStore) return false;
                            if (ordersSup && o.supplier !== ordersSup) return false;
                            if (ordersSupStatus && o.supplierStatus !== ordersSupStatus) return false;
                            if (ordersQuery) {
                              const q = ordersQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          }).map(o => (
                            <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              {COLUMNS.map(col => {
                                const val = o[col.key as keyof OrderRecord] ?? '';
                                let cellContent: React.ReactNode = String(val);

                                if (col.key === 'orderNo') {
                                  cellContent = <span className="font-semibold text-indigo-600">{String(val)}</span>;
                                } else if (col.key === 'store' || col.key === 'supplier' || col.key === 'courier') {
                                  cellContent = renderBadge(String(val));
                                } else if (
                                  col.key === 'orderStatus' || 
                                  col.key === 'supplierStatus' || 
                                  col.key === 'courierStatusEWE' || 
                                  col.key === 'courierStatusTFM' || 
                                  col.key === 'returnReceived'
                                ) {
                                  cellContent = renderStatusBadge(String(val));
                                } else if (col.key === 'dispatchFromOffice') {
                                  const isDisp = isDispatchedFromOffice(o);
                                  cellContent = isDisp ? (
                                    <span className="px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded font-semibold text-[10px]">Yes</span>
                                  ) : (
                                    <span className="px-2 py-0.5 bg-zinc-50 text-zinc-400 border border-zinc-200 rounded font-semibold text-[10px]">No</span>
                                  );
                                } else if (col.key === 'trackingEWE' || col.key === 'trackingTFM') {
                                  cellContent = <span className="font-mono text-[10px]">{String(val)}</span>;
                                }

                                return (
                                  <td
                                    key={col.key}
                                    style={{ width: colWidths[col.key] || 100, minWidth: colWidths[col.key] || 100, maxWidth: colWidths[col.key] || 100 }}
                                    className="p-2.5 border-r border-zinc-100 truncate whitespace-nowrap text-left"
                                    title={typeof val === 'string' ? val : ''}
                                  >
                                    {cellContent}
                                  </td>
                                );
                              })}
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 4. PENDING SUPPLIER */}
              {activeView === 'pending' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={pendingStore}
                      onChange={e => setPendingStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={pendingSup}
                      onChange={e => setPendingSup(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All suppliers</option>
                      {[...new Set(db.map(o => o.supplier))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <input
                      type="text"
                      value={pendingQuery}
                      onChange={e => setPendingQuery(e.target.value)}
                      placeholder="Search order no..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order date</th>
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">SKU</th>
                          <th className="p-3 font-medium">Supplier</th>
                          <th className="p-3 font-medium">Supplier status</th>
                          <th className="p-3 font-medium">Days waiting</th>
                          <th className="p-3 font-medium">Comment</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const today = new Date();
                          const candidates = activeRecords.filter(o => {
                            const ss = o.supplierStatus || '';
                            if (ss !== 'Not Dispatched' && !ss.startsWith('Partially dispatched')) return false;
                            if (isCancelled(o)) return false;
                            if (isDelivered(o)) return false;
                            if (pendingStore && o.store !== pendingStore) return false;
                            if (pendingQuery && !(o.orderNo || '').toLowerCase().includes(pendingQuery.toLowerCase().trim())) return false;
                            return true;
                          });

                          // Extract each pending SKU line
                          const lines: any[] = [];
                          candidates.forEach(o => {
                            const parts = (o.sku || '').split(',').map(s => s.trim()).filter(Boolean);
                            const list = parts.length ? parts : [o.sku || ''];
                            list.forEach(part => {
                              const code = skuCode(part);
                              const skuDetailsObj = o.skuDetails ? (o.skuDetails instanceof Map ? Object.fromEntries(o.skuDetails.entries()) : o.skuDetails) : {};
                              const hasDetail = !!(skuDetailsObj && Object.prototype.hasOwnProperty.call(skuDetailsObj, code));
                              const detail = hasDetail ? skuDetailsObj[code] : {};

                              const supplier = hasDetail ? (detail.supplier || o.supplier || '-') : (o.supplier || '-');
                              const supplierStatus = hasDetail ? (detail.supplierStatus || '') : (o.supplierStatus || '');
                              const dispatchDate = (hasDetail ? (detail.supplierDispatchDate || '') : (o.supplierDispatchDate || '')).trim();

                              if (dispatchDate) return;
                              if (supplierStatus && supplierStatus !== 'Not Dispatched' && !supplierStatus.startsWith('Partially')) return;
                              if (pendingSup && supplier !== pendingSup) return;

                              lines.push({ o, part, supplier });
                            });
                          });

                          if (lines.length === 0) {
                            return <tr><td colSpan={10} className="p-6 text-center text-zinc-400">No pending items found</td></tr>;
                          }

                          return lines.map((line, index) => {
                            const days = line.o.orderDate ? Math.max(0, Math.floor((today.getTime() - new Date(line.o.orderDate).getTime()) / 86400000)) : '-';
                            const dc = typeof days === 'number' ? (days > 5 ? 'bg-rose-100 text-rose-800' : days > 2 ? 'bg-amber-100 text-amber-800' : 'bg-zinc-100 text-zinc-600') : 'bg-zinc-100 text-zinc-600';
                            return (
                              <tr key={`${line.o._id || line.o.orderNo}-${line.part}-${index}`} className="hover:bg-zinc-50 border-b border-zinc-100">
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{line.o.orderDate || '-'}</td>
                                <td className="p-3 font-medium text-indigo-600 cursor-pointer" onClick={() => handleRowClick(line.o)}>{line.o.orderNo}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{renderBadge(line.o.store)}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{line.o.customer || '-'}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{line.o.mobile || '-'}</td>
                                <td className="p-3 truncate max-w-[200px]" onClick={() => handleRowClick(line.o)}>{line.part}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{renderBadge(line.supplier)}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>{renderStatusBadge('Not Dispatched')}</td>
                                <td className="p-3" onClick={() => handleRowClick(line.o)}>
                                  <span className={`px-2 py-0.5 rounded font-semibold text-[10px] ${dc}`}>{days} days</span>
                                </td>
                                <td className="p-3">
                                  <input
                                    type="text"
                                    defaultValue={line.o.comment || ''}
                                    placeholder="Add comment..."
                                    onBlur={e => handleUpdateComment(line.o.orderNo, line.o.sku, e.target.value)}
                                    className="px-2 py-1 border border-zinc-200 rounded text-xs bg-white focus:outline-none focus:border-indigo-500 w-full max-w-[150px]"
                                  />
                                </td>
                              </tr>
                            );
                          });
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 5. NOT DISPATCHED 5+ DAYS (WH) */}
              {activeView === 'wh' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={whStore}
                      onChange={e => setWhStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={whSup}
                      onChange={e => setWhSup(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All suppliers</option>
                      {[...new Set(db.map(o => o.supplier))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <input
                      type="text"
                      value={whQuery}
                      onChange={e => setWhQuery(e.target.value)}
                      placeholder="Search order no / customer / SKU..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order date</th>
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">SKU</th>
                          <th className="p-3 font-medium">Supplier</th>
                          <th className="p-3 font-medium">Supplier status</th>
                          <th className="p-3 font-medium">Supplier dispatch date</th>
                          <th className="p-3 font-medium">Order value</th>
                          <th className="p-3 font-medium">Days pending</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const today = new Date();
                          const rows = activeRecords.filter(o => {
                            if (isCancelled(o) || isDelivered(o) || isReturned(o) || isDispatchedFromOffice(o)) return false;
                            const refDate = o.supplierDispatchDate || o.orderDate;
                            if (!refDate) return false;
                            const d = new Date(refDate);
                            if (isNaN(d.getTime())) return false;
                            const days = Math.floor((today.getTime() - d.getTime()) / 86400000);
                            if (days <= 5) return false;

                            if (whStore && o.store !== whStore) return false;
                            if (whSup && o.supplier !== whSup) return false;
                            if (whQuery) {
                              const q = whQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          }).map(o => {
                            const refDate = o.supplierDispatchDate || o.orderDate;
                            const days = Math.floor((today.getTime() - new Date(refDate).getTime()) / 86400000);
                            return { ...o, _days: days };
                          }).sort((a, b) => b._days - a._days);

                          if (rows.length === 0) {
                            return <tr><td colSpan={11} className="p-6 text-center text-zinc-400">Nothing pending 5+ days</td></tr>;
                          }

                          return rows.map(o => (
                            <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              <td className="p-3">{o.orderDate || '-'}</td>
                              <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                              <td className="p-3">{renderBadge(o.store)}</td>
                              <td className="p-3">{o.customer || '-'}</td>
                              <td className="p-3">{o.mobile || '-'}</td>
                              <td className="p-3 truncate max-w-[200px]">{o.sku}</td>
                              <td className="p-3">{renderBadge(o.supplier)}</td>
                              <td className="p-3">{renderStatusBadge(o.supplierStatus || 'Not Dispatched')}</td>
                              <td className="p-3">{o.supplierDispatchDate || '-'}</td>
                              <td className="p-3 font-medium text-zinc-900">{o.orderValue || '-'}</td>
                              <td className="p-3"><span className="px-2 py-0.5 bg-rose-100 text-rose-800 rounded font-semibold text-[10px]">{o._days} days</span></td>
                            </tr>
                          ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 6. INSTOCK QUEUE */}
              {activeView === 'instock' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={instockStore}
                      onChange={e => setInstockStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <input
                      type="text"
                      value={instockQuery}
                      onChange={e => setInstockQuery(e.target.value)}
                      placeholder="Search order no / customer / SKU..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order date</th>
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">SKU</th>
                          <th className="p-3 font-medium">Order value</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const rows = activeRecords.filter(o => {
                            if (o.supplierStatus !== 'INSTOCK' || isDispatchedFromOffice(o) || isCancelled(o)) return false;
                            if (instockStore && o.store !== instockStore) return false;
                            if (instockQuery) {
                              const q = instockQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          });

                          if (rows.length === 0) {
                            return <tr><td colSpan={7} className="p-6 text-center text-zinc-400">No instock items pending dispatch</td></tr>;
                          }

                          return rows.map(o => (
                            <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              <td className="p-3">{o.orderDate || '-'}</td>
                              <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                              <td className="p-3">{renderBadge(o.store)}</td>
                              <td className="p-3">{o.customer || '-'}</td>
                              <td className="p-3">{o.mobile || '-'}</td>
                              <td className="p-3 truncate max-w-[200px]">{o.sku}</td>
                              <td className="p-3 font-medium text-zinc-900">{o.orderValue || '-'}</td>
                            </tr>
                          ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 7. STOCK OUT */}
              {activeView === 'stockout' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 text-xs text-zinc-500">
                    Items marked "Stock out" or "Sold out" on the supplier bill. These are auto-marked Received in WH = Yes.
                  </div>
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={stockoutStore}
                      onChange={e => setStockoutStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <input
                      type="text"
                      value={stockoutQuery}
                      onChange={e => setStockoutQuery(e.target.value)}
                      placeholder="Search order no / customer / SKU..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order date</th>
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">SKU</th>
                          <th className="p-3 font-medium">Supplier</th>
                          <th className="p-3 font-medium">Bill / note</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const candidates = activeRecords.filter(o => {
                            if (o.supplierStatus !== 'STOCK OUT' || isCancelled(o)) return false;
                            if (stockoutStore && o.store !== stockoutStore) return false;
                            if (stockoutQuery) {
                              const q = stockoutQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          });

                          // Extract each SKU row
                          const lines: any[] = [];
                          candidates.forEach(o => {
                            const parts = (o.sku || '').split(',').map(s => s.trim()).filter(Boolean);
                            const list = parts.length ? parts : [o.sku || ''];
                            list.forEach(part => {
                              const code = skuCode(part);
                              const skuDetailsObj = o.skuDetails ? (o.skuDetails instanceof Map ? Object.fromEntries(o.skuDetails.entries()) : o.skuDetails) : {};
                              const hasDetail = !!(skuDetailsObj && Object.prototype.hasOwnProperty.call(skuDetailsObj, code));
                              const detail = hasDetail ? skuDetailsObj[code] : {};

                              const supplier = hasDetail ? (detail.supplier || o.supplier || '-') : (o.supplier || '-');
                              const status = hasDetail ? (detail.supplierStatus || '') : (o.supplierStatus || '');

                              if (hasDetail && status !== 'STOCK OUT') return;
                              const note = (hasDetail ? (detail.supplierDispatchDate || '') : (o.supplierDispatchDate || '')) || '-';

                              lines.push({ o, part, supplier, note });
                            });
                          });

                          if (lines.length === 0) {
                            return <tr><td colSpan={8} className="p-6 text-center text-zinc-400">No stock out items</td></tr>;
                          }

                          return lines.map((line, index) => (
                            <tr key={`${line.o._id || line.o.orderNo}-${line.part}-${index}`} onClick={() => handleRowClick(line.o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              <td className="p-3">{line.o.orderDate || '-'}</td>
                              <td className="p-3 font-medium text-indigo-600">{line.o.orderNo}</td>
                              <td className="p-3">{renderBadge(line.o.store)}</td>
                              <td className="p-3">{line.o.customer || '-'}</td>
                              <td className="p-3">{line.o.mobile || '-'}</td>
                              <td className="p-3 truncate max-w-[200px]">{line.part}</td>
                              <td className="p-3">{renderBadge(line.supplier)}</td>
                              <td className="p-3 text-zinc-500">{line.note}</td>
                            </tr>
                          ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 8. COURIER — EWE */}
              {activeView === 'courier-ewe' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={eweStore}
                      onChange={e => setEweStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={eweStatus}
                      onChange={e => setEweStatus(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All statuses</option>
                      <option value="Delivered">Delivered</option>
                      <option value="In Warehouse">In Warehouse</option>
                      <option value="Order in the way">Order in the way</option>
                      <option value="Recalled">Recalled</option>
                      <option value="Cancelled">Cancelled</option>
                      <option value="OFD">OFD</option>
                      <option value="Pending">Pending</option>
                    </select>

                    <input
                      type="text"
                      value={eweQuery}
                      onChange={e => setEweQuery(e.target.value)}
                      placeholder="Search order no / customer / tracking..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">Dispatch date</th>
                          <th className="p-3 font-medium">Courier status</th>
                          <th className="p-3 font-medium">Tracking ID</th>
                          <th className="p-3 font-medium">COD status</th>
                          <th className="p-3 font-medium">COD amt received</th>
                          <th className="p-3 font-medium">Replaced / Return</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const rows = activeRecords.filter(o => {
                            if (!o.courier && !o.courierStatusEWE) return false;
                            if (isCancelled(o)) return false;
                            if (eweStore && o.store !== eweStore) return false;
                            if (eweStatus && o.courierStatusEWE !== eweStatus) return false;
                            if (eweQuery) {
                              const q = eweQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.trackingEWE].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          });

                          if (rows.length === 0) {
                            return <tr><td colSpan={10} className="p-6 text-center text-zinc-400">No EWE data yet — import EWE from the Import page</td></tr>;
                          }

                          return rows.map(o => (
                            <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                              <td className="p-3">{renderBadge(o.store)}</td>
                              <td className="p-3">{o.customer || '-'}</td>
                              <td className="p-3">{o.mobile || '-'}</td>
                              <td className="p-3">{o.dispatchDateEWE || '-'}</td>
                              <td className="p-3">{renderStatusBadge(o.courierStatusEWE)}</td>
                              <td className="p-3 font-mono text-[10px]">{o.trackingEWE || '-'}</td>
                              <td className="p-3">{o.codStatus || '-'}</td>
                              <td className="p-3">{o.codAmountReceived || '-'}</td>
                              <td className="p-3">
                                {o.replaced === 'Yes' ? (
                                  <span className="px-2 py-0.5 bg-rose-100 text-rose-800 rounded font-semibold text-[10px]">Replaced/Return</span>
                                ) : (
                                  renderStatusBadge(o.returnReceived)
                                )}
                              </td>
                            </tr>
                          ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 9. COURIER — TFM */}
              {activeView === 'courier-tfm' && (
                <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                  {/* Filter Bar */}
                  <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex flex-wrap gap-3 items-center">
                    <select
                      value={tfmStore}
                      onChange={e => setTfmStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <select
                      value={tfmStatus}
                      onChange={e => setTfmStatus(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All statuses</option>
                      <option value="Delivered">Delivered</option>
                      <option value="In Warehouse">In Warehouse</option>
                      <option value="Order in the way">Order in the way</option>
                      <option value="Recalled">Recalled</option>
                      <option value="Cancelled">Cancelled</option>
                      <option value="OFD">OFD</option>
                      <option value="Created">Created</option>
                    </select>

                    <input
                      type="text"
                      value={tfmQuery}
                      onChange={e => setTfmQuery(e.target.value)}
                      placeholder="Search order no / customer / tracking..."
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                          <th className="p-3 font-medium">Order no.</th>
                          <th className="p-3 font-medium">Store</th>
                          <th className="p-3 font-medium">Customer</th>
                          <th className="p-3 font-medium">Mobile</th>
                          <th className="p-3 font-medium">Dispatch date</th>
                          <th className="p-3 font-medium">Courier status</th>
                          <th className="p-3 font-medium">Tracking ID</th>
                          <th className="p-3 font-medium">Last comment</th>
                          <th className="p-3 font-medium">COD status</th>
                          <th className="p-3 font-medium">COD amt received</th>
                          <th className="p-3 font-medium">Replaced / Return</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const rows = activeRecords.filter(o => {
                            if (!o.tfmCourier && !o.courierStatusTFM) return false;
                            if (isCancelled(o)) return false;
                            if (tfmStore && o.store !== tfmStore) return false;
                            if (tfmStatus && o.courierStatusTFM !== tfmStatus) return false;
                            if (tfmQuery) {
                              const q = tfmQuery.toLowerCase();
                              return [o.orderNo, o.customer, o.trackingTFM].some(v => (v || '').toLowerCase().includes(q));
                            }
                            return true;
                          });

                          if (rows.length === 0) {
                            return <tr><td colSpan={11} className="p-6 text-center text-zinc-400">No TFM data yet — import TFM from the Import page</td></tr>;
                          }

                          return rows.map(o => (
                            <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                              <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                              <td className="p-3">{renderBadge(o.store)}</td>
                              <td className="p-3">{o.customer || '-'}</td>
                              <td className="p-3">{o.mobile || '-'}</td>
                              <td className="p-3">{o.dispatchDateTFM || '-'}</td>
                              <td className="p-3">{renderStatusBadge(o.courierStatusTFM)}</td>
                              <td className="p-3 font-mono text-[10px]">{o.trackingTFM || '-'}</td>
                              <td className="p-3 text-[10px] text-zinc-500 max-w-[200px] truncate" title={o.lastComment}>{o.lastComment || '-'}</td>
                              <td className="p-3">{o.codStatus || '-'}</td>
                              <td className="p-3">{o.codAmountReceived || '-'}</td>
                              <td className="p-3">
                                {o.replaced === 'Yes' ? (
                                  <span className="px-2 py-0.5 bg-rose-100 text-rose-800 rounded font-semibold text-[10px]">Replaced/Return</span>
                                ) : (
                                  renderStatusBadge(o.returnReceived)
                                )}
                              </td>
                            </tr>
                          ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* 10. ORDER STATUS SEARCH */}
              {activeView === 'orderstatus' && (
                <div className="space-y-6 w-full">
                  {/* Search Bar */}
                  <div className="bg-white border border-zinc-200 rounded-xl p-5 shadow-sm">
                    <div className="max-w-2xl">
                      <label className="block text-xs font-bold text-zinc-500 uppercase tracking-wider mb-2">Track Order Status</label>
                      <div className="relative">
                        <Search className="absolute left-3 top-3 w-4 h-4 text-zinc-400" />
                        <input
                          type="text"
                          value={statusQuery}
                          onChange={e => setStatusQuery(e.target.value)}
                          placeholder="Search order number, mobile number, customer name..."
                          className="w-full pl-10 pr-4 py-2.5 border border-zinc-200 bg-zinc-50 focus:bg-white rounded-lg text-sm shadow-sm focus:outline-none focus:border-indigo-500 transition-all"
                        />
                      </div>
                    </div>
                  </div>

                  {(() => {
                    const q = statusQuery.trim().toLowerCase();
                    if (!q) {
                      return (
                        <div className="text-center py-16 bg-white border border-zinc-200 rounded-xl shadow-sm flex flex-col items-center justify-center gap-3">
                          <Search className="w-10 h-10 text-zinc-300 animate-pulse" />
                          <span className="text-xs text-zinc-400 font-medium">Search using reference filters above to check shipment status details.</span>
                        </div>
                      );
                    }

                    const digits = (v: string) => (v || '').replace(/\D/g, '');
                    const qDigits = digits(q);

                    let matches = db.filter(o => (o.orderNo || '').toLowerCase() === q || (qDigits.length >= 7 && digits(o.mobile) === qDigits));
                    if (matches.length === 0) {
                      const partial = db.filter(o =>
                        (o.orderNo || '').toLowerCase().includes(q) ||
                        (o.customer || '').toLowerCase().includes(q) ||
                        (qDigits.length >= 4 && digits(o.mobile).includes(qDigits))
                      );

                      if (partial.length === 0) {
                        return (
                          <div className="text-center py-16 bg-white border border-zinc-200 rounded-xl shadow-sm text-xs text-zinc-400">
                            No matching order found. Please double check the reference or mobile number.
                          </div>
                        );
                      }

                      const orderNos = [...new Set(partial.map(o => o.orderNo))];
                      if (orderNos.length <= 5) {
                        matches = partial;
                      } else {
                        return (
                          <div className="text-center py-12 bg-white border border-zinc-200 rounded-xl shadow-sm text-xs text-zinc-400">
                            No exact match. Did you mean: <strong className="text-indigo-600">{orderNos.slice(0, 8).join(', ')}</strong>?
                          </div>
                        );
                      }
                    }

                    return (
                      <div className="space-y-6">
                        {matches.map(o => {
                          const isCancelledVal = isCancelled(o);
                          const isSupDisp = o.supplierDispatchDate || o.supplierStatus === 'Dispatched' || isDispatchedFromOffice(o);
                          const isOfficeDisp = isDispatchedFromOffice(o);
                          const isDel = isDelivered(o);

                          const customerFields = [
                            ['Store', o.store],
                            ['Customer Name', o.customer],
                            ['Mobile Number', o.mobile],
                            ['Delivery City', o.city]
                          ];
                          const supplierFields = [
                            ['Supplier Name', o.supplier],
                            ['Supplier Status', o.supplierStatus],
                            ['Dispatch Date', o.supplierDispatchDate],
                            ['Received in WH', o.receivedInWH]
                          ];
                          const courierFields = [
                            ['Courier Name', o.courier || o.tfmCourier || '—'],
                            ['Tracking (EWE)', o.trackingEWE],
                            ['Tracking (TFM)', o.trackingTFM],
                            ['Courier Status EWE', o.courierStatusEWE],
                            ['Courier Status TFM', o.courierStatusTFM],
                            ['Dispatch Date (EWE)', o.dispatchDateEWE],
                            ['Dispatch Date (TFM)', o.dispatchDateTFM]
                          ];
                          const financeFields = [
                            ['Order Value', o.orderValue ? `${o.orderValue} AED` : '—'],
                            ['COD Status', o.codStatus],
                            ['COD Received Status', o.codReceived],
                            ['Replaced Shipment', o.replaced],
                            ['Return Status', o.returnReceived]
                          ];

                          return (
                            <div key={o._id} className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm flex flex-col gap-6 p-6">
                              {/* Header Summary */}
                              <div className="flex flex-wrap justify-between items-start border-b border-zinc-100 pb-4 gap-3">
                                <div>
                                  <div className="flex items-center gap-2 flex-wrap">
                                    <h2 className="font-extrabold text-zinc-900 text-lg tracking-tight">{o.orderNo}</h2>
                                    {renderBadge(o.store)}
                                    {isCancelledVal && <span className="bg-rose-100 text-rose-800 text-[10px] font-bold px-2 py-0.5 rounded-full">Cancelled</span>}
                                  </div>
                                  <p className="text-xs text-zinc-400 mt-0.5">Order Date: <span className="font-semibold text-zinc-600">{o.orderDate || '—'}</span></p>
                                </div>
                                <div className="text-right">
                                  <span className="text-[10px] uppercase font-bold text-zinc-400 block tracking-wide">Total Value</span>
                                  <span className="text-lg font-bold text-indigo-600">{o.orderValue || '—'} AED</span>
                                </div>
                              </div>

                              {/* Progress Timeline */}
                              {!isCancelledVal && (
                                <div className="bg-zinc-50 p-4 rounded-xl border border-zinc-100">
                                  <h4 className="text-[10px] uppercase font-bold text-zinc-400 tracking-wider mb-4">Shipment Lifecycle Timeline</h4>
                                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                                    {[
                                      { label: 'Order Received', date: o.orderDate, completed: true, style: 'border-emerald-500 text-emerald-700 bg-emerald-50' },
                                      { label: 'Supplier Dispatched', date: o.supplierDispatchDate, completed: isSupDisp, style: isSupDisp ? 'border-emerald-500 text-emerald-700 bg-emerald-50' : 'border-zinc-200 text-zinc-400 bg-zinc-100/50' },
                                      { label: 'Courier Shipped', date: o.dispatchDateEWE || o.dispatchDateTFM, completed: isOfficeDisp, style: isOfficeDisp ? 'border-emerald-500 text-emerald-700 bg-emerald-50' : 'border-zinc-200 text-zinc-400 bg-zinc-100/50' },
                                      { label: 'Delivered', date: o.dispatchDateEWE || o.dispatchDateTFM, completed: isDel, style: isDel ? 'border-emerald-500 text-emerald-700 bg-emerald-50' : 'border-zinc-200 text-zinc-400 bg-zinc-100/50' }
                                    ].map((step, idx) => (
                                      <div key={idx} className={`p-3 rounded-lg border flex flex-col justify-between min-h-[72px] transition-all shadow-sm ${step.style}`}>
                                        <div className="flex justify-between items-center">
                                          <span className="text-xs font-bold">{step.label}</span>
                                          <span className="text-[9px] uppercase tracking-wider font-extrabold opacity-75">Step {idx + 1}</span>
                                        </div>
                                        <span className="text-[10px] font-medium opacity-85 mt-1">{step.completed ? (step.date || 'Completed') : 'Pending'}</span>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              )}

                              {/* Details Grid */}
                              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                                {/* Customer & Store */}
                                <div className="bg-white border border-zinc-200 rounded-xl p-4 shadow-sm">
                                  <h3 className="font-bold text-zinc-700 text-xs mb-3 border-b border-zinc-100 pb-2 flex items-center gap-1.5">
                                    <Info className="w-3.5 h-3.5 text-indigo-500" />
                                    Customer & Store Info
                                  </h3>
                                  <table className="w-full text-xs">
                                    <tbody>
                                      {customerFields.map(([label, val]) => (
                                        <tr key={label} className="border-b border-zinc-50 last:border-0">
                                          <td className="py-1.5 text-zinc-400 font-medium">{label}</td>
                                          <td className="py-1.5 text-zinc-800 font-semibold text-right">{val || '—'}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>

                                {/* Supplier details */}
                                <div className="bg-white border border-zinc-200 rounded-xl p-4 shadow-sm">
                                  <h3 className="font-bold text-zinc-700 text-xs mb-3 border-b border-zinc-100 pb-2 flex items-center gap-1.5">
                                    <Clock className="w-3.5 h-3.5 text-amber-500" />
                                    Supplier Tracking
                                  </h3>
                                  <table className="w-full text-xs">
                                    <tbody>
                                      {supplierFields.map(([label, val]) => (
                                        <tr key={label} className="border-b border-zinc-50 last:border-0">
                                          <td className="py-1.5 text-zinc-400 font-medium">{label}</td>
                                          <td className="py-1.5 text-zinc-800 font-semibold text-right">
                                            {label === 'Supplier Status' ? renderStatusBadge(val) : (val || '—')}
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>

                                {/* Logistics Details */}
                                <div className="bg-white border border-zinc-200 rounded-xl p-4 shadow-sm">
                                  <h3 className="font-bold text-zinc-700 text-xs mb-3 border-b border-zinc-100 pb-2 flex items-center gap-1.5">
                                    <Truck className="w-3.5 h-3.5 text-indigo-500" />
                                    Courier Logistics
                                  </h3>
                                  <table className="w-full text-xs">
                                    <tbody>
                                      {courierFields.map(([label, val]) => (
                                        <tr key={label} className="border-b border-zinc-50 last:border-0">
                                          <td className="py-1.5 text-zinc-400 font-medium truncate max-w-[100px]" title={label}>{label}</td>
                                          <td className="py-1.5 text-zinc-800 font-semibold text-right truncate max-w-[120px]" title={val}>{val || '—'}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>

                                {/* Finance Details */}
                                <div className="bg-white border border-zinc-200 rounded-xl p-4 shadow-sm">
                                  <h3 className="font-bold text-zinc-700 text-xs mb-3 border-b border-zinc-100 pb-2 flex items-center gap-1.5">
                                    <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-500" />
                                    COD & Returns
                                  </h3>
                                  <table className="w-full text-xs">
                                    <tbody>
                                      {financeFields.map(([label, val]) => (
                                        <tr key={label} className="border-b border-zinc-50 last:border-0">
                                          <td className="py-1.5 text-zinc-400 font-medium">{label}</td>
                                          <td className="py-1.5 text-zinc-800 font-semibold text-right">{val || '—'}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              {/* Items Breakdown list */}
                              <div>
                                <h4 className="text-[10px] uppercase font-bold text-zinc-400 tracking-wider mb-2">Order Items Breakdown</h4>
                                {itemsBreakdown(o)}
                              </div>

                              {/* Comments Log notes */}
                              <div className="bg-zinc-50 rounded-xl border border-zinc-100 p-4 flex flex-col gap-3">
                                <div className="flex justify-between items-center border-b border-zinc-200/50 pb-2">
                                  <h4 className="text-[10px] uppercase font-bold text-zinc-400 tracking-wider">Comments Log</h4>
                                  {o.lastComment && <span className="text-[10px] font-medium text-zinc-500">TFM Courier Comment: <strong className="text-zinc-700">{o.lastComment}</strong></span>}
                                </div>
                                <div className="flex gap-3 items-start">
                                  <textarea
                                    defaultValue={o.comment || ''}
                                    onBlur={async (e) => {
                                      const val = e.target.value;
                                      if (val === o.comment) return;
                                      try {
                                        const res = await fetch('/api/orders/comment', {
                                          method: 'PUT',
                                          headers: { 'Content-Type': 'application/json' },
                                          body: JSON.stringify({ orderNo: o.orderNo, comment: val })
                                        });
                                        const commentData = await res.json();
                                        if (commentData.success) {
                                          showToast('Comment updated successfully');
                                          o.comment = val;
                                        }
                                      } catch (err) {
                                        console.error('Failed to update comment:', err);
                                        showToast('Could not save comment update', false);
                                      }
                                    }}
                                    placeholder="Write a comment note for this order (press tab/click away to save)..."
                                    className="w-full text-xs p-3 border border-zinc-200 bg-white rounded-lg focus:outline-none focus:border-indigo-500 shadow-sm"
                                    rows={2}
                                  />
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })()}
                </div>
              )}

              {/* 10.5. IMPORT LOGS */}
              {activeView === 'importlogs' && (
                <div className="space-y-6 w-full">
                  <div className="flex justify-between items-center flex-wrap gap-4">
                    <div>
                      <h2 className="font-extrabold text-zinc-900 text-lg tracking-tight">Import Diagnostic Logs</h2>
                      <p className="text-xs text-zinc-500 mt-0.5">Review skipped records, parsing mismatches, and integrity checks from your spreadsheet uploads.</p>
                    </div>
                    <button
                      onClick={async () => {
                        await loadImportLogs();
                        showToast('Logs refreshed');
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white border border-zinc-200 hover:bg-zinc-50 text-zinc-700 rounded-lg text-xs font-semibold shadow-sm transition"
                    >
                      <RefreshCcw className="w-3.5 h-3.5" />
                      Refresh Logs
                    </button>
                  </div>

                  {importLogs.length === 0 ? (
                    <div className="text-center py-20 bg-white border border-zinc-200 rounded-xl shadow-sm flex flex-col items-center justify-center gap-3">
                      <div className="p-4 bg-emerald-50 text-emerald-600 rounded-full">
                        <CheckCircle className="w-10 h-10" />
                      </div>
                      <h3 className="font-bold text-zinc-800 text-sm">All Records Clear</h3>
                      <p className="text-zinc-500 text-xs max-w-xs mx-auto leading-relaxed">
                        No rows were skipped or flagged in recent uploads. All orders have been matched and imported successfully!
                      </p>
                    </div>
                  ) : (
                    <div className="bg-white border border-zinc-200 rounded-xl shadow-sm overflow-hidden">
                      <div className="p-4 border-b border-zinc-100 bg-zinc-50/50 flex justify-between items-center flex-wrap gap-3">
                        <span className="text-xs font-bold text-zinc-600">Skipped Rows / Diagnostic Reports ({importLogs.length})</span>
                      </div>
                      
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse text-xs">
                          <thead>
                            <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase font-bold border-b border-zinc-200">
                              <th className="p-3 font-semibold">Import Type</th>
                              <th className="p-3 font-semibold">Spreadsheet Row</th>
                              <th className="p-3 font-semibold">Diagnostic Exception Reason</th>
                              <th className="p-3 font-semibold">Row Data Details</th>
                            </tr>
                          </thead>
                          <tbody>
                            {importLogs.map((log: any) => {
                              let typeColor = 'bg-zinc-100 text-zinc-700 border-zinc-200';
                              if (log.importType === 'orders' || log.importType === 'at') typeColor = 'bg-indigo-50 text-indigo-700 border-indigo-200';
                              else if (log.importType === 'suppliers' || log.importType === 'sup') typeColor = 'bg-amber-50 text-amber-700 border-amber-200';
                              else if (log.importType === 'ewe' || log.importType === 'tfm') typeColor = 'bg-blue-50 text-blue-700 border-blue-200';
                              else if (log.importType === 'cod') typeColor = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                              else if (log.importType === 'returns' || log.importType === 'ret') typeColor = 'bg-rose-50 text-rose-700 border-rose-200';

                              return (
                                <tr key={log._id} className="border-b border-zinc-100 hover:bg-zinc-50/50">
                                  <td className="p-3">
                                    <span className={`px-2 py-0.5 text-[10px] font-bold border rounded-md capitalize ${typeColor}`}>
                                      {log.importType}
                                    </span>
                                  </td>
                                  <td className="p-3 font-semibold text-zinc-700">Row {log.rowNumber}</td>
                                  <td className="p-3 font-medium text-rose-600">{log.reason}</td>
                                  <td className="p-3">
                                    <details className="text-[10px] text-zinc-500 cursor-pointer">
                                      <summary className="font-semibold text-indigo-600 hover:underline">View row attributes</summary>
                                      <pre className="mt-2 p-2 bg-zinc-50 border border-zinc-100 rounded text-[9px] font-mono overflow-auto max-w-lg max-h-40">
                                        {JSON.stringify(log.rawRowData, null, 2)}
                                      </pre>
                                    </details>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* 11. REPORTS */}
              {activeView === 'reports' && (
                <div className="space-y-6">
                  {/* Filter Bar */}
                  <div className="p-4 bg-white border border-zinc-200 rounded-xl flex flex-wrap gap-3 items-center shadow-sm">
                    <select
                      value={reportsStore}
                      onChange={e => setReportsStore(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none"
                    >
                      <option value="">All stores</option>
                      {[...new Set(db.map(o => o.store))].filter(Boolean).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>

                    <input
                      type="text"
                      value={reportsQuery}
                      onChange={e => setReportsQuery(e.target.value)}
                      placeholder="Search order no / customer / SKU..."
                      className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-xs shadow-sm focus:outline-none flex-1 max-w-xs"
                    />
                  </div>

                  {(() => {
                    const matchFilter = (o: OrderRecord) => {
                      if (isCancelled(o)) return false;
                      if (reportsStore && o.store !== reportsStore) return false;
                      if (reportsQuery) {
                        const q = reportsQuery.toLowerCase();
                        return [o.orderNo, o.customer, o.sku].some(v => (v || '').toLowerCase().includes(q));
                      }
                      return true;
                    };

                    const replReturnPending = activeRecords.filter(o => o.replaced === 'Yes' && (o.returnReceived || '').trim().toLowerCase() !== 'received').filter(matchFilter);
                    const returnKeys = new Set(replReturnPending.map(o => (o.orderNo || '') + '||' + (o.sku || '')));
                    const codPendingEWE = activeRecords.filter(o => isDeliveredEWE(o) && (o.codStatus || '').trim().toLowerCase() !== 'received').filter(matchFilter);
                    const codPendingTFM = activeRecords.filter(o => isDeliveredTFM(o) && (o.codStatus || '').trim().toLowerCase() !== 'received').filter(matchFilter);

                    const onHoldEWE = activeRecords.filter(o => {
                      if (!(o.courier === 'EWE' || o.courierStatusEWE)) return false;
                      if (isDeliveredEWE(o) || isReturned(o)) return false;
                      if ((o.courierStatusEWE || '').trim().toLowerCase() === 'pending') return false;
                      if (returnKeys.has((o.orderNo || '') + '||' + (o.sku || ''))) return false;
                      return true;
                    }).filter(matchFilter);

                    const onHoldTFM = activeRecords.filter(o => {
                      if (!(o.tfmCourier === 'TFM' || o.courierStatusTFM)) return false;
                      if (isDeliveredTFM(o) || isReturned(o)) return false;
                      if ((o.courierStatusTFM || '').trim().toLowerCase() === 'created') return false;
                      if (returnKeys.has((o.orderNo || '') + '||' + (o.sku || ''))) return false;
                      return true;
                    }).filter(matchFilter);

                    const sumAmount = (list: OrderRecord[]) => {
                      return list.reduce((t, o) => {
                        const v = parseFloat((o.codAmountReceived || o.orderValue || '0').replace(/[^0-9.\-]/g, ''));
                        return t + (isNaN(v) ? 0 : v);
                      }, 0);
                    };

                    const sumOrderValue = (list: OrderRecord[]) => {
                      return list.reduce((t, o) => {
                        const v = parseFloat((o.orderValue || '0').replace(/[^0-9.\-]/g, ''));
                        return t + (isNaN(v) ? 0 : v);
                      }, 0);
                    };

                    return (
                      <div className="space-y-6">
                        {/* Report Table 1: Replacement returns */}
                        <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                          <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                            <h3 className="font-semibold text-zinc-800 text-xs">Replacement / return — pending back to warehouse</h3>
                            <span className="bg-rose-100 text-rose-800 font-bold px-2 py-0.5 rounded-full text-[10px]">{replReturnPending.length}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                              <thead>
                                <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                                  <th className="p-3 font-medium">Order no.</th>
                                  <th className="p-3 font-medium">Store</th>
                                  <th className="p-3 font-medium">Customer</th>
                                  <th className="p-3 font-medium">Mobile</th>
                                  <th className="p-3 font-medium">Supplier</th>
                                  <th className="p-3 font-medium">Courier</th>
                                  <th className="p-3 font-medium">Return status</th>
                                </tr>
                              </thead>
                              <tbody>
                                {replReturnPending.length === 0 ? (
                                  <tr><td colSpan={7} className="p-6 text-center text-zinc-400">Nothing pending</td></tr>
                                ) : (
                                  replReturnPending.map(o => (
                                    <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                      <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                      <td className="p-3">{renderBadge(o.store)}</td>
                                      <td className="p-3">{o.customer || '-'}</td>
                                      <td className="p-3">{o.mobile || '-'}</td>
                                      <td className="p-3">{renderBadge(o.supplier)}</td>
                                      <td className="p-3">{renderBadge(o.courier || o.tfmCourier)}</td>
                                      <td className="p-3">{renderStatusBadge(o.returnReceived || 'Pending')}</td>
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Report Table 2: EWE Delivered - COD pending */}
                        <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                          <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                            <h3 className="font-semibold text-zinc-800 text-xs">Delivered — COD not received (Easyway)</h3>
                            <span className="bg-rose-100 text-rose-800 font-bold px-2 py-0.5 rounded-full text-[10px]">{codPendingEWE.length}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                              <thead>
                                <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                                  <th className="p-3 font-medium">Order no.</th>
                                  <th className="p-3 font-medium">Store</th>
                                  <th className="p-3 font-medium">Customer</th>
                                  <th className="p-3 font-medium">Mobile</th>
                                  <th className="p-3 font-medium">COD amount</th>
                                  <th className="p-3 font-medium">COD status</th>
                                </tr>
                              </thead>
                              <tbody>
                                {codPendingEWE.length === 0 ? (
                                  <tr><td colSpan={6} className="p-6 text-center text-zinc-400">None</td></tr>
                                ) : (
                                  <>
                                    {codPendingEWE.map(o => (
                                      <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                        <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                        <td className="p-3">{renderBadge(o.store)}</td>
                                        <td className="p-3">{o.customer || '-'}</td>
                                        <td className="p-3">{o.mobile || '-'}</td>
                                        <td className="p-3 font-medium text-zinc-900">{o.codAmountReceived || o.orderValue || '-'}</td>
                                        <td className="p-3">{renderStatusBadge(o.codStatus || 'Pending')}</td>
                                      </tr>
                                    ))}
                                    <tr className="bg-zinc-50 font-semibold border-t border-zinc-300">
                                      <td colSpan={4} className="p-3 text-right">Total:</td>
                                      <td colSpan={2} className="p-3 text-indigo-600 font-bold">{sumAmount(codPendingEWE).toLocaleString()} AED</td>
                                    </tr>
                                  </>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Report Table 3: TFM Delivered - COD pending */}
                        <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                          <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                            <h3 className="font-semibold text-zinc-800 text-xs">Delivered — COD not received (TFM)</h3>
                            <span className="bg-rose-100 text-rose-800 font-bold px-2 py-0.5 rounded-full text-[10px]">{codPendingTFM.length}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                              <thead>
                                <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                                  <th className="p-3 font-medium">Order no.</th>
                                  <th className="p-3 font-medium">Store</th>
                                  <th className="p-3 font-medium">Customer</th>
                                  <th className="p-3 font-medium">Mobile</th>
                                  <th className="p-3 font-medium">COD amount</th>
                                  <th className="p-3 font-medium">COD status</th>
                                </tr>
                              </thead>
                              <tbody>
                                {codPendingTFM.length === 0 ? (
                                  <tr><td colSpan={6} className="p-6 text-center text-zinc-400">None</td></tr>
                                ) : (
                                  <>
                                    {codPendingTFM.map(o => (
                                      <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                        <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                        <td className="p-3">{renderBadge(o.store)}</td>
                                        <td className="p-3">{o.customer || '-'}</td>
                                        <td className="p-3">{o.mobile || '-'}</td>
                                        <td className="p-3 font-medium text-zinc-900">{o.codAmountReceived || o.orderValue || '-'}</td>
                                        <td className="p-3">{renderStatusBadge(o.codStatus || 'Pending')}</td>
                                      </tr>
                                    ))}
                                    <tr className="bg-zinc-50 font-semibold border-t border-zinc-300">
                                      <td colSpan={4} className="p-3 text-right">Total:</td>
                                      <td colSpan={2} className="p-3 text-indigo-600 font-bold">{sumAmount(codPendingTFM).toLocaleString()} AED</td>
                                    </tr>
                                  </>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Report Table 4: On Hold EWE */}
                        <div id="sec-rep-holdewe" className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                          <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                            <h3 className="font-semibold text-zinc-800 text-xs">On hold with Easyway</h3>
                            <span className="bg-amber-100 text-amber-800 font-bold px-2 py-0.5 rounded-full text-[10px]">{onHoldEWE.length}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                              <thead>
                                <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                                  <th className="p-3 font-medium">Order no.</th>
                                  <th className="p-3 font-medium">Store</th>
                                  <th className="p-3 font-medium">Customer</th>
                                  <th className="p-3 font-medium">Mobile</th>
                                  <th className="p-3 font-medium">Supplier</th>
                                  <th className="p-3 font-medium">Courier status</th>
                                  <th className="p-3 font-medium">Order value</th>
                                </tr>
                              </thead>
                              <tbody>
                                {onHoldEWE.length === 0 ? (
                                  <tr><td colSpan={7} className="p-6 text-center text-zinc-400">Nothing on hold</td></tr>
                                ) : (
                                  <>
                                    {onHoldEWE.map(o => (
                                      <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                        <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                        <td className="p-3">{renderBadge(o.store)}</td>
                                        <td className="p-3">{o.customer || '-'}</td>
                                        <td className="p-3">{o.mobile || '-'}</td>
                                        <td className="p-3">{renderBadge(o.supplier)}</td>
                                        <td className="p-3">{renderStatusBadge(o.courierStatusEWE)}</td>
                                        <td className="p-3 font-medium text-zinc-900">{o.orderValue || '-'}</td>
                                      </tr>
                                    ))}
                                    <tr className="bg-zinc-50 font-semibold border-t border-zinc-300">
                                      <td colSpan={6} className="p-3 text-right">Total Order Value:</td>
                                      <td className="p-3 text-indigo-600 font-bold">{sumOrderValue(onHoldEWE).toLocaleString()} AED</td>
                                    </tr>
                                  </>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Report Table 5: On Hold TFM */}
                        <div id="sec-rep-holdtfm" className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                          <div className="px-4 py-3 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                            <h3 className="font-semibold text-zinc-800 text-xs">On hold with TFM</h3>
                            <span className="bg-amber-100 text-amber-800 font-bold px-2 py-0.5 rounded-full text-[10px]">{onHoldTFM.length}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                              <thead>
                                <tr className="bg-zinc-50 text-[10px] text-zinc-500 uppercase border-b border-zinc-200">
                                  <th className="p-3 font-medium">Order no.</th>
                                  <th className="p-3 font-medium">Store</th>
                                  <th className="p-3 font-medium">Customer</th>
                                  <th className="p-3 font-medium">Mobile</th>
                                  <th className="p-3 font-medium">Supplier</th>
                                  <th className="p-3 font-medium">Courier status</th>
                                  <th className="p-3 font-medium">Order value</th>
                                </tr>
                              </thead>
                              <tbody>
                                {onHoldTFM.length === 0 ? (
                                  <tr><td colSpan={7} className="p-6 text-center text-zinc-400">Nothing on hold</td></tr>
                                ) : (
                                  <>
                                    {onHoldTFM.map(o => (
                                      <tr key={o._id} onClick={() => handleRowClick(o)} className="hover:bg-zinc-50 border-b border-zinc-100 cursor-pointer">
                                        <td className="p-3 font-medium text-indigo-600">{o.orderNo}</td>
                                        <td className="p-3">{renderBadge(o.store)}</td>
                                        <td className="p-3">{o.customer || '-'}</td>
                                        <td className="p-3">{o.mobile || '-'}</td>
                                        <td className="p-3">{renderBadge(o.supplier)}</td>
                                        <td className="p-3">{renderStatusBadge(o.courierStatusTFM)}</td>
                                        <td className="p-3 font-medium text-zinc-900">{o.orderValue || '-'}</td>
                                      </tr>
                                    ))}
                                    <tr className="bg-zinc-50 font-semibold border-t border-zinc-300">
                                      <td colSpan={6} className="p-3 text-right">Total Order Value:</td>
                                      <td className="p-3 text-indigo-600 font-bold">{sumOrderValue(onHoldTFM).toLocaleString()} AED</td>
                                    </tr>
                                  </>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}
            </>
          )}
        </div>
      </main>

      {/* Detailed Modal Dialog */}
      {selectedOrder && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
          {/* Backdrop with modern glassmorphism blur */}
          <div
            onClick={() => setSelectedOrder(null)}
            className="absolute inset-0 bg-zinc-900/60 backdrop-blur-sm transition-opacity duration-300"
          />

          {/* Modal Container */}
          <div className="relative bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col animate-in fade-in zoom-in-95 duration-200 border border-zinc-100 z-50">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-zinc-150 flex justify-between items-center bg-zinc-50 flex-shrink-0">
              <div>
                <h3 className="font-bold text-zinc-900 text-sm flex items-center gap-2">
                  Order <span className="text-indigo-600">#{selectedOrder.orderNo}</span>
                  {renderBadge(selectedOrder.store)}
                </h3>
                <p className="text-[10px] text-zinc-400 font-semibold mt-0.5">Order Date: {selectedOrder.orderDate || '—'}</p>
              </div>
              <button
                onClick={() => setSelectedOrder(null)}
                className="p-1.5 hover:bg-zinc-200 rounded-full transition text-zinc-400 hover:text-zinc-700"
              >
                <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            {/* Modal Scrollable Body */}
            <div className="flex-1 overflow-y-auto p-6 space-y-6">
              {/* Info Grid */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                
                {/* Section 1: Customer Details */}
                <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100 space-y-2.5">
                  <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-2 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                    <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" /></svg>
                    Customer & Store Info
                  </h4>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Customer Name</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.customer || '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Mobile Number</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.mobile || '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">City / Emirate</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.city || '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Order Value</span>
                    <span className="font-semibold text-indigo-600">{selectedOrder.orderValue ? `${selectedOrder.orderValue} AED` : '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1">
                    <span className="text-zinc-500">No. Items</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.noItems || '—'}</span>
                  </div>
                </div>

                {/* Section 2: Supplier Details */}
                <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100 space-y-2.5">
                  <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-2 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                    <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" /></svg>
                    Supplier Tracking
                  </h4>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Supplier Name</span>
                    <span>{renderBadge(selectedOrder.supplier)}</span>
                  </div>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Supplier Status</span>
                    <span>{renderStatusBadge(selectedOrder.supplierStatus)}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Supplier Dispatch Date</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.supplierDispatchDate || '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Instock Status</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.instock || '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1">
                    <span className="text-zinc-500">Received in WH</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.receivedInWH || '—'}</span>
                  </div>
                </div>

                {/* Section 3: Courier Details */}
                <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100 space-y-2.5">
                  <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-2 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                    <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                    Courier Logistics & Status
                  </h4>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Primary Courier</span>
                    <span>{renderBadge(selectedOrder.courier || selectedOrder.tfmCourier)}</span>
                  </div>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Office Dispatch</span>
                    <span>
                      {isDispatchedFromOffice(selectedOrder) ? (
                        <span className="px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded font-semibold text-[10px]">Yes</span>
                      ) : (
                        <span className="px-2 py-0.5 bg-zinc-50 text-zinc-400 border border-zinc-200 rounded font-semibold text-[10px]">No</span>
                      )}
                    </span>
                  </div>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">EWE Status</span>
                    <span className="font-semibold text-zinc-850 flex items-center gap-1.5">
                      {renderStatusBadge(selectedOrder.courierStatusEWE)}
                      {selectedOrder.trackingEWE && (
                        <span className="text-[10px] text-zinc-400 font-mono">({selectedOrder.trackingEWE})</span>
                      )}
                    </span>
                  </div>
                  <div className="flex justify-between items-center text-xs py-1">
                    <span className="text-zinc-500">TFM Status</span>
                    <span className="font-semibold text-zinc-850 flex items-center gap-1.5">
                      {renderStatusBadge(selectedOrder.courierStatusTFM)}
                      {selectedOrder.trackingTFM && (
                        <span className="text-[10px] text-zinc-400 font-mono">({selectedOrder.trackingTFM})</span>
                      )}
                    </span>
                  </div>
                </div>

                {/* Section 4: COD & Return Status */}
                <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100 space-y-2.5">
                  <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-2 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                    <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                    Finance & Returns
                  </h4>
                  <div className="flex justify-between items-center text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">COD Status</span>
                    <span>{renderStatusBadge(selectedOrder.codStatus || 'Pending')}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">COD Amount Received</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.codAmountReceived ? `${selectedOrder.codAmountReceived} AED` : '—'}</span>
                  </div>
                  <div className="flex justify-between text-xs py-1 border-b border-zinc-100">
                    <span className="text-zinc-500">Replaced Status</span>
                    <span className="font-semibold text-zinc-800">{selectedOrder.replaced || '—'}</span>
                  </div>
                  <div className="flex justify-between items-center text-xs py-1">
                    <span className="text-zinc-500">Return Received</span>
                    <span>{renderStatusBadge(selectedOrder.returnReceived)}</span>
                  </div>
                </div>

              </div>

              {/* Items Breakdown list */}
              <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100">
                <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-3 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                  <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 10h16M4 14h16M4 18h16" /></svg>
                  Items Breakdown (SKUs)
                </h4>
                {itemsBreakdown(selectedOrder)}
              </div>

              {/* Inline Comment Editor */}
              <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-100">
                <h4 className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-2 border-b border-zinc-100 pb-1 flex items-center gap-1.5">
                  <svg className="w-3.5 h-3.5 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 12h.01M12 12h.01M16 12h.01M21 12c0 4.418-4.03 8-9 8a9.863 9.863 0 01-4.255-.949L3 20l1.395-3.72C3.512 15.042 3 13.574 3 12c0-4.418 4.03-8 9-8s9 3.582 9 8z" /></svg>
                  Order Comments / Notes
                </h4>
                {selectedOrder.lastComment && (
                  <div className="text-[11px] text-zinc-500 bg-amber-50 border border-amber-200/50 rounded-lg p-2.5 mb-3">
                    <span className="font-semibold text-amber-800">TFM Last Comment:</span> {selectedOrder.lastComment}
                  </div>
                )}
                <div className="flex gap-3">
                  <textarea
                    value={editingComment}
                    onChange={e => setEditingComment(e.target.value)}
                    placeholder="Enter order tracking notes or remarks..."
                    className="flex-1 min-h-[70px] border border-zinc-200 rounded-lg p-2.5 text-xs bg-white focus:outline-none focus:border-indigo-500 resize-y"
                  />
                  <button
                    onClick={() => handleUpdateComment(selectedOrder.orderNo, selectedOrder.sku, editingComment)}
                    className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold self-end transition shadow-md hover:shadow-indigo-600/10"
                  >
                    Save Comment
                  </button>
                </div>
              </div>

            </div>
          </div>
        </div>
      )}
    </div>
  );
}
