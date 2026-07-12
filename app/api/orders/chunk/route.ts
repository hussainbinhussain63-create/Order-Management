import { NextResponse } from 'next/server';
import { dbConnect } from '@/lib/mongodb';
import Order from '@/models/Order';
import SupplierImport from '@/models/SupplierImport';
import CourierImport from '@/models/CourierImport';
import CODImport from '@/models/CODImport';
import ReturnImport from '@/models/ReturnImport';
import ImportLog from '@/models/ImportLog';

// Helper: Normalize headers
function norm(h: string): string {
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
}

// Helper: Derive supplier status
function deriveSupStatus(r: any): string {
  if ((r.orderStatus || '').toLowerCase().includes('cancel')) return 'Cancelled';
  const dispatchNote = (r.supplierDispatchDate || '').toLowerCase();
  const whNote = (r.receivedInWH || '').toLowerCase();
  const sold = (r.soldOut || '').toLowerCase().includes('yes') || /sold|stock ?out/.test(dispatchNote) || /sold|stock ?out/.test(whNote);
  if (sold) return 'STOCK OUT';
  const insVal = (r.instock || '').toLowerCase();
  const ins = insVal === 'yes' || insVal.includes('full') || (!r.supplier && (r.receivedInWH || '').toLowerCase() === 'yes');
  if (ins) return 'INSTOCK';
  if (r.supplierStatus) return r.supplierStatus;
  if (r.supplier && r.supplierDispatchDate) return 'Dispatched';
  if (r.supplier && !r.supplierDispatchDate) return 'Not Dispatched';
  return '';
}

// Helper: Find value from multiple possible header mappings
function findVal(r: any, subs: string[]): string {
  for (const s of subs) { if (r[s] !== undefined && r[s] !== null) return String(r[s]); }
  for (const k of Object.keys(r)) {
    const kl = k.toLowerCase();
    if (subs.some(s => kl.includes(s))) return String(r[k]);
  }
  return '';
}

// Helper: Truthy values check
function isTruthy(v: any): boolean {
  return /^(y|yes|true|1|received|done|x|checked|✓|paid)$/i.test((v || '').toString().trim());
}

// Helper: SKU Code extractor
function skuCode(s: string): string {
  return (s || '').split('|')[0].trim().toUpperCase();
}

export async function POST(req: Request) {
  try {
    await dbConnect();
    const body = await req.json();
    const { type, chunk, jobId, startIndex, forceImport } = body;

    if (!type || !Array.isArray(chunk) || chunk.length === 0 || !jobId) {
      return NextResponse.json({ success: false, error: 'Invalid batch configuration parameters' }, { status: 400 });
    }

    let added = 0;
    let updated = 0;
    const startIdx = typeof startIndex === 'number' ? startIndex : 0;
    const logsToInsert: any[] = [];
    let indexInChunk = 0;

    if (type === 'orders' || type === 'at') {
      const ops = chunk.map(r => {
        const rowIndex = startIdx + indexInChunk + 2;
        indexInChunk++;

        const orderNo = (r.orderNo || '').trim();
        const sku = (r.sku || '').trim();

        if (!orderNo && !sku) {
          logsToInsert.push({
            jobId,
            importType: 'orders',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Empty row: missing both orderNo and SKU'
          });
          return null;
        }
        if (!orderNo) {
          logsToInsert.push({
            jobId,
            importType: 'orders',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: orderNo is empty or missing'
          });
          return null;
        }
        if (!sku) {
          logsToInsert.push({
            jobId,
            importType: 'orders',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: SKU code is empty or missing'
          });
          return null;
        }
        
        const setFields = {
          orderDate: r.orderDate || '',
          store: (r.store || '').trim(),
          customer: r.customer || '',
          mobile: r.mobile || r.mobile2 || '',
          city: r.city || '',
          orderValue: r.orderValue || '',
          noItems: r.noItems || '',
          orderStatus: r.orderStatus || ''
        };
        
        return {
          updateOne: {
            filter: { orderNo, sku },
            update: { $set: setFields },
            upsert: true
          }
        };
      }).filter(Boolean) as any[];

      if (ops.length > 0) {
        const bulkRes = await Order.bulkWrite(ops);
        added = bulkRes.upsertedCount;
        updated = bulkRes.modifiedCount;
      }

    } else if (type === 'suppliers' || type === 'sup') {
      const refs = chunk.map((r: any) => (r.orderNo || '').trim()).filter(Boolean);
      const [existingSuppliers, coreOrders] = await Promise.all([
        SupplierImport.find({ orderNo: { $in: refs } }),
        Order.find({ orderNo: { $in: refs } })
      ]);
      const existingMap = new Map(existingSuppliers.map(s => [s.orderNo, s]));
      const coreOrderNos = new Set(coreOrders.map(o => o.orderNo));

      const ops = [];
      for (const r of chunk) {
        const rowIndex = startIdx + indexInChunk + 2;
        indexInChunk++;

        const ref = (r.orderNo || '').trim();
        if (!ref) {
          logsToInsert.push({
            jobId,
            importType: 'suppliers',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: Order number is empty or missing'
          });
          continue;
        }
        if (!coreOrderNos.has(ref)) {
          if (forceImport) {
            await Order.updateOne(
              { orderNo: ref },
              {
                $setOnInsert: {
                  orderNo: ref,
                  sku: r.sku || 'UNKNOWN',
                  orderDate: '',
                  store: '',
                  customer: 'Auto-Created Placeholder',
                  mobile: '',
                  city: '',
                  orderValue: '0',
                  noItems: '0',
                  orderStatus: 'Pending Import'
                }
              },
              { upsert: true }
            );
            coreOrderNos.add(ref);
          } else {
            logsToInsert.push({
              jobId,
              importType: 'suppliers',
              rowNumber: rowIndex,
              rawRowData: r,
              reason: `Skipped: Order number "${ref}" not found in database (import Airtable orders first)`
            });
            continue;
          }
        }

        const supStatus = deriveSupStatus(r);
        const existing = existingMap.get(ref);
        const skuDetailsMap = existing && existing.skuDetails instanceof Map 
          ? existing.skuDetails 
          : new Map(existing && existing.skuDetails ? Object.entries(existing.skuDetails) : []);
        
        if (r.sku) {
          const code = skuCode(r.sku);
          const prev: any = skuDetailsMap.get(code) || {};
          skuDetailsMap.set(code, {
            supplier: r.supplier || prev.supplier || '',
            instock: r.instock || prev.instock || '',
            supplierStatus: supStatus || prev.supplierStatus || '',
            supplierDispatchDate: r.supplierDispatchDate || prev.supplierDispatchDate || '',
            receivedInWH: supStatus === 'STOCK OUT' ? 'Yes' : (r.receivedInWH || prev.receivedInWH || '')
          });
        }

        const detailsList = Array.from(skuDetailsMap.values()) as any[];
        let supplierStatus = supStatus || (existing ? existing.supplierStatus : 'Not Dispatched');
        let supplierDispatchDate = r.supplierDispatchDate || (existing ? existing.supplierDispatchDate : '');
        let receivedInWH = (supStatus === 'STOCK OUT' ? 'Yes' : r.receivedInWH) || (existing ? existing.receivedInWH : '');

        if (detailsList.length > 0) {
          const dates = detailsList.map((d: any) => (d.supplierDispatchDate || '').trim());
          const allDispatched = dates.every(Boolean);
          const allWH = detailsList.every((d: any) => d.receivedInWH === 'Yes');
          
          if (detailsList.every((d: any) => d.supplierStatus === 'STOCK OUT')) {
            supplierStatus = 'STOCK OUT';
          } else if (allDispatched) {
            supplierStatus = 'Dispatched';
            supplierDispatchDate = dates[dates.length - 1];
          } else if (dates.some(Boolean)) {
            supplierStatus = 'Partially dispatched';
          }
          if (allWH) {
            receivedInWH = 'Yes';
          }
        }

        ops.push({
          updateOne: {
            filter: { orderNo: ref },
            update: {
              $set: {
                supplier: r.supplier || (existing ? existing.supplier : ''),
                supplierStatus,
                supplierDispatchDate,
                receivedInWH,
                skuDetails: skuDetailsMap
              }
            },
            upsert: true
          }
        });
      }

      if (ops.length > 0) {
        const bulkRes = await SupplierImport.bulkWrite(ops);
        added = bulkRes.upsertedCount;
        updated = bulkRes.modifiedCount;
      }

    } else if (type === 'ewe' || type === 'tfm') {
      const refs = chunk.map((r: any) => {
        let ref = '';
        if (type === 'ewe') {
          ref = findVal(r, ['referencenumber', 'reference']);
        } else {
          ref = findVal(r, ['shipperref', 'shipper']);
        }
        return (ref || '').trim();
      }).filter(Boolean);

      const coreOrders = await Order.find({ orderNo: { $in: refs } });
      const coreOrderNos = new Set(coreOrders.map(o => o.orderNo));

      const ops = [];
      for (const r of chunk) {
        const rowIndex = startIdx + indexInChunk + 2;
        indexInChunk++;

        let ref = '', status = '', tracking = '', date = '';
        if (type === 'ewe') {
          ref = findVal(r, ['referencenumber', 'reference']);
          status = findVal(r, ['statusname', 'status']);
          tracking = findVal(r, ['barcode']) || (r._raw && r._raw[0] ? r._raw[0].trim() : '');
          date = findVal(r, ['servicedate', 'createddate', 'modifieddate', 'date']);
        } else {
          ref = findVal(r, ['shipperref', 'shipper']);
          status = findVal(r, ['status']);
          tracking = findVal(r, ['awb']);
          date = findVal(r, ['latestofddate', 'firstofddate', 'created', 'date']);
        }
        ref = (ref || '').trim();

        if (!ref) {
          logsToInsert.push({
            jobId,
            importType: type,
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: Courier reference number is empty or missing'
          });
          continue;
        }
        if (!coreOrderNos.has(ref)) {
          if (forceImport) {
            await Order.updateOne(
              { orderNo: ref },
              {
                $setOnInsert: {
                  orderNo: ref,
                  sku: 'UNKNOWN',
                  orderDate: '',
                  store: '',
                  customer: 'Auto-Created Placeholder',
                  mobile: '',
                  city: '',
                  orderValue: '0',
                  noItems: '0',
                  orderStatus: 'Pending Import'
                }
              },
              { upsert: true }
            );
            coreOrderNos.add(ref);
          } else {
            logsToInsert.push({
              jobId,
              importType: type,
              rowNumber: rowIndex,
              rawRowData: r,
              reason: `Skipped: Reference "${ref}" does not exist in the database (import Airtable orders first)`
            });
            continue;
          }
        }

        const replacedRaw = findVal(r, ['replaced', 'replacement']);
        const isReplaced = isTruthy(replacedRaw);
        const lastComment = type === 'tfm' ? (r.tfmLastComment || findVal(r, ['lastcomment']) || '') : '';

        const setFields: any = {};
        if (type === 'ewe') {
          setFields.courier = 'EWE';
          if (status) setFields.courierStatusEWE = status;
          if (tracking) setFields.trackingEWE = tracking;
          if (date) setFields.dispatchDateEWE = date;
        } else {
          setFields.tfmCourier = 'TFM';
          if (status) setFields.courierStatusTFM = status;
          if (tracking) setFields.trackingTFM = tracking;
          if (date) setFields.dispatchDateTFM = date;
          if (lastComment) setFields.lastComment = lastComment;
        }

        if (isReplaced) {
          setFields.replaced = 'Yes';
          setFields.returnReceived = 'Pending';
        }

        ops.push({
          updateOne: {
            filter: { orderNo: ref },
            update: { $set: setFields },
            upsert: true
          }
        });
      }

      if (ops.length > 0) {
        const bulkRes = await CourierImport.bulkWrite(ops);
        added = bulkRes.upsertedCount;
        updated = bulkRes.modifiedCount;
      }

    } else if (type === 'cod') {
      const refs = chunk.map((r: any) => findVal(r, ['orderno', 'order_no', 'orderNo']).trim()).filter(Boolean);
      const coreOrders = await Order.find({ orderNo: { $in: refs } });
      const coreOrderNos = new Set(coreOrders.map(o => o.orderNo));

      const ops = [];
      for (const r of chunk) {
        const rowIndex = startIdx + indexInChunk + 2;
        indexInChunk++;

        const ref = findVal(r, ['orderno', 'order_no', 'orderNo']).trim();
        if (!ref) {
          logsToInsert.push({
            jobId,
            importType: 'cod',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: Order number is empty or missing'
          });
          continue;
        }
        if (!coreOrderNos.has(ref)) {
          if (forceImport) {
            await Order.updateOne(
              { orderNo: ref },
              {
                $setOnInsert: {
                  orderNo: ref,
                  sku: 'UNKNOWN',
                  orderDate: '',
                  store: '',
                  customer: 'Auto-Created Placeholder',
                  mobile: '',
                  city: '',
                  orderValue: '0',
                  noItems: '0',
                  orderStatus: 'Pending Import'
                }
              },
              { upsert: true }
            );
            coreOrderNos.add(ref);
          } else {
            logsToInsert.push({
              jobId,
              importType: 'cod',
              rowNumber: rowIndex,
              rawRowData: r,
              reason: `Skipped: Order number "${ref}" not found in database (import Airtable orders first)`
            });
            continue;
          }
        }

        const received = isTruthy(findVal(r, ['received']));
        const amount = findVal(r, ['amount', 'codamount', 'cod_amount']);
        ops.push({
          updateOne: {
            filter: { orderNo: ref },
            update: {
              $set: {
                codAmountReceived: amount || '',
                codReceived: received ? 'Received' : 'Pending',
                codStatus: received ? 'Received' : 'Pending'
              }
            },
            upsert: true
          }
        });
      }

      if (ops.length > 0) {
        const bulkRes = await CODImport.bulkWrite(ops);
        added = bulkRes.upsertedCount;
        updated = bulkRes.modifiedCount;
      }

    } else if (type === 'returns' || type === 'ret') {
      const refs = chunk.map((r: any) => findVal(r, ['orderno', 'order_no', 'orderNo']).trim()).filter(Boolean);
      const coreOrders = await Order.find({ orderNo: { $in: refs } });
      const coreOrderNos = new Set(coreOrders.map(o => o.orderNo));

      const ops = [];
      for (const r of chunk) {
        const rowIndex = startIdx + indexInChunk + 2;
        indexInChunk++;

        const ref = findVal(r, ['orderno', 'order_no', 'orderNo']).trim();
        if (!ref) {
          logsToInsert.push({
            jobId,
            importType: 'returns',
            rowNumber: rowIndex,
            rawRowData: r,
            reason: 'Skipped: Order number is empty or missing'
          });
          continue;
        }
        if (!coreOrderNos.has(ref)) {
          if (forceImport) {
            await Order.updateOne(
              { orderNo: ref },
              {
                $setOnInsert: {
                  orderNo: ref,
                  sku: 'UNKNOWN',
                  orderDate: '',
                  store: '',
                  customer: 'Auto-Created Placeholder',
                  mobile: '',
                  city: '',
                  orderValue: '0',
                  noItems: '0',
                  orderStatus: 'Pending Import'
                }
              },
              { upsert: true }
            );
            coreOrderNos.add(ref);
          } else {
            logsToInsert.push({
              jobId,
              importType: 'returns',
              rowNumber: rowIndex,
              rawRowData: r,
              reason: `Skipped: Order number "${ref}" not found in database (import Airtable orders first)`
            });
            continue;
          }
        }

        const received = isTruthy(findVal(r, ['received']));
        ops.push({
          updateOne: {
            filter: { orderNo: ref },
            update: {
              $set: {
                returnReceived: received ? 'Received' : 'Not Received',
                replaced: 'Yes'
              }
            },
            upsert: true
          }
        });
      }

      if (ops.length > 0) {
        const bulkRes = await ReturnImport.bulkWrite(ops);
        added = bulkRes.upsertedCount;
        updated = bulkRes.modifiedCount;
      }
    }

    if (logsToInsert.length > 0) {
      await ImportLog.insertMany(logsToInsert);
    }

    return NextResponse.json({
      success: true,
      added,
      updated
    });

  } catch (error: any) {
    console.error('[API/Orders/Chunk] Error processing batch:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
