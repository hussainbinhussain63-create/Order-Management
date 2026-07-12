import { NextResponse } from 'next/server';
import { dbConnect } from '@/lib/mongodb';
import Order from '@/models/Order';

export async function GET() {
  try {
    await dbConnect();
    console.log('[API/Reports/Dashboard] Fetching orders for aggregate metrics...');
    
    // Select only specific fields to keep the query fast and low memory
    const orders = await Order.find(
      {},
      'orderDate store supplier orderValue courier courierStatusEWE courierStatusTFM codStatus codReceived'
    ).lean();

    console.log(`[API/Reports/Dashboard] Processing aggregates for ${orders.length} orders.`);

    // 1. Monthly Trend
    const trendMap: Record<string, { month: string; orders: number; value: number }> = {};
    orders.forEach(o => {
      const dateStr = o.orderDate;
      if (!dateStr) return;
      
      let monthKey = '';
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        // Format: MM/DD/YYYY or M/D/YYYY
        const year = parts[2].length === 4 ? parts[2] : '20' + parts[2];
        const month = parts[0].padStart(2, '0');
        monthKey = `${year}-${month}`;
      } else {
        const d = new Date(dateStr);
        if (!isNaN(d.getTime())) {
          monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        }
      }
      
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
    orders.forEach(o => {
      const s = o.store || 'Unknown';
      storeMap[s] = (storeMap[s] || 0) + 1;
    });
    const storeDistribution = Object.entries(storeMap).map(([store, count]) => ({ store, count }));

    // 3. Supplier Distribution
    const supplierMap: Record<string, number> = {};
    orders.forEach(o => {
      const s = o.supplier || 'Unassigned';
      supplierMap[s] = (supplierMap[s] || 0) + 1;
    });
    const supplierDistribution = Object.entries(supplierMap).map(([supplier, count]) => ({ supplier, count }));

    // 4. Courier Performance
    let eweDelivered = 0, eweReturned = 0, eweHold = 0, ewePending = 0;
    let tfmDelivered = 0, tfmReturned = 0, tfmHold = 0, tfmPending = 0;
    
    orders.forEach(o => {
      const isDel = (status: string) => /deliver/i.test(status || '');
      const isRet = (status: string) => /return|rto/i.test(status || '');
      const isHoldVal = (status: string) => /hold/i.test(status || '');
      
      if (o.courier === 'EWE' || o.courierStatusEWE) {
        const s = o.courierStatusEWE || '';
        if (isDel(s)) eweDelivered++;
        else if (isRet(s)) eweReturned++;
        else if (isHoldVal(s)) eweHold++;
        else ewePending++;
      }
      if (o.tfmCourier === 'TFM' || o.courierStatusTFM) {
        const s = o.courierStatusTFM || '';
        if (isDel(s)) tfmDelivered++;
        else if (isRet(s)) tfmReturned++;
        else if (isHoldVal(s)) tfmHold++;
        else tfmPending++;
      }
    });
    const courierStats = [
      { name: 'Delivered', EWE: eweDelivered, TFM: tfmDelivered },
      { name: 'Returned', EWE: eweReturned, TFM: tfmReturned },
      { name: 'On Hold', EWE: eweHold, TFM: tfmHold },
      { name: 'Pending', EWE: ewePending, TFM: tfmPending }
    ];

    // 5. COD Cashflow
    let totalReceived = 0;
    let totalPending = 0;
    orders.forEach(o => {
      const amt = parseFloat(o.orderValue) || 0;
      if (o.codStatus === 'Received' || o.codReceived === 'Received') {
        totalReceived += amt;
      } else {
        totalPending += amt;
      }
    });
    const codSummary = { totalReceived, totalPending };

    return NextResponse.json({
      success: true,
      monthlyTrend,
      storeDistribution,
      supplierDistribution,
      courierStats,
      codSummary
    });

  } catch (error: any) {
    console.error('[API/Reports/Dashboard] Aggregate analysis failed:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
