import { NextResponse } from 'next/server';
import { dbConnect } from '@/lib/mongodb';
import ImportLog from '@/models/ImportLog';

export async function GET() {
  try {
    await dbConnect();
    console.log('[API/Import/Logs] Fetching skipped logs...');
    // Limit to latest 1000 logs to prevent payload blowups
    const logs = await ImportLog.find({}).sort({ createdAt: -1 }).limit(1000).lean();
    return NextResponse.json({ success: true, logs });
  } catch (error: any) {
    console.error('[API/Import/Logs] Error fetching logs:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
