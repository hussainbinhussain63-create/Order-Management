import { NextResponse } from 'next/server';
import { dbConnect } from '@/lib/mongodb';
import ImportJob from '@/models/ImportJob';

export async function GET(request: Request) {
  try {
    await dbConnect();
    const { searchParams } = new URL(request.url);
    const jobId = searchParams.get('jobId');

    if (!jobId) {
      return NextResponse.json({ success: false, error: 'jobId is required' }, { status: 400 });
    }

    console.log(`[API/Import/Status] Checking import job status: ${jobId}`);
    const job = await ImportJob.findOne({ jobId }).lean();

    if (!job) {
      return NextResponse.json({ success: false, error: 'Job not found' }, { status: 404 });
    }

    return NextResponse.json({
      success: true,
      jobId: job.jobId,
      status: job.status,
      progress: job.progress,
      processedRows: job.processedRows,
      totalRows: job.totalRows,
      added: job.added,
      updated: job.updated,
      error: job.error
    });

  } catch (error: any) {
    console.error('[API/Import/Status] Failed checking status:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
