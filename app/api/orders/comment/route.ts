import { NextResponse } from 'next/server';
import { dbConnect } from '@/lib/mongodb';
import CommentRecord from '@/models/CommentRecord';

export async function PUT(req: Request) {
  try {
    await dbConnect();
    const body = await req.json();
    const { orderNo, comment } = body;

    if (!orderNo) {
      return NextResponse.json({ success: false, error: 'Missing orderNo' }, { status: 400 });
    }

    console.log(`[API/Orders/Comment] Updating comment for order: ${orderNo}`);
    
    // Update or insert comments for the specific order number in decoupled Comments collection
    await CommentRecord.updateOne(
      { orderNo },
      { $set: { comment: comment || '' } },
      { upsert: true }
    );

    return NextResponse.json({ success: true, message: 'Comment updated successfully' });

  } catch (error: any) {
    console.error('[API/Orders/Comment] Error updating comment:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
