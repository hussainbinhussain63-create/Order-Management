import mongoose, { Schema, Document } from 'mongoose';

export interface ICommentRecord extends Document {
  orderNo: string;
  comment: string;
  lastComment: string;
}

const CommentRecordSchema = new Schema<ICommentRecord>({
  orderNo: { type: String, required: true, index: true },
  comment: { type: String, default: '' },
  lastComment: { type: String, default: '' }
}, {
  timestamps: true
});

CommentRecordSchema.index({ orderNo: 1 }, { unique: true });

export default mongoose.models.CommentRecord || mongoose.model<ICommentRecord>('CommentRecord', CommentRecordSchema);
