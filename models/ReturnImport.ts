import mongoose, { Schema, Document } from 'mongoose';

export interface IReturnImport extends Document {
  orderNo: string;
  returnReceived: string;
  replaced: string;
}

const ReturnImportSchema = new Schema<IReturnImport>({
  orderNo: { type: String, required: true, index: true },
  returnReceived: { type: String, default: '' },
  replaced: { type: String, default: '' }
}, {
  timestamps: true
});

ReturnImportSchema.index({ orderNo: 1 }, { unique: true });

export default mongoose.models.ReturnImport || mongoose.model<IReturnImport>('ReturnImport', ReturnImportSchema);
