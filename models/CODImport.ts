import mongoose, { Schema, Document } from 'mongoose';

export interface ICODImport extends Document {
  orderNo: string;
  codStatus: string;
  codAmountReceived: string;
  codReceived: string;
}

const CODImportSchema = new Schema<ICODImport>({
  orderNo: { type: String, required: true, index: true },
  codStatus: { type: String, default: 'Pending' },
  codAmountReceived: { type: String, default: '' },
  codReceived: { type: String, default: '' }
}, {
  timestamps: true
});

CODImportSchema.index({ orderNo: 1 }, { unique: true });

export default mongoose.models.CODImport || mongoose.model<ICODImport>('CODImport', CODImportSchema);
