import mongoose, { Schema, Document } from 'mongoose';

export interface IImportLog extends Document {
  jobId: string;
  importType: string;
  rowNumber: number;
  rawRowData: Record<string, any>;
  reason: string;
}

const ImportLogSchema = new Schema<IImportLog>({
  jobId: { type: String, required: true, index: true },
  importType: { type: String, required: true, index: true },
  rowNumber: { type: Number, required: true },
  rawRowData: { type: Schema.Types.Mixed, default: {} },
  reason: { type: String, required: true }
}, {
  timestamps: true
});

export default mongoose.models.ImportLog || mongoose.model<IImportLog>('ImportLog', ImportLogSchema);
