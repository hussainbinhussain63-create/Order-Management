import mongoose, { Schema, Document } from 'mongoose';

export interface IImportJob extends Document {
  jobId: string;
  status: 'pending' | 'processing' | 'completed' | 'failed';
  progress: number;
  processedRows: number;
  totalRows: number;
  added: number;
  updated: number;
  error?: string;
}

const ImportJobSchema = new Schema<IImportJob>({
  jobId: { type: String, required: true, index: true, unique: true },
  status: { type: String, enum: ['pending', 'processing', 'completed', 'failed'], default: 'pending' },
  progress: { type: Number, default: 0 },
  processedRows: { type: Number, default: 0 },
  totalRows: { type: Number, default: 0 },
  added: { type: Number, default: 0 },
  updated: { type: Number, default: 0 },
  error: { type: String, default: '' }
}, {
  timestamps: true
});

export default mongoose.models.ImportJob || mongoose.model<IImportJob>('ImportJob', ImportJobSchema);
