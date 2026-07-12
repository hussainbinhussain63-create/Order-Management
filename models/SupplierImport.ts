import mongoose, { Schema, Document } from 'mongoose';

export interface ISupplierImport extends Document {
  orderNo: string;
  supplier: string;
  supplierStatus: string;
  supplierDispatchDate: string;
  receivedInWH: string;
  skuDetails: Record<string, {
    supplier?: string;
    instock?: string;
    supplierStatus?: string;
    supplierDispatchDate?: string;
    receivedInWH?: string;
  }>;
}

const SupplierImportSchema = new Schema<ISupplierImport>({
  orderNo: { type: String, required: true, index: true },
  supplier: { type: String, default: '' },
  supplierStatus: { type: String, default: 'Not Dispatched' },
  supplierDispatchDate: { type: String, default: '' },
  receivedInWH: { type: String, default: '' },
  skuDetails: { type: Map, of: Schema.Types.Mixed, default: {} }
}, {
  timestamps: true
});

SupplierImportSchema.index({ orderNo: 1 }, { unique: true });

export default mongoose.models.SupplierImport || mongoose.model<ISupplierImport>('SupplierImport', SupplierImportSchema);
