import mongoose, { Schema, Document } from 'mongoose';

export interface ICourierImport extends Document {
  orderNo: string;
  courier: string;
  trackingEWE: string;
  courierStatusEWE: string;
  dispatchDateEWE: string;
  tfmCourier: string;
  trackingTFM: string;
  courierStatusTFM: string;
  dispatchDateTFM: string;
  lastComment: string;
  replaced: string;
  returnReceived: string;
}

const CourierImportSchema = new Schema<ICourierImport>({
  orderNo: { type: String, required: true, index: true },
  courier: { type: String, default: '' },
  trackingEWE: { type: String, default: '' },
  courierStatusEWE: { type: String, default: '' },
  dispatchDateEWE: { type: String, default: '' },
  tfmCourier: { type: String, default: '' },
  trackingTFM: { type: String, default: '' },
  courierStatusTFM: { type: String, default: '' },
  dispatchDateTFM: { type: String, default: '' },
  lastComment: { type: String, default: '' },
  replaced: { type: String, default: '' },
  returnReceived: { type: String, default: '' }
}, {
  timestamps: true
});

CourierImportSchema.index({ orderNo: 1 }, { unique: true });

export default mongoose.models.CourierImport || mongoose.model<ICourierImport>('CourierImport', CourierImportSchema);
