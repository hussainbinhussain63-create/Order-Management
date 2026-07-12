import mongoose, { Schema, Document } from 'mongoose';

export interface IOrder extends Document {
  orderDate: string;
  orderNo: string;
  store: string;
  customer: string;
  mobile: string;
  city: string;
  orderValue: string;
  noItems: string;
  sku: string;
  supplier: string;
  instock: string;
  orderStatus: string;
  supplierStatus: string;
  supplierDispatchDate: string;
  receivedInWH: string;
  dispatchFromOffice: string;
  dispatchDateEWE: string;
  dispatchDateTFM: string;
  courier: string;
  tfmCourier: string;
  courierStatusEWE: string;
  courierStatusTFM: string;
  codStatus: string;
  codAmountReceived: string;
  codReceived?: string;
  returnReceived: string;
  replaced: string;
  trackingEWE: string;
  trackingTFM: string;
  skuDetails: Record<string, {
    supplier?: string;
    instock?: string;
    supplierStatus?: string;
    supplierDispatchDate?: string;
    receivedInWH?: string;
  }>;
  comment?: string;
  lastComment?: string;
}

const OrderSchema = new Schema<IOrder>({
  orderDate: { type: String, default: '' },
  orderNo: { type: String, required: true },
  store: { type: String, default: '' },
  customer: { type: String, default: '' },
  mobile: { type: String, default: '' },
  city: { type: String, default: '' },
  orderValue: { type: String, default: '' },
  noItems: { type: String, default: '' },
  sku: { type: String, default: '' },
  supplier: { type: String, default: '' },
  instock: { type: String, default: '' },
  orderStatus: { type: String, default: '' },
  supplierStatus: { type: String, default: '' },
  supplierDispatchDate: { type: String, default: '' },
  receivedInWH: { type: String, default: '' },
  dispatchFromOffice: { type: String, default: '' },
  dispatchDateEWE: { type: String, default: '' },
  dispatchDateTFM: { type: String, default: '' },
  courier: { type: String, default: '' },
  tfmCourier: { type: String, default: '' },
  courierStatusEWE: { type: String, default: '' },
  courierStatusTFM: { type: String, default: '' },
  codStatus: { type: String, default: '' },
  codAmountReceived: { type: String, default: '' },
  codReceived: { type: String, default: '' },
  returnReceived: { type: String, default: '' },
  replaced: { type: String, default: '' },
  trackingEWE: { type: String, default: '' },
  trackingTFM: { type: String, default: '' },
  skuDetails: { type: Map, of: Schema.Types.Mixed, default: {} },
  comment: { type: String, default: '' },
  lastComment: { type: String, default: '' }
}, {
  timestamps: true
});

// Compound unique index on orderNo + sku to identify line items and enable high-speed upserts
OrderSchema.index({ orderNo: 1, sku: 1 }, { unique: true });

export default mongoose.models.Order || mongoose.model<IOrder>('Order', OrderSchema);
