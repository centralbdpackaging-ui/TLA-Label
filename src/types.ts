export interface SampleData {
  id: string;
  siNo: number | string;
  samplePo: string;
  poNo: string;
  piNo: string;
  customer: string;
  sampleType: string;
}

export interface CustomerContact {
  customerName: string;
  contactPerson: string;
  phoneNumber: string;
}

export interface LabelData extends SampleData {
  contactPerson?: string;
  phoneNumber?: string;
}
