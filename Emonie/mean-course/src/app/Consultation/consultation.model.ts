// The consultation need to obtain the specific data with relevant data type
export interface Consultation {
  _id?: string;
  patientName: string;
  doctorName: string;
  date: Date;
  timeSlot: string;
  notes: string;
  consultationType: string;
}
