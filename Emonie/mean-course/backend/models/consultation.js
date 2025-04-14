// This is to store the data type of consultation
const mongoose = require('mongoose');

const consultationSchema = new mongoose.Schema({
  // required: true, means this section is compulsory to be inserted 
  patientName: { type: String, required: true },
  doctorName: { type: String, required: true },
  date: { type: Date, required: true },
  timeSlot: { type: String, required: true },
  notes: { type: String },
  consultationType: { type: String, required: true }
}, {
  // Record the creation time and last update time of the consultation record.
  timestamps: true
});

// this can transmit the data to the MongoDB and do CRUD
module.exports = mongoose.model('Consultation', consultationSchema); 