const mongoose = require('mongoose');
 
// Define the schema for the diary entry
const diarySchema = new mongoose.Schema({
  title: { type: String, required: true }, // Title of the diary entry (required)
  content: { type: String, required: true }, // Content of the diary entry (required)
  emotion: { type: String, required: true }, // Emotion associated with the diary entry (required)
  date: { type: Date, default: Date.now } // Date of the diary entry, defaults to the current date and time
}, {
  timestamps: true // Automatically adds createdAt and updatedAt fields
});
 
// Export the Diary model based on the schema
module.exports = mongoose.model('Diary', diarySchema);
 
 