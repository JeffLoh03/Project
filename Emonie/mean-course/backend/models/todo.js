//server-side application
const mongoose = require('mongoose');
//the schema is define structure of the documents that store in MongoDB
const todoSchema = new mongoose.Schema({
  //title is required when adding a new todo
  title: { type: String, required: true },
  category: { 
    type: String, 
    required: true,
    //enum:Ensures that the category field can only have one of these predefined values.
    enum: ['Mindfulness', 'Exercise', 'Self Care']
  },
  //show the todo is completed or not
  completed: { type: Boolean, default: false },
  dueDate: { type: Date, required: true }
}, {
  //Mongoose will automatically add two additional fields, createdAt and updatedAt, to store the document creation and last update times.
  timestamps: true
});
//create a database call Todo
//exports: export todo and can used in other parts 
module.exports = mongoose.model('Todo', todoSchema); 