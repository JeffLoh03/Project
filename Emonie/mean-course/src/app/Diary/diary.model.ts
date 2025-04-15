export interface Diary {
  _id?: string; // Optional field for the diary's unique identifier (used for updates or deletion)
  title: string; // Title of the diary entry
  content: string; // Content of the diary entry
  date: Date; // Date of the diary entry (should be a valid Date object)
  emotion: 'Joy' | 'Sadness' | 'Anger' | 'Anxiety' | 'Numb'; // Emotion associated with the diary entry, restricted to a set of predefined values
}
 
 