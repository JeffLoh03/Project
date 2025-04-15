import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DiaryService } from './diary.service';
import { Diary } from './diary.model';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import {
  faSmile,
  faSadTear,
  faAngry,
  faFlushed,
  faMeh,
  faTrash,
  faPencilAlt,
  faSave,
  faTimes,
  faEdit
} from '@fortawesome/free-solid-svg-icons';
 
type Emotion = 'Joy' | 'Sadness' | 'Anger' | 'Anxiety' | 'Numb';
 
@Component({
  selector: 'app-diary',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    FontAwesomeModule,
    MatButtonModule,
    MatIconModule
  ],
  providers: [DiaryService], // Injecting the DiaryService
  templateUrl: './diary.component.html',
  styleUrls: ['./diary.component.css']
})
export class DiaryComponent implements OnInit {
  // Font Awesome icons
  faSmile = faSmile; // Icon for joy emotion
  faSadTear = faSadTear; // Icon for sadness emotion
  faAngry = faAngry; // Icon for anger emotion
  faFlushed = faFlushed; // Icon for anxiety emotion
  faMeh = faMeh; // Icon for numb emotion
  faTrash = faTrash; // Icon for delete action
  faPencilAlt = faPencilAlt; // Icon for write action
  faSave = faSave; // Icon for save action
  faTimes = faTimes; // Icon for cancel action
  faEdit = faEdit; // Icon for edit action
 
  diaries: Diary[] = []; // Array to hold all diaries
  filteredDiaries: Diary[] = []; // Array to hold filtered diaries based on emotion
  showCreateForm = false; // Flag to toggle the display of the create form
  selectedFilter = ''; // Selected emotion filter
  errorMessage = ''; // Holds error message if any
  editingDiary: Diary | null = null; // Holds the diary being edited
 
  emotions: Emotion[] = ['Joy', 'Sadness', 'Anger', 'Anxiety', 'Numb']; // List of emotions available for diary entry
 
  newDiary: Diary = {
    title: '',
    content: '',
    emotion: 'Joy', // Default emotion
    date: new Date() // Default date to current date
  };
 
  constructor(private diaryService: DiaryService) {} // Injecting the DiaryService
 
  ngOnInit(): void {
    this.loadDiaries(); // Load diaries when the component is initialized
  }
 
  loadDiaries(): void {
    // Fetching diaries from the service
    this.diaryService.getDiaries().subscribe({
      next: (diaries) => {
        this.diaries = diaries; // Assign fetched diaries to the diaries array
        this.filterDiaries(); // Filter diaries based on the selected emotion
        this.errorMessage = ''; // Clear any previous error messages
      },
      error: (error) => {
        console.error('Error loading diaries:', error); // Log the error
        this.errorMessage = 'Failed to load diary entries. Please try again.'; // Display error message
      }
    });
  }
 
  filterDiaries(): void {
    // Filtering diaries based on selected emotion filter
    this.filteredDiaries = this.selectedFilter
      ? this.diaries.filter(diary => diary.emotion === this.selectedFilter) // Filter by selected emotion
      : this.diaries; // If no filter, display all diaries
  }
 
  addDiary(): void {
    if (!this.validateDiary(this.newDiary)) { // Validate the new diary before adding
      return;
    }
 
    // Adding new diary entry through the service
    this.diaryService.addDiary(this.newDiary).subscribe({
      next: () => {
        this.loadDiaries(); // Reload diaries after adding the new entry
        this.resetForm(); // Reset the form after successful submission
        this.errorMessage = ''; // Clear any error messages
      },
      error: (error) => {
        console.error('Error adding diary:', error); // Log the error
        this.errorMessage = 'Failed to save diary entry. Please try again.'; // Display error message
      }
    });
  }
 
  validateDiary(diary: Diary): boolean {
    // Validate the diary form fields before adding
    if (!diary.title.trim()) {
      this.errorMessage = 'Please enter a title.'; // Show error if title is empty
      return false;
    }
    if (!diary.content.trim()) {
      this.errorMessage = 'Please enter some content.'; // Show error if content is empty
      return false;
    }
    if (!diary.emotion) {
      this.errorMessage = 'Please select an emotion.'; // Show error if emotion is not selected
      return false;
    }
    return true; // Return true if all fields are valid
  }
 
  resetForm(): void {
    // Reset the newDiary form to initial state
    this.newDiary = {
      title: '',
      content: '',
      emotion: 'Joy', // Reset to default emotion
      date: new Date() // Reset date to current date
    };
    this.showCreateForm = false; // Hide the create form
    this.editingDiary = null; // Clear editing diary
  }
 
  startEdit(diary: Diary): void {
    this.editingDiary = { ...diary }; // Copy the diary being edited
    this.showCreateForm = false; // Hide the create form when editing
  }
 
  cancelEdit(): void {
    this.editingDiary = null; // Clear editing state
  }
 
  saveEdit(): void {
    if (!this.editingDiary || !this.editingDiary._id) {
      return; // Return if no diary or invalid id
    }
 
    if (!this.validateDiary(this.editingDiary)) { // Validate before saving changes
      return;
    }
 
    // Update the edited diary through the service
    this.diaryService.updateDiary(this.editingDiary._id, this.editingDiary).subscribe({
      next: () => {
        this.loadDiaries(); // Reload diaries after saving changes
        this.editingDiary = null; // Clear editing state
        this.errorMessage = ''; // Clear any error messages
      },
      error: (error) => {
        console.error('Error updating diary:', error); // Log the error
        this.errorMessage = 'Failed to update diary entry. Please try again.'; // Display error message
      }
    });
  }
 
  deleteDiary(id: string | undefined): void {
    if (!id) {
      this.errorMessage = 'Cannot delete diary: Invalid ID'; // Show error if ID is invalid
      return;
    }
    
    // Delete the diary entry by its ID
    this.diaryService.deleteDiary(id).subscribe({
      next: () => {
        this.loadDiaries(); // Reload diaries after deletion
        this.errorMessage = ''; // Clear any error messages
      },
      error: (error) => {
        console.error('Error deleting diary:', error); // Log the error
        this.errorMessage = 'Failed to delete diary entry. Please try again.'; // Display error message
      }
    });
  }
 
  getEmotionIcon(emotion: Emotion) {
    // Return the appropriate icon based on the emotion
    switch (emotion) {
      case 'Joy':
        return this.faSmile; // Return smile icon for joy
      case 'Sadness':
        return this.faSadTear; // Return sad tear icon for sadness
      case 'Anger':
        return this.faAngry; // Return angry icon for anger
      case 'Anxiety':
        return this.faFlushed; // Return flushed icon for anxiety
      case 'Numb':
        return this.faMeh; // Return meh icon for numbness
      default:
        return this.faSmile; // Default to smile icon
    }
  }
}
 
 