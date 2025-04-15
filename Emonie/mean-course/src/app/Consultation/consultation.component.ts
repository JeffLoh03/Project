import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ConsultationService } from './consultation.service';
import { Consultation } from './consultation.model';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
// Import the material from angular
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
// Import different icon, like edit, delete, save, close, add.
import { faPencil,faTrashCan,faSave,faTimes,faPlus,faCheck } from '@fortawesome/free-solid-svg-icons';

// Create the interface for doctor
interface Doctor {
  name: string;
  specialty: string;
  // [] means the string array which can include many slot
  availability: string[];
  timeSlots: string[];
}

@Component({
  selector: 'app-consultation',
  // The component can works independently, so can be directly imported and used by other components
  standalone: true,
  imports: [
    CommonModule, 
    FormsModule, 
    FontAwesomeModule,
    MatInputModule,
    MatFormFieldModule,
    MatButtonModule,
    MatIconModule
  ],
  providers: [ConsultationService],
  templateUrl: './consultation.component.html',
  styleUrls: ['./consultation.component.css']
})
export class ConsultationComponent implements OnInit {
  // Edit icon
  faPencil = faPencil;
  // Delete icon
  faTrashCan = faTrashCan;
  // Save icon
  faSave = faSave;
  // Close icon
  faTimes = faTimes;
  // Add icon
  faPlus = faPlus;
  // Confirm icon
  faCheck = faCheck;

  // Create a consultation array to store the consultation deatils
  consultations: Consultation[] = [];
  // Stores form data for new consultation records, new data will replace the original value
  newConsultation: Consultation = {
    patientName: '',
    doctorName: '',
    date: new Date(),
    timeSlot: '',
    notes: '',
    consultationType: 'Regular'
  };
  // Clear error messages
  errorMessage: string = '';
  // The form is hidden initially when false
  showForm: boolean = false;
  // Indicates the date selected by the user, the initial value is the current date
  selectedDate: Date = new Date();
  availableTimeSlots: string[] = [];
  editingConsultation: Consultation | null = null;

  // Create the doctor array with the available time slot and date.
  doctors: Doctor[] = [
    {
      name: 'Dr. Sarah Johnson',
      specialty: 'Psychiatrist',
      availability: ['Monday', 'Wednesday', 'Friday'],
      timeSlots: ['09:00 AM', '10:00 AM', '11:00 AM', '02:00 PM', '03:00 PM', '04:00 PM']
    },
    {
      name: 'Dr. Michael Chen',
      specialty: 'Clinical Psychologist',
      availability: ['Tuesday', 'Thursday'],
      timeSlots: ['09:00 AM', '10:00 AM', '11:00 AM', '02:00 PM', '03:00 PM']
    },
    {
      name: 'Dr. Emily Williams',
      specialty: 'Therapist',
      availability: ['Monday', 'Tuesday', 'Thursday'],
      timeSlots: ['10:00 AM', '11:00 AM', '02:00 PM', '03:00 PM', '04:00 PM']
    },
    {
      name: 'Dr. James Wilson',
      specialty: 'Counselor',
      availability: ['Wednesday', 'Friday'],
      timeSlots: ['09:00 AM', '10:00 AM', '11:00 AM', '02:00 PM', '03:00 PM']
    }
  ];

  // Provide different types of consultation type
  consultationTypes = [
    'Regular',
    'Emergency',
    'Follow-up',
    'Initial Assessment'
  ];

  // Transmit the data into private method, to achieve interaction with backend performs 
  constructor(private consultationService: ConsultationService) {}

  // when component initialize, it will automatically reload the consultation records
  ngOnInit(): void {
    this.loadConsultations();
  }

  loadConsultations(): void {
    // Retrieve data from backend
    this.consultationService.getConsultations().subscribe({
      // The retrieved data will replace the preset data for presenting
      next: (consultations) => {
        this.consultations = consultations;
        this.errorMessage = '';
      },
      // If happen error, it will show notification
      error: (error) => {
        console.error('Error loading consultations:', error);
        this.errorMessage = 'Failed to load consultations. Please try again.';
      }
    });
  }

  // Update the relevant doctors information from the current consultation record
  selectDoctor(doctor: Doctor): void {
    this.newConsultation.doctorName = doctor.name;
    this.updateAvailableTimeSlots();
  }

  // Update the date
  onDateChange(): void {
    this.updateAvailableTimeSlots();
  }

  // Update the available time slot
  updateAvailableTimeSlots(): void {
    const consultation = this.editingConsultation || this.newConsultation;
    // Checking the name of doctor is same with the doctor name in the array
    const selectedDoctor = this.doctors.find(d => d.name === consultation.doctorName);
    // If not the doctor time slot will still empty
    if (!selectedDoctor) {
      this.availableTimeSlots = [];
      return;
    }

    // **Call getDayOfWeek() to get the day of the week corresponding to the selected date.
    const selectedDay = this.getDayOfWeek(consultation.date);

    // **If the available date is not equal the selected date, it will not allow the users to select
    if (!selectedDoctor.availability.includes(selectedDay)) {
      this.availableTimeSlots = [];
      this.errorMessage = `${selectedDoctor.name} is not available on ${selectedDay}s. Please select a different date.`;
      return;
    }

    // If the date is free, it will update the timeslot of doctors.
    this.errorMessage = '';
    this.availableTimeSlots = selectedDoctor.timeSlots;
  }

  // Verify the validity of consultation data
  validateConsultation(consultation: Consultation): boolean {
    if (!consultation.patientName.trim()) {
      this.errorMessage = 'Please enter patient name.';
      return false;
    }
    if (!consultation.doctorName) {
      this.errorMessage = 'Please select a doctor.';
      return false;
    }
    if (!consultation.timeSlot) {
      this.errorMessage = 'Please select a time slot.';
      return false;
    }
    // Check the selected date is available date or not
    if (!this.isDateAvailable(consultation.date, consultation.doctorName)) {
      this.errorMessage = 'Selected date is not available for this doctor.';
      return false;
    }
    return true;
  }

  // Check if the doctor is available on the selected date.
  isDateAvailable(date: Date, doctorName: string): boolean {
    const selectedDoctor = this.doctors.find(d => d.name === doctorName);
    if (!selectedDoctor) return false;
    return selectedDoctor.availability.includes(this.getDayOfWeek(date));
  }

  // Add the new consultation records
  addConsultation(): void {
    if (!this.validateConsultation(this.newConsultation)) {
      return;
    }

    this.consultationService.addConsultation(this.newConsultation).subscribe({
      next: () => {
        // Reload the consultations record
        this.loadConsultations();
        // Reset the consultation form
        this.resetForm();
        this.errorMessage = '';
      },
      error: (error) => {
        console.error('Error adding consultation:', error);
        this.errorMessage = 'Failed to add consultation. Please try again.';
      }
    });
  }

  // Edit the existing consultation records
  startEdit(consultation: Consultation): void {
    this.editingConsultation = { ...consultation };
    // showForm = false means hidden, true is visible
    this.showForm = false;
    this.updateAvailableTimeSlots();
  }

  // Cancel the edition
  cancelEdit(): void {
    this.editingConsultation = null;
  }

  // Save the editing records
  saveEdit(): void {
    // Check the editing consultation have significant id or not
    if (!this.editingConsultation || !this.editingConsultation._id) {
      return;
    }

    if (!this.validateConsultation(this.editingConsultation)) {
      return;
    }

    // Update the back end records
    this.consultationService.updateConsultation(this.editingConsultation._id, this.editingConsultation).subscribe({
      next: () => {
        this.loadConsultations();
        this.editingConsultation = null;
        this.errorMessage = '';
      },
      // If fail, show the error
      error: (error) => {
        console.error('Error updating consultation:', error);
        this.errorMessage = 'Failed to update consultation. Please try again.';
      }
    });
  }

  // Delete consultation
  deleteConsultation(id: string | undefined): void {
    // Check the id is existing or not
    if (!id) {
      this.errorMessage = 'Cannot delete consultation: Invalid ID';
      return;
    }
    
    // Error prevention
    if (confirm('Are you sure you want to delete this consultation?')) {
      // Delete the data in the back-end and database
      this.consultationService.deleteConsultation(id).subscribe({
        next: () => {
          this.consultations = this.consultations.filter(c => c._id !== id);
          this.errorMessage = '';
        },
        // Show the error if cannot successfully delete the data
        error: (error) => {
          console.error('Error deleting consultation:', error);
          this.errorMessage = 'Failed to delete consultation. Please try again.';
        }
      });
    }
  }

  // Reset the consultation record
  resetForm(): void {
    this.newConsultation = {
      patientName: '',
      doctorName: '',
      date: new Date(),
      timeSlot: '',
      notes: '',
      consultationType: 'Regular'
    };
    this.showForm = false;
  }

  // Accepts a date and returns the corresponding day of the week
  getDayOfWeek(date: Date): string {
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[new Date(date).getDay()];
  }
}
