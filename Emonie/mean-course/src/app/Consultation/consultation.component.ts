import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ConsultationService } from './consultation.service';
import { Consultation } from './consultation.model';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
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
  imports: [CommonModule, FormsModule, FontAwesomeModule],
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

  onDateChange(): void {
    this.updateAvailableTimeSlots();
  }

  updateAvailableTimeSlots(): void {
    const consultation = this.editingConsultation || this.newConsultation;
    const selectedDoctor = this.doctors.find(d => d.name === consultation.doctorName);
    if (!selectedDoctor) {
      this.availableTimeSlots = [];
      return;
    }

    const selectedDay = this.getDayOfWeek(consultation.date);
    if (!selectedDoctor.availability.includes(selectedDay)) {
      this.availableTimeSlots = [];
      this.errorMessage = `${selectedDoctor.name} is not available on ${selectedDay}s. Please select a different date.`;
      return;
    }

    this.errorMessage = '';
    this.availableTimeSlots = selectedDoctor.timeSlots;
  }

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
    if (!this.isDateAvailable(consultation.date, consultation.doctorName)) {
      this.errorMessage = 'Selected date is not available for this doctor.';
      return false;
    }
    return true;
  }

  isDateAvailable(date: Date, doctorName: string): boolean {
    const selectedDoctor = this.doctors.find(d => d.name === doctorName);
    if (!selectedDoctor) return false;
    return selectedDoctor.availability.includes(this.getDayOfWeek(date));
  }

  addConsultation(): void {
    if (!this.validateConsultation(this.newConsultation)) {
      return;
    }

    this.consultationService.addConsultation(this.newConsultation).subscribe({
      next: () => {
        this.loadConsultations();
        this.resetForm();
        this.errorMessage = '';
      },
      error: (error) => {
        console.error('Error adding consultation:', error);
        this.errorMessage = 'Failed to add consultation. Please try again.';
      }
    });
  }

  startEdit(consultation: Consultation): void {
    this.editingConsultation = { ...consultation };
    this.showForm = false;
    this.updateAvailableTimeSlots();
  }

  cancelEdit(): void {
    this.editingConsultation = null;
  }

  saveEdit(): void {
    if (!this.editingConsultation || !this.editingConsultation._id) {
      return;
    }

    if (!this.validateConsultation(this.editingConsultation)) {
      return;
    }

    this.consultationService.updateConsultation(this.editingConsultation._id, this.editingConsultation).subscribe({
      next: () => {
        this.loadConsultations();
        this.editingConsultation = null;
        this.errorMessage = '';
      },
      error: (error) => {
        console.error('Error updating consultation:', error);
        this.errorMessage = 'Failed to update consultation. Please try again.';
      }
    });
  }

  deleteConsultation(id: string | undefined): void {
    if (!id) {
      this.errorMessage = 'Cannot delete consultation: Invalid ID';
      return;
    }
    
    if (confirm('Are you sure you want to delete this consultation?')) {
      this.consultationService.deleteConsultation(id).subscribe({
        next: () => {
          this.consultations = this.consultations.filter(c => c._id !== id);
          this.errorMessage = '';
        },
        error: (error) => {
          console.error('Error deleting consultation:', error);
          this.errorMessage = 'Failed to delete consultation. Please try again.';
        }
      });
    }
  }

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

  getDayOfWeek(date: Date): string {
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[new Date(date).getDay()];
  }
}
