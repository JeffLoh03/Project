import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { catchError, tap } from 'rxjs/operators';
import { Consultation } from './consultation.model';

@Injectable({
  providedIn: 'root'
})

// Enables the service to send HTTP requests through the http instance
export class ConsultationService {
  private apiUrl = '/api/consultations';  // Using proxy configuration

  constructor(private http: HttpClient) {}

  //  Error handling methods
  private handleError(error: HttpErrorResponse) {
    console.error('An error occurred:', error);
    if (error.error instanceof ErrorEvent) {
      // Client-side error (like network issues)
      console.error('Client error:', error.error.message);
    } else {
      // Backend error (like 404)
      console.error(
        `Backend returned code ${error.status}, ` +
        `body was: ${JSON.stringify(error.error)}`);
    }
    return throwError(() => new Error('Something went wrong; please try again later.'));
  }

  // Retrieve data from database
  getConsultations(): Observable<Consultation[]> {
    console.log('Fetching consultations from:', this.apiUrl);
    return this.http.get<Consultation[]>(this.apiUrl)
      .pipe(
        // Function of 'tap' is obtaining the consultation record, the record will be printed to the console.
        tap(consultations => console.log('Fetched consultations:', consultations)),
        catchError(this.handleError)
      );
  }

  // Add new consultation records to database
  addConsultation(consultation: Consultation): Observable<Consultation> {
    console.log('Adding consultation:', consultation);
    return this.http.post<Consultation>(this.apiUrl, consultation)
      .pipe(
        // Function of 'tap' is to add the new consultation records
        tap(newConsultation => console.log('Added consultation:', newConsultation)),
        catchError(this.handleError)
      );
  }

  // Update the current existing consultation records
  updateConsultation(id: string, consultation: Consultation): Observable<Consultation> {
    console.log('Updating consultation:', id, consultation);
    // 'PUT' request the updating consultation record for a specific ID.
    return this.http.put<Consultation>(`${this.apiUrl}/${id}`, consultation)
      .pipe(
        // Function of 'tap' is to record the updated consultation record
        tap(updatedConsultation => console.log('Updated consultation:', updatedConsultation)),
        catchError(this.handleError)
      );
  }

  // Delete the consultation records
  deleteConsultation(id: string): Observable<void> {
    console.log('Deleting consultation:', id);
    return this.http.delete<void>(`${this.apiUrl}/${id}`)
      .pipe(
        // The function of 'tap' is to record the deleting consultation id
        tap(() => console.log('Deleted consultation:', id)),
        catchError(this.handleError)
      );
  }
}
