import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { catchError, tap } from 'rxjs/operators';
import { Diary } from './diary.model';

@Injectable({
  providedIn: 'root' // Provides the service at the root level
})
export class DiaryService {
  private apiUrl = '/api/diaries'; // URL for the API endpoint for diaries

  constructor(private http: HttpClient) {} // Injecting HttpClient service

  private handleError(error: HttpErrorResponse) {
    console.error('An error occurred:', error); // Log error
    if (error.error instanceof ErrorEvent) {
      console.error('Client error:', error.error.message); // Log client-side error
    } else {
      console.error(
        `Backend returned code ${error.status}, ` +
        `body was: ${JSON.stringify(error.error)}`); // Log server-side error
    }
    return throwError(() => new Error('Something went wrong; please try again later.')); // Return user-friendly error message
  }

  getDiaries(): Observable<Diary[]> {
    console.log('Fetching diaries from:', this.apiUrl); // Log the API URL being called
    return this.http.get<Diary[]>(this.apiUrl) // Make GET request to fetch diaries
      .pipe(
        tap(diaries => console.log('Fetched diaries:', diaries)), // Log the fetched diaries
        catchError(this.handleError) // Catch errors and handle them
      );
  }

  addDiary(diary: Diary): Observable<Diary> {
    console.log('Adding diary:', diary); // Log the diary being added
    return this.http.post<Diary>(this.apiUrl, diary) // Make POST request to add a new diary
      .pipe(
        tap(newDiary => console.log('Added diary:', newDiary)), // Log the newly added diary
        catchError(this.handleError) // Catch errors and handle them
      );
  }

  updateDiary(id: string, diary: Diary): Observable<Diary> {
    console.log('Updating diary:', id, diary); // Log the diary being updated
    return this.http.put<Diary>(`${this.apiUrl}/${id}`, diary) // Make PUT request to update an existing diary
      .pipe(
        tap(updatedDiary => console.log('Updated diary:', updatedDiary)), // Log the updated diary
        catchError(this.handleError) // Catch errors and handle them
      );
  }

  deleteDiary(id: string): Observable<void> {
    console.log('Deleting diary:', id); // Log the diary being deleted
    return this.http.delete<void>(`${this.apiUrl}/${id}`) // Make DELETE request to remove the diary
      .pipe(
        tap(() => console.log('Deleted diary:', id)), // Log the ID of the deleted diary
        catchError(this.handleError) // Catch errors and handle them
      );
  }
}
