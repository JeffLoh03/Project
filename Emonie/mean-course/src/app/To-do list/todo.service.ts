import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
//catchError: Catches and handles errors in observable streams.
//tap: Allows you to perform side-effects (like logging) without changing the data.
import { catchError, tap } from 'rxjs/operators';
import { Todo } from './todo.model';
//Registers the service globally (only one instance for the whole app)
@Injectable({
  providedIn: 'root'
})
export class TodoService {
  private apiUrl = '/api/todos';

  constructor(private http: HttpClient) {}

  private handleError(error: HttpErrorResponse) {
    //logs the full error
    console.error('An error occurred:', error);

    if (error.error instanceof ErrorEvent) {
      // Client-side error
      console.error('Client error:', error.error.message);
      return throwError(() => error.error.message);
    } else {
      // Server-side error
      console.error(`Backend returned code ${error.status}, body was:`, error.error);
      return throwError(() => error);
    }
  }
  //get a list of todo from MongoDB
  getTodos(): Observable<Todo[]> {
    //Get request
    return this.http.get<Todo[]>(this.apiUrl)
      .pipe(
        //logs fetched todos
        tap(todos => console.log('Fetched todos:', todos)),
        catchError(this.handleError)
      );
  }
  //send a new todo to MongoDB
  addTodo(todo: Todo): Observable<Todo> {
    console.log('Adding todo:', todo);
    return this.http.post<Todo>(this.apiUrl, todo)
      .pipe(
        tap(newTodo => console.log('Added todo:', newTodo)),
        catchError(this.handleError)
      );
  }
  //sends updated data for a specific todo
  updateTodo(id: string, todo: Todo): Observable<Todo> {
    console.log('Updating todo:', id, todo);
    //put: updates the todo using its ID
    return this.http.put<Todo>(`${this.apiUrl}/${id}`, todo)
      .pipe(
        tap(updatedTodo => console.log('Updated todo:', updatedTodo)),
        catchError(this.handleError)
      );
  }
  //delete request for a todo by ID
  deleteTodo(id: string): Observable<void> {
    console.log('Deleting todo:', id);
    return this.http.delete<void>(`${this.apiUrl}/${id}`)
      .pipe(
        tap(() => console.log('Deleted todo:', id)),
        catchError(this.handleError)
      );
  }
}
