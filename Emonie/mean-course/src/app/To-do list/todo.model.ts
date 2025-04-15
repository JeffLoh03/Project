//client-side application
//declares a reusable TS interface
export interface Todo {
  //? mean unique ID of the todo, assigned by the MongoDB
  _id?: string;
  title: string;
  //Restricts the category to only one of three fixed string values
  category: 'Mindfulness' | 'Exercise' | 'Self Care';
  //user can click on the task, and it will show finished which is true and not yet done is false
  completed: boolean;
  dueDate: Date;
}
