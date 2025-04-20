# Emonie - Personal Wellness Application

A modern web application built with Angular for managing personal wellness tasks, including a To-do list and Diary features.

## Features

- 📝 To-do List Management
- 📅 Task Categories (Mindfulness, Exercise, Self Care)
- 🎯 Quick Action Tasks
- 📊 Task Filtering
- 🎨 Modern UI with Material Design
- ✨ Beautiful Animations and Transitions

## Prerequisites

Before you begin, ensure you have the following installed:
- Node.js (v14.0.0 or higher)
- npm (v6.0.0 or higher)
- Angular CLI (v15.0.0 or higher)

## Installation

1. Clone the repository:
```bash
git clone [your-repository-url]
cd mean-course
```

2. Install dependencies:
```bash
npm install
```

3. Install required Angular Material packages:
```bash
ng add @angular/material
npm install @fortawesome/angular-fontawesome @fortawesome/fontawesome-svg-core @fortawesome/free-solid-svg-icons
```

## Required Imports

### Angular Material Imports
```typescript
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
```

### Font Awesome Imports
```typescript
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { faPlus, faSave, faTimes, faEdit, faTrash } from '@fortawesome/free-solid-svg-icons';
```

### Angular Core Imports
```typescript
import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
```

## Running the Application

1. Start the development server:
```bash
ng serve
```

2. Open your browser and navigate to:
```
http://localhost:4200
```

## Project Structure

```
mean-course/
├── src/
│   ├── app/
│   │   ├── To-do list/
│   │   │   ├── todo.component.ts
│   │   │   ├── todo.component.html
│   │   │   ├── todo.component.css
│   │   │   ├── todo.service.ts
│   │   │   └── todo.model.ts
│   │   └── ...
│   ├── assets/
│   └── ...
├── package.json
└── angular.json
```

## Styling Features

### Button Styles
- Gradient borders
- Hover animations
- Ripple effects
- Custom checkbox design
- Category badges
- Icon animations

### Color Scheme
- Primary: #3498db
- Secondary: #2980b9
- Warning: #E53935
- Background: #f8fafc

## Development Guidelines

1. **Component Structure**
   - Use standalone components
   - Implement proper TypeScript interfaces
   - Follow Angular best practices

2. **Styling**
   - Use CSS custom properties for theming
   - Implement responsive design
   - Follow BEM naming convention

3. **State Management**
   - Use services for data management
   - Implement proper error handling
   - Follow reactive programming patterns

## Troubleshooting

### Common Issues

1. **Angular Material Not Loading**
   - Ensure @angular/material is properly installed
   - Check for proper imports in app.module.ts
   - Verify theme is included in angular.json

2. **Font Awesome Icons Not Showing**
   - Verify FontAwesomeModule is imported
   - Check icon imports in component
   - Ensure proper icon registration

3. **Styling Issues**
   - Clear browser cache
   - Check for CSS specificity conflicts
   - Verify Material theme is properly loaded

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please open an issue in the repository or contact the development team.

---


