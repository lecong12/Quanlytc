# AI Coding Assistant Instructions

## Project Overview
This is a Vietnamese family expense tracking web application built with Google Apps Script. It uses Google Sheets as the database and provides a responsive web interface for managing income/expense records.

## Architecture
- **Backend**: Google Apps Script (Code.js) - server-side logic
- **Frontend**: HTML/CSS/JavaScript files served via HtmlService
- **Database**: Google Sheets ("Data" sheet for transactions, "Users" sheet for authentication)
- **Deployment**: Web App accessible via URL (anonymous access configured)

## Core Data Structure
**Transactions Sheet ("Data"):**
- Column A: ID (timestamp-based unique identifier)
- Column B: Date (dd/MM/yyyy format in sheets, yyyy-MM-dd from HTML inputs)
- Column C: Type ("Thu" for income, "Chi" for expense)
- Column D: Content (transaction description)
- Column E: Amount (numeric value)
- Column F: Created timestamp

**Users Sheet ("Users"):**
- Simple authentication with username/password stored in sheet

## Critical Patterns

### Date Handling
```javascript
// Parse multiple date formats
function parseToDate(s){
    if(/^\d{4}-\d{2}-\d{2}$/.test(s)){ // yyyy-MM-dd from HTML input
        const p = s.split('-').map(Number);
        return new Date(p[0], p[1]-1, p[2]);
    }
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)){ // dd/MM/yyyy from sheet
        const p = s.split('/').map(Number);
        return new Date(p[2], p[1]-1, p[0]);
    }
}
```

### Amount Formatting
```javascript
// Vietnamese locale with dot separators
const fmt = n => new Intl.NumberFormat('vi-VN').format(n || 0);
```

### Authentication Flow
- All data modification operations require login
- Authentication state managed client-side with `isLoggedIn` boolean
- UI elements disabled/enabled based on login status
- Login modal auto-fills demo credentials (thuy/1393)

### CRUD Operations
- **Create**: `addRow()` generates timestamp-based ID, sorts sheet by date after insert
- **Read**: `listData()` returns formatted objects with date conversion
- **Update**: `updateRow()` modifies existing row, re-sorts entire data range
- **Delete**: `deleteRow()` removes row by ID

### Filtering Logic
Complex multi-tier filtering in `processFilterLogic()`:
1. Type filter ("Thu"/"Chi"/"Tất cả")
2. Date range filtering (date, month, quarter, year modes)
3. Results sorted by date (newest first)

### Export Functionality
- **Excel**: Client-side using XLSX.js library
- **PDF**: Client-side using html2pdf.js with custom styling
- **Email**: Server-side GmailApp integration with HTML table formatting

## Key Files
- `Code.js`: Server-side Apps Script functions
- `Index.html`: Main UI layout and structure
- `Javascript.html`: Client-side logic (CRUD, filtering, export)
- `Css.html`: Responsive styling with mobile optimizations
- `appsscript.json`: Manifest configuration (V8 runtime, webapp settings)

## Development Workflow
1. Edit files locally in VS Code
2. Deploy via Google Apps Script editor or clasp
3. Test web app URL after deployment
4. Data persists in linked Google Sheet

## Common Patterns
- Use `google.script.run` for client-server communication
- Handle authentication state for UI element visibility
- Format amounts with Vietnamese locale for display
- Parse dates carefully between HTML inputs and sheet storage
- Include login checks before destructive operations
- Use Bootstrap classes for responsive design
- Apply red styling (`tr-chi`) for expense rows

## Authentication Notes
- Demo credentials: username "thuy", password "1393"
- Login modal auto-fills these values for development
- Authentication required for: Add, Edit, Delete, Export, Email operations
- View-only access available without login

## Deployment
- Configure as web app with "Execute as: User deploying"
- Set access to "Anyone, even anonymous" for public access
- Use V8 runtime (configured in appsscript.json)
- Logging enabled to Stackdriver</content>
<parameter name="filePath">c:\Users\129\quanlithuchi\.github\copilot-instructions.md