# Crayon Costing Application

A web-based costing application for Crayon's License Sales team, replacing Excel-based costing sheets with a centralized platform.

## Features

- **Costing Form**: Enter customer details, line items with Excel paste support
- **Merged Sheet**: Auto-generated financial summary with all calculations
- **Final Price Table**: Professional quotation output
- **Excel Export**: Export all 3 sheets to Excel for client review
- **Region Support**: Middle East and Africa with GP split logic

## Quick Start

### Prerequisites
- Node.js 18+ (https://nodejs.org/)
- npm (comes with Node.js)

### Installation

1. **Extract the zip file** to your desired location

2. **Open terminal/command prompt** and navigate to the folder:
   ```bash
   cd crayon-costing-app
   ```

3. **Install dependencies**:
   ```bash
   npm install
   ```

4. **Start the development server**:
   ```bash
   npm run dev
   ```

5. **Open your browser** and go to:
   ```
   http://localhost:3000
   ```

## Usage Guide

### Loading Sample Data
- Click "Load Sample Data" button to populate with FEWA sample data

### Excel Paste Support
1. Copy rows from your Excel costing sheet
2. Click on the Part Number field in Line Items
3. Paste (Ctrl+V) - data will auto-populate

### Exporting to Excel
- Click the green "Export Excel" button at the bottom
- File will download with 3 sheets: Costing, Merged, FinalPriceTable

### Region Selection
- **Middle East**: Standard GP calculation
- **Africa**: Shows Crayon GP vs Partner GP split columns

## Tabs

1. **Costing Form**: Main data entry
2. **Merged**: Financial summary (auto-generated)
3. **Final Price Table**: Customer quotation (auto-generated)

## Project Structure

```
crayon-costing-app/
├── src/
│   ├── App.jsx          # Main application component
│   ├── main.jsx         # Entry point
│   └── index.css        # Tailwind CSS styles
├── index.html           # HTML template
├── package.json         # Dependencies
├── vite.config.js       # Vite configuration
├── tailwind.config.js   # Tailwind configuration
└── README.md            # This file
```

## Tech Stack

- React 18
- Vite (build tool)
- Tailwind CSS (styling)
- Lucide React (icons)
- SheetJS (Excel export)
- FileSaver.js (file download)

## Troubleshooting

### "npm not found"
Install Node.js from https://nodejs.org/

### Port 3000 in use
Edit `vite.config.js` and change the port number

### Module not found errors
Run `npm install` again

## Support

For issues or questions, contact the development team.

---
Version: 1.0.0
Last Updated: January 2025
