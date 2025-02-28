# Advanced Web Sheet – Google Sheets Clone

## Project Overview

The **Advanced Web Sheet – Google Sheets Clone** is a web-based spreadsheet application designed to closely mimic the functionality and appearance of Google Sheets. This project challenges you to develop a feature-rich, interactive spreadsheet that supports:

- **Dynamic Grid Generation:** Editable cells with dynamically generated rows and columns.
- **Real-Time Formula Evaluation:** Functions such as `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, and `MEDIAN` with cell dependency tracking.
- **Cell Formatting:** Formatting options including bold, italic, underline, adjustable font size, and color selection.
- **Data Quality Functions:** Functions like `TRIM`, `UPPER`, and `LOWER` to maintain clean data entries.
- **File Operations:** Saving and loading of sheet data via `localStorage` along with the ability to import Excel files using SheetJS.
- **Data Visualization:** Integration with Chart.js to dynamically generate charts based on the spreadsheet data.
- **Bonus Features:** Dark mode support, absolute referencing in formulas, and additional functions for an enhanced user experience.

## Getting Started

### Clone or Download the Repository:

To get started, clone or download the repository from GitHub:

```bash

git clone https://github.com/GTK-THANGELLA-17/spreadsheet-app
Run the Application:
Open the Application:
Open the index.html file directly in your web browser.
Alternatively, you can launch a local web server (using tools like VSCode Live Server or Python's http.server) and navigate to the provided URL.
Using the Application:
Editing and Formatting: Use the toolbar to add or delete rows/columns, format cells, and access the formula palette.
Entering Formulas: Type formulas (e.g., =SUM(A1:A5)) into the formula bar and press Enter to update the selected cell.
Saving and Loading Data: Click Save to store data in localStorage and Load to retrieve it.
Excel File Upload: Click the Upload Excel button, select an Excel file (.xlsx or .xls), and the data will populate the grid.
Dark Mode: Toggle dark mode by clicking the Dark Mode button.
Data Visualization: Observe the dynamically updating chart based on data in Column A.
Technical Details
Frontend Technologies:
HTML5: Provides the structure of the web application.
CSS3: Used for styling and layout of the spreadsheet and user interface.
JavaScript (ES6): Implements the logic for grid generation, formula evaluation, file handling, and dynamic updates.
Libraries and Frameworks:
Bootstrap 5: For responsive design and UI components.
Bootstrap Icons: For toolbar and navigation icons.
SweetAlert2: For enhanced alert messages.
SheetJS (XLSX): To import and parse Excel files, converting Excel data into the spreadsheet format.
Chart.js: For rendering interactive charts based on data in the spreadsheet.
Core Functionalities:
Grid Generation: Dynamic creation of rows and columns with unique cell identification. Users can edit and manipulate the data directly in the grid.
Formula Evaluation: Real-time calculation of formulas, including support for absolute references (e.g., $A$1).
Cell Formatting: The ability to apply bold, italic, underline styles, adjust font sizes, and apply color selections.
Data Persistence: Uses localStorage for saving and loading sheet data, allowing users to persist their data locally without a server.
Excel File Upload: Users can upload an Excel file (.xlsx or .xls) which is converted into the application’s internal data structure.
Chart Visualization: As data in Column A is updated, a live chart is dynamically rendered to represent the changes.
Project Structure
bash
Copy
Advanced-Web-Sheet/
├── index.html         # Main HTML file with the application layout
├── styles.css         # Custom CSS styles including dark mode support
├── script.js          # JavaScript code implementing the spreadsheet functionality
└── README.md          # This documentation file
index.html: Contains the overall layout, including the navigation bar, toolbar, formula bar, spreadsheet table, chart canvas, and the formula palette modal.
styles.css: Provides the styling for the page, spreadsheet, dark mode, and interactive elements such as the fill handle and resize handle.
script.js: Implements dynamic grid generation, formula parsing, cell formatting, data persistence, Excel file upload, chart updating, and bonus features (dark mode, absolute references).
Bonus Features
Excel Upload Capability: Utilizes SheetJS to allow users to import Excel files, automatically converting data into the spreadsheet format.
Data Visualization: Integrates Chart.js to generate and update a live chart based on data in Column A.
Dark Mode: Implements a dark mode toggle to switch the UI theme for enhanced usability in low-light conditions.
Enhanced Formula Support: Adds support for additional functions such as MEDIAN and handles absolute references using $ signs (e.g., $A$1).
Design Decisions
Modular Architecture: Separation of HTML, CSS, and JavaScript enhances maintainability and scalability of the codebase.
User-Centered Interface: The UI design mimics Google Sheets to ensure a familiar and intuitive user experience.
Leveraging Established Libraries: Using Bootstrap, SweetAlert2, SheetJS, and Chart.js minimizes development overhead and maximizes reliability.
Local Data Storage: Utilizing localStorage allows for simple data persistence without the need for server-side infrastructure.
Conclusion
The Advanced Web Sheet – Google Sheets Clone project demonstrates a robust, interactive spreadsheet application built using modern web technologies. With dynamic grid generation, real-time formula evaluation, comprehensive cell formatting, Excel file upload, and live data visualization, this project not only fulfills the assignment requirements but also incorporates several bonus features that enhance its functionality and user experience. The thoughtful design decisions ensure that the application is modular, scalable, and user-friendly, setting a solid foundation for further enhancements such as advanced formula parsing, collaborative editing, or cloud integration.

Candidate Name: Gadidamalla Thangella
Date & Time: 28th February 2025

Key Changes and Additions:
Technologies Used: Detailed information on the frontend technologies (HTML5, CSS3, JavaScript) and libraries (Bootstrap, SweetAlert2, SheetJS, Chart.js) are outlined.
How to Run the Application: Instructions for cloning the repository and running the application locally.
Detailed Core Functionalities: Added detailed descriptions of all core features like grid generation, formula evaluation, cell formatting, Excel file upload, and chart visualization.
Project Structure: Clear structure showing how the project is organized in terms of files.
Bonus Features: Listed all bonus features implemented in the project.
Design Decisions: Explained the rationale behind architectural choices.
