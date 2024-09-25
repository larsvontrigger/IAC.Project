# Production Plan Automation with PyQt and Excel Integration

This project is designed to automate the workflow in a quality audit environment for automotive manufacturing. It uses Python and PyQt to extract daily production plans from Excel sheets and enhance productivity by linking relevant part measurement documents. 

## Overview

In the automotive industry, quality auditors frequently deal with repetitive tasks such as locating production plans, selecting shift-specific parts, and manually opening Excel files to record measurements. This tool automates that process by:
- Extracting daily production plans from a structured Excel sheet.
- Filtering parts based on the selected work shift (morning, afternoon, night).
- Consolidating duplicate parts (e.g., left and right components).
- Providing clickable hyperlinks to the relevant Excel files where part-specific measurement data is stored.
- Supporting multiple Excel sheets, with parts linked to specific worksheets when required.


##  Technology Stack

- **Python**: Core logic for handling data extraction and processing.
- **PyQt**: GUI interface for selecting shifts, displaying parts, and interacting with files.
- **Pandas**: Data manipulation library to work with Excel sheets.
- **OpenPyXL**: For Excel file reading and manipulation.
- **Platform support**: MacOS, Windows.

##  How It Works

1. **Daily Production Plan Extraction**: 
   The tool reads from a daily production plan stored in an Excel sheet. Each project and its associated parts are listed for a given day.
   
2. **Shift-Based Filtering**:
   Users select their shift (morning, afternoon, or night) from a dropdown menu. The program then filters parts relevant to that shift and consolidates them if necessary (e.g., left/right parts).
   
3. **Part-File Mapping**:
   Each part is linked to its specific Excel file for recording measurements. In cases where multiple parts share the same Excel file but different sheets, the corresponding worksheet is also specified in the link.

4. **Hyperlinking**:
   The UI provides clickable hyperlinks that open the relevant Excel file, directing the user to the correct sheet for inputting measurement data.

