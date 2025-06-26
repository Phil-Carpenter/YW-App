AACT Customer Shipment Overview
This Streamlit application provides a quick and interactive overview of AACT customer shipments based on local Excel data files. It allows users to visualize shipment volumes over a selected date range, filter by part type and customer, and export the data with applied styling.

Table of Contents
Features

Local Setup

Prerequisites

Installation

Running the Application

Data Requirements

Part Key File (PartKey.xlsx)

Monthly Data Files

App Usage

Styling Conventions

Known Limitations & Future Considerations

Features
Dynamic Shipment Overview: Displays a consolidated view of customer shipment volumes over a configurable date range.

Part Key Integration: Merges shipment data with a master part key for detailed descriptions, part types, and customer information.

Filtering: Filter data by "Part Type" and "Customer" via sidebar multiselect options.

Default Date Range: Shows a default 2-week period, with the current day strategically placed near the end for immediate visibility of recent and upcoming data.

Conditional Styling:

Highlights the current day's column in lightgreen.

Highlights weekend columns in lightgrey.

Excel Export: Allows downloading the displayed table data as an Excel file, preserving the conditional coloring.

Data Refresh: A button to clear the data cache and reload all data from the configured local files.

Local Setup
Prerequisites
Python 3.8+ (The code was developed/tested with Python 3.13, but generally compatible with 3.8+)

pip (Python package installer)

Installation
Clone the repository (or download the streamlit_app.py file).

Navigate to the project directory in your terminal or command prompt:

Bash

cd C:\Users\PCarpenter\Desktop\Python # Or wherever you saved the file
Install the required Python packages:

Bash

pip install streamlit pandas openpyxl
Running the Application
Once the prerequisites are installed and you are in the project directory:

Bash

python -m streamlit run streamlit_app.py
This command will open the Streamlit application in your default web browser (usually at http://localhost:8501).

Data Requirements
The application expects two types of Excel files located in specific local directories.

Part Key File (PartKey.xlsx)
Location: Defined directly in the code as C:\Users\PCarpenter\Desktop\Python\PartKey.xlsx.

Filename: Must be named PartKey.xlsx.

Expected Columns (case-insensitive, trimmed whitespace):

Customer Part Number (or variations like customer part #)

Description (or Part Description, Part Desc)

AACT Part Number (or variations containing aact and part number)

Part Type (or variations containing part type)

Customer (or variations containing customer, excluding customer part number)

Data Cleaning:

Customer Part Number and Description columns are stripped of .0 decimal suffixes and whitespace.

Rows with Description values like 'N/A', 'NA', 'NONE' (case-insensitive) are removed.

Duplicate Customer Part Number entries are removed, keeping the last occurrence.

AACT Part Number is explicitly converted to string to avoid display issues.

Monthly Data Files
Location: Defined directly in the code as C:\Users\PCarpenter\Desktop\Python.

Filename Convention: Monthly data files are expected to be named [MonthNumber].xlsx (e.g., 1.xlsx for January, 2.xlsx for February, etc.).

Sheet Naming Convention: Each Excel file is expected to contain sheets named [DayNumber] (e.g., Sheet1 for Jan 1st, Sheet2 for Jan 2nd, Sheet31 for Jan 31st). The app extracts the day number from the sheet name.

Expected Columns within each Sheet (case-sensitive, must match exactly):

Customer Part #

Actual Ship Qty

Header Row: Data is expected to start after 3 header rows, so header=3 is used when reading sheets.

Part Number Normalization: The app attempts to normalize Customer Part # values from these files against the PartKey.xlsx by checking for exact matches, or by adding/removing "00" suffixes (e.g., 123 vs 12300).

App Usage
Review Data Locations: The fixed paths for your Part Key file and Monthly Data directory are displayed at the top of the sidebar.

Select Date Range: Use the "Start date" and "End date" pickers in the sidebar to define the period you wish to view. By default, it shows a 2-week window with the current date visible towards the end.

Apply Filters: Use the "Filter by Part Type" and "Filter by Customer" multiselect boxes to narrow down the displayed data.

Refresh Data: If you update any of your local Excel data files, click the "Refresh Customer Order Data" button at the bottom of the sidebar to clear the cache and reload the latest information.

Download Data: Click the "Download as Excel (with colors)" button at the bottom of the main display area to get an Excel file of the currently displayed table, complete with conditional styling.

Styling Conventions
The table uses visual cues for easy readability:

Today's Date: Columns corresponding to the current day are highlighted in lightgreen.

Weekends: Columns representing Saturday or Sunday are highlighted in lightgrey.

Known Limitations & Future Considerations
Fixed Data Paths: The application currently relies on hardcoded local file paths.

Future Enhancement: Integrate with cloud storage solutions like SharePoint or OneDrive for easier data management and collaboration, removing the need for local file paths and VPN access. This would involve using Python libraries for SharePoint/Microsoft Graph API access and implementing robust authentication.

No Sticky Columns: The st.dataframe component in Streamlit does not natively support sticky or frozen columns for horizontal scrolling.

Future Enhancement: To achieve sticky columns, integration with a more advanced table component like streamlit-aggrid would be necessary. This would require significant code changes as streamlit-aggrid handles its own styling and data presentation.

Multiselect Tag Colors: Due to Streamlit's internal CSS structure, the selected tags in the multiselect filters currently display in the default theme color (which may appear red in some environments) and cannot be easily changed without risking unintended styling side effects on other parts of the application. This has been left as-is to ensure core grid styling functions correctly.







