# sign-in-metrics-aggregator

**Excel Sheet Compiler and Metric Finder**
A powerful VBA macro to aggregate data from multiple Excel workbooks and calculate key metrics like total entries, resolved/unresolved counts, and average resolution time. This tool is designed to work in conjunction with the Excel Sign-In Sheet with FT# Lookup.

**Features**
**Automated Data Aggregation:**
Reads data from .xlsm sign-in logs in a specified folder.

**Metrics Calculation:**
Tracks total entries, resolved/unresolved counts, and average resolution time.

**User-Friendly Output:**
Displays results in a formatted message box.

**How to Use**
**Prepare Sign-In Logs:**
Ensure you are using the Excel Sign-In Sheet with FT# Lookup to generate consistent .xlsm logs.

**Organize Logs:**
Place all .xlsm files in a folder.

**Update the Folder Path:**
Open the VBA editor (Alt + F11).
Update the folderPath variable in the macro to point to the folder containing your sign-in logs.

**Run the Macro:**
Run the AggregateDataFromWorkbooks macro to aggregate data and calculate metrics.

**Requirements**
**Microsoft Excel 2016 or later**
**Sign-In Logs Generated by the Excel Sign-In Sheet with FT# Lookup:**
- Columns in the logs:
  - Resolved? (Yes/No): Column I
  - Resolution Time: Column J
  - Entry Date: Column B
  - Entry Time: Column C
    
**Example Output**

Total Entries: 100
Resolved (Yes): 70
Unresolved (No): 30
Average Resolution Time: 1 day(s) 3 hour(s) 45 minute(s)

**License**
See the LICENSE file for details.

**Credits**
Created by Joseph Simpson.
