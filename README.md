# Excel Difference App

This simple Tkinter application compares columns from two Excel reports by a common project name and calculates the difference between selected numeric columns.

## Features
- Choose two Excel files for comparison.
- Select project and value columns from each file.
- Calculate differences and display results.
- Save processed data to a new Excel file.

## Usage
1. Run the application:
   ```bash
   python excel_diff_app.py
   ```
2. Use the **Browse** buttons to select two report files.
3. Choose matching project columns and value columns.
4. Specify the output file and click **Compute Difference**.
5. The resulting differences will be displayed and saved to the specified output file.

## Requirements
- pandas
- openpyxl
- tkinter (standard library)
