# H.E.D.N.O_auto_app Multiple Choice Excel Data Extractor  
This Python script offers a GUI-based tool to extract and consolidate specific data from multiple Excel sheets based on user-selected choices.  

**Overview**   
The script utilizes the tkinter library to create a graphical user interface (GUI) that allows users to:  

  Select an Excel file to process.  
  Choose specific data criteria from a multiple-choice list.  
  Extract and consolidate the chosen data from multiple sheets into a new Excel file.  

**Features**  
**File Selection:** Users can either manually input the file path or use the browse button to select an Excel file.    
**Multiple-Choice Interface:** Upon file selection, a window pops up displaying multiple-choice options parsed from the Excel file. Users can select/deselect items for data extraction.    
**Data Extraction:** Upon confirming selections, the script extracts the chosen data from multiple sheets of the Excel file, consolidates it, and saves it into a new Excel file named combined_sheets.xlsx.    
**Error Handling:** The script includes basic error handling, notifying users in case of an exception or if no data is available.  


    
**Usage**    
Run the script using Python.    
Input the path to the Excel file or use the browse button to select it.  
The multiple-choice window will display data options parsed from the file's second sheet.  
Select specific data items for extraction or leave all unchecked to extract all data.  
Click "Process Selection" to initiate data extraction and consolidation.  
Upon completion, a message box will confirm the process.  


    
**Requirements**
Python 3.x  
pandas, openpyxl libraries  
  
     
**How to Run**
```bash
python script_name.py
```  
Replace script_name.py with the actual name of the Python script.  

**Notes**  
Used pyinstaller to make it a standalone appication.  
The script assumes the relevant data is available in the second sheet of the provided Excel file.  
Customize start and end range indices as needed for data parsing.  
