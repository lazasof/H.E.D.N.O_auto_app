#!/usr/bin/env python
# coding: utf-8

# In[31]:


import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import openpyxl

selected_options = []

def on_ok():
    file_path = entry.get()
    print("Path entered:", file_path)
    open_multiple_choice_window(file_path)
    
def browse_path():
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)  # Clear any existing text in the entry
    entry.insert(tk.END, file_path)

def process_selection(selected_options, file_path, result_df):
    selected_values = [option[0] for option in selected_options if option[1].get() == 1]
    print("Selected values:", selected_values)
    
    # Retrieve index and value for selected options
    selected_indices = []
    for value in selected_values:
        index = result_df[result_df['Value'] == value]['Index'].values[0]
        selected_indices.append((index, value))
    numbers_only = [item[0] for item in selected_indices]
    selected_values= [item[1] for item in selected_indices]
    
    numbers_only_minus_one = [x + 1 for x in numbers_only] 
    print(numbers_only)
    print(selected_values)

    output_file = 'combined_sheets.xlsx'

    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names[1:]

    # Get user input for rows to extract
    
    # Parse the second sheet of the Excel file
    second_sheet_df = xls.parse(1)  # Index 1 refers to the second sheet (0-indexed)

    # Prepend 'Τοποθεσία' to the list of selected values
    selected_values.insert(0, 'Τοποθεσία')

    #print("Selected Values from the First Column of the Second Sheet:")
    print(selected_values)


    # Initialize a list to store data from each sheet
    data_frames = []

    for idx, sheet_name in enumerate(sheet_names):
        df = xls.parse(sheet_name, header=None)
        df.drop([0, 1], axis=1, inplace=True)
        df = df.iloc[numbers_only_minus_one]
        df = df.transpose()
        df.insert(0, 'Sheet Name', sheet_name)  # Insert sheet name as a column
        data_frames.append(df)

    # Concatenate all data frames along rows
    combined_data = pd.concat(data_frames, axis=0, ignore_index=True)
    combined_data.columns = selected_values

    # Remove rows where all columns except the first have null values
    cleaned_df = combined_data.dropna(subset=combined_data.columns[1:], how='all')


    cleaned_df.to_excel(output_file, index=False)
    print('finished succesfully!')
    messagebox.showinfo("Pappas BI", "Finished your request!")


    # Perform actions with the selected indices, values, and path here

def open_multiple_choice_window(file_path):
    choice_window = tk.Toplevel(root)
    choice_window.title("Check nothing to bring all data")


    try:
        second_sheet_df = pd.read_excel(file_path, sheet_name=1)  # Assuming data is in the second sheet

        start_range_1 = 0  # Replace with your start index
        end_range_1 = len(second_sheet_df)   # Replace with your end index

        indices = []
        values = []
        
        first_occurrences = {value: True for value in ['Φυσικοχημική Ανάλυση Ελαίου',
                                                       'Αεριοχρωματογραφική Ανάλυση Διαλυμένων Αερίων', 
                                                       'ΙΔΙΟΤΗΤΑ', 
                                                       'ΠΕΡΙΕΚΤΙΚΟΤΗΤΕΣ ΑΕΡΙΩΝ', 
                                                       'Λόγοι ROGERS']}

        for i in range(start_range_1, end_range_1):
            cell_value = second_sheet_df.iloc[i, 0]  # Access the cell in the first column

            if pd.notnull(cell_value) and (cell_value not in first_occurrences.keys() or first_occurrences[cell_value]):
                indices.append(i)
                values.append(cell_value)

                if cell_value in first_occurrences.keys():
                    first_occurrences[cell_value] = False

        result_df = pd.DataFrame({'Index': indices, 'Value': values})
        values_to_remove = [1, 5, 27, 33,34, 48,49]

        # Deleting rows based on specified values in 'Índex' column
        result_df = result_df[~result_df['Index'].isin(values_to_remove)]
        result_df = result_df.reset_index(drop=True)
        print(result_df)

        canvas = tk.Canvas(choice_window)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(choice_window, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        checkbox_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=checkbox_frame, anchor=tk.NW)

        

        for idx, option in enumerate(result_df['Value']):
            var = tk.IntVar()
            check_btn = tk.Checkbutton(checkbox_frame, text=option, variable=var)
            check_btn.pack(anchor=tk.W)
            selected_options.append((result_df.loc[idx]['Value'], var))  # Store (value, var) tuple
        

        checkbox_frame.update_idletasks()
        canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
        
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind('<Configure>', on_configure)

        def on_button_click():
            global selected_options
            
            # Get the selected options
            updated_options = [(var[0], var[1]) for var in selected_options]
            
            # Check if no checkboxes are selected
            if not any(isinstance(var[1], tk.IntVar) and var[1].get() == 1 for var in updated_options):
                # Treat it as if all checkboxes were selected
                updated_options = [(var[0], tk.IntVar(value=1)) for var in updated_options]
            
            process_selection(updated_options, file_path, result_df)
            # Show a message
        
        button = tk.Button(choice_window, text="Process Selection", command=on_button_click)
        button.pack()

    except Exception as e:
        print("Error:", e)
        options = ["No data available"]

root = tk.Tk()
root.title("Pappas BI")

label = tk.Label(root, text="Enter path:")
label.pack()

entry = tk.Entry(root, width=40)
entry.pack()

browse_button = tk.Button(root, text="Browse", command=browse_path)
browse_button.pack()

ok_button = tk.Button(root, text="OK", command=on_ok)
ok_button.pack()

root.mainloop()


