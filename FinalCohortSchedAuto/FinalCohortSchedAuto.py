from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.utils import column_index_from_string
from openpyxl.styles import NamedStyle
from openpyxl.cell.cell import MergedCell
from tkinter import filedialog
import shutil
import tkinter as tk
import os

class App:
    def __init__(self, root):
        self.source_file_path = ""
        self.source_file_name = ""
        
        self.target_file_path = ""
        self.target_file_name = ""
#/--------------------------------GUI------------------------------/
        #setting title
        root.title("Tammy's Cohort Sched Tool")
        #setting window size
        root.geometry("450x180")
        root.resizable(width=False, height=False)

        # Source file label
        source_label = tk.Label(root, text="Calendar:")
        source_label.place(x=30,y=30)

        self.source_file_label = tk.Label(root, text='No File Selected')
        self.source_file_label.place(x=150,y=30)

        # Source file browse button
        source_browse_button = tk.Button(root, text="Browse", command=self.browse_source, bg="#8AC6D1", fg="white", width=10, height=1, relief='raised')
        source_browse_button.place(x=330,y=30)
        
        # Target file label
        target_label = tk.Label(root, text="Cohort Schedule:")
        target_label.place(x=30,y=70)

        self.target_file_label = tk.Label(root, text='No File Selected')
        self.target_file_label.place(x=150,y=70)

        # Target file browse button
        target_browse_button = tk.Button(root, text="Browse", command=self.browse_target, bg="#8AC6D1", fg="white", width=10, height=1, relief='raised')
        target_browse_button.place(x=330,y=70)

        # Submit button
        submit_button = tk.Button(root, text="Submit", command=self.submit, bg="#8AC6D1", fg="white", width=20, height=2)
        submit_button.place(x=150,y=115)
        
        #Backup Checkbox
        self.create_backup_var = tk.BooleanVar()
        create_backup_checkbox = tk.Checkbutton(root, text="Create Backup", variable=self.create_backup_var)
        create_backup_checkbox.pack()
        create_backup_checkbox.place(x=30,y=120)

        #Status notification
        self.status_label = tk.Label(root, text="")
        self.status_label.place(x=330, y=125)

    def browse_source(self):
        source_file = filedialog.askopenfilename(title = "Select a File", filetypes = (("Excel files", "*.xlsx"),("Excel files", "*.xls")))
        self.source_file_path = os.path.realpath(source_file)
        self.source_file_label.config(text=os.path.basename(source_file))
        self.source_file_name = os.path.basename(source_file)
        self.status_label.config(text="")

    def browse_target(self):
        target_file = filedialog.askopenfilename(title = "Select a File", filetypes = (("Excel files", "*.xlsx"),("Excel files", "*.xls")))
        self.target_file_path = os.path.realpath(target_file)
        self.target_file_label.config(text=os.path.basename(target_file))
        self.target_file_name = os.path.basename(target_file)
        self.status_label.config(text="")
#/--------------------------------Submit Algorithm------------------------------/
    #Submit button function. Will attempt to load CST.xlsx file first, then impor the data
    #from the calendar into the CST.xlsx file.
    def submit(self):
        try:
            # here you have access to the source and target files
            wb1 = load_workbook(self.target_file_path)
            wb2 = load_workbook(self.source_file_path)
            
            #greened out bottom lines 4/18 11.58am
            #ungreened out, it was working with this greened out. this code is needed to find errors.
            wb1.save(self.target_file_path)
            self.status_label.config(text="Loading..")
        except UnboundLocalError as e:
            print(e)
            self.status_label.config(text="Update files to .xlsx")
        except FileNotFoundError as e:
            print(e)
            self.status_label.config(text="File not found")
        except Exception as e: 
            print(e)
            self.status_label.config(text="Error: Close Excel Files")

        # here you have access to the source and target files
        source_file_name = self.source_file_label['text']
        target_file_name = self.target_file_label['text']

        source_file_path = self.source_file_path
        target_file_path = self.target_file_path

        # Prepare Cohort Sched excel workbook and sheet to be added
        wb1 = load_workbook(self.target_file_path)
        wb2 = load_workbook(self.source_file_path)

         # Create backup if the user has selected that option
        if self.create_backup_var.get():
            backup_folder = tk.filedialog.askdirectory()
            if backup_folder:  # Make sure the user selected a directory
                backup_path = os.path.join(backup_folder, os.path.basename(target_file_path))
                shutil.copy2(target_file_path, backup_path)


        # Obtain primary worksheet from source
        wsC = wb2.worksheets[0]

        # Split name of wb2 to isolate name of program
        targetWS = self.source_file_name.split()
        ws = wb1[targetWS[0]]

        #Find number of rows to be copied and inserted. compare val of a and b cells, 
        #increase b by one (one row down) and repeat until section is over. a starts at 3 to make up for header
        #b starts at i (4) right under 3, I is incremented each comparison.
        #The section name is used to find the number of courses.
        #!!!!!!!!!!!!!!!!!!!!!UPDATE:5/8 THE ABOVE DESCRIPTION IS NO LONGER TRUE, HAS CHANGED BC OF LATEST AI ADDITION
        i = 4
        a = ws.cell(row=3, column=1)
        b = ws.cell(row=i, column=1)
        counter = 2

        while (a.value == b.value):
            counter += 1
            i += 1
            b = ws.cell(row=i, column=1)  
            
       
        # Read existing rows from the target worksheet and store them in memory
        existing_rows = []
        merged_cells_info = []
        hidden_rows = []
        for row in ws.iter_rows(min_row=2):
            row_data = {}
            for cell in row:
                if isinstance(cell, MergedCell):  # Skip merged cells
                    continue
                row_data[cell.column_letter] = cell.value
                if ws.row_dimensions[cell.row].hidden:
                    row_data["hidden"] = True
                    hidden_rows.append(cell.row)   # Keep track of hidden rows
            existing_rows.append(row_data)
            # Store merged cells' information (cell range and data)
            for merged_cell_range in ws.merged_cells.ranges:
                top_left_cell = ws[merged_cell_range.start_cell.coordinate]
                merged_cells_info.append((merged_cell_range, top_left_cell.value))

        # Add new rows at the top of the target worksheet
        ws.insert_rows(2, counter)

        for row_index, row_data in enumerate(existing_rows, start=counter + 2):
            for col, value in row_data.items():
                if col != "hidden":
                    cell = ws.cell(row=row_index, column=column_index_from_string(col)) 
                    cell.value = value
            if "hidden" in row_data and row_data["hidden"]:
                ws.row_dimensions[row_index].hidden = True  



        # Reapply merged cell ranges and restore their data
        print("Before loop, merged_cells_info:", len(merged_cells_info))

        merged_cells_info_copy = merged_cells_info.copy()
        merged_cells_info_copy = list(set(merged_cells_info_copy))
        while merged_cells_info_copy:
            merged_cell_range, cell_value = merged_cells_info_copy[0]
            print("Inside loop, merged_cell_range:", merged_cell_range)
            ws.merge_cells(start_row=merged_cell_range.bounds[1] + counter, start_column=merged_cell_range.bounds[0],
                           end_row=merged_cell_range.bounds[3] + counter, end_column=merged_cell_range.bounds[2])
            top_left_cell = ws.cell(row=merged_cell_range.bounds[1] + counter, column=merged_cell_range.bounds[0])
            top_left_cell.value = cell_value
            merged_cells_info_copy.pop(0)
            # Remove the processed merged cell range from the list
            for i, (merged_cell_range_copy, _) in enumerate(merged_cells_info_copy):
                if merged_cell_range_copy == merged_cell_range:
                    del merged_cells_info_copy[i]
                    break
            print("merged_cells_info_copy:", merged_cells_info_copy)

        print("After loop, merged_cells_info:", len(merged_cells_info_copy))

        # After reapplying merged cells, adjust hidden rows
        new_hidden_rows = [row + counter for row in hidden_rows]

        # Unhide original rows and hide new ones
        for row in hidden_rows:
            ws.row_dimensions[row].hidden = False

        for row in new_hidden_rows:
            ws.row_dimensions[row].hidden = True


        # Title for new calendar added is program name + section 
        ws.cell(row=2, column=1).value = targetWS[0] + ' ' + targetWS[1]  

        #obtain section from copied ws and paste under 'program'
        for row in ws.iter_rows(min_row=3, max_col=1, max_row=counter+1):
            for cell in row:
                cell.value = targetWS[1]

        #obtain day and paste under Day section
        for row in ws.iter_rows(min_row=3, min_col=4, max_col=4, max_row=counter+1):
            for cell in row:
                cell.value = targetWS[2]
                if ".xlsx" in cell.value:
                    cell.value = cell.value.replace(".xlsx", "") 

        #Find cells with a 'credit' value, meaning rows with course listing on them. Obtain global 'term' for current term as wel
        def iterate():
            for row in wsC.iter_rows():
                for cell in row:
                    if cell in cellChecker:  # make sure not to repeat the same cell
                        continue

                    if cell.value is not None and not (isinstance(cell.value, str) and cell.value.strip() == ""):
                        a = str(cell.value).split()
                        if a[0] == 'Term':
                            global term
                            cellChecker.append(cell)
                            term = a[2]
                        # Find a course row based on a number existing on column B (the credits cell)
                        if cell.column_letter == 'B' and str(cell.value).isnumeric():
                            cellChecker.append(cell)
                            print(f"Returning cell {cell.coordinate}: {cell.value}")  # Add this line for debugging
                            return cell
            return None


    
        rowCounter=0
        row = 3
        col = 2
        currentRow = 0
        currentCol = 0
        global cellChecker
        cellChecker = []

        #row, col determine coordinates in CST doc. currentRow currentCol determine cooridnates on schedule to be added.
        #Counter is subtracted 1 to account for op header rows 1 and 2 
        while(rowCounter != counter-1):
            #Find cells with a 'credit' value, meaning rows with course listing on them. Obtain global 'term' for current term as well.
            currentCell = iterate()
            
            if currentCell is None:
            # Handle the case when no matching cell is found
                print("No more matching cells found.")
                break

            #Write down course name and number
            ws.cell(row, col).value = str(currentCell.offset(0, -1).value).split()[0] + " " + str(currentCell.offset(0, -1).value).split()[1].replace(':','')

            #Write down term
            ws.cell(row, col+1).value = term

            #write down weeks
            ws.cell(row, col+3).value = str(currentCell.offset(0, 2).value)

            #Write down dates
            ws.cell(row, col+5).value = str(currentCell.offset(0, 1).value).split()[0]
            ws.cell(row, col+6).value = str(currentCell.offset(0, 1).value).split()[2]         

            #Move to next row for next iteration
            row += 1
            rowCounter += 1
            
        #show complete flag and save.
        try:
            wb1.save(self.target_file_path)
            self.status_label.config(text="Complete!")
        except: 
            self.status_label.config(text="Error: Close Excel Files")

root = tk.Tk()
app = App(root)
root.mainloop()


