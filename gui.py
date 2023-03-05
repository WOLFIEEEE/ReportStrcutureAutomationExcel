import tkinter as tk
from tkinter import filedialog
import pandas as pd
from main import main , delete_whole_column_letter
import openpyxl
workbook2 = openpyxl.load_workbook('Boilerplate.xlsx')
def open_file():
    # create the BooleanVar objects inside the function
    checkbox_values = {
        "Windows": tk.BooleanVar(),
        "macOS": tk.BooleanVar(),
        "Android": tk.BooleanVar(),
        "iOS": tk.BooleanVar()
    }

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        # remove the Open File button
        open_button.pack_forget()
        
        # create a label to display the file name
        file_name = file_path.split("/")[-1]
        label = tk.Label(root, text=f"Selected file: {file_name}")
        label.pack(pady=20)

        # create a checkbox for each platform
        for platform, var in checkbox_values.items():
            checkbox = tk.Checkbutton(root, text=platform, variable=var)
            checkbox.pack()

        def process_checkboxes():
            selected_platforms = [platform for platform, var in checkbox_values.items() if var.get()]
            print(f"Selected platforms: {selected_platforms}")

            # create a list of unchecked platforms
            unchecked_platforms = [platform for platform in checkbox_values.keys() if platform not in selected_platforms]

            # disable the process button and show loading message
            process_button.config(state="disabled", text="Processing...")
            root.update()

            # delete unchecked columns and send report
            column_arr = []
            for i in unchecked_platforms:
                if(i == "Windows"):
                    column_arr.append('B')
                if(i == "macOS"):
                    column_arr.append('C')
                if(i == "Android"):
                    column_arr.append('D')
                if(i == "iOS"):
                    column_arr.append('E')
        
            # delete_whole_column_letter(column_arr)
            main(file_path)

            # enable the process button and show completed message
            process_button.config(state="normal", text="Completed")
            root.update()

            # ask user where to save the new file
            save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

            # check if user selected a file to save
            if save_file_path:
                # read the processed data
                workbook2.save(save_file_path)
                print(f"File saved at {save_file_path}")

        process_button = tk.Button(root, text="Process Checkboxes", command=process_checkboxes)
        process_button.pack(pady=20, padx=10)

# create the main window
root = tk.Tk()
root.geometry("400x400")

# create a button to open the file dialog
open_button = tk.Button(root, text="Open File", command=open_file)
open_button.pack(pady=20, padx=10)

# run the main loop
root.mainloop()
