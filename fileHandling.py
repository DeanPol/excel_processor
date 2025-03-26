import tkinter as tk
from tkinter import filedialog
import json
import os

# Load and Save our excel files
def load_excel_file():
	root = tk.Tk()
	root.withdraw()
	file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
	return file_path

def save_excel_file(df):
	root = tk.Tk()
	root.withdraw()
	file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
	df.to_excel(file_path, index=False)
	print(f"File saved at: {file_path}")
     
def select_json_file():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if not file_path:
        print("No JSON file selected.")
        return None
    return os.path.basename(file_path)
	
# Load and select options from json file
def load_options_from_json(file_path):
    # Load the JSON file
    with open(file_path, 'r') as file:
        data = json.load(file)
    return data['options']  # Return the list of options from the JSON

def select_values():
    file_name = select_json_file()
    # Load options from JSON file
    options = load_options_from_json(file_name)

    # Create the main window
    window = tk.Tk()
    window.title("Select Values")

    # Create a canvas and a vertical scrollbar
    canvas = tk.Canvas(window)
    scrollbar = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Create a frame inside the canvas to hold the checkboxes
    checkbox_frame = tk.Frame(canvas)

    # Add the checkboxes to the frame
    checkbox_vars = {}
    for option in options:
        var = tk.BooleanVar()
        checkbox = tk.Checkbutton(checkbox_frame, text=option, variable=var)
        checkbox.pack(anchor='w', padx=10)
        checkbox_vars[option] = var

    # Create a window on the canvas to hold the frame (this enables scrolling)
    canvas.create_window((0, 0), window=checkbox_frame, anchor="nw")

    # Configure the scrollbar to work with the canvas
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    # Update the scrollable region whenever new widgets are added
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    checkbox_frame.bind("<Configure>", on_frame_configure)

    # Create a function to get selected options when the user clicks "Submit"
    selected_options = []

    def submit():
        selected_options[:] = [option for option, var in checkbox_vars.items() if var.get()]
        window.quit()  # Close the window after selection

    # Add a Submit button
    submit_button = tk.Button(window, text="Submit", command=submit)
    submit_button.pack(pady=10)

    # Add a Select All button
    def select_all():
        for var in checkbox_vars.values():
            var.set(True)  # Set all checkboxes to checked

    select_all_button = tk.Button(window, text="Select All", command=select_all)
    select_all_button.pack(pady=5)

    # Run the window's event loop
    window.mainloop()

    # Return the list of selected options
    return selected_options

def remove_trailing_empty_rows(df):
    return df.iloc[:df.dropna(how="all").index[-1] + 1] if not df.dropna(how="all").empty else df.iloc[:0]

