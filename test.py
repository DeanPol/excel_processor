import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
import json

def load_options_from_json(file_path):
    # Load the JSON file
    with open(file_path, 'r') as file:
        data = json.load(file)
    return data['options']  # Return the list of options from the JSON


def load_excel_file():
	root = tk.Tk()
	root.withdraw()
	file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
	return file_path

def select_values():
    # Load options from JSON file
    options = load_options_from_json('scenarios.json')

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


def save_excel_file(df):
	root = tk.Tk()
	root.withdraw()
	file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
	df.to_excel(file_path, index=False)
	print(f"File saved at: {file_path}")

def remove_row(df):
	# Check if document has at least 2 rows
	if df.shape[0] < 2:
		return False
	
	df.drop(index=1)
	df.reset_index(drop=True, inplace=True)
	return df

def add_columns(df, selected_values):

	# Check if the column already exists

	if 'TS' in df.columns:
		raise ValueError("The column 'TS' already exists in the DataFrame.")
    
	# Try inserting the column at the specified position
	try:
		df.insert(0, 'TS', ['TS'] + selected_values)
	except Exception as e:
		raise ValueError(f"Error inserting the column: {e}")
	return df

def rearrange_column(df):
	if 'requestName' in df.columns:
		df['requestName'] = df['requestName'].sort_values().reset_index(drop=True)
		return df
	else:
		print("Invalid data - Column 'requestName' does not exist")
		return False

def main():

	#selected_values = select_values()


	# Select excel file to process.
	file_path = load_excel_file()
	if not file_path:
		print("No file selected. Exiting...")
		return
	
	df = pd.read_excel(file_path)

	# Remove the second row.
	#df = remove_row(df)
	df.drop(index=1)
	df.reset_index(drop=True, inplace=True)

	# Add column.
	#df = add_columns(df, selected_values)

	"""

	# Perform column sort.
	df = rearrange_column(df)

	"""
	# Save output file.
	if df is not False:
		save_excel_file(df)


if __name__ == "__main__":
	main()
