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

	# Create a dictionary to hold the checkbox states (True or False)
	checkbox_vars = {}

	# This holds the list of selected options
	selected_options = []

	# Function to get selected options when the user clicks "Submit"
	def submit():
		# Collect selected options based on checkbox states
		selected_options[:] = [option for option, var in checkbox_vars.items() if var.get()]
		window.quit()  # Close the window after selection

	# Create checkboxes for each option
	for option in options:
		var = tk.BooleanVar()
		checkbox = tk.Checkbutton(window, text=option, variable=var)
		checkbox.pack(anchor='w')
		checkbox_vars[option] = var

	# Add a Submit button
	submit_button = tk.Button(window, text="Submit", command=submit)
	submit_button.pack()

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

def add_columns(df):
	new_column_values = ['Value1'] * 7

	# Check if the column already exists

	if 'TS' in df.columns:
		raise ValueError("The column 'TS' already exists in the DataFrame.")
    
	# Try inserting the column at the specified position
	try:
		df.insert(0, 'TS', new_column_values)
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

	selected_values = select_values()
	print(selected_values)

	"""
	# Select excel file to process.
	file_path = load_excel_file()
	if not file_path:
		print("No file selected. Exiting...")
		return
	
	df = pd.read_excel(file_path)

	# Remove the second row.
	df = remove_row(df)

	# Add column.
	df = add_columns(df)

	# Perform column sort.
	df = rearrange_column(df)

	# Save output file.
	if df is not False:
		save_excel_file(df)

	"""

if __name__ == "__main__":
	main()
