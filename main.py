import sys
import re
import pandas as pd

from fileHandling import load_excel_file
from fileHandling import save_excel_file
from fileHandling import select_values

def remove_row(df):
	# Check if document has at least 2 rows
	if df.shape[0] < 2:
		raise ValueError("DataFrame must have at least two rows.")
	try: 
		df.drop(0, inplace=True)
		df.reset_index(drop=True, inplace=True)
		return df
	except Exception as e:
		print(f"Error: {e}")
		sys.exit(1)

def rearrange_column(df):
	if 'requestName' in df.columns:
		# Apply extraction function and sort
		df["sort_values"] = df["requestName"].apply(extract_request_numbers)
		
		# Sorting by extracted (X, Y, Z) where the sorting is done numerically
		df.sort_values(by="sort_values", ascending=True, inplace=True)
		
		# Clean up after sorting (drop the helper column)
		df.drop(columns=["sort_values"], inplace=True)

		df = df.reset_index(drop=True)

		return df

	else:
		print("Invalid data - Column 'requestName' does not exist")
		return False
	
	
def add_column(df, selected_values):

	# Check if the column already exists
	if 'TS' in df.columns:
		raise ValueError("The column 'TS' already exists in the DataFrame.")

	# Change df in order to accomodate new column 
	max_rows = max(df.shape[0], len(selected_values))
	try:
		df = df.reindex(range(max_rows))
	except Exception as e:
		raise ValueError(f"Error changing data frame row number: {e}")

	# Try inserting the column at the specified position
	try:
		df.insert(0, 'TS', selected_values)
	except Exception as e:
		raise ValueError(f"Error inserting the column: {e}")
	
	return df


def shift_column_up_from_index(df: pd.DataFrame, column_name: str, start_index: int) -> pd.DataFrame:
    df_copy = df.copy()
    df_copy.loc[start_index:, column_name] = df_copy.loc[start_index:, column_name].shift(-1)
    return df_copy

def extract_scenario_numbers(val):
    if pd.isna(val) or not isinstance(val, str):
        return (0, 0, 0)

    match = re.search(r"S(\d{1,2})\s+TS\s*(\d{1,2})(?:\.(\d))?", val)
    if match:
        X = int(match.group(1))
        Y = int(match.group(2))
        Z = int(match.group(3)) if match.group(3) else 0
        return (X, Y, Z)

    return (0, 0, 0)

def extract_request_numbers(val):
    if pd.isna(val) or not isinstance(val, str):
        return (0, 0, 0)

    match = re.search(r"S(\d{1,2})\s+TS(-?)\s*(\d{1,2})(?:\.(\d))?", val)
    if match:
        X = int(match.group(1))
        Y = int(match.group(3))
        Z = int(match.group(4)) if match.group(4) else 0
        return (X, Y, Z)

    return (0, 0, 0)

# Function to align columns
def align_columns(df):
	blank_row = pd.DataFrame([{col: None for col in df.columns}])  # Create a blank row
	need_loop = True
	
	while need_loop:
		need_loop = False  # Assume no changes will be needed

		for idx, row in df.iterrows():
			scenario_tuple = extract_scenario_numbers(row.TS)        
			request_tuple = extract_request_numbers(row.requestName)

			if scenario_tuple == (0, 0, 0):
				return df  # Exit function when termination condition is met
			
			if request_tuple == (0, 0, 0):
				continue

			if scenario_tuple != request_tuple:
				df = pd.concat([df.iloc[:idx], blank_row, df.iloc[idx:]]).reset_index(drop=True)
				
				# Shift the 'TS' column up from the inserted row
				df.loc[idx:, 'TS'] = df.loc[idx:, 'TS'].shift(-1)
				
				need_loop = True  # A modification was made, so restart the loop
				break  # Exit the for-loop and restart while-loop
		
	return df

def remove_trailing_empty_rows(df):
    return df.iloc[:df.dropna(how="all").index[-1] + 1] if not df.dropna(how="all").empty else df.iloc[:0]


def main():

	selected_values = select_values()

	# Select excel file to process.
	file_path = load_excel_file()
	if not file_path:
		print("No file selected. Exiting...")
		return
	
	df = pd.read_excel(file_path)

	# Remove the second row.
	df = remove_row(df)

	df = rearrange_column(df)


	# Add column.
	df = add_column(df, selected_values)

	df = align_columns(df)

	df = remove_trailing_empty_rows(df)

	# Save output file.
	save_excel_file(df)
	

if __name__ == "__main__":
	main()
