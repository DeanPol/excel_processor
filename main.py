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

# Function to extract (X, Y, Z) with optional '-' and ':', and X can have 1 or 2 digits
def extract_values(request_name):
    # Match the pattern with optional '-' between 'TS' and Y and optional ':' at the end,
    # and X can be 1 or 2 digits
    match = re.match(r"S(\d{1,2})\s+TS[-\s]?(\d+)(?:\.(\d))?(?:\s*(:)?)", request_name)
    if match:
        X = int(match.group(1))  # Extract X (can be 1 or 2 digits)
        Y = int(match.group(2))  # Extract Y
        Z = int(match.group(3)) if match.group(3) else -1  # If Z is missing, use -1
        return X, Y, Z
    return (0, 0, -1)  # Default in case of a mismatch

def rearrange_column(df):
	if 'requestName' in df.columns:
		# Apply extraction function and sort
		df["sort_values"] = df["requestName"].apply(extract_values)
		
		# Sorting by extracted (X, Y, Z) where the sorting is done numerically
		df.sort_values(by="sort_values", ascending=True, inplace=True)
		
		# Clean up after sorting (drop the helper column)
		df.drop(columns=["sort_values"], inplace=True)

		return df

	else:
		print("Invalid data - Column 'requestName' does not exist")
		return False
	
	
def add_columns(df, selected_values):

	# Check if the column already exists
	if 'TS' in df.columns:
		raise ValueError("The column 'TS' already exists in the DataFrame.")
	
	# Grab number of rows
	number_rows = df.shape[0]
    
	# Try inserting the column at the specified position
	try:
		df.insert(0, 'TS', selected_values + [''] * (number_rows - len(selected_values)))
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

def align_columns(df):
    """
    Aligns the second column based on missing values from the first column.
    Instead of modifying while iterating, we collect missing indices and add them at once.
    """
    missing_indices = []
    
    for idx, row in df.iterrows():
        scenario_tuple = extract_scenario_numbers(row.TS)
        request_tuple = extract_request_numbers(row.requestName)

        # If there's a mismatch or missing value, we mark this index
        if request_tuple != scenario_tuple:
            missing_indices.append(idx)
    
    # Insert blank rows in one batch (to avoid modifying during iteration)
    for offset, idx in enumerate(missing_indices):
        blank_row = pd.DataFrame([[df.loc[idx, "TS"], None]], columns=df.columns)
        df = pd.concat([df.iloc[:idx + offset], blank_row, df.iloc[idx + offset:]]).reset_index(drop=True)

    return df

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
	df = add_columns(df, selected_values)

	df = align_columns(df)

	# Save output file.
	save_excel_file(df)
	

if __name__ == "__main__":
	main()
