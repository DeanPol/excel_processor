import sys
import re
import pandas as pd

from fileHandling import load_excel_file
from fileHandling import save_excel_file
from fileHandling import select_values
from fileHandling import remove_trailing_empty_rows

def remove_row(df):
	if df.shape[0] < 2:
		raise ValueError("DataFrame must have at least two rows.")
	try: 
		df = df[df['requestName'] != 'Rescheduling batch job']
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
	# Change df in order to accomodate new column by adding new rows
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
	blank_row = pd.DataFrame([{col: "" for col in df.columns}])  # Create a blank row
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

def convertToMs(val):
	if type(val) != str:
		return val
	if 'ms' in val:
		val = float(val.replace('ms',''))
	elif 's' in val:
		val = float(val.replace('s', ''))
		val = str(val * 1000)
	return val		

def outputType(val):
	print(f"Type of value is: {type(val)}")

def append_Ms_column(df, baseColumnName, newColumnName):
    try:
        df[newColumnName] = df.loc[:, baseColumnName].apply(convertToMs) 
    except KeyError:
        raise ValueError(f"Column '{baseColumnName}' not found in DataFrame")
    except Exception as e:
        raise ValueError(f"Error applying value change: {e}")
    return df

def main():
	# Select json file to process.
	selected_values = select_values()
	if len(selected_values) == 0:
		print("No scenarios selected. Exiting...")
		return None

	# Select excel file to process.
	file_path = load_excel_file()
	if not file_path:
		print("No file selected. Exiting...")
		return
	
	df = pd.read_excel(file_path)

	df.columns = df.columns.astype(str)

	# Row 'Rescheduling batch job' is not needed
	# Remove it if it exists.
	df = remove_row(df)

	# Sort our requestName column.
	df = rearrange_column(df)

	# Add scenarios column.
	df = add_column(df, selected_values)

	# Sort requests according to scenarios
	df = align_columns(df)

	df = remove_trailing_empty_rows(df)

	# Append AvgToMs, MedianToMs and 90% Ms columns
	df = append_Ms_column(df, 'Avg', 'AvgMs')
	df = append_Ms_column(df, 'Median', 'MedianMs')
	df = append_Ms_column(df, '0.9', '90% Ms')

	# Save output file.
	save_excel_file(df)
	

if __name__ == "__main__":
	main()
