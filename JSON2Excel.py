import pandas as pd
import json
import sys

def parse_json_string(data):
    try:
        return json.loads(data)
    except json.JSONDecodeError:
        return data

if len(sys.argv) < 2:
    print("Usage: python JSON2Excel2.py <input_json_file>")
    sys.exit(1)

file = sys.argv[1]

try:
    with open(file) as json_file:
        json_data = json.load(json_file)
except FileNotFoundError:
    print(f"File {file} not found.")
    sys.exit(1)
except json.JSONDecodeError:
    print(f"Error decoding JSON from file {file}.")
    sys.exit(1)

rows = []

# Check if the top-level element is a list
# This is how we handle nested JSON structures
# Ew
if isinstance(json_data, list):
    json_items = json_data
elif isinstance(json_data, dict):
    json_items = json_data.get('inspections', [])
else:
    print("Unsupported JSON format.")
    sys.exit(1)

# Iterate through each JSON object
for item in json_items:
    try:
        # Parse the 'data' field if it is a JSON string
        item_data = parse_json_string(item.get('data', '{}'))
        # Normalize the entire item to flatten the JSON structure
        item_df = pd.json_normalize(item_data)
        rows.append(item_df)
    except Exception as e:
        print(f"Error normalizing item: {item}")
        print(e)
        sys.exit(1)

if not rows:
    print("No valid data to concatenate.")
    sys.exit(1)

# Concatenate all rows
df = pd.concat(rows, ignore_index=True)

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
