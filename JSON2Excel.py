import pandas as pd
import json
import sys

file = sys.argv[1]

# Load your json data
with open(file) as json_file:
    json_data = json.load(json_file)


json_data = json_data['data']

rows = []

# Iterate through each JSON object
for item in json_data:
    #Do all the things to the JSON
    fields = pd.json_normalize(item['fields'])
    item.pop('fields', None)
    item_df = pd.json_normalize(item)
    row = pd.concat([item_df, fields], axis=1)
    rows.append(row)

#Concatenate all rows
df = pd.concat(rows, ignore_index=True)

#Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
