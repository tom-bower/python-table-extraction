import json
import docx

# Open Word document
filename = "doc.docx"
doc = docx.Document(filename)

# Extract all tables in document
tables = doc.tables

print(len(tables))

for table in tables:
    data = []
    keys = None

    for i, row in enumerate(table.rows):

        # Grab all data in the table row
        row_text = (cell.text for cell in row.cells)

        # Assume table headers are in first row, set these as keys
        if i == 0:
            keys = tuple(row_text)

            # Go to next row
            continue

        # Append row data to dict
        row_data = dict(zip(keys, row_text))
        data.append(row_data)

    print(json.dumps(data))
