# Extract placeholders from the first column of each table
def extract_placeholders(doc):
    placeholders = []
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 2:  # Ensure the row has at least two columns
                first_column_text = row.cells[0].text.strip()
                if first_column_text:  # Only add non-empty cells
                    placeholders.append(first_column_text)
    return placeholders


# Fill the data into the document
def fill_data(doc, form_data):
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 2:
                key = row.cells[0].text.strip()
                if key in form_data:
                    row.cells[1].text = form_data[key]  # Fill the second column