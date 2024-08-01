import re
import smartsheet

# Define the path to the Results.txt file
results_file_path = "N:\\Production\\Production reports\\Results.txt"

# Define the Smartsheet API access token
access_token = "FHrmEeDnC1oEBmBaWTCp1SNvtdEFbNdT2pEgc"

# Initialize the Smartsheet client
smartsheet_client = smartsheet.Smartsheet(access_token)

# Define a mapping of sheet names to their corresponding Smartsheet IDs
sheet_mapping = {
    "AngleMaster": "6973054634643332",
    "Beamline": "8633118017671044",
    "Daily": "6630177199050628",
    "Paint": "6705691314048900",
    "PlateTable": "2861850202951556"
}

# Function to read and parse the Results.txt file
def read_results(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()
    results = []
    for line in lines:
        # Extract the sheet name, date, and total weight using regex
        match = re.search(r'(\w+) throughput\\(\d{4}-\d{2}-\d{2}) (\d+\.\d+) Tons.txt', line)
        if match:
            sheet_name = match.group(1)
            date = match.group(2)
            total_weight = float(match.group(3))
            results.append((sheet_name, date, total_weight))
    return results

# Function to get column ID by column title
def get_column_id(sheet, column_title):
    for column in sheet.columns:
        if column.title == column_title:
            return column.id
    raise ValueError(f"Column '{column_title}' not found in the sheet.")

# Function to update Smartsheet with the results
def update_smartsheets(results, sheet_mapping):
    for sheet_name, date, total_weight in results:
        # Get the corresponding sheet ID
        sheet_id = sheet_mapping.get(sheet_name)
        if not sheet_id:
            print(f"Sheet ID for '{sheet_name}' not found. Skipping this entry.")
            continue

        # Get the sheet
        sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

        # Get column IDs for 'Date' and 'Total Weight'
        date_column_id = get_column_id(sheet, 'Date')
        total_weight_column_id = get_column_id(sheet, 'Total Weight')

        rows_to_update = []

        # Find the row with the matching date
        for row in sheet.rows:
            for cell in row.cells:
                if cell.column_id == date_column_id and cell.value == date:
                    # Create an update for the 'Total Weight' cell
                    new_cell = smartsheet.models.Cell({
                        'column_id': total_weight_column_id,
                        'value': total_weight
                    })
                    updated_row = smartsheet.models.Row({
                        'id': row.id,
                        'cells': [new_cell]
                    })
                    rows_to_update.append(updated_row)
                    break

        if rows_to_update:
            # Update the rows in the sheet
            smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)

    print("Smartsheets have been updated successfully.")

# Read the results from the file
results = read_results(results_file_path)

# Update the Smartsheets with the results
update_smartsheets(results, sheet_mapping)
