import re
import os
import smartsheet
import tkinter as tk
import threading
import logging
from datetime import datetime
from queue import Queue
from tkinter.ttk import Progressbar, Style

# Configuration dictionary to store the necessary configuration values
CONFIG = {
    'access_token': os.getenv('SMARTSHEET_ACCESS_TOKEN', "FHrmEeDnC1oEBmBaWTCp1SNvtdEFbNdT2pEgc"),  # Access token for Smartsheet API
    'sheet_mapping': {  # Mapping of sheet names to their respective Smartsheet IDs
        "AngleMaster": "2865724263452548",
        "Beamline": "8633118017671044",
        "Daily": "6630177199050628",
        "Paint": "6705691314048900",
        "PlateTable": "2861850202951556",
        "Shipping": "589005049515908"
    },
    'base_path': r"T:\\Production reports\\",  # Base path for production reports
    'results_file_path': r"T:\\Production reports\\Results.txt"  # Path to store results
}

# Initialize logging with error level and format
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize the Smartsheet client with the access token from the configuration
smartsheet_client = smartsheet.Smartsheet(CONFIG['access_token'])

def start_scripts(log_queue, progress_queue):
    # Start the main processing and updating script
    log_queue.put("Starting PDCReport process...")
    find_and_process_files(CONFIG['base_path'], CONFIG['results_file_path'], log_queue, progress_queue)  # Process files
    log_queue.put("PDCReport process completed. Starting Smartsheet update...")
    results = read_results(CONFIG['results_file_path'])  # Read results from the results file
    update_smartsheets(results, CONFIG['sheet_mapping'], log_queue, progress_queue)  # Update Smartsheet with results
    log_queue.put("Smartsheet update completed. COMPLETE")
    progress_queue.put(100)  # Ensure progress bar reaches 100% at the end
    start_button.config(state=tk.NORMAL)  # Re-enable the start button

def start_button_handler():
    # Handle the start button click event
    start_button.config(state=tk.DISABLED)  # Disable the start button to prevent multiple clicks
    current_log.set("")  # Clear previous log messages
    progress.pack(pady=10)  # Show progress bar
    threading.Thread(target=start_scripts, args=(log_queue, progress_queue)).start()  # Start the script in a new thread
    app.after(100, process_log_queue)  # Schedule log processing
    app.after(100, update_progress)  # Schedule progress bar update

def process_log_queue():
    # Process the log queue and update the GUI
    if not log_queue.empty():
        current_log.set(log_queue.get())  # Get the next log message and set it to the log label
    app.after(100, process_log_queue)  # Schedule the next log queue processing

def update_progress():
    # Update the progress bar
    if not progress_queue.empty():
        progress_value = progress_queue.get()  # Get the next progress value
        progress['value'] = progress_value  # Update the progress bar
        progress_label.set(f'{progress_value}%')  # Update the progress label
        if progress_value == 100:
            style.configure("green.Horizontal.TProgressbar", background="#006400")  # Change progress bar color to green
    app.after(100, update_progress)  # Schedule the next progress bar update

def generate_unique_filename(directory, filename):
    # Generate a unique filename in the specified directory
    base, extension = os.path.splitext(filename)  # Split the filename into base and extension
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):  # Check if the file already exists
        unique_filename = f"{base} ({counter}){extension}"  # Append a counter to the filename
        counter += 1
    return unique_filename  # Return the unique filename

def extract_date_from_content(content, is_shipping=False):
    # Extract the date from the file content
    pattern = r'Date Shipped:\s+(\d{1,2}/\d{1,2}/\d{4})' if is_shipping else r'Date Station Completed:\s+(\d{1,2}/\d{1,2}/\d{4})'
    match = re.search(pattern, content)  # Search for the date pattern in the content
    if match:
        return datetime.strptime(match.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')  # Convert the date to YYYY-MM-DD format
    else:
        raise ValueError("Date not found in file content")  # Raise an error if the date is not found

def process_file(file_path, results_list, progress_queue, total_files, current_file_index):
    # Process a general file and update the results list
    try:
        with open(file_path, 'r', encoding='utf-8') as file:  # Open the file in read mode
            content = file.read()  # Read the content of the file
            lines = content.splitlines()  # Split the content into lines

        date_str = extract_date_from_content(content, is_shipping=False)  # Extract the date from the content
        completed_numbers = []
        regex = re.compile(r'(\d{1,3}(?:,\d{3})*)#')  # Regular expression to find completed numbers

        for i, line in enumerate(lines):
            if "Completed" in line:
                for next_line in lines[i + 1:i + 5]:  # Check the next 5 lines for completed numbers
                    match = regex.search(next_line)
                    if match:
                        completed_numbers.append(int(match.group(1).replace(',', '')))  # Add the completed number to the list

        total_sum = sum(completed_numbers) / 2000  # Calculate the total sum in tons
        tonnage = round(total_sum, 2)  # Round the total sum to 2 decimal places
        new_file_name = generate_unique_filename(os.path.dirname(file_path), f"{date_str} {tonnage} Tons.txt")  # Generate a unique filename
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)  # Get the new file path
        os.rename(file_path, new_file_path)  # Rename the file
        logging.info(f"File renamed to: {new_file_name}")  # Log the new file name
        results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))  # Add the new file path to the results list
    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")  # Log any errors
    finally:
        progress_value = int(((current_file_index + 1) / total_files) * 100)  # Calculate the progress value
        progress_queue.put(progress_value)  # Update the progress queue

def process_shipping_file(file_path, results_list, progress_queue, total_files, current_file_index):
    # Process a shipping file and update the results list
    try:
        with open(file_path, 'r', encoding='utf-8') as file:  # Open the file in read mode
            content = file.read()  # Read the content of the file
            lines = content.splitlines()  # Split the content into lines

        date_str = extract_date_from_content(content, is_shipping=True)  # Extract the date from the content
        total_weight = 0
        for line in lines:
            if "Total shipped to Jobsite:" in line:  # Find the total shipped weight
                hashtag_pos = line.rfind('#')
                if hashtag_pos != -1:
                    last_space_pos = line.rfind(' ', 0, hashtag_pos)
                    if last_space_pos != -1:
                        weight_str = line[last_space_pos + 1:hashtag_pos].replace(',', '')  # Extract the weight
                        total_weight = int(weight_str)
                        break

        total_weight_tons = total_weight / 2000  # Convert the weight to tons
        tonnage = round(total_weight_tons, 2)  # Round the weight to 2 decimal places
        new_file_name = generate_unique_filename(os.path.dirname(file_path), f"{date_str} {tonnage} Tons.txt")  # Generate a unique filename
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)  # Get the new file path
        os.rename(file_path, new_file_path)  # Rename the file
        logging.info(f"File renamed to: {new_file_name}")  # Log the new file name
        results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))  # Add the new file path to the results list
    except Exception as e:
        logging.error(f"Error processing shipping file {file_path}: {e}")  # Log any errors
    finally:
        progress_value = int(((current_file_index + 1) / total_files) * 100)  # Calculate the progress value
        progress_queue.put(progress_value)  # Update the progress queue

def find_and_process_files(base_path, log_file_path, log_queue, progress_queue):
    # Find and process files in the specified base path
    results_list = []

    # Clear the results file at the start
    open(log_file_path, 'w').close()  # Open the results file in write mode and close it to clear its contents

    # Set total files to the average number (12) for smoother progress bar
    total_files = 12
    current_file_index = 0

    for dirpath, _, filenames in os.walk(base_path):  # Walk through the base path directory
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)  # Get the full file path
            if filename == "Results.txt":
                continue  # Skip the results file
            if filename in ["Station Summary-currentday.txt", "Station Summary-prevday.txt"]:
                process_file(file_path, results_list, progress_queue, total_files, current_file_index)  # Process general files
            elif filename in ["ShippingList_by_Job-currentday.txt", "ShippingList_by_Job-prevday.txt"]:
                process_shipping_file(file_path, results_list, progress_queue, total_files, current_file_index)  # Process shipping files
            current_file_index += 1

    # Log results
    with open(log_file_path, 'a', encoding='utf-8') as log_file:  # Open the log file in append mode
        for result in results_list:
            log_file.write(f"{result}\n")  # Write each result to the log file

    # Archive the processed files
    for result in results_list:
        new_file_path = os.path.join(base_path, result)  # Get the new file path
        archive_dir = os.path.join(os.path.dirname(new_file_path), 'Archive')  # Get the archive directory path
        if 'Archive' not in os.path.dirname(new_file_path):
            new_filename = generate_unique_filename(archive_dir, os.path.basename(new_file_path))  # Generate a unique filename for the archive
            os.makedirs(archive_dir, exist_ok=True)  # Create the archive directory if it doesn't exist
            os.rename(new_file_path, os.path.join(archive_dir, new_filename))  # Move the file to the archive directory
            logging.info(f"Archived file to: {os.path.join(archive_dir, new_filename)}")  # Log the archiving
        else:
            logging.warning(f"Skipping file {new_file_path} as it is already in an archive directory.")  # Log if the file is already archived

def read_results(file_path):
    # Read and parse the results from the specified file
    try:
        with open(file_path, 'r', encoding='utf-8') as file:  # Open the file in read mode
            lines = file.read().splitlines()  # Read and split the lines
        results = []
        # Adjusted regex to match the given file paths
        for line in lines:
            match = re.search(r'([\w\s]+) throughput\\(?:Archive\\)?(\d{4}-\d{2}-\d{2}) (\d+\.\d+) Tons(?: \(\d+\))?.txt', line)
            if match:
                sheet_name = match.group(1).strip()  # Extract the sheet name
                date = match.group(2)  # Extract the date
                total_weight = float(match.group(3))  # Extract the total weight
                results.append((sheet_name, date, total_weight))  # Add the result to the list
            else:
                logging.warning(f"Line '{line}' did not match the expected format.")  # Log if the line doesn't match the format
        return results  # Return the results
    except Exception as e:
        logging.error(f"Error reading results file {file_path}: {e}")  # Log any errors
        return []

def get_column_id(sheet, column_title):
    # Get the column ID by column title
    for column in sheet.columns:
        if column.title == column_title:
            return column.id  # Return the column ID if the title matches
    raise ValueError(f"Column '{column_title}' not found in the sheet.")  # Raise an error if the column is not found

def update_smartsheets(results, sheet_mapping, log_queue, progress_queue):
    # Update the Smartsheet with the results
    total_entries = len(results)  # Get the total number of results
    current_entry_index = 0

    for sheet_name, date, total_weight in results:
        try:
            log_queue.put(f"Processing {sheet_name} for date {date} with weight {total_weight}")  # Log the processing start
            sheet_id = sheet_mapping.get(sheet_name)  # Get the sheet ID from the mapping
            if not sheet_id:
                log_queue.put(f"Sheet ID for '{sheet_name}' not found. Skipping this entry.")  # Log if the sheet ID is not found
                continue

            sheet = smartsheet_client.Sheets.get_sheet(sheet_id)  # Get the sheet from Smartsheet
            log_queue.put(f"Retrieved sheet: {sheet_name} (ID: {sheet_id})")  # Log the sheet retrieval

            try:
                date_column_id = get_column_id(sheet, 'Date')  # Get the column ID for the date column
                total_weight_column_id = get_column_id(sheet, 'Total Weight')  # Get the column ID for the total weight column
            except ValueError as e:
                log_queue.put(f"Error: {e}")  # Log any errors in getting the column IDs
                continue

            rows_to_update = []

            for row in sheet.rows:
                for cell in row.cells:
                    if cell.column_id == date_column_id and cell.value == date:  # Check if the date matches
                        log_queue.put(f"Found matching date: {date} in row ID: {row.id}")  # Log the matching row
                        new_cell = smartsheet.models.Cell({
                            'column_id': total_weight_column_id,
                            'value': total_weight  # Create a new cell with the total weight
                        })
                        updated_row = smartsheet.models.Row({
                            'id': row.id,
                            'cells': [new_cell]  # Create a new row with the updated cell
                        })
                        rows_to_update.append(updated_row)  # Add the updated row to the list
                        break

            if rows_to_update:
                smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)  # Update the rows in the sheet
                log_queue.put(f"Updated rows in sheet: {sheet_name}")  # Log the update
            else:
                log_queue.put(f"No rows found with the date {date} in sheet {sheet_name}")  # Log if no rows are found

        except Exception as e:
            log_queue.put(f"Error updating sheet {sheet_name}: {e}")  # Log any errors

        current_entry_index += 1
        progress_value = int(((current_entry_index + 1) / (total_entries + 12)) * 100)  # Calculate the progress value
        progress_queue.put(progress_value)  # Update the progress queue

    log_queue.put("Smartsheets have been updated successfully.")  # Log the completion

# GUI setup
app = tk.Tk()
app.title("PCDC Reporter")

# Dark mode styles
bg_color = "#2e2e2e"
fg_color = "#ffffff"
btn_color = "#444444"
btn_fg_color = "#ffffff"
highlight_color = "#00ff00"
completed_color = "#006400"

app.configure(bg=bg_color)  # Configure the background color of the app

start_button = tk.Button(app, text="Start", command=start_button_handler, bg=btn_color, fg=btn_fg_color, activebackground=highlight_color)
start_button.pack(pady=10)  # Add the start button to the app

style = Style()
style.theme_use('clam')  # Use the 'clam' theme for the style
style.configure("green.Horizontal.TProgressbar", troughcolor=bg_color, background=highlight_color, thickness=10)  # Configure the progress bar style

progress = Progressbar(app, orient=tk.HORIZONTAL, length=300, mode='determinate', style="green.Horizontal.TProgressbar")  # Create the progress bar
progress_label = tk.StringVar()
progress_label.set("0%")  # Initialize the progress label to 0%
progress_label_widget = tk.Label(app, textvariable=progress_label, bg=bg_color, fg=highlight_color)  # Create the progress label widget

progress.pack(pady=10)  # Add the progress bar to the app
progress_label_widget.pack()  # Add the progress label widget to the app

current_log = tk.StringVar()
log_label = tk.Label(app, textvariable=current_log, justify=tk.LEFT, anchor="w", bg=bg_color, fg=fg_color)  # Create the log label
log_label.pack(pady=10)  # Add the log label to the app

app.geometry("400x150")  # Set the geometry of the app
icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')  # Get the path to the app icon
app.iconbitmap(icon_path)  # Set the app icon

log_queue = Queue()  # Create the log queue
progress_queue = Queue()  # Create the progress queue
app.mainloop()  # Start the main loop of the app
