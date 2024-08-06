import re
import os
import smartsheet
import tkinter as tk
import threading
import logging
from datetime import datetime
from queue import Queue
from tkinter.ttk import Progressbar

# Configuration
CONFIG = {
    'access_token': os.getenv('SMARTSHEET_ACCESS_TOKEN', "FHrmEeDnC1oEBmBaWTCp1SNvtdEFbNdT2pEgc"),
    'sheet_mapping': {
        "AngleMaster": "2865724263452548",
        "Beamline": "8633118017671044",
        "Daily": "6630177199050628",
        "Paint": "6705691314048900",
        "PlateTable": "2861850202951556",
        "Shipping": "589005049515908"
    },
    'base_path': r"T:\\Production reports\\",
    'results_file_path': r"T:\\Production reports\\Results.txt"
}

# Initialize logging
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize the Smartsheet client
smartsheet_client = smartsheet.Smartsheet(CONFIG['access_token'])

def start_scripts(log_queue):
    """Start the main processing and updating script."""
    log_queue.put("Starting PDCReport process...\n")
    progress.start()
    find_and_process_files(CONFIG['base_path'], CONFIG['results_file_path'], log_queue)
    log_queue.put("PDCReport process completed. Starting Smartsheet update...\n")
    results = read_results(CONFIG['results_file_path'])
    update_smartsheets(results, CONFIG['sheet_mapping'], log_queue)
    log_queue.put("Smartsheet update completed.\nCOMPLETE")
    progress.stop()
    start_button.config(state=tk.NORMAL)

def start_button_handler():
    """Handle the start button click event."""
    start_button.config(state=tk.DISABLED)
    log_text.set("")  # Clear previous log messages
    threading.Thread(target=start_scripts, args=(log_queue,)).start()
    app.after(100, process_log_queue)

def process_log_queue():
    """Process the log queue and update the GUI."""
    while not log_queue.empty():
        log_text.set(log_text.get() + log_queue.get())
    app.after(100, process_log_queue)

def generate_unique_filename(directory, filename):
    """Generate a unique filename in the specified directory."""
    base, extension = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):
        unique_filename = f"{base} ({counter}){extension}"
        counter += 1
    return unique_filename

def extract_date_from_content(content, is_shipping=False):
    """Extract the date from the file content."""
    pattern = r'Date Shipped:\s+(\d{1,2}/\d{1,2}/\d{4})' if is_shipping else r'Date Station Completed:\s+(\d{1,2}/\d{1,2}/\d{4})'
    match = re.search(pattern, content)
    if match:
        return datetime.strptime(match.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')
    else:
        raise ValueError("Date not found in file content")

def process_file(file_path, results_list):
    """Process a general file and update the results list."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            lines = content.splitlines()

        date_str = extract_date_from_content(content, is_shipping=False)
        completed_numbers = []
        regex = re.compile(r'(\d{1,3}(?:,\d{3})*)#')

        for i, line in enumerate(lines):
            if "Completed" in line:
                for next_line in lines[i + 1:i + 5]:
                    match = regex.search(next_line)
                    if match:
                        completed_numbers.append(int(match.group(1).replace(',', '')))

        total_sum = sum(completed_numbers) / 2000
        tonnage = round(total_sum, 2)
        new_file_name = generate_unique_filename(os.path.dirname(file_path), f"{date_str} {tonnage} Tons.txt")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        os.rename(file_path, new_file_path)
        logging.info(f"File renamed to: {new_file_name}")
        results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))
    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")

def process_shipping_file(file_path, results_list):
    """Process a shipping file and update the results list."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            lines = content.splitlines()

        date_str = extract_date_from_content(content, is_shipping=True)
        total_weight = 0
        for line in lines:
            if "Total shipped to Jobsite:" in line:
                hashtag_pos = line.rfind('#')
                if hashtag_pos != -1:
                    last_space_pos = line.rfind(' ', 0, hashtag_pos)
                    if last_space_pos != -1:
                        weight_str = line[last_space_pos + 1:hashtag_pos].replace(',', '')
                        total_weight = int(weight_str)
                        break

        total_weight_tons = total_weight / 2000
        tonnage = round(total_weight_tons, 2)
        new_file_name = generate_unique_filename(os.path.dirname(file_path), f"{date_str} {tonnage} Tons.txt")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        os.rename(file_path, new_file_path)
        logging.info(f"File renamed to: {new_file_name}")
        results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))
    except Exception as e:
        logging.error(f"Error processing shipping file {file_path}: {e}")

def find_and_process_files(base_path, log_file_path, log_queue):
    """Find and process files in the specified base path."""
    results_list = []

    # Clear the results file at the start
    open(log_file_path, 'w').close()

    for dirpath, _, filenames in os.walk(base_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if filename == "Results.txt":
                continue
            if filename in ["Station Summary-currentday.txt", "Station Summary-prevday.txt"]:
                process_file(file_path, results_list)
            elif filename in ["ShippingList_by_Job-currentday.txt", "ShippingList_by_Job-prevday.txt"]:
                process_shipping_file(file_path, results_list)

    # Log results
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        for result in results_list:
            log_file.write(f"{result}\n")

    # Archive the processed files
    for result in results_list:
        new_file_path = os.path.join(base_path, result)
        archive_dir = os.path.join(os.path.dirname(new_file_path), 'Archive')
        if 'Archive' not in os.path.dirname(new_file_path):
            new_filename = generate_unique_filename(archive_dir, os.path.basename(new_file_path))
            os.makedirs(archive_dir, exist_ok=True)
            os.rename(new_file_path, os.path.join(archive_dir, new_filename))
            logging.info(f"Archived file to: {os.path.join(archive_dir, new_filename)}")
        else:
            logging.warning(f"Skipping file {new_file_path} as it is already in an archive directory.")

def read_results(file_path):
    """Read and parse the results from the specified file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.read().splitlines()
        results = []
        # Adjusted regex to match the given file paths
        for line in lines:
            match = re.search(r'([\w\s]+) throughput\\(?:Archive\\)?(\d{4}-\d{2}-\d{2}) (\d+\.\d+) Tons(?: \(\d+\))?.txt', line)
            if match:
                sheet_name = match.group(1).strip()
                date = match.group(2)
                total_weight = float(match.group(3))
                results.append((sheet_name, date, total_weight))
            else:
                logging.warning(f"Line '{line}' did not match the expected format.")
        return results
    except Exception as e:
        logging.error(f"Error reading results file {file_path}: {e}")
        return []

def get_column_id(sheet, column_title):
    """Get the column ID by column title."""
    for column in sheet.columns:
        if column.title == column_title:
            return column.id
    raise ValueError(f"Column '{column_title}' not found in the sheet.")

def update_smartsheets(results, sheet_mapping, log_queue):
    """Update the Smartsheet with the results."""
    for sheet_name, date, total_weight in results:
        try:
            log_queue.put(f"Processing {sheet_name} for date {date} with weight {total_weight}\n")
            sheet_id = sheet_mapping.get(sheet_name)
            if not sheet_id:
                log_queue.put(f"Sheet ID for '{sheet_name}' not found. Skipping this entry.\n")
                continue

            sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
            log_queue.put(f"Retrieved sheet: {sheet_name} (ID: {sheet_id})\n")

            try:
                date_column_id = get_column_id(sheet, 'Date')
                total_weight_column_id = get_column_id(sheet, 'Total Weight')
            except ValueError as e:
                log_queue.put(f"Error: {e}\n")
                continue

            rows_to_update = []

            for row in sheet.rows:
                for cell in row.cells:
                    if cell.column_id == date_column_id and cell.value == date:
                        log_queue.put(f"Found matching date: {date} in row ID: {row.id}\n")
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
                smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)
                log_queue.put(f"Updated rows in sheet: {sheet_name}\n")
            else:
                log_queue.put(f"No rows found with the date {date} in sheet {sheet_name}\n")

        except Exception as e:
            log_queue.put(f"Error updating sheet {sheet_name}: {e}\n")

    log_queue.put("Smartsheets have been updated successfully.\n")

# GUI setup
app = tk.Tk()
app.title("PCDC Reporter")

# Dark mode styles
bg_color = "#2e2e2e"
fg_color = "#ffffff"
btn_color = "#444444"
btn_fg_color = "#ffffff"
highlight_color = "#666666"

app.configure(bg=bg_color)

start_button = tk.Button(app, text="Start", command=start_button_handler, bg=btn_color, fg=btn_fg_color, activebackground=highlight_color)
start_button.pack(pady=10)

progress = Progressbar(app, orient=tk.HORIZONTAL, length=300, mode='indeterminate', style="TProgressbar")
progress.pack(pady=10)
progress.pack_forget()  # Hide progress bar initially

log_text = tk.StringVar()
log_label = tk.Label(app, textvariable=log_text, justify=tk.LEFT, anchor="w", bg=bg_color, fg=fg_color)
log_label.pack(pady=10)

app.geometry("500x300")
icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')
app.iconbitmap(icon_path)

style = tk.ttk.Style()
style.theme_use('clam')
style.configure("TProgressbar", troughcolor=bg_color, background=highlight_color, thickness=20)

log_queue = Queue()
app.mainloop()
