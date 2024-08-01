import re
import os
from datetime import datetime, timedelta

def process_file(file_path, date_str, results_list):
    # Read the content of the file with the correct encoding
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()

    # Initialize a list to store the matched numbers
    completed_numbers = []

    # Precompile the regex for efficiency
    regex = re.compile(r'(\d{1,3}(?:,\d{3})*)#')

    # Process lines only when "Completed" is found
    for i, line in enumerate(lines):
        if "Completed" in line:
            for next_line in lines[i + 1:i + 5]:
                match = regex.search(next_line)
                if match:
                    completed_numbers.append(int(match.group(1).replace(',', '')))

    # Calculate the total sum
    total_sum = sum(completed_numbers) / 2000

    # Round the tonnage value
    tonnage = round(total_sum, 2)

    # Construct the new file name
    new_file_name = f"{date_str} {tonnage} Tons.txt"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

    # Check for any existing file with the same date and overwrite it
    for existing_file in os.listdir(os.path.dirname(file_path)):
        if existing_file.startswith(date_str):
            os.remove(os.path.join(os.path.dirname(file_path), existing_file))
            break

    # Rename the file
    os.rename(file_path, new_file_path)
    print(f"File renamed to: {new_file_name}")

    # Add the result to the list
    results_list.append(new_file_path.replace(r"N:\Production\Production reports\\", ""))

def process_shipping_file(file_path, date_str, results_list):
    # Read the content of the file with the correct encoding
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()

    # Initialize a variable to store the total weight
    total_weight = 0

    # Precompile the regex for efficiency
    regex = re.compile(r'Total shipped to Jobsite:\s+\d+,\d+\s+(\d{1,3}(?:,\d{3})*)#')

    # Process lines to find the total shipped weight
    for line in lines:
        match = regex.search(line)
        if match:
            total_weight = int(match.group(1).replace(',', ''))
            break

    # Calculate the total shipped weight in tons
    total_weight_tons = total_weight / 2000

    # Round the tonnage value
    tonnage = round(total_weight_tons, 2)

    # Construct the new file name
    new_file_name = f"{date_str} {tonnage} Tons.txt"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

    # Rename the file
    os.rename(file_path, new_file_path)
    print(f"File renamed to: {new_file_name}")

    # Add the result to the list
    results_list.append(new_file_path.replace(r"N:\Production\Production reports\\", ""))

def find_and_process_files(base_path, log_file_path):
    # Get the current date in YYYY-MM-DD format
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Get yesterday's date in YYYY-MM-DD format
    yesterday_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    # Initialize a list to store results
    results_list = []

    # Walk through the directory and its subdirectories
    for dirpath, _, filenames in os.walk(base_path):
        for filename in filenames:
            # Delete files that are not "Station Summary-currentday.txt", "Station Summary-prevday.txt", "ShippingList_by_Job-currentday.txt" or "ShippingList_by_Job-prevday.txt"
            if filename not in ["Station Summary-currentday.txt", "Station Summary-prevday.txt", "ShippingList_by_Job-currentday.txt", "ShippingList_by_Job-prevday.txt"]:
                os.remove(os.path.join(dirpath, filename))
            elif filename == "Station Summary-currentday.txt":
                process_file(os.path.join(dirpath, filename), current_date, results_list)
            elif filename == "Station Summary-prevday.txt":
                process_file(os.path.join(dirpath, filename), yesterday_date, results_list)
            elif filename == "ShippingList_by_Job-currentday.txt":
                process_shipping_file(os.path.join(dirpath, filename), current_date, results_list)
            elif filename == "ShippingList_by_Job-prevday.txt":
                process_shipping_file(os.path.join(dirpath, filename), yesterday_date, results_list)

    # Clear the log file and write all the results
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        for result in results_list:
            log_file.write(f"{result}\n")

# Define the base path to search
base_path = r"N:\Production\Production reports\\"
log_file_path = r"N:\Production\Production reports\Results.txt"

# Find and process the files
find_and_process_files(base_path, log_file_path)
