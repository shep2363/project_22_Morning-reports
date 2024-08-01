import re
import os
from datetime import datetime, timedelta

def process_file(file_path, date_str, results_list, is_shipping=False):
    completed_numbers = []
    total_weight = 0
    
    # Precompile the regex for efficiency
    completed_regex = re.compile(r'(\d{1,3}(?:,\d{3})*)#')
    shipping_regex = re.compile(r'Total shipped to Jobsite:\s+\d+,\d+\s+(\d{1,3}(?:,\d{3})*)#')

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()
        if is_shipping:
            for line in lines:
                match = shipping_regex.search(line)
                if match:
                    total_weight = int(match.group(1).replace(',', ''))
                    break
        else:
            for i, line in enumerate(lines):
                if "Completed" in line:
                    for next_line in lines[i + 1:i + 5]:
                        match = completed_regex.search(next_line)
                        if match:
                            completed_numbers.append(int(match.group(1).replace(',', '')))
    
    if is_shipping:
        tonnage = round(total_weight / 2000, 2)
    else:
        tonnage = round(sum(completed_numbers) / 2000, 2)
    
    new_file_name = f"{date_str} {tonnage} Tons.txt"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    
    os.rename(file_path, new_file_path)
    print(f"File renamed to: {new_file_name}")
    
    results_list.append(new_file_path.replace(r"N:\Production\Production reports\\", ""))

def find_and_process_files(base_path, log_file_path):
    current_date = datetime.now().strftime('%Y-%m-%d')
    yesterday_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    results_list = []

    for dirpath, _, filenames in os.walk(base_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if filename == "Station Summary-currentday.txt":
                process_file(file_path, current_date, results_list)
            elif filename == "Station Summary-prevday.txt":
                process_file(file_path, yesterday_date, results_list)
            elif filename == "ShippingList_by_Job-currentday.txt":
                process_file(file_path, current_date, results_list, is_shipping=True)
            elif filename == "ShippingList_by_Job-prevday.txt":
                process_file(file_path, yesterday_date, results_list, is_shipping=True)
            else:
                os.remove(file_path)

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write("\n".join(results_list) + "\n")

# Define the base path to search
base_path = r"N:\Production\Production reports\\"
log_file_path = r"N:\Production\Production reports\Results.txt"

# Find and process the files
find_and_process_files(base_path, log_file_path)
