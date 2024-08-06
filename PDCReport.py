import re
import os
from datetime import datetime, timedelta

def generate_unique_filename(directory, filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):
        unique_filename = f"{base} ({counter}){extension}"
        counter += 1
    return unique_filename

def process_file(file_path, date_str, results_list):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()

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
    new_file_name = f"{date_str} {tonnage} Tons.txt"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

    for existing_file in os.listdir(os.path.dirname(file_path)):
        if existing_file.startswith(date_str):
            os.remove(os.path.join(os.path.dirname(file_path), existing_file))
            break

    new_file_name = generate_unique_filename(os.path.dirname(file_path), new_file_name)
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    os.rename(file_path, new_file_path)
    print(f"File renamed to: {new_file_name}")
    results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))

def process_shipping_file(file_path, date_str, results_list):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.read().splitlines()

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
    new_file_name = f"Shipping_{date_str} {tonnage} Tons.txt"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

    new_file_name = generate_unique_filename(os.path.dirname(file_path), new_file_name)
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    os.rename(file_path, new_file_path)
    print(f"File renamed to: {new_file_name}")
    results_list.append(new_file_path.replace(r"T:\\Production reports\\", ""))

def find_and_process_files(base_path, log_file_path):
    current_date = datetime.now().strftime('%Y-%m-%d')
    yesterday_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    results_list = []

    for dirpath, _, filenames in os.walk(base_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if filename == "Results.txt":
                continue
            if filename == "Station Summary-currentday.txt":
                process_file(file_path, current_date, results_list)
            elif filename == "Station Summary-prevday.txt":
                process_file(file_path, yesterday_date, results_list)
            elif filename == "ShippingList_by_Job-currentday.txt":
                process_shipping_file(file_path, current_date, results_list)
            elif filename == "ShippingList_by_Job-prevday.txt":
                process_shipping_file(file_path, yesterday_date, results_list)
            else:
                archive_dir = os.path.join(dirpath, 'archive')
                if os.path.exists(archive_dir):
                    new_filename = generate_unique_filename(archive_dir, filename)
                    os.rename(file_path, os.path.join(archive_dir, new_filename))
                else:
                    print(f"Archive directory does not exist in {dirpath}. Skipping file: {filename}")

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        for result in results_list:
            log_file.write(f"{result}\n")

    for result in results_list:
        new_file_path = os.path.join(base_path, result)
        archive_dir = os.path.join(os.path.dirname(new_file_path), 'archive')
        if os.path.exists(archive_dir):
            new_filename = generate_unique_filename(archive_dir, os.path.basename(new_file_path))
            os.rename(new_file_path, os.path.join(archive_dir, new_filename))
        else:
            print(f"Archive directory does not exist for {new_file_path}. Skipping file.")

# Define the base path to search
base_path = r"T:\\Production reports\\"
log_file_path = r"T:\\Production reports\\Results.txt"

# Find and process the files
find_and_process_files(base_path, log_file_path)
