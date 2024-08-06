import os

# List of paths to check
paths_to_check = [
    r"T:\\Production reports\\",
    r"T:\\Production reports\\Archive\\",
    r"T:\\Production reports\\Results.txt"
]

# Check each path
for path in paths_to_check:
    if os.path.exists(path):
        print(f"Path exists: {path}")
    else:
        print(f"Path does NOT exist: {path}")

# Check if the error-inducing file path is valid
error_file_path = r"T:\\Production reports\\AngleMaster throughput\\Archive\\2024-08-01 1.27 Tons.txt"
if os.path.exists(error_file_path):
    print(f"Error-inducing file path exists: {error_file_path}")
else:
    print(f"Error-inducing file path does NOT exist: {error_file_path}")
