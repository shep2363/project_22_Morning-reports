import tkinter as tk
import subprocess
import threading
import os
from queue import Queue

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
pdc_report_path = os.path.join(script_dir, 'PDCReport.py')
smartsheet_update_path = os.path.join(script_dir, 'SmartsheetUpdate.py')
icon_path = os.path.join(script_dir, 'icon.ico')

def run_script(script_path):
    try:
        result = subprocess.run(['python', script_path], check=True, capture_output=True, text=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        return e.stderr

def start_scripts(log_queue):
    log_queue.put("Starting PDCReport.py...\n")
    pdc_report_result = run_script(pdc_report_path)
    log_queue.put(pdc_report_result + "\n")

    if "Error" in pdc_report_result or "Traceback" in pdc_report_result:
        log_queue.put("An error occurred while running PDCReport.py\n")
        return

    log_queue.put("PDCReport.py completed. Starting SmartsheetUpdate.py...\n")
    smartsheet_update_result = run_script(smartsheet_update_path)
    log_queue.put(smartsheet_update_result + "\n")

    if "Error" in smartsheet_update_result or "Traceback" in smartsheet_update_result:
        log_queue.put("An error occurred while running SmartsheetUpdate.py\n")
        return

    log_queue.put("SmartsheetUpdate.py completed.\nCOMPLETE")

def start_button_handler():
    threading.Thread(target=start_scripts, args=(log_queue,)).start()
    app.after(100, process_log_queue)

def process_log_queue():
    while not log_queue.empty():
        log_text.set(log_text.get() + log_queue.get())
    app.after(100, process_log_queue)

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

log_text = tk.StringVar()
log_label = tk.Label(app, textvariable=log_text, justify=tk.LEFT, anchor="w", bg=bg_color, fg=fg_color)
log_label.pack(pady=10)
app.geometry("600x400")
app.iconbitmap(icon_path)
log_queue = Queue()
app.mainloop()