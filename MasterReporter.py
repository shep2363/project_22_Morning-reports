import tkinter as tk
from tkinter import messagebox
import subprocess
import threading

def run_script(script_path):
    try:
        result = subprocess.run(['python', script_path], check=True, capture_output=True, text=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        return e.stderr

def start_scripts():
    log_text.set("Starting PDCReportV2.py...\n")
    update_ui()
    pdc_report_result = run_script("PDCReportV2.py")
    
    log_text.set(log_text.get() + pdc_report_result + "\n")
    update_ui()

    if "Error" in pdc_report_result or "Traceback" in pdc_report_result:
        messagebox.showerror("Error", "An error occurred while running PDCReportV2.py")
        return
    
    log_text.set(log_text.get() + "PDCReportV2.py completed. Starting SmartsheetUpdate.py...\n")
    update_ui()
    smartsheet_update_result = run_script("SmartsheetUpdate.py")
    
    log_text.set(log_text.get() + smartsheet_update_result + "\n")
    update_ui()

    if "Error" in smartsheet_update_result or "Traceback" in smartsheet_update_result:
        messagebox.showerror("Error", "An error occurred while running SmartsheetUpdate.py")
        return
    
    log_text.set(log_text.get() + "SmartsheetUpdate.py completed.\nCOMPLETE")
    update_ui()
    messagebox.showinfo("Complete", "All scripts completed successfully.")

def start_button_handler():
    threading.Thread(target=start_scripts).start()

def update_ui():
    log_label.update()

app = tk.Tk()
app.title("Script Runner")

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
app.mainloop()
