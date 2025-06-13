import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import subprocess
import sys
import os
try:
    from dotenv import dotenv_values, load_dotenv
except:
    from dotenv import main


local_path=os.getcwd()

def generate():
    year = entry_year.get()
    week = entry_week.get()
    
    result = subprocess.run([sys.executable, 'main.py', '--year', year, '--week', week], capture_output=True)
    
    # Mostra l'output nello stesso Tkinter Text Widget
    output_text.delete(1.0, tk.END)  # Pulisci il widget di testo
    output_text.insert(tk.END, f"Output:\n{result.stdout.decode('utf-8')}\n Errors:\n{result.stderr.decode('utf-8')}")


def populate_years_weeks(start_date, end_date, entry_year,entry_week):
    years = set([2022,2023,2024])
    weeks = set()

    current = start_date
    while current <= end_date:
        iso_year, iso_week, _ = current.isocalendar()
        if iso_year <= datetime.today().date().year:
            years.add(iso_year)
            weeks.add(iso_week)
        current += timedelta(days=1)

    entry_year['values'] = sorted(years, reverse=True)
    entry_week['values'] = sorted(weeks)

    if years:
        entry_year.current(0)
    if weeks:
        entry_week.current(0)   

# Main window generation
root = tk.Tk()
root.title("Weekly Report GUI")
root.iconbitmap()
root.iconbitmap(f"{local_path}\\images\\ajt_official.ico")

# Fields to insert the week and the year
ttk.Label(root, text="Year:").grid(row=0, column=0, padx=5, pady=5) 
entry_year=ttk.Combobox(root, state="readonly")
entry_year.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(root, text="Week:").grid(row=1, column=0, padx=5, pady=5) 
entry_week = ttk.Combobox(root, state="readonly")
entry_week.grid(row=1, column=1, padx=5, pady=5)




today = datetime.today().date()
start_of_year = datetime(today.year, 1, 1).date()
end_of_year = datetime(today.year, 12, 31).date()

populate_years_weeks(start_of_year, end_of_year, entry_year,entry_week)

# Button to execute the script
run_button = ttk.Button(root, text="Generate Report ", command=generate) 
run_button.grid(row=2, column=0, columnspan=2, pady=10)


# Widget to the output
output_text = tk.Text(root, height=10, width=50) 
output_text.grid(row=3, column=0, columnspan=2, pady=10)

root.mainloop()


