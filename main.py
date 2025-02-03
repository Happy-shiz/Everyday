import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
from collections import defaultdict
import calendar

# Create or connect to the Excel file
def create_excel():
    try:
        wb = openpyxl.load_workbook("calendar.xlsx")
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "События"
        sheet.append(["id", "title", "description", "date", "time", "category"])
        wb.save("calendar.xlsx")
    return wb

wb = create_excel()
sheet = wb["События"]

# Colors for categories
CATEGORY_COLORS = {
    "Работа": "#ffcc00",
    "Семья": "#ff6699",
    "Здоровье": "#33cc33",
    "Тренировки": "#3399ff",
    "Обучение": "#9933ff",
    "События": "#ff5050"
}

# Function to add an event
def add_event():
    title = title_entry.get()
    description = description_entry.get("1.0", tk.END).strip()
    date = date_entry.get()
    time = time_entry.get()
    category = category_var.get()
    if not title or not date or not time or not category:
        messagebox.showwarning("Ошибка", "Заполните все обязательные поля!")
        return
    # Convert string to datetime.date object
    try:
        event_date = datetime.strptime(date, "%Y-%m-%d").date()
    except ValueError:
        messagebox.showwarning("Ошибка", "Некорректный формат даты!")
        return
    # Write event to Excel
    last_row = sheet.max_row + 1
    sheet.append([last_row, title, description, event_date, time, category])
    wb.save("calendar.xlsx")
    
    update_monthly_view()
    clear_inputs()

# Function to delete an event with confirmation, now linked to event labels
def delete_event(event_id):
    event_title = next((event[1] for event in events if event[0] == event_id), None)
    if event_title:
        confirm = messagebox.askyesno("Подтверждение удаления", f"Вы уверены, что хотите удалить событие '{event_title}'?")
        if confirm:
            # Search and delete event in Excel
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                if row[0].value == event_id:
                    sheet.delete_rows(row[0].row)
                    wb.save("calendar.xlsx")
                    break
            update_monthly_view()

# Function to clear input fields
def clear_inputs():
    title_entry.delete(0, tk.END)
    description_entry.delete("1.0", tk.END)
    date_entry.delete(0, tk.END)
    time_entry.delete(0, tk.END)

# Function to pick date
def pick_date():
    def set_date():
        date_entry.delete(0, tk.END)
        date_entry.insert(0, cal.get_date())
        top.destroy()
    top = tk.Toplevel(root)
    top.title("Выберите дату")
    cal = Calendar(top, date_pattern="yyyy-mm-dd")
    cal.pack(pady=20)
    tk.Button(top, text="Выбрать", command=set_date).pack(pady=10)

# Function to update monthly view with clickable events
def update_monthly_view():
    for widget in monthly_frame.winfo_children():
        widget.destroy()

    # Dictionary to store events by date
    global events  # Declare events as global to use it in delete_event function
    events = defaultdict(list)

    # Read events from Excel
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        event_date, title, category = row[3].value, row[1].value, row[5].value
        if isinstance(event_date, datetime):
            event_date = event_date.date()
        events[event_date].append((row[0].value, f"{title} - {category}"))

    # Create header with days of the week
    days_of_week = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    day_header = '   '.join([day.ljust(4) for day in days_of_week])
    tk.Label(monthly_frame, text=day_header, font=("Courier", 12, "bold")).pack()

    # Get current year and month
    current_month = selected_month
    current_year = selected_year

    # Get the first day of the month and the number of days in the month
    first_day_of_month = datetime(current_year, current_month, 1)
    _, num_days_in_month = calendar.monthrange(current_year, current_month)

    # Create a list of days for the current month
    days_in_month = [i for i in range(1, num_days_in_month + 1)]
    
    # Format the calendar in weeks
    week = []
    for i, day in enumerate(days_in_month):
        week.append(day)
        if (i + 1) % 7 == 0 or i == len(days_in_month) - 1:
            # Print the week
            week_str = ''
            for day in week:
                day_date = datetime(current_year, current_month, day)
                events_for_day = events.get(day_date.date(), [])
                num_tasks = len(events_for_day)
                # Use ljust() to align all numbers and labels
                week_str += f"{str(day).ljust(3)}({str(num_tasks).rjust(2)})   "
            tk.Label(monthly_frame, text=week_str, font=("Courier", 10)).pack()
            week = []

# Function to switch to the previous year
def prev_year():
    global selected_year
    selected_year -= 1
    update_monthly_view()
    year_label.config(text=f"Год: {selected_year}")

# Function to switch to the next year
def next_year():
    global selected_year
    selected_year += 1
    update_monthly_view()
    year_label.config(text=f"Год: {selected_year}")

xui=int(input(selected_month))

def prev_month_num():
    global selected_month
    selected_month -= 1
    update_monthly_view()
    month_label.config(text=f"Месяц: {selected_month}")

def next_month_num():
    global selected_month
    selected_month += 1
    update_monthly_view()
    month_label.config(text=f"Месяц: {selected_month}")

# Function to switch to the previous month
def prev_month():
    xui -= 1
    global selected_month
    if selected_month == 1:
        selected_month = 12
        prev_year()
    else:
        selected_month -= 1
    update_monthly_view()

# Function to switch to the next month
def next_month():
    global selected_month
    if selected_month == 12:
        selected_month = 1
        next_year()
    else:
        selected_month += 1
    update_monthly_view()

# Create GUI
root = tk.Tk()
root.title("Офлайн Календарь")
root.geometry("800x600")

# Set initial year and month
selected_year = datetime.now().year
selected_month = datetime.now().month

# Monthly view frame
monthly_frame = tk.Frame(root)
monthly_frame.pack(expand=True, fill="both", padx=10, pady=10)

# Navigation buttons for year and month
year_nav_frame = tk.Frame(root)
year_nav_frame.pack(pady=10)

prev_year_btn = tk.Button(year_nav_frame, text="<<", command=prev_year)
prev_year_btn.pack(side="left", padx=5)
year_label = tk.Label(year_nav_frame, text=f"Год: {selected_year}", font=("Arial", 12, "bold"))
year_label.pack(side="left")
next_year_btn = tk.Button(year_nav_frame, text=">>", command=next_year)
next_year_btn.pack(side="left", padx=5)

month_nav_frame = tk.Frame(root)
month_nav_frame.pack(pady=5)


prev_month_btn = tk.Button(month_nav_frame, text="<", command=prev_month,)
prev_month_btn.pack(side="left", padx=5)
month_label = tk.Label(month_nav_frame, text=f"Месяц: {calendar.month_name[selected_month]}", font=("Arial", 12, "bold"))
month_label.pack(side="left")
next_month_btn = tk.Button(month_nav_frame, text=">", command=next_month,)
next_month_btn.pack(side="left", padx=5)

# Input fields with labels
tk.Label(root, text="Название события:").pack()
title_entry = tk.Entry(root)
title_entry.pack()
tk.Label(root, text="Описание события:").pack()
description_entry = tk.Text(root, height=3, width=50)
description_entry.pack()
tk.Label(root, text="Дата события:").pack()
date_entry = tk.Entry(root)
date_entry.pack()
tk.Button(root, text="Выбрать дату", command=pick_date).pack()
tk.Label(root, text="Время события:").pack()
time_entry = tk.Entry(root)
time_entry.pack()
tk.Label(root, text="Категория события:").pack()
category_var = tk.StringVar()
category_dropdown = ttk.Combobox(root, textvariable=category_var, values=list(CATEGORY_COLORS.keys()))
category_dropdown.pack()
tk.Button(root, text="Добавить событие", command=add_event).pack()

# Button to refresh the calendar view after deletions
tk.Button(root, text="Обновить календарь", command=update_monthly_view).pack()

update_monthly_view()
root.mainloop()
wb.save("calendar.xlsx")