import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from plyer import notification
import winsound
import threading
import time
from datetime import datetime

# Создание или подключение к Excel файлу
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

# Цвета категорий
CATEGORY_COLORS = {
    "Работа": "#ffcc00",
    "Семья": "#ff6699",
    "Здоровье": "#33cc33",
    "Тренировки": "#3399ff",
    "Обучение": "#9933ff",
    "События": "#ff5050"
}

# Функция добавления события
def add_event():
    title = title_entry.get()
    description = description_entry.get("1.0", tk.END).strip()
    date = date_entry.get()
    time = time_entry.get()
    category = category_var.get()

    if not title or not date or not time or not category:
        messagebox.showwarning("Ошибка", "Заполните все обязательные поля!")
        return

    # Преобразуем строку в объект datetime.date
    try:
        event_date = datetime.strptime(date, "%Y-%m-%d").date()
    except ValueError:
        messagebox.showwarning("Ошибка", "Некорректный формат даты!")
        return

    # Запись события в Excel
    last_row = sheet.max_row + 1
    sheet.append([last_row, title, description, event_date, time, category])
    wb.save("calendar.xlsx")
    
    update_monthly_view()
    update_weekly_view()
    clear_inputs()

# Функция удаления события
def delete_event():
    selected_item = event_list.selection()
    if not selected_item:
        messagebox.showwarning("Ошибка", "Выберите событие для удаления")
        return

    event_id = event_list.item(selected_item, "values")[0]

    # Поиск и удаление события в Excel
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[0].value == int(event_id):
            sheet.delete_rows(row[0].row)
            wb.save("calendar.xlsx")
            break

    update_monthly_view()
    update_weekly_view()

# Очистка полей ввода
def clear_inputs():
    title_entry.delete(0, tk.END)
    description_entry.delete("1.0", tk.END)
    date_entry.delete(0, tk.END)
    time_entry.delete(0, tk.END)

# Выбор даты через календарь
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

# Функция обновления месячного календаря
def update_monthly_view():
    for widget in monthly_frame.winfo_children():
        widget.destroy()

    # Отображение месячного календаря
    cal = Calendar(monthly_frame, selectmode="day", date_pattern="yyyy-mm-dd", showweeknumbers=False)
    cal.pack(padx=10, pady=10)

    # Чтение событий из Excel
    events = {}
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        event_date, title, category = row[3].value, row[1].value, row[5].value
        
        # Если event_date уже является объектом datetime.date, то не нужно его снова преобразовывать
        if isinstance(event_date, datetime):
            event_date = event_date.date()

        if event_date not in events:
            events[event_date] = []
        events[event_date].append(f"{title} ({category})")

    # Пометка событий на календаре
    for event_date, event_list in events.items():
        for event in event_list:
            cal.calevent_create(event_date, event, tags=("event"))

# Функция обновления недельного календаря
def update_weekly_view():
    for widget in weekly_frame.winfo_children():
        widget.destroy()

    # Чтение событий из Excel за неделю
    cursor_date = time.strftime("%Y-%m-%d")
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        event_date, event_time, event_title, event_category = row[3].value, row[4].value, row[1].value, row[5].value
        if event_date == cursor_date:  # Сравнение текущей даты с датой события
            event_day = time.strptime(event_date, "%Y-%m-%d").tm_wday  # Получаем день недели
            event_hour = int(event_time.split(":")[0])  # Час начала события

            label = tk.Label(weekly_frame, text=event_title, bg=CATEGORY_COLORS.get(event_category, "#ffffff"), height=2, relief="solid")
            label.grid(row=event_hour + 1, column=event_day + 1, sticky="nsew")

    # Подгонка размеров ячеек
    for row in range(1, 25):
        weekly_frame.grid_rowconfigure(row, weight=1)
    for col in range(1, 8):
        weekly_frame.grid_columnconfigure(col, weight=1)

# Функция проверки и отправки уведомлений
def check_notifications():
    while True:
        now = time.strftime("%Y-%m-%d %H:%M")
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            event_date, event_time, event_title = row[3].value, row[4].value, row[1].value
            if f"{event_date} {event_time}" == now:
                notification.notify(
                    title="Напоминание",
                    message=f"Событие: {event_title}",
                    app_name="Офлайн Календарь",
                    timeout=10
                )
                winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)

        time.sleep(30)

# Запуск фонового потока уведомлений
threading.Thread(target=check_notifications, daemon=True).start()

# Создание GUI
root = tk.Tk()
root.title("Офлайн Календарь")
root.geometry("1000x600")

# Боковая панель категорий
sidebar = tk.Frame(root, width=150, bg="#333")
sidebar.pack(side="left", fill="y")

tk.Label(sidebar, text="Категории", bg="#333", fg="white").pack(pady=10)
for category, color in CATEGORY_COLORS.items():
    tk.Label(sidebar, text=category, bg=color, fg="black", padx=10).pack(pady=2)

# Основное окно с вкладками
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# Вкладки
monthly_frame = tk.Frame(notebook)
weekly_frame = tk.Frame(notebook)

notebook.add(monthly_frame, text="Месяц")
notebook.add(weekly_frame, text="Неделя")

# Поля ввода с подписями
title_label = tk.Label(root, text="Название события:")
title_label.pack()
title_entry = tk.Entry(root)
title_entry.pack()

description_label = tk.Label(root, text="Описание события:")
description_label.pack()
description_entry = tk.Text(root, height=3)
description_entry.pack()

date_label = tk.Label(root, text="Дата события:")
date_label.pack()
date_entry = tk.Entry(root)
date_entry.pack()
tk.Button(root, text="📅 Выбрать дату", command=pick_date).pack()

time_label = tk.Label(root, text="Время события:")
time_label.pack()
time_entry = tk.Entry(root)
time_entry.pack()

category_label = tk.Label(root, text="Категория события:")
category_label.pack()
category_var = tk.StringVar()
category_dropdown = ttk.Combobox(root, textvariable=category_var, values=list(CATEGORY_COLORS.keys()))
category_dropdown.pack()

tk.Button(root, text="Добавить событие", command=add_event).pack()
tk.Button(root, text="Удалить событие", command=delete_event).pack()
update_monthly_view()
update_weekly_view()

root.mainloop()

wb.save("calendar.xlsx")