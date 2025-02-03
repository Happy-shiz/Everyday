import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from plyer import notification
import winsound
import threading
import time

# Создание или подключение к базе данных
conn = sqlite3.connect("calendar.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    description TEXT,
    date TEXT NOT NULL,
    time TEXT NOT NULL,
    category TEXT NOT NULL
)
""")
conn.commit()

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

    cursor.execute("INSERT INTO events (title, description, date, time, category) VALUES (?, ?, ?, ?, ?)", 
                   (title, description, date, time, category))
    conn.commit()
    update_monthly_view()
    update_weekly_view()
    update_daily_view()
    clear_inputs()

# Функция удаления события
def delete_event():
    selected_item = event_list.selection()
    if not selected_item:
        messagebox.showwarning("Ошибка", "Выберите событие для удаления")
        return

    event_id = event_list.item(selected_item, "values")[0]
    cursor.execute("DELETE FROM events WHERE id=?", (event_id,))
    conn.commit()
    update_monthly_view()
    update_weekly_view()
    update_daily_view()

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

    cursor.execute("SELECT date, title, category FROM events")
    events = cursor.fetchall()

    for event_date, title, category in events:
        label = tk.Label(monthly_frame, text=f"{event_date}: {title}", bg=CATEGORY_COLORS.get(category, "#ffffff"))
        label.pack(anchor="w", padx=5, pady=2)

# Функция обновления недельного календаря
def update_weekly_view():
    for widget in weekly_frame.winfo_children():
        widget.destroy()

    cursor.execute("SELECT date, title, category FROM events WHERE date BETWEEN DATE('now', '-7 days') AND DATE('now', '+7 days')")
    events = cursor.fetchall()

    for event_date, title, category in events:
        label = tk.Label(weekly_frame, text=f"{event_date}: {title}", bg=CATEGORY_COLORS.get(category, "#ffffff"))
        label.pack(anchor="w", padx=5, pady=2)

# Функция обновления дневного календаря
def update_daily_view():
    for widget in daily_frame.winfo_children():
        widget.destroy()

    cursor.execute("SELECT time, title, category FROM events WHERE date = DATE('now')")
    events = cursor.fetchall()

    for event_time, title, category in events:
        label = tk.Label(daily_frame, text=f"{event_time}: {title}", bg=CATEGORY_COLORS.get(category, "#ffffff"))
        label.pack(anchor="w", padx=5, pady=2)

# Функция проверки и отправки уведомлений
def check_notifications():
    while True:
        now = time.strftime("%Y-%m-%d %H:%M")
        cursor.execute("SELECT title FROM events WHERE date || ' ' || time = ?", (now,))
        event = cursor.fetchone()

        if event:
            notification.notify(
                title="Напоминание",
                message=f"Событие: {event[0]}",
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
root.geometry("800x600")

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
daily_frame = tk.Frame(notebook)

notebook.add(monthly_frame, text="Месяц")
notebook.add(weekly_frame, text="Неделя")
notebook.add(daily_frame, text="День")

# Поля ввода
title_entry = tk.Entry(root)
title_entry.pack()

description_entry = tk.Text(root, height=3)
description_entry.pack()

date_entry = tk.Entry(root)
date_entry.pack()
tk.Button(root, text="📅 Выбрать дату", command=pick_date).pack()

time_entry = tk.Entry(root)
time_entry.pack()

category_var = tk.StringVar()
category_dropdown = ttk.Combobox(root, textvariable=category_var, values=list(CATEGORY_COLORS.keys()))
category_dropdown.pack()

tk.Button(root, text="Добавить событие", command=add_event).pack()
tk.Button(root, text="Удалить событие", command=delete_event).pack()

# Таблица событий
columns = ("id", "Название", "Описание", "Дата", "Время", "Категория")
event_list = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    event_list.heading(col, text=col)
event_list.pack(fill="both", expand=True)

update_monthly_view()
update_weekly_view()
update_daily_view()

root.mainloop()

conn.close()