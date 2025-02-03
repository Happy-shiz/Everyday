import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
import datetime

class Entry:
    def __init__(self, date, title, content, owner):
        self.date = date
        self.title = title
        self.content = content
        self.owner = owner

    def __str__(self):
        return f"{self.date.strftime('%Y-%m-%d')} - {self.title} ({self.owner})"

class Task:
    def __init__(self, day, start_time, end_time, title):
        self.day = day
        self.start_time = start_time
        self.end_time = end_time
        self.title = title

    def __str__(self):
        return f"{self.start_time.strftime('%H:%M')} - {self.end_time.strftime('%H:%M')}: {self.title}"

class Diary:
    def __init__(self):
        self.entries = []

    def add_entry(self, date, title, content, owner):
        entry = Entry(date, title, content, owner)
        self.entries.append(entry)

    def view_entries(self):
        return [str(entry) for entry in self.entries]

class WeeklySchedule:
    def __init__(self):
        self.tasks = {}

    def add_task(self, task):
        if task.day not in self.tasks:
            self.tasks[task.day] = []
        self.tasks[task.day].append(task)

    def get_tasks_for_day(self, day):
        return self.tasks.get(day, [])

class CombinedGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Ежедневник и Расписание")
        self.diary = Diary()
        self.schedule = WeeklySchedule()
        self.create_widgets()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        diary_tab = ttk.Frame(self.notebook)
        self.notebook.add(diary_tab, text='Ежедневник')
        self.create_diary_tab(diary_tab)

        schedule_tab = ttk.Frame(self.notebook)
        self.notebook.add(schedule_tab, text='Расписание')
        self.create_schedule_tab(schedule_tab)

    def create_diary_tab(self, frame):
        self.diary_date = Calendar(frame, selectmode='day', year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
        self.diary_date.pack(side=tk.TOP)

        ttk.Label(frame, text="Заголовок:").pack(side=tk.TOP)
        self.diary_title = ttk.Entry(frame)
        self.diary_title.pack(side=tk.TOP)

        ttk.Label(frame, text="Содержание:").pack(side=tk.TOP)
        self.diary_content = ttk.Entry(frame)
        self.diary_content.pack(side=tk.TOP)

        ttk.Label(frame, text="Владелец:").pack(side=tk.TOP)
        self.diary_owner = ttk.Entry(frame)
        self.diary_owner.pack(side=tk.TOP)

        ttk.Button(frame, text="Добавить запись", command=self.add_diary_entry).pack(side=tk.TOP)

        self.diary_list = tk.Listbox(frame, width=100)
        self.diary_list.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def create_schedule_tab(self, frame):
        ttk.Label(frame, text="День:").pack()
        self.schedule_day = tk.StringVar()
        days = [day.strftime('%A') for day in [datetime.date.today() + datetime.timedelta(days=i) for i in range(7)]]
        self.schedule_day_combobox = ttk.Combobox(frame, textvariable=self.schedule_day, values=days, state="readonly")
        self.schedule_day_combobox.pack()

        ttk.Label(frame, text="Время начала (HH:MM):").pack()
        self.schedule_start_time = ttk.Entry(frame)
        self.schedule_start_time.pack()

        ttk.Label(frame, text="Время окончания (HH:MM):").pack()
        self.schedule_end_time = ttk.Entry(frame)
        self.schedule_end_time.pack()

        ttk.Label(frame, text="Заголовок задачи:").pack()
        self.schedule_task_title = ttk.Entry(frame)
        self.schedule_task_title.pack()

        ttk.Button(frame, text="Добавить задачу", command=self.add_schedule_task).pack()

        self.schedule_list = tk.Listbox(frame, width=50)
        self.schedule_list.pack(fill=tk.BOTH, expand=True)

    def add_diary_entry(self):
        date = self.diary_date.selection_get()
        title = self.diary_title.get()
        content = self.diary_content.get()
        owner = self.diary_owner.get()
        self.diary.add_entry(date, title, content, owner)
        self.update_diary_list()

    def update_diary_list(self):
        self.diary_list.delete(0, tk.END)
        for item in self.diary.view_entries():
            self.diary_list.insert(tk.END, item)

    def add_schedule_task(self):
        try:
            day = next(date for date in [datetime.date.today() + datetime.timedelta(days=i) for i in range(7)] if date.strftime('%A') == self.schedule_day.get())
            start_time = datetime.datetime.strptime(self.schedule_start_time.get(), '%H:%M').time()
            end_time = datetime.datetime.strptime(self.schedule_end_time.get(), '%H:%M').time()
            title = self.schedule_task_title.get()

            task = Task(day, start_time, end_time, title)
            self.schedule.add_task(task)
            self.update_schedule_list()
        except ValueError:
            pass

    def update_schedule_list(self):
        self.schedule_list.delete(0, tk.END)
        for task in self.schedule.get_tasks_for_day(datetime.date.today()):
            self.schedule_list.insert(tk.END, str(task))

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedGUI(root)
    root.mainloop()
    