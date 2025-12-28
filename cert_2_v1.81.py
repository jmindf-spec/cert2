import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import pandas as pd
import re
import os
from datetime import datetime
import numpy as np


# классы из условий ДЗ
class Employee:
    def __init__(self, name, position, salary, hours_worked=0, emp_id=None):
        self.id = emp_id
        self.name = name
        self.position = position
        self.salary = float(salary)
        self.hours_worked = float(hours_worked)
    
    def calculate_pay(self):
        hourly_rate = self.salary / 160
        return hourly_rate * self.hours_worked
    
    def update_hours_worked(self, db_connection):
        query = """
            SELECT COALESCE(SUM(hours_required), 0) 
            FROM tasks 
            WHERE employee_id = %s AND status = 'Завершено'
        """
        result = db_connection.execute_query(query, (self.id,), fetch=True)
        if result:
            self.hours_worked = float(result[0][0])


class Task:
    def __init__(self, title, description, status="В процессе", assigned_employee=None,
                 hours_required=0, project_id=None, task_id=None):
        self.id = task_id
        self.title = title
        self.description = description
        self.status = status
        self.assigned_employee = assigned_employee
        self.hours_required = float(hours_required) if hours_required else 0.0
        self.project_id = project_id
    
    def mark_complete(self):
        old_status = self.status
        self.status = "Завершено"
        return old_status != "Завершено"


class Project:
    def __init__(self, title, tasks=None, project_id=None):
        self.id = project_id
        self.title = title
        self.tasks = tasks or []
    
    def add_task(self, task):
        self.tasks.append(task)
        task.project_id = self.id
    
    def project_progress(self):
        if not self.tasks:
            return 0
        completed = sum(1 for task in self.tasks if task.status == "Завершено")
        return (completed / len(self.tasks)) * 100
    
    def to_dict(self):
        total_tasks = len(self.tasks)
        completed_tasks = sum(1 for task in self.tasks if task.status == "Завершено")
        progress = self.project_progress()
        
        return {
            'id': self.id,
            'title': self.title,
            'total_tasks': total_tasks,
            'completed_tasks': completed_tasks,
            'progress': f"{progress:.1f}%"
        }


# Функции для работы с данными
def extract_emails(text):
    """Находит все email-адреса в строке"""
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(pattern, text)


def clean_csv_data(df):
    """Очищает csv от строк с пустыми значениями"""
    df = df.replace(r'^\s*$', np.nan, regex=True) # Без замены на NaN строки с пропусками не удалялись. Михаил, имеет ли смысл работать со строками?
    df_cleaned = df.dropna()
    return df_cleaned

# Работа с БД
class Database:
    def __init__(self):
        self.connection = None
        self.connect()
    
    def connect(self):
        try:
            self.connection = psycopg2.connect(
                host="localhost",
                database="work_time_tracking",
                user="postgres",
                password="12345"
            )
            print("Успешное подключение к БД")
        except Exception as e:
            print(f"Ошибка подключения к БД: {e}")
    
    def execute_query(self, query, params=None, fetch=False):
        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params or ())
            if fetch:
                result = cursor.fetchall()
            else:
                self.connection.commit()
                result = None
            cursor.close()
            return result
        except Exception as e:
            self.connection.rollback()
            print(f"Ошибка выполнения запроса: {e}")
            return None
    
    # Дополнительные методы для сотрудников
    def get_all_employees(self):
        query = "SELECT id, name, position, salary FROM employees ORDER BY id"
        rows = self.execute_query(query, fetch=True)
        employees = []
        for row in rows:
            emp = Employee(row[1], row[2], row[3], 0, row[0])
            emp.update_hours_worked(self)
            employees.append(emp)
        return employees
    
    def get_employee_by_id(self, emp_id):
        query = "SELECT id, name, position, salary FROM employees WHERE id = %s"
        result = self.execute_query(query, (emp_id,), fetch=True)
        if result:
            row = result[0]
            emp = Employee(row[1], row[2], row[3], 0, row[0])
            emp.update_hours_worked(self)
            return emp
        return None
    
    def get_employee_hours_worked(self, emp_id):
        query = """
            SELECT COALESCE(SUM(hours_required), 0) 
            FROM tasks 
            WHERE employee_id = %s AND status = 'Завершено'
        """
        result = self.execute_query(query, (emp_id,), fetch=True)
        return float(result[0][0]) if result else 0.0
    
    def add_employee(self, employee):
        query = """
            INSERT INTO employees (name, position, salary)
            VALUES (%s, %s, %s) RETURNING id
        """
        result = self.execute_query(query, (employee.name, employee.position, 
                                           employee.salary), fetch=True)
        if result:
            employee.id = result[0][0]
            return employee.id
        return None
    
    def update_employee(self, employee):
        query = """
            UPDATE employees 
            SET name = %s, position = %s, salary = %s
            WHERE id = %s
        """
        self.execute_query(query, (employee.name, employee.position, 
                                  employee.salary, employee.id))
    
    def delete_employee(self, emp_id):
        query = "DELETE FROM employees WHERE id = %s"
        self.execute_query(query, (emp_id,))
    
    # Методы для проектов
    def get_all_projects(self):
        query = "SELECT id, title FROM projects ORDER BY id"
        rows = self.execute_query(query, fetch=True)
        projects = []
        for row in rows:
            project = Project(row[1], project_id=row[0])
            tasks_query = """
                SELECT id, title, description, status, hours_required, employee_id
                FROM tasks WHERE project_id = %s
            """
            task_rows = self.execute_query(tasks_query, (row[0],), fetch=True)
            for task_row in task_rows:
                task = Task(
                    task_row[1], task_row[2], task_row[3],
                    hours_required=task_row[4], task_id=task_row[0]
                )
                if task_row[5]:
                    task.assigned_employee = Employee("", "", 0, 0, task_row[5])
                project.add_task(task)
            projects.append(project)
        return projects
    
    def add_project(self, project):
        query = "INSERT INTO projects (title) VALUES (%s) RETURNING id"
        result = self.execute_query(query, (project.title,), fetch=True)
        if result:
            project.id = result[0][0]
            return project.id
        return None
    
    def update_project(self, project):
        query = "UPDATE projects SET title = %s WHERE id = %s"
        self.execute_query(query, (project.title, project.id))
    
    def delete_project(self, project_id):
        query = "DELETE FROM projects WHERE id = %s"
        self.execute_query(query, (project_id,))
    
    # Методы для задач
    def get_all_tasks(self):
        query = """
            SELECT t.id, t.title, t.description, t.status, t.hours_required, 
                   t.employee_id, t.project_id, e.name as employee_name,
                   p.title as project_title
            FROM tasks t
            LEFT JOIN employees e ON t.employee_id = e.id
            LEFT JOIN projects p ON t.project_id = p.id
            ORDER BY t.id
        """
        rows = self.execute_query(query, fetch=True)
        tasks = []
        for row in rows:
            task = Task(
                row[1], row[2], row[3],
                hours_required=row[4], task_id=row[0]
            )
            if row[5]:
                task.assigned_employee = Employee(row[7] or "", "", 0, 0, row[5])
            task.project_id = row[6]
            tasks.append(task)
        return tasks
    
    def add_task(self, task):
        query = """
            INSERT INTO tasks (title, description, status, hours_required, 
                              employee_id, project_id)
            VALUES (%s, %s, %s, %s, %s, %s) RETURNING id
        """
        emp_id = task.assigned_employee.id if task.assigned_employee else None
        result = self.execute_query(query, (
            task.title, task.description, task.status,
            task.hours_required, emp_id, task.project_id
        ), fetch=True)
        if result:
            task.id = result[0][0]
            return task.id
        return None
    
    def update_task(self, task):
        query = """
            UPDATE tasks 
            SET title = %s, description = %s, status = %s, 
                hours_required = %s, employee_id = %s, project_id = %s
            WHERE id = %s
        """
        emp_id = task.assigned_employee.id if task.assigned_employee else None
        self.execute_query(query, (
            task.title, task.description, task.status,
            task.hours_required, emp_id, task.project_id, task.id
        ))
    
    def delete_task(self, task_id):
        query = "DELETE FROM tasks WHERE id = %s"
        self.execute_query(query, (task_id,))
    
    def mark_task_complete(self, task_id):
        """Отмечает задачу как завершенную и обновляет часы сотрудника"""
        query = "SELECT employee_id, hours_required FROM tasks WHERE id = %s"
        result = self.execute_query(query, (task_id,), fetch=True)
        
        if result and result[0][0]:
            employee_id = result[0][0]
            hours = result[0][1] or 0
            
            update_query = "UPDATE tasks SET status = 'Завершено' WHERE id = %s"
            self.execute_query(update_query, (task_id,))
            
            return employee_id, hours
        else:
            update_query = "UPDATE tasks SET status = 'Завершено' WHERE id = %s"
            self.execute_query(update_query, (task_id,))
            return None, 0
    
    def update_employee_hours(self, emp_id):
        """Обновляет поле hours_worked у сотрудника в БД"""
        hours = self.get_employee_hours_worked(emp_id)
        query = "UPDATE employees SET hours_worked = %s WHERE id = %s"
        self.execute_query(query, (hours, emp_id))
    
    def get_tasks_by_employee(self, emp_id, status=None):
        """Получает задачи сотрудника с возможностью фильтрации по статусу"""
        if status:
            query = """
                SELECT id, title, status, hours_required 
                FROM tasks 
                WHERE employee_id = %s AND status = %s
            """
            result = self.execute_query(query, (emp_id, status), fetch=True)
        else:
            query = """
                SELECT id, title, status, hours_required 
                FROM tasks 
                WHERE employee_id = %s
            """
            result = self.execute_query(query, (emp_id,), fetch=True)
        return result if result else []

class TimeTrackingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учет рабочего времени и задач")
        self.root.geometry("1200x700")        
        self.db = Database()       
        self.setup_ui()
        self.load_data()
    
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.employees_frame = ttk.Frame(self.notebook)
        self.tasks_frame = ttk.Frame(self.notebook)
        self.projects_frame = ttk.Frame(self.notebook)
        self.data_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.employees_frame, text='Сотрудники')
        self.notebook.add(self.tasks_frame, text='Задачи')
        self.notebook.add(self.projects_frame, text='Проекты')
        self.notebook.add(self.data_frame, text='Работа с данными')
        
        self.setup_employees_tab()
        self.setup_tasks_tab()
        self.setup_projects_tab()
        self.setup_data_tab()
    
    def setup_employees_tab(self):
        # Верхняя панель с кнопками для сотрубдников
        button_frame = ttk.Frame(self.employees_frame)
        button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(button_frame, text="Добавить", 
                  command=self.add_employee_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Редактировать", 
                  command=self.edit_employee_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Удалить", 
                  command=self.delete_employee).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Обновить часы", 
                  command=self.update_employee_hours).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Показать задачи", 
                  command=self.show_employee_tasks).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Экспорт в CSV", 
                  command=self.export_employees_csv).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Обновить", 
                  command=self.load_employees).pack(side='left', padx=5)
        
        # Таблица сотрудников
        columns = ('ID', 'Имя', 'Должность', 'Зарплата', 'Отработано часов', 'Заработок', 'Завершено задач')
        self.employees_tree = ttk.Treeview(self.employees_frame, columns=columns, show='headings')
        
        for col in columns:
            self.employees_tree.heading(col, text=col)
            self.employees_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(self.employees_frame, orient="vertical", 
                                 command=self.employees_tree.yview)
        self.employees_tree.configure(yscrollcommand=scrollbar.set)
        
        self.employees_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y', padx=5, pady=5)
    
    def setup_tasks_tab(self):
        # Верхняя панель с кнопками для задач
        button_frame = ttk.Frame(self.tasks_frame)
        button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(button_frame, text="Добавить", 
                  command=self.add_task_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Редактировать", 
                  command=self.edit_task_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Удалить", 
                  command=self.delete_task).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Выполнено", 
                  command=self.mark_task_complete).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Экспорт в CSV", 
                  command=self.export_tasks_csv).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Обновить", 
                  command=self.load_tasks).pack(side='left', padx=5)
        
        # Таблица задач
        columns = ('ID', 'Название', 'Статус', 'Часы', 'Сотрудник', 'Проект')
        self.tasks_tree = ttk.Treeview(self.tasks_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tasks_tree.heading(col, text=col)
            self.tasks_tree.column(col, width=150)
        
        scrollbar = ttk.Scrollbar(self.tasks_frame, orient="vertical", 
                                 command=self.tasks_tree.yview)
        self.tasks_tree.configure(yscrollcommand=scrollbar.set)
        
        self.tasks_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y', padx=5, pady=5)
    
    def setup_projects_tab(self):
        # Верхняя панель с кнопкамидля проектов
        button_frame = ttk.Frame(self.projects_frame)
        button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(button_frame, text="Добавить", 
                  command=self.add_project_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Редактировать", 
                  command=self.edit_project_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Удалить", 
                  command=self.delete_project).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Экспорт в CSV", 
                  command=self.export_projects_csv).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Обновить", 
                  command=self.load_projects).pack(side='left', padx=5)
        
        # Таблица проектов
        columns = ('ID', 'Название', 'Всего задач', 'Завершено', 'Прогресс', 'Всего часов')
        self.projects_tree = ttk.Treeview(self.projects_frame, columns=columns, show='headings')
        
        for col in columns:
            self.projects_tree.heading(col, text=col)
            self.projects_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(self.projects_frame, orient="vertical", 
                                 command=self.projects_tree.yview)
        self.projects_tree.configure(yscrollcommand=scrollbar.set)
        
        self.projects_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar.pack(side='right', fill='y', padx=5, pady=5)
    
    def setup_data_tab(self):
        main_frame = ttk.Frame(self.data_frame)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Фрейм для поиска email
        email_frame = ttk.LabelFrame(main_frame, text="Извлечение email-адресов")
        email_frame.pack(fill='x', pady=10)
        
        # Панель с кнопками для email
        email_buttons_frame = ttk.Frame(email_frame)
        email_buttons_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(email_frame, text="Введите текст:").pack(anchor='w', padx=5, pady=5)
        self.email_text = tk.Text(email_frame, height=5, width=80)
        self.email_text.pack(padx=5, pady=5)
        
        ttk.Button(email_buttons_frame, text="Извлечь email", 
                  command=self.extract_emails).pack(side='left', padx=5)
        ttk.Button(email_buttons_frame, text="Очистить", 
                  command=self.clear_email_fields).pack(side='left', padx=5)
        
        ttk.Label(email_frame, text="Результат:").pack(anchor='w', padx=5, pady=(10, 0))
        self.email_result = tk.Text(email_frame, height=5, width=80, state='disabled')
        self.email_result.pack(padx=5, pady=5)
        
        # Фрейм для работы с CSV
        csv_frame = ttk.LabelFrame(main_frame, text="Работа с CSV файлами")
        csv_frame.pack(fill='x', pady=10)
        
        # Панель для выбора файла
        file_selection_frame = ttk.Frame(csv_frame)
        file_selection_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(file_selection_frame, text="Путь к CSV файлу:").pack(side='left', padx=5)
        self.csv_path = ttk.Entry(file_selection_frame, width=60)
        self.csv_path.pack(side='left', padx=5, expand=True, fill='x')
        
        ttk.Button(file_selection_frame, text="Открыть файл", 
                  command=self.browse_csv_file).pack(side='left', padx=5)
        
        # Панель с кнопками для CSV
        csv_buttons_frame = ttk.Frame(csv_frame)
        csv_buttons_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(csv_buttons_frame, text="Загрузить и сохранить CSV", 
                  command=self.load_and_save_csv).pack(side='left', padx=5)
        ttk.Button(csv_buttons_frame, text="Очистить", 
                  command=self.clear_csv_fields).pack(side='left', padx=5)
        
        ttk.Label(csv_frame, text="Результат обработки:").pack(anchor='w', padx=5, pady=(10, 0))
        self.csv_result = tk.Text(csv_frame, height=10, width=80, state='disabled')
        self.csv_result.pack(padx=5, pady=5)
    
    def clear_email_fields(self):
        self.email_text.delete('1.0', 'end')
        self.email_result.config(state='normal')
        self.email_result.delete('1.0', 'end')
        self.email_result.config(state='disabled')
    
    def browse_csv_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filepath:
            self.csv_path.delete(0, 'end')
            self.csv_path.insert(0, filepath)
    
    def clear_csv_fields(self):
        self.csv_path.delete(0, 'end')
        self.csv_result.config(state='normal')
        self.csv_result.delete('1.0', 'end')
        self.csv_result.config(state='disabled')
    
    def load_data(self):
        self.load_employees()
        self.load_tasks()
        self.load_projects()
    
    def load_employees(self):
        for item in self.employees_tree.get_children():
            self.employees_tree.delete(item)
        
        employees = self.db.get_all_employees()
        for emp in employees:
            pay = emp.calculate_pay()
            
            completed_tasks = self.db.get_tasks_by_employee(emp.id, "Завершено")
            completed_count = len(completed_tasks) if completed_tasks else 0
            
            self.employees_tree.insert('', 'end', values=(
                emp.id, emp.name, emp.position, 
                f"{emp.salary:.2f}", f"{emp.hours_worked:.1f}",
                f"{pay:.2f}", completed_count
            ))
    
    def load_tasks(self):
        for item in self.tasks_tree.get_children():
            self.tasks_tree.delete(item)
        
        tasks = self.db.get_all_tasks()
        for task in tasks:
            emp_name = task.assigned_employee.name if task.assigned_employee else "Не назначен"
            project_title = self.get_project_title(task.project_id) if task.project_id else "Не назначен"
            
            self.tasks_tree.insert('', 'end', values=(
                task.id, task.title, task.status, 
                f"{task.hours_required:.1f}", emp_name, project_title
            ))
    
    def load_projects(self):
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        
        projects = self.db.get_all_projects()
        for project in projects:
            project_dict = project.to_dict()
            
            total_hours = sum(task.hours_required for task in project.tasks)
            
            self.projects_tree.insert('', 'end', values=(
                project.id, project.title, 
                project_dict['total_tasks'], project_dict['completed_tasks'],
                project_dict['progress'], f"{total_hours:.1f}"
            ))
    
    def get_project_title(self, project_id):
        if not project_id:
            return "Не назначен"
        query = "SELECT title FROM projects WHERE id = %s"
        result = self.db.execute_query(query, (project_id,), fetch=True)
        return result[0][0] if result else "Неизвестно"
    
    def add_employee_dialog(self):
        self.employee_dialog("Добавить сотрудника", None)
    
    def edit_employee_dialog(self):
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите сотрудника для редактирования")
            return
        
        item = self.employees_tree.item(selection[0])
        emp_id = item['values'][0]       
        employee = self.db.get_employee_by_id(emp_id)
        if employee:
            self.employee_dialog("Редактировать сотрудника", employee)
    
    def employee_dialog(self, title, employee):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x300")
        
        current_hours = self.db.get_employee_hours_worked(employee.id) if employee else 0
        
        ttk.Label(dialog, text="Имя:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        name_var = tk.StringVar(value=employee.name if employee else "")
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Должность:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        position_var = tk.StringVar(value=employee.position if employee else "")
        position_entry = ttk.Entry(dialog, textvariable=position_var, width=30)
        position_entry.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Зарплата:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        salary_var = tk.StringVar(value=str(employee.salary) if employee else "0")
        salary_entry = ttk.Entry(dialog, textvariable=salary_var, width=30)
        salary_entry.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text=f"Отработано часов (авто):").grid(row=3, column=0, padx=10, pady=10, sticky='w')
        ttk.Label(dialog, text=f"{current_hours:.1f} ч").grid(row=3, column=1, padx=10, pady=10, sticky='w')
        
        def save_employee():
            try:
                name = name_var.get()
                position = position_var.get()
                salary = float(salary_var.get())
                
                if not name:
                    messagebox.showerror("Ошибка", "Введите имя сотрудника")
                    return
                
                if employee:
                    employee.name = name
                    employee.position = position
                    employee.salary = salary
                    self.db.update_employee(employee)
                else:
                    new_employee = Employee(name, position, salary, 0)
                    self.db.add_employee(new_employee)
                
                self.load_employees()
                dialog.destroy()
                messagebox.showinfo("Успех", "Сотрудник сохранен")
                
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректные числовые значения")
        
        ttk.Button(dialog, text="Сохранить", command=save_employee).grid(row=4, column=0, columnspan=2, pady=20)
        ttk.Button(dialog, text="Отмена", command=dialog.destroy).grid(row=5, column=0, columnspan=2)
    
    def delete_employee(self):
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите сотрудника для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранного сотрудника?"):
            item = self.employees_tree.item(selection[0])
            emp_id = item['values'][0]
            self.db.delete_employee(emp_id)
            self.load_employees()
            self.load_tasks()
    
    def update_employee_hours(self):
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите сотрудника")
            return
        
        item = self.employees_tree.item(selection[0])
        emp_id = item['values'][0]       
        self.db.update_employee_hours(emp_id)
        self.load_employees()
        messagebox.showinfo("Обновлено", "Часы сотрудника обновлены")
    
    def show_employee_tasks(self):
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите сотрудника")
            return
        
        item = self.employees_tree.item(selection[0])
        emp_id = item['values'][0]
        emp_name = item['values'][1]
        
        tasks = self.db.get_tasks_by_employee(emp_id)
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Задачи сотрудника: {emp_name}")
        dialog.geometry("600x400")
        
        columns = ('ID', 'Название', 'Статус', 'Часы', 'Проект')
        tree = ttk.Treeview(dialog, columns=columns, show='headings')
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)
        
        completed_hours = 0
        for task in tasks:
            task_id, title, status, hours = task
            project_title = self.get_task_project_title(task_id)
            
            tree.insert('', 'end', values=(
                task_id, title, status, f"{hours:.1f}", project_title
            ))
            
            if status == "Завершено":
                completed_hours += hours
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        stats_frame = ttk.Frame(dialog)
        stats_frame.pack(fill='x', padx=10, pady=5)
        
        total_tasks = len(tasks)
        completed_tasks = sum(1 for t in tasks if t[2] == "Завершено")
        
        ttk.Label(stats_frame, text=f"Всего задач: {total_tasks}").pack(side='left', padx=10)
        ttk.Label(stats_frame, text=f"Завершено: {completed_tasks}").pack(side='left', padx=10)
        ttk.Label(stats_frame, text=f"Отработано часов: {completed_hours:.1f}").pack(side='left', padx=10)
    
    def get_task_project_title(self, task_id):
        query = """
            SELECT p.title 
            FROM tasks t 
            JOIN projects p ON t.project_id = p.id 
            WHERE t.id = %s
        """
        result = self.db.execute_query(query, (task_id,), fetch=True)
        return result[0][0] if result else "Не назначен"
    
    def add_task_dialog(self):
        self.task_dialog("Добавить задачу", None)
    
    def edit_task_dialog(self):
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите задачу для редактирования")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        
        query = """
            SELECT id, title, description, status, hours_required, employee_id, project_id
            FROM tasks WHERE id = %s
        """
        result = self.db.execute_query(query, (task_id,), fetch=True)
        if result:
            task_data = result[0]
            task = Task(
                task_data[1], task_data[2], task_data[3],
                hours_required=task_data[4], task_id=task_data[0]
            )
            if task_data[5]:
                task.assigned_employee = Employee("", "", 0, 0, task_data[5])
            task.project_id = task_data[6]
            self.task_dialog("Редактировать задачу", task)
    
    def task_dialog(self, title, task):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("500x400")
        
        employees = self.db.get_all_employees()
        projects = self.db.get_all_projects()
        
        ttk.Label(dialog, text="Название:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        title_var = tk.StringVar(value=task.title if task else "")
        title_entry = ttk.Entry(dialog, textvariable=title_var, width=40)
        title_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Описание:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        description_text = tk.Text(dialog, height=5, width=40)
        description_text.grid(row=1, column=1, padx=10, pady=10)
        if task and task.description:
            description_text.insert('1.0', task.description)
        
        ttk.Label(dialog, text="Статус:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        status_var = tk.StringVar(value=task.status if task else "В процессе")
        status_combo = ttk.Combobox(dialog, textvariable=status_var, 
                                   values=["В процессе", "Завершено"], state="readonly")
        status_combo.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Требуется часов:").grid(row=3, column=0, padx=10, pady=10, sticky='w')
        hours_var = tk.StringVar(value=str(task.hours_required) if task else "0")
        hours_entry = ttk.Entry(dialog, textvariable=hours_var, width=40)
        hours_entry.grid(row=3, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Сотрудник:").grid(row=4, column=0, padx=10, pady=10, sticky='w')
        employee_var = tk.StringVar()
        employee_combo = ttk.Combobox(dialog, textvariable=employee_var, state="readonly")
        employee_combo['values'] = [f"{e.id}: {e.name}" for e in employees]
        employee_combo.grid(row=4, column=1, padx=10, pady=10)
        
        if task and task.assigned_employee:
            employee_combo.set(f"{task.assigned_employee.id}: {task.assigned_employee.name}")
        
        ttk.Label(dialog, text="Проект:").grid(row=5, column=0, padx=10, pady=10, sticky='w')
        project_var = tk.StringVar()
        project_combo = ttk.Combobox(dialog, textvariable=project_var, state="readonly")
        project_combo['values'] = [f"{p.id}: {p.title}" for p in projects]
        project_combo.grid(row=5, column=1, padx=10, pady=10)
        
        if task and task.project_id:
            for p in projects:
                if p.id == task.project_id:
                    project_combo.set(f"{p.id}: {p.title}")
                    break
        
        def save_task():
            try:
                title_val = title_var.get()
                description = description_text.get('1.0', 'end-1c')
                status = status_var.get()
                hours = float(hours_var.get())
                
                emp_id = None
                emp_obj = None
                if employee_var.get():
                    emp_id = int(employee_var.get().split(':')[0])
                    emp_obj = next((e for e in employees if e.id == emp_id), None)
                
                proj_id = None
                if project_var.get():
                    proj_id = int(project_var.get().split(':')[0])
                
                if task:
                    task.title = title_val
                    task.description = description
                    task.status = status
                    task.hours_required = hours
                    if emp_obj:
                        task.assigned_employee = emp_obj
                    task.project_id = proj_id
                    self.db.update_task(task)
                else:
                    new_task = Task(title_val, description, status, 
                                   hours_required=hours)
                    if emp_obj:
                        new_task.assigned_employee = emp_obj
                    new_task.project_id = proj_id
                    self.db.add_task(new_task)
                
                self.load_tasks()
                self.load_projects()
                self.load_employees()
                dialog.destroy()
                messagebox.showinfo("Успех", "Задача сохранена")
                
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректные числовые значения")
        
        ttk.Button(dialog, text="Сохранить", command=save_task).grid(row=6, column=0, columnspan=2, pady=20)
        ttk.Button(dialog, text="Отмена", command=dialog.destroy).grid(row=7, column=0, columnspan=2)
    
    def delete_task(self):
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите задачу для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранную задачу?"):
            item = self.tasks_tree.item(selection[0])
            task_id = item['values'][0]
            
            query = "SELECT employee_id FROM tasks WHERE id = %s"
            result = self.db.execute_query(query, (task_id,), fetch=True)
            
            self.db.delete_task(task_id)
            self.load_tasks()
            self.load_projects()
            self.load_employees()
            messagebox.showinfo("Удалено", "Задача удалена")
    
    def mark_task_complete(self):
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите задачу для отметки как выполненную")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        
        employee_id, hours = self.db.mark_task_complete(task_id)
        
        if employee_id:
            self.db.update_employee_hours(employee_id)
            messagebox.showinfo("Выполнено", f"Задача отмечена как выполненная. Сотруднику добавлено {hours} часов.")
        else:
            messagebox.showinfo("Выполнено", "Задача отмечена как выполненная.")
        
        self.load_tasks()
        self.load_projects()
        self.load_employees()
    
    def add_project_dialog(self):
        self.project_dialog("Добавить проект", None)
    
    def edit_project_dialog(self):
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите проект для редактирования")
            return
        
        item = self.projects_tree.item(selection[0])
        project_id = item['values'][0]
        
        query = "SELECT id, title FROM projects WHERE id = %s"
        result = self.db.execute_query(query, (project_id,), fetch=True)
        if result:
            project_data = result[0]
            project = Project(project_data[1], project_id=project_data[0])
            self.project_dialog("Редактировать проект", project)
    
    def project_dialog(self, title, project):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x200")
        
        ttk.Label(dialog, text="Название проекта:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        title_var = tk.StringVar(value=project.title if project else "")
        title_entry = ttk.Entry(dialog, textvariable=title_var, width=30)
        title_entry.grid(row=0, column=1, padx=10, pady=10)
        
        def save_project():
            title_val = title_var.get()
            
            if not title_val:
                messagebox.showerror("Ошибка", "Введите название проекта")
                return
            
            if project:
                project.title = title_val
                self.db.update_project(project)
            else:
                new_project = Project(title_val)
                self.db.add_project(new_project)
            
            self.load_projects()
            self.load_tasks()
            dialog.destroy()
            messagebox.showinfo("Успех", "Проект сохранен")
        
        ttk.Button(dialog, text="Сохранить", command=save_project).grid(row=1, column=0, columnspan=2, pady=20)
        ttk.Button(dialog, text="Отмена", command=dialog.destroy).grid(row=2, column=0, columnspan=2)
    
    def delete_project(self):
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите проект для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранный проект и все его задачи?"):
            item = self.projects_tree.item(selection[0])
            project_id = item['values'][0]
            self.db.delete_project(project_id)
            self.load_projects()
            self.load_tasks()
            self.load_employees()
            messagebox.showinfo("Удалено", "Проект и все его задачи удалены")
    
    def export_employees_csv(self):
        employees = self.db.get_all_employees()
        data = []
        for emp in employees:
            completed_tasks = self.db.get_tasks_by_employee(emp.id, "Завершено")
            completed_count = len(completed_tasks) if completed_tasks else 0
            
            data.append({
                'ID': emp.id,
                'Имя': emp.name,
                'Должность': emp.position,
                'Зарплата': emp.salary,
                'Отработано часов': emp.hours_worked,
                'Заработок': emp.calculate_pay(),
                'Завершено задач': completed_count
            })
        
        df = pd.DataFrame(data)
        filename = f"employees_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(filename, index=False, encoding='utf-8')
        messagebox.showinfo("Экспорт", f"Данные экспортированы в {filename}")
    
    def export_tasks_csv(self):
        tasks = self.db.get_all_tasks()
        data = []
        for task in tasks:
            emp_name = task.assigned_employee.name if task.assigned_employee else "Не назначен"
            project_title = self.get_project_title(task.project_id)
            
            data.append({
                'ID': task.id,
                'Название': task.title,
                'Описание': task.description,
                'Статус': task.status,
                'Требуется часов': task.hours_required,
                'Сотрудник': emp_name,
                'Проект': project_title
            })
        
        df = pd.DataFrame(data)
        filename = f"tasks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(filename, index=False, encoding='utf-8')
        messagebox.showinfo("Экспорт", f"Данные экспортированы в {filename}")
    
    def export_projects_csv(self):
        projects = self.db.get_all_projects()
        data = []
        for project in projects:
            project_dict = project.to_dict()
            
            total_hours = sum(task.hours_required for task in project.tasks)
            completed_hours = sum(task.hours_required for task in project.tasks if task.status == "Завершено")
            
            data.append({
                'ID': project.id,
                'Название': project.title,
                'Всего задач': project_dict['total_tasks'],
                'Завершено задач': project_dict['completed_tasks'],
                'Прогресс': project_dict['progress'],
                'Всего часов': total_hours,
                'Выполнено часов': completed_hours
            })
        
        df = pd.DataFrame(data)
        filename = f"projects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(filename, index=False, encoding='utf-8')
        messagebox.showinfo("Экспорт", f"Данные экспортированы в {filename}")
    
    # Методы для работы с данными
    def extract_emails(self):
        text = self.email_text.get('1.0', 'end-1c')
        emails = extract_emails(text)
        
        self.email_result.config(state='normal')
        self.email_result.delete('1.0', 'end')
        
        if emails:
            self.email_result.insert('1.0', 'Найденные email-адреса:\n\n')
            for email in emails:
                self.email_result.insert('end', f'{email}\n')
        else:
            self.email_result.insert('1.0', 'Email-адреса не найдены')
        
        self.email_result.config(state='disabled')
    
    def load_and_save_csv(self):
        """Загружает CSV файл, удаляет строки с пропущенными(пустыми) значениями и сохраняет в ту же директорию"""
        filepath = self.csv_path.get()
        if not filepath:
            messagebox.showwarning("Предупреждение", "Выберите CSV файл")
            return
        
        try:
            df_original = pd.read_csv(filepath)
            
            df_cleaned = clean_csv_data(df_original)
            
            file_dir = os.path.dirname(filepath)
            file_name = os.path.basename(filepath)
            name_without_ext = os.path.splitext(file_name)[0]
            cleaned_file_name = f"{name_without_ext}_cleaned.csv"
            cleaned_file_path = os.path.join(file_dir, cleaned_file_name)
            
            df_cleaned.to_csv(cleaned_file_path, index=False, encoding='utf-8')
            
            self.csv_result.config(state='normal')
            self.csv_result.delete('1.0', 'end')
            
            self.csv_result.insert('1.0', f"Исходный файл: {filepath}\n")
            self.csv_result.insert('end', f"Очищенный файл: {cleaned_file_path}\n\n")
            self.csv_result.insert('end', f"Записей в исходном файле: {len(df_original)}\n")
            self.csv_result.insert('end', f"Записей после удаления пустых значений: {len(df_cleaned)}\n")
            self.csv_result.insert('end', f"Удалено записей: {len(df_original) - len(df_cleaned)}\n\n")
            
            df_for_stats = df_original.replace(r'^\s*$', np.nan, regex=True)
            empty_counts = df_for_stats.isna().sum()
            
            if empty_counts.sum() > 0:
                self.csv_result.insert('end', "Статистика по пропущенным значениям:\n")
                for column, count in empty_counts.items():
                    if count > 0:
                        self.csv_result.insert('end', f"  {column}: {count} пропущенных значений\n")
            
            self.csv_result.insert('end', f"\nКолонки: {', '.join(df_cleaned.columns)}\n\n")
            self.csv_result.insert('end', "Первые 10 строк очищенного файла:\n")
            self.csv_result.insert('end', df_cleaned.head(10).to_string())
            
            self.csv_result.config(state='disabled')
            
            messagebox.showinfo(
                "Успех", 
                f"Файл успешно обработан!\n\n"
                f"Исходный файл: {len(df_original)} записей\n"
                f"Очищенный файл: {len(df_cleaned)} записей\n"
                f"Удалено записей с пустыми значениями: {len(df_original) - len(df_cleaned)}\n\n"
                f"Сохранен как: {cleaned_file_path}"
            )
            
        except FileNotFoundError:
            messagebox.showerror("Ошибка", f"Файл не найден: {filepath}")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Ошибка", "Файл пуст или имеет неверный формат")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать CSV файл: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeTrackingApp(root)
    root.mainloop()