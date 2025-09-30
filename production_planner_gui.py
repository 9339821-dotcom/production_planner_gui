import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import json
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from typing import Dict, List, Tuple
import warnings
warnings.filterwarnings('ignore')

class AdvancedProductionPlanner:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.orders_df = None
        self.materials_df = None
        self.machines_df = None
        self.operations_df = None
        self.stock_data = {}
        self.reserved_materials = defaultdict(float)
        self.reserved_orders = set()
        self.machine_capacity = {}
        self.operation_times = {}
        self.load_all_data()
    
    def load_all_data(self):
        """Загрузка всех данных из Excel файла"""
        try:
            print("📂 Загрузка данных из Excel файла...")
            
            # Загружаем все листы
            self.orders_df = pd.read_excel(self.excel_file, sheet_name='Заказы')
            self.materials_df = pd.read_excel(self.excel_file, sheet_name='Потребность материалов')
            
            # Пытаемся загрузить дополнительные данные если они есть
            try:
                self.machines_df = pd.read_excel(self.excel_file, sheet_name='Оборудование')
                self.load_machine_data()
            except:
                print("⚠️ Лист 'Оборудование' не найден, используем стандартные настройки")
                self.set_default_machine_capacity()
            
            try:
                self.operations_df = pd.read_excel(self.excel_file, sheet_name='Операции')
                self.load_operation_data()
            except:
                print("⚠️ Лист 'Операции' не найден, используем стандартные настройки")
                self.set_default_operation_times()
            
            # Создаем словарь остатков на складе
            if 'На складе' in self.materials_df.columns:
                for _, row in self.materials_df.iterrows():
                    material = row['Материал']
                    if pd.notna(material):
                        stock = row['На складе'] if pd.notna(row['На складе']) else 0
                        self.stock_data[str(material).strip()] = float(stock)
            
            print(f"✅ Загружено: {len(self.orders_df)} заказов, {len(self.materials_df)} материалов")
            
        except Exception as e:
            print(f"❌ Ошибка загрузки данных: {e}")
            raise
    
    def set_default_machine_capacity(self):
        """Установка стандартной производительности оборудования"""
        self.machine_capacity = {
            'Резка': 8,  # часов в день
            'Сварка': 10,
            'Сборка': 12,
            'Покраска': 8,
            'Упаковка': 10
        }
    
    def set_default_operation_times(self):
        """Установка стандартного времени операций"""
        # Время в часах на м² для разных типов продукции
        self.operation_times = {
            'Окно': {'Резка': 0.5, 'Сварка': 0.8, 'Сборка': 1.2, 'Покраска': 0.3, 'Упаковка': 0.2},
            'Дверь': {'Резка': 0.7, 'Сварка': 1.0, 'Сборка': 1.5, 'Покраска': 0.4, 'Упаковка': 0.3},
            'Фасад': {'Резка': 0.6, 'Сварка': 0.9, 'Сборка': 1.3, 'Покраска': 0.5, 'Упаковка': 0.25}
        }
    
    def load_machine_data(self):
        """Загрузка данных об оборудовании"""
        if self.machines_df is not None:
            for _, row in self.machines_df.iterrows():
                machine = row['Оборудование']
                capacity = row['Производительность_час']
                self.machine_capacity[machine] = capacity
    
    def load_operation_data(self):
        """Загрузка данных об операциях"""
        if self.operations_df is not None:
            for _, row in self.operations_df.iterrows():
                product_type = row['Тип_продукции']
                operation = row['Операция']
                time_per_sqm = row['Время_на_м2']
                
                if product_type not in self.operation_times:
                    self.operation_times[product_type] = {}
                self.operation_times[product_type][operation] = time_per_sqm
    
    def get_companies(self):
        """Получить список компаний"""
        return sorted([str(x) for x in self.orders_df['Клиент'].unique() if pd.notna(x)])
    
    def get_product_types(self):
        """Получить список типов продукции"""
        return sorted([str(x) for x in self.orders_df['Тип продукции'].unique() if pd.notna(x)])
    
    def calculate_material_requirements(self, selected_order_numbers):
        """Рассчитать потребность в материалах для выбранных заказов"""
        if not selected_order_numbers:
            return {"error": "Не выбраны заказы"}
        
        required_materials = defaultdict(float)
        
        # Получаем все материалы из таблицы потребности
        all_materials = []
        if 'Материал' in self.materials_df.columns:
            all_materials = [str(x).strip() for x in self.materials_df['Материал'] if pd.notna(x)]
        
        # Для каждого выбранного заказа находим его материалы
        for order_num in selected_order_numbers:
            order_num_clean = str(order_num).strip()
            
            # Ищем колонку с этим заказом в таблице материалов
            order_columns = []
            for col in self.materials_df.columns:
                col_str = str(col).strip()
                if order_num_clean == col_str or order_num_clean in col_str.split():
                    order_columns.append(col)
            
            if order_columns:
                for material_idx, material_name in enumerate(all_materials):
                    total_requirement = 0
                    
                    for order_col in order_columns:
                        if order_col in self.materials_df.columns:
                            value = self.materials_df[order_col].iloc[material_idx]
                            if pd.notna(value) and value != 0:
                                try:
                                    total_requirement += float(value)
                                except (ValueError, TypeError):
                                    pass
                    
                    if total_requirement > 0:
                        required_materials[material_name] += total_requirement
        
        # Рассчитываем остатки после резервирования
        material_balance = {}
        purchase_requirements = {}
        urgent_purchase = {}
        
        for material, required in required_materials.items():
            current_stock = self.stock_data.get(material, 0)
            reserved = self.reserved_materials.get(material, 0)
            available_stock = max(0, current_stock - reserved)
            
            balance_after = available_stock - required
            material_balance[material] = {
                'Текущий запас': current_stock,
                'Уже зарезервировано': reserved,
                'Доступно сейчас': available_stock,
                'Требуется для выбранных': required,
                'Остаток после': balance_after
            }
            
            # Если будет дефицит - добавляем в заявку на закупку
            if balance_after < 0:
                purchase_requirements[material] = abs(balance_after)
                # Срочная закупка если дефицит более 50%
                if abs(balance_after) > current_stock * 0.5:
                    urgent_purchase[material] = abs(balance_after)
        
        return {
            'material_requirements': dict(required_materials),
            'material_balance': material_balance,
            'purchase_requirements': purchase_requirements,
            'urgent_purchase': urgent_purchase
        }
    
    def reserve_materials(self, order_numbers):
        """Резервирование материалов для заказов"""
        requirements = self.calculate_material_requirements(order_numbers)
        
        if 'error' in requirements:
            return requirements
        
        # Резервируем материалы
        for material, required in requirements['material_requirements'].items():
            self.reserved_materials[material] += required
        
        # Добавляем заказы в список зарезервированных
        self.reserved_orders.update(order_numbers)
        
        return requirements
    
    def release_materials(self, order_numbers):
        """Освобождение зарезервированных материалов"""
        requirements = self.calculate_material_requirements(order_numbers)
        
        if 'error' in requirements:
            return requirements
        
        # Освобождаем материалы
        for material, required in requirements['material_requirements'].items():
            if material in self.reserved_materials:
                self.reserved_materials[material] = max(0, self.reserved_materials[material] - required)
        
        # Удаляем заказы из списка зарезервированных
        self.reserved_orders.difference_update(order_numbers)
        
        return requirements
    
    def optimize_production_schedule(self, selected_orders, start_date=None):
        """Оптимизация расписания производства"""
        if start_date is None:
            start_date = datetime.now().date()
        
        # Фильтруем выбранные заказы
        selected_orders_data = self.orders_df[self.orders_df['Номер заказа'].isin(selected_orders)].copy()
        
        if selected_orders_data.empty:
            return {"error": "Не найдены данные по выбранным заказам"}
        
        # Расчет времени производства для каждого заказа
        production_times = []
        for _, order in selected_orders_data.iterrows():
            product_type = order.get('Тип продукции', 'Окно')
            area = order.get('Площадь заказа', 1)
            
            if product_type in self.operation_times:
                total_time = 0
                for operation, time_per_sqm in self.operation_times[product_type].items():
                    total_time += time_per_sqm * area
                production_times.append(total_time)
            else:
                # Стандартное время если тип продукции не найден
                production_times.append(area * 2)  # 2 часа на м²
        
        selected_orders_data['production_time_hours'] = production_times
        selected_orders_data['production_time_days'] = [t / 8 for t in production_times]  # 8 часов в день
        
        # Сортировка по приоритету (срочные сначала)
        selected_orders_data['priority'] = selected_orders_data.get('Срочность', 1)
        selected_orders_data = selected_orders_data.sort_values('priority', ascending=True)
        
        # Расчет расписания
        current_date = start_date
        schedule = []
        
        for _, order in selected_orders_data.iterrows():
            order_num = order['Номер заказа']
            days_needed = np.ceil(order['production_time_days'])
            
            schedule.append({
                'Номер заказа': order_num,
                'Клиент': order['Клиент'],
                'Тип продукции': order.get('Тип продукции', ''),
                'Площадь': order.get('Площадь заказа', 0),
                'Начало производства': current_date,
                'Окончание производства': current_date + timedelta(days=int(days_needed)),
                'Дней производства': int(days_needed),
                'Часов производства': order['production_time_hours']
            })
            
            current_date += timedelta(days=int(days_needed) + 1)  # +1 день на перенастройку
        
        return {
            'schedule': schedule,
            'total_orders': len(schedule),
            'total_days': (current_date - start_date).days,
            'total_hours': sum([x['Часов производства'] for x in schedule])
        }
    
    def calculate_machine_utilization(self, schedule):
        """Расчет загрузки оборудования"""
        machine_workload = {machine: 0 for machine in self.machine_capacity.keys()}
        
        for order in schedule:
            product_type = order['Тип продукции']
            area = order['Площадь']
            
            if product_type in self.operation_times:
                for operation, time_per_sqm in self.operation_times[product_type].items():
                    if operation in machine_workload:
                        machine_workload[operation] += time_per_sqm * area
        
        # Расчет процента загрузки
        utilization = {}
        for machine, workload in machine_workload.items():
            capacity = self.machine_capacity.get(machine, 8) * len(schedule)  # часов доступно
            utilization[machine] = {
                'workload_hours': workload,
                'capacity_hours': capacity,
                'utilization_percent': min(100, (workload / capacity * 100)) if capacity > 0 else 0
            }
        
        return utilization
    
    def group_orders_by_product_type(self, selected_orders):
        """Группировка заказов по типу продукции для минимизации переналадок"""
        selected_orders_data = self.orders_df[self.orders_df['Номер заказа'].isin(selected_orders)]
        
        grouped = selected_orders_data.groupby('Тип продукции').agg({
            'Номер заказа': 'count',
            'Площадь заказа': 'sum',
            'Стоимость заказа': 'sum'
        }).reset_index()
        
        grouped = grouped.rename(columns={
            'Номер заказа': 'Количество заказов',
            'Площадь заказа': 'Общая площадь',
            'Стоимость заказа': 'Общая стоимость'
        })
        
        return grouped.to_dict('records')
    
    def estimate_material_price(self, material):
        """Оценочная стоимость материала"""
        material_lower = material.lower()
        
        if any(x in material_lower for x in ['стекло', 'glass']):
            return 1500
        elif any(x in material_lower for x in ['профиль', 'profile']):
            return 800
        elif any(x in material_lower for x in ['аргон', 'argon']):
            return 200
        elif any(x in material_lower for x in ['герметик', 'sealant']):
            return 1500
        elif any(x in material_lower for x in ['лента', 'tape']):
            return 300
        else:
            return 1000

class AdvancedProductionPlannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🏭 Advanced Production Planner - AI Оптимизация")
        self.root.geometry("1600x1000")
        self.root.configure(bg='#f0f0f0')
        
        # Центрирование окна
        self.center_window()
        
        # Стили
        self.setup_styles()
        
        # Загрузка данных
        self.planner = None
        self.load_data()
        
        # Создание интерфейса
        self.setup_ui()
        
    def center_window(self):
        """Центрирование окна на экране"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_styles(self):
        """Настройка стилей для современного внешнего вида"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Кастомные стили
        style.configure('Modern.TFrame', background='#f8f9fa')
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), background='#f8f9fa')
        style.configure('Card.TFrame', background='white', relief='raised', borderwidth=1)
        
        # Стили для кнопок
        style.configure('Primary.TButton', background='#007bff', foreground='white', font=('Arial', 10))
        style.map('Primary.TButton', background=[('active', '#0056b3')])
        
        style.configure('Success.TButton', background='#28a745', foreground='white', font=('Arial', 10))
        style.map('Success.TButton', background=[('active', '#1e7e34')])
        
        style.configure('Warning.TButton', background='#ffc107', foreground='black', font=('Arial', 10))
        style.map('Warning.TButton', background=[('active', '#e0a800')])
        
        style.configure('Danger.TButton', background='#dc3545', foreground='white', font=('Arial', 10))
        style.map('Danger.TButton', background=[('active', '#c82333')])
    
    def load_data(self):
        """Загрузка данных из файла"""
        excel_file = "Объединенная_статистика_заказов.xlsx"
        if not os.path.exists(excel_file):
            messagebox.showerror("Ошибка", f"Файл {excel_file} не найден!\nПоместите файл в ту же папку, что и программу.")
            self.root.destroy()
            return
        
        try:
            self.planner = AdvancedProductionPlanner(excel_file)
            messagebox.showinfo("Успех", f"Данные успешно загружены!\nЗаказов: {len(self.planner.orders_df)}\nМатериалов: {len(self.planner.materials_df)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")
            self.root.destroy()
    
    def setup_ui(self):
        """Создание современного пользовательского интерфейса"""
        # Главный контейнер
        main_container = ttk.Frame(self.root, style='Modern.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Заголовок
        header_frame = ttk.Frame(main_container, style='Modern.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(header_frame, 
                               text="🏭 ADVANCED PRODUCTION PLANNER - AI ОПТИМИЗАЦИЯ", 
                               style='Title.TLabel')
        title_label.pack(pady=10)
        
        # Создание вкладок
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Вкладки
        self.setup_orders_tab(notebook)
        self.setup_planning_tab(notebook)
        self.setup_optimization_tab(notebook)
        self.setup_materials_tab(notebook)
        self.setup_dashboard_tab(notebook)
    
    def setup_orders_tab(self, notebook):
        """Вкладка с обзором заказов"""
        orders_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(orders_frame, text="📋 Заказы")
        
        # Фильтры
        filter_frame = ttk.Frame(orders_frame, style='Card.TFrame')
        filter_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Фильтр по компании
        ttk.Label(filter_frame, text="Компания:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.company_var = tk.StringVar()
        companies = ['Все компании'] + self.planner.get_companies()
        company_combo = ttk.Combobox(filter_frame, textvariable=self.company_var, values=companies, state='readonly', width=20)
        company_combo.set('Все компании')
        company_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        company_combo.bind('<<ComboboxSelected>>', self.filter_orders)
        
        # Фильтр по типу продукции
        ttk.Label(filter_frame, text="Тип продукции:", font=('Arial', 10)).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.product_type_var = tk.StringVar()
        product_types = ['Все типы'] + self.planner.get_product_types()
        product_combo = ttk.Combobox(filter_frame, textvariable=self.product_type_var, values=product_types, state='readonly', width=15)
        product_combo.set('Все типы')
        product_combo.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        product_combo.bind('<<ComboboxSelected>>', self.filter_orders)
        
        # Поиск
        ttk.Label(filter_frame, text="Поиск:", font=('Arial', 10)).grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=20)
        search_entry.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        search_entry.bind('<KeyRelease>', self.filter_orders)
        
        # Таблица заказов
        table_frame = ttk.Frame(orders_frame, style='Modern.TFrame')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создание Treeview с прокруткой
        columns = ('Выбор', 'Номер', 'Клиент', 'Тип продукции', 'Площадь', 'Стоимость', 'Состояние', 'Срочность')
        self.orders_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Настройка колонок
        column_widths = {'Выбор': 50, 'Номер': 100, 'Клиент': 150, 'Тип продукции': 120, 
                        'Площадь': 80, 'Стоимость': 100, 'Состояние': 100, 'Срочность': 80}
        
        for col in columns:
            self.orders_tree.heading(col, text=col)
            self.orders_tree.column(col, width=column_widths.get(col, 100))
        
        # Checkbox для выбора
        self.orders_tree.heading('Выбор', text='☑')
        
        # Прокрутка
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.orders_tree.yview)
        self.orders_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.orders_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Загрузка данных
        self.load_orders_data()
        
        # Кнопки действий
        button_frame = ttk.Frame(orders_frame, style='Modern.TFrame')
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="✅ Добавить выбранные в планирование", 
                  command=self.add_selected_orders, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🔄 Обновить данные", 
                  command=self.load_orders_data, style='Success.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="📊 Статистика заказов", 
                  command=self.show_orders_stats, style='Warning.TButton').pack(side=tk.LEFT, padx=5)
    
    def setup_planning_tab(self, notebook):
        """Вкладка планирования производства"""
        planning_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(planning_frame, text="📅 Планирование")
        
        # Левая панель - выбранные заказы
        left_frame = ttk.Frame(planning_frame, style='Modern.TFrame')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(left_frame, text="Выбранные заказы:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=5)
        
        # Список выбранных заказов
        self.selected_orders_listbox = tk.Listbox(left_frame, height=15, font=('Arial', 10), selectmode=tk.MULTIPLE)
        self.selected_orders_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Кнопки управления выбранными заказами
        order_buttons_frame = ttk.Frame(left_frame, style='Modern.TFrame')
        order_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(order_buttons_frame, text="🗑️ Удалить выбранный", 
                  command=self.remove_selected_order, style='Danger.TButton').pack(side=tk.LEFT, padx=2)
        
        ttk.Button(order_buttons_frame, text="🧹 Очистить все", 
                  command=self.clear_all_orders, style='Danger.TButton').pack(side=tk.LEFT, padx=2)
        
        # Правая панель - управление планированием
        right_frame = ttk.Frame(planning_frame, style='Modern.TFrame')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Дата начала производства
        ttk.Label(right_frame, text="Дата начала производства:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=5)
        
        date_frame = ttk.Frame(right_frame, style='Modern.TFrame')
        date_frame.pack(fill=tk.X, pady=5)
        
        self.production_start_date = DateEntry(date_frame, width=12, background='darkblue',
                                             foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                                             font=('Arial', 10))
        self.production_start_date.pack(side=tk.LEFT, padx=5)
        
        # Кнопки планирования
        planning_buttons_frame = ttk.Frame(right_frame, style='Modern.TFrame')
        planning_buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(planning_buttons_frame, text="🧮 Рассчитать потребности в материалах", 
                  command=self.calculate_requirements, style='Primary.TButton').pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📦 Зарезервировать материалы", 
                  command=self.reserve_materials, style='Success.TButton').pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📄 Сформировать заявку на закупку", 
                  command=self.generate_purchase_order, style='Primary.TButton').pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="🔄 Снять резерв с выбранных", 
                  command=self.release_materials, style='Warning.TButton').pack(fill=tk.X, pady=2)
        
        # Область результатов
        ttk.Label(right_frame, text="Результаты расчета:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=(20, 5))
        
        self.results_text = scrolledtext.ScrolledText(right_frame, height=15, font=('Arial', 9))
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def setup_optimization_tab(self, notebook):
        """Вкладка оптимизации производства"""
        optimization_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(optimization_frame, text="⚙️ Оптимизация")
        
        # Верхняя панель - кнопки оптимизации
        top_frame = ttk.Frame(optimization_frame, style='Modern.TFrame')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(top_frame, text="📊 Оптимизировать расписание производства", 
                  command=self.optimize_schedule, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="🔧 Анализ загрузки оборудования", 
                  command=self.analyze_machine_utilization, style='Success.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="📈 Группировка заказов по типам", 
                  command=self.group_orders_by_type, style='Warning.TButton').pack(side=tk.LEFT, padx=5)
        
        # Область результатов оптимизации
        results_frame = ttk.Frame(optimization_frame, style='Modern.TFrame')
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.optimization_text = scrolledtext.ScrolledText(results_frame, height=20, font=('Arial', 9))
        self.optimization_text.pack(fill=tk.BOTH, expand=True)
    
    def setup_materials_tab(self, notebook):
        """Вкладка анализа материалов"""
        materials_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(materials_frame, text="📦 Материалы")
        
        # Таблица материалов
        table_frame = ttk.Frame(materials_frame, style='Modern.TFrame')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ('Материал', 'На складе', 'Зарезервировано', 'Доступно', 'Статус', 'Рекомендация')
        self.materials_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.materials_tree.yview)
        self.materials_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.materials_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Загрузка данных о материалах
        self.load_materials_data()
        
        # Кнопки
        button_frame = ttk.Frame(materials_frame, style='Modern.TFrame')
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="🔄 Обновить данные", 
                  command=self.load_materials_data, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="📊 Анализ дефицита", 
                  command=self.analyze_material_deficit, style='Warning.TButton').pack(side=tk.LEFT, padx=5)
    
    def setup_dashboard_tab(self, notebook):
        """Вкладка с дашбордом"""
        dashboard_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(dashboard_frame, text="📊 Дашборд")
        
        # Статистика
        stats_frame = ttk.Frame(dashboard_frame, style='Card.TFrame')
        stats_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Здесь можно добавить виджеты статистики
        ttk.Label(stats_frame, text="Дашборд в разработке...", font=('Arial', 12)).pack(pady=20)
    
    def load_orders_data(self):
        """Загрузка данных в таблицу заказов"""
        # Очистка таблицы
        for item in self.orders_tree.get_children():
            self.orders_tree.delete(item)
        
        # Фильтрация данных
        filtered_orders = self.planner.orders_df
        
        company_filter = self.company_var.get()
        if company_filter != 'Все компании':
            filtered_orders = filtered_orders[filtered_orders['Клиент'] == company_filter]
        
        product_type_filter = self.product_type_var.get()
        if product_type_filter != 'Все типы':
            filtered_orders = filtered_orders[filtered_orders['Тип продукции'] == product_type_filter]
        
        search_text = self.search_var.get().lower()
        if search_text:
            filtered_orders = filtered_orders[
                filtered_orders.apply(lambda row: search_text in str(row).lower(), axis=1)
            ]
        
        # Заполнение таблицы
        for _, order in filtered_orders.iterrows():
            status = order.get('Состояние заказа', 'Новый')
            priority = order.get('Срочность', 'Обычный')
            
            self.orders_tree.insert('', tk.END, values=(
                "☐",  # Чекбокс
                order['Номер заказа'],
                order['Клиент'],
                order.get('Тип продукции', ''),
                f"{order.get('Площадь заказа', 0):.2f}",
                f"{order.get('Стоимость заказа', 0):,.2f}",
                status,
                priority
            ))
    
    def filter_orders(self, event=None):
        """Фильтрация заказов"""
        self.load_orders_data()
    
    def add_selected_orders(self):
        """Добавление выбранных заказов в список планирования"""
        selected_items = self.orders_tree.selection()
        if not selected_items:
            messagebox.showwarning("Внимание", "Выберите заказы из таблицы!")
            return
        
        added_count = 0
        for item in selected_items:
            values = self.orders_tree.item(item)['values']
            order_num = values[1]  # Номер заказа во втором столбце
            
            # Добавляем в список, если еще нет
            if order_num not in self.get_selected_orders():
                self.selected_orders_listbox.insert(tk.END, order_num)
                added_count += 1
        
        messagebox.showinfo("Успех", f"Добавлено {added_count} заказов в планирование")
    
    def get_selected_orders(self):
        """Получить список выбранных заказов"""
        return list(self.selected_orders_listbox.get(0, tk.END))
    
    def remove_selected_order(self):
        """Удаление выбранного заказа из списка"""
        selected_indices = self.selected_orders_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Внимание", "Выберите заказ для удаления!")
            return
        
        for index in reversed(selected_indices):
            self.selected_orders_listbox.delete(index)
    
    def clear_all_orders(self):
        """Очистка всех выбранных заказов"""
        self.selected_orders_listbox.delete(0, tk.END)
    
    def calculate_requirements(self):
        """Расчет потребностей в материалах"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы для планирования!")
            return
        
        requirements = self.planner.calculate_material_requirements(selected_orders)
        
        if 'error' in requirements:
            messagebox.showerror("Ошибка", requirements['error'])
            return
        
        # Вывод результатов
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"📊 РЕЗУЛЬТАТЫ ДЛЯ {len(selected_orders)} ЗАКАЗОВ:\n")
        self.results_text.insert(tk.END, "=" * 60 + "\n\n")
        
        # Показываем материалы с наибольшей потребностью
        sorted_materials = sorted(requirements['material_balance'].items(), 
                                 key=lambda x: x[1]['Требуется для выбранных'], reverse=True)
        
        for material, balance in sorted_materials[:15]:  # Показываем топ-15
            if balance['Требуется для выбранных'] > 0:
                self.results_text.insert(tk.END, f"📦 {material}:\n")
                self.results_text.insert(tk.END, f"   Текущий запас: {balance['Текущий запас']:.2f}\n")
                self.results_text.insert(tk.END, f"   Уже зарезервировано: {balance['Уже зарезервировано']:.2f}\n")
                self.results_text.insert(tk.END, f"   Доступно сейчас: {balance['Доступно сейчас']:.2f}\n")
                self.results_text.insert(tk.END, f"   Требуется: {balance['Требуется для выбранных']:.2f}\n")
                
                remaining = balance['Остаток после']
                if remaining >= 0:
                    self.results_text.insert(tk.END, f"   ✅ Остаток после выполнения: {remaining:.2f}\n")
                else:
                    self.results_text.insert(tk.END, f"   ❌ ДЕФИЦИТ: {-remaining:.2f}\n")
                self.results_text.insert(tk.END, "\n")
        
        # Показываем заявку на закупку
        if requirements['purchase_requirements']:
            self.results_text.insert(tk.END, f"🚨 ТРЕБУЕТСЯ ЗАКУПКА ({len(requirements['purchase_requirements'])} материалов)\n")
            self.results_text.insert(tk.END, "Срочная закупка рекомендуется для:\n")
            for material in requirements.get('urgent_purchase', {}):
                self.results_text.insert(tk.END, f"   ⚠️ {material}\n")
    
    def reserve_materials(self):
        """Резервирование материалов"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        result = self.planner.reserve_materials(selected_orders)
        
        if 'error' in result:
            messagebox.showerror("Ошибка", result['error'])
            return
        
        messagebox.showinfo("Успех", f"Материалы зарезервированы для {len(selected_orders)} заказов")
        self.load_materials_data()  # Обновляем данные о материалах
        
        # Показываем обновленные остатки
        self.calculate_requirements()
    
    def release_materials(self):
        """Снятие резерва материалов"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        result = self.planner.release_materials(selected_orders)
        
        if 'error' in result:
            messagebox.showerror("Ошибка", result['error'])
            return
        
        messagebox.showinfo("Успех", f"Резерв снят для {len(selected_orders)} заказов")
        self.load_materials_data()  # Обновляем данные о материалах
        self.calculate_requirements()
    
    def generate_purchase_order(self):
        """Формирование заявки на закупку"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        requirements = self.planner.calculate_material_requirements(selected_orders)
        
        if 'error' in requirements:
            messagebox.showerror("Ошибка", requirements['error'])
            return
        
        if not requirements.get('purchase_requirements'):
            messagebox.showinfo("Информация", "Заявка на закупку не требуется - все материалы в наличии")
            return
        
        # Создание заявки
        purchase_text = f"ЗАЯВКА НА ЗАКУПКУ МАТЕРИАЛОВ\n"
        purchase_text += f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        purchase_text += f"Для заказов: {', '.join(selected_orders)}\n"
        purchase_text += f"Количество заказов: {len(selected_orders)}\n"
        purchase_text += "=" * 50 + "\n\n"
        
        total_cost = 0
        urgent_materials = requirements.get('urgent_purchase', {})
        
        for material, quantity in requirements['purchase_requirements'].items():
            price = self.planner.estimate_material_price(material)
            cost = price * quantity
            total_cost += cost
            
            urgency = "⚡ СРОЧНО " if material in urgent_materials else ""
            purchase_text += f"• {material}: {quantity:.2f} × {price:,.2f} руб. = {cost:,.2f} руб. {urgency}\n"
        
        purchase_text += f"\nОБЩАЯ СТОИМОСТЬ: {total_cost:,.2f} руб.\n"
        purchase_text += f"Срочных материалов: {len(urgent_materials)}\n"
        purchase_text += f"Рекомендуемый срок поставки: {datetime.now().strftime('%d.%m.%Y')}"
        
        # Показ заявки
        self.show_purchase_order(purchase_text)
    
    def show_purchase_order(self, order_text):
        """Показ заявки на закупку в отдельном окне"""
        order_window = tk.Toplevel(self.root)
        order_window.title("Заявка на закупку материалов")
        order_window.geometry("700x600")
        order_window.configure(bg='white')
        
        # Центрирование
        order_window.transient(self.root)
        order_window.grab_set()
        
        # Текст заявки
        text_widget = scrolledtext.ScrolledText(order_window, font=('Arial', 10), wrap=tk.WORD)
        text_widget.insert(1.0, order_text)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        text_widget.config(state=tk.DISABLED)
        
        # Кнопки
        button_frame = ttk.Frame(order_window)
        button_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Button(button_frame, text="💾 Сохранить в файл", 
                  command=lambda: self.save_purchase_order(order_text), style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🖨️ Печать", 
                  command=lambda: self.print_purchase_order(order_text), style='Success.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="✖️ Закрыть", 
                  command=order_window.destroy, style='Danger.TButton').pack(side=tk.RIGHT, padx=5)
    
    def optimize_schedule(self):
        """Оптимизация расписания производства"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы для оптимизации!")
            return
        
        start_date = self.production_start_date.get_date()
        schedule = self.planner.optimize_production_schedule(selected_orders, start_date)
        
        if 'error' in schedule:
            messagebox.showerror("Ошибка", schedule['error'])
            return
        
        # Вывод результатов оптимизации
        self.optimization_text.delete(1.0, tk.END)
        self.optimization_text.insert(tk.END, "⚙️ ОПТИМИЗИРОВАННОЕ РАСПИСАНИЕ ПРОИЗВОДСТВА\n")
        self.optimization_text.insert(tk.END, "=" * 60 + "\n\n")
        
        self.optimization_text.insert(tk.END, f"📅 Дата начала: {start_date.strftime('%d.%m.%Y')}\n")
        self.optimization_text.insert(tk.END, f"📦 Всего заказов: {schedule['total_orders']}\n")
        self.optimization_text.insert(tk.END, f"⏱️ Общее время производства: {schedule['total_days']} дней\n")
        self.optimization_text.insert(tk.END, f"🕒 Общее количество часов: {schedule['total_hours']:.1f} ч\n\n")
        
        self.optimization_text.insert(tk.END, "ДЕТАЛЬНОЕ РАСПИСАНИЕ:\n")
        self.optimization_text.insert(tk.END, "-" * 60 + "\n")
        
        for i, order in enumerate(schedule['schedule'], 1):
            self.optimization_text.insert(tk.END, f"{i}. Заказ {order['Номер заказа']} - {order['Клиент']}\n")
            self.optimization_text.insert(tk.END, f"   Тип: {order['Тип продукции']}, Площадь: {order['Площадь']} м²\n")
            self.optimization_text.insert(tk.END, f"   Начало: {order['Начало производства'].strftime('%d.%m.%Y')}\n")
            self.optimization_text.insert(tk.END, f"   Окончание: {order['Окончание производства'].strftime('%d.%m.%Y')}\n")
            self.optimization_text.insert(tk.END, f"   Длительность: {order['Дней производства']} дней\n")
            self.optimization_text.insert(tk.END, "\n")
    
    def analyze_machine_utilization(self):
        """Анализ загрузки оборудования"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы для анализа!")
            return
        
        start_date = self.production_start_date.get_date()
        schedule = self.planner.optimize_production_schedule(selected_orders, start_date)
        
        if 'error' in schedule:
            messagebox.showerror("Ошибка", schedule['error'])
            return
        
        utilization = self.planner.calculate_machine_utilization(schedule['schedule'])
        
        # Вывод анализа загрузки
        self.optimization_text.delete(1.0, tk.END)
        self.optimization_text.insert(tk.END, "🔧 АНАЛИЗ ЗАГРУЗКИ ОБОРУДОВАНИЯ\n")
        self.optimization_text.insert(tk.END, "=" * 60 + "\n\n")
        
        for machine, data in utilization.items():
            utilization_percent = data['utilization_percent']
            status = "✅ Оптимальная" if 70 <= utilization_percent <= 90 else \
                    "⚠️ Перегрузка" if utilization_percent > 90 else \
                    "❌ Недогрузка"
            
            color = "green" if 70 <= utilization_percent <= 90 else \
                   "red" if utilization_percent > 90 else "orange"
            
            self.optimization_text.insert(tk.END, f"🏭 {machine}:\n")
            self.optimization_text.insert(tk.END, f"   Загрузка: {data['workload_hours']:.1f} ч / {data['capacity_hours']:.1f} ч\n")
            self.optimization_text.insert(tk.END, f"   Использование: {utilization_percent:.1f}% - {status}\n")
            
            # Рекомендации
            if utilization_percent > 90:
                self.optimization_text.insert(tk.END, f"   💡 Рекомендация: Увеличить мощность или распределить нагрузку\n")
            elif utilization_percent < 50:
                self.optimization_text.insert(tk.END, f"   💡 Рекомендация: Рассмотреть дополнительную загрузку\n")
            
            self.optimization_text.insert(tk.END, "\n")
    
    def group_orders_by_type(self):
        """Группировка заказов по типам продукции"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы для группировки!")
            return
        
        grouped = self.planner.group_orders_by_product_type(selected_orders)
        
        # Вывод группировки
        self.optimization_text.delete(1.0, tk.END)
        self.optimization_text.insert(tk.END, "📈 ГРУППИРОВКА ЗАКАЗОВ ПО ТИПАМ ПРОДУКЦИИ\n")
        self.optimization_text.insert(tk.END, "=" * 60 + "\n\n")
        
        total_orders = 0
        total_area = 0
        total_cost = 0
        
        for group in grouped:
            self.optimization_text.insert(tk.END, f"🏷️ {group['Тип продукции']}:\n")
            self.optimization_text.insert(tk.END, f"   Количество заказов: {group['Количество заказов']}\n")
            self.optimization_text.insert(tk.END, f"   Общая площадь: {group['Общая площадь']:.2f} м²\n")
            self.optimization_text.insert(tk.END, f"   Общая стоимость: {group['Общая стоимость']:,.2f} руб.\n")
            
            # Расчет эффективности группировки
            setup_reduction = max(0, (group['Количество заказов'] - 1) * 0.5)  # Экономия 0.5 дня на переналадку
            if setup_reduction > 0:
                self.optimization_text.insert(tk.END, f"   💰 Экономия на переналадках: {setup_reduction:.1f} дней\n")
            
            self.optimization_text.insert(tk.END, "\n")
            
            total_orders += group['Количество заказов']
            total_area += group['Общая площадь']
            total_cost += group['Общая стоимость']
        
        self.optimization_text.insert(tk.END, f"📊 ИТОГО:\n")
        self.optimization_text.insert(tk.END, f"   Всего заказов: {total_orders}\n")
        self.optimization_text.insert(tk.END, f"   Общая площадь: {total_area:.2f} м²\n")
        self.optimization_text.insert(tk.END, f"   Общая стоимость: {total_cost:,.2f} руб.\n")
    
    def analyze_material_deficit(self):
        """Анализ дефицита материалов"""
        # Показываем материалы с низким остатком
        self.optimization_text.delete(1.0, tk.END)
        self.optimization_text.insert(tk.END, "📊 АНАЛИЗ ДЕФИЦИТА МАТЕРИАЛОВ\n")
        self.optimization_text.insert(tk.END, "=" * 60 + "\n\n")
        
        low_stock_materials = []
        for material, stock in self.planner.stock_data.items():
            reserved = self.planner.reserved_materials.get(material, 0)
            available = max(0, stock - reserved)
            
            # Материалы с доступным количеством менее 20% от текущего запаса
            if stock > 0 and available < stock * 0.2:
                low_stock_materials.append((material, stock, reserved, available))
        
        if not low_stock_materials:
            self.optimization_text.insert(tk.END, "✅ Критического дефицита материалов не обнаружено\n")
            return
        
        # Сортируем по уровню дефицита
        low_stock_materials.sort(key=lambda x: x[3] / x[1] if x[1] > 0 else 0)
        
        self.optimization_text.insert(tk.END, "🚨 МАТЕРИАЛЫ С НИЗКИМ ОСТАТКОМ:\n\n")
        
        for material, stock, reserved, available in low_stock_materials[:10]:  # Топ-10
            deficit_level = (1 - available/stock) * 100 if stock > 0 else 100
            status = "⚡ КРИТИЧЕСКИЙ" if deficit_level > 80 else "⚠️ ВЫСОКИЙ" if deficit_level > 50 else "📉 НИЗКИЙ"
            
            self.optimization_text.insert(tk.END, f"📦 {material}:\n")
            self.optimization_text.insert(tk.END, f"   Запас: {stock:.2f}, Резерв: {reserved:.2f}, Доступно: {available:.2f}\n")
            self.optimization_text.insert(tk.END, f"   Уровень дефицита: {deficit_level:.1f}% - {status}\n\n")
    
    def show_orders_stats(self):
        """Показать статистику по заказам"""
        stats_text = "📊 СТАТИСТИКА ПО ЗАКАЗАМ\n"
        stats_text += "=" * 40 + "\n\n"
        
        stats_text += f"📈 Всего заказов: {len(self.planner.orders_df)}\n"
        stats_text += f"🏢 Компаний: {len(self.planner.get_companies())}\n"
        stats_text += f"📦 Типов продукции: {len(self.planner.get_product_types())}\n\n"
        
        # Статистика по компаниям
        company_stats = self.planner.orders_df.groupby('Клиент').agg({
            'Номер заказа': 'count',
            'Площадь заказа': 'sum',
            'Стоимость заказа': 'sum'
        }).sort_values('Стоимость заказа', ascending=False)
        
        stats_text += "ТОП-5 КОМПАНИЙ ПО СТОИМОСТИ:\n"
        for i, (company, row) in enumerate(company_stats.head().iterrows(), 1):
            stats_text += f"{i}. {company}: {row['Стоимость заказа']:,.2f} руб. ({row['Номер заказа']} зак.)\n"
        
        messagebox.showinfo("Статистика заказов", stats_text)
    
    def save_purchase_order(self, order_text):
        """Сохранение заявки в файл"""
        filename = f"Заявка_на_закупку_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(order_text)
            messagebox.showinfo("Успех", f"Заявка сохранена в файл: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")
    
    def print_purchase_order(self, order_text):
        """Печать заявки"""
        messagebox.showinfo("Печать", "Функция печати будет реализована в следующей версии")
    
    def load_materials_data(self):
        """Загрузка данных о материалах"""
        # Очистка таблицы
        for item in self.materials_tree.get_children():
            self.materials_tree.delete(item)
        
        # Заполнение данными
        for material, stock in self.planner.stock_data.items():
            reserved = self.planner.reserved_materials.get(material, 0)
            available = max(0, stock - reserved)
            
            # Определение статуса
            if available == 0:
                status = "🔴 Нет в наличии"
                recommendation = "СРОЧНАЯ ЗАКУПКА"
            elif available < stock * 0.2:
                status = "🟡 Низкий запас"
                recommendation = "Рекомендуется закупка"
            elif reserved > 0:
                status = f"🔵 Зарезервировано ({reserved})"
                recommendation = "В работе"
            else:
                status = "🟢 В наличии"
                recommendation = "Норма"
            
            self.materials_tree.insert('', tk.END, values=(
                material,
                f"{stock:.2f}",
                f"{reserved:.2f}",
                f"{available:.2f}",
                status,
                recommendation
            ))

def main():
    """Запуск приложения"""
    try:
        root = tk.Tk()
        app = AdvancedProductionPlannerGUI(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Не удалось запустить приложение: {e}")

if __name__ == "__main__":
    main()
