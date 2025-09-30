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

class ModernProductionPlanner:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.orders_df = None
        self.materials_df = None
        self.stock_data = {}
        self.reserved_materials = defaultdict(float)
        self.selected_orders = {}
        self.load_data()
    
    def load_data(self):
        """Загрузка данных из Excel файла"""
        try:
            print("📂 Загрузка данных из Excel файла...")
            
            # Загружаем лист с заказами
            self.orders_df = pd.read_excel(self.excel_file, sheet_name='Заказы')
            
            # Загружаем лист с материалами
            self.materials_df = pd.read_excel(self.excel_file, sheet_name='Потребность материалов')
            
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
    
    def get_companies(self):
        """Получить список компаний"""
        return sorted([str(x) for x in self.orders_df['Клиент'].unique() if pd.notna(x)])
    
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
        
        return {
            'material_requirements': dict(required_materials),
            'material_balance': material_balance,
            'purchase_requirements': purchase_requirements
        }

class ModernProductionPlannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🏭 Production Planner - Система планирования производства")
        self.root.geometry("1400x900")
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
        
        # Современная тема
        style.theme_use('clam')
        
        # Кастомные стили
        style.configure('Modern.TFrame', background='#f8f9fa')
        style.configure('Header.TLabel', background='#343a40', foreground='white', font=('Arial', 12, 'bold'))
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), background='#f8f9fa')
        style.configure('Card.TFrame', background='white', relief='raised', borderwidth=1)
        
        # Стили для кнопок
        style.configure('Primary.TButton', background='#007bff', foreground='white', font=('Arial', 10))
        style.map('Primary.TButton', background=[('active', '#0056b3')])
        
        style.configure('Success.TButton', background='#28a745', foreground='white', font=('Arial', 10))
        style.map('Success.TButton', background=[('active', '#1e7e34')])
        
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
            self.planner = ModernProductionPlanner(excel_file)
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
        
        title_label = ttk.Label(header_frame, text="🏭 PRODUCTION PLANNER", style='Title.TLabel')
        title_label.pack(pady=10)
        
        # Создание вкладок
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Вкладка 1: Обзор заказов
        self.setup_orders_tab(notebook)
        
        # Вкладка 2: Планирование производства
        self.setup_planning_tab(notebook)
        
        # Вкладка 3: Анализ материалов
        self.setup_materials_tab(notebook)
        
        # Вкладка 4: Дашборд
        self.setup_dashboard_tab(notebook)
    
    def setup_orders_tab(self, notebook):
        """Вкладка с обзором заказов"""
        orders_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(orders_frame, text="📋 Заказы")
        
        # Фильтры
        filter_frame = ttk.Frame(orders_frame, style='Card.TFrame')
        filter_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(filter_frame, text="Фильтр по компании:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.company_var = tk.StringVar()
        companies = ['Все компании'] + self.planner.get_companies()
        company_combo = ttk.Combobox(filter_frame, textvariable=self.company_var, values=companies, state='readonly')
        company_combo.set('Все компании')
        company_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        company_combo.bind('<<ComboboxSelected>>', self.filter_orders)
        
        # Поиск
        ttk.Label(filter_frame, text="Поиск:", font=('Arial', 10)).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=20)
        search_entry.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        search_entry.bind('<KeyRelease>', self.filter_orders)
        
        # Таблица заказов
        table_frame = ttk.Frame(orders_frame, style='Modern.TFrame')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создание Treeview с прокруткой
        columns = ('Номер', 'Клиент', 'Тип продукции', 'Площадь', 'Стоимость', 'Состояние', 'Выбор')
        self.orders_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Настройка колонок
        for col in columns:
            self.orders_tree.heading(col, text=col)
            self.orders_tree.column(col, width=120)
        
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
        
        ttk.Button(button_frame, text="✅ Выбрать отмеченные для планирования", 
                  command=self.add_selected_orders, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🔄 Обновить данные", 
                  command=self.load_orders_data, style='Success.TButton').pack(side=tk.LEFT, padx=5)
    
    def setup_planning_tab(self, notebook):
        """Вкладка планирования производства"""
        planning_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(planning_frame, text="📅 Планирование")
        
        # Левая панель - выбранные заказы
        left_frame = ttk.Frame(planning_frame, style='Modern.TFrame')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(left_frame, text="Выбранные заказы:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=5)
        
        # Список выбранных заказов
        self.selected_orders_listbox = tk.Listbox(left_frame, height=15, font=('Arial', 10))
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
        
        # Дата отгрузки
        ttk.Label(right_frame, text="Дата отгрузки:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=5)
        
        date_frame = ttk.Frame(right_frame, style='Modern.TFrame')
        date_frame.pack(fill=tk.X, pady=5)
        
        self.shipment_date = DateEntry(date_frame, width=12, background='darkblue',
                                      foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                                      font=('Arial', 10))
        self.shipment_date.pack(side=tk.LEFT, padx=5)
        
        # Кнопки планирования
        planning_buttons_frame = ttk.Frame(right_frame, style='Modern.TFrame')
        planning_buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(planning_buttons_frame, text="🧮 Рассчитать потребности", 
                  command=self.calculate_requirements, style='Primary.TButton').pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📦 Зарезервировать материалы", 
                  command=self.reserve_materials, style='Success.TButton').pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📄 Сформировать заявку на закупку", 
                  command=self.generate_purchase_order, style='Primary.TButton').pack(fill=tk.X, pady=2)
        
        # Область результатов
        ttk.Label(right_frame, text="Результаты расчета:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=(20, 5))
        
        self.results_text = scrolledtext.ScrolledText(right_frame, height=10, font=('Arial', 9))
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def setup_materials_tab(self, notebook):
        """Вкладка анализа материалов"""
        materials_frame = ttk.Frame(notebook, style='Modern.TFrame')
        notebook.add(materials_frame, text="📦 Материалы")
        
        # Таблица материалов
        table_frame = ttk.Frame(materials_frame, style='Modern.TFrame')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ('Материал', 'На складе', 'Зарезервировано', 'Доступно', 'Статус')
        self.materials_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, width=150)
        
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
        
        search_text = self.search_var.get().lower()
        if search_text:
            filtered_orders = filtered_orders[
                filtered_orders.apply(lambda row: search_text in str(row).lower(), axis=1)
            ]
        
        # Заполнение таблицы
        for _, order in filtered_orders.iterrows():
            self.orders_tree.insert('', tk.END, values=(
                order['Номер заказа'],
                order['Клиент'],
                order.get('Тип продукции', ''),
                f"{order.get('Площадь заказа', 0):.2f}",
                f"{order.get('Стоимость заказа', 0):,.2f}",
                order.get('Состояние заказа', ''),
                "✅"  # Галочка для выбора
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
        
        for item in selected_items:
            values = self.orders_tree.item(item)['values']
            order_num = values[0]
            
            # Добавляем в список, если еще нет
            if order_num not in self.get_selected_orders():
                self.selected_orders_listbox.insert(tk.END, order_num)
        
        messagebox.showinfo("Успех", f"Добавлено {len(selected_items)} заказов в планирование")
    
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
        self.results_text.insert(tk.END, "=" * 50 + "\n\n")
        
        # Показываем первые 10 материалов с наибольшей потребностью
        sorted_materials = sorted(requirements['material_balance'].items(), 
                                 key=lambda x: x[1]['Требуется для выбранных'], reverse=True)
        
        for material, balance in sorted_materials[:10]:
            if balance['Требуется для выбранных'] > 0:
                self.results_text.insert(tk.END, f"📦 {material}:\n")
                self.results_text.insert(tk.END, f"   Текущий запас: {balance['Текущий запас']:.2f}\n")
                self.results_text.insert(tk.END, f"   Доступно сейчас: {balance['Доступно сейчас']:.2f}\n")
                self.results_text.insert(tk.END, f"   Требуется: {balance['Требуется для выбранных']:.2f}\n")
                
                remaining = balance['Остаток после']
                if remaining >= 0:
                    self.results_text.insert(tk.END, f"   ✅ Остаток: {remaining:.2f}\n")
                else:
                    self.results_text.insert(tk.END, f"   ❌ ДЕФИЦИТ: {-remaining:.2f}\n")
                self.results_text.insert(tk.END, "\n")
        
        # Показываем заявку на закупку
        if requirements['purchase_requirements']:
            self.results_text.insert(tk.END, f"🚨 ТРЕБУЕТСЯ ЗАКУПКА ({len(requirements['purchase_requirements'])} материалов)\n")
    
    def reserve_materials(self):
        """Резервирование материалов"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        # Здесь должна быть логика резервирования
        messagebox.showinfo("Успех", f"Материалы зарезервированы для {len(selected_orders)} заказов")
        self.load_materials_data()  # Обновляем данные о материалах
    
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
        purchase_text += f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        purchase_text += f"Для заказов: {', '.join(selected_orders)}\n\n"
        
        total_cost = 0
        for material, quantity in requirements['purchase_requirements'].items():
            price = self.planner.estimate_material_price(material)
            cost = price * quantity
            total_cost += cost
            purchase_text += f"• {material}: {quantity:.2f} × {price:,.2f} руб. = {cost:,.2f} руб.\n"
        
        purchase_text += f"\nОБЩАЯ СТОИМОСТЬ: {total_cost:,.2f} руб."
        
        # Показ заявки
        self.show_purchase_order(purchase_text)
    
    def show_purchase_order(self, order_text):
        """Показ заявки на закупку в отдельном окне"""
        order_window = tk.Toplevel(self.root)
        order_window.title("Заявка на закупку")
        order_window.geometry("600x500")
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
        # В реальной системе здесь будет логика печати
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
            
            status = "✅ В наличии" if available > 0 else "⚠️ Нет в наличии"
            if reserved > 0:
                status = f"📦 Зарезервировано ({reserved})"
            
            self.materials_tree.insert('', tk.END, values=(
                material,
                f"{stock:.2f}",
                f"{reserved:.2f}",
                f"{available:.2f}",
                status
            ))

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

def main():
    """Запуск приложения"""
    try:
        root = tk.Tk()
        app = ModernProductionPlannerGUI(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Не удалось запустить приложение: {e}")

if __name__ == "__main__":
    main()
